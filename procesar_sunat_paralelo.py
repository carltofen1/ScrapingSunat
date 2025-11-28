import threading
import time
from typing import List, Dict
import config
from modules.excel_manager import ExcelManager
from modules.sunat_scraper import SunatScraper

class WorkerThread(threading.Thread):
    
    def __init__(self, worker_id: int, work_items: List[tuple], 
                 columns: Dict, resultados: List[Dict], lock: threading.Lock, pause_event: threading.Event):
        super().__init__()
        self.worker_id = worker_id
        self.work_items = work_items
        self.columns = columns
        self.resultados = resultados
        self.lock = lock
        self.pause_event = pause_event
        self.scraper = None
        
    def run(self):
        print(f"\n{'='*60}")
        print(f"[Worker {self.worker_id}] INICIANDO - {len(self.work_items)} registros asignados")
        print(f"{'='*60}")
        
        self.scraper = SunatScraper(worker_id=self.worker_id)
        if not self.scraper.initialize_driver():
            print(f"[Worker {self.worker_id}] ERROR: No se pudo inicializar Chrome")
            return
        
        try:
            for i, (idx, row) in enumerate(self.work_items, 1):
                # Esperar si esta pausado
                self.pause_event.wait()
                
                razon = str(row[self.columns['razon']]).strip()
                
                print(f"\n[Worker {self.worker_id}] [{i}/{len(self.work_items)}] Procesando: {razon}")
                
                resultado = {
                    'indice_original': idx,
                    'razon_social_input': razon,
                    'ruc': None,
                    'estado': config.STATUS['PENDING'],
                    'observacion': '',
                    'direccion_original': row[self.columns['direccion']] if self.columns['direccion'] else '',
                    'numero_original': row[self.columns['numero']] if self.columns['numero'] else '',
                    'worker_id': self.worker_id
                }
                
                busqueda = self.scraper.buscar_ruc(razon)
                resultado.update(busqueda)
                
                # DETECTAR ERROR CRITICO DE CONEXION
                if resultado.get('estado') == 'ERROR_CONEXION':
                    print(f"\n{'='*70}")
                    print(f"EMERGENCIA: ERROR DE CONEXION EN WORKER {self.worker_id}")
                    print(f"{'='*70}")
                    print(f"Pausando TODOS los workers para evitar perdida de datos...")
                    self.pause_event.clear()  # PAUSAR TODOS
                    
                    with self.lock:
                        self.resultados.append(resultado)
                    
                    # Cerrar el driver actual (puede estar corrupto)
                    print(f"[Worker {self.worker_id}] Cerrando Chrome dañado...")
                    if self.scraper:
                        try:
                            self.scraper.close_driver()
                        except:
                            pass
                    
                    # NO continuar procesando hasta que se reactive pause_event
                    print(f"[Worker {self.worker_id}] Esperando reanudacion...")
                    self.pause_event.wait()  # Esperar a que se reactive
                    
                    # REINICIALIZAR el driver después de reanudar
                    print(f"[Worker {self.worker_id}] Reanudando... reinicializando Chrome")
                    max_reintentos = 3
                    for intento in range(1, max_reintentos + 1):
                        if self.scraper.initialize_driver():
                            print(f"[Worker {self.worker_id}] Chrome reinicializado exitosamente")
                            break
                        else:
                            print(f"[Worker {self.worker_id}] Intento {intento}/{max_reintentos} falló")
                            if intento < max_reintentos:
                                time.sleep(3)
                            else:
                                print(f"[Worker {self.worker_id}] No se pudo reinicializar. Terminando worker.")
                                return
                    
                    # Continuar con el siguiente registro
                    time.sleep(config.DELAY_BETWEEN_BATCHES)
                    continue
                
                with self.lock:
                    self.resultados.append(resultado)
                
                time.sleep(config.DELAY_BETWEEN_BATCHES)
        
        except Exception as e:
            print(f"[Worker {self.worker_id}] ERROR CRITICO: {e}")
        
        finally:
            if self.scraper:
                self.scraper.close_driver()
            
            print(f"\n[Worker {self.worker_id}] FINALIZADO")


def procesar_paralelo():
    print("="*70)
    print("PROCESAMIENTO PARALELO DE RUCs - SUNAT")
    print(f"Workers: {config.NUM_WORKERS}")
    print("="*70)
    
    excel_manager = ExcelManager()
    
    try:
        df = excel_manager.load_data()
        columns = excel_manager.find_columns(df)
    except Exception as e:
        print(f"ERROR en carga inicial: {e}")
        return
    
    resultados, procesados_indices = excel_manager.load_previous_results()
    
    pendientes = excel_manager.get_pending_records(df, procesados_indices)
    
    if len(pendientes) == 0:
        print("\nTodo ya esta procesado")
        return
    
    print(f"\n{'='*70}")
    print("DEDUPLICACION DE REGISTROS CONSECUTIVOS")
    print(f"{'='*70}")
    pendientes_unicos, mapa_duplicados = excel_manager.deduplicate_consecutive(pendientes, columns['razon'])
    
    print(f"\nDistribuyendo {len(pendientes_unicos)} registros unicos entre {config.NUM_WORKERS} workers:")
    work_distribution = excel_manager.distribute_work(pendientes_unicos, config.NUM_WORKERS)
    
    print("\n" + "="*70)
    if config.HEADLESS_MODE:
        print("MODO HEADLESS ACTIVADO")
        print("Los 5 Chrome se ejecutaran en segundo plano (sin ventanas)")
    else:
        print("ADVERTENCIA: Se abriran 5 ventanas de Chrome simultaneamente")
    print("="*70)
    input("\nPresiona ENTER para comenzar el procesamiento paralelo...")
    
    lock = threading.Lock()
    pause_event = threading.Event()
    pause_event.set() # Inicialmente activo (no pausado)
    
    workers = []
    for worker_id in range(config.NUM_WORKERS):
        work_items = work_distribution[worker_id]
        if work_items:
            worker = WorkerThread(
                worker_id=worker_id,
                work_items=work_items,
                columns=columns,
                resultados=resultados,
                lock=lock,
                pause_event=pause_event
            )
            workers.append(worker)
            worker.start()
            time.sleep(2)
    
    print("\n" + "="*70)
    print("Sistema de guardado automatico activado (cada 30 segundos)")
    print("="*70)
    print("\nADVERTENCIA: Si abres RESULTADOS_FINALES.xlsx durante la ejecucion,")
    print("el programa se PAUSARA automaticamente hasta que cierres el archivo.")
    print("Todos los workers esperaran. No se perdera ningun dato.")
    print("="*70)
    
    last_save_count = len(resultados)
    emergency_detected = False
    
    while any(w.is_alive() for w in workers):
        time.sleep(5)  # Revisar cada 5 segundos (mas frecuente para detectar emergencias)
        
        # Detectar si se activo pausa de emergencia
        if not pause_event.is_set() and not emergency_detected:
            # Buscar si hay algún ERROR_CONEXION en resultados
            tiene_error_conexion = any(r.get('estado') == 'ERROR_CONEXION' for r in resultados)
            
            if tiene_error_conexion:
                emergency_detected = True
                print(f"\n{'!'*70}")
                print("SISTEMA DE EMERGENCIA ACTIVADO")
                print(f"{'!'*70}")
                print("\nSe detecto un error de conexion critico.")
                print("TODOS LOS WORKERS HAN SIDO PAUSADOS")
                print("\nGuardando datos de emergencia...")
                
                # Guardar inmediatamente
                excel_manager.save_results(resultados, force=True, pause_event=None)
                
                print(f"\n{'='*70}")
                print("DATOS GUARDADOS EXITOSAMENTE")
                print(f"{'='*70}")
                print(f"\nRegistros guardados: {len(resultados)}")
                print(f"Archivo: {config.OUTPUT_FILE}")
                
                # Verificar si los workers siguen vivos
                workers_vivos = [w for w in workers if w.is_alive()]
                workers_muertos = len(workers) - len(workers_vivos)
                
                print(f"\n{'='*70}")
                print("ESTADO DE WORKERS:")
                print(f"{'='*70}")
                print(f"Workers activos: {len(workers_vivos)}/{len(workers)}")
                if workers_muertos > 0:
                    print(f"Workers cerrados: {workers_muertos}")
                    print("\nNOTA: Algunos workers se cerraron (posiblemente Chrome crasheó).")
                    print("Los workers restantes intentarán reinicializarse al reanudar.")
                
                print(f"\n{'='*70}")
                print("INSTRUCCIONES:")
                print(f"{'='*70}")
                print("1. Verifica tu conexion a internet")
                print("2. Revisa el archivo de resultados para confirmar que no se perdieron datos")
                print("3. Cuando estes listo, presiona ENTER para REANUDAR el procesamiento")
                print("4. Si todos los workers murieron, el programa terminará y podrás ejecutarlo")
                print("   de nuevo - continuará desde donde quedó")
                print("5. O presiona Ctrl+C para TERMINAR el programa ahora")
                print(f"{'='*70}")
                
                input("\nPresiona ENTER para REANUDAR (o Ctrl+C para salir)...")
                
                # Verificar de nuevo cuántos workers siguen vivos
                workers_vivos = [w for w in workers if w.is_alive()]
                if len(workers_vivos) == 0:
                    print("\n" + "="*70)
                    print("TODOS LOS WORKERS HAN TERMINADO")
                    print("="*70)
                    print("\nNo quedan workers activos para continuar.")
                    print("El progreso ha sido guardado.")
                    print("\nPara continuar procesando, ejecuta el programa de nuevo.")
                    print("Automaticamente continuará desde donde quedó.")
                    print("="*70)
                    return  # Terminar la función principal
                
                print(f"\nReanudando procesamiento con {len(workers_vivos)} workers...")
                pause_event.set()  # Reanudar workers
                emergency_detected = False
        
        # Guardado periodico normal (cada 30 segundos)
        current_count = len(resultados)
        if current_count > last_save_count:
            # Solo guardar si no estamos en emergencia
            if pause_event.is_set():
                print(f"\nGuardando progreso... ({current_count} registros totales)")
                excel_manager.save_results(resultados, force=True, pause_event=pause_event)
                last_save_count = current_count
    
    for worker in workers:
        worker.join()
    
    print("\n" + "="*70)
    print("REPLICANDO RESULTADOS A DUPLICADOS")
    print("="*70)
    
    resultados_replicados = []
    for resultado in resultados:
        idx_original = resultado['indice_original']
        if idx_original in mapa_duplicados:
            indices_duplicados = mapa_duplicados[idx_original]
            for idx_dup in indices_duplicados:
                resultado_copia = resultado.copy()
                resultado_copia['indice_original'] = idx_dup
                if idx_dup != idx_original:
                    resultado_copia['observacion'] += ' (duplicado)'
                resultados_replicados.append(resultado_copia)
    
    print(f"Resultados replicados: {len(resultados)} -> {len(resultados_replicados)}")
    resultados = resultados_replicados
    
    print("\n" + "="*70)
    print("GUARDADO FINAL")
    print("="*70)
    excel_manager.save_results(resultados, force=True, pause_event=pause_event)
    
    print("\n" + "="*70)
    print("PROCESAMIENTO COMPLETADO")
    print("="*70)
    print(f"Total procesado: {len(resultados)} registros")
    print(f"Archivo de salida: {config.OUTPUT_FILE}")
    
    exitosos = sum(1 for r in resultados if r.get('ruc'))
    no_encontrados = sum(1 for r in resultados if r.get('estado') == config.STATUS['NOT_FOUND'])
    errores = sum(1 for r in resultados if r.get('estado') == config.STATUS['ERROR'])
    
    print(f"\nEstadisticas:")
    print(f"  Exitosos: {exitosos}")
    print(f"  No encontrados: {no_encontrados}")
    print(f"  Errores: {errores}")


if __name__ == "__main__":
    try:
        procesar_paralelo()
    except KeyboardInterrupt:
        print("\n\nProceso interrumpido por usuario")
        print("El progreso ha sido guardado automaticamente")
    except Exception as e:
        print(f"\nERROR CRITICO: {e}")
        import traceback
        traceback.print_exc()
