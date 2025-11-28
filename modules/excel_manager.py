import pandas as pd
import os
from typing import List, Dict, Any
import threading
import config

class ExcelManager:
    
    def __init__(self, input_file: str = None, output_file: str = None):
        self.input_file = input_file or config.INPUT_FILE
        self.output_file = output_file or config.OUTPUT_FILE
        self.lock = threading.Lock()
        
    def load_data(self) -> pd.DataFrame:
        try:
            df = pd.read_excel(self.input_file)
            print(f"Datos cargados: {len(df)} registros desde {self.input_file}")
            return df
        except Exception as e:
            print(f"ERROR cargando Excel: {e}")
            raise
    
    def find_columns(self, df: pd.DataFrame) -> Dict[str, str]:
        columns = {}
        
        columns['razon'] = next(
            (c for c in df.columns if 'razon' in c.lower() or 'social' in c.lower()), 
            None
        )
        
        columns['direccion'] = next(
            (c for c in df.columns if 'direccion' in c.lower()), 
            None
        )
        
        columns['numero'] = next(
            (c for c in df.columns if 'numero' in c.lower()), 
            None
        )
        
        if not columns['razon']:
            raise ValueError("No se encontro columna de Razon Social")
        
        print(f"Columna Razon Social: {columns['razon']}")
        print(f"Columna Direccion: {columns['direccion'] or 'No encontrada'}")
        print(f"Columna Numero: {columns['numero'] or 'No encontrada'}")
        
        return columns
    
    def load_previous_results(self) -> tuple[List[Dict], set]:
        resultados = []
        procesados_indices = set()
        
        if os.path.exists(self.output_file):
            try:
                df_prev = pd.read_excel(self.output_file)
                resultados = df_prev.to_dict('records')
                procesados_indices = set(df_prev['indice_original'])
                print(f"Recuperados {len(resultados)} registros previos")
            except Exception as e:
                print(f"ADVERTENCIA: No se pudo leer archivo previo: {e}")
        
        return resultados, procesados_indices
    
    def save_results(self, resultados: List[Dict], force: bool = False, pause_event: threading.Event = None) -> bool:
        with self.lock:
            intentos = 0
            while True:
                try:
                    df = pd.DataFrame(resultados)
                    existing_cols = [col for col in config.OUTPUT_COLUMNS if col in df.columns]
                    df = df[existing_cols]
                    
                    df.to_excel(self.output_file, index=False)
                    if intentos > 0:
                        print(f"\n  Guardado exitoso despues de {intentos} intentos")
                        if pause_event:
                            pause_event.set() # Reanudar workers
                    print(f"  Progreso guardado: {len(resultados)} registros")
                    return True
                    
                except PermissionError:
                    if pause_event:
                        pause_event.clear() # Pausar workers inmediatamente
                        
                    intentos += 1
                    print("\n" + "="*70)
                    print("PAUSA AUTOMATICA - ARCHIVO ABIERTO")
                    print("="*70)
                    print(f"El archivo '{self.output_file}' esta abierto en Excel.")
                    print("TODOS LOS WORKERS ESTAN PAUSADOS esperando que cierres el archivo.")
                    print("\nCierra el Excel y presiona ENTER para continuar...")
                    print("="*70)
                    input()
                    print("\nReintentando guardar...")
                    
                except Exception as e:
                    print(f"\n  ERROR CRITICO guardando: {e}")
                    if force:
                        print("  Reintentando en 5 segundos...")
                        import time
                        time.sleep(5)
                    else:
                        return False
    
    def get_pending_records(self, df: pd.DataFrame, procesados: set) -> pd.DataFrame:
        pendientes = df[~df.index.isin(procesados)]
        print(f"Pendientes de procesar: {len(pendientes)}")
        return pendientes
    
    def deduplicate_consecutive(self, pendientes: pd.DataFrame, col_razon: str) -> tuple:
        """
        Detecta y agrupa razones sociales consecutivas duplicadas.
        Retorna: (df_unicos, mapa_duplicados)
        - df_unicos: DataFrame con solo registros unicos
        - mapa_duplicados: {idx_unico: [idx1, idx2, ...]} indices duplicados
        """
        indices_unicos = []
        mapa_duplicados = {}
        razon_anterior = None
        idx_representante = None
        
        for idx, row in pendientes.iterrows():
            razon_actual = str(row[col_razon]).strip().upper()
            
            if razon_actual != razon_anterior:
                indices_unicos.append(idx)
                idx_representante = idx
                mapa_duplicados[idx_representante] = [idx]
                razon_anterior = razon_actual
            else:
                mapa_duplicados[idx_representante].append(idx)
        
        df_unicos = pendientes.loc[indices_unicos]
        
        total_duplicados = sum(len(v) - 1 for v in mapa_duplicados.values())
        print(f"Deduplicacion: {len(pendientes)} registros -> {len(df_unicos)} unicos")
        print(f"Se evitaran {total_duplicados} busquedas duplicadas")
        
        for idx_rep, indices in mapa_duplicados.items():
            if len(indices) > 1:
                razon = str(pendientes.loc[idx_rep, col_razon]).strip()
                print(f"  '{razon}': {len(indices)} duplicados consecutivos")
        
        return df_unicos, mapa_duplicados
    
    def distribute_work(self, pendientes: pd.DataFrame, num_workers: int) -> Dict[int, List[tuple]]:
        work_distribution = {i: [] for i in range(num_workers)}
        
        for i, (idx, row) in enumerate(pendientes.iterrows()):
            worker_id = i % num_workers
            work_distribution[worker_id].append((idx, row))
        
        for worker_id, items in work_distribution.items():
            print(f"  Worker {worker_id}: {len(items)} registros")
        
        return work_distribution
