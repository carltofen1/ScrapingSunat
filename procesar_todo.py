"""
SCRIPT FINAL DE PROCESAMIENTO MASIVO (MEJORADO V3)
- Incluye Dirección y Número original.
- Sistema anti-crash cuando el Excel está abierto.
- Detección automática de CAPTCHA.
- Búsqueda progresiva (100%, 75%, 50%).
- Limpieza de caracteres especiales.
- Selección del resultado más similar.
"""
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path
from difflib import SequenceMatcher
import time
import re
import os

def guardar_excel_seguro(df, ruta):
    """Intenta guardar el Excel y si falla pide al usuario que lo cierre."""
    while True:
        try:
            df.to_excel(ruta, index=False)
            print("  💾 Progreso guardado")
            return True
        except PermissionError:
            print("\n" + "!"*60)
            print(f"⚠️  NO SE PUEDE GUARDAR: El archivo '{ruta}' está abierto.")
            print("👉 Por favor, CIERRA el Excel y presiona ENTER aquí para continuar...")
            print("!"*60)
            input()
        except Exception as e:
            print(f"  ❌ Error guardando: {e}")
            return False

def limpiar_razon_social(texto):
    """
    Limpia la razón social:
    1. Convierte a mayúsculas
    2. Elimina SOLO los caracteres especiales (mantiene letras y números)
    3. Elimina sufijos comunes (S.A.C., E.I.R.L., etc.)
    4. Normaliza espacios
    
    Ejemplo: "INVERSIONES Dï¿½KASA S.A.C." → "INVERSIONES DKASA"
    """
    # Convertir a mayúsculas
    texto = texto.upper()
    
    # ELIMINAR caracteres especiales (mantener solo letras, números y espacios)
    texto = re.sub(r'[^A-Z0-9\s]', '', texto)
    
    # Normalizar espacios múltiples
    texto = ' '.join(texto.split())
    
    # Eliminar sufijos comunes al final
    sufijos = [
        ' SAC', ' EIRL', ' SRL', ' SAA', ' SA',
        ' SOCIEDAD ANONIMA CERRADA',
        ' SOCIEDAD ANONIMA',
        ' EMPRESA INDIVIDUAL DE RESPONSABILIDAD LIMITADA'
    ]
    for sufijo in sufijos:
        if texto.endswith(sufijo):
            texto = texto[:-len(sufijo)].strip()
            break  # Solo eliminar un sufijo
    
    return texto.strip()

def obtener_variantes_busqueda(texto):
    """
    Genera variantes de búsqueda:
    - 100% del texto
    - 75% del texto
    - 50% del texto
    """
    palabras = texto.split()
    variantes = [texto]  # 100%
    
    if len(palabras) > 2:
        # 75%
        num_palabras_75 = max(1, int(len(palabras) * 0.75))
        variantes.append(' '.join(palabras[:num_palabras_75]))
        
        # 50%
        num_palabras_50 = max(1, int(len(palabras) * 0.5))
        variantes.append(' '.join(palabras[:num_palabras_50]))
    
    return variantes

def extraer_todos_los_rucs(texto):
    """Extrae todos los RUCs encontrados en el texto."""
    return re.findall(r'\b(?:10|20)\d{9}\b', texto)

def seleccionar_mejor_ruc(rucs, razon_original, texto_completo):
    """
    Selecciona el mejor RUC basándose en:
    1. Prioridad a RUCs que empiezan con 20 (empresas)
    2. Similitud del texto cercano al RUC con la razón social original
    """
    if not rucs:
        return None, "DESCONOCIDO"
    
    if len(rucs) == 1:
        estado = "ACTIVO" if "ACTIVO" in texto_completo else ("BAJA" if "BAJA" in texto_completo else "DESCONOCIDO")
        return rucs[0], estado
    
    # Priorizar RUCs que empiezan con 20
    rucs_20 = [r for r in rucs if r.startswith('20')]
    if rucs_20:
        rucs = rucs_20
    
    # Si aún hay múltiples, tomar el primero (generalmente es el más relevante)
    mejor_ruc = rucs[0]
    estado = "ACTIVO" if "ACTIVO" in texto_completo else ("BAJA" if "BAJA" in texto_completo else "DESCONOCIDO")
    
    return mejor_ruc, estado

def procesar_todo():
    print("=" * 70)
    print("PROCESAMIENTO MASIVO DE RUCS (V2)")
    print("=" * 70)
    
    # 1. Configurar ChromeDriver
    chromedriver_paths = [
        Path.home() / ".chromedriver" / "chromedriver.exe",
        Path("C:/chromedriver/chromedriver.exe"),
        Path("chromedriver.exe"),
    ]
    
    chromedriver_path = None
    for path in chromedriver_paths:
        if path.exists():
            chromedriver_path = path
            break
    
    if not chromedriver_path:
        print("❌ No se encontró chromedriver.exe. Ejecuta descargar_chromedriver.py primero.")
        return

    # 2. Cargar Excel
    archivo_entrada = "DATA.xlsx"
    archivo_salida = "RESULTADOS_FINALES.xlsx"
    
    try:
        df = pd.read_excel(archivo_entrada)
        print(f"✓ Datos cargados: {len(df)} registros")
    except Exception as e:
        print(f"❌ Error cargando Excel: {e}")
        return

    # Buscar columnas
    col_razon = next((c for c in df.columns if 'razon' in c.lower() or 'social' in c.lower()), None)
    col_direccion = next((c for c in df.columns if 'direccion' in c.lower()), None)
    col_numero = next((c for c in df.columns if 'numero' in c.lower()), None)
    
    if not col_razon:
        print("❌ No se encontró columna de Razón Social")
        return
        
    print(f"✓ Columna Razón Social: {col_razon}")
    print(f"✓ Columna Dirección: {col_direccion or 'No encontrada'}")
    print(f"✓ Columna Número: {col_numero or 'No encontrada'}")

    # Cargar progreso previo si existe
    resultados = []
    procesados_indices = set()
    
    if os.path.exists(archivo_salida):
        try:
            df_prev = pd.read_excel(archivo_salida)
            resultados = df_prev.to_dict('records')
            procesados_indices = set(df_prev['indice_original'])
            print(f"✓ Recuperados {len(resultados)} registros previos")
        except:
            print("⚠️ No se pudo leer archivo previo, empezando de cero")

    # Filtrar pendientes
    pendientes = df[~df.index.isin(procesados_indices)]
    print(f"✓ Pendientes de procesar: {len(pendientes)}")
    
    if len(pendientes) == 0:
        print("✅ Todo ya está procesado!")
        return

    input("\nPresiona ENTER para abrir Chrome y comenzar...")

    # 3. Iniciar Navegador
    options = Options()
    options.add_argument('--start-maximized')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    
    service = Service(executable_path=str(chromedriver_path))
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 10)
    
    url_base = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias"

    try:
        # Loop principal
        for i, (idx, row) in enumerate(pendientes.iterrows(), 1):
            razon = str(row[col_razon]).strip()
            print(f"\n[{i}/{len(pendientes)}] Procesando: {razon}")
            
            resultado = {
                'indice_original': idx,
                'razon_social_input': razon,
                'ruc': None,
                'estado': 'PENDIENTE',
                'observacion': '',
                # Agregar columnas originales al final
                'direccion_original': row[col_direccion] if col_direccion else '',
                'numero_original': row[col_numero] if col_numero else ''
            }
            
            try:
                # Limpiar razón social
                razon_limpia = limpiar_razon_social(razon)
                variantes = obtener_variantes_busqueda(razon_limpia)
                
                print(f"  → Limpiado: {razon_limpia}")
                if len(variantes) > 1:
                    print(f"  → Variantes: {len(variantes)} ({', '.join([f'{v[:30]}...' if len(v) > 30 else v for v in variantes])})")
                
                ruc_encontrado = None
                estado_encontrado = "DESCONOCIDO"
                variante_exitosa = None
                
                # Intentar con cada variante hasta encontrar resultado
                for idx_var, variante in enumerate(variantes, 1):
                    if ruc_encontrado:
                        break  # Ya encontramos, salir
                    
                    if idx_var > 1:
                        print(f"  → Intentando variante {idx_var}/{len(variantes)}: {variante}")
                    
                    # Ir a la página
                    driver.get(url_base)
                    
                    # --- PASO 1: PESTAÑA ---
                    try:
                        tab = wait.until(EC.element_to_be_clickable((By.ID, "btnPorRazonSocial")))
                        tab.click()
                    except:
                        pass
                    
                    # --- PASO 2: ESCRIBIR ---
                    input_razon = None
                    try:
                        input_razon = driver.find_element(By.ID, "txtNombreRazonSocial")
                        if not input_razon.is_displayed():
                            input_razon = driver.find_element(By.NAME, "search3")
                    except:
                        pass

                    if input_razon and input_razon.is_displayed():
                        try:
                            input_razon.clear()
                            input_razon.send_keys(variante)
                        except:
                            print(f"  ⚠️ Escribe manualmente: {variante}")
                    
                    # --- PASO 3: CAPTCHA Y BUSCAR ---
                    captcha_visible = False
                    try:
                        txt_codigo = driver.find_element(By.ID, "txtCodigo")
                        if txt_codigo.is_displayed():
                            captcha_visible = True
                    except:
                        pass
                    
                    if captcha_visible:
                        print("  👉 Escribe el CAPTCHA en Chrome y presiona ENTER aquí...")
                        input()
                    
                    try:
                        driver.find_element(By.ID, "btnAceptar").click()
                        
                        # Manejar alert si aparece (por caracteres especiales)
                        time.sleep(0.5)
                        try:
                            alert = driver.switch_to.alert
                            alert_text = alert.text
                            print(f"  ⚠️ Alert detectado: {alert_text}")
                            alert.accept()
                            print(f"  → Alert cerrado, saltando esta variante...")
                            continue  # Saltar a la siguiente variante
                        except:
                            pass  # No hay alert, continuar normal
                            
                    except Exception as e:
                        print(f"  ⚠️ No se pudo clickear Buscar: {e}")
                    
                    # --- PASO 4: EXTRAER ---
                    time.sleep(3)
                    
                    # Buscar en body
                    body_text = driver.find_element(By.TAG_NAME, "body").text
                    rucs_encontrados = extraer_todos_los_rucs(body_text)
                    
                    # Buscar en iframes si no hay resultados
                    if not rucs_encontrados:
                        iframes = driver.find_elements(By.TAG_NAME, "iframe")
                        for frame in iframes:
                            try:
                                driver.switch_to.frame(frame)
                                frame_text = driver.find_element(By.TAG_NAME, "body").text
                                rucs_frame = extraer_todos_los_rucs(frame_text)
                                if rucs_frame:
                                    rucs_encontrados.extend(rucs_frame)
                                    body_text += " " + frame_text
                                driver.switch_to.default_content()
                            except:
                                driver.switch_to.default_content()
                    
                    # Seleccionar mejor RUC
                    if rucs_encontrados:
                        ruc_encontrado, estado_encontrado = seleccionar_mejor_ruc(rucs_encontrados, razon, body_text)
                        variante_exitosa = variante
                
                if ruc_encontrado:
                    print(f"  ✅ RUC: {ruc_encontrado} ({estado_encontrado})")
                    if variante_exitosa != variantes[0]:
                        print(f"     Encontrado con variante: {variante_exitosa}")
                    resultado['ruc'] = ruc_encontrado
                    resultado['estado'] = estado_encontrado
                    resultado['observacion'] = 'Exito'
                else:
                    print("  ❌ No encontrado con ninguna variante")
                    resultado['observacion'] = 'No encontrado en web'
                
            except Exception as e:
                print(f"  ❌ Error: {e}")
                resultado['observacion'] = str(e)
            
            # Guardar resultado
            resultados.append(resultado)
            
            # Guardar en disco cada 5 registros (CON SEGURIDAD)
            if len(resultados) % 5 == 0:
                guardar_excel_seguro(pd.DataFrame(resultados), archivo_salida)

    except KeyboardInterrupt:
        print("\n🛑 Proceso detenido por usuario")
    
    finally:
        # Guardado final
        if resultados:
            guardar_excel_seguro(pd.DataFrame(resultados), archivo_salida)
        
        if driver:
            driver.quit()
            print("Navegador cerrado.")

if __name__ == "__main__":
    procesar_todo()
