from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path
import time
import re
import config

class SunatScraper:
    
    def __init__(self, worker_id: int = 0):
        self.worker_id = worker_id
        self.driver = None
        self.wait = None
        
    def initialize_driver(self) -> bool:
        try:
            chromedriver_path = None
            for path_str in config.CHROMEDRIVER_PATHS:
                path = Path(path_str)
                if path.exists():
                    chromedriver_path = path
                    break
            
            if not chromedriver_path:
                print(f"[Worker {self.worker_id}] ERROR: No se encontro chromedriver.exe")
                return False
            
            options = Options()
            
            # Modo headless (solo Workers 1-4, Worker 0 visible para debugging)
            if config.HEADLESS_MODE and self.worker_id != 0:
                options.add_argument('--headless=new')
                print(f"[Worker {self.worker_id}] Modo headless activado")
            elif self.worker_id == 0:
                print(f"[Worker {self.worker_id}] Modo VISIBLE para debugging")
            
            # Optimizaciones de rendimiento
            options.add_argument('--start-maximized')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-blink-features=AutomationControlled')
            
            # Deshabilitar carga de recursos innecesarios (pero mantener JS para SUNAT)
            options.add_argument('--disable-plugins')
            options.add_argument('--disable-extensions')
            
            # Deshabilitar GPU y aceleracion
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-software-rasterizer')
            
            # Optimizaciones de red
            options.add_argument('--disable-background-networking')
            options.add_argument('--disable-default-apps')
            options.add_argument('--disable-sync')
            
            # Preferencias para bloquear recursos pesados (pero permitir JS)
            prefs = {
                'profile.managed_default_content_settings.images': 2,
                'profile.managed_default_content_settings.stylesheets': 2,
                'profile.managed_default_content_settings.javascript': 1,  # 1 = permitir
                'profile.managed_default_content_settings.plugins': 2,
                'profile.managed_default_content_settings.popups': 2,
                'profile.managed_default_content_settings.geolocation': 2,
                'profile.managed_default_content_settings.media_stream': 2,
            }
            options.add_experimental_option('prefs', prefs)
            options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
            options.add_experimental_option('useAutomationExtension', False)
            
            service = Service(executable_path=str(chromedriver_path))
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, config.SELENIUM_TIMEOUT)
            
            print(f"[Worker {self.worker_id}] Chrome inicializado (modo optimizado)")
            return True
            
        except Exception as e:
            print(f"[Worker {self.worker_id}] ERROR inicializando Chrome: {e}")
            return False
    
    def close_driver(self):
        if self.driver:
            try:
                self.driver.quit()
                print(f"[Worker {self.worker_id}] Chrome cerrado")
            except:
                pass
            finally:
                self.driver = None
                self.wait = None
    
    def is_driver_alive(self) -> bool:
        """Verifica si el driver sigue activo y funcional"""
        if not self.driver:
            return False
        try:
            # Intentar obtener el título de la página actual
            _ = self.driver.title
            return True
        except:
            return False
    
    def limpiar_razon_social(self, texto: str) -> str:
        texto = texto.upper()
        texto = re.sub(r'[^A-Z0-9\s]', '', texto)
        texto = ' '.join(texto.split())
        
        for sufijo in config.SUFIJOS_EMPRESAS:
            if texto.endswith(sufijo):
                texto = texto[:-len(sufijo)].strip()
                break
        
        return texto.strip()
    
    def obtener_variantes_busqueda(self, texto: str) -> list:
        palabras = texto.split()
        variantes = [texto]
        
        if len(palabras) > 2:
            num_palabras_75 = max(1, int(len(palabras) * 0.75))
            variantes.append(' '.join(palabras[:num_palabras_75]))
            
            num_palabras_50 = max(1, int(len(palabras) * 0.5))
            variantes.append(' '.join(palabras[:num_palabras_50]))
        
        return variantes
    
    def extraer_todos_los_rucs(self, texto: str) -> list:
        return re.findall(r'\b(?:10|20)\d{9}\b', texto)
    
    def seleccionar_mejor_ruc(self, rucs: list, texto_completo: str) -> tuple:
        if not rucs:
            return None, config.STATUS['NOT_FOUND']
        
        if len(rucs) == 1:
            estado = self._extraer_estado(texto_completo)
            return rucs[0], estado
        
        rucs_20 = [r for r in rucs if r.startswith('20')]
        if rucs_20:
            rucs = rucs_20
        
        mejor_ruc = rucs[0]
        estado = self._extraer_estado(texto_completo)
        
        return mejor_ruc, estado
    
    def _extraer_estado(self, texto: str) -> str:
        texto_upper = texto.upper()
        if "ACTIVO" in texto_upper:
            return "ACTIVO"
        elif "BAJA" in texto_upper:
            return "BAJA"
        elif "SUSPENSION" in texto_upper:
            return "SUSPENDIDO"
        else:
            return "DESCONOCIDO"
    
    def buscar_ruc(self, razon_social: str) -> dict:
        resultado = {
            'ruc': None,
            'estado': config.STATUS['PENDING'],
            'observacion': ''
        }
        
        # Verificar que el driver esté vivo antes de comenzar
        if not self.is_driver_alive():
            print(f"[Worker {self.worker_id}] ERROR: Chrome no está activo")
            resultado['estado'] = 'ERROR_CONEXION'
            resultado['observacion'] = 'Chrome cerrado o no disponible'
            return resultado
        
        try:
            razon_limpia = self.limpiar_razon_social(razon_social)
            variantes = self.obtener_variantes_busqueda(razon_limpia)
            
            print(f"[Worker {self.worker_id}] Limpiado: {razon_limpia}")
            
            ruc_encontrado = None
            estado_encontrado = config.STATUS['NOT_FOUND']
            variante_exitosa = None
            
            for idx_var, variante in enumerate(variantes, 1):
                if ruc_encontrado:
                    break
                
                if idx_var > 1:
                    print(f"[Worker {self.worker_id}] Variante {idx_var}/{len(variantes)}: {variante}")
                
                self.driver.get(config.SUNAT_URL)
                
                try:
                    tab = self.wait.until(EC.element_to_be_clickable((By.ID, "btnPorRazonSocial")))
                    tab.click()
                except:
                    pass
                
                input_razon = None
                try:
                    input_razon = self.driver.find_element(By.ID, "txtNombreRazonSocial")
                    if not input_razon.is_displayed():
                        input_razon = self.driver.find_element(By.NAME, "search3")
                except:
                    pass
                
                if input_razon and input_razon.is_displayed():
                    try:
                        input_razon.clear()
                        input_razon.send_keys(variante)
                    except Exception as e:
                        print(f"[Worker {self.worker_id}] ERROR escribiendo: {e}")
                
                captcha_visible = False
                try:
                    txt_codigo = self.driver.find_element(By.ID, "txtCodigo")
                    if txt_codigo.is_displayed():
                        captcha_visible = True
                except:
                    pass
                
                if captcha_visible:
                    print(f"[Worker {self.worker_id}] CAPTCHA detectado. Escribelo en Chrome y presiona ENTER aqui...")
                    input()
                
                try:
                    self.driver.find_element(By.ID, "btnAceptar").click()
                    
                    time.sleep(0.5)
                    try:
                        alert = self.driver.switch_to.alert
                        alert_text = alert.text
                        print(f"[Worker {self.worker_id}] Alert: {alert_text}")
                        alert.accept()
                        continue
                    except:
                        pass
                        
                except Exception as e:
                    print(f"[Worker {self.worker_id}] ERROR en Buscar: {e}")
                
                time.sleep(config.PAGE_LOAD_WAIT)
                
                body_text = self.driver.find_element(By.TAG_NAME, "body").text
                rucs_encontrados = self.extraer_todos_los_rucs(body_text)
                
                if not rucs_encontrados:
                    iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                    for frame in iframes:
                        try:
                            self.driver.switch_to.frame(frame)
                            frame_text = self.driver.find_element(By.TAG_NAME, "body").text
                            rucs_frame = self.extraer_todos_los_rucs(frame_text)
                            if rucs_frame:
                                rucs_encontrados.extend(rucs_frame)
                                body_text += " " + frame_text
                            self.driver.switch_to.default_content()
                        except:
                            self.driver.switch_to.default_content()
                
                if rucs_encontrados:
                    ruc_encontrado, estado_encontrado = self.seleccionar_mejor_ruc(
                        rucs_encontrados, 
                        body_text
                    )
                    variante_exitosa = variante
            
            if ruc_encontrado:
                print(f"[Worker {self.worker_id}] RUC: {ruc_encontrado} ({estado_encontrado})")
                resultado['ruc'] = ruc_encontrado
                resultado['estado'] = estado_encontrado
                resultado['observacion'] = 'Exito'
                if variante_exitosa != variantes[0]:
                    resultado['observacion'] += f' (variante: {variante_exitosa})'
            else:
                print(f"[Worker {self.worker_id}] No encontrado")
                resultado['estado'] = config.STATUS['NOT_FOUND']
                resultado['observacion'] = 'No encontrado en web'
        
        except Exception as e:
            error_msg = str(e)
            print(f"[Worker {self.worker_id}] ERROR: {e}")
            
            # Detectar errores críticos de conexión
            errores_criticos = [
                'ERR_CONNECTION_RESET',
                'ERR_INTERNET_DISCONNECTED', 
                'ERR_NAME_NOT_RESOLVED',
                'ERR_CONNECTION_REFUSED',
                'ERR_CONNECTION_TIMED_OUT',
                'ERR_NETWORK_CHANGED',
                'ERR_CONNECTION_CLOSED',
                'Max retries exceeded',
                'Connection refused',
                'Connection reset',
                'No se puede establecer una conexión',
                'Failed to establish a new connection',
                'NewConnectionError',
                'Timeout',
                'timed out',
                'chrome not reachable',
                'Session deleted because of page crash',
                'disconnected: not connected to DevTools',
                'invalid session id'
            ]
            
            es_error_critico = any(error in error_msg for error in errores_criticos)
            
            if es_error_critico:
                print(f"\n{'!'*70}")
                print(f"[Worker {self.worker_id}] ERROR CRITICO DE CONEXION DETECTADO")
                print(f"{'!'*70}")
                resultado['estado'] = 'ERROR_CONEXION'
                resultado['observacion'] = f'ERROR CONEXION: {error_msg}'
            else:
                resultado['estado'] = config.STATUS['ERROR']
                resultado['observacion'] = str(e)
        
        return resultado

