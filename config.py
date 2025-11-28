import os
from dotenv import load_dotenv

load_dotenv()

INPUT_FILE = os.getenv('INPUT_FILE', 'DATA.xlsx')
OUTPUT_FILE = os.getenv('OUTPUT_FILE', 'RESULTADOS_FINALES.xlsx')

CHROMEDRIVER_PATHS = [
    os.path.expanduser("~/.chromedriver/chromedriver.exe"),
    "C:/chromedriver/chromedriver.exe",
    "chromedriver.exe"
]

NUM_WORKERS = int(os.getenv('NUM_WORKERS', 5))
BATCH_SIZE = int(os.getenv('BATCH_SIZE', 5))
DELAY_BETWEEN_BATCHES = float(os.getenv('DELAY_BETWEEN_BATCHES', 0.5))

SELENIUM_TIMEOUT = int(os.getenv('SELENIUM_TIMEOUT', 10))
PAGE_LOAD_WAIT = float(os.getenv('PAGE_LOAD_WAIT', 3))
HEADLESS_MODE = os.getenv('HEADLESS_MODE', 'false').lower() == 'true'

SUNAT_URL = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias"

OUTPUT_COLUMNS = [
    'indice_original',
    'razon_social_input',
    'ruc',
    'estado',
    'observacion',
    'direccion_original',
    'numero_original',
    'worker_id'
]

STATUS = {
    'PENDING': 'PENDIENTE',
    'PROCESSING': 'PROCESANDO',
    'COMPLETED': 'COMPLETADO',
    'ERROR': 'ERROR',
    'NOT_FOUND': 'NO ENCONTRADO'
}

SUFIJOS_EMPRESAS = [
    ' SAC', ' EIRL', ' SRL', ' SAA', ' SA',
    ' SOCIEDAD ANONIMA CERRADA',
    ' SOCIEDAD ANONIMA',
    ' EMPRESA INDIVIDUAL DE RESPONSABILIDAD LIMITADA',
    ' SCRL', ' SL', ' LTDA'
]
