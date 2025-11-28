# ScrapingSunat - Extractor Paralelo de RUCs

Extrae RUCs desde SUNAT usando procesamiento paralelo con 5 workers.

## Instalacion

```bash
pip install -r requirements.txt
```

## Uso

### Opcion 1: Limpiar duplicados primero (RECOMENDADO)

Si tu Excel tiene muchos duplicados consecutivos:

```bash
python limpiar_duplicados.py DATA.xlsx
```

Esto creara `DATA_LIMPIO.xlsx` sin duplicados consecutivos.

### Opcion 2: Procesar directamente

El script detecta y maneja duplicados automaticamente:

```bash
python procesar_sunat_paralelo.py
```

El sistema:
1. Detecta razones sociales consecutivas duplicadas
2. Procesa solo UNA vez cada razon social unica
3. Replica automaticamente el resultado a todos los duplicados
4. Ahorra tiempo evitando busquedas repetidas

## Configuracion

Edita `.env` para cambiar parametros:
- `NUM_WORKERS`: Numero de workers paralelos (default: 5)
- `BATCH_SIZE`: Registros por batch (default: 5)
- `INPUT_FILE`: Archivo de entrada (default: DATA.xlsx)
- `OUTPUT_FILE`: Archivo de salida (default: RESULTADOS_FINALES.xlsx)
- `HEADLESS_MODE`: Ejecutar Chrome sin ventanas (default: true)

**IMPORTANTE**: Modo headless esta ACTIVADO por defecto para evitar sobrecarga.
Si quieres ver las ventanas de Chrome, edita `.env` y cambia:
```
HEADLESS_MODE=false
```

## Caracteristicas

- **Deduplicacion automatica**: Detecta y procesa solo una vez razones sociales consecutivas duplicadas
- **Procesamiento paralelo**: 5 Chrome simultaneos para maxima velocidad
- **Guardado automatico**: Cada 30 segundos
- **Busqueda progresiva**: 100%, 75%, 50% del nombre
- **Limpieza automatica**: Elimina caracteres especiales
- **Recuperacion de progreso**: Si se interrumpe, continua donde quedo
- **Manejo de CAPTCHA**: Pausa para resolver manualmente

## Estructura

```
ScrapingSunat/
├── config.py                  # Configuracion
├── procesar_sunat_paralelo.py # Script principal
├── modules/
│   ├── excel_manager.py       # Manejo de Excel
│   └── sunat_scraper.py       # Scraper de SUNAT
├── DATA.xlsx                  # Input
└── RESULTADOS_FINALES.xlsx    # Output
```

## Notas

- Requiere ChromeDriver instalado
- Se abriran 5 ventanas de Chrome simultaneamente
- El proceso puede pausarse con Ctrl+C (progreso guardado)
- Si aparece CAPTCHA, resuelve manualmente en Chrome
