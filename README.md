# 🔍 Buscador de RUC por Razón Social - SUNAT

Script para buscar RUCs en SUNAT a partir de Razones Sociales (nombres de empresas).

## 🎯 ¿Qué hace este script?

Tienes un Excel con:
- ✅ **Razón Social** (nombre de empresa)
- ❌ **RUC vacío**

El script:
1. Lee las Razones Sociales de tu Excel
2. Busca cada una en SUNAT usando su búsqueda oficial
3. Extrae el RUC encontrado
4. Opcionalmente consulta datos completos (dirección, estado, etc.)
5. Guarda todo en un nuevo Excel con los RUCs encontrados

## 📋 Requisitos

```bash
pip install -r requirements.txt
```

Dependencias:
- pandas
- openpyxl
- requests
- beautifulsoup4
- tqdm
- lxml

## 🚀 Uso

### Paso 1: Test Rápido (OBLIGATORIO)

**Siempre ejecuta primero el test** con solo 3 empresas:

```bash
python test_busqueda.py
```

Esto te mostrará:
- Si encuentra la columna "RAZON SOCIAL" correctamente
- Los primeros 3 resultados de búsqueda
- Si SUNAT responde bien
- Un archivo `test_resultado.xlsx` con los 3 resultados

**⏱️ Tiempo: 30-60 segundos**

### Paso 2: Procesamiento Completo

Si el test funcionó, procesa todas las razones sociales:

```bash
python buscar_ruc_sunat.py
```

Esto procesará todas las razones sociales válidas en `DATA.xlsx`.

**⏱️ Tiempo estimado**: 
- ~2 segundos por empresa
- Para 400 empresas: ~13-15 minutos

## 📊 Qué verás durante la ejecución

```
2025-11-26 10:20:15 - INFO - Archivo cargado: 2212 filas
2025-11-26 10:20:15 - INFO - Razones sociales válidas: 389
2025-11-26 10:20:15 - INFO - Filas vacías omitidas: 1823
2025-11-26 10:20:15 - INFO - Iniciando búsqueda en SUNAT...

Buscando RUCs: 45%|████████████      | 175/389 [05:50<06:40, 2.00s/it]

2025-11-26 10:26:05 - INFO - Buscando: ALICORP S.A.A.
2025-11-26 10:26:07 - INFO -   ✓ RUC encontrado: 20100070970 - ALICORP S.A.A.
```

## 🎛️ Opciones de Configuración

### Opción 1: Solo buscar RUC (Rápido)

Edita la última línea de `buscar_ruc_sunat.py`:

```python
procesar_excel("DATA.xlsx", consultar_completo=False)
```

- ✅ Más rápido (~2 seg por empresa)
- ✅ Solo obtiene RUC y nombre
- ❌ No obtiene dirección, estado, etc.

### Opción 2: Buscar RUC + Datos Completos (Lento)

```python
procesar_excel("DATA.xlsx", consultar_completo=True)
```

- ❌ Más lento (~4 seg por empresa)
- ✅ Obtiene RUC, nombre, dirección, estado, condición, etc.
- ✅ Información completa de SUNAT

## 📁 Archivos Generados

- `resultados_ruc_YYYYMMDD_HHMMSS.xlsx` - Resultados finales
- `temp_resultados_ruc_*.xlsx` - Guardado incremental (se borra al terminar)
- `busqueda_ruc.log` - Log detallado
- `test_resultado.xlsx` - Resultados del test (solo con test_busqueda.py)

## 📊 Estructura del Excel de Salida

El archivo resultante tendrá todas las columnas originales MÁS:

- `ruc` - RUC encontrado
- `razon_social_original` - Lo que buscaste
- `razon_social_encontrada` - Lo que SUNAT devolvió
- `fuente` - Siempre "SUNAT"
- `fecha_busqueda` - Cuándo se hizo la búsqueda

Si usaste `consultar_completo=True`, también tendrás:
- `estado` - ACTIVO/INACTIVO
- `condicion` - HABIDO/NO HABIDO
- `direccion` - Domicilio fiscal
- `departamento`
- `provincia`
- `distrito`

## ⚠️ Notas Importantes

### 1. Razones Sociales deben ser exactas o similares

SUNAT busca por coincidencia. Si tu Excel dice:
- ✅ "ALICORP S.A.A." → Encontrará
- ✅ "ALICORP" → Probablemente encontrará
- ❌ "ALICOR" → Puede no encontrar
- ❌ "Empresa de alimentos" → No encontrará

### 2. Algunas empresas pueden no encontrarse

Razones:
- Nombre muy genérico
- Empresa no registrada en SUNAT
- Nombre escrito diferente en SUNAT
- Empresa dada de baja

### 3. El script filtra automáticamente

- ❌ Filas con razón social vacía
- ❌ Filas con solo espacios
- ✅ Solo procesa razones sociales válidas

### 4. Puedes interrumpir con Ctrl+C

- Los datos procesados estarán en el archivo temporal
- Revisa `busqueda_ruc.log` para ver hasta dónde llegó

## 🔧 Ajustar Velocidad

Si quieres que vaya más rápido (con riesgo de bloqueo), edita en `buscar_ruc_sunat.py`:

```python
# Línea ~27-28
self.timeout = 10  # Reducir de 15 a 10
self.delay = 1     # Reducir de 2 a 1
```

⚠️ **Advertencia**: Muy rápido puede hacer que SUNAT bloquee temporalmente tu IP.

## 🐛 Solución de Problemas

### "No se encontró columna 'RAZON SOCIAL'"

Tu Excel debe tener una columna con "RAZON" o "SOCIAL" en el nombre.

Columnas válidas:
- ✅ "RAZON SOCIAL"
- ✅ "Razón Social"
- ✅ "RAZON_SOCIAL"
- ❌ "NOMBRE" (no la detectará automáticamente)

Si tu columna se llama diferente, edita la línea ~196 en `buscar_ruc_sunat.py`:

```python
if 'razon' in col_lower or 'social' in col_lower or 'nombre' in col_lower:
```

### "No hay razones sociales para procesar"

Todas tus filas están vacías en la columna de razón social.

### Tasa de éxito baja (< 50%)

Posibles causas:
- Nombres muy genéricos
- Nombres con errores ortográficos
- Empresas no registradas
- SUNAT está lento/bloqueando

Revisa `busqueda_ruc.log` para ver qué está pasando.

### El script se queda "colgado"

- Espera hasta 15 segundos por búsqueda (timeout)
- Si SUNAT no responde, pasa automáticamente al siguiente
- Revisa `busqueda_ruc.log` en tiempo real

## 📞 Cómo Funciona Internamente

1. **Lee tu Excel** y encuentra la columna "RAZON SOCIAL"
2. **Filtra filas válidas** (no vacías)
3. Para cada razón social:
   - Hace request a SUNAT con `accion=consPorRazonSoc`
   - Parsea el HTML de respuesta
   - Extrae el RUC usando regex (11 dígitos que empiezan con 10 o 20)
   - Si `consultar_completo=True`, hace segunda consulta con el RUC
4. **Guarda progreso** cada 10 empresas
5. **Genera Excel final** con todos los resultados

## 🎯 Recomendaciones

1. **SIEMPRE ejecuta `test_busqueda.py` primero**
2. **Revisa los 3 resultados del test** antes de procesar todo
3. **Usa `consultar_completo=False`** la primera vez (más rápido)
4. **Si necesitas datos completos**, ejecuta de nuevo con `True`
5. **Revisa el log** si algo no funciona como esperas

## 📈 Tasa de Éxito Esperada

- **Empresas grandes/conocidas**: ~90-95%
- **Empresas medianas**: ~70-80%
- **Empresas pequeñas/informales**: ~40-60%
- **Nombres genéricos**: ~20-30%

La tasa depende mucho de qué tan exactos sean los nombres en tu Excel.

## 🚀 Ejemplo de Uso Completo

```bash
# 1. Instalar dependencias
pip install -r requirements.txt

# 2. Test rápido (3 empresas)
python test_busqueda.py

# 3. Si funcionó, procesar todo
python buscar_ruc_sunat.py

# 4. Revisar resultados
# Abre: resultados_ruc_YYYYMMDD_HHMMSS.xlsx
```

¡Listo! 🎉
