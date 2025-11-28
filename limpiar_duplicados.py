import pandas as pd
import sys

def limpiar_duplicados_consecutivos(archivo_entrada, archivo_salida=None):
    """
    Elimina duplicados consecutivos de un Excel.
    Mantiene solo la primera ocurrencia de cada razon social consecutiva.
    """
    if archivo_salida is None:
        archivo_salida = archivo_entrada.replace('.xlsx', '_LIMPIO.xlsx')
    
    print("="*70)
    print("LIMPIADOR DE DUPLICADOS CONSECUTIVOS")
    print("="*70)
    
    try:
        df = pd.read_excel(archivo_entrada)
        print(f"\nArchivo cargado: {archivo_entrada}")
        print(f"Total registros: {len(df)}")
        
        col_razon = None
        for col in df.columns:
            if 'razon' in col.lower() or 'social' in col.lower():
                col_razon = col
                break
        
        if not col_razon:
            print("ERROR: No se encontro columna de Razon Social")
            return
        
        print(f"Columna detectada: {col_razon}")
        
        indices_a_mantener = []
        razon_anterior = None
        duplicados_eliminados = 0
        
        for idx, row in df.iterrows():
            razon_actual = str(row[col_razon]).strip().upper()
            
            if razon_actual != razon_anterior:
                indices_a_mantener.append(idx)
                razon_anterior = razon_actual
            else:
                duplicados_eliminados += 1
        
        df_limpio = df.loc[indices_a_mantener]
        
        print(f"\nResultados:")
        print(f"  Registros originales: {len(df)}")
        print(f"  Duplicados eliminados: {duplicados_eliminados}")
        print(f"  Registros finales: {len(df_limpio)}")
        print(f"  Reduccion: {duplicados_eliminados/len(df)*100:.1f}%")
        
        df_limpio.to_excel(archivo_salida, index=False)
        print(f"\nArchivo guardado: {archivo_salida}")
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        archivo = sys.argv[1]
    else:
        archivo = "DATA.xlsx"
    
    limpiar_duplicados_consecutivos(archivo)
