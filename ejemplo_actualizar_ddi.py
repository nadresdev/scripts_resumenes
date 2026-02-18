
import pandas as pd
import os
from datetime import datetime

# Rutas - Ajustar según convenga
ARCHIVO_ORIGEN = r'C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\DDI\ORIGEN\DDI2 (3).xlsx'
ARCHIVO_DESTINO_DIR = r'C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\DDI\DESTINO'

# Crear carpeta destino si no existe
if not os.path.exists(ARCHIVO_DESTINO_DIR):
    os.makedirs(ARCHIVO_DESTINO_DIR)
    print(f"Directorio creado: {ARCHIVO_DESTINO_DIR}")

def actualizar_ddi(input_path, output_dir=None):
    """
    Lee el archivo DDI, realiza actualizaciones de ejemplo, y guarda una copia.
    """
    print(f"Leyendo archivo: {input_path}...")
    try:
        # Cargar con openpyxl engine
        df = pd.read_excel(input_path, engine='openpyxl')
        
        # --- EJEMPLO DE ACTUALIZACIÓN ---
        # 1. Calcular nueva columna 'Dias_Desde_Alta' si existe 'Fecha alta contrato'
        # Convertir a datetime la columna J (aprox) si tiene nombre 'Fecha alta OCM' o similar
        
        # Normalizar nombres de columnas (strip espacios)
        df.columns = df.columns.str.strip()
        
        print("Columnas encontradas:", df.columns.tolist())
        
        # Ejemplo: Actualizar 'Dias en uso' basado en la fecha actual
        # Asumimos que la columna 'Fecha alta OCM' es numérica (Excel serial date) o datetime
        # Pandastools often interpret Excel dates as float/int if not specified.
        # Aquí solo mostraremos info básica y guardaremos para demostrar el flujo.
        
        info_rows = len(df)
        print(f"Total filas cargadas: {info_rows}")
        
        # Simulamos una actualización: Añadir columna de fecha de procesamiento
        df['Fecha_Actualizacion'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Generar nombre de archivo de salida
        nombre_base = os.path.basename(input_path).split('.')[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{nombre_base}_ACTUALIZADO_{timestamp}.xlsx"
        
        if output_dir:
            output_path = os.path.join(output_dir, output_filename)
        else:
            output_path = os.path.join(os.path.dirname(input_path), output_filename)
            
        print(f"Guardando archivo actualizado en: {output_path}...")
        df.to_excel(output_path, index=False, engine='openpyxl')
        print("¡Proceso completado con éxito!")
        return output_path

    except Exception as e:
        print(f"Error durante el proceso: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    if os.path.exists(ARCHIVO_ORIGEN):
        actualizar_ddi(ARCHIVO_ORIGEN, ARCHIVO_DESTINO_DIR)
    else:
        print(f"No se encontró el archivo origen: {ARCHIVO_ORIGEN}")
