
import pandas as pd
import matplotlib.pyplot as plt
import os

FILE_PATH = r'C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\DDI\ORIGEN\DDI2 (3).xlsx'

def analyze_ddi_usage(file_path):
    if not os.path.exists(file_path):
        print(f"Archivo no encontrado: {file_path}")
        return

    print(f"--- ANALIZANDO: {os.path.basename(file_path)} ---")
    
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        
        # Renombrar primera columna a 'Telefono' si no tiene nombre
        if df.columns[0].startswith('Unnamed'):
            df.rename(columns={df.columns[0]: 'Telefono'}, inplace=True)
            
        print(f"Total Registros: {len(df)}")
        print(f"Columnas: {df.columns.tolist()}")
        
        # 1. Análisis de Estado (Status)
        if 'Estado' in df.columns:
            print("\n--- DISTRIBUCIÓN DE ESTADOS ---")
            print(df['Estado'].value_counts(dropna=False))
            
        # 2. Análisis de Uso (Días y Veces)
        print("\n--- ESTADÍSTICAS DE USO ---")
        cols_uso = ['Dias en uso', 'Días des-uso', 'Veces usado', 'Contacto vs llamadas']
        for col in cols_uso:
            if col in df.columns:
                print(f"\nResumen para '{col}':")
                print(df[col].describe())

        # 3. Análisis de Troncales/Motores
        print("\n--- TRONCALES Y MOTORES ---")
        if 'Troncal' in df.columns:
            print("\nTop 5 Troncales:")
            print(df['Troncal'].value_counts().head(5))
        if 'MOTOR' in df.columns:
            print("\nTop Motores:")
            print(df['MOTOR'].value_counts())

        # 4. Duplicados
        dupes = df.duplicated(subset=['Telefono']).sum()
        print(f"\n--- DUPLICADOS ---")
        print(f"Número de teléfonos duplicados: {dupes}")

        # 5. Fechas (Convertir a datetime si es posible para ver rango)
        # Excel dates are often floats. Need conversion logic if strictly needed, 
        # but for overview describing the range is good.
        # Simple check for now.
        
    except Exception as e:
        print(f"Error analizando: {e}")

if __name__ == "__main__":
    analyze_ddi_usage(FILE_PATH)
