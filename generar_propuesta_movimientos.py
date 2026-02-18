
import pandas as pd
import os
from datetime import datetime

INPUT_FILE = r'C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\DDI\ORIGEN\DDI2 (3).xlsx'
OUTPUT_FILE = r'C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\DDI\PROPUESTA_MOVIMIENTOS_HOY.xlsx'

def generar_propuesta():
    print(f"Leyendo archivo origen: {INPUT_FILE}...")
    try:
        df = pd.read_excel(INPUT_FILE, engine='openpyxl')
        
        # Limpieza de columnas
        df.columns = df.columns.str.strip()
        
        # Renombrar primera columna a Telefono si es necesario
        if df.columns[0].startswith('Unnamed'):
            df.rename(columns={df.columns[0]: 'Telefono'}, inplace=True)
            
        # Asegurar tipos
        df['Telefono'] = df['Telefono'].astype(str).str.strip()
        
        # --- 1. PROPUESTA DE ACTIVACIÓN (ALTA -> EMITIENDO) ---
        # Criterio: Estado = 'ALTA'. Orden: Tal cual aparecen (FIFO)
        cond_activar = (df['Estado'] == 'ALTA')
        candidatos_activar = df[cond_activar].copy()
        
        # Seleccionamos TOP 30 (o todos si son menos)
        top_n_activar = 30
        propuesta_activar = candidatos_activar.head(top_n_activar)
        
        print(f"Candidatos para ACTIVAR encontrados: {len(candidatos_activar)}")
        print(f"Seleccionados para propuesta: {len(propuesta_activar)}")
        
        # Añadir columna con la acción sugerida
        propuesta_activar['ACCION_SUGERIDA'] = 'CAMBIAR A EMITIENDO'
        
        
        # --- 2. PROPUESTA DE ROTACIÓN (LEADS -> LEADS-SPIN) ---
        # Criterio: MOTOR = 'LEADS', Veces usado <= 1, Dias en uso > 30 (aprox, para no rotar cosas de ayer)
        # Nota: Si 'Dias en uso' tiene nulos, los trataremos como 0 o ignoraremos
        
        # Asegurar que 'Dias en uso' es numérico
        df['Dias en uso'] = pd.to_numeric(df['Dias en uso'], errors='coerce').fillna(0)
        df['Veces usado'] = pd.to_numeric(df['Veces usado'], errors='coerce').fillna(0)
        
        cond_rotar = (
            (df['MOTOR'] == 'LEADS') & 
            (df['Veces usado'] >= 1) & 
            (df['Dias en uso'] > 20) & # Umbral arbitrario de "madurez"
            (df['Estado'].isin(['EMITIENDO', 'BAJA OCM', 'BAJA'])) # Estados validos para rotar
        )
        
        candidatos_rotar = df[cond_rotar].copy()
        
        # Priorizar los que llevan más tiempo en uso (más antiguos)
        candidatos_rotar = candidatos_rotar.sort_values('Dias en uso', ascending=False)
        
        # Seleccionamos TOP 50
        top_n_rotar = 50
        propuesta_rotar = candidatos_rotar.head(top_n_rotar)

        print(f"Candidatos para ROTAR encontrados: {len(candidatos_rotar)}")
        print(f"Seleccionados para propuesta: {len(propuesta_rotar)}")
        
        propuesta_rotar['ACCION_SUGERIDA'] = 'CAMBIAR MOTOR A LEADS-SPIN'
        
        
        # --- GUARDAR EN EXCEL ---
        print(f"Generando archivo de propuesta: {OUTPUT_FILE}...")
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            propuesta_activar.to_excel(writer, sheet_name='PROPUESTA_ACTIVAR', index=False)
            propuesta_rotar.to_excel(writer, sheet_name='PROPUESTA_ROTAR', index=False)
            
        print("¡Archivo generado exitosamente!")
        return OUTPUT_FILE

    except Exception as e:
        print(f"Error generando propuesta: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    generar_propuesta()
