
import pandas as pd
import numpy as np
import os

FILE_OLD = r'C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\DDI\ANTERIOR DDI3.xlsx'
FILE_NEW = r'C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\DDI\ORIGEN\DDI2 (3).xlsx'

def analizar_patrones_y_predecir():
    print("--- 1. IDENTIFICANDO PATRONES DE MOVIMIENTO ---")
    try:
        df_old = pd.read_excel(FILE_OLD, engine='openpyxl')
        df_new = pd.read_excel(FILE_NEW, engine='openpyxl')
        
        # Normalizar
        df_old.columns = df_old.columns.str.strip()
        df_new.columns = df_new.columns.str.strip()
        
        # Renombrar primera col a Telefono
        if 'Telefono' not in df_old.columns: df_old.rename(columns={df_old.columns[0]: 'Telefono'}, inplace=True)
        if 'Telefono' not in df_new.columns: df_new.rename(columns={df_new.columns[0]: 'Telefono'}, inplace=True)
        
        df_old['Telefono'] = df_old['Telefono'].astype(str).str.strip()
        df_new['Telefono'] = df_new['Telefono'].astype(str).str.strip()
        
        # Merge
        merged = pd.merge(df_old, df_new, on='Telefono', suffixes=('_OLD', '_NEW'))
        
        # --- PATRÓN 1: Activación (ALTA -> EMITIENDO)
        # Queremos saber: ¿Qué características tenían en _OLD los que pasaron a EMITIENDO?
        cambio_activacion = merged[
            (merged['Estado_OLD'] == 'ALTA') & 
            (merged['Estado_NEW'] == 'EMITIENDO')
        ]
        
        print(f"\n[Patrón: ALTA -> EMITIENDO]")
        print(f"Casos detectados: {len(cambio_activacion)}")
        if len(cambio_activacion) > 0:
            avg_dias = cambio_activacion['Dias en uso_OLD'].mean()
            print(f"Promedio 'Dias en uso' antes de activarse: {avg_dias:.1f}")
            # Ver si hay algún orden (ej. los más antiguos primero)
            # Asumiremos que se activan los que llevan X días en espera
            prioridad_activacion = avg_dias
        else:
            prioridad_activacion = None

        # --- PATRÓN 2: Rotación (LEADS -> LEADS-SPIN)
        cambio_spin = merged[
            (merged['MOTOR_OLD'] == 'LEADS') & 
            (merged['MOTOR_NEW'] == 'LEADS-SPIN')
        ]
        
        print(f"\n[Patrón: LEADS -> LEADS-SPIN]")
        print(f"Casos detectados: {len(cambio_spin)}")
        umbral_spin = 0
        if len(cambio_spin) > 0:
            # Analizar 'Dias en uso' previos al cambio
            dias_antes_spin = cambio_spin['Dias en uso_OLD'].describe()
            print(f"Estadísticas 'Dias en uso' al rotar:\n{dias_antes_spin}")
            umbral_spin = dias_antes_spin['min'] # Tomamos el mínimo como trigger conservador
            print(f"-> Umbral detectado para rotar a SPIN: {umbral_spin} días.")


        # --- 2. PREDICCIÓN PARA HOY ---
        print("\n--- 2. CANDIDATOS PARA MOVERSE HOY (Basado en Patrones) ---")
        
        # Candidatos a ACTIVAR (ALTA -> EMITIENDO)
        # Buscamos en el archivo NUEVO (estado actual) quiénes están en ALTA
        candidatos_activar = df_new[df_new['Estado'] == 'ALTA'].copy()
        if not candidatos_activar.empty:
            # Ordenar por antigüedad (Simulando FIFO/Prioridad encontrada)
            candidatos_activar = candidatos_activar.sort_values('Dias en uso', ascending=False)
            print(f"\n[Candidatos a ACTIVAR hoy (ALTA -> EMITIENDO)]")
            print(f"Total disponibles: {len(candidatos_activar)}")
            print("Top 5 sugerecias (los más antiguos en espera):")
            print(candidatos_activar[['Telefono', 'Dias en uso', 'Estado']].head(5).to_string(index=False))
            
        # Candidatos a ROTAR (LEADS -> LEADS-SPIN)
        # Buscamos en el actual quiénes están en LEADS y superan el umbral
        if umbral_spin > 0:
            candidatos_spin = df_new[
                (df_new['MOTOR'] == 'LEADS') & 
                (df_new['Dias en uso'] >= umbral_spin) &
                (df_new['Estado'] != 'BAJA') # Solo activos o emitiendo
            ].copy()
            
            print(f"\n[Candidatos a ROTAR hoy (LEADS -> LEADS-SPIN)]")
            print(f"Criterio: MOTOR='LEADS' y Dias en uso >= {umbral_spin}")
            print(f"Total candidatos: {len(candidatos_spin)}")
            if len(candidatos_spin) > 0:
                print("Top 5 para rotar (Mayor antigüedad):")
                print(candidatos_spin[['Telefono', 'Dias en uso', 'MOTOR']].sort_values('Dias en uso', ascending=False).head(5).to_string(index=False))
        else:
            print("\nNo hay suficientes datos históricos para predecir rotación a SPIN hoy.")

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analizar_patrones_y_predecir()
