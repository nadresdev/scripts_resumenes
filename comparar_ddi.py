
import pandas as pd
import os
import sys

def comparar_ddis(path_old, path_new):
    print(f"--- COMPARANDO ARCHIVOS ---")
    print(f"ANTERIOR: {path_old}")
    print(f"ACTUAL:   {path_new}")

    if not os.path.exists(path_old) or not os.path.exists(path_new):
        print("Error: Uno o ambos archivos no existen.")
        return

    try:
        # Cargar archivos
        df_old = pd.read_excel(path_old, engine='openpyxl')
        df_new = pd.read_excel(path_new, engine='openpyxl')

        # Normalizar nombres de columnas (strip)
        df_old.columns = df_old.columns.str.strip()
        df_new.columns = df_new.columns.str.strip()

        # Asumir columna 0 es 'Telefono' si no se llama así
        col_old_0 = df_old.columns[0]
        col_new_0 = df_new.columns[0]
        
        # Renombrar para consistencia
        df_old.rename(columns={col_old_0: 'Telefono'}, inplace=True)
        df_new.rename(columns={col_new_0: 'Telefono'}, inplace=True)

        # Convertir a string para asegurar match exacto
        df_old['Telefono'] = df_old['Telefono'].astype(str).str.strip()
        df_new['Telefono'] = df_new['Telefono'].astype(str).str.strip()

        # Sets de teléfonos
        set_old = set(df_old['Telefono'])
        set_new = set(df_new['Telefono'])

        # 1. Nuevos y Perdidos
        added = set_new - set_old
        removed = set_old - set_new
        common = set_new.intersection(set_old)

        print(f"\n--- RESULTADOS GENERALES ---")
        print(f"Total en ANTERIOR: {len(set_old)}")
        print(f"Total en ACTUAL:   {len(set_new)}")
        print(f"Comunes:           {len(common)}")
        print(f"NUEVOS (Agregados): {len(added)}")
        print(f"ELIMINADOS:        {len(removed)}")

        # 2. Análisis de Cambios en Comunes (Estado y Uso)
        print(f"\n--- ANÁLISIS DE CAMBIOS (En {len(common)} registros comunes) ---")
        
        # Merge para comparar
        df_merged = pd.merge(df_old, df_new, on='Telefono', suffixes=('_OLD', '_NEW'), how='inner')
        
        cols_to_compare = ['Estado', 'Dias en uso', 'Veces usado', 'MOTOR']
        
        for col in cols_to_compare:
            col_old = f"{col}_OLD"
            col_new = f"{col}_NEW"
            
            if col_old in df_merged.columns and col_new in df_merged.columns:
                # Detectar cambios
                changes = df_merged[df_merged[col_old] != df_merged[col_new]]
                num_changes = len(changes)
                
                print(f"\nCambios en '{col}': {num_changes}")
                if num_changes > 0:
                    # Mostrar top cambios (ej. BAJA -> ALTA)
                    change_counts = changes.groupby([col_old, col_new]).size().reset_index(name='Count')
                    change_counts = change_counts.sort_values('Count', ascending=False).head(10)
                    print(change_counts.to_string(index=False))

    except Exception as e:
        print(f"Error en comparación: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 2:
        file_old = sys.argv[1]
        file_new = sys.argv[2]
    else:
        # Default placeholder paths - User needs to update these or pass args
        file_new = r'C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\DDI\ORIGEN\DDI2 (3).xlsx'
        file_old = r'C:\Users\dresdev\OneDrive\Desktop\SMART CONECT\DDI\ORIGEN\DDI2_ANTERIOR.xlsx' # Placeholder name

    comparar_ddis(file_old, file_new)
