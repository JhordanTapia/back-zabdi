import pandas as pd


def auditar_excel_terco(ruta_archivo):
    print("Abriendo el Excel en modo a prueba de balas...\n")

    try:
        # header=7 salta directo a la fila 8
        df = pd.read_excel(ruta_archivo, header=7)

        print("--- MAPEO EXACTO DE COLUMNAS PARA BD ---")
        print("-" * 80)
        print(f"{'COLUMNA EXCEL'.ljust(15)} | {'TIPO PARA BD'.ljust(18)} | {'EJEMPLO DEL DATO REAL'}")
        print("-" * 80)

        for col in df.columns:
            # Agarramos los datos ignorando los vacíos
            muestra = df[col].dropna()

            if muestra.empty:
                continue  # Ignoramos las columnas que sí están 100% vacías

            # Sacamos el primer valor real de la fila 9 para que veas qué es
            primer_valor = muestra.iloc[0]

            # Determinamos el tipo
            tipo_pandas = df[col].dtype
            if pd.api.types.is_integer_dtype(tipo_pandas):
                tipo_sql = "INTEGER"
            elif pd.api.types.is_float_dtype(tipo_pandas):
                tipo_sql = "DOUBLE PRECISION"
            else:
                max_len = muestra.astype(str).map(len).max()
                tipo_sql = "TEXT" if max_len > 255 else f"VARCHAR({int(max_len) + 50})"

            # Limpiamos los saltos de línea del ejemplo para que se vea ordenado
            ejemplo_limpio = str(primer_valor).replace('\n', ' ')
            # Cortamos el texto si es muy largo para que no rompa la consola
            if len(ejemplo_limpio) > 40:
                ejemplo_limpio = ejemplo_limpio[:37] + "..."

            print(f"{str(col).ljust(15)} | {tipo_sql.ljust(18)} | {ejemplo_limpio}")

        print("-" * 80)

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    mi_archivo = "COTIZACION GAVILAN  13 ABRIL.xlsx"
    auditar_excel_terco(mi_archivo)