import os
import pandas as pd
from google import genai
import json
import openpyxl
from dotenv import load_dotenv  # <-- 1. Importar dotenv

# Cargar las variables del .env
load_dotenv()  # <-- 2. Inicializar la carga

# 1. Configura tu llave leyendo desde el .env
API_KEY = os.getenv("GEMINI_API_KEY")  # <-- 3. Leer la variable segura
cliente = genai.Client(api_key=API_KEY)


# Fíjate que ahora recibe "moneda_indicada"
def obtener_datos_json(ruta_archivo, moneda_indicada):
    print(f"1. Leyendo y aplastando el Excel: {ruta_archivo}...")

    try:
        # Abrimos el Excel normal
        wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
        ws = wb.active

        data = ws.values
        df = pd.DataFrame(data)
        texto_crudo = df.to_csv(index=False, header=False, sep='\t')

        print("2. Enviando datos a Gemini 2.5 para extracción...\n")

        prompt = f"""
        Eres un sistema avanzado de extracción y auditoría para un astillero naval.
        Analiza el siguiente texto en crudo de una cotización y devuelve UNICAMENTE un JSON.

        Estructura obligatoria:
        {{
            "cliente": "Nombre del cliente o empresa",
            "puerto": "Lugar o puerto del trabajo",
            "actividad": "Tipo de actividad (ej: CALDERERIA)",
            "embarcacion": "Nombre del barco",
            "numero_cotizacion": "Número si existe",
            "fecha": "Fecha del documento",
            "moneda": "{moneda_indicada}", 
            "subtotal_excel": 0.0,
            "igv_excel": 0.0,
            "total_excel": 0.0,
            "items": [
                {{
                    "detalle_actividad": "Descripción",
                    "cantidad": 0.0,
                    "precio_unitario": 0.0,
                    "total_item": 0.0
                }}
            ]
        }}

        Reglas:
        - subtotal_excel, igv_excel y total_excel deben ser los montos finales declarados al final del documento.
        - Los ítems deben ser SOLO los trabajos, ignora filas de totales.
        - Todo monto económico debe ser float.
        - REGLA ABSOLUTA DE MONEDA: El usuario ya confirmó que este documento está en {moneda_indicada}. EL CAMPO "moneda" DEL JSON DEBE SER ESTRICTAMENTE "{moneda_indicada}". PROHIBIDO ADIVINAR OTRA MONEDA.
        - PROHIBIDO USAR TRAILING COMMAS: El JSON debe ser 100% válido y estricto. NO agregues una coma al final del último elemento de una lista o diccionario.

        Texto del Excel:
        {texto_crudo}
        """
        respuesta = cliente.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt
        )

        texto_respuesta = respuesta.text.strip()
        if texto_respuesta.startswith("```json"):
            texto_respuesta = texto_respuesta[7:-3].strip()
        elif texto_respuesta.startswith("```"):
            texto_respuesta = texto_respuesta[3:-3].strip()

        # --- 🛡️ ESCUDO ANTI-COMAS TRAMPOSAS DE GEMINI ---
        import re
        # Esta brujería busca cualquier coma que esté solita antes de un } o un ] y la extermina
        texto_respuesta = re.sub(r',\s*([\]}])', r'\1', texto_respuesta)

        # Ahora sí, Python lo lee sin llorar
        datos_limpios = json.loads(texto_respuesta)

        # --- LÓGICA DE AUDITORÍA (LO QUE PEDISTE) ---

        # 1. Calculamos la suma real de los items
        subtotal_calculado = sum(item['total_item'] for item in datos_limpios.get('items', []))

        # 2. Calculamos el IGV (18%)
        igv_calculado = subtotal_calculado * 0.18

        # 3. Calculamos el Total final
        total_calculado = subtotal_calculado + igv_calculado

        # --- IMPRESIÓN DE RESULTADOS ---
        print("--- AUDITORÍA DE TOTALES ---")

        # Formateo visual para comparar
        print(f"SUBTOTAL:")
        print(f"Declarado en Excel: {datos_limpios.get('subtotal_excel', 0):,.2f}")
        print(f"Suma Real Python:   {subtotal_calculado:,.2f}")

        print(f"\nI.G.V (18%):")
        print(f"Declarado en Excel: {datos_limpios.get('igv_excel', 0):,.2f}")
        print(f"Cálculo Real:       {igv_calculado:,.2f}")

        print(f"\nTOTAL GENERAL:")
        print(f"Declarado en Excel: {datos_limpios.get('total_excel', 0):,.2f}")
        print(f"Cálculo Real:       {total_calculado:,.2f}")

        # Verificación automática
        diferencia = abs(total_calculado - datos_limpios.get('total_excel', 0))
        print("\n--- ESTADO DEL PRESUPUESTO ---")
        if diferencia < 1.0:  # Margen de error por redondeos de centavos
            print("EXCEL PERFECTO: Los totales cuadran matemáticamente.")
        else:
            print(
                f"ALERTA DE CUADRE: Hay una diferencia de {diferencia:,.2f} en el archivo. ¡Revisar fórmulas de Excel!")

        # Opcional: Mostrar el JSON completo si quieres verlo
        print("\nJSON COMPLETO PARA LA BASE DE DATOS:")
        print(json.dumps(datos_limpios, indent=4, ensure_ascii=False))

        # ---> ESTA ES LA LÍNEA MÁGICA QUE TE FALTA <---
        return datos_limpios

    except Exception as e:
        print(f"Error en la Matrix: {e}")
        return None


if __name__ == "__main__":
    mi_archivo = "COTIZACION GAVILAN  13 ABRIL.xlsx"
    obtener_datos_json(mi_archivo)