import pandas as pd

def procesar_excel(entrada, salida):
    df = pd.read_excel(entrada)
    # Aquí iría tu lógica real de procesamiento
    df.to_excel(salida, index=False)
