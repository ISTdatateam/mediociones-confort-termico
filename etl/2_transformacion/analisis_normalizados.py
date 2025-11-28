import os
import pandas as pd

# Carpeta donde están los archivos
carpeta = r"C:\Users\Quantum-Malloco\Desktop\Ventilacion\Casos de prueba 3\Normalizados"  # <-- cámbiala si hace falta

resultados = []

for nombre_archivo in os.listdir(carpeta):
    # Filtrar solo archivos Excel
    if not nombre_archivo.lower().endswith(('.xlsx', '.xlsm', '.xls')):
        continue

    ruta_archivo = os.path.join(carpeta, nombre_archivo)
    print(f"Revisando: {ruta_archivo}")

    # Intentar abrir el archivo y detectar la primera hoja
    try:
        xls = pd.ExcelFile(ruta_archivo)
        primera_hoja = xls.sheet_names[0]  # Solo primera hoja
    except Exception as e:
        print(f"  No se pudo abrir el archivo: {e}")
        continue

    # Leer la columna J completa (sin asumir encabezados)
    try:
        df_j = pd.read_excel(
            ruta_archivo,
            sheet_name=primera_hoja,
            usecols="F",
            header=None
        )
    except Exception as e:
        print(f"  Error leyendo la columna F en {nombre_archivo}: {e}")
        continue

    # Filas 91 a 93 de Excel corresponden a índices 90, 91 y 92 en pandas
    for fila_excel in range(91, 94):  # 91, 92, 93
        idx = fila_excel - 1

        if idx < len(df_j):
            valor = df_j.iloc[idx, 0]
        else:
            valor = None  # Por si el archivo no tiene tantas filas

        resultados.append({
            "Archivo": nombre_archivo,
            "Hoja": primera_hoja,
            "Fila": fila_excel,
            "Columna": "F",
            "Valor": valor
        })

        print(f"  Fila {fila_excel}, columna F: {valor}")

# Guardar resultados en un Excel
if resultados:
    df_resultados = pd.DataFrame(resultados)
    salida = os.path.join(carpeta, "resultado_columnaF_filas_91_93.xlsx")
    df_resultados.to_excel(salida, index=False)
    print("\nResumen guardado en:")
    print(salida)
else:
    print("\nNo se pudo obtener información de la columna F en ningún archivo.")
