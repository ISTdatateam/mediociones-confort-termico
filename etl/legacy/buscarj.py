import os
import pandas as pd

# Carpeta donde están los archivos
carpeta = r"C:\Users\Quantum-Malloco\Desktop\Ventilacion\Casos de prueba"

# Texto a buscar (búsqueda parcial)
texto_buscar = "Identificación de aperturas por donde ingresa"
texto_normalizado = texto_buscar.strip().casefold()

resultados = []

for nombre_archivo in os.listdir(carpeta):
    if not nombre_archivo.lower().endswith(('.xlsx', '.xlsm', '.xls')):
        continue

    ruta_archivo = os.path.join(carpeta, nombre_archivo)
    print(f"Revisando: {ruta_archivo}")

    try:
        xls = pd.ExcelFile(ruta_archivo)
        primera_hoja = xls.sheet_names[0]   # SOLO primera hoja
    except Exception as e:
        print(f"  No se pudo abrir el archivo: {e}")
        continue

    try:
        # Leer columnas I:L (para cubrir celdas fusionadas I-J-K-L)
        df_cols = pd.read_excel(
            ruta_archivo,
            sheet_name=primera_hoja,
            usecols="F:L",
            header=None
        )
    except ValueError:
        print("  La primera hoja no tiene columnas I:L.")
        continue
    except Exception as e:
        print(f"  Error leyendo la primera hoja en {nombre_archivo}: {e}")
        continue

    # Normalizar: convertir a texto, quitar espacios, pasar a minúsculas "robustas"
    df_norm = (
        df_cols
        .fillna("")
        .astype(str)
        .apply(lambda col: col.str.strip().str.casefold())
    )

    # Crear máscara booleana: TRUE si la celda contiene el texto buscado (parcial)
    df_mask = df_norm.apply(lambda col: col.str.contains(texto_normalizado, na=False))

    # Para cada fila, ver si al menos una columna I/J/K/L tiene el texto
    mask_filas = df_mask.any(axis=1)

    # Índices de filas donde se encontró el texto
    indices_coinciden = mask_filas[mask_filas].index

    if len(indices_coinciden) > 0:
        idx_primero = indices_coinciden[0]
        fila_excel = idx_primero + 1  # ajustar a numeración de Excel

        # Identificar en qué columnas se encontró (I, J, K, L)
        letras_columnas = ["F", "G", "H","I", "J", "K", "L"]
        mapeo_col = {df_cols.columns[i]: letras_columnas[i]
                     for i in range(len(df_cols.columns))}

        columnas_en_fila = []
        for col in df_cols.columns:
            if df_mask.loc[idx_primero, col]:
                columnas_en_fila.append(mapeo_col[col])

        columnas_str = ", ".join(columnas_en_fila) if columnas_en_fila else ""

        resultados.append({
            "Archivo": nombre_archivo,
            "Hoja": primera_hoja,
            "Fila": fila_excel,
            "Columnas_donde_aparece": columnas_str
        })

        print(f"  >>> Encontrado en fila {fila_excel}, columna(s): {columnas_str}")
    else:
        print("  No se encontró el texto en la primera hoja (columnas F:L).")

# Guardar resultados
if resultados:
    df_resultados = pd.DataFrame(resultados)
    salida = os.path.join(carpeta, "resultado_busqueda_IL_parcial.xlsx")
    df_resultados.to_excel(salida, index=False)
    print("\nResumen guardado en:")
    print(salida)
else:
    print("\nNo se encontró el texto en ningún archivo.")
