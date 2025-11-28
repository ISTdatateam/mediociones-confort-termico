import os
import pandas as pd

# Carpeta donde están los archivos ORIGINALES
carpeta = r"C:\Users\Quantum-Malloco\Desktop\Ventilacion\Casos de prueba 3\Normalizados"

# Carpeta donde se guardarán los archivos PROCESADOS (con filas 91-93 eliminadas)
carpeta_procesadas = os.path.join(carpeta, "procesadas")
os.makedirs(carpeta_procesadas, exist_ok=True)  # Crea la carpeta si no existe

resultados = []

for nombre_archivo in os.listdir(carpeta):

    # Solo archivos Excel
    if not nombre_archivo.lower().endswith(('.xlsx', '.xlsm', '.xls')):
        continue

    ruta_archivo = os.path.join(carpeta, nombre_archivo)
    print(f"\nRevisando: {ruta_archivo}")

    # Abrir archivo y detectar primera hoja
    try:
        xls = pd.ExcelFile(ruta_archivo)
        primera_hoja = xls.sheet_names[0]
    except Exception as e:
        print(f"  No se pudo abrir el archivo: {e}")
        continue

    # Leer columna F completa (sin asumir encabezados)
    try:
        df_F = pd.read_excel(
            ruta_archivo,
            sheet_name=primera_hoja,
            usecols="F",
            header=None
        )
    except Exception as e:
        print(f"  Error leyendo la columna F en {nombre_archivo}: {e}")
        continue

    # Leer celdas F91, F92 y F93
    valores_F = {}
    for fila_excel in range(91, 94):  # 91, 92, 93
        idx = fila_excel - 1
        valor = df_F.iloc[idx, 0] if idx < len(df_F) else None
        valores_F[fila_excel] = valor

        print(f"  Fila {fila_excel} - Col F: {valor}")

        resultados.append({
            "Archivo": nombre_archivo,
            "Hoja": primera_hoja,
            "Fila": fila_excel,
            "Columna": "F",
            "Valor": valor
        })

    # ----------------------------------------
    #   CONDICIÓN: PROCESAR ARCHIVO
    # ----------------------------------------

    F92 = valores_F.get(92)

    # Si F92 comienza con "Identificación de aperturas"
    if isinstance(F92, str) and F92.startswith("Identificación de aperturas"):
        print("  Condición detectada → Se generará copia procesada.")

        # Cargar hoja completa
        try:
            df_full = pd.read_excel(
                ruta_archivo,
                sheet_name=primera_hoja,
                header=None
            )
        except Exception as e:
            print(f"  Error leyendo hoja completa: {e}")
            continue

        # Eliminar filas 91–93 (índices 90–92)
        df_full_procesado = df_full.drop(index=[90, 91, 92], errors='ignore')

        # Crear nombre nuevo SIEMPRE como .xlsx (para evitar error con .xlsm)
        base, ext = os.path.splitext(nombre_archivo)
        nombre_nuevo = base + "_procesado.xlsx"  # fuerza extensión xlsx
        ruta_nueva = os.path.join(carpeta_procesadas, nombre_nuevo)

        # Guardar archivo nuevo PROCESADO
        try:
            with pd.ExcelWriter(ruta_nueva, engine="openpyxl") as writer:
                df_full_procesado.to_excel(
                    writer,
                    sheet_name=primera_hoja,
                    index=False,
                    header=False
                )
            print(f"  ✔ Archivo procesado guardado como: {ruta_nueva}")
        except Exception as e:
            print(f"  Error guardando archivo procesado: {e}")

    else:
        print("  No cumple condición en F92, no se genera copia procesada.")

# Guardar el resumen de F91-F93
if resultados:
    df_res = pd.DataFrame(resultados)
    salida = os.path.join(carpeta, "resultado_columnaF_filas_91_93.xlsx")
    df_res.to_excel(salida, index=False)
    print("\nResumen guardado en:")
    print(salida)
else:
    print("\nNo se pudo obtener información de la columna F en ningún archivo.")
