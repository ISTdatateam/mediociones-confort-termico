import os
import pandas as pd
from openpyxl import load_workbook

# Carpeta donde están los archivos
carpeta = r"C:\Users\Quantum-Malloco\Desktop\Ventilacion\Casos de prueba 2"

# Texto a buscar (búsqueda parcial)
texto_buscar = "Identificación de aperturas por donde ingresa"
texto_normalizado = texto_buscar.strip().casefold()

# Mapeo de "Posición" (fila donde se encuentra el texto) -> (fila_en_la_que_inserto, cantidad_de_filas)
acciones = {
    50: (47, 28),
    57: (51, 21),
    55: (51, 23),
    54: (51, 24),
    53: (51, 25),
    58: (55, 20),
    62: (59, 16),
    60: (59, 18),
    66: (63, 12),
    71: (71, 7),
    42: (39, 36),
    45: (43, 33),
    49: (47, 29),
    51: (51, 31),
    52: (51, 30),
    59: (54, 19),
    61: (59, 18),
    65: (63, 13),
    67: (64, 11),
    70: (67, 8)
}

# Recorremos los archivos de la carpeta
lista_archivos = [
    f for f in os.listdir(carpeta)
    if os.path.isfile(os.path.join(carpeta, f))
]

for nombre_archivo in lista_archivos:
    # Solo procesar xlsx / xlsm (openpyxl no maneja bien .xls antiguos)
    if not nombre_archivo.lower().endswith(('.xlsx', '.xlsm')):
        continue

    ruta_archivo = os.path.join(carpeta, nombre_archivo)
    print(f"\nRevisando: {ruta_archivo}")

    # ---------------- BUSCAR POSICIÓN DEL TEXTO (con pandas) ----------------
    try:
        # Abrimos el archivo con un contexto para cerrar bien el handle
        with pd.ExcelFile(ruta_archivo) as xls:
            # Primera hoja
            primera_hoja = xls.sheet_names[0]

            # Leer columnas F:L
            df_cols = pd.read_excel(
                xls,
                sheet_name=primera_hoja,
                usecols="F:L",
                header=None
            )

    except ValueError:
        print("  La primera hoja no tiene columnas F:L. Se omite.")
        continue
    except Exception as e:
        print(f"  Error leyendo {nombre_archivo}: {e}")
        continue

    # Normalizar a texto en minúsculas (casefold) y sin espacios extremos
    df_norm = (
        df_cols
        .fillna("")
        .astype(str)
        .apply(lambda col: col.str.strip().str.casefold())
    )

    # Máscara de coincidencia parcial
    df_mask = df_norm.apply(
        lambda col: col.str.contains(texto_normalizado, na=False)
    )

    # Filas en las que aparece el texto al menos en una columna
    mask_filas = df_mask.any(axis=1)
    indices_coinciden = mask_filas[mask_filas].index

    if len(indices_coinciden) == 0:
        print("  No se encontró el texto en la primera hoja (columnas F:L).")
        continue

    # Tomamos la primera fila donde se encontró el texto
    idx_primero = indices_coinciden[0]
    fila_excel = idx_primero + 1  # 1-based como en Excel

    print(f"  >>> Texto encontrado en la fila {fila_excel}")

    # ---------------- DETERMINAR ACCIÓN SEGÚN TABLA ----------------
    if fila_excel not in acciones:
        print("  No hay acción configurada para esta posición. No se inserta nada.")
        continue

    fila_insertar, cantidad_filas = acciones[fila_excel]
    print(f"  Acción: insertar {cantidad_filas} filas en la fila {fila_insertar}")

    # ---------------- INSERTAR FILAS EN EL ARCHIVO (openpyxl) ----------------
    try:
        # keep_vba=True para no romper macros en xlsm
        if nombre_archivo.lower().endswith('.xlsm'):
            wb = load_workbook(ruta_archivo, keep_vba=True)
        else:
            wb = load_workbook(ruta_archivo)

        ws = wb[primera_hoja]

        # insert_rows(row_idx, amount=N) inserta N filas ANTES de row_idx
        ws.insert_rows(fila_insertar, amount=cantidad_filas)

        wb.save(ruta_archivo)
        print(f"  Filas insertadas correctamente en {nombre_archivo}.")

    except Exception as e:
        print(f"  Error insertando filas en {nombre_archivo}: {e}")
        continue

print("\nProceso finalizado.")
