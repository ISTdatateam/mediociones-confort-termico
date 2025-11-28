import os
import pandas as pd
from openpyxl import load_workbook, Workbook

# Carpeta donde están los archivos
carpeta = r"C:\Users\Quantum-Malloco\Desktop\Ventilacion\Casos de prueba 3"

# Subcarpeta de salida
carpeta_salida = os.path.join(carpeta, "Normalizados")
os.makedirs(carpeta_salida, exist_ok=True)

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
    71: (70, 7),
    42: (39, 36),
    45: (43, 33),
    49: (47, 29),
    51: (50, 31),
    52: (51, 26),
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
    # Solo procesar xlsx / xlsm
    if not nombre_archivo.lower().endswith(('.xlsx', '.xlsm')):
        continue

    ruta_archivo = os.path.join(carpeta, nombre_archivo)
    print(f"\nRevisando: {ruta_archivo}")

    # ---------------- BUSCAR POSICIÓN DEL TEXTO (con pandas) ----------------

    nombre_hoja = "Hoja_terreno"
    posicion_hoja = 0  # Si no existe la hoja por nombre, usar esta posición

    try:
        with pd.ExcelFile(ruta_archivo) as xls:

            # Revisar si la hoja existe por nombre
            if nombre_hoja in xls.sheet_names:
                hoja = nombre_hoja
            else:
                # Usar la hoja por posición
                hoja = xls.sheet_names[posicion_hoja]

            df_cols = pd.read_excel(
                xls,
                sheet_name=hoja,
                usecols="F:L",
                header=None
            )

    except ValueError:
        print("  La primera hoja no tiene columnas F:L. Se omite.")
        continue
    except Exception as e:
        print(f"  Error leyendo {nombre_archivo}: {e}")
        continue

    # Normalizar
    df_norm = (
        df_cols
        .fillna("")
        .astype(str)
        .apply(lambda col: col.str.strip().str.casefold())
    )

    df_mask = df_norm.apply(
        lambda col: col.str.contains(texto_normalizado, na=False)
    )

    mask_filas = df_mask.any(axis=1)
    indices_coinciden = mask_filas[mask_filas].index

    if len(indices_coinciden) == 0:
        print("  No se encontró el texto en la primera hoja (columnas F:L).")
        continue

    idx_primero = indices_coinciden[0]
    fila_excel = idx_primero + 1  # 1-based como Excel

    print(f"  >>> Texto encontrado en la fila {fila_excel}")

    # ---------------- DETERMINAR ACCIÓN SEGÚN TABLA ----------------
    if fila_excel not in acciones:
        print("  No hay acción configurada para esta posición. No se genera versión normalizada.")
        continue

    fila_insertar, cantidad_filas = acciones[fila_excel]
    print(f"  Acción: crear versión normalizada con hueco a partir de la fila {fila_insertar}, tamaño {cantidad_filas}.")

    # ---------------- ABRIR WORKBOOK ORIGINAL ----------------
    try:
        if nombre_archivo.lower().endswith('.xlsm'):
            wb = load_workbook(ruta_archivo, keep_vba=True)
        else:
            wb = load_workbook(ruta_archivo)

        ws_orig = wb[hoja]

        max_row = ws_orig.max_row
        max_col = ws_orig.max_column

        # Crear nueva hoja donde se va a armar la versión normalizada
        nombre_hoja_original = ws_orig.title
        ws_new = wb.create_sheet(title=f"{nombre_hoja_original}_normalizado_tmp")

        # ---------------- COPIAR FILAS 1 .. fila_insertar ----------------
        # (Ejemplo: 1..47)
        for row in range(1, fila_insertar + 1):
            for col in range(1, max_col + 1):
                src = ws_orig.cell(row=row, column=col)
                dst = ws_new.cell(row=row, column=col)
                dst.value = src.value
                dst._style = src._style
                dst.number_format = src.number_format
                dst.data_type = src.data_type
                if src.hyperlink:
                    dst.hyperlink = src.hyperlink
                if src.comment:
                    dst.comment = src.comment

        # ---------------- COPIAR FILAS fila_insertar+1 .. max_row ----------------
        # Pegando desde fila_insertar + cantidad_filas
        # Ejemplo: copiar desde 48, pegando desde 75 (47 + 28)
        offset = cantidad_filas - 1  # tal como se deduce de tu ejemplo 47 -> 75 para la fila 48
        for row in range(fila_insertar + 1, max_row + 1):
            new_row = row + offset
            for col in range(1, max_col + 1):
                src = ws_orig.cell(row=row, column=col)
                dst = ws_new.cell(row=new_row, column=col)
                dst.value = src.value
                dst._style = src._style
                dst.number_format = src.number_format
                dst.data_type = src.data_type
                if src.hyperlink:
                    dst.hyperlink = src.hyperlink
                if src.comment:
                    dst.comment = src.comment

        # Opcional: aquí podríamos limpiar explícitamente el bloque
        # de filas "hueco" si quieres asegurarte que estén realmente vacías:
        # for row in range(fila_insertar + 1, fila_insertar + offset + 1):
        #     for col in range(1, max_col + 1):
        #         cell = ws_new.cell(row=row, column=col)
        #         cell.value = None
        #         cell._style = cell._style.__class__()

        # Eliminar la hoja original y renombrar la nueva como la original
        wb.remove(ws_orig)
        ws_new.title = nombre_hoja_original

        # ---------------- GUARDAR EN CARPETA 'Normalizados' ----------------
        nombre_base, ext = os.path.splitext(nombre_archivo)
        nuevo_nombre = f"{nombre_base}_normalizado{ext}"
        ruta_salida = os.path.join(carpeta_salida, nuevo_nombre)

        wb.save(ruta_salida)
        print(f"  Archivo normalizado guardado en: {ruta_salida}")

    except Exception as e:
        print(f"  Error generando versión normalizada para {nombre_archivo}: {e}")
        continue

print("\nProceso finalizado.")
