import os
import shutil
import pandas as pd

# Carpeta donde están los archivos
carpeta = r"C:\Users\Quantum-Malloco\Desktop\Ventilacion\Casos de prueba 3\Normalizados"

# Texto a buscar (búsqueda parcial)
texto_buscar = "Identificación de aperturas por donde ingresa"
texto_normalizado = texto_buscar.strip().casefold()

resultados = []

# Tomamos un "snapshot" de los archivos para que moverlos no afecte el listado
lista_archivos = [
    f for f in os.listdir(carpeta)
    if os.path.isfile(os.path.join(carpeta, f))
]

for nombre_archivo in lista_archivos:
    # Solo procesar archivos Excel
    if not nombre_archivo.lower().endswith(('.xlsx', '.xlsm', '.xls')):
        continue

    ruta_archivo = os.path.join(carpeta, nombre_archivo)
    print(f"Revisando: {ruta_archivo}")

    # ---------- ABRIR Y LEER EXCEL CERRANDO CORRECTAMENTE ----------
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

            df_cols = pd.read_excel (
                xls,
                sheet_name=hoja,
                usecols="F",
                header=None
            )

    except ValueError:
        print("  La primera hoja no tiene columnas F:L.")
        continue
    except Exception as e:
        print(f"  Error leyendo la primera hoja en {nombre_archivo}: {e}")
        continue

    # ---------- NORMALIZAR Y BUSCAR TEXTO ----------
    # Convertir a texto, quitar espacios y pasar a lower "robusto"
    df_norm = (
        df_cols
        .fillna("")
        .astype(str)
        .apply(lambda col: col.str.strip().str.casefold())
    )

    # Máscara: TRUE si la celda contiene el texto buscado (búsqueda parcial)
    df_mask = df_norm.apply(
        lambda col: col.str.contains(texto_normalizado, na=False)
    )

    # Filas donde al menos una columna tiene el texto
    mask_filas = df_mask.any(axis=1)
    indices_coinciden = mask_filas[mask_filas].index

    if len(indices_coinciden) > 0:
        # Tomamos la primera coincidencia
        idx_primero = indices_coinciden[0]
        fila_excel = idx_primero + 1  # ajustar a numeración de Excel

        # Mapeo de columnas reales a letras F:G:H:I:J:K:L
        letras_columnas = ["F", "G", "H", "I", "J", "K", "L"]
        mapeo_col = {
            df_cols.columns[i]: letras_columnas[i]
            for i in range(len(df_cols.columns))
        }

        columnas_en_fila = []
        for col in df_cols.columns:
            if df_mask.loc[idx_primero, col]:
                columnas_en_fila.append(mapeo_col[col])

        columnas_str = ", ".join(columnas_en_fila) if columnas_en_fila else ""

        resultados.append({
            "Archivo": nombre_archivo,
            "Hoja": hoja,
            "Fila": fila_excel,
            "Columnas_donde_aparece": columnas_str
        })

        print(f"  >>> Encontrado en fila {fila_excel}, columna(s): {columnas_str}")

        # ---------- CREAR SUBCARPETA Y MOVER ARCHIVO ----------
        # Nombre de la subcarpeta según la fila
        # Si prefieres otro formato, por ejemplo f"Fila_{fila_excel}", cambia aquí
        nombre_subcarpeta = str(fila_excel)

        ruta_subcarpeta = os.path.join(carpeta, nombre_subcarpeta)
        os.makedirs(ruta_subcarpeta, exist_ok=True)

        ruta_destino = os.path.join(ruta_subcarpeta, nombre_archivo)

        # Evitar sobrescribir si ya existe un archivo igual en la carpeta destino
        if os.path.exists(ruta_destino):
            print(f"  Atención: ya existe un archivo con ese nombre en {ruta_subcarpeta}. No se mueve.")
        else:
            try:
                shutil.move(ruta_archivo, ruta_destino)
                print(f"  Archivo movido a: {ruta_destino}")
            except PermissionError as e:
                print(f"  No se pudo mover el archivo por un problema de permisos: {e}")
    else:
        print("  No se encontró el texto en la primera hoja (columnas F:L).")

# ---------- GUARDAR RESUMEN ----------
if resultados:
    df_resultados = pd.DataFrame(resultados)
    salida = os.path.join(carpeta, "resultado_busqueda_IL_parcial.xlsx")
    df_resultados.to_excel(salida, index=False)
    print("\nResumen guardado en:")
    print(salida)
else:
    print("\nNo se encontró el texto en ningún archivo.")
