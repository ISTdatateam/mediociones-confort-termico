from __future__ import annotations

import os
import unicodedata
import pandas as pd
import mysql.connector
from mysql.connector import Error
from dotenv import load_dotenv

load_dotenv()


def get_connection() -> mysql.connector.MySQLConnection:
    """Usa la misma estructura del script de ventilación."""
    return mysql.connector.connect(
        host=os.getenv("DB_HOST"),
        user=os.getenv("DB_USERNAME"),
        password=os.getenv("DB_PASS"),
        database=os.getenv("DB_NAME"),
        port=int(os.getenv("DB_PORT", 3306)),
        autocommit=False,
    )


def _normalize(name: str) -> str:
    """Normaliza columnas: sin tildes, minúsculas, guion bajo."""
    name = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode("ascii")
    return name.strip().lower().replace(" ", "_")


def _clean(value):
    """Limpia texto, NaN, cadenas vacías."""
    if pd.isna(value):
        return None
    if isinstance(value, str) and value.strip() == "":
        return None
    return value


def cargar_centros_trabajo() -> int:
    # Ruta fija: Locales.xlsx en el mismo directorio del script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    ruta_excel = os.path.join(script_dir, "Locales.xlsx")

    print(f"Leyendo archivo: {ruta_excel}")

    if not os.path.exists(ruta_excel):
        raise FileNotFoundError(f"No se encontró el archivo {ruta_excel}")

    df = pd.read_excel(ruta_excel)

    # Normalizar nombres de columnas
    df.columns = [_normalize(c) for c in df.columns]

    # Mapeo esperado según la tabla
    aliases = {
        "cuv": "cuv",
        "rut": "rut",
        "razon_social": "razon_social",
        "rut2": "rut2",
        "nombre_ct": "nombre_ct",
        "direccion_ct": "direccion_ct",
        "comuna_ct": "comuna_ct",
        "region_ct": "region_ct",
        "region_num_ct": "region_num_ct",
    }

    # Renombrar si vienen columnas con variaciones
    rename_map = {}
    for col in df.columns:
        if col in aliases:
            rename_map[col] = aliases[col]

    df = df.rename(columns=rename_map)

    # Validar columnas requeridas
    columnas_requeridas = list(aliases.values())
    faltantes = [c for c in columnas_requeridas if c not in df.columns]
    if faltantes:
        raise ValueError(f"❌ Faltan columnas requeridas en el Excel: {faltantes}")

    # Limpiar valores
    for c in df.columns:
        df[c] = df[c].apply(_clean)

    # Convertir tipos numéricos
    df["cuv"] = pd.to_numeric(df["cuv"], errors="raise")
    df["region_num_ct"] = pd.to_numeric(df["region_num_ct"], errors="raise")

    print("\nVista previa de datos:")
    print(df.head())

    insert_sql = """
        INSERT INTO centros_trabajo (
            cuv, rut, razon_social, rut2, nombre_ct,
            direccion_ct, comuna_ct, region_ct, region_num_ct
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE
            rut = VALUES(rut),
            razon_social = VALUES(razon_social),
            rut2 = VALUES(rut2),
            nombre_ct = VALUES(nombre_ct),
            direccion_ct = VALUES(direccion_ct),
            comuna_ct = VALUES(comuna_ct),
            region_ct = VALUES(region_ct),
            region_num_ct = VALUES(region_num_ct)
    """

    conn = get_connection()
    cursor = conn.cursor()

    try:
        rows = list(df.itertuples(index=False, name=None))
        print(f"\nRegistros a insertar/actualizar: {len(rows)}")

        cursor.executemany(insert_sql, rows)
        conn.commit()

        print(f"✅ Carga completada. Filas afectadas: {cursor.rowcount}")
        return cursor.rowcount

    except Exception as e:
        conn.rollback()
        print("❌ Error durante la carga:", str(e))
        raise

    finally:
        cursor.close()
        conn.close()


if __name__ == "__main__":
    cargar_centros_trabajo()
