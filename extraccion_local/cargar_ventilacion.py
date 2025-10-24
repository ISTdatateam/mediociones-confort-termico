#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cargar resultados de ventilación (Excel) a la base MySQL (esquema higiene).

Uso:
  python cargar_ventilacion.py --excel ./resultados_unificados_pareo.xlsx \
    --create-missing-usuarios \
    --create-missing-centros

Requisitos:
  - Variables de entorno: DB_HOST, DB_USERNAME, DB_PASS, DB_NAME, DB_PORT
  - Archivo mysql_utils.py disponible en PYTHONPATH (proporcionado por el usuario)
  - Paquetes: pandas, python-dotenv, mysql-connector-python
"""
import argparse
import sys
import uuid
import math
from dataclasses import dataclass
from typing import Dict, Optional, Tuple, List
import os

import pandas as pd
import numpy as np
from mysql.connector import Error

# Forzar credenciales ANTES de importar mysql_utils (se leen en import-time)
os.environ.update({
    "MYSQL_USER": "higiene_user",
    "MYSQL_PASSWORD": "higiene123",
    "MYSQL_HOST": "localhost",
    "MYSQL_PORT": "3335",
    # Mapeos alternativos por si mysql_utils usa otros nombres
    "DB_USERNAME": "higiene_user",
    "DB_PASSWORD": "higiene123",
    "DB_PASS":     "higiene123",
    "DB_HOST":     "localhost",
    "DB_PORT":     "3335",
    "DB_NAME":     "higiene",
})




# Usa tu helper de conexión
from db.mysql_utils import MySQLDatabaseManager


# --------------------- Utilidades ---------------------
def norm_str(x) -> Optional[str]:
    if pd.isna(x):
        return None
    s = str(x).strip()
    return s if s else None


def try_float(x) -> Optional[float]:
    try:
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return None
        return float(str(x).replace(',', '.'))
    except Exception:
        return None


def calc_seccion_cm2(largo_cm: Optional[float], ancho_cm: Optional[float]) -> Optional[float]:
    if largo_cm is None or ancho_cm is None:
        return None
    return float(largo_cm) * float(ancho_cm)


def calc_caudal_m3h(seccion_cm2: Optional[float], vels_m_s: List[float]) -> Optional[float]:
    if not vels_m_s:
        return None
    if seccion_cm2 is None:
        return None
    # cm2 -> m2
    seccion_m2 = seccion_cm2 / 10000.0
    vel_prom = sum(vels_m_s) / len(vels_m_s)
    return vel_prom * seccion_m2 * 3600.0  # m3/s -> m3/h


def area_codigo(visita_id: int, idx_area: int) -> str:
    return f"V{visita_id}-A{idx_area}"


# --------------------- Acceso a BD ---------------------
@dataclass
class DBContext:
    mgr: MySQLDatabaseManager

    @property
    def conn(self):
        return self.mgr.connection

    def cursor(self, dictcur: bool = True):
        return self.conn.cursor(dictionary=dictcur)


def get_cuv(ctx: DBContext, rut: Optional[str], nombre_ct: Optional[str], direccion_ct: Optional[str]) -> Optional[int]:
    """Obtiene CUV por (rut, nombre_ct, direccion_ct). Prueba variaciones si faltan campos."""
    with ctx.cursor() as cur:
        if rut and nombre_ct and direccion_ct:
            cur.execute("""
                SELECT cuv FROM centros_trabajo
                 WHERE rut = %s AND nombre_ct = %s AND direccion_ct = %s
                 LIMIT 1
            """, (rut, nombre_ct, direccion_ct))
            row = cur.fetchone()
            if row: return int(row["cuv"])
        if rut and nombre_ct:
            cur.execute("""
                SELECT cuv FROM centros_trabajo
                 WHERE rut = %s AND nombre_ct = %s
                 ORDER BY cuv ASC
                 LIMIT 1
            """, (rut, nombre_ct))
            row = cur.fetchone()
            if row: return int(row["cuv"])
        if rut:
            cur.execute("""
                SELECT cuv FROM centros_trabajo
                 WHERE rut = %s
                 ORDER BY cuv ASC
                 LIMIT 1
            """, (rut,))
            row = cur.fetchone()
            if row: return int(row["cuv"])
    return None





def ensure_centro(ctx: DBContext, create_missing: bool, rut: Optional[str], razon: Optional[str],
                  nombre_ct: Optional[str], direccion_ct: Optional[str]) -> Optional[int]:
    cuv = get_cuv(ctx, rut, nombre_ct, direccion_ct)
    if cuv is not None:
        return cuv
    if not create_missing:
        return None
    # Crear centro mínimo con PK generada: next_cuv = MAX(cuv)+1
    with ctx.cursor() as cur:
        cur.execute("SELECT COALESCE(MAX(cuv), 0) + 1 AS next_cuv FROM centros_trabajo")
        next_cuv = int(cur.fetchone()["next_cuv"])
        cur.execute("""
            INSERT INTO centros_trabajo
                (cuv, rut, razon_social, rut2, nombre_ct, direccion_ct, comuna_ct, region_ct, region_num_ct)
            VALUES
                (%s,  %s,  %s,          '',   %s,        %s,           '',        '',         0)
        """, (next_cuv, rut or '', razon or '', nombre_ct or '', direccion_ct or ''))
    return next_cuv

def email_exists(ctx: DBContext, email: str) -> bool:
    with ctx.cursor() as cur:
        cur.execute("SELECT 1 FROM usuarios WHERE email = %s LIMIT 1", (email,))
        return cur.fetchone() is not None

def get_usuario_email(ctx: DBContext, nombre: Optional[str]) -> Optional[str]:
    if not nombre:
        return None
    with ctx.cursor() as cur:
        cur.execute("""
            SELECT email FROM usuarios WHERE name = %s LIMIT 1
        """, (nombre,))
        row = cur.fetchone()
        if row:
            return row["email"]
    # Si el valor ya viniera como email en la celda:
    if nombre and '@' in nombre:
        return nombre
    return None


def ensure_usuario(ctx: DBContext, create_missing: bool, nombre: Optional[str]) -> Optional[str]:
    # ¿Ya existe por nombre o viene un email en la celda?
    email = get_usuario_email(ctx, nombre)
    if email or not create_missing:
        return email

    # Normaliza un "base" para el email sintético
    base = (nombre or "consultor").strip().lower()
    base = base.replace(' ', '.')
    base = ''.join(ch for ch in base if ch.isalnum() or ch in '._-') or "consultor"

    # Construye un email único: base@pendiente.local, base1@..., base2@..., etc.
    domain = "pendiente.local"
    candidate = f"{base}@{domain}"
    if email_exists(ctx, candidate):
        i = 1
        while True:
            candidate = f"{base}{i}@{domain}"
            if not email_exists(ctx, candidate):
                break
            i += 1

    # Inserta el usuario mínimo
    with ctx.cursor() as cur:
        cur.execute("""
            INSERT INTO usuarios (email, pass, name, type, zonal, cargo)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (candidate, 'changeme', nombre or 'Consultor', 0, None, None))

    return candidate




def insert_visita(ctx: DBContext, data: dict) -> int:
    with ctx.cursor() as cur:
        cur.execute("""
            INSERT INTO visitas
            (cuv_visita, fecha_visita, hora_visita, motivo_evaluacion, nombre_personal_visita,
             cargo_personal_visita, consultor_ist, note_visita, consultor_cargo, consultor_zonal, tipo_evaluacion)
            VALUES
            (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            data.get("cuv_visita"),
            data.get("fecha_visita"),
            data.get("hora_visita"),
            data.get("motivo_evaluacion"),
            data.get("nombre_personal_visita"),
            data.get("cargo_personal_visita"),
            data.get("consultor_ist"),
            data.get("note_visita"),
            data.get("consultor_cargo"),
            data.get("consultor_zonal"),
            data.get("tipo_evaluacion", "ventilacion"),
        ))
        visita_id = cur.lastrowid
    return int(visita_id)


def insert_v_area(ctx: DBContext, area: dict) -> str:
    with ctx.cursor() as cur:
        cur.execute("""
            INSERT INTO v_areas
            (area_id, visita_id, centro_id, codigo_area, nombre_area,
             uso, piso_nivel, largo_m, ancho_m, alto_m, volumen_m3, m3_porpersona,
             m3_porpersona_594, m3_porpersona_cumple, recambio_hora, recambio_hora_cumple,
             ventilacion_tipo, ventilacion_sistema, ventilacion_estado, aberturas, croquis_url, observaciones)
            VALUES
            (%s, %s, %s, %s, %s,
             %s, %s, %s, %s, %s, %s, %s,
             %s, %s, %s, %s,
             %s, %s, %s, %s, %s, %s)
        """, (
            area["area_id"], area["visita_id"], area["centro_id"], area["codigo_area"], area["nombre_area"],
            area.get("uso"), area.get("piso_nivel"), area.get("largo_m"), area.get("ancho_m"), area.get("alto_m"),
            area.get("volumen_m3"), area.get("m3_porpersona"),
            area.get("m3_porpersona_594", 10.0), area.get("m3_porpersona_cumple", 0),
            area.get("recambio_hora"), area.get("recambio_hora_cumple", 0),
            area.get("ventilacion_tipo"), area.get("ventilacion_sistema"), area.get("ventilacion_estado"),
            area.get("aberturas"), area.get("croquis_url"), area.get("observaciones")
        ))
    return area["area_id"]


def insert_v_punto(ctx: DBContext, p: dict) -> str:
    with ctx.cursor() as cur:
        cur.execute("""
            INSERT INTO v_puntos_medicion
            (punto_id, evaluacion_id, area_id, codigo_punto, tipo_punto, ubicacion_detalle,
             altura_m, distancia_fuente_m, conducto_largo_cm, conducto_ancho_cm, conducto_diametro,
             seccion_conducto_cm2, medicion_caudal_1, medicion_caudal_2, medicion_caudal_3, medicion_caudal_4,
             medicion_caudal_5, caudal, fecha_hora, condiciones_ocupacion, puertas_ventanas_abiertas,
             temperatura_c, humedad_relativa_pct, croquis_url, observaciones)
            VALUES
            (%s, %s, %s, %s, %s, %s,
             %s, %s, %s, %s, %s,
             %s, %s, %s, %s, %s,
             %s, %s, %s, %s, %s,
             %s, %s, %s, %s)
        """, (
            p["punto_id"], p["evaluacion_id"], p["area_id"], p["codigo_punto"], p["tipo_punto"], p.get("ubicacion_detalle"),
            p.get("altura_m"), p.get("distancia_fuente_m"), p.get("conducto_largo_cm"), p.get("conducto_ancho_cm"), p.get("conducto_diametro"),
            p.get("seccion_conducto_cm2"), p.get("medicion_caudal_1"), p.get("medicion_caudal_2"), p.get("medicion_caudal_3"), p.get("medicion_caudal_4"),
            p.get("medicion_caudal_5"), p.get("caudal"), p.get("fecha_hora"), p.get("condiciones_ocupacion"), p.get("puertas_ventanas_abiertas"),
            p.get("temperatura_c"), p.get("humedad_relativa_pct"), p.get("croquis_url"), p.get("observaciones")
        ))
    return p["punto_id"]


# --------------------- Proceso principal ---------------------
def process_excel_row(row, ctx: DBContext, args) -> None:
    # Campos cabecera
    razon = norm_str(row.get("RAZON SOCIAL EMPRESA"))
    rut   = norm_str(row.get("RUT EMPRESA"))
    nombre_ct = norm_str(row.get("NOMBRE DEL CENTRO DE TRABAJO (CT)"))
    direccion_ct = norm_str(row.get("DIRECCIÓN (CT)"))
    fecha_visita = row.get("FECHA DE VISITA")
    hora_visita  = row.get("HORARIO DE MEDICIÓN")
    nombre_persona = norm_str(row.get("NOMBRE PERSONA EMPRESA"))
    cargo_persona  = norm_str(row.get("CARGO PERSONA EMPRESA"))
    nombre_consultor = norm_str(row.get("NOMBRE PROFESIONAL IST"))
    cargo_consultor  = norm_str(row.get("CARGO PROFESIONAL IST"))

    cuv = get_cuv(ctx, rut, nombre_ct, direccion_ct)
    if cuv is None:
        cuv = ensure_centro(ctx, args.create_missing_centros, rut, razon, nombre_ct, direccion_ct)
    if cuv is None:
        raise RuntimeError(f"No se encontró (ni creó) CUV para RUT={rut}, CT={nombre_ct}, DIR={direccion_ct}")

    email_consultor = get_usuario_email(ctx, nombre_consultor)
    if email_consultor is None:
        email_consultor = ensure_usuario(ctx, args.create_missing_usuarios, nombre_consultor)
    if email_consultor is None:
        raise RuntimeError(f"No se encontró (ni creó) usuario para consultor='{nombre_consultor}'")

    # Inserta visita
    visita_id = insert_visita(ctx, {
        "cuv_visita": cuv,
        "fecha_visita": fecha_visita,
        "hora_visita": hora_visita,
        "motivo_evaluacion": None,
        "nombre_personal_visita": nombre_persona,
        "cargo_personal_visita": cargo_persona,
        "consultor_ist": email_consultor,
        "note_visita": None,
        "consultor_cargo": cargo_consultor,
        "consultor_zonal": None,
        "tipo_evaluacion": "ventilacion",
    })

    # Áreas (1..4 si existen)
    area_ids = {}
    for i in range(1, 5):
        area_nombre = norm_str(row.get(f"Area_{i}"))
        if not area_nombre:
            continue

        largo = try_float(row.get(f"Largo_{i}"))
        ancho = try_float(row.get(f"Ancho_{i}"))
        alto  = try_float(row.get(f"Alto_{i}"))
        volumen = try_float(row.get(f"Volumen_{i}"))
        m3pp = try_float(row.get(f"m3xpersona_{i}"))
        resultado = norm_str(row.get(f"Resultado_{i}"))

        cumple = 0
        if resultado:
            cumple = 1 if "cumple" in resultado.lower() and "no" not in resultado.lower() else 0
        elif m3pp is not None:
            cumple = 1 if m3pp >= 10 else 0  # DS 594

        area_id = str(uuid.uuid4())
        area_row = {
            "area_id": area_id,
            "visita_id": visita_id,
            "centro_id": cuv,
            "codigo_area": area_codigo(visita_id, i),
            "nombre_area": area_nombre,
            "largo_m": largo,
            "ancho_m": ancho,
            "alto_m": alto,
            "volumen_m3": volumen,
            "m3_porpersona": m3pp,
            "m3_porpersona_594": 10.0,
            "m3_porpersona_cumple": cumple,
            # recambio_hora (se calculará luego al sumar caudales de puntos)
            "recambio_hora": None,
            "recambio_hora_cumple": 0,
        }
        insert_v_area(ctx, area_row)
        area_ids[i] = area_id

        # Puntos INY/EXT j=1..5, k=1..5 repeticiones de velocidad
        for tipo in ["iny", "ext"]:
            for j in range(1, 6):
                largo_cm = try_float(row.get(f"Area_{i}_{tipo}_largo_{j}"))
                ancho_cm = try_float(row.get(f"Area_{i}_{tipo}_ancho_{j}"))
                # Salta si no hay ducto definido
                if largo_cm is None and ancho_cm is None:
                    continue

                vels = []
                for k in range(1, 6):
                    v = try_float(row.get(f"Area_{i}_{tipo}_vel_{j}_{k}"))
                    if v is not None:
                        vels.append(v)

                seccion_cm2 = calc_seccion_cm2(largo_cm, ancho_cm)
                caudal = calc_caudal_m3h(seccion_cm2, vels)

                punto = {
                    "punto_id": str(uuid.uuid4()),
                    "evaluacion_id": visita_id,
                    "area_id": area_id,
                    "codigo_punto": f"{tipo.upper()}-{j}",
                    "tipo_punto": "inyeccion" if tipo == "iny" else "extraccion",
                    "ubicacion_detalle": norm_str(row.get(f"Area_{i}_{tipo}_{j}")),
                    "conducto_largo_cm": largo_cm,
                    "conducto_ancho_cm": ancho_cm,
                    "seccion_conducto_cm2": seccion_cm2,
                    "medicion_caudal_1": vels[0] if len(vels) > 0 else None,
                    "medicion_caudal_2": vels[1] if len(vels) > 1 else None,
                    "medicion_caudal_3": vels[2] if len(vels) > 2 else None,
                    "medicion_caudal_4": vels[3] if len(vels) > 3 else None,
                    "medicion_caudal_5": vels[4] if len(vels) > 4 else None,
                    "caudal": caudal,
                }
                insert_v_punto(ctx, punto)

        # Recalcula recambios/hora por área (suma de caudales de puntos / volumen)
        with ctx.cursor() as cur:
            cur.execute("""
                SELECT COALESCE(SUM(caudal),0) AS inj
                  FROM v_puntos_medicion
                 WHERE evaluacion_id = %s AND area_id = %s AND tipo_punto='inyeccion'
            """, (visita_id, area_id))
            inj = float(cur.fetchone()["inj"] or 0.0)
            cur.execute("""
                SELECT COALESCE(SUM(caudal),0) AS ext
                  FROM v_puntos_medicion
                 WHERE evaluacion_id = %s AND area_id = %s AND tipo_punto='extraccion'
            """, (visita_id, area_id))
            ext = float(cur.fetchone()["ext"] or 0.0)

        recambio = None
        if volumen and max(inj, ext) > 0:
            recambio = max(inj, ext) / volumen  # ACH ≈ m3/h / m3
        cumple_recambio = 1 if (recambio is not None and 6.0 <= recambio <= 60.0) else 0

        with ctx.cursor() as cur:
            cur.execute("""
                UPDATE v_areas
                   SET recambio_hora = %s, recambio_hora_cumple = %s
                 WHERE area_id = %s
            """, (recambio, cumple_recambio, area_id))


def main():
    """
    Versión sin CLI: usa ruta fija y flags editables en el código.
    - Excel fijo: extraccion_local/resultados_unificados.xlsx
    - Flags por defecto: crear usuarios/centros faltantes = True
    """

    EXCEL_PATH = r"C:\Users\Quantum-Malloco\mediociones-confort-termico\extraccion_local\resultados_unificados.xlsx"
    CREATE_MISSING_USUARIOS = True   # cambia a False si no quieres crear usuarios automáticamente
    CREATE_MISSING_CENTROS = True    # cambia a False si no quieres crear centros automáticamente
    # Conexión a BD usando mysql_utils.py (variables de entorno)
    mgr = MySQLDatabaseManager()
    if mgr.connection is None:
        print("No se pudo conectar a la BD. Revisa tus variables de entorno.")
        sys.exit(1)
    ctx = DBContext(mgr=mgr)
    # Carga del Excel
    df = pd.read_excel(EXCEL_PATH, sheet_name=0)
    # Simula args para reusar process_excel_row(row, ctx, args)
    class Args: pass
    args = Args()
    args.create_missing_usuarios = CREATE_MISSING_USUARIOS
    args.create_missing_centros = CREATE_MISSING_CENTROS
    # Iniciar transacción
    try:
        ctx.conn.start_transaction()
        for idx, row in df.iterrows():
            process_excel_row(row, ctx, args)
        ctx.conn.commit()
        print(f"✔ Carga completada. Filas procesadas: {len(df)}")
    except Error as e:
        ctx.conn.rollback()
        print(f"✖ Error MySQL: {e}")
        sys.exit(2)
    except Exception as ex:
        ctx.conn.rollback()
        print(f"✖ Error general: {ex}")
        sys.exit(3)
    finally:
        mgr.close()

if __name__ == "__main__":
    main()
