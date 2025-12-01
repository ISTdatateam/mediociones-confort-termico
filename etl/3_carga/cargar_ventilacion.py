from __future__ import annotations

import os
from datetime import date, time
import unicodedata
from typing import Dict, Iterable, List, Optional

import mysql.connector
import pandas as pd
from dotenv import load_dotenv

load_dotenv()


def get_connection() -> mysql.connector.MySQLConnection:
    return mysql.connector.connect(
        host=os.getenv("DB_HOST"),
        user=os.getenv("DB_USERNAME"),
        password=os.getenv("DB_PASS"),
        database=os.getenv("DB_NAME"),
        port=int(os.getenv("DB_PORT", 3306)),
        autocommit=False,
    )


def _first_available_sheet(xls: pd.ExcelFile, names: Iterable[str]) -> str:
    for name in names:
        if name in xls.sheet_names:
            return name
    raise ValueError(f"Ninguna hoja encontrada; se esperaban: {', '.join(names)}")


def _normalize_column_name(name: str) -> str:
    """Normaliza un nombre de columna quitando acentos y reemplazando espacios."""

    name = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode("ascii")
    return name.strip().lower().replace(" ", "_")


def _apply_aliases(df: pd.DataFrame, aliases: Dict[str, str]) -> pd.DataFrame:
    """Renombra columnas usando un mapa de alias tras normalizar los nombres."""

    rename_map = {}
    for original in df.columns:
        normalized = _normalize_column_name(original)
        if normalized in aliases:
            rename_map[original] = aliases[normalized]
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def _clean_value(value):
    if pd.isna(value) or (isinstance(value, str) and value.strip() == ""):
        return None
    return value


def _to_float(value):
    if value is None or pd.isna(value):
        return None
    try:
        if isinstance(value, str):
            value = value.replace(",", ".")
        return float(value)
    except (TypeError, ValueError):
        return None


def _with_defaults(area: Dict) -> Dict:
    defaults = {
        "m3_porpersona_594": 10.0,
        "m3_porpersona_cumple": 0,
        "caudal_inyeccion_total": 0.0,
        "caudal_extraccion_total": 0.0,
        "m3_porpersona_hora_594": 20.0,
        "m3_porpersona_hora_cumple": 0,
        "recambio_hora_594_min": 6.0,
        "recambio_hora_594_max": 60.0,
        "recambio_hora_cumple": 0,
    }

    for key, value in defaults.items():
        if area.get(key) is None:
            area[key] = value
    return area


def _create_equipo_medicion(cursor, equipo_id: str, ventilacion_data: Dict) -> str:
    nombre_equipo = _clean_value(ventilacion_data.get("instru_nombre_1")) or equipo_id
    marca_equipo = _clean_value(ventilacion_data.get("instru_marca_1"))
    modelo_equipo = _clean_value(ventilacion_data.get("instru_modelo_1")) or equipo_id
    n_serie_equipo = _clean_value(ventilacion_data.get("instru_nserie_1"))
    num_certificado = _clean_value(ventilacion_data.get("instru_ncertificado_1"))

    cursor.execute(
        """
        INSERT INTO equipos_medicion (
            id_equipo,
            nombre_equipo,
            cod_equipo,
            n_serie_equipo,
            marca_equipo,
            modelo_equipo,
            estado_equipo,
            obs_equipo,
            estado_calibracion,
            fecha_calibracion,
            prox_calibracion,
            empresa_certificadora,
            num_certificado,
            fecha_ingreso,
            observaciones,
            simple_cod,
            url_certificado,
            equipo_dicc,
            tipo,
            patron_tbs,
            patron_tbh,
            patron_tg
        )
        VALUES (
            %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
        )
        """,
        (
            equipo_id,
            nombre_equipo,
            None,
            n_serie_equipo,
            marca_equipo,
            modelo_equipo,
            None,
            None,
            None,
            None,
            None,
            None,
            num_certificado,
            None,
            None,
            None,
            None,
            None,
            "instrumento",
            0.0,
            0.0,
            0.0,
        ),
    )

    return equipo_id


def _resolve_equipo_id(
    cursor, equipo_valor: Optional[str], ventilacion_data: Optional[Dict] = None
) -> Optional[str]:
    codigo = _clean_value(equipo_valor)
    if not codigo:
        return None

    cursor.execute(
        "SELECT id_equipo FROM equipos_medicion WHERE id_equipo = %s OR equipo_dicc = %s",
        (codigo, codigo),
    )
    row = cursor.fetchone()
    if row:
        return row[0]

    if ventilacion_data is None:
        return None

    return _create_equipo_medicion(cursor, codigo, ventilacion_data)


def _insert_ev_ventilacion(cursor, visita_id: int, ventilacion_data: Dict):
    equipo_temp = _resolve_equipo_id(cursor, ventilacion_data.get("equipo_temp"))

    equipo_vel_valor = ventilacion_data.get("instru_nserie_1") or ventilacion_data.get("instru_modelo_1")
    equipo_vel = _resolve_equipo_id(cursor, equipo_vel_valor, ventilacion_data)

    cursor.execute(
        """
        INSERT INTO ev_ventilacion (
            visita_id,
            equipo_temp,
            equipo_vel_air,
            instru_nombre_1,
            instru_marca_1,
            instru_modelo_1,
            instru_nserie_1,
            instru_ncertificado_1,
            instru_nombre_2,
            instru_marca_2,
            instru_modelo_2,
            instru_nserie_2,
            instru_ncertificado_2
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE
            equipo_temp = VALUES(equipo_temp),
            equipo_vel_air = VALUES(equipo_vel_air),
            instru_nombre_1 = VALUES(instru_nombre_1),
            instru_marca_1 = VALUES(instru_marca_1),
            instru_modelo_1 = VALUES(instru_modelo_1),
            instru_nserie_1 = VALUES(instru_nserie_1),
            instru_ncertificado_1 = VALUES(instru_ncertificado_1),
            instru_nombre_2 = VALUES(instru_nombre_2),
            instru_marca_2 = VALUES(instru_marca_2),
            instru_modelo_2 = VALUES(instru_modelo_2),
            instru_nserie_2 = VALUES(instru_nserie_2),
            instru_ncertificado_2 = VALUES(instru_ncertificado_2)
        """,
        (
            visita_id,
            equipo_temp,
            equipo_vel,
            ventilacion_data.get("instru_nombre_1"),
            ventilacion_data.get("instru_marca_1"),
            ventilacion_data.get("instru_modelo_1"),
            ventilacion_data.get("instru_nserie_1"),
            ventilacion_data.get("instru_ncertificado_1"),
            ventilacion_data.get("instru_nombre_2"),
            ventilacion_data.get("instru_marca_2"),
            ventilacion_data.get("instru_modelo_2"),
            ventilacion_data.get("instru_nserie_2"),
            ventilacion_data.get("instru_ncertificado_2"),
        ),
    )


def _compute_area_calculations(area: Dict, puntos_df: pd.DataFrame) -> Dict:
    area = _with_defaults(area)

    largo = _to_float(area.get("largo_m"))
    ancho = _to_float(area.get("ancho_m"))
    alto = _to_float(area.get("alto_m"))

    volumen = _to_float(area.get("volumen_m3"))
    if volumen is None and None not in (largo, ancho, alto):
        volumen = largo * ancho * alto
        area["volumen_m3"] = volumen

    aforo = _to_float(area.get("aforo_permitido"))
    area["aforo_permitido"] = aforo

    puntos_area = puntos_df
    if not puntos_df.empty and "area_id" in puntos_df.columns:
        puntos_area = puntos_df[puntos_df["area_id"] == area.get("area_id")]

    def _sum_caudal(df: pd.DataFrame, tipo: str) -> float:
        if df.empty or "tipo_punto" not in df.columns:
            return 0.0
        mask = df["tipo_punto"].fillna("").astype(str).str.lower() == tipo
        caudales = df.loc[mask, "caudal"] if "caudal" in df.columns else pd.Series(dtype=float)
        return float(
            caudales.apply(_to_float).fillna(0).sum()
        )

    total_inyeccion = _sum_caudal(puntos_area, "inyeccion")
    total_extraccion = _sum_caudal(puntos_area, "extraccion")
    area["caudal_inyeccion_total"] = total_inyeccion
    area["caudal_extraccion_total"] = total_extraccion

    m3_pp_ref = _to_float(area.get("m3_porpersona_594")) or 10.0
    m3_pp_hora_ref = _to_float(area.get("m3_porpersona_hora_594")) or 20.0
    recambio_min = _to_float(area.get("recambio_hora_594_min")) or 0.0
    recambio_max = _to_float(area.get("recambio_hora_594_max")) or 0.0

    m3_pp = (volumen / aforo) if (volumen is not None and aforo and aforo > 0) else None
    area["m3_porpersona"] = m3_pp
    area["m3_porpersona_cumple"] = 1 if (m3_pp is not None and m3_pp >= m3_pp_ref) else 0

    max_caudal = max(total_inyeccion, total_extraccion)
    m3_pp_hora = (max_caudal / aforo) if (aforo and aforo > 0) else None
    area["m3_porpersona_hora"] = m3_pp_hora
    area["m3_porpersona_hora_cumple"] = (
        1 if (m3_pp_hora is not None and m3_pp_hora >= m3_pp_hora_ref) else 0
    )

    recambio_hora = (max_caudal / volumen) if (volumen and volumen > 0) else None
    area["recambio_hora"] = recambio_hora
    area["recambio_hora_cumple"] = (
        1
        if (
            recambio_hora is not None
            and recambio_hora >= recambio_min
            and (recambio_max == 0 or recambio_hora <= recambio_max)
        )
        else 0
    )

    return area


def _parse_date(value) -> Optional[date]:
    if pd.isna(value):
        return None
    if isinstance(value, date):
        return value
    return pd.to_datetime(value).date()


def _parse_time(value) -> Optional[time]:
    if pd.isna(value):
        return None
    if isinstance(value, time):
        return value
    return pd.to_datetime(value).time()


def _dict_from_row(row: pd.Series, expected: List[str]) -> Dict:
    data = {}
    for col in expected:
        data[col] = _clean_value(row.get(col))
    return data


def _insert_visita(cursor, visita: Dict) -> int:
    # Intentar reutilizar una visita existente con misma CUV, fecha, hora y consultor
    lookup_query = (
        "SELECT id_visita FROM visitas "
        "WHERE cuv_visita=%s AND fecha_visita=%s AND hora_visita=%s AND consultor_ist=%s "
        "AND tipo_evaluacion='ventilacion'"
    )
    cursor.execute(
        lookup_query,
        (
            visita["cuv_visita"],
            visita["fecha_visita"],
            visita["hora_visita"],
            visita["consultor_ist"],
        ),
    )
    existing = cursor.fetchone()
    if existing:
        return existing[0]

    insert_query = (
        "INSERT INTO visitas (cuv_visita, fecha_visita, hora_visita, motivo_evaluacion, "
        "nombre_personal_visita, cargo_personal_visita, consultor_ist, note_visita, consultor_cargo, consultor_zonal, tipo_evaluacion) "
        "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 'ventilacion')"
    )
    cursor.execute(
        insert_query,
        (
            visita["cuv_visita"],
            visita["fecha_visita"],
            visita["hora_visita"],
            visita.get("motivo_evaluacion"),
            visita.get("nombre_personal_visita"),
            visita.get("cargo_personal_visita"),
            visita["consultor_ist"],
            visita.get("note_visita"),
            visita.get("consultor_cargo"),
            visita.get("consultor_zonal"),
        ),
    )
    return cursor.lastrowid


def _insert_area(cursor, area: Dict):
    area = _with_defaults(area)
    query = (
        "INSERT INTO v_areas (area_id, visita_id, centro_id, codigo_area, nombre_area, uso, piso_nivel, "
        "largo_m, ancho_m, alto_m, volumen_m3, aforo_permitido, m3_porpersona, m3_porpersona_594, "
        "m3_porpersona_cumple, caudal_inyeccion_total, caudal_extraccion_total, m3_porpersona_hora, "
        "m3_porpersona_hora_594, m3_porpersona_hora_cumple, recambio_hora_594_min, recambio_hora_594_max, "
        "recambio_hora, recambio_hora_cumple, ocupacion_habitual, ventilacion_tipo, ventilacion_sistema, "
        "ventilacion_estado, aberturas, croquis_url, observaciones) "
        "VALUES (%(area_id)s, %(visita_id)s, %(centro_id)s, %(codigo_area)s, %(nombre_area)s, %(uso)s, %(piso_nivel)s, "
        "%(largo_m)s, %(ancho_m)s, %(alto_m)s, %(volumen_m3)s, %(aforo_permitido)s, %(m3_porpersona)s, %(m3_porpersona_594)s, "
        "%(m3_porpersona_cumple)s, %(caudal_inyeccion_total)s, %(caudal_extraccion_total)s, %(m3_porpersona_hora)s, "
        "%(m3_porpersona_hora_594)s, %(m3_porpersona_hora_cumple)s, %(recambio_hora_594_min)s, %(recambio_hora_594_max)s, "
        "%(recambio_hora)s, %(recambio_hora_cumple)s, %(ocupacion_habitual)s, %(ventilacion_tipo)s, %(ventilacion_sistema)s, "
        "%(ventilacion_estado)s, %(aberturas)s, %(croquis_url)s, %(observaciones)s) "
        "ON DUPLICATE KEY UPDATE "
        "codigo_area=VALUES(codigo_area), nombre_area=VALUES(nombre_area), uso=VALUES(uso), piso_nivel=VALUES(piso_nivel), "
        "largo_m=VALUES(largo_m), ancho_m=VALUES(ancho_m), alto_m=VALUES(alto_m), volumen_m3=VALUES(volumen_m3), "
        "aforo_permitido=VALUES(aforo_permitido), m3_porpersona=VALUES(m3_porpersona), m3_porpersona_594=VALUES(m3_porpersona_594), "
        "m3_porpersona_cumple=VALUES(m3_porpersona_cumple), caudal_inyeccion_total=VALUES(caudal_inyeccion_total), "
        "caudal_extraccion_total=VALUES(caudal_extraccion_total), m3_porpersona_hora=VALUES(m3_porpersona_hora), "
        "m3_porpersona_hora_594=VALUES(m3_porpersona_hora_594), m3_porpersona_hora_cumple=VALUES(m3_porpersona_hora_cumple), "
        "recambio_hora_594_min=VALUES(recambio_hora_594_min), recambio_hora_594_max=VALUES(recambio_hora_594_max), "
        "recambio_hora=VALUES(recambio_hora), recambio_hora_cumple=VALUES(recambio_hora_cumple), ocupacion_habitual=VALUES(ocupacion_habitual), "
        "ventilacion_tipo=VALUES(ventilacion_tipo), ventilacion_sistema=VALUES(ventilacion_sistema), ventilacion_estado=VALUES(ventilacion_estado), "
        "aberturas=VALUES(aberturas), croquis_url=VALUES(croquis_url), observaciones=VALUES(observaciones)"
    )
    cursor.execute(query, area)


def _insert_punto(cursor, punto: Dict):
    query = (
        "INSERT INTO v_puntos_medicion (punto_id, evaluacion_id, area_id, codigo_punto, tipo_punto, ubicacion_detalle, "
        "altura_m, distancia_fuente_m, conducto_largo_cm, conducto_ancho_cm, conducto_diametro, seccion_conducto_cm2, "
        "medicion_caudal_1, medicion_caudal_2, medicion_caudal_3, medicion_caudal_4, medicion_caudal_5, medicion_caudal_p, "
        "caudal, fecha_hora, condiciones_ocupacion, puertas_ventanas_abiertas, temperatura_c, humedad_relativa_pct, "
        "croquis_url, observaciones) "
        "VALUES (%(punto_id)s, %(evaluacion_id)s, %(area_id)s, %(codigo_punto)s, %(tipo_punto)s, %(ubicacion_detalle)s, "
        "%(altura_m)s, %(distancia_fuente_m)s, %(conducto_largo_cm)s, %(conducto_ancho_cm)s, %(conducto_diametro)s, %(seccion_conducto_cm2)s, "
        "%(medicion_caudal_1)s, %(medicion_caudal_2)s, %(medicion_caudal_3)s, %(medicion_caudal_4)s, %(medicion_caudal_5)s, %(medicion_caudal_p)s, "
        "%(caudal)s, %(fecha_hora)s, %(condiciones_ocupacion)s, %(puertas_ventanas_abiertas)s, %(temperatura_c)s, %(humedad_relativa_pct)s, "
        "%(croquis_url)s, %(observaciones)s) "
        "ON DUPLICATE KEY UPDATE "
        "codigo_punto=VALUES(codigo_punto), tipo_punto=VALUES(tipo_punto), ubicacion_detalle=VALUES(ubicacion_detalle), "
        "altura_m=VALUES(altura_m), distancia_fuente_m=VALUES(distancia_fuente_m), conducto_largo_cm=VALUES(conducto_largo_cm), "
        "conducto_ancho_cm=VALUES(conducto_ancho_cm), conducto_diametro=VALUES(conducto_diametro), seccion_conducto_cm2=VALUES(seccion_conducto_cm2), "
        "medicion_caudal_1=VALUES(medicion_caudal_1), medicion_caudal_2=VALUES(medicion_caudal_2), medicion_caudal_3=VALUES(medicion_caudal_3), "
        "medicion_caudal_4=VALUES(medicion_caudal_4), medicion_caudal_5=VALUES(medicion_caudal_5), medicion_caudal_p=VALUES(medicion_caudal_p), "
        "caudal=VALUES(caudal), fecha_hora=VALUES(fecha_hora), condiciones_ocupacion=VALUES(condiciones_ocupacion), "
        "puertas_ventanas_abiertas=VALUES(puertas_ventanas_abiertas), temperatura_c=VALUES(temperatura_c), "
        "humedad_relativa_pct=VALUES(humedad_relativa_pct), croquis_url=VALUES(croquis_url), observaciones=VALUES(observaciones)"
    )
    cursor.execute(query, punto)


def _process_visita(cursor, visita_data: Dict, areas_df: pd.DataFrame, puntos_df: pd.DataFrame) -> int:
    visita_id = _insert_visita(cursor, visita_data)

    _insert_ev_ventilacion(
        cursor,
        visita_id,
        {
            "equipo_temp": visita_data.get("equipo_temp"),
            "equipo_vel_air": visita_data.get("equipo_vel_air"),
            "instru_nombre_1": visita_data.get("instru_nombre_1"),
            "instru_marca_1": visita_data.get("instru_marca_1"),
            "instru_modelo_1": visita_data.get("instru_modelo_1"),
            "instru_nserie_1": visita_data.get("instru_nserie_1"),
            "instru_ncertificado_1": visita_data.get("instru_ncertificado_1"),
            "instru_nombre_2": visita_data.get("instru_nombre_2"),
            "instru_marca_2": visita_data.get("instru_marca_2"),
            "instru_modelo_2": visita_data.get("instru_modelo_2"),
            "instru_nserie_2": visita_data.get("instru_nserie_2"),
            "instru_ncertificado_2": visita_data.get("instru_ncertificado_2"),
        },
    )

    area_cols = [
        "area_id",
        "codigo_area",
        "nombre_area",
        "uso",
        "piso_nivel",
        "largo_m",
        "ancho_m",
        "alto_m",
        "volumen_m3",
        "aforo_permitido",
        "m3_porpersona",
        "m3_porpersona_594",
        "m3_porpersona_cumple",
        "caudal_inyeccion_total",
        "caudal_extraccion_total",
        "m3_porpersona_hora",
        "m3_porpersona_hora_594",
        "m3_porpersona_hora_cumple",
        "recambio_hora_594_min",
        "recambio_hora_594_max",
        "recambio_hora",
        "recambio_hora_cumple",
        "ocupacion_habitual",
        "ventilacion_tipo",
        "ventilacion_sistema",
        "ventilacion_estado",
        "aberturas",
        "croquis_url",
        "observaciones",
    ]

    for idx, row in areas_df.iterrows():
        area_data = _dict_from_row(row, area_cols)
        if not area_data.get("area_id"):
            area_data["area_id"] = f"{visita_id}-A{idx + 1}"
        area_data["visita_id"] = visita_id
        area_data["centro_id"] = visita_data["cuv_visita"]
        area_data = _compute_area_calculations(area_data, puntos_df)
        _insert_area(cursor, area_data)

    if not puntos_df.empty:
        punto_cols = [
            "punto_id",
            "area_id",
            "codigo_punto",
            "tipo_punto",
            "ubicacion_detalle",
            "altura_m",
            "distancia_fuente_m",
            "conducto_largo_cm",
            "conducto_ancho_cm",
            "conducto_diametro",
            "seccion_conducto_cm2",
            "medicion_caudal_1",
            "medicion_caudal_2",
            "medicion_caudal_3",
            "medicion_caudal_4",
            "medicion_caudal_5",
            "medicion_caudal_p",
            "caudal",
            "fecha_hora",
            "condiciones_ocupacion",
            "puertas_ventanas_abiertas",
            "temperatura_c",
            "humedad_relativa_pct",
            "croquis_url",
            "observaciones",
        ]

        for idx, row in puntos_df.iterrows():
            punto_data = _dict_from_row(row, punto_cols)
            if not punto_data.get("punto_id"):
                punto_data["punto_id"] = f"{visita_id}-P{idx + 1}"
            punto_data["evaluacion_id"] = visita_id
            fecha_val = punto_data.get("fecha_hora")
            if isinstance(fecha_val, str) and fecha_val:
                punto_data["fecha_hora"] = pd.to_datetime(fecha_val)
            _insert_punto(cursor, punto_data)

    return visita_id


def _build_areas_from_wide_row(row: pd.Series) -> pd.DataFrame:
    areas: List[Dict] = []
    for idx in range(1, 12):
        name = _clean_value(row.get(f"area_{idx}"))
        if not name:
            continue

        area: Dict = {
            "codigo_area": _clean_value(row.get(f"area_{idx}_bis")) or name,
            "nombre_area": name,
            "uso": _clean_value(row.get(f"area_{idx}_desc")),
            "largo_m": _clean_value(row.get(f"largo_{idx}")),
            "ancho_m": _clean_value(row.get(f"ancho_{idx}")),
            "alto_m": _clean_value(row.get(f"alto_{idx}")),
            "volumen_m3": _clean_value(row.get(f"volumen_{idx}")),
            "aforo_permitido": _clean_value(row.get(f"n_personas_{idx}")),
            "m3_porpersona": _clean_value(row.get(f"m3xpersona_{idx}")),
        }
        areas.append(area)
    return pd.DataFrame(areas)


def _parse_visita_row(row: pd.Series) -> Dict:
    visita_data = {
        "cuv_visita": _clean_value(row.get("cuv_visita")),
        "fecha_visita": _parse_date(_clean_value(row.get("fecha_de_visita"))),
        "hora_visita": _parse_time(_clean_value(row.get("horario_de_medicion"))),
        "motivo_evaluacion": None,
        "nombre_personal_visita": _clean_value(row.get("nombre_persona_empresa")),
        "cargo_personal_visita": _clean_value(row.get("cargo_persona_empresa")),
        "consultor_ist": _clean_value(row.get("nombre_profesional_ist")),
        "note_visita": None,
        "consultor_cargo": _clean_value(row.get("cargo_profesional_ist")),
        "consultor_zonal": None,
        "equipo_temp": _clean_value(row.get("equipo_temp")),
        "equipo_vel_air": _clean_value(row.get("equipo_vel_air")),
        "instru_nombre_1": _clean_value(row.get("instru_nombre_1")),
        "instru_marca_1": _clean_value(row.get("instru_marca_1")),
        "instru_modelo_1": _clean_value(row.get("instru_modelo_1")),
        "instru_nserie_1": _clean_value(row.get("instru_nserie_1")),
        "instru_ncertificado_1": _clean_value(row.get("instru_ncertificado_1")),
        "instru_nombre_2": _clean_value(row.get("instru_nombre_2")),
        "instru_marca_2": _clean_value(row.get("instru_marca_2")),
        "instru_modelo_2": _clean_value(row.get("instru_modelo_2")),
        "instru_nserie_2": _clean_value(row.get("instru_nserie_2")),
        "instru_ncertificado_2": _clean_value(row.get("instru_ncertificado_2")),
    }
    return visita_data


def cargar_archivo(ruta_excel: str) -> int:
    xls = pd.ExcelFile(ruta_excel)

    conn = get_connection()
    cursor = conn.cursor()

    try:
        visitas_ids: List[int] = []

        if "Resultados" in xls.sheet_names and not set(xls.sheet_names).intersection(
            {"visita", "visitas", "areas", "v_areas", "puntos", "v_puntos"}
        ):
            resultados_df = xls.parse("Resultados")
            resultados_df = resultados_df.rename(columns=_normalize_column_name)

            for _, row in resultados_df.iterrows():
                visita_data = _parse_visita_row(row)
                areas_df = _build_areas_from_wide_row(row)
                puntos_df = pd.DataFrame()
                visita_id = _process_visita(cursor, visita_data, areas_df, puntos_df)
                visitas_ids.append(visita_id)
        else:
            visita_sheet = _first_available_sheet(xls, ["visita", "visitas"])
            visita_df = xls.parse(visita_sheet)
            visita_df = _apply_aliases(
                visita_df,
                {
                    "cuv": "cuv_visita",
                    "cuv_visita": "cuv_visita",
                    "fecha": "fecha_visita",
                    "fecha_visita": "fecha_visita",
                    "hora": "hora_visita",
                    "hora_visita": "hora_visita",
                    "motivo": "motivo_evaluacion",
                    "motivo_evaluacion": "motivo_evaluacion",
                    "personal_visita": "nombre_personal_visita",
                    "nombre_personal_visita": "nombre_personal_visita",
                    "cargo": "cargo_personal_visita",
                    "cargo_personal_visita": "cargo_personal_visita",
                    "consultor": "consultor_ist",
                    "consultor_ist": "consultor_ist",
                    "nota": "note_visita",
                    "note_visita": "note_visita",
                    "consultor_cargo": "consultor_cargo",
                    "consultor_zonal": "consultor_zonal",
                    "equipo_temp": "equipo_temp",
                    "equipo_vel_air": "equipo_vel_air",
                    "equipo_velocidad": "equipo_vel_air",
                    "equipo_ventilacion": "equipo_vel_air",
                    "instru_nombre_1": "instru_nombre_1",
                    "instru_marca_1": "instru_marca_1",
                    "instru_modelo_1": "instru_modelo_1",
                    "instru_nserie_1": "instru_nserie_1",
                    "instru_ncertificado_1": "instru_ncertificado_1",
                    "instru_nombre_2": "instru_nombre_2",
                    "instru_marca_2": "instru_marca_2",
                    "instru_modelo_2": "instru_modelo_2",
                    "instru_nserie_2": "instru_nserie_2",
                    "instru_ncertificado_2": "instru_ncertificado_2",
                },
            )
            if visita_df.empty:
                raise ValueError("La hoja de visita está vacía")

            visita_cols = [
                "cuv_visita",
                "fecha_visita",
                "hora_visita",
                "motivo_evaluacion",
                "nombre_personal_visita",
                "cargo_personal_visita",
                "consultor_ist",
                "note_visita",
                "consultor_cargo",
                "consultor_zonal",
                "equipo_temp",
                "equipo_vel_air",
                "instru_nombre_1",
                "instru_marca_1",
                "instru_modelo_1",
                "instru_nserie_1",
                "instru_ncertificado_1",
                "instru_nombre_2",
                "instru_marca_2",
                "instru_modelo_2",
                "instru_nserie_2",
                "instru_ncertificado_2",
            ]
            visita_data = _dict_from_row(visita_df.iloc[0], visita_cols)
            visita_data["fecha_visita"] = _parse_date(visita_data["fecha_visita"])
            visita_data["hora_visita"] = _parse_time(visita_data["hora_visita"])

            areas_sheet = _first_available_sheet(xls, ["areas", "v_areas"])
            areas_df = xls.parse(areas_sheet)
            areas_df = _apply_aliases(
                areas_df,
                {
                    "id": "area_id",
                    "area_id": "area_id",
                    "codigo": "codigo_area",
                    "codigo_area": "codigo_area",
                    "nombre": "nombre_area",
                    "nombre_area": "nombre_area",
                    "uso": "uso",
                    "piso": "piso_nivel",
                    "piso_nivel": "piso_nivel",
                    "largo_m": "largo_m",
                    "ancho_m": "ancho_m",
                    "alto_m": "alto_m",
                    "volumen_m3": "volumen_m3",
                    "aforo": "aforo_permitido",
                    "aforo_permitido": "aforo_permitido",
                    "m3_persona": "m3_porpersona",
                    "m3_porpersona": "m3_porpersona",
                    "m3_porpersona_594": "m3_porpersona_594",
                    "m3_porpersona_cumple": "m3_porpersona_cumple",
                    "caudal_inyeccion_total": "caudal_inyeccion_total",
                    "caudal_extraccion_total": "caudal_extraccion_total",
                    "m3_porpersona_hora": "m3_porpersona_hora",
                    "m3_porpersona_hora_594": "m3_porpersona_hora_594",
                    "m3_porpersona_hora_cumple": "m3_porpersona_hora_cumple",
                    "recambio_hora_594_min": "recambio_hora_594_min",
                    "recambio_hora_594_max": "recambio_hora_594_max",
                    "recambio_hora": "recambio_hora",
                    "recambio_hora_cumple": "recambio_hora_cumple",
                    "ocupacion_habitual": "ocupacion_habitual",
                    "ventilacion_tipo": "ventilacion_tipo",
                    "ventilacion_sistema": "ventilacion_sistema",
                    "ventilacion_estado": "ventilacion_estado",
                    "aberturas": "aberturas",
                    "croquis_url": "croquis_url",
                    "observaciones": "observaciones",
                },
            )

            puntos_sheet = _first_available_sheet(xls, ["puntos", "v_puntos"])
            puntos_df = xls.parse(puntos_sheet)
            puntos_df = _apply_aliases(
                puntos_df,
                {
                    "id": "punto_id",
                    "punto_id": "punto_id",
                    "area_id": "area_id",
                    "codigo": "codigo_punto",
                    "codigo_punto": "codigo_punto",
                    "tipo": "tipo_punto",
                    "tipo_punto": "tipo_punto",
                    "ubicacion_detalle": "ubicacion_detalle",
                    "altura_m": "altura_m",
                    "distancia_fuente_m": "distancia_fuente_m",
                    "conducto_largo_cm": "conducto_largo_cm",
                    "conducto_ancho_cm": "conducto_ancho_cm",
                    "conducto_diametro": "conducto_diametro",
                    "seccion_conducto_cm2": "seccion_conducto_cm2",
                    "medicion_caudal_1": "medicion_caudal_1",
                    "medicion_caudal_2": "medicion_caudal_2",
                    "medicion_caudal_3": "medicion_caudal_3",
                    "medicion_caudal_4": "medicion_caudal_4",
                    "medicion_caudal_5": "medicion_caudal_5",
                    "medicion_caudal_p": "medicion_caudal_p",
                    "caudal": "caudal",
                    "fecha_hora": "fecha_hora",
                    "condiciones_ocupacion": "condiciones_ocupacion",
                    "puertas_ventanas_abiertas": "puertas_ventanas_abiertas",
                    "temperatura_c": "temperatura_c",
                    "humedad_relativa_pct": "humedad_relativa_pct",
                    "croquis_url": "croquis_url",
                    "observaciones": "observaciones",
                },
            )

            visita_id = _process_visita(cursor, visita_data, areas_df, puntos_df)
            visitas_ids.append(visita_id)

        conn.commit()
        return visitas_ids[-1] if visitas_ids else 0
    except Exception:
        conn.rollback()
        raise
    finally:
        cursor.close()
        conn.close()


if __name__ == "__main__":
    import argparse

    default_path = os.path.join(os.path.dirname(__file__), "pruebacarga.xlsx")

    parser = argparse.ArgumentParser(description="Carga evaluaciones de ventilación desde un Excel")
    parser.add_argument(
        "archivo",
        nargs="?",
        default=default_path,
        help=(
            "Ruta al archivo .xlsx a cargar (por defecto se usa etl/3_carga/pruebacarga.xlsx)"
        ),
    )
    args = parser.parse_args()

    visita_id = cargar_archivo(args.archivo)
    print(f"Datos cargados para la visita #{visita_id}")
