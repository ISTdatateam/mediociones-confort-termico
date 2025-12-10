"""Genera informes DOCX a partir de un archivo de texto con IDs de visita.

Este script lee un archivo ``.txt`` que contenga un ID de visita por línea y
construye los informes Word utilizando las funciones existentes del proyecto.
Los documentos generados se organizan en carpetas por consultor IST.
"""
from __future__ import annotations

import argparse
import logging
from pathlib import Path
import sys
from typing import Iterable, List, Optional

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

from utils.doc_utils import generar_informe_en_word, generar_informe_ventilacion_en_word
from utils.helpers import (
    get_areas_ventilacion_df,
    get_ct,
    get_equipos,
    get_mediciones,
    get_puntos_ventilacion_df,
    get_visita,
)

LOG_FILE = Path(__file__).with_name("generar_informes.txt")

logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)


def _normalizar_tipo(valor: str) -> str:
    """Normaliza el tipo de informe y mapea variantes a valores admitidos."""

    traducciones = str.maketrans(
        {
            "á": "a",
            "é": "e",
            "í": "i",
            "ó": "o",
            "ú": "u",
            "Á": "a",
            "É": "e",
            "Í": "i",
            "Ó": "o",
            "Ú": "u",
        }
    )
    valor_normalizado = " ".join(valor.strip().translate(traducciones).lower().split())

    if valor_normalizado in {"auto"}:
        return "auto"

    if "ventil" in valor_normalizado:
        return "ventilacion"

    if "confort" in valor_normalizado:
        return "confort"

    return valor_normalizado


def _leer_ids_desde_txt(ruta_txt: Path) -> List[int]:
    ids: List[int] = []
    with ruta_txt.open("r", encoding="utf-8") as archivo:
        for linea in archivo:
            linea = linea.strip()
            if not linea or linea.startswith("#"):
                continue
            try:
                ids.append(int(linea))
            except ValueError:
                logging.warning("Línea ignorada (no es un ID válido): %s", linea)
    return ids


def _nombre_carpeta_consultor(df_visita: pd.DataFrame) -> str:
    fila = df_visita.iloc[0]
    for campo in ("consultor_nombre", "consultor_ist"):
        valor = fila.get(campo)
        if isinstance(valor, str) and valor.strip():
            return valor.strip()
    return "desconocido"


def _generar_informe_confort(
    id_visita: int,
    destino_base: Path,
    df_visita: pd.DataFrame,
    cuv: int,
    df_centro: pd.DataFrame,
) -> Optional[Path]:
    df_mediciones = get_mediciones(id_visita)
    if df_mediciones.empty:
        logging.error("No hay mediciones para la visita %s", id_visita)
        return None

    df_equipos = get_equipos()
    doc_bytes = generar_informe_en_word(df_centro, df_visita, df_mediciones, df_equipos)
    if not doc_bytes:
        logging.error("No se pudo generar el informe para la visita %s", id_visita)
        return None

    carpeta_consultor = destino_base / _nombre_carpeta_consultor(df_visita)
    carpeta_consultor.mkdir(parents=True, exist_ok=True)

    nombre_archivo = f"informe_ventilacion_CECO_{cuv}_local_{df_centro['nombre_ct'].iloc[0]}.docx"

    ruta_archivo = carpeta_consultor / nombre_archivo
    with ruta_archivo.open("wb") as salida:
        salida.write(doc_bytes.getvalue())

    return ruta_archivo


def _generar_informe_ventilacion(
    id_visita: int,
    destino_base: Path,
    df_visita: pd.DataFrame,
    cuv: int,
    df_centro: pd.DataFrame,
) -> Optional[Path]:
    df_areas = get_areas_ventilacion_df(id_visita)
    if df_areas.empty:
        logging.error(
            "No existen áreas de ventilación registradas para la visita %s", id_visita
        )
        return None

    df_puntos_raw = get_puntos_ventilacion_df(id_visita)
    if df_puntos_raw is None:
        df_puntos = pd.DataFrame()
    elif isinstance(df_puntos_raw, pd.DataFrame):
        df_puntos = df_puntos_raw
    else:
        df_puntos = pd.DataFrame(df_puntos_raw)

    df_equipos = get_equipos()

    def _normalizar_cumplimiento(columna: str) -> Optional[pd.Series]:
        if columna not in df_areas.columns:
            return None

        serie = df_areas.get(columna)
        if serie is None:
            return None

        # Normaliza textos como "Cumple" / "No cumple" o valores booleanos/númericos.
        texto_normalizado = (
            serie.astype(str).str.strip().str.lower().replace({"sí": "si"})
        )
        mapa_cumple = {
            "cumple": 1,
            "si": 1,
            "true": 1,
            "1": 1,
            "no cumple": 0,
            "no": 0,
            "false": 0,
            "0": 0,
        }
        mapeado = texto_normalizado.map(mapa_cumple)

        # Completa con intentos numéricos para valores no mapeados.
        numerico = pd.to_numeric(serie, errors="coerce")
        combinado = mapeado.combine_first(numerico)

        if combinado.dropna().empty:
            return None

        # Mantiene los valores no informados como nulos para no exigir puntos innecesariamente.
        return combinado.astype("Int64")

    areas_requieren_puntos: pd.DataFrame
    cumple_series: Optional[pd.Series] = None
    for columna in ("m3_porpersona_cumple", "m3_persona", "m3_persona_cumple"):
        cumple_series = _normalizar_cumplimiento(columna)
        if cumple_series is not None:
            if columna != "m3_porpersona_cumple":
                logging.info(
                    "Usando la columna %s para validar m3/persona en la visita %s.",
                    columna,
                    id_visita,
                )
            break

    if cumple_series is None or cumple_series.dropna().empty:
        logging.info(
            "No se encontró una columna de cumplimiento m3/persona reconocible en la visita %s; se omite la validación de puntos requeridos.",
            id_visita,
        )
        areas_requieren_puntos = pd.DataFrame()
    else:
        areas_requieren_puntos = df_areas.loc[cumple_series == 0]

    if not areas_requieren_puntos.empty:
        if df_puntos.empty:
            nombres_requeridos = areas_requieren_puntos.get(
                "nombre_area", pd.Series(dtype=str)
            )
            nombres = ", ".join(nombre for nombre in nombres_requeridos if nombre)
            if not nombres:
                nombres = ", ".join(
                    areas_requieren_puntos.get(
                        "codigo_area", pd.Series(dtype=str)
                    ).astype(str)
                )
            logging.error(
                "Registra puntos de medición para las áreas que no cumplen el "
                "indicador de m³/persona%s",
                f": {nombres}." if nombres else ".",
            )
            return None

        if "area_id" in df_puntos.columns:
            puntos_por_area = df_puntos.groupby("area_id").size()
            areas_sin_puntos = [
                area_id
                for area_id in areas_requieren_puntos.get(
                    "area_id", pd.Series(dtype=int)
                ).tolist()
                if puntos_por_area.get(area_id, 0) == 0
            ]
        else:
            logging.info(
                "No se encontró columna 'area_id' en puntos de medición; se omite la validación de puntos obligatorios."
            )
            areas_sin_puntos = []

        if areas_sin_puntos:
            nombres_faltantes = []
            for area_id in areas_sin_puntos:
                area_fila = areas_requieren_puntos[
                    areas_requieren_puntos["area_id"] == area_id
                ]
                nombre = area_fila.get("nombre_area")
                if nombre is not None and not nombre.empty:
                    nombres_faltantes.append(str(nombre.iloc[0]))
                else:
                    codigo = area_fila.get("codigo_area")
                    if codigo is not None and not codigo.empty:
                        nombres_faltantes.append(str(codigo.iloc[0]))
                    else:
                        nombres_faltantes.append(str(area_id))

            lista_nombres = ", ".join(nombres_faltantes)
            logging.error(
                "Faltan puntos de medición para completar las áreas sin "
                "cumplimiento: %s.",
                lista_nombres,
            )
            return None

    doc_bytes = generar_informe_ventilacion_en_word(
        df_centro, df_visita, df_areas, df_puntos, df_equipos
    )
    if not doc_bytes:
        logging.error(
            "No se pudo generar el informe de ventilación para la visita %s", id_visita
        )
        return None

    carpeta_consultor = destino_base / _nombre_carpeta_consultor(df_visita)
    carpeta_consultor.mkdir(parents=True, exist_ok=True)

    nombre_archivo = f"informe_ventilacion_cuv_{cuv}_visita_{id_visita}.docx"
    ruta_archivo = carpeta_consultor / nombre_archivo
    with ruta_archivo.open("wb") as salida:
        salida.write(doc_bytes.getvalue())

    return ruta_archivo


def _generar_informe(
    id_visita: int, destino_base: Path, tipo: str = "auto"
) -> Optional[Path]:
    df_visita = get_visita(id_visita)
    if df_visita.empty:
        logging.error("No se encontró la visita con ID %s", id_visita)
        return None

    tipo_lower = _normalizar_tipo(tipo)

    cuv = df_visita.iloc[0].get("cuv_visita")
    if pd.isna(cuv):
        logging.error("La visita %s no tiene CUV asociado", id_visita)
        return None

    df_centro = pd.DataFrame(get_ct(cuv))
    if df_centro.empty:
        logging.error("No se encontró el centro de trabajo para CUV %s", cuv)
        return None

    tipo_evaluacion_visita = _normalizar_tipo(
        str(df_visita.iloc[0].get("tipo_evaluacion", ""))
    )
    if tipo_lower == "auto":
        tipo_lower = tipo_evaluacion_visita or "confort"
    elif tipo_evaluacion_visita and tipo_lower != tipo_evaluacion_visita:
        logging.error(
            "La visita %s está registrada con tipo '%s' (columna tipo_evaluacion de la tabla visitas) "
            "y no coincide con el tipo solicitado '%s'.",
            id_visita,
            tipo_evaluacion_visita,
            tipo_lower,
        )
        return None

    if tipo_lower == "ventilacion":
        return _generar_informe_ventilacion(
            id_visita, destino_base, df_visita, cuv, df_centro
        )

    if tipo_lower != "confort":
        logging.error("Tipo de informe no soportado: %s", tipo)
        return None

    return _generar_informe_confort(
        id_visita, destino_base, df_visita, cuv, df_centro
    )


def generar_informes(ids: Iterable[int], salida: Path, tipo: str = "auto") -> List[Path]:
    salida.mkdir(parents=True, exist_ok=True)
    generados: List[Path] = []
    for id_visita in ids:
        logging.info("Generando informe para visita %s", id_visita)
        ruta = _generar_informe(id_visita, salida, tipo=tipo)
        if ruta:
            logging.info("Informe guardado en %s", ruta)
            generados.append(ruta)
        else:
            logging.error(
                "No se pudo generar el informe solicitado para la visita %s", id_visita
            )
    return generados


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--archivo",
        type=Path,
        default=Path(__file__).with_name("informes.txt"),
        help="Ruta al archivo .txt que contiene un ID de visita por línea (por defecto scripts/informes.txt)",
    )
    parser.add_argument(
        "--salida",
        type=Path,
        default=Path("informes_generados"),
        help="Directorio donde se guardarán los informes generados",
    )
    parser.add_argument(
        "--tipo",
        type=_normalizar_tipo,
        choices=["confort", "ventilacion", "auto"],
        default="auto",
        help=(
            "Tipo de informe a generar: confort térmico, ventilación o auto para "
            "inferirlo desde la visita"
        ),
    )
    return parser.parse_args()


if __name__ == "__main__":
    argumentos = parse_args()
    ids = _leer_ids_desde_txt(argumentos.archivo)
    if not ids:
        logging.error("El archivo no contiene IDs de visita válidos.")
    else:
        rutas = generar_informes(ids, argumentos.salida, tipo=argumentos.tipo)
        logging.info("Se generaron %s informe(s).", len(rutas))
