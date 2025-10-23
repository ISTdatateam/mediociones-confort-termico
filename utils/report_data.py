"""Data formatting and analysis helpers for report generation."""
from datetime import date, datetime
from typing import Iterable, List, Optional, Tuple

import pandas as pd
from pythermalcomfort.models import pmv_ppd_iso


def formatear_fecha(fecha: object) -> str:
    """Format a date-like value using the ``dd-mm-YYYY`` pattern."""
    if not fecha:
        return ""
    if isinstance(fecha, (datetime, date)):
        return fecha.strftime("%d-%m-%Y")
    return datetime.strptime(str(fecha), "%Y-%m-%d").strftime("%d-%m-%Y")


def ftemp(valor: object) -> str:
    """Return a string with decimal separator replaced by comma."""
    if valor in (None, ""):
        return ""
    return str(valor).replace(".", ",")


def format_decimal(valor: object, decimales: int = 2, sufijo: str = "") -> str:
    """Return a number formatted using a comma decimal separator."""
    if valor in (None, ""):
        return ""
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return str(valor)
    texto = f"{numero:.{decimales}f}".replace(".", ",")
    return f"{texto}{sufijo}" if sufijo else texto


def format_columns(df: pd.DataFrame, columns: Iterable[str], mode: str = "title") -> pd.DataFrame:
    """Strip whitespace and apply casing to the provided columns."""
    for col in columns:
        if col not in df.columns:
            continue
        df[col] = df[col].astype(str).str.strip()
        if mode == "title":
            df[col] = df[col].str.lower().str.title()
        elif mode == "capitalize":
            df[col] = df[col].str.lower().str.capitalize()
        elif mode == "upper":
            df[col] = df[col].str.upper()
        else:
            raise ValueError(f"Modo '{mode}' no reconocido. Use 'title', 'capitalize' o 'upper'.")
    return df


def interpret_pmv(pmv_value: float) -> str:
    """Return ``CUMPLE`` when the PMV value falls inside the comfort band."""
    if pmv_value <= -1 or pmv_value >= 1:
        return "NO CUMPLE"
    return "CUMPLE"


def calcular_analisis_area(group: pd.DataFrame) -> str:
    """Determine the comfort status for a group of measurements."""
    if len(group) > 1:
        avg_t_bul = group["t_bul_seco"].astype(float).mean()
        avg_t_globo = group["t_globo"].astype(float).mean()
        avg_hum = group["hum_rel"].astype(float).mean()
        avg_vel = group["vel_air"].astype(float).mean()
        avg_met = group["met"].astype(float).mean() or 1.1
        avg_clo = group["clo"].astype(float).mean() or 0.5
        try:
            results = pmv_ppd_iso(
                tdb=avg_t_bul,
                tr=avg_t_globo,
                vr=avg_vel,
                rh=avg_hum,
                met=avg_met,
                clo=avg_clo,
                model="7730-2005",
                limit_inputs=False,
            )
            pmv = results.pmv if hasattr(results, "pmv") else results.get("pmv", 0)
            analisis = interpret_pmv(float(pmv))
        except Exception:
            analisis = "NO CUMPLE"
    else:
        analisis = group.iloc[0].get("resultado_medicion", "NO CUMPLE").upper()
    return analisis


def procesar_areas(df_mediciones: pd.DataFrame) -> Tuple[List[str], List[str]]:
    """Return the areas that comply or fail according to measurements."""
    areas_cumplen: List[str] = []
    areas_no_cumplen: List[str] = []
    if df_mediciones.empty:
        return areas_cumplen, areas_no_cumplen
    grouped = df_mediciones.groupby("nombre_area")
    for area, group in grouped:
        analisis = calcular_analisis_area(group)
        if analisis == "CUMPLE":
            areas_cumplen.append(area)
        else:
            areas_no_cumplen.append(area)
    return areas_cumplen, areas_no_cumplen


def _float_or_none(value: object) -> Optional[float]:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _m3pp_row_ok(row: pd.Series) -> Optional[bool]:
    m3pp = _float_or_none(row.get("m3_porpersona"))
    if m3pp is None:
        largo = _float_or_none(row.get("largo_m"))
        ancho = _float_or_none(row.get("ancho_m"))
        alto = _float_or_none(row.get("alto_m"))
        n = _float_or_none(row.get("aforo_permitido")) or _float_or_none(row.get("nmp"))
        if None in (largo, ancho, alto) or not n:
            return None
        m3pp = (largo * ancho * alto) / n
    return m3pp >= 10


def todas_cumplen_m3(df_areas: pd.DataFrame) -> bool:
    if df_areas.empty:
        return False
    checks: List[bool] = []
    for _, row in df_areas.iterrows():
        ok = _m3pp_row_ok(row)
        if ok is None:
            return False
        checks.append(ok)
    return all(checks)


__all__ = [
    "formatear_fecha",
    "ftemp",
    "format_decimal",
    "format_columns",
    "interpret_pmv",
    "calcular_analisis_area",
    "procesar_areas",
    "todas_cumplen_m3",
]
