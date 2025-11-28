from io import BytesIO
from datetime import datetime, date
from typing import Union

import pandas as pd
import streamlit as st

from utils.helpers import (
    get_ct,
    get_visita,
    get_visitas_por_cuv,
    get_areas_ventilacion_df,
    get_puntos_ventilacion_df,
    get_equipos,
)
from utils.doc_utils import generar_informe_ventilacion_en_word
from utils.report_data import formatear_fecha


def _normalizar_fecha(valor):
    if isinstance(valor, (datetime, date)):
        return valor
    if not valor:
        return None
    try:
        return datetime.strptime(str(valor), "%Y-%m-%d")
    except ValueError:
        return None


def generar_descarga_informe(cuv: Union[str, int], visita_id: int) -> BytesIO:

    """Genera el archivo DOCX para un informe de ventilación.

    Parameters
    ----------
    cuv:
        Identificador del centro de trabajo. Se acepta el valor numérico en formato
        ``str`` o ``int``; cualquier otro tipo provocará un ``ValueError``.
    visita_id:
        Identificador de la visita de ventilación cuyos datos se utilizarán para
        poblar el informe.

    Returns
    -------
    BytesIO
        Flujo en memoria posicionado al inicio con el documento Word listo para
        ser entregado en una descarga.

    Raises
    ------
    ValueError
        Cuando el CUV no es convertible a entero o cuando falta información clave
        (centro, visita, áreas o puntos) para construir el informe.
    """

    try:
        cuv_int = int(cuv)
    except (TypeError, ValueError) as exc:
        raise ValueError("El CUV ingresado no es válido.") from exc

    df_centro = pd.DataFrame(get_ct(cuv_int))
    if df_centro.empty:
        raise ValueError(
            "No se encontró información del centro de trabajo asociado al CUV ingresado."
        )

    df_visita = get_visita(visita_id)
    if df_visita.empty:
        raise ValueError("No se encontró la visita seleccionada.")

    df_areas = get_areas_ventilacion_df(visita_id)
    if df_areas.empty:
        raise ValueError(
            "No existen áreas de ventilación registradas para la visita seleccionada."
        )

    df_puntos = get_puntos_ventilacion_df(visita_id)
    if not isinstance(df_puntos, pd.DataFrame):
        df_puntos = pd.DataFrame(df_puntos or [])

    df_equipos = get_equipos()

    cumple_series = pd.to_numeric(
        df_areas.get("m3_porpersona_cumple"), errors="coerce"
    ).fillna(0)
    areas_requieren_puntos = df_areas.loc[cumple_series.astype(int) == 0]

    if areas_requieren_puntos.empty:
        pass
    else:
        if df_puntos.empty:
            nombres_requeridos = areas_requieren_puntos.get("nombre_area", pd.Series(dtype=str))
            nombres = ", ".join(nombre for nombre in nombres_requeridos if nombre)
            if not nombres:
                nombres = ", ".join(
                    areas_requieren_puntos.get("codigo_area", pd.Series(dtype=str)).astype(str)
                )
            raise ValueError(
                "Registra puntos de medición para las áreas que no cumplen el indicador de m³/persona"
                + (f": {nombres}." if nombres else ".")
            )

        if "area_id" in df_puntos.columns:
            puntos_por_area = df_puntos.groupby("area_id").size()
        else:
            puntos_por_area = pd.Series(dtype=int)

        areas_sin_puntos = [
            area_id
            for area_id in areas_requieren_puntos.get("area_id", pd.Series(dtype=int)).tolist()
            if puntos_por_area.get(area_id, 0) == 0
        ]

        if areas_sin_puntos:
            nombres_faltantes = []
            for area_id in areas_sin_puntos:
                area_fila = areas_requieren_puntos[areas_requieren_puntos["area_id"] == area_id]
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
            raise ValueError(
                "Faltan puntos de medición para completar las áreas sin cumplimiento: "
                f"{lista_nombres}."
            )

    informe_docx = generar_informe_ventilacion_en_word(
        df_centro, df_visita, df_areas, df_puntos, df_equipos
    )
    if isinstance(informe_docx, BytesIO):
        informe_docx.seek(0)
        return informe_docx

    buffer = BytesIO()
    buffer.write(informe_docx)
    buffer.seek(0)
    return buffer


def generar_informe_ventilacion(cuv, id_visita):
    try:
        return generar_descarga_informe(cuv, id_visita)
    except ValueError as error:
        st.error(str(error))
        return None


def generar_informe_unit(cuv):
    try:
        cuv_int = int(cuv)
    except (TypeError, ValueError):
        st.error("El CUV ingresado no es válido.")
        return None

    df_visitas = get_visitas_por_cuv(cuv_int)
    if df_visitas.empty:
        st.error("No se encontraron visitas registradas para el CUV indicado.")
        return None

    df_filtradas = df_visitas.copy()
    if "tipo_evaluacion" in df_filtradas.columns:
        df_filtradas["tipo_evaluacion"] = df_filtradas["tipo_evaluacion"].astype(str)
        df_filtradas = df_filtradas[df_filtradas["tipo_evaluacion"].str.lower() == "ventilacion"]

    if df_filtradas.empty:
        st.error("No existen visitas de ventilación asociadas al CUV indicado.")
        return None

    visita_reciente = df_filtradas.iloc[0]
    visita_id = int(visita_reciente["id_visita"])
    try:
        return generar_descarga_informe(cuv_int, visita_id)
    except ValueError as error:
        st.error(str(error))
        return None


def main():
    st.header("Informes de Ventilación")
    st.write("Versión 1.0 - Generación de informes técnicos de ventilación")

    cuv_input = st.text_input("Ingresa el CUV", key="ventilacion_cuv")
    informe_generado = None

    visitas_filtradas = pd.DataFrame()
    cuv_int = None

    if cuv_input:
        try:
            cuv_int = int(cuv_input)
        except ValueError:
            st.error("Debes ingresar un CUV numérico.")
        else:
            df_visitas = get_visitas_por_cuv(cuv_int)
            if df_visitas.empty:
                st.info("No se encontraron visitas registradas para el CUV ingresado.")
            else:
                visitas_filtradas = df_visitas.copy()
                if "tipo_evaluacion" in visitas_filtradas.columns:
                    visitas_filtradas["tipo_evaluacion"] = visitas_filtradas["tipo_evaluacion"].astype(str)
                    visitas_filtradas = visitas_filtradas[visitas_filtradas["tipo_evaluacion"].str.lower() == "ventilacion"]

                if visitas_filtradas.empty:
                    st.info("No existen visitas de ventilación para el CUV indicado.")

    if not visitas_filtradas.empty and cuv_int is not None:
        opciones = {}
        for _, row in visitas_filtradas.iterrows():
            visita_id = int(row["id_visita"])
            fecha_valor = row.get("fecha_visita")
            fecha_dt = _normalizar_fecha(fecha_valor)
            fecha_texto = formatear_fecha(fecha_dt) if fecha_dt else (str(fecha_valor) if fecha_valor else "Sin fecha")
            motivo = row.get("motivo_evaluacion", "")
            label = f"Visita {visita_id} - {fecha_texto}"
            if motivo:
                label += f" ({motivo})"
            opciones[label] = visita_id

        if opciones:
            with st.form("form_informe_ventilacion"):
                opcion_seleccionada = st.selectbox("Selecciona la visita de ventilación", options=list(opciones.keys()))
                submit = st.form_submit_button("Generar informe")

            if submit:
                visita_id = opciones[opcion_seleccionada]
                informe_generado = generar_informe_ventilacion(cuv_int, visita_id)
        else:
            st.info("No se encontraron visitas válidas para generar el informe.")

    st.markdown("---")
    st.subheader("Generar con la visita más reciente")
    if st.button("Generar informe automático", type="primary"):
        informe_generado = generar_informe_unit(cuv_input)

    if informe_generado:
        st.success("Informe generado correctamente.")
        nombre_archivo = f"informe_ventilacion_{cuv_input}.docx"
        st.download_button(
            label="Descargar informe",
            data=informe_generado,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


if __name__ == "__main__":
    main()
