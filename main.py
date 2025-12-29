import logging
import streamlit as st
import pandas as pd
from datetime import date, datetime, time as dt_time, timedelta
import time
import os
import io
import uuid
from decimal import Decimal
import math
from pathlib import Path
import shutil
from dotenv import load_dotenv
from streamlit_cookies_controller import CookieController
from informe_ventilacion import generar_descarga_informe
from utils.helpers import (
    autenticar_usuario,
    get_ct,
    precompletar_campos_ct,
    interpreter_pmv,
    get_cuv,
    guardar_visita_inicio,
    guardar_visita_cierre,
    logout,
    insertar_medicion,
    get_met,
    check_resultado_pmv,
    get_areas_options,
    get_motivo_eval,
    get_equipo_vel,
    get_equipo_temp,
    get_sector_especifico,
    get_puesto_trabajo,
    get_posicion_trabajador,
    get_vestimenta_trabajador,
    comparar_patron,
    get_visitas_por_cuv,
    get_visita,
    get_mediciones,
    actualizar_visita_inicio,
    actualizar_medicion,
    get_equipo_dicc_por_id,
    insertar_area_ventilacion,
    generar_siguiente_area_id,
    obtener_areas_ventilacion_por_visita,
    obtener_puntos_ventilacion_por_area,
    obtener_puntos_ventilacion_por_visita,
    insertar_punto_ventilacion,
    recalcular_totales_area_ventilacion,
    guardar_equipos_ventilacion,
)

try:
    from utils.helpers import normalizar_fecha_mysql, normalizar_hora_mysql
except ImportError:
    logging.warning(
        "No se pudieron importar las utilidades de normalización desde utils.helpers; "
        "se utilizarán implementaciones locales de respaldo."
    )


    def normalizar_fecha_mysql(valor):
        if valor is None or valor == "":
            return None

        if isinstance(valor, date) and not isinstance(valor, datetime):
            return valor

        if isinstance(valor, datetime):
            return valor.date()

        if isinstance(valor, str):
            for formato in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
                try:
                    return datetime.strptime(valor, formato).date()
                except ValueError:
                    continue
            try:
                return date.fromisoformat(valor)
            except ValueError:
                return valor

        return valor


    def normalizar_hora_mysql(valor):
        if valor is None or valor == "":
            return None

        if isinstance(valor, dt_time):
            return valor

        if isinstance(valor, datetime):
            return valor.time()

        if isinstance(valor, timedelta):
            base_datetime = datetime.combine(date.today(), dt_time.min) + valor
            return base_datetime.time()

        if isinstance(valor, str):
            for formato in ("%H:%M:%S", "%H:%M"):
                try:
                    return datetime.strptime(valor, formato).time()
                except ValueError:
                    continue
            try:
                return dt_time.fromisoformat(valor)
            except ValueError:
                return valor

        return valor

from pythermalcomfort.models import pmv_ppd_iso
from utils.informe import generar_informe

st.set_page_config(page_title="Informes Confort Térmico", layout="wide")


def ensure_float(value, default=None):
    """Garantiza que los valores numéricos sean ``float`` compatibles con Streamlit."""
    if value is None:
        return default
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        value = value.strip()
        if value == "":
            return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _descomponer_fecha_hora(valor):
    """Devuelve una tupla (fecha, hora) a partir de un valor mixto."""

    fecha_default = date.today()
    hora_default = datetime.now().time().replace(microsecond=0)

    if valor is None or valor == "":
        return fecha_default, hora_default

    if isinstance(valor, datetime):
        return valor.date(), valor.time()

    if isinstance(valor, date):
        return valor, hora_default

    if isinstance(valor, str):
        for formato in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                fecha_parseada = datetime.strptime(valor, formato)
                return fecha_parseada.date(), fecha_parseada.time()
            except ValueError:
                continue

    return fecha_default, hora_default


AREA_IMAGES_DIR = Path("imagenes_pdf") / "areas"
AREA_ALLOWED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}
AREA_FORM_WIDGET_KEYS = {
    "nombre_area": "vent_area_nombre",
    "largo_m": "vent_area_largo",
    "ancho_m": "vent_area_ancho",
    "alto_m": "vent_area_alto",
    "aforo_permitido": "vent_area_aforo",
    "observaciones": "vent_area_observaciones",
}
VENT_AREA_FORM_PENDING_KEY = "vent_area_form_pending_updates"
VENT_PUNTO_FORM_WIDGET_KEYS = {
    "punto_id": "vent_punto_id",
    "codigo_punto": "vent_punto_codigo",
    "tipo_punto": "vent_punto_tipo",
    "ubicacion_detalle": "vent_punto_ubicacion",
    "altura_m": "vent_punto_altura",
    "distancia_fuente_m": "vent_punto_distancia",
    "conducto_largo_cm": "vent_punto_conducto_largo",
    "conducto_ancho_cm": "vent_punto_conducto_ancho",
    "conducto_diametro": "vent_punto_conducto_diametro",
    "medicion_caudal_1": "vent_punto_vel_1",
    "medicion_caudal_2": "vent_punto_vel_2",
    "medicion_caudal_3": "vent_punto_vel_3",
    "medicion_caudal_4": "vent_punto_vel_4",
    "medicion_caudal_5": "vent_punto_vel_5",
    "medicion_caudal_p": "vent_punto_vel_prom",
    "condiciones_ocupacion": "vent_punto_ocupacion",
    "puertas_ventanas_abiertas": "vent_punto_puertas",
    "temperatura_c": "vent_punto_temp",
    "humedad_relativa_pct": "vent_punto_humedad",
    "croquis_url": "vent_punto_croquis",
    "observaciones": "vent_punto_observaciones",
}


def _obtener_directorio_fotografias_area(visita_id, area_id):
    """Devuelve la ruta base para almacenar fotografías de un área específica."""

    return AREA_IMAGES_DIR / str(visita_id) / str(area_id)


def _migrar_fotografias_area_legacy(visita_id, area_id):
    """Traslada fotografías antiguas almacenadas sin la carpeta de visita asociada."""

    if not visita_id or not area_id:
        return

    origen = AREA_IMAGES_DIR / str(area_id)
    destino = _obtener_directorio_fotografias_area(visita_id, area_id)

    if origen == destino or not origen.exists() or not origen.is_dir():
        return

    destino.mkdir(parents=True, exist_ok=True)

    for ruta in origen.iterdir():
        if not ruta.is_file() or ruta.suffix.lower() not in AREA_ALLOWED_EXTENSIONS:
            continue

        destino_final = destino / ruta.name
        if destino_final.exists():
            try:
                ruta.unlink(missing_ok=True)
            except Exception as error:
                logging.warning("No se pudo eliminar la fotografía duplicada %s: %s", ruta, error)
            continue

        try:
            shutil.move(str(ruta), str(destino_final))
        except Exception as error:
            logging.warning("No se pudo migrar la fotografía %s: %s", ruta, error)

    try:
        origen.rmdir()
    except OSError:
        pass


def _resetear_estado_form_area():
    """Resetea el estado del formulario de áreas sin activar el modo create."""
    
    st.session_state.pop("vent_area_form_mode", None)
    st.session_state.pop("vent_area_form_area_id", None)
    st.session_state.pop("vent_area_fotos_uploader", None)
    st.session_state.pop("vent_area_status", None)
    st.session_state.pop(VENT_AREA_FORM_PENDING_KEY, None)


def _limpiar_estado_form_area():
    """Inicializa o limpia los valores del formulario de áreas de ventilación."""

    st.session_state["vent_area_form_mode"] = "create"
    st.session_state["vent_area_form_area_id"] = None
    st.session_state.pop("vent_area_fotos_uploader", None)
    st.session_state.pop("vent_area_status", None)

    valores_por_defecto = {
        AREA_FORM_WIDGET_KEYS["nombre_area"]: "",
        AREA_FORM_WIDGET_KEYS["largo_m"]: 0.0,
        AREA_FORM_WIDGET_KEYS["ancho_m"]: 0.0,
        AREA_FORM_WIDGET_KEYS["alto_m"]: 0.0,
        AREA_FORM_WIDGET_KEYS["aforo_permitido"]: 0,
        AREA_FORM_WIDGET_KEYS["observaciones"]: "",
    }
    st.session_state[VENT_AREA_FORM_PENDING_KEY] = valores_por_defecto


def _cargar_area_en_formulario(area_data):
    """Carga los datos de un área existente en el formulario de edición."""

    if not area_data:
        _limpiar_estado_form_area()
        return

    st.session_state["vent_area_form_mode"] = "edit"
    st.session_state["vent_area_form_area_id"] = area_data.get("area_id")
    st.session_state.pop("vent_area_fotos_uploader", None)

    aforo_valor = area_data.get("aforo_permitido")
    try:
        aforo_normalizado = int(aforo_valor or 0)
    except (TypeError, ValueError):
        aforo_normalizado = 0

    st.session_state[VENT_AREA_FORM_PENDING_KEY] = {
        AREA_FORM_WIDGET_KEYS["nombre_area"]: area_data.get("nombre_area", ""),
        AREA_FORM_WIDGET_KEYS["largo_m"]: ensure_float(area_data.get("largo_m"), 0.0) or 0.0,
        AREA_FORM_WIDGET_KEYS["ancho_m"]: ensure_float(area_data.get("ancho_m"), 0.0) or 0.0,
        AREA_FORM_WIDGET_KEYS["alto_m"]: ensure_float(area_data.get("alto_m"), 0.0) or 0.0,
        AREA_FORM_WIDGET_KEYS["aforo_permitido"]: aforo_normalizado,
        AREA_FORM_WIDGET_KEYS["observaciones"]: area_data.get("observaciones", "") or "",
    }


def _limpiar_estado_form_punto(mantener_area=None):
    """Reinicia los valores del formulario de puntos de medición."""

    if mantener_area is None:
        st.session_state.pop("vent_punto_form_area", None)
    else:
        st.session_state["vent_punto_form_area"] = mantener_area

    st.session_state.pop("vent_punto_en_edicion", None)
    st.session_state.pop("vent_punto_form_last_id", None)
    st.session_state["vent_punto_form_mode"] = "create"

    for key in VENT_PUNTO_FORM_WIDGET_KEYS.values():
        st.session_state.pop(key, None)

    st.session_state.pop("vent_punto_fecha", None)
    st.session_state.pop("vent_punto_hora", None)


def _cargar_punto_en_formulario(punto_data):
    """Carga los datos de un punto existente en el formulario de edición."""

    if not punto_data:
        _limpiar_estado_form_punto()
        return

    st.session_state["vent_punto_form_mode"] = "edit"
    st.session_state["vent_punto_en_edicion"] = punto_data
    st.session_state["vent_punto_form_area"] = punto_data.get("area_id")
    st.session_state["vent_punto_form_last_id"] = punto_data.get("punto_id")

    fecha_prefill, hora_prefill = _descomponer_fecha_hora(punto_data.get("fecha_hora"))

    tipo_valor = (punto_data.get("tipo_punto") or "").strip().lower()
    if tipo_valor.startswith("extra"):
        tipo_normalizado = "Extraccion"
    elif tipo_valor.startswith("inye") or tipo_valor.startswith("inyec"):
        tipo_normalizado = "Inyeccion"
    else:
        tipo_normalizado = punto_data.get("tipo_punto", "Inyeccion") or "Inyeccion"

    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["punto_id"]] = punto_data.get("punto_id", "")
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["codigo_punto"]] = punto_data.get("codigo_punto", "")
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["tipo_punto"]] = tipo_normalizado
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["ubicacion_detalle"]] = punto_data.get("ubicacion_detalle", "")
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["altura_m"]] = ensure_float(punto_data.get("altura_m"), 0.0) or 0.0
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["distancia_fuente_m"]] = ensure_float(
        punto_data.get("distancia_fuente_m"), 0.0
    ) or 0.0
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["conducto_largo_cm"]] = ensure_float(
        punto_data.get("conducto_largo_cm"), 0.0
    ) or 0.0
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["conducto_ancho_cm"]] = ensure_float(
        punto_data.get("conducto_ancho_cm"), 0.0
    ) or 0.0
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["conducto_diametro"]] = ensure_float(
        punto_data.get("conducto_diametro"), 0.0
    ) or 0.0
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["medicion_caudal_1"]] = ensure_float(
        punto_data.get("medicion_caudal_1"), 0.0
    ) or 0.0
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["medicion_caudal_2"]] = ensure_float(
        punto_data.get("medicion_caudal_2"), 0.0
    ) or 0.0
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["medicion_caudal_3"]] = ensure_float(
        punto_data.get("medicion_caudal_3"), 0.0
    ) or 0.0
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["medicion_caudal_4"]] = ensure_float(
        punto_data.get("medicion_caudal_4"), 0.0
    ) or 0.0
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["medicion_caudal_5"]] = ensure_float(
        punto_data.get("medicion_caudal_5"), 0.0
    ) or 0.0
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["medicion_caudal_p"]] = ensure_float(
        punto_data.get("medicion_caudal_p"), 0.0
    ) or 0.0

    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["condiciones_ocupacion"]] = int(
        punto_data.get("condiciones_ocupacion") or 0
    )
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["puertas_ventanas_abiertas"]] = (
        "Sí" if punto_data.get("puertas_ventanas_abiertas") in (1, "1", True, "True") else "No"
    )
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["temperatura_c"]] = ensure_float(
        punto_data.get("temperatura_c"), 20.0
    ) or 20.0
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["humedad_relativa_pct"]] = ensure_float(
        punto_data.get("humedad_relativa_pct"), 50.0
    ) or 50.0
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["croquis_url"]] = punto_data.get("croquis_url", "")
    st.session_state[VENT_PUNTO_FORM_WIDGET_KEYS["observaciones"]] = punto_data.get("observaciones", "")

    st.session_state["vent_punto_fecha"] = fecha_prefill
    st.session_state["vent_punto_hora"] = hora_prefill


def _manejar_cambio_punto(opciones_puntos, area_actual):
    """Callback para sincronizar el formulario al cambiar el punto seleccionado."""

    etiqueta = st.session_state.get("vent_punto_selector")
    punto = opciones_puntos.get(etiqueta)
    if punto:
        _cargar_punto_en_formulario(punto)
    else:
        _limpiar_estado_form_punto(area_actual)


def _aplicar_pendientes_form_area():
    """Sincroniza los valores programados del formulario con ``st.session_state``."""

    valores_pendientes = st.session_state.pop(VENT_AREA_FORM_PENDING_KEY, None)
    if not valores_pendientes:
        return

    for key, value in valores_pendientes.items():
        st.session_state[key] = value


def _guardar_fotografias_area(visita_id, area_id, archivos_subidos):
    """Persiste en disco las fotografías asociadas a un área."""

    if not archivos_subidos:
        return []

    _migrar_fotografias_area_legacy(visita_id, area_id)

    destino_base = _obtener_directorio_fotografias_area(visita_id, area_id)
    destino_base.mkdir(parents=True, exist_ok=True)

    fotografias_guardadas = []
    for archivo in archivos_subidos:
        if not archivo:
            continue
        extension = Path(archivo.name).suffix.lower() or ".png"
        if extension not in AREA_ALLOWED_EXTENSIONS:
            extension = ".png"
        nombre_archivo = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:8]}{extension}"
        ruta_archivo = destino_base / nombre_archivo
        with ruta_archivo.open("wb") as destino:
            destino.write(archivo.getbuffer())
        fotografias_guardadas.append(ruta_archivo)

    return fotografias_guardadas


def _listar_fotografias_area(visita_id, area_id):
    """Devuelve las rutas de fotografías registradas para un área."""

    if not visita_id or not area_id:
        return []

    _migrar_fotografias_area_legacy(visita_id, area_id)

    destinos = [
        _obtener_directorio_fotografias_area(visita_id, area_id),
        AREA_IMAGES_DIR / str(area_id),
    ]

    fotografias = []
    for destino in destinos:
        if not destino.exists():
            continue
        fotografias.extend(
            ruta
            for ruta in destino.iterdir()
            if ruta.is_file() and ruta.suffix.lower() in AREA_ALLOWED_EXTENSIONS
        )

    fotografias_unicas = {ruta: None for ruta in fotografias}
    return sorted(fotografias_unicas.keys())


def _eliminar_directorio_vacio_area(directorio):
    """Elimina directorios vacíos dentro de ``AREA_IMAGES_DIR``."""

    if not directorio:
        return

    directorio = Path(directorio)
    limite = AREA_IMAGES_DIR.resolve()

    while True:
        try:
            if not directorio.exists() or not directorio.is_dir():
                break
            if any(directorio.iterdir()):
                break
            if directorio.resolve() == limite:
                break
            directorio.rmdir()
            directorio = directorio.parent
        except Exception as error:
            logging.warning(
                "No se pudo eliminar el directorio vacío de fotografías %s: %s",
                directorio,
                error,
            )
            break


def _eliminar_fotografia_area(ruta_fotografia):
    """Elimina una fotografía registrada de un área."""

    if not ruta_fotografia:
        return False

    ruta = Path(ruta_fotografia)

    try:
        if ruta.exists() and ruta.is_file():
            ruta.unlink()
            _eliminar_directorio_vacio_area(ruta.parent)
            return True
        return False
    except Exception as error:
        logging.error("No se pudo eliminar la fotografía %s: %s", ruta, error)
        return False


def _mostrar_mensaje_area():
    """Muestra mensajes informativos relacionados con la gestión de áreas."""

    estado = st.session_state.pop("vent_area_status", None)
    if not estado:
        return

    nivel, mensaje = estado
    if nivel == "success":
        st.success(mensaje)
    elif nivel == "warning":
        st.warning(mensaje)
    else:
        st.info(mensaje)


def _registrar_mensaje_area(nivel, mensaje):
    """Almacena mensajes para mostrarlos tras un ``st.rerun``."""

    st.session_state["vent_area_status"] = (nivel, mensaje)


# Inicializa el controlador de cookies
cookie_controller = CookieController(key="app_cookies")

# Inicializar sesión del usuario
if "data_user" not in st.session_state:
    st.session_state["data_user"] = None

# Cargar variables de entorno desde un archivo .env
load_dotenv()


def reset_visita_context():
    """Limpiar los datos asociados a una visita en el session_state."""
    st.session_state["id_visita"] = None
    st.session_state["mostrar_formularios"] = False
    st.session_state["mostrar_caja_verificacion"] = False
    st.session_state["modo_edicion"] = False
    st.session_state["visita_prefill"] = {}
    st.session_state["cierre_prefill"] = {}
    st.session_state["mediciones_ids"] = {}
    st.session_state["areas_data"] = {}
    st.session_state["vent_areas"] = []
    st.session_state["vent_puntos"] = {}
    _resetear_estado_form_area()
    st.session_state.pop("cierre", None)
    st.session_state.pop("visita_finalizada", None)
    st.session_state.pop("visita_a_cargar", None)
    st.session_state.pop("prefill_ready", None)
    st.session_state.pop("visita_seleccionada", None)
    st.session_state.pop("vent_area_seleccionada", None)
    st.session_state.pop("ven_area_selector", None)
    # st.session_state["visitas_disponibles"] = pd.DataFrame()  <-- REMOVED to persist list on new visit creation

    for i in range(1, 11):
        st.session_state.pop(f"area_sector_{i}", None)
        st.session_state.pop(f"espec_sector_{i}", None)
        st.session_state.pop(f"puesto_trabajo_{i}", None)
        st.session_state.pop(f"pos_trabajador_{i}", None)
        st.session_state.pop(f"vestimenta_{i}", None)
        st.session_state.pop(f"tbs_{i}", None)
        st.session_state.pop(f"tg_{i}", None)
        st.session_state.pop(f"hr_{i}", None)
        st.session_state.pop(f"vel_aire_{i}", None)
        st.session_state.pop(f"techumbre_{i}", None)
        st.session_state.pop(f"obs_techumbre_{i}", None)
        st.session_state.pop(f"paredes_{i}", None)
        st.session_state.pop(f"obs_paredes_{i}", None)
        st.session_state.pop(f"ventanales_{i}", None)
        st.session_state.pop(f"obs_ventanales_{i}", None)
        st.session_state.pop(f"aire_acond_{i}", None)
        st.session_state.pop(f"obs_aire_acond_{i}", None)
        st.session_state.pop(f"ventiladores_{i}", None)
        st.session_state.pop(f"obs_ventiladores_{i}", None)
        st.session_state.pop(f"inyeccion_extrac_{i}", None)
        st.session_state.pop(f"obs_inyeccion_{i}", None)
        st.session_state.pop(f"ventanas_{i}", None)
        st.session_state.pop(f"obs_ventanas_{i}", None)
        st.session_state.pop(f"puertas_{i}", None)
        st.session_state.pop(f"obs_puertas_{i}", None)
        st.session_state.pop(f"otras_{i}", None)
        st.session_state.pop(f"obs_otras_{i}", None)

    st.session_state["cod_equipo_t"] = "Seleccione..."
    st.session_state["cod_equipo_v"] = "Seleccione..."
    st.session_state["vent_equipo_temp"] = "Seleccione..."
    st.session_state["vent_equipo_vel"] = "Seleccione..."


def preparar_nueva_visita():
    reset_visita_context()
    st.session_state["mostrar_formularios"] = True
    st.session_state["status_message"] = "Formulario listo para registrar una nueva visita."
    st.rerun()


def cargar_visita_existente(id_visita):
    visita_df = get_visita(id_visita)
    st.session_state["mostrar_formularios"] = True
    if visita_df.empty:
        st.session_state["status_message"] = "No se encontraron datos para la visita seleccionada."
        st.rerun()
        return

    visita = visita_df.iloc[0].to_dict()

    fecha_visita = normalizar_fecha_mysql(visita.get("fecha_visita"))
    hora_visita = normalizar_hora_mysql(visita.get("hora_visita"))

    tipo_evaluacion = visita.get("tipo_evaluacion", "confort") or "confort"
    es_confort = tipo_evaluacion == "confort"

    equipo_temp = get_equipo_dicc_por_id(visita.get("equipo_temp")) if es_confort else "Seleccione..."
    equipo_vel = get_equipo_dicc_por_id(visita.get("equipo_vel_air")) if es_confort else "Seleccione..."
    if equipo_temp is None:
        equipo_temp = "Seleccione..."
    if equipo_vel is None:
        equipo_vel = "Seleccione..."

    st.session_state["id_visita"] = id_visita
    st.session_state["modo_edicion"] = True
    st.session_state["visita_prefill"] = {
        "fecha_visita": fecha_visita or date.today(),
        "hora_visita": hora_visita or dt_time(hour=9, minute=0),
        "motivo_evaluacion": visita.get("motivo_evaluacion", ""),
        "nombre_personal_visita": visita.get("nombre_personal_visita", ""),
        "cargo_personal_visita": visita.get("cargo_personal_visita", ""),
        "consultor_ist": visita.get("consultor_ist", ""),
        "tipo_evaluacion": tipo_evaluacion,
        "temperatura_dia": ensure_float(visita.get("temperatura_dia"), 25.0) if es_confort else None,
        "equipo_temp": equipo_temp,
        "equipo_vel_air": equipo_vel,
        "patron_tbs": ensure_float(visita.get("patron_tbs"), 46.4) if es_confort else None,
        "ver_tbs_ini": ensure_float(visita.get("ver_tbs_ini")) if es_confort else None,
        "patron_tbh": ensure_float(visita.get("patron_tbh"), 12.7) if es_confort else None,
        "ver_tbh_ini": ensure_float(visita.get("ver_tbh_ini")) if es_confort else None,
        "patron_tg": ensure_float(visita.get("patron_tg"), 69.8) if es_confort else None,
        "ver_tg_ini": ensure_float(visita.get("ver_tg_ini")) if es_confort else None,
    }

    if es_confort:
        st.session_state["cod_equipo_t"] = equipo_temp
        st.session_state["cod_equipo_v"] = equipo_vel
        st.session_state["mostrar_caja_verificacion"] = True

        cierre_prefill = {
            "ver_tbs_fin": ensure_float(visita.get("ver_tbs_fin")),
            "ver_tbh_fin": ensure_float(visita.get("ver_tbh_fin")),
            "ver_tg_fin": ensure_float(visita.get("ver_tg_fin")),
            "note_visita": visita.get("note_visita", ""),
        }
        st.session_state["cierre_prefill"] = cierre_prefill

        if all(value is not None for key, value in cierre_prefill.items() if key != "note_visita"):
            st.session_state["cierre"] = {
                "Verificación TBS final": cierre_prefill["ver_tbs_fin"],
                "Verificación TBH final": cierre_prefill["ver_tbh_fin"],
                "Verificación TG final": cierre_prefill["ver_tg_fin"],
                "Comentarios finales de evaluación": cierre_prefill["note_visita"],
            }
        else:
            st.session_state.pop("cierre", None)
    else:
        equipo_temp_vent = get_equipo_dicc_por_id(visita.get("evv_equipo_temp")) or "Seleccione..."
        equipo_vel_vent = get_equipo_dicc_por_id(visita.get("evv_equipo_vel_air")) or "Seleccione..."

        st.session_state["cod_equipo_t"] = "Seleccione..."
        st.session_state["cod_equipo_v"] = "Seleccione..."
        st.session_state["vent_equipo_temp"] = equipo_temp_vent
        st.session_state["vent_equipo_vel"] = equipo_vel_vent
        st.session_state["visita_prefill"].update(
            {
                "evv_equipo_temp": equipo_temp_vent,
                "evv_equipo_vel_air": equipo_vel_vent,
            }
        )
        st.session_state["mostrar_caja_verificacion"] = False
        comentario_final = visita.get("note_visita")
        st.session_state["cierre_prefill"] = {"note_visita": comentario_final or ""}
        if comentario_final is not None:
            st.session_state["cierre"] = {
                "Comentarios finales de evaluación": comentario_final or "",
            }
        else:
            st.session_state.pop("cierre", None)
        _resetear_estado_form_area()
        st.session_state.pop("vent_area_selector", None)
        st.session_state["vent_areas"] = obtener_areas_ventilacion_por_visita(id_visita)
        puntos_visita = obtener_puntos_ventilacion_por_visita(id_visita)
        puntos_por_area = {}
        for punto in puntos_visita:
            puntos_por_area.setdefault(punto.get("area_id"), []).append(punto)
        st.session_state["vent_puntos"] = puntos_por_area
        if st.session_state["vent_areas"] and "vent_area_seleccionada" not in st.session_state:
            st.session_state["vent_area_seleccionada"] = st.session_state["vent_areas"][0]["area_id"]

    if es_confort:
        mediciones_df = get_mediciones(id_visita)
        st.session_state["mediciones_ids"] = {}
        st.session_state["areas_data"] = {}
        for idx, medicion in enumerate(mediciones_df.to_dict("records")):
            st.session_state["mediciones_ids"][idx] = medicion.get("id_medicion")
            area_normalizada = {
                **medicion,
                "t_bul_seco": ensure_float(medicion.get("t_bul_seco")),
                "t_globo": ensure_float(medicion.get("t_globo")),
                "hum_rel": ensure_float(medicion.get("hum_rel")),
                "vel_air": ensure_float(medicion.get("vel_air")),
            }
            st.session_state["areas_data"][idx] = area_normalizada
            form_idx = idx + 1
            st.session_state[f"area_sector_{form_idx}"] = area_normalizada.get("nombre_area", "Seleccione...")
            st.session_state[f"espec_sector_{form_idx}"] = area_normalizada.get("sector_especifico", "Seleccione...")
            st.session_state[f"puesto_trabajo_{form_idx}"] = area_normalizada.get("puesto_trabajo", "Seleccione...")
            st.session_state[f"pos_trabajador_{form_idx}"] = area_normalizada.get("posicion_trabajador", "Seleccione...")
            st.session_state[f"vestimenta_{form_idx}"] = area_normalizada.get("vestimenta_trabajador", "Seleccione...")
            st.session_state[f"tbs_{form_idx}"] = ensure_float(area_normalizada.get("t_bul_seco"), 0.0)
            st.session_state[f"tg_{form_idx}"] = ensure_float(area_normalizada.get("t_globo"), 0.0)
            st.session_state[f"hr_{form_idx}"] = ensure_float(area_normalizada.get("hum_rel"), 0.0)
            st.session_state[f"vel_aire_{form_idx}"] = ensure_float(area_normalizada.get("vel_air"), 0.0)
            st.session_state[f"techumbre_{form_idx}"] = "Sí" if medicion.get("cond_techumbre") else "No"
            st.session_state[f"obs_techumbre_{form_idx}"] = medicion.get("obs_techumbre", "")
            st.session_state[f"paredes_{form_idx}"] = "Sí" if medicion.get("cond_paredes") else "No"
            st.session_state[f"obs_paredes_{form_idx}"] = medicion.get("obs_paredes", "")
            st.session_state[f"ventanales_{form_idx}"] = "Sí" if medicion.get("cond_vantanal") else "No"
            st.session_state[f"obs_ventanales_{form_idx}"] = medicion.get("obs_ventanal", "")
            st.session_state[f"aire_acond_{form_idx}"] = "Sí" if medicion.get("cond_aire_acond") else "No"
            st.session_state[f"obs_aire_acond_{form_idx}"] = medicion.get("obs_aire_acond", "")
            st.session_state[f"ventiladores_{form_idx}"] = "Sí" if medicion.get("cond_ventiladores") else "No"
            st.session_state[f"obs_ventiladores_{form_idx}"] = medicion.get("obs_ventiladores", "")
            st.session_state[f"inyeccion_extrac_{form_idx}"] = "Sí" if medicion.get("cond_inyeccion_extraccion") else "No"
            st.session_state[f"obs_inyeccion_{form_idx}"] = medicion.get("obs_inyeccion_extraccion", "")
            st.session_state[f"ventanas_{form_idx}"] = "Sí" if medicion.get("cond_ventanas") else "No"
            st.session_state[f"obs_ventanas_{form_idx}"] = medicion.get("obs_ventanas", "")
            st.session_state[f"puertas_{form_idx}"] = "Sí" if medicion.get("cond_puertas") else "No"
            st.session_state[f"obs_puertas_{form_idx}"] = medicion.get("obs_puertas", "")
            st.session_state[f"otras_{form_idx}"] = "Sí" if medicion.get("cond_otras") else "No"
            st.session_state[f"obs_otras_{form_idx}"] = medicion.get("obs_otras", "")
    else:
        st.session_state["mediciones_ids"] = {}
        st.session_state["areas_data"] = {}

    st.session_state["status_message"] = f"Visita {id_visita} cargada para edición."
    st.session_state["prefill_ready"] = True
    st.rerun()


def checkear_session():
    user_data = cookie_controller.get("user_data")
    if user_data:
        st.session_state["data_user"] = user_data
        return True
    else:
        st.session_state["data_user"] = None
        return False


# Lógica de login persistente
def show_login():
    st.title("Login")
    username = st.text_input("Usuario:", key="username_input")
    password = st.text_input("Contraseña:", type="password", key="password_input")
    if st.button("Iniciar sesión"):
        user_data = autenticar_usuario(username, password)
        if user_data:
            cookie_controller.set("user_data", user_data)
            st.session_state["data_user"] = user_data
            st.success("Login exitoso")
            time.sleep(1)
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos.")


def mostrar_equipos_ventilacion():
    st.subheader("Equipos utilizados en ventilación")

    id_visita = st.session_state.get("id_visita")
    if not id_visita:
        st.info("Guarda primero los datos generales de la visita para asociar equipos.")
        return

    opciones_temp = ["Seleccione..."] + get_equipo_temp()
    opciones_vel = ["Seleccione..."] + get_equipo_vel()

    equipo_temp = st.session_state.get("vent_equipo_temp", "Seleccione...")
    equipo_vel = st.session_state.get("vent_equipo_vel", "Seleccione...")

    if equipo_temp not in opciones_temp:
        equipo_temp = "Seleccione..."
    if equipo_vel not in opciones_vel:
        equipo_vel = "Seleccione..."

    idx_temp = opciones_temp.index(equipo_temp)
    idx_vel = opciones_vel.index(equipo_vel)

    with st.form("form_equipos_ventilacion"):
        equipo_temp = st.selectbox(
            "Equipo termohigrómetro (opcional)",
            options=opciones_temp,
            index=idx_temp,
        )
        equipo_vel = st.selectbox(
            "Equipo de velocidad de aire",
            options=opciones_vel,
            index=idx_vel,
        )

        submit_equipos = st.form_submit_button(
            label="Guardar equipos de ventilación",
            type="primary",
            width='stretch',
            icon=":material/check_circle:",
        )

    if submit_equipos:
        equipo_temp = equipo_temp or "Seleccione..."
        equipo_vel = equipo_vel or "Seleccione..."

        datos_vent = {
            "equipo_temp": None if equipo_temp == "Seleccione..." else equipo_temp,
            "equipo_vel_air": None if equipo_vel == "Seleccione..." else equipo_vel,
        }

        if all(valor is None for valor in datos_vent.values()):
            st.error("Selecciona al menos un equipo para guardar la información de ventilación.")
            return

        guardado = guardar_equipos_ventilacion(id_visita, datos_vent)
        if guardado:
            st.session_state["vent_equipo_temp"] = datos_vent.get("equipo_temp") or "Seleccione..."
            st.session_state["vent_equipo_vel"] = datos_vent.get("equipo_vel_air") or "Seleccione..."
            st.success("Equipos de ventilación guardados correctamente.")
            st.rerun()
        else:
            st.error("No se pudieron guardar los equipos de ventilación. Intenta nuevamente.")


def mostrar_formularios_ventilacion():
    st.subheader("Áreas de ventilación")

    id_visita = st.session_state.get("id_visita")
    if not id_visita:
        st.warning("Debes guardar primero los datos generales de la visita antes de registrar áreas de ventilación.")
        return

    centro_raw = st.session_state.get("input_cuv_str")
    try:
        centro_id = int(centro_raw) if centro_raw not in (None, "") else None
    except ValueError:
        centro_id = None

    areas_guardadas = st.session_state.get("vent_areas", [])

    # No inicializar automáticamente el modo de formulario
    # El formulario solo se mostrará cuando el usuario presione "Crear nueva área" o "Cargar área seleccionada"

    _aplicar_pendientes_form_area()

    _mostrar_mensaje_area()

    areas_por_id = {}
    area_label_map = {}
    area_id_list = []
    nombre_repetidos = {}

    for area in areas_guardadas:
        area_id = area.get("area_id")
        if area_id is None:
            continue
        areas_por_id[area_id] = area
        nombre_area_registrado = area.get("nombre_area", "Área") or "Área"
        contador = nombre_repetidos.get(nombre_area_registrado, 0) + 1
        nombre_repetidos[nombre_area_registrado] = contador
        etiqueta = (
            nombre_area_registrado
            if contador == 1
            else f"{nombre_area_registrado} ({contador})"
        )
        area_label_map[area_id] = etiqueta
        area_id_list.append(area_id)

    selected_area_id = st.session_state.get("vent_area_selector")
    with st.container(border=True):
        st.markdown("#### Selección de áreas registradas")
        #
        #
        #
        if areas_guardadas:
            df_areas = pd.DataFrame(areas_guardadas)
            columnas = [
                "nombre_area",
                "volumen_m3",
                "aforo_permitido",
                "m3_porpersona",
                "m3_porpersona_cumple"
            ]
            columnas_disponibles = [col for col in columnas if col in df_areas.columns]
            df_vista = df_areas[columnas_disponibles].copy()
            if "m3_porpersona_cumple" in df_vista.columns:
                df_vista["m3_porpersona_cumple"] = df_vista["m3_porpersona_cumple"].map({1: "Cumple", 0: "No cumple"})
            if "m3_porpersona_hora_cumple" in df_vista.columns:
                df_vista["m3_porpersona_hora_cumple"] = df_vista["m3_porpersona_hora_cumple"].map(
                    {1: "Cumple", 0: "No cumple"})
            if "recambio_hora_cumple" in df_vista.columns:
                df_vista["recambio_hora_cumple"] = df_vista["recambio_hora_cumple"].map({1: "Cumple", 0: "No cumple"})
            st.dataframe(df_vista, width='stretch')
        else:
            st.info("Aún no se han registrado áreas para esta visita.")
        #
        #
        #
        if area_id_list:
            selected_area_id = st.selectbox(
                "Selecciona un área para cargarla",
                options=area_id_list,
                format_func=lambda value: area_label_map.get(value, str(value)),
                key="vent_area_selector",
            )
        else:
            st.session_state.pop("vent_area_selector", None)
            selected_area_id = None

        col_sel_1, col_sel_2 = st.columns(2)
        with col_sel_1:
            cargar_disabled = not area_id_list or selected_area_id is None
            if st.button(
                "Cargar área seleccionada",
                width='stretch',
                type="primary",
                disabled=cargar_disabled,
            ):
                area_a_editar = areas_por_id.get(selected_area_id)
                if area_a_editar:
                    _cargar_area_en_formulario(area_a_editar)
                    st.session_state["vent_area_seleccionada"] = selected_area_id
                    _registrar_mensaje_area(
                        "info",
                        f"Área \"{area_a_editar.get('nombre_area', 'Área')}\" lista para edición.",
                    )
                    st.rerun()
        with col_sel_2:
            if st.button(
                "Crear nueva área",
                width='stretch',
                type="secondary",
            ):
                _limpiar_estado_form_area()
                _registrar_mensaje_area(
                    "info",
                    "Formulario listo para registrar una nueva área.",
                )
                st.rerun()

    def _requiere_puntos_medicion(area: dict) -> bool:
        valor = area.get("m3_porpersona_cumple") if isinstance(area, dict) else None
        if isinstance(valor, str):
            valor_limpio = valor.strip().lower()
            if valor_limpio in {"true", "si", "sí"}:
                return False
            if valor_limpio in {"false", "no"}:
                return True
        if valor in (None, ""):
            return True
        try:
            return int(float(valor)) == 0
        except (TypeError, ValueError):
            return True

    REFERENCIA_M3_PERSONA = 10.0
    REFERENCIA_M3_PERSONA_HORA = 20.0
    RECAMBIO_MIN = 6.0
    RECAMBIO_MAX = 60.0

    # Solo mostrar el formulario si se está creando o editando un área
    area_form_mode = st.session_state.get("vent_area_form_mode")
    if area_form_mode in ("create", "edit"):
        allowed_photo_types = [ext.lstrip(".") for ext in sorted(AREA_ALLOWED_EXTENSIONS)]
        with st.form("form_area_ventilacion"):
            nombre_area = st.text_input(
                "Nombre del área o dependencia",
                key=AREA_FORM_WIDGET_KEYS["nombre_area"],
            )

            col_dim_1, col_dim_2, col_dim_3 = st.columns(3)
            with col_dim_1:
                largo_m = st.number_input(
                    "Largo (m)",
                    min_value=0.0,
                    step=0.1,
                    key=AREA_FORM_WIDGET_KEYS["largo_m"],
                )
            with col_dim_2:
                ancho_m = st.number_input(
                    "Ancho (m)",
                    min_value=0.0,
                    step=0.1,
                    key=AREA_FORM_WIDGET_KEYS["ancho_m"],
                )
            with col_dim_3:
                alto_m = st.number_input(
                    "Altura (m)",
                    min_value=0.0,
                    step=0.1,
                    key=AREA_FORM_WIDGET_KEYS["alto_m"],
                )

            aforo_permitido = st.number_input(
                "Aforo máximo permitido (personas)",
                min_value=0,
                step=1,
                key=AREA_FORM_WIDGET_KEYS["aforo_permitido"],
            )
            observaciones = st.text_area(
                "Identificación de aperturas por donde ingresa y/o sale aire, ejemplo celosía, puerta, ventana, etc...",
                height=80,
                key=AREA_FORM_WIDGET_KEYS["observaciones"],
            )
            uploaded_photos = st.file_uploader(
                "Fotografías del área",
                type=allowed_photo_types,
                accept_multiple_files=True,
                key="vent_area_fotos_uploader",
                help="Adjunta imágenes en formato JPG, JPEG, PNG o WEBP.",
            )

            submit_area = st.form_submit_button(
                label="Guardar área",
                type="primary",
                width='stretch',
                icon=":material/save:",
            )

        if submit_area:
            area_id_actual = st.session_state.get("vent_area_form_area_id")
            area_existente = (
                areas_por_id.get(area_id_actual)
                if area_form_mode == "edit" and area_id_actual
                else None
            )

            if not centro_id:
                st.error("No se ha identificado el centro de trabajo. Verifica el CUV seleccionado.")
            elif not nombre_area.strip():
                st.error("El nombre del área es obligatorio.")
            else:
                if area_form_mode != "edit" or not area_id_actual:
                    area_id_actual = generar_siguiente_area_id(id_visita)

                volumen_m3 = None
                if largo_m and ancho_m and alto_m:
                    volumen_m3 = largo_m * ancho_m * alto_m

                m3_porpersona = None
                if volumen_m3 is not None and aforo_permitido > 0:
                    m3_porpersona = volumen_m3 / aforo_permitido

                m3_porpersona_cumple = (
                    1
                    if (
                        m3_porpersona is not None
                        and m3_porpersona >= REFERENCIA_M3_PERSONA
                    )
                    else 0
                )

                area_base = area_existente or {}
                codigo_area_generado = (
                    area_base.get("codigo_area")
                    or nombre_area.strip()[:40]
                    or f"A{area_id_actual}"
                )

                area_data = {
                    "area_id": area_id_actual,
                    "visita_id": id_visita,
                    "centro_id": centro_id,
                    "codigo_area": codigo_area_generado,
                    "nombre_area": nombre_area.strip(),
                    "uso": area_base.get("uso", ""),
                    "piso_nivel": area_base.get("piso_nivel", ""),
                    "largo_m": largo_m,
                    "ancho_m": ancho_m,
                    "alto_m": alto_m,
                    "volumen_m3": volumen_m3,
                    "aforo_permitido": aforo_permitido,
                    "m3_porpersona": m3_porpersona,
                    "m3_porpersona_594": REFERENCIA_M3_PERSONA,
                    "m3_porpersona_cumple": m3_porpersona_cumple,
                    "caudal_inyeccion_total": ensure_float(area_base.get("caudal_inyeccion_total"), 0.0) or 0.0,
                    "caudal_extraccion_total": ensure_float(area_base.get("caudal_extraccion_total"), 0.0) or 0.0,
                    "m3_porpersona_hora": ensure_float(area_base.get("m3_porpersona_hora")),
                    "m3_porpersona_hora_594": REFERENCIA_M3_PERSONA_HORA,
                    "m3_porpersona_hora_cumple": int(area_base.get("m3_porpersona_hora_cumple", 0) or 0),
                    "recambio_hora_594_min": RECAMBIO_MIN,
                    "recambio_hora_594_max": RECAMBIO_MAX,
                    "recambio_hora": ensure_float(area_base.get("recambio_hora")),
                    "recambio_hora_cumple": int(area_base.get("recambio_hora_cumple", 0) or 0),
                    "ocupacion_habitual": area_base.get("ocupacion_habitual"),
                    "ventilacion_tipo": area_base.get("ventilacion_tipo"),
                    "ventilacion_sistema": area_base.get("ventilacion_sistema", ""),
                    "ventilacion_estado": area_base.get("ventilacion_estado"),
                    "aberturas": area_base.get("aberturas", ""),
                    "croquis_url": area_base.get("croquis_url", ""),
                    "observaciones": observaciones.strip(),
                }

                if insertar_area_ventilacion(area_data):
                    recalcular_totales_area_ventilacion(area_data["area_id"])
                    st.session_state["vent_areas"] = obtener_areas_ventilacion_por_visita(id_visita)
                    puntos_actualizados = {}
                    for area in st.session_state["vent_areas"]:
                        puntos_actualizados[area["area_id"]] = obtener_puntos_ventilacion_por_area(area["area_id"])
                    st.session_state["vent_puntos"] = puntos_actualizados
                    st.session_state["vent_area_seleccionada"] = area_data["area_id"]

                    area_actualizada = next(
                        (
                            area
                            for area in st.session_state["vent_areas"]
                            if area.get("area_id") == area_data["area_id"]
                        ),
                        area_data,
                    )
                    _cargar_area_en_formulario(area_actualizada)

                    fotografias_guardadas = _guardar_fotografias_area(
                        id_visita, area_data["area_id"], uploaded_photos
                    )
                    if fotografias_guardadas:
                        _registrar_mensaje_area(
                            "success",
                            f"Área \"{area_data['nombre_area']}\" guardada y {len(fotografias_guardadas)} fotografía(s) adjunta(s).",
                        )
                    else:
                        _registrar_mensaje_area(
                            "success",
                            f"Área \"{area_data['nombre_area']}\" guardada correctamente.",
                        )
                    st.rerun()
                else:
                    st.error("No fue posible guardar el área. Revisa los datos e inténtalo nuevamente.")

        area_id_en_formulario = st.session_state.get("vent_area_form_area_id")
        if area_id_en_formulario:
            with st.expander("Fotografías registradas del área", expanded=False):
                fotos_area = _listar_fotografias_area(id_visita, area_id_en_formulario)
                if fotos_area:
                    for ruta in fotos_area:
                        col_imagen, col_accion = st.columns([5, 1])
                        with col_imagen:
                            st.image(
                                str(ruta),
                                caption=ruta.name,
                                use_container_width=True,
                            )
                        with col_accion:
                            st.markdown("&nbsp;")
                            if st.button(
                                "Eliminar foto",
                                key=f"vent_area_delete_{area_id_en_formulario}_{ruta.name}",
                            ):
                                if _eliminar_fotografia_area(ruta):
                                    _registrar_mensaje_area(
                                        "success",
                                        f"Fotografía \"{ruta.name}\" eliminada correctamente.",
                                    )
                                else:
                                    _registrar_mensaje_area(
                                        "warning",
                                        f"No se pudo eliminar la fotografía \"{ruta.name}\".",
                                    )
                                st.rerun()
                else:
                    st.info("Aún no se han agregado fotografías para esta área.")



    # Solo mostrar la sección de puntos de medición si hay áreas registradas
    if not areas_guardadas:
        return

    st.markdown("---")
    st.subheader("Puntos de medición")

    area_options = {area_label_map[area_id]: area_id for area_id in area_id_list}
    area_labels = list(area_options.keys())
    if not area_labels:
        st.info("No hay áreas disponibles para registrar puntos de medición.")
        return
    seleccion_actual = st.session_state.get("vent_area_seleccionada")
    if seleccion_actual:
        idx_area = 0
        for posicion, etiqueta in enumerate(area_labels):
            if area_options[etiqueta] == seleccion_actual:
                idx_area = posicion
                break
    else:
        idx_area = 0

    etiqueta_area = st.selectbox("Selecciona el área para registrar el punto", options=area_labels, index=idx_area)
    area_seleccionada = area_options[etiqueta_area]
    st.session_state["vent_area_seleccionada"] = area_seleccionada

    puntos_area = st.session_state.get("vent_puntos", {}).get(area_seleccionada, [])

    if puntos_area:
        df_puntos = pd.DataFrame(puntos_area)
        columnas = [
            "punto_id",
            "codigo_punto",
            "tipo_punto",
            "seccion_conducto_cm2",
            "medicion_caudal_p",
            "caudal",
            "fecha_hora",
            "condiciones_ocupacion",
            "puertas_ventanas_abiertas",
        ]
        columnas_disponibles = [col for col in columnas if col in df_puntos.columns]
        df_vista = df_puntos[columnas_disponibles].copy()
        if "puertas_ventanas_abiertas" in df_vista.columns:
            df_vista["puertas_ventanas_abiertas"] = df_vista["puertas_ventanas_abiertas"].map({1: "Sí", 0: "No"})
        st.dataframe(df_vista, width='stretch')
    elif requiere_puntos:
        st.info("El área seleccionada aún no tiene puntos registrados.")





    if st.session_state.get("vent_punto_form_area") != area_seleccionada:
        _limpiar_estado_form_punto(area_seleccionada)

    puntos_area = st.session_state.get("vent_puntos", {}).get(area_seleccionada, [])

    opciones_puntos = {"Registrar nuevo punto": None}
    for punto in puntos_area:
        etiqueta_punto = f"{punto.get('punto_id', 'Punto')} | {punto.get('codigo_punto', '')}".strip()
        opciones_puntos[etiqueta_punto] = punto

    etiqueta_punto_sel = st.selectbox(
        "Selecciona un punto para cargarlo y editarlo",
        options=list(opciones_puntos.keys()),
        key="vent_punto_selector",
        on_change=_manejar_cambio_punto,
        args=(opciones_puntos, area_seleccionada),
    )

    punto_seleccionado = opciones_puntos.get(etiqueta_punto_sel)

    if punto_seleccionado:
        ultimo_cargado = st.session_state.get("vent_punto_form_last_id")
        if ultimo_cargado != punto_seleccionado.get("punto_id"):
            _cargar_punto_en_formulario(punto_seleccionado)
    else:
        _limpiar_estado_form_punto(area_seleccionada)

    col_punto_sel_1, col_punto_sel_2 = st.columns(2)
    with col_punto_sel_1:
        if st.button("Cargar punto seleccionado", type="primary", disabled=punto_seleccionado is None):
            _cargar_punto_en_formulario(punto_seleccionado)
            st.rerun()
    with col_punto_sel_2:
        if st.button("Limpiar formulario", type="secondary"):
            _limpiar_estado_form_punto(area_seleccionada)
            st.rerun()

    area_en_foco = next((area for area in areas_guardadas if area["area_id"] == area_seleccionada), {})
    requiere_puntos = _requiere_puntos_medicion(area_en_foco)

    submit_punto = None
    if requiere_puntos:
        modo_punto = st.session_state.get("vent_punto_form_mode", "create")
        punto_prefill = st.session_state.get("vent_punto_en_edicion") if st.session_state.get("vent_punto_form_area") == area_seleccionada else None
        if modo_punto == "edit" and punto_prefill:
            ultimo_punto_cargado = st.session_state.get("vent_punto_form_last_id")
            if ultimo_punto_cargado != punto_prefill.get("punto_id"):
                _cargar_punto_en_formulario(punto_prefill)
        fecha_prefill, hora_prefill = _descomponer_fecha_hora(
            punto_prefill.get("fecha_hora") if punto_prefill else None
        )
        fecha_prefill = st.session_state.get("vent_punto_fecha", fecha_prefill)
        hora_prefill = st.session_state.get("vent_punto_hora", hora_prefill)

        st.markdown(
            "#### "
            + (
                "Edición de punto de medición" if modo_punto == "edit" and punto_prefill else "Registrar punto de medición"
            )
        )
        with st.form("form_punto_ventilacion"):
            col1, col2 = st.columns(2)
            with col1:
                punto_id = st.text_input(
                    "Identificador del punto (PuntoId)",
                    value=st.session_state.get(
                        "vent_punto_id", punto_prefill.get("punto_id") if punto_prefill else ""
                    ),
                    key="vent_punto_id",
                )
                codigo_punto = st.text_input(
                    "Código del punto",
                    value=st.session_state.get(
                        "vent_punto_codigo", punto_prefill.get("codigo_punto") if punto_prefill else ""
                    ),
                    key="vent_punto_codigo",
                )
                tipo_opciones = ["Inyeccion", "Extraccion"]
                tipo_por_defecto = st.session_state.get(
                    "vent_punto_tipo", punto_prefill.get("tipo_punto") if punto_prefill else tipo_opciones[0]
                )
                tipo_punto = st.selectbox(
                    "Tipo de punto",
                    options=tipo_opciones,
                    index=tipo_opciones.index(tipo_por_defecto) if tipo_por_defecto in tipo_opciones else 0,
                    key="vent_punto_tipo",
                )
                ubicacion_detalle = st.text_area(
                    "Ubicación y detalles",
                    height=80,
                    value=st.session_state.get(
                        "vent_punto_ubicacion",
                        punto_prefill.get("ubicacion_detalle") if punto_prefill else "",
                    ),
                    key="vent_punto_ubicacion",
                )
                altura_m = st.number_input(
                    "Altura de medición (m)",
                    min_value=0.0,
                    step=0.1,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_altura", punto_prefill.get("altura_m") if punto_prefill else 0.0
                        ),
                        0.0,
                    ),
                    key="vent_punto_altura",
                )
                distancia_fuente_m = st.number_input(
                    "Distancia a la fuente (m)",
                    min_value=0.0,
                    step=0.1,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_distancia",
                            punto_prefill.get("distancia_fuente_m") if punto_prefill else 0.0,
                        ),
                        0.0,
                    ),
                    key="vent_punto_distancia",
                )
                conducto_largo_cm = st.number_input(
                    "Conducto largo (cm)",
                    min_value=0.0,
                    step=0.1,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_conducto_largo",
                            punto_prefill.get("conducto_largo_cm") if punto_prefill else 0.0,
                        ),
                        0.0,
                    ),
                    key="vent_punto_conducto_largo",
                )
                conducto_ancho_cm = st.number_input(
                    "Conducto ancho (cm)",
                    min_value=0.0,
                    step=0.1,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_conducto_ancho",
                            punto_prefill.get("conducto_ancho_cm") if punto_prefill else 0.0,
                        ),
                        0.0,
                    ),
                    key="vent_punto_conducto_ancho",
                )
                conducto_diametro = st.number_input(
                    "Conducto diámetro (cm)",
                    min_value=0.0,
                    step=0.1,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_conducto_diametro",
                            punto_prefill.get("conducto_diametro") if punto_prefill else 0.0,
                        ),
                        0.0,
                    ),
                    key="vent_punto_conducto_diametro",
                )
            with col2:
                medicion_caudal_1 = st.number_input(
                    "Velocidad 1 (m/s)",
                    min_value=0.0,
                    step=0.01,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_vel_1", punto_prefill.get("medicion_caudal_1") if punto_prefill else 0.0
                        ),
                        0.0,
                    ),
                    key="vent_punto_vel_1",
                )
                medicion_caudal_2 = st.number_input(
                    "Velocidad 2 (m/s)",
                    min_value=0.0,
                    step=0.01,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_vel_2", punto_prefill.get("medicion_caudal_2") if punto_prefill else 0.0
                        ),
                        0.0,
                    ),
                    key="vent_punto_vel_2",
                )
                medicion_caudal_3 = st.number_input(
                    "Velocidad 3 (m/s)",
                    min_value=0.0,
                    step=0.01,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_vel_3", punto_prefill.get("medicion_caudal_3") if punto_prefill else 0.0
                        ),
                        0.0,
                    ),
                    key="vent_punto_vel_3",
                )
                medicion_caudal_4 = st.number_input(
                    "Velocidad 4 (m/s)",
                    min_value=0.0,
                    step=0.01,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_vel_4", punto_prefill.get("medicion_caudal_4") if punto_prefill else 0.0
                        ),
                        0.0,
                    ),
                    key="vent_punto_vel_4",
                )
                medicion_caudal_5 = st.number_input(
                    "Velocidad 5 (m/s)",
                    min_value=0.0,
                    step=0.01,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_vel_5", punto_prefill.get("medicion_caudal_5") if punto_prefill else 0.0
                        ),
                        0.0,
                    ),
                    key="vent_punto_vel_5",
                )
                medicion_caudal_p = st.number_input(
                    "Velocidad promedio (m/s)",
                    min_value=0.0,
                    step=0.01,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_vel_prom", punto_prefill.get("medicion_caudal_p") if punto_prefill else 0.0
                        ),
                        0.0,
                    ),
                    key="vent_punto_vel_prom",
                )
                condiciones_ocupacion = st.number_input(
                    "Personas presentes",
                    min_value=0,
                    step=1,
                    value=int(
                        st.session_state.get(
                            "vent_punto_ocupacion", punto_prefill.get("condiciones_ocupacion", 0) if punto_prefill else 0
                        )
                    ),
                    key="vent_punto_ocupacion",
                )
                puertas_opciones = ["No", "Sí"]
                puertas_default = st.session_state.get(
                    "vent_punto_puertas",
                    "Sí"
                    if punto_prefill and punto_prefill.get("puertas_ventanas_abiertas") in (1, "1", True, "True")
                    else "No",
                )
                puertas_abiertas = st.selectbox(
                    "Puertas/ventanas abiertas",
                    options=puertas_opciones,
                    index=puertas_opciones.index(puertas_default),
                    key="vent_punto_puertas",
                )
                temperatura_c = st.number_input(
                    "Temperatura ambiente (°C)",
                    min_value=-20.0,
                    max_value=60.0,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_temp", punto_prefill.get("temperatura_c") if punto_prefill else 20.0
                        ),
                        20.0,
                    ),
                    step=0.1,
                    key="vent_punto_temp",
                )
                humedad_relativa = st.number_input(
                    "Humedad relativa (%)",
                    min_value=0.0,
                    max_value=100.0,
                    value=ensure_float(
                        st.session_state.get(
                            "vent_punto_humedad",
                            punto_prefill.get("humedad_relativa_pct") if punto_prefill else 50.0,
                        ),
                        50.0,
                    ),
                    step=0.1,
                    key="vent_punto_humedad",
                )
                fecha_medicion = st.date_input(
                    "Fecha de medición",
                    value=fecha_prefill,
                    key="vent_punto_fecha",
                )
                hora_medicion = st.time_input(
                    "Hora de medición",
                    value=hora_prefill,
                    key="vent_punto_hora",
                )
                croquis_punto = st.text_input(
                    "URL de apoyo (foto/croquis)",
                    value=st.session_state.get(
                        "vent_punto_croquis", punto_prefill.get("croquis_url") if punto_prefill else ""
                    ),
                    key="vent_punto_croquis",
                )
                observaciones_punto = st.text_area(
                    "Observaciones del punto",
                    height=80,
                    value=st.session_state.get(
                        "vent_punto_observaciones", punto_prefill.get("observaciones") if punto_prefill else ""
                    ),
                    key="vent_punto_observaciones",
                )

            submit_punto = st.form_submit_button(
                label="Guardar punto de medición",
                type="primary",
                width='stretch',
                icon=":material/save:",
            )
    else:
        st.success(
            "El área cumple con la referencia de m³/persona; no es necesario registrar puntos de medición."
        )

    if submit_punto:
        if not punto_id or not codigo_punto:
            st.error("Los campos Identificador y Código del punto son obligatorios.")
        else:
            velocidades = [
                medicion_caudal_1,
                medicion_caudal_2,
                medicion_caudal_3,
                medicion_caudal_4,
                medicion_caudal_5,
            ]
            velocidades_validas = [v for v in velocidades if v > 0]
            velocidad_promedio = medicion_caudal_p if medicion_caudal_p > 0 else (sum(velocidades_validas) / len(velocidades_validas) if velocidades_validas else 0)

            seccion_cm2 = None
            if conducto_diametro > 0:
                seccion_cm2 = math.pi * (conducto_diametro / 2) ** 2
            elif conducto_largo_cm > 0 and conducto_ancho_cm > 0:
                seccion_cm2 = conducto_largo_cm * conducto_ancho_cm

            caudal = None
            if velocidad_promedio > 0 and seccion_cm2:
                area_m2 = seccion_cm2 / 10000
                caudal = velocidad_promedio * area_m2 * 3600

            fecha_hora = None
            if fecha_medicion and hora_medicion:
                fecha_hora = datetime.combine(fecha_medicion, hora_medicion)

            punto_data = {
                "punto_id": punto_id.strip(),
                "evaluacion_id": id_visita,
                "area_id": area_seleccionada,
                "codigo_punto": codigo_punto.strip(),
                "tipo_punto": tipo_punto,
                "ubicacion_detalle": ubicacion_detalle.strip(),
                "altura_m": altura_m or None,
                "distancia_fuente_m": distancia_fuente_m or None,
                "conducto_largo_cm": conducto_largo_cm or None,
                "conducto_ancho_cm": conducto_ancho_cm or None,
                "conducto_diametro": conducto_diametro or None,
                "seccion_conducto_cm2": seccion_cm2,
                "medicion_caudal_1": medicion_caudal_1 or None,
                "medicion_caudal_2": medicion_caudal_2 or None,
                "medicion_caudal_3": medicion_caudal_3 or None,
                "medicion_caudal_4": medicion_caudal_4 or None,
                "medicion_caudal_5": medicion_caudal_5 or None,
                "medicion_caudal_p": velocidad_promedio or None,
                "caudal": caudal,
                "fecha_hora": fecha_hora,
                "condiciones_ocupacion": condiciones_ocupacion,
                "puertas_ventanas_abiertas": 1 if puertas_abiertas == "Sí" else 0,
                "temperatura_c": temperatura_c,
                "humedad_relativa_pct": humedad_relativa,
                "croquis_url": croquis_punto.strip(),
                "observaciones": observaciones_punto.strip(),
            }

            if insertar_punto_ventilacion(punto_data):
                recalcular_totales_area_ventilacion(area_seleccionada)
                st.session_state["vent_puntos"][area_seleccionada] = obtener_puntos_ventilacion_por_area(area_seleccionada)
                st.session_state["vent_areas"] = obtener_areas_ventilacion_por_visita(id_visita)
                st.session_state.pop("vent_punto_en_edicion", None)
                st.session_state["vent_punto_form_mode"] = "create"
                st.success(f"Punto {punto_data['punto_id']} guardado correctamente.")
                st.rerun()
            else:
                st.error("No fue posible guardar el punto de medición. Revisa los datos ingresados.")




    st.markdown("---")
    st.subheader("Cierre")

    comentarios_default = st.session_state.get("cierre_prefill", {}).get("note_visita", "")
    with st.form("form_cierre_ventilacion"):
        comentarios_finales = st.text_area(
            "Comentarios finales de la evaluación",
            value=comentarios_default,
            height=120,
        )
        submit_cierre = st.form_submit_button(
            label="Guardar comentarios finales",
            type="primary",
            width='stretch',
            icon=":material/check_circle:",
        )

    if submit_cierre:
        datos_cierre = {
            "note_visita": comentarios_finales.strip(),
            "ver_tbs_fin": None,
            "ver_tbh_fin": None,
            "ver_tg_fin": None,
        }
        if guardar_visita_cierre(id_visita, datos_cierre, "ventilacion"):
            st.session_state["cierre"] = {"Comentarios finales de evaluación": comentarios_finales.strip()}
            st.session_state["cierre_prefill"] = {"note_visita": comentarios_finales.strip()}
            st.success("Comentarios finales guardados correctamente.")
            st.rerun()
        else:
            st.error("No fue posible guardar el cierre de la visita.")

    st.markdown("---")
    st.subheader("Finalizar visita")

    visita_guardada = id_visita is not None
    areas_registradas = st.session_state.get("vent_areas", [])
    total_areas = len(areas_registradas)
    cierre_completado = "cierre" in st.session_state

    puntos_por_area = st.session_state.get("vent_puntos", {})
    areas_requieren_puntos = [
        area for area in areas_registradas if _requiere_puntos_medicion(area)
    ]
    areas_pendientes = [
        area
        for area in areas_requieren_puntos
        if len(puntos_por_area.get(area["area_id"], [])) == 0
    ]

    if visita_guardada and total_areas > 0 and not areas_pendientes and cierre_completado:
        if "visita_finalizada" not in st.session_state:
            st.session_state["visita_finalizada"] = False

        if not st.session_state["visita_finalizada"]:
            if st.button("Finalizar visita", type="primary"):
                st.session_state["visita_finalizada"] = True
                st.success("Visita finalizada correctamente.")
                st.rerun()

        if st.session_state["visita_finalizada"]:
            st.success("Visita finalizada correctamente. Ya puedes generar el informe.")

            if st.button(
                "Generar informe de ventilación",
                type="primary",
                key="generar_informe_ventilacion",
            ):
                cuv_valor = centro_raw or centro_id
                try:
                    informe_docx = generar_descarga_informe(cuv_valor, id_visita)
                except ValueError as error:
                    st.error(str(error))
                else:
                    st.session_state["ventilacion_informe_docx"] = informe_docx
                    st.success("Informe generado correctamente.")

            informe_buffer = st.session_state.get("ventilacion_informe_docx")
            if informe_buffer:
                nombre_cuv = centro_raw or centro_id or "centro"
                st.download_button(
                    label="Descargar informe de ventilación",
                    data=informe_buffer,
                    file_name=f"informe_ventilacion_{nombre_cuv}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="descargar_informe_ventilacion",
                )
    else:
        st.info("Debes completar todos los pasos antes de finalizar la visita:")
        if not visita_guardada:
            st.warning("La visita aún no ha sido guardada.")
        if total_areas == 0:
            st.warning("Registra al menos un área de ventilación.")
        if areas_pendientes:
            nombres_pendientes = ", ".join(area.get("nombre_area", "Área") for area in areas_pendientes)
            st.warning(
                "Registra al menos un punto de medición para las áreas que lo requieren: "
                f"{nombres_pendientes}."
            )
        if not cierre_completado:
            st.warning("Guarda los comentarios finales de la evaluación.")


def main():
    st.header("Herramienta de registro de mediciones | Confort térmico")
    data_user = st.session_state.get("data_user", {})
    nombre_usuario = data_user.get("nombre", "Usuario")
    email_usuario = data_user.get("email", "Email Usuario")

    with st.container(border=True):
        st.markdown(f'''Bienvenid@ :violet[**{email_usuario}**]''')
        if st.button("Cerrar sesión", type="secondary"):
            logout()
            st.session_state["data_user"] = None
            st.rerun()

    if "cod_equipo_t" not in st.session_state:
        st.session_state["cod_equipo_t"] = "Seleccione..."
    if "cod_equipo_v" not in st.session_state:
        st.session_state["cod_equipo_v"] = "Seleccione..."
    if "vent_equipo_temp" not in st.session_state:
        st.session_state["vent_equipo_temp"] = "Seleccione..."
    if "vent_equipo_vel" not in st.session_state:
        st.session_state["vent_equipo_vel"] = "Seleccione..."
    if "visita_prefill" not in st.session_state:
        st.session_state["visita_prefill"] = {}
    if "cierre_prefill" not in st.session_state:
        st.session_state["cierre_prefill"] = {}
    if "visitas_disponibles" not in st.session_state:
        st.session_state["visitas_disponibles"] = pd.DataFrame()
    if "modo_edicion" not in st.session_state:
        st.session_state["modo_edicion"] = False
    if "mostrar_caja_verificacion" not in st.session_state:
        st.session_state["mostrar_caja_verificacion"] = False

    status_message = st.session_state.pop("status_message", None)
    if status_message:
        st.success(status_message)

    # --- Inicialización en session_state ---
    if "df_filtrado" not in st.session_state:
        st.session_state["df_filtrado"] = pd.DataFrame()
    if "df_info_cuv" not in st.session_state:
        st.session_state["df_info_cuv"] = pd.DataFrame()
    if "mostrar_formularios" not in st.session_state:
        st.session_state["mostrar_formularios"] = False
    if "input_cuv_str" not in st.session_state:
        st.session_state["input_cuv_str"] = ""

    # Búsqueda por CUV
    with st.container(border=True):
        st.info("Ingresa un CUV o CECO y haz clic en 'Buscar' para iniciar el registro de medición.")
        input_cuv = st.text_input("Ingresa el CUV o CECO:")
        if st.button("Buscar"):
            st.session_state["input_cuv_str"] = input_cuv.strip()
            reset_visita_context()
            resultados = get_ct(st.session_state["input_cuv_str"])
            if resultados and len(resultados) > 0:
                st.session_state["df_info_cuv"] = pd.DataFrame([resultados[0]])
                cuv = st.session_state["input_cuv_str"]
                st.session_state["visitas_disponibles"] = get_visitas_por_cuv(cuv)
                st.success("Centro de trabajo encontrado en base de datos.")
            else:
                st.session_state["df_info_cuv"] = pd.DataFrame()
                st.session_state["visitas_disponibles"] = pd.DataFrame()
                st.error("No se encontró el centro de trabajo con el CUV ingresado.")

        df_info_cuv = st.session_state["df_info_cuv"]

    if not df_info_cuv.empty:
        st.markdown("---")

        # 1: Datos generales
        st.subheader("Datos generales")
        if not df_info_cuv.empty:
            cuv_info_row = df_info_cuv.iloc[0]
            campos_ct = precompletar_campos_ct(cuv_info_row)
        else:
            razon_social = st.text_input("Razón Social")
            rut = st.text_input("RUT")
            nombre_local = st.text_input("Nombre de Local")
            direccion = st.text_input("Dirección")
            comuna = st.text_input("Comuna")
            region = st.text_input("Región")
            cuv_val = st.text_input("CUV")

        st.markdown("---")

        visitas_df = st.session_state.get("visitas_disponibles", pd.DataFrame())
        with st.container(border=True):
            st.subheader("Visitas registradas")
            if not visitas_df.empty:
                tipo_labels = {
                    "confort": "Confort térmico",
                    "ventilacion": "Ventilación",
                }
                df_visitas_display = visitas_df.copy()
                if "tipo_evaluacion" in df_visitas_display.columns:
                    df_visitas_display["tipo_evaluacion"] = df_visitas_display["tipo_evaluacion"].map(
                        lambda x: tipo_labels.get(x, x if x else "")
                    )

                columnas_resumen = [
                    col
                    for col in ["id_visita", "fecha_visita", "hora_visita", "motivo_evaluacion", "tipo_evaluacion"]
                    if col in df_visitas_display.columns
                ]
                if columnas_resumen:
                    st.dataframe(df_visitas_display[columnas_resumen], width='stretch')

                opciones_map = {}
                for _, row in df_visitas_display.iterrows():
                    visita_id = row.get("id_visita")
                    if visita_id is None:
                        continue
                    fecha_valor = row.get("fecha_visita")
                    if hasattr(fecha_valor, "strftime"):
                        fecha_str = fecha_valor.strftime("%Y-%m-%d")
                    else:
                        fecha_str = str(fecha_valor) if fecha_valor is not None else "Sin fecha"
                    motivo = row.get("motivo_evaluacion") or "Sin motivo"
                    tipo_label = row.get("tipo_evaluacion") or "Sin tipo"
                    opciones_map[visita_id] = f"{fecha_str} - {tipo_label} - {motivo} (ID {visita_id})"

                if opciones_map:
                    selected_visita_id = st.selectbox(
                        "Selecciona una visita para editarla",
                        options=list(opciones_map.keys()),
                        format_func=lambda x: opciones_map.get(x, str(x)),
                        key="visita_seleccionada"
                    )
                else:
                    selected_visita_id = None

                col1, col2 = st.columns(2)
                with col1:
                    if st.button("Cargar visita seleccionada", width='stretch', type="primary",
                                 disabled=not opciones_map):
                        if selected_visita_id is not None:
                            cargar_visita_existente(selected_visita_id)
                with col2:
                    if st.button("Crear nueva visita", width='stretch', type="secondary"):
                        preparar_nueva_visita()
            else:
                st.info("No existen visitas registradas para este CUV. Puedes crear una nueva visita para comenzar.")
                if st.button("Crear nueva visita", width='stretch'):
                    preparar_nueva_visita()

        st.write("")

        if not st.session_state.get("mostrar_formularios", False):
            st.info(
                "Selecciona una visita y presiona 'Cargar visita seleccionada' o pulsa 'Crear nueva visita' para habilitar el formulario.")
            st.stop()

        st.markdown("---")

        # ============================================
        # CAJA 1: DATOS DE LA VISITA
        # ============================================
        visita_prefill = st.session_state.get("visita_prefill", {})
        id_visita = st.session_state.get("id_visita")

        with st.form("form_datos_visita"):
            st.subheader("Datos de la visita")

            fecha_default = normalizar_fecha_mysql(visita_prefill.get("fecha_visita")) or date.today()
            hora_default = normalizar_hora_mysql(visita_prefill.get("hora_visita")) or dt_time(hour=9, minute=0)

            fecha_visita = st.date_input("Fecha de visita", value=fecha_default)
            hora_medicion = st.time_input("Hora de medición", value=hora_default)

            opc_motivos = get_motivo_eval()
            motivos_options = ["Seleccione..."] + opc_motivos
            motivo_default = visita_prefill.get("motivo_evaluacion")
            motivo_index = motivos_options.index(motivo_default) if motivo_default in opc_motivos else 0
            motivo_evaluacion = st.selectbox("Motivo de evaluación",
                                             options=motivos_options,
                                             index=motivo_index)

            nombre_personal = st.text_input("Nombre del personal SMU",
                                            value=visita_prefill.get("nombre_personal_visita", ""))
            cargo_por_defecto = visita_prefill.get("cargo_personal_visita") or "Administrador/a"
            cargo = st.text_input("Cargo", value=cargo_por_defecto)

            st.markdown("---")
            st.subheader("Tipo de evaluación")
            eval_options = {
                "Confort térmico": "confort",
                "Ventilación": "ventilacion",
            }
            etiquetas_eval = list(eval_options.keys())
            tipo_actual = visita_prefill.get("tipo_evaluacion", "confort")
            etiqueta_actual = next((label for label, val in eval_options.items() if val == tipo_actual),
                                   etiquetas_eval[0])
            try:
                idx_tipo = etiquetas_eval.index(etiqueta_actual)
            except ValueError:
                idx_tipo = 0
            etiqueta_seleccionada = st.selectbox(
                "Selecciona el tipo de evaluación",
                options=etiquetas_eval,
                index=idx_tipo,
            )
            tipo_evaluacion = eval_options[etiqueta_seleccionada]

            submit_datos_visita = st.form_submit_button(
                label="Guardar datos de la visita",
                type="primary",
                width='stretch',
                icon=":material/check_circle:"
            )

        if submit_datos_visita:
            cuv_visita = get_cuv(st.session_state["input_cuv_str"])
            motivo_limpio = motivo_evaluacion if motivo_evaluacion != "Seleccione..." else ""

            visita_base_data = {
                "cuv_visita": cuv_visita,
                "fecha_visita": fecha_visita.strftime("%Y-%m-%d"),
                "hora_visita": hora_medicion.strftime("%H:%M:%S"),
                "motivo_evaluacion": motivo_limpio,
                "nombre_personal_visita": nombre_personal,
                "cargo_personal_visita": cargo,
                "consultor_ist": email_usuario,
                "tipo_evaluacion": tipo_evaluacion,
            }

            nuevo_prefill = {
                "fecha_visita": fecha_visita,
                "hora_visita": hora_medicion,
                "motivo_evaluacion": motivo_limpio,
                "nombre_personal_visita": nombre_personal,
                "cargo_personal_visita": cargo,
                "consultor_ist": email_usuario,
                "tipo_evaluacion": tipo_evaluacion,
            }

            if st.session_state.get("modo_edicion") and st.session_state.get("id_visita"):
                # Actualizar visita existente (sin datos de confort)
                actualizado = actualizar_visita_inicio(
                    st.session_state["id_visita"], visita_base_data, None
                )
                if actualizado:
                    st.session_state["visita_prefill"] = nuevo_prefill
                    st.session_state["visitas_disponibles"] = get_visitas_por_cuv(cuv_visita)

                    # Mostrar caja de verificación solo si es confort
                    if tipo_evaluacion == "confort":
                        st.session_state["mostrar_caja_verificacion"] = True
                    else:
                        st.session_state["mostrar_caja_verificacion"] = False

                    st.session_state["status_message"] = (
                        f"Datos de visita actualizados correctamente. ID: {st.session_state['id_visita']}"
                    )
                    st.session_state["expand_mediciones"] = True
                    st.session_state["expand_cierre"] = True
                    st.rerun()
                else:
                    st.error("Error al actualizar los datos de la visita.")
            else:
                # Crear nueva visita (sin datos de confort aún)
                id_visita = guardar_visita_inicio(visita_base_data, None)

                if id_visita is not None:
                    st.session_state["id_visita"] = id_visita
                    st.session_state["modo_edicion"] = True
                    st.session_state["visita_prefill"] = nuevo_prefill
                    st.session_state["visitas_disponibles"] = get_visitas_por_cuv(cuv_visita)

                    # Mostrar caja de verificación solo si es confort
                    if tipo_evaluacion == "confort":
                        st.session_state["mostrar_caja_verificacion"] = True
                    else:
                        st.session_state["mostrar_caja_verificacion"] = False

                    st.session_state["status_message"] = (
                        f"Datos de visita guardados correctamente. ID de visita: {id_visita}"
                    )
                    st.session_state["expand_mediciones"] = True
                    st.session_state["expand_cierre"] = True
                    st.rerun()
                else:
                    st.error("Error al guardar los datos de la visita.")


        id_visita_actual = st.session_state.get("id_visita")
        if not id_visita_actual:
            st.info("💡 Por favor, guarda los datos generales de la visita para continuar con las etapas de verificación y medición.")
            return

        # ============================================
        # CAJA 2: VERIFICACIÓN DE PARÁMETROS (solo si tipo_evaluacion es 'confort')
        # ============================================
        tipo_actual = st.session_state.get("visita_prefill", {}).get("tipo_evaluacion", "confort")

        if st.session_state.get("mostrar_caja_verificacion", False) and tipo_actual == "confort":
            st.markdown("---")

            with st.form("form_verificacion_parametros"):
                st.subheader("Verificación de parámetros")

                temp_default = ensure_float(visita_prefill.get("temperatura_dia"), 25.0)
                temp_max = st.number_input(
                    "Temperatura máxima del día (°C)",
                    min_value=-50.0,
                    max_value=60.0,
                    value=temp_default,
                    step=0.1
                )

                opc_equipos_temp = get_equipo_temp()
                opc_equipos_vel = get_equipo_vel()
                opciones_temp = ["Seleccione..."] + opc_equipos_temp
                opciones_vel = ["Seleccione..."] + opc_equipos_vel

                cod_equipo_t = st.session_state.get("cod_equipo_t", "Seleccione...")
                cod_equipo_v = st.session_state.get("cod_equipo_v", "Seleccione...")

                if cod_equipo_t not in opciones_temp:
                    cod_equipo_t = "Seleccione..."
                if cod_equipo_v not in opciones_vel:
                    cod_equipo_v = "Seleccione..."

                index_temp = opciones_temp.index(cod_equipo_t)
                index_vel = opciones_vel.index(cod_equipo_v)

                cod_equipo_t = st.selectbox(
                    "Equipo temperatura",
                    options=opciones_temp,
                    index=index_temp,
                )
                cod_equipo_v = st.selectbox(
                    "Equipo velocidad aire",
                    options=opciones_vel,
                    index=index_vel,
                )

                patron_tbs_default = ensure_float(visita_prefill.get("patron_tbs"), 46.4)
                patron_tbh_default = ensure_float(visita_prefill.get("patron_tbh"), 12.7)
                patron_tg_default = ensure_float(visita_prefill.get("patron_tg"), 69.8)

                patron_tbs = st.number_input("Patrón TBS", value=patron_tbs_default, step=0.1)
                patron_tbh = st.number_input(
                    "Patrón TBH (Sólo modificar en caso necesario)",
                    value=patron_tbh_default,
                    step=0.1,
                )
                patron_tg = st.number_input("Patrón TG", value=patron_tg_default, step=0.1)

                st.write()

                ver_tbs_ini_default = ensure_float(visita_prefill.get("ver_tbs_ini"), 0.0)
                ver_tbh_ini_default = ensure_float(visita_prefill.get("ver_tbh_ini"), 0.0)
                ver_tg_ini_default = ensure_float(visita_prefill.get("ver_tg_ini"), 0.0)

                verif_tbs_inicial = st.number_input(
                    "Verificación TBS inicial",
                    value=ver_tbs_ini_default,
                    step=0.1,
                )
                verif_tbh_inicial = st.number_input(
                    "Verificación TBH inicial",
                    value=ver_tbh_ini_default,
                    step=0.1,
                )
                verif_tg_inicial = st.number_input(
                    "Verificación TG inicial",
                    value=ver_tg_ini_default,
                    step=0.1,
                )

                submit_verificacion = st.form_submit_button(
                    label="Guardar verificación de parámetros",
                    type="primary",
                    width='stretch',
                    icon=":material/check_circle:"
                )

            if submit_verificacion:
                errores = False

                if verif_tbs_inicial is None or verif_tbh_inicial is None or verif_tg_inicial is None:
                    st.error(
                        "Los campos de verificación son obligatorios. Por favor completa todos los valores antes de guardar."
                    )
                    errores = True
                elif cod_equipo_t == "Seleccione..." or cod_equipo_v == "Seleccione...":
                    st.error("Debes seleccionar ambos equipos (Temperatura y Velocidad) para validar el patrón.")
                    errores = True
                else:
                    st.session_state["cod_equipo_t"] = cod_equipo_t
                    st.session_state["cod_equipo_v"] = cod_equipo_v

                    data_patron_medicion = (
                        verif_tbs_inicial,
                        verif_tbh_inicial,
                        verif_tg_inicial,
                    )

                    verificacion = comparar_patron(data_patron_medicion, cod_equipo_t)
                    campos_alerta = [campo for campo, estado in verificacion.items() if estado == "alerta"]

                    if "error" in verificacion:
                        st.error(f"Error en la comparación de patrón: {verificacion['error']}")
                        errores = True
                    elif "alerta" in verificacion.values():
                        st.error(
                            "No se ha guardado la verificación | La verificación del patrón detecta una diferencia mayor a 0,5°C en: "
                            + ", ".join(campos_alerta)
                        )
                        st.json(verificacion)
                        errores = True

                if not errores:
                    confort_data = {
                        "temperatura_dia": temp_max,
                        "equipo_temp": cod_equipo_t,
                        "equipo_vel_air": cod_equipo_v,
                        "patron_tbs": patron_tbs,
                        "ver_tbs_ini": verif_tbs_inicial,
                        "patron_tbh": patron_tbh,
                        "ver_tbh_ini": verif_tbh_inicial,
                        "patron_tg": patron_tg,
                        "ver_tg_ini": verif_tg_inicial,
                    }

                    # Actualizar la visita con los datos de confort
                    cuv_visita = get_cuv(st.session_state["input_cuv_str"])
                    visita_base_data = {
                        "cuv_visita": cuv_visita,
                        "fecha_visita": st.session_state["visita_prefill"]["fecha_visita"].strftime("%Y-%m-%d"),
                        "hora_visita": st.session_state["visita_prefill"]["hora_visita"].strftime("%H:%M:%S"),
                        "motivo_evaluacion": st.session_state["visita_prefill"].get("motivo_evaluacion", ""),
                        "nombre_personal_visita": st.session_state["visita_prefill"].get("nombre_personal_visita", ""),
                        "cargo_personal_visita": st.session_state["visita_prefill"].get("cargo_personal_visita", ""),
                        "consultor_ist": email_usuario,
                        "tipo_evaluacion": "confort",
                    }

                    actualizado = actualizar_visita_inicio(
                        st.session_state["id_visita"], visita_base_data, confort_data
                    )

                    if actualizado:
                        # Actualizar el prefill con los nuevos datos
                        st.session_state["visita_prefill"].update({
                            "temperatura_dia": temp_max,
                            "equipo_temp": cod_equipo_t,
                            "equipo_vel_air": cod_equipo_v,
                            "patron_tbs": patron_tbs,
                            "ver_tbs_ini": verif_tbs_inicial,
                            "patron_tbh": patron_tbh,
                            "ver_tbh_ini": verif_tbh_inicial,
                            "patron_tg": patron_tg,
                            "ver_tg_ini": verif_tg_inicial,
                        })

                        st.session_state["status_message"] = (
                            f"Verificación de parámetros guardada correctamente."
                        )
                        st.rerun()
                    else:
                        st.error("Error al guardar la verificación de parámetros.")

        # Resto del código (mediciones, cierre, etc.) continúa igual...
        # [El resto del código original permanece sin cambios desde aquí]


        tipo_actual = st.session_state.get("visita_prefill", {}).get("tipo_evaluacion", "confort")
        if tipo_actual != "confort":
            mostrar_equipos_ventilacion()
            st.markdown("---")
            
            # Verificar si se han guardado los equipos de ventilación
            equipo_t_vent = st.session_state.get("vent_equipo_temp", "Seleccione...")
            equipo_v_vent = st.session_state.get("vent_equipo_vel", "Seleccione...")

            if equipo_v_vent == "Seleccione...":
                 st.info("💡 Por favor, selecciona y guarda el equipo de velocidad de aire para continuar.")
                 return

            mostrar_formularios_ventilacion()
            return

        st.subheader("Mediciones de Áreas")
        st.info("Completa y guarda cada área individualmente")

        # Verificar que el ID de la visita existe antes de guardar mediciones
        id_visita = st.session_state.get("id_visita", None)

        # Inicializar un diccionario en session_state para almacenar los IDs de medición si aún no existe
        if "mediciones_ids" not in st.session_state:
            st.session_state["mediciones_ids"] = {}

        if "areas_data" not in st.session_state:
            st.session_state["areas_data"] = {}

        if id_visita:
            opc_areas_medicion = get_areas_options()
            opc_sector_especifico = get_sector_especifico()
            opc_puesto_trabajo = get_puesto_trabajo()
            opc_posicion_trabajador = get_posicion_trabajador()
            opc_ventimenta_trabajador = get_vestimenta_trabajador()
            for i in range(1, 11):  # Iterar por cada área de medición
                area_idx = i - 1
                default_area = st.session_state["areas_data"].get(area_idx, {})
                expanded_default = st.session_state.get("expand_mediciones", False)
                with st.expander(f"Área {i} - Haz clic para expandir", expanded=expanded_default):
                    with st.form(key=f"form_area_{i}"):
                        # Captura de datos del formulario
                        area_key = f"area_sector_{i}"
                        if area_key not in st.session_state:
                            st.session_state[area_key] = default_area.get("nombre_area", "Seleccione...")
                        nombre_area = st.selectbox(
                            f"Área {i}",
                            options=["Seleccione..."] + opc_areas_medicion,
                            key=area_key
                        )

                        sector_key = f"espec_sector_{i}"
                        if sector_key not in st.session_state:
                            st.session_state[sector_key] = default_area.get("sector_especifico", "Seleccione...")
                        sector_especifico = st.selectbox(
                            f"Sector específico {i}",
                            options=["Seleccione..."] + opc_sector_especifico,
                            key=sector_key
                        )

                        puesto_key = f"puesto_trabajo_{i}"
                        if puesto_key not in st.session_state:
                            st.session_state[puesto_key] = default_area.get("puesto_trabajo", "Seleccione...")
                        puesto_trabajo = st.selectbox(
                            f"Puesto de trabajo {i}",
                            options=["Seleccione..."] + opc_puesto_trabajo,
                            key=puesto_key
                        )

                        posicion_key = f"pos_trabajador_{i}"
                        if posicion_key not in st.session_state:
                            st.session_state[posicion_key] = default_area.get("posicion_trabajador", "Seleccione...")
                        posicion_trabajador = st.selectbox(
                            f"Posición {i}",
                            options=["Seleccione..."] + opc_posicion_trabajador,
                            key=posicion_key
                        )

                        vestimenta_key = f"vestimenta_{i}"
                        if vestimenta_key not in st.session_state:
                            st.session_state[vestimenta_key] = default_area.get("vestimenta_trabajador",
                                                                                "Seleccione...")
                        vestimenta_trabajador = st.selectbox(
                            f"Vestimenta {i}",
                            options=["Seleccione..."] + opc_ventimenta_trabajador,
                            key=vestimenta_key
                        )

                        # Mediciones
                        tbs_key = f"tbs_{i}"
                        if tbs_key not in st.session_state:
                            st.session_state[tbs_key] = default_area.get("t_bul_seco")
                        t_bul_seco = st.number_input(
                            f"Temp. bulbo seco (°C) {i}",
                            value=st.session_state[tbs_key],
                            step=0.1,
                            key=tbs_key
                        )

                        tg_key = f"tg_{i}"
                        if tg_key not in st.session_state:
                            st.session_state[tg_key] = default_area.get("t_globo")
                        t_globo = st.number_input(
                            f"Temp. globo (°C) {i}",
                            value=st.session_state[tg_key],
                            step=0.1,
                            key=tg_key
                        )

                        hr_key = f"hr_{i}"
                        if hr_key not in st.session_state:
                            st.session_state[hr_key] = default_area.get("hum_rel")
                        hum_rel = st.number_input(
                            f"Humedad relativa (%) {i}",
                            value=st.session_state[hr_key],
                            step=0.1,
                            key=hr_key
                        )

                        vel_key = f"vel_aire_{i}"
                        if vel_key not in st.session_state:
                            st.session_state[vel_key] = default_area.get("vel_air")
                        vel_air = st.number_input(
                            f"Velocidad del aire (m/s) {i}",
                            value=st.session_state[vel_key],
                            step=0.1,
                            key=vel_key
                        )

                        # Cálculo de PMV y PPD
                        met = get_met(puesto_trabajo)
                        clo = 0.5 if vestimenta_trabajador == "Habitual" else 1.0

                        # Condiciones y observaciones
                        techumbre_key = f"techumbre_{i}"
                        if techumbre_key not in st.session_state:
                            st.session_state[techumbre_key] = "Sí" if default_area.get("cond_techumbre") else "No"
                        cond_techumbre_label = st.radio(
                            f"Techumbre aislante {i}",
                            ["Sí", "No"],
                            key=techumbre_key
                        )
                        obs_techumbre_key = f"obs_techumbre_{i}"
                        if obs_techumbre_key not in st.session_state:
                            st.session_state[obs_techumbre_key] = default_area.get("obs_techumbre", "")
                        obs_techumbre = st.text_input(
                            f"Obs. Techumbre {i}",
                            key=obs_techumbre_key
                        )
                        cond_techumbre = 1 if cond_techumbre_label == "Sí" else 0

                        paredes_key = f"paredes_{i}"
                        if paredes_key not in st.session_state:
                            st.session_state[paredes_key] = "Sí" if default_area.get("cond_paredes") else "No"
                        cond_paredes_label = st.radio(
                            f"Paredes aislantes {i}",
                            ["Sí", "No"],
                            key=paredes_key
                        )
                        obs_paredes_key = f"obs_paredes_{i}"
                        if obs_paredes_key not in st.session_state:
                            st.session_state[obs_paredes_key] = default_area.get("obs_paredes", "")
                        obs_paredes = st.text_input(
                            f"Obs. Paredes {i}",
                            key=obs_paredes_key
                        )
                        cond_paredes = 1 if cond_paredes_label == "Sí" else 0

                        ventanal_key = f"ventanales_{i}"
                        if ventanal_key not in st.session_state:
                            st.session_state[ventanal_key] = "Sí" if default_area.get("cond_vantanal") else "No"
                        cond_vantanal_label = st.radio(
                            f"Ventanas aislantes {i}",
                            ["Sí", "No"],
                            key=ventanal_key
                        )
                        obs_ventanal_key = f"obs_ventanales_{i}"
                        if obs_ventanal_key not in st.session_state:
                            st.session_state[obs_ventanal_key] = default_area.get("obs_ventanal", "")
                        obs_ventanal = st.text_input(
                            f"Obs. Ventanas {i}",
                            key=obs_ventanal_key
                        )
                        cond_vantanal = 1 if cond_vantanal_label == "Sí" else 0

                        aire_key = f"aire_acond_{i}"
                        if aire_key not in st.session_state:
                            st.session_state[aire_key] = "Sí" if default_area.get("cond_aire_acond") else "No"
                        cond_aire_label = st.radio(
                            f"Aire acondicionado {i}",
                            ["Sí", "No"],
                            key=aire_key
                        )
                        obs_aire_key = f"obs_aire_acond_{i}"
                        if obs_aire_key not in st.session_state:
                            st.session_state[obs_aire_key] = default_area.get("obs_aire_acond", "")
                        obs_aire_acond = st.text_input(
                            f"Obs. Aire Acondicionado {i}",
                            key=obs_aire_key
                        )
                        cond_aire_acond = 1 if cond_aire_label == "Sí" else 0

                        ventiladores_key = f"ventiladores_{i}"
                        if ventiladores_key not in st.session_state:
                            st.session_state[ventiladores_key] = "Sí" if default_area.get(
                                "cond_ventiladores") else "No"
                        cond_ventiladores_label = st.radio(
                            f"Ventiladores {i}",
                            ["Sí", "No"],
                            key=ventiladores_key
                        )
                        obs_ventiladores_key = f"obs_ventiladores_{i}"
                        if obs_ventiladores_key not in st.session_state:
                            st.session_state[obs_ventiladores_key] = default_area.get("obs_ventiladores", "")
                        obs_ventiladores = st.text_input(
                            f"Obs. Ventiladores {i}",
                            key=obs_ventiladores_key
                        )
                        cond_ventiladores = 1 if cond_ventiladores_label == "Sí" else 0

                        inyeccion_key = f"inyeccion_extrac_{i}"
                        if inyeccion_key not in st.session_state:
                            st.session_state[inyeccion_key] = "Sí" if default_area.get(
                                "cond_inyeccion_extraccion") else "No"
                        cond_inyeccion_label = st.radio(
                            f"Inyección/Extracción {i}",
                            ["Sí", "No"],
                            key=inyeccion_key
                        )
                        obs_inyeccion_key = f"obs_inyeccion_{i}"
                        if obs_inyeccion_key not in st.session_state:
                            st.session_state[obs_inyeccion_key] = default_area.get("obs_inyeccion_extraccion", "")
                        obs_inyeccion_extraccion = st.text_input(
                            f"Obs. Inyección {i}",
                            key=obs_inyeccion_key
                        )
                        cond_inyeccion_extraccion = 1 if cond_inyeccion_label == "Sí" else 0

                        ventanas_key = f"ventanas_{i}"
                        if ventanas_key not in st.session_state:
                            st.session_state[ventanas_key] = "Sí" if default_area.get("cond_ventanas") else "No"
                        cond_ventanas_label = st.radio(
                            f"Ventanas abiertas {i}",
                            ["Sí", "No"],
                            key=ventanas_key
                        )
                        obs_ventanas_key = f"obs_ventanas_{i}"
                        if obs_ventanas_key not in st.session_state:
                            st.session_state[obs_ventanas_key] = default_area.get("obs_ventanas", "")
                        obs_ventanas = st.text_input(
                            f"Obs. Ventanas {i}",
                            key=obs_ventanas_key
                        )
                        cond_ventanas = 1 if cond_ventanas_label == "Sí" else 0

                        puertas_key = f"puertas_{i}"
                        if puertas_key not in st.session_state:
                            st.session_state[puertas_key] = "Sí" if default_area.get("cond_puertas") else "No"
                        cond_puertas_label = st.radio(
                            f"Puertas abiertas {i}",
                            ["Sí", "No"],
                            key=puertas_key
                        )
                        obs_puertas_key = f"obs_puertas_{i}"
                        if obs_puertas_key not in st.session_state:
                            st.session_state[obs_puertas_key] = default_area.get("obs_puertas", "")
                        obs_puertas = st.text_input(
                            f"Obs. Puertas {i}",
                            key=obs_puertas_key
                        )
                        cond_puertas = 1 if cond_puertas_label == "Sí" else 0

                        otras_key = f"otras_{i}"
                        if otras_key not in st.session_state:
                            st.session_state[otras_key] = "Sí" if default_area.get("cond_otras") else "No"
                        cond_otras_label = st.radio(
                            f"Otras condiciones {i}",
                            ["Sí", "No"],
                            key=otras_key
                        )
                        obs_otras_key = f"obs_otras_{i}"
                        if obs_otras_key not in st.session_state:
                            st.session_state[obs_otras_key] = default_area.get("obs_otras", "")
                        obs_otras = st.text_input(
                            f"¿Se identifican otras condiciones que pueden considerarse como disconfort térmico? {i}",
                            key=obs_otras_key
                        )
                        cond_otras = 1 if cond_otras_label == "Sí" else 0

                        # Guardar medición
                        if st.form_submit_button(f"Guardar Área {i}"):
                            campos_incompletos = (
                                    nombre_area == "Seleccione..." or
                                    sector_especifico == "Seleccione..." or
                                    puesto_trabajo == "Seleccione..." or
                                    posicion_trabajador == "Seleccione..."
                            )
                            mediciones_incompletas = any(
                                value is None for value in [t_bul_seco, t_globo, hum_rel, vel_air])

                            if campos_incompletos:
                                st.warning(f"Completa todos los campos antes de guardar el Área {i}.")
                            elif mediciones_incompletas:
                                st.warning(f"Debes completar todas las mediciones numéricas para el Área {i}.")
                            else:
                                resultados = pmv_ppd_iso(
                                    tdb=t_bul_seco,
                                    tr=t_globo,
                                    vr=vel_air,
                                    rh=hum_rel,
                                    met=met,
                                    clo=clo,
                                    model="7730-2005",
                                    limit_inputs=False
                                )
                                pmv = float(resultados.pmv)
                                ppd = float(resultados.ppd)
                                resultado_medicion = check_resultado_pmv(pmv)

                                area_data_guardada = {
                                    "nombre_area": nombre_area,
                                    "sector_especifico": sector_especifico,
                                    "puesto_trabajo": puesto_trabajo,
                                    "posicion_trabajador": posicion_trabajador,
                                    "vestimenta_trabajador": vestimenta_trabajador,
                                    "t_bul_seco": t_bul_seco,
                                    "t_globo": t_globo,
                                    "hum_rel": hum_rel,
                                    "vel_air": vel_air,
                                    "ppd": ppd,
                                    "pmv": pmv,
                                    "resultado_medicion": resultado_medicion,
                                    "cond_techumbre": cond_techumbre,
                                    "obs_techumbre": obs_techumbre,
                                    "cond_paredes": cond_paredes,
                                    "obs_paredes": obs_paredes,
                                    "cond_vantanal": cond_vantanal,
                                    "obs_ventanal": obs_ventanal,
                                    "cond_aire_acond": cond_aire_acond,
                                    "obs_aire_acond": obs_aire_acond,
                                    "cond_ventiladores": cond_ventiladores,
                                    "obs_ventiladores": obs_ventiladores,
                                    "cond_inyeccion_extraccion": cond_inyeccion_extraccion,
                                    "obs_inyeccion_extraccion": obs_inyeccion_extraccion,
                                    "cond_ventanas": cond_ventanas,
                                    "obs_ventanas": obs_ventanas,
                                    "cond_puertas": cond_puertas,
                                    "obs_puertas": obs_puertas,
                                    "cond_otras": cond_otras,
                                    "obs_otras": obs_otras,
                                    "met": met,
                                    "clo": clo,
                                }

                                medicion_existente = st.session_state["mediciones_ids"].get(area_idx)
                                if medicion_existente:
                                    actualizado = actualizar_medicion(
                                        medicion_existente,
                                        nombre_area,
                                        sector_especifico,
                                        puesto_trabajo, posicion_trabajador,
                                        vestimenta_trabajador, t_bul_seco,
                                        t_globo, hum_rel,
                                        vel_air, ppd, pmv,
                                        resultado_medicion, cond_techumbre,
                                        obs_techumbre,
                                        cond_paredes, obs_paredes,
                                        cond_vantanal, obs_ventanal,
                                        cond_aire_acond,
                                        obs_aire_acond, cond_ventiladores,
                                        obs_ventiladores,
                                        cond_inyeccion_extraccion,
                                        obs_inyeccion_extraccion, cond_ventanas,
                                        obs_ventanas, cond_puertas, obs_puertas,
                                        cond_otras, obs_otras, met, clo
                                    )

                                    if actualizado:
                                        st.session_state["areas_data"][area_idx] = area_data_guardada
                                        st.session_state["status_message"] = (
                                            f"Área {i} actualizada con éxito (ID medición {medicion_existente})."
                                        )
                                        st.rerun()
                                    else:
                                        st.error(f"No se pudo actualizar la medición para el área {i}.")
                                else:
                                    id_medicion = insertar_medicion(
                                        id_visita, nombre_area,
                                        sector_especifico,
                                        puesto_trabajo, posicion_trabajador,
                                        vestimenta_trabajador, t_bul_seco,
                                        t_globo, hum_rel,
                                        vel_air, ppd, pmv,
                                        resultado_medicion, cond_techumbre,
                                        obs_techumbre,
                                        cond_paredes, obs_paredes,
                                        cond_vantanal, obs_ventanal,
                                        cond_aire_acond,
                                        obs_aire_acond, cond_ventiladores,
                                        obs_ventiladores,
                                        cond_inyeccion_extraccion,
                                        obs_inyeccion_extraccion, cond_ventanas,
                                        obs_ventanas, cond_puertas, obs_puertas,
                                        cond_otras, obs_otras, met, clo
                                    )

                                    if id_medicion:
                                        st.session_state["mediciones_ids"][area_idx] = id_medicion
                                        st.session_state["areas_data"][area_idx] = area_data_guardada
                                        st.session_state["status_message"] = (
                                            f"Área {i} guardada con éxito. ID de la medición: {id_medicion}"
                                        )
                                        st.rerun()
                                    else:
                                        st.error(f"No se pudo guardar la medición para el área {i}.")

            if st.session_state["mediciones_ids"]:
                with st.container(border=True):
                    st.markdown("**Mediciones registradas**")
                    for area_idx, medicion_id in sorted(st.session_state["mediciones_ids"].items()):
                        st.write(f"Área {area_idx + 1}: ID Medición {medicion_id}")
        else:
            st.warning("Debes guardar primero los datos de la visita antes de registrar mediciones.")

        # 4: Cierre
        with st.expander("Cierre", expanded=st.session_state.get("expand_cierre", False)):
            cierre_prefill = st.session_state.get("cierre_prefill", {})

            with st.form("visita_data_cierre"):
                ver_tbs_fin_default = ensure_float(cierre_prefill.get("ver_tbs_fin"), 0.0)
                ver_tbh_fin_default = ensure_float(cierre_prefill.get("ver_tbh_fin"), 0.0)
                ver_tg_fin_default = ensure_float(cierre_prefill.get("ver_tg_fin"), 0.0)

                verif_tbs_final = st.number_input(
                    "Verificación TBS final",
                    value=ver_tbs_fin_default,
                    step=0.1
                )
                verif_tbh_final = st.number_input(
                    "Verificación TBH final",
                    value=ver_tbh_fin_default,
                    step=0.1
                )
                verif_tg_final = st.number_input(
                    "Verificación TG final",
                    value=ver_tg_fin_default,
                    step=0.1
                )
                comentarios_finales = st.text_area(
                    "Comentarios finales de evaluación",
                    value=cierre_prefill.get("note_visita", ""),
                    max_chars=1000
                )

                cierre_submitted = st.form_submit_button(
                    label="Guardar verificación final",
                    type="primary",
                    width='stretch',
                    icon=":material/check_circle:"
                )
                if cierre_submitted:
                    if verif_tbs_final is None or verif_tbh_final is None or verif_tg_final is None:
                        st.error(
                            "Los campos de verificación son obligatorios. Por favor completa todos los valores antes de guardar."
                        )
                    else:
                        equipo = st.session_state["cod_equipo_t"]
                        if not equipo or equipo == "Seleccione...":
                            st.error("No se encontró el equipo de temperatura para comparar el patrón.")
                        else:
                            data_patron_medicion = (
                                verif_tbs_final,
                                verif_tbh_final,
                                verif_tg_final
                            )
                            verificacion = comparar_patron(data_patron_medicion, equipo)
                            campos_alerta = [campo for campo, estado in verificacion.items() if estado == "alerta"]
                            if "error" in verificacion:
                                st.error(f"Error en la comparación de patrón: {verificacion['error']}")
                            elif "alerta" in verificacion.values():
                                st.error(
                                    f"No se ha guardado la verificación | La verificación del patrón detecta una diferencia mayor a 0,5°C. en: {', '.join(campos_alerta)}"
                                )
                                st.json(verificacion)
                            else:
                                visita_cierre_data = (
                                    verif_tbs_final,
                                    verif_tbh_final,
                                    verif_tg_final,
                                    comentarios_finales,
                                )

                                st.session_state["cierre"] = {
                                    "Verificación TBS final": verif_tbs_final,
                                    "Verificación TBH final": verif_tbh_final,
                                    "Verificación TG final": verif_tg_final,
                                    "Comentarios finales de evaluación": comentarios_finales
                                }

                                st.session_state["cierre_prefill"] = {
                                    "ver_tbs_fin": verif_tbs_final,
                                    "ver_tbh_fin": verif_tbh_final,
                                    "ver_tg_fin": verif_tg_final,
                                    "note_visita": comentarios_finales,
                                }

                                id_visita = st.session_state.get("id_visita")
                                if id_visita is not None:
                                    tipo_eval = st.session_state.get("visita_prefill", {}).get("tipo_evaluacion",
                                                                                               "confort")
                                    actualizado = guardar_visita_cierre(id_visita, visita_cierre_data, tipo_eval)
                                    if actualizado:
                                        st.session_state["visita_actualizada"] = True
                                        st.session_state["status_message"] = (
                                            f"Verificación final guardada correctamente para la visita {id_visita}."
                                        )
                                        st.rerun()
                                    else:
                                        st.error("Error al actualizar la visita.")
                                else:
                                    st.error("No se encontró el ID de la visita para actualizar.")
        st.markdown("---")
        # 5: Generación informe
        st.subheader("Finalizar Visita")

        visita_guardada = "id_visita" in st.session_state and st.session_state["id_visita"] is not None
        hay_mediciones = "mediciones_ids" in st.session_state and len(st.session_state["mediciones_ids"]) > 0
        cierre_completado = "cierre" in st.session_state

        if visita_guardada and hay_mediciones and cierre_completado:
            if "visita_finalizada" not in st.session_state:
                st.session_state["visita_finalizada"] = False

            if not st.session_state["visita_finalizada"]:
                if st.button("Finalizar Visita", type="primary"):
                    st.session_state["visita_finalizada"] = True
                    st.success("Visita finalizada correctamente. Ya puedes generar el informe.")
                    st.rerun()

            if st.session_state["visita_finalizada"]:
                st.subheader("Generar Informe")
                st.write("La visita se ha finalizado. ¿Deseas generar el informe basado en la base de datos?")
                if st.button("Sí, generar informe automáticamente", key="generar_informe"):
                    cuv = st.session_state.get("input_cuv_str")
                    id_visita = st.session_state.get("id_visita")
                    informe_docx = generar_informe(cuv, id_visita)
                    if informe_docx:
                        st.session_state["informe_docx"] = informe_docx
                        st.success("Informe generado correctamente.")
                        st.download_button(
                            label="Descargar Informe",
                            data=st.session_state["informe_docx"],
                            file_name=f"informe_{cuv}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="descargar_informe"
                        )
        else:
            st.info("Debes completar todos los pasos antes de finalizar la visita:")
            if not visita_guardada:
                st.warning("La visita aún no ha sido guardada.")
            if not hay_mediciones:
                st.warning("Debes ingresar al menos una medición.")
            if not cierre_completado:
                st.warning("Debes completar y guardar la verificación final (cierre).")


if checkear_session():
    main()
else:
    show_login()