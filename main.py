import logging
import streamlit as st
import pandas as pd
from datetime import date, datetime, time as dt_time, timedelta
import time
import os
import io
from decimal import Decimal
from dotenv import load_dotenv
from streamlit_cookies_controller import CookieController
from utils.helpers import (autenticar_usuario, get_ct, precompletar_campos_ct, interpreter_pmv, get_cuv,
                           guardar_visita_inicio, guardar_visita_cierre, logout, insertar_medicion, get_met,
                           check_resultado_pmv, get_areas_options, get_motivo_eval, get_equipo_vel, get_equipo_temp,
                           get_sector_especifico, get_puesto_trabajo, get_posicion_trabajador,
                           get_vestimenta_trabajador, comparar_patron, get_visitas_por_cuv, get_visita,
                           get_mediciones, actualizar_visita_inicio, actualizar_medicion, get_equipo_dicc_por_id)

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
    st.session_state["modo_edicion"] = False
    st.session_state["visita_prefill"] = {}
    st.session_state["cierre_prefill"] = {}
    st.session_state["mediciones_ids"] = {}
    st.session_state["areas_data"] = {}
    st.session_state.pop("cierre", None)
    st.session_state.pop("visita_finalizada", None)
    st.session_state.pop("visita_a_cargar", None)
    st.session_state.pop("prefill_ready", None)
    st.session_state.pop("visita_seleccionada", None)
    st.session_state["visitas_disponibles"] = pd.DataFrame()

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
        st.session_state["cod_equipo_t"] = "Seleccione..."
        st.session_state["cod_equipo_v"] = "Seleccione..."
        st.session_state["cierre_prefill"] = {"note_visita": visita.get("note_visita", "")}
        st.session_state.pop("cierre", None)

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
    if "visita_prefill" not in st.session_state:
        st.session_state["visita_prefill"] = {}
    if "cierre_prefill" not in st.session_state:
        st.session_state["cierre_prefill"] = {}
    if "visitas_disponibles" not in st.session_state:
        st.session_state["visitas_disponibles"] = pd.DataFrame()
    if "modo_edicion" not in st.session_state:
        st.session_state["modo_edicion"] = False

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
        st.info("Ingresa un CUV y haz clic en 'Buscar' para iniciar el registro de medición.")
        input_cuv = st.text_input("Ingresa el CUV:")
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
                    st.dataframe(df_visitas_display[columnas_resumen], use_container_width=True)

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
                    if st.button("Cargar visita seleccionada", use_container_width=True, type="primary",
                                  disabled=not opciones_map):
                        if selected_visita_id is not None:
                            cargar_visita_existente(selected_visita_id)
                with col2:
                    if st.button("Crear nueva visita", use_container_width=True, type="secondary"):
                        preparar_nueva_visita()
            else:
                st.info("No existen visitas registradas para este CUV. Puedes crear una nueva visita para comenzar.")
                if st.button("Crear nueva visita", use_container_width=True):
                    preparar_nueva_visita()

        st.write("")

        if not st.session_state.get("mostrar_formularios", False):
            st.info(
                "Selecciona una visita y presiona 'Cargar visita seleccionada' o pulsa 'Crear nueva visita' para habilitar el formulario.")
            st.stop()

        st.markdown("---")
        # Formulario 1: Visita - datos visita + calibración inicial
        visita_prefill = st.session_state.get("visita_prefill", {})
        with st.form("visita_data_inicio"):
            # 2: Inicio
            st.subheader("Datos de la visita")

            fecha_default = normalizar_fecha_mysql(visita_prefill.get("fecha_visita")) or date.today()
            hora_default = normalizar_hora_mysql(visita_prefill.get("hora_visita")) or dt_time(hour=9, minute=0)

            temp_default = ensure_float(visita_prefill.get("temperatura_dia"), 25.0)
            fecha_visita = st.date_input("Fecha de visita", value=fecha_default)
            hora_medicion = st.time_input("Hora de medición", value=hora_default)
            temp_max = st.number_input("Temperatura máxima del día (°C)", min_value=-50.0, max_value=60.0,
                                       value=temp_default, step=0.1)

            opc_motivos = get_motivo_eval()
            motivos_options = ["Seleccione..."] + opc_motivos
            motivo_default = visita_prefill.get("motivo_evaluacion")
            motivo_index = motivos_options.index(motivo_default) if motivo_default in opc_motivos else 0
            motivo_evaluacion = st.selectbox("Motivo de evaluación",
                                             options=motivos_options,
                                             index=motivo_index)

            nombre_personal = st.text_input("Nombre del personal SMU",
                                            value=visita_prefill.get("nombre_personal_visita", ""))
            cargo_por_defecto = visita_prefill.get("cargo_personal_visita") or "Administador/a"
            cargo = st.text_input("Cargo", value=cargo_por_defecto)

            st.markdown("---")
            st.subheader("Tipo de evaluación")
            eval_options = {
                "Confort térmico": "confort",
                "Ventilación": "ventilacion",
            }
            etiquetas_eval = list(eval_options.keys())
            tipo_actual = visita_prefill.get("tipo_evaluacion", "confort")
            etiqueta_actual = next((label for label, val in eval_options.items() if val == tipo_actual), etiquetas_eval[0])
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
            es_confort = tipo_evaluacion == "confort"

            cod_equipo_t = st.session_state.get("cod_equipo_t", "Seleccione...")
            cod_equipo_v = st.session_state.get("cod_equipo_v", "Seleccione...")
            patron_tbs = visita_prefill.get("patron_tbs") if es_confort else None
            patron_tbh = visita_prefill.get("patron_tbh") if es_confort else None
            patron_tg = visita_prefill.get("patron_tg") if es_confort else None
            verif_tbs_inicial = visita_prefill.get("ver_tbs_ini") if es_confort else None
            verif_tbh_inicial = visita_prefill.get("ver_tbh_ini") if es_confort else None
            verif_tg_inicial = visita_prefill.get("ver_tg_ini") if es_confort else None

            if es_confort:
                # 3: Calibración
                st.markdown("---")
                st.subheader("Verificación de parámetros")
                opc_equipos_temp = get_equipo_temp()
                opc_equipos_vel = get_equipo_vel()
                opciones_temp = ["Seleccione..."] + opc_equipos_temp
                opciones_vel = ["Seleccione..."] + opc_equipos_vel
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
            else:
                cod_equipo_t = "Seleccione..."
                cod_equipo_v = "Seleccione..."
                st.info(
                    "El formulario detallado para evaluaciones de ventilación estará disponible próximamente."
                )
            submit_visita_inicio = st.form_submit_button(label="Guardar información visita",
                                                        type="primary",
                                                        use_container_width=True,
                                                        icon=":material/check_circle:")

        if submit_visita_inicio:
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

            errores = False
            confort_data = None

        if es_confort:
            if verif_tbs_inicial is None or verif_tbh_inicial is None or verif_tg_inicial is None:
                st.error(
                    "Los campos de verificación son obligatorios. Por favor completa todos los valores antes de guardar."
                )
                errores = True
            elif cod_equipo_t == "Seleccione...":
                st.error("Debes seleccionar un equipo de temperatura para validar el patrón.")
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
                        "No se ha guardado la visita | La verificación del patrón detecta una diferencia mayor a 0,5°C en: "
                        + ", ".join(campos_alerta)
                    )
                    st.json(verificacion)
                    errores = True
                else:
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
        else:
            st.session_state["cod_equipo_t"] = "Seleccione..."
            st.session_state["cod_equipo_v"] = "Seleccione..."

        if errores:
            st.stop()

        nuevo_prefill = {
            "fecha_visita": fecha_visita,
            "hora_visita": hora_medicion,
            "motivo_evaluacion": motivo_limpio,
            "nombre_personal_visita": nombre_personal,
            "cargo_personal_visita": cargo,
            "consultor_ist": email_usuario,
            "tipo_evaluacion": tipo_evaluacion,
        }

        if es_confort and confort_data:
            nuevo_prefill.update({
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

        if st.session_state.get("modo_edicion") and st.session_state.get("id_visita"):
            actualizado = actualizar_visita_inicio(
                st.session_state["id_visita"], visita_base_data, confort_data
            )
            if actualizado:
                st.session_state["visita_prefill"] = nuevo_prefill
                st.session_state["visitas_disponibles"] = get_visitas_por_cuv(cuv_visita)
                st.session_state["status_message"] = (
                    f"Visita {st.session_state['id_visita']} actualizada correctamente."
                )
                st.rerun()
            else:
                st.error("Error al actualizar la visita.")
        else:
            id_visita = guardar_visita_inicio(visita_base_data, confort_data)

        if id_visita is not None:
            st.session_state["id_visita"] = id_visita
            st.session_state["modo_edicion"] = True
            st.session_state["visita_prefill"] = nuevo_prefill
            st.session_state["visitas_disponibles"] = get_visitas_por_cuv(cuv_visita)
            st.session_state["status_message"] = (
                f"Datos de visita guardados correctamente. ID de visita: {id_visita}"
            )
            st.rerun()
        else:
            st.error("Error al guardar los datos de la visita.")
            # fin formulario 1

        # 3. Formulario 2: Mediciones de Áreas (Formularios Independientes)
        tipo_actual = st.session_state.get("visita_prefill", {}).get("tipo_evaluacion", "confort")
        if tipo_actual != "confort":
            st.subheader("Mediciones")
            st.info(
                "El registro de mediciones para evaluaciones de ventilación estará disponible próximamente."
            )
            st.markdown("---")
            st.subheader("Cierre")
            st.info("El cierre de visitas para este tipo de evaluación se habilitará en futuras versiones.")
            st.markdown("---")
            st.subheader("Finalizar Visita")
            st.warning(
                "Para visitas de ventilación aún no es posible registrar mediciones ni finalizar el proceso desde la aplicación."
            )
            st.stop()

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

                with st.expander(f"Área {i} - Haz clic para expandir", expanded=False):
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
                            st.session_state[vestimenta_key] = default_area.get("vestimenta_trabajador", "Seleccione...")
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
                            st.session_state[ventiladores_key] = "Sí" if default_area.get("cond_ventiladores") else "No"
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
                            st.session_state[inyeccion_key] = "Sí" if default_area.get("cond_inyeccion_extraccion") else "No"
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
                            mediciones_incompletas = any(value is None for value in [t_bul_seco, t_globo, hum_rel, vel_air])

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
        st.subheader("Cierre")
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
                use_container_width=True,
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
                                tipo_eval = st.session_state.get("visita_prefill", {}).get("tipo_evaluacion", "confort")
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
