import streamlit as st
import pandas as pd
from datetime import datetime, date, time as dt_time
import time
import os
import io
from dotenv import load_dotenv
from streamlit_cookies_controller import CookieController
from utils.helpers import autenticar_usuario, get_ct, precompletar_campos_ct, interpreter_pmv, get_cuv, \
    guardar_visita_inicio, guardar_visita_cierre, logout, insertar_medicion, get_met, check_resultado_pmv, \
    get_areas_options, get_motivo_eval, get_equipo_vel, get_equipo_temp, get_sector_especifico, get_puesto_trabajo, \
    get_posicion_trabajador, get_vestimenta_trabajador, comparar_patron
from pythermalcomfort.models import pmv_ppd_iso
from utils.informe import generar_informe

st.set_page_config(page_title="Informes Confort Térmico", layout="wide")

# Inicializa el controlador de cookies
cookie_controller = CookieController(key="app_cookies")

# Inicializar sesión del usuario
if "data_user" not in st.session_state:
    st.session_state["data_user"] = None

# Cargar variables de entorno desde un archivo .env
load_dotenv()

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

    # --- Inicialización en session_state ---
    if "df_filtrado" not in st.session_state:
        st.session_state["df_filtrado"] = pd.DataFrame()
    if "df_info_cuv" not in st.session_state:
        st.session_state["df_info_cuv"] = pd.DataFrame()
    if "input_cuv_str" not in st.session_state:
        st.session_state["input_cuv_str"] = ""

    # Búsqueda por CUV
    with st.container(border=True):
        st.info("Ingresa un CUV y haz clic en 'Buscar' para iniciar el registro de medición.")
        input_cuv = st.text_input("Ingresa el CUV:")
        if st.button("Buscar"):
            st.session_state["input_cuv_str"] = input_cuv.strip()
            resultados = get_ct(st.session_state["input_cuv_str"])
            if resultados and len(resultados) > 0:
                st.session_state["df_info_cuv"] = pd.DataFrame([resultados[0]])
                cuv = st.session_state["input_cuv_str"]
                st.success("Centro de trabajo encontrado en base de datos.")
            else:
                st.session_state["df_info_cuv"] = pd.DataFrame()
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

        st.write("")
        st.markdown("---")
        # formulario 1: Visita - datos visita + calibración inicial
        with st.form("visita_data_inicio"):
            # 2: Inicio
            st.subheader("Datos de la visita")

            fecha_visita = st.date_input("Fecha de visita", value=date.today())
            hora_medicion = st.time_input("Hora de medición", value=dt_time(hour=9, minute=0))
            temp_max = st.number_input("Temperatura máxima del día (°C)", min_value=-50.0, max_value=60.0,
                                       value=25.0, step=0.1)
            opc_motivos = get_motivo_eval()
            motivo_evaluacion = st.selectbox("Motivo de evaluación",
                                             options=["Seleccione..."]+ opc_motivos,
                                             index=0)
            nombre_personal = st.text_input("Nombre del personal SMU")
            cargo = st.text_input("Cargo", value="Administador/a")

            # 3: Calibración
            st.markdown("---")
            st.subheader("Verificación de parámetros")
            opc_equipos_temp = get_equipo_temp()
            opc_equipos_vel = get_equipo_vel()
            cod_equipo_t = st.selectbox("Equipo temperatura",
                                        options=["Seleccione..."] + opc_equipos_temp, index=0, key="cod_equipo_t")
            cod_equipo_v = st.selectbox("Equipo velocidad aire",
                                        options=["Seleccione..."] + opc_equipos_vel,
                                        index=0)
            patron_tbs = st.number_input("Patrón TBS", value=46.4, step=0.1)
            patron_tbh = st.number_input("Patrón TBH (Sólo modificar en caso necesario)", value=12.7, step=0.1)
            patron_tg = st.number_input("Patrón TG", value=69.8, step=0.1)
            st.write()
            verif_tbs_inicial = st.number_input("Verificación TBS inicial", value=None, step=0.1)
            verif_tbh_inicial = st.number_input("Verificación TBH inicial", value=None, step=0.1)
            verif_tg_inicial = st.number_input("Verificación TG inicial", value=None, step=0.1)
            submit_visita_inicio = st.form_submit_button(label="Guardar información visita",
                                                  type="primary",
                                                  use_container_width=True,
                                                  icon=":material/check_circle:")
            if submit_visita_inicio:
                if verif_tbs_inicial is None or verif_tbh_inicial is None or verif_tg_inicial is None:
                    st.error(
                        "Los campos de verificación son obligatorios. Por favor completa todos los valores antes de guardar.")
                elif cod_equipo_t == "Seleccione...":
                    st.error("Debes seleccionar un equipo de temperatura para validar el patrón.")
                else:
                    data_patron_medicion = (verif_tbs_inicial, verif_tbh_inicial, verif_tg_inicial)

                    verificacion = comparar_patron(data_patron_medicion, cod_equipo_t)
                    campos_alerta = [campo for campo, estado in verificacion.items() if estado == "alerta"]

                    if "error" in verificacion:
                        st.error(f"Error en la comparación de patrón: {verificacion['error']}")
                    elif "alerta" in verificacion.values():
                        st.error(
                            f"No se ha guardado la visita | La verificación del patrón detecta una diferencia mayor a 0,5°C en: {', '.join(campos_alerta)}")
                        st.json(verificacion)
                    else:
                        cuv_visita = get_cuv(st.session_state["input_cuv_str"])
                        visita_inicio_data = (
                            cuv_visita,  # cuv_visita
                            fecha_visita.strftime("%Y-%m-%d"),  # fecha_visita (formato ISO)
                            hora_medicion.strftime("%H:%M:%S"),  # hora_visita (formato 24h)
                            temp_max,  # temperatura_dia
                            motivo_evaluacion,  # motivo_evaluacion
                            nombre_personal,  # nombre_personal_visita
                            cargo,  # cargo_personal_visita
                            email_usuario,  # consultor_ist
                            cod_equipo_t,  # equipo_temp
                            cod_equipo_v,  # equipos_vel_air
                            patron_tbs,  # patron_tbs
                            verif_tbs_inicial,  # ver_tbs_ini
                            patron_tbh,  # patron_tbh
                            verif_tbh_inicial,  # ver_tbh_ini
                            patron_tg,  # patron_tg
                            verif_tg_inicial  # ver_th_ini
                        )

                        id_visita = guardar_visita_inicio(visita_inicio_data)
                        if id_visita is not None:
                            st.session_state["id_visita"] = id_visita  # Guardamos el id en session_state
                            st.success(f"Datos de visita guardados correctamente. ID de visita: {id_visita}")
                        else:
                            st.error("Error al guardar los datos de la visita.")

            # fin formulario 1

        # 3. Formulario 2: Mediciones de Áreas (Formularios Independientes)
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
                default_area = st.session_state.areas_data[area_idx] if area_idx < len(
                    st.session_state.areas_data) else {}

                with st.expander(f"Área {i} - Haz clic para expandir", expanded=False):
                    with st.form(key=f"form_area_{i}"):
                        # Captura de datos del formulario
                        nombre_area = st.selectbox(
                            f"Área {i}",
                            options=["Seleccione..."] + opc_areas_medicion,
                            key=f"area_sector_{i}"
                        )
                        sector_especifico = st.selectbox(f"Sector específico {i}",
                                                         options=["Seleccione..."] + opc_sector_especifico,
                                                         key=f"espec_sector_{i}")
                        puesto_trabajo = st.selectbox(f"Puesto de trabajo {i}",
                                                      options=["Seleccione..."] + opc_puesto_trabajo,
                                                      key=f"puesto_trabajo_{i}")
                        posicion_trabajador = st.selectbox(f"Posición {i}",
                                                           options=["Seleccione..."] + opc_posicion_trabajador,
                                                           key=f"pos_trabajador_{i}")
                        vestimenta_trabajador = st.selectbox(f"Vestimenta {i}",
                                                             options=["Seleccione..."] + opc_ventimenta_trabajador,
                                                             key=f"vestimenta_{i}")

                        # Mediciones
                        t_bul_seco = st.number_input(f"Temp. bulbo seco (°C) {i}", value=None,step=0.1,
                                                     key=f"tbs_{i}")
                        t_globo = st.number_input(f"Temp. globo (°C) {i}", value=None, step=0.1, key=f"tg_{i}")
                        hum_rel = st.number_input(f"Humedad relativa (%) {i}",value=None, step=0.1, key=f"hr_{i}")
                        vel_air = st.number_input(f"Velocidad del aire (m/s) {i}",Value=None, step=0.1,
                                                  key=f"vel_aire_{i}")

                        # Cálculo de PMV y PPD
                        met = get_met(puesto_trabajo)  # Puede depender del puesto de trabajo
                        clo = 0.5 if vestimenta_trabajador == "Habitual" else 1.0
                        resultados = pmv_ppd_iso(tdb=t_bul_seco, tr=t_globo, vr=vel_air, rh=hum_rel,
                                                 met=met, clo=clo,
                                                 model="7730-2005", limit_inputs=False)
                        pmv = resultados.pmv
                        ppd = resultados.ppd
                        resultado_medicion = check_resultado_pmv(pmv)

                        # Condiciones y observaciones
                        cond_techumbre = st.radio(f"Techumbre aislante {i}", ["Sí", "No"],
                                                  key=f"techumbre_{i}")
                        obs_techumbre = st.text_input(f"Obs. Techumbre {i}", key=f"obs_techumbre_{i}")
                        cond_techumbre = 1 if cond_techumbre == "Sí" else 0

                        cond_paredes = st.radio(f"Paredes aislantes {i}", ["Sí", "No"],
                                                key=f"paredes_{i}")
                        obs_paredes = st.text_input(f"Obs. Paredes {i}", key=f"obs_paredes_{i}")
                        cond_paredes = 1 if cond_paredes == "Sí" else 0

                        cond_vantanal = st.radio(f"Ventanas aislantes {i}", ["Sí", "No"],
                                                 key=f"ventanales_{i}")
                        obs_ventanal = st.text_input(f"Obs. Ventanas {i}", key=f"obs_ventanales_{i}")
                        cond_vantanal = 1 if cond_vantanal == "Sí" else 0

                        cond_aire_acond = st.radio(f"Aire acondicionado {i}", ["Sí", "No"],
                                                   key=f"aire_acond_{i}")
                        obs_aire_acond = st.text_input(f"Obs. Aire Acondicionado {i}",
                                                       key=f"obs_aire_acond_{i}")
                        cond_aire_acond = 1 if cond_aire_acond == "Sí" else 0

                        cond_ventiladores = st.radio(f"Ventiladores {i}", ["Sí", "No"],
                                                     key=f"ventiladores_{i}")
                        obs_ventiladores = st.text_input(f"Obs. Ventiladores {i}",
                                                         key=f"obs_ventiladores_{i}")
                        cond_ventiladores = 1 if cond_ventiladores == "Sí" else 0

                        cond_inyeccion_extraccion = st.radio(f"Inyección/Extracción {i}", ["Sí", "No"],
                                                             key=f"inyeccion_extrac_{i}")
                        obs_inyeccion_extraccion = st.text_input(f"Obs. Inyección {i}",
                                                                 key=f"obs_inyeccion_{i}")
                        cond_inyeccion_extraccion = 1 if cond_inyeccion_extraccion == "Sí" else 0

                        cond_ventanas = st.radio(f"Ventanas abiertas {i}", ["Sí", "No"],
                                                 key=f"ventanas_{i}")
                        obs_ventanas = st.text_input(f"Obs. Ventanas {i}", key=f"obs_ventanas_{i}")
                        cond_ventanas = 1 if cond_ventanas == "Sí" else 0

                        cond_puertas = st.radio(f"Puertas abiertas {i}", ["Sí", "No"],
                                                key=f"puertas_{i}")
                        obs_puertas = st.text_input(f"Obs. Puertas {i}", key=f"obs_puertas_{i}")
                        cond_puertas = 1 if cond_puertas == "Sí" else 0

                        cond_otras = st.radio(f"Puertas abiertas {i}", ["Sí", "No"], key=f"otras_{i}")
                        obs_otras = st.text_input(
                            f"¿Se identifican otras condiciones que pueden considerarse como disconfort térmico? {i}",
                            key=f"obs_otras_{i}")
                        cond_otras = 1 if cond_otras == "Sí" else 0

                        # Guardar medición
                        if st.form_submit_button(f"Guardar Área {i}"):
                            # Solo insertar si todos los datos están completos
                            if nombre_area != "Seleccione..." and sector_especifico != "Seleccione..." and puesto_trabajo != "Seleccione..." and posicion_trabajador != "Seleccione...":
                                id_medicion = insertar_medicion(id_visita, nombre_area,
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
                                                                cond_otras, obs_otras, met, clo)

                                if id_medicion:
                                    # Almacenar el ID de la medición en session_state pareado con el número de formulario
                                    st.session_state["mediciones_ids"][f"medicion_{i}"] = id_medicion
                                    st.success(
                                        f"Área {i} guardada con éxito. ID de la medición: {id_medicion}")
                                    for key, id_medicion in st.session_state["mediciones_ids"].items():
                                        st.write(
                                            f"**{key.replace('_', ' ').capitalize()}** - ID Medición: {id_medicion}")
                                else:
                                    st.error(f"No se pudo guardar la medición para el área {i}.")
                            else:
                                st.warning(f"Completa todos los campos antes de guardar el Área {i}.")

        # 4: Cierre
        st.subheader("Cierre")
        with st.form("visita_data_cierre"):
            verif_tbs_final = st.number_input("Verificación TBS inicial", value=None, step=0.1)
            verif_tbh_final = st.number_input("Verificación TBH inicial", value=None, step=0.1)
            verif_tg_final = st.number_input("Verificación TG inicial", value=None, step=0.1)
            comentarios_finales = st.text_area("Comentarios finales de evaluación", max_chars=1000)


            cierre_submitted = st.form_submit_button(label="Guardar verificación final", type="primary", use_container_width=True, icon=":material/check_circle:")
            if cierre_submitted:
                if verif_tbs_final is None or verif_tbh_final is None or verif_tg_final is None:
                    st.error(
                        "Los campos de verificación son obligatorios. Por favor completa todos los valores antes de guardar.")
                else:
                    equipo = st.session_state["cod_equipo_t"]
                    if not equipo:
                        st.error("No se encontró el equipo de temperatura para comparar el patrón.")
                    else:
                        data_patron_medicion = (verif_tbs_final,
                                               verif_tbh_final,
                                               verif_tg_final)
                        verificacion = comparar_patron(data_patron_medicion, equipo)
                        campos_alerta = [campo for campo, estado in verificacion.items() if estado == "alerta"]
                        if "error" in verificacion:
                            st.error(f"Error en la comparación de patrón: {verificacion['error']}")
                        elif "alerta" in verificacion.values():
                            st.error(f"No se ha guardado la verificación | La verificación del patrón detecta una diferencia mayor a 0,5°C. en: {', '.join(campos_alerta)}")
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

                            id_visita = st.session_state.get("id_visita")
                            st.write(id_visita)
                            if id_visita is not None:
                                actualizado = guardar_visita_cierre(id_visita, visita_cierre_data)
                                if actualizado:
                                    st.success(f"Visita actualizada correctamente. (ID: {id_visita})")
                                    st.session_state["visita_actualizada"] = True
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