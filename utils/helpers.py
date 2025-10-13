from db.mysql_utils import MySQLDatabaseManager
import streamlit as st
from streamlit.runtime.scriptrunner import get_script_run_ctx
import bcrypt
import re
import time
from streamlit_cookies_controller import CookieController
import logging
import pandas as pd


def validar_rut(rut):
    pass

def autenticar_usuario(username, password):
    db = MySQLDatabaseManager()
    try:
        query = "SELECT name, pass, email FROM usuarios WHERE email = %s"
        user = db.fetch_one(query, (username,))
        if user and bcrypt.checkpw(password.encode(), user["pass"].encode()):
            return {"nombre": user["name"], "email": user["email"]}
        return None
    finally:
        db.close()


def make_sidebar():
        if st.session_state.get("data_user", None):
            st.write(f"👋 Hola, {st.session_state['data_user']['name']}!")

            if st.button("Cerrar sesión"):
                cookie_controller = CookieController()
                cookie_controller.set("user_data", "", max_age=0)
                # cookie_controller.remove("user_data")
                st.session_state["data_user"] = None
                st.success("Has cerrado sesión correctamente.")
                time.sleep(2)
                st.rerun()

        #elif get_current_page_name() != "app":
            # If anyone tries to access a secret page without being logged in,
            # redirect them to the login page


def logout():
    cookie_controller = CookieController()
    cookie_controller.set("user_data", "", max_age=0)
    #cookie_controller.remove("user_data")
    del st.session_state["data_user"]
    st.session_state.clear()
    st.success("Has cerrado sesión correctamente.")
    time.sleep(2)
    st.rerun()


def login(username, password):
    user_data = autenticar_usuario(username, password)
    if user_data:
        # Guardar datos en cookies
        #cookie_controller.set("user_data", user_data, max_age=7200)
        #st.session_state["user_data"] = user_data
        st.success("Login exitoso")
        time.sleep(1)
        #st.switch_page("pages/home.py")
        st.rerun()
    else:
        st.error("Usuario o contraseña incorrectos.")

# Checkear sesión al inicio
def check_session():
    cookie_controller = CookieController()
    #st.write("check de cookie")
    user_data = cookie_controller.get("user_data")
    #st.write(cookie_controller.get("user_data"))
    if user_data:
        st.session_state["data_user"] = user_data
        #st.write(st.session_state["user_data"])
    return user_data

def get_ct(cuv):
    db = MySQLDatabaseManager()
    try:
        query = """
            SELECT cuv, rut, razon_social, rut2, nombre_ct, direccion_ct, comuna_ct, region_ct, region_num_ct
            FROM centros_trabajo 
            WHERE cuv = %s
        """
        db.cursor.execute(query, (cuv,))
        resultados = db.cursor.fetchall()
        return resultados
    finally:
        db.close()

def get_visita(id_visita):
    db = MySQLDatabaseManager()
    try:
        query = """
            SELECT v.*, u.name AS consultor_nombre
            FROM visitas v
            JOIN usuarios u ON v.consultor_ist = u.email
            WHERE v.id_visita = %s
            """
        db.cursor.execute(query, (id_visita,))
        resultados = db.cursor.fetchall()
        return pd.DataFrame(resultados)
    finally:
        db.close()

def get_visitas_por_cuv(cuv):
    db = MySQLDatabaseManager()
    try:
        query = """
            SELECT * FROM visitas
            WHERE cuv_visita = %s
            ORDER BY fecha_visita DESC
        """
        db.cursor.execute(query, (int(cuv),))
        resultados = db.cursor.fetchall()
        return pd.DataFrame(resultados)
    finally:
        db.close()

def get_mediciones(id_visita):
    db = MySQLDatabaseManager()
    try:
        query = "SELECT * FROM mediciones WHERE visita_id = %s"
        db.cursor.execute(query, (id_visita,))
        resultados = db.cursor.fetchall()
        return pd.DataFrame(resultados)
    finally:
        db.close()

def get_equipos():
    db = MySQLDatabaseManager()
    try:
        query = "SELECT * FROM equipos_medicion"
        db.cursor.execute(query)
        resultados = db.cursor.fetchall()
        return pd.DataFrame(resultados)
    finally:
        db.close()

def get_all_cuvs_with_visits():
    db = MySQLDatabaseManager()
    try:
        query = "SELECT DISTINCT cuv_visita FROM visitas"
        db.cursor.execute(query)
        resultados = db.cursor.fetchall()
        return pd.DataFrame(resultados)
    finally:
        db.close()

def precompletar_campos_ct(data_ct):
    razon_social = st.text_input("Razón Social",
                                 value=str(data_ct.get("razon_social", "")),
                                 disabled=True)
    rut = st.text_input("RUT",
                        value=str(data_ct.get("rut", "")),
                        disabled=True)
    nombre_local = st.text_input("Nombre de Local",
                                 value=str(data_ct.get("nombre_ct", "")),
                                 disabled=True)
    direccion = st.text_input("Dirección",
                              value=str(data_ct.get("direccion_ct", "")),
                              disabled=True)
    comuna = st.text_input("Comuna",
                           value=str(data_ct.get("comuna_ct", "")),
                           disabled=True)
    region = st.text_input("Región",
                           value=str(data_ct.get("region_ct", "")),
                           disabled=True)
    cuv_val = st.text_input("CUV",
                            value=str(data_ct.get("cuv", "")),
                            disabled=True)

    return {
        "razon_social": razon_social,
        "rut": rut,
        "nombre_local": nombre_local,
        "direccion": direccion,
        "comuna": comuna,
        "region": region,
        "cuv": cuv_val
    }

def interpreter_pmv(pmv):
    if pmv >= 2.5:
        interpretacion = "Calurosa"
    elif pmv >= 1.5:
        interpretacion = "Cálida"
    elif pmv >= 0.5:
        interpretacion = "Ligeramente cálida"
    elif pmv > -0.5:
        interpretacion = "Neutra - Confortable"
    elif pmv > -1.5:
        interpretacion = "Ligeramente fresca"
    elif pmv > -2.5:
        interpretacion = "Fresca"
    else:
        interpretacion = "Fría"
    return interpretacion

@st.dialog("Confirmación de guardado")
def dialogo_confirmacion():
    st.write("¿Estás segur@ que quieres guardar los datos de la medición?")
    with st.container():
        col1, col2 = st.columns(2)
        aceptar = col1.button("Aceptar", key="dialog_aceptar")
        cancelar = col2.button("Cancelar", key="dialog_cancelar")
    if aceptar:
        st.session_state["confirm_save"] = True
        st.rerun()
    if cancelar:
        st.session_state["confirm_save"] = False
        st.rerun()

def confirmar_guardado():
    if "confirm_save" not in st.session_state:
        st.session_state["confirm_save"] = None
    dialogo_confirmacion()
    return st.session_state.get("confirm_save")

def get_cuv(campo_cuv):
    cuv = campo_cuv
    return cuv

def guardar_visita_inicio(visita_data):
    db = MySQLDatabaseManager()
    try:
        (cuv_visita, fecha_visita, hora_visita, temperatura_dia, motivo_evaluacion,
         nombre_personal_visita, cargo_personal_visita, consultor_ist, equipo_temp, equipo_vel_air,
         patron_tbs, ver_tbs_ini, patron_tbh, ver_tbh_ini, patron_tg, ver_tg_ini) = visita_data
        # paso 1: Obtener datos consultor
        query_user = "SELECT name, cargo, zonal FROM usuarios WHERE email = %s"
        db.cursor.execute(query_user, (consultor_ist,))
        user_row = db.cursor.fetchone()
        if not user_row:
            logging.error("No se encontró el usuario con email %s", consultor_ist)
            return None
        consultor_name, consultor_cargo, consultor_zonal = user_row['name'], user_row['cargo'], user_row['zonal']

        # paso2: Obtener el id del equipo de temperatura
        query_equipo_temp = "SELECT id_equipo FROM equipos_medicion WHERE equipo_dicc = %s"
        db.cursor.execute(query_equipo_temp, (equipo_temp,))
        row_temp = db.cursor.fetchone()
        if not row_temp:
            logging.error("No se encontró el equipo de temperatura con equipo_dicc %s", equipo_temp)
            return None
        id_equipo_temp = row_temp['id_equipo']

        # paso 3: Obtener el id del equipo de velocidad
        query_equipo_vel = "SELECT id_equipo FROM equipos_medicion WHERE equipo_dicc = %s"
        db.cursor.execute(query_equipo_vel, (equipo_vel_air,))
        row_vel = db.cursor.fetchone()
        if not row_vel:
            logging.error("No se encontró el equipo de velocidad con equipo_dicc %s", equipo_vel_air)
            return None
        id_equipo_vel = row_vel['id_equipo']

        # paso 4: Preparo la consulta
        query = """
            INSERT INTO visitas (
                cuv_visita,
                fecha_visita,
                hora_visita,
                temperatura_dia,
                motivo_evaluacion,
                nombre_personal_visita,
                cargo_personal_visita,
                consultor_ist,
                equipo_temp,
                equipo_vel_air,
                patron_tbs,
                ver_tbs_ini,
                patron_tbh,
                ver_tbh_ini,
                patron_tg,
                ver_tg_ini,
                consultor_cargo,
                consultor_zonal
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        params = (
            cuv_visita,
            fecha_visita,
            hora_visita,
            temperatura_dia,
            motivo_evaluacion,
            nombre_personal_visita,
            cargo_personal_visita,
            consultor_ist,
            id_equipo_temp,
            id_equipo_vel,
            patron_tbs,
            ver_tbs_ini,
            patron_tbh,
            ver_tbh_ini,
            patron_tg,
            ver_tg_ini,
            consultor_cargo,
            consultor_zonal
        )
        # paso 5: Ejecuto la consulta
        db.cursor.execute(query, params)
        db.connection.commit()
        id_visita = db.cursor.lastrowid
        return id_visita

    except Exception as e:
        logging.error(f"Error al guardar la visita: {e}")
        return None
    finally:
        db.close()


def actualizar_visita_inicio(id_visita, visita_data):
    db = MySQLDatabaseManager()
    try:
        (cuv_visita, fecha_visita, hora_visita, temperatura_dia, motivo_evaluacion,
         nombre_personal_visita, cargo_personal_visita, consultor_ist, equipo_temp, equipo_vel_air,
         patron_tbs, ver_tbs_ini, patron_tbh, ver_tbh_ini, patron_tg, ver_tg_ini) = visita_data

        query_user = "SELECT name, cargo, zonal FROM usuarios WHERE email = %s"
        db.cursor.execute(query_user, (consultor_ist,))
        user_row = db.cursor.fetchone()
        if not user_row:
            logging.error("No se encontró el usuario con email %s", consultor_ist)
            return False
        consultor_name, consultor_cargo, consultor_zonal = user_row['name'], user_row['cargo'], user_row['zonal']

        query_equipo_temp = "SELECT id_equipo FROM equipos_medicion WHERE equipo_dicc = %s"
        db.cursor.execute(query_equipo_temp, (equipo_temp,))
        row_temp = db.cursor.fetchone()
        if not row_temp:
            logging.error("No se encontró el equipo de temperatura con equipo_dicc %s", equipo_temp)
            return False
        id_equipo_temp = row_temp['id_equipo']

        query_equipo_vel = "SELECT id_equipo FROM equipos_medicion WHERE equipo_dicc = %s"
        db.cursor.execute(query_equipo_vel, (equipo_vel_air,))
        row_vel = db.cursor.fetchone()
        if not row_vel:
            logging.error("No se encontró el equipo de velocidad con equipo_dicc %s", equipo_vel_air)
            return False
        id_equipo_vel = row_vel['id_equipo']

        query = """
            UPDATE visitas
            SET cuv_visita = %s,
                fecha_visita = %s,
                hora_visita = %s,
                temperatura_dia = %s,
                motivo_evaluacion = %s,
                nombre_personal_visita = %s,
                cargo_personal_visita = %s,
                consultor_ist = %s,
                equipo_temp = %s,
                equipo_vel_air = %s,
                patron_tbs = %s,
                ver_tbs_ini = %s,
                patron_tbh = %s,
                ver_tbh_ini = %s,
                patron_tg = %s,
                ver_tg_ini = %s,
                consultor_cargo = %s,
                consultor_zonal = %s
            WHERE id_visita = %s
        """

        params = (
            cuv_visita,
            fecha_visita,
            hora_visita,
            temperatura_dia,
            motivo_evaluacion,
            nombre_personal_visita,
            cargo_personal_visita,
            consultor_ist,
            id_equipo_temp,
            id_equipo_vel,
            patron_tbs,
            ver_tbs_ini,
            patron_tbh,
            ver_tbh_ini,
            patron_tg,
            ver_tg_ini,
            consultor_cargo,
            consultor_zonal,
            id_visita
        )

        db.cursor.execute(query, params)
        db.connection.commit()
        return True

    except Exception as e:
        logging.error(f"Error al actualizar la visita: {e}")
        return False
    finally:
        db.close()


def guardar_visita_cierre(id_visita, dato_cierre):
    query = """
        UPDATE visitas
        SET ver_tbs_fin = %s,
            ver_tbh_fin = %s,
            ver_tg_fin = %s,
            note_visita = %s
        WHERE id_visita = %s
    """
    params = (
        dato_cierre[0],  # ver_tbs_fin
        dato_cierre[1],  # ver_tbh_fin
        dato_cierre[2],  # ver_tg_fin
        dato_cierre[3],  # note_visita
        id_visita  # id_visita para la cláusula WHERE
    )

    db = MySQLDatabaseManager()
    try:
        db.cursor.execute(query, params)
        db.connection.commit()
        return True
    except Exception as e:
        print("Error al actualizar la visita:", e)
        return False
    finally:
        db.close()

def get_met(puesto_trabajo):
    if puesto_trabajo == "Cajera":
        return 1.1
    elif puesto_trabajo == "Reponedor":
        return 1.2
    elif puesto_trabajo in ["Bodeguero", "Recepcionista"]:
        return 1.89
    else:
        return 1.1

def check_resultado_pmv(pmv):
    return "Cumple" if -1 <= pmv <= 1 else "No cumple"

def interpret_pmv(pmv_value):
    if pmv_value >= 2.5:
        return "Calurosa"
    elif pmv_value >= 1.5:
        return "Cálida"
    elif pmv_value >= 0.5:
        return "Ligeramente cálida"
    elif pmv_value > -0.5:
        return "Neutra - Confortable"
    elif pmv_value > -1.5:
        return "Ligeramente fresca"
    elif pmv_value > -2.5:
        return "Fresca"
    else:
        return "Fría"


def insertar_medicion(
                        id_visita,
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
                        cond_otras, obs_otras, met, clo):
    ppd = float(ppd)
    pmv= float(pmv)
    db = MySQLDatabaseManager()
    try:
        # Validación básica
        if not id_visita or nombre_area == "Seleccione..." or sector_especifico == "Seleccione..." or puesto_trabajo == "Seleccione...":
            logging.error("Datos de medición incompletos. No se insertará en la base de datos.")
            return None

        # Preparar la consulta INSERT
        query = """
            INSERT INTO mediciones (
                visita_id, nombre_area, sector_especifico, puesto_trabajo, 
                posicion_trabajador, vestimenta_trabajador, t_bul_seco, t_globo,
                hum_rel, vel_air, ppd, pmv, resultado_medicion, cond_techumbre, obs_techumbre,
                cond_paredes, obs_paredes, cond_vantanal, obs_ventanal, cond_aire_acond, obs_aire_acond,
                cond_ventiladores, obs_ventiladores, cond_inyeccion_extraccion, obs_inyeccion_extraccion,
                cond_ventanas, obs_ventanas, cond_puertas, obs_puertas, cond_otras, obs_otras, met, clo
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
            )
        """


        db.cursor.execute(query, (
            id_visita, nombre_area, sector_especifico, puesto_trabajo,
            posicion_trabajador, vestimenta_trabajador, t_bul_seco, t_globo,
            hum_rel, vel_air, ppd, pmv, resultado_medicion, cond_techumbre,
            obs_techumbre, cond_paredes, obs_paredes, cond_vantanal, obs_ventanal,
            cond_aire_acond, obs_aire_acond, cond_ventiladores, obs_ventiladores,
            cond_inyeccion_extraccion, obs_inyeccion_extraccion, cond_ventanas,
            obs_ventanas, cond_puertas, obs_puertas, cond_otras, obs_otras, met, clo
        ))
        db.connection.commit()

        id_medicion = db.cursor.lastrowid
        logging.info(f"Medición insertada con éxito. ID: {id_medicion}")

        return id_medicion

    except Exception as e:
        logging.error(f"Error al insertar la medición: {e}")
        return None
    finally:
        db.close()


def actualizar_medicion(
                        id_medicion,
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
                        cond_otras, obs_otras, met, clo):
    db = MySQLDatabaseManager()
    try:
        query = """
            UPDATE mediciones
            SET nombre_area = %s,
                sector_especifico = %s,
                puesto_trabajo = %s,
                posicion_trabajador = %s,
                vestimenta_trabajador = %s,
                t_bul_seco = %s,
                t_globo = %s,
                hum_rel = %s,
                vel_air = %s,
                ppd = %s,
                pmv = %s,
                resultado_medicion = %s,
                cond_techumbre = %s,
                obs_techumbre = %s,
                cond_paredes = %s,
                obs_paredes = %s,
                cond_vantanal = %s,
                obs_ventanal = %s,
                cond_aire_acond = %s,
                obs_aire_acond = %s,
                cond_ventiladores = %s,
                obs_ventiladores = %s,
                cond_inyeccion_extraccion = %s,
                obs_inyeccion_extraccion = %s,
                cond_ventanas = %s,
                obs_ventanas = %s,
                cond_puertas = %s,
                obs_puertas = %s,
                cond_otras = %s,
                obs_otras = %s,
                met = %s,
                clo = %s
            WHERE id_medicion = %s
        """

        params = (
            nombre_area,
            sector_especifico,
            puesto_trabajo,
            posicion_trabajador,
            vestimenta_trabajador,
            t_bul_seco,
            t_globo,
            hum_rel,
            vel_air,
            float(ppd),
            float(pmv),
            resultado_medicion,
            cond_techumbre,
            obs_techumbre,
            cond_paredes,
            obs_paredes,
            cond_vantanal,
            obs_ventanal,
            cond_aire_acond,
            obs_aire_acond,
            cond_ventiladores,
            obs_ventiladores,
            cond_inyeccion_extraccion,
            obs_inyeccion_extraccion,
            cond_ventanas,
            obs_ventanas,
            cond_puertas,
            obs_puertas,
            cond_otras,
            obs_otras,
            met,
            clo,
            id_medicion
        )

        db.cursor.execute(query, params)
        db.connection.commit()
        return True

    except Exception as e:
        logging.error(f"Error al actualizar la medición: {e}")
        return False
    finally:
        db.close()


def get_areas_options():
    db = MySQLDatabaseManager()
    try:
        query = "SELECT DISTINCT nombre_area FROM areas_medicion"
        db.cursor.execute(query)
        resultados = db.cursor.fetchall()
        return [row['nombre_area'] for row in resultados]
    finally:
        db.close()

def get_motivo_eval():
    db = MySQLDatabaseManager()
    try:
        query = "SELECT DISTINCT nombre_motivo FROM motivo_evaluacion"
        db.cursor.execute(query)
        resultados = db.cursor.fetchall()
        return [row['nombre_motivo'] for row in resultados]
    finally:
        db.close()

def get_equipo_temp():
    db = MySQLDatabaseManager()
    try:
        query = "SELECT DISTINCT equipo_dicc FROM equipos_medicion WHERE tipo = 'Temperatura'"
        db.cursor.execute(query)
        resultados = db.cursor.fetchall()
        return [row['equipo_dicc'] for row in resultados]
    finally:
        db.close()

def get_equipo_vel():
    db = MySQLDatabaseManager()
    try:
        query = "SELECT DISTINCT equipo_dicc FROM equipos_medicion WHERE tipo = 'Velocidad'"
        db.cursor.execute(query)
        resultados = db.cursor.fetchall()
        return [row['equipo_dicc'] for row in resultados]
    finally:
        db.close()

def get_equipo_dicc_por_id(id_equipo):
    if id_equipo is None:
        return None
    db = MySQLDatabaseManager()
    try:
        query = "SELECT equipo_dicc FROM equipos_medicion WHERE id_equipo = %s"
        db.cursor.execute(query, (id_equipo,))
        resultado = db.cursor.fetchone()
        return resultado['equipo_dicc'] if resultado else None
    finally:
        db.close()

def get_sector_especifico():
    db = MySQLDatabaseManager()
    try:
        query = "SELECT DISTINCT nombre_sector_especifico FROM sector_especifico"
        db.cursor.execute(query)
        resultados = db.cursor.fetchall()
        return [row['nombre_sector_especifico'] for row in resultados]
    finally:
        db.close()

def get_puesto_trabajo():
    db = MySQLDatabaseManager()
    try:
        query = "SELECT DISTINCT nombre_puesto_trabajo FROM puestos_trabajo"
        db.cursor.execute(query)
        resultados = db.cursor.fetchall()
        return [row['nombre_puesto_trabajo'] for row in resultados]
    finally:
        db.close()

def get_posicion_trabajador():
    db = MySQLDatabaseManager()
    try:
        query = "SELECT DISTINCT nombre_posicion_trabajador FROM posicion_trabajador"
        db.cursor.execute(query)
        resultados = db.cursor.fetchall()
        return [row['nombre_posicion_trabajador'] for row in resultados]
    finally:
        db.close()

def get_vestimenta_trabajador():
    db = MySQLDatabaseManager()
    try:
        query = "SELECT DISTINCT nombre_vestimenta_trabajador FROM vestimenta_trabajador"
        db.cursor.execute(query)
        resultados = db.cursor.fetchall()
        return [row['nombre_vestimenta_trabajador'] for row in resultados]
    finally:
        db.close()


def comparar_patron(data_patron_medicion, equipo_dicc):
    db = MySQLDatabaseManager()
    try:
        query = "SELECT patron_tbs, patron_tbh, patron_tg FROM equipos_medicion WHERE equipo_dicc = %s"
        db.cursor.execute(query, (equipo_dicc,))
        patron = db.cursor.fetchone()

        if not patron:
            return {"error": "No se encontró el equipo en la base de datos."}

        resultado = {}
        keys = ['patron_tbs', 'patron_tbh', 'patron_tg']

        for idx, key in enumerate(keys):
            valor_formulario = float(data_patron_medicion[idx])
            valor_patron = float(patron[key])
            diferencia = abs(valor_formulario - valor_patron)
            resultado[key] = "ok" if diferencia <= 0.5 else "alerta"

        return resultado

    except Exception as e:
        return {"error": str(e)}
    finally:
        db.close()
