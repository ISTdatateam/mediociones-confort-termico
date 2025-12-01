from db.mysql_utils import MySQLDatabaseManager
import streamlit as st
from streamlit.runtime.scriptrunner import get_script_run_ctx
import bcrypt
import re
import time
from streamlit_cookies_controller import CookieController
import logging
import pandas as pd
from datetime import datetime, date, time as dt_time, timedelta
import math


def normalizar_fecha_mysql(valor):
    """Convierte valores devueltos por MySQL a objetos ``date`` consistentes."""
    if valor is None or valor == "":
        return None

    if isinstance(valor, date) and not isinstance(valor, datetime):
        return valor

    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, str):
        # Intentamos varios formatos comunes; si fallan, dejamos el valor original
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
    """Convierte valores de hora devueltos por MySQL a objetos ``time``."""
    if valor is None or valor == "":
        return None

    if isinstance(valor, dt_time):
        return valor

    if isinstance(valor, datetime):
        return valor.time()

    if isinstance(valor, timedelta):
        # MySQL Connector entrega campos TIME como timedelta
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
            SELECT
                v.id_visita,
                v.cuv_visita,
                v.fecha_visita,
                v.hora_visita,
                v.motivo_evaluacion,
                v.nombre_personal_visita,
                v.cargo_personal_visita,
                v.consultor_ist,
                v.consultor_cargo,
                v.consultor_zonal,
                v.note_visita,
                v.tipo_evaluacion,
                u.name AS consultor_nombre,
                ec.temperatura_dia,
                ec.equipo_temp,
                ec.equipo_vel_air,
                ec.patron_tbs,
                ec.ver_tbs_ini,
                ec.patron_tbh,
                ec.ver_tbh_ini,
                ec.patron_tg,
                ec.ver_tg_ini,
                ec.ver_tbs_fin,
                ec.ver_tbh_fin,
                ec.ver_tg_fin,
                evv.equipo_temp AS evv_equipo_temp,
                evv.equipo_vel_air AS evv_equipo_vel_air
            FROM visitas v
            JOIN usuarios u ON v.consultor_ist = u.email
            LEFT JOIN ev_confort ec ON ec.visita_id = v.id_visita
            LEFT JOIN ev_ventilacion evv ON evv.visita_id = v.id_visita
            WHERE v.id_visita = %s
            """
        db.cursor.execute(query, (id_visita,))
        resultados = db.cursor.fetchall()
        df = pd.DataFrame(resultados)
        if not df.empty:
            if "fecha_visita" in df.columns:
                df["fecha_visita"] = df["fecha_visita"].apply(normalizar_fecha_mysql)
            if "hora_visita" in df.columns:
                df["hora_visita"] = df["hora_visita"].apply(normalizar_hora_mysql)
        return df
    finally:
        db.close()

def get_visitas_por_cuv(cuv):
    db = MySQLDatabaseManager()
    try:
        query = """
            SELECT
                v.id_visita,
                v.fecha_visita,
                v.hora_visita,
                v.motivo_evaluacion,
                v.tipo_evaluacion
            FROM visitas v
            WHERE v.cuv_visita = %s
            ORDER BY v.fecha_visita DESC
        """
        db.cursor.execute(query, (int(cuv),))
        resultados = db.cursor.fetchall()
        df = pd.DataFrame(resultados)
        if not df.empty:
            if "fecha_visita" in df.columns:
                df["fecha_visita"] = df["fecha_visita"].apply(normalizar_fecha_mysql)
            if "hora_visita" in df.columns:
                df["hora_visita"] = df["hora_visita"].apply(normalizar_hora_mysql)
        return df
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

def _obtener_datos_consultor(db, consultor_ist):
    """Obtiene el cargo y zonal asociados a un consultor."""
    query_user = "SELECT cargo, zonal FROM usuarios WHERE email = %s"
    db.cursor.execute(query_user, (consultor_ist,))
    return db.cursor.fetchone()


def _obtener_id_equipo_por_dicc(db, equipo_dicc):
    """Retorna el identificador del equipo a partir de su código público."""
    if not equipo_dicc or equipo_dicc == "Seleccione...":
        return None
    query_equipo = "SELECT id_equipo FROM equipos_medicion WHERE equipo_dicc = %s"
    db.cursor.execute(query_equipo, (equipo_dicc,))
    row = db.cursor.fetchone()
    return row['id_equipo'] if row else None


def _upsert_ev_ventilacion(db, visita_id, ventilacion_data):
    """Inserta o actualiza los equipos de una evaluación de ventilación."""

    if not ventilacion_data:
        return

    equipo_temp_id = _obtener_id_equipo_por_dicc(db, ventilacion_data.get("equipo_temp"))
    equipo_vel_id = _obtener_id_equipo_por_dicc(db, ventilacion_data.get("equipo_vel_air"))

    if equipo_temp_id is None and equipo_vel_id is None:
        logging.warning(
            "No se proporcionaron equipos válidos para la evaluación de ventilación %s",
            visita_id,
        )
        return

    query_ventilacion = """
        INSERT INTO ev_ventilacion (
            visita_id,
            equipo_temp,
            equipo_vel_air
        ) VALUES (%s, %s, %s)
        ON DUPLICATE KEY UPDATE
            equipo_temp = VALUES(equipo_temp),
            equipo_vel_air = VALUES(equipo_vel_air)
    """

    db.cursor.execute(query_ventilacion, (visita_id, equipo_temp_id, equipo_vel_id))


def guardar_visita_inicio(visita_data, confort_data=None, ventilacion_data=None):
    """Inserta una visita y, de ser necesario, su información de confort térmico."""
    db = MySQLDatabaseManager()
    try:
        consultor_info = _obtener_datos_consultor(db, visita_data["consultor_ist"])
        if not consultor_info:
            logging.error("No se encontró el usuario con email %s", visita_data["consultor_ist"])
            return None

        tipo_evaluacion = visita_data.get("tipo_evaluacion", "confort")
        equipo_temp_id = None
        equipo_vel_id = None

        if tipo_evaluacion == "confort" and not confort_data:
            logging.error("Se requieren datos de confort para guardar una visita de tipo confort.")
            return None

        if tipo_evaluacion == "confort" and confort_data:
            equipo_temp_id = _obtener_id_equipo_por_dicc(db, confort_data.get("equipo_temp"))
            equipo_vel_id = _obtener_id_equipo_por_dicc(db, confort_data.get("equipo_vel_air"))
            if equipo_temp_id is None or equipo_vel_id is None:
                logging.error("No se encontraron los equipos indicados para la visita de confort.")
                return None

        query_visita = """
            INSERT INTO visitas (
                cuv_visita,
                fecha_visita,
                hora_visita,
                motivo_evaluacion,
                nombre_personal_visita,
                cargo_personal_visita,
                consultor_ist,
                consultor_cargo,
                consultor_zonal,
                note_visita,
                tipo_evaluacion
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """

        params_visita = (
            visita_data["cuv_visita"],
            visita_data["fecha_visita"],
            visita_data["hora_visita"],
            visita_data.get("motivo_evaluacion"),
            visita_data.get("nombre_personal_visita"),
            visita_data.get("cargo_personal_visita"),
            visita_data["consultor_ist"],
            consultor_info["cargo"],
            consultor_info["zonal"],
            visita_data.get("note_visita"),
            tipo_evaluacion,
        )

        db.cursor.execute(query_visita, params_visita)
        id_visita = db.cursor.lastrowid

        if tipo_evaluacion == "confort" and confort_data:
            query_confort = """
                INSERT INTO ev_confort (
                    visita_id,
                    temperatura_dia,
                    equipo_temp,
                    equipo_vel_air,
                    patron_tbs,
                    ver_tbs_ini,
                    patron_tbh,
                    ver_tbh_ini,
                    patron_tg,
                    ver_tg_ini,
                    ver_tbs_fin,
                    ver_tbh_fin,
                    ver_tg_fin
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """

            params_confort = (
                id_visita,
                confort_data.get("temperatura_dia"),
                equipo_temp_id,
                equipo_vel_id,
                confort_data.get("patron_tbs"),
                confort_data.get("ver_tbs_ini"),
                confort_data.get("patron_tbh"),
                confort_data.get("ver_tbh_ini"),
                confort_data.get("patron_tg"),
                confort_data.get("ver_tg_ini"),
                confort_data.get("ver_tbs_fin"),
                confort_data.get("ver_tbh_fin"),
                confort_data.get("ver_tg_fin"),
            )

            db.cursor.execute(query_confort, params_confort)
        elif tipo_evaluacion == "ventilacion" and ventilacion_data:
            _upsert_ev_ventilacion(db, id_visita, ventilacion_data)

        db.connection.commit()
        return id_visita

    except Exception as e:
        logging.error(f"Error al guardar la visita: {e}")
        if db.connection:
            db.connection.rollback()
        return None
    finally:
        db.close()


def actualizar_visita_inicio(id_visita, visita_data, confort_data=None, ventilacion_data=None):
    """Actualiza la información base de una visita y sus datos de confort."""
    db = MySQLDatabaseManager()
    try:
        consultor_info = _obtener_datos_consultor(db, visita_data["consultor_ist"])
        if not consultor_info:
            logging.error("No se encontró el usuario con email %s", visita_data["consultor_ist"])
            return False

        tipo_evaluacion = visita_data.get("tipo_evaluacion", "confort")
        equipo_temp_id = None
        equipo_vel_id = None

        if tipo_evaluacion == "confort" and not confort_data:
            logging.error("Se requieren datos de confort para actualizar una visita de tipo confort.")
            return False

        if tipo_evaluacion == "confort" and confort_data:
            equipo_temp_id = _obtener_id_equipo_por_dicc(db, confort_data.get("equipo_temp"))
            equipo_vel_id = _obtener_id_equipo_por_dicc(db, confort_data.get("equipo_vel_air"))
            if equipo_temp_id is None or equipo_vel_id is None:
                logging.error("No se encontraron los equipos indicados para la visita de confort.")
                return False

        query_visita = """
            UPDATE visitas
            SET cuv_visita = %s,
                fecha_visita = %s,
                hora_visita = %s,
                motivo_evaluacion = %s,
                nombre_personal_visita = %s,
                cargo_personal_visita = %s,
                consultor_ist = %s,
                consultor_cargo = %s,
                consultor_zonal = %s,
                tipo_evaluacion = %s
            WHERE id_visita = %s
        """

        params_visita = (
            visita_data["cuv_visita"],
            visita_data["fecha_visita"],
            visita_data["hora_visita"],
            visita_data.get("motivo_evaluacion"),
            visita_data.get("nombre_personal_visita"),
            visita_data.get("cargo_personal_visita"),
            visita_data["consultor_ist"],
            consultor_info["cargo"],
            consultor_info["zonal"],
            tipo_evaluacion,
            id_visita,
        )

        db.cursor.execute(query_visita, params_visita)

        if tipo_evaluacion == "confort" and confort_data:
            query_existe = "SELECT 1 FROM ev_confort WHERE visita_id = %s"
            db.cursor.execute(query_existe, (id_visita,))
            existe_confort = db.cursor.fetchone() is not None

            if existe_confort:
                query_update = """
                    UPDATE ev_confort
                    SET temperatura_dia = %s,
                        equipo_temp = %s,
                        equipo_vel_air = %s,
                        patron_tbs = %s,
                        ver_tbs_ini = %s,
                        patron_tbh = %s,
                        ver_tbh_ini = %s,
                        patron_tg = %s,
                        ver_tg_ini = %s
                    WHERE visita_id = %s
                """

                params_update = (
                    confort_data.get("temperatura_dia"),
                    equipo_temp_id,
                    equipo_vel_id,
                    confort_data.get("patron_tbs"),
                    confort_data.get("ver_tbs_ini"),
                    confort_data.get("patron_tbh"),
                    confort_data.get("ver_tbh_ini"),
                    confort_data.get("patron_tg"),
                    confort_data.get("ver_tg_ini"),
                    id_visita,
                )

                db.cursor.execute(query_update, params_update)
            else:
                query_insert = """
                    INSERT INTO ev_confort (
                        visita_id,
                        temperatura_dia,
                        equipo_temp,
                        equipo_vel_air,
                        patron_tbs,
                        ver_tbs_ini,
                        patron_tbh,
                        ver_tbh_ini,
                        patron_tg,
                        ver_tg_ini
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """

                params_insert = (
                    id_visita,
                    confort_data.get("temperatura_dia"),
                    equipo_temp_id,
                    equipo_vel_id,
                    confort_data.get("patron_tbs"),
                    confort_data.get("ver_tbs_ini"),
                    confort_data.get("patron_tbh"),
                    confort_data.get("ver_tbh_ini"),
                    confort_data.get("patron_tg"),
                    confort_data.get("ver_tg_ini"),
                )

                db.cursor.execute(query_insert, params_insert)
            db.cursor.execute("DELETE FROM ev_ventilacion WHERE visita_id = %s", (id_visita,))
        elif tipo_evaluacion == "ventilacion":
            if ventilacion_data:
                _upsert_ev_ventilacion(db, id_visita, ventilacion_data)
            db.cursor.execute("DELETE FROM ev_confort WHERE visita_id = %s", (id_visita,))
        else:
            db.cursor.execute("DELETE FROM ev_confort WHERE visita_id = %s", (id_visita,))
            db.cursor.execute("DELETE FROM ev_ventilacion WHERE visita_id = %s", (id_visita,))

        db.connection.commit()
        return True

    except Exception as e:
        logging.error(f"Error al actualizar la visita: {e}")
        if db.connection:
            db.connection.rollback()
        return False
    finally:
        db.close()


def guardar_visita_cierre(id_visita, dato_cierre, tipo_evaluacion):
    """Almacena la información de cierre de una visita según su tipo de evaluación."""
    db = MySQLDatabaseManager()
    try:
        comentario = ""
        ver_tbs_fin = None
        ver_tbh_fin = None
        ver_tg_fin = None

        if isinstance(dato_cierre, dict):
            comentario = dato_cierre.get("note_visita", "")
            ver_tbs_fin = dato_cierre.get("ver_tbs_fin")
            ver_tbh_fin = dato_cierre.get("ver_tbh_fin")
            ver_tg_fin = dato_cierre.get("ver_tg_fin")
        elif isinstance(dato_cierre, (list, tuple)):
            if len(dato_cierre) > 0:
                ver_tbs_fin = dato_cierre[0]
            if len(dato_cierre) > 1:
                ver_tbh_fin = dato_cierre[1]
            if len(dato_cierre) > 2:
                ver_tg_fin = dato_cierre[2]
            if len(dato_cierre) > 3:
                comentario = dato_cierre[3]
        elif dato_cierre is not None:
            comentario = str(dato_cierre)

        query_visita = "UPDATE visitas SET note_visita = %s WHERE id_visita = %s"
        db.cursor.execute(query_visita, (comentario, id_visita))

        if tipo_evaluacion == "confort":
            query_confort = """
                INSERT INTO ev_confort (
                    visita_id,
                    ver_tbs_fin,
                    ver_tbh_fin,
                    ver_tg_fin
                ) VALUES (%s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE
                    ver_tbs_fin = VALUES(ver_tbs_fin),
                    ver_tbh_fin = VALUES(ver_tbh_fin),
                    ver_tg_fin = VALUES(ver_tg_fin)
            """

            db.cursor.execute(
                query_confort,
                (
                    id_visita,
                    ver_tbs_fin,
                    ver_tbh_fin,
                    ver_tg_fin,
                ),
            )

        db.connection.commit()
        return True
    except Exception as e:
        logging.error("Error al actualizar el cierre de la visita: %s", e)
        if db.connection:
            db.connection.rollback()
        return False
    finally:
        db.close()


def guardar_equipos_ventilacion(visita_id, ventilacion_data):
    """Guarda los equipos asociados a una visita de ventilación."""

    db = MySQLDatabaseManager()
    try:
        _upsert_ev_ventilacion(db, visita_id, ventilacion_data or {})
        db.connection.commit()
        return True
    except Exception as e:
        logging.error("Error al guardar equipos de ventilación: %s", e)
        if db.connection:
            db.connection.rollback()
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


def obtener_areas_ventilacion_por_visita(id_visita):
    db = MySQLDatabaseManager()
    try:
        query = """
            SELECT *
            FROM v_areas
            WHERE visita_id = %s
            ORDER BY nombre_area
        """
        db.cursor.execute(query, (id_visita,))
        return db.cursor.fetchall()
    finally:
        db.close()


def generar_siguiente_area_id(visita_id):
    """Obtiene un identificador incremental para registrar una nueva área."""

    db = MySQLDatabaseManager()
    try:
        query = """
            SELECT MAX(CAST(area_id AS UNSIGNED))
            FROM v_areas
            WHERE visita_id = %s AND area_id REGEXP '^[0-9]+$'
        """
        db.cursor.execute(query, (visita_id,))
        resultado = db.cursor.fetchone()
        maximo = resultado[0] if resultado else None

        if maximo is None:
            return "1"

        try:
            siguiente = int(maximo) + 1
        except (TypeError, ValueError):
            logging.warning(
                "No fue posible interpretar el área máxima existente (%s); se utilizará 1.",
                maximo,
            )
            return "1"

        return str(siguiente)
    except Exception as error:
        logging.error("Error al generar un nuevo identificador de área: %s", error)
        return str(int(time.time()))
    finally:
        db.close()


def insertar_area_ventilacion(area_data):
    db = MySQLDatabaseManager()
    try:
        query = """
            INSERT INTO v_areas (
                area_id, visita_id, centro_id, codigo_area, nombre_area, uso,
                piso_nivel, largo_m, ancho_m, alto_m, volumen_m3, aforo_permitido,
                m3_porpersona, m3_porpersona_594, m3_porpersona_cumple,
                caudal_inyeccion_total, caudal_extraccion_total, m3_porpersona_hora,
                m3_porpersona_hora_594, m3_porpersona_hora_cumple,
                recambio_hora_594_min, recambio_hora_594_max, recambio_hora,
                recambio_hora_cumple, ocupacion_habitual, ventilacion_tipo,
                ventilacion_sistema, ventilacion_estado, aberturas, croquis_url,
                observaciones
            ) VALUES (
                %(area_id)s, %(visita_id)s, %(centro_id)s, %(codigo_area)s,
                %(nombre_area)s, %(uso)s, %(piso_nivel)s, %(largo_m)s,
                %(ancho_m)s, %(alto_m)s, %(volumen_m3)s, %(aforo_permitido)s,
                %(m3_porpersona)s, %(m3_porpersona_594)s, %(m3_porpersona_cumple)s,
                %(caudal_inyeccion_total)s, %(caudal_extraccion_total)s,
                %(m3_porpersona_hora)s, %(m3_porpersona_hora_594)s,
                %(m3_porpersona_hora_cumple)s, %(recambio_hora_594_min)s,
                %(recambio_hora_594_max)s, %(recambio_hora)s,
                %(recambio_hora_cumple)s, %(ocupacion_habitual)s,
                %(ventilacion_tipo)s, %(ventilacion_sistema)s,
                %(ventilacion_estado)s, %(aberturas)s, %(croquis_url)s,
                %(observaciones)s
            )
            ON DUPLICATE KEY UPDATE
                visita_id = VALUES(visita_id),
                centro_id = VALUES(centro_id),
                codigo_area = VALUES(codigo_area),
                nombre_area = VALUES(nombre_area),
                uso = VALUES(uso),
                piso_nivel = VALUES(piso_nivel),
                largo_m = VALUES(largo_m),
                ancho_m = VALUES(ancho_m),
                alto_m = VALUES(alto_m),
                volumen_m3 = VALUES(volumen_m3),
                aforo_permitido = VALUES(aforo_permitido),
                m3_porpersona = VALUES(m3_porpersona),
                m3_porpersona_594 = VALUES(m3_porpersona_594),
                m3_porpersona_cumple = VALUES(m3_porpersona_cumple),
                caudal_inyeccion_total = VALUES(caudal_inyeccion_total),
                caudal_extraccion_total = VALUES(caudal_extraccion_total),
                m3_porpersona_hora = VALUES(m3_porpersona_hora),
                m3_porpersona_hora_594 = VALUES(m3_porpersona_hora_594),
                m3_porpersona_hora_cumple = VALUES(m3_porpersona_hora_cumple),
                recambio_hora_594_min = VALUES(recambio_hora_594_min),
                recambio_hora_594_max = VALUES(recambio_hora_594_max),
                recambio_hora = VALUES(recambio_hora),
                recambio_hora_cumple = VALUES(recambio_hora_cumple),
                ocupacion_habitual = VALUES(ocupacion_habitual),
                ventilacion_tipo = VALUES(ventilacion_tipo),
                ventilacion_sistema = VALUES(ventilacion_sistema),
                ventilacion_estado = VALUES(ventilacion_estado),
                aberturas = VALUES(aberturas),
                croquis_url = VALUES(croquis_url),
                observaciones = VALUES(observaciones)
        """

        db.cursor.execute(query, area_data)
        db.connection.commit()
        return True
    except Exception as e:
        logging.error("Error al insertar o actualizar el área de ventilación: %s", e)
        if db.connection:
            db.connection.rollback()
        return False
    finally:
        db.close()


def obtener_puntos_ventilacion_por_area(area_id):
    db = MySQLDatabaseManager()
    try:
        query = """
            SELECT *
            FROM v_puntos_medicion
            WHERE area_id = %s
            ORDER BY codigo_punto
        """
        db.cursor.execute(query, (area_id,))
        return db.cursor.fetchall()
    finally:
        db.close()


def obtener_puntos_ventilacion_por_visita(id_visita):
    db = MySQLDatabaseManager()
    try:
        query = """
            SELECT *
            FROM v_puntos_medicion
            WHERE evaluacion_id = %s
            ORDER BY area_id, codigo_punto
        """
        db.cursor.execute(query, (id_visita,))
        return db.cursor.fetchall()
    finally:
        db.close()


def get_areas_ventilacion_df(id_visita):
    """Obtiene las áreas de ventilación asociadas a una visita como DataFrame."""
    registros = obtener_areas_ventilacion_por_visita(id_visita) or []
    df = pd.DataFrame(registros)
    return df


def get_puntos_ventilacion_df(id_visita):
    """Obtiene los puntos de medición de ventilación asociados a una visita como DataFrame."""
    registros = obtener_puntos_ventilacion_por_visita(id_visita) or []
    df = pd.DataFrame(registros)
    return df


def insertar_punto_ventilacion(punto_data):
    db = MySQLDatabaseManager()
    try:
        query = """
            INSERT INTO v_puntos_medicion (
                punto_id, evaluacion_id, area_id, codigo_punto, tipo_punto,
                ubicacion_detalle, altura_m, distancia_fuente_m, conducto_largo_cm,
                conducto_ancho_cm, conducto_diametro, seccion_conducto_cm2,
                medicion_caudal_1, medicion_caudal_2, medicion_caudal_3,
                medicion_caudal_4, medicion_caudal_5, medicion_caudal_p, caudal,
                fecha_hora, condiciones_ocupacion, puertas_ventanas_abiertas,
                temperatura_c, humedad_relativa_pct, croquis_url, observaciones
            ) VALUES (
                %(punto_id)s, %(evaluacion_id)s, %(area_id)s, %(codigo_punto)s,
                %(tipo_punto)s, %(ubicacion_detalle)s, %(altura_m)s,
                %(distancia_fuente_m)s, %(conducto_largo_cm)s, %(conducto_ancho_cm)s,
                %(conducto_diametro)s, %(seccion_conducto_cm2)s,
                %(medicion_caudal_1)s, %(medicion_caudal_2)s, %(medicion_caudal_3)s,
                %(medicion_caudal_4)s, %(medicion_caudal_5)s, %(medicion_caudal_p)s,
                %(caudal)s, %(fecha_hora)s, %(condiciones_ocupacion)s,
                %(puertas_ventanas_abiertas)s, %(temperatura_c)s,
                %(humedad_relativa_pct)s, %(croquis_url)s, %(observaciones)s
            )
            ON DUPLICATE KEY UPDATE
                codigo_punto = VALUES(codigo_punto),
                tipo_punto = VALUES(tipo_punto),
                ubicacion_detalle = VALUES(ubicacion_detalle),
                altura_m = VALUES(altura_m),
                distancia_fuente_m = VALUES(distancia_fuente_m),
                conducto_largo_cm = VALUES(conducto_largo_cm),
                conducto_ancho_cm = VALUES(conducto_ancho_cm),
                conducto_diametro = VALUES(conducto_diametro),
                seccion_conducto_cm2 = VALUES(seccion_conducto_cm2),
                medicion_caudal_1 = VALUES(medicion_caudal_1),
                medicion_caudal_2 = VALUES(medicion_caudal_2),
                medicion_caudal_3 = VALUES(medicion_caudal_3),
                medicion_caudal_4 = VALUES(medicion_caudal_4),
                medicion_caudal_5 = VALUES(medicion_caudal_5),
                medicion_caudal_p = VALUES(medicion_caudal_p),
                caudal = VALUES(caudal),
                fecha_hora = VALUES(fecha_hora),
                condiciones_ocupacion = VALUES(condiciones_ocupacion),
                puertas_ventanas_abiertas = VALUES(puertas_ventanas_abiertas),
                temperatura_c = VALUES(temperatura_c),
                humedad_relativa_pct = VALUES(humedad_relativa_pct),
                croquis_url = VALUES(croquis_url),
                observaciones = VALUES(observaciones)
        """

        db.cursor.execute(query, punto_data)
        db.connection.commit()
        return True
    except Exception as e:
        logging.error("Error al insertar o actualizar el punto de medición: %s", e)
        if db.connection:
            db.connection.rollback()
        return False
    finally:
        db.close()


def recalcular_totales_area_ventilacion(area_id):
    db = MySQLDatabaseManager()
    try:
        query_area = "SELECT * FROM v_areas WHERE area_id = %s"
        db.cursor.execute(query_area, (area_id,))
        area = db.cursor.fetchone()
        if not area:
            return False

        query_totales = """
            SELECT
                SUM(CASE WHEN tipo_punto = 'Inyeccion' THEN caudal ELSE 0 END) AS total_inyeccion,
                SUM(CASE WHEN tipo_punto = 'Extraccion' THEN caudal ELSE 0 END) AS total_extraccion
            FROM v_puntos_medicion
            WHERE area_id = %s
        """
        db.cursor.execute(query_totales, (area_id,))
        totales = db.cursor.fetchone() or {}

        def _to_float(valor):
            if valor is None:
                return None
            try:
                return float(valor)
            except (TypeError, ValueError):
                return None

        total_inyeccion = _to_float(totales.get("total_inyeccion")) or 0.0
        total_extraccion = _to_float(totales.get("total_extraccion")) or 0.0
        aforo = _to_float(area.get("aforo_permitido")) or 0.0
        volumen = _to_float(area.get("volumen_m3")) or 0.0
        m3_pp_ref = _to_float(area.get("m3_porpersona_594")) or 10.0
        m3_pp_hora_ref = _to_float(area.get("m3_porpersona_hora_594")) or 20.0
        recambio_min = _to_float(area.get("recambio_hora_594_min")) or 0.0
        recambio_max = _to_float(area.get("recambio_hora_594_max")) or 0.0

        m3_pp = (volumen / aforo) if aforo > 0 else None
        m3_pp_cumple = 1 if (m3_pp is not None and m3_pp >= m3_pp_ref) else 0

        max_caudal = max(total_inyeccion, total_extraccion)
        m3_pp_hora = (max_caudal / aforo) if aforo > 0 else None
        m3_pp_hora_cumple = 1 if (m3_pp_hora is not None and m3_pp_hora >= m3_pp_hora_ref) else 0

        recambio_hora = (max_caudal / volumen) if volumen > 0 else None
        recambio_cumple = 1 if (
            recambio_hora is not None and recambio_hora >= recambio_min and (
                recambio_max == 0 or recambio_hora <= recambio_max
            )
        ) else 0

        query_update = """
            UPDATE v_areas
            SET caudal_inyeccion_total = %s,
                caudal_extraccion_total = %s,
                m3_porpersona = %s,
                m3_porpersona_cumple = %s,
                m3_porpersona_hora = %s,
                m3_porpersona_hora_cumple = %s,
                recambio_hora = %s,
                recambio_hora_cumple = %s
            WHERE area_id = %s
        """

        db.cursor.execute(
            query_update,
            (
                total_inyeccion,
                total_extraccion,
                m3_pp,
                m3_pp_cumple,
                m3_pp_hora,
                m3_pp_hora_cumple,
                recambio_hora,
                recambio_cumple,
                area_id,
            ),
        )
        db.connection.commit()
        return True
    except Exception as e:
        logging.error("Error al recalcular totales del área de ventilación: %s", e)
        if db.connection:
            db.connection.rollback()
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
