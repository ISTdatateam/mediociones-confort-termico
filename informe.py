import streamlit as st
import pandas as pd
import io
import zipfile
from utils.helpers import get_ct, get_visita, get_mediciones, get_equipos, get_all_cuvs_with_visits, get_visitas_por_cuv
from utils.doc_utils import generar_informe_en_word
from db.mysql_utils import MySQLDatabaseManager


def generar_informe(cuv, id_visita):
    """Genera un informe basado en el CUV y devuelve el archivo en formato BytesIO."""
    df_centro = pd.DataFrame(get_ct(cuv))
    df_visitas = get_visita(id_visita)
    df_mediciones = get_mediciones(id_visita) if not df_visitas.empty else pd.DataFrame()
    df_equipos = get_equipos()

    if df_centro.empty or df_visitas.empty:
        st.error(f"No se encontró suficiente información para generar el informe del CUV {cuv}.")
        return None

    informe_docx = generar_informe_en_word(df_centro, df_visitas, df_mediciones, df_equipos)
    return informe_docx

def generar_informe_unit(cuv):
    """
    Genera un informe basado en el CUV, usando automáticamente la visita más reciente.
    """
    try:
        cuv = int(cuv)
    except ValueError:
        st.error("El CUV ingresado no es válido.")
        return None

    df_centro = pd.DataFrame(get_ct(cuv))
    df_visitas = get_visitas_por_cuv(cuv)

    if df_centro.empty or df_visitas.empty:
        st.error(f"No se encontró suficiente información para generar el informe del CUV {cuv}.")
        return None

    visita_reciente = df_visitas.iloc[0]
    id_visita = int(visita_reciente["id_visita"])

    df_mediciones = get_mediciones(id_visita)
    df_equipos = get_equipos()

    informe_docx = generar_informe_en_word(df_centro, pd.DataFrame([visita_reciente]), df_mediciones, df_equipos)
    return informe_docx


def get_visita_por_cuv(cuv):
    db = MySQLDatabaseManager()
    try:
        query = "SELECT * FROM visitas WHERE cuv_visita = %s ORDER BY fecha_visita DESC"
        db.cursor.execute(query, (int(cuv),))
        resultados = db.cursor.fetchall()
        return pd.DataFrame(resultados)
    finally:
        db.close()


def generar_informes_para_todos():
    cuvs = get_all_cuvs_with_visits()

    if cuvs.empty:
        print("No hay CUVs con visitas registradas.")
        return None

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for row in cuvs.itertuples():
            cuv = str(row.cuv_visita)
            print(f"🔍 Procesando CUV {cuv}...")

            try:
                df_ct = pd.DataFrame(get_ct(cuv))
                df_visita = get_visita(int(cuv))  # <-- Asegúrate que esto retorna visitas, no solo el cuv

                if df_ct.empty:
                    print(f"CUV {cuv}: sin datos de centro.")
                    continue
                if df_visita.empty:
                    print(f"CUV {cuv}: sin datos de visita.")
                    continue

                visita_id = df_visita.iloc[0]["id_visita"]
                df_mediciones = get_mediciones(visita_id)
                df_equipos = get_equipos()

                if df_mediciones.empty:
                    print(f"CUV {cuv}: sin mediciones.")
                    continue

                doc_bytes = generar_informe_en_word(df_ct, df_visita, df_mediciones, df_equipos)

                if doc_bytes:
                    zip_file.writestr(f"informe_CUV_{cuv}.docx", doc_bytes.getvalue())
                    print(f"Informe agregado al ZIP para CUV {cuv}")
                else:
                    print(f"No se generó el documento para CUV {cuv}")

            except Exception as e:
                print(f"Error generando informe para CUV {cuv}: {str(e)}")

    zip_buffer.seek(0)
    return zip_buffer

def generar_informes_masivos():
    """Genera informes para todos los CUVs con visitas registradas y los empaqueta en un archivo ZIP."""
    cuvs = get_all_cuvs_with_visits()
    total = len(cuvs)

    if total == 0:
        st.warning("No hay CUVs con visitas registradas.")
        return None

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        progress_bar = st.progress(0)

        for i, row in cuvs.iterrows():
            cuv = str(row["cuv_visita"])  # Asegura que sea str para consultas

            try:
                # Obtener datos del centro de trabajo
                df_centro = pd.DataFrame(get_ct(cuv))
                if df_centro.empty:
                    st.warning(f"CUV {cuv}: sin datos de centro.")
                    continue

                # Obtener visita asociada
                df_visitas = get_visita(cuv)  # O get_visita(int(cuv)) si esperas ID
                if df_visitas.empty:
                    st.warning(f"CUV {cuv}: sin datos de visita.")
                    continue

                visita_id = int(df_visitas.iloc[0]["id_visita"])

                # Obtener mediciones y equipos
                df_mediciones = get_mediciones(visita_id)
                df_equipos = get_equipos()

                # Generar documento Word
                doc_bytes = generar_informe_en_word(df_centro, df_visitas, df_mediciones, df_equipos)

                # Guardar el archivo en el ZIP
                zip_file.writestr(f"informe_{cuv}.docx", doc_bytes.getvalue())

            except Exception as e:
                st.error(f"Error generando informe para CUV {cuv}: {str(e)}")

            progress_bar.progress((i + 1) / total)

    zip_buffer.seek(0)
    return zip_buffer



def main():
    st.header("Informes Confort Térmico")
    st.write("Versión 4.0 (Generación Automática)")
    st.write("Bienvenido Rodrigo... (usuario)")

    # Sección para generación manual
    st.subheader("Generar Informe Individual")
    input_cuv = st.text_input("Ingresa el CUV: ej. 178050")

    if st.button("Buscar y Generar Informe"):
        informe = generar_informe_unit(input_cuv)
        if informe:
            st.success("Informe generado correctamente.")
            st.download_button("Descargar Informe", data=informe, file_name=f"informe_{input_cuv}.docx")



    if st.button("Generar informes para todos los CUVs 2"):
        zip_file = generar_informes_para_todos()
        if zip_file:
            st.download_button(
                label="Descargar todos los informes",
                data=zip_file,
                file_name="informes_confort_termico.zip",
                mime="application/zip"
            )


# Permite que `supermain.py` se pueda ejecutar como script independiente
if __name__ == "__main__":
    main()
