import streamlit as st
from pythermalcomfort.models import pmv_ppd_iso
from pythermalcomfort.utilities import v_relative

# https://pythermalcomfort.readthedocs.io/en/latest/documentation/models.html#predicted-mean-vote-pmv-and-predicted-percentage-of-dissatisfied-ppd


# Título de la aplicación
st.title("Calculadora de PMV y PPD")
st.write("Ajusta los parámetros y presiona **Calcular** para ver los resultados.")

# Entradas de usuario con valores por defecto (variables de valor fijo)
tdb = st.number_input("Temperatura de bulbo seco (°C):", value=30.0)
tr = st.number_input("Temperatura radiante (°C):", value=30.0)
rh = st.number_input("Humedad relativa (%):", value=32.0)
v = st.number_input("Velocidad del aire (m/s):", value=0.8)
met = st.number_input("Tasa metabólica (met):", value=1.1)
clo_dynamic = st.number_input("Aislamiento de la ropa (clo):", value=0.5)

# Botón de cálculo
if st.button("Calcular"):
    # Calcular la velocidad del aire relativa
    v_r = v_relative(v=v, met=met)

    # Calcular PMV y PPD usando el modelo especificado
    results = pmv_ppd_iso(
        tdb=tdb,
        tr=tr,
        vr=v,
        rh=rh,
        met=met,
        clo=clo_dynamic,
        model="7730-2005",
        limit_inputs=False,
        round_output=True
    )

    # Mostrar los resultados
    st.subheader("Resultados")
    st.write(f"**PMV:** {results.pmv}")
    st.write(f"**PPD:** {results.ppd}")
