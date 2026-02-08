import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO

st.title("🤖 Resumen semanal")
st.markdown("La finalidad de esta segunda parte, es ingresar todo tip de movimiento registrado en mi día a día", unsafe_allow_html=True)



TRACKING_PATH = Path(r"C:\Users\omen0\OneDrive\Documentos\OneDrive\FinApp\Tracking.xlsx")

if not TRACKING_PATH.exists():
    st.error(f"❌ No se encontró el archivo: {TRACKING_PATH}")
    st.stop()

# Leer SOLO la hoja Movimientos
df_tracking = pd.read_excel(TRACKING_PATH, sheet_name="Registro")

# Inputs
Nombre = st.text_input("🖋️ Ingresa la Descripción:")
Cantidad = st.number_input("💲Ingresa el monto:", min_value=0.0, step=1.0)
Categoria = st.text_input("🍻 Ingresa la categoría:")
fecha = st.date_input("🗓️ Selecciona la fecha:")

Submit = st.button("Ingresar")

if Submit:
    new_row = pd.DataFrame([{
        "Fecha": pd.to_datetime(fecha),
        "Categoría": Categoria,
        "Descripción": Nombre,
        "Monto": float(Cantidad)
    }])

    df_tracking = pd.concat([df_tracking, new_row], ignore_index=True)

    # Guardar de vuelta AL MISMO archivo (sobrescribe solo Movimientos)
    with pd.ExcelWriter(TRACKING_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_tracking.to_excel(writer, index=False, sheet_name="Registro")

    st.success("✅ Guardado en Tracking.xlsx (hoja Registro)")
    st.rerun()

