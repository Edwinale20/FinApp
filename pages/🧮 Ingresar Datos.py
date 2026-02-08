import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO

st.title("🤖 Resumen semanal")
st.markdown("La finalidad de esta segunda parte, es ingresar todo tip de movimiento registrado en mi día a día", unsafe_allow_html=True)

# -----------------------------------------------------------------------------------------------------------------------------

cfg = st.secrets["onedrive"]
CLIENT_ID = cfg["client_id"]
CLIENT_SECRET = cfg["client_secret"]
REFRESH_TOKEN = cfg["refresh_token"]
REDIRECT_URI = cfg["redirect_uri"]

def get_access_token():
    url = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": REFRESH_TOKEN,
        "grant_type": "refresh_token",
        "redirect_uri": REDIRECT_URI,
        "scope": "Files.ReadWrite Files.Read.All User.Read offline_access"
    }
    r = requests.post(url, data=data)
    return r.json()

@st.cache_data
def list_excel_files(access_token):
    url = "https://graph.microsoft.com/v1.0/me/drive/root:/FinApp:/children"
    headers = {"Authorization": f"Bearer {access_token}"}
    r = requests.get(url, headers=headers).json()

    return [f for f in r.get("value", []) if f["name"].lower().endswith(".xlsx")]

@st.cache_data
def download_excel_df(access_token, file_id):
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
    headers = {"Authorization": f"Bearer {access_token}"}
    content = requests.get(url, headers=headers).content
    return pd.read_excel(io.BytesIO(content))


# -----------------------------------------------------------------------------------------------------------------------------

token = get_access_token()

if "access_token" not in token:
    st.error("❌ Error obteniendo access_token")
    st.code(token)
else:
    access_token = token["access_token"]

    files = list_excel_files(access_token)

    # Buscar el archivo específico llamado Tracking.xlsx
    tracking_file = next((f for f in files if f["name"] == "Tracking.xlsx"), None)
    
    if tracking_file:
        df_tracking = download_excel_df(access_token, tracking_file["id"])
        #df_tracking2 = download_excel_df(access_token, tracking_file["id"], sheet_name="Deudas")  # segunda hoja
        st.success("✅ Archivo 'Tracking.xlsx' cargado correctamente.")
    else:
        st.error("❌ No se encontró el archivo 'Tracking.xlsx' en la carpeta FinApp.")

# -----------------------------------------------------------------------------------------------------------------------------

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

