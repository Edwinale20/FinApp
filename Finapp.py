import streamlit as st
import requests
import pandas as pd
import openpyxl
import io
import plotly.express as px

st.set_page_config(page_title="FinApp", page_icon="💸")

st.title("💸 FinApp de Pepe")
st.markdown("✅ Datos en tiempo real", unsafe_allow_html=True)
st.markdown("🧮 KPI´s principales", unsafe_allow_html=True)

# ---------------- CONFIG ----------------
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

@st.cache_data
def venta(venta_semanal):
    concat_venta = pd.DataFrame()

    for df2 in venta_semanal:
        df2 = df2.loc[:, ~df2.columns.str.contains('^Unnamed')]

        if "Semana Contable" not in df2.columns:
            continue

        df2["Semana Contable"] = df2["Semana Contable"].astype(str)
        columnas_a_eliminar = ['Metrics']
        df2 = df2.drop(columns=[col for col in columnas_a_eliminar if col in df2.columns], errors='ignore')
        concat_venta = pd.concat([concat_venta, df2], ignore_index=True)

    return concat_venta


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
        st.success("✅ Archivo 'Tracking.xlsx' cargado correctamente.")
    else:
        st.error("❌ No se encontró el archivo 'Tracking.xlsx' en la carpeta FinApp.")



# Calcular métricas
df_tracking["Fecha"] = pd.to_datetime(df_tracking["Fecha"], dayfirst=True, errors='coerce')
mes_actual = df_tracking["Fecha"].dt.month_name().mode()[0]


total_ingresos = df_tracking[df_tracking["Concepto"] == "Ingreso"]["Monto"].sum()
total_gasto_fijo = df_tracking[df_tracking["Concepto"] == "Gasto"]["Monto"].sum()
balance = total_ingresos - total_gasto_fijo 


c7, c8, c9 = st.columns([4,3,4])
with c7:
    total_ingresos = df_tracking[df_tracking["Concepto"] == "Ingreso"]["Monto"].sum()    
    st.metric(label="🚨 Ingresos", value=f"${total_ingresos:,.0f}")

with c8:
    total_gasto_fijo = df_tracking[df_tracking["Concepto"] == "Gasto"]["Monto"].sum()    
    st.metric(label="🚨 Gastos", value=f"${total_gasto_fijo:,.0f}")

with c9: 
    st.metric(label="🚨 Balance", value=f"${balance:,.0f}")

st.divider()
st.write("📊 **Base consolidada:**")

# -----------------------------------------------------------------------------------------------------------------------------


def figura1():
    df_tracking["Fecha"] = pd.to_datetime(df_tracking["Fecha"], dayfirst=True, errors="coerce")

    df_filtrado = df_tracking[df_tracking["Concepto"].isin(["Ingreso", "Gasto"])]

    df_diario = df_filtrado.groupby(["Fecha", "Concepto"]).sum().reset_index()

    fig = px.line(
        df_diario,
        x="Fecha",
        y="Monto",
        color="Concepto",
        title="Balance diario: ingresos vs gastos",
        markers=True
    )

    fig.update_layout(title_x=0.5)
    return fig


# -----------------------------------------------------------------------------------------------------------------------------

def figura2():
    df_tracking["Fecha"] = pd.to_datetime(df_tracking["Fecha"], dayfirst=True, errors="coerce")

    df_filtrado = df_tracking[df_tracking["Concepto"].isin(["Ingreso", "Gasto"])]

    df_diario = df_filtrado.groupby(["Fecha", "Concepto"]).sum().reset_index()

    fig = px.line(
        df_diario,
        x="Fecha",
        y="Monto",
        color="Concepto",
        title="Balance diario: ingresos vs gastos",
        markers=True
    )

    fig.update_layout(title_x=0.5)
    return fig


# -----------------------------------------------------------------------------------------------------------------------------

def figura3():
    # Leer datos
    df_tracking
    df_grafica3 = df_tracking.groupby(["Concepto"])["Monto"].sum().reset_index()

    # Crear gráfica
    fig = px.bar(
        df_grafica3,
        x="Concepto",
        y="Monto",
        color="Concepto",
        text="Monto",
        title="Tracking de deudas",
        #labels={'VENTA_PERDIDA_PESOS': 'Venta Perdida en Pesos (M)'},
        #hover_data={'% Venta Perdida': ':.1f'}
    )

    fig.update_layout(
        xaxis_title="Concepto",
        yaxis_title="Monto ($)",
        #title_x=0.5,
        barmode='stack',
        title_font=dict(size=20),
        #height=400,
    )
    
    return fig


figura1_grafica = figura1()
figura2_grafica = figura2()
figura3_grafica = figura3()

c1, c2, c3 = st.columns([4, 3, 4])
with c1:
    st.plotly_chart(figura1_grafica, use_container_width=True)

with c2:
    st.plotly_chart(figura2_grafica, use_container_width=True)

with c3:
    st.plotly_chart(figura3_grafica, use_container_width=True)

