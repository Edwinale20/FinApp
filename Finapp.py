import streamlit as st
import requests
import pandas as pd
import openpyxl
import io
import plotly.express as px

st.set_page_config(page_title="FinApp", page_icon="💸", layout="wide")

st.title("💸 FinApp de Pepe")
st.markdown("✅ Datos en tiempo real", unsafe_allow_html=True)
st.subheader("📋 KPI´s principales")

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



# Calcular métricas
df_tracking["Fecha"] = pd.to_datetime(df_tracking["Fecha"], dayfirst=True, errors='coerce')
mes_actual = df_tracking["Fecha"].dt.month_name().mode()[0]


total_ingresos = df_tracking[df_tracking["Concepto"] == "Ingreso"]["Monto"].sum()
total_gasto_fijo = df_tracking[df_tracking["Concepto"] == "Gasto"]["Monto"].sum()
balance = total_ingresos - total_gasto_fijo 


c7, c8, c9 = st.columns([4,3,4])
with c7:
    total_ingresos = df_tracking[df_tracking["Concepto"] == "Ingreso"]["Monto"].sum()    
    st.metric(label="📈 Ingresos", value=f"${total_ingresos:,.0f}")

with c8:
    total_gasto_fijo = df_tracking[df_tracking["Concepto"] == "Gasto"]["Monto"].sum()    
    st.metric(label="📉 Gastos", value=f"${total_gasto_fijo:,.0f}")

with c9: 
    st.metric(label="💵 Balance", value=f"${balance:,.0f}")



st.sidebar.title("Filtros 🔠")


# Paso 1: Crear una lista de opciones para el filtro, incluyendo "Ninguno"
Opcion_Categoría  = ['Ninguno'] + list(df_tracking['Categoría'].unique())
Categoria = st.sidebar.selectbox('Seleccione la Categoría', Opcion_Categoría)


df_tracking["Mes"] = pd.to_datetime(df_tracking["Fecha"], dayfirst=True, errors="coerce").dt.to_period("M").astype(str)
Mes = st.sidebar.selectbox("Seleccione el mes", ["Ninguno"] + sorted(df_tracking["Mes"].dropna().unique()))


if Categoria != 'Ninguno':
    df_tracking = df_tracking[df_tracking['Categoría'] == Categoria]

df_tracking["Mes"] = pd.to_datetime(df_tracking["Fecha"], dayfirst=True, errors="coerce").dt.to_period("M").astype(str)

if Mes != "Ninguno": df_tracking = df_tracking[df_tracking["Mes"] == Mes]

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

    fig.update_layout(title_font=dict(size=20))

    return fig


# -----------------------------------------------------------------------------------------------------------------------------
#Ejemplo de una buena grafica de barras!!!
def figura2():
    df_tracking["Fecha"] = pd.to_datetime(df_tracking["Fecha"], dayfirst=True, errors="coerce")

    df_gastos = df_tracking[df_tracking["Concepto"] == "Gasto"].copy()
    df_gastos["Mes"] = df_gastos["Fecha"].dt.to_period("M").astype(str)

    df_mes_cat = df_gastos.groupby(["Mes", "Categoría"])["Monto"].sum().reset_index()
    
    custom_colors = [
        '#00712D', '#FF9800', '#000080', '#FF6347', '#000000',
        '#FFD700', '#008080', '#CD5C5C', '#FF7F50', '#006400',
        '#8B0000', '#FFDEAD', '#ADFF2F', '#2F4F4F', '#33A85C']
    

    fig = px.bar(
        df_mes_cat,
        x="Mes",
        y="Monto",
        title="Gastos por mes por categoría",
        text="Monto",
        color='Categoría',  # Usamos DESCRIPCIÓN en lugar de ARTICULO
        color_discrete_sequence = ['#007074', '#FFBF00', '#9694FF', '#222831', '#004225', '#1230AE', '#8D0B41', '#522258', 
         '#1F7D53', '#EB5B00', '#0D1282', '#09122C', '#ADFF2F', '#2F4F4F', "#7C00FE", "#D10363", "#16404D"],)


    fig.update_layout(title_font=dict(size=20))

    fig.update_traces(
        texttemplate='$%{text:,.0f}', 
        textposition='outside',
        hovertemplate='Mes: %{x}<br>Gasto: $%{y:,.0f}<extra>%{fullData.name}</extra>' )  

    return fig


# -----------------------------------------------------------------------------------------------------------------------------
def figura3():
    df_tracking["Fecha"] = pd.to_datetime(df_tracking["Fecha"], dayfirst=True, errors="coerce")

    df = df_tracking[df_tracking["Concepto"].isin(["Ingreso", "Gasto"])].copy()
    df["Mes"] = df["Fecha"].dt.to_period("M").astype(str)

    df_mes = df.groupby(["Mes", "Concepto"])["Monto"].sum().reset_index()

    fig = px.bar(
        df_mes,
        x="Mes",
        y="Monto",
        color="Concepto",
        title="Balance mensual",
        text="Monto",
        barmode="group"   # usa "stack" si quieres apilado
    )

    fig.update_layout(title_font=dict(size=20))

    fig.update_traces(
        texttemplate='$%{text:,.0f}',
        textposition='outside',
        hovertemplate='Mes: %{x}<br>%{fullData.name}: $%{y:,.0f}<extra></extra>'
    )

    return fig


def figura4():

    df = df_tracking[df_tracking["Concepto"] == "Gasto"].copy()
    #df["Mes"] = df["Fecha"].dt.to_period("M").astype(str)

    df_mes = df.groupby(["Categoría", "Concepto"])["Monto"].sum().reset_index()

    fig = px.bar(
        df_mes,
        x="Categoría",
        y="Monto",
        color="Concepto",
        title="Seguimiento de deudas",
        text="Monto",
        barmode="group"   # usa "stack" si quieres apilado
    )

    fig.update_layout(title_font=dict(size=20))

    fig.update_traces(
        texttemplate='$%{text:,.0f}',
        textposition='outside',
        hovertemplate='Mes: %{x}<br>%{fullData.name}: $%{y:,.0f}<extra></extra>'
    )

    return fig
# -----------------------------------------------------------------------------------------------------------------------------

#Listado de graficas
figura1_grafica = figura1()
figura2_grafica = figura2()
figura3_grafica = figura3()
figura4_grafica = figura4()
#figura6_grafica = figura6()

#Barra 1
st.divider()
st.subheader("_Seguimiento de_ :green[Finanzas personales]")
c3, c2, c1 = st.columns([4, 3, 4])
with c1:
    st.plotly_chart(figura1_grafica, use_container_width=True)

with c2:
    st.plotly_chart(figura2_grafica, use_container_width=True)

with c3:
    st.plotly_chart(figura3_grafica, use_container_width=True)

#Barra 2
st.divider()
st.subheader("_Seguimiento de_ :blue[metas] :sunglasses:")
c4, c5, c6 = st.columns([4, 3, 4])
with c4:
    st.plotly_chart(figura4_grafica, use_container_width=True)

