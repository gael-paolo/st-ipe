import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests
from datetime import datetime, timedelta

# =========================================================
# CONFIGURACIÓN
# =========================================================
st.set_page_config(page_title="IPE Tracking", layout="wide", page_icon="🚗")

# =========================================================
# LOGIN
# =========================================================
st.sidebar.title("🔐 Acceso")
clave = st.sidebar.text_input("Ingrese clave", type="password")

ROL = None
if clave == st.secrets["clave_comercial"]:
    ROL = "COMERCIAL"
elif clave == st.secrets["clave_taller"]:
    ROL = "TALLER"
elif clave:
    st.sidebar.error("Clave incorrecta")

if not ROL:
    st.stop()

# =========================================================
# FUNCIONES
# =========================================================
def enviar_telegram(mensaje):
    try:
        token = st.secrets["telegram_token"]
        chat_id = st.secrets["telegram_chat_id"]

        url = f"https://api.telegram.org/bot{token}/sendMessage"
        payload = {"chat_id": chat_id, "text": mensaje, "parse_mode": "HTML"}

        response = requests.post(url, json=payload)
        return response.status_code == 200

    except Exception as e:
        st.error(f"Error Telegram: {e}")
        return False


def enviar_correo_taller(idv, cliente, apv, punto, marca, modelo):
    usuario = st.secrets["smtp_user"]
    password = st.secrets["smtp_pass"]

    destinatario = st.secrets["correo_to"]
    copia = st.secrets["correo_cc"]

    msg = MIMEMultipart()
    msg['Subject'] = f"🔔 NUEVA SOLICITUD - IDV: {idv}"
    msg['From'] = usuario
    msg['To'] = destinatario
    msg['Cc'] = copia

    cuerpo = f"""
Se ha generado una nueva solicitud de IPE.

DETALLES DEL VEHÍCULO:
- IDV: {idv}
- Unidad: {marca} {modelo}

DATOS DE VENTA:
- Cliente: {cliente}
- APV: {apv}
- Sucursal de Entrega: {punto}
- Fecha de registro: {datetime.now().strftime('%d/%m/%Y %H:%M')}

Por favor, proceder a abrir una OT y procesarlo.
Muchas gracias de antemano

RPA Logystics IPE
🤓
"""
    msg.attach(MIMEText(cuerpo, 'plain'))

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(usuario, password)

            destinatarios = [destinatario] + copia.split(",")
            server.sendmail(usuario, destinatarios, msg.as_string())

        return True

    except Exception as e:
        st.error(f"Error Correo: {e}")
        return False


def enviar_correo_confirmacion(apv_email, data):
    usuario = st.secrets["smtp_user"]
    password = st.secrets["smtp_pass"]

    msg = MIMEMultipart()
    msg['Subject'] = f"✅ Confirmación Solicitud IPE - IDV: {data['IDV']}"
    msg['From'] = usuario
    msg['To'] = apv_email

    tabla = f"""
    <table border="1" cellpadding="6" cellspacing="0">
        <tr><th>Campo</th><th>Valor</th></tr>
        <tr><td>IDV</td><td>{data['IDV']}</td></tr>
        <tr><td>Cliente</td><td>{data['Cliente']}</td></tr>
        <tr><td>Vehículo</td><td>{data['Marca']} {data['Modelo']}</td></tr>
        <tr><td>Color</td><td>{data['Color']}</td></tr>
        <tr><td>APV</td><td>{data['APV']}</td></tr>
        <tr><td>Sucursal</td><td>{data['Punto']}</td></tr>
        <tr><td>Fecha Promesa</td><td>{data['Fecha_Promesa']}</td></tr>
        <tr><td>Implementaciones</td><td>{data['Implementaciones']}</td></tr>
        <tr><td>Condiciones</td><td>{data['Condiciones']}</td></tr>
    </table>
    """

    cuerpo = f"""
    <html>
    <body>
        <p>Hola,</p>
        <p>Tu solicitud IPE ha sido registrada correctamente:</p>
        {tabla}
        <br>
        <p>Equipo IPE 🚗</p>
    </body>
    </html>
    """

    msg.attach(MIMEText(cuerpo, 'html'))

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(usuario, password)
            server.sendmail(usuario, apv_email, msg.as_string())

        return True

    except Exception as e:
        st.error(f"Error correo confirmación: {e}")
        return False


def obtener_email_apv(df_maestros, apv):
    fila = df_maestros[df_maestros["APV"] == apv]
    if not fila.empty:
        return fila.iloc[0]["Email"]
    return None


def sumar_dias_habiles(fecha, dias):
    f = fecha
    contador = 0
    while contador < dias:
        f += timedelta(days=1)
        if f.weekday() < 5:
            contador += 1
    return f

# =========================================================
# CONEXIÓN
# =========================================================
conn = st.connection("gsheets", type=GSheetsConnection)
URL_SHEET = st.secrets["spreadsheet"]

# =========================================================
# CARGA DE LISTAS
# =========================================================
@st.cache_data(ttl=300)
def cargar_maestros():
    df = conn.read(spreadsheet=URL_SHEET, worksheet="Maestros")
    df = df.fillna("")
    return df

df_maestros = cargar_maestros()

marcas = sorted(df_maestros["Marca"].unique())
modelos = sorted(df_maestros["Modelo"].unique())
apvs = sorted(df_maestros["APV"].unique())
puntos = sorted(df_maestros["Punto"].unique())

# =========================================================
# HEADER
# =========================================================
st.markdown("## 🚗 Gestión de IPE")
st.divider()

# =========================================================
# TABS
# =========================================================
if ROL == "COMERCIAL":
    tabs = st.tabs(["📝 Solicitudes", "🔍 Consulta de Status"])
else:
    tabs = st.tabs([
        "📝Solicitudes",
        "📊 Carga de Reporte (Supervisor)",
        "🔍 Consulta de Status"
    ])

# =========================================================
# 1. REGISTRO
# =========================================================
with tabs[0]:
    st.header("Formulario de Solicitud")
    st.markdown("⚠️ Presione el botón para enviar")

    with st.form("registro", clear_on_submit=True):

        with st.expander("🚗 Datos del Vehículo", expanded=True):
            col1, col2 = st.columns(2)

            with col1:
                idv = st.text_input("IDV*").upper().strip()
                marca = st.selectbox("Marca", marcas)
                modelo = st.selectbox("Modelo", modelos)

            with col2:
                color = st.text_input("Color")
                apv = st.selectbox("APV", apvs)

        with st.expander("📋 Datos de Entrega", expanded=True):
            col1, col2 = st.columns(2)

            with col1:
                punto = st.selectbox("Sucursal de Entrega", puntos)

                hoy = datetime.now()
                fecha_minima = sumar_dias_habiles(hoy, 10)

                fecha_p = st.date_input(
                    "Fecha Promesa",
                    min_value=fecha_minima.date()
                )

            with col2:
                cliente = st.text_input("Cliente")

        with st.expander("⚙️ Detalles adicionales"):
            impl = st.text_area("Implementaciones")
            cond = st.text_area("Condiciones")

        btn = st.form_submit_button("🚀 Enviar Solicitud")

        if btn:
            if not idv or not cliente:
                st.error("⚠️ Campos obligatorios faltantes")
            else:
                try:
                    data_actual = conn.read(spreadsheet=URL_SHEET, worksheet="Solicitudes")

                    data_actual['IDV'] = (
                        data_actual['IDV']
                        .astype(str)
                        .str.replace('.0', '', regex=False)
                        .str.strip()
                    )

                    if idv in data_actual['IDV'].values:
                        st.error(f"⚠️ El IDV {idv} ya existe")
                    else:
                        nueva = pd.DataFrame([{
                            "IDV": idv,
                            "Marca": marca,
                            "Modelo": modelo,
                            "Color": color,
                            "APV": apv,
                            "Punto": punto,
                            "Fecha_Promesa": str(fecha_p),
                            "Cliente": cliente,
                            "Implementaciones": impl,
                            "Condiciones": cond,
                            "Fecha_Registro": datetime.now().strftime("%Y-%m-%d %H:%M")
                        }])

                        data_nueva = pd.concat([data_actual, nueva], ignore_index=True)
                        conn.update(spreadsheet=URL_SHEET, worksheet="Solicitudes", data=data_nueva)

                        mensaje_html = f"""
<b>🚀 NUEVA SOLICITUD DE ALISTAMIENTO</b>
────────────────────────
<b>🆔 IDV:</b> <code>{idv}</code>
<b>🚗 Vehículo:</b> {marca} {modelo} ({color})
<b>👤 Cliente:</b> {cliente}
────────────────────────
<b>📋 DETALLES DE VENTA:</b>
• <b>APV:</b> {apv}
• <b>Sucursal de Entrega:</b> {punto}
• <b>Fecha Promesa:</b> {fecha_p.strftime('%d/%m/%Y')}

<b>🛠 IMPLEMENTACIONES:</b>
<i>{impl if impl else 'Sin implementaciones adicionales'}</i>

<b>⚠️ CONDICIONES ESPECIALES:</b>
<i>{cond if cond else 'Ninguna'}</i>
────────────────────────
<i>Registrado el {datetime.now().strftime('%d/%m/%Y %H:%M')}</i>
                        """

                        enviar_telegram(mensaje_html)
                        enviar_correo_taller(idv, cliente, apv, punto, marca, modelo)

                        data_dict = {
                            "IDV": idv,
                            "Marca": marca,
                            "Modelo": modelo,
                            "Color": color,
                            "APV": apv,
                            "Punto": punto,
                            "Fecha_Promesa": fecha_p.strftime('%d/%m/%Y'),
                            "Cliente": cliente,
                            "Implementaciones": impl,
                            "Condiciones": cond
                        }

                        email_apv = obtener_email_apv(df_maestros, apv)

                        if email_apv:
                            enviar_correo_confirmacion(email_apv, data_dict)
                        else:
                            st.warning("⚠️ APV sin correo registrado en Maestros")

                        st.success("✅ Solicitud registrada correctamente")
                        st.balloons()

                except Exception as e:
                    st.error(f"Error: {e}")

# =========================================================
# 2. TALLER
# =========================================================
if ROL == "TALLER":
    with tabs[1]:
        st.header("Carga de Reporte")

        with st.expander("📂 Subir archivo de taller", expanded=True):
            file = st.file_uploader("Archivo .xls", type=["xls"])

        if file:
            try:
                df = pd.read_excel(file, engine='xlrd', dtype=str)
                df.columns = df.columns.str.strip()

                with st.expander("🔍 Vista previa del archivo", expanded=True):
                    st.dataframe(df.head())

                if 'IDV' in df.columns and 'Estad' in df.columns:

                    df['IDV'] = df['IDV'].str.replace('.0', '').str.strip()
                    df['Estad'] = df['Estad'].str.upper().str.strip()

                    def resolver(estados):
                        if any(~estados.isin(['T','TE'])):
                            return "🛠 En Proceso"
                        elif any(estados=='T'):
                            return "🏁 Terminado (En Taller)"
                        elif any(estados=='TE'):
                            return "✅ Terminado y Enviado"
                        return "🛠 En Proceso"

                    df_estado = df.groupby('IDV')['Estad'].apply(resolver).reset_index(name='Estado_Calculado')
                    df_final = df.merge(df_estado, on='IDV', how='left')

                    with st.expander("📊 Resultado procesado"):
                        st.dataframe(df_final.head())

                    if st.button("Confirmar carga"):
                        conn.update(spreadsheet=URL_SHEET, worksheet="ReporteTaller", data=df_final)
                        st.success("Base actualizada")

                else:
                    st.error("Columnas requeridas: IDV y Estad")

            except Exception as e:
                st.error(f"Error: {e}")

# =========================================================
# 3. CONSULTA
# =========================================================
with tabs[-1]:
    st.header("Consulta de Estado")

    @st.cache_data(ttl=60)
    def cargar():
        v = conn.read(spreadsheet=URL_SHEET, worksheet="Solicitudes")
        s = conn.read(spreadsheet=URL_SHEET, worksheet="ReporteTaller")

        for df in [v, s]:
            df['IDV'] = df['IDV'].astype(str).str.replace('.0','').str.strip()

        return v, s

    df_v, df_s = cargar()

    df_s = df_s.astype(str)
    df_v = df_v.astype(str)

    with st.expander("🔎 Búsqueda", expanded=True):
        query = st.text_input("Ingrese IDV")

        col1, col2 = st.columns(2)

        with col1:
            buscar = st.button("🔍 Consultar")

        with col2:
            refresh = st.button("🔄 Actualizar")

    if refresh:
        st.cache_data.clear()
        st.success("Datos actualizados")

    if buscar and query:
        res_v = df_v[df_v['IDV'] == query]

        with st.expander("📄 Resultado", expanded=True):

            if not res_v.empty:
                v = res_v.iloc[0]

                c1, c2, c3 = st.columns(3)
                c1.info(f"Cliente: {v['Cliente']}")
                c2.info(f"Vehículo: {v['Marca']} {v['Modelo']}")
                c3.info(f"Vendedor: {v['APV']}")

                res_s = df_s[df_s['IDV'] == query]

                if not res_s.empty:
                    st.success(res_s.iloc[0]['Estado_Calculado'])

                    with st.expander("🧾 Bitácora de taller"):
                        st.write(res_s)

                else:
                    st.warning("No en taller")

            else:
                res_s = df_s[df_s['IDV'] == query]

                if not res_s.empty:
                    st.info(res_s.iloc[0]['Estado_Calculado'])

                    with st.expander("🧾 Bitácora de taller"):
                        st.write(res_s)

                else:
                    st.error("No encontrado")