import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time

st.set_page_config(page_title="Mail Merge TRIS XXIII", page_icon="✉️")

st.title("✉️ Mail Merge - TRIS XXIII")

# --- 1. CONFIGURACIÓN DEL EMISOR ---
with st.sidebar:
    st.header("1. Configuración de Correo")
    email_usuario = st.text_input("Tu Correo (UANL/Gmail)", placeholder="ejemplo@uanl.edu.mx")
    email_password = st.text_input("Contraseña de Aplicación", type="password")
    
    st.divider()
    servidor_smtp = st.selectbox("Servidor SMTP", ["smtp.office365.com", "smtp.gmail.com"])
    st.caption("UANL usa smtp.office365.com")

# --- 2. CARGA DE DATOS (Múltiples hojas) ---
st.header("2. Carga de Datos")
archivo = st.file_uploader("Sube tu Excel o CSV", type=['csv', 'xlsx'])

if archivo:
    try:
        if archivo.name.endswith('.csv'):
            df = pd.read_csv(archivo, sep=None, engine='python', encoding='latin-1')
        else:
            excel_file = pd.ExcelFile(archivo)
            hoja_seleccionada = st.selectbox("Selecciona la pestaña del LOG:", excel_file.sheet_names)
            df = pd.read_excel(archivo, sheet_name=hoja_seleccionada)
        
        st.success("Datos cargados.")
        
        cols = list(df.columns)
        c1, c2, c3 = st.columns(3)
        with c1: col_equipo = st.selectbox("Columna EQUIPO", cols, index=0)
        with c2: col_correo = st.selectbox("Columna CORREO", cols, index=1 if len(cols)>1 else 0)
        with c3: col_link = st.selectbox("Columna LINK", cols, index=2 if len(cols)>2 else 0)

        # --- 3. VISTA PREVIA Y PRUEBA ---
        st.divider()
        st.header("3. Vista Previa y Prueba")
        asunto = st.text_input("Asunto del correo", "Carpeta de Entregables - TRIS XXIII")
        
        ejemplo_idx = st.number_input("Fila para previsualizar:", 0, len(df)-1, 0)
        fila = df.iloc[ejemplo_idx]

        cuerpo_txt = f"""Hola Coach del equipo {fila[col_equipo]},

Esperamos que se encuentren muy bien. Les compartimos el enlace para subir sus entregables:

Link: {fila[col_link]}

Instrucciones:
1. Den clic al link.
2. Si pide acceso, den clic en "Solicitar acceso".
3. Suban sus archivos a la carpeta.

¡Mucho éxito!"""

        with st.expander("👁️ Ver cómo quedará el correo"):
            st.code(f"Para: {fila[col_correo]}\nAsunto: {asunto}\n\n{cuerpo_txt}")

        # BOTÓN DE PRUEBA
        if st.button("📧 Enviarme un correo de prueba"):
            if not email_usuario or not email_password:
                st.error("Configura tus credenciales en el menú lateral.")
            else:
                try:
                    server = smtplib.SMTP(servidor_smtp, 587)
                    server.starttls()
                    server.login(email_usuario, email_password)
                    
                    msg = MIMEMultipart()
                    msg['From'] = email_usuario
                    msg['To'] = email_usuario # Se envía a ti mismo
                    msg['Subject'] = "[PRUEBA] " + asunto
                    msg.attach(MIMEText(cuerpo_txt, 'plain'))
                    
                    server.send_message(msg)
                    server.quit()
                    st.success(f"✅ ¡Correo de prueba enviado a {email_usuario}! Revisa tu bandeja.")
                except Exception as e:
                    st.error(f"Error en prueba: {e}")

        # --- 4. ENVÍO MASIVO ---
        st.divider()
        st.header("4. Envío Masivo")
        st.warning("Esto enviará correos a todos los destinatarios de la lista.")
        
        confirmar = st.toggle("Confirmar que todo está listo")
        
        if st.button("🚀 INICIAR ENVÍO A TODOS", disabled=not confirmar):
            try:
                server = smtplib.SMTP(servidor_smtp, 587)
                server.starttls()
                server.login(email_usuario, email_password)

                bar = st.progress(0)
                status = st.empty()

                for i, r in df.iterrows():
                    msg = MIMEMultipart()
                    msg['From'] = email_usuario
                    msg['To'] = str(r[col_correo]).strip()
                    msg['Subject'] = asunto
                    cuerpo_envio = cuerpo_txt.replace(str(fila[col_equipo]), str(r[col_equipo])).replace(str(fila[col_link]), str(r[col_link]))
                    msg.attach(MIMEText(cuerpo_envio, 'plain'))
                    
                    server.send_message(msg)
                    status.text(f"Enviando: {r[col_correo]}")
                    bar.progress((i + 1) / len(df))
                    time.sleep(2) # Pausa anti-spam

                server.quit()
                st.success("¡Envío masivo terminado!")
            except Exception as e:
                st.error(f"Error: {e}")

    except Exception as e:
        st.error(f"Error al procesar: {e}")