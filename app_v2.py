import streamlit as st
import pandas as pd
import qrcode
import io
import zipfile
import smtplib
from email.message import EmailMessage
import time

# Configuración de la página
st.set_page_config(page_title="Gestor Torneo - QRs y Correos", page_icon="🏆", layout="wide")

# --- FUNCIONES AUXILIARES ---

def limpiar_dato(dato):
    """Limpia datos vacíos, espacios y formatos numéricos."""
    if pd.isna(dato):
        return ""
    txt = str(dato).strip()
    if txt.endswith(".0"):
        return txt[:-2]
    return txt

def generar_qr_bytes(dato):
    """Genera la imagen QR y devuelve los bytes."""
    qr = qrcode.QRCode(box_size=10, border=4)
    qr.add_data(dato)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()

def procesar_equipos(uploaded_file):
    """
    Lee el Excel y estructura los datos por EQUIPO.
    Retorna una lista de diccionarios con toda la info lista para ZIP o Email.
    """
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl', dtype=str)
    except Exception as e:
        return None, f"Error al leer archivo: {e}"

    equipos_procesados = []

    # Indices de columnas (A=0, B=1, ... J=9, etc.)
    # B(1): Escuela, D(3): Equipo, E(4): Categoria
    # I(8): Celular Asesor (QR), J(9): Correo Asesor
    
    # Columnas de Alumnos (Matricula): K(10), R(17), Y(24), AF(31), AN(39)
    columnas_matriculas = [10, 17, 24, 31, 39]

    for index, row in df.iterrows():
        # 1. Datos del Equipo
        escuela = limpiar_dato(row.iloc[1])
        equipo = limpiar_dato(row.iloc[3])
        categoria = limpiar_dato(row.iloc[4])
        
        if not escuela or not equipo:
            continue # Saltar filas vacías

        # Nombre carpeta limpio
        nombre_carpeta = f"{escuela} {equipo} {categoria}".strip()
        for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
            nombre_carpeta = nombre_carpeta.replace(char, '-')

        # 2. Datos del Asesor (Coach)
        celular_asesor = limpiar_dato(row.iloc[8]) # Columna I
        correo_asesor = limpiar_dato(row.iloc[9])  # Columna J
        
        # Lista para guardar las imágenes de este equipo
        imagenes_equipo = []

        # Generar QR Asesor
        if celular_asesor:
            qr_bytes = generar_qr_bytes(celular_asesor)
            imagenes_equipo.append({
                "nombre_archivo": f"Coach_{celular_asesor}.png",
                "bytes": qr_bytes
            })

        # 3. Datos Alumnos (Iterar las 5 posiciones)
        for col_idx in columnas_matriculas:
            if col_idx < len(row):
                matricula = limpiar_dato(row.iloc[col_idx])
                if matricula:
                    qr_bytes = generar_qr_bytes(matricula)
                    imagenes_equipo.append({
                        "nombre_archivo": f"Alumno_{matricula}.png",
                        "bytes": qr_bytes
                    })

        # Guardar todo el paquete del equipo
        equipos_procesados.append({
            "Carpeta": nombre_carpeta,
            "Escuela": escuela,
            "Equipo": equipo,
            "Correo_Coach": correo_asesor,
            "Imagenes": imagenes_equipo
        })

    return equipos_procesados, "Ok"

# --- INTERFAZ DE USUARIO ---

st.title("Gestor de Torneo: QRs y Envíos 🚀")

uploaded_file = st.file_uploader("Cargar Excel (.xlsx)", type=["xlsx"])

if "equipos_data" not in st.session_state:
    st.session_state.equipos_data = None

if uploaded_file:
    datos, msg = procesar_equipos(uploaded_file)
    if datos:
        st.session_state.equipos_data = datos
        st.success(f"✅ Se procesaron {len(datos)} equipos correctamente.")
    else:
        st.error(msg)

if st.session_state.equipos_data:
    datos = st.session_state.equipos_data
    
    st.divider()

    # --- ACCIÓN 1: DESCARGAR ZIP ---
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.subheader("📂 1. Descargar Archivos")
        st.write("Genera un ZIP con carpetas organizadas para tu respaldo.")
        
        if st.button("Generar ZIP"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for equipo in datos:
                    for img in equipo["Imagenes"]:
                        # Ruta: CarpetaEquipo/Archivo.png
                        ruta = f"{equipo['Carpeta']}/{img['nombre_archivo']}"
                        zip_file.writestr(ruta, img['bytes'])
            
            st.download_button(
                label="⬇️ Descargar ZIP Completo",
                data=zip_buffer.getvalue(),
                file_name="QRs_Torneo_Completo.zip",
                mime="application/zip",
                type="primary"
            )

    # --- ACCIÓN 2: ENVIAR CORREOS ---
    with col_b:
        st.subheader("📧 2. Enviar a Coaches")
        st.write("Envía un correo por equipo al Asesor (Columna J).")
        
        # Filtro: Solo equipos que tienen correo de coach y al menos una imagen
        equipos_validos = [e for e in datos if e['Correo_Coach'] and "@" in e['Correo_Coach'] and e['Imagenes']]
        st.info(f"Equipos listos para enviar: **{len(equipos_validos)}** de {len(datos)}")

        with st.expander("🔐 Configurar Correo (Gmail/Outlook)"):
            email_user = st.text_input("Tu Correo")
            email_pass = st.text_input("Contraseña de Aplicación", type="password")
            proveedor = st.selectbox("Proveedor", ["Gmail", "Outlook/Office365", "Yahoo"])
            asunto = st.text_input("Asunto", value="QRs de Acceso - Torneo de Robótica")
            mensaje_base = st.text_area("Mensaje", value="Estimado Coach,\n\nAdjunto encontrará los códigos QR de acceso para su equipo.\n\nSaludos.")

        if st.button("🚀 Enviar Correos a Coaches"):
            if not email_user or not email_pass:
                st.error("Faltan credenciales.")
            else:
                barra = st.progress(0)
                status = st.empty()
                errores = []
                enviados = 0
                
                # Configurar SMTP
                host, port = {
                    "Gmail": ("smtp.gmail.com", 465),
                    "Outlook/Office365": ("smtp.office365.com", 587),
                    "Yahoo": ("smtp.mail.yahoo.com", 465)
                }[proveedor]

                try:
                    # Conexión al servidor
                    if proveedor == "Outlook/Office365":
                        server = smtplib.SMTP(host, port)
                        server.starttls()
                    else:
                        server = smtplib.SMTP_SSL(host, port)
                    
                    server.login(email_user, email_pass)
                    
                    # Iterar envíos
                    total = len(equipos_validos)
                    for i, equipo in enumerate(equipos_validos):
                        progreso = (i + 1) / total
                        barra.progress(progreso)
                        status.text(f"Enviando a equipo: {equipo['Equipo']} ({equipo['Correo_Coach']})")
                        
                        msg = EmailMessage()
                        msg['Subject'] = f"{asunto} - {equipo['Equipo']}"
                        msg['From'] = email_user
                        msg['To'] = equipo['Correo_Coach']
                        msg.set_content(mensaje_base)
                        
                        # Adjuntar TODAS las imágenes del equipo
                        for img in equipo['Imagenes']:
                            msg.add_attachment(img['bytes'], maintype='image', subtype='png', filename=img['nombre_archivo'])
                        
                        try:
                            server.send_message(msg)
                            enviados += 1
                            time.sleep(1) # Pausa anti-spam
                        except Exception as e:
                            errores.append(f"{equipo['Equipo']}: {str(e)}")
                    
                    server.quit()
                    st.success(f"Finalizado. Enviados: {enviados}/{total}")
                    if errores:
                        st.error("Errores:")
                        st.write(errores)
                        
                except Exception as e:
                    st.error(f"Error de conexión general: {e}")

    st.divider()
    
    # Tabla de revisión
    with st.expander("🔍 Ver detalle de datos detectados"):
        resumen = []
        for e in datos:
            resumen.append({
                "Escuela": e['Escuela'],
                "Equipo": e['Equipo'],
                "Correo Coach": e['Correo_Coach'],
                "QRs Generados": len(e['Imagenes'])
            })
        st.dataframe(pd.DataFrame(resumen))