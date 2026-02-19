import streamlit as st
import pandas as pd
import qrcode
import io
import zipfile
import smtplib
from email.message import EmailMessage
import time
import xlsxwriter
import os

# Configuración de la página
st.set_page_config(page_title="Sistema de Registro", page_icon="🎓", layout="wide")

# --- FUNCIONES AUXILIARES Y DE LÓGICA (Sin cambios) ---

def limpiar_dato(dato):
    if pd.isna(dato): return ""
    txt = str(dato).strip()
    return txt[:-2] if txt.endswith(".0") else txt

def generar_qr_bytes(dato):
    qr = qrcode.QRCode(box_size=10, border=4)
    qr.add_data(dato)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()

def cargar_dataframe(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl', header=None)
        df = df.ffill()
        df = df.iloc[1:].reset_index(drop=True)
        return df
    except Exception as e:
        st.error(f"Error: {e}")
        return None

def generar_excel_resumen(df_original):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    
    # 1. ASESORES
    sheet_asesores = workbook.add_worksheet("Asesores")
    cols_asesor = ["Escuela", "Nombre", "Ap. Paterno", "Ap. Materno", "Celular", "Correo"]
    for c, val in enumerate(cols_asesor): sheet_asesores.write(0, c, val, header_fmt)
        
    asesores_unicos = set()
    row_asesor = 1
    for _, row in df_original.iterrows():
        nombre = limpiar_dato(row.iloc[5])
        celular = limpiar_dato(row.iloc[8])
        if nombre and (nombre, celular) not in asesores_unicos:
            asesores_unicos.add((nombre, celular))
            datos = [limpiar_dato(row.iloc[1]), nombre, limpiar_dato(row.iloc[6]), 
                     limpiar_dato(row.iloc[7]), celular, limpiar_dato(row.iloc[9])]
            for c, val in enumerate(datos): sheet_asesores.write(row_asesor, c, val)
            row_asesor += 1

    # 2. ALUMNOS (TRANSPOSICIÓN)
    config_pos = [[10, 11, 12, 13, 16], [17, 18, 19, 20, 23], [24, 25, 26, 27, 30], 
                  [31, 32, 33, 34, 37], [39, 40, 41, 42, 45]]
    headers_al = ["Escuela", "Equipo", "Categoría", "Matrícula", "Ap. Paterno", "Ap. Materno", "Nombre", "Correo Inst."]
    
    sheets = {
        "Línea":      {"obj": workbook.add_worksheet("Línea"), "row": 1, "max": 4},
        "Laberinto":  {"obj": workbook.add_worksheet("Laberinto"), "row": 1, "max": 4},
        "Escenario":  {"obj": workbook.add_worksheet("Escenario"), "row": 1, "max": 5},
    }
    for k in sheets:
        for c, val in enumerate(headers_al): sheets[k]["obj"].write(0, c, val, header_fmt)

    for _, row in df_original.iterrows():
        escuela = limpiar_dato(row.iloc[1])
        equipo = limpiar_dato(row.iloc[3])
        cat_txt = limpiar_dato(row.iloc[4])
        if not escuela or not equipo: continue

        target = None
        if "línea" in cat_txt.lower() or "linea" in cat_txt.lower(): target = "Línea"
        elif "laberinto" in cat_txt.lower(): target = "Laberinto"
        elif "escenario" in cat_txt.lower(): target = "Escenario"
        
        if target:
            cfg = sheets[target]
            for i in range(cfg["max"]):
                idx = config_pos[i]
                if idx[0] < len(row):
                    mat = limpiar_dato(row.iloc[idx[0]])
                    if mat:
                        d = [escuela, equipo, target, mat, limpiar_dato(row.iloc[idx[1]]), 
                             limpiar_dato(row.iloc[idx[2]]), limpiar_dato(row.iloc[idx[3]]), 
                             limpiar_dato(row.iloc[idx[4]])]
                        for c, v in enumerate(d): cfg["obj"].write(cfg["row"], c, v)
                        cfg["row"] += 1
    workbook.close()
    return output.getvalue(), len(asesores_unicos)

def procesar_zip_correo(df):
    equipos = []
    cols_mat = [10, 17, 24, 31, 39] 
    for _, row in df.iterrows():
        esc = limpiar_dato(row.iloc[1])
        eq = limpiar_dato(row.iloc[3])
        cat = limpiar_dato(row.iloc[4])
        if not esc or not eq: continue

        nom_carpeta = "".join([c if c.isalnum() or c in " -_" else "-" for c in f"{esc} {eq} {cat}".strip()])
        cel_coach = limpiar_dato(row.iloc[8])
        mail_coach = limpiar_dato(row.iloc[9])
        
        imgs = []
        if cel_coach: imgs.append({"name": f"Coach_{cel_coach}.png", "bytes": generar_qr_bytes(cel_coach)})
        
        max_al = 5 if "escenario" in str(cat).lower() else 4
        for i, c_idx in enumerate(cols_mat):
            if i >= max_al: break
            if c_idx < len(row):
                mat = limpiar_dato(row.iloc[c_idx])
                if mat: imgs.append({"name": f"Alumno_{mat}.png", "bytes": generar_qr_bytes(mat)})

        equipos.append({"Carpeta": nom_carpeta, "Equipo": eq, "Correo": mail_coach, "Imagenes": imgs})
    return equipos

# --- INTERFAZ DE USUARIO (NUEVA ESTRUCTURA) ---

# 1. ENCABEZADO INSTITUCIONAL
c_img_izq, c_titulo, c_img_der = st.columns([1, 4, 1], gap="medium")

with c_img_izq:
    # Intenta cargar jpg o png por si acaso
    if os.path.exists("assets/UANL-color-negro.jpg"):
        st.image("assets/UANL-color-negro.jpg", use_container_width=True)
    elif os.path.exists("assets/UANL-color-negro.png"):
        st.image("assets/UANL-color-negro.png", use_container_width=True)
    else:
        st.warning("Logo UANL no encontrado")

with c_titulo:
    st.markdown("""
    <h1 style='text-align: center; color: #333; font-size: 32px;'>
        Procesamiento de Datos de Registro
    </h1>
    """, unsafe_allow_html=True)

with c_img_der:
    if os.path.exists("assets/Logo-Excelencia-Negro.png"):
        st.image("assets/Logo-Excelencia-Negro.png", use_container_width=True)
    else:
        st.warning("Logo Excelencia no encontrado")

st.markdown("---")

# 2. CARGA DE ARCHIVO
uploaded_file = st.file_uploader("📂 Cargar Archivo Excel Maestro (.xlsx)", type=["xlsx"])

if "df_master" not in st.session_state: st.session_state.df_master = None
if "datos_proc" not in st.session_state: st.session_state.datos_proc = []

if uploaded_file:
    with st.spinner("Analizando estructura..."):
        df = cargar_dataframe(uploaded_file)
        if df is not None:
            st.session_state.df_master = df
            st.session_state.datos_proc = procesar_zip_correo(df)
            st.success(f"✅ Archivo cargado exitosamente. Se detectaron {len(st.session_state.datos_proc)} equipos.")

# MOSTRAR SECCIONES SOLO SI HAY DATOS
if st.session_state.df_master is not None:
    df = st.session_state.df_master
    datos = st.session_state.datos_proc
    
    # --- PARTE 2: REPORTES EXCEL (Uniformemente distribuido) ---
    st.write("### 📊 Generación de Reportes")
    with st.container(border=True):
        col_excel_1, col_excel_2 = st.columns([1, 2])
        
        excel_bytes, n_asesores = generar_excel_resumen(df)
        
        with col_excel_1:
            st.metric(label="Asesores Únicos", value=n_asesores)
            st.caption("Total de profesores sin repetir.")
            
        with col_excel_2:
            st.info("Descarga el reporte clasificado por categorías (Vertical).")
            st.download_button(
                label="📥 Descargar Reporte Excel Clasificado",
                data=excel_bytes,
                file_name="Reporte_Torneo_Vertical.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

    st.write("") # Espacio
    
    # --- PARTE 3: ACCIONES (ZIP Y EMAIL) ---
    # Distribuidos uniformemente debajo de la parte 2
    st.write("### 🚀 Acciones de Salida")
    
    col_izq, col_der = st.columns(2, gap="large")

    # COLUMNA IZQUIERDA: ZIP
    with col_izq:
        with st.container(border=True):
            st.subheader("📂 Descargar QRs")
            st.write("Genera un archivo ZIP con carpetas organizadas por equipo.")
            if st.button("Generar ZIP de Imágenes", use_container_width=True):
                b = io.BytesIO()
                with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
                    for eq in datos:
                        for img in eq["Imagenes"]:
                            z.writestr(f"{eq['Carpeta']}/{img['name']}", img['bytes'])
                st.download_button("⬇️ Guardar ZIP en PC", b.getvalue(), "QRs_Torneo.zip", "application/zip", use_container_width=True)

    # COLUMNA DERECHA: EMAIL
    with col_der:
        with st.container(border=True):
            st.subheader("📧 Enviar a Asesores")
            validos = [e for e in datos if e.get('Correo') and "@" in str(e.get('Correo'))]
            st.markdown(f"**{len(validos)} equipos** listos para envío.")
            
            with st.expander("⚙️ Configurar Envío", expanded=True):
                user = st.text_input("Tu Correo (Gmail/Outlook)")
                pwd = st.text_input("Contraseña de Aplicación", type="password")
                prov = st.selectbox("Proveedor", ["Gmail", "Outlook", "Yahoo"])
                
                # NUEVO: PERSONALIZACIÓN DEL MENSAJE
                st.markdown("**Mensaje para el Asesor:**")
                asunto_base = st.text_input("Asunto del correo", value="Accesos QR - Torneo de Robótica")
                mensaje_cuerpo = st.text_area("Cuerpo del correo", value="Estimado Coach,\n\nAdjunto a este correo encontrará los códigos QR de acceso para los integrantes de su equipo.\n\nEstos deberán ser presentados por sus alumnos en el momento de registrarse.\n\nFavor de distribuirlos.\n\nEn el caso de faltar alguno o presentar problemas, por favor notificar al coordinador del torneo.\n\nSaludos cordiales.", height=150)

            if st.button("✈️ Enviar Correos Masivos", type="primary", use_container_width=True):
                if not user or not pwd:
                    st.error("Faltan credenciales.")
                else:
                    progreso = st.progress(0)
                    estado = st.empty()
                    host, port = {
                        "Gmail": ("smtp.gmail.com", 465),
                        "Outlook": ("smtp.office365.com", 587),
                        "Yahoo": ("smtp.mail.yahoo.com", 465)
                    }[prov]
                    
                    try:
                        server = smtplib.SMTP(host, port) if prov == "Outlook" else smtplib.SMTP_SSL(host, port)
                        if prov == "Outlook": server.starttls()
                        server.login(user, pwd)
                        
                        enviados_count = 0
                        for i, eq in enumerate(validos):
                            # Actualizar barra
                            progreso.progress((i + 1) / len(validos))
                            estado.text(f"Enviando a: {eq['Equipo']} ({eq['Correo']})")
                            
                            msg = EmailMessage()
                            msg['Subject'] = f"{asunto_base} - {eq['Equipo']}"
                            msg['From'] = user
                            msg['To'] = eq['Correo']
                            msg.set_content(mensaje_cuerpo) # Usamos el mensaje personalizado
                            
                            for img in eq['Imagenes']:
                                msg.add_attachment(img['bytes'], maintype='image', subtype='png', filename=img['name'])
                            
                            server.send_message(msg)
                            enviados_count += 1
                            time.sleep(1.5) # Pausa leve anti-spam
                        
                        server.quit()
                        st.balloons()
                        st.success(f"¡Proceso finalizado! Se enviaron {enviados_count} correos exitosamente.")
                    except Exception as e:
                        st.error(f"Error de conexión: {e}")