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
import cv2 
import numpy as np 
from datetime import datetime

# Configuración de página
st.set_page_config(page_title="Sistema Integral de Torneo", page_icon="📲", layout="wide")

# ==========================================
# 1. INICIALIZACIÓN DE MEMORIA
# ==========================================
if "df_master" not in st.session_state:
    st.session_state.df_master = None
if "datos_proc" not in st.session_state:
    st.session_state.datos_proc = []

# ==========================================
# 2. FUNCIONES AUXILIARES (BACKEND ADAPTADO)
# ==========================================

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
        # Lee el archivo respetando la fila de encabezados
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        # Elimina filas donde la Matrícula (Columna index 3) esté vacía
        df = df.dropna(subset=[df.columns[3]])
        return df
    except Exception as e:
        st.error(f"Error al cargar el archivo: {e}")
        return None

def generar_excel_resumen(df_original):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    
    # Hoja 1: Todos los alumnos centralizados
    sheet_todos = workbook.add_worksheet("Todos los Participantes")
    headers = list(df_original.columns)
    for c, val in enumerate(headers): 
        sheet_todos.write(0, c, str(val), header_fmt)
        
    for r, (_, row) in enumerate(df_original.iterrows(), start=1):
        for c, val in enumerate(row):
            sheet_todos.write(r, c, str(val))
            
    # Hojas Inteligentes Dinámicas por cada Categoría encontrada
    categorias = df_original.iloc[:, 2].dropna().unique()
    for cat in categorias:
        cat_clean = str(cat).strip()
        if not cat_clean: continue
        sheet_name = "".join([c for c in cat_clean if c.isalnum() or c in " "])[:30]
        sheet = workbook.add_worksheet(sheet_name)
        
        for c, val in enumerate(headers): 
            sheet.write(0, c, str(val), header_fmt)
            
        df_cat = df_original[df_original.iloc[:, 2].astype(str).str.strip() == cat_clean]
        for r, (_, row) in enumerate(df_cat.iterrows(), start=1):
            for c, val in enumerate(row):
                sheet.write(r, c, str(val))
                
    workbook.close()
    total_equipos = df_original.iloc[:, 1].nunique()
    return output.getvalue(), total_equipos

def procesar_zip_correo(df):
    equipos_dict = {}
    for _, row in df.iterrows():
        esc = limpiar_dato(row.iloc[0])
        eq = limpiar_dato(row.iloc[1])
        cat = limpiar_dato(row.iloc[2])
        mat = limpiar_dato(row.iloc[3])
        ap_pat = limpiar_dato(row.iloc[4])
        nom = limpiar_dato(row.iloc[6])
        correo = limpiar_dato(row.iloc[7])
        
        if not eq or not mat: continue
        
        key = f"{esc}_{eq}_{cat}"
        if key not in equipos_dict:
            nom_carpeta = "".join([c if c.isalnum() or c in " -_" else "-" for c in f"{esc} {eq} {cat}".strip()])
            equipos_dict[key] = {
                "Carpeta": nom_carpeta,
                "Equipo": eq,
                "Correos": set([correo]) if correo else set(),
                "Imagenes": []
            }
        else:
            if correo: equipos_dict[key]["Correos"].add(correo)
        
        # Generar QR único usando la Matrícula
        qr_bytes = generar_qr_bytes(mat)
        equipos_dict[key]["Imagenes"].append({
            "name": f"Alumno_{mat}_{nom}_{ap_pat}.png",
            "bytes": qr_bytes
        })
        
    lista_resultado = []
    for key, info in equipos_dict.items():
        info["Correo"] = ", ".join(info["Correos"]) if info["Correos"] else ""
        lista_resultado.append(info)
    return lista_resultado

def registrar_asistencia(codigo_leido, df_master):
    archivo_log = "asistencia_torneo.csv"
    if not os.path.exists(archivo_log):
        df_log = pd.DataFrame(columns=["Fecha", "Hora", "Codigo", "Nombre", "Rol", "Equipo"])
        df_log.to_csv(archivo_log, index=False)
    
    df_log = pd.read_csv(archivo_log)
    codigo_leido_str = str(codigo_leido).strip()
    
    if codigo_leido_str in df_log["Codigo"].astype(str).values:
        return "DUPLICADO", None

    persona_encontrada = None
    for _, row in df_master.iterrows():
        if limpiar_dato(row.iloc[3]) == codigo_leido_str:
            persona_encontrada = row
            break

    if persona_encontrada is not None:
        escuela = limpiar_dato(persona_encontrada.iloc[0])
        equipo = limpiar_dato(persona_encontrada.iloc[1])
        categoria = limpiar_dato(persona_encontrada.iloc[2])
        nombre_completo = f"{limpiar_dato(persona_encontrada.iloc[6])} {limpiar_dato(persona_encontrada.iloc[4])} {limpiar_dato(persona_encontrada.iloc[5])}".strip()
        
        nuevo_registro = {
            "Fecha": datetime.now().strftime("%Y-%m-%d"),
            "Hora": datetime.now().strftime("%H:%M:%S"),
            "Codigo": codigo_leido_str,
            "Nombre": nombre_completo,
            "Rol": f"Participante ({categoria})",
            "Equipo": f"{equipo} ({escuela})"
        }
        pd.DataFrame([nuevo_registro]).to_csv(archivo_log, mode='a', header=False, index=False)
        return "EXITO", nuevo_registro
    else:
        return "NO_ENCONTRADO", None

# ==========================================
# 3. INTERFAZ DE USUARIO
# ==========================================

c1, c2, c3 = st.columns([1, 4, 1])
with c2:
    st.markdown("<h2 style='text-align: center;'>Sistema de Gestión de Torneo (Módulo Vertical)</h2>", unsafe_allow_html=True)
st.markdown("---")

st.sidebar.title("Menú de Opciones")
modo = st.sidebar.radio("Navegación:", ["1. Gestión de Datos (Admin)", "2. Escáner de Asistencia"])

if modo == "1. Gestión de Datos (Admin)":
    st.header("📂 Carga de Datos")
    uploaded_file = st.file_uploader("Cargar Excel Maestro Vertical (.xlsx)", type=["xlsx"])
    
    if uploaded_file:
        with st.spinner("Procesando estructura..."):
            df = cargar_dataframe(uploaded_file)
            if df is not None:
                st.session_state.df_master = df 
                st.session_state.datos_proc = procesar_zip_correo(df)
                st.success("✅ Base de datos cargada y limpia en memoria.")

    if st.session_state.df_master is not None:
        df = st.session_state.df_master
        datos = st.session_state.datos_proc
        
        st.write("### 📊 Generación de Reportes")
        with st.container(border=True):
            col_excel_1, col_excel_2 = st.columns([1, 2])
            excel_bytes, total_eq = generar_excel_resumen(df)
            
            with col_excel_1:
                st.metric(label="Total Equipos Únicos", value=total_eq)
            with col_excel_2:
                st.download_button("📥 Descargar Reporte Segmentado", excel_bytes, "Reporte_Torneo_Final.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)

        st.write("### 🚀 Acciones de Salida")
        col_izq, col_der = st.columns(2, gap="large")

        with col_izq:
            with st.container(border=True):
                st.subheader("📂 Descargar QRs por Equipos")
                if st.button("Generar ZIP Compilado", use_container_width=True):
                    b = io.BytesIO()
                    with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
                        for eq in datos:
                            for img in eq["Imagenes"]:
                                z.writestr(f"{eq['Carpeta']}/{img['name']}", img['bytes'])
                    st.download_button("⬇️ Guardar ZIP", b.getvalue(), "QRs_Torneo.zip", "application/zip", use_container_width=True)

        with col_der:
            with st.container(border=True):
                st.subheader("📧 Notificaciones por Correo")
                validos = [e for e in datos if e.get('Correo')]
                st.markdown(f"**{len(validos)} canales de envío detectados.**")
                
                with st.expander("⚙️ Configuración SMTP"):
                    user = st.text_input("Tu Correo")
                    pwd = st.text_input("App Password", type="password")
                    prov = st.selectbox("Servidor", ["Gmail", "Outlook", "Yahoo"])
                    asunto_base = st.text_input("Asunto", value="Accesos QR - Torneo de Robótica")
                    mensaje_cuerpo = st.text_area("Mensaje", value="Estimado Competidor,\n\nAdjunto a este correo encontrarás tu código QR oficial de acceso para el torneo.\n\nSaludos cordiales.")

                if st.button("✈️ Enviar Correos Masivos", type="primary", use_container_width=True):
                    if not user or not pwd:
                        st.error("Faltan credenciales.")
                    else:
                        progreso = st.progress(0)
                        estado = st.empty()
                        host, port = {"Gmail": ("smtp.gmail.com", 465), "Outlook": ("smtp.office365.com", 587), "Yahoo": ("smtp.mail.yahoo.com", 465)}[prov]
                        
                        try:
                            server = smtplib.SMTP(host, port) if prov == "Outlook" else smtplib.SMTP_SSL(host, port)
                            if prov == "Outlook": server.starttls()
                            server.login(user, pwd)
                            
                            enviados_count = 0
                            for i, eq in enumerate(validos):
                                progreso.progress((i + 1) / len(validos))
                                estado.text(f"Enviando a: {eq['Equipo']}")
                                
                                msg = EmailMessage()
                                msg['Subject'] = f"{asunto_base} - {eq['Equipo']}"
                                msg['From'] = user
                                msg['To'] = eq['Correo']
                                msg.set_content(mensaje_cuerpo)
                                
                                for img in eq['Imagenes']:
                                    msg.add_attachment(img['bytes'], maintype='image', subtype='png', filename=img['name'])
                                server.send_message(msg)
                                enviados_count += 1
                                time.sleep(1.2)
                            
                            server.quit()
                            st.balloons()
                            st.success(f"¡Éxito! {enviados_count} correos procesados.")
                        except Exception as e:
                            st.error(f"Error en envío: {e}")

elif modo == "2. Escáner de Asistencia":
    st.header("📱 Escáner de Entrada")
    
    if st.session_state.df_master is None:
        st.warning("⚠️ Primero debes cargar el Excel Maestro en la sección 'Gestión de Datos'.")
    else:
        df_master = st.session_state.df_master
        col_cam, col_res = st.columns([1, 2])
        
        with col_cam:
            st.write("**Capturar QR:**")
            img_file_buffer = st.camera_input("Enfoca el código del participante")
        
        with col_res:
            st.write("**Resultado del Escaneo:**")
            if img_file_buffer is not None:
                codigo = leer_qr_desde_imagen(img_file_buffer)
                if codigo:
                    st.info(f"Matrícula leída: `{codigo}`")
                    status, info = registrar_asistencia(codigo, df_master)
                    
                    if status == "EXITO":
                        st.success("✅ ACCESO CONCEDIDO")
                        st.markdown(f"**Nombre:** {info['Nombre']} <br>**Asignación:** {info['Rol']} <br>**Equipo:** {info['Equipo']}", unsafe_allow_html=True)
                    elif status == "DUPLICADO":
                        st.warning("⚠️ REGISTRO PREVIO EXISTENTE HOY.")
                    else:
                        st.error("❌ MATRÍCULA NO LOCALIZADA EN EL PADRÓN MAESTRO.")
                else:
                    st.error("No se detecta un QR legible. Intenta ajustar el enfoque.")

        st.divider()
        st.subheader("📋 Registro de Ingresos en Vivo")
        if os.path.exists("asistencia_torneo.csv"):
            df_asistencia = pd.read_csv("asistencia_torneo.csv")
            st.dataframe(df_asistencia.sort_values(by="Hora", ascending=False), use_container_width=True)
            
            csv = df_asistencia.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Descargar Corte de Lista", csv, "Asistencia_Torneo_Corte.csv", "text/csv")