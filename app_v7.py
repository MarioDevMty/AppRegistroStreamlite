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
import re

# Configuración de página
st.set_page_config(page_title="Sistema de Gestión de Torneos UANL", page_icon="📲", layout="wide")

# ==========================================
# 1. CONFIGURACIÓN DE BASE DE DATOS LOCAL
# ==========================================
DB_DIR = "database_torneos"
if not os.path.exists(DB_DIR):
    os.makedirs(DB_DIR)

EVENTOS_REGISTRY = os.path.join(DB_DIR, "registro_eventos.csv")
if not os.path.exists(EVENTOS_REGISTRY):
    df_init = pd.DataFrame(columns=[
        "id_evento", "nombre", "fecha_inicio", "horario_inicio", "fecha_fin", "horario_fin",
        "dependencia", "sede", "domicilio", "magnitud", "dirigido", "clasificacion", "tipo",
        "comite", "responsables"
    ])
    df_init.to_csv(EVENTOS_REGISTRY, index=False)

USUARIOS = {
    "adminuanl": {"password": "Uanl2026Admin", "role": "admin", "nombre": "Administrador General"},
    "creadoruanl": {"password": "Uanl2026Creador", "role": "creador", "nombre": "Organizador de Eventos"}
}

# Variables de Estado de Sesión
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_role" not in st.session_state:
    st.session_state.user_role = None
if "user_name" not in st.session_state:
    st.session_state.user_name = None
if "id_evento_activo" not in st.session_state:
    st.session_state.id_evento_activo = None

# ==========================================
# 2. FUNCIONES DE BASE DE DATOS
# ==========================================

def slugify(text):
    text = text.lower().strip()
    text = re.sub(r'[^a-z0-9_ -]', '', text)
    text = re.sub(r'[- ]+', '_', text)
    return f"{text}_{int(time.time())}"

def obtener_eventos():
    return pd.read_csv(EVENTOS_REGISTRY)

def guardar_nuevo_evento(info):
    df = obtener_eventos()
    id_ev = slugify(info["nombre"])
    info["id_evento"] = id_ev
    df = pd.concat([df, pd.DataFrame([info])], ignore_index=True)
    df.to_csv(EVENTOS_REGISTRY, index=False)
    return id_ev

def eliminar_evento_completo(id_evento):
    df = obtener_eventos()
    df = df[df["id_evento"] != id_evento]
    df.to_csv(EVENTOS_REGISTRY, index=False)
    
    for ext in [f"{id_evento}_master.xlsx", f"{id_evento}_asistencia.csv"]:
        path = os.path.join(DB_DIR, ext)
        if os.path.exists(path):
            os.remove(path)

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

def leer_qr_desde_imagen(img_buffer):
    try:
        bytes_data = img_buffer.getvalue()
        cv_img = cv2.imdecode(np.frombuffer(bytes_data, np.uint8), cv2.IMREAD_COLOR)
        detector = cv2.QRCodeDetector()
        data, _, _ = detector.detectAndDecode(cv_img)
        return data if data else None
    except Exception:
        return None

def registrar_asistencia(codigo_leido, id_evento, df_master):
    archivo_log = os.path.join(DB_DIR, f"{id_evento}_asistencia.csv")
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
                "Carpeta": nom_carpeta, "Equipo": eq, "Correos": set([correo]) if correo else set(), "Imagenes": []
            }
        else:
            if correo: equipos_dict[key]["Correos"].add(correo)
        
        qr_bytes = generar_qr_bytes(mat)
        equipos_dict[key]["Imagenes"].append({
            "name": f"Alumno_{mat}_{nom}_{ap_pat}.png", "bytes": qr_bytes
        })
    lista_resultado = []
    for key, info in equipos_dict.items():
        info["Correo"] = ", ".join(info["Correos"]) if info["Correos"] else ""
        lista_resultado.append(info)
    return lista_resultado

def generar_excel_resumen(df_original):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    sheet_todos = workbook.add_worksheet("Todos los Participantes")
    headers = list(df_original.columns)
    for c, val in enumerate(headers): sheet_todos.write(0, c, str(val), header_fmt)
    for r, (_, row) in enumerate(df_original.iterrows(), start=1):
        for c, val in enumerate(row): sheet_todos.write(r, c, str(val))
    
    categorias = df_original.iloc[:, 2].dropna().unique()
    for cat in categorias: # <--- CORREGIDO AQUÍ
        cat_clean = str(cat).strip()
        if not cat_clean: continue
        sheet_name = "".join([c for c in cat_clean if c.isalnum() or c in " "])[:30]
        sheet = workbook.add_worksheet(sheet_name)
        for c, val in enumerate(headers): sheet.write(0, c, str(val), header_fmt)
        df_cat = df_original[df_original.iloc[:, 2].astype(str).str.strip() == cat_clean]
        for r, (_, row) in enumerate(df_cat.iterrows(), start=1):
            for c, val in enumerate(row): sheet.write(r, c, str(val))
    workbook.close()
    return output.getvalue(), df_original.iloc[:, 1].nunique()

# ==========================================
# 3. INTERFAZ DE LOGUEO O PANEL PRINCIPAL
# ==========================================

if not st.session_state.logged_in:
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.write("")
        with st.container(border=True):
            st.markdown("<h3 style='text-align: center;'>🔐 Acceso al Sistema</h3>", unsafe_allow_html=True)
            usuario_input = st.text_input("Usuario")
            password_input = st.text_input("Contraseña", type="password")
            if st.button("Iniciar Sesión", type="primary", use_container_width=True):
                user = usuario_input.strip().lower()
                if user in USUARIOS and USUARIOS[user]["password"] == password_input:
                    st.session_state.logged_in = True
                    st.session_state.user_role = USUARIOS[user]["role"]
                    st.session_state.user_name = USUARIOS[user]["nombre"]
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas.")
else:
    df_evs = obtener_eventos()
    ev_activo = None
    if st.session_state.id_evento_activo:
        match = df_evs[df_evs["id_evento"] == st.session_state.id_evento_activo]
        if not match.empty:
            ev_activo = match.iloc[0].to_dict()

    # Banner Superior
    if ev_activo:
        st.markdown(f"<h2 style='text-align: center; color: #1E3A8A;'>💻 [EVENTO ACTIVO] {ev_activo['nombre']}</h2>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align: center; color: #6B7280;'>Sede: {ev_activo['sede']} | {ev_activo['fecha_inicio']} ({ev_activo['horario_inicio']})</p>", unsafe_allow_html=True)
    else:
        st.markdown("<h2 style='text-align: center; color: #DC2626;'>⚠️ Ningún Evento Seleccionado</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #6B7280;'>Por favor, ve a 'Historial / Selección' para activar un entorno de trabajo.</p>", unsafe_allow_html=True)
    st.markdown("---")

    # Sidebar
    st.sidebar.title("📲 Menú Principal")
    st.sidebar.write(f"Usuario: **{st.session_state.user_name}**")
    
    opciones_menu = ["📁 Historial / Selección", "✨ Crear Nuevo Evento", "📱 Escáner de Asistencia"]
    if st.session_state.user_role == "admin":
        opciones_menu.append("🛡️ Base de Datos General (Admin)")
    modo = st.sidebar.radio("Navegación:", opciones_menu)
    
    st.sidebar.markdown("---")
    if st.sidebar.button("Cerrar Sesión", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.id_evento_activo = None
        st.rerun()

    # --- ENTORNO 1: HISTORIAL Y SELECCIÓN ---
    if modo == "📁 Historial / Selección":
        st.header("🗂️ Listado Central de Eventos Creados")
        if df_evs.empty:
            st.info("Aún no existen eventos en la base de datos. Ve a 'Crear Nuevo Evento' para empezar.")
        else:
            st.write("Selecciona el evento con el que deseas trabajar en el escáner o enviar correos:")
            
            opciones_select = {row["nombre"]: row["id_evento"] for _, row in df_evs.iterrows()}
            ev_elegido = st.selectbox("Elegir Evento para Activar:", list(opciones_select.keys()))
            
            if st.button("🚀 Activar Evento Seleccionado", type="primary"):
                st.session_state.id_evento_activo = opciones_select[ev_elegido]
                st.success(f"Entorno cambiado a: {ev_elegido}")
                time.sleep(0.5)
                st.rerun()
                
            st.subheader("📋 Catálogo Maestro Almacenado")
            st.dataframe(df_evs[["nombre", "fecha_inicio", "sede", "dependencia", "tipo"]], use_container_width=True)

    # --- ENTORNO 2: CREAR NUEVO EVENTO ---
    elif modo == "✨ Crear Nuevo Evento":
        st.header("⚡ Registro de Nueva Ficha Técnica")
        
        tab_gen, tab_org, tab_xls = st.tabs(["📝 1. Ficha Técnica", "👥 2. Estructura Organizativa", "📊 3. Vincular Alumnos"])
        
        with tab_gen:
            with st.container(border=True):
                n_nombre = st.text_input("Nombre del Evento:", placeholder="Ej. XXII Torneo de Robótica Interpreparatorias STEM")
                c1, c2, c3, c4 = st.columns(4)
                with c1: n_fi = st.text_input("Fecha Inicio:", "02/04/2025")
                with c2: n_hi = st.text_input("Horario Inicio:", "09:00 a 18:00 horas")
                with c3: n_ff = st.text_input("Fecha Fin:", "02/04/2025")
                with c4: n_hf = st.text_input("Horario Fin:", "9:00 a 14:00 horas")
                n_dep = st.text_input("Dependencia Organizadora:", "Dirección del Sistema de Estudios del Nivel Medio Superior.")
                n_sed = st.text_input("Sede del Evento:", "Preparatoria 23 unidad Santa Catarina de la UANL")
                n_dom = st.text_area("Domicilio:", "Ave. San Francisco No. 198 de la colonia La Fama...")
                
                c_sel1, c_sel2, c_sel3, c_sel4 = st.columns(4)
                with c_sel1: n_mag = st.selectbox("Magnitud:", ["Local", "Nacional", "Internacional"])
                with c_sel2: n_dir = st.text_input("Dirigido a:", "Estudiantes de Bachillerato")
                with c_sel3: n_cla = st.text_input("Clasificación:", "Académico")
                with c_sel4: n_tip = st.text_input("Tipo de Evento:", "Torneo")

        with tab_org:
            with st.container(border=True):
                n_com = st.text_input("Comité Organizador:", "CAD de TIC")
                n_res = st.text_area("Responsables del Evento (Separados por coma):", "Diana Janeth Amaro Fernández, Alejandro Ojeda Ramírez, Lisbeth Cortez Hernandez")
                
                if st.button("💾 Crear Evento en Base de Datos", type="primary"):
                    if not n_nombre:
                        st.error("El nombre del evento es obligatorio.")
                    else:
                        nuevo_id = guardar_nuevo_evento({
                            "nombre": n_nombre, "fecha_inicio": n_fi, "horario_inicio": n_hi, "fecha_fin": n_ff, "horario_fin": n_hf,
                            "dependencia": n_dep, "sede": n_sed, "domicilio": n_dom, "magnitud": n_mag, "dirigido": n_dir,
                            "clasificacion": n_cla, "tipo": n_tip, "comite": n_com, "responsables": n_res
                        })
                        st.session_state.id_evento_activo = nuevo_id
                        st.success("¡Evento guardado e indexado!")
                        time.sleep(0.5)
                        st.rerun()

        with tab_xls:
            if not ev_activo:
                st.warning("⚠️ Debes activar o crear un evento en las pestañas previas antes de subir su Excel.")
            else:
                st.markdown(f"Subiendo lista de competidores para: **{ev_activo['nombre']}**")
                up_file = st.file_uploader("Cargar Excel Vertical (.xlsx)", type=["xlsx"])
                if up_file:
                    path_destino = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_master.xlsx")
                    with open(path_destino, "wb") as f:
                        f.write(up_file.getbuffer())
                    st.success("✅ Archivo guardado y enlazado al ID del torneo.")
                
                path_master_existente = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_master.xlsx")
                if os.path.exists(path_master_existente):
                    st.info("💡 Este evento ya cuenta con un padrón de alumnos cargado.")
                    df_m = pd.read_excel(path_master_existente).dropna(subset=[pd.read_excel(path_master_existente).columns[3]])
                    datos_proc = procesar_zip_correo(df_m)
                    
                    col_ex1, col_ex2 = st.columns(2)
                    with col_ex1:
                        if st.button("📦 Generar ZIP de QRs de este Evento"):
                            b = io.BytesIO()
                            with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
                                for eq in datos_proc:
                                    for img in eq["Imagenes"]:
                                        z.writestr(f"{eq['Carpeta']}/{img['name']}", img['bytes'])
                            st.download_button("⬇️ Descargar ZIP", b.getvalue(), f"QRs_{ev_activo['id_evento']}.zip", "application/zip", use_container_width=True)
                    with col_ex2:
                        ex_b, t_eq = generar_excel_resumen(df_m)
                        st.download_button("📥 Descargar Reporte Segmentado", ex_b, f"Reporte_{ev_activo['id_evento']}.xlsx", use_container_width=True)

    # --- ENTORNO 3: ESCÁNER DE ASISTENCIA ---
    elif modo == "📱 Escáner de Asistencia":
        if not ev_activo:
            st.error("⚠️ No hay ningún evento activo. Ve al menú 'Historial / Selección' primero.")
        else:
            path_m = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_master.xlsx")
            if not os.path.exists(path_m):
                st.warning("⚠️ Este evento no tiene una lista de alumnos vinculada. Súbela en la sección 'Crear Nuevo Evento'.")
            else:
                df_master = pd.read_excel(path_m).dropna(subset=[pd.read_excel(path_m).columns[3]])
                
                c_cam, c_res = st.columns([1, 2])
                with c_cam:
                    img_buffer = st.camera_input("Enfoca el código QR del alumno")
                with c_res:
                    if img_buffer is not None:
                        codigo = leer_qr_desde_imagen(img_buffer)
                        if codigo:
                            status, info = registrar_asistencia(codigo, ev_activo['id_evento'], df_master)
                            if status == "EXITO":
                                st.success("✅ ACCESO AUTORIZADO")
                                st.markdown(f"**Nombre:** {info['Nombre']} <br>**Asignación:** {info['Rol']} <br>**Equipo:** {info['Equipo']}", unsafe_allow_html=True)
                            elif status == "DUPLICADO":
                                st.warning("⚠️ ESTA MATRÍCULA YA INGRESÓ HOY.")
                            else:
                                st.error("❌ NO ENCONTRADO EN ESTE TORNEO.")
                        else:
                            st.error("Código ilegible.")
                
                st.subheader("📋 Lista de Accesos Registrados en Vivo")
                path_a = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_asistencia.csv")
                if os.path.exists(path_a):
                    df_as = pd.read_csv(path_a)
                    st.dataframe(df_as.sort_values(by="Hora", ascending=False), use_container_width=True)

    # --- ENTORNO 4: PANEL DE CONTROL DE CAMBIOS ADMIN ---
    elif modo == "🛡️ Base de Datos General (Admin)":
        st.header("🛡️ Consola de Administración Relacional")
        st.write("Monitoreo absoluto de las bases de datos locales.")
        
        if df_evs.empty:
            st.info("No hay registros almacenados.")
        else:
            for _, row in df_evs.iterrows():
                with st.container(border=True):
                    col_t, col_b = st.columns([3, 1])
                    with col_t:
                        st.markdown(f"#### 🏆 {row['nombre']}")
                        st.caption(f"ID del Sistema: `{row['id_evento']}` | Sede: {row['sede']}")
                    with col_b:
                        if st.button("🗑️ Eliminar Torneo de Raíz", key=row['id_evento'], type="secondary", use_container_width=True):
                            eliminar_evento_completo(row['id_evento'])
                            st.success(f"Torneo {row['id_evento']} eliminado.")
                            time.sleep(0.5)
                            st.rerun()