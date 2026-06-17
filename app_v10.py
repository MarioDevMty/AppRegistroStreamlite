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
import json

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

if "logged_in" not in st.session_state: st.session_state.logged_in = False
if "user_role" not in st.session_state: st.session_state.user_role = None
if "user_name" not in st.session_state: st.session_state.user_name = None
if "id_evento_activo" not in st.session_state: st.session_state.id_evento_activo = None

# ==========================================
# 2. FUNCIONES DE CORE Y PROCESAMIENTO
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
    for ext in [f"{id_evento}_master.xlsx", f"{id_evento}_asistencia.csv", f"{id_evento}_mapping.json"]:
        path = os.path.join(DB_DIR, ext)
        if os.path.exists(path): os.remove(path)

def limpiar_dato(dato):
    if pd.isna(dato): return ""
    
    # 1. Detectar y convertir formatos de fecha/Timestamp de Pandas
    if isinstance(dato, datetime) or hasattr(dato, 'strftime'):
        return dato.strftime('%Y-%m-%d')
        
    txt = str(dato).strip()
    if txt.lower() == "nan" or txt.lower() == "none" or txt == "": return ""
    
    # 2. Limpieza de flotantes espurios (.0) en matrículas
    if txt.endswith(".0"): 
        txt = txt[:-2]
        
    # 3. Convertir números de serie de Excel a fechas reales AAAA-MM-DD
    if txt.isdigit() and len(txt) == 5:
        try:
            num_serie = int(txt)
            fecha_real = pd.to_datetime(num_serie, unit='D', origin='1899-12-30')
            return fecha_real.strftime('%Y-%m-%d')
        except Exception:
            pass
            
    # 4. Limpieza si la fecha viene como texto con la hora pegada "2008-02-08 00:00:00"
    if " 00:00:00" in txt:
        txt = txt.replace(" 00:00:00", "")
        
    return txt

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
    except Exception: return None

def registrar_asistencia(codigo_leido, id_evento, df_master):
    archivo_log = os.path.join(DB_DIR, f"{id_evento}_asistencia.csv")
    if not os.path.exists(archivo_log):
        df_log = pd.DataFrame(columns=["Fecha", "Hora", "Codigo", "Nombre", "Rol", "Agrupacion"])
        df_log.to_csv(archivo_log, index=False)
    
    df_log = pd.read_csv(archivo_log)
    codigo_leido_str = str(codigo_leido).strip()
    
    if codigo_leido_str in df_log["Codigo"].astype(str).values:
        return "DUPLICADO", None

    persona_encontrada = None
    for _, row in df_master.iterrows():
        if limpiar_dato(row["Matrícula"]) == codigo_leido_str:
            persona_encontrada = row
            break

    if persona_encontrada is not None:
        nuevo_registro = {
            "Fecha": datetime.now().strftime("%Y-%m-%d"),
            "Hora": datetime.now().strftime("%H:%M:%S"),
            "Codigo": codigo_leido_str,
            "Nombre": limpiar_dato(persona_encontrada["Nombre"]),
            "Rol": f"Participante ({limpiar_dato(persona_encontrada['Categoría'])})",
            "Agrupacion": f"{limpiar_dato(persona_encontrada['Equipo'])} - {limpiar_dato(persona_encontrada['Procedencia'])}"
        }
        pd.DataFrame([nuevo_registro]).to_csv(archivo_log, mode='a', header=False, index=False)
        return "EXITO", nuevo_registro
    else:
        return "NO_ENCONTRADO", None

def procesar_zip_descarga(df):
    equipos_dict = {}
    for _, row in df.iterrows():
        prepa = limpiar_dato(row["Procedencia"])
        equipo = limpiar_dato(row["Equipo"])
        cat = limpiar_dato(row["Categoría"])
        mat = limpiar_dato(row["Matrícula"])
        nom = limpiar_dato(row["Nombre"])
        correo = limpiar_dato(row["Correo"])
        
        if not mat: continue
        key = f"{prepa}_{equipo}_{cat}"
        if key not in equipos_dict:
            # Reemplaza caracteres raros para la creación de carpetas físicas pero mantiene nombres entendibles
            nom_carpeta = "".join([c if c.isalnum() or c in " -_" else "-" for c in f"{prepa} {equipo} {cat}".strip()])
            equipos_dict[key] = {
                "Carpeta": nom_carpeta, "Equipo": equipo, "Correos": set([correo]) if correo else set(), "Imagenes": []
            }
        else:
            if correo: equipos_dict[key]["Correos"].add(correo)
        
        qr_bytes = generar_qr_bytes(mat)
        equipos_dict[key]["Imagenes"].append({
            "name": f"Acceso_{mat}_{nom}.png", "bytes": qr_bytes
        })
    return list(equipos_dict.values())

def generar_excel_reporte(df_vertical):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    
    sheet_todos = workbook.add_worksheet("Padrón General")
    headers = list(df_vertical.columns)
    for c, val in enumerate(headers): sheet_todos.write(0, c, str(val), header_fmt)
    for r, (_, row) in enumerate(df_vertical.iterrows(), start=1):
        for c, val in enumerate(row): sheet_todos.write(r, c, str(val))
        
    if "Categoría" in df_vertical.columns:
        categorias = df_vertical["Categoría"].dropna().unique()
        for cat in categorias:
            cat_clean = str(cat).strip()
            if not cat_clean or cat_clean == "General" or cat_clean == "": continue
            sheet_name = "".join([c for c in cat_clean if c.isalnum() or c in " "])[:30]
            sheet = workbook.add_worksheet(sheet_name)
            for c, val in enumerate(headers): sheet.write(0, c, str(val), header_fmt)
            df_cat = df_vertical[df_vertical["Categoría"].astype(str).str.strip() == cat_clean]
            for r, (_, row) in enumerate(df_cat.iterrows(), start=1):
                for c, val in enumerate(row): sheet.write(r, c, str(val))
    workbook.close()
    return output.getvalue()

# ==========================================
# 3. INTERFAZ GRÁFICA (STREAMLIT)
# ==========================================

if not st.session_state.logged_in:
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.write("")
        with st.container(border=True):
            st.markdown("<h3 style='text-align: center;'>🔐 Acceso al Sistema</h3>", unsafe_allow_html=True)
            usuario_input = st.text_input("Usuario")
            password_input = st.text_input("Contraseña", type="password")
            if st.button("Iniciar Sesión", type="primary", width='stretch'):
                user = usuario_input.strip().lower()
                if user in USUARIOS and USUARIOS[user]["password"] == password_input:
                    st.session_state.logged_in = True
                    st.session_state.user_role = USUARIOS[user]["role"]
                    st.session_state.user_name = USUARIOS[user]["nombre"]
                    st.rerun()
                else: st.error("Credenciales incorrectas.")
else:
    df_evs = obtener_eventos()
    ev_activo = None
    if st.session_state.id_evento_activo:
        match = df_evs[df_evs["id_evento"] == st.session_state.id_evento_activo]
        if not match.empty: ev_activo = match.iloc[0].to_dict()

    if ev_activo:
        st.markdown(f"<h2 style='text-align: center; color: #1E3A8A;'>💻 [EVENTO ACTIVO] {ev_activo['nombre']}</h2>", unsafe_allow_html=True)
    else:
        st.markdown("<h2 style='text-align: center; color: #DC2626;'>⚠️ Ningún Evento Seleccionado</h2>", unsafe_allow_html=True)
    st.markdown("---")

    st.sidebar.title("📲 Menú Principal")
    modo = st.sidebar.radio("Navegación:", ["📁 Historial / Selección", "✨ Crear Nuevo Evento", "📱 Escáner de Asistencia"] + (["🛡️ Base de Datos General (Admin)"] if st.session_state.user_role == "admin" else []))
    
    if st.sidebar.button("Cerrar Sesión", width='stretch'):
        st.session_state.logged_in = False
        st.session_state.id_evento_activo = None
        st.rerun()

    # --- ENTORNO 1: HISTORIAL ---
    if modo == "📁 Historial / Selección":
        st.header("🗂️ Listado Central de Eventos Creados")
        if df_evs.empty: st.info("No hay eventos creados.")
        else:
            opciones_select = {row["nombre"]: row["id_evento"] for _, row in df_evs.iterrows()}
            ev_elegido = st.selectbox("Elegir Evento para Activar:", list(opciones_select.keys()))
            if st.button("🚀 Activar Evento Seleccionado", type="primary"):
                st.session_state.id_evento_activo = opciones_select[ev_elegido]
                st.success(f"Entorno cambiado a: {ev_elegido}")
                time.sleep(0.5)
                st.rerun()
            st.dataframe(df_evs[["nombre", "fecha_inicio", "sede", "tipo"]], width='stretch')

    # --- ENTORNO 2: CREAR EVENTO E INGESTIÓN ---
    elif modo == "✨ Crear Nuevo Evento":
        st.header("⚡ Ficha Técnica y Carga de Competidores")
        tab_gen, tab_org, tab_xls = st.tabs(["📝 1. Ficha Técnica", "👥 2. Estructura Organizativa", "📊 3. Ingestión Inteligente de Excel"])
        
        with tab_gen:
            with st.container(border=True):
                n_nombre = st.text_input("Nombre del Evento:")
                c1, c2 = st.columns(2)
                with c1: n_fi = st.text_input("Fecha Inicio:", "02/04/2025")
                with c2: n_sed = st.text_input("Sede del Evento:", "Preparatoria UANL")
                n_dep = st.text_input("Dependencia Organizadora:", "UANL")
                n_dom = st.text_area("Domicilio Sede:")
                c3, c4 = st.columns(2)
                with c3: n_mag = st.selectbox("Magnitud:", ["Local", "Nacional"])
                with c4: n_tip = st.text_input("Tipo de Evento:", "Torneo")

        with tab_org:
            with st.container(border=True):
                n_com = st.text_input("Comité Organizador:", "CAD de TIC")
                n_res = st.text_area("Responsables del Evento:")
                if st.button("💾 Guardar Torneo", type="primary"):
                    if not n_nombre: st.error("Ingresa el nombre.")
                    else:
                        nuevo_id = guardar_nuevo_evento({
                            "nombre": n_nombre, "fecha_inicio": n_fi, "horario_inicio": "09:00", "fecha_fin": n_fi, "horario_fin": "18:00",
                            "dependencia": n_dep, "sede": n_sed, "domicilio": n_dom, "magnitud": n_mag, "dirigido": "Estudiantes",
                            "clasificacion": "Académico", "tipo": n_tip, "comite": n_com, "responsables": n_res
                        })
                        st.session_state.id_evento_activo = nuevo_id
                        st.success("¡Guardado!")
                        time.sleep(0.5)
                        st.rerun()

        with tab_xls:
            if not ev_activo: st.warning("⚠️ Primero activa o crea un torneo.")
            else:
                st.markdown(f"#### Configuración del padrón de: **{ev_activo['nombre']}**")
                
                tipo_formato = st.radio("📐 Estructura del Excel cargado:", [
                    "Vertical (Cada fila representa a un participante individual)",
                    "Horizontal por Equipos (Cada fila representa un equipo completo con múltiples columnas de alumnos)"
                ])
                
                up_file = st.file_uploader("Cargar Archivo de Respuestas (.xlsx)", type=["xlsx"])
                
                if up_file:
                    df_raw = pd.read_excel(up_file)
                    
                    for col in df_raw.columns:
                        if df_raw[col].dtype == 'object':
                            df_raw[col] = df_raw[col].apply(lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if hasattr(x, 'strftime') else x)
                        elif pd.api.types.is_datetime64_any_dtype(df_raw[col]):
                            df_raw[col] = df_raw[col].astype(str)
                    
                    columnas_raw = list(df_raw.columns)
                    
                    with st.container(border=True):
                        st.markdown("### 🧩 Mapeador Dinámico de Campos")
                        st.write("Identifica las columnas raíz para la procedencia y clasificación del grupo:")
                        
                        c_a1, c_a2, c_a3 = st.columns(3)
                        with c_a1: m_prepa = st.selectbox("🏫 Columna de la Preparatoria / Sede / Colegio:", columnas_raw)
                        with c_a2: m_equipo = st.selectbox("👥 Columna Nombre del Equipo (Elige la misma de escuela si es individual):", options=["No aplica"] + columnas_raw)
                        with c_a3: m_cat = st.selectbox("🏆 Columna de la Categoría:", options=["No aplica"] + columnas_raw)
                        
                        st.markdown("---")
                        
                        # --- FORMATEADOR VERTICAL ---
                        if "Vertical" in tipo_formato:
                            st.write("**Mapeo de Filas de Alumno Único:**")
                            c_v1, c_v2 = st.columns(2)
                            with c_v1: m_mat = st.selectbox("🔑 Columna Matrícula:", columnas_raw)
                            with c_v2: m_cor = st.selectbox("📧 Columna Correo:", options=["No aplica"] + columnas_raw)
                            
                            st.write("👤 **Segmentación del Nombre:**")
                            c_n1, c_n2, c_n3 = st.columns(3)
                            with c_n1: m_pat = st.selectbox("Columna Apellido Paterno:", columnas_raw)
                            with c_n2: m_mat_n = st.selectbox("Columna Apellido Materno:", columnas_raw)
                            with c_n3: m_nom_n = st.selectbox("Columna Nombre(s):", columnas_raw)
                            
                            st.write("📋 **Atributos Escolares Integrados:**")
                            c_ex1, c_ex2 = st.columns(2)
                            with c_ex1: m_f_nac = st.selectbox("Columna Fecha de Nacimiento:", options=["No aplica"] + columnas_raw)
                            with c_ex2: m_sem = st.selectbox("Columna Semestre:", options=["No aplica"] + columnas_raw)
                            
                            if st.button("🔗 Procesar y Verticalizar Padrón (Modo Vertical)", type="primary", width='stretch'):
                                rows_verticales = []
                                for _, row in df_raw.iterrows():
                                    mat_val = limpiar_dato(row[m_mat])
                                    pat_val = limpiar_dato(row[m_pat])
                                    nom_n_val = limpiar_dato(row[m_nom_n])
                                    proc_val = limpiar_dato(row[m_prepa])
                                    
                                    # Solo exige Matrícula, Nombre e Institución real (sea cual sea su nombre)
                                    if mat_val == "" or nom_n_val == "" or proc_val == "": 
                                        continue
                                    
                                    nombre_completo = f"{pat_val} {limpiar_dato(row[m_mat_n])} {nom_n_val}".strip()
                                    nombre_completo = re.sub(r'\s+', ' ', nombre_completo)
                                    
                                    data_fila = {
                                        "Procedencia": proc_val,
                                        "Equipo": "Individual" if m_equipo == "No aplica" else limpiar_dato(row[m_equipo]),
                                        "Categoría": "General" if m_cat == "No aplica" else limpiar_dato(row[m_cat]),
                                        "Matrícula": mat_val,
                                        "Nombre": nombre_completo,
                                        "Correo": "" if m_cor == "No aplica" else limpiar_dato(row[m_cor]),
                                        "Fecha de Nacimiento": "" if m_f_nac == "No aplica" else limpiar_dato(row[m_f_nac]),
                                        "Semestre": "" if m_sem == "No aplica" else limpiar_dato(row[m_sem])
                                    }
                                    rows_verticales.append(data_fila)
                                
                                df_vertical = pd.DataFrame(rows_verticales)
                                path_destino = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_master.xlsx")
                                df_vertical.to_excel(path_destino, index=False)
                                
                                with open(os.path.join(DB_DIR, f"{ev_activo['id_evento']}_mapping.json"), "w") as f:
                                    json.dump({"mapeado": True}, f)
                                st.success("¡Padrón indexado de forma correcta!")
                                time.sleep(0.5)
                                st.rerun()
                                
                        # --- FORMATEADOR HORIZONTAL POR EQUIPOS ---
                        else:
                            st.write("🔬 **Mapeo Dinámico por Listas de Integrantes:**")
                            
                            m_mats = st.multiselect("🔍 Columnas de MATRÍCULAS:", options=columnas_raw)
                            m_pats = st.multiselect("🔹 Columnas de APELLIDO PATERNO:", options=columnas_raw)
                            m_mats_n = st.multiselect("🔹 Columnas de APELLIDO MATERNO (Opcional):", options=columnas_raw)
                            m_noms_n = st.multiselect("🔹 Columnas de NOMBRE(S):", options=columnas_raw)
                            
                            m_cors = st.multiselect("📧 Columnas de CORREOS (Opcional):", options=columnas_raw)
                            m_fnacs = st.multiselect("📅 Columnas de FECHA DE NACIMIENTO:", options=columnas_raw)
                            m_sems = st.multiselect("🏫 Columnas de SEMESTRE:", options=columnas_raw)
                            
                            if st.button("🌀 Aplanar y Normalizar Base de Datos Completa", type="primary", width='stretch'):
                                if not m_mats or not m_pats or not m_noms_n:
                                    st.error("⚠️ Debes rellenar las columnas requeridas (Matrículas, Paternos y Nombres).")
                                elif len(m_mats) != len(m_pats) or len(m_mats) != len(m_noms_n):
                                    st.error("⚠️ Desajuste: Las listas de Matrículas, Paternos y Nombres deben tener la misma cantidad de columnas.")
                                else:
                                    rows_verticales = []
                                    
                                    for _, row in df_raw.iterrows():
                                        proc_val = limpiar_dato(row[m_prepa])
                                        eq_val = "Individual" if m_equipo == "No aplica" else limpiar_dato(row[m_equipo])
                                        cat_val = "General" if m_cat == "No aplica" else limpiar_dato(row[m_cat])
                                        
                                        # Si la fila entera no tiene escuela asignada, se descarta
                                        if proc_val == "":
                                            continue
                                            
                                        for i in range(len(m_mats)):
                                            mat_alumno = limpiar_dato(row.get(m_mats[i], ""))
                                            pat_alumno = limpiar_dato(row.get(m_pats[i], ""))
                                            nom_alumno = limpiar_dato(row.get(m_noms_n[i], ""))
                                            
                                            # El filtro evalúa únicamente la existencia real del estudiante en la ranura
                                            if mat_alumno == "" or nom_alumno == "":
                                                continue
                                            
                                            mat_n_alumno = limpiar_dato(row.get(m_mats_n[i], "")) if i < len(m_mats_n) else ""
                                            nombre_completo = f"{pat_alumno} {mat_n_alumno} {nom_alumno}".strip()
                                            nombre_completo = re.sub(r'\s+', ' ', nombre_completo)
                                            
                                            cor_alumno = limpiar_dato(row.get(m_cors[i], "")) if i < len(m_cors) else ""
                                            fnac_alumno = limpiar_dato(row.get(m_fnacs[i], "")) if i < len(m_fnacs) else ""
                                            sem_alumno = limpiar_dato(row.get(m_sems[i], "")) if i < len(m_sems) else ""
                                            
                                            data_final_alumno = {
                                                "Procedencia": proc_val, # Guarda "CIDEB", "Don Bosco", etc. de forma íntegra
                                                "Equipo": eq_val,
                                                "Categoría": cat_val,
                                                "Matrícula": mat_alumno,
                                                "Nombre": nombre_completo,
                                                "Correo": cor_alumno,
                                                "Fecha de Nacimiento": fnac_alumno,
                                                "Semestre": sem_alumno
                                            }
                                            rows_verticales.append(data_final_alumno)
                                    
                                    df_final_vertical = pd.DataFrame(rows_verticales)
                                    path_destino = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_master.xlsx")
                                    df_final_vertical.to_excel(path_destino, index=False)
                                    
                                    with open(os.path.join(DB_DIR, f"{ev_activo['id_evento']}_mapping.json"), "w") as f:
                                        json.dump({"mapeado": True}, f)
                                        
                                    st.success(f"¡Éxito! Se procesaron {len(df_final_vertical)} alumnos reales incluyendo todas las Preparatorias, Colegios y Centros sin distinción.")
                                    time.sleep(0.5)
                                    st.rerun()

                # --- ZONA DE DESCARGAS CENTRALIZADA ---
                path_master_existente = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_master.xlsx")
                path_map_existente = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_mapping.json")
                if os.path.exists(path_master_existente) and os.path.exists(path_map_existente):
                    st.write("---")
                    st.markdown("### 📦 Descargas Estructuradas Disponibles")
                    df_m = pd.read_excel(path_master_existente)
                    
                    c_d1, c_d2 = st.columns(2)
                    with c_d1:
                        if st.button("📦 Generar y Descargar Paquete de Códigos QR"):
                            datos_proc = procesar_zip_descarga(df_m)
                            b = io.BytesIO()
                            with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
                                for eq in datos_proc:
                                    for img in eq["Imagenes"]:
                                        z.writestr(f"{eq['Carpeta']}/{img['name']}", img['bytes'])
                            st.download_button("⬇️ Descargar ZIP Completo", b.getvalue(), f"QRs_{ev_activo['id_evento']}.zip", "application/zip", width='stretch')
                    with c_d2:
                        ex_b = generar_excel_reporte(df_m)
                        st.download_button("📥 Descargar Reporte Segmentado por Pestañas", ex_b, f"Reporte_{ev_activo['id_evento']}.xlsx", width='stretch')
                    st.dataframe(df_m, width='stretch')

    # --- ENTORNO 3: ESCÁNER ---
    elif modo == "📱 Escáner de Asistencia":
        if not ev_activo: st.error("⚠️ No hay evento activo.")
        else:
            path_m = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_master.xlsx")
            if not os.path.exists(path_m): st.warning("⚠️ Carga la lista primero en 'Crear Nuevo Evento'.")
            else:
                df_master = pd.read_excel(path_m)
                c_cam, c_res = st.columns([1, 2])
                with c_cam: img_buffer = st.camera_input("Enfoca el QR")
                with c_res:
                    if img_buffer is not None:
                        codigo = leer_qr_desde_imagen(img_buffer)
                        if codigo:
                            status, info = registrar_asistencia(codigo, ev_activo['id_evento'], df_master)
                            if status == "EXITO":
                                st.success("✅ ACCESO CONFIRMADO")
                                st.markdown(f"**Nombre:** {info['Nombre']} <br>**Asignación:** {info['Rol']} <br>**Agrupación:** {info['Agrupacion']}", unsafe_allow_html=True)
                            elif status == "DUPLICADO": st.warning("⚠️ ESTA MATRÍCULA YA INGRESÓ.")
                            else: st.error("❌ NO REGISTRADO EN ESTE TORNEO.")
                        else: st.error("QR ilegible.")
                
                st.subheader("📋 Accesos en Tiempo Real")
                path_a = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_asistencia.csv")
                if os.path.exists(path_a):
                    st.dataframe(pd.read_csv(path_a).sort_values(by="Hora", ascending=False), width='stretch')

    # --- ENTORNO 4: ADMIN ---
    elif modo == "🛡️ Base de Datos General (Admin)":
        st.header("🛡️ Consola Administrativa General")
        if df_evs.empty: st.info("Sin registros.")
        else:
            for _, row in df_evs.iterrows():
                with st.container(border=True):
                    col_t, col_b = st.columns([3, 1])
                    with col_t:
                        st.markdown(f"#### 🏆 {row['nombre']}")
                        st.caption(f"ID Interno: `{row['id_evento']}` | Sede: {row['sede']}")
                    with col_b:
                        if st.button("🗑️ Eliminar de Raíz", key=row['id_evento'], type="secondary", width='stretch'):
                            eliminar_evento_completo(row['id_evento'])
                            st.success("Purgado.")
                            time.sleep(0.5)
                            st.rerun()