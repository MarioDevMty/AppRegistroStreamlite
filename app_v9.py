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
# 2. FUNCIONES DE BACKEND Y BASES DE DATOS
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

def registrar_asistencia(codigo_leido, id_evento, df_master, mapping):
    archivo_log = os.path.join(DB_DIR, f"{id_evento}_asistencia.csv")
    if not os.path.exists(archivo_log):
        df_log = pd.DataFrame(columns=["Fecha", "Hora", "Codigo", "Nombre", "Rol", "Relacion_Agrupador"])
        df_log.to_csv(archivo_log, index=False)
    
    df_log = pd.read_csv(archivo_log)
    codigo_leido_str = str(codigo_leido).strip()
    
    if codigo_leido_str in df_log["Codigo"].astype(str).values:
        return "DUPLICADO", None

    # Búsqueda dinámica basada en la columna mapeada como Identificador/Matrícula
    col_id = mapping["col_id"]
    persona_encontrada = None
    
    for _, row in df_master.iterrows():
        if limpiar_dato(row[col_id]) == codigo_leido_str:
            persona_encontrada = row
            break

    if persona_encontrada is not None:
        nombre = limpiar_dato(persona_encontrada[mapping["col_nombre"]])
        prepa = limpiar_dato(persona_encontrada[mapping["col_prepa"]])
        
        # Elementos opcionales de relación
        rol_info = "Participante"
        if mapping.get("col_categoria") and mapping["col_categoria"] in persona_encontrada:
            rol_info += f" ({limpiar_dato(persona_encontrada[mapping['col_categoria']])})"
            
        agrupador = f"{prepa}"
        if mapping.get("col_equipo") and mapping["col_equipo"] in persona_encontrada:
            eq_nome = limpiar_dato(persona_encontrada[mapping["col_equipo"]])
            if comedy_name := eq_nome: agrupador = f"Equipo: {comedy_name} ({prepa})"
            
        nuevo_registro = {
            "Fecha": datetime.now().strftime("%Y-%m-%d"),
            "Hora": datetime.now().strftime("%H:%M:%S"),
            "Codigo": codigo_leido_str,
            "Nombre": nombre,
            "Rol": rol_info,
            "Relacion_Agrupador": agrupador
        }
        pd.DataFrame([nuevo_registro]).to_csv(archivo_log, mode='a', header=False, index=False)
        return "EXITO", nuevo_registro
    else:
        return "NO_ENCONTRADO", None

def procesar_zip_correo(df, mapping):
    equipos_dict = {}
    col_id = mapping["col_id"]
    col_nom = mapping["col_nombre"]
    col_prepa = mapping["col_prepa"]
    col_mail = mapping.get("col_correo")
    col_eq = mapping.get("col_equipo")
    col_cat = mapping.get("col_categoria")

    for _, row in df.iterrows():
        mat = limpiar_dato(row[col_id])
        nom = limpiar_dato(row[col_nom])
        prepa = limpiar_dato(row[col_prepa])
        correo = limpiar_dato(row[col_mail]) if col_mail else ""
        equipo = limpiar_dato(row[col_eq]) if col_eq else "Individual"
        cat = limpiar_dato(row[col_cat]) if col_cat else "General"
        
        if not mat: continue
        
        # La clave de empaquetado depende de si es por equipo o individual por escuela
        key = f"{prepa}_{equipo}_{cat}"
        if key not in equipos_dict:
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
        
    lista_resultado = []
    for key, info in equipos_dict.items():
        info["Correo"] = ", ".join(info["Correos"]) if info["Correos"] else ""
        lista_resultado.append(info)
    return lista_resultado

def generar_excel_resumen_dinamico(df_original, mapping):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    
    # Pestaña 1: Datos Base
    sheet_todos = workbook.add_worksheet("Registros Completos")
    headers = list(df_original.columns)
    for c, val in enumerate(headers): sheet_todos.write(0, c, str(val), header_fmt)
    for r, (_, row) in enumerate(df_original.iterrows(), start=1):
        for c, val in enumerate(row): sheet_todos.write(r, c, str(val))
    
    # Segmentación Inteligente: Si hay categorías segmenta por categoría, si no, segmenta por Escuela (Preparatoria)
    col_segmentar = mapping.get("col_categoria") if mapping.get("col_categoria") else mapping["col_prepa"]
    
    if col_segmentar in df_original.columns:
        bloques = df_original[col_segmentar].dropna().unique()
        for blk in bloques:
            blk_clean = str(blk).strip()
            if not blk_clean: continue
            sheet_name = "".join([c for c in blk_clean if c.isalnum() or c in " "])[:30]
            sheet = workbook.add_worksheet(sheet_name)
            
            for c, val in enumerate(headers): sheet.write(0, c, str(val), header_fmt)
            df_filter = df_original[df_original[col_segmentar].astype(str).str.strip() == blk_clean]
            for r, (_, row) in enumerate(df_filter.iterrows(), start=1):
                for c, val in enumerate(row): sheet.write(r, c, str(val))
                
    workbook.close()
    
    # Conteo de entidades únicas (Equipos si existen, si no, total de participantes)
    total_unidades = df_original[mapping["col_equipo"]].nunique() if mapping.get("col_equipo") else len(df_original)
    return output.getvalue(), total_unidades

# ==========================================
# 3. INTERFAZ DE USUARIO (STREAMLIT)
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
    mapping_activo = None

    if st.session_state.id_evento_activo:
        match = df_evs[df_evs["id_evento"] == st.session_state.id_evento_activo]
        if not match.empty:
            ev_activo = match.iloc[0].to_dict()
            # Cargar mapeo si existe
            path_map = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_mapping.json")
            if os.path.exists(path_map):
                with open(path_map, "r", encoding="utf-8") as f:
                    mapping_activo = json.load(f)

    # Banner Superior
    if ev_activo:
        st.markdown(f"<h2 style='text-align: center; color: #1E3A8A;'>💻 [EVENTO] {ev_activo['nombre']}</h2>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align: center; color: #6B7280;'>Sede: {ev_activo['sede']} | Modalidad estructurada dinámicamente</p>", unsafe_allow_html=True)
    else:
        st.markdown("<h2 style='text-align: center; color: #DC2626;'>⚠️ Ningún Evento Seleccionado</h2>", unsafe_allow_html=True)
        st.markdown("---")

    # Sidebar Menu
    st.sidebar.title("📲 Menú Principal")
    modo = st.sidebar.radio("Navegación:", ["📁 Historial / Selección", "✨ Crear Nuevo Evento", "📱 Escáner de Asistencia"] + (["🛡️ Base de Datos General (Admin)"] if st.session_state.user_role == "admin" else []))
    
    st.sidebar.markdown("---")
    if st.sidebar.button("Cerrar Sesión", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.id_evento_activo = None
        st.rerun()

    # --- RUTA 1: HISTORIAL ---
    if modo == "📁 Historial / Selección":
        st.header("🗂️ Listado Central de Eventos Creados")
        if df_evs.empty:
            st.info("No hay eventos creados todavía.")
        else:
            opciones_select = {row["nombre"]: row["id_evento"] for _, row in df_evs.iterrows()}
            ev_elegido = st.selectbox("Elegir Evento para Activar:", list(opciones_select.keys()))
            
            if st.button("🚀 Activar Evento Seleccionado", type="primary"):
                st.session_state.id_evento_activo = opciones_select[ev_elegido]
                st.success(f"Entorno cambiado a: {ev_elegido}")
                time.sleep(0.5)
                st.rerun()
            st.dataframe(df_evs[["nombre", "fecha_inicio", "sede", "tipo"]], use_container_width=True)

    # --- RUTA 2: CREAR Y CONFIGURAR EVENTO ---
    elif modo == "✨ Crear Nuevo Evento":
        st.header("⚡ Configuración Avanzada y Registro de Fichas")
        tab_gen, tab_org, tab_xls = st.tabs(["📝 1. Ficha Técnica", "👥 2. Estructura Organizativa", "📊 3. Ingestión y Mapeo de Excel"])
        
        with tab_gen:
            with st.container(border=True):
                n_nombre = st.text_input("Nombre del Evento:")
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
                n_res = st.text_area("Responsables del Evento:", "Diana Janeth Amaro Fernández, Alejandro Ojeda Ramírez")
                
                if st.button("💾 Crear Evento en Base de Datos", type="primary"):
                    if not n_nombre: st.error("Falta el nombre.")
                    else:
                        nuevo_id = guardar_nuevo_evento({
                            "nombre": n_nombre, "fecha_inicio": n_fi, "horario_inicio": n_hi, "fecha_fin": n_ff, "horario_fin": n_hf,
                            "dependencia": n_dep, "sede": n_sed, "domicilio": n_dom, "magnitud": n_mag, "dirigido": n_dir,
                            "clasificacion": n_cla, "tipo": n_tip, "comite": n_com, "responsables": n_res
                        })
                        st.session_state.id_evento_activo = nuevo_id
                        st.success("¡Evento guardado!")
                        time.sleep(0.5)
                        st.rerun()

        with tab_xls:
            if not ev_activo:
                st.warning("⚠️ Selecciona o crea un evento primero.")
            else:
                st.markdown(f"#### Ingestión de datos para: **{ev_activo['nombre']}**")
                up_file = st.file_uploader("Cargar cualquier Excel de Registro (.xlsx)", type=["xlsx"])
                
                if up_file:
                    df_preview = pd.read_excel(up_file)
                    path_destino = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_master.xlsx")
                    df_preview.to_excel(path_destino, index=False)
                    st.success("Archivo base guardado. Procede a configurar las columnas abajo.")
                    st.rerun()
                
                path_master_existente = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_master.xlsx")
                if os.path.exists(path_master_existente):
                    df_m = pd.read_excel(path_master_existente)
                    columnas_disponibles = list(df_m.columns)
                    
                    st.write("---")
                    st.markdown("### 🧩 Pre-selección e Inteligencia de Campos (Mapeo Relacional)")
                    st.caption("Dile al sistema qué columna de tu Excel representa cada entidad para poder procesarlo:")
                    
                    # Rellenar con lo existente si ya se configuró previamente
                    def_id = columnas_disponibles.index(mapping_activo["col_id"]) if mapping_activo and mapping_activo["col_id"] in columnas_disponibles else 0
                    def_nom = columnas_disponibles.index(mapping_activo["col_nombre"]) if mapping_activo and mapping_activo["col_nombre"] in columnas_disponibles else 0
                    def_pre = columnas_disponibles.index(mapping_activo["col_prepa"]) if mapping_activo and mapping_activo["col_prepa"] in columnas_disponibles else 0
                    
                    c_m1, c_m2, c_m3 = st.columns(3)
                    with c_m1: m_id = st.selectbox("🔑 Identificador Único / Matrícula (Para generar QR):", columnas_disponibles, index=def_id)
                    with c_m2: m_nom = st.selectbox("👤 Nombre Completo del Participante:", columnas_disponibles, index=def_nom)
                    with c_m3: m_pre = st.selectbox("🏫 Escuela / Dependencia / Procedencia:", columnas_disponibles, index=def_pre)
                    
                    st.write("**Campos Opcionales (Déjalos en 'No aplica' si el evento no los contempla):**")
                    opciones_con_vacio = ["No aplica"] + columnas_disponibles
                    
                    def_eq = opciones_con_vacio.index(mapping_activo["col_equipo"]) if mapping_activo and mapping_activo.get("col_equipo") in opciones_con_vacio else 0
                    def_cat = opciones_con_vacio.index(mapping_activo["col_categoria"]) if mapping_activo and mapping_activo.get("col_categoria") in opciones_con_vacio else 0
                    def_cor = opciones_con_vacio.index(mapping_activo["col_correo"]) if mapping_activo and mapping_activo.get("col_correo") in opciones_con_vacio else 0
                    
                    c_m4, c_m5, c_m6 = st.columns(3)
                    with c_m4: m_eq = st.selectbox("👥 Nombre del Equipo / Grupo:", opciones_con_vacio, index=def_eq)
                    with c_m5: m_cat = st.selectbox("🏆 Categoría / Nivel / Rama del Torneo:", opciones_con_vacio, index=def_cat)
                    with c_m6: m_cor = st.selectbox("📧 Correo Electrónico (Para notificaciones):", opciones_con_vacio, index=def_cor)
                    
                    if st.button("🔗 Confirmar y Guardar Mapeo del Padrón", type="primary"):
                        map_data = {
                            "col_id": m_id, "col_nombre": m_nom, "col_prepa": m_pre,
                            "col_equipo": None if m_eq == "No aplica" else m_eq,
                            "col_categoria": None if m_cat == "No aplica" else m_cat,
                            "col_correo": None if m_cor == "No aplica" else m_cor
                        }
                        path_map_dest = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_mapping.json")
                        with open(path_map_dest, "w", encoding="utf-8") as f:
                            json.dump(map_data, f, ensure_ascii=False, indent=4)
                        st.success("¡Estructura mapeada con éxito!")
                        time.sleep(0.5)
                        st.rerun()
                        
                    if mapping_activo:
                        st.write("---")
                        st.markdown("### 📦 Descargas Disponibles para este Torneo")
                        df_limpio = df_m.dropna(subset=[mapping_activo["col_id"]])
                        datos_proc = procesar_zip_correo(df_limpio, mapping_activo)
                        
                        col_ex1, col_ex2 = st.columns(2)
                        with col_ex1:
                            if st.button("📦 Descargar Paquete ZIP de QRs"):
                                b = io.BytesIO()
                                with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
                                    for eq in datos_proc:
                                        for img in eq["Imagenes"]:
                                            z.writestr(f"{eq['Carpeta']}/{img['name']}", img['bytes'])
                                st.download_button("⬇️ Guardar ZIP", b.getvalue(), f"QRs_{ev_activo['id_evento']}.zip", "application/zip", width='content')
                        with col_ex2:
                            ex_b, t_unidades = generar_excel_resumen_dinamico(df_limpio, mapping_activo)
                            label_met = "Equipos Totales" if mapping_activo["col_equipo"] else "Participantes Totales"
                            st.metric(label=label_met, value=t_unidades)
                            st.download_button("📥 Descargar Reporte Consolidado", ex_b, f"Reporte_{ev_activo['id_evento']}.xlsx", width='content')

    # --- RUTA 3: ESCÁNER ---
    elif modo == "📱 Escáner de Asistencia":
        if not ev_activo or not mapping_activo:
            st.error("⚠️ Evento activo sin mapear. Configura la sección 'Crear Nuevo Evento -> Pestaña 3' primero.")
        else:
            path_m = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_master.xlsx")
            df_master = pd.read_excel(path_m).dropna(subset=[mapping_activo["col_id"]])
            
            c_cam, c_res = st.columns([1, 2])
            with c_cam: img_buffer = st.camera_input("Enfoca el código QR")
            with c_res:
                if img_buffer is not None:
                    codigo = leer_qr_desde_imagen(img_buffer)
                    if codigo:
                        status, info = registrar_asistencia(codigo, ev_activo['id_evento'], df_master, mapping_activo)
                        if status == "EXITO":
                            st.success("✅ ENTRADA REGISTRADA")
                            st.markdown(f"**Nombre:** {info['Nombre']} <br>**Detalles:** {info['Rol']} <br>**Entidad Relacionada:** {info['Relacion_Agrupador']}", unsafe_allow_html=True)
                        elif status == "DUPLICADO": st.warning("⚠️ ACCESO DUPLICADO RECHAZADO.")
                        else: st.error("❌ MATRÍCULA NO ENCONTRADA EN ESTE TORNEO.")
                    else: st.error("Código QR no detectado.")
            
            st.subheader("📋 Accesos Registrados de este Torneo")
            path_a = os.path.join(DB_DIR, f"{ev_activo['id_evento']}_asistencia.csv")
            if os.path.exists(path_a):
                df_as = pd.read_csv(path_a)
                st.dataframe(df_as.sort_values(by="Hora", ascending=False), width='content')

    # --- RUTA 4: ADMIN ---
    elif modo == "🛡️ Base de Datos General (Admin)":
        st.header("🛡️ Consola de Administración Relacional")
        if df_evs.empty: st.info("Sin registros.")
        else:
            for _, row in df_evs.iterrows():
                with st.container(border=True):
                    col_t, col_b = st.columns([3, 1])
                    with col_t:
                        st.markdown(f"#### 🏆 {row['nombre']}")
                        st.caption(f"ID del Sistema: `{row['id_evento']}` | Sede: {row['sede']}")
                    with col_b:
                        if st.button("🗑️ Eliminar Torneo de Raíz", key=row['id_evento'], type="secondary", width='content'):
                            eliminar_evento_completo(row['id_evento'])
                            st.success("Torneo eliminado.")
                            time.sleep(0.5)
                            st.rerun()