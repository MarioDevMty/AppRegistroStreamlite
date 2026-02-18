import streamlit as st
import pandas as pd
import qrcode
import io
import zipfile
import smtplib
from email.message import EmailMessage
import time
import xlsxwriter

st.set_page_config(page_title="Gestor Torneo - QRs y Reportes", page_icon="📊", layout="wide")

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
    """Genera la imagen QR."""
    qr = qrcode.QRCode(box_size=10, border=4)
    qr.add_data(dato)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()

def cargar_dataframe(uploaded_file):
    """Carga Excel, rellena celdas combinadas y prepara datos."""
    try:
        # Leemos sin header para usar indices numéricos absolutos
        df = pd.read_excel(uploaded_file, engine='openpyxl', header=None)
        # Rellenar celdas combinadas (ffill) hacia abajo
        df = df.ffill()
        # Quitar la primera fila que suelen ser los encabezados textuales
        df = df.iloc[1:].reset_index(drop=True)
        return df
    except Exception as e:
        st.error(f"Error leyendo el archivo: {e}")
        return None

# --- LÓGICA CORE: TRASPONER DATOS (HORIZONTAL A VERTICAL) ---

def generar_excel_resumen(df_original):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Estilos
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    
    # ---------------------------------------------------------
    # 1. HOJA DE ASESORES (COACHES)
    # ---------------------------------------------------------
    sheet_asesores = workbook.add_worksheet("Asesores")
    cols_asesor = ["Escuela", "Nombre", "Ap. Paterno", "Ap. Materno", "Celular", "Correo"]
    
    for c, val in enumerate(cols_asesor):
        sheet_asesores.write(0, c, val, header_fmt)
        
    asesores_unicos = set()
    row_asesor = 1
    
    # Indices Asesor: B=1, F=5, G=6, H=7, I=8, J=9
    for _, row in df_original.iterrows():
        nombre = limpiar_dato(row.iloc[5])
        celular = limpiar_dato(row.iloc[8])
        
        # Llave única para evitar duplicados
        if nombre and (nombre, celular) not in asesores_unicos:
            asesores_unicos.add((nombre, celular))
            datos = [
                limpiar_dato(row.iloc[1]), # Escuela
                nombre,
                limpiar_dato(row.iloc[6]), # Pat
                limpiar_dato(row.iloc[7]), # Mat
                celular,
                limpiar_dato(row.iloc[9])  # Correo
            ]
            for c, val in enumerate(datos):
                sheet_asesores.write(row_asesor, c, val)
            row_asesor += 1

    # ---------------------------------------------------------
    # 2. HOJAS DE ALUMNOS (LÓGICA DE TRANSPOSICIÓN)
    # ---------------------------------------------------------
    
    # Definición de las 5 posiciones posibles de alumnos en el Excel horizontal
    # Formato: [Matricula, Paterno, Materno, Nombre, Correo]
    config_posiciones = [
        [10, 11, 12, 13, 16], # Alumno 1
        [17, 18, 19, 20, 23], # Alumno 2
        [24, 25, 26, 27, 30], # Alumno 3
        [31, 32, 33, 34, 37], # Alumno 4
        [39, 40, 41, 42, 45]  # Alumno 5
    ]
    
    headers_alumno = ["Escuela", "Equipo", "Categoría", "Matrícula", "Ap. Paterno", "Ap. Materno", "Nombre", "Correo Inst."]

    # Crear las hojas y contadores de fila
    sheets = {
        "Línea":      {"obj": workbook.add_worksheet("Línea"), "row": 1, "max": 4},
        "Laberinto":  {"obj": workbook.add_worksheet("Laberinto"), "row": 1, "max": 4},
        "Escenario":  {"obj": workbook.add_worksheet("Escenario"), "row": 1, "max": 5},
    }

    # Escribir encabezados en todas las hojas
    for key in sheets:
        for c, val in enumerate(headers_alumno):
            sheets[key]["obj"].write(0, c, val, header_fmt)

    # --- BARRIDO FILA POR FILA (EQUIPOS) ---
    for _, row in df_original.iterrows():
        # BLOQUE 1: DATOS REPETITIVOS (Escuela, Equipo, Categoria)
        escuela = limpiar_dato(row.iloc[1])
        equipo = limpiar_dato(row.iloc[3])
        categoria_txt = limpiar_dato(row.iloc[4])
        
        if not escuela or not equipo: continue

        # Identificar hoja destino
        cat_lower = categoria_txt.lower()
        target = None
        if "línea" in cat_lower or "linea" in cat_lower: target = "Línea"
        elif "laberinto" in cat_lower: target = "Laberinto"
        elif "escenario" in cat_lower: target = "Escenario"
        
        if target:
            hoja_info = sheets[target]
            worksheet = hoja_info["obj"]
            max_alumnos = hoja_info["max"]
            
            # BLOQUE 2: ITERAR COLUMNAS DE ALUMNOS (TRANSPOSICIÓN)
            # Recorremos del alumno 0 al max permitido (4 o 5)
            for i in range(max_alumnos):
                indices = config_posiciones[i]
                
                # Verificar que el índice existe en el excel
                if indices[0] < len(row):
                    matricula = limpiar_dato(row.iloc[indices[0]])
                    
                    # ¡AQUÍ ESTÁ LA CLAVE! 
                    # Si existe matrícula, creamos una NUEVA FILA en el reporte
                    if matricula:
                        paterno = limpiar_dato(row.iloc[indices[1]])
                        materno = limpiar_dato(row.iloc[indices[2]])
                        nombre = limpiar_dato(row.iloc[indices[3]])
                        correo = limpiar_dato(row.iloc[indices[4]])
                        
                        datos_fila = [escuela, equipo, target, matricula, paterno, materno, nombre, correo]
                        
                        # Escribir fila en la hoja correspondiente
                        fila_actual = hoja_info["row"]
                        for col_idx, valor in enumerate(datos_fila):
                            worksheet.write(fila_actual, col_idx, valor)
                        
                        # Avanzar contador de fila para el siguiente alumno (del mismo o diferente equipo)
                        sheets[target]["row"] += 1

    workbook.close()
    return output.getvalue(), len(asesores_unicos)

def procesar_logica_zip_correo(df):
    """Extrae estructura para ZIP y Correos."""
    equipos = []
    cols_mat = [10, 17, 24, 31, 39] 

    for _, row in df.iterrows():
        escuela = limpiar_dato(row.iloc[1])
        equipo = limpiar_dato(row.iloc[3])
        categoria = limpiar_dato(row.iloc[4])
        
        if not escuela or not equipo: continue

        nombre_carpeta = f"{escuela} {equipo} {categoria}".strip()
        nombre_carpeta = "".join([c if c.isalnum() or c in " -_" else "-" for c in nombre_carpeta])

        celular_asesor = limpiar_dato(row.iloc[8])
        correo_asesor = limpiar_dato(row.iloc[9])
        
        imgs = []
        if celular_asesor:
            imgs.append({"nombre_archivo": f"Coach_{celular_asesor}.png", "bytes": generar_qr_bytes(celular_asesor)})

        max_al = 5 if "escenario" in str(categoria).lower() else 4
        for i, col_idx in enumerate(cols_mat):
            if i >= max_al: break
            if col_idx < len(row):
                mat = limpiar_dato(row.iloc[col_idx])
                if mat:
                    imgs.append({"nombre_archivo": f"Alumno_{mat}.png", "bytes": generar_qr_bytes(mat)})

        equipos.append({
            "Carpeta": nombre_carpeta,
            "Escuela": escuela,
            "Equipo": equipo,
            "Correo_Coach": correo_asesor,
            "Imagenes": imgs
        })
    return equipos

# --- INTERFAZ ---

st.title("Gestor de Torneo 🏆")

if "df_original" not in st.session_state:
    st.session_state.df_original = None
if "equipos_data" not in st.session_state:
    st.session_state.equipos_data = []

uploaded_file = st.file_uploader("Cargar Excel Master (.xlsx)", type=["xlsx"])

if uploaded_file:
    with st.spinner("Procesando estructura..."):
        df = cargar_dataframe(uploaded_file)
        if df is not None:
            st.session_state.df_original = df
            st.session_state.equipos_data = procesar_logica_zip_correo(df)
            st.success(f"✅ Archivo cargado. Equipos detectados: {len(st.session_state.equipos_data)}")

if st.session_state.df_original is not None:
    df = st.session_state.df_original
    datos_equipos = st.session_state.equipos_data
    
    st.divider()
    
    # 1. REPORTES EXCEL
    st.header("1. Reportes")
    excel_bytes, num_asesores = generar_excel_resumen(df)
    
    c1, c2 = st.columns([1, 2])
    c1.metric("Asesores Únicos", num_asesores)
    c2.download_button(
        "📥 Descargar Reporte Clasificado (Vertical)",
        data=excel_bytes,
        file_name="Reporte_Alumnos_Vertical.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    
    st.divider()

    # 2. QRS y CORREOS
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.subheader("2. QRs (ZIP)")
        if st.button("Generar ZIP de QRs"):
            b = io.BytesIO()
            with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
                for eq in datos_equipos:
                    for img in eq["Imagenes"]:
                        z.writestr(f"{eq['Carpeta']}/{img['nombre_archivo']}", img['bytes'])
            st.download_button("⬇️ Bajar ZIP", b.getvalue(), "QRs_Torneo.zip", "application/zip")

    with col_b:
        st.subheader("3. Envíos Coach")
        validos = [e for e in datos_equipos if e.get('Correo_Coach') and "@" in str(e.get('Correo_Coach'))]
        st.caption(f"Equipos listos para enviar: {len(validos)}")
        
        with st.expander("Configuración Email"):
            user = st.text_input("Usuario")
            pwd = st.text_input("App Password", type="password")
            prov = st.selectbox("Servidor", ["Gmail", "Outlook", "Yahoo"])
            
        if st.button("Enviar Correos"):
            if not user or not pwd:
                st.error("Faltan datos.")
            else:
                bar = st.progress(0)
                st_txt = st.empty()
                host, port = {
                    "Gmail": ("smtp.gmail.com", 465),
                    "Outlook": ("smtp.office365.com", 587),
                    "Yahoo": ("smtp.mail.yahoo.com", 465)
                }[prov]
                
                try:
                    s = smtplib.SMTP(host, port) if prov == "Outlook" else smtplib.SMTP_SSL(host, port)
                    if prov == "Outlook": s.starttls()
                    s.login(user, pwd)
                    
                    for i, eq in enumerate(validos):
                        bar.progress((i+1)/len(validos))
                        st_txt.text(f"Enviando: {eq['Equipo']}")
                        msg = EmailMessage()
                        msg['Subject'] = f"QRs - {eq['Equipo']}"
                        msg['From'] = user
                        msg['To'] = eq['Correo_Coach']
                        msg.set_content("Adjunto QRs.")
                        for img in eq['Imagenes']:
                            msg.add_attachment(img['bytes'], maintype='image', subtype='png', filename=img['nombre_archivo'])
                        s.send_message(msg)
                        time.sleep(1)
                    s.quit()
                    st.success("Listo!")
                except Exception as e:
                    st.error(f"Error: {e}")