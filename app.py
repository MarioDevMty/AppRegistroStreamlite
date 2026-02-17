import streamlit as st
import pandas as pd
import qrcode
import io
import zipfile

# Configuración de la página
st.set_page_config(page_title="Generador QR Torneo", page_icon="🤖", layout="centered")

def limpiar_dato(dato):
    """Convierte cualquier dato a string limpio, quitando decimales .0 y espacios extra."""
    if pd.isna(dato):
        return ""
    txt = str(dato).strip()
    if txt.endswith(".0"):
        return txt[:-2]
    return txt

def generar_imagen_qr(dato):
    """Genera los bytes de una imagen QR a partir de un dato."""
    qr = qrcode.QRCode(box_size=10, border=4)
    qr.add_data(dato)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    return img_byte_arr.getvalue()

def procesar_excel_y_zip(uploaded_file):
    """Procesa el Excel y genera el ZIP directamente."""
    
    # Leemos el archivo asegurando que todo sea texto para no perder ceros o formatos
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl', dtype=str)
    except Exception as e:
        return None, f"Error al leer el archivo: {str(e)}"

    zip_buffer = io.BytesIO()
    log_errores = []
    total_equipos = 0
    total_imagenes = 0

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        
        # Iteramos por cada fila (cada equipo)
        for index, row in df.iterrows():
            # --- 1. Extraer Datos del Equipo ---
            # B(1)=Escuela, D(3)=Equipo, E(4)=Categoria
            escuela = limpiar_dato(row.iloc[1])
            equipo = limpiar_dato(row.iloc[3])
            categoria = limpiar_dato(row.iloc[4])

            # Si faltan datos clave del equipo, saltamos la fila
            if not escuela or not equipo:
                continue
            
            total_equipos += 1

            # Crear nombre de la carpeta: "Prepa Equipo Categoria"
            # Limpiamos caracteres prohibidos para carpetas (/ \ : *)
            nombre_carpeta = f"{escuela} {equipo} {categoria}".strip()
            for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
                nombre_carpeta = nombre_carpeta.replace(char, '-')

            # Set para evitar duplicados DENTRO de la misma carpeta
            archivos_en_carpeta = set()

            # --- 2. Procesar ASESOR (Celular) ---
            # Columna I es el índice 8 (A=0 ... I=8)
            celular_asesor = limpiar_dato(row.iloc[8])
            
            if celular_asesor:
                nombre_archivo = f"{celular_asesor}.png"
                ruta_zip = f"{nombre_carpeta}/{nombre_archivo}"
                
                # Generar y guardar
                try:
                    img_bytes = generar_imagen_qr(celular_asesor)
                    zip_file.writestr(ruta_zip, img_bytes)
                    archivos_en_carpeta.add(nombre_archivo)
                    total_imagenes += 1
                except Exception as e:
                    log_errores.append(f"Error QR Asesor {equipo}: {e}")

            # --- 3. Procesar ALUMNOS (Matrículas) ---
            # Indices de columnas de MATRÍCULAS según tu tabla:
            # K (10), R (17), Y (24), AF (31), AN (39)
            cols_matriculas = [10, 17, 24, 31, 39]

            for col_idx in cols_matriculas:
                # Verificamos que la columna exista en el excel (por si el excel es más corto)
                if col_idx < len(row):
                    matricula = limpiar_dato(row.iloc[col_idx])
                    
                    if matricula:
                        nombre_archivo = f"{matricula}.png"
                        
                        # Evitar sobrescribir si por error pusieron la misma matrícula dos veces en el equipo
                        if nombre_archivo in archivos_en_carpeta:
                            nombre_archivo = f"{matricula}_duplicado.png"
                        
                        ruta_zip = f"{nombre_carpeta}/{nombre_archivo}"
                        
                        try:
                            img_bytes = generar_imagen_qr(matricula)
                            zip_file.writestr(ruta_zip, img_bytes)
                            archivos_en_carpeta.add(nombre_archivo)
                            total_imagenes += 1
                        except:
                            pass

    return zip_buffer, f"Proceso completado. {total_equipos} equipos procesados, {total_imagenes} imágenes generadas."

# --- INTERFAZ GRÁFICA ---
st.title("Generador de QRs por Equipo 🏆")
st.markdown("""
**Instrucciones:**
1. Sube el archivo Excel con el formato del torneo.
2. El sistema generará una carpeta por equipo: `[Escuela] [Equipo] [Categoría]`.
3. Dentro estarán los QRs del celular del asesor y las matrículas de los alumnos.
""")

uploaded_file = st.file_uploader("Cargar Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    if st.button("Generar QRs y ZIP"):
        with st.spinner("Procesando equipos y generando códigos..."):
            zip_final, mensaje = procesar_excel_y_zip(uploaded_file)
            
            if zip_final:
                st.success(mensaje)
                st.download_button(
                    label="⬇️ Descargar Archivo ZIP",
                    data=zip_final.getvalue(),
                    file_name="Codigos_QR_Torneo.zip",
                    mime="application/zip",
                    type="primary"
                )
            else:
                st.error(mensaje)