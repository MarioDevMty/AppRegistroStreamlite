import os
import csv

def limpiar_texto(texto):
    if not texto: return ""
    # Mismo formato que usaremos en Google: minúsculas y guiones bajos
    return texto.strip().lower().replace(" ", "_")

def proceso_total():
    archivo_origen = os.path.join('Archivos de trabajo', 'datos.csv') 
    archivo_salida_drive = 'datos_para_drive.csv'
    carpeta_raiz = "entregables_tris_xxiii"

    if not os.path.exists(archivo_origen):
        print("Error: No se encuentra el archivo datos.csv")
        return

    datos_drive = []

    with open(archivo_origen, mode='r', encoding='latin-1') as f:
        dialect = csv.Sniffer().sniff(f.read(1024))
        f.seek(0)
        reader = csv.DictReader(f, dialect=dialect)
        
        for fila in reader:
            # Normalizamos los nombres una sola vez aquí
            cat_limpia = limpiar_texto(fila['categoria'])
            equipo_limpio = limpiar_texto(fila['equipo'])
            prepa_limpia = limpiar_texto(fila['preparatoria'])
            correo = fila['correo'].strip()
            
            nombre_carpeta_equipo = f"{equipo_limpio}_{prepa_limpia}"
            ruta_final = os.path.join(carpeta_raiz, cat_limpia, nombre_carpeta_equipo)
            
            # 1. Crear carpetas
            os.makedirs(ruta_final, exist_ok=True)
            
            # 2. Crear archivo de instrucciones
            with open(os.path.join(ruta_final, "INSTRUCCIONES.txt"), "w", encoding="utf-8") as txt:
                txt.write(f"Equipo: {equipo_limpio}\nEntregue sus archivos aqui.")

            # 3. Guardar en la lista para el CSV de Drive
            datos_drive.append({
                'categoria_carpeta': cat_limpia,
                'equipo_carpeta': nombre_carpeta_equipo,
                'correo': correo
            })

    # Generar el CSV que subirás a Google Sheets
    with open(archivo_salida_drive, mode='w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['categoria_carpeta', 'equipo_carpeta', 'correo'])
        writer.writeheader()
        writer.writerows(datos_drive)

    print(f"\nListo! Carpetas creadas y archivo '{archivo_salida_drive}' generado.")

if __name__ == "__main__":
    proceso_total()