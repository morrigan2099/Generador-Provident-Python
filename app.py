import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from pptx.util import Pt, Inches
import os, subprocess, tempfile, zipfile, unicodedata
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"
MESES_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

if 'raw_records' not in st.session_state: st.session_state.raw_records = []
if 'map_memoria' not in st.session_state: st.session_state.map_memoria = {}

def proper_elegante(texto):
    if not texto or str(texto).lower() == "none": return ""
    texto = ''.join(c for c in unicodedata.normalize('NFD', str(texto))
                  if unicodedata.category(c) != 'Mn')
    palabras = texto.lower().split()
    excepciones = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del']
    resultado = []
    for i, p in enumerate(palabras):
        if i > 0 and p in excepciones: resultado.append(p)
        else: resultado.append(p.capitalize())
    return " ".join(resultado)

def formatear_fecha_partes(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        # Parte 1: "dd de Mes" | Parte 2: "de aaaa"
        p1 = f"{dt.day:02d} de {MESES_ES[dt.month-1]}"
        p2 = f"de {dt.year}"
        return p1, p2
    except: return "", ""

def formatear_hora_mx(hora_raw):
    if not hora_raw: return ""
    try:
        t = datetime.strptime(str(hora_raw).strip(), "%H:%M")
        return t.strftime("%I:%M %p").lower().replace("am", "a.m.").replace("pm", "p.m.")
    except: return str(hora_raw)

def descargar_imagen(url):
    try:
        r = requests.get(url, headers={"Authorization": f"Bearer {TOKEN}"})
        if r.status_code == 200: return BytesIO(r.content)
    except: return None
    return None

def generar_pdf(pptx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        path = tmp.name
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(path), path], check=True)
        pdf_path = path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f: data = f.read()
        os.remove(path)
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return data
    except: return None

# --- APP ---
st.set_page_config(page_title="Provident Report Pro", layout="wide")
st.title("🚀 Generador de Reportes y Postales JPG")

# (Sidebar de conexión Airtable se mantiene igual que en versiones anteriores)
# ... [Código de Sidebar omitido por brevedad] ...

if st.session_state.raw_records:
    df_display = pd.DataFrame([{"Tipo": r['fields'].get("Tipo"), "Sucursal": r['fields'].get("Sucursal"), "Fecha": r['fields'].get("Fecha")} for r in st.session_state.raw_records])
    df_display.insert(0, "Seleccionar", False)
    df_edit = st.data_editor(df_display, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        op_salida = st.radio("Acción:", ["Postales", "Reportes"], horizontal=True)
        folder_fisica = os.path.join(BASE_DIR, op_salida.upper())
        
        if os.path.exists(folder_fisica):
            archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
            tipos_unicos = df_edit.loc[sel_idx, "Tipo"].unique()
            for t_p in tipos_unicos:
                st.session_state.map_memoria[t_p] = st.selectbox(f"Plantilla para {t_p}:", archivos_pptx)

            if st.button("🔥 EJECUTAR"):
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                    for idx in sel_idx:
                        fields = st.session_state.raw_records[idx]['fields']
                        
                        # --- LÓGICA CONFECHOR ---
                        p1, p2 = formatear_fecha_partes(fields.get('Fecha'))
                        hora = formatear_hora_mx(fields.get('Hora'))
                        confechor = f"{p1}\n{p2}, {hora}"

                        # --- REEMPLAZOS DE TEXTO ---
                        reemplazos = {
                            "<<Tipo>>": str(fields.get('Tipo')).upper(),
                            "<<confechor>>": confechor,
                            "<<Consuc>>": proper_elegante(f"Sucursal {fields.get('Sucursal')}, {fields.get('Municipio', '')}"),
                            "<<Concat>>": proper_elegante(f"{fields.get('Punto de reunion', '')}, {fields.get('Ruta a seguir', '')}, {fields.get('Municipio', '')}"),
                            "<<Sucursal>>": proper_elegante(fields.get('Sucursal'))
                        }

                        # --- CARGAR PLANTILLA ---
                        prs = Presentation(os.path.join(folder_fisica, st.session_state.map_memoria[fields.get('Tipo')]))
                        
                        for slide in prs.slides:
                            # 1. Reemplazo de Texto y Fotos (Placeholders de imagen)
                            for shape in slide.shapes:
                                # Manejo de Texto
                                if shape.has_text_frame:
                                    for paragraph in shape.text_frame.paragraphs:
                                        for tag, val in reemplazos.items():
                                            if tag in paragraph.text:
                                                paragraph.text = paragraph.text.replace(tag, val)
                                                # Mantener tamaño 32pt para confechor, 42pt para otros
                                                paragraph.font.size = Pt(32) if tag == "<<confechor>>" else Pt(42)

                                # Manejo de Fotos (Si el nombre del shape coincide con el tag)
                                photo_tags = ["<<Foto de equipo>>", "<<Foto 01>>", "<<Foto 02>>", "<<Foto 03>>", 
                                              "<<Reporte firmado>>", "<<Lista de asistencia>>"]
                                for p_tag in photo_tags:
                                    if p_tag in shape.name or (shape.has_text_frame and p_tag in shape.text):
                                        field_name = p_tag.replace("<<", "").replace(">>", "")
                                        img_data = fields.get(field_name)
                                        if img_data and isinstance(img_data, list):
                                            img_url = img_data[0].get('url')
                                            img_file = descargar_imagen(img_url)
                                            if img_file:
                                                # Insertar imagen sobre el shape actual
                                                slide.shapes.add_picture(img_file, shape.left, shape.top, shape.width, shape.height)

                        # --- NOMENCLATURA Y GUARDADO ---
                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        
                        if pdf_data:
                            # Nomenclatura del archivo solicitada
                            nombre_base = f"{fields.get('Fecha')} - {fields.get('Tipo')} {fields.get('Sucursal')}"
                            nombre_final = proper_elegante(nombre_base) + (".jpg" if op_salida == "Postales" else ".pdf")
                            
                            dt = datetime.strptime(fields.get('Fecha'), '%Y-%m-%d')
                            ruta_zip = f"Provident/{dt.year}/{dt.month:02d} - {MESES_ES[dt.month-1]}/{op_salida}/{proper_elegante(fields.get('Sucursal'))}/{nombre_final}"
                            
                            if op_salida == "Reportes":
                                zip_f.writestr(ruta_zip, pdf_data)
                            else:
                                imgs = convert_from_bytes(pdf_data)
                                if imgs:
                                    img_io = BytesIO()
                                    imgs[0].convert('RGB').save(img_io, format='JPEG', quality=85)
                                    zip_f.writestr(ruta_zip, img_io.getvalue())

                st.download_button("📥 DESCARGAR RESULTADOS", zip_buf.getvalue(), "Provident_Pro.zip")
