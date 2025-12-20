import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import os, subprocess, tempfile, zipfile, unicodedata
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"
AZUL_PROVIDENT = RGBColor(0, 32, 96) # Ajusta este RGB al azul de tu marca
MESES_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
DIAS_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

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

def interpretar_hora(hora_txt):
    if not hora_txt: return ""
    hora_txt = str(hora_txt).strip().lower()
    for fmt in ["%H:%M", "%I:%M %p", "%H%M", "%I%p", "%I:%M%p"]:
        try:
            dt_hora = datetime.strptime(hora_txt.replace(" ", ""), fmt.replace(" ", ""))
            return dt_hora.strftime("%I:%M %p").lower().replace("am", "a.m.").replace("pm", "p.m.")
        except: continue
    return hora_txt

def formatear_confechor(fecha_str, hora_txt):
    try:
        dt = datetime.strptime(str(fecha_str).strip(), '%Y-%m-%d')
        dia_semana = DIAS_ES[dt.weekday()]
        mes_nombre = MESES_ES[dt.month-1]
        hora_formateada = interpretar_hora(hora_txt)
        l1 = f"{dia_semana} {mes_nombre} {dt.day:02d}"
        l2 = f"de {dt.year}, {hora_formateada}"
        return f"{l1}\n{l2}"
    except: return f"{fecha_str}\n{hora_txt}"

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

# --- UI STREAMLIT ---
st.set_page_config(page_title="Provident Pro Ultra", layout="wide")
st.title("🚀 Generador Pro: Formato Centrado y Azul")

# (Lógica de sidebar igual a la anterior...)
with st.sidebar:
    headers = {"Authorization": f"Bearer {TOKEN}"}
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base:", list(base_opts.keys()))
        if base_sel:
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
            if r_tab.status_code == 200:
                tabla_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
                tabla_sel = st.selectbox("Tabla:", list(tabla_opts.keys()))
                if st.button("🔄 CARGAR DATOS"):
                    r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
                    st.session_state.raw_records = r_reg.json().get("records", [])
                    st.rerun()

if st.session_state.raw_records:
    df_prev = pd.DataFrame([{"Tipo": r['fields'].get("Tipo"), "Sucursal": r['fields'].get("Sucursal"), "Fecha": r['fields'].get("Fecha")} for r in st.session_state.raw_records])
    df_prev.insert(0, "Seleccionar", False)
    df_edit = st.data_editor(df_prev, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        modo = st.radio("Acción:", ["Postales", "Reportes"], horizontal=True)
        folder_fisica = os.path.join(BASE_DIR, modo.upper())
        
        if os.path.exists(folder_fisica):
            archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
            tipos_unicos = df_edit.loc[sel_idx, "Tipo"].unique()
            for t in tipos_unicos:
                st.session_state.map_memoria[t] = st.selectbox(f"Plantilla {t}:", archivos_pptx, key=f"p_{t}")

            if st.button("🔥 GENERAR ARCHIVOS"):
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                    for idx in sel_idx:
                        f = st.session_state.raw_records[idx]['fields']
                        
                        # Preparar reemplazos
                        reemplazos = {
                            "<<Tipo>>": str(f.get('Tipo', '')).upper(),
                            "<<Confechor>>": formatear_confechor(f.get('Fecha'), f.get('Hora')),
                            "<<Consuc>>": proper_elegante(f"Sucursal {f.get('Sucursal', '')}, {f.get('Municipio', '')}"),
                            "<<Concat>>": proper_elegante(f"{f.get('Punto de reunion', '')}, {f.get('Ruta a seguir', '')}, {f.get('Municipio', '')}"),
                            "<<Sucursal>>": proper_elegante(f.get('Sucursal', ''))
                        }

                        prs = Presentation(os.path.join(folder_fisica, st.session_state.map_memoria[f.get('Tipo')]))
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    # Centrado vertical del cuadro
                                    shape.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                                    
                                    for paragraph in shape.text_frame.paragraphs:
                                        # Centrado horizontal
                                        paragraph.alignment = PP_ALIGN.CENTER
                                        
                                        full_text = "".join(run.text for run in paragraph.runs)
                                        for tag, val in reemplazos.items():
                                            if tag in full_text:
                                                full_text = full_text.replace(tag, val)
                                                # Limpiar y reescribir con estilo
                                                for r in paragraph.runs: r.text = ""
                                                run = paragraph.add_run() if not paragraph.runs else paragraph.runs[0]
                                                run.text = full_text
                                                
                                                # Estilo: Color y Tamaño
                                                run.font.color.rgb = AZUL_PROVIDENT
                                                run.font.bold = True
                                                
                                                if tag == "<<Confechor>>": run.font.size = Pt(34)
                                                elif tag == "<<Sucursal>>": run.font.size = Pt(40)
                                                else: run.font.size = Pt(50) # Tamaño máximo para Concat y Tipo

                                # Fotos
                                if modo == "Reportes":
                                    tags_f = ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]
                                    for tf in tags_f:
                                        if (shape.has_text_frame and f"<<{tf}>>" in shape.text) or (tf in shape.name):
                                            adj = f.get(tf)
                                            if adj and isinstance(adj, list):
                                                r_img = requests.get(adj[0].get('url'))
                                                if r_img.status_code == 200:
                                                    slide.shapes.add_picture(BytesIO(r_img.content), shape.left, shape.top, shape.width, shape.height)

                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        if pdf_data:
                            dt = datetime.strptime(f.get('Fecha', '2025-01-01'), '%Y-%m-%d')
                            nom_f = f"{f.get('Fecha')} - {str(f.get('Tipo')).upper()} {str(f.get('Sucursal')).upper()}"
                            ruta_zip = f"Provident/{dt.year}/{dt.month:02d}/{modo}/{proper_elegante(f.get('Sucursal'))}/{nom_f}.jpg"
                            
                            if modo == "Reportes":
                                zip_f.writestr(ruta_zip.replace(".jpg", ".pdf"), pdf_data)
                            else:
                                imgs = convert_from_bytes(pdf_data)
                                if imgs:
                                    img_io = BytesIO()
                                    imgs[0].convert('RGB').save(img_io, format='JPEG', quality=90)
                                    zip_f.writestr(ruta_zip, img_io.getvalue())

                st.success("✅ Archivos generados con alineación centrada y color azul.")
                st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_Pro.zip")
