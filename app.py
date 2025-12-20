import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
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

def formatear_hora_mx(hora_raw):
    if not hora_raw: return "00:00"
    try:
        # Airtable a veces envía la hora como "14:30" o similar
        t_str = str(hora_raw).strip()
        t = datetime.strptime(t_str, "%H:%M")
        return t.strftime("%I:%M %p").lower().replace("am", "a.m.").replace("pm", "p.m.")
    except:
        return str(hora_raw)

def formatear_confechor(fecha_str, hora_raw):
    try:
        dt = datetime.strptime(str(fecha_str), '%Y-%m-%d')
        hora = formatear_hora_mx(hora_raw)
        # Salto de línea exacto pedido
        linea1 = f"{dt.day:02d} de {MESES_ES[dt.month-1]}"
        linea2 = f"de {dt.year}, {hora}"
        return f"{linea1}\n{linea2}"
    except:
        return "Fecha/Hora Inválida"

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

def descargar_imagen(url):
    try:
        r = requests.get(url, timeout=15)
        if r.status_code == 200: return BytesIO(r.content)
    except: return None

# --- UI ---
st.set_page_config(page_title="Provident Pro FIX", layout="wide")
st.title("🚀 Generador Pro (Corrección de Datos)")

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
        modo = st.radio("Formato:", ["Postales", "Reportes"], horizontal=True)
        folder_fisica = os.path.join(BASE_DIR, modo.upper())
        
        if os.path.exists(folder_fisica):
            archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
            tipos_unicos = df_edit.loc[sel_idx, "Tipo"].unique()
            for t in tipos_unicos:
                st.session_state.map_memoria[t] = st.selectbox(f"Plantilla {t}:", archivos_pptx, key=f"p_{t}")

            if st.button("🔥 GENERAR ZIP"):
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                    for idx in sel_idx:
                        f = st.session_state.raw_records[idx]['fields']
                        
                        # REEMPLAZOS
                        reemplazos = {
                            "<<Tipo>>": str(f.get('Tipo', '')).upper(),
                            "<<confechor>>": formatear_confechor(f.get('Fecha'), f.get('Hora')),
                            "<<Consuc>>": proper_elegante(f"Sucursal {f.get('Sucursal', '')}, {f.get('Municipio', '')}"),
                            "<<Confecha>>": proper_elegante(f"{f.get('Fecha', '')}"),
                            "<<Concat>>": proper_elegante(f"{f.get('Punto de reunion', '')}, {f.get('Ruta a seguir', '')}, {f.get('Municipio', '')}"),
                            "<<Sucursal>>": proper_elegante(f.get('Sucursal', ''))
                        }

                        prs = Presentation(os.path.join(folder_fisica, st.session_state.map_memoria[f.get('Tipo')]))
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    for paragraph in shape.text_frame.paragraphs:
                                        # Truco para que encuentre el tag aunque esté fragmentado en el XML
                                        full_text = "".join(run.text for run in paragraph.runs)
                                        for tag, val in reemplazos.items():
                                            if tag in full_text:
                                                new_text = full_text.replace(tag, val)
                                                # Limpiar runs y poner el nuevo texto
                                                for r in paragraph.runs: r.text = ""
                                                run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
                                                run.text = new_text
                                                run.font.size = Pt(32) if tag == "<<confechor>>" else Pt(42)

                                # Fotos (Solo Reportes)
                                if modo == "Reportes":
                                    tags_fotos = ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]
                                    for tag_f in tags_fotos:
                                        if (shape.has_text_frame and f"<<{tag_f}>>" in shape.text) or (tag_f in shape.name):
                                            adjuntos = f.get(tag_f)
                                            if adjuntos and isinstance(adjuntos, list):
                                                img_io = descargar_imagen(adjuntos[0].get('url'))
                                                if img_io:
                                                    slide.shapes.add_picture(img_io, shape.left, shape.top, shape.width, shape.height)

                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        if pdf_data:
                            dt = datetime.strptime(str(f.get('Fecha', '2025-01-01')), '%Y-%m-%d')
                            nombre_archivo = proper_elegante(f"{f.get('Fecha')} {f.get('Tipo')} {f.get('Sucursal')}") + (".jpg" if modo == "Postales" else ".pdf")
                            ruta_zip = f"Provident/{dt.year}/{dt.month:02d}/{modo}/{nombre_archivo}"
                            
                            if modo == "Reportes":
                                zip_f.writestr(ruta_zip, pdf_data)
                            else:
                                imgs = convert_from_bytes(pdf_data)
                                if imgs:
                                    img_opt = BytesIO()
                                    imgs[0].convert('RGB').save(img_opt, format='JPEG', quality=85)
                                    zip_f.writestr(ruta_zip, img_opt.getvalue())

                st.success("✅ ZIP Listo")
                st.download_button("📥 DESCARGAR", zip_buf.getvalue(), "Provident_Pro.zip")
