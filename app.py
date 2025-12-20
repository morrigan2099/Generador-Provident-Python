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
        if i > 0 and p in excepciones:
            resultado.append(p)
        else:
            resultado.append(p.capitalize())
    return " ".join(resultado)

def formatear_fecha_mx(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        dias = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]
        res = f"{dias[dt.weekday()]} {dt.day} de {meses[dt.month-1]} de {dt.year}"
        return proper_elegante(res)
    except: return ""

def formatear_hora_mx(hora_raw):
    if not hora_raw: return ""
    try:
        t = datetime.strptime(str(hora_raw).strip(), "%H:%M")
        return t.strftime("%I:%M %p").lower().replace("am", "a.m.").replace("pm", "p.m.")
    except: return str(hora_raw)

# --- MOTOR DE CONVERSIÓN ---
def generar_pdf(pptx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        path = tmp.name
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(path), path], check=True)
        pdf_path = path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f: data = f.read()
        os.remove(path); os.remove(pdf_path) if os.path.exists(pdf_path) else None
        return data
    except: return None

# --- INTERFAZ ---
st.set_page_config(page_title="Provident Pro", layout="wide")
st.title("🚀 Generador Pro: Estilo Preservado")

with st.sidebar:
    headers = {"Authorization": f"Bearer {TOKEN}"}
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base:", list(base_opts.keys()))
        r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
        if r_tab.status_code == 200:
            tabla_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
            tabla_sel = st.selectbox("Tabla:", list(tabla_opts.keys()))
            if st.button("🔄 Cargar"):
                r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
                st.session_state.raw_records = r_reg.json().get("records", [])
                st.rerun()

if st.session_state.raw_records:
    # Mostramos la tabla estilizada
    data_list = []
    for r in st.session_state.raw_records:
        f = r['fields']
        data_list.append({
            "Tipo": proper_elegante(f.get("Tipo")),
            "Sucursal": proper_elegante(f.get("Sucursal")),
            "Municipio": proper_elegante(f.get("Municipio")),
            "Fecha": f.get("Fecha", "")
        })
    df_display = pd.DataFrame(data_list)
    df_display.insert(0, "Seleccionar", False)
    df_edit = st.data_editor(df_display, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        opcion_salida = st.radio("Tipo a generar:", ["Postales", "Reportes"], horizontal=True)
        folder_fisica = os.path.join(BASE_DIR, opcion_salida.upper())
        archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
        
        # Mapeo de plantillas por tipo
        tipos_sel = df_edit.loc[sel_idx, "Tipo"].unique()
        for t_p in tipos_sel:
            if t_p not in st.session_state.map_memoria:
                st.session_state.map_memoria[t_p] = archivos_pptx[0] if archivos_pptx else ""

        if st.button("🔥 GENERAR ZIP"):
            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                for idx in sel_idx:
                    rec = st.session_state.raw_records[idx]['fields']
                    tipo_p = proper_elegante(rec.get('Tipo'))
                    
                    # 1. PREPARAR DATOS (CONCATENACIÓN)
                    suc = str(rec.get('Sucursal') or '').strip()
                    muni = str(rec.get('Municipio') or '').strip()
                    punto = str(rec.get('Punto de reunion') or '').strip()
                    ruta = str(rec.get('Ruta a seguir') or '').strip()
                    secc = str(rec.get('Seccion') or '').strip()

                    reemplazos = {
                        "<<Consuc>>": proper_elegante(f"Sucursal {suc}" + (f", {muni}" if muni else "")),
                        "<<Confecha>>": formatear_fecha_mx(rec.get('Fecha')),
                        "<<Conhora>>": formatear_hora_mx(rec.get('Hora')),
                        "<<Concat>>": proper_elegante(", ".join([p for p in [punto, ruta] if p]) + (f", Municipio {muni}" if muni else "") + (f", Seccion {secc}" if secc else "")),
                        "<<Sucursal>>": proper_elegante(suc)
                    }

                    # 2. PROCESAR PPTX CON PRESERVACIÓN DE ESTILO
                    prs = Presentation(os.path.join(folder_fisica, st.session_state.map_memoria[tipo_p]))
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for paragraph in shape.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        for tag, val in reemplazos.items():
                                            if tag in run.text:
                                                # Guardamos el estilo original antes del reemplazo
                                                font_size = run.font.size
                                                font_name = run.font.name
                                                font_bold = run.font.bold
                                                font_color = run.font.color.rgb if run.font.color and hasattr(run.font.color, 'rgb') else None
                                                
                                                # Realizamos el reemplazo
                                                run.text = run.text.replace(tag, val)
                                                
                                                # Forzamos la restauración del estilo
                                                run.font.size = font_size
                                                run.font.name = font_name
                                                run.font.bold = font_bold
                                                if font_color: run.font.color.rgb = font_color

                    # 3. GUARDADO
                    pp_io = BytesIO(); prs.save(pp_io)
                    pdf_data = generar_pdf(pp_io.getvalue())
                    if pdf_data:
                        dt = datetime.strptime(str(rec.get('Fecha', '2024-01-01')), '%Y-%m-%d')
                        nombre = proper_elegante(f"{dt.day} de {dt.month} {rec.get('Tipo')} {suc}") + (".png" if opcion_salida == "Postales" else ".pdf")
                        ruta_final = f"Provident/{dt.year}/{proper_elegante(dt.strftime('%m - %B'))}/{opcion_salida}/{proper_elegante(suc)}/{nombre}"
                        zip_f.writestr(ruta_final, pdf_data)

            st.download_button("📥 DESCARGAR", zip_buf.getvalue(), "Provident.zip")
