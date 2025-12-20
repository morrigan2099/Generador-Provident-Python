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

# --- INICIALIZACIÓN DE ESTADOS (MEMORIA) ---
if 'raw_records' not in st.session_state: st.session_state.raw_records = []
if 'map_memoria' not in st.session_state: st.session_state.map_memoria = {}

# --- LÓGICA PROPER ELEGANTE ---
def proper_elegante(texto):
    if not texto: return ""
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

# --- MOTORES DE CONVERSIÓN ---
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

def generar_png(pdf_bytes):
    try:
        images = convert_from_bytes(pdf_bytes)
        if images:
            buf = BytesIO()
            images[0].save(buf, format='PNG')
            return buf.getvalue()
    except: return None

# --- INTERFAZ ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")
st.title("🚀 Generador con Memoria de Plantillas")

with st.sidebar:
    st.header("🔑 Conexión")
    headers = {"Authorization": f"Bearer {TOKEN}"}
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base:", list(base_opts.keys()))
        r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
        if r_tab.status_code == 200:
            tabla_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
            tabla_sel = st.selectbox("Tabla:", list(tabla_opts.keys()))
            if st.button("🔄 Cargar Datos"):
                r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
                st.session_state.raw_records = r_reg.json().get("records", [])
                st.rerun()

if st.session_state.raw_records:
    df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    for c in ["Tipo", "Sucursal", "Fecha", "Hora", "Punto de reunion", "Ruta a seguir", "Municipio", "Seccion"]:
        if c not in df_full.columns: df_full[c] = ""
    
    df_display = df_full[["Tipo", "Sucursal", "Municipio", "Fecha"]].copy()
    df_display.insert(0, "Seleccionar", False)
    df_edit = st.data_editor(df_display, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        uso_final = st.radio("Formato:", ["POSTALES", "REPORTES"], horizontal=True)
        folder = os.path.join(BASE_DIR, uso_final)
        archivos_pptx = [f for f in os.listdir(folder) if f.endswith('.pptx')]
        tipos_sel = df_edit.loc[sel_idx, "Tipo"].unique()
        
        st.write("### 📂 Asignar Plantillas por Tipo")
        cols = st.columns(len(tipos_sel))
        for i, t in enumerate(tipos_sel):
            with cols[i]:
                # Recuperar de memoria si existe, si no, el primero de la lista
                idx_memoria = 0
                if t in st.session_state.map_memoria:
                    if st.session_state.map_memoria[t] in archivos_pptx:
                        idx_memoria = archivos_pptx.index(st.session_state.map_memoria[t])
                
                sel_p = st.selectbox(f"Tipo: {t}", archivos_pptx, index=idx_memoria, key=f"sel_{t}")
                st.session_state.map_memoria[t] = sel_p # GUARDAR EN MEMORIA

        if st.button("🔥 GENERAR ZIP"):
            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                for idx in sel_idx:
                    rec = st.session_state.raw_records[idx]['fields']
                    tipo_reg = rec.get('Tipo', '')
                    
                    with st.status(f"Procesando: {rec.get('Sucursal')}"):
                        # 1. CONTENIDO
                        suc = str(rec.get('Sucursal') or '').strip()
                        muni = str(rec.get('Municipio') or '').strip()
                        consuc_val = f"Sucursal {suc}" + (f", {muni}" if muni else "")
                        c_partes = [str(rec.get(k) or '').strip() for k in ['Punto de reunion', 'Ruta a seguir'] if rec.get(k)]
                        concat_txt = ", ".join(c_partes)
                        if muni: concat_txt += f", Municipio {muni}"
                        if rec.get('Seccion'): concat_txt += f", Seccion {rec.get('Seccion')}"

                        reemplazos = {
                            "<<Consuc>>": proper_elegante(consuc_val),
                            "<<Confecha>>": formatear_fecha_mx(rec.get('Fecha')),
                            "<<Conhora>>": formatear_hora_mx(rec.get('Hora')),
                            "<<Concat>>": proper_elegante(concat_txt),
                            "<<Sucursal>>": proper_elegante(suc)
                        }

                        # 2. PPTX USANDO MAPEO DE MEMORIA
                        prs = Presentation(os.path.join(folder, st.session_state.map_memoria[tipo_reg]))
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    for paragraph in shape.text_frame.paragraphs:
                                        for run in paragraph.runs:
                                            for tag, val in reemplazos.items():
                                                if tag in run.text:
                                                    run.text = run.text.replace(tag, val)

                        # 3. ZIP Y CONVERSIÓN
                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        if pdf_data:
                            f_raw = rec.get('Fecha', '2024-01-01')
                            try: dt = datetime.strptime(str(f_raw), '%Y-%m-%d')
                            except: dt = datetime.now()
                            
                            nombre_arc = proper_elegante(f"{dt.day} de {dt.month} {tipo_reg} {suc}") + (".png" if uso_final == "POSTALES" else ".pdf")
                            path_zip = f"Provident/{dt.year}/{proper_elegante(dt.strftime('%m - %B'))}/{proper_elegante(uso_final)}/{proper_elegante(suc)}/{nombre_arc}"
                            
                            if uso_final == "REPORTES": zip_f.writestr(path_zip, pdf_data)
                            else:
                                img = generar_png(pdf_data)
                                if img: zip_f.writestr(path_zip, img)

            st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_Pro.zip")
