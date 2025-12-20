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

# --- LÓGICA PROPER ELEGANTE ---
def proper_elegante(texto):
    if not texto: return ""
    # Quitar acentos
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

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")
st.title("🚀 Generador Masivo Proper Elegante")

# Inicializar estados
if 'raw_records' not in st.session_state: st.session_state.raw_records = []

with st.sidebar:
    st.header("🔑 Conexión Airtable")
    headers = {"Authorization": f"Bearer {TOKEN}"}
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Selecciona Base:", list(base_opts.keys()))
        
        r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
        if r_tab.status_code == 200:
            tabla_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
            tabla_sel = st.selectbox("Selecciona Tabla:", list(tabla_opts.keys()))
            
            if st.button("🔄 Cargar Registros"):
                r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
                st.session_state.raw_records = r_reg.json().get("records", [])
                st.rerun()

# --- PROCESAMIENTO PRINCIPAL ---
if st.session_state.raw_records:
    df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    
    # Asegurar columnas críticas
    for c in ["Tipo", "Sucursal", "Fecha", "Hora", "Punto de reunion", "Ruta a seguir", "Municipio", "Seccion"]:
        if c not in df_full.columns: df_full[c] = ""
    
    # Editor de selección
    df_display = df_full[["Tipo", "Sucursal", "Municipio", "Fecha"]].copy()
    df_display.insert(0, "Seleccionar", False)
    
    df_edit = st.data_editor(df_display, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        st.divider()
        uso_final = st.radio("¿Qué deseas generar?", ["POSTALES", "REPORTES"], horizontal=True)
        folder = os.path.join(BASE_DIR, uso_final)
        
        if os.path.exists(folder):
            archivos_pptx = [f for f in os.listdir(folder) if f.endswith('.pptx')]
            tipos_seleccionados = df_edit.loc[sel_idx, "Tipo"].unique()
            
            # Mapeo de plantillas
            mapping = {t: st.selectbox(f"Plantilla para {t}", archivos_pptx, key=f"m_{t}") for t in tipos_seleccionados}

            if st.button("🔥 GENERAR ZIP ESTRUCTURADO"):
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                    for idx in sel_idx:
                        rec = st.session_state.raw_records[idx]['fields']
                        
                        with st.status(f"Procesando: {rec.get('Sucursal')}"):
                            # 1. PREPARAR CONTENIDO
                            suc = str(rec.get('Sucursal') or '').strip()
                            muni = str(rec.get('Municipio') or '').strip()
                            consuc_val = f"Sucursal {suc}" + (f", {muni}" if muni else "")
                            
                            c_partes = [str(rec.get(k) or '').strip() for k in ['Punto de reunion', 'Ruta a seguir'] if rec.get(k)]
                            concat_base = ", ".join(c_partes)
                            concat_txt = concat_base
                            if muni: concat_txt += f", Municipio {muni}"
                            if rec.get('Seccion'): concat_txt += f", Seccion {rec.get('Seccion')}"

                            reemplazos = {
                                "<<Consuc>>": proper_elegante(consuc_val),
                                "<<Confecha>>": formatear_fecha_mx(rec.get('Fecha')),
                                "<<Conhora>>": formatear_hora_mx(rec.get('Hora')),
                                "<<Concat>>": proper_elegante(concat_txt),
                                "<<Sucursal>>": proper_elegante(suc)
                            }

                            # 2. PPTX
                            prs = Presentation(os.path.join(folder, mapping[rec.get('Tipo')]))
                            for slide in prs.slides:
                                for shape in slide.shapes:
                                    if shape.has_text_frame:
                                        for paragraph in shape.text_frame.paragraphs:
                                            for run in paragraph.runs:
                                                for tag, val in reemplazos.items():
                                                    if tag in run.text:
                                                        run.text = run.text.replace(tag, val)
                                                        # Ajuste de tamaño para Concat si es muy largo
                                                        if tag == "<<Concat>>" and len(val) > 70:
                                                            run.font.size = Pt(11)

                            # 3. CONVERSIÓN Y RUTA ZIP
                            pp_io = BytesIO(); prs.save(pp_io)
                            pdf_data = generar_pdf(pp_io.getvalue())
                            
                            if pdf_data:
                                f_raw = rec.get('Fecha', '2024-01-01')
                                try: dt = datetime.strptime(str(f_raw), '%Y-%m-%d')
                                except: dt = datetime.now()
                                
                                mes_f = proper_elegante(dt.strftime('%m - %B'))
                                ext = "png" if uso_final == "POSTALES" else "pdf"
                                # Nombre: Día de Mes TIPO SUCURSAL
                                nombre_arc = proper_elegante(f"{dt.day} de {dt.month} {rec.get('Tipo')} {suc}") + f".{ext}"
                                
                                path_zip = (f"Provident/{dt.year}/{mes_f}/{proper_elegante(uso_final)}/"
                                           f"{proper_elegante(suc)}/{nombre_arc}")
                                
                                if uso_final == "REPORTES":
                                    zip_f.writestr(path_zip, pdf_data)
                                else:
                                    img = generar_png(pdf_data)
                                    if img: zip_f.writestr(path_zip, img)

                st.success("✅ ¡Proceso Terminado!")
                st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_Pro.zip")
        else:
            st.error(f"No se encontró la carpeta de plantillas: {folder}")
else:
    st.info("Por favor, carga los registros desde la barra lateral.")
