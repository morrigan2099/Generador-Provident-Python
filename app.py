import streamlit as st
import requests
import pandas as pd
import json
import os
import re
import unicodedata
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import subprocess, tempfile, zipfile
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN DE PERSISTENCIA ---
CONFIG_FILE = "config_app.json"

def cargar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f: 
                cfg = json.load(f)
                if "plantillas" not in cfg: cfg["plantillas"] = {}
                if "columnas_visibles" not in cfg: cfg["columnas_visibles"] = []
                return cfg
        except: pass
    return {"plantillas": {}, "columnas_visibles": []}

def guardar_config_json(config_data):
    with open(CONFIG_FILE, "w") as f: 
        json.dump(config_data, f, indent=4)

if 'config' not in st.session_state:
    st.session_state.config = cargar_config()

# --- MOTOR DE TEXTO MAESTRO ---
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]

def procesar_texto_maestro(texto, campo=""):
    if not texto or str(texto).lower() == "none": return ""
    # Si es una lista (adjuntos), no procesar como texto
    if isinstance(texto, list): return texto
    
    texto = str(texto)
    nfkd = unicodedata.normalize('NFKD', texto)
    texto_limpio = "".join([c for c in nfkd if not unicodedata.combining(c)]).lower()
    texto_limpio = texto_limpio.replace('\n', ' ').replace('\r', ' ')
    texto_limpio = re.sub(r'\s+', ' ', texto_limpio).strip()
    
    if campo == 'Hora': return texto_limpio
    if campo == 'Seccion': return texto_limpio.upper()

    pequenas = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al']
    tokens = re.split(r'(\s+|\.|\(|\))', texto_limpio)
    resultado = []
    for i, t in enumerate(tokens):
        if re.match(r'\s+|\.|\(|\)', t):
            resultado.append(t); continue
        forzar_mayuscula = i == 0
        if not forzar_mayuscula:
            previo = "".join(tokens[:i]).strip()
            if previo.endswith('.') or previo.endswith('('): forzar_mayuscula = True
        
        if forzar_mayuscula: resultado.append(t.capitalize())
        elif t in pequenas: resultado.append(t.lower())
        else: resultado.append(t.capitalize())
    return "".join(resultado)

def generar_pdf(pptx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        path = tmp.name
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(path), path], check=True)
        pdf_path = path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f: data = f.read()
        os.remove(path); os.remove(pdf_path)
        return data
    except: return None

# --- UI ---
st.set_page_config(page_title="Provident Pro v29", layout="wide")
st.title("🚀 Generador Pro: Fix Imágenes y Estilo")

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

with st.sidebar:
    if st.button("💾 Guardar Configuración", use_container_width=True, type="primary"):
        guardar_config_json(st.session_state.config)
        st.toast("Configuración guardada")
    st.divider()
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base:", [""] + list(base_opts.keys()))
        if base_sel:
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
            tabla_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
            tabla_sel = st.selectbox("Tabla:", list(tabla_opts.keys()))
            if st.button("🔄 CARGAR DATOS"):
                r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
                raw = r_reg.json().get("records", [])
                procesados = []
                for r in raw:
                    f = r['fields']
                    row_limpio = {k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v) for k, v in f.items()}
                    procesados.append({'id': r['id'], 'fields': row_limpio})
                st.session_state.raw_records = procesados
                st.rerun()

# --- FLUJO PRINCIPAL ---
modo = st.radio("1. Acción:", ["Postales", "Reportes"], horizontal=True)
folder_fisica = os.path.join("Plantillas", modo.upper())
AZUL_CELESTE = RGBColor(0, 176, 240)

if 'raw_records' in st.session_state:
    df = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    # Limpieza visual de la tabla (fechas y quitar columnas de adjuntos)
    df_view = df.copy()
    for col in df_view.columns:
        if isinstance(df_view[col].iloc[0], list): df_view.drop(col, axis=1, inplace=True)
    
    df_view.insert(0, "Seleccionar", False)
    df_edit = st.data_editor(df_view, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
        tipos_sel = df_view.loc[sel_idx, "Tipo"].unique()
        for t in tipos_sel:
            p_mem = st.session_state.config["plantillas"].get(t)
            idx_def = archivos_pptx.index(p_mem) if p_mem in archivos_pptx else 0
            st.session_state.config["plantillas"][t] = st.selectbox(f"Plantilla para {t}:", archivos_pptx, index=idx_def, key=t)

        if st.button("🔥 GENERAR", use_container_width=True, type="primary"):
            p_bar = st.progress(0); status = st.empty(); zip_buf = BytesIO()
            total = len(sel_idx)
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                for i, idx in enumerate(sel_idx):
                    record = st.session_state.raw_records[idx]['fields']
                    dt = datetime.strptime(record.get('Fecha'), '%Y-%m-%d')
                    f_tipo = record.get('Tipo'); f_suc = record.get('Sucursal')
                    status.text(f"Procesando {i+1}/{total}: {f_tipo} - {f_suc}")
                    
                    # Reemplazos de texto
                    lugar = record.get('Punto de reunion') or record.get('Ruta a seguir')
                    f_confechor = procesar_texto_maestro(f"{DIAS_ES[dt.weekday()]} {dt.day} de {MESES_ES[dt.month-1]} de {dt.year}, {record.get('Hora')}")
                    f_concat = procesar_texto_maestro(f"{lugar}, {record.get('Municipio')}")
                    
                    reemplazos = {"<<Tipo>>": f_tipo, "<<Sucursal>>": f_suc, "<<Seccion>>": record.get('Seccion'), "<<Confechor>>": f_confechor, "<<Concat>>": f_concat}
                    
                    prs = Presentation(os.path.join(folder_fisica, st.session_state.config["plantillas"][f_tipo]))
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for tag, val in reemplazos.items():
                                    if tag in shape.text_frame.text:
                                        tf = shape.text_frame; tf.auto_size = None; tf.clear()
                                        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
                                        run = p.add_run(); run.text = str(val); run.font.bold = True; run.font.color.rgb = AZUL_CELESTE
                                        if tag == "<<Tipo>>": run.font.size = Pt(12)
                                        elif tag == "<<Sucursal>>": run.font.size = Pt(14)
                                        else: run.font.size = Pt(11)

                    # --- LÓGICA DE IMÁGENES (REPORTES) ---
                    if modo == "Reportes":
                        tags_foto = ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]
                        for tf in tags_foto:
                            adj = record.get(tf)
                            if adj and isinstance(adj, list):
                                try:
                                    r_img = requests.get(adj[0].get('url'))
                                    if r_img.status_code == 200:
                                        img_io = BytesIO(r_img.content)
                                        for slide in prs.slides:
                                            for shape in slide.shapes:
                                                if (shape.has_text_frame and f"<<{tf}>>" in shape.text) or (tf in shape.name):
                                                    slide.shapes.add_picture(img_io, shape.left, shape.top, shape.width, shape.height)
                                except: pass

                    nom_arch = f"{dt.day} de {MESES_ES[dt.month-1]} de {dt.year} - {f_tipo}, {f_suc} - {procesar_texto_maestro(lugar)}, {record.get('Municipio')}"
                    pp_io = BytesIO(); prs.save(pp_io)
                    data_out = generar_pdf(pp_io.getvalue())
                    if data_out:
                        ext = ".pdf" if modo == "Reportes" else ".jpg"
                        ruta_zip = f"Provident/{dt.year}/{str(dt.month).zfill(2)} - {MESES_ES[dt.month-1]}/{modo}/{f_suc}/{nom_arch[:140]}{ext}"
                        zip_f.writestr(ruta_zip, data_out if modo == "Reportes" else convert_from_bytes(data_out)[0].tobytes())
                    p_bar.progress((i + 1) / total)
            
            status.success(f"✅ ¡{total} archivos generados con éxito!")
            st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_Images_Fixed.zip", use_container_width=True)
