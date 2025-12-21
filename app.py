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

# --- CONFIGURACIÓN ---
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

# --- MOTOR DE TEXTO SEGURO (REPARADO) ---
def procesar_texto_maestro(texto, campo=""):
    if not texto or str(texto).lower() == "none": return ""
    if isinstance(texto, list): return texto 
    
    # Limpieza básica
    t = str(texto).strip()
    t = re.sub(r'\s+', ' ', t)
    
    if campo == 'Hora': return t.lower()
    if campo == 'Seccion': return t.upper()

    # Capitalización inteligente simple
    pequenas = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al']
    palabras = t.lower().split()
    res = []
    for i, p in enumerate(palabras):
        if i == 0 or p not in pequenas:
            res.append(p.capitalize())
        else:
            res.append(p)
    return " ".join(res)

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
st.set_page_config(page_title="Provident Pro v39", layout="wide")
st.title("🚀 Generador Pro: Reparación Total y Tipo 11pts")

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

with st.sidebar:
    st.header("⚙️ Configuración")
    if st.button("💾 GUARDAR JSON", use_container_width=True, type="primary"):
        guardar_config_json(st.session_state.config)
        st.toast("Configuración guardada en JSON")
    
    st.divider()
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        bases = r_bases.json().get('bases', [])
        b_names = {b['name']: b['id'] for b in bases}
        base_sel = st.selectbox("Base:", [""] + list(b_names.keys()))
        
        if base_sel:
            bid = b_names[base_sel]
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{bid}/tables", headers=headers)
            tablas = r_tab.json().get('tables', [])
            t_names = {t['name']: t['id'] for t in tablas}
            tabla_sel = st.selectbox("Tabla:", list(t_names.keys()))
            
            if st.button("🔄 CARGAR DATOS"):
                r_reg = requests.get(f"https://api.airtable.com/v0/{bid}/{t_names[tabla_sel]}", headers=headers)
                recs = r_reg.json().get("records", [])
                st.session_state.raw_data_original = recs
                st.session_state.raw_records = [
                    {'id': r['id'], 'fields': {k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v) for k, v in r['fields'].items()}} 
                    for r in recs
                ]
                st.rerun()

# --- TABLA Y PROCESAMIENTO ---
if 'raw_records' in st.session_state:
    df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    all_cols = list(df_full.columns)
    
    with st.sidebar:
        st.divider()
        def_cols = [c for c in st.session_state.config.get("columnas_visibles", []) if c in all_cols] or all_cols
        selected_cols = st.multiselect("Columnas Visibles:", all_cols, default=def_cols)
        st.session_state.config["columnas_visibles"] = selected_cols

    df_view = df_full[[c for c in selected_cols if c in df_full.columns]].copy()
    for c in df_view.columns:
        if not df_view.empty and isinstance(df_view[c].iloc[0], list): df_view.drop(c, axis=1, inplace=True)
    
    df_view.insert(0, "Seleccionar", False)
    
    df_edit = st.data_editor(
        df_view, use_container_width=True, hide_index=True,
        column_config={"Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False)}
    )
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        modo = st.radio("Acción:", ["Postales", "Reportes"], horizontal=True)
        folder_fisica = os.path.join("Plantillas", modo.upper())
        AZUL_CELESTE = RGBColor(0, 176, 240)
        
        archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
        tipos_sel = df_view.loc[sel_idx, "Tipo"].unique() if "Tipo" in df_view.columns else []
        for t in tipos_sel:
            p_mem = st.session_state.config["plantillas"].get(t)
            idx_def = archivos_pptx.index(p_mem) if p_mem in archivos_pptx else 0
            st.session_state.config["plantillas"][t] = st.selectbox(f"Plantilla {t}:", archivos_pptx, index=idx_def, key=t)

        if st.button("🔥 GENERAR ARCHIVOS", use_container_width=True, type="primary"):
            p_bar = st.progress(0); zip_buf = BytesIO()
            total = len(sel_idx)
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                for i, idx in enumerate(sel_idx):
                    record = st.session_state.raw_records[idx]['fields']
                    record_orig = st.session_state.raw_data_original[idx]['fields']
                    
                    # Datos para reemplazo
                    f_tipo = record.get('Tipo')
                    f_suc = record.get('Sucursal')
                    
                    reemplazos = {
                        "<<Tipo>>": f_tipo, 
                        "<<Sucursal>>": f_suc, 
                        "<<Seccion>>": record.get('Seccion'),
                        "<<Confechor>>": f"{record.get('Fecha')}, {record.get('Hora')}",
                        "<<Concat>>": f"{record.get('Punto de reunion') or record.get('Ruta a seguir')}, {record.get('Municipio')}"
                    }

                    prs = Presentation(os.path.join(folder_fisica, st.session_state.config["plantillas"][f_tipo]))
                    for slide in prs.slides:
                        # Fotos
                        for shape in list(slide.shapes):
                            txt = shape.text_frame.text if shape.has_text_frame else ""
                            for tf in ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]:
                                if f"<<{tf}>>" in txt or tf == shape.name:
                                    adj = record_orig.get(tf)
                                    if adj and isinstance(adj, list):
                                        try:
                                            img_data = requests.get(adj[0].get('url')).content
                                            slide.shapes.add_picture(BytesIO(img_data), shape.left, shape.top, shape.width, shape.height)
                                            sp = shape._element; sp.getparent().remove(sp)
                                        except: pass

                        # Texto (Tamaños exactos solicitados)
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for tag, val in reemplazos.items():
                                    if tag in shape.text_frame.text:
                                        tf = shape.text_frame; tf.clear()
                                        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
                                        run = p.add_run(); run.text = str(val); run.font.bold = True; run.font.color.rgb = AZUL_CELESTE
                                        
                                        if tag == "<<Tipo>>": run.font.size = Pt(11) # <--- FIJADO EN 11pts
                                        elif tag == "<<Sucursal>>": run.font.size = Pt(14)
                                        else: run.font.size = Pt(11)

                    pp_io = BytesIO(); prs.save(pp_io)
                    data_out = generar_pdf(pp_io.getvalue())
                    if data_out:
                        ext = ".pdf" if modo == "Reportes" else ".jpg"
                        nombre_f = f"{record.get('Fecha')} - {f_tipo} - {f_suc}{ext}"
                        zip_f.writestr(f"{modo}/{f_suc}/{nombre_f}", data_out if modo == "Reportes" else convert_from_bytes(data_out)[0].tobytes())
                    p_bar.progress((i + 1) / total)
            
            st.success("✅ ¡Archivos generados con éxito!")
            st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_v39.zip", use_container_width=True)
