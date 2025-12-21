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

# --- MOTOR DE TEXTO MAESTRO ---
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]

def procesar_texto_maestro(texto, campo=""):
    if not texto or str(texto).lower() == "none": return ""
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
        forzar = (i == 0)
        if not forzar:
            previo = "".join(tokens[:i]).strip()
            if previo.endswith('.') or previo.endswith('('): forzar = True
        
        if forzar: resultado.append(t.capitalize())
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
st.set_page_config(page_title="Provident Pro v36", layout="wide")
st.title("🚀 Generador Pro: Auto-ajuste de Fuente y Fix API")

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

with st.sidebar:
    st.header("⚙️ Configuración")
    if st.button("💾 GUARDAR CONFIGURACIÓN", use_container_width=True):
        guardar_config_json(st.session_state.config)
        st.toast("Configuración guardada")
    
    st.divider()
    try:
        r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
        if r_bases.status_code == 200:
            base_opts = {b['name']: b['id'] for b in r_bases.json().get('bases', [])}
            base_sel = st.selectbox("Base:", [""] + list(base_opts.keys()))
            if base_sel:
                base_id = base_opts[base_sel]
                r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_id}/tables", headers=headers)
                tablas_data = r_tab.json().get('tables')
                
                if tablas_data:
                    tabla_opts = {t['name']: t['id'] for t in tablas_data}
                    tabla_sel = st.selectbox("Tabla:", list(tabla_opts.keys()))
                    if st.button("🔄 CARGAR DATOS"):
                        r_reg = requests.get(f"https://api.airtable.com/v0/{base_id}/{tabla_sel}", headers=headers)
                        recs = r_reg.json().get("records", [])
                        st.session_state.raw_data_original = recs
                        st.session_state.raw_records = [
                            {'id': r['id'], 'fields': {k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v) for k, v in r['fields'].items()}} 
                            for r in recs
                        ]
                        st.rerun()
                else:
                    st.error("No se pudieron cargar las tablas. Revisa los scopes del Token.")
    except Exception as e:
        st.error(f"Error de conexión: {e}")

if 'raw_records' in st.session_state:
    df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    all_cols = list(df_full.columns)
    
    with st.sidebar:
        st.divider()
        def_cols = [c for c in st.session_state.config.get("columnas_visibles", []) if c in all_cols] or all_cols
        selected_cols = st.multiselect("Columnas:", all_cols, default=def_cols)
        st.session_state.config["columnas_visibles"] = selected_cols

    df_view = df_full[[c for c in selected_cols if c in df_full.columns]].copy()
    for c in df_view.columns:
        if len(df_view) > 0 and isinstance(df_view[c].iloc[0], list): df_view.drop(c, axis=1, inplace=True)
    
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

        if st.button("🔥 GENERAR", use_container_width=True, type="primary"):
            p_bar = st.progress(0); status = st.empty(); zip_buf = BytesIO()
            total = len(sel_idx)
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                for i, idx in enumerate(sel_idx):
                    record = st.session_state.raw_records[idx]['fields']
                    record_orig = st.session_state.raw_data_original[idx]['fields']
                    dt = datetime.strptime(record.get('Fecha'), '%Y-%m-%d')
                    f_tipo = record.get('Tipo'); f_suc = record.get('Sucursal')
                    status.text(f"Procesando {i+1}/{total}: {f_tipo}")

                    reemplazos = {
                        "<<Tipo>>": f_tipo, "<<Sucursal>>": f_suc, 
                        "<<Seccion>>": record.get('Seccion'), 
                        "<<Confechor>>": procesar_texto_maestro(f"{DIAS_ES[dt.weekday()]} {dt.day} de {MESES_ES[dt.month-1]} de {dt.year}, {record.get('Hora')}"),
                        "<<Concat>>": procesar_texto_maestro(f"{record.get('Punto de reunion') or record.get('Ruta a seguir')}, {record.get('Municipio')}")
                    }

                    prs = Presentation(os.path.join(folder_fisica, st.session_state.config["plantillas"][f_tipo]))
                    for slide in prs.slides:
                        # IMÁGENES
                        for shape in list(slide.shapes):
                            txt_b = shape.text_frame.text if shape.has_text_frame else ""
                            for tf in ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]:
                                if f"<<{tf}>>" in txt_b or tf == shape.name:
                                    adj = record_orig.get(tf)
                                    if adj and isinstance(adj, list):
                                        try:
                                            img_d = requests.get(adj[0].get('url')).content
                                            slide.shapes.add_picture(BytesIO(img_d), shape.left, shape.top, shape.width, shape.height)
                                            sp = shape._element; sp.getparent().remove(sp)
                                        except: pass

                        # TEXTO CON LÓGICA DE FUENTE 64PT -> 2 LÍNEAS
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for tag, val in reemplazos.items():
                                    if tag in shape.text_frame.text:
                                        tf = shape.text_frame; tf.clear()
                                        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
                                        run = p.add_run(); run.text = str(val); run.font.bold = True; run.font.color.rgb = AZUL_CELESTE
                                        
                                        if tag == "<<Tipo>>":
                                            current_size = 64
                                            run.font.size = Pt(current_size)
                                            # Reducción de fuente hasta que quepa en 2 líneas
                                            # (Simulado mediante longitud de caracteres para evitar loop infinito)
                                            while len(str(val)) * (current_size/20) > 60 and current_size > 20:
                                                current_size -= 4
                                                run.font.size = Pt(current_size)
                                        elif tag == "<<Sucursal>>": run.font.size = Pt(14)
                                        else: run.font.size = Pt(11)

                    nom_arch = f"{dt.day} de {MESES_ES[dt.month-1]} - {f_tipo}, {f_suc}"
                    pp_io = BytesIO(); prs.save(pp_io)
                    data_out = generar_pdf(pp_io.getvalue())
                    if data_out:
                        ext = ".pdf" if modo == "Reportes" else ".jpg"
                        ruta_zip = f"{modo}/{f_suc}/{nom_arch[:100]}{ext}"
                        zip_f.writestr(ruta_zip, data_out if modo == "Reportes" else convert_from_bytes(data_out)[0].tobytes())
                    p_bar.progress((i + 1) / total)
            
            status.success("✅ ¡Listo!")
            st.download_button("📥 DESCARGAR", zip_buf.getvalue(), "Provident_v36.zip", use_container_width=True)
