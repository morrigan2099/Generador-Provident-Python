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

# --- CONFIGURACIÓN Y UTILIDADES ---
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]

def cargar_config():
    if os.path.exists("config_app.json"):
        try:
            with open("config_app.json", "r") as f: return json.load(f)
        except: pass
    return {"plantillas": {}, "columnas_visibles": []}

def procesar_texto_maestro(texto, campo=""):
    if not texto or str(texto).lower() == "none": return ""
    if isinstance(texto, list): return texto 
    t = str(texto).strip().replace('\n', ' ').replace('\r', ' ')
    t = re.sub(r'\s+', ' ', t)
    if campo == 'Seccion': return t.upper()
    palabras = t.lower().split()
    if not palabras: return ""
    prep = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al']
    res = [palabras[0].capitalize()]
    for p in palabras[1:]:
        res.append(p if p in prep else p.capitalize())
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
st.set_page_config(page_title="Provident Pro v49", layout="wide")
if 'config' not in st.session_state: st.session_state.config = cargar_config()

st.title("🚀 Generador Pro v49")

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Configuración")
    if st.button("💾 GUARDAR CONFIG"):
        with open("config_app.json", "w") as f: json.dump(st.session_state.config, f)
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
                recs = r_reg.json().get("records", [])
                st.session_state.raw_data_original = recs
                st.session_state.raw_records = [
                    {'id': r['id'], 'fields': {k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v) for k, v in r['fields'].items()}} 
                    for r in recs
                ]
                st.rerun()

# --- ÁREA PRINCIPAL ---
if 'raw_records' in st.session_state:
    st.subheader("1. Configurar Acción")
    modo = st.radio("Formato de salida:", ["Postales", "Reportes"], horizontal=True)
    
    st.subheader("2. Selección de Registros")
    df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    
    c_sel1, c_sel2, _ = st.columns([1, 1, 4])
    if 'select_all' not in st.session_state: st.session_state.select_all = False
    
    if c_sel1.button("✅ Seleccionar Todo"): 
        st.session_state.select_all = True; st.rerun()
    if c_sel2.button("❌ Desmarcar Todo"): 
        st.session_state.select_all = False; st.rerun()

    df_view = df_full.copy()
    for c in df_view.columns:
        if isinstance(df_view[c].iloc[0], list): df_view.drop(c, axis=1, inplace=True)
    
    df_view.insert(0, "Seleccionar", st.session_state.select_all)
    df_edit = st.data_editor(df_view, use_container_width=True, hide_index=True,
                             column_config={"Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False)})
    
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        st.subheader("3. Asignación de Plantillas")
        folder_fisica = os.path.join("Plantillas", modo.upper())
        archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
        tipos_sel = df_view.loc[sel_idx, "Tipo"].unique()
        for t in tipos_sel:
            p_mem = st.session_state.config["plantillas"].get(t)
            idx_def = archivos_pptx.index(p_mem) if p_mem in archivos_pptx else 0
            st.session_state.config["plantillas"][t] = st.selectbox(f"Plantilla {t}:", archivos_pptx, index=idx_def, key=f"p_{t}")

        if st.button("🔥 GENERAR ARCHIVOS", use_container_width=True, type="primary"):
            p_bar = st.progress(0); status = st.empty(); zip_buf = BytesIO()
            total = len(sel_idx)
            AZUL_CELESTE = RGBColor(0, 176, 240)
            
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                for i, idx in enumerate(sel_idx):
                    record = st.session_state.raw_records[idx]['fields']
                    record_orig = st.session_state.raw_data_original[idx]['fields']
                    dt = datetime.strptime(record.get('Fecha'), '%Y-%m-%d')
                    f_tipo = record.get('Tipo'); f_suc = record.get('Sucursal')
                    
                    status.text(f"Procesando {i+1}/{total}: {f_suc}")
                    
                    f_confechor = f"{DIAS_ES[dt.weekday()]} {dt.day} de {MESES_ES[dt.month-1]} de {dt.year}, {record.get('Hora')}"
                    
                    # LOGICA CONDICIONAL PARA CONCAT Y NOMBRE
                    if f_tipo == "Actividad en Sucursal":
                        f_concat = f"Sucursal {f_suc}"
                        nom_arch = f"{dt.day} de {MESES_ES[dt.month-1]} de {dt.year} - {f_tipo}, {f_suc}"
                    else:
                        lugar = record.get('Punto de reunion') or record.get('Ruta a seguir')
                        f_concat = f"{lugar}, {record.get('Municipio')}"
                        nom_arch = f"{dt.day} de {MESES_ES[dt.month-1]} de {dt.year} - {f_tipo}, {f_suc} - {f_concat}"
                    
                    reemplazos = {"<<Tipo>>": f_tipo, "<<Sucursal>>": f_suc, "<<Seccion>>": record.get('Seccion'), 
                                  "<<Confechor>>": f_confechor, "<<Concat>>": f_concat}

                    prs = Presentation(os.path.join(folder_fisica, st.session_state.config["plantillas"][f_tipo]))
                    for slide in prs.slides:
                        for shape in list(slide.shapes):
                            # IMÁGENES
                            txt_b = shape.text_frame.text if shape.has_text_frame else ""
                            tags_foto = ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]
                            for tf in tags_foto:
                                if f"<<{tf}>>" in txt_b or tf == shape.name:
                                    adj = record_orig.get(tf)
                                    if adj and isinstance(adj, list):
                                        try:
                                            img_d = requests.get(adj[0].get('url')).content
                                            slide.shapes.add_picture(BytesIO(img_d), shape.left, shape.top, shape.width, shape.height)
                                            sp = shape._element; sp.getparent().remove(sp)
                                        except: pass
                        # TEXTO (Sucursal 12, Resto 11)
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for tag, val in reemplazos.items():
                                    if tag in shape.text_frame.text:
                                        tf = shape.text_frame; tf.clear()
                                        run = tf.paragraphs[0].add_run()
                                        run.text = str(val); run.font.bold = True; run.font.color.rgb = AZUL_CELESTE
                                        if tag == "<<Sucursal>>": run.font.size = Pt(12)
                                        else: run.font.size = Pt(11)

                    pp_io = BytesIO(); prs.save(pp_io)
                    data_out = generar_pdf(pp_io.getvalue())
                    if data_out:
                        ext = ".pdf" if modo == "Reportes" else ".jpg"
                        ruta_zip = f"Provident/{dt.year}/{str(dt.month).zfill(2)} - {MESES_ES[dt.month-1]}/{modo}/{f_suc}/{nom_arch[:140]}{ext}"
                        zip_f.writestr(ruta_zip, data_out if modo == "Reportes" else convert_from_bytes(data_out)[0].tobytes())
                    p_bar.progress((i + 1) / total)
            
            status.success(f"✅ ¡{total} archivos listos!")
            st.download_button("📥 DESCARGAR", zip_buf.getvalue(), f"Provident_{datetime.now().strftime('%H%M%S')}.zip", use_container_width=True)
else:
    st.info("💡 Carga los datos para comenzar.")
