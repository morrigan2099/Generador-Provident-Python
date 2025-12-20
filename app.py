import streamlit as st
import requests
import pandas as pd
import json
import os
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import subprocess, tempfile, zipfile
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN DE PERSISTENCIA ---
CONFIG_FILE = "config_app.json"

def cargar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f: return json.load(f)
        except: pass
    return {"plantillas": {}, "columnas_visibles": []}

def guardar_config(config):
    with open(CONFIG_FILE, "w") as f: json.dump(config, f, indent=4)

config = cargar_config()

# --- PARÁMETROS ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"
AZUL_CELESTE = RGBColor(0, 176, 240)
MESES_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
DIAS_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

def proper_elegante(texto):
    if not texto or str(texto).lower() == "none": return ""
    texto = str(texto).strip().lower()
    palabras = texto.split()
    excepciones = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del']
    resultado = [p.capitalize() if i == 0 or p not in excepciones else p for i, p in enumerate(palabras)]
    return " ".join(resultado)

def interpretar_hora(hora_txt):
    if not hora_txt: return ""
    hora_txt = str(hora_txt).strip().lower()
    for fmt in ["%H:%M", "%I:%M %p", "%H%M", "%I%p"]:
        try:
            dt_hora = datetime.strptime(hora_txt.replace(" ", ""), fmt.replace(" ", ""))
            return dt_hora.strftime("%I:%M %p").lower().replace("am", "a.m.").replace("pm", "p.m.")
        except: continue
    return hora_txt

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
st.set_page_config(page_title="Provident Pro v7", layout="wide")
st.title("🚀 Generador Pro: Confechor y Concat Restaurados")

with st.sidebar:
    st.header("🔌 Conexión Airtable")
    headers = {"Authorization": f"Bearer {TOKEN}"}
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
                st.session_state.raw_records = r_reg.json().get("records", [])
                st.rerun()

if 'raw_records' in st.session_state and st.session_state.raw_records:
    record_fields = st.session_state.raw_records[0]['fields']
    all_keys = list(record_fields.keys())
    try:
        idx_end = all_keys.index("Foto de equipo")
        cols_utiles = all_keys[:idx_end]
    except: cols_utiles = all_keys

    df = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    cols_to_show = st.multiselect("Visualizar campos:", cols_utiles, default=config.get("columnas_visibles", cols_utiles))
    
    if cols_to_show != config["columnas_visibles"]:
        config["columnas_visibles"] = cols_to_show
        guardar_config(config)

    df_display = df[cols_to_show].copy()
    df_display.insert(0, "Seleccionar", False)
    df_edit = st.data_editor(df_display, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        modo = st.radio("Acción:", ["Postales", "Reportes"], horizontal=True)
        folder_fisica = os.path.join(BASE_DIR, modo.upper())
        archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
        
        tipos_sel = df_edit.loc[sel_idx, "Tipo"].unique()
        for t in tipos_sel:
            default_p = config["plantillas"].get(t, archivos_pptx[0] if archivos_pptx else "")
            nueva_p = st.selectbox(f"Plantilla para {t}:", archivos_pptx, index=archivos_pptx.index(default_p) if default_p in archivos_pptx else 0, key=f"p_{t}")
            if nueva_p != config["plantillas"].get(t):
                config["plantillas"][t] = nueva_p
                guardar_config(config)

        if st.button("🔥 GENERAR"):
            p_bar = st.progress(0); s_text = st.empty()
            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                for i, idx in enumerate(sel_idx):
                    record = st.session_state.raw_records[idx]['fields']
                    p_bar.progress(i / len(sel_idx))
                    
                    dt = datetime.strptime(record.get('Fecha', '2025-01-01'), '%Y-%m-%d')
                    f_suc, f_muni, f_tipo = [proper_elegante(record.get(k, '')) for k in ['Sucursal', 'Municipio', 'Tipo']]
                    f_punto, f_ruta = record.get('Punto de reunion', ''), record.get('Ruta a seguir', '')
                    lugar_corto = proper_elegante(min([o for o in [f_punto, f_ruta] if o], key=len) if (f_punto or f_ruta) else "")

                    # LÓGICA DE CONFECHOR Y CONCAT
                    hora_f = interpretar_hora(record.get('Hora', ''))
                    confechor = f"{DIAS_ES[dt.weekday()]} {MESES_ES[dt.month-1]} {str(dt.day).zfill(2)} de {dt.year}, {hora_f}"
                    concat_val = f"{f_punto}, {f_ruta}, {f_muni}"

                    reemplazos = {
                        "<<Tipo>>": f_tipo, 
                        "<<Sucursal>>": f_suc, 
                        "<<Seccion>>": str(record.get('Seccion', '')).upper(),
                        "<<Confechor>>": proper_elegante(confechor),
                        "<<Concat>>": proper_elegante(concat_val)
                    }
                    
                    prs = Presentation(os.path.join(folder_fisica, config["plantillas"][record.get('Tipo')]))
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for tag, val in reemplazos.items():
                                    if tag in shape.text_frame.text:
                                        shape.text_frame.clear()
                                        p = shape.text_frame.paragraphs[0]
                                        p.alignment = PP_ALIGN.CENTER
                                        run = p.add_run(); run.text = val; run.font.bold = True; run.font.color.rgb = AZUL_CELESTE
                                        
                                        # REGLA DE ORO: Tipo a 64pt
                                        if tag == "<<Tipo>>":
                                            run.font.size = Pt(64)
                                        else:
                                            run.font.size = Pt(36)
                                        shape.text_frame.word_wrap = True

                    if modo == "Reportes":
                        tags_foto = ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]
                        for tf in tags_foto:
                            adj = record.get(tf)
                            if adj and isinstance(adj, list):
                                r_img = requests.get(adj[0].get('url'))
                                if r_img.status_code == 200:
                                    for slide in prs.slides:
                                        for shape in slide.shapes:
                                            if (shape.has_text_frame and f"<<{tf}>>" in shape.text) or (tf in shape.name):
                                                slide.shapes.add_picture(BytesIO(r_img.content), shape.left, shape.top, shape.width, shape.height)

                    pp_io = BytesIO(); prs.save(pp_io)
                    pdf_data = generar_pdf(pp_io.getvalue())
                    if pdf_data:
                        f_str = f"{MESES_ES[dt.month-1]} {str(dt.day).zfill(2)} de {dt.year}"
                        nom_file = f"{f_str} - {f_tipo}, {f_suc} - {lugar_corto}, {f_muni}"
                        ext = ".pdf" if modo == "Reportes" else ".jpg"
                        
                        mes_folder = f"{str(dt.month).zfill(2)} - {MESES_ES[dt.month-1]}"
                        ruta = f"Provident/{dt.year}/{mes_folder}/{modo}/{f_suc}/{nom_file[:140]}{ext}"
                        zip_f.writestr(ruta, pdf_data if modo == "Reportes" else convert_from_bytes(pdf_data)[0].tobytes())

            st.success("✅ Generación finalizada con todos los campos restaurados.")
            st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_Pro_v7.zip")
