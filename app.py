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
    st.sidebar.success("✅ Configuración guardada en JSON")

if 'config' not in st.session_state:
    st.session_state.config = cargar_config()

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
st.set_page_config(page_title="Provident Pro v12", layout="wide")
st.title("🚀 Generador Pro: Placeholders 50pt y Guardado JSON")

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
    
    st.divider()
    if st.button("💾 GUARDAR CONFIGURACIÓN ACTUAL", use_container_width=True, type="primary"):
        guardar_config_json(st.session_state.config)

if 'raw_records' in st.session_state and st.session_state.raw_records:
    # Detección de columnas
    all_ordered_keys = []
    for r in st.session_state.raw_records:
        for key in r['fields'].keys():
            if key not in all_ordered_keys: all_ordered_keys.append(key)
    
    df = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    
    # Visualización
    saved_cols = st.session_state.config.get("columnas_visibles", [])
    valid_defaults = [c for c in saved_cols if c in all_ordered_keys]
    if not valid_defaults:
        try:
            idx_foto = all_ordered_keys.index("Foto de equipo")
            valid_defaults = all_ordered_keys[:idx_foto]
        except: valid_defaults = all_ordered_keys

    cols_to_show = st.multiselect("Columnas visibles:", all_ordered_keys, default=valid_defaults)
    st.session_state.config["columnas_visibles"] = cols_to_show

    # Asegurar columnas críticas
    cols_para_df = list(cols_to_show)
    for critical in ["Tipo", "Fecha", "Sucursal"]:
        if critical not in cols_para_df: cols_para_df.append(critical)

    df_display = df[[c for c in cols_para_df if c in df.columns]].copy()
    df_display.insert(0, "Seleccionar", False)
    
    column_config = {c: None for c in ["Tipo", "Fecha", "Sucursal"] if c not in cols_to_show}

    df_edit = st.data_editor(df_display, use_container_width=True, hide_index=True, column_config=column_config)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        modo = st.radio("Modo:", ["Postales", "Reportes"], horizontal=True)
        folder_fisica = os.path.join(BASE_DIR, modo.upper())
        archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
        
        tipos_sel = df_edit.loc[sel_idx, "Tipo"].unique()
        for t in tipos_sel:
            t_str = str(t)
            default_p = st.session_state.config["plantillas"].get(t_str, archivos_pptx[0] if archivos_pptx else "")
            nueva_p = st.selectbox(f"Plantilla para {t_str}:", archivos_pptx, 
                                 index=archivos_pptx.index(default_p) if default_p in archivos_pptx else 0, key=f"p_{t_str}")
            st.session_state.config["plantillas"][t_str] = nueva_p

        if st.button("🔥 GENERAR ARCHIVOS"):
            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                for idx in sel_idx:
                    record = st.session_state.raw_records[idx]['fields']
                    dt = datetime.strptime(record.get('Fecha', '2025-01-01'), '%Y-%m-%d')
                    f_suc, f_muni, f_tipo = [proper_elegante(record.get(k, '')) for k in ['Sucursal', 'Municipio', 'Tipo']]
                    
                    f_punto, f_ruta = str(record.get('Punto de reunion', '')).strip(), str(record.get('Ruta a seguir', '')).strip()
                    lugar_corto = proper_elegante(f_punto if f_punto else f_ruta)
                    hora_f = interpretar_hora(record.get('Hora', ''))
                    confechor = f"{DIAS_ES[dt.weekday()]} {MESES_ES[dt.month-1]} {str(dt.day).zfill(2)} de {dt.year}, {hora_f}"
                    concat_val = ", ".join([p for p in [f_punto if f_punto else f_ruta, f_muni] if p])

                    reemplazos = {
                        "<<Tipo>>": f_tipo, "<<Sucursal>>": f_suc, 
                        "<<Seccion>>": str(record.get('Seccion', '')).upper(),
                        "<<Confechor>>": proper_elegante(confechor), "<<Concat>>": proper_elegante(concat_val)
                    }
                    
                    p_nombre = st.session_state.config["plantillas"].get(record.get('Tipo'), archivos_pptx[0])
                    prs = Presentation(os.path.join(folder_fisica, p_nombre))
                    
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for tag, val in reemplazos.items():
                                    if tag in shape.text_frame.text:
                                        shape.text_frame.clear()
                                        p = shape.text_frame.paragraphs[0]
                                        p.alignment = PP_ALIGN.CENTER
                                        run = p.add_run(); run.text = val; run.font.bold = True; run.font.color.rgb = AZUL_CELESTE
                                        # REGLA: Tipo 64pt, demás placeholders 50pt
                                        run.font.size = Pt(64) if tag == "<<Tipo>>" else Pt(50)
                                        shape.text_frame.word_wrap = True

                    # (Lógica de fotos y PDF omitida aquí por brevedad, igual a v11)
                    pp_io = BytesIO(); prs.save(pp_io)
                    pdf_data = generar_pdf(pp_io.getvalue())
                    if pdf_data:
                        f_str = f"{MESES_ES[dt.month-1]} {str(dt.day).zfill(2)} de {dt.year}"
                        nom_file = f"{f_str} - {f_tipo}, {f_suc} - {lugar_corto}, {f_muni}"
                        mes_folder = f"{str(dt.month).zfill(2)} - {MESES_ES[dt.month-1]}"
                        ruta = f"Provident/{dt.year}/{mes_folder}/{modo}/{f_suc}/{nom_file[:135]}.pdf"
                        zip_f.writestr(ruta, pdf_data)

            st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_Pro_v12.zip")
