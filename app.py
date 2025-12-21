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

# --- UTILIDADES DE TEXTO ---
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]

def limpiar_acentos(texto):
    if not texto: return ""
    texto = str(texto)
    # Normalizar y eliminar diacríticos (acentos)
    nfkd_form = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).lower()

def proper_elegante_v2(texto):
    """
    Aplica las reglas:
    - Todo a minúsculas inicial.
    - Palabras pequeñas en minúsculas.
    - Mayúscula al inicio, tras punto '.' o tras '(' .
    - Palabras en paréntesis inician con Mayúscula (si no son pequeñas).
    """
    texto = limpiar_acentos(texto).strip()
    if not texto: return ""
    
    pequenas = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al']
    
    # Tokenizar manteniendo delimitadores (puntos y paréntesis)
    tokens = re.split(r'(\s+|\.|\(|\))', texto)
    resultado = []
    
    for i, t in enumerate(tokens):
        # Si es espacio o delimitador, pasar tal cual
        if re.match(r'\s+|\.|\(|\)', t):
            resultado.append(t)
            continue
        
        # Determinar si debe ir en Mayúscula por posición
        forzar_mayuscula = False
        if i == 0: forzar_mayuscula = True
        else:
            # Buscar el último token no vacío hacia atrás
            prev = "".join(tokens[:i]).strip()
            if prev.endswith('.') or prev.endswith('('):
                forzar_mayuscula = True
        
        if forzar_mayuscula:
            resultado.append(t.capitalize())
        elif t in pequenas:
            resultado.append(t.lower())
        else:
            resultado.append(t.capitalize())
            
    return "".join(resultado)

def format_fecha_es(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        return f"{dt.day} de {MESES_ES[dt.month-1]} de {dt.year}"
    except: return fecha_str

# --- PARÁMETROS STREAMLIT ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"
AZUL_CELESTE = RGBColor(0, 176, 240)

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
st.set_page_config(page_title="Provident Pro v23", layout="wide")
st.title("🚀 Generador Pro: Proper Elegante y Reglas de Estilo")

with st.sidebar:
    st.header("🔌 Conexión")
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
    
    if st.button("💾 GUARDAR CONFIGURACIÓN"):
        guardar_config_json(st.session_state.config)
        st.toast("Configuración guardada")

# --- CUERPO ---
modo = st.radio("1. Formato:", ["Postales", "Reportes"], horizontal=True)
folder_fisica = os.path.join(BASE_DIR, modo.upper())
archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]

if 'raw_records' in st.session_state and st.session_state.raw_records:
    # Procesar dataframe con reglas de visualización
    data_list = []
    for r in st.session_state.raw_records:
        fields = r['fields'].copy()
        # Aplicar formato de fecha para la tabla
        if 'Fecha' in fields: fields['Fecha'] = format_fecha_es(fields['Fecha'])
        data_list.append(fields)
    
    df = pd.DataFrame(data_list)
    all_keys = list(df.columns)
    
    saved_cols = st.session_state.config.get("columnas_visibles", [])
    valid_cols = [c for c in saved_cols if c in all_keys] or all_keys[:8]
    cols_to_show = st.multiselect("Columnas visibles:", all_keys, default=valid_cols)
    st.session_state.config["columnas_visibles"] = cols_to_show

    df_display = df[[c for c in cols_to_show if c in df.columns]].copy()
    if "Tipo" not in df_display.columns: df_display["Tipo"] = df["Tipo"]
    df_display.insert(0, "Seleccionar", False)

    df_edit = st.data_editor(
        df_display,
        column_config={"Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False)},
        use_container_width=True, hide_index=True
    )

    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        st.subheader("2. Plantillas")
        tipos_sel = df_edit.loc[sel_idx, "Tipo"].unique()
        c1, c2 = st.columns(2)
        for i, t in enumerate(tipos_sel):
            t_str = str(t)
            p_mem = st.session_state.config["plantillas"].get(t_str)
            with (c1 if i % 2 == 0 else c2):
                idx_def = archivos_pptx.index(p_mem) if p_mem in archivos_pptx else 0
                sel_p = st.selectbox(f"Tipo: {t_str}", archivos_pptx, index=idx_def, key=f"sel_{t_str}")
                st.session_state.config["plantillas"][t_str] = sel_p

        if st.button("🔥 GENERAR", use_container_width=True, type="primary"):
            p_bar = st.progress(0); status = st.empty(); zip_buf = BytesIO()
            
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                total = len(sel_idx)
                for i, idx in enumerate(sel_idx):
                    record = st.session_state.raw_records[idx]['fields']
                    
                    # --- PROCESAMIENTO DE DATOS SEGÚN REGLAS ---
                    f_tipo = proper_elegante_v2(record.get('Tipo', ''))
                    f_suc = proper_elegante_v2(record.get('Sucursal', ''))
                    f_seccion = str(record.get('Seccion', '')).upper()
                    f_muni = proper_elegante_v2(record.get('Municipio', ''))
                    f_punto = proper_elegante_v2(record.get('Punto de reunion', ''))
                    f_ruta = proper_elegante_v2(record.get('Ruta a seguir', ''))
                    
                    # Fecha y Hora
                    dt = datetime.strptime(record.get('Fecha', '2025-01-01'), '%Y-%m-%d')
                    hora_f = str(record.get('Hora', '')).strip().lower()
                    
                    confechor = f"{DIAS_ES[dt.weekday()]} {dt.day} de {MESES_ES[dt.month-1]} de {dt.year}, {hora_f}"
                    # Aplicar proper al confechor (excepto hora que ya es minúscula)
                    confechor_final = proper_elegante_v2(confechor)
                    
                    lugar = f_punto if f_punto else f_ruta
                    concat_val = proper_elegante_v2(f"{lugar}, {f_muni}")
                    
                    reemplazos = {
                        "<<Tipo>>": f_tipo, "<<Sucursal>>": f_suc, 
                        "<<Seccion>>": f_seccion, "<<Confechor>>": confechor_final, 
                        "<<Concat>>": concat_val
                    }
                    
                    # --- PPTX ---
                    plantilla = st.session_state.config["plantillas"].get(record.get('Tipo'))
                    prs = Presentation(os.path.join(folder_fisica, plantilla))
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for tag, val in reemplazos.items():
                                    if tag in shape.text_frame.text:
                                        tf = shape.text_frame
                                        tf.auto_size = None
                                        tf.clear()
                                        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
                                        run = p.add_run(); run.text = val; run.font.bold = True; run.font.color.rgb = AZUL_CELESTE
                                        # REGLA PERSONALIZADA: Tipo de mayor a menor (64pts iniciales)
                                        # (Ajuste basado en tu instrucción de memoria: mayor a menor hasta 2 líneas)
                                        if tag == "<<Tipo>>":
                                            run.font.size = Pt(64) 
                                        elif tag == "<<Sucursal>>":
                                            run.font.size = Pt(14)
                                        else:
                                            run.font.size = Pt(11)

                    # --- NOMBRE DE ARCHIVO (Reglas aplicadas) ---
                    f_str_nom = f"{dt.day} de {MESES_ES[dt.month-1]} de {dt.year}"
                    lugar_nom = proper_elegante_v2(lugar)
                    nom_file = f"{f_str_nom} - {f_tipo}, {f_suc} - {lugar_nom}, {f_muni}"
                    
                    # Generación y ZIP
                    pp_io = BytesIO(); prs.save(pp_io)
                    data_out = generar_pdf(pp_io.getvalue())
                    if data_out:
                        ext = ".pdf" if modo == "Reportes" else ".jpg"
                        mes_folder = f"{str(dt.month).zfill(2)} - {MESES_ES[dt.month-1]}"
                        ruta_zip = f"Provident/{dt.year}/{mes_folder}/{modo}/{f_suc}/{nom_file[:150]}{ext}"
                        zip_f.writestr(ruta_zip, data_out if modo == "Reportes" else convert_from_bytes(data_out)[0].tobytes())
                    
                    p_bar.progress((i+1)/total)

            st.success("✅ ¡Proceso completado con reglas de estilo elegantes!")
            st.download_button("📥 DESCARGAR", zip_buf.getvalue(), "Provident_Elegante.zip", use_container_width=True)
