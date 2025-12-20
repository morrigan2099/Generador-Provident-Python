import streamlit as st
import requests
import pandas as pd
import json
import os
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_VERTICAL_ANCHOR
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
st.set_page_config(page_title="Provident Pro v21", layout="wide")
st.title("🚀 Generador Pro: Fuente 50pt Real (Sin AutoFit)")

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
                st.session_state.config = cargar_config()
                st.rerun()
    
    st.divider()
    st.header("📂 Vínculos Guardados")
    if st.button("💾 GUARDAR JSON", use_container_width=True, type="primary"):
        guardar_config_json(st.session_state.config)
        st.toast("Configuración guardada")

    for t_key, p_val in list(st.session_state.config["plantillas"].items()):
        c1, c2 = st.columns([4, 1])
        c1.caption(f"**{t_key}**")
        if c2.button("🗑️", key=f"del_{t_key}"):
            del st.session_state.config["plantillas"][t_key]
            guardar_config_json(st.session_state.config)
            st.rerun()

# --- CUERPO PRINCIPAL ---
modo = st.radio("1. Selecciona el formato de salida:", ["Postales", "Reportes"], horizontal=True)
folder_fisica = os.path.join(BASE_DIR, modo.upper())
archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]

if 'raw_records' in st.session_state and st.session_state.raw_records:
    st.subheader("2. Selecciona los registros a generar")
    all_keys = []
    for r in st.session_state.raw_records:
        for key in r['fields'].keys():
            if key not in all_keys: all_keys.append(key)
    
    df = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    saved_cols = st.session_state.config.get("columnas_visibles", [])
    valid_cols = [c for c in saved_cols if c in all_keys] or all_keys[:8]
    
    cols_to_show = st.multiselect("Columnas de la tabla:", all_keys, default=valid_cols)
    st.session_state.config["columnas_visibles"] = cols_to_show

    df_display = df[[c for c in cols_to_show if c in df.columns]].copy()
    if "Tipo" not in df_display.columns:
        df_display["Tipo"] = df["Tipo"]

    df_display.insert(0, "Seleccionar", False)

    df_edit = st.data_editor(
        df_display,
        column_config={
            "Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False, required=True)
        },
        use_container_width=True,
        hide_index=True
    )

    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        st.subheader("3. Verifica las plantillas vinculadas")
        tipos_sel = df_edit.loc[sel_idx, "Tipo"].unique()
        c1, c2 = st.columns(2)
        for i, t in enumerate(tipos_sel):
            t_str = str(t)
            p_mem = st.session_state.config["plantillas"].get(t_str)
            with (c1 if i % 2 == 0 else c2):
                idx_def = archivos_pptx.index(p_mem) if p_mem in archivos_pptx else 0
                sel_p = st.selectbox(f"Plantilla para {t_str}:", archivos_pptx, index=idx_def, key=f"sel_{t_str}")
                st.session_state.config["plantillas"][t_str] = sel_p

        if st.button("🔥 GENERAR ARCHIVOS", use_container_width=True, type="primary"):
            p_bar = st.progress(0)
            status = st.empty()
            zip_buf = BytesIO()
            
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                total = len(sel_idx)
                for i, idx in enumerate(sel_idx):
                    record = st.session_state.raw_records[idx]['fields']
                    tipo_rec = record.get('Tipo')
                    plantilla_final = st.session_state.config["plantillas"].get(tipo_rec)
                    
                    dt = datetime.strptime(record.get('Fecha', '2025-01-01'), '%Y-%m-%d')
                    f_suc, f_muni, f_tipo = [proper_elegante(record.get(k, '')) for k in ['Sucursal', 'Municipio', 'Tipo']]
                    f_punto, f_ruta = str(record.get('Punto de reunion', '')).strip(), str(record.get('Ruta a seguir', '')).strip()
                    lugar_corto = proper_elegante(f_punto if f_punto else f_ruta)
                    
                    status.text(f"Procesando {i+1}/{total}: {f_tipo} - {f_suc}")
                    p_bar.progress((i + 1) / total)

                    hora_f = str(record.get('Hora', '')).strip()
                    confechor = f"{DIAS_ES[dt.weekday()]} {MESES_ES[dt.month-1]} {str(dt.day).zfill(2)} de {dt.year}, {hora_f}"
                    concat_val = ", ".join([p for p in [f_punto if f_punto else f_ruta, f_muni] if p])

                    reemplazos = {
                        "<<Tipo>>": f_tipo, "<<Sucursal>>": f_suc, 
                        "<<Seccion>>": str(record.get('Seccion', '')).upper(),
                        "<<Confechor>>": proper_elegante(confechor), "<<Concat>>": proper_elegante(concat_val)
                    }
                    
                    prs = Presentation(os.path.join(folder_fisica, plantilla_final))
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for tag, val in reemplazos.items():
                                    if tag in shape.text_frame.text:
                                        # ACCIÓN CRÍTICA: Desactivar AutoFit para forzar 50pt
                                        text_frame = shape.text_frame
                                        text_frame.auto_size = None # Elimina ajuste automático
                                        text_frame.clear()
                                        
                                        p = text_frame.paragraphs[0]
                                        p.alignment = PP_ALIGN.CENTER
                                        run = p.add_run()
                                        run.text = val
                                        run.font.bold = True
                                        run.font.color.rgb = AZUL_CELESTE
                                        run.font.size = Pt(50) # FORZADO A 50PT

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

                    pp_io = BytesIO(); prs.save(pp_io)
                    data_out = generar_pdf(pp_io.getvalue())
                    if data_out:
                        f_str = f"{MESES_ES[dt.month-1]} {str(dt.day).zfill(2)} de {dt.year}"
                        nom_file = f"{f_str} - {f_tipo}, {f_suc} - {lugar_corto}, {f_muni}"
                        ext = ".pdf" if modo == "Reportes" else ".jpg"
                        mes_folder = f"{str(dt.month).zfill(2)} - {MESES_ES[dt.month-1]}"
                        ruta_zip = f"Provident/{dt.year}/{mes_folder}/{modo}/{f_suc}/{nom_file[:130]}{ext}"
                        contenido = data_out if modo == "Reportes" else convert_from_bytes(data_out)[0].tobytes()
                        zip_f.writestr(ruta_zip, contenido)

            status.success(f"✅ ¡Completado! {total} archivos procesados con fuente 50pt.")
            st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), f"Provident_{datetime.now().strftime('%H%M')}.zip", use_container_width=True)
