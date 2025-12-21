import streamlit as st
import requests
import pandas as pd
import json
import os
import re
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import subprocess, tempfile, zipfile
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes
from PIL import Image, ImageOps, ImageFilter

# --- CONFIGURACIÓN Y UTILIDADES ---
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]

def cargar_config():
    default_tamanos = {
        "<<Tipo>>": 11,
        "<<Sucursal>>": 11,
        "<<Seccion>>": 11,
        "<<Confechor>>": 11,
        "<<Concat>>": 11
    }
    config = {"plantillas": {}, "tamanos": default_tamanos}
    
    if os.path.exists("config_app.json"):
        try:
            with open("config_app.json", "r") as f:
                saved_config = json.load(f)
                # Fusionar con defaults para evitar llaves faltantes
                config["plantillas"] = saved_config.get("plantillas", {})
                config["tamanos"] = saved_config.get("tamanos", default_tamanos)
        except:
            pass
    return config

def procesar_texto_maestro(texto, campo=""):
    if not texto or str(texto).lower() == "none": return ""
    if isinstance(texto, list): return texto 
    if campo == 'Hora': return str(texto).lower().strip()
    
    t = str(texto).replace('/', ' ').strip().replace('\n', ' ').replace('\r', ' ')
    t = re.sub(r'\s+', ' ', t)
    if campo == 'Seccion': return t.upper()
    
    palabras = t.lower().split()
    if not palabras: return ""
    prep = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al']
    resultado = []
    for i, p in enumerate(palabras):
        es_inicio = (i == 0)
        despues_parentesis = (i > 0 and "(" in palabras[i-1])
        if es_inicio or despues_parentesis or (p not in prep):
            if p.startswith("("): resultado.append("(" + p[1:].capitalize())
            else: resultado.append(p.capitalize())
        else: resultado.append(p)
    return " ".join(resultado)

def crear_imagen_con_fondo_blur(img_data, target_w_pt, target_h_pt):
    target_w = max(1, int(target_w_pt / 9525))
    target_h = max(1, int(target_h_pt / 9525))
    img = Image.open(BytesIO(img_data)).convert("RGB")
    fondo = ImageOps.fit(img, (target_w, target_h), Image.Resampling.LANCZOS)
    fondo = fondo.filter(ImageFilter.GaussianBlur(radius=15))
    img.thumbnail((target_w, target_h), Image.Resampling.LANCZOS)
    offset = ((target_w - img.width) // 2, (target_h - img.height) // 2)
    fondo.paste(img, offset)
    output = BytesIO()
    fondo.save(output, format="JPEG", quality=90)
    output.seek(0)
    return output

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

# --- UI STREAMLIT ---
st.set_page_config(page_title="Provident Pro v57", layout="wide")
if 'config' not in st.session_state: 
    st.session_state.config = cargar_config()

st.title("🚀 Generador Pro v57")

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# --- SIDEBAR: CONFIGURACIÓN ---
with st.sidebar:
    st.header("⚙️ Configuración")
    
    with st.expander("📏 TAMAÑOS DE FUENTE (PTS)", expanded=True):
        # Usamos una lista de llaves fija para evitar errores de iteración si el dict cambia
        llaves_tamanos = ["<<Tipo>>", "<<Sucursal>>", "<<Seccion>>", "<<Confechor>>", "<<Concat>>"]
        for key in llaves_tamanos:
            val_actual = st.session_state.config["tamanos"].get(key, 11)
            st.session_state.config["tamanos"][key] = st.number_input(
                f"Tamaño {key}", 
                min_value=6, max_value=80, 
                value=int(val_actual),
                key=f"size_{key}"
            )

    if st.button("💾 GUARDAR CONFIG"):
        with open("config_app.json", "w") as f:
            json.dump(st.session_state.config, f)
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

# --- CUERPO PRINCIPAL ---
if 'raw_records' in st.session_state:
    modo = st.radio("Salida:", ["Postales", "Reportes"], horizontal=True)
    df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    
    df_view = df_full.copy()
    for c in df_view.columns:
        if isinstance(df_view[c].iloc[0], list): df_view.drop(c, axis=1, inplace=True)
    
    if 'select_all' not in st.session_state: st.session_state.select_all = False
    c1, c2, _ = st.columns([1,1,4])
    if c1.button("✅ Todo"): st.session_state.select_all = True; st.rerun()
    if c2.button("❌ Nada"): st.session_state.select_all = False; st.rerun()

    df_view.insert(0, "Seleccionar", st.session_state.select_all)
    df_edit = st.data_editor(df_view, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        total_items = len(sel_idx)
        folder_fisica = os.path.join("Plantillas", modo.upper())
        archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
        tipos_sel = df_view.loc[sel_idx, "Tipo"].unique()
        for t in tipos_sel:
            p_mem = st.session_state.config["plantillas"].get(t)
            idx_def = archivos_pptx.index(p_mem) if p_mem in archivos_pptx else 0
            st.session_state.config["plantillas"][t] = st.selectbox(f"Plantilla {t}:", archivos_pptx, index=idx_def, key=f"p_{t}")

        if st.button("🔥 GENERAR", use_container_width=True, type="primary"):
            p_bar = st.progress(0); zip_buf = BytesIO()
            AZUL_CELESTE = RGBColor(0, 176, 240)
            
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                for i, idx in enumerate(sel_idx):
                    record = st.session_state.raw_records[idx]['fields']
                    record_orig = st.session_state.raw_data_original[idx]['fields']
                    dt = datetime.strptime(record.get('Fecha'), '%Y-%m-%d')
                    f_tipo = record.get('Tipo'); f_suc = record.get('Sucursal')
                    
                    lugar = record.get('Punto de reunion') or record.get('Ruta a seguir')
                    f_concat = f"Sucursal {f_suc}" if f_tipo == "Actividad en Sucursal" else ", ".join([str(x) for x in [lugar, record.get('Municipio')] if x and str(x).lower() != 'none'])
                    nom_arch = f"{dt.day} de {MESES_ES[dt.month-1]} de {dt.year} - {f_tipo}, {f_suc}" + ("" if f_tipo == "Actividad en Sucursal" else f" - {f_concat}")
                    
                    reemplazos = {
                        "<<Tipo>>": f_tipo, 
                        "<<Sucursal>>": f_suc, 
                        "<<Seccion>>": record.get('Seccion'), 
                        "<<Confechor>>": f"{DIAS_ES[dt.weekday()]} {dt.day} de {MESES_ES[dt.month-1]} de {dt.year}, {record.get('Hora', '').lower()}", 
                        "<<Concat>>": f_concat
                    }

                    prs = Presentation(os.path.join(folder_fisica, st.session_state.config["plantillas"][f_tipo]))
                    
                    if f_tipo == "Actividad en Sucursal":
                        adj_lista = record_orig.get("Lista de asistencia")
                        if not adj_lista or (isinstance(adj_lista, list) and len(adj_lista) == 0):
                            if len(prs.slides) >= 4:
                                xml_slides = prs.slides._sldIdLst
                                xml_slides.remove(xml_slides[3])

                    for slide in prs.slides:
                        for shape in list(slide.shapes):
                            if shape.has_text_frame:
                                for tf in ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]:
                                    if f"<<{tf}>>" in shape.text_frame.text:
                                        adj = record_orig.get(tf)
                                        if adj:
                                            try:
                                                r_img = requests.get(adj[0].get('url')).content
                                                slide.shapes.add_picture(crear_imagen_con_fondo_blur(r_img, shape.width, shape.height), shape.left, shape.top, shape.width, shape.height)
                                                sp = shape._element; sp.getparent().remove(sp)
                                            except: pass

                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for tag, val in reemplazos.items():
                                    if tag in shape.text_frame.text:
                                        tf = shape.text_frame; tf.clear()
                                        run = tf.paragraphs[0].add_run()
                                        run.text = str(val); run.font.bold = True; run.font.color.rgb = AZUL_CELESTE
                                        run.font.size = Pt(st.session_state.config["tamanos"].get(tag, 11))

                    pp_io = BytesIO(); prs.save(pp_io)
                    data_out = generar_pdf(pp_io.getvalue())
                    if data_out:
                        ext = ".pdf" if modo == "Reportes" else ".jpg"
                        ruta_zip = f"Provident/{dt.year}/{str(dt.month).zfill(2)} - {MESES_ES[dt.month-1]}/{modo}/{f_suc}/{nom_arch[:140]}{ext}"
                        zip_f.writestr(ruta_zip, data_out if modo == "Reportes" else convert_from_bytes(data_out)[0].tobytes())
                    
                    if total_items > 0: p_bar.progress((i + 1) / total_items)
            
            st.download_button("📥 DESCARGAR", zip_buf.getvalue(), f"Provident_{datetime.now().strftime('%H%M%S')}.zip", use_container_width=True)
