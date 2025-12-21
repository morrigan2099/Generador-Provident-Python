import streamlit as st
import requests
import pandas as pd
import json
import os
import re
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import subprocess, tempfile, zipfile
from io import BytesIO
from pdf2image import convert_from_bytes

# --- FUNCIONES DE SOPORTE (AISLADAS) ---

def procesar_texto_maestro(v, c=""):
    """Función ultra-simple para evitar errores de línea"""
    if not v or str(v).lower() == "none": return ""
    if isinstance(v, list): return v
    t = str(v).strip()
    if c == 'Seccion': return t.upper()
    return t.capitalize()

def generar_pdf(pptx_data):
    """Generación de PDF segura"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_data)
        p_in = tmp.name
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(p_in), p_in], check=True)
        p_out = p_in.replace(".pptx", ".pdf")
        with open(p_out, "rb") as f: data = f.read()
        if os.path.exists(p_in): os.remove(p_in)
        if os.path.exists(p_out): os.remove(p_out)
        return data
    except: return None

# --- INICIO DE APP ---
st.set_page_config(page_title="Provident Pro v42", layout="wide")

if 'config' not in st.session_state:
    if os.path.exists("config_app.json"):
        with open("config_app.json", "r") as f: st.session_state.config = json.load(f)
    else:
        st.session_state.config = {"plantillas": {}, "columnas_visibles": []}

st.title("🚀 Generador Pro v42")

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Conexión")
    if st.button("💾 GUARDAR CONFIG"):
        with open("config_app.json", "w") as f: json.dump(st.session_state.config, f)
        st.success("Guardado")

    r_b = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_b.status_code == 200:
        bases = {b['name']: b['id'] for b in r_b.json().get('bases', [])}
        b_sel = st.selectbox("Base:", [""] + list(bases.keys()))
        if b_sel:
            r_t = requests.get(f"https://api.airtable.com/v0/meta/bases/{bases[b_sel]}/tables", headers=headers)
            tablas = {t['name']: t['id'] for t in r_t.json().get('tables', [])}
            t_sel = st.selectbox("Tabla:", list(tablas.keys()))
            if st.button("🔄 CARGAR REGISTROS"):
                r_r = requests.get(f"https://api.airtable.com/v0/{bases[b_sel]}/{tablas[t_sel]}", headers=headers)
                recs = r_r.json().get("records", [])
                st.session_state.raw_data_original = recs
                st.session_state.raw_records = [
                    {'id': r['id'], 'fields': {k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v) for k, v in r['fields'].items()}} 
                    for r in recs
                ]
                st.rerun()

# --- CUERPO PRINCIPAL ---
if 'raw_records' in st.session_state:
    df = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    
    with st.sidebar:
        st.divider()
        cols_v = st.multiselect("Columnas:", list(df.columns), default=st.session_state.config.get("columnas_visibles") or list(df.columns))
        st.session_state.config["columnas_visibles"] = cols_v

    # Filtrar columnas y quitar listas para la tabla
    df_view = df[[c for c in cols_v if c in df.columns]].copy()
    for c in df_view.columns:
        if not df_view.empty and isinstance(df_view[c].iloc[0], list): df_view.drop(c, axis=1, inplace=True)
    
    # Agregar columna de selección
    df_view.insert(0, "Seleccionar", False)
    
    # Editor de datos con Checkbox Maestro
    df_edit = st.data_editor(
        df_view, 
        use_container_width=True, 
        hide_index=True, 
        column_config={"Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False)}
    )
    
    # Obtener índices seleccionados de forma segura
    sel_bool = df_edit["Seleccionar"].tolist()
    idx_sel = [i for i, val in enumerate(sel_bool) if val is True]

    if idx_sel:
        st.write(f"✅ Registros seleccionados: {len(idx_sel)}")
        modo = st.radio("Formato:", ["Postales", "Reportes"], horizontal=True)
        folder = os.path.join("Plantillas", modo.upper())
        
        # Asignación de plantillas
        tipos_unicos = df_view.loc[idx_sel, "Tipo"].unique()
        for t in tipos_unicos:
            archs = [f for f in os.listdir(folder) if f.endswith('.pptx')]
            st.session_state.config.setdefault("plantillas", {})[t] = st.selectbox(f"Plantilla para {t}:", archs, key=f"p_{t}")

        if st.button("🔥 GENERAR", use_container_width=True, type="primary"):
            p_bar = st.progress(0); buf = BytesIO()
            with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED) as zf:
                for i, idx in enumerate(idx_sel):
                    r_p = st.session_state.raw_records[idx]['fields']
                    r_o = st.session_state.raw_data_original[idx]['fields']
                    
                    # Cargar PPTX
                    plantilla_path = os.path.join(folder, st.session_state.config["plantillas"][r_p['Tipo']])
                    prs = Presentation(plantilla_path)
                    
                    reemp = {
                        "<<Tipo>>": r_p.get('Tipo'),
                        "<<Sucursal>>": r_p.get('Sucursal'),
                        "<<Seccion>>": r_p.get('Seccion'),
                        "<<Confechor>>": f"{r_p.get('Fecha')}, {r_p.get('Hora')}",
                        "<<Concat>>": f"{r_p.get('Punto de reunion') or r_p.get('Ruta a seguir')}, {r_p.get('Municipio')}"
                    }

                    for slide in prs.slides:
                        # Fotos
                        for shp in list(slide.shapes):
                            txt = shp.text_frame.text if shp.has_text_frame else ""
                            for tag in ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]:
                                if f"<<{tag}>>" in txt or tag == shp.name:
                                    adj = r_o.get(tag)
                                    if adj:
                                        try:
                                            img_data = requests.get(adj[0]['url']).content
                                            slide.shapes.add_picture(BytesIO(img_data), shp.left, shp.top, shp.width, shp.height)
                                            shp.element.getparent().remove(shp.element)
                                        except: pass
                        # Texto
                        for shp in slide.shapes:
                            if shp.has_text_frame:
                                for tag, val in reemp.items():
                                    if tag in shp.text_frame.text:
                                        tf = shp.text_frame; tf.clear()
                                        run = tf.paragraphs[0].add_run()
                                        run.text = str(val); run.font.bold = True
                                        run.font.color.rgb = RGBColor(0, 176, 240)
                                        # APLICACIÓN DE TAMAÑOS SOLICITADOS
                                        if tag == "<<Tipo>>": run.font.size = Pt(11)
                                        elif tag == "<<Sucursal>>": run.font.size = Pt(14)
                                        else: run.font.size = Pt(11)

                    out_io = BytesIO()
                    prs.save(out_io)
                    res_pdf = generar_pdf(out_io.getvalue())
                    if res_pdf:
                        ext = ".pdf" if modo == "Reportes" else ".jpg"
                        nombre = f"{r_p.get('Fecha')} - {r_p.get('Tipo')} - {r_p.get('Sucursal')}{ext}"
                        zf.writestr(f"{modo}/{r_p.get('Sucursal')}/{nombre}", res_pdf if modo == "Reportes" else convert_from_bytes(res_pdf)[0].tobytes())
                    p_bar.progress((i + 1) / len(idx_sel))
            
            st.success("¡Completado!")
            st.download_button("📥 DESCARGAR ZIP", buf.getvalue(), "Provident_v42.zip", use_container_width=True)
