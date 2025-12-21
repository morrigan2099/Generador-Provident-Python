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
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes

# 1. FUNCIÓN DE TEXTO (AISLADA Y LIMPIA)
def procesar_texto_maestro(valor, campo=""):
    """Esta función solo transforma texto. No toca archivos."""
    if not valor or str(valor).lower() == "none":
        return ""
    if isinstance(valor, list):
        return valor
    
    texto = str(valor).strip()
    if campo == 'Seccion':
        return texto.upper()
    
    # Capitalización básica para evitar errores de lógica
    palabras = texto.lower().split()
    if not palabras: return ""
    
    prep = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al']
    res = [palabras[0].capitalize()]
    for p in palabras[1:]:
        res.append(p if p in prep else p.capitalize())
    return " ".join(res)

# 2. FUNCIÓN DE PDF (SEPARADA)
def generar_pdf(pptx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        p_pptx = tmp.name
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(p_pptx), p_pptx], check=True)
        p_pdf = p_pptx.replace(".pptx", ".pdf")
        with open(p_pdf, "rb") as f:
            pdf_d = f.read()
        if os.path.exists(p_pptx): os.remove(p_pptx)
        if os.path.exists(p_pdf): os.remove(p_pdf)
        return pdf_d
    except:
        return None

# 3. CONFIGURACIÓN INICIAL
if 'config' not in st.session_state:
    if os.path.exists("config_app.json"):
        with open("config_app.json", "r") as f: st.session_state.config = json.load(f)
    else:
        st.session_state.config = {"plantillas": {}, "columnas_visibles": []}

# --- INTERFAZ ---
st.set_page_config(page_title="Provident Pro v41", layout="wide")
st.title("🚀 Generador Pro: Estabilidad v41")

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

with st.sidebar:
    st.header("⚙️ Configuración")
    if st.button("💾 GUARDAR"):
        with open("config_app.json", "w") as f: json.dump(st.session_state.config, f)
        st.toast("Configuración guardada")

    st.divider()
    # Carga de Bases
    r_b = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_b.status_code == 200:
        bases = {b['name']: b['id'] for b in r_b.json().get('bases', [])}
        b_sel = st.selectbox("Base:", [""] + list(bases.keys()))
        if b_sel:
            r_t = requests.get(f"https://api.airtable.com/v0/meta/bases/{bases[b_sel]}/tables", headers=headers)
            tablas = {t['name']: t['id'] for t in r_t.json().get('tables', [])}
            t_sel = st.selectbox("Tabla:", list(tablas.keys()))
            if st.button("🔄 CARGAR DATOS"):
                r_r = requests.get(f"https://api.airtable.com/v0/{bases[b_sel]}/{tablas[t_sel]}", headers=headers)
                data = r_r.json().get("records", [])
                st.session_state.raw_data_original = data
                st.session_state.raw_records = [
                    {'id': r['id'], 'fields': {k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v) for k, v in r['fields'].items()}} 
                    for r in data
                ]
                st.rerun()

# --- TABLA ---
if 'raw_records' in st.session_state:
    df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    with st.sidebar:
        st.divider()
        cols_v = st.multiselect("Columnas:", list(df_full.columns), default=st.session_state.config.get("columnas_visibles") or list(df_full.columns))
        st.session_state.config["columnas_visibles"] = cols_v

    df_view = df_full[[c for c in cols_v if c in df_full.columns]].copy()
    for col in df_view.columns:
        if not df_view.empty and isinstance(df_view[col].iloc[0], list): df_view.drop(col, axis=1, inplace=True)
    
    df_view.insert(0, "Seleccionar", False)
    # Checkbox maestro activado automáticamente por el CheckboxColumn
    df_edit = st.data_editor(df_view, use_container_width=True, hide_index=True, column_config={"Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False)})
    
    idx_sel = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if idx_sel:
        modo = st.radio("Acción:", ["Postales", "Reportes"], horizontal=True)
        fol = os.path.join("Plantillas", modo.upper())
        AZUL = RGBColor(0, 176, 240)
        
        # Selección de Plantillas
        tipos = df_view.loc[idx_sel, "Tipo"].unique()
        for t in tipos:
            archs = [f for f in os.listdir(fol) if f.endswith('.pptx')]
            st.session_state.config.setdefault("plantillas", {})[t] = st.selectbox(f"Plantilla {t}:", archs, key=f"sel_{t}")

        if st.button("🔥 GENERAR", use_container_width=True, type="primary"):
            p_bar = st.progress(0); buf = BytesIO()
            with zipfile.ZipFile(buf, "a", zipfile.ZIP_DEFLATED) as zf:
                for i, idx in enumerate(idx_sel):
                    r_p = st.session_state.raw_records[idx]['fields']
                    r_o = st.session_state.raw_data_original[idx]['fields']
                    
                    prs = Presentation(os.path.join(fol, st.session_state.config["plantillas"][r_p['Tipo']]))
                    reemp = {
                        "<<Tipo>>": r_p.get('Tipo'), "<<Sucursal>>": r_p.get('Sucursal'),
                        "<<Seccion>>": r_p.get('Seccion'), "<<Confechor>>": f"{r_p.get('Fecha')}, {r_p.get('Hora')}",
                        "<<Concat>>": f"{r_p.get('Punto de reunion') or r_p.get('Ruta a seguir')}, {r_p.get('Municipio')}"
                    }

                    for slide in prs.slides:
                        # Imágenes
                        for shp in list(slide.shapes):
                            txt = shp.text_frame.text if shp.has_text_frame else ""
                            for tag in ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]:
                                if f"<<{tag}>>" in txt or tag == shp.name:
                                    adj = r_o.get(tag)
                                    if adj:
                                        try:
                                            img = requests.get(adj[0]['url']).content
                                            slide.shapes.add_picture(BytesIO(img), shp.left, shp.top, shp.width, shp.height)
                                            shp.element.getparent().remove(shp.element)
                                        except: pass
                        # Texto
                        for shp in slide.shapes:
                            if shp.has_text_frame:
                                for tag, val in reemp.items():
                                    if tag in shp.text_frame.text:
                                        tf = shp.text_frame; tf.clear()
                                        run = tf.paragraphs[0].add_run()
                                        run.text = str(val); run.font.bold = True; run.font.color.rgb = AZUL
                                        # TAMAÑO FIJO TIPO 11
                                        run.font.size = Pt(14) if tag == "<<Sucursal>>" else Pt(11)

                    out = BytesIO(); prs.save(out)
                    pdf = generar_pdf(out.getvalue())
                    if pdf:
                        ext = ".pdf" if modo == "Reportes" else ".jpg"
                        name = f"{r_p.get('Fecha')} - {r_p.get('Tipo')} - {r_p.get('Sucursal')}{ext}"
                        zf.writestr(f"{modo}/{r_p.get('Sucursal')}/{name}", pdf if modo == "Reportes" else convert_from_bytes(pdf)[0].tobytes())
                    p_bar.progress((i + 1) / len(idx_sel))
            
            st.success("¡Listo!")
            st.download_button("📥 DESCARGAR", buf.getvalue(), "Provident.zip", use_container_width=True)
