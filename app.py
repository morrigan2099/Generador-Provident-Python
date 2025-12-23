import streamlit as st
import requests
import pandas as pd
import json
import os
import re
import numpy as np
import textwrap
import calendar
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
import subprocess, tempfile, zipfile
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes
from PIL import Image, ImageOps, ImageFilter, ImageChops
from collections import Counter

# ============================================================
#  CONFIGURACIÓN GLOBAL
# ============================================================
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]

# ============================================================
#  FUNCIONES TÉCNICAS
# ============================================================

def recorte_inteligente_bordes(img, umbral_negro=60):
    img_gray = img.convert("L")
    arr = np.array(img_gray)
    h, w = arr.shape
    def fila_es_negra(fila): return (np.sum(fila < 35) / fila.size) * 100 > umbral_negro
    def columna_es_negra(col): return (np.sum(col < 35) / col.size) * 100 > umbral_negro
    top = 0
    while top < h and fila_es_negra(arr[top, :]): top += 1
    bottom = h - 1
    while bottom > top and fila_es_negra(arr[bottom, :]): bottom -= 1
    left = 0
    while left < w and columna_es_negra(arr[:, left]): left += 1
    right = w - 1
    while right > left and columna_es_negra(arr[:, right]): right -= 1
    if right <= left or bottom <= top: return img
    return img.crop((left, top, right + 1, bottom + 1))

def procesar_imagen_inteligente(img_data, target_w_pt, target_h_pt, con_blur=False):
    base_w, base_h = int(target_w_pt / 9525), int(target_h_pt / 9525)
    render_w, render_h = base_w * 2, base_h * 2
    img = Image.open(BytesIO(img_data)).convert("RGB")
    img = recorte_inteligente_bordes(img, umbral_negro=60)
    if con_blur:
        fondo = ImageOps.fit(img, (render_w, render_h), Image.Resampling.LANCZOS)
        fondo = fondo.filter(ImageFilter.GaussianBlur(radius=10))
        img.thumbnail((render_w, render_h), Image.Resampling.LANCZOS)
        offset = ((render_w - img.width) // 2, (render_h - img.height) // 2)
        fondo.paste(img, offset)
        img_final = fondo
    else:
        img_final = img.resize((render_w, render_h), Image.Resampling.LANCZOS)
    output = BytesIO()
    img_final.save(output, format="JPEG", quality=90, subsampling=0, optimize=True)
    output.seek(0)
    return output

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
        else:
            resultado.append(p)
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

# --- LÓGICA DE DATOS ---
def obtener_fecha_texto(fecha_dt):
    dia_idx = fecha_dt.weekday()
    return f"{DIAS_ES[dia_idx]} {fecha_dt.day} de {MESES_ES[fecha_dt.month - 1]} de {fecha_dt.year}"

def obtener_hora_texto(hora_str):
    if not hora_str or str(hora_str).lower() == "none": return ""
    s_raw = str(hora_str).lower().strip()
    es_pm = "p.m." in s_raw or "pm" in s_raw or "p. m." in s_raw
    es_am = "a.m." in s_raw or "am" in s_raw or "a. m." in s_raw
    match = re.search(r'(\d{1,2}):(\d{2})', s_raw)
    if match:
        h, m = int(match.group(1)), match.group(2)
        if es_pm and h < 12: h += 12
        if es_am and h == 12: h = 0
        if h == 0 and es_pm: h = 12 
        if 8 <= h <= 11: suf = "de la mañana"
        elif h == 12: suf = "del día"
        elif 13 <= h <= 19: suf = "de la tarde"
        elif h >= 20: suf = "p.m."
        else: suf = "a.m."
        h_show = h - 12 if h > 12 else (12 if h == 0 else h)
        return f"{h_show}:{m} {suf}"
    return hora_str

def obtener_concat_texto(record):
    parts = []
    v_pt = record.get('Punto de reunion')
    if val_punto and str(val_punto).lower() != 'none': parts.append(str(val_punto))
    v_rt = record.get('Ruta a seguir')
    if val_ruta and str(val_ruta).lower() != 'none': parts.append(str(val_ruta))
    v_mun = record.get('Municipio')
    if val_mun and str(val_mun).lower() != 'none': parts.append(f"Municipio {val_mun}")
    v_sec = record.get('Seccion')
    if val_sec and str(val_sec).lower() != 'none': parts.append(f"Sección {str(val_sec).upper()}")
    return ", ".join(parts)

# ============================================================
#  INICIO DE LA APP
# ============================================================

st.set_page_config(page_title="Provident Pro v113", layout="wide")

# 1. BLOQUEO TECLADO (JS)
st.markdown("""
<script>
    const observer = new MutationObserver((mutations) => {
        const inputs = window.parent.document.querySelectorAll('input[type="text"]');
        inputs.forEach(input => {
            if (input.getAttribute('aria-autocomplete') === 'list') {
                input.setAttribute('inputmode', 'none');
                input.setAttribute('readonly', 'true');
            }
        });
    });
    observer.observe(window.parent.document.body, { childList: true, subtree: true });
</script>
""", unsafe_allow_html=True)

# 2. ESTILOS CSS AGRESIVOS (ANTI MODO OSCURO + COLORES FIJOS)
st.markdown("""
<style>
    /* 1. FORZAR MODO CLARO EN EL NAVEGADOR */
    :root {
        color-scheme: light;
    }
    
    /* 2. CONTENEDORES PRINCIPALES */
    [data-testid="stAppViewContainer"], [data-testid="stSidebar"], .stApp {
        background-color: #ffffff !important;
        color: #000000 !important;
    }
    
    /* 3. TEXTOS GENERALES (Asegurar negro sobre blanco) */
    p, label, span, div, h1, h2, h3, h4, h5, h6, li {
        color: #000000 !important;
    }
    
    /* 4. SIDEBAR ESPECÍFICO */
    [data-testid="stSidebar"] {
        background-color: #f8f9fa !important;
        border-right: 1px solid #ddd;
    }
    
    /* 5. BOTONES (Azul Oscuro / Blanco) */
    div.stButton > button {
        background-color: #002060 !important;
        color: #ffffff !important;
        border: none;
        border-radius: 4px;
        font-weight: bold;
    }
    div.stButton > button:hover {
        background-color: #00b0f0 !important; /* Celeste al pasar mouse */
        color: #ffffff !important;
    }
    div.stButton > button p {
        color: #ffffff !important; /* Texto interno del botón blanco */
    }
    
    /* 6. RADIO BUTTONS (Eliminar Naranja, Poner Celeste) */
    /* Texto de las opciones */
    div[role="radiogroup"] label p {
        color: #000000 !important;
    }
    /* Círculo seleccionado */
    div[role="radiogroup"] div[data-checked="true"] {
        background-color: #00b0f0 !important; /* Relleno Celeste */
        border-color: #00b0f0 !important;
    }
    /* Círculo externo del seleccionado (a veces Streamlit usa dos divs) */
    div[role="radiogroup"] div[data-checked="true"] > div {
        background-color: #00b0f0 !important;
    }
    
    /* 7. EXPANDERS */
    .streamlit-expanderHeader {
        background-color: #ffffff !important;
        color: #002060 !important; /* Título azul oscuro */
        font-weight: bold;
    }
    .streamlit-expanderHeader p {
        color: #002060 !important;
    }
    
    /* 8. TABLAS (Headers Celestes) */
    [data-testid="stDataFrameResizable"] th {
        background-color: #00b0f0 !important;
        color: #ffffff !important;
    }
    [data-testid="stDataFrameResizable"] th div {
        color: #ffffff !important;
    }
    
    /* 9. EXCEPCIONES PARA EL CALENDARIO (Para que el CSS interno funcione) */
    .c-head, .c-day, .c-foot, .c-foot-empty {
        color: #ffffff !important; /* Texto blanco forzado en componentes oscuros */
    }
    
    /* ESTILOS DEL CALENDARIO (Pegados aquí para referencia global) */
    .cal-title { text-align: center; font-size: 1.5em; font-weight: bold; margin: 0 !important; padding-bottom: 10px; color: #333 !important; background-color: #fff; }
    .c-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 2px; margin-top:0px !important; }
    .c-head { background: #002060 !important; color: white !important; padding: 4px; text-align: center; font-weight: bold; border-radius: 2px; font-size: 14px; }
    .c-cell { background: white !important; border: 1px solid #ccc; border-radius: 2px; height: 160px; display: flex; flex-direction: column; justify-content: space-between; overflow: hidden; }
    .c-day { flex: 0 0 auto; background: #00b0f0 !important; color: white !important; font-weight: 900; font-size: 1.1em; text-align: center; padding: 2px 0; }
    .c-body { flex-grow: 1; width: 100%; background-position: center; background-repeat: no-repeat; background-size: cover; background-color: #f8f8f8 !important; }
    .c-foot { flex: 0 0 auto; height: 20px; background: #002060 !important; color: #ffffff !important; font-weight: 900; text-align: center; font-size: 0.9em; padding: 1px; white-space: nowrap; overflow: hidden; }
    .c-foot-empty { flex: 0 0 auto; height: 20px; background: #e0e0e0 !important; }
    
    @media (max-width: 600px) {
        .c-cell { height: 110px; }
        .c-day { font-size: 0.9em; }
        .c-foot, .c-foot-empty { font-size: 0.7em; height: 16px; }
    }
</style>
""", unsafe_allow_html=True)

if 'config' not in st.session_state:
    if os.path.exists("config_app.json"):
        with open("config_app.json", "r") as f: st.session_state.config = json.load(f)
    else: st.session_state.config = {"plantillas": {}}

st.title("🚀 Generador Pro v113")
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# --- SIDEBAR ---
with st.sidebar:
    st.header("🔗 Conexión")
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        
        with st.expander("📂 Seleccionar Base", expanded=True):
            base_sel = st.radio("Bases disponibles:", list(base_opts.keys()), label_visibility="collapsed")
        
        if base_sel:
            st.session_state['base_activa_id'] = base_opts[base_sel]
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
            if r_tab.status_code == 200:
                tablas_data = r_tab.json()['tables']
                st.session_state['todas_tablas'] = {t['name']: t['id'] for t in tablas_data}
                
                with st.expander("📅 Seleccionar Tabla (Mes)", expanded=True):
                    tabla_sel = st.radio("Tablas disponibles:", list(st.session_state['todas_tablas'].keys()), label_visibility="collapsed")
                
                if st.button("🔄 CARGAR DATOS", type="primary"):
                    with st.spinner("Conectando..."):
                        r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{st.session_state['todas_tablas'][tabla_sel]}", headers=headers)
                        recs = r_reg.json().get("records", [])
                        st.session_state.raw_data_original = recs
                        st.session_state.raw_records = [
                            {'id': r['id'], 'fields': {k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v) for k, v in r['fields'].items()}}
                            for r in recs
                        ]
                        st.session_state['tabla_actual_nombre'] = tabla_sel
                    st.success(f"Cargados {len(recs)} registros")
                    st.rerun()
    st.divider()
    
    modulo = None
    if 'raw_records' in st.session_state:
        st.header("2. Módulos")
        modulo = st.radio("Ir a:", ["📮 Postales", "📄 Reportes", "📅 Calendario"], index=2)
        if st.button("💾 Guardar Config"):
            with open("config_app.json", "w") as f: json.dump(st.session_state.config, f)
            st.toast("Guardado")

# --- MAIN ---
if 'raw_records' not in st.session_state:
    st.info("👈 Conecta una base.")
else:
    df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    
    # --------------------------------------------------------
    # MÓDULO POSTALES
    # --------------------------------------------------------
    if modulo == "📮 Postales":
        st.subheader("📮 Generador de Postales")
        
        df_view = df_full.copy()
        for c in df_view.columns:
            if isinstance(df_view[c].iloc[0], list): df_view.drop(c, axis=1, inplace=True)
        
        if 'sa_p' not in st.session_state: st.session_state.sa_p = False
        c1,c2,_=st.columns([1,1,5])
        if c1.button("✅ Todo"): st.session_state.sa_p=True; st.rerun()
        if c2.button("❌ Nada"): st.session_state.sa_p=False; st.rerun()
        df_view.insert(0,"Seleccionar",st.session_state.sa_p)
        
        df_edit = st.data_editor(df_view, hide_index=True)
        sel_idx = df_edit.index[df_edit["Seleccionar"]==True].tolist()
        
        if sel_idx:
            folder = os.path.join("Plantillas", "POSTALES")
            if not os.path.exists(folder): os.makedirs(folder)
            archs = [f for f in os.listdir(folder) if f.endswith('.pptx')]
            tipos = df_view.loc[sel_idx, "Tipo"].unique()
            cols = st.columns(len(tipos)) if len(tipos)>0 else [st]
            for i, t in enumerate(tipos):
                mem = st.session_state.config["plantillas"].get(t)
                idx = archs.index(mem) if mem in archs else 0
                st.session_state.config["plantillas"][t] = cols[i].selectbox(f"Plantilla '{t}':", archs, index=idx, key=f"pp_{t}")

            if st.button("🔥 GENERAR POSTALES", type="primary"):
                p_bar = st.progress(0); st.session_state.archivos_en_memoria = []
                TAM_MAPA = {"<<Tipo>>":12, "<<Sucursal>>":44, "<<Seccion>>":12, "<<Conhora>>":32, "<<Concat>>":32, "<<Consuc>>":32, "<<Confechor>>":28}
                
                for i, idx in enumerate(sel_idx):
                    rec = st.session_state.raw_records[idx]['fields']
                    orig = st.session_state.raw_data_original[idx]['fields']
                    try: dt = datetime.strptime(rec.get('Fecha','2025-01-01'), '%Y-%m-%d'); nm = MESES_ES[dt.month - 1]
                    except: dt = datetime.now(); nm = "error"
                    
                    ft = rec.get('Tipo', 'Sin Tipo'); fs = rec.get('Sucursal', '000')
                    tfe = obtener_fecha_texto(dt); tho = obtener_hora_texto(rec.get('Hora',''))
                    fcf = f"{tfe.strip()}\n{tho.strip()}"
                    fcc = f"Sucursal {fs}" if ft == "Actividad en Sucursal" else obtener_concat_texto(rec)
                    ftag = f"Sucursal {fs}" if ft == "Actividad en Sucursal" else fcc
                    narc = re.sub(r'[\\/*?:"<>|]', "", f"{dt.day} de {nm} de {dt.year} - {ft}, {fs} - {ftag}")[:120] + ".png"
                    reps = {"<<Tipo>>":textwrap.fill(ft,width=35), "<<Sucursal>>":fs, "<<Seccion>>":rec.get('Seccion'), "<<Confechor>>":fcf, "<<Concat>>":fcc, "<<Consuc>>":fcc}
                    
                    try: prs = Presentation(os.path.join(folder, st.session_state.config["plantillas"][ft]))
                    except: continue
                    if ft == "Actividad en Sucursal" and not orig.get("Lista de asistencia") and len(prs.slides)>=4: prs.slides._sldIdLst.remove(prs.slides._sldIdLst[3])

                    for slide in prs.slides:
                        for shp in slide.shapes:
                            if shp.has_text_frame:
                                for tag in ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]:
                                    if f"<<{tag}>>" in shp.text_frame.text:
                                        adj = orig.get(tag)
                                        if adj:
                                            try:
                                                io = procesar_imagen_inteligente(requests.get(adj[0]['url']).content, shp.width, shp.height, con_blur=True)
                                                slide.shapes.add_picture(io, shp.left, shp.top, shp.width, shp.height)
                                                shp._element.getparent().remove(shp._element)
                                            except: pass
                        for shp in slide.shapes:
                            if shp.has_text_frame:
                                for tag, val in reps.items():
                                    if tag in shp.text_frame.text:
                                        if tag in ["<<Tipo>>", "<<Sucursal>>"]: shp.top -= Pt(2)
                                        tf = shp.text_frame; tf.word_wrap = True; bp = tf._element.bodyPr
                                        for c in ['spAutoFit', 'normAutofit', 'noAutofit']:
                                            if bp.find(qn(f'a:{c}')) is not None: bp.remove(bp.find(qn(f'a:{c}')))
                                        if tag=="<<Tipo>>": bp.append(tf._element.makeelement(qn('a:spAutoFit')))
                                        else: bp.append(tf._element.makeelement(qn('a:normAutofit')))
                                        tf.clear(); p = tf.paragraphs[0]
                                        if tag in ["<<Confechor>>", "<<Consuc>>"]: p.alignment = PP_ALIGN.CENTER
                                        p.space_before=Pt(0); p.space_after=Pt(0); p.line_spacing=1.0
                                        run = p.add_run(); run.text=str(val); run.font.bold=True; run.font.color.rgb=AZUL
                                        run.font.size=Pt(TAM_MAPA.get(tag,12))
                    
                    pp_io = BytesIO(); prs.save(pp_io)
                    dout = generar_pdf(pp_io.getvalue())
                    if dout:
                        imgs = convert_from_bytes(dout, dpi=170, fmt='jpeg')
                        with BytesIO() as b: imgs[0].save(b, format="JPEG", quality=85, optimize=True, progressive=True); fbytes = b.getvalue()
                        path = f"{dt.year}/{str(dt.month).zfill(2)} - {nm}/Postales/{fs}/{narc}"
                        st.session_state.archivos_en_memoria.append({"Seleccionar":True, "Archivo":narc, "RutaZip":path, "Datos":fbytes, "Sucursal":fs, "Tipo":f_tipo})
                    p_bar.progress((i+1)/len(sel_idx))
                st.success("Hecho")

    # --------------------------------------------------------
    # MÓDULO REPORTES
    # --------------------------------------------------------
    elif modulo == "📄 Reportes":
        st.subheader("📄 Generador de Reportes")
        df_view = df_full.copy()
        for c in df_view.columns:
            if isinstance(df_view[c].iloc[0], list): df_view.drop(c, axis=1, inplace=True)
        
        if 'sa_r' not in st.session_state: st.session_state.sa_r = False
        c1,c2,_=st.columns([1,1,5])
        if c1.button("✅ Todo", key="r1"): st.session_state.sa_r=True; st.rerun()
        if c2.button("❌ Nada", key="r2"): st.session_state.sa_r=False; st.rerun()
        df_view.insert(0,"Seleccionar",st.session_state.sa_r)
        
        df_edit = st.data_editor(df_view, hide_index=True, key="ed_r")
        sel_idx = df_edit.index[df_edit["Seleccionar"]==True].tolist()
        
        if sel_idx:
            folder = os.path.join("Plantillas", "REPORTES")
            if not os.path.exists(folder): os.makedirs(folder)
            archs = [f for f in os.listdir(folder) if f.endswith('.pptx')]
            tipos = df_view.loc[sel_idx, "Tipo"].unique()
            cols = st.columns(len(tipos)) if len(tipos)>0 else [st]
            for i, t in enumerate(tipos):
                mem = st.session_state.config["plantillas"].get(t)
                idx = archs.index(mem) if mem in archs else 0
                st.session_state.config["plantillas"][t] = cols[i].selectbox(f"Plantilla '{t}':", archs, index=idx, key=f"pr_{t}")

            if st.button("🔥 CREAR REPORTES", type="primary"):
                p_bar = st.progress(0); st.session_state.archivos_en_memoria = []
                TAM_MAPA = {"<<Tipo>>":12, "<<Sucursal>>":12, "<<Seccion>>":12, "<<Confecha>>":24, "<<Conhora>>":15, "<<Consuc>>":24}
                
                for i, idx in enumerate(sel_idx):
                    rec = st.session_state.raw_records[idx]['fields']
                    orig = st.session_state.raw_data_original[idx]['fields']
                    try: dt = datetime.strptime(rec.get('Fecha','2025-01-01'), '%Y-%m-%d'); nm = MESES_ES[dt.month - 1]
                    except: dt = datetime.now(); nm = "error"
                    
                    ft = rec.get('Tipo', 'Sin Tipo'); fs = rec.get('Sucursal', '000')
                    tfe = obtener_fecha_texto(dt); tho = obtener_hora_texto(rec.get('Hora',''))
                    fcs = f"Sucursal {fs}" if ft == "Actividad en Sucursal" else ""
                    ftag = fcs if ft == "Actividad en Sucursal" else obtener_concat_texto(rec)
                    narc = re.sub(r'[\\/*?:"<>|]', "", f"{dt.day} de {nm} de {dt.year} - {ft}, {fs} - {ftag}")[:120] + ".pdf"
                    
                    reps = {"<<Tipo>>":textwrap.fill(ft,width=35), "<<Sucursal>>":fs, "<<Seccion>>":rec.get('Seccion'), "<<Confecha>>":tfe, "<<Conhora>>":tho, "<<Consuc>>":fcs}
                    
                    try: prs = Presentation(os.path.join(folder, st.session_state.config["plantillas"][ft]))
                    except: continue
                    if ft == "Actividad en Sucursal" and not orig.get("Lista de asistencia") and len(prs.slides)>=4: prs.slides._sldIdLst.remove(prs.slides._sldIdLst[3])

                    for slide in prs.slides:
                        for shp in slide.shapes:
                            if shp.has_text_frame:
                                for tag in ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]:
                                    if f"<<{tag}>>" in shp.text_frame.text:
                                        adj = orig.get(tag)
                                        if adj:
                                            try:
                                                io = procesar_imagen_inteligente(requests.get(adj[0]['url']).content, shp.width, shp.height, con_blur=True)
                                                slide.shapes.add_picture(io, shp.left, shp.top, shp.width, shp.height)
                                                shp._element.getparent().remove(shp._element)
                                            except: pass
                        for shp in slide.shapes:
                            if shp.has_text_frame:
                                for tag, val in reps.items():
                                    if tag in shp.text_frame.text:
                                        if tag in ["<<Tipo>>", "<<Sucursal>>"]: shp.top -= Pt(2)
                                        tf = shp.text_frame; tf.word_wrap = True; bp = tf._element.bodyPr
                                        for c in ['spAutoFit', 'normAutofit', 'noAutofit']:
                                            if bp.find(qn(f'a:{c}')) is not None: bp.remove(bp.find(qn(f'a:{c}')))
                                        if tag=="<<Tipo>>": bp.append(tf._element.makeelement(qn('a:spAutoFit')))
                                        elif tag=="<<Conhora>>": bp.append(tf._element.makeelement(qn('a:spAutoFit')))
                                        else: bp.append(tf._element.makeelement(qn('a:normAutofit')))
                                        tf.clear(); p = tf.paragraphs[0]
                                        if tag in ["<<Confecha>>", "<<Conhora>>", "<<Consuc>>"]: p.alignment = PP_ALIGN.CENTER
                                        p.space_before=Pt(0); p.space_after=Pt(0); p.line_spacing=1.0
                                        run = p.add_run(); run.text=str(val); run.font.bold=True; run.font.color.rgb=AZUL
                                        run.font.size=Pt(TAM_MAPA.get(tag,12))
                    
                    pp_io = BytesIO(); prs.save(pp_io)
                    dout = generar_pdf(pp_io.getvalue())
                    if dout:
                        path = f"{dt.year}/{str(dt.month).zfill(2)} - {nm}/Reportes/{fs}/{narc}"
                        st.session_state.archivos_en_memoria.append({"Seleccionar":True, "Archivo":narc, "RutaZip":path, "Datos":dout, "Sucursal":fs, "Tipo":f_tipo})
                    p_bar.progress((i+1)/len(sel_idx))
                st.success("Hecho")

    # --------------------------------------------------------
    # MÓDULO CALENDARIO
    # --------------------------------------------------------
    elif modulo == "📅 Calendario":
        st.subheader("📅 Calendario de Actividades")
        
        if 'todas_tablas' in st.session_state:
            nombres_tablas = list(st.session_state['todas_tablas'].keys())
            idx_actual = 0
            if 'tabla_actual_nombre' in st.session_state and st.session_state['tabla_actual_nombre'] in nombres_tablas:
                idx_actual = nombres_tablas.index(st.session_state['tabla_actual_nombre'])
            
            with st.expander("📅 Cambiar Mes (Tabla)", expanded=False):
                nueva_tabla = st.radio("Meses:", nombres_tablas, index=idx_actual, horizontal=True)
            
            if nueva_tabla != st.session_state.get('tabla_actual_nombre'):
                with st.spinner(f"Cargando {nueva_tabla}..."):
                    base_id = st.session_state['base_activa_id']
                    table_id = st.session_state['todas_tablas'][nueva_tabla]
                    r_reg = requests.get(f"https://api.airtable.com/v0/{base_id}/{table_id}", headers=headers)
                    recs = r_reg.json().get("records", [])
                    st.session_state.raw_data_original = recs
                    st.session_state.raw_records = [
                        {'id': r['id'], 'fields': {k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v) for k, v in r['fields'].items()}}
                        for r in recs
                    ]
                    st.session_state['tabla_actual_nombre'] = nueva_tabla
                st.rerun()

        st.divider()

        fechas_oc = {}
        fechas_lista = []
        for r in st.session_state.raw_data_original:
            f = r['fields'].get('Fecha')
            if f:
                f_short = f.split('T')[0]
                if f_short not in fechas_oc: fechas_oc[f_short] = []
                th = None
                if 'Postal' in r['fields']:
                    att = r['fields']['Postal']
                    if isinstance(att, list) and len(att)>0: 
                        th = att[0].get('url') 
                fechas_oc[f_short].append({"id":r['id'], "thumb":th})
                fechas_lista.append(f_short)

        if not fechas_oc:
            st.warning("No hay fechas en esta tabla.")
        else:
            dt_objs = [datetime.strptime(x, '%Y-%m-%d') for x in fechas_lista]
            mc = Counter([(d.year, d.month) for d in dt_objs])
            ay, am = mc.most_common(1)[0][0]
            
            cal = calendar.Calendar(firstweekday=0) 
            weeks = cal.monthdayscalendar(ay, am)
            
            h = f"<div class='cal-title'>📅 {MESES_ES[am-1].capitalize()} {ay}</div>"
            h += "<div class='c-grid'>"
            for d in ["LUN","MAR","MIÉ","JUE","VIE","SÁB","DOM"]: h += f"<div class='c-head'>{d}</div>"
            
            for wk in weeks:
                for d in wk:
                    if d == 0: h += "<div class='c-cell' style='border:none; background:transparent;'></div>"
                    else:
                        k = f"{ay}-{str(am).zfill(2)}-{str(d).zfill(2)}"
                        acts = fechas_oc.get(k, [])
                        
                        h += f"<div class='c-cell'>"
                        
                        # HEADER
                        h += f"<div class='c-day'>{d}</div>"
                        
                        # BODY
                        style_bg = ""
                        if acts and acts[0]['thumb']:
                            style_bg = f"style=\"background-image: url('{acts[0]['thumb']}');\""
                        h += f"<div class='c-body' {style_bg}></div>"
                        
                        # FOOTER LOGIC
                        if len(acts) > 1:
                            h += f"<div class='c-foot'>+ {len(acts)-1} más</div>"
                        elif len(acts) == 1:
                            h += f"<div class='c-foot'>&nbsp;</div>" # Azul sin texto
                        else:
                            h += f"<div class='c-foot-empty'></div>" # Gris vacío
                        
                        h += "</div>"
            h += "</div>"
            st.markdown(h, unsafe_allow_html=True)

    # --- DESCARGAS ---
    if modulo in ["📮 Postales", "📄 Reportes"] and "archivos_en_memoria" in st.session_state and len(st.session_state.archivos_en_memoria)>0:
        st.divider()
        c1,c2,_=st.columns([1,1,3])
        if c1.button("☑️ Todo", key="dt"): 
            for i in st.session_state.archivos_en_memoria: i["Seleccionar"]=True
            st.rerun()
        if c2.button("⬜ Nada", key="dn"): 
            for i in st.session_state.archivos_en_memoria: i["Seleccionar"]=False
            st.rerun()
        
        df_d = pd.DataFrame(st.session_state.archivos_en_memoria)
        ed = st.data_editor(df_d[["Seleccionar", "Archivo", "Sucursal"]], hide_index=True)
        idxs = ed[ed["Seleccionar"]==True].index.tolist()
        fins = [st.session_state.archivos_en_memoria[i] for i in idxs]
        
        if len(fins)>0:
            zb = BytesIO()
            with zipfile.ZipFile(zb, "a", zipfile.ZIP_DEFLATED) as z:
                for f in fins: z.writestr(f["RutaZip"], f["Datos"])
            st.download_button(f"⬇️ DESCARGAR {len(fins)}", zb.getvalue(), f"Pack_{datetime.now().strftime('%H%M%S')}.zip", "application/zip", type="primary")
