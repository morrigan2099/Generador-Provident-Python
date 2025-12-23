import streamlit as st
import requests
import pandas as pd
import json
import os
import re
import numpy as np
import textwrap
import calendar
import urllib.parse
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
import subprocess, tempfile, zipfile
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes
from PIL import Image, ImageOps, ImageFilter
from collections import Counter

# ============================================================
#  CONFIGURACIÓN GLOBAL Y WHATSAPP
# ============================================================
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]

WHATSAPP_GROUPS = {
    "cordoba": {"link": "https://chat.whatsapp.com/EablxD3sq3n65GHWq9uKTg", "name": "CORDOBA"},
    "boca del rio": {"link": "https://chat.whatsapp.com/ES2ufZaP8f3JrNyO9lpFcX", "name": "BOCA DEL RIO"},
    "orizaba": {"link": "https://chat.whatsapp.com/EHITAvFTeYO5hOsS14xXJM", "name": "ORIZABA"},
    "tuxtepec": {"link": "https://chat.whatsapp.com/HkKsqVFYZSn99FPdjZ7Whv", "name": "TUXTEPEC"},
    "oaxaca": {"link": "https://chat.whatsapp.com/JRawICryDEf8eO2RYKqS0T", "name": "OAXACA"},
    "tehuacan": {"link": "https://chat.whatsapp.com/E9z0vLSwfZCB97Ou4A6Chv", "name": "TEHUACAN"},
    "xalapa": {"link": "", "name": "XALAPA"},
}

# ============================================================
#  FUNCIONES TÉCNICAS
# ============================================================

def recorte_inteligente_bordes(img, umbral_negro=60):
    img_gray = img.convert("L")
    arr = np.array(img_gray)
    h, w = arr.shape
    top, bottom = 0, h - 1
    def fila_es_negra(fila): return (np.sum(fila < 35) / fila.size) * 100 > umbral_negro
    while top < h and fila_es_negra(arr[top, :]): top += 1
    while bottom > top and fila_es_negra(arr[bottom, :]): bottom -= 1
    left, right = 0, w - 1
    def col_es_negra(c): return (np.sum(arr[:, c] < 35) / h) * 100 > umbral_negro
    while left < w and col_es_negra(left): left += 1
    while right > left and col_es_negra(right): right -= 1
    return img.crop((left, top, right + 1, bottom + 1))

def procesar_imagen_inteligente(img_data, target_w_pt, target_h_pt, con_blur=False):
    base_w, base_h = int(target_w_pt / 9525), int(target_h_pt / 9525)
    render_w, render_h = base_w * 2, base_h * 2
    img = Image.open(BytesIO(img_data)).convert("RGB")
    img = recorte_inteligente_bordes(img)
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
    img_final.save(output, format="JPEG", quality=85, optimize=True)
    output.seek(0)
    return output

def procesar_texto_maestro(texto, campo=""):
    if not texto or str(texto).lower() == "none": return ""
    t = str(texto).replace('/', ' ').strip().replace('\n', ' ').replace('\r', ' ')
    if campo == 'Seccion': return t.upper()
    palabras = t.lower().split()
    if not palabras: return ""
    prep = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al']
    resultado = []
    for i, p in enumerate(palabras):
        if i == 0 or "(" in (palabras[i-1] if i>0 else "") or p not in prep:
            resultado.append(p.capitalize())
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

def obtener_hora_texto(hora_str):
    if not hora_str or str(hora_str).lower() == "none": return ""
    s_raw = str(hora_str).lower().strip()
    match = re.search(r'(\d{1,2}):(\d{2})', s_raw)
    if match:
        h, m = int(match.group(1)), match.group(2)
        es_pm = any(x in s_raw for x in ["p.m.", "pm", "p. m."])
        if es_pm and h < 12: h += 12
        h_show = h - 12 if h > 12 else (12 if h == 0 else h)
        suf = "a.m." if h < 12 else "p.m."
        return f"{h_show}:{m} {suf}"
    return hora_str

def obtener_concat_texto(record):
    parts = [record.get(f) for f in ['Punto de reunion', 'Ruta a seguir', 'Municipio'] if record.get(f)]
    return ", ".join([str(p) for p in parts if str(p).lower() != 'none'])

# ============================================================
#  INICIO APP
# ============================================================
st.set_page_config(page_title="Provident Pro v141", layout="wide")

if 'active_module' not in st.session_state: st.session_state.active_module = st.query_params.get("view", "Calendario")
if 'dia_seleccionado' not in st.session_state: st.session_state.dia_seleccionado = None
if 'idx_postal' not in st.session_state: st.session_state.idx_postal = 0

# --- CSS PARA ELIMINAR ESPACIOS Y FORZAR FILA ---
st.markdown("""
<style>
    /* ELIMINAR ESPACIOS DE STREAMLIT */
    .block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; }
    div[data-testid="stVerticalBlock"] > div { padding: 0px !important; margin: 0px !important; }
    
    /* TABLA NAV COMPACTA */
    .table-nav { width: 100%; border-collapse: collapse; margin-bottom: 5px; background: white; }
    .table-nav td { padding: 5px; vertical-align: middle; text-align: center; border: none; }
    .nav-title { line-height: 1.1; }
    .nav-title b { color: #002060; font-size: 1.1em; display: block; }
    .nav-title span { color: #333; font-size: 0.9em; display: block; }
    
    /* BOTONES FLECHA HTML */
    .btn-arrow { 
        background: #00b0f0; color: white !important; border-radius: 50%; width: 35px; height: 35px; 
        display: flex; align-items: center; justify-content: center; text-decoration: none !important;
        font-weight: bold; font-size: 1.2em; margin: 0 auto;
    }

    /* BOTÓN VOLVER CELESTE */
    div.stButton > button[key="btn_volver"] { 
        background-color: #00b0f0 !important; color: white !important; font-weight: bold !important; 
        width: 100%; border: none !important; margin-top: 5px !important;
    }

    /* CALENDARIO */
    [data-testid="column"] { min-width: 0px !important; flex: 1 1 0% !important; }
    .c-head { background: #002060; color: white; padding: 4px; text-align: center; font-weight: bold; font-size: 10px; }
    .c-cell-container { position: relative; height: 110px; border: 1px solid #ccc; background: white; }
    .c-day { background: #00b0f0; color: white; font-weight: bold; font-size: 0.9em; text-align: center; }
    .c-body { flex-grow: 1; background-size: cover; background-position: center; }
    .c-foot { height: 16px; background: #002060; color: white; text-align: center; font-size: 0.7em; }
    .stButton > button[key^="day_"] { position: absolute; top: 0; left: 0; width: 100%; height: 100%; background: transparent !important; border: none !important; color: transparent !important; z-index: 10; }
</style>
""", unsafe_allow_html=True)

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# SIDEBAR
with st.sidebar:
    st.header("🔗 Conexión")
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.radio("Base:", list(base_opts.keys()))
        if base_sel:
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
            if r_tab.status_code == 200:
                tab_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
                tabla_sel = st.radio("Mes:", list(tab_opts.keys()))
                if st.session_state.get('tabla_actual') != tabla_sel:
                    r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tab_opts[tabla_sel]}", headers=headers)
                    recs = r_reg.json().get("records", [])
                    st.session_state.raw_data_original = recs
                    st.session_state.raw_records = [{'id': r['id'], 'fields': {k: procesar_texto_maestro(v, k) for k, v in r['fields'].items()}} for r in recs]
                    st.session_state.tabla_actual = tabla_sel
                    st.rerun()
    st.divider()
    if 'raw_records' in st.session_state:
        if st.button("📮 Postales", type="primary" if st.session_state.active_module == "Postales" else "secondary", use_container_width=True): 
            st.session_state.active_module = "Postales"; st.session_state.dia_seleccionado = None; st.rerun()
        if st.button("📄 Reportes", type="primary" if st.session_state.active_module == "Reportes" else "secondary", use_container_width=True): 
            st.session_state.active_module = "Reportes"; st.session_state.dia_seleccionado = None; st.rerun()
        if st.button("📆 Calendario", type="primary" if st.session_state.active_module == "Calendario" else "secondary", use_container_width=True): 
            st.session_state.active_module = "Calendario"; st.session_state.dia_seleccionado = None; st.rerun()

# --- MAIN ---
if 'raw_records' not in st.session_state:
    st.info("👈 Conecta una base.")
else:
    modulo = st.session_state.active_module
    
    # --- POSTALES Y REPORTES ---
    if modulo in ["Postales", "Reportes"]:
        st.subheader(f"📮 Generador de {modulo}")
        df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
        for c in df_full.columns:
            if isinstance(df_full[c].iloc[0], list): df_full.drop(c, axis=1, inplace=True)
        sel_all = st.checkbox("Seleccionar Todo")
        df_full.insert(0, "✅", sel_all)
        df_edit = st.data_editor(df_full, hide_index=True)
        indices = df_edit.index[df_edit["✅"]==True].tolist()
        
        if indices:
            folder_tpl = f"Plantillas/{modulo.upper()}"
            archs = [f for f in os.listdir(folder_tpl) if f.endswith('.pptx')]
            tipos = df_full.loc[indices, "Tipo"].unique()
            cols_t = st.columns(len(tipos))
            for i, t in enumerate(tipos):
                st.session_state.setdefault('config', {})
                st.session_state['config'].setdefault('plantillas', {})
                st.session_state['config']['plantillas'][t] = cols_t[i].selectbox(f"Plantilla {t}:", archs, key=f"p_{t}")

            if st.button("🚀 GENERAR ARCHIVOS"):
                p_bar = st.progress(0); zip_data = []
                for i, idx in enumerate(indices):
                    rec, orig = st.session_state.raw_records[idx]['fields'], st.session_state.raw_data_original[idx]['fields']
                    dt = datetime.strptime(rec.get('Fecha','2025-01-01'), '%Y-%m-%d')
                    ft, fs = rec.get('Tipo', 'Sin Tipo'), rec.get('Sucursal', '000')
                    nm_mes = MESES_ES[dt.month-1]
                    fcf = f"{nm_mes.capitalize()} {dt.day} de {dt.year}\n{obtener_hora_texto(rec.get('Hora',''))}"
                    fcc = f"Sucursal {fs}" if ft == "Actividad en Sucursal" else obtener_concat_texto(rec)
                    base_n = re.sub(r'[\\/*?:"<>|]', "", f"{dt.day} de {nm_mes} de {dt.year} - {ft}, {fs} - {fcc}")[:120]
                    
                    try:
                        prs = Presentation(f"{folder_tpl}/{st.session_state['config']['plantillas'][ft]}")
                        for slide in prs.slides:
                            for shp in slide.shapes:
                                if shp.has_text_frame:
                                    tags_img = ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Reporte firmado", "Lista de asistencia"]
                                    for tag in tags_img:
                                        if f"<<{tag}>>" in shp.text_frame.text and orig.get(tag):
                                            img_io = procesar_imagen_inteligente(requests.get(orig[tag][0]['url']).content, shp.width, shp.height, True)
                                            slide.shapes.add_picture(img_io, shp.left, shp.top, shp.width, shp.height)
                                            shp._element.getparent().remove(shp._element)
                            for shp in slide.shapes:
                                if shp.has_text_frame:
                                    reps = {"<<Tipo>>":ft, "<<Sucursal>>":fs, "<<Confechor>>":fcf, "<<Concat>>":fcc, "<<Consuc>>":fcc}
                                    for tag, val in reps.items():
                                        if tag in shp.text_frame.text:
                                            tf = shp.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                                            tf.clear(); p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
                                            run = p.add_run(); run.text=str(val); run.font.bold=True; run.font.color.rgb=RGBColor(0, 176, 240); run.font.size=Pt(28 if tag=="<<Confechor>>" else 24)
                        buf = BytesIO(); prs.save(buf)
                        dout = generar_pdf(buf.getvalue())
                        if dout:
                            path_zip = f"Provident/{dt.year}/{str(dt.month).zfill(2)} - {nm_mes.capitalize()}/{modulo}/{fs}/"
                            if modulo == "Postales":
                                img_b = convert_from_bytes(dout, dpi=170)[0]
                                with BytesIO() as b:
                                    img_b.save(b, format="JPEG", quality=85); data_f = b.getvalue()
                                zip_data.append({"n": f"{path_zip}{base_n}.jpg", "d": data_f})
                            else: zip_data.append({"n": f"{path_zip}{base_n}.pdf", "d": dout})
                    except: pass
                    p_bar.progress((i+1)/len(indices))
                
                if zip_data:
                    z_buf = BytesIO()
                    with zipfile.ZipFile(z_buf, "w") as z:
                        for f in zip_data: z.writestr(f["n"], f["d"])
                    st.download_button("⬇️ DESCARGAR PACK", z_buf.getvalue(), "Provident.zip", "application/zip")

    # --- CALENDARIO ---
    elif modulo == "Calendario":
        fechas_oc = {}
        for r in st.session_state.raw_data_original:
            f = r['fields'].get('Fecha')
            if f:
                fk = f.split('T')[0]
                if fk not in fechas_oc: fechas_oc[fk] = []
                th = r['fields']['Postal'][0].get('url') if 'Postal' in r['fields'] else None
                fechas_oc[fk].append({"thumb": th, "raw_fields": r['fields']})

        if st.session_state.dia_seleccionado:
            k = st.session_state.dia_seleccionado
            evts = sorted(fechas_oc[k], key=lambda x: x['raw_fields'].get('Hora', ''))
            total = len(evts)
            curr_idx = st.session_state.idx_postal
            evt = evts[curr_idx]
            dt_obj = datetime.strptime(k, '%Y-%m-%d')
            mes_n, dia_n = MESES_ES[dt_obj.month-1].upper(), f"{DIAS_ES[dt_obj.weekday()].capitalize()} {dt_obj.day}"

            # TABLA DE NAVEGACIÓN UNIFICADA (Flecha - Título - Flecha)
            st.markdown(f"""
            <table class="table-nav">
                <tr>
                    <td style="width:60px;">
                        <a href="?nav=prev" target="_self" class="btn-arrow">←</a>
                    </td>
                    <td>
                        <div class="nav-title"><b>{mes_n}</b><span>{dia_n}</span></div>
                    </td>
                    <td style="width:60px;">
                        <a href="?nav=next" target="_self" class="btn-arrow">→</a>
                    </td>
                </tr>
            </table>
            """, unsafe_allow_html=True)
            
            # Lógica de navegación vía URL
            params = st.query_params
            if params.get("nav") == "next":
                st.session_state.idx_postal = (curr_idx + 1) % total
                st.query_params.clear(); st.rerun()
            if params.get("nav") == "prev":
                st.session_state.idx_postal = (curr_idx - 1) % total
                st.query_params.clear(); st.rerun()

            if st.button("🔙 VOLVER AL CALENDARIO", key="btn_volver"):
                st.session_state.dia_seleccionado = None; st.rerun()
            
            st.divider()
            
            # Layout Postal y Detalles (Lado a lado)
            col_a, col_b = st.columns([1, 1.2])
            with col_a:
                if evt['thumb']: st.image(evt['thumb'], use_container_width=True)
            with col_b:
                f = evt['raw_fields']
                suc, tip, hor = f.get('Sucursal',''), f.get('Tipo',''), obtener_hora_texto(f.get('Hora',''))
                ubi = obtener_concat_texto(f)
                st.markdown(f"**🏢 {suc}**\n\n**📌 {tip}**\n\n**⏰ {hor}**\n\n**📍 {ubi}**")
                
                group = WHATSAPP_GROUPS.get(str(suc).lower().strip(), {"link": "", "name": "N/A"})
                msj = f"Excelente día, te esperamos este {dia_n} de {mes_n.capitalize()} para el evento de {tip}, a las {hor} en {ubi}"
                jwa = f"<script>function c(){{navigator.clipboard.writeText(`{msj}`).then(()=>{{window.open('{group['link']}','_blank');}});}}</script><div onclick='c()' style='background:#25D366;color:white;padding:10px;text-align:center;border-radius:5px;cursor:pointer;font-weight:bold;margin-top:10px;'>📲 WhatsApp {group['name']}</div>"
                if group['link']: st.components.v1.html(jwa, height=75)

        else:
            # Vista Calendario (7 columnas forzadas)
            if fechas_oc:
                dt_ref = datetime.strptime(list(fechas_oc.keys())[0], '%Y-%m-%d')
                st.markdown(f"<div class='cal-title'>{MESES_ES[dt_ref.month-1].upper()} {dt_ref.year}</div>", unsafe_allow_html=True)
                ch = st.columns(7)
                for i, d in enumerate(["L","M","M","J","V","S","D"]): ch[i].markdown(f"<div class='c-head'>{d}</div>", unsafe_allow_html=True)
                weeks = calendar.Calendar(0).monthdayscalendar(dt_ref.year, dt_ref.month)
                for week in weeks:
                    cols = st.columns(7)
                    for i, d in enumerate(week):
                        with cols[i]:
                            if d > 0:
                                k = f"{dt_ref.year}-{str(dt_ref.month).zfill(2)}-{str(d).zfill(2)}"
                                evs = fechas_oc.get(k, [])
                                bg = f"background-image: url('{evs[0]['thumb']}');" if evs and evs[0]['thumb'] else ""
                                st.markdown(f"<div class='c-cell-container'><div style='display:flex;flex-direction:column;height:100%;'><div class='c-day'>{d}</div><div class='c-body' style=\"{bg}\"></div><div class='{'c-foot' if evs else ''}'>{f'+{len(evs)-1}' if len(evs)>1 else ''}</div></div></div>", unsafe_allow_html=True)
                                if evs:
                                    if st.button(" ", key=f"day_{k}"): st.session_state.dia_seleccionado=k; st.session_state.idx_postal=0; st.rerun()
