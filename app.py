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
    "cordoba": {"link": "https://chat.whatsapp.com/EablxD3sq3n65GHWq9uKTg", "name": "CORDOBA Volanteo PDV Provident"},
    "boca del rio": {"link": "https://chat.whatsapp.com/ES2ufZaP8f3JrNyO9lpFcX", "name": "BOCA DEL RIO Volanteo PDV Provident"},
    "orizaba": {"link": "https://chat.whatsapp.com/EHITAvFTeYO5hOsS14xXJM", "name": "ORIZABA Volanteo PDV Provident"},
    "tuxtepec": {"link": "https://chat.whatsapp.com/HkKsqVFYZSn99FPdjZ7Whv", "name": "TUXTEPEC Volanteo PDV Provident"},
    "oaxaca": {"link": "https://chat.whatsapp.com/JRawICryDEf8eO2RYKqS0T", "name": "OAXACA Volanteo PDV Provident"},
    "tehuacan": {"link": "https://chat.whatsapp.com/E9z0vLSwfZCB97Ou4A6Chv", "name": "TEHUACAN Volanteo PDV Provident"},
    "xalapa": {"link": "", "name": "XALAPA Volanteo PDV Provident"},
}

# ============================================================
#  FUNCIONES TÉCNICAS (IMAGEN Y TEXTO)
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
    img_final.save(output, format="JPEG", quality=85, subsampling=0, optimize=True)
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

# ============================================================
#  LÓGICA DE DATOS Y FECHAS
# ============================================================

def obtener_fecha_texto(fecha_dt):
    dia_idx = fecha_dt.weekday()
    return f"{DIAS_ES[dia_idx]} {fecha_dt.day} de {MESES_ES[fecha_dt.month - 1]} de {fecha_dt.year}"

def obtener_hora_texto(hora_str):
    if not hora_str or str(hora_str).lower() == "none": return ""
    s_raw = str(hora_str).lower().strip()
    es_pm = any(x in s_raw for x in ["p.m.", "pm", "p. m."])
    match = re.search(r'(\d{1,2}):(\d{2})', s_raw)
    if match:
        h, m = int(match.group(1)), match.group(2)
        if es_pm and h < 12: h += 12
        h_show = h - 12 if h > 12 else (12 if h == 0 else h)
        suf_wa = "a.m." if h < 12 else "p.m."
        return f"{h_show}:{m} {suf_wa}"
    return hora_str

def obtener_concat_texto(record):
    parts = []
    fields = ['Punto de reunion', 'Ruta a seguir', 'Municipio']
    for f in fields:
        val = record.get(f)
        if val and str(val).lower() != 'none' and val != "":
            parts.append(str(val))
    return ", ".join(parts)

# ============================================================
#  DIALOGO DETALLES EVENTO (COPYS & WHATSAPP)
# ============================================================

@st.dialog("Detalles de Actividad", width="large")
def mostrar_detalles_dia(eventos_del_dia, k_fecha):
    if 'idx_evento' not in st.session_state: st.session_state.idx_evento = 0
    if st.session_state.idx_evento >= len(eventos_del_dia): st.session_state.idx_evento = 0
    elif st.session_state.idx_evento < 0: st.session_state.idx_evento = len(eventos_del_dia) - 1
        
    evt = eventos_del_dia[st.session_state.idx_evento]
    fields = evt.get('raw_fields', {})
    
    if len(eventos_del_dia) > 1:
        c1, c2, c3 = st.columns([1, 6, 1])
        if c1.button("⬅️", key="prev_btn"):
            st.session_state.idx_evento -= 1
            st.rerun()
        if c3.button("➡️", key="next_btn"):
            st.session_state.idx_evento += 1
            st.rerun()
        with c2: st.markdown(f"<p style='text-align:center;'>Evento {st.session_state.idx_evento+1} de {len(eventos_del_dia)}</p>", unsafe_allow_html=True)

    if evt.get('thumb'): st.image(evt['thumb'], use_container_width=True)
    
    sucursal_raw = str(fields.get('Sucursal', '')).lower().strip()
    tipo = fields.get('Tipo', 'Sin Tipo')
    dt_obj = datetime.strptime(k_fecha, '%Y-%m-%d')
    fecha_wa = f"{DIAS_ES[dt_obj.weekday()].capitalize()} {dt_obj.day} de {MESES_ES[dt_obj.month-1].capitalize()}"
    hora = obtener_hora_texto(fields.get('Hora', ''))
    ubicacion = obtener_concat_texto(fields)
    
    st.markdown(f"### {fields.get('Sucursal', 'Sucursal')}")
    st.markdown(f"**Tipo:** {tipo} | **Hora:** {hora}")
    st.markdown(f"**Ubicación:** {ubicacion}")

    group_info = WHATSAPP_GROUPS.get(sucursal_raw, {"link": "", "name": "Grupo Desconocido"})
    mensaje_copiar = f"Excelente día, te esperamos este {fecha_wa} para el evento de {tipo}, a las {hora} en {ubicacion}"
    
    js_component = f"""
    <div style="text-align:center;">
        <script>
        function doAction() {{
            const text = `{mensaje_copiar}`;
            navigator.clipboard.writeText(text).then(() => {{
                window.open("{group_info['link']}", "_blank");
            }}).catch(() => {{
                window.open("{group_info['link']}", "_blank");
            }});
        }}
        </script>
        <button onclick="doAction()" style="background-color:#25D366; color:white; border:none; padding:15px; border-radius:10px; font-weight:bold; cursor:pointer; width:100%; font-size:16px;">
            Copia mensaje y abre {group_info['name']}
        </button>
    </div>
    """
    if group_info['link']:
        st.components.v1.html(js_component, height=80)
    else:
        st.warning("⚠️ Sin link de WhatsApp para esta sucursal.")

# ============================================================
#  INICIO DE LA APP
# ============================================================

st.set_page_config(page_title="Provident Pro v131", layout="wide")

if 'config' not in st.session_state:
    if os.path.exists("config_app.json"):
        with open("config_app.json", "r") as f: st.session_state.config = json.load(f)
    else: st.session_state.config = {"plantillas": {}}

if 'active_module' not in st.session_state:
    st.session_state.active_module = st.query_params.get("view", "Calendario")

def navegar_a(modulo):
    st.session_state.active_module = modulo
    st.query_params["view"] = modulo

# Estilos CSS
st.markdown("""
<style>
    .cal-title { text-align: center; font-size: 1.5em; font-weight: bold; margin-bottom: 10px; color: #002060; }
    .c-head { background: #002060; color: white; padding: 4px; text-align: center; font-weight: bold; border-radius: 2px; }
    .c-cell { background: white; border: 1px solid #ccc; border-radius: 2px; height: 160px; display: flex; flex-direction: column; overflow: hidden; }
    .c-day { background: #00b0f0; color: white; font-weight: bold; text-align: center; padding: 2px; }
    .c-body { flex-grow: 1; background-size: cover; background-position: center; background-color: #f0f0f0; }
    .c-foot { background: #002060; color: white; font-size: 0.8em; text-align: center; padding: 2px; }
</style>
""", unsafe_allow_html=True)

# Airtable Config
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

with st.sidebar:
    st.header("🔗 Airtable")
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base:", list(base_opts.keys()))
        if base_sel:
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
            if r_tab.status_code == 200:
                tablas = {t['name']: t['id'] for t in r_tab.json()['tables']}
                tabla_sel = st.selectbox("Mes:", list(tablas.keys()))
                if 'tabla_actual' not in st.session_state or st.session_state.tabla_actual != tabla_sel:
                    with st.spinner("Sincronizando..."):
                        r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tablas[tabla_sel]}", headers=headers)
                        recs = r_reg.json().get("records", [])
                        st.session_state.raw_data_original = recs
                        st.session_state.raw_records = [
                            {'id': r['id'], 'fields': {k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v) for k, v in r['fields'].items()}}
                            for r in recs
                        ]
                        st.session_state.tabla_actual = tabla_sel
                    st.rerun()

    st.divider()
    if 'raw_records' in st.session_state:
        st.button("📮 Postales", on_click=navegar_a, args=("Postales",), use_container_width=True)
        st.button("📄 Reportes", on_click=navegar_a, args=("Reportes",), use_container_width=True)
        st.button("📅 Calendario", on_click=navegar_a, args=("Calendario",), use_container_width=True)

# --- MÓDULOS PRINCIPALES ---

if 'raw_records' not in st.session_state:
    st.info("👈 Selecciona una base y tabla para comenzar.")
else:
    mod = st.session_state.active_module
    AZUL_PRO = RGBColor(0, 176, 240)

    if mod == "Postales" or mod == "Reportes":
        folder = f"Plantillas/{mod.upper()}"
        if not os.path.exists(folder): os.makedirs(folder)
        st.subheader(f"Generador de {mod}")
        
        df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
        for col in df_full.columns:
            if isinstance(df_full[col].iloc[0], list): df_full.drop(col, axis=1, inplace=True)
        
        sel_all = st.checkbox("Seleccionar todo")
        df_full.insert(0, "✅", sel_all)
        df_edit = st.data_editor(df_full, hide_index=True)
        indices = df_edit.index[df_edit["✅"] == True].tolist()

        if indices:
            archs = [f for f in os.listdir(folder) if f.endswith('.pptx')]
            tipos = df_full.loc[indices, "Tipo"].unique()
            for t in tipos:
                st.session_state.config["plantillas"][t] = st.selectbox(f"Plantilla para {t}:", archs, key=f"tpl_{t}")

            if st.button(f"🚀 Generar {len(indices)} {mod}"):
                p_bar = st.progress(0); zip_data = []
                for i, idx in enumerate(indices):
                    rec = st.session_state.raw_records[idx]['fields']
                    orig = st.session_state.raw_data_original[idx]['fields']
                    dt = datetime.strptime(rec.get('Fecha','2025-01-01'), '%Y-%m-%d')
                    ft, fs = rec.get('Tipo',''), rec.get('Sucursal','')
                    
                    tfe, tho = obtener_fecha_texto(dt), obtener_hora_texto(rec.get('Hora',''))
                    fcf = f"{MESES_ES[dt.month-1].capitalize()} {dt.day} de {dt.year}\n{tho}"
                    fcc = f"Sucursal {fs}" if ft == "Actividad en Sucursal" else obtener_concat_texto(rec)
                    
                    try:
                        prs = Presentation(os.path.join(folder, st.session_state.config["plantillas"][ft]))
                        for slide in prs.slides:
                            for shp in slide.shapes:
                                if shp.has_text_frame:
                                    tags = ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Reporte firmado", "Lista de asistencia"]
                                    for tag in tags:
                                        if f"<<{tag}>>" in shp.text_frame.text and orig.get(tag):
                                            img_io = procesar_imagen_inteligente(requests.get(orig[tag][0]['url']).content, shp.width, shp.height, con_blur=True)
                                            slide.shapes.add_picture(img_io, shp.left, shp.top, shp.width, shp.height)
                                            shp._element.getparent().remove(shp._element)
                            for shp in slide.shapes:
                                if shp.has_text_frame:
                                    reps = {"<<Tipo>>": textwrap.fill(ft, 35), "<<Sucursal>>": fs, "<<Confechor>>": fcf, "<<Concat>>": fcc, "<<Confecha>>": tfe, "<<Conhora>>": tho}
                                    for k, v in reps.items():
                                        if k in shp.text_frame.text:
                                            tf = shp.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                                            tf.clear(); p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
                                            run = p.add_run(); run.text = str(v); run.font.bold = True; run.font.color.rgb = AZUL_PRO
                                            run.font.size = Pt(28 if k == "<<Confechor>>" else 24)

                        buf = BytesIO(); prs.save(buf)
                        pdf = generar_pdf(buf.getvalue())
                        if pdf:
                            if mod == "Postales":
                                img_final = convert_from_bytes(pdf, dpi=150)[0]
                                img_buf = BytesIO(); img_final.save(img_buf, format="JPEG", quality=85)
                                zip_data.append({"n": f"{dt.day}_{ft}_{fs}.jpg", "d": img_buf.getvalue()})
                            else:
                                zip_data.append({"n": f"{dt.day}_{ft}_{fs}.pdf", "d": pdf})
                    except Exception as e: st.error(f"Error en {fs}: {e}")
                    p_bar.progress((i+1)/len(indices))
                
                if zip_data:
                    z_buf = BytesIO()
                    with zipfile.ZipFile(z_buf, "w") as z:
                        for f in zip_data: z.writestr(f["n"], f["d"])
                    st.download_button("⬇️ Descargar todo", z_buf.getvalue(), f"{mod}.zip", "application/zip")

    elif mod == "Calendario":
        st.subheader("📅 Planificador Mensual")
        fechas = {}
        for r in st.session_state.raw_data_original:
            f = r['fields'].get('Fecha')
            if f:
                f_key = f.split('T')[0]
                if f_key not in fechas: fechas[f_key] = []
                th = r['fields'].get('Postal', [{}])[0].get('url')
                fechas[f_key].append({"thumb": th, "raw_fields": r['fields']})
        
        if fechas:
            dt_ref = datetime.strptime(list(fechas.keys())[0], '%Y-%m-%d')
            st.markdown(f"<div class='cal-title'>{MESES_ES[dt_ref.month-1].upper()} {dt_ref.year}</div>", unsafe_allow_html=True)
            cal = calendar.Calendar(0).monthdayscalendar(dt_ref.year, dt_ref.month)
            
            h_cols = st.columns(7)
            for i, d in enumerate(["LUN", "MAR", "MIE", "JUE", "VIE", "SAB", "DOM"]): h_cols[i].markdown(f"<div class='c-head'>{d}</div>", unsafe_allow_html=True)
            
            for week in cal:
                cols = st.columns(7)
                for i, day in enumerate(week):
                    with cols[i]:
                        if day > 0:
                            k = f"{dt_ref.year}-{str(dt_ref.month).zfill(2)}-{str(day).zfill(2)}"
                            evs = fechas.get(k, [])
                            bg = f"background-image: url('{evs[0]['thumb']}');" if evs and evs[0]['thumb'] else ""
                            st.markdown(f"<div class='c-cell'><div class='c-day'>{day}</div><div class='c-body' style='{bg}'></div></div>", unsafe_allow_html=True)
                            if evs:
                                if st.button(f"Ver ({len(evs)})", key=f"btn_{k}", use_container_width=True):
                                    mostrar_detalles_dia(evs, k)
