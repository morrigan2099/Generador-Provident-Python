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
    for f in ['Punto de reunion', 'Ruta a seguir', 'Municipio']:
        val = record.get(f)
        if val and str(val).lower() != 'none' and val != "":
            parts.append(str(val))
    return ", ".join(parts)

# ============================================================
#  DIALOGO DETALLES EVENTO (COPYS & WHATSAPP)
# ============================================================

@st.dialog("Actividades del Día", width="large")
def mostrar_detalles_dia(eventos_del_dia, k_fecha):
    if 'idx_evt' not in st.session_state: st.session_state.idx_evt = 0
    if st.session_state.idx_evt >= len(eventos_del_dia): st.session_state.idx_evt = 0
    elif st.session_state.idx_evt < 0: st.session_state.idx_evt = len(eventos_del_dia) - 1
        
    evt = eventos_del_dia[st.session_state.idx_evt]
    fields = evt.get('raw_fields', {})
    
    if len(eventos_del_dia) > 1:
        c1, c2, c3 = st.columns([1, 4, 1])
        if c1.button("⬅️", key="modal_prev"): st.session_state.idx_evt -= 1; st.rerun()
        if c3.button("➡️", key="modal_next"): st.session_state.idx_evt += 1; st.rerun()
        c2.markdown(f"<p style='text-align:center;'>Evento {st.session_state.idx_evt+1} de {len(eventos_del_dia)}</p>", unsafe_allow_html=True)

    if evt.get('thumb'): st.image(evt['thumb'], use_container_width=True)
    
    sucursal_raw = str(fields.get('Sucursal', '')).lower().strip()
    tipo = fields.get('Tipo', 'Sin Tipo')
    dt_obj = datetime.strptime(k_fecha, '%Y-%m-%d')
    fecha_wa = f"{DIAS_ES[dt_obj.weekday()].capitalize()} {dt_obj.day} de {MESES_ES[dt_obj.month-1].capitalize()}"
    hora = obtener_hora_texto(fields.get('Hora', ''))
    ubicacion = obtener_concat_texto(fields)
    
    st.markdown(f"### {fields.get('Sucursal', 'Sucursal')}")
    st.markdown(f"**Evento:** {tipo}\n\n**Hora:** {hora}\n\n**Ubicación:** {ubicacion}")

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
    if group_info['link']: st.components.v1.html(js_component, height=80)
    else: st.warning("⚠️ Sin link de WhatsApp.")

# ============================================================
#  INICIO DE LA APP
# ============================================================

st.set_page_config(page_title="Provident Pro v133", layout="wide")

if 'active_module' not in st.session_state:
    st.session_state.active_module = st.query_params.get("view", "Calendario")

def navegar_a(modulo):
    st.session_state.active_module = modulo
    st.query_params["view"] = modulo

# --- ESTILOS CSS RECONSTRUIDOS ---
st.markdown("""
<style>
    /* CALENDARIO ORIGINAL */
    .cal-title { text-align: center; font-size: 1.5em; font-weight: bold; margin-bottom: 10px; color: #333; background-color: #fff; }
    .c-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 2px; }
    .c-head { background: #002060; color: white; padding: 4px; text-align: center; font-weight: bold; border-radius: 2px; font-size: 14px; }
    
    .c-cell-container { position: relative; height: 160px; border: 1px solid #ccc; border-radius: 2px; overflow: hidden; background: white; }
    .c-cell-content { display: flex; flex-direction: column; height: 100%; width: 100%; justify-content: space-between; }
    
    .c-day { flex: 0 0 auto; background: #00b0f0; color: white; font-weight: 900; font-size: 1.1em; text-align: center; padding: 2px 0; }
    .c-body { flex-grow: 1; width: 100%; background-position: center; background-repeat: no-repeat; background-size: cover; background-color: #f8f8f8; }
    .c-foot { flex: 0 0 auto; height: 20px; background: #002060; color: #ffffff; font-weight: 900; text-align: center; font-size: 0.9em; padding: 1px; white-space: nowrap; overflow: hidden; }
    .c-foot-empty { flex: 0 0 auto; height: 20px; background: #e0e0e0; }

    /* BOTÓN INVISIBLE QUE CUBRE LA CELDA */
    .stButton > button[key^="btn_"] {
        position: absolute; top: 0; left: 0; width: 100%; height: 100%;
        background: transparent !important; border: none !important; color: transparent !important;
        z-index: 10; cursor: pointer;
    }
    
    @media (max-width: 600px) { .c-cell-container { height: 110px; } .c-day { font-size: 0.9em; } }
</style>
""", unsafe_allow_html=True)

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# --- SIDEBAR ORIGINAL ---
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
                if 'tabla_actual' not in st.session_state or st.session_state.tabla_actual != tabla_sel:
                    r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tab_opts[tabla_sel]}", headers=headers)
                    recs = r_reg.json().get("records", [])
                    st.session_state.raw_data_original = recs
                    st.session_state.raw_records = [{'id': r['id'], 'fields': {k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v) for k, v in r['fields'].items()}} for r in recs]
                    st.session_state.tabla_actual = tabla_sel
                    st.rerun()
    st.divider()
    if 'raw_records' in st.session_state:
        st.subheader("⚡ Generar")
        if st.button("📮 Postales", type="primary" if st.session_state.active_module == "Postales" else "secondary", use_container_width=True): navegar_a("Postales"); st.rerun()
        if st.button("📄 Reportes", type="primary" if st.session_state.active_module == "Reportes" else "secondary", use_container_width=True): navegar_a("Reportes"); st.rerun()
        st.subheader("📅 Eventos")
        if st.button("📆 Calendario", type="primary" if st.session_state.active_module == "Calendario" else "secondary", use_container_width=True): navegar_a("Calendario"); st.rerun()

# --- MÓDULOS PRINCIPALES ---
if 'raw_records' not in st.session_state:
    st.info("👈 Conecta una base en el sidebar.")
else:
    modulo = st.session_state.active_module
    AZUL_PRO = RGBColor(0, 176, 240)

    if modulo == "Postales" or modulo == "Reportes":
        st.subheader(f"📮 Generador de {modulo}")
        df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
        for c in df_full.columns:
            if isinstance(df_full[c].iloc[0], list): df_full.drop(c, axis=1, inplace=True)
        sel_all = st.checkbox("Seleccionar Todo")
        df_full.insert(0, "✅", sel_all)
        df_edit = st.data_editor(df_full, hide_index=True)
        sel_idx = df_edit.index[df_edit["✅"]==True].tolist()
        
        if sel_idx:
            folder = f"Plantillas/{modulo.upper()}"
            if not os.path.exists(folder): os.makedirs(folder)
            archs = [f for f in os.listdir(folder) if f.endswith('.pptx')]
            tipos = df_full.loc[sel_idx, "Tipo"].unique()
            cols = st.columns(len(tipos))
            if 'config' not in st.session_state: st.session_state.config = {"plantillas": {}}
            for i, t in enumerate(tipos):
                st.session_state.config["plantillas"][t] = cols[i].selectbox(f"Plantilla para {t}:", archs, key=f"p_{t}")

            if st.button("🔥 GENERAR ARCHIVOS", type="primary"):
                p_bar = st.progress(0); archivos = []
                for i, idx in enumerate(sel_idx):
                    rec, orig = st.session_state.raw_records[idx]['fields'], st.session_state.raw_data_original[idx]['fields']
                    dt = datetime.strptime(rec.get('Fecha','2025-01-01'), '%Y-%m-%d')
                    ft, fs = rec.get('Tipo', 'Sin Tipo'), rec.get('Sucursal', '000')
                    tfe, tho = obtener_fecha_texto(dt), obtener_hora_texto(rec.get('Hora',''))
                    tfe_c = f"{MESES_ES[dt.month-1].capitalize()} {dt.day} de {dt.year}"
                    fcf = f"{tfe_c}\n{tho}"
                    fcc = f"Sucursal {fs}" if ft == "Actividad en Sucursal" else obtener_concat_texto(rec)
                    narc = re.sub(r'[\\/*?:"<>|]', "", f"{dt.day} de {MESES_ES[dt.month-1]} - {ft}, {fs}")
                    
                    try:
                        prs = Presentation(f"{folder}/{st.session_state.config['plantillas'][ft]}")
                        for slide in prs.slides:
                            for shp in slide.shapes:
                                if shp.has_text_frame:
                                    tags = ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Reporte firmado", "Lista de asistencia"]
                                    for tag in tags:
                                        if f"<<{tag}>>" in shp.text_frame.text and orig.get(tag):
                                            img_io = procesar_imagen_inteligente(requests.get(orig[tag][0]['url']).content, shp.width, shp.height, True)
                                            slide.shapes.add_picture(img_io, shp.left, shp.top, shp.width, shp.height)
                                            shp._element.getparent().remove(shp._element)
                            for shp in slide.shapes:
                                if shp.has_text_frame:
                                    reps = {"<<Tipo>>":textwrap.fill(ft,35), "<<Sucursal>>":fs, "<<Confechor>>":fcf, "<<Concat>>":fcc, "<<Consuc>>":fcc, "<<Confecha>>":tfe, "<<Conhora>>":tho}
                                    for tag, val in reps.items():
                                        if tag in shp.text_frame.text:
                                            tf = shp.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE; tf.word_wrap = True
                                            if tag == "<<Confechor>>":
                                                tf.clear()
                                                p1 = tf.paragraphs[0]; p1.text = tfe_c; p1.alignment = PP_ALIGN.CENTER; p1.font.bold = True; p1.font.color.rgb = AZUL_PRO; p1.font.size = Pt(28)
                                                p2 = tf.add_paragraph(); p2.text = tho; p2.alignment = PP_ALIGN.CENTER; p2.font.bold = True; p2.font.color.rgb = AZUL_PRO; p2.font.size = Pt(28)
                                            else:
                                                tf.clear(); p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; run = p.add_run(); run.text=str(val); run.font.bold=True; run.font.color.rgb=AZUL_PRO; run.font.size=Pt(32 if tag!="<<Tipo>>" else 12)
                        buf = BytesIO(); prs.save(buf)
                        dout = generar_pdf(buf.getvalue())
                        if dout:
                            if modulo == "Postales":
                                img_bytes = convert_from_bytes(dout, dpi=170)[0]
                                with BytesIO() as b:
                                    img_bytes.save(b, format="JPEG", quality=85)
                                    archivos.append({"Ruta": f"{dt.year}/Postales/{fs}/{narc}.png", "Data": b.getvalue()})
                            else:
                                archivos.append({"Ruta": f"{dt.year}/Reportes/{fs}/{narc}.pdf", "Data": dout})
                    except: pass
                    p_bar.progress((i+1)/len(sel_idx))
                
                if archivos:
                    zb = BytesIO()
                    with zipfile.ZipFile(zb, "a", zipfile.ZIP_DEFLATED) as z:
                        for f in archivos: z.writestr(f["Ruta"], f["Data"])
                    st.download_button(f"⬇️ DESCARGAR {len(archivos)} ARCHIVOS", zb.getvalue(), "Generados.zip", "application/zip", type="primary")

    elif modulo == "Calendario":
        st.subheader("📅 Calendario de Actividades")
        fechas_oc = {}
        for r in st.session_state.raw_data_original:
            f = r['fields'].get('Fecha')
            if f:
                f_short = f.split('T')[0]
                if f_short not in fechas_oc: fechas_oc[f_short] = []
                th = r['fields']['Postal'][0].get('url') if 'Postal' in r['fields'] else None
                fechas_oc[f_short].append({"thumb": th, "raw_fields": r['fields']})
        
        if fechas_oc:
            dt_list = [datetime.strptime(x, '%Y-%m-%d') for x in fechas_oc.keys()]
            mc = Counter([(d.year, d.month) for d in dt_list])
            ay, am = mc.most_common(1)[0][0]
            st.markdown(f"<div class='cal-title'>📅 {MESES_ES[am-1].upper()} {ay}</div>", unsafe_allow_html=True)
            
            weeks = calendar.Calendar(0).monthdayscalendar(ay, am)
            # Cabeceras
            cols_h = st.columns(7)
            for i, d in enumerate(["LUN","MAR","MIÉ","JUE","VIE","SÁB","DOM"]): 
                cols_h[i].markdown(f"<div class='c-head'>{d}</div>", unsafe_allow_html=True)
            
            # Grid
            for week in weeks:
                cols = st.columns(7)
                for i, d in enumerate(week):
                    with cols[i]:
                        if d == 0:
                            st.markdown("<div style='height:160px;'></div>", unsafe_allow_html=True)
                        else:
                            k = f"{ay}-{str(am).zfill(2)}-{str(d).zfill(2)}"
                            acts = fechas_oc.get(k, [])
                            bg = f"background-image: url('{acts[0]['thumb']}');" if acts and acts[0]['thumb'] else ""
                            
                            foot_txt = "&nbsp;"
                            foot_cls = "c-foot-empty"
                            if len(acts) > 1:
                                foot_txt = f"+ {len(acts)-1} más"
                                foot_cls = "c-foot"
                            elif len(acts) == 1:
                                foot_cls = "c-foot"
                            
                            # ESTRUCTURA VISUAL ORIGINAL
                            st.markdown(f"""
                            <div class="c-cell-container">
                                <div class="c-cell-content">
                                    <div class="c-day">{d}</div>
                                    <div class="c-body" style="{bg}"></div>
                                    <div class="{foot_cls}">{foot_txt}</div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                            
                            # BOTÓN INVISIBLE PARA EL CLIC
                            if acts:
                                if st.button(" ", key=f"btn_{k}"):
                                    mostrar_detalles_dia(acts, k)
