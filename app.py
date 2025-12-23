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
#  FUNCIONES TÉCNICAS (IMAGEN, TEXTO Y PDF)
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

def obtener_fecha_texto(fecha_dt):
    return f"{DIAS_ES[fecha_dt.weekday()]} {fecha_dt.day} de {MESES_ES[fecha_dt.month - 1]} de {fecha_dt.year}"

# ============================================================
#  INICIO APP
# ============================================================
st.set_page_config(page_title="Provident Pro v138", layout="wide")

if 'active_module' not in st.session_state: st.session_state.active_module = st.query_params.get("view", "Calendario")
if 'dia_seleccionado' not in st.session_state: st.session_state.dia_seleccionado = None
if 'idx_postal' not in st.session_state: st.session_state.idx_postal = 0

def navegar_a(modulo):
    st.session_state.active_module = modulo
    st.session_state.dia_seleccionado = None
    st.query_params["view"] = modulo

# CSS FORZADO PARA MÓVIL (Versión 137 Mejorada)
st.markdown("""
<style>
    [data-testid="column"] { min-width: 0px !important; flex: 1 1 0% !important; }
    .cal-title { text-align: center; font-size: 1.3em; font-weight: bold; margin-bottom: 10px; }
    .c-head { background: #002060; color: white; padding: 4px; text-align: center; font-weight: bold; border-radius: 2px; font-size: 10px; }
    .c-cell-container { position: relative; height: 110px; border: 1px solid #ccc; border-radius: 2px; overflow: hidden; background: white; }
    .c-cell-content { display: flex; flex-direction: column; height: 100%; justify-content: space-between; }
    .c-day { background: #00b0f0; color: white; font-weight: 900; font-size: 0.9em; text-align: center; padding: 1px 0; }
    .c-body { flex-grow: 1; background-size: cover; background-position: center; }
    .c-foot { height: 16px; background: #002060; color: #ffffff; font-weight: 900; text-align: center; font-size: 0.7em; padding: 1px; }
    div.stButton > button[key="btn_volver"] { background-color: #00b0f0 !important; color: white !important; border: none !important; font-weight: bold !important; width: 100%; }
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
        st.subheader("⚡ Generar")
        if st.button("📮 Postales", type="primary" if st.session_state.active_module == "Postales" else "secondary", use_container_width=True): navegar_a("Postales"); st.rerun()
        if st.button("📄 Reportes", type="primary" if st.session_state.active_module == "Reportes" else "secondary", use_container_width=True): navegar_a("Reportes"); st.rerun()
        st.subheader("📅 Eventos")
        if st.button("📆 Calendario", type="primary" if st.session_state.active_module == "Calendario" else "secondary", use_container_width=True): navegar_a("Calendario"); st.rerun()

# --- MAIN ---
if 'raw_records' not in st.session_state:
    st.info("👈 Conecta una base.")
else:
    modulo = st.session_state.active_module
    AZUL_PRO = RGBColor(0, 176, 240)

    # ==========================================
    # LÓGICA DE POSTALES Y REPORTES (RESTAURADA)
    # ==========================================
    if modulo in ["Postales", "Reportes"]:
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
            cols_t = st.columns(len(tipos))
            if 'config' not in st.session_state: st.session_state.config = {"plantillas": {}}
            for i, t in enumerate(tipos):
                st.session_state.config["plantillas"][t] = cols_t[i].selectbox(f"Plantilla {t}:", archs, key=f"p_{t}")

            if st.button("🚀 GENERAR ARCHIVOS", type="primary"):
                p_bar = st.progress(0); archivos = []
                for i, idx in enumerate(sel_idx):
                    rec, orig = st.session_state.raw_records[idx]['fields'], st.session_state.raw_data_original[idx]['fields']
                    dt = datetime.strptime(rec.get('Fecha','2025-01-01'), '%Y-%m-%d')
                    ft, fs = rec.get('Tipo', 'Sin Tipo'), rec.get('Sucursal', '000')
                    tfe, tho = obtener_fecha_texto(dt), obtener_hora_texto(rec.get('Hora',''))
                    tfe_c = f"{MESES_ES[dt.month-1].capitalize()} {dt.day} de {dt.year}"
                    fcf = f"{tfe_c}\n{tho}"
                    fcc = f"Sucursal {fs}" if ft == "Actividad en Sucursal" else obtener_concat_texto(rec)
                    narc = re.sub(r'[\\/*?:"<>|]', "", f"{dt.day}_{fs}_{ft}")
                    
                    try:
                        prs = Presentation(f"{folder}/{st.session_state.config['plantillas'][ft]}")
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
                                    reps = {"<<Tipo>>":ft, "<<Sucursal>>":fs, "<<Confechor>>":fcf, "<<Concat>>":fcc, "<<Consuc>>":fcc, "<<Confecha>>":tfe, "<<Conhora>>":tho}
                                    for tag, val in reps.items():
                                        if tag in shp.text_frame.text:
                                            tf = shp.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
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
                                img_b = convert_from_bytes(dout, dpi=150)[0]
                                with BytesIO() as b:
                                    img_b.save(b, format="JPEG", quality=85)
                                    archivos.append({"n": f"{narc}.jpg", "d": b.getvalue()})
                            else: archivos.append({"n": f"{narc}.pdf", "d": dout})
                    except: pass
                    p_bar.progress((i+1)/len(sel_idx))
                
                if archivos:
                    zb = BytesIO()
                    with zipfile.ZipFile(zb, "w") as z:
                        for f in archivos: z.writestr(f["n"], f["d"])
                    st.download_button("⬇️ DESCARGAR ZIP", zb.getvalue(), f"{modulo}.zip", "application/zip")

    # ==========================================
    # LÓGICA DE CALENDARIO (MÓVIL FORZADO)
    # ==========================================
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
            fields = evt['raw_fields']
            dt_obj = datetime.strptime(k, '%Y-%m-%d')
            fecha_full = f"{DIAS_ES[dt_obj.weekday()].capitalize()} {dt_obj.day} de {MESES_ES[dt_obj.month-1].capitalize()}"
            
            # Cabecera forzada 1 fila
            c1, c2, c3 = st.columns([1, 4, 1])
            with c1: 
                if total > 1 and st.button("⬅️", key="p"): st.session_state.idx_postal=(curr_idx-1)%total; st.rerun()
            with c2: st.markdown(f"<h4 style='text-align:center; margin:0;'>{fecha_full}</h4>", unsafe_allow_html=True)
            with c3: 
                if total > 1 and st.button("➡️", key="n"): st.session_state.idx_postal=(curr_idx+1)%total; st.rerun()
            
            if st.button("🔙 VOLVER AL CALENDARIO", key="btn_volver"):
                st.session_state.dia_seleccionado = None; st.rerun()
            
            st.divider()
            cx, cy = st.columns([1, 1.2])
            with cx:
                if evt['thumb']: st.image(evt['thumb'], use_container_width=True)
            with cy:
                suc, tip, hor = fields.get('Sucursal',''), fields.get('Tipo',''), obtener_hora_texto(fields.get('Hora',''))
                ubi = obtener_concat_texto(fields)
                st.markdown(f"**🏢 {suc}**\n\n**📌 {tip}**\n\n**⏰ {hor}**\n\n**📍 {ubi}**")
                
                # WhatsApp
                sk = str(suc).lower().strip()
                gp = WHATSAPP_GROUPS.get(sk, {"link": "", "name": "N/A"})
                msj = f"Excelente día, te esperamos este {fecha_full} para el evento de {tip}, a las {hor} en {ubi}"
                jwa = f"<script>function c(){{navigator.clipboard.writeText(`{msj}`).then(()=>{{window.open('{gp['link']}','_blank');}});}}</script><div onclick='c()' style='background:#25D366;color:white;padding:10px;text-align:center;border-radius:5px;cursor:pointer;font-weight:bold;'>📲 WhatsApp {gp['name']}</div>"
                if gp['link']: st.components.v1.html(jwa, height=60)

        else:
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
                                st.markdown(f"<div class='c-cell-container'><div class='c-cell-content'><div class='c-day'>{d}</div><div class='c-body' style=\"{bg}\"></div><div class='{'c-foot' if evs else 'c-foot-empty'}'>{f'+{len(evs)-1}' if len(evs)>1 else ''}</div></div></div>", unsafe_allow_html=True)
                                if evs:
                                    if st.button(" ", key=f"day_{k}"): st.session_state.dia_seleccionado=k; st.session_state.idx_postal=0; st.rerun()
