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
#  CONFIGURACIÓN GLOBAL
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

def recorte_inteligente_bordes(img):
    img_gray = img.convert("L")
    arr = np.array(img_gray)
    h, w = arr.shape
    top, bottom = 0, h - 1
    while top < h and np.mean(arr[top, :]) < 35: top += 1
    while bottom > top and np.mean(arr[bottom, :]) < 35: bottom -= 1
    left, right = 0, w - 1
    while left < w and np.mean(arr[:, left]) < 35: left += 1
    while right > left and np.mean(arr[:, right]) < 35: right -= 1
    return img.crop((left, top, right + 1, bottom + 1))

def procesar_imagen_inteligente(img_data, target_w_pt, target_h_pt, con_blur=False):
    img = Image.open(BytesIO(img_data)).convert("RGB")
    img = recorte_inteligente_bordes(img)
    base_w, base_h = int(target_w_pt / 9525), int(target_h_pt / 9525)
    render_w, render_h = base_w * 2, base_h * 2
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
    img_final.save(output, format="JPEG", quality=85)
    output.seek(0)
    return output

def procesar_texto_maestro(texto, campo=""):
    if not texto or str(texto).lower() == "none": return ""
    t = str(texto).replace('/', ' ').strip().replace('\n', ' ').replace('\r', ' ')
    palabras = t.lower().split()
    resultado = [p.capitalize() if i == 0 or p not in ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al'] else p for i, p in enumerate(palabras)]
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
#  ESTILOS Y CONFIGURACIÓN UI
# ============================================================
st.set_page_config(page_title="Provident Pro v152", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    /* FORZAR SIDEBAR EN MÓVIL */
    @media (max-width: 768px) {
        section[data-testid="stSidebar"] { width: 80% !important; display: block !important; visibility: visible !important; }
        [data-testid="stSidebarCollapsedControl"] { display: block !important; left: 10px !important; }
    }

    header[data-testid="stHeader"] { visibility: hidden; height: 0; }
    .block-container { padding-top: 20px !important; }

    /* TABLA CALENDARIO INAMOVIBLE */
    .cal-wrapper { width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; }
    .cal-table-fixed { width: 100%; min-width: 320px; border-collapse: collapse; table-layout: fixed; }
    .cal-table-fixed th { background: #002060; color: white; font-size: 10px; padding: 4px 0; text-align: center; }
    .cal-table-fixed td { 
        height: 70px; border: 0.5px solid #ddd; position: relative; padding: 0; vertical-align: top;
        background-color: white; overflow: hidden;
    }
    .td-num { background: #00b0f0; color: white; font-size: 9px; font-weight: bold; text-align: center; height: 14px; }
    .td-img { height: 46px; background-size: cover; background-position: center; width: 100%; }
    .td-link { position: absolute; top:0; left:0; width:100%; height:100%; text-decoration: none !important; color: transparent !important; z-index: 10; }
    
    /* NAV DETALLE */
    .nav-box { display: flex; align-items: center; justify-content: space-between; background: white; padding: 8px; margin-bottom: 5px; border-bottom: 2px solid #00b0f0; }
    .nav-title-text { text-align: center; line-height: 1.1; }
    .nav-title-text b { color: #002060; font-size: 1.1em; display: block; }

    div.stButton > button[key="btn_volver"] { background-color: #00b0f0 !important; color: white !important; font-weight: bold !important; width: 100%; border: none !important; }
</style>
""", unsafe_allow_html=True)

if 'active_module' not in st.session_state: st.session_state.active_module = "Calendario"
if 'dia_seleccionado' not in st.session_state: st.session_state.dia_seleccionado = None
if 'idx_postal' not in st.session_state: st.session_state.idx_postal = 0

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://www.provident.com.mx/content/dam/provident-mexico/logos/logo-provident.png", width=150)
    st.header("🔗 Airtable")
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base:", list(base_opts.keys()))
        if base_sel:
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
            if r_tab.status_code == 200:
                tab_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
                tabla_sel = st.selectbox("Mes:", list(tab_opts.keys()))
                if st.session_state.get('tabla_actual') != tabla_sel:
                    r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tab_opts[tabla_sel]}", headers=headers)
                    recs = r_reg.json().get("records", [])
                    st.session_state.raw_data_original = recs
                    st.session_state.raw_records = [{'id': r['id'], 'fields': {k: procesar_texto_maestro(v, k) for k, v in r['fields'].items()}} for r in recs]
                    st.session_state.tabla_actual = tabla_sel
                    st.rerun()
    st.divider()
    if 'raw_records' in st.session_state:
        if st.button("📆 Calendario", use_container_width=True): st.session_state.active_module = "Calendario"; st.rerun()
        if st.button("📮 Postales", use_container_width=True): st.session_state.active_module = "Postales"; st.rerun()
        if st.button("📄 Reportes", use_container_width=True): st.session_state.active_module = "Reportes"; st.rerun()

# --- MAIN ---
if 'raw_records' not in st.session_state:
    st.info("👈 Conecta una base.")
else:
    mod = st.session_state.active_module

    if mod == "Calendario":
        fechas_oc = {}
        for r in st.session_state.raw_data_original:
            f = r['fields'].get('Fecha')
            if f:
                fk = f.split('T')[0]
                if fk not in fechas_oc: fechas_oc[fk] = []
                th = r['fields'].get('Postal', [{}])[0].get('url') if 'Postal' in r['fields'] else None
                fechas_oc[fk].append({"thumb": th, "raw": r['fields']})

        params = st.query_params
        if "day" in params: st.session_state.dia_seleccionado = params["day"]
        
        # VALIDACIÓN ANTIERRORES LINEA 200
        if st.session_state.dia_seleccionado and st.session_state.dia_seleccionado in fechas_oc:
            k = st.session_state.dia_seleccionado
            evs = sorted(fechas_oc[k], key=lambda x: x['raw'].get('Hora',''))
            curr = st.session_state.idx_postal % len(evs)
            evt = evs[curr]
            dt = datetime.strptime(k, '%Y-%m-%d')
            mes_n, dia_n = MESES_ES[dt.month-1].upper(), f"{DIAS_ES[dt.weekday()].capitalize()} {dt.day}"

            st.markdown(f"""
            <div class="nav-box">
                <a href="?day={k}&nav=prev" target="_self" style="background:#00b0f0; color:white; padding:8px 12px; border-radius:50%; text-decoration:none; font-weight:bold;">←</a>
                <div class="nav-title-text"><b>{mes_n}</b>{dia_n}</div>
                <a href="?day={k}&nav=next" target="_self" style="background:#00b0f0; color:white; padding:8px 12px; border-radius:50%; text-decoration:none; font-weight:bold;">→</a>
            </div>
            """, unsafe_allow_html=True)
            
            if "nav" in params:
                st.session_state.idx_postal += 1 if params["nav"] == "next" else -1
                st.query_params.clear(); st.query_params["day"] = k; st.rerun()

            if st.button("🔙 VOLVER AL MES", key="btn_volver"):
                st.session_state.dia_seleccionado = None; st.query_params.clear(); st.rerun()

            if evt['thumb']: st.image(evt['thumb'], use_container_width=True)
            f = evt['raw']
            suc, tip, hor = f.get('Sucursal',''), f.get('Tipo',''), obtener_hora_texto(f.get('Hora',''))
            ubi = obtener_concat_texto(f)
            st.markdown(f"**🏢 {suc}**\n\n**📌 {tip}** | **⏰ {hor}**\n\n**📍 {ubi}**")
            
            sk = str(suc).lower().strip()
            gp = WHATSAPP_GROUPS.get(sk, {"link":"", "name":"N/A"})
            msj = f"Excelente día, te esperamos este {dia_n} de {mes_n.capitalize()} para el evento de {tip}, a las {hor} en {ubi}"
            jwa = f"<script>function c(){{navigator.clipboard.writeText(`{msj}`).then(()=>{{window.open('{gp['link']}','_blank');}});}}</script><div onclick='c()' style='background:#25D366;color:white;padding:12px;text-align:center;border-radius:8px;cursor:pointer;font-weight:bold;'>📲 WhatsApp {gp['name']}</div>"
            if gp['link']: st.components.v1.html(jwa, height=80)
        else:
            # MOSTRAR CALENDARIO SI NO HAY DÍA O SI EL DÍA SELECCIONADO NO EXISTE
            st.session_state.dia_seleccionado = None
            if fechas_oc:
                dt_ref = datetime.strptime(list(fechas_oc.keys())[0], '%Y-%m-%d')
                st.markdown(f"<div style='background:#002060; color:white; padding:10px; text-align:center; font-weight:bold; border-radius:8px; margin-bottom:5px;'>{MESES_ES[dt_ref.month-1].upper()} {dt_ref.year}</div>", unsafe_allow_html=True)
                
                cal_html = '<div class="cal-wrapper"><table class="cal-table-fixed"><tr><th>L</th><th>M</th><th>X</th><th>J</th><th>V</th><th>S</th><th>D</th></tr>'
                weeks = calendar.Calendar(0).monthdayscalendar(dt_ref.year, dt_ref.month)
                for week in weeks:
                    cal_html += '<tr>'
                    for d in week:
                        if d == 0: cal_html += '<td></td>'
                        else:
                            fk = f"{dt_ref.year}-{str(dt_ref.month).zfill(2)}-{str(d).zfill(2)}"
                            evs = fechas_oc.get(fk, [])
                            bg = f"background-image: url('{evs[0]['thumb']}');" if evs and evs[0]['thumb'] else ""
                            foot = f"+{len(evs)-1}" if len(evs)>1 else ""
                            link = f'<a href="?day={fk}" target="_self" class="td-link">.</a>' if evs else ""
                            cal_html += f'<td><div class="td-num">{d}</div><div class="td-img" style="{bg}"></div><div style="font-size:7px; text-align:center; background:#002060; color:white;">{foot}</div>{link}</td>'
                    cal_html += '</tr>'
                cal_html += '</table></div>'
                st.markdown(cal_html, unsafe_allow_html=True)

    elif mod in ["Postales", "Reportes"]:
        st.subheader(f"Generador de {mod}")
        df = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
        for c in df.columns:
            if isinstance(df[c].iloc[0], list): df.drop(c, axis=1, inplace=True)
        sel_all = st.checkbox("Seleccionar Todo")
        df.insert(0, "✅", sel_all)
        df_edit = st.data_editor(df, hide_index=True)
        idx = df_edit.index[df_edit["✅"]==True].tolist()
        
        if idx:
            f_tpl = f"Plantillas/{mod.upper()}"
            archs = [f for f in os.listdir(f_tpl) if f.endswith('.pptx')]
            tipos = df.loc[idx, "Tipo"].unique()
            if 'config' not in st.session_state: st.session_state.config = {"plantillas": {}}
            for t in tipos:
                st.session_state.config["plantillas"][t] = st.selectbox(f"Plantilla {t}:", archs, key=f"p_{t}")

            if st.button("🚀 GENERAR"):
                p_bar = st.progress(0); zip_data = []
                for i, ix in enumerate(idx):
                    rec, orig = st.session_state.raw_records[ix]['fields'], st.session_state.raw_data_original[ix]['fields']
                    dt = datetime.strptime(rec.get('Fecha','2025-01-01'), '%Y-%m-%d')
                    ft, fs, mn = rec.get('Tipo',''), rec.get('Sucursal',''), MESES_ES[dt.month-1]
                    fcf = f"{mn.capitalize()} {dt.day} de {dt.year}\n{obtener_hora_texto(rec.get('Hora',''))}"
                    fcc = f"Sucursal {fs}" if ft == "Actividad en Sucursal" else obtener_concat_texto(rec)
                    
                    try:
                        prs = Presentation(f"{f_tpl}/{st.session_state.config['plantillas'][ft]}")
                        for slide in prs.slides:
                            for shp in slide.shapes:
                                if shp.has_text_frame:
                                    for tag in ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Reporte firmado", "Lista de asistencia"]:
                                        if f"<<{tag}>>" in shp.text_frame.text and orig.get(tag):
                                            img_io = procesar_imagen_inteligente(requests.get(orig[tag][0]['url']).content, shp.width, shp.height, True)
                                            slide.shapes.add_picture(img_io, shp.left, shp.top, shp.width, shp.height)
                                            shp._element.getparent().remove(shp._element)
                            for shp in slide.shapes:
                                if shp.has_text_frame:
                                    reps = {"<<Tipo>>":ft, "<<Sucursal>>":fs, "<<Confechor>>":fcf, "<<Concat>>":fcc, "<<Consuc>>":fcc}
                                    for t_k, v in reps.items():
                                        if t_k in shp.text_frame.text:
                                            tf = shp.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                                            tf.clear(); p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
                                            run = p.add_run(); run.text=str(v); run.font.bold=True; run.font.color.rgb=RGBColor(0, 176, 240); run.font.size=Pt(28 if t_k=="<<Confechor>>" else 24)
                        buf = BytesIO(); prs.save(buf)
                        dout = generar_pdf(buf.getvalue())
                        if dout:
                            path = f"Provident/{dt.year}/{str(dt.month).zfill(2)} - {mn.capitalize()}/{mod}/{fs}/"
                            name = re.sub(r'[\\/*?:"<>|]', "", f"{dt.day} de {mn} de {dt.year} - {ft}, {fs} - {fcc}")[:120]
                            if mod == "Postales":
                                img_b = convert_from_bytes(dout, dpi=170)[0]
                                with BytesIO() as b:
                                    img_b.save(b, format="JPEG", quality=85); zip_data.append({"n": f"{path}{name}.jpg", "d": b.getvalue()})
                            else: zip_data.append({"n": f"{path}{name}.pdf", "d": dout})
                    except: pass
                    p_bar.progress((i+1)/len(idx))
                if zip_data:
                    z_buf = BytesIO()
                    with zipfile.ZipFile(z_buf, "w") as z:
                        for f in zip_data: z.writestr(f["n"], f["d"])
                    st.download_button("⬇️ DESCARGAR ZIP", z_buf.getvalue(), "Provident.zip", "application/zip")
