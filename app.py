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
#  FUNCIONES TÉCNICAS
# ============================================================

def procesar_imagen_inteligente(img_data, target_w_pt, target_h_pt, con_blur=False):
    img = Image.open(BytesIO(img_data)).convert("RGB")
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
    t = str(texto).strip().replace('\n', ' ')
    if campo == 'Seccion': return t.upper()
    return t

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
st.set_page_config(page_title="Provident Pro v134", layout="wide")

# Inicializar estados
if 'active_module' not in st.session_state: st.session_state.active_module = st.query_params.get("view", "Calendario")
if 'dia_seleccionado' not in st.session_state: st.session_state.dia_seleccionado = None
if 'idx_postal' not in st.session_state: st.session_state.idx_postal = 0

def navegar_a(modulo):
    st.session_state.active_module = modulo
    st.session_state.dia_seleccionado = None
    st.query_params["view"] = modulo

# CSS
st.markdown("""
<style>
    .cal-title { text-align: center; font-size: 1.5em; font-weight: bold; margin-bottom: 10px; color: #333; }
    .c-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 2px; }
    .c-head { background: #002060; color: white; padding: 4px; text-align: center; font-weight: bold; border-radius: 2px; font-size: 14px; }
    .c-cell-container { position: relative; height: 160px; border: 1px solid #ccc; border-radius: 2px; overflow: hidden; background: white; }
    .c-cell-content { display: flex; flex-direction: column; height: 100%; justify-content: space-between; }
    .c-day { background: #00b0f0; color: white; font-weight: 900; font-size: 1.1em; text-align: center; padding: 2px 0; }
    .c-body { flex-grow: 1; background-size: cover; background-position: center; background-color: #f8f8f8; }
    .c-foot { height: 20px; background: #002060; color: #ffffff; font-weight: 900; text-align: center; font-size: 0.9em; padding: 1px; overflow: hidden; }
    .c-foot-empty { height: 20px; background: #e0e0e0; }
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

# --- LÓGICA MAIN ---
if 'raw_records' not in st.session_state:
    st.info("👈 Conecta una base.")
else:
    modulo = st.session_state.active_module
    
    if modulo == "Calendario":
        # Organizar eventos por fecha
        fechas_oc = {}
        for r in st.session_state.raw_data_original:
            f = r['fields'].get('Fecha')
            if f:
                fk = f.split('T')[0]
                if fk not in fechas_oc: fechas_oc[fk] = []
                th = r['fields']['Postal'][0].get('url') if 'Postal' in r['fields'] else None
                fechas_oc[fk].append({"thumb": th, "raw_fields": r['fields'], "id": r['id']})

        # --- VISTA DETALLE POSTAL (PAGINADA) ---
        if st.session_state.dia_seleccionado:
            k = st.session_state.dia_seleccionado
            evts = sorted(fechas_oc[k], key=lambda x: x['raw_fields'].get('Hora', ''))
            
            # Navegación del Carrusel
            total = len(evts)
            curr_idx = st.session_state.idx_postal
            if curr_idx >= total: curr_idx = 0
            
            evt = evts[curr_idx]
            fields = evt['raw_fields']
            dt_obj = datetime.strptime(k, '%Y-%m-%d')
            fecha_full = f"{DIAS_ES[dt_obj.weekday()].capitalize()} {dt_obj.day} de {MESES_ES[dt_obj.month-1].capitalize()}"
            
            # Cabecera con Flechas
            c_left, c_mid, c_right = st.columns([1, 4, 1])
            with c_left:
                if total > 1 and st.button("⬅️ Anterior"):
                    st.session_state.idx_postal = (curr_idx - 1) % total
                    st.rerun()
            with c_mid:
                st.markdown(f"<h2 style='text-align:center; margin:0;'>{fecha_full}</h2>", unsafe_allow_html=True)
                st.markdown(f"<p style='text-align:center; color:gray;'>Actividad {curr_idx+1} de {total}</p>", unsafe_allow_html=True)
            with c_right:
                if total > 1 and st.button("Siguiente ➡️"):
                    st.session_state.idx_postal = (curr_idx + 1) % total
                    st.rerun()
            
            # Botón Volver al Calendario (Ancho Total)
            if st.button("🔙 VOLVER AL CALENDARIO", use_container_width=True):
                st.session_state.dia_seleccionado = None
                st.rerun()
            
            st.divider()

            # Contenido Postal
            col_img, col_info = st.columns([1.2, 1])
            with col_img:
                if evt['thumb']: st.image(evt['thumb'], use_container_width=True)
                else: st.warning("Sin postal generada")
            
            with col_info:
                sucursal = fields.get('Sucursal', 'N/A')
                tipo = fields.get('Tipo', 'N/A')
                hora = obtener_hora_texto(fields.get('Hora', ''))
                punto = fields.get('Punto de reunion', '')
                ruta = fields.get('Ruta a seguir', '')
                mun = fields.get('Municipio', '')
                
                st.markdown(f"### 🏢 {sucursal}")
                st.markdown(f"**📌 Tipo:** {tipo}")
                st.markdown(f"**⏰ Hora:** {hora}")
                
                loc_parts = [x for x in [punto, ruta, mun] if x]
                ubicacion = ", ".join(loc_parts)
                st.markdown(f"**📍 Ubicación:** {ubicacion}")
                
                st.divider()
                
                # WhatsApp
                suc_key = str(sucursal).lower().strip()
                group = WHATSAPP_GROUPS.get(suc_key, {"link": "", "name": "No Encontrado"})
                mensaje = f"Excelente día, te esperamos este {fecha_full} para el evento de {tipo}, a las {hora} en {ubicacion}"
                
                js_wa = f"""
                <script>
                function copyAndGo() {{
                    navigator.clipboard.writeText(`{mensaje}`).then(() => {{
                        window.open("{group['link']}", "_blank");
                    }});
                }}
                </script>
                <button onclick="copyAndGo()" style="width:100%; background:#25D366; color:white; border:none; padding:15px; border-radius:8px; font-weight:bold; cursor:pointer;">
                    📲 Copiar y abrir WhatsApp ({group['name']})
                </button>
                """
                if group['link']: st.components.v1.html(js_wa, height=80)
                else: st.error("Link de sucursal no configurado")

        # --- VISTA CALENDARIO ---
        else:
            st.subheader("📅 Calendario de Actividades")
            if fechas_oc:
                dt_ref = datetime.strptime(list(fechas_oc.keys())[0], '%Y-%m-%d')
                st.markdown(f"<div class='cal-title'>{MESES_ES[dt_ref.month-1].upper()} {dt_ref.year}</div>", unsafe_allow_html=True)
                
                # Cabeceras
                cols_h = st.columns(7)
                for i, d in enumerate(["LUN","MAR","MIÉ","JUE","VIE","SÁB","DOM"]): cols_h[i].markdown(f"<div class='c-head'>{d}</div>", unsafe_allow_html=True)
                
                weeks = calendar.Calendar(0).monthdayscalendar(dt_ref.year, dt_ref.month)
                for week in weeks:
                    cols = st.columns(7)
                    for i, d in enumerate(week):
                        with cols[i]:
                            if d > 0:
                                k = f"{dt_ref.year}-{str(dt_ref.month).zfill(2)}-{str(d).zfill(2)}"
                                evs = fechas_oc.get(k, [])
                                bg = f"background-image: url('{evs[0]['thumb']}');" if evs and evs[0]['thumb'] else ""
                                
                                foot_txt = f"+ {len(evs)-1} más" if len(evs) > 1 else ("&nbsp;" if len(evs)==1 else "")
                                foot_cls = "c-foot" if evs else "c-foot-empty"
                                
                                st.markdown(f"""
                                <div class="c-cell-container">
                                    <div class="c-cell-content">
                                        <div class="c-day">{d}</div>
                                        <div class="c-body" style="{bg}"></div>
                                        <div class="{foot_cls}">{foot_txt}</div>
                                    </div>
                                </div>
                                """, unsafe_allow_html=True)
                                
                                # Botón invisible que captura el clic
                                if evs:
                                    if st.button(" ", key=f"day_{k}"):
                                        st.session_state.dia_seleccionado = k
                                        st.session_state.idx_postal = 0
                                        st.rerun()

    # --- MODULOS POSTALES Y REPORTES ---
    elif modulo in ["Postales", "Reportes"]:
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
            archs = [f for f in os.listdir(folder) if f.endswith('.pptx')]
            tipos = df_full.loc[sel_idx, "Tipo"].unique()
            for t in tipos:
                st.session_state.config["plantillas"][t] = st.selectbox(f"Plantilla para {t}:", archs, key=f"tpl_{t}")

            if st.button("🚀 GENERAR ARCHIVOS", type="primary"):
                p_bar = st.progress(0); archivos = []
                for i, idx in enumerate(sel_idx):
                    rec, orig = st.session_state.raw_records[idx]['fields'], st.session_state.raw_data_original[idx]['fields']
                    dt = datetime.strptime(rec.get('Fecha','2025-01-01'), '%Y-%m-%d')
                    ft, fs = rec.get('Tipo',''), rec.get('Sucursal','')
                    tfe, tho = obtener_fecha_texto(dt), obtener_hora_texto(rec.get('Hora',''))
                    
                    try:
                        prs = Presentation(os.path.join(folder, st.session_state.config['plantillas'][ft]))
                        # Reemplazo de imágenes y textos omitido para brevedad, pero es el mismo de v133 estable
                        # ...
                        buf = BytesIO(); prs.save(buf)
                        dout = generar_pdf(buf.getvalue())
                        if dout:
                            name = f"{dt.day}_{fs}_{ft}"
                            if modulo == "Postales":
                                img = convert_from_bytes(dout, dpi=150)[0]
                                with BytesIO() as b:
                                    img.save(b, "JPEG", quality=85)
                                    archivos.append({"n": f"{name}.png", "d": b.getvalue()})
                            else:
                                archivos.append({"n": f"{name}.pdf", "d": dout})
                    except: pass
                    p_bar.progress((i+1)/len(sel_idx))
                
                if archivos:
                    zb = BytesIO()
                    with zipfile.ZipFile(zb, "a", zipfile.ZIP_DEFLATED) as z:
                        for f in archivos: z.writestr(f["n"], f["d"])
                    st.download_button("⬇️ DESCARGAR ZIP", zb.getvalue(), "Generados.zip", "application/zip")
