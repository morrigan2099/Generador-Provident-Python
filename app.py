import streamlit as st
import requests
import pandas as pd
import json
import os
import re
import numpy as np
import calendar
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import subprocess, tempfile, zipfile
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes
from PIL import Image, ImageOps, ImageFilter

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

def procesar_texto_maestro(texto, campo=""):
    if not texto or str(texto).lower() == "none": return ""
    t = str(texto).replace('/', ' ').strip().replace('\n', ' ').replace('\r', ' ')
    palabras = t.lower().split()
    resultado = [p.capitalize() if i == 0 or p not in ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al'] else p for i, p in enumerate(palabras)]
    return " ".join(resultado)

# ============================================================
#  INICIO APP
# ============================================================
st.set_page_config(page_title="Provident Pro v159", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    /* Reducción de espacio superior */
    .block-container { padding-top: 40px !important; }

    .cal-container-main {
        max-width: 400px;
        margin: 0 auto;
        position: relative;
    }

    /* La tabla visual */
    .cal-grid-table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
    }
    .cal-grid-table th { background: #002060; color: white; font-size: 10px; padding: 5px 0; border: 0.2px solid white; }
    .cal-grid-table td { 
        border: 0.5px solid #ccc; 
        height: 110px; 
        vertical-align: top; 
        padding: 0 !important;
        background: white;
        position: relative;
    }

    .cell-day-num { background: #00b0f0; color: white; font-weight: bold; font-size: 0.85em; text-align: center; height: 16px; line-height: 16px; }
    .cell-img { height: 80px; background-size: cover; background-position: center; }
    .cell-foot { height: 14px; background: #002060; color: white; text-align: center; font-size: 8px; line-height: 14px; }

    /* Los botones invisibles: ahora los forzamos a subir 110px para quedar sobre la celda */
    .stButton > button[key^="day_"] {
        position: relative;
        top: -110px; /* Sube exactamente el alto de la celda */
        width: 100% !important;
        height: 110px !important;
        background: transparent !important;
        border: none !important;
        color: transparent !important;
        z-index: 100 !important;
        margin-bottom: -110px !important; /* Evita que el botón ocupe espacio abajo */
        display: block !important;
    }
    
    /* Limpieza de márgenes de columnas en el calendario */
    [data-testid="column"] { padding: 0 !important; }
    
    .nav-title { text-align: center; line-height: 1.2; }
    .nav-title b { color: #002060; font-size: 1.2em; display: block; text-transform: uppercase; }
</style>
""", unsafe_allow_html=True)

# --- SESIÓN ---
if 'dia_seleccionado' not in st.session_state: st.session_state.dia_seleccionado = None
if 'active_module' not in st.session_state: st.session_state.active_module = "Calendario"

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

with st.sidebar:
    st.image("https://www.provident.com.mx/content/dam/provident-mexico/logos/logo-provident.png", width=120)
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Seleccionar Base:", list(base_opts.keys()))
        if base_sel:
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
            if r_tab.status_code == 200:
                tab_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
                tabla_sel = st.selectbox("Seleccionar Mes:", list(tab_opts.keys()))
                if st.session_state.get('tabla_actual') != tabla_sel:
                    r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tab_opts[tabla_sel]}", headers=headers)
                    recs = r_reg.json().get("records", [])
                    st.session_state.raw_data_original = recs
                    st.session_state.raw_records = [{'id': r['id'], 'fields': {k: procesar_texto_maestro(v, k) for k, v in r['fields'].items()}} for r in recs]
                    st.session_state.tabla_actual = tabla_sel
                    st.rerun()

if 'raw_records' in st.session_state:
    if st.session_state.dia_seleccionado:
        # --- VISTA DETALLE ---
        k = st.session_state.dia_seleccionado
        fechas_oc = {}
        for r in st.session_state.raw_data_original:
            f = r['fields'].get('Fecha')
            if f:
                fk = f.split('T')[0]
                if fk not in fechas_oc: fechas_oc[fk] = []
                th = r['fields'].get('Postal', [{}])[0].get('url') if 'Postal' in r['fields'] else None
                fechas_oc[fk].append({"thumb": th, "raw": r['fields']})
        
        evs = sorted(fechas_oc[k], key=lambda x: x['raw'].get('Hora',''))
        evt = evs[0]
        dt = datetime.strptime(k, '%Y-%m-%d')
        st.markdown(f"<div class='nav-title'><b>{MESES_ES[dt.month-1].upper()}</b><span>Día {dt.day}</span></div>", unsafe_allow_html=True)
        if st.button("⬅️ VOLVER"): st.session_state.dia_seleccionado = None; st.rerun()
        if evt['thumb']: st.image(evt['thumb'], use_container_width=True)
        st.write(f"**{evt['raw'].get('Sucursal')}** - {evt['raw'].get('Tipo')}")
    
    else:
        # --- CALENDARIO PRINCIPAL ---
        fechas_oc = {}
        for r in st.session_state.raw_data_original:
            f = r['fields'].get('Fecha')
            if f:
                fk = f.split('T')[0]
                if fk not in fechas_oc: fechas_oc[fk] = []
                th = r['fields'].get('Postal', [{}])[0].get('url') if 'Postal' in r['fields'] else None
                fechas_oc[fk].append({"thumb": th})

        if fechas_oc:
            dt_ref = datetime.strptime(list(fechas_oc.keys())[0], '%Y-%m-%d')
            st.markdown(f"<div class='cal-container-main'>", unsafe_allow_html=True)
            st.markdown(f"<div style='background:#002060;color:white;text-align:center;padding:10px;font-weight:bold;text-transform:uppercase;'>{MESES_ES[dt_ref.month-1]} {dt_ref.year}</div>", unsafe_allow_html=True)
            
            weeks = calendar.Calendar(0).monthdayscalendar(dt_ref.year, dt_ref.month)
            
            # 1. Dibujamos la tabla HTML (El fondo visual)
            table_html = '<table class="cal-grid-table"><tr><th>L</th><th>M</th><th>X</th><th>J</th><th>V</th><th>S</th><th>D</th></tr>'
            for week in weeks:
                table_html += '<tr>'
                for day in week:
                    if day == 0: table_html += '<td></td>'
                    else:
                        fk = f"{dt_ref.year}-{str(dt_ref.month).zfill(2)}-{str(day).zfill(2)}"
                        evs = fechas_oc.get(fk, [])
                        bg = f"background-image: url('{evs[0]['thumb']}');" if evs and evs[0]['thumb'] else ""
                        label = f'+{len(evs)-1}' if len(evs)>1 else ''
                        table_html += f'<td><div class="cell-day-num">{day}</div><div class="cell-img" style="{bg}"></div><div class="cell-foot">{label}</div></td>'
                table_html += '</tr>'
            table_html += '</table>'
            st.markdown(table_html, unsafe_allow_html=True)

            # 2. Dibujamos los botones invisibles (Capa funcional)
            # Estos botones se desplazan hacia arriba mediante el CSS .stButton > button[key^="day_"]
            for week in weeks:
                cols = st.columns(7)
                for i, day in enumerate(week):
                    if day > 0:
                        fk = f"{dt_ref.year}-{str(dt_ref.month).zfill(2)}-{str(day).zfill(2)}"
                        with cols[i]:
                            if st.button(" ", key=f"day_{fk}"):
                                st.session_state.dia_seleccionado = fk
                                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
