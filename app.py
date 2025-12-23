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
#  CONFIGURACIÓN
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
#  FUNCIONES
# ============================================================
def procesar_texto_maestro(texto, campo=""):
    if not texto or str(texto).lower() == "none": return ""
    t = str(texto).replace('/', ' ').strip().replace('\n', ' ').replace('\r', ' ')
    if campo == 'Seccion': return t.upper()
    palabras = t.lower().split()
    resultado = [p.capitalize() if i == 0 or p not in ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al'] else p for i, p in enumerate(palabras)]
    return " ".join(resultado)

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
#  APP STREAMLIT
# ============================================================
st.set_page_config(page_title="Provident Pro v147", layout="wide")

st.markdown("""
<style>
    /* MARGEN SUPERIOR FIJO */
    .block-container { 
        padding-top: 90px !important; 
        padding-left: 5px !important;
        padding-right: 5px !important;
    }
    
    /* CALENDARIO: 7 COLUMNAS SIN ESPACIOS LATERALES */
    [data-testid="column"] { min-width: 0px !important; flex: 1 1 0% !important; padding: 1px !important; }
    div[data-testid="stHorizontalBlock"] { gap: 0px !important; width: 100% !important; max-width: 100% !important; }

    /* TÍTULO DEL MES */
    .cal-title-container {
        background: linear-gradient(135deg, #002060 0%, #00b0f0 100%);
        padding: 12px; border-radius: 10px; margin-bottom: 15px; text-align: center;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    .cal-title-text { color: white !important; font-size: 1.3em; font-weight: 800; text-transform: uppercase; margin: 0; }

    /* CONTENEDOR DE CELDA Y BOTÓN INVISIBLE */
    .cell-wrapper {
        position: relative;
        width: 100%;
        height: 80px; /* Altura ideal para móvil vertical */
        border: 0.5px solid #ddd;
        overflow: hidden;
    }
    
    .c-day-num { background: #00b0f0; color: white; font-weight: bold; font-size: 0.75em; text-align: center; height: 16px; line-height: 16px; }
    .c-img-body { height: 64px; background-size: cover; background-position: center; }

    /* FORZAR BOTÓN DE STREAMLIT A SER INVISIBLE Y CUBRIR TODA LA CELDA */
    .cell-wrapper div[data-testid="stButton"] {
        position: absolute !important;
        top: 0 !important;
        left: 0 !important;
        width: 100% !important;
        height: 100% !important;
        z-index: 5;
    }
    .cell-wrapper button {
        width: 100% !important;
        height: 100% !important;
        background-color: transparent !important;
        border: none !important;
        color: transparent !important;
        padding: 0 !important;
        margin: 0 !important;
    }
    
    /* CABECERA DIAS SEMANA */
    .c-week-head { background: #002060; color: white; font-size: 10px; font-weight: bold; text-align: center; padding: 4px 0; }
</style>
""", unsafe_allow_html=True)

if 'active_module' not in st.session_state: st.session_state.active_module = "Calendario"
if 'dia_seleccionado' not in st.session_state: st.session_state.dia_seleccionado = None

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# SIDEBAR (Simplificado para el ejemplo)
with st.sidebar:
    st.header("⚙️ Configuración")
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_sel = st.selectbox("Base:", [b['name'] for b in r_bases.json()['bases']])
        # ... (aquí iría el resto de la lógica de carga de Airtable que ya tienes)

# --- MÓDULO CALENDARIO ---
if 'raw_records' in st.session_state:
    if st.session_state.active_module == "Calendario":
        fechas_oc = {}
        for r in st.session_state.raw_data_original:
            f = r['fields'].get('Fecha')
            if f:
                fk = f.split('T')[0]
                if fk not in fechas_oc: fechas_oc[fk] = []
                th = r['fields'].get('Postal', [{}])[0].get('url')
                fechas_oc[fk].append({"thumb": th, "fields": r['fields']})

        if st.session_state.dia_seleccionado:
            # VISTA DETALLE (Postales del día)
            if st.button("⬅️ VOLVER AL MES"): 
                st.session_state.dia_seleccionado = None
                st.rerun()
            # ... (Lógica de detalle aquí)
        else:
            # VISTA MES (Calendario Vertical)
            if fechas_oc:
                dt_ref = datetime.strptime(list(fechas_oc.keys())[0], '%Y-%m-%d')
                st.markdown(f"<div class='cal-title-container'><p class='cal-title-text'>{MESES_ES[dt_ref.month-1].upper()} {dt_ref.year}</p></div>", unsafe_allow_html=True)

                cols_h = st.columns(7)
                for i, d in enumerate(["L","M","X","J","V","S","D"]):
                    cols_h[i].markdown(f"<div class='c-week-head'>{d}</div>", unsafe_allow_html=True)

                cal = calendar.Calendar(0)
                weeks = cal.monthdayscalendar(dt_ref.year, dt_ref.month)
                
                for week in weeks:
                    cols = st.columns(7)
                    for i, day in enumerate(week):
                        with cols[i]:
                            if day > 0:
                                k = f"{dt_ref.year}-{str(dt_ref.month).zfill(2)}-{str(day).zfill(2)}"
                                evs = fechas_oc.get(k, [])
                                has_ev = len(evs) > 0
                                bg = f"background-image: url('{evs[0]['thumb']}');" if has_ev and evs[0]['thumb'] else ""
                                
                                # HTML de la celda + Botón invisible inyectado dentro del wrapper
                                st.markdown(f"""
                                <div class="cell-wrapper">
                                    <div class="c-day-num">{day}</div>
                                    <div class="c-img-body" style="{bg}"></div>
                                </div>
                                """, unsafe_allow_html=True)
                                
                                if has_ev:
                                    # El botón se coloca "encima" por CSS (wrapper div[data-testid="stButton"])
                                    if st.button(" ", key=f"day_{k}"):
                                        st.session_state.dia_seleccionado = k
                                        st.rerun()
