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
#  CONFIGURACIÓN
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
#  APP STREAMLIT
# ============================================================

st.set_page_config(page_title="Provident Pro v99", layout="wide")

if 'config' not in st.session_state:
    if os.path.exists("config_app.json"):
        with open("config_app.json", "r") as f: st.session_state.config = json.load(f)
    else: st.session_state.config = {"plantillas": {}}

st.title("🚀 Generador Pro v99 - Nativo")
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# --- SIDEBAR ---
with st.sidebar:
    st.header("🔗 Conexión")
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base:", [""]+list(base_opts.keys()))
        if base_sel:
            st.session_state['base_activa_id'] = base_opts[base_sel]
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
            if r_tab.status_code == 200:
                tablas_data = r_tab.json()['tables']
                st.session_state['todas_tablas'] = {t['name']: t['id'] for t in tablas_data}
                tabla_sel = st.selectbox("Tabla Inicial:", list(st.session_state['todas_tablas'].keys()))
                
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
    AZUL = RGBColor(0, 176, 240)

    # --------------------------------------------------------
    # MÓDULO POSTALES
    # --------------------------------------------------------
    if modulo == "📮 Postales":
        st.subheader("📮 Generador de Postales")
        # ... (Código de Postales idéntico al anterior, omitido para brevedad en esta respuesta crítica, pero usa el de v98 si lo necesitas)
        # Te pongo la estructura básica para que no rompa:
        st.info("Módulo de postales activo. Configura en el sidebar.")
        # Aquí iría el código completo de postales si lo necesitas, pero enfoquémonos en el calendario.

    # --------------------------------------------------------
    # MÓDULO CALENDARIO (NATIVO STREAMLIT)
    # --------------------------------------------------------
    elif modulo == "📅 Calendario Visual":
        st.subheader("📅 Calendario de Actividades")
        
        # Selector Tabla
        if 'todas_tablas' in st.session_state:
            nt = list(st.session_state['todas_tablas'].keys())
            curr = st.session_state.get('tabla_actual_nombre')
            idx = nt.index(curr) if curr in nt else 0
            new_t = st.selectbox("Seleccionar Mes (Tabla):", nt, index=idx)
            if new_t != curr:
                # Recarga simple
                bid = st.session_state['base_activa_id']
                tid = st.session_state['todas_tablas'][new_t]
                rreg = requests.get(f"https://api.airtable.com/v0/{bid}/{tid}", headers=headers)
                recs = rreg.json().get("records", [])
                st.session_state.raw_data_original = recs
                st.session_state.raw_records = [{'id':r['id'],'fields':r['fields']} for r in recs]
                st.session_state['tabla_actual_nombre'] = new_t
                st.rerun()

        st.divider()

        # Datos
        fechas_oc = {}
        fechas_lista = []
        for r in st.session_state.raw_data_original:
            f = r['fields'].get('Fecha')
            if f:
                fs = f.split('T')[0]
                if fs not in fechas_oc: fechas_oc[fs] = []
                th = None
                if 'Postal' in r['fields']:
                    att = r['fields']['Postal']
                    if isinstance(att, list) and len(att)>0: th = att[0].get('thumbnails',{}).get('large',{}).get('url')
                fechas_oc[fs].append({"id":r['id'], "thumb":th})
                fechas_lista.append(fs)

        # DIAGNÓSTICO
        with st.expander("🕵️ Diagnóstico de Datos (Si no ves el calendario, abre aquí)"):
            st.write(f"Total eventos encontrados: {len(fechas_lista)}")
            st.write("Fechas detectadas:", fechas_lista)

        if not fechas_oc:
            st.warning("No hay fechas en esta tabla.")
        else:
            # Detectar mes
            dt_objs = [datetime.strptime(x, '%Y-%m-%d') for x in fechas_lista]
            mc = Counter([(d.year, d.month) for d in dt_objs])
            ay, am = mc.most_common(1)[0][0]
            st.markdown(f"### 📅 {MESES_ES[am-1].capitalize()} {ay}")

            # DIBUJAR CALENDARIO CON NATIVOS (st.columns)
            # Esto es indestructible visualmente.
            cal = calendar.Calendar(firstweekday=0) 
            weeks = cal.monthdayscalendar(ay, am)
            
            # Cabecera Días
            cols = st.columns(7)
            for i, d in enumerate(["LUN","MAR","MIÉ","JUE","VIE","SÁB","DOM"]):
                cols[i].markdown(f"<div style='text-align:center; font-weight:bold; background:#eee;'>{d}</div>", unsafe_allow_html=True)
            
            # Cuerpo Semanas
            for week in weeks:
                # Usamos un contenedor para cada semana
                with st.container():
                    cols = st.columns(7)
                    for i, day in enumerate(week):
                        with cols[i]:
                            if day != 0:
                                # Contenedor del día (Borde visual simple con CSS nativo si es posible, si no, layout limpio)
                                st.markdown(f"**{day}**") # HEADER: DÍA NEGRILLA
                                
                                k = f"{ay}-{str(am).zfill(2)}-{str(day).zfill(2)}"
                                acts = fechas_oc.get(k, [])
                                
                                if acts:
                                    # BODY: IMAGEN (NATIVO)
                                    if acts[0]['thumb']:
                                        st.image(acts[0]['thumb'], use_container_width=True)
                                    else:
                                        st.caption("Sin postal")
                                    
                                    # FOOTER: MAS (NATIVO)
                                    if len(acts) > 1:
                                        st.error(f"+ {len(acts)-1} más")
                                else:
                                    # Espacio vacío para mantener alineación (opcional)
                                    st.write("") 
                            else:
                                st.write("") # Día vacío del mes anterior/siguiente
                st.divider() # Línea separadora entre semanas
