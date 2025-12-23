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
import subprocess, tempfile, zipfile
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes
from PIL import Image, ImageOps, ImageFilter

# ============================================================
#  FUNCIONES TÉCNICAS (LÓGICA DE PROCESAMIENTO)
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
#  INTERFAZ Y ESTILOS CSS PRO (SOLUCIÓN OVERLAY)
# ============================================================
st.set_page_config(page_title="Provident Pro v163", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    .block-container { padding-top: 20px !important; }
    
    /* EL CONTENEDOR QUE LO CORRIGE TODO */
    .cal-container-fixed {
        position: relative;
        max-width: 400px;
        margin: 0 auto;
    }

    /* LA TABLA VISUAL (FONDO) */
    .cal-grid-table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        z-index: 1;
    }
    .cal-grid-table th { background: #002060; color: white; font-size: 10px; padding: 5px 0; border: 0.1px solid white; }
    .cal-grid-table td { border: 0.5px solid #ccc; height: 110px; vertical-align: top; padding: 0 !important; background: white; }

    .cell-day-num { background: #00b0f0; color: white; font-weight: bold; font-size: 0.85em; text-align: center; height: 16px; line-height: 16px; }
    .cell-img { height: 80px; background-size: cover; background-position: center; }
    .cell-foot { height: 14px; background: #002060; color: white; text-align: center; font-size: 8px; line-height: 14px; }

    /* LA CAPA DE BOTONES (FRENTE) */
    /* Usamos position: absolute para que no ocupe espacio físico abajo */
    .cal-button-layer {
        position: absolute;
        top: 26px; /* Bajamos los botones para que empiecen después del L-M-X... */
        left: 0;
        width: 100%;
        z-index: 999;
        pointer-events: none; /* Permite que el contenedor ignore clicks pero sus hijos no */
    }
    
    .cal-button-layer [data-testid="column"] { padding: 0 !important; }

    /* Estilo de los botones invisibles de Streamlit */
    .cal-button-layer button {
        width: 100% !important;
        height: 110px !important;
        background: transparent !important;
        border: none !important;
        color: transparent !important;
        margin: 0 !important;
        padding: 0 !important;
        pointer-events: auto; /* El botón sí recibe el click */
        box-shadow: none !important;
    }
    .cal-button-layer button:hover {
        background: rgba(0, 176, 240, 0.1) !important; /* Ligero efecto hover para feedback */
    }

</style>
""", unsafe_allow_html=True)

# --- VARIABLES DE ESTADO ---
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]

if 'active_module' not in st.session_state: st.session_state.active_module = "Calendario"
if 'dia_seleccionado' not in st.session_state: st.session_state.dia_seleccionado = None
if 'idx_postal' not in st.session_state: st.session_state.idx_postal = 0

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://www.provident.com.mx/content/dam/provident-mexico/logos/logo-provident.png", width=120)
    st.header("⚙️ Configuración")
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base Airtable:", list(base_opts.keys()))
        if base_sel:
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
            if r_tab.status_code == 200:
                tab_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
                tabla_sel = st.selectbox("Mes (Tabla):", list(tab_opts.keys()))
                if st.session_state.get('tabla_actual') != tabla_sel:
                    r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tab_opts[tabla_sel]}", headers=headers)
                    recs = r_reg.json().get("records", [])
                    st.session_state.raw_data_original = recs
                    st.session_state.raw_records = [{'id': r['id'], 'fields': {k: procesar_texto_maestro(v, k) for k, v in r['fields'].items()}} for r in recs]
                    st.session_state.tabla_actual = tabla_sel
                    st.rerun()
    st.divider()
    if 'raw_records' in st.session_state:
        if st.button("📆 Ver Calendario", use_container_width=True): st.session_state.active_module = "Calendario"; st.session_state.dia_seleccionado = None; st.rerun()
        if st.button("📮 Hacer Postales", use_container_width=True): st.session_state.active_module = "Postales"; st.rerun()
        if st.button("📄 Hacer Reportes", use_container_width=True): st.session_state.active_module = "Reportes"; st.rerun()

# --- MOTOR PRINCIPAL ---
if 'raw_records' not in st.session_state:
    st.info("👋 Hola! Por favor conecta tu base de Airtable en el menú de la izquierda.")
else:
    mod = st.session_state.active_module

    if mod == "Calendario":
        # Agrupar por fecha
        fechas_oc = {}
        for r in st.session_state.raw_data_original:
            f = r['fields'].get('Fecha')
            if f:
                fk = f.split('T')[0]
                if fk not in fechas_oc: fechas_oc[fk] = []
                th = r['fields'].get('Postal', [{}])[0].get('url') if 'Postal' in r['fields'] else None
                fechas_oc[fk].append({"thumb": th, "raw": r['fields']})

        if st.session_state.dia_seleccionado:
            # --- VISTA DETALLE DEL DÍA ---
            k = st.session_state.dia_seleccionado
            evs = sorted(fechas_oc[k], key=lambda x: x['raw'].get('Hora',''))
            curr = st.session_state.idx_postal % len(evs)
            evt = evs[curr]
            dt = datetime.strptime(k, '%Y-%m-%d')
            mes_n, dia_n = MESES_ES[dt.month-1].upper(), f"{DIAS_ES[dt.weekday()].capitalize()} {dt.day}"

            if st.button("⬅️ VOLVER AL CALENDARIO"): st.session_state.dia_seleccionado = None; st.rerun()
            
            st.markdown(f"### {dia_n} de {mes_n.capitalize()}")
            if evt['thumb']: st.image(evt['thumb'], use_container_width=True)
            
            f_d = evt['raw']
            st.info(f"🏢 **Sucursal:** {f_d.get('Sucursal')}\n\n📌 **Actividad:** {f_d.get('Tipo')}\n\n⏰ **Hora:** {obtener_hora_texto(f_d.get('Hora'))}\n\n📍 **Lugar:** {obtener_concat_texto(f_d)}")

        else:
            # --- VISTA MES CON OVERLAY CORREGIDO ---
            if fechas_oc:
                dt_ref = datetime.strptime(list(fechas_oc.keys())[0], '%Y-%m-%d')
                weeks = calendar.Calendar(0).monthdayscalendar(dt_ref.year, dt_ref.month)
                
                # Encabezado del Mes
                st.markdown(f"<div style='background:#002060; color:white; padding:10px; border-radius:8px; text-align:center; font-weight:bold; margin:0 auto 10px auto; max-width:400px;'>{MESES_ES[dt_ref.month-1].upper()} {dt_ref.year}</div>", unsafe_allow_html=True)

                # INICIO DEL CONTENEDOR RELATIVO
                st.markdown('<div class="cal-container-fixed">', unsafe_allow_html=True)
                
                # 1. DIBUJAMOS LA TABLA (CAPA FONDO)
                table_html = '<table class="cal-grid-table"><tr><th>L</th><th>M</th><th>X</th><th>J</th><th>V</th><th>S</th><th>D</th></tr>'
                for week in weeks:
                    table_html += '<tr>'
                    for day in week:
                        if day == 0: table_html += '<td></td>'
                        else:
                            fk = f"{dt_ref.year}-{str(dt_ref.month).zfill(2)}-{str(day).zfill(2)}"
                            evs = fechas_oc.get(fk, [])
                            bg = f"background-image: url('{evs[0]['thumb']}');" if evs and evs[0]['thumb'] else ""
                            table_html += f'<td><div class="cell-day-num">{day}</div><div class="cell-img" style="{bg}"></div><div class="cell-foot">{f"+{len(evs)-1}" if len(evs)>1 else ""}</div></td>'
                    table_html += '</tr>'
                st.markdown(table_html + '</table>', unsafe_allow_html=True)

                # 2. DIBUJAMOS LOS BOTONES (CAPA FRENTE - ABSOLUTA)
                st.markdown('<div class="cal-button-layer">', unsafe_allow_html=True)
                for week in weeks:
                    cols = st.columns(7)
                    for i, day in enumerate(week):
                        if day > 0:
                            fk = f"{dt_ref.year}-{str(dt_ref.month).zfill(2)}-{str(day).zfill(2)}"
                            with cols[i]:
                                # El botón existe físicamente aquí pero CSS lo encima en la rejilla
                                if st.button(" ", key=f"btn_{fk}"):
                                    st.session_state.dia_seleccionado = fk
                                    st.rerun()
                st.markdown('</div></div>', unsafe_allow_html=True)

    elif mod in ["Postales", "Reportes"]:
        # (Se mantiene la lógica funcional de las versiones anteriores)
        st.subheader(f"Módulo de {mod}")
        df = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
        for c in df.columns:
            if isinstance(df[c].iloc[0], list): df.drop(c, axis=1, inplace=True)
        sel_all = st.checkbox("Seleccionar Todos")
        df.insert(0, "✅", sel_all)
        df_edit = st.data_editor(df, hide_index=True)
        idx_list = df_edit.index[df_edit["✅"]==True].tolist()
        
        if idx_list:
            f_tpl = f"Plantillas/{mod.upper()}"
            archs = [f for f in os.listdir(f_tpl) if f.endswith('.pptx')]
            tipos = df.loc[idx_list, "Tipo"].unique()
            if 'config' not in st.session_state: st.session_state.config = {"plantillas": {}}
            for t in tipos: st.session_state.config["plantillas"][t] = st.selectbox(f"Plantilla para {t}:", archs, key=f"p_{t}")

            if st.button("🚀 INICIAR PROCESAMIENTO"):
                p_bar = st.progress(0); zip_data = []
                for i, ix in enumerate(idx_list):
                    rec, orig = st.session_state.raw_records[ix]['fields'], st.session_state.raw_data_original[ix]['fields']
                    dt = datetime.strptime(rec.get('Fecha','2025-01-01'), '%Y-%m-%d')
                    ft, fs, mes_n = rec.get('Tipo',''), rec.get('Sucursal',''), MESES_ES[dt.month-1]
                    fcf = f"{mes_n.capitalize()} {dt.day} de {dt.year}\n{obtener_hora_texto(rec.get('Hora',''))}"
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
                        pdf = generar_pdf(buf.getvalue())
                        if pdf:
                            path = f"Provident/{dt.year}/{str(dt.month).zfill(2)} - {mes_n.capitalize()}/{mod}/{fs}/"
                            name = re.sub(r'[\\/*?:"<>|]', "", f"{dt.day} de {mes_n} de {dt.year} - {ft}, {fs} - {fcc}")[:120]
                            if mod == "Postales":
                                img_b = convert_from_bytes(pdf, dpi=170)[0]
                                with BytesIO() as b:
                                    img_b.save(b, format="JPEG", quality=85); zip_data.append({"n": f"{path}{name}.jpg", "d": b.getvalue()})
                            else: zip_data.append({"n": f"{path}{name}.pdf", "d": pdf})
                    except: pass
                    p_bar.progress((i+1)/len(idx_list))
                if zip_data:
                    z_buf = BytesIO()
                    with zipfile.ZipFile(z_buf, "w") as z:
                        for f in zip_data: z.writestr(f["n"], f["d"])
                    st.download_button("⬇️ DESCARGAR RESULTADOS (ZIP)", z_buf.getvalue(), f"Pack_Provident.zip", "application/zip")
