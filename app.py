import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import os, subprocess, tempfile, zipfile, unicodedata
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"

if 'raw_records' not in st.session_state: st.session_state.raw_records = []
if 'map_memoria' not in st.session_state: st.session_state.map_memoria = {}

def proper_elegante(texto):
    if not texto or str(texto).lower() == "none": return ""
    texto = ''.join(c for c in unicodedata.normalize('NFD', str(texto))
                  if unicodedata.category(c) != 'Mn')
    palabras = texto.lower().split()
    excepciones = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del']
    resultado = []
    for i, p in enumerate(palabras):
        if i > 0 and p in excepciones: resultado.append(p)
        else: resultado.append(p.capitalize())
    return " ".join(resultado)

def formatear_fecha_mx(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        dias = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]
        res = f"{dias[dt.weekday()]} {dt.day} de {meses[dt.month-1]} de {dt.year}"
        return proper_elegante(res)
    except: return ""

def formatear_hora_mx(hora_raw):
    if not hora_raw: return ""
    try:
        t = datetime.strptime(str(hora_raw).strip(), "%H:%M")
        return t.strftime("%I:%M %p").lower().replace("am", "a.m.").replace("pm", "p.m.")
    except: return str(hora_raw)

def generar_pdf(pptx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        path = tmp.name
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(path), path], check=True)
        pdf_path = path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f: data = f.read()
        os.remove(path)
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return data
    except: return None

# --- APP ---
st.set_page_config(page_title="Provident Pro", layout="wide")
st.title("🚀 Generador Pro: Concat 4 Líneas y Máximo Estilo")

with st.sidebar:
    headers = {"Authorization": f"Bearer {TOKEN}"}
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base:", list(base_opts.keys()))
        r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
        if r_tab.status_code == 200:
            tabla_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
            tabla_sel = st.selectbox("Tabla:", list(tabla_opts.keys()))
            if st.button("🔄 Cargar Registros"):
                r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
                st.session_state.raw_records = r_reg.json().get("records", [])
                st.rerun()

if st.session_state.raw_records:
    data_list = [{"Tipo": proper_elegante(r['fields'].get("Tipo")), "Sucursal": proper_elegante(r['fields'].get("Sucursal")), "Municipio": proper_elegante(r['fields'].get("Municipio")), "Fecha": r['fields'].get("Fecha", "")} for r in st.session_state.raw_records]
    df_display = pd.DataFrame(data_list)
    df_display.insert(0, "Seleccionar", False)
    df_edit = st.data_editor(df_display, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        st.divider()
        opcion_salida = st.radio("Tipo de salida:", ["Postales", "Reportes"], horizontal=True)
        folder_fisica = os.path.join(BASE_DIR, opcion_salida.upper())
        
        if os.path.exists(folder_fisica):
            archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
            tipos_unicos = df_edit.loc[sel_idx, "Tipo"].unique()
            
            st.write("### 📂 Selección de Plantillas")
            cols = st.columns(len(tipos_unicos))
            for i, t_p in enumerate(tipos_unicos):
                with cols[i]:
                    idx_mem = 0
                    if t_p in st.session_state.map_memoria:
                        if st.session_state.map_memoria[t_p] in archivos_pptx:
                            idx_mem = archivos_pptx.index(st.session_state.map_memoria[t_p])
                    st.session_state.map_memoria[t_p] = st.selectbox(f"Archivo para {t_p}:", archivos_pptx, index=idx_mem, key=f"s_{t_p}")

            if st.button("🔥 GENERAR ARCHIVOS"):
                progreso_bar = st.progress(0)
                status_text = st.empty()
                total = len(sel_idx)
                
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                    for i, idx in enumerate(sel_idx):
                        rec = st.session_state.raw_records[idx]['fields']
                        suc_raw = rec.get('Sucursal')
                        
                        progreso_bar.progress((i + 1) / total)
                        status_text.info(f"🔨 Procesando: **{proper_elegante(suc_raw)}**")

                        # --- LOGICA DE CONCAT A 4 LINEAS ---
                        punto = str(rec.get('Punto de reunion') or '').strip()
                        ruta = str(rec.get('Ruta a seguir') or '').strip()
                        muni = str(rec.get('Municipio') or '').strip()
                        secc = str(rec.get('Seccion') or '').strip()

                        # Forzamos los saltos de línea para asegurar las 4 líneas
                        concat_multiline = (
                            f"{proper_elegante(punto)}\n"
                            f"{proper_elegante(ruta)}\n"
                            f"Municipio {proper_elegante(muni)}\n"
                            f"Seccion {proper_elegante(secc)}"
                        )

                        reemplazos = {
                            "<<Consuc>>": proper_elegante(f"Sucursal {suc_raw}" + (f", {muni}" if muni else "")),
                            "<<Confecha>>": formatear_fecha_mx(rec.get('Fecha')),
                            "<<Conhora>>": formatear_hora_mx(rec.get('Hora')),
                            "<<Concat>>": concat_multiline,
                            "<<Sucursal>>": proper_elegante(suc_raw)
                        }

                        # --- PPTX ---
                        prs = Presentation(os.path.join(folder_fisica, st.session_state.map_memoria[proper_elegante(rec.get('Tipo'))]))
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    shape.text_frame.word_wrap = True
                                    for paragraph in shape.text_frame.paragraphs:
                                        full_p = "".join(r.text for r in paragraph.runs)
                                        for tag, val in reemplazos.items():
                                            if tag in full_p:
                                                new_txt = full_p.replace(tag, val)
                                                # Escalado dinámico desde 42pt
                                                if tag == "<<Conhora>>": size = 42
                                                elif tag == "<<Concat>>": size = 32 if len(new_txt) < 80 else 24
                                                else: size = 42 if len(new_txt) < 25 else 32
                                                
                                                run = paragraph.runs[0]
                                                run.text = new_txt
                                                run.font.size = Pt(size)
                                                for r_idx in range(1, len(paragraph.runs)): paragraph.runs[r_idx].text = ""

                        # --- FINALIZACIÓN ---
                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        if pdf_data:
                            dt = datetime.strptime(str(rec.get('Fecha', '2024-01-01')), '%Y-%m-%d')
                            nombre = proper_elegante(f"{dt.day} de {dt.month} {rec.get('Tipo')} {suc_raw}") + (".png" if opcion_salida == "Postales" else ".pdf")
                            ruta_zip = f"Provident/{dt.year}/{proper_elegante(dt.strftime('%m - %B'))}/{opcion_salida}/{proper_elegante(suc_raw)}/{nombre}"
                            
                            if opcion_salida == "Reportes":
                                zip_f.writestr(ruta_zip, pdf_data)
                            else:
                                imgs = convert_from_bytes(pdf_data)
                                if imgs:
                                    img_io = BytesIO()
                                    imgs[0].save(img_io, format='PNG')
                                    zip_f.writestr(ruta_zip, img_io.getvalue())

                status_text.success("✅ Generación finalizada exitosamente.")
                st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), f"Provident_Final.zip")
