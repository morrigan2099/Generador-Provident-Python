import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from pptx.util import Pt, Inches
import os, subprocess, tempfile, zipfile, unicodedata
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN DE MESES Y TOKEN ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"
MESES_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

if 'raw_records' not in st.session_state: st.session_state.raw_records = []
if 'map_memoria' not in st.session_state: st.session_state.map_memoria = {}

# --- FUNCIONES DE FORMATO ---
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
        return f"{dt.day:02d} de {MESES_ES[dt.month-1]} de {dt.year}"
    except: return ""

def formatear_confechor(fecha_str, hora_raw):
    try:
        dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        hora = formatear_hora_mx(hora_raw)
        # Línea 1: dd de Mes | Línea 2: de aaaa, hora
        linea1 = f"{dt.day:02d} de {MESES_ES[dt.month-1]}"
        linea2 = f"de {dt.year}, {hora}"
        return f"{linea1}\n{linea2}"
    except: return ""

def formatear_hora_mx(hora_raw):
    if not hora_raw: return ""
    try:
        # Intenta parsear la hora (asumiendo formato HH:MM desde Airtable)
        t = datetime.strptime(str(hora_raw).strip(), "%H:%M")
        return t.strftime("%I:%M %p").lower().replace("am", "a.m.").replace("pm", "p.m.")
    except:
        return str(hora_raw).lower()

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

def descargar_imagen(url):
    try:
        r = requests.get(url, timeout=10)
        if r.status_code == 200: return BytesIO(r.content)
    except: return None
    return None

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Provident Pro Ultra", layout="wide")
st.title("🚀 Generador Provident Pro")

with st.sidebar:
    st.header("Conexión Airtable")
    headers = {"Authorization": f"Bearer {TOKEN}"}
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base:", list(base_opts.keys()))
        r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
        if r_tab.status_code == 200:
            tabla_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
            tabla_sel = st.selectbox("Tabla:", list(tabla_opts.keys()))
            if st.button("🔄 Cargar Datos"):
                r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
                st.session_state.raw_records = r_reg.json().get("records", [])
                st.success(f"Cargados {len(st.session_state.raw_records)} registros")

if st.session_state.raw_records:
    df_prev = pd.DataFrame([{"Tipo": r['fields'].get("Tipo"), "Sucursal": r['fields'].get("Sucursal"), "Fecha": r['fields'].get("Fecha")} for r in st.session_state.raw_records])
    df_prev.insert(0, "Seleccionar", False)
    df_edit = st.data_editor(df_prev, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        modo = st.radio("Formato de salida:", ["Postales", "Reportes"], horizontal=True)
        folder_fisica = os.path.join(BASE_DIR, modo.upper())
        
        if os.path.exists(folder_fisica):
            archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
            tipos_unicos = df_edit.loc[sel_idx, "Tipo"].unique()
            cols = st.columns(len(tipos_unicos))
            for i, t in enumerate(tipos_unicos):
                with cols[i]:
                    st.session_state.map_memoria[t] = st.selectbox(f"Plantilla {t}:", archivos_pptx, key=f"p_{t}")

            if st.button("🔥 GENERAR ARCHIVOS"):
                progreso = st.progress(0)
                status = st.empty()
                zip_buf = BytesIO()
                
                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                    for i, idx in enumerate(sel_idx):
                        f = st.session_state.raw_records[idx]['fields']
                        suc_p = proper_elegante(f.get('Sucursal'))
                        status.info(f"Procesando: {suc_p}")

                        # Textos para placeholders
                        punto = str(f.get('Punto de reunion') or '').strip()
                        ruta = str(f.get('Ruta a seguir') or '').strip()
                        muni = str(f.get('Municipio') or '').strip()
                        secc = str(f.get('Seccion') or '').strip()
                        
                        concat_txt = ", ".join([p for p in [punto, ruta] if p])
                        if muni: concat_txt += f", Municipio {muni}"
                        if secc: concat_txt += f", Seccion {secc}"

                        reemplazos = {
                            "<<Tipo>>": str(f.get('Tipo')).upper(),
                            "<<confechor>>": formatear_confechor(f.get('Fecha'), f.get('Hora')),
                            "<<Consuc>>": proper_elegante(f"Sucursal {f.get('Sucursal')}" + (f", {muni}" if muni else "")),
                            "<<Confecha>>": proper_elegante(formatear_fecha_mx(f.get('Fecha'))),
                            "<<Conhora>>": formatear_hora_mx(f.get('Hora')),
                            "<<Concat>>": proper_elegante(concat_txt),
                            "<<Sucursal>>": proper_elegante(f.get('Sucursal'))
                        }

                        prs = Presentation(os.path.join(folder_fisica, st.session_state.map_memoria[f.get('Tipo')]))
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    shape.text_frame.word_wrap = True
                                    for paragraph in shape.text_frame.paragraphs:
                                        full_p = "".join(r.text for r in paragraph.runs)
                                        for tag, val in reemplazos.items():
                                            if tag in full_p:
                                                new_txt = full_p.replace(tag, val)
                                                # Tamaño de fuente dinámico
                                                if tag == "<<confechor>>": fs = 32
                                                elif tag == "<<Concat>>": fs = 42 if len(new_txt) < 60 else 32
                                                else: fs = 42 if len(new_txt) < 30 else 36
                                                
                                                run = paragraph.runs[0]
                                                run.text = new_txt
                                                run.font.size = Pt(fs)
                                                for r_idx in range(1, len(paragraph.runs)):
                                                    paragraph.runs[r_idx].text = ""

                                if modo == "Reportes":
                                    tags_fotos = ["<<Foto de equipo>>", "<<Foto 01>>", "<<Foto 02>>", "<<Foto 03>>", 
                                                  "<<Foto 04>>", "<<Foto 05>>", "<<Foto 06>>", "<<Foto 07>>",
                                                  "<<Reporte firmado>>", "<<Lista de asistencia>>"]
                                    for t_foto in tags_fotos:
                                        if (shape.has_text_frame and t_foto in shape.text) or (t_foto in shape.name):
                                            campo_airtable = t_foto.replace("<<", "").replace(">>", "")
                                            adjuntos = f.get(campo_airtable)
                                            if adjuntos and isinstance(adjuntos, list):
                                                img_data = descargar_imagen(adjuntos[0].get('url'))
                                                if img_data:
                                                    slide.shapes.add_picture(img_data, shape.left, shape.top, shape.width, shape.height)

                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        if pdf_data:
                            dt = datetime.strptime(f.get('Fecha'), '%Y-%m-%d')
                            # Nomenclatura archivo
                            partes_n = [p for p in [punto, ruta] if p]
                            if muni: partes_n.append(muni)
                            nom_f = f"{formatear_fecha_mx(f.get('Fecha'))} - {str(f.get('Tipo')).upper()} {str(f.get('Sucursal')).upper()} - {', '.join(partes_n)}"
                            
                            # Recorte si el nombre es muy largo
                            if len(nom_f) > 150:
                                res_n = [ruta] if (punto and ruta and len(punto)<len(ruta)) else ([punto] if punto else [])
                                if muni: res_n.append(muni)
                                nom_f = f"{formatear_fecha_mx(f.get('Fecha'))} - {str(f.get('Tipo')).upper()} {str(f.get('Sucursal')).upper()} - {', '.join(res_n)}"
                            
                            nombre_final = proper_elegante(nom_f) + (".jpg" if modo == "Postales" else ".pdf")
                            folder_mes = f"{dt.month:02d} - {MESES_ES[dt.month-1]}"
                            ruta_zip = f"Provident/{dt.year}/{folder_mes}/{modo}/{proper_elegante(f.get('Sucursal'))}/{nombre_final}"
                            
                            if modo == "Reportes":
                                zip_f.writestr(ruta_zip, pdf_data)
                            else:
                                imgs = convert_from_bytes(pdf_data)
                                if imgs:
                                    img_io = BytesIO()
                                    imgs[0].convert('RGB').save(img_io, format='JPEG', quality=85, optimize=True)
                                    zip_f.writestr(ruta_zip, img_io.getvalue())
                        progreso.progress((i + 1) / len(sel_idx))
                
                st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_Pro_Final.zip")
