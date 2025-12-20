import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from io import BytesIO
import os
import subprocess
import tempfile
import zipfile
import unicodedata
from datetime import datetime
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"

# --- TRANSFORMACIÓN DE TEXTO ---
def procesar_texto_elegante(texto):
    if not texto or str(texto).lower() == "none":
        return ""
    texto = str(texto).replace('\n', ' ').replace('\r', ' ')
    texto = ' '.join(texto.split())
    # Quitar acentos
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto)
                  if unicodedata.category(c) != 'Mn').lower()
    return texto.title()

def formatear_fecha_es(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        dias = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]
        # Formato: jueves 20 de mayo de 2024 (a dos líneas se maneja en el diseño del PPTX)
        res = f"{dias[dt.weekday()]} {dt.day} de {meses[dt.month-1]} de {dt.year}"
        return procesar_texto_elegante(res)
    except: return ""

def formatear_hora_es(hora_raw):
    if not hora_raw: return ""
    try:
        # Intenta convertir formato HH:MM a 12h con a.m./p.m.
        t = datetime.strptime(str(hora_raw).strip(), "%H:%M")
        formato = t.strftime("%I:%M %p").lower()
        return formato.replace("am", "a.m.").replace("pm", "p.m.")
    except: return str(hora_raw)

# --- CONVERSIÓN ---
def generar_pdf(pptx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        tmp_path = tmp.name
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(tmp_path), tmp_path], check=True)
        pdf_path = tmp_path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f: data = f.read()
        os.remove(tmp_path)
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return data
    except: return None

def generar_png(pdf_bytes):
    try:
        images = convert_from_bytes(pdf_bytes)
        if images:
            buf = BytesIO()
            images[0].save(buf, format='PNG')
            return buf.getvalue()
    except: return None

# --- APP ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")
st.title("🚀 Generador Masivo (Fix: Etiquetas Invisibles)")

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
            r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
            st.session_state.raw_records = r_reg.json().get("records", [])

if 'raw_records' in st.session_state and st.session_state.raw_records:
    df = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    for c in ["Tipo", "Sucursal", "Fecha", "Hora", "Punto de reunion", "Ruta a seguir", "Municipio", "Seccion"]:
        if c not in df.columns: df[c] = ""
    
    df.insert(0, "Seleccionar", False)
    df_edit = st.data_editor(df, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        uso_final = st.radio("Formato de Salida:", ["POSTALES", "REPORTES"], horizontal=True)
        folder = os.path.join(BASE_DIR, uso_final)
        archivos_pptx = [f for f in os.listdir(folder) if f.endswith('.pptx')]
        tipos = df_edit.loc[sel_idx, "Tipo"].unique()
        mapping = {t: st.selectbox(f"Plantilla para {t}:", archivos_pptx, key=t) for t in tipos}

        if st.button("🔥 GENERAR ZIP ESTRUCTURADO"):
            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                for idx in sel_idx:
                    rec = st.session_state.raw_records[idx]['fields']
                    
                    with st.status(f"Procesando: {rec.get('Sucursal')}"):
                        # 1. LÓGICA DE CONCATENACIÓN (PLACEHOLDERS)
                        punto = str(rec.get('Punto de reunion') or '').strip()
                        ruta = str(rec.get('Ruta a seguir') or '').strip()
                        muni = str(rec.get('Municipio') or '').strip()
                        secc = str(rec.get('Seccion') or '').strip()
                        
                        # Concat inteligente sin comas huérfanas
                        c_partes = [p for p in [punto, ruta] if p]
                        centro = ", ".join(c_partes)
                        concat_txt = centro
                        if muni: 
                            concat_txt += f", Municipio {muni}" if concat_txt else f"Municipio {muni}"
                        if secc: 
                            concat_txt += f", Seccion {secc}" if concat_txt else f"Seccion {secc}"

                        reemplazos = {
                            "<<Sucursal>>": procesar_texto_elegante(rec.get('Sucursal')),
                            "<<Confecha>>": formatear_fecha_es(rec.get('Fecha')),
                            "<<Conhora>>": formatear_hora_es(rec.get('Hora')),
                            "<<Concat>>": procesar_texto_elegante(concat_txt)
                        }

                        # 2. REEMPLAZO AGRESIVO EN PPTX
                        prs = Presentation(os.path.join(folder, mapping[rec.get('Tipo')]))
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    for paragraph in shape.text_frame.paragraphs:
                                        # Leemos el texto completo del párrafo para evitar el error de etiquetas rotas
                                        full_text = "".join(run.text for run in paragraph.runs)
                                        for tag, val in reemplazos.items():
                                            if tag in full_text:
                                                # Si la etiqueta existe, reemplazamos en el primer run y vaciamos los demás
                                                full_text = full_text.replace(tag, val)
                                                paragraph.runs[0].text = full_text
                                                for r in range(1, len(paragraph.runs)):
                                                    paragraph.runs[r].text = ""

                        # 3. GUARDADO Y ZIP
                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        
                        if pdf_data:
                            try: dt = datetime.strptime(str(rec.get('Fecha')), '%Y-%m-%d')
                            except: dt = datetime.now()
                            
                            ext = "png" if uso_final == "POSTALES" else "pdf"
                            
                            # Nombre archivo elegante
                            nombre_arc = f"{dt.strftime('%d de %m de %Y')} - {rec.get('Tipo')} {rec.get('Sucursal')}"
                            nombre_arc = procesar_texto_elegante(nombre_arc) + f".{ext}"
                            
                            # Estructura de carpetas
                            mes_f = dt.strftime('%m - %B').lower()
                            sub_uso = "Postales" if uso_final == "POSTALES" else "Reportes"
                            ruta_zip = f"Provident/{dt.year}/{procesar_texto_elegante(mes_f)}/{sub_uso}/{procesar_texto_elegante(rec.get('Sucursal'))}/{nombre_arc}"
                            
                            if uso_final == "REPORTES":
                                zip_f.writestr(ruta_zip, pdf_data)
                            else:
                                png = generar_png(pdf_data)
                                if png: zip_f.writestr(ruta_zip, png)

            st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident.zip", "application/zip")
