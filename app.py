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
    """Minúsculas, sin acentos, sin saltos, sin dobles espacios, luego Proper Case"""
    if not texto or str(texto).lower() == "none":
        return ""
    # 1. Quitar saltos de línea y dobles espacios
    texto = str(texto).replace('\n', ' ').replace('\r', ' ')
    texto = ' '.join(texto.split())
    # 2. Quitar acentos y minúsculas
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto)
                  if unicodedata.category(c) != 'Mn').lower()
    # 3. Proper Case (Elegante)
    return texto.title()

def formatear_fecha_es(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        dias = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]
        # Formato solicitado: EEEE dd de mmmm de aaaa (sin acentos por procesar_texto)
        res = f"{dias[dt.weekday()]} {dt.day} de {meses[dt.month-1]} de {dt.year}"
        return procesar_texto_elegante(res)
    except: return ""

def formatear_hora_es(hora_raw):
    if not hora_raw: return ""
    try:
        # Asumiendo entrada HH:MM
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

# --- APP PRINCIPAL ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")
st.title("🚀 Sistema de Generación Masiva")

# 1. AIRTABLE (BARRA LATERAL)
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
            st.session_state.raw_data = r_reg.json().get("records", [])

# 2. PROCESAMIENTO
if 'raw_data' in st.session_state and st.session_state.raw_data:
    df = pd.DataFrame([r['fields'] for r in st.session_state.raw_data])
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
                    # Extraer datos reales del registro de Airtable
                    rec = st.session_state.raw_data[idx]['fields']
                    
                    with st.status(f"Procesando: {rec.get('Sucursal')}"):
                        # --- LÓGICA DE CONTENIDO ---
                        punto = str(rec.get('Punto de reunion') or '').strip()
                        ruta = str(rec.get('Ruta a seguir') or '').strip()
                        muni = str(rec.get('Municipio') or '').strip()
                        secc = str(rec.get('Seccion') or '').strip()
                        
                        # Concat inteligente
                        c_partes = [p for p in [punto, ruta] if p]
                        concat_txt = ", ".join(c_partes)
                        if muni: concat_txt += f", Municipio {muni}"
                        if secc: concat_txt += f", Seccion {secc}"

                        tags = {
                            "<<Sucursal>>": procesar_texto_elegante(rec.get('Sucursal')),
                            "<<Confecha>>": formatear_fecha_es(rec.get('Fecha')),
                            "<<Conhora>>": formatear_hora_es(rec.get('Hora')),
                            "<<Concat>>": procesar_texto_elegante(concat_txt)
                        }

                        # --- PPTX ---
                        prs = Presentation(os.path.join(folder, mapping[rec.get('Tipo')]))
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    for p in shape.text_frame.paragraphs:
                                        for run in p.runs:
                                            for tag, val in tags.items():
                                                if tag in run.text:
                                                    run.text = run.text.replace(tag, val)
                        
                        # --- CONVERSIÓN ---
                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        
                        if pdf_data:
                            # --- NOMBRE DE ARCHIVO Y CARPETAS ---
                            try: dt = datetime.strptime(rec.get('Fecha'), '%Y-%m-%d')
                            except: dt = datetime.now()
                            
                            ext = "png" if uso_final == "POSTALES" else "pdf"
                            
                            # Nombre de archivo dinámico
                            nom_f = f"{dt.strftime('%d de %m de %Y')} - {rec.get('Tipo')} {rec.get('Sucursal')} - {concat_txt}"
                            if len(nom_f) > 150: # Si es muy largo, recortamos el concat
                                nom_f = f"{dt.strftime('%d de %m de %Y')} - {rec.get('Tipo')} {rec.get('Sucursal')} - {punto if len(punto)<len(ruta) else ruta}"
                            
                            nombre_limpio = f"{procesar_texto_elegante(nom_f)}.{ext}"
                            
                            # Ruta interna ZIP: Provident/Año/Mes/Uso/Sucursal/
                            uso_sub = "Postales" if uso_final == "POSTALES" else "Reportes"
                            mes_f = dt.strftime('%m - %B').lower()
                            ruta_zip = f"Provident/{dt.year}/{procesar_texto_elegante(mes_f)}/{uso_sub}/{procesar_texto_elegante(rec.get('Sucursal'))}/{nombre_limpio}"
                            
                            if uso_final == "REPORTES":
                                zip_f.writestr(ruta_zip, pdf_data)
                            else:
                                png = generar_png(pdf_data)
                                if png: zip_f.writestr(ruta_zip, png)

            st.success("✅ ZIP generado correctamente.")
            st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_Estructurado.zip", "application/zip")
