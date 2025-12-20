import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
import os, subprocess, tempfile, zipfile, unicodedata, locale
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes

# Intentar establecer locale a MX para fechas
try: locale.setlocale(locale.LC_TIME, "es_MX.utf8")
except: pass

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"

# --- LÓGICA PROPER ELEGANTE ---
def proper_elegante(texto):
    if not texto: return ""
    # Quitar acentos
    texto = ''.join(c for c in unicodedata.normalize('NFD', str(texto))
                  if unicodedata.category(c) != 'Mn')
    palabras = texto.lower().split()
    excepciones = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del']
    resultado = []
    for i, p in enumerate(palabras):
        if i > 0 and p in excepciones:
            resultado.append(p)
        else:
            resultado.append(p.capitalize())
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

# --- MOTORES ---
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

# --- APP ---
st.title("🚀 Generador Pro: Estilo y Formato Preservado")

# (Carga de datos de Airtable omitida para brevedad, misma lógica que ya funciona)
# ... [Bloque Sidebar y Data Editor] ...

if 'raw_records' in st.session_state and st.session_state.raw_records:
    # ... [Preparación de DataFrame y Selección] ...
    
    if sel_idx:
        uso_final = st.radio("Uso:", ["POSTALES", "REPORTES"], horizontal=True)
        folder = os.path.join(BASE_DIR, uso_final)
        archivos_pptx = [f for f in os.listdir(folder) if f.endswith('.pptx')]
        tipos_seleccionados = df_edit.loc[sel_idx, "Tipo"].unique()
        
        # EL MAPEO QUE NO SE PIERDE
        mapping = {t: st.selectbox(f"Plantilla para {t}", archivos_pptx, key=f"m_{t}") for t in tipos_seleccionados}

        if st.button("🔥 EJECUTAR GENERACIÓN MASIVA"):
            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                for idx in sel_idx:
                    rec = st.session_state.raw_records[idx]['fields']
                    tipo_reg = rec.get('Tipo', '')
                    
                    # 1. DATOS Y CONCATENACIONES
                    suc = str(rec.get('Sucursal') or '').strip()
                    muni = str(rec.get('Municipio') or '').strip()
                    # Consuc: "Sucursal" Nombre, Municipio
                    consuc_val = f"Sucursal {suc}" + (f", {muni}" if muni else "")
                    
                    # Concat
                    c_partes = [str(rec.get(k) or '').strip() for k in ['Punto de reunion', 'Ruta a seguir'] if rec.get(k)]
                    concat_txt = ", ".join(c_partes)
                    if muni: concat_txt += f", Municipio {muni}"
                    if rec.get('Seccion'): concat_txt += f", Seccion {rec.get('Seccion')}"

                    reemplazos = {
                        "<<Consuc>>": proper_elegante(consuc_val),
                        "<<Confecha>>": formatear_fecha_mx(rec.get('Fecha')),
                        "<<Conhora>>": formatear_hora_mx(rec.get('Hora')),
                        "<<Concat>>": proper_elegante(concat_txt),
                        "<<Sucursal>>": proper_elegante(suc)
                    }

                    # 2. PROCESO DE PPTX PRESERVANDO FORMATO
                    prs = Presentation(os.path.join(folder, mapping[tipo_reg]))
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for paragraph in shape.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        for tag, val in reemplazos.items():
                                            if tag in run.text:
                                                # CAMBIO DE TEXTO SIN TOCAR COLOR/FUENTE
                                                run.text = run.text.replace(tag, val)
                                                # Ajuste preventivo solo para Concat si es muy largo
                                                if tag == "<<Concat>>" and len(val) > 60:
                                                    run.font.size = Pt(12)

                    # 3. GUARDADO Y ESTRUCTURA DE CARPETAS
                    pp_io = BytesIO(); prs.save(pp_io)
                    pdf_data = generar_pdf(pp_io.getvalue())
                    
                    if pdf_data:
                        f_raw = rec.get('Fecha', '2024-01-01')
                        dt = datetime.strptime(f_raw, '%Y-%m-%d')
                        mes_nombre = proper_elegante(dt.strftime('%m - %B'))
                        
                        ext = "png" if uso_final == "POSTALES" else "pdf"
                        nombre_file = proper_elegante(f"{dt.day} de {dt.month} {tipo_reg} {suc}") + f".{ext}"
                        
                        path_zip = (f"Provident/{dt.year}/{mes_nombre}/"
                                   f"{proper_elegante(uso_final)}/{proper_elegante(suc)}/{nombre_file}")
                        
                        if uso_final == "REPORTES":
                            zip_f.writestr(path_zip, pdf_data)
                        else:
                            img = generar_png(pdf_data)
                            if img: zip_f.writestr(path_zip, img)

            st.success("✅ Generación exitosa con formato Proper Elegante.")
            st.download_button("📥 DESCARGAR ZIP FINAL", zip_buf.getvalue(), "Provident_Elegante.zip")
