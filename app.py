import streamlit as st
import requests
import pandas as pd
import json
import os
import re
import numpy as np
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import subprocess, tempfile, zipfile
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes
from PIL import Image, ImageOps, ImageFilter, ImageChops

# --- CONFIGURACIÓN DE TAMAÑOS ESTÁTICOS (FIJOS) ---
TAM_TIPO      = 64  # SIEMPRE 64pts como solicitaste
TAM_SUCURSAL  = 11
TAM_SECCION   = 11
TAM_CONFECHOR = 11
TAM_CONCAT    = 11

# --- CONSTANTES ---
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]

# --- FUNCIONES DE IMAGEN ---

def recorte_inteligente_bordes(img, umbral_negro=60, margen_analisis=0.05):
    """Analiza si los bordes tienen más del 60% de negro y los recorta."""
    w, h = img.size
    img_gray = img.convert('L') 
    arr = np.array(img_gray)
    
    top, bottom, left, right = 0, h, 0, w
    v_strip = int(h * margen_analisis)
    h_strip = int(w * margen_analisis)

    def es_franja_negra(seccion):
        if seccion.size == 0: return False
        pue_negros = np.sum(seccion < 35) / seccion.size * 100
        return pue_negros > umbral_negro

    if es_franja_negra(arr[0:v_strip, :]): top = v_strip
    if es_franja_negra(arr[h-v_strip:h, :]): bottom = h - v_strip
    if es_franja_negra(arr[:, 0:h_strip]): left = h_strip
    if es_franja_negra(arr[:, w-h_strip:w]): right = w - h_strip

    return img.crop((left, top, right, bottom))

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
        # Estirado directo (Páginas 3 y 4)
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
        else: resultado.append(p)
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

# --- UI STREAMLIT ---
st.set_page_config(page_title="Provident Pro v67", layout="wide")
if 'config' not in st.session_state:
    if os.path.exists("config_app.json"):
        with open("config_app.json", "r") as f: st.session_state.config = json.load(f)
    else: st.session_state.config = {"plantillas": {}}

st.title("🚀 Generador Pro v67 - Estático 64pts")

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# ... (Sidebar y carga de datos igual a v65) ...

if 'raw_records' in st.session_state:
    modo = st.radio("Salida:", ["Postales", "Reportes"], horizontal=True)
    # ... (Editor de datos y selección) ...

    if st.button("🔥 GENERAR", use_container_width=True, type="primary"):
        p_bar = st.progress(0)
        status_text = st.empty()
        zip_buf = BytesIO()
        AZUL_CELESTE = RGBColor(0, 176, 240)
        mapa_tamanos = {"<<Tipo>>": TAM_TIPO, "<<Sucursal>>": TAM_SUCURSAL, "<<Seccion>>": TAM_SECCION, "<<Confechor>>": TAM_CONFECHOR, "<<Concat>>": TAM_CONCAT}

        with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
            for i, idx in enumerate(sel_idx):
                record = st.session_state.raw_records[idx]['fields']
                record_orig = st.session_state.raw_data_original[idx]['fields']
                
                # Manejo seguro de fecha para evitar IndexError
                try:
                    raw_fecha = record.get('Fecha', '2025-01-01')
                    dt = datetime.strptime(raw_fecha, '%Y-%m-%d')
                    idx_mes = max(0, min(11, dt.month - 1))
                    nombre_mes = MESES_ES[idx_mes]
                except:
                    dt = datetime.now()
                    nombre_mes = "enero"

                f_tipo = record.get('Tipo', 'Sin Tipo')
                f_suc = record.get('Sucursal', '000')
                lugar = record.get('Punto de reunion') or record.get('Ruta a seguir')
                f_concat = f"Sucursal {f_suc}" if f_tipo == "Actividad en Sucursal" else ", ".join([str(x) for x in [lugar, record.get('Municipio')] if x and str(x).lower() != 'none'])
                
                nom_arch = f"{dt.day} de {nombre_mes} de {dt.year} - {f_tipo}, {f_suc}" + ("" if f_tipo == "Actividad en Sucursal" else f" - {f_concat}")
                status_text.markdown(f"**Procesando:** `{nom_arch}`")
                
                reemplazos = {"<<Tipo>>": f_tipo, "<<Sucursal>>": f_suc, "<<Seccion>>": record.get('Seccion'), 
                              "<<Confechor>>": f"{DIAS_ES[dt.weekday()]} {dt.day} de {nombre_mes} de {dt.year}, {record.get('Hora', '').lower()}", 
                              "<<Concat>>": f_concat}

                prs = Presentation(os.path.join(folder_fisica, st.session_state.config["plantillas"][f_tipo]))
                
                # ... (Lógica de eliminación de diapositiva 4 si aplica) ...

                for s_idx, slide in enumerate(prs.slides):
                    for shape in list(slide.shapes):
                        if shape.has_text_frame:
                            # Reemplazo de imágenes con Auto-Recorte
                            tags_foto = ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]
                            for tf in tags_foto:
                                if f"<<{tf}>>" in shape.text_frame.text:
                                    adj_data = record_orig.get(tf)
                                    thumbnails = adj_data[0].get('thumbnails', {}) if adj_data else {}
                                    url_hd = thumbnails.get('full', {}).get('url') or thumbnails.get('large', {}).get('url') or (adj_data[0].get('url') if adj_data else None)
                                    if url_hd:
                                        try:
                                            r_img_bytes = requests.get(url_hd).content
                                            img_io = procesar_imagen_inteligente(r_img_bytes, shape.width, shape.height, con_blur=(s_idx < 2))
                                            slide.shapes.add_picture(img_io, shape.left, shape.top, shape.width, shape.height)
                                            sp = shape._element; sp.getparent().remove(sp)
                                        except: pass

                    # Reemplazo de textos con tamaños estáticos
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for tag, val in reemplazos.items():
                                if tag in shape.text_frame.text:
                                    tf = shape.text_frame; tf.clear()
                                    run = tf.paragraphs[0].add_run()
                                    run.text = str(val); run.font.bold = True; run.font.color.rgb = AZUL_CELESTE
                                    run.font.size = Pt(mapa_tamanos.get(tag, 11))

                # Guardado y PDF
                pp_io = BytesIO(); prs.save(pp_io)
                data_out = generar_pdf(pp_io.getvalue())
                if data_out:
                    ext = ".pdf" if modo == "Reportes" else ".jpg"
                    ruta_zip = f"Provident/{dt.year}/{str(dt.month).zfill(2)} - {nombre_mes}/{modo}/{f_suc}/{nom_arch[:140]}{ext}"
                    zip_f.writestr(ruta_zip, data_out if modo == "Reportes" else convert_from_bytes(data_out)[0].tobytes())
                p_bar.progress((i + 1) / len(sel_idx))
        
        status_text.success("✅ Generación completa.")
        st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), f"Provident_{datetime.now().strftime('%H%M%S')}.zip", use_container_width=True)
