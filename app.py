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

# ============================================================
#  CONFIGURACIÓN Y CONSTANTES GLOBALES
# ============================================================
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]

# ============================================================
#  FUNCIONES DE APOYO (REUTILIZABLES)
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
    nombre_dia = DIAS_ES[dia_idx]
    nombre_mes = MESES_ES[fecha_dt.month - 1]
    return f"{nombre_dia} {fecha_dt.day} de {nombre_mes} de {fecha_dt.year}"

def obtener_hora_texto(hora_str):
    if not hora_str or str(hora_str).lower() == "none": return ""
    s_raw = str(hora_str).lower().strip()
    es_pm = "p.m." in s_raw or "pm" in s_raw or "p. m." in s_raw
    es_am = "a.m." in s_raw or "am" in s_raw or "a. m." in s_raw
    match = re.search(r'(\d{1,2}):(\d{2})', s_raw)
    if match:
        h = int(match.group(1))
        m = match.group(2)
        if es_pm and h < 12: h += 12
        if es_am and h == 12: h = 0
        if h == 0 and es_pm: h = 12 
        if 8 <= h <= 11: sufijo = "de la mañana"
        elif h == 12: sufijo = "del día"
        elif 13 <= h <= 19: sufijo = "de la tarde"
        elif h >= 20: sufijo = "p.m."
        else: sufijo = "a.m."
        h_mostrar = h
        if h > 12: h_mostrar -= 12
        if h == 0: h_mostrar = 12
        return f"{h_mostrar}:{m} {sufijo}"
    return hora_str

def obtener_concat_texto(record):
    parts = []
    val_punto = record.get('Punto de reunion')
    if val_punto and str(val_punto).lower() != 'none': parts.append(str(val_punto))
    val_ruta = record.get('Ruta a seguir')
    if val_ruta and str(val_ruta).lower() != 'none': parts.append(str(val_ruta))
    val_muni = record.get('Municipio')
    if val_muni and str(val_muni).lower() != 'none': parts.append(f"Municipio {val_muni}")
    val_secc = record.get('Seccion')
    if val_secc and str(val_secc).lower() != 'none': parts.append(f"Sección {str(val_secc).upper()}")
    return ", ".join(parts)

# ============================================================
#  INICIO DE LA APLICACIÓN
# ============================================================

st.set_page_config(page_title="Provident Pro v87", layout="wide")

if 'config' not in st.session_state:
    if os.path.exists("config_app.json"):
        with open("config_app.json", "r") as f: st.session_state.config = json.load(f)
    else: st.session_state.config = {"plantillas": {}}

st.title("🚀 Generador Pro v87")
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# ------------------------------------------------------------
# SIDEBAR: SELECCIÓN GLOBAL
# ------------------------------------------------------------
with st.sidebar:
    st.header("1. Conexión Airtable")
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Selecciona Base:", [""] + list(base_opts.keys()))
        
        if base_sel:
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
            tabla_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
            tabla_sel = st.selectbox("Selecciona Tabla:", list(tabla_opts.keys()))
            
            if st.button("🔄 CARGAR DATOS", type="primary"):
                with st.spinner("Descargando datos..."):
                    r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
                    recs = r_reg.json().get("records", [])
                    st.session_state.raw_data_original = recs
                    st.session_state.raw_records = [
                        {'id': r['id'], 'fields': {k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v) for k, v in r['fields'].items()}}
                        for r in recs
                    ]
                st.success("Datos cargados")
                st.rerun()
    
    st.divider()
    
    # MENÚ DE MÓDULOS (Solo aparece si hay datos)
    modulo_seleccionado = None
    if 'raw_records' in st.session_state:
        st.header("2. Selecciona Módulo")
        modulo_seleccionado = st.radio(
            "Herramienta:",
            ["📮 Generar Postales", "📄 Generar Reportes", "📅 Calendario Visual"],
            index=0
        )
        
        st.divider()
        if st.button("💾 Guardar Configuración"):
            with open("config_app.json", "w") as f: json.dump(st.session_state.config, f)
            st.toast("Plantillas guardadas")

# ------------------------------------------------------------
# ÁREA PRINCIPAL
# ------------------------------------------------------------

if 'raw_records' not in st.session_state:
    st.info("👈 Por favor, selecciona una Base y Tabla en el menú lateral para comenzar.")
else:
    df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    AZUL_CELESTE = RGBColor(0, 176, 240)

    # ========================================================
    #  MÓDULO 1: GENERAR POSTALES
    # ========================================================
    if modulo_seleccionado == "📮 Generar Postales":
        st.subheader("📮 Generador de Postales (Imagen)")
        st.info("Configuración: Imágenes JPEG optimizadas (<1MB), Tamaños grandes para relleno, Formato Confechor unido.")

        # --- Selector de Registros ---
        df_view = df_full.copy()
        for c in df_view.columns:
            if isinstance(df_view[c].iloc[0], list): df_view.drop(c, axis=1, inplace=True)
        
        if 'select_all' not in st.session_state: st.session_state.select_all = False
        c1, c2, _ = st.columns([1, 1, 5])
        if c1.button("✅ Todo"): st.session_state.select_all = True; st.rerun()
        if c2.button("❌ Nada"): st.session_state.select_all = False; st.rerun()
        
        df_view.insert(0, "Seleccionar", st.session_state.select_all)
        df_edit = st.data_editor(df_view, use_container_width=True, hide_index=True)
        sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

        if sel_idx:
            # Configuración Plantillas Postales
            folder_fisica = os.path.join("Plantillas", "POSTALES")
            if not os.path.exists(folder_fisica): os.makedirs(folder_fisica)
            archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
            
            tipos_sel = df_view.loc[sel_idx, "Tipo"].unique()
            cols_p = st.columns(len(tipos_sel))
            for i, t in enumerate(tipos_sel):
                p_mem = st.session_state.config["plantillas"].get(t)
                idx_def = archivos_pptx.index(p_mem) if p_mem in archivos_pptx else 0
                st.session_state.config["plantillas"][t] = cols_p[i].selectbox(f"Plantilla '{t}':", archivos_pptx, index=idx_def, key=f"p_post_{t}")

            if st.button("🔥 CREAR POSTALES", type="primary"):
                p_bar = st.progress(0)
                st.session_state.archivos_en_memoria = []
                
                # Mapa de Tamaños POSTALES
                TAM_MAPA_POSTAL = {
                    "<<Tipo>>": 12,
                    "<<Sucursal>>": 44, # Tope para rellenar tipo "Boca del Río"
                    "<<Seccion>>": 12,
                    "<<Conhora>>": 32, 
                    "<<Concat>>": 32,
                    "<<Consuc>>": 32,
                    "<<Confechor>>": 28
                }

                for i, idx in enumerate(sel_idx):
                    record = st.session_state.raw_records[idx]['fields']
                    record_orig = st.session_state.raw_data_original[idx]['fields']
                    
                    try: dt = datetime.strptime(record.get('Fecha','2025-01-01'), '%Y-%m-%d'); nombre_mes = MESES_ES[dt.month - 1]
                    except: dt = datetime.now(); nombre_mes = "error"
                    
                    f_tipo = record.get('Tipo', 'Sin Tipo'); f_suc = record.get('Sucursal', '000')
                    
                    # Lógica Postal
                    txt_fecha = obtener_fecha_texto(dt)
                    txt_hora = obtener_hora_texto(record.get('Hora', ''))
                    f_confechor = f"{txt_fecha.strip()}\n{txt_hora.strip()}"
                    f_concat = f"Sucursal {f_suc}" if f_tipo == "Actividad en Sucursal" else obtener_concat_texto(record)
                    
                    f_filename_tag = f_concat
                    nom_arch = re.sub(r'[\\/*?:"<>|]', "", f"{dt.day} de {nombre_mes} de {dt.year} - {f_tipo}, {f_suc} - {f_filename_tag}")[:120] + ".png"
                    
                    reemplazos = {
                        "<<Tipo>>": textwrap.fill(f_tipo, width=35), "<<Sucursal>>": f_suc, 
                        "<<Seccion>>": record.get('Seccion'),
                        "<<Confechor>>": f_confechor, "<<Concat>>": f_concat, "<<Consuc>>": f_concat
                    }
                    
                    try:
                        prs = Presentation(os.path.join(folder_fisica, st.session_state.config["plantillas"][f_tipo]))
                    except:
                        st.error(f"Falta plantilla para {f_tipo}"); continue

                    # Eliminar slide lista si no hay lista
                    if f_tipo == "Actividad en Sucursal":
                        adj_lista = record_orig.get("Lista de asistencia")
                        if not adj_lista or len(adj_lista) == 0:
                            if len(prs.slides) >= 4: prs.slides._sldIdLst.remove(prs.slides._sldIdLst[3])

                    # Renderizado
                    for slide in prs.slides:
                        # Imagenes
                        for shape in list(slide.shapes):
                            if shape.has_text_frame:
                                for tf_tag in ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]:
                                    if f"<<{tf_tag}>>" in shape.text_frame.text:
                                        adj_data = record_orig.get(tf_tag)
                                        if adj_data:
                                            thumbs = adj_data[0].get('thumbnails', {})
                                            url = thumbs.get('full', {}).get('url') or thumbs.get('large', {}).get('url') or adj_data[0].get('url')
                                            try:
                                                r_img = requests.get(url).content
                                                img_io = procesar_imagen_inteligente(r_img, shape.width, shape.height, con_blur=True)
                                                slide.shapes.add_picture(img_io, shape.left, shape.top, shape.width, shape.height)
                                                sp = shape._element; sp.getparent().remove(sp)
                                            except: pass
                        # Textos
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for tag, val in reemplazos.items():
                                    if tag in shape.text_frame.text:
                                        if tag in ["<<Tipo>>", "<<Sucursal>>"]: shape.top = shape.top - Pt(2)
                                        tf = shape.text_frame; tf.word_wrap = True
                                        bodyPr = tf._element.bodyPr
                                        for child in ['spAutoFit', 'normAutofit', 'noAutofit']:
                                            if bodyPr.find(qn(f'a:{child}')) is not None: bodyPr.remove(bodyPr.find(qn(f'a:{child}')))
                                        
                                        if tag == "<<Tipo>>": bodyPr.append(tf._element.makeelement(qn('a:spAutoFit')))
                                        else: bodyPr.append(tf._element.makeelement(qn('a:normAutofit')))
                                        
                                        tf.clear(); p = tf.paragraphs[0]
                                        if tag in ["<<Confechor>>", "<<Consuc>>"]: p.alignment = PP_ALIGN.CENTER
                                        p.space_before = Pt(0); p.space_after = Pt(0); p.line_spacing = 1.0
                                        run = p.add_run(); run.text = str(val); run.font.bold = True; run.font.color.rgb = AZUL_CELESTE
                                        run.font.size = Pt(TAM_MAPA_POSTAL.get(tag, 12))

                    pp_io = BytesIO(); prs.save(pp_io)
                    data_out = generar_pdf(pp_io.getvalue())
                    if data_out:
                        images = convert_from_bytes(data_out, dpi=170, fmt='jpeg')
                        img_pil = images[0]
                        with BytesIO() as img_buf:
                            img_pil.save(img_buf, format="JPEG", quality=85, optimize=True, progressive=True)
                            bytes_finales = img_buf.getvalue()
                        
                        ruta = f"{dt.year}/{str(dt.month).zfill(2)} - {nombre_mes}/Postales/{f_suc}/{nom_arch}"
                        st.session_state.archivos_en_memoria.append({"Seleccionar":True, "Archivo":nom_arch, "RutaZip":ruta, "Datos":bytes_finales, "Sucursal":f_suc, "Tipo":f_tipo})
                    p_bar.progress((i+1)/len(sel_idx))
                
                st.success("✅ Postales Generadas")

    # ========================================================
    #  MÓDULO 2: GENERAR REPORTES
    # ========================================================
    elif modulo_seleccionado == "📄 Generar Reportes":
        st.subheader("📄 Generador de Reportes (PDF)")
        st.info("Configuración: Formato PDF multipágina, Tamaños estándar, Fecha y Hora separadas.")

        # --- Selector de Registros (Duplicado para independencia) ---
        df_view = df_full.copy()
        for c in df_view.columns:
            if isinstance(df_view[c].iloc[0], list): df_view.drop(c, axis=1, inplace=True)
        
        if 'select_all' not in st.session_state: st.session_state.select_all = False
        c1, c2, _ = st.columns([1, 1, 5])
        if c1.button("✅ Todo", key="rt"): st.session_state.select_all = True; st.rerun()
        if c2.button("❌ Nada", key="rn"): st.session_state.select_all = False; st.rerun()
        
        df_view.insert(0, "Seleccionar", st.session_state.select_all)
        df_edit = st.data_editor(df_view, use_container_width=True, hide_index=True, key="ed_rep")
        sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

        if sel_idx:
            folder_fisica = os.path.join("Plantillas", "REPORTES")
            if not os.path.exists(folder_fisica): os.makedirs(folder_fisica)
            archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
            
            tipos_sel = df_view.loc[sel_idx, "Tipo"].unique()
            cols_p = st.columns(len(tipos_sel))
            for i, t in enumerate(tipos_sel):
                p_mem = st.session_state.config["plantillas"].get(t)
                idx_def = archivos_pptx.index(p_mem) if p_mem in archivos_pptx else 0
                st.session_state.config["plantillas"][t] = cols_p[i].selectbox(f"Plantilla '{t}':", archivos_pptx, index=idx_def, key=f"p_rep_{t}")

            if st.button("🔥 CREAR REPORTES", type="primary"):
                p_bar = st.progress(0)
                st.session_state.archivos_en_memoria = []
                
                # Mapa de Tamaños REPORTES
                TAM_MAPA_REPORTE = {
                    "<<Tipo>>": 12, "<<Sucursal>>": 12,
                    "<<Seccion>>": 12, "<<Confecha>>": 24,
                    "<<Conhora>>": 15, "<<Consuc>>": 24
                }

                for i, idx in enumerate(sel_idx):
                    record = st.session_state.raw_records[idx]['fields']
                    record_orig = st.session_state.raw_data_original[idx]['fields']
                    try: dt = datetime.strptime(record.get('Fecha','2025-01-01'), '%Y-%m-%d'); nombre_mes = MESES_ES[dt.month - 1]
                    except: dt = datetime.now(); nombre_mes = "error"
                    
                    f_tipo = record.get('Tipo', 'Sin Tipo'); f_suc = record.get('Sucursal', '000')
                    
                    # Lógica Reporte
                    txt_fecha = obtener_fecha_texto(dt)
                    txt_hora = obtener_hora_texto(record.get('Hora', ''))
                    f_consuc = f"Sucursal {f_suc}" if f_tipo == "Actividad en Sucursal" else ""
                    tag_name = f_consuc if f_tipo == "Actividad en Sucursal" else obtener_concat_texto(record)
                    
                    nom_arch = re.sub(r'[\\/*?:"<>|]', "", f"{dt.day} de {nombre_mes} de {dt.year} - {f_tipo}, {f_suc} - {tag_name}")[:120] + ".pdf"
                    
                    reemplazos = {
                        "<<Tipo>>": textwrap.fill(f_tipo, width=35), "<<Sucursal>>": f_suc,
                        "<<Seccion>>": record.get('Seccion'), "<<Confecha>>": txt_fecha,
                        "<<Conhora>>": txt_hora, "<<Consuc>>": f_consuc
                    }

                    try: prs = Presentation(os.path.join(folder_fisica, st.session_state.config["plantillas"][f_tipo]))
                    except: continue

                    if f_tipo == "Actividad en Sucursal":
                        adj_lista = record_orig.get("Lista de asistencia")
                        if not adj_lista or len(adj_lista) == 0:
                            if len(prs.slides) >= 4: prs.slides._sldIdLst.remove(prs.slides._sldIdLst[3])

                    for slide in prs.slides:
                        # Imagenes
                        for shape in list(slide.shapes):
                            if shape.has_text_frame:
                                for tf_tag in ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]:
                                    if f"<<{tf_tag}>>" in shape.text_frame.text:
                                        adj_data = record_orig.get(tf_tag)
                                        if adj_data:
                                            thumbs = adj_data[0].get('thumbnails', {})
                                            url = thumbs.get('full', {}).get('url') or thumbs.get('large', {}).get('url') or adj_data[0].get('url')
                                            try:
                                                r_img = requests.get(url).content
                                                img_io = procesar_imagen_inteligente(r_img, shape.width, shape.height, con_blur=True)
                                                slide.shapes.add_picture(img_io, shape.left, shape.top, shape.width, shape.height)
                                                sp = shape._element; sp.getparent().remove(sp)
                                            except: pass
                        # Texto
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for tag, val in reemplazos.items():
                                    if tag in shape.text_frame.text:
                                        if tag in ["<<Tipo>>", "<<Sucursal>>"]: shape.top = shape.top - Pt(2)
                                        tf = shape.text_frame; tf.word_wrap = True
                                        bodyPr = tf._element.bodyPr
                                        for child in ['spAutoFit', 'normAutofit', 'noAutofit']:
                                            if bodyPr.find(qn(f'a:{child}')) is not None: bodyPr.remove(bodyPr.find(qn(f'a:{child}')))
                                        
                                        if tag == "<<Tipo>>": bodyPr.append(tf._element.makeelement(qn('a:spAutoFit')))
                                        elif tag == "<<Conhora>>": bodyPr.append(tf._element.makeelement(qn('a:spAutoFit')))
                                        else: bodyPr.append(tf._element.makeelement(qn('a:normAutofit')))
                                        
                                        tf.clear(); p = tf.paragraphs[0]
                                        if tag in ["<<Confecha>>", "<<Conhora>>", "<<Consuc>>"]: p.alignment = PP_ALIGN.CENTER
                                        p.space_before = Pt(0); p.space_after = Pt(0); p.line_spacing = 1.0
                                        run = p.add_run(); run.text = str(val); run.font.bold = True; run.font.color.rgb = AZUL_CELESTE
                                        run.font.size = Pt(TAM_MAPA_REPORTE.get(tag, 12))

                    pp_io = BytesIO(); prs.save(pp_io)
                    data_out = generar_pdf(pp_io.getvalue())
                    if data_out:
                        ruta = f"{dt.year}/{str(dt.month).zfill(2)} - {nombre_mes}/Reportes/{f_suc}/{nom_arch}"
                        st.session_state.archivos_en_memoria.append({"Seleccionar":True, "Archivo":nom_arch, "RutaZip":ruta, "Datos":data_out, "Sucursal":f_suc, "Tipo":f_tipo})
                    p_bar.progress((i+1)/len(sel_idx))
                
                st.success("✅ Reportes Generados")

    # ========================================================
    #  DESCARGAS (COMÚN PARA AMBOS MÓDULOS)
    # ========================================================
    if modulo_seleccionado in ["📮 Generar Postales", "📄 Generar Reportes"]:
        if "archivos_en_memoria" in st.session_state and len(st.session_state.archivos_en_memoria) > 0:
            st.divider()
            c1, c2, _ = st.columns([1, 1, 3])
            if c1.button("☑️ Marcar Todos"): 
                for i in st.session_state.archivos_en_memoria: i["Seleccionar"]=True
                st.rerun()
            if c2.button("⬜ Desmarcar Todos"): 
                for i in st.session_state.archivos_en_memoria: i["Seleccionar"]=False
                st.rerun()
            
            df_disp = pd.DataFrame(st.session_state.archivos_en_memoria)
            edited_df = st.data_editor(df_disp[["Seleccionar", "Archivo", "Sucursal", "Tipo"]], hide_index=True, use_container_width=True, column_config={"Seleccionar":st.column_config.CheckboxColumn(required=True)})
            
            indices = edited_df[edited_df["Seleccionar"]==True].index.tolist()
            finales = [st.session_state.archivos_en_memoria[i] for i in indices]
            
            if len(finales) > 0:
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                    for item in finales: zip_file.writestr(item["RutaZip"], item["Datos"])
                st.download_button(label=f"📦 DESCARGAR {len(finales)} ARCHIVOS (ZIP)", data=zip_buffer.getvalue(), file_name=f"Pack_{datetime.now().strftime('%H%M%S')}.zip", mime="application/zip", type="primary", use_container_width=True)

    # ========================================================
    #  MÓDULO 3: CALENDARIO VISUAL
    # ========================================================
    elif modulo_seleccionado == "📅 Calendario Visual":
        st.subheader("📅 Calendario de Actividades")
        
        fechas_ocupadas = {}
        for r in st.session_state.raw_data_original:
            f_str = r['fields'].get('Fecha')
            if f_str:
                if f_str not in fechas_ocupadas: fechas_ocupadas[f_str] = []
                thumb_url = None
                # REGLA ESTRICTA: SOLO CAMPO 'POSTAL'
                if 'Postal' in r['fields']:
                    att = r['fields']['Postal']
                    if isinstance(att, list) and len(att)>0: thumb_url = att[0].get('thumbnails', {}).get('small', {}).get('url')
                fechas_ocupadas[f_str].append({"id": r['id'], "thumb": thumb_url})

        if not fechas_ocupadas:
            st.warning("No se encontraron fechas en los datos cargados.")
        else:
            fechas_dt = [datetime.strptime(k, '%Y-%m-%d') for k in fechas_ocupadas.keys()]
            anios = sorted(list(set([d.year for d in fechas_dt])))
            anio_sel = st.selectbox("Año:", anios)
            
            meses_disp = sorted(list(set([d.month for d in fechas_dt if d.year == anio_sel])))
            mes_nombres = [MESES_ES[m-1].capitalize() for m in meses_disp]
            mes_sel_nom = st.selectbox("Mes:", mes_nombres)
            mes_sel_idx = MESES_ES.index(mes_sel_nom.lower()) + 1

            st.divider()
            
            cal = calendar.Calendar(firstweekday=6)
            weeks = cal.monthdayscalendar(anio_sel, mes_sel_idx)
            
            cols = st.columns(7)
            dias_header = ["DOM", "LUN", "MAR", "MIÉ", "JUE", "VIE", "SÁB"]
            for i, d in enumerate(dias_header):
                cols[i].markdown(f"<div style='text-align:center; font-weight:bold; background:#f0f2f6; padding:5px; border-radius:5px;'>{d}</div>", unsafe_allow_html=True)
            
            for week in weeks:
                cols = st.columns(7)
                for i, day in enumerate(week):
                    with cols[i]:
                        if day != 0:
                            f_key = f"{anio_sel}-{str(mes_sel_idx).zfill(2)}-{str(day).zfill(2)}"
                            acts = fechas_ocupadas.get(f_key, [])
                            bg = "#e8f4f8" if acts else "#ffffff"
                            bord = "2px solid #00b0f0" if acts else "1px solid #ddd"
                            
                            st.markdown(f"""
                            <div style='height:100px; border:{bord}; background:{bg}; border-radius:8px; padding:5px; display:flex; flex-direction:column; align-items:center; overflow:hidden;'>
                                <div style='font-weight:bold; font-size:14px; margin-bottom:2px;'>{day}</div>
                            </div>""", unsafe_allow_html=True)
                            
                            if acts:
                                if acts[0]['thumb']: st.image(acts[0]['thumb'], use_container_width=True)
                                else: st.caption("Sin Postal")
                                if len(acts) > 1: st.markdown(f"<div style='text-align:center; font-size:12px; font-weight:bold; color:red;'>+ {len(acts)-1} más</div>", unsafe_allow_html=True)
