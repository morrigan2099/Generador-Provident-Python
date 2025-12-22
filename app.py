import streamlit as st
import requests
import pandas as pd
import json
import os
import re
import numpy as np
import textwrap
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

# --- CONFIGURACIÓN DE TAMAÑOS ---
TAM_TIPO_BASE = 12
TAM_SUCURSAL  = 13
TAM_SECCION   = 11
TAM_CONFECHOR = 11.5 
TAM_CONCAT    = 11

# --- CONSTANTES ---
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]

# ============================================================
#  LÓGICA DE RECORTE IMAGEN
# ============================================================

def recorte_inteligente_bordes(img, umbral_negro=60):
    img_gray = img.convert("L")
    arr = np.array(img_gray)
    h, w = arr.shape

    def fila_es_negra(fila):
        return (np.sum(fila < 35) / fila.size) * 100 > umbral_negro

    def columna_es_negra(col):
        return (np.sum(col < 35) / col.size) * 100 > umbral_negro

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

# --- FUNCIONES DE LÓGICA DE NEGOCIO ---

def procesar_confechor_logica(fecha_dt, hora_str):
    nombre_mes = MESES_ES[fecha_dt.month - 1]
    dia = fecha_dt.day
    anio = fecha_dt.year
    
    # 🔴 CAMBIO SOLICITADO: Formato "mmmm dd 'de' aaaa"
    # Ejemplo: "diciembre 22 de 2025"
    linea_fecha = f"{nombre_mes} {dia} de {anio}"
    
    hora_final = ""
    if hora_str and str(hora_str).lower() != "none":
        hh_mm = str(hora_str)[0:5] 
        try:
            parts = hh_mm.split(':')
            h_24 = int(parts[0])
            minutos = parts[1]
            
            if 8 <= h_24 <= 11: sufijo = "de la mañana"
            elif h_24 == 12: sufijo = "del día"
            elif 13 <= h_24 <= 19: sufijo = "de la tarde"
            else: sufijo = "p.m." if h_24 >= 12 else "a.m."

            h_mostrar = h_24
            if h_24 > 12: h_mostrar = h_24 - 12
            elif h_24 == 0: h_mostrar = 12
            
            h_str = str(h_mostrar)
            hora_final = f"{h_str}:{minutos} {sufijo}"
        except:
            hora_final = hh_mm
            
    return f"{linea_fecha}\n{hora_final}"

# --- FUNCIONES DE IMAGEN Y TEXTO ---

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
        subprocess.run(
            ['soffice', '--headless', '--convert-to', 'pdf',
             '--outdir', os.path.dirname(path), path],
            check=True
        )
        pdf_path = path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f:
            data = f.read()
        os.remove(path)
        os.remove(pdf_path)
        return data
    except:
        return None

# --- UI STREAMLIT ---
st.set_page_config(page_title="Provident Pro v69", layout="wide")

if 'config' not in st.session_state:
    if os.path.exists("config_app.json"):
        with open("config_app.json", "r") as f:
            st.session_state.config = json.load(f)
    else:
        st.session_state.config = {"plantillas": {}}

st.title("🚀 Generador Pro v69 - Final")

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

with st.sidebar:
    st.header("⚙️ Configuración")
    if st.button("💾 GUARDAR PLANTILLAS"):
        with open("config_app.json", "w") as f:
            json.dump(st.session_state.config, f)
        st.toast("Plantillas guardadas")

    st.divider()

    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base:", [""] + list(base_opts.keys()))
        if base_sel:
            r_tab = requests.get(
                f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables",
                headers=headers
            )
            tabla_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
            tabla_sel = st.selectbox("Tabla:", list(tabla_opts.keys()))
            if st.button("🔄 CARGAR DATOS"):
                r_reg = requests.get(
                    f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}",
                    headers=headers
                )
                recs = r_reg.json().get("records", [])
                st.session_state.raw_data_original = recs
                st.session_state.raw_records = [
                    {
                        'id': r['id'],
                        'fields': {
                            k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v)
                            for k, v in r['fields'].items()
                        }
                    }
                    for r in recs
                ]
                st.rerun()

if 'raw_records' in st.session_state:
    modo = st.radio("Salida:", ["Postales", "Reportes"], horizontal=True)

    df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    df_view = df_full.copy()

    for c in df_view.columns:
        if isinstance(df_view[c].iloc[0], list):
            df_view.drop(c, axis=1, inplace=True)

    if 'select_all' not in st.session_state:
        st.session_state.select_all = False

    c1, c2, _ = st.columns([1, 1, 4])
    if c1.button("✅ Todo"):
        st.session_state.select_all = True
        st.rerun()
    if c2.button("❌ Nada"):
        st.session_state.select_all = False
        st.rerun()

    df_view.insert(0, "Seleccionar", st.session_state.select_all)
    df_edit = st.data_editor(df_view, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        folder_fisica = os.path.join("Plantillas", modo.upper())
        if not os.path.exists(folder_fisica):
            os.makedirs(folder_fisica)

        archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]

        tipos_sel = df_view.loc[sel_idx, "Tipo"].unique()
        for t in tipos_sel:
            p_mem = st.session_state.config["plantillas"].get(t)
            idx_def = archivos_pptx.index(p_mem) if p_mem in archivos_pptx else 0
            st.session_state.config["plantillas"][t] = st.selectbox(
                f"Plantilla {t}:",
                archivos_pptx,
                index=idx_def,
                key=f"p_{t}"
            )

        # --- SECCIÓN DE GENERACIÓN Y LISTA DE SELECCIÓN ---
        if st.button("🔥 GENERAR", use_container_width=True, type="primary"):
            p_bar = st.progress(0)
            status_text = st.empty()
            
            # 1. Limpiamos la memoria anterior
            st.session_state.archivos_en_memoria = []
            
            AZUL_CELESTE = RGBColor(0, 176, 240)
            
            mapa_tamanos = {
                "<<Tipo>>": TAM_TIPO_BASE,
                "<<Sucursal>>": TAM_SUCURSAL,
                "<<Seccion>>": TAM_SECCION,
                "<<Confechor>>": TAM_CONFECHOR,
                "<<Concat>>": TAM_CONCAT
            }

            for i, idx in enumerate(sel_idx):
                record = st.session_state.raw_records[idx]['fields']
                record_orig = st.session_state.raw_data_original[idx]['fields']

                try:
                    raw_fecha = record.get('Fecha', '2025-01-01')
                    dt = datetime.strptime(raw_fecha, '%Y-%m-%d')
                    nombre_mes = MESES_ES[max(0, min(11, dt.month - 1))]
                except:
                    dt = datetime.now()
                    nombre_mes = "error"

                f_tipo = record.get('Tipo', 'Sin Tipo')
                f_suc = record.get('Sucursal', '000')
                lugar = record.get('Punto de reunion') or record.get('Ruta a seguir')

                texto_lugares = ", ".join([
                    str(x) for x in [lugar, record.get('Municipio')]
                    if x and str(x).lower() != 'none'
                ])
                
                if f_tipo == "Actividad en Sucursal":
                    f_concat = f"Sucursal {f_suc}"
                else:
                    f_concat = texto_lugares

                # Nombre base del archivo (sin extensión aún)
                nom_arch_base = (
                    f"{dt.day} de {nombre_mes} de {dt.year} - {f_tipo}, {f_suc}"
                    + ("" if f_tipo == "Actividad en Sucursal" else f" - {f_concat}")
                )
                nom_arch_base = re.sub(r'[\\/*?:"<>|]', "", nom_arch_base)

                status_text.markdown(f"**Procesando:** `{nom_arch_base}`")

                f_tipo_procesado = textwrap.fill(f_tipo, width=35)
                f_confechor_procesado = procesar_confechor_logica(dt, record.get('Hora', ''))

                reemplazos = {
                    "<<Tipo>>": f_tipo_procesado,
                    "<<Sucursal>>": f_suc,
                    "<<Seccion>>": record.get('Seccion'),
                    "<<Confechor>>": f_confechor_procesado,
                    "<<Concat>>": f_concat
                }

                prs = Presentation(
                    os.path.join(folder_fisica, st.session_state.config["plantillas"][f_tipo])
                )

                if f_tipo == "Actividad en Sucursal":
                    adj_lista = record_orig.get("Lista de asistencia")
                    if not adj_lista or len(adj_lista) == 0:
                        if len(prs.slides) >= 4:
                            xml_slides = prs.slides._sldIdLst
                            xml_slides.remove(xml_slides[3])

                for s_idx, slide in enumerate(prs.slides):
                    # IMÁGENES
                    for shape in list(slide.shapes):
                        if shape.has_text_frame:
                            tags_foto = [
                                "Foto de equipo", "Foto 01", "Foto 02", "Foto 03",
                                "Foto 04", "Foto 05", "Foto 06", "Foto 07",
                                "Reporte firmado", "Lista de asistencia"
                            ]
                            for tf_tag in tags_foto:
                                if f"<<{tf_tag}>>" in shape.text_frame.text:
                                    adj_data = record_orig.get(tf_tag)
                                    if adj_data:
                                        thumbs = adj_data[0].get('thumbnails', {})
                                        url = (
                                            thumbs.get('full', {}).get('url')
                                            or thumbs.get('large', {}).get('url')
                                            or adj_data[0].get('url')
                                        )
                                        try:
                                            r_img = requests.get(url).content
                                            img_io = procesar_imagen_inteligente(
                                                r_img, shape.width, shape.height, con_blur=(s_idx < 2)
                                            )
                                            slide.shapes.add_picture(
                                                img_io, shape.left, shape.top, shape.width, shape.height
                                            )
                                            sp = shape._element
                                            sp.getparent().remove(sp)
                                        except: pass

                    # TEXTO
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for tag, val in reemplazos.items():
                                if tag in shape.text_frame.text:
                                    
                                    if tag in ["<<Tipo>>", "<<Sucursal>>"]:
                                        shape.top = shape.top - Pt(2)
                                    
                                    tf = shape.text_frame
                                    tf.word_wrap = True 
                                    
                                    bodyPr = tf._element.bodyPr
                                    for child in ['spAutoFit', 'normAutofit', 'noAutofit']:
                                        existing = bodyPr.find(qn(f'a:{child}'))
                                        if existing is not None:
                                            bodyPr.remove(existing)

                                    if tag == "<<Concat>>":
                                        bodyPr.append(tf._element.makeelement(qn('a:normAutofit')))
                                    elif tag == "<<Tipo>>":
                                        bodyPr.append(tf._element.makeelement(qn('a:spAutoFit')))
                                    
                                    tf.clear()
                                    p = tf.paragraphs[0]
                                    if tag == "<<Confechor>>": p.alignment = PP_ALIGN.CENTER
                                    
                                    run = p.add_run()
                                    run.text = str(val)
                                    run.font.bold = True
                                    run.font.color.rgb = AZUL_CELESTE
                                    final_size = mapa_tamanos.get(tag, 12)
                                    run.font.size = Pt(final_size)

                pp_io = BytesIO()
                prs.save(pp_io)
                data_out = generar_pdf(pp_io.getvalue())
                
                if data_out:
                    ext = ".pdf" if modo == "Reportes" else ".jpg"
                    bytes_finales = data_out if modo == "Reportes" else convert_from_bytes(data_out)[0].tobytes()
                    nombre_archivo = f"{nom_arch_base[:120]}{ext}"
                    
                    # Ruta Interna del ZIP (ESTRUCTURA DE ARBOL)
                    # Quitamos "Provident/" del inicio para evitar la carpeta duplicada al descomprimir
                    ruta_zip = f"{dt.year}/{str(dt.month).zfill(2)} - {nombre_mes}/{modo}/{f_suc}/{nombre_archivo}"

                    st.session_state.archivos_en_memoria.append({
                        "Seleccionar": True, # Checkbox activado por defecto
                        "Archivo": nombre_archivo,
                        "RutaZip": ruta_zip,
                        "Datos": bytes_finales,
                        "Sucursal": f_suc,
                        "Tipo": f_tipo
                    })

                p_bar.progress((i + 1) / len(sel_idx))
            
            status_text.success("✅ Archivos procesados. Selecciona abajo cuáles quieres descargar.")

        # --- GESTOR DE DESCARGA CON CHECKBOXES ---
        if "archivos_en_memoria" in st.session_state and len(st.session_state.archivos_en_memoria) > 0:
            st.divider()
            st.subheader("📋 Selecciona los archivos a descargar")
            
            # Controles Masivos
            c_all, c_none, _ = st.columns([1, 1, 3])
            if c_all.button("☑️ Marcar Todos"):
                for item in st.session_state.archivos_en_memoria: item["Seleccionar"] = True
                st.rerun()
            if c_none.button("⬜ Desmarcar Todos"):
                for item in st.session_state.archivos_en_memoria: item["Seleccionar"] = False
                st.rerun()

            # Tabla Interactiva (Data Editor)
            df_display = pd.DataFrame(st.session_state.archivos_en_memoria)
            # Mostramos solo columnas visuales
            edited_df = st.data_editor(
                df_display[["Seleccionar", "Archivo", "Sucursal", "Tipo"]],
                hide_index=True,
                use_container_width=True,
                column_config={
                    "Seleccionar": st.column_config.CheckboxColumn(required=True),
                    "Archivo": st.column_config.TextColumn(width="large")
                }
            )
            
            # Recuperar índices seleccionados
            selected_rows = edited_df[edited_df["Seleccionar"] == True]
            indices_seleccionados = selected_rows.index.tolist()
            
            archivos_finales = [st.session_state.archivos_en_memoria[i] for i in indices_seleccionados]
            cant = len(archivos_finales)
            
            st.divider()
            st.write(f"📥 **{cant} archivos seleccionados**")
            
            if cant > 0:
                # Creación del ZIP al vuelo solo con lo seleccionado
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                    for item in archivos_finales:
                        # Usamos 'RutaZip' que contiene el árbol (Año/Mes/etc)
                        zip_file.writestr(item["RutaZip"], item["Datos"])
                
                fecha_zip = datetime.now().strftime('%Y%m%d_%H%M%S')
                
                st.download_button(
                    label=f"📦 DESCARGAR SELECCIONADOS (Mantiene Carpetas)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Reportes_{fecha_zip}.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )
                st.caption("ℹ️ Este ZIP contiene el árbol de carpetas exacto. Al descomprimirlo, los archivos se acomodarán en sus carpetas correspondientes.")
