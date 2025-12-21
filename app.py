import streamlit as st
import requests
import pandas as pd
import json
import os
import re
import unicodedata
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import subprocess, tempfile, zipfile
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN ---
CONFIG_FILE = "config_app.json"

def cargar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f: 
                cfg = json.load(f)
                return cfg
        except: pass
    return {"plantillas": {}, "columnas_visibles": []}

def guardar_config_json(config_data):
    with open(CONFIG_FILE, "w") as f: 
        json.dump(config_data, f, indent=4)

if 'config' not in st.session_state:
    st.session_state.config = cargar_config()

# --- MOTOR DE TEXTO (RECONSTRUIDO DESDE CERO) ---
def procesar_texto_maestro(valor_crudo, nombre_campo=""):
    if not valor_crudo or str(valor_crudo).lower() == "none":
        return ""
    if isinstance(valor_crudo, list):
        return valor_crudo # Mantiene los adjuntos intactos
    
    texto = str(valor_crudo).strip()
    
    if nombre_campo == 'Seccion':
        return texto.upper()
    
    # Capitalización simple: Primera letra mayúscula, resto minúscula
    # Se aplica a campos generales para evitar errores de lógica compleja
    palabras = texto.lower().split()
    if not palabras:
        return ""
    
    preposiciones = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al']
    resultado = []
    for i, p in enumerate(palabras):
        if i == 0 or p not in preposiciones:
            resultado.append(p.capitalize())
        else:
            resultado.append(p)
            
    return " ".join(resultado)

# --- FUNCIÓN PDF ---
def generar_pdf(pptx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        path_pptx = tmp.name
    try:
        # Comando para Linux (Streamlit Cloud)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(path_pptx), path_pptx], check=True)
        path_pdf = path_pptx.replace(".pptx", ".pdf")
        with open(path_pdf, "rb") as f:
            pdf_data = f.read()
        if os.path.exists(path_pptx): os.remove(path_pptx)
        if os.path.exists(path_pdf): os.remove(path_pdf)
        return pdf_data
    except Exception:
        return None

# --- INTERFAZ ---
st.set_page_config(page_title="Provident Pro v40", layout="wide")
st.title("🚀 Generador Pro: Fix de Estructura y Tipo 11pts")

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

with st.sidebar:
    st.header("⚙️ Panel de Control")
    if st.button("💾 GUARDAR CONFIGURACIÓN"):
        guardar_config_json(st.session_state.config)
        st.toast("Guardado")

    st.divider()
    # Conexión Airtable
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        bases_map = {b['name']: b['id'] for b in r_bases.json().get('bases', [])}
        base_sel = st.selectbox("Base:", [""] + list(bases_map.keys()))
        
        if base_sel:
            bid = bases_map[base_sel]
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{bid}/tables", headers=headers)
            tablas_map = {t['name']: t['id'] for t in r_tab.json().get('tables', [])}
            tabla_sel = st.selectbox("Tabla:", list(tablas_map.keys()))
            
            if st.button("🔄 CARGAR DATOS"):
                r_reg = requests.get(f"https://api.airtable.com/v0/{bid}/{tablas_map[tabla_sel]}", headers=headers)
                recs = r_reg.json().get("records", [])
                st.session_state.raw_data_original = recs
                # Procesamiento de campos
                st.session_state.raw_records = [
                    {'id': r['id'], 'fields': {k: (procesar_texto_maestro(v, k) if k != 'Fecha' else v) for k, v in r['fields'].items()}} 
                    for r in recs
                ]
                st.rerun()

# --- VISTA DE TABLA ---
if 'raw_records' in st.session_state:
    df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    
    with st.sidebar:
        st.divider()
        visibles = [c for c in st.session_state.config.get("columnas_visibles", []) if c in df_full.columns] or list(df_full.columns)
        selected_cols = st.multiselect("Columnas Visibles:", list(df_full.columns), default=visibles)
        st.session_state.config["columnas_visibles"] = selected_cols

    df_view = df_full[[c for c in selected_cols if c in df_full.columns]].copy()
    # Eliminar columnas que contienen listas (adjuntos) para no romper la tabla
    for col in df_view.columns:
        if not df_view.empty and isinstance(df_view[col].iloc[0], list):
            df_view.drop(col, axis=1, inplace=True)
    
    df_view.insert(0, "Seleccionar", False)
    
    # Editor con Checkbox Maestro en el Header
    df_edit = st.data_editor(
        df_view, use_container_width=True, hide_index=True,
        column_config={"Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False)}
    )
    
    indices_seleccionados = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if indices_seleccionados:
        modo = st.radio("Formato:", ["Postales", "Reportes"], horizontal=True)
        path_plantillas = os.path.join("Plantillas", modo.upper())
        AZUL_CORP = RGBColor(0, 176, 240)
        
        archivos_pptx = [f for f in os.listdir(path_plantillas) if f.endswith('.pptx')]
        tipos_en_seleccion = df_view.loc[indices_seleccionados, "Tipo"].unique()
        
        for t in tipos_en_seleccion:
            p_previa = st.session_state.config.get("plantillas", {}).get(t)
            idx = archivos_pptx.index(p_previa) if p_previa in archivos_pptx else 0
            st.session_state.config.setdefault("plantillas", {})[t] = st.selectbox(f"Plantilla para {t}:", archivos_pptx, index=idx)

        if st.button("🔥 GENERAR SELECCIONADOS", use_container_width=True, type="primary"):
            p_bar = st.progress(0); zip_buffer = BytesIO()
            total = len(indices_seleccionados)
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                for i, idx in enumerate(indices_seleccionados):
                    rec_proc = st.session_state.raw_records[idx]['fields']
                    rec_orig = st.session_state.raw_data_original[idx]['fields']
                    
                    f_tipo = rec_proc.get('Tipo')
                    f_suc = rec_proc.get('Sucursal')
                    
                    # Cargar Plantilla
                    prs = Presentation(os.path.join(path_plantillas, st.session_state.config["plantillas"][f_tipo]))
                    
                    reemplazos = {
                        "<<Tipo>>": f_tipo,
                        "<<Sucursal>>": f_suc,
                        "<<Seccion>>": rec_proc.get('Seccion'),
                        "<<Confechor>>": f"{rec_proc.get('Fecha')}, {rec_proc.get('Hora')}",
                        "<<Concat>>": f"{rec_proc.get('Punto de reunion') or rec_proc.get('Ruta a seguir')}, {rec_proc.get('Municipio')}"
                    }

                    for slide in prs.slides:
                        # Inserción de Imágenes
                        for shape in list(slide.shapes):
                            nombre_shape = shape.text_frame.text if shape.has_text_frame else ""
                            tags_img = ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]
                            
                            for tag in tags_img:
                                if f"<<{tag}>>" in nombre_shape or tag == shape.name:
                                    adjuntos = rec_orig.get(tag)
                                    if adjuntos and isinstance(adjuntos, list):
                                        try:
                                            resp = requests.get(adjuntos[0].get('url'))
                                            slide.shapes.add_picture(BytesIO(resp.content), shape.left, shape.top, shape.width, shape.height)
                                            # Borrar placeholder
                                            el = shape._element
                                            el.getparent().remove(el)
                                        except: pass

                        # Reemplazo de Texto
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for tag, valor in reemplazos.items():
                                    if tag in shape.text_frame.text:
                                        tf = shape.text_frame; tf.clear()
                                        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
                                        run = p.add_run()
                                        run.text = str(valor)
                                        run.font.bold = True
                                        run.font.color.rgb = AZUL_CORP
                                        # TAMAÑOS: Tipo 11, Sucursal 14, Otros 11
                                        if tag == "<<Tipo>>": run.font.size = Pt(11)
                                        elif tag == "<<Sucursal>>": run.font.size = Pt(14)
                                        else: run.font.size = Pt(11)

                    # Guardado
                    pp_out = BytesIO(); prs.save(pp_out)
                    final_data = generar_pdf(pp_out.getvalue())
                    
                    if final_data:
                        ext = ".pdf" if modo == "Reportes" else ".jpg"
                        nombre_final = f"{rec_proc.get('Fecha')} - {f_tipo} - {f_suc}{ext}"
                        zip_file.writestr(f"{modo}/{f_suc}/{nombre_final}", final_data if modo == "Reportes" else convert_from_bytes(final_data)[0].tobytes())
                    
                    p_bar.progress((i + 1) / total)
            
            st.success("¡Generación Exitosa!")
            st.download_button("📥 DESCARGAR RESULTADOS", zip_buffer.getvalue(), "Provident_v40.zip", use_container_width=True)
