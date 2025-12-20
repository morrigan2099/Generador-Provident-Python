import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from io import BytesIO
import os
import subprocess
import tempfile
import zipfile
from datetime import datetime
import locale
from pdf2image import convert_from_bytes

# Intentar establecer idioma español para los nombres de los meses
try:
    locale.setlocale(locale.LC_TIME, "es_ES.utf8")
except:
    pass

# --- CONFIGURACIÓN FIJA ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"

# --- FUNCIONES DE APOYO ---
def limpiar_adjuntos(valor):
    if isinstance(valor, list):
        return ", ".join([f.get("filename", "") for f in valor])
    return str(valor) if valor else ""

def generar_pdf(pptx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        tmp_path = tmp.name
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', 
                        '--outdir', os.path.dirname(tmp_path), tmp_path], check=True)
        pdf_path = tmp_path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f:
            pdf_data = f.read()
        if os.path.exists(tmp_path): os.remove(tmp_path)
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return pdf_data
    except: return None

def generar_png(pdf_bytes):
    try:
        images = convert_from_bytes(pdf_bytes)
        if images:
            img_byte_arr = BytesIO()
            images[0].save(img_byte_arr, format='PNG')
            return img_byte_arr.getvalue()
    except: return None

def formatear_nombre_archivo(fila, ext):
    """Lógica de nombres de archivo con manejo de nulos y limpieza de comas"""
    # Manejo seguro de nulos para campos de texto
    punto = str(fila.get('Punto de reunion') or '').strip()
    ruta = str(fila.get('Ruta a seguir') or '').strip()
    municipio = str(fila.get('Municipio') or '').strip()
    tipo = str(fila.get('Tipo') or 'TIPO').strip()
    sucursal = str(fila.get('Sucursal') or 'SUCURSAL').strip()
    
    # Manejo de Fecha
    try:
        dt = datetime.strptime(fila['Fecha'], '%Y-%m-%d')
        fecha_str = dt.strftime('%A %B %d de %Y').lower()
    except:
        fecha_str = datetime.now().strftime('%d de %m de %Y')

    # Lógica de longitud: si el nombre es muy largo, elegir el más corto entre punto y ruta
    if len(punto) + len(ruta) > 80:
        if len(punto) > len(ruta) and ruta:
            punto = ""
        elif punto:
            ruta = ""

    # Construcción del bloque central (limpieza de comas automática)
    partes_centro = [p for p in [punto, ruta] if p]
    centro = " - ".join(partes_centro)
    
    # Nombre final base
    nombre = f"{fecha_str} - {tipo} {sucursal}"
    if centro:
        nombre += f" - {centro}"
    if municipio:
        nombre += f", {municipio}"
    
    # Limpiar caracteres prohibidos en Windows/Linux
    for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
        nombre = nombre.replace(char, '')
    
    return f"{nombre[:200]}.{ext}" # Limitar a 200 caracteres para seguridad de ruta

# --- INTERFAZ ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")
st.title("🚀 Generador Masivo con Estructura de Carpetas")

with st.sidebar:
    st.header("🔑 Conexión Airtable")
    headers = {"Authorization": f"Bearer {TOKEN}"}
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    
    if r_bases.status_code == 200:
        bases = r_bases.json().get("bases", [])
        base_opts = {b['name']: b['id'] for b in bases}
        base_sel = st.selectbox("Selecciona Base:", list(base_opts.keys()))
        
        r_tablas = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
        if r_tablas.status_code == 200:
            tablas = r_tablas.json().get("tables", [])
            tabla_opts = {t['name']: t['id'] for t in tablas}
            tabla_sel = st.selectbox("Selecciona Tabla:", list(tabla_opts.keys()))
            
            r_regs = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
            st.session_state.registros_raw = r_regs.json().get("records", [])
            
            df = pd.DataFrame([r['fields'] for r in st.session_state.registros_raw])
            # Asegurar que existan las columnas para evitar KeyError
            for col in ["Tipo", "Sucursal", "Municipio", "Fecha", "Punto de reunion", "Ruta a seguir"]:
                if col not in df.columns: df[col] = ""
            
            cols_v = ["Tipo", "Sucursal", "Municipio", "Fecha"]
            df_display = df[cols_v].copy()
            df_display.insert(0, "Seleccionar", False)
            st.session_state.df_trabajo = df_display

# --- PANEL PRINCIPAL ---
if not st.session_state.get('df_trabajo', pd.DataFrame()).empty:
    st.subheader("1. Selección de Registros")
    df_editado = st.data_editor(st.session_state.df_trabajo, use_container_width=True, hide_index=True)
    seleccionados_indices = df_editado.index[df_editado["Seleccionar"] == True].tolist()

    if seleccionados_indices:
        st.divider()
        uso_final = st.radio("Seleccione uso final:", ["POSTALES", "REPORTES"], horizontal=True)
        folder_path = os.path.join(BASE_DIR, uso_final)
        
        if os.path.exists(folder_path):
            archivos_pptx = [f for f in os.listdir(folder_path) if f.endswith('.pptx')]
            tipos_unicos = df_editado.loc[seleccionados_indices, "Tipo"].unique()
            mapping_manual = {t: st.selectbox(f"Plantilla para {uso_final} TIPO: {t}", archivos_pptx, key=t) for t in tipos_unicos}

            if st.button("🔥 GENERAR ZIP ESTRUCTURADO"):
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                    for idx in seleccionados_indices:
                        fila_original = st.session_state.registros_raw[idx]['fields']
                        with st.status(f"Procesando {fila_original.get('Sucursal', 'Sin nombre')}..."):
                            
                            # Carpeta dinámica
                            try:
                                f_dt = datetime.strptime(fila_original.get('Fecha', ''), '%Y-%m-%d')
                            except:
                                f_dt = datetime.now()
                            
                            path_zip = f"Provident/{f_dt.year}/{f_dt.strftime('%m - %B').lower()}/{uso_final.capitalize()}/{fila_original.get('Sucursal', 'General')}/"
                            
                            # Procesar PPTX
                            prs = Presentation(os.path.join(folder_path, mapping_manual[fila_original.get('Tipo')]))
                            for slide in prs.slides:
                                for shape in slide.shapes:
                                    if shape.has_text_frame:
                                        for p in shape.text_frame.paragraphs:
                                            for run in p.runs:
                                                for k, v in fila_original.items():
                                                    tag = f"{{{{{k}}}}}"
                                                    if tag in run.text:
                                                        run.text = run.text.replace(tag, limpiar_adjuntos(v))
                            
                            pp_io = BytesIO(); prs.save(pp_io)
                            pdf_bin = generar_pdf(pp_io.getvalue())
                            
                            if pdf_bin:
                                ext = "pdf" if uso_final == "REPORTES" else "png"
                                nombre_arc = formatear_nombre_archivo(fila_original, ext)
                                
                                if uso_final == "REPORTES":
                                    zip_file.writestr(path_zip + nombre_arc, pdf_bin)
                                else:
                                    png_bin = generar_png(pdf_bin)
                                    if png_bin: zip_file.writestr(path_zip + nombre_arc, png_bin)

                st.success("✅ ¡ZIP Generado!")
                st.download_button("📥 DESCARGAR RESULTADOS", zip_buffer.getvalue(), f"Provident_{datetime.now().strftime('%Y%m%d')}.zip", "application/zip")
