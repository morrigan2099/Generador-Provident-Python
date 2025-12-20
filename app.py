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
from pdf2image import convert_from_bytes

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
    """Lógica compleja de nombres de archivo y limpieza de comas"""
    # 1. Obtener fecha de Airtable o actual
    try:
        dt = datetime.strptime(fila['Fecha'], '%Y-%m-%d')
    except:
        dt = datetime.now()
    
    fecha_str = dt.strftime('%d de %B de %Y').lower() # dddd mmmm dd de aaaa (aprox)
    tipo = fila.get('Tipo', 'TIPO')
    sucursal = fila.get('Sucursal', 'SUCURSAL')
    punto = fila.get('Punto de reunion', '').strip()
    ruta = fila.get('Ruta a seguir', '').strip()
    municipio = fila.get('Municipio', '').strip()

    # 2. Lógica de selección por longitud (si es muy largo, elegir el más corto)
    if len(punto) + len(ruta) > 100:
        if len(punto) > len(ruta) and ruta:
            punto = ""
        elif punto:
            ruta = ""

    # 3. Construcción con limpieza de comas
    componentes = [punto, ruta]
    componentes = [c for c in componentes if c] # Elimina vacíos
    centro = ", ".join(componentes)
    
    final_str = f"{fecha_str} - {tipo} {sucursal}"
    if centro:
        final_str += f" - {centro}"
    if municipio:
        final_str += f", {municipio}"
    
    # Limpiar caracteres prohibidos en nombres de archivo
    for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
        final_str = final_str.replace(char, '')

    return f"{final_str}.{ext}"

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
            # Incluimos los campos necesarios para el nombre del archivo
            campos_necesarios = ["Tipo", "Sucursal", "Seccion", "Municipio", "Fecha", "Punto de reunion", "Ruta a seguir"]
            cols_v = [c for c in campos_necesarios if c in df.columns]
            df_display = df[cols_v].copy()
            df_display.insert(0, "Seleccionar", False)
            st.session_state.df_trabajo = df_display

if not st.session_state.get('df_trabajo', pd.DataFrame()).empty:
    st.subheader("1. Selección de Registros")
    df_editado = st.data_editor(st.session_state.df_trabajo, use_container_width=True, hide_index=True)
    seleccionados = df_editado[df_editado["Seleccionar"] == True]

    if not seleccionados.empty:
        st.divider()
        uso_final = st.radio("Seleccione uso final:", ["POSTALES", "REPORTES"], horizontal=True)
        folder_path = os.path.join(BASE_DIR, uso_final)
        
        if os.path.exists(folder_path):
            archivos_pptx = [f for f in os.listdir(folder_path) if f.endswith('.pptx')]
            tipos_unicos = seleccionados["Tipo"].unique()
            mapping_manual = {t: st.selectbox(f"Plantilla para {uso_final} TIPO: {t}", archivos_pptx, key=t) for t in tipos_unicos}

            if st.button("🔥 GENERAR ZIP ESTRUCTURADO"):
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                    for idx, fila in seleccionados.iterrows():
                        with st.status(f"Procesando {fila['Sucursal']}..."):
                            # Carpeta dinámica
                            try:
                                fecha_dt = datetime.strptime(fila['Fecha'], '%Y-%m-%d')
                            except:
                                fecha_dt = datetime.now()
                            
                            anio = fecha_dt.strftime('%Y')
                            mes = fecha_dt.strftime('%m - %B').lower()
                            uso_folder = "Reportes" if uso_final == "REPORTES" else "Postales"
                            sucursal_folder = fila['Sucursal']
                            
                            # Ruta interna ZIP: Provident/AAAA/mm - mmmm/Uso/Sucursal/
                            ruta_zip_base = f"Provident/{anio}/{mes}/{uso_folder}/{sucursal_folder}/"
                            
                            # Procesar PPTX
                            datos_record = st.session_state.registros_raw[idx]['fields']
                            prs = Presentation(os.path.join(folder_path, mapping_manual[fila['Tipo']]))
                            for slide in prs.slides:
                                for shape in slide.shapes:
                                    if shape.has_text_frame:
                                        for p in shape.text_frame.paragraphs:
                                            for run in p.runs:
                                                for k, v in datos_record.items():
                                                    if f"{{{{{k}}}}}" in run.text:
                                                        run.text = run.text.replace(f"{{{{{k}}}}}", limpiar_adjuntos(v))
                            
                            pptx_io = BytesIO(); prs.save(pptx_io)
                            pdf_data = generar_pdf(pptx_io.getvalue())
                            
                            if pdf_data:
                                ext = "pdf" if uso_final == "REPORTES" else "png"
                                nombre_f = formatear_nombre_archivo(fila, ext)
                                
                                if uso_final == "REPORTES":
                                    zip_file.writestr(ruta_zip_base + nombre_f, pdf_data)
                                else:
                                    png_data = generar_png(pdf_data)
                                    if png_data: zip_file.writestr(ruta_zip_base + nombre_f, png_data)

                st.success("✅ ZIP Estructurado creado con éxito")
                st.download_button("📥 DESCARGAR ZIP COMPLETO", zip_buffer.getvalue(), f"Provident_{uso_final}.zip", "application/zip")
