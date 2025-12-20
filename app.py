import streamlit as st
import requests
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os
import subprocess
import sys
import tempfile
from io import BytesIO
import cloudinary
import cloudinary.uploader
from PIL import Image

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador Provident", layout="wide")

# --- ESTADO DE LA SESIÓN ---
if 'registros' not in st.session_state:
    st.session_state.registros = []
if 'record_map' not in st.session_state:
    st.session_state.record_map = {}

# --- FUNCIONES DE APOYO ---

def obtener_bases(token):
    url = "https://api.airtable.com/v0/meta/bases"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    return response.json().get("bases", []) if response.status_code == 200 else []

def obtener_tablas(token, base_id):
    url = f"https://api.airtable.com/v0/meta/bases/{base_id}/tables"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    return response.json().get("tables", []) if response.status_code == 200 else []

def cargar_datos_airtable(token, base_id, table_name):
    url = f"https://api.airtable.com/v0/{base_id}/{table_name}"
    headers = {"Authorization": f"Bearer {token}"}
    params = {"maxRecords": 100, "view": "Grid view"} # Ajusta según tu vista
    response = requests.get(url, headers=headers, params=params)
    if response.status_code == 200:
        data = response.json().get("records", [])
        return data
    return []

def convertir_a_pdf(pptx_data):
    """Convierte bytes de PPTX a PDF usando LibreOffice"""
    with tempfile.TemporaryDirectory() as tmpdirname:
        pptx_path = os.path.join(tmpdirname, "temp.pptx")
        with open(pptx_path, "wb") as f:
            f.write(pptx_data)
        
        # Comando para LibreOffice (funciona en Streamlit Cloud si usas packages.txt)
        try:
            subprocess.run([
                'soffice', '--headless', '--convert-to', 'pdf', 
                '--outdir', tmpdirname, pptx_path
            ], check=True)
            
            pdf_path = os.path.join(tmpdirname, "temp.pdf")
            with open(pdf_path, "rb") as f:
                return f.read()
        except Exception as e:
            st.error(f"Error al convertir a PDF: {e}")
            return None

# --- BARRA LATERAL (CONFIGURACIÓN) ---
st.sidebar.header("🔑 Configuración")
airtable_token = st.sidebar.text_input("Airtable Token", type="password")

if airtable_token:
    bases = obtener_bases(airtable_token)
    if bases:
        base_options = {b['name']: b['id'] for b in bases}
        selected_base_name = st.sidebar.selectbox("Selecciona Base", list(base_options.keys()))
        base_id = base_options[selected_base_name]
        
        tablas = obtener_tablas(airtable_token, base_id)
        if tablas:
            tabla_options = {t['name']: t['id'] for t in tablas}
            selected_table_name = st.sidebar.selectbox("Selecciona Tabla", list(tabla_options.keys()))
            
            if st.sidebar.button("🔄 Cargar Registros"):
                with st.spinner("Leyendo Airtable..."):
                    data = cargar_datos_airtable(airtable_token, base_id, selected_table_name)
                    st.session_state.registros = data
                    st.session_state.record_map = {r['fields'].get('Nombre', 'Sin Nombre'): r for r in data}
                    st.sidebar.success(f"{len(data)} registros cargados.")

st.sidebar.divider()
st.sidebar.header("☁️ Cloudinary (Opcional)")
cloud_name = st.sidebar.text_input("Cloud Name")
api_key = st.sidebar.text_input("API Key")
api_secret = st.sidebar.text_input("API Secret", type="password")

if cloud_name and api_key and api_secret:
    cloudinary.config(cloud_name=cloud_name, api_key=api_key, api_secret=api_secret)

# --- PANEL PRINCIPAL ---
st.title("🚀 Generador de Presentaciones Provident")

if not st.session_state.registros:
    st.info("Configura tu Token de Airtable en la izquierda para comenzar.")
else:
    # Selección de registros a procesar
    seleccionados = st.multiselect(
        "Selecciona los clientes para generar presentación:",
        options=list(st.session_state.record_map.keys())
    )

    plantilla_file = st.file_uploader("Sube la plantilla PowerPoint (.pptx)", type="pptx")

    if plantilla_file and seleccionados:
        if st.button("🪄 Generar y Convertir"):
            for nombre in seleccionados:
                record = st.session_state.record_map[nombre]
                fields = record['fields']
                
                with st.status(f"Procesando a {nombre}...", expanded=True) as status:
                    # 1. Leer plantilla
                    prs = Presentation(plantilla_file)
                    
                    # 2. Lógica de reemplazo (Ejemplo basado en tu código)
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for paragraph in shape.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        # Ejemplo de reemplazo: {{Nombre}} -> Valor en Airtable
                                        if "{{Nombre}}" in run.text:
                                            run.text = run.text.replace("{{Nombre}}", fields.get("Nombre", ""))

                    # 3. Guardar en memoria
                    pptx_io = BytesIO()
                    prs.save(pptx_io)
                    pptx_bytes = pptx_io.getvalue()
                    
                    st.write("✅ PowerPoint generado.")
                    
                    # 4. Conversión a PDF (Opcional)
                    pdf_bytes = convertir_a_pdf(pptx_bytes)
                    
                    # 5. Botones de descarga
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(f"⬇️ Descargar PPTX ({nombre})", pptx_bytes, f"{nombre}.pptx")
                    with col2:
                        if pdf_bytes:
                            st.download_button(f"📄 Descargar PDF ({nombre})", pdf_bytes, f"{nombre}.pdf")
                    
                    status.update(label=f"¡Listo para {nombre}!", state="complete")
