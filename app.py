import streamlit as st
import requests
from pptx import Presentation
import os
import subprocess
import tempfile
from io import BytesIO
import cloudinary
import cloudinary.uploader
import pandas as pd

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")

if 'registros' not in st.session_state:
    st.session_state.registros = []
if 'record_map' not in st.session_state:
    st.session_state.record_map = {}

# --- FUNCIONES DE PROCESAMIENTO (Tu lógica original adaptada) ---

def reemplazar_texto(slide, reemplazos):
    """Recorre formas y tablas para reemplazar etiquetas {{Campo}}"""
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for tag, valor in reemplazos.items():
                        if tag in run.text:
                            run.text = run.text.replace(tag, str(valor if valor else ""))
        
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for paragraph in cell.text_frame.paragraphs:
                        for run in paragraph.runs:
                            for tag, valor in reemplazos.items():
                                if tag in run.text:
                                    run.text = run.text.replace(tag, str(valor if valor else ""))

def convertir_a_pdf(pptx_bytes):
    """Usa LibreOffice en el servidor para convertir"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_pptx:
        tmp_pptx.write(pptx_bytes)
        tmp_pptx_path = tmp_pptx.name
    
    try:
        # Comando para Streamlit Cloud (requiere packages.txt con 'libreoffice')
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', 
                        os.path.dirname(tmp_pptx_path), tmp_pptx_path], check=True)
        pdf_path = tmp_pptx_path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f:
            return f.read()
    except:
        return None

# --- INTERFAZ ---
st.title("🚀 Generador de Presentaciones")

with st.sidebar:
    st.header("🔑 Conexión")
    token = st.text_input("Airtable Token", type="password")
    
    if token:
        # Lógica de carga de bases/tablas (simplificada para el ejemplo)
        # Aquí puedes mantener tu lógica de selectores del mensaje anterior
        base_id = st.text_input("Base ID (appXXXXXXXX)")
        table_name = st.text_input("Nombre de la Tabla")
        
        if st.button("🔄 Cargar Datos de Airtable"):
            url = f"https://api.airtable.com/v0/{base_id}/{table_name}"
            headers = {"Authorization": f"Bearer {token}"}
            res = requests.get(url, headers=headers)
            if res.status_code == 200:
                data = res.json().get("records", [])
                st.session_state.registros = data
                # Usamos el campo 'Nombre' o el ID como clave
                st.session_state.record_map = {r['fields'].get('Nombre', r['id']): r for r in data}
                st.success(f"Cargados {len(data)} registros")

st.divider()

# --- ÁREA DE TRABAJO ---
if st.session_state.registros:
    col_a, col_b = st.columns([1, 2])
    
    with col_a:
        st.subheader("1. Selección")
        seleccionados = st.multiselect("Clientes a procesar:", options=list(st.session_state.record_map.keys()))
        plantilla = st.file_uploader("2. Subir Plantilla PPTX", type="pptx")
        
        do_pdf = st.checkbox("Convertir también a PDF")

    with col_b:
        st.subheader("3. Procesamiento")
        if st.button("🔥 GENERAR TODO") and plantilla and seleccionados:
            
            # Leemos la plantilla una vez para tenerla en memoria
            template_bytes = plantilla.read()
            
            for nombre in seleccionados:
                with st.expander(f"Procesando: {nombre}", expanded=True):
                    record = st.session_state.record_map[nombre]
                    fields = record['fields']
                    
                    # Crear una copia de la presentación en memoria
                    prs = Presentation(BytesIO(template_bytes))
                    
                    # Preparar diccionario de reemplazos {{Campo}} -> Valor
                    # Esto mapea automáticamente cualquier columna de Airtable
                    reemplazos = {f"{{{{{k}}}}}" : v for k, v in fields.items()}
                    
                    for slide in prs.slides:
                        reemplazar_texto(slide, reemplazos)
                    
                    # Guardar resultado
                    output_pptx = BytesIO()
                    prs.save(output_pptx)
                    pptx_final = output_pptx.getvalue()
                    
                    # Botón de descarga PPTX
                    st.download_button(f"📥 Descargar PPTX - {nombre}", pptx_final, f"{nombre}.pptx")
                    
                    if do_pdf:
                        with st.spinner("Convirtiendo a PDF..."):
                            pdf_final = convertir_a_pdf(pptx_final)
                            if pdf_final:
                                st.download_button(f"📄 Descargar PDF - {nombre}", pdf_final, f"{nombre}.pdf")
                            else:
                                st.error("No se pudo generar el PDF (LibreOffice no detectado)")

else:
    st.info("Configura los datos en el panel izquierdo para ver los registros.")
