import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from io import BytesIO
import os
import json
import subprocess
import tempfile
from pdf2image import convert_from_bytes
import cloudinary
import cloudinary.uploader

# --- CONFIGURACIÓN DE ARCHIVOS ---
CONFIG_FILE = "config_plantillas.json"

# Cargar mapeo de plantillas desde JSON
if 'mapping' not in st.session_state:
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            st.session_state.mapping = json.load(f)
    else:
        st.session_state.mapping = {}

# --- FUNCIONES DE PROCESAMIENTO ---

def procesar_pptx(plantilla_bytes, fields):
    """Reemplaza etiquetas {{Campo}} en el PowerPoint"""
    prs = Presentation(BytesIO(plantilla_bytes))
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for key, value in fields.items():
                            tag = f"{{{{{key}}}}}"
                            if tag in run.text:
                                # Manejo de valores nulos o listas
                                val_str = str(value) if value and not isinstance(value, list) else ""
                                run.text = run.text.replace(tag, val_str)
    out = BytesIO()
    prs.save(out)
    return out.getvalue()

def generar_pdf(pptx_bytes):
    """Convierte PPTX a PDF usando LibreOffice"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        tmp_path = tmp.name
    
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', 
                        '--outdir', os.path.dirname(tmp_path), tmp_path], check=True)
        pdf_path = tmp_path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f:
            pdf_data = f.read()
        os.remove(tmp_path)
        os.remove(pdf_path)
        return pdf_data
    except Exception as e:
        st.error(f"Error en PDF: {e}")
        return None

def generar_png(pdf_bytes):
    """Convierte la primera página del PDF a PNG"""
    try:
        # Convertir PDF a lista de imágenes de PIL
        images = convert_from_bytes(pdf_bytes)
        if images:
            img_byte_arr = BytesIO()
            images[0].save(img_byte_arr, format='PNG')
            return img_byte_arr.getvalue()
    except Exception as e:
        st.error(f"Error en PNG: {e}")
        return None

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Generador Provident", layout="wide")
st.title("🚀 Generador de Postales (PNG) y Reportes (PDF)")

# [Aquí iría tu lógica de Sidebar para conectar Airtable que ya definimos]
# Supongamos que ya tenemos st.session_state.df_trabajo cargado...

if not st.session_state.get('df_trabajo', pd.DataFrame()).empty:
    
    # Tabla con Checkboxes
    df_editado = st.data_editor(
        st.session_state.df_trabajo,
        use_container_width=True,
        hide_index=True,
        column_config={"Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False)},
        disabled=[c for c in st.session_state.df_trabajo.columns if c != "Seleccionar"]
    )

    seleccionados = df_editado[df_editado["Seleccionar"] == True]

    if not seleccionados.empty:
        tipos_en_seleccion = seleccionados["Tipo"].unique()
        
        # --- GESTIÓN DE PLANTILLAS POR TIPO ---
        st.subheader("📁 Configuración de Plantillas por Tipo")
        config_completa = True
        
        for t in tipos_en_seleccion:
            if t not in st.session_state.mapping:
                st.warning(f"No hay plantilla para el tipo: **{t}**")
                file = st.file_uploader(f"Subir PPTX para {t}", type="pptx", key=f"p_{t}")
                if file:
                    # Guardar archivo localmente
                    p_path = f"plantilla_{t}.pptx"
                    with open(p_path, "wb") as f:
                        f.write(file.getbuffer())
                    st.session_state.mapping[t] = p_path
                    with open(CONFIG_FILE, 'w') as f:
                        json.dump(st.session_state.mapping, f)
                    st.rerun()
                config_completa = False
        
        if config_completa:
            if st.button("🔥 GENERAR TODO"):
                for _, fila in seleccionados.iterrows():
                    with st.expander(f"Procesando: {fila['Sucursal']}", expanded=True):
                        # 1. Cargar plantilla según tipo
                        path_p = st.session_state.mapping[fila["Tipo"]]
                        with open(path_p, "rb") as f:
                            p_bytes = f.read()
                        
                        # 2. Generar PPTX con datos
                        pptx_res = procesar_pptx(p_bytes, fila.to_dict())
                        
                        # 3. Generar PDF (Base para ambos)
                        pdf_res = generar_pdf(pptx_res)
                        
                        if pdf_res:
                            col1, col2 = st.columns(2)
                            
                            # Opción Reporte (PDF)
                            col1.download_button(
                                "📄 Descargar PDF (Reporte)", 
                                pdf_res, 
                                f"Reporte_{fila['Sucursal']}.pdf", 
                                mime="application/pdf"
                            )
                            
                            # Opción Postal (PNG)
                            png_res = generar_png(pdf_res)
                            if png_res:
                                col2.download_button(
                                    "🖼️ Descargar PNG (Postal)", 
                                    png_res, 
                                    f"Postal_{fila['Sucursal']}.png", 
                                    mime="image/png"
                                )
