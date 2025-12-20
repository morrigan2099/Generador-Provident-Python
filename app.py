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

# --- CONFIGURACIÓN DE ARCHIVOS Y ESTADO ---
CONFIG_FILE = "config_plantillas.json"

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
                                # Limpieza básica de datos de Airtable
                                if isinstance(value, list): # Caso de adjuntos
                                    val_str = ", ".join([f.get("filename", "") for f in value])
                                else:
                                    val_str = str(value) if value else ""
                                run.text = run.text.replace(tag, val_str)
    out = BytesIO()
    prs.save(out)
    return out.getvalue()

def generar_pdf(pptx_bytes):
    """Convierte PPTX a PDF usando LibreOffice (soffice)"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        tmp_path = tmp.name
    
    try:
        # Comando para Streamlit Cloud
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', 
                        '--outdir', os.path.dirname(tmp_path), tmp_path], check=True)
        pdf_path = tmp_path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f:
            pdf_data = f.read()
        os.remove(tmp_path)
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return pdf_data
    except Exception as e:
        st.error(f"Error en conversión PDF: {e}")
        return None

def generar_png(pdf_bytes):
    """Convierte la primera diapositiva (del PDF) a imagen PNG"""
    try:
        images = convert_from_bytes(pdf_bytes)
        if images:
            img_byte_arr = BytesIO()
            images[0].save(img_byte_arr, format='PNG')
            return img_byte_arr.getvalue()
    except Exception as e:
        st.error(f"Error en conversión PNG: {e}")
        return None

# --- INTERFAZ ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")
st.title("🚀 Generador de Postales y Reportes")

# [Aquí iría tu lógica de Sidebar para cargar registros de Airtable definida antes]

if not st.session_state.get('df_trabajo', pd.DataFrame()).empty:
    
    # 1. TABLA DE SELECCIÓN
    st.subheader("1. Selección de Registros")
    c1, c2 = st.columns([1, 5])
    if c1.button("✅ Todo"): st.session_state.df_trabajo["Seleccionar"] = True; st.rerun()
    if c1.button("❌ Nada"): st.session_state.df_trabajo["Seleccionar"] = False; st.rerun()

    df_editado = st.data_editor(
        st.session_state.df_trabajo,
        use_container_width=True,
        hide_index=True,
        column_config={"Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False)},
        disabled=[c for c in st.session_state.df_trabajo.columns if c != "Seleccionar"]
    )

    seleccionados = df_editado[df_editado["Seleccionar"] == True]

    if not seleccionados.empty:
        st.divider()
        
        # 2. CONFIGURACIÓN DE PLANTILLAS POR TIPO
        tipos_en_seleccion = seleccionados["Tipo"].unique()
        st.subheader("2. Configuración de Plantillas")
        
        config_ok = True
        for t in tipos_en_seleccion:
            if t not in st.session_state.mapping or not os.path.exists(st.session_state.mapping[t]):
                st.warning(f"Falta plantilla para el tipo: **{t}**")
                file = st.file_uploader(f"Subir PPTX para el tipo '{t}'", type="pptx", key=f"p_{t}")
                if file:
                    p_path = f"plantilla_{t}.pptx"
                    with open(p_path, "wb") as f:
                        f.write(file.getbuffer())
                    st.session_state.mapping[t] = p_path
                    with open(CONFIG_FILE, 'w') as f:
                        json.dump(st.session_state.mapping, f)
                    st.success(f"Plantilla para {t} guardada.")
                    st.rerun()
                config_ok = False
        
        if config_ok:
            st.success("✅ Todas las plantillas están vinculadas.")
            
            # 3. SELECCIÓN DE FORMATO FINAL
            st.divider()
            st.subheader("3. Formato de Salida")
            formato = st.radio("¿Qué deseas generar?", ["🖼️ Postales (PNG)", "📄 Reportes (PDF)", "🔄 Ambos (PNG y PDF)"], horizontal=True)

            if st.button("🔥 INICIAR PROCESAMIENTO MASIVO"):
                for idx, fila in seleccionados.iterrows():
                    with st.status(f"Procesando: {fila['Sucursal']} ({fila['Tipo']})", expanded=False) as status:
                        
                        # Obtener plantilla
                        path_p = st.session_state.mapping[fila["Tipo"]]
                        with open(path_p, "rb") as f:
                            p_bytes = f.read()
                        
                        # Procesar PowerPoint en memoria
                        pptx_res = procesar_pptx(p_bytes, fila.to_dict())
                        
                        # Generar PDF (base para ambos formatos)
                        pdf_res = generar_pdf(pptx_res)
                        
                        if pdf_res:
                            st.write(f"✅ Archivos listos para {fila['Sucursal']}")
                            col_a, col_b = st.columns(2)
                            
                            # Lógica según selección de formato
                            if "PDF" in formato or "Ambos" in formato:
                                col_a.download_button(f"📥 PDF - {fila['Sucursal']}", pdf_res, f"Reporte_{fila['Sucursal']}.pdf", key=f"pdf_{idx}")
                            
                            if "PNG" in formato or "Ambos" in formato:
                                png_res = generar_png(pdf_res)
                                if png_res:
                                    col_b.download_button(f"📥 PNG - {fila['Sucursal']}", png_res, f"Postal_{fila['Sucursal']}.png", key=f"png_{idx}")
                        
                        status.update(label=f"Completado: {fila['Sucursal']}", state="complete")
