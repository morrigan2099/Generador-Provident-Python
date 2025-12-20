import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from io import BytesIO
import os
import subprocess
import tempfile
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN FIJA ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"

# --- FUNCIONES DE APOYO ---
def limpiar_adjuntos(valor):
    if isinstance(valor, list):
        return ", ".join([f.get("filename", "") for f in valor])
    return str(valor) if valor else ""

def generar_pdf(pptx_bytes):
    """Convierte binario de PPTX a PDF usando LibreOffice"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        tmp_path = tmp.name
    try:
        # Comando para servidores Linux (Streamlit Cloud)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', 
                        '--outdir', os.path.dirname(tmp_path), tmp_path], check=True)
        pdf_path = tmp_path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f:
            pdf_data = f.read()
        os.remove(tmp_path)
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return pdf_data
    except Exception as e:
        st.error(f"Error en conversión a PDF: {e}")
        return None

def generar_png(pdf_bytes):
    """Convierte la primera página de un PDF a imagen PNG"""
    try:
        images = convert_from_bytes(pdf_bytes)
        if images:
            img_byte_arr = BytesIO()
            images[0].save(img_byte_arr, format='PNG')
            return img_byte_arr.getvalue()
    except Exception as e:
        st.error(f"Error en conversión a PNG: {e}")
        return None

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")
st.title("🚀 Generador de Postales y Reportes")

# --- CARGA AUTOMÁTICA (BARRA LATERAL) ---
with st.sidebar:
    st.header("🔑 Conexión Airtable")
    headers = {"Authorization": f"Bearer {TOKEN}"}
    
    # 1. Obtener Bases automáticamente
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        bases = r_bases.json().get("bases", [])
        base_opts = {b['name']: b['id'] for b in bases}
        base_sel = st.selectbox("Selecciona Base:", list(base_opts.keys()))
        
        # 2. Obtener Tablas automáticamente
        r_tablas = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
        if r_tablas.status_code == 200:
            tablas = r_tablas.json().get("tables", [])
            tabla_opts = {t['name']: t['id'] for t in tablas}
            tabla_sel = st.selectbox("Selecciona Tabla:", list(tabla_opts.keys()))
            
            # 3. Cargar Registros automáticamente
            r_regs = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
            st.session_state.registros_raw = r_regs.json().get("records", [])
            
            # Crear DataFrame de visualización
            raw_data = [r['fields'] for r in st.session_state.registros_raw]
            df = pd.DataFrame(raw_data)
            
            # Columnas a mostrar
            cols_deseadas = ["Tipo", "Sucursal", "Seccion", "Municipio", "Fecha", "Hora"]
            cols_visibles = [c for c in cols_deseadas if c in df.columns]
            df_display = df[cols_visibles].copy()
            
            # Limpiar datos para la tabla
            for col in df_display.columns:
                df_display[col] = df_display[col].apply(limpiar_adjuntos)
            
            # Insertar columna de selección
            if "Seleccionar" not in df_display.columns:
                df_display.insert(0, "Seleccionar", False)
            
            st.session_state.df_trabajo = df_display
    else:
        st.error("Error al conectar con Airtable. Verifica el Token.")

# --- PANEL PRINCIPAL ---
if not st.session_state.get('df_trabajo', pd.DataFrame()).empty:
    st.subheader("1. Selección de Registros")
    
    # Botones de selección masiva
    c1, c2, _ = st.columns([1, 1, 4])
    if c1.button("✅ Todo"):
        st.session_state.df_trabajo["Seleccionar"] = True
        st.rerun()
    if c2.button("❌ Nada"):
        st.session_state.df_trabajo["Seleccionar"] = False
        st.rerun()

    # Tabla interactiva
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
        st.subheader("2. Selección de Formato y Plantillas")
        
        # Selección de USO
        uso_final = st.radio("Seleccione uso final:", ["Postal", "Reporte"], horizontal=True)
        
        # Listar archivos PPTX del repositorio
        archivos_repo = [f for f in os.listdir('.') if f.endswith('.pptx')]
        
        if not archivos_repo:
            st.error("⚠️ No se encontraron archivos .pptx en el repositorio.")
        else:
            plantillas_map = {}
            tipos_unicos = seleccionados["Tipo"].unique()
            
            # Cuadro de diálogo interactivo por cada TIPO
            for t in tipos_unicos:
                plantillas_map[t] = st.selectbox(
                    f"Seleccione plantilla para {uso_final} de TIPO: {t}",
                    options=archivos_repo,
                    key=f"sel_{t}"
                )
            
            st.divider()
            if st.button(f"🔥 GENERAR {uso_final.upper()}S"):
                for idx, fila in seleccionados.iterrows():
                    with st.status(f"Procesando {fila['Sucursal']}...", expanded=False):
                        
                        # Datos originales de Airtable
                        datos_record = st.session_state.registros_raw[idx]['fields']
                        archivo_pptx = plantillas_map[fila['Tipo']]
                        
                        # Cargar plantilla y reemplazar
                        prs = Presentation(archivo_pptx)
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    for paragraph in shape.text_frame.paragraphs:
                                        for run in paragraph.runs:
                                            for k, v in datos_record.items():
                                                tag = f"{{{{{k}}}}}"
                                                if tag in run.text:
                                                    run.text = run.text.replace(tag, limpiar_adjuntos(v))
                        
                        # Guardar PPTX temporalmente en memoria
                        pptx_io = BytesIO()
                        prs.save(pptx_io)
                        
                        # Convertir a PDF (Base para ambos usos)
                        pdf_data = generar_pdf(pptx_io.getvalue())
                        
                        if pdf_data:
                            if uso_final == "Reporte":
                                st.download_button(
                                    label=f"📥 PDF - {fila['Sucursal']}",
                                    data=pdf_data,
                                    file_name=f"Reporte_{fila['Sucursal']}.pdf",
                                    key=f"dl_pdf_{idx}"
                                )
                            else:
                                png_data = generar_png(pdf_data)
                                if png_data:
                                    st.download_button(
                                        label=f"📥 PNG - {fila['Sucursal']}",
                                        data=png_data,
                                        file_name=f"Postal_{fila['Sucursal']}.png",
                                        key=f"dl_png_{idx}"
                                    )
else:
    st.info("Esperando conexión con Airtable para mostrar registros...")
