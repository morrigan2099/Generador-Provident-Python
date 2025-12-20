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
# Ruta base según tu estructura: Carpeta raíz 'Plantillas'
BASE_DIR = "Plantillas"

# --- FUNCIONES DE APOYO ---
def limpiar_adjuntos(valor):
    """Limpia los datos de Airtable, especialmente si son listas de adjuntos"""
    if isinstance(valor, list):
        return ", ".join([f.get("filename", "") for f in valor])
    return str(valor) if valor else ""

def generar_pdf(pptx_bytes):
    """Convierte PPTX a PDF usando LibreOffice (Headless)"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        tmp_path = tmp.name
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', 
                        '--outdir', os.path.dirname(tmp_path), tmp_path], check=True)
        pdf_path = tmp_path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f:
            pdf_data = f.read()
        # Limpiar temporales
        if os.path.exists(tmp_path): os.remove(tmp_path)
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return pdf_data
    except Exception as e:
        st.error(f"Error al generar PDF: {e}")
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
        st.error(f"Error al generar PNG: {e}")
        return None

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")
st.title("🚀 Generador de Postales y Reportes")

# --- CARGA AUTOMÁTICA DE AIRTABLE (BARRA LATERAL) ---
with st.sidebar:
    st.header("🔑 Conexión Airtable")
    headers = {"Authorization": f"Bearer {TOKEN}"}
    
    # 1. Obtener Bases
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        bases = r_bases.json().get("bases", [])
        base_opts = {b['name']: b['id'] for b in bases}
        base_sel = st.selectbox("Selecciona Base:", list(base_opts.keys()))
        
        # 2. Obtener Tablas
        r_tablas = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
        if r_tablas.status_code == 200:
            tablas = r_tablas.json().get("tables", [])
            tabla_opts = {t['name']: t['id'] for t in tablas}
            tabla_sel = st.selectbox("Selecciona Tabla:", list(tabla_opts.keys()))
            
            # 3. Obtener Registros
            r_regs = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
            st.session_state.registros_raw = r_regs.json().get("records", [])
            
            # Preparar DataFrame para visualización
            df_raw = pd.DataFrame([r['fields'] for r in st.session_state.registros_raw])
            cols_v = [c for c in ["Tipo", "Sucursal", "Seccion", "Municipio", "Fecha"] if c in df_raw.columns]
            df_display = df_raw[cols_v].copy()
            for col in df_display.columns:
                df_display[col] = df_display[col].apply(limpiar_adjuntos)
            
            if "Seleccionar" not in df_display.columns:
                df_display.insert(0, "Seleccionar", False)
            st.session_state.df_trabajo = df_display
    else:
        st.error("No se pudo conectar con Airtable. Verifica el Token.")

# --- PANEL PRINCIPAL ---
if not st.session_state.get('df_trabajo', pd.DataFrame()).empty:
    st.subheader("1. Selección de Registros")
    
    c1, c2, _ = st.columns([1, 1, 4])
    if c1.button("✅ Seleccionar Todo"):
        st.session_state.df_trabajo["Seleccionar"] = True
        st.rerun()
    if c2.button("❌ Deseleccionar Todo"):
        st.session_state.df_trabajo["Seleccionar"] = False
        st.rerun()

    # Editor de tabla para selección manual
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
        st.subheader("2. Selección de Formato y Plantillas del Repositorio")
        
        # PASO A: Seleccionar uso final (Define la subcarpeta)
        uso_final = st.radio("¿Qué deseas generar?", ["POSTALES", "REPORTES"], horizontal=True)
        uso_label = "Postal" if uso_final == "POSTALES" else "Reporte"
        
        # Ruta dinámica: Plantillas/POSTALES o Plantillas/REPORTES
        folder_path = os.path.join(BASE_DIR, uso_final)
        
        # PASO B: Escanear carpeta seleccionada
        if os.path.exists(folder_path):
            archivos_pptx = [f for f in os.listdir(folder_path) if f.endswith('.pptx')]
        else:
            archivos_pptx = []
            st.error(f"⚠️ La carpeta '{folder_path}' no existe en el repositorio.")

        if archivos_pptx:
            # PASO C: Selección manual de plantilla por cada TIPO único seleccionado
            mapping_manual = {}
            tipos_unicos = seleccionados["Tipo"].unique()
            
            st.info(f"Asigna una plantilla para cada tipo de {uso_label}:")
            cols_sel = st.columns(len(tipos_unicos) if len(tipos_unicos) < 3 else 3)
            
            for i, t in enumerate(tipos_unicos):
                with cols_sel[i % 3]:
                    mapping_manual[t] = st.selectbox(
                        f"Plantilla para {uso_label} TIPO: {t}",
                        options=archivos_pptx,
                        key=f"sel_{uso_final}_{t}"
                    )
            
            st.divider()
            # PASO D: Generación de archivos
            if st.button(f"🔥 INICIAR GENERACIÓN DE {uso_final}"):
                for idx, fila in seleccionados.iterrows():
                    with st.status(f"Procesando: {fila['Sucursal']} ({fila['Tipo']})...", expanded=False):
                        
                        # Obtener datos de Airtable
                        datos_record = st.session_state.registros_raw[idx]['fields']
                        # Obtener archivo elegido manualmente
                        nombre_archivo = mapping_manual[fila['Tipo']]
                        full_path = os.path.join(folder_path, nombre_archivo)
                        
                        # Lógica de reemplazo en PPTX
                        prs = Presentation(full_path)
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    for paragraph in shape.text_frame.paragraphs:
                                        for run in paragraph.runs:
                                            for k, v in datos_record.items():
                                                tag = f"{{{{{k}}}}}"
                                                if tag in run.text:
                                                    run.text = run.text.replace(tag, limpiar_adjuntos(v))
                        
                        # Guardar resultado temporal
                        pptx_io = BytesIO()
                        prs.save(pptx_io)
                        pdf_data = generar_pdf(pptx_io.getvalue())
                        
                        if pdf_data:
                            if uso_final == "REPORTES":
                                st.download_button(
                                    label=f"📥 Descargar PDF - {fila['Sucursal']}",
                                    data=pdf_data,
                                    file_name=f"Reporte_{fila['Sucursal']}.pdf",
                                    key=f"dl_pdf_{idx}"
                                )
                            else:
                                png_data = generar_png(pdf_data)
                                if png_data:
                                    st.download_button(
                                        label=f"📥 Descargar PNG - {fila['Sucursal']}",
                                        data=png_data,
                                        file_name=f"Postal_{fila['Sucursal']}.png",
                                        key=f"dl_png_{idx}"
                                    )
        else:
            st.warning(f"No hay archivos .pptx en la carpeta '{folder_path}'")
else:
    st.info("Conectando con Airtable para mostrar los registros...")
