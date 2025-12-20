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

# --- CONFIGURACIÓN FIJA ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"

# --- FUNCIONES DE PROCESAMIENTO ---
def limpiar_adjuntos(valor):
    if isinstance(valor, list):
        return ", ".join([f.get("filename", "") for f in valor])
    return str(valor) if valor else ""

def procesar_pptx(plantilla_path, fields):
    """Carga la plantilla desde el repositorio y reemplaza etiquetas"""
    if not os.path.exists(plantilla_path):
        return None
        
    prs = Presentation(plantilla_path)
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for key, value in fields.items():
                            tag = f"{{{{{key}}}}}"
                            if tag in run.text:
                                val_str = limpiar_adjuntos(value)
                                run.text = run.text.replace(tag, val_str)
    out = BytesIO()
    prs.save(out)
    return out.getvalue()

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
        os.remove(tmp_path)
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return pdf_data
    except Exception as e:
        st.error(f"Error en PDF: {e}")
        return None

def generar_png(pdf_bytes):
    try:
        images = convert_from_bytes(pdf_bytes)
        if images:
            img_byte_arr = BytesIO()
            images[0].save(img_byte_arr, format='PNG')
            return img_byte_arr.getvalue()
    except Exception as e:
        st.error(f"Error en PNG: {e}")
        return None

# --- INTERFAZ ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")
st.title("🚀 Generador de Postales y Reportes")

# --- CARGA AUTOMÁTICA EN BARRA LATERAL ---
with st.sidebar:
    st.header("🔑 Conexión Automática")
    headers = {"Authorization": f"Bearer {TOKEN}"}
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    
    if r_bases.status_code == 200:
        bases = r_bases.json().get("bases", [])
        base_opts = {b['name']: b['id'] for b in bases}
        base_sel = st.selectbox("Base:", list(base_opts.keys()))
        
        r_tablas = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
        if r_tablas.status_code == 200:
            tablas = r_tablas.json().get("tables", [])
            tabla_opts = {t['name']: t['id'] for t in tablas}
            tabla_sel = st.selectbox("Tabla:", list(tabla_opts.keys()))
            
            # Carga automática de registros
            r_regs = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
            st.session_state.registros_raw = r_regs.json().get("records", [])
            
            raw_fields = [r['fields'] for r in st.session_state.registros_raw]
            df = pd.DataFrame(raw_fields)
            columnas_orden = ["Tipo", "Sucursal", "Seccion", "Municipio", "Fecha"]
            cols_finales = [c for c in columnas_orden if c in df.columns]
            df_display = df[cols_finales].copy()
            for col in df_display.columns: df_display[col] = df_display[col].apply(limpiar_adjuntos)
            df_display.insert(0, "Seleccionar", False)
            st.session_state.df_trabajo = df_display

# --- PANEL PRINCIPAL ---
if not st.session_state.get('df_trabajo', pd.DataFrame()).empty:
    st.subheader("1. Selecciona los registros")
    
    c1, c2, _ = st.columns([1, 1, 4])
    if c1.button("✅ Todo"): st.session_state.df_trabajo["Seleccionar"] = True; st.rerun()
    if c2.button("❌ Nada"): st.session_state.df_trabajo["Seleccionar"] = False; st.rerun()

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
        
        # --- NUEVA SECCIÓN: SELECCIÓN DE FORMATO ANTES DE GENERAR ---
        st.subheader("2. Opciones de Formato Final")
        formato_final = st.radio(
            "¿Qué archivos deseas generar para los registros seleccionados?",
            ["🖼️ Solo Postales (PNG)", "📄 Solo Reportes (PDF)", "🔄 Ambos (PNG y PDF)"],
            horizontal=True
        )

        st.info("💡 La aplicación usará automáticamente las plantillas almacenadas en el repositorio según el campo 'Tipo'.")

        if st.button("🔥 GENERAR ARCHIVOS"):
            for idx, fila in seleccionados.iterrows():
                with st.status(f"Procesando {fila['Sucursal']}...", expanded=False):
                    
                    # Buscar la plantilla en el repositorio
                    # Asumimos que los archivos se llaman exactamente como el "Tipo" (ej: Mensual.pptx)
                    nombre_plantilla = f"{fila['Tipo']}.pptx"
                    
                    if not os.path.exists(nombre_plantilla):
                        st.error(f"No se encontró el archivo '{nombre_plantilla}' en el repositorio.")
                        continue

                    # Obtener datos originales y procesar
                    datos_originales = st.session_state.registros_raw[idx]['fields']
                    pptx_res = procesar_pptx(nombre_plantilla, datos_originales)
                    
                    if pptx_res:
                        pdf_res = generar_pdf(pptx_res)
                        
                        if pdf_res:
                            st.write(f"✅ Archivos listos para **{fila['Sucursal']}**")
                            col_a, col_b = st.columns(2)
                            
                            if "PDF" in formato_final or "Ambos" in formato_final:
                                col_a.download_button(f"📥 Reporte (PDF) - {fila['Sucursal']}", pdf_res, f"Reporte_{fila['Sucursal']}.pdf", key=f"pdf_{idx}")
                            
                            if "PNG" in formato_final or "Ambos" in formato_final:
                                png_res = generar_png(pdf_res)
                                if png_res:
                                    col_b.download_button(f"📥 Postal (PNG) - {fila['Sucursal']}", png_res, f"Postal_{fila['Sucursal']}.png", key=f"png_{idx}")
