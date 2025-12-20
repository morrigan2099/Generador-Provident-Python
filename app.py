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
# He insertado tu token aquí para que sea automático
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
CONFIG_FILE = "config_plantillas.json"

# --- ESTADOS DE SESIÓN ---
if 'mapping' not in st.session_state:
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            st.session_state.mapping = json.load(f)
    else:
        st.session_state.mapping = {}

# --- FUNCIONES DE PROCESAMIENTO ---
def limpiar_adjuntos(valor):
    if isinstance(valor, list):
        return ", ".join([f.get("filename", "") for f in valor])
    return str(valor) if valor else ""

def procesar_pptx(plantilla_bytes, fields):
    prs = Presentation(BytesIO(plantilla_bytes))
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
    
    # Carga de Bases
    headers = {"Authorization": f"Bearer {TOKEN}"}
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    
    if r_bases.status_code == 200:
        bases = r_bases.json().get("bases", [])
        base_opts = {b['name']: b['id'] for b in bases}
        base_sel = st.selectbox("Base detectada:", list(base_opts.keys()))
        base_id = base_opts[base_sel]
        
        # Carga de Tablas
        r_tablas = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_id}/tables", headers=headers)
        if r_tablas.status_code == 200:
            tablas = r_tablas.json().get("tables", [])
            tabla_opts = {t['name']: t['id'] for t in tablas}
            tabla_sel = st.selectbox("Tabla detectada:", list(tabla_opts.keys()))
            tabla_id = tabla_opts[tabla_sel]
            
            # Carga automática de registros
            r_regs = requests.get(f"https://api.airtable.com/v0/{base_id}/{tabla_id}", headers=headers)
            registros_raw = r_regs.json().get("records", [])
            
            # Formatear tabla
            raw_fields = [r['fields'] for r in registros_raw]
            df = pd.DataFrame(raw_fields)
            
            # Columnas requeridas
            columnas_orden = ["Tipo", "Sucursal", "Seccion", "Punto de reunion", "Municipio", "Fecha", "Hora"]
            cols_finales = [c for c in columnas_orden if c in df.columns]
            df_display = df[cols_finales].copy()
            for col in df_display.columns: df_display[col] = df_display[col].apply(limpiar_adjuntos)
            df_display.insert(0, "Seleccionar", False)
            st.session_state.df_trabajo = df_display
            st.session_state.registros_raw = registros_raw

# --- PANEL PRINCIPAL ---
if not st.session_state.get('df_trabajo', pd.DataFrame()).empty:
    st.subheader("1. Selección de Registros")
    
    c1, c2, c3 = st.columns([1, 1, 4])
    if c1.button("✅ Seleccionar Todo"): st.session_state.df_trabajo["Seleccionar"] = True; st.rerun()
    if c2.button("❌ Desmarcar Todo"): st.session_state.df_trabajo["Seleccionar"] = False; st.rerun()

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
        
        # 2. CONFIGURACIÓN DE PLANTILLAS
        tipos = seleccionados["Tipo"].unique()
        st.subheader("2. Configuración de Plantillas")
        config_ok = True
        for t in tipos:
            if t not in st.session_state.mapping or not os.path.exists(st.session_state.mapping[t]):
                st.warning(f"Sube plantilla para el tipo: **{t}**")
                file = st.file_uploader(f"PPTX para '{t}'", type="pptx", key=f"p_{t}")
                if file:
                    p_path = f"plantilla_{t}.pptx"
                    with open(p_path, "wb") as f: f.write(file.getbuffer())
                    st.session_state.mapping[t] = p_path
                    with open(CONFIG_FILE, 'w') as f: json.dump(st.session_state.mapping, f)
                    st.rerun()
                config_ok = False
        
        if config_ok:
            st.success("✅ Plantillas listas.")
            
            # 3. SELECCIÓN DE FORMATO FINAL (Lo que faltaba)
            st.divider()
            st.subheader("3. Formato de Salida Final")
            formato = st.radio(
                "¿Qué deseas generar para los seleccionados?",
                ["🖼️ Postales (PNG)", "📄 Reportes (PDF)", "🔄 Ambos (PNG y PDF)"],
                horizontal=True
            )

            if st.button("🔥 INICIAR GENERACIÓN MASIVA"):
                for idx, fila in seleccionados.iterrows():
                    with st.status(f"Generando {fila['Sucursal']}...", expanded=False):
                        # Cargar datos originales del registro
                        datos_originales = st.session_state.registros_raw[idx]['fields']
                        path_p = st.session_state.mapping[fila["Tipo"]]
                        
                        with open(path_p, "rb") as f: p_bytes = f.read()
                        
                        pptx_res = procesar_pptx(p_bytes, datos_originales)
                        pdf_res = generar_pdf(pptx_res)
                        
                        if pdf_res:
                            col_a, col_b = st.columns(2)
                            if "PDF" in formato or "Ambos" in formato:
                                col_a.download_button(f"📥 PDF - {fila['Sucursal']}", pdf_res, f"Reporte_{fila['Sucursal']}.pdf", key=f"pdf_{idx}")
                            if "PNG" in formato or "Ambos" in formato:
                                png_res = generar_png(pdf_res)
                                if png_res:
                                    col_b.download_button(f"📥 PNG - {fila['Sucursal']}", png_res, f"Postal_{fila['Sucursal']}.png", key=f"png_{idx}")
