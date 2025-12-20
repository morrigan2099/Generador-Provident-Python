import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from io import BytesIO
import cloudinary
import cloudinary.uploader
import os
import tempfile
import subprocess

# --- CONFIGURACIÓN DE ESTADO (La memoria de la App) ---
if 'bases' not in st.session_state:
    st.session_state.bases = []
if 'tablas' not in st.session_state:
    st.session_state.tablas = []
if 'registros' not in st.session_state:
    st.session_state.registros = []
if 'record_map' not in st.session_state:
    st.session_state.record_map = {}

# --- FUNCIONES DE CONEXIÓN ---
def fetch_bases(token):
    headers = {"Authorization": f"Bearer {token}"}
    url = "https://api.airtable.com/v0/meta/bases"
    r = requests.get(url, headers=headers)
    return r.json().get("bases", []) if r.status_code == 200 else []

def fetch_tablas(token, base_id):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://api.airtable.com/v0/meta/bases/{base_id}/tables"
    r = requests.get(url, headers=headers)
    return r.json().get("tables", []) if r.status_code == 200 else []

def fetch_records(token, base_id, table_id):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}"
    r = requests.get(url, headers=headers)
    return r.json().get("records", []) if r.status_code == 200 else []

# --- LÓGICA DE PROCESAMIENTO ---
def procesar_pptx(plantilla_bytes, fields):
    prs = Presentation(BytesIO(plantilla_bytes))
    # Reemplazo de texto simple {{Campo}}
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for key, value in fields.items():
                            tag = f"{{{{{key}}}}}"
                            if tag in run.text:
                                run.text = run.text.replace(tag, str(value))
    
    output = BytesIO()
    prs.save(output)
    return output.getvalue()

# --- INTERFAZ ---
st.set_page_config(page_title="Generador Provident", layout="wide")
st.title("🚀 Generador Provident - Versión Web")

# BARRA LATERAL
with st.sidebar:
    st.header("1. Conexión Airtable")
    token = st.text_input("Airtable Token", type="password")
    
    # Botón para cargar bases
    if st.button("Cargar Mis Bases"):
        if token:
            st.session_state.bases = fetch_bases(token)
            if not st.session_state.bases:
                st.error("No se encontraron bases. Revisa el Token.")
        else:
            st.warning("Ingresa un Token")

    # Selector de Bases
    if st.session_state.bases:
        base_names = {b['name']: b['id'] for b in st.session_state.bases}
        base_sel = st.selectbox("Selecciona Base", options=list(base_names.keys()))
        
        if st.button("Ver Tablas"):
            st.session_state.tablas = fetch_tablas(token, base_names[base_sel])
            st.session_state.base_id_activa = base_names[base_sel]

    # Selector de Tablas
    if st.session_state.tablas:
        tabla_names = {t['name']: t['id'] for t in st.session_state.tablas}
        tabla_sel = st.selectbox("Selecciona Tabla", options=list(tabla_names.keys()))
        
        if st.button("📥 CARGAR REGISTROS"):
            with st.spinner("Leyendo datos..."):
                regs = fetch_records(token, st.session_state.base_id_activa, tabla_names[tabla_sel])
                st.session_state.registros = regs
                # Mapeo por el campo 'Nombre' para el multiselect
                st.session_state.record_map = {r['fields'].get('Nombre', r['id']): r for r in regs}
                st.success(f"Cargados {len(regs)} registros.")

# ÁREA CENTRAL
if st.session_state.registros:
    st.subheader("Configuración de Generación")
    
    col1, col2 = st.columns(2)
    with col1:
        seleccionados = st.multiselect("Selecciona Clientes:", options=list(st.session_state.record_map.keys()))
    with col2:
        archivo_pptx = st.file_uploader("Sube tu plantilla (.pptx)", type="pptx")

    if seleccionados and archivo_pptx:
        if st.button("🔥 GENERAR PRESENTACIONES"):
            plantilla_bytes = archivo_pptx.read()
            
            for nombre in seleccionados:
                with st.expander(f"Resultado: {nombre}", expanded=True):
                    record_data = st.session_state.record_map[nombre]
                    
                    # Generar PPTX
                    pptx_res = procesar_pptx(plantilla_bytes, record_data['fields'])
                    
                    st.write(f"✅ Presentación para {nombre} lista.")
                    st.download_button(
                        label=f"⬇️ Descargar PPTX - {nombre}",
                        data=pptx_res,
                        file_name=f"Presentacion_{nombre}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
else:
    st.info("👈 Comienza configurando tu Token y seleccionando Base/Tabla en el panel izquierdo.")

# --- SECCIÓN CLOUDINARY (Opcional, configurada en sidebar) ---
with st.sidebar:
    st.divider()
    st.header("2. Cloudinary (Opcional)")
    c_name = st.text_input("Cloud Name")
    c_key = st.text_input("API Key")
    c_secret = st.text_input("API Secret", type="password")
    if c_name and c_key and c_secret:
        cloudinary.config(cloud_name=c_name, api_key=c_key, api_secret=c_secret)
        st.caption("✅ Cloudinary conectado")
