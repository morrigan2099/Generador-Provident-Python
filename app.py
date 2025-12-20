import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from io import BytesIO
import os
import tempfile
import subprocess
import cloudinary
import cloudinary.uploader

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")

# Inicialización de estados en sesión (Memoria de la App)
if 'df_trabajo' not in st.session_state: st.session_state.df_trabajo = pd.DataFrame()
if 'registros_raw' not in st.session_state: st.session_state.registros_raw = []
if 'bases' not in st.session_state: st.session_state.bases = []
if 'tablas' not in st.session_state: st.session_state.tablas = []

# --- FUNCIONES DE APOYO ---
def limpiar_adjuntos(valor):
    """Extrae el nombre del archivo de los objetos de Airtable para la tabla"""
    if isinstance(valor, list):
        return ", ".join([f.get("filename", "archivo") for f in valor])
    return str(valor) if valor else ""

def procesar_pptx(plantilla_bytes, fields):
    """Lógica de reemplazo de etiquetas {{Campo}} en PowerPoint"""
    prs = Presentation(BytesIO(plantilla_bytes))
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for key, value in fields.items():
                            val_str = limpiar_adjuntos(value)
                            tag = f"{{{{{key}}}}}"
                            if tag in run.text:
                                run.text = run.text.replace(tag, val_str)
    out = BytesIO()
    prs.save(out)
    return out.getvalue()

# --- BARRA LATERAL (CONEXIÓN Y CONFIGURACIÓN) ---
with st.sidebar:
    st.header("🔑 Conexión Airtable")
    token = st.text_input("Airtable Token", type="password")
    
    if st.button("🔄 Cargar Bases"):
        headers = {"Authorization": f"Bearer {token}"}
        r = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
        if r.status_code == 200:
            st.session_state.bases = r.json().get("bases", [])
            st.success("Bases cargadas")
        else:
            st.error("Error: Token inválido o sin permisos")

    if st.session_state.bases:
        base_opts = {b['name']: b['id'] for b in st.session_state.bases}
        base_sel = st.selectbox("Selecciona Base", list(base_opts.keys()))
        
        if st.button("📂 Ver Tablas"):
            headers = {"Authorization": f"Bearer {token}"}
            r = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
            st.session_state.tablas = r.json().get("tables", [])
            st.session_state.base_id_activa = base_opts[base_sel]

    if st.session_state.tablas:
        tabla_opts = {t['name']: t['id'] for t in st.session_state.tablas}
        tabla_sel = st.selectbox("Selecciona Tabla", list(tabla_opts.keys()))
        
        if st.button("📥 CARGAR DATOS"):
            headers = {"Authorization": f"Bearer {token}"}
            url = f"https://api.airtable.com/v0/{st.session_state.base_id_activa}/{tabla_opts[tabla_sel]}"
            r = requests.get(url, headers=headers)
            regs = r.json().get("records", [])
            st.session_state.registros_raw = regs
            
            # Crear DataFrame con los campos solicitados
            raw_fields = [r['fields'] for r in regs]
            df = pd.DataFrame(raw_fields)
            
            columnas_orden = [
                "Tipo", "Sucursal", "Seccion", "Punto de reunion", "Ruta a seguir", 
                "Municipio", "Fecha", "Hora", "Am responsable", "Teléfono AM", 
                "Dm responsable", "Teléfono DM", "Foto de equipo", "Foto 01", 
                "Foto 02", "Reporte firmado", "Lista de asistencia"
            ]
            
            cols_finales = [c for c in columnas_orden if c in df.columns]
            df_filtrado = df[cols_finales].copy()
            
            # Limpiar visualización de fotos/adjuntos
            for col in df_filtrado.columns:
                df_filtrado[col] = df_filtrado[col].apply(limpiar_adjuntos)
            
            # Insertar columna de selección al inicio
            df_filtrado.insert(0, "Seleccionar", False)
            st.session_state.df_trabajo = df_filtrado

# --- CUERPO PRINCIPAL ---
st.title("🚀 Generador de Presentaciones Provident")

if not st.session_state.df_trabajo.empty:
    st.subheader("Registros Cargados")
    
    # Botones de selección masiva
    c1, c2, _ = st.columns([1, 1, 4])
    if c1.button("✅ Seleccionar Todo"):
        st.session_state.df_trabajo["Seleccionar"] = True
        st.rerun()
    if c2.button("❌ Desmarcar Todo"):
        st.session_state.df_trabajo["Seleccionar"] = False
        st.rerun()

    # Tabla interactiva con editor de datos
    df_editado = st.data_editor(
        st.session_state.df_trabajo,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Seleccionar": st.column_config.CheckboxColumn(
                "Seleccionar",
                help="Marca para procesar este registro",
                default=False,
            )
        },
        # Solo la columna "Seleccionar" es editable
        disabled=[c for c in st.session_state.df_trabajo.columns if c != "Seleccionar"]
    )

    st.divider()

    # Lógica de Generación
    seleccionados_indices = df_editado[df_editado["Seleccionar"] == True].index.tolist()
    
    if seleccionados_indices:
        st.success(f"Seleccionados: {len(seleccionados_indices)} registros.")
        plantilla = st.file_uploader("Sube tu plantilla PowerPoint (.pptx)", type="pptx")
        
        if st.button("🔥 GENERAR SELECCIONADOS") and plantilla:
            plantilla_bytes = plantilla.read()
            
            for idx in seleccionados_indices:
                # Obtenemos los datos originales (sin limpiar) del registro raw
                registro_full = st.session_state.registros_raw[idx]
                nombre_cliente = registro_full['fields'].get('Sucursal', f'Registro_{idx}')
                
                with st.status(f"Procesando {nombre_cliente}...", expanded=False):
                    pptx_final = procesar_pptx(plantilla_bytes, registro_full['fields'])
                    
                    st.write(f"✅ Archivo generado para {nombre_cliente}")
                    st.download_button(
                        label=f"📥 Descargar {nombre_cliente}",
                        data=pptx_final,
                        file_name=f"Reporte_{nombre_cliente}.pptx",
                        key=f"dl_{idx}"
                    )
    else:
        st.warning("Selecciona al menos un registro de la tabla arriba para continuar.")

else:
    st.info("Utiliza el panel de la izquierda para conectar Airtable y cargar los registros.")
