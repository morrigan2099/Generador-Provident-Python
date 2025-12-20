import streamlit as st
import requests
import pandas as pd # Usaremos pandas para mostrar la tabla bonito

# --- CONFIGURACIÓN Y ESTADO ---
st.set_page_config(page_title="Generador Provident", layout="wide")

if 'registros' not in st.session_state:
    st.session_state.registros = []
if 'record_map' not in st.session_state:
    st.session_state.record_map = {}

# --- FUNCIONES API ---
def obtener_bases(token):
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    return response.json().get("bases", []) if response.status_code == 200 else []

def obtener_tablas(token, base_id):
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_id}/tables", headers=headers)
    return response.json().get("tables", []) if response.status_code == 200 else []

def cargar_datos_airtable(token, base_id, table_id_or_name):
    # Airtable prefiere el ID de la tabla o el nombre codificado
    url = f"https://api.airtable.com/v0/{base_id}/{table_id_or_name}"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json().get("records", [])
    else:
        st.error(f"Error al cargar registros: {response.status_code}")
        return []

# --- BARRA LATERAL ---
st.sidebar.header("🔑 Configuración")
token = st.sidebar.text_input("Airtable Token", type="password")

if token:
    bases = obtener_bases(token)
    if bases:
        base_opts = {b['name']: b['id'] for b in bases}
        base_sel = st.sidebar.selectbox("Selecciona Base", list(base_opts.keys()))
        
        tablas = obtener_tablas(token, base_opts[base_sel])
        if tablas:
            tabla_opts = {t['name']: t['id'] for t in tablas}
            tabla_sel = st.sidebar.selectbox("Selecciona Tabla", list(tabla_opts.keys()))
            
            # BOTÓN CRÍTICO
            if st.sidebar.button("🔄 Cargar Registros"):
                with st.spinner("Descargando datos..."):
                    datos = cargar_datos_airtable(token, base_opts[base_sel], tabla_opts[tabla_sel])
                    if datos:
                        st.session_state.registros = datos
                        # Mapeamos por el campo 'Nombre' (asegúrate que exista en tu Airtable)
                        st.session_state.record_map = {
                            r['fields'].get('Nombre', r['id']): r for r in datos
                        }
                        st.sidebar.success(f"Cargados {len(datos)} registros")
                    else:
                        st.sidebar.warning("La tabla está vacía o el campo 'Nombre' no existe.")

# --- CUERPO PRINCIPAL ---
st.title("🚀 Panel de Control Provident")

if st.session_state.registros:
    st.subheader("Registros detectados en Airtable")
    
    # Convertimos los datos a un formato que Streamlit pueda mostrar en tabla (Dataframe)
    # Extraemos solo los campos de 'fields' de cada registro
    lista_para_tabla = [r['fields'] for r in st.session_state.registros]
    df = pd.DataFrame(lista_para_tabla)
    
    # Mostramos la tabla interactiva
    st.dataframe(df, use_container_width=True)

    st.divider()
    
    # Selector para procesar
    seleccionados = st.multiselect(
        "Selecciona los clientes para generar presentación:",
        options=list(st.session_state.record_map.keys())
    )
    
    if seleccionados:
        st.write(f"Has seleccionado {len(seleccionados)} registros.")
        # Aquí iría el resto de tu lógica de PowerPoint...
else:
    st.info("👈 Selecciona una base y tabla en la izquierda y presiona 'Cargar Registros'")
