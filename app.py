import streamlit as st
import requests
import pandas as pd
import json
import os
import re
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import subprocess, tempfile, zipfile
from io import BytesIO
from pdf2image import convert_from_bytes

# --- FUNCIONES NUEVAS PARA EVITAR CACHÉ ---

def limpiar_campo_texto(valor, campo=""):
    """Nueva función para evitar conflictos con versiones anteriores"""
    if not valor or str(valor).lower() == "none": 
        return ""
    if isinstance(valor, list): 
        return valor
    
    txt = str(valor).strip()
    if campo == 'Seccion': 
        return txt.upper()
    
    # Capitalización estándar
    palabras = txt.lower().split()
    if not palabras: 
        return ""
    
    preposiciones = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al']
    resultado = [palabras[0].capitalize()]
    for p in palabras[1:]:
        resultado.append(p if p in preposiciones else p.capitalize())
    return " ".join(resultado)

def convertir_a_pdf(pptx_data):
    """Función de conversión aislada"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_data)
        p_in = tmp.name
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(p_in), p_in], check=True)
        p_out = p_in.replace(".pptx", ".pdf")
        with open(p_out, "rb") as f: 
            data = f.read()
        if os.path.exists(p_in): os.remove(p_in)
        if os.path.exists(p_out): os.remove(p_out)
        return data
    except: 
        return None

# --- INICIALIZACIÓN ---
st.set_page_config(page_title="Provident Pro v43", layout="wide")

if 'config' not in st.session_state:
    if os.path.exists("config_app.json"):
        try:
            with open("config_app.json", "r") as f: 
                st.session_state.config = json.load(f)
        except:
            st.session_state.config = {"plantillas": {}, "columnas_visibles": []}
    else:
        st.session_state.config = {"plantillas": {}, "columnas_visibles": []}

st.title("🚀 Generador Pro v43")

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configuración")
    if st.button("💾 GUARDAR JSON"):
        with open("config_app.json", "w") as f: 
            json.dump(st.session_state.config, f)
        st.toast("Guardado correctamente")

    # Airtable Logic
    try:
        r_b = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
        if r_b.status_code == 200:
            bases = {b['name']: b['id'] for b in r_b.json().get('bases', [])}
            b_sel = st.selectbox("Seleccionar Base:", [""] + list(bases.keys()))
            if b_sel:
                bid = bases[b_sel]
                r_t = requests.get(f"https://api.airtable.com/v0/meta/bases/{bid}/tables", headers=headers)
                tablas = {t['name']: t['id'] for t in r_t.json().get('tables', [])}
                t_sel = st.selectbox("Seleccionar Tabla:", list(tablas.keys()))
                
                if st.button("🔄 CARGAR REGISTROS"):
                    r_r = requests.get(f"https://api.airtable.com/v0/{bid}/{tablas[t_sel]}", headers=headers)
                    recs = r_r.json().get("records", [])
                    st.session_state.raw_data_original = recs
                    st.session_state.raw_records = [
                        {'id': r['id'], 'fields': {k: (limpiar_campo_texto(v, k) if k != 'Fecha' else v) for k, v in r['fields'].items()}} 
                        for r in recs
                    ]
                    st.rerun()
    except Exception as e:
        st.error(f"Error de conexión: {e}")

# --- TABLA Y PROCESAMIENTO ---
if 'raw_records' in st.session_state:
    df_base = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    
    with st.sidebar:
        st.divider()
        todas_las_cols = list(df_base.columns)
        default_cols = [c for c in st.session_state.config.get("columnas_visibles", []) if c in todas_las_cols] or todas_las_cols
        cols_v = st.multiselect("Campos visibles:", todas_las_cols, default=default_cols)
        st.session_state.config["columnas_visibles"] = cols_v

    # Limpieza para tabla
    df_view = df_base[[c for c in cols_v if c in df_base.columns]].copy()
    for col in df_view.columns:
        if not df_view.empty and isinstance(df_view[col].iloc[0], list):
            df_view.drop(col, axis=1, inplace=True)
    
    df_view.insert(0, "Seleccionar", False)
    
    # Editor con Checkbox Maestro
    df_edit = st.data_editor(
        df_view, 
        use_container_width=True, 
        hide_index=True, 
        column_config={"Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False)}
    )
    
    idx_sel = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if idx_sel:
        modo = st.radio("Acción:", ["Postales", "Reportes"], horizontal=True)
        fol_plantillas = os.path.join("Plantillas", modo.upper())
        AZUL_CORP = RGBColor(0, 176, 240)
        
        # Mapeo de plantillas por Tipo
        tipos_en_sel = df_view.loc[idx_sel, "Tipo"].unique()
        archivos_disp = [f for f in os.listdir(fol_plantillas) if f.endswith('.pptx')]
        
        for t in tipos_en_sel:
            p_guardada = st.session_state.config.get("plantillas", {}).get(t)
            idx_def = archivos_disp.index(p_guardada) if p_guardada in archivos_disp else 0
            st.session_state.config.setdefault("plantillas", {})[t] = st.selectbox(f"Plantilla para {t}:", archivos_disp, index=idx_def)

        if st.button("🔥 GENERAR SELECCIÓN", use_container_width=True, type="primary"):
            progress = st.progress(0)
            zip_buf = BytesIO()
            
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zf:
                for i, idx in enumerate(idx_sel):
                    r_proc = st.session_state.raw_records[idx]['fields']
                    r_orig = st.session_state.raw_data_original[idx]['fields']
                    
                    # Carga y reemplazo
                    prs = Presentation(os.path.join(fol_plantillas, st.session_state.config["plantillas"][r_proc['Tipo']]))
                    
                    data_tags = {
                        "<<Tipo>>": r_proc.get('Tipo'),
                        "<<Sucursal>>": r_proc.get('Sucursal'),
                        "<<Seccion>>": r_proc.get('Seccion'),
                        "<<Confechor>>": f"{r_proc.get('Fecha')}, {r_proc.get('Hora')}",
                        "<<Concat>>": f"{r_proc.get('Punto de reunion') or r_proc.get('Ruta a seguir')}, {r_proc.get('Municipio')}"
                    }

                    for slide in prs.slides:
                        # Procesar Imágenes
                        for shp in list(slide.shapes):
                            txt_shp = shp.text_frame.text if shp.has_text_frame else ""
                            tags_fotos = ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]
                            for tag in tags_fotos:
                                if f"<<{tag}>>" in txt_shp or tag == shp.name:
                                    adj = r_orig.get(tag)
                                    if adj and isinstance(adj, list):
                                        try:
                                            img_bytes = requests.get(adj[0]['url']).content
                                            slide.shapes.add_picture(BytesIO(img_bytes), shp.left, shp.top, shp.width, shp.height)
                                            shp.element.getparent().remove(shp.element)
                                        except: pass

                        # Procesar Texto
                        for shp in slide.shapes:
                            if shp.has_text_frame:
                                for tag, val in data_tags.items():
                                    if tag in shp.text_frame.text:
                                        tf = shp.text_frame
                                        tf.clear()
                                        run = tf.paragraphs[0].add_run()
                                        run.text = str(val)
                                        run.font.bold = True
                                        run.font.color.rgb = AZUL_CORP
                                        # TAMAÑOS FIJOS SOLICITADOS
                                        if tag == "<<Tipo>>": 
                                            run.font.size = Pt(11) # <--- SIEMPRE 11pts
                                        elif tag == "<<Sucursal>>": 
                                            run.font.size = Pt(14)
                                        else: 
                                            run.font.size = Pt(11)

                    # Exportación
                    pp_out = BytesIO()
                    prs.save(pp_out)
                    pdf_res = convertir_a_pdf(pp_out.getvalue())
                    
                    if pdf_res:
                        ext = ".pdf" if modo == "Reportes" else ".jpg"
                        nombre_final = f"{r_proc.get('Fecha')} - {r_proc.get('Tipo')} - {r_proc.get('Sucursal')}{ext}"
                        # Carpeta por sucursal dentro del ZIP
                        zf.writestr(f"{modo}/{r_proc.get('Sucursal')}/{nombre_final}", pdf_res if modo == "Reportes" else convert_from_bytes(pdf_res)[0].tobytes())
                    
                    progress.progress((i + 1) / len(idx_sel))

            st.success("✅ Generación terminada")
            st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_v43.zip", use_container_width=True)
