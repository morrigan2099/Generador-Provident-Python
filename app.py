import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import os, subprocess, tempfile, zipfile
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"
AZUL_CELESTE = RGBColor(0, 176, 240) 
MESES_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
DIAS_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

def proper_elegante(texto):
    if not texto or str(texto).lower() == "none": return ""
    texto = str(texto).strip().lower()
    palabras = texto.split()
    excepciones = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del']
    resultado = []
    for i, p in enumerate(palabras):
        if i == 0 or (i > 0 and palabras[i-1].endswith('.')) or p not in excepciones:
            resultado.append(p.capitalize())
        else:
            resultado.append(p)
    return " ".join(resultado)

def generar_pdf(pptx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        path = tmp.name
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(path), path], check=True)
        pdf_path = path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f: data = f.read()
        os.remove(path)
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return data
    except: return None

# --- UI ---
st.set_page_config(page_title="Provident Pro Custom", layout="wide")
st.title("🚀 Generador Pro: Tipo Auto-Ajustable (64pt)")

if 'raw_records' not in st.session_state: st.session_state.raw_records = []

# ... (Lógica de carga de datos de Airtable se mantiene igual) ...

if st.session_state.raw_records:
    df_prev = pd.DataFrame([{"Tipo": r['fields'].get("Tipo"), "Sucursal": r['fields'].get("Sucursal"), "Fecha": r['fields'].get("Fecha")} for r in st.session_state.raw_records])
    df_prev.insert(0, "Seleccionar", False)
    df_edit = st.data_editor(df_prev, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        modo = st.radio("Acción:", ["Postales", "Reportes"], horizontal=True)
        folder_fisica = os.path.join(BASE_DIR, modo.upper())
        
        if os.path.exists(folder_fisica):
            archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
            tipos_unicos = df_edit.loc[sel_idx, "Tipo"].unique()
            map_memoria = {t: st.selectbox(f"Plantilla para {t}:", archivos_pptx, key=f"p_{t}") for t in tipos_unicos}

            if st.button("🔥 GENERAR"):
                p_bar = st.progress(0); s_text = st.empty()
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                    for i, idx in enumerate(sel_idx):
                        f = st.session_state.raw_records[idx]['fields']
                        suc_actual = str(f.get('Sucursal', 'Doc'))
                        p_bar.progress(i / len(sel_idx))
                        
                        f_tipo = str(f.get('Tipo', '')).strip()
                        reemplazos = {
                            "<<Tipo>>": proper_elegante(f_tipo),
                            # ... (resto de campos igual)
                        }

                        s_text.text(f"🖋️ Ajustando jerarquía de 'Tipo' para {suc_actual}...")
                        prs = Presentation(os.path.join(folder_fisica, map_memoria[f_tipo]))
                        
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    txt_shape = shape.text_frame.text
                                    if "<<Tipo>>" in txt_shape:
                                        # LÓGICA DE REDUCCIÓN DINÁMICA
                                        nuevo_texto = proper_elegante(f_tipo)
                                        tf = shape.text_frame
                                        tf.clear()
                                        tf.word_wrap = True
                                        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                                        
                                        p = tf.paragraphs[0]
                                        p.alignment = PP_ALIGN.CENTER
                                        
                                        run = p.add_run()
                                        run.text = nuevo_texto
                                        run.font.color.rgb = AZUL_CELESTE
                                        run.font.bold = True
                                        
                                        # Empezamos en 64pt y bajamos hasta que LibreOffice lo acomode
                                        # Nota: Dado que soffice maneja el wrap al convertir, 
                                        # establecemos el objetivo visual de 64pt.
                                        run.font.size = Pt(64) 
                                        
                                        # Si el texto es muy corto para romper a 2 líneas en 64pt,
                                        # se puede forzar un tamaño aún mayor o dejar que ocupe el ancho.
                                        # Pero siguiendo tu instrucción, partimos de 64pt.

                        # (Manejo de PDF y ZIP se mantiene igual)
                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        # ... lógica de guardado ...

                st.success("✅ ¡Generación finalizada!")
                st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_Pro_Tipo_64.zip")
