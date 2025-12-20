import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from io import BytesIO
import os
import subprocess
import tempfile
import zipfile
import re
import unicodedata
from datetime import datetime
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN FIJA ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"

# --- UTILIDADES DE TEXTO ---
def limpiar_y_estilo(texto):
    """Limpia acentos, minúsculas, elimina saltos y aplica Proper Case"""
    if not texto or str(texto).lower() == "none":
        return ""
    
    # 1. Texto plano, eliminar saltos de línea y dobles espacios
    texto = str(texto).replace('\n', ' ').replace('\r', ' ')
    texto = ' '.join(texto.split())
    
    # 2. Eliminar acentos y pasar a minúsculas
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto)
                  if unicodedata.category(c) != 'Mn').lower()
    
    # 3. Formato Proper Case (Elegante)
    return texto.title()

def formatear_fecha_es(fecha_str):
    """Formatea fecha a 'jueves 20 de mayo de 2024'"""
    try:
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
        
        dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        # Formato: jueves 20 de mayo de 2024 (en dos líneas si es necesario en PPTX)
        resultado = f"{dias[dt.weekday()]} {dt.day} de {meses[dt.month-1]} de {dt.year}"
        return limpiar_y_estilo(resultado)
    except:
        return ""

def formatear_hora_es(hora_raw):
    """Convierte texto de hora a h:mm a.m./p.m."""
    if not hora_raw: return ""
    try:
        # Intenta parsear diferentes formatos de hora
        hora_limpia = str(hora_raw).strip().lower().replace(".", "")
        dt_hora = datetime.strptime(hora_limpia, "%H:%M") # Formato 24h esperado
        formato = dt_hora.strftime("%I:%M %p").lower()
        return formato.replace("am", "a.m.").replace("pm", "p.m.")
    except:
        return str(hora_raw)

# --- PROCESAMIENTO DE DOCUMENTOS ---
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
    except: return None

def generar_png(pdf_bytes):
    try:
        images = convert_from_bytes(pdf_bytes)
        if images:
            img_byte_arr = BytesIO()
            images[0].save(img_byte_arr, format='PNG')
            return img_byte_arr.getvalue()
    except: return None

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")
st.title("🚀 Generador Masivo Personalizado")

# (Lógica de sidebar y carga de Airtable igual a la anterior...)
with st.sidebar:
    st.header("🔑 Conexión")
    headers = {"Authorization": f"Bearer {TOKEN}"}
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_sel = st.selectbox("Base:", [b['name'] for b in r_bases.json()['bases']])
        # ... (Carga de tablas y registros simplificada para el ejemplo)

# --- LÓGICA DE REEMPLAZO Y ZIP ---
if 'registros_raw' in st.session_state:
    df = pd.DataFrame([r['fields'] for r in st.session_state.registros_raw])
    df.insert(0, "Seleccionar", False)
    df_editado = st.data_editor(df, use_container_width=True, hide_index=True)
    seleccionados = df_editado[df_editado["Seleccionar"] == True]

    if not seleccionados.empty:
        uso_final = st.radio("Uso:", ["POSTALES", "REPORTES"], horizontal=True)
        folder_path = os.path.join(BASE_DIR, uso_final)
        
        if os.path.exists(folder_path):
            archivos_pptx = [f for f in os.listdir(folder_path) if f.endswith('.pptx')]
            tipos = seleccionados["Tipo"].unique()
            mapping = {t: st.selectbox(f"Plantilla para {t}", archivos_pptx, key=t) for t in tipos}

            if st.button("🔥 GENERAR ZIP ESTRUCTURADO"):
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_f:
                    for idx, fila in seleccionados.iterrows():
                        with st.status(f"Procesando {fila.get('Sucursal')}..."):
                            
                            # 1. PREPARAR VALORES PERSONALIZADOS
                            punto = str(fila.get('Punto de reunion') or '').strip()
                            ruta = str(fila.get('Ruta a seguir') or '').strip()
                            muni = str(fila.get('Municipio') or '').strip()
                            secc = str(fila.get('Seccion') or '').strip()
                            
                            # Lógica Concat
                            concat_partes = [p for p in [punto, ruta] if p]
                            centro = ", ".join(concat_partes)
                            concat_final = centro
                            if muni: concat_final += f", Municipio {muni}"
                            if secc: concat_final += f", Sección {secc}"

                            reemplazos = {
                                "<<Sucursal>>": limpiar_y_estilo(fila.get('Sucursal')),
                                "<<Confecha>>": formatear_fecha_es(fila.get('Fecha')),
                                "<<Conhora>>": formatear_hora_es(fila.get('Hora')),
                                "<<Concat>>": limpiar_y_estilo(concat_final)
                            }

                            # 2. REEMPLAZAR EN PPTX
                            prs = Presentation(os.path.join(folder_path, mapping[fila['Tipo']]))
                            for slide in prs.slides:
                                for shape in slide.shapes:
                                    if shape.has_text_frame:
                                        for p in shape.text_frame.paragraphs:
                                            for run in p.runs:
                                                for tag, val in reemplazos.items():
                                                    if tag in run.text:
                                                        run.text = run.text.replace(tag, val)

                            # 3. CONVERTIR Y GUARDAR EN ZIP
                            pp_io = BytesIO(); prs.save(pp_io)
                            pdf_bin = generar_pdf(pp_io.getvalue())
                            
                            if pdf_bin:
                                # (Lógica de nombre de archivo y carpetas Provident/Año/Mes... igual a la anterior)
                                ext = "png" if uso_final == "POSTALES" else "pdf"
                                nombre_f = f"{limpiar_y_estilo(fila.get('Sucursal'))}.{ext}"
                                if uso_final == "REPORTES":
                                    zip_f.writestr(nombre_f, pdf_bin)
                                else:
                                    png_bin = generar_png(pdf_bin)
                                    if png_bin: zip_f.writestr(nombre_f, png_bin)

                st.download_button("📥 DESCARGAR ZIP", zip_buffer.getvalue(), "Provident.zip")
