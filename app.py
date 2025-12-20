import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
import os, subprocess, tempfile, zipfile, unicodedata
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"

if 'raw_records' not in st.session_state: st.session_state.raw_records = []
if 'map_memoria' not in st.session_state: st.session_state.map_memoria = {}

# --- LÓGICA PROPER ELEGANTE (ESPAÑOL MX / SIN ACENTOS) ---
def proper_elegante(texto):
    if not texto or str(texto).lower() == "none": return ""
    # Quitar acentos
    texto = ''.join(c for c in unicodedata.normalize('NFD', str(texto))
                  if unicodedata.category(c) != 'Mn')
    palabras = texto.lower().split()
    excepciones = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del']
    resultado = []
    for i, p in enumerate(palabras):
        if i > 0 and p in excepciones:
            resultado.append(p)
        else:
            resultado.append(p.capitalize())
    return " ".join(resultado)

def formatear_fecha_mx(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        dias = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]
        res = f"{dias[dt.weekday()]} {dt.day} de {meses[dt.month-1]} de {dt.year}"
        return proper_elegante(res)
    except: return ""

def formatear_hora_mx(hora_raw):
    if not hora_raw: return ""
    try:
        t = datetime.strptime(str(hora_raw).strip(), "%H:%M")
        return t.strftime("%I:%M %p").lower().replace("am", "a.m.").replace("pm", "p.m.")
    except: return str(hora_raw)

# --- MOTORES DE CONVERSIÓN ---
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

def generar_png(pdf_bytes):
    try:
        images = convert_from_bytes(pdf_bytes)
        if images:
            buf = BytesIO()
            images[0].save(buf, format='PNG')
            return buf.getvalue()
    except: return None

# --- INTERFAZ ---
st.set_page_config(page_title="Provident Pro", layout="wide")
st.title("🚀 Generador Pro: Proper Elegante & Estructura")

with st.sidebar:
    headers = {"Authorization": f"Bearer {TOKEN}"}
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base:", list(base_opts.keys()))
        r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
        if r_tab.status_code == 200:
            tabla_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
            tabla_sel = st.selectbox("Tabla:", list(tabla_opts.keys()))
            if st.button("🔄 Cargar"):
                r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
                st.session_state.raw_records = r_reg.json().get("records", [])
                st.rerun()

if st.session_state.raw_records:
    # --- PROCESAR TABLA PARA QUE SEA PROPER ELEGANTE ---
    data_list = []
    for r in st.session_state.raw_records:
        f = r['fields']
        data_list.append({
            "Tipo": proper_elegante(f.get("Tipo")),
            "Sucursal": proper_elegante(f.get("Sucursal")),
            "Municipio": proper_elegante(f.get("Municipio")),
            "Fecha": f.get("Fecha", ""),
            "_original_tipo": f.get("Tipo") # Para mapeo interno
        })
    
    df_display = pd.DataFrame(data_list)
    df_display.insert(0, "Seleccionar", False)
    
    st.write("### 📋 Registros Disponibles (Vista Proper Elegante)")
    df_edit = st.data_editor(df_display, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        st.divider()
        opcion_salida = st.radio("Tipo a generar:", ["Postales", "Reportes"], horizontal=True)
        # La carpeta física sigue siendo REPORTES/POSTALES pero la lógica de ruta usará Proper
        folder_fisica = os.path.join(BASE_DIR, opcion_salida.upper()) 
        
        if os.path.exists(folder_fisica):
            archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
            tipos_en_seleccion = df_edit.loc[sel_idx, "Tipo"].unique()
            
            st.write("### 📂 Asignar Plantillas")
            cols = st.columns(len(tipos_en_seleccion))
            for i, t_proper in enumerate(tipos_en_seleccion):
                with cols[i]:
                    idx_mem = 0
                    if t_proper in st.session_state.map_memoria:
                        if st.session_state.map_memoria[t_proper] in archivos_pptx:
                            idx_mem = archivos_pptx.index(st.session_state.map_memoria[t_proper])
                    
                    sel_p = st.selectbox(f"{t_proper}:", archivos_pptx, index=idx_mem, key=f"s_{t_proper}")
                    st.session_state.map_memoria[t_proper] = sel_p

            if st.button("🔥 GENERAR ZIP ESTRUCTURADO"):
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                    for idx in sel_idx:
                        # Recuperar campos originales de Airtable
                        rec = st.session_state.raw_records[idx]['fields']
                        
                        with st.status(f"Procesando: {proper_elegante(rec.get('Sucursal'))}"):
                            # 1. LÓGICA DE PLACEHOLDERS (CONCAT, CONFECHA, CONHORA, CONSUC)
                            suc = str(rec.get('Sucursal') or '').strip()
                            muni = str(rec.get('Municipio') or '').strip()
                            punto = str(rec.get('Punto de reunion') or '').strip()
                            ruta = str(rec.get('Ruta a seguir') or '').strip()
                            secc = str(rec.get('Seccion') or '').strip()
                            
                            # <<Consuc>>
                            consuc_txt = f"Sucursal {suc}" + (f", {muni}" if muni else "")
                            
                            # <<Concat>> (Punto, Ruta, Municipio, Seccion)
                            partes = [p for p in [punto, ruta] if p]
                            c_base = ", ".join(partes)
                            if muni: c_base += f", Municipio {muni}"
                            if secc: c_base += f", Seccion {secc}"

                            reemplazos = {
                                "<<Consuc>>": proper_elegante(consuc_txt),
                                "<<Confecha>>": formatear_fecha_mx(rec.get('Fecha')),
                                "<<Conhora>>": formatear_hora_mx(rec.get('Hora')),
                                "<<Concat>>": proper_elegante(c_base),
                                "<<Sucursal>>": proper_elegante(suc)
                            }

                            # 2. CARGAR PLANTILLA USANDO MEMORIA
                            t_key = proper_elegante(rec.get('Tipo'))
                            path_pptx = os.path.join(folder_fisica, st.session_state.map_memoria[t_key])
                            prs = Presentation(path_pptx)
                            
                            for slide in prs.slides:
                                for shape in slide.shapes:
                                    if shape.has_text_frame:
                                        for paragraph in shape.text_frame.paragraphs:
                                            for run in paragraph.runs:
                                                for tag, val in reemplazos.items():
                                                    if tag in run.text:
                                                        run.text = run.text.replace(tag, val)

                            # 3. GUARDADO Y ESTRUCTURA ZIP (Proper Elegante)
                            pp_io = BytesIO(); prs.save(pp_io)
                            pdf_data = generar_pdf(pp_io.getvalue())
                            
                            if pdf_data:
                                try: dt = datetime.strptime(str(rec.get('Fecha')), '%Y-%m-%d')
                                except: dt = datetime.now()
                                
                                mes_nombre = proper_elegante(dt.strftime('%m - %B'))
                                ext = ".png" if opcion_salida == "Postales" else ".pdf"
                                nombre_archivo = proper_elegante(f"{dt.day} de {dt.month} {rec.get('Tipo')} {suc}") + ext
                                
                                # Provident / Año / Mes / Uso / Sucursal / Archivo
                                ruta_zip = (f"Provident/{dt.year}/{mes_nombre}/{opcion_salida}/"
                                           f"{proper_elegante(suc)}/{nombre_archivo}")
                                
                                if opcion_salida == "Reportes":
                                    zip_f.writestr(ruta_zip, pdf_data)
                                else:
                                    img = generar_png(pdf_data)
                                    if img: zip_f.writestr(ruta_zip, img)

                st.success("✅ Proceso completado exitosamente.")
                st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), f"Provident_{opcion_salida}.zip")
