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
MESES_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

if 'raw_records' not in st.session_state: st.session_state.raw_records = []
if 'map_memoria' not in st.session_state: st.session_state.map_memoria = {}

def proper_elegante(texto):
    if not texto or str(texto).lower() == "none": return ""
    texto = ''.join(c for c in unicodedata.normalize('NFD', str(texto))
                  if unicodedata.category(c) != 'Mn')
    palabras = texto.lower().split()
    excepciones = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del']
    resultado = []
    for i, p in enumerate(palabras):
        if i > 0 and p in excepciones: resultado.append(p)
        else: resultado.append(p.capitalize())
    return " ".join(resultado)

def formatear_confechor(fecha_str, hora_texto):
    try:
        # Airtable: 2025-11-18
        dt = datetime.strptime(str(fecha_str).strip(), '%Y-%m-%d')
        # Línea 1: dd de Mes | Línea 2: de aaaa, hora
        l1 = f"{dt.day:02d} de {MESES_ES[dt.month-1]}"
        l2 = f"de {dt.year}, {str(hora_texto).lower()}"
        return f"{l1}\n{l2}"
    except:
        return f"{fecha_str}\n{hora_texto}"

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

def descargar_imagen(url):
    try:
        r = requests.get(url, timeout=15)
        if r.status_code == 200: return BytesIO(r.content)
    except: return None

# --- APP ---
st.set_page_config(page_title="Provident Pro Final", layout="wide")
st.title("🚀 Generador Pro: Estructura y Formatos Restaurados")

with st.sidebar:
    headers = {"Authorization": f"Bearer {TOKEN}"}
    r_bases = requests.get("https://api.airtable.com/v0/meta/bases", headers=headers)
    if r_bases.status_code == 200:
        base_opts = {b['name']: b['id'] for b in r_bases.json()['bases']}
        base_sel = st.selectbox("Base:", list(base_opts.keys()))
        if base_sel:
            r_tab = requests.get(f"https://api.airtable.com/v0/meta/bases/{base_opts[base_sel]}/tables", headers=headers)
            if r_tab.status_code == 200:
                tabla_opts = {t['name']: t['id'] for t in r_tab.json()['tables']}
                tabla_sel = st.selectbox("Tabla:", list(tabla_opts.keys()))
                if st.button("🔄 Cargar Datos"):
                    r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
                    st.session_state.raw_records = r_reg.json().get("records", [])
                    st.rerun()

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
            for t in tipos_unicos:
                st.session_state.map_memoria[t] = st.selectbox(f"Plantilla para {t}:", archivos_pptx, key=f"p_{t}")

            if st.button("🔥 GENERAR TODO"):
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                    for idx in sel_idx:
                        fields = st.session_state.raw_records[idx]['fields']
                        
                        # 1. VALORES PARA NOMENCLATURA Y REEMPLAZO
                        f_fecha = fields.get('Fecha', '2025-01-01')
                        dt_obj = datetime.strptime(f_fecha, '%Y-%m-%d')
                        f_suc = str(fields.get('Sucursal', '')).strip()
                        f_muni = str(fields.get('Municipio', '')).strip()
                        f_punto = str(fields.get('Punto de reunion', '')).strip()
                        f_ruta = str(fields.get('Ruta a seguir', '')).strip()
                        f_tipo = str(fields.get('Tipo', '')).upper()

                        # Nomenclatura específica
                        partes_nom = [p for p in [f_punto, f_ruta] if p]
                        if f_muni: partes_nom.append(f_muni)
                        str_nom = ", ".join(partes_nom)
                        nombre_archivo = f"{f_fecha} - {f_tipo} {f_suc.upper()} - {str_nom}"
                        
                        # 2. DICCIONARIO DE REEMPLAZOS
                        reemplazos = {
                            "<<Tipo>>": f_tipo,
                            "<<confechor>>": formatear_confechor(f_fecha, fields.get('Hora', '')),
                            "<<Consuc>>": proper_elegante(f"Sucursal {f_suc}, {f_muni}"),
                            "<<Concat>>": proper_elegante(f"{f_punto}, {f_ruta}, {f_muni}"),
                            "<<Sucursal>>": proper_elegante(f_suc)
                        }

                        # 3. PROCESAR POWERPOINT
                        prs = Presentation(os.path.join(folder_fisica, st.session_state.map_memoria[fields.get('Tipo')]))
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    for paragraph in shape.text_frame.paragraphs:
                                        full_text = "".join(run.text for run in paragraph.runs)
                                        for tag, val in reemplazos.items():
                                            if tag in full_text:
                                                new_val = full_text.replace(tag, val)
                                                for r in paragraph.runs: r.text = ""
                                                if paragraph.runs:
                                                    run = paragraph.runs[0]
                                                    run.text = new_val
                                                    run.font.size = Pt(32) if "\n" in val else Pt(42)

                                if modo == "Reportes":
                                    tags_f = ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]
                                    for tf in tags_f:
                                        if (shape.has_text_frame and f"<<{tf}>>" in shape.text) or (tf in shape.name):
                                            adj = fields.get(tf)
                                            if adj and isinstance(adj, list):
                                                img_data = descargar_imagen(adj[0].get('url'))
                                                if img_data:
                                                    slide.shapes.add_picture(img_data, shape.left, shape.top, shape.width, shape.height)

                        # 4. EXPORTACIÓN Y ESTRUCTURA DE CARPETAS
                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        if pdf_data:
                            ext = ".jpg" if modo == "Postales" else ".pdf"
                            nom_final = proper_elegante(nombre_archivo) + ext
                            folder_mes = f"{dt_obj.month:02d} - {MESES_ES[dt_obj.month-1]}"
                            
                            # RUTA: Provident / Año / Mes / Modo / Sucursal / Archivo
                            ruta_en_zip = f"Provident/{dt_obj.year}/{folder_mes}/{modo}/{proper_elegante(f_suc)}/{nom_final}"
                            
                            if modo == "Reportes":
                                zip_f.writestr(ruta_en_zip, pdf_data)
                            else:
                                imgs = convert_from_bytes(pdf_data)
                                if imgs:
                                    img_io = BytesIO()
                                    imgs[0].convert('RGB').save(img_io, format='JPEG', quality=85)
                                    zip_f.writestr(ruta_en_zip, img_io.getvalue())

                st.success("✅ Estructura de directorios y formatos generados correctamente.")
                st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_Pro_Completo.zip")
