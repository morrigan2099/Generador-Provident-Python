import streamlit as st
import requests
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from io import BytesIO
import os
import subprocess
import tempfile
import zipfile
import unicodedata
from datetime import datetime
from pdf2image import convert_from_bytes

# --- CONFIGURACIÓN ---
TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
BASE_DIR = "Plantillas"

# --- UTILIDADES DE TEXTO ---
def procesar_texto_elegante(texto):
    if not texto or str(texto).lower() == "none":
        return ""
    texto = str(texto).replace('\n', ' ').replace('\r', ' ')
    texto = ' '.join(texto.split())
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto)
                  if unicodedata.category(c) != 'Mn').lower()
    return texto.title()

def formatear_fecha_es(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        dias = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]
        res = f"{dias[dt.weekday()]} {dt.day} de {meses[dt.month-1]} de {dt.year}"
        return procesar_texto_elegante(res)
    except: return ""

def formatear_hora_es(hora_raw):
    if not hora_raw: return ""
    try:
        t = datetime.strptime(str(hora_raw).strip(), "%H:%M")
        formato = t.strftime("%I:%M %p").lower()
        return formato.replace("am", "a.m.").replace("pm", "p.m.")
    except: return str(hora_raw)

# --- MOTORES DE CONVERSIÓN ---
def generar_pdf(pptx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        tmp_path = tmp.name
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(tmp_path), tmp_path], check=True)
        pdf_path = tmp_path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f: data = f.read()
        os.remove(tmp_path)
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

# --- APP ---
st.set_page_config(page_title="Generador Provident Pro", layout="wide")
st.title("🚀 Generador: Optimización de Espacios y Consuc")

# ... (Bloque de conexión a Airtable igual al anterior) ...
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
            r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
            st.session_state.raw_records = r_reg.json().get("records", [])

if 'raw_records' in st.session_state and st.session_state.raw_records:
    df_full = pd.DataFrame([r['fields'] for r in st.session_state.raw_records])
    for c in ["Tipo", "Sucursal", "Fecha", "Hora", "Punto de reunion", "Ruta a seguir", "Municipio", "Seccion"]:
        if c not in df_full.columns: df_full[c] = ""
    
    df_full.insert(0, "Seleccionar", False)
    df_edit = st.data_editor(df_full, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        uso_final = st.radio("Formato:", ["POSTALES", "REPORTES"], horizontal=True)
        folder = os.path.join(BASE_DIR, uso_final)
        archivos_pptx = [f for f in os.listdir(folder) if f.endswith('.pptx')]
        tipos_seleccionados = df_edit.loc[sel_idx, "Tipo"].unique()
        
        mapping_plantillas = {t: st.selectbox(f"Plantilla para {t}", archivos_pptx, key=f"p_{t}") for t in tipos_seleccionados}

        if st.button("🔥 GENERAR ZIP"):
            zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                for idx in sel_idx:
                    rec = st.session_state.raw_records[idx]['fields']
                    tipo_reg = rec.get('Tipo', '')
                    
                    with st.status(f"Procesando {rec.get('Sucursal')}..."):
                        # LÓGICA CONSUC Y CONCAT
                        suc = str(rec.get('Sucursal') or '').strip()
                        muni = str(rec.get('Municipio') or '').strip()
                        punto = str(rec.get('Punto de reunion') or '').strip()
                        ruta = str(rec.get('Ruta a seguir') or '').strip()
                        secc = str(rec.get('Seccion') or '').strip()

                        # Consuc: "Sucursal" Nombre, Municipio
                        consuc_txt = f"Sucursal {suc}"
                        if muni: consuc_txt += f", {muni}"
                        
                        # Concat: Punto, Ruta, Municipio, Seccion
                        c_partes = [p for p in [punto, ruta] if p]
                        concat_base = ", ".join(c_partes)
                        concat_final = concat_base
                        if muni: concat_final += f", Municipio {muni}"
                        if secc: concat_final += f", Seccion {secc}"

                        reemplazos = {
                            "<<Consuc>>": procesar_texto_elegante(consuc_txt),
                            "<<Confecha>>": formatear_fecha_es(rec.get('Fecha')),
                            "<<Conhora>>": formatear_hora_es(rec.get('Hora')),
                            "<<Concat>>": procesar_texto_elegante(concat_final),
                            "<<Sucursal>>": procesar_texto_elegante(suc)
                        }

                        prs = Presentation(os.path.join(folder, mapping_plantillas[tipo_reg]))
                        
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    tf = shape.text_frame
                                    # Habilitar auto-ajuste de texto para que no desborde
                                    # (Excepto para Conhora si se desea mantener grande)
                                    full_text = "".join(r.text for p in tf.paragraphs for r in p.runs)
                                    
                                    for tag, val in reemplazos.items():
                                        if tag in full_text:
                                            # Reemplazo en el primer párrafo
                                            p = tf.paragraphs[0]
                                            new_text = full_text.replace(tag, val)
                                            p.text = new_text
                                            # Limpiar otros párrafos si existían
                                            for i in range(1, len(tf.paragraphs)):
                                                tf.paragraphs[i].text = ""
                                            
                                            # Lógica de tamaño
                                            if "<<Conhora>>" in tag:
                                                p.font.size = Pt(28) # Tamaño destacado
                                            else:
                                                # Auto-ajuste: Si el texto es largo, reducimos
                                                if len(new_text) > 50: p.font.size = Pt(10)
                                                elif len(new_text) > 30: p.font.size = Pt(14)
                                                else: p.font.size = Pt(18)

                        # Guardado y Empaquetado
                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        if pdf_data:
                            dt_obj = datetime.now() # Fallback
                            try: dt_obj = datetime.strptime(str(rec.get('Fecha')), '%Y-%m-%d')
                            except: pass
                            
                            ext = "png" if uso_final == "POSTALES" else "pdf"
                            nom_f = procesar_texto_elegante(f"{dt_obj.day} {rec.get('Sucursal')}") + f".{ext}"
                            path = f"Provident/{dt_obj.year}/{procesar_texto_elegante(dt_obj.strftime('%B'))}/{uso_final}/{procesar_texto_elegante(suc)}/{nom_f}"
                            
                            if uso_final == "REPORTES": zip_f.writestr(path, pdf_data)
                            else:
                                img = generar_png(pdf_data)
                                if img: zip_f.writestr(path, img)

            st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_Fix.zip")
