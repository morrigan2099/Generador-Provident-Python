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

def formatear_fecha_mx(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, '%Y-%m-%d')
        return proper_elegante(f"{dt.day:02d} de {MESES_ES[dt.month-1]} de {dt.year}")
    except: return ""

def formatear_hora_mx(hora_raw):
    if not hora_raw: return ""
    try:
        t = datetime.strptime(str(hora_raw).strip(), "%H:%M")
        return t.strftime("%I:%M %p").lower().replace("am", "a.m.").replace("pm", "p.m.")
    except: return str(hora_raw)

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

# --- APP ---
st.set_page_config(page_title="Provident Pro", layout="wide")
st.title("🚀 Generador Pro: Nomenclatura y Contenido Restaurado")

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
            if st.button("🔄 Cargar Registros"):
                r_reg = requests.get(f"https://api.airtable.com/v0/{base_opts[base_sel]}/{tabla_opts[tabla_sel]}", headers=headers)
                st.session_state.raw_records = r_reg.json().get("records", [])
                st.rerun()

if st.session_state.raw_records:
    data_list = [{"Tipo": proper_elegante(r['fields'].get("Tipo")), "Sucursal": proper_elegante(r['fields'].get("Sucursal")), "Municipio": proper_elegante(r['fields'].get("Municipio")), "Fecha": r['fields'].get("Fecha", "")} for r in st.session_state.raw_records]
    df_display = pd.DataFrame(data_list)
    df_display.insert(0, "Seleccionar", False)
    df_edit = st.data_editor(df_display, use_container_width=True, hide_index=True)
    sel_idx = df_edit.index[df_edit["Seleccionar"] == True].tolist()

    if sel_idx:
        op_salida = st.radio("Formato de salida:", ["Postales", "Reportes"], horizontal=True)
        folder_fisica = os.path.join(BASE_DIR, op_salida.upper())
        
        if os.path.exists(folder_fisica):
            archivos_pptx = [f for f in os.listdir(folder_fisica) if f.endswith('.pptx')]
            tipos_unicos = df_edit.loc[sel_idx, "Tipo"].unique()
            cols = st.columns(len(tipos_unicos))
            for i, t_p in enumerate(tipos_unicos):
                with cols[i]:
                    st.session_state.map_memoria[t_p] = st.selectbox(f"Plantilla {t_p}:", archivos_pptx, key=f"s_{t_p}")

            if st.button("🔥 GENERAR TODO"):
                progreso_bar = st.progress(0)
                status_text = st.empty()
                zip_buf = BytesIO()
                
                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                    for i, idx in enumerate(sel_idx):
                        rec = st.session_state.raw_records[idx]['fields']
                        
                        # --- 1. LÓGICA DE DATOS (CONCATENACIÓN CONTINUA) ---
                        punto = str(rec.get('Punto de reunion') or '').strip()
                        ruta = str(rec.get('Ruta a seguir') or '').strip()
                        muni = str(rec.get('Municipio') or '').strip()
                        secc = str(rec.get('Seccion') or '').strip()
                        suc_raw = str(rec.get('Sucursal') or '').strip()

                        # Texto para el Placeholder (Continuo)
                        c_list = [p for p in [punto, ruta] if p]
                        if muni: c_list.append(f"Municipio {muni}")
                        if secc: c_list.append(f"Seccion {secc}")
                        concat_placeholder = proper_elegante(", ".join(c_list))

                        reemplazos = {
                            "<<Consuc>>": proper_elegante(f"Sucursal {suc_raw}" + (f", {muni}" if muni else "")),
                            "<<Confecha>>": formatear_fecha_mx(rec.get('Fecha')),
                            "<<Conhora>>": formatear_hora_mx(rec.get('Hora')),
                            "<<Concat>>": concat_placeholder,
                            "<<Sucursal>>": proper_elegante(suc_raw)
                        }

                        # --- 2. PROCESAMIENTO PPTX (RESTAURADO) ---
                        prs = Presentation(os.path.join(folder_fisica, st.session_state.map_memoria[proper_elegante(rec.get('Tipo'))]))
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    shape.text_frame.word_wrap = True
                                    for paragraph in shape.text_frame.paragraphs:
                                        full_txt = "".join(r.text for r in paragraph.runs)
                                        for tag, val in reemplazos.items():
                                            if tag in full_txt:
                                                new_txt = full_txt.replace(tag, val)
                                                # Escalado Gigante
                                                if tag == "<<Concat>>": size = 42 if len(new_txt) < 60 else 36
                                                elif tag == "<<Conhora>>": size = 42
                                                else: size = 42 if len(new_txt) < 30 else 32
                                                
                                                run = paragraph.runs[0]
                                                run.text = new_txt
                                                run.font.size = Pt(size)
                                                for r_idx in range(1, len(paragraph.runs)):
                                                    paragraph.runs[r_idx].text = ""

                        # --- 3. NOMENCLATURA DEL ARCHIVO ---
                        fecha_f = formatear_fecha_mx(rec.get('Fecha'))
                        tipo_f, suc_f = str(rec.get('Tipo') or '').upper(), suc_raw.upper()
                        
                        partes_nom = [p for p in [punto, ruta] if p]
                        if muni: partes_nom.append(muni)
                        
                        str_nom = ", ".join(partes_nom)
                        nombre_base = f"{fecha_f} - {tipo_f} {suc_f} - {str_nom}"
                        
                        if len(nombre_base) > 150:
                            partes_red = [ruta] if (punto and ruta and len(punto) < len(ruta)) else ([punto] if punto else [])
                            if muni: partes_red.append(muni)
                            nombre_base = f"{fecha_f} - {tipo_f} {suc_f} - {', '.join(partes_red)}"

                        nombre_final = proper_elegante(nombre_base) + (".png" if op_salida == "Postales" else ".pdf")

                        # --- 4. GUARDADO ---
                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        if pdf_data:
                            dt = datetime.strptime(str(rec.get('Fecha', '2024-01-01')), '%Y-%m-%d')
                            folder_mes = f"{dt.month:02d} - {MESES_ES[dt.month-1]}"
                            ruta_zip = f"Provident/{dt.year}/{folder_mes}/{op_salida}/{proper_elegante(suc_raw)}/{nombre_final}"
                            
                            if op_salida == "Reportes":
                                zip_f.writestr(ruta_zip, pdf_data)
                            else:
                                imgs = convert_from_bytes(pdf_data)
                                if imgs:
                                    img_io = BytesIO(); imgs[0].save(img_io, format='PNG')
                                    zip_f.writestr(ruta_zip, img_io.getvalue())
                        
                        progreso_bar.progress((i + 1) / len(sel_idx))
                        status_text.info(f"Listo: {nombre_final[:40]}...")

                st.success("✅ Proceso completado con éxito")
                st.download_button("📥 DESCARGAR ZIP FINAL", zip_buf.getvalue(), "Provident_Pro_Final.zip")
