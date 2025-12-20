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

def forzar_dos_lineas(texto):
    if not texto or "\n" in texto: return texto
    palabras = texto.split()
    if len(palabras) < 2: return texto
    mitad = len(palabras) // 2
    return " ".join(palabras[:mitad]) + "\n" + " ".join(palabras[mitad:])

def interpretar_hora(hora_txt):
    if not hora_txt: return ""
    hora_txt = str(hora_txt).strip().lower()
    for fmt in ["%H:%M", "%I:%M %p", "%H%M", "%I%p"]:
        try:
            dt_hora = datetime.strptime(hora_txt.replace(" ", ""), fmt.replace(" ", ""))
            return dt_hora.strftime("%I:%M %p").lower().replace("am", "a.m.").replace("pm", "p.m.")
        except: continue
    return hora_txt

def formatear_confechor(fecha_str, hora_txt):
    try:
        dt = datetime.strptime(str(fecha_str).strip(), '%Y-%m-%d')
        dia_semana = DIAS_ES[dt.weekday()]
        mes_nombre = MESES_ES[dt.month-1]
        hora_formateada = interpretar_hora(hora_txt)
        l1 = proper_elegante(f"{dia_semana} {mes_nombre} {dt.day:02d}")
        l2 = f"de {dt.year}, {hora_formateada}"
        return f"{l1}\n{l2}"
    except: return f"{fecha_str}\n{hora_txt}"

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
st.set_page_config(page_title="Provident Pro Ultra", layout="wide")
st.title("🚀 Generador Pro: Formato de Cuadro Centrado")

if 'raw_records' not in st.session_state: st.session_state.raw_records = []

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
                if st.button("🔄 CARGAR DATOS"):
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
            map_memoria = {t: st.selectbox(f"Plantilla para {t}:", archivos_pptx, key=f"p_{t}") for t in tipos_unicos}

            if st.button("🔥 GENERAR"):
                p_bar = st.progress(0); s_text = st.empty()
                zip_buf = BytesIO()
                with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED) as zip_f:
                    for i, idx in enumerate(sel_idx):
                        f = st.session_state.raw_records[idx]['fields']
                        suc_actual = str(f.get('Sucursal', 'Doc'))
                        p_bar.progress(i / len(sel_idx))
                        
                        f_fecha = f.get('Fecha', '2025-01-01')
                        dt = datetime.strptime(f_fecha, '%Y-%m-%d')
                        f_suc, f_muni, f_tipo = [str(f.get(k, '')).strip() for k in ['Sucursal', 'Municipio', 'Tipo']]
                        f_punto, f_ruta = [str(f.get(k, '')).strip() for k in ['Punto de reunion', 'Ruta a seguir']]
                        f_seccion = str(f.get('Seccion', '')).upper()
                        
                        lugar_corto = min([o for o in [f_punto, f_ruta] if o], key=len) if (f_punto or f_ruta) else ""

                        reemplazos = {
                            "<<Tipo>>": forzar_dos_lineas(proper_elegante(f_tipo)),
                            "<<Confechor>>": formatear_confechor(f_fecha, f.get('Hora', '')),
                            "<<Consuc>>": forzar_dos_lineas(proper_elegante(f"Sucursal {f_suc}, {f_muni}")),
                            "<<Concat>>": forzar_dos_lineas(proper_elegante(f"{f_punto}, {f_ruta}, {f_muni}")),
                            "<<Sucursal>>": proper_elegante(f_suc),
                            "<<Seccion>>": f_seccion 
                        }

                        s_text.text(f"🖋️ Aplicando formato centrado a {suc_actual}...")
                        prs = Presentation(os.path.join(folder_fisica, map_memoria[f_tipo]))
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    # CONFIGURACIÓN DEL CUADRO DE TEXTO
                                    tf = shape.text_frame
                                    txt_old = tf.text
                                    tag_hit = next((tag for tag in reemplazos if tag in txt_old), None)
                                    
                                    if tag_hit:
                                        # Aplicamos formato al contenedor
                                        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                                        tf.word_wrap = True
                                        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
                                        
                                        # Limpieza e inserción
                                        txt_nuevo = txt_old.replace(tag_hit, reemplazos[tag_hit])
                                        tf.clear()
                                        p = tf.paragraphs[0]
                                        p.alignment = PP_ALIGN.CENTER # Centrado horizontal del párrafo
                                        
                                        run = p.add_run()
                                        run.text = txt_nuevo
                                        run.font.color.rgb = AZUL_CELESTE
                                        run.font.bold = True
                                        
                                        # Ajuste de tamaños
                                        if tag_hit in ["<<Confechor>>", "<<Tipo>>"]: run.font.size = Pt(32)
                                        elif tag_hit == "<<Sucursal>>": run.font.size = Pt(36)
                                        elif tag_hit == "<<Seccion>>": run.font.size = Pt(38)
                                        else: run.font.size = Pt(46)

                        # Procesamiento de fotos y guardado...
                        if modo == "Reportes":
                            for tf in ["Foto de equipo", "Foto 01", "Foto 02", "Foto 03", "Foto 04", "Foto 05", "Foto 06", "Foto 07", "Reporte firmado", "Lista de asistencia"]:
                                adj = f.get(tf)
                                if adj and isinstance(adj, list):
                                    s_text.text(f"📸 Descargando {tf} para {suc_actual}...")
                                    r_img = requests.get(adj[0].get('url'))
                                    if r_img.status_code == 200:
                                        for slide in prs.slides:
                                            for shape in slide.shapes:
                                                if (shape.has_text_frame and f"<<{tf}>>" in shape.text) or (tf in shape.name):
                                                    slide.shapes.add_picture(BytesIO(r_img.content), shape.left, shape.top, shape.width, shape.height)

                        pp_io = BytesIO(); prs.save(pp_io)
                        pdf_data = generar_pdf(pp_io.getvalue())
                        if pdf_data:
                            nom_file = f"{proper_elegante(MESES_ES[dt.month-1] + ' ' + str(dt.day).zfill(2) + ' de ' + str(dt.year))} - {proper_elegante(lugar_corto)}, {proper_elegante(f_muni)} - {proper_elegante(f_tipo)}, {proper_elegante(f_suc)}"
                            ext = ".pdf" if modo == "Reportes" else ".jpg"
                            ruta_zip = f"Provident/{dt.year}/{str(dt.month).zfill(2)}/{modo}/{proper_elegante(f_suc)}/{nom_file[:140] + ext}"
                            zip_f.writestr(ruta_zip, pdf_data if modo == "Reportes" else convert_from_bytes(pdf_data)[0].tobytes())

                p_bar.progress(1.0); s_text.text("✅ ¡Proceso completado!")
                st.download_button("📥 DESCARGAR ZIP", zip_buf.getvalue(), "Provident_Pro_Final.zip")
