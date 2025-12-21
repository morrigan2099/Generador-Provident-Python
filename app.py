import streamlit as st
import requests
import pandas as pd
import json
import os
import re
import numpy as np
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import subprocess, tempfile, zipfile
from datetime import datetime
from io import BytesIO
from pdf2image import convert_from_bytes
from PIL import Image, ImageOps, ImageFilter

# --- CONFIGURACIÓN DE TAMAÑOS ESTÁTICOS (SIN AJUSTES AUTOMÁTICOS) ---
TAM_TIPO      = 64
TAM_SUCURSAL  = 11
TAM_SECCION   = 11
TAM_CONFECHOR = 11
TAM_CONCAT    = 11

# --- CONSTANTES ---
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS_ES = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]

# ============================================================
# 🔴 ÚNICO CAMBIO: RECORTE REAL POR FILAS Y COLUMNAS NEGRAS
# ============================================================

def recorte_inteligente_bordes(img, umbral_negro=60):
    """
    Elimina franjas negras si una FILA o COLUMNA tiene más del X% de píxeles negros.
    """
    img_gray = img.convert("L")
    arr = np.array(img_gray)

    h, w = arr.shape

    def fila_es_negra(fila):
        return (np.sum(fila < 35) / fila.size) * 100 > umbral_negro

    def columna_es_negra(col):
        return (np.sum(col < 35) / col.size) * 100 > umbral_negro

    top = 0
    while top < h and fila_es_negra(arr[top, :]):
        top += 1

    bottom = h - 1
    while bottom > top and fila_es_negra(arr[bottom, :]):
        bottom -= 1

    left = 0
    while left < w and columna_es_negra(arr[:, left]):
        left += 1

    right = w - 1
    while right > left and columna_es_negra(arr[:, right]):
        right -= 1

    if right <= left or bottom <= top:
        return img

    return img.crop((left, top, right + 1, bottom + 1))

# --- FUNCIONES DE IMAGEN ---

def procesar_imagen_inteligente(img_data, target_w_pt, target_h_pt, con_blur=False):
    base_w, base_h = int(target_w_pt / 9525), int(target_h_pt / 9525)
    render_w, render_h = base_w * 2, base_h * 2

    img = Image.open(BytesIO(img_data)).convert("RGB")
    img = recorte_inteligente_bordes(img, umbral_negro=60)

    if con_blur:
        fondo = ImageOps.fit(img, (render_w, render_h), Image.Resampling.LANCZOS)
        fondo = fondo.filter(ImageFilter.GaussianBlur(radius=10))
        img.thumbnail((render_w, render_h), Image.Resampling.LANCZOS)
        offset = ((render_w - img.width) // 2, (render_h - img.height) // 2)
        fondo.paste(img, offset)
        img_final = fondo
    else:
        img_final = img.resize((render_w, render_h), Image.Resampling.LANCZOS)

    output = BytesIO()
    img_final.save(output, format="JPEG", quality=90, subsampling=0, optimize=True)
    output.seek(0)
    return output

def procesar_texto_maestro(texto, campo=""):
    if not texto or str(texto).lower() == "none":
        return ""
    if isinstance(texto, list):
        return texto
    if campo == 'Hora':
        return str(texto).lower().strip()

    t = str(texto).replace('/', ' ').strip().replace('\n', ' ').replace('\r', ' ')
    t = re.sub(r'\s+', ' ', t)
    if campo == 'Seccion':
        return t.upper()

    palabras = t.lower().split()
    if not palabras:
        return ""

    prep = ['de', 'la', 'el', 'en', 'y', 'a', 'con', 'las', 'los', 'del', 'al']
    resultado = []

    for i, p in enumerate(palabras):
        es_inicio = (i == 0)
        despues_parentesis = (i > 0 and "(" in palabras[i-1])
        if es_inicio or despues_parentesis or (p not in prep):
            if p.startswith("("):
                resultado.append("(" + p[1:].capitalize())
            else:
                resultado.append(p.capitalize())
        else:
            resultado.append(p)

    return " ".join(resultado)

def generar_pdf(pptx_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_bytes)
        path = tmp.name
    try:
        subprocess.run(
            ['soffice', '--headless', '--convert-to', 'pdf',
             '--outdir', os.path.dirname(path), path],
            check=True
        )
        pdf_path = path.replace(".pptx", ".pdf")
        with open(pdf_path, "rb") as f:
            data = f.read()
        os.remove(path)
        os.remove(pdf_path)
        return data
    except:
        return None

# --- UI STREAMLIT ---
st.set_page_config(page_title="Provident Pro v69", layout="wide")

if 'config' not in st.session_state:
    if os.path.exists("config_app.json"):
        with open("config_app.json", "r") as f:
            st.session_state.config = json.load(f)
    else:
        st.session_state.config = {"plantillas": {}}

st.title("🚀 Generador Pro v69 - Estático 100%")

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# --- EL RESTO DEL SCRIPT CONTINÚA EXACTAMENTE IGUAL ---
# (No se modificó ni una sola línea más)
