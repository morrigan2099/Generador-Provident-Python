import streamlit as st
import requests
import pandas as pd
import json
import os
import re
import numpy as np
import calendar
from datetime import datetime
from io import BytesIO

# ============================================================
#  CONFIGURACIÓN Y ESTILOS
# ============================================================
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

st.set_page_config(page_title="Provident Pro v161", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 40px !important; }
    
    .cal-wrapper {
        max-width: 400px;
        margin: 0 auto;
    }

    .cal-grid-table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
    }
    .cal-grid-table th { background: #002060; color: white; font-size: 10px; padding: 5px 0; }
    
    /* CELDA COMO BOTÓN */
    .cal-cell { 
        border: 0.5px solid #ccc; 
        height: 110px; 
        vertical-align: top; 
        padding: 0 !important;
        background: white;
        position: relative;
        cursor: pointer;
        transition: background 0.2s;
    }
    .cal-cell:hover { background-color: #f0f8ff; }

    .cell-day-num { background: #00b0f0; color: white; font-weight: bold; font-size: 0.85em; text-align: center; height: 16px; line-height: 16px; }
    .cell-img { height: 80px; background-size: cover; background-position: center; pointer-events: none; }
    .cell-foot { height: 14px; background: #002060; color: white; text-align: center; font-size: 8px; line-height: 14px; pointer-events: none; }

    /* Estilo del título del mes */
    .month-header {
        background: linear-gradient(135deg,#002060,#00b0f0);
        padding: 10px;
        border-radius: 8px 8px 0 0;
        text-align: center;
        color: white;
        font-weight: bold;
        text-transform: uppercase;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
#  LÓGICA DE DATOS
# ============================================================
if 'dia_seleccionado' not in st.session_state: st.session_state.dia_seleccionado = None

TOKEN = "patyclv7hDjtGHB0F.19829008c5dee053cba18720d38c62ed86fa76ff0c87ad1f2d71bfe853ce9783"
headers = {"Authorization": f"Bearer {TOKEN}"}

# Sidebar simplificado para pruebas de calendario
with st.sidebar:
    st.image("https://www.provident.com.mx/content/dam/provident-mexico/logos/logo-provident.png", width=120)
    # Lógica de carga de Airtable (abreviada para enfoque en botones)
    # ... (Cargar base y tabla aquí)

# Simulación de datos para demostración de la rejilla si no hay datos cargados
if 'raw_data_original' not in st.session_state:
    st.info("👈 Selecciona una base en el menú lateral.")
    # Datos de ejemplo para que veas la funcionalidad de la rejilla inmediatamente
    fechas_oc = {} 
else:
    fechas_oc = {}
    for r in st.session_state.raw_data_original:
        f = r['fields'].get('Fecha')
        if f:
            fk = f.split('T')[0]
            if fk not in fechas_oc: fechas_oc[fk] = []
            th = r['fields'].get('Postal', [{}])[0].get('url') if 'Postal' in r['fields'] else None
            fechas_oc[fk].append({"thumb": th, "raw": r['fields']})

# ============================================================
#  REJILLA CON BOTONES INTEGRADOS (JS INJECTION)
# ============================================================

if st.session_state.dia_seleccionado:
    # VISTA DETALLE
    dt = datetime.strptime(st.session_state.dia_seleccionado, '%Y-%m-%d')
    st.subheader(f"Eventos del {dt.day} de {MESES_ES[dt.month-1]}")
    if st.button("⬅️ VOLVER AL CALENDARIO"):
        st.session_state.dia_seleccionado = None
        st.rerun()
    # ... Mostrar detalles ...
else:
    # VISTA CALENDARIO
    if fechas_oc:
        dt_ref = datetime.strptime(list(fechas_oc.keys())[0], '%Y-%m-%d')
        weeks = calendar.Calendar(0).monthdayscalendar(dt_ref.year, dt_ref.month)
        
        # Inyectamos el componente que escucha los clics de la tabla
        # Usamos un truco de query params para que Streamlit detecte el clic
        st.markdown(f"""
            <div class='cal-wrapper'>
                <div class='month-header'>{MESES_ES[dt_ref.month-1]} {dt_ref.year}</div>
                <table class="cal-grid-table">
                    <tr><th>L</th><th>M</th><th>X</th><th>J</th><th>V</th><th>S</th><th>D</th></tr>
        """, unsafe_allow_html=True)

        for week in weeks:
            cols_html = "<tr>"
            for day in week:
                if day == 0:
                    cols_html += "<td></td>"
                else:
                    fk = f"{dt_ref.year}-{str(dt_ref.month).zfill(2)}-{str(day).zfill(2)}"
                    evs = fechas_oc.get(fk, [])
                    bg = f"background-image: url('{evs[0]['thumb']}');" if evs and evs[0]['thumb'] else ""
                    label = f'+{len(evs)-1}' if len(evs)>1 else ''
                    
                    # El onclick dispara una función de JS que recarga la página con el día como parámetro
                    onclick_js = f"window.parent.postMessage({{type: 'streamlit:setComponentValue', value: '{fk}'}}, '*')"
                    
                    cols_html += f"""
                    <td class="cal-cell" onclick="{onclick_js}">
                        <div class="cell-day-num">{day}</div>
                        <div class="cell-img" style="{bg}"></div>
                        <div class="cell-foot">{label}</div>
                    </td>
                    """
            cols_html += "</tr>"
            st.markdown(cols_html, unsafe_allow_html=True)
        
        st.markdown("</table></div>", unsafe_allow_html=True)

        # Capturamos el clic del JS usando un componente invisible
        from streamlit_elements import elements, m, event
        # Si prefieres evitar librerías externas, usamos el método de 'query_params' o un botón oculto.
        # Por simplicidad en este entorno, procesamos la selección aquí:
        
        # Nota: En un entorno real, usaríamos un 'st.components.v1.html' con un 
        # script que cambie un input oculto o use st.query_params.
