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

# ===============================
# FUNCIONES AUXILIARES
# ===============================

def limpiar_campo_texto(valor, campo=""):
    if not valor or str(valor).lower() == "none":
        return ""
    if isinstance(valor, list):
        return valor

    txt = str(valor).strip()

    if campo == "Seccion":
        return txt.upper()

    palabras = txt.lower().split()
    if not palabras:
        return ""

    preposiciones = ["de", "la", "el", "en", "y", "a", "con", "las", "los", "del", "al"]
    resultado = [palabras[0].capitalize()]
    for p in palabras[1:]:
        resultado.append(p if p in preposiciones else p.capitalize())

    return " ".join(resultado)


def convertir_a_pdf(pptx_data):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(pptx_data)
        p_in = tmp.name

    try:
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir",
             os.path.dirname(p_in), p_in],
            check=True
        )
        p_out = p_in.replace(".pptx", ".pdf")

        with open(p_out, "rb") as f:
            data = f.read()

        if os.path.exists(p_in):
            os.remove(p_in)
        if os.path.exists(p_out):
            os.remove(p_out)

        return data
    except Exception:
        return None


# ===============================
# INICIALIZACIÓN STREAMLIT
# ===============================

st.set_page_config(
    page_title="Provident Pro v43",
    layout="wide"
)

st.title("🚀 Generador Pro v43")

# ===============================
# CONFIGURACIÓN
# ===============================

if "config" not in st.session_state:
    if os.path.exists("config_app.json"):
        try:
            with open("config_app.json", "r") as f:
                st.session_state.config = json.load(f)
        except Exception:
            st.session_state.config = {"plantillas": {}, "columnas_visibles": []}
    else:
        st.session_state.config = {"plantillas": {}, "columnas_visibles": []}

TOKEN = "TU_TOKEN_AIRTABLE_AQUI"
headers = {"Authorization": f"Bearer {TOKEN}"}

# ===============================
# SIDEBAR
# ===============================

with st.sidebar:
    st.header("⚙️ Configuración")

    if st.button("💾 GUARDAR JSON"):
        with open("config_app.json", "w") as f:
            json.dump(st.session_state.config, f)
        st.success("Configuración guardada")

    st.markdown("---")

    try:
        r_b = requests.get(
            "https://api.airtable.com/v0/meta/bases",
            headers=headers
        )

        if r_b.status_code == 200:
            bases = {b["name"]: b["id"] for b in r_b.json().get("bases", [])}
            base_sel = st.selectbox("Seleccionar Base:", [""] + list(bases.keys()))

            if base_sel:
                base_id = bases[base_sel]
                r_t = requests.get(
                    f"https://api.airtable.com/v0/meta/bases/{base_id}/tables",
                    headers=headers
                )

                tablas = {t["name"]: t["id"] for t in r_t.json().get("tables", [])}
                tabla_sel = st.selectbox("Seleccionar Tabla:", list(tablas.keys()))

                if st.button("🔄 CARGAR REGISTROS"):
                    r_r = requests.get(
                        f"https://api.airtable.com/v0/{base_id}/{tablas[tabla_sel]}",
                        headers=headers
                    )
                    registros = r_r.json().get("records", [])

                    st.session_state.raw_data_original = registros
                    st.session_state.raw_records = [
                        {
                            "id": r["id"],
                            "fields": {
                                k: limpiar_campo_texto(v, k) if k != "Fecha" else v
                                for k, v in r["fields"].items()
                            },
                        }
                        for r in registros
                    ]

                    st.experimental_rerun()

    except Exception as e:
        st.error(f"Error Airtable: {e}")

# ===============================
# TABLA PRINCIPAL
# ===============================

if "raw_records" in st.session_state:

    df_base = pd.DataFrame([r["fields"] for r in st.session_state.raw_records])

    with st.sidebar:
        st.markdown("---")
        todas = list(df_base.columns)
        visibles = st.session_state.config.get("columnas_visibles", todas)
        visibles = [c for c in visibles if c in todas]

        columnas_sel = st.multiselect(
            "Campos visibles:",
            todas,
            default=visibles
        )

        st.session_state.config["columnas_visibles"] = columnas_sel

    df_view = df_base[columnas_sel].copy()

    for col in df_view.columns:
        if not df_view.empty and isinstance(df_view[col].iloc[0], list):
            df_view.drop(col, axis=1, inplace=True)

    df_view.insert(0, "Seleccionar", False)

    df_edit = st.data_editor(
        df_view,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Seleccionar": st.column_config.CheckboxColumn(
                "Seleccionar", default=False
            )
        },
    )

    idx_sel = df_edit.index[df_edit["Seleccionar"]].tolist()

    # ===============================
    # GENERACIÓN
    # ===============================

    if idx_sel:

        modo = st.radio("Acción:", ["Postales", "Reportes"], horizontal=True)
        carpeta_plantillas = os.path.join("Plantillas", modo.upper())
        AZUL = RGBColor(0, 176, 240)

        tipos = df_view.loc[idx_sel, "Tipo"].unique()
        plantillas_disp = [f for f in os.listdir(carpeta_plantillas) if f.endswith(".pptx")]

        for t in tipos:
            guardada = st.session_state.config.get("plantillas", {}).get(t)
            idx = plantillas_disp.index(guardada) if guardada in plantillas_disp else 0

            st.session_state.config.setdefault("plantillas", {})[t] = st.selectbox(
                f"Plantilla para {t}",
                plantillas_disp,
                index=idx,
                key=f"tpl_{t}"
            )

        if st.button("🔥 GENERAR SELECCIÓN", type="primary", use_container_width=True):

            progreso = st.progress(0)
            zip_buffer = BytesIO()

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:

                for i, idx in enumerate(idx_sel):
                    r_proc = st.session_state.raw_records[idx]["fields"]
                    r_orig = st.session_state.raw_data_original[idx]["fields"]

                    prs = Presentation(
                        os.path.join(
                            carpeta_plantillas,
                            st.session_state.config["plantillas"][r_proc["Tipo"]],
                        )
                    )

                    tags = {
                        "<<Tipo>>": r_proc.get("Tipo"),
                        "<<Sucursal>>": r_proc.get("Sucursal"),
                        "<<Seccion>>": r_proc.get("Seccion"),
                        "<<Confechor>>": f"{r_proc.get('Fecha')}, {r_proc.get('Hora')}",
                        "<<Concat>>": f"{r_proc.get('Punto de reunion') or r_proc.get('Ruta a seguir')}, {r_proc.get('Municipio')}",
                    }

                    for slide in prs.slides:

                        # IMÁGENES
                        for shp in list(slide.shapes):
                            if not shp.has_text_frame:
                                continue

                            for tag in [
                                "Foto de equipo", "Foto 01", "Foto 02",
                                "Foto 03", "Foto 04", "Foto 05",
                                "Foto 06", "Foto 07",
                                "Reporte firmado", "Lista de asistencia"
                            ]:
                                if f"<<{tag}>>" in shp.text_frame.text or shp.name == tag:
                                    adj = r_orig.get(tag)
                                    if adj and isinstance(adj, list):
                                        img = requests.get(adj[0]["url"]).content
                                        slide.shapes.add_picture(
                                            BytesIO(img),
                                            shp.left, shp.top,
                                            shp.width, shp.height
                                        )
                                        shp.element.getparent().remove(shp.element)

                        # TEXTO
                        for shp in slide.shapes:
                            if shp.has_text_frame:
                                for tag, val in tags.items():
                                    if tag in shp.text_frame.text:
                                        tf = shp.text_frame
                                        tf.clear()
                                        run = tf.paragraphs[0].add_run()
                                        run.text = str(val)
                                        run.font.bold = True
                                        run.font.color.rgb = AZUL
                                        run.font.size = Pt(14 if tag == "<<Sucursal>>" else 11)

                    salida = BytesIO()
                    prs.save(salida)

                    pdf = convertir_a_pdf(salida.getvalue())

                    if pdf:
                        nombre = f"{r_proc.get('Fecha')} - {r_proc.get('Tipo')} - {r_proc.get('Sucursal')}"
                        zf.writestr(
                            f"{modo}/{r_proc.get('Sucursal')}/{nombre}.pdf",
                            pdf
                        )

                    progreso.progress((i + 1) / len(idx_sel))

            st.success("✅ Generación completada")
            st.download_button(
                "📥 DESCARGAR ZIP",
                zip_buffer.getvalue(),
                "Provident_v43.zip",
                use_container_width=True
            )
