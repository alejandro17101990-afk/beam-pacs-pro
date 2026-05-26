import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile
import io
import os
import re

# =========================================================
# CONFIG
# =========================================================

st.set_page_config(
    page_title="AURA Radiology AI",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================================================
# THEME
# =========================================================

THEME = {
    "bg": "#0b0f14",
    "panel": "#11161d",
    "panel2": "#151b23",
    "border": "#202833",
    "text": "#e8edf2",
    "muted": "#7d8b99",
    "accent": "#3ba4ff",
}

# =========================================================
# CSS
# =========================================================

st.markdown(f"""
<style>

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

html, body, .stApp {{
    background: {THEME["bg"]};
    color: {THEME["text"]};
    font-family: 'Inter', sans-serif;
}}

header, footer, #MainMenu {{
    visibility: hidden;
}}

.block-container {{
    padding-top: 1.5rem;
    padding-bottom: 1rem;
    max-width: 1600px;
}}

h1,h2,h3,h4 {{
    color: {THEME["text"]};
}}

.stTextArea textarea {{
    background: {THEME["panel"]} !important;
    border: 1px solid {THEME["border"]} !important;
    border-radius: 16px !important;
    color: {THEME["text"]} !important;
    font-size: 15px !important;
    line-height: 1.7 !important;
    padding: 18px !important;
}}

.stTextArea textarea:focus {{
    border: 1px solid {THEME["accent"]} !important;
    box-shadow: none !important;
}}

.stButton button {{
    background: {THEME["panel2"]} !important;
    border: 1px solid {THEME["border"]} !important;
    color: {THEME["text"]} !important;
    border-radius: 12px !important;
    height: 48px !important;
    font-weight: 500 !important;
    transition: 0.2s;
}}

.stButton button:hover {{
    border: 1px solid {THEME["accent"]} !important;
    color: white !important;
}}

.stDownloadButton button {{
    background: {THEME["accent"]} !important;
    color: white !important;
    border-radius: 12px !important;
    border: none !important;
    height: 48px !important;
}}

[data-testid="stFileUploader"] {{
    background: {THEME["panel"]};
    border: 1px dashed {THEME["border"]};
    border-radius: 14px;
    padding: 10px;
}}

hr {{
    border-color: {THEME["border"]};
}}

.section-card {{
    background: {THEME["panel"]};
    border: 1px solid {THEME["border"]};
    border-radius: 18px;
    padding: 20px;
}}

.title {{
    font-size: 34px;
    font-weight: 600;
    letter-spacing: -0.03em;
}}

.subtitle {{
    color: {THEME["muted"]};
    font-size: 14px;
    margin-top: -8px;
}}

.small-label {{
    font-size: 12px;
    color: {THEME["muted"]};
    margin-bottom: 6px;
}}

</style>
""", unsafe_allow_html=True)

# =========================================================
# MODELS
# =========================================================

MODELS = {
    "DeepSeek Chat": {
        "url": "https://api.deepseek.com",
        "id": "deepseek-chat"
    },
    "GPT-4o Mini": {
        "url": None,
        "id": "gpt-4o-mini"
    },
    "GPT-4.1 Mini": {
        "url": None,
        "id": "gpt-4.1-mini"
    }
}

# =========================================================
# SESSION STATE
# =========================================================

defaults = {
    "dictado": "",
    "reporte": "",
    "defs": "",
    "modelo": "DeepSeek Chat",
    "audio_id": None,
    "historial": []
}

for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# =========================================================
# API KEY
# =========================================================

try:
    api_key = st.secrets["deepseek_key"]
except:
    api_key = os.environ.get("OPENAI_API_KEY", "")

# =========================================================
# HELPERS
# =========================================================

def client_ia():

    cfg = MODELS[st.session_state.modelo]

    if cfg["url"]:
        return OpenAI(
            api_key=api_key,
            base_url=cfg["url"]
        )

    return OpenAI(api_key=api_key)

def parse_sections(text):

    sections = {}
    current = None
    buffer = []

    for line in text.split("\n"):

        s = line.strip()

        if s.isupper() and 4 < len(s) < 80:

            if current:
                sections[current] = "\n".join(buffer).strip()

            current = s
            buffer = []

        else:

            if current:
                buffer.append(s)

    if current:
        sections[current] = "\n".join(buffer).strip()

    return sections

def generar_docx(texto):

    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    for line in texto.split("\n"):

        s = line.strip()

        if not s:
            doc.add_paragraph()
            continue

        if s.isupper() and len(s) < 80:

            h = doc.add_heading(s, level=1)
            h.alignment = WD_ALIGN_PARAGRAPH.LEFT

        else:

            doc.add_paragraph(s)

    bio = io.BytesIO()
    doc.save(bio)

    return bio.getvalue()

def transcribir_audio(audio):

    cfg = MODELS[st.session_state.modelo]

    client = OpenAI(
        api_key=api_key,
        base_url=cfg["url"]
    ) if cfg["url"] else OpenAI(api_key=api_key)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp:

        tmp.write(audio.read())
        path = tmp.name

    with open(path, "rb") as f:

        result = client.audio.transcriptions.create(
            model="whisper-1",
            file=f,
            language="es"
        )

    os.unlink(path)

    return result.text

# =========================================================
# HEADER
# =========================================================

st.markdown("""
<div class="title">AURA</div>
<div class="subtitle">
Radiology Intelligence System
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# =========================================================
# TOP SETTINGS
# =========================================================

top1, top2, top3 = st.columns([1,1,1])

with top1:

    st.markdown('<div class="small-label">MODELO IA</div>', unsafe_allow_html=True)

    modelo = st.selectbox(
        "modelo",
        list(MODELS.keys()),
        label_visibility="collapsed",
        index=list(MODELS.keys()).index(st.session_state.modelo)
    )

    st.session_state.modelo = modelo

with top2:

    st.markdown('<div class="small-label">PLANTILLA .DOCX</div>', unsafe_allow_html=True)

    plantilla = st.file_uploader(
        "docx",
        type=["docx"],
        label_visibility="collapsed"
    )

with top3:

    st.markdown('<div class="small-label">ACCIONES</div>', unsafe_allow_html=True)

    if st.button("Nuevo informe", use_container_width=True):

        st.session_state.dictado = ""
        st.session_state.reporte = ""
        st.session_state.defs = ""

        st.rerun()

st.markdown("<hr>", unsafe_allow_html=True)

# =========================================================
# MAIN LAYOUT
# =========================================================

col1, col2 = st.columns([1, 1.25])

# =========================================================
# LEFT PANEL
# =========================================================

with col1:

    st.markdown("""
    <div class="section-card">
    """, unsafe_allow_html=True)

    st.markdown("### Dictado")

    audio = st.audio_input(
        "audio",
        label_visibility="collapsed"
    )

    if audio:

        audio_id = hash(audio.read())
        audio.seek(0)

        if audio_id != st.session_state.audio_id:

            if api_key:

                with st.spinner("Transcribiendo..."):

                    txt = transcribir_audio(audio)

                    st.session_state.dictado += " " + txt
                    st.session_state.audio_id = audio_id

                    st.rerun()

    dictado = st.text_area(
        "dictado",
        value=st.session_state.dictado,
        height=500,
        label_visibility="collapsed",
        placeholder="Dicta o escribe hallazgos radiológicos..."
    )

    st.session_state.dictado = dictado

    c1, c2 = st.columns(2)

    with c1:

        generar = st.button(
            "Generar informe",
            use_container_width=True
        )

    with c2:

        limpiar = st.button(
            "Limpiar",
            use_container_width=True
        )

        if limpiar:

            st.session_state.dictado = ""
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# REPORT GENERATION
# =========================================================

if generar:

    if not api_key:

        st.error("API Key requerida.")

    elif not st.session_state.dictado.strip():

        st.warning("Ingresa hallazgos o dictado.")

    else:

        client = client_ia()

        model_id = MODELS[st.session_state.modelo]["id"]

        PROMPT = f"""
Eres AURA, un sistema experto de interpretación radiológica.

Analiza el dictado y detecta automáticamente:

- modalidad
- región anatómica
- lateralidad
- protocolo
- clasificaciones radiológicas relevantes

Genera un informe radiológico estructurado profesional.

REGLAS:

- Lenguaje médico experto
- Morfológicamente preciso
- Sin redundancias
- Sin lenguaje ambiguo
- Sin markdown
- Títulos en MAYÚSCULAS

ESTRUCTURA:

INDICACIÓN
TÉCNICA
HALLAZGOS
IMPRESIÓN DIAGNÓSTICA

DICTADO:

{st.session_state.dictado}
"""

        with st.spinner("Generando informe..."):

            response = client.chat.completions.create(
                model=model_id,
                messages=[
                    {
                        "role": "system",
                        "content": PROMPT
                    }
                ],
                temperature=0.1,
                max_tokens=2500
            )

            report = response.choices[0].message.content

            st.session_state.reporte = report

            st.session_state.historial.insert(0, report)

            if len(st.session_state.historial) > 15:
                st.session_state.historial = st.session_state.historial[:15]

            st.rerun()

# =========================================================
# RIGHT PANEL
# =========================================================

with col2:

    st.markdown("""
    <div class="section-card">
    """, unsafe_allow_html=True)

    st.markdown("### Informe")

    reporte = st.text_area(
        "reporte",
        value=st.session_state.reporte,
        height=700,
        label_visibility="collapsed"
    )

    st.session_state.reporte = reporte

    sec1, sec2, sec3 = st.columns(3)

    # -----------------------------------------------------

    with sec1:

        if st.button(
            "Optimizar conclusión",
            use_container_width=True
        ):

            if st.session_state.reporte:

                client = client_ia()

                model_id = MODELS[st.session_state.modelo]["id"]

                prompt = f"""
Mejora únicamente la IMPRESIÓN DIAGNÓSTICA.

Hazla:

- más precisa
- más clínica
- más accionable

Conserva el resto intacto.

INFORME:

{st.session_state.reporte}
"""

                with st.spinner("Optimizando..."):

                    response = client.chat.completions.create(
                        model=model_id,
                        messages=[
                            {
                                "role": "user",
                                "content": prompt
                            }
                        ],
                        temperature=0.2,
                        max_tokens=2500
                    )

                    st.session_state.reporte = response.choices[0].message.content

                    st.rerun()

    # -----------------------------------------------------

    with sec2:

        if st.button(
            "Definiciones",
            use_container_width=True
        ):

            if st.session_state.reporte:

                client = client_ia()

                model_id = MODELS[st.session_state.modelo]["id"]

                prompt = f"""
Analiza el informe radiológico y proporciona:

- Clasificaciones usadas
- Definiciones
- Correlación clínica
- Referencias

INFORME:

{st.session_state.reporte}
"""

                with st.spinner("Analizando..."):

                    response = client.chat.completions.create(
                        model=model_id,
                        messages=[
                            {
                                "role": "user",
                                "content": prompt
                            }
                        ],
                        temperature=0.1,
                        max_tokens=2000
                    )

                    st.session_state.defs = response.choices[0].message.content

                    st.rerun()

    # -----------------------------------------------------

    with sec3:

        if st.session_state.reporte:

            docx_file = generar_docx(
                st.session_state.reporte
            )

            st.download_button(
                "Exportar DOCX",
                data=docx_file,
                file_name="AURA_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# DEFINITIONS PANEL
# =========================================================

if st.session_state.defs:

    st.markdown("<br>", unsafe_allow_html=True)

    with st.expander(
        "Definiciones · Clasificaciones · Referencias",
        expanded=True
    ):

        st.text_area(
            "defs",
            value=st.session_state.defs,
            height=350,
            label_visibility="collapsed"
        )

# =========================================================
# HISTORY
# =========================================================

if st.session_state.historial:

    st.markdown("<br>", unsafe_allow_html=True)

    with st.expander("Historial reciente", expanded=False):

        for i, rep in enumerate(st.session_state.historial):

            if st.button(
                f"Informe #{i+1}",
                key=f"hist_{i}",
                use_container_width=True
            ):

                st.session_state.reporte = rep
                st.rerun()
