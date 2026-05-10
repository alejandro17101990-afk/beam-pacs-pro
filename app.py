import streamlit as st
from docx import Document
import speech_recognition as sr
import io
from openai import OpenAI
from datetime import datetime

# =========================================================
# CONFIGURACIÓN PREMIUM
# =========================================================

st.set_page_config(
    page_title="Lumen Core AI",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# ESTILO VISUAL ULTRA PREMIUM
# =========================================================

st.markdown("""
<style>

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(124,58,237,0.15), transparent 30%),
        radial-gradient(circle at bottom right, rgba(59,130,246,0.12), transparent 30%),
        #09090b;
    color: #f4f4f5;
}

/* HEADER */
header {
    visibility: hidden;
}

/* SIDEBAR */
[data-testid="stSidebar"] {
    background: rgba(12,12,15,0.95);
    border-right: 1px solid rgba(255,255,255,0.06);
    backdrop-filter: blur(20px);
}

/* TITLES */
h1, h2, h3 {
    letter-spacing: -0.03em;
}

/* CARDS */
.block-container {
    padding-top: 2rem;
}

/* CHAT MESSAGE */
[data-testid="stChatMessage"] {
    background: rgba(24,24,27,0.65);
    border: 1px solid rgba(255,255,255,0.05);
    border-radius: 18px;
    padding: 1rem;
    backdrop-filter: blur(14px);
    margin-bottom: 1rem;
}

/* INPUT */
[data-testid="stChatInput"] {
    background: rgba(24,24,27,0.9) !important;
    border: 1px solid rgba(255,255,255,0.08) !important;
    border-radius: 20px !important;
    padding: 0.5rem !important;
    backdrop-filter: blur(20px);
}

/* TEXTAREA */
textarea {
    color: white !important;
}

/* BUTTONS */
.stButton>button {
    background: linear-gradient(135deg, #7c3aed, #6366f1) !important;
    color: white !important;
    border: none !important;
    border-radius: 14px !important;
    padding: 0.75rem 1rem !important;
    font-weight: 600 !important;
    transition: all 0.25s ease;
    box-shadow: 0 0 25px rgba(124,58,237,0.35);
}

.stButton>button:hover {
    transform: translateY(-2px);
    box-shadow: 0 0 35px rgba(124,58,237,0.55);
}

/* SELECTBOX */
.stSelectbox div[data-baseweb="select"] {
    background: rgba(24,24,27,0.8) !important;
    border-radius: 14px !important;
    border: 1px solid rgba(255,255,255,0.06) !important;
}

/* FILE UPLOADER */
.stFileUploader > div {
    background: rgba(24,24,27,0.7);
    border-radius: 16px;
    border: 1px dashed rgba(255,255,255,0.08);
}

/* DOWNLOAD BUTTON */
.stDownloadButton>button {
    background: linear-gradient(135deg,#2563eb,#4f46e5) !important;
}

/* DIVIDERS */
hr {
    border-color: rgba(255,255,255,0.06);
}

/* METRIC CARDS */
.metric-card {
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.05);
    border-radius: 18px;
    padding: 1rem;
    backdrop-filter: blur(12px);
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# FUNCIONES
# =========================================================

def leer_word_multimodal(file):
    doc = Document(file)
    contenido = []

    for para in doc.paragraphs:
        if para.text.strip():
            contenido.append(para.text)

    for table in doc.tables:
        contenido.append("\\n[TABLA DETECTADA]")
        for row in table.rows:
            fila = " | ".join([cell.text for cell in row.cells])
            contenido.append(fila)

    return '\\n'.join(contenido)

def generar_docx(texto):
    doc = Document()

    titulo = doc.add_heading('Informe Radiológico', level=1)
    titulo.style.font.size = 240000

    for linea in texto.split('\\n'):
        doc.add_paragraph(linea)

    bio = io.BytesIO()
    doc.save(bio)

    return bio.getvalue()

def transcribir_audio(audio_file):
    r = sr.Recognizer()

    with sr.AudioFile(audio_file) as source:
        audio = r.record(source)

    try:
        texto = r.recognize_google(audio, language="es-MX")
        return texto
    except:
        return "[No fue posible transcribir el audio]"

# =========================================================
# SESSION STATE
# =========================================================

if "mensajes" not in st.session_state:
    st.session_state.mensajes = []

if "ultimo_reporte" not in st.session_state:
    st.session_state.ultimo_reporte = ""

# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown("# 🧠 Lumen Core AI")
    st.caption("Radiology Intelligence Workspace")

    st.markdown("---")

    try:
        api_key = st.secrets["deepseek_key"]
        st.success("● DeepSeek conectado")
    except:
        api_key = st.text_input(
            "DeepSeek API Key",
            type="password"
        )

    st.markdown("---")

    st.markdown("### Modalidad")

    modalidad = st.selectbox(
        "Seleccionar estudio",
        [
            "Resonancia Magnética",
            "Tomografía Computarizada",
            "Radiografía",
            "Ultrasonido",
            "PET-CT",
            "Mastografía",
            "Neurorradiología",
            "Musculoesquelético",
            "Cardiotorácico"
        ]
    )

    estilo = st.selectbox(
        "Perfil radiológico",
        [
            "Académico",
            "Alta Especialidad",
            "Conciso",
            "Fellowship",
            "Internacional"
        ]
    )

    st.markdown("---")

    st.markdown("### Plantilla Base")

    archivo_plantilla = st.file_uploader(
        "Subir plantilla .docx",
        type=["docx"]
    )

    plantilla_procesada = ""

    if archivo_plantilla:
        plantilla_procesada = leer_word_multimodal(archivo_plantilla)
        st.success("Plantilla cargada")

    st.markdown("---")

    st.markdown("### Dictado Inteligente")

    audio_dictado = st.audio_input("Grabar hallazgos")

    st.markdown("---")

    st.markdown(f"""
    <div class="metric-card">
        <h4>Estado</h4>
        <p>🟢 Workspace operativo</p>
        <p>{datetime.now().strftime("%d/%m/%Y %H:%M")}</p>
    </div>
    """, unsafe_allow_html=True)

# =========================================================
# HEADER PRINCIPAL
# =========================================================

col1, col2 = st.columns([3,1])

with col1:
    st.title("Radiology Report Generator")

with col2:
    st.markdown("""
    <div style="
        text-align:right;
        padding-top:20px;
        color:#a1a1aa;
        font-size:14px;
    ">
    AI-Powered Reporting
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# =========================================================
# EXPORTACIÓN
# =========================================================

if st.session_state.ultimo_reporte:

    st.download_button(
        label="📥 Exportar Informe Word",
        data=generar_docx(st.session_state.ultimo_reporte),
        file_name="Informe_Radiologico.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# =========================================================
# HISTORIAL CHAT
# =========================================================

for mensaje in st.session_state.mensajes:

    with st.chat_message(mensaje["role"]):
        st.markdown(mensaje["content"])

# =========================================================
# IA ENGINE
# =========================================================

if api_key:

    client = OpenAI(
        api_key=api_key,
        base_url="https://api.deepseek.com",
        timeout=120.0
    )

    prompt_usuario = st.chat_input(
        "Describe hallazgos, modifica conclusiones o dicta el informe..."
    )

    if prompt_usuario or audio_dictado:

        texto_transcrito = ""

        if audio_dictado:
            texto_transcrito = transcribir_audio(audio_dictado)

        entrada = ""

        if texto_transcrito:
            entrada += f"🎙️ Dictado:\\n{texto_transcrito}\\n\\n"

        if prompt_usuario:
            entrada += f"⌨️ Instrucción:\\n{prompt_usuario}"

        st.session_state.mensajes.append({
            "role":"user",
            "content":entrada
        })

        with st.chat_message("user"):
            st.markdown(entrada)

        with st.chat_message("assistant"):

            with st.spinner("Analizando hallazgos y estructurando informe..."):

                prompt_sistema = f"""
Eres Lumen Core AI.

Un asistente radiológico de clase mundial especializado en generación avanzada de informes médicos.

MODALIDAD:
{modalidad}

ESTILO:
{estilo}

OBJETIVOS:
- Redacción médica elegante
- Terminología fellowship
- Alta precisión anatómica
- Conclusiones sofisticadas
- Evitar redundancias
- Lenguaje académico profesional

ESTRUCTURA:
Técnica:
Hallazgos:
Impresión diagnóstica:

REGLAS:
- No inventar hallazgos
- Mantener coherencia clínica
- Usar lenguaje radiológico avanzado
- Corregir gramática automáticamente
- Optimizar claridad diagnóstica

PLANTILLA:
{plantilla_procesada if plantilla_procesada else "Sin plantilla personalizada"}
"""

                try:

                    response = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[
                            {
                                "role":"system",
                                "content":prompt_sistema
                            },
                            {
                                "role":"user",
                                "content":entrada
                            }
                        ],
                        temperature=0.15,
                        max_tokens=3500
                    )

                    respuesta = response.choices[0].message.content

                    st.markdown(respuesta)

                    st.session_state.mensajes.append({
                        "role":"assistant",
                        "content":respuesta
                    })

                    st.session_state.ultimo_reporte = respuesta

                    st.rerun()

                except Exception as e:

                    st.error(f"Error DeepSeek: {e}")

else:

    st.info("Introduce tu API Key de DeepSeek para iniciar.")
