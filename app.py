import streamlit as st
from docx import Document
import speech_recognition as sr
import io
from openai import OpenAI

# ==========================================
# CONFIG
# ==========================================
st.set_page_config(page_title="Beam AI | Radiology OS", layout="wide", initial_sidebar_state="collapsed")

# ==========================================
# CLASIFICACIONES RADIOLÓGICAS
# ==========================================
CLASIFICACIONES = {
    "Menisco · Stoller": [
        ("0", "Normal — señal homogénea"),
        ("I", "Señal focal intrameniscal, no articular"),
        ("II", "Señal lineal, no alcanza superficie"),
        ("III", "Señal alcanza superficie articular → desgarro"),
    ],
    "Cartílago · ICRS": [
        ("0", "Normal"),
        ("I", "Fibrilación o fisuras superficiales"),
        ("II", "Defecto <50% del grosor"),
        ("III", "Defecto >50% del grosor"),
        ("IV", "Hueso subcondral expuesto"),
    ],
    "Artrosis · Kellgren-Lawrence": [
        ("I", "Posible osteofito marginal"),
        ("II", "Osteofito definido, sin pinzamiento"),
        ("III", "Pinzamiento moderado"),
        ("IV", "Pinzamiento grave, esclerosis"),
    ],
    "LCA · Hope & Feagin": [
        ("Parcial", "Fibras continuas, señal aumentada"),
        ("Completa", "Discontinuidad total de fibras"),
        ("Crónica", "Fibras atróficas o laxas"),
    ],
    "Columna · Pfirrmann": [
        ("I", "Núcleo brillante, homogéneo"),
        ("II", "Señal alta, zona no clara"),
        ("III", "Señal gris, distinción borrosa"),
        ("IV", "Señal baja, sin distinción"),
        ("V", "Señal muy baja, sin espacio discal"),
    ],
    "TIRADS · ACR (Tiroides)": [
        ("1", "Benigno — sin nódulo"),
        ("2", "No sospechoso"),
        ("3", "Levemente sospechoso"),
        ("4", "Moderadamente sospechoso"),
        ("5", "Altamente sospechoso de malignidad"),
    ],
    "BI-RADS · ACR (Mama)": [
        ("0", "Estudio incompleto"),
        ("1", "Negativo"),
        ("2", "Hallazgo benigno"),
        ("3", "Probablemente benigno"),
        ("4", "Sospechoso"),
        ("5", "Altamente sugestivo de malignidad"),
        ("6", "Malignidad conocida"),
    ],
    "Hígado · LI-RADS": [
        ("LR-1", "Definitivamente benigno"),
        ("LR-2", "Probablemente benigno"),
        ("LR-3", "Intermedio"),
        ("LR-4", "Probablemente HCC"),
        ("LR-5", "Definitivamente HCC"),
        ("LR-M", "Probable malignidad no-HCC"),
    ],
}

SUGERENCIAS_CONTEXTO = {
    "Resonancia Magnética": [
        "Stoller III → desgarro meniscal confirmado",
        "Extrusión >3 mm → relevancia clínica significativa",
        "ICRS III → candidato a condroplastia",
        "Gonartrosis K-L III-IV → valorar reemplazo",
        "Edema óseo subcondral → lesión activa",
        "Rotura LCA + contusión ósea → pivot shift",
    ],
    "Tomografía Computarizada": [
        "Escala Hounsfield: hueso 700 UH, agua 0 UH",
        "Adenopatía >1 cm eje corto → significativa",
        "Nódulo pulmonar sólido >6 mm → seguimiento Fleischner",
        "Embolia pulmonar → escala WELLS / score GENEVA",
    ],
    "Radiografía": [
        "Kellgren-Lawrence → estadificación artrosis",
        "Cobb >10° → escoliosis diagnóstica",
        "Índice cardiotorácico >0.5 → cardiomegalia",
    ],
    "Ultrasonido": [
        "TIRADS 4-5 → considerar BAAF",
        "BI-RADS US 4A/4B/4C → biopsia según riesgo",
        "Resistividad >0.7 → vasculatura maligna",
    ],
    "PET-CT": [
        "SUVmax >2.5 → actividad metabólica significativa",
        "LI-RADS 5 → HCC definitivo",
        "Captación focal vs difusa → diferencial maligno/inflamatorio",
    ],
}

# ==========================================
# CSS CINEMATOGRÁFICO PACS
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=Syne:wght@400;600;700&display=swap');

.stApp {
    background: #070a10 !important;
    font-family: 'IBM Plex Mono', monospace !important;
}
.block-container { padding: 0 !important; max-width: 100% !important; }
header { visibility: hidden !important; }
[data-testid="stToolbar"] { display: none !important; }
footer { display: none !important; }

/* TOPBAR */
.beam-topbar {
    background: #0b0f1a;
    border-bottom: 1px solid #1a2333;
    padding: 10px 24px;
    display: flex;
    align-items: center;
    gap: 16px;
    position: sticky;
    top: 0;
    z-index: 999;
}
.beam-logo {
    font-family: 'Syne', sans-serif;
    font-weight: 700;
    font-size: 15px;
    color: #e2edf8;
    letter-spacing: 0.1em;
    display: flex;
    align-items: center;
    gap: 8px;
}
.logo-dot {
    width: 8px; height: 8px; border-radius: 50%;
    background: #3b8bd4;
    display: inline-block;
    box-shadow: 0 0 8px #3b8bd4;
}
.beam-badge {
    font-size: 11px; color: #4a9ed4;
    background: #0d1e30; border: 1px solid #1a3a58;
    border-radius: 4px; padding: 2px 10px;
    font-family: 'IBM Plex Mono', monospace;
}
.beam-status {
    margin-left: auto;
    font-size: 11px; color: #3d6a5a;
    font-family: 'IBM Plex Mono', monospace;
    display: flex; align-items: center; gap: 6px;
}
.status-dot {
    width: 6px; height: 6px; border-radius: 50%;
    background: #2ecc71; display: inline-block;
    box-shadow: 0 0 5px #2ecc71;
}

/* PANEL LABELS */
.panel-label {
    font-size: 9px !important;
    letter-spacing: 0.2em !important;
    color: #2a4a6a !important;
    text-transform: uppercase !important;
    font-family: 'IBM Plex Mono', monospace !important;
    margin-bottom: 4px !important;
    margin-top: 0 !important;
}

/* SELECTBOX */
[data-testid="stSelectbox"] > div > div {
    background: #0f1520 !important;
    border: 1px solid #1a2f48 !important;
    border-radius: 6px !important;
    color: #7ab0cc !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 12px !important;
}
[data-testid="stSelectbox"] > div > div:hover {
    border-color: #2a5a8a !important;
}

/* TEXTAREA */
.stTextArea textarea {
    background: #0f1520 !important;
    border: 1px solid #1a2f48 !important;
    border-radius: 8px !important;
    color: #b8d0e8 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 12px !important;
    line-height: 1.7 !important;
}
.stTextArea textarea:focus {
    border-color: #2a5580 !important;
    box-shadow: none !important;
}

/* AUDIO */
[data-testid="stAudioInput"] {
    background: #0f1a28 !important;
    border: 1px solid #1a3a58 !important;
    border-radius: 8px !important;
}

/* FILE UPLOADER */
[data-testid="stFileUploader"] {
    background: #0f1520 !important;
    border: 1px dashed #1d3550 !important;
    border-radius: 8px !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: #2a5a8a !important;
}
[data-testid="stFileUploader"] * { color: #3d6090 !important; font-family: 'IBM Plex Mono', monospace !important; font-size: 11px !important; }

/* BOTÓN PRINCIPAL */
.btn-generar > div > button {
    background: #1a3a60 !important;
    border: 1px solid #2a5a90 !important;
    color: #7ac2f0 !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 13px !important;
    letter-spacing: 0.08em !important;
    border-radius: 8px !important;
    padding: 0.7rem 1rem !important;
    transition: all 0.2s ease !important;
    width: 100% !important;
}
.btn-generar > div > button:hover {
    background: #1f4878 !important;
    border-color: #3a78c0 !important;
}

/* BOTONES SECUNDARIOS */
.stButton > button {
    background: #0d1828 !important;
    border: 1px solid #1a3050 !important;
    color: #4a8abf !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 11px !important;
    border-radius: 6px !important;
    transition: all 0.15s ease !important;
}
.stButton > button:hover {
    background: #102038 !important;
    border-color: #2a5a8a !important;
    color: #7ac2f0 !important;
}

/* EXPANDER (Clasificaciones) */
[data-testid="stExpander"] {
    background: #0b0f1a !important;
    border: 1px solid #1a2840 !important;
    border-radius: 8px !important;
    margin-bottom: 4px !important;
}
[data-testid="stExpander"] summary {
    color: #4a8abf !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 11px !important;
    letter-spacing: 0.05em !important;
}
[data-testid="stExpander"] summary:hover { color: #7ac2f0 !important; }

/* CHIPS DE SUGERENCIAS */
.sug-chip-container {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    margin-bottom: 10px;
}
.sug-chip {
    font-size: 11px;
    font-family: 'IBM Plex Mono', monospace;
    color: #4a8abf;
    background: #0a1828;
    border: 1px solid #1a3a58;
    padding: 4px 10px;
    border-radius: 4px;
    display: inline-block;
}

/* CLASIF ITEMS */
.clasif-item {
    background: #0a1020;
    border: 1px solid #111e30;
    border-radius: 5px;
    padding: 5px 9px;
    margin-bottom: 3px;
    font-size: 11px;
    font-family: 'IBM Plex Mono', monospace;
    color: #3d6090;
    line-height: 1.5;
    cursor: default;
}
.clasif-item.active {
    background: #0d1e32;
    border-color: #1d4a78;
    color: #7ac2f0;
}
.clasif-grade { color: #3b8bd4; font-weight: 500; }

/* BARRA DE COMPLETITUD */
.completitud-bar-bg {
    background: #0f1828;
    border-radius: 3px;
    height: 4px;
    margin: 4px 0 2px;
    overflow: hidden;
}
.completitud-bar-fill {
    height: 100%;
    background: linear-gradient(90deg, #1a4a80, #3b8bd4);
    border-radius: 3px;
    transition: width 0.4s ease;
}

/* DIVIDERS */
hr { border-color: #111a28 !important; margin: 8px 0 !important; }

/* SPINNER */
[data-testid="stSpinner"] { color: #4a8abf !important; }

/* DOWNLOAD BUTTON */
[data-testid="stDownloadButton"] > button {
    background: #0d1a28 !important;
    border: 1px solid #1a3a58 !important;
    color: #3a8abf !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 11px !important;
    border-radius: 6px !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #102030 !important;
    color: #7ac2f0 !important;
}

/* METRICS / COLUMNAS */
[data-testid="column"] {
    background: transparent !important;
}

/* SCROLLBAR */
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: #070a10; }
::-webkit-scrollbar-thumb { background: #1a2f48; border-radius: 2px; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# TOPBAR
# ==========================================
st.markdown("""
<div class="beam-topbar">
    <div class="beam-logo">
        <span class="logo-dot"></span> BEAM AI
    </div>
    <span class="beam-badge">v2.1 · MSK · Neuro · TX · PACS-Mode</span>
    <div class="beam-status">
        <span class="status-dot"></span> DeepSeek · en línea
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# ESTADO
# ==========================================
if 'dictado_actual' not in st.session_state: st.session_state.dictado_actual = ""
if 'reporte_final' not in st.session_state: st.session_state.reporte_final = ""
if 'clasif_activas' not in st.session_state: st.session_state.clasif_activas = {}

# ==========================================
# HELPERS
# ==========================================
def leer_plantilla(file):
    doc = Document(file)
    return '\n'.join([p.text for p in doc.paragraphs if p.text.strip()])

def generar_docx(texto):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    for linea in texto.split('\n'):
        p = doc.add_paragraph(linea)
        if linea.isupper() and linea.strip():
            p.runs[0].bold = True if p.runs else None
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def transcribir_voz(audio_file):
    r = sr.Recognizer()
    with sr.AudioFile(audio_file) as source:
        try: return r.recognize_google(r.record(source), language="es-MX")
        except: return ""

def calcular_completitud(texto):
    secciones = ["TÉCNICA", "HALLAZGOS", "IMPRESIÓN"]
    encontradas = sum(1 for s in secciones if s in texto.upper())
    palabras = len(texto.split())
    score = min(100, int((encontradas / 3) * 60 + min(palabras / 150, 1) * 40))
    return score

# ==========================================
# API KEY
# ==========================================
try: api_key = st.secrets["deepseek_key"]
except: api_key = ""

# ==========================================
# LAYOUT PRINCIPAL
# ==========================================
col_izq, col_centro, col_der = st.columns([1, 2.4, 0.9], gap="small")

# ─────────────────────────────────────────
# COLUMNA IZQUIERDA — Control
# ─────────────────────────────────────────
with col_izq:
    st.markdown("<br>", unsafe_allow_html=True)

    if not api_key:
        api_key = st.text_input("🔑 API Key DeepSeek", type="password", label_visibility="collapsed",
                                placeholder="sk-... DeepSeek API Key")

    st.markdown('<p class="panel-label">MODALIDAD</p>', unsafe_allow_html=True)
    modalidad = st.selectbox("Modalidad", [
        "Resonancia Magnética", "Tomografía Computarizada",
        "Radiografía", "Ultrasonido", "PET-CT"
    ], label_visibility="collapsed")

    st.markdown('<p class="panel-label">REGIÓN ANATÓMICA</p>', unsafe_allow_html=True)
    region = st.selectbox("Región", [
        "Rodilla", "Columna lumbar", "Columna cervical", "Hombro",
        "Cadera", "Tobillo / Pie", "Muñeca / Mano", "Codo",
        "Cerebro", "Columna dorsal", "Tórax", "Abdomen / Pelvis",
        "Mama", "Tiroides", "Hígado"
    ], label_visibility="collapsed")

    st.markdown('<p class="panel-label">PLANTILLA BASE</p>', unsafe_allow_html=True)
    archivo_base = st.file_uploader("Plantilla", type=["docx"], label_visibility="collapsed")
    plantilla_txt = leer_plantilla(archivo_base) if archivo_base else ""

    st.markdown('<p class="panel-label">DICTADO DE VOZ</p>', unsafe_allow_html=True)
    audio_data = st.audio_input("Voz", label_visibility="collapsed")
    if audio_data:
        nuevo = transcribir_voz(audio_data)
        if nuevo and nuevo not in st.session_state.dictado_actual:
            st.session_state.dictado_actual += " " + nuevo

    st.markdown('<p class="panel-label">SEÑAL / DICTADO</p>', unsafe_allow_html=True)
    dictado_verificable = st.text_area("Dictado", value=st.session_state.dictado_actual,
                                       height=130, label_visibility="collapsed",
                                       placeholder="Ej: Articulaciones cigapofisarias, cambios osteocondrales, grasa de Hoffa...")

    st.markdown('<p class="panel-label">DIRECTRICES</p>', unsafe_allow_html=True)
    instrucciones_estilo = st.text_area("Directrices", height=70,
                                        label_visibility="collapsed",
                                        value="Lenguaje médico experto. Infiere diagnósticos. Expande clasificaciones. Sin asteriscos.")

    st.markdown('<div class="btn-generar">', unsafe_allow_html=True)
    procesar = st.button("⬡  PROCESAR INFORME", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    col_b1, col_b2 = st.columns(2)
    with col_b1:
        if st.button("⌫ Purgar", use_container_width=True):
            st.session_state.dictado_actual = ""
            st.rerun()
    with col_b2:
        if st.button("⟳ Limpiar", use_container_width=True):
            st.session_state.reporte_final = ""
            st.rerun()

# ─────────────────────────────────────────
# PROCESAMIENTO IA
# ─────────────────────────────────────────
if procesar:
    if api_key and dictado_verificable.strip():
        client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

        clasif_contexto = ""
        if st.session_state.clasif_activas:
            clasif_contexto = "CLASIFICACIONES SELECCIONADAS POR EL RADIÓLOGO:\n"
            for k, v in st.session_state.clasif_activas.items():
                clasif_contexto += f"- {k}: {v}\n"

        prompt = f"""
Eres Beam AI, asistente experto en interpretación radiológica. Redacta un informe de {modalidad} para {region}.

LÓGICA DE SÍNTESIS:
1. SÍNTESIS: Agrupa e infiere diagnósticos a partir de los hallazgos descritos.
2. EXPANSIÓN: Si recibes diagnósticos (ej. Gonartrosis IV), redacta hallazgos anatómicos esperados.
3. CLASIFICACIONES: Usa las clasificaciones activas del radiólogo si se proporcionan.
4. TERMINOLOGÍA: Respeta terminología anatómica avanzada de MSK/Neuro/TX.
5. IMPRESIÓN: Haz la impresión diagnóstica elegante, concluyente y clínicamente accionable.

{clasif_contexto}

FORMATO DEL INFORME:
- TÍTULOS EN MAYÚSCULAS
- SIN asteriscos ni markdown
- SUBTÍTULOS de secciones en mayúsculas
- Redacción fluida y profesional
- Plantilla base a respetar: {plantilla_txt if plantilla_txt else "Técnica / Hallazgos / Impresión diagnóstica"}

REGLAS DE ESTILO: {instrucciones_estilo}

DICTADO / INPUT DEL RADIÓLOGO:
{dictado_verificable}
"""
        with st.spinner("Sintetizando modelo de datos..."):
            try:
                res = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "system", "content": prompt}],
                    temperature=0.1
                )
                st.session_state.reporte_final = res.choices[0].message.content
                st.rerun()
            except Exception as e:
                st.error(f"Error de enlace: {e}")
    elif not api_key:
        st.warning("Ingresa tu API Key de DeepSeek.")
    else:
        st.warning("Ingresa dictado o descripción clínica.")

# ─────────────────────────────────────────
# COLUMNA CENTRAL — Editor
# ─────────────────────────────────────────
with col_centro:
    st.markdown("<br>", unsafe_allow_html=True)

    # Sugerencias contextuales
    sugs = SUGERENCIAS_CONTEXTO.get(modalidad, [])
    if sugs:
        chips_html = '<div class="sug-chip-container">'
        chips_html += '<span style="font-size:10px;color:#2a4a6a;font-family:\'IBM Plex Mono\',monospace;margin-right:4px;">Sugerencias IA →</span>'
        for s in sugs:
            chips_html += f'<span class="sug-chip">{s}</span>'
        chips_html += '</div>'
        st.markdown(chips_html, unsafe_allow_html=True)

    # Editor principal
    st.markdown('<p class="panel-label">⬡ PROYECCIÓN DEL INFORME</p>', unsafe_allow_html=True)
    reporte_editado = st.text_area(
        "Informe",
        value=st.session_state.reporte_final,
        height=540,
        label_visibility="collapsed",
        placeholder="El informe generado aparecerá aquí para revisión y edición...\n\nPuedes:\n• Dictar en el panel izquierdo → Procesar Informe\n• Editar manualmente cualquier sección\n• Usar Optimizar Conclusión para refinar la impresión diagnóstica\n• Aplicar clasificaciones del panel derecho"
    )

    # Barra de completitud
    if reporte_editado.strip():
        score = calcular_completitud(reporte_editado)
        st.markdown(f"""
        <div style="display:flex;align-items:center;gap:10px;margin-top:6px;">
            <span style="font-size:10px;color:#2a4a6a;font-family:'IBM Plex Mono',monospace;">Completitud</span>
            <div class="completitud-bar-bg" style="flex:1;">
                <div class="completitud-bar-fill" style="width:{score}%"></div>
            </div>
            <span style="font-size:10px;color:#3a7abf;font-family:'IBM Plex Mono',monospace;">{score}%</span>
        </div>
        """, unsafe_allow_html=True)

    # Acciones
    st.markdown("<br>", unsafe_allow_html=True)
    if st.session_state.reporte_final:
        c1, c2, c3 = st.columns([1.4, 1.4, 1])
        with c1:
            if st.button("⟡ Optimizar conclusión", use_container_width=True):
                if api_key and reporte_editado.strip():
                    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                    with st.spinner("Refinando impresión diagnóstica..."):
                        try:
                            prompt_ref = f"""
Eres el optimizador diagnóstico de Beam AI.

TAREA: Lee el informe y MEJORA ÚNICAMENTE el bloque IMPRESIÓN DIAGNÓSTICA.
- Hazla más elegante, concluyente y médicamente precisa.
- Incluye clasificaciones radiológicas pertinentes (Stoller, ICRS, K-L, Pfirrmann, etc.)
- Agrega correlación clínica si es pertinente.
- Sugiere seguimiento o manejo si el grado de lesión lo amerita.

REGLAS ESTRICTAS:
- Devuelve el informe COMPLETO. Conserva Técnica y Hallazgos intactos.
- Solo sustituye el bloque de Impresión.
- CERO asteriscos. Títulos en MAYÚSCULAS.

REPORTE:
{reporte_editado}
"""
                            res_ref = client.chat.completions.create(
                                model="deepseek-chat",
                                messages=[{"role": "user", "content": prompt_ref}],
                                temperature=0.3
                            )
                            st.session_state.reporte_final = res_ref.choices[0].message.content
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error: {e}")

        with c2:
            if st.button("✦ Sugerir definiciones operativas", use_container_width=True):
                if api_key and reporte_editado.strip():
                    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                    with st.spinner("Buscando definiciones y clasificaciones..."):
                        try:
                            prompt_def = f"""
Eres un radiólogo experto. Lee el siguiente informe e identifica:

1. CLASIFICACIONES RADIOLÓGICAS usadas (correctas o incorrectas) y su significado clínico.
2. DEFINICIONES OPERATIVAS de los hallazgos mencionados.
3. SUGERENCIAS para añadir clasificaciones faltantes o más específicas.
4. CORRELACIÓN CLÍNICA breve si aplica.

Responde en formato estructurado, claro y conciso. SIN asteriscos.

INFORME:
{reporte_editado}
"""
                            res_def = client.chat.completions.create(
                                model="deepseek-chat",
                                messages=[{"role": "user", "content": prompt_def}],
                                temperature=0.2
                            )
                            st.info(res_def.choices[0].message.content)
                        except Exception as e:
                            st.error(f"Error: {e}")

        with c3:
            st.download_button("↓ Exportar .docx", generar_docx(reporte_editado),
                               "BeamAI_Informe.docx", use_container_width=True)

# ─────────────────────────────────────────
# COLUMNA DERECHA — Clasificaciones
# ─────────────────────────────────────────
with col_der:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<p class="panel-label">CLASIFICACIONES RADIOLÓGICAS</p>', unsafe_allow_html=True)
    st.markdown('<p style="font-size:10px;color:#2a4060;font-family:\'IBM Plex Mono\',monospace;margin-bottom:8px;">Haz clic en un grado para marcarlo como activo en el informe</p>', unsafe_allow_html=True)

    for nombre_clasif, items in CLASIFICACIONES.items():
        with st.expander(nombre_clasif, expanded=False):
            for grado, descripcion in items:
                key = f"{nombre_clasif}_{grado}"
                activo = st.session_state.clasif_activas.get(nombre_clasif) == f"Grado {grado}: {descripcion}"
                clase_css = "clasif-item active" if activo else "clasif-item"
                st.markdown(
                    f'<div class="{clase_css}"><span class="clasif-grade">{grado}</span> · {descripcion}</div>',
                    unsafe_allow_html=True
                )
                if st.button(f"▸ {grado}", key=key, use_container_width=True):
                    if activo:
                        del st.session_state.clasif_activas[nombre_clasif]
                    else:
                        st.session_state.clasif_activas[nombre_clasif] = f"Grado {grado}: {descripcion}"
                    st.rerun()

    if st.session_state.clasif_activas:
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<p class="panel-label">ACTIVAS EN ESTE INFORME</p>', unsafe_allow_html=True)
        for k, v in st.session_state.clasif_activas.items():
            st.markdown(
                f'<div class="clasif-item active" style="margin-bottom:4px;"><span style="color:#2a5a8a;font-size:10px;">{k[:20]}...</span><br>{v}</div>',
                unsafe_allow_html=True
            )
        if st.button("✕ Limpiar clasificaciones", use_container_width=True):
            st.session_state.clasif_activas = {}
            st.rerun()

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown(
        '<p style="font-size:9px;color:#1a3050;font-family:\'IBM Plex Mono\',monospace;text-align:center;line-height:1.6;">'
        'Beam AI · MSK / Neuro / TX v2.1<br>'
        'ACR · RSNA · ISAKOS · ICRS<br>'
        'Sistemas de clasificación validados</p>',
        unsafe_allow_html=True
    )
