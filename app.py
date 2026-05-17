import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import speech_recognition as sr
import io
import json
from openai import OpenAI

# ==========================================
# CONFIG
# ==========================================
st.set_page_config(
    page_title="Beam AI | PACS Editor v3",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# SUGERENCIAS POR MODALIDAD
# ==========================================
SUGERENCIAS = {
    "Resonancia Magnética": [
        "Stoller III → desgarro confirmado",
        "Extrusión meniscal >3 mm",
        "ICRS III → >50% grosor condral",
        "Kellgren-Lawrence III gonartrosis",
        "Edema óseo subcondral activo",
        "Lesión LCA Hope & Feagin parcial",
        "Pfirrmann IV discopatía",
        "Contusión ósea por impacto",
    ],
    "Tomografía Computarizada": [
        "Hounsfield: hueso ~700 UH",
        "Adenopatía >1 cm eje corto",
        "Nódulo Fleischner >6 mm sólido",
        "WELLS alta probabilidad TEP",
        "Murphy score apendicitis",
        "ASPECTS ACV isquémico",
    ],
    "Radiografía": [
        "Kellgren-Lawrence I-IV artrosis",
        "Cobb >10° escoliosis",
        "Índice cardiotorácico >0.5",
        "Radiopacidad lobar consolidación",
        "Línea pleural neumotórax",
    ],
    "Ultrasonido": [
        "TIRADS 4 → considerar BAAF",
        "BI-RADS 4B → biopsia indicada",
        "Resistividad >0.7 sospecha maligna",
        "Murphy positivo colecistitis",
        "McBurney dolor apendicitis",
    ],
    "PET-CT": [
        "SUVmax >2.5 actividad metabólica",
        "LI-RADS 5 → HCC definitivo",
        "Captación focal vs difusa",
        "Respuesta PERCIST criterios",
    ],
}

CLASIFICACIONES = {
    "Menisco · Stoller": [
        ("I", "Señal focal intrameniscal"),
        ("II", "Señal lineal, no articular"),
        ("III", "Alcanza superficie → desgarro"),
    ],
    "Cartílago · ICRS": [
        ("I", "Fibrilación superficial"),
        ("II", "<50% grosor"),
        ("III", ">50% grosor"),
        ("IV", "Hueso subcondral expuesto"),
    ],
    "Artrosis · Kellgren-Lawrence": [
        ("I", "Posible osteofito"),
        ("II", "Osteofito definido"),
        ("III", "Pinzamiento moderado"),
        ("IV", "Pinzamiento grave"),
    ],
    "LCA · Hope & Feagin": [
        ("Parcial", "Fibras continuas, señal ↑"),
        ("Completa", "Discontinuidad total"),
        ("Crónica", "Fibras atróficas"),
    ],
    "Columna · Pfirrmann": [
        ("I", "Núcleo brillante homogéneo"),
        ("II", "Señal alta, zona no clara"),
        ("III", "Señal gris, distinción borrosa"),
        ("IV", "Señal baja, sin distinción"),
        ("V", "Sin espacio discal"),
    ],
    "TIRADS · ACR": [
        ("2", "No sospechoso"),
        ("3", "Levemente sospechoso"),
        ("4", "Moderadamente sospechoso"),
        ("5", "Altamente sospechoso"),
    ],
    "BI-RADS · ACR": [
        ("2", "Benigno"),
        ("3", "Probablemente benigno"),
        ("4A/4B/4C", "Sospechoso — biopsia"),
        ("5", "Altamente maligno"),
    ],
}

REGIONES = [
    "Rodilla", "Columna lumbar", "Columna cervical", "Hombro",
    "Cadera", "Tobillo / Pie", "Muñeca / Mano", "Codo",
    "Cerebro", "Columna dorsal", "Tórax", "Abdomen / Pelvis",
    "Mama", "Tiroides", "Hígado",
]

# ==========================================
# ESTADO
# ==========================================
for k, v in {
    "dictado": "",
    "reporte_html": "",
    "reporte_texto": "",
    "clasif_activas": {},
    "definiciones_resultado": "",
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ==========================================
# HELPERS
# ==========================================
def leer_plantilla(file):
    doc = Document(file)
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

def generar_docx(html_texto):
    """Convierte HTML básico a DOCX con estilos."""
    from html.parser import HTMLParser

    class HTMLtoDocx(HTMLParser):
        def __init__(self):
            super().__init__()
            self.doc = Document()
            style = self.doc.styles["Normal"]
            style.font.name = "Arial"
            style.font.size = Pt(11)
            self.bold = False
            self.italic = False
            self.underline = False
            self.current_para = None
            self.in_li = False

        def handle_starttag(self, tag, attrs):
            if tag in ("b", "strong"):
                self.bold = True
            elif tag in ("i", "em"):
                self.italic = True
            elif tag == "u":
                self.underline = True
            elif tag == "br":
                if self.current_para is None:
                    self.current_para = self.doc.add_paragraph()
            elif tag in ("p", "div"):
                self.current_para = self.doc.add_paragraph()
            elif tag == "li":
                self.current_para = self.doc.add_paragraph(style="List Bullet")
                self.in_li = True
            elif tag == "hr":
                p = self.doc.add_paragraph()
                p.paragraph_format.space_after = Pt(6)

        def handle_endtag(self, tag):
            if tag in ("b", "strong"):
                self.bold = False
            elif tag in ("i", "em"):
                self.italic = False
            elif tag == "u":
                self.underline = False
            elif tag == "li":
                self.in_li = False
                self.current_para = None

        def handle_data(self, data):
            text = data.strip()
            if not text:
                return
            if self.current_para is None:
                self.current_para = self.doc.add_paragraph()
            run = self.current_para.add_run(text + " ")
            run.bold = self.bold
            run.italic = self.italic
            run.underline = self.underline

    import re
    # Limpiar y normalizar HTML básico
    clean = html_texto.replace("\n", " ").strip()
    parser = HTMLtoDocx()
    try:
        parser.feed(clean)
    except Exception:
        # Fallback: texto plano
        doc = Document()
        import re as _re
        plain = _re.sub(r"<[^>]+>", "", html_texto)
        for line in plain.split("\n"):
            doc.add_paragraph(line)
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()

    bio = io.BytesIO()
    parser.doc.save(bio)
    return bio.getvalue()

def transcribir_voz(audio_file):
    r = sr.Recognizer()
    with sr.AudioFile(audio_file) as source:
        try:
            return r.recognize_google(r.record(source), language="es-MX")
        except Exception:
            return ""

def calcular_completitud(texto):
    secciones = ["TÉCNICA", "HALLAZGOS", "IMPRESIÓN"]
    encontradas = sum(1 for s in secciones if s in texto.upper())
    palabras = len(texto.split())
    return min(100, int((encontradas / 3) * 60 + min(palabras / 150, 1) * 40))

# ==========================================
# API KEY
# ==========================================
try:
    api_key = st.secrets["deepseek_key"]
except Exception:
    api_key = ""

# ==========================================
# CSS GLOBAL
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=Syne:wght@700&display=swap');

html, .stApp { background: #070a10 !important; }
.block-container { padding: 0 !important; max-width: 100% !important; }
header, footer { display: none !important; }
[data-testid="stToolbar"] { display: none !important; }

/* TOPBAR */
.beam-topbar {
    background: #0b0f1a;
    border-bottom: 1px solid #1a2333;
    padding: 10px 24px;
    display: flex; align-items: center; gap: 16px;
    position: sticky; top: 0; z-index: 9999;
}
.beam-logo {
    font-family: 'Syne', sans-serif; font-weight: 700; font-size: 15px;
    color: #e2edf8; letter-spacing: .1em; display: flex; align-items: center; gap: 8px;
}
.ldot { width: 8px; height: 8px; border-radius: 50%; background: #3b8bd4; display: inline-block; }
.tbadge {
    font-size: 11px; color: #4a9ed4; background: #0d1e30;
    border: 1px solid #1a3a58; border-radius: 4px; padding: 2px 10px;
    font-family: 'IBM Plex Mono', monospace;
}
.tstat {
    margin-left: auto; font-size: 11px; color: #2a5a3a;
    font-family: 'IBM Plex Mono', monospace; display: flex; align-items: center; gap: 6px;
}
.sdot { width: 6px; height: 6px; border-radius: 50%; background: #2ecc71; display: inline-block; }

/* PANEL LABELS */
.plabel {
    font-size: 9px !important; letter-spacing: .2em !important; color: #2a4a6a !important;
    text-transform: uppercase !important; font-family: 'IBM Plex Mono', monospace !important;
    margin-bottom: 3px !important; margin-top: 0 !important; display: block;
}

/* INPUTS */
[data-testid="stSelectbox"] > div > div {
    background: #0f1520 !important; border: 1px solid #1a2f48 !important;
    border-radius: 6px !important; color: #7ab0cc !important;
    font-family: 'IBM Plex Mono', monospace !important; font-size: 12px !important;
}
[data-testid="stSelectbox"] > div > div:hover { border-color: #2a5a8a !important; }

.stTextArea textarea {
    background: #0f1520 !important; border: 1px solid #1a2f48 !important;
    border-radius: 7px !important; color: #b8d0e8 !important;
    font-family: 'IBM Plex Mono', monospace !important; font-size: 11px !important; line-height: 1.6 !important;
}
.stTextArea textarea:focus { border-color: #2a5580 !important; box-shadow: none !important; }

[data-testid="stAudioInput"] {
    background: #0f1a28 !important; border: 1px solid #1a3a58 !important; border-radius: 8px !important;
}

[data-testid="stFileUploader"] {
    background: #0f1520 !important; border: 1px dashed #1d3550 !important; border-radius: 8px !important;
}
[data-testid="stFileUploader"] * {
    color: #3d6090 !important; font-family: 'IBM Plex Mono', monospace !important; font-size: 11px !important;
}
[data-testid="stFileUploader"]:hover { border-color: #2a5a8a !important; }

/* BOTÓN PRINCIPAL */
.btn-main > div > button {
    background: #1a3a60 !important; border: 1px solid #2a5a90 !important;
    color: #7ac2f0 !important; font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important; font-size: 13px !important; letter-spacing: .06em !important;
    border-radius: 8px !important; padding: .75rem 1rem !important; width: 100% !important;
    transition: all .2s !important;
}
.btn-main > div > button:hover { background: #1f4878 !important; border-color: #3a78c0 !important; }

/* BOTONES SECUNDARIOS */
.stButton > button {
    background: #0d1828 !important; border: 1px solid #1a3050 !important;
    color: #3a6a9f !important; font-family: 'IBM Plex Mono', monospace !important;
    font-size: 11px !important; border-radius: 6px !important; transition: all .15s !important;
}
.stButton > button:hover { background: #102030 !important; border-color: #2a5a8a !important; color: #7ac2f0 !important; }

/* EXPANDER */
[data-testid="stExpander"] {
    background: #0b0f1a !important; border: 1px solid #1a2840 !important;
    border-radius: 7px !important; margin-bottom: 3px !important;
}
[data-testid="stExpander"] summary {
    color: #3a6a9f !important; font-family: 'IBM Plex Mono', monospace !important;
    font-size: 11px !important; letter-spacing: .06em !important;
}
[data-testid="stExpander"] summary:hover { color: #7ac2f0 !important; }
[data-testid="stExpander"] summary svg { color: #2a4a6a !important; }

/* DOWNLOAD BUTTON */
[data-testid="stDownloadButton"] > button {
    background: #0d1a28 !important; border: 1px solid #1a3a58 !important;
    color: #3a8abf !important; font-family: 'IBM Plex Mono', monospace !important;
    font-size: 11px !important; border-radius: 6px !important;
}
[data-testid="stDownloadButton"] > button:hover { background: #102030 !important; color: #7ac2f0 !important; }

/* CHIPS */
.sug-chips { display: flex; flex-wrap: wrap; gap: 5px; margin-bottom: 6px; }
.sug-chip {
    font-size: 10px; font-family: 'IBM Plex Mono', monospace;
    color: #3a7aaf; background: #0a1828; border: 1px solid #152a40;
    padding: 3px 9px; border-radius: 4px; display: inline-block;
}

/* BARRA DE COMPLETITUD */
.cbar-wrap { display: flex; align-items: center; gap: 8px; margin-top: 6px; }
.cbar-bg { flex: 1; height: 3px; background: #0f1828; border-radius: 2px; overflow: hidden; }
.cbar-fill { height: 100%; background: #1a4a80; border-radius: 2px; transition: width .4s; }
.cpct { font-size: 10px; color: #2a5a9a; font-family: 'IBM Plex Mono', monospace; }

/* SCROLL */
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: #070a10; }
::-webkit-scrollbar-thumb { background: #1a2f48; border-radius: 2px; }

hr { border-color: #111a28 !important; margin: 8px 0 !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# TOPBAR
# ==========================================
st.markdown("""
<div class="beam-topbar">
    <div class="beam-logo"><span class="ldot"></span> BEAM AI</div>
    <span class="tbadge">v3.0 · PACS Editor · MSK / Neuro / TX</span>
    <div class="tstat"><span class="sdot"></span> DeepSeek · activo</div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# LAYOUT
# ==========================================
col_izq, col_centro = st.columns([1, 2.6], gap="small")

# ─────────────────────────────────────────
# IZQUIERDA
# ─────────────────────────────────────────
with col_izq:
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    if not api_key:
        api_key = st.text_input("API Key", type="password",
                                label_visibility="collapsed",
                                placeholder="sk-... DeepSeek API Key")

    # ── Modalidad y región ──
    with st.expander("⊞  MODALIDAD & REGIÓN", expanded=True):
        st.markdown('<span class="plabel">MODALIDAD</span>', unsafe_allow_html=True)
        modalidad = st.selectbox("Modalidad", list(SUGERENCIAS.keys()), label_visibility="collapsed")

        st.markdown('<span class="plabel">REGIÓN ANATÓMICA</span>', unsafe_allow_html=True)
        region = st.selectbox("Región", REGIONES, label_visibility="collapsed")

    # ── Dictado ──
    with st.expander("⊞  DICTADO DE VOZ", expanded=True):
        audio_data = st.audio_input("Voz", label_visibility="collapsed")
        if audio_data:
            nuevo = transcribir_voz(audio_data)
            if nuevo and nuevo not in st.session_state.dictado:
                st.session_state.dictado += " " + nuevo

        st.markdown('<span class="plabel">SEÑAL TRANSCRITA</span>', unsafe_allow_html=True)
        dictado = st.text_area("Dictado", value=st.session_state.dictado,
                               height=120, label_visibility="collapsed",
                               placeholder="Dictado o escritura manual...\n\nEj: Desgarro horizontal menisco medial grado III Stoller, extrusión 3 mm...")

    # ── Sugerencias ──
    with st.expander("⊞  SUGERENCIAS IA", expanded=True):
        sugs = SUGERENCIAS.get(modalidad, [])
        chips_html = '<div class="sug-chips">'
        for s in sugs:
            chips_html += f'<span class="sug-chip">{s}</span>'
        chips_html += '</div>'
        st.markdown(chips_html, unsafe_allow_html=True)
        st.caption("Haz clic en un chip en el editor para insertar")

    # ── Clasificaciones ──
    with st.expander("⊞  CLASIFICACIONES", expanded=False):
        for nombre, items in CLASIFICACIONES.items():
            st.markdown(f'<span class="plabel">{nombre}</span>', unsafe_allow_html=True)
            for grado, desc in items:
                key = f"cls_{nombre}_{grado}"
                activo = st.session_state.clasif_activas.get(nombre) == f"Grado {grado}: {desc}"
                label = f"{'✓ ' if activo else ''}{grado} · {desc}"
                if st.button(label, key=key, use_container_width=True):
                    if activo:
                        del st.session_state.clasif_activas[nombre]
                    else:
                        st.session_state.clasif_activas[nombre] = f"Grado {grado}: {desc}"
                    st.rerun()
            st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    # ── Configuración ──
    with st.expander("⊞  CONFIGURACIÓN", expanded=False):
        st.markdown('<span class="plabel">PLANTILLA BASE</span>', unsafe_allow_html=True)
        archivo_base = st.file_uploader("Plantilla", type=["docx"], label_visibility="collapsed")
        plantilla_txt = leer_plantilla(archivo_base) if archivo_base else ""

        st.markdown('<span class="plabel">DIRECTRICES DE ESTILO</span>', unsafe_allow_html=True)
        instrucciones = st.text_area("Directrices", height=70, label_visibility="collapsed",
                                     value="Lenguaje médico experto. Sin asteriscos. Expandir clasificaciones radiológicas.")

    # ── Botones ──
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    st.markdown('<div class="btn-main">', unsafe_allow_html=True)
    procesar = st.button("⬡  PROCESAR INFORME", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("⌫ Purgar dictado", use_container_width=True):
            st.session_state.dictado = ""
            st.rerun()
    with c2:
        if st.button("⟳ Limpiar informe", use_container_width=True):
            st.session_state.reporte_html = ""
            st.session_state.reporte_texto = ""
            st.rerun()

# ─────────────────────────────────────────
# PROCESAMIENTO IA
# ─────────────────────────────────────────
if procesar:
    if api_key and dictado.strip():
        client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

        clasif_ctx = ""
        if st.session_state.clasif_activas:
            clasif_ctx = "\nCLASIFICACIONES ACTIVAS DEL RADIÓLOGO:\n"
            for k, v in st.session_state.clasif_activas.items():
                clasif_ctx += f"- {k}: {v}\n"

        prompt = f"""
Eres Beam AI, asistente experto en interpretación radiológica. Redacta un informe de {modalidad} para región: {region}.

LÓGICA:
1. Agrupa e infiere diagnósticos desde los hallazgos del dictado.
2. Si recibes diagnósticos, expande los hallazgos anatómicos esperados.
3. Usa las clasificaciones radiológicas activas del radiólogo.
4. La IMPRESIÓN DIAGNÓSTICA debe ser elegante, concluyente y clínicamente accionable.
5. Sugiere seguimiento/manejo si el grado de lesión lo amerita.

{clasif_ctx}

FORMATO ESTRICTO:
- TÍTULOS DE SECCIÓN en MAYÚSCULAS (TÉCNICA, HALLAZGOS, IMPRESIÓN DIAGNÓSTICA)
- SIN asteriscos ni markdown
- Subtítulos de subsección en MAYÚSCULAS
- Usa viñetas con "•" para la impresión diagnóstica
- Redacción fluida y profesional
- Respeta esta plantilla base: {plantilla_txt if plantilla_txt else "TÉCNICA / HALLAZGOS / IMPRESIÓN DIAGNÓSTICA"}

DIRECTRICES: {instrucciones}

DICTADO DEL RADIÓLOGO:
{dictado}
"""
        with st.spinner("Sintetizando modelo de datos..."):
            try:
                res = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "system", "content": prompt}],
                    temperature=0.1
                )
                texto = res.choices[0].message.content
                st.session_state.reporte_texto = texto

                # Convertir texto plano a HTML básico para el editor
                html_lines = []
                for line in texto.split("\n"):
                    stripped = line.strip()
                    if not stripped:
                        html_lines.append("<br>")
                    elif stripped.isupper() and len(stripped) < 60:
                        html_lines.append(f"<b>{stripped}</b><br>")
                    elif stripped.startswith("•"):
                        html_lines.append(f"<li>{stripped[1:].strip()}</li>")
                    else:
                        html_lines.append(f"{stripped}<br>")

                st.session_state.reporte_html = "\n".join(html_lines)
                st.rerun()
            except Exception as e:
                st.error(f"Error de enlace: {e}")
    elif not api_key:
        st.warning("Ingresa tu API Key de DeepSeek.")
    else:
        st.warning("Ingresa dictado o descripción clínica.")

# ─────────────────────────────────────────
# CENTRO — Editor rico
# ─────────────────────────────────────────
with col_centro:
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # Sugerencias chips (display)
    sugs = SUGERENCIAS.get(modalidad, [])
    chips_html = '<div class="sug-chips" style="margin-bottom:8px">'
    chips_html += '<span style="font-size:10px;color:#2a4060;font-family:\'IBM Plex Mono\',monospace;margin-right:4px;">Sugerencias →</span>'
    for s in sugs:
        chips_html += f'<span class="sug-chip">{s}</span>'
    chips_html += '</div>'
    st.markdown(chips_html, unsafe_allow_html=True)

    # Contenido inicial del editor
    contenido_inicial = st.session_state.reporte_html or """<b>RESONANCIA MAGNÉTICA DE RODILLA DERECHA</b><br>
<br>
<b>TÉCNICA</b><br>
Secuencias multiplanares en T1, DP con supresión grasa (DPFS), T2 y STIR en planos axial, coronal y sagital, sin contraste.<br>
<br>
<b>HALLAZGOS</b><br>
<br>
<b>MENISCOS</b><br>
Menisco medial: alteración de señal grado III de Stoller en cuerpo y cuerno posterior, compatible con desgarro horizontal. Extrusión de 3 mm.<br>
Menisco lateral: morfología e intensidad de señal conservadas.<br>
<br>
<b>LIGAMENTOS</b><br>
Ligamento cruzado anterior con señal heterogénea en tercio proximal, compatible con lesión parcial grado I de Hope &amp; Feagin.<br>
<br>
<b>CARTÍLAGO</b><br>
Adelgazamiento condral focal grado III de ICRS en platillo tibial medial (12 mm).<br>
<br>
<b>IMPRESIÓN DIAGNÓSTICA</b><br>
<li>Desgarro horizontal de menisco medial, Stoller grado III, con extrusión de 3 mm.</li>
<li>Lesión parcial grado I de LCA según Hope &amp; Feagin.</li>
<li>Condropatía grado III ICRS en compartimento medial con subcondral reactivo.</li>"""

    # Fondos disponibles
    fondos = {
        "clinical": {"bg": "#0a1018", "color": "#d0e4f0", "label": "Clínico"},
        "white": {"bg": "#f8f9fa", "color": "#1a2a3a", "label": "Blanco"},
        "warm": {"bg": "#141008", "color": "#ddd0b8", "label": "Cálido"},
        "contrast": {"bg": "#000409", "color": "#e8f4ff", "label": "Contraste"},
        "slate": {"bg": "#0d1520", "color": "#c8dff0", "label": "Slate"},
    }

    # Editor como componente HTML
    editor_html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=Syne:wght@700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:#08090f;font-family:Arial,sans-serif;overflow-x:hidden}}

.format-bar{{
    background:#0b0f1a;border-bottom:1px solid #1a2333;
    padding:6px 12px;display:flex;align-items:center;gap:5px;flex-wrap:wrap;
    position:sticky;top:0;z-index:100;
}}
.fb-group{{display:flex;align-items:center;gap:2px;padding-right:7px;border-right:1px solid #1a2333}}
.fb-group:last-child{{border-right:none}}
.fb-btn{{
    background:none;border:1px solid transparent;color:#3a6a9f;
    font-size:12px;font-family:Arial,sans-serif;
    padding:4px 7px;border-radius:4px;cursor:pointer;transition:all .15s;
    min-width:26px;text-align:center;line-height:1;
}}
.fb-btn:hover{{background:#0d1828;border-color:#1a3050;color:#7ac2f0}}
.fb-btn.active{{background:#0f1e30;border-color:#1a3a58;color:#7ac2f0}}
.fb-sel{{
    background:#0f1520;border:1px solid #1a2f48;color:#6a9abf;
    font-size:11px;font-family:'IBM Plex Mono',monospace;
    padding:3px 5px;border-radius:4px;outline:none;appearance:none;cursor:pointer;
}}
.cdot{{
    width:16px;height:16px;border-radius:50%;cursor:pointer;
    border:2px solid transparent;transition:border-color .15s;flex-shrink:0;
}}
.cdot:hover,.cdot.sel{{border-color:#3b8bd4}}
.flabel{{font-size:9px;color:#2a4060;font-family:'IBM Plex Mono',monospace;}}

.editor-wrap{{padding:16px 20px;min-height:500px}}
.doc-surface{{
    min-height:480px;padding:24px 32px;
    font-family:Arial,sans-serif;font-size:14px;line-height:1.85;
    color:#d0e4f0;outline:none;
    border-radius:6px;transition:background .3s,color .3s;
    background:#0a1018;
}}
.doc-surface li{{margin-left:20px;margin-bottom:3px}}
.doc-surface b{{color:#e8f4ff}}
.doc-surface hr{{border:none;border-top:1px solid #2a4a6a;margin:10px 0}}

.action-strip{{
    background:#0a0d16;border-top:1px solid #1a2333;
    padding:8px 14px;display:flex;align-items:center;gap:7px;flex-wrap:wrap;
}}
.act-btn{{
    background:#0d1828;border:1px solid #1a3050;color:#3a6a9f;
    font-size:11px;font-family:'IBM Plex Mono',monospace;
    padding:5px 11px;border-radius:5px;cursor:pointer;
    transition:all .15s;display:flex;align-items:center;gap:5px;
}}
.act-btn:hover{{color:#7ac2f0;background:#102030;border-color:#2a4a6a}}
.act-btn.prime{{color:#6ab8f0;border-color:#1a3a58;background:#0e1e30}}
.cbar-wrap{{margin-left:auto;display:flex;align-items:center;gap:7px}}
.cbar-bg{{width:70px;height:3px;background:#0f1828;border-radius:2px;overflow:hidden}}
.cbar-fill{{height:100%;background:#1a4a80;border-radius:2px;transition:width .4s}}
.cpct{{font-size:10px;color:#2a5a9a;font-family:'IBM Plex Mono',monospace}}
</style>
</head>
<body>

<div class="format-bar">
  <div class="fb-group">
    <select class="fb-sel" id="fontSel" onchange="applyFont(this.value)" style="width:90px">
      <option value="Arial" selected>Arial</option>
      <option value="'Courier New'">Courier New</option>
      <option value="Georgia">Georgia</option>
      <option value="'Times New Roman'">Times New Roman</option>
      <option value="'IBM Plex Mono'">IBM Plex Mono</option>
    </select>
    <select class="fb-sel" id="sizeSel" onchange="applySize(this.value)" style="width:44px">
      <option>11</option><option>12</option><option>13</option>
      <option selected>14</option><option>15</option><option>16</option><option>18</option>
    </select>
  </div>
  <div class="fb-group">
    <button class="fb-btn" id="btnB" onclick="fmt('bold')" title="Negrita"><b>B</b></button>
    <button class="fb-btn" id="btnI" onclick="fmt('italic')" title="Cursiva"><i>I</i></button>
    <button class="fb-btn" id="btnU" onclick="fmt('underline')" title="Subrayado"><u>U</u></button>
  </div>
  <div class="fb-group">
    <button class="fb-btn" onclick="fmt('justifyLeft')" title="Izquierda"><i class="ti ti-align-left"></i></button>
    <button class="fb-btn" onclick="fmt('justifyCenter')" title="Centro"><i class="ti ti-align-center"></i></button>
    <button class="fb-btn" onclick="fmt('justifyRight')" title="Derecha"><i class="ti ti-align-right"></i></button>
  </div>
  <div class="fb-group">
    <button class="fb-btn" onclick="fmt('insertUnorderedList')" title="Viñetas"><i class="ti ti-list"></i></button>
    <button class="fb-btn" onclick="fmt('insertOrderedList')" title="Lista numerada"><i class="ti ti-list-numbers"></i></button>
    <button class="fb-btn" onclick="insertHR()" title="Separador"><i class="ti ti-minus"></i></button>
  </div>
  <div class="fb-group" style="align-items:center;gap:5px">
    <span class="flabel">Fondo:</span>
    <div class="cdot sel" style="background:#0a1018" data-bg="clinical" onclick="setBg(this,'#0a1018','#d0e4f0')" title="Clínico"></div>
    <div class="cdot" style="background:#f8f9fa;border:1px solid #bbb" data-bg="white" onclick="setBg(this,'#f8f9fa','#1a2a3a')" title="Blanco"></div>
    <div class="cdot" style="background:#141008" data-bg="warm" onclick="setBg(this,'#141008','#ddd0b8')" title="Cálido"></div>
    <div class="cdot" style="background:#000409" data-bg="contrast" onclick="setBg(this,'#000409','#e8f4ff')" title="Contraste"></div>
    <div class="cdot" style="background:#0d1520" data-bg="slate" onclick="setBg(this,'#0d1520','#c8dff0')" title="Slate"></div>
  </div>
</div>

<div class="editor-wrap">
  <div class="doc-surface" id="doc" contenteditable="true" spellcheck="false">
    {contenido_inicial}
  </div>
</div>

<div class="action-strip">
  <button class="act-btn" onclick="optimize()"><i class="ti ti-wand" style="font-size:13px"></i> Optimizar conclusión</button>
  <button class="act-btn" onclick="getDefiniciones()"><i class="ti ti-book" style="font-size:13px"></i> Definiciones operativas</button>
  <button class="act-btn" onclick="copyText()"><i class="ti ti-copy" style="font-size:13px"></i> Copiar</button>
  <button class="act-btn prime" onclick="exportDoc()"><i class="ti ti-download" style="font-size:13px"></i> Exportar .docx</button>
  <div class="cbar-wrap">
    <div class="cbar-bg"><div class="cbar-fill" id="cbar" style="width:0%"></div></div>
    <span class="cpct" id="cpct">0%</span>
  </div>
</div>

<script>
var docEl = document.getElementById('doc');
var currentFont = 'Arial';

function fmt(cmd) {{
  docEl.focus();
  document.execCommand(cmd, false, null);
  updateState();
}}

function updateState() {{
  ['Bold','Italic','Underline'].forEach(function(c) {{
    var b = document.getElementById('btn'+c[0]);
    if(b) b.classList.toggle('active', document.queryCommandState(c.toLowerCase()));
  }});
}}

function applyFont(f) {{
  currentFont = f;
  docEl.style.fontFamily = f;
}}

function applySize(s) {{
  docEl.style.fontSize = s + 'px';
}}

function setBg(el, bg, col) {{
  docEl.style.background = bg;
  docEl.style.color = col;
  document.querySelectorAll('.cdot').forEach(function(d) {{ d.classList.remove('sel'); }});
  el.classList.add('sel');
}}

function insertHR() {{
  docEl.focus();
  document.execCommand('insertHTML', false, '<hr style="border:none;border-top:1px solid #2a4a6a;margin:12px 0"><br>');
}}

function calcScore() {{
  var t = docEl.innerText.toUpperCase();
  var secs = ['TÉCNICA','HALLAZGOS','IMPRESIÓN'];
  var found = secs.filter(function(s) {{ return t.includes(s); }}).length;
  var words = t.split(/\\s+/).filter(Boolean).length;
  return Math.min(100, Math.round((found/3)*60 + Math.min(words/150,1)*40));
}}

function updateBar() {{
  var s = calcScore();
  document.getElementById('cbar').style.width = s + '%';
  document.getElementById('cpct').textContent = s + '%';
}}

docEl.addEventListener('input', updateBar);
docEl.addEventListener('keyup', updateState);
docEl.addEventListener('mouseup', updateState);
window.addEventListener('load', function() {{ updateBar(); }});

function getContent() {{
  return docEl.innerHTML;
}}

function getText() {{
  return docEl.innerText;
}}

function copyText() {{
  var range = document.createRange();
  range.selectNode(docEl);
  window.getSelection().removeAllRanges();
  window.getSelection().addRange(range);
  document.execCommand('copy');
  window.getSelection().removeAllRanges();
}}

function optimize() {{
  var texto = getText();
  window.parent.postMessage({{type:'optimize', content: texto}}, '*');
}}

function getDefiniciones() {{
  var texto = getText();
  window.parent.postMessage({{type:'definiciones', content: texto}}, '*');
}}

function exportDoc() {{
  var texto = getText();
  window.parent.postMessage({{type:'export', content: texto, html: getContent()}}, '*');
}}
</script>
</body>
</html>
"""

    # Renderizar editor
    components.html(editor_html, height=680, scrolling=False)

    # Barra de completitud externa
    if st.session_state.reporte_texto:
        score = calcular_completitud(st.session_state.reporte_texto)
        st.markdown(f"""
        <div class="cbar-wrap">
            <span style="font-size:10px;color:#2a4a6a;font-family:'IBM Plex Mono',monospace;">Completitud del informe</span>
            <div class="cbar-bg"><div class="cbar-fill" style="width:{score}%"></div></div>
            <span class="cpct">{score}%</span>
        </div>
        """, unsafe_allow_html=True)

    # ── Acciones IA externas ──
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1.4, 1.6, 1])

    with c1:
        if st.button("⟡ Optimizar conclusión", use_container_width=True):
            if api_key and st.session_state.reporte_texto:
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                with st.spinner("Refinando impresión diagnóstica..."):
                    try:
                        prompt_ref = f"""
Eres el optimizador diagnóstico de Beam AI.

TAREA: Lee el informe y MEJORA el bloque IMPRESIÓN DIAGNÓSTICA:
- Más elegante, concluyente y médicamente precisa.
- Incluye clasificaciones radiológicas específicas con grado.
- Agrega correlación clínica y recomendación de manejo si el grado lo amerita.
- Usa viñetas con "•"

REGLAS: Devuelve el informe COMPLETO. Conserva Técnica y Hallazgos intactos.
CERO asteriscos. Títulos en MAYÚSCULAS.

REPORTE:
{st.session_state.reporte_texto}
"""
                        res = client.chat.completions.create(
                            model="deepseek-chat",
                            messages=[{"role": "user", "content": prompt_ref}],
                            temperature=0.25
                        )
                        texto = res.choices[0].message.content
                        st.session_state.reporte_texto = texto
                        lines = []
                        for line in texto.split("\n"):
                            s = line.strip()
                            if not s:
                                lines.append("<br>")
                            elif s.isupper() and len(s) < 60:
                                lines.append(f"<b>{s}</b><br>")
                            elif s.startswith("•"):
                                lines.append(f"<li>{s[1:].strip()}</li>")
                            else:
                                lines.append(f"{s}<br>")
                        st.session_state.reporte_html = "\n".join(lines)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    with c2:
        if st.button("✦ Definiciones y clasificaciones", use_container_width=True):
            if api_key and st.session_state.reporte_texto:
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                with st.spinner("Analizando clasificaciones..."):
                    try:
                        prompt_def = f"""
Eres un radiólogo experto y docente. Lee este informe radiológico y responde:

1. CLASIFICACIONES USADAS: Lista cada clasificación mencionada, su grado y qué significa clínicamente.
2. CLASIFICACIONES FALTANTES: Sugiere clasificaciones específicas que deberían agregarse basadas en los hallazgos.
3. DEFINICIONES OPERATIVAS: Define en 1-2 líneas los términos técnicos más relevantes del informe.
4. CORRELACIÓN CLÍNICA: Impacto clínico esperado y recomendación de manejo concisa.

Formato: estructurado, claro, conciso. SIN asteriscos. Títulos en MAYÚSCULAS.

INFORME:
{st.session_state.reporte_texto}
"""
                        res = client.chat.completions.create(
                            model="deepseek-chat",
                            messages=[{"role": "user", "content": prompt_def}],
                            temperature=0.2
                        )
                        st.session_state.definiciones_resultado = res.choices[0].message.content
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    with c3:
        if st.session_state.reporte_texto:
            docx_bytes = generar_docx(st.session_state.reporte_html or st.session_state.reporte_texto)
            st.download_button(
                "↓ Exportar .docx",
                data=docx_bytes,
                file_name="BeamAI_Informe.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    # Resultado de definiciones
    if st.session_state.definiciones_resultado:
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        with st.expander("⬡  DEFINICIONES OPERATIVAS & CLASIFICACIONES", expanded=True):
            st.markdown(
                f'<div style="font-family:\'IBM Plex Mono\',monospace;font-size:12px;color:#8ab0cc;line-height:1.7;background:#0a1020;padding:14px;border-radius:7px;border:1px solid #1a2840;white-space:pre-wrap">{st.session_state.definiciones_resultado}</div>',
                unsafe_allow_html=True
            )
            if st.button("✕ Cerrar análisis"):
                st.session_state.definiciones_resultado = ""
                st.rerun()
