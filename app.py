import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import tempfile
import os
from openai import OpenAI

# ──────────────────────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AURA · Radiology Intelligence",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ──────────────────────────────────────────────────────────────
# TEMAS HOLOGRÁFICOS
# ──────────────────────────────────────────────────────────────
TEMAS = {
    "Void": {
        "base": "#000308",
        "surface": "#010a14",
        "glass": "rgba(0, 180, 255, 0.03)",
        "glass_border": "rgba(0, 200, 255, 0.08)",
        "glow": "rgba(0, 180, 255, 0.15)",
        "accent": "#00c8ff",
        "accent2": "#0070a0",
        "text": "#a8d8f0",
        "text_dim": "#2a6080",
        "text_ghost": "#0e3050",
        "scan_line": "rgba(0,180,255,0.03)",
    },
    "Plasma": {
        "base": "#04000a",
        "surface": "#0c0118",
        "glass": "rgba(160, 0, 255, 0.04)",
        "glass_border": "rgba(180, 80, 255, 0.10)",
        "glow": "rgba(160, 80, 255, 0.18)",
        "accent": "#b060ff",
        "accent2": "#6020a0",
        "text": "#d0b0f8",
        "text_dim": "#5030a0",
        "text_ghost": "#200a40",
        "scan_line": "rgba(160,0,255,0.03)",
    },
    "Aurora": {
        "base": "#000a08",
        "surface": "#010f10",
        "glass": "rgba(0, 255, 180, 0.03)",
        "glass_border": "rgba(0, 220, 160, 0.09)",
        "glow": "rgba(0, 220, 160, 0.14)",
        "accent": "#00e8b0",
        "accent2": "#008060",
        "text": "#90e8d0",
        "text_dim": "#1a6050",
        "text_ghost": "#082820",
        "scan_line": "rgba(0,220,160,0.025)",
    },
    "Solar": {
        "base": "#080400",
        "surface": "#100800",
        "glass": "rgba(255, 160, 0, 0.03)",
        "glass_border": "rgba(255, 180, 40, 0.09)",
        "glow": "rgba(255, 160, 0, 0.15)",
        "accent": "#ffb030",
        "accent2": "#a06010",
        "text": "#f0d090",
        "text_dim": "#704010",
        "text_ghost": "#301800",
        "scan_line": "rgba(255,160,0,0.025)",
    },
}

REGIONES = [
    "Rodilla", "Columna lumbar", "Columna cervical", "Hombro", "Cadera",
    "Tobillo / Pie", "Muñeca / Mano", "Codo", "Cerebro", "Columna dorsal",
    "Tórax", "Abdomen / Pelvis", "Mama", "Tiroides", "Hígado",
]

MODALIDADES = [
    "Resonancia Magnética", "Tomografía Computarizada",
    "Radiografía", "Ultrasonido", "PET-CT",
]

# MEJORA: Soporte multi-modelo con etiquetas descriptivas
MODELOS = {
    "DeepSeek Chat": {"api_url": "https://api.deepseek.com", "model_id": "deepseek-chat"},
    "GPT-4o Mini": {"api_url": "https://api.openai.com/v1", "model_id": "gpt-4o-mini"},
    "GPT-4.1 Mini": {"api_url": None, "model_id": "gpt-4.1-mini"},  # usa base_url del entorno
}

REGLAS = """
PROTOCOLO CLÍNICO AURA — REGLAS INVIOLABLES:

TERMINOLOGÍA:
· Prohibido: "cambios degenerativos", "cambios crónicos" sin sustrato morfológico.
· Obligatorio: descriptores morfológicos específicos — osteofitos marginales, esclerosis subcondral,
  pinzamiento articular de X mm, condromalacia, fibrosis periarticular.

TABLAS:
· Solo generar tablas si la plantilla proporcionada contiene explícitamente [TABLA].
· Sin plantilla con tabla → cero tablas en el informe.

CLASIFICACIONES:
· Usar solo clasificaciones directamente respaldadas por los hallazgos del dictado.
· Especificar el criterio morfológico que justifica cada grado.
· No asumir grados sin evidencia suficiente.

IMPRESIÓN:
· Diagnósticos morfológicamente precisos, no genéricos.
· Lenguaje sugerente para manejo: "se sugiere correlación clínica", "puede valorarse".
"""

# ──────────────────────────────────────────────────────────────
# ESTADO
# ──────────────────────────────────────────────────────────────
DEFAULTS = {
    "tema": "Void",
    "dictado": "",
    "reporte_html": "",
    "reporte_texto": "",
    "defs_resultado": "",
    "editor_h": 560,
    "modo": "dictado",          # "dictado" | "hallazgos"
    "plantilla_txt": "",
    "tiene_tabla": False,
    "modelo_sel": "DeepSeek Chat",
    "historial": [],            # NUEVO: lista de dicts {modalidad, region, texto, html}
    "audio_procesado_id": None, # NUEVO: evita re-transcribir el mismo audio
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────

def leer_plantilla(file):
    """
    Lee un archivo .docx y extrae texto y tablas en formato Markdown.
    Retorna (texto_plantilla: str, tiene_tabla: bool).
    """
    doc = Document(file)
    partes = []
    n = 0
    try:
        import docx.text.paragraph as _pp
        import docx.table as _tt
        for el in doc.element.body:
            tag = el.tag.split('}')[-1]
            if tag == 'p':
                p = _pp.Paragraph(el, doc)
                t = p.text.strip()
                if t:
                    partes.append(t)
            elif tag == 'tbl':
                n += 1
                tbl = _tt.Table(el, doc)
                rows = [
                    "| " + " | ".join(c.text.strip() for c in r.cells) + " |"
                    for r in tbl.rows
                ]
                partes.append(f"[TABLA {n}]\n" + "\n".join(rows) + "\n[/TABLA]")
    except Exception:
        partes = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n".join(partes), n > 0


def texto_a_html(texto):
    """
    Convierte texto plano con Markdown básico a HTML para el editor.
    Soporta tablas Markdown, listas, encabezados en mayúsculas y viñetas.
    """
    lines, buf, in_tbl = [], [], False
    for line in texto.split("\n"):
        s = line.strip()
        if not s:
            if in_tbl:
                lines.append(_tbl_html(buf))
                buf = []
                in_tbl = False
            lines.append("<br>")
        elif re.match(r'^\|.+\|$', s):
            # Ignorar líneas separadoras de tabla Markdown (|---|---|)
            if all(c in '-| :' for c in s):
                continue
            in_tbl = True
            buf.append(s)
        else:
            if in_tbl:
                lines.append(_tbl_html(buf))
                buf = []
                in_tbl = False
            # Encabezados en mayúsculas (longitud < 70 y no comienzan con viñeta)
            if s.isupper() and len(s) < 70 and not s.startswith("•"):
                lines.append(f"<b>{s}</b><br>")
            elif s.startswith("•") or s.startswith("·"):
                lines.append(f"<li>{s[1:].strip()}</li>")
            # MEJORA: soporte para negritas Markdown **texto**
            elif "**" in s:
                s_fmt = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', s)
                lines.append(f"{s_fmt}<br>")
            else:
                lines.append(f"{s}<br>")
    if in_tbl and buf:
        lines.append(_tbl_html(buf))
    return "\n".join(lines)


def _tbl_html(rows):
    """Convierte filas Markdown a tabla HTML con estilos básicos."""
    h = '<table style="border-collapse:collapse;width:100%;margin:8px 0;font-size:12px">'
    for i, row in enumerate(rows):
        cols = [c.strip() for c in row.strip("|").split("|")]
        tag = "th" if i == 0 else "td"
        h += "<tr>" + "".join(
            f"<{tag} style='border:1px solid #ccc;padding:4px 9px'>{c}</{tag}>"
            for c in cols
        ) + "</tr>"
    return h + "</table>"


def generar_docx_desde_markdown(texto_md: str, modalidad: str = "", region: str = "") -> bytes:
    """
    MEJORA CRÍTICA: Genera un archivo .docx directamente desde el texto Markdown/plano
    que devuelve la IA, en lugar de parsear HTML frágil.
    Esto garantiza fidelidad de formato, tablas y tipografía.
    """
    doc = Document()

    # Estilos base
    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Calibri"
    style_normal.font.size = Pt(11)

    # Encabezado del documento
    if modalidad or region:
        titulo = doc.add_heading(f"{modalidad.upper()} — {region.upper()}", level=1)
        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    in_table_block = False
    table_rows = []

    def flush_table():
        nonlocal table_rows
        if not table_rows:
            return
        cols = max(len(r) for r in table_rows)
        t = doc.add_table(rows=len(table_rows), cols=cols)
        t.style = "Table Grid"
        for i, row in enumerate(table_rows):
            for j, cell_text in enumerate(row):
                if j < cols:
                    cell = t.rows[i].cells[j]
                    cell.text = cell_text
                    if i == 0:
                        for run in cell.paragraphs[0].runs:
                            run.bold = True
        table_rows = []

    for line in texto_md.split("\n"):
        s = line.strip()

        # Detectar bloque de tabla Markdown
        if re.match(r'^\|.+\|$', s):
            if all(c in '-| :' for c in s):
                continue  # separador de tabla
            in_table_block = True
            cols = [c.strip() for c in s.strip("|").split("|")]
            table_rows.append(cols)
            continue
        else:
            if in_table_block:
                flush_table()
                in_table_block = False

        if not s:
            doc.add_paragraph()
            continue

        # Encabezados en mayúsculas → Heading 2
        if s.isupper() and len(s) < 80 and not s.startswith("•") and not s.startswith("·"):
            h = doc.add_heading(s, level=2)
            h.runs[0].font.color.rgb = RGBColor(0x1a, 0x3a, 0x6a)
            continue

        # Viñetas
        if s.startswith("•") or s.startswith("·"):
            p = doc.add_paragraph(style="List Bullet")
            _add_formatted_run(p, s[1:].strip())
            continue

        # Línea normal con posible Markdown en línea
        p = doc.add_paragraph()
        _add_formatted_run(p, s)

    if in_table_block:
        flush_table()

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def _add_formatted_run(paragraph, text: str):
    """
    Añade texto a un párrafo de python-docx respetando negritas (**texto**)
    e itálicas (*texto*) de Markdown inline.
    """
    # Patrón para detectar **negrita** y *itálica*
    pattern = re.compile(r'(\*\*(.+?)\*\*|\*(.+?)\*)')
    last = 0
    for m in pattern.finditer(text):
        # Texto antes del marcador
        if m.start() > last:
            paragraph.add_run(text[last:m.start()])
        raw = m.group(0)
        if raw.startswith("**"):
            run = paragraph.add_run(m.group(2))
            run.bold = True
        else:
            run = paragraph.add_run(m.group(3))
            run.italic = True
        last = m.end()
    if last < len(text):
        paragraph.add_run(text[last:])


def transcribir_whisper(audio_file, client: OpenAI) -> str:
    """
    MEJORA CRÍTICA: Transcripción de audio con Whisper (OpenAI).
    Usa un prompt inicial con terminología radiológica para mejorar la precisión.
    Retorna el texto transcrito o una cadena vacía si falla.
    """
    PROMPT_RADIOLOGICO = (
        "Transcripción de dictado radiológico en español. "
        "Términos esperados: Stoller, Kellgren-Lawrence, ICRS, LCA, LCP, LCM, LCL, "
        "menisco, condromalacia, osteofito, esclerosis subcondral, extrusión, "
        "pinzamiento, STIR, DP, T1, T2, PET-CT, resonancia magnética, tomografía."
    )
    try:
        # Guardar el archivo de audio en un temporal para enviarlo a la API
        suffix = ".wav"
        if hasattr(audio_file, "name"):
            ext = os.path.splitext(audio_file.name)[-1].lower()
            if ext in [".mp3", ".mp4", ".m4a", ".ogg", ".webm", ".flac"]:
                suffix = ext

        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(audio_file.read())
            tmp_path = tmp.name

        with open(tmp_path, "rb") as f:
            result = client.audio.transcriptions.create(
                model="whisper-1",
                file=f,
                language="es",
                prompt=PROMPT_RADIOLOGICO
            )
        os.unlink(tmp_path)
        return result.text.strip()
    except Exception as e:
        # Si Whisper falla (p.ej. modelo no disponible en la URL configurada),
        # intentar con speech_recognition como fallback
        try:
            import speech_recognition as sr
            audio_file.seek(0)
            r = sr.Recognizer()
            with sr.AudioFile(audio_file) as src:
                return r.recognize_google(r.record(src), language="es-MX")
        except Exception:
            return ""


def get_openai_client(api_key: str, modelo_nombre: str) -> OpenAI:
    """
    Crea y retorna un cliente OpenAI configurado según el modelo seleccionado.
    """
    cfg = MODELOS.get(modelo_nombre, MODELOS["DeepSeek Chat"])
    if cfg["api_url"]:
        return OpenAI(api_key=api_key, base_url=cfg["api_url"])
    return OpenAI(api_key=api_key)


def get_model_id(modelo_nombre: str) -> str:
    return MODELOS.get(modelo_nombre, MODELOS["DeepSeek Chat"])["model_id"]


def completitud(texto: str) -> int:
    """Calcula un porcentaje de completitud del informe (0-100)."""
    secs = sum(1 for s in ["TÉCNICA", "HALLAZGOS", "IMPRESIÓN"] if s in texto.upper())
    words = len(texto.split())
    return min(100, int((secs / 3) * 60 + min(words / 150, 1) * 40))


def guardar_en_historial(modalidad: str, region: str, texto: str, html: str):
    """Agrega el informe actual al historial de la sesión (máximo 10 entradas)."""
    entry = {"modalidad": modalidad, "region": region, "texto": texto, "html": html}
    st.session_state.historial.insert(0, entry)
    if len(st.session_state.historial) > 10:
        st.session_state.historial = st.session_state.historial[:10]


# ──────────────────────────────────────────────────────────────
# API KEY — Prioridad: secrets → variable de entorno → input manual
# ──────────────────────────────────────────────────────────────
try:
    api_key = st.secrets["deepseek_key"]
except Exception:
    api_key = os.environ.get("OPENAI_API_KEY", "")

T = TEMAS[st.session_state.tema]

# ──────────────────────────────────────────────────────────────
# CSS HOLOGRÁFICO (v2 — limpieza de conflictos con Streamlit)
# ──────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600&family=JetBrains+Mono:wght@300;400;500&display=swap');

/* ── RESET & BASE ── */
html, body, .stApp {{ background: {T['base']} !important; }}
.block-container {{ padding: 0 !important; max-width: 100% !important; }}
header, footer, [data-testid="stToolbar"] {{ display: none !important; }}
* {{ font-family: 'JetBrains Mono', monospace !important; }}

/* ── SCAN LINE EFFECT ── */
.stApp::before {{
    content: '';
    position: fixed; inset: 0; pointer-events: none; z-index: 0;
    background: repeating-linear-gradient(
        0deg,
        {T['scan_line']} 0px,
        transparent 1px,
        transparent 3px
    );
}}

/* ── TOPBAR ── */
.aura-bar {{
    position: sticky; top: 0; z-index: 9999;
    background: {T['base']};
    border-bottom: 1px solid {T['glass_border']};
    padding: 0 24px;
    height: 44px;
    display: flex; align-items: center; gap: 0;
}}
.aura-logo {{
    font-family: 'Space Grotesk', sans-serif !important;
    font-weight: 300; font-size: 17px; letter-spacing: .35em;
    color: {T['accent']}; text-transform: uppercase;
    display: flex; align-items: center; gap: 10px;
}}
.aura-pulse {{
    width: 6px; height: 6px; border-radius: 50%;
    background: {T['accent']};
    box-shadow: 0 0 8px {T['accent']}, 0 0 20px {T['accent']};
    animation: pulse 2.4s ease-in-out infinite;
}}
@keyframes pulse {{
    0%, 100% {{ opacity: 1; transform: scale(1); }}
    50% {{ opacity: .4; transform: scale(.7); }}
}}
.aura-sep {{
    width: 1px; height: 18px;
    background: {T['glass_border']};
    margin: 0 18px;
}}
.aura-meta {{
    font-size: 9px; letter-spacing: .2em; color: {T['text_dim']};
    text-transform: uppercase;
}}
.aura-status {{
    margin-left: auto; display: flex; align-items: center; gap: 7px;
    font-size: 9px; letter-spacing: .15em; color: {T['text_dim']};
}}
.aura-online {{
    width: 5px; height: 5px; border-radius: 50%;
    background: {T['accent']};
    box-shadow: 0 0 6px {T['accent']};
}}

/* ── GLASS PANELS ── */
.glass-panel {{
    background: {T['glass']};
    border: 1px solid {T['glass_border']};
    border-radius: 2px;
    backdrop-filter: blur(8px);
    position: relative;
}}
.glass-panel::before {{
    content: '';
    position: absolute; top: 0; left: 0; right: 0;
    height: 1px;
    background: linear-gradient(90deg, transparent, {T['accent']}, transparent);
    opacity: .3;
}}

/* ── SECTION LABELS ── */
.sec-lbl {{
    font-size: 8px !important;
    letter-spacing: .25em !important;
    color: {T['text_ghost']} !important;
    text-transform: uppercase !important;
    margin-bottom: 4px !important;
    display: block;
}}

/* ── SELECTBOX ── */
[data-testid="stSelectbox"] > div > div {{
    background: transparent !important;
    border: 1px solid {T['glass_border']} !important;
    border-radius: 2px !important;
    color: {T['text']} !important;
    font-size: 11px !important;
}}
[data-testid="stSelectbox"] > div > div:hover {{
    border-color: {T['accent2']} !important;
}}

/* ── TEXTAREA ── */
.stTextArea textarea {{
    background: transparent !important;
    border: 1px solid {T['glass_border']} !important;
    border-radius: 2px !important;
    color: {T['text']} !important;
    font-size: 11px !important;
    line-height: 1.6 !important;
    caret-color: {T['accent']} !important;
}}
.stTextArea textarea:focus {{
    border-color: {T['accent2']} !important;
    box-shadow: 0 0 12px {T['glow']} !important;
}}
.stTextArea textarea::placeholder {{ color: {T['text_dim']} !important; }}

/* ── TEXT INPUT ── */
.stTextInput input {{
    background: transparent !important;
    border: 1px solid {T['glass_border']} !important;
    border-radius: 2px !important;
    color: {T['text']} !important;
    font-size: 11px !important;
    caret-color: {T['accent']} !important;
}}
.stTextInput input:focus {{
    border-color: {T['accent2']} !important;
    box-shadow: 0 0 10px {T['glow']} !important;
}}
.stTextInput input::placeholder {{ color: {T['text_dim']} !important; }}

/* ── BUTTONS ── */
.stButton > button {{
    background: transparent !important;
    border: 1px solid {T['glass_border']} !important;
    border-radius: 2px !important;
    color: {T['text_dim']} !important;
    font-size: 9px !important;
    letter-spacing: .18em !important;
    text-transform: uppercase !important;
    padding: 5px 10px !important;
    transition: all .15s !important;
}}
.stButton > button:hover {{
    border-color: {T['accent2']} !important;
    color: {T['text']} !important;
    box-shadow: 0 0 10px {T['glow']} !important;
}}

/* ── BOTÓN PRIMARIO ── */
.btn-primary .stButton > button {{
    border-color: {T['accent']} !important;
    color: {T['accent']} !important;
    box-shadow: 0 0 12px {T['glow']} !important;
    font-size: 10px !important;
    padding: 7px 14px !important;
}}
.btn-primary .stButton > button:hover {{
    background: {T['glass']} !important;
    box-shadow: 0 0 20px {T['glow']} !important;
}}

/* ── DOWNLOAD BUTTON ── */
.stDownloadButton > button {{
    background: transparent !important;
    border: 1px solid {T['accent']} !important;
    border-radius: 2px !important;
    color: {T['accent']} !important;
    font-size: 9px !important;
    letter-spacing: .18em !important;
    text-transform: uppercase !important;
    padding: 5px 10px !important;
    transition: all .15s !important;
    box-shadow: 0 0 8px {T['glow']} !important;
}}
.stDownloadButton > button:hover {{
    background: {T['glass']} !important;
    box-shadow: 0 0 18px {T['glow']} !important;
}}

/* ── EXPANDERS ── */
.streamlit-expanderHeader {{
    background: transparent !important;
    border: 1px solid {T['glass_border']} !important;
    border-radius: 2px !important;
    color: {T['text_dim']} !important;
    font-size: 9px !important;
    letter-spacing: .2em !important;
    text-transform: uppercase !important;
    padding: 5px 10px !important;
}}
.streamlit-expanderHeader:hover {{
    border-color: {T['accent2']} !important;
    color: {T['text']} !important;
}}
.streamlit-expanderContent {{
    border: 1px solid {T['glass_border']} !important;
    border-top: none !important;
    background: {T['glass']} !important;
    padding: 10px !important;
}}

/* ── FILE UPLOADER ── */
[data-testid="stFileUploader"] {{
    border: 1px dashed {T['glass_border']} !important;
    border-radius: 2px !important;
    background: {T['glass']} !important;
}}
[data-testid="stFileUploader"] label {{
    color: {T['text_dim']} !important;
    font-size: 9px !important;
    letter-spacing: .15em !important;
}}

/* ── SLIDER ── */
[data-testid="stSlider"] [data-baseweb="slider"] div[role="slider"] {{
    background: {T['accent']} !important;
    box-shadow: 0 0 6px {T['accent']} !important;
}}
[data-testid="stSlider"] [data-baseweb="slider"] div[data-testid="stThumbValue"] {{
    color: {T['text']} !important;
    font-size: 9px !important;
}}

/* ── ALERTS ── */
.stAlert {{
    background: {T['glass']} !important;
    border: 1px solid {T['glass_border']} !important;
    border-radius: 2px !important;
    color: {T['text']} !important;
    font-size: 10px !important;
}}

/* ── DEFS BOX ── */
.defs-box {{
    font-size: 11px;
    line-height: 1.75;
    color: {T['text']};
    white-space: pre-wrap;
    padding: 10px;
    background: {T['glass']};
    border: 1px solid {T['glass_border']};
    border-radius: 2px;
}}
.defs-box b {{ color: {T['text']}; }}

/* ── HISTORIAL BADGE ── */
.hist-badge {{
    display: inline-block;
    font-size: 8px; letter-spacing: .15em;
    color: {T['accent']}; border: 1px solid {T['accent2']};
    padding: 2px 7px; border-radius: 1px;
    margin-bottom: 4px; cursor: pointer;
}}

/* ── SCROLLBAR ── */
::-webkit-scrollbar {{ width: 2px; height: 2px; }}
::-webkit-scrollbar-track {{ background: transparent; }}
::-webkit-scrollbar-thumb {{ background: {T['accent2']}; border-radius: 1px; }}

/* ── MISC ── */
hr {{ border: none; border-top: 1px solid {T['glass_border']} !important; margin: 6px 0 !important; }}
[data-testid="stSpinner"] > div {{ border-top-color: {T['accent']} !important; }}
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
# TOPBAR
# ──────────────────────────────────────────────────────────────
modelo_activo = st.session_state.modelo_sel
st.markdown(f"""
<div class="aura-bar">
  <div class="aura-logo">
    <div class="aura-pulse"></div>
    AURA
  </div>
  <div class="aura-sep"></div>
  <span class="aura-meta">Radiology Intelligence · v2.0</span>
  <div class="aura-status">
    <div class="aura-online"></div>
    {modelo_activo.upper()} · ONLINE
  </div>
</div>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
# LAYOUT
# ──────────────────────────────────────────────────────────────
col_l, col_r = st.columns([1, 2.8], gap="small")

# ══════════════════════════════════════════════════════════════
# PANEL IZQUIERDO
# ══════════════════════════════════════════════════════════════
with col_l:
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # ── API KEY ──
    if not api_key:
        st.markdown('<span class="sec-lbl">API KEY</span>', unsafe_allow_html=True)
        api_key = st.text_input(
            "k", type="password", label_visibility="collapsed",
            placeholder="sk- ···  DeepSeek / OpenAI API Key"
        )
        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    # ── ESTUDIO ──
    with st.expander("▸  ESTUDIO", expanded=True):
        st.markdown('<span class="sec-lbl">MODALIDAD</span>', unsafe_allow_html=True)
        modalidad = st.selectbox("M", MODALIDADES, label_visibility="collapsed")
        st.markdown('<span class="sec-lbl">REGIÓN</span>', unsafe_allow_html=True)
        region = st.selectbox("R", REGIONES, label_visibility="collapsed")

        # NUEVO: Selector de modelo IA
        st.markdown('<span class="sec-lbl">MODELO IA</span>', unsafe_allow_html=True)
        modelo_sel = st.selectbox(
            "Modelo", list(MODELOS.keys()),
            index=list(MODELOS.keys()).index(st.session_state.modelo_sel),
            label_visibility="collapsed"
        )
        if modelo_sel != st.session_state.modelo_sel:
            st.session_state.modelo_sel = modelo_sel
            st.rerun()

    # ── MODO DE ENTRADA ──
    with st.expander("▸  MODO DE ENTRADA", expanded=True):
        modo_label = "DICTADO DE VOZ" if st.session_state.modo == "dictado" else "HALLAZGOS ESCRITOS"
        st.markdown(f'<span class="sec-lbl">{modo_label}</span>', unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⊙ VOZ", use_container_width=True):
                st.session_state.modo = "dictado"
                st.rerun()
        with c2:
            if st.button("⊙ TEXTO", use_container_width=True):
                st.session_state.modo = "hallazgos"
                st.rerun()

        if st.session_state.modo == "dictado":
            # MEJORA: Transcripción con Whisper
            audio = st.audio_input("_", label_visibility="collapsed")
            if audio:
                # MEJORA: Evitar re-transcribir el mismo archivo de audio
                audio_id = hash(audio.read())
                audio.seek(0)
                if audio_id != st.session_state.audio_procesado_id:
                    if api_key:
                        with st.spinner("Transcribiendo con Whisper..."):
                            client_tmp = get_openai_client(api_key, st.session_state.modelo_sel)
                            txt = transcribir_whisper(audio, client_tmp)
                        if txt:
                            st.session_state.dictado += (" " + txt).strip()
                            st.session_state.audio_procesado_id = audio_id
                            st.rerun()
                        else:
                            st.warning("No se pudo transcribir el audio.")
                    else:
                        st.warning("Ingresa tu API Key para usar la transcripción.")
        else:
            st.markdown('<span class="sec-lbl">HALLAZGOS / IMPRESIÓN</span>', unsafe_allow_html=True)

        dictado = st.text_area(
            "_",
            value=st.session_state.dictado,
            height=140,
            label_visibility="collapsed",
            placeholder=(
                "Dicta o escribe hallazgos, diagnósticos o ambos...\n\n"
                "Ej: Desgarro horizontal menisco medial Stoller III, extrusión 3 mm. "
                "Osteofitos marginales tibiofemorales mediales."
            ),
            key="dictado_area"
        )
        # Sincronizar el textarea con el estado de sesión sin rerun innecesario
        if dictado != st.session_state.dictado:
            st.session_state.dictado = dictado

    # ── PLANTILLA ──
    with st.expander("▸  PLANTILLA", expanded=False):
        st.markdown('<span class="sec-lbl">ARCHIVO .DOCX</span>', unsafe_allow_html=True)
        f = st.file_uploader("_", type=["docx"], label_visibility="collapsed")
        if f:
            st.session_state.plantilla_txt, st.session_state.tiene_tabla = leer_plantilla(f)
            icono = "◈  CON TABLAS" if st.session_state.tiene_tabla else "◇  CARGADA"
            st.markdown(
                f'<span style="font-size:9px;letter-spacing:.15em;color:{T["accent"]}">{icono}</span>',
                unsafe_allow_html=True
            )
        st.markdown('<span class="sec-lbl">DIRECTRICES ADICIONALES</span>', unsafe_allow_html=True)
        instrucciones = st.text_area(
            "_", height=56, label_visibility="collapsed",
            value="Lenguaje médico experto. Sin asteriscos. Solo clasificaciones respaldadas.",
            key="instrucciones_area"
        )

    # ── APARIENCIA ──
    with st.expander("▸  APARIENCIA", expanded=False):
        st.markdown('<span class="sec-lbl">TEMA</span>', unsafe_allow_html=True)
        for nombre in TEMAS:
            activo = nombre == st.session_state.tema
            lbl = f"{'▶ ' if activo else '  '}{nombre.upper()}"
            if st.button(lbl, key=f"t_{nombre}", use_container_width=True):
                st.session_state.tema = nombre
                st.rerun()

        st.markdown('<span class="sec-lbl">ALTURA EDITOR</span>', unsafe_allow_html=True)
        h = st.slider("_", 280, 1100, st.session_state.editor_h, 40,
                      label_visibility="collapsed")
        if h != st.session_state.editor_h:
            st.session_state.editor_h = h
            st.rerun()

    # ── HISTORIAL (NUEVO) ──
    if st.session_state.historial:
        with st.expander(f"▸  HISTORIAL ({len(st.session_state.historial)})", expanded=False):
            st.markdown('<span class="sec-lbl">INFORMES RECIENTES</span>', unsafe_allow_html=True)
            for i, entry in enumerate(st.session_state.historial):
                label = f"{i+1}. {entry['modalidad'][:3].upper()} · {entry['region']}"
                if st.button(label, key=f"hist_{i}", use_container_width=True):
                    st.session_state.reporte_texto = entry["texto"]
                    st.session_state.reporte_html = entry["html"]
                    st.rerun()

    # ── CTA ──
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    procesar = st.button("◈  GENERAR INFORME", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
    ca, cb = st.columns(2)
    with ca:
        if st.button("PURGAR DICTADO", use_container_width=True):
            st.session_state.dictado = ""
            st.session_state.audio_procesado_id = None
            st.rerun()
    with cb:
        if st.button("LIMPIAR EDITOR", use_container_width=True):
            st.session_state.reporte_html = ""
            st.session_state.reporte_texto = ""
            st.rerun()

# ──────────────────────────────────────────────────────────────
# PROCESAMIENTO PRINCIPAL
# ──────────────────────────────────────────────────────────────
if procesar:
    if not api_key:
        st.warning("⚠ API Key requerida para generar el informe.")
    elif not st.session_state.dictado.strip():
        st.warning("⚠ Ingresa hallazgos o realiza un dictado antes de generar.")
    else:
        client = get_openai_client(api_key, st.session_state.modelo_sel)
        model_id = get_model_id(st.session_state.modelo_sel)
        plantilla = st.session_state.plantilla_txt
        tiene_tabla = st.session_state.tiene_tabla
        tabla_instruc = (
            "La plantilla contiene tablas marcadas con [TABLA]. Complétalas con valores del dictado en Markdown."
            if tiene_tabla else
            "NO hay tablas en la plantilla. PROHIBIDO generar tablas bajo ninguna circunstancia."
        )

        # MEJORA: Prompt con estructura explícita y Few-Shot para mayor consistencia
        plantilla_default = "TÉCNICA\nHALLAZGOS\nIMPRESIÓN DIAGNÓSTICA"
        plantilla_usar = plantilla if plantilla else plantilla_default
        prompt_sistema = f"""Eres AURA, sistema de inteligencia radiológica de alta precisión clínica.
Tu tarea es redactar un informe radiológico estructurado de {modalidad} para la región: {region}.

{REGLAS}

TABLAS: {tabla_instruc}

PLANTILLA A SEGUIR:
{plantilla_usar}

DIRECTRICES ADICIONALES: {instrucciones}

FORMATO DE SALIDA:
- Usa MAYÚSCULAS para los títulos de sección (TÉCNICA, HALLAZGOS, IMPRESIÓN DIAGNÓSTICA).
- Usa • para viñetas en la impresión diagnóstica.
- No uses asteriscos Markdown (*) para negritas; usa MAYÚSCULAS para énfasis.
- Sé morfológicamente preciso: incluye medidas, grados de clasificación y localización anatómica.
"""

        with st.spinner("Generando informe..."):
            try:
                res = client.chat.completions.create(
                    model=model_id,
                    messages=[
                        {"role": "system", "content": prompt_sistema},
                        {"role": "user", "content": f"DICTADO / HALLAZGOS DEL RADIÓLOGO:\n{st.session_state.dictado}"}
                    ],
                    temperature=0.1,
                    max_tokens=2048,
                )
                txt = res.choices[0].message.content
                html = texto_a_html(txt)
                st.session_state.reporte_texto = txt
                st.session_state.reporte_html = html
                # NUEVO: Guardar en historial automáticamente
                guardar_en_historial(modalidad, region, txt, html)
                st.rerun()
            except Exception as e:
                st.error(f"Error al generar el informe: {e}")

# ══════════════════════════════════════════════════════════════
# PANEL DERECHO — EDITOR
# ══════════════════════════════════════════════════════════════
with col_r:
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    contenido = st.session_state.reporte_html or """<b>RESONANCIA MAGNÉTICA · RODILLA DERECHA</b><br><br>
<b>TÉCNICA</b><br>
Secuencias multiplanares T1, DP con supresión grasa, T2 y STIR en planos axial, coronal y sagital. Sin contraste.<br><br>
<b>HALLAZGOS</b><br><br>
<b>MENISCOS</b><br>
Menisco medial: señal grado III de Stoller en cuerpo y cuerno posterior — desgarro horizontal. Extrusión de 3 mm en plano coronal.<br>
Menisco lateral: morfología e intensidad conservadas.<br><br>
<b>LIGAMENTOS</b><br>
LCA: señal heterogénea en tercio proximal, lesión parcial grado I (Hope &amp; Feagin). LCP, LCM, LCL íntegros.<br><br>
<b>CARTÍLAGO</b><br>
Adelgazamiento condral focal grado III ICRS en platillo tibial medial (12 mm). Esclerosis subcondral y edema óseo reactivo.<br><br>
<b>ESPACIO ARTICULAR</b><br>
Pinzamiento femorotibial medial de 3 mm. Osteofitos marginales en cóndilos femorales y platillos tibiales, predominio medial.<br><br>
<b>IMPRESIÓN DIAGNÓSTICA</b><br>
<li>Desgarro horizontal de menisco medial, Stoller grado III, extrusión de 3 mm — significativo.</li>
<li>Lesión parcial LCA grado I Hope &amp; Feagin.</li>
<li>Condropatía grado III ICRS en compartimento medial con esclerosis subcondral. Gonartrosis grado II Kellgren-Lawrence.</li>"""

    eH = st.session_state.editor_h
    frameH = eH + 92

    html_editor = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500&family=JetBrains+Mono:wght@300;400;500&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --base:{T['base']};--surface:{T['surface']};
  --glass:{T['glass']};--border:{T['glass_border']};
  --glow:{T['glow']};--accent:{T['accent']};--accent2:{T['accent2']};
  --text:{T['text']};--dim:{T['text_dim']};--ghost:{T['text_ghost']};
}}
html,body{{
  height:{frameH}px;overflow:hidden;
  display:flex;flex-direction:column;
  background:var(--base);
  font-family:'JetBrains Mono',monospace;
}}
body::before{{
  content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
  background:repeating-linear-gradient(0deg,{T['scan_line']} 0px,transparent 1px,transparent 3px);
}}
/* ── TOOLBAR ── */
.tb{{
  flex-shrink:0;position:relative;z-index:10;
  background:var(--base);
  border-bottom:1px solid var(--border);
  padding:4px 12px;
  display:flex;align-items:center;gap:4px;
  flex-wrap:nowrap;overflow-x:auto;
}}
.tg{{display:flex;align-items:center;gap:2px;padding-right:8px;border-right:1px solid var(--ghost);flex-shrink:0}}
.tg:last-child{{border-right:none}}
.tb-btn{{
  background:none;border:1px solid transparent;
  color:var(--dim);font-size:10px;
  padding:3px 5px;border-radius:1px;cursor:pointer;
  transition:all .15s;min-width:22px;text-align:center;
  font-family:'JetBrains Mono',monospace;letter-spacing:.02em;
}}
.tb-btn:hover{{border-color:var(--border);color:var(--text);}}
.tb-btn.on{{border-color:var(--accent2);color:var(--accent);box-shadow:0 0 6px var(--glow);}}
.tb-sel{{
  background:transparent;border:1px solid var(--border);
  color:var(--dim);font-size:9px;letter-spacing:.05em;
  font-family:'JetBrains Mono',monospace;
  padding:2px 4px;border-radius:1px;outline:none;appearance:none;cursor:pointer;
}}
.tb-sel:focus{{border-color:var(--accent2);}}
.cd{{
  width:12px;height:12px;border-radius:50%;cursor:pointer;
  border:1px solid transparent;transition:all .12s;flex-shrink:0;
}}
.cd:hover,.cd.on{{border-color:var(--accent);box-shadow:0 0 5px var(--glow);}}
.tl{{font-size:8px;letter-spacing:.2em;color:var(--ghost);white-space:nowrap;}}
/* ── EDITOR AREA ── */
.ew{{
  flex:1;overflow-y:auto;
  padding:16px 20px;min-height:0;
  background:var(--surface);
  scrollbar-width:thin;scrollbar-color:var(--accent2) transparent;
}}
.ew::-webkit-scrollbar{{width:2px;}}
.ew::-webkit-scrollbar-thumb{{background:var(--accent2);}}
.doc{{
  min-height:100%;padding:24px 32px;
  outline:none;border-radius:1px;
  font-family:'JetBrains Mono',monospace;
  font-size:12.5px;line-height:1.8;
  color:#1a1a1a;background:#ffffff;
  transition:background .25s,color .25s;
}}
.doc b,.doc strong{{font-weight:600;}}
.doc li{{margin-left:18px;margin-bottom:2px;}}
.doc hr{{border:none;border-top:1px solid #e0e0e0;margin:10px 0;}}
.doc table{{border-collapse:collapse;width:100%;margin:8px 0;font-size:11.5px;}}
.doc td,.doc th{{border:1px solid #e0e0e0;padding:4px 10px;}}
.doc th{{background:#f8f8f8;font-weight:600;}}
/* ── ACTION STRIP ── */
.as{{
  flex-shrink:0;position:relative;z-index:10;
  background:var(--base);
  border-top:1px solid var(--border);
  padding:6px 12px;
  display:flex;align-items:center;gap:6px;
}}
.as::before{{
  content:'';position:absolute;top:0;left:0;right:0;height:1px;
  background:linear-gradient(90deg,transparent,var(--accent),transparent);
  opacity:.2;
}}
.as-btn{{
  background:transparent;border:1px solid var(--border);
  color:var(--dim);font-size:8.5px;letter-spacing:.18em;
  text-transform:uppercase;
  padding:4px 10px;border-radius:1px;cursor:pointer;
  transition:all .15s;display:flex;align-items:center;gap:4px;
  font-family:'JetBrains Mono',monospace;
}}
.as-btn:hover{{border-color:var(--accent2);color:var(--text);box-shadow:0 0 8px var(--glow);}}
.prog{{margin-left:auto;display:flex;align-items:center;gap:6px;}}
.prog-bg{{width:56px;height:1px;background:var(--ghost);position:relative;}}
.prog-fill{{position:absolute;top:0;left:0;height:100%;background:var(--accent);transition:width .4s;box-shadow:0 0 4px var(--accent);}}
.prog-pct{{font-size:8px;letter-spacing:.1em;color:var(--dim);}}
/* ── WORD COUNT ── */
.wc{{font-size:8px;letter-spacing:.1em;color:var(--ghost);margin-left:8px;}}
</style>
</head>
<body>

<div class="tb">
  <div class="tg">
    <select class="tb-sel" id="fnt" onchange="applyFont(this.value)" style="width:80px">
      <option value="'JetBrains Mono',monospace" selected>JetBrains</option>
      <option value="'Space Grotesk',sans-serif">Grotesk</option>
      <option value="'Georgia',serif">Georgia</option>
      <option value="'Calibri',sans-serif">Calibri</option>
      <option value="'Arial',sans-serif">Arial</option>
      <option value="'Times New Roman',serif">Times</option>
    </select>
    <select class="tb-sel" id="sz" onchange="applySize(this.value)" style="width:38px">
      <option value="9">9</option><option value="10">10</option>
      <option value="11">11</option><option value="12">12</option>
      <option value="12.5" selected>12.5</option>
      <option value="13">13</option><option value="14">14</option>
      <option value="16">16</option><option value="18">18</option>
    </select>
  </div>
  <div class="tg">
    <button class="tb-btn" id="bB" onclick="fmt('bold')" title="Negrita"><b>B</b></button>
    <button class="tb-btn" id="bI" onclick="fmt('italic')" title="Cursiva"><i>I</i></button>
    <button class="tb-btn" id="bU" onclick="fmt('underline')" title="Subrayado"><u>U</u></button>
  </div>
  <div class="tg">
    <button class="tb-btn" onclick="fmt('justifyLeft')" title="Izquierda"><i class="ti ti-align-left"></i></button>
    <button class="tb-btn" onclick="fmt('justifyCenter')" title="Centro"><i class="ti ti-align-center"></i></button>
    <button class="tb-btn" onclick="fmt('justifyRight')" title="Derecha"><i class="ti ti-align-right"></i></button>
    <button class="tb-btn" onclick="fmt('justifyFull')" title="Justificado"><i class="ti ti-align-justified"></i></button>
  </div>
  <div class="tg">
    <button class="tb-btn" onclick="fmt('insertUnorderedList')" title="Viñetas"><i class="ti ti-list"></i></button>
    <button class="tb-btn" onclick="fmt('insertOrderedList')" title="Numerada"><i class="ti ti-list-numbers"></i></button>
    <button class="tb-btn" onclick="insHR()" title="Separador">—</button>
  </div>
  <div class="tg" style="gap:5px;align-items:center;">
    <span class="tl">FONDO</span>
    <div class="cd on" style="background:#fff;border:1px solid #ddd" onclick="setBg(this,'#fff','#1a1a1a')" title="Blanco"></div>
    <div class="cd" style="background:#f5f0e8" onclick="setBg(this,'#f5f0e8','#1a0e00')" title="Pergamino"></div>
    <div class="cd" style="background:#0a1018" onclick="setBg(this,'#0a1018','#c8e8f8')" title="DICOM"></div>
    <div class="cd" style="background:#000000" onclick="setBg(this,'#000','#00e8b0')" title="Terminal"></div>
    <div class="cd" style="background:#f0f4fa" onclick="setBg(this,'#f0f4fa','#1a2540')" title="Clínico"></div>
  </div>
  <div class="tg">
    <button class="tb-btn" onclick="copyClean()" title="Copiar texto limpio"><i class="ti ti-copy"></i></button>
    <button class="tb-btn" onclick="printDoc()" title="Imprimir / Guardar PDF"><i class="ti ti-printer"></i></button>
  </div>
</div>

<div class="ew">
  <div class="doc" id="doc" contenteditable="true" spellcheck="false">{contenido}</div>
</div>

<div class="as">
  <span class="wc" id="wc">0 palabras</span>
  <div class="prog">
    <div class="prog-bg"><div class="prog-fill" id="pf" style="width:0%"></div></div>
    <span class="prog-pct" id="pp">0%</span>
  </div>
</div>

<script>
var doc=document.getElementById('doc');
doc.style.background='#ffffff';
doc.style.color='#1a1a1a';

function fmt(c){{doc.focus();document.execCommand(c,false,null);upd();}}
function upd(){{
  ['Bold','Italic','Underline'].forEach(function(c){{
    var b=document.getElementById('b'+c[0]);
    if(b)b.classList.toggle('on',document.queryCommandState(c.toLowerCase()));
  }});
}}
function applyFont(f){{doc.style.fontFamily=f;}}
function applySize(s){{doc.style.fontSize=s+'px';}}
function setBg(el,bg,col){{
  doc.style.background=bg;doc.style.color=col;
  document.querySelectorAll('.cd').forEach(function(d){{d.classList.remove('on');}});
  el.classList.add('on');
}}
function insHR(){{
  doc.focus();
  document.execCommand('insertHTML',false,'<hr style="border:none;border-top:1px solid #e0e0e0;margin:10px 0"><br>');
}}
function calcPct(){{
  var t=doc.innerText.toUpperCase();
  var f=['TÉCNICA','HALLAZGOS','IMPRESIÓN'].filter(function(s){{return t.includes(s);}}).length;
  var w=t.split(/\\s+/).filter(Boolean).length;
  return Math.min(100,Math.round((f/3)*60+Math.min(w/150,1)*40));
}}
function updBar(){{
  var s=calcPct();
  document.getElementById('pf').style.width=s+'%';
  document.getElementById('pp').textContent=s+'%';
  // NUEVO: contador de palabras
  var words=doc.innerText.trim().split(/\\s+/).filter(Boolean).length;
  document.getElementById('wc').textContent=words+' palabras';
}}
doc.addEventListener('input',updBar);
doc.addEventListener('keyup',upd);
doc.addEventListener('mouseup',upd);
window.addEventListener('load',function(){{updBar();}});

function copyClean(){{
  var t=doc.innerText;
  if(navigator.clipboard&&navigator.clipboard.writeText){{
    navigator.clipboard.writeText(t).then(function(){{toast('COPIADO');}});
  }}else{{
    var ta=document.createElement('textarea');
    ta.value=t;ta.style.cssText='position:fixed;opacity:0';
    document.body.appendChild(ta);ta.select();
    document.execCommand('copy');document.body.removeChild(ta);
    toast('COPIADO');
  }}
}}

// NUEVO: Imprimir / exportar a PDF desde el navegador
function printDoc(){{
  var w=window.open('','_blank');
  w.document.write('<html><head><title>AURA · Informe Radiológico</title>');
  w.document.write('<style>body{{font-family:Calibri,sans-serif;font-size:12pt;line-height:1.7;margin:2cm;color:#111}}');
  w.document.write('b,strong{{font-weight:600}}table{{border-collapse:collapse;width:100%}}');
  w.document.write('td,th{{border:1px solid #ccc;padding:4px 8px}}th{{background:#f0f0f0}}');
  w.document.write('</style></head><body>');
  w.document.write(doc.innerHTML);
  w.document.write('</body></html>');
  w.document.close();
  w.focus();
  setTimeout(function(){{w.print();}},400);
}}

function toast(m){{
  var el=document.createElement('div');
  el.textContent=m;
  el.style.cssText='position:fixed;bottom:52px;left:50%;transform:translateX(-50%);'
    +'background:transparent;color:var(--accent);border:1px solid var(--accent);'
    +'padding:4px 14px;border-radius:1px;font-size:8px;letter-spacing:.2em;'
    +'font-family:JetBrains Mono,monospace;z-index:9999;pointer-events:none;'
    +'box-shadow:0 0 12px var(--glow);';
  document.body.appendChild(el);
  setTimeout(function(){{document.body.removeChild(el);}},1600);
}}
</script>
</body>
</html>"""

    components.html(html_editor, height=frameH, scrolling=False)

    # ── ACCIONES IA ──
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1.5, 1.8, 1])

    with c1:
        if st.button("◈  OPTIMIZAR CONCLUSIÓN", use_container_width=True):
            if not api_key:
                st.warning("API Key requerida.")
            elif not st.session_state.reporte_texto:
                st.warning("Genera un informe primero.")
            else:
                client = get_openai_client(api_key, st.session_state.modelo_sel)
                model_id = get_model_id(st.session_state.modelo_sel)
                with st.spinner("Optimizando impresión diagnóstica..."):
                    try:
                        r = client.chat.completions.create(
                            model=model_id,
                            messages=[{
                                "role": "user",
                                "content": f"""Eres AURA — optimizador diagnóstico radiológico.
Mejora ÚNICAMENTE el bloque IMPRESIÓN DIAGNÓSTICA del siguiente informe.

{REGLAS}

REGLAS ADICIONALES:
· Morfológicamente precisa y clínicamente accionable.
· Solo clasificaciones con evidencia directa en hallazgos (especifica el criterio morfológico).
· Usa "•" para viñetas. Lenguaje sugerente para seguimiento.
· Devuelve el informe COMPLETO. Conserva exactamente la sección TÉCNICA y HALLAZGOS sin modificarlas.
· Sin asteriscos. Títulos en MAYÚSCULAS.

REPORTE ACTUAL:
{st.session_state.reporte_texto}
"""
                            }],
                            temperature=0.2,
                            max_tokens=2048,
                        )
                        txt = r.choices[0].message.content
                        st.session_state.reporte_texto = txt
                        st.session_state.reporte_html = texto_a_html(txt)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    with c2:
        if st.button("◇  DEFINICIONES & CLASIFICACIONES", use_container_width=True):
            if not api_key:
                st.warning("API Key requerida.")
            elif not st.session_state.reporte_texto:
                st.warning("Genera un informe primero.")
            else:
                client = get_openai_client(api_key, st.session_state.modelo_sel)
                model_id = get_model_id(st.session_state.modelo_sel)
                with st.spinner("Analizando clasificaciones y definiciones..."):
                    try:
                        r = client.chat.completions.create(
                            model=model_id,
                            messages=[{
                                "role": "user",
                                "content": f"""Analiza el siguiente informe radiológico.
Responde con este formato EXACTO. Sin líneas en blanco entre ítems de la misma sección. Una línea en blanco entre secciones.

CLASIFICACIONES USADAS
· Nombre: [nombre completo · autor/sociedad]
· Grado: [grado] — [significado clínico, 1 línea]
· Justificación: [hallazgo específico del texto]
· Ref: [Autor, año, revista]
· URL: [PubMed o sociedad oficial]

CLASIFICACIONES SUGERIDAS
[Solo si hay hallazgo directo que las justifique. Si no: "Ninguna adicional justificada."]
· Nombre: [clasificación]
· Hallazgo que la justifica: [del texto]
· Ref: [Autor, año]
· URL: [URL]

DEFINICIONES
· [Término]: [definición morfológica, 1-2 líneas]

CORRELACIÓN CLÍNICA
[2-3 líneas. Lenguaje sugerente, no prescriptivo.]

Sin asteriscos. Sin negritas markdown.

INFORME:
{st.session_state.reporte_texto}
"""
                            }],
                            temperature=0.15,
                            max_tokens=2048,
                        )
                        st.session_state.defs_resultado = r.choices[0].message.content
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    with c3:
        # MEJORA CRÍTICA: Exportar desde el texto Markdown limpio, no desde HTML
        if st.session_state.reporte_texto:
            docx_bytes = generar_docx_desde_markdown(
                st.session_state.reporte_texto,
                modalidad=modalidad,
                region=region
            )
            st.download_button(
                "↓  EXPORTAR .DOCX",
                data=docx_bytes,
                file_name=f"AURA_{region.replace(' ','_').replace('/','_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    # ── DEFINICIONES ──
    if st.session_state.defs_resultado:
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        with st.expander("◈  DEFINICIONES · CLASIFICACIONES · REFERENCIAS", expanded=True):
            st.markdown(
                f'<div class="defs-box">{st.session_state.defs_resultado}</div>',
                unsafe_allow_html=True
            )
            if st.button("✕  CERRAR DEFINICIONES"):
                st.session_state.defs_resultado = ""
                st.rerun()
