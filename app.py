import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import speech_recognition as sr
import io
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
    "modo": "dictado",   # "dictado" | "hallazgos"
    "plantilla_txt": "",
    "tiene_tabla": False,
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────
def leer_plantilla(file):
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
                if t: partes.append(t)
            elif tag == 'tbl':
                n += 1
                tbl = _tt.Table(el, doc)
                rows = ["| " + " | ".join(c.text.strip() for c in r.cells) + " |" for r in tbl.rows]
                partes.append(f"[TABLA {n}]\n" + "\n".join(rows) + "\n[/TABLA]")
    except Exception:
        partes = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n".join(partes), n > 0

def texto_a_html(texto):
    import re
    lines, buf, in_tbl = [], [], False
    for line in texto.split("\n"):
        s = line.strip()
        if not s:
            if in_tbl:
                lines.append(_tbl_html(buf)); buf = []; in_tbl = False
            lines.append("<br>")
        elif re.match(r'^\|.+\|$', s):
            if all(c in '-| :' for c in s): continue
            in_tbl = True; buf.append(s)
        else:
            if in_tbl:
                lines.append(_tbl_html(buf)); buf = []; in_tbl = False
            if s.isupper() and len(s) < 70 and not s.startswith("•"):
                lines.append(f"<b>{s}</b><br>")
            elif s.startswith("•"):
                lines.append(f"<li>{s[1:].strip()}</li>")
            else:
                lines.append(f"{s}<br>")
    if in_tbl and buf: lines.append(_tbl_html(buf))
    return "\n".join(lines)

def _tbl_html(rows):
    h = '<table style="border-collapse:collapse;width:100%;margin:8px 0;font-size:12px">'
    for i, row in enumerate(rows):
        cols = [c.strip() for c in row.strip("|").split("|")]
        tag = "th" if i == 0 else "td"
        h += "<tr>" + "".join(f"<{tag} style='border:1px solid #ccc;padding:4px 9px'>{c}</{tag}>" for c in cols) + "</tr>"
    return h + "</table>"

def generar_docx(html):
    from html.parser import HTMLParser
    import re

    class P(HTMLParser):
        def __init__(self):
            super().__init__()
            self.doc = Document()
            s = self.doc.styles["Normal"]
            s.font.name = "Calibri"; s.font.size = Pt(11)
            self.b = self.i = self.u = False
            self.para = None; self.in_tbl = False
            self.rows = []; self.row = []; self.cell = ""

        def handle_starttag(self, tag, attrs):
            if tag in ("b","strong"):   self.b = True
            elif tag in ("i","em"):     self.i = True
            elif tag == "u":            self.u = True
            elif tag in ("p","div"):    self.para = self.doc.add_paragraph()
            elif tag == "br":
                if not self.para: self.para = self.doc.add_paragraph()
            elif tag == "li":           self.para = self.doc.add_paragraph(style="List Bullet")
            elif tag == "table":        self.in_tbl = True; self.rows = []
            elif tag == "tr":           self.row = []
            elif tag in ("td","th"):    self.cell = ""

        def handle_endtag(self, tag):
            if tag in ("b","strong"):   self.b = False
            elif tag in ("i","em"):     self.i = False
            elif tag == "u":            self.u = False
            elif tag in ("td","th"):    self.row.append(self.cell); self.cell = ""
            elif tag == "tr":           self.rows.append(self.row)
            elif tag == "table":
                self.in_tbl = False
                if self.rows:
                    cols = max(len(r) for r in self.rows)
                    t = self.doc.add_table(rows=len(self.rows), cols=cols)
                    t.style = "Table Grid"
                    for i, r in enumerate(self.rows):
                        for j, c in enumerate(r):
                            if j < cols: t.rows[i].cells[j].text = c
                self.rows = []

        def handle_data(self, data):
            t = data.strip()
            if not t: return
            if self.in_tbl: self.cell += t; return
            if not self.para: self.para = self.doc.add_paragraph()
            run = self.para.add_run(t + " ")
            run.bold = self.b; run.italic = self.i; run.underline = self.u

    parser = P()
    try:
        parser.feed(html.replace("\n", " ").strip())
    except Exception:
        d = Document()
        for line in re.sub(r"<[^>]+>", "", html).split("\n"):
            d.add_paragraph(line)
        bio = io.BytesIO(); d.save(bio); return bio.getvalue()
    bio = io.BytesIO(); parser.doc.save(bio); return bio.getvalue()

def transcribir(audio):
    r = sr.Recognizer()
    with sr.AudioFile(audio) as src:
        try: return r.recognize_google(r.record(src), language="es-MX")
        except: return ""

def completitud(texto):
    secs = sum(1 for s in ["TÉCNICA","HALLAZGOS","IMPRESIÓN"] if s in texto.upper())
    words = len(texto.split())
    return min(100, int((secs/3)*60 + min(words/150,1)*40))

try:
    api_key = st.secrets["deepseek_key"]
except:
    api_key = ""

T = TEMAS[st.session_state.tema]

# ──────────────────────────────────────────────────────────────
# CSS HOLOGRÁFICO
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
.stTextArea textarea::placeholder {{ color: {T['text_ghost']} !important; }}

/* ── TEXT INPUT ── */
[data-testid="stTextInput"] input {{
    background: transparent !important;
    border: 1px solid {T['glass_border']} !important;
    border-radius: 2px !important;
    color: {T['text']} !important;
    font-size: 11px !important;
    caret-color: {T['accent']} !important;
}}
[data-testid="stTextInput"] input:focus {{
    border-color: {T['accent2']} !important;
    box-shadow: 0 0 10px {T['glow']} !important;
}}

/* ── AUDIO INPUT ── */
[data-testid="stAudioInput"] {{
    background: transparent !important;
    border: 1px solid {T['glass_border']} !important;
    border-radius: 2px !important;
}}

/* ── FILE UPLOADER ── */
[data-testid="stFileUploader"] {{
    background: transparent !important;
    border: 1px dashed {T['glass_border']} !important;
    border-radius: 2px !important;
}}
[data-testid="stFileUploader"] * {{
    color: {T['text_dim']} !important;
    font-size: 10px !important;
}}
[data-testid="stFileUploader"]:hover {{
    border-color: {T['accent2']} !important;
}}

/* ── PRIMARY BUTTON ── */
.btn-primary > div > button {{
    background: transparent !important;
    border: 1px solid {T['accent']} !important;
    border-radius: 2px !important;
    color: {T['accent']} !important;
    font-size: 10px !important;
    letter-spacing: .2em !important;
    text-transform: uppercase !important;
    padding: .65rem 1rem !important;
    width: 100% !important;
    transition: all .2s !important;
    box-shadow: 0 0 16px {T['glow']} !important;
}}
.btn-primary > div > button:hover {{
    background: {T['glass']} !important;
    box-shadow: 0 0 30px {T['glow']}, inset 0 0 20px {T['glow']} !important;
}}

/* ── SECONDARY BUTTONS ── */
.stButton > button {{
    background: transparent !important;
    border: 1px solid {T['glass_border']} !important;
    border-radius: 2px !important;
    color: {T['text_dim']} !important;
    font-size: 9px !important;
    letter-spacing: .12em !important;
    text-transform: uppercase !important;
}}
.stButton > button:hover {{
    border-color: {T['accent2']} !important;
    color: {T['text']} !important;
}}

/* ── EXPANDERS ── */
[data-testid="stExpander"] {{
    background: {T['glass']} !important;
    border: 1px solid {T['glass_border']} !important;
    border-radius: 2px !important;
    margin-bottom: 3px !important;
}}
[data-testid="stExpander"] summary {{
    color: {T['text_dim']} !important;
    font-size: 9px !important;
    letter-spacing: .18em !important;
    text-transform: uppercase !important;
}}
[data-testid="stExpander"] summary:hover {{ color: {T['text']} !important; }}
[data-testid="stExpander"] summary svg {{ display: none !important; }}

/* ── DOWNLOAD ── */
[data-testid="stDownloadButton"] > button {{
    background: transparent !important;
    border: 1px solid {T['glass_border']} !important;
    border-radius: 2px !important;
    color: {T['text_dim']} !important;
    font-size: 9px !important;
    letter-spacing: .12em !important;
}}
[data-testid="stDownloadButton"] > button:hover {{
    border-color: {T['accent']} !important;
    color: {T['accent']} !important;
    box-shadow: 0 0 12px {T['glow']} !important;
}}

/* ── SLIDER ── */
[data-testid="stSlider"] > div {{ padding: 0 !important; }}
[data-testid="stSlider"] [role="slider"] {{
    background: {T['accent']} !important;
    box-shadow: 0 0 8px {T['accent']} !important;
}}

/* ── DEFS BOX ── */
.defs-box {{
    font-size: 10.5px; line-height: 1.45;
    color: {T['text_dim']}; white-space: pre-wrap;
    padding: 14px 16px;
    background: {T['glass']};
    border: 1px solid {T['glass_border']};
    border-radius: 2px;
}}
.defs-box b {{ color: {T['text']}; }}

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
st.markdown(f"""
<div class="aura-bar">
  <div class="aura-logo">
    <div class="aura-pulse"></div>
    AURA
  </div>
  <div class="aura-sep"></div>
  <span class="aura-meta">Radiology Intelligence · v1.0</span>
  <div class="aura-status">
    <div class="aura-online"></div>
    DEEPSEEK · ONLINE
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

    # API Key
    if not api_key:
        st.markdown('<span class="sec-lbl">API KEY</span>', unsafe_allow_html=True)
        api_key = st.text_input("k", type="password", label_visibility="collapsed",
                                placeholder="sk- ···  DeepSeek API Key")
        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    # ── ESTUDIO ──
    with st.expander("▸  ESTUDIO", expanded=True):
        st.markdown('<span class="sec-lbl">MODALIDAD</span>', unsafe_allow_html=True)
        modalidad = st.selectbox("M", MODALIDADES, label_visibility="collapsed")
        st.markdown('<span class="sec-lbl">REGIÓN</span>', unsafe_allow_html=True)
        region = st.selectbox("R", REGIONES, label_visibility="collapsed")

    # ── MODO DE ENTRADA ──
    with st.expander("▸  MODO DE ENTRADA", expanded=True):
        modo_label = "DICTADO DE VOZ" if st.session_state.modo == "dictado" else "HALLAZGOS ESCRITOS"
        st.markdown(f'<span class="sec-lbl">{modo_label}</span>', unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⊙ VOZ", use_container_width=True):
                st.session_state.modo = "dictado"; st.rerun()
        with c2:
            if st.button("⊙ TEXTO", use_container_width=True):
                st.session_state.modo = "hallazgos"; st.rerun()

        if st.session_state.modo == "dictado":
            audio = st.audio_input("_", label_visibility="collapsed")
            if audio:
                txt = transcribir(audio)
                if txt and txt not in st.session_state.dictado:
                    st.session_state.dictado += " " + txt
        else:
            st.markdown('<span class="sec-lbl">HALLAZGOS / IMPRESIÓN</span>', unsafe_allow_html=True)

        dictado = st.text_area(
            "_", value=st.session_state.dictado, height=140,
            label_visibility="collapsed",
            placeholder="Dicta o escribe hallazgos, diagnósticos o ambos...\n\nEj: Desgarro horizontal menisco medial Stoller III, extrusión 3 mm. Osteofitos marginales tibiofemorales mediales."
        )

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
        st.markdown('<span class="sec-lbl">DIRECTRICES</span>', unsafe_allow_html=True)
        instrucciones = st.text_area(
            "_", height=56, label_visibility="collapsed",
            value="Lenguaje médico experto. Sin asteriscos. Solo clasificaciones respaldadas."
        )

    # ── APARIENCIA ──
    with st.expander("▸  APARIENCIA", expanded=False):
        st.markdown('<span class="sec-lbl">TEMA</span>', unsafe_allow_html=True)
        for nombre in TEMAS:
            activo = nombre == st.session_state.tema
            acc = TEMAS[nombre]["accent"]
            lbl = f"{'▶ ' if activo else '  '}{nombre.upper()}"
            if st.button(lbl, key=f"t_{nombre}", use_container_width=True):
                st.session_state.tema = nombre; st.rerun()

        st.markdown('<span class="sec-lbl">ALTURA EDITOR</span>', unsafe_allow_html=True)
        h = st.slider("_", 280, 1100, st.session_state.editor_h, 40,
                       label_visibility="collapsed")
        if h != st.session_state.editor_h:
            st.session_state.editor_h = h; st.rerun()

    # ── CTA ──
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    procesar = st.button("◈  GENERAR INFORME", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
    ca, cb = st.columns(2)
    with ca:
        if st.button("PURGAR", use_container_width=True):
            st.session_state.dictado = ""; st.rerun()
    with cb:
        if st.button("LIMPIAR", use_container_width=True):
            st.session_state.reporte_html = ""
            st.session_state.reporte_texto = ""; st.rerun()

# ──────────────────────────────────────────────────────────────
# PROCESAMIENTO
# ──────────────────────────────────────────────────────────────
if procesar:
    if api_key and dictado.strip():
        client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
        plantilla = st.session_state.plantilla_txt
        tiene_tabla = st.session_state.tiene_tabla
        tabla_instruc = (
            "La plantilla contiene tablas marcadas con [TABLA]. Complétalas con valores del dictado en Markdown."
            if tiene_tabla else
            "NO hay tablas en la plantilla. PROHIBIDO generar tablas bajo ninguna circunstancia."
        )
        prompt = f"""
Eres AURA, sistema de inteligencia radiológica. Redacta un informe de {modalidad} — región: {region}.

{REGLAS}

TABLAS: {tabla_instruc}

PLANTILLA:
{plantilla if plantilla else "TÉCNICA / HALLAZGOS / IMPRESIÓN DIAGNÓSTICA"}

DIRECTRICES: {instrucciones}

ENTRADA DEL RADIÓLOGO:
{dictado}
"""
        with st.spinner(""):
            try:
                res = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "system", "content": prompt}],
                    temperature=0.1
                )
                txt = res.choices[0].message.content
                st.session_state.reporte_texto = txt
                st.session_state.reporte_html = texto_a_html(txt)
                st.rerun()
            except Exception as e:
                st.error(f"{e}")
    elif not api_key:
        st.warning("API Key requerida")
    else:
        st.warning("Ingresa dictado o hallazgos")

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

/* scanlines */
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
.as-btn.prime{{border-color:var(--accent);color:var(--accent);box-shadow:0 0 10px var(--glow);}}
.prog{{margin-left:auto;display:flex;align-items:center;gap:6px;}}
.prog-bg{{width:56px;height:1px;background:var(--ghost);position:relative;}}
.prog-fill{{position:absolute;top:0;left:0;height:100%;background:var(--accent);transition:width .4s;box-shadow:0 0 4px var(--accent);}}
.prog-pct{{font-size:8px;letter-spacing:.1em;color:var(--dim);}}
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
    <button class="tb-btn" onclick="copyClean()" title="Copiar texto"><i class="ti ti-copy"></i></button>
  </div>
</div>

<div class="ew">
  <div class="doc" id="doc" contenteditable="true" spellcheck="false">{contenido}</div>
</div>

<div class="as">
  <button class="as-btn" onclick="optimize()">◈ OPTIMIZAR CONCLUSIÓN</button>
  <button class="as-btn" onclick="getDefs()">◇ DEFINICIONES</button>
  <button class="as-btn prime" onclick="exportDoc()">↓ EXPORTAR .DOCX</button>
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
function optimize(){{window.parent.postMessage({{type:'optimize',content:doc.innerText}},'*');}}
function getDefs(){{window.parent.postMessage({{type:'defs',content:doc.innerText}},'*');}}
function exportDoc(){{window.parent.postMessage({{type:'export',html:doc.innerHTML}},'*');}}
</script>
</body>
</html>"""

    components.html(html_editor, height=frameH, scrolling=False)

    # ── ACCIONES IA ──
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1.5, 1.8, 1])

    with c1:
        if st.button("◈  OPTIMIZAR CONCLUSIÓN", use_container_width=True):
            if api_key and st.session_state.reporte_texto:
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                with st.spinner(""):
                    try:
                        r = client.chat.completions.create(
                            model="deepseek-chat",
                            messages=[{"role":"user","content":f"""
Eres AURA — optimizador diagnóstico.
Mejora ÚNICAMENTE el bloque IMPRESIÓN DIAGNÓSTICA.

{REGLAS}

· Morfológicamente precisa y clínicamente accionable.
· Solo clasificaciones con evidencia directa en hallazgos (especifica el criterio).
· Usa "•" para viñetas. Lenguaje sugerente para seguimiento.
· Devuelve el informe COMPLETO. Conserva Técnica y Hallazgos.
· Cero asteriscos. Títulos en MAYÚSCULAS.

REPORTE:
{st.session_state.reporte_texto}
"""}], temperature=0.2)
                        txt = r.choices[0].message.content
                        st.session_state.reporte_texto = txt
                        st.session_state.reporte_html = texto_a_html(txt)
                        st.rerun()
                    except Exception as e:
                        st.error(str(e))

    with c2:
        if st.button("◇  DEFINICIONES & CLASIFICACIONES", use_container_width=True):
            if api_key and st.session_state.reporte_texto:
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                with st.spinner(""):
                    try:
                        r = client.chat.completions.create(
                            model="deepseek-chat",
                            messages=[{"role":"user","content":f"""
Analiza el informe. Responde con este formato EXACTO.
Sin líneas en blanco entre items de la misma sección. Una línea en blanco entre secciones.

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
"""}], temperature=0.15)
                        st.session_state.defs_resultado = r.choices[0].message.content
                        st.rerun()
                    except Exception as e:
                        st.error(str(e))

    with c3:
        if st.session_state.reporte_texto:
            st.download_button(
                "↓  EXPORTAR",
                data=generar_docx(st.session_state.reporte_html),
                file_name="AURA_Informe.docx",
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
            if st.button("✕  CERRAR"):
                st.session_state.defs_resultado = ""; st.rerun()
