import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Pt
import speech_recognition as sr
import io
from openai import OpenAI

# ==========================================
# CONFIG
# ==========================================
st.set_page_config(
    page_title="Beam AI | PACS Editor v5",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# TEMAS DE INTERFAZ (estilo Eden PACS)
# ==========================================
TEMAS = {
    "Eden Dark": {
        "app_bg": "#070c14",
        "panel_bg": "#0a1020",
        "panel_border": "#162030",
        "topbar_bg": "#060e1a",
        "topbar_border": "#112030",
        "input_bg": "#0c1828",
        "input_border": "#1a3050",
        "input_color": "#7ab8d4",
        "text_primary": "#c8dff0",
        "text_secondary": "#3a6080",
        "text_muted": "#1e3a58",
        "accent": "#1a7abf",
        "accent_light": "#4ab0e8",
        "accent_glow": "#1a5a9a",
        "btn_bg": "#0c1e34",
        "btn_border": "#1a3a5a",
        "btn_color": "#3a90c0",
        "logo_dot": "#2a9ad4",
        "status_color": "#1a9a5a",
        "scrollbar": "#1a3050",
        "expander_bg": "#090f1e",
        "expander_border": "#142030",
        "defs_bg": "#06101e",
        "defs_color": "#6ab0cc",
        "label": "Eden Dark",
    },
    "Eden Light": {
        "app_bg": "#eef2f7",
        "panel_bg": "#f8fafd",
        "panel_border": "#d0dce8",
        "topbar_bg": "#ffffff",
        "topbar_border": "#d0dce8",
        "input_bg": "#f0f5fa",
        "input_border": "#c0d0e0",
        "input_color": "#2a5a80",
        "text_primary": "#1a3050",
        "text_secondary": "#5a7a9a",
        "text_muted": "#8aaccc",
        "accent": "#1a7abf",
        "accent_light": "#0a60a0",
        "accent_glow": "#3a9ad4",
        "btn_bg": "#e8f2fa",
        "btn_border": "#b0cce0",
        "btn_color": "#1a6090",
        "logo_dot": "#1a7abf",
        "status_color": "#1a8a50",
        "scrollbar": "#c0d4e8",
        "expander_bg": "#f0f6fc",
        "expander_border": "#c8daea",
        "defs_bg": "#f0f5fa",
        "defs_color": "#2a5a7a",
        "label": "Eden Light",
    },
    "PACS Clásico": {
        "app_bg": "#0a0a0a",
        "panel_bg": "#111111",
        "panel_border": "#222222",
        "topbar_bg": "#080808",
        "topbar_border": "#1e1e1e",
        "input_bg": "#161616",
        "input_border": "#2a2a2a",
        "input_color": "#aaaaaa",
        "text_primary": "#cccccc",
        "text_secondary": "#555555",
        "text_muted": "#333333",
        "accent": "#00aa66",
        "accent_light": "#00dd88",
        "accent_glow": "#008850",
        "btn_bg": "#141414",
        "btn_border": "#2a2a2a",
        "btn_color": "#00aa66",
        "logo_dot": "#00cc77",
        "status_color": "#00aa66",
        "scrollbar": "#222222",
        "expander_bg": "#0e0e0e",
        "expander_border": "#1e1e1e",
        "defs_bg": "#0c0c0c",
        "defs_color": "#888888",
        "label": "PACS Clásico",
    },
    "Radiology Blue": {
        "app_bg": "#040d18",
        "panel_bg": "#071220",
        "panel_border": "#0e2035",
        "topbar_bg": "#030c16",
        "topbar_border": "#0a1e30",
        "input_bg": "#081828",
        "input_border": "#102840",
        "input_color": "#5ab8e8",
        "text_primary": "#a8d8f8",
        "text_secondary": "#2a6080",
        "text_muted": "#0e3050",
        "accent": "#0a8ad8",
        "accent_light": "#3ab8f8",
        "accent_glow": "#0a6ab8",
        "btn_bg": "#081828",
        "btn_border": "#103858",
        "btn_color": "#2a8ac8",
        "logo_dot": "#0ab0e8",
        "status_color": "#0aaa70",
        "scrollbar": "#0e3050",
        "expander_bg": "#060f1c",
        "expander_border": "#0a1c30",
        "defs_bg": "#050e1a",
        "defs_color": "#4aa8d8",
        "label": "Radiology Blue",
    },
    "Warm Clinical": {
        "app_bg": "#100e0a",
        "panel_bg": "#181410",
        "panel_border": "#2a2018",
        "topbar_bg": "#0e0c08",
        "topbar_border": "#221a10",
        "input_bg": "#1e1812",
        "input_border": "#302418",
        "input_color": "#c8a878",
        "text_primary": "#e0c8a0",
        "text_secondary": "#705030",
        "text_muted": "#3a2818",
        "accent": "#c07830",
        "accent_light": "#e0a060",
        "accent_glow": "#a05820",
        "btn_bg": "#1a1408",
        "btn_border": "#2e2010",
        "btn_color": "#b07830",
        "logo_dot": "#d09040",
        "status_color": "#60a050",
        "scrollbar": "#2a2010",
        "expander_bg": "#141008",
        "expander_border": "#241c0e",
        "defs_bg": "#120e08",
        "defs_color": "#a08860",
        "label": "Warm Clinical",
    },
}

# ==========================================
# CONSTANTES
# ==========================================
REGIONES = [
    "Rodilla", "Columna lumbar", "Columna cervical", "Hombro",
    "Cadera", "Tobillo / Pie", "Muñeca / Mano", "Codo",
    "Cerebro", "Columna dorsal", "Tórax", "Abdomen / Pelvis",
    "Mama", "Tiroides", "Hígado",
]

REGLAS_CLINICAS = """
REGLAS CLÍNICAS ESTRICTAS — NUNCA VIOLAR:

1. TERMINOLOGÍA PRECISA:
   - NO uses "cambios degenerativos" como término genérico.
     Especifica el hallazgo morfológico real: osteofitos marginales, esclerosis subcondral,
     disminución del espacio articular, condromalacia, fibrosis periarticular, etc.
   - NO uses "cambios crónicos" sin especificar el sustrato morfológico.
   - USA descriptores anatómico-morfológicos: "osteofitos marginales tibiofemorales mediales",
     "esclerosis subcondral en platillo tibial medial", "pinzamiento articular de X mm".

2. TABLAS DE MEDIDAS:
   - SOLO genera tablas si la plantilla proporcionada contiene explícitamente una sección [TABLA].
   - Si NO hay plantilla con tabla, NO generes ninguna tabla bajo ninguna circunstancia.
   - Si hay tabla en la plantilla, complétala con los valores mencionados en el dictado,
     en formato Markdown (| columna | columna |).

3. CLASIFICACIONES:
   - Solo incluye clasificaciones que estén directamente respaldadas por los hallazgos del dictado.
   - No asumas grados si no tienes la información suficiente.
   - Especifica el criterio morfológico que justifica el grado asignado.

4. IMPRESIÓN DIAGNÓSTICA:
   - Diagnósticos específicos y morfológicamente precisos.
   - Correlación anatómica-funcional cuando sea pertinente.
   - Lenguaje sugerente para seguimiento: "se sugiere correlación clínica", "puede valorarse".
"""

# ==========================================
# ESTADO
# ==========================================
ESTADO_DEFAULTS = {
    "dictado": "",
    "reporte_html": "",
    "reporte_texto": "",
    "definiciones_resultado": "",
    "editor_height": 580,
    "tema_actual": "Eden Dark",
}
for k, v in ESTADO_DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ==========================================
# HELPERS
# ==========================================
def leer_plantilla(file):
    """Lee párrafos Y tablas del docx. Las tablas se marcan con [TABLA n]."""
    doc = Document(file)
    secciones = []
    tabla_count = 0
    try:
        import docx.text.paragraph as _p
        import docx.table as _t
        for element in doc.element.body:
            tag = element.tag.split('}')[-1]
            if tag == 'p':
                para = _p.Paragraph(element, doc)
                texto = para.text.strip()
                if texto:
                    secciones.append(texto)
            elif tag == 'tbl':
                tabla_count += 1
                tabla = _t.Table(element, doc)
                filas_txt = []
                for row in tabla.rows:
                    celdas = [c.text.strip() for c in row.cells]
                    filas_txt.append("| " + " | ".join(celdas) + " |")
                secciones.append(f"[TABLA {tabla_count}]\n" + "\n".join(filas_txt) + "\n[/TABLA]")
    except Exception:
        secciones = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n".join(secciones)

def texto_a_html(texto):
    """Convierte texto del modelo a HTML. Las tablas solo si vienen del modelo."""
    import re
    lines = []
    in_table = False
    table_rows = []

    for line in texto.split("\n"):
        s = line.strip()
        if not s:
            if in_table:
                # Cerrar tabla
                html_table = '<table style="border-collapse:collapse;width:100%;margin:8px 0">'
                for i, row in enumerate(table_rows):
                    cols = [c.strip() for c in row.strip("|").split("|")]
                    tag = "th" if i == 0 else "td"
                    html_table += "<tr>" + "".join(f"<{tag} style='border:1px solid #ccc;padding:4px 10px'>{c}</{tag}>" for c in cols) + "</tr>"
                html_table += "</table>"
                lines.append(html_table)
                in_table = False
                table_rows = []
            lines.append("<br>")
        elif re.match(r'^\|.+\|$', s):
            if all(c in '-| ' for c in s):
                continue  # saltar separador de tabla markdown
            in_table = True
            table_rows.append(s)
        else:
            if in_table:
                html_table = '<table style="border-collapse:collapse;width:100%;margin:8px 0">'
                for i, row in enumerate(table_rows):
                    cols = [c.strip() for c in row.strip("|").split("|")]
                    tag = "th" if i == 0 else "td"
                    html_table += "<tr>" + "".join(f"<{tag} style='border:1px solid #ccc;padding:4px 10px'>{c}</{tag}>" for c in cols) + "</tr>"
                html_table += "</table>"
                lines.append(html_table)
                in_table = False
                table_rows = []
            if s.isupper() and len(s) < 70 and not s.startswith("•"):
                lines.append(f"<b>{s}</b><br>")
            elif s.startswith("•"):
                lines.append(f"<li>{s[1:].strip()}</li>")
            else:
                lines.append(f"{s}<br>")

    if in_table and table_rows:
        html_table = '<table style="border-collapse:collapse;width:100%;margin:8px 0">'
        for i, row in enumerate(table_rows):
            cols = [c.strip() for c in row.strip("|").split("|")]
            tag = "th" if i == 0 else "td"
            html_table += "<tr>" + "".join(f"<{tag} style='border:1px solid #ccc;padding:4px 10px'>{c}</{tag}>" for c in cols) + "</tr>"
        html_table += "</table>"
        lines.append(html_table)

    return "\n".join(lines)

def generar_docx(html_texto):
    from html.parser import HTMLParser

    class HTMLtoDocx(HTMLParser):
        def __init__(self):
            super().__init__()
            self.doc = Document()
            s = self.doc.styles["Normal"]
            s.font.name = "Arial"
            s.font.size = Pt(11)
            self.bold = self.italic = self.underline = False
            self.current_para = None
            self.in_table = False
            self.table_rows = []
            self.current_row = []
            self.current_cell = ""

        def handle_starttag(self, tag, attrs):
            if tag in ("b","strong"):    self.bold = True
            elif tag in ("i","em"):      self.italic = True
            elif tag == "u":             self.underline = True
            elif tag in ("p","div"):     self.current_para = self.doc.add_paragraph()
            elif tag == "br":
                if not self.current_para: self.current_para = self.doc.add_paragraph()
            elif tag == "li":            self.current_para = self.doc.add_paragraph(style="List Bullet")
            elif tag == "table":         self.in_table = True; self.table_rows = []
            elif tag == "tr":            self.current_row = []
            elif tag in ("td","th"):     self.current_cell = ""
            elif tag == "hr":            self.doc.add_paragraph()

        def handle_endtag(self, tag):
            if tag in ("b","strong"):    self.bold = False
            elif tag in ("i","em"):      self.italic = False
            elif tag == "u":             self.underline = False
            elif tag in ("td","th"):
                self.current_row.append(self.current_cell); self.current_cell = ""
            elif tag == "tr":            self.table_rows.append(self.current_row)
            elif tag == "table":
                self.in_table = False
                if self.table_rows:
                    cols = max(len(r) for r in self.table_rows)
                    tbl = self.doc.add_table(rows=len(self.table_rows), cols=cols)
                    tbl.style = "Table Grid"
                    for i, row in enumerate(self.table_rows):
                        for j, ct in enumerate(row):
                            if j < cols: tbl.rows[i].cells[j].text = ct
                self.table_rows = []

        def handle_data(self, data):
            text = data.strip()
            if not text: return
            if self.in_table: self.current_cell += text; return
            if self.current_para is None: self.current_para = self.doc.add_paragraph()
            run = self.current_para.add_run(text + " ")
            run.bold = self.bold; run.italic = self.italic; run.underline = self.underline

    import re
    clean = html_texto.replace("\n", " ").strip()
    parser = HTMLtoDocx()
    try:
        parser.feed(clean)
    except Exception:
        doc = Document()
        plain = re.sub(r"<[^>]+>", "", html_texto)
        for line in plain.split("\n"): doc.add_paragraph(line)
        bio = io.BytesIO(); doc.save(bio); return bio.getvalue()
    bio = io.BytesIO(); parser.doc.save(bio); return bio.getvalue()

def transcribir_voz(audio_file):
    r = sr.Recognizer()
    with sr.AudioFile(audio_file) as source:
        try: return r.recognize_google(r.record(source), language="es-MX")
        except: return ""

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
except:
    api_key = ""

# ==========================================
# TEMA ACTIVO
# ==========================================
T = TEMAS[st.session_state.tema_actual]

# ==========================================
# CSS DINÁMICO SEGÚN TEMA
# ==========================================
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

html, .stApp {{ background: {T['app_bg']} !important; }}
.block-container {{ padding: 0 !important; max-width: 100% !important; }}
header, footer {{ display: none !important; }}
[data-testid="stToolbar"] {{ display: none !important; }}

/* ── TOPBAR ── */
.beam-topbar {{
    background: {T['topbar_bg']}; border-bottom: 1px solid {T['topbar_border']};
    padding: 9px 20px; display: flex; align-items: center; gap: 14px;
    position: sticky; top: 0; z-index: 9999;
}}
.beam-logo {{
    font-family: 'Inter', sans-serif; font-weight: 600; font-size: 13px;
    color: {T['text_primary']}; letter-spacing: .15em;
    display: flex; align-items: center; gap: 7px;
}}
.ldot {{ width: 7px; height: 7px; border-radius: 50%; background: {T['logo_dot']}; display: inline-block; box-shadow: 0 0 6px {T['logo_dot']}; }}
.tbadge {{
    font-size: 10px; color: {T['accent_light']}; background: {T['btn_bg']};
    border: 1px solid {T['btn_border']}; border-radius: 3px; padding: 2px 8px;
    font-family: 'IBM Plex Mono', monospace; letter-spacing: .04em;
}}
.tstat {{
    margin-left: auto; font-size: 10px; color: {T['status_color']};
    font-family: 'IBM Plex Mono', monospace; display: flex; align-items: center; gap: 5px;
}}
.sdot {{ width: 5px; height: 5px; border-radius: 50%; background: {T['status_color']}; display: inline-block; box-shadow: 0 0 4px {T['status_color']}; }}

/* ── TEMA SELECTOR PILLS ── */
.tema-pills {{ display: flex; gap: 4px; flex-wrap: wrap; margin-bottom: 6px; }}
.tema-pill {{
    font-size: 10px; font-family: 'IBM Plex Mono', monospace;
    color: {T['text_secondary']}; background: {T['btn_bg']};
    border: 1px solid {T['btn_border']}; border-radius: 3px;
    padding: 2px 8px; cursor: pointer;
}}
.tema-pill.active {{ color: {T['accent_light']}; border-color: {T['accent']}; background: {T['btn_bg']}; }}

/* ── PANEL LABELS ── */
.plabel {{
    font-size: 9px !important; letter-spacing: .18em !important;
    color: {T['text_muted']} !important; text-transform: uppercase !important;
    font-family: 'IBM Plex Mono', monospace !important;
    margin-bottom: 2px !important; margin-top: 0 !important; display: block;
}}

/* ── INPUTS ── */
[data-testid="stSelectbox"] > div > div {{
    background: {T['input_bg']} !important; border: 1px solid {T['input_border']} !important;
    border-radius: 5px !important; color: {T['input_color']} !important;
    font-family: 'IBM Plex Mono', monospace !important; font-size: 11px !important;
}}
[data-testid="stSelectbox"] > div > div:hover {{ border-color: {T['accent']} !important; }}

.stTextArea textarea {{
    background: {T['input_bg']} !important; border: 1px solid {T['input_border']} !important;
    border-radius: 6px !important; color: {T['input_color']} !important;
    font-family: 'IBM Plex Mono', monospace !important; font-size: 11px !important; line-height: 1.55 !important;
}}
.stTextArea textarea:focus {{ border-color: {T['accent']} !important; box-shadow: none !important; }}

[data-testid="stTextInput"] input {{
    background: {T['input_bg']} !important; border: 1px solid {T['input_border']} !important;
    border-radius: 5px !important; color: {T['input_color']} !important;
    font-family: 'IBM Plex Mono', monospace !important; font-size: 11px !important;
}}

[data-testid="stAudioInput"] {{
    background: {T['input_bg']} !important; border: 1px solid {T['input_border']} !important;
    border-radius: 7px !important;
}}
[data-testid="stFileUploader"] {{
    background: {T['input_bg']} !important; border: 1px dashed {T['input_border']} !important;
    border-radius: 6px !important;
}}
[data-testid="stFileUploader"] * {{
    color: {T['text_secondary']} !important; font-family: 'IBM Plex Mono', monospace !important; font-size: 10px !important;
}}

/* ── BOTÓN PRINCIPAL ── */
.btn-main > div > button {{
    background: {T['accent_glow']} !important; border: 1px solid {T['accent']} !important;
    color: {T['accent_light']} !important; font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important; font-size: 12px !important; letter-spacing: .06em !important;
    border-radius: 6px !important; padding: .7rem 1rem !important; width: 100% !important;
}}
.btn-main > div > button:hover {{ filter: brightness(1.15) !important; }}

/* ── BOTONES SECUNDARIOS ── */
.stButton > button {{
    background: {T['btn_bg']} !important; border: 1px solid {T['btn_border']} !important;
    color: {T['btn_color']} !important; font-family: 'IBM Plex Mono', monospace !important;
    font-size: 10px !important; border-radius: 5px !important;
}}
.stButton > button:hover {{ border-color: {T['accent']} !important; color: {T['accent_light']} !important; }}

/* ── EXPANDERS ── */
[data-testid="stExpander"] {{
    background: {T['expander_bg']} !important; border: 1px solid {T['expander_border']} !important;
    border-radius: 6px !important; margin-bottom: 3px !important;
}}
[data-testid="stExpander"] summary {{
    color: {T['btn_color']} !important; font-family: 'IBM Plex Mono', monospace !important;
    font-size: 10px !important; letter-spacing: .06em !important;
}}
[data-testid="stExpander"] summary:hover {{ color: {T['accent_light']} !important; }}

/* ── DOWNLOAD ── */
[data-testid="stDownloadButton"] > button {{
    background: {T['btn_bg']} !important; border: 1px solid {T['btn_border']} !important;
    color: {T['btn_color']} !important; font-family: 'IBM Plex Mono', monospace !important;
    font-size: 10px !important; border-radius: 5px !important;
}}
[data-testid="stDownloadButton"] > button:hover {{ color: {T['accent_light']} !important; }}

/* ── SLIDER ── */
[data-testid="stSlider"] > div {{ padding: 0 !important; }}

/* ── DEFINICIONES ── */
.defs-box {{
    font-family: 'IBM Plex Mono', monospace; font-size: 11px;
    color: {T['defs_color']}; line-height: 1.45;
    background: {T['defs_bg']}; padding: 14px 16px;
    border-radius: 6px; border: 1px solid {T['expander_border']};
    white-space: pre-wrap;
}}

::-webkit-scrollbar {{ width: 3px; height: 3px; }}
::-webkit-scrollbar-track {{ background: {T['app_bg']}; }}
::-webkit-scrollbar-thumb {{ background: {T['scrollbar']}; border-radius: 2px; }}
hr {{ border-color: {T['panel_border']} !important; margin: 6px 0 !important; }}
</style>
""", unsafe_allow_html=True)

# ==========================================
# TOPBAR
# ==========================================
st.markdown(f"""
<div class="beam-topbar">
    <div class="beam-logo"><span class="ldot"></span> BEAM AI</div>
    <span class="tbadge">v5.0 · PACS Editor · {T['label']}</span>
    <div class="tstat"><span class="sdot"></span> DeepSeek · activo</div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# LAYOUT
# ==========================================
col_izq, col_centro = st.columns([1, 2.7], gap="small")

# ─────────────────────────────────────────
# PANEL IZQUIERDO
# ─────────────────────────────────────────
with col_izq:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    if not api_key:
        api_key = st.text_input("API Key", type="password",
                                label_visibility="collapsed",
                                placeholder="sk-... DeepSeek API Key")

    # ── Selector de tema ──
    with st.expander("⊞  TEMA DE INTERFAZ", expanded=False):
        st.markdown('<span class="plabel">APARIENCIA</span>', unsafe_allow_html=True)
        for nombre_tema in TEMAS:
            activo = nombre_tema == st.session_state.tema_actual
            if st.button(
                f"{'▶ ' if activo else '  '}{nombre_tema}",
                key=f"tema_{nombre_tema}",
                use_container_width=True
            ):
                st.session_state.tema_actual = nombre_tema
                st.rerun()

    with st.expander("⊞  MODALIDAD & REGIÓN", expanded=True):
        st.markdown('<span class="plabel">MODALIDAD</span>', unsafe_allow_html=True)
        modalidad = st.selectbox("Modalidad", [
            "Resonancia Magnética", "Tomografía Computarizada",
            "Radiografía", "Ultrasonido", "PET-CT"
        ], label_visibility="collapsed")
        st.markdown('<span class="plabel">REGIÓN ANATÓMICA</span>', unsafe_allow_html=True)
        region = st.selectbox("Región", REGIONES, label_visibility="collapsed")

    with st.expander("⊞  DICTADO DE VOZ", expanded=True):
        audio_data = st.audio_input("Voz", label_visibility="collapsed")
        if audio_data:
            nuevo = transcribir_voz(audio_data)
            if nuevo and nuevo not in st.session_state.dictado:
                st.session_state.dictado += " " + nuevo
        st.markdown('<span class="plabel">SEÑAL TRANSCRITA</span>', unsafe_allow_html=True)
        dictado = st.text_area("Dictado", value=st.session_state.dictado,
                               height=130, label_visibility="collapsed",
                               placeholder="Dictado o escritura manual...\n\nEj: Desgarro horizontal menisco medial Stoller III, extrusión 3 mm, osteofitos marginales tibiales...")

    with st.expander("⊞  CONFIGURACIÓN", expanded=False):
        st.markdown('<span class="plabel">PLANTILLA BASE (.docx)</span>', unsafe_allow_html=True)
        archivo_base = st.file_uploader("Plantilla", type=["docx"], label_visibility="collapsed")
        plantilla_txt = leer_plantilla(archivo_base) if archivo_base else ""
        tiene_tabla = "[TABLA" in plantilla_txt if plantilla_txt else False
        if plantilla_txt:
            st.markdown(
                f'<span style="font-size:10px;color:{"#2ecc71" if tiene_tabla else T["btn_color"]};'
                f'font-family:\'IBM Plex Mono\',monospace;">'
                f'{"✓ Plantilla con tablas" if tiene_tabla else "✓ Plantilla cargada"}</span>',
                unsafe_allow_html=True
            )

        st.markdown('<span class="plabel">DIRECTRICES DE ESTILO</span>', unsafe_allow_html=True)
        instrucciones = st.text_area("Directrices", height=65, label_visibility="collapsed",
                                     value="Lenguaje médico experto. Sin asteriscos. Solo clasificaciones respaldadas por los hallazgos.")

    with st.expander("⊞  TAMAÑO DEL EDITOR", expanded=False):
        nueva_altura = st.slider(
            "Altura", min_value=300, max_value=1200,
            value=st.session_state.editor_height, step=50,
            label_visibility="collapsed"
        )
        if nueva_altura != st.session_state.editor_height:
            st.session_state.editor_height = nueva_altura
            st.rerun()
        st.markdown(
            f'<span style="font-size:10px;color:{T["text_secondary"]};font-family:\'IBM Plex Mono\',monospace;">'
            f'Altura actual: {st.session_state.editor_height}px</span>',
            unsafe_allow_html=True
        )

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    st.markdown('<div class="btn-main">', unsafe_allow_html=True)
    procesar = st.button("⬡  PROCESAR INFORME", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("⌫ Purgar", use_container_width=True):
            st.session_state.dictado = ""
            st.rerun()
    with c2:
        if st.button("⟳ Limpiar", use_container_width=True):
            st.session_state.reporte_html = ""
            st.session_state.reporte_texto = ""
            st.rerun()

# ─────────────────────────────────────────
# PROCESAMIENTO IA
# ─────────────────────────────────────────
if procesar:
    if api_key and dictado.strip():
        client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

        instruccion_tabla = (
            "La plantilla contiene tablas. Complétalas con los valores del dictado en formato Markdown."
            if tiene_tabla else
            "NO hay tablas en la plantilla. NO generes ninguna tabla en el informe bajo ninguna circunstancia."
        )

        prompt = f"""
Eres Beam AI, asistente experto en interpretación radiológica.
Redacta un informe de {modalidad} para región: {region}.

{REGLAS_CLINICAS}

INSTRUCCIÓN SOBRE TABLAS: {instruccion_tabla}

PLANTILLA BASE:
{plantilla_txt if plantilla_txt else "TÉCNICA / HALLAZGOS / IMPRESIÓN DIAGNÓSTICA"}

DIRECTRICES ADICIONALES: {instrucciones}

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
                st.session_state.reporte_html = texto_a_html(texto)
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
    elif not api_key:
        st.warning("Ingresa tu API Key.")
    else:
        st.warning("Ingresa dictado o descripción clínica.")

# ─────────────────────────────────────────
# PANEL CENTRAL — Editor
# ─────────────────────────────────────────
with col_centro:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    contenido_inicial = st.session_state.reporte_html or """<b>RESONANCIA MAGNÉTICA DE RODILLA DERECHA</b><br>
<br><b>TÉCNICA</b><br>
Secuencias multiplanares en T1, DP con supresión grasa (DPFS), T2 y STIR en planos axial, coronal y sagital, sin contraste.<br>
<br><b>HALLAZGOS</b><br>
<br><b>MENISCOS</b><br>
Menisco medial: alteración de señal grado III de Stoller en cuerpo y cuerno posterior, compatible con desgarro horizontal. Extrusión de 3 mm en el plano coronal.<br>
Menisco lateral: morfología e intensidad de señal conservadas.<br>
<br><b>LIGAMENTOS</b><br>
Ligamento cruzado anterior con señal heterogénea en tercio proximal, compatible con lesión parcial grado I de Hope &amp; Feagin. LCP, LCM y LCL sin alteraciones.<br>
<br><b>CARTÍLAGO</b><br>
Adelgazamiento condral focal grado III de ICRS en platillo tibial medial, extensión de 12 mm. Esclerosis subcondral y edema óseo reactivo asociado.<br>
<br><b>ESPACIO ARTICULAR</b><br>
Pinzamiento femorotibial medial de 3 mm. Osteofitos marginales en cóndilos femorales y platillos tibiales de predominio medial.<br>
<br><b>IMPRESIÓN DIAGNÓSTICA</b><br>
<li>Desgarro horizontal de menisco medial, grado III de Stoller, con extrusión de 3 mm.</li>
<li>Lesión parcial del LCA, grado I de Hope &amp; Feagin.</li>
<li>Condropatía grado III ICRS en compartimento femorotibial medial con esclerosis subcondral. Compatible con gonartrosis grado II de Kellgren-Lawrence.</li>"""

    editor_h = st.session_state.editor_height
    iframe_h = editor_h + 95

    # Colores del editor pasados como variables JS
    editor_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{height:{iframe_h}px;overflow:hidden;display:flex;flex-direction:column;
  background:{T['panel_bg']};font-family:'Inter',Arial,sans-serif}}

/* Barra de formato */
.fmt{{
  flex-shrink:0;background:{T['topbar_bg']};
  border-bottom:1px solid {T['topbar_border']};
  padding:4px 10px;display:flex;align-items:center;gap:3px;
  flex-wrap:nowrap;overflow-x:auto;
}}
.fg{{display:flex;align-items:center;gap:2px;padding-right:6px;
     border-right:1px solid {T['panel_border']};flex-shrink:0}}
.fg:last-child{{border-right:none}}
.fb{{background:none;border:1px solid transparent;color:{T['text_secondary']};
    font-size:11px;padding:3px 5px;border-radius:3px;cursor:pointer;
    transition:all .1s;min-width:22px;text-align:center;line-height:1;
    font-family:'Inter',Arial,sans-serif}}
.fb:hover{{background:{T['btn_bg']};border-color:{T['btn_border']};color:{T['accent_light']}}}
.fb.on{{background:{T['btn_bg']};border-color:{T['accent']};color:{T['accent_light']}}}
.fs{{background:{T['input_bg']};border:1px solid {T['input_border']};
    color:{T['input_color']};font-size:10px;font-family:'IBM Plex Mono',monospace;
    padding:2px 4px;border-radius:3px;outline:none;appearance:none;cursor:pointer}}
.cd{{width:13px;height:13px;border-radius:50%;cursor:pointer;
    border:2px solid transparent;transition:border-color .1s;flex-shrink:0}}
.cd:hover,.cd.on{{border-color:{T['accent_light']}}}
.fl{{font-size:9px;color:{T['text_muted']};font-family:'IBM Plex Mono',monospace;white-space:nowrap}}

/* Zona edición */
.ew{{flex:1;overflow-y:auto;padding:12px 16px;min-height:0;
    background:{T['panel_bg']}}}
.doc{{
  min-height:100%;padding:20px 28px;outline:none;
  font-family:Arial,sans-serif;font-size:13px;line-height:1.75;
  color:#1a1a1a;background:#ffffff;
  border-radius:4px;transition:background .2s,color .2s;
}}
.doc li{{margin-left:18px;margin-bottom:2px}}
.doc b,.doc strong{{font-weight:700}}
.doc hr{{border:none;border-top:1px solid #ccc;margin:8px 0}}
.doc table{{border-collapse:collapse;width:100%;margin:8px 0;font-size:12px}}
.doc table td,.doc table th{{border:1px solid #ccc;padding:4px 8px;text-align:left}}
.doc table th{{background:#f0f4f8;font-weight:600}}

/* Action strip */
.as{{flex-shrink:0;background:{T['topbar_bg']};border-top:1px solid {T['topbar_border']};
    padding:6px 10px;display:flex;align-items:center;gap:5px;flex-wrap:wrap}}
.ab{{background:{T['btn_bg']};border:1px solid {T['btn_border']};color:{T['btn_color']};
    font-size:10px;font-family:'IBM Plex Mono',monospace;
    padding:4px 9px;border-radius:4px;cursor:pointer;
    transition:all .1s;display:flex;align-items:center;gap:3px}}
.ab:hover{{color:{T['accent_light']};border-color:{T['accent']}}}
.ab.p{{color:{T['accent_light']};border-color:{T['accent']};background:{T['btn_bg']}}}
.cw{{margin-left:auto;display:flex;align-items:center;gap:5px}}
.cbg{{width:60px;height:2px;background:{T['panel_border']};border-radius:1px;overflow:hidden}}
.cbf{{height:100%;background:{T['accent']};border-radius:1px;transition:width .4s}}
.cp{{font-size:9px;color:{T['text_secondary']};font-family:'IBM Plex Mono',monospace}}
</style>
</head>
<body>

<div class="fmt">
  <div class="fg">
    <select class="fs" id="fnt" onchange="applyFont(this.value)" style="width:82px">
      <option value="Arial" selected>Arial</option>
      <option value="'Courier New'">Courier New</option>
      <option value="Georgia">Georgia</option>
      <option value="'Times New Roman'">Times New Roman</option>
      <option value="'IBM Plex Mono'">Mono</option>
    </select>
    <select class="fs" id="sz" onchange="applySize(this.value)" style="width:40px">
      <option value="9">9</option><option value="10">10</option>
      <option value="11">11</option><option value="12">12</option>
      <option value="13" selected>13</option><option value="14">14</option>
      <option value="15">15</option><option value="16">16</option><option value="18">18</option>
    </select>
  </div>
  <div class="fg">
    <button class="fb" id="bB" onclick="fmt('bold')" title="Negrita"><b>B</b></button>
    <button class="fb" id="bI" onclick="fmt('italic')" title="Cursiva"><i>I</i></button>
    <button class="fb" id="bU" onclick="fmt('underline')" title="Subrayado"><u>U</u></button>
  </div>
  <div class="fg">
    <button class="fb" onclick="fmt('justifyLeft')" title="Izquierda"><i class="ti ti-align-left"></i></button>
    <button class="fb" onclick="fmt('justifyCenter')" title="Centro"><i class="ti ti-align-center"></i></button>
    <button class="fb" onclick="fmt('justifyRight')" title="Derecha"><i class="ti ti-align-right"></i></button>
    <button class="fb" onclick="fmt('justifyFull')" title="Justificado"><i class="ti ti-align-justified"></i></button>
  </div>
  <div class="fg">
    <button class="fb" onclick="fmt('insertUnorderedList')" title="Viñetas"><i class="ti ti-list"></i></button>
    <button class="fb" onclick="fmt('insertOrderedList')" title="Numerada"><i class="ti ti-list-numbers"></i></button>
    <button class="fb" onclick="insHR()" title="Separador">—</button>
  </div>
  <div class="fg" style="align-items:center;gap:4px">
    <span class="fl">Fondo:</span>
    <div class="cd on" style="background:#ffffff;border-color:{T['accent_light']}" onclick="setBg(this,'#ffffff','#1a1a1a')" title="Blanco"></div>
    <div class="cd" style="background:#0a1018" onclick="setBg(this,'#0a1018','#d0e4f0')" title="Clínico"></div>
    <div class="cd" style="background:#f5f0e8" onclick="setBg(this,'#f5f0e8','#2a1a0a')" title="Pergamino"></div>
    <div class="cd" style="background:#f0f4f8" onclick="setBg(this,'#f0f4f8','#1a2a3a')" title="Gris suave"></div>
    <div class="cd" style="background:#000409" onclick="setBg(this,'#000409','#e8f4ff')" title="Contraste"></div>
  </div>
  <div class="fg">
    <button class="fb" onclick="copyClean()" title="Copiar texto"><i class="ti ti-copy"></i></button>
  </div>
</div>

<div class="ew">
  <div class="doc" id="doc" contenteditable="true" spellcheck="false">{contenido_inicial}</div>
</div>

<div class="as">
  <button class="ab" onclick="optimize()"><i class="ti ti-wand" style="font-size:11px"></i> Optimizar conclusión</button>
  <button class="ab" onclick="getDefs()"><i class="ti ti-book" style="font-size:11px"></i> Definiciones</button>
  <button class="ab p" onclick="exportDoc()"><i class="ti ti-download" style="font-size:11px"></i> Exportar .docx</button>
  <div class="cw">
    <div class="cbg"><div class="cbf" id="cbf" style="width:0%"></div></div>
    <span class="cp" id="cp">0%</span>
  </div>
</div>

<script>
var doc=document.getElementById('doc');
doc.style.background='#ffffff';
doc.style.color='#1a1a1a';

function fmt(c){{doc.focus();document.execCommand(c,false,null);updS();}}
function updS(){{
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
  document.execCommand('insertHTML',false,'<hr style="border:none;border-top:1px solid #ccc;margin:8px 0"><br>');
}}
function calcS(){{
  var t=doc.innerText.toUpperCase();
  var f=['TÉCNICA','HALLAZGOS','IMPRESIÓN'].filter(function(s){{return t.includes(s);}}).length;
  var w=t.split(/\\s+/).filter(Boolean).length;
  return Math.min(100,Math.round((f/3)*60+Math.min(w/150,1)*40));
}}
function updBar(){{
  var s=calcS();
  document.getElementById('cbf').style.width=s+'%';
  document.getElementById('cp').textContent=s+'%';
}}
doc.addEventListener('input',updBar);
doc.addEventListener('keyup',updS);
doc.addEventListener('mouseup',updS);
window.addEventListener('load',function(){{updBar();}});

function copyClean(){{
  var text=doc.innerText;
  if(navigator.clipboard&&navigator.clipboard.writeText){{
    navigator.clipboard.writeText(text).then(function(){{toast('Copiado ✓');}});
  }}else{{
    var ta=document.createElement('textarea');
    ta.value=text;ta.style.cssText='position:fixed;opacity:0';
    document.body.appendChild(ta);ta.select();
    document.execCommand('copy');document.body.removeChild(ta);
    toast('Copiado ✓');
  }}
}}
function toast(msg){{
  var t=document.createElement('div');
  t.textContent=msg;
  t.style.cssText='position:fixed;bottom:56px;left:50%;transform:translateX(-50%);'
    +'background:{T["accent_glow"]};color:{T["accent_light"]};border:1px solid {T["accent"]};'
    +'padding:5px 14px;border-radius:4px;font-size:10px;font-family:IBM Plex Mono,monospace;z-index:9999;pointer-events:none';
  document.body.appendChild(t);
  setTimeout(function(){{document.body.removeChild(t);}},1800);
}}
function optimize(){{window.parent.postMessage({{type:'optimize',content:doc.innerText}},'*');}}
function getDefs(){{window.parent.postMessage({{type:'definiciones',content:doc.innerText}},'*');}}
function exportDoc(){{window.parent.postMessage({{type:'export',content:doc.innerText,html:doc.innerHTML}},'*');}}
</script>
</body>
</html>"""

    components.html(editor_html, height=iframe_h, scrolling=False)

    # ── Acciones IA ──
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1.4, 1.8, 1])

    with c1:
        if st.button("⟡ Optimizar conclusión", use_container_width=True):
            if api_key and st.session_state.reporte_texto:
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                with st.spinner("Refinando impresión..."):
                    try:
                        prompt_ref = f"""
Eres el optimizador diagnóstico de Beam AI. Mejora ÚNICAMENTE el bloque IMPRESIÓN DIAGNÓSTICA.

{REGLAS_CLINICAS}

CRITERIOS ADICIONALES:
- Más concisa, morfológicamente precisa y clínicamente accionable.
- Incluye solo clasificaciones directamente respaldadas (con el criterio que las justifica).
- Usa "•" para viñetas. Lenguaje sugerente para seguimiento.
- Devuelve el informe COMPLETO. Conserva Técnica y Hallazgos intactos.
- CERO asteriscos. Títulos en MAYÚSCULAS.

REPORTE:
{st.session_state.reporte_texto}
"""
                        res = client.chat.completions.create(
                            model="deepseek-chat",
                            messages=[{"role": "user", "content": prompt_ref}],
                            temperature=0.2
                        )
                        texto = res.choices[0].message.content
                        st.session_state.reporte_texto = texto
                        st.session_state.reporte_html = texto_a_html(texto)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    with c2:
        if st.button("✦ Definiciones y clasificaciones", use_container_width=True):
            if api_key and st.session_state.reporte_texto:
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                with st.spinner("Analizando y generando referencias..."):
                    try:
                        prompt_def = f"""
Eres un radiólogo experto. Analiza el informe y responde con este formato EXACTO.
Usa líneas simples, SIN espacios en blanco extra entre elementos de la misma sección.

CLASIFICACIONES USADAS
[Para cada clasificación encontrada:]
· Nombre: [nombre completo con autor/sociedad]
· Grado asignado: [grado] — [significado clínico en 1 línea]
· Hallazgo que lo justifica: [cita del texto]
· Referencia: [Autor, año, revista/sociedad]
· URL: [PubMed, ACR, RSNA o sociedad correspondiente]

CLASIFICACIONES SUGERIDAS
[Solo si hay hallazgos que las justifiquen EXPLÍCITAMENTE en el texto. Si no hay, escribe "Ninguna adicional justificada por los hallazgos descritos."]
· Nombre: [clasificación]
· Hallazgo que la justifica: [hallazgo específico del informe]
· Referencia: [Autor, año]
· URL: [URL]

DEFINICIONES OPERATIVAS
[Solo los 3-4 términos más relevantes, en líneas compactas:]
· [Término]: [definición morfológica en 1-2 líneas]

CORRELACIÓN CLÍNICA
[2-3 líneas. Impacto clínico y recomendación de manejo. Lenguaje sugerente, no prescriptivo.]

FORMATO ESTRICTO: SIN asteriscos. SIN líneas en blanco entre items de la misma sección. Solo una línea en blanco entre secciones principales.

INFORME:
{st.session_state.reporte_texto}
"""
                        res = client.chat.completions.create(
                            model="deepseek-chat",
                            messages=[{"role": "user", "content": prompt_def}],
                            temperature=0.15
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

    # ── Definiciones ──
    if st.session_state.definiciones_resultado:
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        with st.expander("⬡  DEFINICIONES · CLASIFICACIONES · REFERENCIAS", expanded=True):
            st.markdown(
                f'<div class="defs-box">{st.session_state.definiciones_resultado}</div>',
                unsafe_allow_html=True
            )
            if st.button("✕ Cerrar"):
                st.session_state.definiciones_resultado = ""
                st.rerun()
