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
    page_title="Beam AI | PACS Editor v4",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# CONSTANTES
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
    ],
    "Tomografía Computarizada": [
        "Hounsfield: hueso ~700 UH",
        "Adenopatía >1 cm eje corto",
        "Nódulo Fleischner >6 mm sólido",
        "WELLS alta probabilidad TEP",
        "ASPECTS ACV isquémico",
    ],
    "Radiografía": [
        "Kellgren-Lawrence I-IV artrosis",
        "Cobb >10° escoliosis",
        "Índice cardiotorácico >0.5",
    ],
    "Ultrasonido": [
        "TIRADS 4 → considerar BAAF",
        "BI-RADS 4B → biopsia indicada",
        "Murphy positivo colecistitis",
    ],
    "PET-CT": [
        "SUVmax >2.5 actividad metabólica",
        "LI-RADS 5 → HCC definitivo",
        "Respuesta PERCIST criterios",
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
    "definiciones_resultado": "",
    "editor_height": 600,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ==========================================
# HELPERS
# ==========================================
def leer_plantilla(file):
    doc = Document(file)
    secciones = []
    tabla_count = 0
    for element in doc.element.body:
        tag = element.tag.split('}')[-1]
        if tag == 'p':
            from docx.oxml.ns import qn
            import docx
            para = docx.text.paragraph.Paragraph(element, doc)
            texto = para.text.strip()
            if texto:
                secciones.append(texto)
        elif tag == 'tbl':
            tabla_count += 1
            from docx.table import Table
            tabla = Table(element, doc)
            filas_txt = []
            for row in tabla.rows:
                celdas = [c.text.strip() for c in row.cells]
                filas_txt.append(" | ".join(celdas))
            secciones.append(f"[TABLA {tabla_count}]\n" + "\n".join(filas_txt) + "\n[/TABLA]")
    return "\n".join(secciones)

def texto_a_html(texto):
    """Convierte texto plano del modelo a HTML para el editor."""
    lines = []
    for line in texto.split("\n"):
        s = line.strip()
        if not s:
            lines.append("<br>")
        elif s.isupper() and len(s) < 70:
            lines.append(f"<b>{s}</b><br>")
        elif s.startswith("•"):
            lines.append(f"<li>{s[1:].strip()}</li>")
        elif s.startswith("|") and "|" in s[1:]:
            # Línea de tabla markdown → fila HTML
            celdas = [c.strip() for c in s.strip("|").split("|")]
            lines.append("<tr>" + "".join(f"<td>{c}</td>" for c in celdas) + "</tr>")
        else:
            lines.append(f"{s}<br>")
    # Agrupar filas de tabla
    html = "\n".join(lines)
    if "<tr>" in html:
        html = html.replace("<tr>", "", 1)
        import re
        html = re.sub(r"(<tr>.*?</tr>\n?)+", lambda m: f'<table style="border-collapse:collapse;width:100%;margin:8px 0">' + m.group(0) + '</table>', html, flags=re.DOTALL)
    return html

def generar_docx(html_texto):
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
            self.in_table = False
            self.table_rows = []
            self.current_row = []
            self.current_cell = ""

        def handle_starttag(self, tag, attrs):
            if tag in ("b", "strong"):   self.bold = True
            elif tag in ("i", "em"):     self.italic = True
            elif tag == "u":             self.underline = True
            elif tag in ("p", "div"):    self.current_para = self.doc.add_paragraph()
            elif tag == "br":
                if not self.current_para:
                    self.current_para = self.doc.add_paragraph()
            elif tag == "li":
                self.current_para = self.doc.add_paragraph(style="List Bullet")
            elif tag == "table":
                self.in_table = True
                self.table_rows = []
            elif tag == "tr":
                self.current_row = []
            elif tag in ("td", "th"):
                self.current_cell = ""
            elif tag == "hr":
                self.doc.add_paragraph()

        def handle_endtag(self, tag):
            if tag in ("b", "strong"):   self.bold = False
            elif tag in ("i", "em"):     self.italic = False
            elif tag == "u":             self.underline = False
            elif tag in ("td", "th"):
                self.current_row.append(self.current_cell)
                self.current_cell = ""
            elif tag == "tr":
                self.table_rows.append(self.current_row)
            elif tag == "table":
                self.in_table = False
                if self.table_rows:
                    cols = max(len(r) for r in self.table_rows)
                    tbl = self.doc.add_table(rows=len(self.table_rows), cols=cols)
                    tbl.style = "Table Grid"
                    for i, row in enumerate(self.table_rows):
                        for j, cell_text in enumerate(row):
                            if j < cols:
                                tbl.rows[i].cells[j].text = cell_text
                self.table_rows = []

        def handle_data(self, data):
            text = data.strip()
            if not text:
                return
            if self.in_table:
                self.current_cell += text
                return
            if self.current_para is None:
                self.current_para = self.doc.add_paragraph()
            run = self.current_para.add_run(text + " ")
            run.bold = self.bold
            run.italic = self.italic
            run.underline = self.underline

    clean = html_texto.replace("\n", " ").strip()
    parser = HTMLtoDocx()
    try:
        parser.feed(clean)
    except Exception:
        import re
        doc = Document()
        plain = re.sub(r"<[^>]+>", "", html_texto)
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
# PROMPT PRINCIPAL — reglas clínicas mejoradas
# ==========================================
REGLAS_CLINICAS = """
REGLAS CLÍNICAS ESTRICTAS — NUNCA VIOLAR:

1. TERMINOLOGÍA PRECISA:
   - NO uses "cambios degenerativos" como término genérico. En su lugar especifica:
     el hallazgo anatomopatológico real: osteofitos marginales, esclerosis subcondral,
     disminución del espacio articular, condromalacia, fibrosis, etc.
   - NO uses "cambios crónicos" sin especificar el sustrato morfológico.
   - USA términos descriptivos y morfológicos: "osteofitos marginales tibiofemorales",
     "esclerosis subcondral en platillo medial", "pinzamiento articular de X mm".

2. TABLAS DE MEDIDAS:
   - Si la plantilla incluye tablas (marcadas como [TABLA]), RESPÉTALAS y complétalas
     con los valores del dictado. Devuelve la tabla en formato Markdown (| col | col |).
   - Si no hay plantilla pero el dictado menciona medidas, agrúpalas en una tabla
     al final de la sección HALLAZGOS.

3. CLASIFICACIONES:
   - Solo incluye clasificaciones que apliquen directamente a los hallazgos descritos.
   - No inventes grados si no tienes la información suficiente del dictado.
   - Cuando uses una clasificación, especifica el criterio que justifica ese grado.

4. IMPRESIÓN DIAGNÓSTICA:
   - Diagnósticos específicos, no genéricos.
   - Correlación anatómica-funcional cuando aplique.
   - Si el hallazgo requiere seguimiento o manejo, mencionarlo con lenguaje sugerente
     (no prescriptivo): "se sugiere correlación clínica", "puede valorarse".
"""

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

.beam-topbar {
    background: #0b0f1a; border-bottom: 1px solid #1a2333;
    padding: 10px 24px; display: flex; align-items: center; gap: 16px;
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

.plabel {
    font-size: 9px !important; letter-spacing: .2em !important; color: #2a4a6a !important;
    text-transform: uppercase !important; font-family: 'IBM Plex Mono', monospace !important;
    margin-bottom: 3px !important; margin-top: 0 !important; display: block;
}

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

.btn-main > div > button {
    background: #1a3a60 !important; border: 1px solid #2a5a90 !important;
    color: #7ac2f0 !important; font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important; font-size: 13px !important; letter-spacing: .06em !important;
    border-radius: 8px !important; padding: .75rem 1rem !important; width: 100% !important;
}
.btn-main > div > button:hover { background: #1f4878 !important; }

.stButton > button {
    background: #0d1828 !important; border: 1px solid #1a3050 !important;
    color: #3a6a9f !important; font-family: 'IBM Plex Mono', monospace !important;
    font-size: 11px !important; border-radius: 6px !important;
}
.stButton > button:hover { background: #102030 !important; color: #7ac2f0 !important; }

[data-testid="stExpander"] {
    background: #0b0f1a !important; border: 1px solid #1a2840 !important;
    border-radius: 7px !important; margin-bottom: 3px !important;
}
[data-testid="stExpander"] summary {
    color: #3a6a9f !important; font-family: 'IBM Plex Mono', monospace !important; font-size: 11px !important;
}
[data-testid="stExpander"] summary:hover { color: #7ac2f0 !important; }

[data-testid="stDownloadButton"] > button {
    background: #0d1a28 !important; border: 1px solid #1a3a58 !important;
    color: #3a8abf !important; font-family: 'IBM Plex Mono', monospace !important; font-size: 11px !important;
    border-radius: 6px !important;
}
[data-testid="stDownloadButton"] > button:hover { color: #7ac2f0 !important; }

/* Slider de altura */
[data-testid="stSlider"] > div { padding: 0 !important; }
[data-testid="stSlider"] .stSlider > div > div { background: #1a3a58 !important; }

.cbar-wrap { display: flex; align-items: center; gap: 8px; margin-top: 6px; }
.cbar-bg { flex: 1; height: 3px; background: #0f1828; border-radius: 2px; overflow: hidden; }
.cbar-fill { height: 100%; background: #1a4a80; border-radius: 2px; }
.cpct { font-size: 10px; color: #2a5a9a; font-family: 'IBM Plex Mono', monospace; }

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
    <span class="tbadge">v4.0 · PACS Editor · MSK / Neuro / TX</span>
    <div class="tstat"><span class="sdot"></span> DeepSeek · activo</div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# LAYOUT
# ==========================================
col_izq, col_centro = st.columns([1, 2.6], gap="small")

# ─────────────────────────────────────────
# PANEL IZQUIERDO
# ─────────────────────────────────────────
with col_izq:
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    if not api_key:
        api_key = st.text_input("API Key", type="password",
                                label_visibility="collapsed",
                                placeholder="sk-... DeepSeek API Key")

    with st.expander("⊞  MODALIDAD & REGIÓN", expanded=True):
        st.markdown('<span class="plabel">MODALIDAD</span>', unsafe_allow_html=True)
        modalidad = st.selectbox("Modalidad", list(SUGERENCIAS.keys()), label_visibility="collapsed")
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
                               placeholder="Dictado o escritura manual...\n\nEj: Desgarro horizontal menisco medial grado III Stoller, extrusión 3 mm, osteofitos marginales tibiales...")

    with st.expander("⊞  CONFIGURACIÓN", expanded=False):
        st.markdown('<span class="plabel">PLANTILLA BASE (.docx)</span>', unsafe_allow_html=True)
        archivo_base = st.file_uploader("Plantilla", type=["docx"], label_visibility="collapsed")
        plantilla_txt = leer_plantilla(archivo_base) if archivo_base else ""
        if plantilla_txt:
            tiene_tabla = "[TABLA]" in plantilla_txt
            st.markdown(
                f'<span style="font-size:10px;color:{"#2ecc71" if tiene_tabla else "#3a6a9f"};font-family:\'IBM Plex Mono\',monospace;">'
                f'{"✓ Plantilla con tablas detectada" if tiene_tabla else "Plantilla cargada"}</span>',
                unsafe_allow_html=True
            )

        st.markdown('<span class="plabel">DIRECTRICES DE ESTILO</span>', unsafe_allow_html=True)
        instrucciones = st.text_area("Directrices", height=70, label_visibility="collapsed",
                                     value="Lenguaje médico experto. Sin asteriscos. Solo clasificaciones directamente respaldadas por los hallazgos.")

    # Control de altura del editor
    with st.expander("⊞  TAMAÑO DEL EDITOR", expanded=False):
        st.markdown('<span class="plabel">ALTURA DEL ÁREA DE INFORME</span>', unsafe_allow_html=True)
        nueva_altura = st.slider(
            "Altura", min_value=300, max_value=1200, value=st.session_state.editor_height,
            step=50, label_visibility="collapsed",
            help="Arrastra para ajustar el tamaño del editor de informe"
        )
        if nueva_altura != st.session_state.editor_height:
            st.session_state.editor_height = nueva_altura
            st.rerun()
        st.markdown(
            f'<span style="font-size:10px;color:#2a5a9a;font-family:\'IBM Plex Mono\',monospace;">{st.session_state.editor_height}px</span>',
            unsafe_allow_html=True
        )

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

        prompt = f"""
Eres Beam AI, asistente experto en interpretación radiológica.
Redacta un informe de {modalidad} para región: {region}.

{REGLAS_CLINICAS}

LÓGICA DE SÍNTESIS:
1. Agrupa e infiere diagnósticos desde los hallazgos del dictado.
2. Si recibes diagnósticos, expande los hallazgos anatómicos morfológicos esperados
   (sin usar "cambios degenerativos" — ver REGLAS CLÍNICAS).
3. Si la plantilla incluye tablas, complétalas con los valores del dictado.
4. La IMPRESIÓN DIAGNÓSTICA debe ser concisa, morfológica y clínicamente accionable.

FORMATO ESTRICTO:
- TÍTULOS DE SECCIÓN en MAYÚSCULAS
- SIN asteriscos ni markdown de negritas
- Usa "•" para viñetas en la impresión
- Si hay tabla en la plantilla, devuélvela en formato Markdown (| col | col |)
- Plantilla base a respetar: {plantilla_txt if plantilla_txt else "TÉCNICA / HALLAZGOS / IMPRESIÓN DIAGNÓSTICA"}

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
                st.error(f"Error de enlace: {e}")
    elif not api_key:
        st.warning("Ingresa tu API Key de DeepSeek.")
    else:
        st.warning("Ingresa dictado o descripción clínica.")

# ─────────────────────────────────────────
# PANEL CENTRAL — Editor
# ─────────────────────────────────────────
with col_centro:
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    contenido_inicial = st.session_state.reporte_html or """<b>RESONANCIA MAGNÉTICA DE RODILLA DERECHA</b><br>
<br>
<b>TÉCNICA</b><br>
Secuencias multiplanares en T1, DP con supresión grasa (DPFS), T2 y STIR en planos axial, coronal y sagital, sin contraste.<br>
<br>
<b>HALLAZGOS</b><br>
<br>
<b>MENISCOS</b><br>
Menisco medial: alteración de señal grado III de Stoller en cuerpo y cuerno posterior, compatible con desgarro horizontal. Extrusión de 3 mm en el plano coronal.<br>
Menisco lateral: morfología e intensidad de señal conservadas.<br>
<br>
<b>LIGAMENTOS</b><br>
Ligamento cruzado anterior con señal heterogénea en tercio proximal, compatible con lesión parcial grado I de Hope &amp; Feagin. LCP, LCM y LCL sin alteraciones.<br>
<br>
<b>CARTÍLAGO</b><br>
Adelgazamiento condral focal grado III de ICRS en platillo tibial medial, extensión de 12 mm en el plano coronal. Esclerosis subcondral y edema óseo subcondral reactivo asociado.<br>
<br>
<b>ESPACIO ARTICULAR</b><br>
Pinzamiento femorotibial medial de aproximadamente 3 mm. Osteofitos marginales en cóndilos femorales y platillos tibiales de predominio medial.<br>
<br>
<b>IMPRESIÓN DIAGNÓSTICA</b><br>
<li>Desgarro horizontal de menisco medial, grado III de Stoller, con extrusión de 3 mm — clínicamente significativa.</li>
<li>Lesión parcial del LCA, grado I de Hope &amp; Feagin.</li>
<li>Condropatía grado III ICRS en compartimento femorotibial medial con esclerosis subcondral reactiva. Hallazgos compatibles con gonartrosis grado II de Kellgren-Lawrence.</li>"""

    # iframe height desde session_state
    editor_h = st.session_state.editor_height
    iframe_h = editor_h + 100  # barra formato + action strip

    editor_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=Syne:wght@700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{height:{iframe_h}px;overflow:hidden;display:flex;flex-direction:column;background:#08090f;font-family:Arial,sans-serif}}

/* ── Barra de formato ── */
.fmt-bar{{
    flex-shrink:0;background:#0b0f1a;border-bottom:1px solid #1a2333;
    padding:5px 10px;display:flex;align-items:center;gap:4px;flex-wrap:nowrap;overflow-x:auto;
}}
.fbg{{display:flex;align-items:center;gap:2px;padding-right:6px;border-right:1px solid #1a2333;flex-shrink:0}}
.fbg:last-child{{border-right:none}}
.fb{{background:none;border:1px solid transparent;color:#3a6a9f;font-size:12px;
     padding:3px 6px;border-radius:4px;cursor:pointer;transition:all .12s;min-width:24px;text-align:center;line-height:1}}
.fb:hover{{background:#0d1828;border-color:#1a3050;color:#7ac2f0}}
.fb.on{{background:#0f1e30;border-color:#1a3a58;color:#7ac2f0}}
.fsel{{background:#0f1520;border:1px solid #1a2f48;color:#6a9abf;
       font-size:11px;font-family:'IBM Plex Mono',monospace;
       padding:3px 4px;border-radius:4px;outline:none;appearance:none;cursor:pointer}}
.cdot{{width:15px;height:15px;border-radius:50%;cursor:pointer;
       border:2px solid transparent;transition:border-color .12s;flex-shrink:0}}
.cdot:hover,.cdot.on{{border-color:#3b8bd4}}
.flbl{{font-size:9px;color:#2a4060;font-family:'IBM Plex Mono',monospace;white-space:nowrap}}

/* ── Zona de edición ── */
.ew{{flex:1;overflow-y:auto;padding:14px 18px;min-height:0}}
.doc{{
    min-height:100%;padding:22px 30px;
    font-family:Arial,sans-serif;font-size:14px;line-height:1.85;
    color:#1a1a1a;outline:none;border-radius:6px;
    background:#ffffff;transition:background .25s,color .25s;
}}
.doc li{{margin-left:20px;margin-bottom:3px}}
.doc b,.doc strong{{font-weight:700}}
.doc hr{{border:none;border-top:1px solid #bbb;margin:10px 0}}
.doc table{{border-collapse:collapse;width:100%;margin:10px 0;font-size:13px}}
.doc table td,.doc table th{{border:1px solid #ccc;padding:5px 10px;text-align:left}}
.doc table th{{background:#f0f4f8;font-weight:600}}

/* ── Action strip ── */
.as{{flex-shrink:0;background:#0a0d16;border-top:1px solid #1a2333;
    padding:7px 12px;display:flex;align-items:center;gap:6px;flex-wrap:wrap}}
.ab{{background:#0d1828;border:1px solid #1a3050;color:#3a6a9f;
    font-size:11px;font-family:'IBM Plex Mono',monospace;
    padding:5px 10px;border-radius:5px;cursor:pointer;
    transition:all .12s;display:flex;align-items:center;gap:4px}}
.ab:hover{{color:#7ac2f0;background:#102030;border-color:#2a4a6a}}
.ab.prime{{color:#6ab8f0;border-color:#1a3a58;background:#0e1e30}}
.cbw{{margin-left:auto;display:flex;align-items:center;gap:6px}}
.cbg{{width:64px;height:3px;background:#0f1828;border-radius:2px;overflow:hidden}}
.cbf{{height:100%;background:#1a4a80;border-radius:2px;transition:width .4s}}
.cpct{{font-size:10px;color:#2a5a9a;font-family:'IBM Plex Mono',monospace}}
</style>
</head>
<body>

<div class="fmt-bar">
  <div class="fbg">
    <select class="fsel" id="fnt" onchange="applyFont(this.value)" style="width:88px">
      <option value="Arial" selected>Arial</option>
      <option value="'Courier New'">Courier New</option>
      <option value="Georgia">Georgia</option>
      <option value="'Times New Roman'">Times New Roman</option>
      <option value="'IBM Plex Mono'">Mono</option>
    </select>
    <select class="fsel" id="sz" onchange="applySize(this.value)" style="width:42px">
      <option>11</option><option>12</option><option>13</option>
      <option selected>14</option><option>15</option><option>16</option><option>18</option>
    </select>
  </div>
  <div class="fbg">
    <button class="fb" id="bB" onclick="fmt('bold')" title="Negrita"><b>B</b></button>
    <button class="fb" id="bI" onclick="fmt('italic')" title="Cursiva"><i>I</i></button>
    <button class="fb" id="bU" onclick="fmt('underline')" title="Subrayado"><u>U</u></button>
  </div>
  <div class="fbg">
    <button class="fb" onclick="fmt('justifyLeft')" title="Izquierda"><i class="ti ti-align-left"></i></button>
    <button class="fb" onclick="fmt('justifyCenter')" title="Centro"><i class="ti ti-align-center"></i></button>
    <button class="fb" onclick="fmt('justifyRight')" title="Derecha"><i class="ti ti-align-right"></i></button>
    <button class="fb" onclick="fmt('justifyFull')" title="Justificado"><i class="ti ti-align-justified"></i></button>
  </div>
  <div class="fbg">
    <button class="fb" onclick="fmt('insertUnorderedList')" title="Viñetas"><i class="ti ti-list"></i></button>
    <button class="fb" onclick="fmt('insertOrderedList')" title="Numerada"><i class="ti ti-list-numbers"></i></button>
    <button class="fb" onclick="insHR()" title="Separador"><i class="ti ti-minus"></i></button>
  </div>
  <div class="fbg" style="align-items:center;gap:4px">
    <span class="flbl">Fondo:</span>
    <div class="cdot on" style="background:#ffffff;border:1px solid #aaa" onclick="setBg(this,'#ffffff','#1a1a1a')" title="Blanco"></div>
    <div class="cdot" style="background:#0a1018" onclick="setBg(this,'#0a1018','#d0e4f0')" title="Clínico"></div>
    <div class="cdot" style="background:#f5f0e8;border:1px solid #ccc" onclick="setBg(this,'#f5f0e8','#2a1a0a')" title="Pergamino"></div>
    <div class="cdot" style="background:#000409" onclick="setBg(this,'#000409','#e8f4ff')" title="Contraste"></div>
    <div class="cdot" style="background:#0d1520" onclick="setBg(this,'#0d1520','#c8dff0')" title="Slate"></div>
  </div>
  <div class="fbg">
    <button class="fb" onclick="copyClean()" title="Copiar sin fondo"><i class="ti ti-copy"></i> Copiar</button>
  </div>
</div>

<div class="ew">
  <div class="doc" id="doc" contenteditable="true" spellcheck="false">
    {contenido_inicial}
  </div>
</div>

<div class="as">
  <button class="ab" onclick="optimize()"><i class="ti ti-wand" style="font-size:12px"></i> Optimizar conclusión</button>
  <button class="ab" onclick="getDefs()"><i class="ti ti-book" style="font-size:12px"></i> Definiciones operativas</button>
  <button class="ab prime" onclick="exportDoc()"><i class="ti ti-download" style="font-size:12px"></i> Exportar .docx</button>
  <div class="cbw">
    <div class="cbg"><div class="cbf" id="cbf" style="width:0%"></div></div>
    <span class="cpct" id="cpct">0%</span>
  </div>
</div>

<script>
var doc=document.getElementById('doc');

// Fondo blanco por defecto al cargar
doc.style.background='#ffffff';
doc.style.color='#1a1a1a';

function fmt(cmd){{doc.focus();document.execCommand(cmd,false,null);updState();}}

function updState(){{
  ['Bold','Italic','Underline'].forEach(function(c){{
    var b=document.getElementById('b'+c[0]);
    if(b)b.classList.toggle('on',document.queryCommandState(c.toLowerCase()));
  }});
}}

function applyFont(f){{doc.style.fontFamily=f;}}
function applySize(s){{doc.style.fontSize=s+'px';}}

function setBg(el,bg,col){{
  doc.style.background=bg; doc.style.color=col;
  document.querySelectorAll('.cdot').forEach(function(d){{d.classList.remove('on');}});
  el.classList.add('on');
}}

function insHR(){{
  doc.focus();
  document.execCommand('insertHTML',false,'<hr style="border:none;border-top:1px solid #bbb;margin:10px 0"><br>');
}}

function calcScore(){{
  var t=doc.innerText.toUpperCase();
  var found=['TÉCNICA','HALLAZGOS','IMPRESIÓN'].filter(function(s){{return t.includes(s);}}).length;
  var words=t.split(/\\s+/).filter(Boolean).length;
  return Math.min(100,Math.round((found/3)*60+Math.min(words/150,1)*40));
}}
function updBar(){{
  var s=calcScore();
  document.getElementById('cbf').style.width=s+'%';
  document.getElementById('cpct').textContent=s+'%';
}}
doc.addEventListener('input',updBar);
doc.addEventListener('keyup',updState);
doc.addEventListener('mouseup',updState);
window.addEventListener('load',function(){{updBar();}});

// Copiar SIN fondo: selecciona el texto limpio y lo pone en portapapeles
function copyClean(){{
  var prevBg=doc.style.background;
  var prevColor=doc.style.color;
  // Crea blob temporal con texto plano para evitar arrastrar el background
  var text=doc.innerText;
  if(navigator.clipboard && navigator.clipboard.writeText){{
    navigator.clipboard.writeText(text).then(function(){{
      showToast('Texto copiado sin fondo ✓');
    }});
  }} else {{
    var ta=document.createElement('textarea');
    ta.value=text; ta.style.position='fixed'; ta.style.opacity='0';
    document.body.appendChild(ta); ta.select();
    document.execCommand('copy'); document.body.removeChild(ta);
    showToast('Texto copiado ✓');
  }}
}}

function showToast(msg){{
  var t=document.createElement('div');
  t.textContent=msg;
  t.style.cssText='position:fixed;bottom:60px;left:50%;transform:translateX(-50%);background:#1a3a60;color:#7ac2f0;border:1px solid #2a5a90;padding:6px 16px;border-radius:6px;font-size:11px;font-family:IBM Plex Mono,monospace;z-index:9999;pointer-events:none';
  document.body.appendChild(t);
  setTimeout(function(){{document.body.removeChild(t);}},2000);
}}

function getText(){{return doc.innerText;}}
function getHTML(){{return doc.innerHTML;}}

function optimize(){{
  window.parent.postMessage({{type:'optimize',content:getText()}},'*');
}}
function getDefs(){{
  window.parent.postMessage({{type:'definiciones',content:getText()}},'*');
}}
function exportDoc(){{
  window.parent.postMessage({{type:'export',content:getText(),html:getHTML()}},'*');
}}
</script>
</body>
</html>"""

    components.html(editor_html, height=iframe_h, scrolling=False)

    # ── Acciones IA ──
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1.4, 1.8, 1])

    with c1:
        if st.button("⟡ Optimizar conclusión", use_container_width=True):
            if api_key and st.session_state.reporte_texto:
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                with st.spinner("Refinando impresión diagnóstica..."):
                    try:
                        prompt_ref = f"""
Eres el optimizador diagnóstico de Beam AI.

TAREA: Mejora ÚNICAMENTE el bloque IMPRESIÓN DIAGNÓSTICA de este informe.

{REGLAS_CLINICAS}

ADICIONAL PARA LA IMPRESIÓN:
- Más elegante, concluyente y morfológicamente precisa.
- Incluye clasificaciones directamente respaldadas por los hallazgos (con el criterio que las justifica).
- Correlación clínico-funcional si aplica.
- Lenguaje sugerente para seguimiento: "se sugiere correlación clínica", "puede valorarse".
- Usa viñetas con "•"

REGLAS: Devuelve el informe COMPLETO. Conserva Técnica y Hallazgos intactos. CERO asteriscos. Títulos en MAYÚSCULAS.

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
                with st.spinner("Analizando clasificaciones y generando referencias..."):
                    try:
                        prompt_def = f"""
Eres un radiólogo experto y docente. Lee este informe radiológico y responde en este formato exacto:

CLASIFICACIONES USADAS
Para cada clasificación encontrada en el informe:
- Nombre de la clasificación: [nombre completo con autor/sociedad]
- Grado asignado: [grado] — Significado clínico: [qué implica morfológica y funcionalmente]
- Justificación en el texto: [cita el hallazgo concreto que sustenta ese grado]
- Referencia: [autor principal, año, revista o sociedad — ej: "Stoller DW, Radiology 1987" o "ACR 2017"]
- URL de referencia: [URL directa a PubMed, ACR, RSNA, o fuente oficial. Si no tienes URL exacta, da la búsqueda en PubMed o la web de la sociedad]

CLASIFICACIONES SUGERIDAS (SOLO SI HAY EVIDENCIA DIRECTA EN LOS HALLAZGOS)
Solo sugiere clasificaciones adicionales si el texto contiene hallazgos que las justifican explícitamente.
Para cada una:
- Nombre: [nombre de la clasificación]
- Por qué aplica: [hallazgo específico del informe que la justifica]
- Referencia: [autor, año, fuente]
- URL: [URL o búsqueda sugerida]

Si NO hay evidencia directa en el informe para una clasificación, NO la sugieras.

DEFINICIONES OPERATIVAS
Para los 3-5 términos técnicos más relevantes del informe:
- Término: [término]
- Definición: [1-2 líneas, morfológica y clínicamente precisa]

CORRELACIÓN CLÍNICA
- Impacto clínico esperado y recomendación de manejo en 2-3 líneas. Lenguaje sugerente, no prescriptivo.

FORMATO: SIN asteriscos. Títulos en MAYÚSCULAS. Claro y estructurado.

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

    # ── Panel de definiciones y clasificaciones ──
    if st.session_state.definiciones_resultado:
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        with st.expander("⬡  DEFINICIONES · CLASIFICACIONES · REFERENCIAS", expanded=True):
            st.markdown(
                f'<div style="font-family:\'IBM Plex Mono\',monospace;font-size:12px;color:#8ab0cc;'
                f'line-height:1.8;background:#0a1020;padding:16px;border-radius:7px;'
                f'border:1px solid #1a2840;white-space:pre-wrap">'
                f'{st.session_state.definiciones_resultado}</div>',
                unsafe_allow_html=True
            )
            if st.button("✕ Cerrar análisis"):
                st.session_state.definiciones_resultado = ""
                st.rerun()
