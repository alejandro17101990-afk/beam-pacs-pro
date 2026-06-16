"""
AURA — Radiology Copilot
UI inspired by Quillr: dark theme, left dictation panel, rich central editor, collapsible right AI panel.
"""

import streamlit as st
import streamlit.components.v1 as components
import speech_recognition as sr
import io, re, json, datetime
from openai import OpenAI
from docx import Document
from docx.shared import Pt

st.set_page_config(
    page_title="AURA · Radiology Copilot",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ══════════════════════════════════════════════════════════════════
# KNOWLEDGE BASE
# ══════════════════════════════════════════════════════════════════
KB_CLASIF = {
    "Menisco · Stoller": {
        "desc": "Clasificación RM de desgarro meniscal",
        "grados": {"I":"Señal focal intrameniscal, no articular.","II":"Señal lineal, no alcanza superficie articular.","III":"Alcanza superficie articular → DESGARRO."},
        "cmd": "/stoller",
    },
    "Cartílago · ICRS": {
        "desc": "International Cartilage Repair Society",
        "grados": {"I":"Fibrilación superficial, <50% grosor.","II":"Lesión hasta 50% grosor.","III":"Lesión >50%, sin exposición ósea.","IV":"Hueso subcondral expuesto."},
        "cmd": "/icrs",
    },
    "Artrosis · Kellgren-Lawrence": {
        "desc": "Clasificación radiográfica de osteoartrosis",
        "grados": {"0":"Sin cambios.","I":"Osteofito posible.","II":"Osteofito definido.","III":"Pinzamiento moderado.","IV":"Pinzamiento grave + deformidad."},
        "cmd": "/kl",
    },
    "LCA · Hope & Feagin": {
        "desc": "Lesión de ligamento cruzado anterior en RM",
        "grados": {"Parcial":"Fibras continuas, señal aumentada.","Completa":"Discontinuidad total.","Crónica":"Fibras atróficas."},
        "cmd": "/lca",
    },
    "Disco · Pfirrmann": {
        "desc": "Degeneración discal en columna",
        "grados": {"I":"Núcleo brillante homogéneo.","II":"Señal ligeramente reducida.","III":"Señal gris, distinción borrosa.","IV":"Señal baja, altura reducida.","V":"Espacio discal colapsado."},
        "cmd": "/pfirrmann",
    },
    "Modic": {
        "desc": "Cambios en placa terminal vertebral",
        "grados": {"I":"Edema/inflamación activa.","II":"Sustitución grasa.","III":"Esclerosis ósea."},
        "cmd": "/modic",
    },
    "ACR TIRADS": {
        "desc": "Nódulo tiroideo en ecografía",
        "grados": {"1":"Normal.","2":"Benigno.","3":"Levemente sospechoso.","4":"Moderadamente sospechoso. BAAF ≥1.5cm.","5":"Altamente sospechoso. BAAF ≥1cm."},
        "cmd": "/tirads",
    },
    "ACR BI-RADS": {
        "desc": "Hallazgos mamarios",
        "grados": {"0":"Evaluación incompleta.","1":"Negativo.","2":"Benigno.","3":"Probablemente benigno.","4":"Sospechoso — biopsia.","5":"Altamente maligno.","6":"Malignidad confirmada."},
        "cmd": "/birads",
    },
    "LI-RADS": {
        "desc": "Lesión hepática en paciente de alto riesgo",
        "grados": {"LR-1":"Definitivamente benigno.","LR-2":"Probablemente benigno.","LR-3":"Indeterminado.","LR-4":"Probablemente HCC.","LR-5":"Definitivamente HCC.","LR-M":"Maligno, no específico."},
        "cmd": "/lirads",
    },
    "Fleischner 2017": {
        "desc": "Seguimiento de nódulos pulmonares",
        "grados": {"<6mm bajo riesgo":"Sin seguimiento.","<6mm alto riesgo":"TC opcional 12m.","6-8mm":"TC 6-12m, luego 18-24m.",">8mm":"TC 3m / PET / biopsia.","Subsólido ≥6mm":"TC 6-12m para confirmar."},
        "cmd": "/fleischner",
    },
}

KB_REGIONES = {
    "Rodilla": {"clasif":["Menisco · Stoller","Cartílago · ICRS","Artrosis · Kellgren-Lawrence","LCA · Hope & Feagin"],"omisiones":["menisco medial","menisco lateral","LCA","cartílago"]},
    "Columna lumbar": {"clasif":["Disco · Pfirrmann","Modic"],"omisiones":["discos","canal espinal","foramina"]},
    "Columna cervical": {"clasif":["Disco · Pfirrmann","Modic"],"omisiones":["discos","médula","canal"]},
    "Hombro": {"clasif":["Cartílago · ICRS"],"omisiones":["supraespinoso","infraespinoso","bursa","labrum"]},
    "Cadera": {"clasif":["Cartílago · ICRS","Artrosis · Kellgren-Lawrence"],"omisiones":["labrum","cartílago","espacio articular"]},
    "Tobillo / Pie": {"clasif":["Cartílago · ICRS"],"omisiones":["Aquiles","ligamentos laterales","cartílago talar"]},
    "Cerebro": {"clasif":[],"omisiones":["parénquima","ventrículos","línea media","sustancia blanca"]},
    "Tórax": {"clasif":["Fleischner 2017"],"omisiones":["parénquima","pleuras","mediastino"]},
    "Abdomen / Pelvis": {"clasif":["LI-RADS"],"omisiones":["hígado","riñones","páncreas","aorta"]},
    "Mama": {"clasif":["ACR BI-RADS"],"omisiones":["composición","BI-RADS","ganglios axilares"]},
    "Tiroides": {"clasif":["ACR TIRADS"],"omisiones":["TIRADS","adenopatías"]},
    "Muñeca / Mano": {"clasif":["Cartílago · ICRS"],"omisiones":["TFCC","tendones"]},
}

MODALIDADES = ["Resonancia Magnética","Tomografía Computarizada","Radiografía","Ultrasonido","PET-CT","Intervencionismo"]
REGIONES    = sorted(KB_REGIONES.keys())

# ══════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════
_D = {
    "dictado":"","reporte_html":"","reporte_texto":"",
    "copilot_txt":"","copilot_tipo":"",
    "historial":[],"estilo_ejemplos":[],"estilo_pref":"",
    "clasif_activas":{},"api_key":"","proveedor":"deepseek",
    "modalidad":"Tomografía Computarizada","region":"Abdomen / Pelvis",
    "instrucciones":"Lenguaje médico experto en español de México. Impresión diagnóstica concluyente.",
    "estudio_id":f"EST-{datetime.datetime.now().strftime('%Y%m%d-%H%M')}",
    "plantilla_txt":"","qa":{},"right_open":True,
    "ai_toggle":True,
    "editor_content": "",
    "titulo_estudio":"TC ABDOMEN Y PELVIS CON CONTRASTE",
}
for k,v in _D.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ══════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════
def get_client():
    key = st.session_state.api_key
    try:
        if not key: key = st.secrets.get("deepseek_key","") or st.secrets.get("openai_key","")
    except: pass
    if not key: return None,None
    p = st.session_state.proveedor
    if p=="openai_mini": return OpenAI(api_key=key),"gpt-4o-mini"
    if p=="openai_4":    return OpenAI(api_key=key),"gpt-4.1-mini"
    return OpenAI(api_key=key, base_url="https://api.deepseek.com"),"deepseek-chat"

def transcribir(audio):
    r = sr.Recognizer()
    try:
        with sr.AudioFile(audio) as s: return r.recognize_google(r.record(s), language="es-MX")
    except: return ""

def leer_plantilla(f):
    doc = Document(f)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

def texto_a_html(txt):
    lines=[]
    for line in txt.split("\n"):
        s=line.strip()
        if not s: lines.append("<br>")
        elif s.isupper() and len(s)<80 and not s.startswith("•"):
            lines.append(f'<h3 class="sec">{s}</h3>')
        elif s.startswith("•") or s.startswith("-"):
            lines.append(f"<li>{s[1:].strip()}</li>")
        else: lines.append(f"<p>{s}</p>")
    return "\n".join(lines)

def calcular_qa(texto, region):
    up=texto.upper()
    secs={s:s in up for s in ["TÉCNICA","HALLAZGOS","IMPRESIÓN"]}
    words=len(texto.split())
    omis=[e for e in KB_REGIONES.get(region,{}).get("omisiones",[]) if e.lower() not in texto.lower()]
    found=sum(secs.values())
    score=min(100,int((found/3)*40+min(words/200,1)*35+(found==3)*25))
    return {"score":score,"secciones":secs,"omisiones":omis}

def build_system_prompt():
    r=st.session_state.region; m=st.session_state.modalidad
    clasif_info=""
    for c in KB_REGIONES.get(r,{}).get("clasif",[]):
        d=KB_CLASIF.get(c,{})
        clasif_info+=f"\n{c}: {d.get('desc','')}\n"
        for g,desc in d.get("grados",{}).items(): clasif_info+=f"  Gr.{g}: {desc}\n"
    clasif_activas=""
    if st.session_state.clasif_activas:
        clasif_activas="\nCLASIFICACIONES DEL CASO:\n"
        for k,v in st.session_state.clasif_activas.items(): clasif_activas+=f"• {k}: Grado {v}\n"
    estilo=""
    if st.session_state.estilo_ejemplos:
        estilo="\nESTILO DEL RADIÓLOGO:\n"+"".join(f"Ejemplo:\n{e[:500]}\n" for e in st.session_state.estilo_ejemplos[-3:])
    if st.session_state.estilo_pref: estilo+=f"\nPreferencias: {st.session_state.estilo_pref}"
    return f"""Eres AURA, copiloto de redacción radiológica de nivel experto.
MODALIDAD: {m} | REGIÓN: {r}
CLASIFICACIONES:{clasif_info}{clasif_activas}{estilo}
DIRECTRICES: {st.session_state.instrucciones}
PLANTILLA: {st.session_state.plantilla_txt or "TÉCNICA / HALLAZGOS / IMPRESIÓN DIAGNÓSTICA"}
REGLAS: Títulos en MAYÚSCULAS, sin asteriscos, redacción narrativa en hallazgos, impresión con viñetas "•", español de México."""

def generar_docx(html_txt, plain_txt):
    from html.parser import HTMLParser
    class P(HTMLParser):
        def __init__(self):
            super().__init__(); self.doc=Document()
            st2=self.doc.styles["Normal"]; st2.font.name="Arial"; st2.font.size=Pt(11)
            self.bold=self.italic=self.ul=False; self.para=None
        def handle_starttag(self,tag,attrs):
            if tag in("b","strong"): self.bold=True
            elif tag in("i","em"): self.italic=True
            elif tag=="u": self.ul=True
            elif tag in("p","div","h3"): self.para=self.doc.add_paragraph()
            elif tag=="li": self.para=self.doc.add_paragraph(style="List Bullet")
            elif tag=="br":
                if not self.para: self.para=self.doc.add_paragraph()
        def handle_endtag(self,tag):
            if tag in("b","strong"): self.bold=False
            elif tag in("i","em"): self.italic=False
            elif tag=="u": self.ul=False
        def handle_data(self,data):
            t=data.strip()
            if not t: return
            if not self.para: self.para=self.doc.add_paragraph()
            run=self.para.add_run(t+" ")
            run.bold=self.bold; run.italic=self.italic; run.underline=self.ul
    p=P()
    try: p.feed((html_txt or plain_txt).replace("\n"," "))
    except:
        d=Document()
        for line in re.sub(r"<[^>]+>","",plain_txt).split("\n"): d.add_paragraph(line)
        bio=io.BytesIO(); d.save(bio); return bio.getvalue()
    bio=io.BytesIO(); p.doc.save(bio); return bio.getvalue()

# ══════════════════════════════════════════════════════════════════
# DEFAULT EDITOR CONTENT
# ══════════════════════════════════════════════════════════════════
DEFAULT_CONTENT = """<h3 class="sec">HALLAZGOS</h3>
<p><strong>Hígado:</strong> De tamaño, forma y contornos normales, con atenuación homogénea. No se identifican lesiones focales. Vía biliar intra y extrahepática no dilatada.</p>
<p><strong>Vesícula biliar:</strong> De paredes delgadas, sin litiasis radiopacas.</p>
<p><strong>Páncreas:</strong> De aspecto normal, sin lesiones focales ni dilatación del conducto pancreático.</p>
<p><strong>Bazo:</strong> De tamaño normal, con atenuación homogénea, sin lesiones focales.</p>
<p><strong>Glándulas suprarrenales:</strong> De aspecto normal.</p>
<p><strong>Riñones:</strong> De tamaño y morfología conservada, con adecuada captación y eliminación del contraste. No se observan litiasis ni hidronefrosis.</p>
<p><strong>Estómago:</strong> Moderadamente distendido con contenido líquido. Asas intestinales permeables, sin dilatación ni engrosamiento parietal. Apéndice cecal de aspecto normal.</p>
<p><strong>Útero y anexos:</strong> De aspecto normal para la edad.</p>
<p>No se identifica líquido libre intraabdominal ni colecciones.</p>
<p><strong>Pared abdominal:</strong> Sin defectos. No se identifican lesiones óseas destructivas.</p>
<h3 class="sec">CONCLUSIÓN</h3>
<li>Estudio sin hallazgos tomográficos sugestivos de patología abdominal o pélvica aguda.</li>
<li>Hallazgos descritos.</li>"""

# ══════════════════════════════════════════════════════════════════
# GLOBAL CSS
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

/* ── RESET ── */
*, *::before, *::after { box-sizing: border-box; }
html, body, .stApp {
    background: #1C1C1E !important;
    font-family: 'Inter', sans-serif !important;
    color: #E5E5EA !important;
    margin: 0; padding: 0;
}
.block-container { padding: 0 !important; max-width: 100% !important; }
header, footer, [data-testid="stToolbar"],
[data-testid="stDecoration"], [data-testid="stStatusWidget"] { display: none !important; }
section[data-testid="stSidebar"] { display: none !important; }
[data-testid="stHorizontalBlock"] { gap: 0 !important; align-items: stretch !important; }
[data-testid="column"] { padding: 0 !important; }

/* ── STREAMLIT WIDGETS DARK ── */
.stTextArea textarea {
    background: #2C2C2E !important;
    border: 1px solid #3A3A3C !important;
    border-radius: 8px !important;
    color: #E5E5EA !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 13px !important;
    line-height: 1.65 !important;
    resize: none !important;
}
.stTextArea textarea:focus {
    border-color: #F5C518 !important;
    box-shadow: 0 0 0 3px rgba(245,197,24,.1) !important;
}
[data-testid="stSelectbox"] > div > div {
    background: #2C2C2E !important;
    border: 1px solid #3A3A3C !important;
    border-radius: 8px !important;
    color: #E5E5EA !important;
    font-size: 12px !important;
}
[data-testid="stTextInput"] input {
    background: #2C2C2E !important;
    border: 1px solid #3A3A3C !important;
    border-radius: 8px !important;
    color: #E5E5EA !important;
    font-size: 12px !important;
}
[data-testid="stAudioInput"] {
    background: #2C2C2E !important;
    border: 1px solid #3A3A3C !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploader"] {
    background: #2C2C2E !important;
    border: 1px dashed #3A3A3C !important;
    border-radius: 8px !important;
}
[data-testid="stFileUploader"] * { color: #636366 !important; font-size: 11px !important; }
[data-testid="stRadio"] > div { gap: 4px !important; }
[data-testid="stRadio"] label {
    background: #2C2C2E !important;
    border: 1px solid #3A3A3C !important;
    border-radius: 6px !important;
    padding: 4px 10px !important;
    font-size: 11px !important;
    color: #8E8E93 !important;
}
[data-testid="stRadio"] label:has(input:checked) {
    border-color: #F5C518 !important;
    color: #F5C518 !important;
    background: rgba(245,197,24,.08) !important;
}

/* ── BUTTONS ── */
.stButton > button {
    background: #2C2C2E !important;
    border: 1px solid #3A3A3C !important;
    color: #AEAEB2 !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 12px !important;
    font-weight: 500 !important;
    border-radius: 8px !important;
    padding: 6px 14px !important;
    transition: all .15s !important;
    white-space: nowrap !important;
}
.stButton > button:hover {
    background: #3A3A3C !important;
    color: #E5E5EA !important;
    border-color: #48484A !important;
}
.btn-yellow .stButton > button {
    background: #F5C518 !important;
    border: none !important;
    color: #1C1C1E !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    border-radius: 10px !important;
    padding: 14px !important;
    letter-spacing: .02em !important;
}
.btn-yellow .stButton > button:hover {
    background: #FFD700 !important;
    box-shadow: 0 4px 16px rgba(245,197,24,.35) !important;
}
.btn-update .stButton > button {
    background: #3A3A3C !important;
    border: none !important;
    color: #E5E5EA !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    border-radius: 10px !important;
    padding: 12px !important;
}
.btn-update .stButton > button:hover {
    background: #48484A !important;
}
.btn-primary .stButton > button {
    background: #F5C518 !important;
    border: none !important;
    color: #1C1C1E !important;
    font-weight: 700 !important;
    border-radius: 8px !important;
}
.btn-primary .stButton > button:hover {
    background: #FFD700 !important;
    box-shadow: 0 4px 12px rgba(245,197,24,.3) !important;
}
.btn-ghost .stButton > button {
    background: transparent !important;
    border: 1px solid transparent !important;
    color: #636366 !important;
    padding: 4px 8px !important;
}
.btn-ghost .stButton > button:hover {
    background: #2C2C2E !important;
    color: #AEAEB2 !important;
}
[data-testid="stDownloadButton"] > button {
    background: #2C2C2E !important;
    border: 1px solid #3A3A3C !important;
    color: #AEAEB2 !important;
    font-size: 12px !important;
    border-radius: 8px !important;
}
[data-testid="stDownloadButton"] > button:hover {
    border-color: #F5C518 !important;
    color: #F5C518 !important;
}

/* ── EXPANDER ── */
[data-testid="stExpander"] {
    background: transparent !important;
    border: none !important;
    border-bottom: 1px solid #2C2C2E !important;
    border-radius: 0 !important;
    margin: 0 !important;
}
[data-testid="stExpander"] summary {
    color: #636366 !important;
    font-size: 10px !important;
    font-weight: 700 !important;
    letter-spacing: .06em !important;
    text-transform: uppercase !important;
    padding: 10px 0 !important;
}
[data-testid="stExpander"] summary:hover { color: #AEAEB2 !important; }
[data-testid="stExpander"] > div > div { padding: 4px 0 10px !important; }

/* ── SCROLLBAR ── */
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #3A3A3C; border-radius: 2px; }
::-webkit-scrollbar-thumb:hover { background: #48484A; }

/* ── DIVIDER ── */
hr { border-color: #2C2C2E !important; margin: 8px 0 !important; }

/* ── LABEL ── */
.lbl {
    font-size: 10px; font-weight: 700; color: #636366;
    text-transform: uppercase; letter-spacing: .06em;
    display: block; margin-bottom: 4px; margin-top: 12px;
}

/* ── LEFT PANEL ── */
.left-panel {
    background: #141414;
    border-right: 1px solid #2C2C2E;
    height: 100vh;
    display: flex; flex-direction: column;
    overflow: hidden;
}
.lp-header {
    padding: 14px 14px 10px;
    border-bottom: 1px solid #2C2C2E;
    display: flex; align-items: center; gap: 10px;
}
.lp-logo {
    display: flex; align-items: center; gap: 8px;
}
.logo-leaf {
    font-size: 18px; line-height: 1;
}
.logo-text {
    font-size: 16px; font-weight: 700; color: #E5E5EA;
    letter-spacing: -.01em;
}
.lp-section {
    font-size: 10px; font-weight: 700; color: #48484A;
    letter-spacing: .1em; text-transform: uppercase;
    padding: 14px 14px 6px;
}
.lp-body { flex: 1; overflow-y: auto; padding: 0; }
.lp-footer {
    padding: 10px 14px 14px;
    border-top: 1px solid #2C2C2E;
}
.hist-item {
    padding: 8px 14px; cursor: pointer;
    transition: background .12s;
    border-left: 2px solid transparent;
}
.hist-item:hover { background: #1C1C1E; border-left-color: #F5C518; }
.hist-item.active { background: #1C1C1E; border-left-color: #F5C518; }
.hi-title { font-size: 12px; font-weight: 600; color: #C7C7CC; line-height: 1.3; }
.hi-meta { font-size: 10px; color: #48484A; margin-top: 1px; font-family: 'JetBrains Mono', monospace; }
.hi-badge {
    display: inline-block; padding: 1px 5px; border-radius: 3px;
    font-size: 9px; font-weight: 700; margin-top: 3px;
}
.badge-active { background: rgba(245,197,24,.12); color: #F5C518; }
.badge-ok { background: rgba(52,199,89,.1); color: #34C759; }

/* ── TOPBAR ── */
.aura-topbar {
    background: #1C1C1E;
    border-bottom: 1px solid #2C2C2E;
    height: 50px;
    padding: 0 18px;
    display: flex; align-items: center; gap: 10px;
    position: sticky; top: 0; z-index: 200;
}
.tb-study-select {
    background: #2C2C2E;
    border: 1px solid #3A3A3C;
    border-radius: 8px;
    color: #E5E5EA;
    font-size: 12px; font-weight: 500;
    padding: 6px 32px 6px 12px;
    outline: none; cursor: pointer;
    font-family: 'Inter', sans-serif;
    appearance: none;
    min-width: 160px;
}
.tb-template-btn {
    background: #2C2C2E; border: 1px solid #3A3A3C;
    color: #AEAEB2; font-size: 12px; font-weight: 500;
    padding: 6px 14px; border-radius: 8px; cursor: pointer;
    font-family: 'Inter', sans-serif; white-space: nowrap;
    transition: all .12s;
}
.tb-template-btn:hover { background: #3A3A3C; color: #E5E5EA; }
.tb-btn-icon {
    background: #2C2C2E; border: 1px solid #3A3A3C;
    color: #AEAEB2; font-size: 12px; font-weight: 500;
    padding: 6px 12px; border-radius: 8px; cursor: pointer;
    font-family: 'Inter', sans-serif; display: flex; align-items: center; gap: 6px;
    transition: all .12s; white-space: nowrap;
}
.tb-btn-icon:hover { background: #3A3A3C; color: #E5E5EA; }
.tb-new-btn {
    background: #2C2C2E; border: 1px solid #3A3A3C;
    color: #AEAEB2; font-size: 12px; font-weight: 600;
    padding: 6px 14px; border-radius: 8px; cursor: pointer;
    font-family: 'Inter', sans-serif; display: flex; align-items: center; gap: 6px;
    transition: all .12s; white-space: nowrap;
}
.tb-new-btn:hover { background: #3A3A3C; color: #E5E5EA; }
.tb-reports-btn {
    background: #2C2C2E; border: 1px solid #3A3A3C;
    color: #AEAEB2; font-size: 12px; font-weight: 600;
    padding: 6px 14px; border-radius: 8px; cursor: pointer;
    display: flex; align-items: center; gap: 6px;
    font-family: 'Inter', sans-serif; transition: all .12s;
}
.tb-reports-btn:hover { background: #3A3A3C; color: #E5E5EA; }
.tb-user {
    background: #2C2C2E; border: 1px solid #3A3A3C;
    color: #C7C7CC; font-size: 12px; font-weight: 500;
    padding: 6px 12px; border-radius: 8px;
    display: flex; align-items: center; gap: 6px;
    font-family: 'Inter', sans-serif; white-space: nowrap;
}
.tb-user-dot {
    width: 22px; height: 22px; border-radius: 50%;
    background: linear-gradient(135deg,#F5C518,#FF9500);
    display: flex; align-items: center; justify-content: center;
    font-size: 10px; font-weight: 700; color: #1C1C1E;
}
.tb-sep { width: 1px; height: 20px; background: #3A3A3C; flex-shrink: 0; }

/* ── RIGHT PANEL TAB TRIGGER ── */
.rp-trigger {
    position: fixed; right: 0; top: 50%;
    transform: translateY(-50%);
    background: #F5C518; color: #1C1C1E;
    width: 22px; padding: 12px 0;
    border-radius: 8px 0 0 8px;
    display: flex; align-items: center; justify-content: center;
    cursor: pointer; z-index: 500;
    font-size: 11px; font-weight: 700;
    box-shadow: -4px 0 12px rgba(0,0,0,.4);
    writing-mode: vertical-rl;
    letter-spacing: .1em;
}

/* ── COPILOT RESULT ── */
.copilot-box {
    background: #1C1C1E; border: 1px solid #2C2C2E;
    border-radius: 10px; padding: 14px; margin-bottom: 10px;
}
.copilot-box-title {
    font-size: 10px; font-weight: 700; color: #F5C518;
    text-transform: uppercase; letter-spacing: .08em;
    margin-bottom: 10px; display: flex; align-items: center; gap: 6px;
}
.copilot-box-title::before {
    content: ''; width: 6px; height: 6px; border-radius: 50%;
    background: #F5C518; display: inline-block; flex-shrink: 0;
}
.copilot-text {
    font-size: 11px; color: #AEAEB2; line-height: 1.75;
    white-space: pre-wrap; font-family: 'Inter', sans-serif;
}

/* ── QA SCORE ── */
.qa-box {
    background: #1C1C1E; border: 1px solid #2C2C2E;
    border-radius: 10px; padding: 12px 14px; margin-bottom: 10px;
}
.qa-score-num { font-size: 32px; font-weight: 700; line-height: 1; }
.qa-lbl { font-size: 9px; font-weight: 700; color: #48484A; text-transform: uppercase; letter-spacing: .1em; margin-top: 2px; }
.qa-track { height: 3px; background: #2C2C2E; border-radius: 2px; margin-top: 6px; overflow: hidden; }
.qa-fill { height: 100%; border-radius: 2px; transition: width .5s; }
.sec-row { display:flex;align-items:center;gap:6px;margin-bottom:3px; }
.sec-ok { color:#34C759;font-size:12px; }
.sec-no { color:#FF453A;font-size:12px; }
.sec-name { font-size:11px;color:#8E8E93;font-weight:500; }

/* ── WARN BLOCKS ── */
.warn-box { background:rgba(255,159,10,.06);border:1px solid rgba(255,159,10,.15);border-radius:8px;padding:10px 12px;margin-bottom:8px; }
.warn-title { font-size:10px;font-weight:700;color:#FF9F0A;text-transform:uppercase;letter-spacing:.04em;margin-bottom:5px; }
.warn-item { font-size:11px;color:#8E8E93;margin-bottom:2px; }
.err-box { background:rgba(255,69,58,.04);border:1px solid rgba(255,69,58,.12);border-radius:8px;padding:10px 12px;margin-bottom:8px; }
.err-title { font-size:10px;font-weight:700;color:#FF453A;text-transform:uppercase;letter-spacing:.04em;margin-bottom:5px; }
.err-item { font-size:11px;color:#8E8E93;margin-bottom:2px; }

/* ── CLASIF CARD ── */
.clasif-card { border:1px solid #2C2C2E;border-radius:8px;margin-bottom:8px;overflow:hidden; }
.clasif-hdr { padding:8px 10px;background:#1C1C1E;font-size:11px;font-weight:600;color:#C7C7CC;border-bottom:1px solid #2C2C2E; }
.clasif-body { padding:8px 10px; }
.clasif-grade { display:flex;gap:8px;padding:4px 0;border-bottom:1px solid #1C1C1E;font-size:11px;color:#8E8E93; }
.clasif-grade:last-child { border-bottom:none; }
.grade-n { min-width:28px;font-weight:700;color:#F5C518;font-family:'JetBrains Mono',monospace;font-size:10px;padding-top:1px; }
.clasif-ref { font-size:9px;color:#48484A;margin-top:6px;padding-top:6px;border-top:1px solid #1C1C1E;font-family:'JetBrains Mono',monospace; }

/* ── SUGGEST CARD ── */
.sug-card { border:1px solid #2C2C2E;border-radius:8px;padding:10px 12px;margin-bottom:8px;background:#1C1C1E;transition:border-color .12s; }
.sug-card:hover { border-color:#3A3A3C; }
.sug-tag { font-size:9px;font-weight:700;color:#48484A;text-transform:uppercase;letter-spacing:.06em; }
.sug-title { font-size:12px;font-weight:600;color:#C7C7CC;margin:3px 0; }
.sug-desc { font-size:11px;color:#636366; }
.sug-action { font-size:10px;color:#F5C518;margin-top:5px;cursor:pointer; }

/* ── AI TOGGLE ── */
.ai-toggle-wrap {
    display: flex; align-items: center; gap: 8px;
    background: #2C2C2E; border: 1px solid #3A3A3C;
    border-radius: 8px; padding: 5px 10px;
    font-size: 12px; font-weight: 600; color: #F5C518;
    cursor: pointer; transition: all .12s;
}
.ai-toggle-wrap:hover { background: #3A3A3C; }

/* ── RIGHT PANEL HEADER ── */
.rp-header {
    padding: 12px 14px;
    border-bottom: 1px solid #2C2C2E;
    display: flex; align-items: center; justify-content: space-between;
    background: #141414;
    position: sticky; top: 0; z-index: 10;
}
.rp-title { font-size: 13px; font-weight: 700; color: #E5E5EA; }
.rp-body { padding: 12px; overflow-y: auto; }

/* ── BOTTOM BAR ── */
.bottom-bar {
    border-top: 1px solid #2C2C2E;
    padding: 8px 14px;
    display: flex; align-items: center; gap: 8px;
    background: #1C1C1E;
    flex-wrap: wrap;
}
.bb-pill {
    display: inline-flex; align-items: center; gap: 5px;
    padding: 5px 12px; border-radius: 20px;
    font-size: 11px; font-weight: 500; cursor: pointer;
    border: 1px solid #3A3A3C; background: #2C2C2E; color: #8E8E93;
    transition: all .12s; white-space: nowrap;
    font-family: 'Inter', sans-serif;
}
.bb-pill:hover { border-color: #F5C518; color: #F5C518; }
.bb-pill.prime { background: #F5C518; border-color: #F5C518; color: #1C1C1E; font-weight: 700; }
.bb-pill.prime:hover { background: #FFD700; box-shadow: 0 4px 12px rgba(245,197,24,.3); }
.bb-saved { font-size: 11px; color: #34C759; margin-left: auto; display: flex; align-items: center; gap: 4px; }
.bb-saved::before { content: '✓'; }
.bb-score { display:flex;align-items:center;gap:8px; }
.bb-score-num { font-size:16px;font-weight:700; }
.bb-score-lbl { font-size:9px;font-weight:700;color:#48484A;text-transform:uppercase;letter-spacing:.1em; }
.bb-track { width:44px;height:3px;background:#2C2C2E;border-radius:2px;margin-top:4px;overflow:hidden; }
.bb-fill { height:100%;border-radius:2px;transition:width .5s; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# LAYOUT — 3 columns: left panel | main | right panel
# ══════════════════════════════════════════════════════════════════
rp_open = st.session_state.right_open
if rp_open:
    col_left, col_main, col_right = st.columns([1.05, 3.4, 1.55], gap="small")
else:
    col_left, col_main = st.columns([1.05, 4.95], gap="small")
    col_right = None

# ══════════════════════════════════════════════════════════════════
# LEFT PANEL
# ══════════════════════════════════════════════════════════════════
with col_left:
    # Render logo header
    st.markdown("""
    <div class="left-panel">
      <div class="lp-header">
        <div class="lp-logo">
          <span class="logo-leaf">🌿</span>
          <span class="logo-text">AURA</span>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # New report button
    st.markdown("<div style='padding:10px 10px 0'>", unsafe_allow_html=True)
    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    if st.button("＋  Nuevo informe", use_container_width=True, key="btn_nuevo"):
        st.session_state.reporte_html = ""
        st.session_state.reporte_texto = ""
        st.session_state.dictado = ""
        st.session_state.clasif_activas = {}
        st.session_state.qa = {}
        st.session_state.copilot_txt = ""
        st.session_state.estudio_id = f"EST-{datetime.datetime.now().strftime('%Y%m%d-%H%M')}"
        st.rerun()
    st.markdown('</div></div>', unsafe_allow_html=True)

    # Estudios recientes
    st.markdown('<div class="lp-section" style="padding:14px 10px 6px">Estudios recientes</div>', unsafe_allow_html=True)
    if st.session_state.historial:
        for i, h in enumerate(reversed(st.session_state.historial[-7:])):
            ms = {"Resonancia Magnética":"RM","Tomografía Computarizada":"TC","Radiografía":"Rx","Ultrasonido":"US","PET-CT":"PET"}.get(h.get("modalidad",""),"?")
            sc = h.get("score",0)
            badge_cls = "badge-ok" if sc>=70 else "badge-active"
            st.markdown(f"""
            <div class="hist-item">
                <div class="hi-title">{ms} {h.get('region','')}</div>
                <div class="hi-meta">{h.get('ts','')} · ID: {h.get('id','')[-8:]}</div>
                <span class="hi-badge {badge_cls}">QA {sc}%</span>
            </div>
            """, unsafe_allow_html=True)
            if st.button("Abrir", key=f"open_{i}", use_container_width=True):
                st.session_state.reporte_texto = h.get("texto","")
                st.session_state.reporte_html  = texto_a_html(h.get("texto",""))
                st.session_state.qa = h.get("qa",{})
                st.rerun()
    else:
        st.markdown('<div style="padding:8px 14px;font-size:12px;color:#3A3A3C">Sin estudios previos</div>', unsafe_allow_html=True)

    # Spacer then dictation zone
    st.markdown("<div style='flex:1'></div>", unsafe_allow_html=True)
    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # Voice input zone — dark card
    st.markdown('<div style="padding:0 10px">', unsafe_allow_html=True)
    st.markdown('<span class="lbl">Dictado de voz</span>', unsafe_allow_html=True)
    audio = st.audio_input("Grabar", label_visibility="collapsed", key="audio_in")
    if audio:
        t = transcribir(audio)
        if t and t not in st.session_state.dictado:
            st.session_state.dictado += (" " if st.session_state.dictado else "") + t
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div style='padding:0 10px;margin-top:8px'>", unsafe_allow_html=True)
    # Bottom row: history icon, add icon, AI toggle
    rc1, rc2, rc3 = st.columns([1,1,2], gap="small")
    with rc1:
        st.markdown('<div class="btn-ghost">', unsafe_allow_html=True)
        if st.button("🕐", use_container_width=True): pass
        st.markdown('</div>', unsafe_allow_html=True)
    with rc2:
        st.markdown('<div class="btn-ghost">', unsafe_allow_html=True)
        if st.button("＋", use_container_width=True): pass
        st.markdown('</div>', unsafe_allow_html=True)
    with rc3:
        ai_label = "AI  ●" if st.session_state.ai_toggle else "AI  ○"
        st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
        if st.button(ai_label, use_container_width=True, key="ai_toggle_btn"):
            st.session_state.ai_toggle = not st.session_state.ai_toggle
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Big yellow MIC / GENERATE button
    st.markdown("<div style='padding:10px 10px 0'>", unsafe_allow_html=True)
    st.markdown('<div class="btn-yellow">', unsafe_allow_html=True)
    procesar = st.button("🎙  Generar informe", use_container_width=True, key="btn_gen")
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Update report button
    st.markdown("<div style='padding:8px 10px 14px'>", unsafe_allow_html=True)
    st.markdown('<div class="btn-update">', unsafe_allow_html=True)
    if st.button("Actualizar informe", use_container_width=True, key="btn_update"):
        pass  # future: sync editor content back
    st.markdown('</div></div>', unsafe_allow_html=True)

    # Settings expander
    with st.expander("CONFIGURACIÓN"):
        st.text_input("API Key", type="password", key="api_key", label_visibility="collapsed", placeholder="DeepSeek / OpenAI key")
        st.radio("Modelo", ["deepseek","openai_mini","openai_4"], key="proveedor",
                 format_func=lambda x:{"deepseek":"DeepSeek","openai_mini":"GPT-4o Mini","openai_4":"GPT-4.1 Mini"}[x],
                 label_visibility="collapsed")
        pf = st.file_uploader("Plantilla .docx", type=["docx"], label_visibility="collapsed")
        if pf: st.session_state.plantilla_txt = leer_plantilla(pf)

    with st.expander("MI ESTILO"):
        ej = st.text_area("Ejemplo", height=70, label_visibility="collapsed", placeholder="Pega un informe tuyo…")
        if st.button("Aprender", use_container_width=True):
            if ej.strip():
                st.session_state.estilo_ejemplos.append(ej.strip())
                st.success(f"{len(st.session_state.estilo_ejemplos)} ejemplo(s)")
        st.text_area("Preferencias", key="estilo_pref", height=40, label_visibility="collapsed",
                     placeholder="Ej: Empezar siempre por meniscos…")

    with st.expander("ESTUDIO"):
        st.selectbox("Modalidad", MODALIDADES, key="modalidad", label_visibility="collapsed")
        st.selectbox("Región", REGIONES, key="region", label_visibility="collapsed")
        st.text_area("Dictado manual", key="dictado", height=80, label_visibility="collapsed",
                     placeholder="Escribe los hallazgos aquí…")
        if st.session_state.dictado:
            st.markdown(f'<div style="font-size:10px;color:#48484A;padding:4px 0">{len(st.session_state.dictado.split())} palabras dictadas</div>', unsafe_allow_html=True)

    with st.expander("CLASIFICACIONES"):
        c_sel = st.selectbox("Clasificación", list(KB_CLASIF.keys()), label_visibility="collapsed", key="c_sel")
        if c_sel:
            g_sel = st.selectbox("Grado", list(KB_CLASIF[c_sel]["grados"].keys()), label_visibility="collapsed", key="g_sel")
            if g_sel:
                st.markdown(f'<div style="font-size:11px;color:#636366;padding:4px 0">{KB_CLASIF[c_sel]["grados"][g_sel]}</div>', unsafe_allow_html=True)
            if st.button("+ Agregar al caso"):
                st.session_state.clasif_activas[c_sel] = g_sel
                st.rerun()
        if st.session_state.clasif_activas:
            for k,v in list(st.session_state.clasif_activas.items()):
                cc1,cc2 = st.columns([4,1])
                with cc1: st.markdown(f'<div style="font-size:11px;color:#8E8E93;padding:2px 0"><span style="color:#F5C518;font-weight:700">{k}</span> Gr.{v}</div>', unsafe_allow_html=True)
                with cc2:
                    if st.button("×", key=f"rm_{k}"): del st.session_state.clasif_activas[k]; st.rerun()

# ══════════════════════════════════════════════════════════════════
# PROCESSING
# ══════════════════════════════════════════════════════════════════
if procesar:
    client, model_name = get_client()
    dictado = st.session_state.dictado
    if client and dictado.strip():
        with st.spinner("Generando informe AURA…"):
            try:
                res = client.chat.completions.create(
                    model=model_name,
                    messages=[
                        {"role":"system","content":build_system_prompt()},
                        {"role":"user","content":f"DICTADO:\n{dictado}"}
                    ], temperature=0.1
                )
                texto = res.choices[0].message.content
                st.session_state.reporte_texto = texto
                st.session_state.reporte_html  = texto_a_html(texto)
                qa = calcular_qa(texto, st.session_state.region)
                st.session_state.qa = qa
                ts  = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                eid = st.session_state.estudio_id
                st.session_state.historial.append({"texto":texto,"qa":qa,"region":st.session_state.region,
                    "modalidad":st.session_state.modalidad,"score":qa["score"],"ts":ts,"id":eid})
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
    elif not client:
        st.warning("Ingresa tu API Key en Configuración.")
    else:
        st.warning("Escribe o dicta los hallazgos primero.")

# ══════════════════════════════════════════════════════════════════
# MAIN PANEL — topbar + editor
# ══════════════════════════════════════════════════════════════════
with col_main:
    now_str = datetime.datetime.now().strftime("%d %b %Y, %I:%M %p")
    region    = st.session_state.region
    modalidad = st.session_state.modalidad
    ms = {"Resonancia Magnética":"RM","Tomografía Computarizada":"TC","Radiografía":"Rx",
          "Ultrasonido":"US","PET-CT":"PET"}.get(modalidad, modalidad[:2])
    titulo = f"{ms} {region.upper()}"

    # ── Topbar ──
    st.markdown(f"""
    <div class="aura-topbar">
        <button class="tb-template-btn">📤 Upload Templates</button>
        <div style="position:relative">
            <select class="tb-study-select">
                <option>Seleccionar plantilla</option>
                <option>{modalidad}</option>
            </select>
        </div>
        <div class="tb-sep"></div>
        <div style="margin-left:auto;display:flex;align-items:center;gap:8px">
            <button class="tb-reports-btn">📊 Reports</button>
            <div class="tb-sep"></div>
            <button class="tb-new-btn">📄 New Report</button>
            <div class="tb-sep"></div>
            <div class="tb-user">
                <div class="tb-user-dot">A</div>
                Dr. Alejandro M.
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Rich Editor ──
    content = st.session_state.reporte_html or DEFAULT_CONTENT

    editor_html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{
    background: #1C1C1E;
    font-family: 'Inter', sans-serif;
    height: 100vh; overflow: hidden;
    display: flex; flex-direction: column;
    color: #E5E5EA;
}}

/* TOOLBAR */
.toolbar {{
    background: #141414;
    border-bottom: 1px solid #2C2C2E;
    padding: 8px 14px;
    display: flex; align-items: center; gap: 4px; flex-wrap: wrap;
    position: sticky; top: 0; z-index: 50;
}}
.tg {{
    display: flex; align-items: center; gap: 2px;
    padding-right: 8px; border-right: 1px solid #2C2C2E;
}}
.tg:last-of-type {{ border-right: none; padding-right: 0; }}
.tb {{
    background: none; border: 1px solid transparent; color: #636366;
    font-size: 12px; font-family: 'Inter', sans-serif;
    padding: 4px 8px; border-radius: 5px; cursor: pointer;
    transition: all .1s; min-width: 28px; text-align: center;
    line-height: 1.4; white-space: nowrap;
}}
.tb:hover {{ background: #2C2C2E; color: #C7C7CC; border-color: #3A3A3C; }}
.tb.on {{ background: rgba(245,197,24,.12); border-color: rgba(245,197,24,.25); color: #F5C518; }}
.tsel {{
    background: #2C2C2E; border: 1px solid #3A3A3C; color: #8E8E93;
    font-size: 11px; font-family: 'Inter', sans-serif;
    padding: 4px 6px; border-radius: 5px; outline: none; cursor: pointer;
    appearance: none;
}}
.tsel:focus {{ border-color: #F5C518; }}
.cbg {{
    width: 15px; height: 15px; border-radius: 50%;
    cursor: pointer; border: 2px solid transparent; flex-shrink: 0;
    transition: all .12s;
}}
.cbg:hover, .cbg.on {{ border-color: #F5C518; box-shadow: 0 0 0 2px rgba(245,197,24,.2); }}
.tlabel {{ font-size: 10px; color: #48484A; white-space: nowrap; font-family: 'Inter', sans-serif; }}
.tb-icon {{ font-size: 13px; line-height: 1; }}
.copy-btn {{
    margin-left: auto;
    background: #2C2C2E; border: 1px solid #3A3A3C; color: #8E8E93;
    font-size: 11px; font-family: 'Inter', sans-serif; font-weight: 500;
    padding: 4px 10px; border-radius: 6px; cursor: pointer;
    display: flex; align-items: center; gap: 5px; transition: all .12s;
    white-space: nowrap;
}}
.copy-btn:hover {{ color: #E5E5EA; border-color: #48484A; }}

/* EDITOR */
.editor-wrap {{ flex: 1; overflow-y: auto; }}
#doc {{
    min-height: 100%; padding: 30px 40px;
    font-family: 'Inter', sans-serif; font-size: 14px;
    line-height: 1.9; color: #C7C7CC;
    outline: none; background: #1C1C1E;
    transition: background .3s, color .3s;
}}
#doc h3.sec {{
    font-family: 'Inter', sans-serif;
    font-size: 11px; font-weight: 700;
    color: #E5E5EA; letter-spacing: .08em;
    text-transform: uppercase; margin: 22px 0 8px;
    padding-bottom: 7px;
    border-bottom: 1px solid #2C2C2E;
}}
#doc p {{ margin-bottom: 8px; color: #C7C7CC; line-height: 1.85; }}
#doc strong, #doc b {{ color: #E5E5EA; font-weight: 600; }}
#doc li {{
    margin-left: 18px; margin-bottom: 5px;
    color: #C7C7CC; list-style-type: disc; padding-left: 4px;
}}
#doc hr {{ border: none; border-top: 1px solid #2C2C2E; margin: 16px 0; }}

/* BOTTOM BAR */
.bbar {{
    background: #141414; border-top: 1px solid #2C2C2E;
    padding: 8px 14px; display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
}}
.bp {{
    display: inline-flex; align-items: center; gap: 5px;
    padding: 5px 13px; border-radius: 20px;
    font-size: 11px; font-weight: 500; cursor: pointer;
    border: 1px solid #3A3A3C; background: #2C2C2E; color: #8E8E93;
    transition: all .12s; white-space: nowrap;
    font-family: 'Inter', sans-serif;
}}
.bp:hover {{ border-color: #F5C518; color: #F5C518; }}
.bp.prime {{
    background: #F5C518; border-color: #F5C518; color: #1C1C1E; font-weight: 700;
}}
.bp.prime:hover {{ background: #FFD700; box-shadow: 0 4px 12px rgba(245,197,24,.3); }}

/* SCORE */
.score-wrap {{ margin-left: auto; display: flex; align-items: center; gap: 10px; }}
.score-n {{ font-size: 18px; font-weight: 700; color: #E5E5EA; font-family: 'Inter', sans-serif; }}
.score-l {{ font-size: 9px; font-weight: 700; color: #48484A; text-transform: uppercase; letter-spacing: .1em; }}
.score-track {{ width: 44px; height: 3px; background: #2C2C2E; border-radius: 2px; margin-top: 5px; overflow: hidden; }}
.score-bar {{ height: 100%; border-radius: 2px; transition: width .5s; }}
.saved {{ font-size: 11px; color: #34C759; display: flex; align-items: center; gap: 4px; }}

/* SLASH MENU */
.sm {{
    position: fixed; background: #1C1C1E;
    border: 1px solid #3A3A3C; border-radius: 10px;
    box-shadow: 0 12px 32px rgba(0,0,0,.7);
    z-index: 1000; min-width: 280px; overflow: hidden; display: none;
}}
.sm-hdr {{
    padding: 8px 14px 5px; font-size: 10px; font-weight: 700;
    color: #48484A; letter-spacing: .08em; text-transform: uppercase;
    border-bottom: 1px solid #2C2C2E;
}}
.si {{
    padding: 9px 14px; cursor: pointer; transition: background .1s;
    font-size: 12px; color: #C7C7CC;
    display: flex; align-items: center; gap: 10px;
    font-family: 'Inter', sans-serif;
}}
.si:hover, .si.sel {{ background: #2C2C2E; }}
.sk {{ color: #F5C518; font-weight: 700; font-family: 'JetBrains Mono', monospace; font-size: 11px; min-width: 95px; }}
.sd {{ color: #636366; font-size: 11px; }}
</style>
</head>
<body>

<!-- TOOLBAR -->
<div class="toolbar">
  <div class="tg">
    <button class="tb tb-icon" onclick="fmt('undo')" title="Deshacer">↩</button>
    <button class="tb tb-icon" onclick="fmt('redo')" title="Rehacer">↪</button>
  </div>
  <div class="tg">
    <button class="tb" id="btnB" onclick="fmt('bold')" title="Negrita (Ctrl+B)"><b>B</b></button>
    <button class="tb" id="btnI" onclick="fmt('italic')" title="Cursiva"><i>I</i></button>
    <button class="tb" id="btnU" onclick="fmt('underline')" title="Subrayado"><u>U</u></button>
    <button class="tb" id="btnS" onclick="fmt('strikeThrough')" title="Tachado"><s>S</s></button>
  </div>
  <div class="tg">
    <button class="tb" onclick="fmt('insertUnorderedList')" title="Viñetas">• Lista</button>
    <button class="tb" onclick="fmt('insertOrderedList')" title="Numerada">1. Lista</button>
  </div>
  <div class="tg">
    <button class="tb" onclick="fmt('justifyLeft')" title="Izquierda">⬛</button>
    <button class="tb" onclick="fmt('justifyCenter')" title="Centro">▣</button>
    <button class="tb" onclick="fmt('justifyFull')" title="Justificado">▤</button>
  </div>
  <div class="tg">
    <button class="tb" onclick="insSection()" title="Sección">§ Sec</button>
    <button class="tb" onclick="insHR()" title="Separador">— HR</button>
  </div>
  <div class="tg" style="gap:8px;align-items:center">
    <span class="tlabel">AI:</span>
    <select class="tsel" style="width:70px" title="Estilo de lenguaje">
      <option>Académico</option><option>Conservador</option><option>Directo</option>
    </select>
    <select class="tsel" style="width:90px" title="Destinatario">
      <option>Médico tratante</option><option>Paciente</option><option>Referente</option>
    </select>
  </div>
  <div class="tg" style="gap:5px;align-items:center">
    <span class="tlabel">Fondo</span>
    <div class="cbg on"  style="background:#1C1C1E" onclick="setBg(this,'#1C1C1E','#C7C7CC')" title="Dark"></div>
    <div class="cbg"     style="background:#0A1018" onclick="setBg(this,'#0A1018','#C8DAF0')" title="PACS"></div>
    <div class="cbg"     style="background:#FFFFFF" onclick="setBg(this,'#FFFFFF','#2D3748')" title="Blanco"></div>
    <div class="cbg"     style="background:#FAFAF8" onclick="setBg(this,'#FAFAF8','#2D3748')" title="Warm"></div>
  </div>
  <button class="copy-btn" onclick="copyAll()">⎘ Copiar informe</button>
</div>

<!-- EDITOR -->
<div class="editor-wrap">
  <div id="doc" contenteditable="true" spellcheck="true" lang="es">
    {content}
  </div>
</div>

<!-- SLASH MENU -->
<div id="sm" class="sm">
  <div class="sm-hdr">Plantillas & Clasificaciones — escribe /comando</div>
</div>

<!-- BOTTOM BAR -->
<div class="bbar">
  <button class="bp" onclick="send('suggest_conclusion')">＋ Sugerir conclusión</button>
  <button class="bp" onclick="send('optimize')">✦ Optimizar</button>
  <button class="bp" onclick="send('differential')">⊕ Dif. diagnóstico</button>
  <button class="bp" onclick="send('definiciones')">◎ Definiciones</button>
  <button class="bp" onclick="send('qa_full')">◈ Auditar QA</button>
  <button class="bp prime" onclick="send('export')">↓ Exportar .docx</button>
  <span class="saved" id="savedLbl" style="display:none">Guardado {now_str.split(",")[1].strip()}</span>
  <div class="score-wrap">
    <div>
      <div class="score-n" id="sNum">—</div>
      <div class="score-l">Calidad</div>
      <div class="score-track"><div class="score-bar" id="sBar" style="width:0;background:#48484A"></div></div>
    </div>
  </div>
</div>

<script>
var doc=document.getElementById('doc');
var sm=document.getElementById('sm');
var slashOn=false, slashSel=0;

var CMDS=[
  {{k:'/stoller',l:'Stoller',d:'Clasificación meniscal RM',t:'<h3 class="sec">MENISCOS</h3><p><strong>Menisco medial/lateral:</strong> alteración de señal grado [I/II/III] de Stoller en [cuerpo / cuerno], compatible con [cambio degenerativo / desgarro]. [Extrusión de __ mm].</p>'}},
  {{k:'/icrs',l:'ICRS',d:'Cartílago articular',t:'<p><strong>Cartílago articular:</strong> adelgazamiento condral focal grado [I/II/III/IV] de ICRS en [localización], extensión de __ mm. [Edema subcondral reactivo].</p>'}},
  {{k:'/kl',l:'Kellgren-Lawrence',d:'Osteoartrosis',t:'<p>Hallazgos compatibles con osteoartrosis grado [I/II/III/IV] de Kellgren-Lawrence en [compartimento].</p>'}},
  {{k:'/lca',l:'Hope & Feagin',d:'Lesión LCA',t:'<p><strong>Ligamento cruzado anterior:</strong> [señal heterogénea / discontinuidad], compatible con lesión [parcial / completa / crónica] según Hope &amp; Feagin.</p>'}},
  {{k:'/pfirrmann',l:'Pfirrmann',d:'Degeneración discal',t:'<p>Disco [L_-L_]: degeneración grado [I/II/III/IV/V] de Pfirrmann. [Descripción de señal y altura discal].</p>'}},
  {{k:'/modic',l:'Modic',d:'Placa terminal vertebral',t:'<p>Cambios tipo Modic [I/II/III] en placa terminal de [nivel], indicativos de [edema/grasa/esclerosis].</p>'}},
  {{k:'/tirads',l:'ACR TIRADS',d:'Nódulo tiroideo',t:'<p><strong>Nódulo tiroideo</strong> [lóbulo]: [descripción]. Clasificación ACR TIRADS [2-5]. [Recomendación según tamaño].</p>'}},
  {{k:'/birads',l:'BI-RADS',d:'Hallazgo mamario',t:'<p><strong>Hallazgo mamario:</strong> [descripción]. Categoría ACR BI-RADS [1-6]. [Recomendación].</p>'}},
  {{k:'/lirads',l:'LI-RADS',d:'Lesión hepática',t:'<p><strong>Lesión focal hepática:</strong> [descripción]. Clasificación LI-RADS [LR-1 a LR-5/M].</p>'}},
  {{k:'/fleischner',l:'Fleischner 2017',d:'Nódulo pulmonar',t:'<p><strong>Nódulo pulmonar</strong> [sólido/subsólido] de __ mm en [lóbulo]. Criterios Fleischner 2017: [recomendación de seguimiento].</p>'}},
  {{k:'/tecnica',l:'Técnica estándar',d:'Plantilla de técnica',t:'<h3 class="sec">TÉCNICA</h3><p>Estudio de [modalidad] de [región] en equipo de [__] Tesla. Secuencias [describir] en planos [axial/coronal/sagital]. [Sin / Con] contraste.</p>'}},
  {{k:'/impresion',l:'Impresión DX',d:'Plantilla de conclusión',t:'<h3 class="sec">IMPRESIÓN DIAGNÓSTICA</h3><li>Hallazgo principal con clasificación y grado. Recomendación de manejo.</li><li>Hallazgo secundario.</li><li>Seguimiento o estudios complementarios.</li>'}},
  {{k:'/conclusion',l:'Conclusión rápida',d:'Bloque de conclusión simple',t:'<h3 class="sec">CONCLUSIÓN</h3><li>.</li><li>Hallazgos descritos.</li>'}},
];

function fmt(cmd){{ doc.focus(); document.execCommand(cmd,false,null); updBtns(); }}
function updBtns(){{
  [['bold','btnB'],['italic','btnI'],['underline','btnU'],['strikeThrough','btnS']].forEach(function(p){{
    var b=document.getElementById(p[1]);
    if(b) b.classList.toggle('on',document.queryCommandState(p[0]));
  }});
}}
function setBg(el,bg,col){{
  doc.style.background=bg; doc.style.color=col;
  document.querySelectorAll('.cbg').forEach(function(d){{d.classList.remove('on')}});
  el.classList.add('on');
  var isDark=bg!=='#FFFFFF'&&bg!=='#FAFAF8';
  document.querySelectorAll('#doc h3.sec').forEach(function(t){{
    t.style.color=isDark?'#E5E5EA':'#1A202C';
    t.style.borderBottomColor=isDark?'#2C2C2E':'#E2E8F0';
  }});
  document.querySelectorAll('#doc strong,#doc b').forEach(function(t){{
    t.style.color=isDark?'#E5E5EA':'#1A202C';
  }});
}}
function insSection(){{doc.focus();document.execCommand('insertHTML',false,'<h3 class="sec">NUEVA SECCIÓN</h3><p></p>');}}
function insHR(){{doc.focus();document.execCommand('insertHTML',false,'<hr>');}}

function calcScore(){{
  var t=doc.innerText.toUpperCase();
  var secs=['TÉCNICA','HALLAZGOS','IMPRESIÓN','CONCLUSIÓN'];
  var found=secs.filter(function(s){{return t.indexOf(s)>=0}}).length;
  var words=doc.innerText.trim().split(' ').filter(Boolean).length;
  return Math.min(100,Math.round((found/4)*40+Math.min(words/200,1)*40+(found>=2?20:0)));
}}

function updScore(){{
  var s=calcScore();
  document.getElementById('sNum').textContent=s+'%';
  var b=document.getElementById('sBar');
  b.style.width=s+'%';
  b.style.background=s>=80?'#34C759':s>=50?'#FF9F0A':'#FF453A';
  if(s>0) document.getElementById('savedLbl').style.display='flex';
}}

/* SLASH COMMAND */
function showSlash(filter){{
  var items=filter?CMDS.filter(function(c){{return c.k.includes(filter)||c.l.toLowerCase().includes(filter)}}):CMDS;
  if(!items.length){{hideSlash();return;}}
  var sel=window.getSelection();
  if(!sel.rangeCount)return;
  var r=sel.getRangeAt(0).getBoundingClientRect();
  var html=items.map(function(c,i){{
    return '<div class="si'+(i===0?' sel':'')+'" onclick="insCmd(\''+c.k+'\')"><span class="sk">'+c.k+'</span><div><div style="font-weight:600;font-size:12px">'+c.l+'</div><span class="sd">'+c.d+'</span></div></div>';
  }}).join('');
  sm.innerHTML='<div class="sm-hdr">Plantillas &amp; Clasificaciones</div>'+html;
  sm.style.display='block';
  sm.style.left=Math.max(8,r.left)+'px';
  sm.style.top=(r.bottom+6)+'px';
  slashOn=true; slashSel=0;
}}
function hideSlash(){{sm.style.display='none';slashOn=false;slashSel=0;}}
function insCmd(cmd){{
  hideSlash(); doc.focus();
  var sel=window.getSelection();
  if(sel.rangeCount){{
    var range=sel.getRangeAt(0);
    var txt=(range.startContainer.textContent||'').slice(0,range.startOffset);
    var si=txt.lastIndexOf('/');
    if(si>=0){{range.setStart(range.startContainer,si);range.deleteContents();}}
  }}
  var c=CMDS.find(function(x){{return x.k===cmd;}});
  if(c) document.execCommand('insertHTML',false,c.t);
}}

doc.addEventListener('keydown',function(e){{
  if(slashOn){{
    var its=sm.querySelectorAll('.si');
    if(e.key==='Escape'){{hideSlash();e.preventDefault();return;}}
    if(e.key==='Enter'){{e.preventDefault();if(its[slashSel])its[slashSel].click();return;}}
    if(e.key==='ArrowDown'){{e.preventDefault();slashSel=(slashSel+1)%its.length;its.forEach(function(el,i){{el.classList.toggle('sel',i===slashSel)}});return;}}
    if(e.key==='ArrowUp'){{e.preventDefault();slashSel=(slashSel-1+its.length)%its.length;its.forEach(function(el,i){{el.classList.toggle('sel',i===slashSel)}});return;}}
  }}
}});
doc.addEventListener('keyup',function(e){{
  updBtns(); updScore();
  var sel=window.getSelection();
  if(!sel.rangeCount)return;
  var r=sel.getRangeAt(0);
  var txt=((r.startContainer.textContent||'').slice(0,r.startOffset));
  var m=txt.match(new RegExp('/(\\w*)$'));
  if(m)showSlash(m[1]); else hideSlash();
}});
doc.addEventListener('mouseup',function(){{updBtns();hideSlash();}});
document.addEventListener('click',function(e){{if(!sm.contains(e.target))hideSlash();}});

function copyAll(){{
  var r=document.createRange();r.selectNode(doc);
  window.getSelection().removeAllRanges();
  window.getSelection().addRange(r);
  document.execCommand('copy');
  window.getSelection().removeAllRanges();
}}

function send(type){{
  window.parent.postMessage({{type:type,html:doc.innerHTML,text:doc.innerText}},'*');
}}

window.addEventListener('load',function(){{updScore();}});
</script>
</body>
</html>"""

    components.html(editor_html, height=700, scrolling=False)

    # ── Action buttons below editor ──
    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
    ba1, ba2, ba3, ba4 = st.columns(4, gap="small")
    with ba1:
        if st.button("✦ Optimizar conclusión", use_container_width=True, key="b_opt"):
            client, model_name = get_client()
            if client and st.session_state.reporte_texto:
                with st.spinner("Refinando…"):
                    try:
                        res = client.chat.completions.create(model=model_name, messages=[
                            {"role":"system","content":build_system_prompt()},
                            {"role":"user","content":f"MEJORA SOLO la IMPRESIÓN/CONCLUSIÓN. Conserva Técnica y Hallazgos intactos. Devuelve el informe COMPLETO.\n\n{st.session_state.reporte_texto}"}
                        ], temperature=0.2)
                        t = res.choices[0].message.content
                        st.session_state.reporte_texto = t
                        st.session_state.reporte_html  = texto_a_html(t)
                        st.session_state.qa = calcular_qa(t, st.session_state.region)
                        st.rerun()
                    except Exception as e: st.error(str(e))
    with ba2:
        if st.button("⊕ Diagnóstico diferencial", use_container_width=True, key="b_dd"):
            client, model_name = get_client()
            if client and st.session_state.reporte_texto:
                with st.spinner("Generando DD…"):
                    try:
                        res = client.chat.completions.create(model=model_name, messages=[
                            {"role":"user","content":f"Como radiólogo experto, genera diagnóstico diferencial estructurado (3-5 opciones) con argumentos. Luego sugiere estudios complementarios. Sin asteriscos.\n\nINFORME:\n{st.session_state.reporte_texto}"}
                        ], temperature=0.3)
                        st.session_state.copilot_txt  = res.choices[0].message.content
                        st.session_state.copilot_tipo = "differential"
                        st.session_state.right_open   = True
                        st.rerun()
                    except Exception as e: st.error(str(e))
    with ba3:
        if st.button("◎ Definiciones", use_container_width=True, key="b_def"):
            client, model_name = get_client()
            if client and st.session_state.reporte_texto:
                with st.spinner("Analizando…"):
                    try:
                        res = client.chat.completions.create(model=model_name, messages=[
                            {"role":"user","content":f"Analiza el informe. Proporciona: (1) CLASIFICACIONES USADAS con significado exacto, (2) FALTANTES, (3) DEFINICIONES de términos técnicos, (4) BIBLIOGRAFÍA. Conciso.\n\nINFORME:\n{st.session_state.reporte_texto}"}
                        ], temperature=0.15)
                        st.session_state.copilot_txt  = res.choices[0].message.content
                        st.session_state.copilot_tipo = "definiciones"
                        st.session_state.right_open   = True
                        st.rerun()
                    except Exception as e: st.error(str(e))
    with ba4:
        if st.session_state.reporte_texto:
            docx_b = generar_docx(st.session_state.reporte_html, st.session_state.reporte_texto)
            reg_s  = st.session_state.region.replace(" ","_").replace("/","_")
            st.download_button("↓ Exportar .docx", data=docx_b,
                file_name=f"AURA_{reg_s}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True)

# ══════════════════════════════════════════════════════════════════
# RIGHT PANEL
# ══════════════════════════════════════════════════════════════════
# Toggle trigger (always visible)
st.markdown("""
<div class="rp-trigger" onclick="window.parent.postMessage({type:'toggle_right'},'*')">◀</div>
""", unsafe_allow_html=True)

if col_right:
    with col_right:
        # Header with tabs
        st.markdown("""
        <div class="rp-header">
          <span class="rp-title">Asistente IA</span>
        </div>
        """, unsafe_allow_html=True)

        rp_t1, rp_t2, rp_t3 = st.columns(3, gap="small")
        with rp_t1:
            if st.button("Sugerencias", use_container_width=True, key="tab_sug"):
                st.session_state.copilot_tipo = "sugerencias"
                st.rerun()
        with rp_t2:
            if st.button("Clasificaciones", use_container_width=True, key="tab_cla"):
                st.session_state.copilot_tipo = "clasificaciones"
                st.rerun()
        with rp_t3:
            if st.button("Definiciones", use_container_width=True, key="tab_def"):
                st.session_state.copilot_tipo = "definiciones_tab"
                st.rerun()

        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

        # ── Close panel ──
        st.markdown('<div class="btn-ghost">', unsafe_allow_html=True)
        if st.button("✕ Cerrar panel", use_container_width=True, key="close_rp"):
            st.session_state.right_open = False
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

        # ── QA Score ──
        qa = st.session_state.qa
        score = qa.get("score",0)
        sc = "#34C759" if score>=80 else "#FF9F0A" if score>=50 else "#FF453A"
        st.markdown(f"""
        <div class="qa-box">
          <div style="display:flex;align-items:center;gap:12px">
            <div>
              <div class="qa-score-num" style="color:{sc}">{score}</div>
              <div class="qa-lbl">QA Score</div>
              <div class="qa-track" style="width:52px">
                <div class="qa-fill" style="width:{score}%;background:{sc}"></div>
              </div>
            </div>
            <div>
              {''.join([
                f'<div class="sec-row"><span class="{"sec-ok" if v else "sec-no"}">{"✓" if v else "✗"}</span><span class="sec-name">{k}</span></div>'
                for k,v in qa.get("secciones",{"TÉCNICA":False,"HALLAZGOS":False,"IMPRESIÓN":False}).items()
              ])}
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # ── Omisiones ──
        omisiones = qa.get("omisiones",[])
        if omisiones:
            items_html = "".join([f'<div class="warn-item">· {o}</div>' for o in omisiones])
            st.markdown(f'<div class="warn-box"><div class="warn-title">⚠ Omisiones anatómicas</div>{items_html}</div>', unsafe_allow_html=True)

        # ── Copilot result ──
        if st.session_state.copilot_txt:
            tipo_lbl = {"differential":"Diagnóstico Diferencial","definiciones":"Definiciones & Clasificaciones","qa_full":"Auditoría QA"}.get(st.session_state.copilot_tipo,"Asistente AURA")
            st.markdown(f"""
            <div class="copilot-box">
                <div class="copilot-box-title">{tipo_lbl}</div>
                <div class="copilot-text">{st.session_state.copilot_txt}</div>
            </div>
            """, unsafe_allow_html=True)
            st.markdown('<div class="btn-ghost">', unsafe_allow_html=True)
            if st.button("✕ Cerrar resultado", use_container_width=True, key="close_copilot"):
                st.session_state.copilot_txt = ""
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

        # Sugerencias contextuales
        if not st.session_state.copilot_txt:
            region_sel = st.session_state.region
            clasif_region = KB_REGIONES.get(region_sel,{}).get("clasif",[])
            if clasif_region:
                st.markdown('<div style="font-size:10px;font-weight:700;color:#48484A;text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px">Sugerencias contextuales</div>', unsafe_allow_html=True)
                for cn in clasif_region[:3]:
                    cd = KB_CLASIF.get(cn,{})
                    st.markdown(f"""
                    <div class="sug-card">
                        <div class="sug-tag">Clasificación sugerida</div>
                        <div class="sug-title">{cn}</div>
                        <div class="sug-desc">{cd.get('desc','')}</div>
                        <div class="sug-action">Ver grados →</div>
                    </div>
                    """, unsafe_allow_html=True)

        with st.expander("REFERENCIA RÁPIDA", expanded=False):
            for cn in KB_REGIONES.get(st.session_state.region,{}).get("clasif",[]):
                cd2 = KB_CLASIF.get(cn,{})
                st.markdown(f"""
                <div class="clasif-card">
                    <div class="clasif-hdr">{cn}</div>
                    <div class="clasif-body">
                        {''.join([f'<div class="clasif-grade"><span class="grade-n">{g}</span><span>{desc}</span></div>' for g,desc in cd2.get("grados",{}).items()])}
                    </div>
                </div>
                """, unsafe_allow_html=True)

        with st.expander("AUDITORÍA QA", expanded=False):
            if st.button("◈ Auditar informe", use_container_width=True, key="qa_full"):
                client, model_name = get_client()
                if client and st.session_state.reporte_texto:
                    with st.spinner("Auditando…"):
                        try:
                            res = client.chat.completions.create(model=model_name, messages=[
                                {"role":"user","content":f"QA como radiólogo experto: estructura, completitud anatómica de {st.session_state.region}, clasificaciones, terminología, impresión diagnóstica. Califica 0-100. Sin asteriscos.\n\nINFORME:\n{st.session_state.reporte_texto}"}
                            ], temperature=0.15)
                            st.session_state.copilot_txt  = res.choices[0].message.content
                            st.session_state.copilot_tipo = "qa_full"
                            st.rerun()
                        except Exception as e:
                            st.error(str(e))
                else:
                    st.warning("Genera un informe primero.")

else:
    # Panel cerrado — mostrar botón para reabrir
    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
    if st.button("▶ Abrir panel IA", key="open_rp"):
        st.session_state.right_open = True
        st.rerun()
