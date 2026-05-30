import streamlit as st
import streamlit.components.v1 as components
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile, io, os, re, json, datetime

st.set_page_config(page_title="AURA", layout="wide", initial_sidebar_state="collapsed")

# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# CONSTANTES
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
MODELS = {
    "DeepSeek Chat": {"url": "https://api.deepseek.com", "id": "deepseek-chat"},
    "GPT-4o Mini":   {"url": None, "id": "gpt-4o-mini"},
    "GPT-4.1 Mini":  {"url": None, "id": "gpt-4.1-mini"},
}

MODALIDADES = [
    "Resonancia MagnÃ©tica", "TomografÃ­a Computarizada", "RadiografÃ­a",
    "Ultrasonido", "PET-CT", "MamografÃ­a", "FluoroscopÃ­a", "AngiografÃ­a",
]

REGIONES = {
    "Extremidades inferiores": ["Rodilla", "Cadera", "Tobillo", "Pie", "Muslo", "Pierna"],
    "Extremidades superiores": ["Hombro", "Codo", "MuÃ±eca", "Mano", "Brazo", "Antebrazo"],
    "Columna":                 ["Col. cervical", "Col. dorsal", "Col. lumbar", "Sacro/CÃ³ccix"],
    "CrÃ¡neo y cuello":         ["Cerebro", "Cuello", "Tiroides", "Ãrbitas", "OÃ­dos", "Silla turca"],
    "TÃ³rax":                   ["TÃ³rax", "PulmÃ³n", "CorazÃ³n", "Mediastino", "Mama"],
    "Abdomen y pelvis":        ["Abdomen", "Pelvis", "HÃ­gado", "PÃ¡ncreas", "RiÃ±ones",
                                "Vejiga", "PrÃ³stata", "Ãtero/Anexos", "Suprarrenales", "Bazo"],
}

HCOLS = ["#7c6af7", "#22c55e", "#f59e0b", "#ec4899", "#38bdf8", "#fb923c"]

PLANTILLAS_DEFAULT = {
    "Sin plantilla": "",
    "MusculoesquelÃ©tico RM": """INDICACIÃN
[Motivo de estudio]

TÃCNICA
Estudio de resonancia magnÃ©tica de [regiÃ³n], realizado en equipo de [campo] Tesla, con secuencias [SE/FSE/GRE] en planos [axial/sagital/coronal], con y sin saturaciÃ³n grasa. [Contraste: se administrÃ³/no se administrÃ³ gadolinio].

HALLAZGOS
Partes blandas periarticulares:
Hueso subcondral y mÃ©dula Ã³sea:
CartÃ­lago articular:
Meniscos / FibrocartÃ­lago:
Ligamentos:
Tendones:
LÃ­quido articular:
Hallazgos adicionales:

IMPRESIÃN DIAGNÃSTICA
""",
    "Neuro RM Cerebro": """INDICACIÃN
[Motivo de estudio]

TÃCNICA
Estudio de resonancia magnÃ©tica cerebral realizado en equipo de [campo] Tesla. Secuencias obtenidas: T1, T2, FLAIR, difusiÃ³n (DWI/ADC), T2* [y secuencias adicionales]. [Contraste: se administrÃ³/no se administrÃ³ gadolinio intravenoso].

HALLAZGOS
ParÃ©nquima supratentorial:
ParÃ©nquima infratentorial:
Sistema ventricular y espacios subaracnoideos:
Estructuras de la lÃ­nea media:
Vasculatura intracraneal (si aplica):
Senos paranasales y base de crÃ¡neo:

IMPRESIÃN DIAGNÃSTICA
""",
    "TC Abdomen": """INDICACIÃN
[Motivo de estudio]

TÃCNICA
TomografÃ­a computarizada de abdomen [y pelvis], adquirida en fase [simple/arterial/portal/excretora], con colimaciÃ³n de [X] mm. [Contraste: se administrÃ³ contraste yodado IV/oral/no se administrÃ³].

HALLAZGOS
HÃ­gado:
VÃ­as biliares y vesÃ­cula:
P¡ncreas:
Bazo:
Suprarrenales:
RiÃ±ones y vÃ­as urinarias:
Retroperitoneo y vasos:
Asas intestinales:
Pelvis:
Pared abdominal:

IMPRESIÃN DIAGNÃSTICA
""",
    "TÃ³rax Rx/TC": """INDICACIÃN
[Motivo de estudio]

TÃCNICA
[RadiografÃ­a de tÃ³rax PA y lateral / TomografÃ­a computarizada de tÃ³rax] realizada en [posiciÃ³n/fase respiratoria]. [Contraste: se administrÃ³/no se administrÃ³].

HALLAZGOS
ParÃ©nquima pulmonar:
Hilios pulmonares:
Mediastino:
Silueta cardÃ­aca:
Pleura y espacios pleurales:
Pared torÃ¡cica y estructuras Ã³seas:
Hallazgos subdiafragmÃ¡ticos:

IMPRESIÃN DIAGNÃSTICA
""",
}

DEFAULTS = {
    "dictado": "", "reporte": "", "defs": "", "mentor_feedback": "",
    "modelo": "DeepSeek Chat", "audio_id": None,
    "historial": [], "plantilla_txt": "", "plantilla_nombre": "Sin plantilla",
    "tema": "dark",
    "estilo_aprendido": [],
    "nav_active": "informe",
    "plantillas_custom": {},
    "diccionario": {},
    "sel_mod": "Resonancia MagnÃ©tica",
    "sel_grupo": "Extremidades inferiores",
    "sel_reg": "Rodilla",
    "reg_custom": "",
    "panel_dict_open": False,
    "sugerencias_activas": [],
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

try:    api_key = st.secrets["deepseek_key"]
except: api_key = os.environ.get("OPENAI_API_KEY", "")

# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# HELPERS
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
def get_client():
    cfg = MODELS[st.session_state.modelo]
    return OpenAI(api_key=api_key, base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)

def leer_plantilla_docx(f):
    doc = Document(f); partes = []; n = 0
    try:
        import docx.text.paragraph as pp
        import docx.table as tt
        for el in doc.element.body:
            tag = el.tag.split('}')[-1]
            if tag == 'p':
                p = pp.Paragraph(el, doc); tx = p.text.strip()
                if tx: partes.append(tx)
            elif tag == 'tbl':
                n += 1; tbl = tt.Table(el, doc)
                rows = ["| " + " | ".join(c.text.strip() for c in r.cells) + " |"
                        for r in tbl.rows]
                partes.append(f"[TABLA {n}]\n" + "\n".join(rows) + "\n[/TABLA]")
    except:
        partes = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n".join(partes)

def limpiar_md(texto):
    texto = re.sub(r'\*\*(.+?)\*\*', r'\1', texto)
    texto = re.sub(r'\*(.+?)\*',     r'\1', texto)
    texto = re.sub(r'^\*\s+', 'â¢ ',  texto, flags=re.MULTILINE)
    texto = re.sub(r'\*+', '',        texto)
    texto = re.sub(r'^#+\s+', '',     texto, flags=re.MULTILINE)
    return texto.strip()

def completitud(texto):
    t = re.sub(r'\*+', '', texto)
    secs = sum(1 for s in ["TÃCNICA", "HALLAZGOS", "IMPRESIÃN"] if s in t.upper())
    words = len(t.split())
    return min(100, int((secs / 3) * 60 + min(words / 150, 1) * 40)), words

def generar_docx(texto):
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    for line in texto.split("\n"):
        s = re.sub(r'\*+', '', line).strip()
        if not s:
            doc.add_paragraph()
            continue
        if s.isupper() and len(s) < 80:
            h = doc.add_heading(s, level=1)
            h.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif s.startswith("â¢"):
            doc.add_paragraph(s[1:].strip(), style="List Bullet")
        else:
            doc.add_paragraph(s)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def build_estilo_context():
    ejemplos = st.session_state.estilo_aprendido
    if not ejemplos:
        return ""
    bloques = []
    for i, e in enumerate(ejemplos[-5:], 1):
        bloques.append(f"â Ejemplo {i} â\n{e['reporte']}")
    return "\n\n".join(bloques)

def hora_saludo():
    h = datetime.datetime.now().hour
    if h < 12:   return "Buenos dÃ­as"
    if h < 19:   return "Buenas tardes"
    return "Buenas noches"

def transcribir(audio):
    cfg = MODELS[st.session_state.modelo]
    cl = OpenAI(api_key=api_key, base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp:
        tmp.write(audio.read()); path = tmp.name
    with open(path, "rb") as f:
        res = cl.audio.transcriptions.create(
            model="whisper-1", file=f, language="es",
            prompt="Dictado radiolÃ³gico: Stoller, ICRS, LCA, menisco, Pfirrmann, Modic, NASCET, BI-RADS, PI-RADS, TIRADS."
        )
    os.unlink(path)
    return res.text.strip()

# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# TEMA
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
DARK = {
    "bg":       "#0e0e11",
    "sidebar":  "#131317",
    "panel":    "#17171c",
    "card":     "#1c1c23",
    "border":   "#26262f",
    "accent":   "#7c6af7",
    "accent2":  "#a89ef9",
    "text":     "#e8e8f0",
    "muted":    "#52526a",
    "muted2":   "#38384e",
    "green":    "#22c55e",
    "amber":    "#f59e0b",
    "red":      "#ef4444",
    "ed_bg":    "#111116",
    "surface":  "#1e1e26",
    "genBtn":   "linear-gradient(135deg,#7c6af7,#5b4de0)",
}
LIGHT = {
    "bg":       "#f4f4f7",
    "sidebar":  "#ffffff",
    "panel":    "#ffffff",
    "card":     "#f0f0f5",
    "border":   "#e0e0ea",
    "accent":   "#5b4de0",
    "accent2":  "#7c6af7",
    "text":     "#18181f",
    "muted":    "#8888a8",
    "muted2":   "#c8c8dc",
    "green":    "#16a34a",
    "amber":    "#d97706",
    "red":      "#dc2626",
    "ed_bg":    "#ffffff",
    "surface":  "#f8f8fc",
    "genBtn":   "linear-gradient(135deg,#5b4de0,#7c6af7)",
}

def th():
    return DARK if st.session_state.tema == "dark" else LIGHT

# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# CSS GLOBAL
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
def css():
    T = th()
    return f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700&family=DM+Serif+Display:ital@0;1&display=swap');

*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
html,body,.stApp{{
  background:{T['bg']};color:{T['text']};
  font-family:'Sora',sans-serif;font-size:13px;
}}
header,footer,#MainMenu,.stDeployButton{{visibility:hidden!important;height:0!important}}
.block-container{{padding:0!important;max-width:100vw!important;overflow:hidden}}
[data-testid="column"]{{padding:0!important}}

/* ââ SIDEBAR NAV ââ */
.sidebar{{
  width:168px;min-width:168px;height:100vh;
  background:{T['sidebar']};
  border-right:1px solid {T['border']};
  display:flex;flex-direction:column;
  padding:0;overflow:hidden;
  position:sticky;top:0;
}}
.sb-logo{{
  padding:18px 18px 14px;
  font-family:'DM Serif Display',serif;
  font-size:22px;color:{T['accent']};
  letter-spacing:.04em;
  border-bottom:1px solid {T['border']};
}}
.sb-greet{{
  padding:8px 18px 14px;
  font-size:11px;color:{T['muted']};
  border-bottom:1px solid {T['border']};
}}
.sb-greet strong{{color:{T['text']};font-weight:500}}
.sb-section{{
  padding:14px 12px 4px;
  font-size:9px;font-weight:600;
  letter-spacing:.18em;text-transform:uppercase;
  color:{T['muted']};
}}
.sb-item{{
  display:flex;align-items:center;gap:9px;
  padding:8px 14px;border-radius:7px;
  margin:1px 6px;cursor:pointer;
  font-size:12px;font-weight:400;
  color:{T['muted']};
  transition:all .15s;
  text-decoration:none;
}}
.sb-item:hover{{background:{T['card']};color:{T['text']}}}
.sb-item.active{{
  background:{T['accent']}18;
  color:{T['accent']};font-weight:500;
}}
.sb-item .badge{{
  margin-left:auto;font-size:9px;
  background:{T['card']};border:1px solid {T['border']};
  border-radius:10px;padding:1px 6px;color:{T['muted']};
}}
.sb-item.active .badge{{background:{T['accent']}22;border-color:{T['accent']}40;color:{T['accent2']}}}
.sb-bottom{{margin-top:auto;padding:12px 6px;border-top:1px solid {T['border']};
  display:flex;gap:6px;align-items:center;}}
.sb-icon-btn{{
  width:28px;height:28px;border-radius:6px;
  background:none;border:1px solid transparent;
  color:{T['muted']};font-size:14px;cursor:pointer;
  display:flex;align-items:center;justify-content:center;
  transition:all .15s;
}}
.sb-icon-btn:hover{{background:{T['card']};color:{T['text']};border-color:{T['border']}}}

/* ââ PANEL HEADER ââ */
.ph{{
  height:44px;display:flex;align-items:center;
  padding:0 16px;gap:10px;
  border-bottom:1px solid {T['border']};
  background:{T['panel']};flex-shrink:0;
}}
.ph-title{{font-size:11px;font-weight:600;letter-spacing:.1em;
  text-transform:uppercase;color:{T['muted']}}}
.ph-chip{{
  font-size:10px;padding:2px 8px;border-radius:4px;
  border:1px solid {T['border']};color:{T['muted']};
  background:{T['card']};cursor:default;
}}
.ph-chip.active{{
  background:{T['accent']}18;border-color:{T['accent']}40;
  color:{T['accent2']};
}}
.ph-right{{margin-left:auto;display:flex;align-items:center;gap:6px}}

/* ââ FORM CONTROLS ââ */
[data-testid="stSelectbox"]>div>div{{
  background:{T['card']}!important;border:1px solid {T['border']}!important;
  border-radius:7px!important;color:{T['text']}!important;
  font-size:12px!important;font-family:'Sora',sans-serif!important;
}}
[data-testid="stSelectbox"]>div>div:hover{{border-color:{T['accent']}55!important}}
.stTextInput input{{
  background:{T['card']}!important;border:1px solid {T['border']}!important;
  border-radius:7px!important;color:{T['text']}!important;
  font-size:12px!important;padding:7px 11px!important;
  font-family:'Sora',sans-serif!important;
}}
.stTextInput input:focus{{border-color:{T['accent']}55!important;box-shadow:none!important}}
.stTextInput input::placeholder{{color:{T['muted']}!important}}
.stTextArea textarea{{
  background:{T['ed_bg']}!important;border:1px solid {T['border']}!important;
  border-radius:8px!important;color:{T['text']}!important;
  font-size:12.5px!important;line-height:1.76!important;
  padding:13px 15px!important;
  font-family:'Sora',sans-serif!important;
  caret-color:{T['accent']}!important;
}}
.stTextArea textarea:focus{{border-color:{T['accent']}45!important;box-shadow:none!important}}
.stTextArea textarea::placeholder{{color:{T['muted']}!important}}
[data-testid="stAudioInput"]{{
  background:{T['card']}!important;border:1px solid {T['border']}!important;
  border-radius:9px!important;
}}
[data-testid="stFileUploader"]{{
  background:{T['card']};border:1px dashed {T['border']};
  border-radius:8px;padding:6px;
}}
[data-testid="stFileUploader"] *{{color:{T['muted']}!important;font-size:11px!important}}

/* ââ BUTTONS ââ */
.stButton button{{
  background:{T['card']}!important;border:1px solid {T['border']}!important;
  color:{T['text']}!important;border-radius:7px!important;
  font-size:12px!important;font-weight:400!important;
  font-family:'Sora',sans-serif!important;
  transition:all .15s!important;
}}
.stButton button:hover{{
  border-color:{T['accent']}55!important;
  background:{T['surface']}!important;color:{T['accent2']}!important;
}}
.btn-gen .stButton button{{
  background:{T['genBtn']}!important;
  border:none!important;color:#fff!important;
  font-weight:600!important;font-size:13px!important;
  border-radius:10px!important;
  box-shadow:0 4px 18px {T['accent']}44!important;
  transition:all .2s!important;
}}
.btn-gen .stButton button:hover{{
  box-shadow:0 6px 24px {T['accent']}66!important;
  transform:translateY(-1px)!important;
  opacity:.95!important;
}}
.stDownloadButton button{{
  background:transparent!important;
  border:1px solid {T['accent']}50!important;
  color:{T['accent']}!important;border-radius:7px!important;
  font-size:12px!important;font-family:'Sora',sans-serif!important;
}}
.stDownloadButton button:hover{{background:{T['accent']}0e!important}}

/* ââ EXPANDER ââ */
[data-testid="stExpander"]{{
  background:{T['card']}!important;border:1px solid {T['border']}!important;
  border-radius:8px!important;margin-bottom:5px!important;
}}
[data-testid="stExpander"] summary{{
  color:{T['muted']}!important;font-size:12px!important;
  padding:9px 13px!important;font-family:'Sora',sans-serif!important;
}}
[data-testid="stExpander"] summary:hover{{color:{T['text']}!important}}

/* ââ TABS ââ */
[data-testid="stTabs"] [role="tablist"]{{
  border-bottom:1px solid {T['border']}!important;
  background:transparent!important;gap:0!important;
}}
[data-testid="stTabs"] [role="tab"]{{
  background:transparent!important;border:none!important;
  color:{T['muted']}!important;font-size:11px!important;
  padding:7px 12px!important;border-bottom:2px solid transparent!important;
  border-radius:0!important;font-family:'Sora',sans-serif!important;
  letter-spacing:.03em;
}}
[data-testid="stTabs"] [role="tab"][aria-selected="true"]{{
  color:{T['accent']}!important;border-bottom-color:{T['accent']}!important;
}}
[data-testid="stTabs"] [data-baseweb="tab-panel"]{{
  background:transparent!important;padding:10px 0 0!important;
}}

/* ââ LABELS ââ */
.lbl{{font-size:9px;font-weight:600;letter-spacing:.16em;text-transform:uppercase;
  color:{T['muted']};margin-bottom:4px;display:block}}
.sec{{font-size:9px;font-weight:600;letter-spacing:.16em;text-transform:uppercase;
  color:{T['muted']};margin-bottom:9px;display:flex;align-items:center;gap:5px}}
.sec::before{{content:'';width:4px;height:4px;border-radius:50%;
  background:{T['accent']};flex-shrink:0}}

/* ââ PROGRESS ââ */
.prog{{display:flex;align-items:center;gap:8px;margin-bottom:6px}}
.prog-t{{flex:1;height:2px;background:{T['border']};border-radius:1px;overflow:hidden}}
.prog-f{{height:100%;background:{T['accent']};border-radius:1px;transition:width .5s}}
.prog-l{{font-size:10px;color:{T['muted']};white-space:nowrap}}

/* ââ HISTORIAL ââ */
.h-item{{display:flex;align-items:center;gap:8px;
  padding:7px 10px;background:{T['surface']};
  border:1px solid {T['border']};border-radius:7px;
  margin-bottom:4px;cursor:pointer;transition:border-color .15s}}
.h-item:hover{{border-color:{T['accent']}45}}
.h-dot{{width:7px;height:7px;border-radius:50%;flex-shrink:0}}
.h-name{{font-size:11px;color:{T['text']}}}
.h-sub{{font-size:10px;color:{T['muted']}}}

/* ââ SUGERENCIA CARD ââ */
.sug-card{{
  background:{T['surface']};
  border:1px solid {T['accent']}28;
  border-left:3px solid {T['accent']};
  border-radius:0 8px 8px 0;
  padding:11px 13px;margin-bottom:8px;
  font-size:11.5px;line-height:1.68;color:{T['text']};
}}
.sug-label{{font-size:9px;font-weight:600;letter-spacing:.14em;
  text-transform:uppercase;color:{T['accent']};margin-bottom:5px}}
.sug-term{{color:{T['accent2']};font-weight:500}}
.sug-ref{{font-size:10px;color:{T['muted']};margin-top:4px;font-style:italic}}

/* ââ DEFS BOX ââ */
.defs-box{{background:{T['ed_bg']};border:1px solid {T['border']};
  border-radius:8px;padding:13px;
  font-size:11.5px;line-height:1.65;color:{T['muted']};white-space:pre-wrap}}

::-webkit-scrollbar{{width:3px;height:3px}}
::-webkit-scrollbar-thumb{{background:{T['border']};border-radius:2px}}
hr{{border:none;border-top:1px solid {T['border']}!important;margin:10px 0!important}}
</style>
"""

st.markdown(css(), unsafe_allow_html=True)
T = th()

# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# LAYOUT PRINCIPAL: sidebar + col_mid + col_right
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# Todas las plantillas disponibles
todas_plantillas = {**PLANTILLAS_DEFAULT, **st.session_state.plantillas_custom}
n_plantillas = len(todas_plantillas) - 1  # sin "Sin plantilla"
n_diccionario = len(st.session_state.diccionario)

col_sb, col_mid, col_right = st.columns([0.18, 0.42, 0.40], gap="small")

# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# SIDEBAR
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
with col_sb:
    nav = st.session_state.nav_active

    st.markdown(f"""
    <div style="background:{T['sidebar']};border-right:1px solid {T['border']};
      min-height:100vh;padding-bottom:20px">
      <div style="padding:18px 18px 12px;border-bottom:1px solid {T['border']}">
        <div style="font-family:'DM Serif Display',serif;font-size:22px;
          color:{T['accent']};letter-spacing:.04em">AURA</div>
      </div>
      <div style="padding:8px 18px 12px;border-bottom:1px solid {T['border']};
        font-size:11px;color:{T['muted']}">
        {hora_saludo()}
      </div>
      <div style="padding:12px 4px 4px;font-size:9px;font-weight:600;
        letter-spacing:.18em;text-transform:uppercase;
        color:{T['muted']};padding-left:14px">Biblioteca</div>
    </div>
    """, unsafe_allow_html=True)

    # Nav items
    def nav_btn(key, label, emoji, badge=None):
        active = st.session_state.nav_active == key
        badge_html = f' <span style="margin-left:auto;font-size:9px;background:{""+T["accent"]+"22" if active else T["card"]};border:1px solid {""+T["accent"]+"40" if active else T["border"]};border-radius:10px;padding:1px 6px;color:{""+T["accent2"] if active else T["muted"]}">{badge}</span>' if badge is not None else ""
        color = T['accent'] if active else T['muted']
        bg = T['accent'] + "18" if active else "transparent"
        st.markdown(f"""
        <div style="display:flex;align-items:center;gap:8px;padding:8px 12px;
          border-radius:7px;margin:1px 4px;background:{bg};cursor:pointer;
          font-size:12px;color:{color};font-weight:{'500' if active else '400'}">
          <span>{emoji}</span><span>{label}</span>{badge_html}
        </div>""", unsafe_allow_html=True)
        if st.button(label, key=f"nav_{key}", use_container_width=True,
                     help=f"Ir a {label}"):
            st.session_state.nav_active = key
            st.rerun()

    st.markdown(f"""
    <style>
    [data-testid="stButton"][key^="nav_"] button{{
      background:transparent!important;border:none!important;
      color:transparent!important;height:0!important;padding:0!important;
      margin-top:-34px!important;opacity:0!important;
    }}
    </style>""", unsafe_allow_html=True)

    # Nav items as invisible buttons overlaid on custom HTML
    items = [
        ("informe",    "Informe",    "ð", None),
        ("plantillas", "Plantillas", "ð", n_plantillas if n_plantillas else None),
        ("diccionario","Diccionario","ð", n_diccionario if n_diccionario else None),
        ("historial",  "Historial",  "ð", len(st.session_state.historial) if st.session_state.historial else None),
        ("config",     "Ajustes",    "âï¸", None),
    ]

    for key, label, emoji, badge in items:
        active = nav == key
        badge_txt = f" ({badge})" if badge else ""
        color = T['accent'] if active else T['muted']
        bg = T['accent'] + "18" if active else "transparent"
        fw = "500" if active else "400"
        st.markdown(f"""
        <div onclick="" style="display:flex;align-items:center;gap:8px;
          padding:8px 12px;border-radius:7px;margin:1px 4px;
          background:{bg};font-size:12px;color:{color};font-weight:{fw};">
          {emoji} {label}{"" if not badge else f' <span style="margin-left:auto;font-size:9px;padding:1px 7px;border-radius:10px;background:{T["accent"]+"22" if active else T["card"]};border:1px solid {T["accent"]+"40" if active else T["border"]};color:{T["accent2"] if active else T["muted"]}">{badge}</span>'}
        </div>""", unsafe_allow_html=True)
        if st.button(f"{label}", key=f"nav_{key}",
                     use_container_width=True):
            st.session_state.nav_active = key
            st.rerun()

    st.markdown("<hr>", unsafe_allow_html=True)

    # Tema toggle
    tema_label = "âï¸ Claro" if st.session_state.tema == "dark" else "ð Oscuro"
    if st.button(tema_label, key="tog_tema", use_container_width=True):
        st.session_state.tema = "light" if st.session_state.tema == "dark" else "dark"
        st.rerun()

    # Modelo badge
    st.markdown(f'<div style="padding:6px 10px;font-size:10px;color:{T["muted"]};'
                f'background:{T["card"]};border:1px solid {T["border"]};'
                f'border-radius:6px;margin-top:4px;text-align:center">'
                f'{st.session_state.modelo}</div>', unsafe_allow_html=True)

# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# PANEL CENTRAL â cambia segÃºn nav_active
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
with col_mid:
    nav = st.session_state.nav_active

    # âââ NAV: INFORME âââââââââââââââââââââââââââââââââââââââ
    if nav == "informe":
        # Header
        pt_nombre = st.session_state.plantilla_nombre
        pt_activa = pt_nombre != "Sin plantilla"
        n_dic = len(st.session_state.diccionario)
        st.markdown(f"""
        <div class="ph">
          <span class="ph-title">TranscripciÃ³n</span>
          <span class="ph-chip {'active' if pt_activa else ''}">
            {'ð ' + pt_nombre[:16] if pt_activa else 'Sin plantilla activa'}
          </span>
          <span class="ph-chip {'active' if n_dic else ''}">
            {'ð ' + str(n_dic) + ' tÃ©rminos' if n_dic else 'Sin diccionario activo'}
          </span>
          <div class="ph-right">
            <button onclick="" style="background:none;border:none;
              font-size:12px;color:{T['muted']};cursor:pointer"
              title="Nuevo estudio">+ Nuevo</button>
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<div style='padding:14px 14px 0'>", unsafe_allow_html=True)

        # SelecciÃ³n de modalidad y regiÃ³n
        with st.expander("ð¬  Configurar estudio", expanded=False):
            st.markdown('<span class="lbl">Modalidad</span>', unsafe_allow_html=True)
            st.selectbox("mod", MODALIDADES, label_visibility="collapsed", key="sel_mod")
            c1, c2 = st.columns(2)
            with c1:
                st.markdown('<span class="lbl">Grupo</span>', unsafe_allow_html=True)
                grupo = st.selectbox("grp", list(REGIONES.keys()),
                                     label_visibility="collapsed", key="sel_grupo")
            with c2:
                st.markdown('<span class="lbl">RegiÃ³n</span>', unsafe_allow_html=True)
                st.selectbox("reg", REGIONES[grupo],
                             label_visibility="collapsed", key="sel_reg")
            st.markdown('<span class="lbl" style="margin-top:5px">RegiÃ³n libre</span>',
                        unsafe_allow_html=True)
            st.text_input("rc", label_visibility="collapsed", key="reg_custom",
                          placeholder="Ej: ArticulaciÃ³n glenohumeral derecha")

        st.markdown("<hr style='margin:10px 0'>", unsafe_allow_html=True)

        # Tabs: Dictado / Teclado
        tab_voz, tab_kbd = st.tabs(["ð Dictado", "â¨ Teclado"])

        with tab_voz:
            st.markdown(f"""
            <div style="display:flex;flex-direction:column;align-items:center;
              padding:24px 0 16px;gap:10px">
              <div style="position:relative;width:82px;height:82px">
                <div style="position:absolute;inset:-16px;border-radius:50%;
                  border:1.5px solid {T['accent']}22;
                  animation:rp 2.6s ease-out infinite"></div>
                <div style="position:absolute;inset:-7px;border-radius:50%;
                  border:1.5px solid {T['accent']}38;
                  animation:rp 2.6s ease-out infinite .65s"></div>
                <div style="width:82px;height:82px;border-radius:50%;
                  background:{T['accent']};
                  display:flex;align-items:center;justify-content:center;
                  box-shadow:0 8px 28px {T['accent']}55;cursor:pointer">
                  <svg width="30" height="30" viewBox="0 0 24 24" fill="none"
                    stroke="#fff" stroke-width="1.8" stroke-linecap="round">
                    <rect x="9" y="2" width="6" height="12" rx="3"/>
                    <path d="M5 10a7 7 0 0 0 14 0"/>
                    <line x1="12" y1="19" x2="12" y2="22"/>
                    <line x1="9" y1="22" x2="15" y2="22"/>
                  </svg>
                </div>
              </div>
              <div style="text-align:center">
                <div style="font-size:13px;font-weight:500;color:{T['text']}">
                  Pulsa el botÃ³n para comenzar a dictar</div>
                <div style="font-size:11px;color:{T['muted']};margin-top:3px">
                  <span style="color:{T['accent']};font-weight:500">AURA</span>
                  interpretarÃ¡ el audio al generar el informe.
                </div>
              </div>
            </div>
            <style>@keyframes rp{{
              0%{{transform:scale(1);opacity:.5}}
              100%{{transform:scale(1.5);opacity:0}}
            }}</style>
            """, unsafe_allow_html=True)
            audio = st.audio_input("rec", label_visibility="collapsed")
            if audio:
                aid = hash(audio.read()); audio.seek(0)
                if aid != st.session_state.audio_id:
                    if api_key:
                        with st.spinner("Transcribiendo..."):
                            txt = transcribir(audio)
                        if txt:
                            st.session_state.dictado += (" " + txt).strip()
                            st.session_state.audio_id = aid
                            st.rerun()
                    else:
                        st.warning("Configura tu API Key en Ajustes.")

            if st.session_state.dictado:
                st.markdown('<span class="lbl" style="margin-top:10px">TranscripciÃ³n</span>',
                            unsafe_allow_html=True)
                d = st.text_area("dt", value=st.session_state.dictado,
                                 height=120, label_visibility="collapsed",
                                 key="dictado_ta_voz",
                                 placeholder="La transcripciÃ³n aparece aquÃ­â¦")
                if d != st.session_state.dictado:
                    st.session_state.dictado = d

        with tab_kbd:
            st.markdown('<span class="lbl" style="margin-top:6px">Hallazgos</span>',
                        unsafe_allow_html=True)
            d = st.text_area("dk", value=st.session_state.dictado,
                             height=200, label_visibility="collapsed",
                             key="dictado_ta_kbd",
                             placeholder="Escribe los hallazgos directamente.\n\nEj: Desgarro horizontal menisco medial cuerno posterior, Stoller III, extrusiÃ³n 3 mm. Ligamento cruzado anterior Ã­ntegroâ¦")
            if d != st.session_state.dictado:
                st.session_state.dictado = d

        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

        # BOTÃN GENERAR
        c_gen, c_clr = st.columns([3, 1])
        with c_gen:
            st.markdown('<div class="btn-gen">', unsafe_allow_html=True)
            generar = st.button("â¦  Generar Informe", key="btn_gen",
                                use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c_clr:
            if st.button("âº Nuevo", use_container_width=True, key="btn_clr"):
                st.session_state.dictado = ""
                st.session_state.reporte = ""
                st.session_state.audio_id = None
                st.session_state.mentor_feedback = ""
                st.session_state.sugerencias_activas = []
                st.rerun()

    # âââ NAV: PLANTILLAS ââââââââââââââââââââââââââââââââââââ
    elif nav == "plantillas":
        st.markdown(f"""
        <div class="ph">
          <span class="ph-title">Plantillas</span>
          <span style="font-size:11px;color:{T['muted']}">{len(todas_plantillas)-1} disponibles</span>
        </div>""", unsafe_allow_html=True)
        st.markdown("<div style='padding:14px'>", unsafe_allow_html=True)

        st.markdown('<span class="lbl">Seleccionar plantilla activa</span>',
                    unsafe_allow_html=True)
        sel = st.selectbox("plt_sel", list(todas_plantillas.keys()),
                           index=list(todas_plantillas.keys()).index(
                               st.session_state.plantilla_nombre)
                           if st.session_state.plantilla_nombre in todas_plantillas else 0,
                           label_visibility="collapsed", key="plt_sel_box")
        if sel != st.session_state.plantilla_nombre:
            st.session_state.plantilla_nombre = sel
            st.session_state.plantilla_txt = todas_plantillas[sel]
            st.rerun()

        if todas_plantillas[sel]:
            st.markdown(f"""
            <div style="background:{T['ed_bg']};border:1px solid {T['border']};
              border-radius:8px;padding:13px;margin-top:8px;
              font-size:11.5px;line-height:1.7;color:{T['muted']};
              white-space:pre-wrap;max-height:280px;overflow-y:auto">
              {todas_plantillas[sel]}
            </div>""", unsafe_allow_html=True)

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<span class="lbl">Cargar plantilla .docx</span>',
                    unsafe_allow_html=True)
        f_up = st.file_uploader("plt_up", type=["docx"],
                                label_visibility="collapsed", key="plt_uploader")
        if f_up:
            contenido = leer_plantilla_docx(f_up)
            nombre = f_up.name.replace(".docx", "")
            st.session_state.plantillas_custom[nombre] = contenido
            st.session_state.plantilla_nombre = nombre
            st.session_state.plantilla_txt = contenido
            st.success(f"â Plantilla '{nombre}' cargada")
            st.rerun()

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<span class="lbl">Nueva plantilla manual</span>',
                    unsafe_allow_html=True)
        nuevo_nombre = st.text_input("plt_nombre", label_visibility="collapsed",
                                     placeholder="Nombre de la plantilla",
                                     key="plt_nombre_input")
        nuevo_contenido = st.text_area("plt_contenido",
                                       height=120, label_visibility="collapsed",
                                       placeholder="Escribe la estructura de la plantillaâ¦",
                                       key="plt_contenido_input")
        if st.button("Guardar plantilla", key="btn_save_plt", use_container_width=True):
            if nuevo_nombre and nuevo_contenido:
                st.session_state.plantillas_custom[nuevo_nombre] = nuevo_contenido
                st.success(f"â Guardada: {nuevo_nombre}")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
        generar = False

    # âââ NAV: DICCIONARIO âââââââââââââââââââââââââââââââââââ
    elif nav == "diccionario":
        st.markdown(f"""
        <div class="ph">
          <span class="ph-title">Diccionario radiolÃ³gico</span>
          <span style="font-size:11px;color:{T['muted']}">{n_diccionario} tÃ©rminos</span>
        </div>""", unsafe_allow_html=True)
        st.markdown("<div style='padding:14px'>", unsafe_allow_html=True)
        st.markdown(f"""
        <p style="font-size:11px;color:{T['muted']};line-height:1.7;margin-bottom:10px">
        Define tus tÃ©rminos preferidos, abreviaturas y estilo de redacciÃ³n.
        AURA los utilizarÃ¡ como referencia al generar informes.
        </p>""", unsafe_allow_html=True)

        st.markdown('<span class="lbl">Agregar tÃ©rmino</span>', unsafe_allow_html=True)
        c_t, c_d = st.columns([1, 2])
        with c_t:
            nuevo_term = st.text_input("dic_term", label_visibility="collapsed",
                                       placeholder="TÃ©rmino", key="dic_term_input")
        with c_d:
            nuevo_def = st.text_input("dic_def", label_visibility="collapsed",
                                      placeholder="DefiniciÃ³n / instrucciÃ³n de uso",
                                      key="dic_def_input")
        if st.button("Agregar", key="btn_add_dic", use_container_width=True):
            if nuevo_term and nuevo_def:
                st.session_state.diccionario[nuevo_term] = nuevo_def
                st.rerun()

        st.markdown("<hr>", unsafe_allow_html=True)
        if st.session_state.diccionario:
            for term, defn in list(st.session_state.diccionario.items()):
                c1, c2, c3 = st.columns([1.2, 2.5, 0.4])
                with c1:
                    st.markdown(f'<span style="font-size:11px;color:{T["accent2"]};'
                                f'font-weight:500">{term}</span>', unsafe_allow_html=True)
                with c2:
                    st.markdown(f'<span style="font-size:11px;color:{T["muted"]}">{defn}</span>',
                                unsafe_allow_html=True)
                with c3:
                    if st.button("â", key=f"del_dic_{term}"):
                        del st.session_state.diccionario[term]; st.rerun()
        else:
            st.markdown(f'<p style="font-size:11px;color:{T["muted"]};text-align:center;'
                        f'padding:20px 0">Sin tÃ©rminos definidos aÃºn.</p>',
                        unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        generar = False

    # âââ NAV: HISTORIAL âââââââââââââââââââââââââââââââââââââ
    elif nav == "historial":
        st.markdown(f"""
        <div class="ph">
          <span class="ph-title">Historial</span>
          <span style="font-size:11px;color:{T['muted']}">{len(st.session_state.historial)} estudios</span>
        </div>""", unsafe_allow_html=True)
        st.markdown("<div style='padding:14px'>", unsafe_allow_html=True)
        if st.session_state.historial:
            for i, e in enumerate(st.session_state.historial):
                color = HCOLS[i % len(HCOLS)]
                c1, c2 = st.columns([3, 1])
                with c1:
                    st.markdown(f"""
                    <div class="h-item">
                      <div class="h-dot" style="background:{color}"></div>
                      <div>
                        <div class="h-name">{e['region']}</div>
                        <div class="h-sub">{e['modalidad'][:22]} Â· {e.get('fecha','')}</div>
                      </div>
                    </div>""", unsafe_allow_html=True)
                with c2:
                    if st.button("Cargar", key=f"h_load_{i}",
                                 use_container_width=True):
                        st.session_state.reporte = e['texto']
                        st.session_state.nav_active = "informe"
                        st.rerun()
        else:
            st.markdown(f'<p style="font-size:11px;color:{T["muted"]};text-align:center;'
                        f'padding:30px 0">No hay estudios en el historial.</p>',
                        unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        generar = False

    # âââ NAV: AJUSTES âââââââââââââââââââââââââââââââââââââââ
    elif nav == "config":
        st.markdown(f"""
        <div class="ph"><span class="ph-title">Ajustes</span></div>
        """, unsafe_allow_html=True)
        st.markdown("<div style='padding:14px'>", unsafe_allow_html=True)

        st.markdown('<span class="lbl">Modelo IA</span>', unsafe_allow_html=True)
        m = st.selectbox("cfg_model", list(MODELS.keys()),
                         index=list(MODELS.keys()).index(st.session_state.modelo),
                         label_visibility="collapsed", key="cfg_model_sel")
        if m != st.session_state.modelo:
            st.session_state.modelo = m; st.rerun()

        if not api_key:
            st.markdown('<span class="lbl" style="margin-top:8px">API Key</span>',
                        unsafe_allow_html=True)
            api_key = st.text_input("cfg_key", type="password",
                                    label_visibility="collapsed",
                                    placeholder="sk- Â·Â·Â·", key="cfg_key_input")

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<span class="lbl">Estilo aprendido</span>', unsafe_allow_html=True)
        n_ej = len(st.session_state.estilo_aprendido)
        st.markdown(f'<p style="font-size:11px;color:{T["muted"]};margin-bottom:8px">'
                    f'{n_ej} ejemplo{"s" if n_ej!=1 else ""} guardado{"s" if n_ej!=1 else ""}.</p>',
                    unsafe_allow_html=True)
        if n_ej:
            if st.button("ð Borrar ejemplos de estilo", use_container_width=True,
                         key="btn_del_estilo"):
                st.session_state.estilo_aprendido = []; st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)
        generar = False
    else:
        generar = False

# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# PROCESAMIENTO IA
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
if generar:
    if not api_key:
        st.warning("Configura tu API Key en Ajustes.")
    elif not st.session_state.dictado.strip():
        st.warning("Escribe o dicta los hallazgos primero.")
    else:
        cl  = get_client()
        mid = MODELS[st.session_state.modelo]["id"]
        pt  = st.session_state.plantilla_txt

        mod_sel = st.session_state.get("sel_mod", "")
        reg_sel = (st.session_state.get("reg_custom", "").strip()
                   or st.session_state.get("sel_reg", ""))

        # InstrucciÃ³n de plantilla
        if pt:
            instruc_plantilla = f"""PLANTILLA ESTRUCTURAL OBLIGATORIA
Debes respetar EXACTAMENTE esta estructura, secciones y orden.
Completa el contenido de cada secciÃ³n con los hallazgos del dictado.
No aÃ±adas ni elimines secciones. Si una secciÃ³n no tiene datos en el dictado,
escrÃ­bela con "Sin alteraciones relevantes." o la frase apropiada para esa secciÃ³n.

{pt}"""
        else:
            instruc_plantilla = """ESTRUCTURA:
INDICACIÃN
TÃCNICA
HALLAZGOS
IMPRESIÃN DIAGNÃSTICA"""

        # InstrucciÃ³n de diccionario
        dic = st.session_state.diccionario
        instruc_dic = ""
        if dic:
            terminos = "\n".join([f"Â· {k}: {v}" for k, v in dic.items()])
            instruc_dic = f"""
DICCIONARIO PERSONAL DEL RADIÃLOGO (aplica siempre):
{terminos}
"""

        # Estilo aprendido
        estilo_ctx = build_estilo_context()
        instruc_estilo = ""
        if estilo_ctx:
            instruc_estilo = f"""
ESTILO PERSONAL APRENDIDO:
Replika fielmente el estilo narrativo, nivel de detalle y terminologÃ­a de estos informes previos aprobados:

{estilo_ctx}
"""

        # InstrucciÃ³n tabla
        instruc_tabla = (
            "La plantilla tiene tablas [TABLA]. ComplÃ©talas con datos del dictado."
            if "[TABLA" in pt else
            "No uses tablas salvo que sean estrictamente necesarias."
        )

        prompt = f"""Eres AURA, sistema experto de interpretaciÃ³n radiolÃ³gica de nivel subespecialista.
ActÃºas como: radiÃ³logo subespecialista senior + editor acadÃ©mico mÃ©dico + copiloto de redacciÃ³n.

MODALIDAD: {mod_sel}
REGIÃN: {reg_sel}

ââââââââââââââââââââââââââââââââââââââ
FORMATO DE SALIDA â ABSOLUTAMENTE OBLIGATORIO
ââââââââââââââââââââââââââââââââââââââ
1. CERO asteriscos (*). Absolutamente prohibido.
2. CERO markdown (sin #, sin **, sin *).
3. TÃ­tulos de secciÃ³n: MAYÃSCULAS puras, solos en su lÃ­nea.
4. HALLAZGOS: pÃ¡rrafos narrativos, prosa corrida, oraciones completas.
   NUNCA uses guiones, viÃ±etas o listas en HALLAZGOS.
   Los hallazgos deben fluir como texto mÃ©dico de alta calidad, no como checklist.
   Conecta estructuras anatÃ³micas, establece relaciones, describe hallazgos en contexto.
5. IMPRESIÃN DIAGNÃSTICA: usa â¢ para cada viÃ±eta jerarquizada.
6. Lenguaje: voz activa, tiempo presente, oraciones precisas y elegantes.
   Nivel: publicable en revista indexada, auditable por comitÃ© de pares.
7. Cuantifica siempre: mm, %, grados, scores.

ââââââââââââââââââââââââââââââââââââââ
CLASIFICACIONES APLICABLES
ââââââââââââââââââââââââââââââââââââââ
Aplica SOLO las respaldadas por el dictado:
MeniscoâStoller(I-III), CartÃ­lagoâICRS/Outerbridge, LCAâgrado por continuidad,
ColumnaâPfirrmann/Modic/Meyerding, HombroâBigliani/Goutallier/Sugaya,
CaderaâTÃ¶nnis/alpha-angle, CerebroâASPECTS/Fazekas,
MamaâBI-RADS, PrÃ³stataâPI-RADS, TiroidesâTIRADS.
{instruc_dic}
{instruc_estilo}
{instruc_plantilla}

DICTADO DEL RADIÃLOGO:
{st.session_state.dictado}

{instruc_tabla}"""

        with st.spinner("Generando informe..."):
            try:
                res = cl.chat.completions.create(
                    model=mid,
                    messages=[{"role": "system", "content": prompt}],
                    temperature=0.12, max_tokens=3000
                )
                report = limpiar_md(res.choices[0].message.content)
                st.session_state.reporte = report
                st.session_state.mentor_feedback = ""
                st.session_state.sugerencias_activas = []
                fecha = datetime.datetime.now().strftime("%d/%m %H:%M")
                st.session_state.historial.insert(0, {
                    "modalidad": mod_sel[:20] if mod_sel else "RM",
                    "region":    reg_sel      if reg_sel  else "General",
                    "texto":     report,
                    "fecha":     fecha,
                })
                if len(st.session_state.historial) > 20:
                    st.session_state.historial = st.session_state.historial[:20]
                st.rerun()
            except Exception as e:
                st.error(str(e))

# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# PANEL DERECHO â Editor de informe
# âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
with col_right:
    rep = st.session_state.reporte
    T = th()

    # Header del panel informe
    pct, words = completitud(rep) if rep else (0, 0)
    pt_label = st.session_state.plantilla_nombre
    has_pt = pt_label != "Sin plantilla"

    st.markdown(f"""
    <div class="ph">
      <span class="ph-title">Informe</span>
      {'<span class="ph-chip active">ð ' + pt_label[:14] + '</span>' if has_pt else ''}
      <div class="ph-right">
        <div style="display:flex;align-items:center;gap:6px">
          <div style="width:60px;height:2px;background:{T['border']};border-radius:1px;overflow:hidden">
            <div style="width:{pct}%;height:100%;background:{T['accent']};border-radius:1px;transition:width .4s"></div>
          </div>
          <span style="font-size:10px;color:{T['muted']}">{pct}%</span>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ââ Editor rico via components.html ââ
    def text_to_editor_html(texto):
        if not texto:
            return ""
        html_parts = []
        for line in texto.split("\n"):
            s = line.strip()
            if not s:
                html_parts.append("<p><br></p>")
            elif s.isupper() and 2 < len(s) < 75:
                html_parts.append(f"<h2>{s}</h2>")
            elif s.startswith("â¢"):
                html_parts.append(f"<li>{s[1:].strip()}</li>")
            else:
                html_parts.append(f"<p>{s}</p>")
        return "\n".join(html_parts)

    contenido = text_to_editor_html(rep)
    T = th()

    editor_h = 520

    editor_html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600&family=DM+Serif+Display:ital@0;1&display=swap" rel="stylesheet">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{
  width:100%;
  height:{editor_h + 76}px;
  display:flex;flex-direction:column;
  background:{T['bg']};
  font-family:'Sora',sans-serif;
  overflow:hidden;
}}

/* TOOLBAR */
.tb{{
  flex-shrink:0;height:36px;
  background:{T['panel']};
  border-bottom:1px solid {T['border']};
  padding:0 8px;
  display:flex;align-items:center;gap:2px;
  overflow-x:auto;
}}
.tb::-webkit-scrollbar{{height:2px}}
.tb::-webkit-scrollbar-thumb{{background:{T['border']}}}
.tg{{
  display:flex;align-items:center;gap:1px;
  padding-right:5px;margin-right:3px;
  border-right:1px solid {T['border']};
  flex-shrink:0;
}}
.tg:last-child{{border-right:none}}
.tb-btn{{
  width:24px;height:24px;border:1px solid transparent;
  background:none;color:{T['muted']};font-size:11px;
  border-radius:4px;cursor:pointer;transition:all .1s;
  display:flex;align-items:center;justify-content:center;
  font-family:'Sora',sans-serif;
}}
.tb-btn:hover{{background:{T['card']};color:{T['text']};border-color:{T['border']}}}
.tb-btn.on{{background:{T['accent']}20;color:{T['accent']};border-color:{T['accent']}44}}
.tb-sel{{
  height:24px;background:{T['card']};border:1px solid {T['border']};
  color:{T['muted']};font-size:10px;border-radius:4px;
  padding:0 4px;outline:none;cursor:pointer;
  font-family:'Sora',sans-serif;
}}
.tb-sel:focus{{border-color:{T['accent']}55}}

/* SCROLL AREA */
.scroll{{
  flex:1;overflow-y:auto;overflow-x:hidden;
  background:{T['card']};
  min-height:0;
}}
.scroll::-webkit-scrollbar{{width:3px}}
.scroll::-webkit-scrollbar-thumb{{background:{T['border']};border-radius:2px}}
.scroll::-webkit-scrollbar-thumb:hover{{background:{T['accent']}55}}

/* PAPER */
.paper{{
  min-height:100%;
  padding:28px 32px 40px;
  background:{T['ed_bg']};
  outline:none;
  font-family:'Sora',sans-serif;
  font-size:12.5px;
  line-height:1.84;
  color:{T['text']};
  word-break:break-word;
}}
.paper:empty::before{{
  content:'El informe aparecerÃ¡ aquÃ­.\\ADicta los hallazgos y presiona Generar Informe.';
  color:{T['muted']};white-space:pre;pointer-events:none;display:block;
  text-align:center;padding-top:60px;font-size:12px;
}}
.paper h2{{
  font-family:'Sora',sans-serif;
  font-size:11px;font-weight:600;
  letter-spacing:.16em;text-transform:uppercase;
  color:{T['accent']};
  margin:22px 0 8px;
  padding-bottom:6px;
  border-bottom:1px solid {T['border']};
}}
.paper p{{margin:1px 0;}}
.paper li{{margin-left:16px;margin-bottom:3px}}
.paper ul,
.paper ol{{margin:4px 0 4px 16px}}
.paper hr{{border:none;border-top:1px solid {T['border']};margin:12px 0}}
.paper table{{border-collapse:collapse;width:100%;margin:10px 0;font-size:12px}}
.paper td,.paper th{{border:1px solid {T['border']};padding:6px 10px;color:{T['text']}}}
.paper th{{background:{T['surface']};font-weight:600;
  color:{T['accent']};font-size:10px;letter-spacing:.08em}}
.paper tr:nth-child(even) td{{background:{T['surface']}60}}

/* EMPTY STATE */
.empty-state{{
  display:flex;flex-direction:column;align-items:center;
  justify-content:center;height:100%;min-height:300px;
  gap:10px;opacity:.6;
}}
.empty-icon{{
  width:44px;height:44px;border-radius:12px;
  background:{T['accent']}18;border:1px solid {T['accent']}30;
  display:flex;align-items:center;justify-content:center;
  font-size:18px;
}}

/* STATUS BAR */
.sbar{{
  flex-shrink:0;height:28px;
  background:{T['panel']};border-top:1px solid {T['border']};
  display:flex;align-items:center;padding:0 10px;gap:8px;
}}
.sbar-txt{{font-size:10px;color:{T['muted']}}}
.sbar-track{{width:80px;height:2px;background:{T['border']};border-radius:1px;overflow:hidden}}
.sbar-fill{{height:100%;background:{T['accent']};border-radius:1px;transition:width .4s}}
.sbar-pct{{font-size:10px;color:{T['muted']}}}
</style>
</head>
<body>

<!-- TOOLBAR -->
<div class="tb">
  <div class="tg">
    <select class="tb-sel" style="width:76px" onchange="setFont(this.value)">
      <option value="'Sora',sans-serif" selected>Sora</option>
      <option value="'DM Serif Display',serif">DM Serif</option>
      <option value="Georgia,serif">Georgia</option>
      <option value="Calibri,sans-serif">Calibri</option>
      <option value="'Courier New',monospace">Courier</option>
    </select>
    <select class="tb-sel" style="width:36px" onchange="setSize(this.value)">
      <option>10</option><option>11</option><option>12</option>
      <option selected>13</option><option>14</option>
      <option>15</option><option>16</option><option>18</option>
    </select>
  </div>
  <div class="tg">
    <button class="tb-btn" id="bB" onclick="fmt('bold')" title="Negrita"><b>B</b></button>
    <button class="tb-btn" id="bI" onclick="fmt('italic')" title="Cursiva"><i>I</i></button>
    <button class="tb-btn" id="bU" onclick="fmt('underline')" title="Subrayado"><u>U</u></button>
    <button class="tb-btn" onclick="fmt('strikeThrough')" title="Tachado" style="text-decoration:line-through;font-size:10px">S</button>
  </div>
  <div class="tg">
    <button class="tb-btn" onclick="fmt('justifyLeft')"   title="Izq">&#8676;</button>
    <button class="tb-btn" onclick="fmt('justifyCenter')" title="Centro">&#9644;</button>
    <button class="tb-btn" onclick="fmt('justifyRight')"  title="Der">&#8677;</button>
    <button class="tb-btn" onclick="fmt('justifyFull')"   title="Justificado">&#9776;</button>
  </div>
  <div class="tg">
    <button class="tb-btn" onclick="fmt('insertUnorderedList')" title="ViÃ±etas">&#8226;</button>
    <button class="tb-btn" onclick="fmt('insertOrderedList')"   title="Num">1.</button>
    <button class="tb-btn" onclick="insHR()" title="Separador" style="font-size:9px">HR</button>
    <button class="tb-btn" onclick="insTable()" title="Tabla" style="font-size:9px">Tbl</button>
  </div>
  <div class="tg" style="gap:3px">
    <label title="Color texto" style="display:flex;align-items:center;cursor:pointer">
      <span style="font-size:9px;color:{T['muted']};margin-right:2px">A</span>
      <input type="color" value="{T['text']}" onchange="fmt('foreColor',this.value)"
        style="width:16px;height:16px;padding:0;border:none;border-radius:3px;cursor:pointer">
    </label>
    <label title="Resaltar" style="display:flex;align-items:center;cursor:pointer">
      <span style="font-size:9px;color:{T['muted']};margin-right:2px">HL</span>
      <input type="color" value="{T['accent']}" onchange="fmt('hiliteColor',this.value)"
        style="width:16px;height:16px;padding:0;border:none;border-radius:3px;cursor:pointer">
    </label>
  </div>
  <div class="tg">
    <button class="tb-btn" onclick="copyAll()" title="Copiar">&#10697;</button>
    <button class="tb-btn" onclick="printDoc()" title="Imprimir/PDF">&#128438;</button>
    <button class="tb-btn" onclick="document.execCommand('undo')" title="Deshacer">&#8617;</button>
    <button class="tb-btn" onclick="document.execCommand('redo')" title="Rehacer">&#8618;</button>
  </div>
</div>

<!-- EDITOR -->
<div class="scroll" id="scroll">
  <div class="paper" id="paper" contenteditable="true" spellcheck="false"
    oninput="syncBar()" onkeyup="syncState()" onmouseup="syncState()">
    {contenido if contenido else ''}
  </div>
</div>

<!-- STATUS BAR -->
<div class="sbar">
  <span class="sbar-txt" id="wc">0 palabras</span>
  <div style="margin-left:auto;display:flex;align-items:center;gap:6px">
    <div class="sbar-track"><div class="sbar-fill" id="sf" style="width:{pct}%"></div></div>
    <span class="sbar-pct" id="sp">{pct}%</span>
  </div>
</div>

<script>
var paper = document.getElementById('paper');

function fmt(cmd,val){{
  paper.focus();
  document.execCommand(cmd,false,val||null);
  syncState();
}}
function setFont(f){{paper.style.fontFamily=f}}
function setSize(s){{paper.style.fontSize=s+'px'}}
function insHR(){{
  paper.focus();
  document.execCommand('insertHTML',false,
    '<hr style="border:none;border-top:1px solid {T["border"]};margin:12px 0"><br>');
}}
function insTable(){{
  var r=parseInt(prompt('Filas:','3'))||3;
  var c=parseInt(prompt('Columnas:','3'))||3;
  var h='<table><thead><tr>';
  for(var i=0;i<c;i++) h+='<th>Col '+(i+1)+'</th>';
  h+='</tr></thead><tbody>';
  for(var j=0;j<r-1;j++){{
    h+='<tr>';
    for(var k=0;k<c;k++) h+='<td>&nbsp;</td>';
    h+='</tr>';
  }}
  h+='</tbody></table><p><br></p>';
  paper.focus();
  document.execCommand('insertHTML',false,h);
}}
function syncBar(){{
  var t=paper.innerText||'';
  var w=t.trim().split(/[ \\t\\n]+/).filter(Boolean).length;
  var secs=['TÃCNICA','HALLAZGOS','IMPRESIÃN'].filter(function(s){{
    return t.toUpperCase().includes(s);
  }}).length;
  var p=Math.min(100,Math.round((secs/3)*60+Math.min(w/150,1)*40));
  document.getElementById('wc').textContent=w+' palabras';
  document.getElementById('sf').style.width=p+'%';
  document.getElementById('sp').textContent=p+'%';
}}
function syncState(){{
  syncBar();
  ['Bold','Italic','Underline'].forEach(function(c){{
    var b=document.getElementById('b'+c[0]);
    if(b) b.classList.toggle('on',document.queryCommandState(c.toLowerCase()));
  }});
}}
function copyAll(){{
  var text=paper.innerText;
  if(navigator.clipboard){{
    navigator.clipboard.writeText(text).then(function(){{toast('â Copiado')}});
  }} else {{
    var ta=document.createElement('textarea');
    ta.value=text;ta.style.cssText='position:fixed;opacity:0';
    document.body.appendChild(ta);ta.select();
    document.execCommand('copy');document.body.removeChild(ta);
    toast('â Copiado');
  }}
}}
function printDoc(){{
  var w=window.open('','_blank');
  w.document.write('<html><head><title>AURA â Informe</title>');
  w.document.write('<style>body{{font-family:Calibri,sans-serif;font-size:12pt;');
  w.document.write('line-height:1.75;margin:2.5cm;color:#111}}');
  w.document.write('h2{{font-size:11pt;font-weight:600;text-transform:uppercase;');
  w.document.write('letter-spacing:.1em;border-bottom:1px solid #ccc;');
  w.document.write('padding-bottom:4px;margin:18px 0 7px;color:#111}}');
  w.document.write('li{{margin-left:18px;margin-bottom:3px}}');
  w.document.write('table{{border-collapse:collapse;width:100%;margin:10px 0}}');
  w.document.write('td,th{{border:1px solid #ccc;padding:6px 10px}}');
  w.document.write('th{{background:#f4f4f4;font-weight:600}}');
  w.document.write('</style></head><body>');
  w.document.write(paper.innerHTML);
  w.document.write('</body></html>');
  w.document.close();
  setTimeout(function(){{w.print()}},350);
}}
function toast(msg){{
  var el=document.createElement('div');
  el.textContent=msg;
  el.style.cssText='position:fixed;bottom:36px;left:50%;transform:translateX(-50%);'
    +'background:{T["surface"]};color:{T["accent"]};'
    +'border:1px solid {T["accent"]}55;'
    +'padding:5px 14px;border-radius:5px;font-size:11px;z-index:9999;'
    +'pointer-events:none;font-family:Sora,sans-serif';
  document.body.appendChild(el);
  setTimeout(function(){{document.body.removeChild(el)}},1600);
}}
paper.addEventListener('keydown',function(e){{
  if(e.key==='Tab'){{
    e.preventDefault();
    document.execCommand('insertHTML',false,'&nbsp;&nbsp;&nbsp;&nbsp;');
  }}
}});
window.addEventListener('load',function(){{syncBar();}});
</script>
</body></html>"""

    components.html(editor_html, height=editor_h + 76, scrolling=False)

    # ââ Acciones ââââââââââââââââââââââââââââââââââââââââââââ
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    ca, cb, cc, cd = st.columns(4)

    with ca:
        if st.button("â¦ Optimizar", use_container_width=True, key="btn_opt",
                     help="Optimizar conclusiÃ³n"):
            if rep and api_key:
                cl = get_client(); mid = MODELS[st.session_state.modelo]["id"]
                estilo_ctx = build_estilo_context()
                estilo_note = f"\nESTILO (aplica):\n{estilo_ctx}" if estilo_ctx else ""
                with st.spinner("Optimizandoâ¦"):
                    try:
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role": "user", "content":
                                f"""RadiÃ³logo subespecialista senior. Mejora ÃNICAMENTE la IMPRESIÃN DIAGNÃSTICA.
REGLAS: CERO asteriscos. Sin markdown. TÃ­tulos MAYÃSCULAS. ViÃ±etas con â¢.
Devuelve informe COMPLETO. No alteres TÃCNICA ni HALLAZGOS. Prosa corrida en hallazgos.
CRITERIOS: jerarquÃ­a diagnÃ³stica, grado+implicaciÃ³n en cada â¢, Ãºltima â¢=manejo.
{estilo_note}
INFORME:\n{rep}"""}],
                            temperature=0.15, max_tokens=3000
                        )
                        st.session_state.reporte = limpiar_md(r.choices[0].message.content)
                        st.rerun()
                    except Exception as e:
                        st.error(str(e))

    with cb:
        if st.button("â Mentor", use_container_width=True, key="btn_mentor",
                     help="AnÃ¡lisis editorial del informe"):
            if rep and api_key:
                cl = get_client(); mid = MODELS[st.session_state.modelo]["id"]
                with st.spinner("Analizandoâ¦"):
                    try:
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role": "user", "content":
                                f"""Eres un mentor de redacciÃ³n radiolÃ³gica de Ã©lite.
Analiza este informe. EspaÃ±ol. Sin asteriscos ni markdown.

EVALUACIÃN EDITORIAL (hasta 5 puntos):
Para cada punto: QUÃ Â· POR QUÃ Â· CÃMO MEJORAR (con reescritura sugerida)
Detecta: prosa dÃ©bil, listas donde deberÃ­a haber narrativa, clasificaciones ausentes,
redundancias, hedge words innecesarios, conclusiones no accionables.

NIVEL ACTUAL: [BÃ¡sico/Residente/Fellow/Subespecialista/Publicable]
FORTALEZAS: [2-3 lÃ­neas]
POTENCIAL: [2-3 lÃ­neas especÃ­ficas]
PUNTUACIÃN: [X/10] â [justificaciÃ³n en una lÃ­nea]

INFORME:\n{rep}"""}],
                            temperature=0.2, max_tokens=2000
                        )
                        st.session_state.mentor_feedback = r.choices[0].message.content
                        st.rerun()
                    except Exception as e:
                        st.error(str(e))

    with cc:
        if st.button("ð§  Aprender", use_container_width=True, key="btn_learn",
                     help="Guardar como ejemplo de mi estilo"):
            if rep:
                st.session_state.estilo_aprendido.append({"reporte": rep})
                if len(st.session_state.estilo_aprendido) > 10:
                    st.session_state.estilo_aprendido = st.session_state.estilo_aprendido[-10:]
                st.success(f"â {len(st.session_state.estilo_aprendido)} ejemplo(s) en memoria")

    with cd:
        if rep:
            st.download_button(
                "â .docx", data=generar_docx(rep),
                file_name="AURA_Informe.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, key="btn_docx"
            )

    # ââ Sugerencias de estilo basadas en literatura âââââââââ
    if st.button("ð¡ Sugerencias de redacciÃ³n", use_container_width=True,
                 key="btn_sug",
                 help="Sugerencias basadas en literatura y definiciones operativas"):
        if rep and api_key:
            cl = get_client(); mid = MODELS[st.session_state.modelo]["id"]
            dic = st.session_state.diccionario
            dic_ctx = "\n".join([f"Â· {k}: {v}" for k, v in dic.items()]) if dic else "No definido."
