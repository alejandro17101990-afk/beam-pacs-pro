import streamlit as st
import streamlit.components.v1 as components
from openai import OpenAI
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile, io, os, re

st.set_page_config(page_title="AURA", layout="wide", initial_sidebar_state="collapsed")

# ─────────────────────────────────────────────────────────────
# TEMAS
# ─────────────────────────────────────────────────────────────
THEMES = {
    "AURA Dark": {
        "bg":"#0b0f14","panel":"#111720","card":"#161d27","border":"#1e2a38",
        "accent":"#3b9eff","text":"#dce8f4","muted":"#5a7a96","ed_bg":"#0e1520","green":"#22c55e",
    },
    "Aurora": {
        "bg":"#0d0a1a","panel":"#130e24","card":"#1a1230","border":"#2a1f4a",
        "accent":"#a78bfa","text":"#e8e0ff","muted":"#6b5ca5","ed_bg":"#100c20","green":"#34d399",
    },
    "Bosque": {
        "bg":"#0a1209","panel":"#0f1a0d","card":"#142112","border":"#1e3a1a",
        "accent":"#4ade80","text":"#d4e8d0","muted":"#4a7a44","ed_bg":"#0c1810","green":"#86efac",
    },
    "Océano": {
        "bg":"#040f1a","panel":"#061624","card":"#081e30","border":"#0e2d45",
        "accent":"#22d3ee","text":"#cff5fc","muted":"#276a80","ed_bg":"#050f1a","green":"#34d399",
    },
    "Claro": {
        "bg":"#f0f4f8","panel":"#ffffff","card":"#f8fafc","border":"#e2e8f0",
        "accent":"#2563eb","text":"#1e293b","muted":"#64748b","ed_bg":"#f8fafc","green":"#16a34a",
    },
}

MODELS = {
    "DeepSeek Chat": {"url":"https://api.deepseek.com","id":"deepseek-chat"},
    "GPT-4o Mini":   {"url":None,"id":"gpt-4o-mini"},
    "GPT-4.1 Mini":  {"url":None,"id":"gpt-4.1-mini"},
}

MODALIDADES = [
    "Resonancia Magnética","Tomografía Computarizada","Radiografía",
    "Ultrasonido","PET-CT","Mamografía","Fluoroscopía","Angiografía",
]

REGIONES = {
    "Extremidades inferiores": ["Rodilla","Cadera","Tobillo","Pie","Muslo","Pierna"],
    "Extremidades superiores": ["Hombro","Codo","Muñeca","Mano","Brazo","Antebrazo"],
    "Columna":                 ["Col. cervical","Col. dorsal","Col. lumbar","Sacro / Cóccix"],
    "Cráneo y cuello":        ["Cerebro","Cuello","Tiroides","Órbitas","Oídos","Glándulas salivales"],
    "Tórax":                   ["Tórax","Pulmón","Corazón","Mediastino","Mama"],
    "Abdomen y pelvis":        ["Abdomen","Pelvis","Hígado","Páncreas","Riñones","Vejiga","Útero / Anexos"],
}

HCOLS = ["#3b9eff","#22c55e","#f59e0b","#ec4899","#8b5cf6","#06b6d4"]

DEFAULTS = {
    "dictado":"","reporte":"","defs":"",
    "modelo":"DeepSeek Chat","audio_id":None,
    "historial":[],"plantilla_txt":"",
    "panel_izq":True,"panel_der":True,"tema":"AURA Dark",
}
for k,v in DEFAULTS.items():
    if k not in st.session_state: st.session_state[k]=v

try:    api_key = st.secrets["deepseek_key"]
except: api_key = os.environ.get("OPENAI_API_KEY","")

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
def T(): return THEMES[st.session_state.tema]

def get_client():
    cfg = MODELS[st.session_state.modelo]
    return OpenAI(api_key=api_key,base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)

def leer_plantilla(f):
    doc = Document(f); partes = []; n = 0
    try:
        import docx.text.paragraph as pp, docx.table as tt
        for el in doc.element.body:
            tag = el.tag.split('}')[-1]
            if tag == 'p':
                p = pp.Paragraph(el,doc); t = p.text.strip()
                if t: partes.append(t)
            elif tag == 'tbl':
                n += 1; tbl = tt.Table(el,doc)
                rows = ["| "+" | ".join(c.text.strip() for c in r.cells)+" |" for r in tbl.rows]
                partes.append(f"[TABLA {n}]\n"+"\n".join(rows)+"\n[/TABLA]")
    except:
        partes = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n".join(partes), n > 0

def generar_docx(texto):
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    for line in texto.split("\n"):
        s = line.strip()
        if not s: doc.add_paragraph(); continue
        # Limpiar asteriscos Markdown residuales
        s_clean = re.sub(r'\*+', '', s).strip()
        if s_clean.isupper() and len(s_clean) < 80:
            h = doc.add_heading(s_clean, level=1)
            h.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif s_clean.startswith(("•","·")):
            doc.add_paragraph(s_clean[1:].strip(), style="List Bullet")
        else:
            doc.add_paragraph(s_clean)
    bio = io.BytesIO(); doc.save(bio); return bio.getvalue()

def transcribir(audio):
    cfg = MODELS[st.session_state.modelo]
    cl = OpenAI(api_key=api_key,base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)
    with tempfile.NamedTemporaryFile(delete=False,suffix=".wav") as tmp:
        tmp.write(audio.read()); path = tmp.name
    with open(path,"rb") as f:
        res = cl.audio.transcriptions.create(
            model="whisper-1",file=f,language="es",
            prompt="Dictado radiológico: Stoller, ICRS, LCA, menisco, condromalacia, osteofito, Kellgren-Lawrence."
        )
    os.unlink(path); return res.text.strip()

def completitud(texto):
    # Limpiar asteriscos para el conteo
    t = re.sub(r'\*+','',texto)
    secs = sum(1 for s in ["TÉCNICA","HALLAZGOS","IMPRESIÓN"] if s in t.upper())
    words = len(t.split())
    return min(100, int((secs/3)*60 + min(words/150,1)*40)), words

def limpiar_asteriscos(texto):
    """Elimina asteriscos Markdown del texto del modelo."""
    # Reemplaza **texto** → texto (negrita Markdown)
    texto = re.sub(r'\*\*(.+?)\*\*', r'\1', texto)
    # Reemplaza *texto* → texto (cursiva Markdown)
    texto = re.sub(r'\*(.+?)\*', r'\1', texto)
    # Elimina asteriscos sueltos al inicio de línea (usados como viñetas)
    texto = re.sub(r'^\*\s+', '• ', texto, flags=re.MULTILINE)
    # Elimina cualquier asterisco restante
    texto = re.sub(r'\*+', '', texto)
    return texto.strip()

# ─────────────────────────────────────────────────────────────
# CSS DINÁMICO
# ─────────────────────────────────────────────────────────────
def render_css():
    t = T()
    BG=t["bg"]; PANEL=t["panel"]; CARD=t["card"]; BORDER=t["border"]
    ACC=t["accent"]; TXT=t["text"]; MUT=t["muted"]; EBG=t["ed_bg"]; GRN=t["green"]
    st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

html,body,.stApp{{background:{BG};color:{TXT};font-family:'Inter',sans-serif}}
header,footer,#MainMenu{{visibility:hidden}}
.block-container{{padding:0!important;max-width:100%!important}}
*{{box-sizing:border-box}}

/* TOPBAR */
.topbar{{height:50px;background:{PANEL};border-bottom:1px solid {BORDER};
  display:flex;align-items:center;padding:0 24px;gap:14px;
  position:sticky;top:0;z-index:100}}
.logo{{font-size:17px;font-weight:600;color:{ACC};letter-spacing:.14em;
  display:flex;align-items:center;gap:8px}}
.logo-dot{{width:7px;height:7px;border-radius:50%;background:{ACC};
  animation:dp 2s ease-in-out infinite}}
@keyframes dp{{0%,100%{{opacity:1}}50%{{opacity:.2}}}}
.t-sep{{width:1px;height:16px;background:{BORDER}}}
.t-meta{{font-size:12px;color:{MUT}}}
.t-badge{{font-size:10px;color:{ACC};background:{ACC}18;
  border:1px solid {ACC}40;border-radius:5px;padding:2px 8px}}
.t-right{{margin-left:auto;display:flex;align-items:center;gap:8px}}
.t-online{{width:6px;height:6px;border-radius:50%;background:{GRN};box-shadow:0 0 4px {GRN}}}

/* SELECTS */
[data-testid="stSelectbox"]>div>div{{
  background:{CARD}!important;border:1px solid {BORDER}!important;
  border-radius:8px!important;color:{TXT}!important;font-size:13px!important}}
[data-testid="stSelectbox"]>div>div:hover{{border-color:{ACC}60!important}}

/* INPUTS */
.stTextInput input{{background:{CARD}!important;border:1px solid {BORDER}!important;
  border-radius:8px!important;color:{TXT}!important;font-size:13px!important;padding:8px 12px!important}}
.stTextInput input:focus{{border-color:{ACC}50!important;box-shadow:none!important}}
.stTextInput input::placeholder{{color:{MUT}!important}}

/* TEXTAREA */
.stTextArea textarea{{
  background:{EBG}!important;border:1px solid {BORDER}!important;
  border-radius:10px!important;color:{TXT}!important;
  font-size:13.5px!important;line-height:1.78!important;
  padding:18px 20px!important;caret-color:{ACC}!important}}
.stTextArea textarea:focus{{border-color:{ACC}50!important;box-shadow:0 0 0 3px {ACC}12!important}}
.stTextArea textarea::placeholder{{color:{MUT}!important}}

/* AUDIO */
[data-testid="stAudioInput"]{{background:{CARD}!important;border:1px solid {BORDER}!important;border-radius:10px!important}}

/* FILE UPLOADER */
[data-testid="stFileUploader"]{{background:{CARD};border:1px dashed {BORDER};border-radius:10px;padding:6px}}
[data-testid="stFileUploader"] *{{color:{MUT}!important;font-size:12px!important}}

/* BOTONES */
.stButton button{{
  background:{CARD}!important;border:1px solid {BORDER}!important;
  color:{TXT}!important;border-radius:8px!important;
  font-size:13px!important;font-weight:500!important;transition:all .15s!important}}
.stButton button:hover{{border-color:{ACC}60!important;background:{PANEL}!important}}
.btn-primary .stButton button{{
  background:{ACC}!important;border-color:{ACC}!important;
  color:#fff!important;font-weight:600!important}}
.btn-primary .stButton button:hover{{opacity:.88!important}}
.stDownloadButton button{{
  background:transparent!important;border:1px solid {ACC}!important;
  color:{ACC}!important;border-radius:8px!important;font-size:13px!important}}
.stDownloadButton button:hover{{background:{ACC}18!important}}

/* EXPANDER */
[data-testid="stExpander"]{{background:{CARD}!important;border:1px solid {BORDER}!important;
  border-radius:10px!important;margin-bottom:8px!important}}
[data-testid="stExpander"] summary{{color:{MUT}!important;font-size:13px!important;padding:10px 14px!important}}
[data-testid="stExpander"] summary:hover{{color:{TXT}!important}}

/* TABS */
[data-testid="stTabs"] [role="tablist"]{{border-bottom:1px solid {BORDER}!important;background:transparent!important;gap:0!important}}
[data-testid="stTabs"] [role="tab"]{{background:transparent!important;border:none!important;
  color:{MUT}!important;font-size:13px!important;padding:8px 14px!important;
  border-bottom:2px solid transparent!important;border-radius:0!important}}
[data-testid="stTabs"] [role="tab"][aria-selected="true"]{{color:{ACC}!important;border-bottom-color:{ACC}!important}}
[data-testid="stTabs"] [data-baseweb="tab-panel"]{{background:transparent!important;padding:10px 0 0!important}}

/* PROGRESS */
.prog-row{{display:flex;align-items:center;gap:10px;margin-bottom:10px}}
.prog-bg{{flex:1;height:3px;background:{BORDER};border-radius:2px;overflow:hidden}}
.prog-fill{{height:100%;background:{ACC};border-radius:2px;transition:width .4s}}
.prog-txt{{font-size:11px;color:{MUT};white-space:nowrap}}

/* HISTORIAL */
.h-row{{display:flex;align-items:center;gap:8px;padding:7px 10px;border-radius:8px;
  background:{PANEL};border:1px solid {BORDER};margin-bottom:4px}}
.h-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0}}
.h-name{{font-size:12px;color:{TXT}}}
.h-sub{{font-size:11px;color:{MUT}}}

/* DEFS */
.defs-box{{background:{EBG};border:1px solid {BORDER};border-radius:10px;
  padding:14px;font-size:12.5px;line-height:1.58;color:{MUT};white-space:pre-wrap}}

/* SECTION TITLE */
.sec-title{{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;
  color:{MUT};margin-bottom:10px;display:flex;align-items:center;gap:6px}}
.sec-dot{{width:5px;height:5px;border-radius:50%;background:{ACC}}}

/* LABEL */
.lbl{{font-size:10px;color:{MUT};letter-spacing:.1em;text-transform:uppercase;margin-bottom:4px;display:block}}

::-webkit-scrollbar{{width:3px}}
::-webkit-scrollbar-thumb{{background:{BORDER};border-radius:2px}}
hr{{border:none;border-top:1px solid {BORDER}!important;margin:12px 0!important}}
[data-testid="column"]{{padding:0!important}}
</style>
""", unsafe_allow_html=True)

render_css()
t = T()
ACC=t["accent"]; TXT=t["text"]; MUT=t["muted"]; BRD=t["border"]
PNL=t["panel"]; CRD=t["card"]; GRN=t["green"]; EBG=t["ed_bg"]

# ─────────────────────────────────────────────────────────────
# TOPBAR
# ─────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="topbar">
  <div class="logo"><div class="logo-dot"></div>AURA</div>
  <div class="t-sep"></div>
  <span class="t-meta">Radiology Intelligence</span>
  <div class="t-right">
    <span class="t-badge">{st.session_state.modelo}</span>
    <div class="t-online"></div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# RATIOS DE COLUMNAS SEGÚN ESTADO
# ─────────────────────────────────────────────────────────────
lo = st.session_state.panel_izq
ro = st.session_state.panel_der

if   lo and ro:      ratios = [1.05, 2.5, 0.95]
elif lo and not ro:  ratios = [1.05, 3.4, 0.07]
elif not lo and ro:  ratios = [0.07, 3.4, 0.95]
else:                ratios = [0.07, 5.2, 0.07]

col_l, col_c, col_r = st.columns(ratios, gap="small")

# ═══════════════════════════════════════════════════════════════
# PANEL IZQUIERDO
# ═══════════════════════════════════════════════════════════════
with col_l:
    if st.button("◀" if lo else "▶", key="tog_l"):
        st.session_state.panel_izq = not lo; st.rerun()

    if not lo:
        st.markdown(f"""<div style="display:flex;flex-direction:column;
          align-items:center;gap:16px;padding:14px 0">
          <span style="font-size:15px;color:{MUT}" title="Dictado">🎙</span>
          <span style="font-size:15px;color:{MUT}" title="Historial">📋</span>
          <span style="font-size:15px;color:{MUT}" title="Configuración">⚙</span>
        </div>""", unsafe_allow_html=True)
        generar = False

    else:
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

        # ── Estudio ─────────────────────────────────────────
        st.markdown(f'<div class="sec-title"><div class="sec-dot"></div>Estudio</div>',
                    unsafe_allow_html=True)

        st.markdown('<span class="lbl">Modalidad</span>', unsafe_allow_html=True)
        st.selectbox("mod", MODALIDADES, label_visibility="collapsed", key="sel_mod")

        c_grupo, c_reg = st.columns(2)
        with c_grupo:
            st.markdown('<span class="lbl">Grupo</span>', unsafe_allow_html=True)
            grupo = st.selectbox("grp", list(REGIONES.keys()),
                                 label_visibility="collapsed", key="sel_grupo")
        with c_reg:
            st.markdown('<span class="lbl">Región</span>', unsafe_allow_html=True)
            st.selectbox("reg", REGIONES[grupo],
                         label_visibility="collapsed", key="sel_reg")

        st.markdown('<span class="lbl" style="margin-top:6px">Región libre (opcional)</span>',
                    unsafe_allow_html=True)
        st.text_input("rc", label_visibility="collapsed", key="reg_custom",
                      placeholder="Ej: Articulación glenohumeral derecha")

        st.markdown("<hr>", unsafe_allow_html=True)

        # ── Dictado ─────────────────────────────────────────
        st.markdown(f'<div class="sec-title"><div class="sec-dot"></div>Dictado</div>',
                    unsafe_allow_html=True)

        tab_voz, tab_txt = st.tabs(["🎙 Voz", "⌨ Texto"])

        with tab_voz:
            # Botón de micrófono visual
            st.markdown(f"""
            <div style="display:flex;flex-direction:column;align-items:center;
              padding:16px 0 10px;gap:8px">
              <div style="position:relative;width:68px;height:68px">
                <div style="position:absolute;inset:-12px;border-radius:50%;
                  border:1.5px solid {ACC}28;animation:rp 2.4s ease-out infinite"></div>
                <div style="position:absolute;inset:-5px;border-radius:50%;
                  border:1.5px solid {ACC}42;animation:rp 2.4s ease-out infinite .55s"></div>
                <div style="width:68px;height:68px;border-radius:50%;
                  background:{CRD};border:2px solid {ACC};
                  display:flex;align-items:center;justify-content:center">
                  <svg width="24" height="24" viewBox="0 0 24 24" fill="none"
                    stroke="{ACC}" stroke-width="1.8"
                    stroke-linecap="round" stroke-linejoin="round">
                    <rect x="9" y="2" width="6" height="12" rx="3"/>
                    <path d="M5 10a7 7 0 0 0 14 0"/>
                    <line x1="12" y1="19" x2="12" y2="22"/>
                    <line x1="9"  y1="22" x2="15" y2="22"/>
                  </svg>
                </div>
              </div>
              <span style="font-size:11px;color:{MUT}">Pulsa para grabar</span>
            </div>
            <style>
            @keyframes rp{{
              0%  {{transform:scale(1);  opacity:.6}}
              100%{{transform:scale(1.4);opacity:0}}
            }}
            </style>
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
                        st.info("Configura tu API Key.")

        with tab_txt:
            st.markdown(f'<p style="font-size:12px;color:{MUT}">Escribe hallazgos directamente.</p>',
                        unsafe_allow_html=True)

        st.markdown('<span class="lbl" style="margin-top:8px">Señal de entrada</span>',
                    unsafe_allow_html=True)
        dictado = st.text_area(
            "d", value=st.session_state.dictado, height=190,
            label_visibility="collapsed",
            placeholder="El dictado transcrito aparece aquí.\nTambién puedes escribir directamente.\n\nEj: Desgarro horizontal menisco medial Stoller III, extrusión 3 mm, osteofitos marginales.",
            key="dictado_ta"
        )
        if dictado != st.session_state.dictado:
            st.session_state.dictado = dictado

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        ba, bb = st.columns([1.6, 1])
        with ba:
            st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
            generar = st.button("✦  Generar informe", use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with bb:
            if st.button("Limpiar", use_container_width=True):
                st.session_state.dictado = ""
                st.session_state.audio_id = None
                st.rerun()

        st.markdown("<hr>", unsafe_allow_html=True)

        # ── Configuración avanzada ───────────────────────────
        with st.expander("⚙  Configuración", expanded=False):
            st.markdown('<span class="lbl">Modelo IA</span>', unsafe_allow_html=True)
            m = st.selectbox("m", list(MODELS.keys()),
                             index=list(MODELS.keys()).index(st.session_state.modelo),
                             label_visibility="collapsed")
            if m != st.session_state.modelo:
                st.session_state.modelo = m; st.rerun()

            st.markdown('<span class="lbl" style="margin-top:8px">Plantilla .DOCX</span>',
                        unsafe_allow_html=True)
            f_up = st.file_uploader("plt", type=["docx"], label_visibility="collapsed")
            if f_up:
                st.session_state.plantilla_txt, _ = leer_plantilla(f_up)
                st.success("✓ Plantilla cargada")

            if not api_key:
                st.markdown('<span class="lbl" style="margin-top:8px">API Key</span>',
                            unsafe_allow_html=True)
                api_key = st.text_input("k", type="password",
                                        label_visibility="collapsed", placeholder="sk- ···")

            st.markdown('<span class="lbl" style="margin-top:8px">Tema</span>',
                        unsafe_allow_html=True)
            cols_t = st.columns(3)
            for i, nombre in enumerate(THEMES):
                with cols_t[i % 3]:
                    active = "✓ " if nombre == st.session_state.tema else ""
                    if st.button(f"{active}{nombre}", key=f"tm{i}", use_container_width=True):
                        st.session_state.tema = nombre; st.rerun()

        # ── Historial ────────────────────────────────────────
        if st.session_state.historial:
            with st.expander(f"📋  Historial  ·  {len(st.session_state.historial)}", expanded=False):
                for i, e in enumerate(st.session_state.historial):
                    color = HCOLS[i % len(HCOLS)]
                    st.markdown(f"""<div class="h-row">
                      <div class="h-dot" style="background:{color}"></div>
                      <div>
                        <div class="h-name">{e['region']}</div>
                        <div class="h-sub">{e['modalidad'][:18]}</div>
                      </div>
                    </div>""", unsafe_allow_html=True)
                    if st.button("Cargar", key=f"h{i}", use_container_width=True):
                        st.session_state.reporte = e['texto']; st.rerun()

# ─────────────────────────────────────────────────────────────
# PROCESAMIENTO IA
# ─────────────────────────────────────────────────────────────
if generar:
    if not api_key:
        st.warning("Ingresa tu API Key en Configuración.")
    elif not st.session_state.dictado.strip():
        st.warning("Escribe o dicta hallazgos primero.")
    else:
        cl  = get_client()
        mid = MODELS[st.session_state.modelo]["id"]
        pt  = st.session_state.plantilla_txt
        instruc_tabla = (
            "La plantilla tiene tablas [TABLA]. Completa esas tablas en Markdown con los valores del dictado."
            if "[TABLA" in pt else
            "NO generes tablas bajo ninguna circunstancia. No hay plantilla con tabla."
        )
        mod_sel = st.session_state.get("sel_mod", "")
        reg_sel = (st.session_state.get("reg_custom","").strip()
                   or st.session_state.get("sel_reg",""))

        # INSTRUCCIÓN DE PLANTILLA: si hay plantilla, la estructura DEBE respetarse
        if pt:
            instruc_plantilla = f"""RESPETA ESTRICTAMENTE esta plantilla. Usa exactamente sus secciones,
su orden y sus encabezados. Completa el contenido dentro de cada sección:

{pt}"""
        else:
            instruc_plantilla = "ESTRUCTURA:\nINDICACIÓN\nTÉCNICA\nHALLAZGOS\nIMPRESIÓN DIAGNÓSTICA"

        prompt = f"""Eres AURA, sistema experto de interpretación radiológica de nivel subespecialista.
Genera un informe radiológico estructurado y profesional.

MODALIDAD: {mod_sel}
REGIÓN: {reg_sel}

══════════════════════════════════════
FORMATO DE SALIDA — OBLIGATORIO
══════════════════════════════════════
1. CERO asteriscos (*). Absolutamente prohibido. No uses * para nada.
2. CERO markdown. Sin #, sin **, sin *.
3. Títulos de sección: en MAYÚSCULAS puras, solos en su línea.
4. Viñetas en la impresión: usa el carácter • (bullet), no guiones, no asteriscos.
5. Texto corrido en hallazgos: oraciones completas, sin simbolos especiales.

══════════════════════════════════════
REGLAS CLÍNICAS
══════════════════════════════════════
· PROHIBIDO: "cambios degenerativos" sin sustrato morfológico específico.
· Usa descriptores exactos: osteofitos marginales, esclerosis subcondral, pinzamiento de X mm.
· Incluye clasificaciones SOLO cuando los hallazgos del dictado las respalden directamente.
· {instruc_tabla}

{instruc_plantilla}

DICTADO DEL RADIÓLOGO:
{st.session_state.dictado}"""

        with st.spinner("Generando informe..."):
            try:
                res = cl.chat.completions.create(
                    model=mid,
                    messages=[{"role":"system","content":prompt}],
                    temperature=0.1, max_tokens=2500
                )
                report = limpiar_asteriscos(res.choices[0].message.content)
                st.session_state.reporte = report
                st.session_state.historial.insert(0, {
                    "modalidad": mod_sel[:18] if mod_sel else "RM",
                    "region":    reg_sel      if reg_sel  else "General",
                    "texto":     report
                })
                if len(st.session_state.historial) > 12:
                    st.session_state.historial = st.session_state.historial[:12]
                st.rerun()
            except Exception as e:
                st.error(str(e))

# ═══════════════════════════════════════════════════════════════
# PANEL CENTRAL — Editor con formato real (components.html)
# ═══════════════════════════════════════════════════════════════
with col_c:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    rep = st.session_state.reporte

    # Barra de completitud
    if rep:
        pct, words = completitud(rep)
        st.markdown(f"""<div class="prog-row">
          <div class="prog-bg"><div class="prog-fill" style="width:{pct}%"></div></div>
          <span class="prog-txt">{pct}% · {words} palabras</span>
        </div>""", unsafe_allow_html=True)
    else:
        st.markdown(f'<p style="font-size:13px;color:{MUT};margin-bottom:10px">'
                    'El informe generado aparece aquí.</p>', unsafe_allow_html=True)

    # Contenido inicial del editor — escapar caracteres conflictivos
    def prep_html(texto):
        """Convierte texto plano a HTML seguro para el editor."""
        if not texto: return ""
        html = []
        for line in texto.split("\n"):
            s = line.strip()
            if not s:
                html.append("<br>")
            elif s.isupper() and len(s) < 75:
                html.append(f"<b style='display:block;margin-top:12px'>{s}</b>")
            elif s.startswith("•"):
                html.append(f"<li>{s[1:].strip()}</li>")
            else:
                html.append(f"<p style='margin:2px 0'>{s}</p>")
        return "\n".join(html)

    contenido_inicial = prep_html(rep)
    eH = 520
    fH = eH + 90   # toolbar + action strip

    editor_html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{
  height:{fH}px;overflow:hidden;
  display:flex;flex-direction:column;
  background:{t['bg']};font-family:'Inter',sans-serif;
}}

/* ── TOOLBAR ── */
.tb{{
  flex-shrink:0;
  background:{t['panel']};
  border-bottom:1px solid {t['border']};
  padding:5px 10px;
  display:flex;align-items:center;gap:3px;
  flex-wrap:nowrap;overflow-x:auto;
}}
.tb::-webkit-scrollbar{{height:2px}}
.tb::-webkit-scrollbar-thumb{{background:{t['border']}}}

.tg{{display:flex;align-items:center;gap:2px;
    padding-right:6px;border-right:1px solid {t['border']};flex-shrink:0}}
.tg:last-child{{border-right:none}}

.btn{{
  background:none;border:1px solid transparent;
  color:{t['muted']};font-size:11px;
  padding:4px 6px;border-radius:5px;cursor:pointer;
  transition:all .12s;min-width:24px;text-align:center;
  font-family:'Inter',sans-serif;
}}
.btn:hover{{background:{t['card']};color:{t['text']};border-color:{t['border']}}}
.btn.on{{background:{t['accent']}22;color:{t['accent']};border-color:{t['accent']}55}}

.sel{{
  background:{t['card']};border:1px solid {t['border']};
  color:{t['muted']};font-size:11px;font-family:'Inter',sans-serif;
  padding:3px 5px;border-radius:5px;outline:none;cursor:pointer;
}}
.sel:focus{{border-color:{t['accent']}60}}

.dot{{
  width:13px;height:13px;border-radius:50%;cursor:pointer;
  border:2px solid transparent;transition:all .12s;flex-shrink:0;
}}
.dot:hover,.dot.on{{border-color:{t['accent']}}}

.tl{{font-size:9px;color:{t['border']};letter-spacing:.12em;white-space:nowrap}}

/* ── EDITOR AREA ── */
.ew{{
  flex:1;overflow-y:auto;padding:14px 16px;min-height:0;
  background:{t['card']};
}}
.ew::-webkit-scrollbar{{width:3px}}
.ew::-webkit-scrollbar-thumb{{background:{t['border']};border-radius:2px}}

.doc{{
  min-height:100%;padding:20px 26px;outline:none;border-radius:6px;
  font-family:'Inter',sans-serif;font-size:13.5px;line-height:1.78;
  color:{t['ed_bg'] if t['ed_bg'] == t['bg'] else t['text']};
  background:{t['ed_bg']};
  transition:background .2s,color .2s;
}}
.doc b,.doc strong{{color:{t['text']};font-weight:600}}
.doc li{{margin-left:18px;margin-bottom:2px}}
.doc hr{{border:none;border-top:1px solid {t['border']};margin:10px 0}}
.doc p{{margin:1px 0}}
.doc table{{border-collapse:collapse;width:100%;margin:10px 0;font-size:12.5px}}
.doc td,.doc th{{border:1px solid {t['border']};padding:5px 10px;color:{t['text']}}}
.doc th{{background:{t['card']};font-weight:600;color:{t['accent']}}}

/* ── ACTION STRIP ── */
.as{{
  flex-shrink:0;background:{t['panel']};
  border-top:1px solid {t['border']};
  padding:5px 10px;display:flex;align-items:center;gap:5px;
}}
.ab{{
  background:none;border:1px solid {t['border']};
  color:{t['muted']};font-size:10px;
  padding:4px 9px;border-radius:5px;cursor:pointer;
  transition:all .12s;font-family:'Inter',sans-serif;
}}
.ab:hover{{color:{t['text']};border-color:{t['accent']}60}}
.ab.prime{{color:{t['accent']};border-color:{t['accent']}60}}
.pw{{margin-left:auto;display:flex;align-items:center;gap:6px}}
.pl{{width:50px;height:2px;background:{t['border']};border-radius:1px;overflow:hidden}}
.pf{{height:100%;background:{t['accent']};border-radius:1px;transition:width .4s}}
.pp{{font-size:10px;color:{t['muted']}}}
.wc{{font-size:10px;color:{t['border']}}}
</style>
</head>
<body>

<!-- TOOLBAR -->
<div class="tb">
  <div class="tg">
    <select class="sel" id="fnt" onchange="setFont(this.value)" style="width:82px">
      <option value="'Inter',sans-serif" selected>Inter</option>
      <option value="Georgia,serif">Georgia</option>
      <option value="'Courier New',monospace">Courier</option>
      <option value="Arial,sans-serif">Arial</option>
      <option value="'Times New Roman',serif">Times</option>
    </select>
    <select class="sel" id="fsz" onchange="setSize(this.value)" style="width:40px">
      <option value="9">9</option>
      <option value="10">10</option>
      <option value="11">11</option>
      <option value="12">12</option>
      <option value="13" selected>13</option>
      <option value="14">14</option>
      <option value="15">15</option>
      <option value="16">16</option>
      <option value="18">18</option>
    </select>
  </div>
  <div class="tg">
    <button class="btn" id="btnB" onclick="fmt('bold')" title="Negrita"><b>B</b></button>
    <button class="btn" id="btnI" onclick="fmt('italic')" title="Cursiva"><i>I</i></button>
    <button class="btn" id="btnU" onclick="fmt('underline')" title="Subrayado"><u>U</u></button>
  </div>
  <div class="tg">
    <button class="btn" onclick="fmt('justifyLeft')"   title="Izquierda"><i class="ti ti-align-left"></i></button>
    <button class="btn" onclick="fmt('justifyCenter')" title="Centro"><i class="ti ti-align-center"></i></button>
    <button class="btn" onclick="fmt('justifyRight')"  title="Derecha"><i class="ti ti-align-right"></i></button>
    <button class="btn" onclick="fmt('justifyFull')"   title="Justificado"><i class="ti ti-align-justified"></i></button>
  </div>
  <div class="tg">
    <button class="btn" onclick="fmt('insertUnorderedList')" title="Viñetas"><i class="ti ti-list"></i></button>
    <button class="btn" onclick="fmt('insertOrderedList')"   title="Numerada"><i class="ti ti-list-numbers"></i></button>
    <button class="btn" onclick="insHR()" title="Separador">—</button>
    <button class="btn" onclick="insTable()" title="Insertar tabla"><i class="ti ti-table"></i></button>
  </div>
  <div class="tg" style="gap:4px;align-items:center">
    <span class="tl">BG</span>
    <div class="dot on"   style="background:{t['ed_bg']}"      onclick="setBg(this,'{t['ed_bg']}',   '{t['text']}')"   title="Tema"></div>
    <div class="dot"      style="background:#0a1018"             onclick="setBg(this,'#0a1018',       '#c8e8f8')"        title="DICOM"></div>
    <div class="dot"      style="background:#f5f5f0;border:1px solid #ccc" onclick="setBg(this,'#f5f5f0','#1a1a1a')" title="Claro"></div>
    <div class="dot"      style="background:#f5f0e8;border:1px solid #ddd" onclick="setBg(this,'#f5f0e8','#1a0e00')" title="Pergamino"></div>
  </div>
  <div class="tg">
    <button class="btn" onclick="copyText()" title="Copiar texto"><i class="ti ti-copy"></i></button>
    <button class="btn" onclick="printDoc()" title="Imprimir / PDF"><i class="ti ti-printer"></i></button>
  </div>
</div>

<!-- EDITOR -->
<div class="ew">
  <div class="doc" id="doc" contenteditable="true" spellcheck="false">{contenido_inicial}</div>
</div>

<!-- ACTION STRIP -->
<div class="as">
  <span class="wc" id="wc">—</span>
  <div class="pw">
    <div class="pl"><div class="pf" id="pf" style="width:0%"></div></div>
    <span class="pp" id="pp">0%</span>
  </div>
</div>

<script>
var doc = document.getElementById('doc');

// ── Formato ──
function fmt(cmd) {{
  doc.focus();
  document.execCommand(cmd, false, null);
  upd();
}}
function upd() {{
  ['Bold','Italic','Underline'].forEach(function(c) {{
    var b = document.getElementById('btn' + c[0]);
    if (b) b.classList.toggle('on', document.queryCommandState(c.toLowerCase()));
  }});
}}
function setFont(f) {{
  doc.style.fontFamily = f;
}}
function setSize(s) {{
  doc.style.fontSize = s + 'px';
}}
function setBg(el, bg, col) {{
  doc.style.background = bg;
  doc.style.color = col;
  document.querySelectorAll('.dot').forEach(function(d) {{ d.classList.remove('on'); }});
  el.classList.add('on');
}}
function insHR() {{
  doc.focus();
  document.execCommand('insertHTML', false,
    '<hr style="border:none;border-top:1px solid {t['border']};margin:10px 0"><br>');
}}
function insTable() {{
  var r = parseInt(prompt('Filas:', '3')) || 3;
  var c = parseInt(prompt('Columnas:', '3')) || 3;
  var html = '<table><thead><tr>';
  for (var i = 0; i < c; i++) html += '<th>Col ' + (i+1) + '</th>';
  html += '</tr></thead><tbody>';
  for (var j = 0; j < r - 1; j++) {{
    html += '<tr>';
    for (var k = 0; k < c; k++) html += '<td>&nbsp;</td>';
    html += '</tr>';
  }}
  html += '</tbody></table><p><br></p>';
  doc.focus();
  document.execCommand('insertHTML', false, html);
}}

// ── Completitud ──
function calcPct() {{
  var t = doc.innerText.toUpperCase();
  var f = ['TÉCNICA','HALLAZGOS','IMPRESIÓN'].filter(function(s) {{ return t.includes(s); }}).length;
  var w = t.split(/\s+/).filter(Boolean).length;
  return Math.min(100, Math.round((f/3)*60 + Math.min(w/150,1)*40));
}}
function updBar() {{
  var s = calcPct();
  document.getElementById('pf').style.width = s + '%';
  document.getElementById('pp').textContent = s + '%';
  var w = doc.innerText.trim().split(/\s+/).filter(Boolean).length;
  document.getElementById('wc').textContent = w + ' palabras';
}}
doc.addEventListener('input',  updBar);
doc.addEventListener('keyup',  upd);
doc.addEventListener('mouseup',upd);
window.addEventListener('load', function() {{ updBar(); }});

// ── Copiar texto limpio ──
function copyText() {{
  var text = doc.innerText;
  if (navigator.clipboard && navigator.clipboard.writeText) {{
    navigator.clipboard.writeText(text).then(function() {{ toast('Copiado'); }});
  }} else {{
    var ta = document.createElement('textarea');
    ta.value = text; ta.style.cssText = 'position:fixed;opacity:0';
    document.body.appendChild(ta); ta.select();
    document.execCommand('copy'); document.body.removeChild(ta);
    toast('Copiado');
  }}
}}

// ── Imprimir / PDF ──
function printDoc() {{
  var w = window.open('', '_blank');
  w.document.write('<html><head><title>AURA Informe</title>');
  w.document.write('<style>body{{font-family:Calibri,sans-serif;font-size:12pt;line-height:1.7;margin:2cm;color:#111}}');
  w.document.write('b{{font-weight:600}}table{{border-collapse:collapse;width:100%}}');
  w.document.write('td,th{{border:1px solid #ccc;padding:5px 10px}}th{{background:#f0f4f8;font-weight:600}}');
  w.document.write('</style></head><body>');
  w.document.write(doc.innerHTML);
  w.document.write('</body></html>');
  w.document.close();
  setTimeout(function() {{ w.print(); }}, 350);
}}

// ── Toast ──
function toast(m) {{
  var el = document.createElement('div');
  el.textContent = m;
  el.style.cssText = 'position:fixed;bottom:50px;left:50%;transform:translateX(-50%);'
    + 'background:{t['card']};color:{t['accent']};border:1px solid {t['accent']}60;'
    + 'padding:5px 14px;border-radius:5px;font-size:11px;z-index:9999;pointer-events:none';
  document.body.appendChild(el);
  setTimeout(function() {{ document.body.removeChild(el); }}, 1500);
}}

doc.addEventListener('keydown', function(e) {{
  if (e.key === 'Tab') {{
    e.preventDefault();
    document.execCommand('insertHTML', false, '&nbsp;&nbsp;&nbsp;&nbsp;');
  }}
}});
</script>
</body></html>"""

    components.html(editor_html, height=fH, scrolling=False)

    # ── Acciones IA ──
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    a1, a2, a3 = st.columns(3)

    with a1:
        if st.button("✦  Optimizar conclusión", use_container_width=True):
            if rep and api_key:
                cl = get_client()
                mid = MODELS[st.session_state.modelo]["id"]
                with st.spinner("Optimizando..."):
                    try:
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":
                                f"""Eres un radiólogo subespecialista. Mejora ÚNICAMENTE la IMPRESIÓN DIAGNÓSTICA.

REGLAS DE FORMATO — OBLIGATORIAS:
1. CERO asteriscos. Sin *, sin **, sin markdown.
2. Títulos en MAYÚSCULAS, solos en su línea.
3. Viñetas: usa • (bullet), no guiones ni asteriscos.
4. Devuelve el informe COMPLETO. Conserva TÉCNICA y HALLAZGOS exactamente.

CRITERIOS DE CALIDAD:
· Jerarquía: hallazgo principal primero.
· Cada viñeta: estructura + diagnóstico + clasificación/grado + implicación clínica.
· Última viñeta: orientación de manejo sugerente.
· Sin hedge words ("podría","posible") salvo diferencial genuino.

INFORME:
{st.session_state.reporte}"""}],
                            temperature=0.2, max_tokens=2500
                        )
                        st.session_state.reporte = limpiar_asteriscos(r.choices[0].message.content)
                        st.rerun()
                    except Exception as e: st.error(str(e))

    with a2:
        if st.button("◇  Definiciones", use_container_width=True):
            if rep and api_key:
                cl = get_client()
                mid = MODELS[st.session_state.modelo]["id"]
                with st.spinner("Analizando..."):
                    try:
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":
                                f"""Analiza el informe radiológico. Responde en español. Sin asteriscos ni markdown.
Sin líneas en blanco entre ítems de la misma sección. Una línea en blanco entre secciones.

CLASIFICACIONES UTILIZADAS
· Sistema: [nombre · sociedad]
· Grado: [grado] — [significado clínico]
· Evidencia en el informe: [hallazgo que lo justifica]
· Referencia: [Autor, Revista, Año]
· URL: [PubMed o sociedad oficial]

CLASIFICACIONES ADICIONALES RECOMENDADAS
[Si no aplica: "El informe usa los sistemas apropiados."]
· Sistema sugerido: [nombre] — Hallazgo que lo justifica: [descripción]

GLOSARIO
· [Término]: [definición en 1-2 líneas]

CORRELACIÓN CLÍNICA
[3-4 líneas. Dirigido al médico tratante. Lenguaje sugerente, no prescriptivo.]

INFORME:
{st.session_state.reporte}"""}],
                            temperature=0.15, max_tokens=2000
                        )
                        st.session_state.defs = r.choices[0].message.content
                        st.session_state.panel_der = True
                        st.rerun()
                    except Exception as e: st.error(str(e))

    with a3:
        if rep:
            st.download_button(
                "↓  Exportar .docx",
                data=generar_docx(st.session_state.reporte),
                file_name="AURA_Informe.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

# ═══════════════════════════════════════════════════════════════
# PANEL DERECHO — Definiciones
# ═══════════════════════════════════════════════════════════════
with col_r:
    if st.button("▶" if ro else "◀", key="tog_r"):
        st.session_state.panel_der = not ro; st.rerun()

    if not ro:
        st.markdown(f"""<div style="display:flex;flex-direction:column;
          align-items:center;gap:14px;padding:12px 0">
          <span style="font-size:15px;color:{MUT}" title="Definiciones">📖</span>
          <span style="font-size:15px;color:{MUT}" title="Referencias">🔗</span>
        </div>""", unsafe_allow_html=True)
    else:
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        st.markdown(f'<div class="sec-title"><div class="sec-dot"></div>Definiciones y referencias</div>',
                    unsafe_allow_html=True)

        if st.session_state.defs:
            st.markdown(
                f'<div class="defs-box">{st.session_state.defs}</div>',
                unsafe_allow_html=True
            )
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
            if st.button("Cerrar", key="close_defs"):
                st.session_state.defs = ""; st.rerun()
        else:
            st.markdown(f"""<div style="padding:24px 0;text-align:center">
              <p style="font-size:12px;color:{MUT};line-height:1.7">
                Genera un informe y presiona<br>
                <strong style="color:{TXT}">◇ Definiciones</strong><br>
                para ver clasificaciones,<br>definiciones y referencias.
              </p>
            </div>""", unsafe_allow_html=True)
