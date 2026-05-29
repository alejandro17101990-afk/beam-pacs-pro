import streamlit as st
import streamlit.components.v1 as components
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile, io, os, re, json

st.set_page_config(page_title="AURA", layout="wide", initial_sidebar_state="collapsed")

# ─────────────────────────────────────────────────────────────
# TEMAS
# ─────────────────────────────────────────────────────────────
THEMES = {
    "AURA Dark": {
        "bg":"#080c11","panel":"#0d1219","card":"#111820","border":"#1a2636",
        "accent":"#3b9eff","text":"#d8eaf8","muted":"#4a6a88","ed_bg":"#0a0f17","green":"#22c55e",
        "accent2":"#7dd3fc","surface":"#141e2b",
    },
    "Aurora": {
        "bg":"#09071a","panel":"#100d22","card":"#16112e","border":"#261e45",
        "accent":"#a78bfa","text":"#ede8ff","muted":"#5e4e8a","ed_bg":"#0e0b1e","green":"#34d399",
        "accent2":"#c4b5fd","surface":"#1a1538",
    },
    "Obsidian": {
        "bg":"#0a0a0a","panel":"#111111","card":"#161616","border":"#242424",
        "accent":"#e2e8f0","text":"#e2e8f0","muted":"#4a5568","ed_bg":"#0d0d0d","green":"#68d391",
        "accent2":"#f7fafc","surface":"#1a1a1a",
    },
    "Océano": {
        "bg":"#030d18","panel":"#051525","card":"#071d30","border":"#0c2e48",
        "accent":"#06b6d4","text":"#caf0f8","muted":"#1e6a80","ed_bg":"#040f1e","green":"#34d399",
        "accent2":"#67e8f9","surface":"#091e2e",
    },
    "Claro": {
        "bg":"#f8fafc","panel":"#ffffff","card":"#f1f5f9","border":"#e2e8f0",
        "accent":"#2563eb","text":"#0f172a","muted":"#94a3b8","ed_bg":"#ffffff","green":"#16a34a",
        "accent2":"#60a5fa","surface":"#f8fafc",
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
    "Cráneo y cuello":         ["Cerebro","Cuello","Tiroides","Órbitas","Oídos","Silla turca","Glándulas salivales"],
    "Tórax":                   ["Tórax","Pulmón","Corazón","Mediastino","Mama"],
    "Abdomen y pelvis":        ["Abdomen","Pelvis","Hígado","Páncreas","Riñones","Vejiga","Próstata","Útero / Anexos","Suprarrenales","Bazo"],
}

HCOLS = ["#3b9eff","#22c55e","#f59e0b","#ec4899","#8b5cf6","#06b6d4"]

DEFAULTS = {
    "dictado":"","reporte":"","defs":"","mentor_feedback":"",
    "modelo":"DeepSeek Chat","audio_id":None,
    "historial":[],"plantilla_txt":"",
    "panel_izq":True,"panel_der":False,"tema":"AURA Dark",
    "estilo_aprendido":[],   # lista de ejemplos de estilo del usuario
    "feedback_pendiente":False,
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
    return OpenAI(api_key=api_key, base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)

def leer_plantilla(f):
    doc = Document(f); partes = []; n = 0
    try:
        import docx.text.paragraph as pp, docx.table as tt
        for el in doc.element.body:
            tag = el.tag.split('}')[-1]
            if tag == 'p':
                p = pp.Paragraph(el,doc); tx = p.text.strip()
                if tx: partes.append(tx)
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
        s = re.sub(r'\*+','',line).strip()
        if not s: doc.add_paragraph(); continue
        if s.isupper() and len(s) < 80:
            h = doc.add_heading(s, level=1); h.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif s.startswith("•"):
            doc.add_paragraph(s[1:].strip(), style="List Bullet")
        else:
            doc.add_paragraph(s)
    bio = io.BytesIO(); doc.save(bio); return bio.getvalue()

def transcribir(audio):
    cfg = MODELS[st.session_state.modelo]
    cl = OpenAI(api_key=api_key, base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp:
        tmp.write(audio.read()); path = tmp.name
    with open(path,"rb") as f:
        res = cl.audio.transcriptions.create(
            model="whisper-1", file=f, language="es",
            prompt="Dictado radiológico: Stoller, ICRS, LCA, menisco, condromalacia, Pfirrmann, Modic, NASCET, BI-RADS, PI-RADS, TIRADS."
        )
    os.unlink(path); return res.text.strip()

def limpiar(texto):
    texto = re.sub(r'\*\*(.+?)\*\*', r'\1', texto)
    texto = re.sub(r'\*(.+?)\*',     r'\1', texto)
    texto = re.sub(r'^\*\s+', '• ',  texto, flags=re.MULTILINE)
    texto = re.sub(r'\*+', '',        texto)
    texto = re.sub(r'^#+\s+', '',     texto, flags=re.MULTILINE)
    return texto.strip()

def completitud(texto):
    t = re.sub(r'\*+','',texto)
    secs = sum(1 for s in ["TÉCNICA","HALLAZGOS","IMPRESIÓN"] if s in t.upper())
    words = len(t.split())
    return min(100, int((secs/3)*60 + min(words/150,1)*40)), words

def build_estilo_context():
    """Construye el bloque de contexto de estilo aprendido."""
    ejemplos = st.session_state.estilo_aprendido
    if not ejemplos:
        return ""
    bloques = []
    for i, e in enumerate(ejemplos[-5:], 1):   # últimos 5
        bloques.append(f"--- Ejemplo {i} (aprobado por el radiólogo) ---\n{e['reporte']}")
    return "\n\n".join(bloques)

def guardar_estilo(reporte_corregido):
    """Guarda un reporte corregido como ejemplo de estilo."""
    st.session_state.estilo_aprendido.append({"reporte": reporte_corregido})
    if len(st.session_state.estilo_aprendido) > 10:
        st.session_state.estilo_aprendido = st.session_state.estilo_aprendido[-10:]

# ─────────────────────────────────────────────────────────────
# CSS GLOBAL
# ─────────────────────────────────────────────────────────────
def render_css():
    t = T()
    st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Serif+Display:ital@0;1&display=swap');

html,body,.stApp{{
  background:{t['bg']};color:{t['text']};
  font-family:'DM Sans',sans-serif;
}}
header,footer,#MainMenu{{visibility:hidden}}
.block-container{{padding:0!important;max-width:100%!important}}
*{{box-sizing:border-box}}

/* ── TOPBAR ── */
.topbar{{
  height:48px;background:{t['panel']};
  border-bottom:1px solid {t['border']};
  display:flex;align-items:center;padding:0 20px;gap:12px;
  position:sticky;top:0;z-index:200;
  backdrop-filter:blur(12px);
}}
.logo{{
  font-family:'DM Serif Display',serif;
  font-size:18px;color:{t['accent']};letter-spacing:.18em;
  display:flex;align-items:center;gap:8px;
}}
.logo-pulse{{
  width:6px;height:6px;border-radius:50%;
  background:{t['accent']};
  box-shadow:0 0 0 0 {t['accent']}66;
  animation:ping 2.2s ease infinite;
}}
@keyframes ping{{
  0%  {{box-shadow:0 0 0 0 {t['accent']}66}}
  70% {{box-shadow:0 0 0 8px {t['accent']}00}}
  100%{{box-shadow:0 0 0 0 {t['accent']}00}}
}}
.t-sep{{width:1px;height:14px;background:{t['border']}}}
.t-label{{font-size:11px;color:{t['muted']};letter-spacing:.06em}}
.t-right{{margin-left:auto;display:flex;align-items:center;gap:10px}}
.t-chip{{
  font-size:10px;color:{t['accent']};
  background:{t['accent']}14;border:1px solid {t['accent']}30;
  border-radius:4px;padding:2px 8px;letter-spacing:.05em;
}}
.t-dot{{
  width:6px;height:6px;border-radius:50%;
  background:{t['green']};box-shadow:0 0 6px {t['green']}88;
}}

/* ── FORM CONTROLS ── */
[data-testid="stSelectbox"]>div>div{{
  background:{t['card']}!important;border:1px solid {t['border']}!important;
  border-radius:7px!important;color:{t['text']}!important;
  font-size:12px!important;font-family:'DM Sans',sans-serif!important;
}}
[data-testid="stSelectbox"]>div>div:hover{{border-color:{t['accent']}55!important}}
.stTextInput input{{
  background:{t['card']}!important;border:1px solid {t['border']}!important;
  border-radius:7px!important;color:{t['text']}!important;
  font-size:12px!important;padding:7px 11px!important;
  font-family:'DM Sans',sans-serif!important;
}}
.stTextInput input:focus{{border-color:{t['accent']}50!important;box-shadow:none!important}}
.stTextInput input::placeholder{{color:{t['muted']}!important}}
.stTextArea textarea{{
  background:{t['ed_bg']}!important;border:1px solid {t['border']}!important;
  border-radius:9px!important;color:{t['text']}!important;
  font-size:12.5px!important;line-height:1.75!important;padding:14px 16px!important;
  font-family:'DM Sans',sans-serif!important;caret-color:{t['accent']}!important;
}}
.stTextArea textarea:focus{{border-color:{t['accent']}45!important;box-shadow:none!important}}
.stTextArea textarea::placeholder{{color:{t['muted']}!important}}
[data-testid="stAudioInput"]{{
  background:{t['card']}!important;border:1px solid {t['border']}!important;
  border-radius:9px!important;
}}
[data-testid="stFileUploader"]{{
  background:{t['card']};border:1px dashed {t['border']};
  border-radius:9px;padding:6px;
}}
[data-testid="stFileUploader"] *{{color:{t['muted']}!important;font-size:11px!important}}

/* ── BUTTONS ── */
.stButton button{{
  background:{t['card']}!important;border:1px solid {t['border']}!important;
  color:{t['text']}!important;border-radius:7px!important;
  font-size:12px!important;font-weight:500!important;
  font-family:'DM Sans',sans-serif!important;
  transition:all .15s!important;letter-spacing:.01em!important;
}}
.stButton button:hover{{
  border-color:{t['accent']}55!important;background:{t['surface']}!important;
  color:{t['accent2']}!important;
}}
.btn-primary .stButton button{{
  background:{t['accent']}!important;border-color:{t['accent']}!important;
  color:#fff!important;font-weight:600!important;
}}
.btn-primary .stButton button:hover{{opacity:.85!important}}
.stDownloadButton button{{
  background:transparent!important;border:1px solid {t['accent']}55!important;
  color:{t['accent']}!important;border-radius:7px!important;font-size:12px!important;
}}
.stDownloadButton button:hover{{background:{t['accent']}10!important}}

/* ── EXPANDER ── */
[data-testid="stExpander"]{{
  background:{t['card']}!important;border:1px solid {t['border']}!important;
  border-radius:9px!important;margin-bottom:6px!important;
}}
[data-testid="stExpander"] summary{{
  color:{t['muted']}!important;font-size:12px!important;padding:9px 13px!important;
  font-family:'DM Sans',sans-serif!important;
}}
[data-testid="stExpander"] summary:hover{{color:{t['text']}!important}}

/* ── TABS ── */
[data-testid="stTabs"] [role="tablist"]{{
  border-bottom:1px solid {t['border']}!important;background:transparent!important;
}}
[data-testid="stTabs"] [role="tab"]{{
  background:transparent!important;border:none!important;
  color:{t['muted']}!important;font-size:12px!important;
  padding:7px 12px!important;border-bottom:2px solid transparent!important;
  border-radius:0!important;font-family:'DM Sans',sans-serif!important;
}}
[data-testid="stTabs"] [role="tab"][aria-selected="true"]{{
  color:{t['accent']}!important;border-bottom-color:{t['accent']}!important;
}}
[data-testid="stTabs"] [data-baseweb="tab-panel"]{{
  background:transparent!important;padding:10px 0 0!important;
}}

/* ── LABELS ── */
.lbl{{
  font-size:9px;font-weight:600;letter-spacing:.14em;
  text-transform:uppercase;color:{t['muted']};
  margin-bottom:4px;display:block;
}}
.sec-hdr{{
  font-size:9px;font-weight:600;letter-spacing:.14em;
  text-transform:uppercase;color:{t['muted']};
  margin-bottom:10px;display:flex;align-items:center;gap:5px;
}}
.sec-dot{{width:4px;height:4px;border-radius:50%;background:{t['accent']};flex-shrink:0}}

/* ── PROGRESS ── */
.prog-wrap{{display:flex;align-items:center;gap:8px;margin-bottom:8px}}
.prog-track{{flex:1;height:2px;background:{t['border']};border-radius:1px;overflow:hidden}}
.prog-bar{{height:100%;background:{t['accent']};border-radius:1px;transition:width .5s ease}}
.prog-label{{font-size:10px;color:{t['muted']};white-space:nowrap}}

/* ── MENTOR CARD ── */
.mentor-card{{
  background:{t['surface']};border:1px solid {t['accent']}30;
  border-left:3px solid {t['accent']};
  border-radius:0 9px 9px 0;
  padding:14px 16px;margin-top:10px;
  font-size:12px;line-height:1.7;color:{t['text']};
}}
.mentor-header{{
  font-size:9px;font-weight:600;letter-spacing:.14em;
  text-transform:uppercase;color:{t['accent']};
  margin-bottom:8px;display:flex;align-items:center;gap:5px;
}}

/* ── HISTORIAL ── */
.h-item{{
  display:flex;align-items:center;gap:8px;
  padding:7px 9px;background:{t['surface']};
  border:1px solid {t['border']};border-radius:7px;
  margin-bottom:4px;cursor:pointer;
  transition:border-color .15s;
}}
.h-item:hover{{border-color:{t['accent']}45}}
.h-dot{{width:7px;height:7px;border-radius:50%;flex-shrink:0}}
.h-name{{font-size:11px;color:{t['text']}}}
.h-sub{{font-size:10px;color:{t['muted']}}}

/* ── DEFS BOX ── */
.defs-box{{
  background:{t['ed_bg']};border:1px solid {t['border']};
  border-radius:9px;padding:14px;
  font-size:11.5px;line-height:1.65;color:{t['muted']};
  white-space:pre-wrap;
}}

::-webkit-scrollbar{{width:3px}}
::-webkit-scrollbar-thumb{{background:{t['border']};border-radius:2px}}
hr{{border:none;border-top:1px solid {t['border']}!important;margin:10px 0!important}}
[data-testid="column"]{{padding:0!important}}
</style>
""", unsafe_allow_html=True)

render_css()
t = T()

# ─────────────────────────────────────────────────────────────
# TOPBAR
# ─────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="topbar">
  <div class="logo"><div class="logo-pulse"></div>AURA</div>
  <div class="t-sep"></div>
  <span class="t-label">Radiology Intelligence</span>
  <div class="t-right">
    <span class="t-chip">{st.session_state.modelo}</span>
    <div class="t-dot" title="Conectado"></div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# LAYOUT
# ─────────────────────────────────────────────────────────────
lo = st.session_state.panel_izq
ro = st.session_state.panel_der

if   lo and ro:      ratios=[1.0, 2.6, 0.85]
elif lo and not ro:  ratios=[1.0, 3.6, 0.06]
elif not lo and ro:  ratios=[0.06, 3.6, 0.85]
else:                ratios=[0.06, 5.4, 0.06]

col_l, col_c, col_r = st.columns(ratios, gap="small")

# ═══════════════════════════════════════════════════════════════
# PANEL IZQUIERDO
# ═══════════════════════════════════════════════════════════════
with col_l:
    if st.button("◀" if lo else "▶", key="tog_l", help="Colapsar panel"):
        st.session_state.panel_izq = not lo; st.rerun()

    if not lo:
        st.markdown(f"""
        <div style="display:flex;flex-direction:column;align-items:center;gap:14px;padding:12px 0">
          <span style="font-size:14px;color:{t['muted']}">🎙</span>
          <span style="font-size:14px;color:{t['muted']}">📋</span>
          <span style="font-size:14px;color:{t['muted']}">⚙</span>
        </div>""", unsafe_allow_html=True)
        generar = False
    else:
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

        # ── Estudio ─────────────────────────────────────────
        st.markdown(f'<div class="sec-hdr"><div class="sec-dot"></div>Estudio</div>', unsafe_allow_html=True)

        st.markdown('<span class="lbl">Modalidad</span>', unsafe_allow_html=True)
        st.selectbox("mod", MODALIDADES, label_visibility="collapsed", key="sel_mod")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<span class="lbl">Grupo</span>', unsafe_allow_html=True)
            grupo = st.selectbox("grp", list(REGIONES.keys()), label_visibility="collapsed", key="sel_grupo")
        with c2:
            st.markdown('<span class="lbl">Región</span>', unsafe_allow_html=True)
            st.selectbox("reg", REGIONES[grupo], label_visibility="collapsed", key="sel_reg")

        st.markdown('<span class="lbl" style="margin-top:5px">Región libre</span>', unsafe_allow_html=True)
        st.text_input("rc", label_visibility="collapsed", key="reg_custom",
                      placeholder="Ej: Articulación glenohumeral derecha")

        st.markdown("<hr>", unsafe_allow_html=True)

        # ── Dictado ─────────────────────────────────────────
        st.markdown(f'<div class="sec-hdr"><div class="sec-dot"></div>Dictado</div>', unsafe_allow_html=True)

        tab_voz, tab_txt = st.tabs(["🎙 Voz", "⌨ Texto"])
        with tab_voz:
            st.markdown(f"""
            <div style="display:flex;flex-direction:column;align-items:center;padding:14px 0 8px;gap:7px">
              <div style="position:relative;width:60px;height:60px">
                <div style="position:absolute;inset:-11px;border-radius:50%;
                  border:1px solid {t['accent']}25;animation:rp 2.5s ease-out infinite"></div>
                <div style="position:absolute;inset:-4px;border-radius:50%;
                  border:1px solid {t['accent']}40;animation:rp 2.5s ease-out infinite .6s"></div>
                <div style="width:60px;height:60px;border-radius:50%;
                  background:{t['card']};border:1.5px solid {t['accent']};
                  display:flex;align-items:center;justify-content:center">
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="none"
                    stroke="{t['accent']}" stroke-width="1.8" stroke-linecap="round">
                    <rect x="9" y="2" width="6" height="12" rx="3"/>
                    <path d="M5 10a7 7 0 0 0 14 0"/>
                    <line x1="12" y1="19" x2="12" y2="22"/>
                    <line x1="9" y1="22" x2="15" y2="22"/>
                  </svg>
                </div>
              </div>
              <span style="font-size:10px;color:{t['muted']}">Pulsa para grabar</span>
            </div>
            <style>@keyframes rp{{0%{{transform:scale(1);opacity:.5}}100%{{transform:scale(1.45);opacity:0}}}}</style>
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
            st.markdown(f'<p style="font-size:11px;color:{t["muted"]};margin-bottom:4px">Escribe los hallazgos directamente.</p>', unsafe_allow_html=True)

        st.markdown('<span class="lbl" style="margin-top:6px">Señal de entrada</span>', unsafe_allow_html=True)
        dictado = st.text_area(
            "d", value=st.session_state.dictado, height=180,
            label_visibility="collapsed",
            placeholder="El dictado transcrito aparece aquí.\nTambién puedes escribir directamente.\n\nEj: Desgarro horizontal menisco medial Stoller III, extrusión 3 mm, sin líquido articular significativo.",
            key="dictado_ta"
        )
        if dictado != st.session_state.dictado:
            st.session_state.dictado = dictado

        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
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

        # ── Aprendizaje de estilo ────────────────────────────
        with st.expander("🧠  Mi estilo aprendido", expanded=False):
            n_ej = len(st.session_state.estilo_aprendido)
            st.markdown(f'<p style="font-size:11px;color:{t["muted"]};margin-bottom:8px">'
                        f'{n_ej} ejemplo{"s" if n_ej!=1 else ""} guardado{"s" if n_ej!=1 else ""}. '
                        'Aprueba informes desde el editor central para que AURA aprenda tu estilo.</p>',
                        unsafe_allow_html=True)
            if n_ej > 0:
                if st.button("🗑  Borrar ejemplos de estilo", use_container_width=True):
                    st.session_state.estilo_aprendido = []; st.rerun()

        # ── Configuración ───────────────────────────────────
        with st.expander("⚙  Configuración", expanded=False):
            st.markdown('<span class="lbl">Modelo IA</span>', unsafe_allow_html=True)
            m = st.selectbox("m", list(MODELS.keys()),
                             index=list(MODELS.keys()).index(st.session_state.modelo),
                             label_visibility="collapsed")
            if m != st.session_state.modelo:
                st.session_state.modelo = m; st.rerun()

            st.markdown('<span class="lbl" style="margin-top:7px">Plantilla .DOCX</span>', unsafe_allow_html=True)
            f_up = st.file_uploader("plt", type=["docx"], label_visibility="collapsed")
            if f_up:
                st.session_state.plantilla_txt, _ = leer_plantilla(f_up)
                st.success("✓ Plantilla cargada")

            if not api_key:
                st.markdown('<span class="lbl" style="margin-top:7px">API Key</span>', unsafe_allow_html=True)
                api_key = st.text_input("k", type="password",
                                        label_visibility="collapsed", placeholder="sk- ···")

            st.markdown('<span class="lbl" style="margin-top:7px">Tema visual</span>', unsafe_allow_html=True)
            cols_t = st.columns(3)
            for i, nombre in enumerate(THEMES):
                with cols_t[i % 3]:
                    active = "✓ " if nombre == st.session_state.tema else ""
                    if st.button(f"{active}{nombre}", key=f"tm{i}", use_container_width=True):
                        st.session_state.tema = nombre; st.rerun()

        # ── Historial ────────────────────────────────────────
        if st.session_state.historial:
            with st.expander(f"📋  Historial · {len(st.session_state.historial)}", expanded=False):
                for i, e in enumerate(st.session_state.historial):
                    color = HCOLS[i % len(HCOLS)]
                    st.markdown(f"""<div class="h-item">
                      <div class="h-dot" style="background:{color}"></div>
                      <div>
                        <div class="h-name">{e['region']}</div>
                        <div class="h-sub">{e['modalidad'][:20]}</div>
                      </div>
                    </div>""", unsafe_allow_html=True)
                    if st.button("Cargar", key=f"h{i}", use_container_width=True):
                        st.session_state.reporte = e['texto']; st.rerun()

# ─────────────────────────────────────────────────────────────
# PROCESAMIENTO — GENERACIÓN IA
# ─────────────────────────────────────────────────────────────
if generar:
    if not api_key:
        st.warning("Ingresa tu API Key en Configuración.")
    elif not st.session_state.dictado.strip():
        st.warning("Escribe o dicta los hallazgos primero.")
    else:
        cl  = get_client()
        mid = MODELS[st.session_state.modelo]["id"]
        pt  = st.session_state.plantilla_txt

        mod_sel = st.session_state.get("sel_mod", "")
        reg_sel = (st.session_state.get("reg_custom","").strip()
                   or st.session_state.get("sel_reg",""))

        instruc_tabla = (
            "La plantilla incluye tablas [TABLA]. Complétalas con los valores del dictado."
            if "[TABLA" in pt else
            "No uses tablas a menos que sean estrictamente necesarias para comparar datos."
        )

        instruc_plantilla = (
            f"RESPETA ESTA PLANTILLA en estructura, secciones y encabezados:\n\n{pt}"
            if pt else
            "Estructura: INDICACIÓN / TÉCNICA / HALLAZGOS / IMPRESIÓN DIAGNÓSTICA"
        )

        estilo_ctx = build_estilo_context()
        instruc_estilo = ""
        if estilo_ctx:
            instruc_estilo = f"""
══════════════════════════════════════
ESTILO PERSONAL DEL RADIÓLOGO (APRENDIDO)
══════════════════════════════════════
Los siguientes son informes aprobados previamente por este radiólogo.
Replica fielmente su estilo de redacción, nivel de detalle, estructura narrativa y terminología:

{estilo_ctx}

Fin de ejemplos de estilo. Aplica este patrón al nuevo informe.
"""

        prompt = f"""Eres AURA, sistema experto de interpretación radiológica de nivel subespecialista.
Actúas simultáneamente como: radiólogo subespecialista senior, editor académico médico,
y copiloto de redacción de alto nivel.

MODALIDAD: {mod_sel}
REGIÓN: {reg_sel}

══════════════════════════════════════
FORMATO DE SALIDA — ABSOLUTAMENTE OBLIGATORIO
══════════════════════════════════════
1. CERO asteriscos (*). Nunca. Absolutamente prohibido.
2. CERO markdown (sin #, sin **, sin *).
3. Los títulos de sección van en MAYÚSCULAS, solos en su línea, sin dos puntos al final.
4. En la IMPRESIÓN DIAGNÓSTICA usa • para cada viñeta. En el resto: PROSA CORRIDA.
5. JAMÁS uses listas de viñetas o guiones en TÉCNICA o HALLAZGOS.
   Los hallazgos se redactan en párrafos narrativos, con oraciones completas y fluidas.
   Ejemplo correcto: "El menisco medial presenta desgarro horizontal en su cuerno posterior,
   Stoller grado III, con extrusión de 3 mm. El ligamento cruzado anterior se observa íntegro,
   con señal y morfología conservadas. El cartílago articular..."
   Ejemplo PROHIBIDO: "- Menisco: desgarro\n- LCA: íntegro\n- Cartílago: normal"
6. Coherencia narrativa: conecta los hallazgos entre sí, establece relaciones anatómicas
   y fisiopatológicas cuando corresponda.

══════════════════════════════════════
ESTÁNDARES CLÍNICOS
══════════════════════════════════════
· PROHIBIDO: "cambios degenerativos" sin describir morfología exacta.
· Cuantifica siempre: dimensiones, porcentajes, grados, scores.
· Clasificaciones aplicables según hallazgo:
  Menisco→Stoller, Cartílago→ICRS/Outerbridge, Columna→Pfirrmann/Modic/Meyerding,
  Hombro→Bigliani/Goutallier/Sugaya, Cadera→Tönnis/alpha-angle,
  Cerebro→ASPECTS/Fazekas, Mama→BI-RADS, Próstata→PI-RADS, Tiroides→TIRADS.
· Solo usa clasificaciones respaldadas directamente por los hallazgos del dictado.
· Nivel de redacción: publicable en revista indexada, auditable por comité de pares.
{instruc_estilo}

{instruc_plantilla}

DICTADO:
{st.session_state.dictado}

{instruc_tabla}"""

        with st.spinner("Generando informe..."):
            try:
                res = cl.chat.completions.create(
                    model=mid,
                    messages=[{"role":"system","content":prompt}],
                    temperature=0.12, max_tokens=3000
                )
                report = limpiar(res.choices[0].message.content)
                st.session_state.reporte = report
                st.session_state.mentor_feedback = ""
                st.session_state.historial.insert(0, {
                    "modalidad": mod_sel[:18] if mod_sel else "RM",
                    "region":    reg_sel      if reg_sel  else "General",
                    "texto":     report,
                })
                if len(st.session_state.historial) > 12:
                    st.session_state.historial = st.session_state.historial[-12:]
                st.rerun()
            except Exception as e:
                st.error(str(e))

# ═══════════════════════════════════════════════════════════════
# PANEL CENTRAL — Editor
# ═══════════════════════════════════════════════════════════════
with col_c:
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    rep = st.session_state.reporte

    # Progress bar
    if rep:
        pct, words = completitud(rep)
        st.markdown(f"""
        <div class="prog-wrap">
          <div class="prog-track"><div class="prog-bar" style="width:{pct}%"></div></div>
          <span class="prog-label">{pct}% · {words} palabras</span>
        </div>""", unsafe_allow_html=True)

    # ── Editor rico via components.html ──────────────────────
    def text_to_html(texto):
        if not texto: return ""
        html = []
        for line in texto.split("\n"):
            s = line.strip()
            if not s:
                html.append("<p style='margin:2px 0'><br></p>")
            elif s.isupper() and 2 < len(s) < 75:
                html.append(f"<h2>{s}</h2>")
            elif s.startswith("•"):
                html.append(f"<li>{s[1:].strip()}</li>")
            else:
                html.append(f"<p>{s}</p>")
        return "\n".join(html)

    contenido = text_to_html(rep)
    h_editor = 580

    editor_html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Serif+Display:ital@0;1&display=swap" rel="stylesheet">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{
  width:100%;height:{h_editor+80}px;
  display:flex;flex-direction:column;
  background:{t['bg']};font-family:'DM Sans',sans-serif;
  overflow:hidden;
}}

/* ── TOOLBAR ── */
.tb{{
  flex-shrink:0;
  background:{t['panel']};
  border-bottom:1px solid {t['border']};
  padding:4px 10px;
  display:flex;align-items:center;gap:2px;flex-wrap:nowrap;
  overflow-x:auto;height:38px;
}}
.tb::-webkit-scrollbar{{height:2px}}
.tb::-webkit-scrollbar-thumb{{background:{t['border']}}}
.tg{{display:flex;align-items:center;gap:1px;
    padding-right:6px;margin-right:2px;
    border-right:1px solid {t['border']};flex-shrink:0}}
.tg:last-child{{border-right:none}}
.btn{{
  background:none;border:1px solid transparent;
  color:{t['muted']};font-size:11px;
  padding:3px 6px;border-radius:5px;cursor:pointer;
  transition:all .1s;min-width:22px;text-align:center;
  font-family:'DM Sans',sans-serif;line-height:1.3;
}}
.btn:hover{{background:{t['card']};color:{t['text']};border-color:{t['border']}}}
.btn.on{{background:{t['accent']}1a;color:{t['accent']};border-color:{t['accent']}44}}
.sel{{
  background:{t['card']};border:1px solid {t['border']};
  color:{t['muted']};font-size:11px;font-family:'DM Sans',sans-serif;
  padding:3px 5px;border-radius:5px;outline:none;cursor:pointer;height:24px;
}}
.sel:focus{{border-color:{t['accent']}55}}

/* ── SCROLL WRAPPER ── */
.scroll-wrap{{
  flex:1;
  overflow-y:auto;
  overflow-x:hidden;
  padding:0;
  min-height:0;
  background:{t['card']};
}}
.scroll-wrap::-webkit-scrollbar{{width:4px}}
.scroll-wrap::-webkit-scrollbar-thumb{{
  background:{t['border']};border-radius:2px;
}}
.scroll-wrap::-webkit-scrollbar-thumb:hover{{
  background:{t['accent']}60;
}}

/* ── PAPER ── */
.paper{{
  min-height:100%;
  margin:0 auto;
  max-width:760px;
  padding:32px 40px 48px;
  background:{t['ed_bg']};
  outline:none;
  font-family:'DM Sans',sans-serif;
  font-size:13.5px;
  line-height:1.82;
  color:{t['text']};
  word-break:break-word;
}}
.paper:focus{{outline:none}}
.paper h2{{
  font-family:'DM Serif Display',serif;
  font-size:14px;font-weight:400;letter-spacing:.08em;
  text-transform:uppercase;color:{t['accent']};
  margin:20px 0 6px;padding-bottom:5px;
  border-bottom:1px solid {t['border']};
}}
.paper h3{{font-size:13px;font-weight:600;margin:12px 0 4px;color:{t['text']}}}
.paper p{{margin:2px 0;}}
.paper li{{margin-left:20px;margin-bottom:3px}}
.paper ul,
.paper ol{{margin:4px 0 4px 20px}}
.paper hr{{border:none;border-top:1px solid {t['border']};margin:12px 0}}
.paper table{{border-collapse:collapse;width:100%;margin:10px 0;font-size:12.5px}}
.paper td,.paper th{{border:1px solid {t['border']};padding:6px 10px;color:{t['text']}}}
.paper th{{background:{t['surface']};font-weight:600;color:{t['accent']};font-size:11px;letter-spacing:.06em}}
.paper tr:nth-child(even) td{{background:{t['surface']}60}}

/* ── BOTTOM BAR ── */
.bbar{{
  flex-shrink:0;height:30px;
  background:{t['panel']};border-top:1px solid {t['border']};
  display:flex;align-items:center;padding:0 12px;gap:8px;
}}
.bstat{{font-size:10px;color:{t['muted']}}}
.btrack{{flex:1;height:2px;background:{t['border']};border-radius:1px;overflow:hidden}}
.bfill{{height:100%;background:{t['accent']};border-radius:1px;transition:width .4s}}
.bpct{{font-size:10px;color:{t['muted']}}}
</style>
</head>
<body>

<!-- TOOLBAR -->
<div class="tb">
  <div class="tg">
    <select class="sel" id="fnt" onchange="setFont(this.value)" style="width:80px">
      <option value="'DM Sans',sans-serif" selected>DM Sans</option>
      <option value="'DM Serif Display',serif">DM Serif</option>
      <option value="Georgia,serif">Georgia</option>
      <option value="'Courier New',monospace">Courier</option>
      <option value="Arial,sans-serif">Arial</option>
      <option value="Calibri,sans-serif">Calibri</option>
    </select>
    <select class="sel" id="fsz" onchange="setFontSize(this.value)" style="width:38px">
      <option>10</option><option>11</option><option>12</option>
      <option selected>13</option><option>14</option><option>15</option>
      <option>16</option><option>18</option><option>20</option>
    </select>
  </div>
  <div class="tg">
    <button class="btn" id="btnB" onclick="fmt('bold')" title="Negrita (Ctrl+B)"><b>B</b></button>
    <button class="btn" id="btnI" onclick="fmt('italic')" title="Cursiva (Ctrl+I)"><i>I</i></button>
    <button class="btn" id="btnU" onclick="fmt('underline')" title="Subrayado (Ctrl+U)"><u>U</u></button>
    <button class="btn" onclick="fmt('strikeThrough')" title="Tachado" style="text-decoration:line-through">S</button>
  </div>
  <div class="tg">
    <button class="btn" onclick="fmt('justifyLeft')"   title="Izquierda">&#8676;</button>
    <button class="btn" onclick="fmt('justifyCenter')" title="Centro">&#8596;</button>
    <button class="btn" onclick="fmt('justifyRight')"  title="Derecha">&#8677;</button>
    <button class="btn" onclick="fmt('justifyFull')"   title="Justificado">&#9776;</button>
  </div>
  <div class="tg">
    <button class="btn" onclick="fmt('insertUnorderedList')" title="Viñetas">&#8226;&#8212;</button>
    <button class="btn" onclick="fmt('insertOrderedList')"   title="Numerada">1.</button>
    <button class="btn" onclick="insHR()" title="Separador">&#8212;</button>
    <button class="btn" onclick="insTable()" title="Tabla" style="font-size:10px">Tabla</button>
  </div>
  <div class="tg" style="gap:4px">
    <label style="font-size:9px;color:{t['muted']};display:flex;align-items:center;gap:2px;cursor:pointer">
      A<input type="color" id="fc" value="{t['text']}" onchange="fmt('foreColor',this.value)"
        style="width:18px;height:18px;padding:0;border:none;border-radius:3px;cursor:pointer;background:none">
    </label>
    <label style="font-size:9px;color:{t['muted']};display:flex;align-items:center;gap:2px;cursor:pointer">
      HL<input type="color" id="hc" value="{t['accent']}" onchange="fmt('hiliteColor',this.value)"
        style="width:18px;height:18px;padding:0;border:none;border-radius:3px;cursor:pointer;background:none">
    </label>
  </div>
  <div class="tg">
    <button class="btn" onclick="copyAll()" title="Copiar texto">&#10697;</button>
    <button class="btn" onclick="printDoc()" title="Imprimir / Guardar PDF">&#128438;</button>
    <button class="btn" onclick="undo()" title="Deshacer">&#8617;</button>
    <button class="btn" onclick="redo()" title="Rehacer">&#8618;</button>
  </div>
</div>

<!-- EDITOR SCROLLABLE -->
<div class="scroll-wrap" id="scrollWrap">
  <div class="paper" id="paper" contenteditable="true" spellcheck="false">
    {contenido if contenido else '<p style="color:{t[&apos;muted&apos;]}">Genera un informe o escribe directamente aquí...</p>'}
  </div>
</div>

<!-- BARRA INFERIOR -->
<div class="bbar">
  <span class="bstat" id="wc">—</span>
  <div class="btrack"><div class="bfill" id="bf" style="width:0%"></div></div>
  <span class="bpct" id="bp">0%</span>
</div>

<script>
var paper = document.getElementById('paper');

function fmt(cmd, val) {{
  paper.focus();
  document.execCommand(cmd, false, val || null);
  sync();
}}
function undo() {{ paper.focus(); document.execCommand('undo'); sync(); }}
function redo() {{ paper.focus(); document.execCommand('redo'); sync(); }}

function setFont(f) {{
  paper.style.fontFamily = f;
}}
function setFontSize(s) {{
  paper.style.fontSize = s + 'px';
  paper.style.lineHeight = (parseFloat(s) <= 12 ? 1.9 : 1.78).toString();
}}

function insHR() {{
  paper.focus();
  document.execCommand('insertHTML', false,
    '<hr style="border:none;border-top:1px solid {t["border"]};margin:12px 0"><br>');
}}

function insTable() {{
  var r = parseInt(prompt('Número de filas:', '3')) || 3;
  var c = parseInt(prompt('Número de columnas:', '3')) || 3;
  var h = '<table><thead><tr>';
  for (var i=0;i<c;i++) h += '<th>Col ' + (i+1) + '</th>';
  h += '</tr></thead><tbody>';
  for (var j=0;j<r-1;j++) {{
    h += '<tr>';
    for (var k=0;k<c;k++) h += '<td>&nbsp;</td>';
    h += '</tr>';
  }}
  h += '</tbody></table><p><br></p>';
  paper.focus();
  document.execCommand('insertHTML', false, h);
}}

function sync() {{
  var secs = ['TÉCNICA','HALLAZGOS','IMPRESIÓN'].filter(function(s){{
    return paper.innerText.toUpperCase().includes(s);
  }}).length;
  var w = paper.innerText.trim().split(/[ \t\n]+/).filter(Boolean).length;
  var p = Math.min(100, Math.round((secs/3)*60 + Math.min(w/150,1)*40));
  document.getElementById('bf').style.width = p + '%';
  document.getElementById('bp').textContent = p + '%';
  document.getElementById('wc').textContent = w + ' pal.';

  // Toggle bold/italic/underline
  ['Bold','Italic','Underline'].forEach(function(c) {{
    var b = document.getElementById('btn'+c[0]);
    if(b) b.classList.toggle('on', document.queryCommandState(c.toLowerCase()));
  }});
}}

function copyAll() {{
  var text = paper.innerText;
  if (navigator.clipboard) {{
    navigator.clipboard.writeText(text).then(function(){{ toast('✓ Copiado'); }});
  }} else {{
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.style.cssText = 'position:fixed;opacity:0';
    document.body.appendChild(ta); ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    toast('✓ Copiado');
  }}
}}

function printDoc() {{
  var w = window.open('','_blank');
  w.document.write('<html><head><title>AURA — Informe Radiológico</title>');
  w.document.write('<style>');
  w.document.write('body{{font-family:Calibri,sans-serif;font-size:12pt;line-height:1.75;margin:2.5cm;color:#111}}');
  w.document.write('h2{{font-size:13pt;font-weight:600;text-transform:uppercase;letter-spacing:.06em;');
  w.document.write('border-bottom:1px solid #ccc;padding-bottom:4px;margin:18px 0 6px;color:#1a1a1a}}');
  w.document.write('li{{margin-left:18px;margin-bottom:3px}}');
  w.document.write('table{{border-collapse:collapse;width:100%;margin:10px 0}}');
  w.document.write('td,th{{border:1px solid #ccc;padding:6px 10px}}');
  w.document.write('th{{background:#f0f4f8;font-weight:600}}');
  w.document.write('@media print{{body{{margin:2cm}}}}');
  w.document.write('</style></head><body>');
  w.document.write(paper.innerHTML);
  w.document.write('</body></html>');
  w.document.close();
  setTimeout(function(){{ w.print(); }}, 400);
}}

function toast(msg) {{
  var el = document.createElement('div');
  el.textContent = msg;
  el.style.cssText = 'position:fixed;bottom:44px;left:50%;transform:translateX(-50%);'
    +'background:{t["surface"]};color:{t["accent"]};border:1px solid {t["accent"]}50;'
    +'padding:5px 14px;border-radius:5px;font-size:11px;z-index:9999;pointer-events:none;'
    +'font-family:DM Sans,sans-serif;';
  document.body.appendChild(el);
  setTimeout(function(){{ document.body.removeChild(el); }}, 1600);
}}

paper.addEventListener('input', sync);
paper.addEventListener('keyup', sync);
paper.addEventListener('mouseup', sync);
paper.addEventListener('keydown', function(e) {{
  if (e.key === 'Tab') {{
    e.preventDefault();
    document.execCommand('insertHTML', false, '&nbsp;&nbsp;&nbsp;&nbsp;');
  }}
}});

window.addEventListener('load', function() {{
  sync();
  paper.focus();
}});
</script>
</body></html>"""

    components.html(editor_html, height=h_editor+80, scrolling=False)

    # ── Acciones bajo el editor ──────────────────────────────
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        if st.button("✦  Optimizar conclusión", use_container_width=True):
            if rep and api_key:
                cl = get_client(); mid = MODELS[st.session_state.modelo]["id"]
                with st.spinner("Optimizando..."):
                    try:
                        estilo_ctx = build_estilo_context()
                        estilo_note = f"\nESTILO DEL RADIÓLOGO (aplica):\n{estilo_ctx}" if estilo_ctx else ""
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":
                                f"""Eres un radiólogo subespecialista senior revisando la IMPRESIÓN DIAGNÓSTICA.

REGLAS DE FORMATO — OBLIGATORIAS:
· CERO asteriscos. Sin markdown. Títulos en MAYÚSCULAS.
· Viñetas: usa • únicamente.
· Devuelve el informe COMPLETO. No alteres TÉCNICA ni HALLAZGOS.
· Los HALLAZGOS deben ser texto corrido, no listas.

CRITERIOS DE EXCELENCIA:
· Jerarquía: hallazgo principal → secundarios → incidentales.
· Cada • : estructura + diagnóstico específico + clasificación/grado + implicación clínica.
· Última •: correlación clínico-radiológica + orientación de manejo concreta.
· Sin hedge words salvo diferencial genuino y justificado.
· Lenguaje de fellow/subespecialista: preciso, elegante, no robótico.
{estilo_note}

INFORME:
{st.session_state.reporte}"""}],
                            temperature=0.18, max_tokens=3000
                        )
                        st.session_state.reporte = limpiar(r.choices[0].message.content)
                        st.rerun()
                    except Exception as e: st.error(str(e))

    with c2:
        if st.button("◇  Análisis mentor", use_container_width=True):
            if rep and api_key:
                cl = get_client(); mid = MODELS[st.session_state.modelo]["id"]
                with st.spinner("Analizando como mentor..."):
                    try:
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":
                                f"""Eres un mentor de redacción radiológica de élite y editor académico senior.
Analiza este informe con el ojo crítico de un jefe de residentes en un centro de referencia internacional.
Responde en español. Sin asteriscos ni markdown.

Tu análisis tiene dos partes:

PARTE 1 — FEEDBACK EDITORIAL (máx. 5 puntos específicos)
Para cada punto:
· Qué: [frase o sección problemática, citada literalmente si es corta]
· Por qué es mejorable: [explicación académica concreta]
· Cómo mejorarla: [reescritura sugerida o instrucción precisa]

Detecta específicamente:
- Frases débiles, vagas o con hedge words innecesarios
- Hallazgos descritos en lista donde debería haber prosa narrativa
- Clasificaciones incompletas o ausentes donde corresponderían
- Redundancias o información repetida entre secciones
- Oportunidades para elevar el nivel semántico o la precisión anatómica
- Conclusiones que no orientan al clínico de forma accionable

PARTE 2 — EVALUACIÓN GLOBAL
· Nivel actual: [Básico / Residente / Fellow / Subespecialista / Publicable]
· Fortalezas: [2-3 líneas]
· Potencial de mejora: [2-3 líneas específicas]
· Puntuación global: [X/10] con justificación en una línea.

Sé específico, académico y constructivo. No des elogios genéricos.

INFORME:
{st.session_state.reporte}"""}],
                            temperature=0.2, max_tokens=2000
                        )
                        st.session_state.mentor_feedback = r.choices[0].message.content
                        st.session_state.panel_der = True
                        st.rerun()
                    except Exception as e: st.error(str(e))

    with c3:
        if st.button("🧠  Aprobar y aprender", use_container_width=True, help="Guarda este informe como ejemplo de tu estilo"):
            if rep:
                guardar_estilo(rep)
                st.success(f"✓ Guardado. {len(st.session_state.estilo_aprendido)} ejemplo(s) en memoria.")

    with c4:
        if rep:
            st.download_button(
                "↓  Exportar .docx",
                data=generar_docx(st.session_state.reporte),
                file_name="AURA_Informe.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    # ── Feedback del mentor ──────────────────────────────────
    if st.session_state.mentor_feedback and not st.session_state.panel_der:
        st.markdown(f"""
        <div class="mentor-card">
          <div class="mentor-header">
            <div class="sec-dot"></div>ANÁLISIS DEL MENTOR
          </div>
          <div style="white-space:pre-wrap;font-size:12px;line-height:1.72">
            {st.session_state.mentor_feedback}
          </div>
        </div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# PANEL DERECHO — Análisis mentor + Definiciones
# ═══════════════════════════════════════════════════════════════
with col_r:
    if st.button("▶" if ro else "◀", key="tog_r", help="Expandir panel de análisis"):
        st.session_state.panel_der = not ro; st.rerun()

    if not ro:
        st.markdown(f"""
        <div style="display:flex;flex-direction:column;align-items:center;gap:12px;padding:10px 0">
          <span style="font-size:13px;color:{t['muted']}">🧠</span>
          <span style="font-size:13px;color:{t['muted']}">📖</span>
        </div>""", unsafe_allow_html=True)
    else:
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

        tab_mentor, tab_defs = st.tabs(["🧠 Mentor", "📖 Definiciones"])

        with tab_mentor:
            if st.session_state.mentor_feedback:
                st.markdown(f"""
                <div style="font-size:11.5px;line-height:1.7;color:{t['muted']};
                  white-space:pre-wrap;padding:4px 0">
                  {st.session_state.mentor_feedback}
                </div>""", unsafe_allow_html=True)
                if st.button("Cerrar análisis", key="close_mentor"):
                    st.session_state.mentor_feedback = ""; st.rerun()
            else:
                st.markdown(f"""
                <div style="padding:20px 0;text-align:center">
                  <p style="font-size:11px;color:{t['muted']};line-height:1.8">
                    Genera un informe y presiona<br>
                    <strong style="color:{t['text']}">◇ Análisis mentor</strong><br>
                    para recibir feedback editorial<br>de nivel académico.
                  </p>
                </div>""", unsafe_allow_html=True)

        with tab_defs:
            rep2 = st.session_state.reporte
            if st.button("Generar definiciones", use_container_width=True, key="btn_defs"):
                if rep2 and api_key:
                    cl = get_client(); mid = MODELS[st.session_state.modelo]["id"]
                    with st.spinner("Analizando..."):
                        try:
                            r = cl.chat.completions.create(
                                model=mid,
                                messages=[{"role":"user","content":
                                    f"""Analiza el informe radiológico con profundidad académica.
Responde en español. Sin asteriscos ni markdown.

CLASIFICACIONES UTILIZADAS
Para cada clasificación:
· Sistema: [nombre completo · sociedad]
· Grado asignado: [grado] — [significado clínico concreto]
· Evidencia: [hallazgo del informe que lo justifica]
· Referencia: [Autor, Revista, Año]
· Relevancia: [implicación para el manejo]

CLASIFICACIONES ADICIONALES RECOMENDADAS
[Si no hay: "El informe utiliza los sistemas apropiados."]

GLOSARIO
· [Término]: [definición precisa en 2 líneas, contexto anatómico y diagnóstico]

FISIOPATOLOGÍA
[3-4 líneas: mecanismo subyacente a los hallazgos principales. Nivel fellow.]

ORIENTACIÓN AL CLÍNICO
[4-5 líneas: implicaciones, opciones terapéuticas, estudios complementarios, seguimiento.]

INFORME:
{rep2}"""}],
                                temperature=0.15, max_tokens=2000
                            )
                            st.session_state.defs = r.choices[0].message.content
                            st.rerun()
                        except Exception as e: st.error(str(e))

            if st.session_state.defs:
                st.markdown(f'<div class="defs-box">{st.session_state.defs}</div>',
                            unsafe_allow_html=True)
                if st.button("Cerrar", key="close_defs"):
                    st.session_state.defs = ""; st.rerun()
            else:
                st.markdown(f"""
                <div style="padding:16px 0;text-align:center">
                  <p style="font-size:11px;color:{t['muted']};line-height:1.8">
                    Genera un informe y presiona<br>el botón de arriba para ver<br>
                    clasificaciones y referencias.
                  </p>
                </div>""", unsafe_allow_html=True)
