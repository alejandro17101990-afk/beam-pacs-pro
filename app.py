import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile, io, os, re

st.set_page_config(page_title="AURA", layout="wide", initial_sidebar_state="collapsed")

# ── Paleta ────────────────────────────────────────────────────
BG      = "#0b0f14"
PANEL   = "#111720"
CARD    = "#161d27"
BORDER  = "#1e2a38"
ACCENT  = "#3b9eff"
TEXT    = "#dce8f4"
MUTED   = "#5a7a96"
ED_BG   = "#0e1520"
ED_TEXT = "#c8dff0"
GREEN   = "#22c55e"

MODELS = {
    "DeepSeek Chat": {"url": "https://api.deepseek.com", "id": "deepseek-chat"},
    "GPT-4o Mini":   {"url": None, "id": "gpt-4o-mini"},
    "GPT-4.1 Mini":  {"url": None, "id": "gpt-4.1-mini"},
}
HCOLS = ["#3b9eff","#22c55e","#f59e0b","#ec4899","#8b5cf6","#06b6d4"]

# ── Estado ────────────────────────────────────────────────────
DEFAULTS = {
    "dictado": "", "reporte": "", "defs": "",
    "modelo": "DeepSeek Chat", "audio_id": None,
    "historial": [], "plantilla_txt": "",
    # expansión horizontal de paneles
    "panel_izq": True,   # panel izquierdo visible
    "panel_der": True,   # panel derecho visible
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

try:    api_key = st.secrets["deepseek_key"]
except: api_key = os.environ.get("OPENAI_API_KEY", "")

# ── Helpers ───────────────────────────────────────────────────
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
                p = pp.Paragraph(el, doc); t = p.text.strip()
                if t: partes.append(t)
            elif tag == 'tbl':
                n += 1; tbl = tt.Table(el, doc)
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
        if s.isupper() and len(s) < 80:
            h = doc.add_heading(s, level=1); h.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif s.startswith(("•","·")): doc.add_paragraph(s[1:].strip(), style="List Bullet")
        else: doc.add_paragraph(s)
    bio = io.BytesIO(); doc.save(bio); return bio.getvalue()

def transcribir(audio):
    cfg = MODELS[st.session_state.modelo]
    cl = OpenAI(api_key=api_key, base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp:
        tmp.write(audio.read()); path = tmp.name
    with open(path, "rb") as f:
        res = cl.audio.transcriptions.create(
            model="whisper-1", file=f, language="es",
            prompt="Dictado radiológico: Stoller, ICRS, LCA, menisco, condromalacia, osteofito, Kellgren-Lawrence."
        )
    os.unlink(path); return res.text.strip()

def completitud(texto):
    secs  = sum(1 for s in ["TÉCNICA","HALLAZGOS","IMPRESIÓN"] if s in texto.upper())
    words = len(texto.split())
    return min(100, int((secs/3)*60 + min(words/150,1)*40)), words

# ── CSS ───────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

html, body, .stApp {{ background:{BG}; color:{TEXT}; font-family:'Inter',sans-serif; }}
header, footer, #MainMenu {{ visibility:hidden; }}
.block-container {{ padding:0 !important; max-width:100% !important; }}
* {{ box-sizing:border-box; }}

/* ── TOPBAR ── */
.topbar {{
    height:52px; background:{PANEL}; border-bottom:1px solid {BORDER};
    display:flex; align-items:center; padding:0 24px; gap:16px;
    position:sticky; top:0; z-index:100;
}}
.logo {{ font-size:18px; font-weight:600; color:{ACCENT}; letter-spacing:.12em;
         display:flex; align-items:center; gap:9px; }}
.logo-dot {{ width:8px; height:8px; border-radius:50%; background:{ACCENT};
             animation:dp 2s ease-in-out infinite; }}
@keyframes dp {{ 0%,100%{{opacity:1}} 50%{{opacity:.2}} }}
.t-sep {{ width:1px; height:18px; background:{BORDER}; }}
.t-meta {{ font-size:12px; color:{MUTED}; }}
.t-badge {{ font-size:11px; color:{ACCENT}; background:{ACCENT}18;
            border:1px solid {ACCENT}40; border-radius:6px; padding:3px 10px; }}
.t-right {{ margin-left:auto; display:flex; align-items:center; gap:10px; }}
.t-dot-on {{ width:6px; height:6px; border-radius:50%;
             background:{GREEN}; box-shadow:0 0 5px {GREEN}; }}

/* ── LAYOUT ── */
.main-wrap {{
    display:flex; height:calc(100vh - 52px); overflow:hidden;
}}

/* Panel izquierdo */
.pane-left {{
    display:flex; flex-direction:column;
    background:{PANEL}; border-right:1px solid {BORDER};
    transition:width .25s ease; overflow:hidden; flex-shrink:0;
}}
.pane-left.open  {{ width:380px; min-width:380px; }}
.pane-left.closed{{ width:48px;  min-width:48px;  }}

/* Panel central (editor) */
.pane-center {{
    flex:1; display:flex; flex-direction:column;
    background:{BG}; overflow:hidden; min-width:0;
}}

/* Panel derecho */
.pane-right {{
    display:flex; flex-direction:column;
    background:{PANEL}; border-left:1px solid {BORDER};
    transition:width .25s ease; overflow:hidden; flex-shrink:0;
}}
.pane-right.open  {{ width:320px; min-width:320px; }}
.pane-right.closed{{ width:48px;  min-width:48px;  }}

/* Toggle button para paneles */
.pane-toggle {{
    display:flex; align-items:center; justify-content:center;
    height:36px; cursor:pointer;
    background:transparent; border:none; color:{MUTED};
    font-size:16px; transition:color .15s;
    padding:0; width:100%;
}}
.pane-toggle:hover {{ color:{TEXT}; }}

/* Scrollable interior del panel */
.pane-scroll {{
    flex:1; overflow-y:auto; padding:16px;
    scrollbar-width:thin; scrollbar-color:{BORDER} transparent;
}}
.pane-scroll::-webkit-scrollbar {{ width:3px; }}
.pane-scroll::-webkit-scrollbar-thumb {{ background:{BORDER}; border-radius:2px; }}

/* Collapsed icon strip */
.icon-strip {{
    display:flex; flex-direction:column; align-items:center;
    gap:18px; padding:16px 0;
}}
.icon-strip span {{
    font-size:16px; color:{MUTED}; cursor:pointer; transition:color .15s;
    writing-mode:horizontal-tb;
}}
.icon-strip span:hover {{ color:{ACCENT}; }}

/* ── SECCIÓN DENTRO DE PANEL ── */
.section {{
    background:{CARD}; border:1px solid {BORDER};
    border-radius:12px; padding:16px; margin-bottom:12px;
}}
.section-title {{
    font-size:11px; font-weight:600; letter-spacing:.1em;
    text-transform:uppercase; color:{MUTED}; margin-bottom:12px;
    display:flex; align-items:center; gap:6px;
}}
.section-title .dot {{ width:6px; height:6px; border-radius:50%; background:{ACCENT}; }}

/* ── INPUTS ── */
[data-testid="stSelectbox"] > div > div {{
    background:{CARD} !important; border:1px solid {BORDER} !important;
    border-radius:8px !important; color:{TEXT} !important; font-size:13px !important;
}}
[data-testid="stSelectbox"] > div > div:hover {{ border-color:{ACCENT}50 !important; }}

.stTextArea textarea {{
    background:{ED_BG} !important; border:1px solid {BORDER} !important;
    border-radius:10px !important; color:{ED_TEXT} !important;
    font-size:13.5px !important; line-height:1.75 !important;
    padding:14px !important; caret-color:{ACCENT} !important;
    font-family:'Inter',sans-serif !important;
}}
.stTextArea textarea:focus {{
    border-color:{ACCENT}50 !important;
    box-shadow:0 0 0 3px {ACCENT}12 !important;
}}
.stTextArea textarea::placeholder {{ color:{MUTED} !important; }}

[data-testid="stAudioInput"] {{
    background:{CARD} !important; border:1px solid {BORDER} !important;
    border-radius:10px !important;
}}
[data-testid="stFileUploader"] {{
    background:{CARD}; border:1px dashed {BORDER}; border-radius:10px; padding:6px;
}}
[data-testid="stFileUploader"] * {{ color:{MUTED} !important; font-size:12px !important; }}

/* ── BOTONES ── */
.stButton button {{
    background:{CARD} !important; border:1px solid {BORDER} !important;
    color:{TEXT} !important; border-radius:8px !important;
    font-size:13px !important; font-weight:500 !important; transition:all .15s !important;
}}
.stButton button:hover {{
    border-color:{ACCENT}60 !important; background:{PANEL} !important;
}}
.btn-primary .stButton button {{
    background:{ACCENT} !important; border-color:{ACCENT} !important;
    color:#fff !important; font-weight:600 !important;
}}
.btn-primary .stButton button:hover {{ opacity:.88 !important; }}

.stDownloadButton button {{
    background:transparent !important; border:1px solid {ACCENT} !important;
    color:{ACCENT} !important; border-radius:8px !important; font-size:13px !important;
}}
.stDownloadButton button:hover {{ background:{ACCENT}18 !important; }}

/* ── TABS ── */
[data-testid="stTabs"] [role="tablist"] {{
    border-bottom:1px solid {BORDER} !important; gap:0 !important;
    background:transparent !important;
}}
[data-testid="stTabs"] [role="tab"] {{
    background:transparent !important; border:none !important;
    color:{MUTED} !important; font-size:13px !important;
    padding:8px 14px !important; border-radius:0 !important;
    border-bottom:2px solid transparent !important;
}}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {{
    color:{ACCENT} !important; border-bottom-color:{ACCENT} !important;
}}
[data-testid="stTabs"] [data-baseweb="tab-panel"] {{
    background:transparent !important; padding:12px 0 0 !important;
}}

/* ── EXPANDER ── */
[data-testid="stExpander"] {{
    background:{CARD} !important; border:1px solid {BORDER} !important;
    border-radius:10px !important; margin-bottom:8px !important;
}}
[data-testid="stExpander"] summary {{
    color:{MUTED} !important; font-size:13px !important;
}}
[data-testid="stExpander"] summary:hover {{ color:{TEXT} !important; }}

/* ── BARRA COMPLETITUD ── */
.prog-row {{ display:flex; align-items:center; gap:10px; margin-bottom:12px; }}
.prog-bg  {{ flex:1; height:3px; background:{BORDER}; border-radius:2px; overflow:hidden; }}
.prog-fill{{ height:100%; background:{ACCENT}; border-radius:2px; transition:width .4s; }}
.prog-txt {{ font-size:11px; color:{MUTED}; white-space:nowrap; }}

/* ── HISTORIAL ── */
.h-row {{ display:flex; align-items:center; gap:8px;
          padding:7px 10px; border-radius:8px;
          background:{PANEL}; border:1px solid {BORDER}; margin-bottom:4px; }}
.h-dot {{ width:8px; height:8px; border-radius:50%; flex-shrink:0; }}
.h-name {{ font-size:12px; color:{TEXT}; }}
.h-sub  {{ font-size:11px; color:{MUTED}; }}

/* ── DEFS ── */
.defs-box {{
    background:{ED_BG}; border:1px solid {BORDER}; border-radius:10px;
    padding:14px; font-size:12.5px; line-height:1.6; color:{MUTED}; white-space:pre-wrap;
}}

::-webkit-scrollbar {{ width:3px; }}
::-webkit-scrollbar-thumb {{ background:{BORDER}; border-radius:2px; }}
hr {{ border:none; border-top:1px solid {BORDER} !important; margin:12px 0 !important; }}

/* Forzar columnas de Streamlit sin padding extra */
[data-testid="column"] {{ padding:0 !important; }}
</style>
""", unsafe_allow_html=True)

# ── TOPBAR ────────────────────────────────────────────────────
ok = bool(api_key)
st.markdown(f"""
<div class="topbar">
  <div class="logo"><div class="logo-dot"></div>AURA</div>
  <div class="t-sep"></div>
  <span class="t-meta">Radiology Intelligence</span>
  <div class="t-right">
    <span class="t-badge">{st.session_state.modelo}</span>
    <div class="t-dot-on" title="{'Conectado' if ok else 'Sin API Key'}"></div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── CÁLCULO DE ANCHOS SEGÚN ESTADO DE PANELES ─────────────────
left_open   = st.session_state.panel_izq
right_open  = st.session_state.panel_der

# Ratios: panel colapsado = 0.04 (~48px), expandido reparte el espacio
if left_open and right_open:
    ratios = [1.05, 2.5, 0.95]
elif left_open and not right_open:
    ratios = [1.05, 3.3, 0.08]
elif not left_open and right_open:
    ratios = [0.08, 3.3, 0.95]
else:
    ratios = [0.08, 5.0, 0.08]

col_l, col_c, col_r = st.columns(ratios, gap="small")

# ════════════════════════════════════════════════════════════
# PANEL IZQUIERDO — Entrada + Historial
# ════════════════════════════════════════════════════════════
with col_l:
    # Botón de toggle horizontal (colapsar/expandir)
    toggle_lbl = "◀" if left_open else "▶"
    st.markdown(f"""
    <div style="display:flex;justify-content:{'flex-end' if left_open else 'center'};
                padding:10px 10px 0;border-bottom:1px solid {BORDER};padding-bottom:8px">
    </div>""", unsafe_allow_html=True)

    tl_col, _ = st.columns([1, 0.01]) if left_open else st.columns([1, 0.01])
    if st.button(toggle_lbl, key="tog_l", help="Expandir / colapsar panel de entrada"):
        st.session_state.panel_izq = not left_open
        st.rerun()

    if not left_open:
        # Panel colapsado: iconos verticales
        st.markdown(f"""
        <div class="icon-strip">
          <span title="Entrada">🎙</span>
          <span title="Historial">📋</span>
          <span title="Config">⚙</span>
        </div>""", unsafe_allow_html=True)
    else:
        # ── Configuración de estudio ──
        st.markdown(f"""
        <div class="section-title" style="margin-top:8px">
          <div class="dot"></div>Estudio
        </div>""", unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            modalidades = ["RM","TC","Rx","US","PET-CT"]
            mod_full    = ["Resonancia Magnética","Tomografía Computarizada",
                           "Radiografía","Ultrasonido","PET-CT"]
            st.markdown(f'<p style="font-size:11px;color:{MUTED};margin-bottom:4px">MODALIDAD</p>',
                        unsafe_allow_html=True)
            idx_m = st.selectbox("mod", mod_full, label_visibility="collapsed", key="sel_mod")
        with c2:
            regiones = ["Rodilla","Col. lumbar","Col. cervical","Hombro","Cadera",
                        "Tobillo","Muñeca","Codo","Cerebro","Tórax","Abdomen","Mama","Tiroides","Hígado"]
            st.markdown(f'<p style="font-size:11px;color:{MUTED};margin-bottom:4px">REGIÓN</p>',
                        unsafe_allow_html=True)
            idx_r = st.selectbox("reg", regiones, label_visibility="collapsed", key="sel_reg")

        st.markdown("<hr>", unsafe_allow_html=True)

        # ── Entrada de dictado ──
        st.markdown(f"""
        <div class="section-title">
          <div class="dot"></div>Dictado
        </div>""", unsafe_allow_html=True)

        tab_voz, tab_texto = st.tabs(["🎙 Voz", "⌨ Texto"])

        with tab_voz:
            # Micrófono visual
            st.markdown(f"""
            <div style="display:flex;flex-direction:column;align-items:center;padding:20px 0 14px;gap:10px">
              <div style="position:relative;width:76px;height:76px">
                <div style="position:absolute;inset:-14px;border-radius:50%;
                  border:1.5px solid {ACCENT}28;
                  animation:rp 2.4s ease-out infinite"></div>
                <div style="position:absolute;inset:-7px;border-radius:50%;
                  border:1.5px solid {ACCENT}45;
                  animation:rp 2.4s ease-out infinite .5s"></div>
                <div style="width:76px;height:76px;border-radius:50%;
                  background:{CARD};border:2px solid {ACCENT};
                  display:flex;align-items:center;justify-content:center">
                  <svg width="26" height="26" viewBox="0 0 24 24" fill="none"
                       stroke="{ACCENT}" stroke-width="1.8"
                       stroke-linecap="round" stroke-linejoin="round">
                    <rect x="9" y="2" width="6" height="12" rx="3"/>
                    <path d="M5 10a7 7 0 0 0 14 0"/>
                    <line x1="12" y1="19" x2="12" y2="22"/>
                    <line x1="9"  y1="22" x2="15" y2="22"/>
                  </svg>
                </div>
              </div>
              <span style="font-size:11px;color:{MUTED}">Pulsa para grabar</span>
            </div>
            <style>
            @keyframes rp {{
              0%   {{ transform:scale(1);   opacity:.6 }}
              100% {{ transform:scale(1.35);opacity:0  }}
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
                        st.info("Ingresa tu API Key en Configuración.")

        with tab_texto:
            st.markdown(f'<p style="font-size:12px;color:{MUTED};margin-bottom:6px">'
                        'Escribe hallazgos directamente.</p>', unsafe_allow_html=True)

        # Señal transcrita
        st.markdown(f'<p style="font-size:11px;color:{MUTED};margin:8px 0 4px">'
                    'SEÑAL DE ENTRADA</p>', unsafe_allow_html=True)
        dictado = st.text_area(
            "d", value=st.session_state.dictado, height=200,
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
            generar = st.button("Generar informe", use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with bb:
            if st.button("Limpiar", use_container_width=True):
                st.session_state.dictado = ""
                st.session_state.audio_id = None
                st.rerun()

        st.markdown("<hr>", unsafe_allow_html=True)

        # ── Configuración avanzada ──
        with st.expander("⚙  Configuración avanzada", expanded=False):
            st.markdown(f'<p style="font-size:11px;color:{MUTED};margin-bottom:4px">MODELO IA</p>',
                        unsafe_allow_html=True)
            m = st.selectbox("m", list(MODELS.keys()),
                             index=list(MODELS.keys()).index(st.session_state.modelo),
                             label_visibility="collapsed")
            if m != st.session_state.modelo:
                st.session_state.modelo = m; st.rerun()

            st.markdown(f'<p style="font-size:11px;color:{MUTED};margin:8px 0 4px">PLANTILLA .DOCX</p>',
                        unsafe_allow_html=True)
            f_up = st.file_uploader("plt", type=["docx"], label_visibility="collapsed")
            if f_up:
                st.session_state.plantilla_txt, _ = leer_plantilla(f_up)
                st.markdown(f'<p style="font-size:11px;color:{ACCENT}">✓ Plantilla cargada</p>',
                            unsafe_allow_html=True)

            if not api_key:
                st.markdown(f'<p style="font-size:11px;color:{MUTED};margin:8px 0 4px">API KEY</p>',
                            unsafe_allow_html=True)
                api_key = st.text_input("k", type="password", label_visibility="collapsed",
                                        placeholder="sk- ···")

        # ── Historial ──
        if st.session_state.historial:
            with st.expander(f"📋  Historial  ·  {len(st.session_state.historial)}", expanded=False):
                for i, e in enumerate(st.session_state.historial):
                    color = HCOLS[i % len(HCOLS)]
                    st.markdown(f"""
                    <div class="h-row">
                      <div class="h-dot" style="background:{color}"></div>
                      <div>
                        <div class="h-name">{e['region']}</div>
                        <div class="h-sub">{e['modalidad'][:18]}</div>
                      </div>
                    </div>""", unsafe_allow_html=True)
                    if st.button(f"Cargar", key=f"h{i}", use_container_width=True):
                        st.session_state.reporte = e['texto']
                        st.rerun()

    if not left_open:
        generar = False   # sin botón cuando el panel está colapsado

# ── PROCESAMIENTO ─────────────────────────────────────────────
if generar:
    if not api_key:
        st.warning("Ingresa tu API Key en Configuración avanzada.")
    elif not st.session_state.dictado.strip():
        st.warning("Escribe o dicta hallazgos primero.")
    else:
        cl  = get_client()
        mid = MODELS[st.session_state.modelo]["id"]
        pt  = st.session_state.plantilla_txt
        instruc_tabla = (
            "La plantilla tiene tablas [TABLA]. Complétalas en Markdown."
            if "[TABLA" in pt else
            "NO generes tablas bajo ninguna circunstancia."
        )
        try:
            mod_sel = idx_m
            reg_sel = idx_r
        except:
            mod_sel = ""; reg_sel = ""

        prompt = f"""Eres AURA, asistente de interpretación radiológica de alta precisión.
Genera un informe radiológico estructurado y profesional.

MODALIDAD: {mod_sel}
REGIÓN: {reg_sel}

REGLAS:
· Lenguaje médico preciso. Sin ambigüedad.
· PROHIBIDO: "cambios degenerativos" sin sustrato morfológico. Usa descriptores específicos.
· Solo clasificaciones respaldadas por los hallazgos.
· {instruc_tabla}
· Sin markdown. Títulos en MAYÚSCULAS. Usa • para viñetas en la impresión.

PLANTILLA:
{pt if pt else "INDICACIÓN\\nTÉCNICA\\nHALLAZGOS\\nIMPRESIÓN DIAGNÓSTICA"}

DICTADO:
{st.session_state.dictado}"""

        with st.spinner("Generando informe..."):
            try:
                res = cl.chat.completions.create(
                    model=mid,
                    messages=[{"role":"system","content":prompt}],
                    temperature=0.1, max_tokens=2500
                )
                report = res.choices[0].message.content
                st.session_state.reporte = report
                st.session_state.historial.insert(0, {
                    "modalidad": mod_sel[:18] if mod_sel else "RM",
                    "region":    reg_sel if reg_sel else "General",
                    "texto":     report
                })
                if len(st.session_state.historial) > 12:
                    st.session_state.historial = st.session_state.historial[:12]
                st.rerun()
            except Exception as e:
                st.error(str(e))

# ════════════════════════════════════════════════════════════
# PANEL CENTRAL — Editor
# ════════════════════════════════════════════════════════════
with col_c:
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    rep = st.session_state.reporte

    # Barra superior del editor
    if rep:
        pct, words = completitud(rep)
        st.markdown(f"""
        <div class="prog-row">
          <div class="prog-bg"><div class="prog-fill" style="width:{pct}%"></div></div>
          <span class="prog-txt">{pct}% · {words} palabras</span>
        </div>""", unsafe_allow_html=True)
    else:
        st.markdown(f'<p style="font-size:13px;color:{MUTED};margin-bottom:12px">'
                    'El informe generado aparece aquí.</p>', unsafe_allow_html=True)

    # Calcular altura del textarea según espacio disponible
    # Ocupa casi toda la pantalla, ajustado al panel
    editor_h = 580 if (left_open and right_open) else 620

    reporte = st.text_area(
        "informe",
        value=rep,
        height=editor_h,
        label_visibility="collapsed",
        placeholder=(
            "Genera un informe para comenzar a editar.\n\n"
            "Puedes modificar libremente cualquier sección.\n"
            "Fondo oscuro — sin deslumbramiento."
        ),
        key="reporte_ta"
    )
    if reporte != st.session_state.reporte:
        st.session_state.reporte = reporte

    # Acciones bajo el editor
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    a1, a2, a3 = st.columns(3)

    with a1:
        if st.button("✦ Optimizar conclusión", use_container_width=True):
            if rep and api_key:
                cl = get_client()
                mid = MODELS[st.session_state.modelo]["id"]
                with st.spinner("Optimizando..."):
                    try:
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":
                                f"""Mejora ÚNICAMENTE la IMPRESIÓN DIAGNÓSTICA.
· Morfológicamente precisa y clínicamente accionable.
· Solo clasificaciones con evidencia directa en los hallazgos.
· Usa "•" para viñetas. Lenguaje sugerente para manejo.
· Devuelve el informe COMPLETO. Sin asteriscos. Títulos en MAYÚSCULAS.
INFORME:
{st.session_state.reporte}"""}],
                            temperature=0.2, max_tokens=2500
                        )
                        st.session_state.reporte = r.choices[0].message.content
                        st.rerun()
                    except Exception as e: st.error(str(e))

    with a2:
        if st.button("◇ Definiciones", use_container_width=True):
            if rep and api_key:
                cl = get_client()
                mid = MODELS[st.session_state.modelo]["id"]
                with st.spinner("Analizando..."):
                    try:
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":
                                f"""Analiza el informe. Formato exacto.
Sin líneas en blanco entre ítems. Una línea entre secciones.

CLASIFICACIONES USADAS
· Nombre: [nombre · autor/sociedad]
· Grado: [grado] — [significado clínico]
· Justificación: [hallazgo del texto]
· Ref: [Autor, año, revista]
· URL: [PubMed o sociedad]

CLASIFICACIONES SUGERIDAS
[Solo si hay hallazgo directo. Si no: "Ninguna adicional justificada."]
· Nombre / Hallazgo / Ref / URL

DEFINICIONES
· [Término]: [1-2 líneas]

CORRELACIÓN CLÍNICA
[2-3 líneas. Lenguaje sugerente.]

Sin asteriscos.
INFORME:
{st.session_state.reporte}"""}],
                            temperature=0.15, max_tokens=2000
                        )
                        st.session_state.defs = r.choices[0].message.content
                        # Abrir panel derecho automáticamente
                        st.session_state.panel_der = True
                        st.rerun()
                    except Exception as e: st.error(str(e))

    with a3:
        if rep:
            st.download_button(
                "↓ Exportar .docx",
                data=generar_docx(st.session_state.reporte),
                file_name="AURA_Informe.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

# ════════════════════════════════════════════════════════════
# PANEL DERECHO — Definiciones / Referencias
# ════════════════════════════════════════════════════════════
with col_r:
    # Toggle
    toggle_r = "▶" if right_open else "◀"
    if st.button(toggle_r, key="tog_r", help="Expandir / colapsar panel de definiciones"):
        st.session_state.panel_der = not right_open
        st.rerun()

    if not right_open:
        st.markdown(f"""
        <div class="icon-strip">
          <span title="Definiciones">📖</span>
          <span title="Referencias">🔗</span>
        </div>""", unsafe_allow_html=True)
    else:
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        st.markdown(f"""
        <div class="section-title">
          <div class="dot"></div>Definiciones y referencias
        </div>""", unsafe_allow_html=True)

        if st.session_state.defs:
            st.markdown(
                f'<div class="defs-box">{st.session_state.defs}</div>',
                unsafe_allow_html=True
            )
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
            if st.button("Cerrar", key="close_defs"):
                st.session_state.defs = ""; st.rerun()
        else:
            st.markdown(f"""
            <div style="padding:20px 0;text-align:center">
              <p style="font-size:13px;color:{MUTED};line-height:1.6">
                Genera un informe y presiona<br>
                <strong style="color:{TEXT}">◇ Definiciones</strong><br>
                para ver clasificaciones,<br>definiciones y referencias.
              </p>
            </div>""", unsafe_allow_html=True)
