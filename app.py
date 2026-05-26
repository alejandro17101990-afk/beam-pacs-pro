import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile, io, os, re

st.set_page_config(page_title="AURA", layout="wide", initial_sidebar_state="collapsed")

# ── Paleta ──────────────────────────────────────────────────
BG      = "#0b0f14"
PANEL   = "#111720"
CARD    = "#161d27"
BORDER  = "#1e2a38"
ACCENT  = "#3b9eff"
TEXT    = "#dce8f4"
MUTED   = "#5a7a96"
GHOST   = "#1e3450"
ED_BG   = "#0e1520"   # fondo editor — oscuro suave, sin deslumbramiento
ED_TEXT = "#c8dff0"

# ── Modelos ─────────────────────────────────────────────────
MODELS = {
    "DeepSeek Chat": {"url": "https://api.deepseek.com", "id": "deepseek-chat"},
    "GPT-4o Mini":   {"url": None,                       "id": "gpt-4o-mini"},
    "GPT-4.1 Mini":  {"url": None,                       "id": "gpt-4.1-mini"},
}

# ── Estado ───────────────────────────────────────────────────
for k, v in {"dictado": "", "reporte": "", "defs": "",
             "modelo": "DeepSeek Chat", "audio_id": None,
             "historial": [], "plantilla_txt": ""}.items():
    if k not in st.session_state:
        st.session_state[k] = v

try:    api_key = st.secrets["deepseek_key"]
except: api_key = os.environ.get("OPENAI_API_KEY", "")

# ── Helpers ──────────────────────────────────────────────────
def get_client():
    cfg = MODELS[st.session_state.modelo]
    return OpenAI(api_key=api_key, base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)

def leer_plantilla(f):
    doc = Document(f)
    partes = []; n = 0
    try:
        import docx.text.paragraph as pp, docx.table as tt
        for el in doc.element.body:
            tag = el.tag.split('}')[-1]
            if tag == 'p':
                p = pp.Paragraph(el, doc); t = p.text.strip()
                if t: partes.append(t)
            elif tag == 'tbl':
                n += 1; tbl = tt.Table(el, doc)
                rows = ["| " + " | ".join(c.text.strip() for c in r.cells) + " |"
                        for r in tbl.rows]
                partes.append(f"[TABLA {n}]\n" + "\n".join(rows) + "\n[/TABLA]")
    except:
        partes = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    tiene_tabla = n > 0
    return "\n".join(partes), tiene_tabla

def generar_docx(texto):
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    for line in texto.split("\n"):
        s = line.strip()
        if not s: doc.add_paragraph(); continue
        if s.isupper() and len(s) < 80:
            h = doc.add_heading(s, level=1)
            h.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif s.startswith(("•", "·")):
            doc.add_paragraph(s[1:].strip(), style="List Bullet")
        else:
            doc.add_paragraph(s)
    bio = io.BytesIO(); doc.save(bio); return bio.getvalue()

def transcribir(audio):
    cfg = MODELS[st.session_state.modelo]
    cl = OpenAI(api_key=api_key, base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp:
        tmp.write(audio.read()); path = tmp.name
    with open(path, "rb") as f:
        res = cl.audio.transcriptions.create(
            model="whisper-1", file=f, language="es",
            prompt="Dictado radiológico. Términos: Stoller, ICRS, LCA, menisco, condromalacia, osteofito, Kellgren-Lawrence."
        )
    os.unlink(path); return res.text.strip()

# ── CSS ──────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

html, body, .stApp {{
    background: {BG};
    color: {TEXT};
    font-family: 'Inter', sans-serif;
}}
header, footer, #MainMenu {{ visibility: hidden; }}
.block-container {{ padding: 1.5rem 2rem 1rem; max-width: 100%; }}

/* ── Selectbox ── */
[data-testid="stSelectbox"] > div > div {{
    background: {PANEL} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 10px !important;
    color: {TEXT} !important;
    font-size: 14px !important;
}}
[data-testid="stSelectbox"] > div > div:hover {{ border-color: {ACCENT}60 !important; }}

/* ── Textarea ── */
.stTextArea textarea {{
    background: {ED_BG} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 12px !important;
    color: {ED_TEXT} !important;
    font-size: 14px !important;
    line-height: 1.75 !important;
    padding: 16px !important;
    caret-color: {ACCENT} !important;
}}
.stTextArea textarea:focus {{
    border-color: {ACCENT}50 !important;
    box-shadow: 0 0 0 3px {ACCENT}15 !important;
}}
.stTextArea textarea::placeholder {{ color: {MUTED} !important; }}

/* ── Audio input ── */
[data-testid="stAudioInput"] {{
    background: {PANEL} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 12px !important;
}}

/* ── File uploader ── */
[data-testid="stFileUploader"] {{
    background: {PANEL};
    border: 1px dashed {BORDER};
    border-radius: 12px;
    padding: 8px;
}}
[data-testid="stFileUploader"] * {{ color: {MUTED} !important; font-size: 13px !important; }}

/* ── Botones secundarios ── */
.stButton button {{
    background: {CARD} !important;
    border: 1px solid {BORDER} !important;
    color: {TEXT} !important;
    border-radius: 10px !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    transition: all .15s !important;
}}
.stButton button:hover {{
    border-color: {ACCENT}60 !important;
    color: white !important;
}}

/* ── Botón primario (wrapper) ── */
.btn-primary .stButton button {{
    background: {ACCENT} !important;
    border-color: {ACCENT} !important;
    color: white !important;
    font-weight: 600 !important;
}}
.btn-primary .stButton button:hover {{
    opacity: .88 !important;
}}

/* ── Download ── */
.stDownloadButton button {{
    background: transparent !important;
    border: 1px solid {ACCENT} !important;
    color: {ACCENT} !important;
    border-radius: 10px !important;
    font-size: 13px !important;
    font-weight: 500 !important;
}}
.stDownloadButton button:hover {{
    background: {ACCENT}20 !important;
}}

/* ── Expander ── */
[data-testid="stExpander"] {{
    background: {PANEL} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 12px !important;
}}
[data-testid="stExpander"] summary {{
    color: {MUTED} !important;
    font-size: 13px !important;
}}
[data-testid="stExpander"] summary:hover {{ color: {TEXT} !important; }}

/* ── Tabs ── */
[data-testid="stTabs"] [role="tablist"] {{
    border-bottom: 1px solid {BORDER} !important;
    gap: 0 !important;
}}
[data-testid="stTabs"] [role="tab"] {{
    background: transparent !important;
    border: none !important;
    color: {MUTED} !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    padding: 8px 16px !important;
    border-bottom: 2px solid transparent !important;
    border-radius: 0 !important;
}}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {{
    color: {ACCENT} !important;
    border-bottom-color: {ACCENT} !important;
}}
[data-testid="stTabs"] [data-baseweb="tab-panel"] {{
    background: transparent !important;
    padding: 14px 0 0 !important;
}}

/* ── Card ── */
.card {{
    background: {PANEL};
    border: 1px solid {BORDER};
    border-radius: 16px;
    padding: 20px;
}}

/* ── Labels ── */
.lbl {{
    font-size: 11px;
    font-weight: 500;
    letter-spacing: .08em;
    text-transform: uppercase;
    color: {MUTED};
    margin-bottom: 6px;
    display: block;
}}

/* ── Barra de completitud ── */
.prog-wrap {{
    display: flex; align-items: center; gap: 10px; margin-bottom: 14px;
}}
.prog-bg {{
    flex: 1; height: 3px; background: {BORDER}; border-radius: 2px; overflow: hidden;
}}
.prog-fill {{
    height: 100%; background: {ACCENT}; border-radius: 2px; transition: width .4s;
}}
.prog-meta {{
    font-size: 11px; color: {MUTED}; white-space: nowrap;
}}

/* ── Hist item ── */
.hist-row {{
    display: flex; align-items: center; gap: 8px;
    padding: 7px 12px; border-radius: 8px;
    background: {CARD}; border: 1px solid {BORDER};
    margin-bottom: 5px;
}}
.hist-dot {{ width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }}
.hist-text {{ font-size: 12px; color: {TEXT}; line-height: 1.35; }}
.hist-sub  {{ font-size: 11px; color: {MUTED}; }}

/* ── Defs box ── */
.defs-box {{
    background: {ED_BG};
    border: 1px solid {BORDER};
    border-radius: 12px;
    padding: 16px;
    font-size: 13px;
    line-height: 1.6;
    color: {MUTED};
    white-space: pre-wrap;
}}

/* ── Topbar ── */
.topbar {{
    display: flex; align-items: center; gap: 14px;
    padding-bottom: 20px; border-bottom: 1px solid {BORDER};
    margin-bottom: 24px;
}}
.logo {{
    font-size: 22px; font-weight: 600;
    color: {ACCENT}; letter-spacing: .1em;
    display: flex; align-items: center; gap: 8px;
}}
.logo-dot {{
    width: 8px; height: 8px; border-radius: 50%;
    background: {ACCENT};
    animation: p 2s ease-in-out infinite;
}}
@keyframes p {{ 0%,100%{{opacity:1}} 50%{{opacity:.25}} }}
.logo-sep {{ width: 1px; height: 18px; background: {BORDER}; }}
.logo-meta {{ font-size: 12px; color: {MUTED}; }}
.badge {{
    font-size: 11px; color: {ACCENT};
    background: {ACCENT}18;
    border: 1px solid {ACCENT}40;
    border-radius: 6px; padding: 3px 10px;
    font-weight: 500;
}}

::-webkit-scrollbar {{ width: 4px; }}
::-webkit-scrollbar-thumb {{ background: {BORDER}; border-radius: 2px; }}
hr {{ border: none; border-top: 1px solid {BORDER} !important; margin: 16px 0 !important; }}
</style>
""", unsafe_allow_html=True)

# ── Topbar ───────────────────────────────────────────────────
st.markdown(f"""
<div class="topbar">
  <div class="logo">
    <div class="logo-dot"></div>AURA
  </div>
  <div class="logo-sep"></div>
  <span class="logo-meta">Radiology Intelligence</span>
  <span class="badge">{st.session_state.modelo}</span>
</div>
""", unsafe_allow_html=True)

# ── Fila de ajustes ──────────────────────────────────────────
s1, s2, s3, s4 = st.columns([1.2, 1.2, 1.2, 0.8])

with s1:
    st.markdown('<span class="lbl">Modelo IA</span>', unsafe_allow_html=True)
    modelo = st.selectbox("m", list(MODELS.keys()),
                          index=list(MODELS.keys()).index(st.session_state.modelo),
                          label_visibility="collapsed")
    if modelo != st.session_state.modelo:
        st.session_state.modelo = modelo; st.rerun()

with s2:
    st.markdown('<span class="lbl">Plantilla .docx</span>', unsafe_allow_html=True)
    plantilla_f = st.file_uploader("p", type=["docx"], label_visibility="collapsed")
    if plantilla_f:
        st.session_state.plantilla_txt, _ = leer_plantilla(plantilla_f)

with s3:
    if not api_key:
        st.markdown('<span class="lbl">API Key</span>', unsafe_allow_html=True)
        api_key = st.text_input("k", type="password", label_visibility="collapsed",
                                placeholder="sk- ···")
    else:
        st.markdown('<span class="lbl">Estado</span>', unsafe_allow_html=True)
        st.markdown(f'<p style="font-size:13px;color:{ACCENT};padding-top:6px">● Conectado</p>',
                    unsafe_allow_html=True)

with s4:
    st.markdown('<span class="lbl">&nbsp;</span>', unsafe_allow_html=True)
    if st.button("＋ Nuevo informe", use_container_width=True):
        st.session_state.dictado = ""
        st.session_state.reporte = ""
        st.session_state.defs    = ""
        st.rerun()

st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

# ── Layout principal: izq + der ──────────────────────────────
col_l, col_r = st.columns([1, 1.4], gap="large")

# ═══════════════════════════════════════
# COLUMNA IZQUIERDA — Entrada
# ═══════════════════════════════════════
with col_l:

    st.markdown('<div class="card">', unsafe_allow_html=True)

    # Tabs de entrada
    tab_voz, tab_texto = st.tabs(["🎙  Voz", "⌨  Texto / Hallazgos"])

    with tab_voz:
        # Botón de micrófono visual
        st.markdown(f"""
        <div style="
            display:flex; flex-direction:column;
            align-items:center; gap:10px;
            padding: 24px 0 16px;
        ">
          <!-- Círculo de grabación -->
          <div style="
            position:relative;
            width:80px; height:80px;
          ">
            <!-- Onda exterior animada -->
            <div style="
              position:absolute; inset:-12px;
              border-radius:50%;
              border:1.5px solid {ACCENT}30;
              animation:wave1 2.2s ease-out infinite;
            "></div>
            <div style="
              position:absolute; inset:-6px;
              border-radius:50%;
              border:1.5px solid {ACCENT}50;
              animation:wave1 2.2s ease-out infinite .4s;
            "></div>
            <!-- Botón principal -->
            <div style="
              width:80px; height:80px;
              border-radius:50%;
              background:{CARD};
              border:2px solid {ACCENT};
              display:flex; align-items:center; justify-content:center;
            ">
              <svg width="28" height="28" viewBox="0 0 24 24" fill="none"
                   stroke="{ACCENT}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <rect x="9" y="2" width="6" height="12" rx="3"/>
                <path d="M5 10a7 7 0 0 0 14 0"/>
                <line x1="12" y1="19" x2="12" y2="22"/>
                <line x1="9"  y1="22" x2="15" y2="22"/>
              </svg>
            </div>
          </div>
          <span style="font-size:12px;color:{MUTED}">
            Presiona para grabar tu dictado
          </span>
        </div>
        <style>
        @keyframes wave1 {{
          0%   {{ transform:scale(1);   opacity:.7 }}
          100% {{ transform:scale(1.3); opacity:0  }}
        }}
        </style>
        """, unsafe_allow_html=True)

        audio = st.audio_input("grabación", label_visibility="collapsed")

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
                    st.info("Ingresa tu API Key para transcribir.")

    with tab_texto:
        st.markdown(f'<p style="font-size:12px;color:{MUTED};margin-bottom:8px">'
                    'Escribe hallazgos o diagnósticos directamente.</p>',
                    unsafe_allow_html=True)

    # Área de dictado compartida
    st.markdown('<span class="lbl" style="margin-top:12px">Señal de entrada</span>',
                unsafe_allow_html=True)
    dictado = st.text_area(
        "d",
        value=st.session_state.dictado,
        height=260,
        label_visibility="collapsed",
        placeholder=(
            "El dictado transcrito aparece aquí.\n"
            "También puedes escribir directamente.\n\n"
            "Ej: Desgarro horizontal menisco medial Stoller III,\n"
            "extrusión 3 mm, osteofitos marginales tibiofemorales."
        ),
        key="dictado_ta"
    )
    if dictado != st.session_state.dictado:
        st.session_state.dictado = dictado

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    # Acciones
    ba, bb = st.columns([1.6, 1])
    with ba:
        st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
        generar = st.button("Generar informe", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with bb:
        if st.button("Limpiar todo", use_container_width=True):
            st.session_state.dictado = ""
            st.session_state.audio_id = None
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)  # /card

    # Historial
    if st.session_state.historial:
        st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
        HCOLS = ["#3b9eff","#22c55e","#f59e0b","#ec4899","#8b5cf6","#06b6d4"]
        with st.expander(f"Historial  ·  {len(st.session_state.historial)} informes", expanded=False):
            for i, e in enumerate(st.session_state.historial):
                color = HCOLS[i % len(HCOLS)]
                st.markdown(f"""
                <div class="hist-row">
                  <div class="hist-dot" style="background:{color}"></div>
                  <div>
                    <div class="hist-text">{e['region']}</div>
                    <div class="hist-sub">{e['modalidad'][:20]}</div>
                  </div>
                </div>""", unsafe_allow_html=True)
                if st.button(f"Cargar #{i+1}", key=f"h{i}", use_container_width=True):
                    st.session_state.reporte = e['texto']
                    st.rerun()

# ── PROCESAMIENTO ──────────────────────────────────────────
if generar:
    if not api_key:
        st.warning("Ingresa tu API Key.")
    elif not st.session_state.dictado.strip():
        st.warning("Escribe o dicta hallazgos primero.")
    else:
        cl  = get_client()
        mid = MODELS[st.session_state.modelo]["id"]
        pt  = st.session_state.plantilla_txt
        instruc_tabla = (
            "La plantilla contiene tablas [TABLA]. Complétalas en Markdown con los valores del dictado."
            if "[TABLA" in pt else
            "NO generes tablas. No hay plantilla con tabla."
        )
        prompt = f"""Eres AURA, asistente de interpretación radiológica.
Analiza el dictado y detecta: modalidad, región, lateralidad, protocolo y clasificaciones relevantes.
Genera un informe radiológico estructurado y profesional.

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

                # Detectar modalidad y región del texto para historial
                lines = [l.strip() for l in report.split("\n") if l.strip()]
                mod = "RM"; reg = "General"
                for l in lines[:5]:
                    for m in ["Resonancia","Tomografía","Radiografía","Ultrasonido","PET"]:
                        if m.lower() in l.lower(): mod = m[:3].upper(); break
                    for r in ["Rodilla","Columna","Hombro","Cadera","Cerebro","Tórax","Abdomen"]:
                        if r.lower() in l.lower(): reg = r; break

                st.session_state.historial.insert(0, {"modalidad": mod, "region": reg, "texto": report})
                if len(st.session_state.historial) > 12:
                    st.session_state.historial = st.session_state.historial[:12]
                st.rerun()
            except Exception as e:
                st.error(str(e))

# ═══════════════════════════════════════
# COLUMNA DERECHA — Editor del informe
# ═══════════════════════════════════════
with col_r:

    # Barra de estado del informe
    rep = st.session_state.reporte
    if rep:
        secs  = sum(1 for s in ["TÉCNICA","HALLAZGOS","IMPRESIÓN"] if s in rep.upper())
        words = len(rep.split())
        pct   = min(100, int((secs/3)*60 + min(words/150,1)*40))
        st.markdown(f"""
        <div class="prog-wrap">
          <div class="prog-bg">
            <div class="prog-fill" style="width:{pct}%"></div>
          </div>
          <span class="prog-meta">{pct}% completo · {words} palabras</span>
        </div>""", unsafe_allow_html=True)
    else:
        st.markdown(f'<p style="font-size:13px;color:{MUTED};margin-bottom:14px">'
                    'El informe aparece aquí una vez generado.</p>',
                    unsafe_allow_html=True)

    # Editor de texto
    reporte = st.text_area(
        "informe",
        value=st.session_state.reporte,
        height=560,
        label_visibility="collapsed",
        placeholder=(
            "El informe generado aparece aquí.\n\n"
            "Puedes editarlo libremente.\n"
            "Fondo oscuro para no deslumbrar en la oscuridad."
        ),
        key="reporte_ta"
    )
    if reporte != st.session_state.reporte:
        st.session_state.reporte = reporte

    # Acciones del informe
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    a1, a2, a3 = st.columns(3)

    with a1:
        if st.button("Optimizar conclusión", use_container_width=True):
            if rep and api_key:
                cl  = get_client()
                mid = MODELS[st.session_state.modelo]["id"]
                with st.spinner("Optimizando..."):
                    try:
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":f"""Mejora ÚNICAMENTE la IMPRESIÓN DIAGNÓSTICA.
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
        if st.button("Definiciones", use_container_width=True):
            if rep and api_key:
                cl  = get_client()
                mid = MODELS[st.session_state.modelo]["id"]
                with st.spinner("Analizando..."):
                    try:
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":f"""Analiza el informe. Responde con este formato exacto.
Sin líneas en blanco entre ítems de la misma sección. Una línea entre secciones.

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
                        st.rerun()
                    except Exception as e: st.error(str(e))

    with a3:
        if rep:
            st.download_button(
                "Exportar .docx",
                data=generar_docx(st.session_state.reporte),
                file_name="AURA_Informe.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    # Panel de definiciones
    if st.session_state.defs:
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
        with st.expander("Definiciones · Clasificaciones · Referencias", expanded=True):
            st.markdown(f'<div class="defs-box">{st.session_state.defs}</div>',
                        unsafe_allow_html=True)
            if st.button("Cerrar"):
                st.session_state.defs = ""; st.rerun()
