import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile, io, os, re

st.set_page_config(page_title="AURA", layout="wide", initial_sidebar_state="collapsed")

# ── Paleta base ───────────────────────────────────────────────
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
    "Rosa": {
        "bg":"#1a0912","panel":"#220f1a","card":"#2a1220","border":"#3d1a2e",
        "accent":"#f472b6","text":"#fce7f3","muted":"#9d4f75","ed_bg":"#1e0d17","green":"#34d399",
    },
    "Océano": {
        "bg":"#040f1a","panel":"#061624","card":"#081e30","border":"#0e2d45",
        "accent":"#22d3ee","text":"#cff5fc","muted":"#276a80","ed_bg":"#050f1a","green":"#34d399",
    },
    "Claro": {
        "bg":"#f0f4f8","panel":"#ffffff","card":"#f8fafc","border":"#e2e8f0",
        "accent":"#2563eb","text":"#1e293b","muted":"#64748b","ed_bg":"#ffffff","green":"#16a34a",
    },
}

MODELS = {
    "DeepSeek Chat": {"url": "https://api.deepseek.com", "id": "deepseek-chat"},
    "GPT-4o Mini":   {"url": None, "id": "gpt-4o-mini"},
    "GPT-4.1 Mini":  {"url": None, "id": "gpt-4.1-mini"},
}

MODALIDADES = [
    "Resonancia Magnética","Tomografía Computarizada","Radiografía",
    "Ultrasonido","PET-CT","Mamografía","Fluoroscopía","Angiografía",
]

REGIONES = {
    "Extremidades inferiores": ["Rodilla","Cadera","Tobillo","Pie","Muslo","Pierna"],
    "Extremidades superiores": ["Hombro","Codo","Muñeca","Mano","Brazo","Antebrazo"],
    "Columna": ["Col. cervical","Col. dorsal","Col. lumbar","Sacro / Cóccix"],
    "Cráneo y cuello": ["Cerebro","Silla turca","Órbitas","Oídos / Peñascos","Cuello","Tiroides","Glándulas salivales"],
    "Tórax": ["Tórax","Pulmón","Corazón","Mediastino","Mama"],
    "Abdomen y pelvis": ["Abdomen","Pelvis","Hígado","Páncreas","Riñones","Suprarrenales","Bazo","Vejiga","Próstata","Útero / Anexos"],
}

HCOLS = ["#3b9eff","#22c55e","#f59e0b","#ec4899","#8b5cf6","#06b6d4"]

DEFAULTS = {
    "dictado":"","reporte":"","defs":"",
    "modelo":"DeepSeek Chat","audio_id":None,
    "historial":[],"plantilla_txt":"",
    "panel_izq":True,"panel_der":True,
    "tema":"AURA Dark",
}
for k,v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k]=v

try:    api_key=st.secrets["deepseek_key"]
except: api_key=os.environ.get("OPENAI_API_KEY","")

# ── Helpers ───────────────────────────────────────────────────
def get_client():
    cfg=MODELS[st.session_state.modelo]
    return OpenAI(api_key=api_key,base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)

def get_theme():
    return THEMES[st.session_state.tema]

def leer_plantilla(f):
    doc=Document(f); partes=[]; n=0
    try:
        import docx.text.paragraph as pp, docx.table as tt
        for el in doc.element.body:
            tag=el.tag.split('}')[-1]
            if tag=='p':
                p=pp.Paragraph(el,doc); t=p.text.strip()
                if t: partes.append(t)
            elif tag=='tbl':
                n+=1; tbl=tt.Table(el,doc)
                rows=["| "+" | ".join(c.text.strip() for c in r.cells)+" |" for r in tbl.rows]
                partes.append(f"[TABLA {n}]\n"+"\n".join(rows)+"\n[/TABLA]")
    except:
        partes=[p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n".join(partes), n>0

def generar_docx(texto):
    doc=Document()
    doc.styles["Normal"].font.name="Calibri"
    doc.styles["Normal"].font.size=Pt(11)
    for line in texto.split("\n"):
        s=line.strip()
        if not s: doc.add_paragraph(); continue
        if s.isupper() and len(s)<80:
            h=doc.add_heading(s,level=1); h.alignment=WD_ALIGN_PARAGRAPH.LEFT
        elif s.startswith(("•","·")): doc.add_paragraph(s[1:].strip(),style="List Bullet")
        else: doc.add_paragraph(s)
    bio=io.BytesIO(); doc.save(bio); return bio.getvalue()

def transcribir(audio):
    cfg=MODELS[st.session_state.modelo]
    cl=OpenAI(api_key=api_key,base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)
    with tempfile.NamedTemporaryFile(delete=False,suffix=".wav") as tmp:
        tmp.write(audio.read()); path=tmp.name
    with open(path,"rb") as f:
        res=cl.audio.transcriptions.create(
            model="whisper-1",file=f,language="es",
            prompt="Dictado radiológico: Stoller, ICRS, LCA, menisco, condromalacia, osteofito, Kellgren-Lawrence."
        )
    os.unlink(path); return res.text.strip()

def completitud(texto):
    secs=sum(1 for s in ["TÉCNICA","HALLAZGOS","IMPRESIÓN"] if s in texto.upper())
    words=len(texto.split())
    return min(100,int((secs/3)*60+min(words/150,1)*40)), words

def get_region():
    custom=st.session_state.get("reg_custom","").strip()
    if custom: return custom
    grupo=st.session_state.get("sel_grupo","Extremidades inferiores")
    reg=st.session_state.get("sel_reg","Rodilla")
    return reg

# ── CSS dinámico ──────────────────────────────────────────────
def render_css():
    T=get_theme()
    BG=T["bg"]; PANEL=T["panel"]; CARD=T["card"]; BORDER=T["border"]
    ACCENT=T["accent"]; TEXT=T["text"]; MUTED=T["muted"]; ED_BG=T["ed_bg"]; GREEN=T["green"]
    st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=JetBrains+Mono:wght@400;500&family=Outfit:wght@300;400;500;600;700&display=swap');

html,body,.stApp{{background:{BG};color:{TEXT};font-family:'Outfit',sans-serif}}
header,footer,#MainMenu{{visibility:hidden}}
.block-container{{padding:0!important;max-width:100%!important}}
*{{box-sizing:border-box}}

/* TOPBAR */
.topbar{{height:52px;background:{PANEL};border-bottom:1px solid {BORDER};
  display:flex;align-items:center;padding:0 24px;gap:16px;
  position:sticky;top:0;z-index:100}}
.logo{{font-size:18px;font-weight:700;color:{ACCENT};letter-spacing:.15em;
  display:flex;align-items:center;gap:9px;font-family:'DM Serif Display',serif}}
.logo-dot{{width:8px;height:8px;border-radius:50%;background:{ACCENT};
  animation:dp 2s ease-in-out infinite}}
@keyframes dp{{0%,100%{{opacity:1}}50%{{opacity:.2}}}}
.t-sep{{width:1px;height:18px;background:{BORDER}}}
.t-meta{{font-size:12px;color:{MUTED}}}
.t-badge{{font-size:11px;color:{ACCENT};background:{ACCENT}18;
  border:1px solid {ACCENT}40;border-radius:6px;padding:3px 10px}}
.t-right{{margin-left:auto;display:flex;align-items:center;gap:10px}}
.t-dot-on{{width:6px;height:6px;border-radius:50%;background:{GREEN};box-shadow:0 0 5px {GREEN}}}

/* SELECTS & INPUTS */
[data-testid="stSelectbox"]>div>div{{
  background:{CARD}!important;border:1px solid {BORDER}!important;
  border-radius:8px!important;color:{TEXT}!important;font-size:13px!important}}
[data-testid="stSelectbox"]>div>div:hover{{border-color:{ACCENT}50!important}}
.stTextInput input{{
  background:{CARD}!important;border:1px solid {BORDER}!important;
  border-radius:8px!important;color:{TEXT}!important;font-size:13px!important}}
.stTextInput input:focus{{border-color:{ACCENT}50!important}}

/* EDITOR PRINCIPAL */
.stTextArea textarea{{
  background:{ED_BG}!important;border:1px solid {BORDER}!important;
  border-radius:10px!important;color:{TEXT}!important;
  font-size:14px!important;line-height:1.8!important;padding:20px!important;
  caret-color:{ACCENT}!important;font-family:'Outfit',sans-serif!important}}
.stTextArea textarea:focus{{border-color:{ACCENT}50!important;box-shadow:0 0 0 3px {ACCENT}12!important}}
.stTextArea textarea::placeholder{{color:{MUTED}!important}}

/* TOOLBAR */
.toolbar-wrap{{
  background:{PANEL};border:1px solid {BORDER};border-radius:10px;
  padding:10px 14px;margin-bottom:10px;display:flex;flex-wrap:wrap;align-items:center;gap:6px}}
.tool-btn{{
  background:{CARD};border:1px solid {BORDER};border-radius:6px;
  color:{MUTED};padding:5px 10px;font-size:12px;font-weight:600;cursor:pointer;
  font-family:'Outfit',sans-serif;transition:all .15s}}
.tool-btn:hover{{border-color:{ACCENT}60;color:{TEXT}}}
.tool-select{{
  background:{CARD};border:1px solid {BORDER};border-radius:6px;
  color:{TEXT};font-size:12px;padding:5px 8px;font-family:'Outfit',sans-serif}}
.tool-sep{{width:1px;height:20px;background:{BORDER};margin:0 2px}}

/* BOTONES */
.stButton button{{
  background:{CARD}!important;border:1px solid {BORDER}!important;
  color:{TEXT}!important;border-radius:8px!important;
  font-size:13px!important;font-weight:500!important;transition:all .15s!important}}
.stButton button:hover{{border-color:{ACCENT}60!important;background:{PANEL}!important}}
.btn-primary .stButton button{{
  background:{ACCENT}!important;border-color:{ACCENT}!important;
  color:#fff!important;font-weight:600!important}}
.btn-primary .stButton button:hover{{opacity:.88!important}}
.stDownloadButton button{{
  background:transparent!important;border:1px solid {ACCENT}!important;
  color:{ACCENT}!important;border-radius:8px!important;font-size:13px!important}}
.stDownloadButton button:hover{{background:{ACCENT}18!important}}

/* SECCIONES */
.section{{background:{CARD};border:1px solid {BORDER};border-radius:12px;padding:16px;margin-bottom:12px}}
.section-title{{font-size:11px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;
  color:{MUTED};margin-bottom:12px;display:flex;align-items:center;gap:6px}}
.section-title .dot{{width:6px;height:6px;border-radius:50%;background:{ACCENT}}}

/* PROGRESS */
.prog-row{{display:flex;align-items:center;gap:10px;margin-bottom:10px}}
.prog-bg{{flex:1;height:3px;background:{BORDER};border-radius:2px;overflow:hidden}}
.prog-fill{{height:100%;background:{ACCENT};border-radius:2px;transition:width .4s}}
.prog-txt{{font-size:11px;color:{MUTED};white-space:nowrap}}

/* HISTORIAL */
.h-row{{display:flex;align-items:center;gap:8px;padding:7px 10px;border-radius:8px;
  background:{PANEL};border:1px solid {BORDER};margin-bottom:4px}}
.h-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0}}
.h-name{{font-size:12px;color:{TEXT}}}
.h-sub{{font-size:11px;color:{MUTED}}}

/* DEFS */
.defs-box{{background:{ED_BG};border:1px solid {BORDER};border-radius:10px;
  padding:14px;font-size:12.5px;line-height:1.6;color:{MUTED};white-space:pre-wrap}}

/* TEMAS */
.tema-btn{{
  background:{CARD};border:2px solid {BORDER};border-radius:8px;
  padding:8px;cursor:pointer;text-align:center;font-size:11px;
  color:{MUTED};transition:all .15s;margin-bottom:6px}}
.tema-btn:hover{{border-color:{ACCENT};color:{TEXT}}}
.tema-btn.active{{border-color:{ACCENT};color:{ACCENT}}}
.tema-swatch{{width:100%;height:14px;border-radius:4px;margin-bottom:4px}}

/* AUDIO */
[data-testid="stAudioInput"]{{background:{CARD}!important;border:1px solid {BORDER}!important;border-radius:10px!important}}
[data-testid="stFileUploader"]{{background:{CARD};border:1px dashed {BORDER};border-radius:10px;padding:6px}}
[data-testid="stFileUploader"] *{{color:{MUTED}!important;font-size:12px!important}}

/* EXPANDER */
[data-testid="stExpander"]{{background:{CARD}!important;border:1px solid {BORDER}!important;border-radius:10px!important;margin-bottom:8px!important}}
[data-testid="stExpander"] summary{{color:{MUTED}!important;font-size:13px!important}}
[data-testid="stExpander"] summary:hover{{color:{TEXT}!important}}

/* TABS */
[data-testid="stTabs"] [role="tablist"]{{border-bottom:1px solid {BORDER}!important;background:transparent!important}}
[data-testid="stTabs"] [role="tab"]{{background:transparent!important;border:none!important;
  color:{MUTED}!important;font-size:13px!important;padding:8px 14px!important;
  border-bottom:2px solid transparent!important}}
[data-testid="stTabs"] [role="tab"][aria-selected="true"]{{color:{ACCENT}!important;border-bottom-color:{ACCENT}!important}}

::-webkit-scrollbar{{width:3px}}
::-webkit-scrollbar-thumb{{background:{BORDER};border-radius:2px}}
hr{{border:none;border-top:1px solid {BORDER}!important;margin:12px 0!important}}
[data-testid="column"]{{padding:0!important}}
</style>
""", unsafe_allow_html=True)

render_css()

T=get_theme()
ACCENT=T["accent"]; TEXT=T["text"]; MUTED=T["muted"]; BORDER=T["border"]
PANEL=T["panel"]; CARD=T["card"]; GREEN=T["green"]; ED_BG=T["ed_bg"]

# ── TOPBAR ────────────────────────────────────────────────────
ok=bool(api_key)
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

# ── LAYOUT COLUMNAS ───────────────────────────────────────────
left_open=st.session_state.panel_izq
right_open=st.session_state.panel_der

if left_open and right_open:    ratios=[1.1,2.5,0.95]
elif left_open and not right_open: ratios=[1.1,3.3,0.07]
elif not left_open and right_open: ratios=[0.07,3.3,0.95]
else:                              ratios=[0.07,5.0,0.07]

col_l,col_c,col_r=st.columns(ratios,gap="small")

# ═══════════════════════════════════════════════════════════════
# PANEL IZQUIERDO
# ═══════════════════════════════════════════════════════════════
with col_l:
    tl=st.columns([1,0.01])[0]
    toggle_lbl="◀" if left_open else "▶"
    if st.button(toggle_lbl,key="tog_l",help="Colapsar / expandir panel"):
        st.session_state.panel_izq=not left_open; st.rerun()

    if not left_open:
        st.markdown(f"""<div style="display:flex;flex-direction:column;align-items:center;gap:18px;padding:16px 0">
          <span style="font-size:16px;color:{MUTED}">🎙</span>
          <span style="font-size:16px;color:{MUTED}">📋</span>
          <span style="font-size:16px;color:{MUTED}">⚙</span>
        </div>""",unsafe_allow_html=True)
        generar=False
    else:
        # ── Estudio ──
        st.markdown(f'<div class="section-title" style="margin-top:8px"><div class="dot"></div>Estudio</div>',unsafe_allow_html=True)

        st.markdown(f'<p style="font-size:10px;color:{MUTED};text-transform:uppercase;letter-spacing:.08em;margin-bottom:4px">Modalidad</p>',unsafe_allow_html=True)
        st.selectbox("mod",MODALIDADES,label_visibility="collapsed",key="sel_mod")

        st.markdown(f'<p style="font-size:10px;color:{MUTED};text-transform:uppercase;letter-spacing:.08em;margin-bottom:4px;margin-top:8px">Grupo anatómico</p>',unsafe_allow_html=True)
        grupo=st.selectbox("grp",list(REGIONES.keys()),label_visibility="collapsed",key="sel_grupo")

        st.markdown(f'<p style="font-size:10px;color:{MUTED};text-transform:uppercase;letter-spacing:.08em;margin-bottom:4px;margin-top:8px">Región</p>',unsafe_allow_html=True)
        st.selectbox("reg",REGIONES[grupo],label_visibility="collapsed",key="sel_reg")

        st.markdown(f'<p style="font-size:10px;color:{MUTED};text-transform:uppercase;letter-spacing:.08em;margin-bottom:4px;margin-top:8px">Región libre (opcional)</p>',unsafe_allow_html=True)
        st.text_input("rc",label_visibility="collapsed",key="reg_custom",placeholder="Ej: Articulación glenohumeral derecha")

        st.markdown("<hr>",unsafe_allow_html=True)

        # ── Dictado ──
        st.markdown(f'<div class="section-title"><div class="dot"></div>Dictado</div>',unsafe_allow_html=True)

        tab_voz,tab_texto=st.tabs(["🎙 Voz","⌨ Texto"])
        with tab_voz:
            st.markdown(f"""<div style="display:flex;flex-direction:column;align-items:center;padding:18px 0 12px;gap:8px">
              <div style="position:relative;width:72px;height:72px">
                <div style="position:absolute;inset:-12px;border-radius:50%;border:1.5px solid {ACCENT}28;animation:rp 2.4s ease-out infinite"></div>
                <div style="position:absolute;inset:-6px;border-radius:50%;border:1.5px solid {ACCENT}40;animation:rp 2.4s ease-out infinite .5s"></div>
                <div style="width:72px;height:72px;border-radius:50%;background:{CARD};border:2px solid {ACCENT};
                  display:flex;align-items:center;justify-content:center">
                  <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="{ACCENT}" stroke-width="1.8"
                    stroke-linecap="round" stroke-linejoin="round">
                    <rect x="9" y="2" width="6" height="12" rx="3"/>
                    <path d="M5 10a7 7 0 0 0 14 0"/><line x1="12" y1="19" x2="12" y2="22"/>
                    <line x1="9" y1="22" x2="15" y2="22"/>
                  </svg>
                </div>
              </div>
              <span style="font-size:11px;color:{MUTED}">Pulsa para grabar</span>
            </div>
            <style>@keyframes rp{{0%{{transform:scale(1);opacity:.6}}100%{{transform:scale(1.35);opacity:0}}}}</style>""",
            unsafe_allow_html=True)
            audio=st.audio_input("rec",label_visibility="collapsed")
            if audio:
                aid=hash(audio.read()); audio.seek(0)
                if aid!=st.session_state.audio_id:
                    if api_key:
                        with st.spinner("Transcribiendo..."):
                            txt=transcribir(audio)
                        if txt:
                            st.session_state.dictado+=(" "+txt).strip()
                            st.session_state.audio_id=aid
                            st.rerun()
                    else:
                        st.info("Configura tu API Key.")

        with tab_texto:
            st.markdown(f'<p style="font-size:11px;color:{MUTED};margin-bottom:4px">Escribe hallazgos directamente.</p>',unsafe_allow_html=True)

        st.markdown(f'<p style="font-size:10px;color:{MUTED};text-transform:uppercase;letter-spacing:.08em;margin:8px 0 4px">Señal de entrada</p>',unsafe_allow_html=True)
        dictado=st.text_area("d",value=st.session_state.dictado,height=180,
            label_visibility="collapsed",
            placeholder="El dictado transcrito aparece aquí.\nTambién puedes escribir directamente.\n\nEj: Desgarro horizontal menisco medial Stoller III, extrusión 3mm, osteofitos marginales.",
            key="dictado_ta")
        if dictado!=st.session_state.dictado:
            st.session_state.dictado=dictado

        st.markdown("<div style='height:6px'></div>",unsafe_allow_html=True)
        ba,bb=st.columns([1.6,1])
        with ba:
            st.markdown('<div class="btn-primary">',unsafe_allow_html=True)
            generar=st.button("✦ Generar informe",use_container_width=True)
            st.markdown('</div>',unsafe_allow_html=True)
        with bb:
            if st.button("Limpiar",use_container_width=True):
                st.session_state.dictado=""; st.session_state.audio_id=None; st.rerun()

        st.markdown("<hr>",unsafe_allow_html=True)

        # ── Temas visuales ──
        with st.expander("🎨  Tema visual",expanded=False):
            st.markdown(f'<p style="font-size:11px;color:{MUTED};margin-bottom:8px">Selecciona el tema de color de la interfaz.</p>',unsafe_allow_html=True)
            cols_t=st.columns(3)
            temas_list=list(THEMES.items())
            for i,(nombre,t) in enumerate(temas_list):
                with cols_t[i%3]:
                    active="✓ " if nombre==st.session_state.tema else ""
                    if st.button(f"{active}{nombre}",key=f"tema_{i}",use_container_width=True):
                        st.session_state.tema=nombre; st.rerun()

        # ── Configuración avanzada ──
        with st.expander("⚙  Configuración avanzada",expanded=False):
            st.markdown(f'<p style="font-size:10px;color:{MUTED};text-transform:uppercase;letter-spacing:.08em;margin-bottom:4px">Modelo IA</p>',unsafe_allow_html=True)
            m=st.selectbox("m",list(MODELS.keys()),
                index=list(MODELS.keys()).index(st.session_state.modelo),
                label_visibility="collapsed")
            if m!=st.session_state.modelo:
                st.session_state.modelo=m; st.rerun()

            st.markdown(f'<p style="font-size:10px;color:{MUTED};text-transform:uppercase;letter-spacing:.08em;margin:8px 0 4px">Plantilla .DOCX</p>',unsafe_allow_html=True)
            f_up=st.file_uploader("plt",type=["docx"],label_visibility="collapsed")
            if f_up:
                st.session_state.plantilla_txt,_=leer_plantilla(f_up)
                st.markdown(f'<p style="font-size:11px;color:{ACCENT}">✓ Plantilla cargada</p>',unsafe_allow_html=True)

            if not api_key:
                st.markdown(f'<p style="font-size:10px;color:{MUTED};text-transform:uppercase;letter-spacing:.08em;margin:8px 0 4px">API Key</p>',unsafe_allow_html=True)
                api_key=st.text_input("k",type="password",label_visibility="collapsed",placeholder="sk- ···")

        # ── Historial ──
        if st.session_state.historial:
            with st.expander(f"📋  Historial  ·  {len(st.session_state.historial)}",expanded=False):
                for i,e in enumerate(st.session_state.historial):
                    color=HCOLS[i%len(HCOLS)]
                    st.markdown(f"""<div class="h-row">
                      <div class="h-dot" style="background:{color}"></div>
                      <div><div class="h-name">{e['region']}</div><div class="h-sub">{e['modalidad'][:18]}</div></div>
                    </div>""",unsafe_allow_html=True)
                    if st.button(f"Cargar",key=f"h{i}",use_container_width=True):
                        st.session_state.reporte=e['texto']; st.rerun()

# ── PROCESAMIENTO ─────────────────────────────────────────────
if generar:
    if not api_key:
        st.warning("Ingresa tu API Key en Configuración avanzada.")
    elif not st.session_state.dictado.strip():
        st.warning("Escribe o dicta hallazgos primero.")
    else:
        cl=get_client()
        mid=MODELS[st.session_state.modelo]["id"]
        pt=st.session_state.plantilla_txt
        instruc_tabla=("La plantilla tiene tablas [TABLA]. Complétalas en Markdown."
            if "[TABLA" in pt else "NO generes tablas bajo ninguna circunstancia.")
        mod_sel=st.session_state.get("sel_mod","")
        reg_sel=st.session_state.get("reg_custom","").strip() or st.session_state.get("sel_reg","")

        prompt=f"""Eres AURA, sistema experto de interpretación radiológica de nivel subespecialista.
Tu redacción es la de un radiólogo con fellowship en imagen musculoesquelética, neuroradiología o body imaging,
con más de 15 años de práctica en centros de referencia de alto volumen.

MODALIDAD: {mod_sel}
REGIÓN: {reg_sel}

════════════════════════════════════════
ESTÁNDARES DE REDACCIÓN OBLIGATORIOS
════════════════════════════════════════

PRECISIÓN MORFOLÓGICA:
· Cada hallazgo debe incluir: localización exacta, extensión (mm o %), morfología, señal/densidad y repercusión estructural.
· PROHIBIDO usar "cambios degenerativos" sin especificar: tipo, grado, distribución y correlato estructural.
· PROHIBIDO: "alteración inespecífica", "podría corresponder". Usa afirmaciones diagnósticas directas o diagnósticos diferenciales jerarquizados.
· Cuantifica siempre: dimensiones, porcentajes, grados, scores validados.

CLASIFICACIONES OBLIGATORIAS SEGÚN HALLAZGO:
· Menisco: Stoller (I-III), localización, tipo morfológico (horizontal/radial/complejo/raíz)
· Cartílago: ICRS, Outerbridge, MOAKS
· LCA/LCP: continuidad de fibras, edema óseo, ángulo de Blumensaat
· Columna: Pfirrmann, Modic, Meyerding, NASCET
· Hombro: Bigliani, Goutallier, Sugaya
· Cadera: Tönnis, ángulo alpha (FAI)
· Otros órganos: TIRADS, BI-RADS, PI-RADS, ASPECTS, Fazekas según corresponda
· Siempre: grado + significado clínico del grado.

ESTRUCTURA (títulos en MAYÚSCULAS, sin markdown, sin asteriscos):

INDICACIÓN
[Motivo clínico explícito o inferido. Si no se proporciona, redacta uno coherente con los hallazgos.]

TÉCNICA
[Descripción técnica estándar: secuencias/protocolos, planos, contraste, campo magnético. Nivel de detalle publicable.]

HALLAZGOS
[Sistemático por compartimentos o estructuras anatómicas. Cada estructura mencionada: normal o con hallazgo detallado. Nunca omitir estructuras principales sin aclarar su estado.]

IMPRESIÓN DIAGNÓSTICA
[Síntesis jerarquizada. Hallazgo principal primero. Usa • por punto.
Cada viñeta: diagnóstico + clasificación/grado + implicación clínica.
Última viñeta: correlación clínico-radiológica y orientación de manejo.]

ESTILO: Voz activa, tiempo presente, oraciones cortas y precisas. Nivel: publicable en revista indexada o auditable por comité de pares.
{instruc_tabla}

PLANTILLA:
{pt if pt else "INDICACIÓN\nTÉCNICA\nHALLAZGOS\nIMPRESIÓN DIAGNÓSTICA"}

DICTADO DEL RADIÓLOGO:
{st.session_state.dictado}"""
        with st.spinner("Generando informe..."):
            try:
                res=cl.chat.completions.create(
                    model=mid,
                    messages=[{"role":"system","content":prompt}],
                    temperature=0.1,max_tokens=2500
                )
                report=res.choices[0].message.content
                st.session_state.reporte=report
                st.session_state.historial.insert(0,{
                    "modalidad":mod_sel[:18] if mod_sel else "RM",
                    "region":reg_sel if reg_sel else "General",
                    "texto":report
                })
                if len(st.session_state.historial)>12:
                    st.session_state.historial=st.session_state.historial[:12]
                st.rerun()
            except Exception as e:
                st.error(str(e))

# ═══════════════════════════════════════════════════════════════
# PANEL CENTRAL — Editor rico
# ═══════════════════════════════════════════════════════════════
with col_c:
    st.markdown("<div style='height:10px'></div>",unsafe_allow_html=True)
    rep=st.session_state.reporte

    # Progress
    if rep:
        pct,words=completitud(rep)
        st.markdown(f"""<div class="prog-row">
          <div class="prog-bg"><div class="prog-fill" style="width:{pct}%"></div></div>
          <span class="prog-txt">{pct}% · {words} palabras</span>
        </div>""",unsafe_allow_html=True)

    # ── Toolbar de formato (HTML + JS) ──
    st.markdown(f"""
<div class="toolbar-wrap" id="toolbar">
  <select class="tool-select" style="width:110px" onchange="execFormatCmd('fontName',this.value)">
    <option value="Outfit">Outfit</option>
    <option value="'DM Serif Display'">DM Serif</option>
    <option value="'JetBrains Mono'">Mono</option>
    <option value="Georgia">Georgia</option>
    <option value="Arial">Arial</option>
    <option value="'Times New Roman'">Times New Roman</option>
  </select>
  <select class="tool-select" style="width:54px" onchange="execFormatCmd('fontSize',this.value)">
    <option value="1">10</option><option value="2">12</option>
    <option value="3" selected>14</option><option value="4">16</option>
    <option value="5">18</option><option value="6">24</option><option value="7">36</option>
  </select>
  <select class="tool-select" style="width:84px" onchange="applyHeading(this.value)">
    <option value="p">Párrafo</option>
    <option value="h1">Título 1</option><option value="h2">Título 2</option><option value="h3">Título 3</option>
  </select>
  <div class="tool-sep"></div>
  <button class="tool-btn" title="Negrita" onclick="execFormatCmd('bold')"><b>B</b></button>
  <button class="tool-btn" title="Cursiva" onclick="execFormatCmd('italic')"><i>I</i></button>
  <button class="tool-btn" title="Subrayado" onclick="execFormatCmd('underline')"><u style="text-decoration:underline">U</u></button>
  <button class="tool-btn" title="Tachado" onclick="execFormatCmd('strikeThrough')" style="text-decoration:line-through">S</button>
  <div class="tool-sep"></div>
  <button class="tool-btn" title="Izquierda" onclick="execFormatCmd('justifyLeft')">⬛</button>
  <button class="tool-btn" title="Centro" onclick="execFormatCmd('justifyCenter')">▬</button>
  <button class="tool-btn" title="Derecha" onclick="execFormatCmd('justifyRight')">⬛</button>
  <button class="tool-btn" title="Justificado" onclick="execFormatCmd('justifyFull')">≡</button>
  <div class="tool-sep"></div>
  <button class="tool-btn" title="Lista con viñetas" onclick="execFormatCmd('insertUnorderedList')">• —</button>
  <button class="tool-btn" title="Lista numerada" onclick="execFormatCmd('insertOrderedList')">1.</button>
  <div class="tool-sep"></div>
  <label title="Color de texto" style="display:flex;align-items:center;gap:3px;font-size:11px;color:{MUTED}">
    A <input type="color" value="#dce8f4" style="width:22px;height:22px;border:none;border-radius:4px;cursor:pointer;padding:0"
      onchange="execFormatCmd('foreColor',this.value)">
  </label>
  <label title="Resaltar" style="display:flex;align-items:center;gap:3px;font-size:11px;color:{MUTED}">
    HL <input type="color" value="#3b9eff" style="width:22px;height:22px;border:none;border-radius:4px;cursor:pointer;padding:0"
      onchange="execFormatCmd('hiliteColor',this.value)">
  </label>
  <div class="tool-sep"></div>
  <button class="tool-btn" onclick="insertTable()" style="padding:5px 10px;font-size:11px">+ Tabla</button>
  <button class="tool-btn" onclick="execFormatCmd('insertHorizontalRule')" style="padding:5px 8px;font-size:11px">— HR</button>
  <div class="tool-sep"></div>
  <button class="tool-btn" onclick="execFormatCmd('undo')" title="Deshacer">↩</button>
  <button class="tool-btn" onclick="execFormatCmd('redo')" title="Rehacer">↪</button>
</div>

<!-- EDITOR RICO -->
<div style="position:relative">
<div style="background:{ED_BG};border:1px solid {BORDER};border-radius:10px;
  min-height:480px;height:540px;padding:32px 40px;font-size:14px;line-height:1.8;
  color:{TEXT};font-family:\'Outfit\',sans-serif;outline:none;overflow-y:auto;
  resize:vertical"
  id="richEditor" contenteditable="true"
  data-placeholder="Genera un informe para comenzar a editar, o escribe directamente aquí."
  oninput="updateToolbar()" onclick="updateToolbar()" onkeyup="updateToolbar()">
  {rep.replace(chr(10),"<br>") if rep else ""}
</div>
</div>

<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=JetBrains+Mono&family=Outfit:wght@300;400;500;600;700&display=swap');
#richEditor:empty::before{{content:attr(data-placeholder);color:{MUTED};pointer-events:none}}
#richEditor::-webkit-resizer{{
  background:{CARD};border-radius:0 0 10px 0;
}}
#richEditor{{
  resize:vertical;
  overflow:auto;
}}
#richEditor h1{{font-size:22px;font-weight:700;margin:16px 0 8px;color:{ACCENT};font-family:'DM Serif Display',serif}}
#richEditor h2{{font-size:17px;font-weight:600;margin:14px 0 6px;text-transform:uppercase;letter-spacing:.06em;color:{TEXT}}}
#richEditor h3{{font-size:14px;font-weight:600;margin:10px 0 4px;color:{TEXT}}}
#richEditor table{{border-collapse:collapse;width:100%;margin:12px 0;font-size:13px}}
#richEditor table td,#richEditor table th{{border:1px solid {BORDER};padding:8px 12px;text-align:left}}
#richEditor table th{{background:color-mix(in srgb,{ACCENT} 15%,transparent);color:{ACCENT};font-weight:600}}
#richEditor ul{{margin:4px 0 4px 20px}} #richEditor ol{{margin:4px 0 4px 20px}}
</style>

<script>
function execFormatCmd(cmd,val){{
  document.getElementById('richEditor').focus();
  document.execCommand(cmd,false,val||null);
  updateToolbar();
  syncEditor();
}}
function applyHeading(tag){{
  document.getElementById('richEditor').focus();
  document.execCommand('formatBlock',false,tag);
  syncEditor();
}}
function updateToolbar(){{
  ['bold','italic','underline'].forEach(cmd=>{{
    document.querySelectorAll('.tool-btn').forEach(b=>{{
      const t=b.title.toLowerCase();
      if(t===cmd) b.style.background=document.queryCommandState(cmd)?'rgba(59,158,255,.2)':'';
    }});
  }});
}}
function insertTable(){{
  const rows=parseInt(prompt('Número de filas:','3'))||3;
  const cols=parseInt(prompt('Número de columnas:','3'))||3;
  let html='<table><thead><tr>';
  for(let c=0;c<cols;c++) html+=`<th>Columna ${{c+1}}</th>`;
  html+='</tr></thead><tbody>';
  for(let r=0;r<rows-1;r++){{
    html+='<tr>';
    for(let c=0;c<cols;c++) html+='<td>&nbsp;</td>';
    html+='</tr>';
  }}
  html+='</tbody></table><p><br></p>';
  document.getElementById('richEditor').focus();
  document.execCommand('insertHTML',false,html);
  syncEditor();
}}
function syncEditor(){{
  // Sync con Streamlit via hidden input
  const content=document.getElementById('richEditor').innerHTML;
  const ta=document.querySelector('textarea[aria-label="reporte_sync"]');
  if(ta){{ta.value=content;ta.dispatchEvent(new Event('input',{{bubbles:true}}));}}
}}
// Restaurar contenido al cargar
window.addEventListener('load',()=>{{
  const ed=document.getElementById('richEditor');
  if(!ed.innerHTML.trim()||ed.innerHTML==='<br>') {{
    const stored=window._auraContent||'';
    if(stored) ed.innerHTML=stored;
  }}
  updateToolbar();
}});
document.getElementById('richEditor').addEventListener('keydown',e=>{{
  if(e.key==='Tab'){{e.preventDefault();execFormatCmd('insertHTML','&nbsp;&nbsp;&nbsp;&nbsp;');}}
}});
</script>
""",unsafe_allow_html=True)

    st.markdown("<div style='height:8px'></div>",unsafe_allow_html=True)
    a1,a2,a3=st.columns(3)

    with a1:
        if st.button("✦ Optimizar conclusión",use_container_width=True):
            rep2=st.session_state.reporte
            if rep2 and api_key:
                cl=get_client(); mid=MODELS[st.session_state.modelo]["id"]
                with st.spinner("Optimizando..."):
                    try:
                        r=cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":
                                f"""Eres un radiólogo subespecialista senior revisando la IMPRESIÓN DIAGNÓSTICA de un colega.
Tu objetivo: elevarla al nivel de un informe de centro de referencia internacional.

CRITERIOS DE EXCELENCIA para la IMPRESIÓN DIAGNÓSTICA:
1. Jerarquía diagnóstica: hallazgo principal primero, secundarios después, incidentales al final.
2. Cada viñeta (•) debe contener: [Estructura] + [Diagnóstico específico] + [Clasificación/grado validado] + [Implicación clínica o sugerencia de manejo].
3. La última viñeta debe integrar correlación clínico-radiológica y orientar al clínico (ej: "Se sugiere valoración por ortopedia/neurología, considerar artroscopia/infiltración/seguimiento en X meses").
4. Lenguaje afirmativo. Evita hedge words ("podría", "posible") salvo diagnóstico diferencial genuino.
5. Si hay diagnósticos diferenciales, listarlos en orden de probabilidad con argumento morfológico.
6. Sin asteriscos. Títulos en MAYÚSCULAS. Devuelve el informe COMPLETO sin alterar otras secciones.

INFORME ACTUAL:\n{rep2}"""}],
                            temperature=0.2,max_tokens=2500
                        )
                        st.session_state.reporte=r.choices[0].message.content; st.rerun()
                    except Exception as e: st.error(str(e))

    with a2:
        if st.button("◇ Definiciones",use_container_width=True):
            rep2=st.session_state.reporte
            if rep2 and api_key:
                cl=get_client(); mid=MODELS[st.session_state.modelo]["id"]
                with st.spinner("Analizando..."):
                    try:
                        r=cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":
                                f"""Eres un radiólogo subespecialista y docente universitario. Analiza el siguiente informe radiológico con profundidad académica y clínica. Responde en español. Sin asteriscos ni markdown.

Usa este formato exacto:

CLASIFICACIONES UTILIZADAS
Para cada clasificación mencionada en el informe:
  · Sistema: [nombre completo del sistema · sociedad que lo avala]
  · Grado asignado: [grado] — Significado: [descripción clínica del grado]
  · Evidencia en el informe: [cita textual del hallazgo que lo justifica]
  · Referencia seminal: [Autor(es), Título abreviado, Revista, Año, DOI si disponible]
  · Relevancia clínica: [qué implica este grado para el manejo del paciente]

CLASIFICACIONES ADICIONALES RECOMENDADAS
[Solo si el informe describe hallazgos que se podrían gradificar con sistemas no utilizados. Si no aplica: "El informe utiliza los sistemas apropiados para los hallazgos descritos."]
  · Sistema sugerido: [nombre] — Hallazgo que lo justifica: [descripción]
  · Por qué añadiría valor: [impacto clínico o quirúrgico]

GLOSARIO DE TÉRMINOS TÉCNICOS
Para cada término especializado del informe:
  · [Término]: [definición precisa en 2-3 líneas, con contexto anatómico y relevancia diagnóstica]

FISIOPATOLOGÍA RELEVANTE
[Párrafo de 3-4 líneas explicando el mecanismo fisiopatológico subyacente a los hallazgos principales. Nivel: residente avanzado / fellow.]

CORRELACIÓN CLÍNICA Y ORIENTACIÓN AL TRATANTE
[Párrafo de 4-5 líneas dirigido al médico clínico: qué implican estos hallazgos para el paciente, qué opciones terapéuticas se abren, qué estudios complementarios podrían ser útiles, y cuál sería el seguimiento radiológico recomendado.]

INFORME A ANALIZAR:
{rep2}"""}],
                            temperature=0.15,max_tokens=2000
                        )
                        st.session_state.defs=r.choices[0].message.content
                        st.session_state.panel_der=True; st.rerun()
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

# ═══════════════════════════════════════════════════════════════
# PANEL DERECHO
# ═══════════════════════════════════════════════════════════════
with col_r:
    toggle_r="▶" if right_open else "◀"
    if st.button(toggle_r,key="tog_r",help="Expandir / colapsar definiciones"):
        st.session_state.panel_der=not right_open; st.rerun()

    if not right_open:
        st.markdown(f"""<div style="display:flex;flex-direction:column;align-items:center;gap:14px;padding:12px 0">
          <span style="font-size:15px;color:{MUTED}">📖</span>
          <span style="font-size:15px;color:{MUTED}">🔗</span>
        </div>""",unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="section-title" style="margin-top:6px"><div class="dot"></div>Definiciones y referencias</div>',unsafe_allow_html=True)
        if st.session_state.defs:
            st.markdown(f'<div class="defs-box">{st.session_state.defs}</div>',unsafe_allow_html=True)
            st.markdown("<div style='height:8px'></div>",unsafe_allow_html=True)
            if st.button("Cerrar",key="close_defs"):
                st.session_state.defs=""; st.rerun()
        else:
            st.markdown(f"""<div style="padding:20px 0;text-align:center">
              <p style="font-size:12px;color:{MUTED};line-height:1.7">
                Genera un informe y presiona<br>
                <strong style="color:{TEXT}">◇ Definiciones</strong><br>
                para ver clasificaciones,<br>definiciones y referencias.
              </p>
            </div>""",unsafe_allow_html=True)
