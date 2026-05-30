
import streamlit as st
import streamlit.components.v1 as components
from openai import OpenAI
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile, io, os, re, json, datetime
 
st.set_page_config(page_title="AURA", layout="wide", initial_sidebar_state="collapsed")
 
# ── Modelos ──────────────────────────────────────────────────
MODELS = {
    "DeepSeek Chat": {"url": "https://api.deepseek.com", "id": "deepseek-chat"},
    "GPT-4o Mini":   {"url": None, "id": "gpt-4o-mini"},
    "GPT-4.1 Mini":  {"url": None, "id": "gpt-4.1-mini"},
}
 
MODALIDADES = [
    "Resonancia Magnetica", "Tomografia Computarizada", "Radiografia",
    "Ultrasonido", "PET-CT", "Mamografia", "Fluoroscopia", "Angiografia",
]
 
REGIONES = {
    "Extremidades inferiores": ["Rodilla","Cadera","Tobillo","Pie","Muslo","Pierna"],
    "Extremidades superiores": ["Hombro","Codo","Muneca","Mano","Brazo","Antebrazo"],
    "Columna":                 ["Col. cervical","Col. dorsal","Col. lumbar","Sacro"],
    "Craneo y cuello":         ["Cerebro","Cuello","Tiroides","Orbitas","Oidos"],
    "Torax":                   ["Torax","Pulmon","Corazon","Mediastino","Mama"],
    "Abdomen y pelvis":        ["Abdomen","Pelvis","Higado","Pancreas","Rinones",
                                "Vejiga","Prostata","Utero/Anexos","Suprarrenales"],
}
 
PLANTILLAS = {
    "Ninguna": "",
    "MSK - RM": "INDICACION\n\nTECNICA\nRM de [region] en equipo de [campo] Tesla. Secuencias SE/FSE en planos axial, sagital y coronal, con y sin saturacion grasa.\n\nHALLAZGOS\nPartes blandas periarticulares:\nHueso y medula osea:\nCartilago articular:\nMeniscos / Fibrocartilago:\nLigamentos:\nTendones:\nLiquido articular:\n\nIMPRESION DIAGNOSTICA\n",
    "Neuro - RM Cerebro": "INDICACION\n\nTECNICA\nRM cerebral en equipo de [campo] Tesla. Secuencias T1, T2, FLAIR, DWI/ADC, T2*. [Contraste: si/no].\n\nHALLAZGOS\nParenquima supratentorial:\nParenquima infratentorial:\nSistema ventricular:\nLinea media:\nVasculatura:\nSenos paranasales:\n\nIMPRESION DIAGNOSTICA\n",
    "TC Abdomen": "INDICACION\n\nTECNICA\nTC de abdomen [y pelvis], fase [simple/arterial/portal]. Contraste yodado IV [si/no].\n\nHALLAZGOS\nHigado:\nVias biliares y vesicula:\nPancreas:\nBazo:\nSuprarrenales:\nRinones y vias urinarias:\nRetroperitoneo:\nAsas intestinales:\nPelvis:\n\nIMPRESION DIAGNOSTICA\n",
    "Torax Rx/TC": "INDICACION\n\nTECNICA\n[Radiografia PA y lateral / TC de torax]. [Fase respiratoria]. [Contraste: si/no].\n\nHALLAZGOS\nParenquima pulmonar:\nHilios:\nMediastino:\nSilueta cardiaca:\nPleura:\nPared toracica y huesos:\n\nIMPRESION DIAGNOSTICA\n",
}
 
DEFAULTS = {
    "dictado": "", "reporte": "", "mentor_fb": "", "sugerencias": [],
    "modelo": "DeepSeek Chat", "audio_id": None,
    "historial": [], "plantilla_key": "Ninguna",
    "plantillas_extra": {},
    "diccionario": {},
    "estilo": [],
    "tema": "dark",
    "vista": "transcripcion",
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v
 
try:    api_key = st.secrets["deepseek_key"]
except: api_key = os.environ.get("OPENAI_API_KEY", "")
 
# ── Helpers ──────────────────────────────────────────────────
def get_client():
    cfg = MODELS[st.session_state.modelo]
    return OpenAI(api_key=api_key, base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)
 
def limpiar(txt):
    txt = re.sub(r'\*\*(.+?)\*\*', r'\1', txt)
    txt = re.sub(r'\*(.+?)\*',     r'\1', txt)
    txt = re.sub(r'^\*\s+', '• ',  txt, flags=re.MULTILINE)
    txt = re.sub(r'\*+', '',        txt)
    txt = re.sub(r'^#+\s*', '',     txt, flags=re.MULTILINE)
    return txt.strip()
 
def completitud(txt):
    up = txt.upper()
    s = sum(1 for x in ["TECNICA","HALLAZGOS","IMPRESION"] if x in up)
    w = len(txt.split())
    return min(100, int((s/3)*60 + min(w/150,1)*40)), w
 
def generar_docx(texto):
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    for line in texto.split("\n"):
        s = re.sub(r'\*+','',line).strip()
        if not s: doc.add_paragraph(); continue
        if s.isupper() and len(s) < 80:
            h = doc.add_heading(s, level=1)
            h.alignment = WD_ALIGN_PARAGRAPH.LEFT
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
    with open(path, "rb") as f:
        res = cl.audio.transcriptions.create(
            model="whisper-1", file=f, language="es",
            prompt="Dictado radiologico: Stoller, ICRS, LCA, Pfirrmann, Modic, TIRADS, BIRADS, PIRADS."
        )
    os.unlink(path)
    return res.text.strip()
 
def estilo_ctx():
    items = st.session_state.estilo[-5:]
    if not items: return ""
    return "\n\n".join([f"Ejemplo {i+1}:\n{e['r']}" for i,e in enumerate(items)])
 
def todas_plantillas():
    return {**PLANTILLAS, **st.session_state.plantillas_extra}
 
# ── Tema ─────────────────────────────────────────────────────
D = {
    "bg":      "#0c0c0f",
    "sb":      "#111115",
    "panel":   "#141418",
    "card":    "#1a1a20",
    "border":  "#242430",
    "acc":     "#7c6af7",
    "acc2":    "#a89cf9",
    "txt":     "#e4e4f0",
    "mut":     "#4a4a68",
    "mut2":    "#2e2e40",
    "green":   "#22c55e",
    "edbg":    "#0f0f13",
    "surf":    "#1e1e26",
    "genbtn":  "linear-gradient(135deg,#7c6af7 0%,#5040d0 100%)",
}
L = {
    "bg":      "#f5f5f8",
    "sb":      "#ffffff",
    "panel":   "#ffffff",
    "card":    "#ededf3",
    "border":  "#dcdce8",
    "acc":     "#5b4de0",
    "acc2":    "#7c6af7",
    "txt":     "#141420",
    "mut":     "#8888a8",
    "mut2":    "#d0d0e0",
    "green":   "#16a34a",
    "edbg":    "#ffffff",
    "surf":    "#f0f0f5",
    "genbtn":  "linear-gradient(135deg,#5b4de0 0%,#7c6af7 100%)",
}
 
def TH(): return D if st.session_state.tema == "dark" else L
 
# ── CSS ───────────────────────────────────────────────────────
def apply_css():
    t = TH()
    st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body, .stApp {{
    background: {t['bg']} !important;
    color: {t['txt']} !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 13px !important;
    overflow: hidden !important;
    height: 100vh !important;
}}
header, footer, #MainMenu, .stDeployButton {{ visibility: hidden !important; height: 0 !important; }}
.block-container {{ padding: 0 !important; max-width: 100vw !important; }}
[data-testid="column"] {{ padding: 0 !important; }}
 
/* ── Scrollbars ── */
::-webkit-scrollbar {{ width: 4px; height: 4px; }}
::-webkit-scrollbar-track {{ background: transparent; }}
::-webkit-scrollbar-thumb {{ background: {t['border']}; border-radius: 2px; }}
::-webkit-scrollbar-thumb:hover {{ background: {t['acc']}66; }}
 
/* ── Inputs ── */
[data-testid="stSelectbox"] > div > div {{
    background: {t['card']} !important; border: 1px solid {t['border']} !important;
    border-radius: 8px !important; color: {t['txt']} !important;
    font-size: 12px !important; font-family: 'Inter', sans-serif !important;
}}
[data-testid="stSelectbox"] > div > div:hover {{ border-color: {t['acc']}55 !important; }}
.stTextInput input {{
    background: {t['card']} !important; border: 1px solid {t['border']} !important;
    border-radius: 8px !important; color: {t['txt']} !important;
    font-size: 12px !important; padding: 8px 12px !important;
    font-family: 'Inter', sans-serif !important;
}}
.stTextInput input:focus {{ border-color: {t['acc']}55 !important; box-shadow: none !important; }}
.stTextInput input::placeholder {{ color: {t['mut']} !important; }}
.stTextArea textarea {{
    background: {t['edbg']} !important; border: 1px solid {t['border']} !important;
    border-radius: 8px !important; color: {t['txt']} !important;
    font-size: 12.5px !important; line-height: 1.75 !important;
    padding: 12px 14px !important; font-family: 'Inter', sans-serif !important;
    caret-color: {t['acc']} !important;
}}
.stTextArea textarea:focus {{ border-color: {t['acc']}44 !important; box-shadow: none !important; }}
.stTextArea textarea::placeholder {{ color: {t['mut']} !important; }}
[data-testid="stAudioInput"] {{
    background: {t['card']} !important; border: 1px solid {t['border']} !important;
    border-radius: 8px !important;
}}
[data-testid="stFileUploader"] {{
    background: {t['card']}; border: 1px dashed {t['border']}; border-radius: 8px; padding: 6px;
}}
[data-testid="stFileUploader"] * {{ color: {t['mut']} !important; font-size: 11px !important; }}
 
/* ── Botones ── */
.stButton button {{
    background: {t['card']} !important; border: 1px solid {t['border']} !important;
    color: {t['txt']} !important; border-radius: 8px !important;
    font-size: 12px !important; font-weight: 400 !important;
    font-family: 'Inter', sans-serif !important; transition: all .15s !important;
    white-space: nowrap !important;
}}
.stButton button:hover {{
    border-color: {t['acc']}55 !important; background: {t['surf']} !important;
    color: {t['acc2']} !important;
}}
.btn-gen .stButton button {{
    background: {t['genbtn']} !important; border: none !important;
    color: #fff !important; font-weight: 600 !important; font-size: 13px !important;
    border-radius: 10px !important;
    box-shadow: 0 4px 20px {t['acc']}44 !important;
}}
.btn-gen .stButton button:hover {{ opacity: .9 !important; box-shadow: 0 6px 26px {t['acc']}66 !important; }}
.stDownloadButton button {{
    background: transparent !important; border: 1px solid {t['acc']}55 !important;
    color: {t['acc']} !important; border-radius: 8px !important; font-size: 12px !important;
    font-family: 'Inter', sans-serif !important;
}}
.stDownloadButton button:hover {{ background: {t['acc']}0e !important; }}
 
/* ── Expander ── */
[data-testid="stExpander"] {{
    background: {t['card']} !important; border: 1px solid {t['border']} !important;
    border-radius: 8px !important; margin-bottom: 6px !important;
}}
[data-testid="stExpander"] summary {{
    color: {t['mut']} !important; font-size: 12px !important;
    font-family: 'Inter', sans-serif !important; padding: 8px 12px !important;
}}
[data-testid="stExpander"] summary:hover {{ color: {t['txt']} !important; }}
 
/* ── Tabs ── */
[data-testid="stTabs"] [role="tablist"] {{
    border-bottom: 1px solid {t['border']} !important; background: transparent !important;
}}
[data-testid="stTabs"] [role="tab"] {{
    background: transparent !important; border: none !important; color: {t['mut']} !important;
    font-size: 12px !important; padding: 7px 12px !important;
    border-bottom: 2px solid transparent !important; border-radius: 0 !important;
    font-family: 'Inter', sans-serif !important;
}}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {{
    color: {t['acc']} !important; border-bottom-color: {t['acc']} !important;
}}
[data-testid="stTabs"] [data-baseweb="tab-panel"] {{
    background: transparent !important; padding: 10px 0 0 !important;
}}
 
/* ── Componentes custom ── */
.lbl {{ font-size: 10px; font-weight: 600; letter-spacing: .12em; text-transform: uppercase;
    color: {t['mut']}; margin-bottom: 4px; display: block; }}
.chip {{ display: inline-flex; align-items: center; gap: 4px; font-size: 10px;
    padding: 2px 8px; border-radius: 20px; border: 1px solid {t['border']};
    color: {t['mut']}; background: {t['card']}; white-space: nowrap; }}
.chip.on {{ border-color: {t['acc']}44; color: {t['acc2']}; background: {t['acc']}14; }}
.sug-card {{
    background: {t['surf']}; border: 1px solid {t['acc']}22;
    border-left: 3px solid {t['acc']}; border-radius: 0 8px 8px 0;
    padding: 11px 13px; margin-bottom: 8px; font-size: 11.5px; line-height: 1.65;
}}
.mentor-card {{
    background: {t['surf']}; border: 1px solid {t['border']};
    border-left: 3px solid {t['acc']}; border-radius: 0 8px 8px 0;
    padding: 13px; font-size: 11.5px; line-height: 1.7;
    color: {t['txt']}; white-space: pre-wrap;
}}
hr {{ border: none !important; border-top: 1px solid {t['border']} !important; margin: 10px 0 !important; }}
</style>
""", unsafe_allow_html=True)
 
apply_css()
t = TH()
 
# ═══════════════════════════════════════════════════════════════
# LAYOUT: col_left (entrada) | col_right (editor informe)
# Layout 2 columnas: sin sidebar saturada
# ═══════════════════════════════════════════════════════════════
col_L, col_R = st.columns([0.42, 0.58], gap="small")
 
# ═══════════════════════════════════════════════════════════════
# COLUMNA IZQUIERDA — Entrada + controles
# ═══════════════════════════════════════════════════════════════
with col_L:
    # ── Topbar izquierda ─────────────────────────────────────
    st.markdown(f"""
    <div style="height:48px;background:{t['panel']};border-bottom:1px solid {t['border']};
        border-right:1px solid {t['border']};
        display:flex;align-items:center;padding:0 16px;gap:12px;position:sticky;top:0;z-index:10">
        <span style="font-size:18px;font-weight:700;color:{t['acc']};letter-spacing:.08em">AURA</span>
        <div style="width:1px;height:16px;background:{t['border']}"></div>
        <span style="font-size:11px;color:{t['mut']}">Dictado</span>
        <div style="margin-left:auto;display:flex;gap:6px;align-items:center">
            <span class="chip {'on' if st.session_state.plantilla_key != 'Ninguna' else ''}">
                {'Plantilla: ' + st.session_state.plantilla_key[:12] if st.session_state.plantilla_key != 'Ninguna' else 'Sin plantilla'}
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)
 
    # ── Contenido scrolleable ─────────────────────────────────
    # Usamos un contenedor con altura fija y overflow
    st.markdown(f"""
    <div style="height:calc(100vh - 48px);overflow-y:auto;overflow-x:hidden;
        background:{t['bg']};border-right:1px solid {t['border']};padding:16px;">
    """, unsafe_allow_html=True)
 
    # Configuracion estudio (colapsable)
    with st.expander("Configurar estudio", expanded=False):
        st.markdown('<span class="lbl">Modalidad</span>', unsafe_allow_html=True)
        st.selectbox("mod", MODALIDADES, label_visibility="collapsed", key="sel_mod")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<span class="lbl">Grupo</span>', unsafe_allow_html=True)
            grupo = st.selectbox("grp", list(REGIONES.keys()),
                                 label_visibility="collapsed", key="sel_grupo")
        with c2:
            st.markdown('<span class="lbl">Region</span>', unsafe_allow_html=True)
            st.selectbox("reg", REGIONES[grupo],
                         label_visibility="collapsed", key="sel_reg")
        st.markdown('<span class="lbl" style="margin-top:5px">Region personalizada</span>',
                    unsafe_allow_html=True)
        st.text_input("rc", label_visibility="collapsed", key="reg_custom",
                      placeholder="Ej: Articulacion glenohumeral derecha")
 
    # Plantilla activa
    with st.expander("Plantilla", expanded=False):
        pts = todas_plantillas()
        sel_pt = st.selectbox("pt_sel", list(pts.keys()),
                              index=list(pts.keys()).index(st.session_state.plantilla_key)
                              if st.session_state.plantilla_key in pts else 0,
                              label_visibility="collapsed", key="pt_sel_box")
        if sel_pt != st.session_state.plantilla_key:
            st.session_state.plantilla_key = sel_pt; st.rerun()
        if pts[sel_pt]:
            st.markdown(f"""<div style="background:{t['edbg']};border:1px solid {t['border']};
                border-radius:8px;padding:10px;font-size:11px;line-height:1.6;
                color:{t['mut']};white-space:pre-wrap;max-height:200px;overflow-y:auto;
                margin-top:8px">{pts[sel_pt]}</div>""", unsafe_allow_html=True)
        st.markdown("<hr>", unsafe_allow_html=True)
        f_up = st.file_uploader("Cargar .docx", type=["docx"], key="pt_uploader")
        if f_up:
            doc = Document(f_up)
            txt = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
            nombre = f_up.name.replace(".docx","")
            st.session_state.plantillas_extra[nombre] = txt
            st.session_state.plantilla_key = nombre
            st.success(f"Plantilla '{nombre}' cargada")
            st.rerun()
 
    # Diccionario
    with st.expander(f"Diccionario ({len(st.session_state.diccionario)} terminos)", expanded=False):
        ct, cd = st.columns([1,2])
        with ct: nt = st.text_input("dterm", label_visibility="collapsed",
                                     placeholder="Termino", key="dic_t")
        with cd: nd = st.text_input("ddef", label_visibility="collapsed",
                                     placeholder="Definicion", key="dic_d")
        if st.button("+ Agregar termino", key="btn_dic_add", use_container_width=True):
            if nt and nd:
                st.session_state.diccionario[nt] = nd; st.rerun()
        if st.session_state.diccionario:
            for term, defn in list(st.session_state.diccionario.items()):
                ca, cb = st.columns([3,1])
                with ca: st.markdown(f'<span style="font-size:11px;color:{t["acc2"]}">'
                                     f'<b>{term}</b>: <span style="color:{t["mut"]}">{defn}</span></span>',
                                     unsafe_allow_html=True)
                with cb:
                    if st.button("x", key=f"del_{term}"):
                        del st.session_state.diccionario[term]; st.rerun()
 
    st.markdown("<hr>", unsafe_allow_html=True)
 
    # ── Entrada de dictado ────────────────────────────────────
    tab_voz, tab_kbd = st.tabs(["Dictado", "Teclado"])
 
    with tab_voz:
        st.markdown(f"""
        <div style="display:flex;flex-direction:column;align-items:center;
            padding:20px 0 14px;gap:10px">
            <div style="position:relative;width:72px;height:72px">
                <div style="position:absolute;inset:-12px;border-radius:50%;
                    border:1px solid {t['acc']}22;animation:rp 2.6s ease-out infinite"></div>
                <div style="position:absolute;inset:-5px;border-radius:50%;
                    border:1px solid {t['acc']}38;animation:rp 2.6s ease-out infinite .6s"></div>
                <div style="width:72px;height:72px;border-radius:50%;
                    background:{t['acc']};
                    display:flex;align-items:center;justify-content:center;
                    box-shadow:0 6px 24px {t['acc']}44">
                    <svg width="26" height="26" viewBox="0 0 24 24" fill="none"
                        stroke="#fff" stroke-width="1.8" stroke-linecap="round">
                        <rect x="9" y="2" width="6" height="12" rx="3"/>
                        <path d="M5 10a7 7 0 0 0 14 0"/>
                        <line x1="12" y1="19" x2="12" y2="22"/>
                        <line x1="9" y1="22" x2="15" y2="22"/>
                    </svg>
                </div>
            </div>
            <span style="font-size:12px;color:{t['mut']}">Pulsa para grabar</span>
        </div>
        <style>@keyframes rp{{0%{{transform:scale(1);opacity:.5}}100%{{transform:scale(1.5);opacity:0}}}}</style>
        """, unsafe_allow_html=True)
 
        audio = st.audio_input("", label_visibility="collapsed", key="audio_rec")
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
                    st.warning("Configura tu API Key.")
 
        if st.session_state.dictado:
            st.markdown('<span class="lbl">Transcripcion</span>', unsafe_allow_html=True)
            d = st.text_area("dt_v", value=st.session_state.dictado,
                             height=120, label_visibility="collapsed", key="dictado_voz")
            if d != st.session_state.dictado:
                st.session_state.dictado = d
 
    with tab_kbd:
        d = st.text_area("dt_k", value=st.session_state.dictado,
                         height=220, label_visibility="collapsed", key="dictado_kbd",
                         placeholder="Escribe los hallazgos.\n\nEj: Desgarro horizontal menisco medial cuerno posterior Stoller III, extrusion 3mm. LCA integro.")
        if d != st.session_state.dictado:
            st.session_state.dictado = d
 
    # Configuracion y modelo
    with st.expander("Ajustes", expanded=False):
        st.markdown('<span class="lbl">Modelo IA</span>', unsafe_allow_html=True)
        m = st.selectbox("cfg_m", list(MODELS.keys()),
                         index=list(MODELS.keys()).index(st.session_state.modelo),
                         label_visibility="collapsed", key="cfg_model")
        if m != st.session_state.modelo:
            st.session_state.modelo = m; st.rerun()
        if not api_key:
            st.markdown('<span class="lbl" style="margin-top:6px">API Key</span>',
                        unsafe_allow_html=True)
            api_key = st.text_input("cfg_k", type="password",
                                    label_visibility="collapsed",
                                    placeholder="sk-...", key="cfg_key")
        tema_lbl = "Modo claro" if st.session_state.tema == "dark" else "Modo oscuro"
        if st.button(tema_lbl, key="btn_tema", use_container_width=True):
            st.session_state.tema = "light" if st.session_state.tema == "dark" else "dark"
            st.rerun()
        n_est = len(st.session_state.estilo)
        st.markdown(f'<p style="font-size:11px;color:{t["mut"]};margin-top:8px">'
                    f'Estilo aprendido: {n_est} ejemplo{"s" if n_est!=1 else ""}</p>',
                    unsafe_allow_html=True)
        if n_est and st.button("Borrar ejemplos", key="btn_del_est"):
            st.session_state.estilo = []; st.rerun()
 
    # Historial
    if st.session_state.historial:
        with st.expander(f"Historial ({len(st.session_state.historial)})", expanded=False):
            HCOLS = ["#7c6af7","#22c55e","#f59e0b","#ec4899","#38bdf8","#fb923c"]
            for i, e in enumerate(st.session_state.historial):
                color = HCOLS[i % len(HCOLS)]
                ca, cb = st.columns([3,1])
                with ca:
                    st.markdown(f"""<div style="display:flex;align-items:center;gap:7px;
                        padding:6px 8px;background:{t['surf']};border:1px solid {t['border']};
                        border-radius:7px;margin-bottom:3px">
                        <div style="width:6px;height:6px;border-radius:50%;background:{color};flex-shrink:0"></div>
                        <div>
                            <div style="font-size:11px;color:{t['txt']}">{e['region']}</div>
                            <div style="font-size:10px;color:{t['mut']}">{e['mod'][:18]} · {e.get('fecha','')}</div>
                        </div>
                    </div>""", unsafe_allow_html=True)
                with cb:
                    if st.button("Cargar", key=f"h_{i}"):
                        st.session_state.reporte = e['txt']; st.rerun()
 
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
 
    # ── Boton Generar ─────────────────────────────────────────
    cg, cn = st.columns([3,1])
    with cg:
        st.markdown('<div class="btn-gen">', unsafe_allow_html=True)
        generar = st.button("Generar Informe", key="btn_gen", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with cn:
        if st.button("Nuevo", key="btn_nuevo", use_container_width=True):
            st.session_state.dictado = ""
            st.session_state.reporte = ""
            st.session_state.audio_id = None
            st.session_state.mentor_fb = ""
            st.session_state.sugerencias = []
            st.rerun()
 
    st.markdown("</div>", unsafe_allow_html=True)
 
# ─────────────────────────────────────────────────────────────
# PROCESAMIENTO IA
# ─────────────────────────────────────────────────────────────
if generar:
    if not api_key:
        st.warning("Configura tu API Key en Ajustes.")
    elif not st.session_state.dictado.strip():
        st.warning("Escribe o dicta los hallazgos primero.")
    else:
        cl  = get_client()
        mid = MODELS[st.session_state.modelo]["id"]
        pt  = todas_plantillas().get(st.session_state.plantilla_key, "")
 
        mod_sel = st.session_state.get("sel_mod", "")
        reg_sel = (st.session_state.get("reg_custom","").strip()
                   or st.session_state.get("sel_reg",""))
 
        instruc_pt = (
            f"ESTRUCTURA OBLIGATORIA — respeta exactamente estas secciones:\n\n{pt}"
            if pt else
            "Estructura: INDICACION / TECNICA / HALLAZGOS / IMPRESION DIAGNOSTICA"
        )
 
        dic = st.session_state.diccionario
        instruc_dic = (
            "DICCIONARIO DEL RADIOLOGO:\n" +
            "\n".join(f"- {k}: {v}" for k,v in dic.items())
            if dic else ""
        )
 
        ec = estilo_ctx()
        instruc_estilo = (
            f"ESTILO APRENDIDO (replica este patron de redaccion):\n{ec}"
            if ec else ""
        )
 
        prompt = f"""Eres AURA, sistema experto de interpretacion radiologica de nivel subespecialista.
 
MODALIDAD: {mod_sel}
REGION: {reg_sel}
 
REGLAS DE FORMATO — OBLIGATORIAS:
1. CERO asteriscos (*). Prohibido absolutamente.
2. CERO markdown (sin #, sin **, sin *).
3. Titulos de seccion: MAYUSCULAS, solos en su linea.
4. HALLAZGOS: prosa narrativa continua. Jamas listas ni guiones.
   Los hallazgos fluyen como parrafos de texto medico sofisticado.
   Conecta estructuras anatomicas, establece relaciones, describe en contexto.
   Ejemplo CORRECTO: "El menisco medial presenta desgarro horizontal en el cuerno posterior, Stoller grado III, con extrusion de 3 mm. El compartimento medial muestra reduccion del espacio articular con esclerosis subcondral y osteofitos marginales. El ligamento cruzado anterior se encuentra integro..."
   Ejemplo PROHIBIDO: "- Menisco: desgarro\n- LCA: integro"
5. IMPRESION DIAGNOSTICA: usa punto-bullet (•) para cada hallazgo jerarquizado.
   Cada bullet: estructura + diagnostico especifico + clasificacion/grado + implicacion clinica.
   Ultimo bullet: correlacion clinico-radiologica y orientacion de manejo.
6. Cuantifica siempre: mm, %, grados, scores validados.
7. Clasificaciones segun hallazgo: Stoller, ICRS/Outerbridge, Pfirrmann, Modic, Meyerding,
   Bigliani, Goutallier, Sugaya, Tonnis, ASPECTS, Fazekas, BIRADS, PIRADS, TIRADS.
   Solo aplica las respaldadas directamente por el dictado.
8. Nivel: publicable en revista indexada.
 
{instruc_dic}
{instruc_estilo}
{instruc_pt}
 
DICTADO:
{st.session_state.dictado}"""
 
        with st.spinner("Generando informe..."):
            try:
                res = cl.chat.completions.create(
                    model=mid,
                    messages=[{"role":"system","content":prompt}],
                    temperature=0.12, max_tokens=3000
                )
                report = limpiar(res.choices[0].message.content)
                st.session_state.reporte = report
                st.session_state.mentor_fb = ""
                st.session_state.sugerencias = []
                fecha = datetime.datetime.now().strftime("%d/%m %H:%M")
                st.session_state.historial.insert(0, {
                    "mod": mod_sel[:20], "region": reg_sel or "General",
                    "txt": report, "fecha": fecha,
                })
                if len(st.session_state.historial) > 20:
                    st.session_state.historial = st.session_state.historial[:20]
                st.rerun()
            except Exception as e:
                st.error(str(e))
 
# ═══════════════════════════════════════════════════════════════
# COLUMNA DERECHA — Editor de informe
# ═══════════════════════════════════════════════════════════════
with col_R:
    rep = st.session_state.reporte
    t = TH()
    pct, words = completitud(rep) if rep else (0, 0)
 
    # ── Topbar derecha ────────────────────────────────────────
    st.markdown(f"""
    <div style="height:48px;background:{t['panel']};border-bottom:1px solid {t['border']};
        display:flex;align-items:center;padding:0 14px;gap:10px;position:sticky;top:0;z-index:10">
        <span style="font-size:11px;font-weight:600;letter-spacing:.1em;
            text-transform:uppercase;color:{t['mut']}">Informe</span>
        <div style="margin-left:auto;display:flex;align-items:center;gap:8px">
            <div style="width:80px;height:2px;background:{t['border']};
                border-radius:1px;overflow:hidden">
                <div style="width:{pct}%;height:100%;background:{t['acc']};
                    border-radius:1px;transition:width .4s"></div>
            </div>
            <span style="font-size:10px;color:{t['mut']}">{pct}%</span>
            <span style="font-size:10px;color:{t['mut']}">{words} pal.</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
 
    # ── Editor rico ───────────────────────────────────────────
    def to_html(texto):
        if not texto: return ""
        parts = []
        for line in texto.split("\n"):
            s = line.strip()
            if not s:
                parts.append("<p><br></p>")
            elif s.isupper() and 2 < len(s) < 76:
                parts.append(f"<h2>{s}</h2>")
            elif s.startswith("•"):
                parts.append(f"<li>{s[1:].strip()}</li>")
            else:
                parts.append(f"<p>{s}</p>")
        return "\n".join(parts)
 
    contenido = to_html(rep)
    t = TH()
 
    # Altura: viewport menos topbars y acciones
    editor_h = 480
 
    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{
    width: 100%; height: {editor_h + 76}px;
    display: flex; flex-direction: column;
    background: {t['bg']}; font-family: 'Inter', sans-serif;
    overflow: hidden;
}}
 
/* TOOLBAR */
.tb {{
    flex-shrink: 0; height: 38px;
    background: {t['panel']};
    border-bottom: 1px solid {t['border']};
    padding: 0 10px;
    display: flex; align-items: center; gap: 2px;
    overflow-x: auto;
}}
.tb::-webkit-scrollbar {{ height: 2px; }}
.tb::-webkit-scrollbar-thumb {{ background: {t['border']}; }}
.tg {{
    display: flex; align-items: center; gap: 1px;
    padding-right: 6px; margin-right: 2px;
    border-right: 1px solid {t['border']}; flex-shrink: 0;
}}
.tg:last-child {{ border-right: none; }}
.b {{
    width: 26px; height: 26px;
    background: none; border: 1px solid transparent;
    color: {t['mut']}; font-size: 11px;
    border-radius: 5px; cursor: pointer;
    transition: all .1s; display: flex;
    align-items: center; justify-content: center;
    font-family: 'Inter', sans-serif;
}}
.b:hover {{ background: {t['card']}; color: {t['txt']}; border-color: {t['border']}; }}
.b.on {{ background: {t['acc']}1a; color: {t['acc']}; border-color: {t['acc']}44; }}
.s {{
    height: 26px; background: {t['card']};
    border: 1px solid {t['border']};
    color: {t['mut']}; font-size: 10px;
    border-radius: 5px; padding: 0 5px;
    outline: none; cursor: pointer;
    font-family: 'Inter', sans-serif;
}}
.s:focus {{ border-color: {t['acc']}55; }}
 
/* SCROLL WRAPPER — este es el que hace scroll */
.sw {{
    flex: 1;
    overflow-y: auto;
    overflow-x: hidden;
    background: {t['card']};
    min-height: 0;
}}
.sw::-webkit-scrollbar {{ width: 4px; }}
.sw::-webkit-scrollbar-track {{ background: transparent; }}
.sw::-webkit-scrollbar-thumb {{ background: {t['border']}; border-radius: 2px; }}
.sw::-webkit-scrollbar-thumb:hover {{ background: {t['acc']}66; }}
 
/* PAPER — el editor en si */
.paper {{
    min-height: 100%;
    padding: 28px 32px 48px;
    background: {t['edbg']};
    outline: none;
    font-family: 'Inter', sans-serif;
    font-size: 13px;
    line-height: 1.82;
    color: {t['txt']};
    word-break: break-word;
}}
.paper:focus {{ outline: none; }}
.paper:empty::before {{
    content: attr(data-ph);
    color: {t['mut']}; pointer-events: none;
    font-size: 12px; opacity: .7;
}}
.paper h2 {{
    font-size: 10px; font-weight: 600;
    letter-spacing: .16em; text-transform: uppercase;
    color: {t['acc']};
    margin: 22px 0 8px;
    padding-bottom: 5px;
    border-bottom: 1px solid {t['border']};
}}
.paper p {{ margin: 1px 0; }}
.paper li {{ margin-left: 18px; margin-bottom: 3px; }}
.paper ul, .paper ol {{ margin: 4px 0 4px 18px; }}
.paper hr {{ border: none; border-top: 1px solid {t['border']}; margin: 12px 0; }}
.paper table {{ border-collapse: collapse; width: 100%; margin: 10px 0; font-size: 12px; }}
.paper td, .paper th {{ border: 1px solid {t['border']}; padding: 6px 10px; color: {t['txt']}; }}
.paper th {{ background: {t['surf']}; font-weight: 600; color: {t['acc']}; font-size: 10px; letter-spacing: .08em; }}
.paper tr:nth-child(even) td {{ background: {t['surf']}55; }}
 
/* STATUS BAR */
.sb {{
    flex-shrink: 0; height: 28px;
    background: {t['panel']}; border-top: 1px solid {t['border']};
    display: flex; align-items: center; padding: 0 12px; gap: 8px;
}}
.sbw {{ font-size: 10px; color: {t['mut']}; }}
.sbt {{ margin-left: auto; width: 80px; height: 2px;
    background: {t['border']}; border-radius: 1px; overflow: hidden; }}
.sbf {{ height: 100%; background: {t['acc']}; border-radius: 1px; transition: width .4s; }}
.sbp {{ font-size: 10px; color: {t['mut']}; }}
</style>
</head>
<body>
 
<!-- TOOLBAR -->
<div class="tb">
    <div class="tg">
        <select class="s" style="width:82px" onchange="setFont(this.value)">
            <option value="'Inter',sans-serif" selected>Inter</option>
            <option value="Georgia,serif">Georgia</option>
            <option value="Calibri,sans-serif">Calibri</option>
            <option value="'Courier New',monospace">Courier</option>
        </select>
        <select class="s" style="width:36px" onchange="setSize(this.value)">
            <option>10</option><option>11</option><option>12</option>
            <option selected>13</option><option>14</option>
            <option>15</option><option>16</option><option>18</option>
        </select>
    </div>
    <div class="tg">
        <button class="b" id="bB" onclick="fmt('bold')" title="Negrita (Ctrl+B)"><b>B</b></button>
        <button class="b" id="bI" onclick="fmt('italic')" title="Cursiva"><i>I</i></button>
        <button class="b" id="bU" onclick="fmt('underline')" title="Subrayado"><u>U</u></button>
        <button class="b" onclick="fmt('strikeThrough')" title="Tachado"
            style="text-decoration:line-through;font-size:10px">S</button>
    </div>
    <div class="tg">
        <button class="b" onclick="fmt('justifyLeft')"   title="Izquierda">&#8676;</button>
        <button class="b" onclick="fmt('justifyCenter')" title="Centro">&#9644;</button>
        <button class="b" onclick="fmt('justifyRight')"  title="Derecha">&#8677;</button>
        <button class="b" onclick="fmt('justifyFull')"   title="Justificado">&#9776;</button>
    </div>
    <div class="tg">
        <button class="b" onclick="fmt('insertUnorderedList')" title="Vinetas">&#8226;</button>
        <button class="b" onclick="fmt('insertOrderedList')"   title="Numerada">1.</button>
        <button class="b" onclick="insHR()" title="Separador" style="font-size:9px">HR</button>
        <button class="b" onclick="insTable()" title="Tabla" style="font-size:9px">Tbl</button>
    </div>
    <div class="tg" style="gap:3px">
        <label title="Color" style="display:flex;align-items:center;cursor:pointer">
            <span style="font-size:9px;color:{t['mut']};margin-right:2px">A</span>
            <input type="color" value="{t['txt']}" onchange="fmt('foreColor',this.value)"
                style="width:16px;height:16px;padding:0;border:none;border-radius:3px;cursor:pointer">
        </label>
        <label title="Resaltar" style="display:flex;align-items:center;cursor:pointer">
            <span style="font-size:9px;color:{t['mut']};margin-right:2px">HL</span>
            <input type="color" value="{t['acc']}" onchange="fmt('hiliteColor',this.value)"
                style="width:16px;height:16px;padding:0;border:none;border-radius:3px;cursor:pointer">
        </label>
    </div>
    <div class="tg">
        <button class="b" onclick="copyAll()" title="Copiar todo">&#10697;</button>
        <button class="b" onclick="printDoc()" title="Imprimir / PDF">&#128438;</button>
        <button class="b" onclick="document.execCommand('undo')" title="Deshacer">&#8617;</button>
        <button class="b" onclick="document.execCommand('redo')"  title="Rehacer">&#8618;</button>
    </div>
</div>
 
<!-- EDITOR -->
<div class="sw" id="sw">
    <div class="paper" id="paper" contenteditable="true" spellcheck="false"
        data-ph="El informe aparecera aqui.&#10;Dicta los hallazgos y presiona Generar Informe."
        oninput="sync()" onkeyup="upd()" onmouseup="upd()">
        {contenido}
    </div>
</div>
 
<!-- STATUS BAR -->
<div class="sb">
    <span class="sbw" id="sbw">0 palabras</span>
    <div style="margin-left:auto;display:flex;align-items:center;gap:6px">
        <div class="sbt"><div class="sbf" id="sbf" style="width:{pct}%"></div></div>
        <span class="sbp" id="sbp">{pct}%</span>
    </div>
</div>
 
<script>
var paper = document.getElementById('paper');
 
function fmt(cmd, val) {{
    paper.focus();
    document.execCommand(cmd, false, val || null);
    upd();
}}
 
function setFont(f) {{ paper.style.fontFamily = f; }}
function setSize(s) {{ paper.style.fontSize = s + 'px'; }}
 
function insHR() {{
    paper.focus();
    document.execCommand('insertHTML', false,
        '<hr style="border:none;border-top:1px solid {t["border"]};margin:12px 0"><br>');
}}
 
function insTable() {{
    var r = parseInt(prompt('Filas:', '3')) || 3;
    var c = parseInt(prompt('Columnas:', '3')) || 3;
    var h = '<table><thead><tr>';
    for (var i = 0; i < c; i++) h += '<th>Col ' + (i+1) + '</th>';
    h += '</tr></thead><tbody>';
    for (var j = 0; j < r-1; j++) {{
        h += '<tr>';
        for (var k = 0; k < c; k++) h += '<td>&nbsp;</td>';
        h += '</tr>';
    }}
    h += '</tbody></table><p><br></p>';
    paper.focus();
    document.execCommand('insertHTML', false, h);
}}
 
function sync() {{
    var tx = paper.innerText || '';
    var up = tx.toUpperCase();
    var s  = ['TECNICA','HALLAZGOS','IMPRESION'].filter(function(x){{ return up.includes(x); }}).length;
    var w  = tx.trim().split(/[ \\t\\n]+/).filter(Boolean).length;
    var p  = Math.min(100, Math.round((s/3)*60 + Math.min(w/150,1)*40));
    document.getElementById('sbw').textContent = w + ' palabras';
    document.getElementById('sbf').style.width = p + '%';
    document.getElementById('sbp').textContent = p + '%';
}}
 
function upd() {{
    sync();
    ['Bold','Italic','Underline'].forEach(function(c) {{
        var b = document.getElementById('b' + c[0]);
        if (b) b.classList.toggle('on', document.queryCommandState(c.toLowerCase()));
    }});
}}
 
function copyAll() {{
    var text = paper.innerText;
    if (navigator.clipboard) {{
        navigator.clipboard.writeText(text).then(function() {{ toast('Copiado'); }});
    }} else {{
        var ta = document.createElement('textarea');
        ta.value = text; ta.style.cssText = 'position:fixed;opacity:0';
        document.body.appendChild(ta); ta.select();
        document.execCommand('copy'); document.body.removeChild(ta);
        toast('Copiado');
    }}
}}
 
function printDoc() {{
    var w = window.open('', '_blank');
    w.document.write('<html><head><title>AURA Informe</title>');
    w.document.write('<style>body{{font-family:Calibri,sans-serif;font-size:12pt;line-height:1.75;margin:2.5cm;color:#111}}');
    w.document.write('h2{{font-size:11pt;font-weight:600;text-transform:uppercase;letter-spacing:.1em;');
    w.document.write('border-bottom:1px solid #ccc;padding-bottom:4px;margin:18px 0 7px}}');
    w.document.write('li{{margin-left:18px;margin-bottom:3px}}');
    w.document.write('table{{border-collapse:collapse;width:100%;margin:10px 0}}');
    w.document.write('td,th{{border:1px solid #ccc;padding:6px 10px}}');
    w.document.write('th{{background:#f4f4f4;font-weight:600}}');
    w.document.write('</style></head><body>');
    w.document.write(paper.innerHTML);
    w.document.write('</body></html>');
    w.document.close();
    setTimeout(function() {{ w.print(); }}, 350);
}}
 
function toast(msg) {{
    var el = document.createElement('div');
    el.textContent = msg;
    el.style.cssText = 'position:fixed;bottom:36px;left:50%;transform:translateX(-50%);'
        + 'background:{t["surf"]};color:{t["acc"]};border:1px solid {t["acc"]}55;'
        + 'padding:5px 14px;border-radius:5px;font-size:11px;z-index:9999;'
        + 'pointer-events:none;font-family:Inter,sans-serif;';
    document.body.appendChild(el);
    setTimeout(function() {{ document.body.removeChild(el); }}, 1600);
}}
 
paper.addEventListener('keydown', function(e) {{
    if (e.key === 'Tab') {{
        e.preventDefault();
        document.execCommand('insertHTML', false, '&nbsp;&nbsp;&nbsp;&nbsp;');
    }}
}});
 
window.addEventListener('load', function() {{ sync(); }});
</script>
</body></html>"""
 
    components.html(html, height=editor_h + 76, scrolling=False)
 
    # ── Acciones del informe ──────────────────────────────────
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    ca, cb, cc, cd = st.columns(4)
 
    with ca:
        if st.button("Optimizar", key="btn_opt", use_container_width=True,
                     help="Mejora la Impresion Diagnostica"):
            if rep and api_key:
                ec = estilo_ctx()
                with st.spinner("Optimizando..."):
                    try:
                        cl = get_client(); mid = MODELS[st.session_state.modelo]["id"]
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":
                                f"""Radiologo subespecialista. Mejora UNICAMENTE la IMPRESION DIAGNOSTICA.
REGLAS: cero asteriscos, sin markdown, titulos MAYUSCULAS, vinetas con •.
Devuelve informe completo. No cambies TECNICA ni HALLAZGOS. Prosa en hallazgos.
Criterios: jerarquia, grado+implicacion en cada •, ultimo •=manejo.
{"Estilo: " + ec if ec else ""}
INFORME:\n{rep}"""}],
                            temperature=0.15, max_tokens=3000
                        )
                        st.session_state.reporte = limpiar(r.choices[0].message.content)
                        st.rerun()
                    except Exception as e: st.error(str(e))
 
    with cb:
        if st.button("Mentor", key="btn_mentor", use_container_width=True,
                     help="Analisis editorial del informe"):
            if rep and api_key:
                with st.spinner("Analizando..."):
                    try:
                        cl = get_client(); mid = MODELS[st.session_state.modelo]["id"]
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":
                                f"""Mentor de redaccion radiologica de elite. Analiza este informe.
Sin asteriscos ni markdown. Hasta 4 puntos especificos:
Para cada punto: QUE · POR QUE · COMO MEJORAR (con reescritura).
Detecta: prosa debil, listas donde deberia haber narrativa,
clasificaciones ausentes, hedge words, conclusiones no accionables.
Final: NIVEL [Basico/Residente/Fellow/Subespecialista/Publicable] · PUNTUACION [X/10]
INFORME:\n{rep}"""}],
                            temperature=0.2, max_tokens=1800
                        )
                        st.session_state.mentor_fb = r.choices[0].message.content
                        st.rerun()
                    except Exception as e: st.error(str(e))
 
    with cc:
        if st.button("Aprender", key="btn_learn", use_container_width=True,
                     help="Guardar este informe como ejemplo de mi estilo"):
            if rep:
                st.session_state.estilo.append({"r": rep})
                if len(st.session_state.estilo) > 10:
                    st.session_state.estilo = st.session_state.estilo[-10:]
                st.success(f"Guardado ({len(st.session_state.estilo)})")
 
    with cd:
        if rep:
            st.download_button(
                "Exportar .docx", data=generar_docx(rep),
                file_name="AURA_Informe.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, key="btn_dl"
            )
 
    # ── Sugerencias de redaccion basadas en literatura ────────
    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
    if st.button("Sugerencias de redaccion", key="btn_sug", use_container_width=True,
                 help="Sugerencias basadas en conceptos operativos y literatura"):
        if rep and api_key:
            dic = st.session_state.diccionario
            dic_ctx = "\n".join(f"- {k}: {v}" for k,v in dic.items()) if dic else "No definido."
            with st.spinner("Generando sugerencias..."):
                try:
                    cl = get_client(); mid = MODELS[st.session_state.modelo]["id"]
                    r = cl.chat.completions.create(
                        model=mid,
                        messages=[{"role":"user","content":
                            f"""Experto en redaccion radiologica academica.
Analiza este informe y devuelve EXACTAMENTE este JSON (sin texto fuera del JSON):
{{
  "sugerencias": [
    {{
      "tipo": "Precision semantica|Narrativa|Clasificacion|Correlacion clinica|Terminologia",
      "frase": "fragmento del informe (max 12 palabras)",
      "problema": "por que es suboptimo (1-2 lineas)",
      "mejora": "reescritura sugerida (2-3 lineas)",
      "concepto": "principio de redaccion radiologica",
      "referencia": "Autor, Revista, Ano o estandar ACR/RSNA/ESR"
    }}
  ]
}}
Genera 3 sugerencias. Solo JSON, sin markdown.
Diccionario: {dic_ctx}
INFORME:\n{rep}"""}],
                        temperature=0.2, max_tokens=1200
                    )
                    raw = r.choices[0].message.content.strip()
                    raw = re.sub(r'^```[a-z]*\s*', '', raw)
                    raw = re.sub(r'\s*```$', '', raw)
                    st.session_state.sugerencias = json.loads(raw).get("sugerencias", [])
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al parsear sugerencias: {e}")
 
    # Mostrar sugerencias
    if st.session_state.sugerencias:
        t = TH()
        st.markdown(f'<span style="font-size:9px;font-weight:600;letter-spacing:.14em;'
                    f'text-transform:uppercase;color:{t["mut"]}">Sugerencias de redaccion</span>',
                    unsafe_allow_html=True)
        for sug in st.session_state.sugerencias:
            st.markdown(f"""
            <div style="background:{t['surf']};border:1px solid {t['acc']}22;
                border-left:3px solid {t['acc']};border-radius:0 8px 8px 0;
                padding:11px 13px;margin:6px 0;font-size:11.5px;line-height:1.65">
                <div style="font-size:9px;font-weight:600;letter-spacing:.12em;
                    text-transform:uppercase;color:{t['acc']};margin-bottom:4px">
                    {sug.get('tipo','Sugerencia')}
                </div>
                <div style="color:{t['mut']};margin-bottom:5px">
                    <em>"{sug.get('frase','')}"</em>
                </div>
                <div style="color:{t['txt']};margin-bottom:3px">
                    <b style="color:{t['acc2']}">Problema:</b> {sug.get('problema','')}
                </div>
                <div style="color:{t['txt']};margin-bottom:4px">
                    <b style="color:{t['acc2']}">Mejora:</b> {sug.get('mejora','')}
                </div>
                <div style="font-size:10px;color:{t['mut']}">
                    {sug.get('concepto','')}
                    {'&nbsp;·&nbsp;' + sug.get('referencia','') if sug.get('referencia') else ''}
                </div>
            </div>""", unsafe_allow_html=True)
        if st.button("Cerrar sugerencias", key="btn_close_sug"):
            st.session_state.sugerencias = []; st.rerun()
 
    # Mostrar feedback mentor
    if st.session_state.mentor_fb:
        t = TH()
        st.markdown(f'<span style="font-size:9px;font-weight:600;letter-spacing:.14em;'
                    f'text-transform:uppercase;color:{t["mut"]}">Analisis del mentor</span>',
                    unsafe_allow_html=True)
        st.markdown(f"""
        <div style="background:{t['surf']};border:1px solid {t['border']};
            border-left:3px solid {t['acc']};border-radius:0 8px 8px 0;
            padding:13px;font-size:11.5px;line-height:1.72;
            color:{t['txt']};white-space:pre-wrap;margin-top:6px">
            {st.session_state.mentor_fb}
        </div>""", unsafe_allow_html=True)
        if st.button("Cerrar analisis", key="btn_close_mentor"):
            st.session_state.mentor_fb = ""; st.rerun()
