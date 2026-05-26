import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io, re, os, tempfile
from openai import OpenAI

# ─────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Beam AI · Radiology Intelligence",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────────────────────
# TEMAS
# ─────────────────────────────────────────────────────────────
TEMAS = {
    "Void":   {"base":"#000308","surface":"#030d18","panel":"#050f1c",
               "border":"rgba(0,180,255,0.15)","glow":"rgba(0,180,255,0.25)",
               "accent":"#00c8ff","accent2":"#0060a0","dim":"#1e5070",
               "text":"#a8d8f0","ghost":"#0a2840","scan":"rgba(0,180,255,0.025)"},
    "Plasma": {"base":"#04000a","surface":"#0a0118","panel":"#0d0120",
               "border":"rgba(180,60,255,0.15)","glow":"rgba(160,60,255,0.25)",
               "accent":"#b060ff","accent2":"#6020a0","dim":"#401880",
               "text":"#d0b0f8","ghost":"#180840","scan":"rgba(160,0,255,0.025)"},
    "Aurora": {"base":"#000a08","surface":"#010f10","panel":"#031410",
               "border":"rgba(0,210,150,0.15)","glow":"rgba(0,210,150,0.25)",
               "accent":"#00e8b0","accent2":"#007850","dim":"#0e5040",
               "text":"#90e8d0","ghost":"#052820","scan":"rgba(0,210,150,0.022)"},
    "Solar":  {"base":"#080400","surface":"#100800","panel":"#160a00",
               "border":"rgba(255,170,20,0.15)","glow":"rgba(255,160,0,0.25)",
               "accent":"#ffb030","accent2":"#905010","dim":"#603010",
               "text":"#f0d090","ghost":"#2a1400","scan":"rgba(255,160,0,0.022)"},
}

MODALIDADES = ["Resonancia Magnética","Tomografía Computarizada","Radiografía","Ultrasonido","PET-CT"]
REGIONES = ["Rodilla","Columna lumbar","Columna cervical","Hombro","Cadera",
            "Tobillo / Pie","Muñeca / Mano","Codo","Cerebro","Columna dorsal",
            "Tórax","Abdomen / Pelvis","Mama","Tiroides","Hígado"]
MODELOS = {
    "DeepSeek Chat": {"url":"https://api.deepseek.com","id":"deepseek-chat"},
    "GPT-4o Mini":   {"url":None,"id":"gpt-4o-mini"},
    "GPT-4.1 Mini":  {"url":None,"id":"gpt-4.1-mini"},
}

REGLAS = """PROTOCOLO CLÍNICO BEAM AI:
· PROHIBIDO: "cambios degenerativos", "cambios crónicos" sin sustrato morfológico.
· USAR: descriptores morfológicos — osteofitos marginales, esclerosis subcondral, pinzamiento articular de X mm.
· TABLAS: solo si la plantilla contiene [TABLA]. Sin plantilla → cero tablas.
· CLASIFICACIONES: únicamente las respaldadas por los hallazgos. Especificar criterio morfológico.
· IMPRESIÓN: morfológicamente precisa. Lenguaje sugerente: "se sugiere correlación clínica"."""

SECCIONES_DEF = ["INDICACIÓN","TÉCNICA","HALLAZGOS","IMPRESIÓN DIAGNÓSTICA"]

# ─────────────────────────────────────────────────────────────
# ESTADO
# ─────────────────────────────────────────────────────────────
for k,v in {
    "tema":"Void","dictado":"","reporte_secciones":{},"reporte_texto":"",
    "defs":"","editor_h":600,"modo":"voz","plantilla_txt":"","tiene_tabla":False,
    "modelo":"DeepSeek Chat","historial":[],"audio_id":None,"grabando":False,
    "nueva_sesion":True,
}.items():
    if k not in st.session_state: st.session_state[k]=v

# ─────────────────────────────────────────────────────────────
# HELPERS (Sin cambios en lógica, solo retocados para Beam AI)
# ─────────────────────────────────────────────────────────────
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
    except: partes=[p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n".join(partes), n>0

def parsear_secciones(texto):
    secciones={}; current=None; buf=[]
    for line in texto.split("\n"):
        s=line.strip()
        es_titulo = s.isupper() and 4<len(s)<80 and not s.startswith("•") and not s.startswith("·")
        if es_titulo:
            if current and buf: secciones[current]=" ".join(buf).strip()
            current=s; buf=[]
        elif current:
            if s: buf.append(s)
    if current and buf: secciones[current]=" ".join(buf).strip()
    return secciones

def texto_a_html(texto):
    lines=[]; buf=[]; in_tbl=False
    for line in texto.split("\n"):
        s=line.strip()
        if not s:
            if in_tbl: lines.append(_tbl(buf)); buf=[]; in_tbl=False
            lines.append("<br>")
        elif re.match(r'^\|.+\|$',s):
            if all(c in '-| :' for c in s): continue
            in_tbl=True; buf.append(s)
        else:
            if in_tbl: lines.append(_tbl(buf)); buf=[]; in_tbl=False
            if s.isupper() and len(s)<70 and not s.startswith("•"):
                lines.append(f"<b>{s}</b><br>")
            elif s.startswith("•") or s.startswith("·"):
                lines.append(f"<li>{s[1:].strip()}</li>")
            else:
                s2=re.sub(r'\*\*(.+?)\*\*',r'<b>\1</b>',s)
                lines.append(f"{s2}<br>")
    if in_tbl and buf: lines.append(_tbl(buf))
    return "\n".join(lines)

def _tbl(rows):
    h='<table style="border-collapse:collapse;width:100%;margin:8px 0;font-size:12px;border:1px solid var(--brd)">'
    for i,row in enumerate(rows):
        cols=[c.strip() for c in row.strip("|").split("|")]
        tag="th" if i==0 else "td"
        h+="<tr>"+"".join(f"<{tag} style='border:1px solid var(--brd);padding:4px 9px'>{c}</{tag}>" for c in cols)+"</tr>"
    return h+"</table>"

def generar_docx(texto, modalidad="", region=""):
    doc=Document()
    sn=doc.styles["Normal"]; sn.font.name="Calibri"; sn.font.size=Pt(11)
    if modalidad:
        t=doc.add_heading(f"{modalidad.upper()} — {region.upper()}",level=1)
        t.alignment=WD_ALIGN_PARAGRAPH.CENTER
    in_tbl=False; rows=[]
    def flush():
        nonlocal rows
        if not rows: return
        cols=max(len(r) for r in rows)
        tbl=doc.add_table(rows=len(rows),cols=cols); tbl.style="Table Grid"
        for i,r in enumerate(rows):
            for j,c in enumerate(r):
                if j<cols:
                    cell=tbl.rows[i].cells[j]; cell.text=c
                    if i==0:
                        for run in cell.paragraphs[0].runs: run.bold=True
        rows=[]
    for line in texto.split("\n"):
        s=line.strip()
        if re.match(r'^\|.+\|$',s):
            if all(c in '-| :' for c in s): continue
            in_tbl=True; rows.append([c.strip() for c in s.strip("|").split("|")]); continue
        else:
            if in_tbl: flush(); in_tbl=False
        if not s: doc.add_paragraph(); continue
        if s.isupper() and len(s)<80 and not s.startswith("•"):
            h=doc.add_heading(s,level=2)
            if h.runs: h.runs[0].font.color.rgb=RGBColor(0x1a,0x3a,0x6a)
            continue
        if s.startswith("•") or s.startswith("·"):
            p=doc.add_paragraph(style="List Bullet"); _run(p,s[1:].strip()); continue
        p=doc.add_paragraph(); _run(p,s)
    if in_tbl: flush()
    bio=io.BytesIO(); doc.save(bio); return bio.getvalue()

def _run(p,text):
    pat=re.compile(r'(\*\*(.+?)\*\*|\*(.+?)\*)'); last=0
    for m in pat.finditer(text):
        if m.start()>last: p.add_run(text[last:m.start()])
        r=p.add_run(m.group(2) if m.group(0).startswith("**") else m.group(3))
        r.bold=m.group(0).startswith("**"); r.italic=not m.group(0).startswith("**")
        last=m.end()
    if last<len(text): p.add_run(text[last:])

def transcribir(audio, api_key, modelo):
    PROMPT="Dictado radiológico: Stoller, ICRS, LCA, Kellgren-Lawrence, STIR, menisco, condromalacia, osteofito."
    try:
        cfg=MODELOS[modelo]; client=OpenAI(api_key=api_key,base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)
        suf=".wav"
        if hasattr(audio,"name"):
            ext=os.path.splitext(audio.name)[-1].lower()
            if ext in [".mp3",".mp4",".m4a",".ogg",".webm",".flac"]: suf=ext
        with tempfile.NamedTemporaryFile(delete=False,suffix=suf) as tmp:
            tmp.write(audio.read()); tmp_path=tmp.name
        with open(tmp_path,"rb") as f:
            res=client.audio.transcriptions.create(model="whisper-1",file=f,language="es",prompt=PROMPT)
        os.unlink(tmp_path); return res.text.strip()
    except:
        return ""

def client_ia(api_key, modelo):
    cfg=MODELOS[modelo]
    return OpenAI(api_key=api_key,base_url=cfg["url"]) if cfg["url"] else OpenAI(api_key=api_key)

def guardar_historial(modalidad, region, texto, secciones):
    st.session_state.historial.insert(0,{"m":modalidad,"r":region,"t":texto,"s":secciones})
    if len(st.session_state.historial)>12: st.session_state.historial=st.session_state.historial[:12]

try: api_key=st.secrets["deepseek_key"]
except: api_key=os.environ.get("OPENAI_API_KEY","")

T=TEMAS[st.session_state.tema]

# ─────────────────────────────────────────────────────────────
# CSS MINIMALISTA & LIMPIO
# ─────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600&family=JetBrains+Mono:wght@300;400;500&display=swap');

html, body, .stApp {{ background:{T['base']} !important; color:{T['text']} !important; font-family:'JetBrains Mono',monospace !important; }}
header [data-testid="stToolbar"] {{ display:none !important; }}

/* Scanlines sutiles */
.stApp::before {{
    content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
    background:repeating-linear-gradient(0deg,{T['scan']} 0px,transparent 1px,transparent 4px);
}}

/* Barra Superior Minimalista */
.beam-top {{
    border-bottom: 1px solid {T['border']}; padding: 12px 20px;
    display:flex; align-items:center; gap: 15px; margin-bottom: 20px;
    background: {T['surface']}; border-radius: 4px;
}}
.beam-logo {{ font-family:'Space Grotesk',sans-serif; font-weight:400; font-size:18px; letter-spacing:0.3em; color:{T['accent']}; }}
.beam-status {{ margin-left:auto; font-size: 11px; color:{T['dim']}; letter-spacing:0.1em; display:flex; align-items:center; gap:6px; }}
.dot-pulse {{ width:8px; height:8px; background:{T['accent']}; border-radius:50%; box-shadow:0 0 10px {T['glow']}; animation: pulse 2s infinite; }}
@keyframes pulse {{ 0% {{ opacity: 1; transform:scale(1); }} 50% {{ opacity: 0.5; transform:scale(0.8); }} 100% {{ opacity: 1; transform:scale(1); }} }}

/* Estilizando el componente de Audio de Streamlit para que no desentone */
[data-testid="stAudioInput"] {{ 
    border: 1px solid {T['accent']} !important; 
    border-radius: 6px !important; 
    background: {T['panel']} !important;
    box-shadow: 0 0 15px {T['glow']} !important;
    transition: all 0.3s ease;
}}
[data-testid="stAudioInput"]:hover {{ border-color: {T['accent2']} !important; box-shadow: 0 0 25px {T['glow']} !important; }}

/* Botones y TextAreas generales */
.stTextArea textarea {{
    background: {T['surface']} !important; border: 1px solid {T['border']} !important;
    color: {T['text']} !important; font-family: 'JetBrains Mono', monospace !important; font-size: 13px !important;
    border-radius: 4px !important; line-height: 1.6 !important;
}}
.stTextArea textarea:focus {{ border-color: {T['accent']} !important; box-shadow: 0 0 10px {T['glow']} !important; }}
.stButton>button {{
    background: transparent !important; border: 1px solid {T['border']} !important;
    color: {T['text']} !important; font-size: 11px !important; letter-spacing: 0.1em !important;
    border-radius: 4px !important; transition: all 0.2s;
}}
.stButton>button:hover {{ border-color: {T['accent']} !important; box-shadow: 0 0 12px {T['glow']} !important; color: {T['accent']} !important; }}
.stDownloadButton>button {{ border-color: {T['accent']} !important; color: {T['accent']} !important; }}

/* Sidebar (Para aprovechar la función nativa de Streamlit de colapsar) */
[data-testid="stSidebar"] {{ background: {T['surface']} !important; border-right: 1px solid {T['border']} !important; }}
[data-testid="stSidebar"] * {{ color: {T['text']} !important; font-family:'JetBrains Mono',monospace !important; }}

</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SIDEBAR NATIVO (Permite expandir/colapsar horizontalmente)
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f'<div style="font-family:Space Grotesk; font-size:24px; font-weight:300; letter-spacing:0.2em; color:{T["accent"]}; text-align:center; margin-bottom:20px;">BEAM AI</div>', unsafe_allow_html=True)
    
    if st.button("⊕ NUEVO INFORME", use_container_width=True):
        st.session_state.dictado = ""
        st.session_state.reporte_texto = ""
        st.session_state.reporte_secciones = {}
        st.session_state.defs = ""
        st.rerun()
    
    st.markdown("<hr style='opacity:0.2;'>", unsafe_allow_html=True)
    st.markdown(f"<span style='font-size:10px; color:{T['dim']}; letter-spacing:0.2em;'>CONFIGURACIÓN</span>", unsafe_allow_html=True)
    
    m_sel = st.selectbox("Modelo IA", list(MODELOS.keys()), index=list(MODELOS.keys()).index(st.session_state.modelo))
    if m_sel != st.session_state.modelo: st.session_state.modelo = m_sel; st.rerun()
    
    tema_sel = st.selectbox("Tema Visual", list(TEMAS.keys()), index=list(TEMAS.keys()).index(st.session_state.tema))
    if tema_sel != st.session_state.tema: st.session_state.tema = tema_sel; st.rerun()
    
    if not api_key:
        api_key = st.text_input("API Key", type="password", placeholder="sk-...")

    st.markdown("<hr style='opacity:0.2;'>", unsafe_allow_html=True)
    st.markdown(f"<span style='font-size:10px; color:{T['dim']}; letter-spacing:0.2em;'>HISTORIAL</span>", unsafe_allow_html=True)
    
    if not st.session_state.historial:
        st.caption("Sin informes aún.")
    else:
        for i, entry in enumerate(st.session_state.historial):
            if st.button(f"#{i+1} · {entry['m'][:10]}... {entry['r']}", key=f"h{i}", use_container_width=True):
                st.session_state.reporte_texto = entry['t']
                st.session_state.reporte_secciones = entry['s']
                st.rerun()

# ─────────────────────────────────────────────────────────────
# MAIN CANVAS (Layout dividido 2 Columnas Limpias)
# ─────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="beam-top">
    <div class="beam-logo">BEAM AI <span style="font-size:10px; color:{T['dim']}">(v3.1)</span></div>
    <div class="beam-status"><div class="dot-pulse"></div> EN LÍNEA · {st.session_state.modelo.upper()}</div>
</div>
""", unsafe_allow_html=True)

col_input, col_output = st.columns([1, 1.4], gap="large")

# ══════════════════════════════════════════════════════════
# COLUMNA IZQUIERDA — Dictado y Parámetros
# ══════════════════════════════════════════════════════════
with col_input:
    # Parámetros base sin usar st.expander para evitar ruido visual
    st.markdown(f"<span style='font-size:12px; color:{T['accent']};'>◈ CONTEXTO CLÍNICO</span>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1: modalidad = st.selectbox("Modalidad", MODALIDADES)
    with c2: region = st.selectbox("Región", REGIONES)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    tab_voz, tab_txt, tab_plantilla = st.tabs(["VOZ", "TEXTO MANUAL", "PLANTILLA"])

    with tab_voz:
        audio = st.audio_input("Dictado", label_visibility="collapsed")
        if audio:
            audio_id = hash(audio.read()); audio.seek(0)
            if audio_id != st.session_state.audio_id:
                if api_key:
                    with st.spinner("Procesando señal acústica..."):
                        txt = transcribir(audio, api_key, st.session_state.modelo)
                    if txt:
                        st.session_state.dictado += (" " + txt).strip()
                        st.session_state.audio_id = audio_id
                        st.rerun()
                else:
                    st.warning("API Key requerida para transcripción.")

    with tab_plantilla:
        f_up = st.file_uploader("Subir Plantilla .docx", type=["docx"])
        if f_up:
            st.session_state.plantilla_txt, st.session_state.tiene_tabla = leer_plantilla(f_up)
            st.success("Plantilla cargada.")
        instrucciones = st.text_area("Directrices Extra", value="Lenguaje médico experto. Sin asteriscos. Solo clasificaciones respaldadas.", height=80)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f"<span style='font-size:10px; color:{T['dim']}; letter-spacing:0.1em;'>SEÑAL TRANSCRITA / ENTRADA:</span>", unsafe_allow_html=True)
    dictado = st.text_area("Dictado", value=st.session_state.dictado, height=200, label_visibility="collapsed", placeholder="Dicta o escribe hallazgos aquí...")
    if dictado != st.session_state.dictado:
        st.session_state.dictado = dictado

    c_btn1, c_btn2 = st.columns([2, 1])
    with c_btn1:
        generar = st.button("◈ GENERAR INFORME", use_container_width=True)
    with c_btn2:
        if st.button("LIMPIAR", use_container_width=True):
            st.session_state.dictado = ""; st.session_state.audio_id = None; st.rerun()

# ─────────────────────────────────────────────────────────────
# PROCESAMIENTO IA
# ─────────────────────────────────────────────────────────────
if generar:
    if not api_key: st.warning("API Key requerida.")
    elif not st.session_state.dictado.strip(): st.warning("Ingresa dictado o hallazgos.")
    else:
        cl = client_ia(api_key, st.session_state.modelo)
        mid = MODELOS[st.session_state.modelo]["id"]
        tiene_tabla = st.session_state.tiene_tabla
        plantilla = st.session_state.plantilla_txt
        tabla_i = ("La plantilla contiene tablas [TABLA]. Complétalas en Markdown." if tiene_tabla else "PROHIBIDO generar tablas. No hay plantilla con tablas.")
        prompt_sys = f"""Eres Beam AI, sistema de inteligencia radiológica de alta precisión.
Redacta un informe de {modalidad} — región: {region}.

{REGLAS}
TABLAS: {tabla_i}

PLANTILLA:
{plantilla if plantilla else "INDICACIÓN\\nTÉCNICA\\nHALLAZGOS\\nIMPRESIÓN DIAGNÓSTICA"}

DIRECTRICES: {instrucciones if 'instrucciones' in locals() else ''}

FORMATO:
- Títulos de sección en MAYÚSCULAS.
- Usa • para viñetas en la impresión.
- Sin asteriscos Markdown.
- Medidas, grados y localización anatómica específicos."""
        with st.spinner("Sintetizando informe..."):
            try:
                res = cl.chat.completions.create(
                    model=mid,
                    messages=[
                        {"role":"system","content":prompt_sys},
                        {"role":"user","content":f"DICTADO:\n{st.session_state.dictado}"}
                    ],
                    temperature=0.1, max_tokens=2048
                )
                txt = res.choices[0].message.content
                secs = parsear_secciones(txt)
                st.session_state.reporte_texto = txt
                st.session_state.reporte_secciones = secs
                guardar_historial(modalidad, region, txt, secs)
                st.rerun()
            except Exception as e:
                st.error(str(e))

# ══════════════════════════════════════════════════════════
# COLUMNA DERECHA — Editor HTML y Acciones
# ══════════════════════════════════════════════════════════
with col_output:
    c_out1, c_out2 = st.columns([2, 1])
    with c_out1: st.markdown(f"<span style='font-size:12px; color:{T['accent']};'>◈ VISOR DE INFORME</span>", unsafe_allow_html=True)
    with c_out2:
        if st.session_state.reporte_texto:
            docx_b = generar_docx(st.session_state.reporte_texto, modalidad, region)
            st.download_button("↓ EXPORTAR .DOCX", data=docx_b, file_name=f"BeamAI_{region.replace(' ','_')}.docx", use_container_width=True)

    if st.session_state.reporte_texto:
        contenido_html = texto_a_html(st.session_state.reporte_texto)
        eH = st.session_state.editor_h
        
        # HTML EDITOR: Ahora inicia en modo oscuro heredando las variables de color CSS.
        html_ed = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;500&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
:root{{
    --acc:{T['accent']}; --acc2:{T['accent2']}; --brd:{T['border']}; --glow:{T['glow']};
    --txt:{T['text']}; --dim:{T['dim']}; --surface:{T['surface']}; --base:{T['base']};
}}
html,body{{ height:{eH}px; overflow:hidden; display:flex; flex-direction:column;
  background:var(--base); font-family:'JetBrains Mono',monospace; }}

/* toolbar simplificada */
.tb{{ flex-shrink:0; background:var(--surface); border-bottom:1px solid var(--brd); padding:8px 12px; display:flex; align-items:center; gap:8px; border-radius:4px 4px 0 0; }}
.tb-b{{ background:none; border:1px solid transparent; color:var(--dim); padding:4px 8px; border-radius:3px; cursor:pointer; transition:all 0.2s; font-size:14px; }}
.tb-b:hover{{ border-color:var(--brd); color:var(--txt); }}
.tb-b.on{{ border-color:var(--acc2); color:var(--acc); box-shadow:0 0 5px var(--glow); }}

/* editor con modo oscuro default */
.ew{{ flex:1; overflow-y:auto; padding:0; }}
.doc{{ 
    min-height:100%; padding:25px 30px; outline:none; 
    font-size:13px; line-height:1.8; 
    background:var(--surface); color:var(--txt); 
    border-bottom-left-radius: 4px; border-bottom-right-radius: 4px;
    border: 1px solid var(--brd); border-top: none;
    transition: background 0.3s, color 0.3s;
}}
.doc b, .doc strong {{ color: var(--acc); font-weight:500; }}
.doc table {{ border-collapse:collapse; width:100%; margin:10px 0; border: 1px solid var(--brd); }}
.doc td, .doc th {{ border: 1px solid var(--brd); padding:6px 10px; }}
.doc th {{ background: var(--base); }}

/* scrollbar minimalista */
::-webkit-scrollbar {{ width: 6px; }}
::-webkit-scrollbar-thumb {{ background: var(--acc2); border-radius: 3px; }}
::-webkit-scrollbar-track {{ background: transparent; }}
</style></head><body>

<div class="tb">
  <button class="tb-b" id="bB" onclick="fmt('bold')" title="Negrita"><i class="ti ti-bold"></i></button>
  <button class="tb-b" id="bI" onclick="fmt('italic')" title="Cursiva"><i class="ti ti-italic"></i></button>
  <button class="tb-b" id="bU" onclick="fmt('underline')" title="Subrayado"><i class="ti ti-underline"></i></button>
  <span style="border-left:1px solid var(--brd); height:16px; margin:0 5px;"></span>
  <button class="tb-b" onclick="fmt('insertUnorderedList')" title="Viñetas"><i class="ti ti-list"></i></button>
  <button class="tb-b" onclick="insHR()" title="Separador"><i class="ti ti-minus"></i></button>
  <span style="border-left:1px solid var(--brd); height:16px; margin:0 5px;"></span>
  <button class="tb-b" onclick="toggleTheme()" title="Cambiar Fondo Editor"><i class="ti ti-moon-stars"></i> / <i class="ti ti-sun"></i></button>
  <button class="tb-b" onclick="copyClean()" title="Copiar al Portapapeles" style="margin-left:auto;"><i class="ti ti-copy"></i></button>
</div>

<div class="ew">
  <div class="doc" id="doc" contenteditable="true" spellcheck="false">{contenido_html}</div>
</div>

<script>
var doc=document.getElementById('doc');
var isDark = true; 
// Inicializa en modo oscuro basado en el tema
doc.style.background='var(--surface)'; doc.style.color='var(--txt)';

function toggleTheme() {{
    isDark = !isDark;
    if(isDark) {{ doc.style.background='var(--surface)'; doc.style.color='var(--txt)'; }} 
    else {{ doc.style.background='#ffffff'; doc.style.color='#1a1a1a'; }}
}}
function fmt(c){{doc.focus();document.execCommand(c,false,null);upd();}}
function upd(){{['Bold','Italic','Underline'].forEach(function(c){{
  var b=document.getElementById('b'+c[0]);
  if(b)b.classList.toggle('on',document.queryCommandState(c.toLowerCase()));
}});}}
function insHR(){{doc.focus();document.execCommand('insertHTML',false,'<hr style="border:none;border-top:1px solid var(--brd);margin:12px 0"><br>');}}
doc.addEventListener('keyup',upd);doc.addEventListener('mouseup',upd);
function copyClean(){{var t=doc.innerText;
  if(navigator.clipboard&&navigator.clipboard.writeText){{
    navigator.clipboard.writeText(t).then(function(){{toast('TEXTO COPIADO');}});
  }}else{{var ta=document.createElement('textarea');ta.value=t;ta.style.cssText='position:fixed;opacity:0';
    document.body.appendChild(ta);ta.select();document.execCommand('copy');document.body.removeChild(ta);toast('TEXTO COPIADO');}}}}
function toast(m){{var el=document.createElement('div');el.textContent=m;
  el.style.cssText='position:fixed;bottom:20px;left:50%;transform:translateX(-50%);background:var(--surface);color:var(--acc);border:1px solid var(--acc);padding:6px 16px;border-radius:4px;font-size:11px;letter-spacing:0.1em;font-family:JetBrains Mono,monospace;z-index:9999;box-shadow:0 0 15px var(--glow)';
  document.body.appendChild(el);setTimeout(function(){{document.body.removeChild(el);}},2000);}}
</script></body></html>"""
        components.html(html_ed, height=eH, scrolling=False)
        
        # Botones de Acciones Rápidas
        st.markdown("<br>", unsafe_allow_html=True)
        c_act1, c_act2 = st.columns(2)
        with c_act1:
            if st.button("◈ OPTIMIZAR IMPRESIÓN", use_container_width=True):
                # (Misma lógica de optimización)
                pass # Aquí mantienes tu código original de optimizar
        with c_act2:
            if st.button("◇ DEFINICIONES & REFS", use_container_width=True):
                # (Misma lógica de definiciones)
                pass # Aquí mantienes tu código original de definiciones
    else:
        st.info("El visor se habilitará al procesar el primer informe.")
