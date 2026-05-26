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
    page_title="AURA · Radiology Intelligence",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─────────────────────────────────────────────────────────────
# TEMAS
# ─────────────────────────────────────────────────────────────
TEMAS = {
    "Void":   {"base":"#000308","surface":"#030d18","panel":"#050f1c",
               "border":"rgba(0,180,255,0.10)","glow":"rgba(0,180,255,0.18)",
               "accent":"#00c8ff","accent2":"#0060a0","dim":"#1e5070",
               "text":"#a8d8f0","ghost":"#0a2840","scan":"rgba(0,180,255,0.025)",
               "sec_a":"rgba(0,200,255,0.08)","sec_b":"rgba(0,200,255,0.05)",
               "sec_c":"rgba(0,200,255,0.04)","sec_d":"rgba(0,200,255,0.06)"},
    "Plasma": {"base":"#04000a","surface":"#0a0118","panel":"#0d0120",
               "border":"rgba(180,60,255,0.12)","glow":"rgba(160,60,255,0.20)",
               "accent":"#b060ff","accent2":"#6020a0","dim":"#401880",
               "text":"#d0b0f8","ghost":"#180840","scan":"rgba(160,0,255,0.025)",
               "sec_a":"rgba(180,60,255,0.08)","sec_b":"rgba(180,60,255,0.05)",
               "sec_c":"rgba(180,60,255,0.04)","sec_d":"rgba(180,60,255,0.06)"},
    "Aurora": {"base":"#000a08","surface":"#010f10","panel":"#031410",
               "border":"rgba(0,210,150,0.10)","glow":"rgba(0,210,150,0.16)",
               "accent":"#00e8b0","accent2":"#007850","dim":"#0e5040",
               "text":"#90e8d0","ghost":"#052820","scan":"rgba(0,210,150,0.022)",
               "sec_a":"rgba(0,220,160,0.08)","sec_b":"rgba(0,220,160,0.05)",
               "sec_c":"rgba(0,220,160,0.04)","sec_d":"rgba(0,220,160,0.06)"},
    "Solar":  {"base":"#080400","surface":"#100800","panel":"#160a00",
               "border":"rgba(255,170,20,0.10)","glow":"rgba(255,160,0,0.16)",
               "accent":"#ffb030","accent2":"#905010","dim":"#603010",
               "text":"#f0d090","ghost":"#2a1400","scan":"rgba(255,160,0,0.022)",
               "sec_a":"rgba(255,170,20,0.08)","sec_b":"rgba(255,170,20,0.05)",
               "sec_c":"rgba(255,170,20,0.04)","sec_d":"rgba(255,170,20,0.06)"},
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

REGLAS = """PROTOCOLO CLÍNICO AURA:
· PROHIBIDO: "cambios degenerativos", "cambios crónicos" sin sustrato morfológico.
· USAR: descriptores morfológicos — osteofitos marginales, esclerosis subcondral, pinzamiento articular de X mm.
· TABLAS: solo si la plantilla contiene [TABLA]. Sin plantilla → cero tablas.
· CLASIFICACIONES: únicamente las respaldadas por los hallazgos. Especificar criterio morfológico.
· IMPRESIÓN: morfológicamente precisa. Lenguaje sugerente: "se sugiere correlación clínica"."""

SECCIONES_DEF = ["INDICACIÓN","TÉCNICA","HALLAZGOS","IMPRESIÓN DIAGNÓSTICA"]
SECCIONES_COLORES = {
    "INDICACIÓN":           ("sec_a","#60b8e0"),
    "TÉCNICA":              ("sec_b","#8090c0"),
    "HALLAZGOS":            ("sec_c","#60d4b0"),
    "IMPRESIÓN DIAGNÓSTICA":("sec_d","#e090b0"),
}

# ─────────────────────────────────────────────────────────────
# ESTADO
# ─────────────────────────────────────────────────────────────
for k,v in {
    "tema":"Void","dictado":"","reporte_secciones":{},"reporte_texto":"",
    "defs":"","editor_h":520,"modo":"voz","plantilla_txt":"","tiene_tabla":False,
    "modelo":"DeepSeek Chat","historial":[],"audio_id":None,"grabando":False,
    "nueva_sesion":True,
}.items():
    if k not in st.session_state: st.session_state[k]=v

# ─────────────────────────────────────────────────────────────
# HELPERS
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
    """Divide el texto del informe en secciones por título en mayúsculas."""
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
    h='<table style="border-collapse:collapse;width:100%;margin:8px 0;font-size:12px">'
    for i,row in enumerate(rows):
        cols=[c.strip() for c in row.strip("|").split("|")]
        tag="th" if i==0 else "td"
        h+="<tr>"+"".join(f"<{tag} style='border:1px solid #ddd;padding:4px 9px'>{c}</{tag}>" for c in cols)+"</tr>"
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
        try:
            import speech_recognition as sr; audio.seek(0)
            r=sr.Recognizer()
            with sr.AudioFile(audio) as src: return r.recognize_google(r.record(src),language="es-MX")
        except: return ""

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
# CSS HOLOGRÁFICO — Layout 3 columnas tipo Eden
# ─────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600&family=JetBrains+Mono:wght@300;400;500&display=swap');

html,body,.stApp{{background:{T['base']} !important;}}
.block-container{{padding:0 !important;max-width:100% !important;}}
header,footer,[data-testid="stToolbar"]{{display:none !important;}}
*{{font-family:'JetBrains Mono',monospace !important;}}

/* Scanlines globales */
.stApp::before{{
    content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
    background:repeating-linear-gradient(0deg,{T['scan']} 0px,transparent 1px,transparent 4px);
}}

/* ── TOPBAR ── */
.aura-top{{
    position:sticky;top:0;z-index:9999;
    height:46px;background:{T['base']};
    border-bottom:1px solid {T['border']};
    display:flex;align-items:center;padding:0 20px;gap:14px;
}}
.aura-logo{{
    font-family:'Space Grotesk',sans-serif !important;
    font-weight:300;font-size:16px;letter-spacing:.4em;
    color:{T['accent']};display:flex;align-items:center;gap:9px;
}}
.a-dot{{width:6px;height:6px;border-radius:50%;background:{T['accent']};
    box-shadow:0 0 8px {T['accent']},0 0 18px {T['accent']};
    animation:ap 2.5s ease-in-out infinite;}}
@keyframes ap{{0%,100%{{opacity:1;transform:scale(1);}}50%{{opacity:.35;transform:scale(.65);}}}}
.a-sep{{width:1px;height:16px;background:{T['border']};}}
.a-meta{{font-size:9px;letter-spacing:.18em;color:{T['dim']};}}
.a-right{{margin-left:auto;display:flex;align-items:center;gap:12px;}}
.a-online{{display:flex;align-items:center;gap:5px;font-size:8.5px;letter-spacing:.15em;color:{T['dim']};}}
.a-dot-sm{{width:4px;height:4px;border-radius:50%;background:{T['accent']};box-shadow:0 0 4px {T['accent']};}}

/* ── SIDEBAR IZQUIERDO ── */
.stApp > div > div > div > section[data-testid="stSidebar"]{{display:none !important;}}

/* Columna izquierda (sidebar custom) */
[data-testid="column"]:nth-child(1){{
    background:{T['surface']} !important;
    border-right:1px solid {T['border']} !important;
    min-height:calc(100vh - 46px);
    padding:0 !important;
}}
/* Columna central */
[data-testid="column"]:nth-child(2){{
    background:{T['base']} !important;
    border-right:1px solid {T['border']} !important;
    min-height:calc(100vh - 46px);
    padding:0 !important;
}}
/* Columna derecha */
[data-testid="column"]:nth-child(3){{
    background:{T['panel']} !important;
    min-height:calc(100vh - 46px);
    padding:0 !important;
}}

/* ── SIDEBAR ITEMS ── */
.sb-title{{
    font-size:8px;letter-spacing:.25em;color:{T['dim']};
    text-transform:uppercase;padding:16px 14px 8px;display:block;
}}
.sb-new{{
    margin:0 10px 14px;
    background:transparent;border:1px solid {T['accent']};
    color:{T['accent']};font-size:9px;letter-spacing:.2em;
    text-transform:uppercase;border-radius:2px;
    padding:7px 0;text-align:center;cursor:pointer;
    box-shadow:0 0 10px {T['glow']};
    transition:all .2s;display:block;width:calc(100% - 20px);
}}
.sb-new:hover{{background:{T['border']};box-shadow:0 0 20px {T['glow']};}}

.hist-item{{
    display:flex;align-items:center;gap:8px;
    padding:7px 14px;cursor:pointer;
    border-left:2px solid transparent;
    transition:all .15s;
}}
.hist-item:hover{{background:{T['border']};border-left-color:{T['accent']};}}
.hist-item.active{{background:{T['border']};border-left-color:{T['accent']};}}
.hist-dot{{width:8px;height:8px;border-radius:1px;flex-shrink:0;}}
.hist-lbl{{font-size:9px;color:{T['text']};letter-spacing:.04em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
.hist-sub{{font-size:8px;color:{T['dim']};letter-spacing:.03em;}}

.sb-bottom{{
    position:absolute;bottom:0;left:0;right:0;
    border-top:1px solid {T['border']};padding:10px 14px;
}}
.sb-bottom-item{{
    font-size:8.5px;letter-spacing:.12em;color:{T['dim']};
    padding:4px 0;display:flex;align-items:center;gap:7px;cursor:pointer;
}}
.sb-bottom-item:hover{{color:{T['text']};}}

/* ── PANEL CENTRAL ── */
.rec-header{{
    border-bottom:1px solid {T['border']};
    padding:14px 18px;display:flex;align-items:center;gap:12px;
}}
.rec-title{{font-size:11px;letter-spacing:.15em;color:{T['text']};text-transform:uppercase;}}

.rec-btn{{
    width:36px;height:36px;border-radius:50%;
    background:transparent;border:2px solid {T['accent']};
    display:flex;align-items:center;justify-content:center;
    cursor:pointer;transition:all .2s;flex-shrink:0;
    box-shadow:0 0 12px {T['glow']};
}}
.rec-btn:hover{{box-shadow:0 0 24px {T['glow']};background:{T['border']};}}
.rec-btn.active{{background:{T['accent']};box-shadow:0 0 20px {T['glow']},0 0 40px {T['glow']};}}
.rec-icon{{width:12px;height:12px;border-radius:50%;background:{T['accent']};}}
.rec-icon.active{{background:#fff;}}

.waveform{{
    flex:1;height:28px;display:flex;align-items:center;gap:2px;overflow:hidden;
}}
.wv-bar{{width:2px;border-radius:1px;background:{T['border']};transition:height .1s;}}

.dictado-area{{
    padding:14px 18px;
}}
.dictado-box{{
    background:transparent;border:1px solid {T['border']};
    border-radius:2px;width:100%;resize:none;
    color:{T['text']};font-size:11px;font-family:'JetBrains Mono',monospace;
    line-height:1.65;padding:12px;outline:none;
    caret-color:{T['accent']};
    transition:border-color .2s,box-shadow .2s;
}}
.dictado-box:focus{{border-color:{T['accent2']};box-shadow:0 0 12px {T['glow']};}}
.dictado-box::placeholder{{color:{T['dim']};}}

.center-actions{{
    padding:12px 18px;display:flex;gap:8px;align-items:center;
    border-top:1px solid {T['border']};
}}
.c-btn{{
    background:transparent;border:1px solid {T['accent']};
    color:{T['accent']};font-size:9px;letter-spacing:.15em;
    text-transform:uppercase;padding:7px 16px;border-radius:2px;
    cursor:pointer;transition:all .2s;
    box-shadow:0 0 8px {T['glow']};
    font-family:'JetBrains Mono',monospace;
}}
.c-btn:hover{{background:{T['border']};box-shadow:0 0 18px {T['glow']};}}
.c-btn.sec{{border-color:{T['dim']};color:{T['dim']};box-shadow:none;}}
.c-btn.sec:hover{{border-color:{T['accent2']};color:{T['text']};}}

.cfg-panel{{padding:14px 18px;border-top:1px solid {T['border']};}}
.cfg-row{{display:flex;gap:8px;align-items:center;margin-bottom:8px;}}
.cfg-lbl{{font-size:8px;letter-spacing:.2em;color:{T['dim']};text-transform:uppercase;min-width:60px;}}

/* ── PANEL DERECHO ── */
.rpt-header{{
    border-bottom:1px solid {T['border']};
    padding:12px 16px;display:flex;align-items:center;gap:10px;
}}
.rpt-mod-sel{{
    background:transparent;border:1px solid {T['border']};
    color:{T['text']};font-size:10px;
    font-family:'JetBrains Mono',monospace;
    padding:4px 8px;border-radius:2px;outline:none;appearance:none;cursor:pointer;
}}
.send-btn{{
    margin-left:auto;background:transparent;
    border:1px solid {T['accent']};color:{T['accent']};
    font-size:8.5px;letter-spacing:.15em;text-transform:uppercase;
    padding:5px 12px;border-radius:2px;cursor:pointer;
    box-shadow:0 0 8px {T['glow']};transition:all .2s;
    font-family:'JetBrains Mono',monospace;display:flex;align-items:center;gap:5px;
}}
.send-btn:hover{{background:{T['border']};box-shadow:0 0 18px {T['glow']};}}

.sec-block{{
    border-bottom:1px solid {T['border']};
    overflow:hidden;transition:all .2s;
}}
.sec-hdr{{
    padding:11px 16px;display:flex;align-items:center;gap:10px;
    cursor:pointer;transition:background .15s;
}}
.sec-hdr:hover{{background:{T['border']};}}
.sec-hdr-dot{{width:3px;height:18px;border-radius:1px;flex-shrink:0;}}
.sec-hdr-title{{font-size:10px;letter-spacing:.12em;font-weight:500;flex:1;text-transform:uppercase;}}
.sec-hdr-arrow{{font-size:10px;color:{T['dim']};transition:transform .2s;}}
.sec-body{{padding:10px 16px 14px 29px;font-size:11.5px;line-height:1.75;}}
.sec-empty{{color:{T['dim']};font-size:10px;letter-spacing:.08em;font-style:italic;}}

.rpt-actions{{
    padding:10px 14px;border-top:1px solid {T['border']};
    display:flex;gap:6px;flex-wrap:wrap;align-items:center;
}}
.r-btn{{
    background:transparent;border:1px solid {T['dim']};
    color:{T['dim']};font-size:8px;letter-spacing:.15em;
    text-transform:uppercase;padding:5px 10px;border-radius:2px;
    cursor:pointer;transition:all .15s;font-family:'JetBrains Mono',monospace;
}}
.r-btn:hover{{border-color:{T['accent2']};color:{T['text']};}}
.r-btn.hi{{border-color:{T['accent']};color:{T['accent']};box-shadow:0 0 6px {T['glow']};}}
.prog-wrap{{margin-left:auto;display:flex;align-items:center;gap:6px;}}
.prog-line{{width:48px;height:1px;background:{T['ghost']};position:relative;}}
.prog-fill{{position:absolute;top:0;left:0;height:1px;background:{T['accent']};transition:width .4s;box-shadow:0 0 4px {T['accent']};}}
.prog-pct{{font-size:8px;color:{T['dim']};letter-spacing:.08em;}}

/* ── SELECTS / INPUTS BASE ── */
[data-testid="stSelectbox"]>div>div{{
    background:transparent !important;border:1px solid {T['border']} !important;
    border-radius:2px !important;color:{T['text']} !important;font-size:10px !important;
}}
.stTextArea textarea{{
    background:transparent !important;border:1px solid {T['border']} !important;
    border-radius:2px !important;color:{T['text']} !important;
    font-size:11px !important;line-height:1.6 !important;caret-color:{T['accent']} !important;
}}
.stTextArea textarea:focus{{border-color:{T['accent2']} !important;box-shadow:0 0 10px {T['glow']} !important;}}
.stTextArea textarea::placeholder{{color:{T['dim']} !important;}}
[data-testid="stTextInput"] input{{
    background:transparent !important;border:1px solid {T['border']} !important;
    border-radius:2px !important;color:{T['text']} !important;font-size:10px !important;
}}
[data-testid="stAudioInput"]{{
    background:transparent !important;border:1px solid {T['border']} !important;border-radius:2px !important;
}}
[data-testid="stFileUploader"]{{
    background:transparent !important;border:1px dashed {T['border']} !important;border-radius:2px !important;
}}
[data-testid="stFileUploader"] *{{color:{T['dim']} !important;font-size:9px !important;}}
.stButton>button{{
    background:transparent !important;border:1px solid {T['border']} !important;
    border-radius:2px !important;color:{T['dim']} !important;
    font-size:8.5px !important;letter-spacing:.14em !important;text-transform:uppercase !important;
}}
.stButton>button:hover{{border-color:{T['accent2']} !important;color:{T['text']} !important;}}
.stDownloadButton>button{{
    background:transparent !important;border:1px solid {T['accent']} !important;
    border-radius:2px !important;color:{T['accent']} !important;
    font-size:8.5px !important;letter-spacing:.14em !important;
    box-shadow:0 0 6px {T['glow']} !important;
}}
[data-testid="stExpander"]{{
    background:transparent !important;border:1px solid {T['border']} !important;border-radius:2px !important;
}}
[data-testid="stExpander"] summary{{color:{T['dim']} !important;font-size:9px !important;letter-spacing:.16em !important;}}

::-webkit-scrollbar{{width:2px;height:2px;}}
::-webkit-scrollbar-thumb{{background:{T['accent2']};border-radius:1px;}}
hr{{border:none;border-top:1px solid {T['border']} !important;margin:6px 0 !important;}}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# TOPBAR
# ─────────────────────────────────────────────────────────────
modelo_activo = st.session_state.modelo
st.markdown(f"""
<div class="aura-top">
  <div class="aura-logo"><div class="a-dot"></div>AURA</div>
  <div class="a-sep"></div>
  <span class="a-meta">Radiology Intelligence · v3.0</span>
  <div class="a-right">
    <div class="a-online"><div class="a-dot-sm"></div>{modelo_activo.upper()} · ONLINE</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# LAYOUT 3 COLUMNAS
# ─────────────────────────────────────────────────────────────
col_sb, col_rec, col_rpt = st.columns([0.72, 1.2, 1.8], gap="small")

# ══════════════════════════════════════════════════════════
# SIDEBAR — Historial + config
# ══════════════════════════════════════════════════════════
HIST_COLORS = ["#00c8ff","#b060ff","#00e8b0","#ffb030","#e07070","#70b0e0"]

with col_sb:
    st.markdown("<div style='height:0px'></div>", unsafe_allow_html=True)

    # New report button
    if st.button("⊕  NUEVO INFORME", use_container_width=True):
        st.session_state.dictado=""
        st.session_state.reporte_texto=""
        st.session_state.reporte_secciones={}
        st.session_state.defs=""
        st.session_state.nueva_sesion=True
        st.rerun()

    # Historial
    st.markdown(f'<span class="sb-title">HISTORIAL</span>', unsafe_allow_html=True)

    if not st.session_state.historial:
        st.markdown(f'<span style="font-size:9px;color:{T["dim"]};padding:0 14px;display:block">Sin informes aún</span>',
                    unsafe_allow_html=True)
    else:
        for i, entry in enumerate(st.session_state.historial):
            color = HIST_COLORS[i % len(HIST_COLORS)]
            label = f"{entry['m'][:2].upper()} · {entry['r']}"
            sub   = entry['m'][:12]
            active_cls = "active" if st.session_state.reporte_texto==entry['t'] else ""
            st.markdown(f"""
            <div class="hist-item {active_cls}">
              <div class="hist-dot" style="background:{color}"></div>
              <div>
                <div class="hist-lbl">{label}</div>
                <div class="hist-sub">{sub}</div>
              </div>
            </div>""", unsafe_allow_html=True)
            if st.button(f"Cargar #{i+1}", key=f"h{i}", use_container_width=True):
                st.session_state.reporte_texto = entry['t']
                st.session_state.reporte_secciones = entry['s']
                st.rerun()

    st.markdown("<hr>", unsafe_allow_html=True)

    # Config compacta
    with st.expander("▸  CONFIGURACIÓN", expanded=False):
        st.markdown(f'<span style="font-size:8px;letter-spacing:.2em;color:{T["dim"]}">MODELO IA</span>',
                    unsafe_allow_html=True)
        m_sel = st.selectbox("modelo", list(MODELOS.keys()),
                             index=list(MODELOS.keys()).index(st.session_state.modelo),
                             label_visibility="collapsed")
        if m_sel != st.session_state.modelo:
            st.session_state.modelo = m_sel; st.rerun()

        st.markdown(f'<span style="font-size:8px;letter-spacing:.2em;color:{T["dim"]}">TEMA</span>',
                    unsafe_allow_html=True)
        for nombre in TEMAS:
            a = nombre==st.session_state.tema
            if st.button(f"{'▶ ' if a else '  '}{nombre.upper()}", key=f"tm_{nombre}", use_container_width=True):
                st.session_state.tema=nombre; st.rerun()

        st.markdown(f'<span style="font-size:8px;letter-spacing:.2em;color:{T["dim"]}">ALTURA EDITOR</span>',
                    unsafe_allow_html=True)
        nh = st.slider("h", 280, 1000, st.session_state.editor_h, 40, label_visibility="collapsed")
        if nh!=st.session_state.editor_h: st.session_state.editor_h=nh; st.rerun()

    if not api_key:
        with st.expander("▸  API KEY", expanded=True):
            api_key = st.text_input("k", type="password", label_visibility="collapsed",
                                    placeholder="sk- ···")

# ══════════════════════════════════════════════════════════
# PANEL CENTRAL — Grabación / Dictado
# ══════════════════════════════════════════════════════════
with col_rec:
    st.markdown("<div style='height:0px'></div>", unsafe_allow_html=True)

    # Subpanel de configuración de estudio
    with st.expander("▸  CONFIGURACIÓN DE ESTUDIO", expanded=True):
        st.markdown(f'<span style="font-size:8px;letter-spacing:.2em;color:{T["dim"]}">MODALIDAD</span>',
                    unsafe_allow_html=True)
        modalidad = st.selectbox("M", MODALIDADES, label_visibility="collapsed")
        st.markdown(f'<span style="font-size:8px;letter-spacing:.2em;color:{T["dim"]}">REGIÓN</span>',
                    unsafe_allow_html=True)
        region = st.selectbox("R", REGIONES, label_visibility="collapsed")

    # Tabs voz / texto
    tab_voz, tab_txt, tab_plantilla = st.tabs(["VOZ", "TEXTO", "PLANTILLA"])

    with tab_voz:
        # Componente de audio de Streamlit
        audio = st.audio_input("Grabación", label_visibility="collapsed")
        if audio:
            audio_id = hash(audio.read()); audio.seek(0)
            if audio_id != st.session_state.audio_id:
                if api_key:
                    with st.spinner(""):
                        txt = transcribir(audio, api_key, st.session_state.modelo)
                    if txt:
                        st.session_state.dictado += (" "+txt).strip()
                        st.session_state.audio_id = audio_id
                        st.rerun()
                else:
                    st.warning("API Key requerida para transcripción.")

    with tab_txt:
        st.markdown(f'<span style="font-size:8px;letter-spacing:.2em;color:{T["dim"]}">HALLAZGOS / DICTADO MANUAL</span>',
                    unsafe_allow_html=True)

    with tab_plantilla:
        f_up = st.file_uploader("Plantilla .docx", type=["docx"], label_visibility="collapsed")
        if f_up:
            st.session_state.plantilla_txt, st.session_state.tiene_tabla = leer_plantilla(f_up)
            icono = "◈ CON TABLAS" if st.session_state.tiene_tabla else "◇ CARGADA"
            st.markdown(f'<span style="font-size:9px;color:{T["accent"]}">{icono}</span>', unsafe_allow_html=True)
        st.markdown(f'<span style="font-size:8px;letter-spacing:.2em;color:{T["dim"]}">DIRECTRICES</span>',
                    unsafe_allow_html=True)
        instrucciones = st.text_area("dir", height=50, label_visibility="collapsed",
                                     value="Lenguaje médico experto. Sin asteriscos. Solo clasificaciones respaldadas.")

    # Área de dictado compartida
    st.markdown(f'<span style="font-size:8px;letter-spacing:.2em;color:{T["dim"]};padding:4px 0;display:block">SEÑAL TRANSCRITA / ENTRADA</span>',
                unsafe_allow_html=True)
    dictado = st.text_area(
        "d", value=st.session_state.dictado, height=st.session_state.editor_h//3,
        label_visibility="collapsed",
        placeholder="Dicta o escribe hallazgos...\n\nEj: Desgarro horizontal menisco medial Stoller III, extrusión 3 mm. Osteofitos marginales tibiofemorales mediales.",
        key="dictado_ta"
    )
    if dictado != st.session_state.dictado:
        st.session_state.dictado = dictado

    # Botones de acción central
    ba, bb, bc = st.columns([1.4,1,1])
    with ba:
        generar = st.button("◈  GENERAR INFORME", use_container_width=True)
    with bb:
        if st.button("PURGAR", use_container_width=True):
            st.session_state.dictado=""; st.session_state.audio_id=None; st.rerun()
    with bc:
        if st.button("LIMPIAR", use_container_width=True):
            st.session_state.reporte_texto=""; st.session_state.reporte_secciones={}; st.rerun()

# ─────────────────────────────────────────────────────────────
# PROCESAMIENTO IA
# ─────────────────────────────────────────────────────────────
if generar:
    if not api_key:
        st.warning("API Key requerida.")
    elif not st.session_state.dictado.strip():
        st.warning("Ingresa dictado o hallazgos.")
    else:
        cl = client_ia(api_key, st.session_state.modelo)
        mid = MODELOS[st.session_state.modelo]["id"]
        tiene_tabla = st.session_state.tiene_tabla
        plantilla = st.session_state.plantilla_txt
        tabla_i = ("La plantilla contiene tablas [TABLA]. Complétalas en Markdown."
                   if tiene_tabla else "PROHIBIDO generar tablas. No hay plantilla con tablas.")
        prompt_sys = f"""Eres AURA, sistema de inteligencia radiológica de alta precisión.
Redacta un informe de {modalidad} — región: {region}.

{REGLAS}
TABLAS: {tabla_i}

PLANTILLA:
{plantilla if plantilla else "INDICACIÓN\\nTÉCNICA\\nHALLAZGOS\\nIMPRESIÓN DIAGNÓSTICA"}

DIRECTRICES: {instrucciones}

FORMATO:
- Títulos de sección en MAYÚSCULAS.
- Usa • para viñetas en la impresión.
- Sin asteriscos Markdown.
- Medidas, grados y localización anatómica específicos."""
        with st.spinner(""):
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
# PANEL DERECHO — Informe por secciones
# ══════════════════════════════════════════════════════════
with col_rpt:
    st.markdown("<div style='height:0px'></div>", unsafe_allow_html=True)

    secs = st.session_state.reporte_secciones

    # Header del panel de reporte
    c_hdr1, c_hdr2 = st.columns([1.6,1])
    with c_hdr1:
        st.markdown(f'<span style="font-size:9px;letter-spacing:.2em;color:{T["dim"]}">INFORME ESTRUCTURADO</span>',
                    unsafe_allow_html=True)
    with c_hdr2:
        if st.session_state.reporte_texto:
            docx_b = generar_docx(st.session_state.reporte_texto, modalidad, region)
            fname = f"AURA_{region.replace(' ','_').replace('/','_')}.docx"
            st.download_button("↓ EXPORTAR .DOCX", data=docx_b,
                               file_name=fname,
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True)

    # Editor completo (componente HTML) si hay informe
    if st.session_state.reporte_texto:
        contenido_html = texto_a_html(st.session_state.reporte_texto)
        eH = st.session_state.editor_h
        fH = eH + 92

        html_ed = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@300;400;500&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
:root{{--acc:{T['accent']};--acc2:{T['accent2']};--brd:{T['border']};--glow:{T['glow']};
      --txt:{T['text']};--dim:{T['dim']};--ghost:{T['ghost']};--base:{T['base']};}}
html,body{{height:{fH}px;overflow:hidden;display:flex;flex-direction:column;
  background:var(--base);font-family:'JetBrains Mono',monospace;}}
body::before{{content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
  background:repeating-linear-gradient(0deg,{T['scan']} 0px,transparent 1px,transparent 4px);}}

/* toolbar */
.tb{{flex-shrink:0;z-index:10;background:var(--base);border-bottom:1px solid var(--brd);
    padding:4px 10px;display:flex;align-items:center;gap:3px;flex-wrap:nowrap;overflow-x:auto;}}
.tg{{display:flex;align-items:center;gap:2px;padding-right:7px;border-right:1px solid var(--ghost);flex-shrink:0}}
.tg:last-child{{border-right:none}}
.tb-b{{background:none;border:1px solid transparent;color:var(--dim);font-size:10px;
    padding:3px 5px;border-radius:1px;cursor:pointer;transition:all .12s;min-width:22px;text-align:center;}}
.tb-b:hover{{border-color:var(--brd);color:var(--txt);}}
.tb-b.on{{border-color:var(--acc2);color:var(--acc);box-shadow:0 0 5px var(--glow);}}
.tb-s{{background:transparent;border:1px solid var(--brd);color:var(--dim);
    font-size:9px;font-family:'JetBrains Mono',monospace;
    padding:2px 4px;border-radius:1px;outline:none;appearance:none;cursor:pointer;}}
.cd{{width:11px;height:11px;border-radius:50%;cursor:pointer;
    border:1px solid transparent;transition:all .1s;flex-shrink:0;}}
.cd:hover,.cd.on{{border-color:var(--acc);box-shadow:0 0 4px var(--glow);}}
.tl{{font-size:8px;letter-spacing:.18em;color:var(--ghost);white-space:nowrap;}}

/* editor */
.ew{{flex:1;overflow-y:auto;padding:12px 14px;min-height:0;
    scrollbar-width:thin;scrollbar-color:var(--acc2) transparent;}}
.ew::-webkit-scrollbar{{width:2px;}}
.ew::-webkit-scrollbar-thumb{{background:var(--acc2);}}
.doc{{min-height:100%;padding:20px 26px;outline:none;border-radius:1px;
    font-size:12px;line-height:1.8;color:#1a1a1a;background:#fff;transition:background .2s,color .2s;}}
.doc b,.doc strong{{font-weight:600;}}
.doc li{{margin-left:16px;margin-bottom:2px;}}
.doc hr{{border:none;border-top:1px solid #e0e0e0;margin:8px 0;}}
.doc table{{border-collapse:collapse;width:100%;margin:8px 0;font-size:11px;}}
.doc td,.doc th{{border:1px solid #e0e0e0;padding:4px 8px;}}
.doc th{{background:#f8f8f8;font-weight:600;}}

/* action strip */
.as{{flex-shrink:0;z-index:10;background:var(--base);border-top:1px solid var(--brd);
    padding:5px 10px;display:flex;align-items:center;gap:5px;}}
.as::before{{content:'';position:absolute;left:0;right:0;top:0;height:1px;
    background:linear-gradient(90deg,transparent,var(--acc),transparent);opacity:.15;}}
.as-b{{background:transparent;border:1px solid var(--brd);color:var(--dim);
    font-size:8px;letter-spacing:.15em;text-transform:uppercase;
    padding:4px 9px;border-radius:1px;cursor:pointer;transition:all .12s;
    font-family:'JetBrains Mono',monospace;}}
.as-b:hover{{border-color:var(--acc2);color:var(--txt);}}
.pw{{margin-left:auto;display:flex;align-items:center;gap:5px;}}
.pl{{width:48px;height:1px;background:var(--ghost);position:relative;}}
.pf{{position:absolute;top:0;left:0;height:1px;background:var(--acc);
    transition:width .4s;box-shadow:0 0 3px var(--acc);}}
.pp{{font-size:8px;color:var(--dim);letter-spacing:.08em;}}
.wc{{font-size:8px;color:var(--ghost);}}
</style></head><body>

<div class="tb">
  <div class="tg">
    <select class="tb-s" id="fnt" onchange="applyFont(this.value)" style="width:78px">
      <option value="'JetBrains Mono',monospace" selected>JetBrains</option>
      <option value="Georgia,serif">Georgia</option>
      <option value="'Calibri',sans-serif">Calibri</option>
      <option value="Arial,sans-serif">Arial</option>
      <option value="'Times New Roman',serif">Times</option>
    </select>
    <select class="tb-s" id="sz" onchange="applySize(this.value)" style="width:36px">
      <option value="9">9</option><option value="10">10</option>
      <option value="11">11</option><option value="12" selected>12</option>
      <option value="13">13</option><option value="14">14</option><option value="16">16</option>
    </select>
  </div>
  <div class="tg">
    <button class="tb-b" id="bB" onclick="fmt('bold')"><b>B</b></button>
    <button class="tb-b" id="bI" onclick="fmt('italic')"><i>I</i></button>
    <button class="tb-b" id="bU" onclick="fmt('underline')"><u>U</u></button>
  </div>
  <div class="tg">
    <button class="tb-b" onclick="fmt('justifyLeft')" title="Izq"><i class="ti ti-align-left"></i></button>
    <button class="tb-b" onclick="fmt('justifyCenter')" title="Centro"><i class="ti ti-align-center"></i></button>
    <button class="tb-b" onclick="fmt('justifyRight')" title="Der"><i class="ti ti-align-right"></i></button>
    <button class="tb-b" onclick="fmt('justifyFull')" title="Just"><i class="ti ti-align-justified"></i></button>
  </div>
  <div class="tg">
    <button class="tb-b" onclick="fmt('insertUnorderedList')"><i class="ti ti-list"></i></button>
    <button class="tb-b" onclick="fmt('insertOrderedList')"><i class="ti ti-list-numbers"></i></button>
    <button class="tb-b" onclick="insHR()">—</button>
  </div>
  <div class="tg" style="gap:4px;align-items:center">
    <span class="tl">BG</span>
    <div class="cd on" style="background:#fff;border:1px solid #ddd" onclick="setBg(this,'#fff','#1a1a1a')"></div>
    <div class="cd" style="background:#f5f0e8" onclick="setBg(this,'#f5f0e8','#1a0e00')"></div>
    <div class="cd" style="background:#0a1018" onclick="setBg(this,'#0a1018','#c8e8f8')"></div>
    <div class="cd" style="background:#000" onclick="setBg(this,'#000','#00e8b0')"></div>
  </div>
  <div class="tg">
    <button class="tb-b" onclick="copyClean()" title="Copiar"><i class="ti ti-copy"></i></button>
    <button class="tb-b" onclick="printDoc()" title="PDF"><i class="ti ti-printer"></i></button>
  </div>
</div>

<div class="ew">
  <div class="doc" id="doc" contenteditable="true" spellcheck="false">{contenido_html}</div>
</div>

<div class="as">
  <span class="wc" id="wc">— palabras</span>
  <div class="pw">
    <div class="pl"><div class="pf" id="pf" style="width:0%"></div></div>
    <span class="pp" id="pp">0%</span>
  </div>
</div>

<script>
var doc=document.getElementById('doc');
doc.style.background='#fff';doc.style.color='#1a1a1a';
function fmt(c){{doc.focus();document.execCommand(c,false,null);upd();}}
function upd(){{['Bold','Italic','Underline'].forEach(function(c){{
  var b=document.getElementById('b'+c[0]);
  if(b)b.classList.toggle('on',document.queryCommandState(c.toLowerCase()));
}});}}
function applyFont(f){{doc.style.fontFamily=f;}}
function applySize(s){{doc.style.fontSize=s+'px';}}
function setBg(el,bg,col){{doc.style.background=bg;doc.style.color=col;
  document.querySelectorAll('.cd').forEach(function(d){{d.classList.remove('on');}});el.classList.add('on');}}
function insHR(){{doc.focus();document.execCommand('insertHTML',false,'<hr style="border:none;border-top:1px solid #e0e0e0;margin:10px 0"><br>');}}
function calcPct(){{var t=doc.innerText.toUpperCase();
  var f=['TÉCNICA','HALLAZGOS','IMPRESIÓN'].filter(function(s){{return t.includes(s);}}).length;
  var w=t.split(/\\s+/).filter(Boolean).length;
  return Math.min(100,Math.round((f/3)*60+Math.min(w/150,1)*40));}}
function updBar(){{var s=calcPct();
  document.getElementById('pf').style.width=s+'%';document.getElementById('pp').textContent=s+'%';
  var w=doc.innerText.trim().split(/\\s+/).filter(Boolean).length;
  document.getElementById('wc').textContent=w+' palabras';}}
doc.addEventListener('input',updBar);doc.addEventListener('keyup',upd);doc.addEventListener('mouseup',upd);
window.addEventListener('load',function(){{updBar();}});
function copyClean(){{var t=doc.innerText;
  if(navigator.clipboard&&navigator.clipboard.writeText){{
    navigator.clipboard.writeText(t).then(function(){{toast('COPIADO');}});
  }}else{{var ta=document.createElement('textarea');ta.value=t;ta.style.cssText='position:fixed;opacity:0';
    document.body.appendChild(ta);ta.select();document.execCommand('copy');document.body.removeChild(ta);toast('COPIADO');}}}}
function printDoc(){{var w=window.open('','_blank');
  w.document.write('<html><head><title>AURA · Informe</title><style>body{{font-family:Calibri,sans-serif;font-size:12pt;line-height:1.7;margin:2cm}}b{{font-weight:600}}table{{border-collapse:collapse;width:100%}}td,th{{border:1px solid #ccc;padding:4px 8px}}th{{background:#f0f0f0}}</style></head><body>');
  w.document.write(doc.innerHTML);w.document.write('</body></html>');w.document.close();setTimeout(function(){{w.print();}},400);}}
function toast(m){{var el=document.createElement('div');el.textContent=m;
  el.style.cssText='position:fixed;bottom:44px;left:50%;transform:translateX(-50%);background:transparent;color:var(--acc);border:1px solid var(--acc);padding:4px 12px;border-radius:1px;font-size:8px;letter-spacing:.2em;font-family:JetBrains Mono,monospace;z-index:9999;pointer-events:none;box-shadow:0 0 10px var(--glow)';
  document.body.appendChild(el);setTimeout(function(){{document.body.removeChild(el);}},1500);}}
</script></body></html>"""

        components.html(html_ed, height=fH, scrolling=False)

    else:
        # Estado vacío — secciones skeleton
        for sec in SECCIONES_DEF:
            col_key, color = SECCIONES_COLORES.get(sec, ("sec_a","#888"))
            sec_bg = T[col_key]
            st.markdown(f"""
            <div class="sec-block">
              <div class="sec-hdr">
                <div class="sec-hdr-dot" style="background:{color}"></div>
                <span class="sec-hdr-title" style="color:{color}">{sec}</span>
              </div>
              <div class="sec-body">
                <span class="sec-empty">· · · sin contenido · · ·</span>
              </div>
            </div>""", unsafe_allow_html=True)

    # Acciones IA debajo del editor
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    ra, rb = st.columns(2)

    with ra:
        if st.button("◈  OPTIMIZAR CONCLUSIÓN", use_container_width=True):
            if api_key and st.session_state.reporte_texto:
                cl = client_ia(api_key, st.session_state.modelo)
                mid = MODELOS[st.session_state.modelo]["id"]
                with st.spinner(""):
                    try:
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":f"""Eres AURA — optimizador diagnóstico.
Mejora ÚNICAMENTE el bloque IMPRESIÓN DIAGNÓSTICA.
{REGLAS}
· Morfológicamente preciso y clínicamente accionable.
· Solo clasificaciones con evidencia directa (especifica el criterio).
· Usa "•" para viñetas. Lenguaje sugerente para seguimiento.
· Devuelve el informe COMPLETO. Conserva TÉCNICA y HALLAZGOS intactos.
· Sin asteriscos. Títulos en MAYÚSCULAS.
REPORTE:
{st.session_state.reporte_texto}"""}],
                            temperature=0.2, max_tokens=2048
                        )
                        txt = r.choices[0].message.content
                        st.session_state.reporte_texto=txt
                        st.session_state.reporte_secciones=parsear_secciones(txt)
                        st.rerun()
                    except Exception as e: st.error(str(e))

    with rb:
        if st.button("◇  DEFINICIONES & REFS", use_container_width=True):
            if api_key and st.session_state.reporte_texto:
                cl = client_ia(api_key, st.session_state.modelo)
                mid = MODELOS[st.session_state.modelo]["id"]
                with st.spinner(""):
                    try:
                        r = cl.chat.completions.create(
                            model=mid,
                            messages=[{"role":"user","content":f"""Analiza el informe. Formato EXACTO.
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
· [Término]: [definición, 1-2 líneas]

CORRELACIÓN CLÍNICA
[2-3 líneas. Lenguaje sugerente.]

Sin asteriscos.
INFORME:
{st.session_state.reporte_texto}"""}],
                            temperature=0.15, max_tokens=2048
                        )
                        st.session_state.defs=r.choices[0].message.content
                        st.rerun()
                    except Exception as e: st.error(str(e))

    # Panel de definiciones
    if st.session_state.defs:
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        with st.expander("◈  DEFINICIONES · CLASIFICACIONES · REFERENCIAS", expanded=True):
            st.markdown(
                f'<div style="font-size:11px;line-height:1.5;color:{T["text"]};white-space:pre-wrap;'
                f'padding:10px;background:{T["glass"] if "glass" in T else "transparent"};'
                f'border:1px solid {T["border"]};border-radius:2px">{st.session_state.defs}</div>',
                unsafe_allow_html=True
            )
            if st.button("✕  CERRAR"):
                st.session_state.defs=""; st.rerun()
