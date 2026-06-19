import io
import json
import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from core.ai import RadiologyCopilot
from core.knowledge import KnowledgeBase
from services.storage import DraftStorage
from data.constants import SUGERENCIAS, CLASIFICACIONES, REGIONES

st.set_page_config(page_title="Beam AI", layout="wide", initial_sidebar_state="collapsed")

with open("ui/styles.css", "r", encoding="utf-8") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

for k, v in {
    "theme": "dark",
    "report": "",
    "findings": "",
    "modalidad": "Resonancia Magnética",
    "region": "Rodilla",
    "dictado_transcripcion": "",
    "last_saved": "",
    "clasif_activas": {},
    "definiciones": "",
    "uploaded_template": "",
}.items():
    st.session_state.setdefault(k, v)

storage = DraftStorage()
if not st.session_state.report:
    loaded = storage.load_draft()
    if loaded:
        st.session_state.report = loaded

kb = KnowledgeBase("data/knowledge_base.json")
copilot = RadiologyCopilot(
    api_key=st.secrets.get("deepseek_key", ""),
    knowledge_base=kb,
)

@st.cache_data(ttl=10)
def export_docx(text: str) -> bytes:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = 11
    for line in text.splitlines():
        s = line.strip()
        if not s:
            doc.add_paragraph()
        elif s.startswith("•"):
            doc.add_paragraph(s, style="List Bullet")
        else:
            doc.add_paragraph(s)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def leer_plantilla(file):
    doc = Document(file)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

def save_report(text: str):
    st.session_state.report = text
    storage.save_draft(text)

def render_chips(items):
    html = '<div style="display:flex;flex-wrap:wrap;gap:4px;margin-bottom:4px">'
    for item in items:
        html += f'<span style="font-size:10px;font-family:\'IBM Plex Mono\',monospace;color:var(--accent);background:var(--accent-bg);border:1px solid var(--accent-border);padding:2px 8px;border-radius:4px;cursor:pointer" onclick="navigator.clipboard.writeText(`{item}`)">{item}</span>'
    html += "</div>"
    return html

st.markdown(
    f"<script>document.documentElement.setAttribute('data-theme','{st.session_state.theme}')</script>",
    unsafe_allow_html=True,
)

c1, c2, c3, c4, c5, c6, c7 = st.columns([0.4, 0.8, 0.8, 1.0, 0.6, 0.5, 0.4])

with c1:
    st.markdown('<div class="brand"><span class="brand-dot"></span>BEAM AI</div>', unsafe_allow_html=True)

with c2:
    label = "Ocultar" if st.session_state.get("show_input", True) else "Mostrar"
    if st.button(f"⊞ Panel {label}", key="toggle_panel", use_container_width=True):
        st.session_state.show_input = not st.session_state.get("show_input", True)
        st.rerun()

with c3:
    theme_label = "☀️" if st.session_state.theme == "dark" else "🌙"
    if st.button(f"Tema {theme_label}", key="theme_btn", use_container_width=True):
        st.session_state.theme = "light" if st.session_state.theme == "dark" else "dark"

with c4:
    st.markdown(
        f'<div class="badge">DeepSeek · {"conectado" if copilot.is_available() else "sin API"}</div>',
        unsafe_allow_html=True,
    )

with c5:
    if st.button("Copiar", key="copy_btn", use_container_width=True):
        st.session_state.copy_trigger = True

with c6:
    st.download_button(
        "DOCX",
        data=export_docx(st.session_state.report),
        file_name="BeamAI_Informe.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key="export_btn",
    )

with c7:
    pass

if st.session_state.pop("copy_trigger", False):
    safe = json.dumps(st.session_state.report)
    components.html(f"<script>navigator.clipboard.writeText({safe})</script>", height=0)
    st.toast("Copiado al portapapeles")

if st.session_state.last_saved:
    st.markdown(
        f'<div style="text-align:right;padding:2px 20px;font-size:10px;color:var(--text-muted);border-bottom:1px solid var(--border);font-family:var(--font-mono)">{st.session_state.last_saved}</div>',
        unsafe_allow_html=True,
    )

left, right = st.columns([1.1, 2.6] if st.session_state.get("show_input", True) else [0, 1])

with left:
    if st.session_state.get("show_input", True):
        with st.expander("⊞ MODALIDAD & REGIÓN", expanded=True):
            st.markdown('<span class="section-label">Modalidad</span>', unsafe_allow_html=True)
            modalidad = st.selectbox("", list(SUGERENCIAS.keys()), key="modalidad", label_visibility="collapsed")
            st.markdown('<span class="section-label">Región anatómica</span>', unsafe_allow_html=True)
            region = st.selectbox("", REGIONES, key="region", label_visibility="collapsed")

        with st.expander("⊞ SUGERENCIAS POR MODALIDAD", expanded=True):
            st.markdown(render_chips(SUGERENCIAS[st.session_state.modalidad]), unsafe_allow_html=True)

        with st.expander("⊞ HALLAZGOS / DICTADO", expanded=True):
            findings = st.text_area(
                "",
                value=st.session_state.findings,
                placeholder="Escribe los hallazgos clínicos aquí...",
                key="findings_input",
                label_visibility="collapsed",
                height=140,
            )
            st.session_state.findings = findings

            dictation_component = """<div style="display:flex;gap:6px;align-items:center;margin-top:4px">
<button onclick="dStart()" style="border:1px solid var(--border-light);border-radius:4px;background:var(--bg-card);color:var(--accent);font-size:10px;font-family:var(--font-mono);padding:4px 10px;cursor:pointer">🎤 Dictar</button>
<button onclick="dStop()" style="border:1px solid var(--border-light);border-radius:4px;background:var(--bg-card);color:var(--accent);font-size:10px;font-family:var(--font-mono);padding:4px 10px;cursor:pointer" disabled>⏹ Detener</button>
<span id="ds" style="font-size:10px;color:var(--text-muted);font-family:var(--font-mono)">Inactivo</span></div>
<div id="dp" style="font-size:10px;color:var(--text-muted);margin-top:4px;min-height:16px;font-family:var(--font-mono)"></div>
<script>
var SR=window.SpeechRecognition||window.webkitSpeechRecognition,r=null,final='';
if(!SR){document.getElementById('ds').textContent='No compatible';}
function dStart(){if(!SR)return;r=new SR();r.lang='es-ES';r.continuous=true;r.interimResults=true;final='';
r.onresult=function(e){var inter='';for(var i=e.resultIndex;i<e.results.length;i++){if(e.results[i].isFinal)final+=e.results[i][0].transcript+' ';else inter+=e.results[i][0].transcript;}
document.getElementById('dp').innerHTML=final+'<em>'+inter+'</em>';};
r.onend=function(){document.getElementById('ds').textContent='Detenido';document.querySelectorAll('[onclick*=\"dStart\"]')[0].disabled=false;document.querySelectorAll('[onclick*=\"dStop\"]')[0].disabled=true;
if(final.trim())Streamlit.setComponentValue({type:'dictation',text:final.trim()});};
r.start();document.getElementById('ds').textContent='Escuchando...';document.querySelectorAll('[onclick*=\"dStart\"]')[0].disabled=true;document.querySelectorAll('[onclick*=\"dStop\"]')[0].disabled=false;}
function dStop(){if(r){r.stop();document.getElementById('ds').textContent='Procesando...';}}
</script>"""
            dv = components.html(dictation_component, height=70)
            if dv and isinstance(dv, dict) and dv.get("type") == "dictation":
                txt = dv.get("text", "")
                if txt:
                    cur = st.session_state.findings
                    st.session_state.findings = (cur + " " + txt).strip()
                    st.rerun()

        with st.expander("⊞ CLASIFICACIONES", expanded=True):
            clasif_name = st.selectbox(
                "", list(CLASIFICACIONES.keys()),
                key="clasif_sel", label_visibility="collapsed",
            )
            for grado, desc in CLASIFICACIONES[clasif_name]:
                cols = st.columns([1, 3])
                with cols[0]:
                    if st.button(grado, key=f"g_{grado}", use_container_width=True):
                        st.session_state.clasif_activas[clasif_name] = f"{grado}: {desc}"
                        st.rerun()
                with cols[1]:
                    st.markdown(
                        f'<span style="font-size:10px;font-family:var(--font-mono);color:var(--text-muted);line-height:32px">{desc}</span>',
                        unsafe_allow_html=True,
                    )

        with st.expander("⊞ PLANTILLA DOCX", expanded=False):
            archivo = st.file_uploader("Subir plantilla .docx", type=["docx"], label_visibility="collapsed")
            if archivo:
                txt = leer_plantilla(archivo)
                if txt:
                    st.session_state.uploaded_template = txt
                    st.success(f"Plantilla cargada ({len(txt)} caracteres)")

        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
        generar = st.button("⬡ GENERAR INFORME", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("Limpiar hallazgos", use_container_width=True):
                st.session_state.findings = ""
                st.rerun()
        with c2:
            if st.button("Refinar informe", use_container_width=True):
                st.session_state.refine_trigger = True

with right:
    ac1, ac2, ac3, ac4 = st.columns([1, 1, 1, 1])
    with ac1:
        if st.button("✦ Definiciones", key="def_btn", use_container_width=True):
            st.session_state.def_trigger = True
    with ac2:
        refinar = st.button("⟡ Optimizar", key="opt_btn", use_container_width=True)
    with ac3:
        if st.button("⟳ Limpiar", key="clear_btn", use_container_width=True):
            save_report("")
            st.session_state.last_saved = "Informe limpiado"
            st.rerun()
    with ac4:
        pass

    if generar:
        text = st.session_state.findings or st.session_state.report
        if text.strip() and copilot.is_available():
            with st.spinner("Generando informe..."):
                result = copilot.generate_report(text)
            if result and not result.startswith("Error"):
                save_report(result)
                st.session_state.last_saved = "Generado por IA"
        elif not copilot.is_available():
            st.warning("API key no configurada")

    if refinar or st.session_state.pop("refine_trigger", False):
        if st.session_state.report.strip() and copilot.is_available():
            with st.spinner("Refinando..."):
                result = copilot.refine_report(st.session_state.report)
            if result and not result.startswith("Error"):
                save_report(result)
                st.session_state.last_saved = "Refinado por IA"

    if st.session_state.pop("def_trigger", False):
        if st.session_state.report.strip() and copilot.is_available():
            with st.spinner("Analizando clasificaciones..."):
                client = copilot.client
                prompt = f"""Eres un radiólogo docente. Lee este informe y responde en formato estructurado:
1. CLASIFICACIONES USADAS: lista cada una, su grado y significado clínico.
2. CLASIFICACIONES FALTANTES: sugiere cuáles agregar según los hallazgos.
3. DEFINICIONES: define en 1-2 líneas los términos técnicos clave.
4. CORRELACIÓN CLÍNICA: impacto y manejo esperado.

INFORME:\n{st.session_state.report}"""
                try:
                    res = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[{"role": "user", "content": prompt}],
                        temperature=0.2,
                    )
                    st.session_state.definiciones = res.choices[0].message.content
                except Exception as e:
                    st.error(f"Error: {e}")

    report = st.text_area(
        "",
        value=st.session_state.report,
        placeholder="El informe generado aparecerá aquí. Edítalo directamente...",
        key="report_editor",
        label_visibility="collapsed",
    )
    if report != st.session_state.report:
        save_report(report)
        st.session_state.last_saved = "Autoguardado"

    wc = len(report.split()) if report.strip() else 0
    cc = len(report)
    st.markdown(
        f'<div style="text-align:right;font-size:10px;color:var(--text-muted);padding:2px 4px 0;font-family:var(--font-mono);border-top:1px solid var(--border)">{wc} palabras · {cc} caracteres</div>',
        unsafe_allow_html=True,
    )

    if st.session_state.definiciones:
        with st.expander("⬡ DEFINICIONES & CLASIFICACIONES", expanded=True):
            st.markdown(
                f'<div style="font-family:var(--font-mono);font-size:11px;color:var(--text);line-height:1.7;background:var(--bg-card);padding:12px;border-radius:var(--radius);border:1px solid var(--border-light);white-space:pre-wrap">{st.session_state.definiciones}</div>',
                unsafe_allow_html=True,
            )
            if st.button("Cerrar"):
                st.session_state.definiciones = ""
                st.rerun()
