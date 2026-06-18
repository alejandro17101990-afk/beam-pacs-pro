import json
import streamlit as st
import streamlit.components.v1 as components
from core.ai import RadiologyCopilot
from core.knowledge import KnowledgeBase
from services.export import export_to_docx
from services.storage import DraftStorage

st.set_page_config(page_title="Beam AI", layout="wide", initial_sidebar_state="collapsed")

with open("ui/styles.css", "r", encoding="utf-8") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

st.session_state.setdefault("theme", "dark")

defaults = {
    "draft_report": "",
    "findings_input": "",
    "layout_mode": "horizontal",
    "input_width": 35,
    "show_input": True,
    "last_saved": "",
    "selected_template": "RM Rodilla Estándar",
}
for k, v in defaults.items():
    st.session_state.setdefault(k, v)

storage = DraftStorage()
if not st.session_state.draft_report:
    loaded = storage.load_draft()
    if loaded:
        st.session_state.draft_report = loaded

kb = KnowledgeBase("data/knowledge_base.json")
copilot = RadiologyCopilot(api_key=st.secrets.get("deepseek_key", ""), knowledge_base=kb)

# ─── TOP BAR ──────────────────────────────────────────────────────────

c1, c2, c3, c4, c5, c6, c7, c8 = st.columns([0.35, 0.65, 0.65, 1.1, 1.8, 0.55, 0.9, 0.35])

with c1:
    st.markdown('<div class="brand">BEAM AI</div>', unsafe_allow_html=True)

with c2:
    is_horiz = st.session_state.layout_mode == "horizontal"
    if st.button("☰ Horizontal" if is_horiz else "☷ Vertical", key="layout_btn", use_container_width=True):
        st.session_state.layout_mode = "vertical" if is_horiz else "horizontal"

with c3:
    label = "Ocultar" if st.session_state.show_input else "Mostrar"
    if st.button(f"⊞ {label}", key="toggle_input", use_container_width=True):
        st.session_state.show_input = not st.session_state.show_input

with c4:
    wc1, wc2, wc3 = st.columns([0.2, 0.6, 0.2])
    with wc1:
        if st.button("−", key="narrower"):
            st.session_state.input_width = max(15, st.session_state.input_width - 5)
    with wc2:
        st.markdown(
            f'<div class="width-indicator">{st.session_state.input_width}%</div>',
            unsafe_allow_html=True,
        )
    with wc3:
        if st.button("+", key="wider"):
            st.session_state.input_width = min(65, st.session_state.input_width + 5)

with c5:
    generated = st.button("Generar", key="generate", type="primary", use_container_width=True)

with c6:
    if st.button("Copiar", key="copy_btn"):
        st.session_state.copy_trigger = True

with c7:
    col_a, col_b = st.columns([1, 0.4])
    with col_a:
        docx_data = export_to_docx(st.session_state.draft_report)
        st.download_button(
            "DOCX",
            data=docx_data,
            file_name="BeamAI_Informe.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="export_docx",
        )
    with col_b:
        theme_label = "☀️" if st.session_state.theme == "dark" else "🌙"
        if st.button(theme_label, key="theme_btn"):
            st.session_state.theme = "light" if st.session_state.theme == "dark" else "dark"

with c8:
    pass

st.markdown(
    f"<script>document.documentElement.setAttribute('data-theme','{st.session_state.theme}')</script>",
    unsafe_allow_html=True,
)

# ─── AUTO-SAVE BAR ───────────────────────────────────────────────────

if st.session_state.last_saved:
    st.markdown(
        f'<div class="autosave-bar">{st.session_state.last_saved}</div>',
        unsafe_allow_html=True,
    )

# ─── GENERATE LOGIC ──────────────────────────────────────────────────

if generated:
    text = st.session_state.findings_input or st.session_state.draft_report
    if text.strip():
        with st.spinner("Generando informe estructurado..."):
            result = copilot.generate_report(text)
        st.session_state.draft_report = result
        storage.save_draft(result)
        st.session_state.last_saved = "Generado por IA"

# ─── COPY LOGIC ──────────────────────────────────────────────────────

if st.session_state.pop("copy_trigger", False):
    safe_text = json.dumps(st.session_state.draft_report)
    components.html(
        f"""<script>
navigator.clipboard.writeText({safe_text}).then(() => {{
const t=document.createElement('div');
t.textContent='Copiado al portapapeles';
t.style.cssText='position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#252525;color:#e5e5e5;padding:10px 20px;border-radius:10px;font-size:13px;z-index:9999;border:1px solid #333;';
document.body.appendChild(t);setTimeout(()=>t.remove(),2000);
}});
</script>""",
        height=0,
    )

# ─── DICTATION COMPONENT ─────────────────────────────────────────────

dictation_html = """<div style="display:flex;gap:6px;align-items:center;">
<button id="dStart" style="border:1px solid var(--border-color);border-radius:6px;background:var(--bg-tertiary);color:var(--text-primary);font-size:12px;padding:4px 12px;height:32px;cursor:pointer;">🎤 Iniciar</button>
<button id="dStop" style="border:1px solid var(--border-color);border-radius:6px;background:var(--bg-tertiary);color:var(--text-primary);font-size:12px;padding:4px 12px;height:32px;cursor:pointer;" disabled>⏹ Detener</button>
<span id="dStatus" style="font-size:11px;color:var(--text-tertiary);">Inactivo</span>
</div>
<div id="dPartial" style="font-size:12px;color:var(--text-tertiary);margin-top:6px;min-height:20px;"></div>
<script>
(function(){
const SR=window.SpeechRecognition||window.webkitSpeechRecognition;
if(!SR){document.getElementById('dStatus').textContent='No compatible';return;}
const r=new SR();r.lang='es-ES';r.continuous=true;r.interimResults=true;
let final='';
r.onresult=function(e){
  let interim='';
  for(let i=e.resultIndex;i<e.results.length;i++){
    if(e.results[i].isFinal) final+=e.results[i][0].transcript+' ';
    else interim+=e.results[i][0].transcript;
  }
  document.getElementById('dPartial').innerHTML=final+'<em>'+interim+'</em>';
};
r.onend=function(){
  document.getElementById('dStatus').textContent='Detenido';
  document.getElementById('dStart').disabled=false;
  document.getElementById('dStop').disabled=true;
  if(final.trim()) Streamlit.setComponentValue({type:'dictation',text:final.trim()});
};
document.getElementById('dStart').onclick=function(){
  final='';document.getElementById('dPartial').textContent='';r.start();
  document.getElementById('dStatus').textContent='Escuchando...';this.disabled=true;
  document.getElementById('dStop').disabled=false;
};
document.getElementById('dStop').onclick=function(){r.stop();document.getElementById('dStatus').textContent='Procesando...';};
})();
</script>"""

# ─── DICTATION HANDLER ───────────────────────────────────────────────

def render_dictation():
    dict_val = components.html(dictation_html, height=100)
    if dict_val and isinstance(dict_val, dict) and dict_val.get("type") == "dictation":
        txt = dict_val.get("text", "")
        if txt:
            cur = st.session_state.findings_input
            st.session_state.findings_input = (cur + " " + txt).strip()
            st.rerun()

# ─── TEMPLATES AND COMMON UI HELPERS ─────────────────────────────────

def render_template_ui(key_suffix):
    templates = kb.list_templates()
    sel_idx = 0
    if st.session_state.selected_template in templates:
        sel_idx = templates.index(st.session_state.selected_template)
    selected = st.selectbox(
        "",
        options=templates,
        index=sel_idx,
        key=f"tpl_{key_suffix}",
        label_visibility="collapsed",
    )
    st.session_state.selected_template = selected
    tc1, tc2 = st.columns(2)
    with tc1:
        if st.button("Cargar", key=f"load_{key_suffix}", use_container_width=True):
            st.session_state.draft_report = kb.render_template(selected)
            storage.save_draft(st.session_state.draft_report)
            st.session_state.last_saved = f"Plantilla: {selected}"
    with tc2:
        if st.button("Insertar", key=f"ins_{key_suffix}", use_container_width=True):
            tpl = kb.render_template(selected)
            cur = st.session_state.draft_report
            st.session_state.draft_report = f"{cur}\n\n{tpl}" if cur else tpl
            storage.save_draft(st.session_state.draft_report)
            st.session_state.last_saved = f"Insertada: {selected}"


def render_commands(key_suffix, cols_per_row=2):
    cmds = kb.get_slash_commands()
    cmd_items = list(cmds.items())
    for i in range(0, len(cmd_items), cols_per_row):
        cols = st.columns(cols_per_row)
        for j in range(cols_per_row):
            idx = i + j
            if idx < len(cmd_items):
                k, v = cmd_items[idx]
                lbl = v.get("label", k) if isinstance(v, dict) else v
                if ":" in lbl:
                    lbl = lbl.split(":")[0]
                with cols[j]:
                    if st.button(lbl, key=f"cmd_{key_suffix}_{k}", use_container_width=True):
                        tpl = v.get("template", v) if isinstance(v, dict) else v
                        cur = st.session_state.draft_report
                        st.session_state.draft_report = f"{cur}\n\n{tpl}" if cur else tpl
                        storage.save_draft(st.session_state.draft_report)


# ─── MAIN LAYOUT ─────────────────────────────────────────────────────

if st.session_state.layout_mode == "horizontal":
    show = st.session_state.show_input
    iw = st.session_state.input_width / 100 if show else 0.0
    input_col, editor_col = st.columns([iw, 1.0 - iw])

    with input_col:
        if show:
            st.markdown('<div class="section-header">Hallazgos</div>', unsafe_allow_html=True)
            findings = st.text_area(
                "",
                value=st.session_state.findings_input,
                placeholder="Escribe los hallazgos clínicos aquí...",
                key="findings_h",
                label_visibility="collapsed",
            )
            st.session_state.findings_input = findings

            st.markdown('<div class="section-header">Dictado por voz</div>', unsafe_allow_html=True)
            render_dictation()

            st.markdown('<div class="section-header">Plantilla</div>', unsafe_allow_html=True)
            render_template_ui("h")

            st.markdown('<div class="section-header">Comandos</div>', unsafe_allow_html=True)
            render_commands("h", cols_per_row=2)

    with editor_col:
        report = st.text_area(
            "",
            value=st.session_state.draft_report,
            placeholder="Informe radiológico — edita directamente aquí",
            key="editor_h",
            label_visibility="collapsed",
        )
        if report != st.session_state.draft_report:
            st.session_state.draft_report = report
            storage.save_draft(report)
            st.session_state.last_saved = "Autoguardado"

else:
    # ── VERTICAL MODE ────────────────────────────────────────────────
    if st.session_state.show_input:
        vc1, vc2, vc3 = st.columns([2, 1, 1])
        with vc1:
            st.markdown('<div class="section-header">Hallazgos</div>', unsafe_allow_html=True)
            findings = st.text_area(
                "",
                value=st.session_state.findings_input,
                placeholder="Escribe los hallazgos clínicos aquí...",
                key="findings_v",
                label_visibility="collapsed",
            )
            st.session_state.findings_input = findings
        with vc2:
            st.markdown('<div class="section-header">Dictado</div>', unsafe_allow_html=True)
            render_dictation()
        with vc3:
            st.markdown('<div class="section-header">Plantilla</div>', unsafe_allow_html=True)
            render_template_ui("v")

        st.markdown('<div class="section-header">Comandos</div>', unsafe_allow_html=True)
        render_commands("v", cols_per_row=4)

    report = st.text_area(
        "",
        value=st.session_state.draft_report,
        placeholder="Informe radiológico — edita directamente aquí",
        key="editor_v",
        label_visibility="collapsed",
    )
    if report != st.session_state.draft_report:
        st.session_state.draft_report = report
        storage.save_draft(report)
        st.session_state.last_saved = "Autoguardado"
