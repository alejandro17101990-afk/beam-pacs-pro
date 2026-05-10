import streamlit as st
from docx import Document
import speech_recognition as sr
import io
from openai import OpenAI

# ==========================================
# 1. ESTÉTICA BEAM AI (MODO ESTACIÓN)
# ==========================================
st.set_page_config(page_title="Beam AI | Inteligencia Radiológica", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    .stApp { background-color: #0b0b0f; color: #ededed; font-family: 'Inter', sans-serif; }
    [data-testid="stSidebar"] { background-color: #111116; border-right: 1px solid #1f1f2e; }
    .stTextArea textarea { 
        background-color: #16161d !important; color: #ffffff !important; 
        border: 1px solid #2a2a35 !important; border-radius: 12px !important;
        font-size: 15px !important; line-height: 1.6 !important;
    }
    .primary-btn > div > button { 
        background: linear-gradient(135deg, #7c3aed 0%, #6d28d9 100%) !important; 
        color: white !important; border-radius: 8px !important; border: none !important;
        font-weight: 600 !important; width: 100%; transition: all 0.3s; padding: 0.8rem !important;
    }
    .primary-btn > div > button:hover { box-shadow: 0 0 20px rgba(124, 58, 237, 0.5); }
    h1, h2, h3, h4 { color: #ffffff !important; font-weight: 600 !important; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. FUNCIONES DE LÓGICA MÉDICA
# ==========================================
def leer_plantilla(file):
    doc = Document(file)
    return '\n'.join([p.text for p in doc.paragraphs if p.text.strip()])

def generar_docx(texto):
    doc = Document()
    for linea in texto.split('\n'):
        doc.add_paragraph(linea)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def transcribir_voz(audio_file):
    r = sr.Recognizer()
    with sr.AudioFile(audio_file) as source:
        try: return r.recognize_google(r.record(source), language="es-MX")
        except: return ""

if 'dictado_actual' not in st.session_state: st.session_state.dictado_actual = ""
if 'reporte_final' not in st.session_state: st.session_state.reporte_final = ""

# ==========================================
# 3. BARRA LATERAL: REGLAS Y CONFIGURACIÓN
# ==========================================
with st.sidebar:
    st.markdown("### 🔮 Beam AI")
    st.caption("Estación de Creación Multimodal")
    
    try: api_key = st.secrets["deepseek_key"]
    except: api_key = st.text_input("DeepSeek Key", type="password")
    
    st.divider()
    modalidad = st.selectbox("Estudio", ["Resonancia", "Tomografía", "Rayos X", "Ultrasonido", "PET-CT"])
    archivo_base = st.file_uploader("Subir Plantilla (.docx)", type=["docx"])
    plantilla_txt = leer_plantilla(archivo_base) if archivo_base else ""
    
    st.divider()
    st.markdown("**Reglas de Redacción (Prompts)**")
    instrucciones_estilo = st.text_area(
        "Instrucciones de estilo:", 
        height=150, 
        value="Usa un lenguaje médico formal. No utilices asteriscos para negritas. Si menciono una clasificación, expande la descripción técnica en los hallazgos.",
        placeholder="Ej: No usar gerundios. Estilo conciso..."
    )

# ==========================================
# 4. WORKSPACE DE ALTA PRODUCTIVIDAD
# ==========================================
st.markdown("## 🖥️ Estación de Interpretación")

col_input, col_editor = st.columns([1, 1.8], gap="large")

with col_input:
    st.markdown("#### 🎙️ Dictado Inteligente")
    audio_data = st.audio_input("Grabar hallazgos")
    
    if audio_data:
        nuevo = transcribir_voz(audio_data)
        if nuevo and nuevo not in st.session_state.dictado_actual:
            st.session_state.dictado_actual += " " + nuevo

    dictado_verificable = st.text_area("Hallazgos / Clasificaciones detectadas:", 
                                     value=st.session_state.dictado_actual, 
                                     height=280)
    
    st.markdown('<div class="primary-btn">', unsafe_allow_html=True)
    if st.button("✨ Procesar e Inferir"):
        if api_key and dictado_verificable:
            client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
            
            # EL CEREBRO: PROMPT DE EXPANSIÓN Y SÍNTESIS
            prompt_cerebro = f"""
            Eres un experto radiólogo. Tu misión es redactar un informe de {modalidad}.
            
            COMPORTAMIENTO INTELIGENTE:
            1. EXPANSIÓN: Si el usuario dicta una clasificación (ej. Gonartrosis grado 4, Bosniak II, Kellgren III), tú debes redactar la descripción técnica completa en la sección de HALLAZGOS basándote en los criterios médicos internacionales.
            2. INFERENCIA: Si el usuario dicta descripciones detalladas, tú debes proponer el diagnóstico o clasificación correspondiente en la IMPRESIÓN DIAGNÓSTICA.
            3. ESTILO: Respeta estrictamente estas reglas: {instrucciones_estilo}.
            4. FORMATO: Títulos en MAYÚSCULAS. Prohibido usar asteriscos (**).
            
            PLANTILLA: {plantilla_txt}
            DICTADO: {dictado_verificable}
            """
            
            with st.spinner("IA Pensando y Estructurando..."):
                try:
                    res = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[{"role": "system", "content": prompt_cerebro}],
                        temperature=0.1
                    )
                    st.session_state.reporte_final = res.choices[0].message.content
                    st.rerun()
                except Exception as e: st.error(f"Error: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

with col_editor:
    st.markdown("#### 📄 Editor de Informe Profesional")
    reporte_editado = st.text_area("Workspace", value=st.session_state.reporte_final, height=620, label_visibility="collapsed")
    
    if st.session_state.reporte_final:
        b1, b2 = st.columns(2)
        with b1:
            st.download_button("📥 Exportar Word", generar_docx(reporte_editado), "Reporte_Beam_AI.docx")
        with b2:
            if st.button("🔄 Reformular Conclusión"):
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                with st.spinner("Estilizando propuesta diagnóstica..."):
                    try:
                        prompt_ref = f"""Actúa como un Jefe de Radiología. Toma el siguiente reporte y reescribe ÚNICAMENTE la sección de IMPRESIÓN DIAGNÓSTICA. 
                        No repitas los hallazgos. Propón una síntesis elegante y jerarquizada basada en la evidencia descrita.
                        REPORTE COMPLETO: \n\n{reporte_editado}"""
                        
                        res_ref = client.chat.completions.create(
                            model="deepseek-chat",
                            messages=[{"role": "user", "content": prompt_ref}],
                            temperature=0.3
                        )
                        # Reemplazar solo la parte de la conclusión o actualizar el estado
                        st.session_state.reporte_final = res_ref.choices[0].message.content
                        st.rerun()
                    except: st.error("Error al reformular.")
