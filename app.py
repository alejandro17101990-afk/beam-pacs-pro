import streamlit as st
from docx import Document
import speech_recognition as sr
import io
from openai import OpenAI

# 1. CONFIGURACIÓN DE INTERFAZ (Inspiración v0 / Perplexity)
st.set_page_config(page_title="Lumen AI | Estación de Trabajo", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,300;0,400;0,600;1,400&display=swap');
    .stApp { background-color: #09090b; color: #f4f4f5; font-family: 'Inter', sans-serif; }
    
    /* Panel lateral elegante */
    [data-testid="stSidebar"] { background-color: #0c0c0f; border-right: 1px solid #27272a; }
    
    /* Contenedores de edición */
    .stTextArea textarea { 
        background-color: #121214 !important; color: #ffffff !important; 
        border: 1px solid #27272a !important; border-radius: 12px !important;
        font-size: 16px !important; line-height: 1.6 !important;
    }
    .stTextArea textarea:focus { border-color: #6366f1 !important; box-shadow: 0 0 0 1px #6366f1 !important; }

    /* Botones Premium */
    .stButton>button { 
        background: #ffffff !important; color: #09090b !important; 
        border-radius: 10px !important; font-weight: 600 !important;
        padding: 0.6rem 2rem !important; border: none !important;
    }
    .stButton>button:hover { background: #e4e4e7 !important; transform: translateY(-1px); }
    
    /* Botón Secundario (Reformular) */
    .secondary-btn button { background: #18181b !important; color: #ffffff !important; border: 1px solid #27272a !important; }

    /* Estilo de la barra de dictado */
    [data-testid="stChatInput"] { background-color: #121214 !important; border: 1px solid #27272a !important; border-radius: 20px !important; }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNCIONES DE LÓGICA CLÍNICA
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

def transcribir_instantaneo(audio_file):
    r = sr.Recognizer()
    with sr.AudioFile(audio_file) as source:
        try:
            return r.recognize_google(r.record(source), language="es-MX")
        except: return ""

# Inicialización de memoria
if 'reporte_editable' not in st.session_state: st.session_state.reporte_editable = ""
if 'dictado_acumulado' not in st.session_state: st.session_state.dictado_acumulado = ""

# 3. BARRA LATERAL (Configuración y Plantillas)
with st.sidebar:
    st.markdown("### 🧬 Lumen Core")
    st.caption("Configuración de Especialidad")
    
    try:
        api_key = st.secrets["deepseek_key"]
    except:
        api_key = st.text_input("DeepSeek Key", type="password")
    
    st.divider()
    modalidad = st.selectbox("Modalidad", ["Resonancia", "Tomografía", "Radiografía", "Ultrasonido", "PET-CT"])
    archivo_base = st.file_uploader("Cargar Plantilla Institucional", type=["docx"])
    plantilla_txt = leer_plantilla(archivo_base) if archivo_base else ""
    
    st.divider()
    st.info("💡 **Consejo:** El dictado ahora prioriza tus hallazgos sobre lo que diga la plantilla base.")

# 4. WORKSPACE DE DOBLE PANEL
col_dictado, col_reporte = st.columns([1, 1.3], gap="large")

with col_dictado:
    st.markdown("#### 🎙️ Dictado y Hallazgos")
    
    # Dictado de voz con feedback inmediato
    audio_data = st.audio_input("Dictar ahora")
    if audio_data:
        transcripcion = transcribir_instantaneo(audio_data)
        if transcripcion and transcripcion not in st.session_state.dictado_acumulado:
            st.session_state.dictado_acumulado += " " + transcripcion
    
    # El usuario puede ver y editar lo que se ha dictado antes de procesar
    hallazgos_finales = st.text_area("Hallazgos para procesar (puedes editarlos):", 
                                     value=st.session_state.dictado_acumulado, 
                                     height=250,
                                     placeholder="Lo que dictes aparecerá aquí. Corrígelo si es necesario...")
    
    if st.button("✨ Generar Informe"):
        if api_key and hallazgos_finales:
            client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
            
            prompt_medico = f"""
            Eres un radiólogo experto. Genera un informe de {modalidad}.
            
            INSTRUCCIÓN CRUCIAL:
            - Prioriza los HALLAZGOS DICTADOS. 
            - Si la PLANTILLA BASE contiene patología pero los HALLAZGOS dicen que es normal, el informe DEBE ser normal.
            - Corrige errores fonéticos (ej: si dice 'hoja' en contexto de rodilla, cámbialo a 'Grasa de Hoffa').
            
            PLANTILLA BASE: {plantilla_txt}
            HALLAZGOS DICTADOS: {hallazgos_finales}
            """
            
            with st.spinner("IA redactando..."):
                res = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "system", "content": prompt_medico}],
                    temperature=0.1
                )
                st.session_state.reporte_editable = res.choices[0].message.content
                st.rerun()

with col_reporte:
    st.markdown("#### 📄 Informe Radiológico")
    
    # AQUÍ ESTÁ EL EDITOR: El informe es totalmente editable
    reporte_final = st.text_area("Editor de Informe Final:", 
                                 value=st.session_state.reporte_editable, 
                                 height=600)
    
    if st.session_state.reporte_editable:
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 Descargar Word", generar_docx(reporte_final), "Reporte_Radiologico.docx")
        with c2:
            st.markdown('<div class="secondary-btn">', unsafe_allow_html=True)
            if st.button("🔄 Reformular Conclusión"):
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                with st.spinner("Mejorando conclusión..."):
                    res_ref = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[{"role": "user", "content": f"Reescribe solo la conclusión de este reporte para que sea más experta y estructurada, mantén el resto igual: {reporte_final}"}],
                        temperature=0.3
                    )
                    st.session_state.reporte_editable = res_ref.choices[0].message.content
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

# Limpiar dictado
if st.sidebar.button("🗑️ Limpiar Sesión"):
    st.session_state.dictado_acumulado = ""
    st.session_state.reporte_editable = ""
    st.rerun()
