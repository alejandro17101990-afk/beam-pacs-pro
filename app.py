import streamlit as st
from docx import Document
import speech_recognition as sr
import io
from openai import OpenAI

# ==========================================
# 1. CONFIGURACIÓN VISUAL PREMIUM (BEAM AI)
# ==========================================
st.set_page_config(page_title="Beam AI | Radiology Station", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    
    .stApp { background-color: #0b0b0f; color: #ededed; font-family: 'Inter', sans-serif; }
    
    /* Panel lateral oscuro */
    [data-testid="stSidebar"] { background-color: #111116; border-right: 1px solid #1f1f2e; }
    
    /* Text Areas Amplios y Limpios */
    .stTextArea textarea { 
        background-color: #16161d !important; color: #f8fafc !important; 
        border: 1px solid #2a2a35 !important; border-radius: 8px !important;
        font-size: 15px !important; line-height: 1.6 !important;
    }
    .stTextArea textarea:focus { border-color: #7c3aed !important; box-shadow: 0 0 0 1px #7c3aed !important; }

    /* Botones de Acción (Violeta Beam AI) */
    .primary-btn > div > button { 
        background: linear-gradient(135deg, #7c3aed 0%, #6d28d9 100%) !important; 
        color: white !important; border-radius: 8px !important; border: none !important;
        font-weight: 500 !important; width: 100%; transition: all 0.3s;
    }
    .primary-btn > div > button:hover { box-shadow: 0 0 15px rgba(124, 58, 237, 0.4); }
    
    /* Botones Secundarios */
    .stButton>button { background: #1a1a24 !important; color: #ededed !important; border: 1px solid #2a2a35 !important; border-radius: 8px !important; }
    .stButton>button:hover { border-color: #7c3aed !important; }
    
    h1, h2, h3, h4 { color: #ffffff !important; font-weight: 600 !important; }
    hr { border-color: #1f1f2e !important; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. FUNCIONES CLÍNICAS NÚCLEO
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

# Memoria de la sesión
if 'dictado_actual' not in st.session_state: st.session_state.dictado_actual = ""
if 'reporte_final' not in st.session_state: st.session_state.reporte_final = ""

# ==========================================
# 3. MENÚ LATERAL (Configuración)
# ==========================================
with st.sidebar:
    st.markdown("### 🔮 Beam AI")
    st.caption("Radiology Workstation")
    
    try:
        api_key = st.secrets["deepseek_key"]
    except:
        api_key = st.text_input("DeepSeek Key", type="password")
    
    st.divider()
    modalidad = st.selectbox("Estudio en curso", ["Resonancia Magnética", "Tomografía Computarizada", "Radiografía", "Ultrasonido", "PET-CT"])
    archivo_plantilla = st.file_uploader("Subir Formato/Plantilla", type=["docx"])
    texto_plantilla = leer_plantilla(archivo_plantilla) if archivo_plantilla else ""

# ==========================================
# 4. ÁREA DE TRABAJO PRINCIPAL
# ==========================================
st.markdown("## 🖥️ Estación de Interpretación")

# Hacemos la columna derecha mucho más ancha para simular el editor principal
col_izquierda, col_derecha = st.columns([1, 1.8], gap="large")

with col_izquierda:
    st.markdown("#### 🎙️ Registro de Dictado")
    
    audio_data = st.audio_input("Presiona para dictar")
    if audio_data:
        nuevo_texto = transcribir_voz(audio_data)
        if nuevo_texto and nuevo_texto not in st.session_state.dictado_actual:
            st.session_state.dictado_actual += " " + nuevo_texto

    st.caption("Verifica y corrige términos médicos aquí (Ej. Hoja -> Hoffa):")
    texto_editable_dictado = st.text_area(
        "Log de transcripción:", 
        value=st.session_state.dictado_actual, 
        height=300, 
        label_visibility="collapsed"
    )
    
    st.markdown('<div class="primary-btn">', unsafe_allow_html=True)
    if st.button("✨ Estructurar Informe"):
        if api_key and texto_editable_dictado:
            client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
            
            # PROMPT ESTRICTO PARA EVITAR ASTERISCOS Y PRIORIZAR EL DICTADO
            prompt_maestro = f"""
            Eres un experto radiólogo redactando un informe de {modalidad}.
            
            REGLAS DE FORMATO ESTRICTAS:
            1. NO USES ASTERISCOS (** o *) para negritas bajo ninguna circunstancia.
            2. Usa MAYÚSCULAS para los títulos de las secciones (ej. TÉCNICA, HALLAZGOS, IMPRESIÓN DIAGNÓSTICA).
            3. Escribe texto limpio, plano y profesional.
            
            REGLAS CLÍNICAS:
            1. El DICTADO DEL MÉDICO es la verdad absoluta. Si el dictado dice que algo es normal, SOBREESCRIBE cualquier patología que venga en la plantilla.
            2. Corrige errores de transcripción anatómica por contexto.
            
            PLANTILLA INSTITUCIONAL: {texto_plantilla}
            DICTADO DEL MÉDICO: {texto_editable_dictado}
            """
            
            with st.spinner("Procesando hallazgos..."):
                try:
                    res = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[{"role": "system", "content": prompt_maestro}],
                        temperature=0.1
                    )
                    st.session_state.reporte_final = res.choices[0].message.content
                    st.session_state.dictado_actual = texto_editable_dictado # Guardar correcciones
                    st.rerun()
                except Exception as e:
                    st.error(f"Error de conexión: {e}")
    st.markdown('</div>', unsafe_allow_html=True)
    
    if st.button("🗑️ Limpiar Dictado"):
        st.session_state.dictado_actual = ""
        st.rerun()

with col_derecha:
    st.markdown("#### 📄 Editor de Informe")
    
    # Text area gigante para trabajar cómodamente
    reporte_editado = st.text_area(
        "Workspace", 
        value=st.session_state.reporte_final, 
        height=650, 
        label_visibility="collapsed"
    )
    
    if st.session_state.reporte_final:
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            st.download_button("📥 Exportar a Word", generar_docx(reporte_editado), "Informe_BeamAI.docx")
        with col_btn2:
            if st.button("🔄 Reformular Conclusión"):
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                with st.spinner("Sintetizando nueva conclusión..."):
                    try:
                        prompt_ref = f"""Reescribe ÚNICAMENTE la IMPRESIÓN DIAGNÓSTICA de este texto para que sea más experta y estructurada. 
                        NO uses asteriscos (**). Escribe en texto limpio. Mantén los hallazgos intactos: \n\n{reporte_editado}"""
                        
                        res_ref = client.chat.completions.create(
                            model="deepseek-chat",
                            messages=[{"role": "user", "content": prompt_ref}],
                            temperature=0.3
                        )
                        st.session_state.reporte_final = res_ref.choices[0].message.content
                        st.rerun()
                    except:
                        st.error("Error al reformular.")
