import streamlit as st
from docx import Document
import speech_recognition as sr
import io
from openai import OpenAI

# ==========================================
# 1. ESTÉTICA PREMIUM (ESTILO GEMINI / CHATGPT)
# ==========================================
st.set_page_config(page_title="Beam AI | Multimodal", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    
    /* Fondo oscuro estilo Gemini y reset de márgenes */
    .block-container { padding-top: 1.5rem !important; max-width: 98% !important; }
    .stApp { background-color: #0e0e11; color: #e3e3e8; font-family: 'Inter', sans-serif; }
    
    /* Contenedores elegantes y expanders */
    div[data-testid="stExpander"] { background-color: #18181c !important; border: 1px solid #27272f !important; border-radius: 12px !important; }
    div[data-testid="stExpander"] summary { color: #a2a2b0 !important; font-weight: 500 !important; }
    
    /* Text Areas (El lienzo del documento) */
    .stTextArea textarea { 
        background-color: #131317 !important; color: #ffffff !important; 
        border: 1px solid #2a2a35 !important; border-radius: 16px !important;
        font-size: 15px !important; line-height: 1.7 !important; padding: 1rem !important;
    }
    .stTextArea textarea:focus { border-color: #8b5cf6 !important; box-shadow: 0 0 0 1px #8b5cf6 !important; }
    
    /* Entradas de archivos y selectores */
    .stSelectbox div[data-baseweb="select"], .stFileUploader > div { background-color: #18181c !important; border: 1px solid #27272f !important; border-radius: 12px !important; }
    
    /* Botón Principal (Estilo Generativo) */
    .generar-btn > div > button { 
        background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%) !important; 
        color: white !important; border-radius: 12px !important; border: none !important;
        font-weight: 600 !important; font-size: 16px !important; padding: 0.8rem !important; transition: 0.3s;
    }
    .generar-btn > div > button:hover { box-shadow: 0 4px 20px rgba(139, 92, 246, 0.4); transform: translateY(-2px); }
    
    /* Botones secundarios */
    .stButton>button { background: #1f1f26 !important; color: #e3e3e8 !important; border: 1px solid #333340 !important; border-radius: 10px !important; }
    .stButton>button:hover { background: #2a2a35 !important; border-color: #8b5cf6 !important; }
    
    h1, h2, h3 { color: #ffffff !important; font-weight: 600 !important; letter-spacing: -0.5px; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. MOTORES CLÍNICOS Y DE IA
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

try: api_key = st.secrets["deepseek_key"]
except: api_key = ""

# ==========================================
# 3. INTERFAZ PRINCIPAL (LAYOUT GEMINI)
# ==========================================
st.markdown("## ✨ Beam AI Studio")

# Layout asimétrico: 35% Controles / 65% Editor
col_izq, col_der = st.columns([1.2, 2], gap="large")

with col_izq:
    if not api_key:
        api_key = st.text_input("🔑 DeepSeek API Key", type="password")
        
    # PANEL DE CONFIGURACIÓN (A la vista, no escondido)
    with st.expander("⚙️ Parámetros del Modelo (Plantilla y Reglas)", expanded=True):
        modalidad = st.selectbox("Modalidad", ["Resonancia Magnética", "Tomografía Computarizada", "Radiografía", "Ultrasonido", "PET-CT"])
        archivo_base = st.file_uploader("Subir formato .docx", type=["docx"])
        plantilla_txt = leer_plantilla(archivo_base) if archivo_base else ""
        
        instrucciones_estilo = st.text_area(
            "Prompts / Instrucciones de Estilo:", 
            height=100, 
            value="Lenguaje médico formal. Si describo patología, propón el diagnóstico en la conclusión. Si doy el diagnóstico, expande los hallazgos anatómicos.",
            help="Estas reglas se aplicarán a todos los informes generados."
        )

    st.markdown("#### 💬 Área de Dictado")
    
    # Entrada Multimodal
    audio_data = st.audio_input("Dictar hallazgos")
    if audio_data:
        nuevo = transcribir_voz(audio_data)
        if nuevo and nuevo not in st.session_state.dictado_actual:
            st.session_state.dictado_actual += " " + nuevo

    dictado_verificable = st.text_area("O escribe/corrige tu dictado aquí:", 
                                     value=st.session_state.dictado_actual, 
                                     height=180)
    
    # Botón Principal
    st.markdown('<div class="generar-btn">', unsafe_allow_html=True)
    if st.button("✦ Generar Informe Radiológico"):
        if api_key and dictado_verificable:
            client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
            
            prompt_cerebro = f"""
            Eres Beam AI, un modelo radiológico experto. Redacta un informe de {modalidad}.
            
            REGLAS DE INTELIGENCIA:
            1. SÍNTESIS/INFERENCIA: Si el usuario te describe hallazgos en detalle, agrupa la información y propón un diagnóstico concluyente y estilizado en la IMPRESIÓN DIAGNÓSTICA.
            2. EXPANSIÓN: Si el usuario te dicta una clasificación directa (ej. Gonartrosis grado 4), tú debes redactar la descripción morfológica detallada en los HALLAZGOS.
            3. PROMPTS DEL USUARIO: {instrucciones_estilo}
            4. FORMATO: NO USES ASTERISCOS (**). Títulos en mayúsculas (ej. HALLAZGOS:).
            
            PLANTILLA BASE A RESPETAR: {plantilla_txt}
            DICTADO/INPUT: {dictado_verificable}
            """
            
            with st.spinner("Procesando multimodalidad..."):
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
    
    if st.button("🗑️ Limpiar Dictado"):
        st.session_state.dictado_actual = ""
        st.rerun()

with col_der:
    st.markdown("#### 📄 Documento Interactivo")
    
    reporte_editado = st.text_area(
        "Workspace", 
        value=st.session_state.reporte_final, 
        height=700, 
        label_visibility="collapsed"
    )
    
    if st.session_state.reporte_final:
        b1, b2 = st.columns([1, 1])
        with b1:
            st.download_button("📥 Descargar Reporte (.docx)", generar_docx(reporte_editado), "BeamAI_Reporte.docx", use_container_width=True)
        with b2:
            if st.button("✨ Estilizar Conclusión", use_container_width=True):
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                with st.spinner("Pensando diagnóstico..."):
                    try:
                        prompt_ref = f"""
                        Actúa como un Jefe de Radiología de alta especialidad. 
                        Lee este informe y MEJORA ÚNICAMENTE la IMPRESIÓN DIAGNÓSTICA.
                        Analiza los hallazgos y propón una conclusión elegante, integradora y diagnóstica.
                        
                        REGLA ABSOLUTA: Devuelve el informe COMPLETO. Conserva la Técnica y los Hallazgos exactamente como te los entrego, y solo cambia la parte final.
                        SIN asteriscos (**).
                        
                        REPORTE: \n\n{reporte_editado}
                        """
                        
                        res_ref = client.chat.completions.create(
                            model="deepseek-chat",
                            messages=[{"role": "user", "content": prompt_ref}],
                            temperature=0.3
                        )
                        st.session_state.reporte_final = res_ref.choices[0].message.content
                        st.rerun()
                    except Exception as e: st.error("Error al mejorar.")
