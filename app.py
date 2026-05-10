import streamlit as st
from docx import Document
import speech_recognition as sr
import io
from openai import OpenAI

# ==========================================
# 1. MOTOR GRÁFICO CINEMATOGRÁFICO (GLASSMORPHISM & GLOW)
# ==========================================
st.set_page_config(page_title="Beam AI | Core", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    
    /* Fondo Abismal con Luz Volumétrica Superior */
    .stApp { 
        background: radial-gradient(ellipse at top, #110f1c 0%, #030305 100%); 
        color: #e2e8f0; 
        font-family: 'Inter', sans-serif; 
    }
    .block-container { padding-top: 1.5rem !important; max-width: 96% !important; }
    header { visibility: hidden; }
    
    /* Paneles Flotantes (Glassmorphism) */
    div[data-testid="stExpander"] { 
        background: rgba(255, 255, 255, 0.01) !important; 
        backdrop-filter: blur(12px) !important;
        border: 1px solid rgba(139, 92, 246, 0.1) !important; 
        border-radius: 16px !important; 
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.5);
    }
    div[data-testid="stExpander"] summary { color: #a78bfa !important; font-weight: 500 !important; }
    
    /* Text Areas (El Lienzo Holográfico) */
    .stTextArea textarea { 
        background: rgba(10, 10, 15, 0.4) !important; 
        backdrop-filter: blur(8px) !important;
        color: #ffffff !important; 
        border: 1px solid rgba(139, 92, 246, 0.15) !important; 
        border-radius: 16px !important;
        font-size: 15px !important; line-height: 1.7 !important; padding: 1rem !important;
        transition: all 0.3s ease;
    }
    .stTextArea textarea:focus { 
        border-color: #8b5cf6 !important; 
        box-shadow: 0 0 25px rgba(139, 92, 246, 0.15) !important; 
    }
    
    /* Botón de Fusión Térmica (Generar) */
    .btn-generar > div > button { 
        background: linear-gradient(135deg, #4c1d95 0%, #7c3aed 100%) !important; 
        border: 1px solid rgba(167, 139, 250, 0.3) !important;
        box-shadow: 0 0 20px rgba(124, 58, 237, 0.25) !important;
        color: white !important; border-radius: 12px !important; 
        font-weight: 600 !important; font-size: 16px !important; 
        padding: 0.8rem !important; transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .btn-generar > div > button:hover { 
        box-shadow: 0 0 35px rgba(139, 92, 246, 0.5) !important; 
        transform: translateY(-2px); 
    }
    
    /* Botones Secundarios Glass */
    .stButton>button { 
        background: rgba(255, 255, 255, 0.03) !important; 
        color: #a78bfa !important; 
        border: 1px solid rgba(139, 92, 246, 0.2) !important; 
        border-radius: 10px !important; 
        backdrop-filter: blur(5px);
    }
    .stButton>button:hover { background: rgba(139, 92, 246, 0.1) !important; }
    
    /* Tipografía y Brillos */
    h1, h2, h3 { color: #f8fafc !important; font-weight: 300 !important; letter-spacing: 0.5px; }
    h2 strong { font-weight: 600; background: -webkit-linear-gradient(#c4b5fd, #8b5cf6); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. NÚCLEO LÓGICO
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
# 3. INTERFAZ: EL SISTEMA OPERATIVO
# ==========================================
st.markdown("## Beam AI <strong>OS</strong>", unsafe_allow_html=True)

col_izq, col_der = st.columns([1.2, 2], gap="large")

with col_izq:
    if not api_key:
        api_key = st.text_input("🔑 DeepSeek API Key", type="password")
        
    with st.expander("⎈ Parámetros de Calibración", expanded=False):
        modalidad = st.selectbox("Protocolo", ["Resonancia Magnética", "Tomografía Computarizada", "Radiografía", "Ultrasonido", "PET-CT"])
        archivo_base = st.file_uploader("Inyectar Estructura Base (.docx)", type=["docx"])
        plantilla_txt = leer_plantilla(archivo_base) if archivo_base else ""
        
        instrucciones_estilo = st.text_area(
            "Directrices Óptimas:", 
            height=100, 
            value="Lenguaje médico experto. Infiere diagnósticos de la descripción. Expande clasificaciones en hallazgos anatómicos.",
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### ∿ Consola de Entrada Neural")
    
    # UNIFICACIÓN VISUAL: Audio y Texto juntos
    audio_data = st.audio_input("Registro Biométrico (Voz)", label_visibility="collapsed")
    if audio_data:
        nuevo = transcribir_voz(audio_data)
        if nuevo and nuevo not in st.session_state.dictado_actual:
            st.session_state.dictado_actual += " " + nuevo

    dictado_verificable = st.text_area("Señal transcrita / Inyección manual:", 
                                     value=st.session_state.dictado_actual, 
                                     height=180,
                                     placeholder="Ej. Articulaciones cigapofisarias, cambios osteocondrales, grasa de Hoffa...",
                                     label_visibility="collapsed")
    
    st.markdown('<div class="btn-generar">', unsafe_allow_html=True)
    if st.button("✧ Procesar Algoritmo Clínico"):
        if api_key and dictado_verificable:
            client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
            
            prompt_cerebro = f"""
            Eres el núcleo de inteligencia de Beam AI. Redacta un informe de {modalidad}.
            
            LÓGICA MATRICIAL:
            1. SÍNTESIS: Si recibes descripciones (ej. alteración de señal meniscal), agrupa e infiere el diagnóstico en la IMPRESIÓN DIAGNÓSTICA.
            2. EXPANSIÓN: Si recibes una conclusión (ej. Gonartrosis IV), redacta los cambios anatómicos esperados en los HALLAZGOS.
            3. PRECISIÓN MSK: Reconoce y respeta terminología anatómica avanzada (ej. no confundir "Hoffa" con "hoja").
            4. REGLAS DE USUARIO: {instrucciones_estilo}
            
            ESTRUCTURA:
            - NO USES ASTERISCOS (**). Títulos en MAYÚSCULAS.
            - DEBES respetar el formato de esta plantilla base: {plantilla_txt}
            
            INPUT DETECTADO: {dictado_verificable}
            """
            
            with st.spinner("Sintetizando modelo de datos..."):
                try:
                    res = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[{"role": "system", "content": prompt_cerebro}],
                        temperature=0.1
                    )
                    st.session_state.reporte_final = res.choices[0].message.content
                    st.rerun()
                except Exception as e: st.error(f"Error de enlace: {e}")
    st.markdown('</div>', unsafe_allow_html=True)
    
    if st.button("⌫ Purgar Caché"):
        st.session_state.dictado_actual = ""
        st.rerun()

with col_der:
    st.markdown("#### ⬡ Proyección Holográfica del Informe")
    
    reporte_editado = st.text_area(
        "Workspace", 
        value=st.session_state.reporte_final, 
        height=720, 
        label_visibility="collapsed"
    )
    
    if st.session_state.reporte_final:
        b1, b2 = st.columns([1, 1])
        with b1:
            st.download_button("↓ Extraer Documento", generar_docx(reporte_editado), "BeamAI_Core.docx", use_container_width=True)
        with b2:
            if st.button("⟡ Optimizar Criterio Diagnóstico", use_container_width=True):
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                with st.spinner("Reevaluando red neuronal..."):
                    try:
                        prompt_ref = f"""
                        Eres el sistema de optimización diagnóstica. 
                        Lee este informe y MEJORA ÚNICAMENTE la IMPRESIÓN DIAGNÓSTICA (hazla más elegante, concluyente y médica).
                        
                        PROTOCOLO ESTRICTO: Devuelve el informe COMPLETO. Conserva la Técnica y los Hallazgos intactos, solo sustituye el bloque final.
                        CERO asteriscos (**).
                        
                        REPORTE BASE: \n\n{reporte_editado}
                        """
                        
                        res_ref = client.chat.completions.create(
                            model="deepseek-chat",
                            messages=[{"role": "user", "content": prompt_ref}],
                            temperature=0.3
                        )
                        st.session_state.reporte_final = res_ref.choices[0].message.content
                        st.rerun()
                    except Exception as e: st.error("Error estructural.")
