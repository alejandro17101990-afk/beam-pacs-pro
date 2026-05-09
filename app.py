import streamlit as st
from docx import Document
import speech_recognition as sr
from openai import OpenAI
import io

# 1. CONFIGURACIÓN DEL WORKSPACE (Modo Amplio)
st.set_page_config(page_title="Beam AI | Radiology Workspace", layout="wide", initial_sidebar_state="expanded")

# 2. MOTOR UI/UX: ESTÉTICA PREMIUM (Dark Mode Médico)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    .stApp { background-color: #0b0b0f; color: #e2e8f0; font-family: 'Inter', sans-serif; }
    [data-testid="stSidebar"] { background-color: #111116; border-right: 1px solid #1f1f2e; }
    .stButton>button { 
        background: linear-gradient(135deg, #7c3aed 0%, #6d28d9 100%) !important;
        color: white !important; border-radius: 8px !important; border: 1px solid #8b5cf6 !important;
        padding: 0.6rem 1rem !important; font-weight: 500 !important; width: 100%; transition: all 0.3s;
    }
    .stButton>button:hover { box-shadow: 0 0 15px rgba(124, 58, 237, 0.4); }
    div[data-testid="stVerticalBlock"] > div[style*="flex-direction: column;"] { 
        background-color: #12121a; border: 1px solid #1f1f2e; border-radius: 12px; padding: 24px; 
    }
    .stTextInput input, .stSelectbox div[data-baseweb="select"], .stTextArea textarea { 
        background-color: #16161d !important; color: #f8fafc !important; 
        border: 1px solid #2a2a35 !important; border-radius: 8px !important; font-size: 15px !important;
    }
    h1, h2, h3, h4 { color: #ffffff !important; font-weight: 600 !important; }
    </style>
    """, unsafe_allow_html=True)

# 3. FUNCIONES DE PROCESAMIENTO AVANZADO
def leer_word_con_tablas(file):
    doc = Document(file)
    contenido = []
    for para in doc.paragraphs:
        if para.text.strip(): contenido.append(para.text)
    for table in doc.tables:
        contenido.append("\n[FORMATO DE TABLA DETECTADO - DEBES LLENARLA]")
        for row in table.rows:
            fila_texto = " | ".join([cell.text.replace('\n', ' ') for cell in row.cells])
            contenido.append(fila_texto)
    return '\n'.join(contenido)

def generar_docx(texto_limpio):
    doc = Document()
    for linea in texto_limpio.split('\n'):
        doc.add_paragraph(linea)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def transcribir_audio(audio_file):
    r = sr.Recognizer()
    with sr.AudioFile(audio_file) as source:
        try: return r.recognize_google(r.record(source), language="es-MX")
        except: return "[No se detectó voz clara]"

# Variables de memoria para no perder datos
if 'reporte_generado' not in st.session_state: st.session_state.reporte_generado = ""

# 4. PANEL LATERAL: CONFIGURACIÓN Y APRENDIZAJE
with st.sidebar:
    st.markdown("### ⚕️ Beam AI Pro")
    st.caption("Estación de Alta Especialidad")
    
    try:
        api_key = st.secrets["deepseek_key"]
        st.success("Motor DeepSeek: Activo 🟢")
    except:
        api_key = st.text_input("DeepSeek API Key", type="password")
    
    st.divider()
    st.markdown("📁 **Gestor de Plantillas**")
    archivo_plantilla = st.file_uploader("Sube tu machote (.docx)", type=["docx"])
    plantilla_contenido = leer_word_con_tablas(archivo_plantilla) if archivo_plantilla else ""
    
    st.markdown("🧠 **Memoria del Radiólogo**")
    reglas_usuario = st.text_area(
        "Instrucciones permanentes:", 
        height=180, 
        value="1. Mantener precisión absoluta en descripción de articulaciones cigapofisarias y recesos.\n2. Evitar gerundios.\n3. Agrupar la patología degenerativa en la conclusión.",
        help="El sistema leerá esto en cada reporte para adaptarse a tu estilo."
    )

# 5. ESPACIO DE TRABAJO PRINCIPAL
if api_key:
    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com", timeout=60.0)
    
    st.markdown("## ⚡ Redacción Asistida")
    col_input, col_output = st.columns([1, 1.2], gap="large")

    with col_input:
        st.markdown("#### 🎙️ Entrada Clínica")
        mod_col, sub_col = st.columns(2)
        with mod_col:
            modalidad = st.selectbox("Modalidad", ["Tomografía", "Resonancia", "Radiografía", "Ultrasonido", "PET-CT"])
        with sub_col:
            estilo = st.selectbox("Estilo", ["Académico", "Conciso", "Institucional"])
        
        audio_file = st.audio_input("Dictado de Hallazgos")
        notas_texto = st.text_area("Hallazgos Manuales / Ajustes:", height=300, placeholder="Ej. L4-L5 con osteocondrosis y artrosis facetaria bilateral...")

        if st.button("✨ Procesar e Integrar Informe"):
            texto_dictado = transcribir_audio(audio_file) if audio_file else ""
            if not notas_texto and not texto_dictado:
                st.warning("⚠️ Faltan datos clínicos para procesar.")
            else:
                prompt_sistema = f"""
                Eres un Médico Radiólogo de Alta Especialidad. Tu tarea es redactar un informe estructurado de {modalidad} en estilo {estilo}.
                - ESTRUCTURA CLÍNICA: Técnica, Hallazgos, Impresión Diagnóstica.
                - INSTRUCCIONES DEL MEDICO (ESTRICTAS): {reglas_usuario}
                - PLANTILLA INSTITUCIONAL BASE: {plantilla_contenido if plantilla_contenido else 'Ninguna'}
                - TABLAS: Si la plantilla incluye tablas separadas por '|', DEBES rellenarlas manteniendo ese formato.
                - FIRMA OBLIGATORIA: Al final del documento, debes estampar exactamente esta firma:
                  **ATENTAMENTE**
                  DR. IGNACIO F. ALEJANDRO LÓPEZ BELTRÁN
                  MÉDICO ESP. EN IMAGENOLOGÍA DIAGNÓSTICA Y TERAPÉUTICA.
                """
                
                with st.spinner("🧠 Integrando hallazgos anatómicos y generando estructura..."):
                    try:
                        response = client.chat.completions.create(
                            model="deepseek-chat",
                            messages=[
                                {"role": "system", "content": prompt_sistema}, 
                                {"role": "user", "content": f"Redacta el informe con base en esto. Texto transcrito: {texto_dictado} | Notas añadidas: {notas_texto}"}
                            ],
                            temperature=0.2
                        )
                        st.session_state.reporte_generado = response.choices[0].message.content
                    except Exception as e: 
                        st.error(f"❌ Error de servidor: {str(e)}")

    with col_output:
        st.markdown("#### 📄 Informe Consolidado")
        texto_final = st.text_area("Revisión final:", value=st.session_state.reporte_generado, height=580)
        
        if st.session_state.reporte_generado:
            col_dl, col_ref = st.columns(2)
            with col_dl:
                st.download_button(
                    label="📥 Exportar Documento Word", 
                    data=generar_docx(texto_final), 
                    file_name="Reporte_Radiologico.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with col_ref:
                if st.button("🔄 Mejorar Conclusión"):
                    with st.spinner("Refinando la síntesis diagnóstica..."):
                        try:
                            prompt_ref = f"Reescribe y mejora ÚNICAMENTE la Impresión Diagnóstica del siguiente reporte para hacerla más estructurada y jerárquica. No alteres los Hallazgos ni la firma. REPORTE:\n{texto_final}"
                            response_ref = client.chat.completions.create(
                                model="deepseek-chat", 
                                messages=[{"role": "system", "content": "Experto en síntesis radiológica."}, {"role": "user", "content": prompt_ref}], 
                                temperature=0.3
                            )
                            st.session_state.reporte_generado = response_ref.choices[0].message.content
                            st.rerun()
                        except Exception as e: 
                            st.error("Error de conexión al refinar.")
else: 
    st.info("👈 Ingresa la credencial API para iniciar la estación.")
