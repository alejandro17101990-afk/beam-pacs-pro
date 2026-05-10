import streamlit as st
from docx import Document
import io
from openai import OpenAI

# 1. CONFIGURACIÓN MODO VERCEL/PERPLEXITY
st.set_page_config(page_title="Lumen PACS | Next Gen", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    /* Colores y fondo oscuro absoluto tipo Vercel */
    .stApp { background-color: #09090b; color: #ededed; font-family: 'Inter', sans-serif; }
    header { visibility: hidden; } /* Esconder la barra superior por defecto */
    
    /* Sidebar minimalista */
    [data-testid="stSidebar"] { background-color: #0c0c0f; border-right: 1px solid #1f1f22; }
    
    /* Botones primarios limpios (Shadcn UI style) */
    .stButton>button { 
        background-color: #ededed !important; color: #09090b !important; 
        border-radius: 8px !important; font-weight: 500 !important; transition: all 0.2s; border: none !important;
    }
    .stButton>button:hover { background-color: #d4d4d8 !important; }
    
    /* Barra de entrada de chat tipo Perplexity */
    [data-testid="stChatInput"] { background-color: #18181b !important; border: 1px solid #27272a !important; border-radius: 16px !important; }
    [data-testid="stChatInput"] textarea { color: #ededed !important; }
    
    /* Mensajes de IA */
    [data-testid="stChatMessage"] { background-color: transparent !important; border: none !important; }
    div[data-testid="stChatMessage"]:nth-child(even) { background-color: #121214 !important; border-radius: 12px; }
    
    hr { border-color: #27272a !important; }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNCIONES DE EXPORTACIÓN
def generar_docx(texto):
    doc = Document()
    for linea in texto.split('\n'):
        doc.add_paragraph(linea)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# 3. INTERFAZ LATERAL (Configuración Oculta)
with st.sidebar:
    st.markdown("### ⚕️ Lumen PACS")
    st.caption("Arquitectura V1.0")
    
    try:
        api_key = st.secrets["deepseek_key"]
        st.success("🟢 Conexión Segura")
    except:
        api_key = st.text_input("DeepSeek API Key", type="password")
        
    st.divider()
    modalidad = st.selectbox("Modalidad Activa", ["RM Rodilla", "TC Lumbar", "Radiografía", "Ultrasonido", "PET-CT"])
    estilo = st.selectbox("Tono Clínico", ["Conciso y Directo", "Académico Detallado"])
    
    st.divider()
    st.caption("Si usas plantilla, pégala aquí:")
    plantilla = st.text_area("Formato Base:", height=100)

# 4. MEMORIA DEL CHAT
if "mensajes" not in st.session_state:
    st.session_state.mensajes = [{"role": "assistant", "content": "Hola, Dr. López Beltrán. Estación lista. ¿Qué hallazgos procesamos hoy?"}]
if "ultimo_reporte" not in st.session_state:
    st.session_state.ultimo_reporte = ""

# 5. RENDERIZADO DEL CHAT EN PANTALLA
st.title("Generación Asistida")
st.caption("Dicta tus hallazgos en lenguaje natural y la IA estructurará el informe.")

# Botón de exportación rápido (solo aparece si hay un reporte)
if st.session_state.ultimo_reporte:
    st.download_button(
        label="📥 Descargar Último Reporte a Word", 
        data=generar_docx(st.session_state.ultimo_reporte), 
        file_name="Reporte_Lumen_PACS.docx", 
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.markdown("---")

# Mostrar historial
for msg in st.session_state.mensajes:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# 6. CAJA DE ENTRADA INFERIOR (Estilo Perplexity)
if api_key:
    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com", timeout=60.0)
    
    prompt = st.chat_input("🎙️ Escribe o dicta tus hallazgos aquí (Ej. Desgarro del menisco medial, grado 2...)")
    
    if prompt:
        # Agregar lo que el usuario escribió
        st.session_state.mensajes.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
            
        # Generar respuesta de la IA
        with st.chat_message("assistant"):
            with st.spinner("Procesando estructura clínica..."):
                instruccion_sistema = f"""
                Eres Lumen, un asistente de IA radiológica. 
                Genera un informe médico de {modalidad} en estilo {estilo}.
                Estructura: Técnica, Hallazgos, Impresión Diagnóstica.
                Plantilla a respetar: {plantilla if plantilla else 'Ninguna'}.
                """
                
                try:
                    response = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[
                            {"role": "system", "content": instruccion_sistema},
                            {"role": "user", "content": prompt}
                        ],
                        temperature=0.2
                    )
                    respuesta_ia = response.choices[0].message.content
                    st.markdown(respuesta_ia)
                    st.session_state.mensajes.append({"role": "assistant", "content": respuesta_ia})
                    st.session_state.ultimo_reporte = respuesta_ia
                    st.rerun() # Recarga la página suavemente para mostrar el botón de Word
                except Exception as e:
                    st.error(f"Error de red: {e}")
else:
    st.info("Ingresa tu API Key en el panel izquierdo para iniciar.")
