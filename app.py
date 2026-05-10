import streamlit as st
from docx import Document
import speech_recognition as sr
import io
from openai import OpenAI

# ==========================================
# 1. MOTOR UI/UX (Estética Perplexity / Vercel)
# ==========================================
st.set_page_config(page_title="Lumen Core AI | Multimodal", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    
    /* Tema oscuro profundo y tipografía fluida */
    .stApp { background-color: #09090b; color: #ededed; font-family: 'Inter', sans-serif; }
    header { visibility: hidden; }
    
    /* Panel lateral de controles */
    [data-testid="stSidebar"] { background-color: #0c0c0f; border-right: 1px solid #1f1f22; }
    
    /* Cajas de herramientas y selectores */
    .stSelectbox div[data-baseweb="select"], .stFileUploader > div { 
        background-color: #121214 !important; border: 1px solid #27272a !important; border-radius: 8px !important; 
    }
    
    /* Botones de acción */
    .stButton>button { 
        background-color: #ededed !important; color: #09090b !important; 
        border-radius: 8px !important; font-weight: 600 !important; border: none !important; transition: 0.2s;
    }
    .stButton>button:hover { background-color: #a1a1aa !important; }
    
    /* Interfaz de Chat centralizada */
    [data-testid="stChatInput"] { 
        background-color: #18181b !important; border: 1px solid #27272a !important; 
        border-radius: 16px !important; padding: 0.5rem !important; 
    }
    [data-testid="stChatInput"] textarea { color: #ededed !important; }
    
    /* Burbujas de IA */
    [data-testid="stChatMessage"]:nth-child(even) { background-color: #121214 !important; border-radius: 12px; border: 1px solid #1f1f22; }
    [data-testid="stChatMessage"] { background-color: transparent !important; border: none !important; }
    
    hr { border-color: #27272a !important; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. FUNCIONES DE PROCESAMIENTO (El Cerebro)
# ==========================================
def leer_word_multimodal(file):
    doc = Document(file)
    contenido = []
    for para in doc.paragraphs:
        if para.text.strip(): contenido.append(para.text)
    for table in doc.tables:
        contenido.append("\n[INSTRUCCIÓN INTERNA: FORMATO DE TABLA DETECTADO. RESPETA ESTA CUADRÍCULA]")
        for row in table.rows:
            fila = " | ".join([cell.text.replace('\n', ' ') for cell in row.cells])
            contenido.append(fila)
    return '\n'.join(contenido)

def generar_docx(texto):
    doc = Document()
    for linea in texto.split('\n'):
        doc.add_paragraph(linea)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def transcribir_audio(audio_file):
    r = sr.Recognizer()
    with sr.AudioFile(audio_file) as source:
        try: return r.recognize_google(r.record(source), language="es-MX")
        except: return "[No se detectó voz clara. Verifica el micrófono o dicta de nuevo.]"

# ==========================================
# 3. ESTADOS DE MEMORIA GLOBAL
# ==========================================
if "mensajes" not in st.session_state:
    st.session_state.mensajes = [{"role": "assistant", "content": "Bienvenido a Lumen Core. La plataforma multimodal está lista. Sube tu plantilla, usa el micrófono para dictar, o escribe directamente tus instrucciones."}]
if "ultimo_reporte" not in st.session_state:
    st.session_state.ultimo_reporte = ""

# ==========================================
# 4. PANEL LATERAL: CONTROLES MULTIMODALES
# ==========================================
with st.sidebar:
    st.markdown("### 🧬 Lumen Core AI")
    st.caption("Motor Generativo Radiológico")
    
    try:
        api_key = st.secrets["deepseek_key"]
        st.success("🟢 API Conectada")
    except:
        api_key = st.text_input("DeepSeek API Key", type="password")
        
    st.divider()
    st.markdown("**1. Parámetros Clínicos**")
    modalidad = st.selectbox("Modalidad", [
        "Resonancia Magnética", "Tomografía Computarizada", "Radiografía Convencional", 
        "Ultrasonido", "PET-CT", "Mastografía", "Fluoroscopía", "Medicina Nuclear", "General"
    ])
    estilo = st.selectbox("Perfil de Redacción", [
        "Estándar Internacional", "Académico Detallado", "Conciso y Directo", "Alta Especialidad"
    ])
    
    st.divider()
    st.markdown("**2. Documento Base**")
    archivo_plantilla = st.file_uploader("Subir Plantilla (.docx)", type=["docx"])
    plantilla_procesada = leer_word_multimodal(archivo_plantilla) if archivo_plantilla else ""
    
    st.divider()
    st.markdown("**3. Entrada de Voz**")
    audio_dictado = st.audio_input("Grabar hallazgos")

# ==========================================
# 5. WORKSPACE CENTRAL (Chat & Exportación)
# ==========================================
st.title("Generación de Informes Clínicos")

# Botón de exportación dinámico (flota arriba cuando hay un reporte listo)
if st.session_state.ultimo_reporte:
    st.download_button(
        label="📥 Exportar Informe a Word (.docx)", 
        data=generar_docx(st.session_state.ultimo_reporte), 
        file_name="Informe_Radiologico_Lumen.docx", 
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.markdown("---")

# Renderizar el historial de interacción
for msg in st.session_state.mensajes:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# ==========================================
# 6. MOTOR DE INFERENCIA Y CHAT
# ==========================================
if api_key:
    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com", timeout=60.0)
    
    # El usuario puede interactuar escribiendo o enviando el audio grabado
    prompt_texto = st.chat_input("Escribe instrucciones o presiona Enter para procesar el audio grabado...")
    
    if prompt_texto or audio_dictado:
        texto_transcrito = transcribir_audio(audio_dictado) if audio_dictado else ""
        
        # Consolidar la entrada del usuario (Voz + Texto)
        entrada_consolidada = ""
        if texto_transcrito: entrada_consolidada += f"🎙️ **Voz transcrita:** {texto_transcrito}\n\n"
        if prompt_texto: entrada_consolidada += f"⌨️ **Instrucción:** {prompt_texto}"
        
        # Mostrar en pantalla lo que el usuario envió
        st.session_state.mensajes.append({"role": "user", "content": entrada_consolidada})
        with st.chat_message("user"):
            st.markdown(entrada_consolidada)
            
        # Generar respuesta de la IA
        with st.chat_message("assistant"):
            with st.spinner("Procesando datos multimodales y estructurando informe..."):
                
                # El Prompt Maestro de Sistema
                instruccion_sistema = f"""
                Eres Lumen, un modelo de Inteligencia Artificial Generativa de clase mundial especializado en radiología.
                Tarea: Crear o modificar un informe radiológico de {modalidad}.
                Estilo de redacción: {estilo}.
                
                REGLAS ESTRICTAS:
                1. Mantén un tono médico formal, objetivo y preciso.
                2. Estructura básica obligatoria: Técnica, Hallazgos, Impresión Diagnóstica (salvo que la plantilla indique otra cosa).
                3. Si el usuario provee instrucciones de corrección (ej. "mejora la conclusión"), aplica el cambio manteniendo el resto del reporte intacto.
                4. NO inventes nombres de pacientes ni médicos. Usa marcadores genéricos (ej. [Nombre del Paciente]) si es necesario.
                
                PLANTILLA A RESPETAR:
                {plantilla_procesada if plantilla_procesada else "No se proporcionó plantilla. Usa formato estándar."}
                """
                
                try:
                    response = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[
                            {"role": "system", "content": instruccion_sistema},
                            {"role": "user", "content": entrada_consolidada}
                        ],
                        temperature=0.2
                    )
                    respuesta_ia = response.choices[0].message.content
                    st.markdown(respuesta_ia)
                    
                    # Guardar en memoria
                    st.session_state.mensajes.append({"role": "assistant", "content": respuesta_ia})
                    st.session_state.ultimo_reporte = respuesta_ia
                    
                    # Recargar suavemente para que aparezca el botón de exportación arriba
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ Error de procesamiento: {e}")
else:
    st.info("👈 Ingresa tu API Key en el panel lateral para activar el motor generativo.")
