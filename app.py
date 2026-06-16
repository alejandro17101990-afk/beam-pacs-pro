import streamlit as st

# 1. CONFIGURACIÓN DE LA PÁGINA
# Esto hace que la app ocupe todo el ancho de la pantalla
st.set_page_config(layout="wide", page_title="Beam AI")

# --- COLUMNA IZQUIERDA (BARRA LATERAL) ---
with st.sidebar:
    st.markdown("### 🟣 BEAM AI")
    
    if st.button("+ Nuevo informe", use_container_width=True, type="primary"):
        pass # Aquí irá la lógica para limpiar el editor
        
    st.write("---")
    
    st.markdown("**ESTUDIOS RECIENTES**")
    st.button("TC Abdomen y Pelvis\n12/05/2025", use_container_width=True)
    st.button("RM Rodilla Derecha\n11/05/2025", use_container_width=True)
    st.button("TC Tórax sin contraste\n10/05/2025", use_container_width=True)


# --- COLUMNA CENTRAL (ÁREA PRINCIPAL) ---

# Encabezado (Título y botones superiores)
col_titulo, col_botones = st.columns([3, 1]) # El título toma más espacio que los botones

with col_titulo:
    st.markdown("#### TC ABDOMEN Y PELVIS CON CONTRASTE")
    st.caption("ID: EST-2025-0805678 • 12 Mayo 2025, 10:30 AM")

with col_botones:
    # Contenedor para alinear los botones a la derecha
    btn1, btn2 = st.columns(2)
    with btn1:
        st.button("🎙️ Dictado")
    with btn2:
        # Usamos type="primary" para que resalte como el botón de IA en tu diseño
        st.button("✨ IA activa", type="primary")

st.divider() # Línea separadora

# Área del Editor de Texto
st.markdown("**HALLAZGOS**")

# Texto por defecto (tu plantilla)
texto_inicial = """Hígado: De tamaño, forma y contornos normales, con atenuación homogénea. No se identifican lesiones focales. Vía biliar intra y extrahepática no dilatada.

Vesícula biliar: De paredes delgadas, sin litiasis radiopacas.

Páncreas: De aspecto normal, sin lesiones focales ni dilatación del conducto pancreático.

Bazo: De tamaño normal, con atenuación homogénea, sin lesiones focales.

Riñones: De tamaño y morfología conservada, con adecuada captación y eliminación del contraste. No se observan litiasis ni hidronefrosis.

CONCLUSIÓN
1. Estudio sin hallazgos tomográficos sugestivos de patología abdominal o pélvica aguda.
2. Hallazgos descritos."""

# El text_area de Streamlit funcionará como tu editor principal
reporte = st.text_area(
    "Editor", 
    value=texto_inicial, 
    height=500, # Altura del cuadro de texto
    label_visibility="collapsed" # Ocultamos la etiqueta "Editor" para que se vea más limpio
)

# Botones inferiores (Sugerir conclusión, etc.)
col_inf1, col_inf2, col_inf3, col_inf4 = st.columns([2, 2, 2, 4])

with col_inf1:
    st.button("+ Sugerir conclusión")
with col_inf2:
    st.selectbox("Estilo", ["Estilo académico", "Conservador", "Directo"], label_visibility="collapsed")
with col_inf3:
    st.selectbox("Destinatario", ["Para médico tratante", "Para paciente"], label_visibility="collapsed")
with col_inf4:
    st.markdown("<p style='text-align: right; color: gray; margin-top: 10px;'>✓ Guardado 10:30 AM</p>", unsafe_allow_html=True)
