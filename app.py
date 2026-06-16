import streamlit as st
from streamlit_quill import st_quill

st.set_page_config(
    page_title="Beam AI",
    page_icon="🩻",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ======================================
# ESTADO
# ======================================

if "show_left" not in st.session_state:
    st.session_state.show_left = True

if "show_right" not in st.session_state:
    st.session_state.show_right = True

# ======================================
# CSS
# ======================================

st.markdown("""
<style>

html, body, .stApp{
    background:#0b0f14;
    color:white;
}

header,
footer,
#MainMenu{
    visibility:hidden;
}

.block-container{
    max-width:100%;
    padding:0.8rem 1rem;
}

/* TOPBAR */

.topbar{
    background:#11151c;
    border:1px solid #1f242d;
    border-radius:18px;
    height:62px;
    padding:0 25px;

    display:flex;
    align-items:center;
    justify-content:space-between;

    margin-bottom:15px;
}

.logo{
    font-size:24px;
    font-weight:700;
}

.status{
    color:#8b949e;
    font-size:14px;
}

/* CARDS */

.card{
    background:#11151c;
    border:1px solid #1f242d;
    border-radius:20px;
    padding:20px;
}

/* BOTONES */

.stButton>button{
    border:none;
    border-radius:12px;
    background:#171b22;
    color:white;
}

.stButton>button:hover{
    background:#222831;
}

/* EXPANDERS */

.streamlit-expanderHeader{
    font-size:14px;
}

/* FOOTER */

.footer{
    text-align:center;
    color:#7d8590;
    font-size:13px;
    margin-top:10px;
}

</style>
""", unsafe_allow_html=True)

# ======================================
# TOPBAR
# ======================================

st.markdown("""
<div class="topbar">

<div class="logo">
🩻 BEAM AI
</div>

<div class="status">
● Autoguardado
</div>

</div>
""", unsafe_allow_html=True)

# ======================================
# CONTROLES
# ======================================

c1, c2, c3, c4 = st.columns([1,1,6,1])

with c1:

    if st.button("☰"):

        st.session_state.show_left = (
            not st.session_state.show_left
        )

with c2:

    if st.button("🤖"):

        st.session_state.show_right = (
            not st.session_state.show_right
        )

# ======================================
# LAYOUT DINÁMICO
# ======================================

if (
    st.session_state.show_left
    and
    st.session_state.show_right
):

    left, center, right = st.columns(
        [1.1, 5, 1.5]
    )

elif st.session_state.show_left:

    left, center = st.columns(
        [1.2, 6]
    )

    right = None

elif st.session_state.show_right:

    center, right = st.columns(
        [6, 1.5]
    )

    left = None

else:

    center = st.container()
    left = None
    right = None

# ======================================
# SIDEBAR
# ======================================

if left:

    with left:

        st.markdown(
            "<div class='card'>",
            unsafe_allow_html=True
        )

        st.subheader("Navegación")

        st.button("➕ Nuevo")
        st.button("🕒 Recientes")
        st.button("📄 Plantillas")
        st.button("⭐ Favoritos")
        st.button("⚙ Configuración")

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

# ======================================
# EDITOR
# ======================================

with center:

    st.markdown(
        """
        <div class='card'>
        <h3>Informe Radiológico</h3>
        </div>
        """,
        unsafe_allow_html=True
    )

    a1, a2, a3, a4, a5 = st.columns(5)

    with a1:
        st.button("Plantilla")

    with a2:
        st.button("Optimizar")

    with a3:
        st.button("Conclusión")

    with a4:
        st.button("Clasificaciones")

    with a5:
        st.button("Exportar")

    informe = st_quill(
        value="""
<h2>TÉCNICA</h2>
<p></p>

<h2>HALLAZGOS</h2>
<p></p>

<h2>CONCLUSIÓN</h2>
<p></p>
""",
        html=True,
        toolbar=[
            ['bold', 'italic', 'underline'],
            [{'header':[1,2,3,False]}],
            [{'list':'ordered'},
             {'list':'bullet'}],
            ['clean']
        ]
    )

# ======================================
# IA
# ======================================

if right:

    with right:

        st.markdown(
            "<div class='card'>",
            unsafe_allow_html=True
        )

        st.subheader("Asistente IA")

        with st.expander(
            "Sugerencias"
        ):
            st.write(
                "Las sugerencias aparecerán aquí."
            )

        with st.expander(
            "Clasificaciones"
        ):
            st.write(
                "Stoller, BI-RADS, TIRADS..."
            )

        with st.expander(
            "Conclusiones"
        ):
            st.write(
                "Conclusiones sugeridas."
            )

        with st.expander(
            "Calidad"
        ):
            st.success(
                "Sin omisiones detectadas."
            )

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

# ======================================
# FOOTER
# ======================================

st.markdown("""
<div class="footer">
✓ Autoguardado • IA preparada • Calidad del informe: Excelente
</div>
""", unsafe_allow_html=True)
