"""
AURA v4 — Radiology Copilot
Arquitectura enterprise: 3 paneles, aprendizaje de estilo, QA contextual, KB modular.
"""

import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Pt, RGBColor
import speech_recognition as sr
import io
import json
import re
from openai import OpenAI

# ══════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="AURA · Radiology Copilot",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ══════════════════════════════════════════════════════════════
# BASE DE CONOCIMIENTO MODULAR
# ══════════════════════════════════════════════════════════════

KB = {
    "clasificaciones": {
        "Menisco · Stoller": {
            "descripcion": "Clasificación RM de degeneración/desgarro meniscal",
            "grados": {
                "I": "Señal focal intrameniscal globular o puntiforme, no articular. Sin implicación clínica directa.",
                "II": "Señal lineal intrameniscal horizontal, no alcanza superficie articular. Cambio degenerativo.",
                "III": "Señal anormal alcanza una o ambas superficies articulares → DESGARRO confirmado.",
            },
            "clave": "/stoller",
            "referencia": "Stoller DW. Magnetic Resonance Imaging in Orthopaedics and Sports Medicine. 2007.",
        },
        "Cartílago · ICRS": {
            "descripcion": "International Cartilage Repair Society — clasificación de condropatía",
            "grados": {
                "I": "Fibrilación superficial o fisuras. <50% grosor.",
                "II": "Lesión que alcanza hasta el 50% del grosor condral.",
                "III": "Lesión >50% grosor condral, sin exposición ósea.",
                "IV": "Pérdida completa del cartílago, hueso subcondral expuesto.",
            },
            "clave": "/icrs",
            "referencia": "Brittberg M, et al. Clin Orthop Relat Res. 2001.",
        },
        "Artrosis · Kellgren-Lawrence": {
            "descripcion": "Clasificación radiográfica de osteoartrosis",
            "grados": {
                "0": "Sin cambios osteoartrósicos.",
                "I": "Osteofito posible; dudoso estrechamiento articular.",
                "II": "Osteofito definido; posible estrechamiento del espacio articular.",
                "III": "Osteofitos moderados; estrechamiento moderado; esclerosis subcondral.",
                "IV": "Osteofitos grandes; marcado estrechamiento; esclerosis severa; deformidad ósea.",
            },
            "clave": "/kl",
            "referencia": "Kellgren JH, Lawrence JS. Ann Rheum Dis. 1957.",
        },
        "LCA · Hope & Feagin": {
            "descripcion": "Clasificación RM de lesión de ligamento cruzado anterior",
            "grados": {
                "Parcial": "Fibras continuas con señal aumentada. Ruptura <50% fibras.",
                "Completa": "Discontinuidad total de fibras. Edema periligamentoso.",
                "Crónica": "Fibras atróficas o ausentes. Cicatrización anormal.",
            },
            "clave": "/lca",
            "referencia": "Hope M, Feagin JA. Am J Sports Med. 1995.",
        },
        "Disco · Pfirrmann": {
            "descripcion": "Clasificación RM de degeneración discal en columna",
            "grados": {
                "I": "Núcleo pulposo brillante, homogéneo, señal alta en T2. Altura normal.",
                "II": "Señal ligeramente reducida, zona fibrosa no clara. Altura normal.",
                "III": "Señal gris, distinción pulpo-anular borrosa. Altura normal o ligeramente reducida.",
                "IV": "Señal baja, sin distinción pulpo-anular. Altura normal o moderadamente reducida.",
                "V": "Señal muy baja. Espacio discal colapsado.",
            },
            "clave": "/pfirrmann",
            "referencia": "Pfirrmann CW, et al. Spine. 2001.",
        },
        "Tiroides · ACR TIRADS": {
            "descripcion": "ACR TI-RADS — clasificación ecográfica de nódulo tiroideo",
            "grados": {
                "1": "Normal. Sin nódulos.",
                "2": "Benigno. Sin componentes sospechosos.",
                "3": "Levemente sospechoso. Seguimiento en 1-3 años.",
                "4": "Moderadamente sospechoso. BAAF si ≥1.5 cm.",
                "5": "Altamente sospechoso. BAAF si ≥1 cm.",
            },
            "clave": "/tirads",
            "referencia": "Tessler FN, et al. J Am Coll Radiol. 2017.",
        },
        "Mama · BI-RADS": {
            "descripcion": "ACR BI-RADS — clasificación de hallazgos mamarios",
            "grados": {
                "0": "Evaluación incompleta. Requiere imágenes adicionales.",
                "1": "Negativo. Control rutinario.",
                "2": "Benigno. Control rutinario.",
                "3": "Probablemente benigno. Control en 6 meses.",
                "4A": "Baja sospecha de malignidad. Biopsia sugerida.",
                "4B": "Sospecha intermedia. Biopsia indicada.",
                "4C": "Sospecha moderada-alta. Biopsia indicada.",
                "5": "Altamente sugestivo de malignidad. Biopsia obligatoria.",
                "6": "Malignidad histológicamente probada.",
            },
            "clave": "/birads",
            "referencia": "D'Orsi CJ, et al. ACR BI-RADS Atlas, 5th Ed. 2013.",
        },
        "Columna · Modic": {
            "descripcion": "Cambios de señal en médula ósea vertebral adyacente a discos",
            "grados": {
                "I": "Hipointenso T1 / Hiperintenso T2. Edema/inflamación activa.",
                "II": "Hiperintenso T1 / Hiperintenso T2. Sustitución grasa.",
                "III": "Hipointenso T1 / Hipointenso T2. Esclerosis ósea.",
            },
            "clave": "/modic",
            "referencia": "Modic MT, et al. Radiology. 1988.",
        },
        "Hígado · LI-RADS": {
            "descripcion": "Liver Imaging Reporting and Data System — HCC en pacientes de alto riesgo",
            "grados": {
                "LR-1": "Definitivamente benigno.",
                "LR-2": "Probablemente benigno.",
                "LR-3": "Indeterminado.",
                "LR-4": "Probablemente HCC.",
                "LR-5": "Definitivamente HCC.",
                "LR-M": "Probablemente/definitivamente maligno, no específico de HCC.",
                "LR-TIV": "Invasión tumoral venosa.",
            },
            "clave": "/lirads",
            "referencia": "ACR LI-RADS v2018.",
        },
        "Nódulo pulmonar · Fleischner": {
            "descripcion": "Guías de seguimiento de nódulos pulmonares incidentales",
            "grados": {
                "<6mm sólido bajo riesgo": "Sin seguimiento rutinario.",
                "<6mm sólido alto riesgo": "TC opcional a 12 meses.",
                "6-8mm sólido": "TC a 6-12 meses, luego 18-24 meses.",
                ">8mm sólido": "TC a 3 meses / PET-CT / biopsia.",
                "Subsólido puro <6mm": "Sin seguimiento.",
                "Subsólido puro ≥6mm": "TC a 6-12 meses para confirmar persistencia.",
            },
            "clave": "/fleischner",
            "referencia": "MacMahon H, et al. Radiology. 2017.",
        },
    },

    "anatomia_regiones": {
        "Rodilla": {
            "estructuras": ["menisco medial", "menisco lateral", "LCA", "LCP", "LCM", "LCL",
                           "cartílago femoral", "cartílago tibial", "rótula", "tendón rotuliano",
                           "tendón cuadricipital", "bursas", "cuerpos libres", "grasa de Hoffa"],
            "clasificaciones_relevantes": ["Menisco · Stoller", "Cartílago · ICRS", "Artrosis · Kellgren-Lawrence", "LCA · Hope & Feagin"],
            "hallazgos_frecuentes": ["desgarro meniscal", "condropatía", "edema óseo", "osteoartrosis", "lesión ligamentosa"],
            "omisiones_criticas": ["menisco medial", "menisco lateral", "LCA", "cartílago articular"],
        },
        "Columna lumbar": {
            "estructuras": ["discos intervertebrales L1-S1", "núcleo pulposo", "anillo fibroso",
                           "facetas articulares", "foramina", "canal espinal", "ligamento amarillo",
                           "médula/cono medular", "raíces nerviosas", "platillos vertebrales"],
            "clasificaciones_relevantes": ["Disco · Pfirrmann", "Columna · Modic", "Artrosis · Kellgren-Lawrence"],
            "hallazgos_frecuentes": ["hernia discal", "estenosis foraminal", "estenosis del canal", "espondiloartrosis", "listesis"],
            "omisiones_criticas": ["discos intervertebrales", "canal espinal", "foramina", "facetas"],
        },
        "Hombro": {
            "estructuras": ["manguito rotador (supraespinoso, infraespinoso, subescapular, redondo menor)",
                           "tendón del bíceps (porción larga)", "articulación glenohumeral",
                           "articulación acromioclavicular", "bursa subacromial-subdeltoidea",
                           "espacio subacromial", "labrum glenoideo", "cápsula articular"],
            "clasificaciones_relevantes": ["Cartílago · ICRS"],
            "hallazgos_frecuentes": ["desgarro del manguito", "tendinosis", "bursitis", "impingement", "lesión SLAP", "Bankart"],
            "omisiones_criticas": ["supraespinoso", "infraespinoso", "subescapular", "bursa subacromial", "labrum"],
        },
        "Columna cervical": {
            "estructuras": ["discos C2-C7", "uncovertebral", "facetas articulares", "foramina",
                           "canal espinal", "médula espinal", "ligamento amarillo", "apófisis odontoides"],
            "clasificaciones_relevantes": ["Disco · Pfirrmann", "Columna · Modic"],
            "hallazgos_frecuentes": ["hernia discal", "uncoartrosis", "mielopatía", "estenosis foraminal"],
            "omisiones_criticas": ["discos intervertebrales", "médula espinal", "canal cervical"],
        },
        "Cadera": {
            "estructuras": ["cabeza femoral", "acetábulo", "labrum acetabular", "cartílago articular",
                           "tendón iliopsoas", "glúteos", "trocánter mayor", "espacio articular"],
            "clasificaciones_relevantes": ["Cartílago · ICRS", "Artrosis · Kellgren-Lawrence"],
            "hallazgos_frecuentes": ["coxartrosis", "desgarro del labrum", "impingement femoroacetabular", "necrosis avascular"],
            "omisiones_criticas": ["labrum acetabular", "cartílago articular", "espacio articular"],
        },
        "Tobillo / Pie": {
            "estructuras": ["tendón de Aquiles", "tendón tibial posterior", "tendones peroneos",
                           "ligamentos laterales (LTFA, LCF, LTFP)", "ligamento deltoideo",
                           "sindesmosis", "cartílago talar", "seno del tarso"],
            "clasificaciones_relevantes": ["Cartílago · ICRS"],
            "hallazgos_frecuentes": ["rotura de Aquiles", "tendinosis", "esguince ligamentoso", "osteoartrosis"],
            "omisiones_criticas": ["tendón de Aquiles", "ligamentos laterales", "cartílago talar"],
        },
        "Cerebro": {
            "estructuras": ["parénquima cerebral", "sistema ventricular", "surcos y cisuras",
                           "sustancia blanca", "sustancia gris", "ganglios basales", "tálamos",
                           "cerebelo", "tronco encefálico", "estructuras de línea media",
                           "espacio subaracnoideo", "senos venosos"],
            "clasificaciones_relevantes": [],
            "hallazgos_frecuentes": ["lesión isquémica", "hemorragia", "neoplasia", "leucoaraiosis", "hidrocefalia"],
            "omisiones_criticas": ["parénquima", "sistema ventricular", "estructuras de línea media", "sustancia blanca"],
        },
        "Tórax": {
            "estructuras": ["parénquima pulmonar", "árbol bronquial", "pleuras", "mediastino",
                           "estructuras vasculares", "esófago", "pared torácica", "diafragma",
                           "estructuras óseas costales y vertebrales"],
            "clasificaciones_relevantes": ["Nódulo pulmonar · Fleischner"],
            "hallazgos_frecuentes": ["consolidación", "derrame pleural", "nódulo pulmonar", "adenopatías", "ensanchamiento mediastinal"],
            "omisiones_criticas": ["parénquima pulmonar", "pleuras", "mediastino"],
        },
        "Abdomen / Pelvis": {
            "estructuras": ["hígado", "vía biliar", "vesícula", "páncreas", "bazo", "riñones",
                           "suprarrenales", "aorta abdominal", "vena cava", "mesenterio",
                           "asas intestinales", "vejiga", "útero/próstata"],
            "clasificaciones_relevantes": ["Hígado · LI-RADS"],
            "hallazgos_frecuentes": ["lesión focal hepática", "litiasis biliar", "pancreatitis", "adenopatías", "masa pélvica"],
            "omisiones_criticas": ["hígado", "riñones", "páncreas", "aorta abdominal"],
        },
        "Mama": {
            "estructuras": ["tejido fibroglandular", "grasa subcutánea", "piel y pezón",
                           "vasos linfáticos", "ganglios axilares", "implantes (si aplica)"],
            "clasificaciones_relevantes": ["Mama · BI-RADS"],
            "hallazgos_frecuentes": ["nódulo", "microcalcificaciones", "asimetría", "distorsión arquitectural", "adenopatías axilares"],
            "omisiones_criticas": ["composición tisular", "BI-RADS", "ganglios axilares"],
        },
        "Tiroides": {
            "estructuras": ["lóbulo derecho", "lóbulo izquierdo", "istmo",
                           "nódulos", "adenopatías cervicales", "tráquea", "vascularización"],
            "clasificaciones_relevantes": ["Tiroides · ACR TIRADS"],
            "hallazgos_frecuentes": ["nódulo tiroideo", "bocio multinodular", "tiroiditis", "adenopatías"],
            "omisiones_criticas": ["ACR TIRADS", "vascularización", "adenopatías cervicales"],
        },
    },

    "terminologia_correcta": {
        "cambios degenerativos": "osteoartrosis / degeneración discal / tendinosis (especificar)",
        "desgarro": "desgarro (NO 'rotura' para menisco/ligamento en contexto MSK español)",
        "ruptura": "desgarro (meniscos, ligamentos) / rotura (Aquiles completa)",
        "artritis": "osteoartrosis (para artropatía degenerativa)",
        "edema": "edema óseo subcondral / edema medular óseo (especificar)",
        "lesión": "especificar tipo: desgarro / contusión / fractura / neoplasia",
        "alteración": "especificar naturaleza: señal aumentada / disminuida / heterogénea",
        "cambios": "especificar: cambios degenerativos tipo / cambios postquirúrgicos",
        "normal": "sin alteraciones morfológicas ni de señal relevantes (no usar 'normal' solo)",
        "masa": "especificar: lesión focal / masa / proceso expansivo + características",
    },

    "plantillas_informe": {
        "Resonancia Magnética": {
            "secciones": ["TÉCNICA", "HALLAZGOS", "IMPRESIÓN DIAGNÓSTICA"],
            "tecnica_default": "Estudio de resonancia magnética de {region} en equipo de {tesla} Tesla. Secuencias multiplanares en {secuencias}. {contraste}.",
            "secuencias_por_region": {
                "Rodilla": "DP con supresión grasa (DPFS), T1, T2 y STIR en planos coronal, sagital y axial",
                "Columna lumbar": "T1, T2 sagital y axial, con secuencias STIR",
                "Hombro": "DP-FS, T1, T2 y ABER en planos coronal oblicuo, sagital oblicuo y axial",
                "Cerebro": "T1, T2, FLAIR, DWI/ADC, T2*-GRE en planos axial, coronal y sagital",
                "default": "T1, T2 y secuencias con supresión grasa multiplanares",
            },
        },
        "Tomografía Computarizada": {
            "secciones": ["TÉCNICA", "HALLAZGOS", "IMPRESIÓN DIAGNÓSTICA"],
            "tecnica_default": "Estudio de tomografía computarizada de {region}. {contraste}. Reconstrucciones multiplanares.",
        },
        "Radiografía": {
            "secciones": ["TÉCNICA", "HALLAZGOS", "IMPRESIÓN DIAGNÓSTICA"],
            "tecnica_default": "Radiografía de {region} en proyecciones {proyecciones}.",
        },
        "Ultrasonido": {
            "secciones": ["TÉCNICA", "HALLAZGOS", "IMPRESIÓN DIAGNÓSTICA"],
            "tecnica_default": "Estudio ecográfico de {region} con transductor {transductor}. Evaluación con Doppler color.",
        },
    },
}

# ══════════════════════════════════════════════════════════════
# ESTADO DE SESIÓN
# ══════════════════════════════════════════════════════════════
defaults = {
    "dictado": "",
    "reporte_html": "",
    "reporte_texto": "",
    "copilot_panel": "",
    "copilot_tipo": "",
    "historial_reportes": [],
    "estilo_usuario": {
        "ejemplos": [],
        "preferencias": "",
        "terminologia_propia": [],
    },
    "qa_resultado": {},
    "modelo_activo": "deepseek-chat",
    "api_provider": "deepseek",
    "panel_derecho_visible": True,
    "panel_izquierdo_visible": True,
    "clasificaciones_activas": {},
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ══════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════

def get_client():
    key = st.session_state.get("api_key_input", "")
    try:
        key = key or st.secrets.get("deepseek_key", "") or st.secrets.get("openai_key", "")
    except Exception:
        pass
    if not key:
        return None, None
    provider = st.session_state.api_provider
    if provider == "deepseek":
        return OpenAI(api_key=key, base_url="https://api.deepseek.com"), "deepseek-chat"
    elif provider == "openai_mini":
        return OpenAI(api_key=key), "gpt-4o-mini"
    elif provider == "openai_4":
        return OpenAI(api_key=key), "gpt-4.1-mini"
    return OpenAI(api_key=key, base_url="https://api.deepseek.com"), "deepseek-chat"

def leer_plantilla(file):
    doc = Document(file)
    textos = []
    for p in doc.paragraphs:
        if p.text.strip():
            textos.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            row_text = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
            if row_text:
                textos.append(row_text)
    return "\n".join(textos)

def transcribir_voz(audio_file):
    r = sr.Recognizer()
    try:
        with sr.AudioFile(audio_file) as source:
            audio = r.record(source)
            return r.recognize_google(audio, language="es-MX")
    except Exception:
        return ""

def calcular_qa(texto, region):
    """QA modular: evalúa completitud, terminología y estructura."""
    resultado = {
        "score": 0,
        "secciones": {},
        "omisiones": [],
        "terminologia": [],
        "sugerencias": [],
    }
    texto_upper = texto.upper()
    secciones = ["TÉCNICA", "HALLAZGOS", "IMPRESIÓN"]
    encontradas = 0
    for s in secciones:
        presente = s in texto_upper
        resultado["secciones"][s] = presente
        if presente:
            encontradas += 1

    palabras = len(texto.split())
    score_secciones = (encontradas / 3) * 40
    score_palabras = min(palabras / 200, 1) * 30

    # Verificar omisiones anatómicas
    info_region = KB["anatomia_regiones"].get(region, {})
    omisiones_criticas = info_region.get("omisiones_criticas", [])
    omisiones_encontradas = []
    for estructura in omisiones_criticas:
        if estructura.lower() not in texto.lower():
            omisiones_encontradas.append(estructura)
    resultado["omisiones"] = omisiones_encontradas
    score_anatomia = max(0, (1 - len(omisiones_encontradas) / max(len(omisiones_criticas), 1))) * 20

    # Verificar terminología problemática
    for termino_malo, termino_correcto in KB["terminologia_correcta"].items():
        if termino_malo.lower() in texto.lower():
            resultado["terminologia"].append({
                "incorrecto": termino_malo,
                "correcto": termino_correcto
            })

    # Verificar clasificaciones
    clasif_relevantes = info_region.get("clasificaciones_relevantes", [])
    if clasif_relevantes:
        clasif_score = sum(1 for c in clasif_relevantes
                          if any(w in texto for w in c.split(" · "))) / max(len(clasif_relevantes), 1)
        score_clasif = clasif_score * 10
    else:
        score_clasif = 10

    resultado["score"] = min(100, int(score_secciones + score_palabras + score_anatomia + score_clasif))
    return resultado

def texto_a_html(texto):
    """Convierte texto plano estructurado a HTML para el editor."""
    html_lines = []
    for line in texto.split("\n"):
        s = line.strip()
        if not s:
            html_lines.append("<br>")
        elif s.isupper() and len(s) < 80 and not s.startswith("•"):
            html_lines.append(f'<div class="section-title">{s}</div>')
        elif s.startswith("•") or s.startswith("-"):
            html_lines.append(f'<li>{s[1:].strip()}</li>')
        else:
            html_lines.append(f'<p>{s}</p>')
    return "\n".join(html_lines)

def generar_docx(html_texto, reporte_texto):
    """Genera DOCX con estilos radiológicos profesionales."""
    from html.parser import HTMLParser

    class HTMLtoDocx(HTMLParser):
        def __init__(self):
            super().__init__()
            self.doc = Document()
            style = self.doc.styles["Normal"]
            style.font.name = "Arial"
            style.font.size = Pt(11)
            self.bold = self.italic = self.underline = False
            self.current_para = None
            self.in_li = False
            self.in_section_title = False

        def handle_starttag(self, tag, attrs):
            attr_dict = dict(attrs)
            cls = attr_dict.get("class", "")
            if tag in ("b", "strong") or "section-title" in cls:
                self.bold = True
                if tag == "div" and "section-title" in cls:
                    self.current_para = self.doc.add_paragraph()
                    self.in_section_title = True
            elif tag in ("i", "em"):
                self.italic = True
            elif tag == "u":
                self.underline = True
            elif tag in ("p", "div"):
                self.current_para = self.doc.add_paragraph()
            elif tag == "li":
                self.current_para = self.doc.add_paragraph(style="List Bullet")
                self.in_li = True
            elif tag == "br":
                if self.current_para is None:
                    self.current_para = self.doc.add_paragraph()

        def handle_endtag(self, tag):
            if tag in ("b", "strong"):
                self.bold = False
            elif tag in ("i", "em"):
                self.italic = False
            elif tag == "u":
                self.underline = False
            elif tag == "li":
                self.in_li = False
                self.current_para = None
            elif tag == "div":
                self.bold = False
                self.in_section_title = False

        def handle_data(self, data):
            text = data.strip()
            if not text:
                return
            if self.current_para is None:
                self.current_para = self.doc.add_paragraph()
            run = self.current_para.add_run(text + " ")
            run.bold = self.bold
            run.italic = self.italic
            run.underline = self.underline
            if self.in_section_title:
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(0x00, 0x4C, 0x97)

    clean = (html_texto or reporte_texto).replace("\n", " ").strip()
    parser = HTMLtoDocx()
    try:
        parser.feed(clean)
    except Exception:
        doc = Document()
        plain = re.sub(r"<[^>]+>", "", html_texto or reporte_texto)
        for line in plain.split("\n"):
            doc.add_paragraph(line)
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    bio = io.BytesIO()
    parser.doc.save(bio)
    return bio.getvalue()

def construir_sistema_prompt(modalidad, region, instrucciones, plantilla_txt, clasif_activas, estilo_usuario):
    """Construye el system prompt con todo el contexto del KB."""
    clasif_relevantes = KB["anatomia_regiones"].get(region, {}).get("clasificaciones_relevantes", [])
    info_clasif = ""
    for nombre_clasif in clasif_relevantes:
        datos = KB["clasificaciones"].get(nombre_clasif, {})
        if datos:
            info_clasif += f"\n**{nombre_clasif}**: {datos.get('descripcion','')}\n"
            for grado, desc in datos.get("grados", {}).items():
                info_clasif += f"  Grado {grado}: {desc}\n"

    estructuras = KB["anatomia_regiones"].get(region, {}).get("estructuras", [])
    terminologia_mala = list(KB["terminologia_correcta"].keys())

    estilo_ctx = ""
    if estilo_usuario.get("ejemplos"):
        estilo_ctx = f"\nESTILO DEL RADIÓLOGO (aprender y replicar):\n"
        for i, ej in enumerate(estilo_usuario["ejemplos"][-3:], 1):
            estilo_ctx += f"Ejemplo {i}: {ej[:500]}...\n"
    if estilo_usuario.get("preferencias"):
        estilo_ctx += f"\nPreferencias declaradas: {estilo_usuario['preferencias']}\n"

    clasif_ctx = ""
    if clasif_activas:
        clasif_ctx = "\nCLASIFICACIONES ACTIVAS DEL CASO:\n"
        for k, v in clasif_activas.items():
            clasif_ctx += f"• {k}: Grado {v}\n"

    return f"""Eres AURA, el copiloto de redacción radiológica más avanzado disponible. Eres un radiólogo experto de subespecialidad con 20 años de experiencia.

MODALIDAD: {modalidad}
REGIÓN ANATÓMICA: {region}

CLASIFICACIONES RADIOLÓGICAS APLICABLES:
{info_clasif}

ESTRUCTURAS ANATÓMICAS A EVALUAR EN {region.upper()}:
{", ".join(estructuras)}

{clasif_ctx}

TERMINOLOGÍA PROHIBIDA (usar alternativas):
{", ".join(terminologia_mala)}
Razón: son términos vagos. Siempre especificar tipo, grado y localización.

{estilo_ctx}

DIRECTRICES CLÍNICAS:
{instrucciones}

PLANTILLA BASE:
{plantilla_txt if plantilla_txt else "TÉCNICA / HALLAZGOS / IMPRESIÓN DIAGNÓSTICA"}

REGLAS DE FORMATO:
1. Títulos de sección en MAYÚSCULAS y negritas
2. Subtítulos anatómicos en MAYÚSCULAS
3. Sin asteriscos ni markdown
4. Conclusiones en viñetas con "•", concluyentes y clínicamente accionables
5. Incluir grado exacto de cada clasificación usada
6. Agregar recomendación de manejo cuando el grado lo amerite
7. Redacción narrativa fluida, no listados en hallazgos
8. Lenguaje médico formal en español de México (no España)"""

# ══════════════════════════════════════════════════════════════
# ESTILOS CSS GLOBALES
# ══════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
    --bg-0: #06080f;
    --bg-1: #090c15;
    --bg-2: #0c1020;
    --bg-3: #0f1428;
    --bg-4: #121830;
    --border: #1a2540;
    --border-bright: #1e2e50;
    --accent: #00aaff;
    --accent-dim: #005a99;
    --accent-glow: rgba(0,170,255,0.12);
    --text-primary: #c8daf0;
    --text-secondary: #6a8aaf;
    --text-dim: #3a5070;
    --text-bright: #e8f4ff;
    --success: #00cc88;
    --warning: #f0a020;
    --danger: #ff4060;
    --mono: 'JetBrains Mono', monospace;
    --sans: 'Inter', sans-serif;
}

html, body, .stApp { background: var(--bg-0) !important; font-family: var(--sans) !important; }
.block-container { padding: 0 !important; max-width: 100% !important; }
header, footer, [data-testid="stToolbar"] { display: none !important; }
section[data-testid="stSidebar"] { display: none !important; }

/* ── TOPBAR ── */
.aura-topbar {
    background: var(--bg-1);
    border-bottom: 1px solid var(--border);
    padding: 0 20px;
    height: 48px;
    display: flex;
    align-items: center;
    gap: 14px;
    position: sticky;
    top: 0;
    z-index: 9999;
    backdrop-filter: blur(12px);
}
.aura-logo {
    font-family: var(--sans);
    font-weight: 700;
    font-size: 15px;
    color: var(--text-bright);
    letter-spacing: 0.06em;
    display: flex;
    align-items: center;
    gap: 8px;
}
.logo-pulse {
    width: 8px; height: 8px;
    border-radius: 50%;
    background: var(--accent);
    box-shadow: 0 0 8px var(--accent);
    animation: pulse 2.5s ease-in-out infinite;
}
@keyframes pulse {
    0%, 100% { opacity: 1; box-shadow: 0 0 8px var(--accent); }
    50% { opacity: 0.5; box-shadow: 0 0 3px var(--accent); }
}
.tb-sep { width: 1px; height: 20px; background: var(--border); }
.tb-badge {
    font-family: var(--mono);
    font-size: 10px;
    color: var(--accent-dim);
    background: rgba(0,100,180,0.1);
    border: 1px solid rgba(0,100,180,0.2);
    padding: 2px 8px;
    border-radius: 3px;
    letter-spacing: 0.04em;
}
.tb-badge.active {
    color: var(--accent);
    background: var(--accent-glow);
    border-color: rgba(0,170,255,0.25);
}
.tb-status {
    margin-left: auto;
    display: flex;
    align-items: center;
    gap: 6px;
    font-family: var(--mono);
    font-size: 10px;
    color: var(--success);
}
.status-dot {
    width: 5px; height: 5px;
    border-radius: 50%;
    background: var(--success);
    box-shadow: 0 0 5px var(--success);
}

/* ── PANEL LAYOUT ── */
.aura-layout {
    display: flex;
    height: calc(100vh - 48px);
    overflow: hidden;
}

/* ── LABELS ── */
.plabel {
    font-family: var(--mono) !important;
    font-size: 9px !important;
    letter-spacing: 0.18em !important;
    color: var(--text-dim) !important;
    text-transform: uppercase !important;
    display: block;
    margin-bottom: 4px !important;
    margin-top: 10px !important;
}

/* ── SELECTBOX ── */
[data-testid="stSelectbox"] > div > div {
    background: var(--bg-2) !important;
    border: 1px solid var(--border) !important;
    border-radius: 6px !important;
    color: var(--text-secondary) !important;
    font-family: var(--mono) !important;
    font-size: 11px !important;
}
[data-testid="stSelectbox"] > div > div:hover { border-color: var(--border-bright) !important; }
[data-testid="stSelectbox"] svg { fill: var(--text-dim) !important; }

/* ── TEXTAREA ── */
.stTextArea textarea {
    background: var(--bg-2) !important;
    border: 1px solid var(--border) !important;
    border-radius: 6px !important;
    color: var(--text-primary) !important;
    font-family: var(--mono) !important;
    font-size: 11px !important;
    line-height: 1.7 !important;
    resize: none !important;
}
.stTextArea textarea:focus {
    border-color: var(--accent-dim) !important;
    box-shadow: 0 0 0 1px rgba(0,100,180,0.2) !important;
}

/* ── AUDIO INPUT ── */
[data-testid="stAudioInput"] {
    background: var(--bg-2) !important;
    border: 1px solid var(--border) !important;
    border-radius: 6px !important;
}

/* ── FILE UPLOADER ── */
[data-testid="stFileUploader"] {
    background: var(--bg-2) !important;
    border: 1px dashed var(--border) !important;
    border-radius: 6px !important;
}
[data-testid="stFileUploader"] * {
    color: var(--text-dim) !important;
    font-family: var(--mono) !important;
    font-size: 10px !important;
}

/* ── BOTONES ── */
.stButton > button {
    background: var(--bg-2) !important;
    border: 1px solid var(--border) !important;
    color: var(--text-secondary) !important;
    font-family: var(--sans) !important;
    font-size: 11px !important;
    font-weight: 500 !important;
    border-radius: 6px !important;
    padding: 6px 12px !important;
    transition: all 0.15s !important;
    letter-spacing: 0.02em !important;
}
.stButton > button:hover {
    background: var(--bg-3) !important;
    border-color: var(--border-bright) !important;
    color: var(--text-primary) !important;
}

/* BOTÓN PRIMARIO */
.btn-primary .stButton > button {
    background: linear-gradient(135deg, #003a78 0%, #005aaa 100%) !important;
    border: 1px solid #007acc !important;
    color: #7ad4ff !important;
    font-weight: 600 !important;
    font-size: 12px !important;
    letter-spacing: 0.05em !important;
}
.btn-primary .stButton > button:hover {
    background: linear-gradient(135deg, #004a90 0%, #006acc 100%) !important;
    border-color: #00aaff !important;
    color: #aae8ff !important;
    box-shadow: 0 0 14px rgba(0,120,220,0.3) !important;
}

/* BOTÓN PELIGRO */
.btn-danger .stButton > button {
    color: #cc4455 !important;
    border-color: #3a1520 !important;
}
.btn-danger .stButton > button:hover {
    background: #1a0810 !important;
    color: #ff5566 !important;
    border-color: #cc4455 !important;
}

/* ── EXPANDER ── */
[data-testid="stExpander"] {
    background: transparent !important;
    border: none !important;
    border-bottom: 1px solid var(--border) !important;
    border-radius: 0 !important;
    margin-bottom: 0 !important;
}
[data-testid="stExpander"] summary {
    color: var(--text-dim) !important;
    font-family: var(--mono) !important;
    font-size: 9px !important;
    letter-spacing: 0.15em !important;
    text-transform: uppercase !important;
    padding: 10px 0 !important;
}
[data-testid="stExpander"] summary:hover { color: var(--text-secondary) !important; }
[data-testid="stExpander"] summary svg { color: var(--text-dim) !important; }
[data-testid="stExpander"] > div > div { padding: 4px 0 12px 0 !important; }

/* ── DOWNLOAD BUTTON ── */
[data-testid="stDownloadButton"] > button {
    background: var(--bg-2) !important;
    border: 1px solid var(--border) !important;
    color: var(--text-secondary) !important;
    font-family: var(--mono) !important;
    font-size: 10px !important;
    border-radius: 6px !important;
}
[data-testid="stDownloadButton"] > button:hover {
    border-color: var(--accent-dim) !important;
    color: var(--accent) !important;
}

/* ── RADIO BUTTONS ── */
[data-testid="stRadio"] > div { gap: 6px !important; }
[data-testid="stRadio"] label {
    background: var(--bg-2) !important;
    border: 1px solid var(--border) !important;
    border-radius: 5px !important;
    padding: 4px 10px !important;
    font-family: var(--mono) !important;
    font-size: 10px !important;
    color: var(--text-dim) !important;
}
[data-testid="stRadio"] label:has(input:checked) {
    border-color: var(--accent-dim) !important;
    color: var(--accent) !important;
    background: var(--accent-glow) !important;
}

/* ── SCROLLBAR ── */
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 2px; }
::-webkit-scrollbar-thumb:hover { background: var(--border-bright); }

/* ── DIVIDER ── */
hr { border-color: var(--border) !important; margin: 8px 0 !important; }

/* ── PANEL LATERAL DERECHO ── */
.copilot-panel {
    background: var(--bg-1);
    border-left: 1px solid var(--border);
    overflow-y: auto;
    padding: 12px;
    height: 100%;
}
.copilot-block {
    background: var(--bg-2);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 12px;
    margin-bottom: 10px;
}
.copilot-block-title {
    font-family: var(--mono);
    font-size: 9px;
    color: var(--text-dim);
    letter-spacing: 0.15em;
    text-transform: uppercase;
    margin-bottom: 8px;
    display: flex;
    align-items: center;
    gap: 6px;
}
.copilot-block-title::before {
    content: '';
    width: 4px; height: 4px;
    border-radius: 50%;
    background: var(--accent);
    display: inline-block;
}
.qa-item {
    display: flex;
    align-items: flex-start;
    gap: 6px;
    padding: 5px 0;
    border-bottom: 1px solid var(--border);
    font-size: 11px;
    font-family: var(--mono);
}
.qa-ok { color: var(--success); }
.qa-warn { color: var(--warning); }
.qa-err { color: var(--danger); }
.clasif-chip {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    background: rgba(0,80,150,0.12);
    border: 1px solid rgba(0,100,180,0.2);
    border-radius: 4px;
    padding: 3px 8px;
    font-family: var(--mono);
    font-size: 10px;
    color: #4a90cc;
    margin: 2px;
    cursor: pointer;
    transition: all 0.15s;
}
.clasif-chip:hover {
    background: rgba(0,100,200,0.2);
    border-color: rgba(0,150,255,0.3);
    color: #7ac2f0;
}
.section-badge {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 2px 7px;
    border-radius: 3px;
    font-family: var(--mono);
    font-size: 9px;
    letter-spacing: 0.08em;
}
.badge-ok { background: rgba(0,200,100,0.08); border: 1px solid rgba(0,200,100,0.15); color: var(--success); }
.badge-err { background: rgba(255,60,80,0.08); border: 1px solid rgba(255,60,80,0.15); color: var(--danger); }

/* ── SCORE RING ── */
.score-container {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 12px;
}
.score-number {
    font-family: var(--sans);
    font-weight: 700;
    font-size: 32px;
    color: var(--text-bright);
    line-height: 1;
}
.score-label {
    font-family: var(--mono);
    font-size: 9px;
    color: var(--text-dim);
    letter-spacing: 0.1em;
    text-transform: uppercase;
}
.score-bar {
    flex: 1;
    height: 2px;
    background: var(--bg-4);
    border-radius: 1px;
    overflow: hidden;
    margin-top: 6px;
}
.score-fill {
    height: 100%;
    border-radius: 1px;
    transition: width 0.5s ease;
}

/* HISTORIAL */
.hist-item {
    padding: 8px 10px;
    border-radius: 5px;
    border: 1px solid var(--border);
    margin-bottom: 4px;
    cursor: pointer;
    transition: all 0.15s;
    background: var(--bg-2);
}
.hist-item:hover { border-color: var(--border-bright); background: var(--bg-3); }
.hist-title { font-family: var(--sans); font-size: 11px; color: var(--text-secondary); font-weight: 500; }
.hist-meta { font-family: var(--mono); font-size: 9px; color: var(--text-dim); margin-top: 2px; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# TOPBAR
# ══════════════════════════════════════════════════════════════
proveedor_label = {
    "deepseek": "DeepSeek",
    "openai_mini": "GPT-4o Mini",
    "openai_4": "GPT-4.1 Mini",
}.get(st.session_state.api_provider, "DeepSeek")

st.markdown(f"""
<div class="aura-topbar">
    <div class="aura-logo">
        <div class="logo-pulse"></div>
        AURA
    </div>
    <div class="tb-sep"></div>
    <span class="tb-badge">Radiology Copilot v4</span>
    <span class="tb-badge active">MSK · Neuro · TX · Abd</span>
    <div class="tb-sep"></div>
    <div class="tb-status">
        <div class="status-dot"></div>
        {proveedor_label} · activo
    </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# LAYOUT PRINCIPAL — 3 columnas
# ══════════════════════════════════════════════════════════════
col_izq, col_centro, col_der = st.columns([1, 2.4, 1], gap="small")

# ══════════════════════════════════════════════════════════════
# PANEL IZQUIERDO — Entrada clínica
# ══════════════════════════════════════════════════════════════
with col_izq:
    st.markdown("<div style='padding:12px 4px 0 4px'>", unsafe_allow_html=True)

    # API & Modelo
    with st.expander("CONEXIÓN", expanded=False):
        api_key_input = st.text_input(
            "API Key", type="password", key="api_key_input",
            label_visibility="collapsed",
            placeholder="API Key (DeepSeek / OpenAI)"
        )
        st.markdown('<span class="plabel">MODELO</span>', unsafe_allow_html=True)
        modelo = st.radio(
            "Modelo", ["deepseek", "openai_mini", "openai_4"],
            format_func=lambda x: {"deepseek": "DeepSeek Chat", "openai_mini": "GPT-4o Mini", "openai_4": "GPT-4.1 Mini"}[x],
            key="api_provider", horizontal=False, label_visibility="collapsed"
        )

    # Modalidad y región
    with st.expander("ESTUDIO", expanded=True):
        st.markdown('<span class="plabel">MODALIDAD</span>', unsafe_allow_html=True)
        modalidad = st.selectbox(
            "Modalidad",
            ["Resonancia Magnética", "Tomografía Computarizada", "Radiografía", "Ultrasonido", "PET-CT", "Fluoroscopía / Intervencionismo"],
            label_visibility="collapsed"
        )
        st.markdown('<span class="plabel">REGIÓN ANATÓMICA</span>', unsafe_allow_html=True)
        regiones_disponibles = list(KB["anatomia_regiones"].keys()) + ["Muñeca / Mano", "Codo", "Hígado", "Mama", "Tiroides"]
        region = st.selectbox("Región", sorted(set(regiones_disponibles)), label_visibility="collapsed")

        # Indicar clasificaciones disponibles para esta región
        clasif_region = KB["anatomia_regiones"].get(region, {}).get("clasificaciones_relevantes", [])
        if clasif_region:
            chips_html = "".join([f'<span class="clasif-chip" title="{KB["clasificaciones"].get(c,{}).get("descripcion","")}">{c}</span>' for c in clasif_region])
            st.markdown(f'<div style="margin-top:6px">{chips_html}</div>', unsafe_allow_html=True)

    # Dictado de voz
    with st.expander("VOZ & DICTADO", expanded=True):
        audio_data = st.audio_input("Grabar", label_visibility="collapsed")
        if audio_data:
            nuevo = transcribir_voz(audio_data)
            if nuevo and nuevo not in st.session_state.dictado:
                st.session_state.dictado += (" " if st.session_state.dictado else "") + nuevo
                st.rerun()

        st.markdown('<span class="plabel">TRANSCRIPCIÓN / ENTRADA MANUAL</span>', unsafe_allow_html=True)
        dictado = st.text_area(
            "Dictado", value=st.session_state.dictado,
            height=150, label_visibility="collapsed",
            placeholder="Dicta o escribe los hallazgos...\n\nEj: Rodilla derecha. Stoller III menisco medial cuerpo y cuerno posterior. Extrusión 3 mm. ICRS III platillo tibial medial..."
        )
        if dictado != st.session_state.dictado:
            st.session_state.dictado = dictado

        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="btn-danger">', unsafe_allow_html=True)
            if st.button("Purgar dictado", use_container_width=True):
                st.session_state.dictado = ""
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        with c2:
            if st.button("Limpiar informe", use_container_width=True):
                st.session_state.reporte_html = ""
                st.session_state.reporte_texto = ""
                st.session_state.qa_resultado = {}
                st.rerun()

    # Clasificaciones activas del caso
    with st.expander("CLASIFICACIONES DEL CASO", expanded=False):
        st.markdown('<span class="plabel">ACTIVAR PARA ESTE ESTUDIO</span>', unsafe_allow_html=True)
        clasif_disponibles = list(KB["clasificaciones"].keys())
        clasif_sel = st.selectbox("Clasificación", clasif_disponibles, label_visibility="collapsed", key="clasif_sel")

        if clasif_sel:
            datos_clasif = KB["clasificaciones"][clasif_sel]
            grados = list(datos_clasif["grados"].keys())
            grado_sel = st.selectbox("Grado", grados, label_visibility="collapsed", key="grado_sel")

            col_a, col_b = st.columns([3, 1])
            with col_a:
                if grado_sel:
                    desc = datos_clasif["grados"][grado_sel]
                    st.markdown(f'<div style="font-family:var(--mono);font-size:10px;color:#5a8aaf;line-height:1.5;padding:4px 0">{desc}</div>', unsafe_allow_html=True)
            with col_b:
                if st.button("+ Agregar"):
                    st.session_state.clasificaciones_activas[clasif_sel] = grado_sel
                    st.rerun()

        # Mostrar activas
        if st.session_state.clasificaciones_activas:
            st.markdown('<span class="plabel" style="margin-top:10px">ACTIVAS EN ESTE CASO</span>', unsafe_allow_html=True)
            for k, v in list(st.session_state.clasificaciones_activas.items()):
                c1, c2 = st.columns([4, 1])
                with c1:
                    st.markdown(f'<div style="font-family:var(--mono);font-size:10px;color:#4a8abf;padding:3px 0">{k}: <b style="color:#7ac2f0">Grado {v}</b></div>', unsafe_allow_html=True)
                with c2:
                    if st.button("×", key=f"del_{k}"):
                        del st.session_state.clasificaciones_activas[k]
                        st.rerun()

    # Estilo personal
    with st.expander("MI ESTILO DE REDACCIÓN", expanded=False):
        st.markdown('<span class="plabel">EJEMPLO DE MI INFORME</span>', unsafe_allow_html=True)
        ejemplo_estilo = st.text_area(
            "Ejemplo", height=80, label_visibility="collapsed",
            placeholder="Pega aquí un informe tuyo anterior para que AURA aprenda tu estilo..."
        )
        if st.button("Aprender mi estilo"):
            if ejemplo_estilo.strip():
                st.session_state.estilo_usuario["ejemplos"].append(ejemplo_estilo.strip())
                st.success(f"✓ Ejemplo {len(st.session_state.estilo_usuario['ejemplos'])} guardado")

        st.markdown('<span class="plabel">PREFERENCIAS DECLARADAS</span>', unsafe_allow_html=True)
        pref = st.text_area(
            "Prefs", height=60, label_visibility="collapsed",
            value=st.session_state.estilo_usuario.get("preferencias", ""),
            placeholder="Ej: Siempre empezar por meniscos. Usa 'se observa'. Evita listas en hallazgos."
        )
        if pref != st.session_state.estilo_usuario.get("preferencias", ""):
            st.session_state.estilo_usuario["preferencias"] = pref

        st.markdown(f'<div style="font-family:var(--mono);font-size:10px;color:#2a5a4a;margin-top:6px">{len(st.session_state.estilo_usuario["ejemplos"])} ejemplo(s) aprendido(s)</div>', unsafe_allow_html=True)

    # Configuración
    with st.expander("CONFIGURACIÓN", expanded=False):
        st.markdown('<span class="plabel">PLANTILLA BASE (.docx)</span>', unsafe_allow_html=True)
        archivo_base = st.file_uploader("Plantilla", type=["docx"], label_visibility="collapsed")
        plantilla_txt = leer_plantilla(archivo_base) if archivo_base else ""

        st.markdown('<span class="plabel">DIRECTRICES DE ESTILO</span>', unsafe_allow_html=True)
        instrucciones = st.text_area(
            "Directrices", height=60, label_visibility="collapsed",
            value="Lenguaje médico experto. Expandir clasificaciones. Sin abreviaturas sin definir. Impresión diagnóstica concluyente y accionable."
        )

    # ── Botón principal ──
    st.markdown("<div style='padding:12px 0 6px 0'>", unsafe_allow_html=True)
    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    procesar = st.button("⬡  GENERAR INFORME AURA", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Historial
    if st.session_state.historial_reportes:
        with st.expander(f"HISTORIAL ({len(st.session_state.historial_reportes)})", expanded=False):
            for i, h in enumerate(reversed(st.session_state.historial_reportes[-5:])):
                st.markdown(f"""
                <div class="hist-item" onclick="">
                    <div class="hist-title">{h.get('region','?')} · {h.get('modalidad','?')[:2]}</div>
                    <div class="hist-meta">{h.get('timestamp','')[:16]} · {h.get('palabras',0)} palabras</div>
                </div>
                """, unsafe_allow_html=True)
                if st.button("Recuperar", key=f"rec_{i}"):
                    st.session_state.reporte_texto = h.get("texto", "")
                    st.session_state.reporte_html = texto_a_html(h.get("texto", ""))
                    st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# PROCESAMIENTO IA
# ══════════════════════════════════════════════════════════════
if procesar:
    client, model_name = get_client()
    if client and dictado.strip():
        sistema = construir_sistema_prompt(
            modalidad, region, instrucciones, plantilla_txt,
            st.session_state.clasificaciones_activas,
            st.session_state.estilo_usuario
        )
        with st.spinner(""):
            try:
                res = client.chat.completions.create(
                    model=model_name,
                    messages=[
                        {"role": "system", "content": sistema},
                        {"role": "user", "content": f"DICTADO DEL RADIÓLOGO:\n{dictado}"}
                    ],
                    temperature=0.1
                )
                texto = res.choices[0].message.content
                st.session_state.reporte_texto = texto
                st.session_state.reporte_html = texto_a_html(texto)
                st.session_state.qa_resultado = calcular_qa(texto, region)

                # Guardar en historial
                import datetime
                st.session_state.historial_reportes.append({
                    "texto": texto,
                    "region": region,
                    "modalidad": modalidad,
                    "palabras": len(texto.split()),
                    "timestamp": datetime.datetime.now().isoformat(),
                    "score": st.session_state.qa_resultado.get("score", 0),
                })

                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
    elif not client:
        st.warning("Ingresa tu API Key en el panel de Conexión.")
    else:
        st.warning("Ingresa dictado o hallazgos clínicos.")

# ══════════════════════════════════════════════════════════════
# PANEL CENTRO — Editor rico
# ══════════════════════════════════════════════════════════════
with col_centro:
    contenido_inicial = st.session_state.reporte_html or """<div class="section-title">RESONANCIA MAGNÉTICA DE RODILLA DERECHA</div>
<br>
<div class="section-title">TÉCNICA</div>
<p>Estudio de resonancia magnética de rodilla derecha en equipo de 1.5 Tesla. Secuencias multiplanares en DP con supresión grasa (DPFS), T1, T2 y STIR en planos coronal, sagital y axial, sin administración de medio de contraste.</p>
<br>
<div class="section-title">HALLAZGOS</div>
<br>
<div class="section-title">MENISCOS</div>
<p>Menisco medial: alteración de señal grado III de Stoller en cuerpo y cuerno posterior, de orientación horizontal, que alcanza la superficie articular inferior, compatible con desgarro horizontal. Extrusión medial de 3 mm en plano coronal.</p>
<p>Menisco lateral: morfología e intensidad de señal conservadas. Sin evidencia de desgarro.</p>
<br>
<div class="section-title">LIGAMENTOS</div>
<p>Ligamento cruzado anterior con señal heterogénea en tercio proximal y pérdida focal de la arquitectura de las fibras, compatible con lesión parcial grado I de Hope &amp; Feagin. Ligamento cruzado posterior de señal y morfología normales.</p>
<br>
<div class="section-title">CARTÍLAGO ARTICULAR</div>
<p>Adelgazamiento condral focal grado III de ICRS en platillo tibial medial (extensión aproximada de 12 × 8 mm), con señal subcondral reactiva en secuencias STIR.</p>
<br>
<div class="section-title">IMPRESIÓN DIAGNÓSTICA</div>
<li>Desgarro horizontal de menisco medial grado III de Stoller, en cuerpo y cuerno posterior, con extrusión medial de 3 mm. Se recomienda valoración artroscópica.</li>
<li>Lesión parcial grado I del LCA según Hope &amp; Feagin. Control clínico y de imagen en 6 semanas.</li>
<li>Condropatía grado III de ICRS en compartimento medial con edema subcondral reactivo. Correlacionar con cuadro clínico para considerar manejo conservador o intervencionista.</li>"""

    editor_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root {{
    --bg-0:#06080f; --bg-1:#090c15; --bg-2:#0c1020; --bg-3:#0f1428;
    --border:#1a2540; --border-b:#1e2e50;
    --accent:#00aaff; --accent-dim:#005a99; --accent-glow:rgba(0,170,255,0.1);
    --text-p:#c8daf0; --text-s:#6a8aaf; --text-d:#3a5070; --text-b:#e8f4ff;
    --success:#00cc88; --warning:#f0a020; --danger:#ff4060;
    --mono:'JetBrains Mono',monospace; --sans:'Inter',sans-serif;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg-0);font-family:var(--sans);overflow:hidden;display:flex;flex-direction:column;height:100vh}}

/* FORMAT BAR */
.fbar{{
    background:var(--bg-1);border-bottom:1px solid var(--border);
    padding:6px 14px;display:flex;align-items:center;gap:4px;flex-wrap:wrap;
    position:sticky;top:0;z-index:100;
}}
.fg{{display:flex;align-items:center;gap:2px;padding-right:8px;border-right:1px solid var(--border)}}
.fg:last-of-type{{border-right:none;padding-right:0}}
.fb{{
    background:none;border:1px solid transparent;color:var(--text-s);
    font-size:11px;font-family:var(--sans);
    padding:3px 7px;border-radius:4px;cursor:pointer;transition:all .12s;
    min-width:26px;text-align:center;line-height:1.5;white-space:nowrap;
}}
.fb:hover{{background:var(--bg-2);border-color:var(--border);color:var(--text-p)}}
.fb.on{{background:var(--bg-3);border-color:var(--border-b);color:var(--accent)}}
.fsel{{
    background:var(--bg-2);border:1px solid var(--border);color:var(--text-s);
    font-size:10px;font-family:var(--mono);padding:3px 5px;border-radius:4px;
    outline:none;cursor:pointer;appearance:none;
}}
.fsel:focus{{border-color:var(--accent-dim)}}
.clr{{
    width:16px;height:16px;border-radius:50%;cursor:pointer;
    border:2px solid transparent;flex-shrink:0;transition:all .12s;
}}
.clr:hover,.clr.on{{border-color:var(--accent);box-shadow:0 0 6px rgba(0,170,255,0.4)}}
.flabel{{font-size:9px;color:var(--text-d);font-family:var(--mono);white-space:nowrap}}

/* EDITOR */
.editor-wrap{{flex:1;overflow-y:auto;padding:0}}
#doc{{
    min-height:100%;padding:28px 36px;
    font-family:var(--sans);font-size:14px;line-height:1.9;
    color:var(--text-p);outline:none;
    background:#0a1018;transition:background .3s,color .3s;
}}
#doc .section-title{{
    font-family:var(--sans);font-weight:700;font-size:13px;
    color:var(--text-b);letter-spacing:0.05em;
    margin:18px 0 6px 0;padding-bottom:4px;
    border-bottom:1px solid var(--border);
}}
#doc p{{margin-bottom:6px;color:var(--text-p)}}
#doc li{{
    margin-left:16px;margin-bottom:5px;color:var(--text-p);
    list-style-type:disc;padding-left:4px;
}}
#doc b,#doc strong{{color:var(--text-b)}}
#doc hr{{border:none;border-top:1px solid var(--border);margin:14px 0}}

/* ACTION BAR */
.abar{{
    background:var(--bg-1);border-top:1px solid var(--border);
    padding:7px 14px;display:flex;align-items:center;gap:6px;flex-wrap:wrap;
}}
.ab{{
    background:var(--bg-2);border:1px solid var(--border);color:var(--text-s);
    font-size:10px;font-family:var(--sans);font-weight:500;
    padding:5px 11px;border-radius:5px;cursor:pointer;
    transition:all .12s;display:flex;align-items:center;gap:5px;white-space:nowrap;
}}
.ab:hover{{color:var(--text-p);background:var(--bg-3);border-color:var(--border-b)}}
.ab.prime{{
    background:linear-gradient(135deg,#003a78,#005aaa);
    border-color:#0066bb;color:#7ad4ff;font-weight:600;
}}
.ab.prime:hover{{background:linear-gradient(135deg,#004a90,#0068cc);box-shadow:0 0 12px rgba(0,120,220,0.3)}}

/* SCORE PULSE */
.score-wrap{{margin-left:auto;display:flex;align-items:center;gap:10px}}
.score-val{{font-family:var(--sans);font-size:20px;font-weight:700;color:var(--text-b)}}
.score-lbl{{font-family:var(--mono);font-size:9px;color:var(--text-d);text-transform:uppercase;letter-spacing:.12em}}
.score-track{{width:60px;height:2px;background:var(--bg-3);border-radius:1px;overflow:hidden;margin-top:4px}}
.score-bar-f{{height:100%;border-radius:1px;transition:width .5s ease}}

/* SLASH AUTOCOMPLETE */
.slash-menu{{
    position:absolute;background:var(--bg-2);border:1px solid var(--border-b);
    border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,0.5);
    z-index:1000;min-width:240px;overflow:hidden;
}}
.slash-item{{
    padding:8px 12px;cursor:pointer;transition:background .1s;
    font-family:var(--mono);font-size:11px;color:var(--text-p);
    display:flex;align-items:center;gap:8px;
}}
.slash-item:hover,.slash-item.active{{background:var(--bg-3)}}
.slash-key{{color:var(--accent);font-weight:500;min-width:80px}}
.slash-desc{{color:var(--text-s);font-size:10px}}
</style>
</head>
<body>

<!-- FORMAT BAR -->
<div class="fbar">
  <div class="fg">
    <select class="fsel" id="fontSel" onchange="applyFont(this.value)" style="width:80px">
      <option value="Inter">Inter</option>
      <option value="Arial">Arial</option>
      <option value="Georgia">Georgia</option>
      <option value="'JetBrains Mono'">JB Mono</option>
    </select>
    <select class="fsel" id="sizeSel" onchange="applySize(this.value)" style="width:42px">
      <option>11</option><option>12</option><option>13</option>
      <option selected>14</option><option>15</option><option>16</option><option>18</option>
    </select>
  </div>
  <div class="fg">
    <button class="fb" id="btnB" onclick="fmt('bold')" title="Negrita (Ctrl+B)"><b>B</b></button>
    <button class="fb" id="btnI" onclick="fmt('italic')" title="Cursiva (Ctrl+I)"><i>I</i></button>
    <button class="fb" id="btnU" onclick="fmt('underline')" title="Subrayado (Ctrl+U)"><u>U</u></button>
  </div>
  <div class="fg">
    <button class="fb" onclick="fmt('justifyLeft')" title="Izquierda">≡</button>
    <button class="fb" onclick="fmt('justifyCenter')" title="Centro">≡</button>
    <button class="fb" onclick="fmt('justifyFull')" title="Justificado">≡</button>
  </div>
  <div class="fg">
    <button class="fb" onclick="fmt('insertUnorderedList')" title="Viñetas">• Lista</button>
    <button class="fb" onclick="insertHR()" title="Separador">— HR</button>
    <button class="fb" onclick="insertSecTitle()" title="Título sección">§ Sección</button>
  </div>
  <div class="fg" style="align-items:center;gap:6px">
    <span class="flabel">Fondo:</span>
    <div class="clr on" style="background:#0a1018" onclick="setBg(this,'#0a1018','#c8daf0','#e8f4ff')" title="Clínico"></div>
    <div class="clr" style="background:#f8f9fa;border:1px solid #bbb" onclick="setBg(this,'#f8f9fa','#1a2a3a','#0d1a28')" title="Blanco"></div>
    <div class="clr" style="background:#141008" onclick="setBg(this,'#141008','#ddd0b8','#f0e8d0')" title="Cálido"></div>
    <div class="clr" style="background:#000409" onclick="setBg(this,'#000409','#e8f4ff','#ffffff')" title="Contraste"></div>
    <div class="clr" style="background:#0d1520" onclick="setBg(this,'#0d1520','#c8dff0','#e8f8ff')" title="Slate"></div>
  </div>
</div>

<!-- SLASH AUTOCOMPLETE MENU (hidden by default) -->
<div class="slash-menu" id="slashMenu" style="display:none"></div>

<!-- EDITOR -->
<div class="editor-wrap" id="editorWrap">
  <div id="doc" contenteditable="true" spellcheck="true" lang="es">
    {contenido_inicial}
  </div>
</div>

<!-- ACTION BAR -->
<div class="abar">
  <button class="ab" onclick="sendMsg('optimize')">✦ Optimizar</button>
  <button class="ab" onclick="sendMsg('qa')">◈ Auditar QA</button>
  <button class="ab" onclick="sendMsg('differential')">⊕ Dif. Diagnóstico</button>
  <button class="ab" onclick="sendMsg('definiciones')">◎ Definiciones</button>
  <button class="ab" onclick="copyDoc()">⎘ Copiar</button>
  <button class="ab prime" onclick="sendMsg('export')">↓ Exportar .docx</button>
  <div class="score-wrap">
    <div>
      <div class="score-val" id="scoreVal">0%</div>
      <div class="score-lbl">Calidad</div>
      <div class="score-track"><div class="score-bar-f" id="scoreFill" style="width:0%;background:#1a4a80"></div></div>
    </div>
  </div>
</div>

<script>
var doc = document.getElementById('doc');
var slashMenu = document.getElementById('slashMenu');
var slashActive = false;
var slashStart = 0;
var selectedSlashItem = 0;

var SLASH_CMDS = [
    {{key:'/stoller',label:'Stoller',desc:'Clasificación meniscal RM'}},
    {{key:'/icrs',label:'ICRS',desc:'Clasificación cartílago'}},
    {{key:'/kl',label:'Kellgren-Lawrence',desc:'Artrosis radiográfica'}},
    {{key:'/lca',label:'Hope & Feagin',desc:'Lesión LCA'}},
    {{key:'/pfirrmann',label:'Pfirrmann',desc:'Degeneración discal'}},
    {{key:'/modic',label:'Modic',desc:'Cambios placa terminal'}},
    {{key:'/tirads',label:'ACR TIRADS',desc:'Nódulo tiroideo'}},
    {{key:'/birads',label:'BI-RADS',desc:'Hallazgo mamario'}},
    {{key:'/lirads',label:'LI-RADS',desc:'Lesión hepática'}},
    {{key:'/fleischner',label:'Fleischner',desc:'Nódulo pulmonar'}},
    {{key:'/tecnica',label:'Técnica',desc:'Insertar plantilla técnica'}},
    {{key:'/impresion',label:'Impresión DX',desc:'Plantilla impresión diagnóstica'}},
];

function fmt(cmd) {{
    doc.focus();
    document.execCommand(cmd, false, null);
    updateBtns();
}}

function updateBtns() {{
    ['bold','italic','underline'].forEach(function(c) {{
        var b = document.getElementById('btn'+c[0].toUpperCase()+c.slice(1));
        if(b) b.classList.toggle('on', document.queryCommandState(c));
    }});
}}

function applyFont(f) {{ doc.style.fontFamily = f; }}
function applySize(s) {{ doc.style.fontSize = s+'px'; }}

function setBg(el, bg, col, titleCol) {{
    doc.style.background = bg;
    doc.style.color = col;
    document.querySelectorAll('.clr').forEach(function(d) {{ d.classList.remove('on'); }});
    el.classList.add('on');
    document.querySelectorAll('#doc .section-title').forEach(function(t) {{
        t.style.color = titleCol || col;
    }});
}}

function insertHR() {{
    doc.focus();
    document.execCommand('insertHTML', false, '<hr>');
}}

function insertSecTitle() {{
    doc.focus();
    document.execCommand('insertHTML', false, '<div class="section-title">NUEVA SECCIÓN</div><br>');
}}

// ── SLASH MENU ──
function showSlashMenu(filter) {{
    var sel = window.getSelection();
    if (!sel.rangeCount) return;
    var range = sel.getRangeAt(0);
    var rect = range.getBoundingClientRect();
    var items = filter
        ? SLASH_CMDS.filter(function(c) {{ return c.key.includes(filter.toLowerCase()) || c.label.toLowerCase().includes(filter.toLowerCase()); }})
        : SLASH_CMDS;
    if (!items.length) {{ hideSlashMenu(); return; }}
    slashMenu.innerHTML = items.map(function(c, i) {{
        return '<div class="slash-item'+(i===selectedSlashItem?' active':'')+'" onclick="insertSlash(\''+c.key+'\')"><span class="slash-key">'+c.key+'</span><span class="slash-desc">'+c.label+' — '+c.desc+'</span></div>';
    }}).join('');
    var edRect = document.getElementById('editorWrap').getBoundingClientRect();
    slashMenu.style.display = 'block';
    slashMenu.style.left = Math.max(0, rect.left - edRect.left) + 'px';
    slashMenu.style.bottom = (window.innerHeight - rect.top + 4) + 'px';
    slashMenu.style.top = 'auto';
    slashActive = true;
    selectedSlashItem = 0;
}}

function hideSlashMenu() {{
    slashMenu.style.display = 'none';
    slashActive = false;
    selectedSlashItem = 0;
}}

function insertSlash(cmd) {{
    hideSlashMenu();
    doc.focus();
    // Borrar el texto del slash command
    var sel = window.getSelection();
    if (sel.rangeCount) {{
        var range = sel.getRangeAt(0);
        range.setStart(range.startContainer, Math.max(0, range.startOffset - 20));
        var text = range.toString();
        var slashIdx = text.lastIndexOf('/');
        if (slashIdx >= 0) {{
            range.setStart(range.endContainer, range.endOffset - (text.length - slashIdx));
            range.deleteContents();
        }}
    }}
    var templates = {{
        '/stoller': '<div class="section-title">MENISCOS</div><p>Menisco medial/lateral: alteración de señal grado [I/II/III] de Stoller en [cuerpo/cuerno anterior/posterior], compatible con [cambio degenerativo / desgarro]. [Extrusión de ___ mm].</p>',
        '/icrs': '<p>Cartílago articular: adelgazamiento condral focal grado [I/II/III/IV] de ICRS en [localización], extensión aproximada de ___ mm. [Edema subcondral reactivo].</p>',
        '/kl': '<p>Hallazgos compatibles con osteoartrosis grado [I/II/III/IV] de Kellgren-Lawrence, con [osteofitos / pinzamiento articular / esclerosis subcondral] en [compartimento].</p>',
        '/lca': '<p>Ligamento cruzado anterior con [señal heterogénea / discontinuidad de fibras], compatible con lesión [parcial / completa / crónica] grado [I/II/III] de Hope & Feagin.</p>',
        '/pfirrmann': '<p>Disco [L_-L_]: degeneración grado [I/II/III/IV/V] de Pfirrmann, con [señal reducida / pérdida de distinción pulpo-anular / colapso del espacio discal].</p>',
        '/modic': '<p>Cambios de señal en placa terminal de tipo Modic [I/II/III] en [nivel], indicativos de [edema/inflamación / sustitución grasa / esclerosis].</p>',
        '/tirads': '<p>Nódulo tiroideo [lóbulo D/I]: [descripción]. Clasificación ACR TIRADS [2/3/4/5]. [Recomendación de seguimiento / BAAF según criterios de tamaño].</p>',
        '/birads': '<p>Hallazgo mamario: [descripción]. Categoría ACR BI-RADS [2/3/4A/4B/4C/5]. [Recomendación de manejo].</p>',
        '/lirads': '<p>Lesión focal hepática: [descripción]. Clasificación LI-RADS [LR-1/2/3/4/5/M]. [Correlacionar con contexto clínico y marcadores tumorales].</p>',
        '/fleischner': '<p>Nódulo pulmonar [sólido/subporcentaje] de ___ mm en [lóbulo/segmento]. Por criterios Fleischner 2017: [recomendación de seguimiento].</p>',
        '/tecnica': '<div class="section-title">TÉCNICA</div><p>Estudio de [modalidad] de [región anatómica] en equipo de [__] Tesla. Secuencias [describir] en planos [axial/coronal/sagital]. [Sin / Con] administración de medio de contraste [paramagnético/yodado] a dosis estándar.</p>',
        '/impresion': '<div class="section-title">IMPRESIÓN DIAGNÓSTICA</div><li>Hallazgo principal con clasificación específica y grado.</li><li>Hallazgo secundario.</li><li>Recomendación de manejo o seguimiento.</li>',
    }};
    var html = templates[cmd] || '<p>['+cmd+']</p>';
    document.execCommand('insertHTML', false, html);
}}

// ── SCORE ──
function calcScore() {{
    var t = doc.innerText.toUpperCase();
    var secs = ['TÉCNICA','HALLAZGOS','IMPRESIÓN'];
    var found = secs.filter(function(s) {{ return t.includes(s); }}).length;
    var words = t.trim().split(/ +/).filter(Boolean).length;
    return Math.min(100, Math.round((found/3)*40 + Math.min(words/200,1)*35 + (found===3?25:0)));
}}

function updateScore() {{
    var s = calcScore();
    document.getElementById('scoreVal').textContent = s+'%';
    var fill = document.getElementById('scoreFill');
    fill.style.width = s+'%';
    fill.style.background = s>=80 ? '#00cc88' : s>=50 ? '#f0a020' : '#1a4a80';
}}

doc.addEventListener('input', updateScore);
doc.addEventListener('keyup', updateBtns);
doc.addEventListener('mouseup', updateBtns);

// ── SLASH DETECTION ──
doc.addEventListener('keydown', function(e) {{
    if (slashActive) {{
        if (e.key === 'Escape') {{ hideSlashMenu(); return; }}
        if (e.key === 'Enter') {{
            e.preventDefault();
            var items = slashMenu.querySelectorAll('.slash-item');
            if (items[selectedSlashItem]) items[selectedSlashItem].click();
            return;
        }}
        if (e.key === 'ArrowDown') {{
            e.preventDefault();
            selectedSlashItem = (selectedSlashItem + 1) % slashMenu.querySelectorAll('.slash-item').length;
            slashMenu.querySelectorAll('.slash-item').forEach(function(el, i) {{
                el.classList.toggle('active', i === selectedSlashItem);
            }});
            return;
        }}
        if (e.key === 'ArrowUp') {{
            e.preventDefault();
            var n = slashMenu.querySelectorAll('.slash-item').length;
            selectedSlashItem = (selectedSlashItem - 1 + n) % n;
            slashMenu.querySelectorAll('.slash-item').forEach(function(el, i) {{
                el.classList.toggle('active', i === selectedSlashItem);
            }});
            return;
        }}
    }}
}});

doc.addEventListener('keyup', function(e) {{
    var sel = window.getSelection();
    if (!sel.rangeCount) return;
    var range = sel.getRangeAt(0);
    var text = range.startContainer.textContent || '';
    var offset = range.startOffset;
    var textBefore = text.slice(0, offset);
    var slashMatch = textBefore.match(new RegExp('/(\\w*)$'));
    if (slashMatch) {{
        showSlashMenu(slashMatch[1]);
    }} else {{
        hideSlashMenu();
    }}
}});

doc.addEventListener('click', function() {{ hideSlashMenu(); }});

function copyDoc() {{
    var r = document.createRange();
    r.selectNode(doc);
    window.getSelection().removeAllRanges();
    window.getSelection().addRange(r);
    document.execCommand('copy');
    window.getSelection().removeAllRanges();
}}

function sendMsg(type) {{
    window.parent.postMessage({{type: type, html: doc.innerHTML, text: doc.innerText}}, '*');
}}

window.addEventListener('load', function() {{ updateScore(); }});
</script>
</body>
</html>"""

    components.html(editor_html, height=760, scrolling=False)

    # Acciones externas del editor
    st.markdown("<div style='padding:4px 0'>", unsafe_allow_html=True)
    ca, cb, cc, cd = st.columns(4)

    with ca:
        if st.button("✦ Optimizar conclusión", use_container_width=True):
            client, model_name = get_client()
            if client and st.session_state.reporte_texto:
                sistema = construir_sistema_prompt(modalidad, region, instrucciones, plantilla_txt,
                                                   st.session_state.clasificaciones_activas,
                                                   st.session_state.estilo_usuario)
                with st.spinner("Refinando..."):
                    try:
                        res = client.chat.completions.create(
                            model=model_name,
                            messages=[
                                {"role": "system", "content": sistema},
                                {"role": "user", "content": f"MEJORA SOLO el bloque IMPRESIÓN DIAGNÓSTICA del siguiente informe. Hazla más elegante, concluyente, con grados exactos de clasificación y recomendación de manejo. Conserva Técnica y Hallazgos INTACTOS. Devuelve el informe COMPLETO.\n\n{st.session_state.reporte_texto}"}
                            ],
                            temperature=0.2
                        )
                        texto = res.choices[0].message.content
                        st.session_state.reporte_texto = texto
                        st.session_state.reporte_html = texto_a_html(texto)
                        st.session_state.qa_resultado = calcular_qa(texto, region)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    with cb:
        if st.button("⊕ Diagnóstico diferencial", use_container_width=True):
            client, model_name = get_client()
            if client and st.session_state.reporte_texto:
                with st.spinner("Generando DD..."):
                    try:
                        res = client.chat.completions.create(
                            model=model_name,
                            messages=[{
                                "role": "user",
                                "content": f"""Como radiólogo experto, analiza este informe y genera:

1. DIAGNÓSTICO DIFERENCIAL: Lista los 3-5 diagnósticos alternativos más probables con argumentos a favor y en contra de cada uno, basándote en los hallazgos de imagen.
2. CORRELACIÓN CLÍNICA: Hallazgos de imagen que apoyan o descartan cada diagnóstico.
3. PASOS DIAGNÓSTICOS ADICIONALES: Estudios complementarios que agregarían valor.

Sé conciso y clínicamente útil. Sin asteriscos. Títulos en MAYÚSCULAS.

INFORME:
{st.session_state.reporte_texto}"""
                            }],
                            temperature=0.3
                        )
                        st.session_state.copilot_panel = res.choices[0].message.content
                        st.session_state.copilot_tipo = "differential"
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    with cc:
        if st.button("◎ Definiciones", use_container_width=True):
            client, model_name = get_client()
            if client and st.session_state.reporte_texto:
                with st.spinner("Analizando..."):
                    try:
                        res = client.chat.completions.create(
                            model=model_name,
                            messages=[{
                                "role": "user",
                                "content": f"""Analiza este informe radiológico y proporciona:

1. CLASIFICACIONES IDENTIFICADAS: Cada clasificación mencionada, su grado y significado clínico exacto.
2. CLASIFICACIONES FALTANTES: Que deberían incluirse según los hallazgos descritos.
3. DEFINICIONES OPERATIVAS: De los términos técnicos más relevantes (1-2 líneas cada uno).
4. BIBLIOGRAFÍA: Referencias de las clasificaciones usadas.

Sin asteriscos. Conciso y didáctico.

INFORME:
{st.session_state.reporte_texto}"""
                            }],
                            temperature=0.15
                        )
                        st.session_state.copilot_panel = res.choices[0].message.content
                        st.session_state.copilot_tipo = "definiciones"
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

    with cd:
        if st.session_state.reporte_texto:
            docx_bytes = generar_docx(
                st.session_state.reporte_html,
                st.session_state.reporte_texto
            )
            st.download_button(
                "↓ Exportar .docx",
                data=docx_bytes,
                file_name=f"AURA_{region.replace(' ','_').replace('/','_')}_{modalidad[:2]}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# PANEL DERECHO — Copiloto IA contextual
# ══════════════════════════════════════════════════════════════
with col_der:
    st.markdown("<div style='padding:12px 4px 0 4px;height:calc(100vh - 48px);overflow-y:auto'>", unsafe_allow_html=True)

    # ── QA del informe ──
    qa = st.session_state.qa_resultado
    score = qa.get("score", 0)

    score_color = "#00cc88" if score >= 80 else "#f0a020" if score >= 50 else "#ff4060"
    score_bg = "rgba(0,200,100,0.06)" if score >= 80 else "rgba(240,160,30,0.06)" if score >= 50 else "rgba(255,60,80,0.06)"
    score_border = "rgba(0,200,100,0.15)" if score >= 80 else "rgba(240,160,30,0.15)" if score >= 50 else "rgba(255,60,80,0.15)"

    st.markdown(f"""
    <div style="background:{score_bg};border:1px solid {score_border};border-radius:8px;padding:14px 16px;margin-bottom:10px">
        <div style="display:flex;align-items:center;gap:12px">
            <div>
                <div style="font-family:'Inter',sans-serif;font-weight:700;font-size:30px;color:{score_color};line-height:1">{score}</div>
                <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:#3a5070;text-transform:uppercase;letter-spacing:0.1em;margin-top:2px">QA SCORE</div>
                <div style="width:60px;height:2px;background:#0f1428;border-radius:1px;margin-top:6px;overflow:hidden">
                    <div style="width:{score}%;height:100%;background:{score_color};border-radius:1px"></div>
                </div>
            </div>
            <div style="flex:1">
                {''.join([
                    f'<div style="display:flex;align-items:center;gap:5px;padding:2px 0"><span style="font-size:11px;color:{"#00cc88" if v else "#ff4060"};font-family:JetBrains Mono,monospace">{"✓" if v else "✗"}</span><span style="font-family:JetBrains Mono,monospace;font-size:10px;color:{"#4a8a6a" if v else "#8a4a55"}">{k}</span></div>'
                    for k, v in qa.get("secciones", {"TÉCNICA": False, "HALLAZGOS": False, "IMPRESIÓN": False}).items()
                ])}
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Omisiones anatómicas ──
    omisiones = qa.get("omisiones", [])
    if omisiones:
        st.markdown(f"""
        <div style="background:rgba(240,160,30,0.05);border:1px solid rgba(240,160,30,0.12);border-radius:8px;padding:10px 12px;margin-bottom:10px">
            <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:#8a6020;letter-spacing:0.15em;text-transform:uppercase;margin-bottom:8px">⚠ Omisiones anatómicas</div>
            {''.join([f'<div style="font-family:JetBrains Mono,monospace;font-size:10px;color:#c08030;padding:2px 0">· {o}</div>' for o in omisiones])}
        </div>
        """, unsafe_allow_html=True)

    # ── Terminología ──
    terminologia = qa.get("terminologia", [])
    if terminologia:
        st.markdown(f"""
        <div style="background:rgba(255,60,80,0.04);border:1px solid rgba(255,60,80,0.1);border-radius:8px;padding:10px 12px;margin-bottom:10px">
            <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:#803040;letter-spacing:0.15em;text-transform:uppercase;margin-bottom:8px">✗ Terminología</div>
            {''.join([f'<div style="font-family:JetBrains Mono,monospace;font-size:10px;color:#cc5060;padding:3px 0">"{t["incorrecto"]}" → <span style="color:#7aaa9a">{t["correcto"][:40]}...</span></div>' for t in terminologia[:4]])}
        </div>
        """, unsafe_allow_html=True)

    # ── Clasificaciones disponibles para la región ──
    clasif_region = KB["anatomia_regiones"].get(region, {}).get("clasificaciones_relevantes", [])
    if clasif_region:
        st.markdown("""
        <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:#3a5070;letter-spacing:0.15em;text-transform:uppercase;margin-bottom:6px;margin-top:4px">
            Clasificaciones · {region}
        </div>
        """.format(region=region), unsafe_allow_html=True)

        for nombre_c in clasif_region:
            datos_c = KB["clasificaciones"].get(nombre_c, {})
            with st.expander(nombre_c, expanded=False):
                st.markdown(f'<div style="font-family:JetBrains Mono,monospace;font-size:10px;color:#5a8aaf;margin-bottom:6px">{datos_c.get("descripcion","")}</div>', unsafe_allow_html=True)
                for grado, desc in datos_c.get("grados", {}).items():
                    is_active = st.session_state.clasificaciones_activas.get(nombre_c) == grado
                    bg = "rgba(0,100,180,0.15)" if is_active else "transparent"
                    st.markdown(f'<div style="font-family:JetBrains Mono,monospace;font-size:10px;color:#4a7aaf;padding:4px 6px;border-radius:4px;margin:1px 0;background:{bg}"><b style="color:#7aaccc">Gr.{grado}</b> · <span style="color:#5a8aaf">{desc}</span></div>', unsafe_allow_html=True)
                ref = datos_c.get("referencia", "")
                if ref:
                    st.markdown(f'<div style="font-family:JetBrains Mono,monospace;font-size:9px;color:#2a4060;margin-top:6px;padding-top:4px;border-top:1px solid #1a2540">{ref}</div>', unsafe_allow_html=True)

    # ── Auditoría QA botón ──
    if st.button("◈ Auditar informe completo", use_container_width=True):
        client, model_name = get_client()
        if client and st.session_state.reporte_texto:
            with st.spinner("Auditando..."):
                try:
                    res = client.chat.completions.create(
                        model=model_name,
                        messages=[{
                            "role": "user",
                            "content": f"""Eres un radiólogo experto realizando control de calidad (QA) de este informe.

Evalúa y reporta:
1. ESTRUCTURA: ¿Tiene Técnica, Hallazgos e Impresión? ¿Fluye lógicamente?
2. COMPLETITUD ANATÓMICA: ¿Qué estructuras de {region} no fueron mencionadas?
3. CLASIFICACIONES: ¿Están bien usadas y son las apropiadas? ¿Falta alguna?
4. TERMINOLOGÍA: ¿Hay términos vagos o incorrectos en español radiológico?
5. IMPRESIÓN DIAGNÓSTICA: ¿Es concluyente y clínicamente accionable?
6. CALIFICACIÓN: 0-100 con justificación.

Sé crítico y específico. Sin asteriscos.

INFORME:
{st.session_state.reporte_texto}"""
                        }],
                        temperature=0.15
                    )
                    st.session_state.copilot_panel = res.choices[0].message.content
                    st.session_state.copilot_tipo = "qa_full"
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")

    # ── Resultado del copiloto ──
    if st.session_state.copilot_panel:
        tipo_labels = {
            "differential": "DIAGNÓSTICO DIFERENCIAL",
            "definiciones": "DEFINICIONES & CLASIFICACIONES",
            "qa_full": "AUDITORÍA QA COMPLETA",
        }
        tipo_label = tipo_labels.get(st.session_state.copilot_tipo, "COPILOTO AURA")
        tipo_color = {
            "differential": "#4a90cc",
            "definiciones": "#6a9a4a",
            "qa_full": "#c08030",
        }.get(st.session_state.copilot_tipo, "#4a90cc")

        st.markdown(f"""
        <div style="background:#0c1020;border:1px solid #1a2540;border-radius:8px;padding:14px;margin-top:10px">
            <div style="font-family:'JetBrains Mono',monospace;font-size:9px;color:{tipo_color};letter-spacing:0.15em;text-transform:uppercase;margin-bottom:10px;display:flex;align-items:center;gap:6px">
                <div style="width:4px;height:4px;border-radius:50%;background:{tipo_color}"></div>
                {tipo_label}
            </div>
            <div style="font-family:'JetBrains Mono',monospace;font-size:10px;color:#7a9abf;line-height:1.75;white-space:pre-wrap">{st.session_state.copilot_panel}</div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("✕ Cerrar", use_container_width=True):
            st.session_state.copilot_panel = ""
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
