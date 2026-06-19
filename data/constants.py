SUGERENCIAS = {
    "Resonancia Magnética": [
        "Stoller III → desgarro confirmado",
        "Extrusión meniscal >3 mm",
        "ICRS III → >50% grosor condral",
        "Kellgren-Lawrence III gonartrosis",
        "Edema óseo subcondral activo",
        "Lesión LCA Hope & Feagin parcial",
        "Pfirrmann IV discopatía",
        "Contusión ósea por impacto",
    ],
    "Tomografía Computarizada": [
        "Hounsfield: hueso ~700 UH",
        "Adenopatía >1 cm eje corto",
        "Nódulo Fleischner >6 mm sólido",
        "WELLS alta probabilidad TEP",
        "Murphy score apendicitis",
        "ASPECTS ACV isquémico",
    ],
    "Radiografía": [
        "Kellgren-Lawrence I-IV artrosis",
        "Cobb >10° escoliosis",
        "Índice cardiotorácico >0.5",
        "Radiopacidad lobar consolidación",
        "Línea pleural neumotórax",
    ],
    "Ultrasonido": [
        "TIRADS 4 → considerar BAAF",
        "BI-RADS 4B → biopsia indicada",
        "Resistividad >0.7 sospecha maligna",
        "Murphy positivo colecistitis",
        "McBurney dolor apendicitis",
    ],
    "PET-CT": [
        "SUVmax >2.5 actividad metabólica",
        "LI-RADS 5 → HCC definitivo",
        "Captación focal vs difusa",
        "Respuesta PERCIST criterios",
    ],
}

CLASIFICACIONES = {
    "Menisco · Stoller": [
        ("I", "Señal focal intrameniscal"),
        ("II", "Señal lineal, no articular"),
        ("III", "Alcanza superficie → desgarro"),
    ],
    "Cartílago · ICRS": [
        ("I", "Fibrilación superficial"),
        ("II", "<50% grosor"),
        ("III", ">50% grosor"),
        ("IV", "Hueso subcondral expuesto"),
    ],
    "Artrosis · Kellgren-Lawrence": [
        ("I", "Posible osteofito"),
        ("II", "Osteofito definido"),
        ("III", "Pinzamiento moderado"),
        ("IV", "Pinzamiento grave"),
    ],
    "LCA · Hope & Feagin": [
        ("Parcial", "Fibras continuas, señal ↑"),
        ("Completa", "Discontinuidad total"),
        ("Crónica", "Fibras atróficas"),
    ],
    "Columna · Pfirrmann": [
        ("I", "Núcleo brillante homogéneo"),
        ("II", "Señal alta, zona no clara"),
        ("III", "Señal gris, distinción borrosa"),
        ("IV", "Señal baja, sin distinción"),
        ("V", "Sin espacio discal"),
    ],
    "TIRADS · ACR": [
        ("2", "No sospechoso"),
        ("3", "Levemente sospechoso"),
        ("4", "Moderadamente sospechoso"),
        ("5", "Altamente sospechoso"),
    ],
    "BI-RADS · ACR": [
        ("2", "Benigno"),
        ("3", "Probablemente benigno"),
        ("4A/4B/4C", "Sospechoso — biopsia"),
        ("5", "Altamente maligno"),
    ],
}

REGIONES = [
    "Rodilla", "Columna lumbar", "Columna cervical", "Hombro",
    "Cadera", "Tobillo / Pie", "Muñeca / Mano", "Codo",
    "Cerebro", "Columna dorsal", "Tórax", "Abdomen / Pelvis",
    "Mama", "Tiroides", "Hígado",
]
