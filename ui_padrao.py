import streamlit as st

CSS_UI_PADRAO = """
<style>
.block-container {
    padding-top: 3.4rem !important;
    padding-bottom: 2rem;
}

.page-title {
    display: block;
    font-family: Calibri, 'Segoe UI', Arial, sans-serif;
    font-size: 40px;
    font-weight: 700;
    color: black;
    line-height: 1.2;
    margin: 0 0 12px 0;
    padding: 0;
    overflow: visible;
}

/* FUNDO DO MENU LATERAL INTEIRO */
section[data-testid="stSidebar"] {
    background-color: #203764 !important;
}

section[data-testid="stSidebar"] > div {
    background-color: #203764 !important;
}

section[data-testid="stSidebar"] * {
    color: white !important;
}

/* TÍTULOS DE SEÇÃO */
.sectionbar {
    color: #ffffff;
    padding: 8px 12px;
    border-radius: 0px;
    font-weight: 700;
    font-size: 18px;
    margin-top: 12px;
    margin-bottom: 10px;
}

.sectionbar-fundos { background: #0F2D52; }
.sectionbar-ativos { background: #1D5C9E; }
.sectionbar-similaridade { background: #277DD3; }
.sectionbar-exportar { background: #ADCFFD; }
.sectionbar-comparacao { background: #0F2D52; }
.sectionbar-carteira { background: #1D5C9E; }
.sectionbar-movimentacao { background: #277DD3; }
.sectionbar-fundoindividual { background: #0F2D52; }


/* TABS */
button[data-baseweb="tab"] {
    color: #808080 !important;
}

button[data-baseweb="tab"][aria-selected="true"] {
    color: #000000 !important;
}

div[data-baseweb="tab-highlight"] {
    background-color: #203764 !important;
}

/* KPI */
.kpi-box {
    border: 1px solid #d0d0d0;
    border-radius: 0px;
    padding: 10px 12px;
    background: #EDEDED;
    color: #ffffff
    font-weight: 700;
}

/* DATAFRAME */
div[data-testid="stDataFrame"] {
    border-radius: 0px !important;
}

div[data-testid="stDataFrame"] > div {
    border: 1px solid #e0e0e0;
    border-radius: 0px !important;
}

div[data-testid="stDataFrame"] [data-testid="stSkeleton"],
div[data-testid="stDataFrame"] div[role="grid"],
div[data-testid="stDataFrame"] canvas {
    border-radius: 0px !important;
}

/* INPUTS / SELECTS */
div[data-testid="stTextInput"] *,
div[data-testid="stSelectbox"] *,
div[data-testid="stMultiSelect"] *,
div[data-testid="stDateInput"] *,
div[data-testid="stNumberInput"] *,
div[data-baseweb="input"] *,
div[data-baseweb="select"] * {
    border-radius: 0px !important;
}

/* METRICS */
div[data-testid="stMetricLabel"] p {
    font-size: 14px !important;
    font-weight: 500 !important;
}

div[data-testid="stMetricValue"] {
    font-size: 20px !important;
    font-weight: 600 !important;
}
</style>
"""

def aplicar_ui_padrao():
    st.markdown(CSS_UI_PADRAO, unsafe_allow_html=True)

def render_titulo_padrao(texto: str):
    st.markdown(f"<div class='page-title'>{texto}</div>", unsafe_allow_html=True)

def render_sectionbar(texto: str, classe_extra: str = ""):
    classes = f"sectionbar {classe_extra}".strip()
    st.markdown(f"<div class='{classes}'>{texto}</div>", unsafe_allow_html=True)