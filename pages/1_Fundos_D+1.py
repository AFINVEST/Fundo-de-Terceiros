import streamlit as st
from comparador_base import render_comparacao_fundos
from comparador_base import render_pesquisa_ativo_por_data
from ui_padrao import aplicar_ui_padrao, render_titulo_padrao

st.set_page_config(page_title="Fundos D+1", layout="wide")

aplicar_ui_padrao()
render_titulo_padrao("Fundos D+1")

BASE_DIR = Path(__file__).resolve().parent.parent

tab1, tab2 = st.tabs([
    "Comparação entre fundos",
    "Comparação entre ativos"
])

with tab1:
        render_comparacao_fundos(
        pasta=str(BASE_DIR),
        arquivo_xlsx="carteira_fundos_consolidada.xlsx",
        titulo="Fundos D+1",
        abas_permitidas=["GERAES", "ARX Denali", "Capitânia Top", "Riza Lotus", "Western Asset Total", "Daycoval Classic", "Iridium Apollo", "Porto Seguro FIRF", "Nu Reserva Imediata", "Inter Conservador", "Absolute Atenas"]  # opcional
    )

with tab2:
        render_pesquisa_ativo_por_data(
        pasta=str(BASE_DIR),
        arquivo_xlsx="carteira_fundos_consolidada.xlsx",
        abas_permitidas=["GERAES", "ARX Denali", "Capitânia Top", "Riza Lotus", "Western Asset Total", "Daycoval Classic", "Iridium Apollo", "Porto Seguro FIRF", "Nu Reserva Imediata", "Inter Conservador", "Absolute Atenas"]  # mesma lógica da página
    )
