from pathlib import Path
import streamlit as st

from comparador_base import render_analise_fundo_individual
from comparador_base import render_pesquisa_ativo_por_data
from ui_padrao import aplicar_ui_padrao, render_titulo_padrao

st.set_page_config(page_title="Análise de Fundo", layout="wide")

aplicar_ui_padrao()
render_titulo_padrao("Análise de Fundo")

BASE_DIR = Path(__file__).resolve().parent.parent

tab1, tab2 = st.tabs([
    "Análise individual",
    "Comparação entre ativos"
])

with tab1:
        render_analise_fundo_individual(
        pasta=str(BASE_DIR),
        arquivo_xlsx="carteira_fundos_consolidada.xlsx",
        titulo="Análise de Carteira (Fundo Individual)",
        abas_permitidas=["GERAES", "ARX Denali", "Capitânia Top", "Riza Lotus", "Western Asset Total", "Daycoval Classic", "Iridium Apollo", "Porto Seguro FIRF", "Nu Reserva Imediata", "Inter Conservador", "Absolute Atenas", "GERAES 30", "Daycoval Classic 30", "Riza Lotus Plus", "ARX Vinson Advisory", "Sparta Max Advisory", "Polo Crédito Corporativo", "Iridium Titan Advisory", "Porto Seguro Ipê", "Sparta Top Advisory", "Absolute Creta", "Horizonte", "JGP Select", "ARX Everest Advisory", "Polo Total", "Absolute Olimpia"]  # opcional
    )

with tab2:
    render_pesquisa_ativo_por_data(
        pasta=str(BASE_DIR),
        arquivo_xlsx="carteira_fundos_consolidada.xlsx",
        abas_permitidas=["GERAES", "ARX Denali", "Capitânia Top", "Riza Lotus", "Western Asset Total", "Daycoval Classic", "Iridium Apollo", "Porto Seguro FIRF", "Nu Reserva Imediata", "Inter Conservador", "Absolute Atenas", "GERAES 30", "Daycoval Classic 30", "Riza Lotus Plus", "ARX Vinson Advisory", "Sparta Max Advisory", "Polo Crédito Corporativo", "Iridium Titan Advisory", "Porto Seguro Ipê", "Sparta Top Advisory", "Absolute Creta", "Horizonte", "JGP Select", "ARX Everest Advisory", "Polo Total", "Absolute Olimpia"]  # <-- somente fundos D+30
    )

