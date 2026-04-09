import streamlit as st
from comparador_base import render_comparacao_fundos
from comparador_base import render_pesquisa_ativo_por_data
from ui_padrao import aplicar_ui_padrao, render_titulo_padrao
from pathlib import Path

st.set_page_config(page_title="Fundos D+60", layout="wide")

aplicar_ui_padrao()
render_titulo_padrao("Fundos D+60")

BASE_DIR = Path(__file__).resolve().parent.parent

tab1, tab2 = st.tabs([
    "Comparação entre fundos",
    "Comparação entre ativos"
])

with tab1:
        render_comparacao_fundos(
        pasta=r"Z:\Asset Management\Equipe\Lívia\Rascunhos\Site_fundos",
        arquivo_xlsx="carteira_fundos_consolidada.xlsx",
        titulo="Fundos D+60",
        abas_permitidas=["HORIZONTE", "JGP Select", "ARX Everest Advisory", "Polo Total", "Absolute Olimpia"]  # opcional
    )
     
with tab2:
    render_pesquisa_ativo_por_data(
        pasta=r"Z:\Asset Management\Equipe\Lívia\Rascunhos\Site_fundos",
        arquivo_xlsx="carteira_fundos_consolidada.xlsx",
        abas_permitidas=["HORIZONTE", "JGP Select", "ARX Everest Advisory", "Polo Total", "Absolute Olimpia"]  # <-- somente fundos D+60
    )
