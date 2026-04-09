import streamlit as st
from comparador_base import render_comparacao_fundos
from comparador_base import render_pesquisa_ativo_por_data
from ui_padrao import aplicar_ui_padrao, render_titulo_padrao

st.set_page_config(page_title="Fundos D+30", layout="wide")

aplicar_ui_padrao()
render_titulo_padrao("Fundos D+30")

tab1, tab2 = st.tabs([
    "Comparação entre fundos",
    "Comparação entre ativos"
])

with tab1:
        render_comparacao_fundos(
        pasta=r"Z:\Asset Management\Equipe\Lívia\Rascunhos\Site_fundos",
        arquivo_xlsx="carteira_fundos_consolidada.xlsx",
        titulo="Fundos D+30",
        abas_permitidas=["GERAES 30", "Daycoval Classic 30", "Riza Lotus Plus", "ARX Vinson Advisory", "Sparta Max Advisory", "Polo Crédito Corporativo", "Iridium Titam Advisory", "Porto Seguro Ipê", "Sparta Top Advisory", "Absolute Creta"]  # opcional
    )
        
with tab2:
    render_pesquisa_ativo_por_data(
        pasta=r"Z:\Asset Management\Equipe\Lívia\Rascunhos\Site_fundos",
        arquivo_xlsx="carteira_fundos_consolidada.xlsx",
        abas_permitidas=["GERAES 30", "Daycoval Classic 30", "Riza Lotus Plus", "ARX Vinson Advisory", "Sparta Max Advisory", "Polo Crédito Corporativo", "Iridium Titam Advisory", "Porto Seguro Ipê", "Sparta Top Advisory", "Absolute Creta"]  # <-- somente fundos D+30
    )