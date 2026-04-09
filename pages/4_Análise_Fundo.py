from pathlib import Path
import streamlit as st

from comparador_base import render_analise_fundo_individual
from ui_padrao import aplicar_ui_padrao, render_titulo_padrao
from pathlib import Path

PASTA = Path(r"Z:\Asset Management\Equipe\Lívia\Rascunhos\Site_fundos")

st.set_page_config(page_title="Análise de Fundo", layout="wide")

BASE_DIR = Path(__file__).resolve().parent.parent

aplicar_ui_padrao()
render_titulo_padrao("Análise de Fundo")

render_analise_fundo_individual(
    pasta=str(BASE_DIR),
    arquivo_xlsx="carteira_fundos_consolidada.xlsx",
    titulo="Análise de Carteira (Fundo Individual)",
    abas_permitidas=["GERAES", "ARX Denali", "Capitânia Top", "Riza Lotus", "Western Asset Total", "Daycoval Classic", "Iridium Apollo", "Porto Seguro FIRF", "Nu Reserva Imediata", "Inter Conservador", "Absolute Atenas", "GERAES 30", "Daycoval Classic 30", "Riza Lotus Plus", "ARX Vinson Advisory", "Sparta Max Advisory", "Polo Crédito Corporativo", "Iridium Titam Advisory", "Porto Seguro Ipê", "Sparta Top Advisory", "Absolute Creta", "Horizonte", "JGP Select", "ARX Everest Advisory", "Polo Total", "Absolute Olimpia"]  # opcional
)
