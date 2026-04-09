from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import numpy as np
import altair as alt
from ui_padrao import aplicar_ui_padrao, render_sectionbar

# ==========================
# DEFAULTS (você pode sobrescrever na página)
# ==========================
COL_CODIGO = "Codigo"
COL_DESCRICAO = "Descrição"
COLUNAS_EXTRAS_SITE = ["ISIN", "Emissor", "Classe", "Vencimento"]
ARQUIVO_XLSX_PADRAO = "carteira_fundos_consolidada.xlsx"

TIPO_OPCOES = [
    "Tipo de investimento",
    "Tipo de Investimento",
    "Tipo_Investimento",
    "Tipo",
    "Classe",
]

# ==========================
# HELPERS
# ==========================
def encontrar_coluna(df: pd.DataFrame, opcoes: List[str]) -> Optional[str]:
    cols_norm = {str(c).strip().lower(): c for c in df.columns}
    for op in opcoes:
        key = op.strip().lower()
        if key in cols_norm:
            return cols_norm[key]
    return None


def to_number(x) -> float:
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    if s in ("", "-", "–") or s.lower() == "nan":
        return 0.0
    s = s.replace("R$", "").replace(" ", "")
    s = re.sub(r"[^0-9,\.\-]", "", s)
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def parse_date_col(col) -> Optional[pd.Timestamp]:
    if isinstance(col, pd.Timestamp):
        return pd.Timestamp(col).normalize()
    dt = pd.to_datetime(str(col).strip(), errors="coerce", dayfirst=True)
    if pd.isna(dt):
        return None
    return pd.Timestamp(dt).normalize()


def get_date_columns(df: pd.DataFrame) -> List[str]:
    parsed: List[Tuple[str, pd.Timestamp]] = []
    for c in df.columns:
        dt = parse_date_col(c)
        if dt is not None:
            parsed.append((c, dt))
    parsed.sort(key=lambda t: t[1])
    return [c for c, _ in parsed]


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if COL_CODIGO not in df.columns:
        df = df.rename(columns={df.columns[0]: COL_CODIGO})
    if COL_DESCRICAO not in df.columns:
        raise KeyError(f"Não encontrei '{COL_DESCRICAO}'. Colunas: {list(df.columns)}")
    if len(df) > 0 and str(df.iloc[0][COL_DESCRICAO]).strip() == "Descrição":
        df = df.iloc[1:].copy()
    return df

ORDEM_COLUNAS_SITE = [
    "Código",
    "ISIN",
    "Descrição",
    "Emissor",
    "Classe",
    "Tipo de Investimento",
    "Vencimento",
    "Valor_$",
    "%PL",
    "Valor Base (R$)",
    "%PL Base",
    "Variação (R$)",
    "Variação (%PL)",
    "Variação (%)",
]

COLUNAS_PADRAO_INICIAIS = [
    "Código",
    "ISIN",
    "Classe",
    "Vencimento",
    "Valor_$",
    "%PL",
]


def get_colunas_disponiveis(df: pd.DataFrame) -> list[str]:
    return [c for c in ORDEM_COLUNAS_SITE if c in df.columns]


def get_colunas_iniciais(df: pd.DataFrame) -> list[str]:
    return [c for c in COLUNAS_PADRAO_INICIAIS if c in df.columns]

def asset_key(row: pd.Series) -> str:
    cod = str(row.get(COL_CODIGO, "")).strip()
    desc = str(row.get(COL_DESCRICAO, "")).strip()
    if cod and cod.lower() != "nan":
        return cod
    return desc


def build_fund_table(df_raw: pd.DataFrame) -> Tuple[pd.DataFrame, List[str], Dict[str, pd.Timestamp]]:
    df = normalize_columns(df_raw)
    date_cols = get_date_columns(df)
    if not date_cols:
        raise ValueError("Não encontrei colunas de datas nessa aba.")

    date_map = {c: parse_date_col(c) for c in date_cols}
    col_tipo = encontrar_coluna(df, TIPO_OPCOES)

    cols_base = [COL_CODIGO]

    if "ISIN" in df.columns:
        cols_base.append("ISIN")

    cols_base.append(COL_DESCRICAO)

    if "Emissor" in df.columns:
        cols_base.append("Emissor")

    if "Classe" in df.columns:
        cols_base.append("Classe")

    if col_tipo and col_tipo not in cols_base:
        cols_base.append(col_tipo)

    if "Vencimento" in df.columns:
        cols_base.append("Vencimento")

    out = df[cols_base + date_cols].copy()
    out["_ATIVO"] = out.apply(asset_key, axis=1).astype(str).str.strip()

    MARCADORES_FIM = {"SOMA_COLUNA", "PL_CDA_FI_PL", "DIFERENCA", "STATUS"}

    out[COL_CODIGO] = out[COL_CODIGO].astype(str).replace("nan", "").replace("None", "").str.strip()
    out[COL_DESCRICAO] = out[COL_DESCRICAO].astype(str).replace("nan", "").replace("None", "").str.strip()

    out = out[
        ~out[COL_CODIGO].isin(MARCADORES_FIM) &
        ~out[COL_DESCRICAO].isin(MARCADORES_FIM)
    ].copy()

    if col_tipo:
        out["_TIPO"] = (
            out[col_tipo]
            .astype(str)
            .replace("nan", "")
            .replace("None", "")
            .replace("NaT", "")
            .str.strip()
        )
    else:
        out["_TIPO"] = ""

    for c in date_cols:
        out[c] = out[c].apply(to_number)

    return out, date_cols, date_map

@st.cache_data(show_spinner=False)
def load_workbook_from_path(path: str) -> Dict[str, pd.DataFrame]:
    return pd.read_excel(path, sheet_name=None, dtype=object)

@st.cache_data(show_spinner=False)
def load_correlacoes(path_xlsx: str) -> pd.DataFrame:
    """
    Lê a matriz de correlação da aba 'Correlação código' (ou variações).
    Retorna um DataFrame quadrado: index = fundos, colunas = fundos.
    """
    # tenta nomes possíveis (pra não quebrar se mudar)
    sheet_candidates = ["Correlação código", "Correlações código", "Correlacao codigo", "Correlacoes codigo"]

    xl = pd.ExcelFile(path_xlsx)
    sheet = None
    for s in sheet_candidates:
        if s in xl.sheet_names:
            sheet = s
            break
    if sheet is None:
        raise ValueError(f"Não encontrei a aba de correlação. Abas disponíveis: {xl.sheet_names}")

    df = pd.read_excel(path_xlsx, sheet_name=sheet, dtype=object)

    # primeira coluna costuma vir como "Unnamed: 0" (rótulo das linhas)
    first_col = df.columns[0]
    df = df.rename(columns={first_col: "_ROW"}).copy()

    # limpa nomes
    df["_ROW"] = df["_ROW"].astype(str).str.strip()
    df.columns = [str(c).strip() for c in df.columns]

    # index = nomes das linhas
    df = df.set_index("_ROW")

    # garante numérico
    for c in df.columns:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    return df


def get_correlacao(df_corr: pd.DataFrame, fundo_a: str, fundo_b: str) -> Optional[float]:
    """
    Busca correlação na matriz. Tenta match exato e, se não achar, tenta case-insensitive.
    """
    if df_corr is None or df_corr.empty:
        return None

    a = str(fundo_a).strip()
    b = str(fundo_b).strip()

    # match exato
    if a in df_corr.index and b in df_corr.columns:
        v = df_corr.loc[a, b]
        return None if pd.isna(v) else float(v)

    # fallback: case-insensitive
    idx_map = {str(i).strip().casefold(): i for i in df_corr.index}
    col_map = {str(c).strip().casefold(): c for c in df_corr.columns}

    a2 = idx_map.get(a.casefold())
    b2 = col_map.get(b.casefold())

    if a2 is not None and b2 is not None:
        v = df_corr.loc[a2, b2]
        return None if pd.isna(v) else float(v)

    return None

def values_on_date_full(df_fund: pd.DataFrame, date_col: str) -> tuple[pd.DataFrame, float]:
    """
    Retorna tabela com colunas do ativo + Valor na data + %PL.
    IMPORTANTE: soma apenas linhas onde _TIPO está preenchido (só ativos).
    """
    extras_presentes = [c for c in COLUNAS_EXTRAS_SITE if c in df_fund.columns]

    tmp = df_fund[[COL_CODIGO, COL_DESCRICAO, "_TIPO"] + extras_presentes + [date_col]].copy()
    tmp = tmp.rename(columns={date_col: "Valor_$"})

    tmp[COL_CODIGO] = tmp[COL_CODIGO].astype(str).replace("nan", "").str.strip()
    tmp[COL_DESCRICAO] = tmp[COL_DESCRICAO].astype(str).replace("nan", "").str.strip()
    tmp["_TIPO"] = tmp["_TIPO"].astype(str).replace("nan", "").str.strip()

    for col in extras_presentes:
        tmp[col] = tmp[col].astype(str).replace("nan", "").str.strip()

    # só ativos
    tmp = tmp[tmp["_TIPO"] != ""].copy()

    group_cols = [COL_CODIGO, COL_DESCRICAO, "_TIPO"] + extras_presentes

    tmp = tmp.groupby(group_cols, as_index=False, dropna=False)["Valor_$"].sum()
    tmp = tmp[tmp["Valor_$"] != 0].copy()

    total = float(tmp["Valor_$"].sum()) if len(tmp) else 0.0
    tmp["%PL"] = (tmp["Valor_$"] / total) if total != 0 else 0.0

    tmp = tmp.sort_values("Valor_$", ascending=False)

    tmp = tmp.rename(
        columns={
            COL_CODIGO: "Código",
            COL_DESCRICAO: "Descrição",
            "_TIPO": "Tipo de Investimento",
        }
    )

    cols_saida = [
        "Código",
        "ISIN",
        "Descrição",
        "Emissor",
        "Classe",
        "Tipo de Investimento",
        "Vencimento",
    ]
    cols_saida = [c for c in cols_saida if c in tmp.columns] + ["Valor_$", "%PL"]
    tmp = tmp[cols_saida]

    return tmp, total

def values_on_date_raw(df_fund: pd.DataFrame, date_col: str) -> tuple[pd.DataFrame, float]:
    """
    Retorna a carteira da data exatamente como está na planilha,
    sem agrupar e sem somar linhas.
    Usar somente na Análise de fundo individual.
    """
    extras_presentes = [c for c in COLUNAS_EXTRAS_SITE if c in df_fund.columns]

    tmp = df_fund[[COL_CODIGO, COL_DESCRICAO, "_TIPO"] + extras_presentes + [date_col]].copy()
    tmp = tmp.rename(columns={date_col: "Valor_$"})

    tmp[COL_CODIGO] = tmp[COL_CODIGO].astype(str).replace("nan", "").str.strip()
    tmp[COL_DESCRICAO] = tmp[COL_DESCRICAO].astype(str).replace("nan", "").str.strip()
    tmp["_TIPO"] = tmp["_TIPO"].astype(str).replace("nan", "").str.strip()

    for col in extras_presentes:
        tmp[col] = tmp[col].astype(str).replace("nan", "").str.strip()

    tmp = tmp[tmp["_TIPO"] != ""].copy()

    tmp["Valor_$"] = tmp["Valor_$"].apply(to_number)
    tmp = tmp[tmp["Valor_$"] != 0].copy()

    total = float(tmp["Valor_$"].sum()) if len(tmp) else 0.0
    tmp["%PL"] = (tmp["Valor_$"] / total) if total != 0 else 0.0

    tmp = tmp.rename(
        columns={
            COL_CODIGO: "Código",
            COL_DESCRICAO: "Descrição",
            "_TIPO": "Tipo de Investimento",
        }
    )

    cols_saida = [
        "Código",
        "ISIN",
        "Descrição",
        "Emissor",
        "Classe",
        "Tipo de Investimento",
        "Vencimento",
    ]
    cols_saida = [c for c in cols_saida if c in tmp.columns] + ["Valor_$", "%PL"]

    return tmp[cols_saida].copy(), total

def tab_common_only(tab_a: pd.DataFrame, tab_b: pd.DataFrame, mode: str) -> pd.DataFrame:
    """
    mode:
      - 'common' -> somente comuns
      - 'only_a' -> só no A
      - 'only_b' -> só no B
    """
    def mk_key(df: pd.DataFrame) -> pd.Series:
        cod = df["Código"].astype(str).str.strip()
        desc = df["Descrição"].astype(str).str.strip()
        return cod.where(cod != "", desc)

    a = tab_a.copy()
    b = tab_b.copy()
    a["_KEY"] = mk_key(a)
    b["_KEY"] = mk_key(b)

    set_a = set(a["_KEY"])
    set_b = set(b["_KEY"])

    if mode == "common":
        keys = set_a & set_b
        out = a[a["_KEY"].isin(keys)].copy()
    elif mode == "only_a":
        keys = set_a - set_b
        out = a[a["_KEY"].isin(keys)].copy()
    else:
        keys = set_b - set_a
        out = b[b["_KEY"].isin(keys)].copy()

    cols_saida = [
        "Código",
        "ISIN",
        "Descrição",
        "Emissor",
        "Classe",
        "Tipo de Investimento",
        "Vencimento",
    ]
    cols_saida = [c for c in cols_saida if c in out.columns]
    out = out[cols_saida].drop_duplicates()
    out = out.sort_values(["Código", "Descrição"])
    return out


def format_brl(v: float) -> str:
    s = f"{v:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"R$ {s}"

def format_num_tabela(v):
    if pd.isna(v) or v == "":
        return ""
    return f"{float(v):,.2f}"

def format_pct_tabela(v):
    if pd.isna(v) or v == "":
        return ""
    return f"{float(v) * 100:,.2f}%"

MAPA_RENOMEAR_EXIBICAO = {
    "ISIN": "ISIN/CNPJ",
    "Valor_$": "Valor (R$)",
    "Valor Base (R$)": "Valor Base (R$)",
    "Patrimonio": "Patrimônio (R$)",
}

def renomear_para_exibicao(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns=MAPA_RENOMEAR_EXIBICAO)

def get_formatters(df: pd.DataFrame) -> dict:
    formatters = {}
    for c in df.columns:
        if c in ["Valor (R$)", "Valor Base (R$)", "Variação (R$)", "Patrimônio (R$)"]:
            formatters[c] = format_num_tabela
        elif c in ["%PL", "%PL Base", "Variação (%PL)", "Variação (%)"]:
            formatters[c] = format_pct_tabela
    return formatters

def get_column_config_padrao() -> dict:
    return {
        "Valor_$": st.column_config.NumberColumn("Valor (R$)", format="%,.2f"),
        "%PL": st.column_config.NumberColumn("%PL", format="%.2f%%"),
        "Variação (R$)": st.column_config.NumberColumn("Variação (R$)", format="%.2f"),
        "Variação (%PL)": st.column_config.NumberColumn("Variação (%PL)", format="%.2f%%"),
        "Variação (%)": st.column_config.NumberColumn("Variação (%)", format="%.2f%%"),
        "Patrimonio": st.column_config.NumberColumn("Patrimônio (R$)", format="%,.2f"),
    }

def _prep_pesos_unicos(tab: pd.DataFrame) -> pd.DataFrame:
    t = tab.copy()

    # chave robusta: usa Código; se vazio, usa Descrição
    cod = t["Código"].astype(str).str.strip()
    desc = t["Descrição"].astype(str).str.strip()
    t["_KEY"] = cod.where(cod != "", desc)

    # garante numérico
    t["%PL"] = pd.to_numeric(t["%PL"], errors="coerce").fillna(0.0)

    # ✅ colapsa duplicados (1 linha por ativo)
    t = t.groupby("_KEY", as_index=False)["%PL"].sum()

    return t


def calcular_overlap_ponderado_tabs(tab_a: pd.DataFrame, tab_b: pd.DataFrame) -> float:
    a = _prep_pesos_unicos(tab_a)
    b = _prep_pesos_unicos(tab_b)

    base = a.merge(b, on="_KEY", how="outer", suffixes=("_A", "_B")).fillna(0.0)
    overlap = float(np.minimum(base["%PL_A"], base["%PL_B"]).sum())

    # segurança: overlap sempre entre 0 e 1
    return max(0.0, min(1.0, overlap))

    # responde a pergunta: quanto do portfólio está realmente compartilhado pelos dois fundos, considerando os pesos?
    # Somando as partes em comum dos ativos, os dois fundos compartilham X% da carteira.
    # Quanto eles dividem da carteira


def calcular_cosine_tabs(tab_a: pd.DataFrame, tab_b: pd.DataFrame) -> float:
    a = _prep_pesos_unicos(tab_a)
    b = _prep_pesos_unicos(tab_b)

    base = a.merge(b, on="_KEY", how="outer", suffixes=("_A", "_B")).fillna(0.0)

    vec_a = base["%PL_A"].to_numpy(dtype=float)
    vec_b = base["%PL_B"].to_numpy(dtype=float)

    na = np.linalg.norm(vec_a)
    nb = np.linalg.norm(vec_b)
    if na == 0 or nb == 0:
        return 0.0

    cosine = float(np.dot(vec_a, vec_b) / (na * nb))
    return max(0.0, min(1.0, cosine))

    # responde a pergunta: os dois fundos tem uma distribuição parecida de pesos entre os ativos?
    # Não mostra o quanto de carteira está em comum, mas sim o quão alinhados são os vetores de alocação.
    # Quando o desenho da carteira é parecido.

# ==========================
# FUNÇÃO PRINCIPAL (UI)
# ==========================
def render_comparacao_fundos(
    pasta: str,
    arquivo_xlsx: str = ARQUIVO_XLSX_PADRAO,
    titulo: str = "Comparação entre Fundos",
    abas_permitidas: Optional[List[str]] = None,
) -> None:
    """
    Renderiza a comparação entre fundos a partir de um Excel com várias abas (cada aba = fundo).

    - pasta: caminho da pasta onde está o arquivo
    - arquivo_xlsx: nome do arquivo (default tratado.xlsx)
    - titulo: título mostrado na página/aba
    - abas_permitidas: para restringit é só selecionar quais fundos aparecem no selectbox (por nome da aba)
    """
    aplicar_ui_padrao()

    pasta_path = Path(pasta)
    arq = pasta_path / arquivo_xlsx

    if not arq.exists():
        st.error(f"Não encontrei {arq}. Coloque o arquivo nessa pasta.")
        st.stop()

    wb = load_workbook_from_path(str(arq))

    funds_processed: Dict[str, Tuple[pd.DataFrame, List[str], Dict[str, pd.Timestamp]]] = {}
    erros = {}

    for nome, df in wb.items():
        if abas_permitidas is not None and nome not in abas_permitidas:
            continue
        try:
            df_f, dcols, dmap = build_fund_table(df)
            funds_processed[nome] = (df_f, dcols, dmap)
        except Exception as e:
            erros[nome] = str(e)

    if erros:
        with st.expander("Abas com problema"):
            st.write(erros)

    fundos = sorted(list(funds_processed.keys()))
    if len(fundos) < 2:
        st.error("Precisa ter pelo menos 2 fundos (abas) para comparar.")
        st.stop()

    render_sectionbar("Selecione os fundos", "sectionbar-fundos")
    c1, c2, c3 = st.columns([2.5, 2.5, 2])

    with c1:
        fundo_a = st.selectbox("Fundo A:", fundos, index=0, key=f"{titulo}_fundo_a")
    with c2:
        fundo_b = st.selectbox("Fundo B:", fundos, index=1, key=f"{titulo}_fundo_b")

    df_a, dcols_a, dmap_a = funds_processed[fundo_a]
    df_b, dcols_b, dmap_b = funds_processed[fundo_b]

    datas_a = sorted({dmap_a[c] for c in dcols_a})
    datas_b = sorted({dmap_b[c] for c in dcols_b})
    datas_comuns = sorted(list(set(datas_a) & set(datas_b)))

    if not datas_comuns:
        st.error("Não existem datas em comum entre os fundos.")
        st.stop()

    datas_fmt = [d.strftime("%d/%m/%Y") for d in datas_comuns]
    with c3:
        data_escolhida_fmt = st.selectbox("Data escolhida", datas_fmt, index=len(datas_fmt) - 1, key=f"{titulo}_data")

    data_escolhida = pd.to_datetime(data_escolhida_fmt, dayfirst=True).normalize()

    col_data_a = next(c for c in dcols_a if dmap_a[c] == data_escolhida)
    col_data_b = next(c for c in dcols_b if dmap_b[c] == data_escolhida)

    tab_a, total_a = values_on_date_full(df_a, col_data_a)
    tab_b, total_b = values_on_date_full(df_b, col_data_b)

    max_rows = st.selectbox(
        "Quantidade de linhas:",
        options=[10, 20, 50, "Todas"],
        index=2,
        key=f"{titulo}_max_rows"
    )

    def altura_por_linhas(qtd, total_linhas):
        if qtd == "Todas":
            return max(120, int(35 * total_linhas + 40))
        return int(35 * int(qtd) + 40)

    total_linhas_fundos = max(len(tab_a), len(tab_b))
    altura_fundos = altura_por_linhas(max_rows, total_linhas_fundos)
    altura_listas = 300

    colunas_disponiveis_a = get_colunas_disponiveis(tab_a)
    colunas_disponiveis_b = get_colunas_disponiveis(tab_b)
    colunas_disponiveis = [
        c for c in ORDEM_COLUNAS_SITE
        if c in colunas_disponiveis_a or c in colunas_disponiveis_b
    ]

    colunas_default = [c for c in COLUNAS_PADRAO_INICIAIS if c in colunas_disponiveis]

    key_colunas_tabela = f"{titulo}_colunas_tabela"

    if key_colunas_tabela not in st.session_state:
        st.session_state[key_colunas_tabela] = colunas_default.copy()
    else:
        # mantém só colunas que ainda existem na tabela atual
        st.session_state[key_colunas_tabela] = [
            c for c in st.session_state[key_colunas_tabela]
            if c in colunas_disponiveis
        ]

        # se por algum motivo ficar vazio, volta para o default
        if not st.session_state[key_colunas_tabela]:
            st.session_state[key_colunas_tabela] = colunas_default.copy()

    with st.expander("Escolher colunas da tabela"):
        colunas_escolhidas = st.multiselect(
            "Colunas visíveis",
            options=colunas_disponiveis,
            key=key_colunas_tabela,
        )

    cols_tab_a = [c for c in colunas_escolhidas if c in tab_a.columns]
    cols_tab_b = [c for c in colunas_escolhidas if c in tab_b.columns]

    tab_a_show = renomear_para_exibicao(tab_a[cols_tab_a].copy())
    tab_b_show = renomear_para_exibicao(tab_b[cols_tab_b].copy())

    left, right = st.columns(2)
    with left:
        st.markdown(
            f"<div style='font-size:22px; font-weight:700;'>Fundo A — {fundo_a}</div>",
            unsafe_allow_html=True
        )
        st.dataframe(
            tab_a_show.style.format(get_formatters(tab_a_show), na_rep=""),
            use_container_width=True,
            height=altura_fundos,
        )

    with right:
        st.markdown(
            f"<div style='font-size:22px; font-weight:700;'>Fundo B — {fundo_b}</div>",
            unsafe_allow_html=True
        )
        st.dataframe(
            tab_b_show.style.format(get_formatters(tab_b_show), na_rep=""),
            use_container_width=True,
            height=altura_fundos,
        )

    p1, p2 = st.columns(2)
    with p1:
        st.markdown(
            f'<div class="kpi-box">Patrimônio Líquido: {format_brl(total_a)}</div>',
            unsafe_allow_html=True
        )
    with p2:
        st.markdown(
            f'<div class="kpi-box">Patrimônio Líquido: {format_brl(total_b)}</div>',
            unsafe_allow_html=True
        )

    render_sectionbar("Ativos em comum e exclusivos", "sectionbar-ativos")

    t_common = tab_common_only(tab_a, tab_b, "common")
    t_only_a = tab_common_only(tab_a, tab_b, "only_a")
    t_only_b = tab_common_only(tab_a, tab_b, "only_b")

    # colunas visíveis nessa seção
    colunas_disponiveis_ativos = [
        c for c in ORDEM_COLUNAS_SITE
        if c in t_common.columns or c in t_only_a.columns or c in t_only_b.columns
    ]

    colunas_default_ativos = [c for c in get_colunas_iniciais(tab_a) if c in colunas_disponiveis_ativos]

    key_colunas_ativos = (
        f"{titulo}_colunas_ativos_comuns_"
        f"{fundo_a}_{fundo_b}_{data_escolhida_fmt}"
    )

    if key_colunas_ativos not in st.session_state:
        st.session_state[key_colunas_ativos] = colunas_default_ativos.copy()
    else:
        st.session_state[key_colunas_ativos] = [
            c for c in st.session_state[key_colunas_ativos]
            if c in colunas_disponiveis_ativos
        ]

        if not st.session_state[key_colunas_ativos]:
            st.session_state[key_colunas_ativos] = colunas_default_ativos.copy()

    with st.expander("Escolher colunas dos ativos em comum e exclusivos"):
        colunas_escolhidas_ativos = st.multiselect(
            "Colunas visíveis nesta seção",
            options=colunas_disponiveis_ativos,
            key=key_colunas_ativos,
        )

    cols_common = [c for c in colunas_escolhidas_ativos if c in t_common.columns]
    cols_only_a = [c for c in colunas_escolhidas_ativos if c in t_only_a.columns]
    cols_only_b = [c for c in colunas_escolhidas_ativos if c in t_only_b.columns]

    t_common_show = renomear_para_exibicao(t_common[cols_common].copy())
    t_only_a_show = renomear_para_exibicao(t_only_a[cols_only_a].copy())
    t_only_b_show = renomear_para_exibicao(t_only_b[cols_only_b].copy())

    # Totais
    n_common = len(t_common)
    n_only_a = len(t_only_a)
    n_only_b = len(t_only_b)

    a1, a2, a3 = st.columns(3)

    with a1:
        st.markdown("<div style='font-size:20px; font-weight:700;'>Ativos em comum</div>", unsafe_allow_html=True)
        st.metric("Número total", n_common)
        st.dataframe(
            t_common_show.style.format(get_formatters(t_common_show), na_rep=""),
            use_container_width=True,
            height=altura_listas,
        )

    with a2:
        st.markdown(f"<div style='font-size:20px; font-weight:700;'>Ativos exclusivos - {fundo_a}</div>", unsafe_allow_html=True)
        st.metric("Número total", n_only_a)
        st.dataframe(
            t_only_a_show.style.format(get_formatters(t_only_a_show), na_rep=""),
            use_container_width=True,
            height=altura_listas,
        )

    with a3:
        st.markdown(f"<div style='font-size:20px; font-weight:700;'>Ativos - {fundo_b}</div>", unsafe_allow_html=True)
        st.metric("Número total", n_only_b)
        st.dataframe(
            t_only_b_show.style.format(get_formatters(t_only_b_show), na_rep=""),
            use_container_width=True,
            height=altura_listas,
        )

    render_sectionbar("Similaridade entre fundos", "sectionbar-similaridade")

    overlap = calcular_overlap_ponderado_tabs(tab_a, tab_b)
    cosine = calcular_cosine_tabs(tab_a, tab_b)

    # correlação via planilha
    arq_corr = pasta_path / "Capa e correlação - site.xlsx"
    corr = None
    if arq_corr.exists():
        try:
            df_corr = load_correlacoes(str(arq_corr))
            corr = get_correlacao(df_corr, fundo_a, fundo_b)
        except Exception:
            corr = None

    k1, k2, k3 = st.columns(3)
    with k1:
        st.metric("Sobreposição (ponderada)", f"{overlap:.2%}")
    with k2:
        st.metric("Similaridade (Cosine)", f"{cosine:.2f}")
    with k3:
        st.metric("Correlação", "N/A" if corr is None else f"{corr:.2f}")

    render_sectionbar("Exportar", "sectionbar-exportar")

    if st.button("Gerar Excel da comparação", key=f"{titulo}_btn_export"):
        out_path = pasta_path / f"comparacao_{fundo_a}_vs_{fundo_b}_{data_escolhida.strftime('%Y-%m-%d')}.xlsx"
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            tab_a.to_excel(writer, index=False, sheet_name=f"{fundo_a}_carteira"[:31])
            tab_b.to_excel(writer, index=False, sheet_name=f"{fundo_b}_carteira"[:31])
            t_common.to_excel(writer, index=False, sheet_name="Comuns")
            t_only_a.to_excel(writer, index=False, sheet_name=f"So_{fundo_a}"[:31])
            t_only_b.to_excel(writer, index=False, sheet_name=f"So_{fundo_b}"[:31])
            pd.DataFrame(
                {"Fundo": [fundo_a, fundo_b], "Data": [data_escolhida_fmt, data_escolhida_fmt], "Patrimonio": [total_a, total_b]}
            ).to_excel(writer, index=False, sheet_name="Patrimonio")

        st.success(f"Arquivo gerado em: {out_path}")

def _fund_wide_to_long(df_fund: pd.DataFrame, fund_name: str, date_cols: List[str]) -> pd.DataFrame:
    """1 fundo (wide) -> long: Fundo, Codigo, Descrição, Tipo, Data, Ano, Valor"""
    base = df_fund[[COL_CODIGO, COL_DESCRICAO, "_TIPO"] + date_cols].copy()
    long = base.melt(
        id_vars=[COL_CODIGO, COL_DESCRICAO, "_TIPO"],
        value_vars=date_cols,
        var_name="Data",
        value_name="Valor",
    )
    long["Fundo"] = fund_name
    long["Data"] = pd.to_datetime(long["Data"], errors="coerce", dayfirst=True)
    long["Ano"] = long["Data"].dt.year

    # limpeza
    long[COL_CODIGO] = long[COL_CODIGO].astype(str).replace("nan", "").str.strip()
    long[COL_DESCRICAO] = long[COL_DESCRICAO].astype(str).replace("nan", "").str.strip()
    long["_TIPO"] = long["_TIPO"].astype(str).replace("nan", "").str.strip()

    long["Valor"] = long["Valor"].fillna(0.0).apply(to_number)

    # mantém só linhas com tipo preenchido e valor != 0
    long = long[(long["_TIPO"] != "") & (long["Valor"] != 0)].copy()
    return long


@st.cache_data(show_spinner=False)
def _build_long_database(path_xlsx: str) -> pd.DataFrame:
    wb = load_workbook_from_path(path_xlsx)
    parts = []
    for nome, df in wb.items():
        try:
            df_f, dcols, _ = build_fund_table(df)
            parts.append(_fund_wide_to_long(df_f, nome, dcols))
        except Exception:
            continue
    if not parts:
        return pd.DataFrame(columns=["Fundo", COL_CODIGO, COL_DESCRICAO, "_TIPO", "Data", "Ano", "Valor"])
    return pd.concat(parts, ignore_index=True)

def listar_opcoes_busca(wb: Dict[str, pd.DataFrame], nome_coluna: str) -> list[str]:
    valores = []

    for _, df_raw in wb.items():
        try:
            df_f, _, _ = build_fund_table(df_raw)
        except Exception:
            continue

        if nome_coluna not in df_f.columns:
            continue

        serie = df_f[nome_coluna].copy()

        if nome_coluna == "Vencimento":
            serie_fmt = pd.to_datetime(serie, errors="coerce", dayfirst=True)
            serie = serie_fmt.dt.strftime("%d/%m/%Y").fillna(
                serie.astype(str).replace("nan", "").str.strip()
            )
        else:
            serie = serie.astype(str).replace("nan", "").str.strip()

        valores.extend([v for v in serie if str(v).strip() != ""])

    return sorted({str(v).strip() for v in valores if str(v).strip() not in {"", "nan", "None", "NaT"}})

def render_pesquisa_ativo_por_data(
    pasta: str,
    arquivo_xlsx: str = ARQUIVO_XLSX_PADRAO,
    abas_permitidas: Optional[List[str]] = None,
) -> None:
    """
    Busca um ATIVO e mostra:
      - todos os fundos que já tiveram o ativo
      - em quais DATAS ele apareceu
      - valor em cada data (tabela Fundo x Datas)
      - gráfico de evolução com 3 modos: Valor, Normalizado, %PL
    """

    render_sectionbar("Comparação entre ativos", "sectionbar-comparacao")

    pasta_path = Path(pasta)
    arq = pasta_path / arquivo_xlsx
    if not arq.exists():
        st.error(f"Não encontrei {arq}.")
        st.stop()

    wb = load_workbook_from_path(str(arq))
    if abas_permitidas is not None:
        wb = {k: v for k, v in wb.items() if k in abas_permitidas}

        # -------- inputs --------
        filtros_cfg = [
        ("Código", COL_CODIGO),
        ("ISIN", "ISIN"),
        ("Descrição", COL_DESCRICAO),
        ("Emissor", "Emissor"),
        ("Classe", "Classe"),
        ("Tipo de Investimento", "_TIPO"),
        ("Vencimento", "Vencimento"),
    ]

    opcoes_por_coluna = {
        nome_coluna: listar_opcoes_busca(wb, nome_coluna)
        for _, nome_coluna in filtros_cfg
    }

    filtros = {}
    cols_ui = st.columns(4)

    for i, (rotulo, nome_coluna) in enumerate(filtros_cfg):
        with cols_ui[i % 4]:
            digitado = st.text_input(
                rotulo,
                "",
                key=f"pesquisa_ativo_{nome_coluna}"
            ).strip()

            escolhido = None
            if digitado:
                sugestoes = [
                    v for v in opcoes_por_coluna[nome_coluna]
                    if digitado.lower() in str(v).lower()
                ][:200]

                if sugestoes:
                    escolhido = st.selectbox(
                        f"Sugestões de {rotulo.lower()}",
                        options=sugestoes,
                        index=None,
                        placeholder=f"Selecione {rotulo.lower()}...",
                        key=f"pesquisa_ativo_sug_{nome_coluna}"
                    )

            filtros[nome_coluna] = (escolhido or digitado).strip()

    if not any(filtros.values()):
        st.info("Busque em ao menos um item.")
        return

    modo = st.radio(
        "Modo de visualização:",
        ["Valor (R$)", "Normalizado (Base 100)", "% do PL"],
        horizontal=True,
    )

    # -------- resultado --------
    linhas = []
    fundos_com_ativo = []
    plot_rows = []  # para o gráfico (Fundo, Data, Valor/Pct)
    for nome_fundo, df_raw in wb.items():
        try:
            df_f, date_cols, dmap = build_fund_table(df_raw)
        except Exception:
            continue

        df = df_f.copy()

        # garante que todas as colunas de busca existam
        for col in [COL_CODIGO, "ISIN", COL_DESCRICAO, "Emissor", "Classe", "_TIPO", "Vencimento"]:
            if col not in df.columns:
                df[col] = ""

            df[col] = (
                df[col]
                .astype(str)
                .replace("nan", "")
                .str.strip()
            )

        # --- filtro (AND) em todas as colunas preenchidas ---
        mask = pd.Series(True, index=df.index)

        for col, termo in filtros.items():
            if not termo:
                continue

            if col == COL_CODIGO:
                cod_exato = df[col].str.upper() == termo.upper()
                if cod_exato.any():
                    mask &= cod_exato
                else:
                    mask &= df[col].str.contains(termo, case=False, na=False)
            else:
                mask &= df[col].str.contains(termo, case=False, na=False)

        # valores do ATIVO por data (somando caso haja mais de 1 linha)
        sub = df.loc[mask, date_cols].copy()
        if sub.empty:
            continue

        vals = sub.sum(axis=0)
        vals = vals[vals != 0]
        if vals.empty:
            continue

        # total do fundo por data (para %PL)
        total_fundo = df.loc[df["_TIPO"] != "", date_cols].sum(axis=0)
        total_fundo = total_fundo.replace(0, pd.NA)

        fundos_com_ativo.append(nome_fundo)

        # linha (tabela Fundo x Datas)
        row = {"Fundo": nome_fundo}

        for col in vals.index:
            dt = dmap.get(col)
            if dt is None:
                continue
            dt_fmt = dt.strftime("%d/%m/%Y")

            v_abs = float(vals[col])

            if modo == "% do PL":
                v_tot = total_fundo.get(col)
                v = float(v_abs / v_tot) if pd.notna(v_tot) and float(v_tot) != 0 else 0.0
            else:
                v = v_abs

            row[dt_fmt] = v

            # para o gráfico (Data real)
            plot_rows.append(
                {"Fundo": nome_fundo, "Data": pd.Timestamp(dt), "Valor": v}
            )

        linhas.append(row)

    if not linhas:
        st.warning("Não encontrei esse ativo/critério em nenhum fundo.")
        return

    st.write(f"**Fundos que já tiveram o ativo:** {len(set(fundos_com_ativo))}")
    st.write(", ".join(sorted(set(fundos_com_ativo))))

    # -------- TABELA Fundo x Datas --------
    out = pd.DataFrame(linhas).fillna(0.0)
    date_cols_fmt = [c for c in out.columns if c != "Fundo"]
    date_cols_fmt_sorted = sorted(date_cols_fmt, key=lambda x: pd.to_datetime(x, dayfirst=True))
    out_num = out[["Fundo"] + date_cols_fmt_sorted].copy()

    # Normalização base 100 (por fundo)
    if modo == "Normalizado (Base 100)":
        for i in range(len(out_num)):
            serie = out_num.loc[i, date_cols_fmt_sorted].astype(float)
            # primeiro ponto != 0
            first = next((float(x) for x in serie.values if float(x) != 0), None)
            if first is None or first == 0:
                continue
            out_num.loc[i, date_cols_fmt_sorted] = (serie / first) * 100.0

    # formatação para exibição
    out_show = out_num.copy()
    for c in date_cols_fmt_sorted:
        if modo == "% do PL":
            out_show[c] = out_show[c].apply(lambda v: (f"{float(v)*100:,.2f}%".replace(",", "_").replace(".", ",").replace("_", ".")) if float(v) != 0 else "")
        elif modo == "Normalizado (Base 100)":
            out_show[c] = out_show[c].apply(lambda v: (f"{float(v):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")) if float(v) != 0 else "")
        else:
            out_show[c] = out_show[c].apply(lambda v: format_brl(float(v)) if float(v) != 0 else "")

    st.dataframe(out_show, use_container_width=True, height=420)

# -------- GRÁFICO --------
    st.markdown("### Evolução do ativo (ao longo do tempo)")

    plot_df = pd.DataFrame(plot_rows).copy()
    plot_df = plot_df.sort_values("Data")

    if modo == "Normalizado (Base 100)":
        def _normalizar_grupo(grp):
            grp = grp.sort_values("Data").copy()
            base = grp["Valor"].replace(0, pd.NA).dropna()
            if base.empty:
                grp["Valor_plot"] = grp["Valor"]
            else:
                primeiro = float(base.iloc[0])
                grp["Valor_plot"] = (grp["Valor"] / primeiro) * 100.0
            return grp

        plot_df = (
            plot_df.groupby("Fundo", group_keys=False)
            .apply(_normalizar_grupo)
            .reset_index(drop=True)
        )
    else:
        plot_df["Valor_plot"] = plot_df["Valor"]

    if modo == "Valor (R$)":
        titulo_y = "Valor (R$)"
        formato_tooltip = ",.2f"
    elif modo == "Normalizado (Base 100)":
        titulo_y = "Base 100"
        formato_tooltip = ",.2f"
    else:
        titulo_y = "% do PL"
        formato_tooltip = ".2%"

    hover = alt.selection_point(
        nearest=True,
        on="pointermove",
        fields=["Data"],
        empty=False
    )

    base = alt.Chart(plot_df).encode(
        x=alt.X(
            "Data:T",
            title="Data",
            axis=alt.Axis(format="%b/%Y", labelAngle=0)
        ),
        color=alt.Color(
            "Fundo:N",
            title="Fundo",
            legend=alt.Legend(orient="bottom")
        )
    )

    lines = base.mark_line(strokeWidth=2.5).encode(
        y=alt.Y("Valor_plot:Q", title=titulo_y, axis=alt.Axis(grid=True))
    )

    selectors = alt.Chart(plot_df).mark_point(opacity=0).encode(
        x="Data:T"
    ).add_params(hover)

    points = lines.mark_circle(size=70).encode(
        opacity=alt.condition(hover, alt.value(1), alt.value(0))
    )

    rule = alt.Chart(plot_df).mark_rule(opacity=0).encode(
        x="Data:T"
    ).transform_filter(hover)

    labels = alt.Chart(plot_df).mark_text(
        align="left",
        dx=8,
        dy=-8,
        fontSize=11
    ).encode(
        x="Data:T",
        y="Valor_plot:Q",
        text=alt.condition(
            hover,
            alt.Text("Valor_plot:Q", format=formato_tooltip),
            alt.value("")
        ),
        color=alt.Color("Fundo:N", legend=None)
    ).transform_filter(hover)

    tooltip_points = alt.Chart(plot_df).mark_circle(opacity=0, size=220).encode(
        x="Data:T",
        y="Valor_plot:Q",
        tooltip=[
            alt.Tooltip("Data:T", title="Data", format="%d/%m/%Y"),
            alt.Tooltip("Fundo:N", title="Fundo"),
            alt.Tooltip("Valor_plot:Q", title=titulo_y, format=formato_tooltip),
        ]
    ).transform_filter(hover)

    zoom = alt.selection_interval(bind="scales", encodings=["x", "y"])

    chart = (
        alt.layer(
            lines,
            selectors,
            points,
            rule,
            labels,
            tooltip_points
        )
        .add_params(zoom)
        .properties(height=420)
    )

    st.altair_chart(chart, use_container_width=True)

def _mk_key_from_tab(df: pd.DataFrame) -> pd.Series:
    codigo = df["Código"].astype(str).replace("nan", "").str.strip() if "Código" in df.columns else ""
    isin = df["ISIN"].astype(str).replace("nan", "").str.strip() if "ISIN" in df.columns else ""
    descricao = df["Descrição"].astype(str).replace("nan", "").str.strip() if "Descrição" in df.columns else ""
    vencimento = df["Vencimento"].astype(str).replace("nan", "").str.strip() if "Vencimento" in df.columns else ""
    tipo = df["Tipo de Investimento"].astype(str).replace("nan", "").str.strip() if "Tipo de Investimento" in df.columns else ""

    return (
        codigo.fillna("") + "||" +
        isin.fillna("") + "||" +
        descricao.fillna("") + "||" +
        vencimento.fillna("") + "||" +
        tipo.fillna("")
    )

def _get_sorted_dates_from_fund(date_cols: List[str], date_map: Dict[str, pd.Timestamp]) -> List[pd.Timestamp]:
    datas = sorted({date_map[c] for c in date_cols if date_map.get(c) is not None})
    return datas


def render_analise_fundo_individual(
    pasta: str,
    arquivo_xlsx: str = ARQUIVO_XLSX_PADRAO,
    titulo: str = "Análise de Carteira (Fundo Individual)",
    abas_permitidas: Optional[List[str]] = None,
) -> None:
    """
    Página para analisar 1 fundo:
      - Seleciona fundo e data
      - Mostra carteira na data com %PL
      - Calcula variação vs carteira base escolhida
      - Mostra 'Ativos que entraram' e 'Ativos que saíram'
    """
    aplicar_ui_padrao()

    pasta_path = Path(pasta)
    arq = pasta_path / arquivo_xlsx

    if not arq.exists():
        st.error(f"Não encontrei {arq}. Coloque o arquivo nessa pasta.")
        st.stop()

    wb = load_workbook_from_path(str(arq))

    funds_processed: Dict[str, Tuple[pd.DataFrame, List[str], Dict[str, pd.Timestamp]]] = {}
    erros = {}

    for nome, df in wb.items():
        if abas_permitidas is not None and nome not in abas_permitidas:
            continue
        try:
            df_f, dcols, dmap = build_fund_table(df)
            funds_processed[nome] = (df_f, dcols, dmap)
        except Exception as e:
            erros[nome] = str(e)

    if erros:
        with st.expander("Abas com problema"):
            st.write(erros)

    fundos = sorted(list(funds_processed.keys()))
    if not fundos:
        st.error("Não encontrei nenhuma aba válida para análise.")
        st.stop()

    render_sectionbar("Selecione o fundo", "sectionbar-fundoindividual")

    c1, c2, c3 = st.columns([2.8, 1.6, 1.6])

    with c1:
        fundo = st.selectbox("Fundo:", fundos, index=0, key=f"{titulo}_fundo")

    df_f, dcols, dmap = funds_processed[fundo]
    datas = _get_sorted_dates_from_fund(dcols, dmap)

    if not datas:
        st.error("Esse fundo não tem colunas de datas válidas.")
        st.stop()

    datas_fmt = [d.strftime("%d/%m/%Y") for d in datas]

    with c2:
        data_escolhida_fmt = st.selectbox(
            "Data para analisar:",
            datas_fmt,
            index=len(datas_fmt) - 1,
            key=f"{titulo}_data"
        )

    data_escolhida = pd.to_datetime(data_escolhida_fmt, dayfirst=True).normalize()

    datas_comp = [d for d in datas if d != data_escolhida]
    datas_comp_fmt = [d.strftime("%d/%m/%Y") for d in datas_comp]

    idx = datas.index(data_escolhida)
    sugestao = datas[idx - 1] if idx > 0 else (datas_comp[0] if datas_comp else None)

    opcoes_base = ["Não comparar"] + datas_comp_fmt

    if sugestao is not None and sugestao.strftime("%d/%m/%Y") in datas_comp_fmt:
        idx_comp = opcoes_base.index(sugestao.strftime("%d/%m/%Y"))
    else:
        idx_comp = 0

    with c3:
        data_comp_fmt = st.selectbox(
            "Carteira para comparar (base):",
            opcoes_base,
            index=idx_comp,
            key=f"{titulo}_data_comp",
        )

    data_comp = None if data_comp_fmt == "Não comparar" else pd.to_datetime(data_comp_fmt, dayfirst=True).normalize()

    col_data = next((c for c in dcols if dmap.get(c) == data_escolhida), None)
    col_data_prev = next((c for c in dcols if dmap.get(c) == data_comp), None) if data_comp is not None else None

    tab_atual, total_atual = values_on_date_raw(df_f, col_data)

    tab_prev = tab_atual.iloc[0:0].copy()
    total_prev = 0.0
    if col_data_prev is not None:
        tab_prev, total_prev = values_on_date_raw(df_f, col_data_prev)
    
    # ===== cálculo de variação
    tab_atual_calc = tab_atual.copy()
    tab_atual_calc["_KEY"] = _mk_key_from_tab(tab_atual_calc)

    if len(tab_prev):
        tab_prev_calc = tab_prev.copy()
        tab_prev_calc["_KEY"] = _mk_key_from_tab(tab_prev_calc)

        prev_valor_map = tab_prev_calc.groupby("_KEY")["Valor_$"].sum().to_dict()
        prev_pl_map = tab_prev_calc.groupby("_KEY")["%PL"].sum().to_dict()

        tab_atual_calc["Valor Base (R$)"] = tab_atual_calc["_KEY"].map(prev_valor_map).fillna(0.0)
        tab_atual_calc["%PL Base"] = tab_atual_calc["_KEY"].map(prev_pl_map).fillna(0.0)

        tab_atual_calc["Variação (R$)"] = tab_atual_calc["Valor_$"] - tab_atual_calc["Valor Base (R$)"]
        tab_atual_calc["Variação (%PL)"] = tab_atual_calc["%PL"] - tab_atual_calc["%PL Base"]

        tab_atual_calc["Variação (%)"] = pd.NA
        mask = tab_atual_calc["Valor Base (R$)"] != 0
        tab_atual_calc.loc[mask, "Variação (%)"] = (
            tab_atual_calc.loc[mask, "Variação (R$)"] / tab_atual_calc.loc[mask, "Valor Base (R$)"]
        )
    else:
        tab_atual_calc["Valor Base (R$)"] = pd.NA
        tab_atual_calc["%PL Base"] = pd.NA
        tab_atual_calc["Variação (R$)"] = pd.NA
        tab_atual_calc["Variação (%PL)"] = pd.NA
        tab_atual_calc["Variação (%)"] = pd.NA   

    cols_base = [
        "Código",
        "ISIN",
        "Descrição",
        "Emissor",
        "Classe",
        "Tipo de Investimento",
        "Vencimento",
    ]
    cols_base = [c for c in cols_base if c in tab_atual_calc.columns]

    colunas_tabela = cols_base + ["Valor_$", "%PL"]

    if len(tab_prev):
        colunas_tabela += ["Valor Base (R$)", "%PL Base", "Variação (R$)", "Variação (%PL)", "Variação (%)"]

    tab_carteira = tab_atual_calc[colunas_tabela].copy()

    # ===== tabelas de entrou / saiu
    if len(tab_prev):
        keys_atual = set(tab_atual_calc["_KEY"])

        tab_prev_calc = tab_prev.copy()
        tab_prev_calc["_KEY"] = _mk_key_from_tab(tab_prev_calc)
        keys_prev = set(tab_prev_calc["_KEY"])

        entrou_keys = keys_atual - keys_prev
        saiu_keys = keys_prev - keys_atual

        ativos_entraram = tab_atual_calc[tab_atual_calc["_KEY"].isin(entrou_keys)].copy()
        ativos_sairam = tab_prev_calc[tab_prev_calc["_KEY"].isin(saiu_keys)].copy()

        ativos_entraram["Variação (R$)"] = ativos_entraram["Valor_$"]
        ativos_entraram["Variação (%PL)"] = ativos_entraram["%PL"]
        ativos_entraram["Variação (%)"] = pd.NA

        ativos_sairam["Variação (R$)"] = -ativos_sairam["Valor_$"]
        ativos_sairam["Variação (%PL)"] = -ativos_sairam["%PL"]
        ativos_sairam["Variação (%)"] = -1.0

        cols_mov = [
            "Código",
            "ISIN",
            "Descrição",
            "Emissor",
            "Classe",
            "Tipo de Investimento",
            "Vencimento",
        ]
        cols_mov = [c for c in cols_mov if c in ativos_entraram.columns] + [
            "Valor_$", "%PL", "Variação (R$)", "Variação (%PL)", "Variação (%)"
        ]

        ativos_entraram = ativos_entraram[cols_mov].copy()
        ativos_sairam = ativos_sairam[cols_mov].copy()
    else:
        cols_mov = [
            "Código",
            "ISIN",
            "Descrição",
            "Emissor",
            "Classe",
            "Tipo de Investimento",
            "Vencimento",
            "Valor_$",
            "%PL",
        ]
        cols_mov = [c for c in cols_mov if c in tab_atual.columns]

        ativos_entraram = tab_atual.iloc[0:0][cols_mov].copy()
        ativos_sairam = tab_atual.iloc[0:0][cols_mov].copy()

    # ===== exibição (carteira)
    render_sectionbar("Carteira", "sectionbar-carteira")

    max_rows = st.selectbox(
        "Quantidade de linhas:",
        options=[10, 20, 50, "Todas"],
        index=2,
        key=f"{titulo}_max_rows"
    )

    def altura_por_linhas(qtd, total_linhas):
        if qtd == "Todas":
            return int(35 * total_linhas + 40)
        return int(35 * int(qtd) + 40)

    altura = altura_por_linhas(max_rows, len(tab_carteira))
    colunas_disponiveis = get_colunas_disponiveis(tab_carteira)

    colunas_default = [
    c for c in [
        "Código",
        "ISIN",
        "Classe",
        "Vencimento",
        "Valor_$",
        "%PL",
        "Valor Base (R$)",
        "%PL Base",
    ] if c in colunas_disponiveis
]

    key_colunas_tabela = f"{titulo}_colunas_tabela_{fundo}_{data_escolhida_fmt}_{data_comp_fmt or 'sem_base'}"

    if key_colunas_tabela not in st.session_state:
        st.session_state[key_colunas_tabela] = colunas_default.copy()
    else:
        st.session_state[key_colunas_tabela] = [
            c for c in st.session_state[key_colunas_tabela]
            if c in colunas_disponiveis
        ]

        if not st.session_state[key_colunas_tabela]:
            st.session_state[key_colunas_tabela] = colunas_default.copy()

    with st.expander("Escolher colunas da tabela"):
        colunas_escolhidas = st.multiselect(
            "Colunas visíveis",
            options=colunas_disponiveis,
            key=key_colunas_tabela,
        )

    cols_tab_carteira = [c for c in colunas_escolhidas if c in tab_carteira.columns]

    def color_variacao(val):
        if pd.isna(val):
            return ""
        if val > 0:
            return "background-color: rgba(0, 180, 0, 0.18);"
        elif val < 0:
            return "background-color: rgba(255, 0, 0, 0.18);"
        return "background-color: rgba(180, 180, 180, 0.10);"

    with st.expander("Filtrar tabela"):
        f1, f2, f3 = st.columns(3)

        with f1:
            codigos_disponiveis = []
            if "Código" in tab_carteira.columns:
                codigos_disponiveis = sorted(
                    [
                        str(x).strip()
                        for x in tab_carteira["Código"].dropna().astype(str).unique()
                        if str(x).strip() != ""
                    ]
                )

            filtro_codigo_select = st.selectbox(
                "Ver código",
                options=["Todos"] + codigos_disponiveis,
                index=0,
                key=f"{titulo}_filtro_codigo_select"
            )

            filtro_codigo_busca = st.text_input(
                "Buscar dentro do código",
                key=f"{titulo}_filtro_codigo_busca"
            )

            descricoes_disponiveis = []
            if "Descrição" in tab_carteira.columns:
                descricoes_disponiveis = sorted(
                    [
                        str(x).strip()
                        for x in tab_carteira["Descrição"].dropna().astype(str).unique()
                        if str(x).strip() != ""
                    ]
                )

            filtro_descricao_select = st.selectbox(
                "Ver descrição",
                options=["Todos"] + descricoes_disponiveis,
                index=0,
                key=f"{titulo}_filtro_descricao_select"
            )

            filtro_descricao_busca = st.text_input(
                "Buscar dentro da descrição",
                key=f"{titulo}_filtro_descricao_busca"
            )

        with f2:
            emissores_disponiveis = []
            if "Emissor" in tab_carteira.columns:
                emissores_disponiveis = sorted(
                    [
                        str(x).strip()
                        for x in tab_carteira["Emissor"].dropna().astype(str).unique()
                        if str(x).strip() != ""
                    ]
                )

            filtro_emissor_select = st.selectbox(
                "Ver emissor",
                options=["Todos"] + emissores_disponiveis,
                index=0,
                key=f"{titulo}_filtro_emissor_select"
            )

            filtro_emissor_busca = st.text_input(
                "Buscar dentro do emissor",
                key=f"{titulo}_filtro_emissor_busca"
            )

            classes_disponiveis = []
            if "Classe" in tab_carteira.columns:
                classes_disponiveis = sorted(
                    [
                        str(x).strip()
                        for x in tab_carteira["Classe"].dropna().astype(str).unique()
                        if str(x).strip() != ""
                    ]
                )

            filtro_classe_select = st.selectbox(
                "Ver classe",
                options=["Todos"] + classes_disponiveis,
                index=0,
                key=f"{titulo}_filtro_classe_select"
            )

            filtro_classe_busca = st.text_input(
                "Buscar dentro da classe",
                key=f"{titulo}_filtro_classe_busca"
            )

        with f3:
            tipos_disponiveis = []
            if "Tipo de Investimento" in tab_carteira.columns:
                tipos_disponiveis = sorted(
                    [
                        str(x).strip()
                        for x in tab_carteira["Tipo de Investimento"].dropna().astype(str).unique()
                        if str(x).strip() != ""
                    ]
                )

            filtro_tipo_select = st.selectbox(
                "Ver tipo de investimento",
                options=["Todos"] + tipos_disponiveis,
                index=0,
                key=f"{titulo}_filtro_tipo_select"
            )

            filtro_tipo_busca = st.text_input(
                "Buscar dentro do tipo de investimento",
                key=f"{titulo}_filtro_tipo_busca"
            )

            vencimentos_disponiveis = []
            if "Vencimento" in tab_carteira.columns:
                vencimentos_disponiveis = sorted(
                    [
                        str(x).strip()
                        for x in tab_carteira["Vencimento"].dropna().astype(str).unique()
                        if str(x).strip() != ""
                    ]
                )

            filtro_vencimento_select = st.selectbox(
                "Ver vencimento",
                options=["Todos"] + vencimentos_disponiveis,
                index=0,
                key=f"{titulo}_filtro_vencimento_select"
            )

            filtro_vencimento_busca = st.text_input(
                "Buscar dentro do vencimento",
                key=f"{titulo}_filtro_vencimento_busca"
            )

    tab_carteira_filtrada = tab_carteira.copy()

    if filtro_codigo_select != "Todos" and "Código" in tab_carteira_filtrada.columns:
        tab_carteira_filtrada = tab_carteira_filtrada[
            tab_carteira_filtrada["Código"].astype(str).str.strip() == filtro_codigo_select
        ]

    if filtro_codigo_busca and "Código" in tab_carteira_filtrada.columns:
        tab_carteira_filtrada = tab_carteira_filtrada[
            tab_carteira_filtrada["Código"].astype(str).str.contains(filtro_codigo_busca, case=False, na=False)
        ]

    if filtro_descricao_select != "Todos" and "Descrição" in tab_carteira_filtrada.columns:
        tab_carteira_filtrada = tab_carteira_filtrada[
            tab_carteira_filtrada["Descrição"].astype(str).str.strip() == filtro_descricao_select
        ]

    if filtro_descricao_busca and "Descrição" in tab_carteira_filtrada.columns:
        tab_carteira_filtrada = tab_carteira_filtrada[
            tab_carteira_filtrada["Descrição"].astype(str).str.contains(filtro_descricao_busca, case=False, na=False)
        ]

    if filtro_emissor_select != "Todos" and "Emissor" in tab_carteira_filtrada.columns:
        tab_carteira_filtrada = tab_carteira_filtrada[
            tab_carteira_filtrada["Emissor"].astype(str).str.strip() == filtro_emissor_select
        ]

    if filtro_emissor_busca and "Emissor" in tab_carteira_filtrada.columns:
        tab_carteira_filtrada = tab_carteira_filtrada[
            tab_carteira_filtrada["Emissor"].astype(str).str.contains(filtro_emissor_busca, case=False, na=False)
        ]

    if filtro_classe_select != "Todos" and "Classe" in tab_carteira_filtrada.columns:
        tab_carteira_filtrada = tab_carteira_filtrada[
            tab_carteira_filtrada["Classe"].astype(str).str.strip() == filtro_classe_select
        ]

    if filtro_classe_busca and "Classe" in tab_carteira_filtrada.columns:
        tab_carteira_filtrada = tab_carteira_filtrada[
            tab_carteira_filtrada["Classe"].astype(str).str.contains(filtro_classe_busca, case=False, na=False)
        ]

    if filtro_tipo_select != "Todos" and "Tipo de Investimento" in tab_carteira_filtrada.columns:
        tab_carteira_filtrada = tab_carteira_filtrada[
            tab_carteira_filtrada["Tipo de Investimento"].astype(str).str.strip() == filtro_tipo_select
        ]

    if filtro_tipo_busca and "Tipo de Investimento" in tab_carteira_filtrada.columns:
        tab_carteira_filtrada = tab_carteira_filtrada[
            tab_carteira_filtrada["Tipo de Investimento"].astype(str).str.contains(filtro_tipo_busca, case=False, na=False)
        ]

    if filtro_vencimento_select != "Todos" and "Vencimento" in tab_carteira_filtrada.columns:
        tab_carteira_filtrada = tab_carteira_filtrada[
            tab_carteira_filtrada["Vencimento"].astype(str).str.strip() == filtro_vencimento_select
        ]

    if filtro_vencimento_busca and "Vencimento" in tab_carteira_filtrada.columns:
        tab_carteira_filtrada = tab_carteira_filtrada[
            tab_carteira_filtrada["Vencimento"].astype(str).str.contains(filtro_vencimento_busca, case=False, na=False)
        ]

    tab_carteira_show = renomear_para_exibicao(tab_carteira_filtrada[cols_tab_carteira].copy())

    subset_variacao = [c for c in ["Variação (R$)", "Variação (%PL)", "Variação (%)"] if c in cols_tab_carteira]
    subset_variacao_show = [MAPA_RENOMEAR_EXIBICAO.get(c, c) for c in subset_variacao]

    styled_tab = (
        tab_carteira_show.style
        .format(get_formatters(tab_carteira_show), na_rep="")
        .map(color_variacao, subset=subset_variacao_show)
    )

    st.dataframe(
        styled_tab,
        use_container_width=True,
        height=altura,
    )

    k1, k2 = st.columns(2)

    with k1:
        st.markdown(
            f'<div class="kpi-box">Patrimônio Líquido ({data_escolhida_fmt}): {format_brl(total_atual)}</div>',
            unsafe_allow_html=True
        )

    with k2:
        if data_comp_fmt:
            st.markdown(
                f'<div class="kpi-box">Patrimônio Líquido Base ({data_comp_fmt}): {format_brl(total_prev)}</div>',
                unsafe_allow_html=True
            )

    # ===== entrou / saiu
    render_sectionbar("Ativos que entraram / saíram", "sectionbar-movimentacao")

    if col_data_prev is None:
        st.info("Selecione uma 'Carteira para comparar (base)' para calcular variação / entrou / saiu.")

    left, right = st.columns(2)

    cols_entrou = [c for c in colunas_escolhidas if c in ativos_entraram.columns]
    cols_saiu = [c for c in colunas_escolhidas if c in ativos_sairam.columns]

    ativos_entraram_show = renomear_para_exibicao(ativos_entraram[cols_entrou].copy())
    ativos_sairam_show = renomear_para_exibicao(ativos_sairam[cols_saiu].copy())

    with left:
        st.markdown("<div style='font-size:20px; font-weight:700;'>Ativos que entraram</div>", unsafe_allow_html=True)
        st.dataframe(
            ativos_entraram_show.style.format(get_formatters(ativos_entraram_show), na_rep=""),
            use_container_width=True,
            height=300,
        )

    with right:
        st.markdown("<div style='font-size:20px; font-weight:700;'>Ativos que saíram</div>", unsafe_allow_html=True)
        st.dataframe(
            ativos_sairam_show.style.format(get_formatters(ativos_sairam_show), na_rep=""),
            use_container_width=True,
            height=300,
        )
