from __future__ import annotations

from pathlib import Path
from io import BytesIO
import re
import unicodedata
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

from ui_padrao import aplicar_ui_padrao, render_titulo_padrao, render_sectionbar
from comparador_base import (
    build_fund_table,
    values_on_date_full,
    load_workbook_from_path,
    parse_date_col,
    to_number,
)

# =========================================================
# CONFIGURAÇÕES DA PÁGINA
# =========================================================
st.set_page_config(page_title="Comparação por Categorias", layout="wide")
aplicar_ui_padrao()
render_titulo_padrao("Comparação por Categorias")

BASE_DIR = Path(__file__).resolve().parent.parent
ARQUIVO_CARTEIRA = BASE_DIR / "carteira_fundos_consolidada.xlsx"

# Caminho informado pelo usuário. Se o app estiver rodando dentro da própria pasta,
# o fallback BASE_DIR também cobre o mesmo diretório.
CAMINHO_PLANILHA_MANUAL = Path(
    r"Z:\Asset Management\Equipe\Lívia\Rascunhos\Comparação Fundo de Terceiros\Fundo-de-Terceiros"
)
NOME_PLANILHA_MANUAL = "Acompanhamento Fundos Terceiros"

FUNDOS_D1 = [
    "GERAES", "ARX Denali", "Capitânia Top", "Riza Lotus", "Western Asset Total",
    "Daycoval Classic", "Iridium Apollo", "Porto Seguro FIRF", "Nu Reserva Imediata",
    "Inter Conservador", "Absolute Atenas",
]

FUNDOS_D30 = [
    "GERAES 30", "Daycoval Classic 30", "Riza Lotus Plus", "ARX Vinson Advisory",
    "Sparta Max Advisory", "Polo Crédito Corporativo", "Iridium Titan Advisory",
    "Porto Seguro Ipê", "Sparta Top Advisory", "Absolute Creta",
]

# Alias para não quebrar caso a aba antiga esteja com erro de digitação.
ALIASES_FUNDOS = {
    "Iridium Titan Advisory": ["Iridium Titan Advisory", "Iridium Titam Advisory"],
    "Iridium Titam Advisory": ["Iridium Titan Advisory", "Iridium Titam Advisory"],
}

FUNDOS_D60 = [
    "Horizonte", "JGP Select", "ARX Everest Advisory", "Polo Total", "Absolute Olimpia",
]

CATEGORIAS_D1_D30 = ["Caixa", "Bancários", "Corporativo", "FIDC", "FII", "Outros"]
CATEGORIAS_D60 = ["Caixa", "Bancários", "Corporativo", "FIDC", "Bonds", "Outros"]

GRUPOS = {
    "Fundo D+1": {"fundos": FUNDOS_D1, "categorias": CATEGORIAS_D1_D30, "classe_css": "sectionbar-fundos"},
    "Fundo D+30": {"fundos": FUNDOS_D30, "categorias": CATEGORIAS_D1_D30, "classe_css": "sectionbar-ativos"},
    "Fundo D+60": {"fundos": FUNDOS_D60, "categorias": CATEGORIAS_D60, "classe_css": "sectionbar-similaridade"},
}

FONTE_CODIGO = "Código"
FONTE_MANUAL = "Planilha manual"

# =========================================================
# HELPERS GERAIS
# =========================================================
def _norm_txt(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s


def _fmt_brl(v) -> str:
    if pd.isna(v):
        return ""
    s = f"{float(v):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"R$ {s}"


def _fmt_pct(v) -> str:
    if pd.isna(v):
        return ""
    return f"{float(v) * 100:,.2f}%".replace(",", "_").replace(".", ",").replace("_", ".")


def _to_pct(x) -> float:
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
        v = float(x)
        return v / 100 if abs(v) > 1 else v
    s = str(x).strip()
    if s in ("", "-", "–"):
        return 0.0
    tem_pct = "%" in s
    v = to_number(s.replace("%", ""))
    if tem_pct or abs(v) > 1:
        return v / 100
    return v


def _parse_data_lista(texto: str) -> set[pd.Timestamp]:
    datas = set()
    partes = re.split(r"[,;\n]", texto or "")
    for p in partes:
        p = p.strip()
        if not p:
            continue
        dt = pd.to_datetime(p, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            datas.add(pd.Timestamp(dt).normalize())
    return datas


def _resolver_planilha_manual() -> Optional[Path]:
    candidatos = []
    for pasta in [CAMINHO_PLANILHA_MANUAL, BASE_DIR]:
        candidatos.extend([
            pasta / f"{NOME_PLANILHA_MANUAL}.xlsx",
            pasta / f"{NOME_PLANILHA_MANUAL}.xlsm",
            pasta / f"{NOME_PLANILHA_MANUAL}.xls",
        ])
    for c in candidatos:
        if c.exists():
            return c
    return None


def _fundos_alias(fundo: str) -> list[str]:
    return ALIASES_FUNDOS.get(fundo, [fundo])


# =========================================================
# CLASSIFICAÇÃO DAS CATEGORIAS A PARTIR DA CARTEIRA COMPLETA
# =========================================================
def classificar_categoria(row: pd.Series, categorias_validas: list[str]) -> str:
    """
    Classifica o ativo pela coluna Classe.
    Se Classe vier vazia, usa Tipo/Descrição/Código apenas como fallback.
    """
    mapa_classes = {
        # Caixa
        "titulo publico federal": "Caixa",

        # Bancários
        "cdb vinculado": "Bancários",
        "cdb/ rdb": "Bancários",
        "cdb/rdb": "Bancários",
        "dpge": "Bancários",
        "letra financeira": "Bancários",
        "letra de cambio/ letra hipotecaria/ letra imobiliaria": "Bancários",
        "letra de cambio": "Bancários",
        "letra hipotecaria": "Bancários",
        "letra imobiliaria": "Bancários",

        # Corporativo
        "bonus de subscricao": "Corporativo",
        "bonus privado": "Corporativo",
        "cra": "Corporativo",
        "cri": "Corporativo",
        "debenture permutavel": "Corporativo",
        "debenture simples": "Corporativo",
        "nota promissoria/ commercial paper/ export note": "Corporativo",
        "nota promissoria": "Corporativo",
        "commercial paper": "Corporativo",
        "export note": "Corporativo",
        "recibo de subscricao": "Corporativo",
        "outros certificados de recebiveis": "Corporativo",

        # FIDC
        "fidc": "FIDC",

        # FII
        "fi imobiliario": "FII",
        "fi participacoes": "FII",

        # Bonds
        "bonds e treasury": "Bonds",
        "fundos offshore": "Bonds",
        "titulo da divida externa": "Bonds",

        # Outros
        "acao ordinaria": "Outros",
        "acao preferencial": "Outros",
        "contrato futuro": "Outros",
        "depositary receipt no exterior(dr)": "Outros",
        "depositary receipt no exterior (dr)": "Outros",
        "depositary receipt no exterior": "Outros",
        "fundos de indice": "Outros",
        "futuro de dap:cupom de di x ipca": "Outros",
        "futuro de ddi:cupom cambial": "Outros",
        "futuro de di1:di de 1 dia": "Outros",
        "futuro de dol:dolar comercial": "Outros",
        "futuro de t10:us t-note 10 anos": "Outros",
        "opcao de compra": "Outros",
        "outros": "Outros",
    }

    classe = _norm_txt(row.get("Classe", ""))
    if classe in mapa_classes:
        categoria = mapa_classes[classe]
        return categoria if categoria in categorias_validas else "Outros"

    # Fallback para casos em que a coluna Classe venha vazia/incompleta.
    texto = " ".join([
        _norm_txt(row.get("Tipo de Investimento", "")),
        _norm_txt(row.get("Descrição", "")),
        _norm_txt(row.get("Código", "")),
    ])

    if "fidc" in texto:
        return "FIDC"
    if any(x in texto for x in ["titulo publico", "titpub", "tesouro", "lft", "ltn", "ntn"]):
        return "Caixa"
    if any(x in texto for x in ["cdb", "rdb", "dpge", "letra financeira", "letra de cambio", "letra hipotecaria", "letra imobiliaria"]):
        return "Bancários"
    if "Bonds" in categorias_validas and any(x in texto for x in ["bond", "treasury", "titulo da divida externa", "divida externa", "fundos offshore", "offshore"]):
        return "Bonds"
    if "FII" in categorias_validas and any(x in texto for x in ["fi imobiliario", "fii", "fundo imobiliario", "fi participacoes", "fip"]):
        return "FII"
    if any(x in texto for x in ["cra", "cri", "debenture", "nota promissoria", "commercial paper", "export note", "certificado de recebiveis", "certificados de recebiveis", "bonus privado"]):
        return "Corporativo"

    return "Outros"


def resumo_por_categoria_da_carteira(
    df_fundo: pd.DataFrame,
    date_cols: list[str],
    date_map: dict,
    data_ref: pd.Timestamp,
    categorias: list[str],
) -> Optional[dict]:
    col_data = next((c for c in date_cols if date_map[c] == data_ref), None)
    if col_data is None:
        return None

    tab, pl = values_on_date_full(df_fundo, col_data)
    if tab.empty or pl == 0:
        return None

    tab = tab.copy()
    tab["Categoria"] = tab.apply(lambda r: classificar_categoria(r, categorias), axis=1)
    soma = tab.groupby("Categoria", as_index=False)["Valor_$"].sum()

    out = {"PL": float(pl), "Origem": "Carteira completa"}
    for cat in categorias:
        valor = float(soma.loc[soma["Categoria"] == cat, "Valor_$"].sum())
        out[f"{cat} R$"] = valor
        out[f"{cat} %"] = valor / pl if pl else 0.0
    return out


# =========================================================
# LEITURA DA PLANILHA MANUAL
# =========================================================
def _identificar_coluna(cols_norm: dict, nomes: list[str]) -> Optional[str]:
    for nome in nomes:
        n = _norm_txt(nome)
        for col_norm, col_original in cols_norm.items():
            if col_norm == n or n in col_norm:
                return col_original
    return None


def _carregar_manual_formato_base(path: Path) -> pd.DataFrame:
    """
    Preferencial: uma aba em formato base, com colunas:
    Data | Grupo | Fundo | PL | Caixa % | Caixa R$ | Bancários % | Bancários R$ | ...
    """
    partes = []
    try:
        sheets = pd.read_excel(path, sheet_name=None, dtype=object)
    except Exception:
        return pd.DataFrame()

    for _, df in sheets.items():
        if df is None or df.empty:
            continue
        cols_norm = {_norm_txt(c): c for c in df.columns}
        col_data = _identificar_coluna(cols_norm, ["Data", "DataRef", "Data para análise"])
        col_fundo = _identificar_coluna(cols_norm, ["Fundo", "Nome"])
        col_grupo = _identificar_coluna(cols_norm, ["Grupo", "Classificação", "Classificacao"])
        col_pl = _identificar_coluna(cols_norm, ["PL", "Patrimônio", "Patrimonio"])

        if not col_data or not col_fundo or not col_pl:
            continue

        tmp = pd.DataFrame()
        tmp["Data"] = pd.to_datetime(df[col_data], errors="coerce", dayfirst=True).dt.normalize()
        tmp["Fundo"] = df[col_fundo].astype(str).str.strip()
        tmp["Grupo"] = df[col_grupo].astype(str).str.strip() if col_grupo else ""
        tmp["PL"] = df[col_pl].apply(to_number)

        for cat in sorted(set(CATEGORIAS_D1_D30 + CATEGORIAS_D60)):
            cat_norm = _norm_txt(cat)
            pct_col = None
            val_col = None
            for col_norm, col_original in cols_norm.items():
                if cat_norm not in col_norm:
                    continue
                if any(x in col_norm for x in ["%", "pct", "percent", "pl"]):
                    pct_col = col_original
                if any(x in col_norm for x in ["r$", "rs", "valor"]):
                    val_col = col_original
                if col_norm == cat_norm and pct_col is None:
                    pct_col = col_original

            tmp[f"{cat} %"] = df[pct_col].apply(_to_pct) if pct_col else 0.0
            tmp[f"{cat} R$"] = df[val_col].apply(to_number) if val_col else 0.0

        partes.append(tmp)

    if not partes:
        return pd.DataFrame()

    out = pd.concat(partes, ignore_index=True)
    out = out[out["Data"].notna() & (out["Fundo"].str.strip() != "")].copy()
    return out


def _extrair_data_da_sheet(nome_sheet: str, raw: pd.DataFrame) -> Optional[pd.Timestamp]:
    # Abas como Jan25, Fev25, ..., Dez25 viram o último dia do mês.
    meses = {
        "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
        "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
    }
    nome_norm = _norm_txt(nome_sheet).replace(" ", "")
    m = re.match(r"^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)(\d{2}|\d{4})$", nome_norm)
    if m:
        mes = meses[m.group(1)]
        ano_txt = m.group(2)
        ano = int("20" + ano_txt) if len(ano_txt) == 2 else int(ano_txt)
        ultimo_dia = pd.Period(f"{ano}-{mes:02d}", freq="M").end_time.normalize()
        return pd.Timestamp(ultimo_dia)

    dt = parse_date_col(nome_sheet)
    if dt is not None:
        return dt

    # tenta encontrar uma data nas primeiras linhas da aba
    for _, row in raw.head(10).iterrows():
        for val in row.head(12):
            dt = parse_date_col(val)
            if dt is not None and dt.year >= 2020:
                return dt
    return None


def _categoria_por_header(valor) -> Optional[str]:
    bruto = str(valor).strip().lower() if pd.notna(valor) else ""
    if bruto.startswith(("δ", "∆", "Δ")) or bruto.startswith("var"):
        return None

    s = _norm_txt(valor)
    if not s or s.startswith("delta") or s.startswith("var") or s.startswith("obs"):
        return None

    mapa = {
        "caixa": "Caixa",
        "bancarios": "Bancários",
        "bancario": "Bancários",
        "corporativo": "Corporativo",
        "corporativos": "Corporativo",
        "fidc": "FIDC",
        "fii": "FII",
        "bonds": "Bonds",
        "bond": "Bonds",
        "outros": "Outros",
        "outro": "Outros",
    }
    for k, v in mapa.items():
        if s == k:
            return v
    return None


def _carregar_manual_formato_layout(path: Path) -> pd.DataFrame:
    """
    Fallback: tenta ler abas no layout visual da Planilha8.
    Requisito: a data precisa estar no nome da aba ou em alguma célula superior.
    """
    partes = []
    try:
        sheets = pd.read_excel(path, sheet_name=None, dtype=object, header=None)
    except Exception:
        return pd.DataFrame()

    for nome_sheet, raw in sheets.items():
        if raw is None or raw.empty:
            continue

        data_sheet = _extrair_data_da_sheet(nome_sheet, raw)
        if data_sheet is None:
            continue

        nrows, ncols = raw.shape

        for i in range(nrows):
            row_vals = [_norm_txt(x) for x in raw.iloc[i, :].tolist()]
            header_positions = [j for j, v in enumerate(row_vals) if v.startswith("fundo d+")]
            if not header_positions:
                continue

            for fund_col in header_positions:
                grupo = str(raw.iat[i, fund_col]).strip()
                if not grupo:
                    continue

                pl_col = None
                for j in range(fund_col + 1, min(fund_col + 5, ncols)):
                    if _norm_txt(raw.iat[i, j]) == "pl":
                        pl_col = j
                        break
                if pl_col is None:
                    pl_col = fund_col + 1

                cat_pos = []
                for j in range(pl_col + 1, ncols):
                    cat = _categoria_por_header(raw.iat[i, j])
                    if cat:
                        cat_pos.append((cat, j))
                        # O layout principal tem 6 categorias. Depois disso começam observações/variações.
                        if len(cat_pos) >= 6:
                            break

                if not cat_pos:
                    continue

                r = i + 1
                while r < nrows:
                    nome_fundo = str(raw.iat[r, fund_col]).strip() if pd.notna(raw.iat[r, fund_col]) else ""
                    nome_norm = _norm_txt(nome_fundo)

                    if not nome_fundo:
                        # uma linha vazia normalmente encerra o bloco
                        break
                    if nome_norm.startswith("fundo d+"):
                        break
                    if nome_norm == "total":
                        r += 1
                        continue

                    pl = to_number(raw.iat[r, pl_col]) if pl_col < ncols else 0.0
                    linha = {"Data": data_sheet, "Grupo": grupo, "Fundo": nome_fundo, "PL": pl}

                    for cat in sorted(set(CATEGORIAS_D1_D30 + CATEGORIAS_D60)):
                        linha[f"{cat} %"] = 0.0
                        linha[f"{cat} R$"] = 0.0

                    for cat, c_pct in cat_pos:
                        c_val = c_pct + 1
                        pct = _to_pct(raw.iat[r, c_pct]) if c_pct < ncols else 0.0
                        val = to_number(raw.iat[r, c_val]) if c_val < ncols else 0.0
                        if val == 0 and pct != 0 and pl != 0:
                            val = pct * pl
                        if pct == 0 and val != 0 and pl != 0:
                            pct = val / pl
                        linha[f"{cat} %"] = pct
                        linha[f"{cat} R$"] = val

                    partes.append(linha)
                    r += 1

    if not partes:
        return pd.DataFrame()

    return pd.DataFrame(partes)


@st.cache_data(show_spinner=False)
def carregar_planilha_manual(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    if not path.exists():
        return pd.DataFrame()

    base = _carregar_manual_formato_base(path)
    layout = _carregar_manual_formato_layout(path)

    partes = [df for df in [base, layout] if df is not None and not df.empty]
    if not partes:
        return pd.DataFrame()

    out = pd.concat(partes, ignore_index=True)
    out["Data"] = pd.to_datetime(out["Data"], errors="coerce", dayfirst=True).dt.normalize()
    out["Fundo"] = out["Fundo"].astype(str).str.strip()
    out["Grupo"] = out.get("Grupo", "").astype(str).str.strip()
    out = out[out["Data"].notna() & (out["Fundo"] != "")].copy()
    return out


def resumo_manual(
    df_manual: pd.DataFrame,
    fundo: str,
    grupo: str,
    data_ref: pd.Timestamp,
    categorias: list[str],
) -> Optional[dict]:
    if df_manual is None or df_manual.empty:
        return None

    nomes = [_norm_txt(x) for x in _fundos_alias(fundo)]
    m = df_manual[
        (df_manual["Data"] == data_ref)
        & (df_manual["Fundo"].apply(_norm_txt).isin(nomes))
    ].copy()

    if "Grupo" in m.columns and m["Grupo"].astype(str).str.strip().ne("").any():
        grupo_norm = _norm_txt(grupo)
        m_grupo = m[m["Grupo"].apply(_norm_txt).str.contains(grupo_norm, na=False)]
        if not m_grupo.empty:
            m = m_grupo

    if m.empty:
        return None

    row = m.iloc[-1]
    pl = float(to_number(row.get("PL", 0)))
    out = {"PL": pl, "Origem": "Planilha manual"}

    for cat in categorias:
        pct = _to_pct(row.get(f"{cat} %", 0))
        val = to_number(row.get(f"{cat} R$", 0))
        if val == 0 and pct != 0 and pl != 0:
            val = pct * pl
        if pct == 0 and val != 0 and pl != 0:
            pct = val / pl
        out[f"{cat} %"] = pct
        out[f"{cat} R$"] = val

    return out


# =========================================================
# CÁLCULO DA TABELA FINAL
# =========================================================
def carregar_carteiras_processadas(path: Path) -> tuple[dict, set[pd.Timestamp]]:
    if not path.exists():
        st.error(f"Não encontrei o arquivo: {path}")
        st.stop()

    wb = load_workbook_from_path(str(path))
    processados = {}
    datas = set()

    for nome, df_raw in wb.items():
        try:
            df_f, dcols, dmap = build_fund_table(df_raw)
            processados[nome] = (df_f, dcols, dmap)
            datas.update(dmap.values())
        except Exception:
            continue

    return processados, datas


def obter_resumo_fundo(
    fundo: str,
    grupo: str,
    data_ref: pd.Timestamp,
    categorias: list[str],
    fonte_dados: str,
    carteiras_processadas: dict,
    df_manual: pd.DataFrame,
) -> dict:
    # Fonte controlada manualmente na tela. Não há fallback automático.
    if fonte_dados == FONTE_CODIGO:
        for nome in _fundos_alias(fundo):
            if nome in carteiras_processadas:
                df_f, dcols, dmap = carteiras_processadas[nome]
                resumo = resumo_por_categoria_da_carteira(df_f, dcols, dmap, data_ref, categorias)
                if resumo is not None:
                    resumo["Origem"] = FONTE_CODIGO
                    return resumo

    elif fonte_dados == FONTE_MANUAL:
        resumo = resumo_manual(df_manual, fundo, grupo, data_ref, categorias)
        if resumo is not None:
            resumo["Origem"] = FONTE_MANUAL
            return resumo

    out = {"PL": 0.0, "Origem": f"Sem dados ({fonte_dados})"}
    for cat in categorias:
        out[f"{cat} %"] = 0.0
        out[f"{cat} R$"] = 0.0
    return out

def montar_ativos_classificados_grupo(
    grupo: str,
    fundos: list[str],
    categorias: list[str],
    data_ref: pd.Timestamp,
    fonte_dados: str,
    carteiras_processadas: dict,
) -> pd.DataFrame:
    """
    Monta uma tabela de conferência dos ativos usados na classificação por categoria.

    Essa tabela é exportada apenas no Excel, não aparece na tela do Streamlit.
    Quando a fonte escolhida for Planilha manual, não há detalhamento por ativo,
    porque a planilha manual traz os dados já consolidados por categoria.
    """

    if fonte_dados != FONTE_CODIGO:
        return pd.DataFrame([{
            "Grupo": grupo,
            "Data": data_ref.strftime("%d/%m/%Y"),
            "Fonte escolhida": fonte_dados,
            "Observação": (
                "Detalhamento por ativo não disponível para Planilha manual. "
                "A planilha manual traz apenas dados consolidados por categoria."
            )
        }])

    partes = []

    for fundo in fundos:
        encontrado = False

        for nome_aba in _fundos_alias(fundo):
            if nome_aba not in carteiras_processadas:
                continue

            df_f, dcols, dmap = carteiras_processadas[nome_aba]

            col_data = next((c for c in dcols if dmap[c] == data_ref), None)

            if col_data is None:
                continue

            tab, pl = values_on_date_full(df_f, col_data)

            if tab.empty:
                partes.append(pd.DataFrame([{
                    "Grupo": grupo,
                    "Data": data_ref.strftime("%d/%m/%Y"),
                    "Fonte escolhida": fonte_dados,
                    "Fundo": fundo,
                    "Aba Carteira": nome_aba,
                    "Observação": "Carteira encontrada, mas sem ativos para essa data."
                }]))
                encontrado = True
                break

            tab = tab.copy()

            if "Classe" in tab.columns:
                tab["Classe Normalizada"] = tab["Classe"].apply(_norm_txt)
            else:
                tab["Classe Normalizada"] = ""

            tab["Categoria"] = tab.apply(
                lambda r: classificar_categoria(r, categorias),
                axis=1
            )

            tab.insert(0, "Grupo", grupo)
            tab.insert(1, "Data", data_ref.strftime("%d/%m/%Y"))
            tab.insert(2, "Fonte escolhida", fonte_dados)
            tab.insert(3, "Fundo", fundo)
            tab.insert(4, "Aba Carteira", nome_aba)
            tab.insert(5, "PL Fundo", pl)

            ordem = [
                "Grupo",
                "Data",
                "Fonte escolhida",
                "Fundo",
                "Aba Carteira",
                "PL Fundo",
                "Código",
                "ISIN",
                "Descrição",
                "Emissor",
                "Classe",
                "Classe Normalizada",
                "Tipo de Investimento",
                "Vencimento",
                "Categoria",
                "Valor_$",
                "%PL",
            ]

            cols_ordenadas = [c for c in ordem if c in tab.columns]
            outras_cols = [c for c in tab.columns if c not in cols_ordenadas]

            partes.append(tab[cols_ordenadas + outras_cols])

            encontrado = True
            break

        if not encontrado:
            partes.append(pd.DataFrame([{
                "Grupo": grupo,
                "Data": data_ref.strftime("%d/%m/%Y"),
                "Fonte escolhida": fonte_dados,
                "Fundo": fundo,
                "Observação": "Não encontrei carteira consolidada para esse fundo/data."
            }]))

    if not partes:
        return pd.DataFrame()

    return pd.concat(partes, ignore_index=True, sort=False)

def montar_tabela_grupo(
    grupo: str,
    fundos: list[str],
    categorias: list[str],
    data_ref: pd.Timestamp,
    fonte_dados: str,
    carteiras_processadas: dict,
    df_manual: pd.DataFrame,
) -> pd.DataFrame:
    linhas = []

    for fundo in fundos:
        resumo = obter_resumo_fundo(
            fundo=fundo,
            grupo=grupo,
            data_ref=data_ref,
            categorias=categorias,
            fonte_dados=fonte_dados,
            carteiras_processadas=carteiras_processadas,
            df_manual=df_manual,
        )

        linha = {"Fundo": fundo, "PL": resumo["PL"], "Origem": resumo.get("Origem", "")}
        for cat in categorias:
            linha[f"{cat} %"] = resumo.get(f"{cat} %", 0.0)
            linha[f"{cat} R$"] = resumo.get(f"{cat} R$", 0.0)
        linhas.append(linha)

    df = pd.DataFrame(linhas)

    total = {"Fundo": "TOTAL", "PL": float(df["PL"].sum()), "Origem": ""}
    for cat in categorias:
        valor = float(df[f"{cat} R$"].sum())
        total[f"{cat} R$"] = valor
        total[f"{cat} %"] = valor / total["PL"] if total["PL"] else 0.0

    df = pd.concat([df, pd.DataFrame([total])], ignore_index=True)
    return df


def montar_delta(df_atual: pd.DataFrame, df_base: pd.DataFrame, categorias: list[str], tipo: str) -> pd.DataFrame:
    atual = df_atual.set_index("Fundo")
    base = df_base.set_index("Fundo")
    fundos = list(df_atual["Fundo"])

    linhas = []
    for fundo in fundos:
        linha = {"Fundo": fundo}
        for cat in categorias:
            col = f"{cat} R$" if tipo == "valor" else f"{cat} %"
            va = float(atual.loc[fundo, col]) if fundo in atual.index and col in atual.columns else 0.0
            vb = float(base.loc[fundo, col]) if fundo in base.index and col in base.columns else 0.0
            linha[f"Δ {cat}"] = va - vb
        linhas.append(linha)

    return pd.DataFrame(linhas)


def montar_delta_pl(df_atual: pd.DataFrame, df_base: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula a variação do PL entre a data de análise e a data base.
    Var. PL = PL atual - PL base
    % = Var. PL / PL base
    """
    atual = df_atual.set_index("Fundo")
    base = df_base.set_index("Fundo")
    fundos = list(df_atual["Fundo"])

    linhas = []
    for fundo in fundos:
        pl_atual = float(atual.loc[fundo, "PL"]) if fundo in atual.index and "PL" in atual.columns else 0.0
        pl_base = float(base.loc[fundo, "PL"]) if fundo in base.index and "PL" in base.columns else 0.0
        var_pl = pl_atual - pl_base
        var_pct = var_pl / pl_base if pl_base else 0.0
        linhas.append({"Fundo": fundo, "Var. PL": var_pl, "%": var_pct})

    return pd.DataFrame(linhas)

def _negrito_linha_total(row):
    """
    Deixa a linha TOTAL em negrito.
    """
    fundo = str(row.get("Fundo", "")).strip().upper()

    if fundo == "TOTAL":
        return ["font-weight: 700;" for _ in row]

    return ["" for _ in row]

def _formatar_df_principal(df: pd.DataFrame, categorias: list[str]) -> pd.io.formats.style.Styler:
    fmt = {"PL": _fmt_brl}

    for cat in categorias:
        if f"{cat} %" in df.columns:
            fmt[f"{cat} %"] = _fmt_pct
        if f"{cat} R$" in df.columns:
            fmt[f"{cat} R$"] = _fmt_brl

    return (
        df.style
        .format(fmt, na_rep="")
        .apply(_negrito_linha_total, axis=1)
    )


PALETA_EXCEL = [
    "#F8696B",  # 1 - vermelho forte
    "#F97C72",  # 2
    "#FA9276",  # 3
    "#FCAA79",  # 4
    "#FDC47C",  # 5
    "#FFEB84",  # 6 - amarelo
    "#D9E87C",  # 7
    "#B2D978",  # 8
    "#8BCC74",  # 9
    "#63BE7B",  # 10 - verde forte
]


def _cor_texto_por_fundo(hex_color: str) -> str:
    """
    Define a cor do texto conforme o fundo.
    Em fundos mais fortes, usa texto escuro para manter legibilidade.
    """
    return "color:#000000;"


def _estilo_escala_excel_coluna(coluna: pd.Series, df_ref: pd.DataFrame) -> list[str]:
    """
    Aplica escala de cor por coluna, mas NÃO colore a linha TOTAL.
    A linha TOTAL fica sem background e depois recebe negrito pela função _negrito_linha_total.
    """
    valores = pd.to_numeric(coluna, errors="coerce")

    if valores.dropna().empty:
        return [""] * len(coluna)

    fundos = (
        df_ref.loc[coluna.index, "Fundo"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    mascara_total = fundos == "TOTAL"

    # Calcula a escala ignorando a linha TOTAL
    valores_sem_total = valores[~mascara_total]

    if valores_sem_total.dropna().empty:
        return [""] * len(coluna)

    max_abs = valores_sem_total.abs().max()

    estilos = []

    for idx, v in valores.items():
        if mascara_total.loc[idx]:
            estilos.append("")  # TOTAL sem cor
            continue

        if pd.isna(v):
            estilos.append("")
            continue

        if pd.isna(max_abs) or max_abs == 0:
            estilos.append(f"background-color:{PALETA_EXCEL[5]}; color:#000000;")
            continue

        pos = (v + max_abs) / (2 * max_abs)
        pos = max(0, min(1, pos))

        idx_cor = round(pos * (len(PALETA_EXCEL) - 1))
        cor = PALETA_EXCEL[idx_cor]

        estilos.append(f"background-color:{cor}; color:#000000;")

    return estilos


def _formatar_df_delta(df: pd.DataFrame, tipo: str) -> pd.io.formats.style.Styler:
    fmt = {
        c: (_fmt_brl if tipo == "valor" else _fmt_pct)
        for c in df.columns
        if c != "Fundo"
    }

    subset = [c for c in df.columns if c != "Fundo"]

    return (
        df.style
        .format(fmt, na_rep="")
        .apply(_estilo_escala_excel_coluna, subset=subset, axis=0, df_ref=df)
        .apply(_negrito_linha_total, axis=1)
    )


def _formatar_df_delta_pl(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    fmt = {"Var. PL": _fmt_brl, "%": _fmt_pct}

    # Aplica cor somente na coluna de percentual.
    # A coluna Var. PL fica sem escala de cor.
    subset = [c for c in ["%"] if c in df.columns]

    return (
        df.style
        .format(fmt, na_rep="")
        .apply(_estilo_escala_excel_coluna, subset=subset, axis=0, df_ref=df)
        .apply(_negrito_linha_total, axis=1)
    )


def altura_tabela(df: pd.DataFrame, minimo: int = 180, maximo: int = 620) -> int:
    return max(minimo, min(maximo, 38 * (len(df) + 1) + 30))


def preparar_tabela_principal_para_exibicao(
    df: pd.DataFrame,
    categorias: list[str],
    mostrar_financeiro: bool,
) -> pd.DataFrame:
    """
    Remove a coluna Origem da visualização.
    O PL aparece sempre. Por padrão, ficam ocultos apenas os valores em R$ das categorias.
    """
    cols = ["Fundo", "PL"]

    if mostrar_financeiro:
        for cat in categorias:
            cols.extend([f"{cat} %", f"{cat} R$"])
    else:
        cols.extend([f"{cat} %" for cat in categorias])

    cols = [c for c in cols if c in df.columns]
    return df[cols].copy()


def preparar_delta_para_exibicao(df: pd.DataFrame) -> pd.DataFrame:
    # Garante que só as colunas úteis apareçam na visualização.
    return df[[c for c in df.columns if c != "Origem"]].copy()


def remover_origem(df: pd.DataFrame) -> pd.DataFrame:
    return df.drop(columns=["Origem"], errors="ignore").copy()


def _datas_unicas(datas) -> list[pd.Timestamp]:
    out = []
    for d in datas:
        dt = pd.to_datetime(d, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            out.append(pd.Timestamp(dt).normalize())
    return sorted(set(out))


def _opcoes_datas(datas) -> list[str]:
    return [d.strftime("%d/%m/%Y") for d in _datas_unicas(datas)]


def _selectbox_data(label: str, fonte: str, datas_codigo, datas_manual, key_prefix: str, default_penultima: bool = False) -> pd.Timestamp:
    """
    Mostra todas as datas disponíveis, somando Código + Planilha manual,
    mas sem trocar automaticamente a fonte escolhida pelo usuário.

    Exemplo: se o Código vai até 01/2026 e a Manual vai até 04/2026,
    04/2026 aparece normalmente na lista. Se o usuário escolher fonte Código
    nessa data, o aviso é mostrado abaixo dos parâmetros.
    """
    datas_todas = _datas_unicas(list(datas_codigo) + list(datas_manual))
    opcoes = _opcoes_datas(datas_todas)

    if not opcoes:
        st.warning("Não há datas disponíveis no Código nem na Planilha manual.")
        return pd.Timestamp.today().normalize()

    if default_penultima and len(opcoes) >= 2:
        idx = len(opcoes) - 2
    else:
        idx = len(opcoes) - 1

    escolha = st.selectbox(
        label,
        options=opcoes,
        index=idx,
        key=f"{key_prefix}_{_norm_txt(fonte)}",
    )
    return pd.to_datetime(escolha, dayfirst=True).normalize()


def _fonte_tem_data(fonte: str, data_ref: pd.Timestamp, datas_codigo, datas_manual) -> bool:
    data_ref = pd.Timestamp(data_ref).normalize()
    datas_codigo_set = set(_datas_unicas(datas_codigo))
    datas_manual_set = set(_datas_unicas(datas_manual))

    if fonte == FONTE_CODIGO:
        return data_ref in datas_codigo_set
    if fonte == FONTE_MANUAL:
        return data_ref in datas_manual_set
    return False

def _fmt_mes_ano(data_ref: pd.Timestamp) -> str:
    meses = {
        1: "jan", 2: "fev", 3: "mar", 4: "abr",
        5: "mai", 6: "jun", 7: "jul", 8: "ago",
        9: "set", 10: "out", 11: "nov", 12: "dez",
    }
    data_ref = pd.Timestamp(data_ref).normalize()
    return f"{meses[data_ref.month]}/{str(data_ref.year)[-2:]}"


def _avisar_sem_dados_fontes_agregado(
    selecoes: list[tuple[str, str, pd.Timestamp]],
    datas_codigo,
    datas_manual,
) -> None:
    """
    Junta os avisos de datas indisponíveis em uma única mensagem.
    """
    datas_codigo_set = set(_datas_unicas(datas_codigo))
    datas_manual_set = set(_datas_unicas(datas_manual))

    sem_codigo_mas_tem_manual = []
    sem_manual_mas_tem_codigo = []

    for nome_data, fonte, data_ref in selecoes:
        data_ref = pd.Timestamp(data_ref).normalize()

        if fonte == FONTE_CODIGO:
            if data_ref not in datas_codigo_set and data_ref in datas_manual_set:
                sem_codigo_mas_tem_manual.append(data_ref)

        elif fonte == FONTE_MANUAL:
            if data_ref not in datas_manual_set and data_ref in datas_codigo_set:
                sem_manual_mas_tem_codigo.append(data_ref)

    mensagens = []

    if sem_codigo_mas_tem_manual:
        datas_fmt = "; ".join(
            _fmt_mes_ano(d)
            for d in sorted(set(sem_codigo_mas_tem_manual))
        )

        palavra = (
            "Essa data existe"
            if len(set(sem_codigo_mas_tem_manual)) == 1
            else "Essas datas existem"
        )

        mensagens.append(
            f"Não há dados no Código para {datas_fmt}. "
            f"{palavra} apenas na Planilha manual."
        )

    if sem_manual_mas_tem_codigo:
        datas_fmt = "; ".join(
            _fmt_mes_ano(d)
            for d in sorted(set(sem_manual_mas_tem_codigo))
        )

        palavra = (
            "Essa data existe"
            if len(set(sem_manual_mas_tem_codigo)) == 1
            else "Essas datas existem"
        )

        mensagens.append(
            f"Não há dados na Planilha manual para {datas_fmt}. "
            f"{palavra} apenas no Código."
        )

    if mensagens:
        st.warning(" ".join(mensagens))

# =========================================================
# INTERFACE
# =========================================================
carteiras_processadas, datas_carteira = carregar_carteiras_processadas(ARQUIVO_CARTEIRA)
manual_path = _resolver_planilha_manual()
if manual_path:
    df_manual = carregar_planilha_manual(str(manual_path))
else:
    df_manual = pd.DataFrame()

st.markdown(
    """
    <style>
    .sectionbar {
        margin-bottom: 2px !important;
    }

    div[data-testid="stRadio"] {
        margin-top: -6px !important;
    }

    div[data-testid="stCheckbox"] {
        margin-top: -6px !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

render_sectionbar("Comparar", "sectionbar-fundos")

# Datas disponíveis em cada fonte. A escolha de data é sempre por lista fechada,
# nunca por calendário livre.
datas_codigo = _datas_unicas(datas_carteira)
datas_manual = _datas_unicas(df_manual["Data"].unique()) if not df_manual.empty and "Data" in df_manual.columns else []

# Ajuste visual da linha de opções
st.markdown(
    """
    <style>
    div[data-testid="stRadio"] [role="radiogroup"] {
        width: 100%;
        display: flex;
        justify-content: space-between;
    }

    div[data-testid="stCheckbox"] {
        padding-top: 0px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Linha 1: opções principais distribuídas na linha inteira
l1, l2 = st.columns([2, 1], vertical_alignment="center")

with l1:
    modo_fundos = st.radio(
        "Modo de comparação",
        ["Todos os fundos", "Escolher fundos específicos"],
        horizontal=True,
        label_visibility="collapsed",
        key="modo_fundos_comparacao",
    )

with l2:
    mostrar_financeiro = st.checkbox(
        "Mostrar valores em R$ das categorias",
        value=False,
        help="O PL aparece sempre. Marque para exibir também os valores financeiros por categoria na tabela principal.",
        key="mostrar_financeiro_categorias",
    )

# Linha 2: fontes e datas
c1, c2, c3, c4 = st.columns([1.7, 1.4, 1.7, 1.4])

with c1:
    fonte_analise = st.selectbox(
        "Fonte da data para análise:",
        [FONTE_CODIGO, FONTE_MANUAL],
        index=0,
    )

with c2:
    data_analise_ts = _selectbox_data(
        "Data para análise:",
        fonte_analise,
        datas_codigo,
        datas_manual,
        key_prefix="data_analise",
        default_penultima=False,
    )

with c3:
    fonte_comparar = st.selectbox(
        "Fonte da data para comparar:",
        [FONTE_CODIGO, FONTE_MANUAL],
        index=0,
    )

with c4:
    data_comparar_ts = _selectbox_data(
        "Data para comparar:",
        fonte_comparar,
        datas_codigo,
        datas_manual,
        key_prefix="data_comparar",
        default_penultima=True,
    )

if manual_path:
    st.markdown(
        f"""
        <div style='font-size:10px; color:#8a8a8a; margin-top:2px; margin-bottom:4px;'>
            Planilha manual encontrada: {manual_path}
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    st.warning(
        "A planilha manual 'Acompanhamento Fundos Terceiros' não foi encontrada. "
        "As datas da fonte manual não aparecerão na lista."
    )

# Avisos de disponibilidade: as datas aparecem todas na lista,
# mas a fonte escolhida continua sendo respeitada.
# Avisos de disponibilidade agrupados em uma única mensagem
_avisar_sem_dados_fontes_agregado(
    selecoes=[
        ("Data para análise", fonte_analise, data_analise_ts),
        ("Data para comparar", fonte_comparar, data_comparar_ts),
    ],
    datas_codigo=datas_codigo,
    datas_manual=datas_manual,
)

# seleção de fundos
fundos_por_grupo = {}
if modo_fundos == "Todos os fundos":
    for grupo, cfg in GRUPOS.items():
        fundos_por_grupo[grupo] = cfg["fundos"]
else:
    render_sectionbar("Escolha os fundos", "sectionbar-ativos")
    s1, s2, s3 = st.columns(3)
    with s1:
        fundos_por_grupo["Fundo D+1"] = st.multiselect("Fundos D+1", FUNDOS_D1, default=FUNDOS_D1)
    with s2:
        fundos_por_grupo["Fundo D+30"] = st.multiselect("Fundos D+30", FUNDOS_D30, default=FUNDOS_D30)
    with s3:
        fundos_por_grupo["Fundo D+60"] = st.multiselect("Fundos D+60", FUNDOS_D60, default=FUNDOS_D60)

# monta resultados
resultados_atual = {}
resultados_base = {}
deltas_pl = {}
deltas_valor = {}
deltas_pct = {}
ativos_classificados_atual = {}
ativos_classificados_base = {}

for grupo, cfg in GRUPOS.items():
    fundos = fundos_por_grupo.get(grupo, [])
    categorias = cfg["categorias"]

    resultados_atual[grupo] = montar_tabela_grupo(
        grupo, fundos, categorias, data_analise_ts, fonte_analise,
        carteiras_processadas, df_manual,
    )
    resultados_base[grupo] = montar_tabela_grupo(
        grupo, fundos, categorias, data_comparar_ts, fonte_comparar,
        carteiras_processadas, df_manual,
    )
    deltas_valor[grupo] = montar_delta(resultados_atual[grupo], resultados_base[grupo], categorias, "valor")
    deltas_pct[grupo] = montar_delta(resultados_atual[grupo], resultados_base[grupo], categorias, "pct")
    deltas_pl[grupo] = montar_delta_pl(resultados_atual[grupo], resultados_base[grupo])
    
    ativos_classificados_atual[grupo] = montar_ativos_classificados_grupo(
        grupo=grupo,
        fundos=fundos,
        categorias=categorias,
        data_ref=data_analise_ts,
        fonte_dados=fonte_analise,
        carteiras_processadas=carteiras_processadas,
    )

    ativos_classificados_base[grupo] = montar_ativos_classificados_grupo(
        grupo=grupo,
        fundos=fundos,
        categorias=categorias,
        data_ref=data_comparar_ts,
        fonte_dados=fonte_comparar,
        carteiras_processadas=carteiras_processadas,
    )   
# exibição
st.markdown(
    f"<div style='font-weight:700; margin: 8px 0 12px 0;'>"
    f"Data para análise: {data_analise_ts.strftime('%d/%m/%Y')} ({fonte_analise}) &nbsp;&nbsp; | &nbsp;&nbsp; "
    f"Data para comparar: {data_comparar_ts.strftime('%d/%m/%Y')} ({fonte_comparar})"
    f"</div>",
    unsafe_allow_html=True,
)

for grupo, cfg in GRUPOS.items():
    categorias = cfg["categorias"]
    render_sectionbar(grupo, cfg["classe_css"])

    df_atual = resultados_atual[grupo]
    df_atual_show = preparar_tabela_principal_para_exibicao(
        df_atual,
        categorias,
        mostrar_financeiro=mostrar_financeiro,
    )

    st.dataframe(
        _formatar_df_principal(df_atual_show, categorias),
        use_container_width=True,
        height=altura_tabela(df_atual_show),
    )

    st.markdown("<div style='font-size:18px; font-weight:700; margin-top:10px;'>Variação entre as datas por categoria</div>", unsafe_allow_html=True)

    delta_valor_show = preparar_delta_para_exibicao(deltas_valor[grupo])
    delta_pct_show = preparar_delta_para_exibicao(deltas_pct[grupo])

    v1, v2 = st.columns(2)
    with v1:
        st.markdown("<b>Variação em R$</b>", unsafe_allow_html=True)
        st.dataframe(
            _formatar_df_delta(delta_valor_show, "valor"),
            use_container_width=True,
            height=altura_tabela(delta_valor_show, minimo=160, maximo=420),
        )
    with v2:
        st.markdown("<b>Variação em %PL</b>", unsafe_allow_html=True)
        st.dataframe(
            _formatar_df_delta(delta_pct_show, "pct"),
            use_container_width=True,
            height=altura_tabela(delta_pct_show, minimo=160, maximo=420),
        )

    st.markdown("<div style='font-size:18px; font-weight:700; margin-top:10px;'>Variação do PL</div>", unsafe_allow_html=True)
    delta_pl_show = preparar_delta_para_exibicao(deltas_pl[grupo])
    st.dataframe(
        _formatar_df_delta_pl(delta_pl_show),
        use_container_width=True,
        height=altura_tabela(delta_pl_show, minimo=160, maximo=420),
    )

# exportação
render_sectionbar("Exportar", "sectionbar-exportar")
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    for grupo in GRUPOS:
        sheet_prefix = grupo.replace("Fundo ", "").replace("+", "p").replace(" ", "")
        remover_origem(resultados_atual[grupo]).to_excel(writer, index=False, sheet_name=f"{sheet_prefix}_Atual"[:31])
        remover_origem(resultados_base[grupo]).to_excel(writer, index=False, sheet_name=f"{sheet_prefix}_Base"[:31])
        remover_origem(deltas_pl[grupo]).to_excel(writer, index=False, sheet_name=f"{sheet_prefix}_Delta_PL"[:31])
        remover_origem(deltas_valor[grupo]).to_excel(writer, index=False, sheet_name=f"{sheet_prefix}_Delta_RS"[:31])
        remover_origem(deltas_pct[grupo]).to_excel(writer, index=False, sheet_name=f"{sheet_prefix}_Delta_pct"[:31])
        ativos_classificados_atual[grupo].to_excel(writer, index=False, sheet_name=f"{sheet_prefix}_Ativos_Analise"[:31])
        ativos_classificados_base[grupo].to_excel(writer,  index=False, sheet_name=f"{sheet_prefix}_Ativos_Base"[:31])
buffer.seek(0)
st.download_button(
    label="Baixar Excel da comparação por categorias",
    data=buffer,
    file_name=f"comparacao_categorias_{data_analise_ts.strftime('%Y-%m-%d')}_vs_{data_comparar_ts.strftime('%Y-%m-%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
