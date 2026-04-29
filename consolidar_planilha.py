from __future__ import annotations

from pathlib import Path
import calendar
import re
import pandas as pd

from tratar_planilha_consolidada import executar_tratamento

# =========================================================
# CONFIGURAÇÕES
# =========================================================
BASE_DIR = Path(__file__).resolve().parent
PASTA_BASE = BASE_DIR / "Base Carteira Fundos"
ARQUIVO_SAIDA = BASE_DIR / "carteira_fundos_consolidada.xlsx"

FUNDOS = {
    #GERAES": "09.720.734/0001-10",
    #"ARX Denali": "41.888.219/0001-56",
    #"Capitânia Top": "13.615.411/0001-33",
    #"Riza Lotus": "36.498.670/0001-27",
    #"Western Asset Total": "28.320.756/0001-37",
    #"Daycoval Classic": "10.783.480/0001-68",
    #"Iridium Apollo": "26.978.438/0001-32",
    #"Porto Seguro FIRF": "18.719.154/0001-01",
    #"Nu Reserva Imediata": "42.699.466/0001-77",
    #"Inter Conservador": "20.879.578/0001-77",
    #"Absolute Atenas": "48.096.589/0001-08",
    "GERAES 30": "29.044.189/0001-04",
    #"Daycoval Classic 30": "29.250.121/0001-73",
    #"Riza Lotus Plus": "43.917.493/0001-31",
    #"ARX Vinson Advisory": "41.888.492/0001-80",
    #"Sparta Max Advisory": "32.757.330/0001-12",
    #"Polo Crédito Corporativo": "31.455.879/0001-90",
    #"Iridium Titan Advisory": "32.756.812/0001-58",
    #"Porto Seguro Ipê": "35.378.376/0001-19",
    #"Sparta Top Advisory": "32.846.811/0001-02",
    #"Absolute Creta": "48.122.126/0001-65",
    #"Horizonte": "44.025.131/0001-07",
    #"JGP Select": "21.946.695/0001-79",
    #"ARX Everest Advisory": "35.789.436/0001-96",
    #"Polo Total": "23.601.467/0001-92",
    #"Absolute Olimpia": "48.986.106/0001-32",
    
}

# Se deixar vazio [], processa TODOS os fundos do dicionário FUNDOS
FUNDOS_PROCESSAR = [
    #"GERAES",
    #"ARX Denali",
    #"Capitânia Top",
    #"Riza Lotus",
    #"Western Asset Total",
    #"Daycoval Classic",
    #"Iridium Apollo",
    #"Porto Seguro FIRF",
    #"Nu Reserva Imediata",
    #"Inter Conservador",
    #"Absolute Atenas",
    "GERAES 30",
    #"Daycoval Classic 30",
    #"Riza Lotus Plus",
    #"ARX Vinson Advisory",
    #"Sparta Max Advisory",
    #"Polo Crédito Corporativo",
    #"Iridium Titan Advisory",
    #"Porto Seguro Ipê",
    #"Sparta Top Advisory",
    #"Absolute Creta",
    #"Horizonte",
    #"JGP Select",
    #"ARX Everest Advisory",
    #"Polo Total",
    #"Absolute Olimpia",

]

COLUNAS_FIXAS = ["Codigo", "ISIN", "Descrição", "Emissor", "Classe", "Tipo de Investimento", "Vencimento"]
COLUNAS_CHAVE_AGRUPAMENTO = COLUNAS_FIXAS + ["DataRef"]

# Reprocessa automaticamente os últimos X meses encontrados na pasta
REPROCESSAR_ULTIMOS_MESES = 0

# Se quiser forçar meses específicos, coloque aqui
REPROCESSAR_COMPETENCIAS = ["202512"]
# REPROCESSAR_COMPETENCIAS: list[str] = []

# =========================================================
# CONFIGURAÇÕES DO LOOK-THROUGH
# =========================================================
ATIVAR_LOOKTHROUGH_FUNDOS = True
PROFUNDIDADE_MAXIMA_LOOKTHROUGH = 10

# False = substitui a linha da cota do fundo pelos ativos internos do fundo investido
# True  = mantém a linha da cota e também adiciona os ativos internos do fundo investido
MANTER_LINHA_COTA_FUNDO = False

# Para distribuir o valor aplicado no fundo investido, usa apenas posições positivas
# do fundo filho. Isso evita passivos como "valores a pagar" na abertura econômica.
LOOKTHROUGH_SOMENTE_POSITIVOS = False

# Se não conseguir abrir o fundo investido, mantém a linha original da cota
MANTER_COTA_SE_NAO_ABRIR = True


# =========================================================
# MAPA COMPACTO DAS COLUNAS POR COMPETÊNCIA E PLANILHA
# =========================================================
ORDEM_CAMPOS = (
    "CNPJ",
    "Código",
    "ISIN",
    "Descrição",
    "Emissor",
    "Classe",
    "Tipo de Investimento",
    "Vencimento",
    "Valor",
)

PADRAO_COLUNAS_ORDEM = tuple("CNPJ_FUNDO" for _ in ORDEM_CAMPOS)


COLUNAS_INTERNAS = [
    "EhFundoInvestido",
    "CNPJ_Fundo_Investido",
    "Nome_Fundo_Investido",
    "OrigemLookThrough",
    "FundoRaiz",
    "NivelLookThrough",
]


def colunas_base_internas() -> list[str]:
    return COLUNAS_FIXAS + ["DataRef", "VL_MERC_POS_FINAL_NUM"] + COLUNAS_INTERNAS



def montar_mapa_colunas_padrao() -> dict[str, dict[str, tuple[str, ...]]]:
    mapa: dict[str, dict[str, tuple[str, ...]]] = {}

    if PASTA_BASE.exists():
        competencias = sorted(
            p.name for p in PASTA_BASE.iterdir()
            if p.is_dir() and p.name.isdigit() and len(p.name) == 6
        )
    else:
        competencias = []

    for comp in competencias:
        mapa[comp] = {
            str(i): PADRAO_COLUNAS_ORDEM
            for i in range(1, 9)
        }

    return mapa


MAPA_COLUNAS_POR_COMPETENCIA = montar_mapa_colunas_padrao()

CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602 = {
    "1": ("CNPJ_FUNDO_CLASSE", "TP_TITPUB", "CD_ISIN", "-", " -", "TP_ATIVO", "TP_APLIC", "DT_VENC", "VL_MERC_POS_FINAL"),
    "2": ("CNPJ_FUNDO_CLASSE", "-", "CNPJ_FUNDO_CLASSE_COTA", "NM_FUNDO_CLASSE_SUBCLASSE_COTA", "-", "TP_ATIVO", "TP_APLIC", "-", "VL_MERC_POS_FINAL"),
    "3": ("CNPJ_FUNDO_CLASSE", "DS_SWAP", "DS_SWAP", "DS_SWAP", "TP_APLIC", "TP_ATIVO", "TP_APLIC", "-", "VL_MERC_POS_FINAL"),
    "4": ("CNPJ_FUNDO_CLASSE", "CD_ATIVO", "CD_ISIN", "DS_ATIVO", "DS_ATIVO", "TP_ATIVO", "TP_APLIC", "DT_FIM_VIGENCIA", "VL_MERC_POS_FINAL"),
    "5": ("CNPJ_FUNDO_CLASSE", "-", "CNPJ_EMISSOR", "EMISSOR", "EMISSOR", "TP_ATIVO", "TP_APLIC", "DT_VENC", "VL_MERC_POS_FINAL"),
    "6": ("CNPJ_FUNDO_CLASSE", "EMISSOR", "CPF_CNPJ_EMISSOR", "EMISSOR", "EMISSOR", "TP_ATIVO", "TP_APLIC", "DT_VENC", "VL_MERC_POS_FINAL"),
    "7": ("CNPJ_FUNDO_CLASSE", "-", "CD_ATIVO_BV_MERC", "CD_BV_MERC", "EMISSOR", "TP_ATIVO", "TP_APLIC", "DT_VENC", "VL_MERC_POS_FINAL"),
    "8": ("CNPJ_FUNDO_CLASSE", "DS_ATIVO", "CPF_CNPJ_EMISSOR", "DS_ATIVO", "EMISSOR", "TP_ATIVO", "TP_APLIC", "-", "VL_MERC_POS_FINAL"),
}

CONFIG_202301_202302_202303_202304_202305_202306_202307_202308_202309_202310_202311 = {
    "1": ("CNPJ_FUNDO", "TP_TITPUB", "CD_ISIN", "-", " -", "TP_ATIVO", "TP_APLIC", "DT_VENC", "VL_MERC_POS_FINAL"),
    "2": ("CNPJ_FUNDO", "-", "CNPJ_FUNDO_COTA", "NM_FUNDO_COTA", "-", "TP_ATIVO", "TP_APLIC", "-", "VL_MERC_POS_FINAL"),
    "3": ("CNPJ_FUNDO", "DS_SWAP", "DS_SWAP", "DS_SWAP", "TP_APLIC", "TP_ATIVO", "TP_APLIC", "-", "VL_MERC_POS_FINAL"),
    "4": ("CNPJ_FUNDO", "CD_ATIVO", "CD_ISIN", "DS_ATIVO", "DS_ATIVO", "TP_ATIVO", "TP_APLIC", "DT_FIM_VIGENCIA", "VL_MERC_POS_FINAL"),
    "5": ("CNPJ_FUNDO", "-", "CNPJ_EMISSOR", "EMISSOR", "EMISSOR", "TP_ATIVO", "TP_APLIC", "DT_VENC", "VL_MERC_POS_FINAL"),
    "6": ("CNPJ_FUNDO", "EMISSOR", "CPF_CNPJ_EMISSOR", "EMISSOR", "EMISSOR", "TP_ATIVO", "TP_APLIC", "DT_VENC", "VL_MERC_POS_FINAL"),
    "7": ("CNPJ_FUNDO", "-", "CD_ATIVO_BV_MERC", "CD_BV_MERC", "EMISSOR", "TP_ATIVO", "TP_APLIC", "DT_VENC", "VL_MERC_POS_FINAL"),
    "8": ("CNPJ_FUNDO", "DS_ATIVO", "CPF_CNPJ_EMISSOR", "DS_ATIVO", "EMISSOR", "TP_ATIVO", "TP_APLIC", "-", "VL_MERC_POS_FINAL"),
}

MAPA_COLUNAS_POR_COMPETENCIA["202301"] = CONFIG_202301_202302_202303_202304_202305_202306_202307_202308_202309_202310_202311.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202302"] = CONFIG_202301_202302_202303_202304_202305_202306_202307_202308_202309_202310_202311.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202303"] = CONFIG_202301_202302_202303_202304_202305_202306_202307_202308_202309_202310_202311.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202304"] = CONFIG_202301_202302_202303_202304_202305_202306_202307_202308_202309_202310_202311.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202305"] = CONFIG_202301_202302_202303_202304_202305_202306_202307_202308_202309_202310_202311.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202306"] = CONFIG_202301_202302_202303_202304_202305_202306_202307_202308_202309_202310_202311.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202307"] = CONFIG_202301_202302_202303_202304_202305_202306_202307_202308_202309_202310_202311.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202308"] = CONFIG_202301_202302_202303_202304_202305_202306_202307_202308_202309_202310_202311.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202309"] = CONFIG_202301_202302_202303_202304_202305_202306_202307_202308_202309_202310_202311.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202310"] = CONFIG_202301_202302_202303_202304_202305_202306_202307_202308_202309_202310_202311.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202311"] = CONFIG_202301_202302_202303_202304_202305_202306_202307_202308_202309_202310_202311.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202312"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202401"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202402"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202403"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202404"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202405"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202406"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202407"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202408"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202409"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202410"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202411"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202412"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202501"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202502"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202503"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202504"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202504"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202505"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202506"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202507"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202508"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202509"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202510"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202511"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202512"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202601"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()
MAPA_COLUNAS_POR_COMPETENCIA["202602"] = CONFIG_202312_202401_202402_202403_202404_202405_202406_202407_202408_202409_202410_202411_202412_202501_202502_202503_202504_202506_202507_202508_202509_202510_202511_202512_202601_202602.copy()

# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================
def converter_config_bloco_para_dict(config_bloco: tuple[str, ...] | list[str] | dict[str, str] | None) -> dict[str, str]:
    if config_bloco is None:
        return dict(zip(ORDEM_CAMPOS, PADRAO_COLUNAS_ORDEM))

    if isinstance(config_bloco, dict):
        return {campo: str(config_bloco.get(campo, "CNPJ_FUNDO")) for campo in ORDEM_CAMPOS}

    valores = list(config_bloco)

    if len(valores) < len(ORDEM_CAMPOS):
        valores.extend(["CNPJ_FUNDO"] * (len(ORDEM_CAMPOS) - len(valores)))
    elif len(valores) > len(ORDEM_CAMPOS):
        valores = valores[:len(ORDEM_CAMPOS)]

    return {campo: str(valor) for campo, valor in zip(ORDEM_CAMPOS, valores)}



def imprimir_mapa_colunas() -> None:
    print(f"-> ordem {ORDEM_CAMPOS}")

    for comp in sorted(MAPA_COLUNAS_POR_COMPETENCIA):
        print(f"\n{comp}")
        for i in range(1, 9):
            config_bloco = MAPA_COLUNAS_POR_COMPETENCIA.get(comp, {}).get(str(i), PADRAO_COLUNAS_ORDEM)
            config_dict = converter_config_bloco_para_dict(config_bloco)
            valores = [f'"{config_dict[campo]}"' for campo in ORDEM_CAMPOS]
            print(f'Planilha {i}) ' + '; '.join(valores))



def normalizar_nome_aba(nome: str) -> str:
    nome = re.sub(r'[:\\/*?\[\]]', '_', str(nome).strip())
    nome = nome[:31].strip()
    return nome or "ABA"



def normalizar_cnpj(valor) -> str:
    if pd.isna(valor):
        return ""
    return re.sub(r"\D", "", str(valor))



def obter_fundos_para_processar() -> dict[str, str]:
    fundos_normalizados = {
        normalizar_nome_aba(nome): normalizar_cnpj(cnpj)
        for nome, cnpj in FUNDOS.items()
        if str(cnpj).strip() != ""
    }

    if FUNDOS_PROCESSAR:
        nomes_desejados = {normalizar_nome_aba(nome) for nome in FUNDOS_PROCESSAR}
        return {
            nome: cnpj
            for nome, cnpj in fundos_normalizados.items()
            if nome in nomes_desejados
        }

    return fundos_normalizados



def ler_csv_seguro(caminho: Path) -> pd.DataFrame | None:
    tentativas = [
        {"sep": ";", "encoding": "latin1", "dtype": str, "low_memory": False},
        {"sep": ";", "encoding": "utf-8", "dtype": str, "low_memory": False},
        {"sep": ",", "encoding": "latin1", "dtype": str, "low_memory": False},
        {"sep": ",", "encoding": "utf-8", "dtype": str, "low_memory": False},
    ]

    for params in tentativas:
        try:
            df = pd.read_csv(caminho, **params)
            if len(df.columns) > 1:
                return df
        except Exception:
            pass

    return None

def identificar_primeira_coluna_existente(df: pd.DataFrame, candidatos: list[str]) -> str | None:
    cols_norm = {str(col).strip().upper(): col for col in df.columns}

    for nome in candidatos:
        chave = str(nome).strip().upper()
        if chave in cols_norm:
            return cols_norm[chave]

    return None

COLUNAS_CNPJ_PL = [
    "CNPJ_FUNDO_CLASSE",
    "CNPJ_FUNDO",
]

COLUNAS_PATRIMONIO_LIQUIDO = [
    "VL_PATRIM_LIQ",
    "VL_PATRIM_LIQUIDO",
    "VL_PL",
    "PATRIMONIO_LIQUIDO",
    "PATRIMÔNIO LÍQUIDO",
    "PATRIMÔNIO_LÍQUIDO",
    "PL",
]

def identificar_bloco_arquivo(csv_path: Path) -> str | None:
    nome = csv_path.stem.upper()
    match = re.search(r"BLC[_\-]?(\d+)", nome)
    if match:
        return match.group(1)
    return None



def obter_coluna_configurada(pasta_mes: str, csv_path: Path, campo_logico: str) -> str | None:
    config_mes = MAPA_COLUNAS_POR_COMPETENCIA.get(pasta_mes)
    if config_mes is None:
        return None

    bloco = identificar_bloco_arquivo(csv_path)
    if bloco is None:
        return None

    config_bloco = config_mes.get(bloco)
    if config_bloco is None:
        return None

    config_dict = converter_config_bloco_para_dict(config_bloco)
    return config_dict.get(campo_logico)



def identificar_coluna_configurada(df: pd.DataFrame, pasta_mes: str, csv_path: Path, campo_logico: str) -> str | None:
    col_configurada = obter_coluna_configurada(pasta_mes, csv_path, campo_logico)

    if col_configurada is None or str(col_configurada).strip() in {"", "-"}:
        return None

    if col_configurada not in df.columns:
        print(
            f"[AVISO] A coluna configurada '{col_configurada}' para o campo '{campo_logico}' "
            f"não existe em {csv_path.name} na competência {pasta_mes}."
        )
        return None

    return col_configurada



def obter_serie_configurada(
    df: pd.DataFrame,
    pasta_mes: str,
    csv_path: Path,
    campo_logico: str,
    valor_padrao: str = "-",
) -> pd.Series:
    col = identificar_coluna_configurada(df, pasta_mes, csv_path, campo_logico)
    if col is None:
        return pd.Series([valor_padrao] * len(df), index=df.index, dtype="object")

    serie = df[col].fillna(valor_padrao).astype(str).str.strip()
    serie = serie.replace({"": valor_padrao, "nan": valor_padrao, "None": valor_padrao})
    return serie



def inferir_data_da_pasta(pasta_mes: str) -> str:
    ano = int(pasta_mes[:4])
    mes = int(pasta_mes[4:6])
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    return f"{ano:04d}-{mes:02d}-{ultimo_dia:02d}"



def normalizar_data_texto(valor) -> str:
    if pd.isna(valor):
        return "-"

    s = str(valor).strip()
    if s in {"", "nan", "None", "-"}:
        return "-"

    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return s

    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", s):
        dt = pd.to_datetime(s, format="%d/%m/%Y", errors="coerce")
        if pd.isna(dt):
            return s
        return dt.strftime("%Y-%m-%d")

    dt = pd.to_datetime(s, errors="coerce")
    if pd.isna(dt):
        return s

    return dt.strftime("%Y-%m-%d")



def converter_valor(v) -> float:
    if pd.isna(v):
        return 0.0

    s = str(v).strip()
    if s == "":
        return 0.0

    s = s.replace(" ", "")

    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except Exception:
            return 0.0

    if "," in s and "." not in s:
        s = s.replace(",", ".")
        try:
            return float(s)
        except Exception:
            return 0.0

    if "." in s and "," not in s:
        try:
            return float(s)
        except Exception:
            return 0.0

    try:
        return float(s)
    except Exception:
        return 0.0



def escolher_codigo(df: pd.DataFrame, pasta_mes: str, csv_path: Path) -> pd.Series:
    return obter_serie_configurada(df=df, pasta_mes=pasta_mes, csv_path=csv_path, campo_logico="Código", valor_padrao="-")



def escolher_isin(df: pd.DataFrame, pasta_mes: str, csv_path: Path) -> pd.Series:
    return obter_serie_configurada(df=df, pasta_mes=pasta_mes, csv_path=csv_path, campo_logico="ISIN", valor_padrao="-")



def escolher_descricao(df: pd.DataFrame, pasta_mes: str, csv_path: Path) -> pd.Series:
    return obter_serie_configurada(df=df, pasta_mes=pasta_mes, csv_path=csv_path, campo_logico="Descrição", valor_padrao="-")



def escolher_emissor(df: pd.DataFrame, pasta_mes: str, csv_path: Path) -> pd.Series:
    return obter_serie_configurada(df=df, pasta_mes=pasta_mes, csv_path=csv_path, campo_logico="Emissor", valor_padrao="-")



def escolher_classe(df: pd.DataFrame, pasta_mes: str, csv_path: Path) -> pd.Series:
    return obter_serie_configurada(df=df, pasta_mes=pasta_mes, csv_path=csv_path, campo_logico="Classe", valor_padrao="")



def escolher_tipo_investimento(df: pd.DataFrame, pasta_mes: str, csv_path: Path) -> pd.Series:
    return obter_serie_configurada(df=df, pasta_mes=pasta_mes, csv_path=csv_path, campo_logico="Tipo de Investimento", valor_padrao="-")



def escolher_vencimento(df: pd.DataFrame, pasta_mes: str, csv_path: Path) -> pd.Series:
    serie = obter_serie_configurada(df=df, pasta_mes=pasta_mes, csv_path=csv_path, campo_logico="Vencimento", valor_padrao="-")
    return serie.apply(normalizar_data_texto)



def obter_competencias_disponiveis() -> list[str]:
    if not PASTA_BASE.exists():
        return []

    competencias = []
    for p in PASTA_BASE.iterdir():
        if p.is_dir() and p.name.isdigit() and len(p.name) == 6:
            competencias.append(p.name)

    return sorted(competencias)


TOLERANCIA_DIFERENCA_PL = 0.01



def ler_pl_competencia(pasta_mes: str, cnpj_fundo: str) -> float | None:
    pasta = PASTA_BASE / pasta_mes
    arquivos_pl = sorted(pasta.glob("cda_fi_PL_*.csv"))

    if not arquivos_pl:
        print(f"[AVISO] Nenhum arquivo cda_fi_PL_ encontrado em {pasta_mes}.")
        return None

    cnpj_fundo_norm = normalizar_cnpj(cnpj_fundo)

    for arq in arquivos_pl:
        df = ler_csv_seguro(arq)
        if df is None or df.empty:
            continue

        col_cnpj = identificar_primeira_coluna_existente(df, COLUNAS_CNPJ_PL)
        col_pl = identificar_primeira_coluna_existente(df, COLUNAS_PATRIMONIO_LIQUIDO)

        if col_cnpj is None or col_pl is None:
            print(
                f"[AVISO] Estrutura inesperada no arquivo {arq.name}. "
                f"CNPJ encontrado: {col_cnpj} | PL encontrado: {col_pl}"
            )
            continue

        df[col_cnpj] = df[col_cnpj].apply(normalizar_cnpj)

        filtrado = df[df[col_cnpj] == cnpj_fundo_norm].copy()

        if filtrado.empty:
            continue

        pl = filtrado[col_pl].apply(converter_valor).sum()
        return float(pl)

    print(f"[AVISO] Fundo {cnpj_fundo} não encontrado no cda_fi_PL_ de {pasta_mes}.")
    return None


def montar_resumo_conferencia(
    competencias_processar: list[str],
    somas_por_data: dict[str, float],
    cnpj_fundo: str,
) -> pd.DataFrame:
    linhas = []

    for comp in competencias_processar:
        data_ref = inferir_data_da_pasta(comp)
        soma_coluna = float(somas_por_data.get(data_ref, 0.0))
        pl_arquivo = ler_pl_competencia(comp, cnpj_fundo)

        if pl_arquivo is None:
            diferenca = None
            status = "PL_NAO_ENCONTRADO"
        else:
            diferenca = soma_coluna - pl_arquivo
            status = "IGUAL" if abs(diferenca) <= TOLERANCIA_DIFERENCA_PL else "DIFERENCA"

        linhas.append({
            "DataRef": data_ref,
            "Soma_Coluna": soma_coluna,
            "PL_Arquivo": pl_arquivo,
            "Diferenca": diferenca,
            "Status": status,
        })

    return pd.DataFrame(linhas)



def obter_datas_existentes_da_aba(df_aba: pd.DataFrame | None) -> set[str]:
    if df_aba is None or df_aba.empty:
        return set()

    datas = set()
    for col in df_aba.columns:
        if col not in COLUNAS_FIXAS:
            datas.add(str(col))
    return datas



def carregar_todas_abas_existentes() -> dict[str, pd.DataFrame]:
    if not ARQUIVO_SAIDA.exists():
        return {}

    try:
        return pd.read_excel(ARQUIVO_SAIDA, sheet_name=None, dtype=object)
    except Exception:
        return {}



def normalizar_planilha_existente(df: pd.DataFrame) -> pd.DataFrame:
    for col in COLUNAS_FIXAS:
        if col not in df.columns:
            df[col] = ""

    colunas_datas = [c for c in df.columns if c not in COLUNAS_FIXAS]
    for c in colunas_datas:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    return df


# =========================================================
# LEITURA DA CARTEIRA DIRETA
# =========================================================
def linha_representa_fundo_investido(df: pd.DataFrame, pasta_mes: str, csv_path: Path) -> tuple[pd.Series, pd.Series, pd.Series]:
    bloco = identificar_bloco_arquivo(csv_path)

    eh_fundo = pd.Series(False, index=df.index)
    cnpj_fundo_investido = pd.Series([""] * len(df), index=df.index, dtype="object")
    nome_fundo_investido = pd.Series([""] * len(df), index=df.index, dtype="object")

    if bloco != "2":
        return eh_fundo, cnpj_fundo_investido, nome_fundo_investido

    col_cnpj_fundo_investido = identificar_coluna_configurada(df, pasta_mes, csv_path, "ISIN")
    if col_cnpj_fundo_investido is None:
        return eh_fundo, cnpj_fundo_investido, nome_fundo_investido

    cnpj_fundo_investido = df[col_cnpj_fundo_investido].apply(normalizar_cnpj)
    eh_fundo = cnpj_fundo_investido.ne("")

    col_nome_fundo_investido = identificar_coluna_configurada(df, pasta_mes, csv_path, "Descrição")
    if col_nome_fundo_investido is not None:
        nome_fundo_investido = (
            df[col_nome_fundo_investido]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    return eh_fundo, cnpj_fundo_investido, nome_fundo_investido



def processar_arquivo(csv_path: Path, pasta_mes: str, cnpj_fundo: str) -> pd.DataFrame | None:
    df = ler_csv_seguro(csv_path)
    if df is None or df.empty:
        return None

    col_cnpj = identificar_coluna_configurada(df, pasta_mes, csv_path, "CNPJ")
    if col_cnpj is None:
        return None

    df[col_cnpj] = df[col_cnpj].apply(normalizar_cnpj)
    df = df[df[col_cnpj] == cnpj_fundo].copy()

    if df.empty:
        return None

    col_valor = identificar_coluna_configurada(df, pasta_mes, csv_path, "Valor")
    if col_valor is None:
        return None

    eh_fundo_investido, cnpj_fundo_investido, nome_fundo_investido = linha_representa_fundo_investido(df, pasta_mes, csv_path)

    df["DataRef"] = inferir_data_da_pasta(pasta_mes)
    df["Codigo"] = escolher_codigo(df, pasta_mes, csv_path)
    df["ISIN"] = escolher_isin(df, pasta_mes, csv_path)
    df["Descrição"] = escolher_descricao(df, pasta_mes, csv_path)
    df["Emissor"] = escolher_emissor(df, pasta_mes, csv_path)
    df["Classe"] = escolher_classe(df, pasta_mes, csv_path)
    df["Tipo de Investimento"] = escolher_tipo_investimento(df, pasta_mes, csv_path)
    df["Vencimento"] = escolher_vencimento(df, pasta_mes, csv_path)
    df["VL_MERC_POS_FINAL_NUM"] = df[col_valor].apply(converter_valor)

    mask_valores_a_pagar = (
        df["Tipo de Investimento"].fillna("").astype(str).str.strip().str.lower().eq("valores a pagar")
        | df["Descrição"].fillna("").astype(str).str.strip().str.lower().eq("valores a pagar")
    )

    df.loc[mask_valores_a_pagar, "VL_MERC_POS_FINAL_NUM"] = (
        -df.loc[mask_valores_a_pagar, "VL_MERC_POS_FINAL_NUM"].abs()
    )

    df["EhFundoInvestido"] = eh_fundo_investido
    df["CNPJ_Fundo_Investido"] = cnpj_fundo_investido.where(eh_fundo_investido, "")
    df["Nome_Fundo_Investido"] = nome_fundo_investido.where(eh_fundo_investido, "")
    df["OrigemLookThrough"] = "DIRETO"
    df["FundoRaiz"] = cnpj_fundo
    df["NivelLookThrough"] = 0

    return df[colunas_base_internas()].copy()



def carregar_carteira_direta_competencia(
    pasta_mes: str,
    cnpj_fundo: str,
    cache_direta: dict[tuple[str, str], pd.DataFrame],
) -> pd.DataFrame:
    chave_cache = (pasta_mes, normalizar_cnpj(cnpj_fundo))
    if chave_cache in cache_direta:
        return cache_direta[chave_cache].copy()

    pasta = PASTA_BASE / pasta_mes
    data_mes = inferir_data_da_pasta(pasta_mes)
    bases = []

    if not pasta.exists():
        cache_direta[chave_cache] = pd.DataFrame(columns=colunas_base_internas())
        return cache_direta[chave_cache].copy()

    csvs = sorted(pasta.glob("*.csv"))
    print(f"\nLendo carteira direta do fundo {cnpj_fundo} na competência {pasta_mes} ({data_mes})...")

    for csv_path in csvs:
        nome_upper = csv_path.name.upper()

        if "CDA_FI_PL_" in nome_upper:
            continue

        try:
            base = processar_arquivo(csv_path, pasta_mes, cnpj_fundo)
            if base is not None and not base.empty:
                base = base[base["DataRef"] == data_mes].copy()
                if not base.empty:
                    bases.append(base)
                    print(f"[OK] {csv_path.name} | {len(base)} linha(s)")
        except Exception as e:
            print(f"[ERRO] {csv_path.name} | {e}")

    if bases:
        df = pd.concat(bases, ignore_index=True)
    else:
        df = pd.DataFrame(columns=colunas_base_internas())

    cache_direta[chave_cache] = df.copy()
    return df.copy()


# =========================================================
# LOOK-THROUGH DOS FUNDOS INVESTIDOS
# =========================================================
def agrupar_base_interna(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=colunas_base_internas())

    group_cols = COLUNAS_CHAVE_AGRUPAMENTO + [
        "EhFundoInvestido",
        "CNPJ_Fundo_Investido",
        "Nome_Fundo_Investido",
        "OrigemLookThrough",
        "FundoRaiz",
        "NivelLookThrough",
    ]

    df_agg = (
        df.groupby(group_cols, as_index=False)["VL_MERC_POS_FINAL_NUM"]
        .sum()
    )

    return df_agg[colunas_base_internas()].copy()



def distribuir_fundo_investido_em_ativos(
    linha_fundo: pd.Series,
    carteira_filho_aberta: pd.DataFrame,
    fundo_raiz: str,
    pasta_mes: str,
    cnpj_fundo_investido: str,
) -> pd.DataFrame:
    if carteira_filho_aberta is None or carteira_filho_aberta.empty:
        return pd.DataFrame(columns=colunas_base_internas())

    base_distribuicao = carteira_filho_aberta.copy()

    # Base que será distribuída nas linhas finais
    if LOOKTHROUGH_SOMENTE_POSITIVOS:
        base_saida = base_distribuicao[
            pd.to_numeric(base_distribuicao["VL_MERC_POS_FINAL_NUM"], errors="coerce").fillna(0.0) > 0
        ].copy()
    else:
        base_saida = base_distribuicao.copy()

    if base_saida.empty:
        return pd.DataFrame(columns=colunas_base_internas())

    valor_investido_no_fundo = float(linha_fundo.get("VL_MERC_POS_FINAL_NUM", 0.0) or 0.0)

    # Denominador correto = PL do fundo filho
    pl_filho = ler_pl_competencia(pasta_mes, cnpj_fundo_investido)

    # fallback caso não encontre PL
    if pl_filho is None or pl_filho == 0:
        total_base = float(
            pd.to_numeric(base_distribuicao["VL_MERC_POS_FINAL_NUM"], errors="coerce")
            .fillna(0.0)
            .sum()
        )
    else:
        total_base = float(pl_filho)

    if total_base == 0:
        return pd.DataFrame(columns=colunas_base_internas())

    fator = valor_investido_no_fundo / total_base

    base_saida["VL_MERC_POS_FINAL_NUM"] = base_saida["VL_MERC_POS_FINAL_NUM"] * fator
    base_saida["OrigemLookThrough"] = "LOOKTHROUGH"
    base_saida["FundoRaiz"] = fundo_raiz
    base_saida["NivelLookThrough"] = pd.to_numeric(
        base_saida["NivelLookThrough"], errors="coerce"
    ).fillna(0).astype(int) + 1

    return base_saida[colunas_base_internas()].copy()



def abrir_carteira_fundo(
    pasta_mes: str,
    cnpj_fundo: str,
    cache_direta: dict[tuple[str, str], pd.DataFrame],
    cache_aberta: dict[tuple[str, str], pd.DataFrame],
    pilha: tuple[str, ...] | None = None,
    fundo_raiz: str | None = None,
    profundidade: int = 0,
) -> pd.DataFrame:
    cnpj_fundo_norm = normalizar_cnpj(cnpj_fundo)
    fundo_raiz_norm = normalizar_cnpj(fundo_raiz or cnpj_fundo_norm)
    pilha = tuple(pilha or ())

    chave_cache = (pasta_mes, cnpj_fundo_norm)
    if chave_cache in cache_aberta:
        return cache_aberta[chave_cache].copy()

    carteira_direta = carregar_carteira_direta_competencia(
        pasta_mes=pasta_mes,
        cnpj_fundo=cnpj_fundo_norm,
        cache_direta=cache_direta,
    )

    if carteira_direta.empty:
        cache_aberta[chave_cache] = carteira_direta.copy()
        return carteira_direta.copy()

    if (
        not ATIVAR_LOOKTHROUGH_FUNDOS
        or profundidade >= PROFUNDIDADE_MAXIMA_LOOKTHROUGH
        or cnpj_fundo_norm in pilha
    ):
        if profundidade >= PROFUNDIDADE_MAXIMA_LOOKTHROUGH:
            print(
                f"[AVISO] Profundidade máxima de look-through atingida para o fundo {cnpj_fundo_norm} "
                f"na competência {pasta_mes}."
            )
        if cnpj_fundo_norm in pilha:
            print(
                f"[AVISO] Possível circularidade detectada no look-through do fundo {cnpj_fundo_norm} "
                f"na competência {pasta_mes}."
            )

        carteira_retorno = carteira_direta.copy()
        carteira_retorno["FundoRaiz"] = fundo_raiz_norm
        cache_aberta[chave_cache] = carteira_retorno.copy()
        return carteira_retorno.copy()

    proximapilha = pilha + (cnpj_fundo_norm,)

    diretos_nao_fundos = carteira_direta[~carteira_direta["EhFundoInvestido"]].copy()
    cotas_fundos = carteira_direta[carteira_direta["EhFundoInvestido"]].copy()

    bases_resultado = []

    if not diretos_nao_fundos.empty:
        diretos_nao_fundos["FundoRaiz"] = fundo_raiz_norm
        bases_resultado.append(diretos_nao_fundos)

    for _, linha_fundo in cotas_fundos.iterrows():
        cnpj_fundo_investido = normalizar_cnpj(linha_fundo.get("CNPJ_Fundo_Investido", ""))
        nome_fundo_investido = str(linha_fundo.get("Nome_Fundo_Investido", "")).strip()
        valor_aplicado = float(linha_fundo.get("VL_MERC_POS_FINAL_NUM", 0.0) or 0.0)

        if MANTER_LINHA_COTA_FUNDO:
            linha_cota = linha_fundo.to_frame().T.copy()
            linha_cota["FundoRaiz"] = fundo_raiz_norm
            bases_resultado.append(linha_cota[colunas_base_internas()])

        if cnpj_fundo_investido == "":
            if not MANTER_LINHA_COTA_FUNDO and MANTER_COTA_SE_NAO_ABRIR:
                linha_cota = linha_fundo.to_frame().T.copy()
                linha_cota["FundoRaiz"] = fundo_raiz_norm
                bases_resultado.append(linha_cota[colunas_base_internas()])
            continue

        if valor_aplicado == 0:
            continue

        print(
            f"[LOOKTHROUGH] {cnpj_fundo_norm} -> {cnpj_fundo_investido} "
            f"({nome_fundo_investido or 'FUNDO SEM NOME'}) | Valor aplicado: {valor_aplicado:,.2f}"
        )

        carteira_filho_aberta = abrir_carteira_fundo(
            pasta_mes=pasta_mes,
            cnpj_fundo=cnpj_fundo_investido,
            cache_direta=cache_direta,
            cache_aberta=cache_aberta,
            pilha=proximapilha,
            fundo_raiz=fundo_raiz_norm,
            profundidade=profundidade + 1,
        )

        carteira_expandida = distribuir_fundo_investido_em_ativos(
            linha_fundo=linha_fundo,
            carteira_filho_aberta=carteira_filho_aberta,
            fundo_raiz=fundo_raiz_norm,
            pasta_mes=pasta_mes,
            cnpj_fundo_investido=cnpj_fundo_investido,
        )

        if carteira_expandida.empty:
            print(
                f"[AVISO] Não foi possível abrir a carteira do fundo investido {cnpj_fundo_investido} "
                f"na competência {pasta_mes}."
            )
            if not MANTER_LINHA_COTA_FUNDO and MANTER_COTA_SE_NAO_ABRIR:
                linha_cota = linha_fundo.to_frame().T.copy()
                linha_cota["FundoRaiz"] = fundo_raiz_norm
                bases_resultado.append(linha_cota[colunas_base_internas()])
            continue

        bases_resultado.append(carteira_expandida)

    if bases_resultado:
        carteira_aberta = pd.concat(bases_resultado, ignore_index=True)
        carteira_aberta = agrupar_base_interna(carteira_aberta)
    else:
        carteira_aberta = pd.DataFrame(columns=colunas_base_internas())

    cache_aberta[chave_cache] = carteira_aberta.copy()
    return carteira_aberta.copy()


# =========================================================
# CONSOLIDAÇÃO FINAL
# =========================================================
def definir_competencias_para_processar(
    competencias_disponiveis: list[str],
    datas_existentes: set[str],
) -> list[str]:
    if not competencias_disponiveis:
        return []

    if REPROCESSAR_COMPETENCIAS:
        return [c for c in REPROCESSAR_COMPETENCIAS if c in competencias_disponiveis]

    ultimos = set(competencias_disponiveis[-REPROCESSAR_ULTIMOS_MESES:]) if REPROCESSAR_ULTIMOS_MESES > 0 else set()

    competencias_processar = []

    for comp in competencias_disponiveis:
        data_ref = inferir_data_da_pasta(comp)

        if comp in ultimos:
            competencias_processar.append(comp)
            continue

        if data_ref not in datas_existentes:
            competencias_processar.append(comp)

    return competencias_processar



def consolidar_competencias(
    competencias_processar: list[str],
    cnpj_fundo: str,
) -> tuple[pd.DataFrame | None, pd.DataFrame]:
    bases = []
    cache_direta: dict[tuple[str, str], pd.DataFrame] = {}
    cache_aberta: dict[tuple[str, str], pd.DataFrame] = {}

    for comp in competencias_processar:
        pasta = PASTA_BASE / comp
        if not pasta.exists():
            print(f"[AVISO] Pasta {comp} não encontrada.")
            continue

        data_mes = inferir_data_da_pasta(comp)
        print(f"\nProcessando competência {comp} ({data_mes})...")

        base_mes = abrir_carteira_fundo(
            pasta_mes=comp,
            cnpj_fundo=cnpj_fundo,
            cache_direta=cache_direta,
            cache_aberta=cache_aberta,
            pilha=(),
            fundo_raiz=cnpj_fundo,
            profundidade=0,
        )

        if base_mes is None or base_mes.empty:
            print(f"[INFO] {comp} sem dados do fundo.")
            continue

        base_mes = base_mes[base_mes["DataRef"] == data_mes].copy()
        if base_mes.empty:
            print(f"[INFO] {comp} sem dados do fundo após look-through.")
            continue

        bases.append(base_mes)
        print(f"[OK] Competência {comp} consolidada com {len(base_mes)} linha(s) após look-through.")

    if not bases:
        return None, pd.DataFrame(columns=["DataRef", "Soma_Coluna", "PL_Arquivo", "Diferenca", "Status"])

    df = pd.concat(bases, ignore_index=True)

    df_agg = (
        df.groupby(COLUNAS_CHAVE_AGRUPAMENTO, as_index=False)["VL_MERC_POS_FINAL_NUM"]
        .sum()
    )

    df_pivot = df_agg.pivot_table(
        index=COLUNAS_FIXAS,
        columns="DataRef",
        values="VL_MERC_POS_FINAL_NUM",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    somas_por_data = {
        col: float(df_pivot[col].sum())
        for col in df_pivot.columns
        if col not in COLUNAS_FIXAS
    }

    df_conferencia = montar_resumo_conferencia(
        competencias_processar=competencias_processar,
        somas_por_data=somas_por_data,
        cnpj_fundo=cnpj_fundo,
    )

    return df_pivot, df_conferencia



def remover_colunas_reprocessadas(df_antigo: pd.DataFrame, competencias_processar: list[str]) -> pd.DataFrame:
    datas_reprocessadas = {inferir_data_da_pasta(comp) for comp in competencias_processar}
    colunas_remover = [c for c in df_antigo.columns if c in datas_reprocessadas]

    if colunas_remover:
        print(f"Removendo colunas antigas para sobrescrever: {', '.join(colunas_remover)}")
        df_antigo = df_antigo.drop(columns=colunas_remover, errors="ignore")

    colunas_datas_restantes = [c for c in df_antigo.columns if c not in COLUNAS_FIXAS]

    if not colunas_datas_restantes:
        return pd.DataFrame(columns=COLUNAS_FIXAS)

    valores_restantes = (
        df_antigo[colunas_datas_restantes]
        .apply(pd.to_numeric, errors="coerce")
        .fillna(0.0)
        .sum(axis=1)
    )

    df_antigo = df_antigo.loc[valores_restantes != 0].copy()

    return df_antigo



def mesclar_bases(df_antigo: pd.DataFrame | None, df_novo: pd.DataFrame | None, competencias_processar: list[str]) -> pd.DataFrame:
    if df_antigo is None and df_novo is None:
        raise ValueError("Não há dados para consolidar.")

    if df_antigo is not None:
        df_antigo = normalizar_planilha_existente(df_antigo)
        df_antigo = remover_colunas_reprocessadas(df_antigo, competencias_processar)

        if df_antigo.empty:
            df_antigo = None

    if df_antigo is None:
        df_final = df_novo.copy()
    elif df_novo is None:
        df_final = df_antigo.copy()
    else:
        df_final = pd.merge(
            df_antigo,
            df_novo,
            on=COLUNAS_FIXAS,
            how="outer",
        )

    for col in COLUNAS_FIXAS:
        if col not in df_final.columns:
            df_final[col] = ""

    colunas_data = [c for c in df_final.columns if c not in COLUNAS_FIXAS]
    for c in colunas_data:
        df_final[c] = pd.to_numeric(df_final[c], errors="coerce").fillna(0.0)

    colunas_data = sorted(colunas_data)
    df_final = df_final[COLUNAS_FIXAS + colunas_data]

    df_final = df_final.sort_values(
        by=["Tipo de Investimento", "Descrição", "Codigo"],
        ascending=[True, True, True],
    ).reset_index(drop=True)

    return df_final



def formatar_aba_excel(ws) -> None:
    ws.freeze_panes = "H2"

    larguras = {
        "A": 18,
        "B": 18,
        "C": 18,
        "D": 18,
        "E": 18,
        "F": 18,
        "G": 18,
    }

    for col, largura in larguras.items():
        ws.column_dimensions[col].width = largura

    for idx in range(8, ws.max_column + 1):
        letra = ws.cell(row=1, column=idx).column_letter
        ws.column_dimensions[letra].width = 14

    for row in ws.iter_rows(min_row=2, min_col=8, max_col=ws.max_column):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'



def exportar_excel_multiplas_abas(
    abas: dict[str, pd.DataFrame],
    conferencias: dict[str, pd.DataFrame],
) -> None:
    abas_validas = {
        normalizar_nome_aba(nome): df
        for nome, df in abas.items()
        if df is not None and not df.empty
    }

    if not abas_validas:
        print("Nenhuma aba com dados para salvar.")
        return

    with pd.ExcelWriter(ARQUIVO_SAIDA, engine="openpyxl") as writer:
        for nome_aba, df_final in abas_validas.items():
            df_final.to_excel(writer, sheet_name=nome_aba, index=False)
            ws = writer.book[nome_aba]

            formatar_aba_excel(ws)

            df_conferencia = conferencias.get(nome_aba)

            mapa_conf = {}
            if df_conferencia is not None and not df_conferencia.empty:
                for _, r in df_conferencia.iterrows():
                    mapa_conf[str(r["DataRef"])] = {
                        "Soma_Coluna": r["Soma_Coluna"],
                        "PL_Arquivo": r["PL_Arquivo"],
                        "Diferenca": r["Diferenca"],
                        "Status": r["Status"],
                    }

            linha_inicio = ws.max_row + 2

            ws.cell(row=linha_inicio, column=1, value="SOMA_COLUNA")
            ws.cell(row=linha_inicio + 1, column=1, value="PL_CDA_FI_PL")
            ws.cell(row=linha_inicio + 2, column=1, value="DIFERENCA")
            ws.cell(row=linha_inicio + 3, column=1, value="STATUS")

            for idx in range(8, ws.max_column + 1):
                data_coluna = str(ws.cell(row=1, column=idx).value)
                info = mapa_conf.get(data_coluna)

                if info is None:
                    ws.cell(row=linha_inicio + 3, column=idx, value="SEM_CONF")
                    continue

                ws.cell(row=linha_inicio, column=idx, value=info["Soma_Coluna"])
                ws.cell(row=linha_inicio + 1, column=idx, value=info["PL_Arquivo"])
                ws.cell(row=linha_inicio + 2, column=idx, value=info["Diferenca"])
                ws.cell(row=linha_inicio + 3, column=idx, value=info["Status"])

                ws.cell(row=linha_inicio, column=idx).number_format = '#,##0.00'
                ws.cell(row=linha_inicio + 1, column=idx).number_format = '#,##0.00'
                ws.cell(row=linha_inicio + 2, column=idx).number_format = '#,##0.00'

    print(f"\nArquivo salvo em: {ARQUIVO_SAIDA}")



def main():
    print("Mapa de colunas por competência e planilha:")
    imprimir_mapa_colunas()

    print("\nLendo competências disponíveis...")
    competencias_disponiveis = obter_competencias_disponiveis()
    print(f"Competências encontradas: {len(competencias_disponiveis)}")

    print("Carregando planilha atual...")
    abas_existentes = carregar_todas_abas_existentes()
    print(f"Abas já existentes no arquivo: {len(abas_existentes)}")

    fundos_processar = obter_fundos_para_processar()
    if not fundos_processar:
        print("Nenhum fundo válido foi informado em FUNDOS / FUNDOS_PROCESSAR.")
        return

    print("\nFundos que serão processados:")
    for nome_aba, cnpj in fundos_processar.items():
        print(f" - {nome_aba} | {cnpj}")

    abas_finais = abas_existentes.copy()
    conferencias_finais = {}
    houve_atualizacao = False

    for nome_aba, cnpj_fundo in fundos_processar.items():
        print(f"\n{'=' * 70}")
        print(f"Processando fundo: {nome_aba}")
        print(f"CNPJ: {cnpj_fundo}")
        print(f"{'=' * 70}")

        df_antigo = abas_existentes.get(nome_aba)
        datas_existentes = obter_datas_existentes_da_aba(df_antigo)
        print(f"Datas já existentes na aba {nome_aba}: {len(datas_existentes)}")

        competencias_processar = definir_competencias_para_processar(
            competencias_disponiveis=competencias_disponiveis,
            datas_existentes=datas_existentes,
        )

        if not competencias_processar:
            print(f"Nenhuma competência nova ou marcada para reprocessamento para a aba {nome_aba}.")
            if df_antigo is not None:
                abas_finais[nome_aba] = df_antigo
            continue

        print("\nCompetências que serão processadas:")
        for comp in competencias_processar:
            print(f" - {comp} ({inferir_data_da_pasta(comp)})")

        df_novo, df_conferencia = consolidar_competencias(competencias_processar, cnpj_fundo)

        if df_novo is None:
            print(f"Nenhum dado novo encontrado para o fundo {nome_aba}.")
            if df_antigo is not None:
                abas_finais[nome_aba] = df_antigo
            continue

        print("\nMesclando base antiga com dados novos/reprocessados...")
        df_final = mesclar_bases(df_antigo, df_novo, competencias_processar)

        abas_finais[nome_aba] = df_final
        conferencias_finais[nome_aba] = df_conferencia
        houve_atualizacao = True

    if not houve_atualizacao and ARQUIVO_SAIDA.exists():
        print("Nenhuma aba precisou ser atualizada.")
        return

    print("\nSalvando arquivo final...")
    exportar_excel_multiplas_abas(abas_finais, conferencias_finais)

    print("\nAplicando tratamento automático da carteira...")
    executar_tratamento()

    print("Atualização concluída com sucesso.")


if __name__ == "__main__":
    main()
