from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path
from datetime import datetime

import requests


# =========================================================
# CONFIGURAÇÕES
# =========================================================
PASTA_BASE = Path(r"Z:\Asset Management\Equipe\Lívia\Rascunhos\Site_fundos\Base Carteira Fundos")

URL_MENSAL_BASE = "https://dados.cvm.gov.br/dados/FI/DOC/CDA/DADOS"
URL_HIST_BASE = "https://dados.cvm.gov.br/dados/FI/DOC/CDA/DADOS/HIST"

INICIO_ANO = 2022
INICIO_MES = 12

TIMEOUT = 120
HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

REPROCESSAR_COMPETENCIAS = ["202512"]   # deixe [] quando não quiser forçar nada

# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================
def gerar_competencias(inicio_ano: int, inicio_mes: int) -> list[str]:
    """
    Gera lista de competências no formato YYYYMM desde o mês inicial
    até o mês anterior ao mês atual.
    """
    hoje = datetime.today()
    ano = inicio_ano
    mes = inicio_mes
    competencias = []

    while (ano < hoje.year) or (ano == hoje.year and mes < hoje.month):
        competencias.append(f"{ano}{mes:02d}")
        mes += 1
        if mes == 13:
            mes = 1
            ano += 1

    return competencias


def pasta_mes(comp: str) -> Path:
    """
    Retorna o caminho da pasta do mês no formato YYYYMM.
    """
    return PASTA_BASE / comp


def mes_ja_baixado(comp: str) -> bool:
    """
    Considera o mês baixado se a pasta existir e tiver pelo menos 1 arquivo.
    """
    pasta = pasta_mes(comp)
    return pasta.exists() and any(pasta.iterdir())


def listar_meses_pendentes(competencias: list[str]) -> list[str]:
    return [comp for comp in competencias if not mes_ja_baixado(comp)]


def baixar_bytes(url: str) -> bytes:
    resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    if resp.status_code != 200:
        raise RuntimeError(f"Falha ao baixar {url} | Status code: {resp.status_code}")
    return resp.content


def extrair_zip_bytes(zip_bytes: bytes, destino: Path) -> list[str]:
    """
    Extrai todos os arquivos do zip para a pasta destino.
    """
    destino.mkdir(parents=True, exist_ok=True)
    extraidos = []

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for nome in zf.namelist():
            if nome.endswith("/"):
                continue

            arquivo_destino = destino / Path(nome).name
            with zf.open(nome) as origem, open(arquivo_destino, "wb") as saida:
                saida.write(origem.read())

            extraidos.append(arquivo_destino.name)

    return extraidos


def extrair_apenas_mes_do_zip_anual(zip_bytes: bytes, competencia: str, destino: Path) -> list[str]:
    """
    Para o histórico anual (caso de 2022), extrai apenas os arquivos
    relacionados ao mês desejado, por exemplo 202212.
    """
    destino.mkdir(parents=True, exist_ok=True)
    extraidos = []

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        nomes = zf.namelist()

        # tenta achar arquivos que contenham a competência no nome
        candidatos = [n for n in nomes if competencia in n and not n.endswith("/")]

        if not candidatos:
            raise RuntimeError(
                f"Não encontrei arquivos da competência {competencia} dentro do zip anual."
            )

        for nome in candidatos:
            arquivo_destino = destino / Path(nome).name
            with zf.open(nome) as origem, open(arquivo_destino, "wb") as saida:
                saida.write(origem.read())

            extraidos.append(arquivo_destino.name)

    return extraidos


def baixar_mes_mensal(competencia: str) -> list[str]:
    """
    Baixa o zip mensal da CVM e extrai tudo dentro da pasta YYYYMM.
    Exemplo: cda_fi_202301.zip
    """
    url = f"{URL_MENSAL_BASE}/cda_fi_{competencia}.zip"
    destino = pasta_mes(competencia)

    print(f"Baixando {competencia} da pasta mensal...")
    zip_bytes = baixar_bytes(url)
    extraidos = extrair_zip_bytes(zip_bytes, destino)

    return extraidos


def baixar_dez2022_historico() -> list[str]:
    """
    Dez/2022 vem do zip histórico anual cda_fi_2022.zip.
    """
    competencia = "202212"
    destino = pasta_mes(competencia)
    url = f"{URL_HIST_BASE}/cda_fi_2022.zip"

    print("Baixando 202212 do histórico anual...")
    zip_bytes = baixar_bytes(url)
    extraidos = extrair_apenas_mes_do_zip_anual(zip_bytes, competencia, destino)

    return extraidos


# =========================================================
# EXECUÇÃO PRINCIPAL
# =========================================================
def main():
    PASTA_BASE.mkdir(parents=True, exist_ok=True)

    if REPROCESSAR_COMPETENCIAS:
        pendentes = REPROCESSAR_COMPETENCIAS
        print(f"Reprocessando competências forçadas: {', '.join(pendentes)}")
    else:
        competencias = gerar_competencias(INICIO_ANO, INICIO_MES)
        pendentes = listar_meses_pendentes(competencias)
        print(f"Meses pendentes: {', '.join(pendentes)}")

    if not pendentes:
        print("Nenhum mês novo para baixar.")
        return
    print("-" * 80)

    for comp in pendentes:
        try:
            if comp == "202212":
                arquivos = baixar_dez2022_historico()
            else:
                arquivos = baixar_mes_mensal(comp)

            print(f"[OK] {comp} baixado com {len(arquivos)} arquivo(s).")
            for arq in arquivos:
                print(f"   - {arq}")

        except Exception as e:
            print(f"[ERRO] Falha ao processar {comp}: {e}")

        print("-" * 80)

    print("Processo finalizado.")


if __name__ == "__main__":
    main()