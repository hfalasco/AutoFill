"""Certificado Analítico - Água Potável (FO.GSM.LMI.07.01): leitura, planejamento e preenchimento.

Fluxo de um ponto (ex.: Área de vivência 417):
1. `ler_arquivo` lê as abas já emitidas do 'Vivência 417.xlsx' (ou `novo_arquivo` parte do modelo);
2. `planejar` acha a coleta do mês no FO.GSM.LMI.03.08, avalia cada parâmetro e decide a aba;
3. `gerar_xlsx` copia a última aba e grava os valores, sem perder logotipo nem assinatura.
"""

from __future__ import annotations

import calendar
import io
import re
import warnings
from dataclasses import replace
from datetime import date
from pathlib import Path

import openpyxl

from autofill.model.config import CAMINHO_MODELO, MESES, Config, Parametro, carregar_config, normalizar
from autofill.model.dados import (
    Amostra,
    ArquivoCertificado,
    Certificado,
    InfoCertificado,
    LinhaResultado,
)
from autofill.services.analises import Analises
from autofill.services.xlsx import PastaXlsx

LINHAS_LIDAS = 60
FORMATO_DATA = "dd/mm/yyyy"
COLUNAS_LIDAS = 60


class CertificadoInvalido(ValueError):
    """O arquivo não tem nenhuma aba no formato do FO.GSM.LMI.07.01."""


# --------------------------------------------------------------------------- utilidades
def somar_meses(dia: date, meses: int) -> date:
    total = dia.month - 1 + meses
    ano, mes = dia.year + total // 12, total % 12 + 1
    return date(ano, mes, min(dia.day, calendar.monthrange(ano, mes)[1]))


def indice_mes(nome: str) -> int | None:
    alvo = normalizar(nome)
    for i, mes in enumerate(MESES, start=1):
        if normalizar(mes) == alvo:
            return i
    return None


def _inteiro(valor: object) -> int | None:
    try:
        return int(float(str(valor).replace(",", ".")))
    except (TypeError, ValueError):
        return None


def _texto(valor: object) -> str:
    return "" if valor is None else str(valor).strip()


def numero_do_texto(*textos: str) -> str:
    """Último número que aparece no primeiro texto que tiver algum ('Área de vivência - 417' -> '417')."""
    for texto in textos:
        numeros = re.findall(r"\d+", texto or "")
        if numeros:
            return numeros[-1]
    return ""


# --------------------------------------------------------------------------- leitura
def _celulas(ws) -> dict[str, object]:
    valores = {}
    for linha in ws.iter_rows(min_row=1, max_row=LINHAS_LIDAS, max_col=COLUNAS_LIDAS):
        for c in linha:
            if getattr(c, "value", None) is not None:
                valores[c.coordinate] = c.value
    return valores


def _info_aba(nome: str, valores: dict[str, object], config: Config) -> InfoCertificado | None:
    layout = config.certificado
    cel = layout.celulas
    rotulo_local = "A" + re.sub(r"\D", "", cel["local"])
    if not normalizar(valores.get(rotulo_local)).startswith("local de amostragem"):
        return None
    linhas: dict[str, int] = {}
    inicio, fim = layout.linhas_parametros
    for n in range(inicio, fim + 1):
        nome_param = normalizar(valores.get(f"{layout.coluna_nome_parametro}{n}"))
        if not nome_param:
            continue
        for p in config.parametros:
            if p.id not in linhas and any(nome_param.startswith(apelido) for apelido in p.nomes_no_certificado):
                linhas[p.id] = n
                break
    return InfoCertificado(
        aba=nome,
        mes=_texto(valores.get(cel["mes"])),
        ano=_inteiro(valores.get(cel["ano"])),
        numero=_inteiro(valores.get(cel["numero"])),
        versao=valores.get(cel["versao"]),
        local=_texto(valores.get(cel["local"])),
        veiculo=_texto(valores.get(cel["veiculo"])),
        ponto=_texto(valores.get(cel["ponto"])),
        coletor=_texto(valores.get(cel["coletor"])),
        linhas_parametros=linhas,
    )


def _ordem(info: InfoCertificado) -> tuple[int, int]:
    return info.ano or 0, indice_mes(info.mes) or 0


def ler_arquivo(conteudo: bytes, nome_arquivo: str, config: Config | None = None) -> ArquivoCertificado:
    """Lê um 'Vivência NNN.xlsx': as abas de certificado em ordem cronológica e o ponto."""
    config = config or carregar_config()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            wb = openpyxl.load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
        except Exception as erro:
            raise CertificadoInvalido(f"{nome_arquivo}: não foi possível abrir ({erro})") from erro
    try:
        abas = [
            info
            for ws in wb.worksheets
            if ws.sheet_state == "visible" and (info := _info_aba(ws.title, _celulas(ws), config)) is not None
        ]
    finally:
        wb.close()
    if not abas:
        raise CertificadoInvalido(f"{nome_arquivo}: nenhuma aba no formato do Certificado Analítico (FO.GSM.LMI.07.01).")
    abas.sort(key=_ordem)
    ultimo = abas[-1]
    veiculo = "" if ultimo.veiculo in ("", "-") else ultimo.veiculo
    identificacao = numero_do_texto(ultimo.ponto, veiculo, Path(nome_arquivo).stem)
    return ArquivoCertificado(nome_arquivo, conteudo, ultimo.local or config.origem_padrao, identificacao, abas)


def novo_arquivo(identificacao: str, origem: str | None = None, config: Config | None = None) -> ArquivoCertificado:
    """Arquivo para um ponto que ainda não tem planilha, a partir do modelo em assets/."""
    config = config or carregar_config()
    origem = origem or config.origem_padrao
    modelo = ler_arquivo(CAMINHO_MODELO.read_bytes(), CAMINHO_MODELO.name, config)
    base = modelo.ultimo
    info = replace(base, local=origem, veiculo="-", ponto=f"{origem} - {identificacao}", numero=0)
    nome = config.nome_arquivo_novo.format(ponto=identificacao, origem=origem)
    return ArquivoCertificado(nome, modelo.conteudo, origem, identificacao, [info], novo=True)


# --------------------------------------------------------------------------- avaliação
def _numero(valor: object) -> float | None:
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    m = re.fullmatch(r"\s*[<>≤≥]?\s*(-?\d+(?:[.,]\d+)?)\s*", str(valor or ""))
    return float(m.group(1).replace(",", ".")) if m else None


def valor_para_certificado(parametro: Parametro, valor: object) -> object:
    """Como o resultado vai escrito: 'A' vira 'Ausente', '0,5' vira 0.5."""
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        return None
    if parametro.regra.tipo == "ausente":
        texto = normalizar(valor)
        if texto in ("a", "aus", "ausente"):
            return "Ausente"
        if texto in ("p", "pres", "presente"):
            return "Presente"
        return str(valor).strip()
    if isinstance(valor, str):
        numero = _numero(valor)
        if numero is not None and re.fullmatch(r"\s*-?\d+(?:[.,]\d+)?\s*", valor):
            return int(numero) if numero.is_integer() else numero
        return valor.strip()
    if isinstance(valor, float) and valor.is_integer():
        return int(valor)
    return valor


def avaliar(parametro: Parametro, valor: object, config: Config | None = None) -> str:
    config = config or carregar_config()
    if valor is None:
        return config.conferir
    regra = parametro.regra
    if regra.tipo == "ausente":
        texto = normalizar(valor)
        if texto == "ausente":
            return config.aceitavel
        if texto == "presente":
            return config.nao_aceitavel
        return config.conferir
    numero = _numero(valor)
    if numero is None:
        return config.conferir
    if regra.minimo is not None and numero < regra.minimo:
        return config.nao_aceitavel
    if regra.maximo is not None and numero > regra.maximo:
        return config.nao_aceitavel
    return config.aceitavel


def linhas_resultado(amostra: Amostra, config: Config | None = None) -> list[LinhaResultado]:
    config = config or carregar_config()
    linhas = []
    for p in config.parametros:
        valor = valor_para_certificado(p, amostra.resultados.get(p.id))
        linhas.append(LinhaResultado(p.id, p.nome, valor, p.vmp, avaliar(p, valor, config)))
    return linhas


# --------------------------------------------------------------------------- planejamento
def candidatas(arquivo: ArquivoCertificado, analises: Analises, mes: int, ano: int) -> list[Amostra]:
    return sorted(
        (a for a in analises.do_ponto(arquivo.origem, arquivo.identificacao) if (a.data.year, a.data.month) == (ano, mes)),
        key=lambda a: (a.data, a.linha),
    )


def planejar(
    arquivo: ArquivoCertificado,
    analises: Analises,
    mes: int,
    ano: int,
    usar_anterior: bool = False,
    config: Config | None = None,
) -> Certificado:
    """Decide a coleta, a aba e o número do certificado do mês `mes`/`ano` para um ponto."""
    config = config or carregar_config()
    nome_mes = MESES[mes - 1]
    existente = next(
        (a for a in arquivo.abas if indice_mes(a.mes) == mes and a.ano == ano and not arquivo.novo), None
    )
    if existente is not None:
        aba_destino, numero, sobrescreve = existente.aba, existente.numero or 1, True
    else:
        nomes = {normalizar(a.aba) for a in arquivo.abas} if not arquivo.novo else set()
        aba_destino = nome_mes if normalizar(nome_mes) not in nomes else f"{nome_mes} {ano}"
        numero = max((a.numero or 0 for a in arquivo.abas), default=0) + 1
        sobrescreve = False

    opcoes = candidatas(arquivo, analises, mes, ano)
    avisos: list[str] = []
    if not opcoes and usar_anterior:
        fim_do_mes = date(ano, mes, calendar.monthrange(ano, mes)[1])
        anteriores = sorted(
            (a for a in analises.do_ponto(arquivo.origem, arquivo.identificacao) if a.data <= fim_do_mes),
            key=lambda a: (a.data, a.linha),
        )
        if anteriores:
            opcoes = anteriores[-1:]
            avisos.append(f"Sem coleta em {nome_mes}/{ano}: usada a de {opcoes[0].data:%d/%m/%Y}.")
    if not opcoes:
        avisos.append(f"Nenhuma coleta de {arquivo.origem} {arquivo.identificacao} em {nome_mes}/{ano}.")
    if sobrescreve:
        avisos.append(f"A aba '{aba_destino}' ({nome_mes}/{ano}) já existe e será atualizada.")

    cert = Certificado(
        arquivo=arquivo,
        mes=nome_mes,
        ano=ano,
        numero=numero,
        aba_destino=aba_destino,
        sobrescreve=sobrescreve,
        amostra=None,
        candidatas=opcoes,
        coletor=arquivo.ultimo.coletor,
        avisos=avisos,
    )
    if opcoes:
        # Mais de uma coleta no mês costuma ser recoleta depois de 'Presente': vale a mais recente.
        escolher_amostra(cert, opcoes[-1], config)
    return cert


def escolher_amostra(cert: Certificado, amostra: Amostra, config: Config | None = None) -> Certificado:
    config = config or carregar_config()
    cert.amostra = amostra
    cert.validade = somar_meses(amostra.data, config.validade_meses)
    cert.linhas = linhas_resultado(amostra, config)
    cert.avisos = [a for a in cert.avisos if not a.startswith(("Há ", "Resultado "))]
    if len(cert.candidatas) > 1:
        datas = ", ".join(f"{a.data:%d/%m}" for a in cert.candidatas)
        cert.avisos.append(f"Há {len(cert.candidatas)} coletas no mês ({datas}); usada a de {amostra.data:%d/%m/%Y}.")
    for linha in cert.linhas:
        if linha.situacao != config.aceitavel:
            valor = "vazio" if linha.valor is None else linha.valor
            cert.avisos.append(f"Resultado {linha.situacao.lower()}: {linha.nome} = {valor}.")
    return cert


# --------------------------------------------------------------------------- geração
def valores_celulas(cert: Certificado, config: Config | None = None) -> dict[str, object]:
    config = config or carregar_config()
    layout = config.certificado
    cel = layout.celulas
    info = cert.arquivo.ultimo
    valores: dict[str, object] = {
        cel["mes"]: cert.mes,
        cel["ano"]: cert.ano,
        cel["numero"]: cert.numero,
        cel["data_amostragem"]: cert.amostra.data,
        cel["validade"]: cert.validade,
        cel["coletor"]: cert.coletor,
    }
    if cert.arquivo.novo:
        valores |= {cel["local"]: info.local, cel["veiculo"]: info.veiculo, cel["ponto"]: info.ponto}
    for linha in cert.linhas:
        n = info.linhas_parametros.get(linha.parametro_id)
        if n is None:
            continue
        valores[f"{layout.coluna_resultado}{n}"] = linha.valor if linha.valor is not None else "-"
        valores[f"{layout.coluna_aceitavel}{n}"] = linha.situacao
    return valores


def gerar_xlsx(cert: Certificado, config: Config | None = None) -> bytes:
    """O arquivo do ponto com a aba do certificado do mês preenchida (e aberta ao abrir o arquivo)."""
    if not cert.pronto:
        raise ValueError("Certificado sem coleta selecionada.")
    config = config or carregar_config()
    pasta = PastaXlsx(cert.arquivo.conteudo)
    if cert.sobrescreve:
        destino = cert.aba_destino
    elif cert.arquivo.novo:
        pasta.renomear_aba(cert.arquivo.ultimo.aba, cert.aba_destino)
        destino = cert.aba_destino
    else:
        destino = pasta.clonar_aba(cert.arquivo.ultimo.aba, cert.aba_destino).nome
    pasta.definir_celulas(destino, valores_celulas(cert, config))
    cel = config.certificado.celulas
    pasta.formatar_numero(destino, [cel["data_amostragem"], cel["validade"]], FORMATO_DATA)
    pasta.ativar_aba(destino)
    return pasta.salvar()



def ler_lista_pontos(texto: str, limite_faixa: int = 500) -> list[str]:
    """'631, 634; 665' ou '600-605' -> ['631', '634', '665'] / ['600', ..., '605'], sem repetir."""
    pontos: list[str] = []
    for token in re.split(r"[,;\n\s]+", texto or ""):
        token = token.strip()
        if not token:
            continue
        faixa = re.fullmatch(r"(\d+)\s*(?:-|a|até)\s*(\d+)", token)
        if faixa and 0 <= int(faixa.group(2)) - int(faixa.group(1)) <= limite_faixa:
            novos = [str(n) for n in range(int(faixa.group(1)), int(faixa.group(2)) + 1)]
        else:
            novos = [token]
        pontos.extend(p for p in novos if p not in pontos)
    return pontos
