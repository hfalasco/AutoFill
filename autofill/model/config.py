"""Regras do certificado (data/certificado.json) carregadas em objetos."""

from __future__ import annotations

import json
import unicodedata
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
CAMINHO_CONFIG = RAIZ / "data" / "certificado.json"
CAMINHO_MODELO = RAIZ / "assets" / "modelo_certificado.xlsx"

MESES = (
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
)


def normalizar(texto: object) -> str:
    """Minúsculas, sem acento e com espaços simples: base de toda comparação de texto."""
    if texto is None:
        return ""
    sem_acento = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return " ".join(sem_acento.lower().split())


@dataclass(frozen=True)
class Regra:
    tipo: str  # "ausente" ou "faixa"
    minimo: float | None = None
    maximo: float | None = None


@dataclass(frozen=True)
class Parametro:
    id: str
    nome: str
    nomes_no_certificado: tuple[str, ...]
    vmp: str
    regra: Regra


@dataclass(frozen=True)
class Coluna:
    cabecalhos: tuple[str, ...]
    letra: str


@dataclass(frozen=True)
class LayoutCertificado:
    celulas: dict[str, str]
    coluna_nome_parametro: str
    coluna_resultado: str
    coluna_vmp: str
    coluna_aceitavel: str
    linhas_parametros: tuple[int, int]


@dataclass(frozen=True)
class Config:
    validade_meses: int
    nome_arquivo_novo: str
    origem_padrao: str
    aceitavel: str
    nao_aceitavel: str
    conferir: str
    aba_analises: str
    colunas: dict[str, Coluna]
    certificado: LayoutCertificado
    parametros: tuple[Parametro, ...]

    def parametro(self, id_: str) -> Parametro:
        return next(p for p in self.parametros if p.id == id_)


@lru_cache
def carregar_config(caminho: Path = CAMINHO_CONFIG) -> Config:
    dados = json.loads(Path(caminho).read_text(encoding="utf-8"))
    cert = dados["certificado"]
    return Config(
        validade_meses=dados["validade_meses"],
        nome_arquivo_novo=dados["nome_arquivo_novo"],
        origem_padrao=dados["origem_padrao"],
        aceitavel=dados["aceitavel"],
        nao_aceitavel=dados["nao_aceitavel"],
        conferir=dados["conferir"],
        aba_analises=dados["analises"]["aba_preferida"],
        colunas={
            chave: Coluna(tuple(normalizar(c) for c in col["cabecalho"]), col["letra"])
            for chave, col in dados["analises"]["colunas"].items()
        },
        certificado=LayoutCertificado(
            celulas=cert["celulas"],
            coluna_nome_parametro=cert["coluna_nome_parametro"],
            coluna_resultado=cert["coluna_resultado"],
            coluna_vmp=cert["coluna_vmp"],
            coluna_aceitavel=cert["coluna_aceitavel"],
            linhas_parametros=tuple(cert["linhas_parametros"]),
        ),
        parametros=tuple(
            Parametro(
                p["id"],
                p["nome"],
                tuple(normalizar(n) for n in p["nomes_no_certificado"]),
                p["vmp"],
                Regra(p["regra"]["tipo"], p["regra"].get("minimo"), p["regra"].get("maximo")),
            )
            for p in dados["parametros"]
        ),
    )
