"""Objetos do domínio: a amostra lida no FO.GSM.LMI.03.08 e o certificado a preencher."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from autofill.model.config import carregar_config, normalizar


def chave_ponto(origem: object, identificacao: object) -> tuple[str, str]:
    """('area de vivencia', '417'): 417, '417' e '417.0' são o mesmo ponto."""
    ident = identificacao
    if isinstance(ident, float) and ident.is_integer():
        ident = int(ident)
    return normalizar(origem), normalizar(ident)


@dataclass(frozen=True)
class Amostra:
    linha: int  # linha no FO.GSM.LMI.03.08, para o usuário conferir
    data: date
    origem: str
    identificacao: str
    area: str
    setor: str
    resultados: dict[str, object]  # id do parâmetro -> valor como está na planilha
    rubrica: str = ""

    @property
    def chave(self) -> tuple[str, str]:
        return chave_ponto(self.origem, self.identificacao)


@dataclass(frozen=True)
class InfoCertificado:
    """O que se lê de um certificado existente (uma aba de 'Vivência NNN.xlsx')."""

    aba: str
    mes: str
    ano: int | None
    numero: int | None
    versao: object
    local: str
    veiculo: str
    ponto: str
    coletor: str
    linhas_parametros: dict[str, int]  # id do parâmetro -> linha da planilha


@dataclass
class ArquivoCertificado:
    """Um arquivo de certificados de um ponto (existente ou novo, criado a partir do modelo)."""

    nome_arquivo: str
    conteudo: bytes
    origem: str
    identificacao: str
    abas: list[InfoCertificado]
    novo: bool = False  # True quando vem do modelo (não havia planilha do ponto)

    @property
    def chave(self) -> tuple[str, str]:
        return chave_ponto(self.origem, self.identificacao)

    @property
    def ultimo(self) -> InfoCertificado:
        return self.abas[-1]


@dataclass
class LinhaResultado:
    parametro_id: str
    nome: str
    valor: object
    vmp: str
    situacao: str  # "Aceitável", "Não aceitável" ou "Conferir"


@dataclass
class Certificado:
    """Tudo o que vai para a nova aba do certificado."""

    arquivo: ArquivoCertificado
    mes: str
    ano: int
    numero: int
    aba_destino: str
    sobrescreve: bool  # True quando a aba do mesmo mês/ano já existe e será atualizada
    amostra: Amostra | None
    candidatas: list[Amostra] = field(default_factory=list)
    coletor: str = ""
    validade: date | None = None
    linhas: list[LinhaResultado] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)

    @property
    def pronto(self) -> bool:
        return self.amostra is not None

    @property
    def aceitavel(self) -> bool:
        aceitavel = carregar_config().aceitavel
        return self.pronto and all(l.situacao == aceitavel for l in self.linhas)
