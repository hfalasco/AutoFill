"""Edição de .xlsx direto no XML do pacote.

O openpyxl descarta caixas de texto, linhas e imagens EMF (a assinatura do responsável técnico) ao
salvar. Por isso o certificado é editado no XML: clona-se a aba do último certificado e só os
valores das células mudam; logotipo, assinatura, bordas e estilos ficam idênticos ao original.
"""

from __future__ import annotations

import io
import posixpath
import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from xml.sax.saxutils import escape

TIPO_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
TIPO_IMAGEM = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
TIPO_EXTERNO = "External"
_EPOCA_EXCEL = date(1899, 12, 30)

_ATRIBUTO = re.compile(r'([\w:]+)="([^"]*)"')


def _atributos(tag: str) -> dict[str, str]:
    return dict(_ATRIBUTO.findall(tag))


def _desescapar(texto: str) -> str:
    return (
        texto.replace("&quot;", '"').replace("&apos;", "'").replace("&lt;", "<").replace("&gt;", ">")
        .replace("&amp;", "&")
    )


def coluna_para_indice(coluna: str) -> int:
    indice = 0
    for letra in coluna.upper():
        indice = indice * 26 + ord(letra) - 64
    return indice


def _separar_ref(ref: str) -> tuple[str, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref.upper())
    if not m:
        raise ValueError(f"Referência de célula inválida: {ref}")
    return m.group(1), int(m.group(2))


def serial_excel(valor: date | datetime) -> float:
    if isinstance(valor, datetime):
        delta = valor - datetime(1899, 12, 30)
        return delta.days + delta.seconds / 86400
    return float((valor - _EPOCA_EXCEL).days)


@dataclass
class Aba:
    nome: str
    sheet_id: int
    rid: str
    caminho: str  # ex.: "xl/worksheets/sheet2.xml"
    oculta: bool


class PastaXlsx:
    """Um .xlsx aberto em memória, editado parte a parte."""

    def __init__(self, conteudo: bytes):
        with zipfile.ZipFile(io.BytesIO(conteudo)) as z:
            self._ordem = [i.filename for i in z.infolist()]
            self._partes = {nome: z.read(nome) for nome in self._ordem}

    # ------------------------------------------------------------------ partes
    def _ler(self, caminho: str) -> str:
        return self._partes[caminho].decode("utf-8")

    def _gravar(self, caminho: str, texto: str | bytes) -> None:
        if caminho not in self._partes:
            self._ordem.append(caminho)
        self._partes[caminho] = texto.encode("utf-8") if isinstance(texto, str) else texto

    @staticmethod
    def _rels_de(caminho: str) -> str:
        pasta, arquivo = posixpath.split(caminho)
        return posixpath.join(pasta, "_rels", arquivo + ".rels")

    @staticmethod
    def _resolver(origem: str, alvo: str) -> str:
        if alvo.startswith("/"):
            return alvo.lstrip("/")
        return posixpath.normpath(posixpath.join(posixpath.dirname(origem), alvo))

    def _relacoes(self, caminho_rels: str) -> list[dict[str, str]]:
        if caminho_rels not in self._partes:
            return []
        return [_atributos(m) for m in re.findall(r"<Relationship\b[^>]*>", self._ler(caminho_rels))]

    def _proximo_nome(self, modelo: str) -> str:
        """sheet2.xml -> primeiro sheetN.xml livre na mesma pasta."""
        pasta, arquivo = posixpath.split(modelo)
        m = re.fullmatch(r"(.*?)(\d*)(\.\w+)", arquivo)
        base, ext = m.group(1), m.group(3)
        n = 1
        while posixpath.join(pasta, f"{base}{n}{ext}") in self._partes:
            n += 1
        return posixpath.join(pasta, f"{base}{n}{ext}")

    def _registrar_tipo(self, original: str, novo: str) -> None:
        tipos = self._ler("[Content_Types].xml")
        m = re.search(rf'<Override PartName="/{re.escape(original)}" ContentType="([^"]+)"\s*/>', tipos)
        if m:
            tipos = tipos.replace("</Types>", f'<Override PartName="/{novo}" ContentType="{m.group(1)}"/></Types>')
            self._gravar("[Content_Types].xml", tipos)

    # ------------------------------------------------------------------ abas
    def abas(self) -> list[Aba]:
        rels = {r["Id"]: r["Target"] for r in self._relacoes("xl/_rels/workbook.xml.rels")}
        abas = []
        for tag in re.findall(r"<sheet\b[^>]*/>", self._ler("xl/workbook.xml")):
            a = _atributos(tag)
            abas.append(
                Aba(
                    _desescapar(a["name"]),
                    int(a["sheetId"]),
                    a["r:id"],
                    self._resolver("xl/workbook.xml", rels[a["r:id"]]),
                    a.get("state", "visible") != "visible",
                )
            )
        return abas

    def aba(self, nome: str) -> Aba:
        for a in self.abas():
            if a.nome == nome:
                return a
        raise KeyError(nome)

    def clonar_aba(self, origem: str, novo_nome: str) -> Aba:
        """Copia a aba `origem` (com desenhos e imagens) logo após ela, com o nome `novo_nome`."""
        if any(a.nome == novo_nome for a in self.abas()):
            raise ValueError(f"Já existe uma aba chamada {novo_nome!r}")
        aba = self.aba(origem)
        novo_caminho = self._proximo_nome(aba.caminho)
        xml = self._ler(aba.caminho)
        xml = re.sub(r'\sxr:uid="\{[^}]*\}"', "", xml, count=1)
        xml = xml.replace(' tabSelected="1"', "")
        self._gravar(novo_caminho, xml)
        self._registrar_tipo(aba.caminho, novo_caminho)
        self._clonar_relacoes(aba.caminho, novo_caminho)

        rels_wb = self._ler("xl/_rels/workbook.xml.rels")
        usados = {int(n) for n in re.findall(r'Id="rId(\d+)"', rels_wb)}
        rid = f"rId{max(usados, default=0) + 1}"
        alvo = posixpath.relpath(novo_caminho, "xl")
        rels_wb = rels_wb.replace(
            "</Relationships>",
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            f'relationships/worksheet" Target="{alvo}"/></Relationships>',
        )
        self._gravar("xl/_rels/workbook.xml.rels", rels_wb)

        wb = self._ler("xl/workbook.xml")
        sheet_id = max(a.sheet_id for a in self.abas()) + 1
        tag_origem = next(t for t in re.findall(r"<sheet\b[^>]*/>", wb) if _atributos(t)["r:id"] == aba.rid)
        nova_tag = f'<sheet name="{escape(novo_nome, {chr(34): "&quot;"})}" sheetId="{sheet_id}" r:id="{rid}"/>'
        wb = wb.replace(tag_origem, tag_origem + nova_tag, 1)
        wb = self._deslocar_indices_locais(wb, self._indice(aba.rid) + 1)
        self._gravar("xl/workbook.xml", wb)
        return self.aba(novo_nome)

    def _indice(self, rid: str) -> int:
        return next(i for i, a in enumerate(self.abas()) if a.rid == rid)

    @staticmethod
    def _deslocar_indices_locais(wb: str, a_partir: int) -> str:
        """Nomes definidos por aba (localSheetId) e a aba ativa usam a posição: abrir espaço."""

        def ajustar(m: re.Match) -> str:
            n = int(m.group(2))
            return f'{m.group(1)}="{n + 1 if n >= a_partir else n}"'

        return re.sub(r'(localSheetId|activeTab|firstSheet)="(\d+)"', ajustar, wb)

    def _clonar_relacoes(self, original: str, novo: str) -> None:
        rels_original = self._rels_de(original)
        if rels_original not in self._partes:
            return
        texto = self._ler(rels_original)
        for rel in self._relacoes(rels_original):
            # Imagens podem ser compartilhadas; desenhos, configurações de impressão etc. não.
            if rel.get("TargetMode") == TIPO_EXTERNO or rel["Type"] == TIPO_IMAGEM:
                continue
            parte = self._resolver(original, rel["Target"])
            if parte not in self._partes:
                continue
            copia = self._proximo_nome(parte)
            self._gravar(copia, self._partes[parte])
            self._registrar_tipo(parte, copia)
            if rel["Type"] == TIPO_DRAWING:
                self._clonar_relacoes(parte, copia)
            novo_alvo = posixpath.relpath(copia, posixpath.dirname(novo))
            texto = texto.replace(f'Target="{rel["Target"]}"', f'Target="{novo_alvo}"', 1)
        self._gravar(self._rels_de(novo), texto)

    def renomear_aba(self, nome: str, novo_nome: str) -> None:
        if nome == novo_nome:
            return
        if any(a.nome == novo_nome for a in self.abas()):
            raise ValueError(f"Já existe uma aba chamada {novo_nome!r}")
        aba = self.aba(nome)
        wb = self._ler("xl/workbook.xml")
        tag = next(t for t in re.findall(r"<sheet\b[^>]*/>", wb) if _atributos(t)["r:id"] == aba.rid)
        nova = re.sub(r'name="[^"]*"', f'name="{escape(novo_nome, {chr(34): "&quot;"})}"', tag, count=1)
        self._gravar("xl/workbook.xml", wb.replace(tag, nova, 1))

    def ativar_aba(self, nome: str) -> None:
        """Deixa `nome` como a aba aberta ao abrir o arquivo."""
        alvo = self.aba(nome)
        for a in self.abas():
            xml = self._ler(a.caminho).replace(' tabSelected="1"', "")
            if a.rid == alvo.rid:
                xml = re.sub(r"<sheetView\b", '<sheetView tabSelected="1"', xml, count=1)
            self._gravar(a.caminho, xml)
        wb = self._ler("xl/workbook.xml")
        indice = self._indice(alvo.rid)
        if "activeTab=" in wb:
            wb = re.sub(r'activeTab="\d+"', f'activeTab="{indice}"', wb)
        else:
            wb = re.sub(r"<workbookView\b", f'<workbookView activeTab="{indice}"', wb, count=1)
        wb = re.sub(r'\sfirstSheet="\d+"', "", wb)
        self._gravar("xl/workbook.xml", wb)

    def preparar_impressao(self, nome: str) -> None:
        """Só `nome` fica visível e cabe em uma página (usado para gerar o PDF)."""
        alvo = self.aba(nome)
        wb = self._ler("xl/workbook.xml")
        for tag in re.findall(r"<sheet\b[^>]*/>", wb):
            a = _atributos(tag)
            nova = re.sub(r'\sstate="[^"]*"', "", tag)
            if a["r:id"] != alvo.rid:
                nova = nova.replace("<sheet ", '<sheet state="hidden" ', 1)
            wb = wb.replace(tag, nova, 1)
        self._gravar("xl/workbook.xml", wb)
        self.ativar_aba(nome)

        xml = self._ler(alvo.caminho)
        if "<pageSetUpPr" in xml:
            xml = re.sub(r"<pageSetUpPr\b[^>]*/>", '<pageSetUpPr fitToPage="1"/>', xml)
        elif re.search(r"<sheetPr(?:\s[^>]*)?/>", xml):
            xml = re.sub(r"<sheetPr((?:\s[^>]*)?)/>", r'<sheetPr\1><pageSetUpPr fitToPage="1"/></sheetPr>', xml, count=1)
        elif re.search(r"<sheetPr[\s>]", xml):  # cuidado: "<sheetProtection" também começa com "<sheetPr"
            xml = re.sub(r"(<sheetPr(?:\s[^>]*)?>)", r'\1<pageSetUpPr fitToPage="1"/>', xml, count=1)
        else:
            xml = re.sub(r"(<worksheet\b[^>]*>)", r'\1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>', xml, count=1)
        if "<pageSetup" in xml:
            xml = re.sub(r'\s(fitToWidth|fitToHeight|scale)="\d+"', "", xml)
            xml = xml.replace("<pageSetup ", '<pageSetup fitToWidth="1" fitToHeight="1" ', 1)
        self._gravar(alvo.caminho, xml)
        self._definir_area_impressao(alvo)

    def _definir_area_impressao(self, aba: Aba) -> None:
        """Área de impressão = células usadas + imagens (a assinatura do modelo passa da última coluna)."""
        xml = self._ler(aba.caminho)
        m = re.search(r'<dimension ref="[A-Z]+\d+:([A-Z]+)(\d+)"', xml)
        if not m:
            return
        ultima_col, ultima_lin = coluna_para_indice(m.group(1)), int(m.group(2))
        for rel in self._relacoes(self._rels_de(aba.caminho)):
            if rel["Type"] != TIPO_DRAWING:
                continue
            desenho = self._ler(self._resolver(aba.caminho, rel["Target"]))
            for col, lin in re.findall(r"<xdr:to><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>", desenho, re.S):
                ultima_col, ultima_lin = max(ultima_col, int(col) + 1), max(ultima_lin, int(lin) + 1)
        letra = ""
        n = ultima_col
        while n:
            n, resto = divmod(n - 1, 26)
            letra = chr(65 + resto) + letra
        nome = aba.nome.replace("'", "''")
        indice = self._indice(aba.rid)
        definido = (
            f'<definedName name="_xlnm.Print_Area" localSheetId="{indice}">'
            f"'{escape(nome)}'!$A$1:${letra}${ultima_lin}</definedName>"
        )
        wb = self._ler("xl/workbook.xml")
        wb = re.sub(rf'<definedName name="_xlnm.Print_Area" localSheetId="{indice}">.*?</definedName>', "", wb)
        if "<definedNames>" in wb:
            wb = wb.replace("<definedNames>", "<definedNames>" + definido, 1)
        else:
            wb = wb.replace("</sheets>", f"</sheets><definedNames>{definido}</definedNames>", 1)
        self._gravar("xl/workbook.xml", wb)

    # ------------------------------------------------------------------ células
    def _indice_string(self, texto: str) -> int | None:
        caminho = "xl/sharedStrings.xml"
        if caminho not in self._partes:
            return None
        sst = self._ler(caminho)
        indice = len(re.findall(r"<si>|<si\s", sst))
        sst = sst.replace("</sst>", f'<si><t xml:space="preserve">{escape(texto)}</t></si></sst>', 1)

        def incrementar(m: re.Match) -> str:
            return f'{m.group(1)}="{int(m.group(2)) + 1}"'

        sst = re.sub(r'\b(count|uniqueCount)="(\d+)"', incrementar, sst)
        self._gravar(caminho, sst)
        return indice

    def _conteudo(self, valor) -> tuple[str, str]:
        """(atributo t, XML interno) da célula para `valor`."""
        if valor is None or valor == "":
            return "", ""
        if isinstance(valor, bool):
            return ' t="b"', f"<v>{int(valor)}</v>"
        if isinstance(valor, (date, datetime)):
            numero = serial_excel(valor)
            return "", f"<v>{int(numero) if numero == int(numero) else numero}</v>"
        if isinstance(valor, (int, float)):
            return "", f"<v>{repr(valor) if isinstance(valor, float) else valor}</v>"
        texto = str(valor)
        indice = self._indice_string(texto)
        if indice is None:
            return ' t="inlineStr"', f'<is><t xml:space="preserve">{escape(texto)}</t></is>'
        return ' t="s"', f"<v>{indice}</v>"

    def definir_celulas(self, nome_aba: str, valores: dict[str, object]) -> None:
        aba = self.aba(nome_aba)
        xml = self._ler(aba.caminho)
        for ref, valor in valores.items():
            xml = self._definir(xml, ref.upper(), *self._conteudo(valor))
        self._gravar(aba.caminho, xml)

    @staticmethod
    def _definir(xml: str, ref: str, atributo_t: str, interno: str) -> str:
        padrao = re.compile(rf'<c r="{ref}"(?=[\s/>])([^>]*?)(?:/>|>(.*?)</c>)', re.S)
        m = padrao.search(xml)
        if m:
            atributos = re.sub(r'\st="[^"]*"', "", m.group(1))
            nova = f'<c r="{ref}"{atributos}{atributo_t}>{interno}</c>' if interno else f'<c r="{ref}"{atributos}/>'
            return xml[: m.start()] + nova + xml[m.end():]

        coluna, linha = _separar_ref(ref)
        nova = f'<c r="{ref}"{atributo_t}>{interno}</c>' if interno else ""
        if not nova:
            return xml
        m_linha = re.search(rf'<row r="{linha}"(?=[\s>])[^>]*?(?:/>|>(.*?)</row>)', xml, re.S)
        if m_linha is None:
            # Linha inexistente: entra antes da primeira linha de número maior.
            nova_linha = f'<row r="{linha}">{nova}</row>'
            for m_outra in re.finditer(r'<row r="(\d+)"', xml):
                if int(m_outra.group(1)) > linha:
                    return xml[: m_outra.start()] + nova_linha + xml[m_outra.start():]
            if "<sheetData/>" in xml:
                return xml.replace("<sheetData/>", f"<sheetData>{nova_linha}</sheetData>", 1)
            return xml.replace("</sheetData>", nova_linha + "</sheetData>", 1)
        texto_linha = m_linha.group(0)
        if texto_linha.endswith("/>"):
            nova_linha = texto_linha[:-2] + f">{nova}</row>"
        else:
            nova_linha = None
            for m_c in re.finditer(r'<c r="([A-Z]+)\d+"', texto_linha):
                if coluna_para_indice(m_c.group(1)) > coluna_para_indice(coluna):
                    nova_linha = texto_linha[: m_c.start()] + nova + texto_linha[m_c.start():]
                    break
            if nova_linha is None:
                nova_linha = texto_linha[: -len("</row>")] + nova + "</row>"
        return xml[: m_linha.start()] + nova_linha + xml[m_linha.end():]

    def formatar_numero(self, nome_aba: str, refs: list[str], formato: str) -> None:
        """Aplica o formato numérico `formato` (ex.: 'dd/mm/yyyy') às células, mantendo o resto do estilo.

        O modelo usa o formato de data embutido nº 14, que muda com o idioma de quem abre
        (10/7/2026 no LibreOffice em inglês); um formato explícito sai igual em qualquer máquina.
        """
        if "xl/styles.xml" not in self._partes:
            return
        estilos = self._ler("xl/styles.xml")
        existente = re.search(rf'<numFmt numFmtId="(\d+)" formatCode="{re.escape(escape(formato))}"\s*/>', estilos)
        if existente:
            fmt_id = int(existente.group(1))
        else:
            fmt_id = max([163, *(int(n) for n in re.findall(r'<numFmt numFmtId="(\d+)"', estilos))]) + 1
            novo = f'<numFmt numFmtId="{fmt_id}" formatCode="{escape(formato)}"/>'
            if "<numFmts" in estilos:
                estilos = re.sub(r"(<numFmts\b[^>]*>)", r"\1" + novo, estilos, count=1)
                estilos = re.sub(
                    r'(<numFmts\b[^>]*count=")(\d+)"', lambda m: f'{m.group(1)}{int(m.group(2)) + 1}"', estilos, count=1
                )
            else:
                estilos = re.sub(r"(<styleSheet\b[^>]*>)", rf'\1<numFmts count="1">{novo}</numFmts>', estilos, count=1)

        m_xfs = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", estilos, re.S)
        xfs = re.findall(r"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>", m_xfs.group(1), re.S)
        aba = self.aba(nome_aba)
        xml = self._ler(aba.caminho)
        novos: dict[int, int] = {}
        for ref in refs:
            m = re.search(rf'<c r="{ref}"(?=[\s/>])[^>]*?\ss="(\d+)"', xml)
            estilo = int(m.group(1)) if m else 0
            if estilo not in novos and f'numFmtId="{fmt_id}"' in xfs[estilo]:
                novos[estilo] = estilo  # já formatada numa execução anterior
            if estilo not in novos:
                xf = re.sub(r'\snumFmtId="\d+"', "", xfs[estilo])
                xf = re.sub(r'\sapplyNumberFormat="\d"', "", xf)
                xf = xf.replace("<xf ", f'<xf numFmtId="{fmt_id}" applyNumberFormat="1" ', 1)
                novos[estilo] = len(xfs)
                xfs.append(xf)
            if m:
                inicio, fim = m.span(1)
                xml = xml[:inicio] + str(novos[estilo]) + xml[fim:]
        corpo = "".join(xfs)
        estilos = (
            estilos[: m_xfs.start()]
            + re.sub(r'count="\d+"', f'count="{len(xfs)}"', m_xfs.group(0)[: m_xfs.group(0).index(">") + 1], count=1)
            + corpo
            + "</cellXfs>"
            + estilos[m_xfs.end():]
        )
        self._gravar("xl/styles.xml", estilos)
        self._gravar(aba.caminho, xml)

    # ------------------------------------------------------------------ saída
    def salvar(self) -> bytes:
        # A lista de nomes de abas em docProps/app.xml é opcional; removê-la evita que fique
        # desatualizada depois de clonar ou renomear abas.
        if "docProps/app.xml" in self._partes:
            app = self._ler("docProps/app.xml")
            app = re.sub(r"<HeadingPairs>.*?</HeadingPairs>|<TitlesOfParts>.*?</TitlesOfParts>", "", app, flags=re.S)
            self._gravar("docProps/app.xml", app)
        saida = io.BytesIO()
        with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as z:
            for nome in self._ordem:
                z.writestr(nome, self._partes[nome])
        return saida.getvalue()
