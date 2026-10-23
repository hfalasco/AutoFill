"""Saída em lote: o .zip com todas as planilhas atualizadas e, se houver LibreOffice, os PDFs."""

from __future__ import annotations

import io
import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from autofill.model.dados import Certificado
from autofill.services.certificado import gerar_xlsx
from autofill.services.xlsx import PastaXlsx

# Vírgula decimal no PDF, como no Excel em português.
LOCALE_PDF = {"LANG": "pt_BR.UTF-8", "LC_ALL": "pt_BR.UTF-8"}


def soffice() -> str | None:
    encontrado = shutil.which("soffice") or shutil.which("libreoffice")
    if encontrado:
        return encontrado
    for pasta in (os.environ.get("ProgramFiles", r"C:\Program Files"), os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
        candidato = Path(pasta) / "LibreOffice" / "program" / "soffice.exe"
        if candidato.exists():
            return str(candidato)
    return None


def pdf_disponivel() -> bool:
    return soffice() is not None


def gerar_pdfs(planilhas: dict[str, tuple[bytes, str]]) -> dict[str, bytes]:
    """{nome do pdf: (xlsx, aba)} -> {nome do pdf: bytes}. Uma única chamada ao LibreOffice para o lote."""
    executavel = soffice()
    if executavel is None or not planilhas:
        return {}
    with tempfile.TemporaryDirectory() as pasta:
        entradas = []
        for i, (nome, (xlsx, aba)) in enumerate(planilhas.items()):
            pasta_xlsx = PastaXlsx(xlsx)
            pasta_xlsx.preparar_impressao(aba)
            entrada = Path(pasta) / f"cert_{i:03d}.xlsx"
            entrada.write_bytes(pasta_xlsx.salvar())
            entradas.append((nome, entrada))
        perfil = Path(pasta) / "perfil"
        try:
            subprocess.run(
                [executavel, f"-env:UserInstallation={perfil.as_uri()}", "--headless", "--convert-to", "pdf",
                 "--outdir", pasta, *(str(e) for _, e in entradas)],
                check=True, capture_output=True, timeout=60 + 20 * len(entradas), env={**os.environ, "HOME": pasta, **LOCALE_PDF},
            )
        except (subprocess.SubprocessError, OSError):
            return {}
        saida = {}
        for nome, entrada in entradas:
            pdf = entrada.with_suffix(".pdf")
            if pdf.exists():
                saida[nome] = pdf.read_bytes()
        return saida


def nome_pdf(cert: Certificado) -> str:
    return f"{Path(cert.arquivo.nome_arquivo).stem} - {cert.aba_destino}.pdf"


def gerar_lote(certificados: list[Certificado], com_pdf: bool = False) -> tuple[dict[str, bytes], dict[str, bytes]]:
    """(planilhas, pdfs) dos certificados prontos, indexados pelo nome do arquivo de saída."""
    planilhas: dict[str, bytes] = {}
    para_pdf: dict[str, tuple[bytes, str]] = {}
    for cert in certificados:
        if not cert.pronto:
            continue
        xlsx = gerar_xlsx(cert)
        planilhas[cert.arquivo.nome_arquivo] = xlsx
        para_pdf[nome_pdf(cert)] = (xlsx, cert.aba_destino)
    pdfs = gerar_pdfs(para_pdf) if com_pdf else {}
    return planilhas, pdfs


EXTENSOES_PLANILHA = (".xlsx", ".xlsm")
MAX_ARQUIVOS_ZIP = 1000
MAX_TAMANHO_ARQUIVO = 50 * 1024 * 1024  # descompactado, por planilha
MAX_TAMANHO_ZIP = 1024 * 1024 * 1024  # descompactado, somando tudo


def e_planilha_xlsx(conteudo: bytes) -> bool:
    """Um .xlsx também é um zip: o que o diferencia é o [Content_Types].xml na raiz."""
    try:
        with zipfile.ZipFile(io.BytesIO(conteudo)) as z:
            return "[Content_Types].xml" in z.namelist()
    except zipfile.BadZipFile:
        return False


def desempacotar(nome: str, conteudo: bytes) -> tuple[list[tuple[str, bytes]], list[str]]:
    """Planilhas de um envio: o próprio arquivo, ou as de dentro de um .zip (ex.: o do mês anterior).

    Volta (lista de (nome, conteúdo), avisos). Dentro do zip valem só .xlsx/.xlsm; PDFs, pastas,
    temporários do Excel (~$...) e lixo do macOS (__MACOSX) são ignorados.
    """
    if not nome.lower().endswith(".zip") or e_planilha_xlsx(conteudo):
        return [(nome, conteudo)], []
    try:
        zip_enviado = zipfile.ZipFile(io.BytesIO(conteudo))
    except zipfile.BadZipFile:
        return [], [f"{nome}: o .zip está corrompido e não pôde ser aberto."]
    planilhas: list[tuple[str, bytes]] = []
    avisos: list[str] = []
    total = 0
    with zip_enviado:
        for info in zip_enviado.infolist():
            base = Path(info.filename).name
            if (
                info.is_dir()
                or info.filename.startswith("__MACOSX/")
                or base.startswith(("~$", "."))
                or not base.lower().endswith(EXTENSOES_PLANILHA)
            ):
                continue
            if len(planilhas) >= MAX_ARQUIVOS_ZIP:
                avisos.append(f"{nome}: mais de {MAX_ARQUIVOS_ZIP} planilhas; o restante foi ignorado.")
                break
            if info.file_size > MAX_TAMANHO_ARQUIVO or total + info.file_size > MAX_TAMANHO_ZIP:
                avisos.append(f"{nome}: {base} é grande demais e foi ignorado.")
                continue
            total += info.file_size
            planilhas.append((base, zip_enviado.read(info)))
    if not planilhas:
        avisos.append(f"{nome}: nenhuma planilha .xlsx dentro do .zip.")
    return planilhas, avisos


def zipar(planilhas: dict[str, bytes], pdfs: dict[str, bytes] | None = None) -> bytes:
    saida = io.BytesIO()
    with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as z:
        for nome, dados in planilhas.items():
            z.writestr(nome, dados)
        for nome, dados in (pdfs or {}).items():
            z.writestr(f"PDF/{nome}", dados)
    return saida.getvalue()
