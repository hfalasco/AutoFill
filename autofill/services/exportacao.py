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


def zipar(planilhas: dict[str, bytes], pdfs: dict[str, bytes] | None = None) -> bytes:
    saida = io.BytesIO()
    with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as z:
        for nome, dados in planilhas.items():
            z.writestr(nome, dados)
        for nome, dados in (pdfs or {}).items():
            z.writestr(f"PDF/{nome}", dados)
    return saida.getvalue()
