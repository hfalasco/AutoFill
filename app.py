"""Certificado Analítico de Água Potável - São Martinho.

Controlador Streamlit: lê o FO.GSM.LMI.03.08 (model/services), monta os certificados
FO.GSM.LMI.07.01 de cada ponto e entrega as planilhas atualizadas (view).
Execução: streamlit run app.py
"""

from __future__ import annotations

import hashlib
from datetime import date

import pandas as pd
import streamlit as st

from autofill.model.config import MESES, carregar_config, normalizar
from autofill.services import certificado as cert_svc
from autofill.services import exportacao
from autofill.services.analises import PlanilhaInvalida, ler_analises
from autofill.services.certificado import CertificadoInvalido
from autofill.view import componentes as ui
from autofill.view import tema

st.set_page_config(
    page_title="Certificado de Água Potável | São Martinho",
    page_icon=str(tema.CAMINHO_LOGO),
    layout="wide",
    initial_sidebar_state="expanded",
)
tema.aplicar()
config = carregar_config()


@st.cache_data(show_spinner=False)
def _analises(conteudo: bytes):
    return ler_analises(conteudo)


@st.cache_data(show_spinner=False)
def _arquivo(conteudo: bytes, nome: str):
    return cert_svc.ler_arquivo(conteudo, nome)


@st.cache_data(show_spinner=False)
def _novo(identificacao: str, origem: str):
    return cert_svc.novo_arquivo(identificacao, origem)


@st.cache_data(show_spinner=False)
def _pdfs(planilhas: dict[str, tuple[bytes, str]]) -> dict[str, bytes]:
    return exportacao.gerar_pdfs(planilhas)


def _periodo_padrao(analises) -> date:
    """Mês da coleta mais recente que não está no futuro (protege de datas digitadas erradas)."""
    hoje = date.today()
    datas = [a.data for a in analises.amostras if a.data <= hoje] if analises else []
    return max(datas, default=hoje)


# ----------------------------------------------------------------------------- barra lateral
with st.sidebar:
    st.image(str(tema.CAMINHO_LOGO), width="stretch")
    st.markdown("### 1. Planilha de análises")
    arquivo_analises = st.file_uploader(
        "FO.GSM.LMI.03.08 - Análises Microbiológicas de Água Potável",
        type=["xlsx", "xlsm"],
        help="A planilha da safra com todas as coletas (aba 'Análises').",
    )

analises = None
erro_analises = None
if arquivo_analises is not None:
    try:
        with st.spinner("Lendo a planilha de análises..."):
            analises = _analises(arquivo_analises.getvalue())
    except PlanilhaInvalida as erro:
        erro_analises = str(erro)

origens = sorted({a.origem for a in analises.amostras if a.origem}, key=normalizar) if analises else []
if config.origem_padrao not in origens:
    origens.insert(0, config.origem_padrao)
padrao = _periodo_padrao(analises)

with st.sidebar:
    st.markdown("### 2. Certificados dos pontos")
    arquivos_pontos = st.file_uploader(
        "Planilhas dos pontos (ex.: Vivência 417.xlsx) ou o .zip baixado no mês anterior",
        type=["xlsx", "xlsm", "zip"],
        accept_multiple_files=True,
        help="Cada planilha ganha uma aba nova com o certificado do mês, copiada da última aba. "
        "Pode enviar várias de uma vez ou o .zip inteiro que a ferramenta gerou da última vez.",
    )
    texto_pontos = st.text_area(
        "Pontos sem planilha (criados do modelo)",
        key="pontos_novos",
        placeholder="631, 634, 665  ou  600-610",
        height=80,
        help="Números separados por vírgula ou faixas. Ponto que já veio em planilha acima é ignorado aqui.",
    )
    origem_nova = st.selectbox("Origem desses pontos", origens, index=origens.index(config.origem_padrao))
    coletor_novo = st.text_input(
        "Coletor responsável desses pontos",
        value=_novo("0", config.origem_padrao).ultimo.coletor,
        help="Os pontos com planilha mantêm o coletor do último certificado; dá para trocar na conferência.",
    )
    st.markdown("### 3. Período do certificado")
    c1, c2 = st.columns([3, 2])
    mes = c1.selectbox("Mês", range(1, 13), index=padrao.month - 1, format_func=lambda m: MESES[m - 1])
    ano = c2.number_input("Ano", min_value=2000, max_value=2100, value=padrao.year, step=1)
    usar_anterior = st.checkbox(
        "Sem coleta no mês: usar a mais recente anterior",
        help="Desmarcado, o ponto sem coleta no mês fica de fora e aparece como pendência.",
    )
    if exportacao.pdf_disponivel():
        com_pdf = st.checkbox("Gerar também o PDF de cada certificado", value=True)
    else:
        com_pdf = False
        st.caption("⚠️ LibreOffice não instalado: só Excel (sem PDF).")
    st.markdown(
        '<div class="smto-rodape">Origem: FO.GSM.LMI.03.08 → Destino: FO.GSM.LMI.07.01<br>'
        "Preenchimento 100% determinístico (regras em data/certificado.json).</div>",
        unsafe_allow_html=True,
    )

tema.cabecalho(
    "Certificado Analítico de Água Potável",
    "Preenche o FO.GSM.LMI.07.01 de cada ponto a partir das Análises Microbiológicas (FO.GSM.LMI.03.08)",
)

if erro_analises:
    st.error(erro_analises)
    st.stop()
if analises is None:
    ui.estado_vazio()
    st.stop()

ano = int(ano)
nome_periodo = f"{MESES[mes - 1]}/{ano}"
if analises.ignoradas:
    st.caption(
        f"⚠️ {len(analises.ignoradas)} linha(s) da aba '{analises.aba}' sem data legível foram ignoradas: "
        + ", ".join(str(n) for n in analises.ignoradas[:15])
        + ("..." if len(analises.ignoradas) > 15 else "")
    )

# ----------------------------------------------------------------------------- arquivos dos pontos
arquivos = []
chaves_vistas: dict[tuple[str, str], str] = {}
planilhas_enviadas: list[tuple[str, bytes]] = []
for enviado in arquivos_pontos or []:
    extraidas, avisos_zip = exportacao.desempacotar(enviado.name, enviado.getvalue())
    planilhas_enviadas += extraidas
    for aviso in avisos_zip:
        st.warning(aviso)
for nome_enviado, conteudo_enviado in planilhas_enviadas:
    try:
        arq = _arquivo(conteudo_enviado, nome_enviado)
    except CertificadoInvalido as erro:
        st.warning(str(erro))
        continue
    if arq.chave in chaves_vistas:
        st.warning(f"{nome_enviado}: mesmo ponto de {chaves_vistas[arq.chave]} ({arq.origem} {arq.identificacao}); ignorado.")
        continue
    chaves_vistas[arq.chave] = nome_enviado
    arquivos.append(arq)

for ponto in cert_svc.ler_lista_pontos(texto_pontos):
    novo = _novo(ponto, origem_nova)
    if novo.chave not in chaves_vistas:
        chaves_vistas[novo.chave] = novo.nome_arquivo
        arquivos.append(novo)

coletas_mes = sorted(
    (a for a in analises.amostras if (a.data.year, a.data.month) == (ano, mes)), key=lambda a: (a.data, a.linha)
)


def _adicionar_pontos_do_mes() -> None:
    ja = {chave for chave in chaves_vistas}
    faltam = []
    for a in coletas_mes:
        if normalizar(a.origem) == normalizar(origem_nova) and a.chave not in ja and a.identificacao not in faltam:
            faltam.append(a.identificacao)
    atuais = cert_svc.ler_lista_pontos(st.session_state.get("pontos_novos", ""))
    st.session_state["pontos_novos"] = ", ".join(atuais + [p for p in faltam if p not in atuais])


if not arquivos:
    ui.titulo(
        f"Coletas de {nome_periodo}",
        f"{len(coletas_mes)} coletas no mês. Envie as planilhas dos pontos ou digite os números na barra lateral.",
    )
    if coletas_mes:
        st.button(
            f"Criar certificado para todos os pontos de '{origem_nova}' coletados no mês",
            on_click=_adicionar_pontos_do_mes,
            type="primary",
        )
    st.dataframe(ui.tabela_coletas_do_mes(coletas_mes, config), hide_index=True, width="stretch", height=480,
                 column_config={"Data": st.column_config.DateColumn(format="DD/MM/YYYY")})
    st.stop()

# ----------------------------------------------------------------------------- planejamento
certs = [cert_svc.planejar(arq, analises, mes, ano, usar_anterior, config) for arq in arquivos]
for c in certs:
    if c.arquivo.novo:
        c.coletor = coletor_novo.strip()

varias = [c for c in certs if len(c.candidatas) > 1]
if varias:
    with st.expander(f"🔁 {len(varias)} ponto(s) com mais de uma coleta no mês (recoleta?) - escolha qual vai para o certificado"):
        st.caption("Por padrão vale a coleta mais recente do mês, que costuma ser a recoleta depois de um resultado 'Presente'.")
        for c in varias:
            escolhida = st.selectbox(
                f"{c.arquivo.origem} {c.arquivo.identificacao}",
                c.candidatas,
                index=len(c.candidatas) - 1,
                format_func=lambda a: ui.descricao_amostra(a, config),
                key=f"coleta_{c.arquivo.nome_arquivo}_{mes}_{ano}",
            )
            if escolhida is not c.amostra:
                cert_svc.escolher_amostra(c, escolhida, config)

ui.cards_resumo(certs, len(analises.amostras), nome_periodo)

# ----------------------------------------------------------------------------- conferência
ui.titulo(
    "1. Conferência",
    "Cada linha é um certificado. Dá para ajustar o Nº, o coletor e desmarcar o que não deve sair.",
)
assinatura = hashlib.sha1(
    "|".join([f"{mes}/{ano}/{usar_anterior}"] + [f"{c.arquivo.nome_arquivo}:{c.amostra and c.amostra.linha}" for c in certs]).encode()
).hexdigest()[:12]
editado = ui.editor_conferencia(ui.tabela_conferencia(certs, config), config, f"conferencia_{assinatura}")

selecionados = []
for c, linha in zip(certs, editado.to_dict("records")):
    if pd.notna(linha.get("Nº")):
        c.numero = int(linha["Nº"])
    c.coletor = (linha.get("Coletor") or "").strip()
    if linha.get("Gerar") and c.pronto:
        selecionados.append(c)

sem_coleta = [c for c in certs if not c.pronto]
if sem_coleta:
    st.info(
        f"{len(sem_coleta)} ponto(s) sem coleta em {nome_periodo}: "
        + ", ".join(c.arquivo.identificacao for c in sem_coleta)
        + ". Confira o mês ou marque “usar a mais recente anterior”."
    )

# ----------------------------------------------------------------------------- prévia
prontos = [c for c in certs if c.pronto]
if prontos:
    ui.titulo("2. Prévia do certificado")
    escolhido = st.selectbox(
        "Ponto", prontos, format_func=lambda c: f"{c.arquivo.origem} {c.arquivo.identificacao} · aba {c.aba_destino}"
    )
    ui.previa_certificado(escolhido)

# ----------------------------------------------------------------------------- download
ui.titulo("3. Baixar", f"{len(selecionados)} certificado(s) selecionado(s)")
if not selecionados:
    st.warning("Nenhum certificado selecionado para gerar.")
    st.stop()

planilhas, _ = exportacao.gerar_lote(selecionados)
pdfs = {}
if com_pdf:
    with st.spinner(f"Gerando {len(selecionados)} PDF(s)..."):
        pdfs = _pdfs({exportacao.nome_pdf(c): (planilhas[c.arquivo.nome_arquivo], c.aba_destino) for c in selecionados})
    if len(pdfs) < len(selecionados):
        st.caption("⚠️ Não foi possível gerar todos os PDFs; as planilhas Excel estão completas.")

pacote = exportacao.zipar(planilhas, pdfs)
st.download_button(
    f"Baixar tudo (.zip) - {len(planilhas)} planilha(s)" + (f" + {len(pdfs)} PDF(s)" if pdfs else ""),
    pacote,
    f"Certificados_{MESES[mes - 1]}_{ano}.zip",
    "application/zip",
    type="primary",
    width="stretch",
)
with st.expander("Baixar um arquivo por vez"):
    for c in selecionados:
        nome = c.arquivo.nome_arquivo
        c1, c2 = st.columns(2)
        c1.download_button(
            f"📗 {nome}", planilhas[nome], nome,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"x_{nome}", width="stretch",
        )
        nome_pdf = exportacao.nome_pdf(c)
        if nome_pdf in pdfs:
            c2.download_button(f"📕 {nome_pdf}", pdfs[nome_pdf], nome_pdf, "application/pdf", key=f"p_{nome}", width="stretch")
st.caption(
    "Cada planilha volta com o mesmo nome e uma aba nova (ou a do mês atualizada), já aberta nela. "
    "As abas anteriores, o logotipo e a assinatura ficam intactos."
)
