"""Componentes de tela (somente apresentação; nenhuma regra de negócio aqui)."""

from __future__ import annotations

import html

import pandas as pd
import streamlit as st

from autofill.model.config import Config
from autofill.model.dados import Amostra, Certificado

SEM_COLETA = "⛔ Sem coleta"
ICONES = {"Aceitável": "✅", "Não aceitável": "❌", "Conferir": "⚠️"}
CORES = {
    "Aceitável": "background-color: #E3F4EA; color: #14603A; font-weight: 600",
    "Não aceitável": "background-color: #FBE7E4; color: #A93226; font-weight: 600",
    "Conferir": "background-color: #FFF6DA; color: #8A6D00; font-weight: 600",
}


def card(rotulo: str, valor: str, detalhe: str = "", estilo: str = "", grande: bool = False) -> None:
    st.markdown(
        f"""<div class="smto-card {estilo}">
              <div class="rotulo">{html.escape(rotulo)}</div>
              <div class="valor {'grande' if grande else ''}">{html.escape(valor)}</div>
              <div class="detalhe">{html.escape(detalhe)}</div>
            </div>""",
        unsafe_allow_html=True,
    )


def titulo(texto: str, subtitulo: str = "") -> None:
    st.markdown(f'<div class="smto-titulo">{html.escape(texto)}</div>', unsafe_allow_html=True)
    if subtitulo:
        st.markdown(f'<div class="smto-subtitulo">{html.escape(subtitulo)}</div>', unsafe_allow_html=True)


def estado_vazio() -> None:
    st.markdown(
        """<div class="smto-vazio">
             <h3>Carregue a planilha de análises</h3>
             Em três passos, pela barra lateral:
             <ol>
               <li><b>FO.GSM.LMI.03.08</b> - Análises Microbiológicas de Água Potável;</li>
               <li>as planilhas dos pontos (<i>Vivência 417.xlsx</i>, <i>Vivência 631.xlsx</i>...) e/ou os
                   números dos pontos que ainda não têm planilha;</li>
               <li>o mês do certificado.</li>
             </ol><br>
             A ferramenta acha a coleta de cada ponto no mês, cria a aba nova do certificado
             (FO.GSM.LMI.07.01) com resultados, validade e “Aceitável”, e entrega tudo num único .zip.
           </div>""",
        unsafe_allow_html=True,
    )


def situacao(cert: Certificado) -> str:
    if not cert.pronto:
        return SEM_COLETA
    situacoes = {l.situacao for l in cert.linhas}
    for s in ("Não aceitável", "Conferir"):
        if s in situacoes:
            return f"{ICONES[s]} {s}"
    return f"{ICONES['Aceitável']} Aceitável"


def cards_resumo(certs: list[Certificado], total_coletas: int, periodo: str) -> None:
    prontos = [c for c in certs if c.pronto]
    pendentes = [c for c in certs if not c.aceitavel]
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        card("Período", periodo, f"{total_coletas} coletas lidas na planilha de análises", "azul")
    with c2:
        card("Certificados prontos", f"{len(prontos)} de {len(certs)}", "Com coleta encontrada no mês", grande=True)
    with c3:
        novos = sum(1 for c in prontos if c.arquivo.novo)
        card("Arquivos", f"{len(prontos) - novos} atualizado(s)", f"{novos} criado(s) a partir do modelo")
    with c4:
        card("Para conferir", str(len(pendentes)), "Sem coleta, não aceitável ou valor ilegível", "amarelo")


def _formatar(valor) -> str:
    if valor is None:
        return "—"
    if isinstance(valor, float):
        return f"{valor:g}".replace(".", ",")
    return str(valor)


def descricao_amostra(a: Amostra, config: Config) -> str:
    partes = [f"{a.data:%d/%m/%Y}", f"linha {a.linha}"]
    partes += [f"{config.parametro(k).nome.split()[0]} {_formatar(v)}" for k, v in a.resultados.items() if v is not None]
    return " · ".join(partes)


def tabela_conferencia(certs: list[Certificado], config: Config) -> pd.DataFrame:
    linhas = []
    for c in certs:
        valores = {l.parametro_id: _formatar(l.valor) for l in c.linhas}
        linhas.append(
            {
                "Gerar": c.pronto,
                "Ponto": f"{c.arquivo.origem} {c.arquivo.identificacao}",
                "Arquivo": c.arquivo.nome_arquivo + (" (novo)" if c.arquivo.novo else ""),
                "Aba": c.aba_destino + (" (atualizar)" if c.sobrescreve else ""),
                "Nº": c.numero,
                "Coleta": c.amostra.data if c.pronto else pd.NaT,
                "Validade": c.validade or pd.NaT,
                "Coletor": c.coletor,
                **{p.nome: valores.get(p.id, "—") for p in config.parametros},
                "Situação": situacao(c),
                "Avisos": " ".join(c.avisos),
            }
        )
    df = pd.DataFrame(linhas)
    for coluna in ("Coleta", "Validade"):  # datetime de verdade: célula vazia em vez de "None"
        df[coluna] = pd.to_datetime(df[coluna])
    return df


def editor_conferencia(df: pd.DataFrame, config: Config, chave: str) -> pd.DataFrame:
    bloqueadas = [col for col in df.columns if col not in ("Gerar", "Nº", "Coletor")]
    return st.data_editor(
        df,
        key=chave,
        hide_index=True,
        width="stretch",
        disabled=bloqueadas,
        column_config={
            "Gerar": st.column_config.CheckboxColumn(width="small", help="Desmarque para deixar o ponto fora do .zip"),
            "Nº": st.column_config.NumberColumn("Nº cert.", min_value=1, step=1, format="%d", width="small"),
            "Coleta": st.column_config.DateColumn(format="DD/MM/YYYY", width="small"),
            "Validade": st.column_config.DateColumn(format="DD/MM/YYYY", width="small"),
            "Coletor": st.column_config.TextColumn(width="medium", help="Vem do último certificado do ponto"),
            "Avisos": st.column_config.TextColumn(width="large"),
        },
    )


def previa_certificado(cert: Certificado) -> None:
    info = cert.arquivo.ultimo
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"**Mês / Ano**  \n{cert.mes} / {cert.ano}")
    c2.markdown(f"**Nº certificado**  \n{cert.numero}")
    c3.markdown(f"**Data da amostragem**  \n{cert.amostra.data:%d/%m/%Y}")
    c4.markdown(f"**Validade**  \n{cert.validade:%d/%m/%Y}")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f"**Local**  \n{info.local or '—'}")
    c2.markdown(f"**Ponto**  \n{info.ponto or '—'}")
    c3.markdown(f"**Coletor**  \n{cert.coletor or '—'}")
    a = cert.amostra
    c4.markdown(f"**Na planilha de análises**  \nlinha {a.linha} · {a.area or '—'} / {a.setor or '—'}")
    df = pd.DataFrame(
        [{"Parâmetro": l.nome, "Resultado": _formatar(l.valor), "*VMP": l.vmp, "Aceitável": l.situacao} for l in cert.linhas]
    )
    st.dataframe(
        df.style.map(lambda v: CORES.get(v, ""), subset=["Aceitável"]),
        hide_index=True,
        width="stretch",
    )


def tabela_coletas_do_mes(amostras: list[Amostra], config: Config) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Data": a.data,
                "Origem": a.origem,
                "Identificação": a.identificacao,
                "Área": a.area,
                "Setor": a.setor,
                **{p.nome: _formatar(a.resultados.get(p.id)) for p in config.parametros},
                "Linha": a.linha,
            }
            for a in amostras
        ]
    )
