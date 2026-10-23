"""Identidade visual São Martinho (cores e elementos do template institucional)."""

from __future__ import annotations

import base64
from functools import lru_cache
from pathlib import Path

import streamlit as st

VERDE = "#1F7E4E"
VERDE_ESCURO = "#14603A"
VERDE_CLARO = "#8CD3A9"
VERDE_FUNDO = "#EAF5EE"
AZUL_MARINHO = "#0B3A5B"
AZUL = "#0B4F8A"
AMARELO = "#E8C21B"
VERMELHO = "#C0392B"
CINZA_TEXTO = "#4A5A6A"

CAMINHO_LOGO = Path(__file__).resolve().parents[2] / "assets" / "logo_sao_martinho.png"


@lru_cache
def logo_base64() -> str:
    return base64.b64encode(CAMINHO_LOGO.read_bytes()).decode()


CSS = f"""
<style>
.block-container {{ padding-top: 1.6rem; max-width: 1250px; }}

section[data-testid="stSidebar"] > div {{
    background: linear-gradient(180deg, {VERDE} 0%, {VERDE_ESCURO} 55%, {AZUL_MARINHO} 100%);
}}
section[data-testid="stSidebar"] * {{ color: #FFFFFF; }}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {{
    background: rgba(255,255,255,0.10); border: 1.5px dashed rgba(255,255,255,0.55);
}}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {{
    background: #FFFFFF; border: none;
}}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button * {{ color: {VERDE} !important; }}
section[data-testid="stSidebar"] [data-baseweb="select"] > div {{ background: #FFFFFF; }}
section[data-testid="stSidebar"] [data-baseweb="select"] *,
section[data-testid="stSidebar"] [data-testid="stSelectbox"] input,
section[data-testid="stSidebar"] [data-testid="stTextInput"] input,
section[data-testid="stSidebar"] [data-testid="stNumberInput"] input,
section[data-testid="stSidebar"] [data-testid="stTextArea"] textarea,
section[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] *,
section[data-testid="stSidebar"] [data-testid="stFileChip"] * {{
    color: #13263A !important; -webkit-text-fill-color: #13263A;
}}
section[data-testid="stSidebar"] [data-testid="stTextInput"] [data-baseweb="input"],
section[data-testid="stSidebar"] [data-testid="stTextInput"] [data-baseweb="base-input"],
section[data-testid="stSidebar"] [data-testid="stNumberInput"] [data-baseweb="input"],
section[data-testid="stSidebar"] [data-testid="stTextArea"] [data-baseweb="textarea"] {{ background: #FFFFFF; }}
section[data-testid="stSidebar"] [data-testid="stNumberInput"] button {{ background: #FFFFFF; }}
section[data-testid="stSidebar"] [data-testid="stNumberInput"] button * {{ color: #13263A !important; }}
section[data-testid="stSidebar"] [data-testid="stTextInput"] input {{ caret-color: #13263A; }}
section[data-testid="stSidebar"] [data-testid="stTextInput"] input::placeholder {{
    color: #6B7A89 !important; -webkit-text-fill-color: #6B7A89;
}}
section[data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,0.25); }}
div[data-baseweb="popover"] * {{ color: #13263A; }}

.smto-header {{
    position: relative; overflow: hidden; border-radius: 14px; padding: 26px 32px 24px 32px;
    background: linear-gradient(100deg, {VERDE} 0%, #13705A 45%, {AZUL_MARINHO} 100%);
    color: #FFFFFF; margin-bottom: 1.2rem; border-bottom: 3px solid {AMARELO};
}}
.smto-header h1 {{ font-weight: 900; font-size: 1.9rem; margin: 0; color: #FFFFFF; letter-spacing: .3px; }}
.smto-header p {{ margin: 6px 0 0 0; font-size: .98rem; opacity: .92; color: #FFFFFF; }}
.smto-header .faixa {{ position: absolute; top: 0; height: 100%; transform: skewX(-32deg); }}
.smto-header .f1 {{ right: 150px; width: 46px; background: rgba(140,211,169,.35); }}
.smto-header .f2 {{ right: 80px; width: 46px; background: rgba(11,79,138,.85); }}
.smto-header .f3 {{ right: 10px; width: 46px; background: rgba(255,255,255,.12); }}

.smto-titulo {{ color: {VERDE}; font-weight: 800; font-size: 1.25rem; margin: 1.4rem 0 .2rem 0; }}
.smto-subtitulo {{ color: {AZUL_MARINHO}; font-weight: 600; font-size: .92rem; margin-bottom: .7rem; }}

.smto-card {{
    background: #FFFFFF; border-radius: 12px; padding: 16px 18px; height: 100%;
    border: 1px solid #DCE7E0; border-left: 6px solid {VERDE};
    box-shadow: 0 2px 6px rgba(11,58,91,.06);
}}
.smto-card.azul {{ border-left-color: {AZUL}; }}
.smto-card.amarelo {{ border-left-color: {AMARELO}; }}
.smto-card .rotulo {{ font-size: .74rem; text-transform: uppercase; letter-spacing: .6px; color: {CINZA_TEXTO}; font-weight: 600; }}
.smto-card .valor {{ font-size: 1.22rem; font-weight: 800; color: {AZUL_MARINHO}; margin-top: 4px; line-height: 1.25; }}
.smto-card .valor.grande {{ font-size: 2rem; color: {VERDE}; }}
.smto-card .detalhe {{ font-size: .8rem; color: {CINZA_TEXTO}; margin-top: 4px; }}

.smto-barra {{ display: flex; align-items: center; gap: 12px; margin: 7px 0; }}
.smto-barra .nome {{ width: 290px; font-size: .88rem; color: {AZUL_MARINHO}; font-weight: 600; }}
.smto-barra .trilho {{ flex: 1; background: {VERDE_FUNDO}; border-radius: 6px; height: 18px; overflow: hidden; }}
.smto-barra .preenchido {{ height: 100%; border-radius: 6px; background: linear-gradient(90deg, {VERDE_CLARO}, {VERDE}); }}
.smto-barra.destaque .preenchido {{ background: linear-gradient(90deg, {VERDE}, {AZUL}); }}
.smto-barra.destaque .nome {{ color: {VERDE}; font-weight: 800; }}
.smto-barra .pct {{ width: 70px; text-align: right; font-weight: 800; color: {AZUL_MARINHO}; font-size: .92rem; }}

.smto-vazio {{
    border: 2px dashed {VERDE_CLARO}; border-radius: 14px; padding: 34px; text-align: center;
    background: {VERDE_FUNDO}; color: {AZUL_MARINHO};
}}
.smto-vazio h3 {{ color: {VERDE}; font-weight: 800; margin-bottom: 6px; }}
.smto-passo {{
    display: inline-block; background: {AMARELO}; color: {AZUL_MARINHO} !important; font-weight: 900;
    border-radius: 50%; width: 1.5rem; height: 1.5rem; text-align: center; line-height: 1.5rem; margin-right: .4rem;
}}
.smto-vazio ol {{ text-align: left; display: inline-block; margin: .6rem auto 0 auto; }}
.smto-rodape {{ font-size: .75rem; opacity: .85; margin-top: 1rem; }}
@media (max-width: 700px) {{ .smto-barra .nome {{ width: 140px; font-size: .78rem; }} }}
</style>
"""


def aplicar() -> None:
    st.markdown(CSS, unsafe_allow_html=True)


def cabecalho(titulo: str, subtitulo: str) -> None:
    st.markdown(
        f"""
        <div class="smto-header">
            <span class="faixa f1"></span><span class="faixa f2"></span><span class="faixa f3"></span>
            <h1>{titulo}</h1>
            <p>{subtitulo}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
