"""Pagina inicial do fluxo XLS-first do Master 2.0.

A HOME nao repete a navegacao do menu lateral: ela apresenta o fluxo de tres
etapas, destaca o Arquivo Coleta Oficial e devolve o usuario ao menu. O unico
controle interativo e o download da Coleta, que reutiliza exatamente a mesma
origem ja usada antes (gerar_coleta_oficial_preenchida + NOME_DOWNLOAD_COLETA).

Todo HTML injetado e montado como UMA string sem indentacao e sem linha em
branco: indentacao dentro de st.markdown faz o CommonMark do Streamlit fechar o
bloco e renderizar a marcacao como codigo literal.
"""

import streamlit as st

from _coleta_oficial import (
    NOME_ARQUIVO_COLETA_OFICIAL,
    NOME_DOWNLOAD_COLETA,
    TEMPLATE_COLETA_OFICIAL,
    assinatura_template_coleta,
    gerar_coleta_oficial_preenchida,
)
from _ui_utils import _header_data_uri


st.set_page_config(
    page_title="TLB · cl8us — Início",
    page_icon="assets/cl8us_favicon_512.png",
    layout="wide",
)

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@st.cache_data(show_spinner=False)
def _ler_modelo(assinatura: str) -> bytes:  # noqa: ARG001 — chave SHA-256 do cache
    return gerar_coleta_oficial_preenchida(None)


# ---------------------------------------------------------------------------
# Pictogramas — SVG inline, sem dependencia externa e sem emoji.
# Traco azul da marca sobre disco azul muito claro (identidade cl8us).
# ---------------------------------------------------------------------------

_SVG_ABRE = (
    '<svg viewBox="0 0 32 32" width="20" height="20" fill="none" '
    'stroke="currentColor" stroke-width="1.7" stroke-linecap="round" '
    'stroke-linejoin="round" aria-hidden="true" focusable="false">'
)

ICONE_DOCUMENTO = (
    _SVG_ABRE
    + '<path d="M18.5 4H9A2.5 2.5 0 0 0 6.5 6.5v19A2.5 2.5 0 0 0 9 28h14a2.5 2.5 0 0 0 2.5-2.5V11z"/>'
    '<path d="M18.5 4v7h7"/>'
    '<path d="M11 12.5h4"/><path d="M11 17.5h10"/><path d="M11 22h10"/>'
    "</svg>"
)

ICONE_FISCAL = (
    _SVG_ABRE
    + '<circle cx="12.5" cy="10" r="4.3"/>'
    '<path d="M4.5 26.5c0-4.5 3.6-7.4 8-7.4 1.3 0 2.6.3 3.7.8"/>'
    '<path d="M26.4 16.3a1.95 1.95 0 0 1 2.8 2.8l-7.4 7.4-3.6.8.8-3.6z"/>'
    "</svg>"
)

ICONE_APURACAO = (
    _SVG_ABRE
    + '<path d="M16 19.5V5.5"/><path d="M11 10.5 16 5.5l5 5"/>'
    '<path d="M5 20.5v3.8A2.5 2.5 0 0 0 7.5 26.8h8"/>'
    '<circle cx="23" cy="23" r="5.4"/>'
    '<path d="M20.8 23.1 22.5 24.8 25.4 21.6"/>'
    "</svg>"
)

ICONE_LAMPADA = (
    _SVG_ABRE
    + '<path d="M16 3.6a8 8 0 0 0-4.7 14.5c.9.7 1.5 1.7 1.5 2.8v.6h6.4v-.6c0-1.1.6-2.1 1.5-2.8A8 8 0 0 0 16 3.6z"/>'
    '<path d="M12.8 24.6h6.4"/><path d="M13.9 27.6h4.2"/>'
    "</svg>"
)

SETA_ETAPA = (
    '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" '
    'stroke="currentColor" stroke-width="1.9" stroke-linecap="round" '
    'stroke-linejoin="round" aria-hidden="true" focusable="false">'
    '<path d="M4 12h15"/><path d="M13.4 6.4 19.9 12l-6.5 5.6"/>'
    "</svg>"
)

# Pilha de documentos + folha principal + escudo com check.
ILUSTRACAO_COLETA = (
    '<svg viewBox="0 0 122 92" width="72" height="54" fill="none" '
    'stroke="#1F5F8B" stroke-width="1.7" stroke-linecap="round" '
    'stroke-linejoin="round" aria-hidden="true" focusable="false">'
    '<rect x="10" y="18" width="48" height="62" rx="5" fill="#FFFFFF" opacity=".55"/>'
    '<rect x="19" y="12" width="48" height="62" rx="5" fill="#FFFFFF" opacity=".8"/>'
    '<path d="M33 5h25l13 13v49a5 5 0 0 1-5 5H33a5 5 0 0 1-5-5V10a5 5 0 0 1 5-5z" fill="#FFFFFF"/>'
    '<path d="M58 5v13h13"/>'
    '<path d="M36 30h22"/><path d="M36 38h22"/><path d="M36 46h14"/>'
    '<path d="M92 42 106 47 106 58c0 8.2-5.7 13.4-14 16.2C83.7 71.4 78 66.2 78 58V47Z" fill="#E9F2F9"/>'
    '<path d="M86.6 58.4 90.3 62.1 97.6 54.6"/>'
    "</svg>"
)

ETAPAS = (
    (
        ICONE_DOCUMENTO,
        "1. PREPARE",
        "Defina as datas, os ciclos e gere o Arquivo Coleta Oficial.",
    ),
    (
        ICONE_FISCAL,
        "2. PREENCHA",
        "Após baixar o arquivo pré-preenchido, você envia para o fiscal "
        "complementar as informações necessárias.",
    ),
    (
        ICONE_APURACAO,
        "3. APURE",
        "Envie o arquivo completo e obtenha resultados e documentos.",
    ),
)


def _estilos_home() -> None:
    """CSS exclusivo da HOME. Todo seletor e prefixado por `home-` ou ancorado
    em um marcador desta pagina, para nunca vazar para as demais rotas."""
    st.markdown(
        "<style>"
        # -- Cabecalho preservado: logo a esquerda, titulo e badge ------------
        ".home-hero{margin:-.35rem 0 1.25rem;padding:0;text-align:left;}"
        ".home-hero img{width:min(252px,55%);height:auto;object-fit:contain;display:block;"
        "margin:0 0 .7rem;mix-blend-mode:multiply;}"
        ".home-hero-linha{display:flex;align-items:center;flex-wrap:wrap;gap:.45rem .75rem;}"
        ".home-hero h1{color:var(--cl8us-navy);font-size:1.6rem;letter-spacing:-.015em;margin:0;padding:0;}"
        ".home-hero .cl8us-page-privacy{font-size:.65rem;padding:.18rem .5rem;}"
        # -- Tres etapas: compactas, horizontais, mesma altura ---------------
        ".home-fluxo{display:flex;align-items:stretch;gap:.35rem;margin:0 0 .85rem;}"
        ".home-etapa{background:#FFFFFF;border:1px solid rgba(18,59,99,.13);border-radius:13px;"
        "box-shadow:0 2px 9px rgba(18,59,99,.045);display:flex;flex:1 1 0;flex-direction:column;"
        "min-width:0;padding:.52rem .78rem .58rem;text-align:center;}"
        ".home-ico{align-items:center;background:#E9F2F9;border-radius:50%;color:#1F5F8B;"
        "display:inline-flex;flex:0 0 auto;height:34px;justify-content:center;width:34px;}"
        ".home-etapa .home-ico{margin:0 auto .32rem;}"
        ".home-etapa h3{color:var(--cl8us-navy);font-size:.8rem;font-weight:850;"
        "letter-spacing:.05em;margin:0 0 .18rem;padding:0;}"
        ".home-etapa p{color:var(--cl8us-muted);font-size:.9rem;line-height:1.32;margin:0;}"
        ".home-seta{align-items:center;color:rgba(31,95,139,.4);display:flex;flex:0 0 auto;"
        "justify-content:center;padding:0;}"
        # -- Cartoes inferiores: a COLUNA e o cartao, para o botao real de
        # download ficar dentro dele (nenhuma logica de geracao duplicada).
        '[data-testid="stColumn"]:has(.home-coleta-corpo){border-radius:14px;'
        "background:linear-gradient(180deg,#FFFFFF 0%,#F8FBFD 100%);"
        "border:1px solid rgba(18,59,99,.15);box-shadow:0 2px 9px rgba(18,59,99,.05);"
        "padding:.7rem .9rem .78rem;}"
        '[data-testid="stColumn"]:has(.home-dica-corpo){border-radius:14px;background:#F5FAFC;'
        "border:1px solid rgba(31,95,139,.14);display:flex;flex-direction:column;"
        "justify-content:center;padding:.6rem .8rem;}"
        '[data-testid="stColumn"]:has(.home-dica-corpo) [data-testid="stVerticalBlock"]{'
        "justify-content:center;}"
        # Botao dentro do cartao, alinhado ao titulo/texto e sem virar faixa.
        '[data-testid="stColumn"]:has(.home-coleta-corpo) .stDownloadButton{'
        "margin:.5rem 0 0 calc(72px + .85rem);}"
        '[data-testid="stColumn"]:has(.home-coleta-corpo) .stDownloadButton button{'
        "font-size:.83rem;min-height:1.9rem;padding:.12rem 1rem;}"
        ".home-coleta-corpo{display:flex;align-items:center;gap:.85rem;}"
        ".home-coleta-corpo .home-arte{flex:0 0 72px;line-height:0;}"
        ".home-coleta-corpo h2{color:var(--cl8us-navy);font-size:1rem;margin:0 0 .18rem;padding:0;}"
        ".home-coleta-corpo p{color:var(--cl8us-muted);font-size:.89rem;line-height:1.34;margin:0;}"
        ".home-dica-corpo{align-items:center;display:flex;gap:.62rem;}"
        ".home-dica-corpo .home-ico{height:30px;width:30px;background:#E7F0F7;}"
        ".home-dica-corpo p{color:var(--cl8us-muted);font-size:.83rem;font-weight:500;"
        "line-height:1.38;margin:0;}"
        # -- Largura menor: empilha, gira as setas e solta o recuo do botao ---
        "@media (max-width:900px){"
        ".home-fluxo{flex-direction:column;}"
        ".home-seta{transform:rotate(90deg);padding:.1rem 0;}"
        ".home-coleta-corpo{flex-direction:column;text-align:center;gap:.6rem;}"
        '[data-testid="stColumn"]:has(.home-coleta-corpo) .stDownloadButton{margin-left:0;}'
        ".home-hero img{width:min(230px,68%);}"
        ".home-hero h1{font-size:1.36rem;}"
        "}"
        "</style>",
        unsafe_allow_html=True,
    )


def _render_hero() -> None:
    """Cabecalho proprio da HOME: logo oficial a esquerda, no eixo da grade,
    com o titulo logo abaixo e o aviso de sigilo como badge discreta ao lado.

    Reutiliza o MESMO asset do cabecalho global
    (assets/cl8us_header_proporcional.png, via _header_data_uri): simbolo,
    `cl8us` e a assinatura ja fazem parte da arte oficial — nada e redesenhado
    nem recriado aqui.
    """
    st.markdown(
        '<section class="home-hero" aria-label="Reajustes contratuais">'
        f'<img src="{_header_data_uri()}" alt="TLB cl8us - apoio à gestão de contratos" />'
        '<div class="home-hero-linha"><h1>Reajustes contratuais</h1>'
        '<span class="cl8us-page-privacy">'
        "Use apenas para documentos não sigilosos e de livre acesso."
        "</span></div>"
        "</section>",
        unsafe_allow_html=True,
    )


def _render_fluxo() -> None:
    """Tres etapas ligadas por setas discretas. Bloco puramente explicativo:
    sem botoes e sem links — a navegacao permanece no menu lateral."""
    partes = ['<div class="home-fluxo">']
    for indice, (icone, titulo, texto) in enumerate(ETAPAS):
        if indice:
            partes.append(f'<div class="home-seta">{SETA_ETAPA}</div>')
        partes.append(
            '<article class="home-etapa">'
            f'<span class="home-ico">{icone}</span>'
            f"<h3>{titulo}</h3><p>{texto}</p>"
            "</article>"
        )
    partes.append("</div>")
    st.markdown("".join(partes), unsafe_allow_html=True)


_estilos_home()
_render_hero()
_render_fluxo()

coluna_coleta, coluna_dica = st.columns([2.3, 1], gap="medium")

with coluna_coleta:
    st.markdown(
        '<div class="home-coleta-corpo" role="group" aria-label="Arquivo Coleta Oficial">'
        f'<div class="home-arte">{ILUSTRACAO_COLETA}</div>'
        "<div><h2>Arquivo Coleta Oficial</h2>"
        "<p>É o arquivo central da apuração e acompanha o processo da "
        "preparação até os resultados.</p></div>"
        "</div>",
        unsafe_allow_html=True,
    )
    if TEMPLATE_COLETA_OFICIAL.exists():
        st.download_button(
            "Baixar Arquivo Coleta Oficial",
            data=_ler_modelo(assinatura_template_coleta()),
            file_name=NOME_DOWNLOAD_COLETA,
            mime=MIME_XLSX,
            type="primary",
            use_container_width=False,
            key="download_coleta_inicio",
        )
    else:
        st.error(
            "Não foi possível localizar o Arquivo Coleta Oficial neste ambiente. "
            "O download foi bloqueado para evitar o uso de modelo incompatível."
        )

with coluna_dica:
    st.markdown(
        '<div class="home-dica-corpo" role="note">'
        f'<span class="home-ico">{ICONE_LAMPADA}</span>'
        "<p>Use o menu lateral para acessar as funcionalidades do cl8us.</p>"
        "</div>",
        unsafe_allow_html=True,
    )
