import csv
import html as html_lib
import re
from pathlib import Path

import streamlit as st

from _versao import atualizado_em


MESES_PT_EXTENSO = {
    "jan": "janeiro", "fev": "fevereiro", "mar": "março", "abr": "abril",
    "mai": "maio", "jun": "junho", "jul": "julho", "ago": "agosto",
    "set": "setembro", "out": "outubro", "nov": "novembro", "dez": "dezembro",
}

PREFIXO_AVISO_OVERRIDE_EFEITO_FINANCEIRO = "Marcacao de efeito financeiro ajustada manualmente:"


def render_avisos_override_efeito_financeiro(diagnostico):
    """Exibe uma vez cada aviso canonico de override produzido pelo leitor."""
    avisos = diagnostico.get("avisos", ()) if isinstance(diagnostico, dict) else ()
    if isinstance(avisos, str):
        avisos = (avisos,)

    exibidos = []
    vistos = set()
    for aviso in avisos or ():
        texto = str(aviso).strip()
        if not texto.startswith(PREFIXO_AVISO_OVERRIDE_EFEITO_FINANCEIRO) or texto in vistos:
            continue
        vistos.add(texto)
        exibidos.append(texto)
        st.warning(texto)
    return tuple(exibidos)


_HEADER_ASSET = Path(__file__).resolve().parent / "assets" / "cl8us_header_proporcional.png"


def _header_data_uri():
    """Header oficial (assets/cl8us_header_proporcional.png) como data URI cacheado."""
    import base64
    global _HEADER_DATA_URI
    try:
        return _HEADER_DATA_URI
    except NameError:
        pass
    try:
        b64 = base64.b64encode(_HEADER_ASSET.read_bytes()).decode("ascii")
        _HEADER_DATA_URI = f"data:image/png;base64,{b64}"
    except Exception:
        _HEADER_DATA_URI = ""
    return _HEADER_DATA_URI


def render_marca_topo():
    """Marca oficial do Cl8us no topo. Imagem proporcional (943x180), responsiva
    (width: min(420px, 100%); height: auto; object-fit: contain) — marca, nao banner."""
    src = _header_data_uri()
    st.markdown(
        f"""
        <style>
        .cl8us-brand-wrap {{ margin: 0 0 0.55rem 0; padding: 0.9rem 0 0; }}
        .cl8us-brand-img {{ width: min(420px, 100%); height: auto; object-fit: contain; display: block; }}
        .cl8us-brand-aviso {{ display: block; margin-top: 0.42rem; color: #A16207; font-family: "Inter", "Segoe UI", Arial, sans-serif; font-size: 0.74rem; font-style: italic; font-weight: 400; line-height: 1.20; }}
        </style>
        <div class="cl8us-brand-wrap">
            <img class="cl8us-brand-img" src="{src}" alt="TLB cl8us - apoio à gestão de contratos" />
            <div class="cl8us-brand-aviso">AVISO: use apenas para docs não sigilosos e de livre acesso.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_cabecalho_pagina(titulo, descricao=""):
    """Cabeçalho global do Cl8us: marca + título funcional + aviso de privacidade.

    Contrato global (hotfix): a descrição NUNCA é renderizada. O parâmetro
    `descricao` é mantido apenas por compatibilidade temporária com call-sites
    legados e é intencionalmente ignorado — por isso todo cabeçalho passa a
    exibir exclusivamente a marca Cl8us, o título e o aviso de privacidade, sem
    qualquer frase descritiva intermediária.

    O HTML é montado como um bloco íntegro, em uma única string sem indentação
    nem linha em branco dinâmica. Isso evita o bug pós-reboot em que a linha
    vazia (descrição ausente) fechava o bloco HTML no CommonMark do Streamlit e
    fazia o `<div class="cl8us-page-privacy">` ser renderizado literalmente como
    bloco de código.
    """
    # `descricao` recebido apenas por compatibilidade; nunca é renderizado.
    titulo_seguro = html_lib.escape(str(titulo))
    html = (
        f'<section class="cl8us-page-header" aria-label="{titulo_seguro}">'
        f'<img class="cl8us-brand-img" src="{_header_data_uri()}" '
        f'alt="TLB cl8us - apoio à gestão de contratos" '
        f'style="width:min(420px,100%);height:auto;object-fit:contain;display:block;margin:0 0 0.55rem 0;" />'
        f'<h1>{titulo_seguro}</h1>'
        f'<div class="cl8us-page-privacy">Use apenas para documentos não sigilosos e de livre acesso.</div>'
        f'</section>'
    )
    st.markdown(html, unsafe_allow_html=True)


def render_versao_sidebar():
    """Exibe o carimbo obrigatório da versão efetivamente publicada."""
    st.markdown('<div class="cl8us-version-rule"></div>', unsafe_allow_html=True)
    st.caption(f"Atualizado em {atualizado_em()}")


def _normalizar_mes_ano_ist(valor):
    if valor is None:
        return None
    texto = str(valor).strip().lower().replace(".", "").replace("-", "/").replace("_", "/")
    match = re.match(r"^([a-zç]{3,9})\s*/\s*(\d{2,4})$", texto)
    if not match:
        return None
    mes_raw, ano_raw = match.groups()
    mes = mes_raw[:3]
    if mes not in MESES_PT_EXTENSO:
        return None
    ano = int(ano_raw)
    if ano < 100:
        ano += 2000
    ordem_mes = list(MESES_PT_EXTENSO.keys()).index(mes) + 1
    return {"ordem": ano * 100 + ordem_mes, "descricao": f"{MESES_PT_EXTENSO[mes]}/{ano}"}


def obter_ultima_competencia_ist(caminho="ist.csv"):
    caminho_csv = Path(caminho)
    if not caminho_csv.exists():
        return None

    melhor = None
    try:
        with caminho_csv.open("r", encoding="utf-8-sig", newline="") as arquivo:
            leitor = csv.DictReader(arquivo, delimiter=";")
            for linha in leitor:
                mes_ano = linha.get("MES_ANO") or linha.get("mes_ano")
                indice = linha.get("INDICE_NIVEL") or linha.get("indice_nivel")
                info = _normalizar_mes_ano_ist(mes_ano)
                if not info:
                    continue
                if melhor is None or info["ordem"] > melhor["ordem"]:
                    melhor = {**info, "indice": str(indice).strip() if indice is not None else ""}
    except Exception:
        return None
    return melhor


def render_alerta_ist_local():
    ultima = obter_ultima_competencia_ist()
    if ultima:
        texto = f"IST local: última competência constante no sistema: <strong>{ultima['descricao']}</strong>"
        if ultima.get("indice"):
            texto += f" — índice em nível: <strong>{ultima['indice']}</strong>"
        texto += ". Confira se a data-base necessária já está coberta."
    else:
        texto = "IST local: não foi possível identificar a última competência no arquivo <code>ist.csv</code>. Confira a base antes de prosseguir."

    st.markdown(
        f"""
        <div style="background:#F8FAFC; border:1px solid #E5EAF0; border-radius:9px;
                    padding:7px 10px; margin:5px 0 2px 0; color:#475569;
                    font-size:0.80rem; line-height:1.35;">
            {texto}
        </div>
        """,
        unsafe_allow_html=True,
    )




@st.cache_data(ttl=60 * 60, show_spinner=False)
def _obter_ultima_competencia_icti_cache():
    from _indice_utils import obter_ultima_competencia_icti_ipeadata
    return obter_ultima_competencia_icti_ipeadata(timeout=15)


@st.cache_data(ttl=60 * 60, show_spinner=False)
def _obter_ultima_competencia_sgs_cache(serie_codigo):
    from _indice_utils import obter_ultima_competencia_sgs
    return obter_ultima_competencia_sgs(serie_codigo, timeout=15)


def _texto_ultima_competencia_sgs(serie_codigo):
    """§18: 'Última competência disponível: mm/aaaa.' para IPCA(433)/IGP-M(189).

    Consulta a MESMA fonte oficial do cálculo (SGS/BCB) com cache leve. Em falha
    de rede não bloqueia o usuário: exibe aviso discreto sem inventar data.
    """
    try:
        ultima = _obter_ultima_competencia_sgs_cache(serie_codigo)
        return f"Última competência disponível: {ultima['descricao']}."
    except Exception:
        return "Última competência disponível: não foi possível consultar neste momento."


def render_alerta_icti_ipeadata():
    """Mostra alerta discreto sobre a última competência ICTI disponível no Ipeadata."""
    try:
        ultima = _obter_ultima_competencia_icti_cache()
        texto = (
            f"ICTI/Ipeadata: última competência disponível: <strong>{ultima['descricao']}</strong> "
            f"— série <strong>{ultima.get('serie') or ultima.get('sercodigo') or 'DIMAC_ICTI2'}</strong>. "
            "O cálculo usa a competência do mês anterior como índice-base e acumula as taxas mensais do período."
        )
    except Exception:
        texto = (
            "ICTI/Ipeadata: não foi possível consultar a última competência neste momento. "
            "Verifique a conexão com a internet ou tente novamente."
        )

    st.markdown(
        f"""
        <div style="background:#F8FAFC; border:1px solid #E5EAF0; border-radius:9px;
                    padding:7px 10px; margin:5px 0 2px 0; color:#475569;
                    font-size:0.80rem; line-height:1.35;">
            {texto}
        </div>
        """,
        unsafe_allow_html=True,
    )

def render_indice_contrato_selectbox(key=None, index=0, options=None):
    """Renderiza o campo de índice com destaque visual consistente entre os fluxos."""
    if options is None:
        options = ["IST (Série Local)", "ICTI (Ipeadata)", "IPCA (433)", "IGP-M (189)"]

    with st.container(border=True):
        st.markdown(
            '<span class="cl8us-index-marker"></span>'
            '<div class="cl8us-index-title">Índice do contrato</div>',
            unsafe_allow_html=True,
        )

        selecionado = st.selectbox(
            "Índice do contrato",
            options,
            index=index,
            key=key,
            label_visibility="collapsed",
            help="Confirme o índice contratual antes de gerar cálculos, arquivo de coleta ou relatórios.",
        )

        if selecionado == options[index]:
            st.caption(f"Conferência necessária: confirme se o índice contratual correto é **{selecionado}**.")
        else:
            st.caption(f"Índice selecionado para esta análise: **{selecionado}**.")

        selecionado_norm = str(selecionado).strip().upper()
        if selecionado_norm.startswith("IST"):
            render_alerta_ist_local()
        elif selecionado_norm.startswith("ICTI"):
            render_alerta_icti_ipeadata()
        elif selecionado_norm.startswith("IPCA"):
            # §18: mesma fonte oficial do cálculo (SGS 433); IST/ICTI intactos.
            st.caption(_texto_ultima_competencia_sgs(433))
        elif selecionado_norm.startswith("IGP-M") or selecionado_norm.startswith("IGPM"):
            st.caption(_texto_ultima_competencia_sgs(189))

    return selecionado


def render_aviso_privacidade(tem_upload=False, tem_download=False):
    """Aviso resumido de privacidade para páginas com upload e/ou download."""
    if not tem_upload and not tem_download:
        return

    partes = []
    if tem_upload:
        partes.append("Os arquivos enviados são usados para processamento na sessão do app e não são enviados ao repositório. O upload permanece em memória durante o uso da sessão.")
    if tem_download:
        partes.append("Os documentos gerados são disponibilizados apenas para download pelo navegador.")
    partes.append("Ao limpar ou substituir o arquivo, ou ao encerrar a aba/sessão, os dados deixam de ser necessários para o processamento. Evite carregar dados sigilosos desnecessários e confira os arquivos antes de compartilhar.")

    texto = " ".join(partes)
    st.markdown(
        f"""
        <div style="background:#F0F9FF; border:1px solid #BAE6FD; border-radius:10px;
                    padding:9px 13px; margin:8px 0 14px 0; color:#0C4A6E;
                    font-size:0.84rem; line-height:1.45;">
            🔒 <strong>Privacidade:</strong> {texto}
        </div>
        """,
        unsafe_allow_html=True,
    )
