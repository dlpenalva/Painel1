import csv
import re
from pathlib import Path

import streamlit as st


MESES_PT_EXTENSO = {
    "jan": "janeiro", "fev": "fevereiro", "mar": "março", "abr": "abril",
    "mai": "maio", "jun": "junho", "jul": "julho", "ago": "agosto",
    "set": "setembro", "out": "outubro", "nov": "novembro", "dez": "dezembro",
}


def render_marca_topo(titulo_pagina="", subtitulo_pagina=""):
    """Identidade visual cl8us — header unificado para todas as páginas."""
    _sub = str(subtitulo_pagina or "")
    _tit = str(titulo_pagina or "")

    _topbar = (
        '<style>'
        '[data-testid="stAppViewContainer"] > section > div:first-child { padding-top: 0 !important; }'
        '</style>'
        '<div style="display:flex;align-items:center;justify-content:space-between;'
        'padding:12px 0 10px 0;border-bottom:0.5px solid #E2E8F0;margin-bottom:18px">'
        '<div style="display:flex;align-items:center;gap:10px">'
        '<div style="width:32px;height:32px;border-radius:9px;'
        'background:linear-gradient(135deg,#1F4E78 0%,#185FA5 100%);'
        'display:flex;align-items:center;justify-content:center;flex-shrink:0">'
        '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="white"'
        ' stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
        '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'
        '<polyline points="14 2 14 8 20 8"/>'
        '<line x1="16" y1="13" x2="8" y2="13"/>'
        '<line x1="16" y1="17" x2="8" y2="17"/>'
        '<polyline points="10 9 9 9 8 9"/>'
        '</svg></div>'
        '<div style="display:flex;align-items:baseline;gap:6px">'
        '<span style="font-size:15px;font-weight:600;color:#1F4E78;letter-spacing:-.01em">TLB</span>'
        '<span style="font-size:13px;color:#CBD5E1">&nbsp;·&nbsp;</span>'
        '<span style="font-size:13px;font-weight:500;color:#0F172A">cl8us</span>'
        '</div></div>'
        '<div style="display:flex;align-items:center;gap:5px;'
        'background:#F0FDF4;border:0.5px solid #BBF7D0;border-radius:20px;padding:3px 10px">'
        '<div style="width:5px;height:5px;border-radius:50%;background:#16A34A"></div>'
        '<span style="font-size:11px;color:#15803D;font-weight:500">uso para docs não sigilosos</span>'
        '</div></div>'
    )

    _titulo_html = ""
    if _sub:
        _titulo_html += (
            '<p style="font-size:10px;font-weight:600;letter-spacing:.07em;'
            'text-transform:uppercase;color:#94A3B8;margin-bottom:5px">'
            + _sub + '</p>'
        )
    if _tit:
        _titulo_html += (
            '<p style="font-size:22px;font-weight:500;color:#0F172A;'
            'margin-bottom:0;letter-spacing:-.02em">'
            + _tit + '</p>'
        )

    st.markdown(_topbar + _titulo_html, unsafe_allow_html=True)


def render_versao_sidebar():
    """Função mantida por compatibilidade; não renderiza versão fixa."""
    return None


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
    """Retorna a última competência disponível no ist.csv.

    Busca robusta:
    1. caminho informado;
    2. pasta atual de execução;
    3. pasta onde está o próprio _ui_utils.py.

    Mantém compatibilidade com o CSV atual:
    MES_ANO;INDICE_NIVEL
    jan/22;298,959
    """
    candidatos = []
    caminho_csv = Path(caminho)

    candidatos.append(caminho_csv)
    if not caminho_csv.is_absolute():
        candidatos.append(Path.cwd() / caminho)
        try:
            candidatos.append(Path(__file__).resolve().parent / caminho)
        except Exception:
            pass

    vistos = set()
    candidatos_unicos = []
    for c in candidatos:
        try:
            chave = str(c.resolve())
        except Exception:
            chave = str(c)
        if chave not in vistos:
            vistos.add(chave)
            candidatos_unicos.append(c)

    melhor = None

    for arq_csv in candidatos_unicos:
        if not arq_csv.exists():
            continue

        try:
            with arq_csv.open("r", encoding="utf-8-sig", newline="") as arquivo:
                leitor = csv.DictReader(arquivo, delimiter=";")
                for linha in leitor:
                    mes_ano = linha.get("MES_ANO") or linha.get("mes_ano") or linha.get("MÊS_ANO") or linha.get("MES")
                    indice = linha.get("INDICE_NIVEL") or linha.get("indice_nivel") or linha.get("ÍNDICE_NÍVEL")
                    info = _normalizar_mes_ano_ist(mes_ano)
                    if not info:
                        continue
                    if melhor is None or info["data"] > melhor["data"]:
                        melhor = {
                            **info,
                            "indice": str(indice).strip() if indice is not None else "",
                            "arquivo": str(arq_csv),
                        }
        except Exception:
            continue

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




@st.cache_data(ttl=60 * 60)
def _obter_ultima_competencia_icti_cache():
    from _indice_utils import obter_ultima_competencia_icti_ipeadata
    return obter_ultima_competencia_icti_ipeadata(timeout=15)


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
    """Renderiza o campo de índice com fallback visual seguro.

    Garante que os índices principais apareçam mesmo se algum fluxo enviar
    options vazio/None. Não altera a regra de cálculo: a função continua
    retornando o texto do índice selecionado.
    """
    opcoes_padrao = ["IST (Série Local)", "ICTI (Ipeadata)", "IPCA (433)", "IGP-M (189)"]

    if not options:
        options = list(opcoes_padrao)
    else:
        options = [str(o).strip() for o in list(options) if str(o).strip()]
        if not options:
            options = list(opcoes_padrao)

    for opt in opcoes_padrao:
        if opt not in options:
            options.append(opt)

    try:
        index = int(index or 0)
    except Exception:
        index = 0
    if index < 0 or index >= len(options):
        index = 0

    with st.container():
        st.markdown(
            """
            <div class="cl8us-indice-box">
                <div style="font-weight:800; font-size:0.98rem;">Índice do contrato</div>
                <div style="font-size:0.84rem; opacity:0.86; margin-top:2px;">
                    Revise este campo antes de prosseguir. Ele deve corresponder ao índice previsto no contrato.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        selecionado = st.selectbox(
            "Índice:",
            options,
            index=index,
            key=key,
            help="Confirme o índice contratual antes de gerar cálculos, arquivo de coleta ou relatórios.",
        )

        st.caption("Opções disponíveis: " + " · ".join(options))

        selecionado_norm = str(selecionado or "").strip().upper()
        if selecionado_norm:
            st.caption(f"Índice selecionado para esta análise: **{selecionado}**.")
        else:
            st.warning("Nenhum índice foi selecionado. Confira a lista de opções antes de prosseguir.", icon="⚠️")

        if selecionado_norm.startswith("IST"):
            render_alerta_ist_local()

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
