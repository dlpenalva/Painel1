"""Garantia Contratual — uma única linha do tempo, 100% manual.

A página conta a evolução do contrato em ordem cronológica, e a garantia caminha
junto com ela:

    SITUAÇÃO ORIGINAL -> ALTERAÇÕES POSTERIORES + GARANTIA APÓS CADA EVENTO
        -> SITUAÇÃO ATUAL -> RESULTADO -> TEXTO À CONTRATADA

Não há quadro de garantias separado: cada linha de alteração registra o que
aconteceu com o contrato E qual passou a ser a garantia depois daquele evento.
A garantia de cada linha é a FOTOGRAFIA vigente após o evento — nunca uma
parcela a somar às anteriores.

Ferramenta 100% MANUAL: todos os dados são digitados aqui. A página não lê o
VTA, o Valor Global, a Coleta, os RESULTADOS, o XLS nem qualquer outra chave de
sessão de outra página — mesmo que esses dados existam na sessão, são ignorados.
Também não pede identificação de contrato, contratada ou apólice. O único uso de
``st.session_state`` é o estado da própria página: a grade, a navegação de
retorno e o resultado próprio publicado ao final.

Toda a matemática vive no motor puro ``_garantia_calculo`` (Decimal +
ROUND_HALF_UP), permitindo testes focais.
"""
from html import escape

import pandas as pd
import streamlit as st

from _ui_utils import render_cabecalho_pagina
from _garantia_calculo import (
    COLUNA_EVENTO_DATA,
    COLUNA_EVENTO_GARANTIA,
    COLUNA_EVENTO_TIPO,
    COLUNA_EVENTO_VALIDADE,
    COLUNA_EVENTO_VALOR,
    COLUNA_EVENTO_VIGENCIA,
    DIAGNOSTICO_REGULAR,
    DIAGNOSTICO_VALIDADE,
    DIAGNOSTICO_VALOR,
    DIAGNOSTICO_VALOR_E_VALIDADE,
    DIAS_VALIDADE_MINIMA,
    FINANCEIRO_COMPLEMENTAR,
    FINANCEIRO_SUFICIENTE,
    PERCENTUAL_GARANTIA_PADRAO,
    TEMPORAL_NAO_INFORMADA,
    TEMPORAL_SUFICIENTE,
    TIPOS_EVENTO,
    TRACO,
    analisar_garantia,
    calcular_garantia_necessaria,
    calcular_situacao_atual,
    calcular_validade_minima,
    formatar_brl,
    formatar_brl_opcional,
    formatar_data_br,
    formatar_percentual,
    formatar_variacao,
    gerar_texto_comunicacao,
    normalizar_eventos,
    parse_moeda_br,
)

st.set_page_config(page_icon="assets/cl8us_favicon_512.png", page_title="Análises de Reajustes - Garantia", layout="wide")


def css():
    st.markdown(
        """
        <style>
        .garantia-card {
            background: #F6F8FA;
            border: 1px solid #E1E6EB;
            border-radius: 14px;
            padding: 16px 18px;
            margin: 6px 0 12px 0;
            min-height: 104px;
        }
        .garantia-card-destaque {
            background: #EAF2F8;
            border: 1px solid #C8D9E8;
            border-radius: 14px;
            padding: 16px 18px;
            margin: 6px 0 12px 0;
            min-height: 104px;
        }
        .garantia-label { color: #475569; font-size: 0.9rem; margin-bottom: 4px; }
        .garantia-valor { color: #1F2937; font-size: 1.4rem; font-weight: 700; line-height: 1.2; }
        .garantia-valor-destaque { color: #123B63; font-size: 1.6rem; font-weight: 800; line-height: 1.2; }
        .garantia-nota { color: #64748B; font-size: 0.85rem; margin-top: 6px; }
        .garantia-tabela-wrap { width: 100%; overflow-x: auto; margin: 10px 0 18px 0; }
        table.garantia-tabela { width: 100%; border-collapse: collapse; font-size: 0.9rem; }
        table.garantia-tabela th { background: #E6F0F7; color: #173B5D; border: 1px solid #C5D6E2; padding: 8px 10px; text-align: left; font-weight: 700; }
        table.garantia-tabela td { border: 1px solid #E5EAF0; padding: 8px 10px; vertical-align: top; white-space: normal; overflow-wrap: anywhere; word-break: normal; line-height: 1.35; }
        table.garantia-tabela td.valor { white-space: nowrap; overflow-wrap: normal; text-align: right; }
        table.garantia-tabela tr.garantia-linha-marco td { background: #FBFCFD; color: #475569; font-style: italic; }
        .garantia-resumo {
            background: #FFFFFF;
            border: 1px solid #E1E6EB;
            border-radius: 16px;
            padding: 6px 18px;
            margin: 6px 0 18px 0;
        }
        .garantia-resumo-linha {
            display: flex;
            flex-wrap: wrap;
            align-items: baseline;
            justify-content: space-between;
            gap: 6px 20px;
            padding: 15px 14px;
            border-radius: 12px;
            border-bottom: 1px solid #EEF2F6;
        }
        .garantia-resumo-linha:last-child { border-bottom: none; }
        .garantia-resumo-linha.destaque {
            background: #F1F8F3;
            border: 1px solid #CFE6D7;
            box-shadow: inset 4px 0 0 0 #2F7D51;
            margin: 8px 0;
        }
        .garantia-resumo-texto { flex: 1 1 260px; min-width: 0; }
        .garantia-resumo-rotulo { color: #475569; font-size: 0.93rem; font-weight: 600; line-height: 1.3; }
        .garantia-resumo-linha.destaque .garantia-resumo-rotulo { color: #2C5B43; }
        .garantia-resumo-nota { color: #7A8798; font-size: 0.82rem; margin-top: 5px; line-height: 1.4; }
        .garantia-resumo-valor {
            flex: 0 0 auto;
            margin-left: auto;
            color: #1F2937;
            font-size: 1.45rem;
            font-weight: 700;
            line-height: 1.2;
            white-space: nowrap;
        }
        .garantia-resumo-linha.destaque .garantia-resumo-valor { color: #14532D; font-size: 1.62rem; font-weight: 800; }
        .garantia-status {
            background: #F1F8F3;
            border: 1px solid #CFE6D7;
            border-radius: 10px;
            padding: 12px 14px;
            margin: 4px 0 12px 0;
            color: #24402F;
            font-size: 0.92rem;
            line-height: 1.45;
        }
        .garantia-status strong { color: #14532D; }
        .garantia-status-forte { background: #E9F4EE; border: 1px solid #B9DCC6; font-size: 0.98rem; }
        .garantia-status-ambar { background: #FDF6E8; border: 1px solid #EBDCB4; color: #4A3B1A; }
        .garantia-status-ambar strong { color: #7A5A12; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def card(label, valor, nota=None, destaque=False):
    classe = "garantia-card-destaque" if destaque else "garantia-card"
    valor_classe = "garantia-valor-destaque" if destaque else "garantia-valor"
    nota_html = f'<div class="garantia-nota">{escape(nota)}</div>' if nota else ""
    st.markdown(
        f"""
        <div class="{classe}">
            <div class="garantia-label">{escape(label)}</div>
            <div class="{valor_classe}">{escape(valor)}</div>
            {nota_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def _celula(texto, classe=""):
    atributo = f" class='{classe}'" if classe else ""
    return f"<td{atributo}>{escape(texto)}</td>"


def render_evolucao_contrato(situacao):
    """Memória cronológica: o marco original e o que ficou após cada evento.

    A coluna da garantia mostra a FOTOGRAFIA vigente após a linha — os valores
    das linhas nunca se somam.
    """
    linhas = [
        "<tr class='garantia-linha-marco'>"
        + _celula("0")
        + _celula("Assinatura")
        + _celula(TRACO, "valor")
        + _celula(formatar_brl(situacao["valor_original"]), "valor")
        + _celula(TRACO, "valor")
        + _celula(formatar_brl(situacao["garantia_original"]), "valor")
        + _celula(formatar_data_br(situacao["vigencia_original"]), "valor")
        + _celula(formatar_brl_opcional(situacao["garantia_apresentada_original"]), "valor")
        + _celula(formatar_data_br(situacao["validade_apresentada_original"]), "valor")
        + "</tr>"
    ]
    for etapa in situacao["linha_do_tempo"]:
        linhas.append(
            "<tr>"
            + _celula(str(etapa["numero"]))
            + _celula(etapa["tipo"])
            + _celula(formatar_data_br(etapa["data"]), "valor")
            + _celula(formatar_brl(etapa["valor"]), "valor")
            + _celula(formatar_variacao(etapa["variacao"]), "valor")
            + _celula(formatar_brl(etapa["garantia_exigida"]), "valor")
            + _celula(formatar_data_br(etapa["vigencia"]), "valor")
            + _celula(formatar_brl_opcional(etapa["garantia_apresentada"]), "valor")
            + _celula(formatar_data_br(etapa["validade_apresentada"]), "valor")
            + "</tr>"
        )
    cabecalhos = [
        "Nº", "Evento", "Data", "Valor do contrato", "Variação",
        "Garantia exigida", "Término da vigência", "Garantia apresentada", "Validade",
    ]
    larguras = ["4%", "13%", "9%", "13%", "12%", "12%", "11%", "13%", "13%"]
    colunas = "".join(f'<col style="width: {largura};">' for largura in larguras)
    cabecalho = "".join(f"<th>{escape(titulo)}</th>" for titulo in cabecalhos)
    st.markdown(
        f"""
        <div class="garantia-tabela-wrap">
          <table class="garantia-tabela">
            <colgroup>{colunas}</colgroup>
            <thead><tr>{cabecalho}</tr></thead>
            <tbody>{"".join(linhas)}</tbody>
          </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


def status_box(titulo, detalhe="", forte=False, ambar=False):
    """Box informativo próprio, em verde claro suave ou âmbar muito suave.

    Muda a linguagem visual, nunca o diagnóstico: o título continua sendo a
    constante apurada pelo motor e o detalhe, a mesma frase de antes. O verde é
    próprio, e não o verde forte de ``st.success``.

    Os boxes intermediários (situação financeira e situação da validade) são
    sempre verdes — ali a cor apenas acomoda a leitura. Só a CONCLUSÃO
    (``forte``) carrega semântica de cor: verde para situação regular, âmbar
    para necessidade de ação, e o vermelho de ``st.error`` para as duas
    dimensões pendentes ao mesmo tempo.
    """
    classe = "garantia-status"
    if forte:
        classe += " garantia-status-forte"
    if ambar:
        classe += " garantia-status-ambar"
    detalhe_html = f" — {escape(detalhe)}" if detalhe else ""
    st.markdown(
        f'<div class="{classe}"><strong>{escape(titulo)}</strong>{detalhe_html}</div>',
        unsafe_allow_html=True,
    )


def linha_resumo(rotulo, valor, nota=None, destaque=False):
    """Uma leitura do painel-resumo: rótulo (+ nota) à esquerda, valor à direita.

    O ``flex-wrap`` faz o valor descer para a própria linha em larguras menores —
    nunca corta o número nem cria rolagem horizontal.
    """
    classe = "garantia-resumo-linha destaque" if destaque else "garantia-resumo-linha"
    nota_html = f'<div class="garantia-resumo-nota">{escape(nota)}</div>' if nota else ""
    return (
        f'<div class="{classe}">'
        f'<div class="garantia-resumo-texto">'
        f'<div class="garantia-resumo-rotulo">{escape(rotulo)}</div>'
        f"{nota_html}"
        f"</div>"
        f'<div class="garantia-resumo-valor">{escape(valor)}</div>'
        f"</div>"
    )


def render_resumo_situacao_atual(situacao, analise):
    """Painel único: a conclusão executiva da análise em cinco leituras.

    Não há cálculo aqui — cada linha apenas FORMATA um valor que o motor já
    apurou em ``situacao``/``analise``. Os três resultados práticos (garantia
    exigida, complemento e validade mínima) recebem o destaque verde suave; o
    valor do contrato e a última garantia apresentada ficam como contexto.
    """
    tem_garantia = analise["tem_garantia"]
    complemento = analise["complemento"]
    linhas = [
        linha_resumo(
            "Valor atualizado total do contrato",
            formatar_brl(situacao["valor_atual"]),
            formatar_variacao(situacao["variacao_acumulada"]) + " frente ao valor original.",
        ),
        linha_resumo(
            "Valor atualizado total da garantia exigida",
            formatar_brl(situacao["garantia_exigida"]),
            f"{formatar_percentual(situacao['percentual'])}% do valor atualizado total do contrato.",
            destaque=True,
        ),
        linha_resumo(
            "Valor da última garantia apresentada",
            # Ausência não é zero: sem garantia informada a linha diz "Não
            # informada", jamais R$ 0,00.
            formatar_brl(analise["garantia_apresentada"]) if tem_garantia else "Não informada",
            "Última fotografia da garantia na linha do tempo."
            if tem_garantia
            else "Nenhuma garantia apresentada foi informada na linha do tempo.",
        ),
        linha_resumo(
            "Valor do complemento necessário da garantia",
            formatar_brl(complemento),
            "Não há complemento financeiro necessário."
            if complemento == 0
            else "Diferença entre a garantia exigida e a última garantia apresentada.",
            destaque=True,
        ),
        linha_resumo(
            "Validade mínima exigida da garantia",
            formatar_data_br(analise["validade_minima"]),
            f"{DIAS_VALIDADE_MINIMA} dias corridos após o término da vigência "
            f"({formatar_data_br(situacao['vigencia_atual'])}). Validade da garantia "
            f"apresentada: {formatar_data_br(analise['validade_apresentada'])}.",
            destaque=True,
        ),
    ]
    st.markdown(
        f'<div class="garantia-resumo">{"".join(linhas)}</div>',
        unsafe_allow_html=True,
    )


# ============================================================
# Interface
# ============================================================

css()
render_cabecalho_pagina("Garantia Contratual", "")
# Etapa 29C.1: ciclo de vida da origem em dois estados. A ponte de sessão
# "origem_navegacao_garantia" (gravada pela Central de Modelos) é consumida
# por pop já no primeiro render e convertida no query parameter exclusivo
# "origem_garantia", que marca apenas a visita atual: sobrevive a reruns e ao
# reload da própria Garantia, mas é encerrado pela navegação normal entre
# páginas (sidebar/switch_page). Assim não existe origem residual em visitas
# futuras; sem ponte e sem parâmetro, o retorno é Upload e docs.
_PARAM_ORIGEM_GARANTIA = "origem_garantia"
if st.session_state.pop("origem_navegacao_garantia", None) == "modelos_ferramentas":
    st.query_params[_PARAM_ORIGEM_GARANTIA] = "modelos_ferramentas"
if st.query_params.get(_PARAM_ORIGEM_GARANTIA) == "modelos_ferramentas":
    _destino_voltar_garantia = "pages/14_Central_Modelos_Ferramentas.py"
else:
    _destino_voltar_garantia = "pages/03_Valor_Global.py"
if st.button("← Voltar para Central", key="voltar_central_garantia"):
    if _PARAM_ORIGEM_GARANTIA in st.query_params:
        del st.query_params[_PARAM_ORIGEM_GARANTIA]
    st.switch_page(_destino_voltar_garantia)

# ------------------------------------------------------------
# 1) Situação original do contrato — o marco da assinatura
# ------------------------------------------------------------
st.subheader("Situação original do contrato")
col_o1, col_o2, col_o3 = st.columns(3)
with col_o1:
    valor_original_txt = st.text_input(
        "Valor original do contrato",
        value="",
        placeholder="Ex.: 1.000.000,00",
        key="garantia_valor_original",
    ).strip()
with col_o2:
    percentual_pct = st.number_input(
        "Percentual da garantia (%)",
        min_value=0.01,
        max_value=100.0,
        value=float(PERCENTUAL_GARANTIA_PADRAO),
        step=0.25,
        format="%.2f",
        key="garantia_percentual",
    )
with col_o3:
    vigencia_original = st.date_input(
        "Término da vigência original",
        value=None,
        format="DD/MM/YYYY",
        key="garantia_fim_vigencia",
    )

col_o4, col_o5 = st.columns(2)
with col_o4:
    garantia_original_txt = st.text_input(
        "Garantia apresentada na assinatura",
        value="",
        placeholder="Opcional. Ex.: 50.000,00",
        key="garantia_apresentada_original",
    ).strip()
with col_o5:
    validade_original = st.date_input(
        "Validade da garantia",
        value=None,
        format="DD/MM/YYYY",
        key="garantia_validade_original",
    )

valor_original = parse_moeda_br(valor_original_txt) if valor_original_txt else None
garantia_original = parse_moeda_br(garantia_original_txt) if garantia_original_txt else None
_original_valido = valor_original is not None and valor_original > 0

col_m1, col_m2, col_m3 = st.columns(3)
with col_m1:
    card("Valor original", formatar_brl(valor_original) if _original_valido else TRACO)
with col_m2:
    card(
        "Garantia original exigida",
        formatar_brl(calcular_garantia_necessaria(valor_original, percentual_pct))
        if _original_valido
        else TRACO,
        f"{formatar_percentual(percentual_pct)}% do valor original.",
    )
with col_m3:
    card(
        "Validade mínima original",
        formatar_data_br(calcular_validade_minima(vigencia_original)),
        f"{DIAS_VALIDADE_MINIMA} dias corridos após o término da vigência.",
    )

pendencias = []
if valor_original is None:
    if valor_original_txt:
        st.warning(
            f'O valor "{valor_original_txt}" não pôde ser interpretado. Use o formato R$ 1.000.000,00.'
        )
    pendencias.append("informe o **valor original do contrato**")
elif valor_original <= 0:
    st.warning("O valor original do contrato deve ser maior que zero.")
    pendencias.append("corrija o **valor original do contrato**")
if vigencia_original is None:
    pendencias.append("informe o **término da vigência original**")
validade_sem_garantia = False
if garantia_original_txt and garantia_original is None:
    st.warning(
        f'A garantia "{garantia_original_txt}" não pôde ser interpretada. Use o formato R$ 50.000,00.'
    )
    pendencias.append("corrija a **garantia apresentada na assinatura**")
elif garantia_original is None and validade_original is not None:
    # Validade sozinha não descreve uma garantia: mesma regra das alterações
    # posteriores. Garantia e validade ambas vazias seguem válidas — significam
    # que não havia garantia apresentada na assinatura.
    st.warning("Informe também o valor da garantia apresentada na assinatura.")
    validade_sem_garantia = True

if pendencias:
    # A conclusão anterior não pode sobreviver a uma entrada que deixou de
    # fechar: session_state persiste entre reruns e o Saneador leria um
    # resultado que a tela já não sustenta.
    st.session_state.pop("resultado_garantia", None)
    st.info("Para montar a evolução do contrato: " + "; ".join(pendencias) + ".")
    st.stop()

if validade_sem_garantia:
    # Fail-closed com o aviso específico já exibido acima; nenhum alerta global.
    st.session_state.pop("resultado_garantia", None)
    st.stop()

# ------------------------------------------------------------
# 2) Alterações posteriores — contrato e garantia na mesma linha
# ------------------------------------------------------------
# VTA-POT-1: orientação, não acoplamento. A página segue 100% manual e sem ler
# o VTA; apenas avisa que, em contratos apurados por Pedidos de Compra, o valor
# atualizado do contrato já embute a parcela potencial incorporada por critério
# prudencial — para que o fiscal não digite um valor sem ela.
st.caption(
    "Em contratos apurados por Pedidos de Compra, o Valor Total Atualizado "
    "informado pela apuração já inclui o retroativo potencial incorporado por "
    "critério prudencial. Ao digitar aqui o valor atualizado do contrato, "
    "considere essa parcela."
)

st.subheader("Alterações posteriores à assinatura")
st.caption(
    "Informe, em ordem cronológica, o que aconteceu depois da assinatura e como ficou a garantia "
    "após cada evento. O valor do contrato e a garantia são sempre os TOTAIS vigentes após o "
    "evento: as linhas não se somam. Sem garantia informada, a anterior permanece vigente."
)
eventos_padrao = pd.DataFrame(
    {
        COLUNA_EVENTO_TIPO: pd.Series([None, None, None], dtype="object"),
        COLUNA_EVENTO_DATA: pd.Series([pd.NaT, pd.NaT, pd.NaT], dtype="datetime64[ns]"),
        COLUNA_EVENTO_VALOR: pd.Series(["", "", ""], dtype="object"),
        COLUNA_EVENTO_VIGENCIA: pd.Series([pd.NaT, pd.NaT, pd.NaT], dtype="datetime64[ns]"),
        COLUNA_EVENTO_GARANTIA: pd.Series(["", "", ""], dtype="object"),
        COLUNA_EVENTO_VALIDADE: pd.Series([pd.NaT, pd.NaT, pd.NaT], dtype="datetime64[ns]"),
    }
)
eventos_editados = st.data_editor(
    eventos_padrao,
    hide_index=True,
    use_container_width=True,
    num_rows="dynamic",
    key="garantia_eventos_contrato",
    column_config={
        COLUNA_EVENTO_TIPO: st.column_config.SelectboxColumn(
            COLUNA_EVENTO_TIPO, options=list(TIPOS_EVENTO), width="medium"
        ),
        COLUNA_EVENTO_DATA: st.column_config.DateColumn(
            COLUNA_EVENTO_DATA, format="DD/MM/YYYY", width="small"
        ),
        COLUNA_EVENTO_VALOR: st.column_config.TextColumn(COLUNA_EVENTO_VALOR, width="medium"),
        COLUNA_EVENTO_VIGENCIA: st.column_config.DateColumn(
            COLUNA_EVENTO_VIGENCIA, format="DD/MM/YYYY", width="small"
        ),
        COLUNA_EVENTO_GARANTIA: st.column_config.TextColumn(COLUNA_EVENTO_GARANTIA, width="medium"),
        COLUNA_EVENTO_VALIDADE: st.column_config.DateColumn(
            COLUNA_EVENTO_VALIDADE, format="DD/MM/YYYY", width="small"
        ),
    },
)
registros_eventos = eventos_editados.to_dict("records") if isinstance(eventos_editados, pd.DataFrame) else []
eventos, avisos_eventos, pendencias_eventos = normalizar_eventos(registros_eventos)
for aviso in avisos_eventos:
    st.warning(aviso)

situacao = calcular_situacao_atual(
    valor_original=valor_original,
    percentual=percentual_pct,
    fim_vigencia_original=vigencia_original,
    eventos=eventos,
    garantia_original=garantia_original,
    validade_garantia_original=validade_original,
)

st.markdown("**Evolução do contrato e da garantia**")
render_evolucao_contrato(situacao)

# Fail-closed: entrada materialmente incompleta ≠ dado inexistente. Uma linha
# preenchida que não pôde ser considerada deixa a história sabidamente
# incompleta; concluir mesmo assim afirmaria um resultado ignorando um evento
# que o usuário declarou existir. O aviso da própria linha já está na tela —
# nenhum alerta global é acrescentado. Linha totalmente vazia não gera pendência.
if pendencias_eventos:
    st.session_state.pop("resultado_garantia", None)
    st.stop()

# ------------------------------------------------------------
# 3) Situação atual — 100% derivada da linha do tempo
# ------------------------------------------------------------
st.subheader("Situação atual do contrato")

analise = analisar_garantia(
    valor_total_contrato=situacao["valor_atual"],
    percentual=situacao["percentual"],
    data_fim_vigencia=situacao["vigencia_atual"],
    garantia_apresentada=situacao["garantia_apresentada"],
    validade_apresentada=situacao["validade_apresentada"],
)

# Painel único, na ordem em que a pergunta se faz: quanto vale hoje o contrato,
# quanto passa a ser exigido de garantia, o que já foi apresentado, o que falta
# complementar e até quando a garantia precisa valer.
render_resumo_situacao_atual(situacao, analise)

# ------------------------------------------------------------
# 4) Resultado — a qualificação das duas dimensões independentes
# ------------------------------------------------------------
# Os números já estão no painel acima; aqui fica o que ele não diz: como a
# garantia se qualifica no dinheiro e no prazo, e o diagnóstico final.
st.subheader("Resultado da análise")
col_s1, col_s2 = st.columns(2)
with col_s1:
    st.markdown("**Situação financeira**")
    if analise["situacao_financeira"] == FINANCEIRO_COMPLEMENTAR:
        status_box(
            FINANCEIRO_COMPLEMENTAR,
            f"complemento de {formatar_brl(analise['complemento'])}.",
        )
    elif analise["situacao_financeira"] == FINANCEIRO_SUFICIENTE:
        status_box(FINANCEIRO_SUFICIENTE)
    else:
        status_box(
            analise["situacao_financeira"],
            "não há complemento financeiro a exigir. Eventual adequação depende da "
            "análise contratual.",
        )
with col_s2:
    st.markdown("**Situação da validade**")
    if analise["situacao_temporal"] == TEMPORAL_SUFICIENTE:
        status_box(TEMPORAL_SUFICIENTE)
    elif analise["situacao_temporal"] == TEMPORAL_NAO_INFORMADA:
        status_box(
            TEMPORAL_NAO_INFORMADA,
            "o prazo não pôde ser verificado com os dados apresentados.",
        )
    else:
        status_box(
            analise["situacao_temporal"],
            f"a validade deve alcançar, no mínimo, {formatar_data_br(analise['validade_minima'])}.",
        )

diagnostico = analise["diagnostico"]
if diagnostico == DIAGNOSTICO_REGULAR:
    status_box(DIAGNOSTICO_REGULAR, "valor suficiente e validade suficiente.", forte=True)
elif diagnostico == DIAGNOSTICO_VALOR:
    status_box(
        DIAGNOSTICO_VALOR,
        f"complemento de {formatar_brl(analise['complemento'])}.",
        forte=True,
        ambar=True,
    )
elif diagnostico == DIAGNOSTICO_VALIDADE:
    status_box(
        DIAGNOSTICO_VALIDADE,
        "o valor garantido é suficiente, mas a validade deve alcançar "
        f"{formatar_data_br(analise['validade_minima'])}.",
        forte=True,
        ambar=True,
    )
else:
    # Pendência nas DUAS dimensões segue em vermelho: bloqueio real já previsto.
    st.error(
        f"{DIAGNOSTICO_VALOR_E_VALIDADE} — complemento de {formatar_brl(analise['complemento'])} "
        f"e validade mínima até {formatar_data_br(analise['validade_minima'])}."
    )

# ------------------------------------------------------------
# 5) Texto para a contratada
# ------------------------------------------------------------
st.subheader("Texto para a contratada")
texto_comunicacao = gerar_texto_comunicacao(analise)
# ARMADILHA: st.text_area com key fixa só honra ``value=`` no PRIMEIRO render;
# depois o valor guardado em session_state prevalece e o texto congela na
# primeira apuração. O texto é reescrito sempre que a apuração muda — mesmo
# padrão já homologado na Adequação Orçamentária.
if st.session_state.get("garantia_texto_comunicacao") != texto_comunicacao:
    st.session_state["garantia_texto_comunicacao"] = texto_comunicacao
st.text_area(
    "Texto para a contratada",
    height=340,
    key="garantia_texto_comunicacao",
    label_visibility="collapsed",
)

# ------------------------------------------------------------
# Resultado próprio da página (consumido apenas como indicador de
# disponibilidade pelo Saneador; nunca realimenta os campos acima).
# ------------------------------------------------------------
st.session_state["resultado_garantia"] = {
    "valor_original": situacao["valor_original"],
    "vigencia_original": situacao["vigencia_original"],
    "variacao_acumulada": situacao["variacao_acumulada"],
    "quantidade_eventos": situacao["quantidade_eventos"],
    "linha_do_tempo": situacao["linha_do_tempo"],
    "valor_total_contrato": analise["valor_total_contrato"],
    "percentual_garantia": analise["percentual"],
    "garantia_necessaria": analise["garantia_necessaria"],
    "garantia_apresentada": analise["garantia_apresentada"],
    "tem_garantia": analise["tem_garantia"],
    "cobertura_atual": analise["cobertura_atual"],
    "complemento": analise["complemento"],
    "data_fim_vigencia": analise["data_fim_vigencia"],
    "validade_minima": analise["validade_minima"],
    "validade_apresentada": analise["validade_apresentada"],
    "valor_suficiente": analise["valor_suficiente"],
    "validade_suficiente": analise["validade_suficiente"],
    "situacao_financeira": analise["situacao_financeira"],
    "situacao_temporal": analise["situacao_temporal"],
    "diagnostico": analise["diagnostico"],
    "texto_comunicacao": texto_comunicacao,
}
