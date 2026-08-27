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
st.caption("Informe os dados do contrato na assinatura. A garantia é opcional.")
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
if garantia_original_txt and garantia_original is None:
    st.warning(
        f'A garantia "{garantia_original_txt}" não pôde ser interpretada. Use o formato R$ 50.000,00.'
    )
    pendencias.append("corrija a **garantia apresentada na assinatura**")

if pendencias:
    # A conclusão anterior não pode sobreviver a uma entrada que deixou de
    # fechar: session_state persiste entre reruns e o Saneador leria um
    # resultado que a tela já não sustenta.
    st.session_state.pop("resultado_garantia", None)
    st.info("Para montar a evolução do contrato: " + "; ".join(pendencias) + ".")
    st.stop()

# ------------------------------------------------------------
# 2) Alterações posteriores — contrato e garantia na mesma linha
# ------------------------------------------------------------
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
col_a1, col_a2, col_a3 = st.columns(3)
with col_a1:
    card(
        "Valor atual do contrato",
        formatar_brl(situacao["valor_atual"]),
        formatar_variacao(situacao["variacao_acumulada"]) + " frente ao valor original.",
    )
with col_a2:
    card(
        "Garantia exigida",
        formatar_brl(situacao["garantia_exigida"]),
        f"{formatar_percentual(situacao['percentual'])}% do valor atual.",
    )
with col_a3:
    card(
        "Validade mínima da garantia",
        formatar_data_br(situacao["validade_minima"]),
        f"Vigência até {formatar_data_br(situacao['vigencia_atual'])}.",
    )

analise = analisar_garantia(
    valor_total_contrato=situacao["valor_atual"],
    percentual=situacao["percentual"],
    data_fim_vigencia=situacao["vigencia_atual"],
    garantia_apresentada=situacao["garantia_apresentada"],
    validade_apresentada=situacao["validade_apresentada"],
)

# ------------------------------------------------------------
# 4) Resultado — exigida x última fotografia, nas duas dimensões
# ------------------------------------------------------------
st.subheader("Resultado da análise")
col_r1, col_r2, col_r3 = st.columns(3)
with col_r1:
    card(
        "Garantia apresentada",
        formatar_brl_opcional(analise["garantia_apresentada"]),
        "Última situação informada na linha do tempo."
        if analise["tem_garantia"]
        else "Nenhuma garantia informada.",
    )
with col_r2:
    card(
        "Validade apresentada",
        formatar_data_br(analise["validade_apresentada"]),
        f"Mínima necessária: {formatar_data_br(analise['validade_minima'])}.",
    )
with col_r3:
    card(
        "Complemento financeiro necessário",
        formatar_brl(analise["complemento"]),
        "Não há complemento financeiro a exigir."
        if analise["valor_suficiente"]
        else "Diferença frente à garantia exigida.",
        destaque=not analise["valor_suficiente"],
    )

col_s1, col_s2 = st.columns(2)
with col_s1:
    st.markdown("**Situação financeira**")
    if analise["situacao_financeira"] == FINANCEIRO_COMPLEMENTAR:
        st.warning(
            f"{FINANCEIRO_COMPLEMENTAR} — complemento de {formatar_brl(analise['complemento'])}."
        )
    elif analise["situacao_financeira"] == FINANCEIRO_SUFICIENTE:
        st.success(f"{FINANCEIRO_SUFICIENTE}.")
    else:
        st.success(
            f"{analise['situacao_financeira']} — não há complemento financeiro a exigir. "
            "Eventual adequação depende da análise contratual."
        )
with col_s2:
    st.markdown("**Situação da validade**")
    if analise["situacao_temporal"] == TEMPORAL_SUFICIENTE:
        st.success(f"{TEMPORAL_SUFICIENTE}.")
    elif analise["situacao_temporal"] == TEMPORAL_NAO_INFORMADA:
        st.warning(
            f"{TEMPORAL_NAO_INFORMADA} — o prazo não pôde ser verificado com os dados apresentados."
        )
    else:
        st.warning(
            f"{analise['situacao_temporal']} — a validade deve alcançar, no mínimo, "
            f"{formatar_data_br(analise['validade_minima'])}."
        )

diagnostico = analise["diagnostico"]
if diagnostico == DIAGNOSTICO_REGULAR:
    st.success(f"{DIAGNOSTICO_REGULAR} — valor suficiente e validade suficiente.")
elif diagnostico == DIAGNOSTICO_VALOR:
    st.warning(f"{DIAGNOSTICO_VALOR} — complemento de {formatar_brl(analise['complemento'])}.")
elif diagnostico == DIAGNOSTICO_VALIDADE:
    st.warning(
        f"{DIAGNOSTICO_VALIDADE} — o valor garantido é suficiente, mas a validade deve alcançar "
        f"{formatar_data_br(analise['validade_minima'])}."
    )
else:
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
    "Copie e cole no e-mail à contratada",
    height=340,
    key="garantia_texto_comunicacao",
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
