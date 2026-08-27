"""Garantia Contratual — linha do tempo contratual, 100% manual.

A página conta a evolução do contrato em ordem cronológica e só então a confronta
com o que a contratada apresentou:

    SITUAÇÃO ORIGINAL -> ALTERAÇÕES POSTERIORES -> SITUAÇÃO ATUAL
        -> GARANTIA ATUALMENTE APRESENTADA -> RESULTADO -> TEXTO À CONTRATADA

Dois conceitos que não se misturam: reajuste, repactuação, aditivo e prorrogação
são EVENTOS DO CONTRATO (definem a garantia exigida); apólice e endosso são
GARANTIA APRESENTADA (definem a cobertura existente).

Ferramenta 100% MANUAL: todos os dados são digitados aqui. A página não lê o
VTA, o Valor Global, a Coleta, os RESULTADOS, o XLS nem qualquer outra chave de
sessão de outra página — mesmo que esses dados existam na sessão, são ignorados.
O único uso de ``st.session_state`` é o estado da própria página: as duas grades,
a navegação de retorno e o resultado próprio publicado ao final.

Toda a matemática vive no motor puro ``_garantia_calculo`` (Decimal +
ROUND_HALF_UP), permitindo testes focais.
"""
from html import escape

import pandas as pd
import streamlit as st

from _ui_utils import render_cabecalho_pagina
from _garantia_calculo import (
    COLUNA_EVENTO_DATA,
    COLUNA_EVENTO_OBSERVACAO,
    COLUNA_EVENTO_TIPO,
    COLUNA_EVENTO_VALOR,
    COLUNA_EVENTO_VIGENCIA,
    COLUNA_REFERENCIA,
    COLUNA_VALIDADE,
    COLUNA_VALOR,
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
    formatar_data_br,
    formatar_percentual,
    formatar_variacao,
    gerar_texto_comunicacao,
    normalizar_eventos,
    normalizar_garantias,
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
        table.garantia-tabela td.suficiente { color: #1E6B45; font-weight: 700; }
        table.garantia-tabela td.insuficiente { color: #A2432B; font-weight: 700; }
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


def render_tabela(cabecalhos, linhas_html, larguras=None):
    """Tabela de leitura, não editável, no padrão visual da página."""
    colgroup = ""
    if larguras:
        colunas = "".join(f'<col style="width: {largura};">' for largura in larguras)
        colgroup = f"<colgroup>{colunas}</colgroup>"
    cabecalho = "".join(f"<th>{escape(titulo)}</th>" for titulo in cabecalhos)
    corpo = "".join(linhas_html)
    st.markdown(
        f"""
        <div class="garantia-tabela-wrap">
          <table class="garantia-tabela">
            {colgroup}
            <thead><tr>{cabecalho}</tr></thead>
            <tbody>{corpo}</tbody>
          </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _celula(texto, classe=""):
    atributo = f" class='{classe}'" if classe else ""
    return f"<td{atributo}>{escape(texto)}</td>"


def render_evolucao_contrato(situacao):
    """Memória objetiva da evolução: o marco original e cada evento posterior."""
    linhas = [
        "<tr class='garantia-linha-marco'>"
        + _celula("0")
        + _celula("Assinatura (situação original)")
        + _celula(TRACO, "valor")
        + _celula(formatar_brl(situacao["valor_original"]), "valor")
        + _celula(TRACO, "valor")
        + _celula(formatar_brl(situacao["garantia_original"]), "valor")
        + _celula(formatar_data_br(situacao["vigencia_original"]), "valor")
        + _celula(formatar_data_br(situacao["validade_minima_original"]), "valor")
        + "</tr>"
    ]
    for etapa in situacao["linha_do_tempo"]:
        evento = etapa["tipo"]
        if etapa["observacao"]:
            evento = f"{evento} — {etapa['observacao']}"
        linhas.append(
            "<tr>"
            + _celula(str(etapa["numero"]))
            + _celula(evento)
            + _celula(formatar_data_br(etapa["data"]), "valor")
            + _celula(formatar_brl(etapa["valor"]), "valor")
            + _celula(formatar_variacao(etapa["variacao"]), "valor")
            + _celula(formatar_brl(etapa["garantia_exigida"]), "valor")
            + _celula(formatar_data_br(etapa["vigencia"]), "valor")
            + _celula(formatar_data_br(etapa["validade_minima"]), "valor")
            + "</tr>"
        )
    render_tabela(
        [
            "Nº",
            "Evento",
            "Data",
            "Valor do contrato",
            "Variação",
            "Garantia exigida",
            "Término da vigência",
            "Validade mínima",
        ],
        linhas,
        ["4%", "24%", "10%", "14%", "13%", "13%", "11%", "11%"],
    )


def render_tabela_validades(garantias, validade_minima):
    """Validade de cada garantia apresentada frente à validade mínima exigida."""
    linhas = []
    for indice, garantia in enumerate(garantias, start=1):
        referencia = garantia["referencia"] or f"Garantia {indice}"
        if garantia["validade_suficiente"]:
            situacao, classe = "Validade suficiente", "suficiente"
        elif garantia["validade"] is None:
            situacao, classe = "Validade não informada", "insuficiente"
        else:
            situacao, classe = f"Vence antes de {formatar_data_br(validade_minima)}", "insuficiente"
        linhas.append(
            "<tr>"
            + _celula(referencia)
            + _celula(formatar_brl(garantia["valor"]), "valor")
            + _celula(formatar_data_br(garantia["validade"]), "valor")
            + _celula(situacao, classe)
            + "</tr>"
        )
    render_tabela(
        ["Garantia apresentada", "Valor garantido", "Validade", "Situação da validade"],
        linhas,
        ["34%", "22%", "18%", "26%"],
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
# 1) Identificação — apenas o que a comunicação à contratada usa
# ------------------------------------------------------------
st.subheader("Identificação")
col_i1, col_i2 = st.columns(2)
with col_i1:
    numero_contrato = st.text_input(
        "Número do contrato",
        value="",
        placeholder="Ex.: 123/2024",
        help="Opcional. Aparece na referência do texto à contratada.",
        key="garantia_numero_contrato",
    ).strip()
with col_i2:
    contratada = st.text_input(
        "Contratada",
        value="",
        placeholder="Ex.: Empresa Exemplo Ltda.",
        help="Opcional. Aparece na referência do texto à contratada.",
        key="garantia_contratada",
    ).strip()

# ------------------------------------------------------------
# 2) Situação original do contrato — o marco da assinatura
# ------------------------------------------------------------
st.subheader("Situação original do contrato")
st.caption("Informe os dados do contrato na assinatura.")
col_o1, col_o2, col_o3 = st.columns(3)
with col_o1:
    valor_original_txt = st.text_input(
        "Valor original do contrato",
        value="",
        placeholder="Ex.: 1.000.000,00",
        help="Valor total do contrato na assinatura. Ex.: 1000000, 1.000.000,00 ou R$ 1.000.000,00",
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
        help="Percentual previsto no contrato. O padrão é 5,00% e pode ser alterado.",
        key="garantia_percentual",
    )
with col_o3:
    vigencia_original = st.date_input(
        "Término da vigência original",
        value=None,
        format="DD/MM/YYYY",
        help="Data final da vigência prevista na assinatura.",
        key="garantia_fim_vigencia",
    )

valor_original = parse_moeda_br(valor_original_txt) if valor_original_txt else None
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

if pendencias:
    st.info("Para montar a evolução do contrato: " + "; ".join(pendencias) + ".")
    st.stop()

# ------------------------------------------------------------
# 3) Alterações posteriores à assinatura — eventos do contrato
# ------------------------------------------------------------
st.subheader("Alterações posteriores à assinatura")
st.caption(
    "Informe, em ordem cronológica, os eventos ocorridos após a assinatura que alteraram o valor ou "
    "a vigência do contrato. Os primeiros registros devem representar os eventos mais antigos. "
    "Informe sempre o valor TOTAL do contrato após o evento: a variação é calculada automaticamente."
)
eventos_padrao = pd.DataFrame(
    {
        COLUNA_EVENTO_TIPO: pd.Series([None, None, None], dtype="object"),
        COLUNA_EVENTO_DATA: pd.Series([pd.NaT, pd.NaT, pd.NaT], dtype="datetime64[ns]"),
        COLUNA_EVENTO_VALOR: pd.Series(["", "", ""], dtype="object"),
        COLUNA_EVENTO_VIGENCIA: pd.Series([pd.NaT, pd.NaT, pd.NaT], dtype="datetime64[ns]"),
        COLUNA_EVENTO_OBSERVACAO: pd.Series(["", "", ""], dtype="object"),
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
            COLUNA_EVENTO_TIPO,
            options=list(TIPOS_EVENTO),
            help="O que aconteceu com o contrato. A numeração é automática, pela ordem das linhas.",
            width="medium",
        ),
        COLUNA_EVENTO_DATA: st.column_config.DateColumn(
            COLUNA_EVENTO_DATA,
            help="Opcional. Serve à memória do processo e não bloqueia o cálculo.",
            format="DD/MM/YYYY",
            width="small",
        ),
        COLUNA_EVENTO_VALOR: st.column_config.TextColumn(
            COLUNA_EVENTO_VALOR,
            help="Valor total do contrato depois do evento — não o acréscimo. Ex.: 1.100.000,00",
            width="medium",
        ),
        COLUNA_EVENTO_VIGENCIA: st.column_config.DateColumn(
            COLUNA_EVENTO_VIGENCIA,
            help="Obrigatório na prorrogação; opcional nos demais eventos.",
            format="DD/MM/YYYY",
            width="small",
        ),
        COLUNA_EVENTO_OBSERVACAO: st.column_config.TextColumn(
            COLUNA_EVENTO_OBSERVACAO,
            help="Opcional. Ex.: 1º Termo Aditivo, apostilamento IPCA.",
            width="medium",
        ),
    },
)
registros_eventos = eventos_editados.to_dict("records") if isinstance(eventos_editados, pd.DataFrame) else []
eventos, avisos_eventos = normalizar_eventos(registros_eventos)
for aviso in avisos_eventos:
    st.warning(aviso)

situacao = calcular_situacao_atual(
    valor_original=valor_original,
    percentual=percentual_pct,
    fim_vigencia_original=vigencia_original,
    eventos=eventos,
)

st.markdown("**Evolução do contrato**")
render_evolucao_contrato(situacao)

# ------------------------------------------------------------
# 4) Situação atual do contrato — 100% derivada dos blocos acima
# ------------------------------------------------------------
st.subheader("Situação atual do contrato")
st.caption(
    "Resultado da situação original com as alterações posteriores informadas acima. "
    "Nada aqui é redigitado."
)
col_a1, col_a2, col_a3 = st.columns(3)
with col_a1:
    card(
        "Valor atual do contrato",
        formatar_brl(situacao["valor_atual"]),
        f"{situacao['quantidade_eventos']} alteração(ões) considerada(s)."
        if situacao["quantidade_eventos"]
        else "Nenhuma alteração posterior informada.",
    )
with col_a2:
    card(
        "Variação acumulada",
        formatar_variacao(situacao["variacao_acumulada"]),
        "Frente ao valor original.",
    )
with col_a3:
    card(
        "Garantia total exigida",
        formatar_brl(situacao["garantia_exigida"]),
        f"{formatar_percentual(situacao['percentual'])}% do valor atual do contrato.",
    )

col_a4, col_a5 = st.columns(2)
with col_a4:
    card("Término da vigência atual", formatar_data_br(situacao["vigencia_atual"]))
with col_a5:
    card(
        "Validade mínima da garantia",
        formatar_data_br(situacao["validade_minima"]),
        f"{DIAS_VALIDADE_MINIMA} dias corridos após o término da vigência atual.",
    )

# ------------------------------------------------------------
# 5) Garantia atualmente apresentada — o que a contratada entregou
# ------------------------------------------------------------
st.subheader("Garantia atualmente apresentada")
st.caption(
    "Informe a garantia que está atualmente vigente/apresentada pela contratada para comparação com "
    "a situação atual do contrato. Uma linha por garantia independente, sempre com o valor TOTAL "
    "vigente após o último endosso: endossos da mesma garantia atualizam a linha, não criam outra."
)
garantias_padrao = pd.DataFrame(
    {
        COLUNA_REFERENCIA: pd.Series(["", ""], dtype="object"),
        COLUNA_VALOR: pd.Series(["", ""], dtype="object"),
        COLUNA_VALIDADE: pd.Series([pd.NaT, pd.NaT], dtype="datetime64[ns]"),
    }
)
garantias_editadas = st.data_editor(
    garantias_padrao,
    hide_index=True,
    use_container_width=True,
    num_rows="dynamic",
    key="garantia_vigente_linhas",
    column_config={
        COLUNA_REFERENCIA: st.column_config.TextColumn(
            COLUNA_REFERENCIA,
            help="Opcional. Ex.: Apólice 123, Endosso 2, Carta de fiança Banco X, Caução.",
            width="medium",
        ),
        COLUNA_VALOR: st.column_config.TextColumn(
            COLUNA_VALOR,
            help="Valor total vigente da garantia após o último endosso. Ex.: 40.000,00 ou R$ 40.000,00",
            width="medium",
        ),
        COLUNA_VALIDADE: st.column_config.DateColumn(
            COLUNA_VALIDADE,
            help="Opcional para o cálculo financeiro. Sem ela, o prazo fica como não informado.",
            format="DD/MM/YYYY",
            width="small",
        ),
    },
)
registros_garantias = garantias_editadas.to_dict("records") if isinstance(garantias_editadas, pd.DataFrame) else []
linhas_garantias, avisos_garantias = normalizar_garantias(registros_garantias)
for aviso in avisos_garantias:
    st.warning(aviso)

analise = analisar_garantia(
    valor_total_contrato=situacao["valor_atual"],
    percentual=situacao["percentual"],
    data_fim_vigencia=situacao["vigencia_atual"],
    garantias=linhas_garantias,
)
for aviso in analise["avisos_consolidacao"]:
    st.warning(aviso)

# ------------------------------------------------------------
# 6) Resultado — dinheiro e prazo analisados separadamente
# ------------------------------------------------------------
st.subheader("Resultado da análise")
col_r1, col_r2, col_r3 = st.columns(3)
with col_r1:
    card("Garantia exigida atualmente", formatar_brl(analise["garantia_necessaria"]))
with col_r2:
    if analise["quantidade_garantias"] == 0:
        nota_cobertura = "Nenhuma garantia apresentada."
    elif analise["quantidade_garantias"] == 1:
        nota_cobertura = "Valor total vigente da garantia apresentada."
    else:
        nota_cobertura = f"Soma de {analise['quantidade_garantias']} garantias independentes."
    card("Garantia apresentada considerada", formatar_brl(analise["cobertura_atual"]), nota_cobertura)
with col_r3:
    card(
        "Complemento financeiro necessário",
        formatar_brl(analise["complemento"]),
        "Não há complemento financeiro a exigir."
        if analise["valor_suficiente"]
        else "Diferença frente à garantia exigida.",
        destaque=not analise["valor_suficiente"],
    )

validades_apresentadas = [g["validade"] for g in analise["garantias"] if g["validade"] is not None]
if analise["quantidade_garantias"] == 0:
    validade_apresentada, nota_validade = TRACO, "Nenhuma garantia apresentada."
elif len(validades_apresentadas) < analise["quantidade_garantias"]:
    validade_apresentada, nota_validade = TRACO, "Há garantia sem validade informada."
else:
    validade_apresentada = formatar_data_br(min(validades_apresentadas))
    nota_validade = (
        "Menor validade entre as garantias apresentadas."
        if len(validades_apresentadas) > 1
        else "Validade da garantia apresentada."
    )

col_r4, col_r5 = st.columns(2)
with col_r4:
    card(
        "Validade mínima necessária",
        formatar_data_br(analise["validade_minima"]),
        f"{DIAS_VALIDADE_MINIMA} dias corridos após o término da vigência atual.",
    )
with col_r5:
    card("Validade apresentada", validade_apresentada, nota_validade)

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

st.markdown("**Validade de cada garantia apresentada**")
if analise["garantias"]:
    render_tabela_validades(analise["garantias"], analise["validade_minima"])
else:
    st.caption("Nenhuma garantia apresentada informada.")

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
# 7) Texto para a contratada
# ------------------------------------------------------------
st.subheader("Texto para a contratada")
texto_comunicacao = gerar_texto_comunicacao(
    analise,
    numero_contrato=numero_contrato,
    contratada=contratada,
)
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
    "numero_contrato": numero_contrato,
    "contratada": contratada,
    "valor_original": situacao["valor_original"],
    "vigencia_original": situacao["vigencia_original"],
    "variacao_acumulada": situacao["variacao_acumulada"],
    "quantidade_eventos": situacao["quantidade_eventos"],
    "linha_do_tempo": situacao["linha_do_tempo"],
    "valor_total_contrato": analise["valor_total_contrato"],
    "percentual_garantia": analise["percentual"],
    "garantia_necessaria": analise["garantia_necessaria"],
    "cobertura_atual": analise["cobertura_atual"],
    "complemento": analise["complemento"],
    "data_fim_vigencia": analise["data_fim_vigencia"],
    "validade_minima": analise["validade_minima"],
    "valor_suficiente": analise["valor_suficiente"],
    "validade_suficiente": analise["validade_suficiente"],
    "situacao_financeira": analise["situacao_financeira"],
    "situacao_temporal": analise["situacao_temporal"],
    "diagnostico": analise["diagnostico"],
    "garantias": [
        {
            "referencia": garantia["referencia"],
            "valor": garantia["valor"],
            "validade": garantia["validade"],
            "validade_suficiente": garantia["validade_suficiente"],
        }
        for garantia in analise["garantias"]
    ],
    "texto_comunicacao": texto_comunicacao,
}
