"""Garantia Contratual — calculadora manual de fluxo único (Etapa 49).

A página compara a SITUAÇÃO ATUAL DO CONTRATO com a GARANTIA ATUALMENTE
CONSTITUÍDA e responde se falta valor, se falta prazo, se faltam os dois ou se
não há nada a regularizar.

Ferramenta 100% MANUAL: todos os dados são digitados aqui. A página não lê o
VTA, o Valor Global, a Coleta, os RESULTADOS, o XLS nem qualquer outra chave de
sessão de outra página — mesmo que esses dados existam na sessão, são ignorados.
O único uso de ``st.session_state`` é o estado da própria página: a grade de
garantias, a navegação de retorno e o resultado próprio publicado ao final.

Toda a matemática vive no motor puro ``_garantia_calculo`` (Decimal +
ROUND_HALF_UP), permitindo testes focais.
"""
from html import escape

import pandas as pd
import streamlit as st

from _ui_utils import render_cabecalho_pagina
from _garantia_calculo import (
    COLUNA_IDENTIFICACAO,
    COLUNA_VALIDADE,
    COLUNA_VALOR,
    DIAGNOSTICO_REGULAR,
    DIAGNOSTICO_VALIDADE,
    DIAGNOSTICO_VALOR,
    DIAGNOSTICO_VALOR_E_VALIDADE,
    DIAS_VALIDADE_MINIMA,
    PERCENTUAL_GARANTIA_PADRAO,
    analisar_garantia,
    formatar_brl,
    formatar_data_br,
    formatar_percentual,
    gerar_texto_comunicacao,
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
            padding: 18px 20px;
            margin: 6px 0 14px 0;
        }
        .garantia-card-destaque {
            background: #EAF2F8;
            border: 1px solid #C8D9E8;
            border-radius: 14px;
            padding: 20px 22px;
            margin: 8px 0 16px 0;
        }
        .garantia-label { color: #475569; font-size: 0.92rem; margin-bottom: 4px; }
        .garantia-valor { color: #1F2937; font-size: 1.55rem; font-weight: 700; line-height: 1.2; }
        .garantia-valor-destaque { color: #123B63; font-size: 2rem; font-weight: 800; line-height: 1.2; }
        .garantia-nota { color: #64748B; font-size: 0.88rem; margin-top: 6px; }
        .garantia-tabela-wrap { width: 100%; overflow-x: auto; margin: 10px 0 18px 0; }
        table.garantia-tabela { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 0.92rem; }
        table.garantia-tabela th { background: #E6F0F7; color: #173B5D; border: 1px solid #C5D6E2; padding: 9px 10px; text-align: left; font-weight: 700; }
        table.garantia-tabela td { border: 1px solid #E5EAF0; padding: 9px 10px; vertical-align: top; white-space: normal; overflow-wrap: anywhere; word-break: normal; line-height: 1.35; }
        table.garantia-tabela td.valor { white-space: nowrap; overflow-wrap: normal; text-align: right; }
        table.garantia-tabela td.suficiente { color: #1E6B45; font-weight: 700; }
        table.garantia-tabela td.insuficiente { color: #A2432B; font-weight: 700; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def card(label, valor, nota=None, destaque=False):
    classe = "garantia-card-destaque" if destaque else "garantia-card"
    valor_classe = "garantia-valor-destaque" if destaque else "garantia-valor"
    nota_html = f'<div class="garantia-nota">{nota}</div>' if nota else ""
    st.markdown(
        f"""
        <div class="{classe}">
            <div class="garantia-label">{label}</div>
            <div class="{valor_classe}">{valor}</div>
            {nota_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_tabela_validades(garantias, validade_minima):
    """Validade de cada garantia frente à validade mínima exigida."""
    linhas = []
    for indice, garantia in enumerate(garantias, start=1):
        identificacao = escape(garantia["identificacao"] or f"Garantia {indice}")
        valor = escape(formatar_brl(garantia["valor"]))
        validade = escape(formatar_data_br(garantia["validade"]))
        if garantia["validade_suficiente"]:
            situacao, classe = "Validade suficiente", "suficiente"
        elif garantia["validade"] is None:
            situacao, classe = "Validade não informada", "insuficiente"
        else:
            situacao, classe = f"Vence antes de {formatar_data_br(validade_minima)}", "insuficiente"
        linhas.append(
            f"<tr><td>{identificacao}</td><td class='valor'>{valor}</td>"
            f"<td class='valor'>{validade}</td><td class='{classe}'>{escape(situacao)}</td></tr>"
        )
    html = """
    <div class="garantia-tabela-wrap">
      <table class="garantia-tabela">
        <colgroup>
          <col style="width: 34%;">
          <col style="width: 22%;">
          <col style="width: 18%;">
          <col style="width: 26%;">
        </colgroup>
        <thead>
          <tr><th>Garantia</th><th>Valor total garantido</th><th>Validade</th><th>Situação da validade</th></tr>
        </thead>
        <tbody>
          {linhas}
        </tbody>
      </table>
    </div>
    """.format(linhas="\n".join(linhas))
    st.markdown(html, unsafe_allow_html=True)


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
# 1) Garantia vigente (uma linha por garantia independente)
# ------------------------------------------------------------
st.subheader("Garantia vigente")
st.caption(
    "Uma linha por garantia independente. Informe sempre o valor TOTAL atualmente garantido, já "
    "considerando eventuais endossos: endossos da mesma garantia atualizam a linha, não criam outra. "
    "Duas linhas com a mesma identificação são tratadas como a mesma garantia e a última prevalece."
)
garantias_padrao = pd.DataFrame(
    {
        COLUNA_IDENTIFICACAO: pd.Series(["", ""], dtype="object"),
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
        COLUNA_IDENTIFICACAO: st.column_config.TextColumn(
            COLUNA_IDENTIFICACAO,
            help="Opcional. Ex.: Apólice 123, Carta de fiança Banco X, Caução.",
            width="medium",
        ),
        COLUNA_VALOR: st.column_config.TextColumn(
            COLUNA_VALOR,
            help="Valor total vigente da garantia após o último endosso. Ex.: 40.000,00 ou R$ 40.000,00",
            width="medium",
        ),
        COLUNA_VALIDADE: st.column_config.DateColumn(
            COLUNA_VALIDADE,
            help="Data até a qual a garantia é válida.",
            format="DD/MM/YYYY",
            width="small",
        ),
    },
)
registros_garantias = garantias_editadas.to_dict("records") if isinstance(garantias_editadas, pd.DataFrame) else []
linhas_garantias, avisos_garantias = normalizar_garantias(registros_garantias)
for aviso in avisos_garantias:
    st.warning(aviso)

# ------------------------------------------------------------
# 2) Situação atual do contrato (campos exclusivamente manuais)
# ------------------------------------------------------------
st.subheader("Situação atual do contrato")
col_c1, col_c2, col_c3 = st.columns(3)
with col_c1:
    valor_contrato_txt = st.text_input(
        "Valor total atual do contrato",
        value="",
        placeholder="Ex.: 1.000.000,00",
        help="Valor total consolidado do contrato hoje. Ex.: 1000000, 1.000.000,00 ou R$ 1.000.000,00",
        key="garantia_valor_total_contrato",
    ).strip()
with col_c2:
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
with col_c3:
    data_fim_vigencia = st.date_input(
        "Término da vigência contratual",
        value=None,
        format="DD/MM/YYYY",
        help="Data final da vigência do contrato.",
        key="garantia_fim_vigencia",
    )

valor_contrato = parse_moeda_br(valor_contrato_txt) if valor_contrato_txt else None

pendencias = []
if valor_contrato is None:
    if valor_contrato_txt:
        st.warning(
            f'O valor "{valor_contrato_txt}" não pôde ser interpretado. Use o formato R$ 1.000.000,00.'
        )
    pendencias.append("informe o **valor total atual do contrato**")
elif valor_contrato <= 0:
    st.warning("O valor total atual do contrato deve ser maior que zero.")
    pendencias.append("corrija o **valor total atual do contrato**")
if data_fim_vigencia is None:
    pendencias.append("informe o **término da vigência contratual**")

if pendencias:
    st.info("Para calcular a garantia: " + "; ".join(pendencias) + ".")
    st.stop()

analise = analisar_garantia(
    valor_total_contrato=valor_contrato,
    percentual=percentual_pct,
    data_fim_vigencia=data_fim_vigencia,
    garantias=linhas_garantias,
)
for aviso in analise["avisos_consolidacao"]:
    st.warning(aviso)

# ------------------------------------------------------------
# 3) Resultado
# ------------------------------------------------------------
st.subheader("Resultado")
col_r1, col_r2, col_r3 = st.columns(3)
with col_r1:
    card("Valor total atual do contrato", formatar_brl(analise["valor_total_contrato"]))
with col_r2:
    card(
        "Garantia necessária",
        formatar_brl(analise["garantia_necessaria"]),
        f"{formatar_percentual(analise['percentual'])}% do valor total atual do contrato.",
    )
with col_r3:
    if analise["quantidade_garantias"] == 0:
        nota_cobertura = "Nenhuma garantia vigente informada."
    elif analise["quantidade_garantias"] == 1:
        nota_cobertura = "Valor total vigente da garantia informada."
    else:
        nota_cobertura = f"Soma de {analise['quantidade_garantias']} garantias independentes."
    card("Cobertura atualmente apresentada", formatar_brl(analise["cobertura_atual"]), nota_cobertura)

col_r4, col_r5, col_r6 = st.columns(3)
with col_r4:
    card(
        "Complementação necessária",
        formatar_brl(analise["complemento"]),
        "Não há complementação financeira a exigir."
        if analise["valor_suficiente"]
        else "Diferença frente à garantia necessária.",
        destaque=not analise["valor_suficiente"],
    )
with col_r5:
    card("Término da vigência contratual", formatar_data_br(analise["data_fim_vigencia"]))
with col_r6:
    card(
        "Validade mínima exigida",
        formatar_data_br(analise["validade_minima"]),
        f"{DIAS_VALIDADE_MINIMA} dias corridos após o término da vigência.",
        destaque=not analise["validade_suficiente"],
    )

st.markdown("**Validade de cada garantia**")
if analise["garantias"]:
    render_tabela_validades(analise["garantias"], analise["validade_minima"])
else:
    st.caption("Nenhuma garantia vigente informada.")

diagnostico = analise["diagnostico"]
if diagnostico == DIAGNOSTICO_REGULAR:
    st.success(f"{DIAGNOSTICO_REGULAR} — valor suficiente e validade suficiente.")
elif diagnostico == DIAGNOSTICO_VALOR:
    st.warning(
        f"{DIAGNOSTICO_VALOR} — complementação de {formatar_brl(analise['complemento'])}. "
        "A validade das garantias está adequada."
    )
elif diagnostico == DIAGNOSTICO_VALIDADE:
    st.warning(
        f"{DIAGNOSTICO_VALIDADE} — o valor garantido é suficiente, mas a validade deve alcançar "
        f"{formatar_data_br(analise['validade_minima'])}."
    )
else:
    st.error(
        f"{DIAGNOSTICO_VALOR_E_VALIDADE} — complementação de {formatar_brl(analise['complemento'])} "
        f"e validade mínima até {formatar_data_br(analise['validade_minima'])}."
    )

# ------------------------------------------------------------
# 4) Texto para comunicação à contratada
# ------------------------------------------------------------
st.subheader("Texto para comunicação à contratada")
texto_comunicacao = gerar_texto_comunicacao(analise)
st.text_area(
    "Copie e cole no e-mail à contratada",
    value=texto_comunicacao,
    height=340,
    key="garantia_texto_comunicacao",
)

# ------------------------------------------------------------
# Resultado próprio da página (consumido apenas como indicador de
# disponibilidade pelo Saneador; nunca realimenta os campos acima).
# ------------------------------------------------------------
st.session_state["resultado_garantia"] = {
    "valor_total_contrato": analise["valor_total_contrato"],
    "percentual_garantia": analise["percentual"],
    "garantia_necessaria": analise["garantia_necessaria"],
    "cobertura_atual": analise["cobertura_atual"],
    "complemento": analise["complemento"],
    "data_fim_vigencia": analise["data_fim_vigencia"],
    "validade_minima": analise["validade_minima"],
    "valor_suficiente": analise["valor_suficiente"],
    "validade_suficiente": analise["validade_suficiente"],
    "diagnostico": analise["diagnostico"],
    "garantias": [
        {
            "identificacao": garantia["identificacao"],
            "valor": garantia["valor"],
            "validade": garantia["validade"],
            "validade_suficiente": garantia["validade_suficiente"],
        }
        for garantia in analise["garantias"]
    ],
    "texto_comunicacao": texto_comunicacao,
}
