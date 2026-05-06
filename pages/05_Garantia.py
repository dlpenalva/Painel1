from datetime import datetime
from zoneinfo import ZoneInfo

import streamlit as st

st.set_page_config(page_title="Análises de Reajustes - Garantia", layout="wide")


# ============================================================
# Utilitários
# ============================================================

def moeda(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def numero_para_input(valor):
    try:
        return float(valor)
    except Exception:
        return 0.0


def data_hora_brasilia():
    return datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")


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
        .garantia-label {
            color: #475569;
            font-size: 0.92rem;
            margin-bottom: 4px;
        }
        .garantia-valor {
            color: #1F2937;
            font-size: 1.75rem;
            font-weight: 700;
            line-height: 1.2;
        }
        .garantia-valor-destaque {
            color: #123B63;
            font-size: 2.1rem;
            font-weight: 800;
            line-height: 1.2;
        }
        .garantia-nota {
            color: #64748B;
            font-size: 0.9rem;
            margin-top: 6px;
        }
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


def montar_texto_instrucao(
    valor_original,
    percentual_garantia,
    garantia_original,
    valor_atualizado,
    garantia_constituida,
    garantia_exigida,
    endosso,
):
    percentual = percentual_garantia * 100
    if endosso > 0:
        conclusao = (
            f"Considerando a garantia atualmente constituída no valor de {moeda(garantia_constituida)}, "
            f"faz-se necessário o endosso complementar no montante de {moeda(endosso)}."
        )
    else:
        conclusao = (
            f"Considerando a garantia atualmente constituída no valor de {moeda(garantia_constituida)}, "
            "não foi identificada necessidade de endosso complementar, desde que esse valor esteja efetivamente vigente e aceito."
        )

    return f"""Considerando o valor original do contrato de {moeda(valor_original)}, a garantia contratual original correspondente a {percentual:.2f}% equivale a {moeda(garantia_original)}.

Após a atualização do valor contratual para {moeda(valor_atualizado)}, a garantia contratual exigida, calculada pelo mesmo percentual de {percentual:.2f}%, passa a ser de {moeda(garantia_exigida)}.

{conclusao}

A apuração observa a Cláusula Décima do contrato, segundo a qual a garantia contratual corresponde a 5% do valor total do contrato, devendo ser apresentada no prazo de até 5 dias úteis contados do recebimento da convocação pela TELEBRAS, prorrogável por igual período mediante solicitação justificada e aceita pela Gerência de Compras e Contratos.
"""


css()

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Garantia Contratual")
st.write(
    "Este módulo calcula a garantia contratual original, a garantia exigida após atualização do valor do contrato "
    "e o endosso complementar necessário."
)

resultado_valor_global = st.session_state.get("resultado_valor_global", {}) or {}

valor_original_padrao = numero_para_input(resultado_valor_global.get("valor_original_contrato", 0.0))
valor_atualizado_padrao = numero_para_input(
    resultado_valor_global.get(
        "valor_atualizado_contrato",
        resultado_valor_global.get("valor_global_contrato", resultado_valor_global.get("valor_global_estoque", 0.0)),
    )
)

with st.expander("Contexto importado do Valor Global", expanded=True):
    if resultado_valor_global:
        col_a, col_b = st.columns(2)
        with col_a:
            st.metric("Valor original identificado", moeda(valor_original_padrao))
        with col_b:
            st.metric("Valor atualizado identificado", moeda(valor_atualizado_padrao))
        st.caption("Dados herdados da sessão atual do módulo Valor Global. Os campos abaixo permanecem editáveis para conferência.")
    else:
        st.info(
            "Não há dados do Valor Global disponíveis na sessão atual. Informe os valores manualmente para calcular a garantia."
        )

st.subheader("Dados para cálculo")
col1, col2, col3 = st.columns(3)

with col1:
    valor_original = st.number_input(
        "Valor original do contrato",
        min_value=0.0,
        value=valor_original_padrao,
        step=1000.0,
        format="%.2f",
    )

with col2:
    percentual_garantia_pct = st.number_input(
        "Percentual da garantia (%)",
        min_value=0.0,
        max_value=100.0,
        value=5.0,
        step=0.1,
        format="%.2f",
    )

with col3:
    valor_atualizado = st.number_input(
        "Valor atualizado do contrato",
        min_value=0.0,
        value=valor_atualizado_padrao,
        step=1000.0,
        format="%.2f",
    )

percentual_garantia = percentual_garantia_pct / 100

garantia_original = valor_original * percentual_garantia
garantia_exigida = valor_atualizado * percentual_garantia

col4, col5 = st.columns(2)
with col4:
    garantia_constituida = st.number_input(
        "Garantia atualmente constituída",
        min_value=0.0,
        value=garantia_original,
        step=1000.0,
        format="%.2f",
        help="Informe o valor da garantia atualmente vigente/aceita. Por padrão, o sistema usa a garantia original.",
    )

with col5:
    prazo_dias = st.number_input(
        "Prazo contratual para apresentação/endosso (dias úteis)",
        min_value=1,
        max_value=60,
        value=5,
        step=1,
    )

endosso_necessario = max(garantia_exigida - garantia_constituida, 0.0)
excesso_garantia = max(garantia_constituida - garantia_exigida, 0.0)

st.divider()
st.subheader("Resultado")

colr1, colr2, colr3 = st.columns(3)
with colr1:
    card("Garantia original", moeda(garantia_original), f"{percentual_garantia_pct:.2f}% sobre o valor original")
with colr2:
    card("Nova garantia exigida", moeda(garantia_exigida), f"{percentual_garantia_pct:.2f}% sobre o valor atualizado")
with colr3:
    card(
        "Endosso necessário",
        moeda(endosso_necessario),
        "Diferença entre a nova garantia exigida e a garantia atualmente constituída",
        destaque=True,
    )

if endosso_necessario > 0:
    st.warning(
        f"Será necessário solicitar endosso complementar de {moeda(endosso_necessario)}, observado o prazo contratual de {prazo_dias} dias úteis."
    )
elif excesso_garantia > 0:
    st.success(
        f"A garantia atualmente constituída supera a garantia exigida em {moeda(excesso_garantia)}. Verifique se a garantia vigente está válida e aceita."
    )
else:
    st.success("A garantia atualmente constituída corresponde exatamente à garantia exigida.")

st.subheader("Memória de cálculo")
memoria = [
    {"Indicador": "Valor original do contrato", "Valor": moeda(valor_original)},
    {"Indicador": "Percentual da garantia", "Valor": f"{percentual_garantia_pct:.2f}%".replace(".", ",")},
    {"Indicador": "Garantia original", "Valor": moeda(garantia_original)},
    {"Indicador": "Valor atualizado do contrato", "Valor": moeda(valor_atualizado)},
    {"Indicador": "Nova garantia exigida", "Valor": moeda(garantia_exigida)},
    {"Indicador": "Garantia atualmente constituída", "Valor": moeda(garantia_constituida)},
    {"Indicador": "Endosso necessário", "Valor": moeda(endosso_necessario)},
]
st.dataframe(memoria, use_container_width=True, hide_index=True)

st.subheader("Informações para instrução processual")
texto_instrucao = montar_texto_instrucao(
    valor_original,
    percentual_garantia,
    garantia_original,
    valor_atualizado,
    garantia_constituida,
    garantia_exigida,
    endosso_necessario,
).replace("5 dias úteis", f"{prazo_dias} dias úteis")

st.text_area(
    "Texto sugerido",
    value=texto_instrucao,
    height=260,
)

st.download_button(
    "Baixar informações da garantia (TXT)",
    data=(
        "GARANTIA CONTRATUAL\n"
        f"Gerado em: {data_hora_brasilia()}\n\n"
        f"Valor original do contrato: {moeda(valor_original)}\n"
        f"Percentual da garantia: {percentual_garantia_pct:.2f}%\n"
        f"Garantia original: {moeda(garantia_original)}\n"
        f"Valor atualizado do contrato: {moeda(valor_atualizado)}\n"
        f"Nova garantia exigida: {moeda(garantia_exigida)}\n"
        f"Garantia atualmente constituída: {moeda(garantia_constituida)}\n"
        f"Endosso necessário: {moeda(endosso_necessario)}\n\n"
        f"{texto_instrucao}"
    ).encode("utf-8"),
    file_name="garantia_contratual.txt",
    mime="text/plain",
    type="primary",
    use_container_width=False,
)
