"""Solicitacao interna opcional ao fiscal apos a admissibilidade."""

from __future__ import annotations

from textwrap import dedent
from typing import Any, MutableMapping

import streamlit as st


COMANDO_SOLICITACAO_FISCAL = (
    "Solicitar ao fiscal o preenchimento da planilha Coleta"
)

TEXTO_SOLICITACAO_FISCAL = dedent(
    """\
    Assunto: Solicitação de informações para continuidade da análise de reajuste contratual

    Prezados,

    Recebemos o pedido de reajuste apresentado pela contratada e concluímos a etapa inicial de admissibilidade e conferência dos percentuais aplicáveis.

    Para dar continuidade à análise, especialmente ao cálculo de eventual retroativo e à atualização do valor global do contrato, precisamos agora de algumas informações sobre a execução contratual.

    Preferencialmente, solicitamos o preenchimento da planilha de Coleta anexa, observando os seguintes pontos:

    1. Posição do contrato no início do mês do pedido de reajuste — aba itens_Remanesc

    Deverá ser informada a quantidade remanescente de cada item no início do mês em que foi apresentado o pedido de reajuste.

    Essa informação é necessária para registrar a posição do contrato no momento da atualização dos valores e identificar o saldo remanescente a ser considerado nos cálculos.

    Na aba itens_Remanesc, devem ser registradas as quantidades remanescentes dos itens na posição correspondente.

    2. Histórico da execução/pagamentos — aba financeiro ou itens_PC

    Para calcular eventual retroativo, precisamos conhecer o que foi efetivamente executado ou pago ao longo do contrato, desde o início até o período mais recente disponível.

    Essas informações podem ser apresentadas de uma das seguintes formas:

    - na aba financeiro, com os valores pagos/executados por competência; ou
    - na aba itens_PC, com o histórico dos Pedidos de Compra, informando número do PC, data e valor.

    Não é necessário preencher as duas opções caso uma delas já represente de forma completa o histórico da execução.

    3. Posição atual do contrato — aba CICLO_EM_EXECUCAO

    Também precisamos de uma fotografia atual do contrato, para identificar quanto ainda resta de cada item na data mais recente disponível.

    Nessa aba, o preenchimento é basicamente composto por:

    - data da posição atual; e
    - quantidade que ainda resta de cada item nessa data.

    A própria planilha utiliza essas informações para comparar a posição atual com o saldo existente no início do ciclo e demonstrar a evolução da execução.

    Esses dados são necessários para apurarmos o VTA — Valor Total Atualizado do contrato. Em termos práticos, o VTA representa a posição financeira atualizada do contrato, reunindo de forma consistente o que já foi executado, eventual retroativo decorrente do reajuste e o saldo que ainda permanece para execução.

    Por isso, sempre que possível, pedimos que as informações sejam registradas diretamente na planilha de Coleta anexa, pois ela já está estruturada para organizar os dados e realizar as conferências necessárias.

    Caso seja mais conveniente, as informações também podem ser encaminhadas separadamente para que possamos realizar o preenchimento.

    Atenciosamente,
    """
).strip()

_CHAVE_TEXTO = "solicitacao_fiscal_coleta_texto"
_CHAVE_ANALISE = "solicitacao_fiscal_coleta_analise"


def sincronizar_texto_solicitacao_fiscal(
    estado: MutableMapping[str, Any], assinatura_analise: Any
) -> str:
    """Preserva edicao na mesma analise e restaura a base em uma nova."""
    if estado.get(_CHAVE_ANALISE) != assinatura_analise:
        estado[_CHAVE_ANALISE] = assinatura_analise
        estado[_CHAVE_TEXTO] = TEXTO_SOLICITACAO_FISCAL
    elif _CHAVE_TEXTO not in estado:
        estado[_CHAVE_TEXTO] = TEXTO_SOLICITACAO_FISCAL
    return str(estado[_CHAVE_TEXTO])


def render_solicitacao_fiscal_coleta(assinatura_analise: Any) -> None:
    """Renderiza a acao secundaria sem bloquear ou alimentar o fluxo."""
    sincronizar_texto_solicitacao_fiscal(st.session_state, assinatura_analise)
    with st.expander(COMANDO_SOLICITACAO_FISCAL, expanded=False):
        st.caption(
            "Ação opcional. Revise o texto abaixo e copie-o para a comunicação interna."
        )
        st.text_area(
            "Texto da solicitação interna ao fiscal",
            key=_CHAVE_TEXTO,
            height=460,
            label_visibility="collapsed",
        )
