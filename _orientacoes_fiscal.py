"""
_orientacoes_fiscal.py
----------------------
Bloco visual simples para orientar a solicitação de informações ao fiscal antes
do download da Planilha Master/ColetaMestre.

Não altera cálculo, XLS, leitor, documentos ou session_state.
"""

import streamlit as st


def render_orientacoes_fiscal():
    """Renderiza bloco expansível de orientação ao fiscal.

    Objetivo: mostrar, de forma rápida, a relação entre o que será apurado,
    a informação recomendada do fiscal e o resultado esperado.
    """
    st.markdown(
        """
        <style>
        .orienta-fiscal-wrap {
            background: #F8FAFC;
            border: 1px solid #CBD5E1;
            border-radius: 14px;
            padding: 12px 14px;
            margin: 8px 0 12px 0;
        }
        .orienta-fiscal-title {
            color: #0F172A;
            font-weight: 800;
            font-size: 0.98rem;
            margin-bottom: 4px;
        }
        .orienta-fiscal-subtitle {
            color: #475569;
            font-size: 0.86rem;
            line-height: 1.35;
            margin-bottom: 12px;
        }
        .orienta-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 10px;
        }
        .orienta-card {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            overflow: hidden;
        }
        .orienta-card-head {
            background: #0F766E;
            color: #FFFFFF;
            font-weight: 800;
            font-size: 0.78rem;
            letter-spacing: 0.03em;
            text-transform: uppercase;
            padding: 8px 10px;
        }
        .orienta-card-body {
            padding: 10px;
            color: #334155;
            font-size: 0.82rem;
            line-height: 1.34;
        }
        .orienta-label {
            color: #64748B;
            font-size: 0.70rem;
            font-weight: 800;
            letter-spacing: 0.04em;
            text-transform: uppercase;
            margin-top: 8px;
            margin-bottom: 3px;
        }
        .orienta-need {
            background: #FEF3C7;
            border: 1px solid #FDE68A;
            color: #78350F;
            border-radius: 8px;
            padding: 6px 7px;
        }
        .orienta-result {
            background: #DCFCE7;
            border: 1px solid #BBF7D0;
            color: #14532D;
            border-radius: 8px;
            padding: 6px 7px;
        }
        .orienta-mini-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 7px;
            font-size: 0.72rem;
        }
        .orienta-mini-table th {
            background: #F1F5F9;
            color: #334155;
            border: 1px solid #E2E8F0;
            padding: 4px;
            text-align: center;
        }
        .orienta-mini-table td {
            border: 1px solid #E2E8F0;
            padding: 4px;
            text-align: center;
            color: #475569;
        }
        .orienta-footer {
            margin-top: 10px;
            background: #F1F5F9;
            border: 1px solid #E2E8F0;
            border-radius: 10px;
            color: #334155;
            font-size: 0.80rem;
            line-height: 1.35;
            padding: 9px 10px;
        }
        @media (max-width: 1000px) {
            .orienta-grid { grid-template-columns: 1fr; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    with st.expander("ⓘ Orientações para o Fiscal", expanded=False):
        st.markdown(
            """
            <div class="orienta-fiscal-wrap">
                <div class="orienta-fiscal-title">Informações do fiscal para apuração do reajuste</div>
                <div class="orienta-fiscal-subtitle">
                    A planilha deve reunir a base necessária para apurar retroativo, saldo remanescente e Valor Total Atualizado do Contrato.
                </div>
                <div class="orienta-grid">
                    <div class="orienta-card">
                        <div class="orienta-card-head">Retroativo</div>
                        <div class="orienta-card-body">
                            <div class="orienta-label">Informação recomendada</div>
                            <div class="orienta-need">Valores finais reconhecidos por competência mensal.</div>
                            <table class="orienta-mini-table">
                                <tr><th>Compet.</th><th>Ciclo</th><th>Valor</th></tr>
                                <tr><td>01/2026</td><td>C2</td><td>R$ 100 mil</td></tr>
                                <tr><td>02/2026</td><td>C2</td><td>R$ 120 mil</td></tr>
                                <tr><td>03/2026</td><td>C2</td><td>R$ 115 mil</td></tr>
                            </table>
                            <div class="orienta-label">Resultado</div>
                            <div class="orienta-result">Retroativo por ciclo e total.</div>
                        </div>
                    </div>
                    <div class="orienta-card">
                        <div class="orienta-card-head">Remanescentes</div>
                        <div class="orienta-card-body">
                            <div class="orienta-label">Informação recomendada</div>
                            <div class="orienta-need">Itens remanescentes no início de cada ciclo e saldo atual em data definida.</div>
                            <table class="orienta-mini-table">
                                <tr><th>Item</th><th>Rem. C1</th><th>Rem. atual</th></tr>
                                <tr><td>Item 1</td><td>80</td><td>30</td></tr>
                                <tr><td>Item 2</td><td>160</td><td>90</td></tr>
                                <tr><td>Item 3</td><td>40</td><td>10</td></tr>
                            </table>
                            <div class="orienta-label">Resultado</div>
                            <div class="orienta-result">Saldo remanescente atualizado.</div>
                        </div>
                    </div>
                    <div class="orienta-card">
                        <div class="orienta-card-head">Valor Total Atualizado</div>
                        <div class="orienta-card-body">
                            <div class="orienta-label">Informação recomendada</div>
                            <div class="orienta-need">Histórico executado por ciclo + saldo remanescente após a última competência.</div>
                            <table class="orienta-mini-table">
                                <tr><th>Parcela</th><th>Base</th><th>Valor</th></tr>
                                <tr><td>C0</td><td>Execução</td><td>R$ 500 mil</td></tr>
                                <tr><td>C1</td><td>Execução</td><td>R$ 800 mil</td></tr>
                                <tr><td>Saldo</td><td>Itens</td><td>R$ 300 mil</td></tr>
                            </table>
                            <div class="orienta-label">Resultado</div>
                            <div class="orienta-result">Valor Total Atualizado do Contrato.</div>
                        </div>
                    </div>
                </div>
                <div class="orienta-footer">
                    Regra prática: financeiro mensal favorece a apuração do retroativo; remanescente por itens favorece o saldo atualizado; histórico executado + saldo remanescente favorece o Valor Total Atualizado.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
