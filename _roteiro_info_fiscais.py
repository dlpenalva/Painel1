# -*- coding: utf-8 -*-
"""
_roteiro_info_fiscais.py
------------------------
Bloco de UX do Roteiro das Informações dos Fiscais.

Versão estável v7:
- não seta diretamente chaves de widgets no session_state;
- evita aviso do Streamlit sobre widget com default + session_state;
- mantém perguntas excludentes com resposta lógica N/A;
- mantém resumo automático e grava st.session_state["roteiro_info_fiscais"].
"""
from __future__ import annotations

from typing import Dict, Tuple
import html
import streamlit as st

OPCOES_BASE = ["Sim", "Não", "Não sei / confirmar"]


def _safe(texto: object) -> str:
    return html.escape(str(texto or ""))


def _pergunta(numero: int, texto: str, ajuda: str | None = None) -> None:
    ajuda_html = ""
    if ajuda:
        ajuda_html = (
            '<div style="font-size:0.78rem;color:#64748B;margin-top:2px;line-height:1.35;">'
            f'{_safe(ajuda)}</div>'
        )
    st.markdown(
        f"""
        <div style="margin:0.55rem 0 0.18rem 0;">
            <div style="font-size:0.90rem;font-weight:650;color:#0F172A;line-height:1.35;">
                <span style="display:inline-flex;align-items:center;justify-content:center;width:1.35rem;height:1.35rem;border-radius:999px;background:#CCFBF1;color:#0F766E;font-weight:800;font-size:0.78rem;margin-right:0.35rem;">{numero}</span>
                {_safe(texto)}
            </div>
            {ajuda_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def _radio(prefix: str, nome: str, index: int = 2, disabled: bool = False) -> str:
    # Sufixo v7 evita reaproveitar chaves antigas contaminadas por set manual no session_state.
    return st.radio(
        "Resposta",
        options=OPCOES_BASE,
        index=index,
        key=f"{prefix}_{nome}_v7",
        horizontal=True,
        disabled=disabled,
        label_visibility="collapsed",
    )


def _campo_texto(prefix: str, nome: str, placeholder: str = "") -> str:
    return st.text_input(
        "Resposta textual",
        placeholder=placeholder,
        key=f"{prefix}_{nome}_v7",
        label_visibility="collapsed",
    )


def _modelo_sugerido(resp: Dict[str, str]) -> Tuple[str, str, str]:
    tem_fin = resp.get("tem_financeiro_mensal") == "Sim"
    comp_ok = resp.get("financeiro_por_competencia") == "Sim"
    tem_saldo = resp.get("tem_saldo_remanescente") == "Sim"
    tem_consumo = resp.get("tem_itens_consumidos") == "Sim"
    defasagem = resp.get("ha_preco_origem_diferente") == "Sim"
    valores_ant = resp.get("tem_valores_ciclos_anteriores") == "Sim"
    valor_consolidado = resp.get("tem_valor_formalizado_anterior") == "Sim"

    ressalvas = []
    if tem_fin and not comp_ok:
        ressalvas.append("confirmar se os valores financeiros foram organizados pela competência de referência")
    if defasagem:
        ressalvas.append("há possível defasagem entre competência/preço de origem e competência de execução/reconhecimento")
    if not tem_fin and (tem_saldo or tem_consumo):
        ressalvas.append("retroativo financeiro tende a ser estimativo, salvo validação adicional")
    if not valores_ant and valor_consolidado:
        ressalvas.append("histórico anterior poderá ser usado de forma consolidada")
    if tem_saldo:
        ressalvas.append("itens consumidos ficam como N/A quando há saldo remanescente por itens")

    if tem_fin and tem_saldo and defasagem:
        modelo = "ColetaMestre Completa + Preço de Origem"
        base = "financeiro por competência, saldo remanescente datado e tratamento de preço/valor de origem"
    elif tem_fin and tem_saldo:
        modelo = "ColetaMestre Completa"
        base = "financeiro por competência + saldo remanescente por itens"
    elif defasagem and tem_fin:
        modelo = "Coleta com Preço de Origem"
        base = "valor reconhecido em uma competência com preço/valor formado em competência anterior"
    elif tem_fin:
        modelo = "Coleta Financeira para Retroativo"
        base = "valores finais reconhecidos por competência; VTA depende de saldo remanescente"
    elif tem_saldo:
        modelo = "Coleta Estoque/Remanescente"
        base = "saldo remanescente por itens em data de corte; retroativo não é financeiro definitivo"
    elif tem_consumo:
        modelo = "Coleta Consumo por Itens/Ciclo"
        base = "itens/quantidades consumidos por ciclo; resultado sujeito à premissa fiscal"
    elif valor_consolidado:
        modelo = "Coleta com Histórico Consolidado"
        base = "valor formalizado anterior como âncora + informações atuais"
    else:
        modelo = "Roteiro de Saneamento"
        base = "base ainda insuficiente; orientar fiscal antes de gerar coleta definitiva"

    natureza = "; ".join(ressalvas) if ressalvas else "base potencialmente adequada para a planilha sugerida"
    return modelo, base, natureza


def _na_box(texto: str) -> None:
    st.markdown(
        f"""
        <div style="background:#F8FAFC;border:1px dashed #CBD5E1;border-radius:10px;padding:9px 11px;margin:4px 0 10px 0;color:#64748B;font-size:0.88rem;">
            <b>N/A.</b> {_safe(texto)}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_roteiro_info_fiscais(contexto: str = "") -> Dict[str, str]:
    """Renderiza bloco de qualificação e grava respostas no session_state."""
    prefix = f"roteiro_info_fiscais_{contexto}" if contexto else "roteiro_info_fiscais"

    st.markdown(
        """
        <div style="background:#F0FDFA;border:1px solid #99F6E4;border-radius:16px;padding:14px 16px;margin:14px 0 10px 0;">
          <div style="font-size:0.78rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#0F766E;margin-bottom:4px;">
            Preparação da ColetaMestre
          </div>
          <div style="font-size:1.02rem;font-weight:650;color:#0F172A;margin-bottom:4px;">
            Roteiro das Informações dos Fiscais
          </div>
          <div style="font-size:0.90rem;color:#334155;line-height:1.45;">
            Antes de baixar a planilha, informe que tipo de dado será solicitado ao fiscal.
            O objetivo é gerar e interpretar a ColetaMestre com menor risco de erro no retroativo e no Valor Total Atualizado.
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.expander("Preencher roteiro rápido", expanded=True):
        st.caption(
            "Regra operacional: quando o roteiro mencionar valor final reconhecido/pago, considere o valor final reconhecido para pagamento, já com glosas/descontos contratuais e em base bruta econômica."
        )

        c1, c2 = st.columns(2)
        with c1:
            _pergunta(1, "O fiscal conseguirá informar valores finais reconhecidos para pagamento por competência mensal?", "Base preferencial para cálculo do retroativo.")
            tem_fin = _radio(prefix, "tem_fin", index=2)

            _pergunta(2, "Esses valores estarão organizados pela competência de referência, e não apenas pela data do pagamento?")
            comp_ok = _radio(prefix, "comp_ok", index=0)

            _pergunta(3, "Até qual competência/mês haverá informação financeira?")
            ult_fin = _campo_texto(prefix, "ult_fin", placeholder="Ex.: 04/2026")

            _pergunta(4, "O fiscal conseguirá informar saldo remanescente por itens/quantidades?")
            tem_saldo = _radio(prefix, "tem_saldo", index=2)

            _pergunta(
                5,
                "Esse saldo remanescente representa a posição de qual competência/data de corte?",
                "Saldo remanescente é aquilo que ainda falta executar, entregar, consumir ou faturar no contrato a partir de determinada data de corte.",
            )
            data_saldo = _campo_texto(prefix, "data_saldo", placeholder="Ex.: 05/2026 ou 30/05/2026")

        with c2:
            _pergunta(6, "Caso não haja saldo remanescente por itens, o fiscal conseguirá informar itens/quantidades consumidos ou executados por ciclo?")
            if tem_saldo == "Sim":
                tem_consumo = "N/A"
                _na_box("Como há saldo remanescente por itens, esta alternativa não precisa ser preenchida.")
            else:
                tem_consumo = _radio(prefix, "tem_consumo", index=2)

            _pergunta(
                7,
                "Há casos em que o valor informado, pedido de compra ou preço utilizado foi formado em uma competência, mas a entrega, execução, medição ou atesto ocorreu em competência posterior ou em outro ciclo?",
                "Ex.: pedido em fevereiro com preço de fevereiro, mas entrega/atesto em novembro.",
            )
            defasagem = _radio(prefix, "defasagem", index=1)

            _pergunta(8, "O fiscal conseguirá informar os valores finais reconhecidos/pagos dos ciclos anteriores, como C0, C1 ou C2?")
            valores_ant = _radio(prefix, "valores_ant", index=2)

            _pergunta(9, "Caso não consiga informar os valores por ciclo anterior, existe valor formalizado/consolidado anterior que possa servir como ponto de partida da análise?")
            if valores_ant == "Sim":
                valor_consolidado = "N/A"
                _na_box("Como o fiscal informará os valores por ciclo anterior, esta alternativa fica dispensada.")
            else:
                valor_consolidado = _radio(prefix, "valor_consolidado", index=2)

        resp = {
            "tem_financeiro_mensal": tem_fin,
            "financeiro_por_competencia": comp_ok,
            "ultima_competencia_financeira": ult_fin.strip(),
            "tem_saldo_remanescente": tem_saldo,
            "data_corte_saldo_remanescente": data_saldo.strip(),
            "tem_itens_consumidos": tem_consumo,
            "ha_preco_origem_diferente": defasagem,
            "tem_valores_ciclos_anteriores": valores_ant,
            "tem_valor_formalizado_anterior": valor_consolidado,
        }

        modelo, base, natureza = _modelo_sugerido(resp)
        resp["modelo_sugerido"] = modelo
        resp["base_esperada"] = base
        resp["observacao_qualidade"] = natureza

        st.session_state["roteiro_info_fiscais"] = resp

        st.markdown(
            f"""
            <div style="background:#F0FDF4;border:1px solid #BBF7D0;border-radius:14px;padding:12px 14px;margin-top:10px;">
              <div style="font-size:0.78rem;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:.06em;">Resumo automático</div>
              <div style="font-size:0.98rem;color:#0F172A;margin-top:4px;"><b>Modelo sugerido:</b> {_safe(modelo)}</div>
              <div style="font-size:0.90rem;color:#334155;margin-top:3px;"><b>Base esperada:</b> {_safe(base)}</div>
              <div style="font-size:0.86rem;color:#64748B;margin-top:3px;"><b>Observação:</b> {_safe(natureza)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    return st.session_state.get("roteiro_info_fiscais", {})
