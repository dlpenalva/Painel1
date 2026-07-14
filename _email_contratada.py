"""Rascunho da comunicação posterior à validação fiscal do reajuste."""

from __future__ import annotations

from html import escape
from typing import Any, Iterable, Mapping

import streamlit as st


ASSUNTO_EMAIL_CONTRATADA = "Apuração de atualização contratual – Solicitação de manifestação"


def _percentual(valor: Any) -> str:
    try:
        numero = float(valor or 0)
    except (TypeError, ValueError):
        return "[XX,XX%]"
    if abs(numero) <= 1:
        numero *= 100
    return f"{numero:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def gerar_rascunho_email_contratada(
    ciclos: Iterable[Mapping[str, Any]] | None,
    numero_contrato: str | None = None,
) -> tuple[str, str]:
    """Monta o texto sem estimar valores financeiros ausentes da calculadora."""
    contrato = (numero_contrato or "").strip() or "[NÚMERO DO CONTRATO]"
    linhas_indices: list[str] = []
    linhas_retroativo: list[str] = []

    for ciclo in ciclos or []:
        nome = str(ciclo.get("ciclo") or ciclo.get("Ciclo") or "").strip().upper()
        if not nome or nome == "C0":
            continue
        indice_ciclo = _percentual(ciclo.get("percentual_aplicado", ciclo.get("variacao")))
        fator_acumulado = ciclo.get("fator_acumulado")
        indice_acumulado = (
            _percentual(float(fator_acumulado) - 1)
            if fator_acumulado not in (None, "")
            else "[XX,XX%]"
        )
        linhas_indices.append(f"{nome}\t{indice_ciclo}\t{indice_acumulado}")
        linhas_retroativo.append(f"{nome}\tR$ [VALOR VALIDADO]")

    if not linhas_indices:
        linhas_indices = ["[CICLO]\t[XX,XX%]\t[XX,XX%]"]
        linhas_retroativo = ["[CICLO]\tR$ [VALOR VALIDADO]"]

    corpo = f"""Prezados,

Concluímos a apuração dos valores decorrentes da atualização do Contrato nº {contrato}.

Os ciclos contemplados nesta apuração, bem como os respectivos índices de reajuste e índices acumulados aplicados, estão apresentados no quadro abaixo:

Ciclo\tÍndice do ciclo\tÍndice acumulado
{chr(10).join(linhas_indices)}

Em decorrência dessa atualização, foi apurado o seguinte valor de retroativo:

Ciclo\tRetroativo reconhecido a pagar
{chr(10).join(linhas_retroativo)}
Total\tR$ [VALOR TOTAL VALIDADO]

A atualização contratual será formalizada por meio de Termo de Apostila, instrumento utilizado para registrar a atualização dos valores do contrato.

A planilha anexa apresenta o detalhamento da apuração, contendo, entre outras informações:

- valores unitários atualizados;
- valores retroativos por ciclo;
- quantitativos e valores remanescentes do contrato;
- memória de cálculo da atualização.

Solicitamos, por gentileza, a manifestação dessa empresa quanto ao de acordo com os valores apresentados, para que possamos dar prosseguimento ao apostilamento.

Permanecemos à disposição para quaisquer esclarecimentos."""
    return ASSUNTO_EMAIL_CONTRATADA, corpo


def render_email_contratada(
    ciclos: Iterable[Mapping[str, Any]] | None,
    *,
    numero_contrato: str | None = None,
    key: str,
) -> None:
    """Exibe a opção no trecho posterior ao processamento da calculadora."""
    assunto, corpo = gerar_rascunho_email_contratada(ciclos, numero_contrato)
    st.markdown(
        '<div style="background:#FFF9E8;border:1.5px solid #E8B923;border-radius:16px;'
        'padding:20px 24px;margin:18px 0 8px 0;">'
        '<div style="font-size:.82rem;font-weight:800;color:#8A3D18;letter-spacing:.07em;'
        'margin-bottom:10px;">✉ &nbsp;COMUNICAÇÃO À CONTRATADA</div>'
        '<div style="font-size:1rem;color:#8A3D18;line-height:1.55;">'
        'Rascunho pré-redigido com os percentuais desta análise. '
        '<span style="color:#A76948;">Complete os valores entre colchetes e envie somente após a validação do fiscal.</span>'
        '</div>'
        f'<div style="font-size:.92rem;color:#9A461C;margin-top:14px;">'
        f'<strong>Assunto:</strong> {escape(assunto)}</div></div>',
        unsafe_allow_html=True,
    )
    st.download_button(
        "Baixar rascunho (.txt)",
        data=f"ASSUNTO: {assunto}\n\n{corpo}".encode("utf-8-sig"),
        file_name="Comunicacao_Contratada_Reajuste.txt",
        mime="text/plain; charset=utf-8",
        type="primary",
        key=key,
        help="Revise contrato e valores validados antes do envio à contratada.",
    )
