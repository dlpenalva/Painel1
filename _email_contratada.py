"""Comunicacao a Contratada — etapa PRE-apuracao financeira.

Esta comunicacao ocorre ANTES da apuracao financeira definitiva. Comunica a
Contratada os ciclos, o indice, o percentual e a situacao/efeito de cada ciclo,
e solicita concordancia para prosseguimento. NAO menciona valor retroativo, VTA,
remanescente, valores unitarios, planilha anexa, "valores validados", conclusao
da apuracao financeira nem apostilamento pronto (regra documental do pacote).
"""

from __future__ import annotations

import re
from html import escape
from typing import Any, Iterable, Mapping

import streamlit as st

from _sanitizacao_documental import remover_emojis


ASSUNTO_EMAIL_CONTRATADA = (
    "Apuração preliminar dos ciclos de reajuste – Solicitação de manifestação"
)


def _percentual(valor: Any) -> str:
    """Formata percentual em dd,dd%. Aceita fracao (<=1) ou percentual pronto."""
    try:
        numero = float(valor if valor not in (None, "") else 0)
    except (TypeError, ValueError):
        return "[XX,XX%]"
    if abs(numero) <= 1:
        numero *= 100
    return f"{numero:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def _indice_amigavel(indice: Any) -> str:
    """Nome amigavel do indice, sem expor codigos tecnicos (433/189/DIMAC).

    §3.3: IST (Série Local), ICTI (Ipeadata), IPCA, IGP-M.
    """
    texto = remover_emojis(indice).strip()
    if not texto:
        return "[ÍNDICE]"
    norm = texto.upper()
    if norm.startswith("IST"):
        return "IST (Série Local)"
    if norm.startswith("ICTI"):
        return "ICTI (Ipeadata)"
    if norm.startswith("IPCA"):
        return "IPCA"
    if norm.startswith("IGP"):
        return "IGP-M"
    if norm.startswith("INPC"):
        return "INPC"
    # Remove sufixo de codigo tecnico entre parenteses so quando for numerico.
    return re.sub(r"\s*\(\s*\d+\s*\)\s*$", "", texto).strip()


def _numero_ciclo(nome: str) -> str | None:
    """'C1' -> '1'. Ignora C0/base e rotulos nao reconhecidos."""
    limpo = remover_emojis(nome).strip().upper()
    if not limpo.startswith("C"):
        return None
    digito = limpo[1:].strip()
    if digito.isdigit() and digito != "0":
        return digito
    return None


def _situacao_efeito(ciclo: Mapping[str, Any]) -> str:
    """Monta o trecho '[SITUAÇÃO/EFEITO]' de um ciclo, sem inventar dados.

    - precluso -> 'precluso, sem efeitos financeiros'
    - com data segura -> '<situacao>, com efeitos financeiros a partir de dd/mm/aaaa'
    - sem data segura -> '<situacao>, efeitos financeiros pendentes de definição'
    """
    situacao_bruta = str(
        ciclo.get("situacao_aplicada", ciclo.get("situacao", "")) or ""
    )
    situacao = remover_emojis(situacao_bruta).strip(" -–—|").strip().lower()
    situacao_low = situacao

    inicio = str(
        ciclo.get("financeiro_inicio", ciclo.get("inicio_financeiro", "")) or ""
    ).strip()
    data_segura = "/" in inicio and any(ch.isdigit() for ch in inicio) and (
        "sem efeito" not in inicio.lower()
    )

    if "preclu" in situacao_low:
        rotulo = situacao if situacao else "precluso"
        return f"{rotulo}, sem efeitos financeiros".strip(", ")
    if data_segura:
        rotulo = situacao if situacao else "tempestivo"
        return f"{rotulo}, com efeitos financeiros a partir de {inicio}".strip(", ")
    rotulo = situacao if situacao else "situação em análise"
    return f"{rotulo}, efeitos financeiros pendentes de definição".strip(", ")


def _linhas_ciclos(ciclos: Iterable[Mapping[str, Any]] | None) -> list[str]:
    linhas: list[str] = []
    for ciclo in ciclos or []:
        nome = str(ciclo.get("ciclo") or ciclo.get("Ciclo") or "").strip()
        numero = _numero_ciclo(nome)
        if numero is None:
            continue
        pct = ciclo.get("variacao_formatada")
        if pct in (None, ""):
            pct = _percentual(
                ciclo.get("percentual_aplicado", ciclo.get("variacao"))
            )
        else:
            pct = str(pct).strip()
        linhas.append(f"• Ciclo {numero}: {pct} – {_situacao_efeito(ciclo)};")
    if linhas:
        # ultimo item termina em ponto final
        linhas[-1] = linhas[-1].rstrip(";") + "."
    return linhas


def gerar_rascunho_email_contratada(
    ciclos: Iterable[Mapping[str, Any]] | None,
    numero_contrato: str | None = None,
    indice: str | None = None,
) -> tuple[str, str]:
    """Monta a comunicacao pre-apuracao, sem qualquer valor financeiro."""
    contrato = (numero_contrato or "").strip() or "[CONTRATO]"
    indice_txt = _indice_amigavel(indice)
    linhas = _linhas_ciclos(ciclos)
    if not linhas:
        linhas = ["• Ciclo [N]: [XX,XX%] – [situação/efeito]."]

    corpo = (
        "Prezados,\n\n"
        f"Em atenção ao pedido de reajuste do Contrato {contrato}, apresentamos, "
        "para ciência e manifestação da Contratada, a apuração preliminar das "
        "variações aplicáveis a cada ciclo, conforme a metodologia contratual e o "
        f"índice {indice_txt}:\n\n"
        + "\n".join(linhas)
        + "\n\n"
        "Os ciclos preclusos permanecem registrados para fins de histórico e "
        "memória contratual, sem geração de efeitos financeiros retroativos.\n\n"
        "Esta comunicação refere-se exclusivamente à apuração dos ciclos e "
        "percentuais aplicáveis. Após a concordância da Contratada, será dado "
        "prosseguimento à apuração dos valores financeiros correspondentes e às "
        "demais providências necessárias à formalização.\n\n"
        "Solicitamos, assim, manifestação de concordância quanto às informações "
        "acima."
    )
    # Blindagem final: nenhum arquivo entregue pode conter emoji.
    return ASSUNTO_EMAIL_CONTRATADA, remover_emojis(corpo)


def render_email_contratada(
    ciclos: Iterable[Mapping[str, Any]] | None,
    *,
    numero_contrato: str | None = None,
    indice: str | None = None,
    key: str,
) -> None:
    """Exibe a comunicacao pre-apuracao e o botao de download do rascunho."""
    assunto, corpo = gerar_rascunho_email_contratada(ciclos, numero_contrato, indice)
    st.markdown(
        '<div style="background:#FFF9E8;border:1.5px solid #E8B923;border-radius:16px;'
        'padding:20px 24px;margin:18px 0 8px 0;">'
        '<div style="font-size:.82rem;font-weight:800;color:#8A3D18;letter-spacing:.07em;'
        'margin-bottom:10px;">✉ &nbsp;COMUNICAÇÃO À CONTRATADA</div>'
        '<div style="font-size:1rem;color:#8A3D18;line-height:1.55;">'
        'Rascunho pré-redigido com os ciclos e percentuais desta análise, para '
        'ciência e concordância da Contratada '
        '<span style="color:#A76948;">antes da apuração financeira definitiva.</span>'
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
        help="Revise o número do contrato antes do envio à contratada.",
    )
