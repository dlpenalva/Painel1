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

    §3.3: IST (Anatel), ICTI (Ipeadata), IPCA, IGP-M, INPC.
    Legado "IST (Série Local)" continua reconhecido como IST.
    """
    texto = remover_emojis(indice).strip()
    if not texto:
        return "[ÍNDICE]"
    norm = texto.upper()
    if norm.startswith("IST"):
        return "IST (Anatel)"
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


def _competencia_mm_aaaa(iso: Any) -> str | None:
    """'2025-02-01' -> '02/2025'. Sem data valida, sem linha (nao inventa)."""
    texto = str(iso or "").strip()
    partes = texto.split("-")
    if len(partes) < 2 or not (partes[0].isdigit() and partes[1].isdigit()):
        return None
    return f"{int(partes[1]):02d}/{partes[0]}"


def _fator_6_casas(valor: Any) -> str | None:
    try:
        return f"{float(valor):.6f}".replace(".", ",")
    except (TypeError, ValueError):
        return None


def _numero_indice(valor: Any) -> str | None:
    """Numero-indice em 4 casas, como no bloco parametros!N (formato 0.0000)."""
    try:
        return f"{float(valor):.4f}".replace(".", ",")
    except (TypeError, ValueError):
        return None


def _bloco_memoria_ciclo(numero: str, memoria: list[Mapping[str, Any]]) -> str | None:
    """Formata a memoria de UM ciclo. Reapresenta o payload; nao recalcula.

    MES    -> uma linha 'mm/aaaa: x,xx%' por competencia presente na memoria.
    INDICE -> competencia e numero-indice inicial e final (IST nunca vira
              memoria mensal ficticia).
    RESULTADO -> Fator apurado / Variacao apurada / Metodo/Fonte canonicos.
    """
    mensais: list[str] = []
    indices: list[str] = []
    finais: list[str] = []
    registros = [r for r in memoria if isinstance(r, Mapping)]
    tipo_indice = [
        r for r in registros
        if str(r.get("tipo") or "").strip().upper() == "INDICE"
    ]
    for registro in registros:
        tipo = str(registro.get("tipo") or "").strip().upper()
        if tipo == "MES":
            competencia = _competencia_mm_aaaa(registro.get("competencia"))
            taxa = registro.get("valor_indice")
            if competencia is None or taxa in (None, ""):
                continue
            mensais.append(f"{competencia}: {_percentual(taxa)}")
        elif tipo == "RESULTADO":
            fator = _fator_6_casas(registro.get("fator_acumulado"))
            if fator is not None:
                finais.append(f"Fator apurado: {fator}")
            variacao = registro.get("variacao_final")
            if variacao not in (None, ""):
                finais.append(f"Variação apurada: {_percentual(variacao)}")
            metodo_fonte = str(registro.get("metodo_fonte") or "").strip()
            if metodo_fonte:
                finais.append(f"Método/Fonte: {metodo_fonte}")
    # IST: a memoria canonica pode trazer a serie mensal completa de
    # numeros-indice; o TXT apresenta APENAS o inicial e o final (o metodo e
    # a divisao entre eles) — a serie integral permanece na memoria do XLS.
    indices_validos = []
    for registro in tipo_indice:
        competencia = _competencia_mm_aaaa(registro.get("competencia"))
        valor = _numero_indice(registro.get("valor_indice"))
        if competencia is None or valor is None:
            continue
        indices_validos.append((competencia, valor))
    if indices_validos:
        competencia, valor = indices_validos[0]
        indices.append(f"Número-índice inicial ({competencia}): {valor}")
        if len(indices_validos) > 1:
            competencia, valor = indices_validos[-1]
            indices.append(f"Número-índice final ({competencia}): {valor}")

    corpo = mensais or indices
    if not corpo and not finais:
        return None
    partes = [f"Ciclo {numero}", ""]
    if corpo:
        partes.extend(corpo)
        partes.append("")
    partes.extend(finais)
    return "\n".join(partes).rstrip()


def _secao_memoria_calculo(ciclos: Iterable[Mapping[str, Any]] | None) -> str:
    """Secao MEMORIA DE CALCULO: um bloco por ciclo que POSSUI memoria.

    Reutiliza exclusivamente `memoria_calculo` ja produzida/normalizada pelo
    sistema (registros MES/INDICE/RESULTADO). Ciclo sem memoria nao gera
    bloco; sem nenhum bloco, a secao inteira e omitida.
    """
    blocos: list[str] = []
    for ciclo in ciclos or []:
        if not isinstance(ciclo, Mapping):
            continue
        nome = str(ciclo.get("ciclo") or ciclo.get("Ciclo") or "").strip()
        numero = _numero_ciclo(nome)
        if numero is None:
            continue
        memoria = ciclo.get("memoria_calculo")
        if not isinstance(memoria, (list, tuple)) or not memoria:
            continue
        bloco = _bloco_memoria_ciclo(numero, list(memoria))
        if bloco:
            blocos.append(bloco)
    if not blocos:
        return ""
    return "\n\nMEMÓRIA DE CÁLCULO\n\n" + "\n\n".join(blocos)


def _ciclos_com_memoria(ciclos: Iterable[Mapping[str, Any]] | None) -> list[str]:
    """Numeros dos ciclos que POSSUEM memoria de calculo no payload."""
    numeros: list[str] = []
    for ciclo in ciclos or []:
        if not isinstance(ciclo, Mapping):
            continue
        numero = _numero_ciclo(str(ciclo.get("ciclo") or ciclo.get("Ciclo") or ""))
        if numero is None:
            continue
        memoria = ciclo.get("memoria_calculo")
        if isinstance(memoria, (list, tuple)) and memoria:
            numeros.append(numero)
    return numeros


def _conferir_entrega_da_memoria(
    ciclos: Iterable[Mapping[str, Any]] | None, corpo: str
) -> None:
    """FAIL-CLOSED da comunicacao: memoria disponivel TEM de chegar ao TXT.

    Regra funcional: se um ciclo foi processado e a memoria de calculo
    detalhada existe no sistema, o TXT correspondente DEVE conter a memoria.
    Nunca mais "memoria sumiu -> gera TXT silenciosamente sem memoria": se a
    secao (ou o bloco de um ciclo com memoria) nao puder ser gerada, o erro e
    controlado e visivel — nada de TXT incompleto entregue em silencio.
    """
    esperados = _ciclos_com_memoria(ciclos)
    if not esperados:
        return
    if "MEMÓRIA DE CÁLCULO" not in corpo:
        raise ValueError(
            "A memória de cálculo dos ciclos "
            + ", ".join(esperados)
            + " existe no sistema, mas a seção MEMÓRIA DE CÁLCULO não pôde ser "
            "gerada no TXT. Comunicação bloqueada para não entregar documento "
            "incompleto."
        )
    secao = corpo.split("MEMÓRIA DE CÁLCULO", 1)[1]
    faltantes = [n for n in esperados if f"Ciclo {n}\n" not in secao]
    if faltantes:
        raise ValueError(
            "A memória de cálculo do(s) ciclo(s) "
            + ", ".join(faltantes)
            + " existe no sistema, mas o bloco correspondente não pôde ser "
            "gerado no TXT. Comunicação bloqueada para não entregar documento "
            "incompleto."
        )


def _fator_acumulado_final(ciclos: Iterable[Mapping[str, Any]] | None) -> float | None:
    """Ultimo fator_acumulado NUMERICO do payload dos ciclos.

    E a cadeia CANONICA ja calculada pelo sistema (produto dos fatores, nunca
    soma de percentuais). Sem fator numerico no payload, devolve None e a
    linha do acumulado nao e emitida (nada e inventado)."""
    final = None
    for ciclo in ciclos or []:
        if not isinstance(ciclo, Mapping):
            continue
        bruto = ciclo.get("fator_acumulado")
        try:
            final = float(bruto)
        except (TypeError, ValueError):
            continue
    return final


def montar_txt_download(assunto: str, corpo: str) -> bytes:
    """Bytes EXATOS entregues ao botao de download do rascunho (.txt)."""
    return f"ASSUNTO: {assunto}\n\n{corpo}".encode("utf-8-sig")


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
    # Variacao acumulada FINAL: reapresenta a cadeia canonica do sistema
    # (ultimo fator_acumulado do payload). Nunca soma percentuais de ciclos.
    fator_final = _fator_acumulado_final(ciclos)
    if fator_final is None:
        acumulado_txt = ""
    else:
        pct_txt = (f"{(fator_final - 1) * 100:,.2f}%"
                   .replace(",", "X").replace(".", ",").replace("X", "."))
        acumulado_txt = f"\n\nVariação Acumulada Final: {pct_txt}."

    corpo = (
        "Prezados,\n\n"
        f"Em atenção ao pedido de reajuste do Contrato {contrato}, apresentamos, "
        "para ciência e manifestação da Contratada, a apuração preliminar das "
        "variações aplicáveis a cada ciclo, conforme a metodologia contratual e o "
        f"índice {indice_txt}:\n\n"
        + "\n".join(linhas)
        + acumulado_txt
        + "\n\n"
        "Eventuais ciclos preclusos permanecem registrados para fins de "
        "histórico e memória contratual, sem geração de efeitos financeiros "
        "retroativos.\n\n"
        "Esta comunicação refere-se exclusivamente à apuração dos ciclos e "
        "percentuais aplicáveis. Após a concordância da Contratada, será dado "
        "prosseguimento à apuração dos valores financeiros correspondentes e às "
        "demais providências necessárias à formalização.\n\n"
        "Solicitamos, assim, manifestação de concordância quanto às informações "
        "acima."
    )
    # Memoria de calculo DO INDICE (pre-apuracao: percentuais e fatores, nunca
    # valores financeiros do contrato), reapresentada do payload dos ciclos.
    corpo += _secao_memoria_calculo(ciclos)
    # FAIL-CLOSED: memoria disponivel no payload TEM de estar no corpo.
    _conferir_entrega_da_memoria(ciclos, corpo)
    # Blindagem final: nenhum arquivo entregue pode conter emoji.
    return ASSUNTO_EMAIL_CONTRATADA, remover_emojis(corpo)


def render_email_contratada(
    ciclos: Iterable[Mapping[str, Any]] | None,
    *,
    numero_contrato: str | None = None,
    indice: str | None = None,
    key: str,
) -> None:
    """Exibe a comunicacao pre-apuracao e o botao de download do rascunho.

    FAIL-CLOSED: se a memoria de calculo disponivel nao puder ser entregue no
    TXT, o erro e mostrado ao usuario e o download NAO e oferecido — nunca um
    documento silenciosamente incompleto.
    """
    try:
        assunto, corpo = gerar_rascunho_email_contratada(
            ciclos, numero_contrato, indice
        )
    except ValueError as exc:
        st.error(f"Comunicação à Contratada indisponível: {exc}")
        return
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
        data=montar_txt_download(assunto, corpo),
        file_name="Comunicacao_Contratada_Reajuste.txt",
        mime="text/plain; charset=utf-8",
        type="primary",
        key=key,
        help="Revise o número do contrato antes do envio à contratada.",
    )
