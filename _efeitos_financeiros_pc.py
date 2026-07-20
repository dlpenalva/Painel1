"""Regra canonica de efeito financeiro para Pedidos de Compra.

Esta camada nao enquadra o PC em ciclo. Ela apenas decide, depois de o ciclo
temporal ter sido apurado, se a DATA_PC ja alcancou o inicio do efeito
financeiro daquele ciclo. A fonte visivel e ``parametros``; o metadado
``CL8US_INICIO_EFEITO`` funciona como copia de integridade/compatibilidade.
"""
from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime
from typing import Any


CICLOS_PC = ("C0", "C1", "C2", "C3", "C4")
CABECALHO_INICIO_EFEITO = "INICIO_EFEITO_FINANCEIRO"
PREFIXO_METADADO = "CL8US_INICIO_EFEITO:"


def como_data(valor: Any) -> date | None:
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    texto = str(valor or "").strip()
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%Y"):
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            continue
    return None


def _cabecalho(valor: Any) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return re.sub(r"[^A-Z0-9]+", "_", texto.upper()).strip("_")


def inicios_em_parametros(wb) -> tuple[dict[str, date], bool]:
    if "parametros" not in wb.sheetnames:
        return {}, False
    ws = wb["parametros"]
    # Primeira ocorrencia vence: o bloco MEMORIA DE CALCULO (parametros!J:R)
    # repete nomes como CICLO sem sombrear o layout oficial A:H.
    colunas: dict[str, int] = {}
    for c in ws[1]:
        if c.value not in (None, ""):
            colunas.setdefault(_cabecalho(c.value), c.column)
    col_ciclo = colunas.get("CICLO")
    col_inicio = colunas.get(CABECALHO_INICIO_EFEITO)
    if not col_ciclo or not col_inicio:
        return {}, False
    resultado: dict[str, date] = {}
    for linha in range(2, min(ws.max_row, 100) + 1):
        ciclo = str(ws.cell(linha, col_ciclo).value or "").strip().upper()
        inicio = como_data(ws.cell(linha, col_inicio).value)
        if ciclo in CICLOS_PC and inicio is not None:
            resultado[ciclo] = inicio
    return resultado, True


def inicios_em_metadado(wb) -> tuple[dict[str, date], bool]:
    texto = str(wb.properties.keywords or "")
    if PREFIXO_METADADO not in texto:
        return {}, False
    trecho = texto.split(PREFIXO_METADADO, 1)[1].split(";", 1)[0]
    resultado: dict[str, date] = {}
    for ciclo, data_iso in re.findall(r"\b(C[0-4])=(\d{4}-\d{2}-\d{2})\b", trecho):
        try:
            resultado[ciclo] = date.fromisoformat(data_iso)
        except ValueError:
            continue
    return resultado, True


def reconciliar_inicios_efeito(
    wb,
) -> tuple[dict[str, date], list[str], bool, bool]:
    """Retorna fonte reconciliada, erros e presenca das duas representacoes."""
    parametros, tem_parametros = inicios_em_parametros(wb)
    metadado, tem_metadado = inicios_em_metadado(wb)
    erros: list[str] = []
    resultado: dict[str, date] = {}
    for ciclo in CICLOS_PC:
        visivel = parametros.get(ciclo)
        tecnico = metadado.get(ciclo)
        if visivel is not None and tecnico is not None and visivel != tecnico:
            erros.append(
                f"{ciclo}: INICIO_EFEITO_FINANCEIRO em parametros "
                f"({visivel.isoformat()}) diverge de CL8US_INICIO_EFEITO "
                f"({tecnico.isoformat()})."
            )
            continue
        if visivel is not None:
            resultado[ciclo] = visivel
        elif tecnico is not None:
            resultado[ciclo] = tecnico
    return resultado, erros, tem_parametros, tem_metadado


def efeito_financeiro_pc(
    data_pc: Any,
    ciclo: Any,
    parametros_ciclo: dict[str, Any] | None,
) -> str | None:
    """Classifica por data exata; ``None`` significa resultado indeterminado."""
    data_norm = como_data(data_pc)
    ciclo_norm = str(ciclo or "").strip().upper()
    if data_norm is None or ciclo_norm not in CICLOS_PC:
        return None
    if ciclo_norm == "C0":
        return "Nao"
    registro = parametros_ciclo or {}
    computar = str(registro.get("computar_nesta_apuracao") or "").strip().lower()
    if computar not in {"sim", "s", "true", "1", "yes"}:
        return "Nao"
    inicio = como_data(registro.get("inicio_efeito_financeiro"))
    if inicio is None:
        return None
    return "Sim" if data_norm >= inicio else "Nao"
