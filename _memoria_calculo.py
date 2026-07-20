"""Memoria de calculo mensal persistida no XLS oficial (Etapa 4).

Normaliza a memoria JA CALCULADA e exibida pelas Calculadoras (res['dados'])
em registros JSON-safe, grava o bloco parametros!J2:R80 no gerador e le o
bloco de forma opcional no leitor v10. NAO recalcula indice, variacao, VTA,
retroativo ou efeitos financeiros: apenas persiste o que a aplicacao exibiu.

Layout (parametros!J1:R1, cabecalhos gravados no template via Excel COM):
    J=CICLO  K=TIPO_REGISTRO  L=ORDEM  M=COMPETENCIA  N=VALOR_INDICE
    O=FATOR_MENSAL  P=FATOR_ACUMULADO  Q=VARIACAO_FINAL  R=METODO_FONTE

Tipos de registro:
    MES       - competencia mensal; N = taxa mensal como decimal (0,45% -> 0.0045);
                O/P somente quando ja existem na memoria (ICTI); Q vazio.
    INDICE    - IST: numero-indice inicial (ordem 1) e final (ordem 2) em N.
    RESULTADO - ultima linha do ciclo; P = fator final; Q = variacao final
                canonica do payload; R = metodologia/fonte canonica.
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

CABECALHOS_MEMORIA_CALCULO = (
    "CICLO", "TIPO_REGISTRO", "ORDEM", "COMPETENCIA", "VALOR_INDICE",
    "FATOR_MENSAL", "FATOR_ACUMULADO", "VARIACAO_FINAL", "METODO_FONTE",
)
COLUNAS_MEMORIA_CALCULO = ("J", "K", "L", "M", "N", "O", "P", "Q", "R")
LINHA_INICIO_MEMORIA = 2
LINHA_FIM_MEMORIA = 80
CAPACIDADE_MEMORIA_CALCULO = LINHA_FIM_MEMORIA - LINHA_INICIO_MEMORIA + 1  # 79


def _competencia_iso(valor: Any) -> str | None:
    """Converte Timestamp/datetime/date em ISO (YYYY-MM-DD), sem inventar dados."""
    if valor is None:
        return None
    if hasattr(valor, "to_pydatetime"):
        valor = valor.to_pydatetime()
    if isinstance(valor, datetime):
        return valor.date().isoformat()
    if isinstance(valor, date):
        return valor.isoformat()
    return None


def _numero(valor: Any) -> float | None:
    if valor is None or isinstance(valor, bool):
        return None
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return None
    if numero != numero:  # NaN
        return None
    return numero


def normalizar_memoria_calculo(
    res: dict[str, Any] | None,
    fator_final: Any,
    variacao_final: Any,
) -> list[dict[str, Any]] | None:
    """Serializa res['dados'] (DataFrame exibido) em registros JSON-safe.

    fator_final e variacao_final sao os valores CANONICOS do payload do ciclo
    (fator e variacao ja decididos pela Calculadora); nada e recalculado aqui.
    Retorna None quando nao ha memoria (ciclo nao calculado): o ciclo e omitido.
    """
    if not isinstance(res, dict):
        return None
    dados = res.get("dados")
    if dados is None or getattr(dados, "empty", True):
        return None
    colunas = [str(c) for c in getattr(dados, "columns", [])]
    metodo = str(res.get("metodo") or "").strip()
    fonte = str(res.get("sercodigo") or res.get("serie") or "").strip()
    metodo_fonte = f"{metodo} [{fonte}]" if metodo and fonte else (metodo or fonte)

    registros: list[dict[str, Any]] = []
    if "indice" in colunas:
        # IST: duas linhas de numero-indice; nunca 12 meses ficticios.
        for ordem, (_, linha) in enumerate(dados.iterrows(), start=1):
            registros.append({
                "tipo": "INDICE",
                "ordem": ordem,
                "competencia": _competencia_iso(linha.get("data")),
                "valor_indice": _numero(linha.get("indice")),
                "fator_mensal": None,
                "fator_acumulado": None,
                "variacao_final": None,
                "metodo_fonte": None,
            })
    elif "valor" in colunas:
        # ICTI/SGS: uma linha MES por competencia efetivamente presente na
        # memoria exibida. N = taxa mensal em decimal (valor da fonte / 100,
        # sem arredondamento). O/P somente quando a memoria ja os contem.
        tem_fator = "fator_mensal" in colunas
        tem_acum = "fator_acumulado_progressivo" in colunas
        for ordem, (_, linha) in enumerate(dados.iterrows(), start=1):
            taxa = _numero(linha.get("valor"))
            registros.append({
                "tipo": "MES",
                "ordem": ordem,
                "competencia": _competencia_iso(linha.get("data")),
                "valor_indice": taxa / 100 if taxa is not None else None,
                "fator_mensal": _numero(linha.get("fator_mensal")) if tem_fator else None,
                "fator_acumulado": _numero(linha.get("fator_acumulado_progressivo")) if tem_acum else None,
                "variacao_final": None,
                "metodo_fonte": None,
            })
    else:
        return None
    if not registros:
        return None

    registros.append({
        "tipo": "RESULTADO",
        "ordem": len(registros) + 1,
        "competencia": None,
        "valor_indice": None,
        "fator_mensal": None,
        "fator_acumulado": _numero(fator_final),
        "variacao_final": _numero(variacao_final),
        "metodo_fonte": metodo_fonte or None,
    })
    return registros


def template_possui_bloco_memoria(ws_parametros) -> bool:
    """True quando o template ja tem os cabecalhos J1:R1 (arquivos legados: False)."""
    atual = tuple(
        str(ws_parametros[f"{col}1"].value or "").strip()
        for col in COLUNAS_MEMORIA_CALCULO
    )
    return atual == CABECALHOS_MEMORIA_CALCULO


def escrever_memoria_calculo(ws_parametros, ciclos: dict[str, Any]) -> None:
    """Limpa J2:R80 e grava sequencialmente a memoria dos ciclos C0..C4.

    Template sem os cabecalhos (legado) e ignorado sem erro. Overflow acima
    da capacidade reservada gera erro controlado; nunca gravacao parcial
    silenciosa fora do bloco.
    """
    if not template_possui_bloco_memoria(ws_parametros):
        return

    for linha in range(LINHA_INICIO_MEMORIA, LINHA_FIM_MEMORIA + 1):
        for col in COLUNAS_MEMORIA_CALCULO:
            cell = ws_parametros[f"{col}{linha}"]
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue  # formula do template: nunca limpar
            if cell.value is not None:
                cell.value = None

    planos: list[tuple[str, dict[str, Any]]] = []
    for nome in ("C0", "C1", "C2", "C3", "C4"):
        memoria = (ciclos.get(nome) or {}).get("memoria_calculo")
        if not memoria:
            continue
        for registro in memoria:
            if isinstance(registro, dict):
                planos.append((nome, registro))

    if len(planos) > CAPACIDADE_MEMORIA_CALCULO:
        raise ValueError(
            f"Memoria de calculo com {len(planos)} linhas excede a capacidade "
            f"reservada de {CAPACIDADE_MEMORIA_CALCULO} linhas "
            f"(parametros!J{LINHA_INICIO_MEMORIA}:R{LINHA_FIM_MEMORIA})."
        )

    linha = LINHA_INICIO_MEMORIA
    for nome, registro in planos:
        tipo = str(registro.get("tipo") or "").strip().upper()
        ws_parametros[f"J{linha}"] = nome
        ws_parametros[f"K{linha}"] = tipo
        ordem = registro.get("ordem")
        if ordem is not None:
            ws_parametros[f"L{linha}"] = int(ordem)
        competencia = registro.get("competencia")
        if competencia:
            ws_parametros[f"M{linha}"] = datetime.fromisoformat(str(competencia))
            ws_parametros[f"M{linha}"].number_format = "mm/yyyy"
        valor = _numero(registro.get("valor_indice"))
        if valor is not None:
            ws_parametros[f"N{linha}"] = valor
            ws_parametros[f"N{linha}"].number_format = (
                "0.0000" if tipo == "INDICE" else "0.0000%"
            )
        fator_mensal = _numero(registro.get("fator_mensal"))
        if fator_mensal is not None:
            ws_parametros[f"O{linha}"] = fator_mensal
            ws_parametros[f"O{linha}"].number_format = "0.000000"
        fator_acumulado = _numero(registro.get("fator_acumulado"))
        if fator_acumulado is not None:
            ws_parametros[f"P{linha}"] = fator_acumulado
            ws_parametros[f"P{linha}"].number_format = "0.000000"
        variacao = _numero(registro.get("variacao_final"))
        if variacao is not None:
            ws_parametros[f"Q{linha}"] = variacao
            ws_parametros[f"Q{linha}"].number_format = "0.0000%"
        metodo_fonte = registro.get("metodo_fonte")
        if metodo_fonte:
            ws_parametros[f"R{linha}"] = str(metodo_fonte)
        linha += 1


def ler_memoria_calculo(ws_parametros) -> dict[str, list[dict[str, Any]]]:
    """Leitura opcional do bloco: arquivos legados sem cabecalho retornam {}."""
    if not template_possui_bloco_memoria(ws_parametros):
        return {}
    memoria: dict[str, list[dict[str, Any]]] = {}
    for linha in range(LINHA_INICIO_MEMORIA, LINHA_FIM_MEMORIA + 1):
        ciclo = str(ws_parametros[f"J{linha}"].value or "").strip()
        tipo = str(ws_parametros[f"K{linha}"].value or "").strip().upper()
        if not ciclo or not tipo:
            continue
        registro = {
            "tipo": tipo,
            "ordem": _numero(ws_parametros[f"L{linha}"].value),
            "competencia": _competencia_iso(ws_parametros[f"M{linha}"].value),
            "valor_indice": _numero(ws_parametros[f"N{linha}"].value),
            "fator_mensal": _numero(ws_parametros[f"O{linha}"].value),
            "fator_acumulado": _numero(ws_parametros[f"P{linha}"].value),
            "variacao_final": _numero(ws_parametros[f"Q{linha}"].value),
            "metodo_fonte": str(ws_parametros[f"R{linha}"].value or "").strip() or None,
        }
        if registro["ordem"] is not None:
            registro["ordem"] = int(registro["ordem"])
        memoria.setdefault(ciclo, []).append(registro)
    return memoria
