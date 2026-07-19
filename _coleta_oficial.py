"""Template oficial do XLS Coleta (modelo novo, com aba posicao_contratual).

Fonte de verdade estrutural: templates/COLETA_REAJUSTE_OFICIAL.xlsx
(SHA-256 8569357302273d874aefc3cf7e5e25af6d60d216bc94457f953d0dc4a9666c6f,
apos rodada de UX de 17/07/2026 via Excel real: destaque CONTROLE!B1 +
protecao, itens_PC coluna A em Texto com estilo de entrada, DATA_PC
dd/mm/aaaa, CICLO_PC com alerta de data invalida, V:AC ocultas,
posicao_contratual com cores por ciclo, RESULTADOS memoria na linha 52 e
RESULTADOS!B4 com destaque FFF7E7B2 — mesmo padrao de CONTROLE!B1 —
dropdown/validacao/valor/borda preservados).

Correcao aprovada sobre o arquivo fornecido (Coleta_Reajuste_Nova.xlsx,
SHA-256 6353357b...8069ff3): o original possuia uma OMISSAO CRITICA — a
aba itens_PC vinha sem NUMERO_PC, eliminando o identificador documental e
o controle de duplicidade homologados. O template oficial corrigido
reintroduz NUMERO_PC na coluna A (ITEM permanece removido), desloca
formulas/validacoes/estilos/referencias e amplia o CHECK_PC_FINANCEIRO
com a regra de duplicidade de NUMERO_PC (TRIM+UPPER; vazio nao avalia).
A grade do financeiro foi estendida ate a linha 73 (72 competencias).

Este modulo NAO reconstroi o XLS celula a celula: carrega o template oficial
e apenas limpa os dados demonstrativos das celulas de entrada, preservando
todas as formulas, validacoes, estilos, merges e intervalos nomeados.
"""
from __future__ import annotations

import hashlib
import re
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from dateutil.relativedelta import relativedelta

ROOT = Path(__file__).resolve().parent
TEMPLATE_COLETA_OFICIAL = ROOT / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
NOME_ARQUIVO_COLETA_OFICIAL = "COLETA_REAJUSTE_OFICIAL.xlsx"

# Ordem oficial das abas do novo modelo (fonte de verdade: o proprio XLS)
ABAS_COLETA_OFICIAL = [
    "CONTROLE",
    "parametros",
    "financeiro",
    "itens_Remanesc",
    "itens_Consumidos",
    "itens_PC",
    "aditivos",
    "posicao_contratual",
    "itens_RC",
    "historico_VU",
    "RESULTADOS",
]

ABA_POSICAO_CONTRATUAL = "posicao_contratual"

# Colunas oficiais de itens_PC (linha 1, A:L) — NUMERO_PC obrigatorio na
# coluna A como identificador documental e chave de duplicidade. ITEM nao
# existe no modelo oficial (decisao aprovada; nao restaurar).
COLUNAS_ITENS_PC_OFICIAL = [
    "NUMERO_PC",
    "DATA_PC",
    "CICLO_PC",
    "VALOR_PC",
    "FATOR_ACUMULADO",
    "VALOR_ATUALIZADO",
    "PC_PAGO_A_CONTRATADA",
    "RETROATIVO_RECONHECIDO_A_PAGAR",
    "VALOR_ATUALIZADO_EM_ANALISE",
    "DELTA_POTENCIAL",
    "CHECK_PC_FINANCEIRO",
    "EFEITO_FINANCEIRO_PC",
]

# Cabecalhos essenciais da aba posicao_contratual (linha 1, A:X)
COLUNAS_POSICAO_CONTRATUAL = (
    ["ITEM", "VU_ORIGINAL", "QTD_BASE_ORIGINAL"]
    + [
        rotulo
        for n in range(5)
        for rotulo in (
            f"DELTA_C{n}",
            f"QTD_CONTRATADA_C{n}",
            f"QTD_REM_BASE_C{n}",
            f"QTD_REM_AJUSTADA_C{n}",
        )
    ]
    + ["CHECK_POSICAO_CONTRATUAL"]
)

# Celulas de entrada com dados demonstrativos no template — limpar na geracao.
# Somente celulas de VALOR sao limpas; formulas nunca sao tocadas.
_RESIDUOS_POR_ABA: dict[str, list[str]] = {
    # B2=ciclo vigente demo, B3=data de corte demo, B7=indice, B8=data-base
    "CONTROLE": ["B2", "B3", "B7", "B8"],
}
# parametros: linhas 2-6, colunas de entrada (B=CICLO e F=formula preservados)
_RESIDUOS_POR_ABA["parametros"] = [
    f"{col}{lin}" for lin in range(2, 7) for col in ("A", "C", "D", "E", "G", "H")
]
# financeiro: competencia (A), valor pago (C) e efeito (G) — linhas 2-73
# (grade estendida para 72 competencias mensais)
_RESIDUOS_POR_ABA["financeiro"] = [
    f"{col}{lin}" for lin in range(2, 74) for col in ("A", "C", "G")
]
# itens_Remanesc: entradas do fiscal (item, base, VU, remanescentes por ciclo)
_RESIDUOS_POR_ABA["itens_Remanesc"] = [
    f"{col}{lin}" for lin in range(2, 201)
    for col in ("A", "B", "C", "E", "G", "I", "K")
]
# aditivos: eventos demonstrativos (F soh quando valor manual sobrepoe formula)
_RESIDUOS_POR_ABA["aditivos"] = [
    f"{col}{lin}" for lin in range(2, 201)
    for col in ("A", "B", "D", "E", "F", "H", "K")
]


def assinatura_template_coleta(caminho: str | Path | None = None) -> str:
    """SHA-256 dos bytes do template oficial — assinatura robusta de cache.

    Muda sempre que o conteudo binario do XLS muda, mesmo com nome, versao
    da aplicacao, tamanho e timestamp identicos. Template ausente/ilegivel
    retorna "" (o consumidor decide como tratar; obter_coleta_oficial_bytes
    ja falha explicitamente nesse caso).
    """
    alvo = Path(caminho) if caminho is not None else TEMPLATE_COLETA_OFICIAL
    try:
        return hashlib.sha256(alvo.read_bytes()).hexdigest()
    except OSError:
        return ""


def eh_layout_coleta_oficial(wb) -> bool:
    """True quando o workbook segue o novo modelo (aba posicao_contratual)."""
    return ABA_POSICAO_CONTRATUAL in wb.sheetnames


def _limpar_residuos(wb) -> None:
    for aba, celulas in _RESIDUOS_POR_ABA.items():
        if aba not in wb.sheetnames:
            continue
        ws = wb[aba]
        for coord in celulas:
            cell = ws[coord]
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue  # formula do template: nunca limpar
            if cell.value is not None:
                cell.value = None


def _validar_estrutura_itens_pc(wb) -> None:
    """Barreira contra regressao critica: itens_PC jamais pode sair esvaziada.

    Valida cabecalhos A1:L1 e a densidade de formulas da grade (C:L, linhas
    2-100). Se o template em disco for trocado por uma versao sem a estrutura
    homologada, a geracao falha explicitamente em vez de entregar o XLS.
    """
    ws = wb["itens_PC"]
    cabecalhos = [ws.cell(1, c).value for c in range(1, 13)]
    if cabecalhos != COLUNAS_ITENS_PC_OFICIAL:
        raise ValueError(
            "itens_PC invalida: cabecalhos A1:L1 divergem do modelo oficial "
            f"(encontrado: {cabecalhos})"
        )
    formulas = sum(
        1
        for row in ws.iter_rows(min_row=2, max_row=100, min_col=3, max_col=12)
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    )
    if formulas < 700:
        raise ValueError(
            f"itens_PC invalida: apenas {formulas} formulas na grade C2:L100 "
            "(estrutura esvaziada)"
        )
    # Visao salva rolada (ex.: topLeftCell=AD1 com V:AC ocultas e sem grade)
    # faz a aba parecer completamente vazia no Excel — causa raiz do bug de
    # 17/07/2026. A visao deve sempre abrir em A1.
    tlc = ws.sheet_view.topLeftCell
    if tlc not in (None, "A1"):
        raise ValueError(
            f"itens_PC invalida: visao salva rolada para {tlc} — a aba "
            "abriria visualmente vazia; corrija o template (topLeftCell=A1)"
        )


def obter_coleta_oficial_bytes() -> bytes:
    """Retorna o novo XLS Coleta oficial em branco (estrutura integral).

    Carrega o template oficial com data_only=False (formulas preservadas)
    e limpa apenas os dados demonstrativos das celulas de entrada.
    """
    if not TEMPLATE_COLETA_OFICIAL.exists():
        raise FileNotFoundError(
            f"Template oficial nao encontrado: {TEMPLATE_COLETA_OFICIAL}"
        )
    wb = load_workbook(TEMPLATE_COLETA_OFICIAL, data_only=False)

    faltantes = [a for a in ABAS_COLETA_OFICIAL if a not in wb.sheetnames]
    if faltantes:
        raise ValueError(
            f"Template oficial invalido; abas ausentes: {', '.join(faltantes)}"
        )

    _limpar_residuos(wb)
    _validar_estrutura_itens_pc(wb)

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True

    saida = BytesIO()
    wb.save(saida)
    return saida.getvalue()


def _data(valor: Any) -> date | None:
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


def _numero_ciclo(valor: Any) -> int | None:
    match = re.search(r"\bC\s*([0-4])\b", str(valor or "").upper())
    return int(match.group(1)) if match else None


def _percentual(ciclo: dict[str, Any]) -> float | None:
    for chave in ("percentual_aplicado", "percentual_indice", "percentual", "variacao"):
        valor = ciclo.get(chave)
        if valor in (None, "") or isinstance(valor, bool):
            continue
        try:
            numero = float(valor)
        except (TypeError, ValueError):
            continue
        return numero / 100 if abs(numero) > 1 else numero
    try:
        fator = float(ciclo.get("fator"))
    except (TypeError, ValueError):
        return None
    return fator - 1 if fator >= 0.5 else fator


def normalizar_dados_calculadora(dados: dict[str, Any] | None) -> dict[str, Any]:
    """Adapta o estado atual das Calculadoras ao gerador do modelo oficial."""
    origem = dict(dados or {})
    fornecidos: dict[int, dict[str, Any]] = {}
    for bruto in origem.get("ciclos") or []:
        if not isinstance(bruto, dict):
            continue
        numero = _numero_ciclo(bruto.get("ciclo") or bruto.get("Ciclo"))
        if numero is not None and numero > 0:
            fornecidos[numero] = bruto
    if not fornecidos:
        raise ValueError("A Calculadora não informou nenhum ciclo entre C1 e C4.")

    data_base = _data(origem.get("data_base") or origem.get("data_base_original"))
    inicios: dict[int, date] = {}
    for numero, bruto in fornecidos.items():
        inicio = _data(bruto.get("data_inicio") or bruto.get("inicio_ciclo"))
        if inicio is None:
            ancora = _data(bruto.get("data_base") or bruto.get("periodo_inicio"))
            inicio = ancora + relativedelta(months=12) if ancora else None
        if inicio:
            inicios[numero] = inicio.replace(day=1)
    if not inicios and data_base:
        primeiro = min(fornecidos)
        inicios[primeiro] = data_base.replace(day=1) + relativedelta(months=12 * primeiro)
    if not inicios:
        raise ValueError("Não foi possível identificar a data-base dos ciclos da Calculadora.")

    ultimo = max(fornecidos)
    for numero in range(0, ultimo + 1):
        if numero in inicios:
            continue
        referencia = min(inicios, key=lambda existente: abs(existente - numero))
        inicios[numero] = inicios[referencia] + relativedelta(months=12 * (numero - referencia))

    ciclos = []
    for numero in range(1, ultimo + 1):
        bruto = fornecidos.get(numero, {})
        inicio = inicios[numero]
        fim = _data(bruto.get("data_fim")) or (inicio + relativedelta(months=12) - relativedelta(days=1))
        objeto_atual = bool(
            numero in fornecidos
            and bruto.get("objeto_analise_atual", not bruto.get("ciclo_ja_concedido", False))
        )
        ciclos.append({
            **bruto,
            "ciclo": f"C{numero}",
            "data_inicio": inicio,
            "data_fim": fim,
            "data_pedido": _data(bruto.get("data_pedido")),
            # Data final ja decidida pela Calculadora. O gerador apenas a
            # propaga; nao recria tempestividade, negociacao ou excecoes.
            "inicio_efeito_financeiro": _data(
                bruto.get("inicio_efeito_financeiro")
                or bruto.get("financeiro_inicio")
                or bruto.get("inicio_financeiro")
            ),
            "percentual": _percentual(bruto),
            "possui_efeito_financeiro": "Sim" if objeto_atual else "Não",
            "situacao": bruto.get("situacao_aplicada") or bruto.get("situacao") or "",
        })

    marco_inicial = data_base or inicios.get(0) or min(inicios.values())
    data_corte = _data(origem.get("data_corte")) or max(c["data_fim"] for c in ciclos)
    return {
        **origem,
        "ok": True,
        "modo_origem": origem.get("modo_origem") or origem.get("origem") or origem.get("tipo") or "Calculadora",
        "data_base": marco_inicial,
        "data_corte": data_corte,
        "ciclo_vigente": f"C{ultimo}",
        "ciclos": ciclos,
    }


def gerar_coleta_oficial_preenchida(dados_calculadora: dict[str, Any] | None) -> bytes:
    """Gera o XLS oficial; preenche marcos quando a Calculadora tem dados."""
    base = obter_coleta_oficial_bytes()
    if not (dados_calculadora or {}).get("ciclos"):
        return base
    from _gerador_masterfile import gerar_masterfile_preenchido

    return gerar_masterfile_preenchido(
        normalizar_dados_calculadora(dados_calculadora),
        base,
    )
