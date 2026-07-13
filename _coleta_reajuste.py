"""Motor canônico do XLS-first do Master 2.0.

O Excel é a fonte de verdade: este módulo apenas preenche os marcos já
apurados pela calculadora e valida a estrutura no retorno. Os resultados
financeiros permanecem fórmulas da própria planilha.
"""

from __future__ import annotations

from copy import copy
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
import re
import unicodedata
from typing import Any

from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill


NOME_ARQUIVO_COLETA = "Coleta_Reajuste.xlsx"
CAMINHO_MODELO_COLETA = Path(__file__).resolve().parent / "templates" / NOME_ARQUIVO_COLETA

ABAS_CANONICAS = (
    "CONTROLE",
    "parametros",
    "financeiro",
    "itens_Remanesc",
    "itens_Consumidos",
    "itens_PC",
    "aditivos",
    "RESULTADOS",
    "itens_RC",
    "historico_VU",
)

ABAS_PROIBIDAS = ("itens_Execucao_Saldo", "Itens_Execução", "REGRA_NEGOCIO_CLAUS", "Regra")
ERROS_EXCEL = {"#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}

COR_TEXTO = "FF595959"
COR_MARINHO = "FF123B63"
PREENCHIMENTO_AUTOMATICO = PatternFill("solid", fgColor="FFEDEDED")
PREENCHIMENTO_ENTRADA = PatternFill("solid", fgColor="FFFFF2CC")


def _texto_sem_acento(valor: Any) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    return "".join(ch for ch in texto if not unicodedata.combining(ch))


def _data(valor: Any) -> datetime | None:
    if valor in (None, ""):
        return None
    if isinstance(valor, datetime):
        return valor.replace(tzinfo=None)
    if isinstance(valor, date):
        return datetime(valor.year, valor.month, valor.day)
    texto = str(valor).strip()
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%Y"):
        try:
            return datetime.strptime(texto, formato)
        except ValueError:
            continue
    return None


def _primeiro_dia_mes(valor: datetime) -> datetime:
    return valor.replace(day=1, hour=0, minute=0, second=0, microsecond=0)


def _numero_ciclo(valor: Any) -> int | None:
    match = re.search(r"\bC\s*([0-4])\b", str(valor or "").upper())
    return int(match.group(1)) if match else None


def _numero(valor: Any) -> float | None:
    if valor in (None, "") or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace("%", "").replace(" ", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def _percentual_ciclo(ciclo: dict[str, Any]) -> float | None:
    for chave in ("percentual_aplicado", "percentual_indice", "variacao"):
        valor = _numero(ciclo.get(chave))
        if valor is not None:
            return valor / 100 if abs(valor) > 1 else valor
    fator = _numero(ciclo.get("fator"))
    if fator is not None:
        return fator - 1 if fator >= 0.5 else fator
    return None


def _ciclo_em_analise(ciclo: dict[str, Any]) -> bool:
    if "objeto_analise_atual" in ciclo:
        return bool(ciclo.get("objeto_analise_atual"))
    if "ciclo_ja_concedido" in ciclo:
        return not bool(ciclo.get("ciclo_ja_concedido"))
    return True


def _inicio_teorico(ciclo: dict[str, Any]) -> datetime | None:
    inicio_direto = _data(ciclo.get("inicio_ciclo"))
    if inicio_direto:
        return _primeiro_dia_mes(inicio_direto)
    # Na calculadora, data_base/periodo_inicio é o início da janela do índice;
    # o ciclo contratual se inicia no aniversário, doze meses depois.
    ancora = _data(ciclo.get("data_base")) or _data(ciclo.get("periodo_inicio"))
    if ancora:
        return _primeiro_dia_mes(ancora + relativedelta(months=12))
    return None


def _montar_ciclos(dados: dict[str, Any]) -> tuple[list[dict[str, Any]], set[int], list[str]]:
    ciclos_origem = dados.get("ciclos") or []
    fornecidos: dict[int, dict[str, Any]] = {}
    for ciclo in ciclos_origem:
        if not isinstance(ciclo, dict):
            continue
        numero = _numero_ciclo(ciclo.get("ciclo") or ciclo.get("Ciclo"))
        if numero is not None and numero > 0:
            fornecidos[numero] = ciclo
    if not fornecidos:
        raise ValueError("A calculadora não informou nenhum ciclo entre C1 e C4.")

    ultimo = max(fornecidos)
    alvos = {n for n, ciclo in fornecidos.items() if _ciclo_em_analise(ciclo)}
    if not alvos:
        raise ValueError("Nenhum ciclo foi marcado como objeto desta apuração.")

    inicios = {n: inicio for n, ciclo in fornecidos.items() if (inicio := _inicio_teorico(ciclo))}
    if not inicios:
        ancora = _data(dados.get("data_base_original"))
        if not ancora:
            raise ValueError("Não foi possível identificar a data-base dos ciclos.")
        inicios[min(fornecidos)] = _primeiro_dia_mes(ancora + relativedelta(months=12))

    # Preenche lacunas em blocos anuais, preservando os marcos explícitos da calculadora.
    for numero in range(0, ultimo + 1):
        if numero in inicios:
            continue
        referencia = min(inicios, key=lambda existente: abs(existente - numero))
        inicios[numero] = inicios[referencia] + relativedelta(months=12 * (numero - referencia))

    contexto = dados.get("contexto_contratual_anterior") or {}
    ultimo_contexto = _numero_ciclo(contexto.get("ultimo_ciclo_concedido"))
    percentual_contexto = _numero(contexto.get("percentual_ja_aplicado_pct"))
    if percentual_contexto is not None and abs(percentual_contexto) > 1:
        percentual_contexto /= 100

    alertas: list[str] = []
    saida: list[dict[str, Any]] = []
    for numero in range(0, ultimo + 1):
        origem = fornecidos.get(numero, {})
        percentual = 0.0 if numero == 0 else _percentual_ciclo(origem)
        if numero > 0 and percentual is None and numero == ultimo_contexto:
            percentual = percentual_contexto
        inicio = _primeiro_dia_mes(inicios[numero])
        fim = inicio + relativedelta(months=12) - relativedelta(days=1)
        alvo = numero in alvos
        if numero > 0 and percentual is None and numero <= max(alvos):
            alertas.append(
                f"C{numero}: percentual histórico não informado; resultados acumulados dependentes ficarão em branco."
            )
        situacao = (
            "Base"
            if numero == 0
            else str(origem.get("situacao_aplicada") or origem.get("situacao") or "").strip()
            or ("Em análise" if alvo else "Histórico fora desta apuração")
        )
        saida.append(
            {
                "numero": numero,
                "nome": f"C{numero}",
                "inicio": inicio,
                "fim": fim,
                "percentual": percentual,
                "situacao": situacao,
                "computar": alvo,
                "financeiro_inicio": _data(origem.get("financeiro_inicio")),
            }
        )
    return saida, alvos, alertas


def _formulas(wb) -> dict[str, str]:
    return {
        f"{ws.title}!{cell.coordinate}": cell.value
        for ws in wb.worksheets
        for row in ws.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    }


def _normalizar_arquivo(wb) -> None:
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                cell.comment = None
        ws.sheet_view.showGridLines = False
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.active = 0
    if wb.views:
        wb.views[0].activeTab = 0
        wb.views[0].firstSheet = 0


def gerar_coleta_reajuste(dados_admissibilidade: dict[str, Any]) -> bytes:
    """Preenche o modelo canônico sem substituir ou calcular suas fórmulas."""

    if not CAMINHO_MODELO_COLETA.exists():
        raise FileNotFoundError(f"Modelo canônico não encontrado: {CAMINHO_MODELO_COLETA}")
    dados = dados_admissibilidade or {}
    ciclos, alvos, _alertas = _montar_ciclos(dados)

    wb = load_workbook(CAMINHO_MODELO_COLETA, data_only=False)
    if tuple(wb.sheetnames) != ABAS_CANONICAS:
        raise ValueError("O modelo canônico possui abas inesperadas ou fora de ordem.")
    formulas_originais = _formulas(wb)

    ws = wb["CONTROLE"]
    ws["B1"] = "Principal"
    ws["B2"] = f"C{max(alvos)}"
    ws["B3"] = max(ciclo["fim"] for ciclo in ciclos if ciclo["numero"] <= max(alvos))
    ws["B3"].number_format = "mm/yyyy"
    ws["B7"] = str(dados.get("indice") or "").strip()
    ws["B8"] = ciclos[0]["inicio"]
    ws["B8"].number_format = "mm/yyyy"

    ws = wb["parametros"]
    for ciclo in ciclos:
        row = ciclo["numero"] + 2
        ws[f"A{row}"] = "Sim" if ciclo["computar"] else "Nao"
        ws[f"B{row}"] = ciclo["nome"]
        ws[f"C{row}"] = f'{ciclo["inicio"]:%m/%Y} a {ciclo["fim"]:%m/%Y}'
        ws[f"D{row}"] = ciclo["inicio"]
        ws[f"E{row}"] = ciclo["fim"]
        ws[f"F{row}"] = ciclo["percentual"]
        ws[f"H{row}"] = ciclo["situacao"]
        for col in ("A", "B", "C", "D", "E", "F", "H"):
            cell = ws[f"{col}{row}"]
            cell.fill = PREENCHIMENTO_AUTOMATICO
            font = copy(cell.font)
            font.color = COR_MARINHO if ciclo["computar"] else COR_TEXTO
            font.b = bool(ciclo["computar"])
            cell.font = font
        ws[f"C{row}"].number_format = "@"
        for col in ("D", "E"):
            ws[f"{col}{row}"].number_format = "mm/yyyy"
        ws[f"F{row}"].number_format = "0.00%"

    # C1-C4 ainda não alcançados permanecem estruturalmente presentes, mas vazios.
    ultimo = max(ciclo["numero"] for ciclo in ciclos)
    for numero in range(ultimo + 1, 5):
        row = numero + 2
        ws[f"A{row}"] = "Nao"
        ws[f"B{row}"] = f"C{numero}"
        for col in ("C", "D", "E", "F"):
            ws[f"{col}{row}"] = None
        ws[f"H{row}"] = "Não aplicável"

    ws = wb["financeiro"]
    row = 2
    for ciclo in ciclos:
        for deslocamento in range(12):
            competencia = ciclo["inicio"] + relativedelta(months=deslocamento)
            ws[f"A{row}"] = competencia
            ws[f"A{row}"].number_format = "mm/yyyy"
            ws[f"B{row}"] = ciclo["nome"].lower()
            ws[f"C{row}"] = None
            financeiro_inicio = ciclo["financeiro_inicio"]
            efeito = bool(
                ciclo["computar"]
                and financeiro_inicio
                and _primeiro_dia_mes(competencia) >= _primeiro_dia_mes(financeiro_inicio)
            )
            ws[f"G{row}"] = "Sim" if efeito else "Nao"
            for col in ("A", "B", "G"):
                cell = ws[f"{col}{row}"]
                cell.fill = PREENCHIMENTO_AUTOMATICO
                cell.font = Font(
                    name="Calibri",
                    size=10,
                    bold=bool(ciclo["computar"]),
                    color=COR_MARINHO if ciclo["computar"] else COR_TEXTO,
                )
            ws[f"C{row}"].fill = PREENCHIMENTO_ENTRADA
            row += 1
    while row <= 61:
        for col in ("A", "B", "C", "G"):
            ws[f"{col}{row}"] = None
        ws[f"C{row}"].fill = PREENCHIMENTO_ENTRADA
        row += 1

    _normalizar_arquivo(wb)
    formulas_finais = _formulas(wb)
    if formulas_finais != formulas_originais:
        alteradas = sorted(set(formulas_originais) ^ set(formulas_finais))[:5]
        raise RuntimeError(f"A geração alterou a matriz de fórmulas: {alteradas}")

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def _celula_tem_observacao(valor: Any) -> bool:
    texto = _texto_sem_acento(valor).upper()
    return "OBSERVACAO" in texto or texto.strip() == "OBS"


def ler_coleta_reajuste(conteudo: bytes) -> dict[str, Any]:
    """Valida o XLS no upload sem recalcular nem substituir resultados do Excel."""

    if not conteudo:
        raise ValueError("Arquivo vazio.")
    try:
        wb = load_workbook(BytesIO(conteudo), data_only=False, read_only=False)
    except Exception as exc:
        raise ValueError("O arquivo não é um XLSX válido.") from exc

    faltantes = [aba for aba in ABAS_CANONICAS if aba not in wb.sheetnames]
    proibidas = [aba for aba in ABAS_PROIBIDAS if aba in wb.sheetnames]
    pendencias: list[str] = []
    avisos: list[str] = []
    if faltantes:
        pendencias.append("Abas obrigatórias ausentes: " + ", ".join(faltantes))
    if proibidas:
        pendencias.append("Abas excluídas reapareceram: " + ", ".join(proibidas))
    if faltantes:
        return {
            "valido": False,
            "pronto_para_consolidar": False,
            "pendencias": pendencias,
            "avisos": avisos,
            "contagens": {},
            "metadados": {},
        }

    formulas = _formulas(wb)
    if len(formulas) < 1000:
        pendencias.append("A matriz de fórmulas foi removida ou está incompleta.")
    for chave in ("financeiro!D2", "itens_Remanesc!D2", "itens_Consumidos!O2", "itens_PC!D2"):
        if chave not in formulas:
            pendencias.append(f"Fórmula estrutural ausente em {chave}.")
    referencias_quebradas = [chave for chave, formula in formulas.items() if "#REF!" in formula.upper()]
    if referencias_quebradas:
        pendencias.append("Há fórmulas com referência quebrada: " + ", ".join(referencias_quebradas[:5]))

    comentarios = []
    observacoes = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                if cell.comment is not None:
                    comentarios.append(f"{ws.title}!{cell.coordinate}")
                if cell.row == 1 and _celula_tem_observacao(cell.value):
                    observacoes.append(f"{ws.title}!{cell.coordinate}")
    if comentarios:
        pendencias.append("Comentários/observações de célula não são admitidos: " + ", ".join(comentarios[:5]))
    if observacoes:
        pendencias.append("Campos de observação não são admitidos: " + ", ".join(observacoes[:5]))

    parametros = wb["parametros"]
    ativos = [numero for numero in range(1, 5) if str(parametros[f"A{numero + 2}"].value).strip().lower() == "sim"]
    if not ativos:
        pendencias.append("Nenhum ciclo está marcado para computar nesta apuração.")
    else:
        for numero in range(1, max(ativos) + 1):
            if _numero(parametros[f"F{numero + 2}"].value) is None:
                pendencias.append(f"C{numero}: percentual necessário ao acumulado está ausente.")

    contagens = {
        "competencias_com_valor": sum(1 for row in range(2, 62) if _numero(wb["financeiro"][f"C{row}"].value) is not None),
        "itens_remanescentes": sum(1 for row in range(2, 201) if wb["itens_Remanesc"][f"A{row}"].value not in (None, "")),
        "itens_consumidos": sum(1 for row in range(2, 201) if wb["itens_Consumidos"][f"A{row}"].value not in (None, "")),
        "pedidos_de_compra": sum(1 for row in range(2, 101) if wb["itens_PC"][f"B{row}"].value not in (None, "")),
        "aditivos": sum(1 for row in range(2, 201) if wb["aditivos"][f"A{row}"].value not in (None, "")),
        "formulas": len(formulas),
    }
    if contagens["competencias_com_valor"] == 0 and contagens["itens_remanescentes"] == 0:
        avisos.append("Ainda não há valores mensais nem itens remanescentes preenchidos pelo fiscal.")

    try:
        wb_valores = load_workbook(BytesIO(conteudo), data_only=True, read_only=True)
        erros = []
        for ws in wb_valores.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.upper() in ERROS_EXCEL:
                        erros.append(f"{ws.title}!{cell.coordinate}={cell.value}")
        if erros:
            pendencias.append("O Excel salvou erros de cálculo: " + ", ".join(erros[:8]))
    except Exception:
        avisos.append("Não foi possível conferir os valores calculados em cache; abra e salve o arquivo no Excel.")

    metadados = {
        "indice": wb["CONTROLE"]["B7"].value,
        "ciclo_vigente": wb["CONTROLE"]["B2"].value,
        "data_corte": wb["CONTROLE"]["B3"].value,
        "ciclos_em_analise": [f"C{numero}" for numero in ativos],
    }
    return {
        "valido": not pendencias,
        "pronto_para_consolidar": not pendencias and (
            contagens["competencias_com_valor"] > 0 or contagens["itens_remanescentes"] > 0
        ),
        "pendencias": pendencias,
        "avisos": avisos,
        "contagens": contagens,
        "metadados": metadados,
    }


def eh_coleta_reajuste(conteudo: bytes) -> bool:
    try:
        wb = load_workbook(BytesIO(conteudo), read_only=True, data_only=False)
        nomes = set(wb.sheetnames)
        nucleares = {"CONTROLE", "parametros", "financeiro"}
        # Também reconhece uma coleta canônica danificada, para que o validador
        # possa explicar a aba ausente em vez de desviá-la ao leitor legado.
        return nucleares.issubset(nomes) and len(nomes.intersection(ABAS_CANONICAS)) >= 5
    except Exception:
        return False
