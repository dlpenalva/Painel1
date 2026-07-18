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

from _capacidades_apuracao import avaliar_capacidades_apuracao


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
    "posicao_contratual",
    "itens_RC",
    "historico_VU",
    "RESULTADOS",
)

ABAS_OBRIGATORIAS_LEGADO = tuple(
    aba for aba in ABAS_CANONICAS if aba != "posicao_contratual"
)

NOMES_RESULTADOS_OBRIGATORIOS = {
    "METODO_RETROATIVO",
    "TOLERANCIA_DIVERGENCIA",
    "VALOR_MANUAL_RETRO",
    "JUSTIFICATIVA_RETRO",
    "RETRO_FIN",
    "RETRO_PC",
    "RETRO_ITENS",
    "RETRO_OFICIAL",
    "VTA_CALCULADO",
    "AJUSTE_MANUAL_VTA",
    "VTA_MANUAL_OFICIAL",
    "VTA_FINAL",
    "QTD_REM_OFICIAL",
    "REM_BASE_OFICIAL",
    "REM_ATUALIZADO_OFICIAL",
}

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


def _inicio_efeito_definido(wb, ciclo: str) -> date | None:
    texto = str(wb.properties.keywords or "")
    match = re.search(rf"(?:^|[:,;]){re.escape(ciclo)}=(\d{{4}}-\d{{2}}-\d{{2}})", texto)
    if not match:
        return None
    try:
        return date.fromisoformat(match.group(1))
    except ValueError:
        return None


def _tem_metadado_inicio_efeito(wb) -> bool:
    return "CL8US_INICIO_EFEITO:" in str(wb.properties.keywords or "")


def _ciclo_por_competencia_financeira(wb, competencia: Any) -> str:
    data_comp = _data(competencia)
    if data_comp is None:
        return ""
    ws = wb["parametros"]
    for row in range(2, 7):
        ciclo = str(ws[f"B{row}"].value or "").strip().upper()
        inicio = _data(ws[f"C{row}"].value)
        fim = _data(ws[f"D{row}"].value)
        if ciclo and inicio and fim and inicio <= data_comp <= fim:
            return ciclo
    return ""


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


def _validar_resultados_integra(wb, etapa: str) -> dict[str, Any]:
    if "RESULTADOS" not in wb.sheetnames:
        raise ValueError(f"A aba RESULTADOS desapareceu na etapa {etapa}.")
    ws = wb["RESULTADOS"]
    formulas = sum(
        1
        for row in ws.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    )
    conteudo = sum(1 for row in ws.iter_rows() for cell in row if cell.value not in (None, ""))
    if ws.sheet_state != "visible":
        raise ValueError(f"A aba RESULTADOS não está visível na etapa {etapa}.")
    if ws["A1"].value != "RESULTADOS CONSOLIDADOS — REAJUSTE CONTRATUAL":
        raise ValueError(f"A aba RESULTADOS está vazia ou foi substituída na etapa {etapa}.")
    if formulas < 3000 or conteudo < 3300:
        raise ValueError(
            f"A aba RESULTADOS perdeu conteúdo na etapa {etapa}: "
            f"{formulas} fórmulas e {conteudo} células preenchidas."
        )
    return {"visivel": True, "formulas": formulas, "conteudo": conteudo}


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
    _validar_resultados_integra(wb, "logo após o carregamento do template")
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
        ws[f"C{row}"] = ciclo["inicio"]
        ws[f"D{row}"] = ciclo["fim"]
        ws[f"E{row}"] = ciclo["percentual"]
        ws[f"G{row}"] = ciclo["situacao"]
        for col in ("A", "B", "C", "D", "E", "G"):
            cell = ws[f"{col}{row}"]
            cell.fill = PREENCHIMENTO_AUTOMATICO
            font = copy(cell.font)
            font.color = COR_MARINHO if ciclo["computar"] else COR_TEXTO
            font.b = bool(ciclo["computar"])
            cell.font = font
        for col in ("C", "D"):
            ws[f"{col}{row}"].number_format = "mm/yyyy"
        ws[f"E{row}"].number_format = "0.00%"

    # C1-C4 ainda não alcançados permanecem estruturalmente presentes, mas vazios.
    ultimo = max(ciclo["numero"] for ciclo in ciclos)
    for numero in range(ultimo + 1, 5):
        row = numero + 2
        ws[f"A{row}"] = "Nao"
        ws[f"B{row}"] = f"C{numero}"
        for col in ("C", "D", "E"):
            ws[f"{col}{row}"] = None
        ws[f"G{row}"] = "Não aplicável"

    ws = wb["financeiro"]
    row = 2
    for ciclo in ciclos:
        for deslocamento in range(12):
            competencia = ciclo["inicio"] + relativedelta(months=deslocamento)
            ws[f"A{row}"] = competencia
            ws[f"A{row}"].number_format = "mm/yyyy"
            ws[f"C{row}"] = None
            financeiro_inicio = ciclo["financeiro_inicio"]
            efeito = bool(
                ciclo["computar"]
                and financeiro_inicio
                and _primeiro_dia_mes(competencia) >= _primeiro_dia_mes(financeiro_inicio)
            )
            ws[f"G{row}"] = "Sim" if efeito else "Nao"
            for col in ("A", "G"):
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
        for col in ("A", "C", "G"):
            ws[f"{col}{row}"] = None
        ws[f"C{row}"].fill = PREENCHIMENTO_ENTRADA
        row += 1

    _normalizar_arquivo(wb)
    _validar_resultados_integra(wb, "imediatamente antes do salvamento")
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

    faltantes = [aba for aba in ABAS_OBRIGATORIAS_LEGADO if aba not in wb.sheetnames]
    proibidas = [aba for aba in ABAS_PROIBIDAS if aba in wb.sheetnames]
    bloqueios_estruturais: list[str] = []
    bloqueios_criticos: list[str] = []
    lacunas_apuracao: list[str] = []
    avisos: list[str] = []
    if faltantes:
        bloqueios_estruturais.append("Abas obrigatórias ausentes: " + ", ".join(faltantes))
    if proibidas:
        bloqueios_estruturais.append("Abas excluídas reapareceram: " + ", ".join(proibidas))
    if faltantes:
        capacidades = avaliar_capacidades_apuracao({}, {}, bloqueios_estruturais, [])
        return {
            "valido": False,
            "pronto_para_consolidar": False,
            "processamento_progressivo": True,
            "pendencias": bloqueios_estruturais,
            "bloqueios_estruturais": bloqueios_estruturais,
            "bloqueios_criticos": bloqueios_criticos,
            "lacunas_apuracao": [],
            "avisos": avisos,
            "contagens": {},
            "metadados": {},
            "capacidades": capacidades,
        }

    possui_posicao_contratual = "posicao_contratual" in wb.sheetnames
    if not possui_posicao_contratual:
        avisos.append(
            "Arquivo legado sem a camada posicao_contratual; quantidades por ciclo seguem o leiaute historico."
        )

    # Detecta modelo oficial: NUMERO_PC na coluna A de itens_PC desloca CICLO_PC para C2
    _ws_ipc = wb["itens_PC"] if "itens_PC" in wb.sheetnames else None
    _header_a1 = (_ws_ipc["A1"].value or "") if _ws_ipc is not None else ""
    _chave_ciclo_pc = "itens_PC!C2" if str(_header_a1).strip().upper() == "NUMERO_PC" else "itens_PC!B2"

    formulas = _formulas(wb)
    if len(formulas) < 1000:
        bloqueios_estruturais.append("A matriz de fórmulas foi removida ou está incompleta.")
    for chave in (
        "financeiro!D2",
        "itens_Remanesc!D2",
        "itens_Consumidos!O2",
        _chave_ciclo_pc,
        "RESULTADOS!B15",
        "RESULTADOS!B16",
        "RESULTADOS!B23",
        "RESULTADOS!B26",
        "RESULTADOS!B35",
        "RESULTADOS!C35",
        "RESULTADOS!D35",
        "RESULTADOS!F36",
    ):
        if chave not in formulas:
            bloqueios_estruturais.append(f"Fórmula estrutural ausente em {chave}.")
    if possui_posicao_contratual:
        for chave in (
            "aditivos!L2",
            "posicao_contratual!E2",
            "posicao_contratual!I2",
            "posicao_contratual!M2",
            "posicao_contratual!Q2",
            "posicao_contratual!U2",
            "posicao_contratual!X2",
            "itens_Remanesc!F2",
            "itens_RC!C3",
            "historico_VU!N2",
        ):
            if chave not in formulas:
                bloqueios_estruturais.append(f"Fórmula estrutural ausente em {chave}.")
    referencias_quebradas = [chave for chave, formula in formulas.items() if "#REF!" in formula.upper()]
    if referencias_quebradas:
        bloqueios_estruturais.append("Há fórmulas com referência quebrada: " + ", ".join(referencias_quebradas[:5]))

    nomes_definidos = set(wb.defined_names)
    nomes_ausentes = sorted(NOMES_RESULTADOS_OBRIGATORIOS - nomes_definidos)
    if nomes_ausentes:
        bloqueios_estruturais.append("Nomes estruturais da aba RESULTADOS ausentes: " + ", ".join(nomes_ausentes[:8]))
    if wb.sheetnames[-1] != "RESULTADOS":
        avisos.append("A aba RESULTADOS deve permanecer como a última aba do arquivo.")
    abas_coloridas = [ws.title for ws in wb.worksheets if ws.sheet_properties.tabColor is not None]
    if abas_coloridas != ["RESULTADOS"]:
        avisos.append("Somente a guia RESULTADOS deve possuir cor de aba.")

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
        bloqueios_estruturais.append("Comentários/observações de célula não são admitidos: " + ", ".join(comentarios[:5]))
    if observacoes:
        bloqueios_estruturais.append("Campos de observação não são admitidos: " + ", ".join(observacoes[:5]))

    parametros = wb["parametros"]
    ativos = [numero for numero in range(1, 5) if str(parametros[f"A{numero + 2}"].value).strip().lower() == "sim"]
    if not ativos:
        lacunas_apuracao.append("Nenhum ciclo está marcado para computar nesta apuração.")
    else:
        for numero in range(1, max(ativos) + 1):
            if _numero(parametros[f"E{numero + 2}"].value) is None:
                lacunas_apuracao.append(f"C{numero}: percentual necessário ao acumulado está ausente.")

    financeiro = wb["financeiro"]
    divergencias_manuais: list[str] = []
    comparar_ajustes_manuais = _tem_metadado_inicio_efeito(wb)
    for row in range(2, 74):
        competencia = financeiro[f"A{row}"].value
        valor = _numero(financeiro[f"C{row}"].value)
        efeito = str(financeiro[f"G{row}"].value or "")
        ciclo = _ciclo_por_competencia_financeira(wb, competencia)
        data_comp = _data(competencia)
        referencia = (
            f"ciclo {ciclo or 'nao identificado'}, competencia "
            f"{data_comp.strftime('%m/%Y') if data_comp else 'nao informada'}"
        )
        if valor is not None and competencia not in (None, "") and data_comp is None:
            bloqueios_criticos.append(
                f"Competencia invalida na aba financeiro: linha {row}. "
                "Informe uma data mensal valida."
            )
            continue
        if efeito not in ("", "Sim", "Nao"):
            bloqueios_criticos.append(
                f"Efeito financeiro invalido na aba financeiro: {referencia}. "
                "Use o dropdown e selecione exatamente Sim ou Nao."
            )
            continue
        if valor is not None and not efeito:
            bloqueios_criticos.append(
                f"Efeito financeiro nao informado na aba financeiro: {referencia}."
            )
            continue
        if (
            not comparar_ajustes_manuais
            or not efeito
            or not ciclo
            or data_comp is None
        ):
            continue
        esperado = "Nao"
        inicio_efeito = _inicio_efeito_definido(wb, ciclo)
        ciclo_ativo = ciclo in {f"C{numero}" for numero in ativos}
        if ciclo_ativo and inicio_efeito:
            comp = data_comp.date()
            esperado = (
                "Sim"
                if (comp.year, comp.month) >= (inicio_efeito.year, inicio_efeito.month)
                else "Nao"
            )
        if efeito != esperado:
            divergencias_manuais.append(
                "Marcacao de efeito financeiro ajustada manualmente: "
                f"{ciclo} - {data_comp.strftime('%m/%Y')}."
            )
    avisos.extend(divergencias_manuais[:12])
    if len(divergencias_manuais) > 12:
        avisos.append(
            f"Ha mais {len(divergencias_manuais) - 12} marcacoes manuais de efeito financeiro."
        )

    contagens = {
        "competencias_com_valor": sum(1 for row in range(2, 74) if _numero(wb["financeiro"][f"C{row}"].value) is not None),
        "itens_remanescentes": sum(1 for row in range(2, 201) if wb["itens_Remanesc"][f"A{row}"].value not in (None, "")),
        "itens_consumidos": sum(1 for row in range(2, 201) if wb["itens_Consumidos"][f"A{row}"].value not in (None, "")),
        "pedidos_de_compra": sum(1 for row in range(2, 101) if wb["itens_PC"][f"A{row}"].value not in (None, "")),
        "aditivos": sum(1 for row in range(2, 201) if wb["aditivos"][f"A{row}"].value not in (None, "")),
        "formulas": len(formulas),
        "posicao_contratual_itens": 0,
        "posicao_contratual_calculada": 0,
        "historico_vu_itens": 0,
        "historico_vu_calculado": 0,
    }
    if contagens["competencias_com_valor"] == 0 and contagens["itens_remanescentes"] == 0:
        avisos.append("Ainda não há valores mensais nem itens remanescentes preenchidos pelo fiscal.")

    status_resultados: dict[str, Any] = {}
    try:
        wb_valores = load_workbook(BytesIO(conteudo), data_only=True, read_only=True)
        erros = []
        for ws in wb_valores.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.upper() in ERROS_EXCEL:
                        erros.append(f"{ws.title}!{cell.coordinate}={cell.value}")
        if erros:
            lacunas_apuracao.append("O Excel salvou erros de cálculo: " + ", ".join(erros[:8]))
        if possui_posicao_contratual:
            contagens["posicao_contratual_itens"] = sum(
                1 for row in range(2, 201)
                if wb_valores["posicao_contratual"][f"A{row}"].value not in (None, "")
            )
            contagens["posicao_contratual_calculada"] = sum(
                1 for row in range(2, 201)
                if wb_valores["posicao_contratual"][f"A{row}"].value not in (None, "")
                and any(
                    _numero(wb_valores["posicao_contratual"][f"{col}{row}"].value) is not None
                    for col in ("E", "G", "I", "K", "M", "O", "Q", "S", "U", "W")
                )
            )
        contagens["historico_vu_itens"] = sum(
            1 for row in range(2, 201)
            if wb_valores["historico_VU"][f"A{row}"].value not in (None, "")
        )
        contagens["historico_vu_calculado"] = sum(
            1 for row in range(2, 201)
            if wb_valores["historico_VU"][f"A{row}"].value not in (None, "")
            and any(
                _numero(wb_valores["historico_VU"][f"{col}{row}"].value) is not None
                for col in ("B", "D", "F", "H", "J", "L", "N", "P")
            )
        )
        if possui_posicao_contratual:
            alertas_aditivos = [
                f"aditivos!M{row}={wb_valores['aditivos'][f'M{row}'].value}"
                for row in range(2, 201)
                if str(wb_valores["aditivos"][f"M{row}"].value or "").startswith("ALERTA:")
            ]
            alertas_posicao = [
                f"posicao_contratual!X{row}={wb_valores['posicao_contratual'][f'X{row}'].value}"
                for row in range(2, 201)
                if str(wb_valores["posicao_contratual"][f"X{row}"].value or "").startswith("ALERTA:")
            ]
            if alertas_aditivos:
                bloqueios_criticos.append(
                    "Aditivos quantitativos inconsistentes: " + ", ".join(alertas_aditivos[:5])
                )
            if alertas_posicao:
                bloqueios_criticos.append(
                    "Posição contratual inconsistente: " + ", ".join(alertas_posicao[:5])
                )
        resultados_valores = wb_valores["RESULTADOS"]
        status_resultados = {
            "geral": resultados_valores["J4"].value,
            "metodo_retroativo": resultados_valores["B4"].value,
            "origem_retroativo_oficial": resultados_valores["D16"].value,
            "retroativo": resultados_valores["F16"].value,
            "vta": resultados_valores["E26"].value,
            "remanescente": resultados_valores["F36"].value,
            "valores": {
                "retroativo_financeiro": resultados_valores["B15"].value,
                "retroativo_pc": resultados_valores["C15"].value,
                "retroativo_itens": resultados_valores["D15"].value,
                "retroativo_oficial": resultados_valores["B16"].value,
                "vta_base_contratual": resultados_valores["B20"].value,
                "vta_retroativo": resultados_valores["B21"].value,
                "vta_ajuste_remanescente": resultados_valores["B22"].value,
                "vta_calculado": resultados_valores["B23"].value,
                "vta_ajuste_manual": resultados_valores["B24"].value,
                "vta_manual_oficial": resultados_valores["B25"].value,
                "vta_oficial": resultados_valores["B26"].value,
                "quantidade_remanescente": resultados_valores["B35"].value,
                "remanescente_original": resultados_valores["C35"].value,
                "remanescente_atualizado": resultados_valores["D35"].value,
            },
        }
        if not status_resultados["geral"]:
            avisos.append(
                "Os status de RESULTADOS não estão calculados em cache; abra, recalcule e salve o XLS no Excel."
            )
    except Exception:
        avisos.append("Não foi possível conferir os valores calculados em cache; abra e salve o arquivo no Excel.")

    metadados = {
        "indice": wb["CONTROLE"]["B7"].value,
        "ciclo_vigente": wb["CONTROLE"]["B2"].value,
        "data_corte": wb["CONTROLE"]["B3"].value,
        "ciclos_em_analise": [f"C{numero}" for numero in ativos],
        "status_resultados": status_resultados,
        "arquitetura_posicao_contratual": "canonica" if possui_posicao_contratual else "legada",
    }
    bloqueios = bloqueios_estruturais + bloqueios_criticos
    capacidades = avaliar_capacidades_apuracao(
        contagens,
        metadados,
        bloqueios,
        lacunas_apuracao,
    )
    possui_base = capacidades["resumo"]["tem_alguma_evidencia"]
    resultados_seguros = capacidades["resumo"]["apuracao_integral"]
    pendencias = bloqueios + lacunas_apuracao
    return {
        "valido": not bloqueios,
        "pronto_para_consolidar": not bloqueios and possui_base and resultados_seguros,
        "processamento_progressivo": True,
        "pendencias": pendencias,
        "bloqueios_estruturais": bloqueios_estruturais,
        "bloqueios_criticos": bloqueios_criticos,
        "lacunas_apuracao": lacunas_apuracao,
        "avisos": avisos,
        "contagens": contagens,
        "metadados": metadados,
        "capacidades": capacidades,
    }


def eh_coleta_reajuste(conteudo: bytes) -> bool:
    try:
        wb = load_workbook(BytesIO(conteudo), read_only=True, data_only=False)
        nomes = set(wb.sheetnames)
        nucleares = {"CONTROLE", "parametros", "financeiro"}
        # Também reconhece uma coleta canônica danificada, para que o validador
        # possa explicar a aba ausente em vez de desviá-la ao leitor legado.
        return nucleares.issubset(nomes) and len(nomes.intersection(ABAS_OBRIGATORIAS_LEGADO)) >= 5
    except Exception:
        return False
