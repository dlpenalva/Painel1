"""Adapta a coleta canônica aos documentos sem recalcular seus resultados.

O Excel continua sendo a fonte de verdade. Este módulo lê exclusivamente os
valores em cache gravados pelo Excel e os organiza no contrato de dados já
consumido pelos relatórios da aplicação.
"""

from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import Any

import pandas as pd
from openpyxl import load_workbook

from _coleta_reajuste import ler_coleta_reajuste
from _objeto_processo import (
    Aditivos,
    Ciclos,
    Diagnostico,
    Financeiro,
    Identificacao,
    ObjetoProcesso,
    PosicaoContratual,
    Remanescentes,
    Retroativo,
    ValorTotalAtualizado,
)


def _numero(valor: Any, padrao: float = 0.0) -> float:
    if valor in (None, "") or isinstance(valor, bool):
        return padrao
    try:
        return float(valor)
    except (TypeError, ValueError):
        return padrao


def _data_br(valor: Any) -> str:
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    return str(valor or "")


def montar_objeto_processo(conteudo: bytes) -> ObjetoProcesso:
    """Monta o Objeto Processo a partir dos valores salvos em RESULTADOS.

    Único ponto do ClausGC que abre o workbook da Coleta. Depois daqui, a
    Interface e os documentos consomem exclusivamente o objeto retornado.
    """

    diagnostico = ler_coleta_reajuste(conteudo)
    if not diagnostico.get("valido"):
        raise ValueError("A coleta possui pendências estruturais e não pode liberar documentos.")
    capacidades = diagnostico.get("capacidades") or {}

    wb = load_workbook(BytesIO(conteudo), data_only=True, read_only=True)
    controle = wb["CONTROLE"]
    parametros = wb["parametros"]
    financeiro = wb["financeiro"]
    remanescentes = wb["itens_Remanesc"]
    aditivos = wb["aditivos"]
    posicao = wb["posicao_contratual"] if "posicao_contratual" in wb.sheetnames else None
    itens_rc = wb["itens_RC"]
    resultados = wb["RESULTADOS"]

    ciclos_rows = []
    fatores: dict[str, float] = {}
    for row in range(2, 7):
        ciclo = str(parametros[f"B{row}"].value or "").upper()
        if not ciclo:
            continue
        fator = _numero(parametros[f"F{row}"].value, 1.0)
        fatores[ciclo] = fator
        ciclos_rows.append(
            {
                "Ciclo": ciclo,
                "Data-base": _data_br(parametros[f"C{row}"].value),
                "Intervalo do índice": "",
                "Janela de admissibilidade": "",
                "Data do pedido": "",
                "Situação": parametros[f"G{row}"].value or "",
                "Tratamento financeiro do ciclo": "Apurar" if str(parametros[f"A{row}"].value).lower() == "sim" else "Fora da apuração",
                "Variação": _numero(parametros[f"E{row}"].value),
                "Fator": fator,
                "Fator acumulado": fator,
                "Fator acumulado efetivo": fator,
                "Fator ciclo efetivo": fator,
            }
        )
    df_ciclos = pd.DataFrame(ciclos_rows)

    financeiro_rows = []
    for row in range(2, 62):
        valor = financeiro[f"C{row}"].value
        if valor in (None, ""):
            continue
        financeiro_rows.append(
            {
                "Ciclo": str(financeiro[f"B{row}"].value or "").upper(),
                "Competência": financeiro[f"A{row}"].value,
                "Valor pago/faturado": _numero(valor),
                "Fator aplicado": _numero(financeiro[f"D{row}"].value, 1.0),
                "Valor atualizado": _numero(financeiro[f"E{row}"].value),
                "Delta": _numero(financeiro[f"F{row}"].value),
                "Efeito financeiro": financeiro[f"G{row}"].value or "",
            }
        )
    df_financeiro = pd.DataFrame(financeiro_rows)

    fin_ciclos = []
    for ciclo in [row["Ciclo"] for row in ciclos_rows]:
        linhas = df_financeiro[df_financeiro["Ciclo"] == ciclo] if not df_financeiro.empty else pd.DataFrame()
        pago = float(linhas["Valor pago/faturado"].sum()) if not linhas.empty else 0.0
        devido = float(linhas["Valor atualizado"].sum()) if not linhas.empty else 0.0
        delta = float(linhas["Delta"].sum()) if not linhas.empty else 0.0
        fin_ciclos.append(
            {
                "Ciclo": ciclo,
                "Situação": next((r["Situação"] for r in ciclos_rows if r["Ciclo"] == ciclo), ""),
                "Tratamento financeiro": "Apurar",
                "Fator aplicado ao retroativo": fatores.get(ciclo, 1.0),
                "Fator acumulado": fatores.get(ciclo, 1.0),
                "Valor pago efetivo": pago,
                "Valor teórico calculado": devido,
                "Delta do ciclo": delta,
            }
        )
    df_fin_por_ciclo = pd.DataFrame(fin_ciclos)

    valores_rows = []
    qtd_cols = {"C0": "B", "C1": "E", "C2": "G", "C3": "I", "C4": "K"}
    qtd_posicao_cols = {"C0": "G", "C1": "K", "C2": "O", "C3": "S", "C4": "W"}
    total_cols = {"C0": "D", "C1": "F", "C2": "H", "C3": "J", "C4": "L"}
    total_rc_cols = {"C0": "D", "C1": "G", "C2": "J", "C3": "M", "C4": "P"}
    for row in range(2, 201):
        item = remanescentes[f"A{row}"].value
        if item in (None, ""):
            continue
        for ciclo, qtd_col in qtd_cols.items():
            qtd = _numero(
                posicao[f"{qtd_posicao_cols[ciclo]}{row}"].value
                if posicao is not None
                else remanescentes[f"{qtd_col}{row}"].value
            )
            total = _numero(
                itens_rc[f"{total_rc_cols[ciclo]}{row + 1}"].value
                if posicao is not None
                else remanescentes[f"{total_cols[ciclo]}{row}"].value
            )
            valores_rows.append(
                {
                    "Item": item,
                    "Ciclo": ciclo,
                    "Valor unitário": total / qtd if qtd else _numero(remanescentes[f"C{row}"].value) * fatores.get(ciclo, 1.0),
                    "Quantidade": qtd,
                    "Total R$": total,
                    "Ciclo precluso": False,
                }
            )
    df_valores = pd.DataFrame(valores_rows)

    aditivos_rows = []
    for row in range(2, 201):
        item = aditivos[f"A{row}"].value
        if item in (None, ""):
            continue
        aditivos_rows.append(
            {
                "Item": item,
                "Data do aditivo": aditivos[f"B{row}"].value,
                "Ciclo/Marco": aditivos[f"C{row}"].value or "",
                "Tipo de alteração": aditivos[f"D{row}"].value or "",
                "Quantidade da alteração": _numero(aditivos[f"E{row}"].value),
                "Delta quantitativo contratual": _numero(aditivos[f"L{row}"].value),
                "Status da posição": aditivos[f"M{row}"].value or "",
                "Valor do aditivo na assinatura": _numero(aditivos[f"G{row}"].value),
                "Fator aplicado": _numero(aditivos[f"I{row}"].value, 1.0),
                "Valor do aditivo reajustado": _numero(aditivos[f"J{row}"].value),
                "Computa no Valor Global": str(aditivos[f"K{row}"].value or "").lower() == "sim",
            }
        )
    df_aditivos = pd.DataFrame(aditivos_rows)

    posicao_rows = []
    if posicao is not None:
        contratada_cols = {"C0": "E", "C1": "I", "C2": "M", "C3": "Q", "C4": "U"}
        rem_cols = {"C0": "G", "C1": "K", "C2": "O", "C3": "S", "C4": "W"}
        for row in range(2, 201):
            item = posicao[f"A{row}"].value
            if item in (None, ""):
                continue
            for ciclo in ("C0", "C1", "C2", "C3", "C4"):
                posicao_rows.append(
                    {
                        "Item": item,
                        "Ciclo": ciclo,
                        "Quantidade contratada vigente": _numero(posicao[f"{contratada_cols[ciclo]}{row}"].value),
                        "Quantidade remanescente ajustada": _numero(posicao[f"{rem_cols[ciclo]}{row}"].value),
                        "Status": posicao[f"X{row}"].value or "",
                    }
                )
    df_posicao_contratual = pd.DataFrame(posicao_rows)

    calculos = capacidades.get("calculos") or {}
    retro_capacidade = calculos.get("retroativo") or {}
    vta_capacidade = calculos.get("vta") or {}
    rem_capacidade = calculos.get("valor_remanescente") or {}

    valor_original = _numero(resultados["B20"].value)
    retroativo = _numero(retro_capacidade.get("valor"))
    rem_original = _numero(resultados["C35"].value)
    rem_atualizado = _numero(rem_capacidade.get("valor"))
    valor_total = _numero(vta_capacidade.get("valor"))
    pago_total = float(df_financeiro["Valor pago/faturado"].sum()) if not df_financeiro.empty else 0.0
    devido_total = float(df_financeiro["Valor atualizado"].sum()) if not df_financeiro.empty else 0.0
    total_aditivos = float(df_aditivos["Valor do aditivo reajustado"].sum()) if not df_aditivos.empty else 0.0
    fator_final = max(fatores.values(), default=1.0)
    execucao_atualizada = valor_total - rem_atualizado

    df_rem = pd.DataFrame(
        [{
            "Ciclo": controle["B2"].value or "",
            "Remanescente original": rem_original,
            "Fator aplicado": (rem_atualizado / rem_original) if rem_original else fator_final,
            "Remanescente atualizado": rem_atualizado,
            "Observação": "Valores oficiais lidos da aba RESULTADOS.",
        }]
    )
    df_execucao = pd.DataFrame(
        [{
            "Ciclo": controle["B2"].value or "",
            "Valor executado original": max(valor_original - rem_original, 0.0),
            "Valor executado atualizado": execucao_atualizada,
        }]
    )
    df_composicao = pd.DataFrame(
        [
            {"Componente": "Execução atualizada e retroativo", "Valor": execucao_atualizada},
            {"Componente": "Saldo remanescente atualizado", "Valor": rem_atualizado},
            {"Componente": "Valor Total Atualizado do Contrato", "Valor": valor_total},
        ]
    )
    df_comparativo = pd.DataFrame(
        [
            {"Indicador": "Valor original do contrato", "Valor": valor_original},
            {"Indicador": "Valor pago efetivo", "Valor": pago_total},
            {"Indicador": "Valor teórico calculado", "Valor": devido_total},
            {"Indicador": "Valor represado a pagar", "Valor": retroativo},
            {"Indicador": "Saldo remanescente original", "Valor": rem_original},
            {"Indicador": "Saldo remanescente atualizado", "Valor": rem_atualizado},
            {"Indicador": "Valor Total Atualizado do Contrato", "Valor": valor_total},
        ]
    )

    return ObjetoProcesso(
        identificacao=Identificacao(
            origem_coleta="Coleta_Reajuste.xlsx",
            data_processamento=datetime.now().strftime("%d/%m/%Y %H:%M"),
            modo_apuracao="Processamento progressivo pelo XLS",
            indice=controle["B7"].value or "Não informado",
            ciclo_ultimo_remanescente=controle["B2"].value or "",
        ),
        ciclos=Ciclos(
            tabela=df_ciclos,
            quantidade=len(ciclos_rows),
            fator_acumulado=fator_final,
            variacao_acumulada=fator_final - 1.0,
            origem="Coleta_Reajuste.xlsx",
        ),
        financeiro=Financeiro(
            mensal=df_financeiro,
            por_ciclo=df_fin_por_ciclo,
            valor_pago_efetivo=pago_total,
            valor_teorico_calculado=devido_total,
            disponivel=bool(financeiro_rows),
            aviso="" if financeiro_rows else "Financeiro não informado; os demais blocos permanecem utilizáveis.",
            meses_sem_efeito=pd.DataFrame(),
            quantidade_meses_sem_efeito=0,
            valor_total_sem_efeito=0.0,
        ),
        retroativo=Retroativo(
            valor=retroativo,
            disponivel=bool(abs(retroativo) > 0.004),
            capacidade=retro_capacidade,
            estimado_itens_estoque=pd.DataFrame(),
        ),
        vta=ValorTotalAtualizado(
            valor_original_contrato=valor_original,
            valor_total=valor_total,
            execucao_atualizada=execucao_atualizada,
            composicao=df_composicao,
            capacidade=vta_capacidade,
        ),
        remanescentes=Remanescentes(
            original=rem_original,
            reajustado=rem_atualizado,
            fator=(rem_atualizado / rem_original) if rem_original else fator_final,
            tabela=df_rem,
            valores_unitarios=df_valores,
            execucao=df_execucao,
            capacidade=rem_capacidade,
            capacidade_valores_unitarios=calculos.get("valores_unitarios") or {},
            base_itens_disponivel=bool(valores_rows),
        ),
        aditivos=Aditivos(
            tabela=df_aditivos,
            computaveis=(
                df_aditivos[df_aditivos.get("Computa no Valor Global", pd.Series(dtype=bool)) == True]
                if not df_aditivos.empty
                else pd.DataFrame()
            ),
            total_atualizados=total_aditivos,
            quantidade_total=len(aditivos_rows),
            quantidade_computaveis=sum(bool(r["Computa no Valor Global"]) for r in aditivos_rows),
        ),
        posicao_contratual=PosicaoContratual(
            tabela=df_posicao_contratual,
            capacidade=calculos.get("posicao_contratual") or {},
        ),
        diagnostico=Diagnostico(
            capacidades=capacidades,
            status_resultados=diagnostico.get("metadados", {}).get("status_resultados", {}),
            coleta=diagnostico,
            comparativo=df_comparativo,
            auditoria_consistencia=pd.DataFrame(
                [
                    {
                        "Validação": "Blocos independentes do XLS avaliados",
                        "Status": "OK",
                        "Diferença/Valor": "Processamento progressivo",
                    }
                ]
            ),
            ressalva_modo_apuracao="Somente resultados sustentados pelos blocos disponíveis são apresentados.",
        ),
    )


def adaptar_coleta_reajuste_para_documentos(conteudo: bytes) -> dict[str, Any]:
    """Expõe o Objeto Processo no contrato de dados histórico.

    Camada de compatibilidade: preserva a assinatura e o dicionário já
    consumidos pelos documentos e pela Interface. Novos consumidores devem
    preferir `montar_objeto_processo`.
    """

    return montar_objeto_processo(conteudo).como_dicionario()
