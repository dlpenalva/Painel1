from __future__ import annotations

import io
import os
import re
import zipfile
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from _coleta_oficial import gerar_coleta_oficial_preenchida
from _coleta_reajuste_documentos import processar_coleta_oficial_runtime


pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="defina RUN_EXCEL_INTEGRATION=1 para executar o Excel COM",
)


def _dados(*, longo: bool = False) -> dict:
    ciclos = [
        {"ciclo": "C1", "data_inicio": date(2024, 1, 1), "data_fim": date(2024, 12, 31), "data_pedido": date(2024, 1, 1), "percentual": 0.10},
    ]
    if longo:
        ciclos.extend([
            {"ciclo": "C2", "data_inicio": date(2025, 1, 1), "data_fim": date(2025, 12, 31), "data_pedido": date(2025, 1, 1), "percentual": 0.10},
            {"ciclo": "C3", "data_inicio": date(2026, 1, 1), "data_fim": date(2026, 12, 31), "data_pedido": date(2026, 1, 1), "percentual": 0.10},
            {"ciclo": "C4", "data_inicio": date(2027, 1, 1), "data_fim": date(2028, 12, 31), "data_pedido": date(2027, 1, 1), "percentual": 0.10},
        ])
    return {
        "origem": "Teste Excel COM",
        "indice": "IST",
        "data_base_original": "01/01/2023",
        "data_corte": date(2028, 12, 31) if longo else date(2024, 12, 31),
        "ciclos": ciclos,
    }


def _salvar(wb, caminho: Path) -> None:
    buf = io.BytesIO()
    wb.save(buf)
    caminho.write_bytes(buf.getvalue())


def _recalcular_excel(caminho: Path) -> None:
    import gc
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    pasta = None
    try:
        pasta = excel.Workbooks.Open(str(caminho.resolve()))
        excel.CalculateFullRebuild()
        pasta.Save()
        pasta.Close(SaveChanges=False)
        pasta = None
    finally:
        if pasta is not None:
            pasta.Close(SaveChanges=False)
            pasta = None
        excel.Quit()
        excel = None
        gc.collect()
        pythoncom.CoUninitialize()


def _alterar_cache_resultados(caminho: Path, celula: str, valor: float) -> None:
    temporario = caminho.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(caminho, "r") as origem, zipfile.ZipFile(
        temporario, "w", zipfile.ZIP_DEFLATED
    ) as destino:
        for item in origem.infolist():
            dados = origem.read(item.filename)
            if item.filename == "xl/worksheets/sheet11.xml":
                texto = dados.decode("utf-8")
                padrao = re.compile(
                    rf'(<c r="{re.escape(celula)}"[^>]*>.*?<v>)([^<]*)(</v>)',
                    re.DOTALL,
                )
                texto, quantidade = padrao.subn(rf"\g<1>{valor}\g<3>", texto, count=1)
                assert quantidade == 1
                dados = texto.encode("utf-8")
            destino.writestr(item, dados)
    temporario.replace(caminho)


def test_excel_com_linha_73_entra_em_resultados(tmp_path: Path) -> None:
    wb = load_workbook(io.BytesIO(gerar_coleta_oficial_preenchida(_dados(longo=True))))
    wb["financeiro"]["C73"] = 100.0
    wb["financeiro"]["G73"] = "Sim"
    caminho = tmp_path / "linha_73.xlsx"
    _salvar(wb, caminho)
    _recalcular_excel(caminho)

    valores = load_workbook(caminho, data_only=True)
    assert str(valores["financeiro"]["B73"].value).lower() == "c4"
    assert valores["financeiro"]["F73"].value == pytest.approx(46.41, abs=0.01)
    assert valores["RESULTADOS"]["B14"].value == pytest.approx(46.41, abs=0.01)


def test_excel_com_runtime_financeiro_pc_reconciliacao_e_bloqueio(tmp_path: Path) -> None:
    base = gerar_coleta_oficial_preenchida(_dados())

    financeiro = load_workbook(io.BytesIO(base))
    financeiro["financeiro"]["C14"] = 600.0
    financeiro["financeiro"]["G14"] = "Sim"
    financeiro["itens_Remanesc"]["A2"] = "ITEM-1"
    financeiro["itens_Remanesc"]["B2"] = 10.0
    financeiro["itens_Remanesc"]["C2"] = 100.0
    financeiro["itens_Remanesc"]["E2"] = 4.0
    caminho_fin = tmp_path / "financeiro.xlsx"
    _salvar(financeiro, caminho_fin)
    _recalcular_excel(caminho_fin)
    resultado_fin, _ = processar_coleta_oficial_runtime(caminho_fin.read_bytes())
    assert resultado_fin["valor_represado_a_pagar"] == pytest.approx(60.0, abs=0.01)
    assert resultado_fin["remanescente_reajustado"] == pytest.approx(440.0, abs=0.01)
    assert resultado_fin["valor_atualizado_contrato"] == pytest.approx(1100.0, abs=0.01)
    assert resultado_fin["reconciliacao_xls_python"]["status_geral"] == "CONCILIADO"
    assert not resultado_fin["formalizacao_bloqueada"]

    pc = load_workbook(io.BytesIO(base))
    pc["CONTROLE"]["B1"] = "Pedidos de Compras"
    pc["RESULTADOS"]["B4"] = "PCs"
    pc["itens_PC"]["A2"] = "PC-001"
    pc["itens_PC"]["B2"] = date(2024, 1, 15)
    pc["itens_PC"]["D2"] = 600.0
    pc["itens_PC"]["G2"] = "Sim"
    pc["itens_Remanesc"]["A2"] = "ITEM-1"
    pc["itens_Remanesc"]["B2"] = 10.0
    pc["itens_Remanesc"]["C2"] = 100.0
    pc["itens_Remanesc"]["E2"] = 4.0
    caminho_pc = tmp_path / "pc.xlsx"
    _salvar(pc, caminho_pc)
    _recalcular_excel(caminho_pc)
    resultado_pc, _ = processar_coleta_oficial_runtime(caminho_pc.read_bytes())
    assert resultado_pc["df_pedidos_compra"].iloc[0]["numero_pc"] == "PC-001"
    assert resultado_pc["valor_represado_a_pagar"] == pytest.approx(60.0, abs=0.01)
    assert resultado_pc["reconciliacao_xls_python"]["status_geral"] == "CONCILIADO"
    assert not resultado_pc["formalizacao_bloqueada"]

    _alterar_cache_resultados(caminho_pc, "C15", 9999.0)
    divergente, _ = processar_coleta_oficial_runtime(caminho_pc.read_bytes())
    assert divergente["reconciliacao_xls_python"]["divergencias_relevantes"]
    assert divergente["formalizacao_bloqueada"]
    assert not any(
        doc.get("habilitado")
        for doc in (divergente.get("capacidades") or {}).get("documentos", {}).values()
    )


def test_excel_com_pcs_multiciclo_ignora_fator_historico_fora_do_objeto(tmp_path: Path) -> None:
    dados = {
        "origem": "Teste PC multiciclo",
        "indice": "IST",
        "data_base_original": "01/01/2023",
        "data_corte": date(2026, 12, 31),
        "ciclos": [
            {
                "ciclo": "C1", "data_inicio": date(2024, 1, 1),
                "data_fim": date(2024, 12, 31), "percentual": 0.05,
                "ciclo_ja_concedido": True,
                "situacao": "Histórico fora do objeto atual",
            },
            {
                "ciclo": "C2", "data_inicio": date(2025, 1, 1),
                "data_fim": date(2025, 12, 31), "percentual": 0.10,
            },
            {
                "ciclo": "C3", "data_inicio": date(2026, 1, 1),
                "data_fim": date(2026, 12, 31), "percentual": 0.08,
            },
        ],
    }
    wb = load_workbook(io.BytesIO(gerar_coleta_oficial_preenchida(dados)))
    wb["CONTROLE"]["B1"] = "Pedidos de Compras"
    wb["RESULTADOS"]["B4"] = "PCs"
    for row, numero, data_pc, valor in (
        (2, "PC-2001", date(2025, 2, 15), 600.0),
        (3, "PC-3001", date(2026, 3, 20), 800.0),
    ):
        wb["itens_PC"][f"A{row}"] = numero
        wb["itens_PC"][f"B{row}"] = data_pc
        wb["itens_PC"][f"D{row}"] = valor
        wb["itens_PC"][f"G{row}"] = "Sim"
    caminho = tmp_path / "pc_multiciclo_historico.xlsx"
    _salvar(wb, caminho)
    _recalcular_excel(caminho)

    resultado, _ = processar_coleta_oficial_runtime(caminho.read_bytes())
    assert resultado["valor_represado_a_pagar"] == pytest.approx(210.40, abs=0.01)
    assert resultado["reconciliacao_xls_python"]["status_geral"] == "CONCILIADO"
    assert not resultado["reconciliacao_xls_python"]["divergencias_relevantes"]
    assert not any(
        "Divergência relevante XLS × Python" in bloqueio
        for bloqueio in resultado.get("bloqueios_formalizacao", [])
    )
