from __future__ import annotations

import hashlib
import io
import re
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from _coleta_oficial import (
    ABAS_COLETA_OFICIAL,
    NOME_ARQUIVO_COLETA_OFICIAL,
    TEMPLATE_COLETA_OFICIAL,
    gerar_coleta_oficial_preenchida,
)
from _coleta_reajuste_documentos import processar_coleta_oficial_runtime


ROOT = Path(__file__).resolve().parents[1]
# Atualizado na Etapa 4 do hotfix: restauracao da formula VLOOKUP ausente em
# aditivos!F5 via Excel COM (re-save nativo). Layout visual preservado; apenas
# os bytes do template mudaram, por isso o pin de SHA-256 foi reancorado.
# Reancorado novamente no ajuste final: remocao da opcao "Decrescimo" do
# dropdown de aditivos!D2:D200 (Acrescimo,Supressao) via Excel COM (re-save
# nativo, sem reparo). Layout/formulas F2:F200 preservados.
SHA256_TEMPLATE_ESPERADO = "cc197fac223001db65bfab29ff97033e3c848de9e34616d51c1b28af6c82ad2e"


def _dados_calculadora() -> dict:
    return {
        "origem": "Reajuste Simples",
        "indice": "IST",
        "data_base_original": "01/01/2023",
        "ciclos": [{
            "ciclo": "C1",
            "data_base": "01/01/2023",
            "data_pedido": "01/01/2024",
            "situacao": "TEMPESTIVO",
            "percentual_aplicado": 0.10,
            "financeiro_inicio": "01/01/2024",
        }],
    }


def _dia(valor):
    return valor.date() if hasattr(valor, "date") else valor


def test_calculadoras_e_upload_usam_o_mesmo_fluxo_oficial() -> None:
    simples = (ROOT / "pages/01_Calculo_Simples.py").read_text(encoding="utf-8")
    multiplos = (ROOT / "pages/02_Calculo_Represados.py").read_text(encoding="utf-8")
    upload = (ROOT / "pages/03_Valor_Global.py").read_text(encoding="utf-8")
    runtime = (ROOT / "_coleta_reajuste_documentos.py").read_text(encoding="utf-8")
    for pagina in (simples, multiplos):
        assert "gerar_coleta_oficial_preenchida" in pagina
        assert "gerar_coleta_reajuste(" not in pagina
        assert "NOME_ARQUIVO_COLETA_OFICIAL" in pagina
    assert "processar_coleta_oficial_runtime(conteudo_upload)" in upload
    assert "processar_arquivo_coleta(conteudo)" not in upload
    assert "ler_masterfile_v10(conteudo, exigir_modelo_oficial=True)" in runtime
    assert "reconciliacao_xls_python" in runtime
    assert "avaliar_entrega_segura" in runtime


def test_geracao_pos_calculadora_preserva_e_preenche_modelo_oficial() -> None:
    payload = gerar_coleta_oficial_preenchida(_dados_calculadora())
    wb = load_workbook(io.BytesIO(payload), data_only=False)
    assert NOME_ARQUIVO_COLETA_OFICIAL == "COLETA_REAJUSTE_OFICIAL.xlsx"
    assert wb.sheetnames == ABAS_COLETA_OFICIAL
    assert wb["CONTROLE"]["B2"].value == "C1"
    assert wb["CONTROLE"]["B7"].value == "IST"
    assert _dia(wb["CONTROLE"]["B8"].value) == date(2023, 1, 1)
    assert wb["parametros"]["B3"].value == "C1"
    assert _dia(wb["parametros"]["C3"].value) == date(2024, 1, 1)
    assert _dia(wb["financeiro"]["A2"].value) == date(2023, 1, 1)
    assert _dia(wb["financeiro"]["A25"].value) == date(2024, 12, 1)
    assert wb["itens_PC"]["A1"].value == "NUMERO_PC"
    assert wb["itens_PC"]["B1"].value == "DATA_PC"
    assert wb["itens_PC"]["C1"].value == "CICLO_PC"
    assert "ITEM" not in [wb["itens_PC"].cell(1, c).value for c in range(1, 12)]
    assert wb["RESULTADOS"]["A52"].value is not None


def test_multiciclo_iniciado_em_c2_nao_marca_c1_como_objeto_atual() -> None:
    dados = {
        "origem": "Reajustes Múltiplos",
        "indice": "IST",
        "data_base_original": "10/10/2022",
        "ciclos": [
            {
                "ciclo": "C2",
                "data_base": "10/10/2023",
                "data_pedido": "10/10/2024",
                "percentual_aplicado": 0.0488,
                "financeiro_inicio": "01/10/2024",
                "objeto_analise_atual": True,
            },
            {
                "ciclo": "C3",
                "data_base": "10/10/2024",
                "data_pedido": "10/10/2025",
                "percentual_aplicado": 0.0383,
                "financeiro_inicio": "01/10/2025",
                "objeto_analise_atual": True,
            },
        ],
    }
    wb = load_workbook(io.BytesIO(gerar_coleta_oficial_preenchida(dados)), data_only=False)
    parametros = wb["parametros"]
    assert parametros["B3"].value == "C1"
    assert parametros["A3"].value == "Nao"
    assert parametros["B4"].value == "C2"
    assert parametros["A4"].value == "Sim"
    assert parametros["B5"].value == "C3"
    assert parametros["A5"].value == "Sim"


def test_financeiro_comeca_em_c0_na_linha_2_caso_simples() -> None:
    # data-base do indice (2022-10) e 12 meses anterior ao inicio de C0
    # (2023-10): o periodo do indice pertence a memoria de calculo, nunca
    # a grade financeira. Primeira linha ativa = inicio de C0.
    dados = {
        "origem": "Reajuste Simples",
        "indice": "ICTI",
        "data_base_original": "01/10/2022",
        "ciclos": [{
            "ciclo": "C1",
            "data_base": "01/10/2023",
            "data_pedido": "01/10/2024",
            "percentual_aplicado": 0.0623,
            "financeiro_inicio": "01/10/2024",
        }],
    }
    wb = load_workbook(io.BytesIO(gerar_coleta_oficial_preenchida(dados)), data_only=False)
    financeiro = wb["financeiro"]
    inicio_c0 = _dia(wb["parametros"]["C2"].value)
    assert _dia(financeiro["A2"].value) == inicio_c0 == date(2023, 10, 1)
    # nenhuma competencia anterior ao inicio de C0 em toda a grade
    ativas = [r for r in range(2, 74) if financeiro[f"A{r}"].value is not None]
    assert ativas and ativas == list(range(2, ativas[-1] + 1))  # contiguas a partir da linha 2
    for r in ativas:
        assert _dia(financeiro[f"A{r}"].value) >= inicio_c0
    # linhas nao utilizadas permanecem vazias (A, C e G)
    for r in range(ativas[-1] + 1, 74):
        assert financeiro[f"A{r}"].value is None
        assert financeiro[f"C{r}"].value is None
        assert financeiro[f"G{r}"].value is None


def test_financeiro_comeca_em_c0_na_linha_2_multiciclo() -> None:
    dados = {
        "origem": "Reajustes Múltiplos",
        "indice": "ICTI",
        "data_base_original": "01/10/2022",
        "ciclos": [
            {
                "ciclo": "C1",
                "data_base": "01/10/2023",
                "data_pedido": "01/10/2024",
                "percentual_aplicado": 0.0623,
                "financeiro_inicio": "01/10/2023",
                "objeto_analise_atual": True,
            },
            {
                "ciclo": "C2",
                "data_base": "01/10/2024",
                "data_pedido": "01/10/2025",
                "percentual_aplicado": 0.0442,
                "financeiro_inicio": "01/10/2024",
                "objeto_analise_atual": True,
            },
        ],
    }
    wb = load_workbook(io.BytesIO(gerar_coleta_oficial_preenchida(dados)), data_only=False)
    financeiro = wb["financeiro"]
    inicio_c0 = _dia(wb["parametros"]["C2"].value)
    assert _dia(financeiro["A2"].value) == inicio_c0 == date(2023, 10, 1)
    ativas = [r for r in range(2, 74) if financeiro[f"A{r}"].value is not None]
    assert ativas == list(range(2, 38))  # 36 meses: C0 + C1 + C2 ate o corte
    for r in ativas:
        assert _dia(financeiro[f"A{r}"].value) >= inicio_c0
    assert _dia(financeiro[f"A{ativas[-1]}"].value) == date(2026, 9, 1)
    for r in range(38, 74):
        assert financeiro[f"A{r}"].value is None
        assert financeiro[f"G{r}"].value is None


def test_template_tem_72_competencias_e_resultados_alcanca_linha_73() -> None:
    wb = load_workbook(TEMPLATE_COLETA_OFICIAL, data_only=False)
    financeiro = wb["financeiro"]
    assert financeiro.max_row == 74  # linha 74 = TOTAL (B74=TOTAL, C74/E74/F74=SUM)
    for linha in range(2, 74):
        assert str(financeiro[f"B{linha}"].value).startswith("=")
        assert str(financeiro[f"D{linha}"].value).startswith("=")
        assert str(financeiro[f"E{linha}"].value).startswith("=")
        assert str(financeiro[f"F{linha}"].value).startswith("=")

    formulas_resultados = [
        cell.value
        for row in wb["RESULTADOS"].iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    ]
    assert any("financeiro!$B$2:$B$73" in formula for formula in formulas_resultados)
    assert not any(re.search(r"financeiro!\$([A-G])\$2:\$\1\$61", formula, re.I) for formula in formulas_resultados)


def test_template_preserva_layout_visual_e_sha256() -> None:
    wb = load_workbook(TEMPLATE_COLETA_OFICIAL, data_only=False)
    itens_pc = wb["itens_PC"]
    ocultas = [d for d in itens_pc.column_dimensions.values() if d.hidden]
    assert all(any(d.min <= coluna <= d.max for d in ocultas) for coluna in range(22, 30))
    assert itens_pc.sheet_view.topLeftCell in (None, "A1")
    assert wb["financeiro"].sheet_view.topLeftCell in (None, "A1")
    assert wb["RESULTADOS"]["B4"].fill.fgColor.rgb == "FFF7E7B2"
    assert wb["CONTROLE"]["B1"].fill.fgColor.rgb == "FFF7E7B2"
    assert hashlib.sha256(TEMPLATE_COLETA_OFICIAL.read_bytes()).hexdigest() == SHA256_TEMPLATE_ESPERADO


def test_upload_rejeita_modelo_antigo_sem_fallback() -> None:
    antigo = (ROOT / "templates/Coleta_Reajuste.xlsx").read_bytes()
    with pytest.raises(ValueError, match="versão anterior|NUMERO_PC|Template incompativel"):
        processar_coleta_oficial_runtime(antigo)
    pagina = (ROOT / "pages/03_Valor_Global.py").read_text(encoding="utf-8")
    assert "CAMINHO_MODELO_COLETA" not in pagina
    assert "Arquivo legado processado" not in pagina
    assert "download foi bloqueado para evitar o uso de modelo incompatível" in pagina


def test_interface_nao_reintroduz_rotulos_antigos() -> None:
    fontes = "\n".join(
        p.read_text(encoding="utf-8")
        for p in [ROOT / "app.py", *(ROOT / "pages").glob("*.py")]
    )
    assert "Piloto controlado" not in fontes
    assert "Piloto Controlado" not in fontes
    assert "Mesa GCC" not in fontes
    assert "MasterFile de entrada" not in fontes
