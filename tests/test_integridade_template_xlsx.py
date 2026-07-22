"""Testes permanentes de integridade OOXML do template oficial da Coleta.

Protegem contra a regressao de corrupcao identificada na Etapa 3:
mc:Ignorable com prefixos nao declarados, marcador repairLoad, perda de
formulas/estilos e descaracterizacao da aba financeiro.
"""
from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

FORMULAS_POR_ABA = {
    "CONTROLE": 6,
    "parametros": 32,
    "financeiro": 291,
    "itens_Remanesc": 8200,
    "itens_Consumidos": 1806,
    "itens_PC": 834,
    "aditivos": 1393,
    "posicao_contratual": 4776,
    "itens_RC": 3200,
    "historico_VU": 3592,
    "RESULTADOS": 3319,
}


def _partes_xml(z: zipfile.ZipFile) -> list[str]:
    return [n for n in z.namelist() if n.endswith((".xml", ".rels"))]


def _abas_e_partes(z: zipfile.ZipFile) -> list[tuple[str, str]]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid2t = {
        rel.get("Id"): rel.get("Target")
        for rel in rels.findall("rel:Relationship", NS)
    }
    resultado = []
    for aba in wb.findall("m:sheets/m:sheet", NS):
        alvo = rid2t[aba.get("{%s}id" % NS["r"])].lstrip("/")
        if not alvo.startswith("xl/"):
            alvo = "xl/" + alvo
        resultado.append((aba.get("name"), alvo))
    return resultado


def test_template_existe():
    assert TEMPLATE.is_file()


def test_xml_bem_formado_em_todas_as_partes():
    with zipfile.ZipFile(TEMPLATE) as z:
        for nome in _partes_xml(z):
            ET.fromstring(z.read(nome))


def test_mc_ignorable_somente_com_prefixos_declarados():
    padrao = re.compile(rb'mc:Ignorable="([^"]*)"')
    with zipfile.ZipFile(TEMPLATE) as z:
        for nome in _partes_xml(z):
            dados = z.read(nome)
            encontrado = padrao.search(dados)
            if not encontrado:
                continue
            for prefixo in encontrado.group(1).decode().split():
                declaracao = f'xmlns:{prefixo}='.encode()
                assert declaracao in dados, (
                    f"{nome}: mc:Ignorable referencia prefixo nao "
                    f"declarado {prefixo!r}"
                )


def test_sem_marcador_repairload():
    with zipfile.ZipFile(TEMPLATE) as z:
        contaminadas = [n for n in z.namelist() if b"repairLoad" in z.read(n)]
    assert contaminadas == []


def test_sem_vinculos_externos():
    with zipfile.ZipFile(TEMPLATE) as z:
        assert [n for n in z.namelist() if "externalLink" in n] == []
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        externos = [
            rel.get("Target")
            for rel in rels.findall("rel:Relationship", NS)
            if rel.get("TargetMode") == "External"
        ]
        assert externos == []


def test_contagem_de_formulas_por_aba():
    with zipfile.ZipFile(TEMPLATE) as z:
        abas = _abas_e_partes(z)
        assert [nome for nome, _ in abas] == list(FORMULAS_POR_ABA)
        for nome, parte in abas:
            raiz = ET.fromstring(z.read(parte))
            quantidade = len(raiz.findall(".//m:f", NS))
            assert quantidade == FORMULAS_POR_ABA[nome], (
                f"{nome}: esperava {FORMULAS_POR_ABA[nome]} formulas, "
                f"encontrei {quantidade}"
            )


def test_limites_minimos_de_estilos():
    with zipfile.ZipFile(TEMPLATE) as z:
        estilos = ET.fromstring(z.read("xl/styles.xml"))

    def contar(tag: str) -> int:
        elemento = estilos.find("m:" + tag, NS)
        return len(elemento) if elemento is not None else 0

    assert contar("cellXfs") >= 200
    assert contar("numFmts") >= 10
    assert contar("dxfs") >= 17


def test_financeiro_preservada():
    wb = load_workbook(TEMPLATE)
    ws = wb["financeiro"]
    formulas = sum(
        1
        for linha in ws.iter_rows()
        for celula in linha
        if isinstance(celula.value, str) and celula.value.startswith("=")
    )
    assert formulas == 291  # 72 fórmulas B + 216 DEF linhas 2-73 + 3 SUM em C74/E74/F74
    validacoes = [
        (dv.type, str(dv.sqref)) for dv in ws.data_validations.dataValidation
    ]
    assert validacoes == [("list", "G2:G73")]
    condicionais = sorted(str(rng.sqref) for rng in ws.conditional_formatting)
    assert condicionais == ["A2:G73"]


def test_itens_pc_efeito_financeiro_aplicado():
    wb = load_workbook(TEMPLATE)
    ws = wb["itens_PC"]
    assert ws["L1"].value == "EFEITO_FINANCEIRO_PC"
    assert isinstance(ws["L2"].value, str) and ws["L2"].value.startswith("=IF(")
    assert ws["L101"].value is None
    validacoes = [
        (dv.type, str(dv.sqref)) for dv in ws.data_validations.dataValidation
    ]
    assert validacoes == [("list", "G2:G100")]
    par = wb["parametros"]
    assert par["H1"].value == "INICIO_EFEITO_FINANCEIRO"
    assert {par.cell(r, 8).number_format for r in range(2, 7)} == {
        "dd/mm/yyyy;@"
    }


def test_aditivos_dropdown_tipo_alteracao_sem_decrescimo():
    """Ajuste final: aditivos!D2:D200 lista apenas Acrescimo/Supressao.

    O dropdown de TIPO DE ALTERACAO FORMALIZADA nao pode mais oferecer
    "Decrescimo"; deve conter exclusivamente Acrescimo e Supressao, cobrindo
    todo o intervalo D2:D200.
    """
    wb = load_workbook(TEMPLATE)
    ws = wb["aditivos"]
    dvs_d = [
        dv for dv in ws.data_validations.dataValidation
        if dv.type == "list" and "D2:D200" in str(dv.sqref)
    ]
    assert len(dvs_d) == 1, "esperada uma validacao de lista cobrindo D2:D200"
    dv = dvs_d[0]
    assert str(dv.sqref) == "D2:D200"
    itens = [t.strip() for t in dv.formula1.strip('"').split(",")]
    assert itens == ["Acrescimo", "Supressao"]
    assert not any("Decr" in i for i in itens)


def test_abertura_e_reabertura_sem_reparo_no_excel_real():
    client = pytest.importorskip("win32com.client")
    pythoncom = pytest.importorskip("pythoncom")
    pythoncom.CoInitialize()
    excel = client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = True
    try:
        for rodada in range(2):
            wb = excel.Workbooks.Open(str(TEMPLATE), UpdateLinks=0, ReadOnly=True)
            assert wb.Worksheets.Count == 11, f"rodada {rodada}"
            wb.Close(False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()
