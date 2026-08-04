"""Testes permanentes de integridade OOXML do template oficial da Coleta.

Protegem contra a regressao de corrupcao identificada na Etapa 3:
mc:Ignorable com prefixos nao declarados, marcador repairLoad, perda de
formulas/estilos e descaracterizacao da aba financeiro.
"""
from __future__ import annotations

import gc
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
    # comparativo_VTA: aba de referencia adicionada via Excel COM (1a posicao
    # fisica). cobertura_temporal 14->15 (+ linha Metodo de apuracao).
    # Etapa 26F: calculos antigos preservados em MEMORIA_RESULTADOS e nova
    # RESULTADOS executiva com formulas de apresentacao.
    "comparativo_VTA": 1407,
    "CONTROLE": 6,
    "parametros": 32,
    "financeiro": 291,
    # 26H.1/26H.2: +199 formulas pre-semeadas de base zero visual em
    # B2:B200. FRONTEIRA FUNCIONAL: linha 200 = ultima linha de cadastro
    # contratual (A2:A200 em toda a cadeia); linha 201 = linha extra do
    # total dinamico de D/F — FORA da capacidade funcional e sem automacao
    # de input (o off-by-one B201 da 26H.1 foi apontado pela auditoria e
    # removido na 26H.2).
    "itens_Remanesc": 9394,
    "itens_Consumidos": 1806,
    # Etapa 26G: grade escalada para a capacidade canonica (5.000 PCs
    # x 8 colunas de formula) + resumo lateral N2:T6.
    "itens_PC": 40042,
    "aditivos": 1393,
    "posicao_referencia": 2595,
    # 26H: +398 formulas das colunas tecnicas ocultas Y (CICLO_NASCIMENTO)
    # e Z (EH_NOVO_ITEM), linhas 2:200.
    "posicao_contratual": 5174,
    # Etapa VTA-posicoes: +1800 do bloco POSICAO ATUAL (AUTO) Q:Y (9 colunas
    # x 200 linhas 3:202) via INDIRECT+ISERROR sobre CICLO_EM_EXECUCAO.
    "itens_RC": 5000,
    "historico_VU": 3592,
    "cobertura_temporal": 15,
    # 3762 anteriores + 11 referencias para a tabela manual unica.
    # 26G: +5 (T26/T27 completude do remanescente; T28:T30 PCs sem efeito).
    # Etapa VTA-posicoes: +212 do bloco auxiliar das 3 referencias
    # (W41:W52 = 12 formulas + AB2:AB201 = 200 formulas). B26/T25 intactos.
    "MEMORIA_RESULTADOS": 3990,
    # 57 do prototipo + 4 selos por tabela + 1 premissa da estimativa - 1
    # helper J4 removido (status global agora agrega os selos H8/H14/H24/H33).
    # 26G: +5 (linha executiva A23:E23 dos PCs sem efeito financeiro).
    # Etapa VTA-posicoes: +7 liquidas na Tabela 1 (3 referencias +
    # reconciliacao: B10/C10/H10, B11/C11/H11, B12, B13/H13 = 9 novas,
    # menos as 2 antigas B10/B11 substituidas). H8 preservado.
    "RESULTADOS": 73,
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
    # 26G: grade ate a capacidade canonica — L101 tem formula.
    assert str(ws["L101"].value).startswith("=")
    validacoes = [
        (dv.type, str(dv.sqref)) for dv in ws.data_validations.dataValidation
    ]
    assert validacoes == [("list", "G2:G5001")]
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
            assert wb.Worksheets.Count == 15, f"rodada {rodada}"
            wb.Close(False)
            del wb
    finally:
        excel.Quit()
        del excel
        gc.collect()
        pythoncom.CoUninitialize()
