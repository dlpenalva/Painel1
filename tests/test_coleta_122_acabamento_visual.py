"""Coleta 12.2 — acabamento visual de itens_PC, aditivos, CONTROLE e
CICLO_EM_EXECUCAO.

Somente apresentacao. A prova de que nada estrutural mudou e uma ASSINATURA
por aba (sem estilos): valores e formulas por coordenada, mesclas, validacoes
(sem os textos da mensagem de entrada), celulas editaveis, colunas/linhas
ocultas, filtro, congelamento e protecao. As assinaturas esperadas foram
calculadas, com esta mesma funcao, na Coleta 12.1 (main 2ae72bd): template
oficial e aba CICLO_EM_EXECUCAO gerada em runtime.
"""
from __future__ import annotations

import hashlib
import io
import os
import sys
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "tools"))

import aplicar_coleta_122_acabamento_visual as av  # noqa: E402
from _coleta_oficial import COLUNAS_ITENS_PC_OFICIAL  # noqa: E402

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
ABA_CICLO = "CICLO_EM_EXECUCAO"

_CACHE: dict[str, object] = {}


def _template():
    if "template" not in _CACHE:
        _CACHE["template"] = load_workbook(TEMPLATE)
    return _CACHE["template"]


def _gerado():
    if "gerado" not in _CACHE:
        from _coleta_oficial import obter_coleta_oficial_bytes

        _CACHE["gerado"] = load_workbook(io.BytesIO(obter_coleta_oficial_bytes()))
    return _CACHE["gerado"]


def _valor(valor) -> str:
    texto = getattr(valor, "text", None)       # ArrayFormula
    if texto is not None:
        return f"ARRAY[{getattr(valor, 'ref', '')}]{texto}"
    return repr(valor)


def assinatura(ws) -> str:
    """Hash da estrutura da aba, independente de cor, fonte, borda e largura."""
    partes: list[str] = []
    for (linha, coluna), celula in sorted(ws._cells.items()):
        if celula.value is not None:
            partes.append(f"V{linha},{coluna}={_valor(celula.value)}")
        if celula.protection.locked is False:
            partes.append(f"U{linha},{coluna}")
    partes += [f"M{m}" for m in sorted(map(str, ws.merged_cells.ranges))]
    partes += sorted(
        "D" + "|".join(map(str, (
            dv.sqref, dv.type, dv.operator, dv.formula1, dv.formula2, dv.allow_blank,
            dv.showErrorMessage, dv.errorStyle, dv.errorTitle, dv.error,
        )))
        for dv in ws.data_validations.dataValidation
    )
    partes += [f"C{k}" for k, d in sorted(ws.column_dimensions.items()) if d.hidden]
    partes += [f"R{k}" for k, d in sorted(ws.row_dimensions.items()) if d.hidden]
    partes.append(f"F{ws.auto_filter.ref}|{ws.freeze_panes}|{ws.protection.sheet}")
    return hashlib.sha256("\n".join(partes).encode("utf-8")).hexdigest()


# Coleta 12.1 (2ae72bd). Nenhuma aba do template muda de estrutura nesta frente.
# Coleta 12.4: as abas marcadas recebem ROUND(...,2) nas agregacoes monetarias
# e a RESULTADOS ganha SALDO REMANESCENTE — EVOLUCAO (aplicar_coleta_124_precisao_saldo).
# Coleta 12.5: saldo sobe para B65:G72 e o bloco PCs desce para B74:G83; o rotulo
# MEMORIA_RESULTADOS!V68 acompanha (aplicar_coleta_125_saldo_posicao).
ASSINATURAS_TEMPLATE_121: dict[str, str] = {
    "comparativo_VTA": "2a4be709d887f1c136d00c3e918eead47f921d4290070414f67bb78645a06a67",
    "CONTROLE": "8f0c95e47e79daeb831c587957a06a979c3012ac085eb93e2e42b2d822014452",
    "parametros": "e9a30868b83a2e3c276f7d4801ae6c2add0bcc557b92fd79872ca3bab86e38e1",
    "financeiro": "291a637da251e2d579c9dbb85bdc3e01d51814d059ba6b7f1d549d62273c1fed",  # 12.4
    "itens_Remanesc": "09a29c9cacee3851a9288f7986f59d308fc119fd665d8f8a38c0fec8c752f9d1",
    "itens_Consumidos": "f50c1e368498ac92046097cf28b566b1f8a5967e33aaf6ddf7e633bcd2460d63",  # 12.4
    "itens_PC": "fc7fa79ce3c0d519a80239392e7acd161036682c740fdc69bb60396b5b3d1758",  # 12.4
    "aditivos": "b307532ca71a0384856210d49c445db3b479520e7c8e58ebbc5c26c7cd622e61",
    "posicao_referencia": "efbda739d9886cff654c64974511799115730c88f69fe65cbdd52c75cc388ea5",
    "posicao_contratual": "6cbdd37302efba8e84f55718fd2b8fb7b137ea87f26fa48b0698dd0cd4b179f1",
    "itens_RC": "efd87949078538d682dc306a0a33fcb894000e6864386fb8188c65d3ee8ef567",
    "historico_VU": "de7563cfcbda1421db08cb47e4130f5b865c57c2640c0794e8739d54b0aade42",
    "cobertura_temporal": "b073ae3439ee845443e506347cc84777b4470a9c1d910e9bde1485c47e5fdb1d",  # 12.6
    # 12.6: B10:B14/B15/B44 exigem EFEITO_FINANCEIRO comprovado na linha paga.
    "MEMORIA_RESULTADOS": "d0e25ff6745349d34c13bfac4a6d189e05dafef259076137774a78106ff8aa5e",  # 12.6
    "RESULTADOS_DETALHE": "028ce6c370008be3e7802b1d69f7e3262973e8cd256fed00410312959bedae93",  # 12.4
    "RESULTADOS": "29e39099892cd2cd898ff770b3b66cf4ccd6d9dcf4930ac1e8c99d00388a46d8",  # 12.5
}
ASSINATURA_CICLO_121 = "44775bfa3c587759ab8d0f1fc177119355cdef65a06c74bfd032a31721ded638"


# =========================================================================== #
# INTEGRIDADE ESTRUTURAL
# =========================================================================== #
def assinatura_valores(ws, faixa: str) -> str:
    """Hash de valores/formulas e celulas editaveis de uma faixa da aba."""
    partes = [
        f"{c.coordinate}={_valor(c.value)}|{c.protection.locked}"
        for linha in ws[faixa] for c in linha
    ]
    return hashlib.sha256("\n".join(partes).encode("utf-8")).hexdigest()


# CONTROLE ganha o ORIENTADOR DA COLETA (D4:H34 + auxiliares J:L ocultas); o
# bloco A:B funcional segue identico ao da 12.1. A unica excecao visual e o
# cabecalho do RESUMO DO REAJUSTE, alinhado de A6:B6 para A5:B5.
ASSINATURAS_CONTROLE_A_B_121 = {
    "A1:B3": "86458ff2a1f66102f6969fb8d869256281535ec565626a729d8c1435a02170fb",
    "A7:B14": "e9a8686806ae4bfc812c9c628cdd46c34e6159749f5574bccb0be809597cfe52",
}


def test_template_tem_a_mesma_estrutura_da_coleta_121_em_todas_as_abas():
    wb = _template()
    assert wb.sheetnames == list(ASSINATURAS_TEMPLATE_121)
    for aba, esperado in ASSINATURAS_TEMPLATE_121.items():
        if aba == "CONTROLE":
            continue                   # orientador da coleta: teste proprio abaixo
        assert assinatura(wb[aba]) == esperado, aba


def test_controle_bloco_a_b_funcional_identico_ao_da_121():
    for wb in (_template(), _gerado()):
        ws = wb["CONTROLE"]
        for faixa, esperado in ASSINATURAS_CONTROLE_A_B_121.items():
            assert assinatura_valores(ws, faixa) == esperado, faixa


def test_controle_sem_a_orientacao_do_fluxo_antigo_em_a4_b4():
    for wb in (_template(), _gerado()):
        ws = wb["CONTROLE"]
        assert (ws["A4"].value, ws["B4"].value) == (None, None)
        assert ws["D4"].value == "Data-limite considerada nos cálculos e resultados."
        textos = " ".join(str(c.value) for linha in ws.iter_rows() for c in linha if c.value)
        assert "ciclos fora da analise" not in textos and "Preencher FINANCEIRO" not in textos


def test_ciclo_em_execucao_gerada_tem_a_mesma_estrutura_da_121():
    # Mesmas formulas, valores, coordenadas, mesclas, validacoes, celulas
    # editaveis (D5 e C13:C211), colunas tecnicas ocultas e protecao.
    assert assinatura(_gerado()[ABA_CICLO]) == ASSINATURA_CICLO_121


# =========================================================================== #
# AJUSTE 1 — itens_PC: separacao L|M so por estilo
# =========================================================================== #
def test_itens_pc_separa_entrada_e_resultados_por_borda_dupla():
    # Nenhuma coluna inserida: A:L continuam as colunas oficiais (o runtime so
    # troca os rotulos de H1/J1 na copia entregue) e M1 o quadro.
    pc = _template()["itens_PC"]
    assert [pc.cell(1, c).value for c in range(1, 13)] == COLUNAS_ITENS_PC_OFICIAL[:12]
    for wb in (_template(), _gerado()):
        ws = wb["itens_PC"]
        assert ws["M1"].value == av.CABECALHO_PC_ESPERADO["M1"]
        for linha in (1, 2, 9, 10, 33, 2500, 5001):
            lado = ws[f"L{linha}"].border.right
            assert lado.style == "double", linha
            assert lado.color.rgb == "FF" + av.COR_SEPARACAO_PC, linha
        # O resto da grade de dados segue com a borda fina.
        assert ws["K2"].border.right.style == "thin"
        assert ws["A2"].border.left.style == "thin"
        assert ws.auto_filter.ref == "A1:L5001"


# =========================================================================== #
# AJUSTE 2 — aditivos: mensagem NOVO ITEM em A2:A200
# =========================================================================== #
def _validacoes(ws) -> dict:
    return {str(dv.sqref): dv for dv in ws.data_validations.dataValidation}


def test_aditivos_item_mostra_orientacao_de_novo_item():
    for wb in (_template(), _gerado()):
        ws = wb["aditivos"]
        assert ws["A1"].value == "ITEM"                      # cabecalho curto
        item = _validacoes(ws)[av.FAIXA_ITEM_ADITIVOS]
        assert item.type is None                             # qualquer valor
        assert item.showInputMessage and item.promptTitle == "NOVO ITEM"
        mensagem = item.prompt.replace("_x000a_", "\n")
        assert mensagem == av.MENSAGEM_NOVO_ITEM
        assert "Cadastre primeiro em itens_Remanesc como N001, N002..." in mensagem
        assert "A QTD_BASE_ORIGINAL ficará 0 automaticamente." in mensagem
        assert "Informe o VU_ORIGINAL e, depois, registre aqui a data e a quantidade" in mensagem
        assert len(mensagem) <= 255


def test_aditivos_dropdowns_e_demais_validacoes_preservados():
    for wb in (_template(), _gerado()):
        dvs = _validacoes(wb["aditivos"])
        assert set(dvs) == {"H2:H200 M2:M200", "D2:D200", "A2:A200", "E2:E200", "G2:G200"}
        assert (dvs["H2:H200 M2:M200"].type, dvs["H2:H200 M2:M200"].formula1) == (
            "list", '"Sim,Nao"')
        assert dvs["D2:D200"].formula1 == '"Acrescimo,Acréscimo - novo item,Supressao"'
        assert (dvs["G2:G200"].type, dvs["G2:G200"].formula1) == ("list", '"C0,C1,C2,C3,C4"')
        assert dvs["E2:E200"].promptTitle == "QUANTIDADE DO ADITIVO"
        assert dvs["D2:D200"].promptTitle == "ITEM NOVO"


# =========================================================================== #
# AJUSTE 3 — CONTROLE!B1: largura + orientacao em C1 por formatacao condicional
# =========================================================================== #
def _regras_c1(ws):
    # Alerta 12.0 em C1; orientacao do metodo em C1:H1 (C estreita no
    # orientador da coleta: o fundo acompanha o texto que transborda).
    por_faixa = {str(faixa.sqref): faixa.rules for faixa in ws.conditional_formatting}
    return [por_faixa["C1"][0], por_faixa["C1:H1"][0]]


def test_controle_b1_mesmo_dropdown_e_coluna_mais_larga():
    from _metodo_apuracao import OPCOES_DROPDOWN

    # Template: lista homologada; copia entregue: rotulos canonicos reescritos
    # em runtime (_garantir_dropdown_metodo) — ambos inalterados nesta frente.
    opcoes = {"template": av.OPCOES_METODO, "gerado": ",".join(OPCOES_DROPDOWN)}
    for origem, wb in (("template", _template()), ("gerado", _gerado())):
        ws = wb["CONTROLE"]
        lista = [dv for dv in ws.data_validations.dataValidation if str(dv.sqref) == "B1"]
        assert len(lista) == 1
        assert (lista[0].type, lista[0].formula1) == ("list", f'"{opcoes[origem]}"')
        assert max(map(len, opcoes[origem].split(","))) <= av.LARGURA_CONTROLE_B - 8
        assert ws["B1"].value == "SELECIONE AQUI"
        assert ws["B1"].protection.locked is False
        assert ws.column_dimensions["B"].width >= 33


def test_controle_orientacao_nao_sobrescreve_nada():
    for wb in (_template(), _gerado()):
        ws = wb["CONTROLE"]
        assert ws["C1"].value == av.FORMULA_C1               # alerta 12.0 intacto
        assert all(ws[f"{c}1"].value is None for c in "DEFGH")
        regras = _regras_c1(ws)
        assert [r.formula for r in regras] == [
            ['LEFT($C$1,8)="CONFERIR"'], [av.FORMULA_CF_ORIENTACAO[1:]]]
        orientacao = regras[1].dxf
        assert orientacao.numFmt.formatCode == av.FORMATO_ORIENTACAO
        assert "← Selecione o método de apuração" in orientacao.numFmt.formatCode
        assert orientacao.fill.bgColor.rgb.endswith(av.COR_ORIENTACAO_FUNDO)
        # Formula sem funcoes: identica em qualquer idioma; so le CONTROLE.
        assert "(" not in av.FORMULA_CF_ORIENTACAO.replace('($B$1="', "")
        assert "!" not in av.FORMULA_CF_ORIENTACAO


# =========================================================================== #
# AJUSTE 4 — CICLO_EM_EXECUCAO: paleta clara das demais abas
# =========================================================================== #
AMARELO_CLARO, CINZA_CLARO, AZUL_CLARO, VERDE_CLARO = (
    "FFFFF2CC", "FFEDEDED", "FFDDEBF7", "FFE2EFDA")


def _fundo(celula) -> str:
    return celula.fill.fgColor.rgb


def _luminancia(argb: str) -> float:
    r, g, b = (int(argb[i:i + 2], 16) / 255 for i in (2, 4, 6))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def test_ciclo_entradas_amarelas_e_editaveis_calculos_cinza_e_travados():
    ws = _gerado()[ABA_CICLO]
    assert _fundo(ws["D5"]) == AMARELO_CLARO and ws["D5"].protection.locked is False
    for linha in (13, 100, 211):
        entrada = ws.cell(linha, 3)
        assert _fundo(entrada) == AMARELO_CLARO and entrada.protection.locked is False
        for coluna in (1, 2, 4, 5, 6, 7):
            calculo = ws.cell(linha, coluna)
            assert _fundo(calculo) == CINZA_CLARO, calculo.coordinate
            assert calculo.protection.locked is not False, calculo.coordinate
    assert _fundo(ws["C3"]) == _fundo(ws["F3"]) == CINZA_CLARO
    assert _fundo(ws["A11"]) == CINZA_CLARO


def test_ciclo_cabecalhos_azuis_e_total_verde():
    ws = _gerado()[ABA_CICLO]
    for coord in ("A1", "A3", "D3", "A5", "A7", "A12", "B12", "D12", "E12", "F12", "G12"):
        assert _fundo(ws[coord]) == AZUL_CLARO, coord
    assert _fundo(ws["C12"]) == AMARELO_CLARO                 # cabecalho da entrada
    assert _fundo(ws["A9"]) == VERDE_CLARO                    # total validado
    assert ws["A9"].number_format == "R$ #,##0.00"


def test_ciclo_sem_fundos_escuros_e_com_fontes_escuras():
    ws = _gerado()[ABA_CICLO]
    for linha in ws.iter_rows(min_row=1, max_row=211, max_col=7):
        for celula in linha:
            if celula.fill.fill_type == "solid":
                assert _luminancia(_fundo(celula)) > 0.8, celula.coordinate
            cor = celula.font.color.rgb if celula.font.color is not None else None
            if isinstance(cor, str):
                assert _luminancia(cor) < 0.5, celula.coordinate


def test_ciclo_validacao_temporal_so_ganha_cor():
    ws = _gerado()[ABA_CICLO]
    regras = {r.formula[0]: r.dxf.fill.bgColor.rgb for faixa in ws.conditional_formatting
              if str(faixa.sqref) == "A11:G11" for r in faixa.rules}
    assert regras == {
        '$A$11="POSICAO FISICA NA DATA DE CORTE - REFERENCIA VALIDADA."': VERDE_CLARO,
        'OR($A$11="POSICAO FISICA POSTERIOR A DATA DE CORTE.",'
        '$A$11="DATA DE CORTE DA APURACAO NAO INFORMADA EM CONTROLE!B3.")': "FFFBE5E8",
    }
    # As mensagens coloridas sao exatamente as que a formula de A11 produz.
    for chave in regras:
        for texto in chave.split('"')[1::2]:
            assert texto in ws["A11"].value


def test_ciclo_regras_condicionais_anteriores_preservadas():
    ws = _gerado()[ABA_CICLO]
    faixas = sorted(str(f.sqref) for f in ws.conditional_formatting)
    assert faixas == sorted(["D5", "A13:G211", "C13:C211", "A13:I211", "A11:G11"])


def test_ciclo_colunas_mais_confortaveis():
    ws = _gerado()[ABA_CICLO]
    larguras = {c: ws.column_dimensions[c].width for c in "ABCDEFG"}
    assert larguras == {"A": 20, "B": 26, "C": 26, "D": 26, "E": 20, "F": 22, "G": 28}
    assert ws.row_dimensions[12].height == 58


# =========================================================================== #
# VERSAO
# =========================================================================== #
def test_marcador_do_xls_mostra_a_coleta_123_em_a16_b17():
    ws = _gerado()["CONTROLE"]
    assert [ws[c].value for c in ("A16", "B16", "A17", "B17")] == [
        "Modelo de Coleta", "12.6", "Gerado pelo Cl8us", "12.13"]
    assert all(ws[c].value is None for c in ("A24", "B24", "A25", "B25"))
    assert not ws.row_dimensions[16].hidden and not ws.row_dimensions[17].hidden


def test_leitura_do_marcador_por_rotulo_vale_para_a_nova_e_as_anteriores():
    from _compatibilidade_coleta import detectar_linhagem_coleta

    assert detectar_linhagem_coleta(_gerado())["marcador_publico"] == "12.6"
    anterior = RAIZ / "tests" / "fixtures" / "coletas_compat" / "coleta_11_pc.xlsx"
    wb = load_workbook(anterior, data_only=False)
    assert wb["CONTROLE"]["B24"].value == "11.0"                 # marcador antigo
    assert detectar_linhagem_coleta(wb)["marcador_publico"] == "11.0"


# =========================================================================== #
# ORIENTADOR DA COLETA (CONTROLE)
# =========================================================================== #
import aplicar_coleta_122_orientador_controle as orc  # noqa: E402


def test_orientador_formulas_batem_com_o_descritor_e_sao_ascii():
    orc.validar_ascii_e_parenteses()
    for wb in (_template(), _gerado()):
        ws = wb["CONTROLE"]
        for celula, formula in orc.formulas().items():
            assert ws[celula].value == formula, celula
        for linha, (chave, texto) in orc.CONSTANTES.items():
            assert (ws[f"J{linha}"].value, ws[f"K{linha}"].value) == (chave, texto), linha


def test_orientador_le_somente_o_metodo_fora_da_controle():
    externas = set()
    for formula in orc.formulas().values():
        externas |= {ref.split("!")[0] for ref in formula.replace("(", " ").replace(",", " ")
                     .replace("&", " ").replace("=", " ").split() if "!" in ref}
    assert externas == {"MEMORIA_RESULTADOS"}
    # So leitura: nenhuma formula de outra aba aponta para o orientador.
    for aba in _template().worksheets:
        if aba.title == "CONTROLE":
            continue
        for linha in aba.iter_rows():
            for celula in linha:
                if isinstance(celula.value, str) and "CONTROLE!" in celula.value:
                    for alvo in ("$D$", "$E$", "$F$", "$J$", "$K$", "$L$"):
                        assert f"CONTROLE!{alvo}" not in celula.value, celula.coordinate


def test_orientador_textos_fixos_e_layout():
    for wb in (_template(), _gerado()):
        ws = wb["CONTROLE"]
        assert ws["C3"].value is None
        assert ws["D4"].value == "Data-limite considerada nos cálculos e resultados."
        assert ws["D5"].value == "ORIENTADOR DA COLETA — O QUE PEDIR AO FISCAL"
        assert ws["D6"].value.startswith("Use este quadro para definir o mínimo necessário")
        assert ws["A5"].value == "RESUMO DO REAJUSTE"
        assert ws["B5"].value == "VALOR"
        assert ws["A6"].value is None and ws["B6"].value is None
        assert [ws[f"{c}7"].value for c in "DEF"] == [
            "PARÂMETRO COMPLEMENTAR", "SELEÇÃO", "COMO ISSO MUDA O PEDIDO AO FISCAL"]
        assert ws["D8"].value == "Histórico da execução está completo?"
        assert ws["D9"].value == "Aditivos/supressões no período"
        assert [ws[f"{c}10"].value for c in "EFGH"] == ["C1", "C2", "C3", "C4"]
        assert ws["D10"].value == "Pedido de reajuste?"
        assert ws["D11"].value == "Situação"
        assert ws["D12"].value == "Ciclos com pedido nesta análise"
        assert [ws[f"{c}16"].value for c in "DEF"] == [
            "INFORMAÇÃO", "STATUS", "POR QUE / PARA QUE SERVE"]
        assert ws["D25"].value == "O QUE VOCÊ VAI PEDIR AO FISCAL"
        assert "CHAR(10)&CHAR(10)" in ws["D26"].value
        assert "$K$96" in ws["D26"].value
        assert "D26:H34" in {str(r) for r in ws.merged_cells.ranges}
        assert not any(str(r) in {f"D{x}:H{x}" for x in range(26, 35)}
                       for r in ws.merged_cells.ranges)
        conteudo = str([c.value for linha in ws.iter_rows() for c in linha])
        for texto_obsoleto in (
            "Leitura automática (parametros)", "Informar, se não houver leitura", "Considerado"
        ):
            assert texto_obsoleto not in conteudo
        assert "Base histórica do método" not in str(
            [c.value for linha in ws.iter_rows() for c in linha])
        assert ws.column_dimensions["C"].width < 4
        # J:L podem vir agrupadas num unico <col min..max>: confere pela faixa.
        ocultas = {i for d in ws.column_dimensions.values() if d.hidden
                   for i in range(d.min, d.max + 1)}
        assert {10, 11, 12} <= ocultas and not {4, 5, 6, 7, 8} & ocultas
        assert ws.sheet_view.showGridLines is False
        for linha in (4, 15, 16, 17, 18):
            assert not ws.row_dimensions[linha].hidden, linha


def test_orientador_dropdowns_e_campos_editaveis():
    for wb in (_template(), _gerado()):
        ws = wb["CONTROLE"]
        dvs = {str(dv.sqref): dv for dv in ws.data_validations.dataValidation}
        assert (dvs["E8"].type, dvs["E8"].formula1) == ("list", "$K$56:$K$58")
        assert (dvs["E9"].type, dvs["E9"].formula1) == ("list", "$K$60:$K$62")
        assert (dvs["E11:H11"].type, dvs["E11:H11"].formula1) == ("list", "$K$64:$K$65")
        assert [ws[f"K{r}"].value for r in (56, 57, 58)] == ["Sim", "Parcialmente", "Não confirmado"]
        assert [ws[f"K{r}"].value for r in (60, 61, 62)] == ["Não", "Sim", "Não sei"]
        assert [ws[f"K{r}"].value for r in (64, 65)] == ["Sim", "Não"]
        for celula in ("E8", "E9", "E11", "F11", "G11", "H11"):
            assert ws[celula].protection.locked is False, celula
            assert ws[celula].fill.fgColor.rgb.endswith("FFF2CC"), celula
        for celula in ("D26", "G12", "E17", "J48"):
            assert ws[celula].protection.locked is not False, celula


def test_orientador_status_coloridos_pelo_codigo_oculto():
    for wb in (_template(), _gerado()):
        ws = wb["CONTROLE"]
        regras = {r.formula[0] for faixa in ws.conditional_formatting
                  if str(faixa.sqref) == "E17:E23" for r in faixa.rules}
        assert regras == {f"$J17={codigo}" for codigo in orc.CORES_STATUS}
        assert [ws[f"K{r}"].value for r in range(50, 55)] == [
            "OBRIGATÓRIO", "COMPLEMENTAR", "DISPENSÁVEL", "VERIFICAR", "NÃO SE APLICA"]


def test_orientador_regra_abertura_nunca_dispensada_e_sem_literal_de_ano():
    # A posicao na abertura do ultimo ciclo e sempre OBRIGATORIO (codigo 1).
    assert orc.DIAGNOSTICO[3][1] == "=1"
    textos = " ".join(t for _l, (_c, t) in orc.CONSTANTES.items())
    assert "YYYY" not in textos and "até o dia atual" not in textos
    assert "mais recentes disponíveis até a data de corte da apuração" in textos
    # dd/mm sem ano e sem TEXT() (formato de TEXT depende do idioma).
    assert "YEAR" not in orc.ESTADO["J38"] and "TEXT(" not in orc.ESTADO["J38"]


# --------------------------------------------------------------------------- #
# Excel real (RUN_EXCEL_INTEGRATION=1): textos dinamicos por cenario
# --------------------------------------------------------------------------- #
excel_real = pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="defina RUN_EXCEL_INTEGRATION=1 para executar o Excel COM",
)
PC, FIN, IT = "PC (Pedidos de Compra)", "Financeiro (Mensalidade)", "Itens Consumidos"


@pytest.fixture(scope="module")
def cenarios_excel(tmp_path_factory):
    import gc
    import pythoncom
    import win32com.client

    from _coleta_oficial import gerar_coleta_oficial_preenchida
    from test_coleta_118_excel_real import _dados, _tentar

    caminho = tmp_path_factory.mktemp("orientador") / "Coleta_orientador.xlsx"
    caminho.write_bytes(gerar_coleta_oficial_preenchida(_dados()))
    casos = {
        "pc_completo": (PC, "Sim", "Não", "Sim"),
        "pc_parcial_aditivos": (PC, "Parcialmente", "Sim", "Sim"),
        "fin_nao_confirmado_nao_sei": (FIN, "Não confirmado", "Não sei", "Sim"),
        "itens": (IT, "Sim", "Não", "Sim"),
        "sem_metodo": ("SELECIONE AQUI", None, None, None),
    }
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    leituras = {}
    try:
        livro = _tentar(pythoncom, lambda: excel.Workbooks.Open(
            str(caminho.resolve()), UpdateLinks=0, ReadOnly=False, CorruptLoad=0))
        ws = livro.Worksheets("CONTROLE")
        for nome, (metodo, hist, adt, pedidos) in casos.items():
            def passo():
                ws.Unprotect()
                ws.Range("B1").Value = metodo
                ws.Range("B3").Value = datetime(2027, 8, 31)
                ws.Range("E8").Value = hist
                ws.Range("E9").Value = adt
                ws.Range("E11").Value = pedidos
                for coluna in "FGH":                       # C2..C4 sem pedido
                    ws.Range(f"{coluna}11").Value = "Não" if pedidos else None
                excel.CalculateFull()
                return {
                    "ciclo": ws.Range("B2").Text,
                    "corte": ws.Range("B3").Value,
                    "saida": [linha for linha in ws.Range("D26").Text.splitlines() if linha],
                    "ciclos_pedido": ws.Range("G12").Text,
                    "status": [ws.Range(f"E{r}").Text for r in range(17, 24)],
                    "c1": ws.Range("C1").Text,
                }
            leituras[nome] = _tentar(pythoncom, passo)
        _tentar(pythoncom, lambda: livro.Close(SaveChanges=False))
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
        excel = None
        gc.collect()
        pythoncom.CoUninitialize()
    return leituras


def _ddmm(valor) -> str:
    return f"{valor.day:02d}/{valor.month:02d}"


@excel_real
def test_excel_pc_historico_completo_sem_aditivos(cenarios_excel):
    r = cenarios_excel["pc_completo"]
    ciclo = r["ciclo"]
    assert ciclo == "C4"
    assert r["saida"] == [
        "Para concluir a apuração, favor informar:",
        "1. os Pedidos de Compra mais recentes disponíveis até a data de corte da "
        "apuração, indicando quais foram pagos/aceitos e quais não foram;",
        f"2. as quantidades remanescentes por item na abertura do {ciclo};",
        f"3. as quantidades remanescentes por item na data de corte ({_ddmm(r['corte'])}), "
        "correspondente à posição atual.",
        "Objetivos da coleta: apurar retroativo, VU por ciclo, VTA e o valor remanescente "
        "atualizado na abertura do último ciclo.",
    ]
    assert r["ciclos_pedido"] == "C1"
    assert r["status"] == ["OBRIGATÓRIO", "NÃO SE APLICA", "DISPENSÁVEL", "OBRIGATÓRIO",
                           "OBRIGATÓRIO", "NÃO SE APLICA", "DISPENSÁVEL"]


@excel_real
def test_excel_historico_parcial_e_aditivos_acrescentam_itens(cenarios_excel):
    r = cenarios_excel["pc_parcial_aditivos"]
    ciclo = r["ciclo"]
    assert r["saida"][2] == ("2. posição física ou evidência equivalente apenas para os "
                             "períodos não cobertos pelo histórico da execução;")
    assert r["saida"][4] == (f"4. os acréscimos e supressões ocorridos desde a abertura do "
                             f"{ciclo}, com item, quantidade e data de efeito;")
    assert r["saida"][-2].startswith("5. as quantidades remanescentes por item na data de corte")
    assert r["status"][2] == "COMPLEMENTAR" and r["status"][5] == "OBRIGATÓRIO"


@excel_real
def test_excel_financeiro_e_itens(cenarios_excel):
    fin = cenarios_excel["fin_nao_confirmado_nao_sei"]
    assert fin["saida"][1].startswith("1. os valores do financeiro por competência")
    assert any("a confirmação de que o histórico da execução cobre" in t for t in fin["saida"])
    assert any("a confirmação sobre acréscimos ou supressões" in t for t in fin["saida"])
    assert fin["status"][2] == "VERIFICAR" and fin["status"][5] == "VERIFICAR"
    itens = cenarios_excel["itens"]
    assert itens["saida"][1].startswith(
        "1. as quantidades remanescentes por item na abertura de cada ciclo anterior")
    assert itens["status"][:3] == ["NÃO SE APLICA", "NÃO SE APLICA", "OBRIGATÓRIO"]


@excel_real
def test_excel_sem_metodo_orienta_selecao(cenarios_excel):
    r = cenarios_excel["sem_metodo"]
    assert r["saida"] == [
        "Selecione primeiro o método de apuração em CONTROLE!B1.",
        "Objetivos da coleta: apurar retroativo, VU por ciclo, VTA e o valor remanescente "
        "atualizado na abertura do último ciclo.",
    ]
    assert r["c1"] == "← Selecione o método de apuração"
