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
import sys
from pathlib import Path

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
ASSINATURAS_TEMPLATE_121: dict[str, str] = {
    "comparativo_VTA": "2a4be709d887f1c136d00c3e918eead47f921d4290070414f67bb78645a06a67",
    "CONTROLE": "8f0c95e47e79daeb831c587957a06a979c3012ac085eb93e2e42b2d822014452",
    "parametros": "e9a30868b83a2e3c276f7d4801ae6c2add0bcc557b92fd79872ca3bab86e38e1",
    "financeiro": "d48da5ea0c3da9ebbd82fdc0df1ff325826c193f7b0d853f5025e1cecc08702d",
    "itens_Remanesc": "09a29c9cacee3851a9288f7986f59d308fc119fd665d8f8a38c0fec8c752f9d1",
    "itens_Consumidos": "ed2b7648f27b92b2a711b14f4ac24f36047f01b339906a1ec53aa19499ac2aa5",
    "itens_PC": "d1a21adb9af3cfcd488b80d6d027f156e210bb56a4ddcdaaf29ed12c2c68c8ef",
    "aditivos": "b307532ca71a0384856210d49c445db3b479520e7c8e58ebbc5c26c7cd622e61",
    "posicao_referencia": "efbda739d9886cff654c64974511799115730c88f69fe65cbdd52c75cc388ea5",
    "posicao_contratual": "6cbdd37302efba8e84f55718fd2b8fb7b137ea87f26fa48b0698dd0cd4b179f1",
    "itens_RC": "efd87949078538d682dc306a0a33fcb894000e6864386fb8188c65d3ee8ef567",
    "historico_VU": "de7563cfcbda1421db08cb47e4130f5b865c57c2640c0794e8739d54b0aade42",
    "cobertura_temporal": "73a4170a46302eb390ec72fada1c82cc326d1db97c0a47c23bc12979bba6b708",
    "MEMORIA_RESULTADOS": "b905191d1f3c68b3d565a8f88749055a3fb2e455a1f4d9ef7c00857b2b34d9f3",
    "RESULTADOS_DETALHE": "68beb26b643b0d475217a101249b163f4ee9e5fe593d6f33974e5de5761df989",
    "RESULTADOS": "b4ce52bb2564d8691ec7281369c5eea644349615583cd904a64d1bd03af0382f",
}
ASSINATURA_CICLO_121 = "44775bfa3c587759ab8d0f1fc177119355cdef65a06c74bfd032a31721ded638"


# =========================================================================== #
# INTEGRIDADE ESTRUTURAL
# =========================================================================== #
def test_template_tem_a_mesma_estrutura_da_coleta_121_em_todas_as_abas():
    wb = _template()
    assert wb.sheetnames == list(ASSINATURAS_TEMPLATE_121)
    for aba, esperado in ASSINATURAS_TEMPLATE_121.items():
        assert assinatura(wb[aba]) == esperado, aba


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
    return [regra for faixa in ws.conditional_formatting
            if str(faixa.sqref) == "C1" for regra in faixa.rules]


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
        assert all(ws[f"{c}1"].value is None for c in "DEFG")
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
def test_marcador_do_xls_mostra_a_coleta_122():
    ws = _gerado()["CONTROLE"]
    assert (ws["B24"].value, ws["B25"].value) == ("12.2", "12.9")
