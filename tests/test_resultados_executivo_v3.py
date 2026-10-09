"""RESULTADOS-EXECUTIVO-V3 — RESULTADOS executiva + RESULTADOS_DETALHE.

Tres contratos separados, como pedido na frente:

A. ECONOMICO — a aba executiva nao calcula: so espelha nomes definidos e
   celulas canonicas; a camada tecnica (RESULTADOS_DETALHE) manteve, celula a
   celula, as formulas homologadas da antiga RESULTADOS.
B. TECNICO   — nomes definidos, ordem/visibilidade das abas, validacoes e
   acoplamento XLS->XLS dos ajustes manuais; leitores Python resolvem a aba
   tecnica em arquivo novo E em arquivo anterior; seguranca e versionamento.
C. VISUAL    — formatos numericos reais (xx,xx% / dd/mm/aaaa / R$), mesclagens
   simples sem sobreposicao, larguras estaveis, hiperlink dos ajustes.

A prova no Excel REAL (### a 100%, encaixe de texto, reparo) roda com
RUN_EXCEL_INTEGRATION=1 (`tools/validar_visual_resultados_v3.py`).
"""
from __future__ import annotations

import io
import json
import os
import re
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "tools"))

import aplicar_resultados_executivo_v3 as v3  # noqa: E402
from _resultados_abas import (  # noqa: E402
    ABA_RESULTADOS,
    ABA_RESULTADOS_DETALHE,
    TITULO_RESULTADOS_DETALHE,
    TITULO_RESULTADOS_EXECUTIVO,
    aba_resultados_tecnica,
)

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
# Formulas A1:J87 da aba RESULTADOS de origin/main 8a7dcc2 (pre-migracao).
SNAPSHOT_PRE = RAIZ / "tests" / "baseline_resultados" / "inventario_pre_executivo_v3.json"

_CACHE: dict[str, object] = {}


def _template():
    if "template" not in _CACHE:
        _CACHE["template"] = load_workbook(TEMPLATE)
    return _CACHE["template"]


def _gerado():
    """Coleta EFETIVAMENTE entregue pela aplicacao (load/save do openpyxl)."""
    if "gerado" not in _CACHE:
        from _coleta_oficial import obter_coleta_oficial_bytes

        _CACHE["gerado"] = load_workbook(io.BytesIO(obter_coleta_oficial_bytes()))
    return _CACHE["gerado"]


def _bytes(wb) -> bytes:
    saida = io.BytesIO()
    wb.save(saida)
    return saida.getvalue()


def _formulas_executivas(wb) -> dict[str, str]:
    ws = wb[ABA_RESULTADOS]
    return {
        c.coordinate: c.value
        for linha in ws.iter_rows() for c in linha
        if isinstance(c.value, str) and c.value.startswith("=")
    }


def _sem_literais(formula: str) -> str:
    return re.sub(r'"[^"]*"', '""', formula)


# =========================================================================== #
# A. CONTRATO ECONOMICO
# =========================================================================== #
def test_aba_executiva_nao_tem_motor_de_calculo():
    """Nenhuma soma, arredondamento ou operacao aritmetica na aba executiva."""
    proibidas = ("SUM", "ROUND", "PRODUCT", "AVERAGE", "MINIFS", "MAXIFS", "INDIRECT")
    formulas = _formulas_executivas(_template())
    assert len(formulas) == len(v3.formulas_executivo())
    for endereco, formula in formulas.items():
        corpo = _sem_literais(formula)[1:].upper()
        for funcao in proibidas:
            assert funcao + "(" not in corpo, f"{endereco} usa {funcao}: {formula}"
        sem_refs = re.sub(r"[A-Z_]+!\$?[A-Z]{1,3}\$?\d+", "R", corpo)
        assert not re.search(r"[*/+]", sem_refs), f"{endereco} tem aritmetica: {formula}"
        assert not re.search(r"[A-Z0-9)]\s*-\s*[A-Z0-9(]", sem_refs), (
            f"{endereco} tem subtracao: {formula}"
        )


def test_aba_executiva_so_le_fontes_canonicas():
    """Cada referencia aponta para a camada tecnica, MEMORIA, parametros,
    CONTROLE ou um nome definido canonico existente."""
    nomes = set(_template().defined_names)
    abas_ok = {ABA_RESULTADOS_DETALHE, "MEMORIA_RESULTADOS", "parametros", "CONTROLE"}
    canonicos = {"STATUS_RESULTADOS", "VTA_FINAL", "METODO_RETROATIVO",
                 "RETROATIVO_POTENCIAL_VTA", "RETROATIVO_POTENCIAL_APURADO",
                 "RETROATIVO_POTENCIAL_NEGATIVO", "SALDO_REMANESCENTE_ATUAL",
                 "VTA_SEM_POTENCIAL", "CONFERENCIA_FORMACAO_VTA"}
    usados: set[str] = set()
    for endereco, formula in _formulas_executivas(_template()).items():
        corpo = _sem_literais(formula)
        for aba in re.findall(r"([A-Za-z_]+)!", corpo):
            assert aba in abas_ok, f"{endereco} le a aba {aba}"
        usados |= {t for t in re.findall(r"\b[A-Z][A-Z_]{4,}\b", corpo) if t in canonicos}
    assert usados <= nomes
    assert {"VTA_FINAL", "STATUS_RESULTADOS", "RETROATIVO_POTENCIAL_VTA"} <= usados


def test_cards_espelham_as_mesmas_fontes_da_aba_anterior():
    f = v3.formulas_executivo()
    assert "VTA_FINAL" in f["B9"]
    assert "RESULTADOS_DETALHE!$D$22" in f["C9"]          # = RETRO_OFICIAL
    assert "RETROATIVO_POTENCIAL_VTA" in f["D9"]
    assert "SALDO_REMANESCENTE_ATUAL" in f["E9"]
    assert "RESULTADOS_DETALHE!$D$6" in f["F6"] and f["F9"] == "=F6"
    assert "STATUS_RESULTADOS" in f["G6"]


def test_quadro_sem_efeito_e_espelho_integral_do_pr172():
    """E15:H21 do detalhe chegam inteiros; nada entra em VTA ou total."""
    f = v3.formulas_executivo()
    for n, linha in zip(range(1, 5), range(54, 58)):
        origem = 16 + n
        assert f[f"C{linha}"] == (
            f'=IF(RESULTADOS_DETALHE!$F${origem}="","",RESULTADOS_DETALHE!$F${origem})'
        )
        assert f[f"D{linha}"].endswith(f"RESULTADOS_DETALHE!$G${origem})")
        assert f[f"E{linha}"].endswith(f"RESULTADOS_DETALHE!$H${origem})")
        assert f[f"F{linha}"] == f'=RESULTADOS_DETALHE!$E${origem}&""'
    assert f["B58"] == '=RESULTADOS_DETALHE!$E$21&""'
    usados_fora = [
        k for k, formula in f.items()
        if re.search(r"RESULTADOS_DETALHE!\$[FGH]\$(1[7-9]|20)\b", formula)
        and not re.match(r"[C-E](5[4-7])$", k)
    ]
    assert not usados_fora, usados_fora


def test_quadro_2_espelha_tabela_2_inclusive_total_e_ajuste():
    f = v3.formulas_executivo()
    for n in range(5):
        for col_exe, col_det in (("C", "B"), ("D", "C"), ("E", "D")):
            assert f"RESULTADOS_DETALHE!${col_det}${16 + n}" in f[f"{col_exe}{43 + n}"]
    assert "RESULTADOS_DETALHE!$D$21" in f["E48"]
    for col_exe, col_det in (("C", "B"), ("D", "C"), ("E", "D")):
        assert f"RESULTADOS_DETALHE!${col_det}$22" in f[f"{col_exe}49"]


def test_camada_tecnica_preserva_formulas_homologadas():
    """RESULTADOS_DETALHE == RESULTADOS da main (exceto titulos A1/A2)."""
    esperado = json.loads(SNAPSHOT_PRE.read_text(encoding="utf-8"))["formulas"]
    ws = _template()[ABA_RESULTADOS_DETALHE]
    obtido = {
        c.coordinate: c.value for linha in ws.iter_rows(max_row=87, max_col=10)
        for c in linha if c.value is not None
    }
    diferentes = sorted(
        k for k in set(esperado) | set(obtido)
        if esperado.get(k) != obtido.get(k) and k not in ("A1", "A2", "C12")
    )
    assert not diferentes, f"formulas da camada tecnica mudaram: {diferentes[:10]}"
    assert obtido["C12"] == esperado["C12"].replace(
        "x RESULTADOS!H5", "x RESULTADOS_DETALHE!H5"
    )


# =========================================================================== #
# B. CONTRATO TECNICO
# =========================================================================== #
@pytest.mark.parametrize("origem", ["template", "gerado"])
def test_ordem_e_visibilidade_das_abas(origem):
    wb = _template() if origem == "template" else _gerado()
    assert wb.sheetnames[-1] == ABA_RESULTADOS
    assert wb.sheetnames[-2] == ABA_RESULTADOS_DETALHE
    assert wb[ABA_RESULTADOS].sheet_state == "visible"
    # AJUSTES-XLS-UX pos-174: camada tecnica oculta por padrao (hidden normal,
    # reexibivel; nunca veryHidden).
    assert wb[ABA_RESULTADOS_DETALHE].sheet_state == "hidden"
    assert wb["MEMORIA_RESULTADOS"].sheet_state == "hidden"
    assert wb[ABA_RESULTADOS]["B2"].value == TITULO_RESULTADOS_EXECUTIVO
    assert wb[ABA_RESULTADOS_DETALHE]["A1"].value == TITULO_RESULTADOS_DETALHE


def test_nenhum_nome_definido_aponta_para_a_aba_executiva():
    nomes = _template().defined_names
    for nome, dn in nomes.items():
        assert not str(dn.attr_text).startswith("RESULTADOS!"), nome
    assert nomes["STATUS_RESULTADOS"].attr_text == "RESULTADOS_DETALHE!$B$3"
    assert nomes["OPCOES_APLICAR_MANUAL"].attr_text == "RESULTADOS_DETALHE!$J$2:$J$3"
    assert nomes["VTA_FINAL"].attr_text == "MEMORIA_RESULTADOS!$B$26"


def test_ajustes_manuais_continuam_entradas_lidas_pela_memoria():
    """O acoplamento XLS->XLS de C43:G50 seguiu a aba renomeada."""
    wb = _template()
    mem = wb["MEMORIA_RESULTADOS"]
    assert "RESULTADOS_DETALHE!$C$43" in str(mem["B5"].value)
    assert "RESULTADOS_DETALHE!$G$43" in str(mem["B5"].value)
    assert "RESULTADOS_DETALHE!$C$44" in str(mem["B24"].value)
    assert "RESULTADOS_DETALHE!$C$45" in str(mem["B25"].value)
    assert "RESULTADOS_DETALHE!$H$5" in str(wb["comparativo_VTA"]["B208"].value)
    dvs = {str(dv.sqref): dv for dv in wb[ABA_RESULTADOS_DETALHE].data_validations.dataValidation}
    assert dvs["G43:G50"].formula1 == "OPCOES_APLICAR_MANUAL"
    assert "C46:C50" in dvs


def test_nenhuma_formula_do_workbook_aponta_para_a_aba_executiva():
    """Somente a propria aba executiva pode ser lida como RESULTADOS!."""
    padrao = re.compile(r"(?<![A-Za-z_])RESULTADOS!")
    for ws in _template().worksheets:
        if ws.title == ABA_RESULTADOS:
            continue
        for linha in ws.iter_rows():
            for c in linha:
                if isinstance(c.value, str) and c.value.startswith("="):
                    assert not padrao.search(c.value), f"{ws.title}!{c.coordinate}"


def test_aba_tecnica_resolvida_por_versao():
    assert aba_resultados_tecnica(["X", "RESULTADOS_DETALHE", "RESULTADOS"]) == (
        "RESULTADOS_DETALHE"
    )
    assert aba_resultados_tecnica(["X", "MEMORIA_RESULTADOS", "RESULTADOS"]) == "RESULTADOS"
    assert aba_resultados_tecnica(["X"]) is None


def test_gerado_preserva_aba_executiva_no_roundtrip_openpyxl():
    """A Coleta entregue (openpyxl load/save) mantem formulas, mesclas,
    formatacao condicional e a indicacao dos ajustes manuais (texto normal:
    com a camada tecnica oculta, o hiperlink deixou de existir)."""
    ws = _gerado()[ABA_RESULTADOS]
    for endereco, formula in v3.formulas_executivo().items():
        assert ws[endereco].value == formula, endereco
    assert len(ws.merged_cells.ranges) >= 30
    assert len(ws.conditional_formatting) >= 10
    assert ws["D31"].hyperlink is None
    assert "RESULTADOS_DETALHE" in ws["D31"].value and "Reexibir" in ws["D31"].value


def test_formatacao_condicional_executiva_sobrevive_a_geracao():
    """Regra condicional que referencia OUTRA aba e gravada pelo Excel na
    extensao x14, que o openpyxl DESCARTA ao gerar a Coleta (o quadro 3 e o
    card de situacao perdiam a cor). Toda regra da aba executiva tem de ler a
    propria aba ou um nome definido."""
    def _faixas(ws):
        return sorted(str(f.sqref) for f in ws.conditional_formatting)

    for ws in (_template()[ABA_RESULTADOS], _gerado()[ABA_RESULTADOS]):
        regras = [
            (str(faixa.sqref), formula)
            for faixa in ws.conditional_formatting
            for regra in faixa.rules
            for formula in (regra.formula or [])
        ]
        assert len(regras) >= 30
        for faixa, formula in regras:
            assert "!" not in formula, f"{faixa}: {formula}"
    assert _faixas(_template()[ABA_RESULTADOS]) == _faixas(_gerado()[ABA_RESULTADOS])
    # Regras sobrepostas (bordas x ambar) so convivem sem "parar se verdadeiro".
    for ws in (_template()[ABA_RESULTADOS], _gerado()[ABA_RESULTADOS]):
        for faixa in ws.conditional_formatting:
            for regra in faixa.rules:
                assert not regra.stopIfTrue, f"{faixa.sqref}: stopIfTrue"
    # F53:G53 e mesclada: o Excel grava a faixa como B53:F53.
    assert "B53:F53" in _faixas(_template()[ABA_RESULTADOS])
    assert "B54:F57" in _faixas(_template()[ABA_RESULTADOS])
    assert "G8:G10" in _faixas(_template()[ABA_RESULTADOS])


def test_seguranca_aceita_a_estrutura_nova_com_17_abas():
    from _seguranca_xlsx import ABAS_PERMITIDAS, MAX_ABAS_WORKBOOK

    nomes = _gerado().sheetnames
    assert len(nomes) == 17 <= MAX_ABAS_WORKBOOK
    assert set(nomes) <= ABAS_PERMITIDAS


def test_versionamento_e_linhagem():
    from _compatibilidade_coleta import LINHAGEM_COLETA_11, detectar_linhagem_coleta
    from _versao import COLETA_VERSION, COLETA_VERSOES_ACEITAS

    # Coleta 11.3 (PR #175, UX) mantem a arquitetura executiva + detalhe da 11.2.
    assert COLETA_VERSION == "11.7"
    assert {"11.0", "11.1", "11.2", "11.3", "11.4", "11.5", "11.6"} <= set(COLETA_VERSOES_ACEITAS)
    deteccao = detectar_linhagem_coleta(_gerado())
    assert deteccao["codigo"] == LINHAGEM_COLETA_11
    assert deteccao["marcador_publico"] == "11.7"


def _como_arquivo_anterior(marcador: str | None = "11.1"):
    """Coleta anterior simulada: uma unica RESULTADOS tecnica, sem executiva.

    O marcador publico (CONTROLE!B24) acompanha a versao simulada: 11.0/11.1,
    ou nenhum (linhagem PRE_11, sem versionamento publico)."""
    wb = load_workbook(io.BytesIO(_bytes(_gerado())))
    del wb[ABA_RESULTADOS]
    wb[ABA_RESULTADOS_DETALHE].title = ABA_RESULTADOS
    # Num 11.0/11.1 real a unica RESULTADOS (tecnica) e visivel; a 11.2+ nasce
    # com a camada tecnica oculta, entao a simulacao a reexibe.
    wb[ABA_RESULTADOS].sheet_state = "visible"
    wb[ABA_RESULTADOS]["A1"] = "RESULTADOS CONSOLIDADOS — REAJUSTE CONTRATUAL"
    wb["CONTROLE"]["B24"] = marcador
    # `.title` do openpyxl nao reescreve nomes nem formulas (o Excel reescreve):
    # num 11.x anterior eles apontam para RESULTADOS!.
    for dn in wb.defined_names.values():
        if dn.attr_text and f"{ABA_RESULTADOS_DETALHE}!" in dn.attr_text:
            dn.attr_text = dn.attr_text.replace(
                f"{ABA_RESULTADOS_DETALHE}!", f"{ABA_RESULTADOS}!"
            )
    for ws in wb.worksheets:
        for linha in ws.iter_rows():
            for c in linha:
                if isinstance(c.value, str) and f"{ABA_RESULTADOS_DETALHE}!" in c.value:
                    c.value = c.value.replace(
                        f"{ABA_RESULTADOS_DETALHE}!", f"{ABA_RESULTADOS}!"
                    )
    return wb


def _como_112_sem_detalhe():
    """Coleta 11.2 mutilada: RESULTADOS_DETALHE apagada pelo openpyxl.

    O openpyxl nao reescreve formulas nem nomes definidos ao apagar a aba:
    ficam referencias TEXTUAIS pendentes ('RESULTADOS_DETALHE!$B$3'), sem
    nenhum #REF!, e a unica RESULTADOS restante e a pagina executiva."""
    wb = load_workbook(io.BytesIO(_bytes(_gerado())))
    del wb[ABA_RESULTADOS_DETALHE]
    return load_workbook(io.BytesIO(_bytes(wb)))


def _bloqueios_upload(wb) -> list[str]:
    from _coleta_reajuste import ler_coleta_reajuste

    return ler_coleta_reajuste(_bytes(wb))["bloqueios_estruturais"]


def test_p1_coleta_112_integra_e_aceita():
    from _coleta_reajuste import _validar_resultados_integra
    from _compatibilidade_coleta import exige_resultados_detalhe
    from _leitor_masterfile_v10 import ler_masterfile_v10

    wb = _gerado()
    assert exige_resultados_detalhe(wb)
    assert aba_resultados_tecnica(wb) == ABA_RESULTADOS_DETALHE
    assert _validar_resultados_integra(wb, "teste")["visivel"]
    assert not any(ABA_RESULTADOS_DETALHE in b for b in _bloqueios_upload(wb))
    lido = ler_masterfile_v10(_bytes(wb), exigir_modelo_oficial=True)
    assert ABA_RESULTADOS_DETALHE not in (lido.get("abas_ausentes") or [])


def test_p1_coleta_112_sem_detalhe_e_rejeitada():
    from _coleta_reajuste import _validar_resultados_integra

    wb = _como_112_sem_detalhe()
    assert ABA_RESULTADOS in wb.sheetnames
    # Sem fallback para a RESULTADOS executiva.
    assert aba_resultados_tecnica(wb) is None
    with pytest.raises(ValueError, match=ABA_RESULTADOS_DETALHE):
        _validar_resultados_integra(wb, "teste")
    assert any(ABA_RESULTADOS_DETALHE in b for b in _bloqueios_upload(wb))


@pytest.mark.parametrize("marcador", ["11.0", "11.1", None])
def test_p1_versao_anterior_sem_detalhe_continua_aceita(marcador):
    from _coleta_reajuste import _validar_resultados_integra
    from _compatibilidade_coleta import exige_resultados_detalhe
    from _leitor_masterfile_v10 import ler_masterfile_v10

    wb = _como_arquivo_anterior(marcador)
    assert ABA_RESULTADOS_DETALHE not in wb.sheetnames
    assert not exige_resultados_detalhe(wb)
    assert aba_resultados_tecnica(wb) == ABA_RESULTADOS
    assert _validar_resultados_integra(wb, "teste")["visivel"]
    assert not any(ABA_RESULTADOS_DETALHE in b for b in _bloqueios_upload(wb))
    lido = ler_masterfile_v10(_bytes(wb), exigir_modelo_oficial=True)
    assert ABA_RESULTADOS_DETALHE not in (lido.get("abas_ausentes") or [])


def test_p1_coleta_112_mutilada_com_referencias_pendentes_sem_ref_e_rejeitada():
    from _leitor_masterfile_v10 import ler_masterfile_v10

    wb = _como_112_sem_detalhe()
    textos = [str(dn.attr_text) for dn in wb.defined_names.values()]
    textos += [
        c.value
        for ws in wb.worksheets
        for linha in ws.iter_rows()
        for c in linha
        if isinstance(c.value, str) and c.value.startswith("=")
    ]
    # Pre-condicao do cenario do P1: referencias pendentes, nenhum #REF!.
    assert any(f"{ABA_RESULTADOS_DETALHE}!" in t for t in textos)
    assert not any("#REF!" in t for t in textos)

    lido = ler_masterfile_v10(_bytes(wb), exigir_modelo_oficial=True)
    assert ABA_RESULTADOS_DETALHE in (lido.get("abas_ausentes") or [])
    assert ABA_RESULTADOS_DETALHE in (lido.get("erro") or "")
    assert any(ABA_RESULTADOS_DETALHE in b for b in _bloqueios_upload(wb))


def test_validacao_estrutural_do_upload_aceita_arquivo_novo_e_antigo():
    from _coleta_reajuste import _validar_resultados_integra

    novo = _validar_resultados_integra(_gerado(), "teste")
    assert novo["visivel"] and novo["formulas"] >= 40
    assert _validar_resultados_integra(_como_arquivo_anterior(), "teste")["visivel"]


def test_validacao_estrutural_rejeita_executiva_substituida():
    from _coleta_reajuste import _validar_resultados_integra

    wb = load_workbook(io.BytesIO(_bytes(_gerado())))
    wb[ABA_RESULTADOS]["B2"] = "outra coisa"
    with pytest.raises(ValueError, match="executiva"):
        _validar_resultados_integra(wb, "teste")


def test_referencias_vta_leem_a_camada_tecnica_e_nao_a_executiva():
    """Sem os nomes de auditoria, o fallback por coordenada le B10:H13 da aba
    TECNICA — nunca as coordenadas da pagina executiva."""
    from _leitor_masterfile_v10 import _ler_referencias_vta

    for wb in (load_workbook(io.BytesIO(_bytes(_gerado()))), _como_arquivo_anterior()):
        for nome in list(wb.defined_names):
            if nome.startswith("AUDITORIA_"):
                del wb.defined_names[nome]
        tecnica = wb[aba_resultados_tecnica(wb)]
        tecnica["B10"] = 111.11
        tecnica["H13"] = "RECONCILIADO"
        out = _ler_referencias_vta(wb)
        assert out["origem_leitura"] == "legacy_coordinates"
        assert out["forma1_posicao_atual"] == 111.11
        assert out["reconciliacao_status"] == "RECONCILIADO"


# =========================================================================== #
# C. CONTRATO VISUAL
# =========================================================================== #
def test_formulas_ascii_e_balanceadas():
    v3.validar_ascii_e_parenteses(v3.formulas_executivo())


def test_percentuais_datas_e_moeda_sao_numeros_formatados():
    ws = _template()[ABA_RESULTADOS]
    for endereco in ("F6", "F9", "E35", "E36", "E37", "E38", "E39"):
        assert ws[endereco].number_format.startswith("0.00%"), endereco
    for endereco in ("D6", "C35", "D39", "F37", "C28"):
        assert ws[endereco].number_format.startswith("dd/mm/yyyy"), endereco
    for endereco in ("B9", "C9", "D9", "E9", "C14", "C17", "C43", "E49", "C54", "C64"):
        fmt = ws[endereco].number_format
        assert "R$" in fmt and "#,##0.00" in fmt, (endereco, fmt)
    for formula in v3.formulas_executivo().values():
        assert "TEXT(" not in formula.upper()


def test_mesclagens_simples_sem_sobreposicao():
    ws = _template()[ABA_RESULTADOS]
    celulas: set[tuple[int, int]] = set()
    for faixa in ws.merged_cells.ranges:
        assert faixa.min_row == faixa.max_row, f"mesclagem multi-linha {faixa}"
        for linha in range(faixa.min_row, faixa.max_row + 1):
            for coluna in range(faixa.min_col, faixa.max_col + 1):
                assert (linha, coluna) not in celulas, f"sobreposicao em {faixa}"
                celulas.add((linha, coluna))


def test_larguras_e_cards_comportam_valores_altos():
    from openpyxl.utils import column_index_from_string

    ws = _template()[ABA_RESULTADOS]
    # O XLSX agrupa colunas de mesma largura num unico <col min..max>; acessar
    # column_dimensions["D"] criaria uma entrada nova com a largura padrao.
    grupos = [(d.min, d.max, d.width) for d in ws.column_dimensions.values() if d.min]
    for coluna, largura in v3.LARGURAS.items():
        idx = column_index_from_string(coluna)
        obtida = next(w for mn, mx, w in grupos if mn <= idx <= mx)
        assert abs(obtida - largura) < 0.75, (coluna, obtida)
    # "R$ 999.999.999,99" = 17 caracteres a 15pt negrito: >= 25 de largura.
    assert min(v3.LARGURAS[c] for c in "BCDEF") >= 26


def test_secao_detalhes_do_metodo_some_fora_de_pcs():
    f = v3.formulas_executivo()
    for linha in range(62, 72):
        for col in "BCD":
            chave = f"{col}{linha}"
            if chave in f:
                assert f[chave].startswith('=IF(METODO_RETROATIVO="PCs"'), chave


# =========================================================================== #
# Excel REAL (opt-in): ### a 100%, encaixe, reparo.
# =========================================================================== #
@pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="defina RUN_EXCEL_INTEGRATION=1 para executar Excel COM",
)
def test_excel_real_sem_hashes_sem_reparo(tmp_path):
    from _coleta_oficial import obter_coleta_oficial_bytes
    from validar_visual_resultados_v3 import validar

    arquivo = tmp_path / "coleta_branco.xlsx"
    arquivo.write_bytes(obter_coleta_oficial_bytes())
    assert validar(arquivo, tmp_path / "saida", estresse=True) == []
