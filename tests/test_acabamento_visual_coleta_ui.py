"""Etapa UI — acabamento visual da Coleta e destaque do seletor de ciclo.

Protege exatamente o que a etapa alterou, sem ampliar escopo:

  * CONTROLE!B10 declara a ausencia de historico na analise de UM ciclo e o
    multiciclo continua com a formula do template;
  * CONTROLE!B11 (fator historico integral) permanece formula nos dois casos —
    ela alimenta comparativo_VTA!B208 e RESULTADOS!H5/H8;
  * nova semantica de parametros!G10, sem terceira opcao e sem mexer nas
    validacoes Sim/Nao;
  * padronizacao de parametros!G2:G6 e limpeza de A17:G20;
  * acabamento do painel posicao_referencia!H:I encerrando em H11;
  * CICLO_EM_EXECUCAO com laranja no campo de preenchimento (C) e verde no
    derivado (D);
  * legibilidade da coluna C de cobertura_temporal;
  * RESULTADOS!H8 legivel, A8 horizontal sem mesclagem e bordas da tabela 5;
  * preservacao de TODAS as formulas do template fora de CONTROLE!B10;
  * CSS do seletor de ciclo restrito a classe do proprio widget.
"""
from __future__ import annotations

import io
import re
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from _coleta_oficial import gerar_coleta_oficial_preenchida, obter_coleta_oficial_bytes
from _gerador_masterfile import AVISO_HISTORICO_CICLO_UNICO

RAIZ = Path(__file__).resolve().parents[1]
PAGINA_CICLO_UNICO = (RAIZ / "pages" / "01_Calculo_Simples.py").read_text(
    encoding="utf-8"
)

CICLO_C3 = {
    "ciclo": "C3",
    "data_inicio": date(2027, 1, 1),
    "financeiro_inicio": date(2027, 3, 1),
    "percentual": 0.0308,
    "objeto_analise_atual": True,
    "situacao": "Aplicado nesta apuracao",
}
DADOS_UNICO = {
    "data_base_original": "01/01/2024",
    "indice": "IPCA",
    "ciclos": [dict(CICLO_C3)],
}
DADOS_MULTI = {
    "data_base_original": "01/01/2024",
    "indice": "IPCA",
    "ciclos": [
        {"ciclo": "C1", "data_inicio": date(2025, 1, 1),
         "financeiro_inicio": date(2025, 3, 1), "percentual": 0.0449,
         "objeto_analise_atual": True, "situacao": "Aplicado nesta apuracao"},
        {"ciclo": "C2", "financeiro_inicio": date(2026, 3, 1),
         "percentual": 0.0316, "objeto_analise_atual": True,
         "situacao": "Aplicado nesta apuracao"},
        dict(CICLO_C3),
    ],
}


def _formulas(wb) -> dict[str, str]:
    return {
        f"{ws.title}!{cell.coordinate}": cell.value
        for ws in wb.worksheets
        for row in ws.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    }


def _estilo(borda, lado):
    return getattr(getattr(borda, lado, None), "style", None)


def _abrir(dados):
    return load_workbook(io.BytesIO(gerar_coleta_oficial_preenchida(dados)))


@pytest.fixture(scope="module")
def em_branco():
    wb = load_workbook(io.BytesIO(obter_coleta_oficial_bytes()))
    yield wb
    wb.close()


@pytest.fixture(scope="module")
def unico():
    wb = _abrir(DADOS_UNICO)
    yield wb
    wb.close()


@pytest.fixture(scope="module")
def multiciclo():
    wb = _abrir(DADOS_MULTI)
    yield wb
    wb.close()


# ---------------------------------------------------------------------------
# 1. Interface web
# ---------------------------------------------------------------------------

def test_destaque_do_seletor_de_ciclo_e_restrito_ao_proprio_campo():
    assert '"Qual ciclo deseja analisar?"' in PAGINA_CICLO_UNICO
    assert 'key="sim_ciclo_analise"' in PAGINA_CICLO_UNICO
    # O seletor recebe destaque proprio via a classe que o Streamlit publica
    # para widgets com key — nunca por um seletor global de selectbox.
    assert '.st-key-sim_ciclo_analise [data-baseweb="select"]' in PAGINA_CICLO_UNICO
    blocos = [
        bloco
        for bloco in re.findall(r"<style>(.*?)</style>", PAGINA_CICLO_UNICO, re.S)
        if "st-key-sim_ciclo_analise" in bloco
    ]
    assert len(blocos) == 1
    seletores = [
        regra.split("{", 1)[0].strip()
        for regra in blocos[0].split("}")
        if regra.split("{", 1)[0].strip()
    ]
    assert seletores
    for seletor in seletores:
        assert seletor.startswith(".st-key-sim_ciclo_analise"), seletor


# ---------------------------------------------------------------------------
# 2. CONTROLE!B10 / B11
# ---------------------------------------------------------------------------

def test_ciclo_unico_declara_ausencia_de_historico_em_b10(unico):
    ctl = unico["CONTROLE"]
    assert ctl["B10"].value == AVISO_HISTORICO_CICLO_UNICO
    assert (
        AVISO_HISTORICO_CICLO_UNICO
        == "Histórico anterior não incluído nesta análise."
    )
    assert "N/A" not in str(ctl["B10"].value)
    assert ctl["B10"].number_format == "@"


def test_ciclo_unico_reapresenta_a_variacao_apurada_em_c11(unico):
    """A variacao vem do resultado ja apurado; C11 e nota, nao fonte de verdade."""
    ctl = unico["CONTROLE"]
    assert ctl["C11"].value == "Variação apurada do ciclo único (C3): 3,08%"
    # Mesmo numero que o gerador gravou em parametros!E5 (percentual do C3).
    assert unico["parametros"]["E5"].value == pytest.approx(0.0308)
    # C11 e a coluna de notas da propria aba (ver C3) e nao entra em formula.
    assert not str(ctl["C11"].value).startswith("=")


def test_multiciclo_mantem_a_formula_de_b10_e_nao_recebe_nota(multiciclo, em_branco):
    assert multiciclo["CONTROLE"]["B10"].value == em_branco["CONTROLE"]["B10"].value
    assert str(multiciclo["CONTROLE"]["B10"].value).startswith("=")
    assert multiciclo["CONTROLE"]["C11"].value is None


@pytest.mark.parametrize("fixture", ["unico", "multiciclo"])
def test_b11_permanece_formula_pois_alimenta_o_motor(fixture, request, em_branco):
    """B11 e consumida por comparativo_VTA!B208 e RESULTADOS!H5/H8."""
    wb = request.getfixturevalue(fixture)
    assert wb["CONTROLE"]["B11"].value == em_branco["CONTROLE"]["B11"].value
    assert wb["comparativo_VTA"]["B208"].value == (
        em_branco["comparativo_VTA"]["B208"].value
    )


def test_nenhuma_outra_formula_do_template_muda(unico, multiciclo, em_branco):
    base = _formulas(em_branco)
    esperado = {k: v for k, v in base.items() if k != "CONTROLE!B10"}
    for wb in (unico, multiciclo):
        atual = _formulas(wb)
        atual.pop("CONTROLE!B10", None)
        assert atual == esperado


# ---------------------------------------------------------------------------
# 3. parametros
# ---------------------------------------------------------------------------

def test_situacao_g2_g6_segue_o_padrao_da_tabela(em_branco):
    par = em_branco["parametros"]
    for linha in range(2, 7):
        alvo, referencia = par.cell(linha, 7), par.cell(linha, 6)
        assert alvo.font.b == referencia.font.b is False
        assert alvo.font.sz == referencia.font.sz
        assert alvo.font.color.rgb == referencia.font.color.rgb == "FF595959"
        assert alvo.fill.fgColor.rgb == referencia.fill.fgColor.rgb == "FFEDEDED"
        assert _estilo(alvo.border, "left") == "thin"


def test_faixa_a17_g20_fica_visualmente_limpa(em_branco):
    par = em_branco["parametros"]
    for linha in range(17, 21):
        for coluna in range(1, 8):
            celula = par.cell(linha, coluna)
            assert celula.value is None
            assert not any(
                _estilo(celula.border, lado)
                for lado in ("left", "right", "top", "bottom")
            ), celula.coordinate
    # A tabela MEMORIA DO FATOR continua fechada na linha 16.
    assert _estilo(par["A16"].border, "bottom") == "thin"
    assert par.sheet_view.showGridLines is False


def test_g10_pergunta_pela_existencia_do_reajuste_anterior(em_branco):
    par = em_branco["parametros"]
    assert par["G10"].value == (
        "EXISTE REAJUSTE ANTERIOR FORMALIZADO? (Sim/Não; vazio=não comprovado)"
    )
    validacoes = {
        str(dv.sqref): dv.formula1 for dv in par.data_validations.dataValidation
    }
    assert validacoes["G13:G15"] == "OPCOES_SIM_NAO"
    assert validacoes["G12"] == "OPCOES_SIM_NAO_NA"
    # Sem terceira opcao: a lista de origem segue Sim/Nao (+N/A ja existente).
    assert [par.cell(linha, 20).value for linha in range(2, 5)] == ["Sim", "Nao", "N/A"]


# ---------------------------------------------------------------------------
# 4. posicao_referencia
# ---------------------------------------------------------------------------

def test_painel_de_marcos_termina_em_h11(em_branco):
    pos = em_branco["posicao_referencia"]
    modelo_rotulo, modelo_valor = pos["H2"], pos["I2"]
    for linha in range(1, 12):
        assert pos.cell(linha, 8).fill.fgColor.rgb == modelo_rotulo.fill.fgColor.rgb
        assert pos.cell(linha, 9).fill.fgColor.rgb == modelo_valor.fill.fgColor.rgb
        assert _estilo(pos.cell(linha, 8).border, "left") == "thin"
        assert _estilo(pos.cell(linha, 9).border, "right") == "thin"
    assert _estilo(pos["H8"].border, "bottom") is None
    assert _estilo(pos["H11"].border, "bottom") == "thin"
    assert _estilo(pos["I11"].border, "bottom") == "thin"
    # Nada foi acrescentado abaixo do painel.
    assert pos["H12"].value is None
    assert pos["I9"].number_format == "mm-dd-yy"   # data preservada


# ---------------------------------------------------------------------------
# 5. CICLO_EM_EXECUCAO
# ---------------------------------------------------------------------------

def test_coluna_c_e_preenchimento_e_coluna_d_e_automatica(unico):
    ws = unico["CICLO_EM_EXECUCAO"]
    for linha in (13, 100, 211):
        entrada, derivada = ws.cell(linha, 3), ws.cell(linha, 4)
        assert entrada.fill.fgColor.rgb == "FFFCE4D6"       # laranja claro
        assert _estilo(entrada.border, "left") == "medium"
        assert entrada.border.left.color.rgb == "FFE67E22"  # laranja
        assert entrada.protection.locked is False
        assert derivada.fill.fgColor.rgb == "FFE8F5F1"      # verde suave
        assert derivada.value.startswith("=")
    assert ws["C12"].fill.fgColor.rgb == "FFE67E22"
    assert ws["D12"].fill.fgColor.rgb == "FF0F5B50"
    # A area vizinha nao foi repintada.
    assert ws["A13"].fill.fgColor.rgb == "FFF2F2F2"
    assert ws["E13"].fill.fgColor.rgb == "FFDDEBF7"


# ---------------------------------------------------------------------------
# 6. cobertura_temporal
# ---------------------------------------------------------------------------

def test_notas_da_coluna_c_cabem_na_propria_celula(em_branco):
    cob = em_branco["cobertura_temporal"]
    assert 55.0 < cob.column_dimensions["C"].width <= 70.0
    for linha in (8, 14, 17):
        celula = cob.cell(linha, 3)
        assert celula.alignment.wrap_text is True
        linhas = max(1, -(-len(str(celula.value)) // 60))
        assert cob.row_dimensions[linha].height >= linhas * 14.5


# ---------------------------------------------------------------------------
# 7. RESULTADOS
# ---------------------------------------------------------------------------

def test_status_da_tabela1_fica_legivel(em_branco):
    res = em_branco["RESULTADOS"]
    modelo = res["H14"].font
    assert res["H8"].font.color.rgb == modelo.color.rgb == "FF595959"
    assert (res["H8"].font.name, res["H8"].font.sz, res["H8"].font.b) == (
        modelo.name,
        modelo.sz,
        modelo.b,
    )
    assert str(res["H8"].value).startswith("=IF(OR(VTA_FINAL")


def test_titulo_da_tabela1_ocupa_a8_d8_sem_mesclagem(em_branco):
    res = em_branco["RESULTADOS"]
    assert res["A8"].alignment.wrap_text in (False, None)
    assert [res[c].value for c in ("B8", "C8", "D8")] == [None, None, None]
    assert not [m for m in res.merged_cells.ranges if m.min_row <= 8 <= m.max_row]
    assert res["A8"].value == "1. VALOR TOTAL DO CONTRATO — TRES REFERENCIAS DO VTA"


def test_tabela_da_linha_53_tem_bordas(em_branco):
    res = em_branco["RESULTADOS"]
    assert str(res["A53"].value).startswith("5. TOTAIS CANONICOS DE PCs")
    for linha in range(54, 67):
        for coluna in range(1, 4):
            celula = res.cell(linha, coluna)
            for lado in ("left", "right", "top", "bottom"):
                assert _estilo(celula.border, lado) == "thin", celula.coordinate
            assert celula.border.left.color.rgb == "FFB0C4D8"
    assert res.sheet_view.showGridLines is False
