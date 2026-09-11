from copy import deepcopy
from io import BytesIO
from pathlib import Path

import pytest
from docx import Document

from _apresentacao_pc import montar_quadros_pc, SALDO, NOTA_EXECUCAO
from _motor_composicao_vta import montar_composicao_vta
from test_vta_pc_composicao import _leitura


def test_quadros_conciliam_sem_mutar_apuracao_e_sem_somar_reconhecido():
    comp = montar_composicao_vta(_leitura())
    antes = deepcopy(comp)
    q = montar_quadros_pc(comp)
    assert q
    assert comp == antes
    assert sum(l[1] for l in q["composicao"][:3]) == pytest.approx(comp["vta_composicao"])
    assert len(q["composicao"]) == 4
    for linha in q["execucao"]:
        assert linha[1] + linha[3] == pytest.approx(linha[2])


def test_posicao_fisica_nao_e_rotulada_como_original():
    comp = montar_composicao_vta(_leitura())
    comp["execucao_por_ciclo"][0]["fonte"] = "posicao_fisica"
    q = montar_quadros_pc(comp)
    assert q["execucao_cabecalho"] == ["Ciclo", "Valor executado atualizado", "Situação"]
    assert all(len(linha) == 3 for linha in q["execucao"])
    assert q["composicao"][0][2] == NOTA_EXECUCAO


def test_conferencia_nao_corrige_vta_e_nao_exibe_outros_metodos():
    comp = montar_composicao_vta(_leitura())
    assert not montar_quadros_pc(comp, vta=comp["vta_composicao"] + 1)
    comp["metodo"] = "financeiro"
    assert not montar_quadros_pc(comp)


def test_referencias_preservadas_sem_trocar_pelo_saldo_final():
    comp = montar_composicao_vta(_leitura())
    refs = [["C0", 1000., 1000., 0.], ["C1", 800., 880., 80.]]
    q = montar_quadros_pc(comp, referencias_xls=refs)
    assert q["referencias"] == refs
    assert q["saldo_final"] == comp["saldo_remanescente"]["valor_atualizado"]
    assert all("referência do ciclo" in c for c in q["referencia_cabecalho"][1:])


def test_caso_conhecido_runtime_web_documentos():
    from _coleta_reajuste_documentos import processar_coleta_oficial_runtime
    from _sumario_executivo import montar_dados_sumario_executivo
    from _templates_documentos import gerar_termo_apostila, gerar_despacho_saneador
    entrada = Path(
        r"C:\_DesktopReal\PC_Tabelas_Execucao_Remanescente_Evidencias"
        r"\CASO_PC_VALIDACAO_FINAL.xlsx"
    )
    if not entrada.exists():
        pytest.skip("caso conhecido externo ausente")
    r, _ = processar_coleta_oficial_runtime(entrada.read_bytes())
    comp_antes = deepcopy(r["composicao_vta"])
    dados = montar_dados_sumario_executivo(r)
    q = dados["quadros_pc"]
    assert [l[1] for l in q["composicao"]] == [775., 1870., 30., 2675.]
    assert dados["sintese"]["vta_saldo_remanescente_atualizado"] == 1870.
    assert q["saldo_final"] != r["remanescente_reajustado"]
    for gerar in (gerar_termo_apostila, gerar_despacho_saneador):
        doc = Document(BytesIO(gerar(r)))
        tabelas = [[ [c.text for c in linha.cells] for linha in t.rows] for t in doc.tables]
        quadro = next(t for t in tabelas if t[0] == q["composicao_cabecalho"])
        assert [l[1] for l in quadro[1:]] == ["R$ 775,00", "R$ 1.870,00", "R$ 30,00", "R$ 2.675,00"]
        assert quadro[2][0] == SALDO
    assert r["composicao_vta"] == comp_antes
