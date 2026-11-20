"""Adequacao Orcamentaria — projecao futura limitada ao saldo (Cl8us 12.20).

Cenarios pontuais desta entrega:
  1. saldo elevado + historico mensal regular: projecao preservada;
  2. contrato praticamente esgotado com varios meses restantes: cap no saldo;
  3. historico irregular com saldo: futuro nao determinado (sem recorrencia);
  4. historico irregular sem saldo confiavel: futuro nao determinado;
  5. retroativo reconhecido com projecao futura indisponivel: total parcial;
  6. retroativo potencial separado (fora da complementacao e da programacao);
  7. valores manuais acima do saldo: nao reduzidos, excesso exige correcao;
  8. WEB = XLSX (o arquivo baixado na propria pagina).
"""
from __future__ import annotations

import re
import sys
from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from _adequacao_orcamentaria import (  # noqa: E402
    SITUACAO_EXCEDE_SALDO,
    SITUACAO_LIMITADO_SALDO,
    base_prudencial_irregular,
    limitar_projecao_ao_saldo,
)
from _adequacao_ui import (  # noqa: E402
    NAO_DETERMINADA,
    calcular_projecao,
    consolidar_resultado_adequacao,
    cronograma_por_exercicio,
    linhas_resumo_resultado,
    montar_base_editor,
    parse_moeda_br,
    saldo_elegivel_da_apuracao,
)
from test_adequacao_pc_janela_potencial import (  # noqa: E402
    POTENCIAL,
    RECONHECIDO,
    _blob,
    _run,
    _sessao_pc,
    _valor_do_card,
)

FATOR = 1.05
PERIODOS = list(pd.period_range("2026-09", "2027-08", freq="M"))  # 12 meses
ESTIMADA = "COMPLEMENTAÇÃO ORÇAMENTÁRIA ESTIMADA"
PARCIAL = "COMPLEMENTAÇÃO ORÇAMENTÁRIA PARCIAL"


def _saldo(valor):
    return {"valor": valor, "atualizado": None, "origem": "teste"}


def _texto_do_card(blob, rotulo):
    idx = blob.index(rotulo + "</div>")
    return re.sub(r"<[^>]+>", " ", blob[idx:idx + 400])


# ------------------------------------------------------------ 1. saldo elevado

def test_1_saldo_elevado_e_historico_regular_preservam_a_projecao():
    editor = montar_base_editor(PERIODOS, 10000.0)
    sem = calcular_projecao(editor, 10000.0, FATOR)
    com = calcular_projecao(editor, 10000.0, FATOR, saldo_elegivel=1_000_000.0)
    assert list(com["Diferença futura a adequar"]) == list(sem["Diferença futura a adequar"])
    assert set(com["Limite do saldo"]) == {"Dentro do saldo"}
    cons = consolidar_resultado_adequacao(retroativo=1000.0, df_projecao=com,
                                          meses_restantes=12, saldo_info=_saldo(1_000_000.0))
    assert cons["diferenca_futura"] == pytest.approx(6000.0)
    assert cons["diferenca_media_mensal"] == pytest.approx(500.0)
    assert cons["fatura_media_reajustada"] == pytest.approx(10500.0)
    assert cons["complementacao"] == pytest.approx(7000.0)
    assert cons["ressalvas"] == []


# ------------------------------------------------- 2. contrato quase esgotado

def test_2_contrato_quase_esgotado_limita_a_projecao_ao_saldo():
    proj = calcular_projecao(montar_base_editor(PERIODOS, 10000.0), 10000.0, FATOR,
                             saldo_elegivel=15000.0)
    bases = list(proj["Valor base considerado"])
    assert bases[:3] == [10000.0, 5000.0, 0.0]
    assert sum(bases) == pytest.approx(15000.0)  # nunca supera o saldo
    assert proj.loc[1, "Limite do saldo"] == SITUACAO_LIMITADO_SALDO
    assert proj.loc[11, "Saldo elegível após a competência"] == 0.0
    cons = consolidar_resultado_adequacao(retroativo=2000.0, df_projecao=proj,
                                          meses_restantes=12, saldo_info=_saldo(15000.0))
    assert cons["diferenca_futura"] == pytest.approx(750.0)
    assert cons["diferenca_media_mensal"] == pytest.approx(62.5)
    assert cons["fatura_media_reajustada"] is None  # saldo insuficiente: omitida
    assert cons["complementacao"] == pytest.approx(2750.0)  # retroativo nao limitado
    assert any("limitada ao saldo" in r for r in cons["ressalvas"])


# ------------------------------------------------ 3. irregular com saldo

def test_3_irregular_com_saldo_usa_base_prudencial_limitada():
    assert base_prudencial_irregular(2000.0, 12, 1_000_000.0) == pytest.approx(2000.0)
    assert base_prudencial_irregular(2000.0, 12, 6000.0) == pytest.approx(500.0)
    assert base_prudencial_irregular(2000.0, 12, None) is None


def test_3_pagina_irregular_com_saldo_baixo_nao_inventa_mensalidade():
    at = _run(_sessao_pc(saldo=5000.0))
    assert not at.exception, at.exception
    blob = _blob(at)
    assert NAO_DETERMINADA in _texto_do_card(blob, "Diferença futura total projetada")
    assert NAO_DETERMINADA in _texto_do_card(blob, "Diferença média mensal projetada")
    assert PARCIAL in blob and ESTIMADA not in blob
    assert _valor_do_card(blob, PARCIAL) == pytest.approx(RECONHECIDO)
    avisos = " ".join(w.value for w in at.warning)
    assert "não presumirá uma mensalidade" in avisos
    assert "A projeção automática foi limitada ao saldo" not in avisos
    assert "Fatura média reajustada" not in blob


def test_3_mensal_explicita_continua_projetando_com_saldo_alto():
    at = _run(_sessao_pc(saldo=900000.0))
    blob = _blob(at)
    assert PARCIAL in blob and ESTIMADA not in blob
    at.radio(key="adequacao_v3_premissa").set_value("Mensal (média)")
    at.run()
    assert not at.exception, at.exception
    mensal = _blob(at)
    assert ESTIMADA in mensal
    assert _valor_do_card(mensal, "Diferença futura projetada") == pytest.approx(1126.98)
    assert _valor_do_card(mensal, ESTIMADA) == pytest.approx(RECONHECIDO + 1126.98)
    at.radio(key="adequacao_v3_premissa").set_value("Automática (cadência histórica)")
    at.run()
    assert not at.exception, at.exception
    assert PARCIAL in _blob(at)


# ------------------------------------------------ 4. irregular sem saldo

def test_4_pagina_irregular_sem_saldo_nao_inventa_futuro():
    at = _run(_sessao_pc(saldo=None))
    assert not at.exception, at.exception
    blob = _blob(at)
    assert NAO_DETERMINADA in _texto_do_card(blob, "Diferença futura total projetada")
    assert NAO_DETERMINADA in _texto_do_card(blob, "Diferença média mensal projetada")
    assert PARCIAL in blob and ESTIMADA not in blob
    assert _valor_do_card(blob, PARCIAL) == pytest.approx(RECONHECIDO)


def test_4_saldo_ausente_nao_vira_zero():
    assert saldo_elegivel_da_apuracao({"remanescente_original": 0.0})["valor"] is None
    com_posicao = {"remanescente_original": 0.0, "remanescente_reajustado": 0.0,
                   "df_remanescentes": pd.DataFrame([{"Ciclo": "C1"}])}
    assert saldo_elegivel_da_apuracao(com_posicao)["valor"] == 0.0
    # VTA nunca vira saldo
    assert saldo_elegivel_da_apuracao({"valor_atualizado_contrato": 1e6})["valor"] is None


def test_4_saldo_financeiro_inconciliavel_nao_vira_zero():
    info = saldo_elegivel_da_apuracao({
        "valor_original_contrato": 10_000_000.0,
        "valor_executado_original": 10_100_000.0,
    })
    assert info["valor"] is None
    assert "inconciliável" in info["origem"]


def test_4_mensal_sem_saldo_preserva_retroativo_e_nao_projeta_no_xlsx():
    at = _run(_sessao_pc(saldo=None))
    at.radio(key="adequacao_v3_premissa").set_value("Mensal (média)")
    at.run()
    assert not at.exception, at.exception
    blob = _blob(at)
    assert PARCIAL in blob
    assert _valor_do_card(blob, PARCIAL) == pytest.approx(RECONHECIDO)
    assert NAO_DETERMINADA in _texto_do_card(blob, "Diferença futura total projetada")
    wb = load_workbook(BytesIO(_xlsx_da_pagina(at)))
    ws = wb["PROJECAO"]
    cab = [c.value for c in ws[1]]
    col = cab.index("Diferença futura a adequar") + 1
    assert ws.max_row > 1
    assert all(ws.cell(r, col).value == NAO_DETERMINADA
               for r in range(2, ws.max_row + 1))
    resumo = wb["RESUMO"]
    assert any(resumo.cell(r, 2).value == NAO_DETERMINADA
               for r in range(5, resumo.max_row + 1))


# ----------------------------------- 5. retroativo com futuro indisponivel

def test_5_retroativo_mantido_e_futuro_nao_determinado():
    cons = consolidar_resultado_adequacao(retroativo=RECONHECIDO, df_projecao=pd.DataFrame(),
                                          meses_restantes=0, motivo_indeterminado="vigência ausente")
    assert cons["diferenca_futura"] is None and cons["diferenca_media_mensal"] is None
    assert cons["complementacao"] == pytest.approx(RECONHECIDO)
    assert cons["rotulo_complementacao"].startswith("Complementação orçamentária parcial")
    linhas = dict(linhas_resumo_resultado(cons))
    assert linhas["Diferença futura total projetada"] == NAO_DETERMINADA


def test_5_pagina_sem_vigencia_mantem_o_retroativo():
    at = _run(_sessao_pc(), vigencia="")
    assert not at.exception, at.exception
    blob = _blob(at)
    assert _valor_do_card(blob, "Retroativo reconhecido considerado") == pytest.approx(RECONHECIDO)
    assert NAO_DETERMINADA in _texto_do_card(blob, "Diferença futura total projetada")
    assert PARCIAL in blob


def test_5_programacao_separa_retroativo_e_nao_zera_o_futuro():
    proj = calcular_projecao(montar_base_editor(PERIODOS, 10000.0), 10000.0, FATOR)
    cron = cronograma_por_exercicio(proj, 1000.0, futuro_determinado=False)
    assert list(cron["Exercício"]) == ["2026"]
    assert cron.loc[0, "Retroativo reconhecido"] == 1000.0
    assert cron.loc[0, "Diferença futura projetada"] is None
    assert cron.loc[0, "Valor"] == 1000.0


# ---------------------------------------------- 6. potencial separado

def test_6_potencial_fica_fora_da_complementacao_e_da_programacao():
    at = _run(_sessao_pc())
    assert not at.exception, at.exception
    blob = _blob(at)
    parcial = _valor_do_card(blob, PARCIAL)
    assert parcial == pytest.approx(RECONHECIDO)
    assert _valor_do_card(blob, "Cenário com potencial") - parcial == pytest.approx(POTENCIAL, abs=0.01)
    wb = load_workbook(BytesIO(_xlsx_da_pagina(at)))
    ws = wb["PROGRAMACAO"]
    total = next(ws.cell(r, 4).value for r in range(2, ws.max_row + 1)
                 if ws.cell(r, 1).value == "TOTAL")
    assert total == pytest.approx(parcial)
    cab = [ws.cell(1, c).value for c in range(1, 5)]
    assert "Diferença futura projetada" in cab
    assert ws.cell(2, 3).value == NAO_DETERMINADA


# ---------------------------------------------- 7. manual acima do saldo

def test_7_manual_acima_do_saldo_nao_e_reduzido_e_exige_correcao():
    editor = montar_base_editor(PERIODOS[:3], 0.0, base_por_competencia={})
    editor["Valor informado pelo fiscal"] = ["4.000,00", "4.000,00", "4.000,00"]
    proj = calcular_projecao(editor, 0.0, FATOR, base_por_competencia={},
                             origem_automatica="Premissa manual", saldo_elegivel=10000.0)
    assert list(proj["Valor base considerado"]) == [4000.0, 4000.0, 4000.0]
    assert proj.loc[2, "Limite do saldo"] == SITUACAO_EXCEDE_SALDO
    cons = consolidar_resultado_adequacao(retroativo=500.0, df_projecao=proj,
                                          meses_restantes=3, saldo_info=_saldo(10000.0))
    assert cons["excesso_saldo"] == pytest.approx(2000.0)
    assert cons["futuro_determinado"] is False
    assert cons["complementacao"] == pytest.approx(500.0)
    assert "excedem o saldo" in cons["ressalvas"][0]


def test_7_motor_reserva_o_informado_e_limita_o_automatico():
    r = limitar_projecao_ao_saldo([(100.0, False), (300.0, True), (100.0, False)], 350.0)
    assert r["bases"] == [50.0, 300.0, 0.0]
    assert r["excesso"] == 0.0 and r["limitado"] is True


def test_7_pagina_manual_acima_do_saldo_sinaliza_erro():
    at = _run(_sessao_pc(saldo=1000.0))
    at.radio(key="adequacao_v3_premissa").set_value("Manual")
    at.run()
    chave = next(k for k in at.session_state.filtered_state
                 if str(k).startswith("adequacao_v3_editor"))
    at.session_state[chave] = {"edited_rows": {0: {"Valor informado pelo fiscal": "5.000,00"}},
                               "added_rows": [], "deleted_rows": []}
    at.run()
    assert not at.exception, at.exception
    erros = " ".join(e.value for e in at.error)
    assert "EXCEDEM O SALDO CONTRATUAL ELEGÍVEL" in erros
    blob = _blob(at)
    assert NAO_DETERMINADA in _texto_do_card(blob, "Diferença futura total projetada")
    assert PARCIAL in blob


# --------------------------------------------------------- 8. WEB = XLSX

_ARQUIVOS_BAIXAVEIS: dict[str, bytes] = {}


@pytest.fixture(autouse=True)
def _captura_downloads(monkeypatch):
    """Guarda os bytes entregues aos botoes de download (o AppTest descarta o
    runtime simulado ao fim de cada execucao)."""
    from streamlit.runtime.memory_media_file_storage import MemoryMediaFileStorage
    original = MemoryMediaFileStorage.load_and_get_id

    def _espiao(self, path_or_data, mimetype, kind, filename=None):
        file_id = original(self, path_or_data, mimetype, kind, filename)
        if isinstance(path_or_data, bytes):
            _ARQUIVOS_BAIXAVEIS[file_id] = path_or_data
        return file_id

    monkeypatch.setattr(MemoryMediaFileStorage, "load_and_get_id", _espiao)
    _ARQUIVOS_BAIXAVEIS.clear()
    yield


def _xlsx_da_pagina(at):
    """Bytes do XLSX realmente oferecido pelo botao "Baixar XLSX" da pagina."""
    botao = next(b for b in at.get("download_button") if b.proto.label == "Baixar XLSX")
    file_id = botao.proto.url.rsplit("/", 1)[-1].split(".")[0]
    return _ARQUIVOS_BAIXAVEIS[file_id]


@pytest.mark.parametrize("saldo", [900000.0, 5000.0, None])
def test_8_web_igual_xlsx(saldo):
    at = _run(_sessao_pc(saldo=saldo))
    assert not at.exception, at.exception
    blob = _blob(at)
    wb = load_workbook(BytesIO(_xlsx_da_pagina(at)))
    ws = wb["RESUMO"]
    resumo = {ws.cell(r, 1).value: ws.cell(r, 2).value for r in range(5, ws.max_row + 1)}

    def _xlsx_num(rotulo):
        valor = resumo[rotulo]
        return valor if isinstance(valor, (int, float)) else parse_moeda_br(valor)

    assert _xlsx_num("Retroativo reconhecido") == pytest.approx(
        _valor_do_card(blob, "Retroativo reconhecido considerado"), abs=0.005)
    assert resumo["Meses restantes da vigência"] == "18"
    rot_total = PARCIAL  # irregular: nao ha projecao automatica, mesmo com saldo
    rotulo_xlsx = next(k for k in resumo if k and k.upper().startswith(rot_total))
    assert _xlsx_num(rotulo_xlsx) == pytest.approx(_valor_do_card(blob, rot_total), abs=0.005)
    assert _xlsx_num(rotulo_xlsx) == pytest.approx(RECONHECIDO)
    assert _xlsx_num("Retroativo potencial (não reconhecido)") == pytest.approx(POTENCIAL)
    assert _xlsx_num("Cenário de planejamento com potencial") == pytest.approx(
        _valor_do_card(blob, "Cenário com potencial"), abs=0.005)
    assert "parcial" in str(ws["A2"].value).lower()
    assert resumo["Ocorrências previstas no horizonte"] == NAO_DETERMINADA
    assert "não aplicada à projeção automática" in resumo["Uso da referência mensal reajustada"]
    for rotulo in ("Diferença média mensal projetada", "Diferença futura total projetada"):
        assert resumo[rotulo] == NAO_DETERMINADA
        assert NAO_DETERMINADA in _texto_do_card(blob, rotulo)
    if saldo is not None:
        assert resumo["Saldo contratual elegível (sem reajuste)"] == pytest.approx(saldo)
    else:
        assert resumo["Saldo contratual elegível (sem reajuste)"] == "Não disponível"
    ws_p = wb["PROJECAO"]
    cab = [c.value for c in ws_p[1]]
    col = cab.index("Diferença futura a adequar") + 1
    assert "Limite do saldo" in cab if saldo is not None else True
    assert all(ws_p.cell(r, col).value == NAO_DETERMINADA
               for r in range(2, ws_p.max_row + 1))
    ws_prog = wb["PROGRAMACAO"]
    total = next(ws_prog.cell(r, 4).value for r in range(2, ws_prog.max_row + 1)
                 if ws_prog.cell(r, 1).value == "TOTAL")
    assert total == pytest.approx(RECONHECIDO)
    ressalvas = [ws.cell(r, 1).value for r in range(5, ws.max_row + 1)]
    assert "Ressalvas" in ressalvas
