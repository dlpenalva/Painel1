"""VTA-M5 — PREVIA passa a qualificar o VTA, nao o status global.

Testes permanentes cobrindo os cenarios A-J da tarefa: o gate de PREVIA em
`_aplicar_previa_vta` deixou de usar status_resultados["geral"]
(RESULTADOS!B3, agregado que inclui eixos alheios ao VTA — ex.: H33/posicao
fisica atual) e passou a usar exclusivamente o selo especifico do VTA_FINAL
(RESULTADOS!H8, via status_resultados["vta"]/MEMORIA_RESULTADOS!E26).

`_mascarar_sintese_por_divergencia` (mascaramento por divergencia XLS x
Python) NAO foi alterado nesta tarefa — permanece a unica fonte que zera
sintese["vta"] antes de `_aplicar_previa_vta` rodar.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from _sumario_executivo import (  # noqa: E402
    _aplicar_previa_vta,
    _mascarar_sintese_por_divergencia,
    _selo_vta_validado,
)

_VTA = 132581980.10
_STATUS_VALIDADO = "CALCULADO — CONFERIR"
_STATUS_REVISE = "SEM CALCULO"


def _rodar(vta_inicial, status_geral, status_vta, xls=_VTA):
    sintese = {"vta": vta_inicial} if vta_inicial is not None else {}
    resultados_xls = {"valores": {"VTA_FINAL": xls}} if xls is not None else {}
    status_resultados = {"geral": status_geral, "vta": status_vta}
    _aplicar_previa_vta(sintese, resultados_xls, status_resultados)
    return sintese


def test_a_global_validado_e_vta_especifico_validado_e_definitivo():
    sintese = _rodar(_VTA, "VALIDADO", _STATUS_VALIDADO)
    assert sintese["vta"] == _VTA
    assert "vta_previa" not in sintese


def test_b_global_estimado_e_vta_especifico_validado_e_definitivo():
    sintese = _rodar(_VTA, "ESTIMADO", _STATUS_VALIDADO)
    assert sintese["vta"] == _VTA
    assert "vta_previa" not in sintese


def test_c_global_revise_por_motivo_nao_relacionado_e_vta_definitivo():
    """Status geral REVISE (ex.: H14/H24/H43:H50) sem que o selo especifico
    do VTA (H8) tambem seja REVISE -> VTA continua definitivo."""
    sintese = _rodar(_VTA, "REVISE", _STATUS_VALIDADO)
    assert sintese["vta"] == _VTA
    assert "vta_previa" not in sintese


def test_d_vta_especifico_revise_vira_previa_com_numero_utilizavel():
    sintese = _rodar(_VTA, "VALIDADO", _STATUS_REVISE)
    assert sintese["vta"] is None
    assert sintese["vta_previa"] == _VTA


def test_e_vta_divergente_preserva_politica_cautelar_via_mascaramento():
    """`_mascarar_sintese_por_divergencia` (nao alterado) zera sintese["vta"]
    quando VTA_FINAL diverge; `_aplicar_previa_vta` entao usa o numero XLS
    como PREVIA — igual a antes da VTA-M5."""
    sintese = {"vta": _VTA}
    _mascarar_sintese_por_divergencia(sintese, {"VTA_FINAL"})
    assert sintese["vta"] is None
    _aplicar_previa_vta(
        sintese,
        {"valores": {"VTA_FINAL": _VTA}},
        {"geral": "VALIDADO", "vta": _STATUS_VALIDADO},
    )
    assert sintese["vta"] is None
    assert sintese["vta_previa"] == _VTA


def test_f_vta_em_campos_nao_confiaveis_nao_e_definitivo():
    sintese = {"vta": _VTA}
    _mascarar_sintese_por_divergencia(sintese, {"VTA_FINAL"})
    assert sintese["vta"] is None
    assert sintese.get("vta_motivo")


def test_g_vta_ausente_com_numero_xls_utilizavel_preserva_previa():
    sintese = _rodar(None, "VALIDADO", None, xls=_VTA)
    assert sintese.get("vta") is None
    assert sintese["vta_previa"] == _VTA


def test_h_vta_ausente_sem_numero_utilizavel_nao_fabrica_zero():
    sintese = _rodar(None, "VALIDADO", None, xls=None)
    assert sintese.get("vta") is None
    assert "vta_previa" not in sintese


def test_i_vta_zero_real_e_especifico_validado_e_zero_definitivo():
    sintese = _rodar(0.0, "ESTIMADO", _STATUS_VALIDADO)
    assert sintese["vta"] == 0.0
    assert sintese["vta"] is not None
    assert "vta_previa" not in sintese


def test_j_vta_zero_real_mas_especifico_nao_definitivo_aplica_previa_sem_zero_virar_ausencia():
    sintese = _rodar(0.0, "VALIDADO", _STATUS_REVISE)
    assert sintese["vta"] is None
    assert sintese["vta_previa"] == 0.0
    assert sintese["vta_previa"] is not None


def test_selo_vta_validado_reconhece_os_tres_marcadores():
    assert _selo_vta_validado({"vta": "CALCULADO — CONFERIR"})
    assert _selo_vta_validado({"vta": "MANUAL VALIDADO"})
    assert _selo_vta_validado({"vta": "COM AJUSTE MANUAL"})
    assert not _selo_vta_validado({"vta": ""})
    assert not _selo_vta_validado({"vta": None})
    assert not _selo_vta_validado(None)
    assert not _selo_vta_validado({"geral": "VALIDADO"})  # sem chave "vta"
