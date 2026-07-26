"""Testes da composicao automatica do VTA pelo metodo Pedidos de Compra (PC).

Fixture sintetico reproduzindo apenas os dados indispensaveis do caso golden
CLEMAR (o arquivo real do usuario nunca e commitado). Golden aprovado:

    C0 executado (movimentacao fisica x VU_C0) = 1.687.868,00
    C1 executado por PC                        =   111.286,55
    C2 executado atualizado por PC             =   308.057,35
    C3 executado atualizado por PC             =   155.516,79
    C4 remanescente atualizado (item a item)   =   854.564,71
    VTA                                        = 3.117.293,40
"""
from __future__ import annotations

import _motor_composicao_vta as C

# (ITEM, VU_C0, VU_C4, QTD_REM_AJUSTADA_C0, _C1, _C4) — dados indispensaveis.
_ITENS = [
    (1, 15022, 17408.84, 95, 0, 0),
    (2, 2600, 3013.11, 95, 95, 45),
    (3, 6400, 7416.89, 95, 95, 54),
    (4, 2548, 2952.85, 18, 11, 0),
    (5, 2314, 2681.67, 8, 5, 0),
    (6, 4800, 5562.67, 40, 24, 10),
    (7, 2600, 3013.11, 16, 6, 0),
    (8, 3600, 4172, 108, 71, 63),
]

# itens_PC ja itemizados: usamos a soma de VALOR_ATUALIZADO por ciclo (regra
# itens_PC homologada). Uma linha sintetica por ciclo reproduz o golden.
_PC_POR_CICLO = {"C1": 111286.55, "C2": 308057.35, "C3": 155516.79}

GOLDEN = {
    "C0": 1687868.00,
    "C1": 111286.55,
    "C2": 308057.35,
    "C3": 155516.79,
    "C4_REM": 854564.71,
    "VTA": 3117293.40,
}


def _leitura(modo="pc", vigente="C4", pcs=None, posc=True, hist=True):
    posicao = [{
        "ITEM": it, "VU_ORIGINAL": vu0,
        "QTD_REM_AJUSTADA_C0": q0, "QTD_REM_AJUSTADA_C1": q1,
        "QTD_REM_AJUSTADA_C2": 0, "QTD_REM_AJUSTADA_C3": 0,
        "QTD_REM_AJUSTADA_C4": q4,
    } for (it, vu0, vu4, q0, q1, q4) in _ITENS]
    historico = [{
        "item": it,
        "vu_ciclos": {"VU_C0": vu0, "VU_C1": vu0, "VU_C2": vu0,
                      "VU_C3": vu0, "VU_C4": vu4},
    } for (it, vu0, vu4, q0, q1, q4) in _ITENS]
    itens_pc = []
    fonte = pcs if pcs is not None else _PC_POR_CICLO
    for ciclo, valor in fonte.items():
        itens_pc.append({
            "ciclo": ciclo, "valor_pc": valor, "valor_atualizado": valor,
            "entra_no_calculo": "Sim",
        })
    return {
        "controle": {"modo": modo, "ciclo_vigente": vigente},
        "parametros_v10": {"por_ciclo": {}},
        "posicao_contratual": {"itens": posicao if posc else []},
        "historico_vu": {"itens": historico if hist else []},
        "itens_pc_v10": {"itens": itens_pc},
    }


def _por_ciclo(comp):
    d = {l["ciclo"]: l["valor_atualizado"] for l in comp["execucao_por_ciclo"]}
    d["C4_REM"] = comp["saldo_remanescente"]["valor_atualizado"]
    return d


# ---- F. VTA-PC CLEMAR: R$ 3.117.293,40 -------------------------------------
def test_f_vta_pc_golden():
    comp = C.montar_composicao_vta(_leitura())
    assert comp["disponivel"] is True
    assert comp["metodo"] == "pc"
    assert comp["vta_composicao"] == GOLDEN["VTA"]


# ---- G. remanescente CLEMAR: R$ 854.564,71 ---------------------------------
def test_g_remanescente_golden():
    comp = C.montar_composicao_vta(_leitura())
    assert _por_ciclo(comp)["C4_REM"] == GOLDEN["C4_REM"]


# ---- H. C0 CLEMAR: R$ 1.687.868,00 -----------------------------------------
def test_h_c0_golden():
    comp = C.montar_composicao_vta(_leitura())
    assert _por_ciclo(comp)["C0"] == GOLDEN["C0"]


# ---- I/J/K. C1 / C2 / C3 ----------------------------------------------------
def test_ijk_ciclos_pc_golden():
    pc = _por_ciclo(C.montar_composicao_vta(_leitura()))
    assert pc["C1"] == GOLDEN["C1"]
    assert pc["C2"] == GOLDEN["C2"]
    assert pc["C3"] == GOLDEN["C3"]


# ---- L. PC nao cria quantidade fisica --------------------------------------
def test_l_pc_nao_cria_quantidade_fisica():
    """Um PC extra nao muda o remanescente: quantidade vem da posicao fisica."""
    base = C.montar_composicao_vta(_leitura())
    com_pc = _PC_POR_CICLO | {"C3": _PC_POR_CICLO["C3"] + 50000.0}
    alt = C.montar_composicao_vta(_leitura(pcs=com_pc))
    assert (_por_ciclo(alt)["C4_REM"]
            == _por_ciclo(base)["C4_REM"] == GOLDEN["C4_REM"])


# ---- M. PC posterior ao corte nao gera dupla contagem ----------------------
def test_m_pc_posterior_sem_dupla_contagem():
    """PC no ciclo vigente (C4) ou posterior nao entra na execucao do VTA."""
    com_posterior = _PC_POR_CICLO | {"C4": 999999.0}
    comp = C.montar_composicao_vta(_leitura(pcs=com_posterior))
    assert comp["vta_composicao"] == GOLDEN["VTA"]  # inalterado
    assert "C4" not in {l["ciclo"] for l in comp["execucao_por_ciclo"]}
    assert any("mesmo corte temporal" in a for a in comp["alertas"])


# ---- Fallback: base insuficiente -> CALCULO MANUAL REQUERIDO ---------------
def test_fallback_calculo_manual_requerido():
    comp = C.montar_composicao_vta(_leitura(posc=False))
    assert comp["disponivel"] is False
    assert "CALCULO MANUAL REQUERIDO" in comp["motivo"]
    assert comp["vta_composicao"] is None


# ---- 4.5. Outros metodos nao usam a composicao PC dedicada -----------------
def test_outros_metodos_nao_disparam_pc():
    comp = C.montar_composicao_vta(_leitura(modo="principal"))
    # Caminho legado por reconciliacao (sem reconciliacao no fixture => vazio),
    # nunca marca metodo == 'pc'.
    assert comp.get("metodo") != "pc"
