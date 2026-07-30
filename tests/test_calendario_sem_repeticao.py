"""Item 8 — remover a mensagem repetida do calendario na WEB.

O apontamento do calendario de ciclos (LINHA_TEMPORAL_*) era exibido duas vezes
na mesma pagina: no Assistente operacional (topo) e, de novo, no bloco
"Pendencias e alertas" do Painel detalhado. Estes testes provam que:

  * o Painel deixa de listar os codigos de calendario (presentation-only);
  * outros alertas continuam visiveis no Painel;
  * o view-model (lista de alertas) NAO e mutado — validacao integral;
  * o Assistente operacional continua traduzindo/exibindo o apontamento,
    preservando a "proxima acao" bloqueante.
"""
from __future__ import annotations

import copy

from _painel_executivo import (
    CODIGOS_CALENDARIO_NO_ASSISTENTE,
    _alertas_visiveis_painel,
)
from _assistente_fiscal import _traduzir_inconsistencias


def _alerta(codigo, nivel="ERRO_GRAVE"):
    return {"nivel": nivel, "codigo": codigo,
            "mensagem": f"tecnico {codigo}",
            "mensagem_negocio": f"negocio {codigo}", "identificador": ""}


def test_painel_omite_codigos_de_calendario():
    alertas = [_alerta("LINHA_TEMPORAL_INVALIDA"), _alerta("PC_SEM_DATA", "ALERTA")]
    visiveis = _alertas_visiveis_painel(alertas)
    codigos_visiveis = {a["codigo"] for a in visiveis}
    assert "LINHA_TEMPORAL_INVALIDA" not in codigos_visiveis   # calendario omitido
    assert "PC_SEM_DATA" in codigos_visiveis                   # demais preservados


def test_painel_omite_ambos_os_codigos_de_calendario():
    alertas = [_alerta(c) for c in sorted(CODIGOS_CALENDARIO_NO_ASSISTENTE)]
    assert _alertas_visiveis_painel(alertas) == []


def test_view_model_nao_e_mutado():
    alertas = [_alerta("LINHA_TEMPORAL_INVALIDA"), _alerta("FATOR_INDETERMINADO")]
    original = copy.deepcopy(alertas)
    _alertas_visiveis_painel(alertas)
    assert alertas == original            # lista original intacta (validacao integral)
    assert len(alertas) == 2


def test_assistente_ainda_exibe_apontamento_de_calendario():
    painel = {"alertas": [_alerta("LINHA_TEMPORAL_INVALIDA")]}
    inconsistencias = _traduzir_inconsistencias(painel)
    codigos = {i.get("codigo") for i in inconsistencias}
    assert "LINHA_TEMPORAL_INVALIDA" in codigos                # ainda visivel no Assistente
    assert any(i["gravidade"] == "bloqueante" for i in inconsistencias)  # continua bloqueante
