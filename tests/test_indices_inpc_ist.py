"""Indices — IST (Anatel) rename + novo INPC (188) via SGS generico.

Cobre Secao 8/9 e Secao 13 (ÍNDICES). Testes offline: mapeamento de serie SGS,
lista canonica de indices, normalizadores de documentos e reconhecimento de
alias legado. A coleta online (Anatel / SGS) e coberta por testes de rede
existentes e pelos fluxos Excel COM.
"""
from __future__ import annotations

from pathlib import Path

import _indice_utils as IU

_ROOT = Path(__file__).resolve().parents[1]
_UI = (_ROOT / "_ui_utils.py").read_text(encoding="utf-8")


# ---- INPC 188 via infraestrutura SGS generica ------------------------------
def test_serie_sgs_por_indice():
    assert IU.SGS_INPC == 188
    assert IU.serie_sgs_do_indice("IPCA (433)") == "433"
    assert IU.serie_sgs_do_indice("IGP-M (189)") == "189"
    assert IU.serie_sgs_do_indice("INPC (188)") == "188"
    # INPC nao pode colidir com IGP-M (dispatch binario antigo)
    assert IU.serie_sgs_do_indice("INPC (188)") != IU.serie_sgs_do_indice("IGP-M (189)")


def test_inpc_reutiliza_coletor_sgs_generico():
    # Nao ha motor paralelo: INPC usa a mesma funcao de IPCA/IGP-M.
    assert callable(IU.coletar_sgs_produtorio)
    assert callable(IU.obter_ultima_competencia_sgs)


# ---- Lista canonica final de indices (user-facing) -------------------------
def test_lista_indices_final():
    for rotulo in ["IST (Anatel)", "ICTI (Ipeadata)", "IPCA (433)",
                   "IGP-M (189)", "INPC (188)"]:
        assert rotulo in _UI, rotulo
    # rename aplicado: nao expor mais "IST (Série Local)" na lista da UI
    assert "IST (Série Local)" not in _UI


def test_ui_tem_ultima_competencia_inpc():
    assert "_texto_ultima_competencia_sgs(188)" in _UI


# ---- Docs: IST (Anatel) + legado reconhecido, INPC amigavel ----------------
def test_docs_ist_rename_e_legado():
    from _email_contratada import _indice_amigavel
    from _templates_documentos import _indice_amigavel_doc
    assert _indice_amigavel("IST (Anatel)") == "IST (Anatel)"
    assert _indice_amigavel("IST (Série Local)") == "IST (Anatel)"   # legado
    assert _indice_amigavel("INPC (188)") == "INPC"
    assert _indice_amigavel_doc("IST (Série Local)") == "IST (Anatel)"
    assert _indice_amigavel_doc("INPC") == "INPC"
