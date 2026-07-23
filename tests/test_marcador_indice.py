"""Marcador do 'Índice do contrato': circulo PRETO e PREENCHIDO (mesmo peso
visual dos marcadores solidos dos campos do formulario), a DIREITA do titulo.

Nao deve replicar o marcador contornado de navegacao do menu lateral
(stPageLink ::before), que e branco/contornado.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / "app.py").read_text(encoding="utf-8")


def _bloco(seletor: str) -> str:
    ini = APP.index(seletor)
    return APP[ini: APP.index("}", ini)]


def test_marcador_indice_e_preto_e_preenchido():
    idx = _bloco(".cl8us-index-title::after")
    assert "background: #1A1A1A" in idx        # preto
    assert "border-radius: 50%" in idx          # circulo
    assert "border:" not in idx                 # preenchido (sem contorno)
    assert "content: \"\"" in idx


def test_marcador_indice_nao_replica_stpagelink():
    idx = _bloco(".cl8us-index-title::after")
    # Nao pode usar o padrao contornado do menu lateral (branco + borda).
    assert "rgba(255, 255, 255, .84)" not in idx
    assert "border-radius: 999px" not in idx
    assert "1.5px solid" not in idx


def test_marcador_indice_fica_a_direita():
    idx = _bloco(".cl8us-index-title::after")
    assert "margin-left" in idx      # espacamento a esquerda da bolinha => fica a direita do texto
    assert "margin-right" not in idx


def test_marker_tecnico_invisivel_preservado():
    assert ".cl8us-index-marker { display: none; }" in APP
    assert ":has(.cl8us-index-marker)" in APP
