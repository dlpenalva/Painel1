"""§11 — A bolinha do 'Índice do contrato' segue o mesmo padrao do marcador
ja existente no menu lateral (stPageLink ::before), sem criar novo padrao.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / "app.py").read_text(encoding="utf-8")


def _bloco(seletor: str) -> str:
    ini = APP.index(seletor)
    return APP[ini: APP.index("}", ini)]


def test_marcador_indice_reutiliza_dimensao_do_menu_lateral():
    ref = _bloco("[data-testid=\"stSidebar\"] [data-testid=\"stPageLink\"] a::before")
    idx = _bloco(".cl8us-index-title::after")
    for prop in ("width: .72rem", "height: .72rem", "border-radius: 999px",
                 "border: 1.5px solid rgba(18, 59, 99, .36)",
                 "background: rgba(255, 255, 255, .84)", "vertical-align: -.06rem"):
        assert prop in ref, f"referencia perdeu {prop!r}"
        assert prop in idx, f"marcador do indice nao reutiliza {prop!r}"


def test_marcador_indice_fica_a_direita():
    idx = _bloco(".cl8us-index-title::after")
    assert "margin-left" in idx      # espacamento a esquerda da bolinha => fica a direita do texto
    assert "margin-right" not in idx


def test_marcador_indice_nao_usa_padrao_antigo_pequeno():
    idx = _bloco(".cl8us-index-title::after")
    assert "width: .5rem" not in idx
    assert "#1A1A1A" not in idx


def test_marker_tecnico_invisivel_preservado():
    assert ".cl8us-index-marker { display: none; }" in APP
    assert ":has(.cl8us-index-marker)" in APP
