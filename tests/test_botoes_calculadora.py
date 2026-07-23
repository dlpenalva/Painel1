"""§10.5 — Botoes da Calculadora 1 ciclo usam o padrao visual global.

A pagina nao deve injetar CSS local divergente para botoes nem renderizar
"botoes" HTML com paleta propria. Os botoes de acao seguem o mesmo padrao
(type primary) ja homologado nas demais calculadoras.
"""
from pathlib import Path

import re

ROOT = Path(__file__).resolve().parents[1]
PAGINA_1CICLO = ROOT / "pages" / "01_Calculo_Simples.py"
PAGINA_MULTI = ROOT / "pages" / "02_Calculo_Represados.py"


def _fonte(p: Path) -> str:
    return p.read_text(encoding="utf-8")


def test_sem_override_local_de_css_de_botao():
    fonte = _fonte(PAGINA_1CICLO)
    assert not re.search(r"stButton\s*>\s*button", fonte)
    assert not re.search(r"stDownloadButton\s*>\s*button", fonte)


def test_nenhum_botao_html_verde_e_renderizado():
    """A funcao do 'botao' HTML verde (#4E6E58) permanece sem chamada ativa."""
    fonte = _fonte(PAGINA_1CICLO)
    chamadas = [
        l for l in fonte.splitlines()
        if "render_botao_download_modelo_consumo(" in l and "def " not in l
    ]
    assert chamadas == [], f"botao HTML divergente sendo renderizado: {chamadas}"


def test_botoes_de_acao_usam_padrao_primary_global():
    fonte = _fonte(PAGINA_1CICLO)
    bloco = fonte[fonte.index("Baixar Arquivo Coleta Oficial"):]
    bloco = bloco[: bloco.index(")") + 400]
    assert 'type="primary"' in bloco
    assert "render_email_contratada(" in fonte


def test_padrao_coerente_com_multiciclo():
    for fonte in (_fonte(PAGINA_1CICLO), _fonte(PAGINA_MULTI)):
        assert "Baixar Arquivo Coleta Oficial" in fonte
        assert 'type="primary"' in fonte
