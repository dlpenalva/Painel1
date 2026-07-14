import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
THEME = (ROOT / "_theme.py").read_text(encoding="utf-8")
CONFIG = (ROOT / ".streamlit" / "config.toml").read_text(encoding="utf-8")
APP = (ROOT / "app.py").read_text(encoding="utf-8")
GARANTIA = (ROOT / "pages" / "05_Garantia.py").read_text(encoding="utf-8")


class TemaClaroTests(unittest.TestCase):
    def test_streamlit_e_forcado_para_tema_claro(self):
        self.assertIn('base = "light"', CONFIG)
        self.assertIn('backgroundColor = "#F7F5EF"', CONFIG)
        self.assertIn('secondaryBackgroundColor = "#F1F6F9"', CONFIG)
        self.assertIn('dataframeHeaderBackgroundColor = "#E6F0F7"', CONFIG)

    def test_tema_global_e_renderizado_antes_das_paginas(self):
        self.assertIn("from _theme import render_cl8us_light_theme", APP)
        self.assertLess(APP.index("render_cl8us_light_theme()"), APP.index("pagina_atual.run()"))

    def test_componentes_obrigatorios_estao_cobertos(self):
        for seletor in (
            'data-baseweb="select"',
            'data-baseweb="popover"',
            'role="option"',
            'data-testid="stTextInput"',
            'data-testid="stNumberInput"',
            'data-testid="stDateInput"',
            'data-testid="stTextArea"',
            'data-testid="stDataFrame"',
            'data-testid="stDataEditor"',
            'data-testid="stTable"',
            'data-testid="stExpander"',
            'data-testid="stVerticalBlockBorderWrapper"',
            'data-testid="stAlertContentInfo"',
            'data-testid="stAlertContentWarning"',
            'data-testid="stAlertContentError"',
            'data-testid="stAlertContentSuccess"',
            'data-testid="stSidebar"',
        ):
            self.assertIn(seletor, THEME)

    def test_tema_global_nao_reestiliza_botoes_de_acao(self):
        self.assertNotIn('.stButton', THEME)
        self.assertNotIn('.stDownloadButton', THEME)
        self.assertNotIn("\n        button {", THEME)

    def test_tabela_customizada_nao_tem_cabecalho_escuro(self):
        self.assertNotIn("table.garantia-tabela th { background: #1F4E78", GARANTIA)
        self.assertIn("table.garantia-tabela th { background: #E6F0F7", GARANTIA)


if __name__ == "__main__":
    unittest.main()
