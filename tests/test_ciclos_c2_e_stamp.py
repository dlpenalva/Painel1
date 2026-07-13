import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / "app.py").read_text(encoding="utf-8")
UI = (ROOT / "_ui_utils.py").read_text(encoding="utf-8")
VERSAO = (ROOT / "_versao.py").read_text(encoding="utf-8")
SIMPLES = (ROOT / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
MULTI = (ROOT / "pages" / "02_Calculo_Represados.py").read_text(encoding="utf-8")


class TestCiclosC2EStamp(unittest.TestCase):
    def test_calculadora_simples_permite_c1_ate_c4(self):
        self.assertIn('opcoes_ciclo_analise = ["C1", "C2", "C3", "C4"]', SIMPLES)
        self.assertIn('"Ciclo desta análise:"', SIMPLES)
        self.assertIn("primeiro_ciclo_num = _ciclo_para_numero(ciclo_analise)", SIMPLES)

    def test_multiciclo_permite_iniciar_em_c2_e_limita_em_c4(self):
        self.assertIn('options=["C1", "C2", "C3", "C4"]', MULTI)
        self.assertIn("range(primeiro_ciclo_num, 5)", MULTI)
        self.assertIn(
            "data_atual = dt_base_original + relativedelta(years=primeiro_ciclo_num - 1)",
            MULTI,
        )
        self.assertNotIn('"C5"', MULTI)
        self.assertNotIn("'C5'", MULTI)

    def test_contexto_historico_foi_removido_das_duas_interfaces(self):
        termos_proibidos = (
            "_render_contexto_contratual_anterior",
            "Preencher/editar Contexto do Contrato",
            "Situação anterior à análise",
            "Observação de rastreabilidade",
        )
        for pagina in (SIMPLES, MULTI):
            with self.subTest(termo=pagina[:30]):
                for termo in termos_proibidos:
                    self.assertNotIn(termo, pagina)
            self.assertIn("contexto_contratual = {}", pagina)

    def test_stamp_tem_fallback_brasileiro_e_e_renderizado(self):
        match = re.search(
            r'ATUALIZADO_EM_FALLBACK = "(\d{2}/\d{2}/\d{4} \d{2}:\d{2})"',
            VERSAO,
        )
        self.assertIsNotNone(match)
        self.assertIn("from _versao import atualizado_em", UI)
        self.assertIn('st.caption(f"Atualizado em {atualizado_em()}")', UI)
        self.assertIn("render_versao_sidebar()", APP)

    def test_sidebar_adota_indicador_visual_do_modelo_3(self):
        self.assertIn('a::before', APP)
        self.assertIn('a[aria-current="page"]::before', APP)
        self.assertIn("--cl8us-sidebar: #C6D9E8;", APP)


if __name__ == "__main__":
    unittest.main()
