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
            "data_atual = _calcular_data_inicial_ciclo(_dt_base_calculo, primeiro_ciclo_num, _contexto_calculo)",
            MULTI,
        )
        self.assertNotIn('"C5"', MULTI)
        self.assertNotIn("'C5'", MULTI)

    def test_historico_anterior_aparece_somente_no_multiciclo_apos_c1(self):
        self.assertIn("if int(primeiro_ciclo_num) > 1:", MULTI)
        self.assertIn("Histórico anterior à análise", MULTI)
        self.assertIn("Situação anterior à análise:", MULTI)
        self.assertIn("Nenhum ciclo anterior concedido", MULTI)
        self.assertIn("Houve ciclo anterior concedido/formalizado", MULTI)
        self.assertIn("Situação desconhecida", MULTI)
        self.assertIn("Último ciclo concedido/formalizado:", MULTI)
        self.assertIn("Marco temporal do último ciclo concedido/formalizado:", MULTI)
        self.assertIn("'contexto_contratual_anterior': contexto_contratual", MULTI)
        self.assertNotIn("Situação anterior à análise:", SIMPLES)

    def test_regra_de_ancoragem_replica_a_versao_3(self):
        self.assertIn("def _calcular_data_inicial_ciclo", MULTI)
        self.assertIn("salto = numero_inicial - ultimo_num - 1", MULTI)
        self.assertIn("return dt_base + relativedelta(years=numero_inicial - 1)", MULTI)

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
        self.assertIn("--cl8us-input: #FFF9E8;", APP)
        self.assertIn("--cl8us-index: #FBE8AD;", APP)
        self.assertIn(":has(.cl8us-index-marker)", APP)
        self.assertIn("cl8us-cycle-step", MULTI)
        self.assertIn("cl8us-interval-box", MULTI)
        self.assertIn("cl8us-index-marker", UI)


if __name__ == "__main__":
    unittest.main()
