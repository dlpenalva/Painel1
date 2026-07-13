from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / "app.py").read_text(encoding="utf-8")
INICIO = (ROOT / "pages" / "00_Calculadora_Reajustes.py").read_text(encoding="utf-8")


class TestCascaXlsFirst(unittest.TestCase):
    def test_menu_principal_tem_somente_as_quatro_rotas_operacionais(self):
        self.assertIn('st.page_link(PAGINA_INICIO, label="Início")', APP)
        self.assertIn('st.page_link(PAGINA_UM_CICLO, label="Calculadora 1 ciclo")', APP)
        self.assertIn('st.page_link(PAGINA_MULTICICLO, label="Calculadora multiciclo")', APP)
        self.assertIn('st.page_link(PAGINA_UPLOAD, label="Upload e resultados")', APP)
        self.assertIn('position="hidden"', APP)

    def test_modulos_legados_permanecem_registrados_mas_recolhidos(self):
        self.assertIn('st.expander("Ferramentas complementares", expanded=False)', APP)
        self.assertIn('("04_Relatorio_Global.py", "Relatórios")', APP)
        self.assertIn('("13_DOU.py", "DOU")', APP)

    def test_inicio_expoe_quatro_boxes_e_os_destinos_corretos(self):
        for numero in range(1, 5):
            self.assertEqual(INICIO.count(f'"{numero} ·'), 1)
        self.assertIn('st.switch_page("pages/03_Valor_Global.py")', INICIO)
        self.assertIn('st.switch_page("pages/01_Calculo_Simples.py")', INICIO)
        self.assertIn('st.switch_page("pages/02_Calculo_Represados.py")', INICIO)

    def test_modelo_xls_e_a_fonte_do_download_inicial(self):
        self.assertIn("CAMINHO_MODELO_COLETA", INICIO)
        self.assertIn("NOME_ARQUIVO_COLETA", INICIO)
        self.assertIn('file_name=NOME_ARQUIVO_COLETA', INICIO)
        self.assertIn('"Baixar Coleta_Reajuste.xlsx"', INICIO)

    def test_inicio_nao_reintroduz_seletor_intermediario_da_versao_antiga(self):
        self.assertNotIn("executar_motor", INICIO)
        self.assertNotIn("fluxo_query", INICIO)
        self.assertNotIn("A análise envolve mais de um ciclo", INICIO)
        self.assertNotIn("runpy", INICIO)


if __name__ == "__main__":
    unittest.main()
