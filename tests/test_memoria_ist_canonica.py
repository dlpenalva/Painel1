"""Etapa 2 do hotfix: memoria canonica IST (serie mensal) alimenta Web + XLS.

O metodo matematico do IST (v_fim / v_ini) NAO muda; apenas a memoria/auditoria
passa a preservar todas as competencias mensais reais do ist.csv no intervalo.
A mesma estrutura res['dados'] abastece a Calculadora Web e o XLS parametros J:R.
"""
import datetime
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from _indice_utils import calcular_ist_numero_indice
from _memoria_calculo import normalizar_memoria_calculo

PAGINA_01 = (ROOT / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
PAGINA_02 = (ROOT / "pages" / "02_Calculo_Represados.py").read_text(encoding="utf-8")


class TestProviderISTSerieMensal(unittest.TestCase):
    def _ciclo(self, ano, mes):
        res = calcular_ist_numero_indice(datetime.date(ano, mes, 1))
        self.assertIsNotNone(res, f"ist.csv sem cobertura para {mes:02d}/{ano}")
        return res

    def test_a_ist_1_ciclo_memoria_mensal_completa(self):
        res = self._ciclo(2023, 8)
        dados = res["dados"]
        # 12 meses => 13 competencias mensais (mes-base ate mes-base+12, inclusive)
        self.assertEqual(len(dados), 13)
        self.assertEqual(list(dados.columns), ["data", "indice"])
        # matematica homologada intacta
        self.assertAlmostEqual(res["variacao"], res["i_fim"] / res["i_ini"] - 1, places=12)
        # endpoints coincidem com i_ini/i_fim
        self.assertAlmostEqual(float(dados["indice"].iloc[0]), res["i_ini"], places=6)
        self.assertAlmostEqual(float(dados["indice"].iloc[-1]), res["i_fim"], places=6)

    def test_c_normalizador_ist_gera_competencias_mensais_mais_resultado(self):
        res = self._ciclo(2023, 8)
        mem = normalizar_memoria_calculo(res, res["variacao"] + 1, res["variacao"])
        tipos = [m["tipo"] for m in mem]
        self.assertEqual(tipos.count("INDICE"), 13)
        self.assertEqual(tipos.count("RESULTADO"), 1)
        self.assertEqual(tipos[-1], "RESULTADO")
        # cada INDICE tem competencia e valor de indice reais (nao fabricado)
        for m in mem[:-1]:
            self.assertIsNotNone(m["competencia"])
            self.assertIsNotNone(m["valor_indice"])
        self.assertAlmostEqual(mem[-1]["variacao_final"], res["variacao"], places=12)

    def test_b_ist_multiciclo_cada_ciclo_tem_memoria_mensal(self):
        for ano, mes in ((2022, 1), (2023, 8)):
            res = self._ciclo(ano, mes)
            self.assertGreaterEqual(len(res["dados"]), 12)
            self.assertAlmostEqual(res["variacao"], res["i_fim"] / res["i_ini"] - 1, places=12)

    def test_nao_fabrica_competencias_serie_e_subconjunto_do_csv(self):
        import pandas as pd
        from _indice_utils import carregar_ist_local
        base = carregar_ist_local(str(ROOT / "ist.csv"))
        res = self._ciclo(2023, 8)
        datas_csv = set(pd.to_datetime(base["data"]).dt.strftime("%Y-%m"))
        datas_mem = set(pd.to_datetime(res["dados"]["data"]).dt.strftime("%Y-%m"))
        self.assertTrue(datas_mem.issubset(datas_csv), "memoria contem competencia inexistente no ist.csv")


class TestWebRenderizaMemoriaIST(unittest.TestCase):
    def test_a_calc_simples_ist_renderiza_dataframe(self):
        # Ancora no bloco da MEMORIA (expander), onde a equacao IST e renderizada.
        ini = PAGINA_01.index("_render_equacao_ist(float(res['i_ini'])")
        trecho = PAGINA_01[ini:PAGINA_01.index('elif "ICTI" in tipo_idx:', ini)]
        self.assertIn("st.dataframe(res['dados']", trecho)

    def test_b_calc_multiciclo_ist_renderiza_dataframe(self):
        idx = PAGINA_02.index("_render_equacao_ist(res_c)")
        trecho = PAGINA_02[idx:PAGINA_02.index('elif "ICTI" in idx_sel:', idx)]
        self.assertIn("st.dataframe(res_c['dados']", trecho)


class TestRegressaoNormalizadorInalterado(unittest.TestCase):
    """O normalizador nao muda: N linhas 'indice' -> N INDICE; fixture de 2 linhas -> 2."""

    def test_fixture_duas_linhas_continua_duas(self):
        import pandas as pd
        res = {
            "metodo": "Divisão de Número-Índice (Série Local)",
            "serie": "IST",
            "dados": pd.DataFrame({"data": [pd.Timestamp(2024, 1, 1), pd.Timestamp(2025, 1, 1)], "indice": [165.1, 172.4]}),
        }
        mem = normalizar_memoria_calculo(res, 1.044, 0.044)
        self.assertEqual([m["tipo"] for m in mem], ["INDICE", "INDICE", "RESULTADO"])


if __name__ == "__main__":
    unittest.main()
