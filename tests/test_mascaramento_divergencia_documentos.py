"""Etapa 5b: mascaramento de campos nao-confiaveis nos documentos liberados.

Regra: uma divergencia relevante XLS x Python esvazia o campo divergente e os
que dependem diretamente dele, sem adotar XLS nem Python. Campos independentes
permanecem preenchidos. Retroativo e VTA/remanescente sao dimensoes distintas.
"""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from _reconciliacao_xls_python import campos_nao_confiaveis_para_documentos
from _sumario_executivo import (
    _mascarar_financeiro_por_divergencia,
    _mascarar_sintese_por_divergencia,
)


def _recon(*campos):
    return {"divergencias_relevantes": [{"campo": c} for c in campos]}


class TestExpansaoDependencias(unittest.TestCase):
    def test_rem_atualizado_propaga_para_vta(self):
        nc = campos_nao_confiaveis_para_documentos(_recon("REM_ATUALIZADO_OFICIAL"))
        self.assertEqual(nc, {"REM_ATUALIZADO_OFICIAL", "VTA_FINAL"})

    def test_rem_base_propaga_ate_vta(self):
        nc = campos_nao_confiaveis_para_documentos(_recon("REM_BASE_OFICIAL"))
        self.assertEqual(nc, {"REM_BASE_OFICIAL", "REM_ATUALIZADO_OFICIAL", "VTA_FINAL"})

    def test_qtd_rem_propaga_toda_a_cadeia(self):
        nc = campos_nao_confiaveis_para_documentos(_recon("QTD_REM_OFICIAL"))
        self.assertEqual(
            nc,
            {"QTD_REM_OFICIAL", "REM_BASE_OFICIAL", "REM_ATUALIZADO_OFICIAL", "VTA_FINAL"},
        )

    def test_retro_metodo_propaga_para_retro_oficial(self):
        self.assertEqual(
            campos_nao_confiaveis_para_documentos(_recon("RETRO_FIN")),
            {"RETRO_FIN", "RETRO_OFICIAL"},
        )

    def test_retroativo_e_vta_sao_independentes(self):
        # Divergencia so no retroativo NAO contamina VTA, e vice-versa.
        nc_retro = campos_nao_confiaveis_para_documentos(_recon("RETRO_OFICIAL"))
        self.assertNotIn("VTA_FINAL", nc_retro)
        nc_vta = campos_nao_confiaveis_para_documentos(_recon("VTA_FINAL"))
        self.assertNotIn("RETRO_OFICIAL", nc_vta)

    def test_sem_divergencia_conjunto_vazio(self):
        self.assertEqual(campos_nao_confiaveis_para_documentos(_recon()), set())
        self.assertEqual(campos_nao_confiaveis_para_documentos(None), set())


class TestMascaramentoSintese(unittest.TestCase):
    def _sintese(self):
        return {"vta": 201090.94, "vta_motivo": None,
                "retroativo_total": 5000.0, "retroativo_estado": "ok"}

    def test_vta_final_esvazia_vta(self):
        s = self._sintese()
        _mascarar_sintese_por_divergencia(s, {"VTA_FINAL"})
        self.assertIsNone(s["vta"])
        self.assertTrue(s["vta_motivo"])            # motivo visivel
        self.assertEqual(s["retroativo_total"], 5000.0)  # independente preservado

    def test_retro_oficial_esvazia_retroativo(self):
        s = self._sintese()
        _mascarar_sintese_por_divergencia(s, {"RETRO_OFICIAL"})
        self.assertIsNone(s["retroativo_total"])
        self.assertEqual(s["vta"], 201090.94)       # independente preservado

    def test_sem_campos_nao_altera(self):
        s = self._sintese()
        _mascarar_sintese_por_divergencia(s, set())
        self.assertEqual(s["vta"], 201090.94)
        self.assertEqual(s["retroativo_total"], 5000.0)


class TestMascaramentoFinanceiro(unittest.TestCase):
    def _fin(self):
        return {"delta_total_financeiro": 100.0, "delta_total_pc": 200.0,
                "retroativo_total": 300.0, "retroativo_estado": "ok"}

    def test_retro_oficial_esvazia_totais_do_documento(self):
        f = self._fin()
        _mascarar_financeiro_por_divergencia(f, {"RETRO_OFICIAL"})
        self.assertIsNone(f["delta_total_financeiro"])
        self.assertIsNone(f["delta_total_pc"])
        self.assertIsNone(f["retroativo_total"])

    def test_retro_fin_esvazia_apenas_financeiro(self):
        f = self._fin()
        _mascarar_financeiro_por_divergencia(f, {"RETRO_FIN"})
        self.assertIsNone(f["delta_total_financeiro"])
        self.assertEqual(f["delta_total_pc"], 200.0)

    def test_vta_final_nao_afeta_retroativo(self):
        f = self._fin()
        _mascarar_financeiro_por_divergencia(f, {"VTA_FINAL"})
        self.assertEqual(f["delta_total_financeiro"], 100.0)
        self.assertEqual(f["delta_total_pc"], 200.0)


if __name__ == "__main__":
    unittest.main()
