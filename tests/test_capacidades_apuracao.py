import io
import unittest

from openpyxl import load_workbook

from _capacidades_apuracao import avaliar_capacidades_apuracao
from _coleta_reajuste import gerar_coleta_reajuste, ler_coleta_reajuste
from _coleta_reajuste_documentos import adaptar_coleta_reajuste_para_documentos


class CapacidadesApuracaoTests(unittest.TestCase):
    DOCUMENTOS_PRINCIPAIS = (
        "planilha_executiva",
        "valores_unitarios",
        "relatorio_executivo",
        "mapa_marcos",
        "minuta_apostilamento",
        "garantia_contratual",
        "dou",
        "checklist_processual",
    )

    def test_somente_financeiro_libera_metodo_sem_inventar_resultado_oficial(self):
        capacidades = avaliar_capacidades_apuracao(
            {"competencias_com_valor": 2},
            {
                "ciclos_em_analise": ["C1"],
                "status_resultados": {
                    "retroativo": "CÁLCULO MANUAL REQUERIDO",
                    "valores": {"retroativo_financeiro": 120.50},
                },
            },
        )
        self.assertEqual(capacidades["blocos"]["financeiro"]["estado"], "completo")
        self.assertEqual(capacidades["blocos"]["pcs"]["estado"], "nao_informado")
        self.assertEqual(capacidades["calculos"]["retroativo"]["estado"], "parcial")
        self.assertTrue(capacidades["calculos"]["retroativo"]["disponivel"])
        self.assertEqual(capacidades["calculos"]["retroativo"]["valor"], 120.50)
        self.assertTrue(capacidades["documentos"]["planilha_executiva"]["habilitado"])
        self.assertTrue(capacidades["documentos"]["relatorio_executivo"]["habilitado"])
        self.assertFalse(capacidades["documentos"]["minuta_apostilamento"]["habilitado"])

    def test_somente_itens_libera_posicao_vu_remanescente_e_memoria(self):
        capacidades = avaliar_capacidades_apuracao(
            {
                "itens_remanescentes": 1,
                "posicao_contratual_itens": 1,
                "posicao_contratual_calculada": 1,
                "historico_vu_itens": 1,
                "historico_vu_calculado": 1,
            },
            {
                "ciclos_em_analise": ["C2"],
                "status_resultados": {
                    "remanescente": "BASE ÚNICA — CONFERIR",
                    "valores": {"remanescente_atualizado": 1040.20},
                },
            },
        )
        self.assertEqual(capacidades["blocos"]["financeiro"]["estado"], "nao_informado")
        self.assertEqual(capacidades["calculos"]["posicao_contratual"]["estado"], "completo")
        self.assertEqual(capacidades["calculos"]["valores_unitarios"]["estado"], "completo")
        self.assertTrue(capacidades["calculos"]["valor_remanescente"]["disponivel"])
        self.assertTrue(capacidades["documentos"]["mapa_marcos"]["habilitado"])
        self.assertTrue(capacidades["documentos"]["valores_unitarios"]["habilitado"])

    def test_defeito_estrutural_bloqueia_sem_confundir_com_dado_ausente(self):
        capacidades = avaliar_capacidades_apuracao(
            {"competencias_com_valor": 1},
            {"ciclos_em_analise": ["C1"]},
            ["Fórmula estrutural ausente"],
        )
        self.assertFalse(capacidades["estruturalmente_valido"])
        self.assertTrue(all(item["estado"] == "bloqueado" for item in capacidades["blocos"].values()))
        documentos_apuracao = (
            "planilha_executiva", "valores_unitarios", "relatorio_executivo", "mapa_marcos",
            "minuta_apostilamento", "garantia_contratual", "dou", "checklist_processual",
        )
        for chave in documentos_apuracao:
            self.assertEqual(capacidades["documentos"][chave]["estado"], "bloqueado")
            self.assertFalse(capacidades["documentos"][chave]["habilitado"])

    def test_adaptador_aceita_upload_parcial_e_preserva_estados_individuais(self):
        dados = {
            "indice": "IPCA",
            "ciclos": [
                {
                    "ciclo": "C1",
                    "data_base": "01/01/2024",
                    "financeiro_inicio": "01/01/2025",
                    "percentual_aplicado": 0.045,
                }
            ],
        }
        wb = load_workbook(io.BytesIO(gerar_coleta_reajuste(dados)), data_only=False)
        wb["financeiro"]["C2"] = 1000.0
        saida = io.BytesIO()
        wb.save(saida)
        payload = saida.getvalue()

        diagnostico = ler_coleta_reajuste(payload)
        resultado = adaptar_coleta_reajuste_para_documentos(payload)

        self.assertTrue(diagnostico["valido"])
        self.assertFalse(diagnostico["pronto_para_consolidar"])
        self.assertTrue(diagnostico["processamento_progressivo"])
        self.assertTrue(resultado["ok"])
        self.assertEqual(resultado["capacidades"]["blocos"]["financeiro"]["estado"], "completo")
        self.assertEqual(resultado["capacidades"]["blocos"]["consumidos"]["estado"], "nao_informado")
        self.assertIn("minuta_apostilamento", resultado["capacidades"]["documentos"])

    def test_cinco_cenarios_preservam_oito_documentos_com_estado_individual(self):
        cenarios = {
            "somente_financeiro": (
                {"competencias_com_valor": 2},
                {
                    "retroativo": "CÁLCULO MANUAL REQUERIDO",
                    "valores": {"retroativo_financeiro": 40.0},
                },
            ),
            "financeiro_remanescentes": (
                {
                    "competencias_com_valor": 2,
                    "itens_remanescentes": 1,
                    "posicao_contratual_itens": 1,
                    "posicao_contratual_calculada": 1,
                    "historico_vu_itens": 1,
                    "historico_vu_calculado": 1,
                },
                {
                    "retroativo": "CALCULADO — CONFERIR",
                    "vta": "CALCULADO — CONFERIR",
                    "remanescente": "BASE ÚNICA — CONFERIR",
                    "valores": {
                        "retroativo_oficial": 40.0,
                        "vta_oficial": 1080.0,
                        "remanescente_atualizado": 1040.0,
                    },
                },
            ),
            "pcs_remanescentes": (
                {
                    "pedidos_de_compra": 1,
                    "itens_remanescentes": 1,
                    "posicao_contratual_itens": 1,
                    "posicao_contratual_calculada": 1,
                    "historico_vu_itens": 1,
                    "historico_vu_calculado": 1,
                },
                {
                    "retroativo": "CALCULADO — CONFERIR",
                    "vta": "CALCULADO — CONFERIR",
                    "remanescente": "BASE ÚNICA — CONFERIR",
                    "valores": {
                        "retroativo_pc": 20.0,
                        "retroativo_oficial": 20.0,
                        "vta_oficial": 1060.0,
                        "remanescente_atualizado": 1040.0,
                    },
                },
            ),
            "financeiro_pcs_remanescentes": (
                {
                    "competencias_com_valor": 2,
                    "pedidos_de_compra": 1,
                    "itens_remanescentes": 1,
                    "posicao_contratual_itens": 1,
                    "posicao_contratual_calculada": 1,
                    "historico_vu_itens": 1,
                    "historico_vu_calculado": 1,
                },
                {
                    "retroativo": "CALCULADO — CONFERIR",
                    "vta": "CALCULADO — CONFERIR",
                    "remanescente": "BASE ÚNICA — CONFERIR",
                    "valores": {
                        "retroativo_financeiro": 40.0,
                        "retroativo_pc": 20.0,
                        "retroativo_oficial": 40.0,
                        "vta_oficial": 1080.0,
                        "remanescente_atualizado": 1040.0,
                    },
                },
            ),
            "processo_completo": (
                {
                    "competencias_com_valor": 2,
                    "pedidos_de_compra": 1,
                    "itens_remanescentes": 1,
                    "itens_consumidos": 1,
                    "aditivos": 1,
                    "posicao_contratual_itens": 1,
                    "posicao_contratual_calculada": 1,
                    "historico_vu_itens": 1,
                    "historico_vu_calculado": 1,
                },
                {
                    "retroativo": "CALCULADO — CONFERIR",
                    "vta": "CALCULADO — CONFERIR",
                    "remanescente": "CALCULADO — CONFERIR",
                    "valores": {
                        "retroativo_financeiro": 40.0,
                        "retroativo_pc": 20.0,
                        "retroativo_itens": 10.0,
                        "retroativo_oficial": 40.0,
                        "vta_oficial": 1080.0,
                        "remanescente_atualizado": 1040.0,
                    },
                },
            ),
        }

        classificacoes = {
            "DISPONÍVEL",
            "DISPONÍVEL COM RESSALVAS",
            "PENDENTE DE DADOS",
            "NÃO APLICÁVEL",
        }
        for nome, (contagens, status) in cenarios.items():
            with self.subTest(cenario=nome):
                capacidades = avaliar_capacidades_apuracao(
                    contagens,
                    {"ciclos_em_analise": ["C1"], "status_resultados": status},
                )
                documentos = capacidades["documentos"]
                self.assertTrue(all(chave in documentos for chave in self.DOCUMENTOS_PRINCIPAIS))
                self.assertTrue(
                    all(documentos[chave]["classificacao"] in classificacoes for chave in self.DOCUMENTOS_PRINCIPAIS)
                )
                self.assertTrue(documentos["checklist_processual"]["habilitado"])
                self.assertTrue(documentos["dou"]["habilitado"])

        somente_financeiro = avaliar_capacidades_apuracao(
            cenarios["somente_financeiro"][0],
            {"ciclos_em_analise": ["C1"], "status_resultados": cenarios["somente_financeiro"][1]},
        )
        self.assertTrue(somente_financeiro["calculos"]["retroativo"]["disponivel"])
        self.assertFalse(somente_financeiro["calculos"]["vta"]["disponivel"])
        self.assertEqual(
            somente_financeiro["documentos"]["minuta_apostilamento"]["classificacao"],
            "PENDENTE DE DADOS",
        )


if __name__ == "__main__":
    unittest.main()
