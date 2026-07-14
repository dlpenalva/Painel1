import copy
import json
import unittest
from pathlib import Path

from _capacidades_apuracao import avaliar_capacidades_apuracao


FIXTURE_PATH = Path(__file__).parent / "fixtures" / "capacidades_apuracao.json"


class CapacidadesRegressaoTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.fixture = json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))
        cls.cenarios = cls.fixture["cenarios"]
        cls.documentos = cls.fixture["documentos_principais"]

    def _avaliar(self, nome):
        entradas = copy.deepcopy(self.cenarios[nome]["entradas"])
        return avaliar_capacidades_apuracao(
            entradas["contagens"],
            entradas["metadados"],
        )

    @staticmethod
    def _capacidade(capacidades, chave):
        if chave in capacidades["blocos"]:
            return capacidades["blocos"][chave]
        return capacidades["calculos"][chave]

    def test_fixture_permanente_registra_os_cinco_cenarios_homologados(self):
        self.assertEqual(self.fixture["versao_fixture"], 1)
        self.assertEqual(len(self.fixture["invariantes"]), 5)
        self.assertEqual(
            set(self.cenarios),
            {
                "somente_financeiro",
                "financeiro_remanescentes",
                "pcs_remanescentes",
                "financeiro_pcs_remanescentes",
                "processo_completo",
            },
        )

        for nome, esperado in self.cenarios.items():
            with self.subTest(cenario=nome):
                capacidades = self._avaliar(nome)
                calculos = capacidades["calculos"]
                resultados = esperado["resultados_esperados"]

                for chave in esperado["capacidades_disponiveis"]:
                    self.assertTrue(self._capacidade(capacidades, chave)["disponivel"], chave)
                for chave in esperado["capacidades_ausentes"]:
                    self.assertFalse(self._capacidade(capacidades, chave)["disponivel"], chave)

                self.assertEqual(calculos["retroativo"]["valor"], resultados["retroativo"])
                self.assertEqual(calculos["vta"]["valor"], resultados["vta"])
                self.assertEqual(
                    calculos["valor_remanescente"]["valor"],
                    resultados["valor_remanescente"],
                )

                documentos_esperados = esperado["documentos_esperados"]
                if documentos_esperados == "TODOS DISPONÍVEIS":
                    self.assertTrue(
                        all(
                            capacidades["documentos"][chave]["classificacao"] == "DISPONÍVEL"
                            for chave in self.documentos
                        )
                    )
                else:
                    self.assertEqual(
                        {
                            chave: capacidades["documentos"][chave]["classificacao"]
                            for chave in self.documentos
                        },
                        documentos_esperados,
                    )

                rastreabilidade = capacidades["rastreabilidade"]
                self.assertTrue(rastreabilidade["somente_auditoria"])
                self.assertEqual(
                    rastreabilidade["resultados"]["vta"]["metodologia"],
                    esperado["metodologia_esperada"]["vta"],
                )
                self.assertIn(
                    esperado["metodologia_esperada"]["retroativo"],
                    rastreabilidade["resultados"]["retroativo"]["metodologia"],
                )
                conteudo_auditavel = json.dumps(capacidades, ensure_ascii=False)
                for mensagem in esperado["mensagens_esperadas"]:
                    self.assertIn(mensagem, conteudo_auditavel)

    def test_todo_resultado_tem_trilha_minima_de_auditoria(self):
        campos = {
            "resultado",
            "valor",
            "estado",
            "metodologia",
            "fontes_consideradas",
            "fontes_ausentes",
            "fontes_excluidas",
            "impacto",
        }
        for nome in self.cenarios:
            with self.subTest(cenario=nome):
                trilhas = self._avaliar(nome)["rastreabilidade"]["resultados"]
                self.assertEqual(
                    set(trilhas),
                    {
                        "retroativo",
                        "vta",
                        "valor_remanescente",
                        "posicao_contratual",
                        "valores_unitarios",
                    },
                )
                for trilha in trilhas.values():
                    self.assertTrue(campos.issubset(trilha))

    def test_vta_e_reproduzivel_pelos_componentes_oficiais(self):
        for nome, esperado in self.cenarios.items():
            vta_esperado = esperado["resultados_esperados"]["vta"]
            trilha = self._avaliar(nome)["rastreabilidade"]["resultados"]["vta"]
            with self.subTest(cenario=nome):
                if vta_esperado is None:
                    self.assertFalse(trilha["reproduzivel"])
                    self.assertIsNone(trilha["valor_reproduzido"])
                else:
                    self.assertTrue(trilha["reproduzivel"])
                    self.assertAlmostEqual(trilha["valor_reproduzido"], vta_esperado, places=2)

    def test_diferencas_de_vta_exigem_componentes_ou_evidencias_diferentes(self):
        trilhas = {
            nome: self._avaliar(nome)["rastreabilidade"]
            for nome in self.cenarios
        }
        nomes = list(trilhas)
        for indice, nome_a in enumerate(nomes):
            for nome_b in nomes[indice + 1:]:
                vta_a = trilhas[nome_a]["resultados"]["vta"]["valor"]
                vta_b = trilhas[nome_b]["resultados"]["vta"]["valor"]
                if vta_a != vta_b:
                    self.assertNotEqual(
                        trilhas[nome_a]["assinatura_evidencias"],
                        trilhas[nome_b]["assinatura_evidencias"],
                    )
                    self.assertNotEqual(
                        trilhas[nome_a]["resultados"]["vta"]["componentes"],
                        trilhas[nome_b]["resultados"]["vta"]["componentes"],
                    )

    def test_mesmas_evidencias_produzem_mesma_assinatura_e_mesmo_vta(self):
        for nome in self.cenarios:
            with self.subTest(cenario=nome):
                primeira = self._avaliar(nome)
                entradas = copy.deepcopy(self.cenarios[nome]["entradas"])
                entradas["contagens"] = dict(reversed(list(entradas["contagens"].items())))
                status = entradas["metadados"]["status_resultados"]
                status["valores"] = dict(reversed(list(status["valores"].items())))
                entradas["metadados"].update(
                    {
                        "caminho_local": r"C:\temporario\arquivo.xlsx",
                        "timestamp_execucao": "2099-12-31T23:59:59",
                        "sessao_aleatoria": "nao-deve-participar-da-assinatura",
                    }
                )
                segunda = avaliar_capacidades_apuracao(
                    entradas["contagens"],
                    entradas["metadados"],
                )
                self.assertEqual(
                    primeira["rastreabilidade"]["assinatura_evidencias"],
                    segunda["rastreabilidade"]["assinatura_evidencias"],
                )
                self.assertEqual(
                    primeira["calculos"]["vta"]["valor"],
                    segunda["calculos"]["vta"]["valor"],
                )

    def test_ausencia_de_fonte_nao_inventa_valor(self):
        capacidades = self._avaliar("somente_financeiro")
        trilha_vta = capacidades["rastreabilidade"]["resultados"]["vta"]
        self.assertIsNone(capacidades["calculos"]["vta"]["valor"])
        self.assertIsNone(trilha_vta["valor_reproduzido"])
        self.assertFalse(trilha_vta["reproduzivel"])
        self.assertIn("Itens remanescentes", {item["fonte"] for item in trilha_vta["fontes_ausentes"]})

    def test_mais_evidencias_nao_reduzem_confianca_nem_documentos(self):
        cadeia = (
            "somente_financeiro",
            "financeiro_remanescentes",
            "financeiro_pcs_remanescentes",
            "processo_completo",
        )
        confianca_vta = []
        documentos_habilitados = []
        for nome in cadeia:
            capacidades = self._avaliar(nome)
            confianca_vta.append(
                capacidades["rastreabilidade"]["resultados"]["vta"]["nivel_confianca"]
            )
            documentos_habilitados.append(
                sum(capacidades["documentos"][chave]["habilitado"] for chave in self.documentos)
            )
        self.assertEqual(confianca_vta, sorted(confianca_vta))
        self.assertEqual(documentos_habilitados, sorted(documentos_habilitados))

    def test_fontes_paralelas_nao_sao_somadas_ao_retroativo(self):
        capacidades = self._avaliar("financeiro_pcs_remanescentes")
        retro = capacidades["calculos"]["retroativo"]
        trilha = capacidades["rastreabilidade"]["resultados"]["retroativo"]
        self.assertEqual(retro["valor"], 40.0)
        self.assertIn("PCs", {item["fonte"] for item in trilha["fontes_excluidas"]})
        self.assertTrue(all("dupla contagem" in item["impacto"] for item in trilha["fontes_excluidas"]))

    def test_processo_completo_preserva_baseline_homologada(self):
        capacidades = self._avaliar("processo_completo")
        self.assertEqual(capacidades["calculos"]["retroativo"]["valor"], 40.0)
        self.assertEqual(capacidades["calculos"]["vta"]["valor"], 1040.0)
        self.assertEqual(capacidades["calculos"]["valor_remanescente"]["valor"], 0.0)
        self.assertTrue(
            all(capacidades["documentos"][chave]["classificacao"] == "DISPONÍVEL" for chave in self.documentos)
        )


if __name__ == "__main__":
    unittest.main()
