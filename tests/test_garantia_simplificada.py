"""Testes focais da Calculadora de Garantia Contratual (Etapa 49).

Fluxo único e 100% manual: a página compara a situação atual do contrato com a
garantia atualmente constituída e classifica o resultado em exatamente quatro
diagnósticos. Os quinze casos obrigatórios do enunciado estão cobertos aqui:
percentual padrão e manual, suficiência/insuficiência financeira e temporal (e
suas combinações), múltiplas garantias independentes, ausência de dupla contagem
de endossos, os +90 dias, os textos de comunicação, o isolamento absoluto de
qualquer dado externo em session_state e o card da Central como Calculadora.
"""
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from _garantia_calculo import (
    CLAUSULA_GARANTIA,
    DIAGNOSTICO_REGULAR,
    DIAGNOSTICO_VALIDADE,
    DIAGNOSTICO_VALOR,
    DIAGNOSTICO_VALOR_E_VALIDADE,
    DIAS_VALIDADE_MINIMA,
    PERCENTUAL_GARANTIA_PADRAO,
    SEM_NECESSIDADE_DE_ATUALIZACAO,
    analisar_garantia,
    arredondar_financeiro,
    calcular_cobertura_atual,
    calcular_complemento,
    calcular_garantia_necessaria,
    calcular_validade_minima,
    consolidar_garantias,
    formatar_brl,
    formatar_data_br,
    formatar_percentual,
    gerar_texto_comunicacao,
    moeda,
    normalizar_garantias,
    parse_data_br,
    parse_moeda_br,
)

ROOT = Path(__file__).resolve().parents[1]
GARANTIA = (ROOT / "pages" / "05_Garantia.py").read_text(encoding="utf-8")
# Código da página sem o docstring do módulo: é ele que descreve, em português,
# justamente as fontes que a página NÃO consulta (VTA, Coleta, RESULTADOS).
CORPO_GARANTIA = GARANTIA.split('"""', 2)[2]
CENTRAL = (ROOT / "pages" / "06_Central_Arquivos.py").read_text(encoding="utf-8")

FIM_VIGENCIA = date(2026, 12, 31)
VALIDADE_MINIMA = date(2027, 3, 31)   # 31/12/2026 + 90 dias corridos


def _garantia(valor, validade, identificacao=""):
    """Linha já normalizada, como a devolvida por ``normalizar_garantias``."""
    return {
        "identificacao": identificacao,
        "valor": arredondar_financeiro(parse_moeda_br(valor)),
        "validade": validade,
    }


def _analise(valor_contrato="1.000.000,00", percentual=PERCENTUAL_GARANTIA_PADRAO,
             fim_vigencia=FIM_VIGENCIA, garantias=()):
    return analisar_garantia(
        valor_total_contrato=valor_contrato,
        percentual=percentual,
        data_fim_vigencia=fim_vigencia,
        garantias=list(garantias),
    )


# ============================================================
# Conversão monetária, percentual e datas
# ============================================================

class ConversaoTests(unittest.TestCase):
    def test_formatos_monetarios_aceitos(self):
        for entrada in ("1000000", "1000000,00", "1.000.000,00", "R$ 1.000.000,00", 1000000):
            self.assertEqual(parse_moeda_br(entrada), Decimal("1000000"), entrada)

    def test_espaco_nao_separavel_e_aceito(self):
        # Valores copiados de Excel ou de páginas web trazem NBSP depois do R$.
        self.assertEqual(parse_moeda_br("R$\u00a01.000,00"), Decimal("1000.00"))
        self.assertEqual(parse_moeda_br("1.000,00\u00a0"), Decimal("1000.00"))

    def test_valores_nao_interpretaveis(self):
        for entrada in ("", "abc", None, "R$"):
            self.assertIsNone(parse_moeda_br(entrada))

    def test_formatacao_brasileira(self):
        self.assertEqual(formatar_brl(Decimal("50000")), "R$ 50.000,00")
        self.assertEqual(moeda(Decimal("1234.5")), "R$ 1.234,50")

    def test_arredondamento_half_up(self):
        self.assertEqual(arredondar_financeiro(Decimal("0.005")), Decimal("0.01"))
        self.assertEqual(arredondar_financeiro(Decimal("2.345")), Decimal("2.35"))

    def test_percentual_formatado_sem_zeros_inuteis(self):
        self.assertEqual(formatar_percentual(Decimal("5.00")), "5")
        self.assertEqual(formatar_percentual(Decimal("4.75")), "4,75")
        self.assertEqual(formatar_percentual(3.5), "3,5")

    def test_datas_aceitas_e_formatadas(self):
        self.assertEqual(parse_data_br("31/12/2026"), FIM_VIGENCIA)
        self.assertEqual(parse_data_br("2026-12-31"), FIM_VIGENCIA)
        self.assertIsNone(parse_data_br(""))
        self.assertEqual(formatar_data_br(FIM_VIGENCIA), "31/12/2026")


# ============================================================
# Casos 1 e 2 — percentual padrão e percentual manual
# ============================================================

class PercentualTests(unittest.TestCase):
    def test_caso1_percentual_padrao_e_cinco_por_cento(self):
        self.assertEqual(PERCENTUAL_GARANTIA_PADRAO, Decimal("5.00"))
        self.assertIn("PERCENTUAL_GARANTIA_PADRAO", GARANTIA)
        self.assertEqual(calcular_garantia_necessaria("1.000.000,00"), Decimal("50000.00"))
        analise = _analise()
        self.assertEqual(analise["percentual"], Decimal("5.00"))
        self.assertEqual(analise["garantia_necessaria"], Decimal("50000.00"))

    def test_caso2_percentual_manual_diferente_de_cinco(self):
        self.assertEqual(
            calcular_garantia_necessaria("1.000.000,00", Decimal("3")), Decimal("30000.00")
        )
        analise = _analise(percentual=Decimal("2.5"))
        self.assertEqual(analise["garantia_necessaria"], Decimal("25000.00"))
        self.assertEqual(formatar_percentual(analise["percentual"]), "2,5")

    def test_percentual_arredonda_meio_centavo_para_cima(self):
        # 1.000.000,10 x 5% = 50.000,005 -> 50.000,01 (ROUND_HALF_UP)
        self.assertEqual(
            calcular_garantia_necessaria("1000000,10", Decimal("5")), Decimal("50000.01")
        )


# ============================================================
# Casos 3, 4 e 13 — suficiência financeira
# ============================================================

class SuficienciaFinanceiraTests(unittest.TestCase):
    def test_caso3_cobertura_insuficiente_gera_complemento(self):
        analise = _analise(garantias=[_garantia("40.000,00", VALIDADE_MINIMA)])
        self.assertEqual(analise["garantia_necessaria"], Decimal("50000.00"))
        self.assertEqual(analise["cobertura_atual"], Decimal("40000.00"))
        self.assertEqual(analise["complemento"], Decimal("10000.00"))
        self.assertFalse(analise["valor_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALOR)

    def test_caso4_cobertura_exatamente_suficiente(self):
        analise = _analise(garantias=[_garantia("50.000,00", VALIDADE_MINIMA)])
        self.assertEqual(analise["complemento"], Decimal("0.00"))
        self.assertTrue(analise["valor_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_cobertura_superior_nao_gera_complemento_negativo(self):
        analise = _analise(garantias=[_garantia("80.000,00", VALIDADE_MINIMA)])
        self.assertEqual(analise["complemento"], Decimal("0.00"))
        self.assertGreaterEqual(analise["complemento"], 0)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_complemento_nunca_negativo(self):
        self.assertEqual(calcular_complemento(Decimal("50000"), Decimal("90000")), Decimal("0.00"))
        self.assertEqual(calcular_complemento(Decimal("50000"), Decimal("50000")), Decimal("0.00"))
        self.assertEqual(calcular_complemento(Decimal("50000"), Decimal("40000")), Decimal("10000.00"))

    def test_caso13_garantia_regular_valor_e_validade_suficientes(self):
        analise = _analise(garantias=[_garantia("50.000,00", date(2027, 6, 30))])
        self.assertTrue(analise["valor_suficiente"])
        self.assertTrue(analise["validade_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_sem_garantia_cadastrada_a_pendencia_e_apenas_de_valor(self):
        analise = _analise(garantias=[])
        self.assertEqual(analise["cobertura_atual"], Decimal("0.00"))
        self.assertEqual(analise["complemento"], Decimal("50000.00"))
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALOR)


# ============================================================
# Casos 5, 6 e 9 — validade mínima e suficiência temporal
# ============================================================

class ValidadeTests(unittest.TestCase):
    def test_caso9_validade_minima_soma_noventa_dias_corridos(self):
        self.assertEqual(DIAS_VALIDADE_MINIMA, 90)
        self.assertEqual(calcular_validade_minima(FIM_VIGENCIA), VALIDADE_MINIMA)
        self.assertEqual((VALIDADE_MINIMA - FIM_VIGENCIA).days, 90)
        # Dias corridos, não "tres meses": 31/12/2027 + 90 dias cai em 30/03/2028.
        self.assertEqual(calcular_validade_minima(date(2027, 12, 31)), date(2028, 3, 30))

    def test_caso5_validade_insuficiente_com_valor_suficiente(self):
        analise = _analise(garantias=[_garantia("50.000,00", date(2027, 1, 31))])
        self.assertTrue(analise["valor_suficiente"])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALIDADE)
        self.assertEqual(analise["complemento"], Decimal("0.00"))

    def test_validade_exatamente_na_data_minima_e_suficiente(self):
        analise = _analise(garantias=[_garantia("50.000,00", VALIDADE_MINIMA)])
        self.assertTrue(analise["validade_suficiente"])

    def test_caso6_valor_e_validade_insuficientes(self):
        analise = _analise(garantias=[_garantia("40.000,00", date(2027, 1, 31))])
        self.assertFalse(analise["valor_suficiente"])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALOR_E_VALIDADE)
        self.assertEqual(analise["complemento"], Decimal("10000.00"))

    def test_uma_garantia_curta_entre_varias_torna_a_validade_insuficiente(self):
        analise = _analise(
            garantias=[
                _garantia("30.000,00", date(2027, 6, 30), "Apolice A"),
                _garantia("20.000,00", date(2027, 1, 15), "Apolice B"),
            ]
        )
        self.assertTrue(analise["valor_suficiente"])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALIDADE)

    def test_validade_nao_informada_e_tratada_como_insuficiente(self):
        analise = _analise(garantias=[_garantia("50.000,00", None)])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALIDADE)

    def test_verificacoes_financeira_e_temporal_sao_independentes(self):
        curta, longa = date(2027, 1, 31), date(2027, 6, 30)
        combinacoes = {
            ("50.000,00", longa): DIAGNOSTICO_REGULAR,
            ("40.000,00", longa): DIAGNOSTICO_VALOR,
            ("50.000,00", curta): DIAGNOSTICO_VALIDADE,
            ("40.000,00", curta): DIAGNOSTICO_VALOR_E_VALIDADE,
        }
        for (valor, validade), esperado in combinacoes.items():
            with self.subTest(valor=valor, validade=validade):
                analise = _analise(garantias=[_garantia(valor, validade)])
                self.assertEqual(analise["diagnostico"], esperado)


# ============================================================
# Casos 7 e 8 — garantias independentes e endossos
# ============================================================

class GarantiasIndependentesTests(unittest.TestCase):
    def test_caso7_multiplas_garantias_independentes_somam(self):
        analise = _analise(
            garantias=[
                _garantia("50.000,00", date(2027, 6, 30), "Apolice A"),
                _garantia("20.000,00", date(2027, 6, 30), "Carta de fianca B"),
            ]
        )
        self.assertEqual(analise["cobertura_atual"], Decimal("70000.00"))
        self.assertEqual(analise["quantidade_garantias"], 2)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_garantias_sem_identificacao_sao_sempre_independentes(self):
        analise = _analise(
            garantias=[
                _garantia("30.000,00", date(2027, 6, 30)),
                _garantia("20.000,00", date(2027, 6, 30)),
            ]
        )
        self.assertEqual(analise["cobertura_atual"], Decimal("50000.00"))
        self.assertEqual(analise["quantidade_garantias"], 2)

    def test_caso8_endossos_da_mesma_garantia_nao_sao_somados(self):
        # Garantia original de 30.000 endossada para 40.000: cobertura = 40.000.
        analise = _analise(
            garantias=[
                _garantia("30.000,00", date(2027, 6, 30), "Apolice 123"),
                _garantia("40.000,00", date(2027, 6, 30), "Apolice 123"),
            ]
        )
        self.assertEqual(analise["cobertura_atual"], Decimal("40000.00"))
        self.assertNotEqual(analise["cobertura_atual"], Decimal("70000.00"))
        self.assertEqual(analise["quantidade_garantias"], 1)
        self.assertEqual(analise["complemento"], Decimal("10000.00"))
        self.assertTrue(analise["avisos_consolidacao"])

    def test_consolidacao_ignora_caixa_e_espacos_na_identificacao(self):
        consolidadas, avisos = consolidar_garantias(
            [
                _garantia("30.000,00", date(2027, 6, 30), " Apolice 123 "),
                _garantia("40.000,00", date(2027, 6, 30), "APOLICE 123"),
            ]
        )
        self.assertEqual(len(consolidadas), 1)
        self.assertEqual(consolidadas[0]["valor"], Decimal("40000.00"))
        self.assertEqual(len(avisos), 1)

    def test_endosso_consolidado_mantem_a_validade_mais_recente(self):
        consolidadas, _ = consolidar_garantias(
            [
                _garantia("30.000,00", date(2027, 6, 30), "Apolice 123"),
                _garantia("40.000,00", date(2027, 1, 31), "Apolice 123"),
            ]
        )
        self.assertEqual(consolidadas[0]["validade"], date(2027, 1, 31))

    def test_cobertura_de_lista_vazia_e_zero(self):
        self.assertEqual(calcular_cobertura_atual([]), Decimal("0.00"))


class NormalizacaoGarantiasTests(unittest.TestCase):
    def test_linha_vazia_ignorada_em_silencio(self):
        linhas, avisos = normalizar_garantias(
            [{"Identificação da garantia": "", "Valor total atualmente garantido": "", "Validade": None}]
        )
        self.assertEqual(linhas, [])
        self.assertEqual(avisos, [])

    def test_linha_valida_normalizada(self):
        linhas, avisos = normalizar_garantias(
            [
                {
                    "Identificação da garantia": " Apolice 123 ",
                    "Valor total atualmente garantido": "R$ 40.000,00",
                    "Validade": "31/03/2027",
                }
            ]
        )
        self.assertEqual(avisos, [])
        self.assertEqual(linhas[0]["identificacao"], "Apolice 123")
        self.assertEqual(linhas[0]["valor"], Decimal("40000.00"))
        self.assertEqual(linhas[0]["validade"], date(2027, 3, 31))

    def test_valor_ausente_ou_invalido_gera_aviso_e_descarta(self):
        linhas, avisos = normalizar_garantias(
            [
                {"Identificação da garantia": "A", "Valor total atualmente garantido": "", "Validade": None},
                {"Identificação da garantia": "B", "Valor total atualmente garantido": "abc", "Validade": None},
                {"Identificação da garantia": "C", "Valor total atualmente garantido": "-10", "Validade": None},
            ]
        )
        self.assertEqual(linhas, [])
        self.assertEqual(len(avisos), 3)

    def test_validade_ausente_mantem_a_linha_e_avisa(self):
        linhas, avisos = normalizar_garantias(
            [{"Identificação da garantia": "A", "Valor total atualmente garantido": "10.000,00", "Validade": None}]
        )
        self.assertEqual(len(linhas), 1)
        self.assertIsNone(linhas[0]["validade"])
        self.assertTrue(any("validade" in aviso.lower() for aviso in avisos))


# ============================================================
# Casos 10, 11 e 12 — textos para comunicação
# ============================================================

class TextoComunicacaoTests(unittest.TestCase):
    def test_caso10_texto_somente_de_valor(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[_garantia("40.000,00", date(2027, 6, 30))]))
        self.assertIn("Prezados,", texto)
        self.assertIn("R$ 1.000.000,00", texto)
        self.assertIn("5%", texto)
        self.assertIn("R$ 50.000,00", texto)
        self.assertIn("A garantia atualmente apresentada é de R$ 40.000,00", texto)
        self.assertIn("complementação no valor de R$ 10.000,00", texto)
        self.assertIn(CLAUSULA_GARANTIA, texto)
        self.assertIn("90 dias após o término da vigência contratual", texto)
        self.assertIn("Gentileza encaminhar o respectivo endosso/comprovante", texto)
        self.assertTrue(texto.rstrip().endswith("Atenciosamente,"))

    def test_caso11_texto_somente_de_validade_nao_menciona_complemento(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[_garantia("50.000,00", date(2027, 1, 31))]))
        self.assertIn("necessidade de atualização da validade", texto)
        self.assertIn("validade mínima até 31/03/2027", texto)
        self.assertIn(CLAUSULA_GARANTIA, texto)
        self.assertNotIn("complementação", texto)
        self.assertNotIn("R$", texto)

    def test_caso12_texto_de_valor_e_validade_menciona_ambos(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[_garantia("40.000,00", date(2027, 1, 31))]))
        self.assertIn("complementação no valor de R$ 10.000,00", texto)
        self.assertIn("atualização da sua validade", texto)
        self.assertIn("31/03/2027", texto)
        self.assertIn(CLAUSULA_GARANTIA, texto)

    def test_texto_regular_nao_pede_nada(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[_garantia("50.000,00", date(2027, 6, 30))]))
        self.assertEqual(texto, SEM_NECESSIDADE_DE_ATUALIZACAO)
        self.assertNotIn("Solicitamos", texto)

    def test_pluralizacao_com_multiplas_garantias(self):
        texto = gerar_texto_comunicacao(
            _analise(
                garantias=[
                    _garantia("20.000,00", date(2027, 6, 30), "A"),
                    _garantia("15.000,00", date(2027, 6, 30), "B"),
                ]
            )
        )
        self.assertIn("As garantias atualmente apresentadas totalizam R$ 35.000,00", texto)
        self.assertNotIn("A garantia atualmente apresentada é", texto)

    def test_texto_sem_garantia_pede_apresentacao(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[]))
        self.assertIn("Não há garantia contratual atualmente apresentada", texto)
        self.assertIn("R$ 50.000,00", texto)

    def test_percentual_manual_aparece_no_texto(self):
        texto = gerar_texto_comunicacao(
            _analise(percentual=Decimal("3"), garantias=[_garantia("10.000,00", date(2027, 6, 30))])
        )
        self.assertIn("correspondente a 3% passa a ser de R$ 30.000,00", texto)


# ============================================================
# Página: fluxo único, manual, sem PDF
# ============================================================

class PaginaFluxoUnicoTests(unittest.TestCase):
    def test_quatro_blocos_na_ordem_do_enunciado(self):
        posicoes = [
            GARANTIA.index('st.subheader("Garantia vigente")'),
            GARANTIA.index('st.subheader("Situação atual do contrato")'),
            GARANTIA.index('st.subheader("Resultado")'),
            GARANTIA.index('st.subheader("Texto para comunicação à contratada")'),
        ]
        self.assertEqual(posicoes, sorted(posicoes))

    def test_nao_existe_escolha_de_modo_nem_historico_de_instrumentos(self):
        for residuo in (
            "st.radio",
            "Valores totais",
            "Acréscimos ou reduções",
            "linha_do_tempo",
            "Memória de cálculo",
            "garantia_historico_valores_totais",
            "garantia_historico_alteracoes",
            "Instrumento",
        ):
            self.assertNotIn(residuo, GARANTIA, f"resíduo do modelo antigo: {residuo}")

    def test_pagina_nao_gera_pdf_txt_nem_download(self):
        for residuo in (
            "gerar_pdf_garantia",
            "download_button",
            "arquivo_garantia_pdf",
            "montar_txt_bytes",
            "reportlab",
            "REPORTLAB_OK",
            "application/pdf",
        ):
            self.assertNotIn(residuo, GARANTIA, f"geração de arquivo remanescente: {residuo}")
        self.assertIn("st.text_area(", GARANTIA)

    def test_motor_nao_expoe_mais_pdf_nem_vta(self):
        import _garantia_calculo as motor

        for removido in ("gerar_pdf_garantia", "extrair_vta", "montar_txt_bytes", "REPORTLAB_OK"):
            self.assertFalse(hasattr(motor, removido), f"{removido} deveria ter sido removido do motor")

    def test_campos_manuais_presentes(self):
        from _garantia_calculo import COLUNA_IDENTIFICACAO, COLUNA_VALIDADE, COLUNA_VALOR

        self.assertIn("Valor total atual do contrato", GARANTIA)
        self.assertIn("Percentual da garantia (%)", GARANTIA)
        self.assertIn("Término da vigência contratual", GARANTIA)
        self.assertEqual(COLUNA_VALOR, "Valor total atualmente garantido")
        self.assertEqual(COLUNA_IDENTIFICACAO, "Identificação da garantia")
        self.assertEqual(COLUNA_VALIDADE, "Validade")
        for constante in ("COLUNA_IDENTIFICACAO", "COLUNA_VALOR", "COLUNA_VALIDADE"):
            self.assertIn(constante, CORPO_GARANTIA)
        self.assertIn('num_rows="dynamic"', GARANTIA)

    def test_navegacao_de_retorno_preservada(self):
        self.assertIn("← Voltar para Central", GARANTIA)
        self.assertIn("st.switch_page(_destino_voltar_garantia)", GARANTIA)
        self.assertIn('st.session_state.pop("origem_navegacao_garantia", None)', GARANTIA)

    def test_pagina_publica_apenas_resultado_proprio(self):
        self.assertIn('st.session_state["resultado_garantia"]', GARANTIA)


# ============================================================
# Caso 14 — isolamento absoluto de dados externos
# ============================================================

SESSAO_EXTERNA = {
    "resultado_valor_global": {
        "valor_atualizado_contrato": 999_999_999.00,
        "valor_global_financeiro": 999_999_999.00,
        "valor_executado_atualizado": 888_888_888.00,
        "remanescente_reajustado": 777_777_777.00,
    },
    "diagnostico_coleta_v2": {"capacidades": {}, "metadados": {"indice": "IPCA"}},
    "dados_admissibilidade": {"data_base": "01/01/2020"},
    "resultado_adequacao_orcamentaria": {"total": 123_456.78},
    "input_ciclos": [{"ciclo": "C1", "data": "01/08/2024"}],
    "assinatura_processada_upload_docs": "sig-externa",
}


class IsolamentoSessionStateTests(unittest.TestCase):
    def test_pagina_nao_referencia_nenhuma_chave_externa(self):
        for chave in (
            "resultado_valor_global",
            "valor_atualizado_contrato",
            "valor_global_financeiro",
            "extrair_vta",
            "vta_claus",
            "VTA",
            "diagnostico_coleta_v2",
            "input_ciclos",
            "dados_admissibilidade",
            "Coleta",
            "RESULTADOS",
        ):
            self.assertNotIn(chave, CORPO_GARANTIA, f"a Garantia não pode consultar {chave}")

    def test_motor_nao_referencia_nenhuma_chave_externa(self):
        fonte = (ROOT / "_garantia_calculo.py").read_text(encoding="utf-8")
        for chave in ("resultado_valor_global", "valor_atualizado_contrato", "session_state", "streamlit"):
            self.assertNotIn(chave, fonte, f"o motor não pode depender de {chave}")

    def test_sessao_externa_nao_preenche_a_garantia(self):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(str(ROOT / "pages" / "05_Garantia.py"), default_timeout=60)
        for chave, valor in SESSAO_EXTERNA.items():
            at.session_state[chave] = valor
        at.run()
        self.assertFalse(at.exception)

        # Os campos manuais continuam vazios / com o default próprio.
        self.assertEqual(at.text_input(key="garantia_valor_total_contrato").value, "")
        self.assertEqual(at.number_input(key="garantia_percentual").value, 5.0)
        self.assertIsNone(at.date_input(key="garantia_fim_vigencia").value)

        # Nenhum valor externo vazou para a tela e a página segue pedindo os
        # dados manualmente (não há resultado calculado a partir da sessão).
        textos = " ".join(
            [elemento.value for elemento in at.markdown]
            + [elemento.value for elemento in at.info]
            + [elemento.value for elemento in at.warning]
        )
        for vazado in ("999.999.999", "888.888.888", "777.777.777", "123.456,78"):
            self.assertNotIn(vazado, textos)
        self.assertIn("valor total atual do contrato", " ".join(e.value for e in at.info))
        self.assertNotIn("resultado_garantia", at.session_state)

    def test_pagina_calcula_somente_com_entrada_manual(self):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(str(ROOT / "pages" / "05_Garantia.py"), default_timeout=60)
        for chave, valor in SESSAO_EXTERNA.items():
            at.session_state[chave] = valor
        at.run()
        at.text_input(key="garantia_valor_total_contrato").set_value("1.000.000,00").run()
        at.date_input(key="garantia_fim_vigencia").set_value(FIM_VIGENCIA).run()
        self.assertFalse(at.exception)

        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["valor_total_contrato"], Decimal("1000000.00"))
        self.assertEqual(resultado["garantia_necessaria"], Decimal("50000.00"))
        self.assertEqual(resultado["validade_minima"], VALIDADE_MINIMA)
        self.assertEqual(resultado["diagnostico"], DIAGNOSTICO_VALOR)


# ============================================================
# Caso 15 — card da Central de Arquivos como Calculadora
# ============================================================

class CardCentralTests(unittest.TestCase):
    def test_card_permanece_na_mesma_posicao_do_catalogo(self):
        import re

        catalogo = CENTRAL[: CENTRAL.index("def aplicar_css_central")]
        nomes = re.findall(r'"nome":\s*"([^"]+)"', catalogo)
        self.assertEqual(
            nomes,
            [
                "Sumário Executivo",
                "Adequação Orçamentária",
                "Despacho Saneador",
                "Termo de Apostila",
                "Garantia Contratual",
                "DOU",
            ],
        )

    def test_card_da_garantia_e_uma_calculadora(self):
        catalogo = CENTRAL[: CENTRAL.index("def aplicar_css_central")]
        inicio = catalogo.index('"nome": "Garantia Contratual"')
        bloco = catalogo[inicio : catalogo.index('"nome": "DOU"')]
        self.assertIn('"formato": "Calculadora"', bloco)
        self.assertIn("Calcule e confira o valor e a validade da garantia contratual.", bloco)
        self.assertIn('"pagina": "pages/05_Garantia.py"', bloco)
        self.assertIn('"ferramenta": True', bloco)
        self.assertIn('"sempre_acessivel": True', bloco)
        self.assertNotIn("arquivo_garantia_pdf", bloco)
        self.assertNotIn("application/pdf", bloco)
        self.assertIn("Abrir calculadora", CENTRAL)

    def test_demais_cards_continuam_entregando_arquivo(self):
        for chave in (
            "arquivo_sumario_executivo_pdf",
            "arquivo_adequacao_orcamentaria_xlsx",
            "arquivo_despacho_saneador_docx",
            "arquivo_termo_apostila_docx",
            "arquivo_dou_docx",
        ):
            self.assertIn(chave, CENTRAL)
        self.assertEqual(CENTRAL.count('"ferramenta": True'), 1)

    def test_card_de_ferramenta_nunca_consulta_a_sessao(self):
        # Um PDF antigo remanescente em session_state não pode transformar o
        # card em download: a flag "ferramenta" zera a consulta à sessão antes
        # de qualquer decisão de renderização.
        render = CENTRAL[CENTRAL.index("def render_documento") : CENTRAL.index("render_marca_topo()")]
        self.assertIn('ferramenta = bool(documento.get("ferramenta"))', render)
        self.assertIn(
            'arquivo = None if ferramenta else st.session_state.get(documento["session_key"])', render
        )
        self.assertIn('label="Abrir calculadora" if ferramenta else "Gerar e baixar"', render)


if __name__ == "__main__":
    unittest.main()
