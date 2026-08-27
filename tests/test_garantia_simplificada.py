"""Testes focais da Calculadora de Garantia Contratual — linha do tempo única.

A página conta a história do contrato em ordem cronológica e a garantia caminha
junto: cada linha de alteração registra o que aconteceu com o contrato E qual
passou a ser a garantia depois daquele evento. A garantia de cada linha é a
FOTOGRAFIA vigente após o evento — nunca uma parcela a somar às anteriores.

Cobertos aqui: percentual padrão e manual, suficiência financeira e temporal (e
suas combinações), fotografia sem soma, herança da fotografia anterior, os +90
dias, os textos de comunicação, o fail-closed de linha incompleta, a ausência de
qualquer campo de identificação, o isolamento absoluto de dados externos em
session_state e o card da Central como Calculadora.
"""
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from _garantia_calculo import (
    CLAUSULA_GARANTIA,
    COLUNA_EVENTO_DATA,
    COLUNA_EVENTO_GARANTIA,
    COLUNA_EVENTO_TIPO,
    COLUNA_EVENTO_VALIDADE,
    COLUNA_EVENTO_VALOR,
    COLUNA_EVENTO_VIGENCIA,
    DIAGNOSTICO_REGULAR,
    DIAGNOSTICO_VALIDADE,
    DIAGNOSTICO_VALOR,
    DIAGNOSTICO_VALOR_E_VALIDADE,
    DIAS_VALIDADE_MINIMA,
    FINANCEIRO_COMPLEMENTAR,
    FINANCEIRO_SUFICIENTE,
    FINANCEIRO_SUPERIOR,
    PERCENTUAL_GARANTIA_PADRAO,
    SEM_NECESSIDADE_DE_ATUALIZACAO,
    TEMPORAL_INSUFICIENTE,
    TEMPORAL_NAO_INFORMADA,
    TEMPORAL_SUFICIENTE,
    TIPO_ADITIVO,
    TIPO_OUTRO,
    TIPO_PRORROGACAO,
    TIPO_REAJUSTE,
    TIPO_REPACTUACAO,
    TIPOS_EVENTO,
    analisar_garantia,
    arredondar_financeiro,
    calcular_complemento,
    calcular_garantia_necessaria,
    calcular_situacao_atual,
    calcular_validade_minima,
    data_ausente,
    formatar_brl,
    formatar_brl_opcional,
    formatar_data_br,
    formatar_percentual,
    formatar_variacao,
    gerar_texto_comunicacao,
    moeda,
    montar_linha_do_tempo,
    normalizar_eventos,
    parse_data_br,
    parse_moeda_br,
)

ROOT = Path(__file__).resolve().parents[1]
GARANTIA = (ROOT / "pages" / "05_Garantia.py").read_text(encoding="utf-8")
MOTOR = (ROOT / "_garantia_calculo.py").read_text(encoding="utf-8")
# Código da página sem o docstring do módulo: é ele que descreve, em português,
# justamente as fontes que a página NÃO consulta (VTA, Coleta, RESULTADOS).
CORPO_GARANTIA = GARANTIA.split('"""', 2)[2]
CENTRAL = (ROOT / "pages" / "06_Central_Arquivos.py").read_text(encoding="utf-8")

FIM_VIGENCIA = date(2026, 12, 31)
VALIDADE_MINIMA = date(2027, 3, 31)          # 31/12/2026 + 90 dias corridos
VIGENCIA_PRORROGADA = date(2027, 12, 31)
VALIDADE_MINIMA_PRORROGADA = date(2028, 3, 30)   # 31/12/2027 + 90 dias corridos


def _evento(tipo, valor=None, vigencia=None, data=None, garantia=None, validade=None):
    """Linha crua do quadro de alterações, como o ``st.data_editor`` a devolve."""
    return {
        COLUNA_EVENTO_TIPO: tipo,
        COLUNA_EVENTO_DATA: data,
        COLUNA_EVENTO_VALOR: valor,
        COLUNA_EVENTO_VIGENCIA: vigencia,
        COLUNA_EVENTO_GARANTIA: garantia,
        COLUNA_EVENTO_VALIDADE: validade,
    }


def _situacao(valor_original="1.000.000,00", percentual=PERCENTUAL_GARANTIA_PADRAO,
              vigencia=FIM_VIGENCIA, registros=(), garantia=None, validade=None):
    eventos, avisos, pendencias = normalizar_eventos(list(registros))
    situacao = calcular_situacao_atual(
        valor_original=valor_original,
        percentual=percentual,
        fim_vigencia_original=vigencia,
        eventos=eventos,
        garantia_original=garantia,
        validade_garantia_original=validade,
    )
    return situacao, avisos, pendencias


def _analise_da_situacao(situacao):
    return analisar_garantia(
        valor_total_contrato=situacao["valor_atual"],
        percentual=situacao["percentual"],
        data_fim_vigencia=situacao["vigencia_atual"],
        garantia_apresentada=situacao["garantia_apresentada"],
        validade_apresentada=situacao["validade_apresentada"],
    )


def _analise(valor_contrato="1.000.000,00", percentual=PERCENTUAL_GARANTIA_PADRAO,
             fim_vigencia=FIM_VIGENCIA, garantia=None, validade=None):
    return analisar_garantia(
        valor_total_contrato=valor_contrato,
        percentual=percentual,
        data_fim_vigencia=fim_vigencia,
        garantia_apresentada=garantia,
        validade_apresentada=validade,
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
        self.assertEqual(parse_moeda_br("R$ 1.000,00"), Decimal("1000.00"))
        self.assertEqual(parse_moeda_br("1.000,00 "), Decimal("1000.00"))

    def test_valores_nao_interpretaveis(self):
        for entrada in ("", "abc", None, "R$"):
            self.assertIsNone(parse_moeda_br(entrada))

    def test_formatacao_brasileira(self):
        self.assertEqual(formatar_brl(Decimal("50000")), "R$ 50.000,00")
        self.assertEqual(moeda(Decimal("1234.5")), "R$ 1.234,50")

    def test_ausencia_de_valor_nunca_vira_zero(self):
        self.assertEqual(formatar_brl_opcional(None), "—")
        self.assertEqual(formatar_brl_opcional(Decimal("0")), "R$ 0,00")

    def test_arredondamento_half_up(self):
        self.assertEqual(arredondar_financeiro(Decimal("0.005")), Decimal("0.01"))
        self.assertEqual(arredondar_financeiro(Decimal("2.345")), Decimal("2.35"))

    def test_percentual_formatado_sem_zeros_inuteis(self):
        self.assertEqual(formatar_percentual(Decimal("5.00")), "5")
        self.assertEqual(formatar_percentual(Decimal("4.75")), "4,75")
        self.assertEqual(formatar_percentual(3.5), "3,5")

    def test_variacao_com_sinal_explicito(self):
        self.assertEqual(formatar_variacao(Decimal("100000")), "+ R$ 100.000,00")
        self.assertEqual(formatar_variacao(Decimal("-200000")), "- R$ 200.000,00")
        self.assertEqual(formatar_variacao(Decimal("0")), "Sem alteração")

    def test_datas_aceitas_e_formatadas(self):
        self.assertEqual(parse_data_br("31/12/2026"), FIM_VIGENCIA)
        self.assertEqual(parse_data_br("2026-12-31"), FIM_VIGENCIA)
        self.assertIsNone(parse_data_br(""))
        self.assertEqual(formatar_data_br(FIM_VIGENCIA), "31/12/2026")


# ============================================================
# Percentual padrão e percentual manual
# ============================================================

class PercentualTests(unittest.TestCase):
    def test_percentual_padrao_e_cinco_por_cento(self):
        self.assertEqual(PERCENTUAL_GARANTIA_PADRAO, Decimal("5.00"))
        self.assertIn("PERCENTUAL_GARANTIA_PADRAO", GARANTIA)
        self.assertEqual(calcular_garantia_necessaria("1.000.000,00"), Decimal("50000.00"))
        analise = _analise()
        self.assertEqual(analise["percentual"], Decimal("5.00"))
        self.assertEqual(analise["garantia_necessaria"], Decimal("50000.00"))

    def test_percentual_manual_diferente_de_cinco(self):
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

    def test_percentual_atravessa_toda_a_linha_do_tempo(self):
        situacao, _, _ = _situacao(
            percentual=Decimal("3"),
            registros=[_evento(TIPO_REAJUSTE, "2.000.000,00", garantia="40.000,00")],
        )
        self.assertEqual(situacao["garantia_original"], Decimal("30000.00"))
        self.assertEqual(situacao["linha_do_tempo"][0]["garantia_exigida"], Decimal("60000.00"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("60000.00"))

        analise = _analise_da_situacao(situacao)
        self.assertEqual(analise["complemento"], Decimal("20000.00"))
        texto = gerar_texto_comunicacao(analise)
        self.assertIn("correspondente a 3% passa a ser de R$ 60.000,00", texto)
        self.assertNotIn("5%", texto)


# ============================================================
# Suficiência financeira
# ============================================================

class SuficienciaFinanceiraTests(unittest.TestCase):
    def test_cobertura_insuficiente_gera_complemento(self):
        analise = _analise(garantia="40.000,00", validade=VALIDADE_MINIMA)
        self.assertEqual(analise["garantia_necessaria"], Decimal("50000.00"))
        self.assertEqual(analise["cobertura_atual"], Decimal("40000.00"))
        self.assertEqual(analise["complemento"], Decimal("10000.00"))
        self.assertFalse(analise["valor_suficiente"])
        self.assertEqual(analise["situacao_financeira"], FINANCEIRO_COMPLEMENTAR)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALOR)

    def test_cobertura_exatamente_suficiente(self):
        analise = _analise(garantia="50.000,00", validade=VALIDADE_MINIMA)
        self.assertEqual(analise["complemento"], Decimal("0.00"))
        self.assertTrue(analise["valor_suficiente"])
        self.assertEqual(analise["situacao_financeira"], FINANCEIRO_SUFICIENTE)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_cobertura_superior_nao_pede_nada_nem_manda_devolver(self):
        analise = _analise(garantia="80.000,00", validade=VALIDADE_MINIMA)
        self.assertEqual(analise["complemento"], Decimal("0.00"))
        self.assertEqual(analise["situacao_financeira"], FINANCEIRO_SUPERIOR)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)
        texto = gerar_texto_comunicacao(analise)
        for proibido in ("devolução", "devolver", "redução da garantia", "Solicitamos"):
            self.assertNotIn(proibido, texto)

    def test_complemento_nunca_negativo(self):
        self.assertEqual(calcular_complemento(Decimal("50000"), Decimal("90000")), Decimal("0.00"))
        self.assertEqual(calcular_complemento(Decimal("50000"), Decimal("50000")), Decimal("0.00"))
        self.assertEqual(calcular_complemento(Decimal("50000"), Decimal("40000")), Decimal("10000.00"))

    def test_sem_garantia_a_pendencia_e_apenas_de_valor(self):
        analise = _analise()
        self.assertFalse(analise["tem_garantia"])
        self.assertIsNone(analise["garantia_apresentada"])
        self.assertEqual(analise["cobertura_atual"], Decimal("0.00"))
        self.assertEqual(analise["complemento"], Decimal("50000.00"))
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_NAO_INFORMADA)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALOR)


# ============================================================
# Validade mínima e suficiência temporal
# ============================================================

class ValidadeTests(unittest.TestCase):
    def test_validade_minima_soma_noventa_dias_corridos(self):
        self.assertEqual(DIAS_VALIDADE_MINIMA, 90)
        self.assertEqual(calcular_validade_minima(FIM_VIGENCIA), VALIDADE_MINIMA)
        self.assertEqual((VALIDADE_MINIMA - FIM_VIGENCIA).days, 90)
        # Dias corridos, não "tres meses": 31/12/2027 + 90 dias cai em 30/03/2028.
        self.assertEqual(calcular_validade_minima(date(2027, 12, 31)), VALIDADE_MINIMA_PRORROGADA)

    def test_validade_insuficiente_com_valor_suficiente(self):
        analise = _analise(garantia="50.000,00", validade=date(2027, 1, 31))
        self.assertTrue(analise["valor_suficiente"])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_INSUFICIENTE)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALIDADE)
        self.assertEqual(analise["complemento"], Decimal("0.00"))

    def test_validade_exatamente_na_data_minima_e_suficiente(self):
        analise = _analise(garantia="50.000,00", validade=VALIDADE_MINIMA)
        self.assertTrue(analise["validade_suficiente"])
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_SUFICIENTE)

    def test_valor_e_validade_insuficientes(self):
        analise = _analise(garantia="40.000,00", validade=date(2027, 1, 31))
        self.assertFalse(analise["valor_suficiente"])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALOR_E_VALIDADE)
        self.assertEqual(analise["complemento"], Decimal("10000.00"))

    def test_garantia_sem_validade_conta_no_dinheiro_e_nao_no_prazo(self):
        analise = _analise(garantia="50.000,00", validade=None)
        self.assertEqual(analise["cobertura_atual"], Decimal("50000.00"))
        self.assertTrue(analise["valor_suficiente"])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_NAO_INFORMADA)
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
                self.assertEqual(_analise(garantia=valor, validade=validade)["diagnostico"], esperado)


# ============================================================
# Situação original e linha do tempo
# ============================================================

class SituacaoOriginalTests(unittest.TestCase):
    def test_sem_eventos_a_situacao_atual_e_a_original(self):
        situacao, avisos, pendencias = _situacao()
        self.assertEqual((avisos, pendencias), ([], []))
        self.assertEqual(situacao["quantidade_eventos"], 0)
        self.assertEqual(situacao["valor_atual"], Decimal("1000000.00"))
        self.assertEqual(situacao["variacao_acumulada"], Decimal("0.00"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("50000.00"))
        self.assertEqual(situacao["vigencia_atual"], FIM_VIGENCIA)
        self.assertEqual(situacao["validade_minima"], VALIDADE_MINIMA)

    def test_garantia_e_validade_da_assinatura_sao_opcionais(self):
        situacao, _, _ = _situacao()
        self.assertIsNone(situacao["garantia_apresentada_original"])
        self.assertIsNone(situacao["garantia_apresentada"])
        self.assertIsNone(situacao["validade_apresentada"])

    def test_garantia_da_assinatura_vira_a_primeira_fotografia(self):
        situacao, _, _ = _situacao(garantia="50.000,00", validade=VALIDADE_MINIMA)
        self.assertEqual(situacao["garantia_apresentada"], Decimal("50000.00"))
        self.assertEqual(situacao["validade_apresentada"], VALIDADE_MINIMA)
        self.assertEqual(_analise_da_situacao(situacao)["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_validade_sem_garantia_na_assinatura_nao_carrega_prazo_orfao(self):
        situacao, _, _ = _situacao(validade=VALIDADE_MINIMA)
        self.assertIsNone(situacao["garantia_apresentada"])
        self.assertIsNone(situacao["validade_apresentada"])

    def test_linhas_vazias_do_editor_sao_ignoradas_em_silencio(self):
        situacao, avisos, pendencias = _situacao(
            registros=[_evento(None), _evento(""), _evento(None, valor="")]
        )
        self.assertEqual((avisos, pendencias), ([], []))
        self.assertEqual(situacao["quantidade_eventos"], 0)


class LinhaDoTempoTests(unittest.TestCase):
    def test_reajuste_move_valor_variacao_e_garantia_exigida(self):
        situacao, avisos, _ = _situacao(
            registros=[_evento(TIPO_REAJUSTE, "1.100.000,00", data=date(2025, 3, 1))]
        )
        self.assertEqual(avisos, [])
        etapa = situacao["linha_do_tempo"][0]
        self.assertEqual(etapa["numero"], 1)
        self.assertEqual(etapa["valor_anterior"], Decimal("1000000.00"))
        self.assertEqual(etapa["valor"], Decimal("1100000.00"))
        self.assertEqual(etapa["variacao"], Decimal("100000.00"))
        self.assertEqual(etapa["garantia_exigida"], Decimal("55000.00"))
        self.assertEqual(situacao["valor_atual"], Decimal("1100000.00"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("55000.00"))

    def test_reajuste_sem_valor_nao_entra_e_avisa(self):
        situacao, avisos, pendencias = _situacao(registros=[_evento(TIPO_REAJUSTE)])
        self.assertEqual(situacao["quantidade_eventos"], 0)
        self.assertEqual(len(pendencias), 1)
        self.assertIn("informe o valor do contrato após este evento", avisos[0])

    def test_prorrogacao_sem_valor_preserva_valor_e_recalcula_validade(self):
        situacao, avisos, _ = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, "1.100.000,00"),
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA),
            ]
        )
        self.assertEqual(avisos, [])
        prorrogacao = situacao["linha_do_tempo"][1]
        self.assertFalse(prorrogacao["valor_informado"])
        self.assertEqual(prorrogacao["valor"], Decimal("1100000.00"))
        self.assertEqual(prorrogacao["variacao"], Decimal("0.00"))
        self.assertEqual(prorrogacao["garantia_exigida"], Decimal("55000.00"))
        self.assertEqual(situacao["valor_atual"], Decimal("1100000.00"))
        self.assertEqual(situacao["vigencia_atual"], VIGENCIA_PRORROGADA)
        self.assertEqual(situacao["validade_minima"], VALIDADE_MINIMA_PRORROGADA)

    def test_prorrogacao_pode_alterar_o_valor_quando_informado(self):
        situacao, avisos, _ = _situacao(
            registros=[_evento(TIPO_PRORROGACAO, "1.200.000,00", VIGENCIA_PRORROGADA)]
        )
        self.assertEqual(avisos, [])
        self.assertEqual(situacao["valor_atual"], Decimal("1200000.00"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("60000.00"))
        self.assertEqual(situacao["vigencia_atual"], VIGENCIA_PRORROGADA)

    def test_prorrogacao_sem_nova_vigencia_nao_entra_e_avisa(self):
        situacao, avisos, pendencias = _situacao(registros=[_evento(TIPO_PRORROGACAO)])
        self.assertEqual(situacao["quantidade_eventos"], 0)
        self.assertEqual(len(pendencias), 1)
        self.assertIn("novo término da vigência", avisos[0])

    def test_reducao_de_valor(self):
        situacao, _, _ = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, "1.100.000,00"),
                _evento(TIPO_ADITIVO, "900.000,00"),
            ]
        )
        self.assertEqual(situacao["linha_do_tempo"][1]["variacao"], Decimal("-200000.00"))
        self.assertEqual(situacao["valor_atual"], Decimal("900000.00"))
        self.assertEqual(situacao["variacao_acumulada"], Decimal("-100000.00"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("45000.00"))

    def test_sequencia_automatica_na_ordem_de_insercao(self):
        situacao, avisos, _ = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, "1.100.000,00", data=date(2025, 3, 1)),
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA, data=date(2026, 11, 20)),
                _evento(TIPO_ADITIVO, "1.300.000,00", data=date(2027, 2, 10)),
            ]
        )
        self.assertEqual(avisos, [])
        self.assertEqual([e["numero"] for e in situacao["linha_do_tempo"]], [1, 2, 3])
        self.assertEqual(situacao["valor_atual"], Decimal("1300000.00"))
        self.assertEqual(situacao["vigencia_atual"], VIGENCIA_PRORROGADA)
        self.assertEqual(situacao["garantia_exigida"], Decimal("65000.00"))

    def test_ordem_de_insercao_nunca_e_reordenada_pela_data(self):
        situacao, _, _ = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, "1.100.000,00", data=date(2027, 5, 1)),
                _evento(TIPO_ADITIVO, "1.300.000,00", data=date(2025, 1, 1)),
            ]
        )
        self.assertEqual([e["tipo"] for e in situacao["linha_do_tempo"]], [TIPO_REAJUSTE, TIPO_ADITIVO])
        self.assertEqual(situacao["valor_atual"], Decimal("1300000.00"))

    def test_excluir_evento_anterior_recalcula_toda_a_cadeia(self):
        situacao, _, _ = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, "1.100.000,00"),
                _evento(TIPO_ADITIVO, "1.300.000,00"),
            ]
        )
        self.assertEqual([e["numero"] for e in situacao["linha_do_tempo"]], [1, 2])
        self.assertEqual(situacao["valor_atual"], Decimal("1300000.00"))
        self.assertEqual(situacao["vigencia_atual"], FIM_VIGENCIA)   # sem prorrogação

    def test_alterar_evento_anterior_reflete_nos_posteriores(self):
        editado, _, _ = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, "1.500.000,00"),
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA),
            ]
        )
        self.assertEqual(editado["valor_atual"], Decimal("1500000.00"))
        self.assertEqual(editado["linha_do_tempo"][1]["garantia_exigida"], Decimal("75000.00"))

    def test_todos_os_cinco_tipos_sao_aceitos(self):
        self.assertEqual(
            TIPOS_EVENTO,
            (TIPO_REAJUSTE, TIPO_REPACTUACAO, TIPO_ADITIVO, TIPO_PRORROGACAO, TIPO_OUTRO),
        )
        situacao, avisos, _ = _situacao(
            registros=[
                _evento(TIPO_REPACTUACAO, "1.050.000,00"),
                _evento(TIPO_OUTRO, garantia="52.500,00"),
            ]
        )
        self.assertEqual(avisos, [])
        self.assertEqual(situacao["quantidade_eventos"], 2)


# ============================================================
# Garantia por FOTOGRAFIA — nunca soma, herda quando ausente
# ============================================================

class FotografiaDaGarantiaTests(unittest.TestCase):
    def test_garantias_de_eventos_sucessivos_nao_somam(self):
        situacao, _, _ = _situacao(
            garantia="50.000,00", validade=VALIDADE_MINIMA,
            registros=[
                _evento(TIPO_REAJUSTE, "1.100.000,00", garantia="55.000,00",
                        validade=VALIDADE_MINIMA),
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA, garantia="55.000,00",
                        validade=VALIDADE_MINIMA_PRORROGADA),
            ],
        )
        self.assertEqual(situacao["garantia_apresentada"], Decimal("55000.00"))
        self.assertNotEqual(situacao["garantia_apresentada"], Decimal("160000.00"))
        self.assertEqual(situacao["validade_apresentada"], VALIDADE_MINIMA_PRORROGADA)
        analise = _analise_da_situacao(situacao)
        self.assertEqual(analise["cobertura_atual"], Decimal("55000.00"))
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_ultima_fotografia_informada_prevalece(self):
        situacao, _, _ = _situacao(
            garantia="50.000,00", validade=VALIDADE_MINIMA,
            registros=[
                _evento(TIPO_REAJUSTE, "1.100.000,00", garantia="52.000,00"),
                _evento(TIPO_ADITIVO, "1.200.000,00", garantia="60.000,00"),
            ],
        )
        self.assertEqual(situacao["garantia_apresentada"], Decimal("60000.00"))
        etapas = situacao["linha_do_tempo"]
        self.assertEqual(etapas[0]["garantia_apresentada"], Decimal("52000.00"))
        self.assertEqual(etapas[0]["garantia_apresentada_anterior"], Decimal("50000.00"))
        self.assertEqual(etapas[1]["garantia_apresentada"], Decimal("60000.00"))

    def test_evento_sem_garantia_herda_a_fotografia_anterior(self):
        situacao, _, _ = _situacao(
            garantia="50.000,00", validade=VALIDADE_MINIMA,
            registros=[
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA),
                _evento(TIPO_ADITIVO, "1.100.000,00"),
            ],
        )
        for etapa in situacao["linha_do_tempo"]:
            self.assertFalse(etapa["garantia_informada"])
            self.assertEqual(etapa["garantia_apresentada"], Decimal("50000.00"))
            self.assertEqual(etapa["validade_apresentada"], VALIDADE_MINIMA)
        self.assertEqual(situacao["garantia_apresentada"], Decimal("50000.00"))
        self.assertEqual(situacao["validade_apresentada"], VALIDADE_MINIMA)

    def test_nova_garantia_com_validade_vazia_nao_herda_validade_antiga(self):
        situacao, _, _ = _situacao(
            garantia="50.000,00", validade=VALIDADE_MINIMA,
            registros=[_evento(TIPO_REAJUSTE, "1.100.000,00", garantia="55.000,00")],
        )
        self.assertEqual(situacao["garantia_apresentada"], Decimal("55000.00"))
        self.assertIsNone(situacao["validade_apresentada"])
        analise = _analise_da_situacao(situacao)
        self.assertTrue(analise["valor_suficiente"])            # dinheiro considerado
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_NAO_INFORMADA)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALIDADE)

    def test_validade_sem_valor_de_garantia_pede_o_valor(self):
        situacao, avisos, pendencias = _situacao(
            registros=[_evento(TIPO_REAJUSTE, "1.100.000,00", validade=VALIDADE_MINIMA)]
        )
        self.assertEqual(situacao["quantidade_eventos"], 0)
        self.assertEqual(len(pendencias), 1)
        self.assertIn("informe também o valor da garantia vigente", avisos[0])

    def test_garantia_negativa_ou_ilegivel_e_pendencia(self):
        for entrada in ("-10", "abc"):
            with self.subTest(entrada=entrada):
                _, _, pendencias = _situacao(
                    registros=[_evento(TIPO_REAJUSTE, "1.100.000,00", garantia=entrada)]
                )
                self.assertEqual(len(pendencias), 1)

    def test_exemplo_do_enunciado_com_dois_eventos(self):
        # Assinatura 1.000.000/50.000 -> Reajuste 1.100.000/55.000
        # -> Prorrogação (valor mantido)/55.000 com nova validade.
        situacao, avisos, _ = _situacao(
            garantia="50.000,00", validade=VALIDADE_MINIMA,
            registros=[
                _evento(TIPO_REAJUSTE, "1.100.000,00", garantia="55.000,00"),
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA, garantia="55.000,00",
                        validade=VALIDADE_MINIMA_PRORROGADA),
            ],
        )
        self.assertEqual(avisos, [])
        self.assertEqual(situacao["valor_atual"], Decimal("1100000.00"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("55000.00"))
        self.assertEqual(situacao["garantia_apresentada"], Decimal("55000.00"))
        analise = _analise_da_situacao(situacao)
        self.assertEqual(analise["situacao_financeira"], FINANCEIRO_SUFICIENTE)
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_SUFICIENTE)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)


# ============================================================
# Ausência de data e NaT
# ============================================================

class DataAusenteTests(unittest.TestCase):
    def test_nat_e_instancia_de_datetime_mas_e_ausencia(self):
        import pandas as pd

        self.assertTrue(data_ausente(pd.NaT))
        self.assertIsNone(parse_data_br(pd.NaT))
        analise = _analise(garantia="50.000,00", validade=pd.NaT)
        self.assertIsNone(analise["validade_apresentada"])
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_NAO_INFORMADA)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALIDADE)

    def test_nat_do_editor_nao_derruba_a_linha_do_tempo(self):
        import pandas as pd

        grade = pd.DataFrame(
            {
                COLUNA_EVENTO_TIPO: pd.Series([TIPO_REAJUSTE], dtype="object"),
                COLUNA_EVENTO_DATA: pd.Series([pd.NaT], dtype="datetime64[ns]"),
                COLUNA_EVENTO_VALOR: pd.Series(["1.100.000,00"], dtype="object"),
                COLUNA_EVENTO_VIGENCIA: pd.Series([pd.NaT], dtype="datetime64[ns]"),
                COLUNA_EVENTO_GARANTIA: pd.Series(["55.000,00"], dtype="object"),
                COLUNA_EVENTO_VALIDADE: pd.Series([pd.NaT], dtype="datetime64[ns]"),
            }
        )
        situacao, avisos, pendencias = _situacao(registros=grade.to_dict("records"))
        self.assertEqual((avisos, pendencias), ([], []))
        etapa = situacao["linha_do_tempo"][0]
        self.assertIsNone(etapa["data"])
        self.assertEqual(etapa["vigencia"], FIM_VIGENCIA)     # herdada
        self.assertEqual(situacao["garantia_apresentada"], Decimal("55000.00"))
        self.assertIsNone(situacao["validade_apresentada"])
        self.assertEqual(_analise_da_situacao(situacao)["situacao_temporal"], TEMPORAL_NAO_INFORMADA)

    def test_none_e_as_demais_formas_de_ausencia(self):
        for ausente in (None, float("nan"), "", "   ", "NaT", "nan", "None", "<NA>"):
            with self.subTest(ausente=ausente):
                self.assertTrue(data_ausente(ausente))
                self.assertIsNone(parse_data_br(ausente))

    def test_timestamp_valido_compara_normalmente(self):
        import pandas as pd

        self.assertFalse(data_ausente(pd.Timestamp("2027-06-30")))
        self.assertEqual(parse_data_br(pd.Timestamp("2027-06-30")), date(2027, 6, 30))
        analise = _analise(garantia="50.000,00", validade=pd.Timestamp("2027-06-30"))
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_SUFICIENTE)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_datas_validas_nunca_sao_confundidas_com_ausencia(self):
        from datetime import datetime as _datetime

        for presente in (FIM_VIGENCIA, _datetime(2026, 12, 31, 23, 59), "31/12/2026"):
            with self.subTest(presente=presente):
                self.assertFalse(data_ausente(presente))
                self.assertEqual(parse_data_br(presente), FIM_VIGENCIA)


# ============================================================
# Pendências — fail-closed
# ============================================================

class PendenciasNoMotorTests(unittest.TestCase):
    def test_todo_aviso_de_evento_e_pendencia(self):
        _, avisos, pendencias = _situacao(
            registros=[_evento(TIPO_REAJUSTE), _evento(TIPO_PRORROGACAO)]
        )
        self.assertEqual(len(pendencias), 2)
        self.assertEqual(avisos, pendencias)

    def test_evento_sem_tipo_e_pendencia(self):
        _, avisos, pendencias = _situacao(registros=[_evento(None, valor="1.100.000,00")])
        self.assertEqual(len(pendencias), 1)
        self.assertIn("selecione o tipo do evento", avisos[0])

    def test_evento_valido_nao_gera_pendencia(self):
        _, avisos, pendencias = _situacao(registros=[_evento(TIPO_REAJUSTE, "1.100.000,00")])
        self.assertEqual((avisos, pendencias), ([], []))


# ============================================================
# Texto à contratada
# ============================================================

class TextoComunicacaoTests(unittest.TestCase):
    def test_texto_comeca_direto_em_prezados_sem_referencia(self):
        texto = gerar_texto_comunicacao(_analise(garantia="40.000,00", validade=date(2027, 6, 30)))
        self.assertTrue(texto.startswith("Prezados,"))
        for proibido in ("Ref.:", "Contrato nº", "Contratada"):
            self.assertNotIn(proibido, texto)

    def test_sem_garantia_apresentada_indica_a_garantia_total(self):
        texto = gerar_texto_comunicacao(_analise())
        self.assertIn("Não há garantia contratual atualmente apresentada", texto)
        self.assertIn("R$ 50.000,00", texto)
        self.assertIn(CLAUSULA_GARANTIA, texto)

    def test_insuficiencia_financeira_indica_o_complemento(self):
        texto = gerar_texto_comunicacao(_analise(garantia="40.000,00", validade=date(2027, 6, 30)))
        self.assertIn("R$ 1.000.000,00", texto)
        self.assertIn("5%", texto)
        self.assertIn("A garantia atualmente apresentada é de R$ 40.000,00", texto)
        self.assertIn("complementação no valor de R$ 10.000,00", texto)
        self.assertIn("90 dias após o término da vigência contratual", texto)
        self.assertTrue(texto.rstrip().endswith("Atenciosamente,"))

    def test_apenas_validade_nao_inventa_complemento(self):
        texto = gerar_texto_comunicacao(_analise(garantia="50.000,00", validade=date(2027, 1, 31)))
        self.assertIn("necessidade de atualização da validade", texto)
        self.assertIn("validade mínima até 31/03/2027", texto)
        self.assertNotIn("complementação", texto)
        self.assertNotIn("R$", texto)

    def test_valor_e_validade_menciona_ambos(self):
        texto = gerar_texto_comunicacao(_analise(garantia="40.000,00", validade=date(2027, 1, 31)))
        self.assertIn("complementação no valor de R$ 10.000,00", texto)
        self.assertIn("atualização da sua validade", texto)
        self.assertIn("31/03/2027", texto)

    def test_tudo_suficiente_nao_declara_aceitacao_juridica(self):
        texto = gerar_texto_comunicacao(_analise(garantia="50.000,00", validade=date(2027, 6, 30)))
        self.assertEqual(texto, SEM_NECESSIDADE_DE_ATUALIZACAO)
        self.assertIn("Com os dados apresentados", texto)
        for proibido in ("aceita", "aprovada", "homologada", "Solicitamos"):
            self.assertNotIn(proibido, texto)

    def test_texto_menciona_a_vigencia_atual_apurada(self):
        situacao, _, _ = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, "1.100.000,00"),
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA),
            ]
        )
        texto = gerar_texto_comunicacao(_analise_da_situacao(situacao))
        self.assertIn("R$ 1.100.000,00", texto)
        self.assertIn("encerrada em 31/12/2027", texto)


# ============================================================
# Página: estrutura, ausência de identificação e fail-closed
# ============================================================

class PaginaEstruturaTests(unittest.TestCase):
    def test_cinco_blocos_na_ordem_cronologica(self):
        posicoes = [
            GARANTIA.index('st.subheader("Situação original do contrato")'),
            GARANTIA.index('st.subheader("Alterações posteriores à assinatura")'),
            GARANTIA.index('st.subheader("Situação atual do contrato")'),
            GARANTIA.index('st.subheader("Resultado da análise")'),
            GARANTIA.index('st.subheader("Texto para a contratada")'),
        ]
        self.assertEqual(posicoes, sorted(posicoes))

    def test_nenhum_campo_de_identificacao_existe(self):
        # Rótulos, chaves de widget e parâmetros — a palavra "contratada" segue
        # legítima na prosa ("Texto para a contratada"), o que não pode existir
        # é campo destinado a identificar o caso.
        for residuo in (
            'st.subheader("Identificação")',
            "Número do contrato",
            'st.text_input(\n        "Contratada"',
            "garantia_numero_contrato",
            "garantia_contratada",
            "numero_contrato",
            "_linha_referencia",
            "Ref.:",
            "Apólice",
            "endosso / referência",
            "COLUNA_REFERENCIA",
        ):
            self.assertNotIn(residuo, GARANTIA, f"campo de identificação remanescente: {residuo}")
            self.assertNotIn(residuo, MOTOR, f"campo de identificação remanescente no motor: {residuo}")
        # Nenhum widget de entrada além dos da situação original e da grade.
        self.assertEqual(GARANTIA.count("st.text_input("), 2)
        self.assertEqual(GARANTIA.count("st.date_input("), 2)

    def test_reajuste_apostila_nao_existe_mais(self):
        self.assertEqual(TIPO_REAJUSTE, "Reajuste")
        for fonte in (GARANTIA, MOTOR):
            self.assertNotIn("Reajuste / Apostila", fonte)
            self.assertNotIn("Apostila", fonte)

    def test_nao_existe_bloco_de_garantia_apresentada_separado(self):
        for residuo in (
            'st.subheader("Garantia atualmente apresentada")',
            "Garantia atualmente apresentada",
            "garantia_vigente_linhas",
            "normalizar_garantias",
            "consolidar_garantias",
            "calcular_cobertura_atual",
        ):
            self.assertNotIn(residuo, GARANTIA, f"resíduo do quadro separado: {residuo}")
        # Uma única grade na página inteira.
        self.assertEqual(GARANTIA.count("st.data_editor("), 1)
        self.assertIn("garantia_eventos_contrato", GARANTIA)

    def test_motor_nao_expoe_mais_o_quadro_separado(self):
        import _garantia_calculo as motor

        for removido in (
            "normalizar_garantias", "consolidar_garantias", "calcular_cobertura_atual",
            "COLUNA_REFERENCIA", "gerar_pdf_garantia", "extrair_vta",
        ):
            self.assertFalse(hasattr(motor, removido), f"{removido} deveria ter sido removido")

    def test_colunas_da_grade_unica(self):
        self.assertEqual(COLUNA_EVENTO_TIPO, "Tipo")
        self.assertEqual(COLUNA_EVENTO_DATA, "Data")
        self.assertEqual(COLUNA_EVENTO_VALOR, "Valor do contrato após o evento")
        self.assertEqual(COLUNA_EVENTO_VIGENCIA, "Novo término da vigência")
        self.assertEqual(COLUNA_EVENTO_GARANTIA, "Garantia após o evento")
        self.assertEqual(COLUNA_EVENTO_VALIDADE, "Validade da garantia")
        for constante in (
            "COLUNA_EVENTO_TIPO", "COLUNA_EVENTO_DATA", "COLUNA_EVENTO_VALOR",
            "COLUNA_EVENTO_VIGENCIA", "COLUNA_EVENTO_GARANTIA", "COLUNA_EVENTO_VALIDADE",
            "TIPOS_EVENTO",
        ):
            self.assertIn(constante, CORPO_GARANTIA)
        self.assertNotIn("Observação", GARANTIA)
        self.assertIn('num_rows="dynamic"', GARANTIA)

    def test_campos_da_situacao_original(self):
        self.assertIn("Valor original do contrato", GARANTIA)
        self.assertIn("Percentual da garantia (%)", GARANTIA)
        self.assertIn("Término da vigência original", GARANTIA)
        self.assertIn("Garantia apresentada na assinatura", GARANTIA)
        self.assertIn("Validade da garantia", GARANTIA)

    def test_nenhum_widget_da_pagina_usa_help(self):
        # As bolinhas de ajuda ao lado dos rótulos vinham do argumento help=.
        self.assertNotIn("help=", GARANTIA)

    def test_frase_global_de_pendencia_nao_existe(self):
        for proibido in (
            "Há alteração contratual com dados pendentes",
            "Complete ou remova a linha",
            "Há garantia apresentada com dados pendentes",
        ):
            self.assertNotIn(proibido, GARANTIA, f"alerta global proibido: {proibido}")

    def test_pagina_nao_gera_pdf_txt_nem_download(self):
        for residuo in (
            "gerar_pdf_garantia", "download_button", "arquivo_garantia_pdf",
            "montar_txt_bytes", "reportlab", "REPORTLAB_OK", "application/pdf",
        ):
            self.assertNotIn(residuo, GARANTIA, f"geração de arquivo remanescente: {residuo}")
        self.assertIn("st.text_area(", GARANTIA)

    def test_situacao_atual_e_derivada_e_nao_redigitada(self):
        self.assertIn("calcular_situacao_atual(", GARANTIA)
        self.assertIn('valor_total_contrato=situacao["valor_atual"]', GARANTIA)
        self.assertIn('data_fim_vigencia=situacao["vigencia_atual"]', GARANTIA)
        self.assertIn('garantia_apresentada=situacao["garantia_apresentada"]', GARANTIA)
        self.assertIn('validade_apresentada=situacao["validade_apresentada"]', GARANTIA)

    def test_navegacao_de_retorno_preservada(self):
        self.assertIn("← Voltar para Central", GARANTIA)
        self.assertIn("st.switch_page(_destino_voltar_garantia)", GARANTIA)
        self.assertIn('st.session_state.pop("origem_navegacao_garantia", None)', GARANTIA)

    def test_pagina_publica_apenas_resultado_proprio(self):
        self.assertIn('st.session_state["resultado_garantia"]', GARANTIA)

    def test_texto_da_contratada_nao_congela_na_primeira_apuracao(self):
        # ARMADILHA: st.text_area com key fixa só honra value= no primeiro
        # render; sem a ressincronia por session_state o texto congela na
        # primeira apuração e passa a mentir sobre o resultado exibido acima.
        bloco = GARANTIA[GARANTIA.index("Texto para a contratada"):]
        self.assertIn(
            'if st.session_state.get("garantia_texto_comunicacao") != texto_comunicacao:', bloco
        )
        inicio = bloco.index("st.text_area(")
        self.assertNotIn("value=", bloco[inicio:inicio + 260], "value= com key fixa congela o texto")


# ============================================================
# Isolamento absoluto de dados externos
# ============================================================

SESSAO_EXTERNA = {
    "resultado_valor_global": {
        "valor_atualizado_contrato": 999_999_999.00,
        "valor_global_financeiro": 999_999_999.00,
    },
    "diagnostico_coleta_v2": {"capacidades": {}, "metadados": {"indice": "IPCA"}},
    "dados_admissibilidade": {"data_base": "01/01/2020"},
    "resultado_adequacao_orcamentaria": {"total": 123_456.78},
    "input_ciclos": [{"ciclo": "C1", "data": "01/08/2024"}],
}


class IsolamentoSessionStateTests(unittest.TestCase):
    def test_pagina_nao_referencia_nenhuma_chave_externa(self):
        for chave in (
            "resultado_valor_global", "valor_atualizado_contrato", "valor_global_financeiro",
            "extrair_vta", "vta_claus", "VTA", "diagnostico_coleta_v2", "input_ciclos",
            "dados_admissibilidade", "Coleta", "RESULTADOS",
        ):
            self.assertNotIn(chave, CORPO_GARANTIA, f"a Garantia não pode consultar {chave}")

    def test_motor_nao_referencia_nenhuma_chave_externa(self):
        for chave in ("resultado_valor_global", "valor_atualizado_contrato", "session_state", "streamlit"):
            self.assertNotIn(chave, MOTOR, f"o motor não pode depender de {chave}")

    def test_sessao_externa_nao_preenche_a_garantia(self):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(str(ROOT / "pages" / "05_Garantia.py"), default_timeout=90)
        for chave, valor in SESSAO_EXTERNA.items():
            at.session_state[chave] = valor
        at.run()
        self.assertFalse(at.exception)

        self.assertEqual(at.text_input(key="garantia_valor_original").value, "")
        self.assertEqual(at.text_input(key="garantia_apresentada_original").value, "")
        self.assertEqual(at.number_input(key="garantia_percentual").value, 5.0)
        self.assertIsNone(at.date_input(key="garantia_fim_vigencia").value)

        textos = " ".join(
            [e.value for e in at.markdown] + [e.value for e in at.info] + [e.value for e in at.warning]
        )
        for vazado in ("999.999.999", "123.456,78"):
            self.assertNotIn(vazado, textos)
        self.assertIn("valor original do contrato", " ".join(e.value for e in at.info))
        self.assertNotIn("resultado_garantia", at.session_state)


class PaginaFluxoTests(unittest.TestCase):
    """A página em execução: conclusão, fail-closed e ressincronia do texto."""

    def _linha(self, tipo=None, valor=None, vigencia=None, garantia=None, validade=None):
        return {
            COLUNA_EVENTO_TIPO: tipo,
            COLUNA_EVENTO_DATA: None,
            COLUNA_EVENTO_VALOR: valor,
            COLUNA_EVENTO_VIGENCIA: None if vigencia is None else vigencia.isoformat(),
            COLUNA_EVENTO_GARANTIA: garantia,
            COLUNA_EVENTO_VALIDADE: None if validade is None else validade.isoformat(),
        }

    def _rodar(self, eventos=(), garantia="", validade=None, percentual=None):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(str(ROOT / "pages" / "05_Garantia.py"), default_timeout=90)
        at.run()
        at.text_input(key="garantia_valor_original").set_value("1.000.000,00").run()
        at.date_input(key="garantia_fim_vigencia").set_value(FIM_VIGENCIA).run()
        if garantia:
            at.text_input(key="garantia_apresentada_original").set_value(garantia).run()
        if validade is not None:
            at.date_input(key="garantia_validade_original").set_value(validade).run()
        if percentual is not None:
            at.number_input(key="garantia_percentual").set_value(percentual).run()
        if eventos:
            at.session_state["garantia_eventos_contrato"] = {
                "edited_rows": {}, "added_rows": list(eventos), "deleted_rows": [0, 1, 2],
            }
        at.run()
        self.assertFalse(at.exception)
        return at

    def test_pagina_abre_vazia_sem_excecao(self):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(str(ROOT / "pages" / "05_Garantia.py"), default_timeout=90)
        at.run()
        self.assertFalse(at.exception)
        self.assertIn("valor original do contrato", " ".join(e.value for e in at.info))

    def test_somente_situacao_original_ja_conclui(self):
        at = self._rodar()
        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["valor_total_contrato"], Decimal("1000000.00"))
        self.assertEqual(resultado["garantia_necessaria"], Decimal("50000.00"))
        self.assertEqual(resultado["quantidade_eventos"], 0)
        self.assertFalse(resultado["tem_garantia"])
        self.assertEqual(resultado["situacao_temporal"], TEMPORAL_NAO_INFORMADA)

    def test_garantia_da_assinatura_conclui_regular(self):
        at = self._rodar(garantia="50.000,00", validade=VALIDADE_MINIMA)
        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["garantia_apresentada"], Decimal("50000.00"))
        self.assertEqual(resultado["diagnostico"], DIAGNOSTICO_REGULAR)
        self.assertEqual(resultado["texto_comunicacao"], SEM_NECESSIDADE_DE_ATUALIZACAO)

    def test_evento_valido_com_garantia_continua_calculando(self):
        at = self._rodar(
            garantia="50.000,00", validade=VALIDADE_MINIMA,
            eventos=[self._linha(TIPO_REAJUSTE, "1.100.000,00", garantia="55.000,00",
                                 validade=VALIDADE_MINIMA)],
        )
        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["quantidade_eventos"], 1)
        self.assertEqual(resultado["valor_total_contrato"], Decimal("1100000.00"))
        self.assertEqual(resultado["garantia_necessaria"], Decimal("55000.00"))
        self.assertEqual(resultado["garantia_apresentada"], Decimal("55000.00"))
        self.assertEqual(resultado["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_linha_incompleta_e_fail_closed_sem_alerta_global(self):
        at = self._rodar(eventos=[self._linha(TIPO_REAJUSTE)])
        avisos = [e.value for e in at.warning]
        self.assertTrue(any("informe o valor do contrato após este evento" in a for a in avisos))
        # Só o aviso da própria linha: nenhum alerta global acrescentado.
        self.assertEqual([e.value for e in at.info], [])
        self.assertNotIn("resultado_garantia", at.session_state)

    def test_conclusao_anterior_nao_sobrevive_a_pendencia(self):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(str(ROOT / "pages" / "05_Garantia.py"), default_timeout=90)
        at.run()
        at.text_input(key="garantia_valor_original").set_value("1.000.000,00").run()
        at.date_input(key="garantia_fim_vigencia").set_value(FIM_VIGENCIA).run()
        self.assertIn("resultado_garantia", at.session_state)

        at.session_state["garantia_eventos_contrato"] = {
            "edited_rows": {}, "added_rows": [self._linha(TIPO_REAJUSTE)], "deleted_rows": [0, 1, 2],
        }
        at.run()
        self.assertFalse(at.exception)
        self.assertNotIn("resultado_garantia", at.session_state)

    def test_a_garantia_e_validade_vazias_na_assinatura_sao_permitidas(self):
        # Ambas vazias significam simplesmente que não havia garantia na
        # assinatura: a análise conclui normalmente.
        at = self._rodar()
        resultado = at.session_state["resultado_garantia"]
        self.assertFalse(resultado["tem_garantia"])
        self.assertIsNone(resultado["garantia_apresentada"])
        self.assertEqual(resultado["situacao_temporal"], TEMPORAL_NAO_INFORMADA)
        self.assertEqual([e.value for e in at.warning if "assinatura" in e.value], [])

    def test_b_garantia_sem_validade_na_assinatura_e_permitida(self):
        at = self._rodar(garantia="50.000,00")
        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["garantia_apresentada"], Decimal("50000.00"))
        self.assertTrue(resultado["valor_suficiente"])          # dinheiro considerado
        self.assertIsNone(resultado["validade_apresentada"])
        self.assertEqual(resultado["situacao_temporal"], TEMPORAL_NAO_INFORMADA)
        self.assertEqual(resultado["diagnostico"], DIAGNOSTICO_VALIDADE)

    def test_c_validade_sem_garantia_na_assinatura_e_fail_closed(self):
        at = self._rodar(validade=VALIDADE_MINIMA)
        avisos = [e.value for e in at.warning]
        self.assertIn("Informe também o valor da garantia apresentada na assinatura.", avisos)
        # Só o aviso específico: nenhum alerta global acrescentado.
        self.assertEqual([e.value for e in at.info], [])
        self.assertNotIn("resultado_garantia", at.session_state)
        # Os campos preenchidos permanecem na tela.
        self.assertEqual(at.date_input(key="garantia_validade_original").value, VALIDADE_MINIMA)
        self.assertEqual(at.text_input(key="garantia_valor_original").value, "1.000.000,00")

    def test_c_conclusao_anterior_nao_sobrevive_a_validade_orfa(self):
        at = self._rodar(garantia="50.000,00", validade=VALIDADE_MINIMA)
        self.assertIn("resultado_garantia", at.session_state)
        at.text_input(key="garantia_apresentada_original").set_value("").run()
        self.assertFalse(at.exception)
        self.assertNotIn("resultado_garantia", at.session_state)

    def test_linha_vazia_continua_ignorada_e_nao_bloqueia(self):
        at = self._rodar(eventos=[self._linha()])
        self.assertEqual([e.value for e in at.warning if e.value.startswith("Linha ")], [])
        self.assertEqual([e.value for e in at.info], [])
        self.assertEqual(at.session_state["resultado_garantia"]["quantidade_eventos"], 0)

    def test_texto_da_contratada_acompanha_o_novo_resultado(self):
        at = self._rodar(garantia="50.000,00", validade=VALIDADE_MINIMA)
        self.assertEqual(
            at.session_state["resultado_garantia"]["texto_comunicacao"],
            SEM_NECESSIDADE_DE_ATUALIZACAO,
        )
        # Um reajuste sem novo aporte de garantia derruba a suficiência: o texto
        # tem de acompanhar no mesmo rerun, sem congelar na apuração anterior.
        at.session_state["garantia_eventos_contrato"] = {
            "edited_rows": {}, "added_rows": [self._linha(TIPO_REAJUSTE, "1.100.000,00")],
            "deleted_rows": [0, 1, 2],
        }
        at.run()
        self.assertFalse(at.exception)
        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["diagnostico"], DIAGNOSTICO_VALOR)
        self.assertIn("complementação no valor de R$ 5.000,00", resultado["texto_comunicacao"])
        self.assertEqual(at.text_area(key="garantia_texto_comunicacao").value,
                         resultado["texto_comunicacao"])

    def test_percentual_manual_chega_ao_texto_da_pagina(self):
        at = self._rodar(percentual=3.0)
        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["garantia_necessaria"], Decimal("30000.00"))
        self.assertIn("correspondente a 3% passa a ser de R$ 30.000,00", resultado["texto_comunicacao"])
        self.assertNotIn("5%", resultado["texto_comunicacao"])


# ============================================================
# Card da Central de Arquivos como Calculadora
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
        self.assertIn('"pagina": "pages/05_Garantia.py"', bloco)
        self.assertIn('"ferramenta": True', bloco)
        self.assertIn('"sempre_acessivel": True', bloco)
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


if __name__ == "__main__":
    unittest.main()
