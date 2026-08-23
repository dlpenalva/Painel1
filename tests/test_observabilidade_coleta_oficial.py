"""Prova de observabilidade minima em processar_coleta_oficial_runtime.

Cobre: (1) processamento bem sucedido emite INICIO/ETAPA_CONCLUIDA/FIM com id
consistente e duracao numerica, via stdout best-effort (sem logging/config
global); (2) nenhuma string sensivel (contratada, contrato, valor, arquivo,
PC) nem mensagem completa de excecao aparece na saida, mesmo quando a
excecao funcional as contem; (3) o fail-safe existente de cobertura
temporal pode emitir ETAPA_FALHA_FAILSAFE sem alterar seu comportamento;
(4) a excecao fatal propagada ao chamador permanece exatamente a mesma
(identidade do objeto, mensagem, sem try/except externo capturando-a).
"""

import io
import re
import unittest
from contextlib import redirect_stdout
from unittest.mock import patch

import _coleta_reajuste_documentos as _crd


def _mocks_sucesso_minimo():
    return dict(
        garantir_xlsx_validado=lambda conteudo: conteudo,
        ler_masterfile_v10=lambda conteudo, exigir_modelo_oficial=True: {
            "ok": True,
            "reconciliacao_xls_python": {},
            "avisos": [],
        },
        ler_coleta_reajuste=lambda conteudo: {"valido": True},
        adaptar_coleta_reajuste_para_documentos=lambda conteudo, leitura, diagnostico: {
            "capacidades": {}
        },
        avaliar_entrega_segura=lambda leitura: {"bloqueios": []},
        campos_nao_confiaveis_para_documentos=lambda reconciliacao: [],
        montar_resultado_consolidado=lambda resultado, diagnostico: {},
    )


def _capturar_stdout(func, *args, **kwargs):
    buffer = io.StringIO()
    with redirect_stdout(buffer):
        retorno = func(*args, **kwargs)
    return retorno, buffer.getvalue()


class TestObservabilidadeSucesso(unittest.TestCase):
    def test_processamento_bem_sucedido_emite_eventos_com_id_consistente(self):
        with patch.multiple("_coleta_reajuste_documentos", **_mocks_sucesso_minimo()):
            (resultado, diagnostico), saida = _capturar_stdout(
                _crd.processar_coleta_oficial_runtime, b"conteudo-fake"
            )

        linhas = [linha for linha in saida.splitlines() if linha.strip()]
        inicio = [l for l in linhas if l.startswith("PROCESSAMENTO_INICIO")]
        etapas = [l for l in linhas if l.startswith("ETAPA_CONCLUIDA")]
        fim = [l for l in linhas if l.startswith("PROCESSAMENTO_FIM")]

        self.assertEqual(len(inicio), 1)
        self.assertGreaterEqual(len(etapas), 1)
        self.assertEqual(len(fim), 1)
        # Nenhum PROCESSAMENTO_ERRO nesta implementacao (removido por decisao).
        self.assertNotIn("PROCESSAMENTO_ERRO", saida)

        ids = {re.search(r"id=(\S+)", l).group(1) for l in linhas}
        self.assertEqual(len(ids), 1, f"id deveria ser consistente em toda a execucao: {ids}")

        for linha in etapas + fim:
            duracao = re.search(r"duracao_s=(\S+)", linha).group(1)
            self.assertIsNotNone(float(duracao))

        # Resultado/diagnostico nao recebem nenhuma chave de observabilidade.
        for chave in ("id_proc", "id", "duracao_s", "etapa", "observabilidade"):
            self.assertNotIn(chave, resultado)
            self.assertNotIn(chave, diagnostico)

    def test_fail_safe_cobertura_temporal_pode_emitir_evento_sem_mudar_comportamento(self):
        mocks = _mocks_sucesso_minimo()
        with patch.multiple("_coleta_reajuste_documentos", **mocks):
            with patch(
                "_motor_cobertura_temporal.montar_cobertura_temporal",
                side_effect=RuntimeError("falha interna do motor temporal"),
            ):
                (resultado, diagnostico), saida = _capturar_stdout(
                    _crd.processar_coleta_oficial_runtime, b"conteudo-fake"
                )

        # Comportamento fail-safe original preservado: nao derruba o processamento,
        # e o diagnostico["cobertura_temporal"] recebe o formato de erro existente.
        self.assertFalse(diagnostico["cobertura_temporal"]["ok"])
        self.assertIn("erro", diagnostico["cobertura_temporal"])
        self.assertIn("ETAPA_FALHA_FAILSAFE", saida)
        self.assertIn("etapa=cobertura_temporal", saida)
        self.assertIn("tipo=RuntimeError", saida)
        # A mensagem interna da excecao nao vaza para o evento tecnico.
        self.assertNotIn("falha interna do motor temporal", saida)
        # O processamento chega ate o fim normalmente (fail-safe, nao fatal).
        self.assertIn("PROCESSAMENTO_FIM", saida)


class TestObservabilidadePrivacidade(unittest.TestCase):
    SENTINELAS = (
        "EMPRESA_SIGILOSA_XYZ",
        "CONTRATO_SIGILOSO_123",
        "987654321.99",
        "arquivo_ultrassecreto.xlsx",
        "PC_SIGILOSO_999",
    )

    def test_nenhuma_sentinela_sensivel_aparece_na_saida_de_sucesso(self):
        mocks = _mocks_sucesso_minimo()
        mocks["ler_coleta_reajuste"] = lambda conteudo: {
            "valido": True,
            # Payload de diagnostico com dados de negocio: nao deve vazar a saida.
            "contratada": "EMPRESA_SIGILOSA_XYZ",
            "contrato": "CONTRATO_SIGILOSO_123",
            "valor_pc": "987654321.99",
            "numero_pc": "PC_SIGILOSO_999",
        }
        with patch.multiple("_coleta_reajuste_documentos", **mocks):
            _, saida = _capturar_stdout(
                _crd.processar_coleta_oficial_runtime,
                b"conteudo do arquivo arquivo_ultrassecreto.xlsx",
            )

        for sentinela in self.SENTINELAS:
            self.assertNotIn(sentinela, saida)

    def test_excecao_fatal_com_dados_sensiveis_nao_vaza_mensagem_na_saida(self):
        mocks = _mocks_sucesso_minimo()
        mensagem_sensivel = (
            "Divergencia na contratada EMPRESA_SIGILOSA_XYZ, contrato "
            "CONTRATO_SIGILOSO_123, PC PC_SIGILOSO_999, valor 987654321.99, "
            "arquivo arquivo_ultrassecreto.xlsx"
        )

        def _leitura_com_erro_sensivel(conteudo, exigir_modelo_oficial=True):
            raise ValueError(mensagem_sensivel)

        mocks["ler_masterfile_v10"] = _leitura_com_erro_sensivel

        buffer = io.StringIO()
        with patch.multiple("_coleta_reajuste_documentos", **mocks):
            with redirect_stdout(buffer):
                with self.assertRaises(ValueError) as ctx:
                    _crd.processar_coleta_oficial_runtime(b"conteudo-fake")
        saida = buffer.getvalue()

        # A excecao propagada ao chamador permanece intacta (com a mensagem sensivel:
        # isso e responsabilidade do chamador/UI, nao da observabilidade tecnica).
        self.assertEqual(str(ctx.exception), mensagem_sensivel)

        for sentinela in self.SENTINELAS:
            self.assertNotIn(sentinela, saida)
        self.assertNotIn(mensagem_sensivel, saida)
        # Sem PROCESSAMENTO_ERRO (removido): so o ultimo ETAPA_CONCLUIDA registrado
        # (validacao_xlsx) permite saber ate onde o fluxo chegou.
        self.assertIn("ETAPA_CONCLUIDA id=", saida)
        self.assertIn("etapa=validacao_xlsx", saida)
        self.assertNotIn("etapa=leitura_masterfile", saida)
        self.assertNotIn("PROCESSAMENTO_FIM", saida)


class TestObservabilidadeExcecaoIdentica(unittest.TestCase):
    def test_excecao_propagada_preserva_identidade_sem_try_except_externo(self):
        mocks = _mocks_sucesso_minimo()
        excecao_original = ValueError("falha especifica de leitura")

        def _falha(conteudo, exigir_modelo_oficial=True):
            raise excecao_original

        mocks["ler_masterfile_v10"] = _falha

        buffer = io.StringIO()
        with patch.multiple("_coleta_reajuste_documentos", **mocks):
            with redirect_stdout(buffer):
                with self.assertRaises(ValueError) as ctx:
                    _crd.processar_coleta_oficial_runtime(b"conteudo-fake")

        # Sem try/except externo: a excecao que chega ao chamador e o MESMO objeto
        # (identidade), nao uma copia/reconstrucao — prova de que nada a intercepta.
        self.assertIs(ctx.exception, excecao_original)
        self.assertEqual(str(ctx.exception), "falha especifica de leitura")


if __name__ == "__main__":
    unittest.main()
