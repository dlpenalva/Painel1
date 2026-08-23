"""Prova de observabilidade minima em processar_coleta_oficial_runtime.

Cobre: (1) processamento bem sucedido emite INICIO/ETAPA_CONCLUIDA/FIM com id
consistente e duracao numerica; (2) nenhuma string sensivel (contratada,
contrato, valor, arquivo, PC) nem mensagem completa de excecao aparece nos
logs, mesmo quando a excecao funcional as contem; (3) a excecao propagada ao
chamador permanece exatamente a mesma (tipo, mensagem, encadeamento).
"""

import logging
import re
import unittest
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


class TestObservabilidadeSucesso(unittest.TestCase):
    def test_processamento_bem_sucedido_emite_eventos_com_id_consistente(self):
        with patch.multiple("_coleta_reajuste_documentos", **_mocks_sucesso_minimo()):
            with self.assertLogs("_coleta_reajuste_documentos", level="WARNING") as captura:
                resultado, diagnostico = _crd.processar_coleta_oficial_runtime(b"conteudo-fake")

        mensagens = [registro.getMessage() for registro in captura.records]

        inicio = [m for m in mensagens if m.startswith("PROCESSAMENTO_INICIO")]
        etapas = [m for m in mensagens if m.startswith("ETAPA_CONCLUIDA")]
        fim = [m for m in mensagens if m.startswith("PROCESSAMENTO_FIM")]

        self.assertEqual(len(inicio), 1)
        self.assertGreaterEqual(len(etapas), 1)
        self.assertEqual(len(fim), 1)

        ids = {re.search(r"id=(\S+)", m).group(1) for m in mensagens}
        self.assertEqual(len(ids), 1, f"id deveria ser consistente em toda a execucao: {ids}")

        for mensagem in etapas + fim:
            duracao = re.search(r"duracao_s=(\S+)", mensagem).group(1)
            self.assertIsNotNone(float(duracao))

        # Resultado/diagnostico nao recebem nenhuma chave de observabilidade.
        for chave in ("id_proc", "id", "duracao_s", "etapa", "observabilidade"):
            self.assertNotIn(chave, resultado)
            self.assertNotIn(chave, diagnostico)


class TestObservabilidadePrivacidade(unittest.TestCase):
    SENTINELAS = (
        "EMPRESA_SIGILOSA_XYZ",
        "CONTRATO_SIGILOSO_123",
        "987654321.99",
        "arquivo_ultrassecreto.xlsx",
        "PC_SIGILOSO_999",
    )

    def test_nenhuma_sentinela_sensivel_aparece_no_log_de_sucesso(self):
        mocks = _mocks_sucesso_minimo()
        mocks["ler_coleta_reajuste"] = lambda conteudo: {
            "valido": True,
            # Payload de diagnostico com dados de negocio: nao deve vazar ao log.
            "contratada": "EMPRESA_SIGILOSA_XYZ",
            "contrato": "CONTRATO_SIGILOSO_123",
            "valor_pc": "987654321.99",
            "numero_pc": "PC_SIGILOSO_999",
        }
        with patch.multiple("_coleta_reajuste_documentos", **mocks):
            with self.assertLogs("_coleta_reajuste_documentos", level="WARNING") as captura:
                _crd.processar_coleta_oficial_runtime(
                    b"conteudo do arquivo arquivo_ultrassecreto.xlsx"
                )

        texto_log = "\n".join(registro.getMessage() for registro in captura.records)
        for sentinela in self.SENTINELAS:
            self.assertNotIn(sentinela, texto_log)

    def test_excecao_funcional_com_dados_sensiveis_nao_vaza_mensagem_no_log(self):
        mocks = _mocks_sucesso_minimo()
        mensagem_sensivel = (
            "Divergencia na contratada EMPRESA_SIGILOSA_XYZ, contrato "
            "CONTRATO_SIGILOSO_123, PC PC_SIGILOSO_999, valor 987654321.99, "
            "arquivo arquivo_ultrassecreto.xlsx"
        )

        def _leitura_com_erro_sensivel(conteudo, exigir_modelo_oficial=True):
            raise ValueError(mensagem_sensivel)

        mocks["ler_masterfile_v10"] = _leitura_com_erro_sensivel

        with patch.multiple("_coleta_reajuste_documentos", **mocks):
            with self.assertLogs("_coleta_reajuste_documentos", level="WARNING") as captura:
                with self.assertRaises(ValueError) as ctx:
                    _crd.processar_coleta_oficial_runtime(b"conteudo-fake")

        # A excecao propagada ao chamador permanece intacta (com a mensagem sensivel:
        # isso e responsabilidade do chamador/UI, nao do log tecnico lateral).
        self.assertEqual(str(ctx.exception), mensagem_sensivel)

        texto_log = "\n".join(registro.getMessage() for registro in captura.records)
        for sentinela in self.SENTINELAS:
            self.assertNotIn(sentinela, texto_log)
        self.assertNotIn(mensagem_sensivel, texto_log)
        self.assertIn("PROCESSAMENTO_ERRO", texto_log)
        self.assertIn("tipo=ValueError", texto_log)
        self.assertIn("etapa=leitura_masterfile", texto_log)


class TestObservabilidadeExcecaoIdentica(unittest.TestCase):
    def test_excecao_propagada_preserva_tipo_mensagem_e_encadeamento(self):
        mocks = _mocks_sucesso_minimo()
        excecao_original = ValueError("falha especifica de leitura")

        def _falha(conteudo, exigir_modelo_oficial=True):
            raise excecao_original

        mocks["ler_masterfile_v10"] = _falha

        with patch.multiple("_coleta_reajuste_documentos", **mocks):
            logging.getLogger("_coleta_reajuste_documentos").disabled = True
            try:
                with self.assertRaises(ValueError) as ctx:
                    _crd.processar_coleta_oficial_runtime(b"conteudo-fake")
            finally:
                logging.getLogger("_coleta_reajuste_documentos").disabled = False

        self.assertIs(ctx.exception, excecao_original)
        self.assertEqual(str(ctx.exception), "falha especifica de leitura")


if __name__ == "__main__":
    unittest.main()
