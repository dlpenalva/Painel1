"""Contexto de UMA execucao do fluxo oficial da Coleta.

O upload oficial materializava QUATRO vezes o mesmo XLSX: uma por leitor e
mais uma por varredura, cada uma repagando o parse XML integral de um pacote
com ~16 MB de XML de worksheet. O custo dominante do upload nao era logica
financeira, era abrir o mesmo arquivo de novo.

Este contexto abre cada REPRESENTACAO do arquivo no maximo uma vez por
execucao e a entrega pronta aos leitores:

* ``workbook_valores``  — ``data_only=True``: os valores calculados que o
  Excel gravou em cache.
* ``workbook_formulas`` — ``data_only=False``: as proprias formulas.

As duas representacoes continuam separadas de proposito. Elas respondem a
perguntas diferentes ("quanto deu?" x "existe formula aqui?") e uni-las apenas
para reduzir a contagem de aberturas trocaria desempenho por semantica.

XSEC-09 (fronteira de seguranca): os bytes passam por
``garantir_xlsx_validado`` na construcao e CADA workbook passa por
``validar_geometria_workbook`` dentro do proprio metodo de abertura, antes de
o objeto ser devolvido a qualquer chamador. Nenhum acesso a celula acontece
antes do gate, e nenhuma abertura escapa dele — a garantia fica mais forte que
a atual, em que cada leitor precisava lembrar de aplicar o gate.

Escopo: uma execucao, um usuario, um upload. O contexto NAO entra em cache
global, NAO entra em ``session_state``, NAO persiste em disco e NAO sobrevive
entre requisicoes — um mesmo contrato pode conter informacao sensivel e o
compartilhamento e estritamente intra-execucao.
"""

from __future__ import annotations

from io import BytesIO
from typing import Any

from openpyxl import load_workbook

from _seguranca_xlsx import garantir_xlsx_validado, validar_geometria_workbook


class ContextoColeta:
    """Compartilha os workbooks ja validados de UMA execucao da Coleta.

    Uso tipico (o unico previsto): criado no ponto de entrada do runtime,
    repassado aos leitores e fechado ao final da mesma chamada.

        with ContextoColeta(conteudo) as contexto:
            leitura = ler_masterfile_v10(contexto.conteudo, contexto=contexto)

    Chamadas isoladas dos leitores (testes e consumidores internos) continuam
    validas sem contexto: cada leitor abre os proprios bytes como antes.
    """

    __slots__ = ("_conteudo", "_wb_valores", "_wb_formulas")

    def __init__(self, conteudo: bytes) -> None:
        self._conteudo: bytes = garantir_xlsx_validado(conteudo)
        self._wb_valores: Any = None
        self._wb_formulas: Any = None

    @property
    def conteudo(self) -> bytes:
        """Os bytes ja aprovados na fronteira ZIP/OOXML."""
        return self._conteudo

    def _abrir(self, *, data_only: bool):
        wb = load_workbook(BytesIO(self._conteudo), data_only=data_only)
        try:
            # Gate de geometria ANTES de devolver o objeto: quem recebe o
            # workbook deste contexto ja o recebe dentro de um retangulo
            # aprovado.
            validar_geometria_workbook(wb)
        except BaseException:
            # Arquivo reprovado na fronteira nao fica pendurado no processo: o
            # caminho hostil e justamente onde o workbook nao pode sobreviver.
            wb.close()
            raise
        return wb

    @property
    def workbook_valores(self):
        """Workbook ``data_only=True`` — valores em cache do Excel."""
        if self._wb_valores is None:
            self._wb_valores = self._abrir(data_only=True)
        return self._wb_valores

    @property
    def workbook_formulas(self):
        """Workbook ``data_only=False`` — as formulas, nao o cache."""
        if self._wb_formulas is None:
            self._wb_formulas = self._abrir(data_only=False)
        return self._wb_formulas

    def fechar(self) -> None:
        """Libera os workbooks da execucao. Idempotente."""
        for atributo in ("_wb_valores", "_wb_formulas"):
            wb = getattr(self, atributo)
            setattr(self, atributo, None)
            if wb is None:
                continue
            try:
                wb.close()
            except Exception:
                # Fechar e higiene de recurso: nunca pode derrubar um upload
                # que ja produziu resultado.
                pass

    def __enter__(self) -> "ContextoColeta":
        return self

    def __exit__(self, *_excecao) -> bool:
        self.fechar()
        return False
