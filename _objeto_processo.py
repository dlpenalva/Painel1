"""Modelo de domínio único da apuração — o Objeto Processo.

Este módulo não lê o XLS, não calcula e não decide regra de negócio: ele apenas
dá nome e tipo ao que a Coleta já produz. A montagem a partir do workbook
continua em `_coleta_reajuste_documentos`, que é o único ponto que abre o
arquivo.

`ObjetoProcesso.como_dicionario()` reproduz integralmente o contrato de dados
histórico consumido pelos documentos e pela Interface. Enquanto essa camada de
compatibilidade existir, nenhum consumidor precisa ser alterado; a migração para
os campos tipados pode ocorrer módulo a módulo.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import pandas as pd


@dataclass(frozen=True)
class Identificacao:
    """De onde veio a apuração e sob qual índice ela corre."""

    origem_coleta: str
    data_processamento: str
    modo_apuracao: str
    indice: str
    ciclo_ultimo_remanescente: str


@dataclass(frozen=True)
class Ciclos:
    """Os ciclos da apuração e o fator que eles compõem."""

    tabela: pd.DataFrame
    quantidade: int
    fator_acumulado: float
    variacao_acumulada: float
    origem: str


@dataclass(frozen=True)
class Financeiro:
    """Execução mensal informada pelo fiscal, por competência e por ciclo."""

    mensal: pd.DataFrame
    por_ciclo: pd.DataFrame
    valor_pago_efetivo: float
    valor_teorico_calculado: float
    disponivel: bool
    aviso: str
    meses_sem_efeito: pd.DataFrame
    quantidade_meses_sem_efeito: int
    valor_total_sem_efeito: float


@dataclass(frozen=True)
class Retroativo:
    """Produto 1 do sistema: o retroativo e sua rastreabilidade."""

    valor: float
    disponivel: bool
    capacidade: dict[str, Any] = field(default_factory=dict)
    estimado_itens_estoque: pd.DataFrame = field(default_factory=pd.DataFrame)


@dataclass(frozen=True)
class ValorTotalAtualizado:
    """Produto 3 do sistema: o VTA — conceito pétreo — e sua composição."""

    valor_original_contrato: float
    valor_total: float
    execucao_atualizada: float
    composicao: pd.DataFrame
    capacidade: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class Remanescentes:
    """Produtos 2 e 4: VU atualizado por ciclo e o saldo remanescente."""

    original: float
    reajustado: float
    fator: float
    tabela: pd.DataFrame
    valores_unitarios: pd.DataFrame
    execucao: pd.DataFrame
    capacidade: dict[str, Any] = field(default_factory=dict)
    capacidade_valores_unitarios: dict[str, Any] = field(default_factory=dict)
    base_itens_disponivel: bool = False


@dataclass(frozen=True)
class Aditivos:
    """Aditivos e seu efeito — ou ausência de efeito — no Valor Global."""

    tabela: pd.DataFrame
    computaveis: pd.DataFrame
    total_atualizados: float
    quantidade_total: int
    quantidade_computaveis: int


@dataclass(frozen=True)
class PosicaoContratual:
    """Quantidades contratadas e remanescentes ajustadas por ciclo."""

    tabela: pd.DataFrame
    capacidade: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class Diagnostico:
    """O que o leitor apurou sobre a própria Coleta."""

    capacidades: dict[str, Any] = field(default_factory=dict)
    status_resultados: dict[str, Any] = field(default_factory=dict)
    coleta: dict[str, Any] = field(default_factory=dict)
    comparativo: pd.DataFrame = field(default_factory=pd.DataFrame)
    auditoria_consistencia: pd.DataFrame = field(default_factory=pd.DataFrame)
    ressalva_modo_apuracao: str = ""


@dataclass(frozen=True)
class ObjetoProcesso:
    """Representação única do contrato para a Interface e para os documentos.

    Depois de montado, nenhum consumidor precisa voltar ao XLS.
    """

    identificacao: Identificacao
    ciclos: Ciclos
    financeiro: Financeiro
    retroativo: Retroativo
    vta: ValorTotalAtualizado
    remanescentes: Remanescentes
    aditivos: Aditivos
    posicao_contratual: PosicaoContratual
    diagnostico: Diagnostico

    def como_dicionario(self) -> dict[str, Any]:
        """Expõe o contrato histórico, chave a chave, sem alterar valores.

        Camada de compatibilidade: existe para que documentos e Interface
        continuem funcionando sem qualquer alteração enquanto migram para os
        campos tipados acima.
        """

        identificacao = self.identificacao
        ciclos = self.ciclos
        financeiro = self.financeiro
        retroativo = self.retroativo
        vta = self.vta
        remanescentes = self.remanescentes
        aditivos = self.aditivos
        diagnostico = self.diagnostico

        return {
            "ok": True,
            "origem_coleta": identificacao.origem_coleta,
            "data_processamento": identificacao.data_processamento,
            "modo_apuracao": identificacao.modo_apuracao,
            "base_execucao_mensal_disponivel": financeiro.disponivel,
            "base_itens_disponivel": remanescentes.base_itens_disponivel,
            "aviso_base_execucao": financeiro.aviso,
            "ressalva_modo_apuracao": diagnostico.ressalva_modo_apuracao,
            "config_ciclo_em_execucao": {},
            "corte_operacional_solicitado": False,
            "corte_operacional_aplicado": False,
            "origem_ciclos": ciclos.origem,
            "indice": identificacao.indice,
            "fator_acumulado": ciclos.fator_acumulado,
            "variacao_acumulada": ciclos.variacao_acumulada,
            "quantidade_ciclos": ciclos.quantidade,
            "valor_original_contrato": vta.valor_original_contrato,
            "contexto_contratual_anterior": {},
            "valor_formalizado_anterior": vta.valor_original_contrato,
            "impacto_analise_atual": vta.valor_total - vta.valor_original_contrato,
            "valor_pago_efetivo": financeiro.valor_pago_efetivo,
            "total_pago_faturado": financeiro.valor_pago_efetivo,
            "valor_teorico_calculado": financeiro.valor_teorico_calculado,
            "total_devido_reajustado": financeiro.valor_teorico_calculado,
            "delta_total": retroativo.valor,
            "delta_acumulado": retroativo.valor,
            "valor_represado_a_pagar": retroativo.valor,
            "valor_retroativo_estimado_itens_estoque": retroativo.valor,
            "retroativo_estimado_itens_estoque_disponivel": retroativo.disponivel,
            "quantidade_meses_sem_efeito_financeiro": financeiro.quantidade_meses_sem_efeito,
            "valor_total_sem_efeito_financeiro": financeiro.valor_total_sem_efeito,
            "remanescente_original": remanescentes.original,
            "remanescente_reajustado": remanescentes.reajustado,
            "fator_remanescente": remanescentes.fator,
            "valor_executado_atualizado": vta.execucao_atualizada,
            "valor_calculado_sem_aditivos": vta.valor_total,
            "valor_atualizado_contrato": vta.valor_total,
            "valor_global_financeiro": vta.valor_total,
            "total_aditivos_atualizados": aditivos.total_atualizados,
            "total_aditivos_informativos": 0.0,
            "aditivos_somados_ao_valor_total": False,
            "quantidade_aditivos_total": aditivos.quantidade_total,
            "quantidade_aditivos_marcados_computaveis": aditivos.quantidade_computaveis,
            "ciclo_ultimo_remanescente": identificacao.ciclo_ultimo_remanescente,
            "df_ciclos": ciclos.tabela,
            "df_financeiro_mensal": financeiro.mensal,
            "df_financeiro_mensal_corte_operacional": financeiro.mensal,
            "df_financeiro_mensal_tratado": financeiro.mensal,
            "df_meses_sem_efeito_financeiro": financeiro.meses_sem_efeito,
            "df_financeiro_por_ciclo": financeiro.por_ciclo,
            "df_delta_por_ciclo": financeiro.por_ciclo.copy(),
            "df_execucao_atualizada": remanescentes.execucao,
            "df_retroativo_estimado_itens_estoque": retroativo.estimado_itens_estoque,
            "df_composicao_valor_total": vta.composicao,
            "df_remanescentes": remanescentes.tabela,
            "df_valores_unitarios_ciclo": remanescentes.valores_unitarios,
            "df_aditivos": aditivos.tabela,
            "df_aditivos_executivo": aditivos.tabela,
            "df_aditivos_computaveis": aditivos.computaveis,
            "df_aditivos_informativos": pd.DataFrame(),
            "df_posicao_contratual": self.posicao_contratual.tabela,
            "df_comparativo": diagnostico.comparativo,
            "df_auditoria_consistencia": diagnostico.auditoria_consistencia,
            "status_resultados": diagnostico.status_resultados,
            "capacidades": diagnostico.capacidades,
            "diagnostico_coleta": diagnostico.coleta,
            "resultados_progressivos": {
                "retroativo": retroativo.capacidade,
                "vta": vta.capacidade,
                "valor_remanescente": remanescentes.capacidade,
                "posicao_contratual": self.posicao_contratual.capacidade,
                "valores_unitarios": remanescentes.capacidade_valores_unitarios,
            },
            "_resultado_lido_do_excel": True,
        }
