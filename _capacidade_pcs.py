"""Capacidade canonica de PCs (itens_PC) — fonte unica de verdade.

Etapa 26G: o limite historico de 100 linhas era arquitetural e estava
espalhado em formulas, validacoes e leitores. A capacidade passa a ser
definida UNICAMENTE aqui e consumida por:

* tools/aplicar_escala_pcs_26g.py (owner das faixas no template oficial);
* _coleta_reajuste.py (validacao de upload, duplicidade e contagens);
* _leitor_masterfile_v10.py (limite operacional de leitura);
* _coleta_oficial.py (barreira de densidade da grade);
* testes de fronteira.

Regra fail-closed: PCs alem da capacidade NUNCA sao truncados em silencio.
O upload e bloqueado explicitamente quando ha conteudo apos a ultima linha
da grade (ver _coleta_reajuste.ler_coleta_reajuste).
"""
from __future__ import annotations

# Capacidade maxima de PCs na grade itens_PC (linhas de dados).
# Caso real homologado: 2.118 PCs — margem de ~2,4x.
CAPACIDADE_PCS = 5000

# Grade de dados: linha 1 e cabecalho; dados em 2..ULTIMA_LINHA_PCS.
PRIMEIRA_LINHA_PCS = 2
ULTIMA_LINHA_PCS = CAPACIDADE_PCS + 1  # 5001
