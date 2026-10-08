"""Versoes publicas e carimbo da versao publicada do Cl8us.

O commit mais recente e a fonte primaria no Streamlit Cloud. O fallback deve
ser atualizado em toda entrega para manter o marcador mesmo quando o Git nao
estiver disponivel no ambiente de execucao.

Politica de versionamento publico (sempre no formato ``XX.X``; regra completa
em ``docs/VERSIONAMENTO.md``):

* TODA entrega em ``main`` que altere algo percebido pelo usuario incrementa
  ``CL8US_VERSION`` — no MESMO PR da alteracao;
* TODA entrega que altere o XLSX entregue incrementa ``COLETA_VERSION`` (e a
  nova versao entra em ``COLETA_VERSOES_ACEITAS``);
* ``ATUALIZADO_EM_FALLBACK`` e atualizado junto com o bump;
* numero ja consumido por uma entrega nunca e reutilizado.

``tools/verificar_versionamento.py`` faz o CI falhar quando a politica e
violada; nada e incrementado automaticamente.

``MASTERFILE_VERSION`` continua sendo um marcador tecnico legado onde ainda
for necessario; nao deve ser substituido automaticamente por estas versoes.
"""

from __future__ import annotations

import subprocess
from datetime import datetime
from pathlib import Path


# 11.8: aditivos com fator vigente carregado (item existente x novo item),
# CONTROLE!B2 derivado da data de corte (ciclo em execucao), novos itens no
# CICLO_EM_EXECUCAO e opcao "Acrescimo - novo item" — muda o XLS entregue:
# COLETA_VERSION 11.4.
# 11.7: quatro ajustes de UX da Coleta (PR #175): aviso de percentuais
# historicos, fronteira IST explicada, bloco opcional X:AG, RESULTADOS_DETALHE
# oculta — muda o XLS entregue: COLETA_VERSION 11.3. O bump foi esquecido no
# PR original e aplicado no hotfix de versionamento obrigatorio.
# 11.6: hotfix temporal da ancora historica do 1o ciclo (PR #174) — so app;
# a Coleta seguiu 11.2. Numero consumido logicamente pelo PR #174 (bump
# esquecido no PR original); por isso a entrega seguinte e 11.7.
# 11.5: nova aba RESULTADOS executiva; a RESULTADOS anterior passa a se chamar
# RESULTADOS_DETALHE (mesmas formulas/coordenadas) — mudanca estrutural do XLS:
# COLETA_VERSION 11.2.
# 11.4: aba RESULTADOS ganha o quadro informativo "Execucao sem efeito financeiro"
# (Financeiro, PCs e Itens) — mudanca estrutural do XLS: COLETA_VERSION 11.1.
# 11.3: modelos em branco do Despacho Saneador e do Termo de Apostila com a mesma
# estrutura dos documentos gerados. 11.2: memoria de calculo da garantia em XLSX.
CL8US_VERSION = "11.8"
COLETA_VERSION = "11.4"
# Modelos de Coleta da familia 11.x aceitos SEM adaptacao: o 11.1 so acrescentou
# um quadro informativo na RESULTADOS/MEMORIA_RESULTADOS (formulas); a Coleta 11.0
# nao o possui e continua valida. A 11.2 so separa a apresentacao (RESULTADOS
# executiva) da camada tecnica (RESULTADOS_DETALHE): o motor e o mesmo e os
# leitores resolvem a aba tecnica por `_resultados_abas`. Nao remover versoes
# desta lista sem decisao expressa — um arquivo 11.0/11.1 nunca e bloqueado por
# nao ter a aba executiva nova. A 11.3 so muda apresentacao (aviso, destaque,
# X:AG, RESULTADOS_DETALHE oculta): mesma estrutura de leitura da 11.2. A 11.4
# so muda formulas (CONTROLE!B2, aditivos!I/J/M) e o dropdown de aditivos!D:
# mesmas abas e coordenadas de leitura.
COLETA_VERSOES_ACEITAS = ("11.0", "11.1", "11.2", "11.3", "11.4")
# Marcadores publicos ANTERIORES a camada tecnica RESULTADOS_DETALHE: so neles
# a RESULTADOS ainda e a aba tecnica. Qualquer outro marcador (11.2 ou
# posterior) exige RESULTADOS_DETALHE (fail-closed). Arquivos PRE_11 nao tem
# marcador e seguem a compatibilidade legada.
COLETA_VERSOES_SEM_RESULTADOS_DETALHE = ("11.0", "11.1")
# Janela de compatibilidade retroativa: a versao atual e DUAS linhagens anteriores
# homologadas (PRE_11_L1 e PRE_11_L2, ver _compatibilidade_coleta). Estrutura fora
# desta janela e rejeitada. A formalizacao de Coleta compatibilizada e decidida
# em _formalizacao_compatibilidade, por evidencia tecnica e nunca por versao.
COLETA_COMPATIBILIDADE_ANTERIORES = 2

ATUALIZADO_EM_FALLBACK = "08/10/2026 15:00"


def _data_ultimo_commit() -> str | None:
    try:
        resultado = subprocess.run(
            ["git", "log", "-1", "--format=%cI"],
            cwd=Path(__file__).resolve().parent,
            capture_output=True,
            text=True,
            timeout=5,
        )
        valor = (resultado.stdout or "").strip()
        if not valor:
            return None
        return datetime.fromisoformat(valor).strftime("%d/%m/%Y %H:%M")
    except Exception:
        return None


def atualizado_em() -> str:
    """Retorna o carimbo visivel no formato brasileiro."""
    data_commit = _data_ultimo_commit()
    if not data_commit:
        return ATUALIZADO_EM_FALLBACK
    formato = "%d/%m/%Y %H:%M"
    return max((data_commit, ATUALIZADO_EM_FALLBACK), key=lambda valor: datetime.strptime(valor, formato))
