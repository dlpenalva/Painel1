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

Regra do carimbo de publicacao:

* se o Git estiver disponivel, o horario do ultimo commit e SEMPRE a fonte
  primaria e deve ser exibido sem comparacao com o fallback;
* ``ATUALIZADO_EM_FALLBACK`` so pode ser usado quando o Git nao estiver
  disponivel ou nao retornar um commit valido;
* o fallback nunca pode sobrescrever um horario real obtido do Git.

``tools/verificar_versionamento.py`` faz o CI falhar quando a politica e
violada; nada e incrementado automaticamente.

``MASTERFILE_VERSION`` continua sendo um marcador tecnico legado onde ainda
for necessario; nao deve ser substituido automaticamente por estas versoes.
"""

from __future__ import annotations

import subprocess
from datetime import datetime
from pathlib import Path


# 12.14: Painel Executivo e Assistente Fiscal exibem o VTA canonico (o mesmo da
# pagina Valor Global) em vez da referencia historica do resumo, que no formato
# Coleta virava R$ 0,00; ausencia de VTA aparece como "Indisponivel" (#195).
# 12.13: Coleta 12.6 — Financeiro reconhece retroativo zero quando ha execucao
# sem parcela com efeito; referencias temporais incluem a execucao ate a data
# real da fotografia; abas tecnicas permanecem ocultas por padrao.
# 12.12: Coleta 12.5 — RESULTADOS: a tabela SALDO REMANESCENTE — EVOLUCAO sobe
# para B65:G72, logo apos a observacao final; o bloco DETALHES DO METODO — PCs
# desce para B74:G83. So posicao: calculos, valores e estilos inalterados.
# 12.11: Coleta 12.4 — precisao financeira canonica: agregacoes monetarias do
# XLS fecham em 2 casas (sem residuo binario nem -R$ 0,00) e o web normaliza o
# zero negativo na exibicao; RESULTADOS ganha a tabela SALDO REMANESCENTE —
# EVOLUCAO (espelho das fontes canonicas). Percentual, fator e regras inalterados.
# 12.10: Coleta 12.3 — CONTROLE simplificada (dropdowns visiveis e caixa verde
# unica) e carimbo de versao no final logico da sidebar. Nenhuma regra financeira
# ou estrutural das abas de calculo muda.
# 12.9: Coleta 12.2 — acabamento visual: separacao L|M em itens_PC, mensagem
# NOVO ITEM em aditivos!A2:A200, CONTROLE!B mais larga com orientacao do metodo
# em C1 (formatacao condicional) e CICLO_EM_EXECUCAO na paleta clara das demais
# abas. Nenhuma regra financeira ou estrutural muda.
# 12.8: Coleta 12.1 — RESULTADOS ganha o card RETRO. RECONHECIDO + POTENCIAL
# (soma-resumo dos dois cards, sem nova parcela financeira) e a SITUACAO DA
# APURACAO vira faixa de largura total com tres estados (concluida, concluida
# com retroativo potencial, revise). As abas financeiro, itens_Consumidos e
# itens_PC ganham no topo (linha 1) a orientacao do metodo selecionado em
# quatro estados (ativo, outro metodo, outro metodo com dados, sem metodo).
# Nenhum calculo financeiro muda.
# 12.7: Coleta 12.0 — itens_PC com filtro nativo (aba protegida), PC_PAGO
# normalizado na coluna tecnica oculta AD (Sim/Nao sem diferenca de maiusculas,
# acento ou espacos; demais grafias invalidas) e alerta de coerencia entre o
# metodo (CONTROLE!B1) e a fonte de dados: fonte do metodo vazia bloqueia o VTA
# e o retroativo, sem usar outra fonte automaticamente.
# 12.6: Coleta 11.9 — CONTROLE!B2 volta a ser estritamente automatico (sem
# dropdown residual) e parametros!U:W ficam ocultas.
# 12.5: Coleta 11.8 — ajustes da Coleta em uso: VTA (Financeiro/PCs) usa a
# posicao atual da CICLO_EM_EXECUCAO quando falta a fotografia do ciclo
# vigente; PC_PAGO aceita Nao/Não; PC de ciclo PRECLUSO sem inicio de efeito
# fica sem efeito proprio; ajustes visuais (parametros, posicao_referencia,
# Nxxx verde, cobertura_temporal).
# 12.4: Coleta 11.7 — itens Nxxx usam a base economica do VU (aditivos!N)
# tambem em historico_VU e no VU da CICLO_EM_EXECUCAO, e o VTA passa a
# valoriza-los pelo VU reajustado desde essa base (nao mais desde o nascimento).
# 12.3: Coleta 11.6 passa a usar a ultima referencia fisica conhecida por item,
# separa a base economica do VU do nascimento fisico e destaca itens Nxxx.
# 12.2: hotfix do carimbo de publicacao: quando o Git esta disponivel, o horario
# real do ultimo commit passa a ter prioridade absoluta; o fallback so e usado
# quando o Git nao fornece um commit valido. COLETA_VERSION permanece 11.5.
# 12.1: parametros!A8 so pede percentual de ciclo ANTERIOR a um ciclo
# computado (nao mais ate o ciclo em execucao) e mensagens de entrada orientam
# itens novos por aditivo em itens_Remanesc/aditivos — muda o XLS entregue:
# COLETA_VERSION 11.5.
# 12.0: pagina "Analisar um unico ciclo" deixa de exibir "Nao houve pedido da
# contratada" (estado residual descartado) — so app. Segue 11.9 no formato XX.X.
# 11.9: Reajustes Multiplos oculta "Nao houve pedido da contratada" quando a
# analise atual tem UM unico ciclo (estado residual descartado) — so app.
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
CL8US_VERSION = "12.14"
COLETA_VERSION = "12.6"
# Modelos de Coleta da familia 11.x aceitos SEM adaptacao: o 11.1 so acrescentou
# um quadro informativo na RESULTADOS/MEMORIA_RESULTADOS (formulas); a Coleta 11.0
# nao o possui e continua valida. A 11.2 so separa a apresentacao (RESULTADOS
# executiva) da camada tecnica (RESULTADOS_DETALHE): o motor e o mesmo e os
# leitores resolvem a aba tecnica por `_resultados_abas`. Nao remover versoes
# desta lista sem decisao expressa — um arquivo 11.0/11.1 nunca e bloqueado por
# nao ter a aba executiva nova. A 11.3 so muda apresentacao (aviso, destaque,
# X:AG, RESULTADOS_DETALHE oculta): mesma estrutura de leitura da 11.2. A 11.4
# so muda formulas (CONTROLE!B2, aditivos!I/J/M) e o dropdown de aditivos!D:
# mesmas abas e coordenadas de leitura. A 11.5 so muda a formula do aviso
# parametros!A8, a regra de destaque de parametros!E3:E6 e mensagens de entrada.
# A 11.6 acrescenta aditivos!N/O fora das coordenadas de leitura; a 11.7 so
# muda formulas de historico_VU!D:G. A 11.8 so muda formulas/apresentacao e
# acrescenta MEMORIA_RESULTADOS!S49:T53 fora das coordenadas de leitura. A
# 11.9 remove somente a validacao residual de CONTROLE!B2 e oculta
# parametros!U:W. A 12.0 acrescenta itens_PC!AD (oculta), MEMORIA_RESULTADOS
# S54:T65/AG43:AG64 e CONTROLE!C1 fora das coordenadas de leitura. A 12.1 so
# muda a RESULTADOS executiva (apresentacao) e acrescenta MEMORIA_RESULTADOS
# S66:T67/AF65:AG79 e a orientacao na linha 1 de financeiro (H1:L1),
# itens_Consumidos (AH1:AL1) e itens_PC (AE1:AI1), fora das coordenadas de
# leitura. A 12.2 so muda apresentacao (bordas, larguras, cores, mensagem de
# entrada e formatacao condicional): mesmas coordenadas de leitura da 12.1.
# A 12.4 so fecha agregacoes monetarias em 2 casas (ROUND) e acrescenta
# RESULTADOS!B76:G83 e MEMORIA_RESULTADOS!V68:W70/AF81:AG83, fora das
# coordenadas de leitura. A 12.5 so troca a posicao de dois blocos da
# RESULTADOS executiva (B65:G83), que nenhum leitor consulta. A 12.6 corrige
# formulas da memoria financeira e das referencias temporais, sem mudar as
# coordenadas ou a metodologia dos demais modos.
COLETA_VERSOES_ACEITAS = (
    "11.0", "11.1", "11.2", "11.3", "11.4", "11.5", "11.6", "11.7", "11.8", "11.9",
    "12.0", "12.1", "12.2", "12.3", "12.4", "12.5", "12.6",
)
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

# Fallback conservador da ultima publicacao confirmada antes deste hotfix.
# Quando o Git esta disponivel, este valor nunca prevalece sobre o commit real.
ATUALIZADO_EM_FALLBACK = "20/10/2026 18:00"


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
    """Retorna o carimbo visivel no formato brasileiro.

    O Git e a fonte primaria. O fallback so existe para ambientes em que o
    repositorio/commit nao esteja disponivel. Nunca se compara os dois horarios,
    evitando que um fallback incorreto ou futuro sobrescreva a publicacao real.
    """
    data_commit = _data_ultimo_commit()
    return data_commit if data_commit else ATUALIZADO_EM_FALLBACK
