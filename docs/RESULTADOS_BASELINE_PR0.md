# RESULTADOS-BASELINE — PR 0

Rede de segurança que permite refatorar a aba **RESULTADOS** sem alterar
metodologia, matemática, VTA, retroativo, remanescente, ciclos, percentuais,
fatores, efeito financeiro, status, comportamento da web ou documentos.

Este PR **não altera nada de produção**. Ele apenas mede, congela e documenta o
comportamento de hoje.

---

## CHECKPOINT PRÉ-RESULTADOS

```
f8296f7c2962352716edd22044ed9573f5eeee8a
```

Estado para retorno caso qualquer etapa futura apresente regressão. Não
alterar. O SHA também está gravado em `tests/test_baseline_resultados.py`
(`CHECKPOINT_PRE_RESULTADOS`) e em cada snapshot dos 12 cenários.

---

## 1. O achado central: o contrato da aba é maior do que parece

A aba `RESULTADOS` tem 87 linhas e 270 células preenchidas. Apenas **28 delas
são contrato**; todo o resto é apresentação, livre para ser reescrita no PR 1.

### 1.1. Nove coordenadas lidas pelo runtime Python

| Coordenada | Quem lê | Para quê |
|---|---|---|
| `A1` | `_coleta_reajuste.py:352` | Título literal — gate de integridade; texto diferente reprova o arquivo |
| `B3` | `_coleta_reajuste.py:926` | `status_resultados["geral"]` — o status canônico da apuração |
| `B10` | `_leitor_masterfile_v10.py:3463` | Referência auditável — situação atual do contrato |
| `B11` | `_leitor_masterfile_v10.py:3464` | Referência auditável — última referência de abertura |
| `B12` | `_leitor_masterfile_v10.py:3465` | Comparativo — contrato integralmente reajustado |
| `B13` | `_leitor_masterfile_v10.py:3466` | Diferença entre as duas referências |
| `H10` | `_leitor_masterfile_v10.py:3467` | Situação da referência atual |
| `H11` | `_leitor_masterfile_v10.py:3467` | Situação da referência de abertura |
| `H13` | `_leitor_masterfile_v10.py:3467` | Status da conferência entre referências |

### 1.2. Dezenove células lidas por fórmulas de outra aba (acoplamento XLS→XLS)

O bloco **"5. AJUSTES MANUAIS"** (linhas 43-50) é **entrada do usuário**, e
`MEMORIA_RESULTADOS` a consome dentro do próprio Excel para compor VTA e
retroativo:

```
RESULTADOS!C43, D43, G43     RESULTADOS!C46, G46
RESULTADOS!C44, D44, G44     RESULTADOS!C47, G47
RESULTADOS!C45, D45, G45     RESULTADOS!C48, G48
                             RESULTADOS!C49, G49
                             RESULTADOS!C50, G50
```

> **Este é o acoplamento mais perigoso da frente.** Nenhum teste de Python o
> enxerga: renumerar ou mover essas linhas quebraria o cálculo dentro do Excel
> **sem produzir um único vermelho** na suíte. O baseline levanta o mapa a
> partir do arquivo (`mapear_referencias_de_outras_abas`) e o protege em
> `test_as_entradas_de_ajuste_manual_sao_contrato_dentro_do_excel`.

### 1.3. Cinco intervalos nomeados ancorados na aba

| Nome | Destino | Consumidor |
|---|---|---|
| `STATUS_RESULTADOS` | `RESULTADOS!$B$3` | leitura do status canônico |
| `VTA_ATUALIZACAO_CHEIA` | `RESULTADOS!$B$12` | comparativo integral |
| `OPCOES_APLICAR_MANUAL` | `RESULTADOS!$J$2:$J$3` | lista de validação da tabela manual |
| `EXECUCAO_ATUALIZADA_CICLO` | `RESULTADOS!$B$36` | **nenhum** |
| `SALDO_REMANESCENTE_ATUAL` | `RESULTADOS!$B$38` | **nenhum** |

Os dois últimos **não têm consumidor algum** — nem em produção, nem em teste.
Ficam fotografados para que o PR 1 decida conscientemente o que fazer com eles,
em vez de os arrastar por inércia ou os perder por descuido.

Os demais nomes do contrato (`VTA_FINAL`, `RETRO_OFICIAL`, `REM_*`,
`QTD_REM_*`, `VTA_CALCULADO`, `METODO_RETROATIVO`, …) apontam para
`MEMORIA_RESULTADOS`, não para a aba executiva — a refatoração do leiaute não
os toca.

---

## 2. O segundo achado: sem recálculo do Excel não há grandeza econômica

A cadeia Python **não** é independente do cache de fórmulas do XLS para VTA,
retroativo, remanescente, posição contratual e valores unitários. Num arquivo
gerado por `openpyxl` (sem recálculo), todas as capacidades retornam
*"Aguardando cálculo do XLS"* e o consolidado devolve
`vta_origem = "indisponivel"`.

Isso é o comportamento **correto** e fail-closed (P0-ROBUSTEZ-VALORES-1:
ausência nunca vira `R$ 0,00`), mas tem uma consequência de método:

> Um baseline de VTA/retroativo/remanescente **exige** arquivo recalculado.

Daí a arquitetura em duas suítes:

| Suíte | Fonte | O que congela | Precisa de Excel? |
|---|---|---|---|
| `test_baseline_resultados` | 12 cenários sintéticos | contrato estrutural, método, status, ciclos, situação, variação, percentual, fatores, efeito financeiro, mensagens, fail-closed | não |
| `test_baseline_resultados_goldens` | Coletas reais homologadas | VTA, retroativo, remanescente, quantidade, execução, composição, convergência XLS×Python | não (os arquivos já vêm recalculados de produção) |

`pywin32` não está instalado em nenhum interpretador desta máquina, e **nada foi
instalado**. A camada de recálculo por Excel COM não foi construída: os goldens
reais entregam o que ela entregaria, com a vantagem de serem números homologados
em produção em vez de sintéticos.

---

## 3. Convergência natural, nunca cópia

O item 6 da especificação proíbe o atalho `web == célula da RESULTADOS`. Ele não
foi usado. A prova de convergência vem de um mecanismo que **já existe em
produção**: `resultado["reconciliacao_xls_python"]`, que para cada grandeza
guarda o número do XLS **e** o número que o motor Python calculou por conta
própria, com tolerância declarada (`0,005`).

O baseline fotografa esse bloco inteiro, preservando as duas colunas. No golden
Financeiro homologado:

```
RETRO_OFICIAL           xls 24.678,92     python 24.678,92     CONCILIADO
VTA_FINAL               xls 8.713.820,26  python 8.713.820,26  CONCILIADO
REM_BASE_OFICIAL        xls 1.349.258,38  python 1.349.258,38  CONCILIADO
REM_ATUALIZADO_OFICIAL  xls 1.388.251,07  python 1.388.251,07  CONCILIADO
QTD_REM_OFICIAL         xls 7,9           python 7,9           CONCILIADO

status_geral = CONCILIADO      divergencias_relevantes = []
```

E, sem cache, o bloco declara `sem_cache = True` e
`status_geral = RESULTADO_XLS_INDISPONIVEL_POR_CACHE` — nunca "CONCILIADO"
sobre o vazio.

---

## 4. Triagem dos testes acoplados

`tests/baseline_resultados/inventario_testes_resultados.json`, recalculado a
cada execução por `test_baseline_inventario_resultados.py`.

**49 arquivos, 372 ocorrências.**

| Classe | Ocorrências | Significado |
|---|---:|---|
| A — contrato legítimo | 61 | intervalo nomeado do contrato vivo, coordenada lida pelo runtime, ou API do consolidado |
| B — teste de leiaute | 38 | apresentação: cor, fonte, borda, largura, merge, formato, visibilidade |
| C — teste legado | 9 | coordenada fora do leiaute atual (A1:J87) ou aplicador histórico de template |
| D — suspeito | 51 | célula do leiaute atual que não é contrato vivo nem verificação de apresentação |
| E — outro | 213 | só `MEMORIA_RESULTADOS`, ou menção textual sem endereço |

**Recorte que reduz o escopo do PR 1:** apenas **43** dos 49 arquivos tocam a
aba executiva. Estes seis referenciam somente `MEMORIA_RESULTADOS` (aba técnica
oculta, fora do escopo da refatoração de leiaute):

```
tests/test_26h_novos_itens.py
tests/test_441_base_fisica_c0_temporal.py
tests/test_44_novos_itens_aditivo_vta.py
tests/test_efeitos_financeiros_metodo_financeiro.py
tests/test_temporalidade_aditivos_data_efeito.py
tests/test_vta_pc_independente_27b.py
```

Regras de decisão gravadas em cada registro: `sobrevive_ao_pr1` (verdadeiro só
para a classe A) e `atualizar_no_pr2` (classes B, C e D). Cada registro traz
`arquivo`, `linha`, `teste`, `referencia`, `coordenadas`, `aba`,
`classificacao` e `justificativa` — a tabela pedida no item 3.

**A classificação é heurística e assumida como tal.** Ela orienta a triagem do
PR 2; não autoriza apagar nada. **Nenhum teste existente foi alterado por este
PR.**

---

## 5. Os 12 cenários

Construídos por `tests/_baseline_cenarios.py` sobre o template oficial, sem data
"hoje", sem aleatório e sem rede.

| # | Cenário | O que exercita | Coberto |
|---|---|---|---|
| 1 | `01_financeiro_normal` | Financeiro, C1 tempestivo, 24 competências pagas | sintético + golden |
| 2 | `02_pc` | PCs dentro do corte, com efeito financeiro → retroativo R$ 14.156,80 | sintético + golden |
| 3 | `03_itens_consumidos` | Método Consumido; retroativo fail-closed (`None`, nunca 0,00) | sintético |
| 4 | `04_multiciclo` | C1+C2+C3 computados, fatores encadeados | sintético + golden |
| 5 | `05_reajuste_negativo_aplicado` | percentual −2,18% desce ao fator do ciclo | sintético |
| 6 | `06_reajuste_negativo_neutralizado` | percentual efetivo 0,00% | sintético |
| 7 | `07_sem_ciclo_em_execucao` | sem a fotografia física → VTA `indisponivel` | sintético |
| 8 | `08_situacao_atual_posterior_ao_corte` | posição física de depois da data de corte | sintético |
| 9 | `09_referencia_anterior_ao_corte_estimado` | referência anterior ao corte → projeção | sintético |
| 10 | `10_sem_recalculo_do_excel` | `cache_ausente`, status indisponível, VTA `None` | sintético |
| 11 | `11_pcs_sem_efeito_financeiro` | PCs antes do início do efeito → retroativo 0,00 | sintético |
| 12 | `12_aditivo_no_meio_do_ciclo` | aditivo assinado no meio do ciclo vigente | sintético |

Os cenários 5 e 6 diferem **apenas** no tratamento da variação negativa, e o
teste exige que o percentual efetivo dos dois seja diferente — a decisão do
usuário precisa chegar ao fator.

---

## 6. Baseline das principais grandezas

Golden `financeiro_multiciclo_validado` (Coleta real homologada, ICTI, C1+C2+C3):

| Grandeza | Valor congelado |
|---|---|
| Método | Financeiro |
| Status | VALIDADO (origem: `resultados_xls`) |
| Ciclos considerados | C1, C2, C3 |
| **VTA Oficial** | **R$ 8.713.820,26** (origem: `vta_canonico`) |
| **Retroativo total** | **R$ 24.678,92** |
| Remanescente base | R$ 1.349.258,38 |
| Remanescente atualizado | R$ 1.388.251,07 |
| Quantidade remanescente | 7,9 |
| Execução atualizada | R$ 7.325.569,19 |
| Fator acumulado | 1,028899355437093 |
| Variação acumulada | 0,0288993554370931 |
| Convergência XLS×Python | CONCILIADO, sem divergências |

**Composição do VTA, fechando ao centavo:**

```
Executado apurado             7.300.890,27   (aba financeiro)
(+) Ajustes ainda devidos        24.678,92   (retroativo reconhecido)
(+) Remanescente atualizado   1.388.251,07   (fator 1,028899355437093)
(=) VTA Oficial               8.713.820,26
```

**Referências auditáveis — nunca são o VTA:** situação atual R$ 7.835.180,34;
última referência de abertura R$ 7.835.180,34; contrato integralmente reajustado
R$ 8.499.931,01; conferência entre referências R$ 0,00 (fecha).

Percentuais e fatores são gravados **sem arredondamento**, na precisão
matemática que a produção calculou; monetários, ao centavo.

---

## 7. Cobertura de web e documentos

**Web** — a fotografia sai do mesmo entry point do upload real
(`processar_coleta_oficial_runtime`), e cobre: método, status da apuração,
status de confiabilidade, mensagem de status, formalização e bloqueios, ciclos
considerados, ciclo vigente, situação / variação / percentual efetivo / fator
próprio / fator acumulado / efeito financeiro por ciclo, retroativo por ciclo,
VTA e origem, retroativo total e potencial, remanescente, quantidade, execução
atualizada, composição do VTA, totais canônicos de PC, bloco "fora do corte",
referências auditáveis, convergência, ressalvas, informações, campos não
confiáveis, avisos e pendências.

**Documentos** — conteúdo negocial (valores e decisões), não bytes:

| Documento | Como entra |
|---|---|
| Termo de Apostila | `gerar_termo_apostila` → linhas negociais do DOCX |
| Despacho Saneador | `gerar_despacho_saneador` → linhas negociais do DOCX |
| Sumário Executivo | `montar_dados_sumario_executivo` → síntese, composição, ciclos, observações |
| Comunicação à contratada | `gerar_rascunho_email_contratada` → assunto + linhas negociais |
| DOU | `pages/13_DOU.py` → campos automáticos (valor total, texto do reajuste) |

O teste `test_documentos_publicam_os_mesmos_numeros_da_apuracao` prova que a
síntese do Sumário reproduz exatamente o VTA, o retroativo e o remanescente
apurados, e que o retroativo aparece no Termo de Apostila.

Snapshot de bytes inteiros foi deliberadamente evitado: seria frágil a fonte,
espaçamento, data de geração e ID de apuração, sem proteger nada de substantivo.

---

## 8. Ponto de atenção registrado, não corrigido

No golden Financeiro, `sintese["resumo_executivo"]` do Sumário Executivo diz
*"Resultado: R$ 14.626.459,46"* enquanto o VTA oficial da mesma síntese é
R$ 8.713.820,26. A ordem de grandeza é próxima do defeito já corrigido em VTA-U2
(composição do Financeiro somando a linha TOTAL agregadora).

Está **fotografado no baseline** e **não foi tocado**: corrigir número é alterar
comportamento, e o PR 0 existe para congelar o que existe. Fica registrado para
decisão em frente própria.

---

## 9. Terminologia — regra para as etapas futuras

O baseline **registra** os textos atuais, inclusive "reconciliação" e "posição
física", porque tem de fotografar o estado existente. Nada de produção foi
alterado para aplicar esta regra: isso pertence ao PR 2.

Proibido em **código novo ou documentação nova destinada ao usuário**:
"reconciliação", "reconciliado" e derivados. Usar, conforme o caso:
conferência · conferência cruzada · fechamento · confere · não confere ·
diferença.

Padronização futura aprovada:

| Hoje | PR 2 |
|---|---|
| Posição física atual | Situação atual do contrato |
| Última posição de abertura | Última referência de abertura |
| Conferência entre posições físicas | Conferência entre referências do contrato |
| Referência da posição física | Referência contratual |
| Data da posição física | Data da referência |
| Fotografia da posição física | Fotografia do contrato / situação contratual na data |

---

## 10. Status visual — decisão aprovada para o PR 2 (não implementada)

A faixa superior da Página 1 mostrará o **status canônico** com destaque:

| Estado | Cor |
|---|---|
| VALIDADO | verde destacado |
| ESTIMADO / REVISE / pendência equivalente | amarelo ou âmbar |
| reprovação, bloqueio ou erro | vermelho |
| neutro / informativo | cinza ou azul discreto |

**A cor deriva do MESMO status canônico que XLS e web já usam** — hoje
`RESULTADOS!B3` → `status_resultados["geral"]` → `consolidado["status_apuracao"]`.
Não criar um segundo classificador só para escolher cor; não hardcodar status
independente.

---

## 11. Leiaute aprovado — apenas registrado

Não implementado neste PR.

**PÁGINA 1 — RESULTADO DA APURAÇÃO:** título; faixa superior (status · método ·
ciclos · data de corte · índice); dois destaques (VTA OFICIAL; RETROATIVO TOTAL
A PAGAR); faixa secundária (remanescente · saldo do ciclo · variação acumulada);
quadro CICLOS APURADOS (Ciclo · Situação · Período · Variação apurada ·
Percentual aplicado · Efeito financeiro · Observação); informações e pendências.

**PÁGINA 2 — MEMÓRIA E AUDITORIA:** 1. VTA — metodologia e formação;
2. Retroativo por ciclo; 3. Remanescente por ciclo; 4. Ciclo em execução;
5. Referências para auditoria — **não são o VTA oficial**; 6. Conferência da
execução — auditoria do método Financeiro.

Integração técnica não terá terceira página impressa.

---

## 12. Como usar

```bash
# Rede completa (~11 min): 12 cenários + goldens + inventário
python -m pytest tests/test_baseline_resultados.py \
                tests/test_baseline_resultados_goldens.py \
                tests/test_baseline_inventario_resultados.py

# Só um cenário, no dia a dia
python -m pytest tests/test_baseline_resultados.py -k multiciclo

# Regravar o baseline — ATO DELIBERADO, nunca "para fazer passar"
set RESULTADOS_BASELINE_REGRAVAR=1
python -m pytest tests/test_baseline_resultados.py
```

Goldens reais ficam fora do repositório, em `CL8US_GOLDENS_DIR`
(`C:\Users\danie\Downloads\anthropic-skills` por padrão) e
`CL8US_GOLDENS_PC_DIR`. Ausentes, os testes correspondentes são **pulados** —
nunca falso-verde. As fotografias JSON ficam versionadas, de modo que os números
permanecem auditáveis mesmo sem os arquivos.

**Custo:** cada cenário carrega o template oficial (~16 MB de XML), processa o
upload completo e gera cinco documentos — ~20-25 s por cenário. A suíte é cara e
está deliberadamente fora do CI rápido, que roda apenas os sentinelas.

---

## 13. Arquivos deste PR

Todos novos. **Nenhum arquivo de produção foi alterado.**

```
tests/_baseline_cenarios.py                          construtor dos 12 cenários
tests/_baseline_fotografia.py                        fotógrafo (4 camadas)
tests/test_baseline_resultados.py                    baseline dos 12 cenários
tests/test_baseline_resultados_goldens.py            grandezas econômicas reais
tests/test_baseline_inventario_resultados.py         inventário vivo dos 49
tests/baseline_resultados/*.json                     12 snapshots
tests/baseline_resultados/goldens/*.json             3 snapshots
tests/baseline_resultados/inventario_testes_resultados.json
docs/RESULTADOS_BASELINE_PR0.md                      este documento
```
