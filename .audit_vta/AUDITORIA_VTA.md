# AUDITORIA DO CÁLCULO DO VTA — mapa integral

Base: `origin/main` = `ed7b0ce` (HOMOLOGADO_CICLO_EM_EXECUCAO_20260803).
Template oficial: `templates/COLETA_REAJUSTE_OFICIAL.xlsx`.
Worktree isolado: `C:\_DesktopReal\VTA_Posicoes_Resultados_Itens_RC`.
Branch: `feature/vta-posicoes-resultados-itens-rc`. **Sem push/PR/deploy.**

> Observação metodológica: **não** se alterou nenhuma fórmula. Este documento é
> exclusivamente o mapa exigido pela Seção 4 do enunciado.

---

## 0. Arquitetura geral

O masterfile é um **template XLSX com fórmulas Excel embutidas**. O Python
(`_coleta_oficial.py`, `_gerador_masterfile.py`, `_ciclo_em_execucao.py`) apenas
**preenche abas de entrada** e injeta a aba opcional `CICLO_EM_EXECUCAO` em
runtime. As fórmulas do VTA vivem no XLSX, não no Python.

Abas do template (ordem real):
`comparativo_VTA, CONTROLE, parametros, financeiro, itens_Remanesc,
itens_Consumidos, itens_PC, aditivos, posicao_referencia, posicao_contratual,
itens_RC, historico_VU, cobertura_temporal, MEMORIA_RESULTADOS, RESULTADOS`.
`CICLO_EM_EXECUCAO` **não** está no template — é criada por
`_ciclo_em_execucao.py` (layout `2_ITEMIZADO`), logo após `posicao_referencia`.

---

## 1. Cadeia canônica do VTA (evidência do arquivo real confere)

```
RESULTADOS!B10  = IFERROR(VTA_FINAL,"")
VTA_FINAL       = MEMORIA_RESULTADOS!B26           (nome definido)
B26 (método PCs)= ROUND(T25 + AJUSTE_MANUAL_VTA, 2) [se T25<>"CALCULO MANUAL REQUERIDO"]
T25             = ROUND(T21 + T22 + T23, 2)         [se base válida]
```

Para o arquivo real (método **PCs**, ciclo vigente **C1**, `T20=1`):

| Célula | Fórmula (essência) | Significado econômico | Valor |
|---|---|---|---|
| `T21` | `IF(itens_PC!P2>0, itens_PC!P2, SUM(X2:X201))` | **C0 executado**: PC de C0 se houver, senão execução física implícita por item | 6.034.401,56 |
| `T22` | `SUMPRODUCT((row-2>=1)*(row-2<T20)*itens_PC!P2:P6)` | **Ciclos encerrados** entre C0 e o vigente (nenhum quando vigente=C1) | 0,00 |
| `T23` | `SUM(Y2:Y201)` | **Remanescente físico atualizado na ABERTURA do ciclo vigente** | 7.434.449,85 |
| `T25` | `ROUND(T21+T22+T23,2)` | VTA-PC | **13.468.851,41** |

Colunas auxiliares (MEMORIA_RESULTADOS):
- `X2 = (posicao_contratual!G2 − posicao_contratual!K2) × historico_VU!C2`
  = **(QTD_REM_AJUSTADA_C0 − QTD_REM_AJUSTADA_C1) × VU_C0**
  = **execução física implícita em C0** (já é exatamente a fórmula da Seção 17).
- `Y2 = ROUND(VU_vigente,2) × QREM_vigente`, com `CHOOSE(T20+1, ...)` escolhendo
  a coluna do ciclo vigente em `posicao_contratual` (G/K/O/S/W) e `historico_VU`
  (C/D/E/F/G). É o remanescente **na abertura** do ciclo vigente, item a item.
- `T20 = VALUE(MID(CONTROLE!B2,2,3))` → número do ciclo vigente.
- Guardas: `T24` (base itemizada presente), `T26` (itens sem remanescente do ciclo
  vigente), `T27` (itens sem base p/ C0 físico). `T25` vira `"CALCULO MANUAL
  REQUERIDO"` se `T24=0`, `T20=""`, `T26>0`, ou (sem PC C0 e `T27>0`).

**Conclusão-chave:** o VTA-PC atualmente adotado (`T25`) é conceitualmente
**"VTA pela posição de ABERTURA do ciclo vigente"** — não usa data-de-corte nem
posição física corrente. Ele já corresponde à **FORMA 2** do enunciado, porém
fixada no ciclo vigente (sem fallback para a última abertura completa anterior).

---

## 2. Os três métodos de retroativo (EIXO 1, MEMORIA_RESULTADOS B10:F16)

Por ciclo (linhas 10–14 = C0…C4), três fontes paralelas + coluna oficial:

- **Financeiro** (col B): `SUMIFS(financeiro!F, ciclo, G="Sim")` — valor pago.
- **PCs** (col C): `SUMIFS(itens_PC!H, ciclo=Cn, G="Sim")` — VALOR_ATUALIZADO dos PCs.
- **Itens** (col D): consumo físico × fator do ciclo (via `itens_Consumidos` +
  `parametros!F/D`).
- **Método selecionado** (col E) = `IF(B4="Financeiro",B,IF("PCs",C,D))`, onde
  `B4` deriva de `CONTROLE!B1`.
- `RETRO_OFICIAL (B16)` = `IF(ISNUMBER(B5), B5 /*manual*/, E15 /*total do método*/)`.
- `F16` = status com tolerância `D4=0.005` (DIVERGENTE / CALCULADO / MANUAL…).

Só **um** método é oficial; os demais são conferência. Nada é somado entre métodos.

---

## 3. Composição do VTA (EIXO 3, B20:B26)

| Célula | Nome | Fórmula | Significado |
|---|---|---|---|
| `B20` | — | `SUMIF(itens_Remanesc!A<>"", itens_Remanesc!D)` | Valor original do contrato |
| `B21` | — | `=B16` | Retroativo oficial |
| `B22` | — | `ROUND(D35−C35,2)` | Ajuste do remanescente (com − sem reajuste) |
| `B23` | `VTA_CALCULADO` | `ROUND(B20+B21+B22,2)` | VTA calculado (Financeiro/Itens) |
| `B24` | `AJUSTE_MANUAL_VTA` | de `RESULTADOS!C44` se `G44="Sim"` | Ajuste manual (+/−) |
| `B25` | `VTA_MANUAL_OFICIAL` | de `RESULTADOS!C45` se `G45="Sim"` | VTA manual substitutivo |
| `B26` | `VTA_FINAL` | ver §1 | **VTA oficial** |

Notas:
- Para **PCs**, `B26` usa `T25` (não `B23`) — a via itemizada por abertura.
- Para **Financeiro/Itens**, `B26 = B23 + B24 + N263` (complementos de aditivos).
- `B25` (manual) **prevalece** sobre tudo; `B24`+`B25` juntos ⇒ "ENTRADA MANUAL DUPLA".
- `E26` = status: exige `F16` e `F36` válidos; senão "CÁLCULO MANUAL REQUERIDO".

---

## 4. Remanescente (EIXO 4, B31:F36) e origem por método

- Linha 32 (`itens_Remanesc`): quantidade/valor do remanescente no ciclo vigente,
  com `posicao_contratual` (G/K/O/S/W) × `itens_RC` (D/G/J/M/P).
- Linha 33 (`itens_Consumidos`): remanescente derivado (base − consumido).
- Linha 35 `OFICIAL`: `B35/C35/D35` conforme método; para **PCs**, `D35 = T23`.
- `F36` = status do remanescente (tolerância `D4`).

---

## 5. As três referências do enunciado ↔ o que já existe

| Enunciado | Onde já existe hoje | Situação |
|---|---|---|
| **1. VTA PELA POSIÇÃO ATUAL DO CONTRATO** | **Não existe como VTA.** Insumos crus em `RESULTADOS!B35/B36/B38` (bloco "4. CICLO ATUAL EM EXECUÇÃO") e na aba `CICLO_EM_EXECUCAO` (totais `total_valor_consumido` + `total_valor_remanescente`). | **A CRIAR** |
| **2. VTA PELA ÚLTIMA POSIÇÃO DE ABERTURA DISPONÍVEL** | `MEMORIA_RESULTADOS!T25` / `VTA_FINAL` (PCs), fixado no ciclo **vigente**. | **Existe; falta o fallback** C4→C0 p/ última abertura completa |
| **3. CONTRATO ORIGINAL INTEGRALMENTE REAJUSTADO** | `comparativo_VTA!B208` = `SUMPRODUCT(pos_contratual!B×C) × CONTROLE!B11`; espelhado em `RESULTADOS!B11` e `MEMORIA_RESULTADOS!B29`. | **Existe** (comparativo; nunca oficial) |

---

## 6. Bloco "4. CICLO ATUAL EM EXECUÇÃO" (RESULTADOS 33–39) — insumo da FORMA 1

- `B35` (`Data de referência`) = data-de-corte por método (`cobertura_temporal`/`posicao_referencia`).
- `B36` = `EXECUCAO_ATUALIZADA_CICLO` = executado atualizado **até a data** (por método).
- `B37` = remanescente atualizado **na abertura** do ciclo = `INDEX(C26:C30; ciclo)`.
- `B38` = `SALDO_REMANESCENTE_ATUAL` = `ROUND(B37 − B36, 2)`.

**Identidade estrutural (crucial):** `B37 = B36 + B38`. Ou seja, "remanescente na
abertura" = "executado até a data" + "saldo atual". Portanto:

```
FORMA 1 (posição atual) do ciclo vigente = B36 (executado até data) + saldo_atual
FORMA 2 (abertura)       do ciclo vigente = B37 (remanescente na abertura) = T23
```

Quando o **saldo atual é derivado** (`B38 = B37 − B36`), as duas formas são
**iguais por construção** — coerente com a Seção 6/8 ("a execução posterior apenas
transfere base entre executado e remanescente; não deve ser somada de novo") e com
a reconciliação esperada = R$ 0,00.

A divergência (⇒ REVISE) só nasce quando o **remanescente atual vem de fonte física
independente** — a aba `CICLO_EM_EXECUCAO`, coluna C "QTD REMANESCENTE NA DATA ATUAL
(PREENCHER)" informada pelo fiscal — e essa fotografia não bate com `abertura − executado`.

---

## 7. `CICLO_EM_EXECUCAO` (runtime, layout `2_ITEMIZADO`) — motor da FORMA 1

Estrutura (de `_ciclo_em_execucao.py`):
- `C3` = ciclo (`=CONTROLE!B2`); `F3` = início; `H3` = fim técnico; `D5` = **DATA DA
  POSIÇÃO ATUAL (único input de data)**; `A9` = valor total remanescente na data
  (`SUM(G13:G211)`, com guardas de erro/incompleto).
- Itens (linhas 13–211, espelham `itens_Remanesc!2:200`), colunas visíveis A:G:
  - `A` ITEM (AUTO) — oculta item novo cujo nascimento é posterior à data.
  - `B` QTD REM. NO INÍCIO DO CICLO (AUTO) = `posicao_contratual` F/J/N/R/V
    (QTD_REM_**BASE**), com zero para item novo sem aditivo ≤ D5.
  - `C` QTD REM. NA DATA ATUAL (**PREENCHER**) — **único campo manual**; vazio =
    não confirmado; zero = confirmado zero.
  - `D` QTD CONSUMIDA ATÉ A DATA (AUTO) = `B + I − C`.
  - `E` VU ATUALIZADO (AUTO) = `historico_VU` por ciclo.
  - `F` VALOR CONSUMIDO (AUTO) = `D×E`; `G` VALOR REM. ATUALIZADO (AUTO) = `C×E`.
  - Técnicas ocultas: `I` alterações líquidas no período (`aditivos!L` com
    `F3 < data ≤ D5`), `J` CHECK_FISICO = `B+I−D−C`, `K` STATUS, `L/M/N` conferência
    com `itens_Consumidos`, `O` eventos positivos até a data.
- Motor puro `calcular_posicao_ciclo_por_data`: remanescente atual **declarado é
  autoritativo**; movimentos com efeito `≤ t0` já estão na abertura; só
  `t0 < efeito ≤ D` entra no balanço; efeito `> D` é ignorado. Impõe
  `base = consumido + remanescente` (RECONCILIACAO_MONETARIA).
- Leitura `ler_ciclo_em_execucao`: `disponivel` (layout presente), `utilizado`
  (data OU qtd atual), `completo`, `valido`, e totais
  `total_valor_consumido / total_valor_remanescente / total_base_fisica_atualizada`.
  **Estes totais são os insumos oficiais da FORMA 1.**

Ativação (Seção 13) já suportada pela leitura: ausente/legado/itemizado;
utilizado/completo/valido. Nenhuma parcela do VTA homologado é tocada por este módulo.

---

## 8. Tratamento temporal de aditivos / itens novos — o que já existe

`posicao_contratual` colunas `Y (CICLO_NASCIMENTO)` e `Z (EH_NOVO_ITEM, padrão Nxxx)`:
- `Y = IF(E>0,0,IF(I>0,1,IF(M>0,2,IF(Q>0,3,IF(U>0,4,"")))))` — 1º ciclo com posição.
- Remanescente ajustado por ciclo (G/K/O/S/W) é **bloqueado com `""`** quando
  `Y > n` (item ainda não existia) — i.e. **NÃO APLICÁVEL**, não zero. (Seção 19)
- `X (CHECK)`: "NOVO ITEM COM BASE ≠ 0", "NOVO ITEM SEM VU_ORIGINAL",
  "REMANESCENTE_SUPERA_POSICAO", etc.
- `itens_Remanesc!B` força `QTD_BASE_ORIGINAL=0` p/ Nxxx; aditivos entram via
  `SUMIFS(aditivos!L, item, ciclo)` (DELTA_Cn) — **não retroagem** (cada delta no seu ciclo).

**Portanto o núcleo do tratamento temporal de aditivos já está implementado** via
`CICLO_NASCIMENTO`. O que a Seção 18/19 pede de novo é sobretudo **exposição/status**
("NÃO APLICÁVEL" explícito; pendência de abertura vs. saldo confirmado) e o
roteamento da posição atual de itens nascidos no meio do ciclo para
`CICLO_EM_EXECUCAO` (que a coluna `A`/`I` da aba já contempla via `data ≤ D5`).

---

## 9. Divergências / riscos encontrados na auditoria

1. **Abertura BASE vs. AJUSTADA:** `CICLO_EM_EXECUCAO!B` usa `posicao_contratual`
   F/J/N/R/V (QTD_REM_**BASE**, sem aditivo) + coluna `I` (aditivos do período),
   enquanto `T23`/`itens_RC` usam G/K/O/S/W (QTD_REM_**AJUSTADA**, com aditivo do
   ciclo). A reconciliação FORMA1×FORMA2 precisa tratar essa diferença de origem
   (senão aditivos no início do ciclo podem aparecer como "divergência" espúria).
2. **FORMA 2 sem fallback:** hoje `T20` = ciclo vigente fixo. Se a abertura do ciclo
   vigente estiver incompleta, `T25` vira "CALCULO MANUAL REQUERIDO" — não há a
   seleção "última abertura completa C4→C0" da Seção 7.
3. **Tabela 1 só tem 2 linhas** (B10 VTA método; B11 integralmente reajustado).
   Falta a linha "VTA pela posição atual" e a renomeação para as 3 referências
   canônicas.
4. **Reconciliação FORMA1×FORMA2 (diff=0/REVISE)** não existe como célula.
5. **itens_RC** só tem as fotografias de **abertura** por ciclo (C0…C4). Falta o
   bloco "posição atual (AUTO)" consumindo `CICLO_EM_EXECUCAO` (Seções 11–15).

---

## 10. Fontes canônicas x conciliatórias (Seção 5)

| Componente | Fonte canônica | Cobertura temporal | Conciliatórias |
|---|---|---|---|
| C0 executado | `itens_PC!P2` (PC C0); senão físico implícito `X` | até fim C0 | financeiro C0; itens C0 |
| Ciclos encerrados | `itens_PC!P` (VALOR_ATUALIZADO), ciclos 1..vig−1 | aberturas dos ciclos | financeiro; itens |
| Remanescente abertura (FORMA 2) | `T23` = `posic_contratual` G/K/.. × `historico_VU` | abertura do ciclo vigente | — |
| Execução do ciclo até a data (FORMA 1) | `CICLO_EM_EXECUCAO` col F (consumido) | `F3 < t ≤ D5` | `RESULTADOS!B36` |
| Remanescente atual (FORMA 1) | `CICLO_EM_EXECUCAO` col G (declarado) | na data `D5` | `RESULTADOS!B38` (derivado) |

**Regra anti-dupla-contagem:** no método PCs/Financeiro a execução do ciclo vigente
NÃO deve ser somada de novo à FORMA 2 (a abertura já contém toda a base); na FORMA 1
soma-se executado-até-data + remanescente-atual, cuja soma = base da abertura.
```

---

## 11. Correção temporal dos aditivos (DATA_EFEITO) — Regra 2 resolvida

### 11.1 Causa técnica

`posicao_contratual!Y` (CICLO_NASCIMENTO) é `=IF($E2>0,0,IF($I2>0,1,...))` — deriva
da **quantidade contratada acumulada**, não da data. As colunas `DELTA_Cn`
(`D/H/L/P/T`) agregam os aditivos por **rótulo de ciclo**
(`SUMIFS(aditivos!L; item; ciclo="Cn")`), sem qualquer recorte de data dentro do
ciclo. Consequência: um item incluído no 1º dia do C1 e outro incluído no meio do
C1 recebem o mesmo `Y=1` e a mesma abertura `K = J + H`. A fotografia de abertura
passava a conter quantidade que ainda não existia naquela data, e
`itens_Remanesc` classificava como *aplicável/pendente* um campo que deveria ser
**NÃO APLICÁVEL**.

### 11.2 Regra canônica implementada

```
QTD CONTRATUAL NA ABERTURA DE Cn = quantidade-base + Σ deltas com DATA_EFEITO ≤ abertura(Cn)
```

Implementada por **decomposição do delta oficial**, sem alterar nenhuma fórmula da
cadeia oficial:

```
DELTA_POSTERIOR_ABERTURA_Cn = SUMIFS(aditivos!L; item; ciclo="Cn"; data ≥ INT(abertura)+1)
QTD_CONTRATUAL_ABERTURA_Cn  = QTD_CONTRATADA_Cn − DELTA_POSTERIOR_ABERTURA_Cn
CICLO_NASCIMENTO_DATA       = menor n com QTD_CONTRATUAL_ABERTURA_Cn > 0
```

Como todo delta de ciclo anterior tem data ≤ abertura de `Cn`, a subtração isola
exatamente a parcela do próprio ciclo cujo efeito é posterior à abertura. A
fronteira é **por dia** (`≥ INT(abertura)+1`): um aditivo datado no dia da abertura
permanece do lado da abertura mesmo se a célula carregar componente de hora.

### 11.3 Células criadas

| Aba | Colunas | Conteúdo |
|---|---|---|
| `posicao_contratual` | `AA` | `DATA_EFEITO_INICIAL` (informativa; `_xlfn.MINIFS` + `IFERROR`) |
| `posicao_contratual` | `AB:AF` | `DELTA_POSTERIOR_ABERTURA_C0..C4` |
| `posicao_contratual` | `AG:AK` | `QTD_CONTRATUAL_ABERTURA_C0..C4` |
| `posicao_contratual` | `AL` | `CICLO_NASCIMENTO_DATA` |
| `MEMORIA_RESULTADOS` | `AD2:AD201` | componente **alterações posteriores à abertura** (VU × delta posterior) |
| `MEMORIA_RESULTADOS` | `AC2:AC201` | componente **remanescente na abertura** (`AB − AD`, residual exato) |
| `MEMORIA_RESULTADOS` | `W53/W54` | somas dos dois componentes |
| `MEMORIA_RESULTADOS` | `W55/W56` | trava anti-dupla-contagem e status |
| `MEMORIA_RESULTADOS` | `W57` | itens incluídos após a abertura adotada (qtd) |
| `itens_RC` | `Z:AC` | data de efeito, ciclo de nascimento por data, aplicabilidade, qtd na abertura |
| `itens_Remanesc` | `BI` | espelho local de `CICLO_NASCIMENTO_DATA` (ver 11.6) |

Células **alteradas**: `MEMORIA_RESULTADOS!W48` (recomposta) e
`itens_Remanesc!F/H/J/L` (guarda `AL > n` ⇒ valor de abertura vazio).

### 11.4 Aditivos posteriores à abertura e regra contra dupla contagem

A FORMA 2 **não exclui** o item acrescido depois da abertura — ele entra por
componente próprio:

```
W48 = T21 (C0 executado)
    + ciclos encerrados (itens_PC!P, ciclos 1..W46−1)
    + W53 (remanescente na abertura, temporalmente correto)
    + W54 (alterações contratuais posteriores à abertura)
```

`AC` é definido como **residual** de `AB − AD`, portanto
`W53 + W54 ≡ SUM(AB)` por construção — o mesmo total que a FORMA 2 tinha antes.
`W55 = ROUND(SUM(AB) − (W53+W54), 2)` é a trava programática: qualquer parcela
contada duas vezes a torna diferente de zero e `W56` passa a
"REVISE - DUPLA CONTAGEM NA FORMA 2". O tool aborta se `|W55| > 0,005`.

### 11.5 Impacto nas FORMAS 1 e 2

* **FORMA 2** — total inalterado; a **leitura temporal** de cada parcela muda
  (abertura vs. posterior). O fallback C4→C0 (`W46`) continua igual.
* **FORMA 1** — `CICLO_EM_EXECUCAO` já era correta por data na coluna `A`
  (item novo só aparece após a data de efeito) e na coluna `I`
  (`abertura < data ≤ posição`). Corrigiu-se a lacuna do aditivo datado
  **exatamente na abertura**, que não entrava nem na coluna `B` nem na `I`:
  `B` passa a somar os deltas do ciclo com data `< INT(F3)+1` e `I` passa a
  começar em `≥ INT(F3)+1`. Isso fecha a **divergência nº 1 da §9** e faz a
  reconciliação FORMA 1 × FORMA 2 ser exata por construção.
* O motor puro `calcular_posicao_ciclo_por_data` **não** soma delta de abertura.
  `remanescente_inicio` é autoritativo: já reflete tudo com efeito até a abertura,
  inclusive o aditivo datado no próprio dia da abertura. Quem soma esse delta —
  uma única vez — é a coluna `B` da aba `CICLO_EM_EXECUCAO`, que monta o valor a
  partir de `QTD_REM_BASE`. Reaplicá-lo no motor seria dupla contagem, e é isso
  que `test_motor_puro_nao_reaplica_delta_da_abertura` fixa. A janela do período
  no motor permanece `data_inicio < data_efeito ≤ posição`, estritamente
  complementar à abertura.

### 11.6 Por que o espelho `itens_Remanesc!BI`

A formatação condicional dos 4 estados precisava passar a olhar
`CICLO_NASCIMENTO_DATA`. O Excel migra para a **extensão x14** toda regra de CF
que referencie outra planilha, e a x14 é invisível para o openpyxl — a mesma
linhagem que gera cada Coleta no app. Com o espelho local `BI`, a CF continua
OOXML padrão e sobrevive ao round-trip Excel → openpyxl.

### 11.7 Confirmação das travas

`MEMORIA_RESULTADOS!B26` (`VTA_FINAL`) e `T25` permanecem com **fórmula e valor
originais**; `T21`/`T22`/`T23` e as colunas oficiais `G/K/O/S/W` e `Y` não foram
tocadas — nenhuma coluna nova as alimenta. No arquivo real,
`T25 = B26 = 13.468.851,41` antes e depois. O tool `aplicar_temporalidade_aditivos.py`
compara fórmula **e** valor de `B26`/`T25` antes e depois e aborta na divergência;
`tests/test_temporalidade_aditivos_data_efeito.py` repete a trava sobre o arquivo
real e sobre o texto das fórmulas do template.

---

## 12. Regressão final particionada por isolamento do Excel COM

### 12.1 Por que a regressão foi separada em duas trilhas

As rodadas integrais R6, R7 e R8 (suíte inteira num único processo Python, ~80 min)
**não concluíram**. Nenhuma delas apresentou falha de asserção antes da interrupção:
as quedas foram de infraestrutura COM —

* `RPC_E_DISCONNECTED` (0x80010108);
* `RPC_E_CALL_REJECTED` (0x80010001);
* encerramento anômalo do processo Python;
* ausência de sumário final do pytest;
* ausência de `EXIT=0`.

Causa provável: acumulação de muitos módulos Excel COM no **mesmo** processo Python
por tempo prolongado. A resposta não foi mascarar com retentativa nem instalar
plugin de paralelismo (`pytest-xdist`/`pytest-forked` **não** foram instalados),
e sim **particionar** a regressão por natureza de dependência.

### 12.2 Identificação programática dos arquivos COM

Classificação por leitura do fonte, marcadores `DispatchEx`, `win32com`,
`pythoncom`, `Excel.Application`, **mais** a cadeia de imports locais (um teste que
importa `tools/aplicar_*.py` com COM também é COM). Resultado: **19 arquivos COM**
e **61 sem COM**, 80 arquivos no total.

### 12.3 Composição e resultado real das trilhas

| Trilha | Escopo | Processos | Início → Fim | Duração | EXIT | Colet. | Passed | Failed | Skipped |
|---|---|---|---|---|---|---|---|---|---|
| **A** | 61 arquivos sem COM | 1 | 05/08 01:04:16 → 01:40:25 | 36:06 | **0** | 768 | 768 | **0** | 0 |
| **B** | 19 arquivos COM | 19 (um por arquivo) | 05/08 00:38:44 → 01:03:32 | 24:48 | **19 × 0** | 310 | 252 | **0** | 58 |

Trilha A: `pytest -q --no-header -p no:cacheprovider -rf` com 19 `--ignore`,
última linha literal `768 passed in 2166.05s (0:36:06)`.
Log: `.audit_vta/logs/regressao_final_A_sem_com_20260804.log`.

Trilha B: um processo Python **novo** por arquivo, em série, com verificação de
`EXCEL.EXE` visível antes de cada arquivo (aborta e não encerra nada), remoção de
instâncias **invisíveis** residuais e espera de liberação do COM entre arquivos.
Consolidado (versionado): `.audit_vta/logs/regressao_final_B_com_consolidado_20260804.log`.
Logs individuais em `.audit_vta/logs/com_final_20260804/` — evidência **local**, não
versionada: o consolidado já traz, por arquivo, a última linha literal do pytest.

Os 58 skips da Trilha B **não** são regressão: são testes *opt-in* por
`RUN_EXCEL_INTEGRATION=1` (mesma condição em todas as rodadas R1–R10), mais um
golden externo e um legado ausentes. A cobertura profunda de Excel vem do gate
da §12.5, que exercita o COM de ponta a ponta.

### 12.4 Prova de cobertura total

```
768 (sem COM) + 310 (COM) = 1.078
pytest --collect-only -q  ->  1078 tests collected
```

Declaração precisa: **regressão completa validada em duas trilhas por necessidade
de isolamento do Excel COM: 768 testes sem COM em uma execução e 310 testes COM em
19 processos isolados.** Não houve execução integral única aprovada — R6/R7/R8
permanecem apenas como histórico diagnóstico local.

### 12.5 Gate no Excel real (instância nova, não reaproveitada da Trilha B)

Cenários regerados do arquivo real com as ferramentas atuais e submetidos ao ciclo
abrir → recalcular → salvar → fechar → reabrir. Log `EXIT=0` em
`.audit_vta/logs/gate_excel_real_20260805.log` — evidência **local**, não versionada
(os cenários `.xlsx` carregam dados de produção); os resultados estão transcritos
integralmente nas tabelas abaixo.

| Conferência | Cenário A (limpa) | Cenário B (preenchida) | Cenário C (N001/N002/N003) |
|---|---|---|---|
| Erros de fórmula após recálculo | NENHUM | NENHUM | NENHUM |
| Reabertura sem reparo | sim | sim | sim |
| Erros após reabertura | NENHUM | NENHUM | NENHUM |
| `CICLO_EM_EXECUCAO` editável | aba ausente (sem posição atual) | `C13`/`D5` editáveis, `B13` protegida | `C13`/`D5` editáveis, `B13` protegida |
| `itens_RC` alinhado item a item | bloco vazio, como esperado | 21 itens × 5 colunas | 24 itens × 5 colunas |
| `B26` / `T25` | 13.468.851,41 | 13.468.851,41 | 13.474.882,21 (itens sintéticos injetados) |
| FORMA 1 × FORMA 2 | posição não informada | ambas 13.468.851,41 — `W51=0`, **RECONCILIADO** | posição não informada |
| Trava `W55` / `W56` | 0,00 / SEM DUPLA CONTAGEM | 0,00 / SEM DUPLA CONTAGEM | 0,00 / SEM DUPLA CONTAGEM |

"NENHUM" é o resultado de `SpecialCells(xlCellTypeFormulas/Constants, xlErrors)` em
todas as abas: cobre `#REF!`, `#VALUE!`, `#NAME?`, `#DIV/0!` e demais erros.

Classificação temporal confirmada no Excel real (Cenário C, abertura do C1 =
01/02/2026):

| Item | Data de efeito | `AL` | Qtd na abertura C1 | Qtd na abertura C2 | `itens_RC` aplicável | Valor de abertura C1 |
|---|---|---|---|---|---|---|
| N001 | 15/06/2025 (meio do C0) | 1 | 10 | 10 | SIM | 1.030,79 |
| N002 | 12/06/2026 (meio do C1) | 2 | 0 | 20 | NÃO APLICÁVEL | vazio |
| N003 | 01/02/2026 (1º dia do C1) | 1 | 30 | 30 | SIM | 3.092,36 |

N002 não bloqueou a abertura (FORMA 2 fechou normalmente), entrou como alteração
contratual do período por componente próprio (`W54 = 2.000,00`, `W57 = 1`), não
retroagiu e não foi contado duas vezes (`W55 = 0,00`).

### 12.6 Confirmação de `B26`/`T25`

Inalterados nos cenários A e B: `B26 = T25 = 13.468.851,41`, com as fórmulas
homologadas (`B26` referencia `$T$25`; `T25 = ROUND($T$21+$T$22+$T$23,2)`).
No Cenário C o valor muda apenas porque itens sintéticos foram **injetados** de
propósito — não é violação de trava. O SHA-256 do template validado é
`e740a9737245aeaa70573702f6a09d5b5bd59c1c7549a89f75e196711927710e`.
