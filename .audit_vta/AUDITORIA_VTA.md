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
