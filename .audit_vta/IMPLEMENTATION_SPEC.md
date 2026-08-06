# SPEC DE IMPLEMENTAÇÃO — pronta para execução determinística

Fonte: `.audit_vta/AUDITORIA_VTA.md` + memória `project_vta_posicoes_resultados_itens_rc`.
Travas: **NÃO alterar B26/T25**; sem push/PR/merge/deploy; parar após commit local.
Método de mutação: **Excel COM** em cópia temporária → recalc → verificar 0 erros
Excel → reabrir sem reparo → promover template (padrão de
`tools/aplicar_resultados_consolidados_26f.py`). Alternativa aceitável: openpyxl
para autoria + **gate Excel COM** (abrir/recalc/salvar/reabrir sem reparo) antes do commit.

Downstream lê só nomes definidos (`VTA_FINAL`=B26 etc.); `RESULTADOS!B10/B11` e
`VTA_ATUALIZACAO_CHEIA` NÃO são lidos por coordenada → redesenho de Tabela 1 é seguro.

Referência ao runtime `CICLO_EM_EXECUCAO` (aba opcional, ausente no template):
usar **INDIRECT** + ISERROR (nunca referência dura — evita #REF! e reparo).

---

## A. MEMORIA_RESULTADOS — bloco auxiliar novo (motor das 3 referências)

Colunas livres: **V** (rótulo) e **W** (valor). Coluna aux por item: **AB (28)**.
S/T ocupados até linha 30; usar V/W a partir da linha 40 (área livre; A40 tem só
rótulo "APOIO AO CÁLCULO MANUAL", colunas V/W estão vazias).

### A.1 Completude de abertura por ciclo (n=0..4) — W41:W45
Espelha a lógica de `T26`/`T24`, porém por ciclo n (não só vigente). Para cada n,
coluna de remanescente ajustado em posicao_contratual: C0→G, C1→K, C2→O, C3→S, C4→W.
Item só é exigido no ciclo n se `CICLO_NASCIMENTO (Y) <= n`.

```
V41="Abertura C0 completa?" ; W41 =
 =IF(COUNTIF(posicao_contratual!$A$2:$A$201,"<>")=0,0,
   IF(SUMPRODUCT((posicao_contratual!$A$2:$A$201<>"")*(posicao_contratual!$Y$2:$Y$201<=0)*
      (1-ISNUMBER(posicao_contratual!$G$2:$G$201)))>0,0,1))
V42="Abertura C1 completa?" ; W42 = (idem, coluna K, Y<=1)
V43="Abertura C2 completa?" ; W43 = (idem, coluna O, Y<=2)
V44="Abertura C3 completa?" ; W44 = (idem, coluna S, Y<=3)
V45="Abertura C4 completa?" ; W45 = (idem, coluna W, Y<=4)
```
(Cada Wn escreve a coluna literal por n; sem CHOOSE.)

### A.2 Seleção da última abertura disponível ≤ vigente — W46
```
V46="Ciclo da última abertura completa (num)" ; W46 =
 =IF($T$20="","",
   IF(AND($T$20>=4,INDEX($W$41:$W$45,5)=1),4,
   IF(AND($T$20>=3,INDEX($W$41:$W$45,4)=1),3,
   IF(AND($T$20>=2,INDEX($W$41:$W$45,3)=1),2,
   IF(AND($T$20>=1,INDEX($W$41:$W$45,2)=1),1,
   IF(INDEX($W$41:$W$45,1)=1,0,""))))))
V47="Motivo de não adoção das aberturas posteriores" ; W47 =
 =IF(OR($W$46="",$W$46=$T$20),"",
   "Aberturas posteriores a C"&$W$46&" incompletas (ciclo vigente C"&$T$20&").")
```

### A.3 Aux por item: remanescente na abertura selecionada — AB2:AB201
Espelha `Y2` (=ROUND(VU_sel,2)*QREM_sel) mas com `$W$46` no lugar de `$T$20`:
```
AB2 =
 =IF(OR(posicao_contratual!$A2="",$W$46=""),0,
   IF(AND(ISNUMBER(CHOOSE($W$46+1,historico_VU!$C2,historico_VU!$D2,historico_VU!$E2,historico_VU!$F2,historico_VU!$G2)),
          ISNUMBER(CHOOSE($W$46+1,posicao_contratual!$G2,posicao_contratual!$K2,posicao_contratual!$O2,posicao_contratual!$S2,posicao_contratual!$W2))),
     ROUND(ROUND(CHOOSE($W$46+1,historico_VU!$C2,historico_VU!$D2,historico_VU!$E2,historico_VU!$F2,historico_VU!$G2),2)*
           CHOOSE($W$46+1,posicao_contratual!$G2,posicao_contratual!$K2,posicao_contratual!$O2,posicao_contratual!$S2,posicao_contratual!$W2),2),""))
```
(preencher AB2:AB201, 200 fórmulas.)

### A.4 FORMA 2 — VTA pela última posição de abertura disponível — W48
```
V48="FORMA 2 — VTA última abertura disponível" ; W48 =
 =IF($W$46="","",
   ROUND($T$21
     + SUMPRODUCT((ROW(itens_PC!$P$2:$P$6)-2>=1)*(ROW(itens_PC!$P$2:$P$6)-2<$W$46)*itens_PC!$P$2:$P$6)
     + SUM($AB$2:$AB$201),2))
```
**Verificação obrigatória:** quando `W46=T20` (abertura vigente completa),
`W48` deve ser IDÊNTICO a `T25` (=13.468.851,41 no arquivo real). Se divergir,
NÃO ajustar para forçar — investigar (trava Seção 21).

### A.5 FORMA 1 — VTA pela posição atual (via CICLO_EM_EXECUCAO, opcional) — W49:W50
```
V49="CICLO_EM_EXECUCAO disponível/utilizado?" ; W49 =
 =IF(ISERROR(INDIRECT("CICLO_EM_EXECUCAO!A9")),0,
   IF(INDIRECT("CICLO_EM_EXECUCAO!A9")="",0,1))
V50="FORMA 1 — VTA posição atual" ; W50 =
 =IF(OR($W$49=0,$T$21="",NOT(ISNUMBER($T$21))),"",
   ROUND($T$21+$T$22
     + SUM(INDIRECT("CICLO_EM_EXECUCAO!F13:F211"))   /* consumido no ciclo até a data */
     + SUM(INDIRECT("CICLO_EM_EXECUCAO!G13:G211")),2)) /* remanescente atual na data */
```
Nota: A9 vazio ⇒ aba incompleta/erros (fórmula A9 já zera nesses casos) ⇒ FORMA 1 "".

### A.6 Reconciliação FORMA 1 × FORMA 2 — W51:W52
```
V51="Reconciliação (Posição atual − Última abertura)" ; W51 =
 =IF(OR($W$50="",$W$48=""),"",ROUND($W$50-$W$48,2))
V52="Status reconciliação" ; W52 =
 =IF($W$50="","POSIÇÃO ATUAL NÃO INFORMADA",
   IF($W$48="","ÚLTIMA ABERTURA INDISPONÍVEL",
   IF(ABS($W$51)<=$D$4,"RECONCILIADO","REVISE — DECOMPOSIÇÕES DO VTA NÃO RECONCILIADAS")))
```

---

## B. RESULTADOS — Tabela 1 redesenhada (linhas 8–13; NÃO mexer em ≥14)

Manter A8 como título de seção; ampliar cabeçalho (linha 9) e usar 3 linhas de dados
(10,11,12) + reconciliação (13). Colunas: A Referência | B Valor | C:E Composição |
F:G Fontes | H Situação. Mesclar C10:E10 etc. para composição textual.

```
A8 = "1. VALOR TOTAL DO CONTRATO — TRÊS REFERÊNCIAS DO VTA"
H8 = (manter fórmula de status atual; opcional: refletir W52)

A9="Referência" | B9="Valor" | C9="Composição auditável (aba!célula)" |
F9="Fontes" | H9="Situação"        (mesclar C9:E9 e F9:G9)

Linha 10 — FORMA 1:
 A10="VTA PELA POSIÇÃO ATUAL DO CONTRATO"
 B10==IF(MEMORIA_RESULTADOS!$W$50="","",MEMORIA_RESULTADOS!$W$50)
 C10==IF(MEMORIA_RESULTADOS!$W$50="","POSIÇÃO ATUAL NÃO INFORMADA (CICLO_EM_EXECUCAO ausente/incompleta)",
        "C0 exec (MEMORIA!T21) + ciclos encerrados (MEMORIA!T22) + execução até a data + remanescente atual (CICLO_EM_EXECUCAO!F/G)")
 F10=="CICLO_EM_EXECUCAO (fiscal) + itens_PC (C0/encerrados)"
 H10==IF(MEMORIA_RESULTADOS!$W$50="","INDISPONÍVEL — POSIÇÃO ATUAL NÃO INFORMADA",
        IF(MEMORIA_RESULTADOS!$W$52="RECONCILIADO","DISPONÍVEL PARA CONFERÊNCIA",MEMORIA_RESULTADOS!$W$52))

Linha 11 — FORMA 2 (adotada quando FORMA 1 indisponível):
 A11="VTA PELA ÚLTIMA POSIÇÃO DE ABERTURA DISPONÍVEL"
 B11==IF(MEMORIA_RESULTADOS!$W$48="","",MEMORIA_RESULTADOS!$W$48)
 C11=="C0 exec (MEMORIA!T21) + ciclos encerrados (MEMORIA!T22) + remanescente na abertura C"&IF(MEMORIA_RESULTADOS!$W$46="","?",MEMORIA_RESULTADOS!$W$46)&" (MEMORIA!AB)"
 F11=="itens_PC!P + posicao_contratual (G/K/O/S/W) × historico_VU"
 H11==IF(MEMORIA_RESULTADOS!$W$48="","INCOMPLETO",
        IF(MEMORIA_RESULTADOS!$W$46=MEMORIA_RESULTADOS!$T$20,
           "ADOTADO — ABERTURA DO CICLO VIGENTE C"&MEMORIA_RESULTADOS!$T$20,
           "ADOTADO — ÚLTIMA ABERTURA COMPLETA: C"&MEMORIA_RESULTADOS!$W$46))
 (C11 ou coluna F anexar W47 = motivo de não adoção das posteriores, quando houver.)

Linha 12 — FORMA 3:
 A12="CONTRATO ORIGINAL INTEGRALMENTE REAJUSTADO"
 B12==IFERROR(comparativo_VTA!$B$208,"")
 C12=="SUMPRODUCT(posicao_contratual!B×C) × CONTROLE!B11 (fator histórico integral)"
 F12=="comparativo_VTA!B208"
 H12="COMPARATIVO — NÃO É VTA OFICIAL"

Linha 13 — Reconciliação:
 A13="Reconciliação (Posição atual − Última abertura)"
 B13==IF(MEMORIA_RESULTADOS!$W$51="","",MEMORIA_RESULTADOS!$W$51)
 H13==MEMORIA_RESULTADOS!$W$52
```
Nome definido `VTA_ATUALIZACAO_CHEIA`: repointar para `RESULTADOS!$B$12` (mantém
sentido "atualização cheia" = contrato integralmente reajustado). Nada o lê hoje.

Identidade visual CICLO_EM_EXECUCAO: verde=abertura/remanescente, laranja=consumo,
azul=valores, amarelo=data/destaque, "(AUTO)" nos automáticos.

---

## C. itens_RC — bloco "POSIÇÃO ATUAL (AUTO)" à direita (colunas Q+)

Preserva A:P (aberturas C0..C4). Cabeçalho em Q1 (merge Q1:Y1) =
"POSIÇÃO ATUAL NA DATA (AUTO — origem: CICLO_EM_EXECUCAO)". Sub-cabeçalho linha 2.
Origem por item = linha correspondente da aba (itens_RC dados linha 3 ↔ CICLO linha
13; deslocamento +10). Guardar tudo por INDIRECT+ISERROR.

Colunas (Seção 12), linhas 3..202:
```
Q = DATA DA POSIÇÃO ATUAL (AUTO)          = IF(ISERROR(INDIRECT("CICLO_EM_EXECUCAO!$D$5")),"",INDIRECT("CICLO_EM_EXECUCAO!$D$5"))
R = CICLO DA ÚLTIMA ABERTURA (AUTO)       = IF(MEMORIA_RESULTADOS!$W$46="","","C"&MEMORIA_RESULTADOS!$W$46)
S = QTD REMANESCENTE NA ABERTURA (AUTO)   = IF(ISERROR(INDIRECT("CICLO_EM_EXECUCAO!B"&(ROW()+10))),"",INDIRECT("CICLO_EM_EXECUCAO!B"&(ROW()+10)))
T = QTD CONSUMIDA DESDE ABERTURA (AUTO)   = ...INDIRECT("CICLO_EM_EXECUCAO!D"&(ROW()+10))
U = QTD REMANESCENTE NA DATA ATUAL (AUTO) = ...INDIRECT("CICLO_EM_EXECUCAO!C"&(ROW()+10))
V = VU ATUALIZADO NA DATA (AUTO)          = ...INDIRECT("CICLO_EM_EXECUCAO!E"&(ROW()+10))
W = VALOR CONSUMIDO ATUALIZADO (AUTO)     = ...INDIRECT("CICLO_EM_EXECUCAO!F"&(ROW()+10))
X = VALOR REMANESCENTE ATUALIZADO (AUTO)  = ...INDIRECT("CICLO_EM_EXECUCAO!G"&(ROW()+10))
Y = STATUS DA POSIÇÃO ATUAL (AUTO)        = IF(ISERROR(INDIRECT("CICLO_EM_EXECUCAO!K"&(ROW()+10))),"CICLO_EM_EXECUCAO AUSENTE",INDIRECT("CICLO_EM_EXECUCAO!K"&(ROW()+10)))
```
Bloco inteiro vazio quando aba ausente (ISERROR→""); nunca digitação manual aqui
(motor único = CICLO_EM_EXECUCAO). Reconciliação (Seção 15): itens_RC consome os
mesmos valores da aba ⇒ diferença 0 por construção (mesma fonte).
ATENÇÃO: `ROW()+10` só vale se itens_RC dados começam na linha 3 e CICLO na 13;
confirmar antes (linha 3→13, 4→14, ...). Alternativa robusta: passar o número da
linha origem literal por célula (gerado no loop Python).

---

## D. Contagens de integridade a atualizar (tests/test_integridade_template_xlsx.py)

Recontar após mutação e ajustar `FORMULAS_POR_ABA`:
- `itens_RC`: 3200 → +Δ (Q:Y ≈ 9 col × 200 linhas = ~1800; conferir exato).
- `RESULTADOS`: atual (~3773) → +Δ (Tabela 1 novas fórmulas B/C/F/H linhas 10–13).
- `MEMORIA_RESULTADOS`: atual → +Δ (W41:W52 + AB2:AB201 ≈ 212).
Regra: rodar o teste, ler contagem real, fixar o número novo (não estimar).

---

## E. 14 testes obrigatórios (tests/test_vta_posicoes_resultados.py — novo)

1. Arquivo real: T21+T22+T23 = 13.468.851,41 (data_only, arquivo INBLOKO).
2. Posição atual C1 = C0 + exec C1 até data + rem atual C1; comparar c/ última abertura.
3. Posição atual C2 (cenário sintético) = C0 + C1 exec + exec C2 + rem atual C2.
4. Vigente C3, abertura C3 incompleta, C2 completa ⇒ W46=2, W48 usa C2, W47 motivo visível, sem dupla contagem.
5. Sem CICLO_EM_EXECUCAO ⇒ W50="" (FORMA1 indisponível), W48 calculado, itens_RC Q vazio, aberturas preservadas.
6. CICLO_EM_EXECUCAO incompleta ⇒ A9 vazio ⇒ W49=0 ⇒ FORMA1 sem total; W48 disponível.
7. Posição válida ⇒ itens_RC recebe valores idênticos à aba; check 0.
8. Decomposições divergentes ⇒ W52="REVISE...", B13 exibe diferença, sem adoção silenciosa.
9. Item incluído em C0: abertura C1 vazia=pendência; zero=válida; positiva=válida (posicao_contratual F/Y).
10. Item incluído no meio de C1: abertura C1 NÃO APLICÁVEL (Y=1, K guard); aparece na posição atual após data; aparece na abertura C2 se saldo.
11. Item no 1º dia do ciclo: abertura aplicável; aditivo não reaplicado.
12. Execução física implícita C0 (col X): qtd=(G−K), valor=(G−K)×VU_C0; fonte adotada; divergência vs PCs.
13. Contrato integralmente reajustado permanece comparativo; nunca adotado (H12 fixo).
14. Regressão: VTA oficial B26 inalterado; nomes definidos intactos; opcionalidade CICLO_EM_EXECUCAO.

---

## G. Temporalidade dos aditivos por DATA_EFEITO (executada)

Fonte: `tools/aplicar_temporalidade_aditivos.py`.
Cobertura: `tests/test_temporalidade_aditivos_data_efeito.py`.

### G.1 posicao_contratual — colunas novas `AA:AL`, linhas 2:200 (ocultas)

```
AA  DATA_EFEITO_INICIAL
    =IF($A2="","",IF(COUNTIFS(<criterios>)=0,"",
       IFERROR(_xlfn.MINIFS(aditivos!$B$2:$B$200,<criterios>),"")))
    <criterios> = aditivos!$A$2:$A$200,$A2, aditivos!$L$2:$L$200,"<>",
                  aditivos!$L$2:$L$200,"<>0", aditivos!$B$2:$B$200,">0"

AB..AF  DELTA_POSTERIOR_ABERTURA_C0..C4   (parametros!$C$2..$C$6)
    =IF($A2="","",IF(parametros!$C$3="",0,
       ROUND(SUMIFS(aditivos!$L$2:$L$200,
                    aditivos!$A$2:$A$200,$A2,
                    aditivos!$C$2:$C$200,"C1",
                    aditivos!$B$2:$B$200,">="&(INT(parametros!$C$3)+1)),2)))

AG..AK  QTD_CONTRATUAL_ABERTURA_C0..C4    (E/I/M/Q/U menos AB/AC/AD/AE/AF)
    =IF($A2="","",IF(OR(NOT(ISNUMBER($I2)),NOT(ISNUMBER($AC2))),"",
       ROUND($I2-$AC2,2)))

AL  CICLO_NASCIMENTO_DATA
    =IF($A2="","",IF(N($AG2)>0,0,IF(N($AH2)>0,1,IF(N($AI2)>0,2,
       IF(N($AJ2)>0,3,IF(N($AK2)>0,4,""))))))
```

**Fronteira por dia** (`>=INT(abertura)+1`): imuniza a classificação contra
células de data com componente de hora — o modo como o aditivo do 1º dia do ciclo
era indevidamente empurrado para "posterior".

### G.2 MEMORIA_RESULTADOS — decomposição temporal da FORMA 2

```
AD2:AD201  componente ALTERACOES POSTERIORES A ABERTURA
    =IF(OR(posicao_contratual!$A2="",$W$46=""),0,
      IF($AB2="","",IF(NOT(ISNUMBER(<post>)),0,ROUND(ROUND(<vu>,2)*<post>,2))))
    <vu>   = CHOOSE($W$46+1,historico_VU!$C2..$G2)
    <post> = CHOOSE($W$46+1,posicao_contratual!$AB2..$AF2)

AC2:AC201  componente REMANESCENTE NA ABERTURA (residual — additividade exata)
    =IF(OR(posicao_contratual!$A2="",$W$46=""),0,
      IF(OR($AB2="",$AD2=""),"",ROUND($AB2-$AD2,2)))

W53 =SUM($AC$2:$AC$201)
W54 =SUM($AD$2:$AD$201)
W55 =ROUND(SUM($AB$2:$AB$201)-($W$53+$W$54),2)      <- trava, deve ser 0,00
W56 =IF(ABS($W$55)<=$D$4,"SEM DUPLA CONTAGEM","REVISE - DUPLA CONTAGEM NA FORMA 2")
W57 =IF($W$46="","",SUMPRODUCT((posicao_contratual!$A$2:$A$200<>"")
      *ISNUMBER(posicao_contratual!$AL$2:$AL$200)
      *(posicao_contratual!$AL$2:$AL$200>$W$46)))

W48 (REESCRITA)
    =IF($W$46="","",ROUND($T$21
       +SUMPRODUCT((ROW(itens_PC!$P$2:$P$6)-2>=1)*(ROW(itens_PC!$P$2:$P$6)-2<$W$46)
                   *itens_PC!$P$2:$P$6)
       +$W$53+$W$54,2))
```

Como `AC ≡ AB − AD`, vale `W53 + W54 ≡ SUM(AB)`: o **total** da FORMA 2 é
idêntico ao anterior; muda a leitura temporal das parcelas. O aditivo posterior à
abertura **não é excluído** do VTA auditável — entra por componente próprio, uma
única vez.

### G.3 itens_RC — bloco `Z:AC` (linhas 3:202, merge `Z1:AC1`)

`Z` data de efeito · `AA` ciclo de nascimento por data · `AB` `SIM`/`NAO APLICAVEL`
na abertura adotada (`W46`) · `AC` qtd contratual na abertura adotada.
A fotografia oficial `A:P` permanece intacta.

### G.4 itens_Remanesc

* `BI2:BI200` — espelho local de `posicao_contratual!AL`. Necessário porque o
  Excel migra para a extensão **x14** toda regra de formatação condicional que
  referencie outra planilha, e a x14 é invisível ao openpyxl (linhagem que gera
  cada Coleta). Com o espelho, a CF continua OOXML padrão.
* CF dos 4 estados (`E/G/I/K`) passa a usar `$BI2 > n` na regra
  **NÃO APLICÁVEL** (com `stopIfTrue`) e na regra de pendência.
* `F/H/J/L` (VALOR_REM_INICIO_Cn) ganham a guarda
  `AND(ISNUMBER(posicao_contratual!$AL2),posicao_contratual!$AL2>n)` ⇒ vazio.

### G.5 _ciclo_em_execucao.py

* `_formula_abertura` — coluna `B` soma os deltas do ciclo com
  `data < INT($F$3)+1` sobre a base do fiscal (Regra B: delta do 1º dia entra na
  abertura, uma única vez).
* Coluna `I` (alterações do período) passa a `>= INT($F$3)+1`, complementar à
  abertura — sem sobreposição, sem lacuna.
* Motor puro `calcular_posicao_ciclo_por_data` — **inalterado no cálculo**:
  `remanescente_inicio` é autoritativo e já contém o delta com efeito até a
  abertura; o motor **não** o reaplica (seria dupla contagem). A janela do
  período segue `data_inicio < data_efeito ≤ posição`, complementar à abertura.
  Quem soma o delta da abertura, uma única vez, é a coluna `B` da aba.
  Trava: `test_motor_puro_nao_reaplica_delta_da_abertura`.

### G.6 Ordem de mutação (importante)

1. openpyxl: fórmulas **e** formatação condicional;
2. Excel COM: recalcular, verificar zero erros, conferir `W55`, salvar;
3. reabrir somente-leitura: conferir travas `B26`/`T25` (fórmula **e** valor);
4. openpyxl: conferir que a CF sobreviveu como OOXML padrão (aborta se migrou);
5. promover.

A CF **não** pode ser a última escrita (como em `aplicar_ajustes_finais_layout.py`):
openpyxl não regrava valores em cache, e os leitores usam `data_only=True`.

---

## F. Gate final (ordem)
1. Aplicar mutação (Excel COM temp → recalc → 0 erros → salvar → reabrir sem reparo → promover template).
2. `pytest -q` (regressão integral + 14 novos) — todos PASS.
3. Atualizar contagens integridade.
4. Excel COM: abrir template promovido + gerar Coleta do arquivo real, confirmar sem reparo, B26 = 13.468.851,41.
5. `git add` + commit local (mensagem descritiva; Co-Authored-By). SEM push/PR/deploy.
6. Relatório final (deliverables 1–17 da Seção 22).

---

## H. Regressão particionada por isolamento do Excel COM (executada)

### H.1 Motivo

A suíte inteira num único processo Python (~80 min) não conclui: R6/R7/R8 caíram
por `RPC_E_DISCONNECTED` (0x80010108), `RPC_E_CALL_REJECTED` (0x80010001) e
encerramento anômalo do processo, **sem** nenhuma falha de asserção antes da
interrupção — e portanto sem sumário do pytest e sem `EXIT=0`. A causa provável é
o acúmulo de módulos Excel COM no mesmo processo. Nenhuma dependência nova foi
instalada (`pytest-xdist`/`pytest-forked` continuam **fora** do projeto).

### H.2 Como separar (não usar lista manual)

Classificar por leitura do fonte: marcadores `DispatchEx`, `win32com`, `pythoncom`,
`Excel.Application` **e** a cadeia de imports locais (teste que importa
`tools/aplicar_*.py` com COM conta como COM). Hoje: **19 COM / 61 sem COM**.

### H.3 Trilha A — sem COM, uma única rodada

```
python -m pytest -q --no-header -p no:cacheprovider -rf --ignore=<19 arquivos COM>
```

Log: `.audit_vta/logs/regressao_final_A_sem_com_20260804.log`.
Critério: 768 coletados, `EXIT=0`, 0 failed, sumário real presente.

> Armadilha do runner `.bat`: em `echo EXIT=%ERRORLEVEL%>> log` o cmd lê o dígito
> colado no `>>` como **handle** (`0`=stdin, `1`=stdout) — a linha desaparece do log
> exatamente quando o código é `0`. Gravar como `>> "%LOG%" echo EXIT=%ERRORLEVEL%`
> (redirecionamento antes) e usar `cmd /v:on` + `!ERRORLEVEL!` quando houver
> encadeamento na mesma linha.

### H.4 Trilha B — COM, um processo por arquivo

Serial, nunca em paralelo, nunca dois arquivos COM no mesmo processo. Por arquivo:
verificar `EXCEL.EXE` **visível** (se houver, abortar sem encerrar nada), remover
instâncias **invisíveis** residuais, rodar `pytest tests/<arquivo>` em processo novo,
registrar sumário e código de saída, limpar Excel invisível, aguardar liberação do COM.

Consolidado: `.audit_vta/logs/regressao_final_B_com_consolidado_20260804.log`
(colunas: arquivo, coletados, passed, failed, skipped, exit, duração, log individual).
Individuais: `.audit_vta/logs/com_final_20260804/`.

### H.5 Higiene obrigatória do ambiente

Antes de qualquer rodada, confirmar que **não** há regressão integral ativa. Durante
esta etapa foi encontrada uma rodada integral (`R9`) já em curso e, mais tarde, uma
`R10` disparada por mecanismo destacado de sessão anterior, que rodou concorrente à
primeira Trilha B e invalidou a premissa de isolamento — a trilha foi **refeita** em
ambiente limpo e o lançador integral foi neutralizado
(`run_regressao_integral.bat.DESABILITADO`). Detector correto de intruso:
`python.exe -m pytest` **sem** arquivo alvo (procurar pelo nome do `.bat` gera falso
positivo, pois o próprio comando de busca casa consigo mesmo).

### H.6 Prova de cobertura

```
768 + 310 = 1.078  ==  pytest --collect-only -q  ->  1078 tests collected
```

Declarar sempre: **duas trilhas por isolamento do COM**, nunca "uma execução integral".

---

## F. Gate final (ordem)
1. Aplicar mutação (Excel COM temp → recalc → 0 erros → salvar → reabrir sem reparo → promover template).
2. Regressão particionada (Seção H): Trilha A `EXIT=0` + Trilha B 19 × `EXIT=0`, 0 failed.
3. Atualizar contagens integridade.
4. Excel COM em instância NOVA (`DispatchEx`, nunca reaproveitar a da Trilha B):
   Coleta limpa e Coleta preenchida — abrir, recalcular, salvar, fechar, reabrir sem
   reparo; zero erros de fórmula; `CICLO_EM_EXECUCAO` editável; `itens_RC` alinhado;
   N001/N002/N003 classificados; FORMA 1 × FORMA 2 reconciliadas; B26 = 13.468.851,41.
5. `git add` + commit local (mensagem descritiva; Co-Authored-By). SEM push/PR/deploy.
6. Relatório final (deliverables 1–17 da Seção 22).
