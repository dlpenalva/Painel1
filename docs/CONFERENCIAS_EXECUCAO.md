# Conferências de execução financeira e física

Base: d97d0d150a49e24f7c5bc4a29a8b459b569aa974, origin/main na preparação.
Branch isolada: fix/vta-conferencias-c0-c4.
Restauração local: restore/vta-conferencias-d97d0d1.

## Diagnóstico reproduzível

Executar `python tools/auditar_conferencias_execucao.py /caminho/1-luc.xlsx`.
A entrada fica fora do Git. O script lê os caches; não recalcula nem grava o XLSX.

Valores confirmados na entrada:

| Medida | Valor (R$) | Origem |
|---|---:|---|
| Contrato original | 1.252.476,00 | MEMORIA_RESULTADOS!B20 |
| VTA oficial | 1.249.462,09 | MEMORIA_RESULTADOS!B26 |
| Financeiro total | 757.074,44 | financeiro!C74 |
| Financeiro C0 | 643.018,39 | soma das competências de C0 |
| Físico original C0 | 783.824,00 | itens_Remanesc!AC2:AC8 |
| Diferença física menos financeira C0 | 140.805,61 | diferença das bases acima |
| Divergência de decomposições | -32.375,95 | MEMORIA_RESULTADOS!W51 |

C73 usa `SUM(N(intervalo))` em contexto escalar: o cache contém apenas AC2,
596.240,00. A linha de total AC9 também está no intervalo. Trocar apenas por
SUM duplicaria a soma. A correção usa SUMIF com identificador de item não vazio,
mantém a checagem de valores numéricos e trata ausência de itens como insuficiência.
C0–C3 usam suas próprias colunas. C4 permanece insuficiente por ausência de
fotografia de fechamento C5 no schema; não se inventa execução física.

O status original B3 não referencia W52 nem as conferências por ciclo. Isso permite
G6=VALIDADO mesmo com W52 indicando decomposições não reconciliadas.
A correção preserva a prioridade REVISE para erros existentes e ESTIMADO, depois
expõe CONFERIR ou BASE INSUFICIENTE. Uma diferença isolada não bloqueia formalização.

## Períodos e resultado do caso

Financeiro C0: 01/09/2025 até 31/08/2026. Físico: abertura de 01/09/2025 até
fotografia de 15/10/2026. Há competências financeiras vazias (5 de 12 no C0).
Até a Coleta 12.7 elas não equivaliam a zero e C0/C1 ficavam BASE INSUFICIENTE;
a partir da Coleta 12.8 equivalem a R$ 0,00 na conferência (ver a seção final),
e C0 passa a COMPARAÇÃO NÃO APLICÁVEL porque os períodos diferem. C1 continua
BASE INSUFICIENTE (sem fotografia seguinte). C2–C4 futuros: NÃO APLICÁVEL. Não classificar os
140.805,61 como erro de pagamento. Quando as duas bases estão completas e seus
períodos diferem, usar COMPARAÇÃO NÃO APLICÁVEL.

A divergência W51 se reproduz assim:

- W48 = 643.018,39 + 118.014,59 + 488.429,11 = 1.249.462,09.
- W50 = 643.018,39 + 118.014,59 + 456.053,16 = 1.217.086,14.
- W50 − W48 = −32.375,95.

O financeiro atualizado do ciclo vigente é o mesmo nas duas expressões.
A diferença vem da redução do remanescente entre a abertura e 27/10/2026:
no item 7, quantidade 0,62 passa a 0,37; 0,25 × 129.503,77 = 32.375,9425,
arredondado por item a 32.375,94. A diferença entre os dois saldos arredondados
é 32.375,95. A competência financeira mais recente é 01/10/2026 e não documenta,
por si, quais entregas até 27/10 foram faturadas. A inclusão dessa execução no
financeiro não pode ser comprovada só por esses totais. Não acrescentar essa
parcela ao VTA nem substituir o financeiro pela medição física.

## VTA e redução frente ao contrato

A memória oficial da entrada é 757.074,44 + 3.958,54 + 488.429,11 =
1.249.462,09: redução de 3.013,91 em relação ao original. RESULTADOS!D27 apresenta
orientação condicional quando VTA_FINAL < contrato original. Uma redução pode
ser coerente com a composição financeira, mas não elimina ressalvas de cobertura,
períodos e conciliação fiscal. Não se declara a redução como erro automaticamente.

Nenhuma fórmula oficial do VTA foi alterada. Antes: 1.249.462,09 (cache).
Depois: esperado idêntico pelas fórmulas preservadas; recálculo real no Excel
continua obrigatório para homologação, não disponível neste ambiente Linux.

## Células automáticas com constantes

Comparação da entrada com o template original: CONTROLE!B10 e itens_Remanesc!B2:B8.
B10 é escrito explicitamente como percentual por `_gerador_masterfile.py`.
B2:B8 são quantidades-base de itens existentes: o campo aceita entrada manual;
a fórmula do template trata o caso específico de novo item com quantidade zero.
Não há evidência de substituição indevida nesses endereços. Nenhum foi reparado.

## Implementação e limites

- `_conferencias_execucao.py`: soma física, períodos, completude, classificações,
  mensagem de VTA inferior e status geral, sem gravar nos dados fiscais.
- `_coleta_oficial.py`: aplicação idempotente no fluxo oficial de geração.
- `_leitor_masterfile_v10.py`: conferências por ciclo no retorno da leitura.
- `_motor_reconciliacao.py`: conserva zeros explícitos, exige contexto temporal
  antes de classificar divergência e remove bloqueio automático por mera diferença.
- Template oficial: C73:C76, E73:E77, B3 de RESULTADOS_DETALHE;
  S280:X284 de MEMORIA_RESULTADOS; D27 e D32 de RESULTADOS.
- `tools/aplicar_conferencias_execucao.py`: edição seletiva do pacote XML,
  preservando outras partes e extensões (requer lxml para manutenção).

A classificação JUSTIFICADO está disponível na API mediante justificativa
explícita; não foi criada nova célula de entrada no leiaute. O fluxo visual de
registro de justificativas e o recálculo Excel são pendências de homologação.
Não há autorização inferida para alterar produção ou fazer merge.

## Evidências de validação nesta branch

- Foco: 68 testes aprovados, 2 pulados (Excel), 1 excluído por importar
  `pythoncom` incondicionalmente no Linux. Inclui regressões estruturais dos
  métodos Financeiro, PC e Consumido e a nova matriz de conferência.
- Ampliação: 101 aprovados, 2 pulados; 1 falha no inventário legado de quantidade
  de fórmulas (`CONTROLE`: esperava 7, encontrou 54). A base d97d0d1 já tem 54;
  essa aba não foi alterada. Não foi atualizada uma referência antiga fora do escopo.
- Comparação XML com a restauração: somente sheet14.xml, sheet15.xml e sheet16.xml
  alterados. Estilos de todas as células preexistentes, atributos das linhas,
  dimensões de colunas, mesclagens, vistas, validações, formatação condicional e
  configuração de página preservados. Todas as outras partes do ZIP são idênticas.
- `git diff --check` e compilação dos módulos aprovados.
- Teste permanente opt-in `RUN_EXCEL_INTEGRATION=1`: avalia C73 com itens
  sintéticos, total no intervalo e resíduos em linhas vazias. Requer Windows/Excel.

Entrega para revisão como PR em rascunho: sem merge ou implantação. Homologação
Excel, atribuição documental da execução física ao financeiro e eventual interface
para justificativas permanecem explicitamente pendentes.

## Continuação da revisão

A versão proposta passa a Cl8us 12.15 / Coleta 12.7, mantendo a aceitação das
versões anteriores. Os marcadores são gravados na geração oficial, no espaço
reservado de CONTROLE; o template base mantém esse espaço vazio como previsto.
A cobertura anual exige uma competência distinta por mês e valores numéricos em
todos os meses; zero explícito é aceito e campo vazio continua insuficiente.
(Regra substituída na Coleta 12.8: campo vazio equivale a R$ 0,00; ver a seção final.)
Os testes de conferência passam a integrar o CI rápido. A ferramenta de manutenção
do template grava o ZIP por substituição atômica, evitando leitura parcial.

Para homologar com Microsoft Excel instalado em Windows:

```powershell
python -m pip install -r requirements.txt pywin32 lxml pytest
python tools/homologar_conferencias_excel.py "C:\caminho\1-luc.xlsx"
$env:RUN_EXCEL_INTEGRATION="1"
python -m pytest -q tests/test_conferencias_execucao.py
```

O comando compara o cache recebido com o recálculo original e o recálculo da
cópia corrigida, salva/reabre somente cópias temporárias e verifica o hash da
entrada. Emite JSON com os valores e falha se o VTA mudar. Sem Excel/pywin32,
encerra com mensagem clara; essa homologação não foi executada no Linux.

Na revisão local de continuação, a execução conjunta aprovou 93 testes e pulou
1 teste de Excel. Duas falhas eram pinos históricos do teste de rollback,
desatualizados em relação à base atual (T58, posição vigente e contagem de
fórmulas). As expectativas foram alinhadas à base e ao novo status, sem mudar
essas fórmulas econômicas no template; a execução focal do rollback aprovou
os 15 testes. A guarda de versionamento passou contra origin/main. Nova
comparação XML confirmou 42 células permitidas e apenas três partes alteradas,
com todos os estilos, dimensões, validações e fórmulas oficiais preservados.
O inventário legado de `test_integridade_template_xlsx.py` continua registrado
como desatualizado na base; não foi usado para substituir a comparação XML.

A comparação da tolerância em Python usa a diferença arredondada em centavos,
assim como D73:D77 no XLS, evitando alerta por resíduo binário no limite exato
de cinco centavos. A matriz focal ampliada passou em 22 casos (sem depender do
Excel ou da geração completa do template).

## Cl8us 12.16 / Coleta 12.8 — status reconhecido e texto de RESULTADOS!C13

Base: cf0b9d3 (main após o merge do #199). Restauração: tag
`restore/pre-sumario-retro-status-20261116`.

Defeito 1 (Python). `RESULTADOS_DETALHE!B3` passou a emitir CONFERIR e
BASE INSUFICIENTE, mas `_resultado_consolidado._STATUS_OFICIAL_PARA_PAINEL`
só conhecia VALIDADO, VALIDADO COM RESSALVAS, ESTIMADO e REVISE. No 1-luc.xlsx
recalculado no Excel, a web mostrava "status oficial indisponível… recalcule,
salve e reenvie" e a formalização ficava AGUARDA CONFIRMAÇÃO. Agora os dois
códigos são conclusões próprias (não viram VALIDADO, REVISE nem indisponível),
publicam uma ressalva e não bloqueiam a formalização por si sós. Bloqueios
econômicos (`bloqueios_formalizacao`) continuam bloqueando; sem VTA ou sem
método o status continua PENDENTE. O painel usa a classe visual de ressalva.

Defeito 2 (apresentação do XLS). `RESULTADOS!C12/C13` seguem a regra de
pendência de preenchimento (`RESULTADOS_DETALHE!A7`, `$J$5=0`), que não enxerga
as conferências; com G6=CONFERIR, C13 dizia "Nenhuma pendência registrada pelas
validações desta apuração". Só C13 mudou: com A7 = APURAÇÃO CONCLUÍDA e status
CONFERIR/BASE INSUFICIENTE, explica a conferência informativa e anexa o texto
de potencial quando houver; nos demais casos continua igual a
`MEMORIA_RESULTADOS!T67`. C12, T67, B7 e fórmulas econômicas intactos (a cor
da faixa foi tratada depois; ver a seção seguinte).

Retroativo. A divergência alegada do Sumário Executivo não foi reproduzida.
Retroativo total, reconhecido, pago, em análise e potencial conferem com
RESULTADOS!C9:E9 e MEMORIA_RESULTADOS!B16 (R$ 3.958,54; pago e em análise
R$ 0,00; potencial não se aplica ao Financeiro). Origem no PDF: delta mensal de
`financeiro` com EFEITO_FINANCEIRO=Sim → evento "Retroativo Reconhecido" →
`_assistente_fiscal._classificar_retroativo` → `objeto.resultados.retroativo`
→ `_sumario_executivo`. Nenhuma alteração foi feita nessa cadeia.

Regressão: `tests/test_status_conferir_base_insuficiente.py` (três métodos,
ausência, zero legítimo, bloqueio econômico, texto desconhecido e fórmula C13).

### Pendências finais do PR #200 (mesma versão 12.16 / 12.8)

Restauração: tag `restore/pre-pendencias-pr200-20261117` (= `76a23c6`).

Sumário Executivo — "Em análise". O valor vem só de PCs
(`_assistente_fiscal._classificar_retroativo`: natureza "potencial"); sem PCs
ele é 0,0 por construção. A planilha mostra "—" em `RESULTADOS!D9` sempre que
`METODO_RETROATIVO` ≠ "PCs". O PDF passa a exibir "Não aplicável" somente
quando as três condições valem: método conhecido "Financeiro" ou "Itens",
nenhum PC com evidência e valor em análise nulo
(`_sumario_executivo._em_analise_nao_aplicavel`). Método PCs, método
desconhecido ou qualquer valor diferente de zero mantêm o número (zero apurado
continua R$ 0,00). Cálculo, estados e valores não mudam.

RESULTADOS — faixa B12:G13. O verde C6EFCE/006100 é reservado em RESULTADOS a
VALIDADO/RECONCILIADO (G6, C27:C34), mas a faixa ficava verde com CONFERIR e
BASE INSUFICIENTE porque só olhava C12 ("APUR…") e D9. Uma regra nova, logo
após REVISE (sem `stopIfTrue`, como as demais da aba: a prioridade decide),
pinta a faixa com o âmbar de ressalva
(FFEB9C/9C5700, o mesmo dxf de ESTIMADO em G6) quando C12 = APURAÇÃO CONCLUÍDA
e o status é CONFERIR/BASE INSUFICIENTE. Aplicada por
`tools/aplicar_cor_status_resultados.py` (patch XML, idempotente, localiza as
regras pela fórmula porque o Excel as reordena e renumera dxfId). Só
`sheet16.xml` do template muda; styles.xml intacto. Demais estados iguais:
VALIDADO verde, REVISE vermelho, potencial PC oliva, ESTIMADO como antes.

Homologação em cópias do 1-luc.xlsx recalculadas no Excel (Financeiro, PC,
Itens; status forçados BASE INSUFICIENTE, VALIDADO e ESTIMADO): VTA
R$ 1.249.462,09 e retroativo R$ 3.958,54 preservados, nenhum valor de célula
diferente antes × depois, nenhum alerta de reparo. Regressão em
`tests/test_status_conferir_base_insuficiente.py`.

### Competências sem lançamento equivalem a R$ 0,00 (mesma versão 12.16 / 12.8)

Restauração: tag `restore/pre-competencias-vazias-20261117` (= `f51cf5a`).

Regra de negócio: na aba `financeiro`, competência listada no período do ciclo
sem lançamento de valor (VALOR_PAGO vazio) equivale a R$ 0,00 para fins de
conferência. Nenhuma célula é preenchida e nenhum cálculo oficial muda: a soma
de `RESULTADOS_DETALHE!B73:B77` (SUMIFS) já tratava o vazio como zero; o que
mudou foi apenas a regra de cobertura.

- Excel, `MEMORIA_RESULTADOS!X280:X284`: a exigência "12 valores numéricos"
  passou a "nenhum valor preenchido não numérico"
  (`SUMPRODUCT(--(B="cN"),--(C<>""),--NOT(ISNUMBER(C)))=0`).
- Python, `_conferencias_execucao.conferir_workbook`: mesma regra no campo
  `completo` (vazio = `None` ou `""`; qualquer outro valor não numérico invalida).

Preservado: 12 linhas por ciclo com exatamente uma competência por mês (duplicada
ou ausente = BASE INSUFICIENTE), datas válidas no dia 1 dentro do período, texto
no valor = BASE INSUFICIENTE, períodos diferentes = COMPARAÇÃO NÃO APLICÁVEL,
ciclo futuro = NÃO APLICÁVEL, ciclo vigente sem fotografia seguinte = BASE
INSUFICIENTE. Ciclo sem nenhum lançamento continua BASE INSUFICIENTE (B73 vazio):
sem qualquer valor não há base para a conferência.

Homologação no Excel real (cópias do 1-luc): "7 pagos + 5 vazios" e "7 pagos +
5 zeros explícitos" produzem leituras idênticas (X280:X284, B73:E77, status,
VTA, retroativo) no Excel e no Python; antes da mudança, vazio dava BASE
INSUFICIENTE e zeros, COMPARAÇÃO NÃO APLICÁVEL. Texto em C2 volta X280 a 0.
VTA R$ 1.249.462,09 e retroativo R$ 3.958,54 inalterados. Testes em
`tests/test_conferencias_execucao.py`.
