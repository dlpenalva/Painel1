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
fotografia de 15/10/2026. Há competências financeiras vazias; não equivalem a zero.
C0 e C1: BASE INSUFICIENTE. C2–C4 futuros: NÃO APLICÁVEL. Não classificar os
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
