# Relatório UX — Master 2.0: ciclos C2, histórico condicional e carimbo de versão

Data da revisão: 13/07/2026

## Decisão de produto

O Master 2.0 mantém o XLS como produto principal e usa a web somente para:

1. escolher o ciclo ou o intervalo atual;
2. apurar índice, admissibilidade e marcos;
3. baixar o `Coleta_Reajuste.xlsx` parametrizado;
4. receber o mesmo XLS preenchido para validação e resultados.

Quando a análise começa em C2, C3 ou C4, a web coleta somente o contexto mínimo necessário para definir a âncora correta: situação anterior, último ciclo formalizado e marco temporal. Percentuais, valores e demais fatos históricos permanecem no XLS, onde podem ser auditados pelo fiscal/GCC.

## Mudanças realizadas

- O cálculo simples aceita explicitamente C1, C2, C3 ou C4.
- O multiciclo aceita qualquer intervalo contíguo entre C1 e C4, inclusive C2 → C3 ou C2 → C4.
- Ao iniciar em C2, C3 ou C4, a interface replica o fluxo condicional validado do 3.0.
- Sem ciclo anterior concedido, a data-base segue a linha anual da âncora original.
- Com ciclo anterior formalizado, a data-base usa o marco informado e avança apenas os ciclos intermediários.
- Em situação desconhecida, a ferramenta não inventa marco: informa a incerteza e mantém a linha anual.
- O formulário amplo “Contexto do Contrato” permanece removido; apenas o bloco mínimo e condicional foi reintroduzido.
- O XLS continua preservando C0 e os ciclos anteriores necessários à memória e ao fator acumulado.
- A navegação adota o contraste, a sidebar azul-clara e o indicador circular do protótipo 3.0.
- A sidebar exibe permanentemente `Atualizado em dd/mm/aaaa hh:mm`, usando o último commit e fallback de release.

## Checagem crítica

- O usuário só declara histórico na web quando inicia após C1, e apenas o necessário para ancorar o cálculo.
- O campo opcional de rastreabilidade não cria coluna de observação no XLS.
- A escolha do índice permanece destacada e separada dos campos de data.
- C0 é explicado como base sem reajuste e não aparece como ciclo selecionável.
- C5 não é criado nem aceito; o limite do modelo permanece C4.
- O processamento continua bloqueado quando faltam competências do índice.
- O download usa o mesmo nome canônico `Coleta_Reajuste.xlsx`.

## Gates de aprovação

- testes automatizados da casca, ciclos e XLS;
- geração real de C2 com C0/C1 preservados e somente C2 computado;
- abertura no Microsoft Excel sem reparo;
- smoke local das quatro rotas;
- smoke público após publicação, incluindo carimbo de atualização, seleção C2 e os três estados do histórico anterior.
