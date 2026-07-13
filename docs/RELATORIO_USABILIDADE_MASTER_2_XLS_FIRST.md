# Relatório de usabilidade e segurança — Master 2.0 XLS-first

Data: 13/07/2026
Artefato canônico: `Coleta_Reajuste.xlsx`

## Decisão de produto

O XLS passa a conduzir a apuração. O Streamlit fica responsável por:

1. apurar os marcos e percentuais na calculadora;
2. preencher esses dados no modelo canônico;
3. disponibilizar sempre o arquivo `Coleta_Reajuste.xlsx`;
4. receber o mesmo arquivo após o trabalho do fiscal;
5. validar estrutura, fórmulas e suficiência dos dados;
6. sustentar os downloads documentais nas etapas seguintes.

A web não substitui as fórmulas do Excel e não recalcula os resultados finais.

## Auditoria do modelo recebido

Foram mantidas as regras úteis do arquivo e adotadas as seguintes correções:

- exclusão de `itens_Execucao_Saldo` e `REGRA_NEGOCIO_CLAUS`;
- substituição de `historico` por `RESULTADOS`, entregue vazio nesta etapa;
- limpeza dos dados demonstrativos, preservando a matriz de fórmulas;
- correção de `CONS_QTD_TOTAL`, que somava valores monetários;
- correção de `CONS_VALOR_TOTAL`, que somava quantidades;
- remoção de resíduos de tabela que produziam `#REF!` no Excel;
- proteção das fórmulas de ciclos futuros para devolver vazio, e não `#VALUE!`, quando o ciclo ainda não existe;
- exibição, em `CONTROLE`, da variação e do fator acumulado calculados;
- correção dos rótulos que chamavam um ciclo em análise de “concedido/formalizado”;
- adoção de `mm/aaaa` para datas, exceto `DATA_PC`, mantida em `dd/mm/aaaa`;
- ampliação das colunas de `itens_Consumidos` e `aditivos` conforme o número de linhas de cabeçalho solicitado;
- remoção de comentários e campos de observação.

## Como o fiscal trabalha

As células amarelas são entradas operacionais. As células cinzas são automáticas
ou calculadas. O fiscal não precisa identificar manualmente o ciclo mensal: a
calculadora grava os marcos e o XLS classifica/aplica os fatores.

Fluxo recomendado:

1. conferir `CONTROLE` e `parametros`;
2. preencher os valores pagos por competência em `financeiro`, quando houver base mensal;
3. preencher itens e quantidades em `itens_Remanesc` e `itens_Consumidos`;
4. registrar pedidos de compra em `itens_PC` e alterações em `aditivos`, quando existentes;
5. abrir/recalcular e salvar no Excel;
6. subir o mesmo `Coleta_Reajuste.xlsx` na página Valores.

## Como a GCC recebe informação incompleta

O upload classifica a situação sem inventar números:

| Situação recebida | Resposta da ferramenta |
|---|---|
| Estrutura ou fórmula essencial alterada | Bloqueia e indica a célula/aba afetada |
| Percentual necessário ao acumulado ausente | Bloqueia resultados dependentes e identifica o ciclo |
| Estrutura íntegra, mas sem valores/itens | Aceita o arquivo, informa que está incompleto e não apresenta total |
| Parte dos meses ou itens preenchida | Conta o que foi recebido e marca a base como parcial para a futura consolidação |
| Dados suficientes | Libera a base para a etapa de consolidação, mantendo o Excel como fonte dos resultados |

Informações que podem ser exibidas com segurança mesmo em arquivo parcial:

- índice e ciclo da apuração;
- ciclos marcados para análise;
- quantidade de fórmulas preservadas;
- número de competências, itens, PCs e aditivos preenchidos;
- pendências objetivas de estrutura ou de percentual.

Retroativo, valor atualizado total e remanescente consolidado não são emitidos
como definitivos enquanto dependerem de informação ausente. A construção desses
quadros pertence à segunda etapa.

## Financeiro

A aba possui sempre 60 linhas físicas de fórmula, correspondentes à capacidade
de C0 a C4. O gerador preenche apenas C0 até o último ciclo da apuração; linhas
de ciclos posteriores ficam vazias.

Exemplo validado: uma análise apenas de C2 com início em `09/2024` contém:

- C0: `09/2022` a `08/2023`;
- C1: `09/2023` a `08/2024`;
- C2: `09/2024` a `08/2025`;
- C3 e C4: ausentes da visualização preenchida.

As competências do ciclo em análise aparecem em marinho e negrito. O início do
efeito financeiro continua separado do início teórico do ciclo.

## Identidade visual do Streamlit

- lateral: `#C6D9E8`;
- fundo principal: gradiente `#D4E3EF` → `#BED4E5`;
- texto: marinho `#123B63`;
- ação primária: vinho `#7A1733`;
- box de índice: fundo rosado discreto e borda vinho, mantendo a escolha visivelmente distinta.

## Gate executado

- testes automatizados de C2 isolado e C1–C4: aprovados;
- modelo, simples C2 e múltiplo C4 abertos/recalculados/salvos no Excel real;
- 11.576 fórmulas preservadas em cada arquivo;
- zero `#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?` ou reparo de arquivo;
- upload parcial validado sem emissão de total inseguro.

## Limite desta entrega

Esta é a primeira etapa da remodelagem: geração, uso e retorno seguro do XLS.
A aba `RESULTADOS` e a consolidação web serão construídas na segunda etapa,
lendo os resultados já calculados no Excel e sem criar um motor paralelo.
