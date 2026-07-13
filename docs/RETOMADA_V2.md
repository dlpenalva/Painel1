# Retomada oficial do CL8US 2.0

Data da decisão: 13/07/2026

## Fonte oficial

- Aplicação: `https://reajustes.streamlit.app/`
- Repositório: `dlpenalva/Painel1`
- Branch estável: `main`
- Entrada: `app.py`
- Linha local de trabalho: `C:\_DesktopReal\05.Via_Git`

O projeto 3.0 fica congelado e fora da cadeia de execução, publicação e
dependências do 2.0. Nenhum componente do 3.0 deve ser copiado para esta linha
sem uma decisão técnica específica, teste isolado e aprovação posterior.

## Arquitetura preservada

O Streamlit é a porta de entrada e controla a sequência da análise:

1. coleta o contexto mínimo do contrato;
2. orienta a escolha e a conferência do índice;
3. apura admissibilidade e ciclos;
4. gera a planilha de coleta;
5. recebe a planilha preenchida e consolida os resultados.

O XLS permanece como artefato operacional principal. Ele deve ser capaz de
calcular, expor fórmulas auditáveis, aceitar informação parcial da fiscalização
e deixar explícito o que não pode ser concluído com segurança.

## Regra de segurança para informação incompleta

A ferramenta deve distinguir três estados, sem inventar dados:

- **calculado**: há dados suficientes, regra identificada e memória de cálculo;
- **estimado ou parcial**: o resultado é limitado ao conjunto informado e traz
  ressalva visível;
- **não calculável**: falta dado essencial; o campo de resultado não é emitido e
  a pendência é listada com responsável e ação necessária.

Nenhum total, retroativo, remanescente ou atualização deve ser apresentado como
definitivo quando depender de quantidade, competência, valor unitário, marco
financeiro ou índice não confirmado.

## Portões para considerar a versão ativa e segura

1. domínio público acessível e associado ao repositório correto;
2. fluxos de ciclo único e múltiplos ciclos abrem sem erro;
3. processamento mínimo produz resultado e disponibiliza o XLS;
4. dependências de produção têm versões fixadas;
5. fontes Python compilam sem erro;
6. teste automatizado gera XLS real e valida estrutura e fórmulas;
7. arquivo gerado abre no Excel sem reparo;
8. exemplos contratuais de referência são recalculados e comparados;
9. limites de informação parcial são explícitos na tela e no XLS;
10. publicação só ocorre após os portões anteriores aplicáveis passarem.

## Backlog crítico

### P0 - estabilidade e prova

- manter o domínio `reajustes.streamlit.app` sob a aplicação `Painel1`;
- criar fumaça automatizada dos dois fluxos e da geração de planilhas;
- validar os XLS no Excel, incluindo fórmulas, vínculos e ausência de reparo;
- registrar uma amostra conhecida para ciclo único e outra para ciclos
  múltiplos;
- eliminar dependência de estado oculto ou valor demonstrativo pré-preenchido.

### P1 - segurança de cálculo e trabalho entre fiscal e GCC

- definir os campos mínimos por modalidade de coleta;
- emitir uma lista objetiva de pendências quando a fiscalização devolver dados
  incompletos;
- calcular somente os recortes sustentados pelos dados disponíveis;
- separar resultado definitivo, parcial e não calculável;
- consolidar retroativo, atualizações e remanescentes sem duplicidade por ciclo;
- gerar comunicação ao fornecedor somente após validação fiscal registrada.

### P2 - usabilidade e manutenção

- reduzir duplicação entre ciclo único e múltiplos ciclos;
- substituir APIs descontinuadas do Streamlit em mudança isolada;
- simplificar textos e telas sem remover a memória de cálculo;
- documentar a responsabilidade de cada campo e a origem de cada resultado.

## Estado desta retomada

O acesso e os fluxos visíveis foram restaurados. Isso restabelece o serviço,
mas não equivale, sozinho, à homologação integral dos cálculos. A declaração de
segurança funcional dependerá dos portões de XLS real, Excel e casos contratuais
de referência.
