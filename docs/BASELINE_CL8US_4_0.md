# Baseline oficial — Cl8us 4.0

**Marco:** `v4.0-baseline`
**Status:** infraestrutura homologada e fase "Motor da Posição Contratual" encerrada.

## Infraestrutura homologada

XLS Coleta; upload; leitor; processamento progressivo; Motor de Capacidades;
Motor da Posição Contratual; Motor Documental; geração dos oito documentos;
`RESULTADOS`; retrocompatibilidade; fixtures permanentes; invariantes;
auditoria/rastreabilidade; testes Microsoft Excel; homologação cloud.

## Componentes estruturantes protegidos

- upload e leitor;
- processamento progressivo;
- Motor de Capacidades;
- Motor da Posição Contratual;
- `RESULTADOS`;
- geração dos oito documentos;
- retrocompatibilidade.

Qualquer mudança futura nesses componentes deve ser classificada expressamente
como alteração arquitetural, ter sua necessidade justificada, executar a
regressão completa e validar os fixtures permanentes. Alterações silenciosas
nesses componentes não são admitidas.

## Regra de evolução

O Motor da Posição Contratual é uma funcionalidade concluída. Demandas
posteriores devem ser classificadas como **hotfix** ou **nova funcionalidade**.
A arquitetura desta baseline deve permanecer preservada nas próximas frentes de
UX, refinamento documental, inteligência, validações e automações.
