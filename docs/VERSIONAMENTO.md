# Versionamento público do Cl8us — regra pétrea

Fonte única das versões: `_versao.py`. Formato sempre `XX.X` (sem SemVer de três
componentes). A guarda `tools/verificar_versionamento.py` roda no CI rápido e
**falha** quando a regra é violada; ela nunca altera versão sozinha — o
desenvolvedor decide o número, no mesmo PR.

## 1. `CL8US_VERSION`

Toda entrega mergeada em `main` que altere comportamento ou apresentação
percebida pelo usuário incrementa `CL8US_VERSION`.

Exige bump: nova funcionalidade; correção funcional; hotfix; alteração de regra;
alteração de UX ou visual relevante; mudança de fluxo; nova página ou alteração
de página; alteração de cálculo; alteração da Coleta; alteração de documento
gerado, garantia ou DOU; mudança de alerta ou validação percebida pelo usuário.

Normalmente não exige: somente testes; documentação técnica; comentários;
refatoração comprovadamente sem efeito observável; script/ferramenta interna sem
efeito em produção.

## 2. `COLETA_VERSION`

Toda entrega que altere o XLSX entregue ao usuário incrementa `COLETA_VERSION`
(e, por ser user-facing, também `CL8US_VERSION`).

Exige bump: alteração em `templates/COLETA_REAJUSTE_OFICIAL.xlsx`; nova aba ou
remoção; alteração de visibilidade de aba; nova coluna/linha estrutural; mudança
de cabeçalho; novo campo; nova validação; nova fórmula; alteração da memória de
cálculo; mudança estrutural de layout; alteração de campos manuais ou de nomes
definidos; alteração da geração do XLS.

Mudança só no aplicativo, sem efeito no XLS, não exige bump da Coleta.

A nova versão da Coleta entra em `COLETA_VERSOES_ACEITAS` no mesmo PR.
`COLETA_VERSOES_SEM_RESULTADOS_DETALHE` permanece `("11.0", "11.1")`: desde a 11.2
a aba `RESULTADOS_DETALHE` é obrigatória.

## 3. Sequência

- Incremento normal: Cl8us `11.5 -> 11.6 -> 11.7 -> 11.8`; Coleta `11.2 -> 11.3 -> 11.4`.
- Número já consumido por uma entrega nunca é reutilizado. Se o bump foi
  esquecido, a correção respeita a sequência lógica das entregas.
- Caso de origem desta regra: o PR #174 (hotfix temporal, só app) corresponde
  logicamente ao **Cl8us 11.6**; o PR #175 (UX da Coleta) corresponde ao
  **Cl8us 11.7 / Coleta 11.3**. Ambos foram mergeados sem bump; o hotfix de
  versionamento obrigatório levou `main` diretamente a 11.7 / 11.3.

## 4. Bump no mesmo PR

Proibido depender de lembrança do usuário, correção posterior, PR separado ou
ajuste manual depois do merge. Gate antes do merge (também no template de PR):

**VERSÃO CL8US**
- [ ] A entrega altera algo percebido pelo usuário?
- [ ] Se sim, `CL8US_VERSION` foi incrementada?

**VERSÃO COLETA**
- [ ] A Coleta XLSX muda?
- [ ] Se sim, `COLETA_VERSION` foi incrementada?
- [ ] `COLETA_VERSOES_ACEITAS` inclui a nova versão?
- [ ] O marcador gravado no XLS (`CONTROLE!B24`/`B25`) mostra a nova versão?

**FALLBACK**
- [ ] `ATUALIZADO_EM_FALLBACK` foi atualizado?

## 5. Guarda automática (`tools/verificar_versionamento.py`)

No CI (`.github/workflows/ci-pr.yml`, passo "Guarda de versionamento", antes dos
testes sentinela) compara `HEAD^1` (a base) com `HEAD` — no `pull_request`, o
HEAD é o merge do PR sobre `main`.

| Regra | Falha com |
|---|---|
| arquivo user-facing alterado e `CL8US_VERSION` igual à base | "Alteração user-facing detectada sem incremento de CL8US_VERSION." |
| arquivo da Coleta alterado e `COLETA_VERSION` igual à base | "Alteração da Coleta detectada sem incremento de COLETA_VERSION." |
| `COLETA_VERSION` fora de `COLETA_VERSOES_ACEITAS` | "COLETA_VERSION atual não consta em COLETA_VERSOES_ACEITAS." |
| versão alterada que não avança, ou fora de `XX.X` | mensagem específica |
| `CL8US_VERSION` incrementada com `ATUALIZADO_EM_FALLBACK` igual ao da base, anterior a ele ou fora de `dd/mm/aaaa HH:MM` | "CL8US_VERSION incrementada sem atualizar ATUALIZADO_EM_FALLBACK …" |

Superfícies (listas declarativas no próprio script — revisar no PR quando um
módulo novo passar a escrever no XLS):

- **Coleta**: `templates/COLETA_REAJUSTE_OFICIAL.xlsx`, `_coleta_oficial.py`,
  `_gerador_masterfile.py`, `_memoria_calculo.py`, `_ciclo_em_execucao.py`
  (cria a aba `CICLO_EM_EXECUCAO`), `_apresentacao_pc_xls.py` (apresentação de
  `itens_PC`). Avaliados e deixados fora: `_coleta_reajuste.py` (gerador legado
  do `Coleta_Reajuste.xlsx`, não usado pelas páginas) e módulos que só leem ou
  validam o XLS.
- **User-facing**: `app.py`, `pages/**`, todos os módulos `_*.py` da raiz,
  `templates/**`, `assets/**`, `.streamlit/**`, a superfície da Coleta e a lista
  explícita `SEMPRE_PRODUCAO` (avaliada antes das exclusões): `requirements.txt`
  e `tools/atualizar_ist_anatel.py` — importada em runtime por
  `_indice_utils.carregar_ist_anatel` (série IST oficial).
- **Fora de produção** (nunca exigem bump): `tests/**`, `docs/**`, `tools/**`
  (exceto `SEMPRE_PRODUCAO`), `.github/**`, `*.md`, `*.txt` (exceto
  `requirements.txt`), `*.bat`, `teste_*.py` da raiz.
- Auditoria automática: `tests/test_verificar_versionamento.py` falha se
  `app.py`, `pages/**` ou `_*.py` passarem a importar outro módulo de `tools/`
  sem entrada em `SEMPRE_PRODUCAO`.

Única exceção automática: `.py` cuja AST (sem docstrings) é idêntica à da base —
mudança só de comentários/formatação. Não há outra heurística.

Limitações conhecidas (deliberadas, para não criar falso positivo/negativo
obscuro): a guarda não decide se uma mudança é "visualmente relevante" — qualquer
alteração semântica de superfície user-facing exige bump; e não verifica se o
fallback corresponde ao minuto do merge, só que foi atualizado e avançou.
Atualização automática de `icti.csv` pela automação do Ipeadata não aciona a
guarda (dados de índice, fora das superfícies).
