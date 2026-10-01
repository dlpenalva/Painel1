"""Decisao canonica: Coleta anterior compatibilizada pode ser formalizada?

Contexto
--------
A Etapa 2 reconhece a linhagem de uma Coleta anterior, recompoe EM MEMORIA os
derivados pela regra vigente e prova, bloco a bloco, que reproduziu o calculo
antigo do Excel. O valor derivado que o XLS legado grava (RESULTADOS) continua
sendo o da regra antiga: por isso a reconciliacao XLS x Python acusa divergencia
e a formalizacao ficava bloqueada.

O que esta decisao faz
----------------------
Distingue divergencia EXPLICADA de divergencia REAL, pela CAUSA e nao pelo
tamanho. Uma divergencia e explicada quando o MESMO motor Python, alimentado
com os derivados LEGADOS (o cache original do Excel, restaurado em memoria),
reproduz o valor do XLS: ou seja, o numero antigo e consequencia exata da
precisao antiga e a recomposicao pela regra vigente e a unica coisa que mudou.
R$ 0,01 sem essa prova continua divergencia relevante; R$ 1,43 com a prova e
divergencia compatibilizada.

A decisao mora no motor/politica. A UI so a consome.

Nada aqui altera entrada do fiscal, o arquivo fisico ou a auditoria XLS x Python:
o valor do XLS e o valor legado reproduzido continuam registrados em cada
divergencia compatibilizada.
"""

from __future__ import annotations

from typing import Any

STATUS_NAO_APLICAVEL = "NAO_APLICAVEL"
STATUS_SEM_ADAPTACAO = "SEM_ADAPTACAO"
STATUS_FORMALIZAVEL = "FORMALIZAVEL_COMPATIBILIZADA"
STATUS_NAO_COMPATIBILIZADA = "NAO_COMPATIBILIZADA"

# Status de CAMPO na reconciliacao XLS x Python.
STATUS_DIVERGENCIA_COMPATIBILIZADA = "DIVERGENCIA_COMPATIBILIZADA"

CAUSA_PRECISAO = "precisao_percentual_anterior"

# Reproducao do calculo ANTIGO: o motor com a precisao legada precisa chegar ao
# valor do XLS, admitindo so o arredondamento estrutural de 1 centavo (a mesma
# tolerancia maxima da reconciliacao). Isto mede a REPRODUCAO, nao a diferenca
# entre o resultado vigente e o XLS: essa so e compatibilizada se a causa for
# comprovada.
TOLERANCIA_REPRODUCAO_LEGADA = 0.01

MENSAGEM_COMPATIBILIZADA = (
    "Coleta de versão anterior compatibilizada com o modelo atual. Os "
    "resultados foram recalculados pelas regras vigentes."
)


def _num(valor: Any) -> float | None:
    try:
        if valor in (None, ""):
            return None
        return float(valor)
    except (TypeError, ValueError):
        return None


def _decisao_base(leitura: dict[str, Any]) -> dict[str, Any]:
    linhagem = (leitura.get("coleta_linhagem") or {}).get("codigo")
    auditoria = leitura.get("compatibilidade_valores") or {}
    return {
        "elegivel": False,
        "status": STATUS_NAO_APLICAVEL,
        "linhagem": linhagem,
        "motivo": "",
        "motivos": [],
        "ciclos_precisao_bruta": list(auditoria.get("ciclos_precisao_bruta") or []),
        "blocos_adaptados": list(auditoria.get("blocos_adaptados") or []),
        "blocos_nao_reproduziveis": dict(auditoria.get("blocos_nao_reproduziveis") or {}),
        "restricoes": [
            r.get("codigo") for r in (leitura.get("compatibilidade_restricoes") or [])
            if isinstance(r, dict)
        ],
        "divergencias_esperadas": [],
        "divergencias_nao_explicadas": [],
        "mensagem": "",
    }


def decidir_formalizacao_compatibilidade(
    leitura: dict[str, Any],
    leitura_legada: dict[str, Any] | None,
) -> dict[str, Any]:
    """Decide se a Coleta anterior adaptada pode usar o resultado canonico atual.

    `leitura` e a leitura ja adaptada; `leitura_legada` e a MESMA leitura com os
    derivados originais do XLS (replay). Fail-closed: qualquer condicao que nao
    possa ser comprovada mantem a decisao como NAO_COMPATIBILIZADA.
    """
    decisao = _decisao_base(leitura)
    if not leitura.get("compatibilidade_aplicada"):
        decisao["motivo"] = "Coleta no modelo atual: nenhuma compatibilidade necessária."
        return decisao

    auditoria = leitura.get("compatibilidade_valores") or {}
    if not auditoria.get("aplicada"):
        decisao["status"] = STATUS_SEM_ADAPTACAO
        decisao["motivo"] = (
            "Coleta anterior sem precisão de reajuste antiga: nada a recompor; "
            "a conferência XLS × Python segue as regras gerais."
        )
        return decisao

    motivos: list[str] = []
    if decisao["blocos_nao_reproduziveis"]:
        motivos.append(
            "Há bloco que não reproduz o cache do Excel deste arquivo: "
            + ", ".join(sorted(decisao["blocos_nao_reproduziveis"]))
            + "."
        )
    for restricao in leitura.get("compatibilidade_restricoes") or []:
        if isinstance(restricao, dict) and restricao.get("mensagem"):
            motivos.append(str(restricao["mensagem"]))
    if not decisao["ciclos_precisao_bruta"]:
        motivos.append("Não há causa de compatibilidade identificada.")

    atual = leitura.get("reconciliacao_xls_python") or {}
    legado = (leitura_legada or {}).get("reconciliacao_xls_python") or {}
    if not leitura_legada or not leitura_legada.get("ok"):
        motivos.append("Não foi possível reproduzir o resultado do XLS legado.")
    elif (
        not atual.get("disponivel")
        or atual.get("sem_cache")
        or atual.get("status_geral") == "RESULTADO_XLS_INDISPONIVEL_POR_CACHE"
    ):
        # Sem a conferencia nao ha prova: um arquivo sem os resultados do XLS
        # recalculados nao pode ser dado como compatibilizado.
        motivos.append("A conferência XLS × Python não está disponível neste arquivo.")
    else:
        campos_legado = {c.get("campo"): c for c in legado.get("campos") or []}
        for div in atual.get("divergencias_relevantes") or []:
            campo = str(div.get("campo") or "")
            antigo = campos_legado.get(campo) or {}
            python_vigente = _num(div.get("python"))
            python_legado = _num(antigo.get("python"))
            xls_valor = _num(div.get("xls"))
            reproduzido = (
                xls_valor is not None
                and python_legado is not None
                and abs(round(xls_valor - python_legado, 4)) <= TOLERANCIA_REPRODUCAO_LEGADA
            )
            mudou = (
                python_vigente is not None
                and python_legado is not None
                and abs(python_vigente - python_legado) > 0.004
            )
            registro = {
                "campo": campo,
                "rotulo": div.get("rotulo"),
                "xls": _num(div.get("xls")),
                "python_vigente": python_vigente,
                "python_legado_reproduzido": python_legado,
                "diferenca": (
                    round(python_vigente - _num(div.get("xls")), 2)
                    if python_vigente is not None and _num(div.get("xls")) is not None
                    else None
                ),
                "causa": CAUSA_PRECISAO,
            }
            if reproduzido and mudou and decisao["ciclos_precisao_bruta"]:
                decisao["divergencias_esperadas"].append(registro)
            else:
                registro["causa"] = None
                decisao["divergencias_nao_explicadas"].append(registro)
        if decisao["divergencias_nao_explicadas"]:
            campos = ", ".join(d["campo"] for d in decisao["divergencias_nao_explicadas"])
            motivos.append(
                "Há divergência XLS × Python que o modelo de compatibilidade não "
                f"explica: {campos}."
            )

    composicao = leitura.get("composicao_vta") or {}
    if composicao.get("bloqueia_formalizacao"):
        motivos.append("A composição do valor contratual não fecha.")

    decisao["motivos"] = motivos
    if motivos:
        decisao["status"] = STATUS_NAO_COMPATIBILIZADA
        decisao["motivo"] = motivos[0]
        return decisao
    decisao["elegivel"] = True
    decisao["status"] = STATUS_FORMALIZAVEL
    decisao["motivo"] = (
        "Precisão anterior integralmente reproduzida e recomposta pela regra "
        "vigente; divergências XLS × Python explicadas pela causa."
    )
    decisao["mensagem"] = MENSAGEM_COMPATIBILIZADA
    return decisao


def aplicar_reclassificacao(reconciliacao: dict[str, Any], decisao: dict[str, Any]) -> None:
    """Reclassifica as divergencias EXPLICADAS na propria reconciliacao.

    So age quando a decisao e elegivel (todas as divergencias foram explicadas).
    O valor do XLS e o valor legado reproduzido permanecem no registro: a
    auditoria nao e escondida, apenas deixa de ser tratada como erro.
    """
    if not decisao.get("elegivel") or not reconciliacao:
        return
    explicadas = {d["campo"]: d for d in decisao.get("divergencias_esperadas") or []}
    compatibilizadas: list[dict[str, Any]] = []
    for campo in reconciliacao.get("campos") or []:
        registro = explicadas.get(campo.get("campo"))
        if not registro or campo.get("status") != "DIVERGENCIA_RELEVANTE":
            continue
        campo["status"] = STATUS_DIVERGENCIA_COMPATIBILIZADA
        campo["python_legado_reproduzido"] = registro["python_legado_reproduzido"]
        campo["causa"] = registro["causa"]
        campo["nota"] = (
            "O XLS legado calculou este valor com a precisão de reajuste "
            "anterior; o mesmo motor reproduz o valor do XLS com os derivados "
            "legados e o recompõe pela regra vigente."
        )
        compatibilizadas.append(campo)
    reconciliacao["divergencias_relevantes"] = [
        c for c in reconciliacao.get("divergencias_relevantes") or []
        if c.get("campo") not in explicadas
    ]
    reconciliacao["divergencias_compatibilizadas"] = compatibilizadas
    if compatibilizadas and not reconciliacao["divergencias_relevantes"]:
        reconciliacao["status_geral"] = STATUS_DIVERGENCIA_COMPATIBILIZADA
