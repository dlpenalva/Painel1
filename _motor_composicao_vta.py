"""Motor de Composicao do VTA — espelha o Quadro de memoria fiscal das apostilas.

Compoe o Valor Total Atualizado por parcelas auditaveis, linha a linha:

    VTA = soma(executado do ciclo x fator acumulado do ciclo)
        + saldo remanescente do corte (base original) x fator acumulado vigente
        + soma(aditivo/supressao: valor na assinatura x fator do ciclo-marco)

vedada a dupla contagem (JA_REFLETIDO_EM). A execucao por ciclo vem da
reconciliacao (fonte prevalente Financeiro > PC > Consumo), entao a parcela
composta ja e a base unica por ciclo, nunca soma de fontes redundantes.

Camada aditiva: nao altera vta_sombra, historico!B51, formulas oficiais nem
parsing existente. O cenario "valor original x fator acumulado" e exposto
apenas como teorico, com alerta de superestimacao — nunca como VTA.
"""
from __future__ import annotations

from typing import Any

ROTULO_CENARIO_TEORICO = (
    "Cenario teorico (valor original x fator acumulado) — superestima o VTA "
    "porque reprecifica execucao ja paga a precos antigos; nao usar como VTA."
)

# Ciclos reais de apuracao. Qualquer outro rotulo que apareca como "ciclo" numa
# fonte (tipicamente a linha TOTAL das abas de entrada) e um AGREGADOR: consolida
# as proprias parcelas e nunca pode ser somado junto delas.
CICLOS_APURACAO = ("C0", "C1", "C2", "C3", "C4")


def _e_ciclo_de_apuracao(ciclo: Any) -> bool:
    """True somente para C0..C4. TOTAL (e afins) e agregador, nao ciclo."""
    return str(ciclo or "").strip().upper() in CICLOS_APURACAO


def _tofl(valor: Any, default: float = 0.0) -> float:
    try:
        if valor in (None, ""):
            return default
        return float(valor)
    except (TypeError, ValueError):
        return default


def _fator_do_ciclo(por_ciclo: dict[str, Any], ciclo: str) -> float | None:
    reg = por_ciclo.get(str(ciclo or "").strip().upper()) or {}
    fator = _tofl(reg.get("fator_acumulado"), default=0.0)
    return fator or None


def _execucao_por_ciclo(
    leitura: dict[str, Any],
    por_ciclo: dict[str, Any],
    alertas: list[str],
) -> list[dict[str, Any]]:
    registros = (leitura.get("reconciliacao") or {}).get("registros") or []
    linhas: list[dict[str, Any]] = []
    agregadores: list[str] = []
    for reg in registros:
        if reg.get("metodo_apuracao") == "VINCULO_EXPLICITO":
            continue
        ciclo = str(reg.get("ciclo") or "").strip().upper()
        base = _tofl(reg.get("valor_computado"))
        if not ciclo or not base:
            continue
        if not _e_ciclo_de_apuracao(ciclo):
            # Trava anti-dupla-contagem: registros agregadores (a linha TOTAL das
            # abas de entrada, por exemplo) consolidam as parcelas dos proprios
            # ciclos. Soma-los junto de C0..C4 contaria a mesma execucao duas
            # vezes. Ficam de fora da composicao, sempre.
            agregadores.append(ciclo)
            continue
        fator = _fator_do_ciclo(por_ciclo, ciclo)
        if fator is None:
            alertas.append(
                f"Composicao VTA: ciclo {ciclo} sem fator acumulado "
                "parametrizado; execucao composta pela base, sem atualizacao."
            )
            atualizado = round(base, 2)
        else:
            atualizado = round(base * fator, 2)
        linhas.append({
            "ciclo": ciclo,
            "descricao": (
                f"{ciclo} executado" if (fator or 1.0) == 1.0
                else f"{ciclo} executado atualizado"
            ),
            "valor_base": round(base, 2),
            "fator_acumulado": fator,
            "valor_atualizado": atualizado,
            "fonte": reg.get("fonte_principal"),
            "status_reconciliacao": reg.get("status_reconciliacao"),
            "bloqueia_formalizacao": bool(reg.get("bloqueia_formalizacao")),
        })
    if agregadores:
        alertas.append(
            "Composicao VTA: registro(s) agregador(es) "
            + ", ".join(sorted(set(agregadores)))
            + " ficaram fora da soma das parcelas (ja consolidam os ciclos "
            "C0..C4; soma-los duplicaria a execucao)."
        )
    return sorted(linhas, key=lambda l: l["ciclo"])


def _saldo_remanescente(
    leitura: dict[str, Any],
    alertas: list[str],
) -> dict[str, Any] | None:
    potencial = leitura.get("potencial_futuro") or {}
    saldo = _tofl(potencial.get("saldo_remanescente_base"))
    if not saldo:
        return None
    fator = _tofl(potencial.get("fator_vigente"), default=0.0) or None
    atualizado = _tofl(potencial.get("valor_atualizado_vigente"), default=0.0)
    if not atualizado:
        atualizado = round(saldo, 2)
        alertas.append(
            "Composicao VTA: saldo remanescente sem fator vigente; "
            "composto pela base, sem atualizacao."
        )
    return {
        "descricao": "Saldo remanescente atualizado no corte",
        "valor_base": round(saldo, 2),
        "fator_acumulado": fator,
        "valor_atualizado": round(atualizado, 2),
        "ciclo": potencial.get("ciclo_vigente") or "",
        "fonte": "remanescente",
    }


def _aditivos(
    leitura: dict[str, Any],
    alertas: list[str],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    visiveis = leitura.get("aditivos_visiveis") or {}
    computados: list[dict[str, Any]] = []
    nao_computados: list[dict[str, Any]] = []

    if visiveis.get("ok"):
        for item in visiveis.get("itens") or []:
            registro = {
                "descricao": item.get("evento") or "Aditivo/supressao",
                "ciclo": item.get("ciclo_marco") or "",
                "valor_base": _tofl(item.get("valor_assinatura")),
                "fator_acumulado": item.get("fator_acumulado"),
                "valor_atualizado": _tofl(
                    item.get("valor_atualizado"),
                    default=_tofl(item.get("valor_assinatura")),
                ),
                "fonte": "aditivo",
                "ja_refletido_em": item.get("ja_refletido_em"),
            }
            if str(item.get("ja_refletido_em") or "Nao").strip() not in ("", "Nao"):
                registro["motivo"] = (
                    f"Ja refletido em {item.get('ja_refletido_em')}; "
                    "nao soma ao VTA (vedada dupla contagem)."
                )
                nao_computados.append(registro)
            else:
                computados.append(registro)
        return computados, nao_computados

    # Fallback legado: aditivos consolidados da aba oculta (motor sombra),
    # que ja chegam a valor atualizado.
    parcelas = (leitura.get("vta_sombra") or {}).get("parcelas_computadas") or []
    posicao = leitura.get("posicao_contratual") or {}
    if posicao.get("ok") and not posicao.get("cache_ausente"):
        # No modelo oficial, os deltas da aba aditivos ja integram a posicao
        # contratual e, por ela, o remanescente valorado. Soma-los novamente
        # como parcela financeira duplicaria o mesmo efeito no VTA.
        for parcela in parcelas:
            if parcela.get("fonte_parcela") != "Aditivo":
                continue
            nao_computados.append({
                "descricao": parcela.get("identificador") or "Aditivo",
                "ciclo": parcela.get("ciclo") or "",
                "valor_base": parcela.get("valor_original"),
                "fator_acumulado": parcela.get("fator_acumulado"),
                "valor_atualizado": _tofl(parcela.get("valor")),
                "fonte": "aditivo",
                "ja_refletido_em": "posicao_contratual/remanescente",
                "motivo": (
                    "Efeito fisico ja refletido na posicao contratual e no "
                    "remanescente; nao soma novamente ao VTA."
                ),
            })
        return [], nao_computados

    for parcela in parcelas:
        if parcela.get("fonte_parcela") != "Aditivo":
            continue
        computados.append({
            "descricao": parcela.get("identificador") or "Aditivo (aba oculta)",
            "ciclo": "",
            "valor_base": None,
            "fator_acumulado": None,
            "valor_atualizado": _tofl(parcela.get("valor")),
            "fonte": "aditivo",
            "ja_refletido_em": "Nao",
        })
    if computados:
        alertas.append(
            "Composicao VTA: aditivos lidos do bloco oculto homologado "
            "(sem aba visivel ENTRADA_XLS_ADITIVOS); valores ja atualizados."
        )
    return computados, nao_computados


def _ciclo_num(ciclo: Any) -> int | None:
    s = str(ciclo or "").strip().upper()
    if s.startswith("C") and s[1:].isdigit():
        return int(s[1:])
    return None


def _valor_considerado(item: dict[str, Any]) -> float:
    """Valor historico EFETIVAMENTE considerado de um PC.

    Base paga somada ao retroativo reconhecido a pagar. PC sem efeito financeiro
    e PC do intervalo precluso ficam no valor original — o fator matematico
    integral existe como dado tecnico em VALOR_ATUALIZADO, mas nunca alimenta o
    valor efetivo, o retroativo nem o VTA. Arquivos anteriores a essa medida
    caem no comportamento anterior (VALOR_ATUALIZADO).
    """
    considerado = item.get("valor_historico_considerado")
    if considerado not in (None, ""):
        return _tofl(considerado)
    return _tofl(item.get("valor_atualizado"))


def _abertura(posicao: dict[str, Any], ciclo: int) -> Any:
    """Fotografia da abertura do ciclo, sem aditivo de efeito posterior.

    Cai em ``QTD_REM_AJUSTADA_Cn`` quando a camada temporal nao existe no
    arquivo (modelos anteriores), preservando o resultado anterior.
    """
    valor = posicao.get(f"QTD_REM_ABERTURA_C{ciclo}")
    if valor in (None, ""):
        valor = posicao.get(f"QTD_REM_AJUSTADA_C{ciclo}")
    return valor


def _entra_no_calculo(item: dict[str, Any]) -> bool:
    return str(item.get("entra_no_calculo") or "Sim").strip().lower() in (
        "sim", "s", "true", "1"
    )


def _composicao_vta_pc(
    leitura: dict[str, Any],
    por_ciclo: dict[str, Any],
    alertas: list[str],
) -> dict[str, Any]:
    """Composicao automatica do VTA pelo metodo Pedidos de Compra (PC).

    VTA_PC = execucao historica atualizada pelo metodo PC
             + remanescente fisico atualizado no MESMO corte de referencia.

    - C0 executado: quando nao ha PC confiavel para C0, deriva da movimentacao
      fisica (QTD_REM_AJUSTADA_C0 - QTD_REM_AJUSTADA_C1) valorada ao VU_C0.
    - Cn (1<=n<vigente): soma de itens_PC.VALOR_ATUALIZADO do ciclo (regra
      itens_PC ja homologada; nao reaplica reajuste a ciclos preclusos).
    - Remanescente no ciclo vigente: item a item, VU atualizado do ciclo
      vigente arredondado a centavos, multiplicado pela QTD_REM_AJUSTADA do
      ciclo vigente, total do item arredondado, entao somado.

    Regra petrea (mesmo corte temporal): execucao e remanescente pertencem ao
    mesmo corte. PCs no ciclo vigente ou posteriores NAO entram na execucao
    (seriam dupla contagem contra o remanescente integral daquele corte).
    Base insuficiente -> CALCULO MANUAL REQUERIDO (nunca fabrica resultado).
    """
    controle = leitura.get("controle") or {}
    vigente = str(controle.get("ciclo_vigente") or "").strip().upper()
    n_vig = _ciclo_num(vigente)
    posc = ((leitura.get("posicao_contratual") or {}).get("itens")) or []
    hist = ((leitura.get("historico_vu") or {}).get("itens")) or []
    itens_pc = ((leitura.get("itens_pc_v10") or {}).get("itens")) or []

    if not posc or not hist or n_vig is None:
        return {
            "disponivel": False,
            "motivo": (
                "CALCULO MANUAL REQUERIDO: base itemizada (posicao_contratual / "
                "historico_VU) ou ciclo vigente ausente para compor o VTA pelo "
                "metodo PC."
            ),
        }

    vu_por_item = {h.get("item"): (h.get("vu_ciclos") or {}) for h in hist}

    # Execucao PC por ciclo. A medida e o VALOR_HISTORICO_CONSIDERADO (base paga
    # + retroativo reconhecido): PC sem efeito financeiro entra pelo valor
    # original, nunca pelo fator integral.
    pc_atual: dict[int, float] = {}
    pc_base: dict[int, float] = {}
    posteriores = 0
    fora_do_corte = 0
    for it in itens_pc:
        if not _entra_no_calculo(it):
            continue
        if not it.get("dentro_do_corte", True):
            # PC posterior a data de corte unica: permanece no inventario do
            # arquivo, mas nao entra em nenhum resultado ate o corte.
            fora_do_corte += 1
            continue
        considerado = _valor_considerado(it)
        # Nao ha categoria de execucao "entre ciclos": ciclos consecutivos sao
        # contiguos e as competencias anteriores ao inicio do efeito financeiro
        # pertencem ao proprio ciclo (efeito Nao, fator efetivo 1).
        n = _ciclo_num(it.get("ciclo"))
        if n is None:
            continue
        if n >= n_vig:
            # PC no corte vigente ou posterior: cobertura posterior e projecao,
            # nao desloca silenciosamente o corte oficial do VTA.
            posteriores += 1
            continue
        pc_atual[n] = pc_atual.get(n, 0.0) + considerado
        pc_base[n] = pc_base.get(n, 0.0) + _tofl(it.get("valor_pc"))
    if posteriores:
        alertas.append(
            f"Composicao VTA-PC: {posteriores} PC(s) no ciclo vigente ({vigente}) "
            "ou posteriores nao entram na execucao (mesmo corte temporal; evita "
            "dupla contagem contra o remanescente). Seguem como diagnostico/projecao."
        )
    if fora_do_corte:
        alertas.append(
            f"Composicao VTA-PC: {fora_do_corte} PC(s) com DATA_PC posterior a "
            "data de corte do contrato permanecem no inventario do arquivo e "
            "nao compoem nenhum resultado ate o corte."
        )

    execucao: list[dict[str, Any]] = []

    # C0 executado: PC confiavel se houver; senao movimentacao fisica ao VU_C0.
    if 0 in pc_atual:
        execucao.append({
            "ciclo": "C0",
            "descricao": "C0 executado (PC)",
            "valor_base": round(pc_base.get(0, 0.0), 2),
            "fator_acumulado": _fator_do_ciclo(por_ciclo, "C0"),
            "valor_atualizado": round(pc_atual[0], 2),
            "fonte": "pc",
        })
    else:
        c0_fisico = 0.0
        for p in posc:
            item = p.get("ITEM")
            vu_c0 = _tofl((vu_por_item.get(item) or {}).get("VU_C0"))
            # Movimentacao entre as ABERTURAS (fotografia temporalmente correta):
            # aditivo com efeito no meio de C1 nao pode encolher a execucao de C0.
            mov = _tofl(_abertura(p, 0)) - _tofl(_abertura(p, 1))
            c0_fisico += mov * vu_c0
        c0_fisico = round(c0_fisico, 2)
        if c0_fisico:
            execucao.append({
                "ciclo": "C0",
                "descricao": "C0 executado (movimentacao fisica x VU_C0)",
                "valor_base": c0_fisico,
                "fator_acumulado": 1.0,
                "valor_atualizado": c0_fisico,
                "fonte": "fisico",
            })

    # Cn executado por PC (1 <= n < vigente).
    for n in sorted(k for k in pc_atual if 1 <= k < n_vig):
        execucao.append({
            "ciclo": f"C{n}",
            "descricao": f"C{n} executado atualizado (PC)",
            "valor_base": round(pc_base.get(n, 0.0), 2),
            "fator_acumulado": _fator_do_ciclo(por_ciclo, f"C{n}"),
            "valor_atualizado": round(pc_atual[n], 2),
            "fonte": "pc",
        })

    # Ciclo vigente: PRESENTE (execucao fisica ja informada pelo fiscal) e
    # FUTURO (remanescente fisico atual x VU atualizado). A posicao fisica
    # completa PREVALECE sobre qualquer estimativa por PCs; sem ela, mantem-se
    # a fotografia da abertura + alteracoes posteriores (comportamento anterior).
    vu_key = f"VU_C{n_vig}"
    fator_vig = _fator_do_ciclo(por_ciclo, vigente)
    fisico = leitura.get("ciclo_em_execucao") or {}
    fisico_ok = bool(
        fisico.get("disponivel") and fisico.get("completo") and fisico.get("valido")
    )

    rem = 0.0
    rem_base = 0.0
    posterior_valor = 0.0
    for p in posc:
        item = p.get("ITEM")
        vu = round(_tofl((vu_por_item.get(item) or {}).get(vu_key)), 2)
        vu_orig = round(_tofl(p.get("VU_ORIGINAL")), 2)
        qtd = _tofl(_abertura(p, n_vig))
        rem += round(vu * qtd, 2)
        rem_base += round(vu_orig * qtd, 2)
        posterior_valor += round(
            vu * _tofl(p.get(f"ALTERACAO_POSTERIOR_ABERTURA_C{n_vig}")), 2
        )
    rem = round(rem, 2)
    rem_base = round(rem_base, 2)
    posterior_valor = round(posterior_valor, 2)

    presente = None
    saldo = None
    if fisico_ok:
        consumido = round(_tofl(fisico.get("total_valor_consumido")), 2)
        remanescente = round(_tofl(fisico.get("total_valor_remanescente")), 2)
        data_pos = fisico.get("data_posicao")
        if consumido:
            presente = {
                "descricao": (
                    f"{vigente} executado ate o levantamento do quantitativo restante"
                    + (f" em {data_pos:%d/%m/%Y}" if hasattr(data_pos, "year") else "")
                ),
                "valor_base": consumido,
                "fator_acumulado": fator_vig,
                "valor_atualizado": consumido,
                "ciclo": vigente,
                "fonte": "posicao_fisica",
            }
        if remanescente:
            saldo = {
                "descricao": f"{vigente} quantitativo restante atual (por item)",
                "valor_base": remanescente,
                "fator_acumulado": fator_vig,
                "valor_atualizado": remanescente,
                "ciclo": vigente,
                "fonte": "posicao_fisica",
            }
        alertas.append(
            "Composicao VTA-PC: posicao fisica itemizada do ciclo vigente "
            "completa e valida; ela PREVALECE sobre a estimativa por PCs "
            "(nenhuma projecao por dias, media de consumo ou subtracao de PCs)."
        )
    elif rem:
        if posterior_valor:
            # Alteracoes com efeito posterior a abertura entram por componente
            # proprio, uma unica vez (trava anti-dupla-contagem).
            execucao.append({
                "ciclo": vigente,
                "descricao": (
                    f"{vigente} alteracoes contratuais posteriores a abertura"
                ),
                "valor_base": posterior_valor,
                "fator_acumulado": fator_vig,
                "valor_atualizado": posterior_valor,
                "fonte": "aditivo_posterior_abertura",
            })
        saldo = {
            "descricao": f"{vigente} remanescente na abertura (por item)",
            "valor_base": rem_base,
            "fator_acumulado": fator_vig,
            "valor_atualizado": rem,
            "ciclo": vigente,
            "fonte": "remanescente",
        }
    if presente is not None:
        execucao.append(presente)

    if not execucao and saldo is None:
        return {
            "disponivel": False,
            "motivo": (
                "CALCULO MANUAL REQUERIDO: sem execucao PC nem remanescente "
                "itemizado computavel no corte vigente."
            ),
        }

    total_exec = round(sum(l["valor_atualizado"] for l in execucao), 2)
    total_base = round(sum(l["valor_base"] for l in execucao), 2)
    valor_saldo = saldo["valor_atualizado"] if saldo else 0.0
    return {
        "disponivel": True,
        "motivo": "",
        "metodo": "pc",
        "execucao_por_ciclo": execucao,
        "saldo_remanescente": saldo,
        "aditivos": [],
        "aditivos_nao_computados": [],
        "total_execucao_atualizada": total_exec,
        "total_execucao_base": total_base,
        "retroativo_implicito": round(total_exec - total_base, 2),
        "total_aditivos_atualizados": 0.0,
        "vta_composicao": round(total_exec + valor_saldo, 2),
    }


def _remanescente_atualizado_vigente(
    leitura: dict[str, Any],
    por_ciclo: dict[str, Any],
    vigente: str,
) -> tuple[float | None, float | None, str]:
    """Saldo futuro atualizado do ciclo vigente pela fonte canonica do metodo.

    Mesma cadeia do potencial restante do objeto do processo: primeiro o
    ``potencial_futuro`` consolidado; na falta dele, a fotografia itemizada de
    ``posicao_contratual`` valorada pelo VU do ciclo vigente em ``historico_VU``
    (espelho de ``MEMORIA_RESULTADOS!D35``). Nunca fabrica valor.
    """
    potencial = leitura.get("potencial_futuro") or {}
    if potencial.get("disponivel"):
        atualizado = _tofl(potencial.get("valor_atualizado_vigente"), default=0.0)
        if atualizado:
            return (
                round(_tofl(potencial.get("saldo_remanescente_base")), 2),
                round(atualizado, 2),
                "potencial_futuro",
            )

    posicao = leitura.get("posicao_contratual") or {}
    if not (posicao.get("ok") and not posicao.get("cache_ausente")):
        return None, None, ""
    if not _e_ciclo_de_apuracao(vigente):
        return None, None, ""
    vu_por_item = {
        str(h.get("item") or "").strip().upper(): (h.get("vu_ciclos") or {})
        for h in ((leitura.get("historico_vu") or {}).get("itens") or [])
    }
    base = 0.0
    atualizado = 0.0
    itens = 0
    for reg in posicao.get("itens") or []:
        chave = str(reg.get("ITEM") or "").strip().upper()
        qtd = reg.get(f"QTD_REM_AJUSTADA_{vigente}")
        if qtd in (None, ""):
            qtd = _abertura(reg, int(vigente[1:]))
        vu_original = reg.get("VU_ORIGINAL")
        vu_vigente = (vu_por_item.get(chave) or {}).get(f"VU_{vigente}")
        if qtd in (None, "") or vu_original in (None, "") or vu_vigente in (None, ""):
            continue
        base += round(_tofl(qtd) * _tofl(vu_original), 2)
        atualizado += round(_tofl(qtd) * _tofl(vu_vigente), 2)
        itens += 1
    if not itens:
        return None, None, ""
    return round(base, 2), round(atualizado, 2), "posicao_contratual"


def _parcelas_financeiro_por_ciclo(
    leitura: dict[str, Any],
) -> dict[str, dict[str, float]]:
    """Agrega, por ciclo, o pago informado e o valor considerado do Financeiro.

    Fonte canonica unica: as parcelas do financeiro em ``vta_sombra``. Parcelas
    ``:delta:`` sao ignoradas — o retroativo da competencia ja esta dentro do
    ``valor_atualizado`` da parcela-base correspondente; soma-las dobraria o
    ajuste. Rotulos que nao sao ciclo de apuracao (a linha TOTAL da aba) sao
    agregadores e ficam de fora.
    """
    agregado: dict[str, dict[str, float]] = {}
    for parcela in (leitura.get("vta_sombra") or {}).get("parcelas_computadas") or []:
        if parcela.get("fonte_parcela") not in {"Financeiro", "Historico financeiro"}:
            continue
        ident = str(parcela.get("identificador") or "")
        if ident.startswith("financeiro:") and ":delta:" in ident:
            continue
        ciclo = str(parcela.get("ciclo") or "").strip().upper()
        if not _e_ciclo_de_apuracao(ciclo):
            continue
        base = _tofl(parcela.get("valor"))
        atualizado = parcela.get("valor_atualizado")
        considerado = base if atualizado in (None, "") else _tofl(atualizado)
        alvo = agregado.setdefault(ciclo, {"pago": 0.0, "considerado": 0.0})
        alvo["pago"] = round(alvo["pago"] + base, 2)
        alvo["considerado"] = round(alvo["considerado"] + considerado, 2)
    return agregado


def _composicao_vta_financeiro(
    leitura: dict[str, Any],
    por_ciclo: dict[str, Any],
    alertas: list[str],
) -> dict[str, Any]:
    """Composicao economica do VTA pelo metodo Financeiro.

    VTA = executado apurado (pago informado na aba financeiro)
        + ajustes ainda devidos (reajuste reconhecido, ainda nao contido no pago)
        + remanescente atualizado (saldo futuro pela fonte canonica do metodo)

    Espelha, por calculo proprio, a identidade de ``MEMORIA_RESULTADOS!B26``
    (``D20 + B21 + D35``) — sem ler B26 e sem copiar o XLS. O detalhamento por
    ciclo segue como camada de auditoria e NAO produz um segundo total.
    """
    controle = leitura.get("controle") or {}
    vigente = str(controle.get("ciclo_vigente") or "").strip().upper()
    agregado = _parcelas_financeiro_por_ciclo(leitura)
    if not agregado:
        return {
            "disponivel": False,
            "motivo": (
                "CALCULO MANUAL REQUERIDO: sem parcela financeira apurada por "
                "ciclo para compor o VTA pelo metodo Financeiro."
            ),
        }

    detalhamento: list[dict[str, Any]] = []
    for ciclo in CICLOS_APURACAO:
        reg = agregado.get(ciclo)
        if not reg:
            continue
        detalhamento.append({
            "ciclo": ciclo,
            "descricao": f"{ciclo} — pago informado e valor considerado",
            "valor_pago": reg["pago"],
            "valor_considerado": reg["considerado"],
            "ajuste_devido": round(reg["considerado"] - reg["pago"], 2),
            "fator_acumulado": _fator_do_ciclo(por_ciclo, ciclo),
            "fonte": "financeiro",
        })

    executado = round(sum(r["pago"] for r in agregado.values()), 2)
    considerado = round(sum(r["considerado"] for r in agregado.values()), 2)
    ajustes = round(considerado - executado, 2)
    rem_base, rem_atualizado, origem_rem = _remanescente_atualizado_vigente(
        leitura, por_ciclo, vigente
    )
    if rem_atualizado is None:
        return {
            "disponivel": False,
            "motivo": (
                "CALCULO MANUAL REQUERIDO: remanescente atualizado do ciclo "
                "vigente indisponivel para compor o VTA pelo metodo Financeiro."
            ),
        }

    execucao = [{
        "ciclo": "",
        "descricao": "Executado apurado",
        "valor_base": executado,
        "fator_acumulado": None,
        "valor_atualizado": executado,
        "fonte": "financeiro",
        "explicacao": (
            "Valores efetivamente pagos, conforme informados na aba Financeiro."
        ),
    }]
    ajustes_linha = {
        "ciclo": "",
        "descricao": "Ajustes ainda devidos",
        "valor_base": ajustes,
        "fator_acumulado": None,
        "valor_atualizado": ajustes,
        "fonte": "retroativo",
        "explicacao": (
            "Reajuste ja reconhecido e ainda nao contido no valor pago."
            if ajustes else "Nao ha reajuste reconhecido pendente de pagamento."
        ),
    }
    saldo = {
        "ciclo": vigente,
        "descricao": "Remanescente atualizado",
        "valor_base": rem_base,
        "fator_acumulado": _fator_do_ciclo(por_ciclo, vigente),
        "valor_atualizado": rem_atualizado,
        "fonte": origem_rem or "remanescente",
        "explicacao": "Saldo que ainda falta executar, ja atualizado.",
    }
    return {
        "disponivel": True,
        "motivo": "",
        "metodo": "financeiro",
        "execucao_por_ciclo": execucao,
        "ajustes_devidos": ajustes_linha,
        "saldo_remanescente": saldo,
        "detalhamento_por_ciclo": detalhamento,
        "total_execucao_atualizada": executado,
        "total_execucao_base": executado,
        "retroativo_implicito": ajustes,
        "total_aditivos_atualizados": 0.0,
        "vta_composicao": round(executado + ajustes + rem_atualizado, 2),
    }


def _cenario_teorico(
    leitura: dict[str, Any],
    por_ciclo: dict[str, Any],
) -> dict[str, Any] | None:
    itens = (leitura.get("itens_contrato") or {}).get("itens") or []
    valor_original = round(sum(
        _tofl(i.get("qtd_contratada")) * _tofl(i.get("vu_original"))
        for i in itens
    ), 2)
    if not valor_original:
        return None
    ciclo_vigente = str(
        (leitura.get("controle") or {}).get("ciclo_vigente") or ""
    ).strip().upper()
    fator = _fator_do_ciclo(por_ciclo, ciclo_vigente)
    if fator is None:
        return None
    return {
        "rotulo": ROTULO_CENARIO_TEORICO,
        "valor_original": valor_original,
        "ciclo_vigente": ciclo_vigente,
        "fator_acumulado": fator,
        "valor_teorico": round(valor_original * fator, 2),
        "e_vta": False,
    }


def montar_composicao_vta(leitura: dict[str, Any]) -> dict[str, Any]:
    """Monta o quadro COMPOSICAO_VTA a partir da leitura ja consolidada.

    Requer que a leitura ja tenha reconciliacao, potencial_futuro e (quando
    existir) aditivos_visiveis — por isso roda ao final do leitor.
    """
    alertas: list[str] = []
    por_ciclo = (leitura.get("parametros_v10") or {}).get("por_ciclo") or {}

    # Metodo Pedidos de Compra (PC): composicao automatica dedicada, itemizada,
    # com mesmo corte temporal. Nao altera as demais metodologias (Financeiro /
    # Itens Consumidos), que seguem pela composicao por reconciliacao.
    modo = str((leitura.get("controle") or {}).get("modo") or "").strip().lower()
    if modo == "pc":
        pc = _composicao_vta_pc(leitura, por_ciclo, alertas)
        if pc.get("disponivel"):
            execucao = pc["execucao_por_ciclo"]
            saldo = pc["saldo_remanescente"]
            linhas = [dict(l, tipo="execucao") for l in execucao]
            if saldo:
                linhas.append(dict(saldo, tipo="saldo_remanescente"))
            for idx, linha in enumerate(linhas):
                linha["ref"] = chr(ord("A") + idx) if idx < 26 else str(idx + 1)
            return {
                "disponivel": True,
                "motivo": "",
                "metodo": "pc",
                "execucao_por_ciclo": execucao,
                "saldo_remanescente": saldo,
                "aditivos": [],
                "aditivos_nao_computados": [],
                "total_execucao_atualizada": pc["total_execucao_atualizada"],
                "total_execucao_base": pc["total_execucao_base"],
                "retroativo_implicito": pc["retroativo_implicito"],
                "total_aditivos_atualizados": 0.0,
                "vta_composicao": pc["vta_composicao"],
                "cenario_teorico": _cenario_teorico(leitura, por_ciclo),
                "bloqueia_formalizacao": False,
                "linhas": linhas,
                "alertas": alertas,
            }
        # Base insuficiente: nunca fabrica resultado — devolve motivo controlado.
        return {
            "disponivel": False,
            "motivo": pc.get("motivo") or "CALCULO MANUAL REQUERIDO (metodo PC).",
            "metodo": "pc",
            "execucao_por_ciclo": [],
            "saldo_remanescente": None,
            "aditivos": [],
            "aditivos_nao_computados": [],
            "total_execucao_atualizada": 0.0,
            "total_execucao_base": 0.0,
            "retroativo_implicito": 0.0,
            "total_aditivos_atualizados": 0.0,
            "vta_composicao": None,
            "cenario_teorico": _cenario_teorico(leitura, por_ciclo),
            "bloqueia_formalizacao": False,
            "linhas": [],
            "alertas": alertas,
        }

    # Metodo Financeiro: composicao economica dedicada, em tres parcelas que
    # reconciliam exatamente com o VTA oficial (executado apurado + ajustes
    # ainda devidos + remanescente atualizado). O detalhamento por ciclo segue
    # como camada de auditoria, sem gerar um segundo total concorrente.
    # PC e Consumido nao passam por aqui.
    if modo == "principal":
        fin = _composicao_vta_financeiro(leitura, por_ciclo, alertas)
        base_financeiro: dict[str, Any] = {
            "disponivel": False,
            "motivo": fin.get("motivo") or "CALCULO MANUAL REQUERIDO (Financeiro).",
            "metodo": "financeiro",
            "execucao_por_ciclo": [],
            "ajustes_devidos": None,
            "saldo_remanescente": None,
            "detalhamento_por_ciclo": [],
            "aditivos": [],
            "aditivos_nao_computados": [],
            "total_execucao_atualizada": 0.0,
            "total_execucao_base": 0.0,
            "retroativo_implicito": 0.0,
            "total_aditivos_atualizados": 0.0,
            "vta_composicao": None,
            "cenario_teorico": _cenario_teorico(leitura, por_ciclo),
            "bloqueia_formalizacao": False,
            "linhas": [],
            "alertas": alertas,
        }
        if not fin.get("disponivel"):
            # Base insuficiente: nunca fabrica resultado nem cai no caminho
            # generico (que somaria parcelas de semantica diferente).
            return base_financeiro
        linhas = [
            dict(fin["execucao_por_ciclo"][0], tipo="execucao"),
            dict(fin["ajustes_devidos"], tipo="ajustes_devidos"),
            dict(fin["saldo_remanescente"], tipo="saldo_remanescente"),
        ]
        for idx, linha in enumerate(linhas):
            linha["ref"] = chr(ord("A") + idx)
        base_financeiro.update({
            "disponivel": True,
            "motivo": "",
            "execucao_por_ciclo": fin["execucao_por_ciclo"],
            "ajustes_devidos": fin["ajustes_devidos"],
            "saldo_remanescente": fin["saldo_remanescente"],
            "detalhamento_por_ciclo": fin["detalhamento_por_ciclo"],
            "total_execucao_atualizada": fin["total_execucao_atualizada"],
            "total_execucao_base": fin["total_execucao_base"],
            "retroativo_implicito": fin["retroativo_implicito"],
            "vta_composicao": fin["vta_composicao"],
            "linhas": linhas,
        })
        return base_financeiro

    execucao = _execucao_por_ciclo(leitura, por_ciclo, alertas)
    saldo = _saldo_remanescente(leitura, alertas)
    aditivos, aditivos_fora = _aditivos(leitura, alertas)

    saida: dict[str, Any] = {
        "disponivel": False,
        "motivo": "",
        "execucao_por_ciclo": execucao,
        "saldo_remanescente": saldo,
        "aditivos": aditivos,
        "aditivos_nao_computados": aditivos_fora,
        "total_execucao_atualizada": 0.0,
        "total_execucao_base": 0.0,
        "retroativo_implicito": 0.0,
        "total_aditivos_atualizados": 0.0,
        "vta_composicao": None,
        "cenario_teorico": _cenario_teorico(leitura, por_ciclo),
        "bloqueia_formalizacao": any(
            l.get("bloqueia_formalizacao") for l in execucao
        ),
        "linhas": [],
        "alertas": alertas,
    }

    if not execucao and saldo is None and not aditivos:
        saida["motivo"] = (
            "Sem execucao reconciliada, saldo remanescente ou aditivos; "
            "nada a compor."
        )
        return saida

    total_exec = round(sum(l["valor_atualizado"] for l in execucao), 2)
    total_base = round(sum(l["valor_base"] for l in execucao), 2)
    total_aditivos = round(sum(a["valor_atualizado"] for a in aditivos), 2)
    valor_saldo = round(saldo["valor_atualizado"], 2) if saldo else 0.0

    saida.update({
        "disponivel": True,
        "total_execucao_atualizada": total_exec,
        "total_execucao_base": total_base,
        "retroativo_implicito": round(total_exec - total_base, 2),
        "total_aditivos_atualizados": total_aditivos,
        "vta_composicao": round(total_exec + valor_saldo + total_aditivos, 2),
    })

    linhas = [dict(l, tipo="execucao") for l in execucao]
    if saldo:
        linhas.append(dict(saldo, tipo="saldo_remanescente"))
    linhas.extend(dict(a, tipo="aditivo") for a in aditivos)
    for idx, linha in enumerate(linhas):
        linha["ref"] = chr(ord("A") + idx) if idx < 26 else str(idx + 1)
    saida["linhas"] = linhas

    if saida["bloqueia_formalizacao"]:
        alertas.append(
            "Composicao VTA: ha ciclo DIVERGENTE na reconciliacao; o VTA "
            "composto usa o valor prevalente e a formalizacao segue bloqueada "
            "ate decisao da GCC."
        )
    return saida
