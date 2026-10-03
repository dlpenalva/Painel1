"""Deteccao estrutural das linhagens de Coleta homologadas pelo Cl8us 11.0.

Arquivos anteriores ao versionamento publico nao recebem numeros inventados.
Eles sao reconhecidos por uma assinatura estrutural deterministica, formada
por abas e cabecalhos que nao dependem da data, do nome ou do conteudo
preenchido pelo fiscal.
"""

from __future__ import annotations

from hashlib import sha256
import json
import unicodedata
from typing import Any

from _versao import (
    COLETA_COMPATIBILIDADE_ANTERIORES,
    COLETA_VERSION,
    COLETA_VERSOES_ACEITAS,
)


LINHAGEM_COLETA_11 = "COLETA_11"
LINHAGEM_PRE_11_L1 = "PRE_11_L1"
LINHAGEM_PRE_11_L2 = "PRE_11_L2"
LINHAGEM_NAO_HOMOLOGADA = "NAO_HOMOLOGADA"

_ABAS_BASE = {
    "CONTROLE",
    "parametros",
    "financeiro",
    "itens_Remanesc",
    "itens_Consumidos",
    "itens_PC",
    "aditivos",
    "posicao_contratual",
    "itens_RC",
    "historico_VU",
    "RESULTADOS",
}
_ABAS_L1_ADICIONAIS = {
    "comparativo_VTA",
    "posicao_referencia",
    "MEMORIA_RESULTADOS",
}
# `cobertura_temporal` e `CICLO_EM_EXECUCAO` sao opcionais: o runtime sempre
# aceitou a Coleta sem elas (a segunda e acrescentada em runtime; a primeira
# tem fail-safe proprio no motor temporal). Nao podem decidir a linhagem.
# `RESULTADOS_DETALHE` (Coleta 11.2) e a antiga RESULTADOS renomeada; so existe
# a partir da 11.2 e, portanto, tambem nao pode decidir a linhagem.
_ABAS_L1_OPCIONAIS = {"cobertura_temporal", "CICLO_EM_EXECUCAO", "RESULTADOS_DETALHE"}
_ABAS_L1_PERMITIDAS = _ABAS_BASE | _ABAS_L1_ADICIONAIS | _ABAS_L1_OPCIONAIS

_PARAMETROS_BASE = {
    "computar_nesta_apuracao",
    "ciclo",
    "data_inicio",
    "data_fim",
    "percentual_do_ciclo",
    "fator_acumulado",
    "situacao",
}
_PARAMETROS_L1 = _PARAMETROS_BASE | {
    "inicio_efeito_financeiro",
    "tipo_registro",
    "ordem",
    "competencia",
    "valor_indice",
    "fator_mensal",
    "variacao_final",
    "metodo_fonte",
}
_ITENS_PC_COM_NUMERO = {
    "numero_pc",
    "data_pc",
    "ciclo_pc",
    "valor_pc",
    "fator_acumulado",
    "pc_pago_a_contratada",
}
_POSICAO_BASE = {
    "item",
    "vu_original",
    "qtd_base_original",
    "qtd_rem_ajustada_c4",
    "check_posicao_contratual",
}
_POSICAO_L1 = _POSICAO_BASE | {
    "ciclo_nascimento",
    "eh_novo_item",
    "data_efeito_inicial",
    "ciclo_nascimento_data",
}


def _norm(valor: Any) -> str:
    texto = str(valor or "").strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return "_".join(texto.replace("\n", " ").split())


def _cabecalhos(wb, aba: str) -> set[str]:
    if aba not in wb.sheetnames:
        return set()
    return {_norm(c.value) for c in wb[aba][1] if _norm(c.value)}


def _marcador_publico(wb) -> str | None:
    if "CONTROLE" not in wb.sheetnames:
        return None
    ws = wb["CONTROLE"]
    for linha in range(1, min(int(ws.max_row or 1), 60) + 1):
        if _norm(ws.cell(linha, 1).value) == "modelo_de_coleta":
            valor = ws.cell(linha, 2).value
            return str(valor).strip() if valor not in (None, "") else None
    return None


def _estrutura_l1(wb) -> tuple[bool, list[str]]:
    abas = set(wb.sheetnames)
    parametros = _cabecalhos(wb, "parametros")
    pcs = _cabecalhos(wb, "itens_PC")
    posicao = _cabecalhos(wb, "posicao_contratual")
    consumidos = _cabecalhos(wb, "itens_Consumidos")
    evidencias = [
        "abas-base+comparativo/posicao-referencia/memoria",
        "parametros-com-efeito-financeiro-e-memoria",
        "itens-PC-com-NUMERO_PC",
        "posicao-contratual-com-ciclo-de-nascimento",
    ]
    ok = (
        _ABAS_BASE | _ABAS_L1_ADICIONAIS <= abas
        and abas <= _ABAS_L1_PERMITIDAS
        and _PARAMETROS_L1 <= parametros
        and _ITENS_PC_COM_NUMERO <= pcs
        and _POSICAO_L1 <= posicao
    )
    return ok, evidencias


def _estrutura_l2(wb) -> tuple[bool, list[str]]:
    abas = set(wb.sheetnames)
    parametros = _cabecalhos(wb, "parametros")
    pcs = _cabecalhos(wb, "itens_PC")
    posicao = _cabecalhos(wb, "posicao_contratual")
    consumidos = _cabecalhos(wb, "itens_Consumidos")
    evidencias = [
        "onze-abas-base-sem-camadas-L1",
        "parametros-enxutos-sem-inicio-de-efeito",
        "itens-PC-com-NUMERO_PC",
        "posicao-contratual-base-sem-ciclo-de-nascimento",
        "consumidos-sem-bloco-de-ajuste",
    ]
    ok = (
        abas == _ABAS_BASE
        and _PARAMETROS_BASE <= parametros
        and "inicio_efeito_financeiro" not in parametros
        and "tipo_registro" not in parametros
        and _ITENS_PC_COM_NUMERO <= pcs
        and _POSICAO_BASE <= posicao
        and "ciclo_nascimento" not in posicao
        and "ajuste_status" not in consumidos
    )
    return ok, evidencias


def _fingerprint(wb) -> str:
    estrutura = {
        "abas": sorted(wb.sheetnames),
        "parametros": sorted(_cabecalhos(wb, "parametros")),
        "itens_PC": sorted(_cabecalhos(wb, "itens_PC")),
        "posicao_contratual": sorted(_cabecalhos(wb, "posicao_contratual")),
        "itens_Consumidos": sorted(_cabecalhos(wb, "itens_Consumidos")),
    }
    serializado = json.dumps(
        estrutura, ensure_ascii=True, sort_keys=True, separators=(",", ":")
    )
    return sha256(serializado.encode("utf-8")).hexdigest()


def detectar_linhagem_coleta(wb) -> dict[str, Any]:
    """Classifica a Coleta sem usar nome, data, hash do arquivo ou dados fiscais."""
    marcador = _marcador_publico(wb)
    l1, evidencias_l1 = _estrutura_l1(wb)
    l2, evidencias_l2 = _estrutura_l2(wb)
    fingerprint = _fingerprint(wb)

    if marcador in COLETA_VERSOES_ACEITAS and l1:
        # Familia 11.x: 11.0 (sem o quadro "Execucao sem efeito financeiro") e
        # 11.1 (atual) tem a mesma estrutura de abas/cabecalhos e o mesmo motor.
        codigo = LINHAGEM_COLETA_11
        evidencias = [f"marcador-publico-{marcador}", *evidencias_l1]
    elif marcador is None and l1:
        codigo = LINHAGEM_PRE_11_L1
        evidencias = evidencias_l1
    elif marcador is None and l2:
        codigo = LINHAGEM_PRE_11_L2
        evidencias = evidencias_l2
    else:
        codigo = LINHAGEM_NAO_HOMOLOGADA
        evidencias = []

    suportadas = {
        LINHAGEM_COLETA_11,
        LINHAGEM_PRE_11_L1,
        LINHAGEM_PRE_11_L2,
    }
    return {
        "codigo": codigo,
        "suportada": codigo in suportadas,
        "marcador_publico": marcador,
        "modelo_canonico": COLETA_VERSION if codigo in suportadas else None,
        "compatibilidade_aplicada": codigo in {
            LINHAGEM_PRE_11_L1, LINHAGEM_PRE_11_L2
        },
        "fingerprint": fingerprint,
        "evidencias": evidencias,
        "janela_anterior": COLETA_COMPATIBILIDADE_ANTERIORES,
    }


def mensagem_linhagem_nao_homologada(deteccao: dict[str, Any]) -> str:
    marcador = deteccao.get("marcador_publico")
    detalhe = (
        f"marcador Modelo de Coleta={marcador!r} não suportado"
        if marcador is not None
        else "estrutura sem marcador fora das duas linhagens anteriores homologadas"
    )
    return (
        "Arquivo de Coleta de versão anterior ou desconhecida não homologado para "
        "compatibilidade retroativa: "
        f"{detalhe}. Nenhuma versão foi inferida por data ou nome de arquivo."
    )
