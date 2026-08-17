"""Decisão de reidratação da apuração pós-upload no Painel do Valor Global.

Depois que o Arquivo Coleta Oficial é processado, a fonte de verdade passa a ser o
``st.session_state`` — não o widget ``st.file_uploader``, cujo estado é descartado
pelo Streamlit ao navegar para outra página (por exemplo, Adequação Orçamentária) e
voltar. Estas funções puras encapsulam quando a apuração já processada pode ser
reutilizada sem exigir novo upload nem reprocessamento.
"""
from __future__ import annotations

import hashlib
from collections.abc import MutableMapping
from typing import Any, Mapping

CHAVE_ASSINATURA_UPLOAD = "assinatura_upload_docs"
CHAVE_ASSINATURA_PROCESSADA = "assinatura_processada_upload_docs"
CHAVE_RESULTADO = "resultado_valor_global"
CHAVE_DIAGNOSTICO = "diagnostico_coleta_v2"

# Estado pertencente exclusivamente ao caso em analise. A lista e deliberadamente
# explicita: preferencias visuais, navegacao e o widget do upload atual nao entram
# aqui. O namespace da Adequacao e a unica excecao por prefixo porque a propria
# pagina cria chaves dinamicas para editores conforme competencia, janela e fator.
CHAVES_ESTADO_CASO = (
    CHAVE_ASSINATURA_PROCESSADA,
    CHAVE_RESULTADO,
    CHAVE_DIAGNOSTICO,
    "dados_admissibilidade",
    "contexto_contratual_anterior",
    "chave_analise_simples_processada",
    "processar_reajustes_multiplos_key",
    "resultado_garantia",
    "resultado_adequacao_orcamentaria",
    "checklist_processual",
    "avaliacao_aditivos_eventos",
    "infos_previas_df",
    "saneador_complemento_manual",
    "saneador_texto",
    "previa_dou",
    "garantia_historico_valores_totais",
    "garantia_historico_alteracoes",
    "upload_checklist_processual",
    "checklist_processual_editor",
    "avaliacao_aditivos_editor",
    "upload_infos_previas",
    "infos_previas_editor",
    "arquivo_sumario_executivo_pdf",
    "arquivo_despacho_saneador_docx",
    "arquivo_termo_apostila_docx",
    "arquivo_planilha_executiva_xlsx",
    "arquivo_valores_unitarios_xlsx",
    "arquivo_mapa_marcos_pdf",
    "arquivo_relatorio_executivo_pdf",
    "arquivo_minuta_apostilamento_docx",
    "arquivo_garantia_pdf",
    "arquivo_adequacao_orcamentaria_xlsx",
    "arquivo_dou_docx",
    "arquivo_checklist_processual_xlsx",
    "arquivo_avaliacao_aditivos_xlsx",
    "arquivo_infos_previas_xlsx",
    "arquivo_saneador_docx",
)
PREFIXOS_ESTADO_CASO = ("adequacao_v3_",)


def assinatura_conteudo_upload(conteudo: bytes) -> str:
    """Retorna a identidade do caso a partir dos bytes efetivamente enviados."""
    return hashlib.sha256(conteudo).hexdigest()


def id_apuracao_de_assinatura(assinatura: str) -> str:
    """Converte o SHA-256 integral do upload no identificador publico curto."""
    assinatura_normalizada = str(assinatura or "").strip().lower()
    if len(assinatura_normalizada) != 64 or any(
        caractere not in "0123456789abcdef" for caractere in assinatura_normalizada
    ):
        raise ValueError("A assinatura de origem deve ser um SHA-256 hexadecimal completo.")
    return f"APUR-{assinatura_normalizada[:16].upper()}"


def invalidar_estado_caso(
    estado: MutableMapping[str, Any], nova_assinatura: str
) -> bool:
    """Invalida derivados do caso anterior quando o conteudo do XLS muda.

    Retorna ``True`` quando houve troca de caso. Se a assinatura e a mesma, nao
    altera nenhuma chave, preservando resultados e documentos em reruns normais.
    """
    if estado.get(CHAVE_ASSINATURA_UPLOAD) == nova_assinatura:
        return False

    for chave in CHAVES_ESTADO_CASO:
        estado.pop(chave, None)

    for chave in list(estado):
        if any(str(chave).startswith(prefixo) for prefixo in PREFIXOS_ESTADO_CASO):
            estado.pop(chave, None)

    estado[CHAVE_ASSINATURA_UPLOAD] = nova_assinatura
    return True


def apuracao_persistida_valida(estado: Mapping[str, Any]) -> bool:
    """Indica se há apuração processada reutilizável no estado da sessão.

    Considera válida apenas quando a assinatura do upload processado e o resultado
    canônico coexistem — ambos são gravados juntos exclusivamente no caminho de
    sucesso do processamento. Após a invalidação por arquivo diferente/incompatível
    (que remove essas chaves), a função volta a retornar ``False``, forçando novo
    processamento e impedindo reutilização indevida entre arquivos distintos.
    """
    if not estado.get(CHAVE_ASSINATURA_PROCESSADA):
        return False
    return estado.get(CHAVE_RESULTADO) is not None
