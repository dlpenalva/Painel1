"""Guarda do estado de upload: um arquivo, um resultado.

O Painel1 admite uma única entrada oficial — o `Coleta_Reajuste.xlsx` baixado da
Calculadora. Este módulo não lê o XLS, não calcula e não decide regra de
negócio: ele apenas amarra o que já foi apurado ao arquivo exato que o originou,
para que nenhum resultado, confirmação ou documento sobreviva à troca de
contrato no mesmo slot de sessão.

A ausência de dado permanece ausência. Nada aqui converte falta em zero nem em
DataFrame vazio: um arquivo recusado não produz resultado, produz diagnóstico.
"""

from __future__ import annotations

import hashlib
from typing import Any, MutableMapping

# Muda quando o formato de `ObjetoProcesso.como_dicionario()` mudar; resultado
# gravado sob outra versão é descartado em vez de reaproveitado.
VERSAO_CONTRATO_DADOS = "v4.objeto_processo"

ORIGEM_COLETA_OFICIAL = "coleta_reajuste_oficial"
ORIGEM_NAO_RECONHECIDA = "arquivo_nao_reconhecido"

CHAVE_PROCEDENCIA = "procedencia_upload"

#: Tudo que só existe por causa de um upload anterior. A troca de arquivo apaga
#: esta lista inteira. Ficam de fora, deliberadamente, os dados digitados pelo
#: usuário em outras telas (`infos_previas_df`, `checklist_processual`,
#: `avaliacao_aditivos_eventos`) e o contexto herdado da Calculadora
#: (`dados_admissibilidade`): nenhum deles deriva do XLS enviado aqui.
CHAVES_DERIVADAS_DO_UPLOAD = (
    "resultado_valor_global",
    "diagnostico_coleta_v2",
    "resultado_garantia",
    "resultado_adequacao_orcamentaria",
    "saneador_texto",
    "saneador_complemento_manual",
    "previa_dou",
    "arquivo_planilha_executiva_xlsx",
    "arquivo_valores_unitarios_xlsx",
    "arquivo_mapa_marcos_pdf",
    "arquivo_relatorio_executivo_pdf",
    "arquivo_minuta_apostilamento_docx",
    "arquivo_garantia_pdf",
    "arquivo_previsao_orcamentaria_docx",
    "arquivo_dou_docx",
    "arquivo_saneador_docx",
    "arquivo_checklist_processual_xlsx",
    "arquivo_avaliacao_aditivos_xlsx",
    "arquivo_infos_previas_xlsx",
)

MENSAGEM_ARQUIVO_NAO_RECONHECIDO = (
    "Arquivo não reconhecido como Coleta_Reajuste.xlsx oficial. "
    "O Painel processa exclusivamente o arquivo baixado da Calculadora de "
    "Reajustes, cuja estrutura garante que os valores exibidos e os documentos "
    "emitidos venham do próprio XLS."
)

DETALHES_ARQUIVO_NAO_RECONHECIDO = (
    "As abas CONTROLE, parametros e financeiro não foram localizadas juntas no arquivo enviado.",
    "Nenhum valor foi calculado e nenhum documento foi liberado: um arquivo incompatível não produz resultado.",
    "Baixe o modelo em “1 · Baixar arquivo de trabalho”, preencha-o e envie novamente.",
)


def sha256_do_arquivo(conteudo: bytes) -> str:
    """Identidade integral do arquivo recebido — bytes, não nome nem tamanho."""

    return hashlib.sha256(conteudo).hexdigest()


def limpar_estados_derivados(estado: MutableMapping[str, Any]) -> None:
    """Apaga tudo que veio do upload anterior, inclusive buffers documentais."""

    for chave in CHAVES_DERIVADAS_DO_UPLOAD:
        estado.pop(chave, None)


def procedencia_registrada(estado: MutableMapping[str, Any]) -> dict[str, Any] | None:
    """A procedência do último arquivo avaliado, ou None se nunca houve upload."""

    procedencia = estado.get(CHAVE_PROCEDENCIA)
    return procedencia if isinstance(procedencia, dict) else None


def registrar_upload(
    estado: MutableMapping[str, Any],
    *,
    sha256: str,
    origem: str,
    aceito: bool,
    motivo: str = "",
) -> None:
    """Registra a que arquivo o estado atual corresponde."""

    estado[CHAVE_PROCEDENCIA] = {
        "sha256": sha256,
        "origem": origem,
        "versao_contrato": VERSAO_CONTRATO_DADOS,
        "aceito": aceito,
        "motivo": motivo,
    }


def upload_ja_processado(estado: MutableMapping[str, Any], sha256: str) -> bool:
    """Diz se o resultado em sessão pertence a este arquivo e a esta versão.

    Exige o resultado presente: procedência sem resultado é estado incompleto e
    deve ser reprocessada, não reaproveitada.
    """

    procedencia = procedencia_registrada(estado)
    if not procedencia:
        return False
    return (
        procedencia.get("sha256") == sha256
        and procedencia.get("versao_contrato") == VERSAO_CONTRATO_DADOS
        and bool(procedencia.get("aceito"))
        and estado.get("resultado_valor_global") is not None
    )


def documentos_liberados(estado: MutableMapping[str, Any]) -> bool:
    """Documentos só existem sobre um arquivo oficial aceito e ainda vigente."""

    procedencia = procedencia_registrada(estado)
    if not procedencia:
        return False
    return (
        bool(procedencia.get("aceito"))
        and procedencia.get("origem") == ORIGEM_COLETA_OFICIAL
        and procedencia.get("versao_contrato") == VERSAO_CONTRATO_DADOS
        and estado.get("resultado_valor_global") is not None
    )
