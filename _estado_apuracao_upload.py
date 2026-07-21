"""Decisão de reidratação da apuração pós-upload no Painel do Valor Global.

Depois que o Arquivo Coleta Oficial é processado, a fonte de verdade passa a ser o
``st.session_state`` — não o widget ``st.file_uploader``, cujo estado é descartado
pelo Streamlit ao navegar para outra página (por exemplo, Adequação Orçamentária) e
voltar. Estas funções puras encapsulam quando a apuração já processada pode ser
reutilizada sem exigir novo upload nem reprocessamento.
"""
from __future__ import annotations

from typing import Any, Mapping

CHAVE_ASSINATURA_PROCESSADA = "assinatura_processada_upload_docs"
CHAVE_RESULTADO = "resultado_valor_global"
CHAVE_DIAGNOSTICO = "diagnostico_coleta_v2"


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
