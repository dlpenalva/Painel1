"""ConfiguraÃ§Ã£o estrutural do Masterfile v10 RC.

v10 adiciona: historico_VU, itens_Consumidos com Qtd+Valor por ciclo,
itens_PC com suporte a PC UnitÃ¡rio e PC Global/Multi-item (SAP).
NÃ£o substituir _masterfile_config.py (v9 continua ativa na main).
"""

MASTERFILE_VERSION = "v10.2"           # marcador estrutural do template (nao alterar)
MASTERFILE_RELEASE_VERSION = "v10.5.3" # versao de release publicada — usar em labels visiveis e como chave de cache
TEMPLATE_FILENAME  = "MASTERFILE_v10_2_CONSUMIDOS_FINAL.xlsx"

ABAS_OBRIGATORIAS_V10 = [
    "CONTROLE",
    "parametros",
    "financeiro",
    "itens_Remanesc",
    "itens_Consumidos",
    "aditivos",
    "historico",
    "itens_RC",
    "historico_VU",   # nova v10
]

# Modos de leitura reconhecidos (mantÃ©m os da v9 + sem alteraÃ§Ã£o)
MODOS_VALIDOS_V10 = frozenset({"principal", "d", "pc"})

# Nome fÃ­sico das abas novas/evoluÃ­das
ABA_HISTORICO_VU   = "historico_VU"
ABA_ITENS_PC       = "itens_PC"
ABA_VALIDACOES     = "validacoes"

# Tipos de PC suportados
TIPOS_PC_VALIDOS = frozenset({"UnitÃ¡rio", "Global/Multi-item"})

# Colunas canÃ´nicas de historico_VU (coluna A = ITEM, obrigatÃ³rio)
COLUNAS_HISTORICO_VU = [
    "ITEM",
    "DESCRIÃ‡ÃƒO",
    "QTD_BASE_REFERENCIA",
    "VU_ORIGINAL",
    "VU_C0",
    "VU_C1",
    "VU_C2",
    "VU_C3",
    "VU_C4",
    "VU_VIGENTE_ULTIMO_CICLO",
    "FATOR_ACUMULADO_ULTIMO_CICLO",
    "VARIACAO_ACUMULADA",
    "FONTE",
    "OBSERVACAO",
]

# Colunas canonicas de itens_Consumidos v10.1 (A-Q, 17 colunas)
# D=VALOR_TOTAL adicionado apos VU_ORIGINAL.
COLUNAS_ITENS_CONSUMIDOS_V10 = [
    "ITEM", "QTD_CONTRATADA", "VU_ORIGINAL", "VALOR_TOTAL",
    "QTD_CONS_C0", "VALOR_CONS_C0",
    "QTD_CONS_C1", "VALOR_CONS_C1",
    "QTD_CONS_C2", "VALOR_CONS_C2",
    "QTD_CONS_C3", "VALOR_CONS_C3",
    "QTD_CONS_C4", "VALOR_CONS_C4",
    "CONS_QTD_TOTAL", "CONS_VALOR_TOTAL", "CHECK",
]

# Colunas canonicas de itens_PC v10.5.3 Fase 1.
# A:L preserva a estrutura homologada; W:AD prepara auditoria PC/VTA sem calculo.
COLUNAS_ITENS_PC_V10 = [
    "NUMERO_PC",
    "DATA_PC",
    "CICLO_PC",
    "VALOR_PC",
    "FATOR_ACUMULADO",
    "VALOR_ATUALIZADO",
    "PC_PAGO_A_CONTRATADA",
    "RETROATIVO_RECONHECIDO_A_PAGAR",
    "VALOR_ATUALIZADO_EM_ANALISE",
    "DELTA_POTENCIAL",
    "CHECK_PC_FINANCEIRO",
    "COMPUTA_VTA",
    "TIPO_PARCELA",
    "ORIGEM_DADO",
    "TIPO_FINANCEIRO",
    "FONTE_PARCELA",
    "JA_REFLETIDO_EM",
    "STATUS_CONSOLIDACAO",
    "JUSTIFICATIVA_VTA",
]

# Aba opcional v10.1 — entrada fiscal consolidada
ABA_EXECUCAO_SALDO = "itens_Execucao_Saldo"

# Abas obrigatorias devem estar presentes; opcionais apenas alertam se ausentes
ABAS_OPCIONAIS_V10 = [ABA_EXECUCAO_SALDO]

# Colunas canonicas de itens_Execucao_Saldo v10.1 (A-J, 10 colunas)
# Removidas DESCRICAO_REFERENCIAL e OBSERVACAO; REQUISICAO_SAP renomeada para PC.
COLUNAS_EXECUCAO_SALDO_V10 = [
    "ITEM",
    "QTD_CONTRATADA",
    "VU_ORIGINAL",
    "VALOR_TOTAL_ORIGINAL",
    "PC",
    "QTD_EMITIDA",
    "VALOR_EMITIDO",
    "QTD_SALDO",
    "VALOR_SALDO",
    "CHECK_FISICO",
]

# Aliases para compatibilidade com v9
_ALIAS_ABAS_LOWER_V10: dict[str, str] = {
    "itens_a": "itens_remanesc",
    "itens_b": "itens_consumidos",
    "aditivos": "aditivos",  # noqa
}
