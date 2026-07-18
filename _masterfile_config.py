"""Configuração estrutural compartilhada do Masterfile v9."""


# Fonte única da versão do Masterfile. Não espalhar hardcoded fora deste arquivo.
# v9: aba RESUMO removida; indicadores migrados para historico (Bloco 5, L45-56).
MASTERFILE_VERSION = "v9"


ABAS_OBRIGATORIAS = [
    "CONTROLE",
    "parametros",
    "financeiro",
    "itens_Remanesc",
    "itens_Consumidos",
    "aditivos",
    "historico",
    "itens_RC",
]

# Modos de leitura reconhecidos. "pc" = Pedido de Compra (aba itens_PC,
# opcional — não entra em ABAS_OBRIGATORIAS).
MODOS_VALIDOS = frozenset({"principal", "d", "pc"})

# Nome físico da aba de Pedidos de Compra (condicional ao modo "pc").
ABA_ITENS_PC = "itens_PC"
