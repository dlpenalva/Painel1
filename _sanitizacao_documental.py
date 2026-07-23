"""Sanitizacao de emojis/pictogramas para documentos entregues ao usuario.

Regra de negocio (pacote documental pos-homologacao):
    - A interface WEB pode usar emojis livremente.
    - Nenhum ARQUIVO entregue para download (DOCX/XLSX/TXT/PDF) pode conter
      simbolos pictograficos/emoji (checkmark, alertas, circulos coloridos,
      envelope, setas de UI etc.).

O que NAO deve ser removido: acentos, simbolos monetarios (R$), marcadores
textuais como "•", travessao "–"/"—", e demais caracteres normais da lingua
portuguesa. Este modulo remove apenas a faixa pictografica/emoji.
"""
from __future__ import annotations

import re
from typing import Any

# Faixas Unicode de emoji/pictogramas. Preserva pontuacao geral (2000-206F, onde
# vivem "•" U+2022, "–" U+2013 e "—" U+2014) e simbolos monetarios.
_EMOJI_RANGES = (
    "\U0001F300-\U0001FAFF"   # Misc Symbols & Pictographs, Emoticons, Transport,
                               # Supplemental Symbols, Geometric Shapes Extended,
                               # Symbols & Pictographs Extended-A (inclui circulos
                               # coloridos e triangulo invertido usados na UI)
    "\U00002600-\U000027BF"   # Misc Symbols + Dingbats (checkmark, envelope etc.)
    "\U00002B00-\U00002BFF"   # Misc Symbols and Arrows (quadrado preto, estrela)
    "\U0001F000-\U0001F0FF"   # Mahjong/Domino/Playing cards
    "\U0001F1E6-\U0001F1FF"   # Regional indicators (bandeiras)
    "\U0000FE00-\U0000FE0F"   # Variation Selectors (VS1-VS16)
    "\U00002190-\U000021FF"   # Setas (usadas como pictogramas de UI)
    "\U00002300-\U000023FF"   # Misc Technical (relogio, ampulheta ...)
    "\U000025A0-\U000025FF"   # Geometric Shapes (quadrado, triangulo, circulo)
    "\U0000200D"               # Zero Width Joiner (sequencias de emoji)
)
_EMOJI_RE = re.compile(f"[{_EMOJI_RANGES}]")

# Residuos de separadores orfaos deixados apos remover o pictograma (ex.: o
# padrao "<verde> | <invertido> CICLO ..." vira " |  CICLO ...").
_SEP_ORFAO_RE = re.compile(r"^[\s|/–—-]+|[\s|/–—-]+$")
_ESPACOS_RE = re.compile(r"[ \t]{2,}")


def remover_emojis_leve(texto: Any) -> str:
    """Remove SOMENTE emojis/pictogramas, sem tocar em espacos, travessoes ou
    marcadores. Indicado para runs/celulas de documentos onde "—", "•" e o
    espacamento devem ser preservados exatamente.
    """
    if texto is None:
        return ""
    s = str(texto)
    if not s:
        return s
    return _EMOJI_RE.sub("", s)


def remover_emojis(texto: Any) -> str:
    """Remove emojis/pictogramas preservando texto institucional legitimo.

    Nao mexe em quebras de linha, acentos, "R$", "•" ou travessoes. Colapsa
    espacos duplicados (nao quebras de linha) e apara separadores orfaos nas
    bordas de cada linha.
    """
    if texto is None:
        return ""
    s = str(texto)
    if not s:
        return s
    s = _EMOJI_RE.sub("", s)
    # Normaliza cada linha: colapsa espacos e remove separadores orfaos nas
    # bordas, preservando "•" quando for o marcador inicial legitimo.
    linhas_saida = []
    for linha in s.split("\n"):
        limpa = _ESPACOS_RE.sub(" ", linha)
        stripped = limpa.lstrip()
        if stripped.startswith("• "):
            corpo = _SEP_ORFAO_RE.sub("", stripped[2:])
            limpa = "• " + corpo
        else:
            limpa = _SEP_ORFAO_RE.sub("", limpa)
        linhas_saida.append(limpa)
    return "\n".join(linhas_saida)


def contem_emoji(texto: Any) -> bool:
    """True se o texto ainda contiver algum pictograma/emoji proibido."""
    if texto is None:
        return False
    return bool(_EMOJI_RE.search(str(texto)))
