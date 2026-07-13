"""Atualiza o ``ist.csv`` exclusivamente pela tabela oficial da Anatel.

O atualizador é conservador: valida a página, confere a sobreposição histórica
e somente acrescenta competências mensais novas e contínuas. Qualquer
divergência interrompe a execução antes da escrita.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from html.parser import HTMLParser
from io import StringIO
from pathlib import Path
from urllib.parse import urlparse

import requests


FONTE_IST_ANATEL = (
    "https://www.gov.br/anatel/pt-br/regulado/competicao/"
    "tarifas-e-precos/valores-do-ist"
)
ROOT = Path(__file__).resolve().parents[1]
IST_CSV = ROOT / "ist.csv"

MESES = {
    "jan": 1,
    "fev": 2,
    "mar": 3,
    "abr": 4,
    "mai": 5,
    "jun": 6,
    "jul": 7,
    "ago": 8,
    "set": 9,
    "out": 10,
    "nov": 11,
    "dez": 12,
}
MESES_INVERSO = {numero: nome for nome, numero in MESES.items()}


class ErroAtualizacaoIST(RuntimeError):
    """Falha segura na leitura, validação ou atualização do IST."""


@dataclass(frozen=True, order=True)
class RegistroIST:
    competencia: date
    indice: Decimal


class _LeitorTabelas(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tabelas: list[list[list[str]]] = []
        self._nivel_tabela = 0
        self._tabela: list[list[str]] | None = None
        self._linha: list[str] | None = None
        self._celula: list[str] | None = None

    def handle_starttag(self, tag: str, attrs) -> None:  # noqa: ANN001
        del attrs
        tag = tag.lower()
        if tag == "table":
            self._nivel_tabela += 1
            if self._nivel_tabela == 1:
                self._tabela = []
        elif self._nivel_tabela == 1 and tag == "tr":
            self._linha = []
        elif self._nivel_tabela == 1 and tag in {"td", "th"}:
            self._celula = []

    def handle_data(self, data: str) -> None:
        if self._celula is not None:
            self._celula.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if self._nivel_tabela == 1 and tag in {"td", "th"} and self._celula is not None:
            texto = " ".join("".join(self._celula).split())
            if self._linha is not None:
                self._linha.append(texto)
            self._celula = None
        elif self._nivel_tabela == 1 and tag == "tr":
            if self._tabela is not None and self._linha:
                self._tabela.append(self._linha)
            self._linha = None
        elif tag == "table" and self._nivel_tabela:
            if self._nivel_tabela == 1 and self._tabela:
                self.tabelas.append(self._tabela)
                self._tabela = None
            self._nivel_tabela -= 1


def _competencia(texto: str) -> date | None:
    normalizado = texto.strip().lower().replace("maio", "mai")
    encontrado = re.match(r"^([a-zç]{3})\s*/\s*(\d{2}|\d{4})", normalizado)
    if not encontrado:
        return None
    mes = MESES.get(encontrado.group(1))
    if mes is None:
        return None
    ano = int(encontrado.group(2))
    if ano < 100:
        ano += 2000
    return date(ano, mes, 1)


def _indice(texto: str) -> Decimal | None:
    valor = texto.replace("\xa0", " ").strip().replace(" ", "")
    encontrado = re.search(r"(?:\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+(?:[.,]\d+)?)", valor)
    if not encontrado:
        return None
    numero = encontrado.group(0)
    if "," in numero:
        numero = numero.replace(".", "").replace(",", ".")
    try:
        resultado = Decimal(numero)
    except InvalidOperation:
        return None
    return resultado if resultado > 0 else None


def extrair_registros_ist(html: str) -> list[RegistroIST]:
    leitor = _LeitorTabelas()
    leitor.feed(html)

    por_competencia: dict[date, Decimal] = {}
    for tabela in leitor.tabelas:
        for linha in tabela:
            if len(linha) < 3:
                continue
            competencia = _competencia(linha[0])
            indice = _indice(linha[2])
            if competencia is None or indice is None:
                continue
            anterior = por_competencia.get(competencia)
            if anterior is not None and anterior != indice:
                raise ErroAtualizacaoIST(
                    f"A Anatel apresenta dois valores para {formatar_competencia(competencia)}: "
                    f"{anterior} e {indice}."
                )
            por_competencia[competencia] = indice

    return [RegistroIST(comp, por_competencia[comp]) for comp in sorted(por_competencia)]


def _mes_seguinte(competencia: date) -> date:
    if competencia.month == 12:
        return date(competencia.year + 1, 1, 1)
    return date(competencia.year, competencia.month + 1, 1)


def formatar_competencia(competencia: date) -> str:
    return f"{MESES_INVERSO[competencia.month]}/{str(competencia.year)[-2:]}"


def formatar_indice(indice: Decimal) -> str:
    return f"{indice:.3f}".replace(".", ",")


def validar_fonte(registros: list[RegistroIST], hoje: date | None = None) -> None:
    if len(registros) < 48:
        raise ErroAtualizacaoIST(
            f"A página retornou somente {len(registros)} competências; estrutura oficial não reconhecida."
        )
    hoje = hoje or date.today()
    atual = date(hoje.year, hoje.month, 1)
    ultima = registros[-1].competencia
    if ultima > atual:
        raise ErroAtualizacaoIST("A fonte retornou uma competência futura.")
    limite = atual
    for _ in range(6):
        limite = date(limite.year - 1, 12, 1) if limite.month == 1 else date(limite.year, limite.month - 1, 1)
    if ultima < limite:
        raise ErroAtualizacaoIST(
            f"A última competência oficial ({formatar_competencia(ultima)}) está defasada em mais de seis meses."
        )


def baixar_registros_ist(timeout: int = 30) -> list[RegistroIST]:
    resposta = requests.get(
        FONTE_IST_ANATEL,
        headers={
            "User-Agent": "cl8us-ist-monitor/1.0 (+https://reajustes.streamlit.app/)",
            "Accept": "text/html,application/xhtml+xml",
        },
        timeout=timeout,
    )
    resposta.raise_for_status()
    destino = urlparse(resposta.url)
    host = (destino.hostname or "").lower()
    if destino.scheme != "https" or not (host == "gov.br" or host.endswith(".gov.br")):
        raise ErroAtualizacaoIST(f"Redirecionamento fora do domínio oficial: {resposta.url}")
    if "Índice de Serviços de Telecomunicações" not in resposta.text:
        raise ErroAtualizacaoIST("A página recebida não foi reconhecida como a página oficial do IST.")

    registros = extrair_registros_ist(resposta.text)
    validar_fonte(registros)
    return registros


def ler_registros_locais(caminho: Path = IST_CSV) -> list[RegistroIST]:
    conteudo = caminho.read_text(encoding="utf-8-sig")
    leitor = csv.reader(StringIO(conteudo), delimiter=";")
    linhas = list(leitor)
    if not linhas or [item.strip().upper() for item in linhas[0][:2]] != ["MES_ANO", "INDICE_NIVEL"]:
        raise ErroAtualizacaoIST("O ist.csv local não possui o cabeçalho MES_ANO;INDICE_NIVEL.")

    por_competencia: dict[date, Decimal] = {}
    for numero_linha, linha in enumerate(linhas[1:], start=2):
        if not linha or not any(celula.strip() for celula in linha):
            continue
        if len(linha) < 2:
            raise ErroAtualizacaoIST(f"Linha {numero_linha} inválida no ist.csv.")
        competencia = _competencia(linha[0])
        indice = _indice(linha[1])
        if competencia is None or indice is None:
            raise ErroAtualizacaoIST(f"Linha {numero_linha} inválida no ist.csv.")
        if competencia in por_competencia:
            raise ErroAtualizacaoIST(f"Competência duplicada no ist.csv: {formatar_competencia(competencia)}.")
        por_competencia[competencia] = indice

    if not por_competencia:
        raise ErroAtualizacaoIST("O ist.csv local não contém registros válidos.")
    return [RegistroIST(comp, por_competencia[comp]) for comp in sorted(por_competencia)]


def planejar_atualizacao(
    locais: list[RegistroIST], oficiais: list[RegistroIST]
) -> list[RegistroIST]:
    oficial_por_data = {registro.competencia: registro.indice for registro in oficiais}

    # As 24 competências mais recentes protegem contra revisão silenciosa da base.
    for registro in locais[-24:]:
        oficial = oficial_por_data.get(registro.competencia)
        if oficial is None:
            raise ErroAtualizacaoIST(
                f"A competência local {formatar_competencia(registro.competencia)} não foi encontrada na fonte."
            )
        if oficial != registro.indice:
            raise ErroAtualizacaoIST(
                f"Divergência histórica em {formatar_competencia(registro.competencia)}: "
                f"local={registro.indice} oficial={oficial}."
            )

    ultima_local = locais[-1].competencia
    novos = [registro for registro in oficiais if registro.competencia > ultima_local]
    esperado = _mes_seguinte(ultima_local)
    for registro in novos:
        if registro.competencia != esperado:
            raise ErroAtualizacaoIST(
                f"Lacuna entre {formatar_competencia(ultima_local)} e "
                f"{formatar_competencia(registro.competencia)}."
            )
        esperado = _mes_seguinte(esperado)
    return novos


def anexar_registros(caminho: Path, novos: list[RegistroIST]) -> None:
    if not novos:
        return
    bytes_originais = caminho.read_bytes()
    tem_bom_utf8 = bytes_originais.startswith(b"\xef\xbb\xbf")
    conteudo = bytes_originais.decode("utf-8-sig").rstrip("\r\n") + "\n"
    conteudo += "".join(
        f"{formatar_competencia(registro.competencia)};{formatar_indice(registro.indice)}\n"
        for registro in novos
    )
    temporario = caminho.with_name(f".{caminho.name}.tmp")
    try:
        temporario.write_text(
            conteudo,
            encoding="utf-8-sig" if tem_bom_utf8 else "utf-8",
            newline="",
        )
        temporario.replace(caminho)
    finally:
        temporario.unlink(missing_ok=True)


def executar(caminho: Path = IST_CSV, *, dry_run: bool = False) -> list[RegistroIST]:
    locais = ler_registros_locais(caminho)
    oficiais = baixar_registros_ist()
    novos = planejar_atualizacao(locais, oficiais)

    print(f"IST local: {formatar_competencia(locais[-1].competencia)} ({formatar_indice(locais[-1].indice)})")
    print(
        f"IST Anatel: {formatar_competencia(oficiais[-1].competencia)} "
        f"({formatar_indice(oficiais[-1].indice)})"
    )
    if not novos:
        print("Nenhuma competência nova disponível.")
        return []

    resumo = ", ".join(formatar_competencia(registro.competencia) for registro in novos)
    if dry_run:
        print(f"Simulação: seriam acrescentadas {len(novos)} competência(s): {resumo}.")
    else:
        anexar_registros(caminho, novos)
        print(f"ist.csv atualizado com {len(novos)} competência(s): {resumo}.")
    return novos


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Atualiza o IST pela página oficial da Anatel.")
    parser.add_argument("--dry-run", action="store_true", help="Valida e informa sem alterar o arquivo.")
    parser.add_argument("--arquivo", type=Path, default=IST_CSV, help="Caminho do ist.csv.")
    argumentos = parser.parse_args(argv)
    try:
        executar(argumentos.arquivo, dry_run=argumentos.dry_run)
    except (ErroAtualizacaoIST, OSError, requests.RequestException) as exc:
        print(f"ERRO SEGURO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
