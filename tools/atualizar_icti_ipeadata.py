"""Atualiza ``icti.csv`` somente pela serie oficial ICTI do Ipeadata/Ipea.

O arquivo local e uma copia versionada da serie oficial, nao uma fonte
independente. Nenhuma competencia e estimada, interpolada ou repetida.
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
import sys

import requests


ROOT = Path(__file__).resolve().parents[1]
ICTI_CSV = ROOT / "icti.csv"
ICTI_SERCODIGO = "DIMAC12_ICTI2"
ICTI_API_BASES = (
    "https://www.ipeadata.gov.br/api/odata4",
    "http://www.ipeadata.gov.br/api/odata4",
)


class ErroAtualizacaoICTI(RuntimeError):
    """Falha segura: o arquivo oficial local nao deve ser modificado."""


@dataclass(frozen=True, order=True)
class RegistroICTI:
    competencia: date
    taxa_mensal_percentual: Decimal


def _competencia(valor) -> date:
    texto = str(valor or "").strip()
    try:
        data = date.fromisoformat(texto[:10])
    except (TypeError, ValueError) as exc:
        raise ErroAtualizacaoICTI(f"Competência ICTI inválida: {valor!r}.") from exc
    return date(data.year, data.month, 1)


def _taxa(valor) -> Decimal:
    if isinstance(valor, bool) or valor is None:
        raise ErroAtualizacaoICTI(f"Taxa mensal ICTI inválida: {valor!r}.")
    texto = str(valor).strip().replace(",", ".")
    try:
        taxa = Decimal(texto)
    except (InvalidOperation, ValueError) as exc:
        raise ErroAtualizacaoICTI(f"Taxa mensal ICTI inválida: {valor!r}.") from exc
    if not taxa.is_finite():
        raise ErroAtualizacaoICTI(f"Taxa mensal ICTI não finita: {valor!r}.")
    return taxa


def _mes_seguinte(competencia: date) -> date:
    if competencia.month == 12:
        return date(competencia.year + 1, 1, 1)
    return date(competencia.year, competencia.month + 1, 1)


def validar_registros(registros: list[RegistroICTI], *, origem: str) -> None:
    if not registros:
        raise ErroAtualizacaoICTI(f"A série ICTI {origem} está vazia.")
    competencias = [registro.competencia for registro in registros]
    if competencias != sorted(competencias):
        raise ErroAtualizacaoICTI(f"A série ICTI {origem} não está em ordem cronológica.")
    if len(competencias) != len(set(competencias)):
        repetida = next(c for c in competencias if competencias.count(c) > 1)
        raise ErroAtualizacaoICTI(
            f"Competência duplicada na série ICTI {origem}: {repetida:%m/%Y}."
        )
    for anterior, atual in zip(registros, registros[1:]):
        if atual.competencia != _mes_seguinte(anterior.competencia):
            raise ErroAtualizacaoICTI(
                "A série ICTI "
                f"{origem} possui lacuna entre {anterior.competencia:%m/%Y} "
                f"e {atual.competencia:%m/%Y}."
            )


def extrair_registros_ipeadata(payload) -> list[RegistroICTI]:
    if not isinstance(payload, dict) or not isinstance(payload.get("value"), list):
        raise ErroAtualizacaoICTI("A resposta do Ipeadata possui estrutura inválida.")
    registros = []
    for item in payload["value"]:
        if not isinstance(item, dict):
            raise ErroAtualizacaoICTI("A resposta do Ipeadata contém registro inválido.")
        codigo = item.get("SERCODIGO")
        if codigo != ICTI_SERCODIGO:
            raise ErroAtualizacaoICTI(f"Código de série ICTI inesperado: {codigo!r}.")
        registros.append(
            RegistroICTI(
                _competencia(item.get("VALDATA")),
                _taxa(item.get("VALVALOR")),
            )
        )
    validar_registros(registros, origem="recebida do Ipeadata")
    return registros


def baixar_registros_icti(timeout: int = 30) -> list[RegistroICTI]:
    ultimo_erro = None
    endpoint = f"ValoresSerie(SERCODIGO='{ICTI_SERCODIGO}')"
    headers = {"User-Agent": "Mozilla/5.0 cl8us-icti", "Accept": "application/json"}
    for base in ICTI_API_BASES:
        try:
            resposta = requests.get(f"{base}/{endpoint}", headers=headers, timeout=timeout)
            resposta.raise_for_status()
            return extrair_registros_ipeadata(resposta.json())
        except (requests.RequestException, ValueError, ErroAtualizacaoICTI) as exc:
            ultimo_erro = exc
    raise ErroAtualizacaoICTI(
        f"Não foi possível obter uma série ICTI oficial válida. Último erro: {ultimo_erro}"
    ) from ultimo_erro


def ler_registros_locais(caminho: Path = ICTI_CSV) -> list[RegistroICTI]:
    try:
        linhas = caminho.read_text(encoding="utf-8-sig").splitlines()
    except OSError as exc:
        raise ErroAtualizacaoICTI(f"Não foi possível ler {caminho.name}: {exc}") from exc
    if not linhas or linhas[0] != "COMPETENCIA;TAXA_MENSAL_PERCENTUAL":
        raise ErroAtualizacaoICTI(
            "O icti.csv local não possui o cabeçalho esperado "
            "COMPETENCIA;TAXA_MENSAL_PERCENTUAL."
        )
    registros = []
    for numero_linha, linha in enumerate(linhas[1:], start=2):
        partes = linha.split(";")
        if len(partes) != 2:
            raise ErroAtualizacaoICTI(f"Linha {numero_linha} inválida no icti.csv.")
        try:
            competencia = date.fromisoformat(f"{partes[0].strip()}-01")
        except ValueError as exc:
            raise ErroAtualizacaoICTI(
                f"Competência inválida na linha {numero_linha} do icti.csv."
            ) from exc
        registros.append(RegistroICTI(competencia, _taxa(partes[1])))
    validar_registros(registros, origem="local")
    return registros


def planejar_atualizacao(
    locais: list[RegistroICTI], oficiais: list[RegistroICTI]
) -> list[RegistroICTI]:
    validar_registros(locais, origem="local")
    validar_registros(oficiais, origem="oficial")
    oficial_por_competencia = {r.competencia: r for r in oficiais}
    for local in locais:
        oficial = oficial_por_competencia.get(local.competencia)
        if oficial is None:
            raise ErroAtualizacaoICTI(
                "Atualização abortada: a resposta oficial não contém a competência "
                f"local {local.competencia:%m/%Y}."
            )
        if oficial.taxa_mensal_percentual != local.taxa_mensal_percentual:
            raise ErroAtualizacaoICTI(
                "Atualização abortada: divergência histórica na competência "
                f"{local.competencia:%m/%Y}."
            )
    if oficiais[-1].competencia < locais[-1].competencia:
        raise ErroAtualizacaoICTI(
            "Atualização abortada: a série oficial recebida termina antes da série local."
        )
    return [r for r in oficiais if r.competencia > locais[-1].competencia]


def _formatar_taxa(valor: Decimal) -> str:
    return format(valor, "f").replace(".", ",")


def escrever_atomico(caminho: Path, registros: list[RegistroICTI]) -> None:
    validar_registros(registros, origem="a gravar")
    conteudo = "COMPETENCIA;TAXA_MENSAL_PERCENTUAL\n" + "".join(
        f"{r.competencia:%Y-%m};{_formatar_taxa(r.taxa_mensal_percentual)}\n"
        for r in registros
    )
    temporario = caminho.with_name(f".{caminho.name}.tmp")
    try:
        temporario.write_text(conteudo, encoding="utf-8", newline="\n")
        temporario.replace(caminho)
    finally:
        temporario.unlink(missing_ok=True)


def executar(caminho: Path = ICTI_CSV, *, dry_run: bool = False) -> list[RegistroICTI]:
    oficiais = baixar_registros_icti()
    if caminho.exists():
        locais = ler_registros_locais(caminho)
        novos = planejar_atualizacao(locais, oficiais)
    else:
        locais = []
        novos = oficiais
    if not novos:
        print("icti.csv já está atualizado.")
        return []
    resumo = ", ".join(f"{r.competencia:%m/%Y}" for r in novos)
    if dry_run:
        print(f"Simulação: seriam acrescentadas {len(novos)} competência(s): {resumo}.")
    else:
        escrever_atomico(caminho, oficiais)
        print(f"icti.csv atualizado com {len(novos)} competência(s): {resumo}.")
    return novos


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Atualiza o ICTI pelo Ipeadata/Ipea.")
    parser.add_argument("--dry-run", action="store_true", help="Valida sem alterar o arquivo.")
    parser.add_argument("--arquivo", type=Path, default=ICTI_CSV, help="Caminho do icti.csv.")
    argumentos = parser.parse_args(argv)
    try:
        executar(argumentos.arquivo, dry_run=argumentos.dry_run)
    except (ErroAtualizacaoICTI, OSError, requests.RequestException) as exc:
        print(f"ERRO SEGURO: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
