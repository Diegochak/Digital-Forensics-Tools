"""Leitura das planilhas Excel conforme o mapeamento do config.yaml."""
from __future__ import annotations

import datetime as dt
import io
import re
import unicodedata
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

_LETRA = re.compile(r"^[A-Za-z]{1,3}$")


class ErroConfig(Exception):
    pass


def normalizar(valor) -> str:
    """Texto para comparação: sem acento, sem espaços extras, maiúsculo."""
    if valor is None:
        return ""
    txt = unicodedata.normalize("NFKD", str(valor))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return " ".join(txt.split()).upper()


def formatar_valor(valor, formato_data: str) -> str:
    if valor is None:
        return ""
    if isinstance(valor, dt.datetime):
        if valor.time() == dt.time(0, 0):
            return valor.strftime(formato_data)
        return valor.strftime(formato_data + " %H:%M")
    if isinstance(valor, dt.date):
        return valor.strftime(formato_data)
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def _abrir(caminho: Path):
    """Lê o arquivo para a memória antes de abrir: funciona com a planilha aberta
    no Excel e não deixa cópias dos dados na pasta temporária."""
    if not caminho.exists():
        raise ErroConfig(f"Planilha não encontrada: {caminho}")
    return load_workbook(io.BytesIO(caminho.read_bytes()), data_only=True, read_only=True)


def _abas(wb, cfg_aba) -> list:
    if cfg_aba in (None, ""):
        return [wb.worksheets[0]]
    nomes = cfg_aba if isinstance(cfg_aba, list) else [cfg_aba]
    abas = []
    for nome in nomes:
        if nome not in wb.sheetnames:
            raise ErroConfig(
                f"Aba '{nome}' não existe. Abas disponíveis: {', '.join(wb.sheetnames)}"
            )
        abas.append(wb[nome])
    return abas


def resolver_coluna(ref, cabecalho: list) -> int:
    """Converte a referência do config em índice (0-based).

    Aceita: número (3), letra ("C"), nome do cabeçalho ("Risk Level"),
    ou prefixos explícitos "col:C" / "cab:Risk Level".
    """
    if isinstance(ref, int):
        return ref - 1
    ref = str(ref).strip()
    cab_norm = [normalizar(c) for c in cabecalho]

    if ref.lower().startswith("col:"):
        return column_index_from_string(ref[4:].strip().upper()) - 1
    if ref.lower().startswith("cab:"):
        alvo = normalizar(ref[4:])
        if alvo in cab_norm:
            return cab_norm.index(alvo)
        raise ErroConfig(f"Cabeçalho '{ref[4:]}' não encontrado.")

    if normalizar(ref) in cab_norm:  # nome de cabeçalho tem prioridade
        return cab_norm.index(normalizar(ref))
    if _LETRA.match(ref):
        return column_index_from_string(ref.upper()) - 1
    raise ErroConfig(
        f"Coluna '{ref}' não encontrada. Cabeçalhos: "
        + ", ".join(str(c) for c in cabecalho if c)
    )


def ler_planilha(cfg: dict, base: Path, formato_data: str) -> list[dict]:
    """Retorna uma lista de registros {campo: valor} conforme cfg['colunas']."""
    caminho = (base / cfg["arquivo"]).resolve()
    wb = _abrir(caminho)
    linha_cab = int(cfg.get("linha_cabecalho", 1))
    primeira = int(cfg.get("primeira_linha_dados", linha_cab + 1))
    ultima = cfg.get("ultima_linha_dados")
    colunas = cfg.get("colunas") or {}
    if not colunas:
        raise ErroConfig(f"Nenhuma coluna mapeada para {caminho.name}.")

    registros = []
    for ws in _abas(wb, cfg.get("aba")):
        linhas = list(ws.iter_rows(values_only=True))
        if len(linhas) < linha_cab:
            continue
        cabecalho = list(linhas[linha_cab - 1])
        try:
            indices = {campo: resolver_coluna(ref, cabecalho) for campo, ref in colunas.items()}
        except ErroConfig as e:
            raise ErroConfig(f"[{caminho.name} / aba '{ws.title}'] {e}") from None

        fim = int(ultima) if ultima else len(linhas)
        for n in range(primeira, fim + 1):
            if n > len(linhas):
                break
            linha = linhas[n - 1]
            reg = {}
            for campo, idx in indices.items():
                valor = linha[idx] if idx < len(linha) else None
                reg[campo] = formatar_valor(valor, formato_data)
            if not any(reg.values()):
                continue  # linha vazia
            reg["_linha"] = n
            reg["_aba"] = ws.title
            reg["_arquivo"] = caminho.name
            registros.append(reg)
    wb.close()
    return registros


def mapa_colunas(cfg: dict, base: Path) -> str:
    """Texto com letra | cabeçalho | exemplo, para ajudar a montar o config."""
    caminho = (base / cfg["arquivo"]).resolve()
    wb = _abrir(caminho)
    linha_cab = int(cfg.get("linha_cabecalho", 1))
    primeira = int(cfg.get("primeira_linha_dados", linha_cab + 1))
    saida = [f"\n=== {caminho.name} | abas: {', '.join(wb.sheetnames)}"]
    for ws in _abas(wb, cfg.get("aba")):
        linhas = list(ws.iter_rows(values_only=True))
        if len(linhas) < linha_cab:
            saida.append(f"  Aba '{ws.title}': vazia")
            continue
        cab = linhas[linha_cab - 1]
        ex = linhas[primeira - 1] if len(linhas) >= primeira else [None] * len(cab)
        saida.append(f"  Aba '{ws.title}' (cabeçalho na linha {linha_cab}, {len(linhas) - linha_cab} linhas de dados)")
        for i, c in enumerate(cab):
            if c is None and (i >= len(ex) or ex[i] is None):
                continue
            exemplo = formatar_valor(ex[i] if i < len(ex) else None, "%d/%m/%Y")[:50]
            saida.append(f"    {get_column_letter(i + 1):>3} | {str(c or ''):<30} | {exemplo}")
    wb.close()
    return "\n".join(saida)
