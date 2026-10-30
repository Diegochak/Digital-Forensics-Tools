"""Log CSV de tudo que foi gerado — evidência de auditoria e trava contra duplicidade."""
from __future__ import annotations

import csv
import datetime as dt
from pathlib import Path

COLUNAS = [
    "data_hora", "modo", "chave", "arquivo", "aba", "linha",
    "criticidade", "para", "cc", "assunto", "status", "detalhe",
]


class Registro:
    def __init__(self, caminho: Path):
        self.caminho = caminho
        self.caminho.parent.mkdir(parents=True, exist_ok=True)

    def ja_criados(self) -> set[str]:
        if not self.caminho.exists():
            return set()
        with self.caminho.open(encoding="utf-8-sig", newline="") as f:
            return {
                l["chave"] for l in csv.DictReader(f, delimiter=";")
                if l.get("status", "").startswith("RASCUNHO")
            }

    def gravar(self, itens: list[dict], modo: str) -> None:
        novo = not self.caminho.exists()
        agora = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.caminho.open("a", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=COLUNAS, delimiter=";", extrasaction="ignore")
            if novo:
                w.writeheader()
            for it in itens:
                w.writerow({**it, "data_hora": agora, "modo": modo})
