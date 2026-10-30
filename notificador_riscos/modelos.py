"""Carrega os modelos de e-mail (Word .docx ou .htm) e preenche os campos."""
from __future__ import annotations

import html
import re
from pathlib import Path

import jinja2
import jinja2.meta
import mammoth

ESTILO_BASE = """<style>
body, p, td, li { font-family: Aptos, Calibri, Arial, sans-serif; font-size: 11pt; color: #1f1f1f; }
p { margin: 0 0 10px 0; }
td p, th p { margin: 0; }
table { border-collapse: collapse; margin: 8px 0 14px 0; }
td, th { border: 1px solid #c8c8c8; padding: 5px 9px; vertical-align: top; text-align: left; }
th { background: #f2f2f2; }
h1, h2, h3 { font-family: Aptos, Calibri, Arial, sans-serif; margin: 14px 0 8px 0; }
</style>"""

_ASPAS = str.maketrans({"“": '"', "”": '"', "‘": "'", "’": "'"})


class ErroModelo(Exception):
    pass


class Modelo:
    """Corpo HTML do modelo, separado em estilos (head) e conteúdo (body)."""

    def __init__(self, caminho: Path, estilos: str, corpo: str, avisos: list[str]):
        self.caminho = caminho
        self.estilos = estilos
        self.avisos = avisos
        fonte = _limpar_marcadores(corpo)
        env = _ambiente_html()
        try:
            self._tpl = env.from_string(fonte)
            self.campos = jinja2.meta.find_undeclared_variables(env.parse(fonte))
        except jinja2.TemplateSyntaxError as e:
            raise ErroModelo(f"{caminho.name}: marcador mal formado ({e.message})") from None
        texto = re.sub(r"<[^>]+>", "", fonte)
        if texto.count("{{") != texto.count("}}"):
            self.avisos.append("há '{{' sem '}}' correspondente — algum marcador ficou quebrado")

    def preencher(self, dados: dict) -> str:
        try:
            return self._tpl.render(**dados)
        except jinja2.TemplateError as e:
            raise ErroModelo(f"{self.caminho.name}: {_traduz(e)}") from None


def _ambiente_html():
    return jinja2.Environment(autoescape=True, undefined=jinja2.StrictUndefined)


def _traduz(e: Exception) -> str:
    m = re.search(r"'(\w+)' is undefined", str(e))
    if m:
        return f"o campo {{{{{m.group(1)}}}}} não existe no config.yaml"
    return str(e)


def _limpar_marcadores(conteudo: str) -> str:
    """O Word pode quebrar {{ campo }} em trechos com formatação diferente
    ou trocar aspas por aspas curvas. Remove tags e corrige dentro dos marcadores."""

    def limpa(m):
        abre, meio, fecha = m.group(1), m.group(2), m.group(3)
        meio = re.sub(r"<[^>]+>", "", meio)
        meio = html.unescape(meio).translate(_ASPAS)
        return f"{abre}{meio}{fecha}"

    return re.sub(r"(\{\{|\{%)(.*?)(\}\}|%\})", limpa, conteudo, flags=re.S)


def _ler_html(caminho: Path) -> str:
    bruto = caminho.read_bytes()
    m = re.search(rb'charset=["\']?([\w-]+)', bruto[:3000], re.I)
    for cod in ([m.group(1).decode()] if m else []) + ["utf-8", "cp1252"]:
        try:
            return bruto.decode(cod)
        except (UnicodeDecodeError, LookupError):
            continue
    return bruto.decode("utf-8", errors="replace")


def carregar_modelo(caminho: Path) -> Modelo:
    if not caminho.exists():
        raise ErroModelo(f"Modelo não encontrado: {caminho}")
    sufixo = caminho.suffix.lower()

    if sufixo == ".docx":
        with caminho.open("rb") as f:
            res = mammoth.convert_to_html(f)
        avisos = [m.message for m in res.messages if m.type == "warning"]
        return Modelo(caminho, ESTILO_BASE, res.value, avisos)

    if sufixo in (".htm", ".html"):
        doc = _ler_html(caminho)
        head = re.search(r"<head[^>]*>(.*?)</head>", doc, re.S | re.I)
        estilos = "".join(re.findall(r"<style.*?</style>", head.group(1), re.S | re.I)) if head else ""
        body = re.search(r"<body[^>]*>(.*)</body>", doc, re.S | re.I)
        return Modelo(caminho, estilos or ESTILO_BASE, body.group(1) if body else doc, [])

    raise ErroModelo(f"Formato de modelo não suportado: {caminho.name} (use .docx ou .htm)")


def preencher_texto(texto: str, dados: dict) -> str:
    """Para o assunto: sem escape HTML."""
    env = jinja2.Environment(autoescape=False, undefined=jinja2.StrictUndefined)
    try:
        return " ".join(env.from_string(texto).render(**dados).split())
    except jinja2.TemplateError as e:
        raise ErroModelo(f"Config ({texto!r}): {_traduz(e)}") from None


def montar_documento(estilos: str, corpo: str) -> str:
    return (
        '<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">'
        f"{estilos}</head><body>{corpo}</body></html>"
    )


