"""Destinos dos e-mails montados.

Cada remetente expõe o mesmo método `entregar(email) -> (status, detalhe)`.
Para migrar para o Gmail no futuro, basta criar uma classe nova com esse
método; leitura das planilhas e montagem dos e-mails não mudam.
"""
from __future__ import annotations

import base64
import html
import re
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from modelos import montar_documento


@dataclass
class Email:
    chave: str
    para: list[str]
    cc: list[str]
    assunto: str
    estilos: str
    corpo: str
    info: dict = field(default_factory=dict)  # linha, aba, criticidade...

    @property
    def documento(self) -> str:
        return montar_documento(self.estilos, self.corpo)


# --------------------------------------------------------------------- prévia
class RemetentePrevia:
    """Grava cada e-mail como HTML numa pasta, com um índice para revisão."""

    def __init__(self, pasta: Path):
        self.pasta = pasta
        self.pasta.mkdir(parents=True, exist_ok=True)
        for antigo in self.pasta.glob("*.html"):
            antigo.unlink()

    def entregar(self, email: Email) -> tuple[str, str]:
        nome = _nome_arquivo(email)
        cab = (
            '<div style="font:12px Arial;background:#f3f3f3;border:1px solid #ddd;'
            'padding:10px;margin-bottom:16px">'
            f"<b>Para:</b> {html.escape('; '.join(email.para))}<br>"
            f"<b>Cc:</b> {html.escape('; '.join(email.cc)) or '—'}<br>"
            f"<b>Assunto:</b> {html.escape(email.assunto)}</div>"
        )
        (self.pasta / nome).write_text(
            montar_documento(email.estilos, cab + email.corpo), encoding="utf-8"
        )
        email.info["previa"] = nome
        return "PREVIA", nome

    def finalizar(self, itens: list[dict]) -> Path:
        ordem = {"ERRO": 0, "RESSALVA": 1, "PREVIA": 2, "IGNORADO": 3}
        itens = sorted(itens, key=lambda i: (ordem.get(i["status"], 9), i.get("linha", 0)))
        linhas = []
        cores = {"PREVIA": "#1a7f37", "RESSALVA": "#b26a00", "IGNORADO": "#8a8a8a", "ERRO": "#c62828"}
        for it in itens:
            link = (
                f'<a href="{html.escape(it["previa"])}" target="_blank">abrir</a>'
                if it.get("previa") else ""
            )
            linhas.append(
                "<tr>"
                f"<td>{html.escape(str(it.get('aba', '')))}:{it.get('linha', '')}</td>"
                f"<td>{html.escape(it.get('chave', ''))}</td>"
                f"<td>{html.escape(it.get('criticidade', ''))}</td>"
                f"<td>{html.escape(it.get('para', ''))}</td>"
                f"<td>{html.escape(it.get('assunto', ''))}</td>"
                f"<td style=\"color:{cores.get(it['status'], '#000')};font-weight:600\">{it['status']}</td>"
                f"<td>{html.escape(it.get('detalhe', ''))}</td>"
                f"<td>{link}</td></tr>"
            )
        total = len(itens)
        ok = sum(1 for i in itens if i["status"] == "PREVIA")
        ress = sum(1 for i in itens if i["status"] == "RESSALVA")
        indice = f"""<!doctype html><html><head><meta charset="utf-8"><title>Prévia de notificações</title>
<style>body{{font:14px Arial;margin:24px;color:#222}}table{{border-collapse:collapse;width:100%}}
td,th{{border:1px solid #ddd;padding:6px 8px;text-align:left;vertical-align:top}}th{{background:#f3f3f3}}</style></head>
<body><h2>Prévia de notificações</h2><p>{ok} e-mails prontos · {ress} ressalvas para análise futura · {total - ok - ress} só controle/ignorados</p>
<table><tr><th>Linha</th><th>Chave</th><th>Criticidade</th><th>Para</th><th>Assunto</th><th>Status</th><th>Motivo</th><th></th></tr>
{''.join(linhas)}</table></body></html>"""
        destino = self.pasta / "_indice.html"
        destino.write_text(indice, encoding="utf-8")
        return destino


def _nome_arquivo(email: Email) -> str:
    base = f"{email.info.get('linha', 0):04d}_{email.chave}"
    return re.sub(r"[^\w.-]+", "_", base)[:80] + ".html"


# -------------------------------------------------------------------- outlook
_PR_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
_PR_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
_RASCUNHOS = 16  # olFolderDrafts


class RemetenteOutlook:
    """Cria rascunhos no Outlook clássico (Windows) via automação COM."""

    def __init__(self, conta: str | None, pasta: str | None, assinatura: bool):
        try:
            import win32com.client  # type: ignore
        except ImportError:
            raise RuntimeError(
                "pywin32 não está instalado ou o sistema não é Windows. "
                "Rode instalar.bat no computador com o Outlook clássico."
            ) from None
        self.app = win32com.client.Dispatch("Outlook.Application")
        ns = self.app.GetNamespace("MAPI")
        self.conta = None
        if conta:
            for c in ns.Accounts:
                if c.SmtpAddress.lower() == conta.lower():
                    self.conta = c
                    break
            if self.conta is None:
                contas = ", ".join(c.SmtpAddress for c in ns.Accounts)
                raise RuntimeError(f"Conta '{conta}' não encontrada no Outlook. Contas: {contas}")
        raiz = (
            self.conta.DeliveryStore.GetDefaultFolder(_RASCUNHOS)
            if self.conta else ns.GetDefaultFolder(_RASCUNHOS)
        )
        self.pasta = None
        if pasta:
            try:
                self.pasta = raiz.Folders[pasta]
            except Exception:
                self.pasta = raiz.Folders.Add(pasta)
        self.assinatura = assinatura
        self._tmp = Path(tempfile.mkdtemp(prefix="notificador_img_"))

    def entregar(self, email: Email) -> tuple[str, str]:
        mail = self.app.CreateItem(0)
        if self.conta:
            try:
                mail.SendUsingAccount = self.conta
            except Exception:  # contorno conhecido do pywin32
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, self.conta))

        for end in email.para:
            mail.Recipients.Add(end).Type = 1
        for end in email.cc:
            mail.Recipients.Add(end).Type = 2
        nao_resolvidos = []
        if not mail.Recipients.ResolveAll():
            nao_resolvidos = [r.Name for r in mail.Recipients if not r.Resolved]

        mail.Subject = email.assunto
        corpo = self._embutir_imagens(mail, email.corpo)

        if self.assinatura:
            mail.GetInspector  # força o Outlook a inserir a assinatura padrão
            mail.HTMLBody = _mesclar(mail.HTMLBody, email.estilos, corpo)
        else:
            mail.HTMLBody = montar_documento(email.estilos, corpo)

        mail.Save()
        if self.pasta is not None:
            mail.Move(self.pasta)

        if nao_resolvidos:
            return "RASCUNHO_COM_ALERTA", "destinatário não resolvido: " + "; ".join(nao_resolvidos)
        return "RASCUNHO_CRIADO", ""

    def _embutir_imagens(self, mail, corpo: str) -> str:
        """Imagens do Word chegam como data:URI; o Outlook precisa delas anexadas (cid)."""
        contador = [0]

        def troca(m):
            contador[0] += 1
            tipo, dados = m.group(1), m.group(2)
            ext = tipo.split("/")[-1].replace("jpeg", "jpg").split("+")[0]
            cid = f"img{contador[0]:02d}@notificador"
            arq = self._tmp / f"img{contador[0]:02d}.{ext}"
            arq.write_bytes(base64.b64decode(dados))
            anexo = mail.Attachments.Add(str(arq), 1, 0)
            anexo.PropertyAccessor.SetProperty(_PR_CONTENT_ID, cid)
            anexo.PropertyAccessor.SetProperty(_PR_HIDDEN, True)
            return f'src="cid:{cid}"'

        return re.sub(r'src="data:(image/[\w.+-]+);base64,([^"]+)"', troca, corpo)

    def finalizar(self, itens: list[dict]) -> None:
        return None


def _mesclar(doc_assinatura: str, estilos: str, corpo: str) -> str:
    """Coloca o corpo antes da assinatura que o Outlook inseriu."""
    if not re.search(r"<body[^>]*>", doc_assinatura, re.I):
        return montar_documento(estilos, corpo)
    doc = re.sub(r"</head>", estilos + "</head>", doc_assinatura, count=1, flags=re.I)
    return re.sub(r"(<body[^>]*>)", lambda m: m.group(1) + corpo, doc, count=1, flags=re.I)
