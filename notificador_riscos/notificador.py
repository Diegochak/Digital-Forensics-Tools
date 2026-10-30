"""Notificador de software não homologado — monta rascunhos de e-mail por achado.

Uso:
    python notificador.py mapear       mostra letras/cabeçalhos das planilhas
    python notificador.py previa       gera os e-mails em HTML para revisão (não toca no Outlook)
    python notificador.py rascunhos    cria os rascunhos no Outlook

Opções: --config arquivo.yaml  --limite N  --forcar  --sim
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import sys
import webbrowser
from pathlib import Path

import yaml

from modelos import ErroModelo, carregar_modelo, preencher_texto
from planilhas import ErroConfig, ler_planilha, mapa_colunas, normalizar
from registro import Registro
from remetentes import Email, RemetenteOutlook, RemetentePrevia

PASTA = Path(__file__).resolve().parent


# ------------------------------------------------------------------ config
def carregar_config(caminho: Path) -> dict:
    if not caminho.exists():
        raise ErroConfig(f"Arquivo de configuração não encontrado: {caminho}")
    with caminho.open(encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    for chave in ("planilha_analise", "criticidade", "email"):
        if chave not in cfg:
            raise ErroConfig(f"Seção '{chave}' ausente no {caminho.name}")
    return cfg


def _lista(texto: str) -> list[str]:
    return [p.strip() for p in texto.replace(",", ";").split(";") if p.strip()]


# ----------------------------------------------------------------- montagem
def juntar_inventario(achados: list[dict], cfg: dict, base: Path, fmt: str) -> list[str]:
    """Completa cada achado com os dados da planilha de inventário."""
    inv_cfg = cfg.get("planilha_inventario")
    if not inv_cfg or not inv_cfg.get("arquivo"):
        return []
    ligacao = inv_cfg.get("ligacao") or {}
    if not ligacao:
        raise ErroConfig("planilha_inventario.ligacao precisa indicar os campos que ligam as planilhas.")

    indice: dict[tuple, dict] = {}
    for reg in ler_planilha(inv_cfg, base, fmt):
        k = tuple(normalizar(reg.get(c)) for c in ligacao.values())
        indice.setdefault(k, reg)

    avisos = []
    for a in achados:
        for campo in inv_cfg["colunas"]:  # campo existe mesmo sem correspondência
            a.setdefault(campo, "")
        k = tuple(normalizar(a.get(c)) for c in ligacao.keys())
        inv = indice.get(k)
        if inv is None:
            a["_aviso_inventario"] = "não encontrado no inventário"
            avisos.append(a["_linha"])
            continue
        for campo, valor in inv.items():
            if campo.startswith("_"):
                continue
            if not a.get(campo):  # a análise de riscos tem prioridade
                a[campo] = valor
    return avisos


def campos_faltando(reg: dict, obrigatorios: list | None) -> list[str]:
    """Itens da lista podem ser um campo ou uma lista de alternativas (basta um)."""
    faltando = []
    for item in obrigatorios or []:
        opcoes = item if isinstance(item, list) else [item]
        if not any(str(reg.get(c, "")).strip() for c in opcoes):
            faltando.append(" ou ".join(opcoes))
    return faltando


def gravar_ressalvas(itens: list[dict], caminho: Path) -> int:
    """Lista atual de ressalvas (sobrescrita a cada execução) para acompanhamento."""
    ressalvas = [i for i in itens if i.get("status") == "RESSALVA"]
    caminho.parent.mkdir(parents=True, exist_ok=True)
    with caminho.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["data_execucao", "aba", "linha", "chave", "criticidade", "aplicacao", "faltando"])
        agora = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
        for i in ressalvas:
            w.writerow([agora, i["aba"], i["linha"], i["chave"], i["criticidade"],
                        i.get("aplicacao", ""), i.get("faltando", "")])
    return len(ressalvas)


def passa_filtros(reg: dict, filtros: list[dict]) -> str | None:
    """Retorna o motivo da exclusão, ou None se o registro passa."""
    for f in filtros or []:
        valor = normalizar(reg.get(f["campo"], ""))
        if "vazio" in f:
            if f["vazio"] and valor:
                return f"{f['campo']} preenchido"
            if not f["vazio"] and not valor:
                return f"{f['campo']} vazio"
        if "igual" in f and valor not in {normalizar(v) for v in f["igual"]}:
            return f"{f['campo']} = '{reg.get(f['campo'])}'"
        if "contem" in f and not any(normalizar(v) in valor for v in f["contem"]):
            return f"{f['campo']} = '{reg.get(f['campo'])}'"
        if "nao_contem" in f and any(normalizar(v) in valor for v in f["nao_contem"]):
            return f"{f['campo']} = '{reg.get(f['campo'])}'"
    return None


def montar(cfg: dict, base: Path) -> list[tuple[dict, Email | None]]:
    fmt = cfg.get("geral", {}).get("formato_data", "%d/%m/%Y")
    achados = ler_planilha(cfg["planilha_analise"], base, fmt)
    juntar_inventario(achados, cfg, base, fmt)

    crit_cfg = cfg["criticidade"]
    sinonimos = {normalizar(k): normalizar(v) for k, v in (crit_cfg.get("sinonimos") or {}).items()}
    email_cfg = cfg["email"]

    modelos, assuntos = {}, {}
    for nivel, m in (crit_cfg.get("modelos") or {}).items():
        try:
            modelos[normalizar(nivel)] = carregar_modelo((base / m["arquivo"]).resolve())
        except ErroModelo as e:
            modelos[normalizar(nivel)] = e
        assuntos[normalizar(nivel)] = m.get("assunto") or email_cfg.get("assunto_padrao", "")

    notificar = {normalizar(n) for n in (crit_cfg.get("notificar") or [])}
    conhecidos = (set(modelos) | set(sinonimos) | set(sinonimos.values()) | notificar
                  | {"EMERGENCY", "CRITICAL", "HIGH", "MEDIUM", "LOW", "INFORMATIONAL"})

    campos_chave = email_cfg.get("campo_chave", [])
    campos_chave = campos_chave if isinstance(campos_chave, list) else [campos_chave]

    resultado = []
    for reg in achados:
        info = {
            "arquivo": reg["_arquivo"], "aba": reg["_aba"], "linha": reg["_linha"],
            "chave": "-".join(reg.get(c, "") for c in campos_chave) or f"L{reg['_linha']}",
        }
        bruto = reg.get(crit_cfg["campo"], "")
        nivel = normalizar(bruto)
        for parte in nivel.replace("-", " ").split():  # "4 - CRITICAL" -> CRITICAL
            if parte in conhecidos:
                nivel = parte
                break
        nivel = sinonimos.get(nivel, nivel)
        info["criticidade"] = nivel or "(vazio)"

        if notificar and nivel not in notificar:
            motivo = f"{nivel} fica só no controle (não notificável)" if nivel else "criticidade em branco"
            resultado.append(({**info, "status": "IGNORADO", "detalhe": motivo}, None))
            continue

        faltando = campos_faltando(reg, (cfg.get("ressalva") or {}).get("campos_obrigatorios"))
        if faltando:
            info.update(aplicacao=reg.get("aplicacao", ""), faltando=", ".join(faltando))
            resultado.append(({**info, "status": "RESSALVA",
                               "detalhe": "falta " + ", ".join(faltando) + " — análise futura"}, None))
            continue

        motivo = passa_filtros(reg, cfg.get("filtros"))
        if motivo:
            resultado.append(({**info, "status": "IGNORADO", "detalhe": f"filtro: {motivo}"}, None))
            continue
        modelo = modelos.get(nivel)
        if modelo is None:
            motivo = f"sem modelo para a criticidade '{bruto}'" if bruto else "criticidade em branco"
            resultado.append(({**info, "status": "IGNORADO", "detalhe": motivo}, None))
            continue
        if isinstance(modelo, ErroModelo):
            resultado.append(({**info, "status": "ERRO", "detalhe": str(modelo)}, None))
            continue

        dados = {k: v for k, v in reg.items() if not k.startswith("_")}
        dados.update(criticidade=nivel, criticidade_original=bruto,
                     hoje=dt.date.today().strftime(fmt), linha=reg["_linha"])
        try:
            para = _lista(preencher_texto(email_cfg["para"], dados))
            cc = _lista(preencher_texto(email_cfg.get("cc") or "", dados))
            assunto = preencher_texto(assuntos[nivel], dados)
            corpo = modelo.preencher(dados)
        except ErroModelo as e:
            resultado.append(({**info, "status": "ERRO", "detalhe": str(e)}, None))
            continue
        if not para:
            resultado.append(({**info, "status": "ERRO", "detalhe": "destinatário vazio"}, None))
            continue

        alertas = list(modelo.avisos)
        if reg.get("_aviso_inventario"):
            alertas.append(reg["_aviso_inventario"])
        sem_arroba = [p for p in para + cc if "@" not in p]
        if sem_arroba:
            alertas.append("será resolvido pelo catálogo do Outlook: " + "; ".join(sem_arroba))

        info.update(para="; ".join(para), cc="; ".join(cc), assunto=assunto,
                    detalhe=" | ".join(alertas))
        email = Email(info["chave"], para, cc, assunto, modelo.estilos, corpo, info)
        resultado.append((info, email))
    return resultado


# ---------------------------------------------------------------- execução
def executar(modo: str, cfg: dict, base: Path, args) -> int:
    saida = cfg.get("saida", {})
    registro = Registro((base / saida.get("log", "saida/log_notificacoes.csv")).resolve())
    itens = montar(cfg, base)

    if modo == "rascunhos" and not args.forcar:
        criados = registro.ja_criados()
        for info, email in itens:
            if email and info["chave"] in criados:
                info.update(status="IGNORADO", detalhe="rascunho já criado antes (use --forcar)")
        itens = [(i, e if i.get("status") != "IGNORADO" else None) for i, e in itens]

    prontos = [(i, e) for i, e in itens if e is not None]
    if args.limite:
        for i, e in prontos[args.limite:]:
            i.update(status="IGNORADO", detalhe=f"fora do --limite {args.limite}")
        prontos = prontos[: args.limite]

    n_ress = gravar_ressalvas([i for i, _ in itens],
                              (base / saida.get("ressalvas", "saida/ressalvas.csv")).resolve())
    print(f"\n{len(itens)} linhas lidas · {len(prontos)} e-mails a gerar · "
          f"{n_ress} ressalvas · {len(itens) - len(prontos) - n_ress} só controle/ignorados")
    ids_prontos = {id(i) for i, _ in prontos}
    for i, _ in itens:
        if id(i) not in ids_prontos:
            print(f"  [{i['status']}] linha {i['linha']} · {i['chave']} · {i.get('detalhe', '')}")

    if modo == "previa":
        remetente = RemetentePrevia((base / saida.get("pasta_previa", "saida/previa")).resolve())
    else:
        if not prontos:
            print("Nenhum rascunho a criar.")
            registro.gravar([i for i, _ in itens if i.get("status") == "RESSALVA"], modo)
            return 0
        if not args.sim:
            r = input(f"Criar {len(prontos)} rascunhos no Outlook? (s/n) ").strip().lower()
            if r not in ("s", "sim", "y"):
                print("Cancelado.")
                return 1
        ecfg = cfg["email"]
        remetente = RemetenteOutlook(ecfg.get("conta_outlook"), ecfg.get("pasta_rascunhos"),
                                     bool(ecfg.get("incluir_assinatura", True)))

    for info, email in prontos:
        try:
            status, detalhe = remetente.entregar(email)
        except Exception as e:  # um erro não interrompe o lote
            status, detalhe = "ERRO", f"{type(e).__name__}: {e}"
        info["status"] = status
        if modo == "rascunhos" or status == "ERRO":
            info["detalhe"] = " | ".join(x for x in (info.get("detalhe"), detalhe) if x)
        print(f"  [{status}] linha {info['linha']} · {info['chave']} · {info.get('para', '')}")

    todos = [i for i, _ in itens]
    if modo == "previa":
        indice = remetente.finalizar(todos)
        print(f"\nPrévia gerada: {indice}")
        if n_ress:
            print(f"Ressalvas: {(base / saida.get('ressalvas', 'saida/ressalvas.csv')).resolve()}")
        if not args.sem_abrir:
            webbrowser.open(indice.as_uri())
    else:
        registro.gravar(todos, modo)
        print(f"\nLog atualizado: {registro.caminho}")
        print("Revise os rascunhos no Outlook antes de enviar.")
    return 0


def main() -> int:
    p = argparse.ArgumentParser(description="Notificador de software não homologado")
    p.add_argument("modo", choices=["mapear", "previa", "rascunhos"])
    p.add_argument("--config", default=str(PASTA / "config.yaml"))
    p.add_argument("--limite", type=int, help="gera só os N primeiros e-mails (teste)")
    p.add_argument("--forcar", action="store_true", help="recria rascunhos já registrados no log")
    p.add_argument("--sim", action="store_true", help="não pede confirmação")
    p.add_argument("--sem-abrir", action="store_true", help="não abre a prévia no navegador")
    args = p.parse_args()

    caminho = Path(args.config).resolve()
    try:
        cfg = carregar_config(caminho)
        base = caminho.parent
        if args.modo == "mapear":
            print(mapa_colunas(cfg["planilha_analise"], base))
            if (cfg.get("planilha_inventario") or {}).get("arquivo"):
                print(mapa_colunas(cfg["planilha_inventario"], base))
            return 0
        return executar(args.modo, cfg, base, args)
    except (ErroConfig, ErroModelo, RuntimeError) as e:
        print(f"\nERRO: {e}")
        return 2


if __name__ == "__main__":
    sys.exit(main())
