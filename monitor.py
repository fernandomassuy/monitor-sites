"""
Monitor de disponibilidade -> alerta no Microsoft Teams (Workflows webhook).

Variáveis de ambiente:
  SITES              "https://www.vw.com.br/pt.html,https://ofertas.vw.com.br/"
  TEAMS_WEBHOOK_URL  "https://default41eb501af6714ce0a5bfb64168c370.5f.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/df1d8cf592414a6383d47a511b944bc7/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=zVg0baI1tyeMf64x3t9kZYnXNz-W5UdxcDJcxQej8xg"
  MENTIONS           "Fernando Massuyama:fernando.massuyama@omc.com"
"""
import os
import sys
import time
import requests

TIMEOUT = 15      # segundos por tentativa
RETRIES = 2       # tentativas antes de considerar fora do ar
RETRY_WAIT = 20   # segundos entre tentativas


def check(url: str) -> str | None:
    """Retorna None se ok, ou a descrição do erro."""
    erro = None
    for _ in range(RETRIES):
        try:
            r = requests.get(url, timeout=TIMEOUT, headers={"User-Agent": "uptime-monitor"})
            if r.status_code < 400:
                return None
            erro = f"HTTP {r.status_code}"
        except requests.RequestException as e:
            erro = type(e).__name__
        time.sleep(RETRY_WAIT)
    return erro


def parse_mentions(raw: str):
    pessoas = []
    for item in filter(None, (x.strip() for x in raw.split(","))):
        nome, email = item.split(":", 1)
        pessoas.append((nome.strip(), email.strip()))
    return pessoas


def alert(webhook: str, falhas: dict, pessoas):
    linhas = "\n\n".join(f"🔴 **{u}** — {e}" for u, e in falhas.items())
    tags = " ".join(f"<at>{n}</at>" for n, _ in pessoas)
    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": [
            {"type": "TextBlock", "text": "Site(s) fora do ar", "weight": "Bolder", "size": "Large"},
            {"type": "TextBlock", "text": linhas, "wrap": True},
            {"type": "TextBlock", "text": tags, "wrap": True},
        ],
        "msteams": {
            "width": "Full",
            "entities": [
                {"type": "mention", "text": f"<at>{n}</at>", "mentioned": {"id": e, "name": n}}
                for n, e in pessoas
            ],
        },
    }
    payload = {
        "type": "message",
        "attachments": [{"contentType": "application/vnd.microsoft.card.adaptive", "content": card}],
    }
    r = requests.post(webhook, json=payload, timeout=30)
    r.raise_for_status()


def heartbeat(webhook: str, sites):
    lista = "\n\n".join(f"🟢 {u}" for u in sites)
    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": [
            {"type": "TextBlock", "text": "Monitor ativo ✅", "weight": "Bolder"},
            {"type": "TextBlock", "text": lista, "wrap": True, "isSubtle": True},
        ],
    }
    payload = {
        "type": "message",
        "attachments": [{"contentType": "application/vnd.microsoft.card.adaptive", "content": card}],
    }
    requests.post(webhook, json=payload, timeout=30).raise_for_status()


def main():
    sites = [s.strip() for s in os.environ["SITES"].split(",") if s.strip()]
    webhook = os.environ["TEAMS_WEBHOOK_URL"]
    pessoas = parse_mentions(os.environ.get("MENTIONS", ""))
    modo_heartbeat = "--heartbeat" in sys.argv

    falhas = {u: e for u in sites if (e := check(u))}
    for u in sites:
        print(f"{u}: {'FORA - ' + falhas[u] if u in falhas else 'ok'}")

    if falhas:
        alert(webhook, falhas, pessoas)
        sys.exit(1)  # marca a execução como falha no GitHub também
    elif modo_heartbeat:
        heartbeat(webhook, sites)


if __name__ == "__main__":
    main()
