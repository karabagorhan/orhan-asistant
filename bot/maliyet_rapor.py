import json, os, urllib.request, urllib.parse
from pathlib import Path

def _veri():
    key = (Path.home() / ".openrouter_key").read_text().strip()
    req = urllib.request.Request("https://openrouter.ai/api/v1/key",
                                 headers={"Authorization": "Bearer " + key})
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.load(r).get("data", {})

def kullanim():
    d = _veri()
    return float(d.get("usage") or 0), d.get("usage_daily")

def ozet():
    d = _veri()
    s = ["💰 OpenRouter kullanımı"]
    for k, ad in [("usage_daily", "Bugün"), ("usage_weekly", "Bu hafta"),
                  ("usage_monthly", "Bu ay"), ("usage", "Toplam")]:
        if d.get(k) is not None:
            s.append(f"{ad}: ${d[k]:.3f}")
    if d.get("limit") is not None:
        kalan = d.get("limit_remaining")
        s.append(f"Limit: ${d['limit']:.2f}" + (f" | Kalan: ${kalan:.2f}" if kalan is not None else ""))
    else:
        s.append("⚠️ Anahtarda harcama limiti yok — OpenRouter panelinden ekleyin!")
    return "\n".join(s)

if __name__ == "__main__":
    metin = "☀️ Günlük maliyet raporu\n\n" + ozet()
    url = "https://api.telegram.org/bot" + os.environ["TG_TOKEN"] + "/sendMessage"
    veri = urllib.parse.urlencode({"chat_id": os.environ["TG_CHAT_ID"], "text": metin}).encode()
    urllib.request.urlopen(url, data=veri, timeout=20)
