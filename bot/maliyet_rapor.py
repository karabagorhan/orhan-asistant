
def kullanim():
    key = (Path.home() / ".openrouter_key").read_text().strip()
    req = urllib.request.Request("https://openrouter.ai/api/v1/key",
                                 headers={"Authorization": "Bearer " + key})
    with urllib.request.urlopen(req, timeout=15) as r:
        d = json.load(r).get("data", {})
    return float(d.get("usage") or 0), d.get("usage_daily")
