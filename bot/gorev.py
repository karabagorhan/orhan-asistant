import json, os, subprocess, sys, time
from pathlib import Path
import httpx

EV = Path.home()
for satir in (EV / ".telegram_env").read_text().splitlines():
    if "=" in satir:
        k, v = satir.split("=", 1)
        os.environ[k] = v
API = "https://api.telegram.org/bot" + os.environ["TG_TOKEN"]
CHAT = os.environ["TG_CHAT_ID"]

ad = sys.argv[1]
istem = (Path("/opt/asistan/gorevler") / f"{ad}.md").read_text()
bas = time.time() - 1
env = dict(os.environ, PATH=f"{EV}/.npm-global/bin:{EV}/.local/bin:/usr/local/bin:/usr/bin:/bin")
r = subprocess.run(
    [str(EV / ".local/bin/claude-ds"), "-p",
     istem + "\n\n[Zamanlanmış görev. Dosyaları /opt/asistan/ciktilar içine kaydet. Kısa rapor yaz.]",
     "--output-format", "json", "--max-turns", "30", "--dangerously-skip-permissions"],
    cwd="/opt/asistan", env=env, capture_output=True, text=True, timeout=1800)
try:
    metin = json.loads(r.stdout).get("result", "")
except Exception:
    metin = r.stdout or r.stderr[-1500:]
mesaj = f"⏰ Görev: {ad}\n\n{metin or '(boş sonuç)'}"
for i in range(0, len(mesaj), 4000):
    httpx.post(API + "/sendMessage", data={"chat_id": CHAT, "text": mesaj[i:i+4000]}, timeout=30)
for f in sorted(Path("/opt/asistan/ciktilar").glob("*")):
    if f.is_file() and f.stat().st_mtime >= bas:
        with open(f, "rb") as fh:
            httpx.post(API + "/sendDocument", data={"chat_id": CHAT},
                       files={"document": (f.name, fh)}, timeout=60)
