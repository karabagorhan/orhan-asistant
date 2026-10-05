import asyncio, json, os, time
from pathlib import Path
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from maliyet_rapor import ozet, kullanim

TOKEN = os.environ["TG_TOKEN"]
IZINLI = int(os.environ["TG_CHAT_ID"])
ASISTAN = Path("/opt/asistan")
CIKTI, GELEN = ASISTAN / "ciktilar", ASISTAN / "gelen"
CLAUDE = os.path.expanduser("~/.local/bin/claude-ds")
NORMAL_MODEL = "openai/gpt-6-luna"
PRO_MODEL = "openai/gpt-6-luna-pro"
ZAMAN_ASIMI = 1800
OTURUM_SURESI = 1800
MAX_TUR = "40"
kilit = asyncio.Lock()
durum = {"yeni": True, "son": 0.0, "pro": False}

EK = ("\n\n[Telegram üzerinden geldi. Ürettiğin dosyaları /opt/asistan/ciktilar içine kaydet. "
      "Yanıtını kısa ve telefonda okunacak şekilde yaz. Token tasarrufu yap: gereksiz dosya okuma ve uzun açıklama yapma.]")

def yetkili(u: Update) -> bool:
    return u.effective_chat is not None and u.effective_chat.id == IZINLI

def oturum_yeni_mi() -> bool:
    return durum["yeni"] or (time.time() - durum["son"] > OTURUM_SURESI)

def guvenli_kullanim():
    try:
        return kullanim()
    except Exception as e:
        print("MALIYET OKUNAMADI:", e, flush=True)
        return (None, None)

def maliyet_satiri(u: dict, once, sonra, tur) -> str:
    girdi = (u.get("input_tokens") or 0) + (u.get("cache_creation_input_tokens") or 0)
    onbellek = u.get("cache_read_input_tokens") or 0
    cikti = u.get("output_tokens") or 0
    s = f"📊 Girdi: {girdi:,} | Önbellek: {onbellek:,} | Çıktı: {cikti:,} token"
    if tur:
        s += f" | {tur} adım"
    s = s.replace(",", ".")
    if once[0] is not None and sonra[0] is not None:
        s += f"\n💵 Bu mesaj: ${max(sonra[0] - once[0], 0):.4f}"
        if sonra[1] is not None:
            s += f" | Bugün: ${sonra[1]:.3f}"
    return s

async def calistir(prompt: str, model: str):
    yeni = oturum_yeni_mi()
    once = await asyncio.to_thread(guvenli_kullanim)
    env = dict(os.environ, ANA_MODEL=model)
    args = [CLAUDE] + ([] if yeni else ["--continue"]) + \
           ["-p", prompt + EK, "--output-format", "json", "--max-turns", MAX_TUR,
            "--dangerously-skip-permissions"]
    proc = await asyncio.create_subprocess_exec(*args, cwd=str(ASISTAN), env=env,
             stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE)
    try:
        out, err = await asyncio.wait_for(proc.communicate(), ZAMAN_ASIMI)
    except asyncio.TimeoutError:
        proc.kill()
        return "⏱ 30 dakikada bitmedi, iş durduruldu.", ""
    if err:
        print("CLAUDE UYARI:", err.decode(errors="ignore")[-800:], flush=True)
    durum["yeni"] = False
    durum["son"] = time.time()
    ham = out.decode(errors="ignore").strip()
    try:
        j = json.loads(ham)
        metin = (j.get("result") or "").strip()
        u = j.get("usage") or {}
        tur = j.get("num_turns")
    except Exception:
        metin, u, tur = ham, {}, None
    await asyncio.sleep(3)  # OpenRouter kaydının işlenmesi için kısa bekleme
    sonra = await asyncio.to_thread(guvenli_kullanim)
    metin = metin or ("⚠️ " + err.decode(errors="ignore").strip()[-1500:])
    return metin, maliyet_satiri(u, once, sonra, tur)

async def isle(update: Update, prompt: str):
    if kilit.locked():
        await update.message.reply_text("⏳ Önceki iş sürüyor, bitince tekrar yaz.")
        return
    async with kilit:
        model = PRO_MODEL if durum["pro"] else NORMAL_MODEL
        durum["pro"] = False
        etiket = "Pro" if model == PRO_MODEL else "Luna"
        tur = "yeni konuşma" if oturum_yeni_mi() else "devam"
        bas = time.time() - 1
        await update.message.reply_text(f"🛠 Çalışıyorum ({etiket}, {tur})...")
        cevap, alt = await calistir(prompt, model)
        for i in range(0, len(cevap), 4000):
            await update.message.reply_text(cevap[i:i+4000])
        for f in sorted(CIKTI.glob("*")):
            if f.is_file() and f.stat().st_mtime >= bas:
                with open(f, "rb") as fh:
                    await update.message.reply_document(document=fh, filename=f.name)
        if alt:
            await update.message.reply_text(alt)

async def metin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if yetkili(update):
        await isle(update, update.message.text)

async def dosya(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not yetkili(update):
        return
    m = update.message
    if m.document:
        ad, tg = m.document.file_name, await m.document.get_file()
    else:
        ad, tg = f"foto_{int(time.time())}.jpg", await m.photo[-1].get_file()
    hedef = GELEN / ad
    await tg.download_to_drive(str(hedef))
    istek = m.caption or "Bu dosyayı incele, özetle ve önemli noktaları çıkar."
    await isle(update, f"Kullanıcı dosya gönderdi: {hedef}\n\nİstek: {istek}")

async def yeni(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if yetkili(update):
        durum["yeni"] = True
        await update.message.reply_text("🆕 Yeni konuşma başlatıldı.")

async def pro(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if yetkili(update):
        durum["pro"] = True
        await update.message.reply_text("🧠 Sıradaki mesajın Luna Pro ile çalışacak (zor işler için, daha pahalı).")

async def maliyet(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if yetkili(update):
        try:
            m = await asyncio.to_thread(ozet)
        except Exception as e:
            m = f"Maliyet bilgisi alınamadı: {e}"
        await update.message.reply_text(m)

async def basla(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if yetkili(update):
        await update.message.reply_text(
            "Hazırım. Yaz ya da dosya gönder.\n"
            "/yeni → yeni konuşma\n/pro → sıradaki mesaj güçlü modelle\n/maliyet → harcama durumu")

app = Application.builder().token(TOKEN).concurrent_updates(True).build()
app.add_handler(CommandHandler("start", basla))
app.add_handler(CommandHandler("yeni", yeni))
app.add_handler(CommandHandler("pro", pro))
app.add_handler(CommandHandler("maliyet", maliyet))
app.add_handler(MessageHandler(filters.Document.ALL | filters.PHOTO, dosya))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, metin))
app.run_polling()
