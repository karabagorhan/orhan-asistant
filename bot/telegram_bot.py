import asyncio, os, time
from pathlib import Path
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

TOKEN = os.environ["TG_TOKEN"]
IZINLI = int(os.environ["TG_CHAT_ID"])
ASISTAN = Path("/opt/asistan")
CIKTI, GELEN = ASISTAN / "ciktilar", ASISTAN / "gelen"
CLAUDE = os.path.expanduser("~/.local/bin/claude-ds")
ZAMAN_ASIMI = 1800
kilit = asyncio.Lock()
durum = {"yeni": True}

EK = ("\n\n[Telegram üzerinden geldi. Ürettiğin dosyaları /opt/asistan/ciktilar içine kaydet. "
      "Yanıtını kısa ve telefonda okunacak şekilde yaz.]")

def yetkili(u: Update) -> bool:
    return u.effective_chat is not None and u.effective_chat.id == IZINLI

async def calistir(prompt: str) -> str:
    args = [CLAUDE] + ([] if durum["yeni"] else ["--continue"]) + \
           ["-p", prompt + EK, "--output-format", "text", "--dangerously-skip-permissions"]
    proc = await asyncio.create_subprocess_exec(*args, cwd=str(ASISTAN),
             stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE)
    try:
        out, err = await asyncio.wait_for(proc.communicate(), ZAMAN_ASIMI)
    except asyncio.TimeoutError:
        proc.kill()
        return "⏱ 30 dakikada bitmedi, iş durduruldu."
    if err: print("CLAUDE UYARI:", err.decode(errors="ignore")[-800:], flush=True)
    durum["yeni"] = False
    return out.decode(errors="ignore").strip() or ("⚠️ " + err.decode(errors="ignore").strip()[-1500:])

async def isle(update: Update, prompt: str):
    if kilit.locked():
        await update.message.reply_text("⏳ Önceki iş sürüyor, bitince tekrar yaz.")
        return
    async with kilit:
        bas = time.time() - 1
        await update.message.reply_text("🛠 Çalışıyorum...")
        cevap = await calistir(prompt)
        for i in range(0, len(cevap), 4000):
            await update.message.reply_text(cevap[i:i+4000])
        for f in sorted(CIKTI.glob("*")):
            if f.is_file() and f.stat().st_mtime >= bas:
                with open(f, "rb") as fh:
                    await update.message.reply_document(document=fh, filename=f.name)

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

async def basla(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if yetkili(update):
        await update.message.reply_text("Hazırım. Yaz ya da dosya gönder (açıklamaya ne istediğini yaz). /yeni → yeni konuşma")

app = Application.builder().token(TOKEN).concurrent_updates(True).build()
app.add_handler(CommandHandler("start", basla))
app.add_handler(CommandHandler("yeni", yeni))
app.add_handler(MessageHandler(filters.Document.ALL | filters.PHOTO, dosya))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, metin))
app.run_polling()
