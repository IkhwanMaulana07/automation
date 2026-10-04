"""
SouthCity SmartSettlement AI — Telegram Bot Integration
Allows finance managers & controllers to query advance settlement status,
download Excel reports, and converse with Gemini AI directly via Telegram.
"""
import os
import sys
import json
import asyncio
from pathlib import Path
from dotenv import load_dotenv

# Ensure root is in sys.path
BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

load_dotenv()

from src.config import (
    GL_FILE_PATH,
    WP_FILE_PATH,
    OUTPUT_EXCEL_PATH,
    GEMINI_API_KEY,
    GEMINI_MODEL
)
from src.matcher import ReconciliationMatcher

# Initialize Matcher & Cached Data
matcher = ReconciliationMatcher(GL_FILE_PATH, WP_FILE_PATH)
results = matcher.match_all()
metrics = matcher.get_summary_metrics()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()


def query_gemini(user_query: str) -> str:
    """Answers financial questions using Gemini AI with structured context."""
    if not GEMINI_API_KEY:
        return f"Total Advance: Rp {metrics['total_advance']:,.0f}, Realisasi: Rp {metrics['total_realization']:,.0f}, Sisa Saldo: Rp {metrics['total_saldo']:,.0f} (PBB JV 2 Summarecon)."

    try:
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=GEMINI_API_KEY)

        context_data = {
            "metrics": metrics,
            "items": [
                {
                    "deskripsi": r["desc"],
                    "advance": r["amount"],
                    "realisasi": r["realization_amount"],
                    "saldo": r["saldo"],
                    "status": r["status"],
                    "voucher_realisasi": r["realization_voucher"],
                    "tgl_realisasi": r["realization_date"],
                    "jumlah_voucher": r["vouchers_count"]
                }
                for r in results
            ]
        }

        system_prompt = f"""
        Anda adalah asisten AI resmi Finance Controller PT. Setiawan Dwi Tunggal (SouthCity) yang terhubung ke Telegram.
        Data rekonsiliasi Advance Settlement periode April 2026:
        {json.dumps(context_data, indent=2, default=str)}

        Instruksi Telegram:
        - Jawab dengan format rapi dan ringkas (gunakan emoji yang profesional).
        - Format angka dalam Rupiah (contoh: Rp 107.104.216).
        - Jika diminta draf penagihan/follow up, buatkan draf pesan yang formal.
        """

        res = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=f"System Context:\n{system_prompt}\n\nPertanyaan User Telegram: {user_query}",
            config=types.GenerateContentConfig(temperature=0.2)
        )
        if res and res.text:
            return res.text
    except Exception as e:
        return f"Maaf, terjadi kendala saat memproses dengan Gemini AI: {e}"

    return "Tidak ada respons dari AI."


async def start_command(update, context):
    welcome_text = (
        "💼 *SouthCity Finance AI Assistant*\n"
        "Halo! Saya bot asisten otomatisasi rekonsiliasi Uang Muka (Advance Settlement) SouthCity periode April 2026.\n\n"
        "📌 *Perintah yang Tersedia:*\n"
        "• `/status` - Lihat ringkasan metrik rekonsiliasi keuangan\n"
        "• `/summary` - Baca Executive Summary dari Gemini AI\n"
        "• `/unsettled` - Cek transaksi yang masih memiliki sisa saldo\n"
        "• `/excel` - Unduh file `Working_Paper_Result.xlsx` langsung\n\n"
        "💬 *Atau langsung ketik pertanyaan apa pun!* (Contoh: _Berapa sisa saldo Uang Muka PBB?_ atau _Apakah Styling Apartemen sudah lunas?_)"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")


async def status_command(update, context):
    msg = (
        "📊 *RINGKASAN REKONSILIASI KEUANGAN (APRIL 2026)*\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        f"• *Total Transaksi:* {metrics['total_items']} Item\n"
        f"• *Total Nilai Advance:* Rp {metrics['total_advance']:,.0f}\n"
        f"• *Total Realisasi Masuk:* Rp {metrics['total_realization']:,.0f}\n"
        f"• *Sisa Saldo Terbuka:* Rp {metrics['total_saldo']:,.0f}\n"
        f"• *Settlement Rate:* *{metrics['settlement_rate_pct']}%*\n"
        f"• *Status:* {metrics['count_settled']} Lunas | {metrics['count_partial']} Parsial | {metrics['count_unsettled']} Belum\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        "Ketik `/unsettled` untuk detail item yang belum lunas penuh."
    )
    await update.message.reply_text(msg, parse_mode="Markdown")


async def summary_command(update, context):
    summary_path = BASE_DIR / "output" / "Executive_Summary.md"
    if summary_path.exists():
        with open(summary_path, "r", encoding="utf-8") as f:
            text = f.read()
            # Trim if exceeds telegram limit (4096 chars)
            if len(text) > 4000:
                text = text[:4000] + "\n\n*(Laporan terpotong karena batas karakter Telegram. Unduh file Excel untuk versi lengkap)*"
            await update.message.reply_text(text)
    else:
        await update.message.reply_text("Executive Summary belum digenerate. Jalankan `python main.py` terlebih dahulu.")


async def unsettled_command(update, context):
    partial = metrics["partial_details"]
    if partial:
        item = partial[0]
        msg = (
            "⚠️ *ITEM DENGAN SISA SALDO / UNSETTLED*\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            f"• *Deskripsi:* {item['desc']}\n"
            f"• *No Voucher Advance:* `{item['advance_voucher']}`\n"
            f"• *Advance Awal:* Rp {item['amount']:,.0f}\n"
            f"• *Realisasi Diterima:* Rp {item['realization_amount']:,.0f} (`{item['realization_voucher']}`)\n"
            f"• *Sisa Saldo Terbuka:* *Rp {item['saldo']:,.0f}*\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "💡 *Saran Tindak Lanjut:* Hubungi tim Finance/Legal untuk penagihan sisa reimburse ke pihak mitra Summarecon."
        )
    else:
        msg = "✅ Seluruh transaksi Uang Muka telah terselesaikan 100% (Lunas)!"
    await update.message.reply_text(msg, parse_mode="Markdown")


async def excel_command(update, context):
    if not OUTPUT_EXCEL_PATH.exists():
        matcher.generate_excel_result(str(OUTPUT_EXCEL_PATH))

    await update.message.reply_text("⏳ Sedang menyiapkan file Excel...")
    with open(OUTPUT_EXCEL_PATH, "rb") as doc:
        await update.message.reply_document(
            document=doc,
            filename="Working_Paper_Result.xlsx",
            caption="📑 File Hasil Rekonsiliasi Advance Settlement (Tabs: Dashboard & Working_Paper_Result)"
        )


async def handle_message(update, context):
    user_text = update.message.text
    await update.message.reply_text("🤔 AI sedang memeriksa data keuangan...")
    response = query_gemini(user_text)
    await update.message.reply_text(response)


def main():
    if not TELEGRAM_BOT_TOKEN:
        print("=" * 70)
        print("❌ TELEGRAM_BOT_TOKEN BELUM DIISI DI FILE .env!")
        print("=" * 70)
        print("Cara mendapatkan token Telegram Bot gratis (hanya 30 detik):")
        print("1. Buka Telegram dan cari akun: @BotFather")
        print("2. Kirim pesan: /newbot")
        print("3. Beri nama bot Anda (misal: SouthCity Finance AI)")
        print("4. Beri username bot yang berakhiran bot (misal: southcity_finance_bot)")
        print("5. Salin token API yang diberikan BotFather (contoh: 123456789:ABCdef...)")
        print("6. Tempelkan ke file .env:")
        print("   TELEGRAM_BOT_TOKEN=123456789:ABCdef...")
        print("7. Jalankan kembali: python telegram_bot.py")
        print("=" * 70)
        return

    from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters

    print("🚀 Menjalankan Telegram Bot...")
    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("status", status_command))
    app.add_handler(CommandHandler("summary", summary_command))
    app.add_handler(CommandHandler("unsettled", unsettled_command))
    app.add_handler(CommandHandler("excel", excel_command))
    app.add_handler(CommandHandler("download", excel_command))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    print("✅ Bot Telegram aktif dan siap menerima pesan!")
    app.run_polling()


if __name__ == "__main__":
    main()
