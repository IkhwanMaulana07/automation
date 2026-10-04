"""
SouthCity SmartSettlement AI — Integrasi Bot Telegram
Bot untuk manajer keuangan mengecek status rekonsiliasi advance settlement,
download laporan Excel, kirim email, dan tanya jawab dengan Gemini AI via Telegram.
"""
import os
import re
import sys
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from dotenv import load_dotenv

# Pastikan root ada di sys.path
BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.append(str(BASE_DIR))

# Fix encoding Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

load_dotenv()

from src.config import (
    GL_FILE_PATH,
    WP_FILE_PATH,
    OUTPUT_EXCEL_PATH,
    GEMINI_API_KEY,
    GEMINI_MODEL
)
from src.matcher import ReconciliationMatcher
from src.cleaner import bersihkan_markdown_dan_asterik

# Inisialisasi Matcher & Data
matcher = ReconciliationMatcher(GL_FILE_PATH, WP_FILE_PATH)
results = matcher.match_all()
metrics = matcher.get_summary_metrics()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()

# Konfigurasi Email
EMAIL_SENDER = os.getenv("EMAIL_SENDER", "").strip()
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "").strip().replace(" ", "")
EMAIL_SMTP_HOST = os.getenv("EMAIL_SMTP_HOST", "smtp.gmail.com").strip()
EMAIL_SMTP_PORT = int(os.getenv("EMAIL_SMTP_PORT", "587"))


def bersihkan_markdown(teks: str) -> str:
    """Hapus SEMUA formatting Markdown dan asterik dari teks supaya bersih dan rapi."""
    return bersihkan_markdown_dan_asterik(teks)


def fallback_local_ai(user_query: str) -> str:
    """Mesin kecerdasan finansial lokal jika API eksternal limit/offline."""
    q = user_query.lower()

    # 1. Pertanyaan tentang PBB / Saldo / Sisa / Unsettled
    if any(k in q for k in ["pbb", "saldo", "sisa", "unsettled", "partial", "parsial", "summarecon", "kurang"]):
        partial = metrics["partial_details"][0] if metrics["partial_details"] else None
        if partial:
            return (
                "📊 Informasi Saldo & Realisasi PBB JV 2 Summarecon\n"
                "━━━━━━━━━━━━━━━━━━━━━━\n"
                "📌 Berdasarkan data rekonsiliasi Advance Settlement periode April 2026:\n\n"
                f"- Deskripsi Transaksi : {partial['desc']}\n"
                f"- Nilai Advance Awal  : Rp {partial['amount']:,.0f}\n"
                f"- Realisasi Diterima  : Rp {partial['realization_amount']:,.0f}\n"
                f"- No. Voucher Masuk   : {partial['realization_voucher']} ({partial['realization_date']})\n"
                f"- Sisa Saldo Terbuka  : Rp {partial['saldo']:,.0f}\n"
                f"- Status              : PARTIAL (Belum Lunas Penuh)\n\n"
                "💡 Rekomendasi: Koordinasi tim Legal & Finance untuk penagihan sisa reimburse ke pihak Summarecon."
            )

    # 2. Pertanyaan tentang Total / Ringkasan / Rekap / Rate
    if any(k in q for k in ["total", "ringkasan", "status", "rekap", "summary", "advance", "realisasi", "rate", "persen", "persentase"]):
        return (
            "📊 Ringkasan Rekonsiliasi Advance Settlement (April 2026)\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            f"- Total Transaksi       : {metrics['total_items']} Item\n"
            f"- Total Advance Awal    : Rp {metrics['total_advance']:,.0f}\n"
            f"- Total Realisasi Masuk : Rp {metrics['total_realization']:,.0f}\n"
            f"- Total Sisa Saldo      : Rp {metrics['total_saldo']:,.0f}\n"
            f"- Settlement Rate       : {metrics['settlement_rate_pct']}%\n"
            f"- Status Detail         : {metrics['count_settled']} Lunas | {metrics['count_partial']} Parsial | {metrics['count_unsettled']} Belum\n\n"
            "Ketik /unsettled untuk rincian item atau /excel untuk mengunduh laporan lengkap."
        )

    # 3. Default informative response
    return (
        "💼 SouthCity Finance AI Assistant\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        "Asisten siap membantu Anda mengelola data rekonsiliasi keuangan SouthCity:\n\n"
        f"- Total Advance   : Rp {metrics['total_advance']:,.0f} (15 item)\n"
        f"- Realisasi Masuk : Rp {metrics['total_realization']:,.0f}\n"
        f"- Sisa Saldo      : Rp {metrics['total_saldo']:,.0f} (PBB JV 2 Summarecon)\n"
        f"- Tingkat Sukses  : {metrics['settlement_rate_pct']}%\n\n"
        "Perintah yang dapat digunakan:\n"
        "  /status    - Ringkasan metrik\n"
        "  /unsettled - Detail sisa saldo\n"
        "  /excel     - Unduh file Excel\n"
        "  /email     - Kirim laporan ke email"
    )


def query_gemini(user_query: str) -> str:
    """Jawab pertanyaan menggunakan Gemini AI dengan fallback multi-model & local engine tanpa limit."""
    if not GEMINI_API_KEY:
        return fallback_local_ai(user_query)

    MODELS_TO_TRY = ["gemini-3.5-flash-lite", "gemini-flash-lite-latest", "gemini-2.5-flash"]

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

        system_prompt = f"""Anda adalah asisten AI cerdas bernama SouthCity Finance AI Assistant.

Anda memiliki 2 kemampuan:

KEMAMPUAN 1 - AHLI KEUANGAN:
Anda memiliki data rekonsiliasi Advance Settlement PT. Setiawan Dwi Tunggal (SouthCity) periode April 2026:
{json.dumps(context_data, indent=2, default=str)}
Jika pertanyaan user berkaitan dengan data keuangan, rekonsiliasi, advance, settlement, saldo, voucher, atau topik keuangan SouthCity, jawab berdasarkan data di atas.

KEMAMPUAN 2 - ASISTEN UMUM:
Jika pertanyaan user TIDAK berkaitan dengan data keuangan SouthCity (misalnya pertanyaan umum, teknologi, sains, sejarah, tips, dll), jawab dengan pengetahuan umum Anda secara informatif dan membantu.

ATURAN FORMAT WAJIB (SANGAT PENTING):
1. DILARANG KERAS menggunakan format Markdown apapun.
2. DILARANG menggunakan tanda bintang (*) untuk bold atau italic.
3. DILARANG menggunakan tanda pagar (#) untuk header.
4. DILARANG menggunakan backtick (`) untuk code.
5. DILARANG menggunakan underscore (_) untuk italic.
6. Gunakan HANYA teks biasa (plain text).
7. Gunakan emoji untuk penanda visual (contoh: 📊 📌 ✅ ⚠️ 💰).
8. Gunakan tanda strip (-) untuk bullet point.
9. Gunakan garis pemisah (━━━) untuk pembatas bagian.
10. Format angka Rupiah dengan titik pemisah ribuan (contoh: Rp 107.104.216).
11. Jawab dalam Bahasa Indonesia kecuali user bertanya dalam bahasa lain."""

        for model_name in MODELS_TO_TRY:
            try:
                res = client.models.generate_content(
                    model=model_name,
                    contents=f"System Context:\n{system_prompt}\n\nPertanyaan User: {user_query}",
                    config=types.GenerateContentConfig(temperature=0.3, max_output_tokens=8192)
                )
                if res and res.text:
                    return bersihkan_markdown(res.text)
            except Exception:
                # Coba model berikutnya jika model ini terkena limit (429)
                continue

    except Exception:
        pass

    # Jika semua model Gemini offline/limit kuota, gunakan kecerdasan lokal tanpa limit
    return fallback_local_ai(user_query)


def kirim_email(daftar_email: list) -> str:
    """Kirim laporan rekonsiliasi (Executive Summary + Excel) ke satu atau banyak email."""
    if not EMAIL_SENDER or not EMAIL_PASSWORD:
        return (
            "Email belum dikonfigurasi.\n\n"
            "Tambahkan ke file .env:\n"
            "  EMAIL_SENDER=emailkamu@gmail.com\n"
            "  EMAIL_PASSWORD=app_password_16_digit\n\n"
            "Untuk Gmail, buat App Password di:\n"
            "https://myaccount.google.com/apppasswords"
        )

    berhasil = []
    gagal = []

    for tujuan in daftar_email:
        try:
            # Buat email
            msg = MIMEMultipart()
            msg['From'] = EMAIL_SENDER
            msg['To'] = tujuan
            msg['Subject'] = 'Laporan Rekonsiliasi Advance Settlement - April 2026 (SouthCity)'

            # Isi body email
            body = (
                "Yth. Bapak/Ibu,\n\n"
                "Berikut terlampir laporan hasil rekonsiliasi otomatis Advance Settlement "
                "PT. Setiawan Dwi Tunggal (SouthCity) periode April 2026.\n\n"
                "RINGKASAN:\n"
                f"  Total Transaksi       : {metrics['total_items']} Item\n"
                f"  Total Nilai Advance   : Rp {metrics['total_advance']:,.0f}\n"
                f"  Total Realisasi Masuk : Rp {metrics['total_realization']:,.0f}\n"
                f"  Sisa Saldo Terbuka    : Rp {metrics['total_saldo']:,.0f}\n"
                f"  Settlement Rate       : {metrics['settlement_rate_pct']}%\n"
                f"  Status                : {metrics['count_settled']} Lunas | "
                f"{metrics['count_partial']} Parsial | {metrics['count_unsettled']} Belum\n\n"
                "Detail lengkap dapat dilihat pada file Excel terlampir.\n\n"
                "Hormat kami,\n"
                "SouthCity Finance AI Assistant\n"
                "(Laporan ini digenerate secara otomatis)"
            )
            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # Lampirkan file Excel jika ada
            excel_path = OUTPUT_EXCEL_PATH
            if not excel_path.exists():
                matcher.generate_excel_result(str(excel_path))

            if excel_path.exists():
                with open(excel_path, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename="Working_Paper_Result.xlsx"'
                    )
                    msg.attach(part)

            # Lampirkan Executive Summary (bersih tanpa asterik)
            summary_path = BASE_DIR / "output" / "Executive_Summary.md"
            if summary_path.exists():
                with open(summary_path, 'r', encoding='utf-8') as f:
                    clean_summary = bersihkan_markdown_dan_asterik(f.read())
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(clean_summary.encode('utf-8'))
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="Executive_Summary.txt"'
                )
                msg.attach(part)

            # Kirim email via SMTP
            with smtplib.SMTP(EMAIL_SMTP_HOST, EMAIL_SMTP_PORT) as server:
                server.starttls()
                server.login(EMAIL_SENDER, EMAIL_PASSWORD)
                server.send_message(msg)

            berhasil.append(tujuan)

        except smtplib.SMTPAuthenticationError:
            return (
                "Gagal login email. Pastikan:\n"
                "  1. EMAIL_SENDER di .env sudah benar\n"
                "  2. EMAIL_PASSWORD menggunakan App Password (bukan password biasa)\n"
                "  3. Untuk Gmail, buat App Password di: https://myaccount.google.com/apppasswords"
            )
        except Exception as e:
            gagal.append(f"{tujuan} ({e})")

    # Buat laporan hasil kirim
    laporan = ""
    if berhasil:
        laporan += f"✅ Berhasil dikirim ke {len(berhasil)} email:\n"
        for email in berhasil:
            laporan += f"  - {email}\n"
    if gagal:
        laporan += f"\n❌ Gagal dikirim ke {len(gagal)} email:\n"
        for info in gagal:
            laporan += f"  - {info}\n"
    if berhasil:
        laporan += "\nLampiran:\n  - Working_Paper_Result.xlsx\n  - Executive_Summary.txt\n"
        laporan += "\nSilakan cek inbox atau folder spam email tujuan."

    return laporan


async def start_command(update, context):
    welcome_text = (
        "💼 SouthCity Finance AI Assistant\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        "Halo! Saya bot asisten cerdas SouthCity.\n"
        "Saya bisa menjawab pertanyaan seputar rekonsiliasi keuangan "
        "maupun pertanyaan umum lainnya.\n\n"
        "📌 Perintah yang Tersedia:\n"
        "  /status     - Ringkasan metrik rekonsiliasi keuangan\n"
        "  /summary    - Executive Summary dari Gemini AI\n"
        "  /unsettled  - Transaksi yang masih ada sisa saldo\n"
        "  /excel      - Unduh file Working_Paper_Result.xlsx\n"
        "  /email      - Kirim laporan ke satu atau banyak email sekaligus\n"
        "                (contoh: /email tim1@southcity.co.id tim2@southcity.co.id)\n\n"
        "💬 Atau langsung ketik pertanyaan apa pun!\n"
        "Contoh keuangan : Berapa sisa saldo uang muka PBB?\n"
        "Contoh umum     : Apa itu rekonsiliasi dalam akuntansi?"
    )
    await update.message.reply_text(welcome_text)


async def status_command(update, context):
    msg = (
        "📊 RINGKASAN REKONSILIASI KEUANGAN (APRIL 2026)\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        f"  Total Transaksi       : {metrics['total_items']} Item\n"
        f"  Total Nilai Advance   : Rp {metrics['total_advance']:,.0f}\n"
        f"  Total Realisasi Masuk : Rp {metrics['total_realization']:,.0f}\n"
        f"  Sisa Saldo Terbuka    : Rp {metrics['total_saldo']:,.0f}\n"
        f"  Settlement Rate       : {metrics['settlement_rate_pct']}%\n"
        f"  Status                : {metrics['count_settled']} Lunas | "
        f"{metrics['count_partial']} Parsial | {metrics['count_unsettled']} Belum\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        "Ketik /unsettled untuk detail item yang belum lunas penuh."
    )
    await update.message.reply_text(msg)


async def kirim_pesan_aman(update, teks: str):
    """Kirim pesan tanpa batasan karakter Telegram (otomatis dipecah berurutan jika panjang)."""
    MAX_LEN = 3900
    if len(teks) <= MAX_LEN:
        await update.message.reply_text(teks)
        return

    # Bagi teks berdasarkan baris agar pemotongan tetap rapi dan tidak memotong kata
    baris = teks.split('\n')
    buffer = ""
    for line in baris:
        if len(buffer) + len(line) + 1 > MAX_LEN:
            if buffer.strip():
                await update.message.reply_text(buffer.strip())
            buffer = line + "\n"
        else:
            buffer += line + "\n"
    if buffer.strip():
        await update.message.reply_text(buffer.strip())


async def summary_command(update, context):
    summary_path = BASE_DIR / "output" / "Executive_Summary.md"
    if summary_path.exists():
        with open(summary_path, "r", encoding="utf-8") as f:
            text = f.read()
            text = bersihkan_markdown(text)
            await kirim_pesan_aman(update, text)
    else:
        await update.message.reply_text(
            "Executive Summary belum digenerate. Jalankan python main.py terlebih dahulu."
        )


async def unsettled_command(update, context):
    partial = metrics["partial_details"]
    if partial:
        item = partial[0]
        msg = (
            "⚠️ ITEM DENGAN SISA SALDO / UNSETTLED\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            f"  Deskripsi           : {item['desc']}\n"
            f"  No Voucher Advance  : {item['advance_voucher']}\n"
            f"  Advance Awal        : Rp {item['amount']:,.0f}\n"
            f"  Realisasi Diterima  : Rp {item['realization_amount']:,.0f} ({item['realization_voucher']})\n"
            f"  Sisa Saldo Terbuka  : Rp {item['saldo']:,.0f}\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "💡 Saran: Hubungi tim Finance/Legal untuk penagihan sisa "
            "reimburse ke pihak mitra Summarecon."
        )
    else:
        msg = "✅ Seluruh transaksi Uang Muka telah terselesaikan 100% (Lunas)!"
    await update.message.reply_text(msg)


async def excel_command(update, context):
    if not OUTPUT_EXCEL_PATH.exists():
        matcher.generate_excel_result(str(OUTPUT_EXCEL_PATH))

    await update.message.reply_text("⏳ Sedang menyiapkan file Excel...")
    with open(OUTPUT_EXCEL_PATH, "rb") as doc:
        await update.message.reply_document(
            document=doc,
            filename="Working_Paper_Result.xlsx",
            caption="📑 File Hasil Rekonsiliasi Advance Settlement"
        )


async def email_command(update, context):
    """Kirim laporan ke satu atau banyak email. Format: /email email1@gmail.com email2@gmail.com"""
    if not context.args:
        await update.message.reply_text(
            "📧 Cara pakai: /email alamat1@email.com alamat2@email.com\n\n"
            "Bisa kirim ke 1 email atau banyak email sekaligus (pisahkan dengan spasi atau koma):\n"
            "Contoh 1 email: /email finance@southcity.co.id\n"
            "Contoh banyak: /email bos@southcity.co.id finance@southcity.co.id audit@southcity.co.id\n\n"
            "Laporan Excel dan Executive Summary akan otomatis dilampirkan."
        )
        return

    raw_text = " ".join(context.args)
    import re
    # Pisahkan email berdasarkan spasi, koma, titik koma
    email_list = [e.strip() for e in re.split(r'[,;\s]+', raw_text) if e.strip()]

    valid_emails = []
    invalid_emails = []
    for e in email_list:
        if "@" in e and "." in e:
            valid_emails.append(e)
        else:
            invalid_emails.append(e)

    if not valid_emails:
        await update.message.reply_text(
            "Format email tidak valid!\n"
            "Contoh yang benar:\n"
            "/email finance@southcity.co.id\n"
            "Atau banyak:\n"
            "/email tim1@southcity.co.id tim2@southcity.co.id"
        )
        return

    if invalid_emails:
        await update.message.reply_text(
            f"Perhatian: Email berikut diabaikan karena format tidak valid:\n"
            + "\n".join([f"- {inv}" for inv in invalid_emails])
        )

    daftar_tujuan = ", ".join(valid_emails)
    await update.message.reply_text(f"Sedang memproses dan mengirim laporan ke:\n{daftar_tujuan}...")

    # Jalankan pengiriman email
    hasil = kirim_email(valid_emails)
    await update.message.reply_text(hasil)


async def handle_message(update, context):
    user_text = update.message.text.strip()

    # Intersepsi cerdas: jika pesan adalah perintah /email atau permintaan kirim email
    import re
    emails = re.findall(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', user_text)
    if user_text.lower().startswith("/email") or (emails and any(k in user_text.lower() for k in ["email", "kirim", "send"])):
        if emails:
            daftar = ", ".join(emails)
            await update.message.reply_text(f"Sedang memproses dan mengirim laporan ke:\n{daftar}...")
            hasil = kirim_email(emails)
            await update.message.reply_text(hasil)
            return
        else:
            await update.message.reply_text(
                "Alamat email tidak ditemukan.\nContoh penggunaan:\n/email nama@gmail.com"
            )
            return

    await update.message.reply_text("🤔 AI sedang memproses pertanyaan Anda...")
    response = query_gemini(user_text)
    response = bersihkan_markdown(response)
    await kirim_pesan_aman(update, response)


def main():
    if not TELEGRAM_BOT_TOKEN:
        print("=" * 70)
        print("TELEGRAM_BOT_TOKEN BELUM DIISI DI FILE .env!")
        print("=" * 70)
        print("Cara mendapatkan token Telegram Bot (hanya 30 detik):")
        print("1. Buka Telegram dan cari akun: @BotFather")
        print("2. Kirim pesan: /newbot")
        print("3. Beri nama bot (misal: SouthCity Finance AI)")
        print("4. Beri username bot berakhiran bot (misal: southcity_finance_bot)")
        print("5. Salin token API dari BotFather")
        print("6. Tempelkan ke file .env: TELEGRAM_BOT_TOKEN=token_anda")
        print("7. Jalankan kembali: python telegram_bot.py")
        print("=" * 70)
        return

    from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters

    print("Menjalankan Telegram Bot...")
    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("status", status_command))
    app.add_handler(CommandHandler("summary", summary_command))
    app.add_handler(CommandHandler("unsettled", unsettled_command))
    app.add_handler(CommandHandler("excel", excel_command))
    app.add_handler(CommandHandler("download", excel_command))
    app.add_handler(CommandHandler("email", email_command))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_message))

    print("Bot Telegram aktif dan siap menerima pesan!")
    app.run_polling()


if __name__ == "__main__":
    main()
