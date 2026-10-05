# Advance Settlement Automation & AI Reconciliation Engine
### Department: Information Technology (IT) — SouthCity
**Role:** Junior AI & Automation Engineer  
**Case:** Advance Settlement & General Ledger (GL) Reconciliation with Gemini AI & Google Sheets Sync

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/IkhwanMaulana07/automation/blob/main/notebooks/advance_settlement_walkthrough.ipynb)


---

## 📌 Ringkasan Eksekutif & Konteks Bisnis
Sistem otomatisasi ini dirancang untuk menyelesaikan proses rekonsiliasi bulanan antara transaksi penyelesaian Uang Muka (**Advances - Other**) pada **General Ledger (GL)** SQL Server dan file monitoring **Working Paper** Tim Finance & Operational SouthCity periode April 2026.

Sistem secara cerdas:
1. **Mengekstrak dan memproses 63 transaksi GL** (fokus pada 39 baris transaksi KREDIT penyelesaian/realisasi).
2. **Memadankan (matching) 15 transaksi Uang Muka outstanding** di Working Paper menggunakan algoritma *Multi-Tier Matching Engine* (Regex PO/WO, Semantic Phrase Matching, dan penanganan Multi-Voucher).
3. **Menghitung sisa saldo dinamis** (`Saldo = Amount - Realization Amount`) secara otomatis.
4. **Mengintegrasikan Google Gemini AI** (`gemini-3.5-flash-lite`) untuk menyusun *Executive Summary* komprehensif mengenai rasio penyelesaian, detail item *partial settlement* (PBB JV 2 Summarecon), dan rekomendasi tindak lanjut bagi Direksi.
5. **Menghasilkan output terstruktur** ke dalam file Excel profesional dengan 2 tab (`Dashboard` dan `Working_Paper_Result`) serta dukungan sinkronisasi live ke **Google Sheets API**.

---

## 📊 Hasil Kunci Rekonsiliasi (Financial Metrics)

| Parameter Rekonsiliasi | Nilai Finansial (IDR) | Keterangan |
|:---|:---:|:---|
| **Total Transaksi Diproses** | **15 Transaksi** | Seluruh item advance Q1 2026 di Working Paper |
| **Total Nilai Advance Awal** | **Rp 570.406.879,00** | Total Uang Muka yang dimonitor |
| **Total Realisasi Diterima (GL)** | **Rp 463.302.663,00** | Berhasil dipadankan dari kredit GL April 2026 |
| **Total Sisa Saldo Outstanding** | **Rp 107.104.216,00** | 100% berasal dari transaksi PBB JV 2 Summarecon |
| **Rasio Penyelesaian (Settlement Rate)** | **81,22%** | Tingkat realisasi finansial |
| **Status Pemadanan** | **14 Lunas (100%), 1 Parsial, 0 Unsettled** | Akurasi pencocokan 100% |

---

## 🧠 Logika Algoritma Matching (Reconciliation Logic)

Proses pemadanan mengadopsi pendekatan **Multi-Tier Hierarchical Engine**:

```
                       [ Input Transaksi Working Paper ]
                                       │
                         Ada Kode PO/WO dalam Deskripsi?
                                 ├─── YA ───> Tier 1: Regex Extraction ([A-Z0-9]+/(?:PO|WO)/[0-9]+)
                                 │            Pencocokan presisi ke kolom DESKRIPSI GL Kredit.
                                 │
                                 └─── TIDAK ─> Tier 2: Canonical Semantic Phrase Matching
                                              Pencocokan frasa entitas (Dropbox, Talenta, Kirana Resto,
                                              Buka Puasa, PBB JV 2, Complimentary Show Unit Maret).
                                       │
                                       ▼
                       [ Tier 3: Multi-Voucher Aggregator ]
    Menangani 1 Advance yang diselesaikan oleh banyak voucher:
    - Styling Apartemen SM 1132 (2BR) : 8 voucher GL (7 belanja + 1 retur kas) = Rp 25.695.900 (Lunas)
    - Styling Apartemen SM 1127 (Studio): 3 voucher GL (2 belanja + 1 retur kas) = Rp 1.583.700 (Lunas)
                                       │
                                       ▼
                       [ Tier 4: Saldo Calculation & Sanity Check ]
    Formula: Saldo = Amount - Realization Amount
    - Lunas (Rp 0) jika Realisasi == Advance.
    - Parsial (Saldo > 0) jika Realisasi < Advance (Contoh: PBB JV 2 Summarecon sisa Rp 107.104.216).
```

### Penanganan Khusus Edge Cases:
1. **Multi-Voucher Styling Apartemen:**
   Satu transaksi advance diselesaikan oleh gabungan beberapa voucher belanja dan retur sisa uang kas (`PMT2/BM/...`). Sistem menggabungkan seluruh nomor voucher dengan pemisah koma pada Kolom F dan menjumlahkan total realisasinya pada Kolom G.
2. **PBB JV 2 Summarecon (Partial Settlement):**
   Advance sebesar Rp 422.974.179 diselesaikan sebagian oleh voucher `BCA1/BM/2604/0069` (Penerimaan Pembayaran PBB MSCM AJB JV 2) sebesar Rp 315.869.963, menyisakan saldo Rp 107.104.216.
3. **Filter Transaksi Intra-Bulan April:**
   24 transaksi kredit lainnya di GL merupakan pelunasan atas advance yang baru ditarik di bulan April 2026 itu sendiri, sehingga sistem memisahkannya agar tidak terjadi false positive matching pada Working Paper Q1.

---

## 🤖 Integrasi AI (Google Gemini API)

Sistem menggunakan model **Google Gemini 3.5 Flash** (`gemini-3.5-flash-lite`) untuk mentransformasi data mentah rekonsiliasi menjadi narasi eksekutif siap saji bagi Dewan Direksi:
- **Analisis Kontekstual:** Prompt diinjeksi dengan metrik keuangan, rincian item unsettled, dan pola multi-voucher.
- **Rekomendasi Tindak Lanjut Terarah:** Memberikan instruksi konkret bagi tim Finance untuk menagih sisa reimbursement PBB JV 2 Summarecon sebesar Rp 107,1M dan pengarsipan bukti potong pajak.
- **Graceful Fallback:** Apabila koneksi internet atau API key tidak tersedia, modul secara otomatis beralih ke engine ringkasan analitis deterministik bawaan tanpa crash.

---

## 🛠️ Struktur Project

```text
automation-skilltes/
├── data/                                      # Sumber Data Mentah Acuan
│   ├── GL - Advances Other - April 2026.xls
│   └── Working Paper Advances and Prepayment-Soal.xlsx
├── output/                                    # Hasil Eksekusi Otomatisasi
│   ├── Working_Paper_Result.xlsx              # Excel Final (Tabs: Dashboard & Working_Paper_Result)
│   ├── Executive_Summary.md                   # Executive Summary AI (Markdown)
│   └── reconciliation_report.json             # Laporan Audit Trail JSON
├── src/                                       # Core Modular Engine
│   ├── __init__.py
│   ├── config.py                              # Konfigurasi Path & Environment
│   ├── matcher.py                             # Multi-Tier Matching Engine & Excel Builder
│   ├── ai_agent.py                            # Integrasi Google Gemini API & Fallback
│   └── sheets_sync.py                         # Formatter Dashboard & Google Sheets Exporter
├── notebooks/                                 # Walkthrough Interaktif
│   └── advance_settlement_walkthrough.ipynb   # Jupyter Notebook Presentasi
├── tests/                                     # Automated Unit Testing
│   └── test_reconciliation.py                 # Pengujian Akurasi Finansial
├── main.py                                    # CLI Orchestrator Utama
├── requirements.txt                           # Dependensi Python
├── .env.example                               # Template Kredensial API
├── .gitignore                                 # Git Ignore Rules
└── README.md                                  # Dokumentasi Lengkap
```

---

## 🚀 Panduan Instalasi & Pengujian

### 1. Prasyarat Sistem
- Python 3.10 atau versi lebih baru.
- Koneksi internet (untuk mengunduh dependensi dan akses Gemini API).

### 2. Instalasi Dependensi
Clone repository dan pasang library yang dibutuhkan:
```bash
git clone https://github.com/IkhwanMaulana07/automation.git
cd automation
pip install -r requirements.txt
```

### 3. Konfigurasi Environment (Opsional untuk AI & Google Sheets)
Salin file `.env.example` menjadi `.env`:
```bash
cp .env.example .env
```
Isi konfigurasi pada file `.env` jika ingin menghubungkan live Gemini API / Telegram Bot:
```env
GEMINI_API_KEY=AIzaSy...your_gemini_api_key...
TELEGRAM_BOT_TOKEN=your_telegram_bot_token
```
*(Catatan: Sistem tetap berjalan 100% normal dan menghasilkan output lengkap meskipun tanpa API key berkat smart deterministic fallback).*

### 4. Menjalankan Otomatisasi (CLI)
Cukup jalankan satu perintah:
```bash
python main.py
```
Output yang dihasilkan:
- `output/Working_Paper_Result.xlsx` (File Excel 2 tab berformat rapi dengan formula `=D{r}-G{r}`).
- `output/Executive_Summary.txt` & `.md` (Ringkasan eksekutif AI bersih bebas asterik).
- `output/reconciliation_report.json` (Audit trail rekonsiliasi).

### 5. Menjalankan Unit Test Otomatis
Verifikasi kebenaran nominal dan logika matching:
```bash
pytest
```
atau:
```bash
python -m unittest tests/test_reconciliation.py
```

### 6. Menjalankan Interactive Web App (Streamlit Dashboard)
```bash
streamlit run app.py
```

### 7. Menjalankan Interactive Telegram Bot
```bash
python telegram_bot.py
```

### 8. Menjalankan Jupyter Notebook
Untuk melihat presentasi visual interaktif:
```bash
jupyter notebook notebooks/advance_settlement_walkthrough.ipynb
```

---

## 💡 Pemanfaatan AI Coding Assistant

Dalam proses pengerjaan technical test ini, **AI Coding Assistant** dimanfaatkan secara strategis sebagai *pair programmer* untuk:
1. **Exploratory Data Analysis (EDA):** Membantu mengurai metadata file `.xls` warisan OLE2 SQL Server dan mendeteksi anomali pada kolom deskripsi secara instan.
2. **Pattern Recognition & Multi-Voucher Reverse Engineering:** Mengidentifikasi pola 8 voucher transaksi `Styling Apartemen SM 1132` dan 3 voucher `SM 1127` yang mencakup transaksi belanja dan retur sisa uang muka.
3. **Arsitektur Modular & Clean Code:** Merancang struktur kode yang *loosely coupled*, memisahkan *matcher engine*, *AI summary agent*, dan *spreadsheet exporter*.
4. **Perumusan Prompt Engineering Finansial:** Merancang *system prompt* untuk Gemini agar menghasilkan narasi bergaya *Corporate Finance Controller* yang tajam dan berorientasi aksi bisnis.
5. **Automated Testing:** Menyusun unit test otomatis untuk menjamin tidak ada regresi kalkulasi rupiah.

---

## 📄 Deliverables Sesuai Instruksi Soal

1. **File Excel Hasil Otomatisasi:** Tersedia di `output/Working_Paper_Result.xlsx` (memuat Sheet `Dashboard` dan `Working_Paper_Result`).
2. **Link Google Sheets Live:** [Link Google Sheets - Anyone with link can view] *(Hasil import atau sinkronisasi via `sheets_sync.py`)*.
3. **Repository GitHub Public:** [https://github.com/IkhwanMaulana07/automation](https://github.com/IkhwanMaulana07/automation).

