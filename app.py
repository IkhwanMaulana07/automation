"""
SouthCity SmartSettlement AI — Finance Chatbot & Interactive Dashboard
Built with Streamlit and Google Gemini AI
"""
import os
import json
import streamlit as st
import pandas as pd
from pathlib import Path
from src.config import GL_FILE_PATH, WP_FILE_PATH, OUTPUT_EXCEL_PATH, GEMINI_API_KEY, GEMINI_MODEL
from src.matcher import ReconciliationMatcher
from src.ai_agent import ExecutiveSummaryAgent

st.set_page_config(
    page_title="SouthCity Finance AI Assistant",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-header {
        font-size: 26px;
        font-weight: 700;
        color: #1A365D;
        margin-bottom: 0px;
    }
    .sub-header {
        font-size: 14px;
        color: #718096;
        margin-bottom: 20px;
    }
    .metric-card {
        background-color: #F7FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_data
def get_reconciliation_data():
    matcher = ReconciliationMatcher(GL_FILE_PATH, WP_FILE_PATH)
    results = matcher.match_all()
    metrics = matcher.get_summary_metrics()
    return results, metrics


results, metrics = get_reconciliation_data()

# SIDEBAR
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/accounting.png", width=64)
    st.title("SouthCity Finance AI")
    st.caption("Advance Settlement & GL Reconciliation Engine")
    st.divider()

    st.subheader("Ringkasan Cepat")
    st.metric("Total Advance", f"Rp {metrics['total_advance']:,.0f}")
    st.metric("Total Realisasi", f"Rp {metrics['total_realization']:,.0f}")
    st.metric("Sisa Saldo", f"Rp {metrics['total_saldo']:,.0f}", delta=f"{metrics['settlement_rate_pct']}% Selesai")

    st.divider()
    if os.path.exists(OUTPUT_EXCEL_PATH):
        with open(OUTPUT_EXCEL_PATH, "rb") as f:
            st.download_button(
                label="📥 Download Working_Paper_Result.xlsx",
                data=f,
                file_name="Working_Paper_Result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# MAIN CONTENT
st.markdown('<p class="main-header">💼 SouthCity Finance & Automation Assistant</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Tanya jawab interaktif seputar status Uang Muka, Realisasi GL, dan Tindak Lanjut Saldo</p>', unsafe_allow_html=True)

tab_chat, tab_dashboard, tab_table = st.tabs(["💬 Chatbot Keuangan AI", "📊 Executive Dashboard", "📋 Tabel Rekonsiliasi"])

# TAB 1: CHATBOT
with tab_chat:
    st.markdown("##### Ajukan pertanyaan seputar data rekonsiliasi Uang Muka periode April 2026:")
    
    # Prompt suggestions
    col_p1, col_p2, col_p3 = st.columns(3)
    suggested_prompt = None
    if col_p1.button("📌 Berapa sisa saldo Uang Muka yang belum lunas?", use_container_width=True):
        suggested_prompt = "Berapa sisa saldo Uang Muka yang belum lunas dan transaksi apa saja?"
    if col_p2.button("🔍 Bagaimana status Uang Muka Styling Apartemen?", use_container_width=True):
        suggested_prompt = "Bagaimana status penyelesaian Uang Muka Styling Apartemen SM 1132 dan SM 1127?"
    if col_p3.button("✉️ Buatkan draft surat/email penagihan PBB JV 2", use_container_width=True):
        suggested_prompt = "Buatkan draft email penagihan resmi ke pihak Summarecon untuk sisa reimburse PBB JV 2 sebesar Rp 107.104.216."

    # Initialize chat history
    if "messages" not in st.session_state:
        st.session_state.messages = [
            {
                "role": "assistant",
                "content": "Halo! Saya adalah **Finance AI Assistant** SouthCity. Saya dapat membantu Anda menganalisis data rekonsiliasi Uang Muka, memeriksa status multi-voucher, menghitung sisa saldo, atau membuat draf rekomendasi tindak lanjut bagi manajemen. Ada yang bisa saya bantu?"
            }
        ]

    # Display chat messages
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # Chat Input
    user_input = st.chat_input("Ketik pertanyaan Anda di sini (misal: Apakah ada selisih di Uang Muka Dropbox?)...")
    if suggested_prompt:
        user_input = suggested_prompt

    if user_input:
        # Add user message to history
        st.session_state.messages.append({"role": "user", "content": user_input})
        with st.chat_message("user"):
            st.markdown(user_input)

        # Build context
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

        # Query Gemini API or fallback
        with st.chat_message("assistant"):
            with st.spinner("AI sedang menganalisis data keuangan..."):
                response_text = ""
                api_key = os.getenv("GEMINI_API_KEY", GEMINI_API_KEY)
                
                if api_key:
                    try:
                        from google import genai
                        from google.genai import types
                        client = genai.Client(api_key=api_key)
                        
                        system_instruction = f"""
                        Anda adalah asisten AI resmi Finance Controller PT. Setiawan Dwi Tunggal (SouthCity).
                        Anda memiliki data lengkap rekonsiliasi Advance Settlement periode April 2026 sebagai berikut:
                        {json.dumps(context_data, indent=2, default=str)}

                        Instruksi:
                        - Jawablah pertanyaan pengguna dengan sopan, akurat, berdasarkan data di atas.
                        - Format angka selalu dalam format Rupiah (contoh: Rp 107.104.216).
                        - Jika pengguna meminta draft email/surat, buatkan draf formal yang profesional untuk tim Finance/Management.
                        - Berikan alasan atau nomor voucher jika relevan.
                        """

                        gemini_res = client.models.generate_content(
                            model=GEMINI_MODEL,
                            contents=f"System Context:\n{system_instruction}\n\nPertanyaan User: {user_input}",
                            config=types.GenerateContentConfig(temperature=0.2)
                        )
                        if gemini_res and gemini_res.text:
                            response_text = gemini_res.text
                    except Exception as e:
                        response_text = f"Terjadi kendala saat memanggil Gemini API ({e}). Berdasarkan data internal: Total Advance adalah Rp {metrics['total_advance']:,.0f}, Realisasi Rp {metrics['total_realization']:,.0f}, dan Sisa Saldo Rp {metrics['total_saldo']:,.0f} (PBB JV 2 Summarecon)."
                else:
                    response_text = f"Berdasarkan data rekonsiliasi: Total Advance adalah Rp {metrics['total_advance']:,.0f}, Realisasi Rp {metrics['total_realization']:,.0f}, dan Sisa Saldo Rp {metrics['total_saldo']:,.0f} (PBB JV 2 Summarecon)."

                st.markdown(response_text)
                st.session_state.messages.append({"role": "assistant", "content": response_text})

# TAB 2: DASHBOARD
with tab_dashboard:
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Transaksi", f"{metrics['total_items']} Item")
    col2.metric("Total Advance", f"Rp {metrics['total_advance']:,.0f}")
    col3.metric("Total Realisasi", f"Rp {metrics['total_realization']:,.0f}")
    col4.metric("Sisa Saldo", f"Rp {metrics['total_saldo']:,.0f}", delta=f"{metrics['settlement_rate_pct']}% Lunas")

    st.divider()
    st.subheader("Executive Summary (Gemini AI)")
    exec_summary_file = Path(OUTPUT_EXCEL_PATH).parent / "Executive_Summary.md"
    if exec_summary_file.exists():
        with open(exec_summary_file, "r", encoding="utf-8") as f:
            st.markdown(f.read())
    else:
        st.info("Jalankan `python main.py` untuk mengenerate Executive Summary.")

# TAB 3: TABEL REKONSILIASI
with tab_table:
    st.subheader("Data Hasil Pemadanan (Working Paper Result)")
    df_display = pd.DataFrame(results)[[
        'date', 'advance_voucher', 'desc', 'amount',
        'realization_date', 'realization_voucher', 'realization_amount', 'saldo', 'status'
    ]]
    df_display.columns = [
        'Tanggal Advance', 'No Voucher Advance', 'Deskripsi Uang Muka', 'Nominal Advance',
        'Tanggal Realisasi', 'No Voucher Realisasi', 'Nominal Realisasi', 'Sisa Saldo', 'Status'
    ]
    st.dataframe(df_display, use_container_width=True)
