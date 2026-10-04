"""
Advance Settlement Automation Engine
Main Orchestrator Script
"""
import sys
import json
import warnings
from pathlib import Path

# Suppress library deprecation and SDK warnings
warnings.filterwarnings("ignore")

# Fix Windows console encoding for UTF-8
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
from src.config import (
    GL_FILE_PATH,
    WP_FILE_PATH,
    OUTPUT_EXCEL_PATH,
    EXECUTIVE_SUMMARY_PATH,
    RECONCILIATION_REPORT_JSON,
    GEMINI_API_KEY,
    GEMINI_MODEL,
    GOOGLE_SERVICE_ACCOUNT_FILE,
    SPREADSHEET_ID
)
from src.matcher import ReconciliationMatcher
from src.ai_agent import ExecutiveSummaryAgent
from src.sheets_sync import DualSheetExporter


def main():
    print("=" * 80)
    print("🚀 SOUTHCITY IT AUTOMATION: ADVANCE SETTLEMENT RECONCILIATION ENGINE")
    print("=" * 80)

    # 1. Verification of data sources
    gl_path = Path(GL_FILE_PATH)
    wp_path = Path(WP_FILE_PATH)

    if not gl_path.exists():
        print(f"[Error] GL file not found at: {gl_path}")
        sys.exit(1)
    if not wp_path.exists():
        print(f"[Error] Working Paper file not found at: {wp_path}")
        sys.exit(1)

    print(f"[*] Reading GL Data from: {gl_path.name}")
    print(f"[*] Reading Working Paper from: {wp_path.name}")

    # 2. Execution of Matching Engine
    matcher = ReconciliationMatcher(str(gl_path), str(wp_path))
    results = matcher.match_all()
    metrics = matcher.get_summary_metrics()

    print("\n" + "=" * 80)
    print("📊 RECONCILIATION SUMMARY")
    print("=" * 80)
    print(f"Total Transactions Processed : {metrics['total_items']}")
    print(f"Total Advance Value          : Rp {metrics['total_advance']:15,.2f}")
    print(f"Total Realization Settled    : Rp {metrics['total_realization']:15,.2f}")
    print(f"Total Outstanding Saldo      : Rp {metrics['total_saldo']:15,.2f}")
    print(f"Settlement Fulfillment Rate  : {metrics['settlement_rate_pct']}%")
    print(f"Status Breakdown             : {metrics['count_settled']} Settled | {metrics['count_partial']} Partial | {metrics['count_unsettled']} Unsettled")
    print("-" * 80)

    # 3. Generate Excel Working Paper Result
    print(f"\n[*] Generating Working Paper Excel at: {OUTPUT_EXCEL_PATH.name}...")
    matcher.generate_excel_result(str(OUTPUT_EXCEL_PATH))

    # 4. Generate AI Executive Summary
    print(f"[*] Generating AI Executive Summary (Model: {GEMINI_MODEL})...")
    ai_agent = ExecutiveSummaryAgent(api_key=GEMINI_API_KEY, model_name=GEMINI_MODEL)
    executive_summary = ai_agent.generate_summary(metrics, results)

    # Save Executive Summary (Clean without asterisks)
    with open(EXECUTIVE_SUMMARY_PATH, "w", encoding="utf-8") as f:
        f.write(executive_summary)
    print(f"[✓] Saved Executive Summary to: {EXECUTIVE_SUMMARY_PATH.name}")

    summary_txt_path = EXECUTIVE_SUMMARY_PATH.with_suffix(".txt")
    with open(summary_txt_path, "w", encoding="utf-8") as f:
        f.write(executive_summary)
    print(f"[✓] Saved Clean Executive Summary to: {summary_txt_path.name}")

    # Save JSON report
    report_data = {
        "metrics": metrics,
        "results": results
    }
    with open(RECONCILIATION_REPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(report_data, f, indent=2, default=str)
    print(f"[✓] Saved Audit JSON Report to: {RECONCILIATION_REPORT_JSON.name}")

    # 5. Add Dashboard Sheet to Excel
    print("[*] Building 'Dashboard' tab with KPI Cards and AI Executive Summary...")
    exporter = DualSheetExporter(str(OUTPUT_EXCEL_PATH))
    exporter.add_dashboard_sheet(metrics, executive_summary)
    print(f"[✓] Excel completed with 2 tabs: 'Dashboard' and 'Working_Paper_Result'")

    # 6. Live Google Sheets Sync (Optional)
    if SPREADSHEET_ID and GOOGLE_SERVICE_ACCOUNT_FILE:
        print("\n[*] Synchronizing with Live Google Sheets...")
        exporter.sync_to_google_sheets(
            GOOGLE_SERVICE_ACCOUNT_FILE,
            SPREADSHEET_ID,
            results,
            metrics,
            executive_summary
        )

    print("\n" + "=" * 80)
    print("✅ RECONCILIATION PIPELINE COMPLETED SUCCESSFULLY!")
    print(f"📂 Output Folder: {OUTPUT_EXCEL_PATH.parent.resolve()}")
    print("=" * 80)


if __name__ == "__main__":
    main()
