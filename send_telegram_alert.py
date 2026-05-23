"""
Called by the workflow AFTER GitHub Pages is live.
Reads output/summary.json written by the main screener run.
"""
import json, os, requests

TOKEN   = os.environ.get("TELEGRAM_BOT_TOKEN", "")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")
URL     = os.environ.get("PAGES_URL", "")

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
SUMMARY_FILE = os.path.join(BASE_DIR, "output", "summary.json")

if not TOKEN or not CHAT_ID:
    print("Telegram credentials missing — skipping.")
    raise SystemExit(0)

if not os.path.exists(SUMMARY_FILE):
    print(f"summary.json not found at {SUMMARY_FILE}")
    raise SystemExit(1)

with open(SUMMARY_FILE) as f:
    s = json.load(f)

top_tickers = s.get("top_tickers", [])
top_line    = ("  " + "  |  ".join(top_tickers)) if top_tickers else "  —"
dl_line     = f'\n📥 <a href="{URL}">Download Report (Excel)</a>' if URL else ""

msg = (
    f"📊 <b>NSE Monthly Breakout — {s['run_month']}</b>\n"
    f"━━━━━━━━━━━━━━━━━━━━\n"
    f"✅ Breakout stocks  : <b>{s['breakout_count']}</b>\n"
    f"🔥 Full breakout    : <b>{s['full_breakout']}</b>  "
    f"(Close &amp; Low both above prev month)\n"
    f"\n<b>Top Full-Breakout Tickers:</b>\n{top_line}"
    f"{dl_line}\n"
    f"━━━━━━━━━━━━━━━━━━━━\n"
    f"<i>Universe: NSE 500  |  Source: Yahoo Finance</i>"
)

r = requests.post(
    f"https://api.telegram.org/bot{TOKEN}/sendMessage",
    json={"chat_id": CHAT_ID, "text": msg, "parse_mode": "HTML"},
    timeout=10,
)
print("Telegram sent ✓" if r.ok else f"Telegram failed: {r.status_code} {r.text}")
