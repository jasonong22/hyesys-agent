"""
Daily US Tech Stock Alert — @JOstocks_bot
Fetches price data + analyst consensus via yfinance, generates a structured
summary, and sends it to Jason's Telegram at US market close (5 AM SGT).
Completely separate from the HyESys pipeline.
"""

import html
import json
import logging
import os
import time
import urllib.request
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, timedelta
from pathlib import Path

import yfinance as yf

# ─────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────
TELEGRAM_TOKEN   = "8912852669:AAEuo7Y0l4WzEjSRc0iyimL7CAiSL-KktVM"
TELEGRAM_CHAT_ID = "42746142"                       # @Jasonozy

TICKERS = [
    "NVDA", "NFLX", "GOOG", "MSFT", "SPOT",
    "CSPX.L",                                       # LSE listing — analysed at US close time
    "AAPL", "AMZN", "META", "TSM", "PLTR", "VOO",
]

# Display name overrides
DISPLAY_NAME = {
    "CSPX.L": "CSPX",
}

SGT = timezone(timedelta(hours=8))

# Tracks whether an alert has already been sent today
SENT_FLAG_FILE = Path(__file__).parent / ".last_sent_date"

def already_sent_today() -> bool:
    today = datetime.now(SGT).strftime("%Y-%m-%d")
    if SENT_FLAG_FILE.exists():
        return SENT_FLAG_FILE.read_text().strip() == today
    return False

def mark_sent_today():
    today = datetime.now(SGT).strftime("%Y-%m-%d")
    SENT_FLAG_FILE.write_text(today)

LOG_FILE = Path(__file__).parent / "stock_alert.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
log = logging.getLogger("stocks.alert")

# ─────────────────────────────────────────────
# TECHNICAL INDICATORS
# ─────────────────────────────────────────────
def compute_rsi(closes, period=14):
    if len(closes) < period + 1:
        return None
    deltas = [closes[i] - closes[i - 1] for i in range(1, len(closes))]
    gains  = [d if d > 0 else 0 for d in deltas]
    losses = [-d if d < 0 else 0 for d in deltas]
    avg_gain = sum(gains[:period]) / period
    avg_loss = sum(losses[:period]) / period
    for i in range(period, len(deltas)):
        avg_gain = (avg_gain * (period - 1) + gains[i]) / period
        avg_loss = (avg_loss * (period - 1) + losses[i]) / period
    if avg_loss == 0:
        return 100.0
    rs = avg_gain / avg_loss
    return round(100 - (100 / (1 + rs)), 1)

def ma(closes, period):
    if len(closes) < period:
        return None
    return round(sum(closes[-period:]) / period, 2)

# ─────────────────────────────────────────────
# RECOMMENDATION ENGINE
# ─────────────────────────────────────────────
def derive_recommendation(price, ma50, ma200, rsi, analyst_rec):
    """
    Combines technical signals + analyst consensus into BUY / HOLD / SELL.
    """
    signals = []

    if ma50 and ma200:
        if price > ma50 > ma200:
            signals.append("BUY")
        elif price < ma50 < ma200:
            signals.append("SELL")
        else:
            signals.append("HOLD")

    if rsi:
        if rsi < 35:
            signals.append("BUY")
        elif rsi > 70:
            signals.append("SELL")
        else:
            signals.append("HOLD")

    # Analyst consensus from yfinance (strongBuy/buy/hold/sell/strongSell)
    if analyst_rec:
        rec_lower = analyst_rec.lower()
        if "buy" in rec_lower:
            signals.append("BUY")
        elif "sell" in rec_lower:
            signals.append("SELL")
        else:
            signals.append("HOLD")

    if not signals:
        return "HOLD"
    buys  = signals.count("BUY")
    sells = signals.count("SELL")
    if buys > sells:
        return "BUY"
    elif sells > buys:
        return "SELL"
    return "HOLD"

def rec_emoji(rec):
    return {"BUY": "🟢", "HOLD": "🟡", "SELL": "🔴"}.get(rec, "⚪")

def rsi_label(rsi):
    if rsi is None:
        return "N/A"
    if rsi < 35:
        return f"{rsi} (Oversold)"
    if rsi > 70:
        return f"{rsi} (Overbought)"
    return f"{rsi} (Neutral)"

# ─────────────────────────────────────────────
# 6-MONTH OUTLOOK
# ─────────────────────────────────────────────
def six_month_outlook(price, target_price, rec, rsi, ma50, ma200):
    """Generates a concise one-line 6-month projection."""
    lines = []

    if target_price and price:
        upside = ((target_price - price) / price) * 100
        direction = "upside" if upside >= 0 else "downside"
        lines.append(f"Analyst consensus target ${target_price:.2f} implies {abs(upside):.1f}% {direction}")

    if rsi and rsi < 35:
        lines.append("RSI suggests potential recovery")
    elif rsi and rsi > 70:
        lines.append("RSI suggests near-term pullback risk")

    if ma50 and ma200:
        if ma50 > ma200:
            lines.append("golden cross in place — medium-term trend bullish")
        else:
            lines.append("death cross — medium-term trend bearish")

    return ". ".join(lines) if lines else "Insufficient data for projection."

# ─────────────────────────────────────────────
# RISK ASSESSMENT
# ─────────────────────────────────────────────
def assess_risks(ticker, rsi, ma50, ma200, price, week52_high, week52_low):
    risks = []
    if rsi and rsi > 70:
        risks.append("Overbought (RSI &gt; 70) — correction risk")
    if ma50 and ma200 and ma50 < ma200:
        risks.append("Below 200-day MA — bearish trend")
    if week52_high and price and price >= week52_high * 0.97:
        risks.append("Near 52-week high — resistance level")
    if week52_low and price and price <= week52_low * 1.05:
        risks.append("Near 52-week low — support being tested")
    if ticker in ["PLTR", "SPOT", "NFLX"]:
        risks.append("High valuation / volatile growth stock")
    if ticker in ["TSM"]:
        risks.append("Geopolitical risk (Taiwan Strait)")
    if ticker in ["NVDA"]:
        risks.append("AI capex cycle sensitivity")
    if not risks:
        risks.append("No major near-term technical risks identified")
    return risks

# ─────────────────────────────────────────────
# FETCH & ANALYSE ONE TICKER
# ─────────────────────────────────────────────
def analyse_ticker(symbol) -> tuple[str, str]:
    """Returns (formatted_block, recommendation) for ranking purposes."""
    display = DISPLAY_NAME.get(symbol, symbol)
    try:
        tkr  = yf.Ticker(symbol)
        info = tkr.info
        hist = tkr.history(period="1y")

        if hist.empty:
            return f"⚠️ <b>{display}</b> — No data available.\n", "UNKNOWN"

        closes      = list(hist["Close"])
        price       = round(closes[-1], 2)
        prev_close  = round(closes[-2], 2) if len(closes) > 1 else price
        day_chg     = round(price - prev_close, 2)
        day_pct     = round((day_chg / prev_close) * 100, 2) if prev_close else 0
        chg_arrow   = "▲" if day_chg >= 0 else "▼"
        chg_sign    = "+" if day_chg >= 0 else ""

        rsi_val      = compute_rsi(closes)
        ma50_val     = ma(closes, 50)
        ma200_val    = ma(closes, 200)

        week52_high  = info.get("fiftyTwoWeekHigh")
        week52_low   = info.get("fiftyTwoWeekLow")
        analyst_rec  = info.get("recommendationKey", "")
        target_price = info.get("targetMeanPrice")
        currency     = info.get("currency", "USD")

        rec     = derive_recommendation(price, ma50_val, ma200_val, rsi_val, analyst_rec)
        risks   = assess_risks(display, rsi_val, ma50_val, ma200_val, price, week52_high, week52_low)
        outlook = six_month_outlook(price, target_price, rec, rsi_val, ma50_val, ma200_val)

        risk_text    = "\n    • ".join(html.escape(r) for r in risks)
        currency_sym = "£" if currency in ("GBp", "GBP") else "$"

        block = (
            f"{rec_emoji(rec)} <b>{display}</b> — {rec}\n"
            f"  Price: {currency_sym}{price:,.2f}  {chg_arrow} {chg_sign}{day_chg} ({chg_sign}{day_pct}%)\n"
            f"  MA50: {currency_sym}{ma50_val:,.2f} | MA200: {currency_sym}{ma200_val:,.2f}\n"
            f"  RSI(14): {rsi_label(rsi_val)}\n"
            f"  52W: {currency_sym}{week52_low:,.2f} – {currency_sym}{week52_high:,.2f}\n"
            f"  Risks:\n    • {risk_text}\n"
            f"  6M Outlook: {html.escape(outlook)}\n"
        )
        return block, rec

    except Exception as e:
        log.error("Error fetching %s: %s", symbol, e)
        return f"⚠️ <b>{display}</b> — Error: {html.escape(str(e))}\n", "UNKNOWN"

# ─────────────────────────────────────────────
# BUILD FULL MESSAGE
# ─────────────────────────────────────────────
REC_PRIORITY = {"BUY": 0, "HOLD": 1, "SELL": 2, "UNKNOWN": 3}

def build_message():
    now_sgt = datetime.now(SGT)
    header = (
        f"📊 <b>Daily Stock Alert</b>\n"
        f"🗓 {now_sgt.strftime('%A, %d %B %Y')} | US Market Close\n"
        f"{'─' * 32}\n\n"
    )

    # Fetch all tickers in parallel
    results = [None] * len(TICKERS)
    executor = ThreadPoolExecutor(max_workers=6)
    future_to_idx = {
        executor.submit(analyse_ticker, symbol): i
        for i, symbol in enumerate(TICKERS)
    }
    try:
        for future in as_completed(future_to_idx, timeout=90):
            i = future_to_idx[future]
            symbol = TICKERS[i]
            log.info("Received %s", symbol)
            results[i] = future.result()
    except TimeoutError:
        log.warning("One or more tickers hung past 90s — skipping remainder")
    finally:
        # cancel_futures=True drops queued work; wait=False does not block on hung threads
        executor.shutdown(wait=False, cancel_futures=True)

    # Fill any tickers that hung or never completed
    for future, i in future_to_idx.items():
        if results[i] is None:
            symbol = TICKERS[i]
            display = DISPLAY_NAME.get(symbol, symbol)
            log.warning("No result for %s — marking as timed out", symbol)
            results[i] = (f"⚠️ <b>{display}</b> — Timed out (yfinance hung).\n", "UNKNOWN")

    # Sort: BUY → HOLD → SELL → UNKNOWN
    results.sort(key=lambda x: REC_PRIORITY.get(x[1], 3))

    # Group with section dividers
    blocks = []
    current_group = None
    group_labels  = {"BUY": "🟢 BUY", "HOLD": "🟡 HOLD", "SELL": "🔴 SELL", "UNKNOWN": "⚪ OTHER"}
    for block, rec in results:
        group = rec if rec in group_labels else "UNKNOWN"
        if group != current_group:
            blocks.append(f"<b>── {group_labels[group]} ──</b>\n")
            current_group = group
        blocks.append(block)

    footer = (
        f"\n{'─' * 32}\n"
        f"⚡ <i>Analysis based on Yahoo Finance data + technical indicators.</i>\n"
        f"<i>Not financial advice. Always do your own research.</i>"
    )

    return header + "\n".join(blocks) + footer

# ─────────────────────────────────────────────
# SEND TELEGRAM MESSAGE
# ─────────────────────────────────────────────
def send_telegram(message: str):
    if TELEGRAM_CHAT_ID == "FILL_IN_AFTER_MESSAGING_BOT":
        log.error("Chat ID not set. Run get_chat_id.py first.")
        return False

    url  = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    # Split at newline boundaries to avoid cutting mid-HTML-tag
    chunks = []
    remaining = message
    while len(remaining) > 4000:
        split_at = remaining.rfind("\n", 0, 4000)
        if split_at == -1:
            split_at = 4000
        chunks.append(remaining[:split_at])
        remaining = remaining[split_at:].lstrip("\n")
    if remaining:
        chunks.append(remaining)

    for chunk in chunks:
        data = json.dumps({
            "chat_id":    TELEGRAM_CHAT_ID,
            "text":       chunk,
            "parse_mode": "HTML",
        }).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req) as r:
                result = json.loads(r.read())
                if not result.get("ok"):
                    log.error("Telegram error: %s", result)
                    return False
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")
            log.error("Telegram HTTP %s: %s", e.code, body)
            return False

        except Exception as e:
            log.error("Failed to send Telegram message: %s", e)
            return False

    log.info("Alert sent successfully.")
    return True

# ─────────────────────────────────────────────
# INTERNET CONNECTIVITY WAIT
# ─────────────────────────────────────────────
def wait_for_internet(max_wait_secs=300, interval=10) -> bool:
    """Blocks until a live internet connection is confirmed or timeout expires."""
    deadline = time.monotonic() + max_wait_secs
    attempt = 0
    while time.monotonic() < deadline:
        try:
            urllib.request.urlopen("https://www.google.com", timeout=5)
            if attempt > 0:
                log.info("Internet available after %d attempts.", attempt + 1)
            return True
        except Exception:
            attempt += 1
            remaining = int(deadline - time.monotonic())
            log.info("No internet yet (attempt %d) — retrying in %ds … (%ds left)", attempt, interval, remaining)
            time.sleep(interval)
    log.warning("Internet not available after %ds — aborting.", max_wait_secs)
    return False


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    if already_sent_today():
        log.info("Alert already sent today — skipping.")
        return
    if not wait_for_internet():
        os._exit(1)
    log.info("Building stock alert …")
    message = build_message()
    if send_telegram(message):
        mark_sent_today()
    os._exit(0)  # force exit regardless of lingering yfinance threads

if __name__ == "__main__":
    main()
