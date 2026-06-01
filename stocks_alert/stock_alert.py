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
import urllib.request
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, timedelta
from pathlib import Path

import yfinance as yf

# ─────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────
TELEGRAM_TOKEN   = "8912852669:AAEpqdY1492JaZniwuWq5iERbeseDXO1bFc"
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
# HIGH-GROWTH ETF PICKS (Singaporeans via IBKR)
# ─────────────────────────────────────────────
ETF_PICKS = {
    "QQQ": {
        "name": "Invesco QQQ Trust",
        "tracks": "Nasdaq-100 (100 largest non-financial US companies)",
        "expense_ratio": "0.20%",
        "ucits_alt": "CNDX.L / EQQQ.L (0.30–0.33% — preferred for positions >$60K USD)",
        "why": (
            "The definitive high-growth ETF. Concentrated exposure to FAANG+M, NVIDIA, "
            "and AI-era leaders. Most liquid growth ETF globally; top 10 holdings ≈ 50% of NAV. "
            "Outperforms broad market in sustained tech bull cycles."
        ),
        "risks": [
            "Heavy US tech concentration — top 10 stocks ≈ 50% of portfolio",
            "High P/E valuation; sensitive to interest rate hikes",
            "Single-sector drawdowns can be severe (−33% in 2022)",
        ],
        "five_year": (
            "Analyst consensus projects 10–13% CAGR over 5 years, underpinned by AI/cloud "
            "infrastructure spending. Historical 5Y CAGR (2019–2024): ~18%. Forward estimates "
            "are more conservative given elevated valuations and macro uncertainty."
        ),
    },
    "VGT": {
        "name": "Vanguard Information Technology ETF",
        "tracks": "MSCI US IMI Information Technology 25/50 Index",
        "expense_ratio": "0.10%",
        "ucits_alt": "IITU.L (iShares S&P 500 IT Sector UCITS — 0.20%)",
        "why": (
            "Purest low-cost US tech play. Deeper sector cut than QQQ — excludes "
            "non-tech Nasdaq names (e.g. Amazon, Meta). AAPL + NVDA + MSFT ≈ 50% of NAV. "
            "Vanguard's cost discipline keeps drag minimal at 0.10%."
        ),
        "risks": [
            "Zero diversification outside IT — maximum sector concentration risk",
            "Highly correlated with QQQ; holding both adds little diversification",
            "Semiconductor cycle volatility (NVDA, AVGO, QCOM combined weight ≈ 25%)",
        ],
        "five_year": (
            "More tech-concentrated than QQQ → higher beta. Expected 10–14% CAGR if AI/semiconductor "
            "spend sustains. Outperforms QQQ in pure tech bull runs; underperforms during "
            "rotation out of tech. Favoured by investors with high risk tolerance."
        ),
    },
    "SCHG": {
        "name": "Schwab U.S. Large-Cap Growth ETF",
        "tracks": "Dow Jones U.S. Large-Cap Growth Total Stock Market Index",
        "expense_ratio": "0.04%",
        "ucits_alt": "No direct UCITS equivalent — IWMO.L (iShares MSCI World Momentum, 0.30%) is closest",
        "why": (
            "Lowest-fee high-growth ETF available (0.04%). Broader than QQQ/VGT — blends "
            "tech with high-growth healthcare and consumer discretionary. ~260 holdings "
            "reduce single-stock concentration. Strong Sharpe ratio historically."
        ),
        "risks": [
            "Still 55–60% US tech — growth label does not mean full diversification",
            "No direct Ireland-domiciled UCITS equivalent; US estate tax applies above $60K USD",
            "Slightly lower liquidity than QQQ; wider bid/ask spreads intraday",
        ],
        "five_year": (
            "Most diversified of the three; expected 9–12% CAGR. Broader sector mix provides "
            "a cushion in tech corrections. Morningstar rates it 5-star — best risk-adjusted "
            "growth return at the lowest cost. Recommended as a core holding alongside QQQ or VGT."
        ),
    },
}

ETF_TICKERS = list(ETF_PICKS.keys())  # ["QQQ", "VGT", "SCHG"]

def build_etf_section() -> str:
    """Fetches live data for the 3 curated ETFs and appends static research context."""
    lines = [
        f"\n{'─' * 32}\n"
        f"📈 <b>Top 3 High-Growth ETFs</b> — Singaporeans via IBKR\n"
        f"<i>Curated picks: minimum fees, multiple credible sources</i>\n"
        f"{'─' * 32}\n"
    ]

    for symbol in ETF_TICKERS:
        meta = ETF_PICKS[symbol]
        # Fetch live price block (reuse existing logic)
        live_block, rec = analyse_ticker(symbol)

        lines.append(
            f"\n<b>#{ETF_TICKERS.index(symbol) + 1} — {meta['name']} ({symbol})</b>\n"
            f"  Tracks: {html.escape(meta['tracks'])}\n"
            f"  Fee (Expense Ratio): {meta['expense_ratio']}\n"
            f"  UCITS alt for SG: <i>{html.escape(meta['ucits_alt'])}</i>\n"
        )

        # Inline the live price block (strip the leading emoji+ticker header to avoid duplication)
        live_lines = live_block.strip().splitlines()
        # Skip first line (already shown above) — keep price/MA/RSI/52W/Risks/Outlook lines
        for ln in live_lines[1:]:
            lines.append(f"  {ln.strip()}\n")

        lines.append(
            f"\n  <b>Why hold it:</b> {html.escape(meta['why'])}\n"
            f"\n  <b>Risks:</b>\n"
        )
        for r in meta["risks"]:
            lines.append(f"    • {html.escape(r)}\n")

        lines.append(
            f"\n  <b>5-Year Projection:</b> {html.escape(meta['five_year'])}\n"
        )

    lines.append(
        f"\n{'─' * 32}\n"
        f"⚠️ <i>SG Tax Note: US-domiciled ETFs (QQQ/VGT/SCHG) carry 30% US dividend "
        f"withholding tax and US estate tax risk on holdings above USD 60,000. "
        f"For larger positions, consider the UCITS alternatives listed above "
        f"(Ireland-domiciled, 15% WHT, no estate tax). Not financial advice.</i>\n"
    )

    return "".join(lines)


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

    etf_section = build_etf_section()

    return header + "\n".join(blocks) + footer + etf_section

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
# MAIN
# ─────────────────────────────────────────────
def main():
    if already_sent_today():
        log.info("Alert already sent today — skipping.")
        return
    log.info("Building stock alert …")
    message = build_message()
    if send_telegram(message):
        mark_sent_today()
    os._exit(0)  # force exit regardless of lingering yfinance threads

if __name__ == "__main__":
    main()
