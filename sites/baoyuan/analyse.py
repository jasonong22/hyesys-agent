"""
Baoyuan CapBank Pattern Analysis
Run at any time to study accumulated data in baoyuan.db.

Usage: python sites/baoyuan/analyse.py
"""

import sqlite3
import statistics
from collections import defaultdict
from datetime import datetime, timezone, timedelta
from pathlib import Path

DB_PATH = Path(__file__).parent / "data" / "baoyuan.db"
SGT     = timezone(timedelta(hours=8))

SITES = ["BAOYUAN-CAPBANK1", "BAOYUAN-CAPBANK2"]


def load_records(conn: sqlite3.Connection) -> list[dict]:
    rows = conn.execute("""
        SELECT site_id, timestamp, Ia, Ib, Ic, quality_tag
        FROM meter_records
        WHERE quality_tag IN ('CLEAN', 'SUSPECT')
        ORDER BY timestamp
    """).fetchall()
    records = []
    for site_id, ts, ia, ib, ic in [(r[0], r[1], r[2], r[3], r[4]) for r in rows]:
        try:
            dt_utc = datetime.fromisoformat(ts).replace(tzinfo=timezone.utc)
            dt_sgt = dt_utc.astimezone(SGT)
        except ValueError:
            continue
        i_total = (ia or 0) + (ib or 0) + (ic or 0)
        records.append({
            "site_id": site_id,
            "dt":      dt_sgt,
            "Ia": ia or 0, "Ib": ib or 0, "Ic": ic or 0,
            "I_total": i_total,
        })
    return records


def section(title: str):
    print(f"\n{'-' * 60}")
    print(f"  {title}")
    print(f"{'-' * 60}")


def summary_stats(values: list[float]) -> str:
    if not values:
        return "no data"
    return (
        f"mean={statistics.mean(values):7.1f}A  "
        f"min={min(values):7.1f}A  "
        f"max={max(values):7.1f}A  "
        f"stdev={statistics.stdev(values):.1f}A" if len(values) > 1
        else f"mean={statistics.mean(values):7.1f}A  (single record)"
    )


def analyse():
    if not DB_PATH.exists():
        print("baoyuan.db not found — run agent1.py first.")
        return

    conn = sqlite3.connect(str(DB_PATH))
    records = load_records(conn)

    if not records:
        print("No records in database yet.")
        conn.close()
        return

    # -- 1. OVERVIEW ----------------------------------------------
    section("OVERVIEW")
    total = len(records)
    first_dt = min(r["dt"] for r in records)
    last_dt  = max(r["dt"] for r in records)
    span_h   = (last_dt - first_dt).total_seconds() / 3600
    print(f"  Records : {total}")
    print(f"  Period  : {first_dt.strftime('%Y-%m-%d %H:%M')} to {last_dt.strftime('%Y-%m-%d %H:%M')} SGT")
    print(f"  Span    : {span_h:.1f} hours  ({span_h/24:.1f} days)")
    print()
    for site in SITES:
        site_recs = [r for r in records if r["site_id"] == site]
        clean_q = conn.execute(
            "SELECT quality_tag, COUNT(*) FROM meter_records WHERE site_id=? GROUP BY quality_tag",
            (site,)
        ).fetchall()
        tag_str = "  ".join(f"{t}={n}" for t, n in clean_q)
        print(f"  {site:25s}  {len(site_recs):5d} records  |  {tag_str}")

    # -- 2. CURRENT LEVELS PER SITE --------------------------------
    section("CURRENT LEVELS PER SITE (3-phase total)")
    for site in SITES:
        vals = [r["I_total"] for r in records if r["site_id"] == site]
        print(f"  {site}")
        print(f"    {summary_stats(vals)}")
        if vals:
            ia_vals = [r["Ia"] for r in records if r["site_id"] == site]
            ib_vals = [r["Ib"] for r in records if r["site_id"] == site]
            ic_vals = [r["Ic"] for r in records if r["site_id"] == site]
            print(f"    Phase A: mean={statistics.mean(ia_vals):.1f}A")
            print(f"    Phase B: mean={statistics.mean(ib_vals):.1f}A")
            print(f"    Phase C: mean={statistics.mean(ic_vals):.1f}A")

    # -- 3. HOURLY PATTERN (SGT) -----------------------------------
    section("HOURLY AVERAGE CURRENT (I_total A, SGT) - all days combined")
    for site in SITES:
        print(f"\n  {site}")
        hourly: dict[int, list[float]] = defaultdict(list)
        for r in records:
            if r["site_id"] == site:
                hourly[r["dt"].hour].append(r["I_total"])
        if hourly:
            print(f"  {'Hour':>5}  {'Avg I_total':>12}  {'Samples':>8}  Bar")
            max_avg = max(statistics.mean(v) for v in hourly.values())
            for h in sorted(hourly.keys()):
                avg = statistics.mean(hourly[h])
                bar = "#" * int(30 * avg / max_avg) if max_avg > 0 else ""
                print(f"  {h:02d}:00  {avg:10.0f}A  {len(hourly[h]):>8}  {bar}")

    # -- 4. DAILY PATTERN -----------------------------------------
    section("DAILY SUMMARY (SGT)")
    for site in SITES:
        print(f"\n  {site}")
        daily: dict[str, list[float]] = defaultdict(list)
        for r in records:
            if r["site_id"] == site:
                day_key = r["dt"].strftime("%Y-%m-%d %a")
                daily[day_key].append(r["I_total"])
        if daily:
            print(f"  {'Date':>15}  {'Avg I':>10}  {'Min I':>10}  {'Max I':>10}  {'Samples':>8}")
            for day in sorted(daily.keys()):
                vals = daily[day]
                print(
                    f"  {day:>15}  "
                    f"{statistics.mean(vals):>10.0f}  "
                    f"{min(vals):>10.0f}  "
                    f"{max(vals):>10.0f}  "
                    f"{len(vals):>8}"
                )

    # -- 5. PHASE BALANCE -----------------------------------------
    section("PHASE BALANCE ANALYSIS")
    for site in SITES:
        site_recs = [r for r in records if r["site_id"] == site]
        if not site_recs:
            continue
        imbalances = []
        for r in site_recs:
            currents = [r["Ia"], r["Ib"], r["Ic"]]
            active = [c for c in currents if c > 1.0]
            if len(active) == 3:
                spread = (max(active) - min(active)) / max(active) * 100
                imbalances.append(spread)
        if imbalances:
            pct_balanced = sum(1 for i in imbalances if i <= 5) / len(imbalances) * 100
            print(f"\n  {site}")
            print(f"    Mean imbalance : {statistics.mean(imbalances):.1f}%")
            print(f"    Max imbalance  : {max(imbalances):.1f}%")
            print(f"    Within 5%      : {pct_balanced:.0f}% of records")

    # -- 6. CAPBANK EVENTS (offline / tripped) --------------------
    section("CAPBANK OFFLINE EVENTS (I_total < 3A)")
    for site in SITES:
        offline = [r for r in records if r["site_id"] == site and r["I_total"] < 3.0]
        print(f"  {site:25s}  {len(offline)} offline records")
        for r in offline[:5]:
            print(f"    {r['dt'].strftime('%Y-%m-%d %H:%M')} SGT  Ia={r['Ia']:.1f} Ib={r['Ib']:.1f} Ic={r['Ic']:.1f}")
        if len(offline) > 5:
            print(f"    ... and {len(offline)-5} more")

    # -- 7. CORRELATION BETWEEN CAPBANK1 AND CAPBANK2 ------------
    section("CAPBANK1 vs CAPBANK2 CORRELATION")
    ts_cb1 = {r["dt"].replace(second=0, microsecond=0): r["I_total"]
              for r in records if r["site_id"] == "BAOYUAN-CAPBANK1"}
    ts_cb2 = {r["dt"].replace(second=0, microsecond=0): r["I_total"]
              for r in records if r["site_id"] == "BAOYUAN-CAPBANK2"}
    common_ts = sorted(set(ts_cb1.keys()) & set(ts_cb2.keys()))
    if len(common_ts) >= 2:
        cb1_vals = [ts_cb1[t] for t in common_ts]
        cb2_vals = [ts_cb2[t] for t in common_ts]
        mean1, mean2 = statistics.mean(cb1_vals), statistics.mean(cb2_vals)
        num = sum((a - mean1) * (b - mean2) for a, b in zip(cb1_vals, cb2_vals))
        den = (sum((a - mean1)**2 for a in cb1_vals) * sum((b - mean2)**2 for b in cb2_vals)) ** 0.5
        corr = num / den if den > 0 else 0
        ratio = mean2 / mean1 if mean1 > 0 else 0
        print(f"  Matched timestamps : {len(common_ts)}")
        print(f"  Pearson correlation: {corr:.3f}  (1.0 = perfectly in sync)")
        print(f"  CB2 / CB1 ratio    : {ratio:.2f}x  (CB2 carries ~{ratio:.1f}x the current of CB1)")
    else:
        print("  Insufficient matched timestamps for correlation.")

    print(f"\n{'-' * 60}")
    print(f"  Analysis complete — {total} records from {first_dt.strftime('%Y-%m-%d')} to {last_dt.strftime('%Y-%m-%d')}")
    print(f"{'-' * 60}\n")
    conn.close()


if __name__ == "__main__":
    analyse()
