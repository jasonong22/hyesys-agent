"""
Train all site models from validated data in hyesys.db.
Nightly retraining — typically completes in <5 seconds.
Usage: python train.py
"""

import logging
import time
from pathlib import Path
from datetime import datetime

from core.store import get_connection, init_db, read_clean_records, read_sar, get_sites
from models.site_model import SiteModel, MODELS_DIR

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("hyesys.train")

OUTPUT_DIR = Path(__file__).parent / "outputs"


def train_all(conn) -> list[dict]:
    sites    = get_sites(conn)
    reports  = []

    if not sites:
        log.warning("No sites found in database. Ingest CSV data first.")
        return reports

    log.info("Training models for %d site(s): %s", len(sites), sites)

    for site_id in sites:
        t0 = time.time()
        meter_records = [dict(r) for r in read_clean_records(conn, site_id=site_id)]
        sar_records   = [dict(r) for r in read_sar(conn, site_id=site_id)]

        if not meter_records:
            log.warning("[%s] No clean records — skipping.", site_id)
            continue

        model = SiteModel(site_id=site_id)
        model.fit(meter_records, sar_records)
        model.save(MODELS_DIR)

        elapsed = round(time.time() - t0, 2)
        report  = {
            "site_id":   site_id,
            "records":   len(meter_records),
            "sar":       len(sar_records),
            "avg_kW":    model.avg_kw,
            "avg_kVAr":  model.avg_kvar,
            "avg_PF":    model.avg_pf,
            "elapsed_s": elapsed,
        }
        reports.append(report)
        _write_report(model, report)

    return reports


def _write_report(model: SiteModel, stats: dict):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    date_str = datetime.utcnow().strftime("%Y%m%d")
    path     = OUTPUT_DIR / f"train_report_{model.site_id}_{date_str}.txt"

    lines = [
        f"HyESys Site Model Training Report",
        f"Generated: {datetime.utcnow().isoformat()}",
        f"{'=' * 50}",
        model.summary(),
        f"{'─' * 50}",
        f"Training records:  {stats['records']}",
        f"SAR records:       {stats['sar']}",
        f"Elapsed:           {stats['elapsed_s']}s",
        f"{'─' * 50}",
        "Hourly kVAr profile:",
    ]
    for h in range(24):
        count = model.hourly_counts[h]
        bar   = "█" * min(int(abs(model.hourly_avg_kvar[h]) / 5), 20)
        lines.append(f"  {h:02d}:00  {model.hourly_avg_kvar[h]:+7.1f} kVAr  {bar}  (n={count})")

    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    log.info("Report written: %s", path)


def print_summary(reports: list[dict]):
    if not reports:
        return
    print("\n" + "=" * 65)
    print("TRAINING COMPLETE")
    print("=" * 65)
    print(f"{'Site':<30} {'Records':>8} {'SAR':>6} {'Avg PF':>8} {'Time(s)':>8}")
    print("-" * 65)
    for r in reports:
        print(f"{r['site_id']:<30} {r['records']:>8} {r['sar']:>6} {r['avg_PF']:>8.4f} {r['elapsed_s']:>8.2f}")
    print("=" * 65 + "\n")


if __name__ == "__main__":
    conn = get_connection()
    init_db(conn)
    reports = train_all(conn)
    print_summary(reports)
    conn.close()
