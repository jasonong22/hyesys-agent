"""
Agent 1 Simulator — runs Agent 1 on historical CSV files.
Validates all rows, writes CLEAN/SUSPECT to hyesys.db, prints a report.
"""

import logging
from pathlib import Path
from collections import defaultdict

from core.parser import parse_csv, parse_csv_directory
from core.store import get_connection, init_db, write_record
from core.schema import CLEAN, SUSPECT, REJECTED
from agent1.validator import Validator

log = logging.getLogger("hyesys.agent1.simulator")


def run_file(filepath: Path, conn, verbose: bool = False) -> dict:
    """
    Runs Agent 1 on a single CSV file.
    Returns a summary dict with counts per quality tag.
    """
    validator = Validator()
    counts = defaultdict(int)
    site_id = None

    for raw in parse_csv(filepath):
        site_id = raw.get("site_id", "UNKNOWN")
        record, result = validator.validate(raw)
        counts[result.tag] += 1

        if result.tag != REJECTED:
            write_record(conn, record)

        if verbose:
            tag = result.tag
            ts  = record.get("timestamp", "?")
            log.info("[%s] %s %s%s", tag, site_id, ts,
                     f" | {result.reason}" if result.reason else "")

    total = sum(counts.values())
    summary = {
        "file":     filepath.name,
        "site_id":  site_id,
        "total":    total,
        CLEAN:      counts.get(CLEAN,    0),
        SUSPECT:    counts.get(SUSPECT,  0),
        REJECTED:   counts.get(REJECTED, 0),
    }
    return summary


def run_directory(directory: Path, conn, verbose: bool = False) -> list[dict]:
    """Runs Agent 1 on all CSVs in a directory. Returns list of summaries."""
    summaries = []
    for filepath, _ in parse_csv_directory(directory):
        log.info("Processing: %s", filepath.name)
        summary = run_file(filepath, conn, verbose)
        summaries.append(summary)
        _print_summary(summary)
    return summaries


def print_report(summaries: list[dict]):
    total_rows    = sum(s["total"]   for s in summaries)
    total_clean   = sum(s[CLEAN]     for s in summaries)
    total_suspect = sum(s[SUSPECT]   for s in summaries)
    total_rejected = sum(s[REJECTED] for s in summaries)

    print("\n" + "=" * 60)
    print("AGENT 1 — INGESTION REPORT")
    print("=" * 60)
    print(f"{'File':<35} {'Total':>7} {'CLEAN':>7} {'SUSPECT':>8} {'REJECTED':>9}")
    print("-" * 60)
    for s in summaries:
        print(f"{s['file']:<35} {s['total']:>7} {s[CLEAN]:>7} {s[SUSPECT]:>8} {s[REJECTED]:>9}")
    print("-" * 60)
    print(f"{'TOTAL':<35} {total_rows:>7} {total_clean:>7} {total_suspect:>8} {total_rejected:>9}")
    clean_pct = (total_clean / total_rows * 100) if total_rows else 0
    print(f"\nClean rate: {clean_pct:.1f}%")
    print("=" * 60 + "\n")


def _print_summary(s: dict):
    pct = (s[CLEAN] / s["total"] * 100) if s["total"] else 0
    log.info("  %-35s total=%-6d CLEAN=%-6d SUSPECT=%-6d REJECTED=%-6d (%.1f%% clean)",
             s["file"], s["total"], s[CLEAN], s[SUSPECT], s[REJECTED], pct)


if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("data")
    conn = get_connection()
    init_db(conn)

    if path.is_file():
        summary = run_file(path, conn, verbose=True)
        print_report([summary])
    elif path.is_dir():
        summaries = run_directory(path, conn)
        print_report(summaries)
    else:
        print(f"Path not found: {path}")
