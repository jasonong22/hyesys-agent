"""
HyESys Pipeline Entry Point.
Runs Agent 1 (validate + ingest) → Agent 2 (analyse + decide) in sequence.

Usage:
    python main.py --csv data/mysite.csv          # single CSV file
    python main.py --csv-dir data/                # all CSVs in a directory
    python main.py --agent2-only                  # skip ingestion, run Agent 2 on DB
    python main.py --train                        # retrain site models after ingestion
    python main.py --report                       # print SAR summary report
"""

import argparse
import logging
import sys
from pathlib import Path

from core.store import get_connection, init_db, read_clean_records, get_sites, get_sar_summary, get_record_count
from core.schema import SITE_CONFIG
from agent1.simulator import run_file, run_directory, print_report
from agent2.state import build_states_from_rows
from agent2.agent import Agent2
from models.site_model import load_all_models
from models.savings import summarise_savings
from train import train_all, print_summary

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("hyesys.main")


def run_agent1(args, conn):
    """Agent 1 — ingest and validate CSV data."""
    summaries = []

    if args.csv:
        path = Path(args.csv)
        if not path.exists():
            log.error("File not found: %s", path)
            sys.exit(1)
        log.info("Agent 1: processing %s", path)
        summaries.append(run_file(path, conn, verbose=args.verbose))

    elif args.csv_dir:
        path = Path(args.csv_dir)
        if not path.is_dir():
            log.error("Directory not found: %s", path)
            sys.exit(1)
        log.info("Agent 1: processing directory %s", path)
        summaries = run_directory(path, conn, verbose=args.verbose)

    if summaries:
        print_report(summaries)

    return summaries


def run_agent2(conn, site_models: dict):
    """Agent 2 — read DB states and issue decisions for all sites."""
    sites = get_sites(conn)
    if not sites:
        log.warning("No sites in database. Run Agent 1 first.")
        return []

    agent   = Agent2(conn, site_models=site_models)
    all_sar = []

    for site_id in sites:
        rows   = read_clean_records(conn, site_id=site_id)
        states = build_states_from_rows(rows)
        log.info("Agent 2: processing %d states for site %s", len(states), site_id)
        sar = agent.process_batch(states)
        all_sar.extend(sar)
        _print_agent2_summary(site_id, sar)

    return all_sar


def _print_agent2_summary(site_id: str, sar: list[dict]):
    if not sar:
        return
    summary = summarise_savings(sar)
    positive_pct = summary.get("positive_outcome_pct", 0)
    avg_fraction = summary.get("avg_loss_fraction") or 0
    print(f"\n  [{site_id}]")
    print(f"    Decisions:        {summary.get('total_decisions', 0)}")
    print(f"    Positive outcomes:{summary.get('positive_outcomes', 0)} ({positive_pct:.1f}%)")
    print(f"    Negative outcomes:{summary.get('negative_outcomes', 0)}")
    print(f"    Avg loss fraction:{avg_fraction:.4f}  ({avg_fraction*100:.2f}% losses eliminated)")


def print_db_report(conn):
    sites = get_sites(conn)
    print("\n" + "=" * 60)
    print("HYESYS DATABASE REPORT")
    print("=" * 60)
    for site_id in sites:
        counts = get_record_count(conn, site_id)
        sar    = get_sar_summary(conn, site_id)
        total  = sum(counts.values())
        print(f"\n  Site: {site_id}")
        print(f"    Meter records: {total}  "
              f"(CLEAN={counts.get('CLEAN',0)}, "
              f"SUSPECT={counts.get('SUSPECT',0)}, "
              f"REJECTED={counts.get('REJECTED',0)})")
        print(f"    SAR log:       "
              f"POSITIVE={sar.get('POSITIVE',0)}, "
              f"NEUTRAL={sar.get('NEUTRAL',0)}, "
              f"NEGATIVE={sar.get('NEGATIVE',0)}")
    print("=" * 60 + "\n")


def main():
    parser = argparse.ArgumentParser(description="HyESys Multi-Agent Pipeline")
    parser.add_argument("--csv",          help="Path to a single CSV file for Agent 1")
    parser.add_argument("--csv-dir",      help="Directory of CSV files for Agent 1")
    parser.add_argument("--agent2-only",  action="store_true", help="Skip Agent 1, run Agent 2 on existing DB data")
    parser.add_argument("--train",        action="store_true", help="Retrain site models after pipeline")
    parser.add_argument("--report",       action="store_true", help="Print database summary report")
    parser.add_argument("--verbose",      action="store_true", help="Verbose per-row logging")
    args = parser.parse_args()

    # ── Initialise DB ──────────────────────────────────────────────
    conn = get_connection()
    init_db(conn)

    # ── Report only ────────────────────────────────────────────────
    if args.report:
        print_db_report(conn)
        conn.close()
        return

    # ── Load site models ───────────────────────────────────────────
    site_models = load_all_models()
    if site_models:
        log.info("Loaded %d site model(s): %s", len(site_models), list(site_models.keys()))
    else:
        log.info("No pre-trained site models found — Agent 2 will use default rules.")

    # ── Agent 1 ────────────────────────────────────────────────────
    if not args.agent2_only:
        if not args.csv and not args.csv_dir:
            log.info("No CSV input specified. Use --csv or --csv-dir to ingest data.")
            log.info("Running Agent 2 on existing database records.")
        else:
            run_agent1(args, conn)

    # ── Agent 2 ────────────────────────────────────────────────────
    run_agent2(conn, site_models)

    # ── Retrain ────────────────────────────────────────────────────
    if args.train:
        log.info("Retraining site models…")
        reports = train_all(conn)
        print_summary(reports)

    conn.close()
    log.info("Pipeline complete.")


if __name__ == "__main__":
    main()
