"""
CSV ingest and timestamp normalisation.
Handles mixed column naming conventions and multiple timestamp formats.
"""

import csv
from datetime import datetime
from pathlib import Path
from typing import Iterator

from core.schema import COL_SITE_ID, COL_TIMESTAMP, COL_KW, COL_KVAR, COL_PF, COL_VOLTAGE

TIMESTAMP_FORMATS = [
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%dT%H:%M:%S%z",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y %H:%M:%S",
    "%Y/%m/%d %H:%M:%S",
    "%d-%m-%Y %H:%M:%S",
    "%m/%d/%Y %H:%M:%S",
    "%Y-%m-%d %H:%M",
]

# Maps common alternative column names → standard name
COLUMN_ALIASES = {
    COL_KW:      ["kw", "active_power", "active power", "p(kw)", "p", "watt", "kw_total"],
    COL_KVAR:    ["kvar", "reactive_power", "reactive power", "q(kvar)", "q", "kvar_total"],
    COL_PF:      ["pf", "power_factor", "powerfactor", "power factor", "p.f"],
    COL_VOLTAGE: ["voltage_v", "voltage", "v", "volts", "volt", "vac", "v_rms"],
    COL_TIMESTAMP: ["timestamp", "datetime", "time", "date", "date_time", "recorded_at"],
}


def parse_timestamp(raw: str) -> datetime | None:
    raw = raw.strip()
    for fmt in TIMESTAMP_FORMATS:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def normalise_timestamp(raw: str) -> str | None:
    dt = parse_timestamp(raw)
    return dt.strftime("%Y-%m-%dT%H:%M:%S") if dt else None


def _resolve_column(headers_lower: dict[str, str], standard_col: str) -> str | None:
    """Find the actual CSV header that maps to a standard column name."""
    aliases = COLUMN_ALIASES.get(standard_col, [standard_col])
    for alias in aliases:
        if alias in headers_lower:
            return headers_lower[alias]
    return None


def infer_site_id(filepath: Path) -> str:
    """Derive site_id from filename — strips _no_solar suffix for display."""
    name = filepath.stem
    return name.replace("_no_solar", "").strip("_")


def has_solar(filepath: Path) -> bool:
    """Returns False if filename contains _no_solar, else True."""
    return "_no_solar" not in filepath.stem.lower()


def parse_csv(filepath: Path, site_id: str | None = None) -> Iterator[dict]:
    """
    Yields normalised row dicts from a HyESys CSV file.
    Each dict has standard column keys + '_raw' for the original row.
    """
    path = Path(filepath)
    if site_id is None:
        site_id = infer_site_id(path)

    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            return

        # Build lowercase → actual header mapping
        headers_lower = {h.strip().lower(): h.strip() for h in reader.fieldnames}

        # Resolve each standard column once
        col_map = {
            std: _resolve_column(headers_lower, std)
            for std in [COL_TIMESTAMP, COL_KW, COL_KVAR, COL_PF, COL_VOLTAGE]
        }

        for row in reader:
            row = {k.strip(): (v or "").strip() for k, v in row.items()}

            def get(std_col):
                actual = col_map.get(std_col)
                return row.get(actual, "") if actual else ""

            yield {
                COL_SITE_ID:   site_id,
                COL_TIMESTAMP: get(COL_TIMESTAMP),
                COL_KW:        get(COL_KW),
                COL_KVAR:      get(COL_KVAR),
                COL_PF:        get(COL_PF),
                COL_VOLTAGE:   get(COL_VOLTAGE),
                "_raw":        row,
                "_solar":      has_solar(path),
            }


def parse_csv_directory(directory: Path) -> Iterator[tuple[Path, Iterator[dict]]]:
    """Yields (filepath, row_iterator) for every CSV in a directory."""
    for csv_file in sorted(Path(directory).glob("*.csv")):
        yield csv_file, parse_csv(csv_file)
