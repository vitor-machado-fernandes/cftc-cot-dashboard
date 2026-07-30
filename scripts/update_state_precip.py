from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from cotton_weather.state_precip import update_state_precipitation_dataset


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build state-level cotton-weighted precipitation from PRISM over the CDL cotton footprint."
    )
    parser.add_argument("--history-days", type=int, default=45, help="Initial backfill window when no dataset exists.")
    parser.add_argument("--reprocess-days", type=int, default=7, help="Reprocess this many recent days on every run.")
    parser.add_argument("--start-date", type=date.fromisoformat, default=None, help="Optional explicit start date (YYYY-MM-DD).")
    parser.add_argument("--end-date", type=date.fromisoformat, default=None, help="Optional explicit end date (YYYY-MM-DD).")
    parser.add_argument("--footprint-year", type=int, default=2024, help="CDL footprint year.")
    parser.add_argument("--states", nargs="+", default=None, help="Optional subset of state abbreviations.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    def report(progress: dict) -> None:
        current = progress.get("current_day", 0)
        total = progress.get("total_days", 0)
        day = progress.get("date", "")
        phase = progress.get("phase", "running")
        width = 24
        filled = int(width * current / total) if total else 0
        bar = "#" * filled + "-" * (width - filled)
        print(f"[{bar}] {current}/{total} {phase} {day}")

    summary = update_state_precipitation_dataset(
        history_days=args.history_days,
        reprocess_days=args.reprocess_days,
        start_date=args.start_date,
        end_date=args.end_date,
        footprint_year=args.footprint_year,
        states=args.states,
        progress_callback=report,
    )
    print("State-level cotton precipitation update completed.")
    print(f"Requested range: {summary['requested_start']} to {summary['requested_end']}")
    print(f"States: {', '.join(summary['states'])}")
    print(f"Rows written: {summary['rows_written']}")


if __name__ == "__main__":
    main()
