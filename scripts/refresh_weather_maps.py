from __future__ import annotations

import argparse
from datetime import date, timedelta
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from cotton_weather.config import RAW_DIR  # noqa: E402
from cotton_weather.forecast_qpf import refresh_wpc_qpf_image_cache  # noqa: E402
from cotton_weather.precip_maps import prebuild_precipitation_map_previews  # noqa: E402
from cotton_weather.prism import PrismDownloadError, ensure_prism_asset  # noqa: E402


def _date_range(start_date: date, end_date: date) -> list[date]:
    days = (end_date - start_date).days
    if days < 0:
        return []
    return [start_date + timedelta(days=offset) for offset in range(days + 1)]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Refresh the NOAA/WPC forecast map cache and recent PRISM precipitation grids for CoT Weather maps."
    )
    parser.add_argument("--history-days", type=int, default=21, help="Recent PRISM days to keep refreshed.")
    parser.add_argument("--end-date", type=date.fromisoformat, default=None, help="Optional PRISM end date (YYYY-MM-DD).")
    parser.add_argument("--preview-end-dates", type=int, default=7, help="Recent valid end dates to prebuild per map window.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    end_date = args.end_date or (date.today() - timedelta(days=1))
    start_date = end_date - timedelta(days=max(1, args.history_days) - 1)

    prism_assets = []
    missing = []
    for day in _date_range(start_date, end_date):
        try:
            asset = ensure_prism_asset(variable="ppt", date_value=day, raw_dir=RAW_DIR)
        except PrismDownloadError as exc:
            missing.append(f"{day.isoformat()}: {exc}")
            continue
        prism_assets.append(asset)

    wpc_paths = refresh_wpc_qpf_image_cache()
    preview_meta = prebuild_precipitation_map_previews(recent_end_dates=args.preview_end_dates)

    print("CoT weather map refresh completed.")
    print(f"PRISM ppt assets available: {len(prism_assets)}")
    print(f"PRISM map previews cached: {len(preview_meta)}")
    print(f"NOAA/WPC assets refreshed: {len(wpc_paths)}")
    if missing:
        print("PRISM days unavailable:")
        for message in missing:
            print(f"  - {message}")


if __name__ == "__main__":
    main()
