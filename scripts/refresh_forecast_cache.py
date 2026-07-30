from __future__ import annotations

from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from cotton_weather.forecast_qpf import refresh_wpc_qpf_image_cache  # noqa: E402


def main() -> None:
    cached = refresh_wpc_qpf_image_cache()
    print("NOAA/WPC forecast cache refresh completed.")
    print(f"NOAA/WPC assets refreshed: {len(cached)}")
    for path in cached:
        print(path)


if __name__ == "__main__":
    main()
