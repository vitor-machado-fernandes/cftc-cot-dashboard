from __future__ import annotations

import pandas as pd

from cotton_weather.config import LOCATIONS_FILE


def load_locations() -> pd.DataFrame:
    locations = pd.read_csv(LOCATIONS_FILE)
    required_columns = {
        "location_id",
        "location_name",
        "state",
        "county",
        "latitude",
        "longitude",
    }
    missing = required_columns.difference(locations.columns)
    if missing:
        missing_list = ", ".join(sorted(missing))
        raise ValueError(f"Missing required location columns: {missing_list}")
    return locations

