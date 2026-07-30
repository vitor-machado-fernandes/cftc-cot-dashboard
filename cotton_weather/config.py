from __future__ import annotations

from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
REFERENCE_DIR = DATA_DIR / "reference"
RAW_DIR = DATA_DIR / "raw" / "prism"
CDL_RAW_DIR = DATA_DIR / "raw" / "cdl"
BOUNDARY_RAW_DIR = DATA_DIR / "raw" / "boundaries"
USDA_RAW_DIR = DATA_DIR / "raw" / "usda"
STB_RAW_DIR = DATA_DIR / "raw" / "stb"
PROCESSED_DIR = DATA_DIR / "processed"
STATE_WEIGHT_DIR = PROCESSED_DIR / "state_prism_weights"

LOCATIONS_FILE = REFERENCE_DIR / "cotton_locations.csv"
PROCESSED_DATA_FILE = PROCESSED_DIR / "cotton_daily_weather.parquet"
METADATA_FILE = PROCESSED_DIR / "update_metadata.json"
BACKFILL_PLAN_FILE = PROCESSED_DIR / "backfill_plan.json"
STATE_PRECIP_FILE = PROCESSED_DIR / "state_cotton_precipitation.parquet"
STATE_PRECIP_METADATA_FILE = PROCESSED_DIR / "state_cotton_precipitation_metadata.json"
STATE_PRECIP_PROGRESS_FILE = PROCESSED_DIR / "state_cotton_precipitation_progress.json"
STATE_PRECIP_BACKFILL_PLAN_FILE = PROCESSED_DIR / "state_cotton_precipitation_backfill_plan.json"
USDA_COTTON_METRICS_FILE = PROCESSED_DIR / "us_cotton_metrics.parquet"
USDA_COTTON_METRICS_METADATA_FILE = PROCESSED_DIR / "us_cotton_metrics_metadata.json"
USDA_WAREHOUSE_FILE = USDA_RAW_DIR / "wcmd_cotton_warehouses.csv"
US_COTTON_GINS_PDF = USDA_RAW_DIR / "us_cotton_board_gins.pdf"
US_COTTON_GINS_FILE = PROCESSED_DIR / "us_cotton_gins.csv"
US_COTTON_GINS_METADATA_FILE = PROCESSED_DIR / "us_cotton_gins_metadata.json"
COTTON_RAIL_FILE = PROCESSED_DIR / "cotton_rail_2025.geojson"
COTTON_RAIL_METADATA_FILE = PROCESSED_DIR / "cotton_rail_2025_metadata.json"
COTTON_PORT_EXPORTS_FILE = USDA_RAW_DIR / "cotton_port_exports.csv"
COTTON_PORTS_FILE = PROCESSED_DIR / "cotton_ports_2025.csv"
COTTON_PORTS_METADATA_FILE = PROCESSED_DIR / "cotton_ports_2025_metadata.json"
COTTON_COUNTIES_FILE = PROCESSED_DIR / "cotton_counties_2024.geojson"

PRISM_BASE_URL = "https://data.prism.oregonstate.edu/time_series/us/an/800m"
SUPPORTED_VARIABLES = ("ppt", "tmin", "tmax")
USDA_QUICKSTATS_URL = "https://quickstats.nass.usda.gov/api/api_GET/"

CDL_PORTAL_URL = "https://croplandcros.scinet.usda.gov/"
CDL_FAQ_URL = "https://data.nass.usda.gov/Research_and_Science/Cropland/sarsfaqs2.php"
CDL_METADATA_URL = "https://www.nass.usda.gov/Research_and_Science/Cropland/metadata/metadata_Cropland-Data-Layer-2024.htm"
CDL_COTTON_CLASS_CODE = 2
CDL_CONUS_START_YEAR = 2008
CDL_SERVICE_URL = "https://nassgeodata.gmu.edu/axis2/services/CDLService/GetCDLFile"
CENSUS_COUNTY_BOUNDARY_URL = "https://www2.census.gov/geo/tiger/GENZ2024/shp/cb_2024_us_county_500k.zip"

COTTON_STATE_FIPS = {
    "AL": "01",
    "AZ": "04",
    "AR": "05",
    "CA": "06",
    "FL": "12",
    "GA": "13",
    "KS": "20",
    "LA": "22",
    "MO": "29",
    "MS": "28",
    "NC": "37",
    "NM": "35",
    "OK": "40",
    "SC": "45",
    "TN": "47",
    "TX": "48",
    "VA": "51",
}

DEFAULT_HISTORY_DAYS = 45
DEFAULT_REPROCESS_DAYS = 7
DEFAULT_BACKFILL_CHUNK_DAYS = 90
