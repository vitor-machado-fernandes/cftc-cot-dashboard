import os
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px

st.markdown(
    """
    <style>
        .block-container {
            max-width: 1000px;
            padding-left: 2rem;
            padding-right: 2rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------
# Page setup
# ---------------------------------------------------------------------
st.set_page_config(layout="wide")
st.title("CFTC Commitment of Traders Dashboard")


from CoT_updater import run_update_check
from usda_crop_progress_condition_updater import refresh_crop_progress_condition_data
from pages_app.open_interest import render_open_interest
from pages_app.on_call import render_on_call
from pages_app.n_traders import render_n_traders
from pages_app.CoT_position import render_position
from pages_app.concentration import render_concentration
from pages_app.conab_cotton_progress import render_conab_cotton_progress
from pages_app.crop_progress_condition import render_crop_progress_condition
from pages_app.mato_grosso import render_mato_grosso
from pages_app.stocks_use import render_stocks_use
from pages_app.weather import render_weather
from pages_app.ndvi_index import render_ndvi_index


def render_home():
    st.header("Home")
    st.write(
        """
        The Commitment of Traders (CoT) report is a weekly CFTC publication that shows how different types of market participants are positioned in futures and options markets.

        Traders are grouped into categories such as producers, swap dealers, managed money, other reportables, and non-reportables. Looking at those groups over time helps us understand positioning, participation, hedging pressure, and market structure.
        """
    )

    st.markdown(
        """
        **Sections**

        - `Concentration`: Shows how concentrated long and short exposure is among the largest traders.
        - `CONAB`: Tracks Brazilian cotton planting and harvest progress from CONAB bulletins.
        - `Crop Progress & Condition`: Tracks USDA weekly planting, development, and crop-condition data.
        - `Mato Grosso`: Tracks Brazil crop data from IMEA.
        - `NDVI index`
        - `Number of Traders`: Tracks how many traders are active on the long, short, and spread sides.
        - `On-Call`: Focuses on cotton on-call activity, including unfixed sales and purchases.
        - `Open Interest`: Explores open interest trends and seasonal patterns.
        - `Position`: Displays long, short, spread, and net positioning by trader category.
        - `Stocks & Use`: Tracks supply, demand, stocks, and use data.
        - `Weather`
        """
    )

    st.caption("Use the sidebar to navigate between sections.")

# ------------------------------------------------
# Run updater once per session
# ------------------------------------------------
if "cot_update_ran" not in st.session_state:
    st.session_state["cot_update_ran"] = True

    with st.spinner("Checking for CFTC CoT updates..."):
        try:
            result = run_update_check(data_dir=".", force=False)  # set data_dir to where your xlsx files live
            if result["did_update"]:
                st.success(
                    f"CoT files updated (local {result['sentinel_local']} → CFTC {result['sentinel_cftc']})."
                )
                for msg in result["messages"]:
                    st.caption(msg)
            else:
                st.info(f"CoT files already up to date (latest {result['sentinel_local']}).")
        except Exception as e:
            st.error(f"CoT update check failed: {e}")
            st.stop()


# ------------------------------------------------
# Run cotton on-call updater once per session
# ------------------------------------------------
if "on_call_update_ran" not in st.session_state:
    st.session_state["on_call_update_ran"] = True

    with st.spinner("Checking for cotton on-call updates..."):
        try:
            try:
                from cotton_on_call_updater import build_cotton_on_call_parquet
            except ModuleNotFoundError as e:
                missing_module = getattr(e, "name", None) or "a scraper dependency"
                st.info(
                    f"Cotton on-call auto-update skipped because `{missing_module}` is not installed in this Streamlit environment."
                )
                build_cotton_on_call_parquet = None

            if build_cotton_on_call_parquet is not None:
                try:
                    on_call_result = build_cotton_on_call_parquet(
                        data_dir=".",
                        force=False,
                    )
                except Exception:
                    # Work-PC fallback when Python HTTPS/proxy handling blocks a normal request.
                    on_call_result = build_cotton_on_call_parquet(
                        data_dir=".",
                        force=False,
                        trust_env=False,
                        verify=False,
                    )

                if on_call_result.get("did_update"):
                    latest_release = on_call_result.get("latest_release_date")
                    latest_as_of = on_call_result.get("latest_report_date")
                    st.success(
                        "Cotton on-call parquet updated "
                        f"(released {latest_release or 'N/A'}, as of {latest_as_of or 'N/A'})."
                    )
                    if on_call_result["errors"]:
                        st.warning(
                            f"Cotton on-call update completed with {len(on_call_result['errors'])} parsing gaps."
                        )
                else:
                    st.info("Cotton on-call parquet already up to date.")
        except Exception as e:
            st.warning(f"Cotton on-call update check failed: {e}")

# ------------------------------------------------
# Run USDA crop progress / condition updater once per session
# ------------------------------------------------
USDA_UPDATE_CHECK_INTERVAL = timedelta(hours=6)
last_usda_check = st.session_state.get("usda_crop_progress_last_checked_at")
should_run_usda_update = (
    last_usda_check is None
    or datetime.utcnow() - last_usda_check >= USDA_UPDATE_CHECK_INTERVAL
)

if should_run_usda_update:
    st.session_state["usda_crop_progress_last_checked_at"] = datetime.utcnow()

    with st.spinner("Checking for USDA crop progress and condition updates..."):
        try:
            usda_api_key = (
                st.secrets.get("USDA_QUICKSTATS_API_KEY")
                if hasattr(st, "secrets")
                else None
            ) or os.getenv("USDA_QUICKSTATS_API_KEY") or os.getenv("QUICKSTATS_API_KEY")

            crop_result = refresh_crop_progress_condition_data(
                data_dir=".",
                api_key=usda_api_key,
                force=False,
            )

            if crop_result.get("skipped") and crop_result.get("reason") == "missing_api_key":
                st.info(
                    "USDA crop progress auto-update skipped because `USDA_QUICKSTATS_API_KEY` is not configured."
                )
            elif crop_result.get("did_update"):
                remote_latest = crop_result.get("remote_latest")
                remote_label = remote_latest.date() if remote_latest is not None else "N/A"
                st.session_state.pop("crop_progress_report_date", None)
                st.session_state.pop("crop_condition_report_date", None)
                st.success(f"USDA crop progress data updated through {remote_label}.")
                for msg in crop_result["messages"]:
                    st.caption(msg)
            else:
                local_latest = crop_result.get("local_latest")
                local_label = local_latest.date() if local_latest is not None else "N/A"
                st.info(f"USDA crop progress data already up to date (latest {local_label}).")
        except Exception as e:
            st.warning(f"USDA crop progress update check failed: {e}")


# ------------------------------------------------
# ---- Sidebar navigation ----
# ------------------------------------------------
page = st.sidebar.radio(
    "Select section",
    [
        "Home",
        "Concentration",
        "CONAB",
        "Crop Progress & Condition",
        "Mato Grosso",
        "NDVI index",
        "Number of Traders",
        "On-Call",
        "Open Interest",
        "Position",
        "Stocks & Use",
        "Weather",
    ],
)

# ------------------------------------------------
# ---- Route to sub-app ----
# ------------------------------------------------
if page == "Home":
    render_home()
elif page == "Position":
    render_position()
elif page == "Open Interest":
    render_open_interest()
elif page == "On-Call":
    render_on_call()
elif page == "Number of Traders":
    render_n_traders()
elif page == "Concentration":
    render_concentration()
elif page == "CONAB":
    render_conab_cotton_progress()
elif page == "Crop Progress & Condition":
    render_crop_progress_condition()
elif page == "Mato Grosso":
    render_mato_grosso()
elif page == "NDVI index":
    render_ndvi_index()
elif page == "Stocks & Use":
    render_stocks_use()
elif page == "Weather":
    render_weather()

