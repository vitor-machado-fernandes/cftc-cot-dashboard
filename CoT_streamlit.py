import os
import streamlit as st
import pandas as pd
from datetime import datetime
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
from pages_app.open_interest import render_open_interest
from pages_app.on_call import render_on_call
from pages_app.n_traders import render_n_traders
from pages_app.CoT_position import render_position
from pages_app.concentration import render_concentration

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

                if on_call_result["report_count_fetched"] > 0:
                    st.success(
                        f"Cotton on-call parquet updated through {on_call_result['latest_report_date']}."
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
# ---- Sidebar navigation ----
# ------------------------------------------------
page = st.sidebar.radio(
    "Select section",
    ["Open Interest", "On-Call", "Position", "Number of Traders", "Concentration"]
)

# ------------------------------------------------
# ---- Route to sub-app ----
# ------------------------------------------------
if page == "Position":
    render_position()
elif page == "Open Interest":
    render_open_interest()
elif page == "On-Call":
    render_on_call()
elif page == "Number of Traders":
    render_n_traders()
elif page == "Concentration":
    render_concentration()

