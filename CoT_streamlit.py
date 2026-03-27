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
        - `Number of Traders`: Tracks how many traders are active on the long, short, and spread sides.
        - `On-Call`: Focuses on cotton on-call activity, including unfixed sales and purchases.
        - `Open Interest`: Explores open interest trends and seasonal patterns.
        - `Position`: Displays long, short, spread, and net positioning by trader category.
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
    ["Home", "Concentration", "Number of Traders", "On-Call", "Open Interest", "Position"]
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

