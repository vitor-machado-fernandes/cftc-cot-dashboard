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
# ---- Sidebar navigation ----
# ------------------------------------------------
page = st.sidebar.radio(
    "Select section",
    ["Open Interest", "On-Call"] # later add: , "Position", "Concentration", etc]
)

# ------------------------------------------------
# ---- Route to sub-app ----
# ------------------------------------------------
if page == "Open Interest":
    render_open_interest()
elif page == "On-Call":
    render_on_call()