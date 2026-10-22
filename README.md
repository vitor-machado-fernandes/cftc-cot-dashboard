# Fundamentals & CoT Dashboard

## 🚀 Live Demo
👉 **https://cftc-cot-dashboard.streamlit.app**


## Weekly Brazil export updates

`.github/workflows/update-brazil-exports.yml` checks ComexStat every Friday at
14:23 UTC (09:23 Chicago daylight time / 08:23 standard time). It refreshes the
monthly cotton export history and commits changed `brazil_cotton_exports.parquet`
and `brazil_cotton_exports_metadata.json` files to the default branch, so saved
data survives Streamlit rebuilds. This is a weekly check of monthly data.

The workflow must be pushed to the default branch to activate the schedule.
GitHub Actions must be enabled and branch rules must allow the workflow's
`GITHUB_TOKEN` to push commits. No ComexStat API secret is required.
For an immediate check, select **Actions > Update Brazil cotton exports >
Run workflow**. The run summary reports the latest saved month.

API failures (including exhausted rate-limit retries), empty data, invalid or
duplicate rows, and downloads missing previously saved month/code pairs fail the
job before committing. Existing repository data is retained; inspect the failed
run's logs and retry manually if needed. Historical revisions are included on
successful refreshes. The app's existing runtime refresh remains available.
