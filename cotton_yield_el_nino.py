from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


DATA_URL = (
    "https://ourworldindata.org/grapher/cotton-yield.csv"
    "?v=1&csvType=full&useColumnShortNames=false"
)
COUNTRIES = ["India", "Pakistan", "Australia", "United States"]
START_YEAR = 1990
END_YEAR = 2024
EL_NINO_YEARS = {
    1998: "1997/98 El Nino",
    2016: "2015/16 El Nino",
}
OUTPUT_PNG = Path("cotton_yield_el_nino.png")
OUTPUT_SVG = Path("cotton_yield_el_nino.svg")


def find_yield_column(df: pd.DataFrame) -> str:
    """Find the seed cotton yield variable in tonnes per hectare."""
    candidates: list[tuple[int, str]] = []

    for col in df.columns:
        col_lower = col.lower()
        if col in {"Entity", "Code", "Year"}:
            continue
        if "cotton" not in col_lower or "yield" not in col_lower:
            continue
        if not (
            "tonnes per hectare" in col_lower
            or "tonne per hectare" in col_lower
            or "t/ha" in col_lower
        ):
            continue

        score = 0
        if "seed cotton" in col_lower:
            score += 3
        if "tonnes per hectare" in col_lower:
            score += 2
        if pd.api.types.is_numeric_dtype(df[col]):
            score += 1
        candidates.append((score, col))

    if not candidates:
        raise ValueError(
            "Could not identify the seed cotton yield column in tonnes per hectare. "
            f"Columns found: {list(df.columns)}"
        )

    return max(candidates, key=lambda item: item[0])[1]


def prepare_data() -> tuple[pd.DataFrame, str]:
    df = pd.read_csv(DATA_URL)
    yield_col = find_yield_column(df)

    filtered = df.loc[
        df["Entity"].isin(COUNTRIES) & df["Year"].between(START_YEAR, END_YEAR),
        ["Entity", "Year", yield_col],
    ].rename(columns={"Entity": "country", "Year": "year", yield_col: "yield_t_ha"})

    filtered = filtered.sort_values(["country", "year"]).copy()
    filtered["trend_5yr"] = filtered.groupby("country")["yield_t_ha"].transform(
        lambda series: series.rolling(window=5, center=True, min_periods=3).mean()
    )
    filtered["yield_index"] = filtered["yield_t_ha"] / filtered["trend_5yr"] * 100

    return filtered, yield_col


def plot_yield_index(data: pd.DataFrame) -> None:
    plt.style.use("default")
    fig, axes = plt.subplots(2, 2, figsize=(13, 8), sharex=True, sharey=True)
    fig.patch.set_facecolor("white")

    colors = {
        "India": "#1f77b4",
        "Pakistan": "#2ca02c",
        "Australia": "#d62728",
        "United States": "#9467bd",
    }

    for ax, country in zip(axes.flat, COUNTRIES):
        country_data = data.loc[data["country"] == country]
        ax.set_facecolor("white")

        for year, label in EL_NINO_YEARS.items():
            ax.axvspan(year - 0.5, year + 0.5, color="#f4b183", alpha=0.28, lw=0)
            ax.text(
                year,
                122,
                label,
                rotation=90,
                ha="center",
                va="top",
                fontsize=8,
                color="#7a3b00",
            )

        ax.axhline(100, color="#333333", lw=1.0, ls="--", alpha=0.85)
        ax.plot(
            country_data["year"],
            country_data["yield_index"],
            color=colors[country],
            lw=2.2,
        )
        ax.scatter(
            country_data["year"],
            country_data["yield_index"],
            color=colors[country],
            s=16,
            zorder=3,
        )

        ax.set_title(country, fontsize=12, fontweight="bold", pad=8)
        ax.grid(axis="y", color="#e6e6e6", lw=0.8)
        ax.grid(axis="x", color="#f2f2f2", lw=0.5)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#cccccc")
        ax.spines["bottom"].set_color("#cccccc")
        ax.set_xlim(START_YEAR, END_YEAR)
        ax.set_ylim(75, 125)

    for ax in axes[-1, :]:
        ax.set_xlabel("Year")
    for ax in axes[:, 0]:
        ax.set_ylabel("Yield index, 100 = local 5-year trend")

    fig.suptitle(
        "Cotton yield vs 5-year trend during strong El Nino events",
        fontsize=18,
        fontweight="bold",
        y=0.98,
    )
    fig.text(
        0.5,
        0.925,
        (
            "Seed cotton yield; index vs centered 5-year moving average. "
            "Source: FAO via Our World in Data."
        ),
        ha="center",
        va="center",
        fontsize=10,
        color="#555555",
    )

    fig.tight_layout(rect=[0, 0.03, 1, 0.9])
    fig.savefig(OUTPUT_PNG, dpi=300, bbox_inches="tight", facecolor="white")
    fig.savefig(OUTPUT_SVG, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def print_el_nino_table(data: pd.DataFrame) -> None:
    table = (
        data.loc[data["year"].isin(EL_NINO_YEARS), ["country", "year", "yield_index"]]
        .pivot(index="country", columns="year", values="yield_index")
        .reindex(COUNTRIES)
        .rename(columns={1998: "1998 index", 2016: "2016 index"})
        .round(1)
    )

    print("\nYield index in strong El Nino impact years")
    print(table.to_string())


def main() -> None:
    data, yield_col = prepare_data()
    plot_yield_index(data)
    print(f"Detected yield column: {yield_col}")
    print_el_nino_table(data)
    print(f"\nSaved: {OUTPUT_PNG} and {OUTPUT_SVG}")


if __name__ == "__main__":
    main()
