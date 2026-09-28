from __future__ import annotations

import json
import html
import re
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlencode

import altair as alt
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

DATA_PATH = Path("data/university_dashboard_with_dea_efficiency.xlsx")
SHEET_NAME = "Dashboard_Data_with_DEA"

st.set_page_config(
    page_title="University Performance Dashboard",
    page_icon=":bar_chart:",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.6rem; padding-bottom: 2rem;}
    .small-note {color: #7a7f8a; font-size: 0.92rem;}
    .metric-card {
        border: 1px solid rgba(49, 51, 63, 0.12);
        border-radius: 12px;
        padding: 1rem;
        background: rgba(250, 250, 250, 0.65);
    }
    .evidence-card {
        border: 1px solid rgba(49, 51, 63, 0.14);
        border-radius: 10px;
        padding: 0.72rem 0.85rem;
        margin: 0.45rem 0;
        background: rgba(250, 250, 250, 0.72);
    }
    .evidence-id {
        display: inline-block;
        min-width: 2.2rem;
        font-weight: 700;
    }
    .context-link-note {
        border-left: 3px solid rgba(49, 51, 63, 0.35);
        padding-left: 0.75rem;
        margin: 0.25rem 0 0.8rem 0;
        color: #5b606b;
        font-size: 0.9rem;
    }
    .focus-callout {
        border: 2px solid rgba(49, 51, 63, 0.55);
        border-radius: 10px;
        padding: 0.65rem 0.8rem;
        margin: 0.35rem 0 0.8rem 0;
        background: rgba(255, 248, 220, 0.75);
    }
    .focus-value {
        display: inline-block;
        font-weight: 800;
        font-size: 1.05rem;
        border-bottom: 3px solid rgba(49, 51, 63, 0.55);
        padding: 0 0.15rem;
    }
    .companion-data-link,
    .companion-evidence-link {
        color: inherit;
        text-decoration: underline;
        text-decoration-thickness: 2px;
        text-underline-offset: 2px;
        cursor: pointer;
        font-weight: 700;
    }
    .companion-data-link:hover,
    .companion-evidence-link:hover {
        background: rgba(255, 232, 128, 0.45);
        border-radius: 4px;
    }
    [id^="source-"] {
        scroll-margin-top: 90px;
    }
    .evidence-scroll-pulse {
        outline: 2px solid rgba(255, 75, 75, 0.55);
        outline-offset: 8px;
        border-radius: 8px;
        transition: outline-color 0.8s ease;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_data() -> pd.DataFrame:
    df = pd.read_excel(DATA_PATH, sheet_name=SHEET_NAME)
    df["year"] = df["year"].astype(int)
    return df


def clean_value(value: Any) -> Any:
    if pd.isna(value):
        return None
    if hasattr(value, "item"):
        value = value.item()
    if isinstance(value, (int, float, str, bool)):
        return value
    return str(value)


def _mapping_get(value: Any, key: str, default: Any = None) -> Any:
    """Read from dict-like Streamlit event objects without depending on one version."""
    if value is None:
        return default
    try:
        return value.get(key, default)
    except Exception:
        return getattr(value, key, default)


def extract_chart_selection_field(chart_key: str, selection_name: str, field: str) -> Any:
    """Extract one field from a Streamlit/Vega-Lite point-selection state."""
    chart_state = st.session_state.get(chart_key)
    selection_state = _mapping_get(chart_state, "selection", {})
    payload = _mapping_get(selection_state, selection_name)
    if not payload:
        return None

    candidates = payload if isinstance(payload, (list, tuple)) else [payload]
    for item in reversed(candidates):
        candidate = _mapping_get(item, field)
        if isinstance(candidate, (list, tuple)):
            candidate = candidate[-1] if candidate else None
        if candidate is not None and candidate != "":
            return candidate
    return None


def extract_chart_selected_university(chart_key: str, selection_name: str) -> str | None:
    candidate = extract_chart_selection_field(chart_key, selection_name, "university")
    return str(candidate) if candidate else None


def _clear_deep_focus() -> None:
    try:
        if "focus_field" in st.query_params:
            del st.query_params["focus_field"]
    except Exception:
        pass


def _sync_chart_selection_to_sidebar(chart_key: str, selection_name: str, source_page: str) -> None:
    """Make a clicked university mark the active university for the entire dashboard."""
    clicked_university = extract_chart_selected_university(chart_key, selection_name)
    if not clicked_university:
        return

    valid_options = st.session_state.get("_current_university_options", [])
    if valid_options and clicked_university not in valid_options:
        return

    _clear_deep_focus()
    st.session_state["university_selector"] = clicked_university
    st.session_state["chart_selection_meta"] = {
        "university": clicked_university,
        "source_page": source_page,
        "interaction": "chart mark click",
    }
    st.session_state["chart_focus_meta"] = None
    st.session_state["ai_answer"] = ""
    st.session_state["last_ai_signature"] = ""


def _sync_chart_focus(chart_key: str, selection_name: str, source_page: str, field_name: str = "Field", label_name: str = "Label", value_name: str = "Value") -> None:
    """Focus the Companion and visualization on a clicked dimension/indicator bar."""
    field = extract_chart_selection_field(chart_key, selection_name, field_name)
    if not field:
        return
    label = extract_chart_selection_field(chart_key, selection_name, label_name)
    value = extract_chart_selection_field(chart_key, selection_name, value_name)
    _clear_deep_focus()
    st.session_state["chart_focus_meta"] = {
        "source_page": source_page,
        "interaction": "chart bar click",
        "field": str(field),
        "label": clean_value(label) if label is not None else str(field),
        "value": clean_value(value),
    }
    st.session_state["ai_answer"] = ""
    st.session_state["last_ai_signature"] = ""


def on_finance_scatter_select() -> None:
    _sync_chart_selection_to_sidebar("finance_scatter_chart", "finance_point_selection", "Finance Explorer")


def on_dea_scatter_select() -> None:
    _sync_chart_selection_to_sidebar("dea_scatter_chart", "dea_point_selection", "DEA Efficiency Explorer")


def on_overview_top_select() -> None:
    _sync_chart_selection_to_sidebar("overview_top_chart", "overview_university_selection", "Overview")


def on_ranking_bar_select() -> None:
    _sync_chart_selection_to_sidebar("ranking_bar_chart", "ranking_university_selection", "Ranking Explorer")


def on_dea_top_select() -> None:
    _sync_chart_selection_to_sidebar("dea_top_chart", "dea_top_university_selection", "DEA Efficiency Explorer")


def on_teaching_research_scatter_select() -> None:
    _sync_chart_selection_to_sidebar("teaching_research_scatter_chart", "teaching_research_university_selection", "Teaching and Research")


def on_profile_dimension_select() -> None:
    _sync_chart_focus("profile_dimension_chart", "profile_dimension_selection", "University Profile")


def on_finance_indicator_select() -> None:
    _sync_chart_focus("finance_indicator_chart", "finance_indicator_selection", "Finance Explorer", label_name="Indicator")


def on_teaching_indicator_select() -> None:
    _sync_chart_focus("teaching_indicator_chart", "teaching_indicator_selection", "Teaching and Research", label_name="Indicator")


def on_research_indicator_select() -> None:
    _sync_chart_focus("research_indicator_chart", "research_indicator_selection", "Teaching and Research", label_name="Indicator")


def on_teaching_research_score_select() -> None:
    _sync_chart_focus("teaching_research_score_chart", "teaching_research_score_selection", "Teaching and Research")


def on_comparison_score_select() -> None:
    field = extract_chart_selection_field("comparison_score_chart", "comparison_score_selection", "Field")
    dimension = extract_chart_selection_field("comparison_score_chart", "comparison_score_selection", "Dimension")
    value = extract_chart_selection_field("comparison_score_chart", "comparison_score_selection", "Score")
    uni = extract_chart_selection_field("comparison_score_chart", "comparison_score_selection", "University")
    if not field:
        return
    _clear_deep_focus()
    st.session_state["chart_focus_meta"] = {
        "source_page": "University Comparison", "interaction": "comparison bar click",
        "field": str(field), "label": f"{uni} — {dimension}", "value": clean_value(value), "university": clean_value(uni),
    }
    st.session_state["ai_answer"] = ""
    st.session_state["last_ai_signature"] = ""


def on_comparison_gap_select() -> None:
    _sync_chart_focus("comparison_gap_chart", "comparison_gap_selection", "University Comparison", field_name="Field", label_name="Dimension", value_name="Gap")


def on_time_change_select() -> None:
    _sync_chart_focus("time_change_chart", "time_change_selection", "Time Dynamics / What Changed", field_name="Field", label_name="Dimension", value_name="Change")


def on_sidebar_university_change() -> None:
    """A manual sidebar choice supersedes previous chart-origin selections and deep focus."""
    _clear_deep_focus()
    st.session_state["chart_selection_meta"] = None
    st.session_state["chart_focus_meta"] = None
    st.session_state["ai_answer"] = ""
    st.session_state["last_ai_signature"] = ""


def format_number(value: Any, decimals: int = 1) -> str:
    if value is None or pd.isna(value):
        return "n/a"
    return f"{float(value):,.{decimals}f}"


def is_valid_api_key(api_key: Any) -> bool:
    if not isinstance(api_key, str):
        return False
    api_key = api_key.strip()
    if not api_key:
        return False
    try:
        api_key.encode("ascii")
    except UnicodeEncodeError:
        return False
    return api_key.startswith("sk-") and len(api_key) > 20


def score_label(value: float) -> str:
    if value >= 70:
        return "high"
    if value >= 55:
        return "moderate"
    return "low"


def selected_context(row: pd.Series, view_type: str = "University Profile") -> dict:
    fields = [
        "university",
        "year",
        "region",
        "macro_area",
        "size_class",
        "overall_score",
        "overall_rank_year",
        "teaching_score",
        "placement_score",
        "research_score",
        "financial_score",
        "national_avg_overall_score",
        "macro_area_avg_overall_score",
        "ffo_per_student",
        "operating_cost_per_student",
        "personnel_cost_share",
        "public_revenue_share",
        "student_contribution_share",
        "performance_quota_share",
        "economic_financial_sustainability_index",
        "placement_data_completeness",
        "student_staff_ratio",
        "enrolled_students",
        "second_year_retention_pct",
        "inactive_students_reversed_score",
        "graduation_within_standard_pct",
        "employment_index",
        "graduation_intensity",
        "publications_per_teaching_staff",
        "citations_per_publication",
        "h_index",
        "highly_cited_researchers",
        "nature_science_articles",
        "staff_per_1000_students",
        "non_academic_staff_per_1000_students",
        "dea_vrs_efficiency_100",
        "dea_crs_efficiency_100",
        "dea_scale_efficiency_100",
        "dea_vrs_rank_year",
        "dea_crs_rank_year",
        "efficiency_category",
    ]
    context = {field: clean_value(row.get(field)) for field in fields if field in row.index}
    context["view_type"] = view_type
    return context


def make_overview_context(filtered: pd.DataFrame, year: int, macro_area: str, region: str, size_class: str) -> dict:
    top = filtered.sort_values("overall_score", ascending=False).head(3)
    bottom = filtered.sort_values("overall_score", ascending=True).head(3)
    return {
        "view_type": "Overview",
        "year": int(year),
        "macro_area_filter": macro_area,
        "region_filter": region,
        "size_class_filter": size_class,
        "number_of_universities": int(filtered["university"].nunique()),
        "average_overall_score": clean_value(filtered["overall_score"].mean()),
        "average_teaching_score": clean_value(filtered["teaching_score"].mean()),
        "average_placement_score": clean_value(filtered["placement_score"].mean()),
        "average_research_score": clean_value(filtered["research_score"].mean()),
        "average_financial_score": clean_value(filtered["financial_score"].mean()),
        "average_profile_range": clean_value(max(
            filtered["teaching_score"].mean(),
            filtered["placement_score"].mean(),
            filtered["research_score"].mean(),
            filtered["financial_score"].mean(),
        ) - min(
            filtered["teaching_score"].mean(),
            filtered["placement_score"].mean(),
            filtered["research_score"].mean(),
            filtered["financial_score"].mean(),
        )),
        "average_dea_vrs_efficiency": clean_value(filtered["dea_vrs_efficiency_100"].mean()) if "dea_vrs_efficiency_100" in filtered.columns else None,
        "top_universities": top[["university", "overall_score", "overall_rank_year"]].to_dict("records"),
        "lowest_universities": bottom[["university", "overall_score", "overall_rank_year"]].to_dict("records"),
    }


def comparison_context(row_a: pd.Series, row_b: pd.Series) -> dict:
    fields = [
        "university",
        "year",
        "region",
        "macro_area",
        "size_class",
        "overall_score",
        "overall_rank_year",
        "teaching_score",
        "placement_score",
        "research_score",
        "financial_score",
        "ffo_per_student",
        "operating_cost_per_student",
        "personnel_cost_share",
        "public_revenue_share",
        "student_contribution_share",
        "performance_quota_share",
        "economic_financial_sustainability_index",
        "student_staff_ratio",
        "employment_index",
        "graduation_intensity",
        "publications_per_teaching_staff",
        "citations_per_publication",
        "h_index",
        "dea_vrs_efficiency_100",
        "dea_crs_efficiency_100",
        "dea_scale_efficiency_100",
        "dea_vrs_rank_year",
        "efficiency_category",
    ]
    a = {field: clean_value(row_a.get(field)) for field in fields if field in row_a.index}
    b = {field: clean_value(row_b.get(field)) for field in fields if field in row_b.index}
    score_gaps = {
        "overall_gap_a_minus_b": clean_value(row_a.get("overall_score") - row_b.get("overall_score")),
        "teaching_gap_a_minus_b": clean_value(row_a.get("teaching_score") - row_b.get("teaching_score")),
        "placement_gap_a_minus_b": clean_value(row_a.get("placement_score") - row_b.get("placement_score")),
        "research_gap_a_minus_b": clean_value(row_a.get("research_score") - row_b.get("research_score")),
        "financial_gap_a_minus_b": clean_value(row_a.get("financial_score") - row_b.get("financial_score")),
    }
    return {
        "view_type": "University Comparison",
        "year": clean_value(row_a.get("year")),
        "university_a": a,
        "university_b": b,
        "score_gaps_a_minus_b": score_gaps,
    }


def make_score_profile(row: pd.Series, include_overall: bool = True) -> pd.DataFrame:
    rows = []
    if include_overall:
        rows.append({"Dimension": "Overall", "Label": "Overall", "Field": "overall_score", "Score": float(row["overall_score"]), "Value": float(row["overall_score"])})
    for label, field in [("Teaching", "teaching_score"), ("Placement", "placement_score"), ("Research", "research_score"), ("Financial", "financial_score")]:
        rows.append({"Dimension": label, "Label": label, "Field": field, "Score": float(row[field]), "Value": float(row[field])})
    return pd.DataFrame(rows)


def make_indicator_bar(row: pd.Series, indicators: dict[str, str], focus_field: str | None = None) -> alt.Chart:
    data = []
    for col, label in indicators.items():
        if col in row.index and pd.notna(row[col]):
            data.append({"Indicator": label, "Label": label, "Field": col, "Value": float(row[col]), "Focused": bool(focus_field == col)})
    chart_df = pd.DataFrame(data)
    if chart_df.empty:
        chart_df = pd.DataFrame({"Indicator": ["No data"], "Label": ["No data"], "Field": ["none"], "Value": [0.0], "Focused": [False]})
    base = (
        alt.Chart(chart_df)
        .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
        .encode(
            y=alt.Y("Indicator:N", sort="-x", title=None),
            x=alt.X("Value:Q", title="Value"),
            color=alt.condition("datum.Focused", alt.value("#111111"), alt.value("#4c78a8")),
            stroke=alt.condition("datum.Focused", alt.value("#111111"), alt.value(None)),
            strokeWidth=alt.condition("datum.Focused", alt.value(3), alt.value(0)),
            tooltip=["Indicator", "Field", alt.Tooltip("Value:Q", format=",.2f")],
        )
    )
    labels = (
        alt.Chart(chart_df)
        .mark_text(align="left", dx=5)
        .encode(y=alt.Y("Indicator:N", sort="-x"), x=alt.X("Value:Q"), text=alt.Text("Value:Q", format=",.2f"))
    )
    return (base + labels).properties(height=max(260, 38 * len(chart_df)))


def numeric_series(data: pd.DataFrame, column: str) -> pd.Series:
    if column not in data.columns:
        return pd.Series(dtype=float)
    return pd.to_numeric(data[column], errors="coerce").dropna()


def percentile_position(data: pd.DataFrame, column: str, value: Any) -> float | None:
    series = numeric_series(data, column)
    if series.empty or value is None or pd.isna(value):
        return None
    value = float(value)
    return float((series <= value).mean() * 100)


def position_label(percentile: float | None) -> str:
    if percentile is None:
        return "not available"
    if percentile >= 85:
        return "upper tail"
    if percentile >= 65:
        return "above the central range"
    if percentile >= 35:
        return "central range"
    if percentile >= 15:
        return "below the central range"
    return "lower tail"


def median_gap_label(value: Any, median: Any) -> str:
    if value is None or median is None or pd.isna(value) or pd.isna(median):
        return "not available"
    gap = float(value) - float(median)
    if abs(gap) < 1e-9:
        return "close to the median"
    return "above the median" if gap > 0 else "below the median"


def cluster_label(x_percentile: float | None, y_percentile: float | None) -> str:
    if x_percentile is None or y_percentile is None:
        return "not available"
    if 25 <= x_percentile <= 75 and 25 <= y_percentile <= 75:
        return "inside the main cluster"
    if x_percentile >= 90 or x_percentile <= 10 or y_percentile >= 90 or y_percentile <= 10:
        return "near the edge of the distribution"
    return "outside the central band but not an extreme outlier"


def add_profile_position_context(context: dict, filtered_data: pd.DataFrame, row: pd.Series) -> dict:
    for col in ["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"]:
        if col in filtered_data.columns:
            context[f"{col}_median_current_filter"] = clean_value(numeric_series(filtered_data, col).median())
            context[f"{col}_percentile_current_filter"] = clean_value(percentile_position(filtered_data, col, row.get(col)))
    return context


def add_finance_position_context(context: dict, filtered_data: pd.DataFrame, x_col: str, y_col: str, row: pd.Series) -> dict:
    x_series = numeric_series(filtered_data, x_col)
    y_series = numeric_series(filtered_data, y_col)
    x_val = row.get(x_col)
    y_val = row.get(y_col)
    x_pct = percentile_position(filtered_data, x_col, x_val)
    y_pct = percentile_position(filtered_data, y_col, y_val)
    context.update(
        {
            "x_axis_median_current_filter": clean_value(x_series.median()) if not x_series.empty else None,
            "y_axis_median_current_filter": clean_value(y_series.median()) if not y_series.empty else None,
            "x_axis_percentile_current_filter": clean_value(x_pct),
            "y_axis_percentile_current_filter": clean_value(y_pct),
            "x_axis_position_label": position_label(x_pct),
            "y_axis_position_label": position_label(y_pct),
            "scatter_cluster_position": cluster_label(x_pct, y_pct),
            "x_axis_median_gap_label": median_gap_label(x_val, x_series.median()) if not x_series.empty else "not available",
            "y_axis_median_gap_label": median_gap_label(y_val, y_series.median()) if not y_series.empty else "not available",
        }
    )
    return context


def add_teaching_research_position_context(context: dict, filtered_data: pd.DataFrame, row: pd.Series) -> dict:
    pairs = {
        "teaching_score": row.get("teaching_score"),
        "placement_score": row.get("placement_score"),
        "research_score": row.get("research_score"),
        "second_year_retention_pct": row.get("second_year_retention_pct"),
        "graduation_within_standard_pct": row.get("graduation_within_standard_pct"),
        "employment_index": row.get("employment_index"),
        "publications_per_teaching_staff": row.get("publications_per_teaching_staff"),
        "citations_per_publication": row.get("citations_per_publication"),
        "h_index": row.get("h_index"),
    }
    for col, value in pairs.items():
        if col in filtered_data.columns:
            series = numeric_series(filtered_data, col)
            context[f"{col}_median_current_filter"] = clean_value(series.median()) if not series.empty else None
            context[f"{col}_percentile_current_filter"] = clean_value(percentile_position(filtered_data, col, value))
    teaching_pct = context.get("teaching_score_percentile_current_filter")
    research_pct = context.get("research_score_percentile_current_filter")
    context["teaching_research_position"] = cluster_label(teaching_pct, research_pct)
    return context



def dimension_scores(row: pd.Series) -> dict[str, float]:
    return {
        "Teaching": float(row.get("teaching_score", 0) or 0),
        "Placement": float(row.get("placement_score", 0) or 0),
        "Research": float(row.get("research_score", 0) or 0),
        "Financial": float(row.get("financial_score", 0) or 0),
    }


def classify_profile(row: pd.Series) -> str:
    dims = dimension_scores(row)
    values = list(dims.values())
    strongest = max(dims, key=dims.get)
    weakest = min(dims, key=dims.get)
    spread = max(values) - min(values)
    if spread <= 10:
        return "Balanced profile"
    if strongest == "Research" and dims["Research"] - sorted(values)[-2] >= 8:
        return "Research-oriented profile"
    if strongest == "Teaching" and dims["Teaching"] - sorted(values)[-2] >= 8:
        return "Teaching-oriented profile"
    if strongest == "Placement" and dims["Placement"] - sorted(values)[-2] >= 8:
        return "Placement-oriented profile"
    if weakest == "Financial" and max(values) - dims["Financial"] >= 12:
        return "Finance-constrained profile"
    return "Mixed profile"


def strengths_weaknesses(row: pd.Series, comparison_data: pd.DataFrame) -> tuple[list[str], list[str]]:
    strengths: list[str] = []
    weaknesses: list[str] = []
    dims = {
        "Teaching": "teaching_score",
        "Placement": "placement_score",
        "Research": "research_score",
        "Financial": "financial_score",
        "Overall": "overall_score",
    }
    for label, col in dims.items():
        if col in comparison_data.columns and col in row.index:
            pct = percentile_position(comparison_data, col, row.get(col))
            if pct is not None and pct >= 70:
                strengths.append(f"{label} score is in the upper range of the current group ({pct:.0f}th percentile).")
            elif pct is not None and pct <= 30:
                weaknesses.append(f"{label} score is in the lower range of the current group ({pct:.0f}th percentile).")

    dim_values = dimension_scores(row)
    strongest = max(dim_values, key=dim_values.get)
    weakest = min(dim_values, key=dim_values.get)
    if dim_values[strongest] - dim_values[weakest] >= 12:
        strengths.append(f"The strongest dimension is {strongest.lower()} ({dim_values[strongest]:.1f}).")
        weaknesses.append(f"The weakest dimension is {weakest.lower()} ({dim_values[weakest]:.1f}), creating an uneven profile.")

    if "dea_vrs_efficiency_100" in row.index and "dea_vrs_efficiency_100" in comparison_data.columns:
        dea_pct = percentile_position(comparison_data, "dea_vrs_efficiency_100", row.get("dea_vrs_efficiency_100"))
        if dea_pct is not None and dea_pct >= 70:
            strengths.append(f"DEA-VRS efficiency is above most peers in the current group ({dea_pct:.0f}th percentile).")
        elif dea_pct is not None and dea_pct <= 30:
            weaknesses.append(f"DEA-VRS efficiency is below the central peer group ({dea_pct:.0f}th percentile).")

    if not strengths:
        strengths.append("No clear upper-tail strength is visible; the profile is mainly central/moderate.")
    if not weaknesses:
        weaknesses.append("No severe lower-tail weakness is visible in the current filtered group.")
    return strengths[:4], weaknesses[:4]


def percentile_badge(data: pd.DataFrame, column: str, value: Any) -> str:
    pct = percentile_position(data, column, value)
    if pct is None:
        return "Percentile n/a"
    return f"{pct:.0f}th percentile ({position_label(pct)})"


def make_ranking_context(filtered_data: pd.DataFrame, metric_col: str, metric_label: str, selected_row: pd.Series, top_n: int) -> dict:
    ranked = filtered_data.sort_values(metric_col, ascending=False).reset_index(drop=True)
    selected_rank = int(ranked.index[ranked["university"].eq(selected_row["university"])][0] + 1) if selected_row["university"] in ranked["university"].values else None
    return {
        "view_type": "Ranking Explorer",
        "year": clean_value(selected_row.get("year")),
        "metric": metric_label,
        "metric_column": metric_col,
        "number_of_universities": int(filtered_data["university"].nunique()),
        "selected_university": clean_value(selected_row.get("university")),
        "selected_value": clean_value(selected_row.get(metric_col)),
        "selected_rank_in_current_filter": selected_rank,
        "selected_percentile_current_filter": clean_value(percentile_position(filtered_data, metric_col, selected_row.get(metric_col))),
        "top_universities": ranked[["university", metric_col]].head(top_n).to_dict("records"),
        "bottom_universities": ranked[["university", metric_col]].tail(top_n).to_dict("records"),
    }


def make_change_context(full_data: pd.DataFrame, selected_university: str, start_year: int, end_year: int) -> dict:
    rows = full_data[(full_data["university"].eq(selected_university)) & (full_data["year"].isin([start_year, end_year]))].copy()
    context = {
        "view_type": "Time Dynamics / What Changed",
        "university": selected_university,
        "start_year": int(start_year),
        "end_year": int(end_year),
    }
    if rows["year"].nunique() < 2:
        context["note"] = "Selected university does not have both start and end year observations."
        return context
    start = rows[rows["year"].eq(start_year)].iloc[0]
    end = rows[rows["year"].eq(end_year)].iloc[0]
    change_cols = ["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"]
    if "dea_vrs_efficiency_100" in full_data.columns:
        change_cols.append("dea_vrs_efficiency_100")
    changes = {col: clean_value(end.get(col) - start.get(col)) for col in change_cols if col in rows.columns}
    largest_driver = max(changes, key=lambda k: abs(float(changes[k] or 0))) if changes else None
    context.update(
        {
            "start_values": {col: clean_value(start.get(col)) for col in change_cols if col in rows.columns},
            "end_values": {col: clean_value(end.get(col)) for col in change_cols if col in rows.columns},
            "changes_end_minus_start": changes,
            "largest_change_dimension": largest_driver,
        }
    )
    return context


def make_dea_context(filtered_data: pd.DataFrame, selected_row: pd.Series) -> dict:
    context = selected_context(selected_row, "DEA Efficiency Explorer")
    for col in ["dea_vrs_efficiency_100", "dea_crs_efficiency_100", "dea_scale_efficiency_100", "overall_score", "teaching_score", "placement_score", "research_score"]:
        if col in filtered_data.columns and col in selected_row.index:
            series = numeric_series(filtered_data, col)
            context[f"{col}_median_current_filter"] = clean_value(series.median()) if not series.empty else None
            context[f"{col}_percentile_current_filter"] = clean_value(percentile_position(filtered_data, col, selected_row.get(col)))
            context[f"{col}_position_label"] = position_label(context[f"{col}_percentile_current_filter"])
    return context



def _evidence_value_text(value: Any, decimals: int | None = None, suffix: str = "") -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "n/a"
    if isinstance(value, bool):
        text = str(value)
    elif isinstance(value, (int, float)):
        if decimals is None:
            decimals = 0 if isinstance(value, int) else 1
        text = f"{float(value):,.{decimals}f}"
    else:
        text = str(value)
    return f"{text}{suffix}"


def build_evidence_items(context: dict) -> list[dict]:
    """Create a compact evidence registry that links AI statements back to visible dashboard data."""
    items: list[dict] = []

    def add(label: str, value: Any, source: str, field: str = "", anchor: str = "source-current-view", decimals: int | None = None, suffix: str = "") -> None:
        if value is None:
            return
        if isinstance(value, float) and pd.isna(value):
            return
        items.append(
            {
                "id": f"E{len(items) + 1}",
                "label": label,
                "value": clean_value(value),
                "display_value": _evidence_value_text(value, decimals, suffix),
                "source": source,
                "field": field,
                "anchor": anchor,
            }
        )

    view = context.get("view_type")

    if view == "Overview":
        add("Universities in current filter", context.get("number_of_universities"), "Overview summary", "number_of_universities", "source-overview-summary", 0)
        add("Average overall score", context.get("average_overall_score"), "Overview profile", "average_overall_score", "source-overview-profile", 1)
        add("Average teaching score", context.get("average_teaching_score"), "Overview profile", "average_teaching_score", "source-overview-profile", 1)
        add("Average placement score", context.get("average_placement_score"), "Overview profile", "average_placement_score", "source-overview-profile", 1)
        add("Average research score", context.get("average_research_score"), "Overview profile", "average_research_score", "source-overview-profile", 1)
        add("Average financial score", context.get("average_financial_score"), "Overview profile", "average_financial_score", "source-overview-profile", 1)
        add("Average profile spread", context.get("average_profile_range"), "Overview profile", "average_profile_range", "source-overview-profile", 1)
        add("Average DEA-VRS efficiency", context.get("average_dea_vrs_efficiency"), "Overview profile", "average_dea_vrs_efficiency", "source-overview-profile", 1)

    elif view == "University Comparison" and context.get("university_a") and context.get("university_b"):
        a = context["university_a"]
        b = context["university_b"]
        gaps = context.get("score_gaps_a_minus_b", {})
        add(f"{a.get('university')} overall score", a.get("overall_score"), "Comparison summary", "university_a.overall_score", "source-comparison-summary", 1)
        add(f"{b.get('university')} overall score", b.get("overall_score"), "Comparison summary", "university_b.overall_score", "source-comparison-summary", 1)
        add("Overall score gap (A-B)", gaps.get("overall_gap_a_minus_b"), "Score gaps", "score_gaps_a_minus_b.overall_gap_a_minus_b", "source-comparison-scores", 1)
        add("Teaching score gap (A-B)", gaps.get("teaching_gap_a_minus_b"), "Score gaps", "score_gaps_a_minus_b.teaching_gap_a_minus_b", "source-comparison-scores", 1)
        add("Placement score gap (A-B)", gaps.get("placement_gap_a_minus_b"), "Score gaps", "score_gaps_a_minus_b.placement_gap_a_minus_b", "source-comparison-scores", 1)
        add("Research score gap (A-B)", gaps.get("research_gap_a_minus_b"), "Score gaps", "score_gaps_a_minus_b.research_gap_a_minus_b", "source-comparison-scores", 1)
        add("Financial score gap (A-B)", gaps.get("financial_gap_a_minus_b"), "Score gaps", "score_gaps_a_minus_b.financial_gap_a_minus_b", "source-comparison-scores", 1)

    elif view == "Finance Explorer":
        add(context.get("finance_x_axis", "Selected financial indicator"), context.get("selected_x_value"), "Selected scatterplot point", "selected_x_value", "source-finance-scatter", 2)
        add(context.get("score_y_axis", "Selected score"), context.get("selected_y_value"), "Selected scatterplot point", "selected_y_value", "source-finance-scatter", 1)
        add("X-axis median in current filter", context.get("x_axis_median_current_filter"), "Finance scatterplot benchmark", "x_axis_median_current_filter", "source-finance-scatter", 2)
        add("Y-axis median in current filter", context.get("y_axis_median_current_filter"), "Finance scatterplot benchmark", "y_axis_median_current_filter", "source-finance-scatter", 1)
        add("X-axis percentile in current filter", context.get("x_axis_percentile_current_filter"), "Finance scatterplot position", "x_axis_percentile_current_filter", "source-finance-scatter", 0, "th percentile")
        add("Y-axis percentile in current filter", context.get("y_axis_percentile_current_filter"), "Finance scatterplot position", "y_axis_percentile_current_filter", "source-finance-scatter", 0, "th percentile")
        add("FFO per student", context.get("ffo_per_student"), "Finance summary cards", "ffo_per_student", "source-finance-summary", 0)
        add("Operating cost per student", context.get("operating_cost_per_student"), "Finance summary cards", "operating_cost_per_student", "source-finance-summary", 0)
        add("Personnel cost share", context.get("personnel_cost_share"), "Finance summary cards", "personnel_cost_share", "source-finance-summary", 2)
        add("Public revenue share", context.get("public_revenue_share"), "Finance summary cards", "public_revenue_share", "source-finance-summary", 2)

    elif view == "DEA Efficiency Explorer":
        add("DEA-VRS efficiency", context.get("dea_vrs_efficiency_100"), "DEA summary", "dea_vrs_efficiency_100", "source-dea-summary", 1)
        add("DEA-CRS efficiency", context.get("dea_crs_efficiency_100"), "DEA summary", "dea_crs_efficiency_100", "source-dea-summary", 1)
        add("Scale efficiency", context.get("dea_scale_efficiency_100"), "DEA summary", "dea_scale_efficiency_100", "source-dea-summary", 1)
        add("Overall score", context.get("overall_score"), "DEA vs performance context", "overall_score", "source-dea-summary", 1)
        add("DEA-VRS percentile in current filter", context.get("dea_vrs_efficiency_100_percentile_current_filter"), "DEA current-filter position", "dea_vrs_efficiency_100_percentile_current_filter", "source-dea-scatter", 0, "th percentile")
        add("Efficiency category", context.get("efficiency_category"), "DEA summary", "efficiency_category", "source-dea-summary")

    elif view == "Ranking Explorer":
        add(context.get("metric", "Selected ranking metric"), context.get("selected_value"), "Ranking chart", "selected_value", "source-ranking-chart", 1)
        add("Rank in current filter", context.get("selected_rank_in_current_filter"), "Ranking summary", "selected_rank_in_current_filter", "source-ranking-summary", 0)
        add("Percentile in current filter", context.get("selected_percentile_current_filter"), "Ranking summary", "selected_percentile_current_filter", "source-ranking-summary", 0, "th percentile")
        add("Universities in current ranking", context.get("number_of_universities"), "Ranking summary", "number_of_universities", "source-ranking-summary", 0)

    elif view == "Time Dynamics / What Changed":
        starts = context.get("start_values", {})
        ends = context.get("end_values", {})
        changes = context.get("changes_end_minus_start", {})
        add(f"Overall score in {context.get('start_year')}", starts.get("overall_score"), "Time comparison", "start_values.overall_score", "source-time-change", 1)
        add(f"Overall score in {context.get('end_year')}", ends.get("overall_score"), "Time comparison", "end_values.overall_score", "source-time-change", 1)
        add("Overall score change", changes.get("overall_score"), "Change by dimension", "changes_end_minus_start.overall_score", "source-time-change", 1)
        largest = context.get("largest_change_dimension")
        if largest:
            add(f"Largest change: {str(largest).replace('_', ' ')}", changes.get(largest), "Change by dimension", f"changes_end_minus_start.{largest}", "source-time-change", 1)
        add("Teaching score change", changes.get("teaching_score"), "Change by dimension", "changes_end_minus_start.teaching_score", "source-time-change", 1)
        add("Research score change", changes.get("research_score"), "Change by dimension", "changes_end_minus_start.research_score", "source-time-change", 1)
        add("Financial score change", changes.get("financial_score"), "Change by dimension", "changes_end_minus_start.financial_score", "source-time-change", 1)

    elif view == "Teaching and Research":
        add("Teaching score", context.get("teaching_score"), "Teaching and research profile", "teaching_score", "source-teaching-summary", 1)
        add("Placement score", context.get("placement_score"), "Teaching and research profile", "placement_score", "source-teaching-summary", 1)
        add("Research score", context.get("research_score"), "Teaching and research profile", "research_score", "source-teaching-summary", 1)
        add("Teaching percentile in current filter", context.get("teaching_score_percentile_current_filter"), "Teaching vs research positioning", "teaching_score_percentile_current_filter", "source-teaching-research", 0, "th percentile")
        add("Research percentile in current filter", context.get("research_score_percentile_current_filter"), "Teaching vs research positioning", "research_score_percentile_current_filter", "source-teaching-research", 0, "th percentile")
        add("Second-year retention", context.get("second_year_retention_pct"), "Teaching indicators", "second_year_retention_pct", "source-teaching-indicators", 1, "%")
        add("Publications per teaching staff", context.get("publications_per_teaching_staff"), "Research indicators", "publications_per_teaching_staff", "source-research-indicators", 2)

    elif view == "Data and Methodology":
        add("Dataset scope", context.get("dataset_scope"), "Data and Methodology", "dataset_scope", "source-methodology")
        add("Current filtered universities", context.get("current_filter_count"), "Data and Methodology", "current_filter_count", "source-methodology", 0)
        add("Score definition", context.get("score_definition"), "Data and Methodology", "score_definition", "source-methodology")

    else:  # University Profile and other single-university pages
        add("Overall score", context.get("overall_score"), "Dimension profile", "overall_score", "source-profile-dimensions", 1)
        add("Overall rank", context.get("overall_rank_year"), "University profile summary", "overall_rank_year", "source-profile-summary", 0)
        add("Teaching score", context.get("teaching_score"), "Dimension profile", "teaching_score", "source-profile-dimensions", 1)
        add("Placement score", context.get("placement_score"), "Dimension profile", "placement_score", "source-profile-dimensions", 1)
        add("Research score", context.get("research_score"), "Dimension profile", "research_score", "source-profile-dimensions", 1)
        add("Financial score", context.get("financial_score"), "Dimension profile", "financial_score", "source-profile-dimensions", 1)
        add("National average overall score", context.get("national_avg_overall_score"), "Benchmark comparison", "national_avg_overall_score", "source-profile-benchmark", 1)
        add("Macro-area average overall score", context.get("macro_area_avg_overall_score"), "Benchmark comparison", "macro_area_avg_overall_score", "source-profile-benchmark", 1)

    return items


def extract_evidence_ids(answer: str, evidence_items: list[dict]) -> list[str]:
    valid = {item["id"] for item in evidence_items}
    seen: list[str] = []
    for evidence_id in re.findall(r"\[(E\d+)\]", answer or ""):
        if evidence_id in valid and evidence_id not in seen:
            seen.append(evidence_id)
    return seen


def linkify_evidence_citations(answer: str, evidence_items: list[dict]) -> str:
    """Render evidence IDs as same-tab links to the dashboard source.

    Raw HTML anchors are used instead of Markdown links because Streamlit may open
    Markdown links in a new browser tab. target="_self" keeps navigation inside
    the current dashboard tab.
    """
    item_map = {item["id"]: item for item in evidence_items}

    def repl(match: re.Match) -> str:
        evidence_id = match.group(1)
        item = item_map.get(evidence_id)
        if not item:
            return match.group(0)
        url = html.escape(evidence_deep_link(item), quote=True)
        return (
            f"<a class='companion-evidence-link' href='{url}' target='_self' "
            f"title='Show this evidence in the dashboard'>[{evidence_id}]</a>"
        )

    return re.sub(r"\[(E\d+)\]", repl, answer or "")


def evidence_deep_link(item: dict) -> str:
    anchor = item.get("anchor", "source-current-view")
    params = _navigation_state_params()
    params["focus_field"] = str(item.get("field", ""))
    # urlencode preserves spaces, slashes and comparison labels safely while keeping
    # the browser in the same dashboard state after navigation.
    return f"?{urlencode(params)}#{anchor}"


def _same_tab_value_link(display: str, item: dict) -> str:
    """Create an inline evidence value link that stays in the current browser tab."""
    url = html.escape(evidence_deep_link(item), quote=True)
    safe_display = html.escape(display)
    label = html.escape(str(item.get("label", "dashboard evidence")), quote=True)
    return (
        f"<a class='companion-data-link' href='{url}' target='_self' "
        f"title='Show {label} in the dashboard'>{safe_display}</a>"
    )


def linkify_evidence_values(answer: str, evidence_items: list[dict]) -> str:
    """Turn evidence-backed numeric values in Companion prose into same-tab deep links."""
    result = answer or ""
    numeric_items = [item for item in evidence_items if item.get("value") is not None and not isinstance(item.get("value"), str)]

    # First, pair each number with its nearby evidence ID. This disambiguates repeated values such as 100.0.
    for item in numeric_items:
        display = str(item.get("display_value", ""))
        if not display:
            continue
        eid = re.escape(str(item.get("id", "")))
        pattern = rf"(?<![\w\[]){re.escape(display)}(?![\w\]])(?=[^\n]{{0,80}}\[{eid}\])"
        result = re.sub(pattern, lambda _m, d=display, it=item: _same_tab_value_link(d, it), result)

    # Then link values that are unique in the evidence registry even if the model placed the evidence ID farther away.
    counts: dict[str, int] = {}
    for item in numeric_items:
        display = str(item.get("display_value", ""))
        counts[display] = counts.get(display, 0) + 1
    for item in numeric_items:
        display = str(item.get("display_value", ""))
        if not display or counts.get(display) != 1:
            continue
        pattern = rf"(?<![\w\[]){re.escape(display)}(?![\w\]])"
        # Avoid replacing a value that is already inside one of our HTML anchors.
        if f">{html.escape(display)}<" in result:
            continue
        result = re.sub(pattern, lambda _m, d=display, it=item: _same_tab_value_link(d, it), result)
    return result


def normalize_visual_focus_field(field: str | None) -> str | None:
    if not field:
        return None
    mapping = {
        "score_gaps_a_minus_b.overall_gap_a_minus_b": "overall_score",
        "score_gaps_a_minus_b.teaching_gap_a_minus_b": "teaching_score",
        "score_gaps_a_minus_b.placement_gap_a_minus_b": "placement_score",
        "score_gaps_a_minus_b.research_gap_a_minus_b": "research_score",
        "score_gaps_a_minus_b.financial_gap_a_minus_b": "financial_score",
        "changes_end_minus_start.overall_score": "overall_score",
        "changes_end_minus_start.teaching_score": "teaching_score",
        "changes_end_minus_start.placement_score": "placement_score",
        "changes_end_minus_start.research_score": "research_score",
        "changes_end_minus_start.financial_score": "financial_score",
        "changes_end_minus_start.dea_vrs_efficiency_100": "dea_vrs_efficiency_100",
        "start_values.overall_score": "overall_score",
        "end_values.overall_score": "overall_score",
    }
    return mapping.get(field, field)


def current_focus_field() -> str | None:
    try:
        value = st.query_params.get("focus_field")
    except Exception:
        value = None
    if isinstance(value, list):
        value = value[-1] if value else None
    return str(value) if value else None


def _query_param_scalar(name: str) -> str | None:
    try:
        value = st.query_params.get(name)
    except Exception:
        value = None
    if isinstance(value, list):
        value = value[-1] if value else None
    return str(value) if value not in (None, "") else None


def _safe_widget_restore(key: str, options: list, query_name: str, fallback):
    """Restore a widget value after an evidence deep-link reload without overriding live state."""
    if key in st.session_state:
        return
    requested = _query_param_scalar(query_name)
    if requested is not None:
        for option in options:
            if str(option) == requested:
                st.session_state[key] = option
                return
    st.session_state[key] = fallback


def _navigation_state_params() -> dict[str, str]:
    """Serialize the visible dashboard state into evidence links.

    Streamlit deep links can trigger a browser reload. Carrying the navigation state in
    the URL lets the app restore the same page and sidebar selection immediately.
    """
    params: dict[str, str] = {}
    mapping = {
        "page": "dashboard_page_selector",
        "year": "year_selector",
        "macro_area": "macro_area_selector",
        "region": "region_selector",
        "size_class": "size_class_selector",
        "university": "university_selector",
    }
    for query_name, state_key in mapping.items():
        value = st.session_state.get(state_key)
        if value not in (None, ""):
            params[query_name] = str(value)
    return params


def page_focus_field(page_name: str) -> str | None:
    deep = current_focus_field()
    chart_meta = st.session_state.get("chart_focus_meta") or {}
    if chart_meta.get("source_page") == page_name and chart_meta.get("field"):
        return normalize_visual_focus_field(str(chart_meta.get("field")))
    return normalize_visual_focus_field(deep)


def render_focus_callout(field: str, label: str, value: Any) -> None:
    if current_focus_field() != field:
        return
    safe_label = html.escape(str(label))
    safe_value = html.escape(_evidence_value_text(value, 2 if isinstance(value, float) and abs(value) < 10 else 1))
    st.markdown(
        f"<div class='focus-callout'>Linked from the Analysis Companion: <b>{safe_label}</b> = "
        f"<span class='focus-value'>{safe_value}</span>. The corresponding mark is outlined in black below.</div>",
        unsafe_allow_html=True,
    )


def ensure_evidence_references(answer: str, evidence_items: list[dict]) -> str:
    if not evidence_items:
        return answer
    if extract_evidence_ids(answer, evidence_items):
        return answer
    refs = " ".join(f"[{item['id']}]" for item in evidence_items[:6])
    return f"{answer.rstrip()}\n\n**Linked evidence from the current view:** {refs}"


def render_evidence_registry(evidence_items: list[dict], answer: str = "") -> None:
    if not evidence_items:
        return
    cited = extract_evidence_ids(answer, evidence_items)
    shown = [item for item in evidence_items if item["id"] in cited] if cited else evidence_items[:6]
    st.markdown("#### Evidence linked to the interpretation")
    st.caption("Both evidence IDs and evidence values are clickable. They return to the corresponding dashboard section; supported charts outline the linked mark in black.")
    for item in shown:
        label = html.escape(str(item["label"]))
        value = html.escape(str(item["display_value"]))
        source = html.escape(str(item["source"]))
        field = html.escape(str(item.get("field", "")))
        deep_link = evidence_deep_link(item)
        st.markdown(
            f"<div class='evidence-card'><span class='evidence-id'>{item['id']}</span>"
            f"<b>{label}</b><br><a href='{deep_link}' target='_self'><span class='focus-value'>{value}</span></a><br>"
            f"<span class='small-note'>Source: {source} · field: {field}</span></div>",
            unsafe_allow_html=True,
        )



def scroll_to_evidence_anchor(anchor: str | None) -> None:
    """After a deep-link rerun, return the browser to the linked dashboard section.

    Streamlit rebuilds the page after query-parameter navigation, so the browser can
    miss the URL fragment before the target element exists. A tiny client-side helper
    retries after rendering and scrolls the parent document to the evidence anchor.
    """
    if not anchor:
        return

    anchor_json = json.dumps(str(anchor))
    components.html(
        f"""
        <script>
        (() => {{
            const anchorId = {anchor_json};
            let attempts = 0;
            const maxAttempts = 24;

            function locateAndScroll() {{
                attempts += 1;
                try {{
                    const doc = window.parent.document;
                    const el = doc.getElementById(anchorId);
                    if (el) {{
                        // Keep the linked chart comfortably in view rather than placing
                        // the zero-height anchor at the very top edge of the browser.
                        el.scrollIntoView({{behavior: 'smooth', block: 'center', inline: 'nearest'}});

                        // Briefly pulse the Streamlit block that contains the linked section.
                        const block = el.closest('[data-testid="stVerticalBlock"]') || el.parentElement;
                        if (block) {{
                            block.classList.add('evidence-scroll-pulse');
                            window.setTimeout(() => block.classList.remove('evidence-scroll-pulse'), 1800);
                        }}
                        return true;
                    }}
                }} catch (err) {{
                    // Fallback below handles environments where iframe-parent DOM access
                    // is temporarily unavailable.
                }}

                if (attempts < maxAttempts) {{
                    window.setTimeout(locateAndScroll, 125);
                }} else {{
                    try {{ window.parent.location.hash = anchorId; }} catch (err) {{}}
                }}
                return false;
            }}

            // Streamlit/Altair elements are created asynchronously. Multiple scheduled
            // attempts make the scroll robust even when the chart finishes rendering late.
            window.setTimeout(locateAndScroll, 40);
            window.setTimeout(locateAndScroll, 260);
            window.setTimeout(locateAndScroll, 700);
        }})();
        </script>
        """,
        height=0,
        scrolling=False,
    )


def gap_direction(gap: float, threshold: float = 5.0) -> str:
    if gap >= threshold:
        return "clearly above"
    if gap <= -threshold:
        return "clearly below"
    return "close to"


def score_band_analysis(value: float) -> str:
    if value >= 75:
        return "strong"
    if value >= 65:
        return "above-average"
    if value >= 55:
        return "moderate"
    if value >= 45:
        return "relatively weak"
    return "weak"


def spread_analysis(values: list[float]) -> str:
    spread = max(values) - min(values)
    if spread >= 25:
        return "uneven profile with large differences between dimensions"
    if spread >= 12:
        return "mixed profile with visible strengths and weaker areas"
    return "balanced profile with no very large gap between dimensions"


def local_overview_interpretation(context: dict, user_question: str | None = None) -> str:
    top = context.get("top_universities", [])
    low = context.get("lowest_universities", [])
    top_names = ", ".join([x["university"] for x in top])
    low_names = ", ".join([x["university"] for x in low])
    avg_overall = float(context.get("average_overall_score") or 0)
    avg_teaching = float(context.get("average_teaching_score") or 0)
    avg_placement = float(context.get("average_placement_score") or 0)
    avg_research = float(context.get("average_research_score") or 0)
    avg_financial = float(context.get("average_financial_score") or 0)
    dims = {"teaching": avg_teaching, "placement": avg_placement, "research": avg_research, "financial": avg_financial}
    strongest = max(dims, key=dims.get)
    weakest = min(dims, key=dims.get)
    spread = max(dims.values()) - min(dims.values())
    q = f"\n\n**User question**\n\n{user_question}" if user_question else ""
    return f"""
**Overview analysis**

The filtered system contains **{context.get('number_of_universities')} universities** in **{context.get('year')}**. The average overall score is **{avg_overall:.1f}**, which indicates a **{score_band_analysis(avg_overall)}** system-level profile for the current selection.

**Main pattern, not just numbers**

The average profile is strongest in **{strongest}** and weakest in **{weakest}**. The gap between the strongest and weakest average dimensions is **{spread:.1f} points**, so this filtered group looks like a **{spread_analysis(list(dims.values()))}**. This means that the current selection should not be interpreted through the overall score alone: the dashboard suggests that one dimension is shaping the profile more than the others.

**Distribution reading**

The top universities in this view are **{top_names}**, while the lower end of the distribution includes **{low_names}**. This does not mean that the lower-ranked universities are globally weak; it means that, under the current filters and normalized dashboard scores, their multidimensional profile is less favorable than the group leaders.

**What the page suggests**

The most useful next step is to compare one top university with one lower-scoring university. This can reveal whether the difference is driven by research, finance, teaching, or placement rather than by a general performance gap.

**Limit**

This is an exploratory system-level reading. It does not explain causes and does not produce policy recommendations.{q}
"""


def local_profile_interpretation(context: dict, user_question: str | None = None) -> str:
    overall = float(context.get("overall_score") or 0)
    teaching = float(context.get("teaching_score") or 0)
    placement = float(context.get("placement_score") or 0)
    research = float(context.get("research_score") or 0)
    financial = float(context.get("financial_score") or 0)
    national = float(context.get("national_avg_overall_score") or 0)
    macro = float(context.get("macro_area_avg_overall_score") or 0)
    rank = int(context.get("overall_rank_year") or 0)
    dims = {"teaching": teaching, "placement": placement, "research": research, "financial": financial}
    strongest = max(dims, key=dims.get)
    weakest = min(dims, key=dims.get)
    national_gap = overall - national
    macro_gap = overall - macro
    spread = max(dims.values()) - min(dims.values())
    q = f"\n\n**User question**\n\n{user_question}\n\nThe answer should use only the University Profile indicators and benchmark bars." if user_question else ""
    return f"""
**University Profile analysis**

**{context.get('university')}** in **{context.get('year')}** has an overall score of **{overall:.1f}** and rank **{rank}/61**. This places it in a **{score_band_analysis(overall)}** position in the dashboard ranking.

**Benchmark interpretation**

Compared with the national average, the university is **{gap_direction(national_gap)}** the benchmark by **{abs(national_gap):.1f} points**. Compared with the macro-area average, it is **{gap_direction(macro_gap)}** the benchmark by **{abs(macro_gap):.1f} points**. This means the selected university is not only being evaluated in isolation: its profile is being read against both the national system and its territorial context.

**Profile shape**

The profile is strongest in **{strongest}** (**{dims[strongest]:.1f}**) and weakest in **{weakest}** (**{dims[weakest]:.1f}**). The internal spread is **{spread:.1f} points**, so the university shows a **{spread_analysis(list(dims.values()))}**. In practical terms, this page suggests that the university's overall position is driven more by its strongest dimension than by equal performance across all dimensions.

**Interpretive reading**

If **research** is the strongest dimension, the university is more research-oriented in this profile. If **financial** is the weakest, the dashboard suggests that financial conditions should be inspected before interpreting the overall score as fully balanced. If **placement** or **teaching** is weak, the student lifecycle indicators should be checked on the Teaching and Research page.

**Next visual step**

Use the time trend to check whether this profile is stable from 2020 to 2023. If the same strong/weak pattern remains across years, it is more meaningful than a one-year difference.

**Limit**

This is a descriptive dashboard interpretation. It identifies patterns and gaps, but it does not explain their causes.{q}
"""


def local_finance_interpretation(context: dict, user_question: str | None = None) -> str:
    uni = context.get("university")
    year = context.get("year")
    financial = float(context.get("financial_score") or 0)
    overall = float(context.get("overall_score") or 0)
    x_label = context.get("finance_x_axis", "selected financial indicator")
    y_label = context.get("score_y_axis", "selected score")
    x_value = context.get("selected_x_value")
    y_value = context.get("selected_y_value")
    x_median = context.get("x_axis_median_current_filter")
    y_median = context.get("y_axis_median_current_filter")
    x_pct = context.get("x_axis_percentile_current_filter")
    y_pct = context.get("y_axis_percentile_current_filter")
    x_position = context.get("x_axis_position_label", "not available")
    y_position = context.get("y_axis_position_label", "not available")
    cluster_position = context.get("scatter_cluster_position", "not available")
    ffo = context.get("ffo_per_student")
    op_cost = context.get("operating_cost_per_student")
    personnel = context.get("personnel_cost_share")
    public_share = context.get("public_revenue_share")
    student_share = context.get("student_contribution_share")
    perf_share = context.get("performance_quota_share")
    gap = financial - overall
    if gap >= 8:
        finance_reading = "the financial dimension is stronger than the overall profile"
    elif gap <= -8:
        finance_reading = "the financial dimension is weaker than the overall profile"
    else:
        finance_reading = "the financial dimension is broadly aligned with the overall profile"

    if cluster_position == "inside the main cluster":
        cluster_reading = (
            f"The selected point is inside the main cluster: it is not an outlier on this page. "
            f"The interpretation should therefore focus on moderate differences in funding, costs, and revenue structure rather than on an exceptional financial pattern."
        )
    elif cluster_position == "near the edge of the distribution":
        cluster_reading = (
            f"The selected point is close to the edge of the distribution. This makes the university more distinctive in the current filtered group, "
            f"so the selected financial indicator should be inspected carefully before generalizing the profile."
        )
    else:
        cluster_reading = (
            f"The selected point is outside the central band but not an extreme outlier. This suggests a visible but not exceptional financial-performance position."
        )

    q = f"\n\n**User question**\n\n{user_question}\n\nThe answer should focus on the Finance Explorer page only." if user_question else ""
    return f"""
**Finance Explorer analysis**

This page reads the financial position of **{uni}** in **{year}** through the selected scatterplot: **{x_label}** on the x-axis and **{y_label}** on the y-axis.

**Position in the scatterplot**

The selected university has **{format_number(x_value, 2)}** on the x-axis and **{format_number(y_value, 1)}** on the y-axis. Within the current filtered group, this places it in the **{x_position}** for **{x_label}** and in the **{y_position}** for **{y_label}**. The median values in the same filtered group are **{format_number(x_median, 2)}** for the x-axis and **{format_number(y_median, 1)}** for the y-axis.

**What this means analytically**

{cluster_reading} Its financial score is **{financial:.1f}**, while its overall score is **{overall:.1f}**; therefore, **{finance_reading}**. This is more informative than reading the financial score alone, because it shows whether finance is pulling the university's profile up, down, or moving together with the general performance profile.

**Financial structure reading**

The financial profile combines funding intensity, cost pressure, and revenue structure. FFO per student is **{format_number(ffo, 0)}**, operating cost per student is **{format_number(op_cost, 0)}**, and personnel cost share is **{format_number(personnel, 2)}**. Public revenue share is **{format_number(public_share, 2)}**, while student contribution share is **{format_number(student_share, 2)}**. This suggests whether the financial profile is more dependent on public funding, student contributions, or internal cost structure.

Performance quota share is **{format_number(perf_share, 2)}**. This should be read as an additional descriptive signal, not as proof that funding caused the observed score.

**Next visual step**

Change the x-axis from **{x_label}** to another financial indicator. If the university remains in the same relative position across several financial indicators, the interpretation is more robust; if it moves substantially, the financial reading depends strongly on the chosen indicator.

**Limit**

This page shows association and positioning only. It does not estimate the causal impact of funding or costs on performance.{q}
"""

def local_teaching_research_interpretation(context: dict, user_question: str | None = None) -> str:
    uni = context.get("university")
    year = context.get("year")
    teaching = float(context.get("teaching_score") or 0)
    placement = float(context.get("placement_score") or 0)
    research = float(context.get("research_score") or 0)
    retention = context.get("second_year_retention_pct")
    inactive = context.get("inactive_students_reversed_score")
    grad_std = context.get("graduation_within_standard_pct")
    employment = context.get("employment_index")
    pub_staff = context.get("publications_per_teaching_staff")
    cites_pub = context.get("citations_per_publication")
    h_index = context.get("h_index")
    hcr = context.get("highly_cited_researchers")
    ns = context.get("nature_science_articles")
    tr_gap = research - teaching
    rp_gap = research - placement
    teaching_pct = context.get("teaching_score_percentile_current_filter")
    research_pct = context.get("research_score_percentile_current_filter")
    placement_pct = context.get("placement_score_percentile_current_filter")
    tr_position = context.get("teaching_research_position", "not available")

    if tr_gap >= 10:
        orientation = "research-oriented"
        orientation_text = "research performance is visibly stronger than the teaching dimension"
    elif tr_gap <= -10:
        orientation = "teaching-oriented"
        orientation_text = "teaching is visibly stronger than research"
    else:
        orientation = "balanced between teaching and research"
        orientation_text = "teaching and research are relatively close"

    if tr_position == "inside the main cluster":
        position_text = "The university sits inside the main teaching-research cluster, so the profile is not an extreme outlier among the currently filtered institutions."
    elif tr_position == "near the edge of the distribution":
        position_text = "The university is near the edge of the teaching-research distribution, so its academic profile is relatively distinctive in the current filtered group."
    else:
        position_text = "The university is outside the central band but not an extreme case, indicating a visible specialization pattern rather than a fully typical profile."

    q = f"\n\n**User question**\n\n{user_question}\n\nThe answer should focus on the Teaching and Research page indicators." if user_question else ""
    return f"""
**Teaching and Research analysis**

**{uni}** in **{year}** shows a **{orientation}** academic profile. Its teaching score is **{teaching:.1f}**, placement score is **{placement:.1f}**, and research score is **{research:.1f}**. The gap between research and teaching is **{tr_gap:.1f} points**, so **{orientation_text}**.

**Position relative to the current group**

Within the current filtered group, the university is around the **{format_number(teaching_pct, 0)}th percentile** for teaching and the **{format_number(research_pct, 0)}th percentile** for research. {position_text} This is important because the page is not only showing the selected university's indicators; it is also showing whether its teaching-research combination is typical or distinctive compared with similar filtered universities.

**Student lifecycle interpretation**

The student-side profile should be read through retention, graduation, and employment together. Second-year retention is **{format_number(retention, 1)}**, graduation within standard duration is **{format_number(grad_std, 1)}**, and employment index is **{format_number(employment, 1)}**. The inactive-students reversed score is **{format_number(inactive, 1)}**, where a higher value is more favorable. If placement is lower than teaching, the dashboard suggests that the student pathway may look stronger during study progression than after graduation.

**Research intensity interpretation**

The research side is supported by publications per teaching staff (**{format_number(pub_staff, 2)}**), citations per publication (**{format_number(cites_pub, 2)}**), H-index (**{format_number(h_index, 1)}**), highly cited researchers (**{format_number(hcr, 1)}**), and Nature/Science articles (**{format_number(ns, 1)}**). A high research score should therefore be interpreted as a composite research-intensity signal rather than as a single publication count.

**Main analytical reading**

The page suggests that the profile should be interpreted through the balance between student outcomes and research intensity. If research is much higher than teaching or placement, the university appears more research-intensive than student-outcome-oriented in this dashboard view. If teaching and placement are close to research, the profile is more balanced and should not be reduced to a research-only interpretation.

**Next visual step**

Use the Teaching vs Research positioning chart to find universities with a similar balance, then compare one of them with **{uni}** on the University Comparison page.

**Limit**

This page identifies academic profile patterns. It does not explain why research, teaching, or placement indicators differ.{q}
"""

def local_single_interpretation(context: dict, user_question: str | None = None) -> str:
    return local_profile_interpretation(context, user_question)


def local_comparison_interpretation(context: dict, user_question: str | None = None) -> str:
    a = context["university_a"]
    b = context["university_b"]
    gaps = context["score_gaps_a_minus_b"]
    dims = {
        "overall": float(gaps["overall_gap_a_minus_b"]),
        "teaching": float(gaps["teaching_gap_a_minus_b"]),
        "placement": float(gaps["placement_gap_a_minus_b"]),
        "research": float(gaps["research_gap_a_minus_b"]),
        "financial": float(gaps["financial_gap_a_minus_b"]),
    }
    largest_dim = max(dims, key=lambda k: abs(dims[k]))
    overall_gap = dims["overall"]
    leader = a["university"] if overall_gap >= 0 else b["university"]
    follower = b["university"] if overall_gap >= 0 else a["university"]
    a_dims = [float(a["teaching_score"]), float(a["placement_score"]), float(a["research_score"]), float(a["financial_score"])]
    b_dims = [float(b["teaching_score"]), float(b["placement_score"]), float(b["research_score"]), float(b["financial_score"])]
    a_balance = spread_analysis(a_dims)
    b_balance = spread_analysis(b_dims)
    q = f"\n\n**User question**\n\n{user_question}\n\nThe answer should focus only on the two selected universities and the comparison charts." if user_question else ""
    return f"""
**University Comparison analysis**

In **{context.get('year')}**, **{leader}** has the higher overall score, leading **{follower}** by **{abs(overall_gap):.1f} points**. This is not just a ranking difference: the comparison page shows which dimension is responsible for the gap.

**Where the comparison is really different**

The largest visible difference is in **{largest_dim}**, with a gap of **{abs(dims[largest_dim]):.1f} points**. This suggests that the comparison should be interpreted mainly through **{largest_dim}**, rather than only through the overall score.

**Profile interpretation**

- **{a['university']}** has overall **{float(a['overall_score']):.1f}** and rank **{int(a['overall_rank_year'])}/61**. Its internal profile is a **{a_balance}**.
- **{b['university']}** has overall **{float(b['overall_score']):.1f}** and rank **{int(b['overall_rank_year'])}/61**. Its internal profile is a **{b_balance}**.

If one university has a higher research score but similar or weaker financial score, the comparison suggests a specialization pattern rather than uniformly better performance. If one university is higher across all dimensions, the difference is broader and more consistent.

**Next visual step**

Check the trend chart for 2020-2023. A one-year gap should be interpreted cautiously, while a persistent gap across years is a stronger descriptive pattern.

**Limit**

This is a descriptive comparison. It does not explain why one university is ahead and does not imply causal conclusions.{q}
"""


def local_methodology_interpretation(context: dict, user_question: str | None = None) -> str:
    q = f"\n\n**User question**\n\n{user_question}" if user_question else ""
    return f"""
**Data and methodology analysis**

This page explains how the dashboard should be interpreted. The dataset is a curated prototype dataset for Italian universities over **2020-2023**, designed for visual analytics rather than for causal estimation.

**Why this matters analytically**

The score fields are **dashboard-based normalized profile scores**, not DEA or SFA efficiency estimates. This means they are useful for comparison, visualization, and profile interpretation, but they should not be presented as final econometric efficiency measures.

**How to read the dashboard responsibly**

The dashboard is strongest when used to identify patterns: strong and weak dimensions, outliers, benchmark gaps, and differences between universities. It should not be used to claim that one variable causes another variable to improve or decline.

**Methodological value**

The main contribution is the transformation of a complex multidimensional dataset into an interactive visual analytics environment with an AI-assisted interpretation layer. This supports interpretation beyond static rankings and tables.

**Limit**

The dashboard supports exploration and explanation of visible patterns, but it does not produce causal conclusions or autonomous policy recommendations.{q}
"""


def local_ranking_interpretation(context: dict, user_question: str | None = None) -> str:
    metric = context.get("metric")
    selected = context.get("selected_university")
    selected_value = float(context.get("selected_value") or 0)
    rank = context.get("selected_rank_in_current_filter")
    n = context.get("number_of_universities")
    pct = context.get("selected_percentile_current_filter")
    top = context.get("top_universities", [])
    bottom = context.get("bottom_universities", [])
    top_names = ", ".join([x.get("university", "") for x in top[:3]])
    bottom_names = ", ".join([x.get("university", "") for x in bottom[:3]])
    q = f"\n\n**User question**\n\n{user_question}" if user_question else ""
    return f"""
**Ranking Explorer analysis**

This page ranks the current filtered group by **{metric}**. **{selected}** has a value of **{selected_value:.1f}** and is ranked **{rank}/{n}** in the current selection, around the **{format_number(pct, 0)}th percentile**.

**Analytical reading**

The ranking should be read as a dimension-specific benchmark, not as a complete evaluation of the university. If the selected metric is overall score, it summarizes the dashboard profile. If the selected metric is DEA-VRS efficiency, it reads the resource-to-output transformation instead of pure performance level. This distinction matters because a high-performing university is not always the most efficient one.

**Distribution reading**

The top end of the ranking includes **{top_names}**, while the lower end includes **{bottom_names}**. This helps identify whether the selected university is close to the leaders, in the central group, or closer to the lower tail.

**Recommended next check**

Switch the ranking metric from overall score to DEA-VRS efficiency. If the selected university keeps a similar position, its performance and efficiency stories are aligned. If the position changes, the dashboard reveals a difference between performance level and resource-adjusted efficiency.

**Limit**

This page is descriptive. Rankings depend on the selected metric, year, and filters.{q}
"""


def local_change_interpretation(context: dict, user_question: str | None = None) -> str:
    uni = context.get("university")
    start_year = context.get("start_year")
    end_year = context.get("end_year")
    changes = context.get("changes_end_minus_start", {})
    if not changes:
        return f"**Time dynamics analysis**\n\nThe dashboard cannot compute a change for **{uni}** because one of the selected years is missing."
    overall_change = float(changes.get("overall_score") or 0)
    largest = context.get("largest_change_dimension")
    largest_change = float(changes.get(largest) or 0) if largest else 0
    direction = "improved" if overall_change > 1 else "declined" if overall_change < -1 else "remained broadly stable"
    q = f"\n\n**User question**\n\n{user_question}" if user_question else ""
    return f"""
**Time dynamics analysis**

From **{start_year}** to **{end_year}**, **{uni}** **{direction}** in overall score, with a change of **{overall_change:+.1f} points**.

**What changed most**

The largest visible movement is in **{str(largest).replace('_', ' ')}**, with a change of **{largest_change:+.1f} points**. This suggests that the time trend should not be interpreted only through the overall score: the dashboard shows which dimension is driving the movement.

**Analytical interpretation**

If the overall score increased but one dimension declined, the university's profile became stronger but less balanced. If all dimensions moved in the same direction, the change is broader and more consistent. If DEA efficiency moved differently from the performance scores, the university's resource-adjusted position changed differently from its pure performance profile.

**Recommended next check**

Open the University Profile page and inspect whether the same strong/weak dimensions remain visible in the final year. Persistent patterns are more meaningful than a one-year fluctuation.

**Limit**

The page describes changes over time. It does not explain why the changes occurred.{q}
"""


def local_dea_interpretation(context: dict, user_question: str | None = None) -> str:
    uni = context.get("university")
    year = context.get("year")
    vrs = float(context.get("dea_vrs_efficiency_100") or 0)
    crs = float(context.get("dea_crs_efficiency_100") or 0)
    scale = float(context.get("dea_scale_efficiency_100") or 0)
    rank = context.get("dea_vrs_rank_year")
    category = context.get("efficiency_category")
    vrs_pct = context.get("dea_vrs_efficiency_100_percentile_current_filter")
    vrs_pos = context.get("dea_vrs_efficiency_100_position_label", "not available")
    overall = float(context.get("overall_score") or 0)
    overall_pct = context.get("overall_score_percentile_current_filter")
    if vrs >= 99:
        reading = "is on or very close to the DEA-VRS efficient frontier"
    elif vrs >= 90:
        reading = "is close to the efficient frontier, but still has some distance from the best peer combinations"
    elif vrs >= 80:
        reading = "has a moderate efficiency gap relative to the frontier"
    else:
        reading = "has a visible efficiency gap relative to the frontier"

    if scale >= 95:
        scale_reading = "scale efficiency is high, so the difference between CRS and VRS is limited"
    elif scale >= 85:
        scale_reading = "scale efficiency is moderate, so part of the gap may be connected with scale conditions"
    else:
        scale_reading = "scale efficiency is relatively low, so scale conditions should be inspected carefully"
    q = f"\n\n**User question**\n\n{user_question}" if user_question else ""
    return f"""
**DEA Efficiency Explorer analysis**

**{uni}** in **{year}** has a DEA-VRS efficiency score of **{vrs:.1f}/100** and DEA-VRS rank **{rank}/61**. In this specification, the university **{reading}**. Its efficiency category is **{category}**.

**Efficiency versus performance**

The university's overall dashboard score is **{overall:.1f}**, while its DEA-VRS efficiency is **{vrs:.1f}**. This comparison is important because overall score measures performance level, whereas DEA efficiency measures how well selected resources and cost indicators are transformed into teaching, placement, and research outputs.

Within the current filtered group, the DEA-VRS score is in the **{vrs_pos}** around the **{format_number(vrs_pct, 0)}th percentile**. The overall score is around the **{format_number(overall_pct, 0)}th percentile**. If these two percentiles differ, the dashboard suggests a difference between performance strength and resource-adjusted efficiency.

**Scale interpretation**

The DEA-CRS score is **{crs:.1f}**, and scale efficiency is **{scale:.1f}**. Therefore, **{scale_reading}**. VRS is the main score here because universities differ strongly in size and scale.

**Inputs and outputs used**

The DEA layer uses funding/cost/staff indicators as inputs and teaching, placement, and research scores as outputs. Financial score is not used as an output because financial variables already appear on the input side.

**Limit**

DEA results depend on the selected inputs and outputs. The score is a descriptive benchmarking measure, not a causal estimate and not a final policy judgment.{q}
"""


def generate_local_interpretation(context: dict, user_question: str | None = None) -> str:
    view_type = context.get("view_type")
    if view_type == "Overview":
        return local_overview_interpretation(context, user_question)
    if view_type == "University Profile":
        return local_profile_interpretation(context, user_question)
    if view_type == "University Comparison":
        return local_comparison_interpretation(context, user_question)
    if view_type == "Finance Explorer":
        return local_finance_interpretation(context, user_question)
    if view_type == "Teaching and Research":
        return local_teaching_research_interpretation(context, user_question)
    if view_type == "Ranking Explorer":
        return local_ranking_interpretation(context, user_question)
    if view_type == "Time Dynamics / What Changed":
        return local_change_interpretation(context, user_question)
    if view_type == "DEA Efficiency Explorer":
        return local_dea_interpretation(context, user_question)
    if view_type == "Data and Methodology":
        return local_methodology_interpretation(context, user_question)
    return local_single_interpretation(context, user_question)


def generate_ai_interpretation(context: dict, user_question: str | None = None) -> str:
    try:
        api_key = st.secrets.get("OPENAI_API_KEY", None)
    except Exception:
        api_key = None

    evidence_items = context.get("evidence_items", [])

    if OpenAI is None or not is_valid_api_key(api_key):
        local_answer = generate_local_interpretation(context, user_question)
        return ensure_evidence_references(local_answer, evidence_items)

    client = OpenAI(api_key=api_key.strip())
    prompt = {
        "task": f"Interpret the current dashboard page: {context.get('view_type')}.",
        "rules": [
            "Use only the provided dashboard context.",
            "Analyze only the currently selected page/view_type.",
            "Do not refer to other dashboard pages unless suggesting next exploration.",
            "Do not invent missing values.",
            "Do not make causal claims.",
            "Write in concise academic English.",
            "Return a page-specific interpretation: do not use a generic university profile unless the current page is University Profile.",
            "For Finance Explorer, focus on financial indicators and scatterplot axes. For Teaching and Research, focus on teaching/research indicators. For Overview, focus on filtered system-level patterns. For University Comparison, focus on the two selected universities. For Ranking Explorer, focus on ranking position and metric choice. For Time Dynamics, focus on changes between years. For DEA Efficiency Explorer, focus on DEA-VRS, CRS, scale efficiency and resource-to-output interpretation.",
            "Write analysis, not a list of visible numbers. Use a few key numbers only when they support an interpretation.",
            "Use computed percentiles, medians, gap labels, and cluster-position fields whenever they are available.",
            "The context contains an evidence_items registry. For every key quantitative or comparative claim, append one or more evidence IDs in square brackets, for example [E1] or [E2][E3]. Use only IDs that exist in evidence_items.",
            "Evidence IDs are not decorative citations: each one links the explanation back to the visible dashboard section from which the value came.",
            "If interaction_selection is present, treat the chart-clicked university as the explicit user-selected observation and keep the interpretation grounded in that selected mark.",
            "If interaction_focus is present, prioritize the clicked dimension or indicator and explain it in the context of the current page without losing the broader profile.",
            "If evidence_focus is present, explicitly address that evidence item because the user clicked its number in the previous Companion explanation.",
            "Explain what the pattern means for the current page: profile shape, trade-offs, benchmark position, specialization, outlier behavior, or balance between dimensions.",
            "Use phrases such as: this suggests, this indicates, this points to, this should be read as, but do not state causality.",
            "Return: analytical summary, interpretation of the main pattern, strengths/trade-offs, what to inspect next, and limitation.",
        ],
        "current_dashboard_context": context,
        "user_question": user_question,
    }
    try:
        response = client.responses.create(
            model="gpt-4.1-mini",
            input=json.dumps(prompt, ensure_ascii=True),
        )
        return ensure_evidence_references(response.output_text, evidence_items)
    except Exception:
        local_answer = generate_local_interpretation(context, user_question)
        return ensure_evidence_references(local_answer, evidence_items)


@st.cache_data(show_spinner=False, ttl=3600, max_entries=128)
def cached_auto_interpretation(context_json: str) -> str:
    """Cache automatic Companion prose across Streamlit sessions.

    Evidence-number navigation can reload the page in some browsers. This server-side
    cache prevents the same analytical state from calling the model again after that
    navigation. Pure navigation focus is intentionally excluded from context_json.
    """
    context = json.loads(context_json)
    return generate_ai_interpretation(context, None)


def score_bar_chart(score_df: pd.DataFrame, focus_field: str | None = None) -> alt.Chart:
    data = score_df.copy()
    data["Focused"] = data["Field"].eq(focus_field) if "Field" in data.columns else False
    bars = (
        alt.Chart(data)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("Dimension:N", sort=None, title="Dimension"),
            y=alt.Y("Score:Q", title="Score", scale=alt.Scale(domain=[0, 100])),
            color=alt.condition("datum.Focused", alt.value("#111111"), alt.Color("Dimension:N", legend=None)),
            stroke=alt.condition("datum.Focused", alt.value("#111111"), alt.value(None)),
            strokeWidth=alt.condition("datum.Focused", alt.value(3), alt.value(0)),
            tooltip=["Dimension", "Field", alt.Tooltip("Score:Q", format=".1f")],
        )
    )
    labels = alt.Chart(data).mark_text(dy=-8).encode(x=alt.X("Dimension:N", sort=None), y="Score:Q", text=alt.Text("Score:Q", format=".1f"))
    return (bars + labels).properties(height=300)


def benchmark_chart(row: pd.Series, focus_field: str | None = None) -> alt.Chart:
    benchmark_df = pd.DataFrame([
        {"Benchmark":"Selected university", "Field":"overall_score", "Overall score":float(row["overall_score"])},
        {"Benchmark":"National average", "Field":"national_avg_overall_score", "Overall score":float(row["national_avg_overall_score"])},
        {"Benchmark":"Macro-area average", "Field":"macro_area_avg_overall_score", "Overall score":float(row["macro_area_avg_overall_score"])},
    ])
    benchmark_df["Focused"] = benchmark_df["Field"].eq(focus_field)
    bars = alt.Chart(benchmark_df).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
        x=alt.X("Benchmark:N", sort=None, title="Benchmark"),
        y=alt.Y("Overall score:Q", title="Overall score", scale=alt.Scale(domain=[0, 100])),
        color=alt.condition("datum.Focused", alt.value("#111111"), alt.Color("Benchmark:N", legend=None)),
        stroke=alt.condition("datum.Focused", alt.value("#111111"), alt.value(None)),
        strokeWidth=alt.condition("datum.Focused", alt.value(3), alt.value(0)),
        tooltip=["Benchmark", "Field", alt.Tooltip("Overall score:Q", format=".1f")],
    )
    labels = alt.Chart(benchmark_df).mark_text(dy=-8).encode(x=alt.X("Benchmark:N", sort=None), y="Overall score:Q", text=alt.Text("Overall score:Q", format=".1f"))
    return (bars + labels).properties(height=300)


df = load_data()

st.title("AI-enhanced Visual Analytics Dashboard for Italian Universities")
st.caption("Interactive prototype for exploring teaching, placement, research, and financial profiles of Italian universities, 2020-2023.")

with st.sidebar:
    st.header("Filters")
    years = sorted(df["year"].unique())
    _safe_widget_restore("year_selector", years, "year", years[-1])
    year = st.selectbox("Year", years, key="year_selector")

    macro_area_options = ["All"] + sorted(df["macro_area"].dropna().unique())
    _safe_widget_restore("macro_area_selector", macro_area_options, "macro_area", "All")
    macro_area = st.selectbox("Macro-area", macro_area_options, key="macro_area_selector")

    region_base = df[df["year"] == year].copy()
    if macro_area != "All":
        region_base = region_base[region_base["macro_area"] == macro_area]
    region_options = ["All"] + sorted(region_base["region"].dropna().unique())
    if st.session_state.get("region_selector") not in region_options:
        st.session_state.pop("region_selector", None)
    _safe_widget_restore("region_selector", region_options, "region", "All")
    region = st.selectbox("Region", region_options, key="region_selector")

    size_base = region_base.copy()
    if region != "All":
        size_base = size_base[size_base["region"] == region]
    size_options = ["All"] + sorted(size_base["size_class"].dropna().unique())
    if st.session_state.get("size_class_selector") not in size_options:
        st.session_state.pop("size_class_selector", None)
    _safe_widget_restore("size_class_selector", size_options, "size_class", "All")
    size_class = st.selectbox("Size class", size_options, key="size_class_selector")

filtered = df[df["year"] == year].copy()
if macro_area != "All":
    filtered = filtered[filtered["macro_area"] == macro_area]
if region != "All":
    filtered = filtered[filtered["region"] == region]
if size_class != "All":
    filtered = filtered[filtered["size_class"] == size_class]

if filtered.empty:
    st.error("No universities match the selected filters. Please change the filters.")
    st.stop()

university_options = sorted(filtered["university"].unique())
st.session_state["_current_university_options"] = university_options
if "university_selector" not in st.session_state:
    requested_university = _query_param_scalar("university")
    st.session_state["university_selector"] = requested_university if requested_university in university_options else university_options[0]
    st.session_state["chart_selection_meta"] = None
elif st.session_state["university_selector"] not in university_options:
    st.session_state["university_selector"] = university_options[0]
    st.session_state["chart_selection_meta"] = None

with st.sidebar:
    university = st.selectbox(
        "University",
        university_options,
        key="university_selector",
        on_change=on_sidebar_university_change,
    )
    st.caption("Tip: interactive marks are available across the dashboard. Click university marks to change the active university; click dimension/indicator bars to focus the Companion on that metric.")

selected = filtered[filtered["university"] == university].iloc[0]

main_col, ai_col = st.columns([3.2, 1.1], gap="large")

pages = [
    "Overview",
    "University Profile",
    "University Comparison",
    "Finance Explorer",
    "DEA Efficiency Explorer",
    "Ranking Explorer",
    "Time Dynamics / What Changed",
    "Teaching and Research",
    "Data and Methodology",
]

with main_col:
    st.subheader(f"{university} - {year}")
    st.markdown(
        f"<span class='small-note'>Region: <b>{selected['region']}</b> | Macro-area: <b>{selected['macro_area']}</b> | Size class: <b>{selected['size_class']}</b></span>",
        unsafe_allow_html=True,
    )
    st.write("")
    requested_page = _query_param_scalar("page")
    # A focus deep link may reload the Streamlit document. Restore the originating
    # dashboard page before rendering so the user never falls back to Overview.
    if current_focus_field() and requested_page in pages:
        st.session_state["dashboard_page_selector"] = requested_page
    elif "dashboard_page_selector" not in st.session_state:
        st.session_state["dashboard_page_selector"] = requested_page if requested_page in pages else pages[0]
    elif st.session_state["dashboard_page_selector"] not in pages:
        st.session_state["dashboard_page_selector"] = pages[0]

    page = st.radio(
        "Dashboard page",
        pages,
        horizontal=True,
        label_visibility="collapsed",
        key="dashboard_page_selector",
        on_change=_clear_deep_focus,
    )
    st.divider()

    # Default active context, overwritten by page-specific contexts below.
    active_context = selected_context(selected, page)

    if page == "Overview":
        active_context = make_overview_context(filtered, year, macro_area, region, size_class)
        st.markdown("<div id='source-overview-summary'></div>", unsafe_allow_html=True)
        st.markdown("### System overview")
        st.caption("This page summarizes the universities that match the current filters and shows the overall distribution of dashboard-based profile scores.")
        o1, o2, o3, o4 = st.columns(4)
        o1.metric("Universities shown", f"{filtered['university'].nunique()}")
        o2.metric("Average overall", f"{filtered['overall_score'].mean():.1f}")
        o3.metric("Average teaching", f"{filtered['teaching_score'].mean():.1f}")
        o4.metric("Average research", f"{filtered['research_score'].mean():.1f}")

        top10 = filtered.sort_values("overall_score", ascending=False).head(10).copy()
        overview_university_selection = alt.selection_point(name="overview_university_selection", fields=["university"], on="click", clear="dblclick")
        top_chart = (
            alt.Chart(top10)
            .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
            .encode(
                x=alt.X("overall_score:Q", title="Overall score", scale=alt.Scale(domain=[0, 100])),
                y=alt.Y("university:N", sort="-x", title=None),
                color=alt.condition(alt.datum.university == university, alt.value("#111111"), alt.value("#4c78a8")),
                tooltip=["university", "region", alt.Tooltip("overall_score:Q", format=".1f")],
            )
            .add_params(overview_university_selection)
            .properties(height=360)
        )
        top_labels = (
            alt.Chart(top10)
            .mark_text(align="left", dx=5)
            .encode(x=alt.X("overall_score:Q"), y=alt.Y("university:N", sort="-x"), text=alt.Text("overall_score:Q", format=".1f"))
        )
        top_chart = top_chart + top_labels
        macro_avg = (
            filtered.groupby("macro_area", as_index=False)[["overall_score", "teaching_score", "research_score", "financial_score"]]
            .mean()
            .sort_values("overall_score", ascending=False)
        )
        macro_long = macro_avg.melt(id_vars="macro_area", var_name="Score type", value_name="Score")
        macro_chart = (
            alt.Chart(macro_long)
            .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
            .encode(
                x=alt.X("macro_area:N", title="Macro-area"),
                y=alt.Y("Score:Q", scale=alt.Scale(domain=[0, 100])),
                color=alt.Color("Score type:N", title="Score type"),
                tooltip=["macro_area", "Score type", alt.Tooltip("Score:Q", format=".1f")],
            )
            .properties(height=360)
        )
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Top 10 universities by overall score")
            st.caption("Click a university bar to make it the active dashboard selection.")
            st.altair_chart(top_chart, key="overview_top_chart", on_select=on_overview_top_select, selection_mode=["overview_university_selection"], width="stretch")
        with c2:
            st.markdown("<div id='source-overview-profile'></div>", unsafe_allow_html=True)
            st.markdown("#### Current filtered average profile")
            overview_focus = page_focus_field("Overview")
            if overview_focus == "average_profile_range":
                render_focus_callout("average_profile_range", "Average profile spread", active_context.get("average_profile_range"))
            avg_profile = pd.DataFrame([
                {"Dimension":"Overall", "Field":"average_overall_score", "Score":float(filtered["overall_score"].mean())},
                {"Dimension":"Teaching", "Field":"average_teaching_score", "Score":float(filtered["teaching_score"].mean())},
                {"Dimension":"Placement", "Field":"average_placement_score", "Score":float(filtered["placement_score"].mean())},
                {"Dimension":"Research", "Field":"average_research_score", "Score":float(filtered["research_score"].mean())},
                {"Dimension":"Financial", "Field":"average_financial_score", "Score":float(filtered["financial_score"].mean())},
            ])
            avg_profile["Focused"] = avg_profile["Field"].eq(overview_focus)
            avg_bars = alt.Chart(avg_profile).mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
                y=alt.Y("Dimension:N", sort=None, title=None), x=alt.X("Score:Q", scale=alt.Scale(domain=[0,100])),
                color=alt.condition("datum.Focused", alt.value("#111111"), alt.value("#4c78a8")),
                stroke=alt.condition("datum.Focused", alt.value("#111111"), alt.value(None)),
                strokeWidth=alt.condition("datum.Focused", alt.value(3), alt.value(0)),
                tooltip=["Dimension", alt.Tooltip("Score:Q", format=".1f")])
            avg_labels = alt.Chart(avg_profile).mark_text(align="left", dx=5).encode(y=alt.Y("Dimension:N", sort=None), x="Score:Q", text=alt.Text("Score:Q", format=".1f"))
            st.altair_chart((avg_bars+avg_labels).properties(height=240), width="stretch")
            st.markdown("#### Average profile by macro-area")
            st.altair_chart(macro_chart, width="stretch")

        distribution = (
            alt.Chart(filtered)
            .mark_bar()
            .encode(
                x=alt.X("overall_score:Q", bin=alt.Bin(maxbins=18), title="Overall score"),
                y=alt.Y("count():Q", title="Number of universities"),
                tooltip=[alt.Tooltip("count():Q", title="Universities")],
            )
            .properties(height=260)
        )
        st.markdown("#### Distribution of overall scores")
        st.altair_chart(distribution, width="stretch")

    elif page == "University Profile":
        active_context = selected_context(selected, page)
        active_context = add_profile_position_context(active_context, filtered, selected)
        st.markdown("<div id='source-profile-summary'></div>", unsafe_allow_html=True)
        st.markdown("### University profile")
        st.caption("This page focuses on one selected university and compares its overall profile with national and macro-area averages.")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Overall score", f"{selected['overall_score']:.1f}")
        k2.metric("Rank", f"{int(selected['overall_rank_year'])}/61")
        k3.metric("Teaching", f"{selected['teaching_score']:.1f}")
        k4.metric("Research", f"{selected['research_score']:.1f}")
        k5.metric("Financial", f"{selected['financial_score']:.1f}")

        st.markdown("#### Profile classification and percentile badges")
        badge_cols = st.columns(5)
        badge_cols[0].info(f"Overall: {percentile_badge(filtered, 'overall_score', selected['overall_score'])}")
        badge_cols[1].info(f"Teaching: {percentile_badge(filtered, 'teaching_score', selected['teaching_score'])}")
        badge_cols[2].info(f"Placement: {percentile_badge(filtered, 'placement_score', selected['placement_score'])}")
        badge_cols[3].info(f"Research: {percentile_badge(filtered, 'research_score', selected['research_score'])}")
        badge_cols[4].info(f"Financial: {percentile_badge(filtered, 'financial_score', selected['financial_score'])}")

        profile_type = classify_profile(selected)
        strengths, weaknesses = strengths_weaknesses(selected, filtered)
        active_context["profile_type"] = profile_type
        active_context["strengths"] = strengths
        active_context["weaknesses"] = weaknesses
        sw1, sw2, sw3 = st.columns([1, 1, 1])
        with sw1:
            st.markdown("##### Profile type")
            st.success(profile_type)
        with sw2:
            st.markdown("##### Main strengths")
            for item in strengths:
                st.write(f"- {item}")
        with sw3:
            st.markdown("##### Points to inspect")
            for item in weaknesses:
                st.write(f"- {item}")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("<div id='source-profile-dimensions'></div>", unsafe_allow_html=True)
            st.markdown("#### Performance profile")
            profile_focus = page_focus_field("University Profile")
            profile_df = make_score_profile(selected)
            profile_dimension_selection = alt.selection_point(name="profile_dimension_selection", fields=["Field", "Label", "Value"], on="click", clear="dblclick")
            profile_chart = score_bar_chart(profile_df, profile_focus).add_params(profile_dimension_selection)
            st.caption("Click a dimension bar to focus the Analysis Companion on that dimension.")
            st.altair_chart(profile_chart, key="profile_dimension_chart", on_select=on_profile_dimension_select, selection_mode=["profile_dimension_selection"], width="stretch")
        with c2:
            st.markdown("<div id='source-profile-benchmark'></div>", unsafe_allow_html=True)
            st.markdown("#### Benchmark comparison")
            st.altair_chart(benchmark_chart(selected, page_focus_field("University Profile")), width="stretch")

        trend = df[df["university"] == university].sort_values("year")
        trend_long = trend.melt(
            id_vars=["year", "university"],
            value_vars=["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"],
            var_name="Score type",
            value_name="Score",
        )
        trend_chart = (
            alt.Chart(trend_long)
            .mark_line(point=True)
            .encode(
                x=alt.X("year:O", title="Year"),
                y=alt.Y("Score:Q", scale=alt.Scale(domain=[0, 100])),
                color=alt.Color("Score type:N", title="Score type"),
                tooltip=["year", "Score type", alt.Tooltip("Score:Q", format=".1f")],
            )
            .properties(height=330)
        )
        st.markdown("#### Score dynamics over time")
        st.altair_chart(trend_chart, width="stretch")

    elif page == "University Comparison":
        st.markdown("<div id='source-comparison-summary'></div>", unsafe_allow_html=True)
        st.markdown("### University comparison")
        st.caption("Select two universities in the same year and compare their scores, ranks, financial indicators, and time trends. The comparison is exploratory and does not imply causality.")

        universities_available = sorted(filtered["university"].unique())
        default_a = universities_available.index(university) if university in universities_available else 0
        default_b = 1 if len(universities_available) > 1 else 0
        if default_b == default_a and len(universities_available) > 1:
            default_b = 0 if default_a != 0 else 1

        ca, cb = st.columns(2)
        with ca:
            university_a = st.selectbox("University A", universities_available, index=default_a, key="university_a_select")
        with cb:
            university_b = st.selectbox("University B", universities_available, index=default_b, key="university_b_select")

        if university_a == university_b:
            st.warning("Please choose two different universities.")
            active_context = selected_context(selected, page)
            active_context["note"] = "The comparison page is open, but only one university has been selected."
        else:
            row_a = filtered[filtered["university"] == university_a].iloc[0]
            row_b = filtered[filtered["university"] == university_b].iloc[0]
            active_context = comparison_context(row_a, row_b)

            cc1, cc2, cc3, cc4 = st.columns(4)
            cc1.metric(f"{university_a} overall", f"{row_a['overall_score']:.1f}", f"rank {int(row_a['overall_rank_year'])}/61")
            cc2.metric(f"{university_b} overall", f"{row_b['overall_score']:.1f}", f"rank {int(row_b['overall_rank_year'])}/61")
            cc3.metric("Overall gap", f"{row_a['overall_score'] - row_b['overall_score']:.1f}")
            cc4.metric("Comparison year", f"{year}")

            comp_scores = []
            for label, col in [
                ("Overall", "overall_score"),
                ("Teaching", "teaching_score"),
                ("Placement", "placement_score"),
                ("Research", "research_score"),
                ("Financial", "financial_score"),
            ]:
                comp_scores.append({"Dimension": label, "Label": label, "Field": col, "University": university_a, "Score": float(row_a[col]), "Value": float(row_a[col])})
                comp_scores.append({"Dimension": label, "Label": label, "Field": col, "University": university_b, "Score": float(row_b[col]), "Value": float(row_b[col])})
            comp_scores_df = pd.DataFrame(comp_scores)
            comparison_focus = page_focus_field("University Comparison")
            comp_scores_df["Focused"] = comp_scores_df["Field"].eq(comparison_focus)
            comparison_score_selection = alt.selection_point(name="comparison_score_selection", fields=["Field", "Dimension", "University", "Score"], on="click", clear="dblclick")
            comp_chart = (
                alt.Chart(comp_scores_df)
                .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                .encode(
                    x=alt.X("Dimension:N", sort=None, title="Dimension"),
                    y=alt.Y("Score:Q", scale=alt.Scale(domain=[0, 100])),
                    xOffset="University:N",
                    color=alt.Color("University:N", title="University"),
                    stroke=alt.condition("datum.Focused", alt.value("#111111"), alt.value(None)),
                    strokeWidth=alt.condition("datum.Focused", alt.value(3), alt.value(0)),
                    tooltip=["University", "Dimension", "Field", alt.Tooltip("Score:Q", format=".1f")],
                )
                .add_params(comparison_score_selection)
                .properties(height=340)
            )

            comp_labels = alt.Chart(comp_scores_df).mark_text(dy=-8, fontSize=10).encode(
                x=alt.X("Dimension:N", sort=None), y="Score:Q", xOffset="University:N", text=alt.Text("Score:Q", format=".1f")
            )
            comp_chart = comp_chart + comp_labels

            gap_data = []
            for label, col in [
                ("Overall", "overall_score"),
                ("Teaching", "teaching_score"),
                ("Placement", "placement_score"),
                ("Research", "research_score"),
                ("Financial", "financial_score"),
            ]:
                gap = float(row_a[col] - row_b[col])
                gap_data.append(
                    {
                        "Dimension": label,
                        "Label": label,
                        "Field": col,
                        "Gap": gap,
                        "Direction": f"{university_a} higher" if gap >= 0 else f"{university_b} higher",
                    }
                )
            gap_df = pd.DataFrame(gap_data)
            gap_df["Focused"] = gap_df["Field"].eq(comparison_focus)
            comparison_gap_selection = alt.selection_point(name="comparison_gap_selection", fields=["Field", "Dimension", "Gap"], on="click", clear="dblclick")
            gap_chart = (
                alt.Chart(gap_df)
                .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    x=alt.X("Gap:Q", title=f"Score gap: {university_a} minus {university_b}"),
                    y=alt.Y("Dimension:N", sort=None, title="Dimension"),
                    color=alt.Color("Direction:N", title="Direction"),
                    stroke=alt.condition("datum.Focused", alt.value("#111111"), alt.value(None)),
                    strokeWidth=alt.condition("datum.Focused", alt.value(3), alt.value(0)),
                    tooltip=["Dimension", "Field", alt.Tooltip("Gap:Q", format=".1f"), "Direction"],
                )
                .add_params(comparison_gap_selection)
                .properties(height=340)
            )
            gap_labels = alt.Chart(gap_df).mark_text(align="left", dx=5).encode(
                x="Gap:Q", y=alt.Y("Dimension:N", sort=None), text=alt.Text("Gap:Q", format="+.1f")
            )
            gap_chart = gap_chart + gap_labels

            gc1, gc2 = st.columns(2)
            with gc1:
                st.markdown("<div id='source-comparison-scores'></div>", unsafe_allow_html=True)
                st.markdown("#### Side-by-side score profile")
                st.caption("Click a bar to focus the Companion on that university and dimension within the comparison.")
                st.altair_chart(comp_chart, key="comparison_score_chart", on_select=on_comparison_score_select, selection_mode=["comparison_score_selection"], width="stretch")
            with gc2:
                st.markdown("#### Score gaps")
                st.caption("Click a gap bar to focus the Companion on that difference.")
                st.altair_chart(gap_chart, key="comparison_gap_chart", on_select=on_comparison_gap_select, selection_mode=["comparison_gap_selection"], width="stretch")

            trends = df[df["university"].isin([university_a, university_b])].sort_values(["university", "year"])
            trend_comp = (
                alt.Chart(trends)
                .mark_line(point=True)
                .encode(
                    x=alt.X("year:O", title="Year"),
                    y=alt.Y("overall_score:Q", title="Overall score", scale=alt.Scale(domain=[0, 100])),
                    color=alt.Color("university:N", title="University"),
                    tooltip=["university", "year", alt.Tooltip("overall_score:Q", format=".1f")],
                )
                .properties(height=300)
            )
            st.markdown("#### Overall score trend comparison, 2020-2023")
            st.altair_chart(trend_comp, width="stretch")

            comparison_table_cols = [
                "university",
                "year",
                "region",
                "macro_area",
                "size_class",
                "overall_score",
                "overall_rank_year",
                "teaching_score",
                "placement_score",
                "research_score",
                "financial_score",
                "ffo_per_student",
                "operating_cost_per_student",
                "personnel_cost_share",
            ]
            st.markdown("#### Key indicator table")
            st.dataframe(pd.DataFrame([row_a, row_b])[comparison_table_cols], width="stretch", hide_index=True)

    elif page == "Finance Explorer":
        active_context = selected_context(selected, page)
        st.markdown("<div id='source-finance-summary'></div>", unsafe_allow_html=True)
        st.markdown("### Finance explorer")
        st.caption("Use this page to explore observed relationships between financial indicators and profile scores. The scatterplot is exploratory and does not imply causality.")

        f1, f2, f3, f4 = st.columns(4)
        f1.metric("FFO per student", format_number(selected.get("ffo_per_student"), 0))
        f2.metric("Operating cost per student", format_number(selected.get("operating_cost_per_student"), 0))
        f3.metric("Personnel cost share", format_number(selected.get("personnel_cost_share"), 2))
        f4.metric("Public revenue share", format_number(selected.get("public_revenue_share"), 2))

        finance_options = {
            "FFO per student": "ffo_per_student",
            "Operating cost per student": "operating_cost_per_student",
            "Personnel cost share": "personnel_cost_share",
            "Public revenue share": "public_revenue_share",
            "Student contribution share": "student_contribution_share",
            "Performance quota share": "performance_quota_share",
            "Economic-financial sustainability index": "economic_financial_sustainability_index",
        }
        score_options = {
            "Overall score": "overall_score",
            "Teaching score": "teaching_score",
            "Placement score": "placement_score",
            "Research score": "research_score",
            "Financial score": "financial_score",
        }
        fc1, fc2 = st.columns(2)
        with fc1:
            finance_label = st.selectbox("Financial indicator for x-axis", list(finance_options.keys()))
        with fc2:
            score_label_choice = st.selectbox("Score for y-axis", list(score_options.keys()))
        x_col = finance_options[finance_label]
        y_col = score_options[score_label_choice]
        active_context["finance_x_axis"] = finance_label
        active_context["score_y_axis"] = score_label_choice
        active_context["selected_x_value"] = clean_value(selected.get(x_col))
        active_context["selected_y_value"] = clean_value(selected.get(y_col))
        active_context = add_finance_position_context(active_context, filtered, x_col, y_col, selected)

        scatter_data = filtered.copy()
        scatter_data["selected_flag"] = scatter_data["university"].eq(university)
        finance_point_selection = alt.selection_point(
            name="finance_point_selection",
            fields=["university"],
            on="click",
            clear="dblclick",
        )
        base_scatter = (
            alt.Chart(scatter_data)
            .mark_circle(size=85, opacity=0.65)
            .encode(
                x=alt.X(f"{x_col}:Q", title=finance_label),
                y=alt.Y(f"{y_col}:Q", title=score_label_choice, scale=alt.Scale(domain=[0, 100])),
                color=alt.Color("macro_area:N", title="Macro-area"),
                size=alt.Size("enrolled_students:Q", title="Enrolled students"),
                tooltip=[
                    "university",
                    "region",
                    "macro_area",
                    alt.Tooltip(f"{x_col}:Q", title=finance_label, format=",.2f"),
                    alt.Tooltip(f"{y_col}:Q", title=score_label_choice, format=".1f"),
                ],
            )
            .add_params(finance_point_selection)
            .interactive()
        )
        selected_point = (
            alt.Chart(scatter_data[scatter_data["selected_flag"]])
            .mark_circle(size=260, fillOpacity=0, stroke="black", strokeWidth=3)
            .encode(x=alt.X(f"{x_col}:Q"), y=alt.Y(f"{y_col}:Q"), tooltip=["university"])
        )
        scatter_data["selected_label_text"] = scatter_data.apply(
            lambda r: f"{r['university']} | {finance_label}: {float(r[x_col]):,.2f} | {score_label_choice}: {float(r[y_col]):.1f}" if pd.notna(r.get(x_col)) and pd.notna(r.get(y_col)) else str(r["university"]), axis=1
        )
        selected_label = (
            alt.Chart(scatter_data[scatter_data["selected_flag"]])
            .mark_text(align="left", dx=10, dy=-10, fontSize=12, fontWeight="bold")
            .encode(x=alt.X(f"{x_col}:Q"), y=alt.Y(f"{y_col}:Q"), text="selected_label_text:N")
        )
        st.markdown("<div id='source-finance-scatter'></div>", unsafe_allow_html=True)
        st.markdown("#### Financial indicator vs selected score")
        st.caption("Click any university point to make it the active dashboard selection. The sidebar, highlighted point, and Analysis Companion will update together.")
        st.altair_chart(
            (base_scatter + selected_point + selected_label).properties(height=380),
            key="finance_scatter_chart",
            on_select=on_finance_scatter_select,
            selection_mode=["finance_point_selection"],
            width="stretch",
        )
        finance_selection_meta = st.session_state.get("chart_selection_meta") or {}
        if (
            finance_selection_meta.get("source_page") == "Finance Explorer"
            and finance_selection_meta.get("university") == university
        ):
            st.success(f"Selected from the Finance scatterplot: {university}. This selection is now linked to the Analysis Companion.")

        finance_indicators = {
            "personnel_cost_share": "Personnel cost share",
            "public_revenue_share": "Public revenue share",
            "student_contribution_share": "Student contribution share",
            "performance_quota_share": "Performance quota share",
            "economic_financial_sustainability_index": "Economic-financial sustainability index",
        }
        st.markdown("<div id='source-finance-indicators'></div>", unsafe_allow_html=True)
        st.markdown("#### Selected university financial structure")
        finance_indicator_selection = alt.selection_point(name="finance_indicator_selection", fields=["Field", "Indicator", "Value"], on="click", clear="dblclick")
        finance_indicator_chart = make_indicator_bar(selected, finance_indicators, page_focus_field("Finance Explorer")).add_params(finance_indicator_selection)
        st.caption("Click a financial bar to focus the Companion on that indicator.")
        st.altair_chart(finance_indicator_chart, key="finance_indicator_chart", on_select=on_finance_indicator_select, selection_mode=["finance_indicator_selection"], width="stretch")


    elif page == "DEA Efficiency Explorer":
        active_context = make_dea_context(filtered, selected)
        st.markdown("<div id='source-dea-summary'></div>", unsafe_allow_html=True)
        st.markdown("### DEA efficiency explorer")
        st.caption("This page adds an exploratory DEA efficiency layer. DEA-VRS is the main score because universities differ in size and scale.")

        if "dea_vrs_efficiency_100" not in df.columns:
            st.error("DEA columns are not available in the current dataset. Upload university_dashboard_with_dea_efficiency.xlsx and use the Dashboard_Data_with_DEA sheet.")
        else:
            d1, d2, d3, d4, d5 = st.columns(5)
            d1.metric("DEA-VRS efficiency", format_number(selected.get("dea_vrs_efficiency_100"), 1))
            d2.metric("DEA-VRS rank", f"{int(selected.get('dea_vrs_rank_year'))}/61")
            d3.metric("DEA-CRS efficiency", format_number(selected.get("dea_crs_efficiency_100"), 1))
            d4.metric("Scale efficiency", format_number(selected.get("dea_scale_efficiency_100"), 1))
            d5.metric("Category", str(selected.get("efficiency_category")))

            st.markdown("#### Efficiency position")
            ecols = st.columns(3)
            ecols[0].info(f"DEA-VRS: {percentile_badge(filtered, 'dea_vrs_efficiency_100', selected['dea_vrs_efficiency_100'])}")
            ecols[1].info(f"DEA-CRS: {percentile_badge(filtered, 'dea_crs_efficiency_100', selected['dea_crs_efficiency_100'])}")
            ecols[2].info(f"Scale efficiency: {percentile_badge(filtered, 'dea_scale_efficiency_100', selected['dea_scale_efficiency_100'])}")

            dea_x_options = {
                "FFO per student": "ffo_per_student",
                "Operating cost per student": "operating_cost_per_student",
                "Personnel cost share": "personnel_cost_share",
                "Staff per 1000 students": "staff_per_1000_students",
                "Non-academic staff per 1000 students": "non_academic_staff_per_1000_students",
                "Overall score": "overall_score",
            }
            dea_x_label = st.selectbox("X-axis for DEA scatterplot", list(dea_x_options.keys()))
            dea_x_col = dea_x_options[dea_x_label]
            active_context["dea_scatter_x_axis"] = dea_x_label
            active_context["dea_scatter_x_value"] = clean_value(selected.get(dea_x_col))

            dea_data = filtered.copy()
            dea_data["selected_flag"] = dea_data["university"].eq(university)
            dea_point_selection = alt.selection_point(
                name="dea_point_selection",
                fields=["university"],
                on="click",
                clear="dblclick",
            )
            dea_scatter = (
                alt.Chart(dea_data)
                .mark_circle(size=90, opacity=0.65)
                .encode(
                    x=alt.X(f"{dea_x_col}:Q", title=dea_x_label),
                    y=alt.Y("dea_vrs_efficiency_100:Q", title="DEA-VRS efficiency", scale=alt.Scale(domain=[0, 100])),
                    color=alt.Color("efficiency_category:N", title="Efficiency category"),
                    size=alt.Size("enrolled_students:Q", title="Enrolled students"),
                    tooltip=["university", "region", "macro_area", alt.Tooltip(f"{dea_x_col}:Q", title=dea_x_label, format=",.2f"), alt.Tooltip("dea_vrs_efficiency_100:Q", title="DEA-VRS", format=".1f")],
                )
                .add_params(dea_point_selection)
                .interactive()
            )
            selected_dea_point = (
                alt.Chart(dea_data[dea_data["selected_flag"]])
                .mark_circle(size=270, fillOpacity=0, stroke="black", strokeWidth=3)
                .encode(x=alt.X(f"{dea_x_col}:Q"), y=alt.Y("dea_vrs_efficiency_100:Q"), tooltip=["university"])
            )
            dea_data["selected_label_text"] = dea_data.apply(
                lambda r: f"{r['university']} | {dea_x_label}: {float(r[dea_x_col]):,.2f} | DEA-VRS: {float(r['dea_vrs_efficiency_100']):.1f}" if pd.notna(r.get(dea_x_col)) and pd.notna(r.get("dea_vrs_efficiency_100")) else str(r["university"]), axis=1
            )
            selected_dea_label = (
                alt.Chart(dea_data[dea_data["selected_flag"]])
                .mark_text(align="left", dx=10, dy=-10, fontSize=12, fontWeight="bold")
                .encode(x=alt.X(f"{dea_x_col}:Q"), y=alt.Y("dea_vrs_efficiency_100:Q"), text="selected_label_text:N")
            )
            st.markdown("<div id='source-dea-scatter'></div>", unsafe_allow_html=True)
            st.markdown("#### DEA efficiency vs selected variable")
            st.caption("Click any university point to select it directly from the DEA view. The sidebar, highlighted point, DEA context, and Analysis Companion will update together.")
            st.altair_chart(
                (dea_scatter + selected_dea_point + selected_dea_label).properties(height=390),
                key="dea_scatter_chart",
                on_select=on_dea_scatter_select,
                selection_mode=["dea_point_selection"],
                width="stretch",
            )
            dea_selection_meta = st.session_state.get("chart_selection_meta") or {}
            if (
                dea_selection_meta.get("source_page") == "DEA Efficiency Explorer"
                and dea_selection_meta.get("university") == university
            ):
                st.success(f"Selected from the DEA scatterplot: {university}. This selection is now linked to the Analysis Companion.")

            dc1, dc2 = st.columns(2)
            with dc1:
                st.markdown("#### DEA trend, 2020-2023")
                dea_trend = df[df["university"].eq(university)].sort_values("year")
                dea_trend_long = dea_trend.melt(
                    id_vars=["year", "university"],
                    value_vars=["dea_vrs_efficiency_100", "dea_crs_efficiency_100", "dea_scale_efficiency_100"],
                    var_name="DEA measure",
                    value_name="Score",
                )
                dea_trend_chart = (
                    alt.Chart(dea_trend_long)
                    .mark_line(point=True)
                    .encode(
                        x=alt.X("year:O", title="Year"),
                        y=alt.Y("Score:Q", scale=alt.Scale(domain=[0, 100])),
                        color=alt.Color("DEA measure:N", title="DEA measure"),
                        tooltip=["year", "DEA measure", alt.Tooltip("Score:Q", format=".1f")],
                    )
                    .properties(height=320)
                )
                st.altair_chart(dea_trend_chart, width="stretch")
            with dc2:
                st.markdown("#### Top 10 by DEA-VRS efficiency")
                top_dea = filtered.sort_values("dea_vrs_efficiency_100", ascending=False).head(10)
                dea_top_university_selection = alt.selection_point(name="dea_top_university_selection", fields=["university"], on="click", clear="dblclick")
                top_dea_chart = (
                    alt.Chart(top_dea)
                    .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                    .encode(
                        x=alt.X("dea_vrs_efficiency_100:Q", title="DEA-VRS efficiency", scale=alt.Scale(domain=[0, 100])),
                        y=alt.Y("university:N", sort="-x", title=None),
                        color=alt.condition(alt.datum.university == university, alt.value("#111111"), alt.value("#4c78a8")),
                        tooltip=["university", alt.Tooltip("dea_vrs_efficiency_100:Q", format=".1f"), "efficiency_category"],
                    )
                    .add_params(dea_top_university_selection)
                    .properties(height=320)
                )
                top_dea_labels = alt.Chart(top_dea).mark_text(align="left", dx=5).encode(
                    x="dea_vrs_efficiency_100:Q", y=alt.Y("university:N", sort="-x"), text=alt.Text("dea_vrs_efficiency_100:Q", format=".1f")
                )
                top_dea_chart = top_dea_chart + top_dea_labels
                st.caption("Click a university bar to make it the active selection.")
                st.altair_chart(top_dea_chart, key="dea_top_chart", on_select=on_dea_top_select, selection_mode=["dea_top_university_selection"], width="stretch")

            st.markdown("#### DEA inputs and outputs for selected university")
            dea_table_cols = [
                "university", "year", "ffo_per_student", "operating_cost_per_student", "personnel_cost_share",
                "staff_per_1000_students", "non_academic_staff_per_1000_students", "teaching_score", "placement_score",
                "research_score", "dea_vrs_efficiency_100", "dea_crs_efficiency_100", "dea_scale_efficiency_100",
                "efficiency_category"
            ]
            st.dataframe(pd.DataFrame([selected])[dea_table_cols], width="stretch", hide_index=True)
            st.info("DEA efficiency is a benchmarking result based on selected inputs and outputs. It should not be interpreted as a causal estimate or final policy judgment.")

    elif page == "Ranking Explorer":
        st.markdown("<div id='source-ranking-summary'></div>", unsafe_allow_html=True)
        st.markdown("### Ranking explorer")
        st.caption("Rank universities by any dashboard score or DEA efficiency measure within the current filters.")
        metric_options = {
            "Overall score": "overall_score",
            "Teaching score": "teaching_score",
            "Placement score": "placement_score",
            "Research score": "research_score",
            "Financial score": "financial_score",
        }
        if "dea_vrs_efficiency_100" in df.columns:
            metric_options.update({
                "DEA-VRS efficiency": "dea_vrs_efficiency_100",
                "DEA-CRS efficiency": "dea_crs_efficiency_100",
                "Scale efficiency": "dea_scale_efficiency_100",
            })
        rc1, rc2 = st.columns(2)
        with rc1:
            ranking_metric_label = st.selectbox("Ranking metric", list(metric_options.keys()))
        with rc2:
            n_universities = len(filtered)

            if n_universities == 0:
                st.warning("No universities match the current filters.")
                st.stop()
            elif n_universities == 1:
                top_n = 1
                st.info("Only one university matches the current filters, so the ranking contains one institution.")
            else:
                max_top_n = min(25, n_universities)
                default_top_n = min(10, max_top_n)

                top_n = st.slider(
                    "Number of universities",
                    min_value=1,
                    max_value=max_top_n,
                    value=default_top_n,
                    step=1,
                )
        ranking_col = metric_options[ranking_metric_label]
        active_context = make_ranking_context(filtered, ranking_col, ranking_metric_label, selected, top_n)

        ranked = filtered.sort_values(ranking_col, ascending=False).reset_index(drop=True)
        ranked["rank_current_filter"] = ranked.index + 1
        rank_row = ranked[ranked["university"].eq(university)].iloc[0]
        r1, r2, r3 = st.columns(3)
        r1.metric("Selected value", format_number(rank_row[ranking_col], 1))
        r2.metric("Rank in current filter", f"{int(rank_row['rank_current_filter'])}/{len(ranked)}")
        r3.metric("Percentile", f"{percentile_position(filtered, ranking_col, selected[ranking_col]):.0f}th")

        rank_chart_data = ranked.head(top_n).copy()
        rank_chart_data["selected_flag"] = rank_chart_data["university"].eq(university)
        ranking_university_selection = alt.selection_point(name="ranking_university_selection", fields=["university"], on="click", clear="dblclick")
        rank_chart = (
            alt.Chart(rank_chart_data)
            .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
            .encode(
                x=alt.X(f"{ranking_col}:Q", title=ranking_metric_label, scale=alt.Scale(domain=[0, 100])),
                y=alt.Y("university:N", sort="-x", title=None),
                color=alt.condition(alt.datum.selected_flag, alt.value("black"), alt.Color("macro_area:N", title="Macro-area")),
                tooltip=["rank_current_filter", "university", "macro_area", alt.Tooltip(f"{ranking_col}:Q", title=ranking_metric_label, format=".1f")],
            )
            .add_params(ranking_university_selection)
            .properties(height=max(330, 28 * len(rank_chart_data)))
        )
        rank_labels = alt.Chart(rank_chart_data).mark_text(align="left", dx=5).encode(
            x=alt.X(f"{ranking_col}:Q"), y=alt.Y("university:N", sort="-x"), text=alt.Text(f"{ranking_col}:Q", format=".1f")
        )
        rank_chart = rank_chart + rank_labels
        st.markdown("<div id='source-ranking-chart'></div>", unsafe_allow_html=True)
        st.markdown(f"#### Top {top_n} universities by {ranking_metric_label}")
        st.caption("Click a university bar to make it the active dashboard selection.")
        st.altair_chart(rank_chart, key="ranking_bar_chart", on_select=on_ranking_bar_select, selection_mode=["ranking_university_selection"], width="stretch")

        show_cols = ["rank_current_filter", "university", "region", "macro_area", "size_class", ranking_col, "overall_score", "dea_vrs_efficiency_100"]

        # Remove duplicate column names.
        # This is necessary because ranking_col can be "overall_score" or "dea_vrs_efficiency_100",
        # which are also included in the default display columns. PyArrow/Streamlit cannot render
        # a dataframe with duplicate column names.
        show_cols = list(dict.fromkeys(show_cols))

        show_cols = [c for c in show_cols if c in ranked.columns]
        st.markdown("#### Ranking table")
        st.dataframe(ranked[show_cols].copy(), width="stretch", hide_index=True)

    elif page == "Time Dynamics / What Changed":
        st.markdown("<div id='source-time-change'></div>", unsafe_allow_html=True)
        st.markdown("### Time dynamics / What changed?")
        st.caption("Compare how the selected university changed between two years and identify which dimension moved the most.")
        all_years = sorted(df["year"].unique())
        tc1, tc2 = st.columns(2)
        with tc1:
            start_year = st.selectbox("Start year", all_years, index=0)
        with tc2:
            end_year = st.selectbox("End year", all_years, index=len(all_years) - 1)
        if start_year >= end_year:
            st.warning("Choose an end year later than the start year.")
            active_context = {"view_type": "Time Dynamics / What Changed", "university": university, "note": "Invalid year range"}
        else:
            active_context = make_change_context(df, university, start_year, end_year)
            uni_years = df[(df["university"].eq(university)) & (df["year"].isin([start_year, end_year]))].sort_values("year")
            if uni_years["year"].nunique() < 2:
                st.error("The selected university does not have data for both selected years.")
            else:
                start_row = uni_years[uni_years["year"].eq(start_year)].iloc[0]
                end_row = uni_years[uni_years["year"].eq(end_year)].iloc[0]
                change_cols = ["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"]
                if "dea_vrs_efficiency_100" in df.columns:
                    change_cols.append("dea_vrs_efficiency_100")
                change_rows = []
                for col in change_cols:
                    change_rows.append({
                        "Dimension": col.replace("_score", "").replace("dea_vrs_efficiency_100", "DEA-VRS efficiency").replace("_", " ").title(),
                        "Label": col.replace("_score", "").replace("dea_vrs_efficiency_100", "DEA-VRS efficiency").replace("_", " ").title(),
                        "Field": col,
                        "Start": float(start_row[col]),
                        "End": float(end_row[col]),
                        "Change": float(end_row[col] - start_row[col]),
                    })
                change_df = pd.DataFrame(change_rows)
                tcards = st.columns(4)
                tcards[0].metric("Overall change", f"{end_row['overall_score'] - start_row['overall_score']:+.1f}")
                tcards[1].metric("Teaching change", f"{end_row['teaching_score'] - start_row['teaching_score']:+.1f}")
                tcards[2].metric("Research change", f"{end_row['research_score'] - start_row['research_score']:+.1f}")
                if "dea_vrs_efficiency_100" in df.columns:
                    tcards[3].metric("DEA-VRS change", f"{end_row['dea_vrs_efficiency_100'] - start_row['dea_vrs_efficiency_100']:+.1f}")
                else:
                    tcards[3].metric("Financial change", f"{end_row['financial_score'] - start_row['financial_score']:+.1f}")

                time_focus = page_focus_field("Time Dynamics / What Changed")
                change_df["Focused"] = change_df["Field"].eq(time_focus)
                time_change_selection = alt.selection_point(name="time_change_selection", fields=["Field", "Dimension", "Change"], on="click", clear="dblclick")
                change_chart = (
                    alt.Chart(change_df)
                    .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                    .encode(
                        x=alt.X("Change:Q", title=f"Change from {start_year} to {end_year}"),
                        y=alt.Y("Dimension:N", sort=None, title=None),
                        color=alt.condition("datum.Focused", alt.value("#111111"), alt.value("#4c78a8")),
                        stroke=alt.condition("datum.Focused", alt.value("#111111"), alt.value(None)),
                        strokeWidth=alt.condition("datum.Focused", alt.value(3), alt.value(0)),
                        tooltip=["Dimension", "Field", alt.Tooltip("Start:Q", format=".1f"), alt.Tooltip("End:Q", format=".1f"), alt.Tooltip("Change:Q", format="+.1f")],
                    )
                    .add_params(time_change_selection)
                    .properties(height=330)
                )
                change_labels = alt.Chart(change_df).mark_text(align="left", dx=5).encode(
                    x="Change:Q", y=alt.Y("Dimension:N", sort=None), text=alt.Text("Change:Q", format="+.1f")
                )
                change_chart = change_chart + change_labels
                st.markdown("#### Change by dimension")
                st.caption("Click a dimension bar to focus the Companion on that change.")
                st.altair_chart(change_chart, key="time_change_chart", on_select=on_time_change_select, selection_mode=["time_change_selection"], width="stretch")

                trend_df = df[df["university"].eq(university)].sort_values("year")
                trend_cols = ["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"]
                if "dea_vrs_efficiency_100" in df.columns:
                    trend_cols.append("dea_vrs_efficiency_100")
                trend_long = trend_df.melt(id_vars=["year", "university"], value_vars=trend_cols, var_name="Metric", value_name="Value")
                trend_chart = (
                    alt.Chart(trend_long)
                    .mark_line(point=True)
                    .encode(
                        x=alt.X("year:O", title="Year"),
                        y=alt.Y("Value:Q", scale=alt.Scale(domain=[0, 100])),
                        color=alt.Color("Metric:N", title="Metric"),
                        tooltip=["year", "Metric", alt.Tooltip("Value:Q", format=".1f")],
                    )
                    .properties(height=330)
                )
                st.markdown("<div id='source-time-trend'></div>", unsafe_allow_html=True)
                st.markdown("#### Full trend, 2020-2023")
                st.altair_chart(trend_chart, width="stretch")

                st.markdown("#### Universities with largest overall improvement / decline")
                start_group = df[df["year"].eq(start_year)][["university", "overall_score"]].rename(columns={"overall_score": "overall_start"})
                end_group = df[df["year"].eq(end_year)][["university", "overall_score", "region", "macro_area"]].rename(columns={"overall_score": "overall_end"})
                changes_all = start_group.merge(end_group, on="university", how="inner")
                changes_all["overall_change"] = changes_all["overall_end"] - changes_all["overall_start"]
                ca, cb = st.columns(2)
                with ca:
                    st.dataframe(changes_all.sort_values("overall_change", ascending=False).head(10), width="stretch", hide_index=True)
                with cb:
                    st.dataframe(changes_all.sort_values("overall_change", ascending=True).head(10), width="stretch", hide_index=True)

    elif page == "Teaching and Research":
        active_context = selected_context(selected, page)
        active_context = add_teaching_research_position_context(active_context, filtered, selected)
        st.markdown("<div id='source-teaching-summary'></div>", unsafe_allow_html=True)
        st.markdown("### Teaching and research profile")
        st.caption("This page separates teaching, placement, and research indicators so that the profile is not reduced to a single ranking.")
        teaching_indicators = {
            "second_year_retention_pct": "Second-year retention (%)",
            "inactive_students_reversed_score": "Inactive students reversed score",
            "graduation_within_standard_pct": "Graduation within standard duration (%)",
            "graduation_intensity": "Graduation intensity",
            "employment_index": "Employment index",
        }
        research_indicators = {
            "publications_per_teaching_staff": "Publications per teaching staff",
            "citations_per_publication": "Citations per publication",
            "h_index": "H-index",
            "highly_cited_researchers": "Highly cited researchers",
            "nature_science_articles": "Nature and Science articles",
        }
        tr_focus = page_focus_field("Teaching and Research")
        tr_score_df = make_score_profile(selected, include_overall=False)
        tr_score_df = tr_score_df[tr_score_df["Field"].isin(["teaching_score", "placement_score", "research_score"])].copy()
        teaching_research_score_selection = alt.selection_point(name="teaching_research_score_selection", fields=["Field", "Label", "Value"], on="click", clear="dblclick")
        st.markdown("#### Teaching, placement and research scores")
        st.caption("Click a score bar to focus the Companion on that dimension.")
        st.altair_chart(score_bar_chart(tr_score_df, tr_focus).add_params(teaching_research_score_selection), key="teaching_research_score_chart", on_select=on_teaching_research_score_select, selection_mode=["teaching_research_score_selection"], width="stretch")

        tr1, tr2 = st.columns(2)
        with tr1:
            st.markdown("<div id='source-teaching-indicators'></div>", unsafe_allow_html=True)
            st.markdown("#### Teaching and placement indicators")
            teaching_indicator_selection = alt.selection_point(name="teaching_indicator_selection", fields=["Field", "Indicator", "Value"], on="click", clear="dblclick")
            st.caption("Click an indicator bar to focus the Companion on it.")
            st.altair_chart(make_indicator_bar(selected, teaching_indicators, tr_focus).add_params(teaching_indicator_selection), key="teaching_indicator_chart", on_select=on_teaching_indicator_select, selection_mode=["teaching_indicator_selection"], width="stretch")
        with tr2:
            st.markdown("<div id='source-research-indicators'></div>", unsafe_allow_html=True)
            st.markdown("#### Research indicators")
            research_indicator_selection = alt.selection_point(name="research_indicator_selection", fields=["Field", "Indicator", "Value"], on="click", clear="dblclick")
            st.caption("Click an indicator bar to focus the Companion on it.")
            st.altair_chart(make_indicator_bar(selected, research_indicators, tr_focus).add_params(research_indicator_selection), key="research_indicator_chart", on_select=on_research_indicator_select, selection_mode=["research_indicator_selection"], width="stretch")

        teaching_research_university_selection = alt.selection_point(name="teaching_research_university_selection", fields=["university"], on="click", clear="dblclick")
        teaching_research_scatter = (
            alt.Chart(filtered)
            .mark_circle(size=100, opacity=0.7)
            .encode(
                x=alt.X("teaching_score:Q", title="Teaching score", scale=alt.Scale(domain=[0, 100])),
                y=alt.Y("research_score:Q", title="Research score", scale=alt.Scale(domain=[0, 100])),
                color=alt.Color("macro_area:N", title="Macro-area"),
                size=alt.Size("enrolled_students:Q", title="Enrolled students"),
                tooltip=["university", "region", alt.Tooltip("teaching_score:Q", format=".1f"), alt.Tooltip("research_score:Q", format=".1f")],
            )
            .add_params(teaching_research_university_selection)
            .interactive()
        )
        selected_tr = (
            alt.Chart(filtered[filtered["university"] == university])
            .mark_circle(size=260, fillOpacity=0, stroke="black", strokeWidth=3)
            .encode(x="teaching_score:Q", y="research_score:Q")
        )
        st.markdown("<div id='source-teaching-research'></div>", unsafe_allow_html=True)
        st.markdown("#### Teaching vs research positioning")
        st.caption("Click any university point to make it the active dashboard selection.")
        st.altair_chart((teaching_research_scatter + selected_tr).properties(height=360), key="teaching_research_scatter_chart", on_select=on_teaching_research_scatter_select, selection_mode=["teaching_research_university_selection"], width="stretch")

    elif page == "Data and Methodology":
        active_context = {
            "view_type": "Data and Methodology",
            "year": int(year),
            "dataset_scope": "61 Italian universities, 2020-2023",
            "score_definition": "Dashboard-based normalized profile scores plus an exploratory DEA efficiency layer.",
            "current_filter_count": int(filtered["university"].nunique()),
            "macro_area_filter": macro_area,
            "region_filter": region,
            "size_class_filter": size_class,
        }
        st.markdown("<div id='source-methodology'></div>", unsafe_allow_html=True)
        st.markdown("### Data and methodology")
        st.markdown(
            "The dashboard uses a curated prototype dataset covering 61 Italian universities over 2020-2023. "
            "The score fields are dashboard-based normalized profile scores, while the DEA page adds an exploratory efficiency layer. "
            "The DEA score is a descriptive benchmarking measure based on selected inputs and outputs, not a causal estimate."
        )
        st.markdown("#### Current page AI input")
        st.json(active_context)
        st.markdown("#### Current filtered dataset")
        csv_data = filtered.sort_values("overall_score", ascending=False).to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download current filtered dataset as CSV",
            data=csv_data,
            file_name=f"filtered_university_dashboard_{year}.csv",
            mime="text/csv",
        )
        st.dataframe(filtered.sort_values("overall_score", ascending=False), width="stretch", hide_index=True)

# Attach interaction provenance when the current university was chosen directly from a chart.
chart_selection_meta = st.session_state.get("chart_selection_meta") or {}
if (
    chart_selection_meta.get("university") == university
    and chart_selection_meta.get("source_page") == page
):
    active_context["interaction_selection"] = dict(chart_selection_meta)

chart_focus_meta = st.session_state.get("chart_focus_meta") or {}
if chart_focus_meta.get("source_page") == page:
    active_context["interaction_focus"] = dict(chart_focus_meta)

# Build the evidence registry after all page-specific controls have defined the active context.
# This makes the dashboard -> Companion link explicit and keeps the evidence synchronized
# with year, university, filters, selected axes, ranking metric, and comparison choices.
active_context["evidence_items"] = build_evidence_items(active_context)
_focus_field = current_focus_field()
if _focus_field:
    _focused_item = next((item for item in active_context["evidence_items"] if item.get("field") == _focus_field), None)
    if _focused_item:
        active_context["evidence_focus"] = {k: _focused_item.get(k) for k in ["id", "label", "value", "display_value", "source", "field", "anchor"]}

# A deep-link click causes a Streamlit rerun. Once the selected page and visual are
# rebuilt, immediately return the browser to the corresponding evidence section.
if active_context.get("evidence_focus"):
    scroll_to_evidence_anchor(active_context["evidence_focus"].get("anchor"))

with ai_col:
    st.subheader("AI Analysis Companion")
    st.caption("The assistant analyzes only the currently selected dashboard page.")
    st.markdown(
        "<div class='context-link-note'><b>Live dashboard context is linked.</b> "
        "Changing the university, year, filters, axes, ranking metric, or comparison automatically changes the data sent to the Companion and clears the previous interpretation.</div>",
        unsafe_allow_html=True,
    )

    question = st.text_area(
        "Optional question",
        placeholder="Example: What is the main gap on this page?",
        height=110,
    )

    if "ai_answer" not in st.session_state:
        st.session_state.ai_answer = ""
    if "last_ai_signature" not in st.session_state:
        st.session_state.last_ai_signature = ""
    if "ai_auto_cache" not in st.session_state:
        st.session_state.ai_auto_cache = {}

    # Evidence deep-link focus is navigation/highlighting, not a new analytical state.
    # Excluding it prevents the Companion from rewriting the same interpretation just
    # because the user clicked a number to locate it in the visualization.
    analysis_context = dict(active_context)
    analysis_context.pop("evidence_focus", None)
    signature = json.dumps(analysis_context, ensure_ascii=True, sort_keys=True, default=str)
    context_changed = signature != st.session_state.last_ai_signature
    if context_changed:
        # The visible explanation must always correspond to the current analytical state.
        st.session_state.ai_answer = ""
        st.session_state.last_ai_signature = signature

    view_type = active_context.get("view_type", page)
    if view_type == "University Comparison" and "university_a" in active_context and "university_b" in active_context:
        st.markdown(
            f"**Current page:** {view_type}<br>**Context:** {active_context['university_a']['university']} vs {active_context['university_b']['university']}, {active_context.get('year')}",
            unsafe_allow_html=True,
        )
    elif view_type == "Overview":
        st.markdown(
            f"**Current page:** {view_type}<br>**Context:** {active_context.get('number_of_universities')} universities, {active_context.get('year')}",
            unsafe_allow_html=True,
        )
    elif view_type == "Ranking Explorer":
        st.markdown(f"**Current page:** {view_type}<br>**Context:** ranking by {active_context.get('metric')}, {active_context.get('year')}", unsafe_allow_html=True)
    elif view_type == "Time Dynamics / What Changed":
        st.markdown(f"**Current page:** {view_type}<br>**Context:** {active_context.get('university')}, {active_context.get('start_year')} to {active_context.get('end_year')}", unsafe_allow_html=True)
    elif view_type == "DEA Efficiency Explorer":
        st.markdown(f"**Current page:** {view_type}<br>**Context:** {active_context.get('university')}, {active_context.get('year')}", unsafe_allow_html=True)
    elif view_type == "Data and Methodology":
        st.markdown(f"**Current page:** {view_type}<br>**Context:** dataset, scores, and DEA methodology", unsafe_allow_html=True)
    else:
        st.markdown(
            f"**Current page:** {view_type}<br>**Context:** {active_context.get('university')}, {active_context.get('year')}",
            unsafe_allow_html=True,
        )

    interaction_selection = active_context.get("interaction_selection")
    if interaction_selection:
        st.caption(
            f"Interactive selection source: {interaction_selection.get('source_page')} → "
            f"{interaction_selection.get('university')}. The Companion is analyzing the chart-selected observation."
        )
    interaction_focus = active_context.get("interaction_focus")
    if interaction_focus:
        st.caption(
            f"Interactive focus: {interaction_focus.get('label') or interaction_focus.get('field')} "
            f"({interaction_focus.get('value')}). The Companion prioritizes the clicked bar/indicator."
        )
    if active_context.get("evidence_focus"):
        ef = active_context["evidence_focus"]
        st.caption(f"Evidence focus: {ef.get('label')} = {ef.get('display_value')}. The linked visual mark is highlighted in black.")

    with st.expander("Linked data from current view", expanded=False):
        evidence_df = pd.DataFrame(active_context.get("evidence_items", []))
        if not evidence_df.empty:
            visible_cols = [c for c in ["id", "label", "display_value", "source", "field"] if c in evidence_df.columns]
            st.dataframe(evidence_df[visible_cols], width="stretch", hide_index=True)
        else:
            st.caption("No evidence items are available for this view.")

    auto_analysis = st.toggle(
        "Automatic analysis",
        value=True,
        help=(
            "When enabled, the Companion automatically regenerates its interpretation whenever the "
            "dashboard analytical state changes, including chart-point selections, filters, axes, year, "
            "ranking metric, or comparison choices."
        ),
        key="auto_analysis_enabled",
    )

    # Automatic dashboard -> Companion execution. The context signature changes only when the
    # analytical state changes; typing in the optional question does not trigger an API call.
    # Previously generated automatic answers are cached per analytical state to avoid repeated calls
    # when a user navigates back to a context already analyzed during the same session.
    if auto_analysis and (context_changed or not st.session_state.ai_answer):
        cached_answer = st.session_state.ai_auto_cache.get(signature)
        if cached_answer:
            st.session_state.ai_answer = cached_answer
        else:
            with st.spinner("Updating interpretation for the current dashboard selection..."):
                auto_answer = cached_auto_interpretation(signature)
            st.session_state.ai_answer = auto_answer
            st.session_state.ai_auto_cache[signature] = auto_answer
            # Keep the in-session cache bounded.
            if len(st.session_state.ai_auto_cache) > 24:
                oldest_key = next(iter(st.session_state.ai_auto_cache))
                st.session_state.ai_auto_cache.pop(oldest_key, None)

    # A manual follow-up remains available only when the user wants to ask a specific question.
    # It is no longer required for the normal page interpretation.
    if question.strip():
        if st.button("Ask follow-up", type="secondary"):
            with st.spinner("Answering the follow-up in the current dashboard context..."):
                st.session_state.ai_answer = generate_ai_interpretation(active_context, question.strip())

    if st.session_state.ai_answer:
        linked_answer = linkify_evidence_values(st.session_state.ai_answer, active_context.get("evidence_items", []))
        linked_answer = linkify_evidence_citations(linked_answer, active_context.get("evidence_items", []))
        st.markdown(linked_answer, unsafe_allow_html=True)
        render_evidence_registry(active_context.get("evidence_items", []), st.session_state.ai_answer)

    with st.expander("Current AI input"):
        st.json(active_context)

    st.info(
        "The AI assistant supports interpretation only. It analyzes the current page context and does not provide causal conclusions or autonomous policy recommendations."
    )
