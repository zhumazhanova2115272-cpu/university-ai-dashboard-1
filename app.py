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
COMPANION_LOGIC_VERSION = "v21.0-clean-companion"

SCORE_COMPONENTS = {
    "Teaching": [
        ("second_year_retention_pct", "Second-year retention", "direct"),
        ("inactive_students_reversed_score", "Inactive students reversed score", "direct"),
        ("graduation_within_standard_pct", "Graduation within standard duration", "direct"),
        ("graduation_intensity", "Graduation intensity", "normalize"),
    ],
    "Placement": [
        ("employment_index", "Employment index", "direct"),
        ("internship_participation_pct", "Internship participation", "direct"),
        ("job_placement_services_pct", "Job placement services use", "direct"),
    ],
    "Research": [
        ("publications_per_teaching_staff", "Publications per teaching staff", "normalize"),
        ("citations_per_publication", "Citations per publication", "normalize"),
        ("h_index", "H-index", "normalize"),
        ("highly_cited_researchers", "Highly cited researchers", "normalize"),
        ("nature_science_articles", "Nature / Science articles", "normalize"),
    ],
    "Financial": [
        ("ffo_per_student", "FFO per student", "normalize"),
        ("personnel_cost_share", "Personnel cost share", "reverse_normalize"),
        ("performance_quota_share", "Performance quota share", "normalize"),
        ("economic_financial_sustainability_index", "Economic-financial sustainability index", "normalize"),
    ],
}

SENSITIVITY_SCENARIOS = {
    "Equal weights": {"Teaching": 0.25, "Placement": 0.25, "Research": 0.25, "Financial": 0.25},
    "Teaching emphasis": {"Teaching": 0.45, "Placement": 0.20, "Research": 0.20, "Financial": 0.15},
    "Student outcomes emphasis": {"Teaching": 0.35, "Placement": 0.35, "Research": 0.15, "Financial": 0.15},
    "Research emphasis": {"Teaching": 0.15, "Placement": 0.15, "Research": 0.50, "Financial": 0.20},
    "Financial emphasis": {"Teaching": 0.20, "Placement": 0.15, "Research": 0.25, "Financial": 0.40},
}

SCORE_HELP = {
    "overall_score": "Exploratory average of Teaching, Placement, Research and Financial dashboard scores. It is not an official university ranking.",
    "teaching_score": "Average of second-year retention, inactive-students reversed score, graduation within standard duration and normalized graduation intensity.",
    "placement_score": "Average of employment index, internship participation and job-placement-service use; missing components are ignored.",
    "research_score": "Average of year-normalized research output and impact components.",
    "financial_score": "Average of year-normalized FFO per student, reversed personnel cost share, performance quota share and economic-financial sustainability index.",
    "dea_vrs_efficiency_100": "Relative VRS DEA efficiency under the selected yearly sample and model specification. It is a benchmarking measure, not an absolute quality score.",
    "dea_crs_efficiency_100": "Relative CRS DEA efficiency under the selected yearly sample and model specification.",
    "dea_scale_efficiency_100": "Scale efficiency derived from the relationship between CRS and VRS DEA scores.",
}

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
    .companion-data-link {
        color: inherit;
        text-decoration: underline;
        text-decoration-thickness: 2px;
        text-underline-offset: 2px;
        cursor: pointer;
        font-weight: 700;
    }
    .companion-data-link:hover {
        background: rgba(255, 232, 128, 0.45);
        border-radius: 4px;
    }
    [id^="source-"] {
        scroll-margin-top: 90px;
    }
    .linked-scroll-pulse {
        outline: 2px solid rgba(255, 75, 75, 0.55);
        outline-offset: 8px;
        border-radius: 8px;
        transition: outline-color 0.8s ease;
    }
    .insight-badge {
        display: inline-block;
        border: 1px solid rgba(49, 51, 63, 0.16);
        border-radius: 999px;
        padding: 0.26rem 0.58rem;
        margin: 0.16rem 0.18rem 0.16rem 0;
        font-size: 0.82rem;
        background: rgba(245, 247, 250, 0.92);
    }
    .breadcrumb {
        font-size: 0.9rem;
        color: #666b75;
        margin: 0.15rem 0 0.8rem 0;
    }
    .method-note {
        border-left: 3px solid rgba(49, 51, 63, 0.22);
        padding-left: 0.65rem;
        color: #646a74;
        font-size: 0.88rem;
        margin: 0.35rem 0 0.65rem 0;
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
    """Focus the Companion and visualization on a clicked dimension/indicator bar.

    Clearing the Altair selection also clears the analytical focus, so the
    Companion returns to page-level analysis only when the user explicitly
    removes the mark selection.
    """
    field = extract_chart_selection_field(chart_key, selection_name, field_name)
    if not field:
        existing = st.session_state.get("chart_focus_meta") or {}
        if existing.get("source_page") == source_page:
            st.session_state["chart_focus_meta"] = None
            st.session_state["ai_answer"] = ""
            st.session_state["last_ai_signature"] = ""
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


def normalize_value(series: pd.Series, value: Any, reverse: bool = False) -> float | None:
    numeric = pd.to_numeric(series, errors="coerce").dropna()
    if numeric.empty or value is None or pd.isna(value):
        return None
    lo, hi = float(numeric.min()), float(numeric.max())
    if hi <= lo:
        normalized = 50.0
    else:
        normalized = (float(value) - lo) / (hi - lo) * 100.0
    return 100.0 - normalized if reverse else normalized


def score_component_breakdown(full_data: pd.DataFrame, row: pd.Series, dimension: str) -> pd.DataFrame:
    """Reconstruct the exact dashboard score components documented in Score_Methodology."""
    year_data = full_data[full_data["year"].eq(int(row["year"]))]
    records: list[dict] = []
    for field, label, method in SCORE_COMPONENTS.get(dimension, []):
        raw = row.get(field)
        if raw is None or pd.isna(raw):
            component_score = None
        elif method == "direct":
            component_score = float(raw)
        elif method == "reverse_normalize":
            component_score = normalize_value(year_data[field], raw, reverse=True)
        else:
            component_score = normalize_value(year_data[field], raw, reverse=False)
        records.append({
            "Component": label,
            "Field": field,
            "Raw value": clean_value(raw),
            "Component score": clean_value(component_score),
            "Method": method.replace("_", " "),
        })
    return pd.DataFrame(records)


def profile_dispersion(row: pd.Series) -> float:
    values = [float(row[c]) for c in ["teaching_score", "placement_score", "research_score", "financial_score"] if c in row.index and pd.notna(row[c])]
    return float(pd.Series(values).std(ddof=0)) if values else 0.0


def profile_balance_label(dispersion: float) -> str:
    if dispersion <= 5:
        return "Highly balanced"
    if dispersion <= 10:
        return "Moderately balanced"
    if dispersion <= 15:
        return "Uneven"
    return "Highly specialized / uneven"


def benchmark_group(full_data: pd.DataFrame, row: pd.Series, mode: str, custom_universities: list[str] | None = None, current_filtered: pd.DataFrame | None = None) -> tuple[pd.DataFrame, str]:
    year_data = full_data[full_data["year"].eq(int(row["year"]))].copy()
    if mode == "National":
        return year_data, "National average"
    if mode == "Macro-area":
        subset = year_data[year_data["macro_area"].eq(row.get("macro_area"))]
        return subset, f"{row.get('macro_area')} average"
    if mode == "Region":
        subset = year_data[year_data["region"].eq(row.get("region"))]
        return subset, f"{row.get('region')} average"
    if mode == "Same size":
        subset = year_data[year_data["size_class"].eq(row.get("size_class"))]
        return subset, f"{row.get('size_class')} average"
    if mode == "Current filtered group" and current_filtered is not None and not current_filtered.empty:
        return current_filtered.copy(), "Current filtered-group average"
    if mode == "Custom peer group" and custom_universities:
        subset = year_data[year_data["university"].isin(custom_universities)]
        if not subset.empty:
            return subset, f"Custom peers (n={subset['university'].nunique()})"
    return year_data, "National average"


def benchmark_profile_data(row: pd.Series, group: pd.DataFrame, label: str) -> pd.DataFrame:
    records = []
    for dim, field in [("Overall", "overall_score"), ("Teaching", "teaching_score"), ("Placement", "placement_score"), ("Research", "research_score"), ("Financial", "financial_score")]:
        records.append({"Dimension": dim, "Series": str(row.get("university")), "Score": float(row[field])})
        records.append({"Dimension": dim, "Series": label, "Score": float(pd.to_numeric(group[field], errors="coerce").mean())})
    return pd.DataFrame(records)


def sensitivity_table(year_data: pd.DataFrame, selected_university: str) -> pd.DataFrame:
    dimension_fields = {"Teaching": "teaching_score", "Placement": "placement_score", "Research": "research_score", "Financial": "financial_score"}
    rows = []
    work = year_data.copy()
    for scenario, weights in SENSITIVITY_SCENARIOS.items():
        score = sum(pd.to_numeric(work[dimension_fields[d]], errors="coerce") * w for d, w in weights.items())
        rank = score.rank(method="min", ascending=False)
        idx = work.index[work["university"].eq(selected_university)]
        if len(idx) == 0:
            continue
        i = idx[0]
        rows.append({
            "Scenario": scenario,
            "Score": float(score.loc[i]),
            "Rank": int(rank.loc[i]),
            "Universities": int(work["university"].nunique()),
            "Weights": ", ".join(f"{d[0]}={w:.0%}" for d, w in weights.items()),
        })
    return pd.DataFrame(rows)


def insight_flags(row: pd.Series, comparison_data: pd.DataFrame, full_data: pd.DataFrame) -> list[str]:
    flags: list[str] = []
    for label, field in [("Teaching", "teaching_score"), ("Placement", "placement_score"), ("Research", "research_score"), ("Financial", "financial_score")]:
        pct = percentile_position(comparison_data, field, row.get(field))
        if pct is not None and pct >= 90:
            flags.append(f"Top 10% in {label}")
        elif pct is not None and pct <= 25:
            flags.append(f"Lower quartile in {label}")
    if "dea_vrs_efficiency_100" in row.index and pd.notna(row.get("dea_vrs_efficiency_100")) and float(row.get("dea_vrs_efficiency_100")) >= 99.95:
        flags.append("DEA-VRS frontier")
    dispersion = profile_dispersion(row)
    if dispersion >= 12:
        flags.append("High profile imbalance")
    elif dispersion <= 5:
        flags.append("Balanced profile")
    uni_history = full_data[full_data["university"].eq(row.get("university"))].sort_values("year")
    if len(uni_history) >= 2:
        first, last = uni_history.iloc[0], uni_history.iloc[-1]
        teaching_change = float(last.get("teaching_score") - first.get("teaching_score"))
        overall_change = float(last.get("overall_score") - first.get("overall_score"))
        if teaching_change >= 5:
            flags.append("Strong teaching improvement")
        if overall_change >= 5:
            flags.append("Strong overall improvement")
        elif overall_change <= -5:
            flags.append("Marked overall decline")
    return flags[:7]


def render_insight_flags(flags: list[str]) -> None:
    if not flags:
        st.caption("No rule-based flags are triggered in the current context.")
        return
    html_badges = "".join(f"<span class='insight-badge'>{html.escape(flag)}</span>" for flag in flags)
    st.markdown(html_badges, unsafe_allow_html=True)


def request_visual_explanation(source_page: str, label: str, field: str, key: str) -> None:
    if st.button("✦ Explain this visual", key=key, help="Ask the Analysis Companion to explain the visible pattern in this chart using the current dashboard values."):
        _clear_deep_focus()
        st.session_state["chart_focus_meta"] = {
            "source_page": source_page,
            "interaction": "explain visual request",
            "focus_kind": "visual",
            "field": field,
            "label": label,
            "value": None,
        }
        st.session_state["ai_answer"] = ""
        st.session_state["last_ai_signature"] = ""
        st.rerun()


def reset_analytical_state() -> None:
    keys_to_clear = [
        "year_selector", "macro_area_selector", "region_selector", "size_class_selector", "university_selector", "dashboard_page_selector",
        "chart_selection_meta", "chart_focus_meta", "ai_answer", "last_ai_signature", "university_a_select", "university_b_select",
        "profile_benchmark_mode", "profile_custom_peers", "score_breakdown_dimension", "comparison_benchmark_mode", "comparison_custom_peers",
        "show_largest_dimension_changes", "companion_follow_up_question",
        "finance_scatter_chart", "dea_scatter_chart", "overview_top_chart", "ranking_bar_chart", "dea_top_chart",
        "teaching_research_scatter_chart", "profile_dimension_chart", "finance_indicator_chart", "teaching_indicator_chart",
        "research_indicator_chart", "teaching_research_score_chart", "comparison_score_chart", "comparison_gap_chart", "time_change_chart",
        "overview_linked_brush_chart",
    ]
    for key in keys_to_clear:
        st.session_state.pop(key, None)
    try:
        st.query_params.clear()
    except Exception:
        pass


def interval_selection_payload(event: Any, selection_name: str) -> dict:
    selection_state = _mapping_get(event, "selection", {})
    payload = _mapping_get(selection_state, selection_name, {})
    return dict(payload) if payload else {}


def filter_by_interval(data: pd.DataFrame, payload: dict, x_col: str, y_col: str) -> pd.DataFrame:
    if not payload:
        return data.iloc[0:0].copy()
    x_range = payload.get(x_col) or payload.get("x")
    y_range = payload.get(y_col) or payload.get("y")
    if not isinstance(x_range, (list, tuple)) or not isinstance(y_range, (list, tuple)) or len(x_range) != 2 or len(y_range) != 2:
        return data.iloc[0:0].copy()
    return data[data[x_col].between(min(x_range), max(x_range)) & data[y_col].between(min(y_range), max(y_range))].copy()


def build_context_breadcrumb(page: str, year: int, macro_area: str, region: str, size_class: str, university: str) -> str:
    parts = [str(year)]
    for value in [macro_area, region, size_class]:
        if value and value != "All":
            parts.append(str(value))
    if page != "Overview":
        parts.append(str(university))
    parts.append(page)
    focus = st.session_state.get("chart_focus_meta") or {}
    if focus.get("source_page") == page and focus.get("label"):
        parts.append(str(focus.get("label")))
    return " › ".join(parts)


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
    top = filtered.sort_values("overall_score", ascending=False).head(10)
    bottom = filtered.sort_values("overall_score", ascending=True).head(3)
    macro_profiles = (
        filtered.groupby("macro_area", as_index=False)[["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"]]
        .mean()
        .sort_values("overall_score", ascending=False)
    )
    corr = None
    try:
        corr = clean_value(filtered[["teaching_score", "research_score"]].corr().iloc[0, 1])
    except Exception:
        pass
    overall_series = numeric_series(filtered, "overall_score")
    return {
        "view_type": "Overview",
        "year": int(year),
        "macro_area_filter": macro_area,
        "region_filter": region,
        "size_class_filter": size_class,
        "number_of_universities": int(filtered["university"].nunique()),
        "average_overall_score": clean_value(filtered["overall_score"].mean()),
        "median_overall_score": clean_value(overall_series.median()) if not overall_series.empty else None,
        "std_overall_score": clean_value(overall_series.std()) if len(overall_series) > 1 else None,
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
        "top_universities": top.head(3)[["university", "overall_score", "overall_rank_year"]].to_dict("records"),
        "top_10_universities": top[["university", "overall_score", "overall_rank_year"]].to_dict("records"),
        "lowest_universities": bottom[["university", "overall_score", "overall_rank_year"]].to_dict("records"),
        "macro_area_profiles": macro_profiles.to_dict("records"),
        "teaching_research_correlation": corr,
        "teaching_min": clean_value(filtered["teaching_score"].min()),
        "teaching_max": clean_value(filtered["teaching_score"].max()),
        "research_min": clean_value(filtered["research_score"].min()),
        "research_max": clean_value(filtered["research_score"].max()),
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
        brush = context.get("linked_brush", {})
        if brush:
            add("Universities in brushed group", brush.get("selected_universities"), "Linked brushing explorer", "linked_brush.selected_universities", "source-overview-brush", 0)
            add("Brushed-group average overall score", brush.get("average_overall_score"), "Linked brushing explorer", "linked_brush.average_overall_score", "source-overview-brush", 1)

    elif view == "University Comparison" and context.get("university_a") and context.get("university_b"):
        a = context["university_a"]
        b = context["university_b"]
        gaps = context.get("score_gaps_a_minus_b", {})
        # Keep the individual A/B dimension scores in the evidence registry, not
        # only the gaps. This lets a clicked bar such as "Bari — Placement" be
        # cited and linked directly by the Companion.
        for dim_label, col in [
            ("Overall", "overall_score"),
            ("Teaching", "teaching_score"),
            ("Placement", "placement_score"),
            ("Research", "research_score"),
            ("Financial", "financial_score"),
        ]:
            add(f"{a.get('university')} {dim_label.lower()} score", a.get(col), "Side-by-side score profile", f"university_a.{col}", "source-comparison-scores", 1)
            add(f"{b.get('university')} {dim_label.lower()} score", b.get(col), "Side-by-side score profile", f"university_b.{col}", "source-comparison-scores", 1)
        add("Overall score gap (A-B)", gaps.get("overall_gap_a_minus_b"), "Score gaps", "score_gaps_a_minus_b.overall_gap_a_minus_b", "source-comparison-gaps", 1)
        add("Teaching score gap (A-B)", gaps.get("teaching_gap_a_minus_b"), "Score gaps", "score_gaps_a_minus_b.teaching_gap_a_minus_b", "source-comparison-gaps", 1)
        add("Placement score gap (A-B)", gaps.get("placement_gap_a_minus_b"), "Score gaps", "score_gaps_a_minus_b.placement_gap_a_minus_b", "source-comparison-gaps", 1)
        add("Research score gap (A-B)", gaps.get("research_gap_a_minus_b"), "Score gaps", "score_gaps_a_minus_b.research_gap_a_minus_b", "source-comparison-gaps", 1)
        add("Financial score gap (A-B)", gaps.get("financial_gap_a_minus_b"), "Score gaps", "score_gaps_a_minus_b.financial_gap_a_minus_b", "source-comparison-gaps", 1)

    elif view == "Finance Explorer":
        add(context.get("finance_x_axis", "Selected financial indicator"), context.get("selected_x_value"), "Selected scatterplot point", "selected_x_value", "source-finance-scatter", 2)
        add(context.get("score_y_axis", "Selected score"), context.get("selected_y_value"), "Selected scatterplot point", "selected_y_value", "source-finance-scatter", 1)
        add("X-axis median in current filter", context.get("x_axis_median_current_filter"), "Finance scatterplot benchmark", "x_axis_median_current_filter", "source-finance-scatter", 2)
        add("Y-axis median in current filter", context.get("y_axis_median_current_filter"), "Finance scatterplot benchmark", "y_axis_median_current_filter", "source-finance-scatter", 1)
        add("X-axis percentile in current filter", context.get("x_axis_percentile_current_filter"), "Finance scatterplot position", "x_axis_percentile_current_filter", "source-finance-scatter", 0, "th percentile")
        add("Y-axis percentile in current filter", context.get("y_axis_percentile_current_filter"), "Finance scatterplot position", "y_axis_percentile_current_filter", "source-finance-scatter", 0, "th percentile")
        add("FFO per student", context.get("ffo_per_student"), "Finance summary cards", "ffo_per_student", "source-finance-summary", 0)
        add("Operating cost per student", context.get("operating_cost_per_student"), "Finance summary cards", "operating_cost_per_student", "source-finance-summary", 0)
        add("Financial score", context.get("financial_score"), "Finance summary cards", "financial_score", "source-finance-summary", 1)
        add("Overall score", context.get("overall_score"), "Finance summary cards", "overall_score", "source-finance-summary", 1)
        add("Personnel cost share", context.get("personnel_cost_share"), "Financial structure", "personnel_cost_share", "source-finance-indicators", 2)
        add("Public revenue share", context.get("public_revenue_share"), "Financial structure", "public_revenue_share", "source-finance-indicators", 2)
        add("Student contribution share", context.get("student_contribution_share"), "Financial structure", "student_contribution_share", "source-finance-indicators", 2)
        add("Performance quota share", context.get("performance_quota_share"), "Financial structure", "performance_quota_share", "source-finance-indicators", 2)
        add("Economic-financial sustainability index", context.get("economic_financial_sustainability_index"), "Financial structure", "economic_financial_sustainability_index", "source-finance-indicators", 2)

    elif view == "DEA Efficiency Explorer":
        add("DEA-VRS efficiency", context.get("dea_vrs_efficiency_100"), "DEA summary", "dea_vrs_efficiency_100", "source-dea-summary", 1)
        add("DEA-CRS efficiency", context.get("dea_crs_efficiency_100"), "DEA summary", "dea_crs_efficiency_100", "source-dea-summary", 1)
        add("Scale efficiency", context.get("dea_scale_efficiency_100"), "DEA summary", "dea_scale_efficiency_100", "source-dea-summary", 1)
        add("Overall score", context.get("overall_score"), "DEA vs performance context", "overall_score", "source-dea-summary", 1)
        add(context.get("dea_scatter_x_axis", "DEA scatterplot x-axis value"), context.get("dea_scatter_x_value"), "DEA scatterplot", "dea_scatter_x_value", "source-dea-scatter", 2)
        add("DEA x-axis median in current filter", context.get("dea_scatter_x_median_current_filter"), "DEA scatterplot benchmark", "dea_scatter_x_median_current_filter", "source-dea-scatter", 2)
        add("DEA x-axis percentile in current filter", context.get("dea_scatter_x_percentile_current_filter"), "DEA scatterplot position", "dea_scatter_x_percentile_current_filter", "source-dea-scatter", 0, "th percentile")
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
        add("Inactive students reversed score", context.get("inactive_students_reversed_score"), "Teaching indicators", "inactive_students_reversed_score", "source-teaching-indicators", 1)
        add("Graduation within standard duration", context.get("graduation_within_standard_pct"), "Teaching indicators", "graduation_within_standard_pct", "source-teaching-indicators", 1, "%")
        add("Graduation intensity", context.get("graduation_intensity"), "Teaching indicators", "graduation_intensity", "source-teaching-indicators", 2)
        add("Employment index", context.get("employment_index"), "Teaching indicators", "employment_index", "source-teaching-indicators", 1)
        add("Publications per teaching staff", context.get("publications_per_teaching_staff"), "Research indicators", "publications_per_teaching_staff", "source-research-indicators", 2)
        add("Citations per publication", context.get("citations_per_publication"), "Research indicators", "citations_per_publication", "source-research-indicators", 2)
        add("H-index", context.get("h_index"), "Research indicators", "h_index", "source-research-indicators", 1)
        add("Highly cited researchers", context.get("highly_cited_researchers"), "Research indicators", "highly_cited_researchers", "source-research-indicators", 1)
        add("Nature and Science articles", context.get("nature_science_articles"), "Research indicators", "nature_science_articles", "source-research-indicators", 1)

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
        add("Profile dispersion", context.get("profile_dispersion"), "Profile balance", "profile_dispersion", "source-profile-summary", 1)
        dynamic_benchmark = context.get("dynamic_benchmark", {})
        if dynamic_benchmark:
            add(f"{dynamic_benchmark.get('label', 'Benchmark')} overall average", dynamic_benchmark.get("overall_average"), "Dynamic benchmark comparison", "dynamic_benchmark.overall_average", "source-profile-benchmark", 1)
        else:
            add("National average overall score", context.get("national_avg_overall_score"), "Benchmark comparison", "national_avg_overall_score", "source-profile-benchmark", 1)
            add("Macro-area average overall score", context.get("macro_area_avg_overall_score"), "Benchmark comparison", "macro_area_avg_overall_score", "source-profile-benchmark", 1)

    return items



def select_required_evidence_items(context: dict, evidence_items: list[dict]) -> list[dict]:
    """Select the quantitative evidence that a useful Companion answer should explicitly use.

    The full evidence registry can be larger than a readable interpretation.  This helper
    defines a compact page-specific evidence set that is sufficient to support the actual
    analysis. The generated narrative is required to mention every item in this set with its
    visible value and an internal source ID. The ID is hidden from the user after the value is
    converted into a clickable link to the corresponding visualization.
    """
    if not evidence_items:
        return []

    by_field = {str(item.get("field", "")): item for item in evidence_items}
    view = context.get("view_type")
    required_fields: list[str] = []

    # A clicked/focused mark should take precedence over the normal page-wide evidence set.
    focus = context.get("focus_details") or {}
    interaction_focus = context.get("interaction_focus") or {}
    focus_field = str(focus.get("field") or interaction_focus.get("field") or "")
    if context.get("focus_mode") and focus_field and not focus_field.startswith("visual::"):
        if view == "University Comparison" and focus_field in {
            "overall_score", "teaching_score", "placement_score", "research_score", "financial_score"
        }:
            gap_key = {
                "overall_score": "overall_gap_a_minus_b",
                "teaching_score": "teaching_gap_a_minus_b",
                "placement_score": "placement_gap_a_minus_b",
                "research_score": "research_gap_a_minus_b",
                "financial_score": "financial_gap_a_minus_b",
            }[focus_field]
            required_fields = [
                f"university_a.{focus_field}",
                f"university_b.{focus_field}",
                f"score_gaps_a_minus_b.{gap_key}",
            ]
        elif view == "Time Dynamics / What Changed" and focus_field:
            required_fields = [
                f"start_values.{focus_field}",
                f"end_values.{focus_field}",
                f"changes_end_minus_start.{focus_field}",
            ]
        else:
            # Keep the clicked item plus the most useful context for the active page.
            exact = [f for f in by_field if f == focus_field or f.endswith(f".{focus_field}")]
            required_fields.extend(exact[:1])
            if view == "University Profile":
                required_fields += ["overall_score", "profile_dispersion"]
            elif view == "Finance Explorer":
                required_fields += ["financial_score", "overall_score", "x_axis_median_current_filter"]
            elif view == "Teaching and Research":
                required_fields += ["teaching_score", "placement_score", "research_score"]
            elif view == "DEA Efficiency Explorer":
                required_fields += ["dea_vrs_efficiency_100", "dea_crs_efficiency_100", "dea_scale_efficiency_100"]
            elif view == "Ranking Explorer":
                required_fields += ["selected_value", "selected_rank_in_current_filter", "selected_percentile_current_filter"]

    if not required_fields:
        if view == "Overview":
            required_fields = [
                "number_of_universities",
                "average_overall_score",
                "average_teaching_score",
                "average_placement_score",
                "average_research_score",
                "average_financial_score",
            ]
            if context.get("linked_brush"):
                required_fields += ["linked_brush.selected_universities", "linked_brush.average_overall_score"]
        elif view == "University Profile":
            required_fields = [
                "overall_score", "overall_rank_year", "teaching_score", "placement_score",
                "research_score", "financial_score", "profile_dispersion",
            ]
            if context.get("dynamic_benchmark"):
                required_fields.append("dynamic_benchmark.overall_average")
            else:
                required_fields += ["national_avg_overall_score", "macro_area_avg_overall_score"]
        elif view == "University Comparison":
            required_fields = [
                "university_a.overall_score", "university_b.overall_score",
                "score_gaps_a_minus_b.overall_gap_a_minus_b",
                "score_gaps_a_minus_b.teaching_gap_a_minus_b",
                "score_gaps_a_minus_b.placement_gap_a_minus_b",
                "score_gaps_a_minus_b.research_gap_a_minus_b",
                "score_gaps_a_minus_b.financial_gap_a_minus_b",
            ]
        elif view == "Finance Explorer":
            required_fields = [
                "selected_x_value", "selected_y_value",
                "x_axis_median_current_filter", "y_axis_median_current_filter",
                "x_axis_percentile_current_filter", "y_axis_percentile_current_filter",
                "financial_score", "overall_score",
            ]
        elif view == "DEA Efficiency Explorer":
            required_fields = [
                "dea_vrs_efficiency_100", "dea_crs_efficiency_100", "dea_scale_efficiency_100",
                "overall_score", "dea_scatter_x_value", "dea_scatter_x_median_current_filter",
                "dea_scatter_x_percentile_current_filter", "dea_vrs_efficiency_100_percentile_current_filter",
            ]
        elif view == "Ranking Explorer":
            required_fields = [
                "selected_value", "selected_rank_in_current_filter",
                "selected_percentile_current_filter", "number_of_universities",
            ]
        elif view == "Time Dynamics / What Changed":
            required_fields = [
                "start_values.overall_score", "end_values.overall_score",
                "changes_end_minus_start.overall_score",
            ]
            largest = context.get("largest_change_dimension")
            if largest:
                required_fields.append(f"changes_end_minus_start.{largest}")
            required_fields += [
                "changes_end_minus_start.teaching_score",
                "changes_end_minus_start.research_score",
                "changes_end_minus_start.financial_score",
            ]
        elif view == "Teaching and Research":
            required_fields = [
                "teaching_score", "placement_score", "research_score",
                "teaching_score_percentile_current_filter", "research_score_percentile_current_filter",
                "second_year_retention_pct", "graduation_within_standard_pct",
                "employment_index", "publications_per_teaching_staff", "citations_per_publication",
            ]
        elif view == "Data and Methodology":
            required_fields = ["current_filter_count"]

    selected: list[dict] = []
    seen: set[str] = set()
    for field in required_fields:
        item = by_field.get(field)
        if item and item["id"] not in seen:
            selected.append(item)
            seen.add(item["id"])

    # Visual-specific focus sometimes uses derived chart values that are not individual registry
    # fields. In that case keep a compact source-matched set rather than returning no evidence.
    if context.get("focus_mode") and focus_field.startswith("visual::") and not selected:
        visual_id = focus_field.split("::", 1)[-1]
        anchor_map = {
            "overview_average_profile": "source-overview-profile",
            "overview_linked_brush": "source-overview-brush",
            "profile_dimensions": "source-profile-dimensions",
            "profile_benchmark": "source-profile-benchmark",
            "comparison_scores": "source-comparison-scores",
            "comparison_gaps": "source-comparison-gaps",
            "finance_scatter": "source-finance-scatter",
            "finance_indicators": "source-finance-indicators",
            "dea_scatter": "source-dea-scatter",
            "dea_summary": "source-dea-summary",
            "ranking_chart": "source-ranking-chart",
            "time_change": "source-time-change",
            "teaching_research": "source-teaching-research",
            "teaching_indicators": "source-teaching-indicators",
            "research_indicators": "source-research-indicators",
        }
        anchor = anchor_map.get(visual_id)
        if anchor:
            selected = [item for item in evidence_items if item.get("anchor") == anchor][:8]

    return selected


def annotate_inline_evidence_ids(answer: str, evidence_items: list[dict]) -> str:
    """Attach internal evidence IDs only to values that already appear in the prose.

    This function never appends a separate evidence/key-values paragraph. The IDs are used
    internally so the existing numeric value can be turned into a clickable deep link, and are
    hidden before the Companion response is rendered.
    """
    result = answer or ""
    if not result or not evidence_items:
        return result

    numeric_items = [
        item for item in evidence_items
        if item.get("value") is not None and not isinstance(item.get("value"), str)
    ]
    counts: dict[str, int] = {}
    for item in numeric_items:
        display = str(item.get("display_value", ""))
        if display:
            counts[display] = counts.get(display, 0) + 1

    # Only auto-annotate values that map unambiguously to one evidence item.
    # Repeated values (for example several DEA scores equal to 100.0) are left to the
    # model-provided [E#] IDs, which disambiguate the intended visual source.
    for item in numeric_items:
        display = str(item.get("display_value", ""))
        evidence_id = str(item.get("id", ""))
        if not display or not evidence_id or counts.get(display) != 1:
            continue
        if re.search(rf"\[{re.escape(evidence_id)}\]", result):
            continue
        pattern = rf"(?<![\w\[]){re.escape(display)}(?![\w\]])"
        match = re.search(pattern, result)
        if match:
            result = result[:match.end()] + f" [{evidence_id}]" + result[match.end():]

    return result


def extract_evidence_ids(answer: str, evidence_items: list[dict]) -> list[str]:
    valid = {item["id"] for item in evidence_items}
    seen: list[str] = []
    for evidence_id in re.findall(r"\[(E\d+)\]", answer or ""):
        if evidence_id in valid and evidence_id not in seen:
            seen.append(evidence_id)
    return seen


def evidence_deep_link(item: dict) -> str:
    anchor = item.get("anchor", "source-current-view")
    params = _navigation_state_params()
    params["focus_field"] = str(item.get("field", ""))
    # urlencode preserves spaces, slashes and comparison labels safely while keeping
    # the browser in the same dashboard state after navigation.
    return f"?{urlencode(params)}#{anchor}"


def _same_tab_value_link(display: str, item: dict) -> str:
    """Create an inline linked-value anchor that stays in the current browser tab."""
    url = html.escape(evidence_deep_link(item), quote=True)
    safe_display = html.escape(display)
    label = html.escape(str(item.get("label", "dashboard value")), quote=True)
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


def hide_internal_source_ids(answer: str) -> str:
    """Hide internal [E#] mapping tokens while preserving clickable numeric links."""
    cleaned = re.sub(r"\s*\[E\d+\]", "", answer or "")
    cleaned = re.sub(r"[ \t]+([,.;:])", r"\1", cleaned)
    return cleaned


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
                            block.classList.add('linked-scroll-pulse');
                            window.setTimeout(() => block.classList.remove('linked-scroll-pulse'), 1800);
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
    brush = context.get("linked_brush") or {}
    if brush:
        selected_n = brush.get("selected_universities")
        b_overall = _safe_float(brush.get("average_overall_score"))
        b_teaching = _safe_float(brush.get("average_teaching_score"))
        b_research = _safe_float(brush.get("average_research_score"))
        return f"""
**Overview analysis — brushed subgroup**

The current brush selects **{selected_n} universities** from the **{context.get('number_of_universities')}** universities in the filtered system. Their mean Overall score is **{format_number(b_overall,1)}**, compared with **{avg_overall:.1f}** for the full filtered group. Their mean Teaching and Research scores are **{format_number(b_teaching,1)}** and **{format_number(b_research,1)}**, versus **{avg_teaching:.1f}** and **{avg_research:.1f}** in the full group.

**Interpretation**

The brush defines a temporary analytical subgroup. If its Teaching and Research means are both above the full-group averages, it represents a jointly stronger academic cluster; if only one is above, the selected region of the scatterplot reflects specialization rather than uniformly stronger performance.

**What to inspect next**

Move the brush to another region and compare how the linked multidimensional profile changes. This is useful for testing whether the visible pattern is specific to one cluster of universities.

**Limit**

The brushed subgroup is user-defined and exploratory; it is not a statistically estimated cluster.{q}
"""
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
    dynamic_benchmark = context.get("dynamic_benchmark") or {}
    if dynamic_benchmark:
        selected_dims = {"Overall": overall, "Teaching": teaching, "Placement": placement, "Research": research, "Financial": financial}
        benchmark_dims = {
            "Overall": dynamic_benchmark.get("overall_average"), "Teaching": dynamic_benchmark.get("teaching_average"),
            "Placement": dynamic_benchmark.get("placement_average"), "Research": dynamic_benchmark.get("research_average"),
            "Financial": dynamic_benchmark.get("financial_average"),
        }
        gaps = {k: _metric_change(benchmark_dims.get(k), selected_dims.get(k)) for k in selected_dims}
        valid_gaps = {k:v for k,v in gaps.items() if _safe_float(v) is not None}
        largest_gap_dim = max(valid_gaps, key=lambda k: abs(float(valid_gaps[k]))) if valid_gaps else None
        benchmark_sentence = (
            f"Against **{dynamic_benchmark.get('label')}** ({dynamic_benchmark.get('n')} universities), the largest selected-minus-benchmark gap is "
            f"**{largest_gap_dim} {_fmt_signed(valid_gaps.get(largest_gap_dim))} points**."
            if largest_gap_dim else "The selected benchmark does not contain enough comparable data."
        )
    else:
        benchmark_sentence = f"Relative to the national average, the overall gap is **{national_gap:+.1f} points**; relative to the macro-area average, it is **{macro_gap:+.1f} points**."
    return f"""
**University Profile analysis**

**{context.get('university')}** in **{context.get('year')}** has an overall score of **{overall:.1f}** and rank **{rank}/61**. This places it in a **{score_band_analysis(overall)}** position in the dashboard ranking.

**Benchmark interpretation**

{benchmark_sentence} This means the selected university is being read relative to an explicit contextual reference group rather than in isolation.

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


def _visual_id_from_focus(focus: dict) -> str:
    field = str(focus.get("field") or "")
    return field.split("visual::", 1)[1] if field.startswith("visual::") else field


def _safe_float(value: Any) -> float | None:
    try:
        if value is None or pd.isna(value):
            return None
        return float(value)
    except Exception:
        return None


def _metric_change(start: Any, end: Any) -> float | None:
    a, b = _safe_float(start), _safe_float(end)
    return None if a is None or b is None else b - a


def build_visual_focus_details(context: dict, focus: dict) -> dict:
    """Build data-rich context for every 'Explain this visual' button.

    This is intentionally deterministic: even if the external model is unavailable,
    the local Companion can explain the actual chart rather than merely acknowledge
    that the chart was selected.
    """
    visual_id = _visual_id_from_focus(focus)
    view = context.get("view_type")
    details: dict[str, Any] = {
        "mode": "visual",
        "focus_kind": "visual",
        "visual_id": visual_id,
        "view_type": view,
        "label": focus.get("label") or visual_id,
        "interaction": focus.get("interaction"),
    }

    # Overview visuals
    if visual_id == "overview_top10":
        rows = context.get("top_10_universities") or context.get("top_universities") or []
        details["ranking_rows"] = rows
        if rows:
            details["leader"] = rows[0]
            details["last_visible"] = rows[-1]
            details["visible_range"] = _metric_change(rows[-1].get("overall_score"), rows[0].get("overall_score"))
    elif visual_id == "overview_average_profile":
        profile = {
            "Teaching": context.get("average_teaching_score"),
            "Placement": context.get("average_placement_score"),
            "Research": context.get("average_research_score"),
            "Financial": context.get("average_financial_score"),
        }
        details["profile"] = profile
        valid = {k: _safe_float(v) for k, v in profile.items() if _safe_float(v) is not None}
        if valid:
            details["strongest_dimension"] = max(valid, key=valid.get)
            details["weakest_dimension"] = min(valid, key=valid.get)
            details["spread"] = max(valid.values()) - min(valid.values())
        details["overall_average"] = context.get("average_overall_score")
    elif visual_id == "overview_macro_area":
        rows = context.get("macro_area_profiles") or []
        details["macro_area_profiles"] = rows
        if rows:
            details["highest_overall_macro_area"] = max(rows, key=lambda r: _safe_float(r.get("overall_score")) or -1e9)
            details["lowest_overall_macro_area"] = min(rows, key=lambda r: _safe_float(r.get("overall_score")) or 1e9)
    elif visual_id == "overview_linked_brush":
        details["linked_brush"] = context.get("linked_brush")
        details["number_of_universities"] = context.get("number_of_universities")
        details["full_average_overall"] = context.get("average_overall_score")
        details["full_average_teaching"] = context.get("average_teaching_score")
        details["full_average_research"] = context.get("average_research_score")
        details["teaching_research_correlation"] = context.get("teaching_research_correlation")
        details["teaching_range"] = [context.get("teaching_min"), context.get("teaching_max")]
        details["research_range"] = [context.get("research_min"), context.get("research_max")]

    # University Profile visuals
    elif visual_id == "profile_dimensions":
        profile = {k: context.get(v) for k, v in {
            "Teaching": "teaching_score", "Placement": "placement_score", "Research": "research_score", "Financial": "financial_score"
        }.items()}
        details.update({"university": context.get("university"), "overall_score": context.get("overall_score"), "profile": profile,
                        "profile_dispersion": context.get("profile_dispersion"), "profile_type": context.get("profile_type")})
        valid = {k: _safe_float(v) for k, v in profile.items() if _safe_float(v) is not None}
        if valid:
            details["strongest_dimension"] = max(valid, key=valid.get)
            details["weakest_dimension"] = min(valid, key=valid.get)
            details["spread"] = max(valid.values()) - min(valid.values())
    elif visual_id == "profile_benchmark":
        bench = context.get("dynamic_benchmark") or {}
        details.update({"university": context.get("university"), "benchmark": bench, "selected_scores": {
            "Overall": context.get("overall_score"), "Teaching": context.get("teaching_score"), "Placement": context.get("placement_score"),
            "Research": context.get("research_score"), "Financial": context.get("financial_score")}})
        if bench:
            details["gaps_selected_minus_benchmark"] = {
                "Overall": _metric_change(bench.get("overall_average"), context.get("overall_score")),
                "Teaching": _metric_change(bench.get("teaching_average"), context.get("teaching_score")),
                "Placement": _metric_change(bench.get("placement_average"), context.get("placement_score")),
                "Research": _metric_change(bench.get("research_average"), context.get("research_score")),
                "Financial": _metric_change(bench.get("financial_average"), context.get("financial_score")),
            }
    elif visual_id.endswith("_breakdown"):
        details.update({"university": context.get("university"), "score_breakdown": context.get("score_breakdown") or {}})
    elif visual_id == "weight_sensitivity":
        rows = context.get("weight_sensitivity") or []
        details.update({"university": context.get("university"), "scenarios": rows})
        if rows:
            valid = [r for r in rows if _safe_float(r.get("Score")) is not None]
            if valid:
                details["highest_score_scenario"] = max(valid, key=lambda r: float(r.get("Score")))
                details["lowest_score_scenario"] = min(valid, key=lambda r: float(r.get("Score")))
                valid_rank = [r for r in valid if _safe_float(r.get("Rank")) is not None]
                if valid_rank:
                    details["best_rank_scenario"] = min(valid_rank, key=lambda r: float(r.get("Rank")))
                    details["worst_rank_scenario"] = max(valid_rank, key=lambda r: float(r.get("Rank")))
    elif visual_id == "profile_trend":
        details.update({"university": context.get("university"), "trend_summary": context.get("profile_trend_summary") or {}})

    # Comparison visuals
    elif visual_id in {"comparison_scores", "comparison_gaps"}:
        a, b = context.get("university_a") or {}, context.get("university_b") or {}
        gaps = context.get("score_gaps_a_minus_b") or {}
        details.update({"university_a": a, "university_b": b, "gaps": gaps})
        gap_pairs = {
            "Overall": gaps.get("overall_gap_a_minus_b"), "Teaching": gaps.get("teaching_gap_a_minus_b"),
            "Placement": gaps.get("placement_gap_a_minus_b"), "Research": gaps.get("research_gap_a_minus_b"),
            "Financial": gaps.get("financial_gap_a_minus_b")}
        valid = {k: _safe_float(v) for k, v in gap_pairs.items() if _safe_float(v) is not None}
        if valid:
            details["largest_gap_dimension"] = max(valid, key=lambda k: abs(valid[k]))
            details["largest_gap"] = valid[details["largest_gap_dimension"]]
    elif visual_id == "comparison_benchmark":
        details.update({"selected_vs_benchmark": context.get("selected_vs_benchmark") or {}})

    # Finance visuals
    elif visual_id == "finance_scatter":
        details.update({
            "university": context.get("university"), "x_label": context.get("finance_x_axis"), "y_label": context.get("score_y_axis"),
            "x_value": context.get("selected_x_value"), "y_value": context.get("selected_y_value"),
            "x_median": context.get("x_axis_median_current_filter"), "y_median": context.get("y_axis_median_current_filter"),
            "x_percentile": context.get("x_axis_percentile_current_filter"), "y_percentile": context.get("y_axis_percentile_current_filter"),
            "cluster_position": context.get("scatter_cluster_position"),
        })
    elif visual_id == "finance_structure":
        details.update({"university": context.get("university"), "financial_score": context.get("financial_score"), "overall_score": context.get("overall_score"),
                        "indicators": {"Personnel cost share": context.get("personnel_cost_share"), "Public revenue share": context.get("public_revenue_share"),
                                       "Student contribution share": context.get("student_contribution_share"), "Performance quota share": context.get("performance_quota_share"),
                                       "Economic-financial sustainability index": context.get("economic_financial_sustainability_index")}})

    # DEA visuals
    elif visual_id == "dea_scatter":
        details.update({
            "university": context.get("university"), "x_label": context.get("dea_scatter_x_axis"), "x_value": context.get("dea_scatter_x_value"),
            "x_median": context.get("dea_scatter_x_median_current_filter"), "x_percentile": context.get("dea_scatter_x_percentile_current_filter"),
            "vrs": context.get("dea_vrs_efficiency_100"), "vrs_median": context.get("dea_vrs_efficiency_100_median_current_filter"),
            "vrs_percentile": context.get("dea_vrs_efficiency_100_percentile_current_filter"), "category": context.get("efficiency_category"),
            "overall_score": context.get("overall_score"),
        })
    elif visual_id == "dea_trend":
        details.update({"university": context.get("university"), "trend_summary": context.get("dea_trend_summary") or {}})
    elif visual_id == "dea_top10":
        rows = context.get("dea_top10") or []
        details.update({"rows": rows, "selected_university": context.get("university"), "selected_vrs": context.get("dea_vrs_efficiency_100"),
                        "selected_rank": context.get("dea_vrs_rank_year")})

    # Ranking visual
    elif visual_id == "ranking_chart":
        details.update({"metric": context.get("metric"), "selected_university": context.get("selected_university"), "selected_value": context.get("selected_value"),
                        "rank": context.get("selected_rank_in_current_filter"), "n": context.get("number_of_universities"),
                        "percentile": context.get("selected_percentile_current_filter"), "top_universities": context.get("top_universities") or []})

    # Time visuals
    elif visual_id in {"time_change", "time_slope"}:
        details.update({"university": context.get("university"), "start_year": context.get("start_year"), "end_year": context.get("end_year"),
                        "start_values": context.get("start_values") or {}, "end_values": context.get("end_values") or {},
                        "changes": context.get("changes_end_minus_start") or {}, "largest_change_dimension": context.get("largest_change_dimension")})

    # Teaching & research visuals
    elif visual_id == "teaching_research_scores":
        profile = {"Teaching": context.get("teaching_score"), "Placement": context.get("placement_score"), "Research": context.get("research_score")}
        details.update({"university": context.get("university"), "profile": profile})
        valid = {k: _safe_float(v) for k, v in profile.items() if _safe_float(v) is not None}
        if valid:
            details["strongest_dimension"] = max(valid, key=valid.get)
            details["weakest_dimension"] = min(valid, key=valid.get)
    elif visual_id == "teaching_indicators":
        details.update({"university": context.get("university"), "teaching_score": context.get("teaching_score"), "placement_score": context.get("placement_score"),
                        "indicators": {"Second-year retention": context.get("second_year_retention_pct"), "Inactive students reversed score": context.get("inactive_students_reversed_score"),
                                       "Graduation within standard duration": context.get("graduation_within_standard_pct"), "Graduation intensity": context.get("graduation_intensity"),
                                       "Employment index": context.get("employment_index")},
                        "percentiles": {"Retention": context.get("second_year_retention_pct_percentile_current_filter"),
                                        "Graduation": context.get("graduation_within_standard_pct_percentile_current_filter"),
                                        "Employment": context.get("employment_index_percentile_current_filter")}})
    elif visual_id == "research_indicators":
        details.update({"university": context.get("university"), "research_score": context.get("research_score"),
                        "indicators": {"Publications per teaching staff": context.get("publications_per_teaching_staff"), "Citations per publication": context.get("citations_per_publication"),
                                       "H-index": context.get("h_index"), "Highly cited researchers": context.get("highly_cited_researchers"), "Nature and Science articles": context.get("nature_science_articles")},
                        "percentiles": {"Publications/staff": context.get("publications_per_teaching_staff_percentile_current_filter"),
                                        "Citations/publication": context.get("citations_per_publication_percentile_current_filter"), "H-index": context.get("h_index_percentile_current_filter")}})
    elif visual_id == "teaching_research_scatter":
        details.update({"university": context.get("university"), "teaching_score": context.get("teaching_score"), "research_score": context.get("research_score"),
                        "teaching_percentile": context.get("teaching_score_percentile_current_filter"), "research_percentile": context.get("research_score_percentile_current_filter"),
                        "position": context.get("teaching_research_position")})

    return details


def build_focus_details(context: dict) -> dict | None:
    """Create a compact, page-specific description of the clicked visual mark.

    The raw chart callback tells us which field/bar was clicked. This helper adds
    the comparison values needed for a genuinely focused Companion response.
    """
    focus = context.get("interaction_focus") or {}
    if not focus:
        return None

    field = str(focus.get("field") or "")
    label = str(focus.get("label") or field)
    value = focus.get("value")
    view = context.get("view_type")

    if focus.get("focus_kind") == "visual" or field.startswith("visual::") or focus.get("interaction") == "explain visual request":
        return build_visual_focus_details(context, focus)

    details: dict[str, Any] = {
        "mode": "clicked_mark",
        "view_type": view,
        "field": field,
        "label": label,
        "clicked_value": value,
        "interaction": focus.get("interaction"),
    }

    score_labels = {
        "overall_score": "Overall",
        "teaching_score": "Teaching",
        "placement_score": "Placement",
        "research_score": "Research",
        "financial_score": "Financial",
    }

    if view == "University Comparison" and context.get("university_a") and context.get("university_b"):
        a = context["university_a"]
        b = context["university_b"]
        if field in score_labels:
            a_val = clean_value(a.get(field))
            b_val = clean_value(b.get(field))
            gap = None
            try:
                gap = float(a.get(field)) - float(b.get(field))
            except Exception:
                pass
            details.update({
                "dimension": score_labels[field],
                "university_a": a.get("university"),
                "university_b": b.get("university"),
                "value_a": a_val,
                "value_b": b_val,
                "gap_a_minus_b": gap,
                "clicked_university": focus.get("university"),
                "overall_gap_a_minus_b": context.get("score_gaps_a_minus_b", {}).get("overall_gap_a_minus_b"),
            })
        return details

    if view == "University Profile":
        if field in score_labels:
            details.update({
                "dimension": score_labels[field],
                "university": context.get("university"),
                "value": context.get(field),
                "overall_score": context.get("overall_score"),
                "national_avg_overall_score": context.get("national_avg_overall_score"),
                "macro_area_avg_overall_score": context.get("macro_area_avg_overall_score"),
                "profile_dispersion": context.get("profile_dispersion"),
            })
        return details

    if view == "Finance Explorer":
        details.update({
            "university": context.get("university"),
            "value": context.get(field, value),
            "financial_score": context.get("financial_score"),
            "overall_score": context.get("overall_score"),
            "finance_x_axis": context.get("finance_x_axis"),
            "score_y_axis": context.get("score_y_axis"),
            "x_axis_median_current_filter": context.get("x_axis_median_current_filter"),
            "y_axis_median_current_filter": context.get("y_axis_median_current_filter"),
        })
        return details

    if view == "Teaching and Research":
        details.update({
            "university": context.get("university"),
            "value": context.get(field, value),
            "teaching_score": context.get("teaching_score"),
            "placement_score": context.get("placement_score"),
            "research_score": context.get("research_score"),
        })
        return details

    if view == "Time Dynamics / What Changed":
        starts = context.get("start_values", {})
        ends = context.get("end_values", {})
        changes = context.get("changes_end_minus_start", {})
        details.update({
            "university": context.get("university"),
            "start_year": context.get("start_year"),
            "end_year": context.get("end_year"),
            "start_value": starts.get(field),
            "end_value": ends.get(field),
            "change": changes.get(field, value),
        })
        return details

    if view == "Overview":
        details.update({"year": context.get("year"), "number_of_universities": context.get("number_of_universities")})
        return details

    if view == "DEA Efficiency Explorer":
        details.update({
            "university": context.get("university"),
            "value": context.get(field, value),
            "dea_vrs_efficiency_100": context.get("dea_vrs_efficiency_100"),
            "dea_crs_efficiency_100": context.get("dea_crs_efficiency_100"),
            "dea_scale_efficiency_100": context.get("dea_scale_efficiency_100"),
            "overall_score": context.get("overall_score"),
        })
        return details

    if view == "Ranking Explorer":
        details.update({
            "university": context.get("university"),
            "metric": context.get("metric"),
            "selected_value": context.get("selected_value"),
            "rank": context.get("selected_rank_in_current_filter"),
            "percentile": context.get("selected_percentile_current_filter"),
        })
        return details

    return details


def _fmt_signed(value: Any, decimals: int = 1) -> str:
    v = _safe_float(value)
    return "n/a" if v is None else f"{v:+.{decimals}f}"


def _compact_pairs(mapping: dict, decimals: int = 1) -> str:
    parts = []
    for key, value in mapping.items():
        v = _safe_float(value)
        if v is not None:
            parts.append(f"**{key} {v:.{decimals}f}**")
    return ", ".join(parts)


def local_visual_interpretation(context: dict, user_question: str | None = None) -> str:
    """Explain the actual pattern in the visual selected by 'Explain this visual'."""
    focus = context.get("focus_details") or build_focus_details(context) or {}
    vid = focus.get("visual_id")
    label = focus.get("label") or "selected visual"
    q = f"\n\n**User question**\n\n{user_question}" if user_question else ""

    if vid == "overview_top10":
        rows = focus.get("ranking_rows") or []
        if rows:
            leader = rows[0]
            tail = rows[-1]
            gap = (_safe_float(leader.get("overall_score")) or 0) - (_safe_float(tail.get("overall_score")) or 0)
            top3 = ", ".join(f"{r.get('university')} ({format_number(r.get('overall_score'),1)})" for r in rows[:3])
            return f"""**Visual analysis — {label}**

The visible leaders are **{top3}**. The first university, **{leader.get('university')}**, has an overall score of **{format_number(leader.get('overall_score'),1)}**, while the last university shown in this Top-{len(rows)} view has **{format_number(tail.get('overall_score'),1)}**. The visible spread is therefore about **{gap:.1f} points**.

**Interpretation**

This chart is useful for locating the upper end of the current filtered distribution, but it should not be read as a complete quality hierarchy. The overall score aggregates Teaching, Placement, Research and Financial dimensions, so similar bar lengths can conceal very different profiles.

**What to inspect next**

Click one of the bars and open its University Profile to see which dimensions produce the position shown here.

**Limit**

The ranking is specific to the selected year and filters and uses the dashboard's normalized overall score.{q}"""

    if vid == "overview_average_profile":
        profile = focus.get("profile") or {}
        return f"""**Visual analysis — {label}**

The current filtered group has the following average profile: {_compact_pairs(profile)}. The strongest average dimension is **{focus.get('strongest_dimension')}**, while the weakest is **{focus.get('weakest_dimension')}**. Their difference is **{format_number(focus.get('spread'),1)} points**; the average overall score is **{format_number(focus.get('overall_average'),1)}**.

**Interpretation**

The chart shows the internal shape of the filtered system, not just its average overall level. A visible spread means that the institutions in the current selection collectively perform more strongly in some dimensions than others, so the overall average alone would hide that structure.

**What to inspect next**

Use the macro-area chart or Linked Brushing Explorer to check whether the same strong/weak pattern is shared across geographical groups or is driven by a subset of universities.

**Limit**

These are group averages and can hide substantial institution-level heterogeneity.{q}"""

    if vid == "overview_macro_area":
        high, low = focus.get("highest_overall_macro_area") or {}, focus.get("lowest_overall_macro_area") or {}
        rows = focus.get("macro_area_profiles") or []
        profile_text = "; ".join(f"{r.get('macro_area')}: overall {format_number(r.get('overall_score'),1)}" for r in rows)
        return f"""**Visual analysis — {label}**

Across the macro-areas shown, the highest average overall score is **{high.get('macro_area')} ({format_number(high.get('overall_score'),1)})**, while the lowest is **{low.get('macro_area')} ({format_number(low.get('overall_score'),1)})**. The current overall averages are: {profile_text}.

**Interpretation**

The grouped bars show that geographical averages are multidimensional: a macro-area can lead in one score and not in another. The useful reading is therefore the profile composition within each macro-area, rather than a single geographical ranking.

**What to inspect next**

Apply one macro-area as a filter and compare individual universities inside it; this separates between-area differences from institutional heterogeneity within the area.

**Limit**

These are descriptive averages and do not establish territorial causes of performance differences.{q}"""

    if vid == "overview_linked_brush":
        brush = focus.get("linked_brush") or {}
        if brush:
            return f"""**Visual analysis — {label}**

The current brush selects **{brush.get('selected_universities')} universities**. Their mean Teaching score is **{format_number(brush.get('average_teaching_score'),1)}**, mean Research score is **{format_number(brush.get('average_research_score'),1)}**, and mean Overall score is **{format_number(brush.get('average_overall_score'),1)}**. For the full filtered group, the corresponding averages are Teaching **{format_number(focus.get('full_average_teaching'),1)}**, Research **{format_number(focus.get('full_average_research'),1)}**, and Overall **{format_number(focus.get('full_average_overall'),1)}**.

**Interpretation**

The brush is doing more than highlighting points: it defines a temporary analytical subgroup. Comparing the brushed means with the full-sample means shows whether the selected cluster represents a relatively teaching-strong, research-strong, jointly strong, or weaker profile.

**What to inspect next**

Move the brush to a different part of the scatterplot and compare how the linked multidimensional profile changes.

**Limit**

The brushed group is user-defined and exploratory; it is not a statistically estimated cluster.{q}"""
        return f"""**Visual analysis — {label}**

This scatterplot positions **{focus.get('number_of_universities')} universities** by Teaching and Research. Teaching ranges from **{format_number((focus.get('teaching_range') or [None,None])[0],1)}** to **{format_number((focus.get('teaching_range') or [None,None])[1],1)}**, while Research ranges from **{format_number((focus.get('research_range') or [None,None])[0],1)}** to **{format_number((focus.get('research_range') or [None,None])[1],1)}**. The descriptive Teaching–Research correlation is **{format_number(focus.get('teaching_research_correlation'),2)}**.

**Interpretation**

The purpose of the visual is to expose joint positioning and heterogeneity. Dragging a rectangle creates a temporary subgroup whose linked profile can then be compared with the whole filtered system.

**What to inspect next**

Brush a dense central region and then an extreme region to see how their average multidimensional profiles differ.

**Limit**

The correlation and spatial pattern are descriptive and do not imply a causal relationship between Teaching and Research.{q}"""

    if vid == "profile_dimensions":
        profile = focus.get("profile") or {}
        return f"""**Visual analysis — {label}**

For **{focus.get('university')}**, the profile is {_compact_pairs(profile)} with an Overall score of **{format_number(focus.get('overall_score'),1)}**. The strongest dimension is **{focus.get('strongest_dimension')}** and the weakest is **{focus.get('weakest_dimension')}**, separated by **{format_number(focus.get('spread'),1)} points**. The profile dispersion is **{format_number(focus.get('profile_dispersion'),1)}**.

**Interpretation**

This is the core multidimensional view: it shows whether the overall score is supported evenly across dimensions or produced by specialization. A large spread means that the single overall value masks meaningful internal differences.

**What to inspect next**

Click the strongest or weakest bar, or open **Why this score?**, to trace the composite dimension back to its component indicators.

**Limit**

The bars are normalized dashboard scores and should not be interpreted as official external ratings.{q}"""

    if vid == "profile_benchmark":
        bench = focus.get("benchmark") or {}
        gaps = focus.get("gaps_selected_minus_benchmark") or {}
        valid = {k:v for k,v in gaps.items() if _safe_float(v) is not None}
        largest = max(valid, key=lambda k: abs(float(valid[k]))) if valid else None
        return f"""**Visual analysis — {label}**

The selected university is compared with **{bench.get('label')}** (**{bench.get('n')} universities**). The selected-minus-benchmark gaps are {_compact_pairs(valid)}. The largest absolute difference is in **{largest}** at **{_fmt_signed(valid.get(largest))} points**.

**Interpretation**

This chart changes the meaning of comparison by changing the reference group. A positive gap means the selected university is above the chosen benchmark average in that dimension; a negative gap means it is below. The most important information is not only the direction of the overall gap but whether the same pattern persists across dimensions.

**What to inspect next**

Switch from National to Same size or Region. If the gaps change materially, institutional context is affecting the descriptive benchmark position.

**Limit**

Benchmark results depend on the peer definition and should not be treated as causal or normative judgments.{q}"""

    if vid and vid.endswith("_breakdown"):
        bd = focus.get("score_breakdown") or {}
        comps = bd.get("components") or []
        valid = [r for r in comps if _safe_float(r.get("Component score")) is not None]
        strongest = max(valid, key=lambda r: float(r.get("Component score"))) if valid else {}
        weakest = min(valid, key=lambda r: float(r.get("Component score"))) if valid else {}
        return f"""**Visual analysis — {label}**

The reported **{bd.get('dimension')} score is {format_number(bd.get('reported_score'),1)}**, and reconstruction from the displayed components gives **{format_number(bd.get('reconstructed_score'),1)}**. The strongest component is **{strongest.get('Component')} ({format_number(strongest.get('Component score'),1)})**, while the weakest is **{weakest.get('Component')} ({format_number(weakest.get('Component score'),1)})**.

**Interpretation**

The breakdown explains *why* the composite score has its current level. It separates the contribution of the underlying normalized indicators from the final dimension score, making the aggregation transparent rather than treating the score as a black box.

**What to inspect next**

Compare the weakest component's raw value and normalization method with the stronger components; this shows whether the limitation comes from the observed indicator level or from the way its scale is transformed.

**Limit**

Component contributions reflect the dashboard's specified normalization and equal-within-dimension aggregation rules.{q}"""

    if vid == "weight_sensitivity":
        high, low = focus.get("highest_score_scenario") or {}, focus.get("lowest_score_scenario") or {}
        best, worst = focus.get("best_rank_scenario") or {}, focus.get("worst_rank_scenario") or {}
        return f"""**Visual analysis — {label}**

Across the predefined weighting scenarios, the profile score ranges from **{format_number(low.get('Score'),1)}** under **{low.get('Scenario')}** to **{format_number(high.get('Score'),1)}** under **{high.get('Scenario')}**. The best observed rank is **{best.get('Rank')}** under **{best.get('Scenario')}**, while the weakest rank is **{worst.get('Rank')}** under **{worst.get('Scenario')}**.

**Interpretation**

This chart is a robustness check. If score and rank change little across scenarios, the university's relative profile is fairly stable to weighting choices. Larger movements indicate that the aggregate position depends more strongly on which dimension receives priority.

**What to inspect next**

Compare the scenario that improves the rank most with the university's strongest dimension; that reveals which weighting assumption is driving the sensitivity.

**Limit**

These are predefined exploratory scenarios, not alternative official rankings.{q}"""

    if vid == "profile_trend":
        ts = focus.get("trend_summary") or {}
        changes = ts.get("changes") or {}
        largest = ts.get("largest_change_dimension")
        return f"""**Visual analysis — {label}**

For **{focus.get('university')}**, this chart tracks the score profile from **{ts.get('start_year')}** to **{ts.get('end_year')}**. The Overall score changes by **{_fmt_signed(changes.get('overall_score'))} points**. The largest dimension movement is **{str(largest).replace('_score','').title() if largest else 'n/a'}**, changing by **{_fmt_signed(changes.get(largest))} points**.

**Interpretation**

The line chart separates overall stability from internal restructuring. A nearly flat Overall line can coexist with sizable movements in individual dimensions, meaning that gains in one area may offset declines in another.

**What to inspect next**

Use Time Dynamics / What Changed to compare the exact start and end values of the dimension with the largest movement.

**Limit**

The chart describes temporal change but does not identify its causes.{q}"""

    if vid in {"comparison_scores", "comparison_gaps"}:
        a, b, gaps = focus.get("university_a") or {}, focus.get("university_b") or {}, focus.get("gaps") or {}
        dim = focus.get("largest_gap_dimension")
        gap = focus.get("largest_gap")
        return f"""**Visual analysis — {label}**

The two universities differ most in **{dim}**, where the A-minus-B gap is **{_fmt_signed(gap)} points**. Their Overall scores are **{a.get('university')}: {format_number(a.get('overall_score'),1)}** and **{b.get('university')}: {format_number(b.get('overall_score'),1)}**. The full dimension gaps are Overall **{_fmt_signed(gaps.get('overall_gap_a_minus_b'))}**, Teaching **{_fmt_signed(gaps.get('teaching_gap_a_minus_b'))}**, Placement **{_fmt_signed(gaps.get('placement_gap_a_minus_b'))}**, Research **{_fmt_signed(gaps.get('research_gap_a_minus_b'))}**, Financial **{_fmt_signed(gaps.get('financial_gap_a_minus_b'))}**.

**Interpretation**

The chart shows whether the overall difference is broad or concentrated. If one dimension gap is much larger than the others, that dimension is the main descriptive separator; opposing positive and negative gaps indicate different specialization profiles rather than uniform dominance.

**What to inspect next**

Click the largest gap bar to switch from whole-chart interpretation to a focused dimension comparison.

**Limit**

These differences are descriptive and do not explain why the universities differ.{q}"""

    if vid == "comparison_benchmark":
        cmp = focus.get("selected_vs_benchmark") or {}
        gaps = cmp.get("dimension_gaps_selected_minus_benchmark") or {}
        valid = {str(k).replace('_score','').title():v for k,v in gaps.items() if _safe_float(v) is not None}
        largest = max(valid, key=lambda k: abs(float(valid[k]))) if valid else None
        return f"""**Visual analysis — {label}**

The active university **{cmp.get('selected_university')}** is compared with **{cmp.get('benchmark')}** (**{cmp.get('benchmark_n')} universities**). The dimension gaps are {_compact_pairs(valid)}. The largest departure from the benchmark is **{largest} ({_fmt_signed(valid.get(largest))} points)**.

**Interpretation**

This view answers a different question from A-vs-B comparison: it asks how the active university sits relative to a contextual peer average. The sign and size of each gap show where the profile is distinctive within that benchmark group.

**What to inspect next**

Change the benchmark definition and check whether the largest gap remains the same. Stability across peer definitions strengthens the descriptive pattern.

**Limit**

The result is conditional on the selected benchmark group.{q}"""

    if vid == "finance_scatter":
        xv, yv, xm, ym = focus.get("x_value"), focus.get("y_value"), focus.get("x_median"), focus.get("y_median")
        return f"""**Visual analysis — {label}**

**{focus.get('university')}** is plotted at **{focus.get('x_label')} = {format_number(xv,2)}** and **{focus.get('y_label')} = {format_number(yv,1)}**. In the current filtered group, the medians are **{format_number(xm,2)}** and **{format_number(ym,1)}**, respectively. The selected university is around the **{format_number(focus.get('x_percentile'),0)}th percentile** on the x-axis and **{format_number(focus.get('y_percentile'),0)}th percentile** on the y-axis, placing it **{focus.get('cluster_position')}**.

**Interpretation**

This visual should be read as a joint position, not as a causal relationship. The key question is whether the selected university combines an unusually high/low financial condition with an unusually high/low performance score compared with the current peers.

**What to inspect next**

Change either axis while keeping the same university selected. If the university remains unusual across several financial indicators, the pattern is more persistent than a single scatterplot position.

**Limit**

The scatterplot shows association only; movement along the financial axis is not a what-if intervention.{q}"""

    if vid == "finance_structure":
        return f"""**Visual analysis — {label}**

For **{focus.get('university')}**, the displayed financial structure is: {_compact_pairs(focus.get('indicators') or {}, 2)}. Its Financial score is **{format_number(focus.get('financial_score'),1)}**, compared with an Overall score of **{format_number(focus.get('overall_score'),1)}**.

**Interpretation**

These bars describe different financial concepts, so their raw heights should not be compared as if they shared one common unit. Their value is in exposing the composition of the university's resource/revenue structure and providing the raw context behind the normalized Financial score.

**What to inspect next**

Click an individual bar to focus on that indicator, then place the same indicator on the Finance scatterplot x-axis to see its peer position.

**Limit**

Financial variables are contextual/descriptive and should not be treated as direct causes of academic performance.{q}"""

    if vid == "dea_scatter":
        return f"""**Visual analysis — {label}**

**{focus.get('university')}** has **DEA-VRS = {format_number(focus.get('vrs'),1)}** and **{focus.get('x_label')} = {format_number(focus.get('x_value'),2)}**. Within the current filter, its DEA-VRS score is around the **{format_number(focus.get('vrs_percentile'),0)}th percentile** and the x-axis value around the **{format_number(focus.get('x_percentile'),0)}th percentile**. The corresponding medians are DEA-VRS **{format_number(focus.get('vrs_median'),1)}** and x-axis **{format_number(focus.get('x_median'),2)}**.

**Interpretation**

The point shows the university's resource/context position together with relative efficiency. A high x-axis value does not imply high efficiency: DEA evaluates the output-to-input relationship relative to the frontier, so universities with different resource levels can occupy very different efficiency positions.

**What to inspect next**

Change the x-axis variable and see whether the selected university's efficiency position remains distinctive relative to other resource indicators.

**Limit**

DEA-VRS is sample- and specification-dependent; the scatterplot is descriptive and does not identify causal effects.{q}"""

    if vid == "dea_trend":
        ts = focus.get("trend_summary") or {}
        ch = ts.get("changes") or {}
        return f"""**Visual analysis — {label}**

For **{focus.get('university')}**, DEA-VRS changes from **{format_number((ts.get('start_values') or {}).get('dea_vrs_efficiency_100'),1)}** to **{format_number((ts.get('end_values') or {}).get('dea_vrs_efficiency_100'),1)}** between **{ts.get('start_year')}** and **{ts.get('end_year')}** (**{_fmt_signed(ch.get('dea_vrs_efficiency_100'))} points**). Over the same period, DEA-CRS changes by **{_fmt_signed(ch.get('dea_crs_efficiency_100'))}** and Scale efficiency by **{_fmt_signed(ch.get('dea_scale_efficiency_100'))}**.

**Interpretation**

The three lines separate pure VRS efficiency, CRS efficiency and the scale component. If VRS is stable while CRS or scale efficiency changes, the movement is more closely associated with scale conditions than with the VRS frontier position.

**What to inspect next**

Compare the years where CRS and VRS diverge most and inspect the corresponding scale-efficiency value.

**Limit**

Year-to-year DEA changes are relative to each year's comparison frontier and are not causal effects.{q}"""

    if vid == "dea_top10":
        rows = focus.get("rows") or []
        leaders = ", ".join(f"{r.get('university')} ({format_number(r.get('dea_vrs_efficiency_100'),1)})" for r in rows[:5])
        frontier_count = sum(1 for r in rows if (_safe_float(r.get('dea_vrs_efficiency_100')) or 0) >= 99.95)
        return f"""**Visual analysis — {label}**

The visible leading institutions include **{leaders}**. Among the displayed Top-{len(rows)}, **{frontier_count}** have DEA-VRS approximately equal to 100. The active university **{focus.get('selected_university')}** has DEA-VRS **{format_number(focus.get('selected_vrs'),1)}** and yearly rank **{focus.get('selected_rank')}**.

**Interpretation**

A DEA score of 100 indicates frontier membership under the selected VRS specification; it does **not** imply a unique first place. Multiple universities can be efficient simultaneously because each can define a different part of the frontier.

**What to inspect next**

Click a frontier university and compare its inputs/outputs with a lower-efficiency university rather than interpreting the tied score as identical institutional performance.

**Limit**

The frontier depends on the selected variables and yearly sample.{q}"""

    if vid == "ranking_chart":
        top = focus.get("top_universities") or []
        leader = top[0] if top else {}
        return f"""**Visual analysis — {label}**

The current metric is **{focus.get('metric')}**. **{focus.get('selected_university')}** has **{format_number(focus.get('selected_value'),1)}**, rank **{focus.get('rank')}/{focus.get('n')}**, around the **{format_number(focus.get('percentile'),0)}th percentile**. The visible leader is **{leader.get('university')}** with **{format_number(next((v for k,v in leader.items() if k != 'university'), None),1)}**.

**Interpretation**

This chart is metric-specific. A university can move substantially when the metric switches from Overall to Research, Financial or DEA efficiency, which is precisely why the dashboard keeps ranking views separate from the multidimensional profile.

**What to inspect next**

Switch the ranking metric while keeping the same university active and compare the change in position.

**Limit**

The ranking depends on the chosen metric, filters and year; it is not an overall judgment of institutional quality.{q}"""

    if vid in {"time_change", "time_slope"}:
        ch = focus.get("changes") or {}
        largest = focus.get("largest_change_dimension")
        return f"""**Visual analysis — {label}**

For **{focus.get('university')}**, the chart compares **{focus.get('start_year')}** with **{focus.get('end_year')}**. Overall changes by **{_fmt_signed(ch.get('overall_score'))} points**, Teaching by **{_fmt_signed(ch.get('teaching_score'))}**, Placement by **{_fmt_signed(ch.get('placement_score'))}**, Research by **{_fmt_signed(ch.get('research_score'))}**, and Financial by **{_fmt_signed(ch.get('financial_score'))}**. The largest absolute movement is **{str(largest).replace('_score','').title() if largest else 'n/a'}**.

**Interpretation**

The visual distinguishes broad movement from internal rebalancing. If dimensions move in opposing directions, a small Overall change can conceal substantial restructuring inside the profile; if they move together, the temporal pattern is more consistent across dimensions.

**What to inspect next**

Click the bar for the largest change to focus the Companion on its exact start value, end value and change.

**Limit**

These are descriptive changes between observed years, not causal effects.{q}"""

    if vid == "teaching_research_scores":
        profile = focus.get("profile") or {}
        return f"""**Visual analysis — {label}**

For **{focus.get('university')}**, the three displayed scores are {_compact_pairs(profile)}. The strongest is **{focus.get('strongest_dimension')}** and the weakest is **{focus.get('weakest_dimension')}**.

**Interpretation**

This chart reveals whether the academic profile is teaching-oriented, placement-oriented, research-oriented, or relatively balanced across the three dimensions. It should be read as a profile shape rather than as a single league-table position.

**What to inspect next**

Click a score bar, then inspect its underlying indicator block to see which raw measures support that composite score.

**Limit**

The three bars are dashboard composites based on the specified normalization and aggregation rules.{q}"""

    if vid == "teaching_indicators":
        pcts = focus.get("percentiles") or {}
        return f"""**Visual analysis — {label}**

For **{focus.get('university')}**, the teaching/placement indicators shown are: {_compact_pairs(focus.get('indicators') or {},2)}. The Teaching score is **{format_number(focus.get('teaching_score'),1)}** and the Placement score is **{format_number(focus.get('placement_score'),1)}**. Within the current filter, retention is around the **{format_number(pcts.get('Retention'),0)}th percentile**, graduation around the **{format_number(pcts.get('Graduation'),0)}th percentile**, and employment around the **{format_number(pcts.get('Employment'),0)}th percentile**.

**Interpretation**

The indicator bars explain which student-lifecycle measures are relatively strong or weak behind the composite Teaching and Placement scores. Raw magnitudes should be interpreted according to their own definitions rather than compared directly across unlike units.

**What to inspect next**

Click the indicator with the lowest peer percentile to obtain a focused explanation of that specific component.

**Limit**

These are descriptive indicators and do not identify causal determinants of student outcomes.{q}"""

    if vid == "research_indicators":
        pcts = focus.get("percentiles") or {}
        return f"""**Visual analysis — {label}**

For **{focus.get('university')}**, the research indicators shown are: {_compact_pairs(focus.get('indicators') or {},2)}. The composite Research score is **{format_number(focus.get('research_score'),1)}**. Within the current filter, publications per teaching staff are around the **{format_number(pcts.get('Publications/staff'),0)}th percentile**, citations per publication around the **{format_number(pcts.get('Citations/publication'),0)}th percentile**, and H-index around the **{format_number(pcts.get('H-index'),0)}th percentile**.

**Interpretation**

The chart separates research volume, citation impact and visibility-related indicators. A strong composite Research score can therefore arise from several different configurations rather than one single research measure.

**What to inspect next**

Click the indicator that looks most distinctive relative to peers and compare it with the score breakdown on University Profile.

**Limit**

The indicators have different scales and meanings; raw bar heights should not be treated as directly comparable units.{q}"""

    if vid == "teaching_research_scatter":
        return f"""**Visual analysis — {label}**

**{focus.get('university')}** has Teaching **{format_number(focus.get('teaching_score'),1)}** and Research **{format_number(focus.get('research_score'),1)}**. Within the current filtered group, these correspond to approximately the **{format_number(focus.get('teaching_percentile'),0)}th** and **{format_number(focus.get('research_percentile'),0)}th percentiles**, placing the university **{focus.get('position')}**.

**Interpretation**

The scatterplot makes specialization visible. A point high on Research but closer to the middle on Teaching indicates a research-oriented profile relative to current peers; the reverse indicates stronger teaching positioning. Similar percentiles suggest a more balanced academic position.

**What to inspect next**

Click a nearby university point to compare whether a similar overall academic position is produced by the same Teaching–Research balance.

**Limit**

This is relative positioning within the active filter, not a causal relationship between Teaching and Research.{q}"""

    return f"""**Visual analysis — {label}**

The selected chart is linked to the current **{context.get('view_type')}** context, but there is not enough chart-specific data in the current state to provide a reliable numerical interpretation.

**What to inspect next**

Use the chart's selections or filters to create a more specific analytical state, then request the explanation again.

**Limit**

The Companion will not invent values that are not present in the dashboard context.{q}"""


def local_focused_interpretation(context: dict, user_question: str | None = None) -> str:
    """Short interpretation of the exact bar/indicator the user clicked."""
    focus = context.get("focus_details") or build_focus_details(context) or {}
    if focus.get("focus_kind") == "visual" or focus.get("mode") == "visual":
        return local_visual_interpretation(context, user_question)
    view = context.get("view_type")
    label = focus.get("label") or focus.get("field") or "selected mark"
    clicked = focus.get("clicked_value")
    q = f"\n\n**User question**\n\n{user_question}" if user_question else ""

    if view == "University Comparison" and focus.get("dimension"):
        a_name, b_name = focus.get("university_a"), focus.get("university_b")
        a_val, b_val = focus.get("value_a"), focus.get("value_b")
        gap = focus.get("gap_a_minus_b")
        clicked_uni = focus.get("clicked_university")
        dim = focus.get("dimension")
        if gap is not None:
            higher = a_name if gap >= 0 else b_name
            lower = b_name if gap >= 0 else a_name
            gap_sentence = f"**{higher}** is higher than **{lower}** by **{abs(float(gap)):.1f} points** in this dimension."
        else:
            gap_sentence = "The dimension gap could not be calculated from the current context."
        selected_sentence = (
            f"You clicked **{clicked_uni} — {dim}**" + (f" at **{float(clicked):.1f}**." if isinstance(clicked, (int, float)) else ".")
            if clicked_uni else f"You clicked the **{dim}** comparison gap."
        )
        return f"""
**Focused analysis — {label}**

{selected_sentence} In the same year, **{a_name}** has **{format_number(a_val, 1)}** and **{b_name}** has **{format_number(b_val, 1)}** for **{dim.lower()}**. {gap_sentence}

**What this specific difference means**

This mark should be read as a dimension-specific difference, not as a general statement that one university is better overall. Compare it with the overall gap of **{format_number(focus.get('overall_gap_a_minus_b'), 1)} points**: if the {dim.lower()} gap is larger in absolute terms, this dimension is an important contributor to the overall separation; if it is smaller, other dimensions explain more of the overall difference.

**What to inspect next**

Check whether the same {dim.lower()} difference persists over time. A persistent gap is a stronger descriptive pattern than a one-year difference.

**Limit**

This is a focused descriptive comparison of the clicked mark; it does not explain its cause.{q}
"""

    if view == "University Profile":
        return f"""
**Focused analysis — {label}**

You selected **{label} = {format_number(focus.get('value', clicked), 1)}** for **{focus.get('university')}**. The university's overall score is **{format_number(focus.get('overall_score'), 1)}**, so the clicked dimension should be interpreted as one component of the multidimensional profile rather than as a standalone ranking.

**Interpretation of the clicked dimension**

A value above the overall score indicates that this dimension strengthens the profile; a value below it indicates a comparatively weaker component. The profile dispersion is **{format_number(focus.get('profile_dispersion'), 1)}**, which shows how unevenly the four dimensions are distributed.

**What to inspect next**

Open **Why this score?** for the selected dimension to see which normalized indicators contribute to it, then compare the same dimension with the selected benchmark group.

**Limit**

The clicked score is a normalized dashboard profile measure, not a causal estimate.{q}
"""

    if view == "Finance Explorer":
        return f"""
**Focused analysis — {label}**

You selected **{label} = {format_number(focus.get('value', clicked), 2)}** for **{focus.get('university')}**. The financial score is **{format_number(focus.get('financial_score'), 1)}**, while the overall profile score is **{format_number(focus.get('overall_score'), 1)}**.

**Interpretation of this indicator**

This financial variable describes one part of the university's resource or financial structure. It should not be read as a direct cause of the financial or overall score. Its meaning is strongest when compared with the current filtered distribution and with other financial indicators rather than in isolation.

**What to inspect next**

Use this indicator on the Finance scatterplot x-axis, then compare the selected university with the current-filter median. That shows whether the clicked value is typical or distinctive within the active comparison group.

**Limit**

The relationship shown here is descriptive and associative, not causal.{q}
"""

    if view == "Teaching and Research":
        return f"""
**Focused analysis — {label}**

You selected **{label} = {format_number(focus.get('value', clicked), 2)}** for **{focus.get('university')}**. In the same profile, Teaching is **{format_number(focus.get('teaching_score'), 1)}**, Placement is **{format_number(focus.get('placement_score'), 1)}**, and Research is **{format_number(focus.get('research_score'), 1)}**.

**Interpretation of the clicked mark**

The selected value should be interpreted inside its relevant dimension. An individual teaching or research indicator helps explain the composite score, but it does not determine that score by itself. If the clicked mark is one of the three composite scores, its difference from the other two indicates the university's academic orientation in this dashboard view.

**What to inspect next**

Compare the clicked indicator with the other indicators in the same block and then use the Teaching vs Research positioning chart to see whether this profile is common or distinctive.

**Limit**

This is descriptive profile analysis and does not identify causes of the observed difference.{q}
"""

    if view == "Time Dynamics / What Changed":
        return f"""
**Focused analysis — {label}**

You selected the change in **{label}** for **{focus.get('university')}**. It moved from **{format_number(focus.get('start_value'), 1)}** in **{focus.get('start_year')}** to **{format_number(focus.get('end_value'), 1)}** in **{focus.get('end_year')}**, a change of **{format_number(focus.get('change'), 1)} points**.

**Interpretation of this change**

This is the exact temporal movement represented by the clicked bar. It should be compared with changes in the other dimensions to determine whether the university changed broadly or whether the movement was concentrated in this one area.

**Limit**

The chart describes change between observed years; it does not explain what caused that change.{q}
"""

    # Generic focused mode for other interactive charts / Explain-this-visual requests.
    return f"""
**Focused analysis — {label}**

You selected **{label}**{f' = **{clicked}**' if clicked not in (None, 'current visual') else ''}. The Companion is now prioritizing this exact visual mark rather than summarizing the whole **{view}** page.

**Interpretation**

Read this mark in relation to the surrounding observations and the active filters. The selected value is descriptive evidence within the current analytical context; it should not be treated as an isolated quality judgment or causal result.

**Next step**

Compare the selected mark with its closest benchmark or neighboring observations in the same visualization.{q}
"""


def generate_local_interpretation(context: dict, user_question: str | None = None) -> str:
    if context.get("focus_mode") and context.get("interaction_focus"):
        return local_focused_interpretation(context, user_question)
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


def _focused_answer_is_too_generic(answer: str, context: dict) -> bool:
    """Reject meta-only focused responses and fall back to deterministic chart analysis."""
    if not context.get("focus_mode") or not context.get("interaction_focus"):
        return False
    focus = context.get("focus_details") or {}
    text = (answer or "").strip()
    if len(text.split()) < 65:
        return True
    bad_phrases = [
        "companion is now prioritizing",
        "you selected dea efficiency vs selected variable",
        "read this mark in relation to the surrounding observations",
        "compare the selected mark with its closest benchmark",
    ]
    lower = text.lower()
    if any(p in lower for p in bad_phrases):
        return True
    if focus.get("focus_kind") == "visual" or focus.get("mode") == "visual":
        # Rich visual explanations should contain several numbers when the context has them.
        numeric_tokens = re.findall(r"(?<![A-Za-z])[-+]?\d+(?:\.\d+)?", text)
        if len(numeric_tokens) < 2:
            return True
    return False


def local_followup_answer(context: dict, user_question: str) -> str:
    """Provide a useful deterministic answer when the external model is unavailable.

    It does not try to imitate open-ended reasoning; it answers from the current
    dashboard value and explicitly avoids claims the current view cannot support.
    """
    q = (user_question or "").strip()
    view = context.get("view_type")
    if not q:
        return ""

    if context.get("focus_mode") and context.get("focus_details"):
        focus = context.get("focus_details") or {}
        if focus.get("focus_kind") == "visual" or focus.get("mode") == "visual":
            return f"\n\n**Direct answer to your question**\n\nYour question is being answered from the selected visual **{focus.get('label')}**. The chart-specific interpretation above contains the values available for this visual. If the question asks *why* the pattern exists, the current dashboard cannot establish a cause; it can only describe the observed pattern and benchmark position."
        return f"\n\n**Direct answer to your question**\n\nThe current focus is **{focus.get('label')}**. Based on the clicked mark and its linked comparison values, the safest answer is the focused interpretation above. The dashboard can describe the difference and its relative position, but it cannot identify a causal mechanism behind it."

    if view == "Overview":
        dims = {"Teaching": context.get("average_teaching_score"), "Placement": context.get("average_placement_score"), "Research": context.get("average_research_score"), "Financial": context.get("average_financial_score")}
        valid = {k:_safe_float(v) for k,v in dims.items() if _safe_float(v) is not None}
        strong = max(valid, key=valid.get) if valid else "n/a"
        weak = min(valid, key=valid.get) if valid else "n/a"
        return f"\n\n**Direct answer to your question**\n\nFor the current filtered system, the clearest descriptive contrast is **{strong} ({format_number(valid.get(strong),1)})** versus **{weak} ({format_number(valid.get(weak),1)})**. The average Overall score is **{format_number(context.get('average_overall_score'),1)}** across **{context.get('number_of_universities')} universities**. The page supports a system-level descriptive answer, not a causal explanation."

    if view == "University Profile":
        dims = {"Teaching": context.get("teaching_score"), "Placement": context.get("placement_score"), "Research": context.get("research_score"), "Financial": context.get("financial_score")}
        valid = {k:_safe_float(v) for k,v in dims.items() if _safe_float(v) is not None}
        strong = max(valid, key=valid.get) if valid else "n/a"
        weak = min(valid, key=valid.get) if valid else "n/a"
        return f"\n\n**Direct answer to your question**\n\nFor **{context.get('university')}**, the strongest dashboard dimension is **{strong} ({format_number(valid.get(strong),1)})** and the weakest is **{weak} ({format_number(valid.get(weak),1)})**; Overall is **{format_number(context.get('overall_score'),1)}**. This is the main profile evidence available for answering the question from this page."

    if view == "University Comparison":
        gaps = context.get("score_gaps_a_minus_b") or {}
        named = {"Overall":gaps.get("overall_gap_a_minus_b"), "Teaching":gaps.get("teaching_gap_a_minus_b"), "Placement":gaps.get("placement_gap_a_minus_b"), "Research":gaps.get("research_gap_a_minus_b"), "Financial":gaps.get("financial_gap_a_minus_b")}
        valid = {k:_safe_float(v) for k,v in named.items() if _safe_float(v) is not None}
        largest = max(valid, key=lambda k: abs(valid[k])) if valid else "n/a"
        return f"\n\n**Direct answer to your question**\n\nThe largest current difference is in **{largest} ({_fmt_signed(valid.get(largest))} points, A minus B)**. This is the most informative dimension for explaining the visible separation between the two selected profiles; the page does not establish why that difference exists."

    if view == "Finance Explorer":
        return f"\n\n**Direct answer to your question**\n\nThe selected university has **{context.get('finance_x_axis')} = {format_number(context.get('selected_x_value'),2)}** versus a current-filter median of **{format_number(context.get('x_axis_median_current_filter'),2)}**, while **{context.get('score_y_axis')} = {format_number(context.get('selected_y_value'),1)}** versus a median of **{format_number(context.get('y_axis_median_current_filter'),1)}**. This supports a descriptive association only, not a causal answer."

    if view == "DEA Efficiency Explorer":
        return f"\n\n**Direct answer to your question**\n\nThe key DEA evidence is **VRS {format_number(context.get('dea_vrs_efficiency_100'),1)}**, **CRS {format_number(context.get('dea_crs_efficiency_100'),1)}**, and **Scale efficiency {format_number(context.get('dea_scale_efficiency_100'),1)}**. The VRS position is around the **{format_number(context.get('dea_vrs_efficiency_100_percentile_current_filter'),0)}th percentile** in the current filter. These are relative benchmarking results, not causal estimates."

    if view == "Ranking Explorer":
        return f"\n\n**Direct answer to your question**\n\nFor **{context.get('metric')}**, **{context.get('selected_university')}** is **{context.get('selected_rank_in_current_filter')}/{context.get('number_of_universities')}** with value **{format_number(context.get('selected_value'),1)}**, around the **{format_number(context.get('selected_percentile_current_filter'),0)}th percentile**. Any broader judgment would require checking the other dimensions rather than relying on this one ranking metric."

    if view == "Time Dynamics / What Changed":
        changes = context.get("changes_end_minus_start") or {}
        largest = context.get("largest_change_dimension")
        return f"\n\n**Direct answer to your question**\n\nFrom **{context.get('start_year')}** to **{context.get('end_year')}**, the largest recorded movement is **{str(largest).replace('_score','').title() if largest else 'n/a'} ({_fmt_signed(changes.get(largest))} points)**, while Overall changes by **{_fmt_signed(changes.get('overall_score'))} points**. This describes what changed, not why."

    if view == "Teaching and Research":
        vals = {"Teaching":context.get("teaching_score"), "Placement":context.get("placement_score"), "Research":context.get("research_score")}
        valid = {k:_safe_float(v) for k,v in vals.items() if _safe_float(v) is not None}
        strong = max(valid, key=valid.get) if valid else "n/a"
        weak = min(valid, key=valid.get) if valid else "n/a"
        return f"\n\n**Direct answer to your question**\n\nThe clearest academic-profile contrast is **{strong} ({format_number(valid.get(strong),1)})** versus **{weak} ({format_number(valid.get(weak),1)})**. Use the underlying teaching/research indicator bars to identify which observed measures contribute to that profile."

    if view == "Data and Methodology":
        return "\n\n**Direct answer to your question**\n\nThe dashboard uses normalized profile scores for descriptive comparison and a separate DEA layer for relative efficiency benchmarking. It does not support causal inference or official institutional evaluation."

    return "\n\n**Direct answer to your question**\n\nThe current dashboard context does not contain enough structured evidence to answer that question reliably without adding information that is not visible in the selected view."


def generate_ai_interpretation(context: dict, user_question: str | None = None) -> str:
    try:
        api_key = st.secrets.get("OPENAI_API_KEY", None)
    except Exception:
        api_key = None

    evidence_items = context.get("evidence_items", [])

    if OpenAI is None or not is_valid_api_key(api_key):
        local_answer = generate_local_interpretation(context, None)
        if user_question:
            local_answer += local_followup_answer(context, user_question)
        return annotate_inline_evidence_ids(local_answer, evidence_items)

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
            "Write analysis, not a list of visible numbers. Use the quantitative values that are necessary to support the interpretation, but do not dump every available value.",
            "Use computed percentiles, medians, gap labels, benchmark values and cluster-position fields whenever they materially support the interpretation.",
            "The context contains an evidence_items registry. Whenever you state a key quantitative value that exists in evidence_items, put the exact visible value in the prose and append its evidence ID immediately after that value, for example 74.2 [E4]. Use only IDs that exist in evidence_items.",
            "The interface hides the [E#] token and turns the number itself into a clickable link to the corresponding visualization. Therefore the evidence ID must follow the exact numeric value it belongs to.",
            "Do not create a separate Evidence section, Linked evidence section, Key values section, source list, citation list, or appendix of numbers. All useful quantitative values must appear naturally inside the analytical prose where they are interpreted.",
            "If interaction_selection is present, treat the chart-clicked university as the explicit user-selected observation and keep the interpretation grounded in that selected mark.",
            "If focus_mode is true and interaction_focus is present, switch to FOCUSED MODE. Do not give the normal page-wide overview. Use focus_details as the primary evidence source. At least 70% of the answer must interpret the selected mark or visual; broader page context is allowed only when it helps explain it.",
            "If focus_details.focus_kind is 'visual' or focus_details.mode is 'visual', this is an EXPLAIN-THIS-VISUAL request. Do NOT say merely that the user selected a visual, that the Companion is prioritizing it, or that the mark should be read in context. Instead explain the actual visible pattern using the numerical fields in focus_details: identify the main pattern, quantify the most important comparison, explain what that pattern means on this page, suggest one specific next inspection, and state one limitation. Start with 'Visual analysis — <visual label>'.",
            "For EXPLAIN-THIS-VISUAL requests, include at least two concrete values from focus_details whenever two or more numeric values are available. Never invent chart values that are absent from focus_details.",
            "In FOCUSED MODE on University Comparison, compare only the clicked dimension for the two selected universities first, state both values and the exact A-minus-B gap, then explain whether that dimension contributes strongly or weakly to the overall gap. Do not repeat a generic comparison-page summary.",
            "In FOCUSED MODE on Finance Explorer or Teaching and Research, explain the selected indicator and how it relates to the relevant composite profile; do not turn one indicator into a causal explanation.",
            "In FOCUSED MODE on Time Dynamics, explain the selected dimension's start value, end value, and change before discussing anything else.",
            "Keep focused-mode responses concise (roughly 120-180 words) and omit generic sections that are unrelated to the clicked mark.",
            "If evidence_focus is present, explicitly address that linked value because the user clicked its number in the previous Companion explanation.",
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
        model_answer = response.output_text
        if _focused_answer_is_too_generic(model_answer, context):
            model_answer = generate_local_interpretation(context, user_question)
        # Never append a separate list of evidence values. Only annotate values that already
        # occur naturally in the prose so they can become clickable in-place.
        return annotate_inline_evidence_ids(model_answer, evidence_items)
    except Exception:
        local_answer = generate_local_interpretation(context, None)
        if user_question:
            local_answer += local_followup_answer(context, user_question)
        return annotate_inline_evidence_ids(local_answer, evidence_items)


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

st.title("Interactive Visual Analytics Dashboard for Italian Universities")
st.caption("Interactive prototype for exploring multidimensional performance, resources, rankings, and DEA-based efficiency profiles of Italian universities, 2020-2023. The Analysis Companion is a supporting interpretive layer.")

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
    if st.button("Reset analytical state", width="stretch", help="Clear filters, chart selections, focused values and Companion context."):
        reset_analytical_state()
        st.rerun()

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
    st.markdown(f"<div class='breadcrumb'>{html.escape(build_context_breadcrumb(page, year, macro_area, region, size_class, university))}</div>", unsafe_allow_html=True)
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
            request_visual_explanation("Overview", "Top 10 universities by overall score", "visual::overview_top10", "explain_overview_top10")
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
            request_visual_explanation("Overview", "Current filtered average profile", "visual::overview_average_profile", "explain_overview_average_profile")
            st.markdown("#### Average profile by macro-area")
            st.altair_chart(macro_chart, width="stretch")
            request_visual_explanation("Overview", "Average profile by macro-area", "visual::overview_macro_area", "explain_overview_macro")

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

        st.markdown("<div id='source-overview-brush'></div>", unsafe_allow_html=True)
        st.markdown("#### Linked brushing explorer")
        st.caption("Drag a rectangle across the Teaching vs Research scatterplot. The linked profile on the right updates immediately for the brushed group, while the Companion receives the selected ranges and group summary.")
        overview_linked_brush = alt.selection_interval(name="overview_linked_brush", encodings=["x", "y"], clear="dblclick")
        brush_scatter = (
            alt.Chart(filtered)
            .mark_circle(size=95, opacity=0.72)
            .encode(
                x=alt.X("teaching_score:Q", title="Teaching score", scale=alt.Scale(domain=[0, 100])),
                y=alt.Y("research_score:Q", title="Research score", scale=alt.Scale(domain=[0, 100])),
                color=alt.condition(overview_linked_brush, alt.Color("macro_area:N", title="Macro-area"), alt.value("#d4d7dc")),
                tooltip=["university", "region", alt.Tooltip("teaching_score:Q", format=".1f"), alt.Tooltip("research_score:Q", format=".1f")],
            )
            .add_params(overview_linked_brush)
            .properties(height=320)
        )
        linked_profile = (
            alt.Chart(filtered)
            .transform_filter(overview_linked_brush)
            .transform_fold(["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"], as_=["Dimension", "Score"])
            .transform_calculate(
                DimensionLabel="replace(replace(replace(replace(replace(datum.Dimension, 'overall_score', 'Overall'), 'teaching_score', 'Teaching'), 'placement_score', 'Placement'), 'research_score', 'Research'), 'financial_score', 'Financial')"
            )
            .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
            .encode(
                y=alt.Y("DimensionLabel:N", sort=["Overall", "Teaching", "Placement", "Research", "Financial"], title=None),
                x=alt.X("mean(Score):Q", title="Mean score of brushed universities", scale=alt.Scale(domain=[0, 100])),
                tooltip=[alt.Tooltip("mean(Score):Q", title="Mean", format=".1f")],
            )
            .properties(height=320)
        )
        brush_event = st.altair_chart(
            alt.hconcat(brush_scatter, linked_profile).resolve_scale(color="independent"),
            key="overview_linked_brush_chart",
            on_select="rerun",
            selection_mode=["overview_linked_brush"],
            width="stretch",
        )
        request_visual_explanation("Overview", "Linked brushing explorer", "visual::overview_linked_brush", "explain_overview_brush")
        brush_payload = interval_selection_payload(brush_event, "overview_linked_brush")
        brushed_group = filter_by_interval(filtered, brush_payload, "teaching_score", "research_score")
        if not brushed_group.empty:
            active_context["linked_brush"] = {
                "selection": clean_value(brush_payload),
                "selected_universities": int(brushed_group["university"].nunique()),
                "average_overall_score": clean_value(brushed_group["overall_score"].mean()),
                "average_teaching_score": clean_value(brushed_group["teaching_score"].mean()),
                "average_research_score": clean_value(brushed_group["research_score"].mean()),
                "university_names": brushed_group["university"].tolist()[:20],
            }
            st.caption(f"Brushed group: {brushed_group['university'].nunique()} universities · mean overall {brushed_group['overall_score'].mean():.1f}.")

    elif page == "University Profile":
        active_context = selected_context(selected, page)
        active_context = add_profile_position_context(active_context, filtered, selected)
        st.markdown("<div id='source-profile-summary'></div>", unsafe_allow_html=True)
        st.markdown("### University profile")
        st.caption("This page focuses on one selected university and compares its overall profile with national and macro-area averages.")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Overall score", f"{selected['overall_score']:.1f}", help=SCORE_HELP["overall_score"])
        k2.metric("Rank", f"{int(selected['overall_rank_year'])}/61", help="Rank derived from the dashboard overall score within the same year; not an official external ranking.")
        k3.metric("Teaching", f"{selected['teaching_score']:.1f}", help=SCORE_HELP["teaching_score"])
        k4.metric("Research", f"{selected['research_score']:.1f}", help=SCORE_HELP["research_score"])
        k5.metric("Financial", f"{selected['financial_score']:.1f}", help=SCORE_HELP["financial_score"])
        dispersion = profile_dispersion(selected)
        active_context["profile_dispersion"] = clean_value(dispersion)
        active_context["profile_balance_label"] = profile_balance_label(dispersion)
        b1, b2 = st.columns(2)
        b1.metric("Profile dispersion (SD)", f"{dispersion:.1f}", help="Standard deviation across Teaching, Placement, Research and Financial scores. Lower values indicate a more balanced profile; this is descriptive and not a quality score.")
        b2.metric("Profile balance", profile_balance_label(dispersion), help="Descriptive classification based on dispersion across the four dashboard dimensions.")

        st.markdown("#### Rule-based insight flags")
        flags = insight_flags(selected, filtered, df)
        active_context["rule_based_flags"] = flags
        render_insight_flags(flags)

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
            request_visual_explanation("University Profile", "Performance profile", "visual::profile_dimensions", "explain_profile_dimensions")
        with c2:
            st.markdown("<div id='source-profile-benchmark'></div>", unsafe_allow_html=True)
            st.markdown("#### Dynamic benchmark comparison")
            benchmark_mode = st.selectbox(
                "Compare with",
                ["National", "Macro-area", "Region", "Same size", "Current filtered group", "Custom peer group"],
                key="profile_benchmark_mode",
            )
            custom_peers: list[str] = []
            if benchmark_mode == "Custom peer group":
                peer_options = [u for u in sorted(df[df["year"].eq(year)]["university"].unique()) if u != university]
                custom_peers = st.multiselect("Custom peers", peer_options, default=peer_options[: min(5, len(peer_options))], key="profile_custom_peers")
            bench_group, bench_label = benchmark_group(df, selected, benchmark_mode, custom_peers, filtered)
            benchmark_df = benchmark_profile_data(selected, bench_group, bench_label)
            active_context["dynamic_benchmark"] = {
                "mode": benchmark_mode,
                "label": bench_label,
                "n": int(bench_group["university"].nunique()),
                "overall_average": clean_value(bench_group["overall_score"].mean()),
                "teaching_average": clean_value(bench_group["teaching_score"].mean()),
                "placement_average": clean_value(bench_group["placement_score"].mean()),
                "research_average": clean_value(bench_group["research_score"].mean()),
                "financial_average": clean_value(bench_group["financial_score"].mean()),
            }
            benchmark_profile_chart = (
                alt.Chart(benchmark_df)
                .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                .encode(
                    x=alt.X("Dimension:N", sort=["Overall", "Teaching", "Placement", "Research", "Financial"], title="Dimension"),
                    y=alt.Y("Score:Q", scale=alt.Scale(domain=[0, 100])),
                    xOffset="Series:N",
                    color=alt.Color("Series:N", title="Comparison"),
                    tooltip=["Series", "Dimension", alt.Tooltip("Score:Q", format=".1f")],
                )
                .properties(height=300)
            )
            st.altair_chart(benchmark_profile_chart, width="stretch")
            st.caption(f"Benchmark contains {bench_group['university'].nunique()} universities. Comparisons remain descriptive and depend on the selected peer definition.")
            request_visual_explanation("University Profile", "Dynamic benchmark comparison", "visual::profile_benchmark", "explain_profile_benchmark")

        st.markdown("#### Why this score?")
        score_to_explain = st.selectbox("Score breakdown", ["Teaching", "Placement", "Research", "Financial"], key="score_breakdown_dimension")
        breakdown = score_component_breakdown(df, selected, score_to_explain)
        score_field = score_to_explain.lower() + "_score"
        reconstructed = pd.to_numeric(breakdown["Component score"], errors="coerce").mean()
        active_context["score_breakdown"] = {
            "dimension": score_to_explain,
            "reported_score": clean_value(selected.get(score_field)),
            "reconstructed_score": clean_value(reconstructed),
            "components": breakdown.to_dict("records"),
        }
        bd1, bd2 = st.columns([1.4, 1])
        with bd1:
            component_chart = (
                alt.Chart(breakdown.dropna(subset=["Component score"]))
                .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    y=alt.Y("Component:N", sort="-x", title=None),
                    x=alt.X("Component score:Q", title="Component score (0–100)", scale=alt.Scale(domain=[0, 100])),
                    tooltip=["Component", "Field", alt.Tooltip("Raw value:Q", format=",.2f"), alt.Tooltip("Component score:Q", format=".1f"), "Method"],
                )
                .properties(height=max(240, 42 * max(1, len(breakdown))))
            )
            st.altair_chart(component_chart, width="stretch")
        with bd2:
            st.metric(f"Reported {score_to_explain} score", f"{selected.get(score_field):.1f}")
            st.metric("Reconstructed from components", f"{reconstructed:.1f}")
            st.dataframe(breakdown, width="stretch", hide_index=True)
        st.markdown("<div class='method-note'>The breakdown reconstructs the exact dashboard score logic from the workbook methodology: direct percentage-like components are used as-is, while research and selected financial components are normalized within each year. Personnel cost share is reversed before aggregation.</div>", unsafe_allow_html=True)
        request_visual_explanation("University Profile", f"{score_to_explain} score breakdown", f"visual::{score_field}_breakdown", "explain_score_breakdown")

        st.markdown("#### Weight sensitivity / robustness view")
        year_universities = df[df["year"].eq(year)].copy()
        sens_df = sensitivity_table(year_universities, university)
        active_context["weight_sensitivity"] = sens_df.to_dict("records")
        sensitivity_chart = (
            alt.Chart(sens_df)
            .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
            .encode(
                y=alt.Y("Scenario:N", sort=None, title=None),
                x=alt.X("Score:Q", title="Weighted profile score", scale=alt.Scale(domain=[0, 100])),
                color=alt.Color("Rank:Q", title="Rank", scale=alt.Scale(reverse=True)),
                tooltip=["Scenario", alt.Tooltip("Score:Q", format=".1f"), "Rank", "Universities", "Weights"],
            )
            .properties(height=280)
        )
        st.altair_chart(sensitivity_chart, width="stretch")
        st.dataframe(sens_df, width="stretch", hide_index=True)
        st.caption("Sensitivity scenarios are exploratory. They show how the profile score and rank would change under predefined weighting priorities; they do not replace the equal-weight dashboard score.")
        request_visual_explanation("University Profile", "Weight sensitivity scenarios", "visual::weight_sensitivity", "explain_weight_sensitivity")

        trend = df[df["university"] == university].sort_values("year")
        if not trend.empty:
            _trend_fields = ["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"]
            _trend_start = trend.iloc[0]
            _trend_end = trend.iloc[-1]
            _trend_changes = {f: clean_value(_trend_end.get(f) - _trend_start.get(f)) for f in _trend_fields if f in trend.columns}
            _largest_trend = max(_trend_changes, key=lambda f: abs(float(_trend_changes.get(f) or 0))) if _trend_changes else None
            active_context["profile_trend_summary"] = {
                "start_year": int(_trend_start.get("year")),
                "end_year": int(_trend_end.get("year")),
                "start_values": {f: clean_value(_trend_start.get(f)) for f in _trend_fields if f in trend.columns},
                "end_values": {f: clean_value(_trend_end.get(f)) for f in _trend_fields if f in trend.columns},
                "changes": _trend_changes,
                "largest_change_dimension": _largest_trend,
            }
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
        request_visual_explanation("University Profile", "Score dynamics over time", "visual::profile_trend", "explain_profile_trend")

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
                request_visual_explanation("University Comparison", "Side-by-side score profile", "visual::comparison_scores", "explain_comparison_scores")
            with gc2:
                st.markdown("<div id='source-comparison-gaps'></div>", unsafe_allow_html=True)
                st.markdown("#### Score gaps")
                st.caption("Click a gap bar to focus the Companion on that difference.")
                st.altair_chart(gap_chart, key="comparison_gap_chart", on_select=on_comparison_gap_select, selection_mode=["comparison_gap_selection"], width="stretch")
                request_visual_explanation("University Comparison", "Score gaps", "visual::comparison_gaps", "explain_comparison_gaps")

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

            st.markdown("#### Selected university vs benchmark group")
            st.caption("This additional mode compares the active sidebar university with a contextual benchmark instead of requiring a second university.")
            benchmark_mode_cmp = st.selectbox(
                "Benchmark group for selected university",
                ["National", "Macro-area", "Region", "Same size", "Current filtered group", "Custom peer group"],
                key="comparison_benchmark_mode",
            )
            cmp_custom: list[str] = []
            if benchmark_mode_cmp == "Custom peer group":
                cmp_peer_options = [u for u in sorted(df[df["year"].eq(year)]["university"].unique()) if u != university]
                cmp_custom = st.multiselect("Custom benchmark peers", cmp_peer_options, default=cmp_peer_options[: min(5, len(cmp_peer_options))], key="comparison_custom_peers")
            cmp_group, cmp_label = benchmark_group(df, selected, benchmark_mode_cmp, cmp_custom, filtered)
            cmp_benchmark_df = benchmark_profile_data(selected, cmp_group, cmp_label)
            active_context["selected_vs_benchmark"] = {
                "selected_university": university,
                "benchmark": cmp_label,
                "benchmark_n": int(cmp_group["university"].nunique()),
                "dimension_gaps_selected_minus_benchmark": {
                    field: clean_value(float(selected[field]) - float(cmp_group[field].mean()))
                    for field in ["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"]
                },
            }
            cmp_bench_chart = (
                alt.Chart(cmp_benchmark_df)
                .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                .encode(
                    x=alt.X("Dimension:N", sort=["Overall", "Teaching", "Placement", "Research", "Financial"]),
                    y=alt.Y("Score:Q", scale=alt.Scale(domain=[0,100])),
                    xOffset="Series:N",
                    color=alt.Color("Series:N", title="Series"),
                    tooltip=["Series", "Dimension", alt.Tooltip("Score:Q", format=".1f")],
                )
                .properties(height=300)
            )
            st.altair_chart(cmp_bench_chart, width="stretch")
            request_visual_explanation("University Comparison", "Selected university vs benchmark group", "visual::comparison_benchmark", "explain_comparison_benchmark")

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
        request_visual_explanation("Finance Explorer", "Financial indicator vs selected score", "visual::finance_scatter", "explain_finance_scatter")
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
        request_visual_explanation("Finance Explorer", "Selected university financial structure", "visual::finance_structure", "explain_finance_structure")
        st.markdown("<div class='method-note'>Financial scatterplots show observed associations only. Moving along the x-axis does not represent a causal intervention or what-if simulation.</div>", unsafe_allow_html=True)


    elif page == "DEA Efficiency Explorer":
        active_context = make_dea_context(filtered, selected)
        st.markdown("<div id='source-dea-summary'></div>", unsafe_allow_html=True)
        st.markdown("### DEA efficiency explorer")
        st.caption("This page adds an exploratory DEA efficiency layer. DEA-VRS is the main score because universities differ in size and scale.")

        if "dea_vrs_efficiency_100" not in df.columns:
            st.error("DEA columns are not available in the current dataset. Upload university_dashboard_with_dea_efficiency.xlsx and use the Dashboard_Data_with_DEA sheet.")
        else:
            d1, d2, d3, d4, d5 = st.columns(5)
            d1.metric("DEA-VRS efficiency", format_number(selected.get("dea_vrs_efficiency_100"), 1), help=SCORE_HELP["dea_vrs_efficiency_100"])
            d2.metric("DEA-VRS rank", f"{int(selected.get('dea_vrs_rank_year'))}/61", help="Rank of the VRS DEA score within the same year. Ties are possible, especially at the efficient frontier.")
            d3.metric("DEA-CRS efficiency", format_number(selected.get("dea_crs_efficiency_100"), 1), help=SCORE_HELP["dea_crs_efficiency_100"])
            d4.metric("Scale efficiency", format_number(selected.get("dea_scale_efficiency_100"), 1), help=SCORE_HELP["dea_scale_efficiency_100"])
            d5.metric("Category", str(selected.get("efficiency_category")), help="Descriptive efficiency category stored in the dashboard-ready dataset.")

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
            _dea_x_series = numeric_series(filtered, dea_x_col)
            active_context["dea_scatter_x_median_current_filter"] = clean_value(_dea_x_series.median()) if not _dea_x_series.empty else None
            active_context["dea_scatter_x_percentile_current_filter"] = clean_value(percentile_position(filtered, dea_x_col, selected.get(dea_x_col)))

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
            request_visual_explanation("DEA Efficiency Explorer", "DEA efficiency vs selected variable", "visual::dea_scatter", "explain_dea_scatter")
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
                if not dea_trend.empty:
                    _dea_fields = ["dea_vrs_efficiency_100", "dea_crs_efficiency_100", "dea_scale_efficiency_100"]
                    _dea_start = dea_trend.iloc[0]
                    _dea_end = dea_trend.iloc[-1]
                    active_context["dea_trend_summary"] = {
                        "start_year": int(_dea_start.get("year")),
                        "end_year": int(_dea_end.get("year")),
                        "start_values": {f: clean_value(_dea_start.get(f)) for f in _dea_fields if f in dea_trend.columns},
                        "end_values": {f: clean_value(_dea_end.get(f)) for f in _dea_fields if f in dea_trend.columns},
                        "changes": {f: clean_value(_dea_end.get(f) - _dea_start.get(f)) for f in _dea_fields if f in dea_trend.columns},
                    }
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
                request_visual_explanation("DEA Efficiency Explorer", "DEA trend 2020–2023", "visual::dea_trend", "explain_dea_trend")
            with dc2:
                st.markdown("#### Top 10 by DEA-VRS efficiency")
                top_dea = filtered.sort_values("dea_vrs_efficiency_100", ascending=False).head(10)
                active_context["dea_top10"] = top_dea[["university", "dea_vrs_efficiency_100", "efficiency_category"]].to_dict("records")
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
                request_visual_explanation("DEA Efficiency Explorer", "Top 10 by DEA-VRS efficiency", "visual::dea_top10", "explain_dea_top10")

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
        request_visual_explanation("Ranking Explorer", f"Ranking by {ranking_metric_label}", "visual::ranking_chart", "explain_ranking_chart")

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
                request_visual_explanation("Time Dynamics / What Changed", "Change by dimension", "visual::time_change", "explain_time_change")

                slope_rows = []
                for _, r in change_df.iterrows():
                    slope_rows.append({"Dimension": r["Dimension"], "Year": str(start_year), "Score": r["Start"]})
                    slope_rows.append({"Dimension": r["Dimension"], "Year": str(end_year), "Score": r["End"]})
                slope_df = pd.DataFrame(slope_rows)
                slope_chart = (
                    alt.Chart(slope_df)
                    .mark_line(point=True)
                    .encode(
                        x=alt.X("Year:N", sort=[str(start_year), str(end_year)], title=None),
                        y=alt.Y("Score:Q", scale=alt.Scale(domain=[0, 100])),
                        color=alt.Color("Dimension:N", title="Dimension"),
                        detail="Dimension:N",
                        tooltip=["Dimension", "Year", alt.Tooltip("Score:Q", format=".1f")],
                    )
                    .properties(height=300)
                )
                st.markdown("#### Start-to-end slope view")
                st.altair_chart(slope_chart, width="stretch")
                request_visual_explanation("Time Dynamics / What Changed", "Start-to-end slope view", "visual::time_slope", "explain_time_slope")

                largest_change_df = change_df.assign(Absolute_change=change_df["Change"].abs()).sort_values("Absolute_change", ascending=False)
                active_context["largest_dimension_changes"] = largest_change_df.head(3)[["Dimension", "Start", "End", "Change"]].to_dict("records")
                if st.toggle("Show largest changes first", value=False, key="show_largest_dimension_changes"):
                    st.dataframe(largest_change_df[["Dimension", "Start", "End", "Change"]].head(3), width="stretch", hide_index=True)

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
        request_visual_explanation("Teaching and Research", "Teaching, placement and research scores", "visual::teaching_research_scores", "explain_teaching_research_scores")

        tr1, tr2 = st.columns(2)
        with tr1:
            st.markdown("<div id='source-teaching-indicators'></div>", unsafe_allow_html=True)
            st.markdown("#### Teaching and placement indicators")
            teaching_indicator_selection = alt.selection_point(name="teaching_indicator_selection", fields=["Field", "Indicator", "Value"], on="click", clear="dblclick")
            st.caption("Click an indicator bar to focus the Companion on it.")
            st.altair_chart(make_indicator_bar(selected, teaching_indicators, tr_focus).add_params(teaching_indicator_selection), key="teaching_indicator_chart", on_select=on_teaching_indicator_select, selection_mode=["teaching_indicator_selection"], width="stretch")
            request_visual_explanation("Teaching and Research", "Teaching and placement indicators", "visual::teaching_indicators", "explain_teaching_indicators")
        with tr2:
            st.markdown("<div id='source-research-indicators'></div>", unsafe_allow_html=True)
            st.markdown("#### Research indicators")
            research_indicator_selection = alt.selection_point(name="research_indicator_selection", fields=["Field", "Indicator", "Value"], on="click", clear="dblclick")
            st.caption("Click an indicator bar to focus the Companion on it.")
            st.altair_chart(make_indicator_bar(selected, research_indicators, tr_focus).add_params(research_indicator_selection), key="research_indicator_chart", on_select=on_research_indicator_select, selection_mode=["research_indicator_selection"], width="stretch")
            request_visual_explanation("Teaching and Research", "Research indicators", "visual::research_indicators", "explain_research_indicators")

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
        request_visual_explanation("Teaching and Research", "Teaching vs research positioning", "visual::teaching_research_scatter", "explain_teaching_research_scatter")

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
    active_context["focus_mode"] = True
    focus_details = build_focus_details(active_context)
    if focus_details:
        active_context["focus_details"] = focus_details

# Build the evidence registry after all page-specific controls have defined the active context.
# This makes the dashboard -> Companion link explicit and keeps the evidence synchronized
# with year, university, filters, selected axes, ranking metric, and comparison choices.
# The logic version intentionally changes the automatic-analysis signature after Companion
# upgrades so an old in-session cached response cannot survive a code update.
active_context["companion_logic_version"] = COMPANION_LOGIC_VERSION
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
        "Changing the university, year, filters, axes, ranking metric, comparison, brush selection, or analytical focus automatically changes the data sent to the Companion.</div>",
        unsafe_allow_html=True,
    )

    if "ai_answer" not in st.session_state:
        st.session_state.ai_answer = ""
    if "last_ai_signature" not in st.session_state:
        st.session_state.last_ai_signature = ""
    if "ai_auto_cache" not in st.session_state:
        st.session_state.ai_auto_cache = {}

    analysis_context = dict(active_context)
    analysis_context.pop("evidence_focus", None)
    signature = json.dumps(analysis_context, ensure_ascii=True, sort_keys=True, default=str)
    context_changed = signature != st.session_state.last_ai_signature
    if context_changed:
        st.session_state.ai_answer = ""
        st.session_state.last_ai_signature = signature

    view_type = active_context.get("view_type", page)
    if view_type == "University Comparison" and "university_a" in active_context and "university_b" in active_context:
        context_text = f"{active_context['university_a']['university']} vs {active_context['university_b']['university']}, {active_context.get('year')}"
    elif view_type == "Overview":
        context_text = f"{active_context.get('number_of_universities')} universities, {active_context.get('year')}"
    elif view_type == "Ranking Explorer":
        context_text = f"ranking by {active_context.get('metric')}, {active_context.get('year')}"
    elif view_type == "Time Dynamics / What Changed":
        context_text = f"{active_context.get('university')}, {active_context.get('start_year')} to {active_context.get('end_year')}"
    elif view_type == "Data and Methodology":
        context_text = "dataset, scores, and DEA methodology"
    else:
        context_text = f"{active_context.get('university')}, {active_context.get('year')}"
    st.markdown(f"**Current page:** {view_type}<br>**Context:** {context_text}", unsafe_allow_html=True)

    interaction_selection = active_context.get("interaction_selection")
    if interaction_selection:
        st.caption(
            f"Interactive selection source: {interaction_selection.get('source_page')} → "
            f"{interaction_selection.get('university')}. The Companion is analyzing the chart-selected observation."
        )
    interaction_focus = active_context.get("interaction_focus")
    if interaction_focus:
        if interaction_focus.get("focus_kind") == "visual" or interaction_focus.get("interaction") == "explain visual request":
            st.info(
                f"Visual focus: {interaction_focus.get('label') or interaction_focus.get('field')}. "
                "The Companion will explain the actual pattern in this chart using the current dashboard values."
            )
        else:
            st.info(
                f"Focused mode: {interaction_focus.get('label') or interaction_focus.get('field')} "
                f"({interaction_focus.get('value')}). The Companion will analyze this clicked mark instead of summarizing the whole page."
            )
    if active_context.get("evidence_focus"):
        ef = active_context["evidence_focus"]
        st.caption(f"Focused value: {ef.get('label')} = {ef.get('display_value')}. The linked visual mark is highlighted in black.")
    if active_context.get("linked_brush"):
        brush_info = active_context["linked_brush"]
        st.caption(f"Linked brush: {brush_info.get('selected_universities')} universities selected in the Overview visual analytics explorer.")

    auto_analysis = st.toggle(
        "Automatic analysis",
        value=True,
        help="Automatically refresh the Companion when the analytical state changes.",
        key="auto_analysis_enabled",
    )

    if auto_analysis and (context_changed or not st.session_state.ai_answer):
        cached_answer = st.session_state.ai_auto_cache.get(signature)
        if cached_answer:
            st.session_state.ai_answer = cached_answer
        else:
            with st.spinner("Updating interpretation for the current dashboard selection..."):
                auto_answer = cached_auto_interpretation(signature)
            st.session_state.ai_answer = auto_answer
            st.session_state.ai_auto_cache[signature] = auto_answer
            if len(st.session_state.ai_auto_cache) > 24:
                oldest_key = next(iter(st.session_state.ai_auto_cache))
                st.session_state.ai_auto_cache.pop(oldest_key, None)

    summary_tab, ask_tab = st.tabs(["Summary", "Ask"])

    with summary_tab:
        if st.session_state.ai_answer:
            # Internal source IDs remain in the stored response only to map each number
            # back to its dashboard location. The user sees clean prose with clickable values.
            linked_answer = linkify_evidence_values(st.session_state.ai_answer, active_context.get("evidence_items", []))
            linked_answer = hide_internal_source_ids(linked_answer)
            st.markdown(linked_answer, unsafe_allow_html=True)
        else:
            st.caption("No interpretation is available yet. Enable Automatic analysis or use the Ask tab.")

    with ask_tab:
        question = st.text_area(
            "Ask about the current view",
            placeholder="Example: Why is Research stronger than Financial in this profile?",
            height=120,
            key="companion_follow_up_question",
        )
        if question.strip() and st.button("Ask follow-up", type="secondary", key="ask_followup_button"):
            with st.spinner("Answering the follow-up in the current dashboard context..."):
                st.session_state.ai_answer = generate_ai_interpretation(active_context, question.strip())
            st.rerun()

        # Keep the structured Companion input available on demand without adding a separate tab.
        with st.expander("Current Companion input"):
            st.json(active_context)

    snapshot = {
        "dashboard_state": active_context,
        "companion_interpretation": st.session_state.ai_answer,
        "export_note": "Exploratory dashboard snapshot. Scores and DEA measures are descriptive/benchmarking outputs, not causal estimates.",
    }
    snapshot_json = json.dumps(snapshot, ensure_ascii=False, indent=2, default=str)
    st.download_button(
        "Export current analysis",
        data=snapshot_json.encode("utf-8"),
        file_name=f"university_dashboard_analysis_{str(university).replace(' ', '_')}_{year}_{page.replace(' ', '_').replace('/', '-')}.json",
        mime="application/json",
        width="stretch",
        help="Download the current filters, selected university, page-specific context and Companion interpretation.",
    )

    st.info(
        "The AI assistant supports interpretation only. It analyzes the current page context and does not provide causal conclusions or autonomous policy recommendations."
    )
