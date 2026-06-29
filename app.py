from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import altair as alt
import pandas as pd
import streamlit as st

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

DATA_PATH = Path("data/university_dashboard_ready_dataset_v2_checked.xlsx")
SHEET_NAME = "Dashboard_Ready_Data_v2"

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


def make_score_profile(row: pd.Series) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Dimension": ["Teaching", "Placement", "Research", "Financial"],
            "Score": [
                row["teaching_score"],
                row["placement_score"],
                row["research_score"],
                row["financial_score"],
            ],
        }
    )


def make_indicator_bar(row: pd.Series, indicators: dict[str, str]) -> alt.Chart:
    data = []
    for col, label in indicators.items():
        if col in row.index and pd.notna(row[col]):
            data.append({"Indicator": label, "Value": float(row[col])})
    chart_df = pd.DataFrame(data)
    if chart_df.empty:
        chart_df = pd.DataFrame({"Indicator": ["No data"], "Value": [0]})
    return (
        alt.Chart(chart_df)
        .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
        .encode(
            y=alt.Y("Indicator:N", sort="-x", title=None),
            x=alt.X("Value:Q", title="Value"),
            tooltip=["Indicator", alt.Tooltip("Value:Q", format=",.2f")],
        )
        .properties(height=max(260, 38 * len(chart_df)))
    )


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
    if view_type == "Data and Methodology":
        return local_methodology_interpretation(context, user_question)
    return local_single_interpretation(context, user_question)


def generate_ai_interpretation(context: dict, user_question: str | None = None) -> str:
    try:
        api_key = st.secrets.get("OPENAI_API_KEY", None)
    except Exception:
        api_key = None

    if OpenAI is None or not is_valid_api_key(api_key):
        return generate_local_interpretation(context, user_question)

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
            "For Finance Explorer, focus on financial indicators and scatterplot axes. For Teaching and Research, focus on teaching/research indicators. For Overview, focus on filtered system-level patterns. For University Comparison, focus on the two selected universities.",
            "Write analysis, not a list of visible numbers. Use a few key numbers only when they support an interpretation.",
            "Use computed percentiles, medians, gap labels, and cluster-position fields whenever they are available.",
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
        return response.output_text
    except Exception:
        return generate_local_interpretation(context, user_question)


def score_bar_chart(score_df: pd.DataFrame) -> alt.Chart:
    return (
        alt.Chart(score_df)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("Dimension:N", sort=None, title="Dimension"),
            y=alt.Y("Score:Q", title="Score", scale=alt.Scale(domain=[0, 100])),
            color=alt.Color("Dimension:N", legend=None),
            tooltip=["Dimension", alt.Tooltip("Score:Q", format=".1f")],
        )
        .properties(height=300)
    )


def benchmark_chart(row: pd.Series) -> alt.Chart:
    benchmark_df = pd.DataFrame(
        {
            "Benchmark": ["Selected university", "National average", "Macro-area average"],
            "Overall score": [
                row["overall_score"],
                row["national_avg_overall_score"],
                row["macro_area_avg_overall_score"],
            ],
        }
    )
    return (
        alt.Chart(benchmark_df)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("Benchmark:N", sort=None, title="Benchmark"),
            y=alt.Y("Overall score:Q", title="Overall score", scale=alt.Scale(domain=[0, 100])),
            color=alt.Color("Benchmark:N", legend=None),
            tooltip=["Benchmark", alt.Tooltip("Overall score:Q", format=".1f")],
        )
        .properties(height=300)
    )


df = load_data()

st.title("AI-enhanced Visual Analytics Dashboard for Italian Universities")
st.caption("Interactive prototype for exploring teaching, placement, research, and financial profiles of Italian universities, 2020-2023.")

with st.sidebar:
    st.header("Filters")
    years = sorted(df["year"].unique())
    year = st.selectbox("Year", years, index=len(years) - 1)

    macro_area_options = ["All"] + sorted(df["macro_area"].dropna().unique())
    macro_area = st.selectbox("Macro-area", macro_area_options)

    region_base = df[df["year"] == year].copy()
    if macro_area != "All":
        region_base = region_base[region_base["macro_area"] == macro_area]
    region_options = ["All"] + sorted(region_base["region"].dropna().unique())
    region = st.selectbox("Region", region_options)

    size_base = region_base.copy()
    if region != "All":
        size_base = size_base[size_base["region"] == region]
    size_options = ["All"] + sorted(size_base["size_class"].dropna().unique())
    size_class = st.selectbox("Size class", size_options)

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

with st.sidebar:
    university = st.selectbox("University", sorted(filtered["university"].unique()))
    st.caption("Tip: choose the page first, then press Analyze current view.")

selected = filtered[filtered["university"] == university].iloc[0]

main_col, ai_col = st.columns([3.2, 1.1], gap="large")

pages = [
    "Overview",
    "University Profile",
    "University Comparison",
    "Finance Explorer",
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
    page = st.radio("Dashboard page", pages, horizontal=True, label_visibility="collapsed")
    st.divider()

    # Default active context, overwritten by page-specific contexts below.
    active_context = selected_context(selected, page)

    if page == "Overview":
        active_context = make_overview_context(filtered, year, macro_area, region, size_class)
        st.markdown("### System overview")
        st.caption("This page summarizes the universities that match the current filters and shows the overall distribution of dashboard-based profile scores.")
        o1, o2, o3, o4 = st.columns(4)
        o1.metric("Universities shown", f"{filtered['university'].nunique()}")
        o2.metric("Average overall", f"{filtered['overall_score'].mean():.1f}")
        o3.metric("Average teaching", f"{filtered['teaching_score'].mean():.1f}")
        o4.metric("Average research", f"{filtered['research_score'].mean():.1f}")

        top10 = filtered.sort_values("overall_score", ascending=False).head(10)
        top_chart = (
            alt.Chart(top10)
            .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
            .encode(
                x=alt.X("overall_score:Q", title="Overall score", scale=alt.Scale(domain=[0, 100])),
                y=alt.Y("university:N", sort="-x", title=None),
                tooltip=["university", "region", alt.Tooltip("overall_score:Q", format=".1f")],
            )
            .properties(height=360)
        )
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
            st.altair_chart(top_chart, width="stretch")
        with c2:
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
        st.markdown("### University profile")
        st.caption("This page focuses on one selected university and compares its overall profile with national and macro-area averages.")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Overall score", f"{selected['overall_score']:.1f}")
        k2.metric("Rank", f"{int(selected['overall_rank_year'])}/61")
        k3.metric("Teaching", f"{selected['teaching_score']:.1f}")
        k4.metric("Research", f"{selected['research_score']:.1f}")
        k5.metric("Financial", f"{selected['financial_score']:.1f}")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Performance profile")
            st.altair_chart(score_bar_chart(make_score_profile(selected)), width="stretch")
        with c2:
            st.markdown("#### Benchmark comparison")
            st.altair_chart(benchmark_chart(selected), width="stretch")

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
                comp_scores.append({"Dimension": label, "University": university_a, "Score": float(row_a[col])})
                comp_scores.append({"Dimension": label, "University": university_b, "Score": float(row_b[col])})
            comp_scores_df = pd.DataFrame(comp_scores)
            comp_chart = (
                alt.Chart(comp_scores_df)
                .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                .encode(
                    x=alt.X("Dimension:N", sort=None, title="Dimension"),
                    y=alt.Y("Score:Q", scale=alt.Scale(domain=[0, 100])),
                    xOffset="University:N",
                    color=alt.Color("University:N", title="University"),
                    tooltip=["University", "Dimension", alt.Tooltip("Score:Q", format=".1f")],
                )
                .properties(height=340)
            )

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
                        "Gap": gap,
                        "Direction": f"{university_a} higher" if gap >= 0 else f"{university_b} higher",
                    }
                )
            gap_df = pd.DataFrame(gap_data)
            gap_chart = (
                alt.Chart(gap_df)
                .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
                .encode(
                    x=alt.X("Gap:Q", title=f"Score gap: {university_a} minus {university_b}"),
                    y=alt.Y("Dimension:N", sort=None, title="Dimension"),
                    color=alt.Color("Direction:N", title="Direction"),
                    tooltip=["Dimension", alt.Tooltip("Gap:Q", format=".1f"), "Direction"],
                )
                .properties(height=340)
            )
            gc1, gc2 = st.columns(2)
            with gc1:
                st.markdown("#### Side-by-side score profile")
                st.altair_chart(comp_chart, width="stretch")
            with gc2:
                st.markdown("#### Score gaps")
                st.altair_chart(gap_chart, width="stretch")

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
            .interactive()
        )
        selected_point = (
            alt.Chart(scatter_data[scatter_data["selected_flag"]])
            .mark_circle(size=260, fillOpacity=0, stroke="black", strokeWidth=3)
            .encode(x=alt.X(f"{x_col}:Q"), y=alt.Y(f"{y_col}:Q"), tooltip=["university"])
        )
        selected_label = (
            alt.Chart(scatter_data[scatter_data["selected_flag"]])
            .mark_text(align="left", dx=10, dy=-10, fontSize=12, fontWeight="bold")
            .encode(x=alt.X(f"{x_col}:Q"), y=alt.Y(f"{y_col}:Q"), text="university:N")
        )
        st.markdown("#### Financial indicator vs selected score")
        st.altair_chart((base_scatter + selected_point + selected_label).properties(height=380), width="stretch")

        finance_indicators = {
            "personnel_cost_share": "Personnel cost share",
            "public_revenue_share": "Public revenue share",
            "student_contribution_share": "Student contribution share",
            "performance_quota_share": "Performance quota share",
            "economic_financial_sustainability_index": "Economic-financial sustainability index",
        }
        st.markdown("#### Selected university financial structure")
        st.altair_chart(make_indicator_bar(selected, finance_indicators), width="stretch")

    elif page == "Teaching and Research":
        active_context = selected_context(selected, page)
        active_context = add_teaching_research_position_context(active_context, filtered, selected)
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
        tr1, tr2 = st.columns(2)
        with tr1:
            st.markdown("#### Teaching and placement indicators")
            st.altair_chart(make_indicator_bar(selected, teaching_indicators), width="stretch")
        with tr2:
            st.markdown("#### Research indicators")
            st.altair_chart(make_indicator_bar(selected, research_indicators), width="stretch")

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
            .interactive()
        )
        selected_tr = (
            alt.Chart(filtered[filtered["university"] == university])
            .mark_circle(size=260, fillOpacity=0, stroke="black", strokeWidth=3)
            .encode(x="teaching_score:Q", y="research_score:Q")
        )
        st.markdown("#### Teaching vs research positioning")
        st.altair_chart((teaching_research_scatter + selected_tr).properties(height=360), width="stretch")

    elif page == "Data and Methodology":
        active_context = {
            "view_type": "Data and Methodology",
            "year": int(year),
            "dataset_scope": "61 Italian universities, 2020-2023",
            "score_definition": "Dashboard-based normalized profile scores, not DEA or SFA efficiency scores.",
            "current_filter_count": int(filtered["university"].nunique()),
            "macro_area_filter": macro_area,
            "region_filter": region,
            "size_class_filter": size_class,
        }
        st.markdown("### Data and methodology")
        st.markdown(
            "The dashboard uses a curated prototype dataset covering 61 Italian universities over 2020-2023. "
            "The score fields are dashboard-based normalized profile scores, not DEA or SFA efficiency scores. "
            "They are intended for visual exploration and comparison."
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

with ai_col:
    st.subheader("AI Analysis Companion")
    st.caption("The assistant analyzes only the currently selected dashboard page.")

    question = st.text_area(
        "Optional question",
        placeholder="Example: What is the main gap on this page?",
        height=110,
    )

    if "ai_answer" not in st.session_state:
        st.session_state.ai_answer = ""
    if "last_ai_signature" not in st.session_state:
        st.session_state.last_ai_signature = ""

    signature = json.dumps(active_context, ensure_ascii=True, sort_keys=True, default=str)
    if signature != st.session_state.last_ai_signature:
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
    elif view_type == "Data and Methodology":
        st.markdown(f"**Current page:** {view_type}<br>**Context:** dataset and score construction", unsafe_allow_html=True)
    else:
        st.markdown(
            f"**Current page:** {view_type}<br>**Context:** {active_context.get('university')}, {active_context.get('year')}",
            unsafe_allow_html=True,
        )

    if st.button("Analyze current page", type="primary"):
        with st.spinner("Generating page-specific interpretation..."):
            st.session_state.ai_answer = generate_ai_interpretation(active_context, question or None)

    if st.session_state.ai_answer:
        st.markdown(st.session_state.ai_answer)

    with st.expander("Current AI input"):
        st.json(active_context)

    st.info(
        "The AI assistant supports interpretation only. It analyzes the current page context and does not provide causal conclusions or autonomous policy recommendations."
    )
