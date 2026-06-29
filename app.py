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


def selected_context(row: pd.Series) -> dict:
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
        "employment_index",
        "graduation_intensity",
        "publications_per_teaching_staff",
        "citations_per_publication",
        "h_index",
    ]
    return {field: clean_value(row.get(field)) for field in fields if field in row.index}


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


def generate_local_interpretation(
    context: dict, user_question: str | None = None, technical_note: str | None = None
) -> str:
    overall = float(context.get("overall_score") or 0)
    teaching = float(context.get("teaching_score") or 0)
    placement = float(context.get("placement_score") or 0)
    research = float(context.get("research_score") or 0)
    financial = float(context.get("financial_score") or 0)
    national = float(context.get("national_avg_overall_score") or 0)
    macro = float(context.get("macro_area_avg_overall_score") or 0)
    rank = int(context.get("overall_rank_year") or 0)

    dimensions = {
        "teaching": teaching,
        "placement": placement,
        "research": research,
        "financial": financial,
    }
    strongest = max(dimensions, key=dimensions.get)
    weakest = min(dimensions, key=dimensions.get)
    national_gap = overall - national
    macro_gap = overall - macro

    benchmark_sentence = (
        f"The overall score is {abs(national_gap):.1f} points "
        f"{'above' if national_gap >= 0 else 'below'} the national average and "
        f"{abs(macro_gap):.1f} points {'above' if macro_gap >= 0 else 'below'} the macro-area average."
    )

    question_note = ""
    if user_question:
        question_note = (
            f"\n\n**User question**\n\n{user_question}\n\n"
            "Based on the selected indicators, the answer should focus only on the current score profile, "
            "benchmarks, and visual patterns. The dashboard does not provide causal evidence."
        )

    tech = ""  # Technical notes are kept out of the presentation view.

    return f"""
**Short summary**

{context.get('university')} in {context.get('year')} has a {score_label(overall)} overall profile with an overall score of **{overall:.1f}** and rank **{rank}/61**. {benchmark_sentence}

**Strengths**

- The strongest dimension is **{strongest}** with a score of **{dimensions[strongest]:.1f}**.
- The profile should be interpreted together with the university size class: **{context.get('size_class')}**.

**Weaknesses / points to monitor**

- The weakest dimension is **{weakest}** with a score of **{dimensions[weakest]:.1f}**.
- The financial profile should be checked together with FFO per student, operating cost per student, and personnel cost share.

**Suggested next visual exploration**

Compare this university with other institutions in **{context.get('region')}** and **{context.get('macro_area')}**, then inspect whether the same pattern appears across 2020-2023.

**Important limitation**

This is a dashboard-based interpretation. It does not imply causality or policy recommendations.{question_note}
"""


def generate_ai_interpretation(context: dict, user_question: str | None = None) -> str:
    try:
        api_key = st.secrets.get("OPENAI_API_KEY", None)
    except Exception:
        api_key = None

    if OpenAI is None:
        return generate_local_interpretation(context, user_question, "The OpenAI package is not installed.")

    if not is_valid_api_key(api_key):
        return generate_local_interpretation(
            context,
            user_question,
            "No valid external API key was found. Local dashboard interpretation was shown instead.",
        )

    client = OpenAI(api_key=api_key.strip())
    prompt = {
        "task": "Interpret the selected university dashboard profile.",
        "rules": [
            "Use only the provided dashboard context.",
            "Do not invent missing values.",
            "Do not make causal claims.",
            "Write in concise academic English.",
            "Return: short summary, strengths, weaknesses, benchmark comparison, suggested next visual exploration.",
        ],
        "selected_context": context,
        "user_question": user_question,
    }

    try:
        response = client.responses.create(
            model="gpt-4.1-mini",
            input=json.dumps(prompt, ensure_ascii=True),
        )
        return response.output_text
    except Exception as exc:
        return generate_local_interpretation(
            context,
            user_question,
            f"The external AI API call failed. Error type: {type(exc).__name__}.",
        )



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
        "student_staff_ratio",
        "employment_index",
        "publications_per_teaching_staff",
        "citations_per_publication",
        "h_index",
    ]
    return {
        "comparison_year": clean_value(row_a.get("year")),
        "university_a": {field: clean_value(row_a.get(field)) for field in fields if field in row_a.index},
        "university_b": {field: clean_value(row_b.get(field)) for field in fields if field in row_b.index},
    }


def generate_local_comparison_interpretation(context: dict, user_question: str | None = None) -> str:
    a = context["university_a"]
    b = context["university_b"]
    dims = ["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"]

    def f(x: Any) -> float:
        return float(x or 0)

    a_name = a.get("university")
    b_name = b.get("university")
    year_value = context.get("comparison_year")
    a_overall = f(a.get("overall_score"))
    b_overall = f(b.get("overall_score"))
    overall_gap = a_overall - b_overall
    leader = a_name if overall_gap >= 0 else b_name
    lagger = b_name if overall_gap >= 0 else a_name

    gap_rows = []
    for dim in dims:
        gap_rows.append((dim.replace("_score", ""), f(a.get(dim)) - f(b.get(dim))))
    largest_gap_dim, largest_gap = max(gap_rows, key=lambda x: abs(x[1]))

    a_strongest = max(dims[1:], key=lambda dim: f(a.get(dim))).replace("_score", "")
    b_strongest = max(dims[1:], key=lambda dim: f(b.get(dim))).replace("_score", "")
    a_weakest = min(dims[1:], key=lambda dim: f(a.get(dim))).replace("_score", "")
    b_weakest = min(dims[1:], key=lambda dim: f(b.get(dim))).replace("_score", "")

    question_note = ""
    if user_question:
        question_note = (
            f"\n\n**User question**\n\n{user_question}\n\n"
            "The answer is based only on the selected dashboard indicators and should not be interpreted as causal evidence."
        )

    return f"""
**Short comparison summary**

In {year_value}, **{leader}** has the higher overall score. The overall gap between **{a_name}** and **{b_name}** is **{abs(overall_gap):.1f}** points, with **{lagger}** lower on the selected overall profile.

**Main difference visible in the charts**

- The largest score gap is in **{largest_gap_dim}**, equal to **{abs(largest_gap):.1f}** points.
- **{a_name}** is strongest in **{a_strongest}** and weakest in **{a_weakest}**.
- **{b_name}** is strongest in **{b_strongest}** and weakest in **{b_weakest}**.

**What to inspect next**

Compare the financial indicators, especially FFO per student, operating cost per student, and personnel cost share. Then check the time trend to see whether the gap is stable or only appears in {year_value}.

**Important limitation**

This comparison is exploratory and dashboard-based. It does not imply causality or policy recommendations.{question_note}
"""


def generate_ai_comparison_interpretation(context: dict, user_question: str | None = None) -> str:
    try:
        api_key = st.secrets.get("OPENAI_API_KEY", None)
    except Exception:
        api_key = None

    if OpenAI is None or not is_valid_api_key(api_key):
        return generate_local_comparison_interpretation(context, user_question)

    client = OpenAI(api_key=api_key.strip())
    prompt = {
        "task": "Compare two universities in the selected dashboard year.",
        "rules": [
            "Use only the provided dashboard context.",
            "Do not invent missing values.",
            "Do not make causal claims.",
            "Write in concise academic English.",
            "Return: short comparison summary, main visible differences, strengths and weaknesses, suggested next visual exploration.",
        ],
        "selected_comparison_context": context,
        "user_question": user_question,
    }

    try:
        response = client.responses.create(
            model="gpt-4.1-mini",
            input=json.dumps(prompt, ensure_ascii=True),
        )
        return response.output_text
    except Exception:
        return generate_local_comparison_interpretation(context, user_question)


def make_comparison_score_chart(row_a: pd.Series, row_b: pd.Series) -> alt.Chart:
    chart_df = pd.DataFrame(
        [
            {"University": row_a["university"], "Dimension": "Overall", "Score": row_a["overall_score"]},
            {"University": row_a["university"], "Dimension": "Teaching", "Score": row_a["teaching_score"]},
            {"University": row_a["university"], "Dimension": "Placement", "Score": row_a["placement_score"]},
            {"University": row_a["university"], "Dimension": "Research", "Score": row_a["research_score"]},
            {"University": row_a["university"], "Dimension": "Financial", "Score": row_a["financial_score"]},
            {"University": row_b["university"], "Dimension": "Overall", "Score": row_b["overall_score"]},
            {"University": row_b["university"], "Dimension": "Teaching", "Score": row_b["teaching_score"]},
            {"University": row_b["university"], "Dimension": "Placement", "Score": row_b["placement_score"]},
            {"University": row_b["university"], "Dimension": "Research", "Score": row_b["research_score"]},
            {"University": row_b["university"], "Dimension": "Financial", "Score": row_b["financial_score"]},
        ]
    )
    return (
        alt.Chart(chart_df)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("Dimension:N", sort=["Overall", "Teaching", "Placement", "Research", "Financial"], title="Dimension"),
            xOffset="University:N",
            y=alt.Y("Score:Q", scale=alt.Scale(domain=[0, 100]), title="Score"),
            color=alt.Color("University:N", title="University"),
            tooltip=["University", "Dimension", alt.Tooltip("Score:Q", format=".1f")],
        )
        .properties(height=340)
    )


def make_comparison_gap_chart(row_a: pd.Series, row_b: pd.Series) -> alt.Chart:
    a_name = row_a["university"]
    b_name = row_b["university"]
    chart_df = pd.DataFrame(
        {
            "Dimension": ["Overall", "Teaching", "Placement", "Research", "Financial"],
            "Gap": [
                row_a["overall_score"] - row_b["overall_score"],
                row_a["teaching_score"] - row_b["teaching_score"],
                row_a["placement_score"] - row_b["placement_score"],
                row_a["research_score"] - row_b["research_score"],
                row_a["financial_score"] - row_b["financial_score"],
            ],
        }
    )
    chart_df["Interpretation"] = chart_df["Gap"].apply(lambda x: f"{a_name} higher" if x >= 0 else f"{b_name} higher")
    return (
        alt.Chart(chart_df)
        .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
        .encode(
            y=alt.Y("Dimension:N", sort=["Overall", "Teaching", "Placement", "Research", "Financial"], title="Dimension"),
            x=alt.X("Gap:Q", title=f"Score gap: {a_name} minus {b_name}"),
            color=alt.Color("Interpretation:N", title="Direction"),
            tooltip=["Dimension", alt.Tooltip("Gap:Q", format="+.1f"), "Interpretation"],
        )
        .properties(height=300)
    )


def metric_card(label: str, value: Any, help_text: str | None = None) -> None:
    st.metric(label, value, help=help_text)


def make_score_profile_chart(row: pd.Series) -> alt.Chart:
    score_df = pd.DataFrame(
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
    return (
        alt.Chart(score_df)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("Dimension:N", sort=None, title="Dimension"),
            y=alt.Y("Score:Q", scale=alt.Scale(domain=[0, 100]), title="Score"),
            color=alt.Color("Dimension:N", legend=None),
            tooltip=["Dimension", alt.Tooltip("Score:Q", format=".1f")],
        )
        .properties(height=300)
    )


def make_benchmark_chart(row: pd.Series) -> alt.Chart:
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
            y=alt.Y("Overall score:Q", scale=alt.Scale(domain=[0, 100]), title="Overall score"),
            color=alt.Color("Benchmark:N", legend=None),
            tooltip=["Benchmark", alt.Tooltip("Overall score:Q", format=".1f")],
        )
        .properties(height=300)
    )


def make_indicator_bar(row: pd.Series, indicators: dict[str, str]) -> alt.Chart:
    chart_df = pd.DataFrame(
        {
            "Indicator": list(indicators.values()),
            "Value": [row[col] if col in row.index else None for col in indicators.keys()],
        }
    ).dropna()
    return (
        alt.Chart(chart_df)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("Value:Q", title="Value"),
            y=alt.Y("Indicator:N", sort="-x", title=""),
            tooltip=["Indicator", alt.Tooltip("Value:Q", format=".2f")],
        )
        .properties(height=280)
    )


def format_number(value: float, digits: int = 1) -> str:
    if pd.isna(value):
        return "n/a"
    return f"{value:,.{digits}f}"


# Load data
df = load_data()

st.title("AI-enhanced Visual Analytics Dashboard for Italian Universities")
st.caption(
    "Prototype dashboard for exploring teaching, placement, research, and financial profiles "
    "of Italian universities, 2020-2023."
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 2.0rem; padding-bottom: 2rem; }
    div[data-testid="stMetric"] { background-color: #fafafa; padding: 0.7rem 0.8rem; border-radius: 0.7rem; border: 1px solid #eeeeee; }
    div[data-testid="stSidebar"] { background-color: #f5f7fb; }
    .small-note { color: #6b7280; font-size: 0.90rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

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

university = st.sidebar.selectbox("University", sorted(filtered["university"].unique()))
st.sidebar.caption("Tip: change filters first, then choose a university and press Analyze current view.")

selected = filtered[filtered["university"] == university].iloc[0]
context = selected_context(selected)

main_col, ai_col = st.columns([3.2, 1.1], gap="large")

with main_col:
    st.subheader(f"{university} - {year}")
    st.markdown(
        f"<span class='small-note'>Region: <b>{selected['region']}</b> | Macro-area: <b>{selected['macro_area']}</b> | Size class: <b>{selected['size_class']}</b></span>",
        unsafe_allow_html=True,
    )
    st.write("")

    tab_overview, tab_profile, tab_comparison, tab_finance, tab_teaching_research, tab_data = st.tabs(
        ["Overview", "University Profile", "University Comparison", "Finance Explorer", "Teaching and Research", "Data and Methodology"]
    )

    with tab_overview:
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
                y=alt.Y("university:N", sort="-x", title="University"),
                tooltip=["university", "region", "macro_area", alt.Tooltip("overall_score:Q", format=".1f")],
            )
            .properties(height=330)
        )

        macro_summary = (
            filtered.groupby("macro_area", as_index=False)[
                ["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"]
            ]
            .mean()
            .melt(id_vars="macro_area", var_name="Score type", value_name="Average score")
        )
        macro_chart = (
            alt.Chart(macro_summary)
            .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
            .encode(
                x=alt.X("macro_area:N", title="Macro-area"),
                y=alt.Y("Average score:Q", scale=alt.Scale(domain=[0, 100])),
                color=alt.Color("Score type:N"),
                tooltip=["macro_area", "Score type", alt.Tooltip("Average score:Q", format=".1f")],
            )
            .properties(height=330)
        )

        oc1, oc2 = st.columns(2)
        with oc1:
            st.markdown("#### Top 10 universities by overall score")
            st.altair_chart(top_chart, width="stretch")
        with oc2:
            st.markdown("#### Average scores by macro-area")
            st.altair_chart(macro_chart, width="stretch")

        hist = (
            alt.Chart(filtered)
            .mark_bar()
            .encode(
                x=alt.X("overall_score:Q", bin=alt.Bin(maxbins=15), title="Overall score"),
                y=alt.Y("count():Q", title="Number of universities"),
                tooltip=["count()"],
            )
            .properties(height=260)
        )
        st.markdown("#### Distribution of overall scores")
        st.altair_chart(hist, width="stretch")

    with tab_profile:
        st.markdown("### University profile")
        st.caption("This page compares the selected university with national and macro-area benchmarks and shows score dynamics over time.")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Overall score", f"{selected['overall_score']:.1f}")
        k2.metric("Rank", f"{int(selected['overall_rank_year'])}/61")
        k3.metric("Teaching", f"{selected['teaching_score']:.1f}")
        k4.metric("Research", f"{selected['research_score']:.1f}")
        k5.metric("Financial", f"{selected['financial_score']:.1f}")

        if "placement_data_completeness" in selected.index and selected["placement_data_completeness"] != "complete":
            st.warning(
                "Placement score may be based on incomplete placement components. "
                f"Completeness flag: {selected['placement_data_completeness']}"
            )

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Performance profile")
            st.altair_chart(make_score_profile_chart(selected), width="stretch")
        with c2:
            st.markdown("#### Benchmark comparison")
            st.altair_chart(make_benchmark_chart(selected), width="stretch")

        trend = df[df["university"] == university].sort_values("year")
        trend_long = trend.melt(
            id_vars=["university", "year"],
            value_vars=["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"],
            var_name="Score type",
            value_name="Score",
        )
        trend_chart = (
            alt.Chart(trend_long)
            .mark_line(point=True)
            .encode(
                x=alt.X("year:O", title="Year"),
                y=alt.Y("Score:Q", title="Score", scale=alt.Scale(domain=[0, 100])),
                color=alt.Color("Score type:N"),
                tooltip=["year", "Score type", alt.Tooltip("Score:Q", format=".1f")],
            )
            .properties(height=300)
        )
        st.markdown("#### Score dynamics over time")
        st.altair_chart(trend_chart, width="stretch")

    with tab_comparison:
        st.markdown("### University comparison")
        st.caption("Select two universities in the same year and compare their scores, ranks, financial indicators, and time trends. The comparison is exploratory and does not imply causality.")

        comparison_base = df[df["year"] == year].copy()
        if macro_area != "All":
            comparison_base = comparison_base[comparison_base["macro_area"] == macro_area]
        if region != "All":
            comparison_base = comparison_base[comparison_base["region"] == region]
        if size_class != "All":
            comparison_base = comparison_base[comparison_base["size_class"] == size_class]

        comparison_universities = sorted(comparison_base["university"].unique())
        if len(comparison_universities) < 2:
            st.warning("At least two universities are required for comparison. Please broaden the filters.")
        else:
            default_a = comparison_universities.index(university) if university in comparison_universities else 0
            default_b = 1 if default_a == 0 else 0
            cc1, cc2 = st.columns(2)
            with cc1:
                university_a = st.selectbox("University A", comparison_universities, index=default_a)
            with cc2:
                university_b = st.selectbox("University B", comparison_universities, index=default_b)

            if university_a == university_b:
                st.warning("Please choose two different universities.")
            else:
                row_a = comparison_base[comparison_base["university"] == university_a].iloc[0]
                row_b = comparison_base[comparison_base["university"] == university_b].iloc[0]
                comp_context = comparison_context(row_a, row_b)

                m1, m2, m3, m4 = st.columns(4)
                m1.metric(f"{university_a} overall", f"{row_a['overall_score']:.1f}", f"rank {int(row_a['overall_rank_year'])}/61")
                m2.metric(f"{university_b} overall", f"{row_b['overall_score']:.1f}", f"rank {int(row_b['overall_rank_year'])}/61")
                m3.metric("Overall gap", f"{abs(row_a['overall_score'] - row_b['overall_score']):.1f}")
                m4.metric("Comparison year", f"{year}")

                comp_col1, comp_col2 = st.columns(2)
                with comp_col1:
                    st.markdown("#### Side-by-side score profile")
                    st.altair_chart(make_comparison_score_chart(row_a, row_b), width="stretch")
                with comp_col2:
                    st.markdown("#### Score gaps")
                    st.altair_chart(make_comparison_gap_chart(row_a, row_b), width="stretch")

                trend_pair = df[df["university"].isin([university_a, university_b])].sort_values(["university", "year"])
                trend_pair_long = trend_pair.melt(
                    id_vars=["university", "year"],
                    value_vars=["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"],
                    var_name="Score type",
                    value_name="Score",
                )
                selected_trend_score = st.selectbox(
                    "Score trend to compare",
                    ["overall_score", "teaching_score", "placement_score", "research_score", "financial_score"],
                    format_func=lambda x: x.replace("_", " ").title(),
                )
                trend_selected = trend_pair_long[trend_pair_long["Score type"] == selected_trend_score]
                trend_pair_chart = (
                    alt.Chart(trend_selected)
                    .mark_line(point=True)
                    .encode(
                        x=alt.X("year:O", title="Year"),
                        y=alt.Y("Score:Q", title="Score", scale=alt.Scale(domain=[0, 100])),
                        color=alt.Color("university:N", title="University"),
                        tooltip=["university", "year", alt.Tooltip("Score:Q", format=".1f")],
                    )
                    .properties(height=300)
                )
                st.markdown("#### Trend comparison over 2020-2023")
                st.altair_chart(trend_pair_chart, width="stretch")

                comparison_table_cols = [
                    "university",
                    "region",
                    "macro_area",
                    "size_class",
                    "overall_rank_year",
                    "overall_score",
                    "teaching_score",
                    "placement_score",
                    "research_score",
                    "financial_score",
                    "ffo_per_student",
                    "operating_cost_per_student",
                    "personnel_cost_share",
                    "student_staff_ratio",
                    "employment_index",
                    "publications_per_teaching_staff",
                ]
                st.markdown("#### Key indicator table")
                st.dataframe(pd.DataFrame([row_a, row_b])[comparison_table_cols], width="stretch", hide_index=True)

                st.markdown("#### AI comparison interpretation")
                comparison_question = st.text_area(
                    "Optional question for the comparison",
                    placeholder="Example: Which university has a more balanced profile and why?",
                    height=90,
                    key="comparison_question",
                )
                if "comparison_ai_answer" not in st.session_state:
                    st.session_state.comparison_ai_answer = ""
                if st.button("Analyze comparison", type="primary"):
                    with st.spinner("Generating comparison interpretation..."):
                        st.session_state.comparison_ai_answer = generate_ai_comparison_interpretation(
                            comp_context, comparison_question or None
                        )
                if st.session_state.comparison_ai_answer:
                    st.markdown(st.session_state.comparison_ai_answer)
                with st.expander("Selected comparison AI input"):
                    st.json(comp_context)

    with tab_finance:
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
            .encode(
                x=alt.X(f"{x_col}:Q"),
                y=alt.Y(f"{y_col}:Q"),
                tooltip=["university"],
            )
        )
        selected_label = (
            alt.Chart(scatter_data[scatter_data["selected_flag"]])
            .mark_text(align="left", dx=10, dy=-10, fontSize=12, fontWeight="bold")
            .encode(
                x=alt.X(f"{x_col}:Q"),
                y=alt.Y(f"{y_col}:Q"),
                text="university:N",
            )
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

    with tab_teaching_research:
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

    with tab_data:
        st.markdown("### Data and methodology")
        st.markdown(
            "The dashboard uses a curated prototype dataset covering 61 Italian universities over 2020-2023. "
            "The score fields are dashboard-based normalized profile scores, not DEA or SFA efficiency scores. "
            "They are intended for visual exploration and comparison."
        )
        st.markdown("#### Selected AI input")
        st.json(context)
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
    st.caption("Interprets only the selected dashboard context. The output supports exploration, not causal conclusions.")

    question = st.text_area(
        "Optional question",
        placeholder="Example: Why is the financial score lower than the research score?",
        height=110,
    )

    if "ai_answer" not in st.session_state:
        st.session_state.ai_answer = ""

    st.markdown(
        f"**Current context:** {university}, {year}<br>Overall score: **{selected['overall_score']:.1f}** | Rank: **{int(selected['overall_rank_year'])}/61**",
        unsafe_allow_html=True,
    )

    if st.button("Analyze current view", type="primary"):
        with st.spinner("Generating interpretation..."):
            st.session_state.ai_answer = generate_ai_interpretation(context, question or None)

    if st.session_state.ai_answer:
        st.markdown(st.session_state.ai_answer)

    with st.expander("Selected AI input"):
        st.json(context)

    st.info(
        "The AI assistant supports interpretation only. It does not provide causal conclusions "
        "or autonomous policy recommendations."
    )
