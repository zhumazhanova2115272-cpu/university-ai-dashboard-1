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

    tech = ""
    if technical_note:
        tech = f"\n\n**Technical note**\n\n{technical_note}"

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

This is a dashboard-based interpretation. It does not imply causality or policy recommendations.{question_note}{tech}
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
selected = filtered[filtered["university"] == university].iloc[0]
context = selected_context(selected)

main_col, ai_col = st.columns([3.2, 1.1], gap="large")

with main_col:
    st.subheader(f"{university} - {year}")

    tab_overview, tab_profile, tab_finance, tab_teaching_research, tab_data = st.tabs(
        ["Overview", "University Profile", "Finance Explorer", "Teaching and Research", "Data and Methodology"]
    )

    with tab_overview:
        st.markdown("### System overview")
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

    with tab_finance:
        st.markdown("### Finance explorer")

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
        st.markdown("#### Financial indicator vs selected score")
        st.altair_chart((base_scatter + selected_point).properties(height=380), width="stretch")

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
        st.dataframe(filtered.sort_values("overall_score", ascending=False), width="stretch", hide_index=True)

with ai_col:
    st.subheader("AI Analysis Companion")
    st.caption("Interprets only the selected dashboard context.")

    question = st.text_area(
        "Optional question",
        placeholder="Example: Why is the financial score lower than the research score?",
        height=110,
    )

    if "ai_answer" not in st.session_state:
        st.session_state.ai_answer = ""

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
