from __future__ import annotations

import json
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

try:
    from openai import OpenAI
except ImportError:  # The dashboard still works without AI if openai is not installed.
    OpenAI = None


DATA_PATH = Path("data/university_dashboard_ready_dataset_v2_checked.xlsx")
SHEET_NAME = "Dashboard_Ready_Data_v2"


st.set_page_config(
    page_title="University Performance Dashboard",
    page_icon="📊",
    layout="wide",
)


@st.cache_data
def load_data() -> pd.DataFrame:
    df = pd.read_excel(DATA_PATH, sheet_name=SHEET_NAME)
    df["year"] = df["year"].astype(int)
    return df


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
    ]
    context = {}
    for field in fields:
        value = row.get(field)
        if pd.isna(value):
            context[field] = None
        elif isinstance(value, (int, float, str)):
            context[field] = value
        else:
            context[field] = value.item() if hasattr(value, "item") else str(value)
    return context


def generate_ai_interpretation(context: dict, user_question: str | None = None) -> str:
    try:
        api_key = st.secrets.get("OPENAI_API_KEY", None)
    except Exception:
        api_key = None

    if OpenAI is None or not api_key:
        return generate_local_interpretation(context, user_question)

    client = OpenAI(api_key=api_key)
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
    response = client.responses.create(
        model="gpt-4.1-mini",
        input=json.dumps(prompt, ensure_ascii=False),
    )
    return response.output_text


def score_label(value: float) -> str:
    if value >= 70:
        return "high"
    if value >= 55:
        return "moderate"
    return "low"


def generate_local_interpretation(context: dict, user_question: str | None = None) -> str:
    """Fallback interpretation so the button works even before an API key is added."""
    overall = float(context["overall_score"])
    teaching = float(context["teaching_score"])
    placement = float(context["placement_score"])
    research = float(context["research_score"])
    financial = float(context["financial_score"])
    national = float(context["national_avg_overall_score"])
    macro = float(context["macro_area_avg_overall_score"])
    rank = int(context["overall_rank_year"])

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
            f"\n\n**User question:** {user_question}\n\n"
            "Based on the selected indicators, the answer should focus on the score profile and benchmarks only. "
            "This dashboard does not provide causal evidence."
        )

    return f"""
**Short summary**

{context['university']} in {context['year']} has a {score_label(overall)} overall profile with an overall score of **{overall:.1f}** and rank **{rank}/61**.
{benchmark_sentence}

**Strengths**

- The strongest dimension is **{strongest}** with a score of **{dimensions[strongest]:.1f}**.
- Research and teaching should be interpreted together with the university size class: **{context['size_class']}**.

**Weaknesses / points to monitor**

- The weakest dimension is **{weakest}** with a score of **{dimensions[weakest]:.1f}**.
- The financial profile should be checked together with FFO per student, operating cost per student, and personnel cost share.

**Suggested next visual exploration**

Compare this university with other institutions in **{context['region']}** and **{context['macro_area']}**, then inspect whether the same pattern appears across 2020-2023.

**Important limitation**

This is a dashboard-based interpretation. It does not imply causality or policy recommendations.
{question_note}
"""


df = load_data()

st.title("AI-enhanced Visual Analytics Dashboard for Italian Universities")
st.caption(
    "Prototype dashboard for exploring teaching, placement, research, and financial profiles "
    "of Italian universities, 2020-2023."
)

with st.sidebar:
    st.header("Filters")
    year = st.selectbox("Year", sorted(df["year"].unique()), index=len(sorted(df["year"].unique())) - 1)
    macro_area_options = ["All"] + sorted(df["macro_area"].dropna().unique())
    macro_area = st.selectbox("Macro-area", macro_area_options)
    region_options = ["All"] + sorted(df["region"].dropna().unique())
    region = st.selectbox("Region", region_options)
    size_options = ["All"] + sorted(df["size_class"].dropna().unique())
    size_class = st.selectbox("Size class", size_options)

filtered = df[df["year"] == year].copy()
if macro_area != "All":
    filtered = filtered[filtered["macro_area"] == macro_area]
if region != "All":
    filtered = filtered[filtered["region"] == region]
if size_class != "All":
    filtered = filtered[filtered["size_class"] == size_class]

university = st.sidebar.selectbox("University", sorted(filtered["university"].unique()))
selected = filtered[filtered["university"] == university].iloc[0]
context = selected_context(selected)

main_col, ai_col = st.columns([3, 1], gap="large")

with main_col:
    st.subheader(f"{university} · {year}")

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Overall score", f"{selected['overall_score']:.1f}")
    k2.metric("Rank", f"{int(selected['overall_rank_year'])}/61")
    k3.metric("Teaching", f"{selected['teaching_score']:.1f}")
    k4.metric("Research", f"{selected['research_score']:.1f}")
    k5.metric("Financial", f"{selected['financial_score']:.1f}")

    score_df = pd.DataFrame(
        {
            "Dimension": ["Teaching", "Placement", "Research", "Financial"],
            "Score": [
                selected["teaching_score"],
                selected["placement_score"],
                selected["research_score"],
                selected["financial_score"],
            ],
        }
    )
    score_chart = (
        alt.Chart(score_df)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("Dimension:N", sort=None),
            y=alt.Y("Score:Q", scale=alt.Scale(domain=[0, 100])),
            color=alt.Color("Dimension:N", legend=None),
            tooltip=["Dimension", alt.Tooltip("Score:Q", format=".1f")],
        )
        .properties(height=280)
    )

    benchmark_df = pd.DataFrame(
        {
            "Benchmark": ["Selected university", "National average", "Macro-area average"],
            "Overall score": [
                selected["overall_score"],
                selected["national_avg_overall_score"],
                selected["macro_area_avg_overall_score"],
            ],
        }
    )
    benchmark_chart = (
        alt.Chart(benchmark_df)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X("Benchmark:N", sort=None),
            y=alt.Y("Overall score:Q", scale=alt.Scale(domain=[0, 100])),
            color=alt.Color("Benchmark:N", legend=None),
            tooltip=["Benchmark", alt.Tooltip("Overall score:Q", format=".1f")],
        )
        .properties(height=280)
    )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Performance Profile")
        st.altair_chart(score_chart, use_container_width=True)
    with c2:
        st.markdown("#### Benchmark Comparison")
        st.altair_chart(benchmark_chart, use_container_width=True)

    trend = df[df["university"] == university].sort_values("year")
    trend_chart = (
        alt.Chart(trend)
        .mark_line(point=True)
        .encode(
            x=alt.X("year:O", title="Year"),
            y=alt.Y("overall_score:Q", title="Overall score", scale=alt.Scale(domain=[0, 100])),
            tooltip=["year", alt.Tooltip("overall_score:Q", format=".1f")],
        )
        .properties(height=260)
    )
    st.markdown("#### Overall Score Over Time")
    st.altair_chart(trend_chart, use_container_width=True)

    scatter = (
        alt.Chart(filtered)
        .mark_circle(size=90, opacity=0.75)
        .encode(
            x=alt.X("ffo_per_student:Q", title="FFO per student"),
            y=alt.Y("overall_score:Q", title="Overall score", scale=alt.Scale(domain=[0, 100])),
            color=alt.Color("macro_area:N", title="Macro-area"),
            size=alt.Size("enrolled_students:Q", title="Enrolled students"),
            tooltip=[
                "university",
                "region",
                "macro_area",
                alt.Tooltip("ffo_per_student:Q", format=",.0f"),
                alt.Tooltip("overall_score:Q", format=".1f"),
            ],
        )
        .interactive()
        .properties(height=340)
    )
    st.markdown("#### Finance Explorer: Funding Intensity vs Overall Score")
    st.altair_chart(scatter, use_container_width=True)

    st.markdown("#### Selected Data")
    st.dataframe(pd.DataFrame([context]), use_container_width=True)

with ai_col:
    st.subheader("AI Analysis Companion")
    st.caption("Interprets only the selected dashboard context.")

    question = st.text_area(
        "Optional question",
        placeholder="Example: Why is the financial score lower than the research score?",
        height=100,
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

