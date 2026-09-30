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
