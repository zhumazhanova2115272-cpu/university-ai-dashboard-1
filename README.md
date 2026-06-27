# AI-enhanced Visual Analytics Dashboard for Italian Universities

This project is a Streamlit prototype for exploring multidimensional university performance profiles in Italy.

## Dataset

The app uses `data/university_dashboard_ready_dataset_v2_checked.xlsx`, sheet `Dashboard_Ready_Data_v2`.

The dataset includes:
- university, year, region, macro-area, and size class
- teaching, placement, research, financial, and overall scores
- financial indicators
- yearly ranks
- national and macro-area benchmark averages
- placement data completeness flags

## Dashboard Features

- University and year filters
- Performance score cards
- Benchmark comparison with national and macro-area averages
- Overall score trend over time
- Finance explorer scatterplot
- AI Analysis Companion side panel

## AI Companion

The AI assistant receives only the selected dashboard context and generates an interpretation of the current view. It should not be used for causal conclusions or autonomous policy recommendations.

## How to Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Secrets

To enable the AI assistant, add your OpenAI API key in Streamlit Cloud secrets:

```toml
OPENAI_API_KEY = "your_api_key_here"
```
