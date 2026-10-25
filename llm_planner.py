"""
llm_planner.py — Groq LLM → structured JSON slide plan (Day 2)

Sends the aggregated data summary (from data_parser.py) + user prompt to
the Groq API and gets back a structured slide plan as JSON.

The LLM NEVER receives raw row-level data — only pre-computed aggregates.
It decides slide structure, chart types, and writes narrative text.

Pipeline position:
    Excel → [data_parser] → summary dict → [llm_planner] → slide plan JSON → PPTX builder
"""

import json
import os
import re
import sys
from typing import Any, Dict, Optional

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    from groq import Groq
except ImportError:
    Groq = None

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEFAULT_GEMINI_MODEL = "gemini-1.5-flash"
DEFAULT_GROQ_MODEL = "llama-3.3-70b-versatile"
DEFAULT_MODEL = "gemini-1.5-flash"
MAX_RETRIES = 1  # retry once on JSON parse failure

VALID_CHART_TYPES = {"bar", "pie", "line", "table", "text"}

# ---------------------------------------------------------------------------
# System prompt — the core instruction set for the LLM
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """\
You are a professional presentation planner for business/sales decks.

You will receive:
1. An AGGREGATED DATA SUMMARY from a multi-tab Excel workbook. This summary
   contains ONLY computed statistics (sums, means, value counts, date ranges)
   — never raw row-level data. You must plan your slides using ONLY what
   appears in this summary.
2. The USER'S INSTRUCTIONS describing the focus, audience, and tone.

Your job is to produce a structured slide plan as a JSON array.

## RULES — read carefully:

- Respond with ONLY valid JSON. No markdown fences, no preamble, no explanation.
- The JSON must be an array of slide objects.
- Each slide object has these exact keys:
  {
    "title":        "Slide title (string) - MUST be unique, descriptive, and reflect the specific analysis (e.g. 'Supplier Cost Distribution' vs 'Supplier Cost Trend Over Time'). Never duplicate titles.",
    "source_sheet": "Name of the Excel sheet this slide draws data from (string)",
    "chart_type":   "One of: bar, pie, line, table, text (string)",
    "data_columns": ["list", "of", "column", "names", "to", "chart/table"],
    "insight_text": "A structured string starting with a bold header followed by 2-3 bullet points. Do NOT state specific numeric values in insight_text — describe trends, rankings, and relationships in words and placeholders only, because you do not have access to exact computed totals. Instead, you MUST use the following placeholders where appropriate so the server can inject the verified numbers:\n    - {{TOP_CATEGORY}}: Category name with the highest value (e.g., 'History')\n    - {{TOP_VALUE}}: The highest aggregated value (e.g., '514')\n    - {{TOP_PCT}}: The percentage share of the top category (e.g., '36.4%')\n    - {{SECOND_CATEGORY}}: Category name with the second highest value\n    - {{SECOND_VALUE}}: The second highest aggregated value\n    - {{SECOND_PCT}}: The percentage share of the second category\n    - {{BOTTOM_CATEGORY}}: Category name with the lowest value\n    - {{BOTTOM_VALUE}}: The lowest aggregated value\n    - {{TOTAL_VALUE}}: The sum of all values in the main value column\n    - {{MEAN_VALUE}}: The average of all values in the main value column\n    - {{COUNT}}: The count of records/categories\n    Format: **[Short Bold Heading]**\\n• **{{TOP_CATEGORY}}**: Leads with {{TOP_VALUE}} units, representing {{TOP_PCT}} of total...\\n• **{{SECOND_CATEGORY}}**: Represents the second highest share at {{SECOND_VALUE}}...\\n• **Takeaway**: General strategic note.",
    "kpis":         [
                      {"label": "Lowest Category Stock", "type": "extreme_category", "group_by": "Category", "value_col": "Current Stock", "agg": "sum", "which": "min"},
                      {"label": "Average Stock per Supplier", "type": "avg_per_group", "group_by": "Supplier", "value_col": "Current Stock", "agg": "sum"},
                      {"label": "Reorder Rate", "type": "filtered_rate", "filter_col": "Reorder Status", "filter_contains": "Reorder"}
                    ]
  }

- Note: Every data slide (bar/pie/line/table) MUST have 2-3 KPI definitions under the 'kpis' key. Each KPI is a specification dictionary that MUST match one of the following exact shapes (depending on the type of metric needed):
  1) {"label": "Display Label", "type": "extreme_category", "group_by": "<category_column>", "value_col": "<value_column>", "agg": "sum"|"mean"|"count"|"max"|"min", "which": "min"|"max"} (computes min/max category and its value)
  2) {"label": "Display Label", "type": "avg_per_group", "group_by": "<category_column>", "value_col": "<value_column>", "agg": "sum"|"mean"|"count"|"max"|"min"} (computes mean of the grouped sums/means)
  3) {"label": "Display Label", "type": "filtered_rate", "filter_col": "<filter_column>", "filter_contains": "<substring_to_match>"} (computes % of rows matching filter)
  4) {"label": "Display Label", "type": "filtered_sum", "filter_col": "<filter_column>", "filter_contains": "<substring_to_match>", "value_col": "<value_column>"} (sums a column for matching rows only)
  5) {"label": "Display Label", "type": "peak_point", "x_col": "<date_or_x_column>", "value_col": "<numeric_value_column>"} (finds date/x at maximum value)
  Do NOT put numeric values in the slide plan — the system will compute all values using these specifications. For "text" slides, set 'kpis' to [].
- For "text" chart_type slides (e.g. title slide, summary slide), set
  data_columns to an empty list [] and source_sheet to "none".

## VISUALIZATION CHOICES (VARY TYPES & MAP TO USE-CASES):
Vary the chart types across the deck to keep it engaging. Never use the same chart type three times in a row. Choose chart type based on these mappings:
1. **Pie Chart (`pie`)** → Use for part-to-whole share, composition, or proportions (e.g. "genre-wise billing split", "distribution of stock by category").
2. **Line Chart (`line`)** → Use for trends over time, chronological tracking, or sequence-based changes (e.g. "daily total revenue trends", "monthly orders timeline").
3. **Bar Chart (`bar`)** → Use for rankings, comparison of separate categories, or multi-series variables (e.g. "top 5 suppliers by stock levels", "revenue vs spend by department").
4. **Table (`table`)** → Use for detailed lists with 3+ columns, raw parameter comparisons, or status tables (e.g. "reorder status list showing current stock, minimum stock, and unit cost").

## EXAMPLES OF PROMPT → CHART-TYPE MAPPINGS:
- *Prompt:* "Show me where most of our money is spent and how it has changed this quarter."
  → Slide 1 (Pie): "Cost Distribution by Category" (Columns: Category, Cost)
  → Slide 2 (Line): "Weekly Spending Trend" (Columns: Date, Cost)
- *Prompt:* "What are our stock levels and which items are low?"
  → Slide 1 (Bar): "Top 5 Categories by Stock Quantity" (Columns: Category, Current Stock)
  → Slide 2 (Table): "Low Stock Inventory Alerts" (Columns: Title, Current Stock, Minimum Required Stock, Reorder Status)

## PRIVACY / IDENTIFIER RULES — CRITICAL:
- Some columns may have "identifier_like": true. These contain client names, order IDs, etc.
- NEVER put these identifier columns in data_columns and NEVER mention their values in insight_text or KPIs.

## SLIDE PLANNING GUIDELINES:
- Start with a title/overview slide (chart_type: "text").
- End with a key takeaways / summary slide (chart_type: "text").
- Distribute slides across sheets. Ensure titles are descriptive and distinct.
- Avoid leaving slides looking empty: encourage multi-variable combinations. If a sheet has multiple numeric columns (e.g. Revenue and Units Sold), include both in the data_columns list to chart them together rather than just charting a single column!
- Use actual column and sheet names. Produce exactly the number of slides requested.
"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _condense_summary(summary: dict) -> dict:
    """
    Produce a highly compressed version of the data summary to fit within
    strict free-tier token limits (6000 TPM). Strips all non-essential metrics.
    """
    condensed = {}
    for sheet_name, sheet_info in summary.items():
        entry = {
            "row_count": sheet_info.get("row_count", 0),
            "columns": {},
        }
        if "note" in sheet_info:
            entry["note"] = sheet_info["note"]

        for col_name, col_info in sheet_info.get("columns", {}).items():
            slim = {"type": col_info.get("inferred_type", "unknown")}

            # Only copy crucial planning stats
            for key in ("sum", "mean", "min_date", "max_date", "unique_values"):
                if key in col_info and col_info[key] is not None:
                    slim[key] = col_info[key]

            # Limit top values to top 3 to save tokens
            if "top_values" in col_info and col_info["top_values"]:
                slim["top_values"] = dict(list(col_info["top_values"].items())[:3])

            if col_info.get("identifier_like"):
                slim["identifier_like"] = True

            entry["columns"][col_name] = slim

        condensed[sheet_name] = entry
    return condensed


def _strip_markdown_fences(text: str) -> str:
    """
    Strip markdown code fences (```json ... ```) if the model wraps its
    response in them despite being told not to.
    """
    text = text.strip()
    # Remove opening fence with optional language tag
    text = re.sub(r"^```(?:json)?\s*\n?", "", text)
    # Remove closing fence
    text = re.sub(r"\n?```\s*$", "", text)
    return text.strip()


def _parse_slide_plan(raw_text: str) -> list:
    """
    Parse the LLM's raw response into a validated list of slide dicts.
    Raises ValueError if the JSON is invalid or doesn't match the schema.
    """
    cleaned = _strip_markdown_fences(raw_text)

    try:
        parsed = json.loads(cleaned)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON from LLM: {e}\n\nRaw response:\n{raw_text}")

    # Accept both {"slides": [...]} and bare [...]
    if isinstance(parsed, dict) and "slides" in parsed:
        slides = parsed["slides"]
    elif isinstance(parsed, list):
        slides = parsed
    else:
        raise ValueError(
            f"Expected a JSON array of slides or {{\"slides\": [...]}}, "
            f"got {type(parsed).__name__}"
        )

    # Validate each slide
    required_keys = {"title", "source_sheet", "chart_type", "data_columns", "insight_text"}
    for i, slide in enumerate(slides):
        if not isinstance(slide, dict):
            raise ValueError(f"Slide {i} is not a dict: {slide}")
        missing = required_keys - slide.keys()
        if missing:
            raise ValueError(f"Slide {i} missing keys: {missing}")
        # Validate and normalize KPIs to list-of-dicts format
        raw_kpis = slide.get("kpis", [])
        if isinstance(raw_kpis, dict):
            # Legacy format {label: value} — convert to new descriptive format
            raw_kpis = []
        if not isinstance(raw_kpis, list):
            raw_kpis = []
        # Keep only well-formed KPI defs (having 'label' and 'type')
        valid_kpis = []
        for kpi in raw_kpis:
            if isinstance(kpi, dict) and "label" in kpi and "type" in kpi:
                valid_kpis.append(kpi)
        slide["kpis"] = valid_kpis
        # Normalise chart_type
        slide["chart_type"] = slide["chart_type"].lower().strip()
        if slide["chart_type"] not in VALID_CHART_TYPES:
            raise ValueError(
                f"Slide {i} has invalid chart_type '{slide['chart_type']}'. "
                f"Must be one of: {VALID_CHART_TYPES}"
            )
        if not isinstance(slide["data_columns"], list):
            slide["data_columns"] = [slide["data_columns"]]

    return slides


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def _call_gemini_api(api_key: str, system_prompt: str, user_message: str, model_name: str = DEFAULT_GEMINI_MODEL) -> str:
    """Calls Google Gemini API using google-genai or google-generativeai."""
    errors = []
    # Attempt 1: modern google.genai SDK
    try:
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model=model_name,
            contents=user_message,
            config=types.GenerateContentConfig(
                system_instruction=system_prompt,
                temperature=0.4,
                response_mime_type="application/json",
            ),
        )
        if response and response.text:
            return response.text
    except Exception as e:
        errors.append(f"google.genai error: {e}")

    # Attempt 2: google.generativeai SDK
    try:
        import google.generativeai as legacy_genai
        legacy_genai.configure(api_key=api_key)
        model = legacy_genai.GenerativeModel(
            model_name=model_name,
            system_instruction=system_prompt,
            generation_config={
                "temperature": 0.4,
                "response_mime_type": "application/json",
            },
        )
        response = model.generate_content(user_message)
        if response and response.text:
            return response.text
    except Exception as e:
        errors.append(f"google.generativeai error: {e}")

    raise RuntimeError(f"Gemini API call failed: {' | '.join(errors)}")


def _call_groq_api(api_key: str, system_prompt: str, user_message: str, model_name: str = DEFAULT_GROQ_MODEL) -> str:
    """Calls Groq API."""
    if Groq is None:
        raise RuntimeError("groq package is not installed. Install with: pip install groq")
    client = Groq(api_key=api_key)
    response = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        temperature=0.4,
        max_tokens=4096,
    )
    return response.choices[0].message.content


def generate_slide_plan(
    data_summary: dict,
    user_prompt: str,
    num_slides: int = 8,
    model: Optional[str] = None,
    api_key: Optional[str] = None,
) -> dict:
    """
    Generate a structured slide plan using Gemini (default) or Groq.

    Parameters
    ----------
    data_summary : dict
        The aggregate-only summary from data_parser.parse_excel().
    user_prompt : str
        The user's instructions (focus areas, audience, tone, etc.).
    num_slides : int
        Number of slides to generate (default 8).
    model : str | None
        Model ID (default: gemini-1.5-flash for Gemini, llama-3.3-70b-versatile for Groq).
    api_key : str | None
        API key. Falls back to GEMINI_API_KEY / GROQ_API_KEY from secrets or .env.

    Returns
    -------
    dict with key "slides" containing the list of slide plan dicts.
    """
    gemini_key = api_key if (api_key and not api_key.startswith("gsk_")) else None
    groq_key = api_key if (api_key and api_key.startswith("gsk_")) else None

    # Check Streamlit secrets first
    try:
        import streamlit as st
        if not gemini_key:
            gemini_key = st.secrets.get("GEMINI_API_KEY") or st.secrets.get("GOOGLE_API_KEY")
        if not groq_key:
            groq_key = st.secrets.get("GROQ_API_KEY")
    except Exception:
        pass

    # Check environment variables / .env
    if not gemini_key:
        gemini_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not groq_key:
        groq_key = os.environ.get("GROQ_API_KEY")

    provider = None
    resolved_key = None
    chosen_model = model

    if gemini_key and gemini_key.strip():
        provider = "gemini"
        resolved_key = gemini_key.strip()
        if not chosen_model:
            chosen_model = DEFAULT_GEMINI_MODEL
    elif groq_key and groq_key.strip():
        provider = "groq"
        resolved_key = groq_key.strip()
        if not chosen_model:
            chosen_model = DEFAULT_GROQ_MODEL
    else:
        raise RuntimeError(
            "No AI API key found. Please set GEMINI_API_KEY (recommended) or GROQ_API_KEY "
            "in your .env file or Streamlit Cloud Settings -> Secrets."
        )

    # Condense the summary to fit within token limits
    compact_summary = _condense_summary(data_summary)

    # Build the user message with summary + instructions
    user_message = (
        f"DATA SUMMARY:\n"
        f"{json.dumps(compact_summary, default=str)}\n\n"
        f"USER INSTRUCTIONS: {user_prompt}\n\n"
        f"Produce exactly {num_slides} slides. Respond with ONLY the JSON array."
    )

    last_error = None

    for attempt in range(1 + MAX_RETRIES):
        try:
            msg = user_message
            if attempt > 0:
                msg += (
                    "\n\nWARNING: YOUR PREVIOUS RESPONSE WAS NOT VALID JSON. "
                    "Respond with ONLY a raw JSON array — no markdown fences, "
                    "no explanation, no preamble. Just the [ ... ] array."
                )

            if provider == "gemini":
                raw_text = _call_gemini_api(resolved_key, SYSTEM_PROMPT, msg, chosen_model)
            else:
                raw_text = _call_groq_api(resolved_key, SYSTEM_PROMPT, msg, chosen_model)

            slides = _parse_slide_plan(raw_text)

            return {
                "slides": slides,
                "model": chosen_model,
                "provider": provider,
                "num_slides_requested": num_slides,
                "num_slides_returned": len(slides),
            }

        except ValueError as e:
            last_error = e
            if attempt < MAX_RETRIES:
                print(f"[llm_planner] JSON parse failed (attempt {attempt + 1}), retrying...")
                continue
            else:
                raise ValueError(
                    f"Failed to get valid JSON from LLM after {1 + MAX_RETRIES} attempts. "
                    f"Last error: {last_error}"
                ) from last_error

        except Exception as e:
            raise RuntimeError(f"{provider.capitalize()} API call failed: {e}") from e



# ---------------------------------------------------------------------------
# CLI test block
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Use Day 1's data_parser to get a real summary, then generate a plan
    from data_parser import parse_excel
    import numpy as np
    import pandas as pd
    from io import BytesIO

    print("=" * 60)
    print("DAY 2 TEST: LLM Planner")
    print("=" * 60)

    # ---------- Generate the same sample workbook from Day 1 ----------
    print("\n[1/3] Generating sample multi-tab Excel...\n")
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        sales = pd.DataFrame({
            "Region": ["North", "South", "East", "West", "North", "South",
                        "East", "West", "North", "South"] * 3,
            "Product": ["Widget A", "Widget B", "Widget C", "Widget A",
                        "Widget B", "Widget C", "Widget A", "Widget B",
                        "Widget C", "Widget A"] * 3,
            "Revenue": np.random.randint(10_000, 500_000, size=30),
            "Units Sold": np.random.randint(50, 5000, size=30),
            "Date": pd.date_range("2024-01-01", periods=30, freq="M"),
        })
        sales.to_excel(writer, sheet_name="Regional Sales", index=False)

        feedback = pd.DataFrame({
            "Customer Segment": np.random.choice(
                ["Enterprise", "SMB", "Startup", "Government"], size=20
            ),
            "Satisfaction": np.random.choice(
                ["Very Satisfied", "Satisfied", "Neutral", "Dissatisfied"], size=20
            ),
            "NPS Score": np.random.randint(1, 11, size=20),
            "Response Date": pd.date_range("2024-06-01", periods=20, freq="W"),
        })
        feedback.to_excel(writer, sheet_name="Customer Feedback", index=False)

        messy = pd.DataFrame([
            ["", "", "", ""],
            ["", "", "", ""],
            ["Quarter", "Channel", "Spend", "ROI %"],
            ["Q1", "Online", 50000, 12.5],
            ["Q1", "Retail", 35000, 8.3],
            ["Q2", "Online", 62000, 15.1],
            ["Q2", "Retail", 41000, 9.7],
            ["Q3", "Online", 71000, 18.2],
            ["Q3", "Retail", 45000, 11.0],
            ["Q4", "Online", 80000, 20.5],
            ["Q4", "Retail", 52000, 13.4],
        ])
        messy.to_excel(writer, sheet_name="Marketing Spend", index=False, header=False)

    buf.seek(0)

    # ---------- Parse it ----------
    print("[2/3] Parsing Excel with data_parser...\n")
    summary, dfs = parse_excel(buf)
    print(f"Sheets found: {list(summary.keys())}")
    for name, info in summary.items():
        print(f"  {name}: {info.get('row_count', 0)} rows, {info.get('column_count', 0)} cols")

    # ---------- Generate slide plan ----------
    print("\n[3/3] Calling Groq API for slide plan...\n")
    test_prompt = (
        "Focus on regional sales performance and marketing ROI. "
        "This deck is for senior leadership. "
        "Use professional, concise language. "
        "Highlight any standout trends or areas of concern."
    )

    try:
        plan = generate_slide_plan(
            data_summary=summary,
            user_prompt=test_prompt,
            num_slides=5,
        )

        print("[OK] Slide plan generated successfully!\n")
        print(f"Model: {plan['model']}")
        print(f"Slides requested: {plan['num_slides_requested']}")
        print(f"Slides returned:  {plan['num_slides_returned']}")
        print("\n" + "-" * 60)
        print("FULL SLIDE PLAN:")
        print("-" * 60)
        print(json.dumps(plan["slides"], indent=2, ensure_ascii=True))

        # Quick validation summary
        print("\n" + "-" * 60)
        print("SLIDE SUMMARY:")
        print("-" * 60)
        for i, slide in enumerate(plan["slides"]):
            print(f"\n  Slide {i + 1}: {slide['title']}")
            print(f"    Chart:   {slide['chart_type']}")
            print(f"    Sheet:   {slide['source_sheet']}")
            print(f"    Columns: {slide['data_columns']}")
            insight_preview = slide['insight_text'][:100].encode('ascii', 'replace').decode()
            print(f"    Insight: {insight_preview}...")

    except Exception as e:
        print(f"[ERROR] {e}")
        sys.exit(1)
