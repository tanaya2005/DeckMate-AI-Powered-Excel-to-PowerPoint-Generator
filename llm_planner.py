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

from dotenv import load_dotenv
from groq import Groq

load_dotenv()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEFAULT_MODEL = "llama-3.1-8b-instant"
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
    "title":        "Slide title (string)",
    "source_sheet": "Name of the Excel sheet this slide draws data from (string)",
    "chart_type":   "One of: bar, pie, line, table, text (string)",
    "data_columns": ["list", "of", "column", "names", "to", "chart"],
    "insight_text": "2-3 sentence written interpretation of what the data shows. Written in professional business language appropriate for the stated audience."
  }

- For "text" chart_type slides (e.g. title slide, summary slide), set
  data_columns to an empty list [] and source_sheet to "none".
- Choose chart_type thoughtfully:
  • bar  → comparing categories or groups
  • pie  → showing composition / share of a whole
  • line → showing trends over time (requires a date/time column)
  • table → when exact numbers matter more than visual shape
  • text  → for title, agenda, key takeaways, or closing slides

## PRIVACY / IDENTIFIER RULES — CRITICAL:

- Some categorical columns in the summary may be flagged with
  "identifier_like": true. These are high-cardinality fields that likely
  contain personally identifiable data (client names, order IDs, emails, etc.).
- NEVER reference, quote, or use values from identifier-like columns in
  insight_text. Do not name specific clients, people, or identifiers.
- NEVER include identifier-like columns in data_columns for charting.
- If all columns in a sheet are identifier-like, skip that sheet entirely.

## SLIDE PLANNING GUIDELINES:

- Start with a title/overview slide (chart_type: "text").
- End with a key takeaways / summary slide (chart_type: "text").
- Distribute the remaining slides across the available sheets based on
  the user's stated focus areas.
- Write insight_text as if you are presenting to the stated audience.
  Be specific with numbers from the summary (e.g. "Revenue totalled $6.2M
  across 30 transactions, with North region contributing the highest share").
- Use the actual column names and sheet names from the summary — do NOT
  invent columns that don't exist.
- Produce exactly the number of slides requested.
"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _condense_summary(summary: dict) -> dict:
    """
    Produce a leaner version of the data summary to fit within free-tier
    token limits. Strips redundant per-column fields (count, non_null)
    that the LLM doesn't need for slide planning.
    """
    condensed = {}
    for sheet_name, sheet_info in summary.items():
        entry = {
            "sheet_name": sheet_info.get("sheet_name", sheet_name),
            "row_count": sheet_info.get("row_count", 0),
            "columns": {},
        }
        if "note" in sheet_info:
            entry["note"] = sheet_info["note"]

        for col_name, col_info in sheet_info.get("columns", {}).items():
            # Keep only the fields the LLM actually needs
            slim = {"type": col_info.get("inferred_type", "unknown")}

            # Copy key stats, skip verbose ones
            for key in ("sum", "mean", "min", "max", "median",
                        "min_date", "max_date", "date_range_days",
                        "unique_values", "top_values",
                        "identifier_like", "note",
                        "true_pct", "true_count", "false_count"):
                if key in col_info:
                    slim[key] = col_info[key]

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

def generate_slide_plan(
    data_summary: dict,
    user_prompt: str,
    num_slides: int = 8,
    model: str = DEFAULT_MODEL,
    api_key: Optional[str] = None,
) -> dict:
    """
    Call the Groq API to generate a structured slide plan.

    Parameters
    ----------
    data_summary : dict
        The aggregate-only summary from data_parser.parse_excel().
    user_prompt : str
        The user's instructions (focus areas, audience, tone, etc.).
    num_slides : int
        Number of slides to generate (default 8).
    model : str
        Groq model ID (default: llama-3.1-8b-instant).
    api_key : str | None
        Groq API key. Falls back to GROQ_API_KEY env var / .env file.

    Returns
    -------
    dict with key "slides" containing the list of slide plan dicts.
    """
    # Resolve API key
    key = api_key or os.environ.get("GROQ_API_KEY")
    if not key:
        raise RuntimeError(
            "No Groq API key found. Set GROQ_API_KEY in your .env file "
            "or pass api_key= directly."
        )

    client = Groq(api_key=key)

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
            # Add retry hint to user message if this is a retry
            msg = user_message
            if attempt > 0:
                msg += (
                    "\n\nWARNING: YOUR PREVIOUS RESPONSE WAS NOT VALID JSON. "
                    "Respond with ONLY a raw JSON array — no markdown fences, "
                    "no explanation, no preamble. Just the [ ... ] array."
                )

            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": msg},
                ],
                temperature=0.4,
                max_tokens=4096,
            )

            raw_text = response.choices[0].message.content
            slides = _parse_slide_plan(raw_text)

            return {
                "slides": slides,
                "model": model,
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
            raise RuntimeError(f"Groq API call failed: {e}") from e


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
