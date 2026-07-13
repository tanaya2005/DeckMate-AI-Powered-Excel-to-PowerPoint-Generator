# DeckMate — AI-Powered Excel → PowerPoint Generator

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://deckmate.streamlit.app/)

Turns messy, multi-tab sales Excel sheets into a polished, presentation-ready PowerPoint deck — complete with real charts (bar/pie/line), auto-written interpretations, and audience-aware framing (client-facing vs internal/senior-facing).

Built for a real use case: a sales professional who needs to turn long raw Excel data into client/senior-ready decks fast, without manually building charts and writing insights every time.

---

## How it works

```
Excel (.xlsx, multi-tab)
        │
        ▼
[1] pandas/openpyxl → parses every tab, computes aggregate stats
        │   (sums, trends, top-N, groupbys — NEVER raw rows)
        ▼
[2] Data summary + user prompt → Groq LLM
        │   (LLM only ever sees aggregated stats, decides:
        │    which slides, which chart type, what to write)
        ▼
[3] LLM returns structured JSON slide plan
        │
        ▼
[4] python-pptx → builds native, editable charts from the
        │   REAL dataframe (not LLM-guessed numbers)
        ▼
Downloadable .pptx
```

**Core design rule:** the LLM never does math and never sees raw row-level data. It only plans structure and writes narrative text based on pre-computed summaries. All numbers on every chart come straight from pandas.

---

## Tech stack (100% free tier)

| Piece | Tool | Why |
|---|---|---|
| Data parsing | `pandas` + `openpyxl` | Free, handles multi-tab Excel well |
| LLM (planning + text) | Groq API (Llama 3.1 / similar) | Free tier, does **not** train on your data — critical for client privacy |
| Slide generation | `python-pptx` | Native editable charts, not static images |
| Frontend | Streamlit | Fast to build, upload + prompt UI in ~50 lines |
| Hosting | Streamlit Community Cloud / Hugging Face Spaces | Free tier deployment |

---

## Privacy & data handling — read this before deploying

This app is built to be **stateless by design**:

- Uploaded Excel files are processed **in-memory / temp storage only**, never written to persistent disk or a database.
- Temp files (if any are created for library compatibility) are deleted immediately after the response is generated.
- The LLM (Groq) only ever receives **aggregated summaries** of the data (e.g. "Q3 region-wise revenue: +12%"), never raw rows — so client names, deal values, contact info, etc. never leave the user's session in identifiable form.
- No user Excel sheets or generated PPTs are logged, cached, or stored server-side.
- No conversation/session history is persisted between uses.

If you ever add a "save my past decks" feature later, that becomes an explicit opt-in with its own storage and disclosure — not a default.

---

## Project structure

```
deckmate/
├── app.py                  # Streamlit entrypoint (upload + prompt UI)
├── data_parser.py          # Excel → aggregated summary (pandas)
├── llm_planner.py          # Groq API call → structured JSON slide plan
├── ppt_builder.py          # JSON plan + real data → .pptx (python-pptx)
├── requirements.txt
└── README.md
```

---

## Setup

```bash
git clone <your-repo-url>
cd deckmate
pip install -r requirements.txt
```

Create a `.env` file:
```
GROQ_API_KEY=your_key_here
```

Run locally:
```bash
streamlit run app.py

OR 


python -m streamlit run app.py
