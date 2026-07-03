# DeckMate — 7 Day Build Roadmap (2 hrs/day, vibecoded)

Each day = one working, testable chunk. Paste the prompt into Claude Code / Cursor / whatever you're vibecoding with, for that day's file(s) only — don't jump ahead, keep each day's output actually running before moving on.

Assumes: Python, Streamlit, pandas, python-pptx, Groq API key ready.

---

## Day 1 — Excel parsing + data summarization
**Goal:** Given any multi-tab .xlsx, output a clean aggregated summary (no raw rows) that's safe to send to an LLM.

**Prompt:**
> Write a Python module `data_parser.py` that takes an uploaded .xlsx file (multi-tab) and returns a JSON-safe summary dict, one entry per sheet. For each sheet, use pandas to detect column types, and compute: row count, column names + inferred types, for numeric columns give sum/mean/min/max, for categorical columns give top 5 value counts, and for any date-like column give min/max date range. Do NOT include any raw row-level data in the output — only aggregates. Handle sheets with messy headers (e.g. header not in row 1) gracefully. Include a `if __name__ == "__main__"` test block that runs it on a sample multi-tab Excel file and pretty-prints the JSON summary.

**Test:** Feed it a real (or dummy) multi-tab sales sheet, confirm the summary is genuinely aggregate-only — eyeball it yourself to make sure no client names/raw rows leaked through.

---

## Day 2 — Groq API integration + structured JSON slide plan
**Goal:** Send the Day 1 summary + a user prompt to Groq, get back a structured slide plan.

**Prompt:**
> Write a Python module `llm_planner.py` that calls the Groq API (use `groq` python SDK, model llama-3.1-8b-instant or similar). Function `generate_slide_plan(data_summary: dict, user_prompt: str, num_slides: int) -> dict` should send a system prompt instructing the model to act as a presentation planner for a sales deck, given ONLY the aggregated data summary (never raw data) and the user's focus/audience instructions. The model must respond with ONLY valid JSON (no markdown fences, no preamble) matching this schema: a list of slides, each with `title`, `source_sheet`, `chart_type` (one of bar/pie/line/table/text), `data_columns` (which columns from that sheet to chart), and `insight_text` (2-3 sentence written interpretation). Include JSON parsing with error handling that strips markdown fences if the model adds them anyway, and retries once if parsing fails.

**Test:** Run it with Day 1's summary output + a sample prompt like "focus on regional performance, for senior leadership, 5 slides." Confirm valid JSON comes back.

---

## Day 3 — PPTX builder with real charts
**Goal:** Take the JSON slide plan + the ORIGINAL dataframe (not LLM numbers) and build native PowerPoint charts.

**Prompt:**
> Write a Python module `ppt_builder.py` using python-pptx. Function `build_presentation(slide_plan: dict, dataframes: dict, output_path: str)` should loop through each slide in the plan, and for chart slides, pull the real chart data by recomputing it directly from the pandas dataframe specified in `source_sheet` and `data_columns` (never trust numbers from the LLM). Use python-pptx's native `add_chart()` with `XL_CHART_TYPE.PIE`, `COLUMN_CLUSTERED`, or `LINE` matching `chart_type`. Add a title text box and the `insight_text` as a text box below/beside the chart. Use a clean, consistent layout (e.g. title top, chart left, insight text right, or chart top-insight bottom — your call). Save to output_path as .pptx.

**Test:** Manually construct a fake slide_plan + dataframe dict, run it, open the resulting .pptx and confirm charts render correctly with real numbers.

---

## Day 4 — Streamlit frontend
**Goal:** Upload UI + prompt box + config, wired to nothing yet (just UI + state).

**Prompt:**
> Write a Streamlit app `app.py` with: a file uploader restricted to .xlsx, a text area for the user's prompt (focus areas, audience), a number input for slide count (default 8), and a "Generate Deck" button. On button click, just show a spinner and st.write the uploaded filename + prompt for now as a placeholder — don't wire the pipeline yet. Add a sidebar note about the privacy approach: "Your file is processed in-memory and never stored. Only aggregated summaries are sent to the AI model — never raw row-level data."

**Test:** `streamlit run app.py` locally, upload a file, confirm UI flow feels smooth.

---

## Day 5 — Wire the full pipeline + enforce no-persistence
**Goal:** Connect Day 1 → 2 → 3 → 4 into one working end-to-end flow, and make the "we don't save your data" claim actually true in code.

**Prompt:**
> In `app.py`, wire the full pipeline: on "Generate Deck" click, read the uploaded file directly into memory (use `io.BytesIO`, never write the uploaded file to disk), pass it through `data_parser.py` to get the summary, pass the summary + user prompt into `llm_planner.py` to get the slide plan, then pass the plan + in-memory dataframes into `ppt_builder.py`, writing the output pptx to a temp file using Python's `tempfile` module. Offer it via `st.download_button` reading the temp file's bytes into memory, then explicitly delete the temp file right after (use `os.remove` in a `finally` block, or `tempfile.NamedTemporaryFile` with `delete=True` if compatible with python-pptx's save method). Add error handling for corrupt Excel files and LLM JSON failures, showing a friendly Streamlit error message instead of crashing.

**Test:** Run the full flow with a real sheet start to finish, confirm a working .pptx downloads, and confirm (by checking the temp dir) that nothing lingers after the request completes.

---

## Day 6 — Polish, edge cases, branding
**Goal:** Make it demo-ready and resilient.

**Prompt:**
> Improve `deckmate` for robustness and polish: (1) in `data_parser.py`, handle edge cases — empty sheets, sheets with all-text columns, sheets with >50% missing values — without crashing. (2) In `ppt_builder.py`, add a simple consistent color theme/template (define brand colors as constants) and a title slide + closing slide. (3) In `app.py`, add a progress indicator with stage labels ("Reading your data...", "Planning your deck...", "Building slides...") instead of a single spinner, using `st.status` or sequential `st.spinner` blocks. (4) Add input validation — reject files >10MB, show clear error if no numeric data found at all.

**Test:** Try deliberately messy/edge-case Excel files, confirm graceful failures with clear messages, not stack traces.

---

## Day 7 — Deploy + final README pass
**Goal:** Live, shareable link for both your relative and your portfolio.

**Prompt:**
> Help me prepare this Streamlit app for deployment to Streamlit Community Cloud: generate a `requirements.txt` with pinned versions for streamlit, pandas, openpyxl, python-pptx, groq. Add a `.gitignore` for typical Python/Streamlit projects plus `.env`. Walk me through setting the GROQ_API_KEY as a Streamlit secret (not committed to git) using `st.secrets`, and update `llm_planner.py` to read the key from `st.secrets["GROQ_API_KEY"]` with a fallback to `os.environ` for local dev.

**Test:** Deploy, open the live URL on your phone, run one real end-to-end generation, confirm it works exactly like local.

---

## After Day 7 (optional, if you have more time)
- Add a "regenerate this slide only" button (re-run just one slide's LLM call + rebuild)
- Let user pick a chart theme/color palette
- Add basic auth so it's not a fully public link if sensitive data will flow through
- Swap Groq model for a bigger one temporarily if quality needs a boost, compare outputs
