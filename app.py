"""
app.py — Streamlit frontend for DeckMate (Day 5)

Full pipeline wired: Upload → data_parser → llm_planner → ppt_builder → download.
All processing is in-memory. Temp files are deleted immediately after serving.
"""

import io
import os
import tempfile

import streamlit as st

from data_parser import parse_excel
from llm_planner import generate_slide_plan
from ppt_builder import build_presentation

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="DeckMate — AI-Powered Excel → PowerPoint",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------# Custom CSS for a clean, beautiful premium look (White Mode)
# ---------------------------------------------------------------------------

st.markdown("""
<style>
    /* Global Background and Typography */
    .stApp {
        background-color: #ffffff;
        color: #333333;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    
    /* Main container */
    .main .block-container {
        padding-top: 3rem;
        max-width: 960px;
    }

    /* Header styling */
    .hero-title {
        font-size: 2.8rem;
        font-weight: 800;
        color: #1B2A4A;
        margin-bottom: 0.1rem;
        letter-spacing: -0.8px;
    }
    .hero-subtitle {
        font-size: 1.15rem;
        color: #555555;
        margin-bottom: 2.5rem;
        line-height: 1.6;
    }

    /* Clean Card Sections */
    .config-section {
        background: #ffffff;
        border-radius: 12px;
        padding: 1.75rem;
        margin-bottom: 1.5rem;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.02);
    }
    .config-section h3 {
        color: #1B2A4A;
        font-size: 1.1rem;
        font-weight: 700;
        margin-bottom: 0.8rem;
    }

    /* Upload area styling */
    [data-testid="stFileUploader"] {
        border: 1px dashed #cbd5e1;
        border-radius: 12px;
        padding: 1rem;
        background-color: #f8fafc;
    }

    /* Primary Generate button with Teal-to-Navy gradient */
    .stButton > button {
        background: linear-gradient(135deg, #2E86AB 0%, #1B2A4A 100%);
        color: white;
        border: none;
        padding: 0.85rem 2.2rem;
        font-size: 1.15rem;
        font-weight: 600;
        border-radius: 8px;
        width: 100%;
        box-shadow: 0 10px 15px -3px rgba(46, 134, 171, 0.25);
        transition: all 0.2s ease-in-out;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 20px -2px rgba(46, 134, 171, 0.35);
        color: #ffffff !important;
    }
    .stButton > button:active {
        transform: translateY(0);
    }

    /* Sidebar (White Mode styling) */
    [data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
    [data-testid="stSidebar"] .block-container {
        padding-top: 2rem;
    }

    /* Clean white privacy badge */
    .privacy-badge {
        background-color: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 10px;
        padding: 1.25rem;
        margin-top: 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }
    .privacy-badge h4 {
        color: #166534;
        font-size: 0.95rem;
        font-weight: 700;
        margin-top: 0;
        margin-bottom: 0.4rem;
    }
    .privacy-badge p {
        color: #1e3a1e;
        font-size: 0.85rem;
        line-height: 1.5;
        margin: 0;
    }

    /* Pipeline steps indicator */
    .pipeline-step {
        display: flex;
        align-items: center;
        padding: 0.5rem 0;
        color: #475569;
        font-size: 0.9rem;
        font-weight: 500;
    }
    .pipeline-step .step-num {
        background: #2E86AB;
        color: white;
        width: 24px;
        height: 24px;
        border-radius: 50%;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 0.75rem;
        font-weight: 700;
        margin-right: 0.75rem;
        box-shadow: 0 2px 4px rgba(46, 134, 171, 0.2);
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown("### DeckMate")
    st.markdown("*AI-Powered Deck Generator*")

    st.markdown("---")

    st.markdown("#### How it works")
    st.markdown("""
    <div class="pipeline-step">
        <span class="step-num">1</span>
        Upload your Excel file (.xlsx)
    </div>
    <div class="pipeline-step">
        <span class="step-num">2</span>
        Describe your audience & focus
    </div>
    <div class="pipeline-step">
        <span class="step-num">3</span>
        AI plans your slide structure
    </div>
    <div class="pipeline-step">
        <span class="step-num">4</span>
        Real charts built from your data
    </div>
    <div class="pipeline-step">
        <span class="step-num">5</span>
        Download your .pptx deck
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    st.markdown("""
    <div class="privacy-badge">
        <h4>🔒 Privacy First</h4>
        <p>
            Your file is processed <strong>in-memory</strong> and never stored.
            Only aggregated summaries are sent to the AI model — never raw
            row-level data. No files, sessions, or conversation history
            are persisted.
        </p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.caption("Built with Streamlit + Groq + python-pptx")


# ---------------------------------------------------------------------------
# Main content
# ---------------------------------------------------------------------------

st.markdown('<p class="hero-title">DeckMate</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="hero-subtitle">'
    'Turn messy, multi-tab Excel sheets into polished, presentation-ready '
    'PowerPoint decks — with real charts, AI-written insights, and '
    'audience-aware framing.'
    '</p>',
    unsafe_allow_html=True,
)

# ---- File upload ----
st.markdown("#### 📂 Upload your Excel file")
uploaded_file = st.file_uploader(
    "Drag and drop or browse for a .xlsx file",
    type=["xlsx"],
    help="Multi-tab workbooks work best. Max recommended size: 10 MB.",
    label_visibility="collapsed",
)

if uploaded_file:
    file_size_mb = uploaded_file.size / (1024 * 1024)
    st.success(
        f"**{uploaded_file.name}** uploaded "
        f"({file_size_mb:.1f} MB)"
    )

st.markdown("---")

# ---- Prompt input ----
st.markdown("#### 💬 Describe your deck")

col1, col2 = st.columns([3, 1])

with col1:
    user_prompt = st.text_area(
        "What should this deck focus on? Who is the audience?",
        placeholder=(
            "Example: Focus on regional sales performance and marketing ROI. "
            "This is for a senior leadership quarterly review. "
            "Use professional, concise language. "
            "Highlight any standout trends or areas of concern."
        ),
        height=120,
        help="Be specific about focus areas, audience, and tone. "
             "The AI uses this to decide slide structure and writing style.",
    )

with col2:
    num_slides = st.number_input(
        "Number of slides",
        min_value=3,
        max_value=20,
        value=8,
        step=1,
        help="Includes title and summary slides.",
    )

    audience_preset = st.selectbox(
        "Quick preset",
        options=["Custom (use text above)", "Senior Leadership", "Client-Facing",
                 "Internal Team", "Board Presentation"],
        help="Optional: auto-adjusts tone if no custom prompt is given.",
    )

st.markdown("---")

# ---- Generate button ----
generate_disabled = uploaded_file is None

if generate_disabled:
    st.info("Upload an Excel file to get started.")

generate_clicked = st.button(
    "Generate Deck",
    disabled=generate_disabled,
    use_container_width=True,
    type="primary",
)

# ---- Full pipeline (Day 5) ----
if generate_clicked and uploaded_file:

    # --- Clamping slide count in code ---
    num_slides = max(3, min(15, num_slides))

    # --- File size validation ---
    file_size_mb = uploaded_file.size / (1024 * 1024)
    if file_size_mb > 10:
        st.error(
            f"File is {file_size_mb:.1f} MB — please keep uploads under 10 MB "
            f"to avoid memory issues."
        )
        st.stop()

    # --- Build combined prompt (preset + custom text) ---
    custom_text = user_prompt.strip()
    if audience_preset != "Custom (use text above)":
        if custom_text:
            final_prompt = f"Audience: {audience_preset}. {custom_text}"
        else:
            final_prompt = (
                f"Audience: {audience_preset}. "
                f"Create a professional deck appropriate for this audience."
            )
    else:
        final_prompt = custom_text or "Create a professional overview deck."

    # --- Pipeline execution with stage indicators ---
    temp_path = None

    try:
        # Stage 1: Parse Excel
        with st.status("Reading your data...", expanded=True) as status:
            file_bytes = io.BytesIO(uploaded_file.getvalue())
            try:
                summary, dataframes = parse_excel(file_bytes)
            except Exception as e:
                status.update(label="Failed to read file", state="error")
                st.error(f"Could not read this Excel file: {e}")
                st.stop()

            # --- Structural data validation ---
            sheet_names = [s for s, info in summary.items() if info.get("row_count", 0) > 0]
            if not sheet_names:
                status.update(label="Empty Excel File", state="error")
                st.error("No usable data found in this file. All sheets appear to be empty.")
                st.stop()

            # Check if there is at least one usable numeric or categorical column across all sheets
            total_columns = 0
            for s in sheet_names:
                cols_dict = summary[s].get("columns", {})
                total_columns += len(cols_dict)
            
            if total_columns == 0:
                status.update(label="No columns to chart", state="error")
                st.error("The Excel file doesn't contain any columns or headers. Nothing to chart!")
                st.stop()

            status.update(label="Understanding your data structure...")
            st.write(f"✓ Found **{len(sheet_names)}** sheets containing **{total_columns}** columns.")

            # Stage 2: LLM Planning
            status.update(label="Planning your deck with AI...")
            try:
                # Provide a suggested deck title to pass to ppt_builder
                deck_title = f"{os.path.splitext(uploaded_file.name)[0].replace('_', ' ').title()} Deck"
                slide_plan = generate_slide_plan(
                    data_summary=summary,
                    user_prompt=final_prompt,
                    num_slides=num_slides,
                )
                slide_plan["title_slide_title"] = deck_title
            except ValueError as e:
                status.update(label="AI Planning Failed", state="error")
                st.error(
                    f"The AI returned an invalid response. "
                    f"Please try again — this usually works on retry.\n\n"
                    f"Details: {e}"
                )
                st.stop()
            except RuntimeError as e:
                status.update(label="AI Service Error", state="error")
                st.error(f"AI service error: {e}")
                st.stop()

            n_slides = slide_plan.get("num_slides_returned", 0)
            st.write(f"✓ AI slide plan complete ({n_slides} content slides planned).")

            # Stage 3: Build PPTX
            status.update(label="Building slides...")
            # Create temp file — use delete=False so we can read it back
            temp_fd, temp_path = tempfile.mkstemp(suffix=".pptx")
            os.close(temp_fd)  # close the file descriptor, ppt_builder will write to path

            try:
                build_presentation(slide_plan, dataframes, temp_path)
            except Exception as e:
                status.update(label="Slide building failed", state="error")
                st.error(f"Error building the presentation: {e}")
                st.stop()

            status.update(label="Finalizing...")
            # Read the generated file into memory immediately
            with open(temp_path, "rb") as f:
                pptx_bytes = f.read()

            status.update(label="Deck ready!", state="complete", expanded=False)

        # --- Download button ---
        st.markdown("---")
        st.markdown("### Your deck is ready!")

        # Build a clean filename from the uploaded file
        base_name = os.path.splitext(uploaded_file.name)[0]
        download_name = f"{base_name}_deck.pptx"

        st.download_button(
            label=f"Download {download_name}",
            data=pptx_bytes,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            type="primary",
        )

        # --- Show slide plan summary ---
        with st.expander(f"Slide plan ({n_slides} slides)", expanded=False):
            for i, slide in enumerate(slide_plan.get("slides", [])):
                st.markdown(
                    f"**Slide {i+1}:** {slide.get('title', 'Untitled')}  \n"
                    f"*Type:* `{slide.get('chart_type', 'text')}` · "
                    f"*Sheet:* `{slide.get('source_sheet', 'none')}` · "
                    f"*Columns:* `{slide.get('data_columns', [])}`  \n"
                    f"_{slide.get('insight_text', '')}_"
                )
                if i < n_slides - 1:
                    st.markdown("---")

    finally:
        # --- Temp file cleanup (privacy-critical) ---
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass  # best-effort deletion

            # Verify deletion
            if os.path.exists(temp_path):
                st.warning(
                    "Warning: temporary file could not be deleted. "
                    "Please manually remove: " + temp_path
                )

