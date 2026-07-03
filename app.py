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

# ---------------------------------------------------------------------------
# Custom CSS for a premium look
# ---------------------------------------------------------------------------

st.markdown("""
<style>
    /* Main container */
    .main .block-container {
        padding-top: 2rem;
        max-width: 900px;
    }

    /* Header styling */
    .hero-title {
        font-size: 2.4rem;
        font-weight: 800;
        color: #1B2A4A;
        margin-bottom: 0.2rem;
        letter-spacing: -0.5px;
    }
    .hero-subtitle {
        font-size: 1.1rem;
        color: #666;
        margin-bottom: 2rem;
        line-height: 1.6;
    }

    /* Card-like sections */
    .config-section {
        background: #f8f9fa;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        border: 1px solid #e9ecef;
    }
    .config-section h3 {
        color: #1B2A4A;
        font-size: 1rem;
        font-weight: 600;
        margin-bottom: 0.8rem;
    }

    /* Upload area */
    [data-testid="stFileUploader"] {
        border-radius: 12px;
    }

    /* Generate button */
    .stButton > button {
        background: linear-gradient(135deg, #2E86AB 0%, #1B2A4A 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-size: 1.1rem;
        font-weight: 600;
        border-radius: 8px;
        width: 100%;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 15px rgba(46, 134, 171, 0.4);
    }
    .stButton > button:active {
        transform: translateY(0);
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #f8f9fa;
    }

    /* Privacy badge */
    .privacy-badge {
        background: linear-gradient(135deg, #e8f5e9 0%, #f1f8e9 100%);
        border: 1px solid #c8e6c9;
        border-radius: 10px;
        padding: 1rem;
        margin-top: 1rem;
    }
    .privacy-badge h4 {
        color: #2e7d32;
        font-size: 0.9rem;
        margin-bottom: 0.5rem;
    }
    .privacy-badge p {
        color: #555;
        font-size: 0.8rem;
        line-height: 1.5;
        margin: 0;
    }

    /* Pipeline steps */
    .pipeline-step {
        display: flex;
        align-items: center;
        padding: 0.4rem 0;
        color: #666;
        font-size: 0.85rem;
    }
    .pipeline-step .step-num {
        background: #2E86AB;
        color: white;
        width: 22px;
        height: 22px;
        border-radius: 50%;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 0.7rem;
        font-weight: 700;
        margin-right: 0.6rem;
        flex-shrink: 0;
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
        with st.status("Generating your deck...", expanded=True) as status:
            st.write("📖 Reading your data...")
            file_bytes = io.BytesIO(uploaded_file.getvalue())
            try:
                summary, dataframes = parse_excel(file_bytes)
            except ValueError as e:
                st.error(f"Could not read this Excel file: {e}")
                st.stop()

            sheet_names = [s for s, info in summary.items()
                          if info.get("row_count", 0) > 0]
            if not sheet_names:
                st.error(
                    "No usable data found in this file. "
                    "All sheets appear to be empty."
                )
                st.stop()

            st.write(
                f"Found **{len(sheet_names)}** sheets with data: "
                f"{', '.join(sheet_names)}"
            )

            # Stage 2: LLM Planning
            st.write("🧠 Planning your deck...")
            try:
                slide_plan = generate_slide_plan(
                    data_summary=summary,
                    user_prompt=final_prompt,
                    num_slides=num_slides,
                )
            except ValueError as e:
                st.error(
                    f"The AI returned an invalid response. "
                    f"Please try again — this usually works on retry.\n\n"
                    f"Details: {e}"
                )
                st.stop()
            except RuntimeError as e:
                st.error(f"AI service error: {e}")
                st.stop()

            n_slides = slide_plan.get("num_slides_returned", 0)
            st.write(f"Planned **{n_slides}** slides")

            # Stage 3: Build PPTX
            st.write("📊 Building slides & charts...")

            # Create temp file — use delete=False so we can read it back
            temp_fd, temp_path = tempfile.mkstemp(suffix=".pptx")
            os.close(temp_fd)  # close the file descriptor, ppt_builder will write to path

            try:
                build_presentation(slide_plan, dataframes, temp_path)
            except Exception as e:
                st.error(f"Error building the presentation: {e}")
                st.stop()

            # Read the generated file into memory immediately
            with open(temp_path, "rb") as f:
                pptx_bytes = f.read()

            status.update(label="Deck ready!", state="complete", expanded=True)
            st.write("✅ Done!")

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

