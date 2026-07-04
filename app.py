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
# ---------------------------------------------------------------------------
# Custom CSS matching Stitch Design ("DeckMate Generator")
# ---------------------------------------------------------------------------

st.markdown("""
<style>
    /* Google Fonts & Material Icons */
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    @import url('https://fonts.googleapis.com/icon?family=Material+Icons+Outlined');

    /* Global Typography & Palette (Theme Independent overrides) */
    :root {
        --primary-color: #2563eb !important;
        --background-color: #f7f9fb !important;
        --secondary-background-color: #f2f4f6 !important;
        --text-color: #0f172a !important;
    }

    html, body, [data-testid="stAppViewContainer"], .stApp {
        background-color: #f7f9fb !important;
        color: #0f172a !important;
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }

    /* Override Streamlit base light-dark mode variable controls for borders & texts */
    div, span, p, label, select, input, textarea, button {
        border-color: #cbd5e1 !important;
    }

    /* Main container bounds */
    .main .block-container {
        padding-top: 2rem !important;
        padding-bottom: 4rem !important;
        max-width: 1000px !important;
    }

    /* Title & Welcome Hero styling */
    .welcome-badge {
        color: #2563eb !important;
        font-weight: 700;
        font-size: 0.875rem;
        margin-bottom: 0.5rem;
    }
    .welcome-title {
        font-size: 2.25rem;
        font-weight: 800;
        color: #0f172a !important;
        margin-bottom: 0.75rem;
        letter-spacing: -0.5px;
        line-height: 1.25;
    }
    .welcome-desc {
        font-size: 1rem;
        color: #64748b !important;
        line-height: 1.6;
        margin-bottom: 2rem;
    }

    /* Force controls to render high-contrast white mode regardless of system preferences */
    [data-testid="stTextarea"] textarea,
    [data-testid="stSelectbox"] select,
    [data-testid="stNumberInput"] input,
    [data-testid="stTextInput"] input,
    [data-testid="stFileUploader"],
    [data-testid="stFileUploader"] > div,
    [data-testid="stFileUploaderDropzone"],
    .stTextInput input,
    select,
    textarea,
    input {
        color: #0f172a !important;
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
        font-size: 0.875rem !important;
    }

    [data-testid="stTextarea"] textarea::placeholder,
    textarea::placeholder {
        color: #94a3b8 !important;
        opacity: 1 !important;
    }

    label[data-testid="stWidgetLabel"],
    [data-testid="stWidgetLabel"] p,
    label {
        color: #0f172a !important;
        font-weight: 700 !important;
        font-size: 0.75rem !important;
    }

    /* Clean Card Sections (Stitch style) */
    .stitch-card {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 16px !important;
        padding: 1.5rem !important;
        margin-bottom: 1rem !important;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05) !important;
    }
    
    .card-header {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        margin-bottom: 0.5rem;
    }
    .card-icon-box-green {
        background-color: #f0fdf4;
        padding: 0.5rem;
        border-radius: 8px;
        color: #166534;
        display: inline-flex;
    }
    .card-icon-box-blue {
        background-color: #eff6ff;
        padding: 0.5rem;
        border-radius: 8px;
        color: #2563eb;
        display: inline-flex;
    }
    .card-title-text {
        font-weight: 700;
        color: #0f172a !important;
        font-size: 1.05rem;
    }
    .card-subtitle-text {
        font-size: 0.75rem;
        color: #94a3b8 !important;
        margin-top: 0.1rem;
    }

    /* Dropzone pattern styling */
    [data-testid="stFileUploader"] {
        border: 2px dashed #cbd5e1 !important;
        border-radius: 12px !important;
        padding: 1.5rem !important;
        background-color: #fafbfc !important;
        text-align: center;
    }
    [data-testid="stFileUploader"] label {
        color: #0f172a !important;
        font-weight: 600 !important;
    }

    /* Blueprint Info Banner */
    .blueprint-info-banner {
        background-color: #eff6ff !important;
        border: 1px solid #dbeafe !important;
        border-radius: 12px !important;
        padding: 1rem !important;
        display: flex;
        align-items: center;
        gap: 0.75rem;
        margin-bottom: 1rem;
        color: #1e40af !important;
        font-weight: 600;
        font-size: 0.875rem;
    }
    .blueprint-icon-circle {
        background-color: #2563eb;
        color: #ffffff;
        width: 24px;
        height: 24px;
        border-radius: 50%;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 14px;
        font-weight: bold;
    }

    /* Primary Generate button shadow and active effect */
    .stButton > button {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border: none !important;
        padding: 0.9rem 2rem !important;
        font-size: 1rem !important;
        font-weight: 700 !important;
        border-radius: 12px !important;
        width: 100% !important;
        box-shadow: 0 10px 15px -3px rgba(37, 99, 235, 0.2) !important;
        transition: all 0.15s ease-in-out !important;
    }
    .stButton > button:hover {
        background-color: #1d4ed8 !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 12px 20px -2px rgba(37, 99, 235, 0.3) !important;
        color: #ffffff !important;
    }
    .stButton > button:active {
        transform: scale(0.99) !important;
    }

    /* Sidebar Navigation Steps */
    [data-testid="stSidebar"],
    [data-testid="stSidebar"] > div {
        background-color: #f2f4f6 !important;
        border-right: 1px solid #e2e8f0 !important;
    }
    [data-testid="stSidebar"] .block-container {
        padding-top: 2rem;
    }
    
    .sidebar-step-item {
        display: flex;
        gap: 0.75rem;
        align-items: flex-start;
        margin-bottom: 1rem;
    }
    .sidebar-step-num-active {
        background-color: #2563eb !important;
        color: #ffffff !important;
        width: 24px;
        height: 24px;
        border-radius: 50%;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 12px;
        font-weight: 700;
    }
    .sidebar-step-num-inactive {
        background-color: #e2e8f0 !important;
        color: #64748b !important;
        width: 24px;
        height: 24px;
        border-radius: 50%;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 12px;
        font-weight: 700;
    }
    .sidebar-step-icon-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 0.4rem;
        display: inline-flex;
        color: #64748b;
    }
    .sidebar-step-icon-card-active {
        background-color: #ffffff;
        border: 1px solid #dbeafe;
        border-radius: 8px;
        padding: 0.4rem;
        display: inline-flex;
        color: #2563eb;
    }
    .sidebar-step-text-active {
        font-size: 0.875rem;
        font-weight: 700;
        color: #2563eb !important;
        margin-top: 0.15rem;
    }
    .sidebar-step-text-inactive {
        font-size: 0.875rem;
        font-weight: 600;
        color: #475569 !important;
        margin-top: 0.15rem;
    }

    /* Privacy Banner Box */
    .sidebar-privacy-box {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 12px !important;
        padding: 1rem !important;
        margin-top: 2rem !important;
    }
    .sidebar-privacy-header {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        color: #2563eb !important;
        font-weight: 700;
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 0.5rem;
    }
    .sidebar-privacy-desc {
        font-size: 0.725rem;
        color: #64748b !important;
        line-height: 1.5;
        margin: 0;
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Sidebar (Stitch UI layout conversion)
# ---------------------------------------------------------------------------

with st.sidebar:
    # Sidebar Header
    st.markdown("""
    <div style="display: flex; align-items: center; gap: 0.75rem; margin-bottom: 0.25rem;">
        <div style="background-color: #2563eb; padding: 0.4rem; border-radius: 8px; display: inline-flex; color: white;">
            <span class="material-icons-outlined" style="font-size: 1.25rem;">description</span>
        </div>
        <h2 style="font-size: 1.25rem; font-weight: 800; color: #0f172a; margin: 0; tracking-tight: -0.5px;">DeckMate</h2>
    </div>
    <p style="font-size: 0.75rem; color: #64748b; font-weight: 500; margin-top: 0; margin-left: 2.6rem;">AI-Powered Deck Generator</p>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # Step list indicators based on design system
    st.markdown("""
    <p style="font-size: 10px; font-weight: 800; color: #64748b; uppercase: true; letter-spacing: 1px; margin-bottom: 1rem;">HOW IT WORKS</p>
    
    <div class="sidebar-step-item">
        <div class="sidebar-step-num-active">1</div>
        <div class="sidebar-step-icon-card-active">
            <span class="material-icons-outlined" style="font-size: 1.1rem;">upload_file</span>
        </div>
        <span class="sidebar-step-text-active">Upload your Excel file</span>
    </div>
    
    <div class="sidebar-step-item">
        <div class="sidebar-step-num-inactive">2</div>
        <div class="sidebar-step-icon-card">
            <span class="material-icons-outlined" style="font-size: 1.1rem;">groups</span>
        </div>
        <span class="sidebar-step-text-inactive">Describe your audience & focus</span>
    </div>
    
    <div class="sidebar-step-item">
        <div class="sidebar-step-num-inactive">3</div>
        <div class="sidebar-step-icon-card">
            <span class="material-icons-outlined" style="font-size: 1.1rem;">auto_awesome</span>
        </div>
        <span class="sidebar-step-text-inactive">AI plans your slide structure</span>
    </div>
    
    <div class="sidebar-step-item">
        <div class="sidebar-step-num-inactive">4</div>
        <div class="sidebar-step-icon-card">
            <span class="material-icons-outlined" style="font-size: 1.1rem;">bar_chart</span>
        </div>
        <span class="sidebar-step-text-inactive">Real charts built from data</span>
    </div>
    
    <div class="sidebar-step-item">
        <div class="sidebar-step-num-inactive">5</div>
        <div class="sidebar-step-icon-card">
            <span class="material-icons-outlined" style="font-size: 1.1rem;">download</span>
        </div>
        <span class="sidebar-step-text-inactive">Download your .pptx deck</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # Clean privacy box
    st.markdown("""
    <div class="sidebar-privacy-box">
        <div class="sidebar-privacy-header">
            <span class="material-icons-outlined" style="font-size: 1rem;">verified_user</span>
            <span>Privacy First</span>
        </div>
        <p class="sidebar-privacy-desc">
            Your file is processed in-memory and never stored. Only aggregated summaries are sent to the AI model — never raw data. No files, sessions, or conversation history are persisted.
        </p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.caption("Built with Streamlit + Groq + python-pptx")


# ---------------------------------------------------------------------------
# Main content (Stitch UI layout conversion)
# ---------------------------------------------------------------------------

# Welcome Hero Section (Side-by-side title and image using st.columns)
hero_col1, hero_col2 = st.columns([0.65, 0.35], gap="large")

with hero_col1:
    st.markdown("""
    <div class="welcome-badge" style="margin-top: 0.2rem;">Welcome to DeckMate</div>
    <h2 class="welcome-title">
        Create your deck <span style="color: #2563eb;">in minutes</span>
        <span class="material-icons-outlined" style="color: #2563eb; vertical-align: top; font-size: 1.5rem; margin-left: 0.25rem;">auto_awesome</span>
    </h2>
    <p class="welcome-desc">
        Turn messy, multi-tab Excel sheets into polished, presentation-ready PowerPoint decks — with real charts, AI-written insights, and audience-aware framing.
    </p>
    """, unsafe_allow_html=True)

with hero_col2:
    st.image(
        "image.png",
        width="stretch",
    )

# Step 1: Upload Card (Stitch style wrapper)
st.markdown("""
<div class="stitch-card">
    <div class="card-header">
        <div class="card-icon-box-green">
            <span class="material-icons-outlined">table_view</span>
        </div>
        <div>
            <div class="card-title-text">Upload your Excel file</div>
            <div class="card-subtitle-text">Supports .xlsx files up to 10MB</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Drag and drop or browse for a .xlsx file",
    type=["xlsx"],
    label_visibility="collapsed",
)

if uploaded_file:
    file_size_mb = uploaded_file.size / (1024 * 1024)
    st.success(
        f"**{uploaded_file.name}** uploaded "
        f"({file_size_mb:.1f} MB)"
    )

st.markdown("<br>", unsafe_allow_html=True)

# Step 2: Describe Card (Stitch style wrapper)
st.markdown("""
<div class="stitch-card">
    <div class="card-header">
        <div class="card-icon-box-blue">
            <span class="material-icons-outlined">psychology</span>
        </div>
        <div>
            <div class="card-title-text">Describe your deck</div>
            <div class="card-subtitle-text">What should this deck focus on? Who is the audience?</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([3, 1])

with col1:
    user_prompt = st.text_area(
        "Focus Area",
        placeholder=(
            "Example: Focus on regional sales performance and marketing ROI. "
            "This is for a senior leadership quarterly review. "
            "Use professional, concise language. "
            "Highlight any standout trends or areas of concern."
        ),
        height=180,
        label_visibility="collapsed",
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

    st.markdown("<div style='margin-top: 1rem;'></div>", unsafe_allow_html=True)

    audience_preset = st.selectbox(
        "Quick preset",
        options=["Custom (use text above)", "Senior Leadership", "Client-Facing",
                 "Board Presentation", "Internal Team"],
    )

st.markdown("<br>", unsafe_allow_html=True)

# ---- Generate button ----
generate_disabled = uploaded_file is None

if generate_disabled:
    st.markdown("""
    <div class="blueprint-info-banner">
        <div class="blueprint-icon-circle">i</div>
        <span>Upload an Excel file to get started.</span>
    </div>
    """, unsafe_allow_html=True)

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

