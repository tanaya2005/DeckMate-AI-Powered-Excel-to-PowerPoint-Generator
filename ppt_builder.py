"""
ppt_builder.py — JSON slide plan + real DataFrames → .pptx (Day 3)

Takes the structured slide plan from llm_planner.py and the REAL dataframes
from data_parser.py, then builds a native, editable PowerPoint deck using
python-pptx.

ALL chart numbers come from pandas — the LLM only decides structure and
writes narrative text.

Pipeline position:
    Excel → [data_parser] → summary → [llm_planner] → slide plan →
    [ppt_builder] → downloadable .pptx
"""

import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor


# ---------------------------------------------------------------------------
# Brand / Theme constants
# ---------------------------------------------------------------------------

# Slide dimensions (standard 16:9)
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Colour palette
COLOR_PRIMARY    = RGBColor(0x1B, 0x2A, 0x4A)  # Dark navy
COLOR_SECONDARY  = RGBColor(0x2E, 0x86, 0xAB)  # Teal
COLOR_ACCENT     = RGBColor(0xE8, 0x6F, 0x51)  # Coral
COLOR_BG_LIGHT   = RGBColor(0xF5, 0xF5, 0xF5)  # Off-white
COLOR_TEXT_DARK   = RGBColor(0x33, 0x33, 0x33)  # Near-black
COLOR_TEXT_LIGHT  = RGBColor(0xFF, 0xFF, 0xFF)  # White
COLOR_SUBTITLE    = RGBColor(0x66, 0x66, 0x66)  # Grey

# Branded 6-Color Palette
COLOR_TEAL   = RGBColor(0x2E, 0x86, 0xAB)  # Teal
COLOR_CORAL  = RGBColor(0xE8, 0x6F, 0x51)  # Coral
COLOR_AMBER  = RGBColor(0xF2, 0xB7, 0x05)  # Amber
COLOR_SAGE   = RGBColor(0x76, 0x9A, 0x6B)  # Sage Green
COLOR_SLATE  = RGBColor(0x6C, 0x7A, 0x89)  # Slate
COLOR_LAVENDER = RGBColor(0x9B, 0x59, 0xB6) # Purple/Lavender

CHART_COLORS = [
    COLOR_TEAL,
    COLOR_CORAL,
    COLOR_AMBER,
    COLOR_SAGE,
    COLOR_SLATE,
    COLOR_LAVENDER
]

# Chart type mapping
CHART_TYPE_MAP = {
    "bar":  XL_CHART_TYPE.COLUMN_CLUSTERED,
    "pie":  XL_CHART_TYPE.PIE,
    "line": XL_CHART_TYPE.LINE_MARKERS,
}

# Layout positions (Wider elements to fill up whitespace)
TITLE_LEFT   = Inches(0.6)
TITLE_TOP    = Inches(0.4)
TITLE_WIDTH  = Inches(12.133)
TITLE_HEIGHT = Inches(0.8)

# Lowered chart/insight box to accommodate KPI callouts on top
CHART_LEFT   = Inches(0.6)
CHART_TOP    = Inches(2.5)       # Shifted down from 1.4 to 2.5
CHART_WIDTH  = Inches(7.8)
CHART_HEIGHT = Inches(4.1)       # Shrunk height from 5.0 to 4.1

INSIGHT_LEFT   = Inches(8.7)
INSIGHT_TOP    = Inches(2.5)       # Shifted down from 1.4 to 2.5
INSIGHT_WIDTH  = Inches(4.0)
INSIGHT_HEIGHT = Inches(4.1)       # Shrunk height from 5.0 to 4.1

TABLE_LEFT   = Inches(0.6)
TABLE_TOP    = Inches(2.5)
TABLE_WIDTH  = Inches(12.133)

FOOTER_LEFT   = Inches(0.6)
FOOTER_TOP    = Inches(6.9)
FOOTER_WIDTH  = Inches(12.133)
FOOTER_HEIGHT = Inches(0.4)


# ---------------------------------------------------------------------------
# Column name resolution (trust boundary)
# ---------------------------------------------------------------------------

def _normalize(name: str) -> str:
    """Normalize a column name for comparison: lowercase, strip whitespace."""
    return "".join(name.lower().split())


def _resolve_columns(
    requested_cols: List[str],
    df: pd.DataFrame,
) -> Tuple[List[str], List[str]]:
    """
    Resolve LLM-provided column names against the actual DataFrame columns.

    Uses case-insensitive, whitespace-tolerant matching.

    Returns
    -------
    (matched_cols, unmatched_cols)
        matched_cols  — list of ACTUAL DataFrame column names (in the order requested)
        unmatched_cols — list of LLM column names that couldn't be matched
    """
    actual_cols = list(df.columns)
    # Build a lookup: normalized_name -> actual_name
    lookup = {}
    for col in actual_cols:
        norm = _normalize(col)
        if norm not in lookup:  # first match wins (handles duplicates)
            lookup[norm] = col

    matched = []
    unmatched = []

    for req in requested_cols:
        norm_req = _normalize(req)
        if norm_req in lookup:
            matched.append(lookup[norm_req])
        else:
            unmatched.append(req)

    return matched, unmatched


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------

def _add_kpi_callouts(slide, kpis: dict):
    """Draw a row of styled KPI callout boxes at the top of the slide."""
    if not kpis:
        return

    # Filter out empty entries
    valid_kpis = {k: v for k, v in kpis.items() if k and v}
    if not valid_kpis:
        return

    # Limit to max 3 KPIs to fit the width
    items = list(valid_kpis.items())[:3]
    count = len(items)

    # Box coordinates
    kpi_y = Inches(1.3)
    kpi_h = Inches(0.9)
    total_w = Inches(12.133)
    
    # Calculate box width and spacing dynamically
    box_w = Inches(3.6)
    spacing = Inches(0.4)
    
    # Left margin offset to center the KPIs container
    container_w = (box_w * count) + (spacing * (count - 1))
    start_x = Inches(0.6) + (total_w - container_w) / 2

    from pptx.enum.shapes import MSO_SHAPE
    
    for i, (label, val) in enumerate(items):
        box_x = start_x + i * (box_w + spacing)
        
        # Add a rounded rectangle background card for each KPI
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, box_x, kpi_y, box_w, kpi_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_BG_LIGHT
        card.line.color.rgb = COLOR_SECONDARY
        card.line.width = Pt(1.5)
        
        # Add textbox inside the card for label and value
        txBox = slide.shapes.add_textbox(box_x, kpi_y + Inches(0.05), box_w, kpi_h - Inches(0.1))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.auto_size = None
        
        # KPI Value (large, bold)
        p_val = tf.paragraphs[0]
        p_val.text = str(val)
        p_val.font.size = Pt(20)
        p_val.font.bold = True
        p_val.font.color.rgb = COLOR_PRIMARY
        p_val.alignment = PP_ALIGN.CENTER
        
        # KPI Label (small, gray)
        p_lbl = tf.add_paragraph()
        p_lbl.text = str(label).upper()
        p_lbl.font.size = Pt(9)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = COLOR_SECONDARY
        p_lbl.alignment = PP_ALIGN.CENTER

def _add_title_textbox(slide, title_text: str):
    """Add a styled title text box to the top of a slide, stripping markdown bold."""
    txBox = slide.shapes.add_textbox(TITLE_LEFT, TITLE_TOP, TITLE_WIDTH, TITLE_HEIGHT)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    
    parts = title_text.split("**")
    for idx, part in enumerate(parts):
        if not part:
            continue
        run = p.add_run()
        run.text = part
        run.font.size = Pt(28)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY


def _add_insight_textbox(slide, insight_text: str, left=None, top=None,
                         width=None, height=None):
    """Add a styled insight text box, rendering subheadings and bullet points properly over a light card background."""
    x = left or INSIGHT_LEFT
    y = top or INSIGHT_TOP
    w = width or INSIGHT_WIDTH
    h = height or INSIGHT_HEIGHT

    # Add light background card shape (rounded rectangle or rectangle)
    from pptx.enum.shapes import MSO_SHAPE
    card = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, x, y, w, h
    )
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_BG_LIGHT
    card.line.fill.background() # No border
    
    # Add text box on top of the card (with slight margins)
    margin = Inches(0.25)
    txBox = slide.shapes.add_textbox(
        x + margin,
        y + margin,
        w - (margin * 2),
        h - (margin * 2),
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None

    # Title label
    label_p = tf.paragraphs[0]
    label_p.text = "ANALYSIS & KEY INSIGHTS"
    label_p.font.size = Pt(11)
    label_p.font.bold = True
    label_p.font.color.rgb = COLOR_SECONDARY
    label_p.space_after = Pt(12)

    # Split the insight text into segments
    lines = insight_text.replace("\\n", "\n").split("\n")
    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue

        p = tf.add_paragraph()
        p.space_after = Pt(8)
        p.line_spacing = Pt(18)

        # Check if it's a bullet point
        if line_str.startswith("•") or line_str.startswith("-") or line_str.startswith("*"):
            # Strip the bullet char
            content = line_str.lstrip("•-* ").strip()
            # Style bullet paragraph
            p.level = 0
            p.space_before = Pt(4)
            # Add bullet symbol explicitly
            run_bullet = p.add_run()
            run_bullet.text = "•  "
            run_bullet.font.size = Pt(13)
            run_bullet.font.bold = True
            run_bullet.font.color.rgb = COLOR_SECONDARY

            # Parse bold text within the bullet
            _add_formatted_runs(p, content)
        else:
            # Normal line or heading
            _add_formatted_runs(p, line_str)


def _add_formatted_runs(paragraph, text: str):
    """Parse text and add bold/normal runs (looks for **text**)."""
    parts = text.split("**")
    for idx, part in enumerate(parts):
        if not part:
            continue
        run = paragraph.add_run()
        run.text = part
        run.font.size = Pt(12)
        run.font.color.rgb = COLOR_TEXT_DARK

        # Odd indices are inside the ** ** block, so they should be bolded
        if idx % 2 == 1:
            run.font.bold = True
            run.font.color.rgb = COLOR_PRIMARY
        else:
            run.font.bold = False


def _add_footer(slide, slide_num: int, total_slides: int):
    """Add a subtle footer with slide number."""
    txBox = slide.shapes.add_textbox(FOOTER_LEFT, FOOTER_TOP, FOOTER_WIDTH, FOOTER_HEIGHT)
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = f"DeckMate  |  Slide {slide_num} of {total_slides}"
    p.font.size = Pt(9)
    p.font.color.rgb = COLOR_SUBTITLE
    p.alignment = PP_ALIGN.RIGHT


def _style_chart(chart, chart_type: str):
    """Apply consistent styling to a chart."""
    chart.has_legend = True
    chart.legend.include_in_layout = False
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.font.size = Pt(9)

    # Style the series/points colours
    has_points = (chart_type in ("pie", "bar"))
    
    for i, series in enumerate(chart.series):
        # Determine series level color
        color = CHART_COLORS[i % len(CHART_COLORS)]
        
        # If single series and is bar or pie, color individual points differently
        if len(chart.series) == 1 and has_points and hasattr(series, "points"):
            for p_idx in range(len(series.points)):
                pt = series.points[p_idx]
                pt_color = CHART_COLORS[p_idx % len(CHART_COLORS)]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = pt_color
        else:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = color

        # Line charts: style the line and markers
        if chart_type == "line":
            series.format.line.color.rgb = color
            series.format.line.width = Pt(2.5)
            if hasattr(series, 'marker'):
                series.marker.style = 8  # circle
                series.marker.size = 8

    # Value axis formatting (not applicable to pie)
    if chart_type != "pie":
        try:
            value_axis = chart.value_axis
            value_axis.has_title = False
            value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
            value_axis.format.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
            value_axis.tick_labels.font.size = Pt(9)
            value_axis.tick_labels.font.color.rgb = COLOR_SUBTITLE

            category_axis = chart.category_axis
            category_axis.format.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
            category_axis.tick_labels.font.size = Pt(9)
            category_axis.tick_labels.font.color.rgb = COLOR_SUBTITLE
        except Exception:
            pass  # some chart types may not have these axes


def _prepare_chart_data(
    df: pd.DataFrame,
    category_col: str,
    value_cols: List[str],
    chart_type: str,
) -> CategoryChartData:
    """
    Prepare chart data from the REAL dataframe.

    For bar/pie: aggregate by category column, sum the value columns.
    For line: use category as-is (typically dates/quarters), values as series.
    """
    chart_data = CategoryChartData()

    if chart_type in ("bar", "pie"):
        # Aggregate: group by category, sum numeric value columns
        agg_dict = {}
        for vc in value_cols:
            if pd.api.types.is_numeric_dtype(df[vc]):
                agg_dict[vc] = "sum"
            else:
                agg_dict[vc] = "count"

        if agg_dict:
            grouped = df.groupby(category_col, sort=True).agg(agg_dict).reset_index()
        else:
            # All value columns are non-numeric — count occurrences
            grouped = df[category_col].value_counts().reset_index()
            grouped.columns = [category_col, "Count"]
            value_cols = ["Count"]

    elif chart_type == "line":
        # For line charts, sort by the category (usually date/time)
        grouped = df.copy()
        try:
            grouped = grouped.sort_values(category_col)
        except Exception:
            pass

        # If there are many rows, aggregate by category
        if len(grouped) > 20:
            agg_dict = {vc: "mean" for vc in value_cols if pd.api.types.is_numeric_dtype(df[vc])}
            if agg_dict:
                grouped = grouped.groupby(category_col).agg(agg_dict).reset_index()
    else:
        grouped = df.copy()

    # Set categories
    categories = [str(v) for v in grouped[category_col].tolist()]
    chart_data.categories = categories

    # Add series
    for vc in value_cols:
        if vc in grouped.columns:
            values = grouped[vc].tolist()
            # Ensure numeric
            values = [float(v) if pd.notna(v) else 0.0 for v in values]
            chart_data.add_series(vc, values)

    return chart_data


def _build_chart_slide(
    prs: Presentation,
    slide_info: dict,
    df: pd.DataFrame,
    matched_cols: List[str],
    slide_num: int,
    total_slides: int,
):
    """Build a slide with a native chart + insight text with dynamic height budgeting."""
    slide_layout = prs.slide_layouts[6]  # blank layout
    slide = prs.slides.add_slide(slide_layout)

    chart_type = slide_info["chart_type"]
    title = slide_info.get("title", "")
    insight = slide_info.get("insight_text", "")

    # Add title
    _add_title_textbox(slide, title)

    # Add KPI Callouts
    kpis = slide_info.get("kpis", {})
    _add_kpi_callouts(slide, kpis)

    # Determine category vs value columns
    category_col = matched_cols[0]
    value_cols = matched_cols[1:] if len(matched_cols) > 1 else []

    if not value_cols:
        if pd.api.types.is_numeric_dtype(df[category_col]):
            value_cols = [category_col]
            category_col = df.index.name or "Index"
            if category_col not in df.columns:
                df = df.reset_index()
        else:
            counts = df[category_col].value_counts().reset_index()
            counts.columns = [category_col, "Count"]
            df = counts
            value_cols = ["Count"]

    # Dynamic Height Budgeting:
    # Slide height = 7.5. Footer top = 6.9. Title & KPIs take up top 2.3.
    # Total remaining space between y = 2.4 and y = 6.8 is 4.4 inches.
    content_top = Inches(2.4)
    max_content_h = Inches(4.3)
    
    # We fit the chart into the max height
    chart_h = max_content_h

    # Prepare chart data from REAL dataframe
    chart_data = _prepare_chart_data(df, category_col, value_cols, chart_type)
    xl_chart_type = CHART_TYPE_MAP.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    # Add chart
    chart_frame = slide.shapes.add_chart(
        xl_chart_type,
        CHART_LEFT, content_top, CHART_WIDTH, chart_h,
        chart_data,
    )

    # Style the chart
    _style_chart(chart_frame.chart, chart_type)

    # Add insight text box (right side)
    _add_insight_textbox(slide, insight, left=INSIGHT_LEFT, top=content_top, width=INSIGHT_WIDTH, height=chart_h)

    # Add footer
    _add_footer(slide, slide_num, total_slides)


def _build_table_slide(
    prs: Presentation,
    slide_info: dict,
    df: pd.DataFrame,
    matched_cols: List[str],
    slide_num: int,
    total_slides: int,
):
    """Build a slide with a data table + insight text with dynamic row shrinking and height budgeting."""
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)

    title = slide_info.get("title", "")
    insight = slide_info.get("insight_text", "")

    _add_title_textbox(slide, title)

    # Add KPI Callouts
    kpis = slide_info.get("kpis", {})
    _add_kpi_callouts(slide, kpis)

    # Prepare table data (no hard truncate, show all rows requested by plan)
    table_df = df[matched_cols].copy()
    rows = len(table_df) + 1  # +1 for header
    cols = len(matched_cols)

    # Dynamic Height Budgeting:
    # Slide height = 7.5. Available vertical area for content: from y = 2.4 to y = 6.8 (4.4 inches).
    content_top = Inches(2.4)
    max_height_budget = Inches(4.3)
    
    # Reserve space for insights (minimum guaranteed height)
    min_insight_h = Inches(1.3)
    max_table_h = max_height_budget - min_insight_h - Inches(0.2) # leaves buffer

    # Calculate row height dynamically
    default_row_h = Inches(0.4)
    needed_height = default_row_h * rows
    
    font_size = 10
    row_h = default_row_h

    if needed_height > max_table_h:
        # Scale down row height to fit the budget
        row_h = max_table_h / rows
        # Sensible minimum row height (approx 0.2 inches)
        min_row_h = Inches(0.22)
        if row_h < min_row_h:
            print(f"[ppt_builder] WARNING: Table has too many rows ({rows}). Even at min row height, it may overflow.")
            row_h = min_row_h
        
        # Scale font down proportionally (min 8pt, default 10pt)
        ratio = float(row_h) / float(default_row_h)
        font_size = max(8, int(10 * ratio))

    table_height = row_h * rows
    
    # Scale column widths if there are many columns
    default_col_w = TABLE_WIDTH / cols
    col_w = default_col_w
    col_font_size = font_size
    if cols > 6:
        col_font_size = max(8, font_size - 1)

    table_shape = slide.shapes.add_table(
        rows, cols,
        TABLE_LEFT, content_top, TABLE_WIDTH, table_height,
    )
    table = table_shape.table

    # Set row heights explicitly
    for r_idx in range(rows):
        table.rows[r_idx].height = row_h

    # Style header row
    for j, col_name in enumerate(matched_cols):
        cell = table.cell(0, j)
        cell.text = str(col_name)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.font.color.rgb = COLOR_TEXT_LIGHT
        p.font.size = Pt(max(9, col_font_size + 1))
        p.font.bold = True
        p.alignment = PP_ALIGN.LEFT

    # Fill data rows
    for i in range(len(table_df)):
        for j, col_name in enumerate(matched_cols):
            cell = table.cell(i + 1, j)
            val = table_df.iloc[i][col_name]
            if pd.isna(val):
                cell.text = "—"
            elif isinstance(val, float):
                cell.text = f"{val:,.2f}"
            else:
                cell.text = str(val)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(col_font_size)
            p.font.color.rgb = COLOR_TEXT_DARK
            p.alignment = PP_ALIGN.LEFT

            # Alternate row shading
            if i % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = COLOR_BG_LIGHT

    # Position the insight box dynamically BELOW the table
    insight_top = content_top + table_height + Inches(0.15)
    insight_h = max(min_insight_h, max_height_budget - table_height)
    
    _add_insight_textbox(
        slide, insight,
        left=TABLE_LEFT, top=insight_top,
        width=TABLE_WIDTH, height=insight_h,
    )

    _add_footer(slide, slide_num, total_slides)


def _build_text_slide(
    prs: Presentation,
    slide_info: dict,
    slide_num: int,
    total_slides: int,
    is_fallback: bool = False,
):
    """Build a text-only slide (title, takeaways, or fallback for unmatched columns)."""
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)

    title = slide_info.get("title", "")
    insight = slide_info.get("insight_text", "")

    # Centre the title for text slides
    txBox = slide.shapes.add_textbox(
        Inches(1.0), Inches(2.0), Inches(11.0), Inches(1.2),
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    
    parts_title = title.split("**")
    for idx, part in enumerate(parts_title):
        if not part:
            continue
        run = p.add_run()
        run.text = part
        run.font.size = Pt(36)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY

    # Body text
    if insight:
        body_box = slide.shapes.add_textbox(
            Inches(1.5), Inches(3.5), Inches(10.0), Inches(3.0),
        )
        bf = body_box.text_frame
        bf.word_wrap = True
        bp = bf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        bp.line_spacing = Pt(28)
        
        parts_body = insight.replace("\\n", "\n").split("**")
        for idx, part in enumerate(parts_body):
            if not part:
                continue
            run = bp.add_run()
            run.text = part
            run.font.size = Pt(18)
            run.font.color.rgb = COLOR_TEXT_DARK
            if idx % 2 == 1:
                run.font.bold = True
                run.font.color.rgb = COLOR_PRIMARY

    # Fallback notice
    if is_fallback:
        warn_box = slide.shapes.add_textbox(
            Inches(1.5), Inches(6.0), Inches(10.0), Inches(0.5),
        )
        wf = warn_box.text_frame
        wp = wf.paragraphs[0]
        wp.text = "(Chart could not be rendered — column mismatch with data source)"
        wp.font.size = Pt(10)
        wp.font.color.rgb = COLOR_SUBTITLE
        wp.font.italic = True
        wp.alignment = PP_ALIGN.CENTER

    _add_footer(slide, slide_num, total_slides)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_presentation(
    slide_plan: dict,
    dataframes: dict,
    output_path: str,
) -> str:
    """
    Build a PowerPoint presentation from the slide plan and real DataFrames.

    Parameters
    ----------
    slide_plan : dict
        Output from llm_planner.generate_slide_plan(). Must have a "slides" key.
    dataframes : dict
        Dict of {sheet_name: pd.DataFrame} from data_parser.parse_excel().
    output_path : str
        Where to save the .pptx file.

    Returns
    -------
    str — the output path (for chaining).
    """
    prs = Presentation()

    # Set 16:9 slide size
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    slides = slide_plan.get("slides", [])
    total_slides = len(slides)

    for idx, slide_info in enumerate(slides):
        slide_num = idx + 1
        chart_type = slide_info.get("chart_type", "text").lower().strip()
        source_sheet = slide_info.get("source_sheet", "none")
        data_columns = slide_info.get("data_columns", [])

        # ------- TEXT SLIDES (no data needed) -------
        if chart_type == "text" or source_sheet.lower() == "none":
            _build_text_slide(prs, slide_info, slide_num, total_slides)
            continue

        # ------- DATA SLIDES (chart / table) -------
        # Find the source dataframe
        df = dataframes.get(source_sheet)
        if df is None or df.empty:
            # Sheet not found or empty → fallback to text
            print(f"[ppt_builder] WARNING: Sheet '{source_sheet}' not found or empty "
                  f"for slide {slide_num} ('{slide_info.get('title', '')}'). "
                  f"Rendering as text slide.")
            _build_text_slide(prs, slide_info, slide_num, total_slides, is_fallback=True)
            continue

        if not data_columns:
            # No columns specified → text slide
            _build_text_slide(prs, slide_info, slide_num, total_slides)
            continue

        # ------- COLUMN VALIDATION (trust boundary) -------
        matched_cols, unmatched_cols = _resolve_columns(data_columns, df)

        if unmatched_cols:
            print(f"[ppt_builder] WARNING: Slide {slide_num} ('{slide_info.get('title', '')}') "
                  f"— LLM requested columns {unmatched_cols} which don't exist in "
                  f"sheet '{source_sheet}'. Available: {list(df.columns)}")

        if len(matched_cols) < 1:
            # No valid columns at all → fallback to text
            print(f"[ppt_builder] WARNING: No valid columns for slide {slide_num}. "
                  f"Rendering as text-only fallback.")
            _build_text_slide(prs, slide_info, slide_num, total_slides, is_fallback=True)
            continue

        # ------- BUILD THE APPROPRIATE SLIDE TYPE -------
        try:
            df_copy = df.copy()
            if chart_type in ("bar", "pie", "line"):
                _build_chart_slide(
                    prs, slide_info, df_copy, matched_cols, slide_num, total_slides,
                )
            elif chart_type == "table":
                _build_table_slide(
                    prs, slide_info, df_copy, matched_cols, slide_num, total_slides,
                )
            else:
                _build_text_slide(prs, slide_info, slide_num, total_slides)

        except Exception as e:
            # If chart building fails for any reason, fallback to text
            print(f"[ppt_builder] ERROR building slide {slide_num}: {e}. "
                  f"Falling back to text slide.")
            _build_text_slide(prs, slide_info, slide_num, total_slides, is_fallback=True)

    # Save
    prs.save(output_path)
    print(f"[ppt_builder] Saved presentation to: {output_path}")
    return output_path


# ---------------------------------------------------------------------------
# CLI test block
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import json
    import numpy as np

    print("=" * 60)
    print("DAY 3 TEST: PPTX Builder")
    print("=" * 60)

    # Build a fake slide plan + dataframes to test each slide type
    test_df_sales = pd.DataFrame({
        "Region": ["North", "South", "East", "West"] * 3,
        "Product": ["Widget A", "Widget B", "Widget C",
                     "Widget A", "Widget B", "Widget C",
                     "Widget A", "Widget B", "Widget C",
                     "Widget A", "Widget B", "Widget C"],
        "Revenue": [120000, 95000, 145000, 88000,
                    110000, 72000, 130000, 98000,
                    155000, 91000, 105000, 85000],
        "Units Sold": [450, 380, 520, 310,
                       410, 290, 480, 370,
                       560, 340, 400, 320],
    })

    test_df_marketing = pd.DataFrame({
        "Quarter": ["Q1", "Q1", "Q2", "Q2", "Q3", "Q3", "Q4", "Q4"],
        "Channel": ["Online", "Retail", "Online", "Retail",
                     "Online", "Retail", "Online", "Retail"],
        "Spend": [50000, 35000, 62000, 41000, 71000, 45000, 80000, 52000],
        "ROI %": [12.5, 8.3, 15.1, 9.7, 18.2, 11.0, 20.5, 13.4],
    })

    dataframes = {
        "Regional Sales": test_df_sales,
        "Marketing Spend": test_df_marketing,
    }

    slide_plan = {
        "slides": [
            {
                "title": "Sales Performance Overview",
                "source_sheet": "none",
                "chart_type": "text",
                "data_columns": [],
                "insight_text": "This deck presents an overview of regional sales performance and marketing ROI for the fiscal year, prepared for senior leadership review.",
            },
            {
                "title": "Revenue by Region",
                "source_sheet": "Regional Sales",
                "chart_type": "bar",
                "data_columns": ["Region", "Revenue"],
                "insight_text": "The East region led revenue generation with $400K total, while the West region trailed at $264K. North and South regions showed competitive mid-range performance.",
            },
            {
                "title": "Product Revenue Share",
                "source_sheet": "Regional Sales",
                "chart_type": "pie",
                "data_columns": ["Product", "Revenue"],
                "insight_text": "Widget C commands the largest share of total revenue at 38%, followed by Widget A at 35%. Widget B contributed 27% and represents a potential growth area.",
            },
            {
                "title": "Marketing ROI Trend",
                "source_sheet": "Marketing Spend",
                "chart_type": "line",
                "data_columns": ["Quarter", "ROI %"],
                "insight_text": "Marketing ROI has shown consistent quarter-over-quarter improvement, rising from 10.4% in Q1 to 17% in Q4. This upward trajectory validates increased investment in online channels.",
            },
            {
                "title": "Marketing Spend Breakdown",
                "source_sheet": "Marketing Spend",
                "chart_type": "table",
                "data_columns": ["Quarter", "Channel", "Spend", "ROI %"],
                "insight_text": "Online consistently outperforms Retail on ROI, while Retail spend has grown more modestly. Consider reallocating a portion of Retail budget to online channels.",
            },
            {
                # This slide deliberately has a WRONG column name to test
                # the fuzzy-match fallback
                "title": "Hallucinated Column Test",
                "source_sheet": "Regional Sales",
                "chart_type": "bar",
                "data_columns": ["Region", "TotalRevenue"],  # "TotalRevenue" doesn't exist
                "insight_text": "This slide tests the fallback when the LLM hallucinates a column name that doesn't exist in the data.",
            },
            {
                # Test fuzzy match: "ROI%" (no space) vs actual "ROI %"
                "title": "Fuzzy Column Match Test",
                "source_sheet": "Marketing Spend",
                "chart_type": "bar",
                "data_columns": ["Channel", "ROI%"],  # "ROI%" should match "ROI %"
                "insight_text": "This tests that 'ROI%' fuzzy-matches to the actual column 'ROI %' despite missing whitespace.",
            },
            {
                "title": "Key Takeaways",
                "source_sheet": "none",
                "chart_type": "text",
                "data_columns": [],
                "insight_text": "East region is the clear revenue leader. Marketing ROI is trending upward. Online channels significantly outperform Retail on ROI. Recommend increased online investment.",
            },
        ]
    }

    output_file = "test_output.pptx"

    print(f"\nBuilding {len(slide_plan['slides'])} slides...\n")
    build_presentation(slide_plan, dataframes, output_file)

    print(f"\nDone! Open '{output_file}' to verify:")
    print("  - Slide 1: Text (title/overview)")
    print("  - Slide 2: Bar chart (Revenue by Region)")
    print("  - Slide 3: Pie chart (Product Revenue Share)")
    print("  - Slide 4: Line chart (Marketing ROI Trend)")
    print("  - Slide 5: Table (Marketing Spend Breakdown)")
    print("  - Slide 6: Text FALLBACK (hallucinated column 'TotalRevenue')")
    print("  - Slide 7: Bar chart (fuzzy match 'ROI%' -> 'ROI %')")
    print("  - Slide 8: Text (Key Takeaways)")
