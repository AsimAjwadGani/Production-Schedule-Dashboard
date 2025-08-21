import streamlit as st
import calendar
import json
from pathlib import Path
from datetime import date, datetime
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.pdfgen import canvas
import io
import base64
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
import re

# =======================
# CONFIG & CONSTANTS
# =======================
st.set_page_config(page_title="Production Schedule Dashboard", layout="wide")

COLOR_AC225_RUN_EVG = "#A2EBCD"
COLOR_IN111_RUN_EVG = "#F1E183"
COLOR_AC225_RUN_SRX = "#F3B48F"
COLOR_IN111_RUN_SRX = "#EC712A"
COLOR_CARDINAL_TPI_NIOWAVE = "#0ABB21"
COLOR_NMCTG = "#BCA6CA"
COLOR_PLACEHOLDER = "#F1E429"
COLOR_SHUTDOWN = "#F5253A"
COLOR_CONFIRMED = "#5F65BB"
COLOR_PV = "#3CD63C"
COLOR_SRX = "#3ACCC0"
COLOR_PERCEPTIVE = "#75A06B"
COLOR_BWXT = "#3D6E34"
COLOR_MAINTENANCE = "#D10D96"

DASHBOARD_NAME = "production_schedule"
FILENAME = f"{DASHBOARD_NAME}.json"
CONFIG_FILE = Path.home() / ".production_schedule_config.json"
DEFAULT_DIR = Path.home() / "Schedules"  # change if you want a different default

# =======================
# SESSION INIT
# =======================
ss = st.session_state
ss.setdefault("current_month", date.today().month)
ss.setdefault("current_year", date.today().year)
ss.setdefault("week_action_rows", {})  # "YYYY-MM_week_idx" -> row count (month-specific)
ss.setdefault("entries", {})  # "YYYY-MM-DD_row" -> text
ss.setdefault("__autosave_ok__", False)
ss.setdefault("__autosave_error__", "")
ss.setdefault("show_legend", False)  # legend drawer visibility
if "custom_legend_entries" not in ss:
    ss.custom_legend_entries = []

# pending load buffer (used when dir changes or on initial boot load)
ss.setdefault("__pending_entries__", None)
ss.setdefault("__pending_meta__", None)
ss.setdefault("__pending_week_action_rows__", None)

# =======================
# RERUN FUNCTION (compatible)
# =======================
def rerun():
    st.rerun()

# =======================
# HELPERS
# =======================
def _sanitize_dir(raw: str) -> Path:
    s = (raw or "").strip().strip('"').strip("'")
    return Path(s).expanduser()

def load_latest_dir() -> Path:
    if "latest_dir" in ss:
        return Path(ss.latest_dir)
    if CONFIG_FILE.exists():
        try:
            cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            last = cfg.get("latest_dir")
            if last:
                ss.latest_dir = last
                return Path(last)
        except Exception:
            pass
    ss.latest_dir = str(DEFAULT_DIR)
    return DEFAULT_DIR

def save_latest_dir(dir_str: str) -> None:
    ss.latest_dir = dir_str
    try:
        CONFIG_FILE.write_text(json.dumps({"latest_dir": dir_str}, indent=2), encoding="utf-8")
    except Exception as e:
        st.warning(f"Couldn't persist latest directory: {e}")

def get_color(text: str) -> str:
    if not text or not text.strip():
        return "white"
    lower = text.strip().lower()

    if "custom_legend_entries" in st.session_state:
        for item in st.session_state.custom_legend_entries:
            if item["label"].strip().lower() in lower:
                return item["color"]

    if "shutdown" in lower:
        return COLOR_SHUTDOWN
    if "ac225 run-evg" in lower:
        return COLOR_AC225_RUN_EVG
    if "in111 run-evg" in lower:
        return COLOR_IN111_RUN_EVG
    if "ac225 run-srx" in lower:
        return COLOR_AC225_RUN_SRX
    if "in111 run-srx" in lower:
        return COLOR_IN111_RUN_SRX
    if lower.startswith(("cardinal", "tpi", "niowave")):
        return COLOR_CARDINAL_TPI_NIOWAVE
    if "nmctg" in lower:
        return COLOR_NMCTG
    if re.match(r"^\d{5}-p\d", lower):
        return COLOR_PLACEHOLDER
    if re.match(r"^\d{5}-\d{3}", lower):
        if re.search(r"MD[123]$", lower, re.IGNORECASE):
            return COLOR_MAINTENANCE
        return COLOR_CONFIRMED
    if lower.startswith("pv") and "srx" in lower:
        return COLOR_PV
    if lower == "srx maintenance":
        return COLOR_SRX
    if "perceptive" in lower:
        return COLOR_PERCEPTIVE
    if lower == "bwxt order":
        return COLOR_BWXT
    return "white"

def date_key(y: int, m: int, d: int, row_idx: int) -> str:
    """Unique per actual calendar day + row, e.g., '2025-08-13_0'."""
    return f"{date(y, m, d).isoformat()}_{row_idx}"

def _save_payload() -> dict:
    return {
        "meta": {"year": ss.current_year, "month": ss.current_month},
        "entries": ss.get("entries", {}),
        "week_action_rows": ss.get("week_action_rows", {}),
        "custom_legend_entries": ss.get("custom_legend_entries", []),
    }

def _autosave_now() -> Path:
    latest_dir = _sanitize_dir(str(load_latest_dir()))
    file_path = latest_dir / FILENAME
    file_path.parent.mkdir(parents=True, exist_ok=True)
    with file_path.open("w", encoding="utf-8") as f:
        json.dump(_save_payload(), f, indent=2, ensure_ascii=False)
    return file_path

def _commit_and_autosave(dkey: str, widget_key: str):
    """Commit this cell's value and autosave JSON."""
    try:
        ss.entries[dkey] = ss.get(widget_key, "")
        _autosave_now()
        ss["__autosave_ok__"] = True
        ss["__autosave_error__"] = ""
    except Exception as e:
        ss["__autosave_ok__"] = False
        ss["__autosave_error__"] = str(e)

def _commit_all_widgets_and_autosave():
    """Commit all current widget values to entries and autosave."""
    try:
        # Commit all widget values to entries
        for k, v in ss.entries.items():
            widget_key = f"cell_widget_{k}"
            if widget_key in ss:
                ss.entries[k] = ss.get(widget_key, "")
        _autosave_now()
        ss["__autosave_ok__"] = True
        ss["__autosave_error__"] = ""
    except Exception as e:
        ss["__autosave_ok__"] = False
        ss["__autosave_error__"] = str(e)

def _apply_meta_to_calendar(meta: dict) -> bool:
    """If meta has month/year, switch calendar to it. Returns True if changed."""
    try:
        yr = int(meta.get("year"))
        mo = int(meta.get("month"))
        changed = (yr != ss.current_year) or (mo != ss.current_month)
        ss.current_year = yr
        ss.current_month = mo
        return changed
    except Exception:
        return False

def _preload_widgets_from_entries():
    """Set widget values from entries BEFORE widgets are created."""
    for k, v in ss.entries.items():
        ss[f"cell_widget_{k}"] = v

def _sync_widgets_with_entries():
    """Ensure all widget values match their corresponding entries."""
    for k, v in ss.entries.items():
        widget_key = f"cell_widget_{k}"
        if widget_key not in ss or ss[widget_key] != v:
            ss[widget_key] = v

def _ensure_rows_for_current_month(valid_weeks):
    """Ensure enough rows to display all saved entries."""
    day_to_week = {}
    for w_idx, week in enumerate(valid_weeks):
        for d in week:
            if d != 0:
                day_to_week[d] = w_idx

    required_rows = {}
    for k in ss.entries.keys():
        try:
            dpart, rpart = k.split("_", 1)
            dt = datetime.fromisoformat(dpart).date()
            if dt.year == ss.current_year and dt.month == ss.current_month:
                row_i = int(rpart)
                w_idx = day_to_week.get(dt.day, None)
                if w_idx is None:
                    continue
                needed = row_i + 1
                required_rows[w_idx] = max(required_rows.get(w_idx, 1), needed)
        except Exception:
            continue

    for w_idx in range(len(valid_weeks)):
        current = ss.week_action_rows.get(f"{ss.current_year}-{ss.current_month}_{w_idx}", 1)
        ss.week_action_rows[f"{ss.current_year}-{ss.current_month}_{w_idx}"] = max(current, required_rows.get(w_idx, 1))

def get_week_action_rows_for_month(year, month):
    """Get week action rows for a specific month."""
    month_week_rows = {}
    month_key = f"{year}-{month}"
    for key, value in ss.week_action_rows.items():
        if key.startswith(f"{month_key}_"):
            week_idx = int(key.split("_")[1])
            month_week_rows[week_idx] = value
    return month_week_rows

def _try_load_from(path: Path):
    """Read JSON file and return (entries, meta, week_action_rows, full_data) or (None, None, None, None)."""
    try:
        if not path.exists():
            return None, None, None, None
        data = json.loads(path.read_text(encoding="utf-8")) or {}
        if isinstance(data, dict) and "entries" in data:
            entries = data.get("entries", {}) or {}
            meta = data.get("meta") or {}
            week_action_rows = data.get("week_action_rows", {}) or {}
            return entries, meta, week_action_rows, data  # Return full data
    except Exception as e:
        st.error(f"Load error: {e}")
        pass
    return None, None, None, None

# =======================
# PDF GENERATION (FIXED)
# =======================
def generate_pdf_calendar(year, month, entries, week_action_rows):
    """Generate a PDF calendar with full weeks, trailing dates, and wrapped text."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=0.4*inch,
        rightMargin=0.4*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch
    )
    story = []

    styles = getSampleStyleSheet()

    # Title styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=15,
        alignment=1
    )
    month_style = ParagraphStyle(
        'MonthStyle',
        parent=styles['Heading2'],
        fontSize=20,
        spaceAfter=20,
        alignment=1
    )

    story.append(Paragraph("Production Schedule Dashboard", title_style))
    month_name = calendar.month_name[month]
    story.append(Paragraph(f"{month_name} {year}", month_style))

    # --- Build extended weeks with real dates ---
    cal_raw = calendar.monthcalendar(year, month)
    valid_weeks = [week for week in cal_raw if any(d != 0 for d in week)]

    def prev_month(y, m):
        return (y - 1, 12) if m == 1 else (y, m - 1)

    def next_month(y, m):
        return (y + 1, 1) if m == 12 else (y, m + 1)

    extended_weeks = []
    for week in valid_weeks:
        ref_day = next((d for d in week if d != 0), None)
        if ref_day is None:
            extended_weeks.append([None]*7)
            continue
        ref_date = date(year, month, ref_day)
        ref_weekday = ref_date.weekday()  # Mon=0
        week_start = ref_date - timedelta(days=ref_weekday)
        extended_week = [week_start + timedelta(days=i) for i in range(7)]
        extended_weeks.append(extended_week)

    # --- Define styles for wrapped text ---
    cell_style = ParagraphStyle(
        'TableCell',
        fontSize=9,
        leading=10,
        alignment=1,  # Center
        wordWrap='CJK',  # Enable wrapping
        spaceAfter=2,
        textColor=colors.black
    )

    header_style = ParagraphStyle(
        'HeaderCell',
        parent=cell_style,
        fontSize=12,
        textColor=colors.whitesmoke,
        fontName='Helvetica-Bold'
    )

    date_style = ParagraphStyle(
        'DateCell',
        parent=cell_style,
        fontSize=12,
        textColor=colors.black,
        fontName='Helvetica-Bold'
    )

    # --- Table Data ---
    day_headers = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    table_data = [[Paragraph(h, header_style) for h in day_headers]]  # Header row

    row_heights = [0.4]  # Header

    for week_idx, week_dates in enumerate(extended_weeks):
        # Date row
        date_row = []
        for dt in week_dates:
            if dt is None:
                date_row.append("")
            else:
                label = f"{dt.strftime('%b-%d')}"
                date_row.append(Paragraph(label, date_style))
        table_data.append(date_row)
        row_heights.append(0.35)  # Slightly taller for date

        # Activity rows
        num_rows = week_action_rows.get(week_idx, 1)
        for row_idx in range(num_rows):
            activity_row = []
            for dt in week_dates:
                if dt is None:
                    activity_row.append("")
                else:
                    dkey = date_key(dt.year, dt.month, dt.day, row_idx)
                    entry = entries.get(dkey, "").strip()
                    if entry:
                        # Apply color-specific text color later; just wrap text now
                        p_style = cell_style.clone('tmp')
                        color_hex = get_color(entry)
                        if color_hex in [
                            COLOR_AC225_RUN_EVG, COLOR_IN111_RUN_EVG,
                            COLOR_PLACEHOLDER, COLOR_CARDINAL_TPI_NIOWAVE,
                            COLOR_NMCTG, COLOR_PERCEPTIVE, COLOR_IN111_RUN_SRX
                        ]:
                            p_style.textColor = colors.black
                        else:
                            p_style.textColor = colors.white
                        p = Paragraph(entry, p_style)
                        activity_row.append(p)
                    else:
                        activity_row.append("")
            table_data.append(activity_row)
            row_heights.append(0.5)  # Taller for multi-line text

    # --- Create Table ---
    col_widths = [1.1 * inch] * 7
    table = Table(table_data, colWidths=col_widths, rowHeights=[h * inch for h in row_heights])

    # Base table style
    table_style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('BACKGROUND', (0, 1), (-1, 1), colors.lightgrey),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ])

    # Apply formatting per row
    current_row = 1
    for week_idx, week_dates in enumerate(extended_weeks):
        # Date row already styled via Paragraph
        current_row += 1

        num_activity_rows = week_action_rows.get(week_idx, 1)
        for row_offset in range(num_activity_rows):
            for day_idx, dt in enumerate(week_dates):
                if dt is None:
                    continue
                dkey = date_key(dt.year, dt.month, dt.day, row_offset)
                entry = entries.get(dkey, "").strip()
                if not entry:
                    continue
                color_hex = get_color(entry)
                if color_hex == "white":
                    continue
                try:
                    hex_clean = color_hex.lstrip('#')
                    r = int(hex_clean[0:2], 16) / 255.0
                    g = int(hex_clean[2:4], 16) / 255.0
                    b = int(hex_clean[4:6], 16) / 255.0
                    bg_color = colors.Color(r, g, b)
                    table_style.add('BACKGROUND', (day_idx, current_row), (day_idx, current_row), bg_color)
                except Exception:
                    pass
            current_row += 1

    table.setStyle(table_style)
    story.append(table)
    story.append(Spacer(1, 0.3*inch))

    # Build PDF
    doc.build(story)
    pdf_data = buffer.getvalue()
    buffer.close()
    return pdf_data

def generate_ppt_calendar(year, month, entries, week_action_rows):
    """Generate a PowerPoint calendar with compact, text-wrapped cells and proper scaling."""
    from datetime import date as dt, timedelta
    import calendar

    def prev_month(y, m):
        return (y - 1, 12) if m == 1 else (y, m - 1)

    def next_month(y, m):
        return (y + 1, 1) if m == 12 else (y, m + 1)

    # Build extended weeks (Mon–Sun)
    cal_raw = calendar.monthcalendar(year, month)
    valid_weeks = [week for week in cal_raw if any(d != 0 for d in week)]

    extended_weeks = []
    for week in valid_weeks:
        ref_day = next((d for d in week if d != 0), None)
        if ref_day is None:
            extended_weeks.append([None] * 7)
            continue
        ref_date = dt(year, month, ref_day)
        start_of_week = ref_date - timedelta(days=ref_date.weekday())  # Monday
        week_dates = [start_of_week + timedelta(days=i) for i in range(7)]
        extended_weeks.append(week_dates)

    prs = Presentation()
    slide_width = Inches(13.33)
    slide_height = Inches(7.5)
    margin = Inches(0.4)
    prs.slide_width = int(slide_width)
    prs.slide_height = int(slide_height)

    # Size settings — compact but readable
    CELL_WIDTH = (slide_width - 2 * margin) / 7
    HEADER_ROW_HEIGHT = Inches(0.3)      # Smaller header
    DATE_ROW_HEIGHT = Inches(0.28)       # Compact date row
    ACTIVITY_ROW_HEIGHT = Inches(0.35)   # Just enough for wrapped text

    # Available vertical space
    TOP_MARGIN = Inches(1.0)             # Title + subtitle
    BOTTOM_LIMIT = slide_height - margin

    current_week_idx = 0
    slide = None

    while current_week_idx < len(extended_weeks):
        # Start a new slide
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # === Add Title and Subtitle ===
        title_box = slide.shapes.add_textbox(margin, Inches(0.2), slide_width - 2*margin, Inches(0.3))
        title_frame = title_box.text_frame
        title_frame.text = "Production Schedule Dashboard"
        title_frame.paragraphs[0].font.size = Pt(16)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        subtitle_box = slide.shapes.add_textbox(margin, Inches(0.5), slide_width - 2*margin, Inches(0.2))
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = f"{calendar.month_name[month]} {year}"
        subtitle_frame.paragraphs[0].font.size = Pt(12)
        subtitle_frame.paragraphs[0].font.bold = True
        subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        y_current = TOP_MARGIN  # Start below title

        # === Day Headers ===
        day_headers = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        for i, day_name in enumerate(day_headers):
            x = margin + i * CELL_WIDTH
            header_box = slide.shapes.add_textbox(x, y_current, CELL_WIDTH, HEADER_ROW_HEIGHT)
            frame = header_box.text_frame
            frame.text = day_name
            frame.paragraphs[0].font.size = Pt(9)
            frame.paragraphs[0].font.bold = True
            frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            frame.word_wrap = True
            frame.auto_size = True
            header_box.fill.solid()
            header_box.fill.fore_color.rgb = RGBColor(128, 128, 128)
            frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        y_current += HEADER_ROW_HEIGHT

        # Add weeks until we run out of space
        while current_week_idx < len(extended_weeks):
            week_dates = extended_weeks[current_week_idx]
            week_idx = current_week_idx

            # Estimate height needed
            num_activity_rows = week_action_rows.get(week_idx, 1)
            week_height = DATE_ROW_HEIGHT + (num_activity_rows * ACTIVITY_ROW_HEIGHT)

            if y_current + week_height > BOTTOM_LIMIT:
                break  # Force new slide

            # === Date Row ===
            for day_idx, dt_obj in enumerate(week_dates):
                if dt_obj is None:
                    continue
                x = margin + day_idx * CELL_WIDTH
                date_box = slide.shapes.add_textbox(x, y_current, CELL_WIDTH, DATE_ROW_HEIGHT)
                frame = date_box.text_frame
                frame.text = dt_obj.strftime("%b-%d")
                frame.paragraphs[0].font.size = Pt(8)
                frame.paragraphs[0].font.bold = True
                frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                frame.word_wrap = True
                frame.auto_size = True
                date_box.fill.solid()
                date_box.fill.fore_color.rgb = RGBColor(240, 240, 240)
                frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)
            y_current += DATE_ROW_HEIGHT

            # === Activity Rows ===
            for row_idx in range(week_action_rows.get(week_idx, 1)):
                for day_idx, dt_obj in enumerate(week_dates):
                    if dt_obj is None:
                        continue
                    x = margin + day_idx * CELL_WIDTH
                    y = y_current
                    dkey = date_key(dt_obj.year, dt_obj.month, dt_obj.day, row_idx)
                    entry = entries.get(dkey, "").strip()

                    activity_box = slide.shapes.add_textbox(x, y, CELL_WIDTH, ACTIVITY_ROW_HEIGHT)
                    frame = activity_box.text_frame
                    if entry:
                        frame.text = entry
                    else:
                        frame.text = ""

                    # ✅ Enable wrapping and dynamic sizing
                    frame.word_wrap = True
                    frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE  # Shrink text to fit
                    p = frame.paragraphs[0]
                    p.font.size = Pt(8)
                    p.font.bold = True
                    p.alignment = PP_ALIGN.CENTER

                    # Apply color
                    if entry:
                        color_hex = get_color(entry)
                        if color_hex != "white":
                            try:
                                hex_clean = color_hex.lstrip('#')
                                r = int(hex_clean[0:2], 16)
                                g = int(hex_clean[2:4], 16)
                                b = int(hex_clean[4:6], 16)
                                activity_box.fill.solid()
                                activity_box.fill.fore_color.rgb = RGBColor(r, g, b)

                                light_bg = color_hex in [
                                    COLOR_AC225_RUN_EVG, COLOR_IN111_RUN_EVG,
                                    COLOR_PLACEHOLDER, COLOR_CARDINAL_TPI_NIOWAVE,
                                    COLOR_NMCTG, COLOR_PERCEPTIVE, COLOR_IN111_RUN_SRX
                                ]
                                text_color_rgb = RGBColor(0, 0, 0) if light_bg else RGBColor(255, 255, 255)
                                p.font.color.rgb = text_color_rgb
                            except Exception:
                                pass
                y_current += ACTIVITY_ROW_HEIGHT

            current_week_idx += 1

    # Save to bytes
    buffer = io.BytesIO()
    prs.save(buffer)
    ppt_data = buffer.getvalue()
    buffer.close()
    return ppt_data

def generate_excel_calendar(year, month, entries, week_action_rows):
    """Generate an Excel file with the calendar, including trailing dates, events, and colors."""
    from datetime import date as dt, timedelta
    import calendar
    import pandas as pd
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl import Workbook
    import io

    def prev_month(y, m):
        return (y - 1, 12) if m == 1 else (y, m - 1)

    def next_month(y, m):
        return (y + 1, 1) if m == 12 else (y, m + 1)

    # Build extended weeks (Mon–Sun)
    cal_raw = calendar.monthcalendar(year, month)
    valid_weeks = [week for week in cal_raw if any(d != 0 for d in week)]

    extended_weeks = []
    for week in valid_weeks:
        ref_day = next((d for d in week if d != 0), None)
        if ref_day is None:
            extended_weeks.append([None] * 7)
            continue
        ref_date = dt(year, month, ref_day)
        start_of_week = ref_date - timedelta(days=ref_date.weekday())  # Monday
        week_dates = [start_of_week + timedelta(days=i) for i in range(7)]
        extended_weeks.append(week_dates)

    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = f"{calendar.month_name[month]} {year}"

    # Set column widths
    for col_idx in range(1, 8):
        ws.column_dimensions[get_column_letter(col_idx)].width = 15

    # Header row
    day_headers = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    for col_idx, day_name in enumerate(day_headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = day_name
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Convert hex to openpyxl color (RRGGBB)
    def hex_to_xlsx(color_hex):
        return color_hex.lstrip('#').upper() if color_hex != "white" else "FFFFFF"

    # Start from row 2
    current_row = 2

    for week_idx, week_dates in enumerate(extended_weeks):
        # Date row
        for col_idx, dt_obj in enumerate(week_dates, 1):
            if dt_obj is None:
                continue
            cell = ws.cell(row=current_row, column=col_idx)
            cell.value = dt_obj.strftime("%b-%d")
            cell.font = Font(bold=True, size=11, color="000000")
            cell.fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Activity rows
        num_rows = week_action_rows.get(week_idx, 1)
        for row_offset in range(num_rows):
            current_row += 1
            for col_idx, dt_obj in enumerate(week_dates, 1):
                if dt_obj is None:
                    continue
                dkey = date_key(dt_obj.year, dt_obj.month, dt_obj.day, row_offset)
                entry = entries.get(dkey, "").strip()
                cell = ws.cell(row=current_row, column=col_idx)
                cell.value = entry
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.font = Font(size=10, bold=True)

                if entry:
                    color_hex = get_color(entry)
                    if color_hex != "white":
                        bg_color = hex_to_xlsx(color_hex)
                        text_color = "000000" if color_hex in [
                            COLOR_AC225_RUN_EVG, COLOR_IN111_RUN_EVG,
                            COLOR_PLACEHOLDER, COLOR_CARDINAL_TPI_NIOWAVE,
                            COLOR_NMCTG, COLOR_PERCEPTIVE, COLOR_IN111_RUN_SRX
                        ] else "FFFFFF"
                        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
                        cell.font = Font(size=10, bold=True, color=text_color)
                    else:
                        cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                else:
                    cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        current_row += 1  # Gap between weeks

    # Set row heights for readability
    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 15  # Date row
        # You can adjust per row if needed

    # Auto-fit columns (optional)
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 20)
        ws.column_dimensions[col_letter].width = adjusted_width

    # Save to bytes
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()



# =======================
# BOOT LOAD
# =======================
if "__boot_loaded__" not in ss:
    ss["__boot_loaded__"] = True
    try:
        latest_dir = load_latest_dir()
        fp = latest_dir / FILENAME
        entries, meta, week_action_rows, full_data = _try_load_from(fp)
        if entries is not None:
            ss.entries = entries
            if week_action_rows is not None:
                ss.week_action_rows = week_action_rows
            if full_data and "custom_legend_entries" in full_data:
                ss.custom_legend_entries = full_data["custom_legend_entries"]
            changed = _apply_meta_to_calendar(meta or {})
            _preload_widgets_from_entries()
            if changed:
                rerun()
    except Exception as e:
        ss["__boot_error__"] = str(e)

# =======================
# APPLY PENDING LOAD
# =======================
if ss.get("__pending_entries__") is not None:
    try:
        pending_entries = ss.get("__pending_entries__") or {}
        pending_meta = ss.get("__pending_meta__") or {}
        pending_week_action_rows = ss.get("__pending_week_action_rows__") or {}
        ss.entries = pending_entries
        if pending_week_action_rows:
            ss.week_action_rows = pending_week_action_rows
        changed = _apply_meta_to_calendar(pending_meta)
        _preload_widgets_from_entries()
        ss["__pending_entries__"] = None
        ss["__pending_meta__"] = None
        ss["__pending_week_action_rows__"] = None
        if changed:
            rerun()
    except Exception as e:
        ss["__boot_error__"] = str(e)

# =======================
# TITLE & NAV
# =======================
col_title_left, col_title_mid, col_title_right = st.columns([0.4, 4.3, 0.48])
with col_title_left:
    pass
with col_title_mid:
    st.markdown("<h1 style='text-align: center;'>Production Schedule Dashboard</h1>", unsafe_allow_html=True)
with col_title_right:
    if st.button("☰ Legend", key="legend_toggle"):
        ss.show_legend = not ss.show_legend
        rerun()

col_nav_left, col_nav_mid, col_nav_right = st.columns([0.48, 5.04, 0.48])
with col_nav_left:
    if st.button("← Prev", key="prev"):
        # Commit all widgets and autosave before changing month
        _commit_all_widgets_and_autosave()
        ss.current_month -= 1
        if ss.current_month < 1:
            ss.current_month = 12
            ss.current_year -= 1
        rerun()
with col_nav_mid:
    pass
with col_nav_right:
    if st.button("Next →", key="next"):
        # Commit all widgets and autosave before changing month
        _commit_all_widgets_and_autosave()
        ss.current_month += 1
        if ss.current_month > 12:
            ss.current_month = 1
            ss.current_year += 1
        rerun()

month_name = calendar.month_name[ss.current_month]
st.markdown(
    f"<div style='text-align: center; font-size: 26px; font-weight: bold; margin: 10px 0 20px 0;'>{month_name} {ss.current_year}</div>",
    unsafe_allow_html=True
)

# =======================
# EDITABLE LEGEND SIDEBAR (Integrated View)
# =======================
if ss.show_legend:
    st.markdown("""<style>.css-1d391kg { width: 500px !important; }</style>""", unsafe_allow_html=True)
    with st.sidebar:
        st.title("Legends")

        # Initialize custom legend list
        if 'custom_legend_entries' not in ss:
            ss.custom_legend_entries = []

        # All legends: built-in + custom
        all_legends = []

        # Add built-in legends (non-deletable)
        built_in_legends = [
            {"label": "Confirmed Patient", "description": "Confirmed patient dose scheduled", "color": COLOR_CONFIRMED, "builtin": True},
            {"label": "Placeholder Patient", "description": "Placeholder for expected patient dose", "color": COLOR_PLACEHOLDER, "builtin": True},
            {"label": "Shutdown", "description": "Equipment or facility shutdown", "color": COLOR_SHUTDOWN, "builtin": True},
            {"label": "Cardinal/TPI/Niowave", "description": "Ac225 production site activities", "color": COLOR_CARDINAL_TPI_NIOWAVE, "builtin": True},
            {"label": "BWXT Order", "description": "IN-111 Isotope", "color": COLOR_BWXT, "builtin": True},
            {"label": "AC225 Run-EVG", "description": "Scheduled production of Ac225 batches at Evergreen", "color": COLOR_AC225_RUN_EVG, "builtin": True},
            {"label": "IN111 Run-EVG", "description": "Scheduled production of In111 batches at Evergreen", "color": COLOR_IN111_RUN_EVG, "builtin": True},
            {"label": "AC225 Run-SRx", "description": "Scheduled production of Ac225 batches at Spectron Rx", "color": COLOR_AC225_RUN_SRX, "builtin": True},
            {"label": "IN111 Run-SRx", "description": "Scheduled production of In111 batches at Spectron Rx", "color": COLOR_IN111_RUN_SRX, "builtin": True},
            {"label": "NMCTG", "description": "Clinical Site Qualification Event by NMCTG", "color": COLOR_NMCTG, "builtin": True},
            {"label": "Perceptive", "description": "Clinical Site Qualification Event by Perceptive", "color": COLOR_PERCEPTIVE, "builtin": True},
        ]
        all_legends.extend(built_in_legends)
        all_legends.extend([{
            "label": item["label"],
            "description": item["description"],
            "color": item["color"],
            "builtin": False,
            "index": i
        } for i, item in enumerate(ss.custom_legend_entries)])

        # === Display All Legends Together ===
        for item in all_legends:
            cols = st.columns([4, 1])
            with cols[0]:
                text_color = "black" if sum(int(item['color'].lstrip('#')[i:i+2], 16) for i in (0,2,4)) > 300 else "white"
                st.markdown(f"""
                <div style="background-color: {item['color']}; padding: 10px; margin: 6px 0; border-radius: 6px; border: 1px solid #ddd;">
                    <div style="font-weight: 600; color: {text_color}; font-size: 13px;">{item['label']}</div>
                    <div style="color: {text_color}; font-size: 11px;">{item['description']}</div>
                </div>
                """, unsafe_allow_html=True)
            with cols[1]:
                if not item.get("builtin", False):
                    if st.button("🗑️", key=f"del_{item['index']}", help=f"Delete '{item['label']}'"):
                        ss.custom_legend_entries.pop(item['index'])
                        st.rerun()

        # === Add New Legend ===
        st.markdown("### ➕ Add New Legend")
        picked_color = st.color_picker("Choose color:", "#3366cc", key="new_legend_color")

        label_input = st.text_input(
            "Symbol/Label",
            placeholder= "Enter Symbol/Label"
        )

        desc_input = st.text_input(
            "Description",
            placeholder="Enter Description"
        )

        if st.button("➕ Add Legend Item"):
            if label_input.strip():
                ss.custom_legend_entries.append({
                    "label": label_input.strip(),
                    "description": desc_input.strip() if desc_input.strip() else "No description",
                    "color": picked_color
                })
                st.success(f"Added: {label_input}")
                st.rerun()
            else:
                st.warning("Label is required.")

        # Close button
        if st.button("✕ Close Legend", use_container_width=True, type="primary"):
            ss.show_legend = False
            st.rerun()

# =======================
# CALENDAR GRID
# =======================
import calendar
from datetime import datetime, date, timedelta

def prev_month(year, month):
    if month == 1:
        return year - 1, 12
    return year, month - 1

def next_month(year, month):
    if month == 12:
        return year + 1, 1
    return year, month + 1

def get_date_for_day(year, month, day):
    """Safely get a date, even if day is out of range."""
    try:
        return date(year, month, day)
    except ValueError:
        # Handle invalid day by adjusting month
        if day < 1:
            prev_y, prev_m = prev_month(year, month)
            last_day_prev = calendar.monthrange(prev_y, prev_m)[1]
            return date(prev_y, prev_m, last_day_prev + day)
        else:
            next_y, next_m = next_month(year, month)
            return date(next_y, next_m, day - calendar.monthrange(year, month)[1])
    except Exception:
        return None

# Get raw calendar (includes 0s for padding)
cal_raw = calendar.monthcalendar(ss.current_year, ss.current_month)
valid_weeks = [week for week in cal_raw if any(d != 0 for d in week)]

# Extend to include full weeks (even if they span previous/next months)
extended_weeks = []
for week in valid_weeks:
    extended_week = []
    for i, d in enumerate(week):
        if d == 0:
            # This day is in previous or next month
            # We need to infer the actual date
            # Get the first valid day in the week
            ref_day = None
            for x in week:
                if x != 0:
                    ref_day = x
                    break
            if ref_day is None:
                extended_week.append(None)
                continue

            # Find the actual date of that ref_day
            try:
                ref_date = date(ss.current_year, ss.current_month, ref_day)
                # Week starts on Monday (0), so calculate offset
                weekday_of_ref = ref_date.weekday()  # Mon=0, Sun=6
                target_date = ref_date - timedelta(days=weekday_of_ref - i)
                extended_week.append(target_date)
            except Exception:
                extended_week.append(None)
        else:
            try:
                dt = date(ss.current_year, ss.current_month, d)
                extended_week.append(dt)
            except Exception:
                extended_week.append(None)
    extended_weeks.append(extended_week)

# Sync widgets before rendering
_ensure_rows_for_current_month(valid_weeks)
_sync_widgets_with_entries()

# Day of week headers
dow_columns = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
header_cols = st.columns(7)
for i, day_name in enumerate(dow_columns):
    with header_cols[i]:
        st.markdown(f"<div style='font-size:20px; font-weight:800; text-align:center'>{day_name}</div>", unsafe_allow_html=True)
st.markdown("---")

# Render each week
for week_idx, week_dates in enumerate(extended_weeks):
    # Date row (e.g., Jul-30, Aug-01)
    cols = st.columns(7)
    for i, dt in enumerate(week_dates):
        with cols[i]:
            if dt is None:
                st.write("")
            else:
                label = f"{dt.strftime('%b-%d')}"
                st.markdown(f"<div style='font-size:16px; font-weight:600; text-align:center'>{label}</div>", unsafe_allow_html=True)

    # Get number of action rows for this week
    week_key = f"{ss.current_year}-{ss.current_month}_{week_idx}"
    num_rows = ss.week_action_rows.get(week_key, 1)

    # Render action rows
    for row_idx in range(num_rows):
        input_cols = st.columns(7)
        for day_idx, dt in enumerate(week_dates):
            with input_cols[day_idx]:
                if dt is None:
                    st.write("")
                else:
                    # Use actual date (could be from prev/next month)
                    dkey = date_key(dt.year, dt.month, dt.day, row_idx)
                    current_val = ss.entries.get(dkey, "")
                    bg_color = get_color(current_val)

                    # Determine text color
                    if bg_color == "white":
                        text_color = "black"
                    else:
                        light_bgs = [
                            COLOR_AC225_RUN_EVG,
                            COLOR_IN111_RUN_EVG,
                            COLOR_PLACEHOLDER,
                            COLOR_CARDINAL_TPI_NIOWAVE,
                            COLOR_NMCTG,
                            COLOR_PERCEPTIVE,
                            COLOR_IN111_RUN_SRX,
                        ]
                        text_color = "black" if bg_color in light_bgs else "white"

                    label_str = f"cell_{dkey}"
                    widget_key = f"cell_widget_{dkey}"
                    if widget_key not in ss:
                        ss[widget_key] = current_val

                    # Dynamic CSS
                    st.markdown(
                        f"""
                        <style>
                        div[data-testid="stTextInput"] input[aria-label="{label_str}"] {{
                            background-color: {bg_color} !important;
                            color: {text_color} !important;
                            border: 0 !important;
                            height: 40px !important;
                            line-height: 40px !important;
                            text-align: center !important;
                            font-weight: 500 !important;
                            border-radius: 4px !important;
                            box-shadow: none !important;
                            padding: 0 8px !important;
                            margin: 0 !important;
                        }}
                        div[data-testid="stTextInput"] label {{
                            display: none !important;
                        }}
                        div[data-testid="stTextInput"] > div {{
                            margin: 0 !important;
                            padding: 0 !important;
                        }}
                        </style>
                        """,
                        unsafe_allow_html=True,
                    )

                    st.text_input(
                        label=label_str,
                        key=widget_key,
                        label_visibility="collapsed",
                        placeholder="Add event",
                        on_change=lambda dk=dkey, wk=widget_key: _commit_and_autosave(dk, wk),
                    )
    st.markdown('<div style="margin: 20px 0;"></div>', unsafe_allow_html=True)

# =======================
# ADD/REMOVE ROW CONTROL
# =======================
st.markdown("---")
st.markdown("### Add/Remove Row of events for the Week")
week_options = [f"Week {i+1}" for i in range(len(valid_weeks))]
selected_week = st.selectbox("Select Week Number to add/remove a row of events:", options=week_options, key="select_week")

col_add_remove_left, col_add_remove_right = st.columns(2)

with col_add_remove_left:
    if st.button("➕ Add Row to the Selected Week", use_container_width=True):
        selected_index = int(selected_week.split(" ")[1]) - 1
        ss.week_action_rows[f"{ss.current_year}-{ss.current_month}_{selected_index}"] = ss.week_action_rows.get(f"{ss.current_year}-{ss.current_month}_{selected_index}", 1) + 1
        # Autosave the row changes
        try:
            _autosave_now()
            ss["__autosave_ok__"] = True
            ss["__autosave_error__"] = ""
        except Exception as e:
            ss["__autosave_ok__"] = False
            ss["__autosave_error__"] = str(e)
        rerun()

with col_add_remove_right:
    if st.button("➖ Delete Row from the Selected Week", use_container_width=True):
        selected_index = int(selected_week.split(" ")[1]) - 1
        week_key = f"{ss.current_year}-{ss.current_month}_{selected_index}"
        current_rows = ss.week_action_rows.get(week_key, 1)
        
        # Only delete if there's more than one row
        if current_rows > 1:
            # Check if the bottom row is empty
            bottom_row_empty = True
            week = valid_weeks[selected_index]
            
            for day in week:
                if day != 0:  # Skip empty days
                    dkey = date_key(ss.current_year, ss.current_month, day, current_rows - 1)
                    if ss.entries.get(dkey, "").strip():
                        bottom_row_empty = False
                        break
            
            if bottom_row_empty:
                ss.week_action_rows[week_key] = current_rows - 1
                # Clean up any entries from the deleted row
                for day in week:
                    if day != 0:
                        dkey = date_key(ss.current_year, ss.current_month, day, current_rows - 1)
                        if dkey in ss.entries:
                            del ss.entries[dkey]
                # Autosave the row changes
                try:
                    _autosave_now()
                    ss["__autosave_ok__"] = True
                    ss["__autosave_error__"] = ""
                except Exception as e:
                    ss["__autosave_ok__"] = False
                    ss["__autosave_error__"] = str(e)
                rerun()
            else:
                st.warning("Cannot delete row: The bottom row contains events. Please clear the events first.")
        else:
            st.warning("Cannot delete row: This is the only row remaining for this week.")

# =======================
# POWERPOINT DOWNLOAD
# =======================
st.markdown("---")
st.markdown("### 📊 Export Production Schedule Dashboard as PowerPoint")
st.markdown("Generate a PowerPoint presentation of your current month's production schedule")

if st.button("Generate PowerPoint", key="ppt_download", type="primary"):
    with st.spinner("🔄 Generating PowerPoint..."):
        try:
            month_week_rows = get_week_action_rows_for_month(ss.current_year, ss.current_month)
            ppt_data = generate_ppt_calendar(ss.current_year, ss.current_month, ss.entries, month_week_rows)
            
            st.success("✅ PowerPoint generated successfully!")
            
            st.download_button(
                label="Download PowerPoint",
                data=ppt_data,
                file_name=f"production_schedule_{ss.current_year}_{ss.current_month:02d}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="ppt_download_btn"
            )
            
        except Exception as e:
            st.error(f"❌ Error generating PowerPoint: {str(e)}")
            st.info("💡 Make sure you have the required dependencies installed: `pip install python-pptx`")

# =======================
# PDF DOWNLOAD
# =======================
st.markdown("---")
st.markdown("### 📄 Export Production Schedule Dashboard as PDF")
st.markdown("Generate a PDF version for the current months production schedule")

if st.button("Generate PDF", key="pdf_download", type="primary"):
    with st.spinner("🔄 Generating PDF..."):
        try:
            month_week_rows = get_week_action_rows_for_month(ss.current_year, ss.current_month)
            pdf_data = generate_pdf_calendar(ss.current_year, ss.current_month, ss.entries, month_week_rows)
            
            st.success("✅ PDF generated successfully!")
            
            st.download_button(
                label="Download PDF",
                data=pdf_data,
                file_name=f"production_schedule_{ss.current_year}_{ss.current_month:02d}.pdf",
                mime="application/pdf",
                key="pdf_download_btn"
            )
            
        except Exception as e:
            st.error(f"❌ Error generating PDF: {str(e)}")
            st.info("💡 Make sure you have the required dependencies installed: `pip install reportlab`")


# =======================
# EXCEL DOWNLOAD
# =======================
st.markdown("---")
st.markdown("### 📥 Export Production Schedule Dashboard as Excel")
st.markdown("Generate an Excel file for the current months production schedule dashboard")

if st.button("Generate Excel", key="excel_download", type="primary"):
    with st.spinner("🔄 Generating Excel..."):
        try:
            month_week_rows = get_week_action_rows_for_month(ss.current_year, ss.current_month)
            excel_data = generate_excel_calendar(ss.current_year, ss.current_month, ss.entries, month_week_rows)
            
            st.success("✅ Excel generated successfully!")
            
            st.download_button(
                label="Download Excel",
                data=excel_data,
                file_name=f"production_schedule_{ss.current_year}_{ss.current_month:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="excel_download_btn"
            )
            
        except Exception as e:
            st.error(f"❌ Error generating Excel: {str(e)}")
            st.info("💡 Make sure you have the required dependencies installed: `pip install pandas openpyxl`")


# =======================
# RELOAD PREVIOUS
# =======================
st.markdown("---")
st.markdown("### 🔄 Reload Previous Production Schedule")
prefill_dir = str(load_latest_dir())
raw_dir_input = st.text_input(
    "Absolute directory (auto-loads if a saved file exists):",
    value=prefill_dir,
    key="dir_input",
    help="Example (macOS): /Users/you/Schedules | Example (Windows): D:\\Schedules"
)
entered_dir = _sanitize_dir(raw_dir_input)
if str(entered_dir) and str(entered_dir) != prefill_dir:
    save_latest_dir(str(entered_dir))
    new_fp = entered_dir / FILENAME
    try:
        if new_fp.exists():
            data = json.loads(new_fp.read_text(encoding="utf-8"))
            if isinstance(data, dict) and "entries" in data:
                ss["__pending_entries__"] = data.get("entries", {}) or {}
                ss["__pending_meta__"] = data.get("meta") or {}
                ss["__pending_week_action_rows__"] = data.get("week_action_rows", {}) or {}
                st.success(f"Loaded schedule from: {new_fp}")
            else:
                st.info(f"No saved schedule found at: {new_fp}")
        else:
            st.info(f"No saved schedule found at: {new_fp}")
    except Exception as e:
        st.error(f"Failed to read: {e}")
    rerun()

latest_dir = load_latest_dir()
file_path = latest_dir / FILENAME
st.markdown(
    f"<pre style='text-align:left; white-space:pre-wrap; margin:0'>Working file:\n{file_path}</pre>",
    unsafe_allow_html=True
)

# Status display
if ss["__autosave_ok__"]:
    st.caption("✅ Autosaved")
elif ss["__autosave_error__"]:
    st.caption(f"⚠️ Autosave error: {ss['__autosave_error__']}")

# Show any boot errors
if ss.get("__boot_error__"):
    st.error(f"⚠️ Load error: {ss['__boot_error__']}")