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
COLOR_PV = "#3CD63CC0"
COLOR_SRX = "#3ACCC0"
COLOR_PERCEPTIVE = "#75A06B"
COLOR_BWXT = "#3D6E34"

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
ss.setdefault("week_action_rows", {})          # week index -> row count (per month view)
ss.setdefault("entries", {})                   # "YYYY-MM-DD_row" -> text
ss.setdefault("__autosave_ok__", False)
ss.setdefault("__autosave_error__", "")
ss.setdefault("show_legend", False)  # legend drawer visibility

# pending load buffer (used when dir changes or on initial boot load)
ss.setdefault("__pending_entries__", None)
ss.setdefault("__pending_meta__", None)

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
    if "shutdown" in lower:
        return COLOR_SHUTDOWN
    if lower == "ac225 run-evg":
        return COLOR_AC225_RUN_EVG
    if lower == "in111 run-evg":
        return COLOR_IN111_RUN_EVG
    if lower == "ac225 run-srx":
        return COLOR_AC225_RUN_SRX
    if lower == "in111 run-srx":
        return COLOR_IN111_RUN_SRX
    if lower.startswith(("cardinal ac225", "tpi ac225", "niowave ac225")):
        return COLOR_CARDINAL_TPI_NIOWAVE
    if "nmctg" in lower:
        return COLOR_NMCTG
    if lower.startswith("32008-p2"):
        return COLOR_PLACEHOLDER
    if lower.startswith(("3200")) and "p2" not in lower:
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

def _ensure_rows_for_current_month(valid_weeks):
    """
    Make sure each week renders enough rows to display ALL saved entries for this month.
    Looks at saved date-based keys (YYYY-MM-DD_row), finds what week each day belongs to,
    and sets week_action_rows[week_idx] to max(row_idx)+1 for that week (but never less than current).
    """
    # Build day->week_idx map aligned with valid_weeks indexing
    day_to_week = {}
    for w_idx, week in enumerate(valid_weeks):
        for d in week:
            if d != 0:
                day_to_week[d] = w_idx

    # Compute required rows by scanning entries for this month
    required_rows = {}  # week_idx -> max rows needed
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

    # Ensure week_action_rows has at least what's required (default 1)
    for w_idx in range(len(valid_weeks)):
        current = ss.week_action_rows.get(w_idx, 1)
        ss.week_action_rows[w_idx] = max(current, required_rows.get(w_idx, 1))

def _try_load_from(path: Path):
    """Read JSON file and return (entries, meta) or (None, None) if not valid."""
    try:
        if not path.exists():
            return None, None
        data = json.loads(path.read_text(encoding="utf-8")) or {}
        if isinstance(data, dict) and "entries" in data:
            entries = data.get("entries", {}) or {}
            meta = (data.get("meta") if isinstance(data, dict) else None) or {}
            return entries, meta
    except Exception:
        pass
    return None, None

def generate_pdf_calendar(year, month, entries, week_action_rows):
    """Generate a PDF calendar for the specified month with activities and colors."""
    # Create a buffer to store the PDF
    buffer = io.BytesIO()
    
    # Create the PDF document
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    story = []
    
    # Get styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=1  # Center alignment
    )
    
    month_style = ParagraphStyle(
        'MonthStyle',
        parent=styles['Heading2'],
        fontSize=20,
        spaceAfter=20,
        alignment=1
    )
    
    # Add title
    story.append(Paragraph("Production Schedule Dashboard", title_style))
    story.append(Spacer(1, 20))
    
    # Add month and year
    month_name = calendar.month_name[month]
    story.append(Paragraph(f"{month_name} {year}", month_style))
    story.append(Spacer(1, 20))
    
    # Get calendar data
    cal_raw = calendar.monthcalendar(year, month)
    valid_weeks = [week for week in cal_raw if any(d != 0 for d in week)]
    
    # Day headers
    day_headers = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    
    # Create calendar table data with separate rows for dates and activities
    table_data = []
    
    # Add day headers
    table_data.append(day_headers)
    
    # Add calendar weeks with separate rows for dates and activities
    for week_idx, week in enumerate(valid_weeks):
        # Date row (just day numbers)
        date_row = []
        for day in week:
            if day == 0:
                date_row.append("")
            else:
                date_row.append(str(day))
        table_data.append(date_row)
        
        # Activity rows for this week
        num_rows = week_action_rows.get(week_idx, 1)
        for row_idx in range(num_rows):
            activity_row = []
            for day in week:
                if day == 0:
                    activity_row.append("")
                else:
                    dkey = date_key(year, month, day, row_idx)
                    entry = entries.get(dkey, "")
                    activity_row.append(entry.strip() if entry.strip() else "")
            table_data.append(activity_row)
    
    # Calculate row heights: header + date rows + activity rows
    num_weeks = len(valid_weeks)
    total_rows = 1 + num_weeks + sum(week_action_rows.get(i, 1) for i in range(num_weeks))
    
    # Create row heights: header (0.4"), date rows (0.3"), activity rows (0.4")
    row_heights = [0.4]  # Header
    for week_idx in range(num_weeks):
        row_heights.append(0.3)  # Date row
        num_activity_rows = week_action_rows.get(week_idx, 1)
        row_heights.extend([0.4] * num_activity_rows)  # Activity rows
    
    # Create the table
    table = Table(table_data, colWidths=[1.1*inch]*7, rowHeights=[h*inch for h in row_heights])
    
    # Style the table
    table_style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Day headers
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # Day headers background
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # Day headers text
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ])
    
    # Style date rows (bold, light background)
    current_row = 1
    for week_idx in range(num_weeks):
        date_row_idx = current_row
        table_style.add('FONTNAME', (0, date_row_idx), (-1, date_row_idx), 'Helvetica-Bold')
        table_style.add('FONTSIZE', (0, date_row_idx), (-1, date_row_idx), 14)
        table_style.add('BACKGROUND', (0, date_row_idx), (-1, date_row_idx), colors.lightgrey)
        current_row += 1
        
        # Style activity rows with colors
        num_activity_rows = week_action_rows.get(week_idx, 1)
        for activity_row_idx in range(num_activity_rows):
            for day_idx, day in enumerate(week):
                if day != 0:
                    dkey = date_key(year, month, day, activity_row_idx)
                    entry = entries.get(dkey, "")
                    if entry.strip():
                        # Get color for this activity
                        color = get_color(entry)
                        # Convert hex color to reportlab color
                        if color != "white":
                            try:
                                # Convert hex to RGB
                                hex_color = color.lstrip('#')
                                r = int(hex_color[0:2], 16) / 255.0
                                g = int(hex_color[2:4], 16) / 255.0
                                b = int(hex_color[4:6], 16) / 255.0
                                reportlab_color = colors.Color(r, g, b)
                                table_style.add('BACKGROUND', (day_idx, current_row), (day_idx, current_row), reportlab_color)
                                
                                # Set text color based on background brightness
                                if color in ["#A2EBCD", "#F1E183", "#F1E429", "#0ABB21", "#BCA6CA", "#75A06B"]:
                                    table_style.add('TEXTCOLOR', (day_idx, current_row), (day_idx, current_row), colors.black)
                                else:
                                    table_style.add('TEXTCOLOR', (day_idx, current_row), (day_idx, current_row), colors.white)
                            except:
                                pass  # Fallback to default if color conversion fails
            
            current_row += 1
    
    table.setStyle(table_style)
    story.append(table)
    story.append(Spacer(1, 30))
    
    # Add legend with colors
    legend_title = Paragraph("Legend", styles['Heading2'])
    story.append(legend_title)
    story.append(Spacer(1, 15))
    
    # Legend data with colors
    legend_data = [
        ["Confirmed Patient", "Confirmed patient dose scheduled", COLOR_CONFIRMED],
        ["Placeholder Patient", "Placeholder for expected patient dose", COLOR_PLACEHOLDER],
        ["Shutdown", "Equipment or facility shutdown", COLOR_SHUTDOWN],
        ["Cardinal/TPI/Niowave", "Ac225 production site activities", COLOR_CARDINAL_TPI_NIOWAVE],
        ["BWXT Order", "IN-111 Isotope", COLOR_BWXT],
        ["AC225 Run-EVG", "Scheduled production of Ac225 batches at Evergreen", COLOR_AC225_RUN_EVG],
        ["IN111 Run-EVG", "Scheduled production of In111 batches at Evergreen", COLOR_IN111_RUN_EVG],
        ["AC225 Run-SRx", "Scheduled production of Ac225 batches at Spectron Rx", COLOR_AC225_RUN_SRX],
        ["IN111 Run-SRx", "Scheduled production of In111 batches at Spectron Rx", COLOR_IN111_RUN_SRX],
        ["NMCTG", "Clinical Site Qualification Event by NMCTG", COLOR_NMCTG],
        ["Perceptive", "Clinical Site Qualification Event by Perceptive", COLOR_PERCEPTIVE],
    ]
    
    # Create legend table with colored backgrounds
    legend_table_data = []
    for item in legend_data:
        legend_table_data.append([item[0], item[1]])
    
    legend_table = Table(legend_table_data, colWidths=[2*inch, 4*inch])
    legend_style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ])
    
    # Apply colors to legend rows
    for i, item in enumerate(legend_data):
        try:
            color = item[2]
            hex_color = color.lstrip('#')
            r = int(hex_color[0:2], 16) / 255.0
            g = int(hex_color[2:4], 16) / 255.0
            b = int(hex_color[4:6], 16) / 255.0
            reportlab_color = colors.Color(r, g, b)
            legend_style.add('BACKGROUND', (0, i), (-1, i), reportlab_color)
            
            # Set text color based on background brightness
            if color in ["#A2EBCD", "#F1E183", "#F1E429", "#0ABB21", "#BCA6CA", "#75A06B"]:
                legend_style.add('TEXTCOLOR', (0, i), (-1, i), colors.black)
            else:
                legend_style.add('TEXTCOLOR', (0, i), (-1, i), colors.white)
        except:
            pass  # Fallback to default if color conversion fails
    
    legend_table.setStyle(legend_style)
    story.append(legend_table)
    
    # Build the PDF
    doc.build(story)
    
    # Get the PDF data
    pdf_data = buffer.getvalue()
    buffer.close()
    
    return pdf_data

# =======================
# BOOT LOAD (auto-load last saved + jump calendar to saved month/year)
# =======================
if "__boot_loaded__" not in ss:
    ss["__boot_loaded__"] = True
    try:
        latest_dir = load_latest_dir()
        fp = latest_dir / FILENAME
        entries, meta = _try_load_from(fp)
        if entries is not None:
            ss.entries = entries
            changed = _apply_meta_to_calendar(meta or {})
            _preload_widgets_from_entries()
            if changed:
                rerun()
    except Exception as e:
        ss["__boot_error__"] = str(e)

# =======================
# APPLY PENDING LOAD (from directory change) BEFORE widgets render
# =======================
if ss.get("__pending_entries__") is not None:
    try:
        pending_entries = ss.get("__pending_entries__") or {}
        pending_meta = ss.get("__pending_meta__") or {}
        ss.entries = pending_entries
        changed = _apply_meta_to_calendar(pending_meta)
        _preload_widgets_from_entries()
        ss["__pending_entries__"] = None
        ss["__pending_meta__"] = None
        if changed:
            rerun()
    except Exception as e:
        ss["__boot_error__"] = str(e)

# =======================
# TITLE & NAV
# =======================
# Title and legend button in the same row
col_title_left, col_title_mid, col_title_right = st.columns([0.4, 4.3, 0.48])
with col_title_left:
    # Empty space
    pass
with col_title_mid:
    st.markdown("<h1 style='text-align: center;'>Production Schedule Dashboard</h1>", unsafe_allow_html=True)
with col_title_right:
    # Legend toggle aligned with title
    if st.button("☰ Legend", key="legend_toggle"):
        ss.show_legend = not ss.show_legend
        rerun()

# Navigation buttons aligned horizontally
col_nav_left, col_nav_mid, col_nav_right = st.columns([0.48, 5.04, 0.48])
with col_nav_left:
    if st.button("← Prev", key="prev"):
        ss.current_month -= 1
        if ss.current_month < 1:
            ss.current_month = 12
            ss.current_year -= 1
        rerun()
with col_nav_mid:
    # Empty space
    pass
with col_nav_right:
    if st.button("Next →", key="next"):
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

# PDF Download Section
st.markdown("---")
col_pdf_left, col_pdf_mid, col_pdf_right = st.columns([1, 2, 1])
with col_pdf_mid:
    st.markdown("### 📄 Download Calendar as PDF")
    if st.button("📥 Generate & Download PDF", key="pdf_download", type="primary", use_container_width=True):
        try:
            # Generate PDF
            pdf_data = generate_pdf_calendar(ss.current_year, ss.current_month, ss.entries, ss.week_action_rows)
            
            # Create download button
            st.download_button(
                label="💾 Download PDF",
                data=pdf_data,
                file_name=f"production_schedule_{ss.current_year}_{ss.current_month:02d}.pdf",
                mime="application/pdf",
                key="pdf_download_btn",
                use_container_width=True
            )
            
            st.success("✅ PDF generated successfully! Click the download button above to save it.")
            
        except Exception as e:
            st.error(f"❌ Error generating PDF: {str(e)}")
            st.info("💡 Make sure you have the required dependencies installed: `pip install reportlab`")

st.markdown("---")

# =======================
# LEGEND SIDEBAR (USING STREAMLIT NATIVE SIDEBAR)
# =======================
if ss.show_legend:
    # Custom CSS to extend sidebar width
    st.markdown("""
    <style>
    .css-1d391kg {
        width: 500px !important;
    }
    .css-1d391kg > div {
        width: 500px !important;
    }
    </style>
    """, unsafe_allow_html=True)
    
    with st.sidebar:
        st.title("Legend")
        
        # Create a legend table using Streamlit components
        legend_data = [
            {"Symbol/Label": "Confirmed Patient", "Description": "Confirmed patient dose scheduled", "Color": COLOR_CONFIRMED},
            {"Symbol/Label": "Placeholder Patient", "Description": "Placeholder for expected patient dose", "Color": COLOR_PLACEHOLDER},
            {"Symbol/Label": "Shutdown", "Description": "Equipment or facility shutdown", "Color": COLOR_SHUTDOWN},
            {"Symbol/Label": "Cardinal/TPI/Niowave", "Description": "Ac225 production site activities", "Color": COLOR_CARDINAL_TPI_NIOWAVE},
            {"Symbol/Label": "BWXT Order", "Description": "IN-111 Isotope", "Color": COLOR_BWXT},
            {"Symbol/Label": "AC225 Run-EVG", "Description": "Scheduled production of Ac225 batches at Evergreen", "Color": COLOR_AC225_RUN_EVG},
            {"Symbol/Label": "IN111 Run-EVG", "Description": "Scheduled production of In111 batches at Evergreen", "Color": COLOR_IN111_RUN_EVG},
            {"Symbol/Label": "AC225 Run-SRx", "Description": "Scheduled production of Ac225 batches at Spectron Rx", "Color": COLOR_AC225_RUN_SRX},
            {"Symbol/Label": "IN111 Run-SRx", "Description": "Scheduled production of In111 batches at Spectron Rx", "Color": COLOR_IN111_RUN_SRX},
            {"Symbol/Label": "NMCTG", "Description": "Clinical Site Qualification Event by NMCTG", "Color": COLOR_NMCTG},
            {"Symbol/Label": "Perceptive", "Description": "Clinical Site Qualification Event by Perceptive", "Color": COLOR_PERCEPTIVE},
        ]
        
        # Display legend items with colored row backgrounds
        for item in legend_data:
            # Determine text color based on background color brightness
            # Use white text for dark backgrounds, dark text for light backgrounds
            text_color = "#000000" if item['Color'] in [COLOR_AC225_RUN_EVG, COLOR_IN111_RUN_EVG, COLOR_PLACEHOLDER, COLOR_CARDINAL_TPI_NIOWAVE, COLOR_NMCTG, COLOR_PERCEPTIVE] else "#FFFFFF"
            
            # Create a colored row background with stacked layout for better text wrapping
            st.markdown(f"""
             <div style="
                 background-color: {item['Color']}; 
                 padding: 12px 16px;
                 margin: 8px 0;
                 border-radius: 8px;
                 border: 1px solid #ddd;
             ">
                 <div style="
                     display: flex;
                     flex-direction: column;
                     gap: 6px;
                 ">
                     <div style="
                         font-weight: 600;
                         color: {text_color};
                         font-size: 13px;
                         line-height: 1.2;
                         word-wrap: break-word;
                         overflow-wrap: break-word;
                     ">{item['Symbol/Label']}</div>
                     <div style="
                         color: {text_color};
                         font-size: 11px;
                         line-height: 1.3;
                         word-wrap: break-word;
                         overflow-wrap: break-word;
                     ">{item['Description']}</div>
                 </div>
             </div>
             """, unsafe_allow_html=True)
        
        # Close button at the bottom
        if st.button("✕ Close Legend", key="close_legend_btn", use_container_width=True, type="primary"):
            ss.show_legend = False
            rerun()

# =======================
# CALENDAR GRID
# =======================
cal_raw = calendar.monthcalendar(ss.current_year, ss.current_month)
valid_weeks = [week for week in cal_raw if any(d != 0 for d in week)]

# Ensure we render enough rows for all saved entries in this month (but don't shrink user-added rows)
_ensure_rows_for_current_month(valid_weeks)

# headers
dow_columns = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
header_cols = st.columns(7)
for i, day_name in enumerate(dow_columns):
    with header_cols[i]:
        st.markdown(f"<div style='font-size:20px; font-weight:800; text-align:center'>{day_name}</div>", unsafe_allow_html=True)
st.markdown("---")

for week_idx, week in enumerate(valid_weeks):
    # date labels
    cols = st.columns(7)
    for i, d in enumerate(week):
        with cols[i]:
            if d == 0:
                st.write("")
            else:
                month_abbr = calendar.month_abbr[ss.current_month]
                st.markdown(
                    f"<div style='font-size:16px; font-weight:600; text-align:center'>{month_abbr}-{d:02d}</div>",
                    unsafe_allow_html=True
                )

    # default at least 1
    if week_idx not in ss.week_action_rows:
        ss.week_action_rows[week_idx] = 1
    num_rows = ss.week_action_rows[week_idx]

    # rows of editable cells (keys bound to REAL DATE + row)
    for row_idx in range(num_rows):
        input_cols = st.columns(7)
        for day_idx, d in enumerate(week):
            with input_cols[day_idx]:
                if d == 0:
                    st.write("")
                else:
                    dkey = date_key(ss.current_year, ss.current_month, d, row_idx)
                    current_val = ss.entries.get(dkey, "")
                    color = get_color(current_val)

                    label_str = f"cell_{dkey}"           # aria-label
                    widget_key = f"cell_widget_{dkey}"   # widget identity

                    # Per-cell CSS (Mac/Chrome friendly)
                    st.markdown(f"""
                    <style>
                    div[data-testid="stTextInput"] input[aria-label="{label_str}"] {{
                        background-color: {color} !important;
                        border: 0 !important;
                        height: 40px !important;
                        line-height: 40px !important;
                        text-align: center !important;
                        font-weight: 500 !important;
                        border-radius: 4px !important;
                        box-shadow: none !important;
                        padding: 0 8px !important;
                        color: black !important;
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
                    """, unsafe_allow_html=True)

                    # Ensure widget has initial value BEFORE it is instantiated
                    if widget_key not in ss:
                        ss[widget_key] = current_val

                    # Editable input — Enter commits & AUTOSAVES to *that date*
                    st.text_input(
                        label=label_str,
                        key=widget_key,
                        label_visibility="collapsed",
                        placeholder="Add event",
                        on_change=lambda dk=dkey, wk=widget_key: _commit_and_autosave(dk, wk),
                    )

                    # Keep store in sync so color reflects latest
                    ss.entries[dkey] = ss.get(widget_key, "")

    st.markdown('<div style="margin: 20px 0;"></div>', unsafe_allow_html=True)

# =======================
# ADD ROW CONTROL
# =======================
st.markdown("---")
st.markdown("### Add Row of events for the Week")
week_options = [f"Week {i+1}" for i in range(len(valid_weeks))]
selected_week = st.selectbox(
    "Select Week Number to add a row of events:",
    options=week_options,
    key="select_week"
)
if st.button("Add Row to the Selected Week"):
    selected_index = int(selected_week.split(" ")[1]) - 1
    ss.week_action_rows[selected_index] = ss.week_action_rows.get(selected_index, 1) + 1
    rerun()

# =======================
# RELOAD PREVIOUS (BOTTOM)
# =======================
st.markdown("---")
st.markdown("### 🔄 Reload Previous Production Schedule")

prefill_dir = str(load_latest_dir())
raw_dir_input = st.text_input(
    "Absolute directory (auto-loads if a saved file exists):",
    value=prefill_dir,
    key="dir_input",
    help="Example (macOS): /Users/you/Schedules   |   Example (Windows): D:\\Schedules"
)
entered_dir = _sanitize_dir(raw_dir_input)

if str(entered_dir) and str(entered_dir) != prefill_dir:
    # Remember new dir and attempt to auto-load from it
    save_latest_dir(str(entered_dir))
    new_fp = entered_dir / FILENAME
    try:
        data = json.loads(new_fp.read_text(encoding="utf-8")) if new_fp.exists() else None
        if isinstance(data, dict) and "entries" in data:
            ss["__pending_entries__"] = data.get("entries", {}) or {}
            ss["__pending_meta__"] = (data.get("meta") or {})
            st.success(f"Loaded schedule from: {new_fp}")
        else:
            ss["__pending_entries__"] = None
            ss["__pending_meta__"] = None
            st.info(f"No saved schedule found at: {new_fp}")
    except Exception as e:
        ss["__pending_entries__"] = None
        ss["__pending_meta__"] = None
        st.error(f"Failed to read: {e}")
    rerun()

latest_dir = load_latest_dir()
file_path = latest_dir / FILENAME
st.markdown(
    f"<pre style='text-align:left; white-space:pre-wrap; margin:0'>Working file:\n{file_path}</pre>",
    unsafe_allow_html=True
)

# Optional tiny autosave feedback
if ss["__autosave_ok__"]:
    st.caption("✅ Autosaved")
elif ss["__autosave_error__"]:
    st.caption(f"⚠️ Autosave error: {ss['__autosave_error__']}")

