import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date, timedelta
from io import BytesIO
import html

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSHEETS_AVAILABLE = True
except ImportError:
    GSHEETS_AVAILABLE = False

NOTES_FILE = os.path.join(os.path.dirname(__file__), "submission_notes.json")

SUBMISSION_COLS = [
    "submission_id", "Name", "Email", "Phone", "Submission Date", "Submission Time",
    "Sector / Industry", "Primary Company (Ticker)", "Market Capitalization",
    "12-Month Target Price", "Professional Bio", "Your Edge",
    "Investment Summary", "Model or Supporting Materials Work",
]
NOTE_COLS = [
    "submission_id", "status", "meeting_date", "followup_date",
    "meeting_notes", "followup_notes", "starred", "last_updated", "notes_log",
]

STATUS_ICONS = {
    "New": "", "Meeting Scheduled": "🟡", "Met": "🟢",
    "Pass": "🔴", "Follow-up": "🟠",
}
STATUS_OPTIONS = list(STATUS_ICONS.keys())

LOGO_URL = "https://custom-images.strikinglycdn.com/res/hrscywv4p/image/upload/c_limit,fl_lossy,h_630,w_1200,f_auto,q_auto/3075/614271_402360.jpeg"

st.set_page_config(
    page_title="FirstWave | SIN Tracker", page_icon="🌊",
    layout="wide", initial_sidebar_state="expanded",
)
st.markdown("""
<style>
  /* ── Hide top-right toolbar & main menu ── */
  [data-testid="stToolbarActions"] { display: none !important; }
  #MainMenu { display: none !important; }
  [data-testid="stMainMenuButton"] { display: none !important; }
  header[data-testid="stHeader"] { background: transparent !important; }
  /* Hide heading anchor link icons */
  [data-testid="stHeadingActionElements"],
  [data-testid="stHeadingActionElements"] * { display: none !important; visibility: hidden !important; opacity: 0 !important; }
  [data-testid="StyledLinkIconContainer"] { display: none !important; visibility: hidden !important; }
  h1 a, h2 a, h3 a { display: none !important; pointer-events: none !important; }
  h1 svg, h2 svg, h3 svg { display: none !important; }
  /* Remove blue hyperlink styling everywhere in main content */
  .stMarkdown a[href^="mailto"],
  .stMarkdown a { color: inherit !important; text-decoration: none !important; pointer-events: none !important; cursor: default !important; }
  /* No pointer cursor on any non-button element */
  p a, span a, div a:not([data-testid]) { cursor: default !important; }

  /* ── Sidebar background & fixed padding ── */
  section[data-testid="stSidebar"],
  section[data-testid="stSidebar"] > div:first-child,
  section[data-testid="stSidebar"] > div {
      background: linear-gradient(175deg, #0b1220 0%, #1a2744 100%) !important;
  }
  section[data-testid="stSidebar"] > div:first-child {
      padding-left: 1rem !important;
      padding-right: 1rem !important;
      padding-top: 1rem !important;
  }

  /* Sidebar text */
  section[data-testid="stSidebar"] p,
  section[data-testid="stSidebar"] label,
  section[data-testid="stSidebar"] .stMarkdown,
  section[data-testid="stSidebar"] small {
      color: rgba(255,255,255,0.82) !important;
  }
  section[data-testid="stSidebar"] span { color: rgba(255,255,255,0.82) !important; }
  section[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.15) !important; }

  /* Sidebar search input only */
  section[data-testid="stSidebar"] [data-testid="stTextInput"] input {
      background: #ffffff !important;
      border-color: rgba(255,255,255,0.3) !important;
      color: #1a2744 !important;
      caret-color: #1a2744 !important;
      -webkit-text-fill-color: #1a2744 !important;
  }
  section[data-testid="stSidebar"] [data-testid="stTextInput"] input::placeholder { color: #8896a5 !important; }
  /* Multiselect / select boxes — white background, dark text */
  section[data-testid="stSidebar"] [data-baseweb="select"] > div,
  section[data-testid="stSidebar"] [data-baseweb="select"] > div > div {
      background: #ffffff !important;
      border-color: rgba(255,255,255,0.3) !important;
      color: #1a2744 !important;
  }
  section[data-testid="stSidebar"] [data-baseweb="select"] input {
      background: transparent !important;
      color: #1a2744 !important;
      -webkit-text-fill-color: #1a2744 !important;
  }
  section[data-testid="stSidebar"] [data-baseweb="select"] span { color: #1a2744 !important; }
  section[data-testid="stSidebar"] [data-baseweb="select"] [aria-selected] { color: #1a2744 !important; }
  section[data-testid="stSidebar"] [data-baseweb="select"] svg { fill: #8896a5 !important; }
  section[data-testid="stSidebar"] [data-baseweb="tag"] {
      background: rgba(26,39,68,0.12) !important;
      border-color: rgba(26,39,68,0.25) !important;
  }
  section[data-testid="stSidebar"] [data-baseweb="tag"] span { color: #1a2744 !important; }

  /* Radio nav */
  section[data-testid="stSidebar"] [data-testid="stRadio"] label { color: rgba(255,255,255,0.82) !important; }
  section[data-testid="stSidebar"] [data-testid="stRadio"] [data-baseweb="radio"] div { border-color: rgba(255,255,255,0.4) !important; }

  /* File uploader — compact single button */
  section[data-testid="stSidebar"] [data-testid="stFileUploader"] {
      background: transparent !important;
      border: none !important;
      padding: 0 !important;
  }
  section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
      border: 1px solid rgba(201,160,64,0.45) !important;
      background: rgba(201,160,64,0.1) !important;
      border-radius: 6px !important;
      padding: 2px 6px !important;
      min-height: unset !important;
  }
  section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] { display: none !important; }
  section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] > div { gap: 0 !important; padding: 2px 0 !important; }
  section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {
      width: 100% !important; font-size: 0.8rem !important;
      background: transparent !important; border: none !important;
      color: rgba(255,255,255,0.85) !important; padding: 4px 8px !important;
  }
  section[data-testid="stSidebar"] [data-testid="stFileUploader"] * { color: rgba(255,255,255,0.85) !important; }

  /* ── Main layout ── */
  .block-container { padding-top: 1.6rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; min-width: 700px !important; }

  /* ── Submission list buttons ── */
  .stButton > button {
      text-align: left !important;
      border-radius: 6px !important;
      border: 1px solid #e2e6ea !important;
      border-left: 4px solid transparent !important;
      padding: 10px 14px !important;
      background: #ffffff !important;
      color: #1a2744 !important;
      font-size: 0.88rem !important;
      white-space: nowrap !important;
      overflow: hidden !important;
      text-overflow: ellipsis !important;
  }
  .stButton > button:hover { border-color: #c9a040 !important; background: #fdf8ee !important; }
  .stButton > button[kind="primaryFormSubmit"],
  .stButton > button[kind="primary"],
  [data-testid="baseButton-primary"] {
      background: #1a2744 !important; color: #ffffff !important;
      border-color: #1a2744 !important; border-left: 4px solid #c9a040 !important;
  }
  .stButton > button[kind="primary"]:hover,
  [data-testid="baseButton-primary"]:hover { background: #243860 !important; border-color: #243860 !important; border-left-color: #c9a040 !important; }

  /* ── Metrics ── */
  [data-testid="metric-container"] {
      background: #f7f9fc; border: 1px solid #e2e6ea; border-radius: 8px; padding: 10px 14px;
  }
  div[data-testid="stMetricValue"] { font-size: 1rem; font-weight: 600; color: #1a2744; }
  div[data-testid="stMetricLabel"] { font-size: 0.75rem; color: #6c757d; }

  /* ── Misc ── */
  .stAlert { border-radius: 8px !important; }
  [data-testid="stExpander"] { border: 1px solid #e2e6ea !important; border-radius: 8px !important; }
  [data-testid="stDownloadButton"] button {
      background: #f7f9fc !important; color: #1a2744 !important;
      border: 1px solid #c9a040 !important; border-radius: 6px !important; font-weight: 500 !important;
  }
  hr { border-color: #e2e6ea !important; margin: 1.2rem 0 !important; }
  .stCaption { color: #8896a5 !important; font-size: 0.82rem !important; }
</style>
""", unsafe_allow_html=True)


# ── Google Sheets ─────────────────────────────────────────────────────────────

@st.cache_resource
def get_spreadsheet():
    if not GSHEETS_AVAILABLE:
        return None
    try:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]),
            scopes=[
                "https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive",
            ],
        )
        client = gspread.authorize(creds)
        name = st.secrets.get("google_sheet", {}).get("name", "SIN Tracker Notes")
        try:
            return client.open(name)
        except gspread.SpreadsheetNotFound:
            return client.create(name)
    except Exception:
        return None


def get_worksheet(tab_name, headers):
    ss = get_spreadsheet()
    if ss is None:
        return None
    cache_key = f"_ws_{tab_name}"
    if cache_key in st.session_state:
        return st.session_state[cache_key]
    try:
        ws = ss.worksheet(tab_name)
        try:
            existing = ws.row_values(1)
            if len(existing) < len(headers):
                for i in range(len(existing), len(headers)):
                    ws.update_cell(1, i + 1, headers[i])
        except Exception:
            pass
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(tab_name, rows=2000, cols=len(headers))
        ws.append_row(headers)
    st.session_state[cache_key] = ws
    return ws


# ── Submissions ───────────────────────────────────────────────────────────────

def load_submissions_from_sheet():
    if "submissions_cache" in st.session_state:
        return st.session_state.submissions_cache
    ws = get_worksheet("Submissions", SUBMISSION_COLS)
    if ws is None:
        return None
    records = ws.get_all_records()
    if not records:
        return None
    df = pd.DataFrame(records)
    df["Submission Time"] = pd.to_datetime(df["Submission Time"], utc=True, errors="coerce")
    df["Submission Date"] = df["Submission Time"].dt.strftime("%Y-%m-%d")
    st.session_state.submissions_cache = df
    return df


def save_submissions_to_sheet(df):
    ws = get_worksheet("Submissions", SUBMISSION_COLS)
    if ws is None:
        return
    rows = [[safe(row.get(c, "")) for c in SUBMISSION_COLS] for _, row in df.iterrows()]
    ws.clear()
    ws.append_row(SUBMISSION_COLS)
    if rows:
        ws.append_rows(rows)
    st.session_state.submissions_cache = df


# ── Notes ─────────────────────────────────────────────────────────────────────

def load_notes():
    if "notes_cache" in st.session_state:
        return st.session_state.notes_cache
    ws = get_worksheet("Notes", NOTE_COLS)
    if ws is not None:
        records = ws.get_all_records()
        notes = {
            r["submission_id"]: {k: r.get(k, "") for k in NOTE_COLS[1:]}
            for r in records if r.get("submission_id")
        }
    elif os.path.exists(NOTES_FILE):
        with open(NOTES_FILE) as f:
            notes = json.load(f)
    else:
        notes = {}
    st.session_state.notes_cache = notes
    return notes


def save_note(notes_data, sub_id, note):
    notes_data[sub_id] = note
    st.session_state.notes_cache = notes_data
    ws = get_worksheet("Notes", NOTE_COLS)
    if ws is not None:
        row_values = [sub_id] + [note.get(k, "") for k in NOTE_COLS[1:]]
        records = ws.get_all_records()
        existing = next(
            (i + 2 for i, r in enumerate(records) if r.get("submission_id") == sub_id), None
        )
        col_letter = chr(ord("A") + len(NOTE_COLS) - 1)
        if existing:
            ws.update(f"A{existing}:{col_letter}{existing}", [row_values])
        else:
            ws.append_row(row_values)
    else:
        with open(NOTES_FILE, "w") as f:
            json.dump(notes_data, f, indent=2)


def delete_note(notes_data, sub_id):
    notes_data.pop(sub_id, None)
    st.session_state.notes_cache = notes_data
    ws = get_worksheet("Notes", NOTE_COLS)
    if ws is not None:
        records = ws.get_all_records()
        row_idx = next((i + 2 for i, r in enumerate(records) if r.get("submission_id") == sub_id), None)
        if row_idx:
            ws.delete_rows(row_idx)
    else:
        with open(NOTES_FILE, "w") as f:
            json.dump(notes_data, f, indent=2)


# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_name(raw):
    s = str(raw)
    return s.split(",", 1)[1].strip() if "," in s else s.strip()


def safe(val):
    v = str(val).strip()
    return "" if v in ("nan", "None", "") else v


def rl_escape(text):
    return html.escape(str(text))


def parse_upload(uploaded_file):
    name = uploaded_file.name.lower()
    df = pd.read_excel(uploaded_file) if name.endswith((".xlsx", ".xls")) else pd.read_csv(uploaded_file)
    rename = {col: col.split(":", 1)[1].strip() for col in df.columns if ":" in col}
    df = df.rename(columns=rename)
    if "Name" in df.columns:
        df["Name"] = df["Name"].apply(parse_name)
    if "Email" not in df.columns and "Audience Email" in df.columns:
        df["Email"] = df["Audience Email"]
    if "Submission Time" in df.columns:
        df["Submission Time"] = pd.to_datetime(df["Submission Time"], utc=True)
        df["Submission Date"] = df["Submission Time"].dt.strftime("%Y-%m-%d")
    return df


def make_id(row):
    ts = row.get("Submission Time", "")
    ts_str = ts.isoformat() if hasattr(ts, "isoformat") else str(ts)
    return f"{row.get('Email', '')}|{ts_str}"


def is_test_entry(row):
    ticker  = safe(row.get("Primary Company (Ticker)", "")).lower()
    summary = safe(row.get("Investment Summary", "")).lower().strip()
    name    = safe(row.get("Name", "")).lower()
    if ticker == "test":
        return True
    if summary in {"test", "testing", "we are testing the website to see if my submission will work"}:
        return True
    if name in {"gel", "shahan", "shahan test", "rikka", "rikka tindle", "boston bonah"}:
        return True
    return False


def has_notes(note):
    if any(note.get(k) for k in ["meeting_notes", "followup_notes", "meeting_date", "followup_date"]):
        return True
    if note.get("status", "New") != "New":
        return True
    try:
        return bool(json.loads(note.get("notes_log") or "[]"))
    except Exception:
        return False


# ── Notes modal ───────────────────────────────────────────────────────────────

@st.dialog("Internal Notes")
def notes_dialog(sub_id, row, notes_data):
    note  = notes_data.get(sub_id, {})
    today = date.today()

    name   = safe(row.get("Name", ""))
    ticker = safe(row.get("Primary Company (Ticker)", ""))
    st.markdown(f"**{name}** — {ticker}")
    st.divider()

    nc1, nc2 = st.columns(2)
    with nc1:
        cur_status   = note.get("status", "New")
        status_val   = st.selectbox("Status", STATUS_OPTIONS,
            index=STATUS_OPTIONS.index(cur_status) if cur_status in STATUS_OPTIONS else 0)
        raw_mdate    = note.get("meeting_date", "")
        mdate_def    = datetime.strptime(raw_mdate, "%Y-%m-%d").date() if raw_mdate else None
        meeting_date = st.date_input("Meeting Date", value=mdate_def)

    with nc2:
        raw_fdate     = note.get("followup_date", "")
        fdate_def     = datetime.strptime(raw_fdate, "%Y-%m-%d").date() if raw_fdate else None
        followup_date = st.date_input("Follow-up Date", value=fdate_def)
        quick = st.selectbox("Quick follow-up in...",
            ["—", "3 days", "1 week", "2 weeks", "1 month", "3 months"])
        if quick != "—":
            days_map = {"3 days": 3, "1 week": 7, "2 weeks": 14, "1 month": 30, "3 months": 90}
            followup_date = today + timedelta(days=days_map[quick])

    meeting_notes  = st.text_area("Meeting Notes", value=note.get("meeting_notes", ""),
        placeholder="How did it go? Key takeaways, impressions...", height=100)
    followup_notes = st.text_area("Follow-up / Action Items", value=note.get("followup_notes", ""),
        placeholder="Next steps, requests, action items...", height=80)

    # ── Notes log ──
    st.divider()
    st.markdown("**📋 Notes Log**")
    try:
        notes_log = json.loads(note.get("notes_log") or "[]")
    except Exception:
        notes_log = []

    if notes_log:
        for i, entry in enumerate(reversed(notes_log)):
            real_i = len(notes_log) - 1 - i
            c_text, c_del = st.columns([7, 1])
            with c_text:
                ts_str = entry.get("ts", "")[:16].replace("T", " ")
                st.markdown(
                    f"<div style='background:#f7f9fc;border-radius:6px;padding:6px 10px;margin-bottom:4px;'>"
                    f"<small style='color:#8896a5;'>{ts_str}</small><br>"
                    f"<span style='font-size:0.9rem;'>{html.escape(entry.get('text',''))}</span></div>",
                    unsafe_allow_html=True,
                )
            with c_del:
                if st.button("🗑️", key=f"dellog_{sub_id}_{real_i}", help="Delete this entry"):
                    notes_log.pop(real_i)
                    updated = {**note, "notes_log": json.dumps(notes_log), "last_updated": datetime.now().isoformat()}
                    save_note(notes_data, sub_id, updated)
                    st.rerun()
    else:
        st.caption("No log entries yet.")

    add_col, btn_col = st.columns([5, 1])
    with add_col:
        new_entry = st.text_input("New log entry", placeholder="Add a quick note...", label_visibility="collapsed")
    with btn_col:
        if st.button("➕", help="Add entry"):
            if new_entry.strip():
                notes_log.append({"ts": datetime.now().isoformat(), "text": new_entry.strip()})
                updated = {**note, "notes_log": json.dumps(notes_log), "last_updated": datetime.now().isoformat()}
                save_note(notes_data, sub_id, updated)
                st.rerun()

    st.divider()
    save_col, del_col = st.columns([3, 1])
    with save_col:
        if st.button("💾 Save Notes", type="primary", use_container_width=True):
            save_note(notes_data, sub_id, {
                "status":         status_val,
                "meeting_date":   meeting_date.strftime("%Y-%m-%d") if meeting_date else "",
                "followup_date":  followup_date.strftime("%Y-%m-%d") if followup_date else "",
                "meeting_notes":  meeting_notes,
                "followup_notes": followup_notes,
                "notes_log":      json.dumps(notes_log),
                "starred":        note.get("starred", False),
                "last_updated":   datetime.now().isoformat(),
            })
            st.success("Saved!")
            st.rerun()
    with del_col:
        if st.button("🗑️ Delete All", use_container_width=True, help="Remove all notes for this submission"):
            delete_note(notes_data, sub_id)
            st.rerun()


# ── PDF ───────────────────────────────────────────────────────────────────────

def generate_pdf(row, notes_data):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        rightMargin=0.75*inch, leftMargin=0.75*inch,
        topMargin=0.75*inch, bottomMargin=0.75*inch,
    )
    base = getSampleStyleSheet()
    NAVY  = colors.HexColor("#1a2744")
    SLATE = colors.HexColor("#2c3e50")
    MUTED = colors.HexColor("#6c757d")
    LIGHT = colors.HexColor("#f0f4f8")
    GOLD  = colors.HexColor("#fff3cd")

    H1   = ParagraphStyle("H1",   parent=base["Heading1"], fontSize=22, textColor=NAVY,  spaceAfter=2,   leading=26)
    H2   = ParagraphStyle("H2",   parent=base["Heading2"], fontSize=11, textColor=SLATE, spaceBefore=14, spaceAfter=4)
    SUB  = ParagraphStyle("SUB",  parent=base["Normal"],   fontSize=10, textColor=MUTED, spaceAfter=10)
    BODY = ParagraphStyle("BODY", parent=base["Normal"],   fontSize=10, leading=15,      spaceAfter=6)
    FOOT = ParagraphStyle("FOOT", parent=base["Normal"],   fontSize=8,  textColor=MUTED, alignment=TA_CENTER, spaceBefore=6)

    story = []
    ticker   = safe(row.get("Primary Company (Ticker)", ""))
    name     = safe(row.get("Name", ""))
    email    = safe(row.get("Email", ""))
    phone    = safe(row.get("Phone", ""))
    sector   = safe(row.get("Sector / Industry", ""))
    mktcap   = safe(row.get("Market Capitalization", ""))
    target   = safe(row.get("12-Month Target Price", ""))
    sub_date = safe(row.get("Submission Date", ""))

    story.append(Paragraph(f"{rl_escape(ticker)} — {rl_escape(name)}", H1))
    story.append(Paragraph(
        f"{rl_escape(email)}&nbsp;&nbsp;|&nbsp;&nbsp;{rl_escape(phone)}", SUB
    ))
    story.append(HRFlowable(width="100%", thickness=2, color=NAVY, spaceAfter=10))

    tbl_data = [
        ["Sector / Industry", sector or "—"],
        ["Company (Ticker)",  ticker or "—"],
        ["Market Cap",        mktcap or "—"],
        ["12-Month Target",   f"${target}" if target and not target.startswith("$") else target or "—"],
        ["Submitted",         sub_date or "—"],
    ]
    t = Table(tbl_data, colWidths=[1.9*inch, 5.0*inch])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), LIGHT),
        ("FONTNAME",   (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE",   (0, 0), (-1, -1), 9),
        ("PADDING",    (0, 0), (-1, -1), 6),
        ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("ROWBACKGROUNDS", (1, 0), (1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
    ]))
    story.append(t)

    for field, label in [
        ("Professional Bio",   "Professional Background"),
        ("Your Edge",          "Investment Edge"),
        ("Investment Summary", "Investment Summary"),
    ]:
        text = safe(row.get(field, ""))
        if text:
            story.append(Paragraph(label, H2))
            for line in text.split("\n"):
                if line.strip():
                    story.append(Paragraph(rl_escape(line.strip()), BODY))

    doc_link = safe(row.get("Model or Supporting Materials Work", ""))
    if doc_link:
        story.append(Paragraph("Supporting Materials", H2))
        story.append(Paragraph("Document attached — accessible via original submission link.", BODY))

    sub_id = row.get("submission_id", make_id(row))
    note   = notes_data.get(sub_id, {})
    if any(note.get(k) for k in ["status", "meeting_date", "meeting_notes", "followup_date", "followup_notes"]):
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#dee2e6"), spaceBefore=14, spaceAfter=10))
        story.append(Paragraph("Internal Notes", H2))
        meta = [[k.replace("_", " ").title(), note[k]]
                for k in ["status", "meeting_date", "followup_date"] if note.get(k)]
        if meta:
            nt = Table(meta, colWidths=[1.9*inch, 5.0*inch])
            nt.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (0, -1), GOLD),
                ("FONTNAME",   (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTSIZE",   (0, 0), (-1, -1), 9),
                ("PADDING",    (0, 0), (-1, -1), 6),
                ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
            ]))
            story.append(nt)
            story.append(Spacer(1, 8))
        for key, label in [("meeting_notes", "Meeting Notes"), ("followup_notes", "Follow-up Notes")]:
            if note.get(key):
                story.append(Paragraph(f"<b>{label}:</b>", BODY))
                for line in note[key].split("\n"):
                    if line.strip():
                        story.append(Paragraph(rl_escape(line), BODY))

    story.append(Spacer(1, 24))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#dee2e6")))
    story.append(Paragraph(
        f"Generated {date.today().strftime('%B %d, %Y')} — First Wave Capital | Specialist Insights Network",
        FOOT,
    ))
    doc.build(story)
    buffer.seek(0)
    return buffer


# ── Submission detail panel ───────────────────────────────────────────────────

def render_detail(row, notes_data):
    sub_id = row.get("submission_id", make_id(row))
    note   = notes_data.get(sub_id, {})

    name     = safe(row.get("Name", ""))
    ticker   = safe(row.get("Primary Company (Ticker)", ""))
    email    = safe(row.get("Email", ""))
    phone    = safe(row.get("Phone", ""))
    sector   = safe(row.get("Sector / Industry", ""))
    mktcap   = safe(row.get("Market Capitalization", ""))
    target   = safe(row.get("12-Month Target Price", ""))
    sub_date = safe(row.get("Submission Date", ""))
    is_starred = note.get("starred", False)

    # ── Header ──
    h_left, h_right = st.columns([3, 1])
    with h_left:
        star_icon = "⭐" if is_starred else "☆"
        st.markdown(
            f"<h2 style='color:#1a2744;font-family:Georgia,serif;margin-bottom:4px;'>"
            f"{html.escape(ticker)}  —  {html.escape(name)}</h2>",
            unsafe_allow_html=True,
        )
        # Star button inline below title
        star_col, info_col = st.columns([1, 6])
        with star_col:
            if st.button(star_icon, key=f"star_{sub_id}", help="Star this submission"):
                updated = {**note, "starred": not is_starred, "last_updated": datetime.now().isoformat()}
                save_note(notes_data, sub_id, updated)
                st.rerun()
        with info_col:
            st.markdown(
                f"<p style='color:#8896a5;font-size:0.83rem;margin:0;padding-top:6px;'>"
                f"{html.escape(email)}  ·  {html.escape(phone)}  ·  Submitted {sub_date}</p>",
                unsafe_allow_html=True,
            )

    with h_right:
        # Notes button
        if has_notes(note):
            status = note.get("status", "New")
            st.markdown(
                f"<p style='font-size:0.82rem;font-weight:600;color:#1a2744;"
                f"margin-bottom:4px;'>{STATUS_ICONS.get(status,'')} {status}</p>",
                unsafe_allow_html=True,
            )
            if st.button("✏️ Edit Notes", key=f"edit_{sub_id}", use_container_width=True):
                notes_dialog(sub_id, row, notes_data)
        else:
            if st.button("📝 Add Notes", key=f"add_{sub_id}", use_container_width=True, type="primary"):
                notes_dialog(sub_id, row, notes_data)
        # PDF button — top right
        pdf_buf = generate_pdf(row, notes_data)
        st.download_button(
            "📄 Export PDF", data=pdf_buf,
            file_name=f"{ticker}_{name.replace(' ', '_')}.pdf",
            mime="application/pdf", key=f"pdf_{sub_id}",
            use_container_width=True,
        )

    c1, c2, c3 = st.columns(3)
    c1.metric("Market Cap", mktcap or "—")
    c2.metric("12-Mo Target", f"${target}" if target and not target.startswith("$") else target or "—")
    c3.metric("Sector", sector or "—")

    st.divider()

    bio = safe(row.get("Professional Bio", ""))
    if bio:
        with st.expander("Professional Bio"):
            st.write(bio)

    edge = safe(row.get("Your Edge", ""))
    if edge:
        st.markdown("**Investment Edge**")
        st.info(edge)

    summary = safe(row.get("Investment Summary", ""))
    if summary:
        st.markdown("**Investment Summary**")
        st.markdown(summary)

    doc_link = safe(row.get("Model or Supporting Materials Work", ""))
    if doc_link:
        st.markdown(
            f"<p>📎 <a href='{doc_link}' target='_blank' "
            f"style='color:#1a2744;font-weight:500;text-decoration:underline;pointer-events:auto;'>"
            f"Download Supporting Materials</a></p>",
            unsafe_allow_html=True,
        )


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    notes_data = load_notes()

    with st.sidebar:
        st.markdown(
            f"<div style='margin:-4rem -1.2rem 0 -1.2rem;height:160px;overflow:hidden;"
            f"display:flex;align-items:center;pointer-events:none;'>"
            f"<img src='{LOGO_URL}' style='width:100%;display:block;mix-blend-mode:screen;"
            f"transform:scale(1.6);transform-origin:center center;'></div>",
            unsafe_allow_html=True,
        )
        search = st.text_input("Search", placeholder="🔍 Search name, ticker, email", label_visibility="collapsed")
        st.divider()
        view = st.radio("View", ["📋  All Submissions", "📝  Notes", "⭐  Favorites"], label_visibility="collapsed")
        st.divider()
        uploaded = st.file_uploader("📤 Upload new export", type=["csv", "xlsx", "xls"], label_visibility="collapsed")
        st.divider()
        st.markdown("<p style='font-size:0.78rem;opacity:0.6;margin-bottom:4px;'>FILTERS</p>", unsafe_allow_html=True)
        hide_test     = st.checkbox("Hide test entries", value=True)
        sector_filter = []
        status_filter = []

    df = None
    if uploaded is not None:
        df = parse_upload(uploaded)
        df["submission_id"] = df.apply(make_id, axis=1)
        save_submissions_to_sheet(df)
        st.sidebar.success("Submissions saved.")
    else:
        df = load_submissions_from_sheet()

    if df is None or df.empty:
        st.markdown(
            "<h1 style='color:#1a2744;font-family:Georgia,serif;font-weight:700;'>"
            "Specialist Insights Network</h1>"
            "<p style='color:#8896a5;font-size:1rem;margin-top:-10px;'>First Wave Capital</p>",
            unsafe_allow_html=True,
        )
        st.info(
            "No submissions loaded yet. Upload your first Strikingly export from the sidebar.\n\n"
            "**Tip:** Always export *All Time* so previous weeks stay visible."
        )
        return

    if "submission_id" not in df.columns:
        df["submission_id"] = df.apply(make_id, axis=1)

    with st.sidebar:
        if "Sector / Industry" in df.columns:
            sectors = sorted(df["Sector / Industry"].dropna().unique().tolist())
            sector_filter = st.multiselect("Sector", options=sectors)
        status_filter = st.multiselect("Status", options=STATUS_OPTIONS)

    # Apply filters
    filtered = df.copy()
    if hide_test:
        filtered = filtered[~filtered.apply(is_test_entry, axis=1)]
    if search:
        q = search.lower()
        mask = (
            filtered.get("Name", pd.Series(dtype=str)).astype(str).str.lower().str.contains(q, na=False)
            | filtered.get("Primary Company (Ticker)", pd.Series(dtype=str)).astype(str).str.lower().str.contains(q, na=False)
            | filtered.get("Email", pd.Series(dtype=str)).astype(str).str.lower().str.contains(q, na=False)
        )
        filtered = filtered[mask]
    if sector_filter:
        filtered = filtered[filtered["Sector / Industry"].isin(sector_filter)]
    if status_filter:
        filtered = filtered[
            filtered["submission_id"].apply(
                lambda sid: notes_data.get(sid, {}).get("status", "New")
            ).isin(status_filter)
        ]

    # Follow-up alerts
    today = date.today()
    alerts = []
    for _, row in filtered.iterrows():
        fu = notes_data.get(row["submission_id"], {}).get("followup_date", "")
        if fu:
            try:
                delta = (datetime.strptime(fu, "%Y-%m-%d").date() - today).days
                if delta <= 3:
                    label = "today" if delta == 0 else f"in {delta}d" if delta > 0 else f"{abs(delta)}d overdue"
                    alerts.append((safe(row.get("Name", "")), safe(row.get("Primary Company (Ticker)", "")), label, delta))
            except ValueError:
                pass
    if alerts:
        with st.expander(f"⚠️ {len(alerts)} follow-up(s) due soon", expanded=True):
            for person, ticker, label, delta in alerts:
                icon = "🔴" if delta < 0 else "🟠" if delta == 0 else "🟡"
                st.markdown(f"{icon} **{person}** ({ticker}) — {label}")

    if "selected_id" not in st.session_state:
        st.session_state.selected_id = None

    # ── Header ──
    tracked_count = sum(1 for _, r in filtered.iterrows() if has_notes(notes_data.get(r["submission_id"], {})))
    st.markdown(
        "<h2 style='color:#1a2744;font-family:Georgia,serif;margin-top:0;margin-bottom:0.3rem;'>"
        "Specialist Insights Network</h2>",
        unsafe_allow_html=True,
    )
    st.caption(f"{len(filtered)} submission{'s' if len(filtered) != 1 else ''}  ·  {tracked_count} tracked")
    st.divider()

    col_list, col_detail = st.columns([1, 3], gap="large")

    # ── Submission list (left column) ──
    with col_list:
        if "📋" in view:
            rows_to_show = list(filtered.iterrows())
            selected_key = "selected_id"
            if not rows_to_show:
                st.warning("No submissions match the current filters.")
            for _, row in rows_to_show:
                sub_id = row["submission_id"]
                note   = notes_data.get(sub_id, {})
                status = note.get("status", "New")
                icon   = STATUS_ICONS.get(status, "⚪")
                name   = safe(row.get("Name", "Unknown"))
                ticker = safe(row.get("Primary Company (Ticker)", "—"))
                sector = safe(row.get("Sector / Industry", ""))
                sdate  = safe(row.get("Submission Date", ""))
                badge  = "  📝" if has_notes(note) else ""
                label  = f"{icon}**{ticker}**  ·  {name}{badge}"
                if st.button(label, key=f"btn_{sub_id}", use_container_width=True,
                             type="primary" if st.session_state.get(selected_key) == sub_id else "secondary"):
                    st.session_state[selected_key] = sub_id
                    st.rerun()
        elif "📝" in view:
            selected_key = "tracked_selected"
            tracked_rows = [row for _, row in filtered.iterrows()
                            if has_notes(notes_data.get(row["submission_id"], {}))]
            if not tracked_rows:
                st.info("No submissions with notes yet.")
            for row in tracked_rows:
                sub_id = row["submission_id"]
                note   = notes_data.get(sub_id, {})
                status = note.get("status", "New")
                icon   = STATUS_ICONS.get(status, "⚪")
                name   = safe(row.get("Name", "Unknown"))
                ticker = safe(row.get("Primary Company (Ticker)", "—"))
                fu     = note.get("followup_date", "")
                label  = f"{icon}**{ticker}**  ·  {name}"
                if st.button(label, key=f"tracked_{sub_id}", use_container_width=True,
                             type="primary" if st.session_state.get(selected_key) == sub_id else "secondary"):
                    st.session_state[selected_key] = sub_id
                    st.rerun()
        else:
            selected_key = "fav_selected"
            fav_rows = [row for _, row in filtered.iterrows()
                        if notes_data.get(row["submission_id"], {}).get("starred")]
            if not fav_rows:
                st.info("No starred submissions yet. Click ☆ on any submission to add it here.")
            for row in fav_rows:
                sub_id = row["submission_id"]
                note   = notes_data.get(sub_id, {})
                status = note.get("status", "New")
                icon   = STATUS_ICONS.get(status, "⚪")
                name   = safe(row.get("Name", "Unknown"))
                ticker = safe(row.get("Primary Company (Ticker)", "—"))
                sector = safe(row.get("Sector / Industry", ""))
                label  = f"⭐ **{ticker}**  ·  {name}"
                if st.button(label, key=f"fav_{sub_id}", use_container_width=True,
                             type="primary" if st.session_state.get(selected_key) == sub_id else "secondary"):
                    st.session_state[selected_key] = sub_id
                    st.rerun()

    # ── Detail panel (right column) ──
    with col_detail:
        sel_id = st.session_state.get(selected_key)
        if sel_id is None:
            st.markdown(
                "<div style='margin-top:3rem;text-align:center;color:#8896a5;'>"
                "<p style='font-size:2rem;'>←</p>"
                "<p>Select a submission to view details</p></div>",
                unsafe_allow_html=True,
            )
        else:
            if "📋" in view:
                match = filtered[filtered["submission_id"] == sel_id]
                row = match.iloc[0] if not match.empty else None
            elif "📝" in view:
                match = [r for r in tracked_rows if r["submission_id"] == sel_id]
                row = match[0] if match else None
            else:
                match = [r for r in fav_rows if r["submission_id"] == sel_id]
                row = match[0] if match else None
            if row is not None:
                render_detail(row, notes_data)
            else:
                st.info("Select a submission to view details.")


if __name__ == "__main__":
    main()
