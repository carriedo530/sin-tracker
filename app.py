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
    "meeting_notes", "followup_notes", "last_updated",
]

STATUS_ICONS = {
    "New": "🔵", "Meeting Scheduled": "🟡", "Met": "🟢",
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
  /* ── Sidebar ── */
  [data-testid="stSidebar"] {
      background: linear-gradient(175deg, #0b1220 0%, #1a2744 100%) !important;
  }
  [data-testid="stSidebar"] .stMarkdown p,
  [data-testid="stSidebar"] .stMarkdown,
  [data-testid="stSidebar"] label,
  [data-testid="stSidebar"] span,
  [data-testid="stSidebar"] .stCheckbox span {
      color: rgba(255,255,255,0.85) !important;
  }
  [data-testid="stSidebar"] .stTextInput input,
  [data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] {
      background: rgba(255,255,255,0.08) !important;
      border-color: rgba(255,255,255,0.15) !important;
      color: white !important;
  }
  [data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.12) !important; }
  [data-testid="stSidebar"] .stFileUploader {
      background: rgba(255,255,255,0.06) !important;
      border: 1px dashed rgba(255,255,255,0.25) !important;
      border-radius: 8px !important;
  }
  [data-testid="stSidebar"] .stSuccess {
      background: rgba(40,167,69,0.2) !important;
      color: #90ee90 !important;
  }

  /* ── Main layout ── */
  .block-container { padding-top: 1.8rem; max-width: 1400px; }

  /* ── Submission list buttons ── */
  [data-testid="stSidebar"] .stButton > button,
  .stButton > button {
      text-align: left !important;
      border-radius: 6px !important;
      border: 1px solid #e2e6ea !important;
      padding: 10px 14px !important;
      background: #ffffff !important;
      color: #1a2744 !important;
      font-size: 0.88rem !important;
      transition: border-color 0.15s, background 0.15s !important;
  }
  .stButton > button:hover {
      border-color: #c9a040 !important;
      background: #fdf8ee !important;
  }
  /* Active / selected */
  [data-testid="baseButton-primary"] {
      background: #1a2744 !important;
      color: #ffffff !important;
      border-color: #1a2744 !important;
  }
  [data-testid="baseButton-primary"]:hover {
      background: #243860 !important;
      border-color: #243860 !important;
  }

  /* ── Tabs ── */
  .stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 2px solid #e2e6ea; }
  .stTabs [data-baseweb="tab"] {
      background: transparent !important;
      color: #6c757d !important;
      font-weight: 500;
      padding: 8px 18px;
      border-radius: 6px 6px 0 0;
  }
  .stTabs [aria-selected="true"] {
      color: #1a2744 !important;
      border-bottom: 3px solid #c9a040 !important;
      font-weight: 600 !important;
  }

  /* ── Metrics ── */
  [data-testid="metric-container"] {
      background: #f7f9fc;
      border: 1px solid #e2e6ea;
      border-radius: 8px;
      padding: 10px 14px;
  }
  div[data-testid="stMetricValue"] { font-size: 1rem; font-weight: 600; color: #1a2744; }
  div[data-testid="stMetricLabel"] { font-size: 0.75rem; color: #6c757d; }

  /* ── Info / alert boxes ── */
  .stAlert { border-radius: 8px !important; }
  [data-testid="stExpander"] { border: 1px solid #e2e6ea !important; border-radius: 8px !important; }

  /* ── Download button ── */
  [data-testid="stDownloadButton"] button {
      background: #f7f9fc !important;
      color: #1a2744 !important;
      border: 1px solid #c9a040 !important;
      border-radius: 6px !important;
      font-weight: 500 !important;
  }
  [data-testid="stDownloadButton"] button:hover {
      background: #fdf8ee !important;
  }

  /* ── Dividers ── */
  hr { border-color: #e2e6ea !important; margin: 1.2rem 0 !important; }

  /* ── Caption / muted text ── */
  .stCaption, [data-testid="stCaptionContainer"] { color: #8896a5 !important; font-size: 0.82rem !important; }
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
    try:
        ws = ss.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(tab_name, rows=2000, cols=len(headers))
        ws.append_row(headers)
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
        if existing:
            ws.update(f"A{existing}:G{existing}", [row_values])
        else:
            ws.append_row(row_values)
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
    return any(note.get(k) for k in ["meeting_notes", "followup_notes", "meeting_date", "followup_date"]) \
           or note.get("status", "New") != "New"


# ── Notes modal ───────────────────────────────────────────────────────────────

@st.dialog("Internal Notes", width="large")
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
        placeholder="How did it go? Key takeaways, impressions...", height=120)
    followup_notes = st.text_area("Follow-up / Action Items", value=note.get("followup_notes", ""),
        placeholder="Next steps, requests, action items...", height=90)

    if st.button("💾 Save Notes", type="primary", use_container_width=True):
        save_note(notes_data, sub_id, {
            "status":         status_val,
            "meeting_date":   meeting_date.strftime("%Y-%m-%d") if meeting_date else "",
            "followup_date":  followup_date.strftime("%Y-%m-%d") if followup_date else "",
            "meeting_notes":  meeting_notes,
            "followup_notes": followup_notes,
            "last_updated":   datetime.now().isoformat(),
        })
        st.success("Saved!")
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

    story.append(Paragraph(f"{rl_escape(ticker)} — Analyst Submission", H1))
    story.append(Paragraph(
        f"{rl_escape(name)}&nbsp;&nbsp;|&nbsp;&nbsp;{rl_escape(email)}&nbsp;&nbsp;|&nbsp;&nbsp;{rl_escape(phone)}", SUB
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

    # Header row with notes button
    h1, h2 = st.columns([3, 1])
    with h1:
        st.markdown(f"## {ticker}  —  {name}")
        st.caption(f"{email}  ·  {phone}  ·  Submitted {sub_date}")
    with h2:
        if has_notes(note):
            status = note.get("status", "New")
            icon   = STATUS_ICONS.get(status, "⚪")
            st.markdown(f"**{icon} {status}**")
            if note.get("meeting_date"):
                st.caption(f"Met: {note['meeting_date']}")
            if note.get("followup_date"):
                st.caption(f"Follow-up: {note['followup_date']}")
            if st.button("✏️ Edit Notes", key=f"edit_{sub_id}", use_container_width=True):
                notes_dialog(sub_id, row, notes_data)
        else:
            if st.button("📝 Add Notes", key=f"add_{sub_id}", use_container_width=True, type="primary"):
                notes_dialog(sub_id, row, notes_data)

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
        st.markdown(f"📎 [Download Supporting Materials]({doc_link})")

    st.divider()
    pdf_buf = generate_pdf(row, notes_data)
    st.download_button(
        "📄 Export PDF", data=pdf_buf,
        file_name=f"{ticker}_{name.replace(' ', '_')}.pdf",
        mime="application/pdf", key=f"pdf_{sub_id}",
    )


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    notes_data = load_notes()

    with st.sidebar:
        st.image(LOGO_URL, use_container_width=True)
        st.markdown(
            "<p style='color:rgba(255,255,255,0.5);font-size:0.72rem;letter-spacing:0.12em;"
            "text-transform:uppercase;margin:-8px 0 4px 2px;'>Specialist Insights Network</p>",
            unsafe_allow_html=True,
        )
        st.divider()
        uploaded = st.file_uploader("Upload new export (.csv or .xlsx)", type=["csv", "xlsx", "xls"])
        st.divider()
        st.markdown("### Filters")
        search        = st.text_input("🔍 Search name, ticker, email")
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

    st.markdown(
        "<p style='color:#8896a5;font-size:0.8rem;letter-spacing:0.1em;text-transform:uppercase;"
        "margin-bottom:0.2rem;'>First Wave Capital</p>"
        "<h2 style='color:#1a2744;font-family:Georgia,serif;margin-top:0;'>Specialist Insights Network</h2>",
        unsafe_allow_html=True,
    )

    # ── Tabs ──
    tab_all, tab_tracked = st.tabs([
        f"📋 All Submissions ({len(filtered)})",
        f"📌 Tracked ({sum(1 for _, r in filtered.iterrows() if has_notes(notes_data.get(r['submission_id'], {})))})",
    ])

    # ── All Submissions tab ──
    with tab_all:
        if filtered.empty:
            st.warning("No submissions match the current filters.")
        else:
            col_list, col_detail = st.columns([1, 2], gap="large")
            with col_list:
                for _, row in filtered.iterrows():
                    sub_id = row["submission_id"]
                    note   = notes_data.get(sub_id, {})
                    status = note.get("status", "New")
                    icon   = STATUS_ICONS.get(status, "⚪")
                    name   = safe(row.get("Name", "Unknown"))
                    ticker = safe(row.get("Primary Company (Ticker)", "—"))
                    sector = safe(row.get("Sector / Industry", ""))
                    sdate  = safe(row.get("Submission Date", ""))
                    note_badge = "  ✅" if has_notes(note) else ""
                    label  = f"{icon} **{name}**  ·  {ticker}{note_badge}\n{sector}  ·  {sdate}"
                    if st.button(label, key=f"btn_{sub_id}", use_container_width=True,
                                 type="primary" if st.session_state.selected_id == sub_id else "secondary"):
                        st.session_state.selected_id = sub_id
                        st.rerun()

            with col_detail:
                if st.session_state.selected_id is None:
                    st.info("← Select a submission to view details.")
                else:
                    rows = filtered[filtered["submission_id"] == st.session_state.selected_id]
                    if rows.empty:
                        st.info("← Select a submission to view details.")
                    else:
                        render_detail(rows.iloc[0], notes_data)

    # ── Tracked tab ──
    with tab_tracked:
        tracked = [
            row for _, row in filtered.iterrows()
            if has_notes(notes_data.get(row["submission_id"], {}))
        ]
        if not tracked:
            st.info("No submissions with notes yet. Add notes to a submission and it will appear here.")
        else:
            if "tracked_selected" not in st.session_state:
                st.session_state.tracked_selected = None

            col_list, col_detail = st.columns([1, 2], gap="large")
            with col_list:
                for row in tracked:
                    sub_id = row["submission_id"]
                    note   = notes_data.get(sub_id, {})
                    status = note.get("status", "New")
                    icon   = STATUS_ICONS.get(status, "⚪")
                    name   = safe(row.get("Name", "Unknown"))
                    ticker = safe(row.get("Primary Company (Ticker)", "—"))
                    fu     = note.get("followup_date", "")
                    fu_str = f"  ·  Follow-up: {fu}" if fu else ""
                    label  = f"{icon} **{name}**  ·  {ticker}\n{status}{fu_str}"
                    if st.button(label, key=f"tracked_{sub_id}", use_container_width=True,
                                 type="primary" if st.session_state.tracked_selected == sub_id else "secondary"):
                        st.session_state.tracked_selected = sub_id
                        st.rerun()

            with col_detail:
                sel = st.session_state.tracked_selected
                if sel is None:
                    st.info("← Select a tracked submission to view details.")
                else:
                    match = [r for r in tracked if r["submission_id"] == sel]
                    if match:
                        render_detail(match[0], notes_data)
                    else:
                        st.info("← Select a tracked submission to view details.")


if __name__ == "__main__":
    main()
