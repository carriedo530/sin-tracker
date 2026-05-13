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

# Google Sheets (optional — falls back to local JSON if not configured)
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSHEETS_AVAILABLE = True
except ImportError:
    GSHEETS_AVAILABLE = False

NOTES_FILE = os.path.join(os.path.dirname(__file__), "submission_notes.json")
SHEET_HEADERS = ["submission_id", "status", "meeting_date", "followup_date",
                 "meeting_notes", "followup_notes", "last_updated"]

st.set_page_config(
    page_title="SIN Tracker",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    div[data-testid="stMetricValue"] { font-size: 1rem; }
    .stButton > button { text-align: left !important; }
</style>
""", unsafe_allow_html=True)

STATUS_ICONS = {
    "New": "🔵",
    "Meeting Scheduled": "🟡",
    "Met": "🟢",
    "Pass": "🔴",
    "Follow-up": "🟠",
}
STATUS_OPTIONS = list(STATUS_ICONS.keys())


# ── Google Sheets backend ─────────────────────────────────────────────────────

@st.cache_resource
def get_sheet():
    if not GSHEETS_AVAILABLE:
        return None
    try:
        creds_info = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(
            creds_info,
            scopes=[
                "https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive",
            ],
        )
        client = gspread.authorize(creds)
        sheet_name = st.secrets.get("google_sheet", {}).get("name", "SIN Tracker Notes")
        try:
            sh = client.open(sheet_name)
        except gspread.SpreadsheetNotFound:
            sh = client.create(sheet_name)
        ws = sh.sheet1
        if not ws.get_all_values():
            ws.append_row(SHEET_HEADERS)
        return ws
    except Exception:
        return None


def load_notes():
    if "notes_cache" in st.session_state:
        return st.session_state.notes_cache

    sheet = get_sheet()
    if sheet is not None:
        records = sheet.get_all_records()
        notes = {
            r["submission_id"]: {k: r.get(k, "") for k in SHEET_HEADERS[1:]}
            for r in records if r.get("submission_id")
        }
    else:
        if os.path.exists(NOTES_FILE):
            with open(NOTES_FILE, "r") as f:
                notes = json.load(f)
        else:
            notes = {}

    st.session_state.notes_cache = notes
    return notes


def save_note(notes_data, sub_id, note):
    notes_data[sub_id] = note
    st.session_state.notes_cache = notes_data

    sheet = get_sheet()
    if sheet is not None:
        row_values = [sub_id] + [note.get(k, "") for k in SHEET_HEADERS[1:]]
        records = sheet.get_all_records()
        existing = next(
            (i + 2 for i, r in enumerate(records) if r.get("submission_id") == sub_id),
            None,
        )
        if existing:
            sheet.update(f"A{existing}:G{existing}", [row_values])
        else:
            sheet.append_row(row_values)
    else:
        with open(NOTES_FILE, "w") as f:
            json.dump(notes_data, f, indent=2)


# ── CSV / Excel parsing ───────────────────────────────────────────────────────

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
    if name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(uploaded_file)
    else:
        df = pd.read_csv(uploaded_file)

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
    ts = row["Submission Time"]
    return f"{row.get('Email', '')}|{ts.isoformat() if hasattr(ts, 'isoformat') else ts}"


def is_test_entry(row):
    ticker = safe(row.get("Primary Company (Ticker)", "")).lower()
    summary = safe(row.get("Investment Summary", "")).lower().strip()
    name = safe(row.get("Name", "")).lower()
    if ticker == "test":
        return True
    if summary in {"test", "testing", "we are testing the website to see if my submission will work"}:
        return True
    if name in {"gel", "shahan", "shahan test", "rikka", "rikka tindle", "boston bonah"}:
        return True
    return False


# ── PDF generation ────────────────────────────────────────────────────────────

def generate_pdf(row, notes_data):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        rightMargin=0.75 * inch, leftMargin=0.75 * inch,
        topMargin=0.75 * inch, bottomMargin=0.75 * inch,
    )

    base = getSampleStyleSheet()
    NAVY  = colors.HexColor("#1a2744")
    SLATE = colors.HexColor("#2c3e50")
    MUTED = colors.HexColor("#6c757d")
    LIGHT = colors.HexColor("#f0f4f8")
    GOLD  = colors.HexColor("#fff3cd")

    H1   = ParagraphStyle("H1",   parent=base["Heading1"], fontSize=22, textColor=NAVY,  spaceAfter=2,  leading=26)
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
    t = Table(tbl_data, colWidths=[1.9 * inch, 5.0 * inch])
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
        ("Professional Bio",    "Professional Background"),
        ("Your Edge",           "Investment Edge"),
        ("Investment Summary",  "Investment Summary"),
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

    sub_id = make_id(row)
    note   = notes_data.get(sub_id, {})
    has_notes = any(note.get(k) for k in ["status", "meeting_date", "meeting_notes", "followup_date", "followup_notes"])

    if has_notes:
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#dee2e6"), spaceBefore=14, spaceAfter=10))
        story.append(Paragraph("Internal Notes", H2))

        meta = [[k.replace("_", " ").title(), note[k]]
                for k in ["status", "meeting_date", "followup_date"] if note.get(k)]
        if meta:
            nt = Table(meta, colWidths=[1.9 * inch, 5.0 * inch])
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


# ── main app ──────────────────────────────────────────────────────────────────

def main():
    notes_data = load_notes()

    with st.sidebar:
        st.markdown("## 📊 SIN Tracker")
        st.caption("First Wave Capital")
        st.divider()
        uploaded = st.file_uploader("Upload export (.csv or .xlsx)", type=["csv", "xlsx", "xls"])
        st.divider()
        st.markdown("### Filters")
        search      = st.text_input("🔍 Search name, ticker, email")
        hide_test   = st.checkbox("Hide test entries", value=True)
        sector_filter = []
        status_filter = []

    if uploaded is None:
        st.title("Specialist Insights Network")
        st.info(
            "Upload your Strikingly export from the sidebar to get started.\n\n"
            "**Tip:** Always export *All Time* so previous weeks stay visible — notes are preserved automatically."
        )
        return

    df = parse_upload(uploaded)

    with st.sidebar:
        if "Sector / Industry" in df.columns:
            sectors = sorted(df["Sector / Industry"].dropna().unique().tolist())
            sector_filter = st.multiselect("Sector", options=sectors)
        status_filter = st.multiselect("Status", options=STATUS_OPTIONS)

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
            filtered.apply(lambda r: notes_data.get(make_id(r), {}).get("status", "New"), axis=1).isin(status_filter)
        ]

    # follow-up alerts
    today = date.today()
    alerts = []
    for _, row in filtered.iterrows():
        fu = notes_data.get(make_id(row), {}).get("followup_date", "")
        if fu:
            try:
                fu_date = datetime.strptime(fu, "%Y-%m-%d").date()
                delta = (fu_date - today).days
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

    st.markdown(f"### {len(filtered)} Submission{'s' if len(filtered) != 1 else ''}")

    if filtered.empty:
        st.warning("No submissions match the current filters.")
        return

    if "selected_id" not in st.session_state:
        st.session_state.selected_id = None

    col_list, col_detail = st.columns([1, 2], gap="large")

    with col_list:
        for _, row in filtered.iterrows():
            sub_id = make_id(row)
            note   = notes_data.get(sub_id, {})
            status = note.get("status", "New")
            icon   = STATUS_ICONS.get(status, "⚪")
            name   = safe(row.get("Name", "Unknown"))
            ticker = safe(row.get("Primary Company (Ticker)", "—"))
            sector = safe(row.get("Sector / Industry", ""))
            sdate  = safe(row.get("Submission Date", ""))
            label  = f"{icon} **{name}**  ·  {ticker}\n{sector}  ·  {sdate}"
            if st.button(label, key=f"btn_{sub_id}", use_container_width=True,
                         type="primary" if st.session_state.selected_id == sub_id else "secondary"):
                st.session_state.selected_id = sub_id
                st.rerun()

    with col_detail:
        if st.session_state.selected_id is None:
            st.info("← Select a submission to view details and add notes.")
            return

        rows = filtered[filtered.apply(lambda r: make_id(r) == st.session_state.selected_id, axis=1)]
        if rows.empty:
            st.info("← Select a submission to view details.")
            return

        row    = rows.iloc[0]
        sub_id = make_id(row)
        note   = notes_data.get(sub_id, {})

        name     = safe(row.get("Name", ""))
        ticker   = safe(row.get("Primary Company (Ticker)", ""))
        email    = safe(row.get("Email", ""))
        phone    = safe(row.get("Phone", ""))
        sector   = safe(row.get("Sector / Industry", ""))
        mktcap   = safe(row.get("Market Capitalization", ""))
        target   = safe(row.get("12-Month Target Price", ""))
        sub_date = safe(row.get("Submission Date", ""))

        st.markdown(f"## {ticker}  —  {name}")
        st.caption(f"{email}  ·  {phone}  ·  Submitted {sub_date}")

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
        st.markdown("### 📝 Internal Notes")

        nc1, nc2 = st.columns(2)
        with nc1:
            cur_status  = note.get("status", "New")
            status_val  = st.selectbox("Status", STATUS_OPTIONS,
                index=STATUS_OPTIONS.index(cur_status) if cur_status in STATUS_OPTIONS else 0,
                key=f"status_{sub_id}")
            raw_mdate   = note.get("meeting_date", "")
            mdate_def   = datetime.strptime(raw_mdate, "%Y-%m-%d").date() if raw_mdate else None
            meeting_date = st.date_input("Meeting Date", value=mdate_def, key=f"mdate_{sub_id}")

        with nc2:
            raw_fdate   = note.get("followup_date", "")
            fdate_def   = datetime.strptime(raw_fdate, "%Y-%m-%d").date() if raw_fdate else None
            followup_date = st.date_input("Follow-up Date", value=fdate_def, key=f"fdate_{sub_id}")
            quick = st.selectbox("Quick follow-up in...",
                ["—", "3 days", "1 week", "2 weeks", "1 month", "3 months"],
                key=f"quick_{sub_id}")
            if quick != "—":
                days_map = {"3 days": 3, "1 week": 7, "2 weeks": 14, "1 month": 30, "3 months": 90}
                followup_date = today + timedelta(days=days_map[quick])

        meeting_notes = st.text_area("Meeting Notes", value=note.get("meeting_notes", ""),
            placeholder="How did it go? Key takeaways, impressions...",
            height=110, key=f"mnotes_{sub_id}")
        followup_notes = st.text_area("Follow-up / Action Items", value=note.get("followup_notes", ""),
            placeholder="Next steps, requests, action items...",
            height=80, key=f"fnotes_{sub_id}")

        btn1, btn2 = st.columns(2)
        with btn1:
            if st.button("💾 Save Notes", key=f"save_{sub_id}", use_container_width=True, type="primary"):
                save_note(notes_data, sub_id, {
                    "status":        status_val,
                    "meeting_date":  meeting_date.strftime("%Y-%m-%d") if meeting_date else "",
                    "followup_date": followup_date.strftime("%Y-%m-%d") if followup_date else "",
                    "meeting_notes": meeting_notes,
                    "followup_notes": followup_notes,
                    "last_updated":  datetime.now().isoformat(),
                })
                st.success("Notes saved!")
                st.rerun()

        with btn2:
            pdf_buf = generate_pdf(row, notes_data)
            st.download_button(
                "📄 Export PDF",
                data=pdf_buf,
                file_name=f"{ticker}_{name.replace(' ', '_')}.pdf",
                mime="application/pdf",
                key=f"pdf_{sub_id}",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()
