"""
PARTNER OPERATIONS PULSE (3-Page Dashboard)
--------------------------------------------------------
Reads TWO separate Excel files:
  1. Pending_Tickets.xlsx (Open / Pending backlog)
  2. Closed_Tickets.xlsx  (Completed / Closed tickets)

Generates index.html with:
  - Responsive screen layout matching device screen width
  - Strict fixed 3-page poster structure reserved for PDF/Print
  - Page 1: Daily Partner Operations Pulse (includes Closed & 4PM Cutoff KPIs)
  - Page 2: Partner Aging & Accountability Heatmap
  - Page 3: Partner Output Performance & Productivity
            - Substring filter strictly excludes "Adrian" and "Unassigned"
            - Sorted descending by Output % (Highest to Lowest)
            - Output % >= 70% highlighted green

Requires: pip install pandas openpyxl
"""

import html
import json
import math
import re
import sys
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pandas as pd

# ============================================================
# SETTINGS
# ============================================================

DATA_DIR = Path(__file__).resolve().parent

PENDING_FILE = "Pending_Tickets.xlsx"
CLOSED_FILE = "Closed_Tickets.xlsx"
OUTPUT_HTML = "index.html"
HISTORY_FILE = "pending_history.json"

REPORT_TIME = datetime.now()

# Age thresholds
FRESH_MAX = 2                     # 0-2 days: on track
WATCH_MAX = 6                     # 3-6 days: ageing; 7+ days: critical
ESCALATION_FROM = FRESH_MAX + 1

HEATMAP_MAX_AGE = 16              # Ages >= 16 share the last heatmap column
TOP_N_PARTNERS = 10               # Main partners listed individually on Page 2
MIN_TICKETS_FOR_PCT_RANKING = 10  # Minimum queue size for % rankings
FOCUS_CATEGORY = "Maintenance"

# Page 3 Cut-offs and benchmarks
CUTOFF_HOUR = 16                  # 4:00 PM cut-off
OUTPUT_GOOD = 70                  # Output % >= 70% -> Green
OUTPUT_POOR = 40                  # Output % <= 40% -> Red

# Explicit list of partner substrings to exclude from Page 3 Performance Table
EXCLUDE_PARTNERS = ["adrian", "unassigned"]

SERVE = True                      # Launch local http server
PORT = 8000
OPEN_BROWSER = True               # Open in default browser automatically

CATEGORIES = ["Maintenance", "Connections", "Relocation", "FTTB", "PTMP", "Other"]

CAT_COLOR = {
    "Maintenance": "#F28C1B",
    "Connections": "#1E88E5",
    "Relocation": "#2E9B4F",
    "FTTB": "#7B3FE4",
    "PTMP": "#78909C",
    "Other": "#9AA5B1",
}

CAT_ICON = {
    "Maintenance": "&#128295;",
    "Connections": "&#128268;",
    "Relocation": "&#128260;",
    "FTTB": "&#127968;",
    "PTMP": "&#128225;",
    "Other": "&#8943;",
}

CATEGORY_MAP = {
    "Connections": "Connections",
    "FTTB Connection": "FTTB",
    "FTTB Maintenance": "FTTB",
    "Maintenance": "Maintenance",
    "PTMP": "PTMP",
    "Relocation": "Relocation",
}

TEAM_COLUMNS = ["Teams", "Team", "Assigned to", "Assignee", "Resolved by", "Closed by", "Technician", "Engineer"]


# ============================================================
# HELPERS
# ============================================================

def esc(value):
    return html.escape(str(value))


def num(value):
    return f"{int(value):,}" if value is not None else "-"


def pct(part, whole):
    return (part / whole * 100) if whole else 0.0


def clean_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    for character in "_-/\\,.:;|":
        text = text.replace(character, " ")
    return " ".join(text.split())


def find_column(df, name):
    for column in df.columns:
        if str(column).strip().lower() == name.lower():
            return column
    return None


def clean_partner(value):
    if pd.isna(value):
        return "Unassigned"
    text = " ".join(str(value).split())
    text = re.sub(r"^agents?\s+", "", text, flags=re.IGNORECASE)
    return text or "Unassigned"


def raw_partner(value):
    if pd.isna(value):
        return "Unassigned"
    return " ".join(str(value).split()) or "Unassigned"


def hex_to_rgb(color):
    color = color.lstrip("#")
    return tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))


def mix(c1, c2, t):
    t = max(0.0, min(1.0, t))
    a, b = hex_to_rgb(c1), hex_to_rgb(c2)
    return "#%02x%02x%02x" % tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def ramp(value, stops):
    if value <= stops[0][0]:
        return stops[0][1]
    for (p1, c1), (p2, c2) in zip(stops, stops[1:]):
        if value <= p2:
            return mix(c1, c2, (value - p1) / (p2 - p1))
    return stops[-1][1]


def heat_cell(age, value):
    if value == 0:
        return "#F7FAF9", "#A3AEB8"
    if age <= FRESH_MAX:
        t = min(1.0, math.log10(value + 1) / 2.5)
        return mix("#DDF2D8", "#6CC36F", t), "#1D3B22"
    if age <= WATCH_MAX:
        if value <= 3:
            return "#FFF1A6", "#4A3B00"
        if value <= 10:
            return "#FFD966", "#4A3B00"
        if value <= 30:
            return "#FFA94D", "#3D1F00"
        return "#FF7043", "#FFFFFF"
    if value <= 2:
        return "#FFC27A", "#3D1F00"
    if value <= 10:
        return "#FF8A50", "#3D1F00"
    return "#EF4B45", "#FFFFFF"


def colour_fresh(share):
    return ramp(share, [(0.30, "#F6A15C"), (0.55, "#FFE27A"), (0.75, "#C5E58A"), (0.90, "#5DBB63")])


def colour_watch(share):
    return ramp(share, [(0.0, "#FFFBE0"), (0.15, "#FFF0A6"), (0.35, "#FFC861"), (0.55, "#FF9A4D")])


def colour_old(share):
    return ramp(share, [(0.0, "#FFF3E0"), (0.03, "#FFD29A"), (0.06, "#FF9A4D"), (0.10, "#EF5350")])


def colour_aged_count(count, biggest):
    return ramp(count / biggest if biggest else 0, [(0.0, "#FFF5F5"), (1.0, "#F4A3A3")])


def colour_pct_aged(value):
    return ramp(value / 100, [(0.05, "#FFFFFF"), (0.20, "#FFE0E0"), (0.30, "#FFB4B4"), (0.60, "#EF5350")])


# ============================================================
# CLASSIFICATION & DATA INGESTION
# ============================================================

def classify_issue_type(row):
    issue_type = clean_text(row.get("Issue type", ""))
    service = clean_text(row.get("Service", ""))
    description = clean_text(row.get("Short description", ""))
    business_type = clean_text(row.get("Business type", ""))
    location = clean_text(row.get("Location", ""))

    combined = " ".join([issue_type, service, description, business_type, location])

    if "fttb" in combined:
        if "maintenance" in combined or "maint" in combined or "mnt" in combined:
            return "FTTB Maintenance"
        if "connection" in combined or "connect" in combined:
            return "FTTB Connection"

    if "ptmp" in combined:
        return "PTMP"

    if "relocation" in combined or "relocate" in combined or "relocated" in combined:
        return "Relocation"

    if "maintenance" in issue_type or "maint" in issue_type:
        return "Maintenance"

    if "connection" in issue_type or "connection" in service or "connection" in description:
        return "Connections"

    return None


def dashboard_category(row):
    return CATEGORY_MAP.get(classify_issue_type(row), "Other")


def work_type_category(row):
    cat = dashboard_category(row)
    if cat in ("Connections", "FTTB"):
        return "Connection"
    if cat == "Maintenance":
        return "Maintenance"
    return None


def load_two_files():
    pending_path = DATA_DIR / PENDING_FILE
    closed_path = DATA_DIR / CLOSED_FILE

    if not pending_path.exists():
        print(f"ERROR: {PENDING_FILE} not found in {DATA_DIR}")
        sys.exit(1)

    print(f"Reading Pending data from: {pending_path.name}")
    p_excel = pd.ExcelFile(pending_path)
    p_frames = [p_excel.parse(s) for s in p_excel.sheet_names]
    df_pending = pd.concat(p_frames, ignore_index=True)
    df_pending.columns = [str(c).strip() for c in df_pending.columns]

    df_closed = None
    if closed_path.exists():
        print(f"Reading Closed data from: {closed_path.name}")
        c_excel = pd.ExcelFile(closed_path)
        c_frames = [c_excel.parse(s) for s in c_excel.sheet_names]
        df_closed = pd.concat(c_frames, ignore_index=True)
        df_closed.columns = [str(c).strip() for c in df_closed.columns]
    else:
        print(f"WARNING: {CLOSED_FILE} not found. Closed metrics will be zero.")

    return df_pending, df_closed


def prepare_pending(df):
    group_col = find_column(df, "Assignment group")
    created_col = find_column(df, "Created")

    if group_col is None or created_col is None:
        raise ValueError(f"Pending file missing 'Assignment group' or 'Created'. Found: {list(df.columns)}")

    df["Partner"] = df[group_col].apply(clean_partner)
    df["PartnerRaw"] = df[group_col].apply(raw_partner)
    df["Category"] = df.apply(dashboard_category, axis=1)
    df["WorkType"] = df.apply(work_type_category, axis=1)

    created = pd.to_datetime(df[created_col], errors="coerce", format="ISO8601")
    bad = created.isna()
    if bad.any():
        created.loc[bad] = pd.to_datetime(df.loc[bad, created_col], errors="coerce", dayfirst=True)

    df["_Created"] = created
    df["Age"] = (pd.Timestamp(REPORT_TIME) - created).dt.days.fillna(0).clip(lower=0).astype(int)

    cutoff_dt = REPORT_TIME.replace(hour=CUTOFF_HOUR, minute=0, second=0, microsecond=0)
    df["AfterCutoff"] = (created >= cutoff_dt).fillna(False).astype(bool)

    return df


def prepare_closed(cdf):
    if cdf is None or cdf.empty:
        return None, None
    group_col = find_column(cdf, "Assignment group")
    if not group_col:
        print("WARNING: Closed file missing 'Assignment group' column.")
        return None, None

    cdf["Partner"] = cdf[group_col].apply(clean_partner)
    cdf["PartnerRaw"] = cdf[group_col].apply(raw_partner)
    cdf["WorkType"] = cdf.apply(work_type_category, axis=1)

    team_col = None
    for name in TEAM_COLUMNS:
        team_col = find_column(cdf, name)
        if team_col:
            break

    return cdf, team_col


def summarize(sub):
    ages = sub["Age"]
    total = len(sub)
    fresh = int((ages <= FRESH_MAX).sum())
    watch = int(((ages > FRESH_MAX) & (ages <= WATCH_MAX)).sum())
    old = int((ages > WATCH_MAX).sum())
    aged = watch + old
    counts = ages.clip(upper=HEATMAP_MAX_AGE).value_counts().reindex(range(HEATMAP_MAX_AGE + 1), fill_value=0).tolist()
    return {
        "total": total, "fresh": fresh, "watch": watch, "old": old,
        "aged": aged, "pct_aged": pct(aged, total), "counts": counts
    }


def load_history(path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def previous_snapshot(history, today_key):
    keys = sorted(k for k in history if k < today_key)
    return (history[keys[-1]], keys[-1]) if keys else (None, None)


# ============================================================
# HTML BUILDING BLOCKS
# ============================================================

CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
* { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body { 
  background: #DDE5EE; 
  font-family: "Segoe UI", Calibri, Arial, sans-serif; 
  color: #12233F; 
  padding: 12px;
}
.page { 
  width: 100%; 
  max-width: 1000px; 
  margin: 0 auto 24px auto; 
  background: #EDF3FA; 
  position: relative; 
  display: flex; 
  flex-direction: column; 
  box-shadow: 0 6px 24px rgba(10,30,60,.18); 
  border-radius: 8px;
  overflow: hidden;
}
@media print {
  @page { size: 1000px 1333px; margin: 0; }
  body { background: #fff; padding: 0; }
  .page { 
    width: 1000px; 
    height: 1333px; 
    margin: 0; 
    box-shadow: none; 
    page-break-after: always; 
    break-after: page; 
    border-radius: 0;
  }
  .page:last-child { page-break-after: auto; break-after: auto; }
  .row { flex-wrap: nowrap !important; }
}
.hdr { height: 128px; background: linear-gradient(100deg, #0A2A5E 0%, #123F82 60%, #1B5DA8 100%);
       color: #fff; position: relative; padding: 24px 30px 0 40px; flex: none; overflow: hidden; }
.hdr .tag { position: absolute; left: 14px; top: 0; background: #1F9D55; color: #fff; font-size: 13px;
            font-weight: 600; padding: 4px 12px; border-radius: 0 0 8px 8px; }
.hdr h1 { font-size: 40px; font-weight: 800; letter-spacing: -.5px; margin-top: 8px; line-height: 1.1; }
.hdr h1 span { color: var(--accent); }
.hdr .sub { font-size: 17px; margin-top: 8px; opacity: .95; }
.hdr .tagline { position: absolute; right: 92px; top: 52px; font-size: 14px; line-height: 1.35; max-width: 210px; }
.hdr svg.tower { position: absolute; right: 8px; top: 8px; height: 118px; opacity: .85; }
.body { padding: 10px 12px 8px; display: flex; flex-direction: column; gap: 10px; flex: 1; min-height: 0; }
.row { display: flex; gap: 10px; flex-wrap: wrap; }
.row.grow { flex: 1; }
.card { background: #fff; border: 1px solid #D5E1EF; border-radius: 12px; padding: 12px 16px; min-width: 280px; flex: 1; }
.card h3 { font-size: 20px; font-weight: 700; color: #12233F; }
.card h3 small { font-size: 16px; font-weight: 400; color: #4B5E78; }
.kpi { display: flex; align-items: center; gap: 14px; height: 124px; }
.kpi .ico { width: 54px; height: 62px; border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 30px; color: #fff; flex: none; }
.kpi .lbl { font-size: 20px; font-weight: 700; }
.kpi .big { font-size: 68px; font-weight: 800; line-height: 1; letter-spacing: -2px; }
.kpi .chg { font-size: 22px; font-weight: 800; }
.kpi .chg small { display: block; font-size: 13px; font-weight: 400; color: #4B5E78; }
.kpi.alert { background: #FDECEC; border-color: #F6C6C6; }
.kpi.alert .lbl, .kpi.alert .big { color: #E0242B; }
.kpi.alert .ico { background: #E0242B; }
.kpi .pctof { font-size: 16px; font-weight: 700; color: #E0242B; margin-top: 6px; }
.tiles { display: flex; gap: 8px; flex-wrap: wrap; }
.tile { flex: 1; min-width: 130px; border-radius: 10px; overflow: hidden; background: #fff; border: 1px solid #D5E1EF; text-align: center; }
.tile .top { color: #fff; padding: 8px 0 4px; font-weight: 700; font-size: 15px; }
.tile .top i { display: block; font-style: normal; font-size: 22px; margin-bottom: 2px; }
.tile .n { font-size: 32px; font-weight: 800; padding-top: 6px; }
.tile .p { font-size: 14px; color: #4B5E78; padding-bottom: 8px; }
.donut-wrap { display: flex; align-items: center; gap: 18px; margin-top: 6px; flex-wrap: wrap; }
.donut { position: relative; width: 190px; height: 190px; flex: none; }
.donut svg { width: 100%; height: 100%; transform: rotate(0deg); }
.donut .mid { position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; }
.donut .mid b { font-size: 31px; line-height: 1; }
.donut .mid span { font-size: 15px; font-weight: 600; line-height: 1.15; margin-top: 2px; }
.legend div { display: flex; align-items: center; gap: 8px; font-size: 16px; line-height: 1.12; margin: 5px 0; }
.legend i { width: 13px; height: 13px; border-radius: 3px; flex: none; }
.legend b { font-weight: 700; }
.agebar { display: flex; height: 30px; border-radius: 8px; overflow: hidden; margin: 14px 0 10px; }
.agecols { display: flex; text-align: center; }
.agecols div { flex: 1; font-size: 15px; color: #4B5E78; line-height: 1.25; }
.agecols b { display: block; font-size: 26px; }
.callout { display: flex; align-items: center; gap: 14px; margin-top: 12px; background: #FDECEC; border: 1px solid #F6C6C6; border-radius: 10px; padding: 10px 14px; color: #D0222A; }
.callout .big { font-size: 33px; font-weight: 800; }
.callout .txt { font-size: 16px; font-weight: 600; line-height: 1.2; flex: 1; }
.callout .rq { font-size: 13px; font-weight: 600; width: 100px; border-left: 2px solid #F0B3B3; padding-left: 10px; line-height: 1.2; }
.rank { display: flex; align-items: center; gap: 10px; margin: 9px 0; font-size: 17px; }
.rank .bdg { width: 25px; height: 25px; border-radius: 50%; border: 2px solid #F29B3C; background: #FFF4E5; color: #B25E00; font-weight: 700; font-size: 14px; display: flex; align-items: center; justify-content: center; flex: none; }
.rank .nm { width: 150px; flex: none; font-size: 17px; }
.rank .bar { flex: 1; height: 24px; position: relative; }
.rank .bar i { position: absolute; left: 0; top: 0; bottom: 0; border-radius: 2px; }
.rank .val { font-weight: 800; min-width: 44px; text-align: right; }
.rank.sm { margin: 8px 0; font-size: 14px; }
.rank.sm .nm { width: 120px; font-size: 14px; }
.rank.sm .bar { height: 20px; }
.rank.sm .bdg { width: 21px; height: 21px; font-size: 12px; }
.good { color: #1B8E4B; font-weight: 800; }
.good small { font-weight: 600; margin-left: 6px; }
.panel { border-radius: 12px; overflow: hidden; background: #fff; border: 1px solid #D5E1EF; flex: 1; min-width: 280px; }
.panel .ph { color: #fff; font-size: 21px; font-weight: 700; padding: 11px 16px; display: flex; align-items: center; gap: 10px; }
.panel .pb { padding: 6px 16px 10px; }
.item { display: flex; gap: 12px; align-items: flex-start; padding: 8px 0; }
.item .no { width: 30px; height: 30px; border-radius: 50%; color: #fff; font-weight: 700; font-size: 16px; display: flex; align-items: center; justify-content: center; flex: none; margin-top: 2px; }
.item b { display: block; font-size: 17px; line-height: 1.2; }
.item span { font-size: 14.5px; color: #3E5068; line-height: 1.25; display: block; }
.item.win .no { background: #1F9D55; font-size: 17px; }
.foot { font-size: 12px; color: #4B5E78; padding: 0 6px 4px; }
table.tbl { width: 100%; border-collapse: collapse; font-size: 12.5px; }
table.tbl th { background: #E6EEF8; color: #12233F; padding: 6px 4px; font-weight: 700; border-bottom: 2px solid #C9D8EA; }
table.tbl td { padding: 0 4px; height: 26px; text-align: center; border-bottom: 1px solid #E3EBF4; }
table.tbl td.l, table.tbl th.l { text-align: left; padding-left: 8px; }
table.tbl tr.oth td { background: #F4F7FB; }
table.tbl tr.tot td { background: #E6EEF8; font-weight: 800; border-top: 2px solid #C9D8EA; }
table.heat { table-layout: fixed; }
table.heat td { padding: 0; height: 28px; font-size: 13px; }
table.heat th { padding: 6px 0; font-size: 13px; }
.red { color: #D0222A; font-weight: 800; }
.legend2 { display: flex; gap: 14px; font-size: 13px; float: right; }
.legend2 span i { display: inline-block; width: 13px; height: 13px; border-radius: 3px; margin-right: 4px; vertical-align: -2px; border: 1px solid rgba(0,0,0,.2); }
.take { display: flex; gap: 12px; padding: 9px 0; align-items: flex-start; font-size: 15px; line-height: 1.3; }
.take .no { width: 26px; height: 26px; border-radius: 50%; background: #1E6FD0; color: #fff; font-weight: 700; font-size: 14px; display: flex; align-items: center; justify-content: center; flex: none; }
"""

TOWER_SVG = """
<svg class="tower" viewBox="0 0 100 130" xmlns="http://www.w3.org/2000/svg" fill="none" stroke="#9EC7F2" stroke-width="1.6">
  <path d="M50 14 L26 126 M50 14 L74 126"/>
  <path d="M42 50 L58 50 M38 70 L62 70 M34 92 L66 92 M30 112 L70 112"/>
  <path d="M42 50 L62 70 M58 50 L38 70 M38 70 L66 92 M62 70 L34 92 M34 92 L70 112 M66 92 L30 112"/>
  <circle cx="50" cy="12" r="3" fill="#FF6B6B" stroke="none"/>
  <path d="M38 8 q12 -12 24 0 M32 4 q18 -18 36 0" stroke="#7FB5EE"/>
</svg>
"""


def header(page, title, accent, accent_colour, subtitle, tagline):
    return f"""
<div class="hdr" style="--accent:{accent_colour}">
  <div class="tag">Page {page} of 3</div>
  <h1>{title} <span>{accent}</span></h1>
  <div class="sub">{subtitle}</div>
  <div class="tagline">{tagline}</div>
  {TOWER_SVG}
</div>"""


def rank_rows(items, colour, small=False):
    if not items:
        return '<div style="color:#6B7C93;padding:10px 0">Nothing to show.</div>'
    top = max(v for _, v, _ in items) or 1
    out = []
    for i, (name, value, label) in enumerate(items, 1):
        width = max(4, value / top * 100)
        out.append(
            f'<div class="rank{" sm" if small else ""}"><div class="bdg">{i}</div>'
            f'<div class="nm">{esc(name)}</div>'
            f'<div class="bar"><i style="width:{width:.0f}%;background:{colour}"></i></div>'
            f'<div class="val">{label}</div></div>'
        )
    return "".join(out)


# ============================================================
# MAIN EXECUTABLE & RENDER
# ============================================================

def main():
    print()
    print("==============================================")
    print("      PARTNER OPERATIONS PULSE DASHBOARD")
    print("==============================================")
    print()

    df_raw, cdf_raw = load_two_files()
    df = prepare_pending(df_raw)
    cdf, team_col = prepare_closed(cdf_raw)

    total_tickets = len(df)
    stamp = REPORT_TIME.strftime("%d %b %Y, %H:%M")
    today_key = REPORT_TIME.strftime("%Y-%m-%d")

    overall = summarize(df)
    cat_stats = {c: summarize(df[df["Category"] == c]) for c in CATEGORIES}
    partner_stats = {p: summarize(g) for p, g in df.groupby("Partner")}

    by_total = sorted(partner_stats, key=lambda p: (-partner_stats[p]["total"], p.lower()))
    top_names = by_total[:TOP_N_PARTNERS]
    rest_names = by_total[TOP_N_PARTNERS:]
    other_stat = summarize(df[df["Partner"].isin(rest_names)]) if rest_names else None

    top_aged = [p for p in sorted(partner_stats, key=lambda p: (-partner_stats[p]["aged"], p.lower())) if partner_stats[p]["aged"] > 0][:5]
    eligible = [p for p in partner_stats if partner_stats[p]["total"] >= MIN_TICKETS_FOR_PCT_RANKING]
    top_pct = [p for p in sorted(eligible, key=lambda p: (-partner_stats[p]["pct_aged"], p.lower())) if partner_stats[p]["pct_aged"] > 0][:5]

    focus_counts = df[df["Category"] == FOCUS_CATEGORY].groupby("Partner").size()
    focus_top = sorted(focus_counts.items(), key=lambda kv: (-kv[1], kv[0].lower()))[:3]

    history_path = DATA_DIR / HISTORY_FILE
    history = load_history(history_path)
    prev, prev_key = previous_snapshot(history, today_key)

    improvements = []
    change_pct = None

    if prev:
        prev_total = prev.get("total", 0)
        change_pct = pct(total_tickets - prev_total, prev_total) if prev_total else None
        for p in set(partner_stats) | set(prev.get("partners", {})):
            now = partner_stats.get(p, {}).get("total", 0)
            before = prev["partners"].get(p, 0)
            if before > 0 and now < before:
                improvements.append((p, now - before, (now - before) / before * 100))
        improvements.sort(key=lambda x: (x[1], x[2]))
        improvements = improvements[:3]

    history[today_key] = {
        "saved_at": REPORT_TIME.strftime("%Y-%m-%d %H:%M"),
        "total": total_tickets,
        "partners": {p: s["total"] for p, s in partner_stats.items()},
    }

    try:
        history_path.write_text(json.dumps(history, indent=1), encoding="utf-8")
    except Exception:
        pass

    prev_label = datetime.strptime(prev_key, "%Y-%m-%d").strftime("%d %b") if prev_key else ""

    # Totals for KPIs (Total Closed, Before 4PM Pending, After 4PM Pending)
    total_closed_kpi = len(cdf) if cdf is not None else 0
    total_before_kpi = int((~df["AfterCutoff"]).sum())
    total_after_kpi = int(df["AfterCutoff"].sum())

    # ============================================================
    # PAGE 1 RENDER
    # ============================================================

    chg_html = (
        f'<div class="chg" style="color:#E0242B">&#9650; +{change_pct:.0f}%<small>vs {prev_label}*</small></div>' if change_pct and change_pct > 0 else
        (f'<div class="chg" style="color:#1B8E4B">&#9660; {change_pct:.0f}%<small>vs {prev_label}*</small></div>' if change_pct and change_pct < 0 else
         '<div class="chg" style="color:#6B7C93">n/a<small>first snapshot</small></div>')
    )

    kpis_main = f"""
<div class="row">
  <div class="card kpi" style="flex:1.15">
    <div class="ico" style="background:#1E6FD0">&#128196;</div>
    <div><div class="lbl">Total Pending Tickets</div>
      <div style="display:flex;align-items:flex-end;gap:14px"><div class="big">{num(total_tickets)}</div>{chg_html}</div></div>
  </div>
  <div class="card kpi alert" style="flex:1">
    <div class="ico">&#128339;</div>
    <div><div class="lbl">&ge; {ESCALATION_FROM} Days (Escalation)</div>
      <div style="display:flex;align-items:flex-end;gap:14px"><div class="big">{num(overall['aged'])}</div>
      <div class="pctof">{overall['pct_aged']:.1f}% of total</div></div></div>
  </div>
</div>"""

    kpis_cutoff = f"""
<div class="row">
  <div class="card kpi" style="flex:1"><div class="ico" style="background:#1FA34A">&#10004;</div><div><div class="lbl">Total Closed</div><div class="big">{num(total_closed_kpi)}</div></div></div>
  <div class="card kpi" style="flex:1"><div class="ico" style="background:#1E6FD0">&#128339;</div><div><div class="lbl">Before 4PM Pending</div><div class="big">{num(total_before_kpi)}</div></div></div>
  <div class="card kpi" style="flex:1"><div class="ico" style="background:#78909C">&#128347;</div><div><div class="lbl">After 4PM Pending</div><div class="big">{num(total_after_kpi)}</div></div></div>
</div>"""

    tiles = '<div class="tiles">' + "".join(
        f'<div class="tile"><div class="top" style="background:{CAT_COLOR[c]}"><i>{CAT_ICON[c]}</i>{c}</div>'
        f'<div class="n" style="color:{CAT_COLOR[c]}">{num(cat_stats[c]["total"])}</div>'
        f'<div class="p">{pct(cat_stats[c]["total"], total_tickets):.0f}% of total</div></div>'
        for c in CATEGORIES
    ) + "</div>"

    segments, cumulative = [], 0.0
    for c in CATEGORIES:
        share = pct(cat_stats[c]["total"], total_tickets)
        if share > 0:
            segments.append(
                f'<circle cx="21" cy="21" r="15.9155" fill="none" stroke="{CAT_COLOR[c]}" stroke-width="6.5" '
                f'stroke-dasharray="{share:.3f} {100 - share:.3f}" stroke-dashoffset="{25 - cumulative:.3f}"/>'
            )
        cumulative += share

    legend = "".join(
        f'<div><i style="background:{CAT_COLOR[c]}"></i><span><b>{c}</b><br>'
        f'{num(cat_stats[c]["total"])} ({pct(cat_stats[c]["total"], total_tickets):.0f}%)</span></div>'
        for c in CATEGORIES
    )

    backlog_card = f"""
<div class="card" style="flex:.95">
  <h3>Backlog by Category</h3>
  <div class="donut-wrap">
    <div class="donut">
      <svg viewBox="0 0 42 42"><circle cx="21" cy="21" r="15.9155" fill="none" stroke="#EEF2F7" stroke-width="6.5"/>{''.join(segments)}</svg>
      <div class="mid"><b>{num(total_tickets)}</b><span>pending<br>tickets</span></div>
    </div>
    <div class="legend">{legend}</div>
  </div>
</div>"""

    f_w, w_w, o_w = pct(overall["fresh"], total_tickets), pct(overall["watch"], total_tickets), pct(overall["old"], total_tickets)

    aging_card = f"""
<div class="card" style="flex:1.05">
  <h3>Ticket Aging Profile <small>(All Categories)</small></h3>
  <div class="agebar">
    <div style="width:{max(f_w, .5):.1f}%;background:linear-gradient(90deg,#1FA34A,#58C05A)"></div>
    <div style="width:{max(w_w, .5):.1f}%;background:linear-gradient(90deg,#F5D23B,#F2A93B)"></div>
    <div style="width:{max(o_w, .5):.1f}%;background:#EF3B2D"></div>
  </div>
  <div class="agecols">
    <div><b style="color:#1B8E4B">{num(overall['fresh'])}</b>0&ndash;{FRESH_MAX} days<br>({f_w:.1f}%)</div>
    <div><b style="color:#E08A00">{num(overall['watch'])}</b>{FRESH_MAX + 1}&ndash;{WATCH_MAX} days<br>({w_w:.1f}%)</div>
    <div><b style="color:#D0222A">{num(overall['old'])}</b>{WATCH_MAX + 1}+ days<br>({o_w:.1f}%)</div>
  </div>
  <div class="callout">
    <div class="big">{num(overall['aged'])}</div>
    <div class="txt">tickets are &ge; {ESCALATION_FROM} days old<br>{overall['pct_aged']:.1f}% of total backlog</div>
    <div class="rq">Requires focus &amp; acceleration</div>
  </div>
</div>"""

    focus_card = f"""
<div class="card" style="flex:1">
  <h3>Top 3 Partners by Total Backlog<br><small>({esc(FOCUS_CATEGORY)} category)</small></h3>
  {rank_rows([(n, v, num(v)) for n, v in focus_top], "linear-gradient(90deg,#F04B4B,#E0242B)")}
</div>"""

    impr_rows = rank_rows(
        [(p, abs(d), f'<span class="good">{d:+d}<small>({pc:.0f}%)</small></span>') for p, d, pc in improvements],
        "linear-gradient(90deg,#22A94B,#4CC866)"
    ) if improvements else '<div style="color:#6B7C93;padding:14px 0">First snapshot saved - daily changes show from next run.</div>'

    impr_card = f"""
<div class="card" style="flex:1.1">
  <h3>Top 3 Improvements<br><small>(vs previous run)</small></h3>
  {impr_rows}
</div>"""

    biggest_cat = max(CATEGORIES, key=lambda c: cat_stats[c]["total"])
    names_for_focus = ", ".join(top_aged) if top_aged else "the oldest queues"

    focus_panel = f"""
<div class="panel"><div class="ph" style="background:#133B80">&#9678; Today's Focus Areas</div><div class="pb">
  <div class="item"><div class="no" style="background:#E0242B">1</div><div><b>Clear &ge; {ESCALATION_FROM} day aged tickets ({num(overall['aged'])})</b><span>Focus on {esc(names_for_focus)}.</span></div></div>
  <div class="item"><div class="no" style="background:#1E6FD0">2</div><div><b>Accelerate {esc(biggest_cat.lower())} closures</b><span>{esc(biggest_cat)} is {pct(cat_stats[biggest_cat]['total'], total_tickets):.0f}% of total backlog ({num(cat_stats[biggest_cat]['total'])} tickets).</span></div></div>
  <div class="item"><div class="no" style="background:#1E6FD0">3</div><div><b>Review partner capacity and recovery plans</b><span>Address aging concentrations (see Page 2 heatmap).</span></div></div>
</div></div>"""

    wins_panel = f"""
<div class="panel"><div class="ph" style="background:#1B8E4B">&#127942; Wins to Celebrate</div><div class="pb">
  <div class="item win"><div class="no">&#10003;</div><div><span style="font-size:15px">Share of tickets ({f_w:.0f}%) are 0&ndash;{FRESH_MAX} days old.</span></div></div>
  <div class="item win"><div class="no">&#10003;</div><div><span style="font-size:15px">Relocation and PTMP remain low at {pct(cat_stats['Relocation']['total'], total_tickets):.0f}% and {pct(cat_stats['PTMP']['total'], total_tickets):.0f}% of total respectively.</span></div></div>
  <div class="item win"><div style="color:#1B8E4B;font-weight:800;font-size:19px;padding-left:42px">Let's keep the momentum!</div></div>
</div></div>"""

    page1 = f"""
<div class="page">
{header(1, "Daily Partner Operations", "Pulse", "#3DDC84", f"Pending Tickets &nbsp;|&nbsp; As at {stamp}", "Focus. Close. Deliver.<br>Better Customer Experience.")}
<div class="body">
  {kpis_main}
  {kpis_cutoff}
  {tiles}
  <div class="row grow">{backlog_card}{aging_card}</div>
  <div class="row grow">{focus_card}{impr_card}</div>
  <div class="row grow">{focus_panel}{wins_panel}</div>
  <div class="foot">*First snapshot saved - daily change will show from the next run.</div>
</div>
</div>"""

    # ============================================================
    # PAGE 2 RENDER
    # ============================================================

    max_aged = max([partner_stats[p]["aged"] for p in top_names] + [1])

    def summary_row(idx, label, s, cls=""):
        n = s["total"] or 1
        return (
            f'<tr class="{cls}"><td>{idx}</td><td class="l">{esc(label)}</td><td><b>{num(s["total"])}</b></td>'
            f'<td style="background:{colour_fresh(s["fresh"] / n)}">{num(s["fresh"])}</td>'
            f'<td style="background:{colour_watch(s["watch"] / n) if s["watch"] else "#fff"}">{num(s["watch"])}</td>'
            f'<td style="background:{colour_old(s["old"] / n) if s["old"] else "#fff"}">{num(s["old"])}</td>'
            f'<td style="background:{colour_aged_count(s["aged"], max_aged) if s["aged"] else "#fff"}">{num(s["aged"])}</td>'
            f'<td style="background:{colour_pct_aged(s["pct_aged"])}">{s["pct_aged"]:.0f}%</td></tr>'
        )

    srows = [summary_row(i, p, partner_stats[p]) for i, p in enumerate(top_names, 1)]
    if other_stat:
        srows.append(summary_row("", "Other partners", other_stat, "oth"))

    srows.append(
        f'<tr class="tot"><td></td><td class="l">Total</td><td>{num(total_tickets)}</td><td>{num(overall["fresh"])}</td>'
        f'<td>{num(overall["watch"])}</td><td>{num(overall["old"])}</td><td>{num(overall["aged"])}</td>'
        f'<td>{overall["pct_aged"]:.1f}%</td></tr>'
    )

    summary_card = f"""
<div class="card" style="flex:1.75;padding:10px 12px">
  <h3 style="margin-bottom:6px">Partner Aging Summary <small>(All Categories)</small></h3>
  <table class="tbl"><thead><tr><th>#</th><th class="l">Partner</th><th>Total</th>
    <th>0&ndash;{FRESH_MAX} days</th><th>{FRESH_MAX + 1}&ndash;{WATCH_MAX} days</th><th>{WATCH_MAX + 1}+ days</th>
    <th>&ge;{ESCALATION_FROM} days</th><th>% &ge;{ESCALATION_FROM}d</th></tr></thead>
  <tbody>{''.join(srows)}</tbody></table>
</div>"""

    side_cards = f"""
<div style="flex:1;display:flex;flex-direction:column;gap:10px">
  <div class="card" style="flex:1;padding:10px 12px"><h3 style="font-size:17px">Top 5 Partners by &ge; {ESCALATION_FROM} Days<br>(Aged Tickets)</h3>{rank_rows([(p, partner_stats[p]["aged"], f'<span style="font-weight:800">{partner_stats[p]["aged"]}</span>') for p in top_aged], "linear-gradient(90deg,#F04B4B,#E0242B)", small=True)}</div>
  <div class="card" style="flex:1;padding:10px 12px"><h3 style="font-size:17px">Top 5 Partners by % Aged (&ge;{ESCALATION_FROM}d)</h3>{rank_rows([(p, partner_stats[p]["pct_aged"], f'<span style="font-weight:800">{partner_stats[p]["pct_aged"]:.0f}%</span>') for p in top_pct], "linear-gradient(90deg,#F7B500,#F59B00)", small=True)}</div>
</div>"""

    heat_head = "".join(f'<th style="background:#1E5FA8;color:#fff">{a if a < HEATMAP_MAX_AGE else str(a) + "+"}</th>' for a in range(HEATMAP_MAX_AGE + 1))

    def heat_row(idx, label, s, cls=""):
        cells = [f"<td>{v}</td>" if cls else f'<td style="background:{heat_cell(a, v)[0]};color:{heat_cell(a, v)[1]}">{v}</td>' for a, v in enumerate(s["counts"])]
        return f'<tr class="{cls}"><td>{idx}</td><td class="l" style="white-space:nowrap">{esc(label)}</td>{"".join(cells)}<td><b>{num(s["total"])}</b></td><td class="{"red" if s["pct_aged"] >= 25 else ""}">{num(s["aged"])}</td></tr>'

    hrows = [heat_row(i, p, partner_stats[p]) for i, p in enumerate(top_names, 1)]
    if other_stat:
        hrows.append(heat_row("", "Other partners", other_stat, "oth"))
    hrows.append(heat_row("", "Total", overall, "tot"))

    heat_card = f"""
<div class="card" style="padding:10px 12px">
  <h3 style="margin-bottom:6px">Partner Aging Heatmap <small>(Number of Tickets by Age in Days)</small>
    <span class="legend2"><span><i style="background:#8DD58A"></i>Low</span><span><i style="background:#FFD966"></i>Moderate</span>
    <span><i style="background:#FF8A50"></i>High</span><span><i style="background:#EF4B45"></i>Very high</span></span></h3>
  <table class="tbl heat"><thead><tr><th style="width:26px">#</th><th class="l" style="width:150px">Partner</th>{heat_head}
    <th style="width:48px">Total</th><th style="width:44px">&ge;{ESCALATION_FROM}d</th></tr></thead>
  <tbody>{''.join(hrows)}</tbody></table>
</div>"""

    crows = [f'<tr><td class="l" style="color:{CAT_COLOR[c]};font-weight:700">{c}</td><td>{num(cat_stats[c]["total"])}</td><td>{num(cat_stats[c]["fresh"])}</td><td>{num(cat_stats[c]["watch"])}</td><td>{num(cat_stats[c]["old"])}</td><td>{num(cat_stats[c]["aged"])}</td><td style="background:{colour_pct_aged(cat_stats[c]["pct_aged"])};font-weight:700">{cat_stats[c]["pct_aged"]:.1f}%</td></tr>' for c in CATEGORIES]
    crows.append(f'<tr class="tot"><td class="l">Total</td><td>{num(total_tickets)}</td><td>{num(overall["fresh"])}</td><td>{num(overall["watch"])}</td><td>{num(overall["old"])}</td><td>{num(overall["aged"])}</td><td>{overall["pct_aged"]:.1f}%</td></tr>')

    cat_card = f"""
<div class="card" style="flex:1.05;padding:10px 12px">
  <h3 style="margin-bottom:6px">Aging by Category <small>(&ge; {ESCALATION_FROM} Days)</small></h3>
  <table class="tbl"><thead><tr><th class="l">Category</th><th>Total</th><th>0&ndash;{FRESH_MAX}d</th>
    <th>{FRESH_MAX + 1}&ndash;{WATCH_MAX}d</th><th>{WATCH_MAX + 1}+d</th><th>&ge;{ESCALATION_FROM}d</th><th>% &ge;{ESCALATION_FROM}d</th></tr></thead>
  <tbody>{''.join(crows)}</tbody></table>
</div>"""

    take_panel = f"""
<div class="panel" style="flex:.95"><div class="ph" style="background:#133B80">&#128202; Key Takeaways</div>
<div class="pb"><div class="take"><div class="no">1</div><div>{num(overall['aged'])} tickets ({overall['pct_aged']:.1f}%) are &ge; {ESCALATION_FROM} days old (escalation threshold).</div></div>
<div class="take"><div class="no">2</div><div>Focus partner recovery plans on queues with highest 3+ day concentrations (see heatmap).</div></div></div></div>"""

    page2 = f"""
<div class="page">
{header(2, "Partner Aging &amp; Accountability", "Heatmap", "#4FC3F7", f"Pending Tickets by Partner and Age (Days) &nbsp;|&nbsp; As at {stamp}", "Right Partners.<br>Faster Closures.<br>Happier Customers.")}
<div class="body">
  <div class="row">{summary_card}{side_cards}</div>
  {heat_card}
  <div class="row grow">{cat_card}{take_panel}</div>
</div>
</div>"""

    # ============================================================
    # PAGE 3 RENDER (Partner Output Performance - STRICT SUBSTRING FILTER FOR Adrian & Unassigned)
    # ============================================================

    all_partners = set(df["Partner"].unique()) | (set(cdf["Partner"].unique()) if cdf is not None else set())

    # Substring match to filter out any variant containing "adrian" or "unassigned"
    p3_partners = [
        p for p in all_partners 
        if not any(ex in str(p).lower() for ex in EXCLUDE_PARTNERS)
    ]

    partner_records = []
    tot_conn, tot_maint, tot_pend, tot_bef, tot_aft, tot_closed, tot_teams = 0, 0, 0, 0, 0, 0, set()

    for p in p3_partners:
        sub_p = df[df["Partner"] == p]
        c_sub = cdf[cdf["Partner"] == p] if cdf is not None else pd.DataFrame()

        p_conn = int((sub_p["WorkType"] == "Connection").sum())
        p_maint = int((sub_p["WorkType"] == "Maintenance").sum())
        p_pend = len(sub_p)
        p_before = int((~sub_p["AfterCutoff"]).sum())
        p_after = int(sub_p["AfterCutoff"].sum())
        p_closed = len(c_sub)

        p_teams = c_sub[team_col].dropna().astype(str).str.strip().nunique() if (cdf is not None and team_col and not c_sub.empty) else 0
        if cdf is not None and team_col and not c_sub.empty:
            tot_teams.update(c_sub[team_col].dropna().astype(str).str.strip().unique())

        tot_conn += p_conn
        tot_maint += p_maint
        tot_pend += p_pend
        tot_bef += p_before
        tot_aft += p_after
        tot_closed += p_closed

        prod = (p_closed // p_teams) if p_teams > 0 else "-"
        out_pct = (p_closed * 100 // (p_closed + p_before)) if (p_closed + p_before) > 0 else None

        raw_name = sub_p["PartnerRaw"].iloc[0] if not sub_p.empty else (c_sub["PartnerRaw"].iloc[0] if not c_sub.empty else p)

        partner_records.append({
            "partner": p,
            "raw_name": raw_name,
            "p_conn": p_conn,
            "p_maint": p_maint,
            "p_pend": p_pend,
            "p_before": p_before,
            "p_after": p_after,
            "p_closed": p_closed,
            "p_teams": p_teams,
            "prod": prod,
            "out_pct": out_pct,
        })

    # Sort descending by Output % (None values placed at the very end)
    partner_records.sort(key=lambda x: (x["out_pct"] is not None, x["out_pct"] if x["out_pct"] is not None else -1), reverse=True)

    p3_rows = []
    output_rankings = []

    for rec in partner_records:
        out_pct = rec["out_pct"]
        out_pct_str = f"{out_pct}%" if out_pct is not None else "-"

        # Highlight green for Output >= 70%
        if out_pct is not None and out_pct >= OUTPUT_GOOD:
            bg_col = "background:#28A745;color:#FFFFFF;"  # Solid Green
        elif out_pct is not None and out_pct <= OUTPUT_POOR:
            bg_col = "background:#F8D7DA;color:#721C24;"  # Soft Red
        else:
            bg_col = "background:#FFFFFF;"

        if out_pct is not None:
            output_rankings.append((rec["partner"], out_pct, f"{out_pct}%"))

        p3_rows.append(
            f'<tr><td class="l">{esc(rec["raw_name"])}</td>'
            f'<td>{num(rec["p_conn"])}</td><td>{num(rec["p_maint"])}</td>'
            f'<td><b>{num(rec["p_pend"])}</b></td><td>{num(rec["p_before"])}</td>'
            f'<td>{num(rec["p_after"])}</td><td><b>{num(rec["p_closed"])}</b></td>'
            f'<td>{num(rec["p_teams"]) if rec["p_teams"] else "-"}</td>'
            f'<td><b>{rec["prod"]}</b></td>'
            f'<td style="{bg_col}"><b>{out_pct_str}</b></td></tr>'
        )

    tot_prod = (tot_closed // len(tot_teams)) if len(tot_teams) > 0 else "-"
    tot_out_pct = (tot_closed * 100 // (tot_closed + tot_bef)) if (tot_closed + tot_bef) > 0 else None
    tot_out_pct_str = f"{tot_out_pct}%" if tot_out_pct is not None else "-"

    tot_row = (
        f'<tr class="tot"><td class="l">TOTAL</td><td>{num(tot_conn)}</td><td>{num(tot_maint)}</td><td>{num(tot_pend)}</td>'
        f'<td>{num(tot_bef)}</td><td>{num(tot_aft)}</td><td>{num(tot_closed)}</td><td>{len(tot_teams) if tot_teams else "-"}</td>'
        f'<td>{tot_prod}</td><td>{tot_out_pct_str}</td></tr>'
    )

    output_rankings.sort(key=lambda x: -x[1])
    top_output_bars = rank_rows(output_rankings[:5], "linear-gradient(90deg,#28A745,#4CD964)")

    page3 = f"""
<div class="page">
{header(3, "Partner Output", "Performance", "#FFB74D", f"Data as at {stamp}", "Measure.<br>Accelerate.<br>Succeed.")}
<div class="body">
  <div class="row grow">
    <div class="card" style="flex:1.8;padding:10px 12px;overflow-y:auto">
      <h3 style="margin-bottom:6px">Partner Productivity & Output % Table</h3>
      <table class="tbl">
        <thead><tr><th class="l">Partner</th><th>Conn</th><th>Maint</th><th>Pending</th><th>Before 4PM</th><th>After 4PM</th><th>Closed</th><th>Teams</th><th>Prod</th><th>Output %</th></tr></thead>
        <tbody>{''.join(p3_rows)}{tot_row}</tbody>
      </table>
    </div>
    <div class="card" style="flex:1;padding:10px 12px">
      <h3>Top 5 Partners by Output %</h3>
      {top_output_bars}
    </div>
  </div>
</div>
</div>"""

    # ============================================================
    # WRITE FILE & SERVE
    # ============================================================

    document = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Partner Operations Pulse - {stamp}</title>
<style>{CSS}</style></head>
<body>{page1}{page2}{page3}</body></html>"""

    html_path = DATA_DIR / OUTPUT_HTML

    try:
        html_path.write_text(document, encoding="utf-8")
    except PermissionError:
        print(f"ERROR: Please close {OUTPUT_HTML} and run again.")
        sys.exit(1)

    print(f"Dashboard saved : {html_path}")

    print()
    print("==============================================")
    print("                  DONE")
    print("==============================================")

    if SERVE:
        serve(html_path)


# ============================================================
# LOCAL WEB SERVER
# ============================================================

def serve(html_path):
    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            if self.path.split("?")[0] not in ("/", "/" + html_path.name):
                self.send_error(404)
                return

            body = html_path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def log_message(self, *args):
            pass

    server, port = None, PORT

    for port in range(PORT, PORT + 10):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            continue

    if server is None:
        print(f"ERROR: ports {PORT}-{PORT + 9} are all busy.")
        return

    url = f"http://localhost:{port}/"

    print()
    print(f"Open your dashboard: {url}")
    print("Press Ctrl+C to stop the server.")

    if OPEN_BROWSER:
        webbrowser.open(url)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()