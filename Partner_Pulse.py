"""
Partner Operations Pulse - 2-page dashboard built from pending tickets.

Reads Pending_Tickets.xlsx (next to this script), writes
Partner_Operations_Pulse.html and serves it on your own machine.
Run the script, then click the http://localhost link it prints.

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

PENDING_TICKETS_FILE = "Pending_Tickets.xlsx"
OUTPUT_HTML = "Partner_Operations_Pulse.html"

# Daily totals are stored here so "vs yesterday" can be calculated.
HISTORY_FILE = "pending_history.json"

# Report "as at" time. Use e.g. datetime(2026, 10, 8, 8, 10) to force a time.
REPORT_TIME = datetime.now()

# Age bands in whole days since creation
FRESH_MAX = 2                     # 0-2 days: on track
WATCH_MAX = 6                     # 3-6 days: ageing; 7+ days: critical
ESCALATION_FROM = FRESH_MAX + 1

HEATMAP_MAX_AGE = 16              # ages at or above this share the last column
TOP_N_PARTNERS = 10               # remaining partners are grouped as "Other partners"
MIN_TICKETS_FOR_PCT_RANKING = 10  # ignore tiny queues in "Top 5 by % aged"
FOCUS_CATEGORY = "Maintenance"    # category used in "Top 3 Partners by Backlog"

SERVE = True            # serve the report at a local http://localhost link
PORT = 8000             # next free port is used if this one is busy
OPEN_BROWSER = True     # also open the link automatically

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


# ============================================================
# HELPERS
# ============================================================

def esc(value):
    return html.escape(str(value))


def num(value):
    return f"{int(value):,}"


def pct(part, whole):
    return part / whole * 100 if whole else 0.0


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


def hex_to_rgb(color):
    color = color.lstrip("#")
    return tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))


def mix(c1, c2, t):
    t = max(0.0, min(1.0, t))
    a, b = hex_to_rgb(c1), hex_to_rgb(c2)

    return "#%02x%02x%02x" % tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def ramp(value, stops):
    """Colour for `value` along [(position, colour), ...] stops."""
    if value <= stops[0][0]:
        return stops[0][1]

    for (p1, c1), (p2, c2) in zip(stops, stops[1:]):
        if value <= p2:
            return mix(c1, c2, (value - p1) / (p2 - p1))

    return stops[-1][1]


# ============================================================
# CLASSIFICATION
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


# ============================================================
# DATA LOADING AND PREPARATION
# ============================================================

def load_pending(path):
    excel = pd.ExcelFile(path)
    print("Sheets found:", ", ".join(excel.sheet_names))

    collected_frames = []

    # Check for sheets explicitly named 'pending'
    pending_sheets = [s for s in excel.sheet_names if "pending" in s.lower()]

    if pending_sheets:
        for s in pending_sheets:
            print(f"Reading pending sheet: {s}")
            temp = excel.parse(s)
            if not temp.empty:
                collected_frames.append(temp)
    else:
        # Check all sheets for ticket data
        for name in excel.sheet_names:
            temp = excel.parse(name)
            if temp.empty:
                continue

            lookup = {str(c).strip().lower(): c for c in temp.columns}

            if "assignment group" in lookup and "created" in lookup:
                if "state" in lookup:
                    state = temp[lookup["state"]].fillna("").astype(str).str.upper()
                    # Keep all rows except explicitly completed/closed ones
                    active_rows = temp[~state.str.contains(r"RESOLV|CLOSE|CANCEL|COMPLET", regex=True, na=False)]
                    if not active_rows.empty:
                        print(f"Reading active rows from sheet: {name} ({len(active_rows)} rows)")
                        collected_frames.append(active_rows)
                else:
                    print(f"Reading all rows from sheet: {name} ({len(temp)} rows)")
                    collected_frames.append(temp)

    if not collected_frames:
        # Final fallback: concatenate all non-empty sheets in the workbook
        print("Fallback: Combining all non-empty sheets in workbook...")
        for name in excel.sheet_names:
            temp = excel.parse(name)
            if not temp.empty:
                collected_frames.append(temp)

    if not collected_frames:
        raise ValueError("Could not find any ticket data in the provided Excel file.")

    df = pd.concat(collected_frames, ignore_index=True)
    df.columns = [str(c).strip() for c in df.columns]

    return df


def prepare(df):
    group_col = find_column(df, "Assignment group")
    created_col = find_column(df, "Created")

    for label, col in [("Assignment group", group_col), ("Created", created_col)]:
        if col is None:
            raise ValueError(
                f"Pending data has no '{label}' column.\nColumns found: {list(df.columns)}"
            )

    df["Partner"] = df[group_col].apply(clean_partner)
    df["Category"] = df.apply(dashboard_category, axis=1)

    # Robust multi-format date parsing across all records
    created = pd.to_datetime(df[created_col], errors="coerce", format="ISO8601")

    bad = created.isna()
    if bad.any():
        created.loc[bad] = pd.to_datetime(df.loc[bad, created_col], errors="coerce")

    bad = created.isna()
    if bad.any():
        created.loc[bad] = pd.to_datetime(df.loc[bad, created_col], errors="coerce", dayfirst=True)

    bad = created.isna()
    if bad.any():
        created.loc[bad] = pd.to_datetime(
            df.loc[bad, created_col].astype(str), errors="coerce", format="mixed"
        )

    age = (pd.Timestamp(REPORT_TIME) - created).dt.days

    unknown = int(age.isna().sum())
    if unknown:
        print()
        print(f"WARNING: {unknown} ticket(s) have a 'Created' value that could not")
        print("be read as a date. They are counted in totals as age 0.")
        print("Sample values:", df.loc[age.isna(), created_col].head(5).tolist())
        print()

    df["Age"] = age.fillna(0).clip(lower=0).astype(int)

    return df


def summarize(sub):
    ages = sub["Age"]
    total = len(sub)

    fresh = int((ages <= FRESH_MAX).sum())
    watch = int(((ages > FRESH_MAX) & (ages <= WATCH_MAX)).sum())
    old = int((ages > WATCH_MAX).sum())
    aged = watch + old

    counts = (
        ages.clip(upper=HEATMAP_MAX_AGE)
        .value_counts()
        .reindex(range(HEATMAP_MAX_AGE + 1), fill_value=0)
        .tolist()
    )

    return {
        "total": total,
        "fresh": fresh,
        "watch": watch,
        "old": old,
        "aged": aged,
        "pct_aged": pct(aged, total),
        "counts": counts,
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
# COLOUR RULES
# ============================================================

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


def heat_cell(age, value):
    """Return (background, text colour) for one heatmap cell."""
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


# ============================================================
# HTML BUILDING BLOCKS
# ============================================================

CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
* { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
@page { size: 1000px 1333px; margin: 0; }
body { background: #DDE5EE; font-family: "Segoe UI", Calibri, Arial, sans-serif; color: #12233F; }
.page { width: 1000px; height: 1333px; margin: 24px auto; background: #EDF3FA; overflow: hidden;
        position: relative; display: flex; flex-direction: column; box-shadow: 0 6px 24px rgba(10,30,60,.18); }
@media print {
  body { background: #fff; }
  .page { margin: 0; box-shadow: none; page-break-after: always; break-after: page; }
  .page:last-child { page-break-after: auto; break-after: auto; }
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
.row { display: flex; gap: 10px; }
.row.grow { flex: 1; }
.card { background: #fff; border: 1px solid #D5E1EF; border-radius: 12px; padding: 12px 16px; }
.card h3 { font-size: 20px; font-weight: 700; color: #12233F; }
.card h3 small { font-size: 16px; font-weight: 400; color: #4B5E78; }
.kpi { display: flex; align-items: center; gap: 14px; height: 124px; }
.kpi .ico { width: 54px; height: 62px; border-radius: 10px; display: flex; align-items: center;
            justify-content: center; font-size: 30px; color: #fff; flex: none; }
.kpi .lbl { font-size: 20px; font-weight: 700; }
.kpi .big { font-size: 68px; font-weight: 800; line-height: 1; letter-spacing: -2px; }
.kpi .chg { font-size: 22px; font-weight: 800; }
.kpi .chg small { display: block; font-size: 13px; font-weight: 400; color: #4B5E78; }
.kpi.alert { background: #FDECEC; border-color: #F6C6C6; }
.kpi.alert .lbl, .kpi.alert .big { color: #E0242B; }
.kpi.alert .ico { background: #E0242B; }
.kpi .pctof { font-size: 16px; font-weight: 700; color: #E0242B; margin-top: 6px; }
.tiles { display: flex; gap: 8px; }
.tile { flex: 1; border-radius: 10px; overflow: hidden; background: #fff; border: 1px solid #D5E1EF; text-align: center; }
.tile .top { color: #fff; padding: 8px 0 4px; font-weight: 700; font-size: 15px; }
.tile .top i { display: block; font-style: normal; font-size: 22px; margin-bottom: 2px; }
.tile .n { font-size: 32px; font-weight: 800; padding-top: 6px; }
.tile .p { font-size: 14px; color: #4B5E78; padding-bottom: 8px; }
.donut-wrap { display: flex; align-items: center; gap: 18px; margin-top: 6px; }
.donut { position: relative; width: 190px; height: 190px; flex: none; }
.donut svg { width: 100%; height: 100%; transform: rotate(0deg); }
.donut .mid { position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center;
              justify-content: center; text-align: center; }
.donut .mid b { font-size: 31px; line-height: 1; }
.donut .mid span { font-size: 15px; font-weight: 600; line-height: 1.15; margin-top: 2px; }
.legend div { display: flex; align-items: center; gap: 8px; font-size: 16px; line-height: 1.12; margin: 5px 0; }
.legend i { width: 13px; height: 13px; border-radius: 3px; flex: none; }
.legend b { font-weight: 700; }
.agebar { display: flex; height: 30px; border-radius: 8px; overflow: hidden; margin: 14px 0 10px; }
.agecols { display: flex; text-align: center; }
.agecols div { flex: 1; font-size: 15px; color: #4B5E78; line-height: 1.25; }
.agecols b { display: block; font-size: 26px; }
.callout { display: flex; align-items: center; gap: 14px; margin-top: 12px; background: #FDECEC;
           border: 1px solid #F6C6C6; border-radius: 10px; padding: 10px 14px; color: #D0222A; }
.callout .big { font-size: 33px; font-weight: 800; }
.callout .txt { font-size: 16px; font-weight: 600; line-height: 1.2; flex: 1; }
.callout .rq { font-size: 13px; font-weight: 600; width: 100px; border-left: 2px solid #F0B3B3; padding-left: 10px; line-height: 1.2; }
.rank { display: flex; align-items: center; gap: 10px; margin: 9px 0; font-size: 17px; }
.rank .bdg { width: 25px; height: 25px; border-radius: 50%; border: 2px solid #F29B3C; background: #FFF4E5;
             color: #B25E00; font-weight: 700; font-size: 14px; display: flex; align-items: center;
             justify-content: center; flex: none; }
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
.panel { border-radius: 12px; overflow: hidden; background: #fff; border: 1px solid #D5E1EF; flex: 1; }
.panel .ph { color: #fff; font-size: 21px; font-weight: 700; padding: 11px 16px; display: flex; align-items: center; gap: 10px; }
.panel .pb { padding: 6px 16px 10px; }
.item { display: flex; gap: 12px; align-items: flex-start; padding: 8px 0; }
.item .no { width: 30px; height: 30px; border-radius: 50%; color: #fff; font-weight: 700; font-size: 16px;
            display: flex; align-items: center; justify-content: center; flex: none; margin-top: 2px; }
.item b { display: block; font-size: 17px; line-height: 1.2; }
.item span { font-size: 14.5px; color: #3E5068; line-height: 1.25; display: block; }
.item.win .no { background: #1F9D55; font-size: 17px; }
.foot { font-size: 12px; color: #4B5E78; padding: 0 6px 4px; }
table.tbl { width: 100%; border-collapse: collapse; font-size: 14px; }
table.tbl th { background: #E6EEF8; color: #12233F; padding: 6px 4px; font-weight: 700; border-bottom: 2px solid #C9D8EA; }
table.tbl td { padding: 0 4px; height: 29px; text-align: center; border-bottom: 1px solid #E3EBF4; }
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
.take .no { width: 26px; height: 26px; border-radius: 50%; background: #1E6FD0; color: #fff; font-weight: 700;
            font-size: 14px; display: flex; align-items: center; justify-content: center; flex: none; }
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
  <div class="tag">Page {page} of 2</div>
  <h1>{title} <span>{accent}</span></h1>
  <div class="sub">{subtitle}</div>
  <div class="tagline">{tagline}</div>
  {TOWER_SVG}
</div>"""


def rank_rows(items, colour, small=False):
    """Ranked rows with bars. items: [(name, value, label_html)]."""
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
# MAIN
# ============================================================

def main():
    print()
    print("==============================================")
    print("      PARTNER OPERATIONS PULSE DASHBOARD")
    print("==============================================")
    print()

    raw_path = DATA_DIR / PENDING_TICKETS_FILE

    if not raw_path.exists():
        print(f"ERROR: {raw_path} was not found.")
        print("Put Pending_Tickets.xlsx in the same folder as this script.")
        sys.exit(1)

    try:
        df = load_pending(raw_path)
    except PermissionError:
        print("ERROR: Please close Pending_Tickets.xlsx in Excel and run again.")
        sys.exit(1)

    df = prepare(df)
    total_tickets = len(df)

    if total_tickets == 0:
        print("ERROR: the Pending data has no rows.")
        sys.exit(1)

    print(f"Pending tickets : {total_tickets}")

    stamp = REPORT_TIME.strftime("%d %b %Y, %H:%M")
    today_key = REPORT_TIME.strftime("%Y-%m-%d")

    # ---------------- statistics ----------------
    overall = summarize(df)
    cat_stats = {c: summarize(df[df["Category"] == c]) for c in CATEGORIES}
    partner_stats = {p: summarize(g) for p, g in df.groupby("Partner")}

    by_total = sorted(partner_stats, key=lambda p: (-partner_stats[p]["total"], p.lower()))
    top_names = by_total[:TOP_N_PARTNERS]
    rest_names = by_total[TOP_N_PARTNERS:]
    other_stat = summarize(df[df["Partner"].isin(rest_names)]) if rest_names else None

    top_aged = [
        p for p in sorted(partner_stats, key=lambda p: (-partner_stats[p]["aged"], p.lower()))
        if partner_stats[p]["aged"] > 0
    ][:5]

    eligible = [p for p in partner_stats if partner_stats[p]["total"] >= MIN_TICKETS_FOR_PCT_RANKING]
    top_pct = [
        p for p in sorted(eligible, key=lambda p: (-partner_stats[p]["pct_aged"], p.lower()))
        if partner_stats[p]["pct_aged"] > 0
    ][:5]

    focus_counts = df[df["Category"] == FOCUS_CATEGORY].groupby("Partner").size()
    focus_top = sorted(focus_counts.items(), key=lambda kv: (-kv[1], kv[0].lower()))[:3]

    # ---------------- comparison with previous snapshot ----------------
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
    except Exception as e:
        print(f"(Could not save history: {e})")

    prev_label = datetime.strptime(prev_key, "%Y-%m-%d").strftime("%d %b") if prev_key else ""

    # ============================================================
    # PAGE 1
    # ============================================================

    if change_pct is None:
        chg_html = '<div class="chg" style="color:#6B7C93">n/a<small>first snapshot</small></div>'
    elif change_pct > 0:
        chg_html = f'<div class="chg" style="color:#E0242B">&#9650; +{change_pct:.0f}%<small>vs {prev_label}*</small></div>'
    elif change_pct < 0:
        chg_html = f'<div class="chg" style="color:#1B8E4B">&#9660; {change_pct:.0f}%<small>vs {prev_label}*</small></div>'
    else:
        chg_html = f'<div class="chg" style="color:#6B7C93">&#9644; 0%<small>vs {prev_label}*</small></div>'

    kpis = f"""
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

    tiles = '<div class="tiles">' + "".join(
        f'<div class="tile"><div class="top" style="background:{CAT_COLOR[c]}"><i>{CAT_ICON[c]}</i>{c}</div>'
        f'<div class="n" style="color:{CAT_COLOR[c]}">{num(cat_stats[c]["total"])}</div>'
        f'<div class="p">{pct(cat_stats[c]["total"], total_tickets):.0f}% of total</div></div>'
        for c in CATEGORIES
    ) + "</div>"

    # Donut: one stroked circle segment per category, offset by the running total.
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

    f_w = pct(overall["fresh"], total_tickets)
    w_w = pct(overall["watch"], total_tickets)
    o_w = pct(overall["old"], total_tickets)

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

    if improvements:
        impr_rows = rank_rows(
            [(p, abs(d), f'<span class="good">{d:+d}<small>({pc:.0f}%)</small></span>')
             for p, d, pc in improvements],
            "linear-gradient(90deg,#22A94B,#4CC866)",
        )
    elif prev:
        impr_rows = '<div style="color:#6B7C93;padding:14px 0">No partner reduced its pending tickets since the last snapshot.</div>'
    else:
        impr_rows = ('<div style="color:#6B7C93;padding:14px 0">Comparisons appear from the next daily run '
                     '(today is the first snapshot).</div>')

    impr_title = f"vs {prev_label} &ndash; overall pending" if prev else "vs previous run &ndash; overall pending"

    impr_card = f"""
<div class="card" style="flex:1.1">
  <h3>Top 3 Improvements<br><small>({impr_title})</small></h3>
  {impr_rows}
</div>"""

    names_for_focus = ", ".join(top_aged) if top_aged else "the oldest queues"
    biggest_cat = max(CATEGORIES, key=lambda c: cat_stats[c]["total"])

    focus_panel = f"""
<div class="panel"><div class="ph" style="background:#133B80">&#9678; Today's Focus Areas</div><div class="pb">
  <div class="item"><div class="no" style="background:#E0242B">1</div><div><b>Clear &ge; {ESCALATION_FROM} day aged tickets ({num(overall['aged'])})</b>
     <span>Focus on {esc(names_for_focus)}.</span></div></div>
  <div class="item"><div class="no" style="background:#1E6FD0">2</div><div><b>Accelerate {esc(biggest_cat.lower())} closures</b>
     <span>{esc(biggest_cat)} is {pct(cat_stats[biggest_cat]['total'], total_tickets):.0f}% of total backlog ({num(cat_stats[biggest_cat]['total'])} tickets).</span></div></div>
  <div class="item"><div class="no" style="background:#1E6FD0">3</div><div><b>Review partner capacity and recovery plans</b>
     <span>Address aging concentrations (see Page 2 heatmap).</span></div></div>
</div></div>"""

    wins = []

    if improvements:
        wins.append(f"{', '.join(p for p, _, _ in improvements)} reduced total pending tickets since {prev_label}.")

    wins.append(
        f"{'Majority of tickets' if f_w >= 50 else 'Share of tickets'} ({f_w:.0f}%) are 0&ndash;{FRESH_MAX} days old."
    )
    wins.append(
        f"Relocation and PTMP remain low at {pct(cat_stats['Relocation']['total'], total_tickets):.0f}% and "
        f"{pct(cat_stats['PTMP']['total'], total_tickets):.0f}% of total respectively."
    )

    win_items = "".join(
        f'<div class="item win"><div class="no">&#10003;</div><div><span style="font-size:15px">{w}</span></div></div>'
        for w in wins
    )

    wins_panel = f"""
<div class="panel"><div class="ph" style="background:#1B8E4B">&#127942; Wins to Celebrate</div><div class="pb">
  {win_items}
  <div class="item win"><div style="color:#1B8E4B;font-weight:800;font-size:19px;padding-left:42px">Let's keep the momentum!</div></div>
</div></div>"""

    if prev:
        footnote = f"*Change vs the {prev_label} snapshot ({prev.get('saved_at', '')})."
    else:
        footnote = "*First snapshot saved - daily change will show from the next run."

    page1 = f"""
<div class="page">
{header(1, "Daily Partner Operations", "Pulse", "#3DDC84", f"Pending Tickets &nbsp;|&nbsp; As at {stamp}", "Focus. Close. Deliver.<br>Better Customer Experience.")}
<div class="body">
  {kpis}
  {tiles}
  <div class="row grow">{backlog_card}{aging_card}</div>
  <div class="row grow">{focus_card}{impr_card}</div>
  <div class="row grow">{focus_panel}{wins_panel}</div>
  <div class="foot">{footnote}</div>
</div>
</div>"""

    # ============================================================
    # PAGE 2
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

    aged_list = rank_rows(
        [(p, partner_stats[p]["aged"], f'<span style="font-weight:800">{partner_stats[p]["aged"]}</span>')
         for p in top_aged],
        "linear-gradient(90deg,#F04B4B,#E0242B)", small=True,
    )

    pct_list = rank_rows(
        [(p, partner_stats[p]["pct_aged"], f'<span style="font-weight:800">{partner_stats[p]["pct_aged"]:.0f}%</span>')
         for p in top_pct],
        "linear-gradient(90deg,#F7B500,#F59B00)", small=True,
    )

    side_cards = f"""
<div style="flex:1;display:flex;flex-direction:column;gap:10px">
  <div class="card" style="flex:1;padding:10px 12px"><h3 style="font-size:17px">Top 5 Partners by &ge; {ESCALATION_FROM} Days<br>(Aged Tickets)</h3>{aged_list}</div>
  <div class="card" style="flex:1;padding:10px 12px"><h3 style="font-size:17px">Top 5 Partners by % Aged (&ge;{ESCALATION_FROM}d)</h3>{pct_list}</div>
</div>"""

    heat_head = "".join(
        f'<th style="background:#1E5FA8;color:#fff">{a if a < HEATMAP_MAX_AGE else str(a) + "+"}</th>'
        for a in range(HEATMAP_MAX_AGE + 1)
    )

    def heat_row(idx, label, s, cls=""):
        cells = []

        for a, v in enumerate(s["counts"]):
            if cls:
                cells.append(f"<td>{v}</td>")
            else:
                bg, fg = heat_cell(a, v)
                cells.append(f'<td style="background:{bg};color:{fg}">{v}</td>')

        aged_cls = "red" if s["pct_aged"] >= 25 else ""

        return (
            f'<tr class="{cls}"><td>{idx}</td><td class="l" style="white-space:nowrap">{esc(label)}</td>'
            f'{"".join(cells)}<td><b>{num(s["total"])}</b></td><td class="{aged_cls}">{num(s["aged"])}</td></tr>'
        )

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

    crows = []

    for c in CATEGORIES:
        s = cat_stats[c]
        crows.append(
            f'<tr><td class="l" style="color:{CAT_COLOR[c]};font-weight:700">{c}</td><td>{num(s["total"])}</td>'
            f'<td>{num(s["fresh"])}</td><td>{num(s["watch"])}</td><td>{num(s["old"])}</td><td>{num(s["aged"])}</td>'
            f'<td style="background:{colour_pct_aged(s["pct_aged"])};font-weight:700">{s["pct_aged"]:.1f}%</td></tr>'
        )

    crows.append(
        f'<tr class="tot"><td class="l">Total</td><td>{num(total_tickets)}</td><td>{num(overall["fresh"])}</td>'
        f'<td>{num(overall["watch"])}</td><td>{num(overall["old"])}</td><td>{num(overall["aged"])}</td>'
        f'<td>{overall["pct_aged"]:.1f}%</td></tr>'
    )

    cat_card = f"""
<div class="card" style="flex:1.05;padding:10px 12px">
  <h3 style="margin-bottom:6px">Aging by Category <small>(&ge; {ESCALATION_FROM} Days)</small></h3>
  <table class="tbl"><thead><tr><th class="l">Category</th><th>Total</th><th>0&ndash;{FRESH_MAX}d</th>
    <th>{FRESH_MAX + 1}&ndash;{WATCH_MAX}d</th><th>{WATCH_MAX + 1}+d</th><th>&ge;{ESCALATION_FROM}d</th><th>% &ge;{ESCALATION_FROM}d</th></tr></thead>
  <tbody>{''.join(crows)}</tbody></table>
</div>"""

    takeaways = [
        f"{num(overall['aged'])} tickets ({overall['pct_aged']:.1f}%) are &ge; {ESCALATION_FROM} days old (escalation threshold)."
    ]

    if overall["aged"] and top_aged:
        group = top_aged
        group_aged = sum(partner_stats[p]["aged"] for p in group)
        joined = ", ".join(esc(p) for p in group[:-1]) + (" and " if len(group) > 1 else "") + esc(group[-1])
        takeaways.append(
            f"{joined} account for {num(group_aged)} of {num(overall['aged'])} aged tickets "
            f"({pct(group_aged, overall['aged']):.0f}%)."
        )

    big_cats = [c for c in CATEGORIES if c != "Other" and cat_stats[c]["total"] >= MIN_TICKETS_FOR_PCT_RANKING]

    if big_cats:
        worst = max(big_cats, key=lambda c: cat_stats[c]["pct_aged"])
        takeaways.append(f"{esc(worst)} has the highest proportion of aged tickets ({cat_stats[worst]['pct_aged']:.0f}%).")

    takeaways.append("Focus partner recovery plans on queues with the highest 3+ day concentrations (see heatmap).")

    take_html = "".join(
        f'<div class="take"><div class="no">{i}</div><div>{text}</div></div>'
        for i, text in enumerate(takeaways, 1)
    )

    take_panel = f"""
<div class="panel" style="flex:.95"><div class="ph" style="background:#133B80">&#128202; Key Takeaways</div>
<div class="pb">{take_html}</div></div>"""

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
    # WRITE FILES
    # ============================================================

    document = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<title>Partner Operations Pulse - {stamp}</title>
<style>{CSS}</style></head>
<body>{page1}{page2}</body></html>"""

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
    """Serve only the report on localhost; it is re-read on every refresh."""

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
        print(f"ERROR: ports {PORT}-{PORT + 9} are all busy. Change PORT and run again.")
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