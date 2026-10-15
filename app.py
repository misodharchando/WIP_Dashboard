import io
import math
import re
import html
import base64
import json

import pandas as pd
import streamlit as st

try:
    from streamlit_js_eval import streamlit_js_eval
except ImportError:
    streamlit_js_eval = None

st.set_page_config(
    page_title="R4 Line WIP Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

LEVELS = [
    ("High WIP", "Red", "#D32F2F", "#FFFFFF", "Store Violation - Over Issue"),
    ("Standard WIP", "Green", "#00A651", "#FFFFFF", "WIP Secured, No need to issue"),
    ("Reorder Level", "Yellow", "#FFD600", "#111827", "Need to issue to the line"),
    ("Minimum Level", "Gray", "#6B7280", "#FFFFFF", "Line will stop - Need Urgent Input"),
    ("Line Stop", "Black", "#111111", "#FFFFFF", "Line machines stopped"),
]
DEFAULT_THRESH = [1.18, 0.98, 0.78, 0.58]
TARGET_FACTOR = 1.18



# Browser local-storage keys. Files stay in the current browser/profile only.
PLAN_CACHE_KEY = "r4_wip_morning_commitment_v1"
WIP_CACHE_KEY = "r4_wip_system_report_v1"
MAX_BROWSER_CACHE_BYTES = 2_500_000  # conservative limit per file before Base64 expansion


class CachedUpload(io.BytesIO):
    """BytesIO carrying the .name attribute expected by the existing readers."""
    def __init__(self, data: bytes, name: str):
        super().__init__(data)
        self.name = name


def _js_string(value: str) -> str:
    return json.dumps(value)


def load_browser_file(storage_key: str, component_key: str):
    """Read one Base64 encoded uploaded file from this browser's localStorage."""
    if streamlit_js_eval is None:
        return None
    raw = streamlit_js_eval(
        js_expressions=f"localStorage.getItem({_js_string(storage_key)})",
        key=component_key,
        want_output=True,
    )
    if not raw:
        return None
    try:
        payload = json.loads(raw)
        return CachedUpload(base64.b64decode(payload["data"]), payload["name"])
    except (ValueError, KeyError, TypeError, base64.binascii.Error):
        return None


def save_browser_file(uploaded_file, storage_key: str, component_key: str):
    """Save an uploaded file to browser localStorage and return a status message."""
    if streamlit_js_eval is None or uploaded_file is None:
        return None
    data = uploaded_file.getvalue()
    if len(data) > MAX_BROWSER_CACHE_BYTES:
        return "too_large"
    payload = json.dumps({
        "name": uploaded_file.name,
        "data": base64.b64encode(data).decode("ascii"),
    }, separators=(",", ":"))
    result = streamlit_js_eval(
        js_expressions=(
            "try { localStorage.setItem(" + _js_string(storage_key) + ", "
            + _js_string(payload) + "); 'saved'; } "
            "catch (e) { 'error:' + e.name; }"
        ),
        key=component_key,
        want_output=True,
    )
    return result


def clear_browser_cache():
    if streamlit_js_eval is None:
        return None
    return streamlit_js_eval(
        js_expressions=(
            f"localStorage.removeItem({_js_string(PLAN_CACHE_KEY)});"
            f"localStorage.removeItem({_js_string(WIP_CACHE_KEY)});"
            "'cleared';"
        ),
        key="clear_r4_browser_cache",
        want_output=True,
    )

def band_name(n: int) -> str:
    return f"R4-BAND-{int(n):02d}"


def level_of(wip_days: float, thresholds) -> tuple:
    for level, threshold in zip(LEVELS[:4], thresholds):
        if wip_days > threshold:
            return level
    return LEVELS[4]


def plan_from_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    cols = {c.upper(): c for c in df.columns}

    required = ["BAND", "PLAN HOURS", "DAY PLAN"]
    missing = [c for c in required if c not in cols]
    if missing:
        raise ValueError(f"Morning commitment is missing columns: {', '.join(missing)}")

    band_index = list(df.columns).index(cols["BAND"])
    zone_col = None
    if "ZONE" in cols:
        zone_col = cols["ZONE"]
    elif band_index > 0:
        first_col = df.columns[0]
        if first_col.upper() not in ("BAND", "PLAN HOURS", "DAY PLAN"):
            zone_col = first_col

    zone = (
        df[zone_col].astype(object).where(df[zone_col].notna()).ffill()
        if zone_col else pd.Series("All Lines", index=df.index)
    )

    out = pd.DataFrame({
        "Zone": zone,
        "Band No": pd.to_numeric(df[cols["BAND"]], errors="coerce"),
        "Plan Hours": pd.to_numeric(df[cols["PLAN HOURS"]], errors="coerce"),
        "Day Plan": pd.to_numeric(df[cols["DAY PLAN"]], errors="coerce"),
    })
    out["Zone"] = out["Zone"].astype(str).str.strip().replace({"nan": "All Lines", "": "All Lines"})
    out = out[
        out["Band No"].notna()
        & (out["Band No"] % 1 == 0)
        & (out["Band No"] >= 1)
        & out["Day Plan"].notna()
        & (out["Day Plan"] > 5)
    ]
    out["Band No"] = out["Band No"].astype(int)
    return out.drop_duplicates("Band No").reset_index(drop=True)


def wip_from_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    cols = {c.upper(): c for c in df.columns}

    required = ["WORKSTATION", "WIP QTY"]
    missing = [c for c in required if c not in cols]
    if missing:
        raise ValueError(f"System report is missing columns: {', '.join(missing)}")

    out = pd.DataFrame({
        "Workstation": df[cols["WORKSTATION"]].astype(str).str.strip().str.upper(),
        "WIP Qty": pd.to_numeric(df[cols["WIP QTY"]], errors="coerce"),
    })
    return out[
        out["Workstation"].str.startswith("R4-BAND", na=False)
        & out["WIP Qty"].notna()
    ].reset_index(drop=True)


def read_any(file, sheet_hint=None) -> pd.DataFrame:
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file)
    xl = pd.ExcelFile(file)
    sheet = next(
        (s for s in xl.sheet_names if sheet_hint and sheet_hint.lower() in s.lower()),
        xl.sheet_names[0],
    )
    return xl.parse(sheet)


def calculate(plan: pd.DataFrame, wip: pd.DataFrame, thresholds) -> pd.DataFrame:
    wip_sum = wip.groupby("Workstation")["WIP Qty"].sum()
    rows = []
    for _, row in plan.iterrows():
        number = int(row["Band No"])
        name = band_name(number)
        target = float(row["Day Plan"])
        hours = float(row["Plan Hours"]) if pd.notna(row["Plan Hours"]) else 0.0
        hourly = target / hours if hours else 0.0
        system_wip = float(wip_sum.get(name, 0.0))
        wip_hours = system_wip / hourly if hourly else 0.0
        wip_days = system_wip / hourly / hours if hourly and hours else 0.0
        level = level_of(wip_days, thresholds)
        rows.append({
            "Zone": row["Zone"],
            "Band": name,
            "Combined Line": f"B-{number}-IN",
            "Day Pcs Target": target,
            "Day Planned Hrs": hours,
            "Hourly Output Target": hourly,
            "System Line WIP": system_wip,
            "WIP Hours": wip_hours,
            "WIP Days": wip_days,
            "Dashboard Target (pcs)": target * TARGET_FACTOR,
            "Level": level[0],
            "Colour": level[1],
            "Direction to Store": level[4],
        })
    return pd.DataFrame(rows)


def unmatched_bands(plan, wip):
    known = {band_name(n) for n in plan["Band No"]}
    return sorted(set(wip["Workstation"]) - known)


def standards(hourly: float, machines: float, run_hrs: float, in_extra: float = 220):
    d8, d9, d10, d11 = hourly * 4, hourly * 2, machines * 5, hourly * 1.5
    total = math.ceil(round(d8 + d9 + d10 + d11, 9))
    c15 = total + hourly * 2
    c16, d16 = total, c15
    d17, c17 = c16, c16 - hourly * 2
    d18, c18 = c17, c17 - hourly * 2
    d19 = c18

    components = pd.DataFrame({
        "Place": ["Waiting to Move to Line", "Input Rack WIP", "WIP Per Machine", "Endline & Bundling WIP"],
        "Standard": [
            "Issue for 4 hours max",
            "Max of 2 hrs of output capability of the line",
            "5 pcs x number of machines",
            "Max of 1.5 hrs of output capability of the line",
        ],
        "Calculated Qty": [d8, d9, d10, d11],
    })

    per_day = hourly * run_hrs or 1
    levels = pd.DataFrame({
        "Level": ["Maximum WIP", "Standard WIP", "Reorder Level", "Minimum Level", "Line Stop"],
        "Colour": ["Red", "Green", "Yellow", "Gray", "Black"],
        "From (pcs)": [c15, c16, c17, c18, None],
        "To (pcs)": [None, d16, d17, d18, d19],
    })
    levels["From (days)"] = levels["From (pcs)"] / per_day
    levels["To (days)"] = levels["To (pcs)"] / per_day

    j_in = math.ceil(hourly * 4 + hourly * 2 + machines * 5) + in_extra
    k_out = math.ceil(hourly * 2 + machines * 5 + hourly * 1.5)
    combined = {
        "IN": j_in, "OUT": k_out, "Total": j_in + k_out,
        "IN days": j_in / per_day,
        "OUT days": k_out / per_day,
        "Total days": (j_in + k_out) / per_day,
    }
    return components, total, levels, combined


COMPACT_CSS = """
<style>
/* Use almost the entire browser viewport */
.block-container{max-width:100%!important;padding:0.35rem 0.65rem 0.5rem!important}
[data-testid="stHeader"]{height:2rem;background:transparent}
[data-testid="stToolbar"]{right:.4rem}
[data-testid="stSidebar"]{min-width:315px;max-width:315px}
h1{font-size:1.35rem!important;line-height:1.15!important;margin:0 0 .15rem!important}
h2,h3{margin:.12rem 0!important;padding:0!important}
h3{font-size:.95rem!important}
[data-testid="stMetric"]{padding:.2rem .45rem!important}
[data-testid="stMetricLabel"]{font-size:.70rem!important}
[data-testid="stMetricValue"]{font-size:1.15rem!important}
.stTabs [data-baseweb="tab-list"]{gap:.25rem;height:2.25rem}
.stTabs [data-baseweb="tab"]{height:2.1rem;padding:0 .65rem;font-size:.78rem}
.stMultiSelect{margin-bottom:-.35rem}
hr{margin:.25rem 0!important}

.wt-kpis{display:grid;grid-template-columns:repeat(6,minmax(95px,1fr));gap:6px;margin:2px 0 5px}
.wt-kpi{border-radius:8px;padding:5px 8px;color:#fff;box-shadow:0 1px 3px rgba(0,0,0,.16)}
.wt-kpi-label{font-size:.61rem;font-weight:700;text-transform:uppercase;letter-spacing:.04em;opacity:.9}
.wt-kpi-value{font-size:1.15rem;font-weight:900;line-height:1.15}

.wt-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(var(--min,120px),1fr));gap:6px;margin:2px 0 7px}
.wt{border-radius:9px;overflow:hidden;box-shadow:0 2px 5px rgba(15,23,42,.20);color:var(--fg);background:var(--bg);border:1px solid rgba(255,255,255,.32);min-width:0}
.wt-top{display:flex;justify-content:space-between;align-items:center;gap:4px;padding:5px 7px 1px}
.wt-name{font-size:.67rem;font-weight:900;letter-spacing:.035em;white-space:nowrap}
.wt-badge{font-size:.51rem;font-weight:800;padding:1px 5px;border-radius:999px;background:rgba(255,255,255,.22);border:1px solid rgba(255,255,255,.38);white-space:nowrap}
.wt-body{display:grid;grid-template-columns:1fr 1px 1fr;align-items:center;padding:2px 7px}
.wt-div{height:27px;background:currentColor;opacity:.30}
.wt-col{text-align:center;min-width:0}
.wt-lbl{font-size:.50rem;font-weight:800;text-transform:uppercase;letter-spacing:.07em;opacity:.78}
.wt-tgt{font-size:.86rem;font-weight:700;line-height:1.05}
.wt-act{font-size:1.08rem;font-weight:900;line-height:1.05}
.wt-track{position:relative;height:5px;background:rgba(255,255,255,.28);border-radius:4px;margin:3px 7px 2px;overflow:hidden}
.wt-fill{height:100%;background:var(--fg);opacity:.82;border-radius:4px}
.wt-mark{position:absolute;top:0;bottom:0;width:2px;background:var(--fg);left:66.67%}
.wt-foot{display:flex;justify-content:space-between;gap:4px;padding:1px 7px 5px;font-size:.56rem;opacity:.9;white-space:nowrap}
.wt-big{font-size:1.40rem;font-weight:900;text-align:center;line-height:1;padding:4px 0 1px}
.wt-sub{text-align:center;font-size:.52rem;font-weight:700;padding-bottom:5px;text-transform:uppercase;letter-spacing:.06em;opacity:.84}
.wt-legend{display:flex;flex-wrap:wrap;gap:4px;margin:2px 0 5px}
.wt-chip{border-radius:999px;padding:2px 7px;font-size:.57rem;font-weight:800;border:1px solid rgba(0,0,0,.22)}

@media (max-width:1100px){
 .wt-kpis{grid-template-columns:repeat(3,1fr)}
 .wt-grid{grid-template-columns:repeat(auto-fit,minmax(110px,1fr))}
}
@media (max-width:700px){
 .block-container{padding:.25rem!important}
 .wt-kpis{grid-template-columns:repeat(2,1fr)}
 .wt-grid{grid-template-columns:repeat(2,minmax(0,1fr))}
 .wt-badge{display:none}
}
</style>
"""


def badge(level):
    return f"<span class='wt-badge'>{html.escape(level[0])}</span>"


def card(name, target, actual, days, level):
    pct = actual / target if target else 0
    fill = min(max(pct, 0), 1.5) / 1.5 * 100
    safe_name = html.escape(name.replace("R4-", ""))
    return f"""
    <div class='wt' style='--bg:{level[2]};--fg:{level[3]}'>
      <div class='wt-top'><span class='wt-name'>{safe_name}</span>{badge(level)}</div>
      <div class='wt-body'>
        <div class='wt-col'><div class='wt-lbl'>Target</div><div class='wt-tgt'>{target:,.0f}</div></div>
        <div class='wt-div'></div>
        <div class='wt-col'><div class='wt-lbl'>Actual</div><div class='wt-act'>{actual:,.0f}</div></div>
      </div>
      <div class='wt-track'><div class='wt-fill' style='width:{fill:.1f}%'></div><div class='wt-mark'></div></div>
      <div class='wt-foot'><span><b>{days:.2f}</b> days</span><span><b>{pct:.0%}</b> target</span></div>
    </div>"""


def days_card(name, days, level):
    safe_name = html.escape(name.replace("R4-", ""))
    return f"""
    <div class='wt' style='--bg:{level[2]};--fg:{level[3]}'>
      <div class='wt-top'><span class='wt-name'>{safe_name}</span>{badge(level)}</div>
      <div class='wt-big'>{days:.2f}</div><div class='wt-sub'>WIP days</div>
    </div>"""


def grid(cards, min_px=120):
    st.markdown(
        f"<div class='wt-grid' style='--min:{min_px}px'>{''.join(cards)}</div>",
        unsafe_allow_html=True,
    )


def legend(thresholds):
    bounds = [
        f"> {thresholds[0]:.2f}",
        f"> {thresholds[1]:.2f} to {thresholds[0]:.2f}",
        f"> {thresholds[2]:.2f} to {thresholds[1]:.2f}",
        f"> {thresholds[3]:.2f} to {thresholds[2]:.2f}",
        f"≤ {thresholds[3]:.2f}",
    ]
    chips = [
        f"<span class='wt-chip' style='background:{level[2]};color:{level[3]}'>{level[0]} ({bound} d)</span>"
        for level, bound in zip(LEVELS, bounds)
    ]
    st.markdown(f"<div class='wt-legend'>{''.join(chips)}</div>", unsafe_allow_html=True)


def style_level(df):
    colour_map = {level[1]: (level[2], level[3]) for level in LEVELS}

    def row_style(row):
        bg, fg = colour_map[row["Colour"]]
        # Fill the complete row, not only selected cells.
        return [f"background-color:{bg};color:{fg};font-weight:700" for _ in row.index]

    return df.style.apply(row_style, axis=1).format({
        "Day Pcs Target": "{:,.0f}",
        "Day Planned Hrs": "{:.1f}",
        "Hourly Output Target": "{:,.1f}",
        "System Line WIP": "{:,.0f}",
        "WIP Hours": "{:.2f}",
        "WIP Days": "{:.2f}",
    })


def colour_standard_rows(df):
    mapping = {level[1]: (level[2], level[3]) for level in LEVELS}

    def apply_colour(row):
        bg, fg = mapping.get(row["Colour"], ("#FFFFFF", "#111827"))
        return [f"background-color:{bg};color:{fg};font-weight:700" for _ in row.index]

    return df.style.apply(apply_colour, axis=1)


def to_excel(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="WIP LEVEL CALCULATIONS")
        worksheet = writer.sheets["WIP LEVEL CALCULATIONS"]
        for column in worksheet.columns:
            worksheet.column_dimensions[column[0].column_letter].width = min(
                max(len(str(cell.value or "")) for cell in column) + 2, 35
            )
    return buffer.getvalue()


st.markdown(COMPACT_CSS, unsafe_allow_html=True)
st.title("Line Wise WIP Dashboard - R4")

with st.sidebar:
    st.header("Data input")

    if streamlit_js_eval is None:
        st.error("Browser cache requires: pip install streamlit-js-eval")

    uploaded_f1 = st.file_uploader("Morning commitment", type=["xlsx", "csv"])
    uploaded_f2 = st.file_uploader("System report", type=["xlsx", "csv"])

    # Retrieve the last files saved by this browser. A newly uploaded file always wins.
    cached_f1 = load_browser_file(PLAN_CACHE_KEY, "load_plan_cache")
    cached_f2 = load_browser_file(WIP_CACHE_KEY, "load_wip_cache")
    f1 = uploaded_f1 or cached_f1
    f2 = uploaded_f2 or cached_f2

    if uploaded_f1 is not None:
        save_status_1 = save_browser_file(uploaded_f1, PLAN_CACHE_KEY, f"save_plan_{uploaded_f1.name}_{uploaded_f1.size}")
        if save_status_1 == "too_large":
            st.warning("Morning commitment loaded, but it is too large for browser cache.")
    elif cached_f1 is not None:
        st.caption(f"Restored from browser: {cached_f1.name}")

    if uploaded_f2 is not None:
        save_status_2 = save_browser_file(uploaded_f2, WIP_CACHE_KEY, f"save_wip_{uploaded_f2.name}_{uploaded_f2.size}")
        if save_status_2 == "too_large":
            st.warning("System report loaded, but it is too large for browser cache.")
    elif cached_f2 is not None:
        st.caption(f"Restored from browser: {cached_f2.name}")

    if st.button("Clear saved browser data", use_container_width=True):
        st.session_state["clear_browser_cache_requested"] = True

    if st.session_state.get("clear_browser_cache_requested"):
        clear_result = clear_browser_cache()
        if clear_result == "cleared":
            st.session_state.pop("clear_browser_cache_requested", None)
            st.success("Saved browser data cleared.")

    st.caption("Saved locally in this browser profile. Nothing is written to the server database.")

    plan_df = pd.DataFrame(columns=["Zone", "Band No", "Plan Hours", "Day Plan"])
    wip_df = pd.DataFrame(columns=["Workstation", "WIP Qty"])

    try:
        if f1:
            f1.seek(0)
            plan_df = plan_from_df(read_any(f1, "commit"))
        if f2:
            f2.seek(0)
            wip_df = wip_from_df(read_any(f2, "report"))
    except Exception as exc:
        st.error(f"Could not read the file: {exc}")

    st.header("WIP day thresholds")
    th = [
        st.number_input(label, value=value, step=0.01, format="%.2f")
        for label, value in zip(
            ["Red above", "Green above", "Yellow above", "Gray above"],
            DEFAULT_THRESH,
        )
    ]
    if not all(th[i] > th[i + 1] for i in range(3)):
        st.error("Thresholds must be in descending order.")
        st.stop()

if not (f1 and f2):
    st.info("Open the sidebar and upload the Morning commitment and System report files.")
    st.stop()

with st.expander("Edit uploaded data", expanded=False):
    left, right = st.columns(2)
    plan_df = left.data_editor(plan_df, num_rows="dynamic", use_container_width=True, key=f"plan_{f1.name}")
    wip_df = right.data_editor(wip_df, num_rows="dynamic", use_container_width=True, key=f"wip_{f2.name}")
    right.caption(f"Total uploaded system WIP: {pd.to_numeric(wip_df['WIP Qty'], errors='coerce').sum():,.0f} pcs")

plan_df["Band No"] = pd.to_numeric(plan_df["Band No"], errors="coerce")
plan_df["Day Plan"] = pd.to_numeric(plan_df["Day Plan"], errors="coerce")
plan_df["Plan Hours"] = pd.to_numeric(plan_df["Plan Hours"], errors="coerce")
plan_df = plan_df.dropna(subset=["Band No", "Day Plan"])
plan_df["Band No"] = plan_df["Band No"].astype(int)
wip_df["WIP Qty"] = pd.to_numeric(wip_df["WIP Qty"], errors="coerce")
wip_df = wip_df.dropna(subset=["Workstation", "WIP Qty"])
wip_df["Workstation"] = wip_df["Workstation"].astype(str).str.strip().str.upper()

calc = calculate(plan_df, wip_df, th)
if calc.empty:
    st.warning("No valid morning commitment rows were found.")
    st.stop()

bad = unmatched_bands(plan_df, wip_df)
if bad:
    st.warning("Ignored workstations without a morning commitment: " + ", ".join(bad))

kpis = [
    ("Bands", f"{len(calc)}", "#2563EB"),
    ("Total WIP", f"{calc['System Line WIP'].sum():,.0f}", "#7C3AED"),
    ("Avg WIP Days", f"{calc['WIP Days'].mean():.2f}", "#0891B2"),
    ("High WIP", f"{int((calc['Level'] == 'High WIP').sum())}", "#D32F2F"),
    ("Reorder", f"{int((calc['Level'] == 'Reorder Level').sum())}", "#D6A900"),
    ("Line Stop", f"{int((calc['Level'] == 'Line Stop').sum())}", "#111111"),
]
st.markdown(
    "<div class='wt-kpis'>" + "".join(
        f"<div class='wt-kpi' style='background:{bg}'><div class='wt-kpi-label'>{label}</div><div class='wt-kpi-value'>{value}</div></div>"
        for label, value, bg in kpis
    ) + "</div>",
    unsafe_allow_html=True,
)
legend(th)

tab1, tab2, tab3, tab4 = st.tabs([
    "WIP Dashboard (pcs)", "WIP Days", "Calculations", "WIP Standards"
])
zones = list(dict.fromkeys(calc["Zone"].fillna("All Lines").astype(str)))

with tab1:
    zsel = st.multiselect("Zone", zones, default=zones, key="zone_pcs", label_visibility="collapsed")
    filtered = calc[calc["Zone"].isin(zsel)]
    for zone in zsel:
        data = filtered[filtered["Zone"] == zone]
        st.subheader(zone or "All Lines")
        grid([
            card(
                row["Band"], row["Dashboard Target (pcs)"], row["System Line WIP"],
                row["WIP Days"], level_of(row["WIP Days"], th)
            )
            for _, row in data.iterrows()
        ], min_px=120)
    st.caption("Target = Day target × 1.18 | Actual = System line WIP | Full-card colour = WIP level")

with tab2:
    zsel2 = st.multiselect("Zone", zones, default=zones, key="zone_days", label_visibility="collapsed")
    filtered = calc[calc["Zone"].isin(zsel2)]
    for zone in zsel2:
        data = filtered[filtered["Zone"] == zone]
        st.subheader(zone or "All Lines")
        grid([
            days_card(row["Band"], row["WIP Days"], level_of(row["WIP Days"], th))
            for _, row in data.iterrows()
        ], min_px=105)

with tab3:
    show = calc.drop(columns=["Dashboard Target (pcs)"])
    st.dataframe(style_level(show), use_container_width=True, hide_index=True, height=520)
    st.download_button(
        "Download calculations (.xlsx)",
        to_excel(calc),
        "WIP_LEVEL_CALCULATIONS.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with tab4:
    c1, c2, c3, c4 = st.columns(4)
    hourly = c1.number_input("Hourly target", value=110.0, step=1.0)
    machines = c2.number_input("Machines", value=51.0, step=1.0)
    run_hrs = c3.number_input("Running hours", value=10.0, step=0.5)
    extra = c4.number_input("Extra IN pcs", value=220.0, step=10.0)

    components, total, levels, combined = standards(hourly, machines, run_hrs, extra)
    left, right = st.columns([1.05, 1])
    with left:
        st.markdown("**Standard WIP build-up**")
        st.dataframe(
            components.style.format({"Calculated Qty": "{:,.1f}"}),
            hide_index=True,
            use_container_width=True,
            height=178,
        )
        m1, m2 = st.columns(2)
        m1.metric("Total standard", f"{total:,} pcs")
        m2.metric("Standard days", f"{levels.loc[1, 'To (days)']:.2f}")
    with right:
        st.markdown("**Colour-filled WIP levels**")
        styled_levels = colour_standard_rows(levels).format(
            {c: "{:,.2f}" for c in levels.columns if "(" in c}, na_rep="-"
        )
        st.dataframe(styled_levels, hide_index=True, use_container_width=True, height=215)

    st.markdown("**Combined lines (IN + OUT)**")
    st.dataframe(
        pd.DataFrame([combined]).style.format("{:,.2f}"),
        hide_index=True,
        use_container_width=True,
        height=72,
    )
