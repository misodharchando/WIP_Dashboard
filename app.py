"""
R4 Line-wise WIP Dashboard  –  Streamlit version of the Excel workbook
"R4 WIP Dashboard (Standard for line WIP)".

Excel logic reproduced
----------------------
Per band (sheet "WIP LEVEL CALCULATIONS"):
    Day Pcs Target   = Day plan            (Morning commitment, VLOOKUP by band no.)
    Day planned HRS  = Plan hours          (Morning commitment)
    Hourly target    = Day Pcs Target / Day planned HRS
    System line WIP  = SUMIFS(system report WIP Qty, Workstation = band)
    WIP Hours        = WIP / Hourly target
    WIP Days         = WIP / Hourly target / Day planned HRS
Colour of WIP Days:  > 1.18 Red | > 0.98 Green | > 0.78 Yellow | > 0.58 Gray | else Black
Dashboard (pcs):     Target = Day Pcs Target * 1.18 ,  Actual = System line WIP
"WIP Standards" sheet: standard-WIP calculator for a line.

Run:  streamlit run app.py
"""
import io
import math
import re

import pandas as pd
import streamlit as st

st.set_page_config(page_title="R4 Line WIP Dashboard", page_icon="📊", layout="wide")

# --------------------------------------------------------------------------
# Defaults taken from the workbook
# --------------------------------------------------------------------------
# Levels: name, colour, hex, text colour, direction to store   (order = Excel rule priority)
LEVELS = [
    ("High WIP",       "Red",    "#FF0000", "#FFFFFF", "Store Violation - Over Issue"),
    ("Standard WIP",   "Green",  "#00B050", "#FFFFFF", "WIP Secured, No need to issue"),
    ("Reorder Level",  "Yellow", "#FFFF00", "#000000", "Need to issue to the line"),
    ("Minimum Level",  "Gray",   "#808080", "#FFFFFF", "Line will stop - Need Urgent Input"),
    ("Line Stop",      "Black",  "#000000", "#FFFFFF", "Line machines stopped"),
]
DEFAULT_THRESH = [1.18, 0.98, 0.78, 0.58]   # strictly-greater-than cut-offs (as in Excel CF rules)
TARGET_FACTOR = 1.18                        # "Target" on the pcs dashboard = Day target x 1.18


def band_name(n: int) -> str:
    return f"R4-BAND-{int(n):02d}"


def band_no(name) -> int | None:
    m = re.search(r"(\d+)\s*$", str(name))
    return int(m.group(1)) if m else None


def level_of(wip_days: float, th) -> tuple:
    """Same cascade as the Excel conditional formatting (strict '>' tests)."""
    for lvl, t in zip(LEVELS[:4], th):
        if wip_days > t:
            return lvl
    return LEVELS[4]


# --------------------------------------------------------------------------
# Input loading
# --------------------------------------------------------------------------
def plan_from_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise a 'Morning commitment' table (BAND, PLAN HOURS, DAY PLAN, optional zone)."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    cols = {c.upper(): c for c in df.columns}
    zone_col = df.columns[0] if df.columns[0].upper() not in ("BAND", "ZONE") and "BAND" in cols \
        and list(df.columns).index(cols["BAND"]) > 0 else cols.get("ZONE")
    out = pd.DataFrame({
        "Zone": df[zone_col].astype(object).where(df[zone_col].notna()).ffill() if zone_col else "",
        "Band No": pd.to_numeric(df[cols["BAND"]], errors="coerce"),
        "Plan Hours": pd.to_numeric(df[cols["PLAN HOURS"]], errors="coerce"),
        "Day Plan": pd.to_numeric(df[cols["DAY PLAN"]], errors="coerce"),
    })
    out["Zone"] = out["Zone"].astype(str).str.strip().replace("nan", "")
    out = out[out["Band No"].notna() & (out["Band No"] % 1 == 0) & (out["Band No"] >= 1)
              & out["Day Plan"].notna() & (out["Day Plan"] > 5)]   # drops the stray helper rows
    out["Band No"] = out["Band No"].astype(int)
    return out.drop_duplicates("Band No").reset_index(drop=True)


def wip_from_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    cols = {c.upper(): c for c in df.columns}
    out = pd.DataFrame({
        "Workstation": df[cols["WORKSTATION"]].astype(str).str.strip(),
        "WIP Qty": pd.to_numeric(df[cols["WIP QTY"]], errors="coerce"),
    })
    return out[out["Workstation"].str.upper().str.startswith("R4-BAND") & out["WIP Qty"].notna()] \
        .reset_index(drop=True)


def read_any(file, sheet_hint=None) -> pd.DataFrame:
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    xl = pd.ExcelFile(file)
    sheet = next((s for s in xl.sheet_names if sheet_hint and sheet_hint.lower() in s.lower()),
                 xl.sheet_names[0])
    return xl.parse(sheet)


# --------------------------------------------------------------------------
# Core calculation  (== sheet "WIP LEVEL CALCULATIONS")
# --------------------------------------------------------------------------
def calculate(plan: pd.DataFrame, wip: pd.DataFrame, th) -> pd.DataFrame:
    wip_sum = wip.groupby("Workstation")["WIP Qty"].sum()
    rows = []
    for _, r in plan.iterrows():
        n = int(r["Band No"])
        name = band_name(n)
        target = float(r["Day Plan"])
        hrs = float(r["Plan Hours"]) if pd.notna(r["Plan Hours"]) else 0.0
        hourly = target / hrs if hrs else 0.0
        sys_wip = float(wip_sum.get(name, 0.0))
        wip_hours = sys_wip / hourly if hourly else 0.0
        wip_days = sys_wip / hourly / hrs if hourly and hrs else 0.0
        lvl = level_of(wip_days, th)
        rows.append({
            "Zone": r["Zone"], "Band": name, "Combined Line": f"B-{n}-IN",
            "Day Pcs Target": target, "Day Planned Hrs": hrs, "Hourly Output Target": hourly,
            "System Line WIP": sys_wip, "WIP Hours": wip_hours, "WIP Days": wip_days,
            "Dashboard Target (pcs)": target * TARGET_FACTOR,
            "Level": lvl[0], "Colour": lvl[1], "Direction to Store": lvl[4],
        })
    return pd.DataFrame(rows)


def unmatched_bands(plan, wip):
    known = {band_name(n) for n in plan["Band No"]}
    return sorted(set(wip["Workstation"]) - known)


# --------------------------------------------------------------------------
# WIP standards (== sheet "WIP Standards")
# --------------------------------------------------------------------------
def standards(hourly: float, machines: float, run_hrs: float, in_extra: float = 220):
    d8, d9, d10, d11 = hourly * 4, hourly * 2, machines * 5, hourly * 1.5
    total = math.ceil(round(d8 + d9 + d10 + d11, 9))
    c15 = total + hourly * 2           # max WIP (above)
    c16, d16 = total, c15              # standard
    d17 = c16; c17 = d17 - hourly * 2  # reorder
    d18 = c17; c18 = d18 - hourly * 2  # minimum
    d19 = c18                          # line stop (below)
    comp = pd.DataFrame({
        "Place": ["Waiting to Move to Line", "Input Rack WIP", "WIP Per Machine", "Endline & Bundling WIP"],
        "Standard": ["Issue for 4 hours max", "Max of 4 hrs of output capability of the line",
                     "5 Pcs x Number of machines", "Max of 1.5 hrs of output capability of the line"],
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
    combined = {"IN": j_in, "OUT": k_out, "Total": j_in + k_out,
                "IN days": j_in / per_day, "OUT days": k_out / per_day,
                "Total days": (j_in + k_out) / per_day}
    return comp, total, levels, combined


# --------------------------------------------------------------------------
# UI helpers
# --------------------------------------------------------------------------
def card(name, target, actual, days, lvl):
    """Band tile split by a vertical divider: Target (left, neutral) | Actual (right, level colour)."""
    _, _, bg, fg, _ = lvl
    return (
        "<div style='border:1px solid #444;border-radius:6px;overflow:hidden;text-align:center;"
        "line-height:1.25'>"
        f"<div style='background:#262730;color:#fff;font-weight:700;font-size:.8rem;padding:4px 2px;"
        f"border-bottom:1px solid #444'>{name}</div>"
        "<div style='display:flex'>"
        "<div style='flex:1;padding:6px 4px;background:#f0f2f6;color:#111;border-right:2px solid #444'>"
        f"<div style='font-size:.68rem;text-transform:uppercase;opacity:.7'>Target</div>"
        f"<div style='font-size:1rem;font-weight:700'>{target:,.0f}</div></div>"
        f"<div style='flex:1;padding:6px 4px;background:{bg};color:{fg}'>"
        f"<div style='font-size:.68rem;text-transform:uppercase;opacity:.85'>Actual</div>"
        f"<div style='font-size:1rem;font-weight:700'>{actual:,.0f}</div></div>"
        "</div>"
        f"<div style='background:{bg};color:{fg};font-size:.72rem;padding:2px;border-top:1px solid #444'>"
        f"{days:.2f} days</div></div>"
    )


def days_card(name, days, lvl):
    _, _, bg, fg, _ = lvl
    return (
        f"<div style='background:{bg};color:{fg};border:1px solid #444;border-radius:6px;"
        f"padding:6px 8px;text-align:center'>"
        f"<div style='font-weight:700;font-size:.8rem'>{name}</div>"
        f"<div style='font-size:1.25rem;font-weight:700'>{days:.2f}</div></div>"
    )


def grid(cards, per_row=9):
    for i in range(0, len(cards), per_row):
        cols = st.columns(per_row)
        for c, html in zip(cols, cards[i:i + per_row]):
            c.markdown(html, unsafe_allow_html=True)
        st.write("")


def style_level(df, th):
    colour_map = {l[1]: (l[2], l[3]) for l in LEVELS}

    def row_style(r):
        bg, fg = colour_map[r["Colour"]]
        return [f"background-color:{bg};color:{fg}" if c in ("WIP Days", "WIP Hours", "Colour", "Level")
                else "" for c in r.index]
    return df.style.apply(row_style, axis=1).format({
        "Day Pcs Target": "{:,.0f}", "Day Planned Hrs": "{:.0f}", "Hourly Output Target": "{:,.1f}",
        "System Line WIP": "{:,.0f}", "WIP Hours": "{:.2f}", "WIP Days": "{:.2f}",
        "Dashboard Target (pcs)": "{:,.0f}"})


def legend(th):
    parts = []
    bounds = [f"> {th[0]}", f"{th[1]} – {th[0]}", f"{th[2]} – {th[1]}", f"{th[3]} – {th[2]}", f"≤ {th[3]}"]
    for l, b in zip(LEVELS, bounds):
        parts.append(f"<span style='background:{l[2]};color:{l[3]};border:1px solid #444;border-radius:4px;"
                     f"padding:2px 8px;margin-right:6px;font-size:.8rem'>{l[0]} ({b} d)</span>")
    st.markdown("".join(parts), unsafe_allow_html=True)


def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="WIP LEVEL CALCULATIONS")
    return buf.getvalue()


# --------------------------------------------------------------------------
# Sidebar – data input
# --------------------------------------------------------------------------
st.title("Line Wise WIP Dashboard – R4")

with st.sidebar:
    st.header("Data input")
    f1 = st.file_uploader("Morning commitment (BAND, PLAN HOURS, DAY PLAN)", type=["xlsx", "csv"])
    f2 = st.file_uploader("System report (Workstation, WIP Qty)", type=["xlsx", "csv"])
    plan_df = pd.DataFrame(columns=["Zone", "Band No", "Plan Hours", "Day Plan"])
    wip_df = pd.DataFrame(columns=["Workstation", "WIP Qty"])
    try:
        if f1:
            plan_df = plan_from_df(read_any(f1, "commit"))
        if f2:
            wip_df = wip_from_df(read_any(f2, "report"))
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read the file: {e}")

    st.header("Colour thresholds (WIP days)")
    st.caption("Level applies when WIP days is strictly greater than the value.")
    th = [st.number_input(n, value=v, step=0.01, format="%.2f") for n, v in
          zip(["Red above", "Green above", "Yellow above", "Gray above (else Black)"], DEFAULT_THRESH)]
    if not all(th[i] > th[i + 1] for i in range(3)):
        st.warning("Thresholds should be in descending order.")

if not (f1 and f2):
    st.info("Upload the **Morning commitment** and **System report** files in the sidebar to build the dashboard.")
    st.stop()

with st.expander("Edit input tables (Morning commitment / System report)", expanded=False):
    c1, c2 = st.columns(2)
    plan_df = c1.data_editor(plan_df, num_rows="dynamic", width="stretch", key=f"plan_{f1.name}")
    wip_df = c2.data_editor(wip_df, num_rows="dynamic", width="stretch", key=f"wip_{f2.name}")
    c2.caption(f"Total system WIP: {pd.to_numeric(wip_df['WIP Qty'], errors='coerce').sum():,.0f} pcs")

plan_df = plan_df.dropna(subset=["Band No", "Day Plan"]).astype({"Band No": int})
wip_df = wip_df.dropna(subset=["Workstation", "WIP Qty"])
calc = calculate(plan_df, wip_df, th)

bad = unmatched_bands(plan_df, wip_df)
if bad:
    st.warning("System-report workstations with no morning commitment (ignored, as in Excel): "
               + ", ".join(bad))

# --------------------------------------------------------------------------
# KPIs
# --------------------------------------------------------------------------
if calc.empty:
    st.info("No morning-commitment rows found.")
    st.stop()

k = st.columns(6)
k[0].metric("Bands", len(calc))
k[1].metric("Total WIP (pcs)", f"{calc['System Line WIP'].sum():,.0f}")
k[2].metric("Avg WIP days", f"{calc['WIP Days'].mean():.2f}")
for col, name in zip(k[3:], ["High WIP", "Reorder Level", "Line Stop"]):
    col.metric(name, int((calc["Level"] == name).sum()))
legend(th)

tab1, tab2, tab3, tab4 = st.tabs(["WIP Dashboard (pcs)", "WIP Days Dashboard",
                                  "WIP Level Calculations", "WIP Standards"])
zones = list(dict.fromkeys(calc["Zone"]))

# ---- Tab 1 : pcs dashboard ------------------------------------------------
with tab1:
    zsel = st.multiselect("Zone filter", zones, default=zones, key="z1")
    for z in zsel:
        d = calc[calc["Zone"] == z]
        st.subheader(z or "Zone")
        grid([card(r["Band"], r["Dashboard Target (pcs)"], r["System Line WIP"], r["WIP Days"],
                   level_of(r["WIP Days"], th)) for _, r in d.iterrows()])
    st.caption("Target = Day pcs target × 1.18 (max standard WIP); Actual = system line WIP. "
               "Colour is driven by WIP days.")
    chart = calc.set_index("Band")[["Dashboard Target (pcs)", "System Line WIP"]]
    st.bar_chart(chart, height=320)

# ---- Tab 2 : days dashboard ------------------------------------------------
with tab2:
    zsel2 = st.multiselect("Zone filter", zones, default=zones, key="z2")
    for z in zsel2:
        d = calc[calc["Zone"] == z]
        st.subheader(z or "Zone")
        grid([days_card(r["Band"], r["WIP Days"], level_of(r["WIP Days"], th)) for _, r in d.iterrows()])

# ---- Tab 3 : calculations --------------------------------------------------
with tab3:
    show = calc.drop(columns=["Dashboard Target (pcs)"])
    st.dataframe(style_level(show, th), width="stretch", hide_index=True, height=560)
    st.download_button("Download calculations (.xlsx)", to_excel(calc), "WIP_LEVEL_CALCULATIONS.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.markdown("**Level colour coding & directions to store**")
    st.dataframe(pd.DataFrame({
        "Level": [l[0] for l in LEVELS], "Colour": [l[1] for l in LEVELS],
        "From (days)": [f"{th[0]}+", f"{th[1]}+", f"{th[2]}+", f"{th[3]}+", "0"],
        "To (days)": ["Above", str(th[0]), str(th[1]), str(th[2]), str(th[3])],
        "Direction to store": [l[4] for l in LEVELS]}), hide_index=True, width="stretch")

# ---- Tab 4 : standards -----------------------------------------------------
with tab4:
    c = st.columns(3)
    hourly = c[0].number_input("Hourly line target (pcs)", value=110.0, step=1.0)
    machines = c[1].number_input("Number of machines", value=51.0, step=1.0)
    run_hrs = c[2].number_input("Running time (hours)", value=10.0, step=0.5)
    extra = st.number_input("Extra pcs added to combined IN standard (Excel: +220)", value=220.0)
    comp, total, levels, comb = standards(hourly, machines, run_hrs, extra)

    a, b = st.columns(2)
    with a:
        st.markdown("**Standard WIP build-up**")
        st.dataframe(comp.style.format({"Calculated Qty": "{:,.1f}"}), hide_index=True,
                     width="stretch")
        st.metric("Total WIP standard (pcs)", f"{total:,}")
        st.metric("Standard WIP in days", f"{levels.loc[1, 'To (days)']:.2f}")
    with b:
        st.markdown("**Levels (pcs and days)**")
        st.dataframe(levels.style.format({c_: "{:,.1f}" for c_ in levels.columns if "(" in c_},
                                         na_rep="–"), hide_index=True, width="stretch")
    st.markdown("**Combined lines (IN + OUT)**")
    st.dataframe(pd.DataFrame([comb]).style.format("{:,.2f}"), hide_index=True, width="stretch")
