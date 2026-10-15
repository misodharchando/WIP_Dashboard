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
TILE_CSS = """
<style>
.wt-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(var(--min,170px),1fr));gap:12px;margin:4px 0 18px}
.wt-legend{display:flex;flex-wrap:wrap;gap:6px;margin:4px 0 10px}
.wt-chip{border:1px solid #444;border-radius:6px;padding:2px 8px;font-size:.8rem;white-space:nowrap}
@media (max-width:640px){
 .wt-grid{grid-template-columns:repeat(auto-fill,minmax(calc(var(--min,170px) * .85),1fr));gap:8px}
 .wt-act{font-size:1.2rem}.wt-big{font-size:1.5rem}.wt-name{font-size:.7rem}.wt-badge{font-size:.55rem;padding:2px 6px}
 .wt-foot{font-size:.62rem}
}
.wt{background:#fff;border:1px solid #e3e6ee;border-radius:12px;box-shadow:0 1px 3px rgba(20,25,40,.10);
    overflow:hidden;font-family:'Segoe UI',system-ui,sans-serif;color:#1c2333;transition:transform .12s,box-shadow .12s}
.wt:hover{transform:translateY(-2px);box-shadow:0 6px 14px rgba(20,25,40,.18)}
.wt-bar{height:7px}
.wt-top{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:3px;padding:8px 10px 2px}
.wt-name,.wt-badge{white-space:nowrap}
.wt-name{font-weight:800;font-size:.78rem;letter-spacing:.04em;color:#44506b}
.wt-badge{font-size:.62rem;font-weight:700;padding:2px 8px;border-radius:999px;text-transform:uppercase;letter-spacing:.03em}
.wt-body{display:flex;align-items:stretch;padding:6px 10px 4px}
.wt-col{flex:1;text-align:center}
.wt-div{width:0;border-left:2px dashed #cfd5e3;margin:2px 6px}
.wt-lbl{font-size:.6rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#8a93a8}
.wt-tgt{font-size:1rem;font-weight:600;color:#6b7590}
.wt-act{font-size:1.45rem;font-weight:800;line-height:1.1}
.wt-track{position:relative;height:8px;background:#e9ecf3;border-radius:6px;margin:6px 10px 4px;overflow:visible}
.wt-fill{height:100%;border-radius:6px}
.wt-mark{position:absolute;top:-3px;width:2px;height:14px;background:#1c2333;opacity:.55}
.wt-foot{display:flex;justify-content:space-between;padding:2px 10px 9px;font-size:.7rem;color:#6b7590}
.wt-foot b{color:#1c2333}
.wt-big{font-size:1.9rem;font-weight:800;text-align:center;padding:4px 0 2px;line-height:1.1}
.wt-sub{text-align:center;font-size:.64rem;color:#8a93a8;padding-bottom:9px;letter-spacing:.06em;text-transform:uppercase}
</style>
"""


def _badge(lvl):
    name, _, bg, fg, _ = lvl
    bd = "border:1px solid #bbb;" if bg.upper() in ("#FFFF00",) else ""
    return f"<span class='wt-badge' style='background:{bg};color:{fg};{bd}'>{name}</span>"


def card(name, target, actual, days, lvl):
    """Band tile: level accent bar + badge, Target | Actual split by a dashed divider, progress bar."""
    bg = lvl[2]
    pct = (actual / target) if target else 0
    fill = min(pct, 1.5) / 1.5 * 100            # bar spans 0-150 % of target; marker at 100 %
    mark = 100 / 1.5
    return (
        "<div class='wt'>"
        f"<div class='wt-bar' style='background:{bg}'></div>"
        f"<div class='wt-top'><span class='wt-name'>{name.replace('R4-', '')}</span>{_badge(lvl)}</div>"
        "<div class='wt-body'>"
        f"<div class='wt-col'><div class='wt-lbl'>Target</div><div class='wt-tgt'>{target:,.0f}</div></div>"
        "<div class='wt-div'></div>"
        f"<div class='wt-col'><div class='wt-lbl'>Actual</div><div class='wt-act'>{actual:,.0f}</div></div>"
        "</div>"
        f"<div class='wt-track'><div class='wt-fill' style='width:{fill:.1f}%;background:{bg}'></div>"
        f"<div class='wt-mark' style='left:{mark:.1f}%'></div></div>"
        f"<div class='wt-foot'><span><b>{days:.2f}</b> days</span><span><b>{pct:.0%}</b> of target</span></div>"
        "</div>"
    )


def days_card(name, days, lvl):
    bg = lvl[2]
    return (
        "<div class='wt'>"
        f"<div class='wt-bar' style='background:{bg}'></div>"
        f"<div class='wt-top'><span class='wt-name'>{name.replace('R4-', '')}</span>{_badge(lvl)}</div>"
        f"<div class='wt-big'>{days:.2f}</div><div class='wt-sub'>WIP days</div>"
        "</div>"
    )


def grid(cards, min_px=170):
    """Responsive CSS grid: as many tiles per row as fit (wraps to 2-up / 1-up on phones)."""
    st.markdown(f"<div class='wt-grid' style='--min:{min_px}px'>{''.join(cards)}</div>",
                unsafe_allow_html=True)


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
    chips = []
    bounds = [f"> {th[0]}", f"{th[1]} – {th[0]}", f"{th[2]} – {th[1]}", f"{th[3]} – {th[2]}", f"≤ {th[3]}"]
    for l, b in zip(LEVELS, bounds):
        chips.append(f"<span class='wt-chip' style='background:{l[2]};color:{l[3]}'>{l[0]} ({b} d)</span>")
    st.markdown(f"<div class='wt-legend'>{''.join(chips)}</div>", unsafe_allow_html=True)


def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="WIP LEVEL CALCULATIONS")
    return buf.getvalue()


# --------------------------------------------------------------------------
# Sidebar – data input
# --------------------------------------------------------------------------
st.title("Line Wise WIP Dashboard – R4")
st.markdown(TILE_CSS, unsafe_allow_html=True)

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

kpis = [("Bands", f"{len(calc)}"), ("Total WIP (pcs)", f"{calc['System Line WIP'].sum():,.0f}"),
        ("Avg WIP days", f"{calc['WIP Days'].mean():.2f}")] + \
       [(n, f"{int((calc['Level'] == n).sum())}") for n in ["High WIP", "Reorder Level", "Line Stop"]]
st.markdown("<div class='wt-grid' style='--min:115px'>" + "".join(
    f"<div class='wt' style='padding:10px 12px'><div class='wt-lbl'>{n}</div>"
    f"<div style='font-size:1.5rem;font-weight:800'>{v}</div></div>" for n, v in kpis) + "</div>",
    unsafe_allow_html=True)
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
                   level_of(r["WIP Days"], th)) for _, r in d.iterrows()], min_px=170)
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
        grid([days_card(r["Band"], r["WIP Days"], level_of(r["WIP Days"], th)) for _, r in d.iterrows()],
             min_px=120)

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
