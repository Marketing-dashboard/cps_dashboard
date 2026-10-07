"""
CPS Demand & Supply Dashboard Generator
========================================
Run this script whenever you have new data:
    python generate_dashboard.py

It reads:
  - CPS_Demand_supply.xlsx   (demand sheet + Supply sheet)
  - CPS_Apr_may_data.xlsx    (historical data — any number of months)

And writes:
  - index.html               (the live dashboard)

AUTO-DETECTION:
  - Current month + year are read from the Supply sheet dates automatically
  - Historical months are all months found in CPS_Apr_may_data.xlsx automatically
  - Working days per month = 26 (configurable below)
  - Supply days = however many unique dates exist in the Supply sheet
"""

import pandas as pd
import json
import re
from pathlib import Path
from datetime import datetime

# ── CONFIG ──────────────────────────────────────────────────────────
DEMAND_FILE   = "CPS_Demand_supply.xlsx"
HIST_FILE     = "CPS_Apr_may_data.xlsx"
OUTPUT_FILE   = "index.html"
WORKING_DAYS  = 25   # working days assumed per month (for both demand and historical avg)

# Chetak variant names in the historical file — all get summed for combined Bajaj Chetak demand row
CHETAK_HIST_KEYS = [
    "Bajaj Chetak", "Bajaj Chetak 3503", "Bajaj Chetak 3001",
    "Bajaj Chetak C2501", "Bajaj Chetak C3001", "Bajaj Chetak C3503",
]

# ── MODEL MAPPINGS ───────────────────────────────────────────────────
# Key   = model name as it appears in the demand sheet
# Value = "Actual Model" name in the Supply sheet  (None = no supply data)
SUP_MAP = {
    "Apache RTR 310":                    "TVS Apache RR 310",
    "Apache RTR 160":                    "TVS Apache RTR 160",
    "Apache RTR 180":                    "TVS Apache RTR 180",
    "Apache RTR 200 4V":                 "TVS Apache RTR 200 4V",
    "Ronin":                             "TVS Ronin",
    "Jupiter 110":                       "TVS Jupiter",
    "Jupiter 125":                       "TVS Jupiter 125",
    "Ntorq 125":                         "TVS NTORQ 125",
    "Scooty Zest 110":                   "TVS Scooty Zest",
    "Radeon":                            "TVS Radeon",
    "Raider 125":                        "TVS Raider",
    "Sport":                             "TVS Sport",
    "XL100":                             "TVS XL100",
    "iQube":                             "TVS iQube",
    "Bajaj Pulsar 125":                  "Bajaj Pulsar 125",
    "Bajaj Pulsar 150":                  "Bajaj Pulsar 150",
    "Bajaj Pulsar NS 125":               "Bajaj Pulsar NS 125",
    "Bajaj Pulsar NS200":                "Bajaj Pulsar NS200",
    "Bajaj Pulsar N160":                 "Bajaj Pulsar N160",
    "Bajaj Pulsar 220 F":                "Bajaj Pulsar 220 F",
    "Bajaj Pulsar NS160":                "Bajaj Pulsar NS160",
    "Bajaj Pulsar NS400Z":               "Bajaj Pulsar NS400Z",
    "Bajaj Pulsar 180":                  "Bajaj Pulsar 180",
    "KTM 200 Duke":                      "KTM 200 Duke",
    "KTM 250 Duke":                      "KTM 250 Duke",
    "KTM Duke 390":                      "KTM Duke 390",
    "KTM 160 Duke":                      "KTM 160 Duke",
    "Triumph Speed 400":                 "Triumph Speed 400",
    "Triumph Speed T4":                  "Triumph Speed T4",
    "Triumph Tracker 400":               "Triumph Tracker 400",
    "Bajaj Chetak":                      "Bajaj Chetak",
    "Bajaj Maxima Xl Cargo E-TEC 12.0":  "Bajaj Maxima XL Cargo E-TEC 12.0",
    "Bajaj Wego P9018":                  "Bajaj WEGO P9018",
    "Bajaj Wego P70":                    "Bajaj WEGO P70",
    "Bajaj Wego P50":                    None,
    "Bajaj Wego C90":                    "Bajaj WEGO C90",
    "Ather 450 X":                       "Ather 450X",
    "Ather Rizta":                       "Ather Rizta",
    "Jawa 42FJ":                         "Jawa 42 FJ",
    "Jawa 42":                           "Jawa 42",
    "Jawa 42 Bobber":                    "Jawa 42 Bobber",
    "Yezdi Adventure":                   "Yezdi Adventure",
    "Yezdi Scrambler":                   None,
    "Yezdi Roadster":                    "Yezdi Roadster",
    "BSA Scrambler":                     None,
    "Magnus Neo":                        "Ampere Magnus Neo",
    "New Magnus Neo":                    None,
    "Nexus":                             "Ampere Nexus",
    "Magnus G Max":                      "Ampere Magnus G Max",
    "Magnus Grand":                      "Ampere Magnus Grand",
}

# Key   = demand sheet model name
# Value = "Model Projection" in historical file
#         None            = no historical data
#         "__CHETAK__"    = special: sum all CHETAK_HIST_KEYS variants
HIST_MAP = {
    "Apache RTR 310":                    "TVS Apache RTR 310",
    "Apache RTR 160":                    "TVS Apache RTR 160",
    "Apache RTR 180":                    "TVS Apache RTR 180",
    "Apache RTR 200 4V":                 "TVS Apache RTR 200 4V",
    "Ronin":                             "TVS Ronin",
    "Jupiter 110":                       "TVS Jupiter",
    "Jupiter 125":                       "TVS Jupiter 125",
    "Ntorq 125":                         "TVS NTORQ 125",
    "Scooty Zest 110":                   "TVS Scooty Zest",
    "Radeon":                            "TVS Radeon",
    "Raider 125":                        "TVS Raider",
    "Sport":                             "TVS Sport",
    "XL100":                             "TVS XL100",
    "iQube":                             "TVS iQube",
    "Bajaj Pulsar 125":                  "Bajaj Pulsar 125",
    "Bajaj Pulsar 150":                  "Bajaj Pulsar 150",
    "Bajaj Pulsar NS 125":               "Bajaj Pulsar NS 125",
    "Bajaj Pulsar NS200":                "Bajaj Pulsar NS200",
    "Bajaj Pulsar N160":                 "Bajaj Pulsar N160",
    "Bajaj Pulsar 220 F":                "Bajaj Pulsar 220 F",
    "Bajaj Pulsar NS160":                "Bajaj Pulsar NS160",
    "Bajaj Pulsar NS400Z":               "Bajaj Pulsar NS400Z",
    "Bajaj Pulsar 180":                  "Bajaj Pulsar 180",
    "KTM 200 Duke":                      "KTM 200 Duke",
    "KTM 250 Duke":                      "KTM 250 Duke",
    "KTM Duke 390":                      "KTM Duke 390",
    "KTM 160 Duke":                      "KTM 160 Duke",
    "Triumph Speed 400":                 "Triumph Speed 400",
    "Triumph Speed T4":                  "Triumph Speed T4",
    "Triumph Tracker 400":               "Triumph Tracker 400",
    "Bajaj Chetak":                      "__CHETAK__",
    "Bajaj Maxima Xl Cargo E-TEC 12.0":  "Bajaj Maxima XL Cargo E-TEC 12.0",
    "Bajaj Wego P9018":                  None,
    "Bajaj Wego P70":                    None,
    "Bajaj Wego P50":                    None,
    "Bajaj Wego C90":                    None,
    "Ather 450 X":                       "Ather 450X",
    "Ather Rizta":                       "Ather Rizta",
    "Jawa 42FJ":                         "Jawa 42 FJ",
    "Jawa 42":                           "Jawa 42",
    "Jawa 42 Bobber":                    "Jawa 42 Bobber",
    "Yezdi Adventure":                   "Yezdi Adventure",
    "Yezdi Scrambler":                   "Yezdi Scrambler",
    "Yezdi Roadster":                    "Yezdi Roadster",
    "BSA Scrambler":                     None,
    "Magnus Neo":                        "Ampere Magnus Neo",
    "New Magnus Neo":                    "New Ampere Magnus Neo",
    "Nexus":                             "Ampere Nexus",
    "Magnus G Max":                      "Ampere Magnus G Max",
    "Magnus Grand":                      "Ampere Magnus Grand",
}


# ── HELPERS ─────────────────────────────────────────────────────────

def get_brand(m):
    if m.startswith(("Apache","Jupiter","Ntorq","Scooty","Radeon","Raider","Sport","XL100","iQube","Ronin")):
        return "TVS"
    if m.startswith("KTM"):      return "KTM"
    if m.startswith("Triumph"):  return "Triumph"
    if m.startswith("Ather"):    return "Ather"
    if m.startswith(("Jawa","Yezdi","BSA")): return "Jawa/Yezdi/BSA"
    if m.startswith(("Magnus","New Magnus","Nexus","Reo")): return "Ampere"
    return "Bajaj"


def parse_month_label(val):
    """
    Convert month strings like "Mar'26", "Apr'26", "Jun'26" etc.
    into a sortable (year, month_num) tuple and a display label.
    Returns (sort_key, display_label) or None if unparseable.
    """
    MONTH_NUM = {
        "jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
        "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12
    }
    s = str(val).strip()
    # Pattern: "Mar'26" or "Mar26" or "March 2026"
    m = re.match(r"([A-Za-z]+)['\s]?(\d{2,4})", s)
    if not m:
        return None
    mon_str = m.group(1).lower()[:3]
    yr_str  = m.group(2)
    yr = int(yr_str) + 2000 if len(yr_str) == 2 else int(yr_str)
    mon_num = MONTH_NUM.get(mon_str)
    if not mon_num:
        return None
    return (yr * 100 + mon_num, s)   # (sort_key, original_label)


def detect_current_month(sup_df):
    """
    Read the Supply sheet dates (datetime column) and return a display string
    like "June 2026" plus the month/year integers.
    """
    dates = pd.to_datetime(sup_df["Date"], errors="coerce").dropna()
    if dates.empty:
        raise ValueError("No valid dates found in Supply sheet.")
    latest = dates.max()
    return latest.strftime("%B %Y"), latest.month, latest.year


def detect_hist_months(hist_df):
    """
    Find every unique month label in the historical file's Date column,
    parse and sort them, return:
      - sorted list of original labels  (used for filtering rows)
      - human-readable list             (used in dashboard header)
      - total_days = count * WORKING_DAYS
    """
    raw_vals = hist_df["Date"].dropna().unique()
    parsed = []
    for v in raw_vals:
        result = parse_month_label(v)
        if result:
            parsed.append(result)

    if not parsed:
        raise ValueError("Could not parse any month labels from historical file Date column.")

    parsed.sort(key=lambda x: x[0])
    original_labels = [p[1] for p in parsed]
    total_days      = len(original_labels) * WORKING_DAYS

    # Build a nice display like "Mar'26 + Apr'26 + May'26"
    display = " + ".join(original_labels)
    return original_labels, display, total_days


# ── MAIN DATA LOADER ─────────────────────────────────────────────────

def load_data():
    # ── DEMAND ──────────────────────────────────────────────────────
    # Support both "demand" (old) and "Demand" (new capitalized sheet name)
    sheet_names = pd.ExcelFile(DEMAND_FILE).sheet_names
    dem_sheet = "Demand" if "Demand" in sheet_names else "demand"
    dem = pd.read_excel(DEMAND_FILE, sheet_name=dem_sheet)
    if dem.shape[1] == 2:
        dem.columns = ["model", "Total"]
        dem["Google"] = 0
        dem["Facebook"] = 0
    else:
        dem.columns = ["model", "Google", "Facebook", "Total"]
    for c in ["Google", "Facebook", "Total"]:
        dem[c] = pd.to_numeric(dem[c], errors="coerce").fillna(0)
    dem_nz = dem[dem["Total"] > 0].copy()
    dem_nz["daily_demand"] = (dem_nz["Total"] / WORKING_DAYS).round(1)

    # ── SUPPLY ──────────────────────────────────────────────────────
    # Support both old "Supply" sheet and new "Raw" sheet format
    if "Supply" in sheet_names:
        sup = pd.read_excel(DEMAND_FILE, sheet_name="Supply")
        if sup.shape[1] == 5 and sup.iloc[0, 1] != "Medium":
            sup.columns = ["Date", "Actual_Model", "Medium", "Total_Leads", "Process"]
        else:
            sup.columns = ["Date", "Medium", "Total_Leads", "Process", "Actual_Model"]
    else:
        # New format: Raw sheet with named columns
        sup_raw = pd.read_excel(DEMAND_FILE, sheet_name="Raw")
        sup = pd.DataFrame({
            "Date":         sup_raw["Date"],
            "Actual_Model": sup_raw["Model_Map"],
            "Medium":       sup_raw["Medium"],
            "Total_Leads":  sup_raw["Total"],
            "Process":      sup_raw["Process"],
        })
    sup["Total_Leads"] = pd.to_numeric(sup["Total_Leads"], errors="coerce").fillna(0)

    # Auto-detect: current month name + how many days of supply data exist
    current_month_label, cur_month, cur_year = detect_current_month(sup)
    supply_days = sup["Date"].nunique()   # count unique dates present

    print(f"  Current month detected : {current_month_label}")
    print(f"  Supply data days found : {supply_days}")

    sup_tot = sup.groupby("Actual_Model")["Total_Leads"].sum()
    sup_fb  = sup[sup["Medium"] == "Facebook"].groupby("Actual_Model")["Total_Leads"].sum()
    sup_g   = sup[sup["Medium"] == "Google"].groupby("Actual_Model")["Total_Leads"].sum()

    def gsup(key, series):
        if not key or key not in series.index:
            return 0
        return round(float(series[key]) / supply_days, 1)

    # ── HISTORICAL ──────────────────────────────────────────────────
    hist_df = pd.read_excel(HIST_FILE)
    hist_df["Total_Leads"] = pd.to_numeric(hist_df["Total_Leads"], errors="coerce").fillna(0)

    # Auto-detect all months present in the file
    hist_months, hist_display, hist_total_days = detect_hist_months(hist_df)
    num_hist_months = len(hist_months)

    print(f"  Historical months found: {hist_display}  ({num_hist_months} months × {WORKING_DAYS} days = {hist_total_days} days)")

    am3 = hist_df[hist_df["Date"].isin(hist_months)]

    hist_tot = am3.groupby("Model Projection")["Total_Leads"].sum()
    hist_fb  = am3[am3["Medium"] == "Facebook"].groupby("Model Projection")["Total_Leads"].sum()
    hist_g   = am3[am3["Medium"] == "Google"].groupby("Model Projection")["Total_Leads"].sum()

    # Chetak combined historical
    chetak_mask = am3["Model Projection"].isin(CHETAK_HIST_KEYS)
    chetak_tot  = am3[chetak_mask]["Total_Leads"].sum()
    chetak_fb   = am3[chetak_mask & (am3["Medium"] == "Facebook")]["Total_Leads"].sum()
    chetak_g    = am3[chetak_mask & (am3["Medium"] == "Google")]["Total_Leads"].sum()

    def ghist(key, series):
        if not key or key == "__CHETAK__":
            return 0
        if key not in series.index:
            return 0
        return round(float(series[key]) / hist_total_days, 1)

    # ── BUILD MODEL ROWS ────────────────────────────────────────────
    rows = []
    for _, dr in dem_nz.iterrows():
        m  = dr["model"]
        sk = SUP_MAP.get(m)
        hk = HIST_MAP.get(m)

        ds  = gsup(sk, sup_tot)
        dsf = gsup(sk, sup_fb)
        dsg = gsup(sk, sup_g)

        if hk == "__CHETAK__":
            dh  = round(chetak_tot / hist_total_days, 1)
            dhf = round(chetak_fb  / hist_total_days, 1)
            dhg = round(chetak_g   / hist_total_days, 1)
        else:
            dh  = ghist(hk, hist_tot)
            dhf = ghist(hk, hist_fb)
            dhg = ghist(hk, hist_g)

        pct = round((ds / dh * 100) if dh > 0 else 0, 1)

        rows.append({
            "model": m,  "brand": get_brand(m),
            "monthly_demand":      int(dr["Total"]),
            "daily_demand":        dr["daily_demand"],
            "daily_supply":        ds,  "daily_supply_fb": dsf,  "daily_supply_google": dsg,
            "daily_hist":          dh,  "daily_hist_fb":   dhf,  "daily_hist_google":   dhg,
            "pct_hist": pct,
        })

    # ── BRAND AGGREGATION ───────────────────────────────────────────
    brands = {}
    for r in rows:
        b = r["brand"]
        if b not in brands:
            brands[b] = {k: 0 for k in [
                "daily_demand", "daily_supply", "daily_supply_fb", "daily_supply_google",
                "monthly_demand", "daily_hist", "daily_hist_fb", "daily_hist_google", "model_count"
            ]}
            brands[b]["brand"] = b
        for k in ["daily_demand", "daily_supply", "daily_supply_fb", "daily_supply_google",
                  "monthly_demand", "daily_hist", "daily_hist_fb", "daily_hist_google"]:
            brands[b][k] += r[k]
        brands[b]["model_count"] += 1

    brand_rows = []
    for b, br in brands.items():
        ds = round(br["daily_supply"], 1)
        dh = round(br["daily_hist"],   1)
        brand_rows.append({
            "brand": b, "monthly_demand": br["monthly_demand"],
            "daily_demand":        round(br["daily_demand"],        1),
            "daily_supply":        ds,
            "daily_supply_fb":     round(br["daily_supply_fb"],     1),
            "daily_supply_google": round(br["daily_supply_google"], 1),
            "daily_hist":          dh,
            "daily_hist_fb":       round(br["daily_hist_fb"],       1),
            "daily_hist_google":   round(br["daily_hist_google"],   1),
            "pct_hist": round((ds / dh * 100) if dh > 0 else 0, 1),
            "model_count": br["model_count"],
        })
    brand_rows.sort(key=lambda x: -x["monthly_demand"])

    meta = {
        "current_month":   current_month_label,
        "supply_days":     supply_days,
        "hist_display":    hist_display,
        "hist_months":     num_hist_months,
        "hist_total_days": hist_total_days,
    }

    return rows, brand_rows, meta


# ── HTML BUILDER ─────────────────────────────────────────────────────

def build_html(rows, brand_rows, meta):
    import os

    over     = len([r for r in rows if r["pct_hist"] > 100])
    on_track = len([r for r in rows if 90 <= r["pct_hist"] <= 100])
    short    = len([r for r in rows if r["pct_hist"] < 90])

    rows_json  = json.dumps(rows)
    brand_json = json.dumps(brand_rows)

    cm  = meta["current_month"]          # e.g. "June 2026"
    sd  = meta["supply_days"]            # e.g. 9
    hd  = meta["hist_display"]           # e.g. "Mar'26 + Apr'26 + May'26"
    htd = meta["hist_total_days"]        # e.g. 78
    hm  = meta["hist_months"]            # e.g. 3

    # Build dynamic label strings
    dem_label    = f"{cm} Demand (÷{WORKING_DAYS})"
    sup_label    = f"{cm} Supply — {sd}-day avg"
    his_label    = f"Historical Avg — {hd} (÷{htd})"
    his_subtitle = f"{hd} baseline"
    topbar_meta  = (
        f"Demand ÷{WORKING_DAYS} &nbsp;·&nbsp; "
        f"Supply ÷{sd} days &nbsp;·&nbsp; "
        f"Hist = {hd} ÷{htd}"
    )
    data_range   = f"Data: {cm[:3]} 1–{sd}, {cm[-4:]}"

    # Load the HTML template (same directory as this script)
    template_path = Path(__file__).parent / "dashboard_template.html"
    template = template_path.read_text(encoding="utf-8")

    # Inject all dynamic values
    html = (template
        .replace("__BRAND_JSON__",    brand_json)
        .replace("__MODEL_JSON__",    rows_json)
        .replace("__CURRENT_MONTH__", cm)
        .replace("__DATA_RANGE__",    data_range)
        .replace("__SUPPLY_DAYS__",   str(sd))
        .replace("__DEM_LABEL__",     dem_label)
        .replace("__SUP_LABEL__",     sup_label)
        .replace("__HIS_LABEL__",     his_label)
        .replace("__HIS_SUBTITLE__",  his_subtitle)
        .replace("__TOPBAR_META__",   topbar_meta)
        # update model count badge
        .replace(">50 models<",       f">{len(rows)} models<")
        .replace(">50 model<",        f">{len(rows)} model<")
    )

    return html


# ── ENTRY POINT ──────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Loading data...")
    rows, brand_rows, meta = load_data()
    print(f"  {len(rows)} models across {len(brand_rows)} brands")

    print("Building dashboard...")
    html = build_html(rows, brand_rows, meta)
    Path(OUTPUT_FILE).write_text(html, encoding="utf-8")

    print(f"  OK  Written to: {OUTPUT_FILE}")
    print(f"  OK  Title: CPS Demand & Supply Dashboard - {meta['current_month']}")
    print("Done! Push to GitHub to update the live site.")
