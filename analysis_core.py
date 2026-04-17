import pandas as pd
import numpy as np
from datetime import date, datetime
from dateutil.parser import parse

def default_rent_roll_df():
    # lease_end can be a date string or "MTM"
    data = [
        {"suite": "100", "tenant": "Alpha Dental", "sf": 2200, "tenant_since": "2019-04-01",
         "lease_start": "2022-01-01", "lease_end": "2027-12-31", "base_monthly_rent": 7700, "monthly_cams": 1100, "options": "1x5yr"},
        {"suite": "110", "tenant": "Vacant", "sf": 1500, "tenant_since": "",
         "lease_start": "", "lease_end": "", "base_monthly_rent": 0, "monthly_cams": 0, "options": ""},
        {"suite": "120", "tenant": "Bright Logistics", "sf": 4100, "tenant_since": "2021-08-15",
         "lease_start": "2021-09-01", "lease_end": "MTM", "base_monthly_rent": 10250, "monthly_cams": 1850, "options": "MTM"},
        {"suite": "200", "tenant": "Cedar Financial", "sf": 3200, "tenant_since": "2018-06-01",
         "lease_start": "2023-03-01", "lease_end": "2026-02-28", "base_monthly_rent": 9600, "monthly_cams": 1400, "options": "2x3yr"},
        {"suite": "210", "tenant": "Nori Sushi", "sf": 1800, "tenant_since": "2024-05-01",
         "lease_start": "2024-05-01", "lease_end": "2029-04-30", "base_monthly_rent": 7200, "monthly_cams": 900, "options": "1x5yr"},
    ]
    return pd.DataFrame(data)

def default_stress_test_state():
    return {
        # Loan terms
        "loan_source": "New Loan",          # New Loan / Loan Review (shared for Loan Amt + Note Rate)
        "loan_amount": 7500000.0,
        "note_rate": 0.075,                # decimal
        "io": False,
        "amort_years": 25.0,               # used if io=False
        "debt_service_override_on": False,
        "annual_debt_service_override": 0.0,

        # Income / NOI / Value
        "income_source": "Rent Roll",
        "rental_income": 720000.0,         # annual
        "opex_source": "Tax Return",
        "operating_expenses": 240000.0,    # annual
        "cap_source": "Appraisal",
        "cap_rate": 0.065,                 # decimal

        # Breakeven thresholds (POC defaults)
        "target_dscr": 1.20,
        "target_ltv": 0.75,
    }

def _to_date_maybe(x):
    if x is None:
        return None
    if isinstance(x, (datetime, date)):
        return x if isinstance(x, date) else x.date()
    s = str(x).strip()
    if s == "" or s.lower() in ["nan", "none"]:
        return None
    if s.strip().upper() == "MTM":
        return "MTM"
    try:
        return parse(s).date()
    except Exception:
        return None

def years_between(d1, d2):
    if not d1 or not d2 or d1 == "MTM" or d2 == "MTM":
        return np.nan
    days = (d2 - d1).days
    return round(days / 365.25, 2)

def remaining_term_years(as_of, lease_end):
    if lease_end == "MTM" or lease_end is None:
        return 0.0
    if not isinstance(as_of, date) or not isinstance(lease_end, date):
        return 0.0
    days = (lease_end - as_of).days
    return round(max(days / 365.25, 0.0), 2)

def compute_rent_roll(df: pd.DataFrame, as_of: date) -> pd.DataFrame:
    out = df.copy()

    # Ensure required cols exist
    for col in ["suite","tenant","sf","tenant_since","lease_start","lease_end","base_monthly_rent","monthly_cams","options"]:
        if col not in out.columns:
            out[col] = ""

    # Coerce numeric
    out["sf"] = pd.to_numeric(out["sf"], errors="coerce").fillna(0.0)
    out["base_monthly_rent"] = pd.to_numeric(out["base_monthly_rent"], errors="coerce").fillna(0.0)
    out["monthly_cams"] = pd.to_numeric(out["monthly_cams"], errors="coerce").fillna(0.0)

    # Parse lease dates
    lease_start_parsed = out["lease_start"].apply(_to_date_maybe)
    lease_end_parsed = out["lease_end"].apply(_to_date_maybe)

    total_sf = float(out["sf"].sum()) if out["sf"].sum() > 0 else 0.0
    out["tenant_pct_total_sf"] = np.where(total_sf > 0, out["sf"] / total_sf, 0.0)

    # Original lease term (years)
    out["original_lease_term_years"] = [
        years_between(ls if isinstance(ls, date) else None, le if isinstance(le, date) else None)
        if (le != "MTM") else np.nan
        for ls, le in zip(lease_start_parsed, lease_end_parsed)
    ]

    # Remaining term (years) as-of rent roll date
    out["remaining_term_years"] = [
        remaining_term_years(as_of, le if isinstance(le, date) else (None if le != "MTM" else "MTM"))
        for le in lease_end_parsed
    ]

    # Monthly rent / SF (base only)
    out["monthly_rent_per_sf"] = np.where(out["sf"] > 0, out["base_monthly_rent"] / out["sf"], np.nan)

    # Annual rent (base + CAMs)
    out["annual_rent"] = (out["base_monthly_rent"] + out["monthly_cams"]) * 12.0
    total_annual_rent = float(out["annual_rent"].sum()) if out["annual_rent"].sum() > 0 else 0.0
    out["pct_total_rent"] = np.where(total_annual_rent > 0, out["annual_rent"] / total_annual_rent, 0.0)

    # Display-friendly formatting columns (keep as numeric internally)
    return out

def totals_row(df: pd.DataFrame) -> pd.DataFrame:
    # Build a "TOTAL / AVG" row consistent with your requirements
    tr = {c: "" for c in df.columns}
    tr["suite"] = "TOTAL / AVG"
    tr["tenant"] = ""

    # Totals
    tr["sf"] = df["sf"].sum()
    tr["tenant_pct_total_sf"] = df["tenant_pct_total_sf"].sum()
    tr["base_monthly_rent"] = df["base_monthly_rent"].sum()
    tr["monthly_cams"] = df["monthly_cams"].sum()
    tr["annual_rent"] = df["annual_rent"].sum()
    tr["pct_total_rent"] = df["pct_total_rent"].sum()

    # Averages
    tr["original_lease_term_years"] = df["original_lease_term_years"].mean(skipna=True)
    tr["remaining_term_years"] = df["remaining_term_years"].mean(skipna=True)
    tr["monthly_rent_per_sf"] = df["monthly_rent_per_sf"].mean(skipna=True)

    return pd.concat([df, pd.DataFrame([tr])], ignore_index=True)

def lease_expiration_summary(df: pd.DataFrame) -> pd.DataFrame:
    # Buckets: Year 1 (<=1), Year 2 (1-2], Year 3 (2-3]
    buckets = {
        "Year 1": (df["remaining_term_years"] <= 1.0),
        "Year 2": ((df["remaining_term_years"] > 1.0) & (df["remaining_term_years"] <= 2.0)),
        "Year 3": ((df["remaining_term_years"] > 2.0) & (df["remaining_term_years"] <= 3.0)),
    }

    total_sf = df["sf"].sum() if df["sf"].sum() > 0 else 0.0

    rows = []
    for metric in ["# of units", "sq ft", "total %"]:
        r = {"Metric": metric}
        for y, mask in buckets.items():
            if metric == "# of units":
                r[y] = int(mask.sum())
            elif metric == "sq ft":
                r[y] = float(df.loc[mask, "sf"].sum())
            else:
                sf_bucket = float(df.loc[mask, "sf"].sum())
                r[y] = (sf_bucket / total_sf) if total_sf > 0 else 0.0
        rows.append(r)

    return pd.DataFrame(rows)

def vacant_occupied_summary(df: pd.DataFrame) -> pd.DataFrame:
    total_sf = df["sf"].sum() if df["sf"].sum() > 0 else 0.0
    tenant_series = df["tenant"].astype(str).fillna("")
    vacant_mask = tenant_series.str.strip().str.lower().str.contains("vacant")

    vacant_sf = float(df.loc[vacant_mask, "sf"].sum())
    occupied_sf = float(df.loc[~vacant_mask, "sf"].sum())

    return pd.DataFrame([
        {"Status": "Vacant", "sq ft": vacant_sf, "%": (vacant_sf / total_sf) if total_sf > 0 else 0.0},
        {"Status": "Occupied", "sq ft": occupied_sf, "%": (occupied_sf / total_sf) if total_sf > 0 else 0.0},
    ])

def rule_based_bullets(df: pd.DataFrame, borrower: str, address: str, as_of: date) -> str:
    payload = build_rent_roll_ai_payload(df, borrower, address, as_of)
    t = payload["totals"]
    lease = payload["lease_expiration_summary"]
    top = payload["top_tenants_by_rent"]

    # Formatting helpers
    def money(x):
        return "not provided" if x is None else f"${float(x):,.0f}"
    def pct(x):
        return "not provided" if x is None else f"{float(x):.1%}"
    def sf(x):
        return "not provided" if x is None else f"{float(x):,.0f} SF"
    def num(x):
        return "not provided" if x is None else f"{float(x):.2f}"

    total_sf = t.get("total_sf")
    total_rent = t.get("total_annual_rent")
    vacancy_sf = t.get("vacancy_sf")
    vacancy_pct = t.get("vacancy_pct")
    occupied_sf = (total_sf - vacancy_sf) if (total_sf is not None and vacancy_sf is not None) else None
    occupied_pct = (1 - vacancy_pct) if vacancy_pct is not None else None
    wale = t.get("wale_years")
    avg_rpsf = t.get("avg_monthly_rent_per_sf")
    mtm_sf = t.get("mtm_sf")
    mtm_rent_pct = t.get("mtm_rent_pct")

    # Rollover buckets
    y1, y2, y3 = lease["year_1"], lease["year_2"], lease["year_3"]
    y1_sf_pct, y1_rent_pct = y1["sf_pct"], y1["rent_pct"]
    y2_sf_pct, y2_rent_pct = y2["sf_pct"], y2["rent_pct"]
    y3_sf_pct, y3_rent_pct = y3["sf_pct"], y3["rent_pct"]

    # Tenant concentration
    top_lines = []
    for x in top:
        rt = money(x.get("rent"))
        rp = pct(x.get("rent_pct"))
        rem = num(x.get("remaining_term_yrs"))
        top_lines.append(f"{x.get('tenant')} {rt} ({rp}), rem {rem} yrs")
    conc_text = "; ".join(top_lines) if top_lines else "Top tenants not provided."

    top1_share = top[0]["rent_pct"] if len(top) else 0.0

    # Simple risk scoring (pick top 2)
    risks = []
    # Vacancy
    if vacancy_pct is not None:
        sev = 3 if vacancy_pct >= 0.15 else (2 if vacancy_pct >= 0.10 else (1 if vacancy_pct >= 0.05 else 0))
        risks.append((sev, f"Vacancy {sf(vacancy_sf)} ({pct(vacancy_pct)}) may pressure stabilized NOI without near-term leasing traction."))
    # Yr1 rollover
    sev = 3 if y1_rent_pct >= 0.50 else (2 if y1_rent_pct >= 0.35 else (1 if y1_rent_pct >= 0.20 else 0))
    risks.append((sev, f"Front-loaded rollover: Yr1 = {pct(y1_sf_pct)} SF / {pct(y1_rent_pct)} rent; renewal/retention risk is elevated."))
    # MTM
    if mtm_rent_pct is not None:
        sev = 3 if mtm_rent_pct >= 0.25 else (2 if mtm_rent_pct >= 0.15 else (1 if mtm_rent_pct >= 0.08 else 0))
        risks.append((sev, f"MTM exposure {sf(mtm_sf)} ({pct(mtm_rent_pct)} of rent) increases cash-flow volatility and re-tenanting risk."))
    # Concentration
    sev = 3 if top1_share >= 0.35 else (2 if top1_share >= 0.25 else (1 if top1_share >= 0.18 else 0))
    risks.append((sev, f"Tenant concentration: top tenant ≈ {pct(top1_share)} of rent; adverse credit/renewal outcome could materially impact DSCR."))

    # Select top 2 risks by severity
    risks_sorted = sorted(risks, key=lambda x: x[0], reverse=True)
    top_risks = [r[1] for r in risks_sorted[:2]]

    # Underwriting focus suggestions (tailored)
    focus = [
        "Confirm renewal status for Yr1 expirations + any MTM tenants (LOIs, discussions, probability-weighted outcomes).",
        "Request tenant credit quality / sales (if retail) + rent collections/arrears and leasing plan assumptions (downtime, TI/LC, broker opinion).",
        "Validate in-place vs market rent and re-leasing costs; quantify downside case on NOI/DSCR under vacancy + rollover stress."
    ]

    # Build 8 bullets (single-line, memo tone)
    bullets = [
        f"- **Snapshot:** {borrower} ({address}) as-of {as_of.isoformat()}; {sf(total_sf)}; total annual rent {money(total_rent)}; avg base rent {('not provided' if avg_rpsf is None else f'${avg_rpsf:.2f}/SF/mo')}.",
        f"- **Occupancy:** Vacancy {sf(vacancy_sf)} ({pct(vacancy_pct)}); occupied {sf(occupied_sf)} ({pct(occupied_pct)}).",
        f"- **Lease Profile:** WALE {('not provided' if wale is None else f'{wale:.2f} yrs')}; shorter WALE implies higher near-term cash-flow re-pricing/renewal execution risk.",
        f"- **Rollover:** Yr1 {pct(y1_sf_pct)} SF / {pct(y1_rent_pct)} rent; Yr2 {pct(y2_sf_pct)} / {pct(y2_rent_pct)}; Yr3 {pct(y3_sf_pct)} / {pct(y3_rent_pct)}; profile is {'front-loaded' if y1_rent_pct >= 0.35 else 'moderate'}.",
        f"- **MTM Exposure:** MTM {sf(mtm_sf)} representing {pct(mtm_rent_pct)} of rent; stability depends on near-term renewals/mark-to-market execution.",
        f"- **Concentration:** {conc_text}",
        f"- **Key Risks:** (1) {top_risks[0]} (2) {top_risks[1]}",
        f"- **Underwriting Focus:** {focus[0]} {focus[1]} {focus[2]}",
    ]
    return "\n".join(bullets)

def rule_based_stress_test_bullets(payload: dict) -> str:
    t_dscr = payload["targets"]["dscr"]
    t_ltv = payload["targets"]["ltv"]

    loan = payload["loan_terms"]
    noi_in = payload["noi_inputs"]
    col = payload["collateral"]
    base = payload["baseline_metrics"]
    pts = payload["key_stress_points"]

    def money(x): return f"\\${float(x):,.0f}"
    def pct(x): return f"{float(x):.1%}"
    def rate(x): return f"{float(x)*100:.3f}%"
    def dscr_fmt(x): return f"{float(x):.2f}x"

    # Helper to parse DSCR/LTV from stress-point dicts
    def dscr_from(d):
        if not d: return None
        v = d.get("Interest Rate Impact to DSCR") or d.get("Vacancy Rate Impact to DSCR") or d.get("NOI Change Impact to DSCR")
        if v is None: return None
        return float(str(v).replace("x","").strip())

    def ltv_from(d):
        if not d: return None
        v = d.get("Cap Rate Impact to LTV")
        if v is None: return None
        return float(str(v).replace("%","").strip())/100.0

    # Baseline status
    base_dscr = base["dscr"]
    base_ltv = base["ltv"]

    base_dscr_flag = "PASS" if base_dscr >= t_dscr else ("WATCH" if base_dscr >= 1.00 else "FAIL")
    base_ltv_flag = "PASS" if base_ltv <= t_ltv else ("WATCH" if base_ltv <= t_ltv + 0.05 else "FAIL")

    # Stresses
    dscr_200 = dscr_from(pts.get("rate_plus_200bp"))
    dscr_400 = dscr_from(pts.get("rate_plus_400bp"))
    vac_10 = dscr_from(pts.get("vac_plus_10"))
    vac_30 = dscr_from(pts.get("vac_plus_30"))
    noi_m10 = dscr_from(pts.get("noi_minus_10"))
    noi_m30 = dscr_from(pts.get("noi_minus_30"))

    ltv_100 = ltv_from(pts.get("cap_plus_100bp"))
    ltv_400 = ltv_from(pts.get("cap_plus_400bp"))

    # Pick top 2 risks (simple severity ranking)
    risks = []
    if base_dscr < t_dscr: risks.append((3, f"Baseline DSCR {dscr_fmt(base_dscr)} is below target {dscr_fmt(t_dscr)}."))
    if base_ltv > t_ltv: risks.append((3, f"Baseline LTV {pct(base_ltv)} exceeds target {pct(t_ltv)}."))

    if dscr_400 is not None and dscr_400 < 1.00: risks.append((3, f"Rate shock +400 bps drives DSCR to {dscr_fmt(dscr_400)} (sub-1.00x)."))
    elif dscr_400 is not None and dscr_400 < t_dscr: risks.append((2, f"Rate shock +400 bps reduces DSCR to {dscr_fmt(dscr_400)} (< target)."))

    if vac_30 is not None and vac_30 < 1.00: risks.append((3, f"+30% vacancy drives DSCR to {dscr_fmt(vac_30)} (sub-1.00x)."))
    elif vac_30 is not None and vac_30 < t_dscr: risks.append((2, f"+30% vacancy reduces DSCR to {dscr_fmt(vac_30)} (< target)."))

    if noi_m30 is not None and noi_m30 < 1.00: risks.append((3, f"-30% NOI drives DSCR to {dscr_fmt(noi_m30)} (sub-1.00x)."))
    elif noi_m30 is not None and noi_m30 < t_dscr: risks.append((2, f"-30% NOI reduces DSCR to {dscr_fmt(noi_m30)} (< target)."))

    if ltv_400 is not None and ltv_400 > t_ltv + 0.10: risks.append((3, f"+400 bps cap expansion pushes LTV to {pct(ltv_400)} (materially above target)."))
    elif ltv_400 is not None and ltv_400 > t_ltv: risks.append((2, f"+400 bps cap expansion pushes LTV to {pct(ltv_400)} (> target)."))

    risks = sorted(risks, key=lambda x: x[0], reverse=True)
    top_risks = [r[1] for r in risks[:2]] if risks else ["No material risks flagged from provided stresses.", "Not provided."]

    amort_label = "I/O" if loan["io"] else f"{float(loan.get('amort_years', 0.0)):.0f}yr am"

    bullets = [
        f"- **Snapshot:** Loan {money(loan['loan_amount'])} at {rate(loan['note_rate'])}; {amort_label}; annual debt service {money(loan['annual_debt_service'])}.",
        f"- **Baseline Cash Flow:** Income {money(noi_in['rental_income'])}; opex {money(noi_in['operating_expenses'])}; NOI {money(noi_in['noi'])}; DSCR {dscr_fmt(base_dscr)} vs {dscr_fmt(t_dscr)} ({base_dscr_flag}).",
        f"- **Collateral / Leverage:** Cap {rate(col['cap_rate'])} ({col.get('cap_source','not provided')}); value {money(col['estimated_value'])}; LTV {pct(base_ltv)} vs {pct(t_ltv)} ({base_ltv_flag}).",
        f"- **Rate Sensitivity:** DSCR at +200 bps = {('not provided' if dscr_200 is None else dscr_fmt(dscr_200))}; +400 bps = {('not provided' if dscr_400 is None else dscr_fmt(dscr_400))}; monitor refinance/renewal cushion.",
        f"- **Vacancy Sensitivity:** DSCR at +10% vacancy = {('not provided' if vac_10 is None else dscr_fmt(vac_10))}; +30% = {('not provided' if vac_30 is None else dscr_fmt(vac_30))}; downside hinges on leasing stability.",
        f"- **NOI Sensitivity:** DSCR at -10% NOI = {('not provided' if noi_m10 is None else dscr_fmt(noi_m10))}; -30% = {('not provided' if noi_m30 is None else dscr_fmt(noi_m30))}; validates margin for expense creep / income leakage.",
        f"- **Key Risks:** (1) {top_risks[0]} (2) {top_risks[1]}",
        f"- **Underwriting Focus:** Validate sources (income/opex/cap), normalize add-backs, confirm lease durability and sponsor support, and size TI/LC + downtime in a downside case.",
    ]
    return "\n".join(bullets)

def _safe_float(x, default=0.0):
    try:
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return default
        return float(x)
    except Exception:
        return default

def build_rent_roll_ai_payload(df: pd.DataFrame, borrower: str, address: str, as_of: date) -> dict:
    """
    df should be the *computed* rent roll dataframe (NOT the formatted one, NOT the footer row),
    containing: sf, annual_rent, monthly_rent_per_sf, remaining_term_years, tenant, lease_end, tenant_pct_total_sf, pct_total_rent.
    """
    data = df.copy()

    # Basic totals
    total_sf = _safe_float(data["sf"].sum(), 0.0)
    total_annual_rent = _safe_float(data["annual_rent"].sum(), 0.0)

    tenant_series = data["tenant"].astype(str).fillna("")
    vacant_mask = tenant_series.str.strip().str.lower().str.contains("vacant")

    vacancy_sf = _safe_float(data.loc[vacant_mask, "sf"].sum(), 0.0)
    vacancy_pct = (vacancy_sf / total_sf) if total_sf > 0 else 0.0

    # Occupied subset
    occ = data.loc[~vacant_mask].copy()

    # WALE (SF-weighted, occupied only)
    wale_years = None
    if not occ.empty and occ["sf"].sum() > 0:
        wale_years = _safe_float((occ["sf"] * occ["remaining_term_years"]).sum() / occ["sf"].sum(), None)

    # Avg monthly rent / SF (occupied)
    avg_monthly_rent_per_sf = None
    if not occ.empty:
        avg_monthly_rent_per_sf = _safe_float(occ["monthly_rent_per_sf"].mean(), None)

    # MTM exposure (based on user-entered lease_end text)
    lease_end_series = data["lease_end"].astype(str).fillna("").str.strip().str.upper()
    mtm_mask = lease_end_series.eq("MTM")
    mtm_sf = _safe_float(data.loc[mtm_mask, "sf"].sum(), 0.0)
    mtm_rent = _safe_float(data.loc[mtm_mask, "annual_rent"].sum(), 0.0)
    mtm_rent_pct = (mtm_rent / total_annual_rent) if total_annual_rent > 0 else 0.0

    # Lease expiration buckets (Yr1/Yr2/Yr3) with SF% and Rent%
    def bucket(lo, hi):
        if lo is None:
            return data["remaining_term_years"] <= hi
        if hi is None:
            return data["remaining_term_years"] > lo
        return (data["remaining_term_years"] > lo) & (data["remaining_term_years"] <= hi)

    buckets = {
        "year_1": bucket(None, 1.0),
        "year_2": bucket(1.0, 2.0),
        "year_3": bucket(2.0, 3.0),
    }

    lease_exp = {}
    for k, m in buckets.items():
        sf_k = _safe_float(data.loc[m, "sf"].sum(), 0.0)
        rent_k = _safe_float(data.loc[m, "annual_rent"].sum(), 0.0)
        lease_exp[k] = {
            "units": int(m.sum()),
            "sf": sf_k,
            "sf_pct": (sf_k / total_sf) if total_sf > 0 else 0.0,
            "rent_pct": (rent_k / total_annual_rent) if total_annual_rent > 0 else 0.0,
        }

    # Vacancy/Occupied summary
    occupied_sf = _safe_float(data.loc[~vacant_mask, "sf"].sum(), 0.0)
    vacancy_occupied_summary = [
        {"status": "Vacant", "sf": vacancy_sf, "sf_pct": vacancy_pct},
        {"status": "Occupied", "sf": occupied_sf, "sf_pct": (occupied_sf / total_sf) if total_sf > 0 else 0.0},
    ]

    # Top tenants by RENT (exclude vacant)
    top_tenants = []
    if not occ.empty and total_annual_rent > 0:
        g = (
            occ.groupby("tenant", dropna=False)
               .agg(rent=("annual_rent", "sum"),
                    sf=("sf", "sum"),
                    remaining_term_yrs=("remaining_term_years", "mean"))
               .sort_values("rent", ascending=False)
               .head(3)
        )
        for tenant, row in g.iterrows():
            rent = _safe_float(row["rent"], 0.0)
            sf = _safe_float(row["sf"], 0.0)
            top_tenants.append({
                "tenant": str(tenant),
                "rent": rent,
                "rent_pct": (rent / total_annual_rent) if total_annual_rent > 0 else 0.0,
                "sf": sf,
                "sf_pct": (sf / total_sf) if total_sf > 0 else 0.0,
                "remaining_term_yrs": _safe_float(row["remaining_term_yrs"], 0.0),
            })

    payload = {
        "borrower_name": borrower,
        "borrower_address": address,
        "as_of_date": as_of.isoformat(),
        "totals": {
            "total_sf": total_sf,
            "total_annual_rent": total_annual_rent,
            "vacancy_sf": vacancy_sf,
            "vacancy_pct": vacancy_pct,
            "wale_years": wale_years,
            "avg_monthly_rent_per_sf": avg_monthly_rent_per_sf,
            "mtm_sf": mtm_sf,
            "mtm_rent_pct": mtm_rent_pct,
        },
        "lease_expiration_summary": lease_exp,
        "vacancy_occupied_summary": vacancy_occupied_summary,
        "top_tenants_by_rent": top_tenants,
        # Optional: you can add more later (options summary, rent/SF dispersion, tenant since, etc.)
    }
    return payload

def build_construction_ai_payload(
    bt: dict,
    base_takeout_ltv: float,
    base_takeout_dscr: float,
    base_binding: float,
    stress_binding_20: float,
    stress_binding_30: float,
    cap_sens_binding: dict,
    rate_sens_binding: dict,
) -> dict:
    return {
        "date_today": date.today().isoformat(),
        "actual_loan_terms": {
            "loan_commitment": float(bt.get("loan_commitment", 0.0)),
            "proforma_noi": float(bt.get("proforma_noi", 0.0)),
            "appraisal_cap_rate": float(bt.get("appraisal_cap_rate", 0.0)),
        },
        "takeout_terms": {
            "amort_years": float(bt.get("takeout_amort_years", 0.0)),
            "dscr_target": float(bt.get("takeout_dscr", 0.0)),
            "ltv_target": float(bt.get("takeout_ltv", 0.0)),
            "underwriting_rate": float(bt.get("underwriting_rate", 0.0)),
        },
        "baseline_sizing": {
            "max_takeout_ltv": float(base_takeout_ltv),
            "max_takeout_dscr": float(base_takeout_dscr),
            "binding_takeout": float(base_binding),
        },
        "downside_sizing": {
            "binding_takeout_noi_minus_20": float(stress_binding_20),
            "binding_takeout_noi_minus_30": float(stress_binding_30),
        },
        "cap_sensitivity": cap_sens_binding or {},
        "rate_sensitivity": rate_sens_binding or {},
    }

def rule_based_construction_bullets(payload: dict) -> str:
    a = payload["actual_loan_terms"]
    t = payload["takeout_terms"]
    b = payload["baseline_sizing"]
    d = payload["downside_sizing"]

    loan = float(a.get("loan_commitment", 0.0))
    noi = float(a.get("proforma_noi", 0.0))
    cap = float(a.get("appraisal_cap_rate", 0.0))

    base_ltv = float(b.get("max_takeout_ltv", 0.0))
    base_dscr = float(b.get("max_takeout_dscr", 0.0))
    base_bind = float(b.get("binding_takeout", 0.0))

    bind_type = "LTV" if base_ltv <= base_dscr else "DSCR"
    surplus = base_bind - loan

    cap_s = payload.get("cap_sensitivity", {}) or {}
    rate_s = payload.get("rate_sensitivity", {}) or {}

    def money(x): return f"${float(x):,.0f}"
    def pct(x): return f"{float(x)*100:.2f}%"
    def rate(x): return f"{float(x)*100:.2f}%"

    cap_line = "not provided"
    if len(cap_s):
        k = list(cap_s.keys())[0]
        cap_line = f"{k.replace('_',' ')} → {money(cap_s[k].get('binding_takeout',0.0))}"

    rate_line = "not provided"
    if len(rate_s):
        k = list(rate_s.keys())[0]
        rate_line = f"{k.replace('_',' ')} → {money(rate_s[k].get('binding_takeout',0.0))}"

    bullets = [
        f"- **Snapshot:** As-of {payload['date_today']}; bridge {money(loan)}; pro forma NOI {money(noi)}; appraisal cap {pct(cap)}.",
        f"- **Takeout Terms:** {t.get('amort_years',0):.0f}yr am; DSCR {t.get('dscr_target',0):.2f}x; LTV {pct(t.get('ltv_target',0))}; UW rate {rate(t.get('underwriting_rate',0))}.",
        f"- **Base Takeout Sizing:** Max @ LTV {money(base_ltv)}; max @ DSCR {money(base_dscr)}; binding = {bind_type} ({money(base_bind)}).",
        f"- **Refi Feasibility:** Binding takeout {money(base_bind)} vs bridge {money(loan)} → {'surplus' if surplus>=0 else 'shortfall'} {money(abs(surplus))}.",
        f"- **NOI Downside:** -20% NOI {money(d.get('binding_takeout_noi_minus_20',0.0))}; -30% NOI {money(d.get('binding_takeout_noi_minus_30',0.0))}; sensitivity is material.",
        f"- **Cap Sensitivity:** {cap_line}; cap expansion reduces sale price and LTV capacity; validate exit cap support.",
        f"- **Rate Sensitivity:** {rate_line}; higher perm debt rates compress DSCR capacity; confirm takeout market terms.",
        f"- **Underwriting Focus:** Confirm stabilization timing, leasing/pro forma support, contingency & interest reserve, and sponsor liquidity/guaranty capacity.",
    ]
    return "\n".join(bullets)

def clamp_avg_term(avg_term_years: float) -> float:
    # Excel-like: if >5 -> 5; if <2 -> 2; else exact
    if avg_term_years is None or np.isnan(avg_term_years):
        return 3.0
    return float(min(5.0, max(2.0, avg_term_years)))

def get_year_sf_from_lease_expiration(computed_rr: pd.DataFrame) -> dict:
    # Uses your existing bucketing logic (Year1<=1, Year2 (1,2], Year3 (2,3])
    les = lease_expiration_summary(computed_rr)
    # les rows: "# of units", "sq ft", "total %"
    out = {"Year 1": 0.0, "Year 2": 0.0, "Year 3": 0.0}
    try:
        sq_row = les[les["Metric"] == "sq ft"].iloc[0]
        out["Year 1"] = float(sq_row["Year 1"])
        out["Year 2"] = float(sq_row["Year 2"])
        out["Year 3"] = float(sq_row["Year 3"])
    except Exception:
        pass
    return out

def build_rollover_risk_outputs(
    borrower: str,
    address: str,
    as_of: date,
    computed_rr: pd.DataFrame,
    stress_state: dict,
    leasing_commission_pct: float,
    ti_per_sf: float,
    market_rent_per_sf_yr: float,
    rent_loss_months: float,
) -> tuple[pd.DataFrame, dict]:
    # Pull NOI + actual debt service from Stress Test state
    noi = float(stress_state.get("rental_income", 0.0) - stress_state.get("operating_expenses", 0.0))
    debt_service_calc = annual_debt_service(
        stress_state.get("loan_amount", 0.0),
        stress_state.get("note_rate", 0.0),
        stress_state.get("amort_years", 0.0),
        bool(stress_state.get("io", False)),
    )
    debt_service = float(stress_state.get("annual_debt_service_override", 0.0)) if bool(stress_state.get("debt_service_override_on", False)) else float(debt_service_calc)

    baseline_dscr = (noi / debt_service) if debt_service > 0 else 0.0

    # Avg original lease term from Rent Roll (use your computed col; ignore NaNs)
    avg_orig = computed_rr["original_lease_term_years"].mean(skipna=True)
    avg_term_used = clamp_avg_term(float(avg_orig) if avg_orig is not None else np.nan)

    # Avg monthly rent/SF from rent roll (occupied mean already available in payload builder logic, but reuse here)
    # Keep it simple: mean of computed_rr['monthly_rent_per_sf'] excluding NaN
    avg_monthly_rent_per_sf = float(computed_rr["monthly_rent_per_sf"].mean(skipna=True)) if len(computed_rr) else 0.0

    # Year SF (from lease expiration summary)
    year_sf = get_year_sf_from_lease_expiration(computed_rr)

    # Estimated costs by year
    # Leasing Commission: market rent ($/SF/yr) * SF * LC% * avg term (yrs)
    lc_cost = {y: (market_rent_per_sf_yr * sf * leasing_commission_pct * avg_term_used) for y, sf in year_sf.items()}

    # TI: TI ($/SF) * SF
    ti_cost = {y: (ti_per_sf * sf) for y, sf in year_sf.items()}

    # Rent loss: SF * avg monthly rent/SF * months
    rent_loss_cost = {y: (sf * avg_monthly_rent_per_sf * rent_loss_months) for y, sf in year_sf.items()}

    total_cost = {y: (lc_cost[y] + ti_cost[y] + rent_loss_cost[y]) for y in year_sf.keys()}
    adj_noi = {y: (noi - total_cost[y]) for y in year_sf.keys()}
    adj_dscr = {y: (adj_noi[y] / debt_service) if debt_service > 0 else 0.0 for y in year_sf.keys()}

    # Build the display table (matches your Excel layout)
    cols = ["Year 1", "Year 2", "Year 3"]
    table_rows = [
        ("NOI (from Stress Test)", {c: noi for c in cols}),
        ("Estimated Costs", {c: "" for c in cols}),  # header row
        ("Leasing Commissions", {c: lc_cost[c] for c in cols}),
        ("Tenant Improvements (TI)", {c: ti_cost[c] for c in cols}),
        ("Rent Loss", {c: rent_loss_cost[c] for c in cols}),
        ("—", {c: "—" for c in cols}),
        ("Total Estimated Costs", {c: total_cost[c] for c in cols}),
        ("", {c: "" for c in cols}),
        ("Adjusted NOI", {c: adj_noi[c] for c in cols}),
        ("Actual Debt Service", {c: debt_service for c in cols}),
        ("— ", {c: "—" for c in cols}),
        ("Adjusted DSCR", {c: adj_dscr[c] for c in cols}),
    ]

    df = pd.DataFrame([{"Row": name, **vals} for name, vals in table_rows])

    payload = {
        "borrower_name": borrower,
        "borrower_address": address,
        "as_of_date": as_of.isoformat(),
        "noi": noi,
        "annual_debt_service": debt_service,
        "baseline_dscr": baseline_dscr,
        "assumptions": {
            "avg_original_lease_term_years_raw": None if (avg_orig is None or np.isnan(avg_orig)) else float(avg_orig),
            "avg_term_used_years": avg_term_used,
            "leasing_commission_pct": float(leasing_commission_pct),
            "ti_per_sf": float(ti_per_sf),
            "market_rent_per_sf_yr": float(market_rent_per_sf_yr),
            "rent_loss_months": float(rent_loss_months),
            "avg_monthly_rent_per_sf": float(avg_monthly_rent_per_sf),
        },
        "rollover_sf": year_sf,
        "costs": {
            "leasing_commissions": lc_cost,
            "tenant_improvements": ti_cost,
            "rent_loss": rent_loss_cost,
            "total_estimated_costs": total_cost,
        },
        "adjusted": {
            "adjusted_noi": adj_noi,
            "adjusted_dscr": adj_dscr,
        },
        "targets": {
            "dscr": float(stress_state.get("target_dscr", 1.20)),
        }
    }

    return df, payload

def rule_based_rollover_bullets(payload: dict) -> str:
    noi = float(payload.get("noi", 0.0))
    ds = float(payload.get("annual_debt_service", 0.0))
    base_dscr = float(payload.get("baseline_dscr", 0.0))
    t_dscr = float(payload.get("targets", {}).get("dscr", 1.20))

    ysf = payload.get("rollover_sf", {})
    costs = payload.get("costs", {})
    adj = payload.get("adjusted", {})

    def money(x): return f"${float(x):,.0f}"
    def dscr_fmt(x): return f"{float(x):.2f}x"

    # Biggest year by total costs
    tot = costs.get("total_estimated_costs", {})
    biggest_year = max(tot, key=lambda k: float(tot.get(k, 0.0))) if tot else "not provided"

    # Biggest driver overall
    drivers = {
        "Leasing Commissions": sum(float(v) for v in (costs.get("leasing_commissions", {}) or {}).values()),
        "Tenant Improvements": sum(float(v) for v in (costs.get("tenant_improvements", {}) or {}).values()),
        "Rent Loss": sum(float(v) for v in (costs.get("rent_loss", {}) or {}).values()),
    }
    biggest_driver = max(drivers, key=drivers.get) if drivers else "not provided"

    # Worst adjusted DSCR
    adj_dscr = adj.get("adjusted_dscr", {}) or {}
    worst_year = min(adj_dscr, key=lambda k: float(adj_dscr.get(k, 0.0))) if adj_dscr else "not provided"
    worst_dscr = float(adj_dscr.get(worst_year, 0.0)) if worst_year != "not provided" else 0.0

    flag = "PASS" if worst_dscr >= t_dscr else ("WATCH" if worst_dscr >= 1.00 else "FAIL")

    a = payload.get("assumptions", {})
    borrower = payload.get("borrower_name", "not provided")
    address = payload.get("borrower_address", "not provided")
    as_of = payload.get("as_of_date", "not provided")

    bullets = [
        f"- **Snapshot:** {borrower} ({address}) as-of {as_of}; NOI {money(noi)}; debt service {money(ds)}; baseline DSCR {dscr_fmt(base_dscr)}.",
        f"- **Rollover SF:** Yr1 {ysf.get('Year 1',0):,.0f} SF; Yr2 {ysf.get('Year 2',0):,.0f} SF; Yr3 {ysf.get('Year 3',0):,.0f} SF (from lease-expiration buckets).",
        f"- **Cost Assumptions:** term used {float(a.get('avg_term_used_years',0)):.2f} yrs; LC {float(a.get('leasing_commission_pct',0))*100:.2f}%; TI {money(a.get('ti_per_sf',0))}/SF; market {money(a.get('market_rent_per_sf_yr',0))}/SF/yr; rent loss {float(a.get('rent_loss_months',0)):.1f} mos.",
        f"- **Estimated Costs:** largest year = {biggest_year} ({money(tot.get(biggest_year,0))}); primary driver = {biggest_driver}; results are conservative (assumes re-tenanting for all expiring SF).",
        f"- **Adjusted NOI:** Yr1 {money(adj.get('adjusted_noi',{}).get('Year 1',0))}; Yr2 {money(adj.get('adjusted_noi',{}).get('Year 2',0))}; Yr3 {money(adj.get('adjusted_noi',{}).get('Year 3',0))}.",
        f"- **Adjusted DSCR:** Yr1 {dscr_fmt(adj_dscr.get('Year 1',0))}; Yr2 {dscr_fmt(adj_dscr.get('Year 2',0))}; Yr3 {dscr_fmt(adj_dscr.get('Year 3',0))}; worst {worst_year} = {dscr_fmt(worst_dscr)} vs {dscr_fmt(t_dscr)} ({flag}).",
        f"- **Key Risks:** (1) Costs concentrated in {biggest_year} could compress coverage. (2) Assumptions (TI/LC/rent loss) are sensitive; small changes can move DSCR below policy.",
        f"- **Underwriting Focus:** Confirm renewals/LOIs for Yr1 expirations; validate TI/LC budget support and downtime; assess sponsor liquidity to cover rollover costs.",
    ]
    return "\n".join(bullets)

def annual_debt_service(loan_amt: float, note_rate: float, amort_years: float | None, io: bool) -> float:
    """
    loan_amt: principal
    note_rate: annual rate in decimal (e.g., 0.075)
    amort_years: years (ignored if io=True)
    io: interest-only if True
    returns annual debt service ($/yr)
    """
    loan_amt = float(loan_amt or 0.0)
    note_rate = float(note_rate or 0.0)

    if loan_amt <= 0 or note_rate < 0:
        return 0.0

    if io:
        return loan_amt * note_rate

    # amortizing
    if amort_years is None or amort_years <= 0:
        return 0.0

    r = note_rate / 12.0
    n = int(round(amort_years * 12))

    if r == 0:
        pmt = loan_amt / n
    else:
        pmt = loan_amt * (r / (1 - (1 + r) ** (-n)))

    return pmt * 12.0

def fmt_money(x):
    try:
        return f"${float(x):,.0f}"
    except Exception:
        return "—"

def fmt_pct(x):
    try:
        return f"{float(x):.2%}"
    except Exception:
        return "—"

def fmt_rate_pp(x):
    # x as decimal (0.005 -> +0.50%)
    try:
        v = float(x) * 100
        sign = "+" if v > 0 else ""
        return f"{sign}{v:.2f}%"
    except Exception:
        return "—"

def build_stress_test_ai_payload(s: dict, noi: float, debt_service: float, value: float, dscr: float, ltv: float,
                                cap_df: pd.DataFrame, ir_df: pd.DataFrame, vac_df: pd.DataFrame, noi_df: pd.DataFrame) -> dict:
    # Pull out “baseline” and a few key stress points so the model doesn’t have to interpret long tables
    def find_row(df, col, target):
        m = df[df[col].astype(str).str.strip().str.lower() == str(target).strip().lower()]
        return m.iloc[0].to_dict() if len(m) else None

    payload = {
        "date_today": date.today().isoformat(),
        "targets": {"dscr": float(s.get("target_dscr", 1.20)), "ltv": float(s.get("target_ltv", 0.75))},
        "loan_terms": {
            "loan_amount": float(s.get("loan_amount", 0.0)),
            "note_rate": float(s.get("note_rate", 0.0)),
            "io": bool(s.get("io", False)),
            "amort_years": None if bool(s.get("io", False)) else float(s.get("amort_years", 0.0)),
            "annual_debt_service": float(debt_service),
            "loan_source": s.get("loan_source"),
        },
        "noi_inputs": {
            "rental_income": float(s.get("rental_income", 0.0)),
            "operating_expenses": float(s.get("operating_expenses", 0.0)),
            "noi": float(noi),
            "income_source": s.get("income_source"),
            "opex_source": s.get("opex_source"),
        },
        "collateral": {
            "cap_rate": float(s.get("cap_rate", 0.0)),
            "cap_source": s.get("cap_source"),
            "estimated_value": float(value),
        },
        "baseline_metrics": {"dscr": float(dscr), "ltv": float(ltv)},
        "stress_tables": {
            "cap_rate_effect": cap_df.to_dict(orient="records"),
            "interest_rate_effect": ir_df.to_dict(orient="records"),
            "vacancy_rate_effect": vac_df.to_dict(orient="records"),
            "noi_change_effect": noi_df.to_dict(orient="records"),
        },
        "key_stress_points": {
            "rate_plus_200bp": find_row(ir_df, "Interest Rate Change", "+2.00%"),
            "rate_plus_400bp": find_row(ir_df, "Interest Rate Change", "+4.00%"),
            "vac_plus_10": find_row(vac_df, "Vacancy Rate Change", "+10%"),
            "vac_plus_30": find_row(vac_df, "Vacancy Rate Change", "+30%"),
            "noi_minus_10": find_row(noi_df, "NOI Change", "-10%"),
            "noi_minus_30": find_row(noi_df, "NOI Change", "-30%"),
            "cap_plus_100bp": find_row(cap_df, "Cap Rates Change", "+1.00%"),
            "cap_plus_400bp": find_row(cap_df, "Cap Rates Change", "+4.00%"),
        }
    }
    return payload

def max_loan_from_annual_debt_service(annual_ds: float, note_rate: float, amort_years: float) -> float:
    """
    Inverts an amortizing payment to solve for principal:
    annual_ds = PMT_monthly(P, r, n) * 12  -> solve for P.
    """
    annual_ds = float(annual_ds or 0.0)
    note_rate = float(note_rate or 0.0)
    amort_years = float(amort_years or 0.0)

    if annual_ds <= 0 or amort_years <= 0 or note_rate < 0:
        return 0.0

    rm = note_rate / 12.0
    n = int(round(amort_years * 12))

    if n <= 0:
        return 0.0

    if rm == 0:
        # monthly pmt = P/n, annual = 12P/n -> P = annual * n/12
        return annual_ds * n / 12.0

    ann_factor = (rm / (1 - (1 + rm) ** (-n))) * 12.0
    if ann_factor <= 0:
        return 0.0

    return annual_ds / ann_factor

def _build_default_grid(base: float, step: float, rows: int) -> list[float]:
    # base/step are decimals (0.055 = 5.5%)
    rows = int(rows or 5)
    rows = max(1, min(rows, 30))
    return [max(0.0, float(base) + i * float(step)) for i in range(rows)]

def _noi_scenarios(proforma_noi: float) -> dict:
    noi = float(proforma_noi or 0.0)
    return {
        "Pro Forma": noi,
        "Pro Forma (-20%)": noi * 0.80,
        "Pro Forma (-30%)": noi * 0.70,
    }

def _fmt_money_df(df: pd.DataFrame, money_cols: list[str]) -> pd.DataFrame:
    out = df.copy()
    for c in money_cols:
        if c in out.columns:
            out[c] = out[c].apply(lambda x: "" if x == "" or pd.isna(x) else fmt_money(x))
    return out

def _fmt_pct_col(series: pd.Series) -> pd.Series:
    def _f(x):
        try:
            return f"{float(x)*100:.2f}%"
        except Exception:
            return ""
    return series.apply(_f)

def build_stress_tables(s: dict, noi: float, debt_service: float):
    cap_deltas = [0.0] + [i/1000 for i in range(5, 45, 5)]  # 0.00%..+4.00%
    ir_deltas  = cap_deltas[:]
    vac_deltas = [("As is", 0.0)] + [(f"+{i}%", i/100) for i in [5,10,15,20,25,30]]
    noi_deltas = [(f"+{i}%", i/100) for i in [30,20,10]] + [("0.00%", 0.0)] + [(f"-{i}%", -i/100) for i in [10,20,30]]

    # Cap Rate Effect (LTV)
    cap_rows = []
    for d in cap_deltas:
        stressed_cap = float(s.get("cap_rate", 0.0)) + d
        stressed_value = (noi / stressed_cap) if stressed_cap > 0 else 0.0
        stressed_ltv = (float(s.get("loan_amount", 0.0)) / stressed_value) if stressed_value > 0 else 0.0
        cap_rows.append({"Cap Rates Change": fmt_rate_pp(d), "Cap Rate Impact to LTV": fmt_pct(stressed_ltv)})
    cap_df = pd.DataFrame(cap_rows)

    # Interest Rate Effect (DSCR) — uses calc DS, matches your page logic
    ir_rows = []
    for d in ir_deltas:
        stressed_rate = float(s.get("note_rate", 0.0)) + d
        stressed_ds = annual_debt_service(float(s.get("loan_amount", 0.0)), stressed_rate, float(s.get("amort_years", 0.0)), bool(s.get("io", False)))
        stressed_dscr = (noi / stressed_ds) if stressed_ds > 0 else 0.0
        ir_rows.append({"Interest Rate Change": fmt_rate_pp(d), "Interest Rate Impact to DSCR": f"{stressed_dscr:.2f}x"})
    ir_df = pd.DataFrame(ir_rows)

    # Vacancy Rate Effect (DSCR) — uses “debt_service” which may be override
    vac_rows = []
    for label, d in vac_deltas:
        stressed_income = float(s.get("rental_income", 0.0)) * (1 - d)
        stressed_noi = stressed_income - float(s.get("operating_expenses", 0.0))
        stressed_dscr = (stressed_noi / debt_service) if debt_service > 0 else 0.0
        vac_rows.append({"Vacancy Rate Change": label, "Vacancy Rate Impact to DSCR": f"{stressed_dscr:.2f}x"})
    vac_df = pd.DataFrame(vac_rows)

    # NOI Change Effect (DSCR)
    noi_rows = []
    for label, d in noi_deltas:
        stressed_noi = noi * (1 + d)
        stressed_dscr = (stressed_noi / debt_service) if debt_service > 0 else 0.0
        noi_rows.append({"NOI Change": label, "NOI Change Impact to DSCR": f"{stressed_dscr:.2f}x"})
    noi_df = pd.DataFrame(noi_rows)

    return cap_df, ir_df, vac_df, noi_df

