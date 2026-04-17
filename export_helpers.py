"""
Helpers to collect formatted table data from Streamlit session state for Word export.
"""
import pandas as pd
import numpy as np
from datetime import date


def fmt_money(x):
    try:
        return f"${float(x):,.0f}"
    except Exception:
        return ""

def fmt_pct(x):
    try:
        return f"{float(x):.2%}"
    except Exception:
        return ""

def safe_str(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    return str(v)


def collect_rent_roll_table(computed_with_footer: pd.DataFrame) -> list:
    """Format the rent roll + totals row into a list of dicts for JS."""
    rows = []
    for _, row in computed_with_footer.iterrows():
        r = {}
        r["suite"] = safe_str(row.get("suite", ""))
        r["tenant"] = safe_str(row.get("tenant", ""))
        try:
            r["sf"] = f"{float(row['sf']):,.0f}" if row["sf"] != "" and not pd.isna(row["sf"]) else ""
        except Exception:
            r["sf"] = ""
        try:
            r["pct_sf"] = f"{float(row['tenant_pct_total_sf']):.1%}" if row["tenant_pct_total_sf"] != "" and not pd.isna(row["tenant_pct_total_sf"]) else ""
        except Exception:
            r["pct_sf"] = ""
        r["lease_start"] = safe_str(row.get("lease_start", ""))
        r["lease_end"] = safe_str(row.get("lease_end", ""))
        try:
            r["rem_term"] = f"{float(row['remaining_term_years']):.1f}" if row["remaining_term_years"] != "" and not pd.isna(row["remaining_term_years"]) else ""
        except Exception:
            r["rem_term"] = ""
        try:
            r["base_rent"] = fmt_money(row["base_monthly_rent"]) if row["base_monthly_rent"] != "" and not pd.isna(row["base_monthly_rent"]) else ""
        except Exception:
            r["base_rent"] = ""
        try:
            r["cams"] = fmt_money(row["monthly_cams"]) if row["monthly_cams"] != "" and not pd.isna(row["monthly_cams"]) else ""
        except Exception:
            r["cams"] = ""
        try:
            r["annual"] = fmt_money(row["annual_rent"]) if row["annual_rent"] != "" and not pd.isna(row["annual_rent"]) else ""
        except Exception:
            r["annual"] = ""
        try:
            r["pct_rent"] = f"{float(row['pct_total_rent']):.1%}" if row["pct_total_rent"] != "" and not pd.isna(row["pct_total_rent"]) else ""
        except Exception:
            r["pct_rent"] = ""
        rows.append(r)
    return rows


def collect_lease_expiration(les_df: pd.DataFrame) -> list:
    rows = []
    for _, row in les_df.iterrows():
        r = {"Metric": str(row["Metric"])}
        for y in ["Year 1", "Year 2", "Year 3"]:
            if row["Metric"] == "total %":
                r[y] = f"{float(row[y]):.1%}"
            elif row["Metric"] == "sq ft":
                r[y] = f"{float(row[y]):,.0f}"
            else:
                r[y] = str(int(row[y]))
        rows.append(r)
    return rows


def collect_vac_occ(vo_df: pd.DataFrame) -> list:
    rows = []
    for _, row in vo_df.iterrows():
        rows.append({
            "Status": str(row["Status"]),
            "sq_ft": f"{float(row['sq ft']):,.0f}",
            "pct": f"{float(row['%']):.1%}",
        })
    return rows


def collect_stress_entry(s: dict, noi: float, debt_service: float, value: float, dscr: float, ltv: float) -> list:
    return [
        {"Item": "Current Loan Amount", "Source": s.get("loan_source",""), "Terms": fmt_money(s.get("loan_amount",0))},
        {"Item": "Note Rate", "Source": s.get("loan_source",""), "Terms": f"{float(s.get('note_rate',0))*100:.3f}%"},
        {"Item": "Amortization / I-O", "Source": "Actual Terms", "Terms": ("I/O" if s.get("io") else f"{float(s.get('amort_years',0)):.0f}")},
        {"Item": "Annual Debt Service", "Source": "Actual Terms", "Terms": fmt_money(debt_service)},
        {"Item": "", "Source": "", "Terms": ""},
        {"Item": "Rental Income", "Source": s.get("income_source",""), "Terms": fmt_money(s.get("rental_income",0))},
        {"Item": "Operating Expenses", "Source": s.get("opex_source",""), "Terms": fmt_money(s.get("operating_expenses",0))},
        {"Item": "NOI", "Source": "Calculated", "Terms": fmt_money(noi)},
        {"Item": "Cap Rate", "Source": s.get("cap_source",""), "Terms": f"{float(s.get('cap_rate',0))*100:.3f}%"},
        {"Item": "Collateral Value", "Source": "Calculated", "Terms": fmt_money(value)},
        {"Item": "DSCR", "Source": "Calculated", "Terms": f"{dscr:.2f}x"},
        {"Item": "LTV", "Source": "Calculated", "Terms": fmt_pct(ltv)},
    ]


def collect_stress_tables(cap_df, ir_df, vac_df, noi_df):
    return (
        cap_df.to_dict(orient="records") if cap_df is not None else None,
        ir_df.to_dict(orient="records") if ir_df is not None else None,
        vac_df.to_dict(orient="records") if vac_df is not None else None,
        noi_df.to_dict(orient="records") if noi_df is not None else None,
    )


def collect_rollover_table(rollover_df: pd.DataFrame) -> list:
    money_rows = {"NOI (from Stress Test)", "Leasing Commissions", "Tenant Improvements (TI)",
                  "Rent Loss", "Total Estimated Costs", "Adjusted NOI", "Actual Debt Service"}
    dscr_rows = {"Adjusted DSCR"}
    rows = []
    for _, row in rollover_df.iterrows():
        rd = {"Row": str(row["Row"])}
        for c in ["Year 1", "Year 2", "Year 3"]:
            v = row[c]
            rname = str(row["Row"])
            if rname in money_rows:
                rd[c] = "" if v == "" or v == "\u2014" else fmt_money(v)
            elif rname in dscr_rows:
                try:
                    rd[c] = f"{float(v):.2f}x"
                except Exception:
                    rd[c] = str(v)
            else:
                rd[c] = str(v) if v else ""
        rows.append(rd)
    return rows


def collect_construction_tables(bt, scenarios, cap_grid, ir_grid,
                                ltv_takeout, dscr_takeout, max_loan_fn):
    """Build the 3 construction grid tables + sizing + NOI scenarios."""
    constr_noi = {k: fmt_money(v) for k, v in scenarios.items()}

    constr_cap_ltv = []
    for cap in cap_grid:
        r = {"Cap Rate": f"{cap*100:.2f}%"}
        for sc_name, noi_val in scenarios.items():
            sale = (noi_val / cap) if cap > 0 else 0.0
            mx = sale * float(bt.get("takeout_ltv", 0.0))
            r[sc_name] = fmt_money(mx)
        constr_cap_ltv.append(r)

    constr_ir_dscr = []
    for rr in ir_grid:
        r = {"Interest Rate": f"{rr*100:.2f}%"}
        for sc_name, noi_val in scenarios.items():
            max_ads = (noi_val / float(bt.get("takeout_dscr", 1.0))) if float(bt.get("takeout_dscr", 0.0)) > 0 else 0.0
            mx = max_loan_fn(max_ads, rr, float(bt.get("takeout_amort_years", 0.0)))
            r[sc_name] = fmt_money(mx)
        constr_ir_dscr.append(r)

    constr_sale = []
    for cap in cap_grid:
        r = {"Cap Rate": f"{cap*100:.2f}%"}
        for sc_name, noi_val in scenarios.items():
            sale = (noi_val / cap) if cap > 0 else 0.0
            r[sc_name] = fmt_money(sale)
        constr_sale.append(r)

    constr_sizing = {
        "max_ltv": fmt_money(ltv_takeout),
        "max_dscr": fmt_money(dscr_takeout),
        "bridge": fmt_money(float(bt.get("loan_commitment", 0.0))),
    }

    return constr_noi, constr_cap_ltv, constr_ir_dscr, constr_sale, constr_sizing