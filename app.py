
import copy
import html
import re
import tempfile
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

from analysis_core import (
    default_rent_roll_df,
    default_stress_test_state,
    compute_rent_roll,
    totals_row,
    lease_expiration_summary,
    vacant_occupied_summary,
    rule_based_bullets,
    annual_debt_service,
    build_stress_tables,
    build_stress_test_ai_payload,
    rule_based_stress_test_bullets,
    build_rollover_risk_outputs,
    rule_based_rollover_bullets,
    _noi_scenarios,
    _build_default_grid,
    max_loan_from_annual_debt_service,
    build_construction_ai_payload,
    rule_based_construction_bullets,
    fmt_money,
    fmt_pct,
)
from export_credit_report import create_credit_analysis_report
from export_helpers import (
    collect_rent_roll_table,
    collect_lease_expiration,
    collect_vac_occ,
    collect_stress_entry,
    collect_stress_tables,
    collect_rollover_table,
    collect_construction_tables,
)

st.set_page_config(page_title="HB Relationship Review Studio", layout="wide")

APP_TITLE = "HB Relationship Review Studio"
DB_BG = "#EAF2FF"
MANUAL_BG = "#FFF6E8"
DB_BORDER = "#8CB4FF"
MANUAL_BORDER = "#E3B46D"

COMMON_COGNOS_FIELDS = [
    ("client_group_name", "Client Group Name"),
    ("group_number", "Group Number"),
    ("loan_exposure", "Loan Exposure"),
    ("deposit_relationship", "Deposit Relationship"),
    ("tier_level", "Tier Level"),
    ("tenure", "Tenure"),
    ("loan_wair", "Loan WAIR"),
    ("loan_amortization_range", "Loan Amortization Range"),
    ("deposit_wair", "Deposit WAIR"),
    ("deposit_count_by_type", "Deposit Count (By Type)"),
]

LOAN_COGNOS_FIELDS = [
    ("borrower_name", "Borrower Name"),
    ("loan_number", "Loan Number"),
    ("loan_type", "Loan Type"),
    ("loan_balance", "Loan Balance"),
    ("loan_rate", "Loan Rate"),
    ("mla_sub_notes", "MLA Sub Notes"),
    ("loan_maturity_date", "Loan Maturity Date"),
    ("collateral_type", "Collateral Type"),
    ("collateral_location", "Collateral Location"),
    ("loan_recourse", "Loan Recourse"),
    ("appraisal_value", "Appraisal Value"),
    ("ltv", "LTV"),
    ("current_loan_amount", "Current Loan Amount"),
    ("note_rate", "Note Rate"),
    ("amortization_period_or_io", "Amortization Period or I/O"),
    ("annual_debt_service", "Annual Debt Service"),
    ("estimated_collateral_value", "Estimated Collateral Value"),
    ("ltv_ratio", "LTV Ratio"),
    ("loan_commitment", "Loan Commitment"),
    ("current_risk_rating", "Current Risk Rating"),
    ("watch_reason", "Watch Reason"),
]

RISK_RATINGS = ["Pass", "Acceptable", "Watch", "Special Mention", "Substandard"]
RATE_TYPES = ["Fixed", "Variable", "Hybrid", "Step-Up", "Unknown"]
SCOPE_OPTIONS = ["Full", "Split"]
POTENTIAL_OPTIONS = ["Low", "Moderate", "High"]
UPDATE_OPTIONS = ["Stable", "Improving", "Mixed", "Needs Attention"]
SOURCE_OPTIONS = ["Tax Return", "P&L Statement", "Rent Roll", "Appraisal", "Lease Agreements", "Other"]
CAP_SOURCE_OPTIONS = ["Appraisal", "Broker Opinion", "Market Survey", "Other"]
CONDITION_OPTIONS = ["Excellent", "Good", "Average", "Deferred Maintenance"]
COVENANT_OPTIONS = ["In Compliance", "Minor Exception", "Breached", "Not Tested"]
BOOLEAN_OPTIONS = ["No", "Yes"]


def fmt_date(v):
    if isinstance(v, date):
        return v.isoformat()
    return str(v) if v is not None else ""


def money(x):
    try:
        return f"${float(x):,.0f}"
    except Exception:
        return str(x)


def pct(x):
    try:
        return f"{float(x):.1%}"
    except Exception:
        return str(x)


def markdownish_to_plain_text(text: str) -> str:
    if not text:
        return ""
    lines = []
    for raw in str(text).splitlines():
        line = raw.strip()
        if not line:
            continue
        m = re.match(r"^-\s*\*\*([^*]+)\*\*:?[\s]*(.*)$", line)
        if m:
            label = m.group(1).strip().rstrip(":")
            content = m.group(2).strip()
            lines.append(f"- {label}: {content}" if content else f"- {label}")
            continue
        line = line.replace("**", "")
        lines.append(line)
    return "\n".join(lines)


def commentary_to_html(text: str) -> str:
    if not text or not str(text).strip():
        return '<p style="margin:0;color:#6b7280;">No commentary yet.</p>'

    items = []
    for raw in str(text).splitlines():
        line = raw.strip()
        if not line:
            continue

        m = re.match(r"^-\s*\*\*([^*]+)\*\*:?\s*(.*)$", line)
        if m:
            label = html.escape(m.group(1).strip().rstrip(":"))
            content = html.escape(m.group(2).strip())
            items.append(f"<li><strong>{label}:</strong> {content}</li>" if content else f"<li><strong>{label}</strong></li>")
            continue

        m2 = re.match(r"^-\s*([^:]+):\s*(.*)$", line)
        if m2:
            label = html.escape(m2.group(1).strip().rstrip(":"))
            content = html.escape(m2.group(2).strip())
            items.append(f"<li><strong>{label}:</strong> {content}</li>" if content else f"<li><strong>{label}</strong></li>")
            continue

        if line.startswith('- '):
            items.append(f"<li>{html.escape(line[2:].strip())}</li>")
        else:
            items.append(f"<li>{html.escape(line)}</li>")

    return (
        "<div style='background:#f8fafc;border:1px solid #e5e7eb;border-radius:10px;padding:12px 14px;'>"
        "<ul style='margin:0;padding-left:1.2rem;'>" + ''.join(items) + "</ul></div>"
    )


def render_commentary_preview(title: str, text: str):
    st.markdown(f"**{title} Preview**")
    st.markdown(commentary_to_html(text), unsafe_allow_html=True)


def render_editable_commentary_block(label: str, storage: dict, field: str, key_prefix: str, preview_title: str, height: int = 220):
    edit_key = f"{key_prefix}_{field}_editing"
    draft_key = f"{key_prefix}_{field}_draft"

    if edit_key not in st.session_state:
        st.session_state[edit_key] = False
    if draft_key not in st.session_state:
        st.session_state[draft_key] = storage.get(field, "")

    current_value = storage.get(field, "")

    if st.session_state[edit_key]:
        st.session_state[draft_key] = st.text_area(
            label,
            value=st.session_state[draft_key],
            height=height,
            key=f"{key_prefix}_{field}_editor",
        )
        c1, c2 = st.columns([1, 1])
        if c1.button("Save", key=f"{key_prefix}_{field}_save", use_container_width=True):
            storage[field] = st.session_state[draft_key]
            st.session_state[edit_key] = False
            st.rerun()
        if c2.button("Cancel", key=f"{key_prefix}_{field}_cancel", use_container_width=True):
            st.session_state[draft_key] = current_value
            st.session_state[edit_key] = False
            st.rerun()
    else:
        render_commentary_preview(preview_title, current_value)
        if st.button(f"Edit {label}", key=f"{key_prefix}_{field}_edit", use_container_width=True):
            st.session_state[draft_key] = current_value
            st.session_state[edit_key] = True
            st.rerun()


def parse_amortization(value):
    text = str(value).strip().lower()
    if "i/o" in text or "io" == text:
        return True, 25.0
    digits = "".join(ch if ch.isdigit() or ch == "." else " " for ch in text)
    nums = [float(x) for x in digits.split() if x.strip()]
    return False, (nums[0] if nums else 25.0)


def make_rent_roll(seed: int):
    df = default_rent_roll_df().copy()
    df["suite"] = [f"{100 + 10 * i + seed}" for i in range(len(df))]
    df["base_monthly_rent"] = df["base_monthly_rent"] + seed * 120
    df["monthly_cams"] = df["monthly_cams"] + seed * 15
    return df


def build_demo_portfolio():
    today = date.today()
    groups = [
        {
            "group_id": "G1001",
            "group_name": "Harbor Crest Holdings",
            "branch": "Long Beach",
            "officer": "M. Nguyen",
            "tier": "Tier 1",
            "last_review_date": today - timedelta(days=92),
            "next_review_date": today + timedelta(days=273),
            "current_rbr_status": "In Progress",
            "deposit_balances": 3_400_000,
            "loans": [
                {
                    "loan_id": "G1001-L01",
                    "db": {
                        "borrower_name": "Harbor Crest Office LLC",
                        "loan_number": "90124571",
                        "loan_type": "Investor CRE",
                        "loan_balance": 6_850_000,
                        "loan_rate": "6.35%",
                        "mla_sub_notes": "Standard reporting; annual site visit",
                        "loan_maturity_date": today + timedelta(days=410),
                        "collateral_type": "Office",
                        "collateral_location": "221 Pine Ave, Long Beach, CA",
                        "loan_recourse": "Partial",
                        "appraisal_value": 11_400_000,
                        "ltv": 0.60,
                        "current_loan_amount": 6_850_000,
                        "note_rate": 0.0635,
                        "amortization_period_or_io": "25 Years",
                        "annual_debt_service": 549_000,
                        "estimated_collateral_value": 11_400_000,
                        "ltv_ratio": 0.60,
                        "loan_commitment": 7_200_000,
                        "current_risk_rating": "Acceptable",
                        "watch_reason": "None",
                    },
                },
                {
                    "loan_id": "G1001-L02",
                    "db": {
                        "borrower_name": "Harbor Crest Retail LLC",
                        "loan_number": "90124572",
                        "loan_type": "Investor CRE",
                        "loan_balance": 4_300_000,
                        "loan_rate": "6.10%",
                        "mla_sub_notes": "Retail rollover monitored quarterly",
                        "loan_maturity_date": today + timedelta(days=515),
                        "collateral_type": "Retail",
                        "collateral_location": "404 Seaside Blvd, Long Beach, CA",
                        "loan_recourse": "Full",
                        "appraisal_value": 7_000_000,
                        "ltv": 0.61,
                        "current_loan_amount": 4_300_000,
                        "note_rate": 0.061,
                        "amortization_period_or_io": "25 Years",
                        "annual_debt_service": 334_000,
                        "estimated_collateral_value": 7_000_000,
                        "ltv_ratio": 0.61,
                        "loan_commitment": 4_500_000,
                        "current_risk_rating": "Watch",
                        "watch_reason": "Tenant rollover",
                    },
                },
            ],
        },
        {
            "group_id": "G1002",
            "group_name": "Pacific Grove Partners",
            "branch": "Seal Beach",
            "officer": "A. Patel",
            "tier": "Tier 2",
            "last_review_date": today - timedelta(days=131),
            "next_review_date": today + timedelta(days=234),
            "current_rbr_status": "Not Started",
            "deposit_balances": 1_950_000,
            "loans": [
                {
                    "loan_id": "G1002-L01",
                    "db": {
                        "borrower_name": "Pacific Grove Industrial LLC",
                        "loan_number": "90124881",
                        "loan_type": "Industrial",
                        "loan_balance": 8_150_000,
                        "loan_rate": "6.55%",
                        "mla_sub_notes": "Warehouse conversion underway",
                        "loan_maturity_date": today + timedelta(days=760),
                        "collateral_type": "Industrial",
                        "collateral_location": "9500 Industry Way, Carson, CA",
                        "loan_recourse": "Partial",
                        "appraisal_value": 13_100_000,
                        "ltv": 0.62,
                        "current_loan_amount": 8_150_000,
                        "note_rate": 0.0655,
                        "amortization_period_or_io": "30 Years",
                        "annual_debt_service": 611_000,
                        "estimated_collateral_value": 13_100_000,
                        "ltv_ratio": 0.62,
                        "loan_commitment": 8_500_000,
                        "current_risk_rating": "Acceptable",
                        "watch_reason": "Lease-up monitoring",
                    },
                },
                {
                    "loan_id": "G1002-L02",
                    "db": {
                        "borrower_name": "Pacific Grove Multifamily LLC",
                        "loan_number": "90124882",
                        "loan_type": "Multifamily",
                        "loan_balance": 5_900_000,
                        "loan_rate": "5.95%",
                        "mla_sub_notes": "Seasoned asset",
                        "loan_maturity_date": today + timedelta(days=610),
                        "collateral_type": "Multifamily",
                        "collateral_location": "1450 Bayview Dr, Torrance, CA",
                        "loan_recourse": "Limited",
                        "appraisal_value": 9_400_000,
                        "ltv": 0.63,
                        "current_loan_amount": 5_900_000,
                        "note_rate": 0.0595,
                        "amortization_period_or_io": "30 Years",
                        "annual_debt_service": 420_000,
                        "estimated_collateral_value": 9_400_000,
                        "ltv_ratio": 0.63,
                        "loan_commitment": 6_100_000,
                        "current_risk_rating": "Pass",
                        "watch_reason": "None",
                    },
                },
                {
                    "loan_id": "G1002-L03",
                    "db": {
                        "borrower_name": "Pacific Grove Bridge LLC",
                        "loan_number": "90124883",
                        "loan_type": "Bridge / Construction",
                        "loan_balance": 3_700_000,
                        "loan_rate": "8.10%",
                        "mla_sub_notes": "Tenant improvement draw program",
                        "loan_maturity_date": today + timedelta(days=260),
                        "collateral_type": "Mixed Use",
                        "collateral_location": "82 Harbor Point, Redondo Beach, CA",
                        "loan_recourse": "Full",
                        "appraisal_value": 5_600_000,
                        "ltv": 0.66,
                        "current_loan_amount": 3_700_000,
                        "note_rate": 0.0810,
                        "amortization_period_or_io": "I/O",
                        "annual_debt_service": 300_000,
                        "estimated_collateral_value": 5_600_000,
                        "ltv_ratio": 0.66,
                        "loan_commitment": 4_000_000,
                        "current_risk_rating": "Watch",
                        "watch_reason": "Exit timing",
                    },
                },
            ],
        },
        {
            "group_id": "G1003",
            "group_name": "Mission Ridge Capital",
            "branch": "Los Angeles",
            "officer": "S. Romero",
            "tier": "Tier 1",
            "last_review_date": today - timedelta(days=80),
            "next_review_date": today + timedelta(days=285),
            "current_rbr_status": "Draft Ready",
            "deposit_balances": 4_950_000,
            "loans": [
                {
                    "loan_id": "G1003-L01",
                    "db": {
                        "borrower_name": "Mission Ridge Office LLC",
                        "loan_number": "90125031",
                        "loan_type": "Office",
                        "loan_balance": 9_250_000,
                        "loan_rate": "6.20%",
                        "mla_sub_notes": "Large anchor concentration",
                        "loan_maturity_date": today + timedelta(days=480),
                        "collateral_type": "Office",
                        "collateral_location": "1100 Wilshire Blvd, Los Angeles, CA",
                        "loan_recourse": "Partial",
                        "appraisal_value": 14_500_000,
                        "ltv": 0.64,
                        "current_loan_amount": 9_250_000,
                        "note_rate": 0.062,
                        "amortization_period_or_io": "25 Years",
                        "annual_debt_service": 733_000,
                        "estimated_collateral_value": 14_500_000,
                        "ltv_ratio": 0.64,
                        "loan_commitment": 9_500_000,
                        "current_risk_rating": "Watch",
                        "watch_reason": "Near-term rollover",
                    },
                },
                {
                    "loan_id": "G1003-L02",
                    "db": {
                        "borrower_name": "Mission Ridge Retail LLC",
                        "loan_number": "90125032",
                        "loan_type": "Retail",
                        "loan_balance": 5_450_000,
                        "loan_rate": "6.75%",
                        "mla_sub_notes": "Restaurant-heavy tenancy",
                        "loan_maturity_date": today + timedelta(days=370),
                        "collateral_type": "Retail",
                        "collateral_location": "1520 Sunset Blvd, Los Angeles, CA",
                        "loan_recourse": "Full",
                        "appraisal_value": 8_350_000,
                        "ltv": 0.65,
                        "current_loan_amount": 5_450_000,
                        "note_rate": 0.0675,
                        "amortization_period_or_io": "25 Years",
                        "annual_debt_service": 440_000,
                        "estimated_collateral_value": 8_350_000,
                        "ltv_ratio": 0.65,
                        "loan_commitment": 5_700_000,
                        "current_risk_rating": "Acceptable",
                        "watch_reason": "Tenant mix review",
                    },
                },
            ],
        },
    ]

    portfolio = {}
    for g_idx, group in enumerate(groups, start=1):
        total_balance = sum(l["db"]["loan_balance"] for l in group["loans"])
        total_commitment = sum(l["db"]["loan_commitment"] for l in group["loans"])
        wair = np.average([l["db"]["note_rate"] for l in group["loans"]], weights=[l["db"]["loan_balance"] for l in group["loans"]])
        common_db = {
            "client_group_name": group["group_name"],
            "group_number": group["group_id"],
            "loan_exposure": total_balance,
            "deposit_relationship": group["deposit_balances"],
            "tier_level": group["tier"],
            "tenure": f"{5 + g_idx} years",
            "loan_wair": f"{wair*100:.2f}%",
            "loan_amortization_range": "25-30 Years",
            "deposit_wair": f"{(1.95 + g_idx*0.15):.2f}%",
            "deposit_count_by_type": {"DDA": 3 + g_idx, "MMA": 2, "CD": 1},
        }
        common_manual = {
            "relationship_scope": "Full",
            "key_personnel": "CEO, CFO, Asset Manager",
            "future_deposit_potential": "Moderate",
            "future_borrowing_potential": "High",
            "future_property_acquisition": "Moderate",
            "recent_update": "Stable",
            "overall_update": "Stable",
        }
        for l_idx, loan in enumerate(group["loans"], start=1):
            loan["manual"] = {
                "loan_rate_type": "Fixed",
                "loan_term_years": 5,
                "strength_bucket": "Collateral Quality",
                "weakness_bucket": "Tenant Rollover",
                "risk_rating_recommendation": loan["db"]["current_risk_rating"],
                "relationship_risk_assessment": "Moderate",
                "site_inspection_condition": "Good",
                "covenant_review_status": "In Compliance",
                "waiver_request": "No",
                "risk_rating_change_request": "No",
                "additional_repayment_sources": "Guarantor Support",
            }
            loan["stress_manual"] = {
                "income_source": "Rent Roll",
                "rental_income": round(loan["db"]["loan_balance"] * 0.11, 0),
                "opex_source": "Tax Return",
                "operating_expenses": round(loan["db"]["loan_balance"] * 0.03, 0),
                "cap_source": "Appraisal",
                "cap_rate": 0.065 if l_idx % 2 else 0.070,
                "target_dscr": 1.20,
                "target_ltv": 0.75,
            }
            loan["rollover_manual"] = {
                "leasing_commission_pct": 0.05,
                "ti_per_sf": 25.0 + 2 * l_idx,
                "market_rent_per_sf_yr": 42.0 + 4 * l_idx,
                "rent_loss_months": 3.0,
            }
            loan["construction_manual"] = {
                "proforma_noi": round(loan["db"]["loan_balance"] * 0.12, 0),
                "appraisal_cap_rate": 0.065,
                "takeout_amort_years": 25.0,
                "takeout_dscr": 1.25,
                "takeout_ltv": 0.65,
                "underwriting_rate": max(loan["db"]["note_rate"] - 0.0075, 0.05),
                "grid_rows": 5,
            }
            loan["rent_roll"] = make_rent_roll(g_idx + l_idx)
            loan["final"] = {
                "executive_summary": "",
                "primary_repayment": "",
                "rent_roll_commentary": "",
                "stress_commentary": "",
                "rollover_commentary": "",
                "construction_commentary": "",
            }

        portfolio[group["group_id"]] = {
            "summary": {
                "group_id": group["group_id"],
                "group_name": group["group_name"],
                "branch": group["branch"],
                "officer": group["officer"],
                "active_loan_account": len(group["loans"]),
                "loan_balances": total_balance,
                "loan_commitments": total_commitment,
                "deposit_balances": group["deposit_balances"],
                "tier": group["tier"],
                "last_review_date": group["last_review_date"],
                "next_review_date": group["next_review_date"],
                "current_rbr_status": group["current_rbr_status"],
            },
            "common_db": common_db,
            "common_manual": common_manual,
            "loans": group["loans"],
        }
    return portfolio


def inject_css():
    st.markdown(
        f"""
        <style>
            .source-chip {{
                display:inline-block;
                padding: 0.18rem 0.55rem;
                border-radius: 999px;
                font-size: 0.78rem;
                font-weight: 600;
                margin-bottom: 0.35rem;
            }}
            .db-chip {{
                background:{DB_BG};
                border:1px solid {DB_BORDER};
                color:#1D4F91;
            }}
            .manual-chip {{
                background:{MANUAL_BG};
                border:1px solid {MANUAL_BORDER};
                color:#8A5A12;
            }}
            .legend-box {{
                padding:0.75rem 0.9rem;
                border-radius:0.8rem;
                margin-bottom:0.75rem;
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def init_state():
    if "portfolio" not in st.session_state:
        st.session_state.portfolio = build_demo_portfolio()
    if "page" not in st.session_state:
        st.session_state.page = "Home"
    if "selected_group_id" not in st.session_state:
        st.session_state.selected_group_id = None
    if "export_target_loan_id" not in st.session_state:
        st.session_state.export_target_loan_id = None


def get_group():
    gid = st.session_state.get("selected_group_id")
    if not gid:
        return None
    return st.session_state.portfolio[gid]


def get_loan(group, loan_id):
    for loan in group["loans"]:
        if loan["loan_id"] == loan_id:
            return loan
    return None


def field_chip(source):
    cls = "db-chip" if source == "db" else "manual-chip"
    label = "COGNOS / Database" if source == "db" else "Manual Input"
    st.markdown(f'<div class="source-chip {cls}">{label}</div>', unsafe_allow_html=True)


def render_db_value(label, value, key):
    field_chip("db")
    st.text_input(label, value=str(value), disabled=True, key=key)


def render_group_header(group):
    summary = group["summary"]
    st.title(f"{summary['group_id']} · {summary['group_name']}")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Loan Count", summary["active_loan_account"])
    c2.metric("Loan Balances", money(summary["loan_balances"]))
    c3.metric("Commitments", money(summary["loan_commitments"]))
    c4.metric("Deposit Balances", money(summary["deposit_balances"]))
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Branch", summary["branch"])
    c6.metric("Officer", summary["officer"])
    c7.metric("Last Review", fmt_date(summary["last_review_date"]))
    c8.metric("Next Review", fmt_date(summary["next_review_date"]))


def render_common_sections(group):
    st.markdown("### Relationship Information")
    left, right = st.columns(2)

    with left:
        with st.container(border=True):
            st.markdown("#### Relationship-Level COGNOS Data")
            cols = st.columns(2)
            for idx, (key, label) in enumerate(COMMON_COGNOS_FIELDS):
                value = group["common_db"][key]
                if isinstance(value, dict):
                    value = ", ".join(f"{k}: {v}" for k, v in value.items())
                with cols[idx % 2]:
                    render_db_value(label, value, key=f"{group['summary']['group_id']}_common_{key}")

    with right:
        manual = group["common_manual"]
        with st.container(border=True):
            st.markdown("#### Relationship-Level Manual Inputs")
            field_chip("manual")
            manual["relationship_scope"] = st.selectbox(
                "Relationship Scope", SCOPE_OPTIONS,
                index=SCOPE_OPTIONS.index(manual["relationship_scope"]),
                key=f"{group['summary']['group_id']}_relationship_scope",
            )
            field_chip("manual")
            manual["key_personnel"] = st.text_input(
                "Key Personnel",
                value=manual["key_personnel"],
                key=f"{group['summary']['group_id']}_key_personnel",
            )
            c1, c2 = st.columns(2)
            with c1:
                field_chip("manual")
                manual["future_deposit_potential"] = st.selectbox(
                    "Future Deposit Potential",
                    POTENTIAL_OPTIONS,
                    index=POTENTIAL_OPTIONS.index(manual["future_deposit_potential"]),
                    key=f"{group['summary']['group_id']}_future_dep",
                )
                field_chip("manual")
                manual["future_property_acquisition"] = st.selectbox(
                    "Future Property Acquisition",
                    POTENTIAL_OPTIONS,
                    index=POTENTIAL_OPTIONS.index(manual["future_property_acquisition"]),
                    key=f"{group['summary']['group_id']}_future_prop",
                )
            with c2:
                field_chip("manual")
                manual["future_borrowing_potential"] = st.selectbox(
                    "Future Borrowing Potential",
                    POTENTIAL_OPTIONS,
                    index=POTENTIAL_OPTIONS.index(manual["future_borrowing_potential"]),
                    key=f"{group['summary']['group_id']}_future_borr",
                )
                field_chip("manual")
                manual["recent_update"] = st.selectbox(
                    "Recent Update",
                    UPDATE_OPTIONS,
                    index=UPDATE_OPTIONS.index(manual["recent_update"]),
                    key=f"{group['summary']['group_id']}_recent_update",
                )
            field_chip("manual")
            manual["overall_update"] = st.selectbox(
                "Overall Update",
                UPDATE_OPTIONS,
                index=UPDATE_OPTIONS.index(manual["overall_update"]),
                key=f"{group['summary']['group_id']}_overall_update",
            )


def render_loan_overview_block(loan):
    db = loan["db"]
    manual = loan["manual"]
    st.markdown("##### Loan Overview")
    left, right = st.columns(2)
    with left:
        with st.container(border=True):
            st.markdown("###### COGNOS Fields")
            cols = st.columns(2)
            for idx, (key, label) in enumerate(LOAN_COGNOS_FIELDS[:12]):
                value = db[key]
                if isinstance(value, date):
                    value = value.isoformat()
                elif isinstance(value, float) and ("rate" in key or "ltv" in key):
                    value = f"{value*100:.2f}%" if value < 1 and ("rate" in key or "ltv" in key) else money(value)
                elif isinstance(value, (int, float)) and "date" not in key and "number" not in key and "notes" not in key and "type" not in key and "location" not in key and "reason" not in key and "name" not in key and "recourse" not in key:
                    value = money(value)
                with cols[idx % 2]:
                    render_db_value(label, value, key=f"{loan['loan_id']}_overview_{key}")
    with right:
        with st.container(border=True):
            st.markdown("###### Manual Fields")
            c1, c2 = st.columns(2)
            with c1:
                field_chip("manual")
                manual["loan_rate_type"] = st.selectbox(
                    "Loan Rate Type", RATE_TYPES,
                    index=RATE_TYPES.index(manual["loan_rate_type"]),
                    key=f"{loan['loan_id']}_rate_type",
                )
                field_chip("manual")
                manual["loan_term_years"] = st.number_input(
                    "Loan Term (Years)", min_value=1, max_value=30,
                    value=int(manual["loan_term_years"]),
                    key=f"{loan['loan_id']}_loan_term_years",
                )
                field_chip("manual")
                manual["strength_bucket"] = st.selectbox(
                    "Primary Strength", ["Collateral Quality", "Guarantor Support", "Deposit Depth", "Cash Flow Stability", "Market Position"],
                    index=["Collateral Quality", "Guarantor Support", "Deposit Depth", "Cash Flow Stability", "Market Position"].index(manual["strength_bucket"]),
                    key=f"{loan['loan_id']}_strength_bucket",
                )
                field_chip("manual")
                manual["relationship_risk_assessment"] = st.selectbox(
                    "Relationship Risk Assessment",
                    ["Low", "Moderate", "Elevated"],
                    index=["Low", "Moderate", "Elevated"].index(manual["relationship_risk_assessment"]),
                    key=f"{loan['loan_id']}_relationship_risk",
                )
            with c2:
                field_chip("manual")
                manual["weakness_bucket"] = st.selectbox(
                    "Primary Weakness", ["Tenant Rollover", "Leverage", "Liquidity", "Construction Timing", "Concentration"],
                    index=["Tenant Rollover", "Leverage", "Liquidity", "Construction Timing", "Concentration"].index(manual["weakness_bucket"]),
                    key=f"{loan['loan_id']}_weakness_bucket",
                )
                field_chip("manual")
                manual["risk_rating_recommendation"] = st.selectbox(
                    "Risk Rating Recommendation",
                    RISK_RATINGS,
                    index=RISK_RATINGS.index(manual["risk_rating_recommendation"]),
                    key=f"{loan['loan_id']}_risk_reco",
                )
                field_chip("manual")
                manual["covenant_review_status"] = st.selectbox(
                    "Covenant Review", COVENANT_OPTIONS,
                    index=COVENANT_OPTIONS.index(manual["covenant_review_status"]),
                    key=f"{loan['loan_id']}_covenant_review",
                )
                field_chip("manual")
                manual["waiver_request"] = st.selectbox(
                    "Waiver Request", BOOLEAN_OPTIONS,
                    index=BOOLEAN_OPTIONS.index(manual["waiver_request"]),
                    key=f"{loan['loan_id']}_waiver_request",
                )
            field_chip("manual")
            manual["additional_repayment_sources"] = st.selectbox(
                "Additional Sources of Repayment",
                ["Guarantor Support", "Deposit Collateral", "Cross-Collateral", "Operating Cash Flow", "Other"],
                index=["Guarantor Support", "Deposit Collateral", "Cross-Collateral", "Operating Cash Flow", "Other"].index(manual["additional_repayment_sources"]),
                key=f"{loan['loan_id']}_repayment_sources",
            )


def render_rent_roll_section(loan):
    st.markdown("##### Rent Roll Analysis")
    st.caption("All rent roll fields are manual in this mockup. Add or remove rows as needed.")
    rr = loan["rent_roll"]
    rr_editor = st.data_editor(
        rr,
        num_rows="dynamic",
        use_container_width=True,
        key=f"{loan['loan_id']}_rent_roll_editor",
    )
    loan["rent_roll"] = rr_editor.copy()

    as_of = date.today()
    computed = compute_rent_roll(rr_editor, as_of)
    display_cols = [
        "suite","tenant","sf","tenant_pct_total_sf","tenant_since","lease_start","lease_end",
        "original_lease_term_years","remaining_term_years","monthly_rent_per_sf",
        "base_monthly_rent","monthly_cams","annual_rent","pct_total_rent","options"
    ]
    computed = computed[display_cols]
    with_footer = totals_row(computed)

    fmt_df = with_footer.copy()
    pct_cols = ["tenant_pct_total_sf", "pct_total_rent"]
    money_cols = ["base_monthly_rent", "monthly_cams", "annual_rent"]
    for c in pct_cols:
        fmt_df[c] = fmt_df[c].apply(lambda x: "" if pd.isna(x) or x == "" else f"{float(x):.2%}")
    for c in money_cols:
        fmt_df[c] = fmt_df[c].apply(lambda x: "" if pd.isna(x) or x == "" else f"${float(x):,.0f}")
    fmt_df["sf"] = fmt_df["sf"].apply(lambda x: "" if pd.isna(x) or x == "" else f"{float(x):,.0f}")
    for c in ["monthly_rent_per_sf","original_lease_term_years","remaining_term_years"]:
        fmt_df[c] = fmt_df[c].apply(lambda x: "" if pd.isna(x) or x == "" else f"{float(x):.2f}")

    left, right = st.columns([2.2, 1.2])
    with left:
        st.dataframe(fmt_df, use_container_width=True, hide_index=True)
    with right:
        les = lease_expiration_summary(computed)
        # Build a display copy as object dtype so formatted strings do not clash with numeric dtypes
        les_fmt = les.copy().astype(object)
        for y in ["Year 1", "Year 2", "Year 3"]:
            les_fmt.loc[les_fmt["Metric"] == "total %", y] = les_fmt.loc[les_fmt["Metric"] == "total %", y].apply(lambda x: f"{float(x):.2%}")
            les_fmt.loc[les_fmt["Metric"] == "sq ft", y] = les_fmt.loc[les_fmt["Metric"] == "sq ft", y].apply(lambda x: f"{float(x):,.0f}")
        vo = vacant_occupied_summary(computed)
        vo_fmt = vo.copy()
        vo_fmt["sq ft"] = vo_fmt["sq ft"].apply(lambda x: f"{float(x):,.0f}")
        vo_fmt["%"] = vo_fmt["%"].apply(lambda x: f"{float(x):.2%}")
        st.markdown("**Lease Expiration Summary**")
        st.dataframe(les_fmt, use_container_width=True, hide_index=True)
        st.markdown("**Vacancy / Occupancy Summary**")
        st.dataframe(vo_fmt, use_container_width=True, hide_index=True)
    return computed, with_footer, les, vo


def build_stress_state_from_loan(loan):
    db = loan["db"]
    manual = loan["stress_manual"]
    io, amort_years = parse_amortization(db["amortization_period_or_io"])
    return {
        "loan_source": "Loan Review",
        "loan_amount": float(db["current_loan_amount"]),
        "note_rate": float(db["note_rate"]),
        "io": io,
        "amort_years": amort_years,
        "debt_service_override_on": True,
        "annual_debt_service_override": float(db["annual_debt_service"]),
        "income_source": manual["income_source"],
        "rental_income": float(manual["rental_income"]),
        "opex_source": manual["opex_source"],
        "operating_expenses": float(manual["operating_expenses"]),
        "cap_source": manual["cap_source"],
        "cap_rate": float(manual["cap_rate"]),
        "target_dscr": float(manual["target_dscr"]),
        "target_ltv": float(manual["target_ltv"]),
    }


def render_stress_section(loan):
    st.markdown("##### Investor CRE Stress Test")
    manual = loan["stress_manual"]
    db = loan["db"]

    c1, c2, c3 = st.columns(3)
    with c1:
        render_db_value("Current Loan Amount", money(db["current_loan_amount"]), key=f"{loan['loan_id']}_stress_current_loan_amount")
        render_db_value("Note Rate", f"{db['note_rate']*100:.2f}%", key=f"{loan['loan_id']}_stress_note_rate")
    with c2:
        render_db_value("Amortization / I-O", db["amortization_period_or_io"], key=f"{loan['loan_id']}_stress_amortization")
        render_db_value("Annual Debt Service", money(db["annual_debt_service"]), key=f"{loan['loan_id']}_stress_annual_debt_service")
    with c3:
        field_chip("manual")
        manual["income_source"] = st.selectbox(
            "Rental Income Source", SOURCE_OPTIONS,
            index=SOURCE_OPTIONS.index(manual["income_source"]),
            key=f"{loan['loan_id']}_income_source",
        )
        field_chip("manual")
        manual["rental_income"] = st.number_input(
            "Rental Income ($/yr)", min_value=0.0, value=float(manual["rental_income"]), step=10000.0,
            key=f"{loan['loan_id']}_rental_income",
        )

    c4, c5, c6 = st.columns(3)
    with c4:
        field_chip("manual")
        manual["opex_source"] = st.selectbox(
            "Operating Expense Source", ["Tax Return", "P&L Statement", "Appraisal", "Proforma", "Other"],
            index=["Tax Return", "P&L Statement", "Appraisal", "Proforma", "Other"].index(manual["opex_source"]),
            key=f"{loan['loan_id']}_opex_source",
        )
        field_chip("manual")
        manual["operating_expenses"] = st.number_input(
            "Operating Expenses ($/yr)", min_value=0.0, value=float(manual["operating_expenses"]), step=10000.0,
            key=f"{loan['loan_id']}_operating_expenses",
        )
    with c5:
        field_chip("manual")
        manual["cap_source"] = st.selectbox(
            "Cap Rate Source", CAP_SOURCE_OPTIONS,
            index=CAP_SOURCE_OPTIONS.index(manual["cap_source"]),
            key=f"{loan['loan_id']}_cap_source",
        )
        field_chip("manual")
        manual["cap_rate"] = st.number_input(
            "Cap Rate (%)", min_value=0.01, value=float(manual["cap_rate"] * 100), step=0.05,
            key=f"{loan['loan_id']}_cap_rate",
        ) / 100.0
    with c6:
        field_chip("manual")
        manual["target_dscr"] = st.number_input(
            "Target DSCR", min_value=0.50, value=float(manual["target_dscr"]), step=0.05,
            key=f"{loan['loan_id']}_target_dscr",
        )
        field_chip("manual")
        manual["target_ltv"] = st.number_input(
            "Target LTV (%)", min_value=10.0, max_value=100.0, value=float(manual["target_ltv"] * 100), step=1.0,
            key=f"{loan['loan_id']}_target_ltv",
        ) / 100.0

    s = build_stress_state_from_loan(loan)
    noi = float(s["rental_income"] - s["operating_expenses"])
    debt_service_calc = annual_debt_service(s["loan_amount"], s["note_rate"], s["amort_years"], s["io"])
    debt_service = float(s["annual_debt_service_override"]) if s["debt_service_override_on"] else debt_service_calc
    value = (noi / s["cap_rate"]) if s["cap_rate"] > 0 else 0.0
    dscr = (noi / debt_service) if debt_service > 0 else 0.0
    ltv = (s["loan_amount"] / value) if value > 0 else 0.0

    entry_rows = collect_stress_entry(s, noi, debt_service, value, dscr, ltv)
    entry_df = pd.DataFrame(entry_rows)
    cap_df, ir_df, vac_df, noi_df = build_stress_tables(s, noi, debt_service)

    st.dataframe(entry_df, use_container_width=True, hide_index=True)

    a, b, c, d = st.columns(4)
    with a:
        st.markdown("**Cap Rate Effect**")
        st.dataframe(cap_df, use_container_width=True, hide_index=True)
    with b:
        st.markdown("**Interest Rate Effect**")
        st.dataframe(ir_df, use_container_width=True, hide_index=True)
    with c:
        st.markdown("**Vacancy Rate Effect**")
        st.dataframe(vac_df, use_container_width=True, hide_index=True)
    with d:
        st.markdown("**NOI Change Effect**")
        st.dataframe(noi_df, use_container_width=True, hide_index=True)

    payload = build_stress_test_ai_payload(s, noi, debt_service, value, dscr, ltv, cap_df, ir_df, vac_df, noi_df)
    commentary = rule_based_stress_test_bullets(payload)
    return {
        "state": s,
        "noi": noi,
        "debt_service": debt_service,
        "value": value,
        "dscr": dscr,
        "ltv": ltv,
        "entry_rows": entry_rows,
        "cap_df": cap_df,
        "ir_df": ir_df,
        "vac_df": vac_df,
        "noi_df": noi_df,
        "commentary": commentary,
    }


def render_rollover_section(loan, computed_rr, stress_state):
    st.markdown("##### Rollover Risk")
    manual = loan["rollover_manual"]
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        field_chip("manual")
        manual["leasing_commission_pct"] = st.number_input(
            "Leasing Commission (%)", min_value=0.0, max_value=25.0,
            value=float(manual["leasing_commission_pct"] * 100), step=0.25,
            key=f"{loan['loan_id']}_lc_pct",
        ) / 100.0
    with c2:
        field_chip("manual")
        manual["ti_per_sf"] = st.number_input(
            "TI ($/SF)", min_value=0.0, value=float(manual["ti_per_sf"]), step=1.0,
            key=f"{loan['loan_id']}_ti_psf",
        )
    with c3:
        field_chip("manual")
        manual["market_rent_per_sf_yr"] = st.number_input(
            "Market Rent ($/SF/Yr)", min_value=0.0, value=float(manual["market_rent_per_sf_yr"]), step=1.0,
            key=f"{loan['loan_id']}_market_rent",
        )
    with c4:
        field_chip("manual")
        manual["rent_loss_months"] = st.number_input(
            "Rent Loss (Months)", min_value=0.0, max_value=24.0, value=float(manual["rent_loss_months"]), step=0.5,
            key=f"{loan['loan_id']}_rent_loss_mos",
        )

    roll_df, payload = build_rollover_risk_outputs(
        borrower=loan["db"]["borrower_name"],
        address=loan["db"]["collateral_location"],
        as_of=date.today(),
        computed_rr=computed_rr,
        stress_state=stress_state,
        leasing_commission_pct=float(manual["leasing_commission_pct"]),
        ti_per_sf=float(manual["ti_per_sf"]),
        market_rent_per_sf_yr=float(manual["market_rent_per_sf_yr"]),
        rent_loss_months=float(manual["rent_loss_months"]),
    )

    disp = roll_df.copy()
    money_rows = {"NOI (from Stress Test)", "Leasing Commissions", "Tenant Improvements (TI)", "Rent Loss", "Total Estimated Costs", "Adjusted NOI", "Actual Debt Service"}
    dscr_rows = {"Adjusted DSCR"}
    for i, row in disp.iterrows():
        row_name = str(row["Row"])
        for c in ["Year 1", "Year 2", "Year 3"]:
            v = row[c]
            if row_name in money_rows:
                disp.at[i, c] = "" if v == "" or v == "—" else fmt_money(v)
            elif row_name in dscr_rows:
                disp.at[i, c] = "" if v == "" or v == "—" else f"{float(v):.2f}x"
    st.dataframe(disp, use_container_width=True, hide_index=True)
    commentary = rule_based_rollover_bullets(payload)
    return {"table_df": roll_df, "payload": payload, "commentary": commentary}


def render_construction_section(loan):
    st.markdown("##### Construction / Bridge Loan Stress Test")
    db = loan["db"]
    manual = loan["construction_manual"]
    left, right = st.columns(2)
    with left:
        render_db_value("Loan Commitment", money(db["loan_commitment"]), key=f"{loan['loan_id']}_construction_loan_commitment")
        field_chip("manual")
        manual["proforma_noi"] = st.number_input(
            "Pro Forma NOI ($/yr)", min_value=0.0, value=float(manual["proforma_noi"]), step=25000.0,
            key=f"{loan['loan_id']}_proforma_noi",
        )
        field_chip("manual")
        manual["appraisal_cap_rate"] = st.number_input(
            "Appraisal Cap Rate (%)", min_value=0.01, value=float(manual["appraisal_cap_rate"] * 100), step=0.05,
            key=f"{loan['loan_id']}_bridge_cap_rate",
        ) / 100.0
    with right:
        field_chip("manual")
        manual["takeout_amort_years"] = st.number_input(
            "Takeout Amortization (Years)", min_value=1.0, value=float(manual["takeout_amort_years"]), step=1.0,
            key=f"{loan['loan_id']}_takeout_amort",
        )
        field_chip("manual")
        manual["takeout_dscr"] = st.number_input(
            "Takeout DSCR", min_value=0.50, value=float(manual["takeout_dscr"]), step=0.05,
            key=f"{loan['loan_id']}_takeout_dscr",
        )
        field_chip("manual")
        manual["takeout_ltv"] = st.number_input(
            "Takeout LTV (%)", min_value=10.0, max_value=100.0, value=float(manual["takeout_ltv"] * 100), step=1.0,
            key=f"{loan['loan_id']}_takeout_ltv",
        ) / 100.0
        field_chip("manual")
        manual["underwriting_rate"] = st.number_input(
            "Underwriting Rate (%)", min_value=0.01, value=float(manual["underwriting_rate"] * 100), step=0.05,
            key=f"{loan['loan_id']}_underwriting_rate",
        ) / 100.0

    manual["grid_rows"] = int(st.number_input(
        "Rows in Stress Grids", min_value=3, max_value=10, value=int(manual["grid_rows"]), step=1,
        key=f"{loan['loan_id']}_grid_rows",
    ))

    scenarios = _noi_scenarios(float(manual["proforma_noi"]))
    cap_grid = _build_default_grid(float(manual["appraisal_cap_rate"]), 0.005, int(manual["grid_rows"]))
    ir_grid = _build_default_grid(float(manual["underwriting_rate"]), 0.005, int(manual["grid_rows"]))

    ltv_takeout = ((manual["proforma_noi"] / manual["appraisal_cap_rate"]) * manual["takeout_ltv"]) if manual["appraisal_cap_rate"] > 0 else 0.0
    dscr_takeout = max_loan_from_annual_debt_service(
        (manual["proforma_noi"] / manual["takeout_dscr"]) if manual["takeout_dscr"] > 0 else 0.0,
        manual["underwriting_rate"],
        manual["takeout_amort_years"],
    )

    constr_noi, constr_cap, constr_ir, constr_sale, constr_sizing = collect_construction_tables(
        {
            "loan_commitment": float(db["loan_commitment"]),
            "proforma_noi": float(manual["proforma_noi"]),
            "appraisal_cap_rate": float(manual["appraisal_cap_rate"]),
            "takeout_amort_years": float(manual["takeout_amort_years"]),
            "takeout_dscr": float(manual["takeout_dscr"]),
            "takeout_ltv": float(manual["takeout_ltv"]),
            "underwriting_rate": float(manual["underwriting_rate"]),
        },
        scenarios,
        cap_grid,
        ir_grid,
        ltv_takeout,
        dscr_takeout,
        max_loan_from_annual_debt_service,
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Max Takeout @ LTV", constr_sizing["max_ltv"])
    with c2:
        st.metric("Max Takeout @ DSCR", constr_sizing["max_dscr"])
    with c3:
        st.metric("Bridge Commitment", constr_sizing["bridge"])

    left, right = st.columns(2)
    with left:
        st.markdown("**Cap Rate → Max Takeout (LTV Constraint)**")
        st.dataframe(pd.DataFrame(constr_cap), use_container_width=True, hide_index=True)
        st.markdown("**Cap Rate → Sale Prices upon Stabilization**")
        st.dataframe(pd.DataFrame(constr_sale), use_container_width=True, hide_index=True)
    with right:
        st.markdown("**Interest Rate → Max Takeout (DSCR Constraint)**")
        st.dataframe(pd.DataFrame(constr_ir), use_container_width=True, hide_index=True)

    cap_plus_50 = float(manual["appraisal_cap_rate"]) + 0.005
    sale_cap_50 = (float(manual["proforma_noi"]) / cap_plus_50) if cap_plus_50 > 0 else 0.0
    bind_cap_50 = min(sale_cap_50 * float(manual["takeout_ltv"]), dscr_takeout)

    rate_plus_50 = float(manual["underwriting_rate"]) + 0.005
    bind_rate_50 = min(
        ltv_takeout,
        max_loan_from_annual_debt_service(
            (float(manual["proforma_noi"]) / float(manual["takeout_dscr"])) if float(manual["takeout_dscr"]) > 0 else 0.0,
            rate_plus_50,
            float(manual["takeout_amort_years"]),
        ),
    )

    payload = build_construction_ai_payload(
        bt={
            "loan_commitment": float(db["loan_commitment"]),
            "proforma_noi": float(manual["proforma_noi"]),
            "appraisal_cap_rate": float(manual["appraisal_cap_rate"]),
            "takeout_amort_years": float(manual["takeout_amort_years"]),
            "takeout_dscr": float(manual["takeout_dscr"]),
            "takeout_ltv": float(manual["takeout_ltv"]),
            "underwriting_rate": float(manual["underwriting_rate"]),
        },
        base_takeout_ltv=ltv_takeout,
        base_takeout_dscr=dscr_takeout,
        base_binding=min(ltv_takeout, dscr_takeout),
        stress_binding_20=min(ltv_takeout * 0.80, dscr_takeout * 0.80),
        stress_binding_30=min(ltv_takeout * 0.70, dscr_takeout * 0.70),
        cap_sens_binding={"cap_plus_50bp": {"cap_rate": cap_plus_50, "binding_takeout": bind_cap_50}},
        rate_sens_binding={"rate_plus_50bp": {"underwriting_rate": rate_plus_50, "binding_takeout": bind_rate_50}},
    )
    commentary = rule_based_construction_bullets(payload)
    return {
        "payload": payload,
        "commentary": commentary,
        "constr_noi": constr_noi,
        "constr_cap": constr_cap,
        "constr_ir": constr_ir,
        "constr_sale": constr_sale,
        "constr_sizing": constr_sizing,
    }


def compute_all_for_loan(loan):
    rr_comp, rr_with_footer, les, vo = render_rent_roll_section(loan)
    stress = render_stress_section(loan)
    rollover = render_rollover_section(loan, rr_comp, stress["state"])
    construction = render_construction_section(loan)
    return {
        "rr_comp": rr_comp,
        "rr_with_footer": rr_with_footer,
        "les": les,
        "vo": vo,
        "stress": stress,
        "rollover": rollover,
        "construction": construction,
    }


def derive_default_exec_summary(group, loan, analysis):
    summary = group["summary"]
    stress = analysis["stress"]
    return (
        f"{summary['group_name']} ({summary['group_id']}) remains a {summary['tier']} relationship managed by {summary['officer']} out of {summary['branch']}. "
        f"Selected loan {loan['loan_id']} for {loan['db']['borrower_name']} has a current balance of {money(loan['db']['loan_balance'])} "
        f"and matures on {fmt_date(loan['db']['loan_maturity_date'])}. "
        f"Current underwriting indicates DSCR of {stress['dscr']:.2f}x and LTV of {stress['ltv']:.2%}. "
        f"Key strength is {loan['manual']['strength_bucket'].lower()}, while the main watch item is {loan['manual']['weakness_bucket'].lower()}."
    )


def derive_primary_repayment(loan, analysis):
    rr = analysis["rr_comp"]
    total_rent = rr["annual_rent"].sum() if "annual_rent" in rr else 0.0
    occupied_sf = rr.loc[~rr["tenant"].astype(str).str.lower().str.contains("vacant"), "sf"].sum() if len(rr) else 0.0
    total_sf = rr["sf"].sum() if "sf" in rr else 0.0
    occ_pct = (occupied_sf / total_sf) if total_sf else 0.0
    return (
        f"- Primary Source: Stabilized property cash flow supports debt service.\n"
        f"- Annual Rent Base: {money(total_rent)} with occupancy of {occ_pct:.1%}.\n"
        f"- Secondary Support: {loan['manual']['additional_repayment_sources']}."
    )


def validation_issues(group, loan):
    issues = []
    m = loan["manual"]
    required_loan = ["loan_rate_type", "loan_term_years", "risk_rating_recommendation", "relationship_risk_assessment", "covenant_review_status"]
    for field in required_loan:
        if not str(m.get(field, "")).strip():
            issues.append(f"Missing loan field: {field}")
    sm = loan["stress_manual"]
    for field in ["rental_income", "operating_expenses", "cap_rate"]:
        if float(sm.get(field, 0) or 0) <= 0:
            issues.append(f"Missing / invalid stress input: {field}")
    if loan["rent_roll"].empty:
        issues.append("Rent roll must include at least one row.")
    else:
        valid_rows = loan["rent_roll"].fillna("")
        if not ((valid_rows["tenant"].astype(str).str.strip() != "") & (pd.to_numeric(valid_rows["sf"], errors="coerce").fillna(0) > 0)).any():
            issues.append("Rent roll needs at least one tenant row with square footage.")
    return issues


def render_final_review_and_export(group):
    st.divider()
    st.markdown("## Final Review and Word Export")
    st.caption("Review each loan, edit commentary only when needed, check the loan-level review box, and export one combined relationship document.")

    all_issues = []
    loan_tabs = st.tabs([loan["loan_id"] for loan in group["loans"]])
    relationship_exec_key = f"{group['summary']['group_id']}_relationship_exec_summary"
    if relationship_exec_key not in st.session_state:
        st.session_state[relationship_exec_key] = ""

    default_exec = " ".join([
        derive_default_exec_summary(group, loan, compute_analysis_snapshot(loan)) for loan in group["loans"]
    ])
    if not st.session_state[relationship_exec_key]:
        st.session_state[relationship_exec_key] = default_exec

    st.session_state[relationship_exec_key] = st.text_area(
        "Executive Relationship Summary",
        value=st.session_state[relationship_exec_key],
        height=140,
        key=f"{group['summary']['group_id']}_relationship_exec_summary_box",
    )

    for tab, loan in zip(loan_tabs, group["loans"]):
        with tab:
            analysis = compute_analysis_snapshot(loan)
            final = loan["final"]
            if not final.get("primary_repayment"):
                final["primary_repayment"] = derive_primary_repayment(loan, analysis)
            if not final.get("rent_roll_commentary"):
                final["rent_roll_commentary"] = analysis["rent_roll_commentary"]
            if not final.get("stress_commentary"):
                final["stress_commentary"] = analysis["stress_commentary"]
            if not final.get("rollover_commentary"):
                final["rollover_commentary"] = analysis["rollover_commentary"]
            if not final.get("construction_commentary"):
                final["construction_commentary"] = analysis["construction_commentary"]

            st.markdown(f"### {loan['loan_id']} · {loan['db']['borrower_name']}")
            meta_cols = st.columns(4)
            meta_cols[0].metric("Balance", money(loan["db"]["loan_balance"]))
            meta_cols[1].metric("Maturity", fmt_date(loan["db"]["loan_maturity_date"]))
            meta_cols[2].metric("Risk Rating", str(loan["db"]["current_risk_rating"]))
            meta_cols[3].metric("Collateral", str(loan["db"]["collateral_type"]))

            render_editable_commentary_block(
                "Primary Source of Repayment", final, "primary_repayment", loan["loan_id"], "Primary Source of Repayment", height=120
            )
            c1, c2 = st.columns(2)
            with c1:
                render_editable_commentary_block(
                    "Rent Roll Commentary", final, "rent_roll_commentary", loan["loan_id"], "Rent Roll", height=220
                )
                render_editable_commentary_block(
                    "Rollover Commentary", final, "rollover_commentary", loan["loan_id"], "Rollover", height=220
                )
            with c2:
                render_editable_commentary_block(
                    "Stress Test Commentary", final, "stress_commentary", loan["loan_id"], "Stress Test", height=220
                )
                render_editable_commentary_block(
                    "Construction Commentary", final, "construction_commentary", loan["loan_id"], "Construction", height=220
                )

            issues = validation_issues(group, loan)
            all_issues.extend([f"{loan['loan_id']}: {issue}" for issue in issues])
            if issues:
                st.warning("Please complete these items before export:")
                for issue in issues:
                    st.write(f"- {issue}")

            loan["final"]["validated"] = st.checkbox(
                "I checked rent roll, stress, rollover, construction, and commentary for this loan",
                key=f"{loan['loan_id']}_checked_all",
            )

    if all_issues:
        st.warning("Some loans still have missing required inputs. Please resolve them before export.")

    all_loan_checks = all(loan["final"].get("validated", False) for loan in group["loans"])
    can_export = (not all_issues) and all_loan_checks

    if can_export:
        doc_path = build_relationship_word_report(group, st.session_state[relationship_exec_key])
        with open(doc_path, "rb") as f:
            st.download_button(
                "Download Combined Word Document",
                data=f,
                file_name=f"{group['summary']['group_id']}_relationship_rbr.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
            )
    else:
        st.button("Download Combined Word Document", disabled=True, use_container_width=True)


def compute_analysis_snapshot(loan):
    rr_base = loan["rent_roll"]
    rr_comp = compute_rent_roll(rr_base, date.today())
    rr_with_footer = totals_row(rr_comp[[
        "suite","tenant","sf","tenant_pct_total_sf","tenant_since","lease_start","lease_end",
        "original_lease_term_years","remaining_term_years","monthly_rent_per_sf",
        "base_monthly_rent","monthly_cams","annual_rent","pct_total_rent","options"
    ]])
    les = lease_expiration_summary(rr_comp)
    vo = vacant_occupied_summary(rr_comp)
    rent_roll_commentary = rule_based_bullets(rr_comp, loan["db"]["borrower_name"], loan["db"]["collateral_location"], date.today())

    s = build_stress_state_from_loan(loan)
    noi = float(s["rental_income"] - s["operating_expenses"])
    debt_service_calc = annual_debt_service(s["loan_amount"], s["note_rate"], s["amort_years"], s["io"])
    debt_service = float(s["annual_debt_service_override"]) if s["debt_service_override_on"] else debt_service_calc
    value = (noi / s["cap_rate"]) if s["cap_rate"] > 0 else 0.0
    dscr = (noi / debt_service) if debt_service > 0 else 0.0
    ltv = (s["loan_amount"] / value) if value > 0 else 0.0
    entry_rows = collect_stress_entry(s, noi, debt_service, value, dscr, ltv)
    cap_df, ir_df, vac_df, noi_df = build_stress_tables(s, noi, debt_service)
    stress_payload = build_stress_test_ai_payload(s, noi, debt_service, value, dscr, ltv, cap_df, ir_df, vac_df, noi_df)
    stress_commentary = rule_based_stress_test_bullets(stress_payload)

    roll_df, roll_payload = build_rollover_risk_outputs(
        borrower=loan["db"]["borrower_name"],
        address=loan["db"]["collateral_location"],
        as_of=date.today(),
        computed_rr=rr_comp,
        stress_state=s,
        leasing_commission_pct=float(loan["rollover_manual"]["leasing_commission_pct"]),
        ti_per_sf=float(loan["rollover_manual"]["ti_per_sf"]),
        market_rent_per_sf_yr=float(loan["rollover_manual"]["market_rent_per_sf_yr"]),
        rent_loss_months=float(loan["rollover_manual"]["rent_loss_months"]),
    )
    rollover_commentary = rule_based_rollover_bullets(roll_payload)

    constr = loan["construction_manual"]
    scenarios = _noi_scenarios(float(constr["proforma_noi"]))
    cap_grid = _build_default_grid(float(constr["appraisal_cap_rate"]), 0.005, int(constr["grid_rows"]))
    ir_grid = _build_default_grid(float(constr["underwriting_rate"]), 0.005, int(constr["grid_rows"]))
    ltv_takeout = ((constr["proforma_noi"] / constr["appraisal_cap_rate"]) * constr["takeout_ltv"]) if constr["appraisal_cap_rate"] > 0 else 0.0
    dscr_takeout = max_loan_from_annual_debt_service(
        (constr["proforma_noi"] / constr["takeout_dscr"]) if constr["takeout_dscr"] > 0 else 0.0,
        constr["underwriting_rate"],
        constr["takeout_amort_years"],
    )
    constr_noi, constr_cap, constr_ir, constr_sale, constr_sizing = collect_construction_tables(
        {
            "loan_commitment": float(loan["db"]["loan_commitment"]),
            "proforma_noi": float(constr["proforma_noi"]),
            "appraisal_cap_rate": float(constr["appraisal_cap_rate"]),
            "takeout_amort_years": float(constr["takeout_amort_years"]),
            "takeout_dscr": float(constr["takeout_dscr"]),
            "takeout_ltv": float(constr["takeout_ltv"]),
            "underwriting_rate": float(constr["underwriting_rate"]),
        },
        scenarios,
        cap_grid,
        ir_grid,
        ltv_takeout,
        dscr_takeout,
        max_loan_from_annual_debt_service,
    )
    cap_plus_50 = float(constr["appraisal_cap_rate"]) + 0.005
    rate_plus_50 = float(constr["underwriting_rate"]) + 0.005
    constr_payload = build_construction_ai_payload(
        bt={
            "loan_commitment": float(loan["db"]["loan_commitment"]),
            "proforma_noi": float(constr["proforma_noi"]),
            "appraisal_cap_rate": float(constr["appraisal_cap_rate"]),
            "takeout_amort_years": float(constr["takeout_amort_years"]),
            "takeout_dscr": float(constr["takeout_dscr"]),
            "takeout_ltv": float(constr["takeout_ltv"]),
            "underwriting_rate": float(constr["underwriting_rate"]),
        },
        base_takeout_ltv=ltv_takeout,
        base_takeout_dscr=dscr_takeout,
        base_binding=min(ltv_takeout, dscr_takeout),
        stress_binding_20=min(ltv_takeout * 0.80, dscr_takeout * 0.80),
        stress_binding_30=min(ltv_takeout * 0.70, dscr_takeout * 0.70),
        cap_sens_binding={"cap_plus_50bp": {"cap_rate": cap_plus_50, "binding_takeout": min(((constr["proforma_noi"] / cap_plus_50) * constr["takeout_ltv"]) if cap_plus_50 > 0 else 0.0, dscr_takeout)}},
        rate_sens_binding={"rate_plus_50bp": {"underwriting_rate": rate_plus_50, "binding_takeout": min(ltv_takeout, max_loan_from_annual_debt_service((constr["proforma_noi"] / constr["takeout_dscr"]) if constr["takeout_dscr"] > 0 else 0.0, rate_plus_50, constr["takeout_amort_years"]))}},
    )
    construction_commentary = rule_based_construction_bullets(constr_payload)

    return {
        "rr_comp": rr_comp,
        "rr_with_footer": rr_with_footer,
        "les": les,
        "vo": vo,
        "rent_roll_commentary": rent_roll_commentary,
        "stress": {
            "state": s,
            "noi": noi,
            "debt_service": debt_service,
            "value": value,
            "dscr": dscr,
            "ltv": ltv,
        },
        "stress_entry_rows": entry_rows,
        "stress_cap": cap_df,
        "stress_ir": ir_df,
        "stress_vac": vac_df,
        "stress_noi": noi_df,
        "stress_commentary": stress_commentary,
        "roll_df": roll_df,
        "rollover_commentary": rollover_commentary,
        "constr_noi": constr_noi,
        "constr_cap": constr_cap,
        "constr_ir": constr_ir,
        "constr_sale": constr_sale,
        "constr_sizing": constr_sizing,
        "construction_commentary": construction_commentary,
    }


def build_loan_export_payload(group, loan, analysis):
    final = loan["final"]
    rr_table = collect_rent_roll_table(analysis["rr_with_footer"])
    lease_table = collect_lease_expiration(analysis["les"])
    vac_occ = collect_vac_occ(analysis["vo"])
    sc_data, si_data, sv_data, sn_data = collect_stress_tables(
        analysis["stress_cap"], analysis["stress_ir"], analysis["stress_vac"], analysis["stress_noi"]
    )
    ro_data = collect_rollover_table(analysis["roll_df"])
    return {
        "loan_title": f"{loan['loan_id']} · {loan['db']['borrower_name']}",
        "loan_meta": {
            "Loan ID": loan["loan_id"],
            "Loan Number": str(loan["db"]["loan_number"]),
            "Borrower": loan["db"]["borrower_name"],
            "Loan Type": loan["db"]["loan_type"],
            "Loan Balance": money(loan["db"]["loan_balance"]),
            "Maturity": fmt_date(loan["db"]["loan_maturity_date"]),
            "Risk Rating": loan["db"]["current_risk_rating"],
            "Collateral": loan["db"]["collateral_type"],
        },
        "primary_repayment": final["primary_repayment"],
        "rent_roll_rule_based": final["rent_roll_commentary"],
        "rent_roll_genai": "",
        "stress_test_rule_based": final["stress_commentary"],
        "stress_test_genai": "",
        "dscr_calculations": "",
        "rollover_rule_based": final["rollover_commentary"],
        "rollover_genai": "",
        "construction_rule_based": final["construction_commentary"],
        "construction_genai": "",
        "rent_roll_table": rr_table,
        "lease_exp_table": lease_table,
        "vac_occ_table": vac_occ,
        "stress_entry_table": analysis["stress_entry_rows"],
        "stress_cap_table": sc_data,
        "stress_ir_table": si_data,
        "stress_vac_table": sv_data,
        "stress_noi_table": sn_data,
        "rollover_table": ro_data,
        "constr_noi_scenarios": analysis["constr_noi"],
        "constr_cap_ltv_table": analysis["constr_cap"],
        "constr_ir_dscr_table": analysis["constr_ir"],
        "constr_sale_table": analysis["constr_sale"],
        "constr_sizing": analysis["constr_sizing"],
    }


def build_relationship_word_report(group, relationship_exec_summary):
    loan_payloads = []
    for loan in group["loans"]:
        analysis = compute_analysis_snapshot(loan)
        final = loan["final"]
        if not final["primary_repayment"]:
            final["primary_repayment"] = markdownish_to_plain_text(derive_primary_repayment(loan, analysis))
        if not final["rent_roll_commentary"]:
            final["rent_roll_commentary"] = markdownish_to_plain_text(analysis["rent_roll_commentary"])
        if not final["stress_commentary"]:
            final["stress_commentary"] = markdownish_to_plain_text(analysis["stress_commentary"])
        if not final["rollover_commentary"]:
            final["rollover_commentary"] = markdownish_to_plain_text(analysis["rollover_commentary"])
        if not final["construction_commentary"]:
            final["construction_commentary"] = markdownish_to_plain_text(analysis["construction_commentary"])
        loan_payloads.append(build_loan_export_payload(group, loan, analysis))

    doc_path = str(Path(tempfile.gettempdir()) / f"{group['summary']['group_id']}_relationship_rbr.docx")
    return create_credit_analysis_report(
        borrower_name=group["summary"]["group_name"],
        client_number=group["summary"]["group_id"],
        branch_number=group["summary"]["branch"],
        relationship_manager=group["summary"]["officer"],
        loan_exposure=group["common_db"]["loan_exposure"],
        deposit_relationship=group["common_db"]["deposit_relationship"],
        tier_level=group["common_db"]["tier_level"],
        exec_summary=relationship_exec_summary,
        loans_data=loan_payloads,
        output_path=doc_path,
    )


def render_loan_workspace(loan):
    render_loan_overview_block(loan)
    rr_comp, rr_with_footer, les, vo = render_rent_roll_section(loan)
    stress = render_stress_section(loan)
    render_rollover_section(loan, rr_comp, stress["state"])
    render_construction_section(loan)


def render_mock_one(group):
    st.markdown("## Mock 1 · Loan Tabs")
    tabs = st.tabs([loan["loan_id"] for loan in group["loans"]])
    for tab, loan in zip(tabs, group["loans"]):
        with tab:
            render_loan_workspace(loan)


def open_loan_details_dialog(loan):
    if hasattr(st, "dialog"):
        @st.dialog(f"{loan['loan_id']} · Full Loan Details")
        def _loan_dialog():
            st.markdown("### Additional COGNOS Fields")
            rows = []
            shown = {"borrower_name", "loan_number", "loan_type", "loan_balance", "loan_maturity_date", "current_risk_rating", "watch_reason", "collateral_type"}
            for key, label in LOAN_COGNOS_FIELDS:
                if key in shown:
                    continue
                val = loan["db"][key]
                if isinstance(val, date):
                    val = val.isoformat()
                rows.append({"Field": label, "Value": val})
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        _loan_dialog()
    else:
        with st.expander(f"{loan['loan_id']} · Full Loan Details", expanded=True):
            rows = [{"Field": label, "Value": loan["db"][key]} for key, label in LOAN_COGNOS_FIELDS]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def render_mock_two(group):
    st.markdown("## Mock 2 · Consolidated Loan Table")
    st.caption("Important COGNOS columns stay locked. Important manual inputs stay in the table. Additional data can be opened per loan.")
    records = []
    for loan in group["loans"]:
        records.append({
            "loan_id": loan["loan_id"],
            "borrower_name": loan["db"]["borrower_name"],
            "loan_type": loan["db"]["loan_type"],
            "loan_balance": loan["db"]["loan_balance"],
            "loan_maturity_date": fmt_date(loan["db"]["loan_maturity_date"]),
            "current_risk_rating": loan["db"]["current_risk_rating"],
            "watch_reason": loan["db"]["watch_reason"],
            "loan_rate_type": loan["manual"]["loan_rate_type"],
            "loan_term_years": loan["manual"]["loan_term_years"],
            "rental_income": loan["stress_manual"]["rental_income"],
            "operating_expenses": loan["stress_manual"]["operating_expenses"],
            "cap_rate_pct": loan["stress_manual"]["cap_rate"] * 100,
            "risk_rating_recommendation": loan["manual"]["risk_rating_recommendation"],
        })
    df = pd.DataFrame(records)
    edited = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        disabled=["loan_id", "borrower_name", "loan_type", "loan_balance", "loan_maturity_date", "current_risk_rating", "watch_reason"],
        key=f"{group['summary']['group_id']}_summary_editor",
        column_config={
            "loan_balance": st.column_config.NumberColumn("Loan Balance", format="$%0.0f"),
            "rental_income": st.column_config.NumberColumn("Rental Income", format="$%0.0f"),
            "operating_expenses": st.column_config.NumberColumn("Operating Expenses", format="$%0.0f"),
            "cap_rate_pct": st.column_config.NumberColumn("Cap Rate (%)", format="%.2f"),
        },
    )
    for _, row in edited.iterrows():
        loan = get_loan(group, row["loan_id"])
        loan["manual"]["loan_rate_type"] = row["loan_rate_type"]
        loan["manual"]["loan_term_years"] = int(row["loan_term_years"])
        loan["stress_manual"]["rental_income"] = float(row["rental_income"])
        loan["stress_manual"]["operating_expenses"] = float(row["operating_expenses"])
        loan["stress_manual"]["cap_rate"] = float(row["cap_rate_pct"]) / 100.0
        loan["manual"]["risk_rating_recommendation"] = row["risk_rating_recommendation"]

    loan_ids = [loan["loan_id"] for loan in group["loans"]]
    selected = st.selectbox("Open detailed loan workspace", loan_ids, key=f"{group['summary']['group_id']}_detail_selector")
    detail_loan = get_loan(group, selected)
    open_col, _ = st.columns([1, 3])
    with open_col:
        if st.button("Open full loan details", key=f"{selected}_open_dialog", use_container_width=True):
            open_loan_details_dialog(detail_loan)
    with st.expander(f"{selected} · Detailed analysis workspace", expanded=True):
        render_loan_workspace(detail_loan)


def render_home():
    st.title(APP_TITLE)
    st.caption("Synthetic relationship data for UI mockups. Calculations and Word export are reused from the uploaded V1 app.")
    portfolio = st.session_state.portfolio
    rows = []
    for group in portfolio.values():
        s = group["summary"]
        rows.append({
            "group #": s["group_id"],
            "group name": s["group_name"],
            "branch": s["branch"],
            "officer": s["officer"],
            "active loan account": s["active_loan_account"],
            "loan balances": s["loan_balances"],
            "loan commitments": s["loan_commitments"],
            "deposit balances": s["deposit_balances"],
            "tier": s["tier"],
            "last review date": fmt_date(s["last_review_date"]),
            "next review date": fmt_date(s["next_review_date"]),
            "current rbr status": s["current_rbr_status"],
        })
    df = pd.DataFrame(rows)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Group Count", len(df))
    c2.metric("Loan Count", int(df["active loan account"].sum()))
    c3.metric("Total Balances", money(df["loan balances"].sum()))
    c4.metric("Total Commitments", money(df["loan commitments"].sum()))

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "loan balances": st.column_config.NumberColumn("loan balances", format="$%0.0f"),
            "loan commitments": st.column_config.NumberColumn("loan commitments", format="$%0.0f"),
            "deposit balances": st.column_config.NumberColumn("deposit balances", format="$%0.0f"),
        },
    )
    options = [f"{row['group #']} · {row['group name']}" for _, row in df.iterrows()]
    selected = st.selectbox("Open client group", options, key="home_group_select")
    if st.button("Open RBR Page", type="primary", use_container_width=True):
        st.session_state.selected_group_id = selected.split(" · ")[0]
        st.session_state.page = "RBR"
        st.rerun()


def render_rbr_page():
    group = get_group()
    if group is None:
        st.session_state.page = "Home"
        st.rerun()
    top_left, top_right = st.columns([3, 1])
    with top_left:
        if st.button("← Back to Homepage"):
            st.session_state.page = "Home"
            st.rerun()
    with top_right:
        st.caption("Database fields are locked. Manual fields are editable.")
    render_group_header(group)
    render_common_sections(group)
    st.divider()
    mock_mode = st.radio("Second-Page Mockup", ["Mock 1 · Loan Tabs", "Mock 2 · Consolidated Table"], horizontal=True)
    if mock_mode.startswith("Mock 1"):
        render_mock_one(group)
    else:
        render_mock_two(group)
    render_final_review_and_export(group)


def main():
    inject_css()
    init_state()
    if st.session_state.page == "Home":
        render_home()
    else:
        render_rbr_page()


if __name__ == "__main__":
    main()
