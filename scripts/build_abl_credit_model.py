#!/usr/bin/env python3
"""Build the ABL credit model workbook and memo for Genuine Parts Company."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell


REPO_ROOT = Path(__file__).resolve().parents[1]
DATA_PATH = REPO_ROOT / "data" / "gpc_sec_financials.json"
THREAD_ID = os.environ.get("CODEX_THREAD_ID", "local")
OUTPUT_DIR = REPO_ROOT / "outputs" / THREAD_ID
WORKBOOK_PATH = OUTPUT_DIR / "GPC_ABL_Borrowing_Base_Credit_Model.xlsx"
MEMO_PATH = OUTPUT_DIR / "GPC_ABL_Credit_Memo.md"

YEARS = [2021, 2022, 2023, 2024, 2025]
SCENARIOS = ["Base", "Downside", "Stress"]


def load_data() -> dict[str, Any]:
    return json.loads(DATA_PATH.read_text(encoding="utf-8"))


def metric(data: dict[str, Any], key: str, year: int, *, fallback_key: str | None = None) -> float:
    raw = data["metrics"][key]["values"][str(year)]["value"]
    if raw is None and fallback_key:
        raw = data["metrics"][fallback_key]["values"][str(year)]["value"]
    return float(raw or 0) / 1_000_000


def source_for(data: dict[str, Any], key: str, year: int, *, fallback_key: str | None = None) -> dict[str, Any] | None:
    item = data["metrics"][key]["values"][str(year)]
    if item["value"] is None and fallback_key:
        item = data["metrics"][fallback_key]["values"][str(year)]
    return item.get("source")


def fmt_money(value: float) -> str:
    return f"${value:,.0f}mm"


def fmt_pct(value: float) -> str:
    return f"{value:.1%}"


def hist_values(data: dict[str, Any]) -> dict[int, dict[str, float]]:
    values: dict[int, dict[str, float]] = {}
    for year in YEARS:
        interest = metric(data, "interest_expense", year, fallback_key="interest_paid")
        pretax = metric(data, "pretax_income", year)
        da = metric(data, "depreciation_amortization", year)
        revenue = metric(data, "revenue", year)
        cfo = metric(data, "cfo", year)
        capex = metric(data, "capex", year)
        cash = metric(data, "cash", year)
        ar = metric(data, "accounts_receivable", year)
        inv = metric(data, "inventory", year)
        current_assets = metric(data, "current_assets", year)
        current_liabilities = metric(data, "current_liabilities", year)
        st_borrowings = metric(data, "short_term_borrowings", year)
        cp = metric(data, "commercial_paper", year)
        ltd_current = metric(data, "long_term_debt_current", year)
        ltd_noncurrent = metric(data, "long_term_debt_noncurrent", year)
        total_debt = st_borrowings + cp + ltd_current + ltd_noncurrent
        ebitda = pretax + interest + da
        values[year] = {
            "revenue": revenue,
            "pretax_income": pretax,
            "income_tax": metric(data, "income_tax", year),
            "net_income": metric(data, "net_income", year),
            "interest": interest,
            "da": da,
            "ebitda": ebitda,
            "ebitda_margin": ebitda / revenue if revenue else 0,
            "cfo": cfo,
            "capex": capex,
            "fcf": cfo - capex,
            "ar": ar,
            "inventory": inv,
            "cash": cash,
            "current_assets": current_assets,
            "current_liabilities": current_liabilities,
            "st_borrowings": st_borrowings,
            "commercial_paper": cp,
            "ltd_current": ltd_current,
            "ltd_noncurrent": ltd_noncurrent,
            "total_debt": total_debt,
            "net_debt": total_debt - cash,
            "current_ratio": current_assets / current_liabilities if current_liabilities else 0,
            "quick_ratio": (cash + ar) / current_liabilities if current_liabilities else 0,
            "dso": ar / revenue * 365 if revenue else 0,
            "inventory_days": inv / revenue * 365 if revenue else 0,
            "current_liability_days": current_liabilities / revenue * 365 if revenue else 0,
        }
        values[year]["ccc_proxy"] = (
            values[year]["dso"] + values[year]["inventory_days"] - values[year]["current_liability_days"]
        )
        values[year]["debt_to_ebitda"] = total_debt / ebitda if ebitda else 0
        values[year]["interest_coverage"] = ebitda / interest if interest else 0
        if year == YEARS[0]:
            values[year]["revenue_growth"] = 0
        else:
            prior_revenue = values[year - 1]["revenue"]
            values[year]["revenue_growth"] = revenue / prior_revenue - 1 if prior_revenue else 0
    return values


def scenario_inputs() -> dict[str, dict[str, float]]:
    return {
        "Base": {
            "revenue_growth": 0.02,
            "ebitda_margin_change": 0.000,
            "ar_eligible_pct": 0.92,
            "ar_advance_rate": 0.85,
            "inventory_balance_change": 0.02,
            "inventory_eligible_pct": 0.75,
            "inventory_advance_rate": 0.50,
            "dilution_reserve_pct": 0.020,
            "concentration_reserve_pct": 0.015,
            "other_reserves": 150.0,
            "availability_block": 100.0,
            "loan_balance_change": 0.00,
            "interest_rate_uplift": 0.000,
            "cfo_conversion": 0.70,
            "capex_pct_revenue": 0.020,
        },
        "Downside": {
            "revenue_growth": -0.08,
            "ebitda_margin_change": -0.015,
            "ar_eligible_pct": 0.87,
            "ar_advance_rate": 0.825,
            "inventory_balance_change": 0.08,
            "inventory_eligible_pct": 0.65,
            "inventory_advance_rate": 0.45,
            "dilution_reserve_pct": 0.035,
            "concentration_reserve_pct": 0.030,
            "other_reserves": 250.0,
            "availability_block": 150.0,
            "loan_balance_change": 0.10,
            "interest_rate_uplift": 0.010,
            "cfo_conversion": 0.55,
            "capex_pct_revenue": 0.022,
        },
        "Stress": {
            "revenue_growth": -0.18,
            "ebitda_margin_change": -0.030,
            "ar_eligible_pct": 0.80,
            "ar_advance_rate": 0.80,
            "inventory_balance_change": 0.12,
            "inventory_eligible_pct": 0.55,
            "inventory_advance_rate": 0.40,
            "dilution_reserve_pct": 0.050,
            "concentration_reserve_pct": 0.050,
            "other_reserves": 350.0,
            "availability_block": 200.0,
            "loan_balance_change": 0.20,
            "interest_rate_uplift": 0.025,
            "cfo_conversion": 0.35,
            "capex_pct_revenue": 0.025,
        },
    }


def scenario_outputs(hist: dict[int, dict[str, float]], inputs: dict[str, dict[str, float]]) -> dict[str, dict[str, Any]]:
    base = hist[2025]
    loan_base = base["st_borrowings"] + base["commercial_paper"] + base["ltd_current"]
    outputs: dict[str, dict[str, Any]] = {}
    for scenario, assumption in inputs.items():
        gross_ar = base["ar"] * (1 + assumption["revenue_growth"])
        eligible_ar = gross_ar * assumption["ar_eligible_pct"]
        ar_avail = eligible_ar * assumption["ar_advance_rate"]
        gross_inv = base["inventory"] * (1 + assumption["inventory_balance_change"])
        eligible_inv = gross_inv * assumption["inventory_eligible_pct"]
        inv_avail = eligible_inv * assumption["inventory_advance_rate"]
        dilution_reserve = gross_ar * assumption["dilution_reserve_pct"]
        concentration_reserve = gross_ar * assumption["concentration_reserve_pct"]
        total_reserves = (
            dilution_reserve
            + concentration_reserve
            + assumption["other_reserves"]
            + assumption["availability_block"]
        )
        borrowing_base = ar_avail + inv_avail - total_reserves
        loan_balance = loan_base * (1 + assumption["loan_balance_change"])
        excess_availability = borrowing_base - loan_balance
        revenue = base["revenue"] * (1 + assumption["revenue_growth"])
        ebitda_margin = max(base["ebitda_margin"] + assumption["ebitda_margin_change"], 0)
        ebitda = revenue * ebitda_margin
        interest = base["interest"] + loan_balance * assumption["interest_rate_uplift"]
        cfo = ebitda * assumption["cfo_conversion"]
        capex = revenue * assumption["capex_pct_revenue"]
        fcf = cfo - capex
        total_debt = base["total_debt"] + (loan_balance - loan_base)
        min_liquidity = base["cash"] + excess_availability
        debt_to_ebitda = total_debt / ebitda if ebitda else 99
        coverage = ebitda / interest if interest else 99
        if excess_availability < 0 or min_liquidity < 0 or coverage < 1.0 or debt_to_ebitda > 8.0:
            rating = "High Risk"
        elif excess_availability < 500 or coverage < 2.0 or debt_to_ebitda > 5.0:
            rating = "Watch"
        else:
            rating = "Pass"
        outputs[scenario] = {
            "gross_ar": gross_ar,
            "eligible_ar": eligible_ar,
            "ar_availability": ar_avail,
            "gross_inventory": gross_inv,
            "eligible_inventory": eligible_inv,
            "inventory_availability": inv_avail,
            "dilution_reserve": dilution_reserve,
            "concentration_reserve": concentration_reserve,
            "total_reserves": total_reserves,
            "borrowing_base": borrowing_base,
            "loan_balance": loan_balance,
            "excess_availability": excess_availability,
            "utilization": loan_balance / borrowing_base if borrowing_base else 0,
            "revenue": revenue,
            "ebitda_margin": ebitda_margin,
            "ebitda": ebitda,
            "interest": interest,
            "coverage": coverage,
            "total_debt": total_debt,
            "debt_to_ebitda": debt_to_ebitda,
            "cfo": cfo,
            "capex": capex,
            "fcf": fcf,
            "min_liquidity": min_liquidity,
            "rating": rating,
        }
    return outputs


def add_formats(workbook: xlsxwriter.Workbook) -> dict[str, Any]:
    return {
        "title": workbook.add_format(
            {"bold": True, "font_size": 18, "font_color": "#0B2545", "font_name": "Arial"}
        ),
        "subtitle": workbook.add_format({"font_size": 10, "font_color": "#52616B", "font_name": "Arial"}),
        "section": workbook.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#0B2545",
                "border": 0,
                "font_name": "Arial",
            }
        ),
        "subsection": workbook.add_format(
            {"bold": True, "font_color": "#0B2545", "bg_color": "#D9EAF7", "font_name": "Arial"}
        ),
        "header": workbook.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#1F6F78",
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "font_name": "Arial",
            }
        ),
        "input": workbook.add_format(
            {"font_color": "#0000FF", "num_format": "0.0%;[Red](0.0%);-", "font_name": "Arial"}
        ),
        "input_money": workbook.add_format(
            {"font_color": "#0000FF", "num_format": "$#,##0;[Red]($#,##0);-", "font_name": "Arial"}
        ),
        "money": workbook.add_format({"num_format": "$#,##0;[Red]($#,##0);-", "font_name": "Arial"}),
        "money_total": workbook.add_format(
            {
                "bold": True,
                "top": 1,
                "num_format": "$#,##0;[Red]($#,##0);-",
                "font_name": "Arial",
            }
        ),
        "pct": workbook.add_format({"num_format": "0.0%;[Red](0.0%);-", "font_name": "Arial"}),
        "pct_total": workbook.add_format(
            {"bold": True, "top": 1, "num_format": "0.0%;[Red](0.0%);-", "font_name": "Arial"}
        ),
        "multiple": workbook.add_format({"num_format": "0.0x;[Red](0.0x);-", "font_name": "Arial"}),
        "number": workbook.add_format({"num_format": "#,##0;[Red](#,##0);-", "font_name": "Arial"}),
        "number_total": workbook.add_format(
            {"bold": True, "top": 1, "num_format": "#,##0;[Red](#,##0);-", "font_name": "Arial"}
        ),
        "label": workbook.add_format({"font_name": "Arial"}),
        "label_indent": workbook.add_format({"indent": 1, "font_name": "Arial"}),
        "note": workbook.add_format(
            {"font_color": "#52616B", "font_size": 9, "text_wrap": True, "valign": "vcenter", "font_name": "Arial"}
        ),
        "ok": workbook.add_format(
            {"font_color": "#155724", "bg_color": "#D4EDDA", "bold": True, "align": "center", "font_name": "Arial"}
        ),
        "watch": workbook.add_format(
            {"font_color": "#7A5A00", "bg_color": "#FFF3CD", "bold": True, "align": "center", "font_name": "Arial"}
        ),
        "risk": workbook.add_format(
            {"font_color": "#721C24", "bg_color": "#F8D7DA", "bold": True, "align": "center", "font_name": "Arial"}
        ),
        "card_label": workbook.add_format(
            {"font_color": "#52616B", "font_size": 9, "align": "center", "font_name": "Arial"}
        ),
        "card_value": workbook.add_format(
            {"bold": True, "font_size": 16, "font_color": "#0B2545", "align": "center", "font_name": "Arial"}
        ),
        "card_money": workbook.add_format(
            {
                "bold": True,
                "font_size": 16,
                "font_color": "#0B2545",
                "align": "center",
                "num_format": "$#,##0",
                "font_name": "Arial",
            }
        ),
        "card_pct": workbook.add_format(
            {
                "bold": True,
                "font_size": 16,
                "font_color": "#0B2545",
                "align": "center",
                "num_format": "0.0%",
                "font_name": "Arial",
            }
        ),
        "source": workbook.add_format(
            {"font_color": "#008000", "text_wrap": True, "font_size": 9, "font_name": "Arial"}
        ),
        "border": workbook.add_format({"border": 1, "border_color": "#D9E2EC", "font_name": "Arial"}),
    }


def setup_sheet(ws: xlsxwriter.worksheet.Worksheet) -> None:
    ws.hide_gridlines(2)
    ws.set_zoom(90)


def write_formula(ws, row: int, col: int, formula: str, cell_format: Any, value: Any) -> None:
    ws.write_formula(row, col, formula, cell_format, value)


def source_comment(source: dict[str, Any] | None) -> str:
    if not source:
        return "No SEC source found."
    if source.get("synthetic_zero"):
        return f"Modeled as zero because SEC tag {source.get('tag')} was not reported for this fiscal year."
    lines = [
        f"SEC tag: {source.get('tag')}",
        f"Form: {source.get('form')} report date {source.get('report_date')}",
        f"Filed: {source.get('filing_date') or source.get('filed')}",
        f"Accession: {source.get('accession')}",
        f"URL: {source.get('url')}",
    ]
    return "\n".join(str(line) for line in lines if line)


def build_workbook(data: dict[str, Any], hist: dict[int, dict[str, float]], inputs, outputs) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(WORKBOOK_PATH)
    workbook.set_properties(
        {
            "title": "GPC ABL Borrowing Base Credit Model",
            "subject": "Asset-based lending credit analysis",
            "author": "Codex",
            "comments": "Built from SEC companyfacts with explicit ABL assumptions.",
        }
    )
    workbook.set_calc_mode("auto")
    fmt = add_formats(workbook)

    cover = workbook.add_worksheet("Cover")
    sources = workbook.add_worksheet("Sources")
    hist_ws = workbook.add_worksheet("Historical Financials")
    assumptions = workbook.add_worksheet("ABL Assumptions")
    ar_ws = workbook.add_worksheet("AR Aging")
    inv_ws = workbook.add_worksheet("Inventory Eligibility")
    conc_ws = workbook.add_worksheet("Customer Concentration")
    bb_ws = workbook.add_worksheet("Borrowing Base")
    stress_ws = workbook.add_worksheet("Cash Flow Stress")
    dash = workbook.add_worksheet("Dashboard")
    checks = workbook.add_worksheet("Checks")

    for ws in [cover, sources, hist_ws, assumptions, ar_ws, inv_ws, conc_ws, bb_ws, stress_ws, dash, checks]:
        setup_sheet(ws)

    # Cover
    cover.set_column("A:A", 3)
    cover.set_column("B:B", 42)
    cover.set_column("C:H", 15)
    cover.set_row(5, 30)
    cover.set_row(6, 30)
    cover.set_row(7, 30)
    cover.write("B2", "Asset Based Lending Credit Analysis", fmt["title"])
    cover.write("B3", "Borrowing Base and Cash Flow Stress Test | Genuine Parts Company (GPC)", fmt["subtitle"])
    cover.merge_range("B5:H5", "Project Use", fmt["section"])
    cover.merge_range("B6:B8", "Purpose", fmt["subsection"])
    cover.merge_range(
        "C6:H8",
        "Estimate ABL borrowing capacity from SEC financials.\nStress AR, inventory, reserves, liquidity, and cash-flow support.",
        fmt["note"],
    )
    cover.write("B8", "Workbook map", fmt["subsection"])
    map_rows = [
        ("Sources", "SEC filings, XBRL tags, and assumption notes."),
        ("Historical Financials", "Reported history and lender-style liquidity ratios."),
        ("ABL Assumptions", "Editable scenario drivers; blue font marks inputs."),
        ("AR Aging / Inventory / Customer Concentration", "Collateral eligibility schedules for ABL controls."),
        ("Borrowing Base", "Eligible collateral x advance rates, less reserves, vs. loan balance."),
        ("Cash Flow Stress", "Base/downside/stress repayment capacity and liquidity outputs."),
        ("Dashboard", "Credit committee style summary view."),
        ("Checks", "Tie-outs and model-status tests."),
    ]
    for idx, row in enumerate(map_rows, start=9):
        cover.write(idx, 1, row[0], fmt["label"])
        cover.merge_range(idx, 2, idx, 7, row[1], fmt["note"])
    cover.write("B20", "Key modeling caveat", fmt["subsection"])
    cover.merge_range(
        "C20:H22",
        "ABL detail schedules are illustrative.\nScaled to FY2025 reported AR and inventory.",
        fmt["note"],
    )
    latest_10k = data["as_of"]["latest_10k"]
    latest_10q = data["as_of"]["latest_10q"]
    cover.write("B24", "Latest annual filing", fmt["label"])
    cover.write_url("C24", latest_10k["url"], fmt["source"], f"{latest_10k['form']} filed {latest_10k['filing_date']}")
    cover.write("B25", "Latest quarterly filing context", fmt["label"])
    cover.write_url("C25", latest_10q["url"], fmt["source"], f"{latest_10q['form']} filed {latest_10q['filing_date']}")

    # Sources
    sources.set_column("A:A", 18)
    sources.set_column("B:C", 24)
    sources.set_column("D:D", 48)
    sources.set_column("E:G", 18)
    sources.set_column("H:H", 80)
    sources.write("A1", "Sources and Audit Trail", fmt["title"])
    sources.write("A3", "Company", fmt["label"])
    sources.write("B3", data["company"]["name"], fmt["label"])
    sources.write("A4", "CIK", fmt["label"])
    sources.write("B4", data["company"]["cik"], fmt["label"])
    sources.write("A5", "SIC", fmt["label"])
    sources.write("B5", f"{data['company']['sic']} - {data['company']['sic_description']}", fmt["label"])
    sources.write_url("A7", data["as_of"]["companyfacts_url"], fmt["source"], "SEC companyfacts API")
    sources.write_url("A8", data["as_of"]["submissions_url"], fmt["source"], "SEC submissions API")
    sources.write_row("A10", ["Metric", "Year", "Value ($mm)", "SEC Tag", "Form", "Filed", "Accession", "URL / Notes"], fmt["header"])
    r = 10
    for key, item in data["metrics"].items():
        for year in YEARS:
            value = item["values"][str(year)]["value"]
            source = item["values"][str(year)]["source"]
            sources.write(r, 0, item["label"], fmt["label"])
            sources.write(r, 1, year, fmt["number"])
            sources.write(r, 2, (float(value) / 1_000_000) if value is not None else None, fmt["money"])
            sources.write(r, 3, item["tag"], fmt["label"])
            sources.write(r, 4, source.get("form") if source else None, fmt["label"])
            sources.write(r, 5, source.get("filing_date") or source.get("filed") if source else None, fmt["label"])
            sources.write(r, 6, source.get("accession") if source else None, fmt["label"])
            if source and source.get("url"):
                sources.write_url(r, 7, source["url"], fmt["source"], "SEC filing")
            elif source and source.get("synthetic_zero"):
                sources.write(r, 7, "Modeled as zero because the SEC tag was not reported.", fmt["note"])
            else:
                sources.write(r, 7, "No source fact extracted.", fmt["note"])
            r += 1
    sources.write(r + 1, 0, "Assumption sources", fmt["section"])
    sources.write(
        r + 2,
        0,
        "ABL eligibility rates, reserves, concentration thresholds, and scenario stress cases are analyst assumptions designed to demonstrate borrowing-base mechanics. Adjust blue-font inputs on ABL Assumptions to test alternatives.",
        fmt["note"],
    )

    # Historical Financials
    hist_ws.freeze_panes(4, 1)
    hist_ws.set_column("A:A", 34)
    hist_ws.set_column("B:F", 15)
    hist_ws.set_column("G:G", 38)
    hist_ws.write("A1", "Historical Financials", fmt["title"])
    hist_ws.write("A2", "USD in millions except ratios. SEC-sourced inputs include cell comments with tag and filing references.", fmt["subtitle"])
    hist_ws.write_row("A4", ["Fiscal year", *YEARS, "Notes"], fmt["header"])
    hist_row = 4
    hist_row_map: dict[str, int] = {}

    hist_definitions = [
        ("revenue", "Revenue", "money", "revenue", None),
        ("revenue_growth", "Revenue growth", "pct", None, "growth"),
        ("pretax_income", "Pretax income", "money", "pretax_income", None),
        ("interest", "Interest expense / paid fallback", "money", "interest_expense", "interest_paid"),
        ("da", "Depreciation & amortization", "money", "depreciation_amortization", None),
        ("ebitda", "EBITDA proxy", "money_total", None, "ebitda"),
        ("ebitda_margin", "EBITDA margin", "pct", None, "ebitda_margin"),
        ("net_income", "Net income", "money", "net_income", None),
        ("cfo", "Operating cash flow", "money", "cfo", None),
        ("capex", "Capital expenditures", "money", "capex", None),
        ("fcf", "Free cash flow", "money_total", None, "fcf"),
        ("ar", "Accounts receivable", "money", "accounts_receivable", None),
        ("inventory", "Inventory", "money", "inventory", None),
        ("cash", "Cash and restricted cash", "money", "cash", None),
        ("current_assets", "Current assets", "money", "current_assets", None),
        ("current_liabilities", "Current liabilities", "money", "current_liabilities", None),
        ("st_borrowings", "Short-term borrowings", "money", "short_term_borrowings", None),
        ("commercial_paper", "Commercial paper", "money", "commercial_paper", None),
        ("ltd_current", "Current maturities of long-term debt", "money", "long_term_debt_current", None),
        ("ltd_noncurrent", "Long-term debt, noncurrent", "money", "long_term_debt_noncurrent", None),
        ("total_debt", "Total debt", "money_total", None, "total_debt"),
        ("net_debt", "Net debt", "money_total", None, "net_debt"),
        ("current_ratio", "Current ratio", "multiple", None, "current_ratio"),
        ("quick_ratio", "Quick ratio", "multiple", None, "quick_ratio"),
        ("dso", "DSO", "number", None, "dso"),
        ("inventory_days", "Inventory days (sales-based)", "number", None, "inventory_days"),
        ("current_liability_days", "Current liability days (sales-based)", "number", None, "current_liability_days"),
        ("ccc_proxy", "Cash conversion cycle proxy", "number_total", None, "ccc_proxy"),
        ("debt_to_ebitda", "Debt / EBITDA", "multiple", None, "debt_to_ebitda"),
        ("interest_coverage", "EBITDA / interest", "multiple", None, "interest_coverage"),
    ]

    for key, label, fmt_key, source_key, derived in hist_definitions:
        hist_row_map[key] = hist_row
        hist_ws.write(hist_row, 0, label, fmt["label"])
        for col_idx, year in enumerate(YEARS, start=1):
            value = hist[year][key] if key in hist[year] else None
            if source_key:
                hist_ws.write(hist_row, col_idx, value, fmt[fmt_key])
                source = source_for(data, source_key, year, fallback_key=derived if key == "interest" else None)
                hist_ws.write_comment(hist_row, col_idx, source_comment(source), {"width": 360, "height": 130})
            elif derived == "growth":
                if year == YEARS[0]:
                    hist_ws.write_blank(hist_row, col_idx, None, fmt[fmt_key])
                else:
                    rev = hist_row_map["revenue"]
                    formula = f"={xl_rowcol_to_cell(rev, col_idx)}/{xl_rowcol_to_cell(rev, col_idx - 1)}-1"
                    write_formula(ws=hist_ws, row=hist_row, col=col_idx, formula=formula, cell_format=fmt[fmt_key], value=value)
            else:
                if key == "ebitda":
                    formula = (
                        f"={xl_rowcol_to_cell(hist_row_map['pretax_income'], col_idx)}"
                        f"+{xl_rowcol_to_cell(hist_row_map['interest'], col_idx)}"
                        f"+{xl_rowcol_to_cell(hist_row_map['da'], col_idx)}"
                    )
                elif key == "ebitda_margin":
                    formula = f"={xl_rowcol_to_cell(hist_row_map['ebitda'], col_idx)}/{xl_rowcol_to_cell(hist_row_map['revenue'], col_idx)}"
                elif key == "fcf":
                    formula = f"={xl_rowcol_to_cell(hist_row_map['cfo'], col_idx)}-{xl_rowcol_to_cell(hist_row_map['capex'], col_idx)}"
                elif key == "total_debt":
                    formula = (
                        f"=SUM({xl_rowcol_to_cell(hist_row_map['st_borrowings'], col_idx)}:"
                        f"{xl_rowcol_to_cell(hist_row_map['ltd_noncurrent'], col_idx)})"
                    )
                elif key == "net_debt":
                    formula = f"={xl_rowcol_to_cell(hist_row_map['total_debt'], col_idx)}-{xl_rowcol_to_cell(hist_row_map['cash'], col_idx)}"
                elif key == "current_ratio":
                    formula = f"={xl_rowcol_to_cell(hist_row_map['current_assets'], col_idx)}/{xl_rowcol_to_cell(hist_row_map['current_liabilities'], col_idx)}"
                elif key == "quick_ratio":
                    formula = (
                        f"=({xl_rowcol_to_cell(hist_row_map['cash'], col_idx)}+"
                        f"{xl_rowcol_to_cell(hist_row_map['ar'], col_idx)})/"
                        f"{xl_rowcol_to_cell(hist_row_map['current_liabilities'], col_idx)}"
                    )
                elif key == "dso":
                    formula = f"={xl_rowcol_to_cell(hist_row_map['ar'], col_idx)}/{xl_rowcol_to_cell(hist_row_map['revenue'], col_idx)}*365"
                elif key == "inventory_days":
                    formula = f"={xl_rowcol_to_cell(hist_row_map['inventory'], col_idx)}/{xl_rowcol_to_cell(hist_row_map['revenue'], col_idx)}*365"
                elif key == "current_liability_days":
                    formula = f"={xl_rowcol_to_cell(hist_row_map['current_liabilities'], col_idx)}/{xl_rowcol_to_cell(hist_row_map['revenue'], col_idx)}*365"
                elif key == "ccc_proxy":
                    formula = (
                        f"={xl_rowcol_to_cell(hist_row_map['dso'], col_idx)}+"
                        f"{xl_rowcol_to_cell(hist_row_map['inventory_days'], col_idx)}-"
                        f"{xl_rowcol_to_cell(hist_row_map['current_liability_days'], col_idx)}"
                    )
                elif key == "debt_to_ebitda":
                    formula = f"={xl_rowcol_to_cell(hist_row_map['total_debt'], col_idx)}/{xl_rowcol_to_cell(hist_row_map['ebitda'], col_idx)}"
                elif key == "interest_coverage":
                    formula = f"={xl_rowcol_to_cell(hist_row_map['ebitda'], col_idx)}/{xl_rowcol_to_cell(hist_row_map['interest'], col_idx)}"
                else:
                    formula = ""
                write_formula(hist_ws, hist_row, col_idx, formula, fmt[fmt_key], value)
        if key == "interest":
            hist_ws.write(hist_row, 6, "Uses InterestExpense where reported; InterestPaidNet fallback for 2024-2025.", fmt["note"])
        elif key in {"ebitda", "ccc_proxy"}:
            hist_ws.write(
                hist_row,
                6,
                "Analyst proxy for lending review; not company-reported adjusted EBITDA or a formal operating CCC metric.",
                fmt["note"],
            )
        hist_row += 1

    # ABL Assumptions
    assumptions.freeze_panes(4, 1)
    assumptions.set_column("A:A", 34)
    assumptions.set_column("B:D", 15)
    assumptions.set_column("E:E", 58)
    assumptions.write("A1", "ABL Assumptions", fmt["title"])
    assumptions.write("A2", "Blue-font cells are editable scenario assumptions. Values shown are illustrative lender assumptions.", fmt["subtitle"])
    assumptions.write_row("A4", ["Driver", *SCENARIOS, "Rationale"], fmt["header"])
    assumption_rows = {
        "revenue_growth": ("Revenue growth / shock", "pct", "Revenue stress drives AR balance and borrower cash flow."),
        "ebitda_margin_change": ("EBITDA margin change", "pct", "Downside cases compress cash-flow repayment capacity."),
        "ar_eligible_pct": ("AR eligible %", "pct", "Excludes 90+ day, disputed, affiliate, unsupported foreign, and cross-aged receivables."),
        "ar_advance_rate": ("AR advance rate", "pct", "Typical ABL AR advance range for eligible receivables."),
        "inventory_balance_change": ("Inventory balance change", "pct", "Downside assumes inventory builds while eligibility and NOLV support deteriorate."),
        "inventory_eligible_pct": ("Inventory eligible %", "pct", "Excludes obsolete, consigned, slow-moving, and unsupported offsite inventory."),
        "inventory_advance_rate": ("Inventory advance rate", "pct", "Applied against eligible inventory; conservative vs. AR advance."),
        "dilution_reserve_pct": ("Dilution reserve % of gross AR", "pct", "Proxy for credits, disputes, returns, and unapplied cash."),
        "concentration_reserve_pct": ("Concentration reserve % of gross AR", "pct", "Proxy for obligor concentration above borrowing-base limits."),
        "other_reserves": ("Other reserves ($mm)", "money", "Rent, freight, tax, and other collateral-control reserves."),
        "availability_block": ("Availability block ($mm)", "money", "Minimum cushion retained by the lender."),
        "loan_balance_change": ("ABL loan balance change", "pct", "Downside assumes higher revolver draw / working-capital funding need."),
        "interest_rate_uplift": ("Interest rate uplift", "pct", "Sensitivity to higher borrowing cost."),
        "cfo_conversion": ("CFO conversion of EBITDA", "pct", "Working-capital stress lowers EBITDA-to-cash conversion."),
        "capex_pct_revenue": ("Capex % of revenue", "pct", "Simplified maintenance and growth capex assumption."),
    }
    assumption_row_map: dict[str, int] = {}
    row = 4
    for key, (label, ftype, rationale) in assumption_rows.items():
        assumption_row_map[key] = row
        assumptions.write(row, 0, label, fmt["label"])
        for col_idx, scenario in enumerate(SCENARIOS, start=1):
            cell_fmt = fmt["input_money"] if ftype == "money" else fmt["input"]
            assumptions.write(row, col_idx, inputs[scenario][key], cell_fmt)
        assumptions.write(row, 4, rationale, fmt["note"])
        row += 1

    # AR Aging
    ar_ws.set_column("A:A", 26)
    ar_ws.set_column("B:G", 16)
    ar_ws.set_column("H:H", 44)
    ar_ws.write("A1", "AR Aging and Eligibility", fmt["title"])
    ar_ws.write("A2", "Illustrative borrowing-base certificate schedule scaled to reported FY2025 accounts receivable.", fmt["subtitle"])
    ar_ws.write("A4", "Reported 2025 AR", fmt["label"])
    write_formula(ar_ws, 3, 1, "='Historical Financials'!F16", fmt["money"], hist[2025]["ar"])
    ar_ws.write_row("A6", ["Bucket / Exclusion", "% of AR", "Gross AR", "Ineligible %", "Eligible AR", "Advance Rate", "Advance Value", "ABL treatment"], fmt["header"])
    ar_buckets = [
        ("Current", 0.58, 0.00, "Eligible if billed, earned, and not contra/affiliate."),
        ("1-30 days past due", 0.20, 0.00, "Normally eligible subject to concentration and dilution controls."),
        ("31-60 days past due", 0.10, 0.10, "Partial haircut for aging deterioration."),
        ("61-90 days past due", 0.05, 0.25, "Higher ineligibility due to near-threshold aging."),
        ("90+ days past due", 0.04, 1.00, "Typically ineligible under ABL borrowing-base rules."),
        ("Disputed / offset", 0.02, 1.00, "Excluded due to dispute, credit memo, or payable offset."),
        ("Affiliate / unsupported foreign", 0.01, 1.00, "Excluded without acceptable credit support."),
    ]
    ar_total_row = 6 + len(ar_buckets)
    for idx, (bucket, pct, ineligible, treatment) in enumerate(ar_buckets, start=6):
        ar_ws.write(idx, 0, bucket, fmt["label"])
        ar_ws.write(idx, 1, pct, fmt["input"])
        write_formula(ar_ws, idx, 2, f"=$B$4*{xl_rowcol_to_cell(idx, 1)}", fmt["money"], hist[2025]["ar"] * pct)
        ar_ws.write(idx, 3, ineligible, fmt["input"])
        write_formula(ar_ws, idx, 4, f"={xl_rowcol_to_cell(idx, 2)}*(1-{xl_rowcol_to_cell(idx, 3)})", fmt["money"], hist[2025]["ar"] * pct * (1 - ineligible))
        write_formula(ar_ws, idx, 5, "='ABL Assumptions'!B7", fmt["pct"], inputs["Base"]["ar_advance_rate"])
        write_formula(ar_ws, idx, 6, f"={xl_rowcol_to_cell(idx, 4)}*{xl_rowcol_to_cell(idx, 5)}", fmt["money"], hist[2025]["ar"] * pct * (1 - ineligible) * inputs["Base"]["ar_advance_rate"])
        ar_ws.write(idx, 7, treatment, fmt["note"])
    ar_ws.write(ar_total_row, 0, "Total", fmt["label"])
    ar_ws.write_formula(ar_total_row, 1, f"=SUM(B7:B{ar_total_row})", fmt["pct_total"], 1.0)
    ar_ws.write_formula(ar_total_row, 2, f"=SUM(C7:C{ar_total_row})", fmt["money_total"], hist[2025]["ar"])
    ar_ws.write_formula(ar_total_row, 4, f"=SUM(E7:E{ar_total_row})", fmt["money_total"], hist[2025]["ar"] * 0.9175)
    ar_ws.write_formula(ar_total_row, 6, f"=SUM(G7:G{ar_total_row})", fmt["money_total"], hist[2025]["ar"] * 0.9175 * inputs["Base"]["ar_advance_rate"])

    # Inventory Eligibility
    inv_ws.set_column("A:A", 28)
    inv_ws.set_column("B:G", 16)
    inv_ws.set_column("H:H", 44)
    inv_ws.write("A1", "Inventory Eligibility", fmt["title"])
    inv_ws.write("A2", "Illustrative eligibility schedule scaled to reported FY2025 inventory.", fmt["subtitle"])
    inv_ws.write("A4", "Reported 2025 inventory", fmt["label"])
    write_formula(inv_ws, 3, 1, "='Historical Financials'!F17", fmt["money"], hist[2025]["inventory"])
    inv_ws.write_row("A6", ["Inventory segment", "% of inventory", "Gross inventory", "Eligible %", "Eligible inventory", "Advance Rate", "Advance Value", "ABL treatment"], fmt["header"])
    inv_segments = [
        ("Finished goods / parts", 0.72, 1.00, "Primary collateral pool for an auto-parts distributor."),
        ("Slow-moving inventory", 0.10, 0.50, "Partial eligibility haircut due to turnover risk."),
        ("Raw materials", 0.00, 0.00, "Not material for distributor model."),
        ("Work-in-process", 0.00, 0.00, "Not material for distributor model."),
        ("Obsolete inventory", 0.05, 0.00, "Excluded from eligible inventory."),
        ("Consigned inventory", 0.04, 0.00, "Excluded unless lender has acceptable control agreement."),
        ("Unsupported offsite inventory", 0.03, 0.00, "Excluded without field exam/control support."),
        ("Other eligible inventory", 0.06, 1.00, "Eligible after standard reserves and NOLV haircut."),
    ]
    inv_total_row = 6 + len(inv_segments)
    for idx, (segment, pct, eligible_pct, treatment) in enumerate(inv_segments, start=6):
        inv_ws.write(idx, 0, segment, fmt["label"])
        inv_ws.write(idx, 1, pct, fmt["input"])
        write_formula(inv_ws, idx, 2, f"=$B$4*{xl_rowcol_to_cell(idx, 1)}", fmt["money"], hist[2025]["inventory"] * pct)
        inv_ws.write(idx, 3, eligible_pct, fmt["input"])
        write_formula(inv_ws, idx, 4, f"={xl_rowcol_to_cell(idx, 2)}*{xl_rowcol_to_cell(idx, 3)}", fmt["money"], hist[2025]["inventory"] * pct * eligible_pct)
        write_formula(inv_ws, idx, 5, "='ABL Assumptions'!B10", fmt["pct"], inputs["Base"]["inventory_advance_rate"])
        write_formula(inv_ws, idx, 6, f"={xl_rowcol_to_cell(idx, 4)}*{xl_rowcol_to_cell(idx, 5)}", fmt["money"], hist[2025]["inventory"] * pct * eligible_pct * inputs["Base"]["inventory_advance_rate"])
        inv_ws.write(idx, 7, treatment, fmt["note"])
    inv_ws.write(inv_total_row, 0, "Total", fmt["label"])
    inv_ws.write_formula(inv_total_row, 1, f"=SUM(B7:B{inv_total_row})", fmt["pct_total"], 1.0)
    inv_ws.write_formula(inv_total_row, 2, f"=SUM(C7:C{inv_total_row})", fmt["money_total"], hist[2025]["inventory"])
    inv_ws.write_formula(inv_total_row, 4, f"=SUM(E7:E{inv_total_row})", fmt["money_total"], hist[2025]["inventory"] * 0.83)
    inv_ws.write_formula(inv_total_row, 6, f"=SUM(G7:G{inv_total_row})", fmt["money_total"], hist[2025]["inventory"] * 0.83 * inputs["Base"]["inventory_advance_rate"])

    # Customer Concentration
    conc_ws.set_column("A:A", 20)
    conc_ws.set_column("B:G", 16)
    conc_ws.set_column("H:H", 38)
    conc_ws.write("A1", "Customer Concentration Risk", fmt["title"])
    conc_ws.write("A2", "Illustrative obligor schedule. Public filings do not provide AR by customer.", fmt["subtitle"])
    conc_ws.write("A4", "Reported 2025 AR", fmt["label"])
    write_formula(conc_ws, 3, 1, "='Historical Financials'!F16", fmt["money"], hist[2025]["ar"])
    conc_ws.write_row("A6", ["Customer", "% of AR", "AR balance", "Limit", "Excess over limit", "Reserve @ AR advance", "Flag", "Notes"], fmt["header"])
    concentration = [
        ("Customer A", 0.14),
        ("Customer B", 0.11),
        ("Customer C", 0.08),
        ("Customer D", 0.07),
        ("Customer E", 0.06),
        ("Customer F", 0.05),
        ("Customer G", 0.04),
        ("Customer H", 0.03),
        ("Customer I", 0.02),
        ("Customer J", 0.015),
        ("All other customers", 0.385),
    ]
    conc_total_row = 6 + len(concentration)
    for idx, (customer, pct) in enumerate(concentration, start=6):
        conc_ws.write(idx, 0, customer, fmt["label"])
        conc_ws.write(idx, 1, pct, fmt["input"])
        write_formula(conc_ws, idx, 2, f"=$B$4*{xl_rowcol_to_cell(idx, 1)}", fmt["money"], hist[2025]["ar"] * pct)
        conc_ws.write(idx, 3, 0.10, fmt["input"])
        write_formula(conc_ws, idx, 4, f"=MAX(0,{xl_rowcol_to_cell(idx, 2)}-$B$4*{xl_rowcol_to_cell(idx, 3)})", fmt["money"], max(0, hist[2025]["ar"] * pct - hist[2025]["ar"] * 0.10))
        write_formula(conc_ws, idx, 5, f"={xl_rowcol_to_cell(idx, 4)}*'ABL Assumptions'!B7", fmt["money"], max(0, hist[2025]["ar"] * pct - hist[2025]["ar"] * 0.10) * inputs["Base"]["ar_advance_rate"])
        flag = "Breach" if pct > 0.10 else "OK"
        write_formula(conc_ws, idx, 6, f'=IF({xl_rowcol_to_cell(idx, 1)}>{xl_rowcol_to_cell(idx, 3)},"Breach","OK")', fmt["watch"] if flag == "Breach" else fmt["ok"], flag)
        conc_ws.write(idx, 7, "Concentration limit set at 10% for demonstration.", fmt["note"])
    conc_ws.write(conc_total_row, 0, "Total", fmt["label"])
    conc_ws.write_formula(conc_total_row, 1, f"=SUM(B7:B{conc_total_row})", fmt["pct_total"], 1.0)
    conc_ws.write_formula(conc_total_row, 2, f"=SUM(C7:C{conc_total_row})", fmt["money_total"], hist[2025]["ar"])
    conc_ws.write_formula(conc_total_row, 5, f"=SUM(F7:F{conc_total_row})", fmt["money_total"], outputs["Base"]["concentration_reserve"])

    # Borrowing Base
    bb_ws.freeze_panes(4, 1)
    bb_ws.set_column("A:A", 36)
    bb_ws.set_column("B:D", 16)
    bb_ws.set_column("E:E", 48)
    bb_ws.write("A1", "Borrowing Base Calculation", fmt["title"])
    bb_ws.write("A2", "Eligible AR plus eligible inventory, less reserves, compared with estimated ABL loan balance.", fmt["subtitle"])
    bb_ws.write_row("A4", ["Borrowing-base component", *SCENARIOS, "Formula / Control"], fmt["header"])
    bb_rows = [
        ("gross_ar", "Gross AR", "money", "Reported AR scaled by revenue stress."),
        ("ar_eligible_pct", "AR eligible %", "pct", "Scenario assumption."),
        ("eligible_ar", "Eligible AR", "money", "Gross AR x AR eligible %."),
        ("ar_advance_rate", "AR advance rate", "pct", "Scenario assumption."),
        ("ar_availability", "AR availability", "money", "Eligible AR x AR advance rate."),
        ("gross_inventory", "Gross inventory", "money", "Reported inventory scaled by inventory build assumption."),
        ("inventory_eligible_pct", "Inventory eligible %", "pct", "Scenario assumption."),
        ("eligible_inventory", "Eligible inventory", "money", "Gross inventory x inventory eligible %."),
        ("inventory_advance_rate", "Inventory advance rate", "pct", "Scenario assumption."),
        ("inventory_availability", "Inventory availability", "money", "Eligible inventory x inventory advance rate."),
        ("total_collateral", "Total collateral availability", "money_total", "AR availability + inventory availability."),
        ("dilution_reserve", "Dilution reserve", "money", "Gross AR x dilution reserve %."),
        ("concentration_reserve", "Concentration reserve", "money", "Greater of scenario reserve and illustrative concentration schedule for Base."),
        ("other_reserves", "Other reserves", "money", "Scenario assumption."),
        ("availability_block", "Availability block", "money", "Scenario assumption."),
        ("total_reserves", "Total reserves", "money_total", "Sum of reserves."),
        ("borrowing_base", "Borrowing base", "money_total", "Total collateral availability less reserves."),
        ("loan_balance", "Estimated ABL loan balance", "money", "2025 short-term borrowings + CP + current debt, adjusted by scenario."),
        ("excess_availability", "Excess availability", "money_total", "Borrowing base less loan balance."),
        ("utilization", "Utilization", "pct_total", "Loan balance / borrowing base."),
        ("min_liquidity", "Minimum liquidity", "money_total", "Cash + excess availability."),
        ("alert", "Availability alert", "text", "Flags low availability or borrowing-base deficiency."),
    ]
    bb_row_map: dict[str, int] = {}
    for ridx, (key, label, ftype, note) in enumerate(bb_rows, start=4):
        bb_row_map[key] = ridx
        bb_ws.write(ridx, 0, label, fmt["label"] if "total" not in ftype else fmt["label"])
        for cidx, scenario in enumerate(SCENARIOS, start=1):
            out = outputs[scenario]
            assumption = inputs[scenario]
            if key == "gross_ar":
                formula = f"='Historical Financials'!F16*(1+'ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['revenue_growth'], cidx)})"
                value = out[key]
            elif key == "ar_eligible_pct":
                formula = f"='ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['ar_eligible_pct'], cidx)}"
                value = assumption[key]
            elif key == "eligible_ar":
                formula = f"={xl_rowcol_to_cell(bb_row_map['gross_ar'], cidx)}*{xl_rowcol_to_cell(bb_row_map['ar_eligible_pct'], cidx)}"
                value = out[key]
            elif key == "ar_advance_rate":
                formula = f"='ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['ar_advance_rate'], cidx)}"
                value = assumption[key]
            elif key == "ar_availability":
                formula = f"={xl_rowcol_to_cell(bb_row_map['eligible_ar'], cidx)}*{xl_rowcol_to_cell(bb_row_map['ar_advance_rate'], cidx)}"
                value = out[key]
            elif key == "gross_inventory":
                formula = f"='Historical Financials'!F17*(1+'ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['inventory_balance_change'], cidx)})"
                value = out[key]
            elif key == "inventory_eligible_pct":
                formula = f"='ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['inventory_eligible_pct'], cidx)}"
                value = assumption[key]
            elif key == "eligible_inventory":
                formula = f"={xl_rowcol_to_cell(bb_row_map['gross_inventory'], cidx)}*{xl_rowcol_to_cell(bb_row_map['inventory_eligible_pct'], cidx)}"
                value = out[key]
            elif key == "inventory_advance_rate":
                formula = f"='ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['inventory_advance_rate'], cidx)}"
                value = assumption[key]
            elif key == "inventory_availability":
                formula = f"={xl_rowcol_to_cell(bb_row_map['eligible_inventory'], cidx)}*{xl_rowcol_to_cell(bb_row_map['inventory_advance_rate'], cidx)}"
                value = out[key]
            elif key == "total_collateral":
                formula = f"={xl_rowcol_to_cell(bb_row_map['ar_availability'], cidx)}+{xl_rowcol_to_cell(bb_row_map['inventory_availability'], cidx)}"
                value = out["ar_availability"] + out["inventory_availability"]
            elif key == "dilution_reserve":
                formula = f"={xl_rowcol_to_cell(bb_row_map['gross_ar'], cidx)}*'ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['dilution_reserve_pct'], cidx)}"
                value = out[key]
            elif key == "concentration_reserve":
                formula = (
                    f"=MAX({xl_rowcol_to_cell(bb_row_map['gross_ar'], cidx)}*'ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['concentration_reserve_pct'], cidx)},"
                    f"'Customer Concentration'!$F${conc_total_row+1}*{xl_rowcol_to_cell(bb_row_map['gross_ar'], cidx)}/'Customer Concentration'!$C${conc_total_row+1})"
                )
                value = out[key]
            elif key == "other_reserves":
                formula = f"='ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['other_reserves'], cidx)}"
                value = assumption[key]
            elif key == "availability_block":
                formula = f"='ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['availability_block'], cidx)}"
                value = assumption[key]
            elif key == "total_reserves":
                formula = f"=SUM({xl_rowcol_to_cell(bb_row_map['dilution_reserve'], cidx)}:{xl_rowcol_to_cell(bb_row_map['availability_block'], cidx)})"
                value = out[key]
            elif key == "borrowing_base":
                formula = f"={xl_rowcol_to_cell(bb_row_map['total_collateral'], cidx)}-{xl_rowcol_to_cell(bb_row_map['total_reserves'], cidx)}"
                value = out[key]
            elif key == "loan_balance":
                formula = (
                    f"=SUM('Historical Financials'!F21:F23)*(1+'ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['loan_balance_change'], cidx)})"
                )
                value = out[key]
            elif key == "excess_availability":
                formula = f"={xl_rowcol_to_cell(bb_row_map['borrowing_base'], cidx)}-{xl_rowcol_to_cell(bb_row_map['loan_balance'], cidx)}"
                value = out[key]
            elif key == "utilization":
                formula = f"={xl_rowcol_to_cell(bb_row_map['loan_balance'], cidx)}/{xl_rowcol_to_cell(bb_row_map['borrowing_base'], cidx)}"
                value = out[key]
            elif key == "min_liquidity":
                formula = f"='Historical Financials'!F18+{xl_rowcol_to_cell(bb_row_map['excess_availability'], cidx)}"
                value = out[key]
            elif key == "alert":
                formula = f'=IF({xl_rowcol_to_cell(bb_row_map["excess_availability"], cidx)}<0,"Deficiency",IF({xl_rowcol_to_cell(bb_row_map["excess_availability"], cidx)}<500,"Low Availability","OK"))'
                value = "Deficiency" if out["excess_availability"] < 0 else "Low Availability" if out["excess_availability"] < 500 else "OK"
            else:
                formula, value = "", None
            if ftype == "text":
                cell_fmt = fmt["risk"] if value == "Deficiency" else fmt["watch"] if value == "Low Availability" else fmt["ok"]
            else:
                cell_fmt = fmt.get(ftype, fmt["money"])
            write_formula(bb_ws, ridx, cidx, formula, cell_fmt, value)
        bb_ws.write(ridx, 4, note, fmt["note"])

    # Cash Flow Stress
    stress_ws.freeze_panes(4, 1)
    stress_ws.set_column("A:A", 34)
    stress_ws.set_column("B:D", 16)
    stress_ws.set_column("E:E", 48)
    stress_ws.write("A1", "Cash Flow Stress Test", fmt["title"])
    stress_ws.write("A2", "Tests whether cash flow and collateral availability support debt service under weakening cases.", fmt["subtitle"])
    stress_ws.write_row("A4", ["Metric", *SCENARIOS, "Interpretation"], fmt["header"])
    stress_rows = [
        ("revenue", "Revenue", "money", "2025 revenue adjusted by scenario shock."),
        ("ebitda_margin", "EBITDA margin", "pct", "2025 EBITDA proxy margin plus scenario margin change."),
        ("ebitda", "EBITDA proxy", "money_total", "Revenue x EBITDA margin."),
        ("interest", "Interest cash cost", "money", "2025 interest paid plus rate uplift on ABL loan balance."),
        ("coverage", "EBITDA / interest", "multiple", "Primary cash-flow coverage indicator."),
        ("total_debt", "Total debt", "money", "2025 total debt plus incremental scenario draw."),
        ("debt_to_ebitda", "Debt / EBITDA", "multiple", "Leverage pressure indicator."),
        ("cfo", "Operating cash flow proxy", "money", "EBITDA x CFO conversion assumption."),
        ("capex", "Capex", "money", "Revenue x capex % assumption."),
        ("fcf", "Free cash flow", "money_total", "Operating cash flow proxy less capex."),
        ("excess_availability", "Excess availability", "money_total", "Linked to borrowing-base schedule."),
        ("min_liquidity", "Minimum liquidity", "money_total", "Cash + excess availability."),
        ("rating", "Scenario risk rating", "text", "Pass / Watch / High Risk based on liquidity, coverage, and leverage."),
    ]
    stress_row_map = {}
    for ridx, (key, label, ftype, note) in enumerate(stress_rows, start=4):
        stress_row_map[key] = ridx
        stress_ws.write(ridx, 0, label, fmt["label"])
        for cidx, scenario in enumerate(SCENARIOS, start=1):
            out = outputs[scenario]
            if key == "revenue":
                formula = f"='Historical Financials'!F5*(1+'ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['revenue_growth'], cidx)})"
            elif key == "ebitda_margin":
                formula = f"=MAX(0,'Historical Financials'!F11+'ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['ebitda_margin_change'], cidx)})"
            elif key == "ebitda":
                formula = f"={xl_rowcol_to_cell(stress_row_map['revenue'], cidx)}*{xl_rowcol_to_cell(stress_row_map['ebitda_margin'], cidx)}"
            elif key == "interest":
                formula = (
                    f"='Historical Financials'!{xl_rowcol_to_cell(hist_row_map['interest'], 5)}"
                    f"+'Borrowing Base'!{xl_rowcol_to_cell(bb_row_map['loan_balance'], cidx)}"
                    f"*'ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['interest_rate_uplift'], cidx)}"
                )
            elif key == "coverage":
                formula = f"={xl_rowcol_to_cell(stress_row_map['ebitda'], cidx)}/{xl_rowcol_to_cell(stress_row_map['interest'], cidx)}"
            elif key == "total_debt":
                formula = f"='Historical Financials'!F25+('Borrowing Base'!{xl_rowcol_to_cell(bb_row_map['loan_balance'], cidx)}-'Borrowing Base'!B22)"
            elif key == "debt_to_ebitda":
                formula = f"={xl_rowcol_to_cell(stress_row_map['total_debt'], cidx)}/{xl_rowcol_to_cell(stress_row_map['ebitda'], cidx)}"
            elif key == "cfo":
                formula = f"={xl_rowcol_to_cell(stress_row_map['ebitda'], cidx)}*'ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['cfo_conversion'], cidx)}"
            elif key == "capex":
                formula = f"={xl_rowcol_to_cell(stress_row_map['revenue'], cidx)}*'ABL Assumptions'!{xl_rowcol_to_cell(assumption_row_map['capex_pct_revenue'], cidx)}"
            elif key == "fcf":
                formula = f"={xl_rowcol_to_cell(stress_row_map['cfo'], cidx)}-{xl_rowcol_to_cell(stress_row_map['capex'], cidx)}"
            elif key == "excess_availability":
                formula = f"='Borrowing Base'!{xl_rowcol_to_cell(bb_row_map['excess_availability'], cidx)}"
            elif key == "min_liquidity":
                formula = f"='Borrowing Base'!{xl_rowcol_to_cell(bb_row_map['min_liquidity'], cidx)}"
            elif key == "rating":
                formula = (
                    f'=IF(OR({xl_rowcol_to_cell(stress_row_map["excess_availability"], cidx)}<0,'
                    f'{xl_rowcol_to_cell(stress_row_map["min_liquidity"], cidx)}<0,'
                    f'{xl_rowcol_to_cell(stress_row_map["coverage"], cidx)}<1,'
                    f'{xl_rowcol_to_cell(stress_row_map["debt_to_ebitda"], cidx)}>8),"High Risk",'
                    f'IF(OR({xl_rowcol_to_cell(stress_row_map["excess_availability"], cidx)}<500,'
                    f'{xl_rowcol_to_cell(stress_row_map["coverage"], cidx)}<2,'
                    f'{xl_rowcol_to_cell(stress_row_map["debt_to_ebitda"], cidx)}>5),"Watch","Pass"))'
                )
            else:
                formula = ""
            value = out[key]
            if ftype == "text":
                cell_fmt = fmt["risk"] if value == "High Risk" else fmt["watch"] if value == "Watch" else fmt["ok"]
            else:
                cell_fmt = fmt.get(ftype, fmt["money"])
            write_formula(stress_ws, ridx, cidx, formula, cell_fmt, value)
        stress_ws.write(ridx, 4, note, fmt["note"])

    # Dashboard
    dash.set_column("A:A", 2)
    dash.set_column("B:M", 13)
    dash.set_row(0, 26)
    dash.write("B2", "ABL Credit Dashboard: Genuine Parts Company", fmt["title"])
    dash.write("B3", "Borrowing base, liquidity, cash-flow coverage, and downside protection summary.", fmt["subtitle"])
    kpis = [
        ("Base borrowing base", "='Borrowing Base'!B21", outputs["Base"]["borrowing_base"], "money"),
        ("Base excess availability", "='Borrowing Base'!B23", outputs["Base"]["excess_availability"], "money"),
        ("Base utilization", "='Borrowing Base'!B24", outputs["Base"]["utilization"], "pct"),
        ("Stress excess availability", "='Borrowing Base'!D23", outputs["Stress"]["excess_availability"], "money"),
        ("Stress EBITDA / interest", "='Cash Flow Stress'!D9", outputs["Stress"]["coverage"], "multiple"),
        ("Stress risk rating", "='Cash Flow Stress'!D17", outputs["Stress"]["rating"], "text"),
    ]
    start_col = 1
    for i, (label, formula, value, ktype) in enumerate(kpis):
        col = start_col + i * 2
        dash.merge_range(4, col, 4, col + 1, label, fmt["card_label"])
        if ktype == "money":
            dash.merge_range(5, col, 6, col + 1, "", fmt["card_money"])
            write_formula(dash, 5, col, formula, fmt["card_money"], value)
        elif ktype == "pct":
            dash.merge_range(5, col, 6, col + 1, "", fmt["card_pct"])
            write_formula(dash, 5, col, formula, fmt["card_pct"], value)
        elif ktype == "multiple":
            dash.merge_range(5, col, 6, col + 1, "", fmt["card_value"])
            write_formula(dash, 5, col, formula, fmt["multiple"], value)
        else:
            dash.merge_range(5, col, 6, col + 1, "", fmt["card_value"])
            cell_fmt = fmt["risk"] if value == "High Risk" else fmt["watch"] if value == "Watch" else fmt["ok"]
            write_formula(dash, 5, col, formula, cell_fmt, value)

    dash.write("B9", "Risk Alerts", fmt["section"])
    dash.write_row("B10", ["Alert", "Status", "Comment"], fmt["header"])
    alerts = [
        ("Low excess availability", outputs["Stress"]["excess_availability"] < 500, "Stress excess availability below $500mm threshold."),
        ("High past-due AR", True, "Illustrative AR aging has 90+ day and disputed ineligibles; monitor through BBC reporting."),
        ("Inventory ineligibility", True, "Downside/stress eligibility falls materially as slow-moving/obsolete risk rises."),
        ("Customer concentration", True, "Top two illustrative customers exceed 10% threshold."),
        ("Borrowing base deficiency", outputs["Stress"]["excess_availability"] < 0, "Stress case tests deficiency trigger."),
        ("Cash-flow coverage pressure", outputs["Stress"]["coverage"] < 2.0, "Stress EBITDA / interest below 2.0x watch threshold."),
    ]
    for idx, (alert, triggered, comment) in enumerate(alerts, start=10):
        dash.write(idx, 1, alert, fmt["label"])
        dash.write(idx, 2, "Watch" if triggered else "OK", fmt["watch"] if triggered else fmt["ok"])
        dash.write(idx, 3, comment, fmt["note"])

    line_chart = workbook.add_chart({"type": "line"})
    line_chart.add_series(
        {
            "name": "Revenue",
            "categories": ["Historical Financials", 3, 1, 3, 5],
            "values": ["Historical Financials", hist_row_map["revenue"], 1, hist_row_map["revenue"], 5],
            "line": {"color": "#1F6F78", "width": 2.25},
        }
    )
    line_chart.add_series(
        {
            "name": "EBITDA proxy",
            "categories": ["Historical Financials", 3, 1, 3, 5],
            "values": ["Historical Financials", hist_row_map["ebitda"], 1, hist_row_map["ebitda"], 5],
            "line": {"color": "#B7791F", "width": 2.25},
        }
    )
    line_chart.set_title({"name": "Revenue and EBITDA Trend ($mm)"})
    line_chart.set_legend({"position": "bottom"})
    line_chart.set_y_axis({"num_format": "$#,##0"})
    line_chart.set_style(10)
    dash.insert_chart("F10", line_chart, {"x_scale": 1.25, "y_scale": 1.05})

    collateral_chart = workbook.add_chart({"type": "column", "subtype": "stacked"})
    collateral_chart.add_series(
        {
            "name": "AR availability",
            "categories": ["Borrowing Base", 3, 1, 3, 3],
            "values": ["Borrowing Base", bb_row_map["ar_availability"], 1, bb_row_map["ar_availability"], 3],
            "fill": {"color": "#1F6F78"},
        }
    )
    collateral_chart.add_series(
        {
            "name": "Inventory availability",
            "categories": ["Borrowing Base", 3, 1, 3, 3],
            "values": ["Borrowing Base", bb_row_map["inventory_availability"], 1, bb_row_map["inventory_availability"], 3],
            "fill": {"color": "#5B8C5A"},
        }
    )
    collateral_chart.set_title({"name": "Collateral Availability by Scenario ($mm)"})
    collateral_chart.set_legend({"position": "bottom"})
    collateral_chart.set_y_axis({"num_format": "$#,##0"})
    collateral_chart.set_style(11)
    dash.insert_chart("B19", collateral_chart, {"x_scale": 1.18, "y_scale": 1.05})

    availability_chart = workbook.add_chart({"type": "column"})
    availability_chart.add_series(
        {
            "name": "Excess availability",
            "categories": ["Borrowing Base", 3, 1, 3, 3],
            "values": ["Borrowing Base", bb_row_map["excess_availability"], 1, bb_row_map["excess_availability"], 3],
            "fill": {"color": "#0B2545"},
        }
    )
    availability_chart.set_title({"name": "Excess Availability by Scenario ($mm)"})
    availability_chart.set_y_axis({"num_format": "$#,##0"})
    availability_chart.set_legend({"none": True})
    availability_chart.set_style(10)
    dash.insert_chart("F19", availability_chart, {"x_scale": 1.18, "y_scale": 1.05})

    concentration_chart = workbook.add_chart({"type": "bar"})
    concentration_chart.add_series(
        {
            "name": "AR balance",
            "categories": ["Customer Concentration", 6, 0, 10, 0],
            "values": ["Customer Concentration", 6, 2, 10, 2],
            "fill": {"color": "#8B5E34"},
        }
    )
    concentration_chart.set_title({"name": "Top 5 Customer AR Exposure ($mm)"})
    concentration_chart.set_x_axis({"num_format": "$#,##0"})
    concentration_chart.set_legend({"none": True})
    concentration_chart.set_style(10)
    dash.insert_chart("J19", concentration_chart, {"x_scale": 1.0, "y_scale": 1.05})

    # Checks
    checks.set_column("A:A", 38)
    checks.set_column("B:D", 18)
    checks.set_column("E:E", 58)
    checks.write("A1", "Model Checks", fmt["title"])
    checks.write_row("A4", ["Check", "Actual", "Expected / Tolerance", "Status", "Notes"], fmt["header"])
    check_rows = [
        ("AR aging percentages sum to 100%", "='AR Aging'!B14", 1.0, "ABS(B5-C5)<0.0001", "Illustrative aging schedule tie-out."),
        ("Inventory percentages sum to 100%", "='Inventory Eligibility'!B15", 1.0, "ABS(B6-C6)<0.0001", "Illustrative inventory schedule tie-out."),
        ("Customer concentration sums to 100%", "='Customer Concentration'!B18", 1.0, "ABS(B7-C7)<0.0001", "Illustrative concentration schedule tie-out."),
        ("Base excess availability positive", "='Borrowing Base'!B23", 0.0, "B8>C8", "Collateral support covers modeled ABL balance."),
        ("Stress case risk rating produced", "='Cash Flow Stress'!D17", "High Risk", 'B9<>""', "Scenario rating formula is active."),
        (
            "Interest fallback documented",
            f"='Historical Financials'!{xl_rowcol_to_cell(hist_row_map['interest'], 5)}",
            0.0,
            "B10>C10",
            "2024-2025 use InterestPaidNet where InterestExpense is unavailable in companyfacts.",
        ),
    ]
    for idx, (check, actual_formula, expected, test, note) in enumerate(check_rows, start=4):
        checks.write(idx, 0, check, fmt["label"])
        actual_value: Any
        if idx == 4:
            actual_value = 1.0
        elif idx == 5:
            actual_value = 1.0
        elif idx == 6:
            actual_value = 1.0
        elif idx == 7:
            actual_value = outputs["Base"]["excess_availability"]
        elif idx == 8:
            actual_value = outputs["Stress"]["rating"]
        else:
            actual_value = hist[2025]["interest"]
        write_formula(checks, idx, 1, actual_formula, fmt["pct"] if isinstance(actual_value, float) and actual_value <= 1.1 else fmt["money"] if isinstance(actual_value, float) else fmt["label"], actual_value)
        checks.write(idx, 2, expected, fmt["pct"] if isinstance(expected, float) and expected == 1.0 else fmt["money"] if isinstance(expected, float) else fmt["label"])
        status_formula = f'=IF({test},"OK","Review")'
        status_value = "OK"
        write_formula(checks, idx, 3, status_formula, fmt["ok"], status_value)
        checks.write(idx, 4, note, fmt["note"])
    checks.write(12, 0, "Overall model status", fmt["section"])
    write_formula(checks, 12, 1, '=IF(COUNTIF(D5:D10,"Review")=0,"OK","Review")', fmt["ok"], "OK")

    workbook.close()


def build_memo(data: dict[str, Any], hist: dict[int, dict[str, float]], outputs: dict[str, dict[str, Any]]) -> None:
    latest_10k = data["as_of"]["latest_10k"]
    latest_10q = data["as_of"]["latest_10q"]
    base = outputs["Base"]
    downside = outputs["Downside"]
    stress = outputs["Stress"]
    h25 = hist[2025]
    memo = f"""# ABL Credit Memo: Genuine Parts Company

**Borrower:** Genuine Parts Company (NYSE: GPC)  
**Industry:** {data['company']['sic_description']}  
**Purpose:** Demonstration ABL credit analysis using public financial statements and illustrative borrowing-base controls.  
**Latest annual source:** {latest_10k['form']} filed {latest_10k['filing_date']} for period ended {latest_10k['report_date']}  
**Latest quarterly context:** {latest_10q['form']} filed {latest_10q['filing_date']} for period ended {latest_10q['report_date']}

## Recommendation

Approve in the base case with conservative advance rates, monthly borrowing-base reporting, AR aging and dilution controls, inventory eligibility monitoring, and a minimum excess-availability covenant. The borrower has meaningful AR and inventory collateral, but downside protection depends on maintaining collateral quality and cash-flow coverage if inventory builds or margin pressure persists.

## Credit Questions Addressed

1. **How much can be lent against working-capital collateral?** Base-case borrowing base is estimated at **{fmt_money(base['borrowing_base'])}**, with excess availability of **{fmt_money(base['excess_availability'])}** after the modeled ABL loan balance.
2. **Can cash flow service debt?** FY2025 reported revenue was **{fmt_money(h25['revenue'])}** and EBITDA proxy was **{fmt_money(h25['ebitda'])}**. Base-case EBITDA / interest is **{base['coverage']:.1f}x**.
3. **What happens if performance weakens?** Stress-case excess availability falls to **{fmt_money(stress['excess_availability'])}**, EBITDA / interest is **{stress['coverage']:.1f}x**, and the model flags **{stress['rating']}**.

## Collateral Support

- FY2025 reported accounts receivable: **{fmt_money(h25['ar'])}**.
- FY2025 reported inventory: **{fmt_money(h25['inventory'])}**.
- Base case applies **{fmt_pct(0.85)}** AR advance rate and **{fmt_pct(0.50)}** inventory advance rate after eligibility haircuts.
- Reserves include dilution, concentration, other collateral reserves, and an availability block.
- Customer concentration, AR aging, and inventory detail are illustrative schedules because public filings do not provide borrowing-base certificate detail.

## Cash Flow And Liquidity

- FY2025 operating cash flow was **{fmt_money(h25['cfo'])}** and free cash flow after capex was **{fmt_money(h25['fcf'])}**.
- FY2025 total debt was **{fmt_money(h25['total_debt'])}** and net debt was **{fmt_money(h25['net_debt'])}**.
- Current ratio was **{h25['current_ratio']:.1f}x** and quick ratio was **{h25['quick_ratio']:.1f}x**.
- Under the downside case, excess availability is **{fmt_money(downside['excess_availability'])}** and EBITDA / interest is **{downside['coverage']:.1f}x**.

## Key Risks

- Inventory-heavy collateral base may be exposed to slow-moving, obsolete, or lower-NOLV inventory in a downturn.
- FY2025 earnings were pressured, reducing cash-flow cushion relative to prior years.
- Higher short-term borrowings and commercial paper increase refinancing and liquidity monitoring importance.
- Public-company data lacks customer-level AR and inventory field-exam detail, so lender diligence would need a real borrowing-base certificate, AR aging, inventory appraisal, and field exam.

## Mitigants And Controls

- Advance rates are conservative and step down in downside/stress cases.
- Borrowing-base reserves explicitly capture dilution, concentration, rent/other reserves, and an availability block.
- Monthly BBC reporting, AR aging, inventory roll-forward, customer concentration reporting, and minimum excess availability would provide early-warning controls.
- Cash-flow stress is reviewed alongside collateral value, consistent with ABL underwriting that considers both repayment capacity and collateral downside protection.
"""
    MEMO_PATH.write_text(memo, encoding="utf-8")


def main() -> None:
    data = load_data()
    hist = hist_values(data)
    inputs = scenario_inputs()
    outputs = scenario_outputs(hist, inputs)
    build_workbook(data, hist, inputs, outputs)
    build_memo(data, hist, outputs)
    print(f"Wrote {WORKBOOK_PATH}")
    print(f"Wrote {MEMO_PATH}")


if __name__ == "__main__":
    main()
