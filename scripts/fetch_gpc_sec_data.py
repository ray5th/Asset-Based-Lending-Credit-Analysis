#!/usr/bin/env python3
"""Fetch Genuine Parts Company financial facts from SEC companyfacts.

The generated JSON/CSV are intentionally compact and auditable. The workbook
builder uses these files as source inputs and keeps ABL assumptions separate.
"""

from __future__ import annotations

import csv
import json
import os
from pathlib import Path
from typing import Any

import certifi
import requests


CIK_INT = 40987
CIK_PADDED = f"{CIK_INT:010d}"
COMPANYFACTS_URL = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK_PADDED}.json"
SUBMISSIONS_URL = f"https://data.sec.gov/submissions/CIK{CIK_PADDED}.json"
SEC_ARCHIVES_BASE = f"https://www.sec.gov/Archives/edgar/data/{CIK_INT}"
USER_AGENT = os.environ.get(
    "SEC_USER_AGENT",
    "Asset-Based-Lending-Credit-Analysis/1.0 contact@example.com",
)

YEARS = [2021, 2022, 2023, 2024, 2025]

DURATION_METRICS = {
    "revenue": {
        "label": "Revenue",
        "tag": "RevenueFromContractWithCustomerExcludingAssessedTax",
    },
    "pretax_income": {
        "label": "Pretax income",
        "tag": "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
    },
    "income_tax": {"label": "Income tax expense", "tag": "IncomeTaxExpenseBenefit"},
    "net_income": {"label": "Net income", "tag": "NetIncomeLoss"},
    "cfo": {
        "label": "Net cash provided by operating activities",
        "tag": "NetCashProvidedByUsedInOperatingActivities",
    },
    "capex": {
        "label": "Capital expenditures",
        "tag": "PaymentsToAcquirePropertyPlantAndEquipment",
    },
    "interest_expense": {"label": "Interest expense", "tag": "InterestExpense"},
    "interest_paid": {"label": "Interest paid", "tag": "InterestPaidNet"},
    "depreciation_amortization": {
        "label": "Depreciation and amortization",
        "tag": "DepreciationAndAmortization",
    },
}

INSTANT_METRICS = {
    "cash": {
        "label": "Cash, cash equivalents and restricted cash",
        "tag": "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
    },
    "accounts_receivable": {
        "label": "Accounts, notes and loans receivable, net",
        "tag": "AccountsNotesAndLoansReceivableNetCurrent",
    },
    "inventory": {"label": "Inventory, net", "tag": "InventoryNet"},
    "current_assets": {"label": "Current assets", "tag": "AssetsCurrent"},
    "current_liabilities": {"label": "Current liabilities", "tag": "LiabilitiesCurrent"},
    "short_term_borrowings": {
        "label": "Short-term borrowings",
        "tag": "ShortTermBorrowings",
        "missing_as_zero": True,
    },
    "commercial_paper": {
        "label": "Commercial paper",
        "tag": "CommercialPaper",
        "missing_as_zero": True,
    },
    "long_term_debt_current": {
        "label": "Current maturities of long-term debt",
        "tag": "LongTermDebtCurrent",
        "missing_as_zero": True,
    },
    "long_term_debt_noncurrent": {
        "label": "Long-term debt, noncurrent",
        "tag": "LongTermDebtNoncurrent",
        "missing_as_zero": True,
    },
}


def get_json(url: str) -> dict[str, Any]:
    response = requests.get(
        url,
        headers={"User-Agent": USER_AGENT, "Accept-Encoding": "gzip, deflate"},
        timeout=45,
        verify=certifi.where(),
    )
    response.raise_for_status()
    return response.json()


def filing_url(accession: str | None, accession_to_doc: dict[str, str]) -> str | None:
    if not accession:
        return None
    primary_doc = accession_to_doc.get(accession)
    if not primary_doc:
        return f"{SEC_ARCHIVES_BASE}/{accession.replace('-', '')}/"
    return f"{SEC_ARCHIVES_BASE}/{accession.replace('-', '')}/{primary_doc}"


def build_accession_maps(submissions: dict[str, Any]) -> tuple[dict[str, str], dict[str, dict[str, str]]]:
    recent = submissions["filings"]["recent"]
    accession_to_doc: dict[str, str] = {}
    accession_meta: dict[str, dict[str, str]] = {}
    for idx, accession in enumerate(recent["accessionNumber"]):
        accession_to_doc[accession] = recent["primaryDocument"][idx]
        accession_meta[accession] = {
            "form": recent["form"][idx],
            "filing_date": recent["filingDate"][idx],
            "report_date": recent["reportDate"][idx],
            "primary_document": recent["primaryDocument"][idx],
        }
    return accession_to_doc, accession_meta


def choose_fact(
    facts: dict[str, Any],
    tag: str,
    year: int,
    *,
    duration: bool,
    missing_as_zero: bool = False,
) -> dict[str, Any] | None:
    fact = facts.get(tag)
    if not fact:
        if missing_as_zero:
            return {"val": 0, "tag": tag, "unit": "USD", "synthetic_zero": True}
        return None

    unit = "USD" if "USD" in fact.get("units", {}) else next(iter(fact.get("units", {})))
    rows = fact["units"][unit]
    end_date = f"{year}-12-31"
    frame = f"CY{year}" if duration else f"CY{year}Q4I"

    candidates = [
        row
        for row in rows
        if row.get("form") == "10-K" and row.get("fp") == "FY" and row.get("end") == end_date
    ]
    if duration:
        candidates = [row for row in candidates if row.get("start", "").startswith(f"{year}-01-")]

    if not candidates:
        if missing_as_zero:
            return {"val": 0, "tag": tag, "unit": unit, "synthetic_zero": True}
        return None

    def score(row: dict[str, Any]) -> tuple[int, str]:
        frame_score = 2 if row.get("frame") == frame else 1 if row.get("frame") else 0
        return (frame_score, row.get("filed", ""))

    chosen = sorted(candidates, key=score)[-1].copy()
    chosen["tag"] = tag
    chosen["unit"] = unit
    return chosen


def enrich_fact(
    raw: dict[str, Any] | None,
    accession_to_doc: dict[str, str],
    accession_meta: dict[str, dict[str, str]],
) -> dict[str, Any]:
    if raw is None:
        return {"value": None, "source": None}

    accession = raw.get("accn")
    source = {
        "tag": raw.get("tag"),
        "unit": raw.get("unit", "USD"),
        "accession": accession,
        "filed": raw.get("filed"),
        "fy": raw.get("fy"),
        "fp": raw.get("fp"),
        "frame": raw.get("frame"),
        "start": raw.get("start"),
        "end": raw.get("end"),
        "synthetic_zero": bool(raw.get("synthetic_zero")),
        "url": filing_url(accession, accession_to_doc),
    }
    if accession in accession_meta:
        source.update(accession_meta[accession])
    return {"value": raw.get("val"), "source": source}


def main() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    data_dir = repo_root / "data"
    data_dir.mkdir(parents=True, exist_ok=True)

    submissions = get_json(SUBMISSIONS_URL)
    companyfacts = get_json(COMPANYFACTS_URL)
    facts = companyfacts["facts"]["us-gaap"]
    accession_to_doc, accession_meta = build_accession_maps(submissions)

    filings = []
    recent = submissions["filings"]["recent"]
    for idx, form in enumerate(recent["form"]):
        if form in {"10-K", "10-Q"}:
            accession = recent["accessionNumber"][idx]
            filings.append(
                {
                    "form": form,
                    "filing_date": recent["filingDate"][idx],
                    "report_date": recent["reportDate"][idx],
                    "accession": accession,
                    "primary_document": recent["primaryDocument"][idx],
                    "url": filing_url(accession, accession_to_doc),
                }
            )

    metrics: dict[str, dict[str, Any]] = {}
    for metric_key, metric in DURATION_METRICS.items():
        metrics[metric_key] = {"label": metric["label"], "tag": metric["tag"], "values": {}}
        for year in YEARS:
            raw = choose_fact(facts, metric["tag"], year, duration=True)
            metrics[metric_key]["values"][str(year)] = enrich_fact(raw, accession_to_doc, accession_meta)

    for metric_key, metric in INSTANT_METRICS.items():
        metrics[metric_key] = {"label": metric["label"], "tag": metric["tag"], "values": {}}
        for year in YEARS:
            raw = choose_fact(
                facts,
                metric["tag"],
                year,
                duration=False,
                missing_as_zero=metric.get("missing_as_zero", False),
            )
            metrics[metric_key]["values"][str(year)] = enrich_fact(raw, accession_to_doc, accession_meta)

    output = {
        "company": {
            "name": submissions["name"],
            "ticker": "GPC",
            "cik": CIK_PADDED,
            "sic": submissions.get("sic"),
            "sic_description": submissions.get("sicDescription"),
            "fiscal_year_end": submissions.get("fiscalYearEnd"),
        },
        "as_of": {
            "history_years": YEARS,
            "latest_10k": next((filing for filing in filings if filing["form"] == "10-K"), None),
            "latest_10q": next((filing for filing in filings if filing["form"] == "10-Q"), None),
            "companyfacts_url": COMPANYFACTS_URL,
            "submissions_url": SUBMISSIONS_URL,
            "generated_from": "SEC companyfacts and submissions APIs",
        },
        "metrics": metrics,
    }

    json_path = data_dir / "gpc_sec_financials.json"
    json_path.write_text(json.dumps(output, indent=2), encoding="utf-8")

    csv_path = data_dir / "gpc_historical_financials.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["metric_key", "metric_label", "tag", *YEARS])
        for metric_key, metric in metrics.items():
            writer.writerow(
                [
                    metric_key,
                    metric["label"],
                    metric["tag"],
                    *[metric["values"][str(year)]["value"] for year in YEARS],
                ]
            )

    print(f"Wrote {json_path}")
    print(f"Wrote {csv_path}")
    print(f"Latest 10-K: {output['as_of']['latest_10k']}")
    print(f"Latest 10-Q: {output['as_of']['latest_10q']}")


if __name__ == "__main__":
    main()
