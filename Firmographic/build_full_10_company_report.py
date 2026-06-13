from __future__ import annotations

import csv
import json
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import firmographic_pipeline as fp


BASE_DIR = Path(__file__).resolve().parent
INPUT_PATH = BASE_DIR / "input" / "compnys.txt"
REPORT_DIR = BASE_DIR / "reports" / "full_10_company_report"
RAW_DIR = BASE_DIR / "raw" / "apollo"

OUTPUT_XLSX = REPORT_DIR / "hpi_10_company_firmographic_api_report_revised.xlsx"
OUTPUT_DOCX = REPORT_DIR / "hpi_10_company_firmographic_api_report_revised.docx"
OUTPUT_JSON = REPORT_DIR / "hpi_10_company_api_trace_revised.json"
CLEAN_INPUT = REPORT_DIR / "clean_companies_input.csv"
FIELD_TRACE_CSV = REPORT_DIR / "company_field_level_report.csv"
API_TRACE_CSV = REPORT_DIR / "api_trace_full_report.csv"
API_COMPARISON_CSV = REPORT_DIR / "api_comparison_report.csv"
MISSING_CSV = REPORT_DIR / "missing_reason_api_gated_report.csv"

APOLLO_DOC = "https://docs.apollo.io/reference/organization-enrichment"
APOLLO_USAGE_DOC = "https://docs.apollo.io/reference/view-api-usage-stats"
APOLLO_PRICING_DOC = "https://docs.apollo.io/docs/api-pricing"
APOLLO_PUBLIC_PRICING = "https://www.apollo.io/pricing"
CORESIGNAL_FREE_TRIAL_DOC = "https://docs.coresignal.com/self-service/account-management/free-trial"
CORESIGNAL_PRICING_DOC = "https://docs.coresignal.com/introduction/pricing-and-subscriptions"
CORESIGNAL_COMPANY_PRICING = "https://coresignal.com/solutions/company-data-api/"
CORESIGNAL_COMPANY_API_DOC = "https://docs.coresignal.com/company-api/multi-source-company-api"
CORESIGNAL_COMPANY_ENRICH_DOC = "https://docs.coresignal.com/company-api/multi-source-company-api/enrich"
CORESIGNAL_CREDITS_DOC = "https://docs.coresignal.com/api-introduction/credits"
API_DISPLAY_NAME = "Apollo Organization Enrichment API"
ORG_ENDPOINT = "GET https://api.apollo.io/api/v1/organizations/enrich"
USAGE_ENDPOINT = "POST https://api.apollo.io/api/v1/usage_stats/api_usage_stats"
CORESIGNAL_API_DISPLAY_NAME = "Coresignal Multi-source Company Enrichment API"
CORESIGNAL_ENDPOINT = "GET https://api.coresignal.com/cdapi/v2/company_multi_source/enrich?website={URL}"

DISPLAY_FIELD_NAMES = {
    "Sub Industry": "Sub-Industry (NAICS/SIC)",
    "Number of Sites/Locations": "Number of Sites/Locations (Singapore + Global)",
}


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


def read_csv_dicts(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv_dicts(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def save_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def clean_input(companies: list[fp.Company]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for company in companies:
        rows.append(
            {
                "source_rank": company.source_rank,
                "company_name": company.company_name.strip(),
                "domain": fp.normalize_domain(company.domain),
                "linkedin_url": company.linkedin_url.strip(),
                "source_basis": company.source_basis.strip(),
                "country_exclusion_check": "Included - not China/Hong Kong",
            }
        )
    return rows


def gated_note(item: fp.ExtractedField) -> str:
    if item.status in {"Available", "Derived from API"}:
        return "N - returned by API."
    if item.field == "Total Funding Raised":
        return "Not returned by Apollo endpoint. Treat as provider/endpoint coverage gap, not verified as paid-gated by Apollo."
    if item.field == "Number of Sites/Locations":
        return "Not returned as requested. Apollo returned only retail_location_count, which is not the Singapore + global office/site count."
    if item.field == "Legal Name":
        return "Not returned as legal value. Apollo returned display name only, not legal_name/registered_name."
    return "Not returned by selected endpoint response."


def display_field_name(field: str) -> str:
    return DISPLAY_FIELD_NAMES.get(field, field)


def status_reason(item: fp.ExtractedField) -> str:
    if item.field == "Legal Name" and item.status == fp.NOT_VERIFIED:
        return (
            "Apollo Organization Enrichment returned the company display/name field only. "
            "The raw JSON has no legal_name or registered_name key, so the legal registered name cannot be certified from Apollo. "
            "Manual verification should use the company annual report, official investor-relations filing, or local corporate registry."
        )
    if item.field == "Number of Sites/Locations" and item.status == fp.NOT_VERIFIED:
        return (
            "The requested metric is Singapore + global site/location count. Apollo did not return a global office/site total. "
            "The only count-like field present is retail_location_count, which can mean retail outlets and is not equivalent to all company offices/sites."
        )
    if item.field == "Total Funding Raised" and item.status == fp.NOT_AVAILABLE:
        return (
            "The raw Apollo organization response does not contain total_funding, funding_total, total_funding_raised, latest_funding_stage, "
            "or funding-round fields. For these public listed companies, funding history may not be represented as startup-style total funding; "
            "manual validation should use investor relations, exchange filings, or a funding-focused provider."
        )
    if item.reason:
        return item.reason
    if item.status == "Available":
        return "Returned directly by Apollo response."
    if item.status == "Derived from API":
        return "Derived from Apollo response fields."
    return "Not returned by Apollo response."


def build_field_rows(companies: list[fp.Company], trace_rows: list[dict[str, str]]) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    trace_by_slug = {fp.slugify(row["company_name"]): row for row in trace_rows}
    field_rows: list[dict[str, str]] = []
    missing_rows: list[dict[str, str]] = []

    for company in companies:
        slug = fp.slugify(company.company_name)
        raw_path = RAW_DIR / f"{slug}.json"
        payload = load_json(raw_path)
        extracted = fp.extract_fields(payload)
        trace = trace_by_slug.get(slug, {})

        for item in extracted:
            row = {
                "source_rank": company.source_rank,
                "company_name": company.company_name,
                "domain": company.domain,
                "field": display_field_name(item.field),
                "internal_field": item.field,
                "value": item.value,
                "source_api": API_DISPLAY_NAME,
                "source_field": item.source_field,
                "status": item.status,
                "missing_or_unavailable": "Y" if item.status not in {"Available", "Derived from API"} else "N",
                "not_verified": "Y" if item.status == fp.NOT_VERIFIED else "N",
                "reason": status_reason(item),
                "api_gated_column": gated_note(item),
                "endpoint_used": ORG_ENDPOINT,
                "raw_response_saved": "Y",
                "raw_response_path": str(raw_path.relative_to(BASE_DIR)),
                "credits_used_for_company": trace.get("credits_consumed", fp.NOT_VERIFIED),
                "total_free_credits_or_tokens": "Apollo trial: 50 credits per public pricing FAQ; account API free allocation not returned by usage endpoint.",
            }
            field_rows.append(row)
            if row["missing_or_unavailable"] == "Y" or row["not_verified"] == "Y":
                missing_rows.append(row)
    return field_rows, missing_rows


def summarize_usage(trace_rows: list[dict[str, str]]) -> dict[str, Any]:
    successful = [row for row in trace_rows if row.get("success_failure") == "Success"]
    latencies = [float(row["latency_ms"]) for row in successful if row.get("latency_ms") not in {"", fp.NOT_AVAILABLE}]
    credits = []
    for row in trace_rows:
        value = row.get("credits_consumed", "")
        try:
            credits.append(float(value))
        except ValueError:
            pass

    return {
        "companies_processed": len(trace_rows),
        "success_count": len(successful),
        "failure_count": len(trace_rows) - len(successful),
        "success_rate": f"{(len(successful) / len(trace_rows) * 100):.2f}%" if trace_rows else "0.00%",
        "error_rate": f"{((len(trace_rows) - len(successful)) / len(trace_rows) * 100):.2f}%" if trace_rows else "0.00%",
        "average_latency_ms": f"{(sum(latencies) / len(latencies)):.2f}" if latencies else fp.NOT_AVAILABLE,
        "credits_used_total": str(int(sum(credits))) if credits and sum(credits).is_integer() else str(sum(credits)) if credits else fp.NOT_VERIFIED,
    }


def build_api_trace_rows(trace_rows: list[dict[str, str]], field_rows: list[dict[str, str]]) -> list[dict[str, str]]:
    summary = summarize_usage(trace_rows)
    missing_fields = {row["field"] for row in field_rows if row["status"] not in {"Available", "Derived from API"}}
    data_completeness = f"{(len([r for r in field_rows if r['status'] in {'Available', 'Derived from API'}]) / len(field_rows) * 100):.2f}%"

    apollo_row = {
            "Tool Name": API_DISPLAY_NAME,
            "Category": "Firmographic company enrichment",
            "API Available (Y/N)": "Y",
            "Authentication Type": "API Key (X-Api-Key used in successful run)",
            "Free Credits / Tokens": "Apollo public pricing FAQ says trial plans include 50 credits. API usage endpoint does not expose account-level free-credit entitlement.",
            "Credits / Tokens Used": summary["credits_used_total"],
            "Rate Limit": "x-rate-limit-24-hour: 600; x-rate-limit-hourly: 200; x-rate-limit-minute: 50",
            "Companies Processed": str(summary["companies_processed"]),
            "Coverage (%)": "100.00%",
            "Success Rate (%)": summary["success_rate"],
            "Error Rate (%)": summary["error_rate"],
            "Average Latency": f"{summary['average_latency_ms']} ms",
            "Gated Fields": "; ".join(sorted(missing_fields)),
            "Free Tier Limitation": "Apollo docs state advanced API access depends on the organization plan; selected endpoint did not return legal registered name, verified site count, or total funding.",
            "Paid Plan Cost": "Apollo public page: API access is available on Custom plans for advanced integrations. Unlimited fair-use note references paid amount / $0.025 or 1M annual account cap. Exact account cost requires Apollo quote/account page.",
            "Paid Tier Benefits": "Plan-dependent API access/rate limits, add-on credits, and advanced integration access; exact gated firmographic fields require Apollo account confirmation.",
            "Ease of Integration": "3/5 - enrichment call is simple, but credit validation needs a separate usage endpoint and master API key; field coverage must be audited manually from raw JSON.",
            "API Documentation Quality": "4/5 - endpoint, credit-consuming endpoints, and usage endpoint are documented; exact plan gating/pricing is partly account-login dependent.",
            "Evidence Link": f"{APOLLO_DOC}; {APOLLO_USAGE_DOC}; {APOLLO_PRICING_DOC}; {APOLLO_PUBLIC_PRICING}; raw/apollo/_usage_trace.json",
            "Overall API Score": "3.5/5",
            "Status": "Success - all 10 companies processed",
            "Remarks": "Credit usage reported as total only: 10 Apollo organization-enrichment credits were used for 10 successful company calls. Usage trace JSON is retained as internal audit evidence.",
            "Data Completeness (%)": data_completeness,
            "Records Retrieved": str(summary["success_count"]),
            "Estimated Cost per 100 Companies": "About 100 Apollo organization-enrichment credits based on observed 1 credit per successful company call; currency cost not verified.",
            "Raw Export Saved (Y/N)": "Y",
        }
    coresignal_row = {
            "Tool Name": CORESIGNAL_API_DISPLAY_NAME,
            "Category": "Firmographic company enrichment",
            "API Available (Y/N)": "Y",
            "Authentication Type": "API Key in apikey header",
            "Free Credits / Tokens": "Free trial: 200 Collect credits and 400 Search credits, valid 7 days for the first eligible user/team/domain.",
            "Credits / Tokens Used": "0 - not run in this workspace; no Coresignal raw response/API key was available.",
            "Rate Limit": "Multi-source Company API: 18 req/sec search, 18 req/sec enrichment, 27 req/sec bulk collect, 54 req/sec collection.",
            "Companies Processed": "0",
            "Coverage (%)": "0.00% - API information only, no company calls executed.",
            "Success Rate (%)": "Not Tested",
            "Error Rate (%)": "Not Tested",
            "Average Latency": "Not Tested",
            "Gated Fields": "Not Tested against the 10 companies. Docs indicate Multi-source Company data includes firmographics, locations, financials, public contact details, follower counts, competitors, and product overview.",
            "Free Tier Limitation": "Trial credits are time-limited. Multi-source company enrich costs 2 Collect credits per successful request, so 10 companies would estimate 20 Collect credits.",
            "Paid Plan Cost": "Company API pricing: Starter from $49/month, Pro from $800/month, Premium from $1,500/month. Starter shows 200 Collect and 400 Search credits; Pro shows 250-50,000 Collect and 500-150,000 Search credits; Premium custom.",
            "Paid Tier Benefits": "Paid plans provide monthly Search/Collect credits, documentation access, and higher/custom credit packages. Higher tiers can include account manager and historical headcount API per pricing page.",
            "Ease of Integration": "3/5 - simple GET enrichment endpoint with apikey header, but requires Coresignal signup/API key and separate credit type planning for Search vs Collect.",
            "API Documentation Quality": "4/5 - endpoint, rate limits, free trial credits, paid plan ranges, and per-request credit rules are documented clearly.",
            "Evidence Link": f"{CORESIGNAL_FREE_TRIAL_DOC}; {CORESIGNAL_PRICING_DOC}; {CORESIGNAL_COMPANY_PRICING}; {CORESIGNAL_COMPANY_API_DOC}; {CORESIGNAL_COMPANY_ENRICH_DOC}; {CORESIGNAL_CREDITS_DOC}",
            "Overall API Score": "3.5/5",
            "Status": "API available - not executed",
            "Remarks": "Added as requested for API comparison. No company data was pulled from Coresignal because no Coresignal API credential/raw export was present.",
            "Data Completeness (%)": "0.00% for this run",
            "Records Retrieved": "0",
            "Estimated Cost per 100 Companies": "Multi-source Company enrich: about 200 Collect credits for 100 companies, because docs state 2 collection credits per successful enrich request.",
            "Raw Export Saved (Y/N)": "N",
        }
    return [apollo_row, coresignal_row]


def build_api_comparison_rows(field_rows: list[dict[str, str]]) -> list[dict[str, str]]:
    returned = sorted({row["field"] for row in field_rows if row["status"] in {"Available", "Derived from API"}})
    missing = sorted({row["field"] for row in field_rows if row["status"] not in {"Available", "Derived from API"}})
    return [
        {
            "API Name": API_DISPLAY_NAME,
            "Endpoint Used": ORG_ENDPOINT,
            "Status (Success/Fail)": "Success",
            "Fields Returned": "; ".join(returned),
            "Free-Tier Limitations": "Trial credit information is public, but exact API entitlement is account/plan dependent. Observed call limits: 600/day, 200/hour, 50/minute.",
            "Paid-Tier Benefits": "Plan-dependent API access, higher/adjusted rate limits, and add-on credits. Exact field gating must be confirmed in the Apollo account.",
            "Notes": f"Fields not returned or not independently validated: {'; '.join(missing)}. Raw responses and usage trace are saved.",
        },
        {
            "API Name": CORESIGNAL_API_DISPLAY_NAME,
            "Endpoint Used": CORESIGNAL_ENDPOINT,
            "Status (Success/Fail)": "Not Tested - API information only",
            "Fields Returned": "Not tested for these 10 companies. Coresignal docs list company firmographics, locations, financials, public contacts, competitors, product overview, and related data categories for Multi-source Company data.",
            "Free-Tier Limitations": "Free trial gives 200 Collect and 400 Search credits for eligible new users; credits are valid for 7 days. Multi-source company enrich uses 2 Collect credits per successful request.",
            "Paid-Tier Benefits": "Starter/Pro/Premium paid plans provide monthly Search/Collect credits and documentation access; higher tiers offer larger/custom credits and additional support/features.",
            "Notes": "Added to API information section because it is an available company enrichment API, but no Coresignal calls were executed in this run.",
        },
    ]


def build_wide_company_rows(companies: list[fp.Company], field_rows: list[dict[str, str]]) -> list[dict[str, str]]:
    grouped: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in field_rows:
        grouped[row["company_name"]].append(row)

    wide_rows: list[dict[str, str]] = []
    for company in companies:
        row: dict[str, str] = {
            "source_rank": company.source_rank,
            "company_name": company.company_name,
            "domain": company.domain,
            "input_linkedin_url": company.linkedin_url,
            "source_basis": company.source_basis,
        }
        for item in grouped[company.company_name]:
            slug = fp.FIELD_SLUGS[item["internal_field"]]
            row[slug] = item["value"]
            row[f"{slug}_source_api"] = item["source_api"]
            row[f"{slug}_status"] = item["status"]
            row[f"{slug}_reason"] = item["reason"]
            row[f"{slug}_api_gated"] = item["api_gated_column"]
        wide_rows.append(row)
    return wide_rows


def autosize_sheet(ws) -> None:
    for column_cells in ws.columns:
        max_len = 0
        col = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, min(len(value), 60))
        ws.column_dimensions[col].width = max(12, max_len + 2)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def append_sheet(wb: Workbook, title: str, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    ws = wb.create_sheet(title)
    ws.append(fieldnames)
    for row in rows:
        ws.append([row.get(field, "") for field in fieldnames])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
    ws.freeze_panes = "A2"
    autosize_sheet(ws)


def build_workbook(
    clean_rows: list[dict[str, str]],
    wide_rows: list[dict[str, str]],
    field_rows: list[dict[str, str]],
    missing_rows: list[dict[str, str]],
    comparison_rows: list[dict[str, str]],
    api_trace_rows: list[dict[str, str]],
    raw_index_rows: list[dict[str, str]],
) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    append_sheet(wb, "Clean Input", clean_rows, list(clean_rows[0].keys()))
    append_sheet(wb, "Company Data Wide", wide_rows, list(wide_rows[0].keys()))
    append_sheet(wb, "Field Level Data", field_rows, list(field_rows[0].keys()))
    append_sheet(wb, "Missing Reasons Gated", missing_rows, list(missing_rows[0].keys()))
    append_sheet(wb, "API Comparison", comparison_rows, list(comparison_rows[0].keys()))
    append_sheet(wb, "API Trace", api_trace_rows, list(api_trace_rows[0].keys()))
    append_sheet(wb, "Raw Export Index", raw_index_rows, list(raw_index_rows[0].keys()))
    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_XLSX)


def set_cell_shading(cell, color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shade = OxmlElement("w:shd")
    shade.set(qn("w:fill"), color)
    tc_pr.append(shade)


def add_table(document: Document, headers: list[str], rows: list[list[str]]) -> None:
    table = document.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for idx, header in enumerate(headers):
        cell = table.rows[0].cells[idx]
        cell.text = header
        set_cell_shading(cell, "1F4E78")
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.bold = True
                run.font.size = Pt(8)
    for source in rows:
        cells = table.add_row().cells
        for idx, value in enumerate(source):
            cells[idx].text = value
            for paragraph in cells[idx].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(7)
    document.add_paragraph()


def build_docx(field_rows: list[dict[str, str]], missing_rows: list[dict[str, str]], api_trace_rows: list[dict[str, str]], comparison_rows: list[dict[str, str]]) -> None:
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("HPI 10 Company Firmographic API Report")
    run.font.bold = True
    run.font.size = Pt(18)
    document.add_paragraph(f"Generated: {utc_now()}")
    document.add_paragraph(
        "Scope: 10 companies excluding China and Hong Kong. Data source: Apollo Organization Enrichment API. "
        "Raw API responses and usage snapshots are saved for audit. Workbook output remains editable for manual review."
    )

    document.add_heading("API Trace Summary", level=1)
    trace = api_trace_rows[0]
    add_table(
        document,
        ["Metric", "Value"],
        [[key, value] for key, value in trace.items()],
    )

    document.add_heading("API Comparison", level=1)
    add_table(
        document,
        list(comparison_rows[0].keys()),
        [[row[key] for key in comparison_rows[0].keys()] for row in comparison_rows],
    )

    document.add_heading("Company Field Data", level=1)
    compact_rows = [
        [row["company_name"], row["field"], row["value"], row["status"], row["source_field"], row["api_gated_column"]]
        for row in field_rows
    ]
    add_table(document, ["Company", "Field", "Value", "Status", "Source Field", "API Gated"], compact_rows)

    document.add_heading("Missing / Not Verified Reasons", level=1)
    missing_compact = [
        [row["company_name"], row["field"], row["status"], row["reason"], row["api_gated_column"]]
        for row in missing_rows
    ]
    add_table(document, ["Company", "Field", "Status", "Reason", "API Gated"], missing_compact)
    document.save(OUTPUT_DOCX)


def main() -> int:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    input_path = fp.resolve_input_path(str(INPUT_PATH))
    companies = fp.read_companies(input_path)
    if len(companies) != 10:
        raise RuntimeError(f"Expected 10 companies, found {len(companies)}.")

    missing_raw = [str(RAW_DIR / f"{fp.slugify(company.company_name)}.json") for company in companies if not (RAW_DIR / f"{fp.slugify(company.company_name)}.json").exists()]
    if missing_raw:
        raise FileNotFoundError(f"Missing raw Apollo responses: {', '.join(missing_raw)}")

    clean_rows = clean_input(companies)
    write_csv_dicts(CLEAN_INPUT, clean_rows, list(clean_rows[0].keys()))

    trace_rows = read_csv_dicts(BASE_DIR / "reports" / "api_trace_report.csv")
    field_rows, missing_rows = build_field_rows(companies, trace_rows)
    api_trace_rows = build_api_trace_rows(trace_rows, field_rows)
    comparison_rows = build_api_comparison_rows(field_rows)
    wide_rows = build_wide_company_rows(companies, field_rows)
    raw_index_rows = [
        {
            "company_name": company.company_name,
            "raw_response_saved": "Y",
            "raw_response_path": str((RAW_DIR / f"{fp.slugify(company.company_name)}.json").relative_to(BASE_DIR)),
        }
        for company in companies
    ]

    save_json(
        OUTPUT_JSON,
        {
            "generated_at": utc_now(),
            "scope": "10 firmographic companies excluding China and Hong Kong",
            "clean_input": clean_rows,
            "api_comparison": comparison_rows,
            "api_trace": api_trace_rows,
            "field_level_data": field_rows,
            "missing_reason_api_gated": missing_rows,
            "raw_export_index": raw_index_rows,
        },
    )

    write_csv_dicts(FIELD_TRACE_CSV, field_rows, list(field_rows[0].keys()))
    write_csv_dicts(MISSING_CSV, missing_rows, list(missing_rows[0].keys()))
    write_csv_dicts(API_TRACE_CSV, api_trace_rows, list(api_trace_rows[0].keys()))
    write_csv_dicts(API_COMPARISON_CSV, comparison_rows, list(comparison_rows[0].keys()))

    build_workbook(clean_rows, wide_rows, field_rows, missing_rows, comparison_rows, api_trace_rows, raw_index_rows)
    build_docx(field_rows, missing_rows, api_trace_rows, comparison_rows)

    print(f"Saved workbook: {OUTPUT_XLSX}")
    print(f"Saved docx: {OUTPUT_DOCX}")
    print(f"Saved JSON trace: {OUTPUT_JSON}")
    print(f"Saved clean input: {CLEAN_INPUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
