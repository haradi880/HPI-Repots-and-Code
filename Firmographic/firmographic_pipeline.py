from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests


BASE_DIR = Path(__file__).resolve().parent
INPUT_CANDIDATES = [
    BASE_DIR / "input" / "compnys.txt",
    BASE_DIR / "input" / "companies.csv",
    BASE_DIR / "input" / "companies.txt",
]
REPORTS_DIR = BASE_DIR / "reports"
API_NAME = "apollo"
RAW_DIR = BASE_DIR / "raw" / API_NAME
USAGE_TRACE_PATH = RAW_DIR / "_usage_trace.json"

APOLLO_BASE_URL = "https://api.apollo.io/api/v1"
ORG_ENRICH_URL = f"{APOLLO_BASE_URL}/organizations/enrich"
USAGE_STATS_URL = f"{APOLLO_BASE_URL}/usage_stats/api_usage_stats"

NOT_AVAILABLE = "Not Available"
NOT_VERIFIED = "Not Verified"

REQUIRED_INPUT_COLUMNS = ["source_rank", "company_name", "domain", "linkedin_url", "source_basis"]

FIELD_NAMES = [
    "Legal Name",
    "Website Domain",
    "LinkedIn URL",
    "Employee Count",
    "Headcount Growth (1 Year)",
    "Revenue / Revenue Band",
    "Industry",
    "Sub Industry",
    "Headquarters Location",
    "Number of Sites/Locations",
    "Founded Year",
    "Ownership Type",
    "Total Funding Raised",
]

FIELD_SLUGS = {
    "Legal Name": "legal_name",
    "Website Domain": "website_domain",
    "LinkedIn URL": "linkedin_url",
    "Employee Count": "employee_count",
    "Headcount Growth (1 Year)": "headcount_growth_1_year",
    "Revenue / Revenue Band": "revenue_revenue_band",
    "Industry": "industry",
    "Sub Industry": "sub_industry",
    "Headquarters Location": "headquarters_location",
    "Number of Sites/Locations": "number_of_sites_locations",
    "Founded Year": "founded_year",
    "Ownership Type": "ownership_type",
    "Total Funding Raised": "total_funding_raised",
}


@dataclass
class Company:
    source_rank: str
    company_name: str
    domain: str
    linkedin_url: str
    source_basis: str


@dataclass
class ExtractedField:
    field: str
    value: str
    source_field: str
    status: str
    reason: str


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def load_dotenv() -> None:
    env_path = BASE_DIR / ".env"
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def ensure_dirs() -> None:
    REPORTS_DIR.mkdir(exist_ok=True)
    RAW_DIR.mkdir(parents=True, exist_ok=True)


def resolve_input_path(input_arg: str | None) -> Path:
    if input_arg:
        path = Path(input_arg)
        if not path.is_absolute():
            path = BASE_DIR / path
        if not path.exists():
            raise FileNotFoundError(f"Input file not found: {path}")
        return path

    for candidate in INPUT_CANDIDATES:
        if candidate.exists():
            return candidate
    raise FileNotFoundError("No input file found. Expected input/compnys.txt or input/companies.csv.")


def read_companies(path: Path) -> list[Company]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        missing = [column for column in REQUIRED_INPUT_COLUMNS if column not in (reader.fieldnames or [])]
        if missing:
            raise ValueError(f"Input file is missing columns: {', '.join(missing)}")

        companies: list[Company] = []
        for row in reader:
            if not row or not (row.get("company_name") or "").strip():
                continue
            companies.append(
                Company(
                    source_rank=(row.get("source_rank") or "").strip(),
                    company_name=(row.get("company_name") or "").strip(),
                    domain=normalize_domain((row.get("domain") or "").strip()),
                    linkedin_url=(row.get("linkedin_url") or "").strip(),
                    source_basis=(row.get("source_basis") or "").strip(),
                )
            )
    return companies


def normalize_domain(domain: str) -> str:
    domain = domain.strip()
    domain = re.sub(r"^https?://", "", domain, flags=re.IGNORECASE)
    domain = domain.split("/")[0]
    if domain.lower().startswith("www."):
        domain = domain[4:]
    return domain


def slugify(text: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_")
    return slug or "company"


def save_json(path: Path, data: Any) -> None:
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def safe_headers(headers: requests.structures.CaseInsensitiveDict[str]) -> dict[str, str]:
    allowed: dict[str, str] = {}
    for key, value in headers.items():
        lowered = key.lower()
        if lowered in {"date", "content-type"} or "rate" in lowered or "retry-after" in lowered:
            allowed[key] = value
    return allowed


def auth_headers(api_key: str, mode: str) -> dict[str, str]:
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Cache-Control": "no-cache",
    }
    if mode == "x-api-key":
        headers["X-Api-Key"] = api_key
    else:
        headers["Authorization"] = f"Bearer {api_key}"
    return headers


def api_request(
    method: str,
    url: str,
    api_key: str,
    *,
    params: dict[str, Any] | None = None,
    body: dict[str, Any] | None = None,
    timeout: int = 60,
) -> dict[str, Any]:
    preferred = os.getenv("APOLLO_AUTH_MODE", "").strip().lower()
    modes = [preferred] if preferred in {"bearer", "x-api-key"} else ["bearer", "x-api-key"]
    last_result: dict[str, Any] | None = None

    for mode in modes:
        started_at = now_iso()
        start = time.perf_counter()
        try:
            response = requests.request(
                method,
                url,
                headers=auth_headers(api_key, mode),
                params=params,
                json=body,
                timeout=timeout,
            )
            latency_ms = round((time.perf_counter() - start) * 1000, 2)
            response_received_at = now_iso()
            try:
                parsed = response.json()
                text = None
            except ValueError:
                parsed = None
                text = response.text

            result = {
                "ok": 200 <= response.status_code < 300,
                "status_code": response.status_code,
                "auth_mode": mode,
                "started_at": started_at,
                "response_received_at": response_received_at,
                "latency_ms": latency_ms,
                "headers": safe_headers(response.headers),
                "json": parsed,
                "text": text,
                "error": "",
            }
        except requests.RequestException as exc:
            result = {
                "ok": False,
                "status_code": None,
                "auth_mode": mode,
                "started_at": started_at,
                "response_received_at": now_iso(),
                "latency_ms": round((time.perf_counter() - start) * 1000, 2),
                "headers": {},
                "json": None,
                "text": None,
                "error": str(exc),
            }

        last_result = result
        if result["status_code"] != 401:
            return result

    return last_result or {
        "ok": False,
        "status_code": None,
        "auth_mode": NOT_AVAILABLE,
        "started_at": now_iso(),
        "response_received_at": now_iso(),
        "latency_ms": 0,
        "headers": {},
        "json": None,
        "text": None,
        "error": "No request attempted.",
    }


def get_usage_snapshot(api_key: str) -> dict[str, Any]:
    result = api_request("POST", USAGE_STATS_URL, api_key, body={})
    result["endpoint"] = USAGE_STATS_URL
    return result


def build_params(company: Company) -> dict[str, str]:
    params: dict[str, str] = {}
    if company.domain:
        params["domain"] = company.domain
        params["website"] = f"https://{company.domain}"
    if company.linkedin_url:
        params["linkedin_url"] = company.linkedin_url
    if company.company_name:
        params["name"] = company.company_name
    return params


def get_path(data: Any, path: str) -> Any:
    current = data
    for part in path.split("."):
        if isinstance(current, dict) and part in current:
            current = current[part]
        elif isinstance(current, list) and part.isdigit() and int(part) < len(current):
            current = current[int(part)]
        else:
            return None
    return current


def is_missing(value: Any) -> bool:
    return value is None or value == "" or value == [] or value == {}


def extract_organization(payload: Any) -> tuple[dict[str, Any] | None, str]:
    if not isinstance(payload, dict):
        return None, ""
    for key in ("organization", "company", "account"):
        value = payload.get(key)
        if isinstance(value, dict):
            return value, key
    organizations = payload.get("organizations")
    if isinstance(organizations, list) and organizations and isinstance(organizations[0], dict):
        return organizations[0], "organizations.0"
    if any(key in payload for key in ("name", "primary_domain", "website_url")):
        return payload, ""
    return None, ""


def source_path(prefix: str, path: str) -> str:
    return f"{prefix}.{path}" if prefix else path


def first_value(org: dict[str, Any] | None, prefix: str, paths: list[str]) -> tuple[Any, str] | None:
    if org is None:
        return None
    for path in paths:
        value = get_path(org, path)
        if not is_missing(value):
            return value, source_path(prefix, path)
    return None


def format_scalar(value: Any) -> str:
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, list):
        return ", ".join(format_scalar(item) for item in value)
    return str(value)


def format_money(value: Any) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(int(value)) if float(value).is_integer() else str(value)
    return format_scalar(value)


def format_growth(value: Any) -> str:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        percent = value * 100 if -1 < value < 1 else value
        return f"{percent:.2f}%"
    return format_scalar(value)


def available(field: str, value: Any, source: str, formatter=format_scalar) -> ExtractedField:
    return ExtractedField(field, formatter(value), source, "Available", "")


def missing(field: str, reason: str = "Field absent in Apollo response.") -> ExtractedField:
    return ExtractedField(field, NOT_AVAILABLE, NOT_AVAILABLE, NOT_AVAILABLE, reason)


def unverified(field: str, source: str, reason: str) -> ExtractedField:
    return ExtractedField(field, NOT_VERIFIED, source or NOT_AVAILABLE, NOT_VERIFIED, reason)


def extract_fields(payload: Any) -> list[ExtractedField]:
    org, prefix = extract_organization(payload)
    rows: list[ExtractedField] = []

    legal = first_value(org, prefix, ["legal_name", "registered_name"])
    if legal:
        rows.append(available("Legal Name", legal[0], legal[1]))
    else:
        display_name = first_value(org, prefix, ["name"])
        reason = "Apollo did not return an explicit legal_name or registered_name field."
        if display_name:
            reason += f" API name was returned at {display_name[1]}."
            rows.append(unverified("Legal Name", display_name[1], reason))
        else:
            rows.append(missing("Legal Name", reason))

    field_paths = {
        "Website Domain": ["primary_domain", "domain", "website_domain"],
        "LinkedIn URL": ["linkedin_url", "linkedin"],
        "Employee Count": ["estimated_num_employees", "employee_count", "employees"],
        "Headcount Growth (1 Year)": [
            "organization_headcount_twelve_month_growth",
            "headcount_growth_12_month",
            "headcount_growth_1_year",
            "employee_growth_12_month",
        ],
        "Revenue / Revenue Band": [
            "annual_revenue_printed",
            "organization_revenue_printed",
            "annual_revenue",
            "organization_revenue",
            "estimated_annual_revenue",
            "revenue_range",
        ],
        "Industry": ["industry"],
        "Founded Year": ["founded_year", "founded"],
        "Total Funding Raised": ["total_funding", "funding_total", "total_funding_raised"],
    }
    formatters = {
        "Headcount Growth (1 Year)": format_growth,
        "Revenue / Revenue Band": format_money,
        "Total Funding Raised": format_money,
    }

    for field in [
        "Website Domain",
        "LinkedIn URL",
        "Employee Count",
        "Headcount Growth (1 Year)",
        "Revenue / Revenue Band",
        "Industry",
    ]:
        found = first_value(org, prefix, field_paths[field])
        if found:
            rows.append(available(field, found[0], found[1], formatters.get(field, format_scalar)))
        else:
            rows.append(missing(field))

    sub_parts: list[str] = []
    sub_sources: list[str] = []
    secondary = first_value(org, prefix, ["secondary_industries", "industries"])
    if secondary:
        sub_parts.append(format_scalar(secondary[0]))
        sub_sources.append(secondary[1])
    naics = first_value(org, prefix, ["naics_codes", "naics_code"])
    if naics:
        sub_parts.append(f"NAICS: {format_scalar(naics[0])}")
        sub_sources.append(naics[1])
    sic = first_value(org, prefix, ["sic_codes", "sic_code"])
    if sic:
        sub_parts.append(f"SIC: {format_scalar(sic[0])}")
        sub_sources.append(sic[1])
    if sub_parts:
        rows.append(ExtractedField("Sub Industry", " | ".join(sub_parts), ", ".join(sub_sources), "Available", ""))
    else:
        rows.append(missing("Sub Industry", "Sub-industry, NAICS, and SIC fields absent in Apollo response."))

    raw_address = first_value(org, prefix, ["raw_address"])
    if raw_address:
        rows.append(available("Headquarters Location", raw_address[0], raw_address[1]))
    else:
        location_parts: list[str] = []
        location_sources: list[str] = []
        for path in ["street_address", "city", "state", "country", "postal_code"]:
            part = first_value(org, prefix, [path])
            if part:
                location_parts.append(format_scalar(part[0]))
                location_sources.append(part[1])
        if location_parts:
            rows.append(ExtractedField("Headquarters Location", ", ".join(location_parts), ", ".join(location_sources), "Available", ""))
        else:
            rows.append(missing("Headquarters Location"))

    exact_sites = first_value(org, prefix, ["number_of_locations", "num_locations", "locations_count", "site_count"])
    if exact_sites:
        rows.append(available("Number of Sites/Locations", exact_sites[0], exact_sites[1]))
    else:
        retail_locations = first_value(org, prefix, ["retail_location_count"])
        if retail_locations:
            rows.append(
                unverified(
                    "Number of Sites/Locations",
                    retail_locations[1],
                    "Apollo returned retail_location_count, but not a verified total site/location count.",
                )
            )
        else:
            rows.append(missing("Number of Sites/Locations", "Exact site/location count absent in Apollo response."))

    founded = first_value(org, prefix, field_paths["Founded Year"])
    rows.append(available("Founded Year", founded[0], founded[1]) if founded else missing("Founded Year"))

    ownership = first_value(org, prefix, ["ownership_type", "company_type", "type"])
    if ownership:
        rows.append(available("Ownership Type", ownership[0], ownership[1]))
    else:
        ticker = first_value(org, prefix, ["publicly_traded_symbol", "ticker"])
        exchange = first_value(org, prefix, ["publicly_traded_exchange", "exchange"])
        if ticker or exchange:
            sources = ", ".join(item[1] for item in [ticker, exchange] if item)
            rows.append(ExtractedField("Ownership Type", "Public Company", sources, "Derived from API", "Derived from Apollo public listing fields."))
        else:
            rows.append(missing("Ownership Type", "Ownership/listing fields absent in Apollo response."))

    funding = first_value(org, prefix, field_paths["Total Funding Raised"])
    rows.append(available("Total Funding Raised", funding[0], funding[1], format_money) if funding else missing("Total Funding Raised"))

    return rows


def flatten_numbers(data: Any, prefix: str = "") -> dict[str, float]:
    values: dict[str, float] = {}
    if isinstance(data, dict):
        for key, value in data.items():
            next_prefix = f"{prefix}.{key}" if prefix else str(key)
            values.update(flatten_numbers(value, next_prefix))
    elif isinstance(data, list):
        for index, value in enumerate(data):
            next_prefix = f"{prefix}.{index}" if prefix else str(index)
            values.update(flatten_numbers(value, next_prefix))
    elif isinstance(data, (int, float)) and not isinstance(data, bool):
        values[prefix] = float(data)
    return values


def score_usage_path(path: str) -> int:
    lowered = path.lower()
    score = 0
    for token in ["organization", "organizations", "enrich", "enrichment", "credit", "credits", "consumed", "used", "usage"]:
        if token in lowered:
            score += 1
    return score


def clean_metric_path(path: str) -> str:
    return re.sub(r'\["([^"]+)",\s*"([^"]+)"\]', r"\1/\2", path)


def extract_credit_delta(before: dict[str, Any] | None, after: dict[str, Any] | None) -> tuple[str, str]:
    if not before or not after or not before.get("ok") or not after.get("ok"):
        return NOT_VERIFIED, "Usage endpoint unavailable or not authorized."
    before_json = before.get("json")
    after_json = after.get("json")
    if not isinstance(before_json, (dict, list)) or not isinstance(after_json, (dict, list)):
        return NOT_VERIFIED, "Usage endpoint did not return JSON usage metrics."

    before_numbers = flatten_numbers(before_json)
    after_numbers = flatten_numbers(after_json)
    candidates: list[tuple[int, str, float, float, float]] = []
    for path, after_value in after_numbers.items():
        if path not in before_numbers:
            continue
        delta = after_value - before_numbers[path]
        if delta == 0:
            continue
        score = score_usage_path(path)
        if score:
            candidates.append((score, path, before_numbers[path], after_value, delta))

    if not candidates:
        return NOT_VERIFIED, "No changed credit/usage metric found between before and after snapshots."

    candidates.sort(key=lambda item: (item[0], abs(item[4])), reverse=True)
    _, path, before_value, after_value, delta = candidates[0]
    return fmt_number(delta), f"{clean_metric_path(path)}: {fmt_number(before_value)} -> {fmt_number(after_value)}"


def fmt_number(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.4f}".rstrip("0").rstrip(".")


def summarize_rate_limit(api_result: dict[str, Any], usage_after: dict[str, Any] | None) -> str:
    rate_headers = [f"{key}: {value}" for key, value in api_result.get("headers", {}).items() if "rate" in key.lower()]
    if rate_headers:
        return "; ".join(rate_headers)

    usage_json = usage_after.get("json") if usage_after else None
    if isinstance(usage_json, (dict, list)):
        numbers = flatten_numbers(usage_json)
        rate_paths = []
        for path, value in numbers.items():
            lowered = path.lower()
            if "limit" in lowered or "left_over" in lowered or "remaining" in lowered:
                rate_paths.append(f"{path}={fmt_number(value)}")
            if len(rate_paths) >= 8:
                break
        if rate_paths:
            return "; ".join(rate_paths)
    return NOT_AVAILABLE


def records_retrieved(payload: Any) -> int:
    org, _ = extract_organization(payload)
    return 1 if org else 0


def response_body_for_save(result: dict[str, Any]) -> Any:
    if result.get("json") is not None:
        return result["json"]
    return {
        "status_code": result.get("status_code"),
        "text": result.get("text"),
        "error": result.get("error"),
    }


def firmographic_report_row(company: Company, extracted: list[ExtractedField], trace: dict[str, str]) -> dict[str, str]:
    by_field = {item.field: item for item in extracted}
    row: dict[str, str] = {
        "source_rank": company.source_rank,
        "company_name": company.company_name,
        "domain": company.domain,
        "input_linkedin_url": company.linkedin_url,
        "source_basis": company.source_basis,
        "api_name": API_NAME,
        "request_status": trace["success_failure"],
        "raw_response_path": trace["raw_response_path"],
    }
    for field in FIELD_NAMES:
        item = by_field[field]
        slug = FIELD_SLUGS[field]
        row[slug] = item.value
        row[f"{slug}_source_field"] = item.source_field
        row[f"{slug}_status"] = item.status
    return row


def missing_report_rows(company: Company, extracted: list[ExtractedField], raw_path: str) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for item in extracted:
        if item.status in {"Available", "Derived from API"}:
            continue
        rows.append(
            {
                "source_rank": company.source_rank,
                "company_name": company.company_name,
                "domain": company.domain,
                "field": item.field,
                "value": item.value,
                "status": item.status,
                "source_field": item.source_field,
                "reason": item.reason,
                "raw_response_path": raw_path,
            }
        )
    return rows


def write_csv(path: Path, rows: list[dict[str, str]], fieldnames: list[str]) -> None:
    path.parent.mkdir(exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def run_pipeline(input_path: Path, reuse_raw: bool = False, limit: int | None = None) -> int:
    load_dotenv()
    ensure_dirs()
    api_key = os.getenv("APOLLO_API_KEY") or os.getenv("APOLLO_MASTER_API_KEY")
    usage_key = os.getenv("APOLLO_MASTER_API_KEY") or api_key
    if not api_key:
        print("Missing Apollo API key. Set APOLLO_API_KEY or APOLLO_MASTER_API_KEY in .env.", file=sys.stderr)
        return 2

    companies = read_companies(input_path)
    if limit is not None:
        companies = companies[:limit]
    if not companies:
        print(f"No companies found in {input_path}.", file=sys.stderr)
        return 2

    usage_trace: list[dict[str, Any]] = []
    usage_available = False
    previous_usage: dict[str, Any] | None = None
    if usage_key:
        previous_usage = get_usage_snapshot(usage_key)
        usage_available = bool(previous_usage.get("ok"))
        usage_trace.append({"label": "before_run", **previous_usage})

    firmographic_rows: list[dict[str, str]] = []
    trace_rows: list[dict[str, str]] = []
    missing_rows: list[dict[str, str]] = []

    for company in companies:
        raw_path = RAW_DIR / f"{slugify(company.company_name)}.json"
        params = build_params(company)

        if reuse_raw and raw_path.exists():
            payload = load_json(raw_path)
            api_result = {
                "ok": True,
                "status_code": 200,
                "started_at": now_iso(),
                "response_received_at": now_iso(),
                "latency_ms": 0,
                "headers": {},
                "error": "",
                "json": payload,
            }
            credit_delta = NOT_VERIFIED
            credit_evidence = "Reused saved raw file; no live credit consumed in this run."
            usage_after = previous_usage
        else:
            api_result = api_request("GET", ORG_ENRICH_URL, api_key, params=params)
            payload = response_body_for_save(api_result)
            save_json(raw_path, payload)

            usage_after = None
            if usage_available and usage_key:
                usage_after = get_usage_snapshot(usage_key)
                usage_trace.append({"label": f"after_{slugify(company.company_name)}", "company_name": company.company_name, **usage_after})
            credit_delta, credit_evidence = extract_credit_delta(previous_usage, usage_after)
            if usage_available:
                previous_usage = usage_after

        raw_path_display = str(raw_path.relative_to(BASE_DIR))
        record_count = records_retrieved(payload)
        success = bool(api_result.get("ok") and record_count)
        error_message = api_result.get("error") or ""
        if api_result.get("ok") and not record_count:
            error_message = "HTTP success but no organization object returned."
        elif not api_result.get("ok") and not error_message:
            error_message = api_result.get("text") or ""
            if not error_message and isinstance(api_result.get("json"), dict):
                error_message = json.dumps(api_result["json"], ensure_ascii=False)

        trace = {
            "api_name": API_NAME,
            "source_rank": company.source_rank,
            "company_name": company.company_name,
            "domain": company.domain,
            "endpoint_used": ORG_ENRICH_URL,
            "request_params": json.dumps(params, ensure_ascii=False),
            "http_status": str(api_result.get("status_code") or NOT_AVAILABLE),
            "success_failure": "Success" if success else "Failure",
            "latency_ms": str(api_result.get("latency_ms") if api_result.get("latency_ms") is not None else NOT_AVAILABLE),
            "credits_consumed": credit_delta,
            "credit_evidence": credit_evidence,
            "rate_limits": summarize_rate_limit(api_result, usage_after),
            "rate_limit_hit": "Y" if api_result.get("status_code") == 429 else "N",
            "records_retrieved": str(record_count),
            "raw_response_path": raw_path_display,
            "started_at": str(api_result.get("started_at") or NOT_AVAILABLE),
            "response_received_at": str(api_result.get("response_received_at") or NOT_AVAILABLE),
            "error_message": error_message,
        }
        trace_rows.append(trace)

        extracted = extract_fields(payload)
        firmographic_rows.append(firmographic_report_row(company, extracted, trace))
        missing_rows.extend(missing_report_rows(company, extracted, raw_path_display))

        print(f"{trace['success_failure']}: {company.company_name} | HTTP {trace['http_status']} | {trace['latency_ms']} ms | raw={raw_path_display}")

    if usage_trace:
        save_json(USAGE_TRACE_PATH, usage_trace)

    firmographic_fieldnames = [
        "source_rank",
        "company_name",
        "domain",
        "input_linkedin_url",
        "source_basis",
        "api_name",
        "request_status",
        "raw_response_path",
    ]
    for field in FIELD_NAMES:
        slug = FIELD_SLUGS[field]
        firmographic_fieldnames.extend([slug, f"{slug}_source_field", f"{slug}_status"])

    trace_fieldnames = [
        "api_name",
        "source_rank",
        "company_name",
        "domain",
        "endpoint_used",
        "request_params",
        "http_status",
        "success_failure",
        "latency_ms",
        "credits_consumed",
        "credit_evidence",
        "rate_limits",
        "rate_limit_hit",
        "records_retrieved",
        "raw_response_path",
        "started_at",
        "response_received_at",
        "error_message",
    ]
    missing_fieldnames = [
        "source_rank",
        "company_name",
        "domain",
        "field",
        "value",
        "status",
        "source_field",
        "reason",
        "raw_response_path",
    ]

    write_csv(REPORTS_DIR / "firmographic_report.csv", firmographic_rows, firmographic_fieldnames)
    write_csv(REPORTS_DIR / "api_trace_report.csv", trace_rows, trace_fieldnames)
    write_csv(REPORTS_DIR / "missing_fields_report.csv", missing_rows, missing_fieldnames)

    print(f"Saved {REPORTS_DIR / 'firmographic_report.csv'}")
    print(f"Saved {REPORTS_DIR / 'api_trace_report.csv'}")
    print(f"Saved {REPORTS_DIR / 'missing_fields_report.csv'}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Run Apollo firmographic enrichment for input companies.")
    parser.add_argument("--input", help="CSV/TXT input path. Defaults to input/compnys.txt, then input/companies.csv.")
    parser.add_argument("--reuse-raw", action="store_true", help="Build reports from existing raw/apollo/*.json files without live API calls.")
    parser.add_argument("--limit", type=int, help="Process only the first N companies from the input file.")
    args = parser.parse_args()

    try:
        input_path = resolve_input_path(args.input)
        print(f"Input: {input_path}")
        return run_pipeline(input_path, reuse_raw=args.reuse_raw, limit=args.limit)
    except Exception as exc:
        print(f"Pipeline failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
