# HPI Technographic API Evaluation

This folder now has one clean entry point for technographic extraction:

```powershell
python technographic_pipeline.py --limit 10
```

The pipeline uses the two working APIs configured in `.env`:

- `THEIRSTACK_API_KEY`
- `CORESIGNAL_API_KEY`

## What It Extracts

- Detected technologies, full list
- Hardware, device, and endpoint signals
- Print, MPS, collaboration, and UC stack
- IT spend estimate when the provider returns it
- First/last seen or freshness date per technology

## Audit Outputs

Raw API responses are saved here:

- `raw/theirstack/data/{company_slug}.json`
- `raw/coresignal/data/{company_slug}.json`

Per-call API logs are saved here:

- `apilogs/theirstack/{company_slug}.jsonl`
- `apilogs/coresignal/{company_slug}.jsonl`

Reports are saved here:

- `reports/api_comparison_report.csv`
- `reports/api_tracing_report.csv`
- `reports/api_call_log.csv`
- `reports/company_technographics.csv`
- `reports/company_technographics.json`
- `reports/technology_detail.csv`
- `reports/missing_fields_report.csv`
- `reports/hpi_technographic_api_comparison_*.xlsx`
- `reports/hpi_technographic_api_comparison_*.docx`

## Credit And Rate-Limit Handling

TheirStack:

- Endpoint: `POST /v1/companies/technologies`
- Credit balance endpoint: `GET /v0/billing/credit-balance`
- If balance fields are returned, the pipeline records before/after balances.
- If no balance delta is available, successful company lookups are marked with the documented estimate of 3 credits.

Coresignal:

- Endpoint: `GET /company_multi_source/enrich?website={URL}`
- Coresignal documents 2 collection credits per successful Multi-source Company enrich call.
- The pipeline records that documented estimate when no account-level balance endpoint is available.

## Rebuild From Saved Raw Responses

Use this when you want to regenerate reports without spending another API call:

```powershell
python technographic_pipeline.py --limit 10 --reuse-raw
```

The input company file is `input/compnys.txt`.
