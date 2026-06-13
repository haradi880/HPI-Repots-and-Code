# HPI Data Provider Evaluation Reports

This repository contains the cleaned 10-company HPI firmographic and technographic API evaluation package.

## Current Scope

- Input list: `Firmographic/input/compnys.txt` and `Technographic/input/compnys.txt`
- Companies: DBS Group, Singapore Telecommunications, United Overseas Bank, ST Engineering, Toyota Motor Corporation, Sony Group Corporation, Samsung Electronics, SK Hynix, Infosys, and BHP
- Firmographic providers: Apollo and Coresignal
- Technographic providers: TheirStack and Coresignal
- Jobs/Hiring providers: TheirStack, Coresignal, Apollo, PredictLeads status, and LinkUp/Aura status

## Folder Structure

```text
Firmographic/
  input/                         Input company list
  raw/apollo/                    Apollo raw API responses
  raw/coresignal/data/           Coresignal raw API responses
  reports/full_10_company_report/ Final firmographic deliverables and audit CSVs

Technographic/
  input/                         Input company list
  raw/theirstack/data/           TheirStack raw API responses
  raw/coresignal/data/           Coresignal raw API responses
  apilogs/                       Per-company API call logs
  reports/                       Final technographic deliverables and audit CSVs

JobsHiring/
  input/                         Input company list
  raw/                           Raw job/hiring provider responses
  apilogs/                       Per-company API call logs
  reports/                       Final jobs/hiring deliverables and audit CSVs
```

## Final Deliverables

Firmographic:

- `Firmographic/reports/full_10_company_report/hpi_10_company_firmographic_api_comparison_apollo_coresignal.xlsx`
- `Firmographic/reports/full_10_company_report/hpi_10_company_firmographic_api_comparison_apollo_coresignal.docx`
- `Firmographic/reports/full_10_company_report/hpi_10_company_coresignal_firmographic_api_report.xlsx`
- `Firmographic/reports/full_10_company_report/hpi_10_company_coresignal_firmographic_api_report.docx`
- `Firmographic/reports/full_10_company_report/hpi_10_company_firmographic_api_report_revised.xlsx`
- `Firmographic/reports/full_10_company_report/hpi_10_company_firmographic_api_report_revised.docx`
- `Firmographic/reports/full_10_company_report/company_field_level_report.csv`
- `Firmographic/reports/full_10_company_report/api_trace_full_report.csv`
- `Firmographic/reports/full_10_company_report/coresignal_api_trace_report.csv`
- `Firmographic/reports/full_10_company_report/coresignal_firmographic_report.csv`

Technographic:

- `Technographic/reports/hpi_technographic_api_comparison_20260613_125914.xlsx`
- `Technographic/reports/hpi_technographic_api_comparison_20260613_125914.docx`
- `Technographic/reports/company_technographics.csv`
- `Technographic/reports/technology_detail.csv`
- `Technographic/reports/api_tracing_report.csv`
- `Technographic/reports/api_call_log.csv`

Jobs / Hiring:

- `JobsHiring/reports/hpi_jobs_hiring_api_evaluation_20260613_152928.xlsx`
- `JobsHiring/reports/hpi_jobs_hiring_api_evaluation_20260613_152928.docx`
- `JobsHiring/reports/company_jobs_hiring.csv`
- `JobsHiring/reports/job_detail.csv`
- `JobsHiring/reports/api_tracing_report.csv`
- `JobsHiring/reports/api_call_log.csv`

## Rebuild Reports

Rebuild firmographic report from saved raw exports:

```powershell
python Firmographic/build_full_10_company_report.py
```

Rebuild technographic report from saved raw exports without spending API credits:

```powershell
python Technographic/technographic_pipeline.py --limit 10 --reuse-raw
```

Rebuild Jobs/Hiring report from saved raw exports:

```powershell
python JobsHiring/jobs_hiring_pipeline.py --limit 10 --reuse-raw
```

Run live technographic refresh only for empty/missing TheirStack raw responses:

```powershell
python Technographic/technographic_pipeline.py --limit 10 --apis theirstack,coresignal --refresh-empty-raw
```

## Environment

Secrets are intentionally excluded from Git. Use local `.env` files only.

Required keys for live API runs:

```text
APOLLO_API_KEY=
CORESIGNAL_API_KEY=
THEIRSTACK_API_KEY=
PREDICTLEADS_API_KEY=
```

Saved raw responses are committed so the final reports can be regenerated without live API calls.
