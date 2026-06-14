# HPI Jobs / Hiring API Evaluation

This folder evaluates Jobs/Hiring providers for the HPI account-intelligence pilot.

## Providers

- PredictLeads Job Openings API: blocked until a valid API key is available
- TheirStack Job Search API: live API, capped sample size to control credits
- Coresignal: reused saved Multi-source Company raw exports with active job signals
- Apollo Organization Job Postings API: live API using Apollo organization IDs from firmographic enrichment
- LinkUp / Aura: sample/demo on request, not executed in the free-tier run

## Target Fields

- Active job count, total and Singapore
- Roles by function: engineering, design, IT, sales, operations
- Hiring velocity, postings last 90 days or closest available provider signal
- Job locations
- Posting first-seen or posted date
- Source URL

## Outputs

Raw responses:

- `raw/apollo/data/{company_slug}.json`
- `raw/theirstack/data/{company_slug}.json`
- `raw/coresignal/data/{company_slug}.json`
- `raw/predictleads/data/{company_slug}.json`

Reports:

- `reports/company_jobs_hiring.csv`
- `reports/job_detail.csv`
- `reports/api_call_log.csv`
- `reports/api_comparison_report.csv`
- `reports/api_tracing_report.csv`
- `reports/missing_fields_report.csv`
- `reports/hpi_jobs_hiring_api_evaluation_*.xlsx`

## Run

Live/reuse mixed run:

```powershell
python jobs_hiring_pipeline.py --limit 10
```

Rebuild from saved raw responses:

```powershell
python jobs_hiring_pipeline.py --limit 10 --reuse-raw
```

TheirStack job search consumes credits per returned job, so the pipeline defaults to a capped sample of 10 jobs per company.
