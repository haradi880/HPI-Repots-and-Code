# HPI Contact-Level API Evaluation

This folder evaluates contact-level enrichment providers for the HPI account-intelligence pilot.

## Providers

- Apollo People Search API: used to seed named contacts for each company
- FullEnrich Contact Bulk Enrichment API: capped waterfall enrichment for selected seed contacts
- Prospeo Person Enrichment API: capped email and mobile enrichment for selected seed contacts
- SignalHire Candidate Search API: capped at 5 lookups because available credits are very low
- People Data Labs: marked server_down/not executed for this run

## Target Fields

- Verified work email and confidence
- Direct/mobile phone and confidence
- Title, seniority, and department
- Reports-to / org-chart hierarchy
- LinkedIn URL
- Match rate per requested contact

## Outputs

Raw responses:

- `raw/apollo/data/{company_slug}.json`
- `raw/fullenrich/data/{company_slug}.json`
- `raw/fullenrich_search/data/{company_slug}.json`
- `raw/prospeo/data/{company_slug}__{contact_slug}.json`
- `raw/signalhire/data/{company_slug}__{contact_slug}.json`
- `raw/people_data_labs/data/{company_slug}.json`

Reports:

- `reports/company_contact_level.csv`
- `reports/contact_detail.csv`
- `reports/api_call_log.csv`
- `reports/api_comparison_report.csv`
- `reports/api_tracing_report.csv`
- `reports/missing_fields_report.csv`
- `reports/hpi_contact_level_api_evaluation_*.xlsx`
- `reports/hpi_contact_level_api_evaluation_*.docx`

## Run

Live capped run:

```powershell
python contact_level_pipeline.py --limit 10 --enrichment-limit 5 --signalhire-limit 5
```

Rebuild from saved raw responses:

```powershell
python contact_level_pipeline.py --limit 10 --reuse-raw
```

The live run is intentionally credit-controlled. Apollo seeds up to 2 people per company, FullEnrich and Prospeo enrich at most 5 seed contacts by default, and SignalHire is capped at 5 lookups.
