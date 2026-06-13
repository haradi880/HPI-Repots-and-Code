# HPI Firmographic API Evaluation

This folder contains the firmographic evaluation for the 10-company HPI pilot list.

## Providers

- Apollo Organization Enrichment API
- Coresignal Multi-source Company Enrichment API

## Raw Evidence

- Apollo raw responses: `raw/apollo/{company_slug}.json`
- Coresignal raw responses: `raw/coresignal/data/{company_slug}.json`

## Final Report Folder

All current deliverables are in:

```text
reports/full_10_company_report/
```

Key outputs:

- `hpi_10_company_firmographic_api_comparison_apollo_coresignal.xlsx`
- `hpi_10_company_firmographic_api_comparison_apollo_coresignal.docx`
- `hpi_10_company_coresignal_firmographic_api_report.xlsx`
- `hpi_10_company_coresignal_firmographic_api_report.docx`
- `hpi_10_company_firmographic_api_report_revised.xlsx`
- `hpi_10_company_firmographic_api_report_revised.docx`
- `company_field_level_report.csv`
- `api_comparison_report.csv`
- `api_trace_full_report.csv`
- `coresignal_api_trace_report.csv`
- `coresignal_firmographic_report.csv`
- `missing_reason_api_gated_report.csv`

## Rebuild

```powershell
python build_full_10_company_report.py
```

The rebuild uses saved raw Apollo and Coresignal exports and does not require new API calls.
