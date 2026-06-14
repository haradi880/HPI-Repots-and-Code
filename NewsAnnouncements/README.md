# HPI News / Key Announcements API Evaluation

This folder evaluates news and announcement providers for the HPI account-intelligence pilot.

## Providers

- Exa Search API
- Tavily Search API
- Google News RSS
- NewsAPI Everything API
- PredictLeads: blocked until a valid key is available
- GDELT: not used, no API key required but excluded from this run per scope

## Target Fields

- Event headline, URL, and date
- Event type: funding, expansion, launch, leadership, M&A, partnership, other
- Source / publisher
- Relevance / confidence
- Coverage depth: events per account in the last 12 months

## Outputs

Raw responses:

- `raw/exa/data/{company_slug}.json`
- `raw/tavily/data/{company_slug}.json`
- `raw/google_news_rss/data/{company_slug}.json`
- `raw/newsapi/data/{company_slug}.json`
- `raw/predictleads/data/{company_slug}.json`
- `raw/gdelt/data/{company_slug}.json`

Reports:

- `reports/company_news_announcements.csv`
- `reports/event_detail.csv`
- `reports/api_call_log.csv`
- `reports/api_comparison_report.csv`
- `reports/api_tracing_report.csv`
- `reports/missing_fields_report.csv`
- `reports/hpi_news_announcements_api_evaluation_*.xlsx`

## Run

Live capped run:

```powershell
python news_announcements_pipeline.py --limit 10 --per-provider-limit 10
```

Rebuild from saved raw responses:

```powershell
python news_announcements_pipeline.py --limit 10 --reuse-raw
```
