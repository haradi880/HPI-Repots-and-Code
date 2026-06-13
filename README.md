# HPI Report Making

This project generates comprehensive firmographic and technographic reports for companies using multiple data sources and APIs.

## Project Structure

### Firmographic
Generates firmographic reports using Apollo API data.
- `firmographic_pipeline.py` - Main pipeline for processing firmographic data
- `build_full_10_company_report.py` - Generates complete 10-company reports
- `build_dbs_docx_report.py` - Builds DBS-specific DOCX reports
- `generate_apollo_dbs_report.py` - Generates Apollo DBS comparison reports

**Directories:**
- `input/` - Input company lists
- `raw/` - Raw API responses from Apollo
- `reports/` - Generated CSV and JSON reports
- `evidence/` - Supporting evidence files

### Technographic
Generates technographic reports using CoreSignal and PredictLeads APIs.
- `technographic_pipeline.py` - Main pipeline for processing technographic data
- `technographic_collector.py` - Collects technographic data
- `generate_docx_report.py` - Generates DOCX reports
- `PREDICTLEADS_fetch.py` - Fetches PredictLeads data

**Directories:**
- `input/` - Input company lists
- `output/` - Generated reports and monitoring data
- `raw/` - Raw API responses from CoreSignal
- `PredictLeads/` - PredictLeads-specific data and reports

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Configure API keys in `.env`:
   ```
   APOLLO_API_KEY=your_key_here
   CORESIGNAL_API_KEY=your_key_here
   PREDICTLEADS_API_KEY=your_key_here
   ```

3. Add company list to input files (e.g., `Firmographic/input/compnys.txt`)

## Usage

### Generate Firmographic Reports
```bash
python Firmographic/firmographic_pipeline.py
```

### Generate Technographic Reports
```bash
python Technographic/technographic_pipeline.py
```

### Generate Full 10-Company Report
```bash
python Firmographic/build_full_10_company_report.py
```

## Output

Reports are generated in:
- Firmographic: `Firmographic/reports/`
- Technographic: `Technographic/output/`

Output formats include CSV, JSON, and DOCX files.
