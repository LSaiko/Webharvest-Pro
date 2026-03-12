# 🏗 Architecture — WebHarvest Pro

## System Overview

WebHarvest Pro is a **three-layer data pipeline**:

```
┌─────────────────────────────────────────────────┐
│                  INPUT LAYER                    │
│  Web UI (browser)  ←→  CLI (build_excel.py)     │
└─────────────────────┬───────────────────────────┘
                      │ List of URLs
                      ▼
┌─────────────────────────────────────────────────┐
│              COLLECTION LAYER                   │
│                                                 │
│   is_url_safe()  ──→  check_robots_txt()        │
│        │                     │                  │
│     BLOCKED              crawl_delay            │
│                               │                 │
│                         fetch_page()            │
│                               │                 │
│                    paywall detection             │
│                               │                 │
│                 extract_structured_data()        │
│                               │                 │
│                         result dict             │
└─────────────────────┬───────────────────────────┘
                      │ List of result dicts
                      ▼
┌─────────────────────────────────────────────────┐
│               EXPORT LAYER                      │
│                                                 │
│   create_excel_export()   export_to_csv()       │
│   ├─ Summary sheet        ├─ summary.csv        │
│   ├─ Raw Data sheet       ├─ raw_content.csv    │
│   ├─ Tables sheet         └─ links.csv          │
│   ├─ Links sheet                                │
│   ├─ Analytics Template                        │
│   ├─ PowerBI Guide                              │
│   ├─ PivotTable Guide                           │
│   └─ Changelog                                 │
└─────────────────────────────────────────────────┘
```

## Module Responsibilities

### `scraper_engine.py`
**Single responsibility:** Given a URL, return a structured result dict or a failure result.

Key functions:
- `is_url_safe(url)` — URL pattern + domain blocklist validation
- `check_robots_txt(url)` — Fetches and parses robots.txt, returns (allowed, crawl_delay)
- `fetch_page(url)` — HTTP GET with error handling, returns (soup, status, meta)
- `extract_structured_data(soup, url)` — DOM traversal, returns typed data dict
- `scrape_url(url)` — **Orchestrates the above** — the public API of this module

This module has **no knowledge** of how results are exported.

### `excel_exporter.py`
**Single responsibility:** Given a list of result dicts, write a formatted `.xlsx` file.

Key functions:
- `create_excel_export(results, output_path)` — Entry point, creates workbook
- `_create_summary_sheet(wb, results)` — KPI cards + overview table
- `_create_raw_data_sheet(wb, results)` — Flat content table for analysis
- `_create_tables_sheet(wb, results)` — Extracted HTML tables
- `_create_links_sheet(wb, results)` — Link inventory
- `_create_analytics_template_sheet(wb, results)` — Pre-built formulas
- `_create_powerbi_guide_sheet(wb)` — Static guide (no data dependency)
- `_create_pivot_guide_sheet(wb)` — Static guide
- `_create_changelog_sheet(wb, results)` — Run history

Helper factory functions (used everywhere):
- `make_fill(hex)`, `make_font(...)`, `make_border(...)`, `make_center()`, `make_left()`
- `style_header_row(ws, row, cols, ...)` — Applies header styling to a row
- `auto_fit_columns(ws, ...)` — Calculates column widths from content

This module has **no knowledge** of how data was collected.

### `csv_exporter.py`
**Single responsibility:** Given a list of result dicts, write 3 timestamped CSV files.

Outputs: `summary_[ts].csv`, `raw_content_[ts].csv`, `links_[ts].csv`

### `build_excel.py`
**Single responsibility:** Parse user input (CLI args or interactive), call the scraper, call the exporters.

Contains no scraping logic and no export formatting logic — it purely orchestrates.

---

## Data Flow

### The Result Dict (the contract between modules)

```python
{
    # Always present
    'url': str,
    'timestamp': str,       # ISO 8601
    'success': bool,
    'error': str | None,
    'meta': dict,
    'crawl_delay': float,

    # Only present when success=True
    'data': {
        'title': str,
        'description': str,
        'headings': [{'level': str, 'text': str}],
        'paragraphs': [str],
        'tables': [{'headers': [str], 'rows': [[str]]}],
        'links': [{'text': str, 'url': str, 'external': bool}],
        'lists': [[str]],
        'metadata': dict,
        'structured_data': [dict],   # JSON-LD
    }
}
```

All downstream code (exporters) uses `data = r.get('data') or {}` to safely handle failed results.

---

## Extension Points

**Add a new data source (e.g., Playwright for JS-rendered pages):**
Only change `scraper_engine.py`. The result schema stays the same.

**Add a new export format (e.g., Google Sheets API):**
Create a new exporter module. It consumes the same result list.

**Add a new analytics sheet:**
Add a `_create_X_sheet(wb, results)` function in `excel_exporter.py` and call it from `create_excel_export`.

**Add a new ethics check:**
Add a check to `is_url_safe()` or `fetch_page()`. The rest of the pipeline is unchanged.
