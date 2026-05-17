# ⬡ WebHarvest Pro

> **Ethical public-data web scraper** with scheduled monitoring, multi-site collection, and pre-built Excel/PowerBI analytics export — built entirely in Python and vanilla JavaScript.

[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-22c55e?style=flat-square)](LICENSE)
[![Ethics](https://img.shields.io/badge/Ethics-robots.txt%20compliant-6C63FF?style=flat-square)](#ethics--compliance)
[![Export](https://img.shields.io/badge/Export-Excel%20%7C%20CSV%20%7C%20PowerBI-FF6B9D?style=flat-square)](#output-formats)

---

## 📸 What It Does

WebHarvest Pro is a **full-stack data collection pipeline** that:

1. **Validates** every URL against an ethics engine (robots.txt, paywall detection, auth-wall detection)
2. **Scrapes** structured public data — headings, paragraphs, tables, links, metadata
3. **Tracks** content via SHA-256 hashes stored in every export for change auditing (browser UI adds scheduling on top)
4. **Exports** to richly formatted Excel workbooks with pre-built PowerBI/PivotTable analytics templates, or clean analysis-ready CSV files

---

## 🗂 Repository Structure

```
webharvest-pro/
│
├── src/
│   ├── scraper_engine.py      # Core scraping pipeline + ethics enforcement
│   ├── excel_exporter.py      # 8-sheet Excel builder with analytics templates
│   ├── csv_exporter.py        # Analysis-ready CSV export (3 files)
│   └── build_excel.py         # CLI entry point — orchestrates full pipeline
│
├── web/
│   └── index.html             # Browser UI — no install needed
│
├── docs/
│   ├── ARCHITECTURE.md        # System design & module breakdown
│   ├── SKILLS_SHOWCASE.md     # Annotated code — interview talking points
│   └── ETHICS.md              # Compliance policy & scraping ethics
│
├── examples/
│   ├── sample_export.xlsx     # Pre-generated Excel output to explore
│   └── sample_output.json     # Sample raw scrape data
│
├── requirements.txt
├── LICENSE
└── README.md                  ← You are here
```

---

## ⚡ Quick Start

### Web UI (zero install)
```
1. Download index.html
2. Open in any browser
3. Paste URLs → click Scrape → export CSV
```

### Python CLI
```bash
# Install dependencies
pip install -r requirements.txt

# Scrape one or more public URLs
python src/build_excel.py --urls https://news.ycombinator.com https://quotes.toscrape.com

# Build Excel from a previously saved JSON file
python src/build_excel.py --json examples/sample_output.json

# CSV output only, custom output directory
python src/build_excel.py --urls https://example.com --format csv --output ./exports
```

---

## 📊 Output Formats

### Excel Workbook (`.xlsx`) — 8 sheets

| Sheet | Contents |
|---|---|
| 📊 Summary | KPI stat cards, per-site overview table, auto-status badges |
| 📄 Raw Data | All headings, paragraphs, and list items with source attribution |
| 📋 Tables | Every HTML table found, reformatted with column headers |
| 🔗 Links | Full link inventory with internal/external classification and domain |
| 📈 Analytics Template | Pre-written Excel formulas + 6 PivotTable configuration recipes |
| ⚡ PowerBI Guide | DAX formulas + step-by-step dashboard layout guide |
| 🔄 PivotTable Guide | Slicer setup, FILTER/UNIQUE/SORT formulas, chart type recommendations |
| 📝 Changelog | Timestamped run history with content hashes for change tracking |

### CSV Files (3 files)

| File | Contents |
|---|---|
| `summary_[ts].csv` | One row per URL with counts and status |
| `raw_content_[ts].csv` | All content rows with type labels (heading, paragraph, list_item) |
| `links_[ts].csv` | All links with source, text, URL, and external flag |

---

## 🛡 Ethics & Compliance

This tool is built around **four layers of protection** against unethical scraping:

| Layer | Mechanism | What it blocks |
|---|---|---|
| **URL Validation** | Regex pattern matching | Login pages, dashboards, private/admin paths |
| **Domain Blocklist** | Exact domain matching | Gmail, Instagram, LinkedIn DMs, Patreon, etc. |
| **robots.txt** | `urllib.robotparser` | Any path disallowed by the site's crawl policy |
| **Paywall Detection** | Keyword scan of page HTML | Subscribe walls, members-only gates |

**Designed for:** public news sites, government data portals, open academic resources, company public pages, open directories.

**Not for:** paywalled articles, authenticated portals, personal social media pages, or any content where you don't have explicit permission.

---

## 🔧 Technical Stack

| Component | Technology | Why |
|---|---|---|
| HTTP client | `requests` | Robust redirects, timeout handling, session management |
| HTML parsing | `BeautifulSoup4` + `lxml` | Fast, fault-tolerant DOM traversal |
| robots.txt | `urllib.robotparser` (stdlib) | Zero extra dependency, RFC-compliant |
| Excel generation | `openpyxl` | Full formatting API — colors, borders, tables, charts |
| CSV export | `csv` (stdlib) | Zero-dependency CSV writing with clean column schema |
| UI | Vanilla JS + HTML/CSS | Zero-dependency browser app, no build step |
| CLI | `argparse` (stdlib) | Self-documenting command interface |

---

## 💼 Skills Demonstrated

A quick reference for hiring managers — each item links to the relevant code:

| Skill Area | What's built | File |
|---|---|---|
| **Web scraping** | Full pipeline: fetch → parse → structure → export | `scraper_engine.py` |
| **Data engineering** | Structured extraction of tables, links, headings, metadata | `scraper_engine.py` |
| **API / HTTP** | Rate limiting, redirects, User-Agent headers, timeout handling | `scraper_engine.py` |
| **File I/O** | Generate formatted multi-sheet Excel and multi-file CSV | `excel_exporter.py`, `csv_exporter.py` |
| **Security thinking** | Ethics engine with 4 validation layers before any request fires | `scraper_engine.py` |
| **Automation** | Scheduled re-scraping + content-hash change detection | `index.html` (JS scheduler) |
| **Frontend** | Responsive single-file browser app — no framework | `index.html` |
| **CLI design** | Argparse with modes: direct URLs, JSON input, interactive | `build_excel.py` |
| **Data viz prep** | Pre-written DAX, PivotTable recipes, chart recommendations | `excel_exporter.py` |

See [`docs/SKILLS_SHOWCASE.md`](docs/SKILLS_SHOWCASE.md) for annotated code walkthroughs of each area.

---

## 📦 Installation

```bash
git clone https://github.com/LSaiko/Webharvest-Pro.git
cd Webharvest-Pro
pip install -r requirements.txt
python src/build_excel.py --urls https://quotes.toscrape.com
```

**Requirements:**
```
requests>=2.28.0,<3.0.0
beautifulsoup4>=4.11.0,<5.0.0
lxml>=4.9.0,<6.0.0
openpyxl>=3.1.0,<4.0.0
```

---

## 📄 License

MIT — free for personal and commercial use. See [LICENSE](LICENSE).

---

## 🤝 Contributing

Pull requests welcome. Please ensure any new scraping features:
1. Pass the ethics validation pipeline
2. Include a docstring explaining what data is collected
3. Respect the existing rate-limiting patterns

---

*Built as a portfolio project to demonstrate end-to-end data engineering: collection, transformation, and analytics-ready export.*
