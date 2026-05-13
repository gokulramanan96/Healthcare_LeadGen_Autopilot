# Healthcare LeadGen Autopilot

Automated B2B lead-intelligence pipeline for healthcare and early-childhood facility news across GCC, Anglophone Africa, Francophone Africa, and the Caribbean. Replaces a manual market-research workflow that previously consumed several analyst-hours per day.

The pipeline runs unattended each morning, scrapes Google News across 21 countries, uses GPT-4o-mini to classify and score articles by lead quality, enriches qualified companies with contacts via Apollo.io, and writes a marketing-ready spreadsheet — with three-layer deduplication to stop the same event repeating across sources or days.

---

## What it produces

A multi-sheet Excel workbook on every run:

| Sheet | Purpose |
|---|---|
| **Healthcare News** | One row per unique news event with full lead context — company, top-3 contacts, signals, AI summary, research flag |
| **Apollo Results** | One row per unique company — 30-day enrichment cache to avoid repeat API spend |
| **Processed Hashes** | Every URL ever seen with outcome (Written / Low / Event Duplicate) — 30-day rolling |

---

## Architecture

```
       ┌────────────────────────────────────┐
       │  Query Builder                     │
       │  21 countries × (Universal +       │
       │  cluster-specific query templates) │
       └──────────────┬─────────────────────┘
                      │
       ┌──────────────▼──────────────┐
       │  Google News RSS            │
       └──────────────┬──────────────┘
                      │
       ┌──────────────▼──────────────┐
       │  Layer 1 — URL+title hash   │  ← blocks any article already seen
       └──────────────┬──────────────┘
              miss    │
       ┌──────────────▼──────────────┐
       │  GPT-4o-mini Pass           │
       │  → category, news_type,     │
       │    relevance, entity name   │
       └──────────────┬──────────────┘
              High    │
       ┌──────────────▼──────────────┐
       │  Layer 2 — event fingerprint│  ← blocks same event across sources
       │  (entity | country | type)  │      within 7-day window
       └──────────────┬──────────────┘
              miss    │
       ┌──────────────▼──────────────┐
       │  Layer 3 — Apollo cache     │  ← avoids paid re-enrichment
       │  (30-day TTL)               │
       └──────────────┬──────────────┘
              miss    │
       ┌──────────────▼──────────────┐
       │  Apollo: org search →       │
       │  people search → person     │
       │  enrichment → top-3 ranking │
       └──────────────┬──────────────┘
                      │
       ┌──────────────▼──────────────┐
       │  Excel writer (3 sheets)    │
       └─────────────────────────────┘
```

### Query architecture

Two layers feed the news scraper:

- **Universal layer** — runs for every country. Baseline hospital, clinic, and daycare construction and expansion queries in standard English.
- **Cluster-specific layer** — additive queries tuned to each market:
  - **GCC** — PPP, healthcare investment, health city
  - **Anglophone Africa** — World Bank / AfDB / IFC funding, district hospitals, diagnostic centres, ECD centres
  - **Francophone Africa** — French-language queries (hôpital, clinique, crèche) plus English fallback
  - **Caribbean** — health-centre groundbreakings, wellness centres, infant schools

### Dedup model

Three independent layers prevent duplicate work and duplicate output:

1. **URL hash** — `md5(url + title)` blocks any article already processed, written or discarded.
2. **Event fingerprint** — `normalised_entity | country | news_type` blocks the same event reported by different sources within a 7-day window.
3. **Apollo cache** — keyed on `normalised_entity | country`, valid for 30 days, avoids paid re-enrichment of recently looked-up companies.

### Contact ranking

Contacts are scored on two dimensions and the top three by combined score are written to the row:

- **Seniority** (1–5): CEO / Chairman / MD / President = 5, C-suite minus CEO and VP = 4, Director / Head / GM = 3, Manager-level = 2, other = 1.
- **Relevance bonus** (0–2): Development, projects, facilities, procurement, healthcare = +2; commercial, contracts, regional = +1.

---

## Setup

### Prerequisites

- Python 3.10+
- OpenAI API key with GPT-4o-mini access
- Apollo.io API key (any paid tier)

### Install

```bash
git clone https://github.com/gokulramanan96/Healthcare_LeadGen_Autopilot.git
cd Healthcare_LeadGen_Autopilot
pip install feedparser openpyxl requests openai
```

### Configure

Open `20260318_Hospitalleadsautomation.py` and set the keys and output paths at the top of the file:

```python
OPENAI_API_KEY   = "sk-..."
APOLLO_API_KEY   = "..."
EXCEL_FILE_PATH  = os.path.expanduser("~/Desktop/GCC_Healthcare_News_0.18.xlsx")
LOG_FILE_PATH    = os.path.expanduser("~/Desktop/gcc_news.log")
```

Default output path is the user's Desktop. Adjust `EXCEL_FILE_PATH` if running on Windows or a non-default working directory.

### Run

```bash
python 20260318_Hospitalleadsautomation.py
```

### Schedule

Daily automated runs via `cron` (macOS/Linux) or Task Scheduler (Windows). Sample cron entry for 7:00 AM local:

```cron
0 7 * * * /usr/bin/python3 /path/to/20260318_Hospitalleadsautomation.py
```

---

## Configuration

Tunables at the top of the script:

| Constant | Purpose | Default |
|---|---|---|
| `COUNTRIES` | Target country list | 21 countries |
| `MARKET_CLUSTERS` | Country → query-cluster mapping | — |
| `SIMILARITY_THRESHOLD` | Min name-match score to accept an Apollo company | 0.60 |
| `EVENT_DEDUP_DAYS` | Days an event fingerprint blocks re-entry | 7 |
| `APOLLO_CACHE_DAYS` | Days an Apollo enrichment stays valid | 30 |
| `HASH_RETENTION_DAYS` | Days a processed-hash entry is retained | 30 |
| `APOLLO_MIN_GAP` | Min seconds between Apollo calls | 0.15 |
| `FETCH_TIMEOUT` | Network timeout (seconds) | 5 |

---

## Output schema — Healthcare News sheet

```
Date Found | Headline | Source | URL | Country | Category | News Type |
AI Summary | Relevance Score | Relevance Reason | Key Signal |
Management Name | Company Website | Company LinkedIn | Company Size | Company Description |
[Contact 1: Name | Title | Score | Email | Phone | LinkedIn] |
[Contact 2: ...] |
[Contact 3: ...] |
Apollo Match Confidence | Research Flag | Event Fingerprint | Hash
```

Conditional formatting is applied automatically to Relevance, Category, News Type, and Research Flag columns.

---

## Roadmap

- Migrate API keys and paths to environment variables (`.env` + `python-dotenv`) for cross-platform portability
- Add `requirements.txt`, `.gitignore`, and a renamed entry point (`hospital_leads.py`)
- Split single-file script into a small module layout (`config/`, `enrichment/`, `io/`)
- Unit tests for `score_contact`, `name_similarity`, `is_vague_name`, and the three dedup layers
- Replace Excel output with Postgres plus a thin Streamlit dashboard
- Add LinkedIn Sales Navigator as a secondary enrichment source
- Email-digest delivery — top-N High-relevance leads plus the flagged manual-review queue
