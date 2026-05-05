# The Quarry — CRE Comps & Contacts Pipeline

## Purpose
The Quarry consumes the structured JSON handoff produced by App 1's daily digest and writes it into two persistent Excel workbooks: a **Comps database** (Sales, Leases, Loans tabs) and a **Contacts database** (upsert with change-flagging and deal notes). It runs as a separate GitHub Actions workflow triggered after App 1 completes.

## Relationship to App 1
- The Quarry is **read-only with respect to App 1** — it never modifies App 1 files.
- Input: `articles_handoff.json` written by App 1's `digest.py` and deployed to GitHub Pages at `https://altmanr91.github.io/CRE-News-Reader/articles_handoff.json`
- The Quarry fetches this file at runtime; no shared filesystem required.
- Future changes to App 1's pipeline are automatically captured as long as the JSON schema is maintained. If App 1 adds new fields to `ArticleSummary`, The Quarry picks them up on next run without code changes.

## Directory Structure
```
quarry/
├── CLAUDE.md              # This file
├── quarry_pipeline.py     # Main entry point: fetch handoff → write comps → write contacts
├── comps_writer.py        # Append to Excel comps workbook (Sales/Leases/Loans tabs)
├── contacts_writer.py     # Upsert Excel contacts workbook with change-flagging
└── requirements.txt       # openpyxl (+ requests, python-dotenv)
```

GitHub Actions workflow: `.github/workflows/quarry_pipeline.yml`

## Input: articles_handoff.json Schema
App 1 writes this file after each digest run. Schema per article record:

```json
{
  "date": "2026-04-30",
  "articles": [
    {
      "title": "...",
      "source": "Commercial Observer",
      "link": "https://...",
      "published": "Thu, 30 Apr 2026 ...",
      "page_url": "https://altmanr91.github.io/CRE-News-Reader/articles/2026-04-30-001.html",
      "transaction_type": "Sale",
      "article_type": null,
      "market": "Dallas, TX",
      "narrative": "2-4 sentence summary...",
      "data_points": {
        "property_type": "Multifamily",
        "property_name": "The Meadows",
        "address": "123 Main St, Dallas, TX 75201",
        "address_sourced_separately": false,
        "size_sf": 120000,
        "size_units": 250,
        "size_beds": null,
        "size_keys": null,
        "sale_price": 50000000,
        "loan_amount": null,
        "total_project_cost": null,
        "occupancy": 94.5,
        "year_built": 2018,
        "completion": null,
        "land_area_acres": null,
        "notable_features": null,
        "project_notes": null
      },
      "companies_people": [
        {
          "label": "BUYER",
          "firm_name": "Blackstone",
          "people": [{"name": "John Smith", "title": "Managing Director"}]
        }
      ],
      "tenants": ["Amazon", "Starbucks"],
      "financing": "65% LTV senior loan from Wells Fargo.",
      "market_intelligence": null
    }
  ]
}
```

`transaction_type` values from App 1: Sale, Acquisition, Lease, Loan, Refinance, Development, Construction, Promotion
`article_type` is set for non-transaction articles (market research, policy, etc.) — The Quarry ignores these for comps.

## Output: Comps Workbook (comps.xlsx)

### Tab routing
| transaction_type | Tab |
|---|---|
| Sale, Acquisition | Sales |
| Lease | Leases |
| Loan, Refinance | Loans |
| Development, Construction | (skip — no comp value until traded) |
| Promotion | (skip) |
| Market research articles | (skip) |

### Sales tab columns
| Column | Source |
|---|---|
| Date | `date` |
| Property Name | `data_points.property_name` |
| Address | `data_points.address` |
| Market | `market` |
| Property Type | `data_points.property_type` |
| Size (SF) | `data_points.size_sf` |
| Units | `data_points.size_units` |
| Sale Price | `data_points.sale_price` |
| $/SF | calculated (sale_price / size_sf) |
| $/Unit | calculated (sale_price / size_units) |
| Year Built | `data_points.year_built` |
| Occupancy % | `data_points.occupancy` |
| Buyer | firms with label BUYER or ACQUIRER |
| Seller | firms with label SELLER |
| Broker | firms with label BROKER or SALES BROKER |
| Lender | firms with label LENDER |
| Source | `source` |
| Link | `link` |
| Notes | `financing` (capital stack detail) |

### Leases tab columns
| Column | Source |
|---|---|
| Date | `date` |
| Property Name | `data_points.property_name` |
| Address | `data_points.address` |
| Market | `market` |
| Property Type | `data_points.property_type` |
| Size (SF) | `data_points.size_sf` |
| Tenants | `tenants` joined by ", " |
| Landlord | firms with label LANDLORD or OWNER |
| Tenant Rep Broker | firms with label TENANT REP or BROKER |
| Landlord Broker | firms with label LANDLORD REP or BROKER |
| Source | `source` |
| Link | `link` |

### Loans tab columns
| Column | Source |
|---|---|
| Date | `date` |
| Property Name | `data_points.property_name` |
| Address | `data_points.address` |
| Market | `market` |
| Property Type | `data_points.property_type` |
| Size (SF) | `data_points.size_sf` |
| Units | `data_points.size_units` |
| Loan Amount | `data_points.loan_amount` |
| Loan/SF | calculated |
| Loan/Unit | calculated |
| Borrower/Sponsor | firms with label BORROWER, SPONSOR, or DEVELOPER |
| Lender | firms with label LENDER |
| Source | `source` |
| Link | `link` |
| Notes | `financing` |

### Append behavior
- Each daily run appends new rows — never overwrites existing data.
- **Duplicate address detection:** If the same `address` already exists in the tab, apply yellow fill (`FFFF00`) to the new row and to the existing row(s) with that address. This flags for manual review — the same property may have traded or refinanced again, or it may be a duplicate article.
- Calculated columns ($/SF, $/Unit, Loan/SF, etc.) are written as plain values, not Excel formulas, to avoid formula maintenance complexity.

## Output: Contacts Workbook (contacts.xlsx)

### Schema
Single "Contacts" tab with these columns:
| Column | Notes |
|---|---|
| Name | Person's full name — primary key component |
| Title | Most recent title seen |
| Company | Most recent firm — primary key component |
| Role | Label from companies_people (BUYER, LENDER, BROKER, etc.) — most recent seen |
| Market | Most recent market |
| Last Seen | Date of most recent article mentioning them |
| First Seen | Date of first article mentioning them |
| Appearances | Count of articles they've appeared in |
| Notes | Pipe-separated list of deal appearances, e.g. `2026-04-30: $50M sale at The Meadows, Dallas TX | 2026-05-01: $120M loan at ...` |
| Review Flag | "YES" if any field changed since last time; blank otherwise |

### Upsert logic
Match key: **Name + Company** (case-insensitive, stripped).

**New contact:** append new row with all fields populated, Review Flag blank.

**Existing contact (same Name + Company):** update in place:
- If `Title` changed: update cell, set Review Flag = "YES"
- If `Role` changed: update cell, set Review Flag = "YES"
- If `Market` changed: update cell, set Review Flag = "YES"
- Always: update `Last Seen`, increment `Appearances`, append to `Notes`
- Review Flag is never cleared automatically — must be manually reset.

**Name match, different company:** treat as a new row (person may have changed firms). Set Review Flag = "YES" on the new row to draw attention to a possible firm change.

### Notes field format
`YYYY-MM-DD: [role] — [article title snippet] ([market])` — one entry per appearance, pipe-separated:
```
2026-04-30: BUYER — Blackstone acquires The Meadows, 250-unit multifamily... (Dallas, TX) | 2026-05-02: LENDER — ...
```

### Contacts sourcing
- Only include people with a non-empty `name` field.
- Pull from `companies_people[*].people[*]` for all shown (non-filtered) articles.
- Promotions (transaction_type = "Promotion"): also include — they give the most reliable title/company data.
- Skip market research / non-transaction articles with no companies_people data.

## GitHub Actions Workflow

File: `.github/workflows/quarry_pipeline.yml`

Trigger: `workflow_run` on completion of "Daily CRE Digest" workflow (App 1).

```yaml
on:
  workflow_run:
    workflows: ["Daily CRE Digest"]
    types: [completed]
  workflow_dispatch:
```

The workflow:
1. Checks out repo
2. Fetches `articles_handoff.json` from GitHub Pages
3. Downloads `comps.xlsx` and `contacts.xlsx` from the `quarry-data` artifact (previous run)
4. Runs `quarry_pipeline.py` to update both workbooks
5. Uploads updated workbooks as `quarry-data` artifact (retention: 365 days)
6. Emails updated workbooks as attachments (same Gmail SMTP as App 1)

**Note:** The workbooks live entirely in GitHub Actions artifacts — they are not committed to the repo (binary files). The `quarry-data` artifact accumulates all history.

## Traded.co Scraper (Deferred)
A separate, independent scraper for `traded.co` sale data. Not part of this pipeline. Will run on its own schedule and write to the Sales tab of comps.xlsx using the same append + duplicate-address-highlighting logic. Will be built once the App 1 handoff pipeline is stable.

## Market Insights Database (Deferred)
Aggregation of `market_intelligence` and `key_data_points` fields from App 1 articles into a third workbook. Deferred until comps and contacts are stable.

## Local Development
```bash
# Activate App 1's venv (The Quarry shares it for now)
venv\Scripts\activate

# Run The Quarry manually (requires articles_handoff.json in project root)
python quarry/quarry_pipeline.py

# Required env vars (same .env file as App 1)
GMAIL_APP_PASSWORD=...   # for emailing workbooks
```

## Dependencies
- `openpyxl` — Excel read/write
- `requests` — fetch handoff JSON from GitHub Pages
- `python-dotenv` — .env loading
