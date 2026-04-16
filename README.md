# SRKK Document Intelligence — ReconApp

A Streamlit-based financial document intelligence platform built for the **SRKK Financial Department**. The system automates document ingestion, AI-powered data extraction, and three-way reconciliation between Microsoft Billing statements, Partner Center CSV exports, and Purchase Orders.

---

## Purpose

The financial department receives a high volume of documents — vendor invoices, utility bills, rental agreements, bank statements, credit notes, statements of account, and hotel/travel receipts. This app provides:

1. **Automated OCR & Extraction** — upload a PDF and receive structured JSON data extracted by an AI vision model.
2. **Report View** — all processed documents are mapped into a unified spreadsheet format for review and export.
3. **Two-Stage Reconciliation** — match Microsoft Billing statements against Partner Center CSV and Purchase Orders, with per-PO match results and demo Approve/Reject controls.

---

## Getting Started

### Prerequisites

- Python 3.11+
- An Azure OpenAI deployment with a vision-capable model (e.g. `gpt-4o`)
- A `.env` file at the project root:

```
AZURE_OPENAI_ENDPOINT=https://<your-resource>.openai.azure.com/
AZURE_OPENAI_API_KEY=<your-key>
AZURE_OPENAI_DEPLOYMENT=<deployment-name>
AZURE_OPENAI_API_VERSION=2025-04-01-preview
```

### Install & Run

```bash
pip install -r requirements.txt
streamlit run src/app.py
```

---

## Project Structure

```
srkk_ReconApp/
├── requirements.txt
├── README.md
├── reconciliation-logic.md      # Algorithm reference for all reconciliation stages
└── src/
    ├── app.py                          # Main Streamlit application (entry point)
    │
    ├── core/                           # Core processing pipeline
    │   ├── __init__.py
    │   ├── pdf_to_images.py            # PDF → PNG page images (PyMuPDF, 300 DPI)
    │   ├── ocr_agent.py                # Vision OCR via Azure OpenAI
    │   ├── orchestrator.py             # Classifier → extraction agent router
    │   ├── page_tracker.py             # OCR page quota tracker (persistent JSON)
    │   └── reconcile/
    │       ├── microsoft_billing_reconcile.py     # Stage 1: Billing × Excel/CSV
    │       └── microsoft_billing_po_reconcile.py  # Stage 2: Multi-PO matching (approval pending backlog)
    │
    ├── agents/                         # Extraction agents, one per document type
    │   ├── __init__.py                 # Shared Azure OpenAI client & helpers
    │   ├── classifier.py               # LLM-based document type classifier
    │   ├── extraction_invoice.py
    │   ├── extraction_utility.py
    │   ├── extraction_rental.py
    │   ├── extraction_hotel.py
    │   ├── extraction_travel.py
    │   ├── extraction_soa.py
    │   ├── extraction_bank.py
    │   ├── srkk_microsoft_billing.py
    │   ├── srkk_purchase_order.py
    │   └── srkk_vendor_invoice.py
    │
    ├── docs/
    │   ├── page_usage.json             # Persistent OCR page quota store
    │   ├── database/
    │   │   ├── doc_teams.json          # Document → team assignment map
    │   │   └── srkk_po_dir.json        # PO registry — filenames of extracted PO JSONs
    │   └── uploads/
    │       └── microsoft_billing/      # Reference Excel/CSV uploads
    │
    └── output/                         # All pipeline outputs (auto-created)
        ├── ocr/                        # Raw OCR JSON + token usage log
        ├── extraction/                 # Structured extraction JSON per document
        ├── images/                     # PDF page images per document
        └── reconciliation/
            └── po_billing_status.json  # PO billed status tracker (auto-maintained)
```

---

## Application Pages

### 🏠 Dashboard
Overview of OCR page quota usage. Shows pages consumed vs. the 1,000-page limit with a visual progress bar and KPI tiles.

### 📤 Document Processing
The primary workflow page:
1. Upload a PDF (or image) file.
2. The app converts the PDF to page images (300 DPI).
3. Azure OpenAI Vision performs OCR and returns structured text with confidence scores.
4. The document is classified into one of the supported types.
5. A type-specific extraction agent pulls structured fields (invoice number, vendor, amounts, dates, etc.).
6. Results are saved to `output/ocr/` and `output/extraction/`.
7. Previously uploaded documents are listed below with team assignment and processing status.

### 🔍 OCR Viewer
Browse and inspect raw OCR output JSON files. Shows per-section confidence scores and raw extracted text for any previously processed document.

### 📊 Extraction Viewer
Repository-style view of all extraction results. Browse by document and inspect individual field values.

### 📋 Report Format
Unified spreadsheet view mapping all extracted documents into a standard financial report schema. Features:
- Upload date, document type, and company name filters.
- Utility ↔ Rental automatic matching.
- Excel export (`.xlsx`).

### 🔄 Reconciliation — Microsoft Billing

Three-stage reconciliation between a **Microsoft Billing statement**, a **Partner Center Excel/CSV**, and all extracted **Purchase Orders**.

---

#### Stage 1 — Billing × Excel/CSV

Matches each billing line item from the extracted Microsoft Billing document against rows in the reference Excel/CSV file.

| Status | When |
|--------|------|
| 🟢 Matched (Exact) | Single CSV row, name + amount match exactly |
| 🔵 Sum Match | Multiple CSV rows; one customer's summed total matches billing amount |
| 🟠 Near Match | Amount differs by ≤ 0.01 (rounding) |
| 🟡 Ambiguous | No clear single match, or row claimed by multiple billing items |
| 🔴 No Match | No CSV row shares the product name |

See [reconciliation-logic.md](reconciliation-logic.md) for the full matching algorithm.

---

#### Stage 2 — Billing × Purchase Orders (Multi-PO Scan)

Automatically scans all PO JSON files registered in `src/docs/database/srkk_po_dir.json` and matches each unbilled PO line item against the Stage 1 result for the corresponding customer.

**PO Skip Logic (billing-number aware):**

| PO Status | Current billing owns a line? | Action |
|-----------|------------------------------|--------|
| `No` | N/A | Reconcile all lines |
| `Partial` | N/A | Reconcile unbilled lines only |
| `Done` | Yes (re-run same billing) | Reconcile — already-billed lines shown as `already_billed` |
| `Done` | No (different billing owns it) | Skip — fully claimed by another billing |

**Filters available in the UI:**
- **Customer** — select a specific delivery recipient; defaults to All
- **Match Status** — filter to `✅ Matched` (has at least one `found_*` line) or `❌ Not Matched`

Each PO expander shows:
- Green `🟢` prefix in the title when the PO has newly matched lines
- Billing number(s) already recorded on this PO
- **Approve / Reject** buttons _(demo only — no data is written)_
- 5 metrics: Billing Amount / PO Amount (Billed) / Coverage / Variance / Unmatched PO
- Line-item table with colour-coded match status:

| Row colour | Status |
|-----------|--------|
| Grey | `⬜ Already Billed` |
| Green | `✅ Exact Match`, `✅ Date+Amount`, `✅ Name+Amount` |
| Orange | `🟠 Near (all/date/name, ±0.01)` |
| Red | `❌ Not Found` |

**Match tiers (priority order):**

| Tier | Criteria | Status Code |
|------|----------|-------------|
| T1 | Date + Name + Exact amount | `found_exact` |
| T2 | Date + Exact amount | `found_date_amount` |
| T3 | Name + Exact amount | `found_name_amount` |
| T1n | Date + Name + Near amount (±0.01) | `found_near` |
| T2n | Date + Near amount | `found_date_near` |
| T3n | Name + Near amount | `found_name_near` |
| — | No tier matched | `not_found_in_billing` |

---

## Document Types Supported

| Label | Display Name |
|-------|-------------|
| `commercial_invoice` | Commercial Invoice |
| `srkk_vendor_invoice` | SRKK — Vendor Invoice |
| `srkk_purchase_order` | SRKK — Purchase Order |
| `srkk_microsoft_billing` | SRKK — Microsoft Billing |
| `utility` | Utility Bill |
| `rental` | Rental Agreement |
| `hotel` | Hotel Receipt |
| `travel` | Travel Receipt |
| `soa` | Statement of Account |
| `bank_statement` | Bank Statement |
| `credit_note` | Credit Note |

---

## Key Configuration

| Setting | Location | Default |
|---------|----------|---------|
| Azure OpenAI credentials | `.env` at project root | — |
| OCR page quota | `core/page_tracker.py` | 1,000 pages |
| Quota store | `docs/page_usage.json` | auto-created |
| Team assignment map | `docs/database/doc_teams.json` | auto-created |
| PO billing status | `output/reconciliation/po_billing_status.json` | auto-created |

---

## Architecture Overview

```
PDF upload
    │
    ▼
pdf_to_images (core/)          → PNG pages at 300 DPI
    │
    ▼
ocr_agent (core/)              → Azure OpenAI Vision → OCR JSON
    │
    ▼
classifier (agents/)           → document type label
    │
    ▼
extraction_<type> (agents/)    → structured extraction JSON
    │
    ▼
output/extraction/<name>.json

    (for Microsoft Billing — three-stage reconciliation)
    │
    ▼
Stage 1: microsoft_billing_reconcile   ← reference Excel/CSV
    │        Billing × CSV → per-line match statuses
    ▼
Stage 2: reconcile_po_all              ← scans ALL PO extraction JSONs
    │        Skip logic: Done+other billing → skip; else reconcile unbilled lines
    │        Tiered matching (T1–T3n) with consumed-row tracking
    │        Returns pending_billed_detail per line (NOT written to disk yet)
    ▼
Stage 3: approve_po_results            ← user-triggered approval
             Writes billed_detail into PO JSONs
             Updates po_billing_status.json
```
