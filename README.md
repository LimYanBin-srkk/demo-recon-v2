# SRKK Document Intelligence — ReconApp

A Streamlit-based financial document intelligence platform built for the **SRKK Financial Department**. The system automates document ingestion, AI-powered data extraction, and cross-document reconciliation to replace manual data entry and review.

---

## Purpose

The financial department receives a high volume of documents — vendor invoices, utility bills, rental agreements, bank statements, credit notes, statements of account, and hotel/travel receipts. This app provides:

1. **Automated OCR & Extraction** — upload a PDF and receive structured JSON data extracted by an AI vision model.
2. **Report View** — all processed documents are mapped into a unified spreadsheet format for review and export.
3. **Reconciliation** — cross-match extracted documents to surface discrepancies in amounts, dates, and reference numbers.

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
    │       └── microsoft_billing_po_reconcile.py  # Stage 2: Stage 1 result × PO
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
    │   │   └── doc_teams.json          # Document → team assignment map
    │   └── uploads/
    │       └── microsoft_billing/      # Reference Excel/CSV uploads
    │
    └── output/                         # All pipeline outputs (auto-created)
        ├── ocr/                        # Raw OCR JSON + token usage log
        ├── extraction/                 # Structured extraction JSON per document
        ├── images/                     # PDF page images per document
        └── reconciliation/             # Reconciliation output per document
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

Three-way reconciliation between a **Microsoft Billing statement**, a **Partner Center Excel/CSV**, and a **Purchase Order**.

---

#### Stage 1 — Billing × Excel/CSV

Matches each billing line item from the extracted Microsoft Billing document against rows in the reference Excel/CSV file.

##### Matching Steps (per billing line item)

```
Billing line item (product name + amount)
        │
        ▼
Step 1 — Product Name Match
        │  Normalise both sides (lowercase, collapse whitespace).
        │  Try exact match, then partial/contains.
        │
        ├─ No CSV row matches the product name → 🔴 No Match
        │
        ▼
Step 2 — Amount Filter (1-to-1)
        │  Keep CSV rows where |csv_amount − billing_amount| ≤ 0.01
        │
        ├─ Exactly 1 row passes → 🟢 Matched (Exact)
        │
        └─ 0 or 2+ rows pass → Flow 2
                │
                ▼
Step 3 — Flow 2: Per-Customer Sum
                Group name-matched rows by Customer Name.
                Sum each customer's amounts.
                │
                ├─ Exactly 1 customer's sum matches → 🔵 Sum Match
                └─ Otherwise → 🟡 Ambiguous
```

##### Result Statuses

| Status | Colour | When it occurs |
|--------|--------|----------------|
| 🟢 **Matched (Exact)** | Green | Single CSV row matches by product name AND amount (within ±0.01) |
| 🔵 **Sum Match** | Blue | Multiple CSV rows share the product name; one customer's summed amounts equal the billing amount |
| 🟠 **Near Match (±0.01)** | Orange | Amount differs by ≤ 0.01 — rounding adjustment |
| 🟡 **Ambiguous** | Yellow | No single row or customer sum matches, or a matched row is contested by multiple billing items — human review required |
| 🔴 **No Match** | Red | No CSV row shares the product name (exact or partial) |

##### Column Auto-Detection

The engine auto-detects columns in the uploaded Excel/CSV by scanning headers for these keywords (in priority order):

| Field | Keywords |
|-------|----------|
| Product name | `ProductName`, `SkuName`, `product`, `service`, `description`, `item`, `sku` |
| Amount | `Subtotal`, `amount`, `total`, `cost`, `price`, `extended` |
| Charge start date | `ChargeStartDate`, `charge start`, `startdate`, `orderdate`, `date` |
| Charge end date | `ChargeEndDate`, `charge end` |
| Customer | `CustomerName`, `customer name`, `client name`, `customer`, `client` |
| Charge type | `ChargeType`, `charge type`, `type` |
| Order ID | `OrderId`, `order id`, `order_id` |

##### Charge Type Handling

| CSV Charge Type | Label | Notes |
|----------------|-------|-------|
| `new` | New | Standard new subscription |
| `renew` | Renewal | Subscription renewal |
| `cycleCharge` | Cycle Charge | Regular billing cycle |
| `addQuantity` | Add Quantity | Seat addition |
| `moveQuantity` | Move Quantity | Quantity transfer |
| `customerCredit` | ⚠️ Customer Credit | Negative — included in sum |
| `cancelImmediate` | ⚠️ Cancellation | Negative refund — included in sum |
| `creditNote` / `credit note` | *(excluded)* | Excluded entirely |

---

#### Stage 2 — Stage 1 Result × Purchase Order

Matches each PO line item against the Stage 1 billing result for the same customer using **contract dates + amount** (no product name matching).

##### Matching Logic

For each PO line item:
1. Parse `Contract Dates: DD/MM/YYYY - DD/MM/YYYY` from the PO description.
2. Find the first unconsumed CSV row for the matched customer where:
   - `ChargeStartDate` = PO contract start date
   - `ChargeEndDate` = PO contract end date
   - `|csv_amount − po_amount|` ≤ 0.01
3. Each CSV row is consumed at most once (one-to-one assignment).

##### Result Categories

| Category | Meaning |
|----------|---------|
| ✅ Found in Billing | PO line matched a CSV row by customer + charge dates + amount |
| ❌ Not Found in Billing | No CSV row matched the contract dates and amount |

##### Stage 2 Result Columns

| Column | Description |
|--------|-------------|
| PO Line # | Line number from the Purchase Order |
| PO Description | Description text from the PO line item |
| PO Amount | Amount stated in the PO |
| Billing Amount | Matched CSV row amount |
| Order ID | `OrderId` from the matched CSV row |

---

## Document Types Supported

| Label | Display Name |
|-------|-------------|
| `commercial_invoice` | Commercial Invoice |
| `srkk_vendor_invoice` | SRKK - Vendor Invoice |
| `srkk_purchase_order` | SRKK - Purchase Order |
| `srkk_microsoft_billing` | SRKK - Microsoft Billing |
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

    (for Microsoft Billing)
    │
    ▼
Stage 1: microsoft_billing_reconcile   ← reference Excel/CSV
    │
    ▼
Stage 2: microsoft_billing_po_reconcile ← Purchase Order extraction
```


---

## Purpose

The financial department receives a high volume of documents — vendor invoices, utility bills, rental agreements, bank statements, credit notes, statements of account, and hotel/travel receipts. This app provides:

1. **Automated OCR & Extraction** — upload a PDF and receive structured JSON data extracted by an AI vision model.
2. **Report View** — all processed documents are mapped into a unified spreadsheet format for review and export.
3. **Reconciliation** — select any previously processed documents and cross-match their extracted fields to surface discrepancies (e.g. amount mismatches, missing fields, duplicate invoices).

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
└── src/
    ├── app.py                          # Main Streamlit application (entry point)
    ├── Readme.md                       # This file
    │
    ├── core/                           # Core processing pipeline
    │   ├── __init__.py
    │   ├── pdf_to_images.py            # PDF → PNG page images (PyMuPDF, 300 DPI)
    │   ├── ocr_agent.py                # Vision OCR via Azure OpenAI
    │   ├── orchestrator.py             # Classifier → extraction agent router
    │   └── page_tracker.py             # OCR page quota tracker (persistent JSON)
    │
    ├── agents/                         # Extraction agents, one per document type
    │   ├── __init__.py                 # Shared Azure OpenAI client & helpers
    │   ├── classifier.py               # LLM-based document type classifier
    │   ├── extraction_invoice.py       # Commercial invoice
    │   ├── extraction_srkk_vendor_invoice.py   # SRKK vendor invoice
    │   ├── extraction_srkk_customer_invoice.py # SRKK customer invoice
    │   ├── extraction_utility.py       # Utility bills (TNB, Air Selangor, etc.)
    │   ├── extraction_rental.py        # Rental / lease agreements
    │   ├── extraction_hotel.py         # Hotel receipts
    │   ├── extraction_travel.py        # Travel receipts
    │   ├── extraction_soa.py           # Statement of account
    │   └── extraction_bank.py          # Bank statements
    │
    ├── docs/
    │   ├── page_usage.json             # Persistent OCR page quota store
    │   ├── database/
    │   │   └── doc_teams.json          # Document → team assignment map
    │   └── uploads/                    # Uploaded source PDFs
    │       └── reconciliation/
    │           ├── scenario1/          # Sales / Income / Balance Excel files
    │           └── scenario2/          # Invoice & Ledger PDFs
    │
    └── output/                         # All pipeline outputs (auto-created)
        ├── ocr/                        # Raw OCR JSON + token usage log
        ├── extraction/                 # Structured extraction JSON per document
        └── reconciliation/             # Reconciliation output per document label
```

---

## Application Pages

### 🏠 Dashboard
Overview of OCR page quota usage. Shows pages consumed vs. the 1,000-page limit broken down by source, with visual progress bar and KPI tiles.

### 📤 Document Processing
The primary workflow page:
1. **Upload** a PDF (or image) file.
2. The app converts the PDF to page images (300 DPI).
3. Azure OpenAI Vision performs OCR and returns structured text with confidence scores.
4. The document is **classified** into one of 9 supported types.
5. A type-specific **extraction agent** pulls structured fields (invoice number, vendor, amounts, dates, etc.).
6. Results are saved to `output/ocr/` and `output/extraction/` with the upload timestamp embedded.
7. Previously uploaded documents are listed below with team assignment and processing status.

### 🔍 OCR Viewer
Browse and inspect raw OCR output JSON files. Shows per-section confidence scores and raw extracted text for any previously processed document.

### 📊 Extraction Viewer
Repository-style view of all extraction results. Browse by document, inspect individual field values, and review the full extraction JSON.

### 📋 Report Format
Unified spreadsheet view mapping all extracted documents to a standard financial report schema. Features:
- **Upload date filter** — filter documents by when they were uploaded.
- **Document type filter** — show/hide specific types (Invoice, Utility, Rental, etc.).
- **Company name filter** — narrow to specific vendors.
- **Utility ↔ Rental matching** — automatically matches utility bills to corresponding rental agreements.
- **Excel export** — download the filtered table as `.xlsx`.

### 🔄 Reconciliation
Select up to 3 previously processed documents and cross-match their extracted fields. Surfaces discrepancies in amounts, dates, and reference numbers across documents (e.g. comparing a vendor invoice against a statement of account).

---

#### ☁️ Microsoft Billing × Excel/CSV Reconciliation

When one of the selected documents is a **SRKK - Microsoft Billing** statement, you can upload a reference Excel/CSV file (e.g. the Microsoft Partner Center reconciliation file exported by the reseller) and match each billing line item against it.

---

##### Step-by-Step Matching Process

For every line item in the **billing invoice** the engine performs the following steps in order:

```
Billing line item  (product name + amount)
        │
        ▼
Step 1 ── Product Name Match
        │  Normalise both sides: lowercase + collapse whitespace
        │  Try exact match first, then partial/contains match
        │
        ├─ No CSV row shares the product name
        │        └─► 🔴 No Match — stop
        │
        ▼
Step 2 ── Amount Filter
        │  Keep only CSV rows where |csv_amount − billing_amount| ≤ 0.01
        │
        ├─ Exactly 1 CSV row passes
        │        └─► 🟢 Matched (1-to-1) — stop
        │
        └─ 0 or 2+ CSV rows pass → proceed to Flow 2
                │
                ▼
Step 3 ── Flow 2: Per-Customer Sum
                │  Group all name-matched CSV rows by Customer Name
                │  Sum each customer's amounts for this product
                │  Check if |customer_sum − billing_amount| ≤ 0.01
                │
                ├─ Exactly 1 customer's sum matches
                │        └─► 🔵 Sum Match (Flow 2) — stop
                │
                └─ 0 or 2+ customers match
                         └─► 🟡 Ambiguous — all name-matched rows listed
```

---

##### Result Statuses

| Status | Colour | When it occurs |
|--------|--------|----------------|
| 🟢 **Matched (Exact)** | Green | Single CSV row matches by product name AND amount exactly equals the billing amount |
| 🔵 **Sum Match (Flow 2)** | Blue | Multiple CSV rows share the product name; one customer's summed amounts exactly equal the billing amount (reseller pattern: same product billed to multiple customers) |
| 🟡 **Ambiguous** | Yellow | No single row or customer sum matches the billing amount — human review required |
| 🔴 **No Match** | Red | No CSV row contains the product name (exact or partial) |

---

##### Two-Pass Deduplication

After the initial pass, the engine checks for **contested rows** — CSV rows claimed as a 1-to-1 green match by more than one billing item (e.g. the same subscription line appears twice in the invoice):

1. **Pass 1** — determine status and claimed CSV row indices for each billing item.
2. **Pass 2** — any CSV row index claimed by 2+ green items causes **all** those claimants to be downgraded to 🟡 Ambiguous. Their display table excludes rows already held exclusively by surviving green items.

---

##### Column Auto-Detection (Reference File)

The engine automatically detects columns in the uploaded Excel/CSV by scanning column names for keywords:

| Field | Keywords searched (in priority order) |
|-------|-------------------------------------|
| Product name | `ProductName`, `product name`, `SkuName`, `sku name`, `product`, `service`, `description`, `item`, `sku` |
| Amount | `Subtotal`, `amount`, `total`, `cost`, `price`, `extended` |
| Date | `ChargeStartDate`, `charge start`, `startdate`, `orderdate`, `date`, `period`, `month` |
| Customer | `CustomerName`, `customer name`, `client name`, `customer`, `client`, `account`, `company` |
| Charge type | `ChargeType`, `charge type`, `type` |

---

##### Charge Type Labels & Flags

| CSV Charge Type | Label shown | Notes |
|----------------|------------|-------|
| `new` | New | Standard new subscription/purchase |
| `renew` | Renewal | Subscription renewal |
| `cycleCharge` | Cycle Charge | Regular monthly billing |
| `addQuantity` | Add Quantity | Seat/unit addition |
| `moveQuantity` | Move Quantity | Quantity transfer |
| `customerCredit` | ⚠️ Customer Credit | Included in match; negative amount flags a credit adjustment |
| `cancelImmediate` | ⚠️ Cancellation | Included in match; mid-cycle cancellation refund |
| `creditNote` / `credit note` | *(excluded)* | Formal credit note documents — excluded entirely (separate invoices) |

Rows flagged with ⚠️ appear in the match table with an orange prefix so reviewers can identify why the CSV total may differ from the billing amount.

---

##### Result Table (per billing item)

Each billing line item is shown as a collapsible section containing:

| Column | Source | Description |
|--------|--------|-------------|
| **Billing Amount** | Extracted invoice | Amount as stated in the Microsoft billing PDF |
| **CSV Total Amount** | Sum of matched rows | Arithmetic sum of all matched CSV rows (including adjustments) |
| Customer Name | Reference file | Customer the row was billed to |
| CSV Amount | Reference file | Individual row amount |
| Charge Type | Reference file | Transaction type (with ⚠️ for adjustments) |
| Date | Reference file | Charge start date |

---

## Document Types Supported

| Label | Display Name | Description |
|-------|-------------|-------------|
| `commercial_invoice` | Commercial Invoice | Standard vendor invoices |
| `srkk_vendor_invoice` | SRKK - Vendor Invoice | SRKK-format vendor invoices |
| `srkk_purchase_order` | SRKK - Purchase Order | SRKK purchase order documents |
| `srkk_microsoft_billing` | SRKK - Microsoft Billing | Microsoft / Azure cloud billing statements |
| `utility` | Utility | TNB, Air Selangor, and other utility bills |
| `rental` | Rental | Rental / lease / tenancy agreements |
| `hotel` | Hotel | Hotel accommodation receipts |
| `travel` | Travel | Travel and transport receipts |
| `soa` | Statement of Account | Statement of account |
| `bank_statement` | Bank Statement | Bank statements |
| `credit_note` | Credit Note | Credit notes |

---

## Key Configuration

| Setting | Location | Default |
|---------|----------|---------|
| Azure OpenAI credentials | `.env` at project root | — |
| OCR page quota | `core/page_tracker.py` | 1,000 pages |
| Quota store | `docs/page_usage.json` | auto-created |
| Team assignment map | `docs/database/doc_teams.json` | auto-created |

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
output/extraction/<name>.json  → report view + reconciliation
```
