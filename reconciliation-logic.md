# Microsoft Billing Reconciliation — Algorithm Reference

Three-stage matching pipeline:
1. **Billing × CSV** — match billing statement lines against Partner Center export
2. **Billing × POs** — match Stage 1 result against all extracted Purchase Orders (multi-PO, approval-gated)
3. **Review & Approve** — user confirms matches; `billed_detail` is persisted into PO extraction JSONs

---

## Stage 1 — Billing Tax Invoice × Partner Center CSV

**Input:** Microsoft Billing extraction JSON (`invoice.tax_invoice.line_items`) + Partner Center Excel/CSV
**Output:** Per-line-item match status with matched CSV rows and claimed row indices

### Pre-processing: Build CSV Name Lookup

> ⚠️ **Note:** Credit notes (`creditNote` / `credit note` charge type) are **currently ignored** — excluded from all matching, sums, and totals.

For every non-credit-note CSV row:

- Normalise product name: `lowercase + collapse whitespace`
- Index into `name_lookup[normalised_product]`
- Retain: `Product`, `Customer Name`, `Amount`, `Charge Type`, `Order ID`, `Charge Start`, `Charge End`

### Matching Steps (per billing line item)

```
Billing line item: product name + amount
         │
         ▼
 ┌─────────────────────────────────────────────────────────┐
 │  Step 1 — Product Name Match                            │
 │                                                         │
 │  Normalise: lowercase + collapse whitespace             │
 │  1a. Exact lookup in name_lookup                        │
 │  1b. Fallback: scan all keys for substring containment  │
 │      in either direction                                │
 └─────────────────────────────────────────────────────────┘
         │
         └─ Zero rows → 🔴 No Match  (stop)
         │
         ▼
 ┌─────────────────────────────────────────────────────────┐
 │  Step 2 — Amount Filter (1-to-1)                        │
 │                                                         │
 │  Keep name-matched rows where:                          │
 │      |csv_amount − billing_amount| ≤ 0.01               │
 └─────────────────────────────────────────────────────────┘
         │
         ├─ Exactly 1 row, diff = 0        → 🟢 Matched (Exact)  (stop)
         ├─ Exactly 1 row, 0 < diff ≤ 0.01 → 🟠 Near Match       (stop)
         │
         └─ 0 or 2+ rows → Flow 2
                  │
                  ▼
         ┌──────────────────────────────────────────────────────┐
         │  Step 3 — Flow 2: Per-Customer Sum                   │
         │                                                      │
         │  Group all name-matched rows by Customer Name.       │
         │  For each group: sum all row amounts.                │
         │  Check if |group_sum − billing_amount| ≤ 0.01        │
         └──────────────────────────────────────────────────────┘
                  │
                  ├─ Exactly 1 group passes → 🔵 Sum Match   (stop)
                  └─ 0 or 2+ groups pass    → 🟡 Ambiguous
```

### Two-Pass Deduplication

Runs after all line items receive an initial status:

| Pass | Action |
|------|--------|
| Pass 1 | Record which CSV `row_idx` is claimed by each green / near / Sum Match item |
| Pass 2 | Any `row_idx` claimed by 2+ items → **all** those claimants downgraded to 🟡 Ambiguous |

Surviving (non-contested) items retain their `row_idx` as exclusive.

### Result Statuses

| Status | Code | When |
|--------|------|------|
| 🟢 Matched (Exact) | `green` | Single CSV row, name + amount match exactly |
| 🟠 Near Match | `near` | Single CSV row, name matches, amount within ±0.01 |
| 🔵 Sum Match | `flow2` | Multiple CSV rows; one customer's grouped sum matches |
| 🟡 Ambiguous | `yellow` | No clear match, or contested row |
| 🔴 No Match | `red` | No CSV row shares the product name |

### Charge Type Handling

| CSV Value | Label | Behaviour |
|-----------|-------|-----------|
| `new` | New | Included |
| `renew` | Renewal | Included |
| `cycleCharge` | Cycle Charge | Included |
| `addQuantity` | Add Quantity | Included |
| `moveQuantity` | Move Quantity | Included |
| `customerCredit` | ⚠️ Customer Credit | Included, flagged (negative amount) |
| `cancelImmediate` | ⚠️ Cancellation | Included, flagged (negative refund) |
| `creditNote` / `credit note` | *(excluded)* | **Ignored — not matched, not summed** |

### Column Auto-Detection Priority

| Field | Priority order |
|-------|---------------|
| Product | `ProductName` › `SkuName` › `product` › `service` › `description` › `item` › `sku` |
| Amount | `Subtotal` › `amount` › `total` › `cost` › `price` › `extended` |
| Charge Start | `ChargeStartDate` › `charge start` › `startdate` › `orderdate` › `date` |
| Charge End | `ChargeEndDate` › `charge end` |
| Customer | `CustomerName` › `customer name` › `client name` › `customer` › `client` |
| Charge Type | `ChargeType` › `charge type` › `type` |
| Order ID | `OrderId` › `order id` › `order_id` |

---

## Stage 2 — `reconcile_po_all`: Multi-PO Line-Item Matching

**Module:** `core/reconcile/microsoft_billing_po_reconcile.py`
**Function:** `reconcile_po_all(billing_data, billing_result, extraction_dir, status_file)`

**Input:**
- `billing_data` — Microsoft Billing extraction JSON (provides `billing_number`)
- `billing_result` — output of Stage 1 `reconcile()` (contains CSV match rows per billing line)
- `extraction_dir` — path to `output/extraction/` (scans all `*.json` files)
- `status_file` — path to `output/reconciliation/po_billing_status.json`

**Output:** `reconcile_po_all` result dict (not written to disk — pending approval)

### PO Selection & Skip Logic

For every extraction JSON that has `po_number` + `delivery_recipient`:

```
Compute current_billed_status:
    "No"      → 0 lines have billed_detail
    "Partial" → some lines have billed_detail
    "Done"    → all lines have billed_detail

If current_billed_status == "Done":
    po_billing_nos = all billing_no values stored in billed_detail across all lines
    if billing_number NOT in po_billing_nos:
        SKIP — PO fully claimed by a different billing
    else:
        CONTINUE — current billing owns at least one line; allow re-reconcile

If current_billed_status in ("No", "Partial"):
    CONTINUE — reconcile unbilled lines
```

### Step A — Customer Name Match

```
PO: delivery_recipient (string or dict.name)
         │
         Normalise: strip all non-alphanumeric chars + lowercase
         │
         ▼
Compare against all unique CustomerName values from Stage 1 CSV matches.

1. Exact normalised match → customer identified
2. Substring containment fallback

No customer match → all lines → not_found_in_billing
```

### Step B — Build Customer CSV Row Pool

Flatten all Stage 1 CSV match rows for the matched customer into an ordered list.
Deduplicate rows via `_seen_row_keys` before adding to the pool — a row is identified by the tuple `(order_id, product, amount, charge_start, charge_end)` to prevent the same CSV row appearing twice if it appears in multiple billing item matches.

Row consumption across POs is tracked by `global_consumed_row_keys` (same tuple), ensuring each CSV row is assigned to at most one PO line item per run.

### Step C — Tiered Multi-Way Matching (one-to-one)

Six tiers run in priority order. Exact-amount tiers run first; near-amount (±0.01) tiers run last. Each CSV row can be consumed at most once across all POs in the run.

**Matching criteria:**
- **Date match** — `ChargeStartDate == PO start AND ChargeEndDate == PO end` (dates parsed from PO description via regex `DD/MM/YYYY - DD/MM/YYYY`)
- **Amount exact** — `csv_amount == po_amount` (diff = 0)
- **Amount near** — `0 < |csv_amount − po_amount| ≤ 0.01`
- **Product name** — substring containment + distinctive-token fallback (stop-words excluded)

```
For each unbilled PO line (6 passes over unconsumed CSV rows):

T1:  date ✓  name ✓  exact ✓  → found_exact
T2:  date ✓           exact ✓  → found_date_amount
T3:           name ✓  exact ✓  → found_name_amount
T1n: date ✓  name ✓  near  ✓  → found_near
T2n: date ✓           near  ✓  → found_date_near
T3n:          name ✓  near  ✓  → found_name_near

Still unmatched                → not_found_in_billing
```

Lines that already have `billed_detail` (from a prior approval) are always reported as `already_billed` and are never re-matched.

### Stage 2 Result Structure (per PO)

```json
{
  "po_number": "PO-2025-001",
  "po_date": "2025-01-10",
  "po_file": "CompanyName.json",
  "delivery_recipient": "Company Name Sdn Bhd",
  "customer_match": true,
  "current_billed_status": "Partial",
  "proposed_billed_status": "Done",
  "billing_numbers": ["INV-2025-001"],
  "total_items": 4,
  "already_billed": 1,
  "newly_matched": 3,
  "still_unbilled": 0,
  "line_items": [
    {
      "po_line_no": "1",
      "po_description": "...",
      "po_amount": "1234.56",
      "already_billed": false,
      "existing_billed_detail": null,
      "match_status": "found_exact",
      "pending_billed_detail": {
        "billing_no": "INV-2025-001",
        "billing_amount": "1234.56",
        "match_status": "found_exact",
        "order_id": "ORD-000001",
        "charge_start": "2025-01-01",
        "charge_end": "2025-12-31",
        "matched_at": "2026-04-16T08:00:00Z"
      }
    }
  ]
}
```

### Match Status Colour Coding (UI)

| Row colour | Status codes |
|-----------|-------------|
| Grey | `already_billed` |
| Green | `found_exact`, `found_date_amount`, `found_name_amount`, `found_near`, `found_date_near`, `found_name_near` |
| Red | `not_found_in_billing` |

---

## Backlog

The following items are planned but not yet implemented.

### Multi-PO Billing Matching

Currently the engine matches one PO at a time against a single billing. When the same billing contains line items that span **multiple POs** for the same customer (e.g. overlapping product SKUs across PO `S1-001` and `S1-002`), row assignment is first-come-first-served by PO scan order. A proper multi-PO solver should:

- Pool all unbilled line items across all POs for the matched customer simultaneously.
- Solve the assignment problem globally (e.g. bipartite matching or Hungarian algorithm) so that total variance is minimised.
- Respect tier priority: exact matches take precedence over near matches across all POs before any row is consumed.
- Report which rows are shared candidates and surface ambiguity to the user when a CSV row could satisfy lines in more than one PO.

### Post-Approval Behaviour

The Approve / Reject buttons in Stage 2 are currently **demo-only** (no data is written). When implemented, clicking **Approve** should:

1. Write `billed_detail` into each newly matched line item inside the PO extraction JSON (`output/extraction/<po_file>`).
2. Recompute the PO `billed_status` (`No` → `Partial` → `Done`) based on how many lines now have `billed_detail`.
3. Rebuild `output/reconciliation/po_billing_status.json` to reflect the updated statuses for all POs.
4. Register the billing number against the PO so future runs with the same billing skip already-approved lines (instead of re-matching them).
5. Return an approval summary showing: previous status → new status, number of newly billed items, and total billing amount written for each PO.

Clicking **Reject** should record the rejection decision without writing any `billed_detail`, and optionally surface a reason input for audit trail.

---

## Tolerance & Normalisation

All amount comparisons: **`|diff| ≤ 0.01`** (MYR/USD rounding tolerance, `Decimal` arithmetic)

| Function | Transformation | Used in |
|----------|---------------|---------|
| `_normalise(text)` | lowercase + collapse whitespace | Stage 1 product name matching |
| `_normalise_name(text)` | strip all non-alphanumeric + lowercase | Stage 2 customer name matching |
| `_parse_amount(value)` | strip commas/spaces, handle `(x)` = negative, round to 0.01 | All amount parsing |
| `_parse_contract_dates(text)` | regex `DD/MM/YYYY - DD/MM/YYYY` → ISO dates | Stage 2 date matching |

---

## Key Stop-Words (Product Name Matching)

Generic terms excluded from product token matching to prevent false positives across Microsoft SKUs:

```
microsoft, office, plan, plans, annual, contract, dates, renewal,
add, ons, from, with, the, for, and, p1y, p3y, p1, p3,
365, year, commitment
```
