"""
Microsoft Billing × Purchase Order Reconciliation  (v3)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Phase 2 of the three-way match.  Takes the output of
`microsoft_billing_reconcile.reconcile()` (which contains all CSV rows
per billing product) and a Purchase Order extraction JSON, then:

Step A — Customer Name Match
    Collect all unique CustomerName values from the billing_result matches.
    Normalise both sides by stripping all non-alphanumeric characters
    (so "SDN. BHD." == "SDN BHD"), then compare case-insensitively.

Step B — Line Item Match (PO-centric)
    Flatten all CSV match-rows from billing_result that belong to the
    matched customer.  For each PO line item:
      1. Find CSV rows whose product name partially matches the PO description.
      2. Sum their amounts net (negative rows for Cancel / CustomerCredit
         are already signed in the billing_result matches).
      3. If |net_sum − po_amount| ≤ AMOUNT_TOL → found_in_billing.
      4. Otherwise (customer found, amount differs, or no CSV rows) →
         not_found_in_billing.

Output Schema
-------------
{
  "po_match_meta": {
    "po_number":            str,
    "po_date":              str,
    "delivery_recipient":   str,
    "customer_name_match":  bool,
    "matched_customers":    [str],
    "generated_at":         str,
    "total_po_items":       int,
    "found_in_billing":     int,
    "not_found_in_billing": int
  },
  "line_items": [
    {
      "po_line_no":             str,
      "po_description":         str,
      "po_amount":              str,
      "net_csv_amount":         str | null,   # null if no CSV rows found
      "csv_row_count":          int,
      "match_status":           "found_in_billing" | "not_found_in_billing"
    }
  ]
}
"""

import re
from datetime import datetime, timezone
from decimal import ROUND_HALF_UP, Decimal

AMOUNT_TOL = Decimal("0.01")


def _normalise(text: str) -> str:
    """Lowercase + collapse whitespace (keeps symbols)."""
    return re.sub(r"\s+", " ", str(text or "").lower().strip())


def _normalise_name(text: str) -> str:
    """Lowercase + strip all non-alphanumeric chars (removes dots, commas, etc.)."""
    s = re.sub(r"[^a-z0-9\s]", "", str(text or "").lower())
    return re.sub(r"\s+", " ", s).strip()


def _parse_amount(value) -> Decimal:
    if value is None:
        return Decimal("0")
    s = str(value).strip().replace(",", "").replace(" ", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal("0")


# Generic billing terms that appear in almost every product — excluded from matching
_MATCH_STOP = frozenset({
    "microsoft", "office", "plan", "plans", "annual", "contract",
    "dates", "renewal", "add", "ons", "from", "with", "the", "for",
    "and", "p1y", "p3y", "p1", "p3", "365", "year", "commitment",
})


def _word_tokens(text: str) -> list:
    """Extract clean alphanumeric tokens from text."""
    return re.findall(r"[a-z0-9]+", _normalise(text))


def _product_name_match(csv_product: str, po_description: str) -> bool:
    """Return True if csv_product is recognisably the same product as po_description.

    Strategy:
    1. Substring containment in either direction (exact / near-exact names).
    2. Distinctive-token fallback: extract tokens from csv_product that are either
       long words (len >= 5) or short alphanumeric codes (e.g. 'e3', 'f3'), then
       check that at least one appears in po_description.  Generic billing stop-words
       ('microsoft', '365', 'annual', etc.) are excluded so they don't cause
       false-positive matches across different product SKUs.
    """
    norm_csv = _normalise(csv_product)
    norm_po  = _normalise(po_description)

    if not norm_csv or not norm_po:
        return False
    if norm_csv in norm_po or norm_po in norm_csv:
        return True

    tokens_csv = _word_tokens(csv_product)
    tokens_po  = set(_word_tokens(po_description))

    if not tokens_csv:
        return False

    # Distinctive: long words OR short mixed alpha-digit codes (e.g. e3, f3, m365)
    distinctive = [
        t for t in tokens_csv
        if t not in _MATCH_STOP and (
            len(t) >= 5
            or (len(t) >= 2
                and any(c.isdigit() for c in t)
                and any(c.isalpha() for c in t))
        )
    ]
    if not distinctive:
        # Fallback: any token with len >= 4 not in stop list
        distinctive = [t for t in tokens_csv if len(t) >= 4 and t not in _MATCH_STOP]
    if not distinctive:
        return False
    return any(t in tokens_po for t in distinctive)


def _parse_contract_dates(po_desc: str):
    """Extract contract start/end from 'DD/MM/YYYY - DD/MM/YYYY' in PO description.
    Returns (YYYY-MM-DD, YYYY-MM-DD) or (None, None)."""
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})", po_desc)
    if not m:
        return None, None
    def _iso(d):
        dd, mm, yyyy = d.split("/")
        return f"{yyyy}-{mm}-{dd}"
    return _iso(m.group(1)), _iso(m.group(2))


def reconcile_po(billing_data: dict, billing_result: dict, po_data: dict) -> dict:
    """
    Parameters
    ----------
    billing_data   : raw extraction JSON from srkk_microsoft_billing agent
                     (kept for API compatibility; not used for matching)
    billing_result : output of microsoft_billing_reconcile.reconcile()
    po_data        : raw extraction JSON from srkk_purchase_order agent
    """
    # -- PO metadata ----------------------------------------------------------
    po_number     = po_data.get("po_number") or ""
    po_date       = po_data.get("po_date") or ""
    delivery      = (po_data.get("delivery_recipient") or {})
    delivery_name = delivery.get("name") or ""

    # -- Step A: Customer name match (symbol-stripped) ------------------------
    excel_customers: set[str] = set()
    for item in billing_result.get("line_items", []):
        for m in item.get("matches", []):
            cn = m.get("Customer Name", "").strip()
            if cn:
                excel_customers.add(cn)

    norm_delivery = _normalise_name(delivery_name)
    matched_customers = [
        cn for cn in excel_customers
        if _normalise_name(cn) == norm_delivery
    ]
    # Broader fallback: substring containment after symbol stripping
    if not matched_customers:
        matched_customers = [
            cn for cn in excel_customers
            if (
                _normalise_name(cn) in norm_delivery
                or norm_delivery in _normalise_name(cn)
            )
        ]
    customer_match = bool(matched_customers)

    # -- Step B: Build flat list of CSV rows for matched customer -------------
    matched_norms = {_normalise_name(cn) for cn in matched_customers}

    customer_csv_rows: list[dict] = []
    for item in billing_result.get("line_items", []):
        for m in item.get("matches", []):
            if _normalise_name(m.get("Customer Name", "")) in matched_norms:
                customer_csv_rows.append({
                    "product":      m.get("Product", ""),
                    "amount":       _parse_amount(m.get("Amount", "0")),
                    "order_id":     m.get("Order ID", ""),
                    "charge_start": m.get("Charge Start", ""),
                    "charge_end":   m.get("Charge End", ""),
                })

    # -- Step C: Match each PO line item by contract dates + amount ----------
    # One-to-one: each CSV row may be consumed by at most one PO line.
    po_line_items = po_data.get("line_items") or []
    results = []
    found_count = 0
    not_found_count = 0
    consumed_indices: set[int] = set()

    for po_item in po_line_items:
        po_line_no  = str(po_item.get("line_no") or "")
        po_desc     = po_item.get("description") or ""
        po_amount   = _parse_amount(po_item.get("amount") or "0")

        po_start, po_end = _parse_contract_dates(po_desc)

        matched_row = None
        if po_start and po_end:
            for idx, r in enumerate(customer_csv_rows):
                if (
                    idx not in consumed_indices
                    and r.get("charge_start") == po_start
                    and r.get("charge_end") == po_end
                    and abs(r["amount"] - po_amount) <= AMOUNT_TOL
                ):
                    matched_row = (idx, r)
                    break

        if not matched_row:
            results.append({
                "po_line_no":     po_line_no,
                "po_description": po_desc.replace("\n", " | "),
                "po_amount":      str(po_amount),
                "match_status":   "not_found_in_billing",
            })
            not_found_count += 1
            continue

        idx, r = matched_row
        consumed_indices.add(idx)

        results.append({
            "po_line_no":      po_line_no,
            "po_description":  po_desc.replace("\n", " | "),
            "po_amount":       str(po_amount),
            "billing_amount":  str(r["amount"]),
            "order_ids":       r.get("order_id", ""),
            "contract_start":  r.get("charge_start", ""),
            "contract_end":    r.get("charge_end", ""),
            "match_status":    "found_in_billing",
        })
        found_count += 1

    return {
        "po_match_meta": {
            "po_number":            po_number,
            "po_date":              po_date,
            "delivery_recipient":   delivery_name,
            "customer_name_match":  customer_match,
            "matched_customers":    matched_customers,
            "generated_at":         datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "total_po_items":       len(results),
            "found_in_billing":     found_count,
            "not_found_in_billing": not_found_count,
        },
        "line_items": results,
    }
