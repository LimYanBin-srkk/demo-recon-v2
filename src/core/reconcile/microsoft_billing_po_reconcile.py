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

import json
import re
from datetime import datetime, timezone
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

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

    # -- Step C: Tiered multi-way matching (one-to-one consumption) ---------
    # Tiers run in order; exact-amount tiers first, then near (±0.01).
    # Each CSV row may be consumed by at most one PO line.
    po_line_items = po_data.get("line_items") or []
    consumed_indices: set[int] = set()

    # Pre-parse all PO items
    po_parsed: list[dict] = []
    for po_item in po_line_items:
        po_desc   = po_item.get("description") or ""
        po_start, po_end = _parse_contract_dates(po_desc)
        po_parsed.append({
            "po_line_no":  str(po_item.get("line_no") or ""),
            "po_desc":     po_desc,
            "po_amount":   _parse_amount(po_item.get("amount") or "0"),
            "po_start":    po_start,
            "po_end":      po_end,
        })

    # Result slots — None means unmatched so far
    match_results: list[dict | None] = [None] * len(po_parsed)

    def _try_match(po_idx: int, *, need_date: bool, need_name: bool, exact_amount: bool) -> bool:
        """Attempt to match po_parsed[po_idx] against an unconsumed CSV row."""
        p = po_parsed[po_idx]
        for idx, r in enumerate(customer_csv_rows):
            if idx in consumed_indices:
                continue
            # Date check
            if need_date:
                if not p["po_start"] or not p["po_end"]:
                    continue
                if r.get("charge_start") != p["po_start"] or r.get("charge_end") != p["po_end"]:
                    continue
            # Amount check
            diff = abs(r["amount"] - p["po_amount"])
            if exact_amount:
                if diff != Decimal("0"):
                    continue
            else:
                if diff > AMOUNT_TOL or diff == Decimal("0"):
                    continue  # exact already handled; only 0 < diff ≤ 0.01
            # Name check
            if need_name:
                if not _product_name_match(r["product"], p["po_desc"]):
                    continue
            return _assign(po_idx, idx, r, diff)
        return False

    def _assign(po_idx: int, csv_idx: int, r: dict, diff: Decimal) -> bool:
        consumed_indices.add(csv_idx)
        p = po_parsed[po_idx]

        # Determine status based on what matched
        has_date = (p["po_start"] and p["po_end"]
                    and r.get("charge_start") == p["po_start"]
                    and r.get("charge_end") == p["po_end"])
        has_name = _product_name_match(r["product"], p["po_desc"])
        is_near  = diff > Decimal("0")

        if has_date and has_name and not is_near:
            status = "found_exact"
        elif has_date and not is_near:
            status = "found_date_amount"
        elif has_name and not is_near:
            status = "found_name_amount"
        elif has_date and has_name and is_near:
            status = "found_near"
        elif has_date and is_near:
            status = "found_date_near"
        else:
            status = "found_name_near"

        match_results[po_idx] = {
            "po_line_no":      p["po_line_no"],
            "po_description":  p["po_desc"].replace("\n", " | "),
            "po_amount":       str(p["po_amount"]),
            "billing_amount":  str(r["amount"]),
            "order_ids":       r.get("order_id", ""),
            "contract_start":  r.get("charge_start", ""),
            "contract_end":    r.get("charge_end", ""),
            "match_status":    status,
        }
        return True

    # Run tiers in priority order
    tiers = [
        # Exact amount tiers
        {"need_date": True,  "need_name": True,  "exact_amount": True},   # T1
        {"need_date": True,  "need_name": False, "exact_amount": True},   # T2
        {"need_date": False, "need_name": True,  "exact_amount": True},   # T3
        # Near amount tiers (0 < diff ≤ 0.01)
        {"need_date": True,  "need_name": True,  "exact_amount": False},  # T1n
        {"need_date": True,  "need_name": False, "exact_amount": False},  # T2n
        {"need_date": False, "need_name": True,  "exact_amount": False},  # T3n
    ]

    for tier in tiers:
        for po_idx in range(len(po_parsed)):
            if match_results[po_idx] is not None:
                continue
            _try_match(po_idx, **tier)

    # Build final results
    results = []
    for po_idx, p in enumerate(po_parsed):
        if match_results[po_idx] is not None:
            results.append(match_results[po_idx])
        else:
            results.append({
                "po_line_no":     p["po_line_no"],
                "po_description": p["po_desc"].replace("\n", " | "),
                "po_amount":      str(p["po_amount"]),
                "match_status":   "not_found_in_billing",
            })

    found_count     = sum(1 for r in results if r["match_status"] != "not_found_in_billing")
    not_found_count = sum(1 for r in results if r["match_status"] == "not_found_in_billing")

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


def reconcile_po_summary(
    billing_result: dict,
    extraction_dir,
    po_line_results: "list[dict] | None" = None,
) -> dict:
    """
    Stage 3 — Customer-level PO coverage summary.

    Groups by customer, then by PO.  For each PO shows:
      - how many line items matched vs. unmatched
      - matched PO amount vs billing amount and variance
      - unmatched PO amount (items not in this billing — may appear next month)
      - billing amount NOT covered by any PO for this customer

    Parameters
    ----------
    billing_result   : output of microsoft_billing_reconcile.reconcile()
    extraction_dir   : Path to folder containing *.json extraction files
    po_line_results  : list of per-PO Stage 2 result dicts
                       Each dict must have keys: po_match_meta + line_items
                       (the output of reconcile_po).  Optional.

    Returns
    -------
    {
      "customers": [
        {
          "customer_name":        str,
          "billing_total":        str,   # sum of ALL Stage 1 CSV rows for this customer
          "billing_in_po":        str,   # sum of billing_amount from matched PO lines
          "billing_not_in_po":    str,   # billing_total - billing_in_po
          "po_coverage_pct":      float, # billing_in_po / billing_total * 100
          "pos": [
            {
              "po_number":            str,
              "po_date":              str,
              "po_total":             str,   # total_incl_tax from PO file
              "total_items":          int,
              "matched_count":        int,
              "unmatched_count":      int,
              "matched_po_amount":    str,   # sum of po_amount for matched lines
              "matched_billing_amt":  str,   # sum of billing_amount for matched lines
              "variance":             str,   # matched_billing_amt - matched_po_amount
              "unmatched_po_amount":  str,   # sum of po_amount for unmatched lines
              "status":               "⏳ Pending Approval",
              "line_items":           [...], # same as Stage 2 line_items (may be empty)
              "_mock":                bool,
            }
          ],
        }
      ],
      "total_customers":   int,
      "generated_at":      str,
    }
    """
    extraction_dir = Path(extraction_dir)

    # ── Collect unique CustomerNames + billing amounts from Stage 1 ───────
    excel_customers: dict[str, str] = {}       # norm → original
    customer_billing: dict[str, Decimal] = {}  # norm → total billing amount
    # Also track which Stage 1 billing rows belong to each customer
    # so we can compute billing_in_po properly
    customer_billing_rows: dict[str, list[dict]] = {}  # norm → list of match dicts

    for item in billing_result.get("line_items", []):
        for m in item.get("matches", []):
            cn = (m.get("Customer Name") or "").strip()
            if not cn:
                continue
            norm = _normalise_name(cn)
            excel_customers[norm] = cn
            amt = _parse_amount(m.get("Amount") or "0")
            customer_billing[norm] = customer_billing.get(norm, Decimal("0")) + amt
            customer_billing_rows.setdefault(norm, []).append(m)

    # ── Load all PO extraction JSONs, group by delivery_recipient ─────────
    po_files: list[dict] = []
    if extraction_dir.exists():
        for jf in sorted(extraction_dir.glob("*.json")):
            try:
                data = json.loads(jf.read_text(encoding="utf-8"))
                if data.get("po_number") and data.get("delivery_recipient"):
                    po_files.append(data)
            except Exception:
                pass

    # Build stage-2-result index: po_number → line_items list
    s2_by_po: dict[str, list[dict]] = {}
    if po_line_results:
        for s2 in po_line_results:
            meta = s2.get("po_match_meta") or {}
            pn   = meta.get("po_number") or ""
            if pn:
                s2_by_po[pn] = s2.get("line_items", [])

    # ── Build customer cards ───────────────────────────────────────────────
    customers_out: list[dict] = []

    for norm_cn, original_cn in sorted(excel_customers.items(), key=lambda x: x[1]):
        billing_total = customer_billing.get(norm_cn, Decimal("0"))

        # Find all PO files that match this customer
        matched_pos: list[dict] = []
        for po_data in po_files:
            recipient = (po_data.get("delivery_recipient") or {}).get("name") or ""
            norm_rec  = _normalise_name(recipient)
            if norm_rec == norm_cn or norm_cn in norm_rec or norm_rec in norm_cn:
                matched_pos.append(po_data)

        billing_in_po = Decimal("0")
        po_cards: list[dict] = []

        for po_data in matched_pos:
            po_number  = po_data.get("po_number") or ""
            po_date    = po_data.get("po_date") or ""
            po_total   = _parse_amount(po_data.get("total_incl_tax") or "0")
            is_mock    = bool(po_data.get("_mock"))

            # Get Stage 2 line items for this PO (if available)
            line_items = s2_by_po.get(po_number, [])

            # Aggregate per PO
            matched_lines   = [r for r in line_items if r.get("match_status") != "not_found_in_billing"]
            unmatched_lines = [r for r in line_items if r.get("match_status") == "not_found_in_billing"]

            matched_po_amt   = sum(_parse_amount(r.get("po_amount") or "0")      for r in matched_lines)
            matched_bill_amt = sum(_parse_amount(r.get("billing_amount") or "0") for r in matched_lines)
            unmatched_po_amt = sum(_parse_amount(r.get("po_amount") or "0")      for r in unmatched_lines)

            # If no Stage 2 data, treat all PO line items as unmatched
            if not line_items:
                raw_lines = po_data.get("line_items") or []
                unmatched_po_amt = sum(_parse_amount(li.get("amount") or "0") for li in raw_lines)
                unmatched_lines  = [
                    {
                        "po_line_no":     str(li.get("line_no") or ""),
                        "po_description": (li.get("description") or "").replace("\n", " | "),
                        "po_amount":      str(_parse_amount(li.get("amount") or "0")),
                        "match_status":   "not_found_in_billing",
                    }
                    for li in raw_lines
                ]
                line_items = unmatched_lines

            billing_in_po += matched_bill_amt
            variance = matched_bill_amt - matched_po_amt

            po_cards.append({
                "po_number":           po_number,
                "po_date":             po_date,
                "po_total":            str(po_total),
                "total_items":         len(line_items),
                "matched_count":       len(matched_lines),
                "unmatched_count":     len(unmatched_lines),
                "matched_po_amount":   str(matched_po_amt),
                "matched_billing_amt": str(matched_bill_amt),
                "variance":            str(variance),
                "unmatched_po_amount": str(unmatched_po_amt),
                "status":              "⏳ Pending Approval",
                "line_items":          line_items,
                "_mock":               is_mock,
            })

        billing_not_in_po = billing_total - billing_in_po
        coverage_pct = float(billing_in_po / billing_total * 100) if billing_total else 0.0

        customers_out.append({
            "customer_name":     original_cn,
            "billing_total":     str(billing_total),
            "billing_in_po":     str(billing_in_po),
            "billing_not_in_po": str(billing_not_in_po),
            "po_coverage_pct":   round(coverage_pct, 1),
            "pos":               po_cards,
        })

    return {
        "customers":       customers_out,
        "total_customers": len(customers_out),
        "generated_at":    datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    }
