"""
Microsoft Billing Reconciliation — v2
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Matches every line item from a Microsoft Billing extraction JSON against
rows in a Microsoft Partner Center reconciliation Excel/CSV file.

Matching Process (per billing line item)
-----------------------------------------
Step 1 — Product Name Match
    Normalise both sides (lowercase + collapse whitespace).
    Try exact match first, then partial/contains match.
    No match at all → status: red (stop).

Step 2 — Amount Filter
    From the name-matched CSV rows, keep those where
        |csv_row_amount − billing_amount| ≤ AMOUNT_TOL (0.01).
    Exactly 1 row passes → status: green (stop).
    0 or 2+ rows pass → proceed to Flow 2.

Step 3 — Flow 2: Per-Customer Sum
    Group all name-matched CSV rows by CustomerName.
    Sum each group's amounts for the product.
    Exactly 1 customer whose sum satisfies the ≤ AMOUNT_TOL check → status: flow2 (stop).
    0 or 2+ customers match → status: yellow.

Statuses
--------
green  : single CSV row, name + amount match (exact within AMOUNT_TOL).
flow2  : multiple CSV rows; one customer's grouped sum matches (within AMOUNT_TOL).
yellow : ambiguous — zero or multiple customers match in Flow 2, OR a
         green/flow2 row index is contested by 2+ billing items (dedup).
red    : no CSV row shares the product name (exact or partial).

Two-Pass Deduplication
----------------------
Pass 1: determine initial status + claimed CSV row indices per billing item.
Pass 2: any CSV row index claimed by 2+ green/flow2 items causes ALL those
        claimants to be downgraded to yellow.  Downgraded items' display
        tables exclude rows held exclusively by surviving green/flow2 items.

Column Auto-Detection (reference file)
---------------------------------------
Product  : ProductName > product name > SkuName > sku name > product > ...
Amount   : Subtotal > amount > total > cost > price > extended
Date     : ChargeStartDate > charge start > startdate > orderdate > date > ...
Customer : CustomerName > customer name > client name > customer > ...
ChargeType: ChargeType > charge type > type

Charge Type Handling
--------------------
Excluded (never matched):  creditNote, credit note  (formal credit documents)
Included + flagged (⚠️):   customerCredit, cancelImmediate
    — negative-amount rows or known adjustment types receive an ⚠️ prefix
      so reviewers can see adjustments that affect the net total.

Output Schema
-------------
{
  "reconcile_meta": {
    "billing_number", "vendor_name", "billing_period", "document_date",
    "reference_file", "ref_product_col", "ref_amount_col", "ref_date_col",
    "ref_customer_col", "ref_charge_type_col", "generated_at",
    "total_line_items", "green", "flow2", "yellow", "red"
  },
  "line_items": [
    {
      "product":     str,
      "line_amount": str,          # billing amount (serialised Decimal)
      "status":      green|flow2|yellow|red,
      "matches": [                 # [] for red; sorted by CustomerName+Date
        {
          "Product":       str,
          "Customer Name": str,
          "Date":          str,    # YYYY-MM-DD
          "Amount":        str,    # serialised Decimal
          "Charge Type":   str     # friendly label, ⚠️-prefixed if adjustment
        }
      ]
    }
  ]
}
"""

import json
import re
from collections import defaultdict
from datetime import datetime, timezone
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pandas as pd


AMOUNT_TOL = Decimal("0.01")

_EXCLUDED_CHARGE_TYPES = {"creditnote", "credit note"}

_CT_LABELS = {
    "new":             "New",
    "renew":           "Renewal",
    "cycleccharge":    "Cycle Charge",
    "cyclecharge":     "Cycle Charge",
    "addquantity":     "Add Quantity",
    "movequantity":    "Move Quantity",
    "trialconversion": "Trial Conversion",
    "customercredit":  "Customer Credit",
    "cancelimmediate": "Cancellation",
}
_ADJ_TYPES = {"customercredit", "cancelimmediate"}


# -- Helpers -------------------------------------------------------------------

def _parse_amount(value) -> Decimal:
    if value is None:
        return Decimal("0")
    s = str(value).strip().replace(",", "").replace(" ", "")
    if not s or s in ("-", "."):
        return Decimal("0")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        d = Decimal(s)
        if not d.is_finite():
            return Decimal("0")
        return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal("0")


def _fmt(d: Decimal) -> str:
    return str(d)


def _normalise(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "").lower().strip())


def _detect_col(df, *candidates):
    for cand in candidates:
        for col in df.columns:
            if cand.lower() in col.lower():
                return col
    return None


def _charge_label(raw_ct: str, amt: Decimal) -> str:
    label = _CT_LABELS.get(raw_ct.lower(), raw_ct) if raw_ct else ""
    is_adj = amt < Decimal("0") or raw_ct.lower() in _ADJ_TYPES
    if is_adj:
        label = ("⚠️ " + label) if label else "⚠️ Adjustment"
    return label


def _ser(entry: dict) -> dict:
    """Serialise entry for JSON output: Decimal -> str, drop internal row_idx."""
    return {
        k: (_fmt(v) if isinstance(v, Decimal) else v)
        for k, v in entry.items()
        if k != "row_idx"
    }


# -- Core reconcile function ---------------------------------------------------

def reconcile(billing_data: dict, spreadsheet_path) -> dict:
    spreadsheet_path = Path(spreadsheet_path)

    if spreadsheet_path.suffix.lower() == ".csv":
        ref_df = pd.read_csv(spreadsheet_path)
    else:
        ref_df = pd.read_excel(spreadsheet_path)

    ref_df.columns = [str(c).strip() for c in ref_df.columns]

    prod_col        = _detect_col(ref_df, "productname", "product name", "skuname", "sku name", "product", "service", "description", "item", "sku")
    amount_col      = _detect_col(ref_df, "subtotal", "amount", "total", "cost", "price", "extended")
    date_col        = _detect_col(ref_df, "chargestartdate", "charge start", "startdate", "orderdate", "date", "period", "month")
    cust_col        = _detect_col(ref_df, "customername", "customer name", "client name", "customer", "client", "account", "company")
    charge_type_col = _detect_col(ref_df, "chargetype", "charge type", "type")
    order_id_col    = _detect_col(ref_df, "orderid", "order id", "order_id")
    charge_start_col = _detect_col(ref_df, "chargestartdate", "charge start")
    charge_end_col   = _detect_col(ref_df, "chargeenddate", "charge end")

    name_lookup: dict = defaultdict(list)

    for row_idx, row in ref_df.iterrows():
        if charge_type_col:
            ct_raw = str(row[charge_type_col] or "").strip().lower()
            if ct_raw in _EXCLUDED_CHARGE_TYPES:
                continue

        prod_val = str(row[prod_col] or "").strip() if prod_col else ""
        if not prod_val:
            continue

        amt_val    = _parse_amount(row[amount_col]) if amount_col else Decimal("0")
        raw_date   = str(row[date_col] or "").strip() if date_col else ""
        fmt_date   = raw_date.split("T")[0] if "T" in raw_date else raw_date
        raw_ct       = str(row[charge_type_col] or "").strip() if charge_type_col else ""
        order_id     = str(row[order_id_col] or "").strip() if order_id_col else ""
        charge_start = str(row[charge_start_col] or "").strip().split("T")[0] if charge_start_col else ""
        charge_end   = str(row[charge_end_col] or "").strip().split("T")[0] if charge_end_col else ""

        entry = {
            "row_idx":        row_idx,
            "Product":        prod_val,
            "Customer Name":  str(row[cust_col] or "").strip() if cust_col else "",
            "Date":           fmt_date,
            "Amount":         amt_val,
            "Charge Type":    _charge_label(raw_ct, amt_val),
            "Order ID":       order_id,
            "Charge Start":   charge_start,
            "Charge End":     charge_end,
        }
        name_lookup[_normalise(prod_val)].append(entry)

    def _find_name_matches(norm_prod: str) -> list:
        matches = list(name_lookup.get(norm_prod, []))
        if not matches:
            for key, entries in name_lookup.items():
                if norm_prod and (norm_prod in key or key in norm_prod):
                    matches.extend(entries)
        return matches

    def _sort_by_customer(entries: list) -> list:
        return sorted(entries, key=lambda e: (e["Customer Name"].lower(), e["Date"]))

    def _flow2(name_matches: list, billing_dec: Decimal) -> dict:
        by_cust: dict = defaultdict(list)
        for e in name_matches:
            by_cust[e["Customer Name"]].append(e)

        winners = [
            (cust, rows)
            for cust, rows in by_cust.items()
            if abs(sum(r["Amount"] for r in rows) - billing_dec) <= AMOUNT_TOL
        ]

        if len(winners) == 1:
            _, rows = winners[0]
            cust_sum = sum(r["Amount"] for r in rows)
            _st = "flow2" if cust_sum == billing_dec else "near"
            return {
                "status":       _st,
                "matches":      [_ser(r) for r in rows],
                "claimed_idxs": {r["row_idx"] for r in rows},
            }

        sorted_entries = _sort_by_customer(name_matches)
        return {
            "status":       "yellow",
            "matches":      [_ser(e) for e in sorted_entries],
            "claimed_idxs": set(),
        }

    invoice    = billing_data.get("invoice", {}) or {}
    line_items = (invoice.get("tax_invoice") or {}).get("line_items") or []

    pass1 = []

    for item in line_items:
        product     = item.get("product") or ""
        billing_dec = _parse_amount(item.get("amount") or "0")
        norm_prod   = _normalise(product)

        name_matches = _find_name_matches(norm_prod)

        if not name_matches:
            pass1.append({
                "product":           product,
                "line_amount":       _fmt(billing_dec),
                "status":            "red",
                "matches":           [],
                "claimed_idxs":      set(),
                "_all_name_matches": [],
            })
            continue

        amt_matches = [
            e for e in name_matches
            if abs(e["Amount"] - billing_dec) <= AMOUNT_TOL
        ]

        if len(amt_matches) == 1:
            e = amt_matches[0]
            _st = "green" if e["Amount"] == billing_dec else "near"
            pass1.append({
                "product":           product,
                "line_amount":       _fmt(billing_dec),
                "status":            _st,
                "matches":           [_ser(e)],
                "claimed_idxs":      {e["row_idx"]},
                "_all_name_matches": name_matches,
            })
        else:
            flow = _flow2(name_matches, billing_dec)
            pass1.append({
                "product":           product,
                "line_amount":       _fmt(billing_dec),
                "_all_name_matches": name_matches,
                **flow,
            })

    idx_claimants: dict = defaultdict(list)
    for i, r in enumerate(pass1):
        if r["status"] in ("green", "near", "flow2"):
            for idx in r["claimed_idxs"]:
                idx_claimants[idx].append(i)

    contested = {idx for idx, claimants in idx_claimants.items() if len(claimants) > 1}
    downgrade_set = {i for idx in contested for i in idx_claimants[idx]}

    survivor_idxs: set = set()
    for i, r in enumerate(pass1):
        if r["status"] in ("green", "near", "flow2") and i not in downgrade_set:
            survivor_idxs |= r["claimed_idxs"]

    results = []
    for i, r in enumerate(pass1):
        if i in downgrade_set:
            available = [
                e for e in r["_all_name_matches"]
                if e["row_idx"] not in survivor_idxs
            ]
            sorted_entries = _sort_by_customer(available)
            results.append({
                "product":     r["product"],
                "line_amount": r["line_amount"],
                "status":      "yellow",
                "matches":     [_ser(e) for e in sorted_entries],
            })
        else:
            results.append({
                "product":     r["product"],
                "line_amount": r["line_amount"],
                "status":      r["status"],
                "matches":     r["matches"],
            })

    green_count  = sum(1 for r in results if r["status"] == "green")
    near_count   = sum(1 for r in results if r["status"] == "near")
    flow2_count  = sum(1 for r in results if r["status"] == "flow2")
    yellow_count = sum(1 for r in results if r["status"] == "yellow")
    red_count    = sum(1 for r in results if r["status"] == "red")

    return {
        "reconcile_meta": {
            "billing_number":      invoice.get("billing_number"),
            "vendor_name":         invoice.get("vendor_name"),
            "billing_period":      invoice.get("billing_period"),
            "document_date":       invoice.get("document_date"),
            "reference_file":      spreadsheet_path.name,
            "ref_product_col":     prod_col,
            "ref_amount_col":      amount_col,
            "ref_date_col":        date_col,
            "ref_customer_col":    cust_col,
            "ref_charge_type_col": charge_type_col,
            "generated_at":        datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "total_line_items":    len(results),
            "green":               green_count,
            "near":                near_count,
            "flow2":               flow2_count,
            "yellow":              yellow_count,
            "red":                 red_count,
        },
        "line_items": results,
    }


# -- CLI entry point -----------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Reconcile Microsoft Billing JSON against Excel/CSV")
    parser.add_argument("--billing", required=True, help="Path to extraction JSON")
    parser.add_argument("--ref",     required=True, help="Path to reference Excel/CSV")
    parser.add_argument("--out",     default="reconcile_result.json", help="Output JSON path")
    args = parser.parse_args()

    with open(args.billing, encoding="utf-8") as f:
        billing = json.load(f)

    result = reconcile(billing, args.ref)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    meta = result["reconcile_meta"]
    print(f"Done -> {args.out}")
    print(f"  Total : {meta['total_line_items']}")
    print(f"  Green : {meta['green']}")
    print(f"  Yellow: {meta['yellow']}")
    print(f"  Red   : {meta['red']}")
