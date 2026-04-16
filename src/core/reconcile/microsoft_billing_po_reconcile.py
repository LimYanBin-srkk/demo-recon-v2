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


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _delivery_name(po_data: dict) -> str:
    delivery = po_data.get("delivery_recipient")
    if isinstance(delivery, dict):
        return str(delivery.get("name") or "")
    return str(delivery or "")


def _line_billed_detail(line_item: dict):
    bd = line_item.get("billed_detail")
    if isinstance(bd, dict) and bd:
        return bd
    return None


def _po_billing_nos(po_data: dict) -> set[str]:
    """Return all billing_no values stored in billed_detail across all line items."""
    nos: set[str] = set()
    for li in po_data.get("line_items") or []:
        bd = _line_billed_detail(li)
        if bd and bd.get("billing_no"):
            nos.add(str(bd["billing_no"]))
    return nos


def _compute_po_status(po_data: dict) -> tuple[str, int, int, int]:
    lines = po_data.get("line_items") or []
    total = len(lines)
    billed = sum(1 for li in lines if _line_billed_detail(li) is not None)
    unbilled = max(total - billed, 0)
    if total == 0:
        status = "No"
    elif billed == total:
        status = "Done"
    elif billed > 0:
        status = "Partial"
    else:
        status = "No"
    return status, total, billed, unbilled


def _load_status_file(status_file: Path) -> dict:
    if not status_file.exists():
        return {"last_updated": _utc_now(), "po_statuses": {}}
    try:
        raw = json.loads(status_file.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            return {"last_updated": _utc_now(), "po_statuses": {}}
        if "po_statuses" not in raw or not isinstance(raw["po_statuses"], dict):
            raw["po_statuses"] = {}
        return raw
    except Exception:
        return {"last_updated": _utc_now(), "po_statuses": {}}


def _build_status_entry(po_data: dict, po_file: str) -> dict:
    status, total, billed, unbilled = _compute_po_status(po_data)
    return {
        "po_number": str(po_data.get("po_number") or ""),
        "po_file": po_file,
        "delivery_recipient": _delivery_name(po_data),
        "billed_status": status,
        "total_items": total,
        "billed_items": billed,
        "unbilled_items": unbilled,
        "billing_numbers": sorted(_po_billing_nos(po_data)),
    }


def _collect_customer_names(billing_result: dict) -> set[str]:
    out: set[str] = set()
    for item in billing_result.get("line_items", []):
        for m in item.get("matches", []):
            cn = str(m.get("Customer Name") or "").strip()
            if cn:
                out.add(cn)
    return out


def _matched_customers(delivery_name: str, excel_customers: set[str]) -> list[str]:
    norm_delivery = _normalise_name(delivery_name)
    matched = [cn for cn in excel_customers if _normalise_name(cn) == norm_delivery]
    if matched:
        return matched
    return [
        cn for cn in excel_customers
        if _normalise_name(cn) in norm_delivery or norm_delivery in _normalise_name(cn)
    ]


def _po_dir_files(extraction_dir: Path) -> list[Path]:
    """Return ordered list of PO extraction file paths from srkk_po_dir.json registry.

    Falls back to glob scanning extraction_dir for any JSON with po_number if the
    registry file does not yet exist (backwards compatibility).
    """
    # Registry lives next to the extraction dir: src/docs/database/srkk_po_dir.json
    registry = extraction_dir.parent.parent / "docs" / "database" / "srkk_po_dir.json"
    if registry.exists():
        try:
            raw = json.loads(registry.read_text(encoding="utf-8"))
            po_file_names = list((raw.get("po_files") or {}).keys())
            return [
                extraction_dir / fname
                for fname in sorted(po_file_names)
                if (extraction_dir / fname).exists()
            ]
        except Exception:
            pass
    # Fallback: scan all JSONs (slower, for initial migration)
    return sorted(extraction_dir.glob("*.json"))


def reconcile_po_all(
    billing_data: dict,
    billing_result: dict,
    extraction_dir: Path,
    status_file: Path,
) -> dict:
    _ = billing_data

    billing_number = str((billing_data.get("invoice") or {}).get("billing_number") or "").strip()
    if not billing_number:
        raise ValueError("billing_data['invoice']['billing_number'] is required")

    extraction_dir = Path(extraction_dir)
    status_file = Path(status_file)

    status_doc = _load_status_file(status_file)
    po_statuses = dict(status_doc.get("po_statuses") or {})

    po_records: list[dict] = []

    if extraction_dir.exists():
        for jf in _po_dir_files(extraction_dir):
            try:
                po_data = json.loads(jf.read_text(encoding="utf-8"))
            except Exception:
                continue
            if not (po_data.get("po_number") and po_data.get("delivery_recipient")):
                continue

            po_statuses[str(po_data.get("po_number") or "")] = _build_status_entry(po_data, jf.name)
            po_records.append({"file": jf, "data": po_data})

    excel_customers = _collect_customer_names(billing_result)

    global_consumed_row_keys: set[tuple[str, str, str, str, str]] = set()
    po_results: list[dict] = []

    for rec in po_records:
        po_file = rec["file"]
        po_data = rec["data"]
        po_number = str(po_data.get("po_number") or "")
        po_date = str(po_data.get("po_date") or "")
        delivery_name = _delivery_name(po_data)

        current_status, total_items, billed_items, _ = _compute_po_status(po_data)
        if current_status == "Done":
            po_billing_nos = _po_billing_nos(po_data)
            if billing_number not in po_billing_nos:
                # All lines are billed by a different billing — skip
                continue
            # Current billing owns at least one line: allow re-reconcile
            # (already-billed lines will surface as already_billed in output)

        matched_customers = _matched_customers(delivery_name, excel_customers)
        customer_match = bool(matched_customers)
        matched_norms = {_normalise_name(cn) for cn in matched_customers}

        customer_rows: list[dict] = []
        _seen_row_keys: set[tuple] = set()
        for item in billing_result.get("line_items", []):
            for m in item.get("matches", []):
                if _normalise_name(m.get("Customer Name", "")) not in matched_norms:
                    continue
                _rk = (
                    str(m.get("Order ID") or ""),
                    str(m.get("Product") or ""),
                    str(m.get("Amount") or ""),
                    str(m.get("Charge Start") or ""),
                    str(m.get("Charge End") or ""),
                )
                if _rk in _seen_row_keys:
                    continue  # de-duplicate: same CSV row can appear in multiple billing item matches
                _seen_row_keys.add(_rk)
                customer_rows.append({
                    "product": str(m.get("Product") or ""),
                    "amount": _parse_amount(m.get("Amount", "0")),
                    "order_id": str(m.get("Order ID") or ""),
                    "charge_start": str(m.get("Charge Start") or ""),
                    "charge_end": str(m.get("Charge End") or ""),
                })

        line_items_out: list[dict] = []
        pending_matches = 0

        po_lines = po_data.get("line_items") or []
        parsed_unbilled: list[dict] = []
        parsed_to_output_idx: list[int] = []

        for li in po_lines:
            existing_bd = _line_billed_detail(li)
            base = {
                "po_line_no": str(li.get("line_no") or ""),
                "po_description": str(li.get("description") or "").replace("\n", " | "),
                "po_amount": str(_parse_amount(li.get("amount") or "0")),
            }
            if existing_bd is not None:
                line_items_out.append({
                    **base,
                    "already_billed": True,
                    "existing_billed_detail": existing_bd,
                    "match_status": "already_billed",
                    "pending_billed_detail": None,
                })
                continue

            parsed_unbilled.append({
                "po_line_no": base["po_line_no"],
                "po_desc": str(li.get("description") or ""),
                "po_amount": _parse_amount(li.get("amount") or "0"),
                "po_start": _parse_contract_dates(str(li.get("description") or ""))[0],
                "po_end": _parse_contract_dates(str(li.get("description") or ""))[1],
            })
            parsed_to_output_idx.append(len(line_items_out))
            line_items_out.append({
                **base,
                "already_billed": False,
                "existing_billed_detail": None,
                "match_status": "not_found_in_billing",
                "pending_billed_detail": None,
            })

        local_matched: set[int] = set()

        # Status derived strictly from the tier that produced the match,
        # not from re-evaluating attributes post-match.
        _TIER_STATUS: dict[tuple[bool, bool, bool], str] = {
            (True,  True,  True):  "found_exact",
            (True,  False, True):  "found_date_amount",
            (False, True,  True):  "found_name_amount",
            (True,  True,  False): "found_near",
            (True,  False, False): "found_date_near",
            (False, True,  False): "found_name_near",
        }

        def _row_key(r: dict) -> tuple[str, str, str, str, str]:
            return (
                str(r.get("order_id") or ""),
                str(r.get("product") or ""),
                str(r.get("amount") or ""),
                str(r.get("charge_start") or ""),
                str(r.get("charge_end") or ""),
            )

        def _assign_match(po_idx: int, row_idx: int, diff: Decimal,
                          need_date: bool, need_name: bool, exact_amount: bool) -> bool:
            nonlocal pending_matches
            p = parsed_unbilled[po_idx]
            r = customer_rows[row_idx]

            status = _TIER_STATUS[(need_date, need_name, exact_amount)]

            local_matched.add(row_idx)
            global_consumed_row_keys.add(_row_key(r))

            pending_billed_detail = {
                "billing_no": billing_number,
                "billing_amount": str(r["amount"]),
                "match_status": status,
                "order_id": str(r.get("order_id") or ""),
                "charge_start": str(r.get("charge_start") or ""),
                "charge_end": str(r.get("charge_end") or ""),
                "matched_at": _utc_now(),
            }

            out_idx = parsed_to_output_idx[po_idx]
            line_items_out[out_idx]["match_status"] = status
            line_items_out[out_idx]["pending_billed_detail"] = pending_billed_detail
            pending_matches += 1
            return True

        def _try_match(po_idx: int, *, need_date: bool, need_name: bool, exact_amount: bool) -> bool:
            p = parsed_unbilled[po_idx]
            for row_idx, r in enumerate(customer_rows):
                if row_idx in local_matched:
                    continue
                if _row_key(r) in global_consumed_row_keys:
                    continue

                if need_date:
                    if not p["po_start"] or not p["po_end"]:
                        continue
                    if r.get("charge_start") != p["po_start"] or r.get("charge_end") != p["po_end"]:
                        continue

                diff = abs(r["amount"] - p["po_amount"])
                if exact_amount:
                    if diff != Decimal("0"):
                        continue
                else:
                    if diff > AMOUNT_TOL or diff == Decimal("0"):
                        continue

                if need_name and not _product_name_match(r["product"], p["po_desc"]):
                    continue

                return _assign_match(po_idx, row_idx, diff, need_date, need_name, exact_amount)
            return False

        if customer_match and parsed_unbilled:
            tiers = [
                {"need_date": True,  "need_name": True,  "exact_amount": True},
                {"need_date": True,  "need_name": False, "exact_amount": True},
                {"need_date": False, "need_name": True,  "exact_amount": True},
                {"need_date": True,  "need_name": True,  "exact_amount": False},
                {"need_date": True,  "need_name": False, "exact_amount": False},
                {"need_date": False, "need_name": True,  "exact_amount": False},
            ]
            matched_flags: list[bool] = [False] * len(parsed_unbilled)
            for tier in tiers:
                for po_idx in range(len(parsed_unbilled)):
                    if matched_flags[po_idx]:
                        continue
                    matched_flags[po_idx] = _try_match(po_idx, **tier)

        already_billed = billed_items
        newly_matched = pending_matches
        still_unbilled = max(total_items - already_billed - newly_matched, 0)

        if still_unbilled == 0 and total_items > 0:
            proposed_status = "Done"
        elif already_billed + newly_matched > 0:
            proposed_status = "Partial"
        else:
            proposed_status = "No"

        po_results.append({
            "po_number": po_number,
            "po_date": po_date,
            "po_file": po_file.name,
            "delivery_recipient": delivery_name,
            "customer_match": customer_match,
            "current_billed_status": current_status,
            "proposed_billed_status": proposed_status,
            "billing_numbers": sorted(_po_billing_nos(po_data)),
            "total_items": total_items,
            "already_billed": already_billed,
            "newly_matched": newly_matched,
            "still_unbilled": still_unbilled,
            "line_items": line_items_out,
        })

    return {
        "billing_number": billing_number,
        "generated_at": _utc_now(),
        "po_results": po_results,
    }


def approve_po_results(
    reconcile_result: dict,
    extraction_dir: Path,
    status_file: Path,
) -> dict:
    extraction_dir = Path(extraction_dir)
    status_file = Path(status_file)
    status_file.parent.mkdir(parents=True, exist_ok=True)

    po_updates: list[dict] = []
    po_statuses: dict[str, dict] = {}

    for po_res in reconcile_result.get("po_results", []):
        po_file = str(po_res.get("po_file") or "")
        if not po_file:
            continue
        po_path = extraction_dir / po_file
        if not po_path.exists():
            continue

        try:
            po_data = json.loads(po_path.read_text(encoding="utf-8"))
        except Exception:
            continue

        prev_status, _, _, _ = _compute_po_status(po_data)
        newly_billed_items = 0

        po_line_items = po_data.get("line_items") or []
        for line_res in po_res.get("line_items", []):
            pending = line_res.get("pending_billed_detail")
            if not isinstance(pending, dict) or not pending:
                continue

            target_line_no = str(line_res.get("po_line_no") or "")
            target_desc = str(line_res.get("po_description") or "")
            target_amount = _parse_amount(line_res.get("po_amount") or "0")

            chosen_idx = None
            for idx, li in enumerate(po_line_items):
                if str(li.get("line_no") or "") != target_line_no:
                    continue
                if _line_billed_detail(li) is not None:
                    continue
                chosen_idx = idx
                break

            if chosen_idx is None:
                for idx, li in enumerate(po_line_items):
                    if _line_billed_detail(li) is not None:
                        continue
                    li_desc = str(li.get("description") or "").replace("\n", " | ")
                    li_amount = _parse_amount(li.get("amount") or "0")
                    if li_desc == target_desc and li_amount == target_amount:
                        chosen_idx = idx
                        break

            if chosen_idx is None:
                continue

            po_line_items[chosen_idx]["billed_detail"] = pending
            newly_billed_items += 1

        po_data["line_items"] = po_line_items
        po_path.write_text(json.dumps(po_data, indent=2, ensure_ascii=False), encoding="utf-8")

        new_status, _, _, _ = _compute_po_status(po_data)
        po_number = str(po_data.get("po_number") or po_res.get("po_number") or "")
        po_statuses[po_number] = _build_status_entry(po_data, po_file)
        po_updates.append({
            "po_number": po_number,
            "previous_status": prev_status,
            "new_status": new_status,
            "newly_billed_items": newly_billed_items,
        })

    # Refresh status file from all registered PO extraction files to keep it complete.
    if extraction_dir.exists():
        for jf in _po_dir_files(extraction_dir):
            try:
                po_data = json.loads(jf.read_text(encoding="utf-8"))
            except Exception:
                continue
            if not po_data.get("po_number"):
                continue
            po_number = str(po_data.get("po_number") or "")
            po_statuses[po_number] = _build_status_entry(po_data, jf.name)

    status_payload = {
        "last_updated": _utc_now(),
        "po_statuses": po_statuses,
    }
    status_file.write_text(json.dumps(status_payload, indent=2, ensure_ascii=False), encoding="utf-8")

    return {
        "approved_at": _utc_now(),
        "po_updates": po_updates,
    }
