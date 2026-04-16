"""LLM-based document classifier. Reads OCR JSON and returns a document type label."""

import re

from agents import client, DEPLOYMENT

CLASSIFIER_PROMPT = """
You are a document classifier. You receive OCR output (JSON) from a scanned financial document.

Your ONLY task is to classify the document into exactly ONE of these categories:

- "srkk_vendor_invoice"   : Invoices addressed TO SRKK (SRKK is the buyer/bill-to party)
- "srkk_purchase_order"   : Purchase orders issued BY SRKK to a supplier
- "srkk_microsoft_billing": Microsoft/cloud subscription billing statements addressed to SRKK
- "commercial_invoice"   : Product/goods invoices with barcodes, PO numbers, shipping terms, product line items
- "travel"               : Flight tickets, travel agency invoices, itineraries with passenger/routing info
- "rental"               : Rental/lease invoices, mall rent, service charges, tenancy billing
- "hotel"                : Hotel folios/invoices, room charges, accommodation bills
- "utility"              : Utility bills (electricity, water, gas), telecom bills (phone, internet, telephone)
- "soa"                  : Statements of account, outstanding balance summaries, aging reports
- "bank_statement"       : Bank account statements with transaction listings
- "credit_note"          : Credit notes / refund documents
- "unknown"              : Does not fit any category above

CLASSIFICATION HINTS:
- If you see "SRKK" or "Syarikat Reka Kod Kreatif" in the bill-to / recipient field → "srkk_vendor_invoice"
- If SRKK is the buyer and the document is a Purchase Order (has PO Number, Order From, Deliver To) → "srkk_purchase_order"
- If the document is from Microsoft with subscription/cloud billing charges addressed to SRKK → "srkk_microsoft_billing"
- If you see barcodes, PO numbers, trade terms (CFR, FOB), bill of lading → "commercial_invoice"
- If you see flight numbers, routing, passenger names, ticket numbers → "travel"
- If you see base rent, service charge, tenancy period, lot/unit number → "rental"
- If you see room charge, check-in/check-out, folio, guest name → "hotel"
- If you see meter reading, tariff, kWh, billing period, subscriber, telephone bill, water bill, sewerage, broadband usage, internet usage, wifi → "utility"
- IMPORTANT: If the invoice description mentions broadband, internet, telephone, electricity, or water reimbursement/usage, classify as "utility" even if issued by a property/rental company
- If you see list of invoices with aging (30/60/90 days), outstanding balance → "soa"
- If you see bank name, account transactions, debit/credit columns, running balance → "bank_statement"
- If you see "Credit Note" or "CN" in the title → "credit_note"
- Travel agency invoices for hotel bookings (with hotel name, check-in, room type but issued by a travel agent) → "travel"

Return ONLY the category label as a single word. No JSON. No explanation.
"""

VALID_LABELS = {
    "commercial_invoice", "travel", "rental", "hotel",
    "utility", "soa", "bank_statement", "credit_note",
    "srkk_vendor_invoice", "srkk_purchase_order", "srkk_microsoft_billing", "unknown",
}

ALIAS_MAP = {
    "invoice": "commercial_invoice",
    "commercial invoice": "commercial_invoice",
    "product invoice": "commercial_invoice",
    "goods invoice": "commercial_invoice",
    "tax invoice": "commercial_invoice",
    "travel invoice": "travel",
    "travel document": "travel",
    "flight ticket": "travel",
    "itinerary": "travel",
    "rental invoice": "rental",
    "lease invoice": "rental",
    "rent invoice": "rental",
    "hotel invoice": "hotel",
    "hotel folio": "hotel",
    "utility bill": "utility",
    "electricity bill": "utility",
    "water bill": "utility",
    "gas bill": "utility",
    "telecom bill": "utility",
    "phone bill": "utility",
    "internet bill": "utility",
    "telephone bill": "utility",
    "tel bill": "utility",
    "water bill": "utility",
    "sewerage bill": "utility",
    "statement of account": "soa",
    "account statement": "soa",
    "aging report": "soa",
    "bank statement": "bank_statement",
    "credit note": "credit_note",
    "credit memo": "credit_note",
    "cn": "credit_note",
}

KEYWORD_RULES = [
    ("srkk_vendor_invoice", ["srkk vendor invoice", "vendor invoice"]),
    ("srkk_purchase_order", ["srkk purchase order", "purchase order", "order from", "deliver to"]),
    ("srkk_microsoft_billing", ["microsoft billing", "microsoft regional sales", "azure subscription", "cloud billing", "microsoft 365"]),
    ("srkk_vendor_invoice", ["srkk vendor"]),
    ("credit_note", ["credit note", "credit memo", "refund note", "cn"]),
    ("utility", ["utility", "electricity", "water", "gas", "telecom", "telephone", "telefon", "meter", "kwh", "tariff", "sewerage", "indah water", "bil air", "bil telefon", "broadband", "wifi", "internet usage", "broadband usage", "monthly charges", "current charges", "billing period"]),
    ("bank_statement", ["bank statement", "running balance", "account transactions", "opening balance", "closing balance"]),
    ("soa", ["statement of account", "outstanding", "aging", "balance brought forward"]),
    ("hotel", ["hotel", "folio", "check-in", "check out", "room charge", "guest"]),
    ("travel", ["travel", "flight", "ticket", "itinerary", "passenger", "pnr", "routing"]),
    ("rental", ["rental", "lease", "tenancy", "base rent", "service charge", "lot no"]),
    ("commercial_invoice", ["commercial invoice", "tax invoice", "po number", "bill of lading", "barcode"]),
]


def _keyword_match_label(text: str) -> str:
    lower_text = (text or "").lower()
    for canonical, keywords in KEYWORD_RULES:
        if any(keyword in lower_text for keyword in keywords):
            return canonical
    return "unknown"


def _normalize_label(text: str) -> str:
    raw = (text or "").strip().lower().strip('"').strip("'")
    if raw in VALID_LABELS:
        return raw

    cleaned = re.sub(r"\s+", " ", re.sub(r"[_\-]+", " ", raw)).strip()
    if cleaned in ALIAS_MAP:
        return ALIAS_MAP[cleaned]

    for canonical, keywords in KEYWORD_RULES:
        if any(keyword in cleaned for keyword in keywords):
            return canonical

    return "unknown"


def classify_document(ocr_json_str: str) -> str:
    """Classify the OCR output into a document type. Returns the category label string."""
    ocr_excerpt = ocr_json_str[:12000]

    # Fast deterministic fallback based on OCR text itself.
    keyword_guess = _keyword_match_label(ocr_excerpt)

    try:
        completion = client.chat.completions.create(
            model=DEPLOYMENT,
            messages=[
                {"role": "system", "content": CLASSIFIER_PROMPT},
                {
                    "role": "user",
                    "content": (
                        "Classify this document. Return ONLY the category label.\n\n"
                        "OCR OUTPUT:\n" + ocr_json_str[:8000]  # Truncate to save tokens — headers are enough
                    ),
                },
            ],
            temperature=1.0,
            max_tokens=20,
        )
        raw = completion.choices[0].message.content or ""
        normalized = _normalize_label(raw)
        if normalized != "unknown":
            return normalized

        # Fallback: if model returns extra text, use OCR keyword heuristic.
        return keyword_guess
    except Exception:
        # Fallback even when LLM classification fails (auth/content filter/transient errors).
        return keyword_guess
