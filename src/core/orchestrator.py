"""
Orchestrator: classifies an OCR JSON file and routes to the correct extraction agent.
"""

import json
import sys
from pathlib import Path

from agents import call_extraction_agent, maybe_parse_json
from agents.classifier import classify_document
from agents import extraction_invoice
from agents import extraction_travel
from agents import extraction_rental
from agents import extraction_hotel
from agents import extraction_utility
from agents import extraction_soa
from agents import extraction_bank
from agents import srkk_vendor_invoice
from agents import srkk_purchase_order
from agents import srkk_microsoft_billing

# Registry: maps classifier label → (SYSTEM_PROMPT, USER_PROMPT)
AGENT_REGISTRY: dict[str, tuple[str, str]] = {
    "commercial_invoice":       (extraction_invoice.SYSTEM_PROMPT, extraction_invoice.USER_PROMPT),
    "credit_note":              (extraction_invoice.SYSTEM_PROMPT, extraction_invoice.USER_PROMPT),
    "travel":                   (extraction_travel.SYSTEM_PROMPT,  extraction_travel.USER_PROMPT),
    "rental":                   (extraction_rental.SYSTEM_PROMPT,  extraction_rental.USER_PROMPT),
    "hotel":                    (extraction_hotel.SYSTEM_PROMPT,   extraction_hotel.USER_PROMPT),
    "utility":                  (extraction_utility.SYSTEM_PROMPT, extraction_utility.USER_PROMPT),
    "soa":                      (extraction_soa.SYSTEM_PROMPT,     extraction_soa.USER_PROMPT),
    "bank_statement":           (extraction_bank.SYSTEM_PROMPT,    extraction_bank.USER_PROMPT),
    "srkk_vendor_invoice":      (srkk_vendor_invoice.SYSTEM_PROMPT, srkk_vendor_invoice.USER_PROMPT),
    "srkk_purchase_order":      (srkk_purchase_order.SYSTEM_PROMPT, srkk_purchase_order.USER_PROMPT),
    "srkk_microsoft_billing":   (srkk_microsoft_billing.SYSTEM_PROMPT, srkk_microsoft_billing.USER_PROMPT),
}

FALLBACK_SYSTEM_PROMPT = extraction_invoice.SYSTEM_PROMPT
FALLBACK_USER_PROMPT = extraction_invoice.USER_PROMPT


def run(ocr_json_str: str, forced_type: str | None = None) -> tuple[str, object]:
    """Classify and extract."""
    if forced_type:
        doc_type = forced_type
    else:
        doc_type = classify_document(ocr_json_str)

    if doc_type in AGENT_REGISTRY:
        system_prompt, user_prompt = AGENT_REGISTRY[doc_type]
    else:
        system_prompt, user_prompt = FALLBACK_SYSTEM_PROMPT, FALLBACK_USER_PROMPT

    raw_result = call_extraction_agent(system_prompt, user_prompt, ocr_json_str)
    parsed = maybe_parse_json(raw_result)

    if isinstance(parsed, dict) and "document_type" not in parsed:
        parsed["document_type"] = doc_type

    return doc_type, parsed
