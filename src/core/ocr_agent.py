import base64
import argparse
import json
import mimetypes
import os
import io
from pathlib import Path
from datetime import datetime, timezone
from collections import Counter

from dotenv import load_dotenv
from openai import AzureOpenAI, OpenAI

load_dotenv(Path(__file__).resolve().parents[2] / ".env")
from PIL import Image, ImageEnhance, ImageFilter

def _get_config_value(name: str) -> str | None:
  value = os.getenv(name)
  if value:
    return value

  try:
    import streamlit as st

    secret_value = st.secrets.get(name)
    if secret_value:
      return str(secret_value)
  except Exception:
    pass

  return None


def _get_required_env(name: str) -> str:
  value = _get_config_value(name)
  if not value:
    raise RuntimeError(
      f"Missing required config value: {name}. Set it as an environment variable or Streamlit secret."
    )
  return value


def _extract_usage_dict(completion: object) -> dict | None:
  usage = getattr(completion, "usage", None)
  if usage is None:
    return None

  if isinstance(usage, dict):
    prompt_tokens = usage.get("prompt_tokens")
    completion_tokens = usage.get("completion_tokens")
    total_tokens = usage.get("total_tokens")
  else:
    prompt_tokens = getattr(usage, "prompt_tokens", None)
    completion_tokens = getattr(usage, "completion_tokens", None)
    total_tokens = getattr(usage, "total_tokens", None)

  if prompt_tokens is None and completion_tokens is None and total_tokens is None:
    return None

  return {
    "prompt_tokens": prompt_tokens,
    "completion_tokens": completion_tokens,
    "total_tokens": total_tokens,
  }


def _append_token_usage_log(entry: dict) -> None:
  log_path = Path(__file__).resolve().parents[1] / "output" / "ocr" / "token_usage_log.jsonl"
  log_path.parent.mkdir(parents=True, exist_ok=True)
  with open(log_path, "a", encoding="utf-8") as f:
    f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def _log_token_usage(completion: object, request_mode: str, file_names: list[str]) -> None:
  usage = _extract_usage_dict(completion)
  if not usage:
    return

  entry = {
    "timestamp_utc": datetime.now(timezone.utc).isoformat(),
    "request_mode": request_mode,
    "model": deployment,
    "file_count": len(file_names),
    "file_names": file_names,
    **usage,
  }
  _append_token_usage_log(entry)


endpoint = _get_required_env("AZURE_OPENAI_ENDPOINT")
model_name = "gpt-5-chat"
deployment = _get_config_value("AZURE_OPENAI_DEPLOYMENT") or "gpt-5.2-chat"

subscription_key = _get_required_env("AZURE_OPENAI_API_KEY")
api_version = _get_config_value("AZURE_OPENAI_API_VERSION") or "2025-03-01-preview"

client = AzureOpenAI(
    api_version=api_version,
    azure_endpoint=endpoint,
    api_key=subscription_key,
)


def _preprocess_image(image_path: Path) -> bytes:
  """Preprocess image to improve OCR accuracy."""
  img = Image.open(image_path)

  if img.mode != "RGB":
    img = img.convert("RGB")

  min_dim = min(img.size)
  if min_dim < 2000:
    scale = 2000 / min_dim
    new_size = (int(img.width * scale), int(img.height * scale))
    img = img.resize(new_size, Image.LANCZOS)

  img = img.filter(ImageFilter.SHARPEN)

  enhancer = ImageEnhance.Contrast(img)
  img = enhancer.enhance(1.3)

  buf = io.BytesIO()
  img.save(buf, format="PNG", optimize=True)
  return buf.getvalue()


def _image_file_to_data_url(image_path: Path, preprocess: bool = True) -> str:
  """Convert image to base64 data URL, optionally with preprocessing."""
  if preprocess:
    img_bytes = _preprocess_image(image_path)
    mime_type = "image/png"
  else:
    img_bytes = image_path.read_bytes()
    mime_type, _ = mimetypes.guess_type(str(image_path))
    if not mime_type:
      mime_type = "application/octet-stream"
  b64 = base64.b64encode(img_bytes).decode("utf-8")
  return f"data:{mime_type};base64,{b64}"

SYSTEM_PROMPT = """
You are an OCR transcription engine. Your ONLY task is to read every visible character from the provided document image(s) and output a structured JSON transcription.

You are NOT an assistant, analyst, or calculator. You do NOT interpret, summarize, compute, or infer anything.

═══ TRANSCRIPTION RULES ═══

1. COPY EXACTLY what you see: every letter, digit, punctuation mark, symbol, and whitespace.
2. NEVER change, correct, round, recalculate, or reformat any value.
   - If the document says "30,752.88", output "30,752.88" — not "30752.88".
   - If a word is misspelled on the document, reproduce the misspelling.
3. PRESERVE the spatial/logical grouping of text:
   - Key-value pairs (e.g. "CI Number : 26001099") → keep label and value together.
   - Tables → preserve column headers and row alignment.
   - Addresses, notes, footers → keep as contiguous blocks.
4. READ EVERY PAGE. Do NOT skip repeated headers, footers, or boilerplate.
5. If multiple images are provided, treat them as pages of ONE document in the order given.

═══ DIGIT ACCURACY — CRITICAL ═══

Pay EXTREME attention to easily confused characters in reference numbers, account numbers, barcodes,
and other numeric fields. Common confusions to watch for:
- "8" vs "5" vs "6" vs "0" — zoom in mentally on each digit's shape
- "0" (zero) vs "O" (letter O) vs "D"
- "1" (one) vs "l" (lowercase L) vs "I" (uppercase i)
- "9" vs "0" — check whether the top is closed (9) or open (0)
- Do NOT insert or drop digits — count the digits carefully for long numbers
- If a reference number appears on multiple pages, cross-check that your reading is consistent

═══ CONFIDENCE TAGGING ═══

- Assign a confidence score (0.00–1.00) to each section.
- Set confidence = 1.00 ONLY when every character in that section is clearly legible.
- If a character or word is ambiguous, output your best reading and LOWER the confidence.
- If text is completely unreadable, output "[UNREADABLE]" with confidence 0.00.
- If content is visibly cut off at the page edge, output what is visible and append "[PARTIAL]".

═══ STRICT PROHIBITIONS ═══

- DO NOT add, remove, or alter any text that is not on the document.
- DO NOT calculate totals, subtotals, percentages, or derived values.
- DO NOT merge or reconcile data across pages or tables.
- DO NOT infer missing values or fill blanks.
- DO NOT rewrite table data into sentences or summaries.
- DO NOT add any explanation, commentary, or markdown.

═══ OUTPUT SCHEMA ═══

Return ONLY a single valid JSON object. No markdown fences. No text before or after.

{
  "pages": [
    {
      "page_number": <int>,
      "file_name": "<filename if provided>",
      "sections": [
        {
          "type": "<section_type>",
          "content": "<exact transcribed text>",
          "confidence": <float 0.00-1.00>
        }
      ]
    }
  ],
  "metadata": {
    "total_pages": <int>,
    "languages_detected": ["en"],
    "image_quality": "clear | noisy | blurry | low_resolution"
  }
}

Allowed section types:
- "header"       : document titles, company names, page labels (e.g. "Page 1 of 3")
- "address"      : address blocks (bill-to, ship-to, business unit)
- "key_value"    : label-value pairs (e.g. "PO Number : MY2501539")
- "table_header" : column header row of a table
- "table_row"    : one data row of a table, values separated by " | "
- "subtotal"     : subtotal / total / summary lines
- "paragraph"    : free-form text, notes, instructions
- "footer"       : page footers, disclaimers, correspondence addresses
- "signature"    : signature blocks, stamps, seals
- "empty"        : page has no extractable text (include reason in content)

═══ TABLE TRANSCRIPTION RULES ═══

- Output the column header row as one section with type "table_header", values separated by " | ".
- Output each data row as a separate section with type "table_row", values separated by " | ".
- Preserve the column order exactly as it appears on the document.
- If a cell is empty, output an empty string between the delimiters (e.g. "value1 |  | value3").
- DO NOT infer column meanings, merge rows, or reorder columns.

CRITICAL — MULTI-LINE PRODUCT DESCRIPTIONS:
- A single product row often has MULTIPLE lines of text in the description column:
    Line 1: Main product name (e.g. "WATSONS SIDE SEALED COTTON PUFFS")
    Line 2: Product variant/spec (e.g. "WATSONS SIDE SEALED COTTON PUFFS 189S SEA AW2022")
    Line 3: Campaign/collection name (e.g. "WATSONS OOB COTTON RELAUNCH")
- ALL of these lines belong to the SAME product row.
- Combine them into ONE table_row section, joining description sub-lines with \\n.
- The key indicator that lines belong to the same row: they share the SAME barcode, quantity, unit price, and amount on the right side.
- If description text appears on lines below the barcode/quantity/price row, and those lines have NO barcode, NO quantity, NO unit price, and NO amount of their own, they are sub-descriptions of the row above — merge them into that row.
- NEVER output a sub-description line as a separate table_row.
"""


def ocr_image_with_chat_model(image_path: Path, user_prompt: str) -> str:
  data_url = _image_file_to_data_url(image_path)

  try:
    completion = client.chat.completions.create(
      model=deployment,
      messages=[
        {"role": "system", "content": SYSTEM_PROMPT},
        {
          "role": "user",
          "content": [
            {"type": "text", "text": user_prompt},
            {"type": "image_url", "image_url": {"url": data_url, "detail": "high"}},
          ],
        },
      ],
    )
    _log_token_usage(
      completion=completion,
      request_mode="single_image",
      file_names=[image_path.name],
    )
    return completion.choices[0].message.content or ""
  except Exception as e:
    return json.dumps(
      {
        "error": "ocr_failed",
        "message": str(e),
      },
      ensure_ascii=False,
    )


def _extract_key_values_from_page(page_data: dict) -> dict[str, str]:
  kv = {}
  for section in page_data.get("sections", []):
    if section.get("type") == "key_value":
      for line in section["content"].split("\n"):
        if " : " in line:
          key, _, val = line.partition(" : ")
          kv[key.strip()] = val.strip()
  return kv


def _majority_vote(values: list[str]) -> tuple[str, float]:
  if not values:
    return "", 0.0
  counts = Counter(values)
  winner, count = counts.most_common(1)[0]
  return winner, count / len(values)


def _merge_consensus_kv(all_pass_kvs: list[dict[str, str]]) -> dict[str, tuple[str, float]]:
  all_keys = set()
  for kv in all_pass_kvs:
    all_keys.update(kv.keys())

  result = {}
  for key in all_keys:
    values = [kv[key] for kv in all_pass_kvs if key in kv]
    consensus_val, agreement = _majority_vote(values)
    result[key] = (consensus_val, agreement)
  return result


def _apply_consensus_to_page(page_data: dict, consensus_kv: dict[str, tuple[str, float]]) -> dict:
  import copy
  page = copy.deepcopy(page_data)

  for section in page.get("sections", []):
    if section.get("type") != "key_value":
      continue

    new_lines = []
    min_confidence = 1.0
    for line in section["content"].split("\n"):
      if " : " in line:
        key, _, _ = line.partition(" : ")
        key_s = key.strip()
        if key_s in consensus_kv:
          val, agreement = consensus_kv[key_s]
          new_lines.append(f"{key_s} : {val}")
          min_confidence = min(min_confidence, agreement)
        else:
          new_lines.append(line)
      else:
        new_lines.append(line)

    section["content"] = "\n".join(new_lines)
    section["confidence"] = round(min_confidence, 2)

  return page


def ocr_image_multipass(
  image_path: Path,
  user_prompt: str,
  num_passes: int = 3,
) -> str:
  data_url = _image_file_to_data_url(image_path)

  pass_results: list[str] = []
  for i in range(num_passes):
    try:
      completion = client.chat.completions.create(
        model=deployment,
        messages=[
          {"role": "system", "content": SYSTEM_PROMPT},
          {
            "role": "user",
            "content": [
              {"type": "text", "text": user_prompt},
              {"type": "image_url", "image_url": {"url": data_url, "detail": "high"}},
            ],
          },
        ],
      )
      _log_token_usage(
        completion=completion,
        request_mode=f"multipass_{i+1}_of_{num_passes}",
        file_names=[image_path.name],
      )
      pass_results.append(completion.choices[0].message.content or "")
    except Exception as e:
      print(f"  WARNING: Pass {i+1} failed: {e}")
      continue

  if not pass_results:
    return json.dumps({"error": "all_passes_failed"}, ensure_ascii=False)

  parsed_passes = []
  for raw in pass_results:
    try:
      parsed_passes.append(json.loads(raw))
    except Exception:
      continue

  if not parsed_passes:
    return pass_results[0]

  if len(parsed_passes) == 1:
    return json.dumps(parsed_passes[0], ensure_ascii=False)

  base = parsed_passes[0]

  for page_idx, page in enumerate(base.get("pages", [])):
    all_pass_kvs = []
    for parsed in parsed_passes:
      pages = parsed.get("pages", [])
      if page_idx < len(pages):
        all_pass_kvs.append(_extract_key_values_from_page(pages[page_idx]))

    if all_pass_kvs:
      consensus_kv = _merge_consensus_kv(all_pass_kvs)
      base["pages"][page_idx] = _apply_consensus_to_page(page, consensus_kv)

      for key, (val, agreement) in consensus_kv.items():
        if agreement < 1.0:
          all_vals = [kv.get(key, "<missing>") for kv in all_pass_kvs]
          print(f"  CONSENSUS [{agreement:.0%}] {key}: {val}  (all readings: {all_vals})")

  return json.dumps(base, ensure_ascii=False)


def ocr_images_with_chat_model(image_paths: list[Path], user_prompt: str) -> str:
  content: list[dict] = [
    {
      "type": "text",
      "text": (
        f"{user_prompt}\n\n"
        "You will receive multiple images. Treat them as pages of ONE single document in the order given. "
        "Include one entry per image in pages[], preserving the order. "
        "Set page_number starting from 1. Include the filename in a field named file_name. "
        "Do NOT skip any page, even if its content repeats a previous page."
      ),
    }
  ]

  for idx, image_path in enumerate(image_paths, start=1):
    content.append({"type": "text", "text": f"Image {idx} filename: {image_path.name}"})
    content.append({"type": "image_url", "image_url": {"url": _image_file_to_data_url(image_path), "detail": "high"}})

  try:
    completion = client.chat.completions.create(
      model=deployment,
      messages=[
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": content},
      ],
    )
    _log_token_usage(
      completion=completion,
      request_mode="batch",
      file_names=[p.name for p in image_paths],
    )
    return completion.choices[0].message.content or ""
  except Exception as e:
    return json.dumps(
      {
        "error": "ocr_failed",
        "message": str(e),
      },
      ensure_ascii=False,
    )


def _maybe_parse_json(text: str) -> object:
  try:
    return json.loads(text)
  except Exception:
    return text
