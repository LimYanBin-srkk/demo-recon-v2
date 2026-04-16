"""Shared Azure OpenAI client and helpers used by all agents."""

import json
import os
from pathlib import Path
from dotenv import load_dotenv
from openai import AzureOpenAI

# Load .env from project root (two levels up from src/agents/)
load_dotenv(Path(__file__).resolve().parents[2] / ".env")

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


ENDPOINT = _get_required_env("AZURE_OPENAI_ENDPOINT")
DEPLOYMENT = _get_config_value("AZURE_OPENAI_DEPLOYMENT") or "gpt-5.2-chat"
SUBSCRIPTION_KEY = _get_required_env("AZURE_OPENAI_API_KEY")
API_VERSION = _get_config_value("AZURE_OPENAI_API_VERSION") or "2025-03-01-preview"

client = AzureOpenAI(
    api_version=API_VERSION,
    azure_endpoint=ENDPOINT,
    api_key=SUBSCRIPTION_KEY,
)


def call_extraction_agent(system_prompt: str, user_prompt: str, ocr_json_str: str) -> str:
    """Send OCR JSON to an extraction agent and return raw response text."""
    try:
        completion = client.chat.completions.create(
            model=DEPLOYMENT,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt + ocr_json_str},
            ],
        )
        return completion.choices[0].message.content or ""
    except Exception as e:
        return json.dumps(
            {"error": "extraction_failed", "message": str(e)},
            ensure_ascii=False,
        )


def maybe_parse_json(text: str) -> object:
    try:
        return json.loads(text)
    except Exception:
        return text
