from __future__ import annotations

import hashlib
import os
import time
from dataclasses import dataclass
from typing import Any, Dict, Optional

from dotenv import load_dotenv
import httpx

# Load .env file - dotenv searches from current dir up to project root automatically
load_dotenv()


@dataclass
class LLMResponse:
    model: str
    created_ts: float
    prompt_hash: str
    text: str
    raw: Dict[str, Any]


def _get_env(name: str, default: Optional[str] = None) -> str:
    val = os.getenv(name, default)
    if val is None:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return val


def sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def call_openrouter(
    *,
    model: str,
    prompt: str,
    temperature: float = 0.2,
    max_tokens: int = 1200,
    timeout_seconds: int = 60,
    max_retries: int = 3,
    retry_delay_base: float = 1.0,
) -> LLMResponse:
    """Call an arbitrary OpenRouter-compatible model and return structured output.
    
    Implements retry logic with exponential backoff for transient network errors.
    
    Environment variables:
    - OPENROUTER_API_KEY
    - OPENROUTER_BASE_URL (defaults to https://openrouter.ai/api/v1)
    
    Args:
        max_retries: Maximum number of retry attempts (default: 3)
        retry_delay_base: Base delay in seconds for exponential backoff (default: 1.0)
    """
    api_key = _get_env("OPENROUTER_API_KEY")
    base_url = os.getenv("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1")

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    body: Dict[str, Any] = {
        "model": model,
        "messages": [
            {"role": "user", "content": prompt},
        ],
        "temperature": temperature,
        "max_tokens": max_tokens,
    }

    created_ts = time.time()
    prompt_hash = sha256_text(prompt)

    last_error = None
    for attempt in range(max_retries):
        try:
            with httpx.Client(base_url=base_url, timeout=timeout_seconds) as client:
                resp = client.post("/chat/completions", headers=headers, json=body)
                resp.raise_for_status()
                data = resp.json()
                break  # Success, exit retry loop
        except (httpx.NetworkError, httpx.ReadTimeout, httpx.ConnectTimeout, httpx.RemoteProtocolError) as e:
            last_error = e
            if attempt < max_retries - 1:
                # Exponential backoff: delay = base * 2^attempt
                delay = retry_delay_base * (2 ** attempt)
                time.sleep(delay)
                continue
            else:
                # Final attempt failed, re-raise
                raise RuntimeError(f"OpenRouter API call failed after {max_retries} attempts: {e}") from e
        except httpx.HTTPStatusError as e:
            # Don't retry on client/server errors (4xx/5xx) unless it's a 5xx server error
            if e.response.status_code >= 500 and attempt < max_retries - 1:
                # Retry on 5xx errors
                delay = retry_delay_base * (2 ** attempt)
                time.sleep(delay)
                continue
            else:
                # Don't retry on 4xx errors or final attempt
                raise
    else:
        # This should never be reached, but handle just in case
        if last_error:
            raise RuntimeError(f"OpenRouter API call failed after {max_retries} attempts: {last_error}") from last_error
        raise RuntimeError("OpenRouter API call failed for unknown reason")

    # Try to be resilient to different providers' shapes
    text = ""
    try:
        choices = data.get("choices") or []
        if choices:
            msg = choices[0].get("message") or {}
            # Check for content first, then reasoning (for o1-style models)
            text = (msg.get("content") or "").strip()
            if not text:
                # Some models (like GPT-5/o1) use reasoning field
                text = (msg.get("reasoning") or "").strip()
    except Exception:
        text = ""

    return LLMResponse(
        model=model,
        created_ts=created_ts,
        prompt_hash=prompt_hash,
        text=text,
        raw=data,
    )




