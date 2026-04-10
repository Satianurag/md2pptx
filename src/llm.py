from __future__ import annotations
import os
import time
import threading
import logging
from typing import TypeVar, Type
from collections import deque

from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.messages import HumanMessage, SystemMessage
from pydantic import BaseModel
from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger(__name__)


def _extract_text(content) -> str:
    """Extract plain text from LLM response content (may be str or list of dicts)."""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for item in content:
            if isinstance(item, dict) and "text" in item:
                parts.append(item["text"])
            elif isinstance(item, str):
                parts.append(item)
        return "\n".join(parts)
    return str(content)

T = TypeVar("T", bound=BaseModel)

# ── Rate limit constants ──
MAX_CALLS_PER_MINUTE = 15
MAX_TOKENS_PER_MINUTE = 250_000
MAX_RETRIES = 3
RETRY_BASE_DELAY = 4.0  # seconds


class RateLimiter:
    """Thread-safe sliding-window rate limiter for API calls and tokens."""

    def __init__(self, max_calls: int = MAX_CALLS_PER_MINUTE, max_tokens: int = MAX_TOKENS_PER_MINUTE):
        self._max_calls = max_calls
        self._max_tokens = max_tokens
        self._call_timestamps: deque[float] = deque()
        self._token_log: deque[tuple[float, int]] = deque()  # (timestamp, token_count)
        self._lock = threading.Lock()

    def _purge_old(self, now: float) -> None:
        cutoff = now - 60.0
        while self._call_timestamps and self._call_timestamps[0] < cutoff:
            self._call_timestamps.popleft()
        while self._token_log and self._token_log[0][0] < cutoff:
            self._token_log.popleft()

    def _tokens_used(self) -> int:
        return sum(t for _, t in self._token_log)

    def wait_if_needed(self, estimated_tokens: int = 0) -> None:
        """Block until the next call is safe under rate limits."""
        while True:
            with self._lock:
                now = time.time()
                self._purge_old(now)

                calls_ok = len(self._call_timestamps) < self._max_calls
                tokens_ok = (self._tokens_used() + estimated_tokens) <= self._max_tokens

                if calls_ok and tokens_ok:
                    self._call_timestamps.append(now)
                    return

                # Calculate wait time
                if not calls_ok and self._call_timestamps:
                    wait = self._call_timestamps[0] + 60.0 - now + 0.1
                elif not tokens_ok and self._token_log:
                    wait = self._token_log[0][0] + 60.0 - now + 0.1
                else:
                    wait = 1.0

            logger.info(f"Rate limit: waiting {wait:.1f}s")
            time.sleep(max(wait, 0.5))

    def record_tokens(self, token_count: int) -> None:
        with self._lock:
            self._token_log.append((time.time(), token_count))


# ── Singleton LLM instance ──
_llm_instance: ChatGoogleGenerativeAI | None = None
_rate_limiter: RateLimiter | None = None
_init_lock = threading.Lock()


def _get_api_key() -> str:
    key = os.getenv("GOOGLE_API_KEY", "")
    if not key:
        raise ValueError(
            "GOOGLE_API_KEY not set. Create a .env file with GOOGLE_API_KEY=your-key "
            "or set the environment variable."
        )
    return key


def get_llm() -> ChatGoogleGenerativeAI:
    """Return the shared ChatGoogleGenerativeAI singleton."""
    global _llm_instance
    if _llm_instance is None:
        with _init_lock:
            if _llm_instance is None:
                _llm_instance = ChatGoogleGenerativeAI(
                    model="gemini-3.1-flash-lite-preview",
                    google_api_key=_get_api_key(),
                    temperature=0.3,
                    max_output_tokens=16384,
                )
    return _llm_instance


def get_rate_limiter() -> RateLimiter:
    """Return the shared RateLimiter singleton."""
    global _rate_limiter
    if _rate_limiter is None:
        with _init_lock:
            if _rate_limiter is None:
                _rate_limiter = RateLimiter()
    return _rate_limiter


def invoke_llm(
    system_prompt: str,
    user_prompt: str,
    estimated_tokens: int = 5000,
) -> str:
    """Invoke the LLM with rate limiting and retry logic. Returns raw text."""
    llm = get_llm()
    limiter = get_rate_limiter()

    messages = [
        SystemMessage(content=system_prompt),
        HumanMessage(content=user_prompt),
    ]

    for attempt in range(MAX_RETRIES):
        try:
            limiter.wait_if_needed(estimated_tokens)
            response = llm.invoke(messages)

            # Record actual token usage if available
            usage = getattr(response, "usage_metadata", None)
            if usage:
                total = usage.get("total_tokens", estimated_tokens) if isinstance(usage, dict) else getattr(usage, "total_tokens", estimated_tokens)
                limiter.record_tokens(total)
            else:
                limiter.record_tokens(estimated_tokens)

            return _extract_text(response.content)

        except Exception as e:
            delay = RETRY_BASE_DELAY * (2 ** attempt)
            logger.warning(f"LLM call failed (attempt {attempt + 1}/{MAX_RETRIES}): {e}")
            if attempt < MAX_RETRIES - 1:
                logger.info(f"Retrying in {delay:.1f}s...")
                time.sleep(delay)
            else:
                raise RuntimeError(f"LLM call failed after {MAX_RETRIES} attempts: {e}") from e

    return ""  # unreachable


def invoke_llm_structured(
    system_prompt: str,
    user_prompt: str,
    output_schema: Type[T],
    estimated_tokens: int = 8000,
) -> T:
    """Invoke the LLM and parse the response into a Pydantic model using structured output."""
    llm = get_llm()
    limiter = get_rate_limiter()

    structured_llm = llm.with_structured_output(output_schema, method="json_schema")

    messages = [
        SystemMessage(content=system_prompt),
        HumanMessage(content=user_prompt),
    ]

    for attempt in range(MAX_RETRIES):
        try:
            limiter.wait_if_needed(estimated_tokens)
            result = structured_llm.invoke(messages)

            limiter.record_tokens(estimated_tokens)

            if not isinstance(result, output_schema):
                raise ValueError(f"Expected {output_schema.__name__}, got {type(result).__name__}")

            return result

        except Exception as e:
            delay = RETRY_BASE_DELAY * (2 ** attempt)
            logger.warning(f"Structured LLM call failed (attempt {attempt + 1}/{MAX_RETRIES}): {e}")
            if attempt < MAX_RETRIES - 1:
                logger.info(f"Retrying in {delay:.1f}s...")
                time.sleep(delay)
            else:
                raise RuntimeError(
                    f"Structured LLM call failed after {MAX_RETRIES} attempts: {e}"
                ) from e

    raise RuntimeError("Unreachable")
