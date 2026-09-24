"""
llm/groq_client.py — Groq API wrapper with JSON extraction and retry logic.

Responsibilities:
- Wrap the Groq Python SDK for chat completions.
- Extract valid JSON from raw LLM responses (handles markdown fences, preamble).
- Retry on JSON parse failure with a correction prompt.
- Translate Groq SDK exceptions into application-level exceptions.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional

from llm.json_utils import JSONParseError, parse_json_response, retry_json_completion
from utils.logging_utils import get_logger, log_llm_usage

logger = get_logger(__name__)

# ---------------------------------------------------------------------------
# Custom exceptions
# ---------------------------------------------------------------------------

class GroqClientError(Exception):
    """Base exception for all Groq client errors."""


class GroqAuthError(GroqClientError):
    """Invalid or missing API key."""


class GroqRateLimitError(GroqClientError):
    """Rate limit exceeded."""


class GroqAPIError(GroqClientError):
    """Generic API error (network, server, etc.)."""


# JSONParseError, parse_json_response re-exported from llm.json_utils for
# backward compatibility with existing `from llm.groq_client import JSONParseError` call sites.


# ---------------------------------------------------------------------------
# Groq client
# ---------------------------------------------------------------------------

class GroqClient:
    """
    Thin wrapper around the Groq Python SDK.

    Usage:
        client = GroqClient(api_key="gsk_...", model="openai/gpt-oss-120b")
        raw = client.chat_complete(messages=[...])
        data = client.chat_complete_json(messages=[...])
    """

    def __init__(
        self,
        api_key: str,
        model: str = "openai/gpt-oss-120b",
        temperature: float = 0.3,
        max_tokens: int = 4096,
        max_retries: int = 3,
        key_number: Optional[int] = None,
    ) -> None:
        if not api_key or api_key.strip() == "":
            raise GroqAuthError(
                "Groq API key is missing. Set GROQ_API_KEY in your .env file "
                "or enter it in the sidebar."
            )

        try:
            from groq import Groq, APIStatusError, RateLimitError, AuthenticationError
            self._Groq = Groq
            self._APIStatusError = APIStatusError
            self._RateLimitError = RateLimitError
            self._AuthenticationError = AuthenticationError
            self._client = Groq(api_key=api_key.strip())
        except ImportError as e:
            raise GroqClientError(
                "groq package is not installed. Run: pip install groq"
            ) from e

        self.model = model
        self.temperature = temperature
        self.max_tokens = max_tokens
        self.max_retries = max_retries
        self.key_number = key_number
        self._key_label = f"...{api_key.strip()[-4:]}" if len(api_key.strip()) > 4 else "..."

    def chat_complete(
        self,
        messages: List[Dict[str, str]],
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
    ) -> str:
        """
        Send a chat completion request to Groq.
        Returns the raw text response.
        """
        _temp = temperature if temperature is not None else self.temperature
        _max_tok = max_tokens if max_tokens is not None else self.max_tokens

        logger.debug(
            "Groq request — model=%s temp=%.2f max_tokens=%d messages=%d",
            self.model,
            _temp,
            _max_tok,
            len(messages),
        )

        extra_kwargs = {}
        if "gpt-oss" in self.model:
            extra_kwargs["extra_body"] = {"reasoning_effort": "low"}

        try:
            response = self._client.chat.completions.create(
                model=self.model,
                messages=messages,  # type: ignore[arg-type]
                temperature=_temp,
                max_tokens=_max_tok,
                **extra_kwargs,
            )
            content = response.choices[0].message.content or ""
            usage = getattr(response, "usage", None)
            log_llm_usage(
                provider="groq",
                key_number=self.key_number,
                key_label=self._key_label,
                model=self.model,
                prompt_tokens=getattr(usage, "prompt_tokens", None),
                completion_tokens=getattr(usage, "completion_tokens", None),
                total_tokens=getattr(usage, "total_tokens", None),
            )
            logger.debug("Groq response length: %d chars", len(content))
            return content

        except self._AuthenticationError as e:
            raise GroqAuthError(
                "Authentication failed. Please check your Groq API key."
            ) from e

        except self._RateLimitError as e:
            raise GroqRateLimitError(
                "Groq rate limit exceeded. Please wait a moment and try again."
            ) from e

        except self._APIStatusError as e:
            raise GroqAPIError(
                f"Groq API error (status {e.status_code}): {e.message}"
            ) from e

        except Exception as e:
            raise GroqAPIError(f"Unexpected error calling Groq: {e}") from e

    def chat_complete_json(
        self,
        messages: List[Dict[str, str]],
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
    ) -> Dict[str, Any]:
        """
        Send a chat completion request and parse the response as JSON.

        On JSON parse failure, sends a correction prompt and retries
        up to self.max_retries times.
        """
        return retry_json_completion(
            self.chat_complete,
            messages,
            self.max_retries,
            temperature=temperature,
            max_tokens=max_tokens,
        )
