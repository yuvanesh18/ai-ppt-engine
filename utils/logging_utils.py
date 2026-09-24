"""
utils/logging_utils.py — Structured logging configuration.

Provides a consistent logger factory for all modules.
Log level is controlled by the LOG_LEVEL environment variable.
"""

from __future__ import annotations

import json
import logging
import os
import sys
import threading
from datetime import datetime, timezone
from typing import Any, Optional

_usage_logger = logging.getLogger("llm.usage")
_usage_lock = threading.Lock()


def configure_logging() -> None:
    """Configure root logger. Call once at application startup."""
    level_name = os.getenv("LOG_LEVEL", "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)

    logging.basicConfig(
        level=level,
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        stream=sys.stderr,
        force=True,
    )


def get_logger(name: str) -> logging.Logger:
    """Return a module-level logger."""
    return logging.getLogger(name)


def log_llm_usage(
    provider: str,
    key_number: Optional[int],
    key_label: str,
    model: str,
    prompt_tokens: Optional[int] = None,
    completion_tokens: Optional[int] = None,
    total_tokens: Optional[int] = None,
    **extra: Any,
) -> None:
    """
    Log a single LLM call's token usage and which API key served it.

    Emits an INFO line via the "llm.usage" logger and appends a JSON line to
    logs/llm/usage.jsonl for later auditing (e.g. spotting which key/project
    is burning through its token quota).
    """
    record = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "provider": provider,
        "key_number": key_number,
        "key": key_label,
        "model": model,
        "prompt_tokens": prompt_tokens,
        "completion_tokens": completion_tokens,
        "total_tokens": total_tokens,
        **extra,
    }

    _usage_logger.info(
        "LLM call — key=#%s (%s) provider=%s model=%s prompt_tokens=%s completion_tokens=%s total_tokens=%s",
        key_number, key_label, provider, model, prompt_tokens, completion_tokens, total_tokens,
    )

    try:
        import config  # local import to avoid a hard dependency for callers that don't need it

        usage_path = config.LOGS_DIR / "usage.jsonl"
        with _usage_lock:
            with open(usage_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(record) + "\n")
    except Exception:
        _usage_logger.debug("Failed to persist LLM usage record to usage.jsonl", exc_info=True)
