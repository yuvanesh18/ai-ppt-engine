"""
core/content_analyzer.py — Stage 1: LLM-based content analysis.

Takes raw source text + user preferences and returns a ContentAnalysis
object describing the key concepts, statistics, processes, etc. found
in the source — without generating any slide content yet.

This stage helps Stage 2 (presentation planning) make better layout
selection decisions.
"""

from __future__ import annotations

import json
from typing import Optional

from llm.groq_client import GroqClient, JSONParseError
from llm.prompts import CONTENT_ANALYSIS_PROMPT, SYSTEM_ROLE_ANALYST
from llm.schemas import ContentAnalysis, normalize_content_analysis_lists
from utils.logging_utils import get_logger
from utils.text_utils import truncate_text

logger = get_logger(__name__)

# Default values used when user leaves fields empty
_DEFAULTS = {
    "audience": "General",
    "style": "Professional",
    "language": "English",
    "additional_instructions": "None",
}


class ContentAnalysisError(Exception):
    """Raised when content analysis fails."""


def analyze_content(
    source_text: str,
    client: GroqClient,
    audience: str = "General",
    style: str = "Professional",
    slide_count: Optional[int] = None,
    language: str = "English",
    additional_instructions: str = "",
    max_source_chars: int = 14000,
) -> ContentAnalysis:
    """
    Stage 1: Analyze source content using LLM.

    Args:
        source_text:             The raw extracted or user-provided text.
        client:                  Initialized GroqClient.
        audience:                Target audience (e.g., "Business").
        style:                   Presentation style (e.g., "Professional").
        slide_count:             Requested number of slides (None = auto).
        language:                Output language.
        additional_instructions: Extra user instructions.
        max_source_chars:        Truncate source before sending to LLM.

    Returns:
        ContentAnalysis pydantic model.

    Raises:
        ContentAnalysisError: On LLM failure or validation error.
    """
    if not source_text or source_text.strip() == "":
        raise ContentAnalysisError("Source text is empty — nothing to analyze.")

    # Truncate to stay within context window
    truncated = truncate_text(source_text, max_source_chars)
    if len(truncated) < len(source_text):
        logger.info(
            "Source text truncated from %d to %d chars for LLM",
            len(source_text), len(truncated)
        )

    # Determine suggested slide count
    if slide_count is None:
        # Rough heuristic: 1 slide per ~300 chars of content, clamped 5–20
        char_count = len(truncated)
        suggested = max(5, min(20, char_count // 300))
    else:
        suggested = slide_count

    prompt = CONTENT_ANALYSIS_PROMPT.format(
        source_content=truncated,
        audience=audience or _DEFAULTS["audience"],
        style=style or _DEFAULTS["style"],
        slide_count=slide_count or "Automatic",
        language=language or _DEFAULTS["language"],
        additional_instructions=additional_instructions or _DEFAULTS["additional_instructions"],
        suggested_slides=suggested,
    )

    messages = [
        {"role": "system", "content": SYSTEM_ROLE_ANALYST},
        {"role": "user", "content": prompt},
    ]

    logger.info("Running content analysis (Stage 1)...")

    try:
        raw_data = client.chat_complete_json(
            messages=messages,
            temperature=0.2,  # Low temp for factual extraction
            max_tokens=2048,
        )
    except JSONParseError as e:
        raise ContentAnalysisError(f"LLM returned invalid JSON during analysis: {e}") from e
    except Exception as e:
        raise ContentAnalysisError(f"Content analysis LLM call failed: {e}") from e

    raw_data = normalize_content_analysis_lists(raw_data)

    try:
        analysis = ContentAnalysis(**raw_data)
        logger.info(
            "Content analysis complete: topic='%s', detected_types=%s",
            analysis.main_topic,
            analysis.content_types_detected,
        )
        return analysis

    except Exception as e:
        logger.error("ContentAnalysis validation failed: %s. Raw data: %s", e, str(raw_data)[:500])
        # Build a minimal analysis from whatever we got — never let a second
        # bad field crash this fallback path too.
        try:
            return ContentAnalysis(
                main_topic=raw_data.get("main_topic", "Unknown Topic"),
                key_concepts=raw_data.get("key_concepts", []),
                sections=raw_data.get("sections", []),
                statistics=raw_data.get("statistics", []),
                processes=raw_data.get("processes", []),
                comparisons=raw_data.get("comparisons", []),
                timelines=raw_data.get("timelines", []),
                conclusions=raw_data.get("conclusions", []),
                recommendations=raw_data.get("recommendations", []),
                suggested_slide_count=suggested,
                content_types_detected=raw_data.get("content_types_detected", ["TITLE_AND_CONTENT"]),
                summary=raw_data.get("summary", ""),
            )
        except Exception as e2:
            logger.error("ContentAnalysis repair also failed: %s — falling back to minimal stub", e2)
            return ContentAnalysis(
                main_topic=str(raw_data.get("main_topic", "Unknown Topic")),
                suggested_slide_count=suggested,
                content_types_detected=["TITLE_AND_CONTENT"],
                summary=str(raw_data.get("summary", "")),
            )


# ---------------------------------------------------------------------------
# Semantic pipeline entry point (Stage 1a for document uploads)
# ---------------------------------------------------------------------------

class SemanticAnalysisError(Exception):
    """Raised when the full semantic pipeline fails."""


def analyze_content_semantic(
    root_section,           # DocumentSection
    plain_text: str,
    plain_text_filename: str,  # used for document_title fallback
    key_manager,            # KeyManager
    logs_dir,               # Path
    run_id: str,
    audience: str = "General",
    style: str = "Professional",
    slide_count: Optional[int] = None,
    language: str = "English",
    additional_instructions: str = "",
    max_chunk_chars: int = 8000,
    min_section_chars: int = 300,
    progress_callback=None,  # Callable[[str, int, int], None]
) -> tuple:
    """
    Full semantic pipeline for structured document analysis.

    Pipeline:
      1. If root_section has no subsections → auto-infer structure via LLM
      2. Build semantic chunks from the section tree
      3. Save document_structure.json
      4. Analyze each chunk (rolling context, round-robin API keys)
      5. Consolidate all chunk analyses → ContentAnalysis
      6. Return (ContentAnalysis, AnalysisCache, chunks)

    Args:
        root_section:     DocumentSection tree from parse_document_with_structure().
        plain_text:       Full document text (for auto-structure fallback).
        plain_text_filename: Original filename (for document title).
        key_manager:      KeyManager with one or more API keys.
        logs_dir:         Where logs/llm/<run_id>/ files are saved.
        run_id:           Unique run identifier.
        audience/style/slide_count/language/additional_instructions: User prefs.
        max_chunk_chars:  Safety cap for single chunk size.
        min_section_chars: Merge threshold for tiny sections.
        progress_callback: Called with (heading, current, total) per chunk.

    Returns:
        (ContentAnalysis, AnalysisCache, List[SemanticChunk])
    """
    from core.document_structure import (
        build_semantic_chunks,
        auto_structure_from_plain_text,
        count_sections,
        render_structure_text,
    )
    from core.chunk_analyzer import (
        analyze_all_chunks,
        save_document_structure_log,
    )
    from core.consolidator import consolidate

    # Determine document title
    from pathlib import Path as _Path
    doc_title = _Path(plain_text_filename).stem.replace("_", " ").replace("-", " ").title()

    # ── Step 1: Inspect structure and select chunking strategy ───────────
    h1_count, h2_count = count_sections(root_section)
    detection_method = getattr(root_section, "_detection_method", "unknown")

    # Detection methods that provide reliable heading boundaries
    _RELIABLE_HEADING_METHODS = {"docx_styles", "txt_markdown", "txt_numbered"}

    use_dynamic_chunker = detection_method not in _RELIABLE_HEADING_METHODS

    logger.info(
        "Chunking strategy: detection_method='%s', h1=%d, h2+=%d → %s",
        detection_method,
        h1_count,
        h2_count,
        "AI-dynamic" if use_dynamic_chunker else "structural",
    )

    # ── Step 2: Build semantic chunks ─────────────────────────────────────
    if use_dynamic_chunker:
        # Documents with no reliable heading structure → AI-driven dynamic chunking
        from core.dynamic_chunker import run_dynamic_chunking
        chunks = run_dynamic_chunking(
            root_section=root_section,
            plain_text=plain_text,
            document_title=doc_title,
            key_manager=key_manager,
            logs_dir=logs_dir,
            run_id=run_id,
            max_chunk_chars=max_chunk_chars,
            min_chunk_chars=min_section_chars,
            progress_callback=progress_callback,
        )
        # Update detection_method so logs reflect the dynamic path
        detection_method = "ai_dynamic"

    else:
        # Documents with reliable heading structure → existing structural chunker
        if h1_count == 0:
            # No headings despite being in a 'reliable' format — run auto-structure
            logger.info("No headings in %s format — invoking LLM auto-structure", detection_method)
            inference_client = key_manager.get_client(chunk_index=0)
            root_section, detection_method = auto_structure_from_plain_text(
                text=plain_text,
                client=inference_client,
                document_title=doc_title,
            )
            h1_count, h2_count = count_sections(root_section)
            logger.info("Auto-structure: %d H1, %d H2+ sections", h1_count, h2_count)

        chunks = build_semantic_chunks(
            root=root_section,
            max_chunk_chars=max_chunk_chars,
            min_section_chars=min_section_chars,
        )

    if not chunks:
        logger.warning("No chunks produced — falling back to single-pass analysis")
        # Last resort: analyze the whole text in one shot
        single_client = key_manager.get_client(chunk_index=0)
        analysis = analyze_content(
            source_text=plain_text,
            client=single_client,
            audience=audience,
            style=style,
            slide_count=slide_count,
            language=language,
            additional_instructions=additional_instructions,
        )
        return analysis, None, []

    # ── Step 3: Save document structure log ───────────────────────────────
    save_document_structure_log(
        root=root_section,
        chunks=chunks,
        detection_method=detection_method,
        document_title=doc_title,
        logs_dir=logs_dir,
        run_id=run_id,
    )

    # ── Step 4: Per-chunk analysis (rolling context + multi-key) ─────────
    analyses, cache = analyze_all_chunks(
        chunks=chunks,
        key_manager=key_manager,
        document_title=doc_title,
        logs_dir=logs_dir,
        run_id=run_id,
        max_tokens_per_chunk=2048,
        progress_callback=progress_callback,
    )

    # ── Step 5: Consolidation ─────────────────────────────────────────────
    content_analysis = consolidate(
        analyses=analyses,
        cache=cache,
        root=root_section,
        document_title=doc_title,
        key_manager=key_manager,
        logs_dir=logs_dir,
        run_id=run_id,
        audience=audience,
        style=style,
        slide_count=slide_count,
        language=language,
        additional_instructions=additional_instructions,
    )

    return content_analysis, cache, chunks
