"""
core/consolidator.py — Consolidation: merge N chunk analyses → ContentAnalysis.

After all semantic chunks are analyzed, this module:
  1. Formats a hierarchy-aware consolidation prompt
  2. Makes a single LLM call to synthesize → ContentAnalysis (same schema as before)
  3. Saves consolidation.json to the run log directory

The consolidation prompt preserves the document hierarchy:
  Introduction
    Analysis...
  Key Components
    Procurement
      Analysis...
    Manufacturing
      Analysis...
  Conclusion
    Analysis...

This ensures the presentation planner receives rich, structured context.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import List, Optional

from llm.schemas import ChunkAnalysis, ContentAnalysis, normalize_content_analysis_lists
from llm.prompts import CONSOLIDATION_PROMPT, SYSTEM_ROLE_CONSOLIDATOR
from llm.key_manager import KeyManager
from core.document_structure import DocumentSection, render_structure_text
from core.chunk_analyzer import AnalysisCache, _save_json
from utils.logging_utils import get_logger

logger = get_logger(__name__)


class ConsolidationError(Exception):
    """Raised when consolidation LLM call fails unrecoverably."""


# ---------------------------------------------------------------------------
# Hierarchy-aware consolidation prompt builder
# ---------------------------------------------------------------------------

def _build_section_analyses_text(analyses: List[ChunkAnalysis]) -> str:
    """
    Format chunk analyses with their hierarchy preserved for the prompt.
    Each section is indented by its level for visual clarity.
    """
    lines: List[str] = []

    for analysis in analyses:
        indent = "  " * (analysis.level - 1) if analysis.level > 0 else ""
        path_str = " → ".join(analysis.section_path) if analysis.section_path else analysis.heading

        lines.append(f"\n{'─' * 60}")
        lines.append(f"{indent}SECTION: {path_str}")
        lines.append(f"{indent}Chunk ID: {analysis.chunk_id} | Level: {analysis.level} | Importance: {analysis.importance}")
        lines.append(f"{indent}Summary: {analysis.summary}")

        if analysis.key_points:
            lines.append(f"{indent}Key Points: {'; '.join(analysis.key_points[:5])}")
        if analysis.statistics:
            lines.append(f"{indent}Statistics: {'; '.join(analysis.statistics[:5])}")
        if analysis.processes:
            lines.append(f"{indent}Processes: {'; '.join(analysis.processes[:3])}")
        if analysis.comparisons:
            lines.append(f"{indent}Comparisons: {'; '.join(analysis.comparisons[:3])}")
        if analysis.dates:
            lines.append(f"{indent}Dates/Milestones: {'; '.join(analysis.dates[:4])}")
        if analysis.potential_visuals:
            lines.append(f"{indent}Suggested Visuals: {'; '.join(analysis.potential_visuals[:3])}")
        if analysis.content_types:
            lines.append(f"{indent}Content Types: {', '.join(analysis.content_types)}")

    return "\n".join(lines)


def _build_structure_outline(root: DocumentSection) -> str:
    """Simple outline string of the document structure."""
    lines: List[str] = []

    def _walk(sec: DocumentSection, depth: int = 0) -> None:
        if sec.level == 0:
            for sub in sec.subsections:
                _walk(sub, 0)
            return
        prefix = "  " * (sec.level - 1)
        lines.append(f"{prefix}{'H' + str(sec.level)}: {sec.heading}")
        for sub in sec.subsections:
            _walk(sub, depth + 1)

    _walk(root)
    return "\n".join(lines) if lines else "(flat document)"


# ---------------------------------------------------------------------------
# Main consolidation function
# ---------------------------------------------------------------------------

def consolidate(
    analyses: List[ChunkAnalysis],
    cache: AnalysisCache,
    root: DocumentSection,
    document_title: str,
    key_manager: KeyManager,
    logs_dir: Path,
    run_id: str,
    audience: str = "General",
    style: str = "Professional",
    slide_count: Optional[int] = None,
    language: str = "English",
    additional_instructions: str = "",
) -> ContentAnalysis:
    """
    Synthesize all chunk analyses into a single ContentAnalysis object.

    Args:
        analyses:       Ordered list of ChunkAnalysis from chunk_analyzer.
        cache:          In-memory cache (used for completeness check).
        root:           Original DocumentSection tree.
        document_title: Document title.
        key_manager:    Provides the LLM client (uses next_key()).
        logs_dir:       Where to save consolidation.json.
        run_id:         Unique run identifier.
        audience/style/slide_count/language/additional_instructions: User prefs.

    Returns:
        ContentAnalysis — same Pydantic model expected by presentation_planner.
    """
    from llm.groq_client import JSONParseError

    if not analyses:
        raise ConsolidationError("No chunk analyses to consolidate.")

    # Auto-determine slide count if not specified
    if slide_count is None:
        total_chars = sum(len(a.summary) + len(" ".join(a.key_points)) for a in analyses)
        suggested = max(5, min(20, len(analyses) * 2))
    else:
        suggested = slide_count

    # Build the prompt
    structure_outline = _build_structure_outline(root)
    section_analyses_text = _build_section_analyses_text(analyses)

    prompt = CONSOLIDATION_PROMPT.format(
        document_title=document_title,
        structure_outline=structure_outline,
        section_analyses=section_analyses_text,
        audience=audience or "General",
        style=style or "Professional",
        slide_count=slide_count or "Automatic",
        language=language or "English",
        additional_instructions=additional_instructions or "None",
        suggested_slides=suggested,
    )

    messages = [
        {"role": "system", "content": SYSTEM_ROLE_CONSOLIDATOR},
        {"role": "user", "content": prompt},
    ]

    logger.info("Running consolidation LLM call (%d chunk analyses)", len(analyses))

    # Use a fresh client key for consolidation
    client = key_manager.get_client(max_tokens=3000)

    try:
        raw_data = client.chat_complete_json(
            messages=messages,
            temperature=0.2,
            max_tokens=3000,
        )
    except JSONParseError as e:
        raise ConsolidationError(f"Consolidation LLM returned invalid JSON: {e}") from e
    except Exception as e:
        raise ConsolidationError(f"Consolidation LLM call failed: {e}") from e

    # Parse into ContentAnalysis
    raw_data = normalize_content_analysis_lists(raw_data)
    try:
        content_analysis = ContentAnalysis(**raw_data)
    except Exception as e:
        logger.warning("ContentAnalysis validation failed in consolidation: %s — building from raw", e)
        try:
            content_analysis = _build_fallback_analysis(raw_data, analyses, document_title, suggested)
        except Exception as e2:
            logger.error("Fallback analysis build also failed: %s — using minimal stub", e2)
            content_analysis = ContentAnalysis(
                main_topic=str(raw_data.get("main_topic", document_title)),
                suggested_slide_count=suggested,
                content_types_detected=["TITLE_AND_CONTENT"],
                summary=str(raw_data.get("summary", f"Analysis of {document_title}")),
            )

    # Save consolidation log
    consolidation_log = {
        "run_id": run_id,
        "document_title": document_title,
        "chunks_consolidated": len(analyses),
        "detection_method": "semantic",
        "output": content_analysis.model_dump(),
        "raw_llm_response": raw_data,
    }
    _save_json(logs_dir / run_id / "consolidation.json", consolidation_log)

    logger.info(
        "Consolidation complete: topic='%s', %d sections, %d slide(s) suggested",
        content_analysis.main_topic,
        len(content_analysis.sections),
        content_analysis.suggested_slide_count,
    )

    return content_analysis


# ---------------------------------------------------------------------------
# Fallback analysis builder
# ---------------------------------------------------------------------------

def _build_fallback_analysis(
    raw_data: dict,
    analyses: List[ChunkAnalysis],
    document_title: str,
    suggested_slides: int,
) -> ContentAnalysis:
    """Build ContentAnalysis from raw LLM data + chunk analyses when validation fails."""
    # Aggregate from chunks
    all_stats = []
    all_processes = []
    all_comparisons = []
    all_dates = []
    all_concepts = []
    all_conclusions = []
    content_types_seen: set = set()

    for a in analyses:
        all_stats.extend(a.statistics[:2])
        all_processes.extend(a.processes[:2])
        all_comparisons.extend(a.comparisons[:1])
        all_dates.extend(a.dates[:2])
        all_concepts.extend(a.key_points[:2])
        content_types_seen.update(a.content_types)

    sections = [a.heading for a in analyses if a.level <= 1]
    summary_parts = [a.summary for a in analyses if a.summary and a.importance == "high"]
    summary = " ".join(summary_parts[:3]) if summary_parts else f"Analysis of {document_title}"

    return ContentAnalysis(
        main_topic=raw_data.get("main_topic", document_title),
        key_concepts=raw_data.get("key_concepts", all_concepts[:8]) or all_concepts[:8],
        sections=raw_data.get("sections", sections),
        statistics=raw_data.get("statistics", all_stats[:6]) or all_stats[:6],
        processes=raw_data.get("processes", all_processes[:4]) or all_processes[:4],
        comparisons=raw_data.get("comparisons", all_comparisons[:3]) or all_comparisons[:3],
        timelines=raw_data.get("timelines", all_dates[:4]) or all_dates[:4],
        conclusions=raw_data.get("conclusions", []),
        recommendations=raw_data.get("recommendations", []),
        suggested_slide_count=raw_data.get("suggested_slide_count", suggested_slides),
        content_types_detected=raw_data.get(
            "content_types_detected",
            list(content_types_seen) or ["TITLE_AND_CONTENT"]
        ),
        summary=raw_data.get("summary", summary),
    )
