"""
core/presentation_planner_hld_qbr.py — LLM Planner for the HLD QBR Template.

Two-stage planning:
  Stage 1 (_select_archetypes): a small call that picks WHICH archetypes best
    fit this specific source document, sized to the user's requested slide
    count. This is what lets the LLM genuinely choose the best slides/layout
    instead of defaulting to the same handful regardless of content or count.
  Stage 2 (this module's main call): generates rich, enriched content ONLY
    for the archetypes Stage 1 selected, so token budget isn't spread thin
    across 16 possible fields when only a handful were actually chosen.
"""
from __future__ import annotations

import re
from typing import Any, Dict, Optional

from pydantic import ValidationError

from llm.groq_client import GroqClient, JSONParseError
from llm.content_model_schemas import ContentModel
from llm.hld_qbr_schemas import HLDQBRPresentationPlan
from llm.prompts_hld_qbr import (
    REPEATABLE_FIELDS,
    SYSTEM_ROLE_HLD_QBR_ARCHITECT,
    SYSTEM_ROLE_HLD_QBR_SELECTOR,
    build_hld_qbr_archetype_selection_prompt,
    build_hld_qbr_planning_prompt,
    max_selectable_content_slides,
    parse_archetype_selection,
    parse_narrative_order,
)
from llm.schemas import ContentAnalysis
from utils.logging_utils import get_logger
from utils.text_utils import truncate_text

logger = get_logger(__name__)

# Cover, Agenda, Executive Summary, and Closing are always rendered by the builder
# regardless of content — they are mandatory structural slides and must never be
# subtracted from the user's requested content-slide target.
MANDATORY_SLIDE_COUNT = 4
# Upper bound on content slides the template can actually support — derived from
# the archetype registry (see llm/prompts_hld_qbr.py) so it can never silently
# drift out of sync and under-cap a legitimate large request (e.g. 20 slides).
MAX_CONTENT_SLIDES = max_selectable_content_slides()
# Stage 2 is split into multiple calls once a selection exceeds this many slides'
# worth of content, so no single call's token budget gets starved by trying to
# write rich content for too many archetypes at once (see _split_selection_into_batches).
CONTENT_BATCH_SIZE = 7


class HLDQBRPlanningError(Exception):
    """Raised when HLD QBR presentation planning fails."""


def _format_compact_analysis(ca: ContentAnalysis) -> str:
    """Produces a concise, token-efficient text summary of content analysis."""
    parts = [f"Topic: {ca.main_topic}"]
    if ca.key_concepts:
        parts.append(f"Key Concepts: {'; '.join(ca.key_concepts[:10])}")
    if ca.statistics:
        parts.append(f"Statistics & Metrics: {'; '.join(ca.statistics[:16])}")
    if ca.sections:
        parts.append(f"Key Sections: {'; '.join(ca.sections[:10])}")
    if ca.processes:
        parts.append(f"Processes: {'; '.join(ca.processes[:5])}")
    highlights = getattr(ca, "executive_highlights", None)
    if highlights:
        parts.append(f"Highlights: {'; '.join(highlights[:6])}")
    if ca.summary:
        parts.append(f"Summary: {ca.summary[:600]}")
    return "\n".join(parts)


def _select_archetypes(
    client: GroqClient,
    analysis_digest: str,
    source_excerpt: str,
    content_slide_target: Optional[int],
    content_model_json: Optional[str],
) -> tuple[Optional[Dict[str, Any]], list[str]]:
    """Stage 1: asks the LLM which archetypes best fit this document AND in what
    narrative order. Returns (None, []) on failure — None triggers the
    full-catalog fallback and [] triggers the builder's default fixed order —
    so a Stage-1 hiccup never blocks the whole presentation from generating."""
    selection_prompt = build_hld_qbr_archetype_selection_prompt(
        content_analysis=analysis_digest,
        source_content=source_excerpt,
        requested_slide_count=content_slide_target,
        content_model_json=content_model_json,
    )
    messages = [
        {"role": "system", "content": SYSTEM_ROLE_HLD_QBR_SELECTOR},
        {"role": "user", "content": selection_prompt},
    ]
    try:
        raw = client.chat_complete_json(messages=messages, temperature=0.2, max_tokens=800)
        selection = parse_archetype_selection(raw)
        order = parse_narrative_order(raw)
        logger.info("HLD QBR archetype selection: %s | order: %s", {k: v for k, v in selection.items() if v}, order)
        return selection, order
    except Exception as e:
        logger.warning("HLD QBR archetype selection failed (%s) — falling back to full catalog", e)
        return None, []


def _selection_slide_count(selection: Optional[Dict[str, Any]]) -> int:
    """Total rendered slides a selection maps to. Note: kpi_safety_quality and
    kpi_operational share ONE physical dashboard slide (never two), and that
    slide is skipped entirely when kpi_tables is also selected — see the
    if/elif in core/builders/hld_qbr_builder.py's KPI section."""
    if not selection:
        return 0
    total = sum(int(selection.get(f, 0) or 0) for f in REPEATABLE_FIELDS)
    for field, val in selection.items():
        if field in REPEATABLE_FIELDS or field in ("kpi_safety_quality", "kpi_operational"):
            continue
        if val:
            total += 1
    kpi_dashboard_wanted = bool(selection.get("kpi_safety_quality") or selection.get("kpi_operational"))
    if kpi_dashboard_wanted and not selection.get("kpi_tables"):
        total += 1
    return total


def _split_selection_into_batches(
    selection: Dict[str, Any], batch_size: int = CONTENT_BATCH_SIZE
) -> list[Dict[str, Any]]:
    """Splits a large Stage-1 selection into per-call batches so Stage 2 never has
    to write more than ~batch_size slides' worth of enriched content in one LLM
    call. Repeatable fields (charts/kpi_tables) are kept together as one unit
    since splitting a single chart/table set across calls would fragment it."""
    on_fields = [f for f, v in selection.items() if v]
    total_weight = _selection_slide_count(selection)
    if total_weight <= batch_size:
        return [selection]

    repeatable_fields = [f for f in on_fields if f in REPEATABLE_FIELDS]
    boolean_fields = [f for f in on_fields if f not in REPEATABLE_FIELDS]

    batches: list[Dict[str, Any]] = []
    current: Dict[str, Any] = {f: (0 if f in REPEATABLE_FIELDS else False) for f in selection}
    current_weight = 0

    def flush():
        nonlocal current, current_weight
        if current_weight > 0:
            batches.append(current)
        current = {f: (0 if f in REPEATABLE_FIELDS else False) for f in selection}
        current_weight = 0

    for f in repeatable_fields:
        count = selection[f]
        if current_weight and current_weight + count > batch_size:
            flush()
        current[f] = count
        current_weight += count

    for f in boolean_fields:
        if current_weight and current_weight + 1 > batch_size:
            flush()
        current[f] = True
        current_weight += 1

    flush()
    return batches or [selection]


def plan_hld_qbr_presentation(
    client: GroqClient,
    content_analysis: ContentAnalysis,
    source_text: str = "",
    presentation_title: Optional[str] = None,
    facility_name: str = "",
    audience: Optional[str] = None,
    style: Optional[str] = None,
    language: Optional[str] = None,
    additional_instructions: Optional[str] = None,
    slide_count: Optional[int] = None,
    content_model: Optional[ContentModel] = None,
    max_source_chars: int = 14000,
) -> HLDQBRPresentationPlan:
    """Executes the HLD QBR Architect LLM call and returns a validated plan.

    ``slide_count`` is the user's requested number of CONTENT slides.
    Mandatory structural slides (Cover, Agenda, Executive Summary, Closing)
    are always included by the builder and are not subtracted from this count.
    """
    if not presentation_title or not presentation_title.strip():
        presentation_title = content_analysis.main_topic

    content_slide_target: Optional[int] = None
    if slide_count is not None:
        # slide_count is the target number of CONTENT slides
        content_slide_target = max(1, min(MAX_CONTENT_SLIDES, slide_count))

    # Sanitize citation markers and duplicate blank lines from PDF extract
    cleaned_source = re.sub(r"[\u25a0I]cite[\u25a0I][^\u25a0I\n]+[\u25a0I]", "", source_text)
    cleaned_source = re.sub(r"\n{3,}", "\n\n", cleaned_source).strip()

    # Cap context so input tokens (~4,000) + max_tokens (~2,600) stay safely within Groq's 8,000 TPM limit
    effective_max_chars = min(max_source_chars, 8800)
    if slide_count is not None and slide_count <= 4:
        effective_max_chars = min(max_source_chars, 5500)

    truncated_source = truncate_text(cleaned_source, effective_max_chars)
    analysis_digest = _format_compact_analysis(content_analysis)
    content_model_json = content_model.compact_json() if content_model and content_model.content_items else None

    # Stage 1: pick the best-fitting archetypes for THIS document, sized to the
    # requested count, before spending any budget on writing their content.
    selection, narrative_order = _select_archetypes(
        client,
        analysis_digest,
        truncate_text(cleaned_source, 4000),
        content_slide_target,
        content_model_json,
    )

    logger.info("Running HLD QBR Planning for '%s'", presentation_title)

    # Stage 2: split into batches once the selection is large, so no single call's
    # token budget gets starved trying to write rich content for too many slides
    # at once. Small/typical selections (the common case) run as a single batch,
    # identical to the pre-batching behavior.
    batches = _split_selection_into_batches(selection) if selection else [None]
    merged_plan_dict: Dict[str, Any] = {}
    for batch_idx, batch_selection in enumerate(batches):
        batch_prompt = build_hld_qbr_planning_prompt(
            content_analysis=analysis_digest,
            source_content=truncated_source,
            presentation_title=presentation_title,
            facility_name=facility_name or "",
            audience=audience or "Executive Leadership",
            style=style or "Executive Corporate",
            language=language or "English",
            additional_instructions=additional_instructions or "None",
            requested_slide_count=content_slide_target,
            content_model_json=content_model_json,
            selection=batch_selection,
        )
        batch_messages = [
            {"role": "system", "content": SYSTEM_ROLE_HLD_QBR_ARCHITECT},
            {"role": "user", "content": batch_prompt},
        ]

        # Token budget scales with how many slides THIS batch is responsible for
        # (not the whole selection), so each one gets a generous content allowance.
        batch_weight = _selection_slide_count(batch_selection) or (content_slide_target or 6)
        calc_max_tokens = min(3800, max(1500, 900 + batch_weight * 220))

        try:
            raw_data = client.chat_complete_json(messages=batch_messages, temperature=0.3, max_tokens=calc_max_tokens)
        except JSONParseError as e:
            raise HLDQBRPlanningError(f"LLM returned invalid JSON during HLD QBR planning: {e}") from e
        except Exception as e:
            raise HLDQBRPlanningError(f"HLD QBR planning LLM call failed: {e}") from e

        if not isinstance(raw_data, dict):
            logger.warning("LLM returned non-dict at root (%s) — requesting correction turn", type(raw_data).__name__)
            correction_messages = list(batch_messages) + [
                {"role": "assistant", "content": str(raw_data)[:800]},
                {"role": "user", "content": "Error: You returned a JSON list at the root. You MUST return a single JSON OBJECT enclosed in { ... } with keys matching HLDQBRPresentationPlan (e.g. presentation_title, agenda_topics, executive_summary, priorities, achievements, action_tracker, charts, kpi_tables, next_steps). Return only the JSON object."}
            ]
            try:
                raw_data = client.chat_complete_json(messages=correction_messages, temperature=0.2, max_tokens=calc_max_tokens)
            except Exception as e:
                logger.warning("Correction turn failed: %s", e)

        batch_dict = raw_data.get("plan", raw_data) if isinstance(raw_data, dict) else {}
        if not isinstance(batch_dict, dict):
            continue

        if batch_idx == 0:
            # First batch owns presentation-wide fields (titles, agenda, exec summary)
            # plus its own archetype content.
            merged_plan_dict.update(batch_dict)
        else:
            # Later batches only contribute the archetype fields THEY were asked
            # for — their redundant titles/mandatory fields are discarded so the
            # first batch's canonical values always win.
            for field, val in (batch_selection or {}).items():
                weight = val if field in REPEATABLE_FIELDS else (1 if val else 0)
                if weight and field in batch_dict:
                    merged_plan_dict[field] = batch_dict[field]

    plan_dict: Any = merged_plan_dict
    if isinstance(plan_dict, dict):
        plan_dict["narrative_order"] = narrative_order

    try:
        plan = HLDQBRPresentationPlan.model_validate(plan_dict)
    except ValidationError as e:
        # Field-level repair: drop ONLY the offending field(s) so every other
        # correctly-populated section the LLM produced survives, instead of
        # discarding the entire plan for one bad field (e.g. voice_of_customer
        # returned as [] instead of a dict/null).
        logger.warning("HLDQBRPresentationPlan direct validation failed: %s — attempting field-level repair", e)
        if not isinstance(plan_dict, dict):
            raise HLDQBRPlanningError(f"Failed to construct valid HLDQBRPresentationPlan: {e}") from e
        repaired_dict = dict(plan_dict)
        last_err: ValidationError = e
        for _ in range(10):
            try:
                plan = HLDQBRPresentationPlan.model_validate(repaired_dict)
                break
            except ValidationError as retry_err:
                last_err = retry_err
                bad_fields = {err["loc"][0] for err in retry_err.errors() if err["loc"]}
                if not bad_fields or not bad_fields & repaired_dict.keys():
                    raise HLDQBRPlanningError(
                        f"Failed to construct valid HLDQBRPresentationPlan: {retry_err}"
                    ) from retry_err
                for field in bad_fields:
                    logger.warning("Dropping invalid field '%s' during HLD QBR repair (falls back to schema default)", field)
                    repaired_dict.pop(field, None)
        else:
            raise HLDQBRPlanningError(
                f"Failed to construct valid HLDQBRPresentationPlan after repeated repair attempts: {last_err}"
            ) from last_err

    populated_content_slides = _count_populated_content_slides(plan)
    logger.info(
        "HLD QBR Plan validated: %d agenda topics, %d charts, %d tables, %d action rows. "
        "Content slides populated: %d/%s (requested total incl. mandatory: %s).",
        len(plan.agenda_topics), len(plan.charts), len(plan.kpi_tables), len(plan.action_tracker),
        populated_content_slides,
        content_slide_target if content_slide_target is not None else "auto",
        slide_count if slide_count is not None else "auto",
    )
    return plan


def _count_populated_content_slides(plan: HLDQBRPresentationPlan) -> int:
    """Mirrors core.builders.hld_qbr_builder's per-section inclusion checks —
    used only for logging/telemetry, never to gate or trigger replanning."""
    charts_count = len(plan.charts) if plan.charts else (1 if plan.operational_chart else 0)
    tables_count = len(plan.kpi_tables) if plan.kpi_tables else (1 if (plan.kpi_safety_quality or plan.kpi_operational) else 0)
    return (
        charts_count
        + tables_count
        + sum(
            bool(populated)
            for populated in (
                plan.org_structure,
                plan.achievements,
                plan.priorities,
                plan.action_tracker,
                plan.voice_of_customer and plan.voice_of_customer.quote,
                plan.gemba_walk,
                plan.ci_tracker,
                plan.quality_org_structure,
                plan.nc_review_narrative or plan.nc_review_summary,
                plan.nc_tracker,
                plan.next_steps,
            )
        )
    )
