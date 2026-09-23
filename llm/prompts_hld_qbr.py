"""
llm/prompts_hld_qbr.py — Prompts for the HLD QBR Presentation Architect.

Instructs the LLM on the HLD QBR template's archetype catalog (see
HLD_QBR_TEMPLATE_PLAN.md §2) and injects the mined brand/guardrail knowledge
from hld_qbr_guidelines.py, so generated content stays on-brand without ever
needing the template's own "Formatting Help" appendix slides to be shown.

Planning is a TWO-STAGE process to avoid token-budget starvation and to let
the LLM genuinely choose the best-fitting slides instead of defaulting to the
same handful every time:
  Stage 1 (selection): a small, cheap call picks WHICH archetypes to use and
    how many charts/tables, sized to the user's requested slide count.
  Stage 2 (content): a call that only carries instructions/output-keys for
    the SELECTED archetypes, so every chosen slide gets full per-field detail
    and generous token budget instead of competing with 15 unused archetypes.
"""
from __future__ import annotations

import json
from typing import Any, Dict, List, Optional

from llm.hld_qbr_guidelines import build_llm_guidelines_prompt
from llm.hld_qbr_layout_registry import compact_registry_json

SYSTEM_ROLE_HLD_QBR_ARCHITECT = """\
You are an Executive Presentation Architect for UPS Healthcare Logistics & Distribution QBRs.
Extract concrete operational facts, metrics, dollar values, percentages, facility names, and root-cause analyses from the source document to synthesize an authoritative, boardroom-ready Quarterly Business Review.
The source document may come from ANY business domain and will rarely mention the template's own section names — your job is to map its real content onto the fixed brand/template structure by meaning, not by matching vocabulary.

Core Rules:
1. DEEP FACTUAL GROUNDING: Every bullet, achievement, priority, and chart insight MUST cite real metrics, data points, or operational challenges from the source. No vague generalities or corporate fluff.
2. CHART INTEGRITY: Group metrics with identical units/scales (e.g. percentages together, unit counts together). For each chart, provide 2 concise analytical observations explaining trends, variances, and recommended CAPA.
3. TABULAR DATA: Preserve full multi-column datasets with clear headers and units.
4. STRICT JSON: Return ONLY a valid JSON object matching the requested schema. No conversational prose or markdown wrap.
"""

_ARCHETYPE_MAPPING_PHILOSOPHY = """\
ARCHETYPE MAPPING PHILOSOPHY (read first):
Slide archetype names (e.g. 'gemba_walk', 'voice_of_customer', 'nc_tracker') are STRUCTURAL
CONTAINERS inherited from the template's original design — labels for a SHAPE of content (a
quote+attribution, an area+observation list, an issue/status tracker), not literal topics the
source document must mention by name. Map by MEANING AND STRUCTURE, never by keyword-matching
the archetype's name against the source's vocabulary. Examples:
- A direct quote from any customer, user, patient, employee, or survey respondent -> voice_of_customer, even if the source never says "voice of the customer".
- Any facility/site/store/branch visit notes, audit findings, or process walkthroughs, organized as area + finding -> gemba_walk, even if the source never says "Gemba".
- Any tracked list of issues, incidents, defects, complaints, or deviations with status/dates -> nc_tracker, even if the source never says "non-conformance".
- Any list of improvement initiatives, projects, or experiments with a status/value -> ci_tracker, even if the source never says "continuous improvement".
- Any list of people + their roles/titles, for whichever team the source emphasizes -> org_structure (primary team) or quality_org_structure (a second/specialized team).
- Any metric with an actual value and (optionally) a target/goal -> kpi_safety_quality / kpi_operational / kpi_tables, regardless of the metric's domain (financial, technical, clinical, operational, etc.).
Do NOT skip an archetype merely because the source uses different terminology than the
template — check whether the UNDERLYING STRUCTURE matches before deciding there's no evidence."""

# ---------------------------------------------------------------------------
# Per-archetype prompt fragments. Split so Stage 2 can render ONLY the
# archetypes Stage 1 actually selected — every selected slide gets full,
# dedicated instructions instead of being crammed into one dense "optional"
# line that the model tends to ignore.
# ---------------------------------------------------------------------------
_MANDATORY_FIELD_SPECS: Dict[str, Dict[str, str]] = {
    "agenda_topics": {
        "content_line": "- agenda_topics: 3-5 concise topic titles for Slide 2.",
        "json_key": '  "agenda_topics": ["..."],',
    },
    "executive_summary": {
        "content_line": (
            "- executive_summary: 4-5 high-impact bullets formatted as 'UPPERCASE CATEGORY "
            "(2-4 words): Concrete analytical takeaway citing figures, percentages, and business "
            "impact' (e.g. 'ON-TIME SERVICE EXCELLENCE: OTIF improved from 91.8% to 97.2%, "
            "exceeding the 95% target.')."
        ),
        "json_key": '  "executive_summary": ["CATEGORY: Analytical bullet with metrics..."],',
    },
}

_OPTIONAL_FIELD_SPECS: Dict[str, Dict[str, str]] = {
    "priorities": {
        "content_line": (
            '- priorities: up to 3 strategic pillars [{"heading": "Max 3 words", "body": '
            '"2 sentences with operational friction, lever, and target."}].'
        ),
        "json_key": '  "priorities": [{"heading": "...", "body": "..."}],',
    },
    "achievements": {
        "content_line": "- achievements: up to 4 milestone strings with concrete figures/percentages from prior quarter.",
        "json_key": '  "achievements": ["..."],',
    },
    "charts": {
        "content_line": (
            '- charts: [{"chart_title": "Descriptive content-specific title", "categories": '
            '["Cat1", ...], "series": [{"name": "Series", "values": [12.3, ...]}], '
            '"insights_title": "KEY OBSERVATIONS", "insights": ["2 concise analytical bullets '
            'with metrics & root-cause/CAPA"]}]. (Never mix dissimilar scales like % and volume '
            "counts on the same axis)."
        ),
        "json_key": (
            '  "charts": [{"chart_title": "...", "categories": [...], "series": '
            '[{"name": "...", "values": [...]}], "insights_title": "KEY OBSERVATIONS", '
            '"insights": [...]}],'
        ),
    },
    "kpi_tables": {
        "content_line": (
            '- kpi_tables: [{"table_title": "Descriptive content-specific title", "headers": '
            '["Col 1 (Unit)", "Col 2"], "rows": [["Row1", "..."], ...]}] for multi-column '
            "operational datasets."
        ),
        "json_key": '  "kpi_tables": [{"table_title": "...", "headers": [...], "rows": [[...]]}],',
    },
    "action_tracker": {
        "content_line": (
            '- action_tracker: up to 4 operational rows [{"project": "Initiative name", '
            '"owner": "Role title", "next_step": "Measurable action", "comment": "Rationale '
            'citing source data", "status": "Complete" | "In Progress" | "Delayed"}].'
        ),
        "json_key": (
            '  "action_tracker": [{"project": "...", "owner": "...", "next_step": "...", '
            '"comment": "...", "status": "In Progress"}],'
        ),
    },
    "next_steps": {
        "content_line": '- next_steps: up to 4 rows [{"step": "Action", "date": "e.g. Q4 2026"}].',
        "json_key": '  "next_steps": [{"step": "...", "date": "..."}],',
    },
    "org_structure": {
        "content_line": '- org_structure: up to 10 people [{"name": "...", "role": "..."}] for the primary team.',
        "json_key": '  "org_structure": [{"name": "...", "role": "..."}],',
    },
    "quality_org_structure": {
        "content_line": (
            '- quality_org_structure: up to 5 people [{"name": "...", "role": "..."}] for a second/'
            "specialized team, distinct from org_structure."
        ),
        "json_key": '  "quality_org_structure": [{"name": "...", "role": "..."}],',
    },
    "kpi_safety_quality": {
        "content_line": '- kpi_safety_quality: up to 4 rows [{"label": "...", "actual": "...", "target": "..."}].',
        "json_key": '  "kpi_safety_quality": [{"label": "...", "actual": "...", "target": "..."}],',
    },
    "kpi_operational": {
        "content_line": '- kpi_operational: up to 7 rows [{"label": "...", "actual": "...", "target": "..."}].',
        "json_key": '  "kpi_operational": [{"label": "...", "actual": "...", "target": "..."}],',
    },
    "voice_of_customer": {
        "content_line": '- voice_of_customer: {"quote": "Direct quote from the source", "attribution": "Who said it"}.',
        "json_key": '  "voice_of_customer": {"quote": "...", "attribution": "..."},',
    },
    "gemba_walk": {
        "content_line": '- gemba_walk: up to 4 rows [{"area": "...", "observation": "..."}].',
        "json_key": '  "gemba_walk": [{"area": "...", "observation": "..."}],',
    },
    "ci_tracker": {
        "content_line": (
            '- ci_tracker: up to 7 rows [{"activity": "...", "category": "...", "status": "...", '
            '"value": "...", "comment": "..."}].'
        ),
        "json_key": (
            '  "ci_tracker": [{"activity": "...", "category": "...", "status": "...", '
            '"value": "...", "comment": "..."}],'
        ),
    },
    "nc_review_summary": {
        "content_line": (
            '- nc_review_summary: exactly 4 values ["total initiated", "total assigned", '
            '"total closed", "percent complete"].'
        ),
        "json_key": (
            '  "nc_review_summary": ["total NC initiated", "total CAPAs assigned", '
            '"total CAPAs closed", "percent complete"],'
        ),
    },
    "nc_tracker": {
        "content_line": (
            '- nc_tracker: up to 7 rows [{"period": "...", "nc_id": "...", "event": "...", '
            '"due_date": "...", "status": "..."}].'
        ),
        "json_key": (
            '  "nc_tracker": [{"period": "...", "nc_id": "...", "event": "...", '
            '"due_date": "...", "status": "..."}],'
        ),
    },
    "stat_highlights": {
        "content_line": '- stat_highlights: up to 6 hero stats [{"label": "...", "value": "..."}].',
        "json_key": '  "stat_highlights": [{"label": "...", "value": "..."}],',
    },
    "process_flow": {
        "content_line": '- process_flow: up to 6 steps [{"heading": "...", "description": "..."}].',
        "json_key": '  "process_flow": [{"heading": "...", "description": "..."}],',
    },
}

_TITLE_FIELD_SPECS: Dict[str, str] = {
    "agenda_title": '  "agenda_title": "AGENDA SLIDE HEADING (e.g. TODAY\'S AGENDA)",',
    "executive_summary_title": '  "executive_summary_title": "EXECUTIVE SUMMARY HEADING",',
    "section_heading": '  "section_heading": "CONTENT-SPECIFIC SECTION HEADING (uppercase, 4-7 words)",',
    "achievements_title": '  "achievements_title": "ACHIEVEMENTS SLIDE HEADING",',
    "priorities_title": '  "priorities_title": "PRIORITIES SLIDE HEADING",',
    "action_tracker_title": '  "action_tracker_title": "ACTION TRACKER SLIDE HEADING",',
    "ci_section_title": '  "ci_section_title": "CONTINUOUS IMPROVEMENT SECTION HEADING",',
    "quality_section_title": '  "quality_section_title": "QUALITY MANAGEMENT SECTION HEADING",',
    "next_steps_title": '  "next_steps_title": "NEXT STEPS SLIDE HEADING",',
    "org_structure_title": '  "org_structure_title": "ORG STRUCTURE SLIDE HEADING",',
    "quality_org_structure_title": '  "quality_org_structure_title": "QUALITY ORG STRUCTURE SLIDE HEADING",',
    "voice_of_customer_title": '  "voice_of_customer_title": "VOICE OF CUSTOMER SLIDE HEADING",',
    "gemba_walk_title": '  "gemba_walk_title": "GEMBA WALK / SITE OBSERVATION SLIDE HEADING",',
    "ci_tracker_title": '  "ci_tracker_title": "CI TRACKER SLIDE HEADING",',
    "nc_review_title": '  "nc_review_title": "ISSUE REVIEW SLIDE HEADING",',
    "nc_tracker_title": '  "nc_tracker_title": "ISSUE TRACKER SLIDE HEADING",',
    "kpi_dashboard_title": '  "kpi_dashboard_title": "KPI DASHBOARD SLIDE HEADING",',
    "stat_highlights_title": '  "stat_highlights_title": "STAT HIGHLIGHTS SLIDE HEADING",',
    "process_flow_title": '  "process_flow_title": "PROCESS FLOW SLIDE HEADING",',
}

# Optional fields that produce a variable NUMBER of slides (unlike the rest,
# which are 0-or-1). Selection Stage returns an integer count for these.
REPEATABLE_FIELDS = {"charts", "kpi_tables"}
# Upper bound on instances per repeatable field (e.g. at most 4 chart slides).
MAX_REPEATABLE_INSTANCES = 4


def max_selectable_content_slides() -> int:
    """Upper bound on content slides the template can support in a single deck:
    every boolean archetype once, plus each repeatable archetype's max instances.
    Single source of truth so callers never have to hardcode/re-derive this."""
    boolean_count = len(_OPTIONAL_FIELD_SPECS) - len(REPEATABLE_FIELDS)
    return boolean_count + len(REPEATABLE_FIELDS) * MAX_REPEATABLE_INSTANCES

_SLIDE_HEADING_RULES = """\
SLIDE HEADING RULES (CRITICAL):
Every slide heading below MUST be generated specifically for this presentation and content. Do NOT copy template names like 'PERFORMANCE MANAGEMENT UPDATES', 'CONTINUOUS IMPROVEMENT PROGRAM UPDATES', 'TODAY'S DISCUSSION', 'GEMBA WALK SUMMARY', 'VOICE OF THE CUSTOMER', 'NON-CONFORMANCE REVIEW', 'NON-CONFORMANCE TRACKER', 'ORGANIZATIONAL STRUCTURE', 'QUALITY ORGANIZATIONAL STRUCTURE', or 'KEY PERFORMANCE INDICATOR (KPI) DASHBOARD' — those are just the template's own sample/reference wording, not headings to reuse.
- agenda_title: UPPERCASE heading for the agenda slide (e.g. 'TODAY\'S AGENDA', 'MEETING OVERVIEW & OBJECTIVES', 'QUARTERLY REVIEW AGENDA').
- executive_summary_title: UPPERCASE heading for the executive summary slide (e.g. 'EXECUTIVE SUMMARY: STRATEGIC & OPERATIONAL HIGHLIGHTS', 'EXECUTIVE OVERVIEW: KEY OUTCOMES & IMPACT').
- section_heading: A concise 4-7 word UPPERCASE heading for the performance/operations section divider slide. Must reflect the actual content domain (e.g. 'HEALTHCARE LOGISTICS PERFORMANCE REVIEW', 'QUARTERLY OPERATIONS INTELLIGENCE BRIEFING').
- achievements_title: UPPERCASE heading for the achievements slide (e.g. 'Q3 2026 OPERATIONAL MILESTONES', 'PRIOR QUARTER KEY ACHIEVEMENTS').
- priorities_title: UPPERCASE heading for the priorities slide (e.g. 'STRATEGIC PRIORITIES FOR 2026', 'OPERATIONAL EXCELLENCE FOCUS AREAS').
- action_tracker_title: UPPERCASE heading for the action tracker slide (e.g. 'OPEN ACTION ITEMS & ACCOUNTABILITY', 'ACTION REGISTER & OWNERS').
- ci_section_title: UPPERCASE heading for the continuous improvement section divider slide (e.g. 'CONTINUOUS IMPROVEMENT INITIATIVES & VALUE CREATION').
- quality_section_title: UPPERCASE heading for the quality management section divider slide (e.g. 'QUALITY MANAGEMENT & REGULATORY ASSURANCE').
- next_steps_title: UPPERCASE heading for the next steps slide (e.g. 'NEXT STEPS & TARGET TIMELINES', 'UPCOMING COMMITMENTS').
- org_structure_title: UPPERCASE heading for the org structure slide, reflecting who this team actually is (e.g. 'ACCOUNT TEAM & KEY CONTACTS', 'DEDICATED SUPPORT TEAM').
- quality_org_structure_title: UPPERCASE heading for the second/specialized team slide (e.g. 'QUALITY & COMPLIANCE LEADERSHIP', 'TECHNICAL SUPPORT TEAM').
- voice_of_customer_title: UPPERCASE heading for the quote slide (e.g. 'DIRECT CUSTOMER FEEDBACK', 'IN THEIR OWN WORDS').
- gemba_walk_title: UPPERCASE heading for the area/observation slide, matching the actual setting described (e.g. 'FACILITY WALKTHROUGH FINDINGS', 'SITE AUDIT OBSERVATIONS').
- ci_tracker_title: UPPERCASE heading for the improvement-activity tracker slide (e.g. 'IMPROVEMENT INITIATIVE TRACKER', 'LEAN PROJECT STATUS').
- nc_review_title: UPPERCASE heading for the issue-review summary slide (e.g. 'ISSUE REVIEW & CAPA SUMMARY', 'INCIDENT REVIEW SCORECARD').
- nc_tracker_title: UPPERCASE heading for the issue tracker slide (e.g. 'ISSUE & CORRECTIVE ACTION TRACKER', 'DEVIATION LOG').
- kpi_dashboard_title: UPPERCASE heading for the KPI dashboard slide (e.g. 'SAFETY, QUALITY & OPERATIONAL KPIs', 'CORE PERFORMANCE SCORECARD').
- stat_highlights_title: UPPERCASE heading for the hero-stat callout slide (e.g. 'PERFORMANCE AT A GLANCE').
- process_flow_title: UPPERCASE heading for the process/workflow slide (e.g. 'ONBOARDING WORKFLOW', 'ESCALATION PROCESS')."""


# ---------------------------------------------------------------------------
# Stage 1 — Archetype Selection
# ---------------------------------------------------------------------------
SYSTEM_ROLE_HLD_QBR_SELECTOR = """\
You are a presentation editor deciding which slides a fixed, brand-approved QBR template \
should use for a given source document. You do not write slide content — only select and \
size the slide archetypes. Return ONLY a valid JSON object. No prose, no markdown fences.
"""

_SELECTION_PROMPT = """\
{mapping_philosophy}

Cover, Agenda, Executive Summary, and Closing are mandatory and always included — they are
NOT part of this selection.

ARCHETYPE CATALOG (each is a real, always-available template slide; use "supported_content_types"
and "purpose" to judge fit — purpose describes the content SHAPE, not a literal topic):
{catalog_json}

Two additional variable-count archetypes (not in the catalog above):
- "charts": one slide per genuinely chartable numeric series (e.g. trends, comparisons over time/category). Estimate how many DISTINCT chartable datasets the source actually supports (0-4).
- "kpi_tables": one slide per multi-column tabular dataset that isn't better suited to a chart. Estimate how many DISTINCT tables the source actually supports (0-4).

IMPORTANT SLIDE-COUNT NOTE: "kpi_safety_quality" and "kpi_operational" share ONE combined
dashboard slide (not two) when both are selected — count them as 1 slide toward your target,
not 2. That shared dashboard slide is also skipped entirely if "kpi_tables" is selected —
choose EITHER "kpi_tables" OR "kpi_safety_quality"/"kpi_operational" for tabular KPI data, never
both, to avoid wasted/unused content.

{content_model_section}
DOCUMENT ANALYSIS:
{content_analysis}

SOURCE EXCERPT (for gauging topic coverage only):
{source_content}

REQUESTED CONTENT SLIDE COUNT: {slide_count_text}

TASK: Select the BEST-FITTING archetypes for this specific document, sized to the requested
count. Prioritize whichever archetypes have the STRONGEST, most concrete evidence in the source —
it is better to undershoot the target with well-grounded slides than to select an archetype the
source doesn't actually support. Do not select the same handful out of habit — judge each
archetype independently against this document's real content.

THEN order your selected slides as a NARRATIVE, not a catalog listing — e.g. context/priorities
first, evidence (charts/tables/trackers) next, risks or gaps, then resolution/next steps. Insert
"DIVIDER_PERFORMANCE" / "DIVIDER_CI" / "DIVIDER_QUALITY" only where that section's slides are
placed and a real topic shift occurs — never on a fixed schedule, and never if that section is
empty. Use "kpi_dashboard" (not "kpi_safety_quality"/"kpi_operational" individually) in "order".

OUTPUT SPECIFICATION — return ONLY this JSON object:
{{
  "selected": {{
{selection_keys}
  }},
  "order": ["<selected archetype/DIVIDER keys, in your chosen narrative sequence>"],
  "estimated_total_content_slides": <int>
}}
Each boolean field is true only if the source has clear, concrete supporting evidence.
"charts" and "kpi_tables" are integers (0-4), not booleans.
"""


def _build_selection_keys() -> str:
    lines = []
    for field in _OPTIONAL_FIELD_SPECS:
        if field in REPEATABLE_FIELDS:
            lines.append(f'    "{field}": 0,')
        else:
            lines.append(f'    "{field}": false,')
    return "\n".join(lines)


def build_hld_qbr_archetype_selection_prompt(
    *,
    content_analysis: str,
    source_content: str,
    requested_slide_count: Optional[int] = None,
    content_model_json: Optional[str] = None,
) -> str:
    """Stage 1: a small, cheap call that ONLY decides which archetypes to use."""
    slide_count_text = (
        f"~{requested_slide_count} content slides" if requested_slide_count is not None
        else "no specific target — use your best judgement based on available evidence"
    )
    return _SELECTION_PROMPT.format(
        mapping_philosophy=_ARCHETYPE_MAPPING_PHILOSOPHY,
        catalog_json=json.dumps(compact_registry_json(), indent=2),
        content_model_section=_build_content_model_section(content_model_json),
        content_analysis=content_analysis,
        source_content=source_content,
        slide_count_text=slide_count_text,
        selection_keys=_build_selection_keys(),
    )


def parse_archetype_selection(raw: Dict[str, Any]) -> Dict[str, Any]:
    """Validates/clamps a Stage-1 selection dict against the known field catalog."""
    selected_raw = raw.get("selected", raw) if isinstance(raw, dict) else {}
    selection: Dict[str, Any] = {}
    for field in _OPTIONAL_FIELD_SPECS:
        val = selected_raw.get(field) if isinstance(selected_raw, dict) else None
        if field in REPEATABLE_FIELDS:
            try:
                count = int(val) if val is not None else 0
            except (TypeError, ValueError):
                count = 0
            selection[field] = max(0, min(MAX_REPEATABLE_INSTANCES, count))
        else:
            selection[field] = bool(val)
    return selection


# Divider/combined-slide pseudo-keys the builder's dispatch table also recognizes
# in a narrative order, beyond the plain archetype field names.
_DIVIDER_KEYS = ("DIVIDER_PERFORMANCE", "DIVIDER_CI", "DIVIDER_QUALITY")
_ORDER_PSEUDO_KEYS = ("kpi_dashboard",)  # combines kpi_safety_quality + kpi_operational


def valid_order_keys() -> set:
    """All dispatch keys a narrative order may legally reference."""
    return set(_OPTIONAL_FIELD_SPECS.keys()) | set(_DIVIDER_KEYS) | set(_ORDER_PSEUDO_KEYS)


def parse_narrative_order(raw: Dict[str, Any]) -> List[str]:
    """Extracts and validates the Stage-1 'order' array. Invalid/unknown entries
    are dropped rather than failing the whole call — the builder's own fallback
    to its default order handles an empty/garbled result safely."""
    order_raw = raw.get("order") if isinstance(raw, dict) else None
    if not isinstance(order_raw, list):
        return []
    valid = valid_order_keys()
    seen: set = set()
    cleaned: List[str] = []
    for item in order_raw:
        key = str(item).strip() if isinstance(item, str) else ""
        if key in valid and key not in seen:
            cleaned.append(key)
            seen.add(key)
    return cleaned


# ---------------------------------------------------------------------------
# Stage 2 — Content Generation (only for the archetypes Stage 1 selected)
# ---------------------------------------------------------------------------
HLD_QBR_PRESENTATION_PLANNING_PROMPT = """\
Plan a boardroom-grade QBR presentation based strictly on the source document.
The source document can be about ANY business/industry domain — it will rarely use the
same section names as the template. Your job is to fit that document's real content into
the fixed, brand-approved template below; never force the document to already match it.

{mapping_philosophy}
FALLBACK RULE: If a piece of source content is clearly important and well-grounded but does not
match any specialized archetype's structure, still surface it — place narrative points in
'executive_summary' or 'priorities', tabular data in 'kpi_tables', and numeric series in
'charts' (all four are fully generic, domain-agnostic containers). Never silently discard
important, well-grounded source content just because no specialized archetype fits it.

{guidelines}

CONTENT TO GENERATE (write full, enriched, boardroom-grade content for EVERY field below —
this exact set was already selected as the best fit for this document; do not add or omit
fields, and do not pad any field with filler just to fill space):
{content_archetypes}

{slide_heading_rules}

{content_model_section}
DOCUMENT ANALYSIS:
{content_analysis}

SOURCE CONTENT (ground all data strictly in this text):
{source_content}

METADATA:
- Title: {presentation_title} | Facility: {facility_name} | Audience: {audience} | Style: {style}

OUTPUT SPECIFICATION:
Return ONLY a single valid JSON object in this format:
{{
  "presentation_title": "{presentation_title}",
  "facility_name": "{facility_name}",
  "date": "",
{title_json_keys}
{content_json_keys}
}}
"""


def _build_content_model_section(content_model_json: Optional[str]) -> str:
    """Optional Stage-1 traceable items: included only when available."""
    if not content_model_json:
        return ""
    return (
        "CONTENT MODEL (traceable fact items):\n"
        f"{content_model_json}\n\n"
    )


def _selected_fields(selection: Optional[Dict[str, Any]]) -> List[str]:
    """Field names to render in Stage 2, in catalog order. selection=None means
    'no Stage-1 result available' -> fall back to offering every archetype
    (old single-call behavior), so the pipeline never produces an empty deck."""
    if selection is None:
        return list(_OPTIONAL_FIELD_SPECS.keys())
    fields = []
    for field in _OPTIONAL_FIELD_SPECS:
        val = selection.get(field)
        if field in REPEATABLE_FIELDS:
            if val:
                fields.append(field)
        elif val:
            fields.append(field)
    return fields


def build_hld_qbr_planning_prompt(
    *,
    requested_slide_count: Optional[int] = None,
    content_model_json: Optional[str] = None,
    selection: Optional[Dict[str, Any]] = None,
    **kwargs,
) -> str:
    fields = _selected_fields(selection)

    content_lines = [spec["content_line"] for spec in _MANDATORY_FIELD_SPECS.values()]
    json_lines = [spec["json_key"] for spec in _MANDATORY_FIELD_SPECS.values()]
    for field in fields:
        spec = _OPTIONAL_FIELD_SPECS[field]
        line = spec["content_line"]
        if field in REPEATABLE_FIELDS and selection:
            count = selection.get(field, 0)
            line = f"{line} EXACTLY {count} instance(s) — do not add more or fewer."
        content_lines.append(line)
        json_lines.append(spec["json_key"])

    kwargs.pop("slide_count_guidance", None)  # legacy caller kwarg, ignored in Stage 2
    return HLD_QBR_PRESENTATION_PLANNING_PROMPT.format(
        mapping_philosophy=_ARCHETYPE_MAPPING_PHILOSOPHY,
        guidelines=build_llm_guidelines_prompt(),
        content_archetypes="\n".join(content_lines),
        slide_heading_rules=_SLIDE_HEADING_RULES,
        title_json_keys="\n".join(_TITLE_FIELD_SPECS.values()),
        content_json_keys="\n".join(json_lines).rstrip(","),
        content_model_section=_build_content_model_section(content_model_json),
        **kwargs,
    )

