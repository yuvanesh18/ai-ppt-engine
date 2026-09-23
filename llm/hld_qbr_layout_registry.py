"""
llm/hld_qbr_layout_registry.py — Structured Layout/Archetype Registry for the
HLD QBR template.

The HLD QBR template is archetype-based (see llm/hld_qbr_schemas.py): each
entry below documents one named, fixed template slide already implemented in
core/builders/hld_qbr_builder.py. This registry formalizes constraints that
previously only existed as prose in llm/prompts_hld_qbr.py, so they can be
checked deterministically instead of trusted to the LLM.
"""
from __future__ import annotations

from typing import Any, Dict, List, TypedDict


class ArchetypeSpec(TypedDict):
    layout_id: str
    plan_field: str
    purpose: str
    supported_content_types: List[str]
    required_content: List[str]
    optional_content: List[str]
    constraints: Dict[str, int]
    guardrails: List[str]


HLD_QBR_ARCHETYPES: List[ArchetypeSpec] = [
    {
        "layout_id": "ORG_STRUCTURE", "plan_field": "org_structure",
        "purpose": "Show key people/roles for the account team",
        "supported_content_types": ["person", "entity"],
        "required_content": ["name"], "optional_content": ["role"],
        "constraints": {"max_items": 10},
        "guardrails": ["Do not invent people or roles not present in the source"],
    },
    {
        "layout_id": "ACHIEVEMENTS", "plan_field": "achievements",
        "purpose": "Prior-quarter milestones, in chronological order",
        "supported_content_types": ["outcome", "goal", "fact"],
        "required_content": ["text"], "optional_content": [],
        "constraints": {"max_items": 6},
        "guardrails": ["Do not invent milestones or launches not described in the source"],
    },
    {
        "layout_id": "PRIORITIES", "plan_field": "priorities",
        "purpose": "Customer/business priority pillars",
        "supported_content_types": ["goal", "key_message"],
        "required_content": ["heading"], "optional_content": ["body"],
        "constraints": {"max_items": 3},
        "guardrails": ["Do not invent strategic priorities not implied by the source"],
    },
    {
        "layout_id": "ACTION_TRACKER", "plan_field": "action_tracker",
        "purpose": "Open action items with owner/status",
        "supported_content_types": ["step", "process"],
        "required_content": ["project"], "optional_content": ["owner", "next_step", "comment", "status"],
        "constraints": {"max_items": 7},
        "guardrails": ["Do not invent owners, statuses, or projects not present in the source"],
    },
    {
        "layout_id": "KPI_DASHBOARD_SAFETY_QUALITY", "plan_field": "kpi_safety_quality",
        "purpose": "Safety/quality KPI table (actual vs target)",
        "supported_content_types": ["metric"],
        "required_content": ["label", "actual"], "optional_content": ["target"],
        "constraints": {"max_items": 4},
        "guardrails": ["Never invent metric values or targets not stated in the source"],
    },
    {
        "layout_id": "KPI_DASHBOARD_OPERATIONAL", "plan_field": "kpi_operational",
        "purpose": "Operational KPI table (actual vs target)",
        "supported_content_types": ["metric"],
        "required_content": ["label", "actual"], "optional_content": ["target"],
        "constraints": {"max_items": 7},
        "guardrails": ["Never invent metric values or targets not stated in the source"],
    },
    {
        "layout_id": "VOICE_OF_CUSTOMER", "plan_field": "voice_of_customer",
        "purpose": "A single direct quote from any customer, client, stakeholder, or survey respondent (not limited to a literal 'Voice of the Customer' section)",
        "supported_content_types": ["quote"],
        "required_content": ["quote"], "optional_content": ["attribution"],
        "constraints": {"max_items": 1},
        "guardrails": ["Only use a quote that is directly attributable to the source"],
    },
    {
        "layout_id": "GEMBA_WALK", "plan_field": "gemba_walk",
        "purpose": "Any area/location + observation table — site walkthroughs, facility audits, store/branch visits, service reviews, or similar structured observations (regardless of whether the source uses the term 'Gemba')",
        "supported_content_types": ["process", "fact"],
        "required_content": ["area", "observation"], "optional_content": [],
        "constraints": {"max_items": 4},
        "guardrails": ["Do not invent observations not described in the source"],
    },
    {
        "layout_id": "CI_TRACKER", "plan_field": "ci_tracker",
        "purpose": "Any improvement-initiative/activity tracker with status and value — Lean, Kaizen, process improvement, optimization projects, or general initiative logs",
        "supported_content_types": ["process", "step"],
        "required_content": ["activity"], "optional_content": ["category", "status", "value", "comment"],
        "constraints": {"max_items": 7},
        "guardrails": ["Do not invent CI activities not present in the source"],
    },
    {
        "layout_id": "QUALITY_ORG_STRUCTURE", "plan_field": "quality_org_structure",
        "purpose": "A second, smaller people/role list distinct from 'org_structure' — e.g. a specialized team (quality, compliance, technical, security, or any secondary team) if the source describes more than one team",
        "supported_content_types": ["person", "entity"],
        "required_content": ["name"], "optional_content": ["role"],
        "constraints": {"max_items": 5},
        "guardrails": ["Do not invent people or roles not present in the source"],
    },
    {
        "layout_id": "NC_TRACKER", "plan_field": "nc_tracker",
        "purpose": "Any issue/incident/defect/deviation/complaint tracker table with dates and status — not limited to formal 'non-conformance' terminology",
        "supported_content_types": ["risk", "fact"],
        "required_content": ["event"], "optional_content": ["period", "nc_id", "due_date", "status"],
        "constraints": {"max_items": 7},
        "guardrails": ["Never invent NC/CAPA ids, dates, or statuses not present in the source"],
    },
    {
        "layout_id": "NC_REVIEW_SUMMARY", "plan_field": "nc_review_summary",
        "purpose": "A 4-number issue/incident review scorecard (total initiated, total assigned, total closed, percent complete) — a summary companion to nc_tracker's detailed row-by-row table",
        "supported_content_types": ["metric", "fact"],
        "required_content": ["total", "assigned", "closed", "pct"], "optional_content": [],
        "constraints": {"max_items": 4},
        "guardrails": ["Never invent totals/percentages not stated or computable from the source"],
    },
    {
        "layout_id": "NEXT_STEPS", "plan_field": "next_steps",
        "purpose": "Calls to action / next steps",
        "supported_content_types": ["step", "goal"],
        "required_content": ["step"], "optional_content": ["date"],
        "constraints": {"max_items": 4},
        "guardrails": ["Do not invent next steps not implied by the source"],
    },
    {
        "layout_id": "STAT_HIGHLIGHTS", "plan_field": "stat_highlights",
        "purpose": "3-6 hero-metric callout badges (big number + label) for headline stats worth calling out visually — any domain, not detailed data",
        "supported_content_types": ["metric", "outcome"],
        "required_content": ["label", "value"], "optional_content": [],
        "constraints": {"max_items": 6},
        "guardrails": ["Never invent stat values not stated in the source"],
    },
    {
        "layout_id": "PROCESS_FLOW", "plan_field": "process_flow",
        "purpose": "A left-to-right sequence of process/workflow/roadmap steps (chevron segments) — any multi-step process, not limited to a specific domain",
        "supported_content_types": ["process", "step"],
        "required_content": ["heading"], "optional_content": ["description"],
        "constraints": {"max_items": 6},
        "guardrails": ["Do not invent process steps not described in the source"],
    },
]


def get_archetype_max_items() -> Dict[str, int]:
    """{plan_field: max_items} for every archetype that defines the constraint."""
    return {
        a["plan_field"]: a["constraints"]["max_items"]
        for a in HLD_QBR_ARCHETYPES
        if "max_items" in a["constraints"]
    }


def compact_registry_json() -> List[Dict[str, Any]]:
    """Compact view (no guardrail prose) suitable for prompt injection."""
    return [
        {
            "layout_id": a["layout_id"],
            "plan_field": a["plan_field"],
            "purpose": a["purpose"],
            "supported_content_types": a["supported_content_types"],
            "constraints": a["constraints"],
        }
        for a in HLD_QBR_ARCHETYPES
    ]
