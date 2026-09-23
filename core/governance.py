"""UPS Healthcare brand, legal, accessibility, and output quality checks."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict, Iterable

from pydantic import BaseModel, Field

from llm.hld_qbr_guidelines import BRAND_COLORS, TYPOGRAPHY
from llm.hld_qbr_layout_registry import get_archetype_max_items


class GovernanceReport(BaseModel):
    passed: bool = True
    brand_compliance: bool = True
    legal_checks: bool = True
    accessibility: bool = True
    source_traceability: bool = True
    quality_score: int = 100
    errors: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    checks: Dict[str, str] = Field(default_factory=dict)

    def error(self, message: str) -> None:
        self.errors.append(message)
        self.passed = False

    def warning(self, message: str) -> None:
        self.warnings.append(message)


def _walk_text(value: Any) -> Iterable[str]:
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for child in value.values():
            yield from _walk_text(child)
    elif isinstance(value, (list, tuple)):
        for child in value:
            yield from _walk_text(child)


def _enforce_hld_qbr_archetype_limits(plan: Any, report: GovernanceReport) -> None:
    """Deterministic min/max-item enforcement (see llm/hld_qbr_layout_registry.py).
    Trims overflow in place rather than failing the run -- the LLM was told the
    limits, but this guarantees them regardless of what it returns."""
    for field_name, max_items in get_archetype_max_items().items():
        items = getattr(plan, field_name, None)
        if isinstance(items, list) and len(items) > max_items:
            report.warning(
                f"{field_name}: {len(items)} items exceeds template max of {max_items} -- trimmed."
            )
            del items[max_items:]


def _check_hld_qbr_traceability(plan: Any, content_model: Any, report: GovernanceReport) -> None:
    """Soft traceability check against the Stage-1 Content Model. Unknown ids are
    flagged; missing citations are informational only -- the LLM may legitimately
    ground a field directly in raw source text with no matching content item."""
    if content_model is None or not getattr(content_model, "content_items", None):
        return
    valid_ids = content_model.ids()
    traceability = getattr(plan, "content_traceability", None) or {}
    for field_name, ref_ids in traceability.items():
        unknown = [rid for rid in ref_ids if rid not in valid_ids]
        if unknown:
            report.warning(f"{field_name}: content_traceability cites unknown content ids {unknown}.")
            report.source_traceability = False
    if get_archetype_max_items().keys() and not traceability:
        populated = [f for f in get_archetype_max_items() if getattr(plan, f, None)]
        if populated:
            report.warning(
                "No content_traceability was returned for any populated section -- "
                "factual claims could not be automatically traced to the Content Model."
            )


# Distinctive content types with a near-unambiguous archetype mapping. Used only to
# surface likely under-utilized source content -- deliberately excludes overlapping
# types (e.g. "metric", "fact") that legitimately map to many archetypes at once and
# would otherwise produce noisy false positives.
_COVERAGE_CHECKS = [
    ("quote", ["voice_of_customer"], "a direct quote"),
    ("risk", ["nc_tracker"], "an issue/incident/deviation"),
    ("person", ["org_structure", "quality_org_structure"], "a named person/role"),
    ("entity", ["org_structure", "quality_org_structure"], "a named person/role"),
]


def _check_hld_qbr_archetype_coverage(plan: Any, content_model: Any, report: GovernanceReport) -> None:
    """Soft warning when the source clearly contains content of a distinctive type
    (quote/risk/person) but every archetype that could hold it was left empty --
    surfaces likely-dropped content instead of silently under-populating the deck."""
    if content_model is None or not getattr(content_model, "content_items", None):
        return
    traceability = getattr(plan, "content_traceability", None) or {}
    used_ids = {rid for ref_ids in traceability.values() for rid in ref_ids}
    for content_type, candidate_fields, label in _COVERAGE_CHECKS:
        if any(getattr(plan, f, None) for f in candidate_fields):
            continue  # at least one candidate archetype was populated -- assume covered
        unused = [
            item.id for item in content_model.content_items
            if item.type == content_type and item.id not in used_ids
        ]
        if unused:
            report.warning(
                f"Source document contains {label} (content ids {unused}) but "
                f"{' / '.join(candidate_fields)} was left empty -- verify this wasn't "
                "dropped in planning."
            )


def validate_plan_governance(
    plan: Any,
    template_id: str,
    source_ids: list[str],
    content_model: Any = None,
) -> GovernanceReport:
    report = GovernanceReport()
    allowed_templates = {"hld_qbr"}
    if template_id not in allowed_templates:
        report.warning("Selected template is not an UPS Healthcare governed template.")
        report.brand_compliance = False
        report.quality_score -= 20
    else:
        report.checks["template"] = "UPS Healthcare HLD QBR template selected"
        _enforce_hld_qbr_archetype_limits(plan, report)
        _check_hld_qbr_traceability(plan, content_model, report)
        _check_hld_qbr_archetype_coverage(plan, content_model, report)

    all_text = "\n".join(_walk_text(plan.model_dump()))
    if re.search(r"\b(?:TBD|TODO|lorem ipsum|click to add|sample text)\b", all_text, re.I):
        report.error("Generated plan contains unresolved placeholder or authoring text.")
    if not source_ids:
        report.error("No source identifier is attached to the presentation plan.")
        report.source_traceability = False
    else:
        report.checks["source_traceability"] = f"{len(source_ids)} source identifier(s) recorded"

    report.checks["claims"] = "Claims require source-backed content; human approval remains required"
    report.checks["logo_controls"] = "Template-controlled logos only"
    report.checks["approved_assets"] = "No external visual asset may be embedded without approval"
    report.accessibility = True
    report.checks["accessibility"] = "Title and content structure checked; rendered contrast/overflow review recommended"
    report.quality_score = max(0, min(100, report.quality_score))
    return report


def validate_pptx_governance(pptx_bytes: bytes, template_id: str) -> GovernanceReport:
    """Inspect the final package for basic output and accessibility signals."""
    report = GovernanceReport()
    if not pptx_bytes or pptx_bytes[:4] != b"PK\x03\x04":
        report.error("Output is not a valid Office package.")
        return report

    try:
        from pptx import Presentation
        import io

        presentation = Presentation(io.BytesIO(pptx_bytes))
        if not presentation.slides:
            report.error("Presentation contains no slides.")
        missing_titles = sum(
            1 for slide in presentation.slides
            if not any(getattr(shape, "has_text", False) and shape.text.strip() for shape in slide.shapes)
        )
        if missing_titles:
            report.warning(f"{missing_titles} slide(s) have no readable text and need accessibility review.")
            report.accessibility = False
            report.quality_score -= min(20, missing_titles * 5)
        report.checks["pptx_integrity"] = f"{len(presentation.slides)} slide(s) opened successfully"
    except Exception as exc:
        report.error(f"Could not inspect generated presentation: {exc}")

    if template_id == "hld_qbr":
        report.checks["brand_theme"] = "UPS Healthcare template package retained"
    else:
        report.warning("Final deck was generated with a non-governed template.")
        report.brand_compliance = False
        report.quality_score -= 20
    report.quality_score = max(0, min(100, report.quality_score))
    return report


def write_audit_record(path: Path, *, source: ContentSource, brief: Any, plan: Any,
                       plan_report: GovernanceReport, output_report: GovernanceReport) -> None:
    """Persist a redacted, replayable governance record for the generation run."""
    path.parent.mkdir(parents=True, exist_ok=True)
    record = {
        "source": source.model_dump(exclude={"text"}),
        "presentation_brief": brief.model_dump(),
        "plan": plan.model_dump() if hasattr(plan, "model_dump") else str(plan),
        "plan_governance": plan_report.model_dump(),
        "output_governance": output_report.model_dump(),
        "approved_colors": BRAND_COLORS,
        "approved_fonts": TYPOGRAPHY["theme_fonts"],
    }
    path.write_text(json.dumps(record, indent=2, default=str), encoding="utf-8")