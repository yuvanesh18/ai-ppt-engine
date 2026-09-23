"""
llm/hld_qbr_schemas.py — Pydantic schemas for the HLD QBR template's
presentation plan.

Unlike the freeform DynamicPresentationPlan (template1/techm_v3), the HLD QBR
template is archetype-based: each field below maps to one specific, named
template slide (see core/builders/hld_qbr_builder.py). Any field left empty
means that slide is simply omitted from the generated deck — sections are
optional except cover/agenda/closing (always included).
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field, model_validator


# ---------------------------------------------------------------------------
# Unicode Sanitizer — replaces LLM-generated special characters that cause
# CP1252 encoding failures or visual glitches in PowerPoint/python-pptx.
# Applied recursively to all string fields in the final presentation plan.
# ---------------------------------------------------------------------------
_UNICODE_REPLACEMENTS = [
    ("\u2011", "-"),   # non-breaking hyphen → regular hyphen
    ("\u2012", "-"),   # figure dash → hyphen
    ("\u2013", "-"),   # en dash → hyphen
    ("\u2014", " - "), # em dash → spaced hyphen
    ("\u2015", "-"),   # horizontal bar → hyphen
    ("\u00a0", " "),   # non-breaking space → regular space
    ("\u2018", "'"),   # left single quote → apostrophe
    ("\u2019", "'"),   # right single quote → apostrophe
    ("\u201c", '"'),   # left double quote → straight quote
    ("\u201d", '"'),   # right double quote → straight quote
    ("\u2026", "..."), # ellipsis → three dots
    ("\u2022", "*"),   # bullet → asterisk (pptx handles its own bullets)
    ("\u25cf", "*"),   # filled circle → asterisk
    ("\u202f", " "),   # narrow no-break space → regular space
    ("\ufeff", ""),    # byte order mark → empty
]


def _sanitize_str(text: str) -> str:
    """Replace known problematic Unicode characters with safe ASCII equivalents."""
    for char, replacement in _UNICODE_REPLACEMENTS:
        text = text.replace(char, replacement)
    return text


def _sanitize_value(obj: Any) -> Any:
    """Recursively sanitize all string values in dicts, lists, and strings."""
    if isinstance(obj, str):
        return _sanitize_str(obj)
    if isinstance(obj, dict):
        return {k: _sanitize_value(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize_value(item) for item in obj]
    return obj


class OrgPersonItem(BaseModel):
    """One person/role card for an org-structure slide."""
    name: str = ""
    role: str = ""

    @model_validator(mode="before")
    @classmethod
    def _normalize(cls, data):
        if isinstance(data, dict):
            name = data.get("name") or data.get("person") or data.get("full_name") or ""
            role = data.get("role") or data.get("title") or data.get("position") or ""
            return {"name": str(name).strip(), "role": str(role).strip()}
        return data


class PriorityItem(BaseModel):
    """One of up to 3 priority pillars."""
    heading: str = Field(default="Priority", description="Short pillar heading, e.g. 'Expand to New Markets'")
    body: str = Field(default="", description="One-sentence supporting detail")

    @model_validator(mode="before")
    @classmethod
    def _normalize(cls, data):
        if isinstance(data, dict):
            heading = data.get("heading") or data.get("title") or data.get("name") or "Priority"
            body = data.get("body") or data.get("description") or data.get("detail") or ""
            return {"heading": str(heading).strip(), "body": str(body).strip()}
        return data


class ActionTrackerRow(BaseModel):
    project: str = ""
    owner: str = ""
    next_step: str = ""
    comment: str = ""
    status: str = "Not Started"


class KPIRow(BaseModel):
    """One KPI row: label + actual value + target value."""
    label: str = ""
    actual: str = ""
    target: str = ""


class GembaRow(BaseModel):
    area: str = ""
    observation: str = ""


class CIRow(BaseModel):
    """Continuous Improvement activity tracker row."""
    activity: str = ""
    category: str = ""
    status: str = "Not Started"
    value: str = ""
    comment: str = ""


class NonConformanceRow(BaseModel):
    period: str = ""
    nc_id: str = ""
    event: str = ""
    due_date: str = ""
    status: str = ""


class NextStepItem(BaseModel):
    step: str = ""
    date: str = ""


class StatHighlightItem(BaseModel):
    """One hero-metric callout badge (mined from the template's unused 'Performance Summary' slide)."""
    label: str = ""
    value: str = ""

    @model_validator(mode="before")
    @classmethod
    def _normalize(cls, data):
        if isinstance(data, dict):
            label = data.get("label") or data.get("metric") or data.get("name") or ""
            value = data.get("value") or data.get("stat") or data.get("number") or ""
            return {"label": str(label).strip(), "value": str(value).strip()}
        return data


class ProcessFlowStep(BaseModel):
    """One chevron segment (mined from the template's unused 'Process Flow Comparison' slide)."""
    heading: str = ""
    description: str = ""


class VoiceOfCustomer(BaseModel):
    quote: str = ""
    attribution: str = ""


class ChartSeries(BaseModel):
    """One data series in a chart (e.g. 'Previous Month' or 'Current Month')."""
    name: str = ""
    values: List[float] = Field(default_factory=list)

    @model_validator(mode="before")
    @classmethod
    def _clean_values(cls, data):
        if isinstance(data, dict):
            import re
            name = data.get("name") or data.get("series_name") or data.get("label") or ""
            raw_vals = data.get("values") or []
            cleaned = []
            for v in raw_vals:
                if isinstance(v, (int, float)):
                    cleaned.append(float(v))
                elif isinstance(v, str):
                    m = re.search(r"[-+]?\d*\.?\d+", v.replace(",", ""))
                    if m:
                        try:
                            cleaned.append(float(m.group(0)))
                        except ValueError:
                            cleaned.append(0.0)
                    else:
                        cleaned.append(0.0)
                else:
                    cleaned.append(0.0)
            return {"name": str(name).strip(), "values": cleaned}
        return data


class OperationalChart(BaseModel):
    """Clustered column chart slide with side observation bullets (template Slide 13 archetype)."""
    chart_title: str = Field(default="Operational Performance Snapshot", description="Slide and chart title")
    categories: List[str] = Field(default_factory=list, description="Up to 7 metric or category names")
    series: List[ChartSeries] = Field(default_factory=list, description="Up to 3 series, e.g. Previous Month, Current Month")
    insights_title: str = Field(default="KEY OBSERVATIONS", description="Title for the side observation card, e.g. 'KEY OBSERVATIONS' or 'OPERATIONAL INSIGHTS'")
    insights: List[str] = Field(default_factory=list, description="Up to 7 key bullet observations for the side panel")

    @model_validator(mode="before")
    @classmethod
    def _normalize(cls, data):
        if isinstance(data, dict):
            chart_title = data.get("chart_title") or data.get("title") or "Operational Performance Snapshot"
            categories = data.get("categories") or data.get("metrics") or data.get("labels") or []
            series = data.get("series") or []
            insights_title = data.get("insights_title") or data.get("insights_heading") or "KEY OBSERVATIONS"
            insights = data.get("insights") or data.get("observations") or data.get("takeaways") or []
            return {
                "chart_title": str(chart_title).strip(),
                "categories": [str(c).strip() for c in categories if str(c).strip()],
                "series": series,
                "insights_title": str(insights_title).strip(),
                "insights": [str(ins).strip() for ins in insights if str(ins).strip()],
            }
        return data


class KPITableSlide(BaseModel):
    """A structured data or KPI table slide (template Slide 12 archetype)."""
    table_title: str = Field(default="Key Performance Indicator Dashboard", description="Slide title")
    headers: List[str] = Field(default_factory=lambda: ["Metric", "Actual", "Target"], description="Column headers (2 to 5 columns)")
    rows: List[List[str]] = Field(default_factory=list, description="Table data rows (up to 7 rows)")

    @model_validator(mode="before")
    @classmethod
    def _normalize_table(cls, data):
        if isinstance(data, dict):
            title = data.get("table_title") or data.get("title") or "Key Performance Indicator Dashboard"
            headers_raw = data.get("headers") or data.get("columns") or ["Metric", "Actual", "Target"]
            rows_raw = data.get("rows") or data.get("data") or []
            headers = [str(h).strip() for h in headers_raw if str(h).strip()]
            rows = []
            for r in rows_raw:
                if isinstance(r, list):
                    rows.append([str(c).strip() for c in r])
                elif isinstance(r, dict):
                    rows.append([str(r.get(h, "")).strip() for h in headers])
                else:
                    rows.append([str(r).strip()])
            return {
                "table_title": str(title).strip(),
                "headers": headers or ["Metric", "Actual", "Target"],
                "rows": rows,
            }
        return data


class HLDQBRPresentationPlan(BaseModel):
    """Top-level plan consumed by core.builders.hld_qbr_builder.HLDQBRBuilder."""

    presentation_title: str = Field(..., description="Main report title, shown in the right-corner cover box")
    facility_name: str = Field(default="", description="Facility/program label shown on content slides")
    date: Optional[str] = Field(default="", description="Cover date string, e.g. 'September 20th 2026'")

    # Dynamic slide headings — LLM-generated per content (must NOT be template names)
    section_heading: str = Field(default="", description="Custom heading for the Performance section divider slide, e.g. 'SUPPLY CHAIN PERFORMANCE REVIEW'")
    agenda_title: str = Field(default="", description="Custom heading for the Agenda slide, e.g. 'TODAY\'S AGENDA' or 'MEETING OVERVIEW'")
    executive_summary_title: str = Field(default="", description="Custom heading for the Executive Summary slide, e.g. 'EXECUTIVE SUMMARY: STRATEGIC & OPERATIONAL HIGHLIGHTS'")
    achievements_title: str = Field(default="", description="Custom heading for the Achievements slide, e.g. 'Q3 2026 OPERATIONAL MILESTONES'")
    priorities_title: str = Field(default="", description="Custom heading for the Priorities slide, e.g. 'STRATEGIC PRIORITIES & INITIATIVES'")
    action_tracker_title: str = Field(default="", description="Custom heading for the Action Item Tracker slide, e.g. 'OPEN ACTION ITEMS & ACCOUNTABILITY'")
    ci_section_title: str = Field(default="", description="Custom heading for the Continuous Improvement section divider, e.g. 'CONTINUOUS IMPROVEMENT INITIATIVES & VALUE CREATION'")
    quality_section_title: str = Field(default="", description="Custom heading for the Quality Management section divider, e.g. 'QUALITY MANAGEMENT & REGULATORY ASSURANCE'")
    next_steps_title: str = Field(default="", description="Custom heading for the Next Steps slide, e.g. 'IMMEDIATE NEXT STEPS & TIMELINE'")
    org_structure_title: str = Field(default="", description="Custom heading for the Org Structure slide, e.g. 'ACCOUNT TEAM & KEY CONTACTS'")
    quality_org_structure_title: str = Field(default="", description="Custom heading for the Quality Org Structure slide, e.g. 'QUALITY & COMPLIANCE TEAM'")
    voice_of_customer_title: str = Field(default="", description="Custom heading for the Voice of Customer slide, e.g. 'DIRECT CUSTOMER FEEDBACK'")
    gemba_walk_title: str = Field(default="", description="Custom heading for the Gemba Walk slide, e.g. 'FACILITY WALKTHROUGH FINDINGS'")
    ci_tracker_title: str = Field(default="", description="Custom heading for the CI Tracker slide, e.g. 'IMPROVEMENT INITIATIVE TRACKER'")
    nc_review_title: str = Field(default="", description="Custom heading for the Non-Conformance Review slide, e.g. 'ISSUE REVIEW & CAPA SUMMARY'")
    nc_tracker_title: str = Field(default="", description="Custom heading for the Non-Conformance Tracker slide, e.g. 'ISSUE & CORRECTIVE ACTION TRACKER'")
    kpi_dashboard_title: str = Field(default="", description="Custom heading for the KPI Dashboard slide, e.g. 'SAFETY, QUALITY & OPERATIONAL KPIs'")
    stat_highlights_title: str = Field(default="", description="Custom heading for the Stat Highlights slide, e.g. 'PERFORMANCE AT A GLANCE'")
    process_flow_title: str = Field(default="", description="Custom heading for the Process Flow slide, e.g. 'ONBOARDING WORKFLOW'")

    agenda_topics: List[str] = Field(default_factory=list)
    executive_summary: List[str] = Field(default_factory=list, description="3 to 5 key takeaways for Executive Summary slide")

    org_structure: List[OrgPersonItem] = Field(default_factory=list, description="Up to 10 people")
    achievements: List[str] = Field(default_factory=list, description="Up to 6 prior-quarter milestones")
    priorities: List[PriorityItem] = Field(default_factory=list, description="Up to 3 pillars")

    # Multi-chart and multi-table support for data-heavy presentations
    charts: List[OperationalChart] = Field(default_factory=list, description="List of column charts with side observations")
    kpi_tables: List[KPITableSlide] = Field(default_factory=list, description="List of structured data or KPI table slides")

    action_tracker: List[ActionTrackerRow] = Field(default_factory=list, description="Up to 7 rows")

    # Legacy single-instance fields maintained for complete backwards-compatibility
    kpi_safety_quality: List[KPIRow] = Field(default_factory=list, description="Up to 4 rows")
    kpi_operational: List[KPIRow] = Field(default_factory=list, description="Up to 7 rows")
    operational_chart: Optional[OperationalChart] = Field(default=None, description="Comparative monthly/quarterly chart with side observations")

    voice_of_customer: Optional[VoiceOfCustomer] = None

    gemba_walk_intro: str = ""
    gemba_walk: List[GembaRow] = Field(default_factory=list, description="Up to 4 rows")

    ci_tracker: List[CIRow] = Field(default_factory=list, description="Up to 7 rows")

    quality_org_structure: List[OrgPersonItem] = Field(default_factory=list, description="Up to 5 people")

    nc_review_summary: List[str] = Field(
        default_factory=list,
        description="4 values: [total NC initiated, total CAPAs assigned, total CAPAs closed, percent complete]",
    )
    nc_review_narrative: str = ""
    nc_tracker: List[NonConformanceRow] = Field(default_factory=list, description="Up to 7 rows")

    next_steps: List[NextStepItem] = Field(default_factory=list, description="Up to 4 rows")

    stat_highlights: List[StatHighlightItem] = Field(default_factory=list, description="Up to 6 hero metric callouts")
    process_flow: List[ProcessFlowStep] = Field(default_factory=list, description="Up to 6 sequential process/workflow steps")

    content_traceability: Dict[str, List[str]] = Field(
        default_factory=dict,
        description=(
            "Maps each populated archetype field name (e.g. 'priorities', 'kpi_operational') "
            "to the Content Model item ids that support it. Only cite ids that genuinely "
            "support the field's text; omit a field entirely if it was not grounded in the "
            "supplied Content Model."
        ),
    )

    narrative_order: List[str] = Field(
        default_factory=list,
        description=(
            "LLM-authored slide sequence (archetype/DIVIDER dispatch keys) between the "
            "mandatory Executive Summary and Closing slides. Empty means the builder falls "
            "back to its default fixed order."
        ),
    )

    @model_validator(mode="before")
    @classmethod
    def _coerce_and_sanitize(cls, data):
        """1) Unwrap {{plan: ...}} wrapper if LLM used it.  2) Harmonize charts/tables. 3) Sanitize all Unicode."""
        if isinstance(data, dict) and "plan" in data and "presentation_title" not in data:
            data = data["plan"]

        if isinstance(data, dict):
            # Bidirectional harmonization between operational_chart and charts
            charts_val = data.get("charts")
            op_chart = data.get("operational_chart")
            if isinstance(charts_val, list) and charts_val:
                if not op_chart:
                    data["operational_chart"] = charts_val[0]
            elif op_chart and isinstance(op_chart, dict):
                data["charts"] = [op_chart]

            # Harmonize kpi_operational into kpi_tables if kpi_tables is not provided
            kpi_tables_val = data.get("kpi_tables")
            kpi_op = data.get("kpi_operational")
            if (not kpi_tables_val) and isinstance(kpi_op, list) and kpi_op:
                rows = []
                for row in kpi_op:
                    if isinstance(row, dict):
                        rows.append([
                            str(row.get("label", "")),
                            str(row.get("actual", "")),
                            str(row.get("target", "") or "-"),
                        ])
                if rows:
                    data["kpi_tables"] = [{
                        "table_title": "Key Performance Indicator Dashboard",
                        "headers": ["Metric", "Actual", "Target"],
                        "rows": rows,
                    }]

        # Recursively sanitize every string value to avoid CP1252/encoding issues in pptx
        return _sanitize_value(data)
