"""
llm/schemas.py — Pydantic models for all LLM-structured outputs.

These models enforce data contracts between the LLM and the rest of the
application.  If the LLM returns invalid JSON, these models raise
ValidationError before any PowerPoint code is reached.
"""

from __future__ import annotations

from enum import Enum
from typing import Any, Dict, List, Optional, Union

from pydantic import BaseModel, Field, field_validator, model_validator


# ---------------------------------------------------------------------------
# Layout type enum — must stay in sync with config.LAYOUT_REGISTRY
# ---------------------------------------------------------------------------

class LayoutType(str, Enum):
    # ---- Universal visual archetypes (new — LLM-facing) ----
    TITLE             = "TITLE"
    EXECUTIVE_SUMMARY = "EXECUTIVE_SUMMARY"  # slide 2 — always required
    AGENDA            = "AGENDA"             # slide 3 — always required
    BULLETS           = "BULLETS"            # replaces TITLE_AND_CONTENT as default
    SECTION_HEADER    = "SECTION_HEADER"
    CARDS_2_COL       = "CARDS_2_COL"        # replaces TWO_COLUMN / COMPARISON
    CARDS_3_COL       = "CARDS_3_COL"        # replaces PROCESS / THREE_COLUMN
    TABLE             = "TABLE"
    STATS             = "STATS"              # replaces STATISTICS
    IMAGE_TEXT        = "IMAGE_TEXT"         # replaces ARCHITECTURE
    TIMELINE          = "TIMELINE"
    CONCLUSION        = "CONCLUSION"

    # ---- Legacy names (kept for backward compatibility) ----
    TITLE_AND_CONTENT = "TITLE_AND_CONTENT"
    TWO_COLUMN        = "TWO_COLUMN"
    COMPARISON        = "COMPARISON"
    PROCESS           = "PROCESS"
    ARCHITECTURE      = "ARCHITECTURE"
    STATISTICS        = "STATISTICS"


# ---------------------------------------------------------------------------
# Per-layout content shapes
# ---------------------------------------------------------------------------

class TitleContent(BaseModel):
    """Content for TITLE layout."""
    subtitle: Optional[str] = None
    presenter: Optional[str] = None
    date: Optional[str] = None


class ExecutiveSummaryContent(BaseModel):
    """Content for EXECUTIVE_SUMMARY layout (slide 2)."""
    highlights: List[str] = Field(default_factory=list, min_length=1)
    purpose_statement: Optional[str] = None
    key_metric: Optional[str] = None


class AgendaContent(BaseModel):
    """Content for AGENDA layout (slide 3)."""
    agenda_items: List[str] = Field(default_factory=list, min_length=1)
    descriptions: List[str] = Field(default_factory=list)  # optional one-liners per item
    time_allocation: List[str] = Field(default_factory=list)  # optional duration per item


class BulletContent(BaseModel):
    """Content for TITLE_AND_CONTENT layout."""
    bullets: List[str] = Field(default_factory=list, min_length=1)
    body_text: Optional[str] = None  # fallback prose paragraph


class TwoColumnContent(BaseModel):
    """Content for TWO_COLUMN layout."""
    left_heading: Optional[str] = None
    left_bullets: List[str] = Field(default_factory=list)
    right_heading: Optional[str] = None
    right_bullets: List[str] = Field(default_factory=list)


class SectionHeaderContent(BaseModel):
    """Content for SECTION_HEADER layout."""
    section_title: str
    description: Optional[str] = None


class ComparisonContent(BaseModel):
    """Content for COMPARISON layout."""
    left_heading: str = "Option A"
    left_points: List[str] = Field(default_factory=list)
    right_heading: str = "Option B"
    right_points: List[str] = Field(default_factory=list)


class TimelineItem(BaseModel):
    date: str
    event: str
    description: Optional[str] = None


class TimelineContent(BaseModel):
    """Content for TIMELINE layout."""
    items: List[TimelineItem] = Field(default_factory=list, min_length=1)


class ProcessStep(BaseModel):
    step_number: int
    title: str
    description: Optional[str] = None


class ProcessContent(BaseModel):
    """Content for PROCESS layout."""
    steps: List[ProcessStep] = Field(default_factory=list, min_length=1)


class ArchitectureComponent(BaseModel):
    name: str
    description: Optional[str] = None
    sub_components: List[str] = Field(default_factory=list)


class ArchitectureContent(BaseModel):
    """Content for ARCHITECTURE layout."""
    layers: List[ArchitectureComponent] = Field(default_factory=list, min_length=1)
    description: Optional[str] = None


class Statistic(BaseModel):
    value: str         # e.g. "40%", "$2M", "10x"
    label: str         # e.g. "Reduction in processing time"
    context: Optional[str] = None


class StatisticsContent(BaseModel):
    """Content for STATISTICS layout."""
    statistics: List[Statistic] = Field(default_factory=list, min_length=1)
    context_paragraph: Optional[str] = None


class ConclusionContent(BaseModel):
    """Content for CONCLUSION layout."""
    summary_points: List[str] = Field(default_factory=list)
    recommendations: List[str] = Field(default_factory=list)  # NEW — client requirement
    call_to_action: Optional[str] = None
    next_steps: List[str] = Field(default_factory=list)


class ImageTextContent(BaseModel):
    """Content for IMAGE_TEXT layout (text portion only; image is a placeholder)."""
    description: str
    bullets: List[str] = Field(default_factory=list)
    image_caption: Optional[str] = None


# ---------------------------------------------------------------------------
# Union of all content types
# ---------------------------------------------------------------------------

SlideContentUnion = Union[
    TitleContent,
    ExecutiveSummaryContent,
    AgendaContent,
    BulletContent,
    TwoColumnContent,
    SectionHeaderContent,
    ComparisonContent,
    TimelineContent,
    ProcessContent,
    ArchitectureContent,
    StatisticsContent,
    ConclusionContent,
    ImageTextContent,
    Dict[str, Any],  # permissive fallback — handled in pptx_builder
]


# ---------------------------------------------------------------------------
# Slide definition
# ---------------------------------------------------------------------------

class SlideDefinition(BaseModel):
    """A single planned slide."""

    slide_number: int = Field(ge=1)
    purpose: str
    title: str
    layout_type: str = "TITLE_AND_CONTENT"  # keep as str; validated in layer manager
    content: Dict[str, Any] = Field(default_factory=dict)
    speaker_notes: Optional[str] = None

    @field_validator("layout_type", mode="before")
    @classmethod
    def normalise_layout(cls, v: Any) -> str:
        """Uppercase and strip the layout type for robust matching."""
        return str(v).upper().strip()

    @field_validator("title", mode="before")
    @classmethod
    def strip_title(cls, v: Any) -> str:
        return str(v).strip()


# ---------------------------------------------------------------------------
# Full presentation plan (Stage-2 output)
# ---------------------------------------------------------------------------

class PresentationPlan(BaseModel):
    """Complete presentation plan returned by the LLM."""

    title: str
    subtitle: Optional[str] = None
    slides: List[SlideDefinition] = Field(default_factory=list, min_length=1)

    @field_validator("slides", mode="after")
    @classmethod
    def sort_slides(cls, v: List[SlideDefinition]) -> List[SlideDefinition]:
        return sorted(v, key=lambda s: s.slide_number)

    @model_validator(mode="after")
    def renumber_slides(self) -> "PresentationPlan":
        """Ensure slide numbers are sequential starting from 1."""
        for i, slide in enumerate(self.slides, start=1):
            slide.slide_number = i
        return self


class PresentationPlanWrapper(BaseModel):
    """Root wrapper matching the JSON structure: { 'presentation': {...} }"""
    presentation: PresentationPlan


# ---------------------------------------------------------------------------
# Content analysis (Stage-1 output)
# ---------------------------------------------------------------------------

class ContentAnalysis(BaseModel):
    """Structured analysis of the source content."""

    main_topic: str
    key_concepts: List[str] = Field(default_factory=list)
    sections: List[str] = Field(default_factory=list)
    statistics: List[str] = Field(default_factory=list)
    processes: List[str] = Field(default_factory=list)
    comparisons: List[str] = Field(default_factory=list)
    timelines: List[str] = Field(default_factory=list)
    conclusions: List[str] = Field(default_factory=list)
    recommendations: List[str] = Field(default_factory=list)
    suggested_slide_count: int = Field(default=10, ge=3, le=30)
    content_types_detected: List[str] = Field(default_factory=list)
    summary: str = ""
    # NEW fields — used by structural enforcement layer
    executive_highlights: List[str] = Field(default_factory=list)  # 3-5 top takeaways for exec summary slide
    agenda_topics: List[str] = Field(default_factory=list)         # ordered section topics for agenda slide


# Fields the schema requires as List[str] — LLMs sometimes collapse a
# single-item list down to a bare string (e.g. "conclusions": "...").
CONTENT_ANALYSIS_LIST_FIELDS = (
    "key_concepts", "sections", "statistics", "processes", "comparisons",
    "timelines", "conclusions", "recommendations", "content_types_detected",
    "executive_highlights", "agenda_topics",
)


def normalize_content_analysis_lists(raw_data: dict) -> dict:
    """Coerce non-list values in ContentAnalysis's list-typed fields so validation doesn't crash."""
    for field in CONTENT_ANALYSIS_LIST_FIELDS:
        value = raw_data.get(field)
        if isinstance(value, str):
            raw_data[field] = [value] if value.strip() else []
        elif value is not None and not isinstance(value, list):
            raw_data[field] = [str(value)]
    return raw_data


# ---------------------------------------------------------------------------
# Validation result
# ---------------------------------------------------------------------------

class ValidationResult(BaseModel):
    is_valid: bool = True
    errors: List[str] = Field(default_factory=list)
    warnings: List[str] = Field(default_factory=list)

    def add_error(self, msg: str) -> None:
        self.errors.append(msg)
        self.is_valid = False

    def add_warning(self, msg: str) -> None:
        self.warnings.append(msg)


# ---------------------------------------------------------------------------
# Semantic chunking models (new pipeline)
# ---------------------------------------------------------------------------

class ChunkAnalysis(BaseModel):
    """
    LLM output for a single semantic chunk / document section.
    Stored in logs/llm/<run_id>/section_XXX.json.
    """
    chunk_id: str
    heading: str
    level: int = 1
    parent_heading: Optional[str] = None
    section_path: List[str] = Field(default_factory=list)

    # Core extractions
    summary: str = ""
    key_points: List[str] = Field(default_factory=list)
    facts: List[str] = Field(default_factory=list)
    statistics: List[str] = Field(default_factory=list)
    processes: List[str] = Field(default_factory=list)
    comparisons: List[str] = Field(default_factory=list)
    dates: List[str] = Field(default_factory=list)
    examples: List[str] = Field(default_factory=list)
    relationships: List[str] = Field(default_factory=list)
    potential_visuals: List[str] = Field(default_factory=list)
    content_types: List[str] = Field(default_factory=list)  # e.g. STATISTICS, PROCESS
    importance: str = "medium"  # high / medium / low


class DocumentSectionLog(BaseModel):
    """Lightweight representation of a section for JSON logging."""
    heading: str
    level: int
    subsections: List["DocumentSectionLog"] = Field(default_factory=list)


class DocumentStructureLog(BaseModel):
    """Saved to logs/llm/<run_id>/document_structure.json."""
    document_title: str
    detection_method: str  # docx_styles | txt_markdown | txt_numbered | pdf_heuristic | llm_inferred
    total_sections: int
    total_chunks: int
    sections: List[DocumentSectionLog] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# AI-driven dynamic chunking models
# ---------------------------------------------------------------------------

class ChunkBoundary(BaseModel):
    """
    A single chunk boundary returned by the LLM during dynamic chunking.

    The LLM references element IDs (e.g. 'element_012') rather than
    reproducing document content, keeping token usage low.
    """
    chunk_id: str                    # e.g. "chunk_001"
    start_element: str               # e.g. "element_001"
    end_element: str                 # e.g. "element_012"
    topic: str                       # brief topic label for this chunk
    boundary_reason: str             # why the boundary is here
    confidence: float = Field(ge=0.0, le=1.0, default=0.8)

    @field_validator("chunk_id", "start_element", "end_element", mode="before")
    @classmethod
    def strip_str(cls, v: Any) -> str:
        return str(v).strip()

    @field_validator("confidence", mode="before")
    @classmethod
    def clamp_confidence(cls, v: Any) -> float:
        try:
            return max(0.0, min(1.0, float(v)))
        except (TypeError, ValueError):
            return 0.8


class DynamicChunkingPlan(BaseModel):
    """
    LLM output for one sliding window of the document.
    Saved to logs/llm/<run_id>/dynamic_chunking/window_XXX.json.
    """
    window_id: str
    boundaries: List[ChunkBoundary] = Field(default_factory=list)


class DynamicChunkingWindowLog(BaseModel):
    """Full audit record for a single AI chunking window call."""
    run_id: str
    request_id: str
    window_id: str
    timestamp: str
    element_range: List[str]         # [start_element_id, end_element_id]
    element_count: int
    prompt_preview: str              # first 500 chars only — no keys
    raw_llm_response: str
    parsed_boundaries: List[dict] = Field(default_factory=list)
    validation_result: str           # "ok" | "gap_filled" | "overlap_resolved" | "fallback"
    retry_count: int = 0
    token_usage: Optional[dict] = None
    error: Optional[str] = None


class ChunkingSummaryLog(BaseModel):
    """Master summary saved to logs/llm/<run_id>/dynamic_chunking/chunking_summary.json."""
    run_id: str
    document_title: str
    detection_method: str
    total_elements: int
    total_windows: int
    total_chunks: int
    boundary_method_counts: dict     # {"ai_semantic": N, "paragraph_fallback": M}
    chunks: List[dict] = Field(default_factory=list)  # serialised SemanticChunk metadata

