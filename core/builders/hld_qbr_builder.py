"""
core/builders/hld_qbr_builder.py — Deep-clone builder for the HLD QBR Template
(UPS Healthcare Quarterly Business Review format).

Architecture (see HLD_QBR_TEMPLATE_PLAN.md for the full analysis):
  - Loads the 60-slide HLD QBR working copy (config.HLD_QBR_TEMPLATE_FILE).
  - OpenXML deep-clones only the specific CORE archetype slides needed for
    THIS plan (Cover/Agenda/Closing always; every other section only if the
    plan actually has content for it — graceful degradation, no filler).
  - Strips, on every cloned slide: guidance/instruction boxes (regex),
    decorative status-dot icons left over from the source slide's original
    column meaning, and stray text-highlight "replace me" markers.
  - Prunes the original 60 template slides, leaving only the generated ones.
  - Returns binary bytes via io.BytesIO().
"""
from __future__ import annotations

import copy
from datetime import datetime
import io
import math
import re
from typing import Any, Dict, List, Optional

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.parts.chart import ChartPart
from pptx.util import Inches, Pt

import config
from llm.hld_qbr_schemas import HLDQBRPresentationPlan
from utils.logging_utils import get_logger

logger = get_logger(__name__)

def _get_ordinal_date(dt: Optional[datetime] = None) -> str:
    """Returns today's date formatted with ordinal suffix matching template style, e.g. 'September 21st 2026'."""
    if dt is None:
        dt = datetime.now()
    day = dt.day
    if 11 <= (day % 100) <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{dt.strftime('%B')} {day}{suffix} {dt.year}"

IDX_COVER = 0
IDX_AGENDA = 1
IDX_EXECUTIVE_SUMMARY = 3
IDX_ORG_STRUCTURE = 2
IDX_ACHIEVEMENTS = 4
IDX_PRIORITIES = 5
IDX_SECTION_PERF_MGMT = 11
IDX_TRACKER = 9
IDX_OPERATIONAL_CHART = 13
IDX_KPI_DASHBOARD = 12
IDX_VOICE_OF_CUSTOMER = 10
IDX_BLANK_CANVAS = 20  # Template slide 21: 'MASTER DATA MANAGEMENT KPI\'S AND UPDATES' (Title Only layout, white background)
IDX_SECTION_CIP = 21   # Template slide 22: 'CONTINUOUS IMPROVEMENT PROGRAM UPDATES' (1_Section Header_No Image)
IDX_GEMBA_WALK = 22
IDX_CI_TRACKER = 23
IDX_SECTION_QUALITY = 24
IDX_QUALITY_ORG = 25
IDX_NC_REVIEW = 26
IDX_NC_TRACKER = 27
IDX_NEXT_STEPS = 29
IDX_CLOSING = 30

GUIDANCE_REGEX = re.compile(
    r"required content|required slide|\bexample\b|optional \||format option|"
    r"additional slides available|must be updated|formulas in notes",
    re.IGNORECASE,
)

# Official UPS Healthcare Template Theme Palette (accent1, accent2, accent4, accent5, accent6)
TEMPLATE_SERIES_COLORS = [
    RGBColor(14, 37, 84),    # accent1: Deep Navy (#0E2554)
    RGBColor(66, 109, 169),  # accent2: Slate Blue (#426DA9)
    RGBColor(255, 190, 0),   # accent4: UPS Gold (#FFBE00)
    RGBColor(0, 133, 125),   # accent5: UPS Teal (#00857D)
    RGBColor(136, 167, 209), # accent6: Light Blue (#88A7D1)
]



def _clone_slide(prs: Presentation, source_idx: int) -> Any:
    """Deep-clones a slide, remapping media relationships and stripping template
    authoring artifacts (guidance boxes, decorative dots, highlight markers)."""
    source_slide = prs.slides[source_idx]
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)

    if source_slide.background and source_slide.background.fill:
        try:
            if source_slide.background.fill.type == 1:  # SOLID
                new_slide.background.fill.solid()
                new_slide.background.fill.fore_color.rgb = source_slide.background.fill.fore_color.rgb
        except Exception:
            pass

    rId_map: Dict[str, str] = {}
    pkg = prs.part.package
    for rId, rel in source_slide.part.rels.items():
        if (
            "slideLayout" not in rel.target_ref
            and "notesSlide" not in rel.target_ref
            and "notesMaster" not in rel.target_ref
        ):
            try:
                if isinstance(rel.target_part, ChartPart):
                    # Decouple chart part: clone independently so slides never share chart parts or excel workbooks
                    source_chart_part = rel.target_part
                    new_chart_partname = pkg.next_partname("/ppt/charts/chart%d.xml")
                    new_chart_part = ChartPart.load(
                        new_chart_partname,
                        source_chart_part.content_type,
                        pkg,
                        source_chart_part.blob,
                    )
                    new_chart_part._element._remove_externalData()
                    for c_rId, c_rel in source_chart_part.rels.items():
                        if "chartStyle" in c_rel.reltype or "chartColorStyle" in c_rel.reltype:
                            new_chart_part.relate_to(c_rel.target_part, c_rel.reltype)
                    new_rId = new_slide.part.relate_to(new_chart_part, RT.CHART)
                    rId_map[rId] = new_rId
                else:
                    new_rId = new_slide.part.relate_to(rel.target_part, rel.reltype)
                    rId_map[rId] = new_rId
            except Exception:
                pass

    for s in list(new_slide.shapes):
        s.element.getparent().remove(s.element)

    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        for elem in new_el.xpath(".//*[@r:embed or @r:link or @r:id]"):
            for attr in list(elem.attrib.keys()):
                if attr.endswith("embed") or attr.endswith("link") or attr.endswith("}id") or attr == "r:id":
                    old_val = elem.attrib[attr]
                    if old_val in rId_map:
                        elem.attrib[attr] = rId_map[old_val]
        new_slide.shapes._spTree.append(new_el)

    _strip_all_highlights(new_slide)
    return new_slide


def _strip_all_highlights(slide: Any) -> int:
    """Removes <a:highlight> marks (e.g. the template's magenta '[CUSTOMER NAME]'
    replace-me flag) slide-wide, including inside table cells."""
    removed = 0
    text_frames = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            text_frames.append(shape.text_frame)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    text_frames.append(cell.text_frame)
    for tf in text_frames:
        for para in tf.paragraphs:
            for run in para.runs:
                rPr = run._r.find(qn("a:rPr"))
                if rPr is None:
                    continue
                hl = rPr.find(qn("a:highlight"))
                if hl is not None:
                    rPr.remove(hl)
                    removed += 1
    return removed


def _strip_guidance_shapes(slide: Any) -> int:
    """Removes any shape whose text matches the guidance-box vocabulary."""
    removed = 0
    for shape in list(slide.shapes):
        if shape.has_text_frame and GUIDANCE_REGEX.search(shape.text_frame.text):
            shape.element.getparent().remove(shape.element)
            removed += 1
    return removed


def _strip_decorative_connectors(slide: Any) -> int:
    """Removes leftover red/yellow/green 'Flowchart: Connector' status-dot icons
    tied to one specific column's meaning in the ORIGINAL template slide."""
    removed = 0
    for shape in list(slide.shapes):
        if shape.name.startswith("Flowchart: Connector"):
            shape.element.getparent().remove(shape.element)
            removed += 1
    return removed


def _render_stat_highlights(slide: Any, stats: List[Any]) -> None:
    """Composes N hero-metric badges (icon circle + big number + label) on a blank
    canvas — mined from the template's unused 'Performance Summary' slide (34)."""
    n = len(stats)
    if n == 0:
        return
    LEFT_MARGIN = Inches(0.40)
    TOTAL_WIDTH = Inches(12.53)
    GAP = Inches(0.30)
    CARD_W = int((TOTAL_WIDTH - GAP * (n - 1)) / n)
    TOP = Inches(2.30)
    BADGE_D = Inches(0.95)

    for i, stat in enumerate(stats):
        left = int(LEFT_MARGIN + i * (CARD_W + GAP))

        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, left + (CARD_W - BADGE_D) // 2, TOP, BADGE_D, BADGE_D)
        badge.fill.solid()
        badge.fill.fore_color.rgb = RGBColor(0, 43, 73)
        badge.line.fill.background()
        tf_b = badge.text_frame
        tf_b.word_wrap = True
        p_b = tf_b.paragraphs[0]
        p_b.text = str(i + 1)
        p_b.alignment = PP_ALIGN.CENTER
        for r in p_b.runs:
            r.font.name = "Verdana"
            r.font.size = Pt(16)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 190, 0)

        value_box = slide.shapes.add_textbox(left, TOP + BADGE_D + Inches(0.15), CARD_W, Inches(0.60))
        tf_v = value_box.text_frame
        tf_v.word_wrap = True
        p_v = tf_v.paragraphs[0]
        p_v.text = stat.value
        p_v.alignment = PP_ALIGN.CENTER
        for r in p_v.runs:
            r.font.name = "Verdana"
            r.font.size = Pt(22 if len(stat.value) <= 8 else 16)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0, 43, 73)

        label_box = slide.shapes.add_textbox(left, TOP + BADGE_D + Inches(0.75), CARD_W, Inches(0.90))
        tf_l = label_box.text_frame
        tf_l.word_wrap = True
        p_l = tf_l.paragraphs[0]
        p_l.text = stat.label
        p_l.alignment = PP_ALIGN.CENTER
        for r in p_l.runs:
            r.font.name = "Verdana"
            r.font.size = Pt(11)
            r.font.color.rgb = RGBColor(88, 107, 123)


def _render_process_flow(slide: Any, steps: List[Any]) -> None:
    """Composes N left-to-right overlapping chevron segments on a blank canvas —
    mined from the template's unused 'Process Flow Comparison' slide (50)."""
    n = len(steps)
    if n == 0:
        return
    LEFT_MARGIN = Inches(0.40)
    TOTAL_WIDTH = Inches(12.53)
    OVERLAP = Inches(0.18)  # slight overlap reads as a connected flow, not separate boxes
    SEG_W = int((TOTAL_WIDTH + OVERLAP * (n - 1)) / n)
    TOP = Inches(2.60)
    SEG_H = Inches(1.10)

    for i, step in enumerate(steps):
        left = int(LEFT_MARGIN + i * (SEG_W - OVERLAP))
        chevron = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, left, TOP, SEG_W, SEG_H)
        chevron.fill.solid()
        chevron.fill.fore_color.rgb = TEMPLATE_SERIES_COLORS[i % len(TEMPLATE_SERIES_COLORS)]
        chevron.line.fill.background()
        tf = chevron.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.15)
        tf.margin_right = Inches(0.25)
        p_h = tf.paragraphs[0]
        p_h.text = step.heading.upper()
        for r in p_h.runs:
            r.font.name = "Verdana"
            r.font.size = Pt(12)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
        if step.description:
            p_d = tf.add_paragraph()
            p_d.text = step.description
            for r in p_d.runs:
                r.font.name = "Verdana"
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(255, 255, 255)


def _shape_by_name(slide: Any, name: str) -> Optional[Any]:
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def _remove_shapes(slide: Any, predicate) -> int:
    """Removes every shape matching predicate(shape) -> bool. Used to clear a
    cloned slide's fixed decorative/placeholder shapes before freeform rendering."""
    removed = 0
    for shape in list(slide.shapes):
        if predicate(shape):
            shape.element.getparent().remove(shape.element)
            removed += 1
    return removed


def _set_first_run_text(shape: Any, text: str) -> None:
    """Sets paragraph-0 text in-place, preserving existing run formatting."""
    tf = shape.text_frame
    if not tf.paragraphs:
        return
    p0 = tf.paragraphs[0]
    if p0.runs:
        p0.runs[0].text = text
        for r in p0.runs[1:]:
            r.text = ""
    else:
        p0.text = text


def _set_slide_header_and_sub(
    slide: Any,
    title_text: str,
    subtitle_text: Optional[str] = None,
) -> None:
    """Sets the slide title and subtitle using the template's standard placeholders
    (Title 2 and Text Placeholder 3) or creates them if missing.
    Ensures empty placeholders are removed so PowerPoint does not show prompt text like 'Sub-header'.
    """
    title_shape = None
    sub_shape = None

    for sh in list(slide.shapes):
        if sh.name == "Title 2" or (sh.is_placeholder and sh.placeholder_format.type == 1):
            title_shape = sh
        elif sh.name == "Text Placeholder 3" or (sh.is_placeholder and sh.placeholder_format.idx in (3, 13)):
            sub_shape = sh
        elif sh.has_text_frame and "MASTER DATA" in sh.text_frame.text:
            title_shape = sh

    # 1. Title formatting (matches template slide 14, 15, 21: left=0.40", top=0.40", width=12.53", height=0.42")
    if title_shape is None:
        title_shape = slide.shapes.add_textbox(
            Inches(0.40), Inches(0.40), Inches(12.53), Inches(0.42)
        )

    tf_t = title_shape.text_frame
    tf_t.word_wrap = False
    p_t = tf_t.paragraphs[0]
    p_t.text = (title_text or "").upper()
    runs_t = p_t.runs if p_t.runs else [p_t.add_run()]
    for r in runs_t:
        r.font.name = "Verdana"
        r.font.size = Pt(20)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 112, 192)  # UPS brand blue (#0070C0)
    while len(tf_t.paragraphs) > 1:
        p_extra = tf_t.paragraphs[-1]
        p_extra._p.getparent().remove(p_extra._p)

    # 2. Subtitle formatting (matches template slide 14, 15, 21: left=0.40", top=0.83", width=12.53", height=0.24")
    if subtitle_text and subtitle_text.strip():
        if sub_shape is None:
            sub_shape = slide.shapes.add_textbox(
                Inches(0.40), Inches(0.83), Inches(12.53), Inches(0.24)
            )
        tf_s = sub_shape.text_frame
        tf_s.word_wrap = False
        p_s = tf_s.paragraphs[0]
        p_s.text = subtitle_text.strip()
        runs_s = p_s.runs if p_s.runs else [p_s.add_run()]
        for r in runs_s:
            r.font.name = "Verdana"
            r.font.size = Pt(11)
            r.font.bold = False
            r.font.color.rgb = RGBColor(88, 107, 123)  # Template Slate/Dark Gray (#586B7B)
        while len(tf_s.paragraphs) > 1:
            p_extra = tf_s.paragraphs[-1]
            p_extra._p.getparent().remove(p_extra._p)
    else:
        # Crucial: if subtitle is empty, remove the placeholder so PowerPoint does NOT show "Sub-header"
        if sub_shape is not None:
            try:
                slide.shapes._spTree.remove(sub_shape._element)
            except Exception:
                pass


def _set_cover_breadcrumb_blank(cover_slide: Any) -> None:
    """Blanks the center banner text (Rectangle 6) — the title lives in the
    right-corner Title 5 box; a duplicate title crammed into the narrow center
    zone reads as unprofessional (see HLD_QBR_TEMPLATE_PLAN.md follow-ups)."""
    shape = _shape_by_name(cover_slide, "Rectangle 6")
    if shape is None:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            run.text = ""


def _set_cover_title(cover_slide: Any, text: str) -> None:
    """Fills Title 5 (right-corner box). Ships EMPTY in the template (only an
    <a:endParaRPr>, no run) at Verdana 22pt Bold; longer text must shrink."""
    shape = _shape_by_name(cover_slide, "Title 5")
    if shape is None:
        return
    tf = shape.text_frame
    if not tf.paragraphs:
        return
    p0 = tf.paragraphs[0]
    if p0.runs:
        p0.runs[0].text = text
        for r in p0.runs[1:]:
            r.text = ""
        run = p0.runs[0]
    else:
        run = p0.add_run()
        run.text = text
    if len(text) > 24:
        run.font.size = Pt(14)
    elif len(text) > 16:
        run.font.size = Pt(18)


def _set_org_box(shape: Any, name: str, role: str) -> None:
    """Org-chart person card: paragraph 0 = name, paragraph 1 = role."""
    paras = shape.text_frame.paragraphs
    for i, val in enumerate((name, role)):
        if i < len(paras) and paras[i].runs:
            paras[i].runs[0].text = val
            for r in paras[i].runs[1:]:
                r.text = ""


def _set_or_remove_facility_placeholder(slide: Any, facility_name: str) -> None:
    """Fills Text Placeholder 3 with the facility name if provided; otherwise
    removes the placeholder element entirely so PowerPoint doesn't show ghost
    prompts like 'Sub-header' or 'Facility name or location'."""
    ph = _shape_by_name(slide, "Text Placeholder 3")
    if ph is not None:
        if facility_name and facility_name.strip():
            _set_first_run_text(ph, facility_name.strip())
        else:
            slide.shapes._spTree.remove(ph._element)


def _fill_priority_group(
    group_shape: Any,
    heading: str,
    body: str,
    center_x: Optional[int] = None,
    card_width: Optional[int] = None,
    target_body_top: int = 4220000,
) -> None:
    # 1. Update group bounds and child shape positions if center_x and card_width are provided
    if center_x is not None and card_width is not None:
        new_left = center_x - (card_width // 2)
        grp_xfrm = group_shape._element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
        if grp_xfrm is not None:
            off = grp_xfrm.find(qn("a:off"))
            ext = grp_xfrm.find(qn("a:ext"))
            chOff = grp_xfrm.find(qn("a:chOff"))
            chExt = grp_xfrm.find(qn("a:chExt"))
            if off is not None:
                off.set("x", str(new_left))
            if ext is not None:
                ext.set("cx", str(card_width))
            if chOff is not None:
                chOff.set("x", str(new_left))
            if chExt is not None:
                chExt.set("cx", str(card_width))

        # Center the circle and icon inside the group
        for s in group_shape.shapes:
            if s.name.startswith("Oval") or s.name.startswith("Freeform"):
                s.left = center_x - (s.width // 2)

    # 2. Identify heading & body text boxes
    text_boxes = [
        s for s in group_shape.shapes
        if getattr(s, "has_text_frame", False) and (s.shape_type == 17 or "Shape;" in s.name)
    ]
    text_boxes.sort(key=lambda s: s.top)
    if len(text_boxes) >= 2:
        heading_sub = text_boxes[0]
        body_sub = text_boxes[1]
    else:
        heading_sub = None
        body_sub = None
        for sub in group_shape.shapes:
            if not getattr(sub, "has_text_frame", False):
                continue
            current = sub.text_frame.text.strip().upper()
            if current.startswith("PRIORITY"):
                heading_sub = sub
            elif "SUPPORTING" in current:
                body_sub = sub

    if heading_sub and body_sub:
        # Widen heading & body boxes to card_width if provided, or match body width
        if center_x is not None and card_width is not None:
            new_left = center_x - (card_width // 2)
            heading_sub.left = new_left
            heading_sub.width = card_width
            body_sub.left = new_left
            body_sub.width = card_width
        else:
            heading_sub.left = body_sub.left
            heading_sub.width = body_sub.width

        # Heading formatting: STRICT 14PT BOLD across ALL cards (per template guideline)
        heading_sub.text_frame.word_wrap = True
        p_head = heading_sub.text_frame.paragraphs[0]
        p_head.text = heading

        # Strip 150% line spacing from the template paragraph so it renders single-spaced
        pPr = p_head._p.find(qn("a:pPr"))
        if pPr is not None:
            lnSpc = pPr.find(qn("a:lnSpc"))
            if lnSpc is not None:
                pPr.remove(lnSpc)

        if p_head.runs:
            run_h = p_head.runs[0]
            run_h.font.bold = True
            run_h.font.size = Pt(14)  # STRICT 14PT BOLD FOR ALL CARDS PER TEMPLATE GUIDELINE

        # Body formatting: dynamically auto-scale if text volume is high
        body_sub.text_frame.word_wrap = True
        p_body = body_sub.text_frame.paragraphs[0]
        p_body.text = body
        if p_body.runs:
            run_b = p_body.runs[0]
            if len(body) > 220 or target_body_top >= 4450000:
                run_b.font.size = Pt(10.5)
            elif len(body) > 160:
                run_b.font.size = Pt(11.0)
            else:
                run_b.font.size = Pt(12.0)

        # Set body baseline top coordinate to ensure clean spacing and alignment
        body_sub.top = target_body_top
    elif heading_sub:
        _set_first_run_text(heading_sub, heading)
    elif body_sub:
        _set_first_run_text(body_sub, body)



def _apply_table_status_dots(
    slide: Any,
    table_shape: Any,
    status_column_idx: int,
    action_items: List[Any],
) -> None:
    """Renders corporate status indicator dots in the table's status column,
    preserving the template's header legend dots (Green, Amber, Red) and
    dynamically placing centered colored circle dots for each populated row."""
    tbl = table_shape.table
    header_bottom = table_shape.top + tbl.rows[0].height

    # 1. Strip leftover dummy template row connectors (below the header row)
    for s in list(slide.shapes):
        if s.name.startswith("Flowchart: Connector") and s.top >= (header_bottom - Inches(0.04)):
            try:
                slide.shapes._spTree.remove(s._element)
            except Exception:
                pass

    # 2. Horizontal centering in the status column
    col_left = table_shape.left + sum(tbl.columns[c].width for c in range(status_column_idx))
    col_width = tbl.columns[status_column_idx].width
    dot_size = Inches(0.20)  # Standard 182,880 EMU matching template
    dot_left = int(col_left + (col_width - dot_size) // 2)

    # 3. Add dynamic colored status dots for each populated item
    for r_idx, item in enumerate(action_items, start=1):
        if r_idx >= len(tbl.rows):
            break

        # Blank cell text so dot sits clean
        cell = tbl.cell(r_idx, status_column_idx)
        cell.text = ""

        raw_status = (getattr(item, "status", None) or "").strip().lower()
        if any(w in raw_status for w in ("complete", "done", "closed", "on track", "green")):
            color = RGBColor(0, 168, 89)   # Emerald Green
        elif any(w in raw_status for w in ("delay", "block", "overdue", "escalat", "critical", "red", "missed")):
            color = RGBColor(220, 20, 60)   # Crimson Red
        else:
            color = RGBColor(255, 192, 0)   # Amber / Yellow (In Progress, At Risk, Pending)

        row_top = table_shape.top + sum(tbl.rows[k].height for k in range(r_idx))
        dot_top = int(row_top + (tbl.rows[r_idx].height - dot_size) // 2)

        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, dot_left, dot_top, dot_size, dot_size)
        dot.fill.solid()
        dot.fill.fore_color.rgb = color
        dot.line.color.rgb = RGBColor(255, 255, 255)
        dot.line.width = Pt(1.2)


def _fill_table_rows(
    tbl: Any,
    rows: List[List[str]],
    start_row: int = 1,
    prune_unused_rows: bool = True,
    text_color: RGBColor = RGBColor(0, 43, 73),
) -> None:
    """Fills table rows with dynamic typography, cell margins, font color, and row pruning."""
    if not rows:
        return
    max_cell_len = max(len(str(val)) for row in rows for val in row) if rows else 0
    if max_cell_len > 100:
        font_sz = Pt(9.0)
    elif max_cell_len > 60:
        font_sz = Pt(10.0)
    elif max_cell_len > 35:
        font_sz = Pt(10.5)
    else:
        font_sz = Pt(11.0)

    for r_offset, row_values in enumerate(rows):
        r_idx = start_row + r_offset
        if r_idx >= len(tbl.rows):
            break
        for c_idx, val in enumerate(row_values):
            if c_idx < len(tbl.columns):
                cell = tbl.cell(r_idx, c_idx)
                cell.text = str(val)
                cell.margin_top = Inches(0.04)
                cell.margin_bottom = Inches(0.04)
                cell.margin_left = Inches(0.08)
                cell.margin_right = Inches(0.08)
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.name = "Verdana"
                        run.font.size = font_sz
                        run.font.color.rgb = text_color

    used_through = start_row + len(rows)
    if prune_unused_rows:
        while len(tbl.rows) > used_through:
            last_tr = tbl.rows[len(tbl.rows) - 1]._tr
            tbl._tbl.remove(last_tr)
    else:
        for r_idx in range(used_through, len(tbl.rows)):
            for c_idx in range(len(tbl.columns)):
                tbl.cell(r_idx, c_idx).text = ""


def _get_chart_list(plan: HLDQBRPresentationPlan) -> List[Any]:
    return list(plan.charts) if plan.charts else ([plan.operational_chart] if plan.operational_chart else [])


def _get_table_list(plan: HLDQBRPresentationPlan) -> List[Any]:
    return list(plan.kpi_tables) if plan.kpi_tables else []


# ---------------------------------------------------------------------------
# Dispatch functions — one per orderable archetype/divider. Each is a pure
# extraction of what used to be an inline block inside build(), unchanged in
# behavior, so build() can call them in WHATEVER sequence the LLM's narrative
# order specifies instead of a hardcoded chronology. See _resolve_archetype_order.
# ---------------------------------------------------------------------------

def _dispatch_org_structure(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not plan.org_structure:
        return
    org_slide = _clone_slide(prs, IDX_ORG_STRUCTURE)
    org_title = _shape_by_name(org_slide, "Title 1")
    if org_title is not None:
        _set_first_run_text(org_title, (plan.org_structure_title or "ACCOUNT TEAM & KEY CONTACTS").strip().upper())
    org_boxes = sorted(
        (s for s in org_slide.shapes if s.name.startswith("Rectangle: Rounded Corners")),
        key=lambda s: (s.top, s.left),
    )
    for box, person in zip(org_boxes, plan.org_structure):
        if person.name:
            _set_org_box(box, person.name, person.role)
    for box in org_boxes[len(plan.org_structure):]:
        _set_org_box(box, "", "")


def _dispatch_achievements(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not plan.achievements:
        return
    achievements_slide = _clone_slide(prs, IDX_ACHIEVEMENTS)
    _strip_guidance_shapes(achievements_slide)
    _set_or_remove_facility_placeholder(achievements_slide, plan.facility_name)

    ach_title = _shape_by_name(achievements_slide, "Title 2")
    if ach_title is None:
        ach_title = _shape_by_name(achievements_slide, "Title 6")
    if ach_title is not None and ach_title.text_frame.paragraphs:
        heading_text = (
            plan.achievements_title.strip().upper()
            if plan.achievements_title
            else "PRIOR QUARTER MILESTONES & WINS"
        )
        p0 = ach_title.text_frame.paragraphs[0]
        p0.text = heading_text
        if p0.runs:
            p0.runs[0].font.size = Pt(22)
            p0.runs[0].font.bold = True

    milestone_boxes = sorted(
        (s for s in achievements_slide.shapes if s.name == "Content Placeholder 42"),
        key=lambda s: s.top,
    )
    connectors = sorted(
        (s for s in achievements_slide.shapes if s.name.startswith("Flowchart: Connector")),
        key=lambda s: s.top,
    )

    n_ach = len(plan.achievements)
    max_ach_len = max(len(t) for t in plan.achievements) if plan.achievements else 0

    TOP_MIN = 1200000  # ~1.31 inches
    TOP_MAX = 5400000  # ~5.91 inches
    spacing = (TOP_MAX - TOP_MIN) // (n_ach - 1) if n_ach > 1 else 0

    if max_ach_len > 140 or n_ach >= 5:
        ach_font_sz = Pt(11.0)
    elif max_ach_len > 80:
        ach_font_sz = Pt(12.5)
    else:
        ach_font_sz = Pt(13.5)

    for i, (text, box, conn) in enumerate(zip(plan.achievements, milestone_boxes, connectors)):
        new_top = TOP_MIN + i * spacing if n_ach > 1 else (TOP_MIN + TOP_MAX) // 2
        conn.top = int(new_top)
        box.top = int(new_top)
        box.height = int(Inches(0.65))
        box.text_frame.word_wrap = True

        p = box.text_frame.paragraphs[0]
        p.text = text
        pPr = p._p.find(qn("a:pPr"))
        if pPr is not None:
            lnSpc = pPr.find(qn("a:lnSpc"))
            if lnSpc is not None:
                pPr.remove(lnSpc)
        for r in p.runs:
            r.font.name = "Verdana"
            r.font.size = ach_font_sz
            r.font.color.rgb = RGBColor(0, 43, 73)

        p_c = conn.text_frame.paragraphs[0]
        p_c.text = str(i + 1)
        for r in p_c.runs:
            r.font.name = "Verdana"
            r.font.bold = True
            r.font.size = Pt(12)

    for box in milestone_boxes[n_ach:]:
        achievements_slide.shapes._spTree.remove(box._element)
    for conn in connectors[n_ach:]:
        achievements_slide.shapes._spTree.remove(conn._element)


def _dispatch_priorities(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not plan.priorities:
        return
    priorities_slide = _clone_slide(prs, IDX_PRIORITIES)
    _strip_guidance_shapes(priorities_slide)
    _set_or_remove_facility_placeholder(priorities_slide, plan.facility_name)

    priority_title = _shape_by_name(priorities_slide, "Title 2")
    if priority_title is not None and priority_title.text_frame.paragraphs:
        p0 = priority_title.text_frame.paragraphs[0]
        heading_text = (
            plan.priorities_title.strip().upper()
            if plan.priorities_title
            else (
                f"{plan.facility_name.upper()} STRATEGIC PRIORITIES"
                if plan.facility_name
                else "STRATEGIC PRIORITIES & FOCUS AREAS"
            )
        )
        p0.text = heading_text
        if p0.runs:
            p0.runs[0].font.size = Pt(22)
            p0.runs[0].font.bold = True

    n_p = len(plan.priorities)
    SLIDE_WIDTH = 12192000  # 13.333 inches

    if n_p == 2:
        CARD_WIDTH_EMU = int(4.00 * 914400)
        CIRCLE_CENTERS = [int(SLIDE_WIDTH * 0.30), int(SLIDE_WIDTH * 0.70)]
    elif n_p == 1:
        CARD_WIDTH_EMU = int(5.50 * 914400)
        CIRCLE_CENTERS = [int(SLIDE_WIDTH * 0.50)]
    else:
        CARD_WIDTH_EMU = int(3.25 * 914400)
        CIRCLE_CENTERS = [2031252, 5987374, 10056673]

    p_width_in = (CARD_WIDTH_EMU - int(0.20 * 914400)) / 914400
    chars_per_line = int(p_width_in * 9.5)
    max_head_lines = max(
        max(1, math.ceil(len(p.heading.strip()) / chars_per_line))
        for p in plan.priorities
    ) if plan.priorities else 1

    if max_head_lines >= 3:
        target_body_top = 4450000
    elif max_head_lines == 2:
        target_body_top = 4220000
    else:
        target_body_top = 3939696

    priority_groups = sorted(
        (s for s in priorities_slide.shapes if s.shape_type == 6),
        key=lambda s: s.left,
    )
    for col_idx, (group, item) in enumerate(zip(priority_groups, plan.priorities)):
        center_x = (
            CIRCLE_CENTERS[col_idx]
            if col_idx < len(CIRCLE_CENTERS)
            else group.left + (group.width // 2)
        )
        _fill_priority_group(
            group,
            item.heading,
            item.body,
            center_x=center_x,
            card_width=CARD_WIDTH_EMU,
            target_body_top=target_body_top,
        )
    for group in priority_groups[len(plan.priorities):]:
        priorities_slide.shapes._spTree.remove(group._element)


def _dispatch_divider_performance(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    """Only renders if the sections it divides actually have content — a
    divider with nothing to divide is never shown, regardless of order."""
    chart_list = _get_chart_list(plan)
    table_list = _get_table_list(plan)
    has_perf_section = bool(
        plan.action_tracker
        or table_list
        or plan.kpi_safety_quality
        or plan.kpi_operational
        or any(c and c.categories and c.series for c in chart_list)
    )
    if not has_perf_section:
        return
    perf_divider = _clone_slide(prs, IDX_SECTION_PERF_MGMT)
    t = _shape_by_name(perf_divider, "Title 6")
    if t is not None:
        heading = (
            plan.section_heading.strip().upper()
            if plan.section_heading
            else "OPERATIONAL PERFORMANCE & DATA REVIEW"
        )
        _set_first_run_text(t, heading)


def _dispatch_action_tracker(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not plan.action_tracker:
        return
    tracker = _clone_slide(prs, IDX_TRACKER)
    _strip_guidance_shapes(tracker)
    _set_or_remove_facility_placeholder(tracker, plan.facility_name)

    tracker_title = _shape_by_name(tracker, "Title 2")
    if tracker_title is None:
        tracker_title = _shape_by_name(tracker, "Title 6")
    if tracker_title is None:
        tracker_title = _shape_by_name(tracker, "Title 1")
    if tracker_title is not None and tracker_title.text_frame.paragraphs:
        heading_text = (
            plan.action_tracker_title.strip().upper()
            if plan.action_tracker_title
            else "OPEN ACTION ITEMS & ACCOUNTABILITY"
        )
        p0 = tracker_title.text_frame.paragraphs[0]
        p0.text = heading_text
        if p0.runs:
            p0.runs[0].font.size = Pt(22)
            p0.runs[0].font.bold = True

    table_shape = _shape_by_name(tracker, "Table 4")
    if table_shape is not None and table_shape.has_table:
        rows = [[r.project, r.owner, r.next_step, r.comment, ""] for r in plan.action_tracker]
        _fill_table_rows(table_shape.table, rows, start_row=1, prune_unused_rows=False)
        _apply_table_status_dots(tracker, table_shape, status_column_idx=4, action_items=plan.action_tracker)


def _dispatch_charts(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    """Renders one slide per genuinely chartable series (0-4 instances)."""
    chart_list = _get_chart_list(plan)
    for chart_model in chart_list:
        if not (chart_model and chart_model.categories and chart_model.series):
            continue
        chart_slide = _clone_slide(prs, IDX_BLANK_CANVAS)
        _strip_guidance_shapes(chart_slide)
        _strip_decorative_connectors(chart_slide)

        chart_title = (chart_model.chart_title or "OPERATIONAL PERFORMANCE SNAPSHOT").upper()
        _set_slide_header_and_sub(chart_slide, chart_title, plan.facility_name)

        chart_shape = next((s for s in chart_slide.shapes if s.has_chart), None)

        if chart_shape is None:
            from pptx.enum.chart import XL_CHART_TYPE
            chart_data_new = CategoryChartData()
            chart_data_new.categories = chart_model.categories
            for ser in chart_model.series:
                chart_data_new.add_series(ser.name, ser.values)

            has_insights = bool(chart_model.insights)
            if has_insights:
                chart_left = Inches(0.40)
                chart_top = Inches(1.30)
                chart_width = Inches(8.50)
                chart_height = Inches(5.45)
            else:
                chart_left = Inches(0.40)
                chart_top = Inches(1.30)
                chart_width = Inches(12.50)
                chart_height = Inches(5.45)

            chart_shape = chart_slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED,
                chart_left, chart_top, chart_width, chart_height,
                chart_data_new,
            )
            c = chart_shape.chart
            c.has_title = False
            c.has_legend = len(chart_model.series) > 1
            if c.has_legend:
                c.legend.position = XL_LEGEND_POSITION.TOP
                c.legend.include_in_layout = False
                try:
                    c.legend.font.name = "Verdana"
                    c.legend.font.size = Pt(10)
                except Exception:
                    pass
            try:
                c.category_axis.tick_labels.font.name = "Verdana"
                c.category_axis.tick_labels.font.size = Pt(9)
                c.value_axis.tick_labels.font.name = "Verdana"
                c.value_axis.tick_labels.font.size = Pt(9)
            except Exception:
                pass
            try:
                if c.plots:
                    plot = c.plots[0]
                    plot.vary_by_categories = False
                    plot.has_data_labels = True
                    plot.data_labels.font.name = "Verdana"
                    plot.data_labels.font.size = Pt(8)
                    for s_idx, ser in enumerate(plot.series):
                        color = TEMPLATE_SERIES_COLORS[s_idx % len(TEMPLATE_SERIES_COLORS)]
                        try:
                            fill = ser.format.fill
                            fill.solid()
                            fill.fore_color.rgb = color
                        except Exception:
                            pass
            except Exception:
                pass

            if has_insights:
                insights = chart_model.insights
                n_items = len(insights)
                insights_title_text = (getattr(chart_model, "insights_title", None) or "KEY OBSERVATIONS").strip().upper()

                CARD_LEFT = Inches(9.10)
                CARD_TOP = Inches(1.30)
                CARD_WIDTH = Inches(3.80)
                CARD_HEIGHT = Inches(5.45)

                blue_card = chart_slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    CARD_LEFT, CARD_TOP, CARD_WIDTH, CARD_HEIGHT
                )
                blue_card.fill.solid()
                blue_card.fill.fore_color.rgb = RGBColor(0, 43, 73)
                blue_card.line.fill.background()

                header_shape = chart_slide.shapes.add_textbox(
                    CARD_LEFT + Inches(0.20), CARD_TOP + Inches(0.20),
                    CARD_WIDTH - Inches(0.40), Inches(0.40)
                )
                tf_h = header_shape.text_frame
                tf_h.margin_left = tf_h.margin_right = tf_h.margin_top = tf_h.margin_bottom = 0
                p_h = tf_h.paragraphs[0]
                p_h.text = insights_title_text
                for r in p_h.runs:
                    r.font.name = "Verdana"
                    r.font.size = Pt(11.5)
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(255, 190, 0)

                accent_line = chart_slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    CARD_LEFT + Inches(0.20), CARD_TOP + Inches(0.58),
                    CARD_WIDTH - Inches(0.40), Inches(0.03)
                )
                accent_line.fill.solid()
                accent_line.fill.fore_color.rgb = RGBColor(255, 190, 0)
                accent_line.line.fill.background()

                text_box = chart_slide.shapes.add_textbox(
                    CARD_LEFT + Inches(0.20), CARD_TOP + Inches(0.72),
                    CARD_WIDTH - Inches(0.40), CARD_HEIGHT - Inches(0.85)
                )
                tf = text_box.text_frame
                tf.word_wrap = True
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

                estimated_lines = sum(max(1, math.ceil(len(item) / 45)) for item in insights)
                font_sz = (
                    Pt(11.0) if estimated_lines <= 6 and n_items <= 3
                    else Pt(10.0) if estimated_lines <= 9 and n_items <= 5
                    else Pt(9.0) if estimated_lines <= 13
                    else Pt(8.5)
                )
                for i, item in enumerate(insights):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.space_after = Pt(8)
                    p.text = f"• {item}"
                    for r in p.runs:
                        r.font.name = "Verdana"
                        r.font.size = font_sz
                        r.font.color.rgb = RGBColor(255, 255, 255)

            continue

        if chart_shape is not None:
            c = chart_shape.chart
            cd = CategoryChartData()
            cd.categories = chart_model.categories
            for ser in chart_model.series:
                cd.add_series(ser.name, ser.values)
            c.replace_data(cd)
            c.has_title = False

            if len(chart_model.series) > 1:
                c.has_legend = True
                c.legend.position = XL_LEGEND_POSITION.TOP
                c.legend.include_in_layout = False
                try:
                    c.legend.font.name = "Verdana"
                    c.legend.font.size = Pt(10)
                except Exception:
                    pass

            try:
                c.category_axis.tick_labels.font.name = "Verdana"
                c.category_axis.tick_labels.font.size = Pt(9)
                c.value_axis.tick_labels.font.name = "Verdana"
                c.value_axis.tick_labels.font.size = Pt(9)
            except Exception:
                pass

            try:
                if c.plots:
                    plot = c.plots[0]
                    plot.vary_by_categories = False
                    plot.has_data_labels = True
                    plot.data_labels.font.name = "Verdana"
                    plot.data_labels.font.size = Pt(8)
                    for s_idx, ser in enumerate(plot.series):
                        color = TEMPLATE_SERIES_COLORS[s_idx % len(TEMPLATE_SERIES_COLORS)]
                        try:
                            fill = ser.format.fill
                            fill.solid()
                            fill.fore_color.rgb = color
                        except Exception:
                            pass
            except Exception:
                pass

        side_group = None
        for cand_name in ["Group 11", "Group 7"]:
            g = _shape_by_name(chart_slide, cand_name)
            if g is not None:
                side_group = g
                break
        if side_group is None:
            for sh in chart_slide.shapes:
                if sh.shape_type == 6 and any("Rectangle" in sub.name for sub in getattr(sh, "shapes", [])):
                    side_group = sh
                    break

        if side_group is not None:
            if chart_model.insights:
                insights = chart_model.insights
                n_items = len(insights)
                total_chars = sum(len(it) for it in insights)
                avg_chars = total_chars / max(1, n_items)
                max_chars = max(len(it) for it in insights) if insights else 0

                TOTAL_CONTENT_WIDTH = Inches(12.533)
                CONTENT_LEFT = Inches(0.40)
                GAP = Inches(0.25)

                if max_chars > 85 or total_chars > 380 or (n_items >= 6 and avg_chars > 55):
                    CARD_WIDTH = int(Inches(3.85))
                elif max_chars > 50 or total_chars > 200 or n_items >= 4:
                    CARD_WIDTH = int(Inches(3.48))
                else:
                    CARD_WIDTH = int(Inches(3.00))

                CHART_WIDTH = int(TOTAL_CONTENT_WIDTH - CARD_WIDTH - GAP)
                CHART_LEFT = int(CONTENT_LEFT)
                CARD_LEFT = int(CHART_LEFT + CHART_WIDTH + GAP)

                CARD_TOP = int(Inches(1.45))
                CARD_HEIGHT = int(Inches(5.25))

                if chart_shape is not None:
                    chart_shape.left = CHART_LEFT
                    chart_shape.top = int(Inches(1.40))
                    chart_shape.width = CHART_WIDTH
                    chart_shape.height = int(Inches(5.30))

                rect_names = {s.name for s in side_group.shapes if "Rectangle" in s.name or s.shape_type == 1}

                spTree = chart_slide.shapes._spTree
                for sp in list(side_group._element.xpath("p:sp")):
                    spTree.append(sp)
                side_group.element.getparent().remove(side_group.element)

                unpacked_rects = [sh for sh in chart_slide.shapes if sh.name in rect_names]
                if len(unpacked_rects) >= 2:
                    blue_card = max(unpacked_rects, key=lambda r: r.height)
                    badge = min(unpacked_rects, key=lambda r: r.height)
                elif len(unpacked_rects) == 1:
                    blue_card = unpacked_rects[0]
                    badge = None
                else:
                    blue_card = None
                    badge = None

                for s in list(chart_slide.shapes):
                    if s not in (blue_card, badge) and s.has_text_frame and s.text_frame.text:
                        txt = s.text_frame.text.strip().upper()
                        if "INSIGHT" in txt or "OBSERVATION" in txt or "KPI SUMMARY REQUIRED" in txt:
                            try:
                                chart_slide.shapes._spTree.remove(s._element)
                            except Exception:
                                pass

                insights_title = (getattr(chart_model, "insights_title", None) or "KEY OBSERVATIONS").strip().upper()
                raw_badge_width = int(Inches(len(insights_title) * 0.125 + 0.40))
                BADGE_WIDTH = int(min(max(raw_badge_width, Inches(2.20)), CARD_WIDTH - Inches(0.35)))
                BADGE_HEIGHT = int(Inches(0.48))
                BADGE_LEFT = int(CARD_LEFT + (CARD_WIDTH - BADGE_WIDTH) // 2)
                BADGE_TOP = int(CARD_TOP - Inches(0.24))

                if badge is not None:
                    badge.width = BADGE_WIDTH
                    badge.height = BADGE_HEIGHT
                    badge.left = BADGE_LEFT
                    badge.top = BADGE_TOP
                    tf_b = badge.text_frame
                    tf_b.margin_left = Inches(0.05)
                    tf_b.margin_right = Inches(0.05)
                    tf_b.margin_top = Inches(0.05)
                    tf_b.margin_bottom = Inches(0.05)
                    tf_b.word_wrap = False
                    p_b = tf_b.paragraphs[0]
                    p_b.text = insights_title
                    if len(insights_title) > 28:
                        badge_font_sz = Pt(9.5)
                    elif len(insights_title) > 20:
                        badge_font_sz = Pt(10.5)
                    else:
                        badge_font_sz = Pt(12)
                    for r in p_b.runs:
                        r.font.name = "Verdana"
                        r.font.size = badge_font_sz
                        r.font.bold = True
                        r.font.color.rgb = RGBColor(255, 190, 0)

                if blue_card is not None:
                    blue_card.left = CARD_LEFT
                    blue_card.top = CARD_TOP
                    blue_card.width = CARD_WIDTH
                    blue_card.height = CARD_HEIGHT
                    tf = blue_card.text_frame
                    tf.margin_left = Inches(0.18)
                    tf.margin_right = Inches(0.18)
                    tf.margin_top = Inches(0.45)
                    tf.word_wrap = True

                    sample_pPr = None
                    if len(tf.paragraphs) > 1 and tf.paragraphs[1]._p.pPr is not None:
                        sample_pPr = copy.deepcopy(tf.paragraphs[1]._p.pPr)
                    tf.clear()

                    printable_width_in = (CARD_WIDTH - Inches(0.36)) / Inches(1)
                    chars_per_line = int(printable_width_in * 14.5)
                    estimated_lines = sum(max(1, math.ceil(len(item) / chars_per_line)) for item in insights)

                    if estimated_lines <= 6 and n_items <= 3:
                        font_sz = Pt(11.5)
                    elif estimated_lines <= 9 and n_items <= 5:
                        font_sz = Pt(10.5)
                    elif estimated_lines <= 13:
                        font_sz = Pt(9.5)
                    else:
                        font_sz = Pt(8.5)

                    for i, item in enumerate(insights):
                        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                        p.text = item
                        p.font.name = "Verdana"
                        p.font.size = font_sz
                        for r in p.runs:
                            r.font.name = "Verdana"
                            r.font.size = font_sz
                            r.font.color.rgb = RGBColor(255, 255, 255)
                        if sample_pPr is not None:
                            if p._p.pPr is not None:
                                p._p.remove(p._p.pPr)
                            p._p.insert(0, copy.deepcopy(sample_pPr))
            else:
                side_group.element.getparent().remove(side_group.element)
                if chart_shape is not None:
                    chart_shape.left = Inches(0.40)
                    chart_shape.width = Inches(12.50)


def _dispatch_kpi_tables(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    """Renders one slide per structured tabular dataset (0-4 instances)."""
    table_list = _get_table_list(plan)
    for tbl_model in table_list:
        if not (tbl_model and tbl_model.rows):
            continue
        kpi_slide = _clone_slide(prs, IDX_BLANK_CANVAS)
        _strip_guidance_shapes(kpi_slide)

        table_title = (tbl_model.table_title or "KEY PERFORMANCE INDICATOR DASHBOARD").upper()
        _set_slide_header_and_sub(kpi_slide, table_title, plan.facility_name)

        headers = tbl_model.headers or ["Metric", "Actual", "Target"]
        rows = tbl_model.rows
        n_cols = len(headers)
        n_rows = len(rows)

        table_shape = kpi_slide.shapes.add_table(
            n_rows + 1, n_cols,
            Inches(0.40), Inches(1.45), Inches(12.53),
            min(Inches(5.20), int(Inches(0.48 * (n_rows + 1))))
        )
        tbl = table_shape.table

        for c_idx, h in enumerate(headers):
            cell = tbl.cell(0, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0, 43, 73)
            cell.text = str(h)
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.name = "Verdana"
                    r.font.size = Pt(11)
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(255, 255, 255)

        max_c_len = max((len(str(val)) for row in rows for val in row), default=0)
        if max_c_len > 70 or n_rows >= 8:
            row_font_sz = Pt(9.0)
        elif max_c_len > 40:
            row_font_sz = Pt(10.0)
        else:
            row_font_sz = Pt(10.5)

        for r_idx, row in enumerate(rows):
            for c_idx in range(n_cols):
                val = str(row[c_idx]) if c_idx < len(row) else ""
                cell = tbl.cell(r_idx + 1, c_idx)
                cell.text = val
                cell.fill.solid()
                if r_idx % 2 == 0:
                    cell.fill.fore_color.rgb = RGBColor(245, 248, 252)
                else:
                    cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
                cell.margin_left = Inches(0.08)
                cell.margin_right = Inches(0.08)
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.name = "Verdana"
                        r.font.size = row_font_sz
                        r.font.color.rgb = RGBColor(0, 43, 73)


def _dispatch_kpi_dashboard(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    """Mutually exclusive with kpi_tables (matches the template's single dashboard slide)."""
    if _get_table_list(plan):
        return
    if not (plan.kpi_safety_quality or plan.kpi_operational):
        return
    kpi_slide = _clone_slide(prs, IDX_KPI_DASHBOARD)
    _set_or_remove_facility_placeholder(kpi_slide, plan.facility_name)
    kpi_dash_title = _shape_by_name(kpi_slide, "Title 2")
    if kpi_dash_title is not None:
        _set_first_run_text(kpi_dash_title, (plan.kpi_dashboard_title or "KEY PERFORMANCE INDICATOR DASHBOARD").strip().upper())
    kpi_tables = [s for s in kpi_slide.shapes if s.has_table]
    sq_shape = next((s for s in kpi_tables if len(s.table.columns) == 7), None)
    op_shape = next((s for s in kpi_tables if len(s.table.columns) == 3), None)

    if plan.kpi_safety_quality and sq_shape is not None:
        sq_tbl = sq_shape.table
        for i, row in enumerate(plan.kpi_safety_quality):
            r = 2 + i
            if r < len(sq_tbl.rows):
                sq_tbl.cell(r, 0).text = row.label
                for c in (1, 4):
                    sq_tbl.cell(r, c).text = row.actual
                for c in (2, 5):
                    sq_tbl.cell(r, c).text = row.target or "-"
                for c in (3, 6):
                    sq_tbl.cell(r, c).text = row.actual
                for c_idx in range(len(sq_tbl.columns)):
                    for para in sq_tbl.cell(r, c_idx).text_frame.paragraphs:
                        for run in para.runs:
                            run.font.name = "Verdana"
                            run.font.size = Pt(9.5)
                            run.font.color.rgb = RGBColor(0, 43, 73)
        for r in range(2 + len(plan.kpi_safety_quality), len(sq_tbl.rows)):
            for c in range(len(sq_tbl.columns)):
                sq_tbl.cell(r, c).text = ""
    elif sq_shape is not None:
        kpi_slide.shapes._spTree.remove(sq_shape._element)
        if op_shape is not None:
            op_shape.top = Inches(1.40)

    if plan.kpi_operational and op_shape is not None:
        op_tbl = op_shape.table
        op_tbl.cell(0, 0).text = "Operational Metric"
        op_tbl.cell(0, 1).text = "Actual"
        op_tbl.cell(0, 2).text = "Target"
        rows = [[r.label, r.actual, (r.target or "").strip() or "-"] for r in plan.kpi_operational]
        _fill_table_rows(op_tbl, rows, start_row=1)
    elif op_shape is not None and not plan.kpi_operational:
        kpi_slide.shapes._spTree.remove(op_shape._element)


def _dispatch_stat_highlights(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not plan.stat_highlights:
        return
    stat_slide = _clone_slide(prs, IDX_BLANK_CANVAS)
    _strip_guidance_shapes(stat_slide)
    _set_slide_header_and_sub(stat_slide, plan.stat_highlights_title or "PERFORMANCE AT A GLANCE", plan.facility_name)
    _render_stat_highlights(stat_slide, plan.stat_highlights)


def _dispatch_process_flow(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not plan.process_flow:
        return
    flow_slide = _clone_slide(prs, IDX_BLANK_CANVAS)
    _strip_guidance_shapes(flow_slide)
    _set_slide_header_and_sub(flow_slide, plan.process_flow_title or "PROCESS OVERVIEW", plan.facility_name)
    _render_process_flow(flow_slide, plan.process_flow)


def _dispatch_voice_of_customer(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not (plan.voice_of_customer and plan.voice_of_customer.quote):
        return
    voc_slide = _clone_slide(prs, IDX_VOICE_OF_CUSTOMER)
    _remove_shapes(voc_slide, lambda s: s.name == "TextBox 5")  # stale guidance placeholder
    voc_title = _shape_by_name(voc_slide, "Title 2")
    if voc_title is not None:
        _set_first_run_text(voc_title, (plan.voice_of_customer_title or "VOICE OF THE CUSTOMER").strip().upper())
    bubble = _shape_by_name(voc_slide, "Speech Bubble: Rectangle with Corners Rounded 4")
    if bubble is not None:
        paras = bubble.text_frame.paragraphs
        if paras and paras[0].runs:
            paras[0].runs[0].text = plan.voice_of_customer.quote
        if len(paras) > 2 and paras[2].runs:
            paras[2].runs[0].text = plan.voice_of_customer.attribution


def _dispatch_divider_ci(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not (plan.gemba_walk or plan.ci_tracker):
        return
    ci_divider = _clone_slide(prs, IDX_SECTION_CIP)
    t2 = _shape_by_name(ci_divider, "Title 6")
    if t2 is not None:
        _set_first_run_text(t2, plan.ci_section_title or "CONTINUOUS IMPROVEMENT PROGRAM UPDATES")


def _dispatch_gemba_walk(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not plan.gemba_walk:
        return
    gemba_slide = _clone_slide(prs, IDX_GEMBA_WALK)
    _strip_guidance_shapes(gemba_slide)
    gemba_title = _shape_by_name(gemba_slide, "Title 2")
    if gemba_title is not None:
        _set_first_run_text(gemba_title, (plan.gemba_walk_title or "FACILITY WALKTHROUGH FINDINGS").strip().upper())
    gemba_intro = _shape_by_name(gemba_slide, "TextBox 6")
    if gemba_intro is not None and plan.gemba_walk_intro:
        _set_first_run_text(gemba_intro, plan.gemba_walk_intro)
    gemba_table_shape = _shape_by_name(gemba_slide, "Table 5")
    if gemba_table_shape is not None and gemba_table_shape.has_table:
        rows = [[r.area, r.observation] for r in plan.gemba_walk]
        _fill_table_rows(gemba_table_shape.table, rows, start_row=1)


def _dispatch_ci_tracker(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not plan.ci_tracker:
        return
    ci_tracker_slide = _clone_slide(prs, IDX_CI_TRACKER)
    _strip_guidance_shapes(ci_tracker_slide)
    _strip_decorative_connectors(ci_tracker_slide)
    ci_tracker_title = _shape_by_name(ci_tracker_slide, "Title 2")
    if ci_tracker_title is not None:
        _set_first_run_text(ci_tracker_title, (plan.ci_tracker_title or "IMPROVEMENT INITIATIVE TRACKER").strip().upper())
    ci_table_shape = _shape_by_name(ci_tracker_slide, "Table 4")
    if ci_table_shape is not None and ci_table_shape.has_table:
        rows = [[r.activity, r.category, r.status, r.value, r.comment] for r in plan.ci_tracker]
        _fill_table_rows(ci_table_shape.table, rows, start_row=1)


def _dispatch_divider_quality(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not (plan.quality_org_structure or plan.nc_review_narrative or plan.nc_review_summary or plan.nc_tracker):
        return
    quality_divider = _clone_slide(prs, IDX_SECTION_QUALITY)
    t3 = _shape_by_name(quality_divider, "Title 6")
    if t3 is not None:
        _set_first_run_text(t3, plan.quality_section_title or "QUALITY MANAGEMENT SYSTEM UPDATES")


def _dispatch_quality_org_structure(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not plan.quality_org_structure:
        return
    quality_org_slide = _clone_slide(prs, IDX_QUALITY_ORG)
    _strip_guidance_shapes(quality_org_slide)
    quality_title = _shape_by_name(quality_org_slide, "Title 1")
    if quality_title is not None:
        _set_first_run_text(quality_title, (plan.quality_org_structure_title or "QUALITY & COMPLIANCE LEADERSHIP").strip().upper())
    quality_boxes = sorted(
        (s for s in quality_org_slide.shapes if s.name.startswith("Rectangle: Rounded Corners")),
        key=lambda s: (s.top, s.left),
    )
    for box, person in zip(quality_boxes, plan.quality_org_structure):
        _set_org_box(box, person.name, f"{person.role}\nUPS Healthcare")
    for box in quality_boxes[len(plan.quality_org_structure):]:
        _set_org_box(box, "", "")


def _dispatch_nc_review(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not (plan.nc_review_narrative or plan.nc_review_summary):
        return
    nc_review_slide = _clone_slide(prs, IDX_NC_REVIEW)
    _strip_guidance_shapes(nc_review_slide)
    _set_or_remove_facility_placeholder(nc_review_slide, plan.facility_name)
    nc_review_title = _shape_by_name(nc_review_slide, "Title 2")
    if nc_review_title is not None:
        _set_first_run_text(nc_review_title, (plan.nc_review_title or "ISSUE REVIEW & CAPA SUMMARY").strip().upper())
    nc_narrative = _shape_by_name(nc_review_slide, "TextBox 7")
    if nc_narrative is not None and plan.nc_review_narrative:
        nc_narrative.text_frame.text = plan.nc_review_narrative
    nc_summary_table = _shape_by_name(nc_review_slide, "Table 5")
    if nc_summary_table is not None and nc_summary_table.has_table and plan.nc_review_summary:
        tbl = nc_summary_table.table
        for c_idx, val in enumerate(plan.nc_review_summary[: len(tbl.columns)]):
            tbl.cell(1, c_idx).text = val


def _dispatch_nc_tracker(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not plan.nc_tracker:
        return
    nc_tracker_slide = _clone_slide(prs, IDX_NC_TRACKER)
    _strip_guidance_shapes(nc_tracker_slide)
    _strip_decorative_connectors(nc_tracker_slide)
    nc_tracker_title = _shape_by_name(nc_tracker_slide, "Title 2")
    if nc_tracker_title is not None:
        _set_first_run_text(nc_tracker_title, (plan.nc_tracker_title or "ISSUE & CORRECTIVE ACTION TRACKER").strip().upper())
    nc_tracker_table = _shape_by_name(nc_tracker_slide, "Table 4")
    if nc_tracker_table is not None and nc_tracker_table.has_table:
        rows = [[r.period, r.nc_id, r.event, r.due_date, r.status] for r in plan.nc_tracker]
        _fill_table_rows(nc_tracker_table.table, rows, start_row=1)


def _dispatch_next_steps(prs: Any, plan: HLDQBRPresentationPlan) -> None:
    if not plan.next_steps:
        return
    next_steps_slide = _clone_slide(prs, IDX_NEXT_STEPS)
    _set_or_remove_facility_placeholder(next_steps_slide, plan.facility_name)

    ns_title = _shape_by_name(next_steps_slide, "Title 1")
    if ns_title is not None and ns_title.text_frame.paragraphs:
        heading_text = (
            plan.next_steps_title.strip().upper()
            if plan.next_steps_title
            else "NEXT STEPS & TARGET TIMELINES"
        )
        p0 = ns_title.text_frame.paragraphs[0]
        p0.text = heading_text
        if p0.runs:
            p0.runs[0].font.size = Pt(22)
            p0.runs[0].font.bold = True

    ns_table_shape = _shape_by_name(next_steps_slide, "Table 8")
    if ns_table_shape is not None and ns_table_shape.has_table:
        tbl = ns_table_shape.table
        tbl._tbl.tblPr.set("firstRow", "0")
        rows = [
            [s.step, (s.date or "").strip() or "Target Q4 2026"]
            for s in plan.next_steps
        ]
        _fill_table_rows(tbl, rows, start_row=0, text_color=RGBColor(0, 43, 73))


ARCHETYPE_DISPATCH: Dict[str, Any] = {
    "org_structure": _dispatch_org_structure,
    "achievements": _dispatch_achievements,
    "priorities": _dispatch_priorities,
    "DIVIDER_PERFORMANCE": _dispatch_divider_performance,
    "action_tracker": _dispatch_action_tracker,
    "charts": _dispatch_charts,
    "kpi_tables": _dispatch_kpi_tables,
    "kpi_dashboard": _dispatch_kpi_dashboard,
    "stat_highlights": _dispatch_stat_highlights,
    "process_flow": _dispatch_process_flow,
    "voice_of_customer": _dispatch_voice_of_customer,
    "DIVIDER_CI": _dispatch_divider_ci,
    "gemba_walk": _dispatch_gemba_walk,
    "ci_tracker": _dispatch_ci_tracker,
    "DIVIDER_QUALITY": _dispatch_divider_quality,
    "quality_org_structure": _dispatch_quality_org_structure,
    "nc_review_summary": _dispatch_nc_review,
    "nc_tracker": _dispatch_nc_tracker,
    "next_steps": _dispatch_next_steps,
}

# Fallback sequence when the LLM's narrative order is missing/invalid — mirrors
# the deck's original fixed order so behavior is unchanged unless a valid
# LLM-authored order is supplied (see llm/prompts_hld_qbr.py's selection prompt).
DEFAULT_ARCHETYPE_ORDER: List[str] = [
    "org_structure", "achievements", "priorities",
    "DIVIDER_PERFORMANCE", "action_tracker", "charts", "kpi_tables", "kpi_dashboard",
    "stat_highlights", "process_flow", "voice_of_customer",
    "DIVIDER_CI", "gemba_walk", "ci_tracker",
    "DIVIDER_QUALITY", "quality_org_structure", "nc_review_summary", "nc_tracker",
    "next_steps",
]


def _resolve_archetype_order(plan: HLDQBRPresentationPlan) -> List[str]:
    """Validates the LLM-provided narrative order against the known dispatch
    keys; falls back to the default fixed order if missing/invalid so the
    pipeline never breaks on a bad or absent order. Any dispatch key the LLM's
    order omitted is appended at the end so real content is never dropped."""
    order = list(getattr(plan, "narrative_order", None) or [])
    valid = [key for key in order if key in ARCHETYPE_DISPATCH]
    if not valid:
        return DEFAULT_ARCHETYPE_ORDER
    missing = [key for key in DEFAULT_ARCHETYPE_ORDER if key not in valid]
    return valid + missing


class HLDQBRBuilder:
    """Builds an HLD QBR-branded .pptx from a validated HLDQBRPresentationPlan."""
    def __init__(self, template_path=None):
        self.template_path = template_path or config.HLD_QBR_TEMPLATE_FILE

    def build(self, plan: HLDQBRPresentationPlan) -> bytes:
        prs = Presentation(str(self.template_path))
        initial_count = len(prs.slides)

        # 1. Cover (always)
        cover = _clone_slide(prs, IDX_COVER)
        _set_cover_breadcrumb_blank(cover)
        date_ph = _shape_by_name(cover, "Date Placeholder 1")
        # Always stamp today's date when generating the PPTX (matching template ordinal style)
        today_str = _get_ordinal_date(datetime.now())
        if date_ph is not None:
            _set_first_run_text(date_ph, today_str)
        _set_cover_title(cover, plan.presentation_title)

        # 2. Agenda (always; falls back to template's default topics if none given)
        agenda = _clone_slide(prs, IDX_AGENDA)

        # Dynamic agenda slide heading
        agenda_title_shape = _shape_by_name(agenda, "Title 1")
        if agenda_title_shape is not None and agenda_title_shape.text_frame.paragraphs:
            heading_text = (
                plan.agenda_title.strip().upper()
                if getattr(plan, "agenda_title", "") and plan.agenda_title.strip()
                else "TODAY'S AGENDA"
            )
            p0 = agenda_title_shape.text_frame.paragraphs[0]
            p0.text = heading_text
            if p0.runs:
                p0.runs[0].font.name = "Verdana"
                p0.runs[0].font.bold = True

        if plan.agenda_topics:
            agenda_box = _shape_by_name(agenda, "TextBox 7")
            if agenda_box is not None:
                tf = agenda_box.text_frame
                # Extract template's bullet paragraph properties to replicate across all topics
                sample_pPr = None
                if tf.paragraphs and tf.paragraphs[0]._p.find(qn("a:pPr")) is not None:
                    sample_pPr = copy.deepcopy(tf.paragraphs[0]._p.find(qn("a:pPr")))
                tf.clear()
                n_topics = len(plan.agenda_topics)
                if n_topics <= 3:
                    ag_sz = Pt(18)
                    spc_after = Pt(16)
                elif n_topics <= 5:
                    ag_sz = Pt(16)
                    spc_after = Pt(12)
                else:
                    ag_sz = Pt(14)
                    spc_after = Pt(8)

                for i, topic in enumerate(plan.agenda_topics):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.text = topic
                    if sample_pPr is not None:
                        pPr = p._p.find(qn("a:pPr"))
                        if pPr is not None:
                            p._p.remove(pPr)
                        p._p.insert(0, copy.deepcopy(sample_pPr))
                    p.space_after = spc_after
                    for r in p.runs:
                        r.font.name = "Verdana"
                        r.font.size = ag_sz
                        r.font.color.rgb = RGBColor(0, 43, 73)

        # 3. Executive Summary
        # A. Section Header slide (template slide 4 / index 3) — clean photo background with centered title
        exec_header = _clone_slide(prs, IDX_EXECUTIVE_SUMMARY)
        _strip_guidance_shapes(exec_header)
        t6 = _shape_by_name(exec_header, "Title 6")
        if t6 is not None:
            _set_first_run_text(t6, plan.executive_summary_title or "EXECUTIVE SUMMARY")

        # B. Content slide (template slide 21 / index 20) — clean white canvas with refined horizontal cards
        exec_content = _clone_slide(prs, IDX_BLANK_CANVAS)
        _strip_guidance_shapes(exec_content)

        es_title = plan.executive_summary_title or "EXECUTIVE SUMMARY: STRATEGIC & OPERATIONAL HIGHLIGHTS"
        subtitle = plan.facility_name or "Comprehensive Performance Summary & Strategic Milestones | UPS Healthcare"
        _set_slide_header_and_sub(exec_content, es_title, subtitle)

        summary_bullets = plan.executive_summary or []
        if not summary_bullets and plan.priorities:
            summary_bullets = [f"{p.heading}: {p.body}" for p in plan.priorities[:5]]

        def _parse_bullet_item(item_text: str, index: int) -> tuple[str, str]:
            """Extracts (category_tag, body_text) from a summary bullet."""
            if ":" in item_text:
                parts = item_text.split(":", 1)
                tag_candidate = parts[0].strip().upper()
                body_candidate = parts[1].strip()
                if 2 <= len(tag_candidate) <= 32:
                    return tag_candidate, body_candidate

            lower = item_text.lower()
            if any(k in lower for k in ["otif", "on-time", "delivery", "sla", "delay"]):
                tag = "ON-TIME SERVICE EXCELLENCE"
            elif any(k in lower for k in ["cold-chain", "temperature", "temp", "pharma", "compliance", "excursion"]):
                tag = "COLD-CHAIN INTEGRITY"
            elif any(k in lower for k in ["emission", "carbon", "empty mile", "green", "fleet", "transport"]):
                tag = "FLEET & CARBON EFFICIENCY"
            elif any(k in lower for k in ["cost", "saving", "spend", "usd", "dollar", "$", "revenue"]):
                tag = "FINANCIAL & VALUE CREATION"
            elif any(k in lower for k in ["global", "countr", "who", "supply", "volume", "order", "reach"]):
                tag = "GLOBAL SCALE & IMPACT"
            elif any(k in lower for k in ["digital", "rfid", "barcode", "scan", "system", "track"]):
                tag = "DIGITAL INTELLIGENCE"
            else:
                default_tags = [
                    "GLOBAL SCALE & IMPACT",
                    "SERVICE LEVEL EXCELLENCE",
                    "OPERATIONAL EFFICIENCY",
                    "QUALITY & COMPLIANCE",
                    "STRATEGIC VALUE CREATION",
                ]
                tag = default_tags[index % len(default_tags)]
            return tag, item_text.strip()

        if summary_bullets:
            n_b = len(summary_bullets[:5])
            TOP_START = Inches(1.30)
            TOTAL_H = Inches(5.35)
            GAP = Inches(0.12)
            CARD_H = min(Inches(1.10), (TOTAL_H - GAP * (n_b - 1)) / max(n_b, 1))
            LEFT = Inches(0.40)
            WIDTH = Inches(12.53)

            ACCENT_COLORS = [
                RGBColor(255, 190, 0),  # UPS Gold for first card
                RGBColor(0, 43, 73),    # UPS Navy
                RGBColor(0, 112, 192),  # Brand Blue
                RGBColor(0, 43, 73),
                RGBColor(0, 112, 192),
            ]

            for i, raw_b in enumerate(summary_bullets[:5]):
                tag, body = _parse_bullet_item(raw_b, i)
                card_top = TOP_START + i * (CARD_H + GAP)

                # 1. Main Card Container (Rounded Rectangle)
                card = exec_content.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, LEFT, card_top, WIDTH, CARD_H)
                card.fill.solid()
                card.fill.fore_color.rgb = RGBColor(248, 250, 253)
                card.line.color.rgb = RGBColor(218, 228, 238)
                card.line.width = Pt(0.75)

                # 2. Left Accent Bar
                BAR_W = Inches(0.10)
                bar = exec_content.shapes.add_shape(MSO_SHAPE.RECTANGLE, LEFT, card_top, BAR_W, CARD_H)
                bar.fill.solid()
                bar.fill.fore_color.rgb = ACCENT_COLORS[i % len(ACCENT_COLORS)]
                bar.line.fill.background()

                # 3. Number Badge (Circle)
                BADGE_SZ = Inches(0.46)
                badge_top = card_top + (CARD_H - BADGE_SZ) / 2
                badge = exec_content.shapes.add_shape(MSO_SHAPE.OVAL, LEFT + Inches(0.25), badge_top, BADGE_SZ, BADGE_SZ)
                badge.fill.solid()
                badge.fill.fore_color.rgb = RGBColor(0, 43, 73)
                badge.line.fill.background()
                tf_b = badge.text_frame
                tf_b.margin_left = tf_b.margin_right = tf_b.margin_top = tf_b.margin_bottom = Inches(0.01)
                p_b = tf_b.paragraphs[0]
                p_b.text = f"{i+1}"
                p_b.alignment = PP_ALIGN.CENTER
                for r in p_b.runs:
                    r.font.name = "Verdana"
                    r.font.size = Pt(11.5)
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(255, 255, 255)

                # 4. Category Tag Pill
                pill = exec_content.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, LEFT + Inches(0.85), card_top + Inches(0.12), Inches(2.35), Inches(0.32))
                pill.fill.solid()
                pill.fill.fore_color.rgb = RGBColor(234, 243, 252)
                pill.line.color.rgb = RGBColor(190, 215, 240)
                pill.line.width = Pt(0.5)
                tf_p = pill.text_frame
                tf_p.margin_left = tf_p.margin_right = tf_p.margin_top = tf_p.margin_bottom = Inches(0.02)
                p_p = tf_p.paragraphs[0]
                p_p.text = tag
                p_p.alignment = PP_ALIGN.CENTER
                for r in p_p.runs:
                    r.font.name = "Verdana"
                    r.font.size = Pt(8.5)
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(0, 43, 73)

                # 5. Narrative Text
                tb = exec_content.shapes.add_textbox(LEFT + Inches(3.35), card_top + Inches(0.08), WIDTH - Inches(3.50), CARD_H - Inches(0.16))
                tf = tb.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                r = p.add_run()
                r.text = body
                r.font.name = "Verdana"
                r.font.size = Pt(11.5 if len(body) < 130 else 10.5)
                r.font.color.rgb = RGBColor(30, 41, 59)

        # 4-17. Narrative-ordered content archetypes — sequence decided by the
        # LLM's Stage 1 narrative_order (falls back to the template's original
        # fixed order if missing/invalid). See _resolve_archetype_order().
        for archetype_key in _resolve_archetype_order(plan):
            ARCHETYPE_DISPATCH[archetype_key](prs, plan)

        # 18. Mandatory Closing (always, intact)
        _clone_slide(prs, IDX_CLOSING)

        # Prune the original 60 source slides, leaving only the generated ones
        rId_attr = qn("r:id")
        sldIdLst = prs.slides._sldIdLst
        for _ in range(initial_count):
            sldId = sldIdLst[0]
            rId = sldId.get(rId_attr)
            if rId:
                prs.part.drop_rel(rId)
            sldIdLst.remove(sldId)

        logger.info("[HLD QBR Builder] Built %d slides.", len(prs.slides))

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        return buf.read()
