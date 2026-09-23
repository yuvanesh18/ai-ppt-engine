# HLD QBR Template — Integration Plan
### New Pluggable Template: `hld_qbr` ("HLD QBR Template PFv3.potx")

---

## 1. What was analyzed

Source file: `HLD QBR Template PFv3.potx` (PowerPoint **Template** — content type
`presentationml.template.main+xml`, not directly openable by `python-pptx`).

- **60 slides**, 16:9 (12192000 x 6858000 EMU), single slide master with **19 slide layouts**
  (`Cover Blue`, `Co-branding Title Slide`, `Title Only`, `1_Section Header_No Image`,
  `2_Section Header_No Image_Blue BG`, `Title and Content`, `Blank`, `Thank you blue`, etc.).
- This is a real customer-branded **UPS Healthcare Quarterly Business Review (QBR)** deck
  (branding: UPS Healthcare™ / "BIORIDGE PHARMA Monthly Business Review" example cover),
  built for **account teams to fill in per-customer**, not a generic slide-shape library like
  `dark_navy`'s 11 layouts.
- It is dense with **native OOXML features** that go beyond python-pptx's high-level API:
  native `Table` shapes (trackers/scorecards), embedded `Chart` parts (with their own embedded
  XLSX workbook parts), `EMBEDDED_OLE_OBJECT` shapes on the cover, `FREEFORM` org-chart
  connectors, and large `GROUP` shapes (icon libraries). This is why the plan calls for
  **direct OOXML/lxml manipulation**, not just `python-pptx`'s shape-add API — the same
  deep-clone approach already used by [core/builders/template1_builder.py](core/builders/template1_builder.py)
  and [core/builders/techm_v3_builder.py](core/builders/techm_v3_builder.py).
- **Confirmed the exact feature the user asked about**: many slides contain small
  `Rectangle`/`TextBox` shapes that are **pure authoring guidance**, e.g.:
  - `"Required slide | Example text provided"`
  - `"Required content | Format Option 1 | May include customer slides..."`
  - `"Optional | EXAMPLE"`
  - `"Additional slides available to use | Hyperlinks included | Should be relevant to your customer"`
  - `"KPI summary required | EXAMPLE"`
  - `"Example only – | How to summarize performance | Must be updated"`

  These are instructions for the **human/agent building the deck** and must never reach the
  generated output. There is also an entire trailing **appendix of guide-only slides**
  (slides 41–59: "FORMATTING HELP" section — brand guardrails, approved colors, font-size
  cheatsheets, sample tables/charts, and 3 full icon-library slides) that must be excluded
  from generation entirely.

Full per-slide shape/text dump used for this analysis is saved at
[scratch/_hld_dump.txt](scratch/_hld_dump.txt) for reference during implementation.

---

## 2. Slide classification (drives what the builder is allowed to clone)

| # | Title / Purpose | Layout | Class | Guidance text found |
|---|---|---|---|---|
| 0 | Cover | Blank | **CORE** | — |
| 1 | Today's Discussion (agenda) | 9_Section Header | **CORE** | — |
| 2 | Organizational Structure | Title Only | **CORE** | — |
| 3 | Executive Summary | Title Only | **CORE** | — |
| 4 | Previous Quarter Achievements (timeline) | Title Only | **CORE** | "Required slide \| Example text provided" |
| 5 | [Customer] Priorities — Format Option 1 | Title Only | **CORE (choose 1 of 5/6)** | "Required content \| Format Option 1..." |
| 6 | [Customer] Priorities — Format Option 2 | Title Only | **ALT of #5** | "Format Option 2 – example text only" |
| 7 | UPS Healthcare Priorities | Title Only | **CORE** | "Required content \| Example Text Only..." |
| 8 | (supplemental slide gallery pointer) | Title Only | **GUIDE_ONLY** (pointer to external Seismic slides) | "Additional slides available to use..." |
| 9 | Action Item Tracker (table) | Title Only | **CORE** | "Required content \| EXAMPLE" |
| 10 | Voice of the Customer (quote) | Title Only | **OPTIONAL** *(no label — assumption, see §7)* | — |
| 11 | Section: Performance Management Updates | 1_Section Header_No Image | **CORE** (divider) | — |
| 12 | KPI Dashboard (2 tables) | Title Only | **CORE** | — |
| 13 | Inbound Summary (chart) | Title Only | **CORE** | "KPI summary required \| EXAMPLE" |
| 14 | Outbound Summary (chart) | Title Only | **CORE** | "KPI summary required \| EXAMPLE" |
| 15 | Inventory Accuracy (chart) | Title Only | **CORE** | "KPI summary required \| EXAMPLE" |
| 16 | Customer Forecast Accuracy (chart) | Title Only | **CORE** | "Required content, \| Customer provided..." |
| 17 | Space Utilization (chart) | Title Only | **CORE** | "Can be a total, by storage condition..." |
| 18 | Spend Summary (chart) | Title Only | **CORE** | "Can be by quarter, business unit..." |
| 19 | Accounts Receivable Updates (chart) | Title Only | **OPTIONAL** | "Optional \| EXAMPLE" |
| 20 | Master Data Management KPIs | Title Only | **CORE** | "Required content" |
| 21 | Section: Continuous Improvement Program Updates | 1_Section Header_No Image | **CORE** (divider) | — |
| 22 | Gemba Walk Summary (table) | Title Only | **CORE** | "Required content \| EXAMPLE" |
| 23 | CI Activity Tracker (table) | Title Only | **CORE** | "Required content \| EXAMPLE" |
| 24 | Section: Quality Management System Updates | 1_Section Header_No Image | **CORE** (divider) | — |
| 25 | Quality Organizational Structure | Title Only | **CORE** | "Example" |
| 26 | Non-Conformance Review (table) | Title Only | **CORE** | "Required content \| EXAMPLE 1" |
| 27 | Non-Conformance Tracker (table) | Title Only | **CORE** | "Required content \| EXAMPLE 2" |
| 28 | Section: Next Steps | 1_Section Header_No Image | **CORE** (divider) | — |
| 29 | Next Steps (table) | Title Only | **CORE** | — |
| 30 | Brand Closing Slide ("Moving our world forward") | Blank | **CORE — mandatory closing, always appended intact** | — |
| 31 | Section: OPTIONAL SLIDES | 2_Section Header_No Image_Blue BG | **GUIDE_ONLY divider** (only emit if ≥1 optional slide chosen) | — |
| 32–33 | KPI Scorecard Quarterly Summary (x2 layout variants) | Title Only | **OPTIONAL** | — |
| 34 | Performance Summary (brand stat highlights) | Blank | **OPTIONAL** | "Example only – How to summarize... Must be updated" |
| 35 | Value-Centric Approach (cost-avoidance table) | Title Only | **OPTIONAL** | "EXAMPLE – Formulas in Notes" |
| 36 | Financial Value (3 charts) | Title Only | **OPTIONAL** | — |
| 37–38 | Cost / Data / Trend Analysis (x2) | Title Only | **OPTIONAL** | "Examples –..." |
| 39 | Forward Looking Roadmap | Title Only | **OPTIONAL** | "Optional \| ..." / "Example items for discussion" |
| 40 | Technology and Digital Enablement | Title Only | **OPTIONAL** | — |
| 41 | Section: FORMATTING HELP | 2_Section Header_No Image_Blue BG | **GUIDE_ONLY — never emitted** | — |
| 42–59 | Guardrails, Brand colors, font/format cheatsheets, sample tables/charts, 3x icon libraries, Harvey-ball icons | mixed | **GUIDE_ONLY — never emitted, and safe to physically strip from the working copy** | (whole section is instructional) |

**Rule of thumb the builder encodes:** only slides classed `CORE` and (optionally) `OPTIONAL`
are ever cloneable content sources. `GUIDE_ONLY` slides are reference material for humans/LLM
prompt-authoring only and are removed from the distributable template asset entirely.

---

## 3. How the "information / guidance boxes" will be detected & stripped

Rather than relying on fragile text-content sniffing at generation time, guidance shapes will
be **tagged once, at template-conversion time**, then stripped deterministically at build time:

1. **Conversion pass** (one-time script, run now / whenever the source `.potx` changes):
   - Fix content type (`presentationml.template.main+xml` → `...presentation.main+xml`) so
     `python-pptx`/lxml can load it (already validated — see §6).
   - Walk every `CORE`/`OPTIONAL` slide's shape tree; any shape whose text matches the
     guidance vocabulary (regex, case-insensitive) —
     `required content|required slide|example|optional \||format option|additional slides available|must be updated|formulas in notes`
     — gets its XML element **renamed** with a reserved prefix, e.g.
     `sp.name = "__GUIDE__" + original_name`, instead of being deleted immediately.
     (Renaming, not deleting, at conversion time keeps a human-auditable diff and lets a
     reviewer confirm nothing legitimate got tagged before it's baked into the shipped asset.)
   - After manual review, a second pass **physically deletes** all `__GUIDE__*` shapes and all
     `GUIDE_ONLY` slides (8, 31, 41–59) from the working copy, producing the slim distributable
     `templates/hld_qbr_template.pptx` used at runtime. This also shrinks the file drastically
     (icon-library slides alone contain hundreds of freeform shapes).
2. **Runtime safety net**: the builder's `_clone_slide()` additionally re-applies the same
   regex check while copying shape elements, so if a future re-export of the source template
   reintroduces an untagged guidance box, it's still dropped instead of leaking into output.
3. Guidance text mined from slide 42 ("GUARDRAILS") and the brand rules on slide 43 are not
   discarded — they are **converted into LLM system-prompt constraints** (see §5) so the
   generator still respects them (e.g. "do not fabricate customer logos", "claims must be
   sourced") even though the slide itself is never shown to the user.

---

## 4. OpenXML strategy (full-fidelity, not just python-pptx high level API)

Following the pattern already proven in `template1_builder.py` / `techm_v3_builder.py`:

- **Deep clone via `lxml`**: `copy.deepcopy(shape.element)` appended directly to the target
  slide's `spTree`, instead of python-pptx's `shapes.add_*()` helpers — preserves 100% of
  formatting, effects, gradients, and geometry that the high-level API cannot reproduce.
- **Relationship remapping**: for every `r:embed` / `r:link` / `r:id` attribute found via
  `xpath(".//*[@r:embed or @r:link or @r:id]")`, re-`relate_to()` the underlying part on the
  new slide and rewrite the attribute — required for:
  - **Pictures** (logos, icons) — already solved in existing builders.
  - **Charts** — a chart shape's `<c:chart>` part itself owns a *nested* relationship to an
    embedded `.xlsx` workbook part; cloning a chart shape means cloning the chart part **and**
    its embedded package part, not just the top-level picture-style rId. This needs a dedicated
    `_clone_chart_part()` helper beyond what template1/techm_v3 needed (they had no charts).
  - **Embedded OLE objects** (cover slide's `Object 9`) — same nested-part problem; if the
    cover isn't mutated beyond title/date text, this can be cloned as an intact opaque blob.
  - **Tables** — no relationship issues, but python-pptx cannot delete/insert `<a:tr>` rows
    through its object model for arbitrary data sizes; row count changes require direct
    `<a:tbl>` XML surgery (clone a template `<a:tr>`, mutate `<a:t>` runs per cell, append/remove
    rows) — same technique already partly used for `ACTION_TABLE` in `template1_builder.py`.
- **Background/theme fidelity**: copy `slide.background.fill` explicitly per clone (as
  `template1_builder._clone_slide` already does) since `add_slide(layout)` does not inherit a
  slide-level override background.
- **Placeholder + free shape text mutation in place**: reuse the existing `_set_tf_text()`
  pattern (mutate first run, blank out the rest) to preserve every run's font/color/size instead
  of re-creating text frames from scratch.
- **Mandatory closing slide guarantee**: mirrors the existing TechM "Slide 35 guarantee" —
  slide 30 (brand closing) is always cloned intact and appended last, regardless of which
  optional slides were chosen.

---

## 5. New components to build

| File | Purpose |
|---|---|
| `templates/hld_qbr_template.pptx` | Slimmed, guidance-stripped, pptx-content-typed working copy produced by the one-time conversion script (§3, §6). Original `.potx` stays untouched as the source-of-truth reference. |
| `templates/convert_hld_qbr_template.py` | One-time/repeatable conversion script: potx→pptx content-type fix, guidance-shape tagging, `GUIDE_ONLY` slide + tagged-shape stripping, writes the slim asset. |
| `config.py` | Add `HLD_QBR_TEMPLATE_FILE = TEMPLATES_DIR / "hld_qbr_template.pptx"`. |
| `core/template_registry.py` | Register `"hld_qbr"` entry (id/name/description/icon) + `get_builder()` branch instantiating the new builder. |
| `core/builders/hld_qbr_builder.py` | The deep-clone engine: slide-archetype catalog (source index → archetype name), `_clone_slide`, `_clone_chart_part`, `_mutate_table_rows`, guidance-shape regex safety net, optional-slide selection logic, mandatory closing-slide append, final prune of original 60 template slides. |
| `llm/dynamic_layout_schemas.py` (extend) | New Pydantic item types the QBR archetypes need beyond existing ones: `OrgChartPerson`, `ActionTrackerRow`, `NonConformanceRow`, `KPIScorecardRow`, `VoiceOfCustomerQuote`, `MilestoneAchievement`. |
| `llm/prompts_hld_qbr.py` | System + planning prompt describing the QBR brand (UPS Healthcare palette/typography), the CORE/OPTIONAL archetype catalog from §2, and the mined guardrail constraints from slide 42/43 — explicitly instructing the LLM that "format option / example / optional" boxes are structure hints only and must not appear as literal output text. |
| `core/presentation_planner_hld_qbr.py` | Mirrors `presentation_planner_template1.py`: calls the LLM with the above prompts, returns a `DynamicPresentationPlan`, decides which `OPTIONAL` archetypes to include based on user's requested slide count/topics. |
| `app.py` (wire-up) | Add `"hld_qbr"` to the template dropdown; branch like the existing `template_id == "template1"` case to call `plan_hld_qbr_presentation(...)` then `get_builder("hld_qbr")`. |

---

## 6. Known blocker already solved during analysis

`python-pptx` refuses to open `.potx` directly (`ValueError: ... not a PowerPoint file,
content type is 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'`).
Verified fix: unzip → patch `[Content_Types].xml` replacing
`presentationml.template.main+xml` → `presentationml.presentation.main+xml` → rezip. This
exact step is step 1 of `convert_hld_qbr_template.py` (§5).

---

## 7. Open questions / assumptions to confirm before implementation

1. **Slide 10 "Voice of the Customer"** has no explicit Required/Optional guidance box —
   plan currently assumes **OPTIONAL**. Confirm.
2. **Slides 5 vs 6** ("[Customer] Priorities" Format Option 1 vs 2) are alternate visual
   treatments of the same required content — plan assumes the LLM/user picks exactly one,
   not both. Confirm.
3. **Slide 8** (hyperlinked "additional slides available" pointer, referencing an external
   Seismic library not present in this file) — plan treats this as pure guidance with nothing
   to clone (there is no actual slide content behind the hyperlink in this asset). Confirm no

---

## 8. KNOWN BUG (2026-09-23) — slide-count guidance only counts 6 of 16 archetypes

**Symptom:** generated decks always produce the same fixed content slides
(priorities/achievements/action_tracker/next_steps/+charts/kpi_tables) regardless of the
user's requested slide count or what the source document actually contains — the LLM never
picks `org_structure`, `voice_of_customer`, `gemba_walk`, `ci_tracker`,
`quality_org_structure`, `kpi_safety_quality`, `kpi_operational`, `nc_review_summary`, or
`nc_tracker`.

**Root cause:** `llm/prompts_hld_qbr.py::_build_slide_count_guidance()` hardcodes an exact
"Total content slides ... = N" arithmetic instruction using only 6 of the 16 optional
`HLDQBRPresentationPlan` fields as valid terms. The other 10 archetypes are only mentioned
once, earlier in the prompt, as an uncounted aside ("Optional if evidence exists: ...") — so
they never factor into how the LLM decides which slides to populate to hit the target count.
This is a prompt bug only; the schema, layout registry, and builder already fully support all
16 archetypes.

Full root-cause analysis + phased fix plan: see
[HLD_QBR_SLIDE_COUNT_ALLOCATION_BUG.md](HLD_QBR_SLIDE_COUNT_ALLOCATION_BUG.md).

**Status:** ✅ Fixed 2026-09-23 — `_build_slide_count_guidance()` in `llm/prompts_hld_qbr.py`
now generates the archetype list dynamically from `llm.hld_qbr_layout_registry.HLD_QBR_ARCHETYPES`
(plus the 4 fields not in the registry: `charts`, `kpi_tables`, `operational_chart`,
`nc_review_summary`) and uses evidence-first soft-target guidance instead of a strict
6-field equality constraint.
   further asset is expected.
4. Should `GUIDE_ONLY` appendix slides (icons, brand cheatsheet) be preserved *anywhere*
   (e.g. as a hidden reference tab in the app for the user) or fully deleted as this plan
   proposes? Deleting keeps the shipped template small and avoids ever accidentally leaking
   an icon-library slide into output.
5. Confirm target LLM output should keep tables non-scrolling (fixed max row counts likely
   needed per archetype, e.g. Action Tracker) — needs a per-archetype `max_rows` cap the
   planner must respect so cloned tables don't overflow the slide canvas.

---

## 8. Phased delivery roadmap

1. **Phase 1 — Template asset prep**: build `convert_hld_qbr_template.py`, produce and review
   the slimmed `templates/hld_qbr_template.pptx`, confirm zero guidance-box text remains and
   file size/slide count matches the CORE+OPTIONAL slide set from §2.
2. **Phase 2 — Static skeleton builder**: `hld_qbr_builder.py` that clones only the CORE
   slides intact (no dynamic content yet) + mandatory closing slide, registered in
   `template_registry.py` + `config.py`, selectable end-to-end from `app.py` and producing a
   valid `.pptx` — validates the OOXML clone plumbing (tables, charts, OLE cover) works before
   any LLM logic is added.
3. **Phase 3 — Dynamic content population**: schemas + prompts + planner
   (`dynamic_layout_schemas.py`, `prompts_hld_qbr.py`, `presentation_planner_hld_qbr.py`),
   wiring real per-archetype text/table/chart mutation into the Phase 2 builder.
4. **Phase 4 — Optional-slide selection & guardrail enforcement**: planner logic to pick
   relevant `OPTIONAL` archetypes based on source content/requested slide count, plus
   guardrail constraints (from slide 42/43) injected into the system prompt.
5. **Phase 5 — Validation**: a `scratch/` smoke-test script that runs a sample plan through
   the builder and asserts (a) no shape text matches the guidance regex, (b) no `GUIDE_ONLY`
   slide indices were cloned, (c) the closing slide is always last.
