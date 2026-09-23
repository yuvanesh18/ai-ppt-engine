# HLD QBR — "Fixed Slides Regardless of Requested Count" Bug

Status: ✅ Fixed 2026-09-23 (see §4/§5 — `_build_slide_count_guidance()` rewritten).

---

## 1. Symptom

Regardless of the user-requested slide count (or what's actually in the source document),
HLD QBR generation keeps producing the same fixed handful of content slides
(`priorities`, `achievements`, `action_tracker`, `next_steps`, plus `charts`/`kpi_tables` when
evidenced). Other archetypes the template and schema fully support —
`org_structure`, `voice_of_customer`, `gemba_walk`, `ci_tracker`, `quality_org_structure`,
`kpi_safety_quality`, `kpi_operational`, `nc_review_summary`, `nc_tracker` — are essentially
never populated, even when the source document clearly contains that kind of data.

---

## 2. Root cause

`HLDQBRPresentationPlan` ([llm/hld_qbr_schemas.py](llm/hld_qbr_schemas.py#L225)) supports
**16 optional content fields**, and the archetype registry
([llm/hld_qbr_layout_registry.py](llm/hld_qbr_layout_registry.py) `HLD_QBR_ARCHETYPES`) formally
documents 12 named archetypes with their own purpose/constraints — org_structure, achievements,
priorities, action_tracker, kpi_safety_quality, kpi_operational, voice_of_customer, gemba_walk,
ci_tracker, quality_org_structure, nc_tracker, next_steps.

However, the prompt instruction that actually drives how the LLM hits the user's requested
slide count — `_build_slide_count_guidance()` in
[llm/prompts_hld_qbr.py](llm/prompts_hld_qbr.py#L108-L121) — hardcodes the counting/allocation
pool to only **6 of those fields**:

```python
def _build_slide_count_guidance(requested_content_slide_count: Optional[int]) -> str:
    ...
    return (
        f"SLIDE TARGET: Allocate EXACTLY {requested_content_slide_count} content slides across:\n"
        f"- 'charts' (1 slide per chart)\n"
        f"- 'kpi_tables' (1 slide per table)\n"
        f"- 'priorities' (1 slide if populated)\n"
        f"- 'achievements' (1 slide if populated)\n"
        f"- 'action_tracker' (1 slide if populated)\n"
        f"- 'next_steps' (1 slide if populated)\n"
        f"Total content slides: len(charts) + len(kpi_tables) + non-chart slides = {requested_content_slide_count}.\n"
        "(Cover, Agenda, Executive Summary, and Closing are added automatically).\n"
    )
```

All the *other* archetypes (`org_structure`, `voice_of_customer`, `gemba_walk`, `ci_tracker`,
`quality_org_structure`, `kpi_safety_quality`, `kpi_operational`, `nc_review_summary`,
`nc_tracker`) are only mentioned once, earlier in the prompt, as a soft, uncounted aside:

> "Optional if evidence exists: org_structure [...], voice_of_customer {...}, gemba_walk [...],
> ci_tracker [...], nc_review_summary [...], nc_tracker [...]."

### Why the LLM ignores the other archetypes

The model is given a **strict arithmetic contract** ("Total content slides ... = N") that only
names 6 fields as valid terms in that equation. An LLM instructed to satisfy an exact-equality
constraint using a named, closed set of fields has no incentive — and an implicit
disincentive — to add slides from fields **outside** that equation, since doing so would
appear to break the "must equal N" instruction it was just given. The result is a
self-reinforcing bias toward the same 6 fields on every run, independent of:
- the actual requested slide count value, and
- what evidence actually exists in the source document.

This is a **prompt-engineering bug**, not a schema, builder, or validation bug — the schema
(`HLDQBRPresentationPlan`), the registry (`HLD_QBR_ARCHETYPES`), and the builder
(`core/builders/hld_qbr_builder.py`, which already renders all 12+ archetypes conditionally,
e.g. lines 770, 1449, 1469, 1481, 1501, 1520, 1534) are all fully capable of handling every
archetype — the LLM is just never told they count toward the target.

---

## 3. Archetype coverage table

| plan_field | Counted in slide-count guidance? | Rendered by builder? |
|---|---|---|
| `charts` | ✅ | ✅ |
| `kpi_tables` | ✅ | ✅ |
| `priorities` | ✅ | ✅ |
| `achievements` | ✅ | ✅ |
| `action_tracker` | ✅ | ✅ |
| `next_steps` | ✅ | ✅ |
| `org_structure` | ❌ | ✅ |
| `kpi_safety_quality` | ❌ | ✅ |
| `kpi_operational` | ❌ | ✅ |
| `voice_of_customer` | ❌ | ✅ |
| `gemba_walk` | ❌ | ✅ |
| `ci_tracker` | ❌ | ✅ |
| `quality_org_structure` | ❌ | ✅ |
| `nc_review_summary` | ❌ | ✅ |
| `nc_tracker` | ❌ | ✅ |

10 of 16 content archetypes are structurally excluded from ever contributing to the model's
"hit the target count" reasoning, even though every one of them is a real, always-available
slide in the template and builder.

---

## 4. Fix plan (not yet implemented)

**Phase 1 — Single source of truth for the archetype list.**
Generate the slide-count guidance text dynamically from
`llm.hld_qbr_layout_registry.HLD_QBR_ARCHETYPES` (already exists, already lists all 12 named
archetypes + `max_items`) instead of a hardcoded 6-field string in
`_build_slide_count_guidance()`. This guarantees the prompt and the registry can never drift
apart again.

**Phase 2 — Replace strict equality with evidence-first, soft-target guidance.**
Change the instruction from "must equal N exactly using these 6 fields" to something like:
> "Select every archetype below that has clear supporting evidence in the source document.
> Prioritize the ~{N} archetypes with the strongest evidence. Do not fabricate content for an
> archetype just to reach {N} — undershooting is acceptable and expected when the source lacks
> a given archetype's data (per existing repo policy, see `HLD_QBR_TEMPLATE_PLAN.md` §7)."

This preserves the existing, intentional design decision (documented in repo memory) that
undershooting the target is acceptable, while removing the artificial 6-field ceiling.

**Phase 3 — No schema/builder changes required.**
`core/presentation_planner_hld_qbr.py::_count_populated_content_slides()` already accounts for
all archetypes correctly (it's used for logging only) — confirms the bug is isolated to the
prompt text.

**Phase 4 — Regression test.**
Re-run the existing dry-run scripts (`scratch/hld_qbr/dry_run_ups_healthcare.py`,
`scratch/hld_qbr/dry_run_production_builder.py`) against a source document that contains
evidence for at least one currently-excluded archetype (e.g. a customer quote for
`voice_of_customer`, or named quality-team members for `quality_org_structure`) and confirm the
plan now populates it instead of silently dropping it.

**Phase 5 (optional, follow-up).**
Add a non-blocking governance warning in `core/governance.py` if the Content Model
(`content_traceability`) shows unused traceable items of a type that maps to a currently-empty
archetype (e.g. `content_type: "quote"` present but `voice_of_customer` empty) — surfaces
under-population without hard-failing the run.

---

## 5. Files to touch when implementing

- [llm/prompts_hld_qbr.py](llm/prompts_hld_qbr.py) — `_build_slide_count_guidance()` rewrite (Phases 1–2).
- No changes expected to `llm/hld_qbr_schemas.py`, `llm/hld_qbr_layout_registry.py`,
  `core/builders/hld_qbr_builder.py`, or `core/presentation_planner_hld_qbr.py`.
