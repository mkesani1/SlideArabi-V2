# SlideShift v2 — Engineering Documentation

**Version:** 2.1  
**Codebase:** `/home/user/workspace/slideshift_v2/`  
**Total source lines:** 12,176  
**Purpose:** Template-first, deterministic RTL (right-to-left) transformation engine that converts English PowerPoint presentations to Arabic RTL output.

---

## Table of Contents

1. [Architecture Overview](#1-architecture-overview)
2. [Module Reference](#2-module-reference)
3. [RTL Transformation Rules — The 19 Fixes](#3-rtl-transformation-rules--the-19-fixes)
4. [Pipeline Walkthrough](#4-pipeline-walkthrough)
5. [Template Registry](#5-template-registry)
6. [Property Resolution](#6-property-resolution)
7. [Known Limitations](#7-known-limitations)
8. [Key Implementation Notes](#8-key-implementation-notes)
9. [Configuration & Dependencies](#9-configuration--dependencies)
10. [Running the Pipeline](#10-running-the-pipeline)
11. [Phase 1b — Dual-LLM Translation Backend](#11-phase-1b--dual-llm-translation-backend)

---

## 1. Architecture Overview

### Design Philosophy

SlideShift v2 is built around three core principles:

1. **Template-first**: Every slide's layout type is detected first. RTL transforms are determined by layout rules, not by per-shape heuristics.
2. **Deterministic**: Given the same input, the pipeline always produces identical output. No loops, no random sampling, no AI-driven fixes during transformation.
3. **Phased DAG**: The pipeline runs as a directed acyclic graph of phases. Each phase executes exactly once and produces immutable output consumed by subsequent phases.

### High-Level Data Flow

```
English PPTX Input
       │
       ▼
┌─────────────────────────────────────┐
│  Phase 0: Parse & Resolve           │
│  PropertyResolver → ResolvedPresentation │
│  LayoutAnalyzer  → LayoutClassifications │
└─────────────────┬─────────────────┘
                  │
       ┌──────────▼───────────────────────┐
       │  Phase 1: Translate                  │
       │  ┌───────────────────────────────┐ │
       │  │ TokenProtector (pre-processing) │ │
       │  │ GPT-4o  EN→AR translation       │ │
       │  │ Claude Sonnet 4.5 QA pass       │ │
       │  │ Domain Glossary overrides       │ │
       │  └───────────────────────────────┘ │
       │  → translation_map                   │
       └──────────┬───────────────────────┘
                  │
       ┌──────────▼─────────────────────┐
       │  Phase 2: Transform Masters &   │
       │  Layouts (MasterLayoutTransformer)│
       │  Mutates: prs slide masters +   │
       │           slide layouts in-place│
       └──────────┬──────────────────────┘
                  │
       ┌──────────▼─────────────────────┐
       │  Phase 3: Transform Slide       │
       │  Content (SlideContentTransformer)│
       │  Mutates: slide shapes, text,   │
       │           tables, charts in-place│
       └──────────┬──────────────────────┘
                  │
       ┌──────────▼─────────────────────┐
       │  Phase 4: Typography            │
       │  Normalization (TypographyNormalizer)│
       │  Mutates: fonts, sizes, margins │
       └──────────┬──────────────────────┘
                  │
       ┌──────────▼─────────────────────┐
       │  Phase 5: Structural Validation │
       │  (StructuralValidator) — READ   │
       │  ONLY. Produces ValidationReport│
       └──────────┬──────────────────────┘
                  │
                  ▼
           Arabic PPTX Output
```

### Module Dependency Map

```
pipeline.py
 ├── property_resolver.py  (Phase 0)
 │    └── models.py
 ├── layout_analyzer.py    (Phase 0 / Phase 3)
 ├── template_registry.py  (Phase 2 / Phase 3)
 ├── rtl_transforms.py     (Phase 2 + Phase 3)
 │    └── utils.py
 ├── typography.py         (Phase 4)
 │    └── utils.py
 │    └── rtl_transforms.py (TransformReport)
 └── structural_validator.py (Phase 5)

process_single_deck.py  (standalone runner — bypasses pipeline.py)
 ├── layout_analyzer.py
 ├── template_registry.py
 ├── rtl_transforms.py
 ├── typography.py
 ├── property_resolver.py
 └── llm_translator.py      (Phase 1b — dual-LLM, activated by --llm-translate)
      ├── TokenProtector     (pre-processing: placeholder injection)
      ├── GPT-4o             (primary EN→AR translation)
      ├── ClaudeQAPass       (post-processing: pair review)
      └── DomainGlossary     (final deterministic overrides)

batch_process.py
 └── process_single_deck.py

smartart_translator.py  (called on saved PPTX file, post-pipeline)
test_harness.py         (testing / QA tooling)
```

### What Each Phase Owns

| Phase | Class | Scope | Mutates? |
|-------|-------|-------|----------|
| 0 | `PropertyResolver`, `LayoutAnalyzer` | Analysis only | No |
| 1 | (external translate_fn) | Text extraction | No |
| 1b | `LLMTranslator` (optional) | Dual-LLM EN→AR | No (produces translation_map) |
| 2 | `MasterLayoutTransformer` | Masters + layouts | Yes |
| 3 | `SlideContentTransformer` | Slide content | Yes |
| 4 | `TypographyNormalizer` | Fonts, sizes | Yes |
| 5 | `StructuralValidator` | Read-only check | No |

---

## 2. Module Reference

### 2.1 `models.py` (474 lines)

**Purpose:** Immutable data models representing the fully-resolved OOXML presentation structure. Frozen dataclasses guarantee no mutations after Phase 0.

**Key Types:**

| Class | Description |
|-------|-------------|
| `ResolvedRun` | A single text run with all formatting resolved (font size, name, bold, italic, color, underline, source level) |
| `ResolvedParagraph` | A paragraph with resolved alignment, RTL flag, level, bullet type, spacing, and a tuple of `ResolvedRun` |
| `ResolvedShape` | A shape with position (EMU), type, placeholder info, paragraphs, and a reference to the original lxml element |
| `ResolvedLayout` | A slide layout with its placeholders and freeform shapes |
| `ResolvedMaster` | A slide master with placeholders, freeforms, and txStyles dict |
| `ResolvedSlide` | A slide with slide number, layout info, and all resolved shapes |
| `ResolvedPresentation` | The complete immutable snapshot. `total_shapes` and `total_slides` properties. |
| `TransformAction` | A single atomic transform (mirror, swap, set_rtl, set_font, etc.) with shape_id and params |
| `TransformPlan` | A mutable plan containing lists of `TransformAction` keyed by slide, master, or layout |
| `ValidationIssue` | A single validation issue: severity, slide, shape, type, message, expected vs actual |
| `ValidationReport` | Full validation result with error/warning/info counts and `passed` bool |

**OOXML Constants:**
- Font sizes: hundredths of a point (`sz="1800"` = 18pt)
- EMU: 1 inch = 914,400 EMU; 1 pt = 12,700 EMU
- Alignment values: `'l'`, `'r'`, `'ctr'`, `'just'`, `'dist'`

**Design Note:** `ResolvedShape.original_xml_element` is excluded from `hash` and `eq` (via `field(compare=False, hash=False)`) because lxml elements are mutable and unhashable.

**Dependencies:** Standard library only (`dataclasses`, `typing`).

---

### 2.2 `utils.py` (537 lines)

**Purpose:** OOXML namespace helpers, coordinate math, text direction detection, and XML manipulation primitives. Used by `rtl_transforms.py`, `typography.py`, and `property_resolver.py`.

**Key Functions:**

| Function | Description |
|----------|-------------|
| `qn(tag)` | Converts `'a:pPr'` to `'{http://...}pPr'` (Clark notation). Supports `a:`, `p:`, `r:`, `c:` prefixes. |
| `mirror_x(x, w, slide_w)` | Core RTL formula: `slide_w - (x + w)`. Used everywhere for position mirroring. |
| `swap_positions(s1_x, s1_w, s2_x, s2_w, slide_w)` | Returns `(new_x1, new_x2)` by mirroring each shape's position through the other's position. Used for two-column swaps. |
| `has_arabic(text)` | Returns True if any character falls in Arabic Unicode ranges U+0600–U+06FF, U+0750–U+077F, U+FB50–U+FDFF, U+FE70–U+FEFF. |
| `has_latin(text)` | Returns True if any character is A–Z or a–z. |
| `is_bidi_text(text)` | True if text contains both Arabic and Latin characters. |
| `compute_script_ratio(text)` | Returns `{'arabic': float, 'latin': float, 'numeric': float, 'other': float}` summing to 1.0. Ignores whitespace. |
| `ensure_pPr(para_elem)` | Gets or creates `<a:pPr>` as the first child. Required before setting `rtl` or `algn` attributes. |
| `set_rtl_on_paragraph(para_elem)` | Sets `rtl='1'` on `<a:pPr>`. |
| `set_alignment_on_paragraph(para_elem, algn)` | Sets `algn` attribute on `<a:pPr>`. |
| `get_placeholder_info(shape)` | Returns `(ph_type_str, idx)` from python-pptx shape, or None. |
| `get_placeholder_info_from_xml(sp_elem)` | Same but from raw lxml element. Used when no python-pptx wrapper is available. |
| `set_body_pr_rtl_col(txBody)` | Sets `rtlCol='1'` on `<a:bodyPr>`. Controls column direction. |
| `set_defRPr_lang(txBody, lang)` | Sets `lang='ar-SA'` on all `<a:defRPr>` children. Enables Arabic font selection. |
| `bounds_check_emu(value, dim)` | Returns True if `value` is within `[-200_000, dim + 500_000]`. |
| `clamp_emu(value, dim)` | Clamps to `[-200_000, dim + 500_000]`. |
| `emu_to_pt`, `pt_to_emu`, `inches_to_emu`, etc. | Unit conversion helpers. |
| `hundredths_pt_to_pt(val)` | Converts OOXML font size format (e.g., `1800`) to points (`18.0`). |

**Arabic Unicode Ranges detected:**
```python
(0x0600, 0x06FF)  # Arabic
(0x0750, 0x077F)  # Arabic Supplement
(0xFB50, 0xFDFF)  # Arabic Presentation Forms-A
(0xFE70, 0xFEFF)  # Arabic Presentation Forms-B
```

**Dependencies:** `lxml.etree`, standard library.

---

### 2.3 `models.py` — see section 2.1

---

### 2.4 `property_resolver.py` (1,758 lines)

**Purpose:** Implements the 7-level OOXML property inheritance resolver. This is the most complex class in the system. It produces a `ResolvedPresentation` where every text property has a concrete non-None value.

**Key Class:** `PropertyResolver`

**Constructor:**
```python
PropertyResolver(presentation)  # python-pptx Presentation object
```

**Main Entry Point:**
```python
resolved = PropertyResolver(prs).resolve_presentation()
# Returns: ResolvedPresentation
```

**The 7-Level Inheritance Chain (checked in order, first non-None wins):**

| Level | OOXML Source | Description |
|-------|-------------|-------------|
| 1 | `a:rPr` on the run | Explicit run-level formatting |
| 2 | `a:pPr/a:defRPr` on the paragraph | Paragraph default run properties |
| 3 | `a:lstStyle` on the shape's `a:txBody` | Text frame list style |
| 4 | Shape's own inline / distinct `lstStyle` | Shape-level overrides |
| 5 | Layout placeholder `a:lstStyle` (matched by idx, then type) | Layout inheritance |
| 6 | Master placeholder `a:lstStyle` (matched by idx, then type) | Master inheritance |
| 7 | `p:txStyles` → `titleStyle`/`bodyStyle`/`otherStyle` | Master text style defaults |
| — | Hardcoded defaults | 18pt, Calibri, left, no RTL |

**Placeholder Matching Strategy:**
1. First match by `idx` attribute (exact index)
2. Fall back to matching by `type` attribute

**Theme Font Resolution:**
- `+mj-lt` → `a:majorFont/a:latin/@typeface`
- `+mn-lt` → `a:minorFont/a:latin/@typeface`
- `+mj-cs` → `a:majorFont/a:cs/@typeface`
- `+mn-cs` → `a:minorFont/a:cs/@typeface`
- Resolved from `a:theme → a:themeElements → a:fontScheme`

**Per-Property Resolvers:**

| Method | What it resolves |
|--------|-----------------|
| `resolve_font_size()` | Returns `(float_pt, source_level_str)` |
| `resolve_font_name()` | Returns font family string, resolves theme refs |
| `resolve_bold()` | Returns bool |
| `resolve_italic()` | Returns bool |
| `resolve_alignment()` | Returns one of `'l'`, `'r'`, `'ctr'`, `'just'`, `'dist'` |
| `_resolve_rtl()` | Returns bool |
| `_resolve_underline()` | Returns bool |
| `_resolve_color()` | Returns hex string or None |
| `_resolve_bullet_type()` | Returns `'char:•'`, `'auto:arabicPeriod'`, `'blip'`, or None |
| `_resolve_line_spacing()` | Returns multiplier float or None |
| `_resolve_space_before/after()` | Returns points or None |

**Shape Type Classification (`_classify_shape_type`):**
Inspects `shape._element.tag` to determine: `placeholder`, `textbox`, `picture`, `chart`, `table`, `group`, `connector`, `freeform`, `ole`, `smartart`, `media`.

**Local Position Override Detection (`_has_local_position_override`):**
A slide placeholder with explicit `p:spPr/a:xfrm/a:off` coordinates has a local position override. These are removed in Phase 3 so the shape inherits from the RTL-transformed layout.

**Dependencies:** `models.py`, `lxml`.

---

### 2.5 `layout_analyzer.py` (576 lines)

**Purpose:** Classifies each slide's layout into a canonical `ST_SlideLayoutType` string for deterministic transform lookup.

**Key Class:** `LayoutAnalyzer`

**Constructor:**
```python
LayoutAnalyzer(presentation)  # python-pptx Presentation
```

**Main Entry Point:**
```python
classifications = LayoutAnalyzer(prs).analyze_all()
# Returns: Dict[int, LayoutClassification]  (slide_number → classification)
```

**Classification Strategy (priority order):**

1. Read explicit `type` attribute from `<p:sldLayout type="twoColTx">` — confidence = 1.0
2. Infer from placeholder configuration using 19 heuristic rules (confidence 0.6–0.95)
3. Spatial analysis for two-column detection (both body placeholders in different halves)
4. Fall back to `'cust'` (confidence = 0.4) — flagged for AI classification

**Layout Inference Rules (in priority order):**

| Rule | Condition | Type | Confidence |
|------|-----------|------|-----------|
| 0 | No structural placeholders | `blank` | 0.95 |
| 1 | ctrTitle + subTitle | `title` | 0.95 |
| 2 | title only, no content | `titleOnly` | 0.90 |
| 3 | title + table | `tbl` | 0.90 |
| 4 | title + chart, no body | `chart` | 0.90 |
| 5 | title + diagram | `dgm` | 0.85 |
| 6 | title + body + chart | `txAndChart` | 0.85 |
| 7 | title + body + media | `txAndMedia` | 0.80 |
| 8 | title + body + clip art | `txAndClipArt` | 0.80 |
| 9 | title + picture | `picTx` | 0.85 |
| 10 | title + 2 bodies (spatial check) | `twoColTx` | 0.75–0.85 |
| 11 | title + 2 objects | `twoObj` | 0.85 |
| 12 | title + body + 2 objects | `txAndTwoObj` | 0.80 |
| 13 | title + body + object | `txAndObj` | 0.85 |
| 14 | 4 objects | `fourObj` | 0.85 |
| 15 | title + single body | `tx` | 0.90 |
| 16 | title + single object | `obj` | 0.85 |
| 17 | objects, no title | `objOnly` | 0.80 |
| 18 | body only (no title) | `tx` | 0.60 |
| 19 | fallback | `cust` | 0.40 |

**Decorative placeholder types** (`dt`, `ftr`, `sldNum`) are ignored when counting structural placeholders.

**Two-column spatial detection:** If two `body` placeholders exist and one has its center in the left half while the other has its center in the right half, it's classified as `twoColTx` with confidence 0.85 (vs 0.75 without spatial confirmation).

**AI Classification Flag:** Slides with confidence below `0.7` have `requires_ai_classification = True`.

**Caching:** Layout classifications are cached by `id(layout)` to avoid re-classifying shared layouts.

**Dependencies:** Standard library, python-pptx.

---

### 2.6 `template_registry.py` (1,193 lines)

**Purpose:** The "brain" of the deterministic pipeline. Maps each of the 36 standard `ST_SlideLayoutType` values to prescriptive RTL transformation rules.

**Key Classes:**

**`PlaceholderAction`** — a rule for a single placeholder:
- `action`: `'keep_centered'`, `'right_align'`, `'mirror'`, `'swap_with_partner'`, `'keep_position'`, `'reverse_columns'`, `'reverse_axes'`
- `set_rtl`: Whether to set RTL direction
- `set_alignment`: Target alignment (`'r'`, `'l'`, `'ctr'`, or None)
- `swap_partner_idx`: Placeholder idx to swap with
- `mirror_x`: Whether to mirror X coordinate
- `notes`: Human-readable reason

**`LayoutTransformRules`** — complete ruleset for a layout type:
- `layout_type`: Canonical layout type string
- `description`: Human-readable description
- `placeholder_rules`: Dict of `ph_type` or `idx_N` → `PlaceholderAction`
- `freeform_action`: `'mirror'`, `'keep'`, or `'analyze'`
- `mirror_master_elements`: Whether to mirror inherited master shapes
- `swap_columns`: Whether this layout has columns to swap
- `table_action`: `'reverse_columns'`, `'keep'`, or `'rtl_only'`
- `chart_action`: `'reverse_axes'`, `'keep'`, or `'mirror_legend'`

**`TemplateRegistry`** — the registry itself:

```python
registry = TemplateRegistry(slide_width_emu, slide_height_emu)
rules = registry.get_rules('twoColTx')
action = registry.get_placeholder_action('twoColTx', 'body', idx=1)
registry.register_custom_rule('my_layout', custom_rules)
```

**Placeholder Action Lookup Order:**
1. `idx_N` key (most specific — by placeholder index number)
2. `placeholder_type` key (by semantic type)
3. Default: `keep_position` with RTL=True, alignment=`'r'`

**Custom Rules:** Enterprise layouts can be registered at runtime via `register_custom_rule()`. Custom rules take precedence over built-in defaults.

**Font Map (`ARABIC_FONT_MAP`):**
A dictionary mapping Latin/Western fonts to Arabic-capable equivalents. Used in both `template_registry.py` and `typography.py`.

| Latin Font | Arabic Substitute |
|-----------|------------------|
| Calibri, Arial, Times New Roman, Tahoma, Segoe UI | Same (have Arabic glyphs) |
| Cambria, Georgia | Sakkal Majalla |
| Verdana, Trebuchet MS | Tahoma |
| Garamond, Palatino | Traditional Arabic |
| Century Gothic, Futura | Dubai |
| Helvetica, Helvetica Neue | Arial |
| Courier New, Consolas | Courier New |

**Dependencies:** Standard library.

---

### 2.7 `rtl_transforms.py` (3,106 lines)

**Purpose:** The core transformation engine. Contains `MasterLayoutTransformer` (Phase 2) and `SlideContentTransformer` (Phase 3) plus `TransformReport`.

**Key Class: `TransformReport`**

Mutable summary produced by each transformation phase:
```python
report.add('change_type', count)   # increment counter
report.warn('message')              # log warning
report.error('message')             # log error
report.merge(other_report)          # combine reports
# Properties: total_changes, changes_by_type, warnings, errors
```

---

#### `MasterLayoutTransformer` (Phase 2)

**What it does to masters:**
1. Sets `rtlCol='1'` on all `<a:bodyPr>` elements (safe — controls column direction only)
2. Sets `lang='ar-SA'` on all `<a:defRPr>` elements (enables Arabic font selection)
3. Sets `lang='ar-SA'` in txStyles (`titleStyle`, `bodyStyle`, `otherStyle`) level paragraphs
4. Mirrors small logo images (`<p:pic>` shapes with width < 20% of slide width) — position only, no `flipH`
5. Mirrors small brand text elements (non-placeholder text shapes with width < 30% of slide width)

**What it does NOT do (critical design constraints):**
- Does NOT set `rtl='1'` at master or layout level — this corrupts English text by triggering the bidi algorithm on all text
- Does NOT set `algn` at master/layout level — alignment is context-sensitive and must be set at slide paragraph level
- Does NOT apply `flipH` to any shape — this inverts images and corrupts logos

**What it does to layouts:**
1. Applies same RTL direction defaults as master
2. Mirrors all placeholder X positions for standard layouts
3. For two-column layouts (`twoColTx`, `twoObj`, `twoTxTwoObj`, `txAndChart`, `chartAndTx`, `picTx`): swaps the leftmost and rightmost content placeholders

**Two-column swap algorithm:**
- Title placeholders (idx=0) are mirrored individually
- Content placeholders are sorted left-to-right by X position
- The leftmost and rightmost are swapped using `swap_positions()`
- Any remaining middle placeholders are mirrored individually

---

#### `SlideContentTransformer` (Phase 3)

**Per-slide transformation sequence:**
1. Collect all top-level shapes (groups as single units)
2. For each shape:
   - If **placeholder**: remove local `xfrm` element so it inherits from the RTL-transformed layout (Round 4: check for dangerous overlap first; if overlap detected, mirror explicitly and shrink to fit)
   - If **non-placeholder (freeform)**: mirror X position if `_should_mirror_shape()` returns True; reverse directional shapes (Fix 10)
   - If has text frame: apply translation + `_set_rtl_alignment_unconditional()`
   - Validate text box width for short text (Fix 8)
   - If has table: transform table RTL
   - If has chart: transform chart RTL
3. Fix 11: Fix `wrap="none"` text boxes for Arabic text (change to `square`, add normAutofit)
4. Fix 12: Resolve title-body vertical overlap for RTL
5. Fix 14: Right-anchor cover title text on photo-background covers
6. Fix 15: Mirror split-panel layouts (image+text side-by-side)
7. Fix 16: Reverse timeline alternation pattern
8. Fix 17: Reverse logo row ordering
9. Fix 18: Center text in circular/bounded container shapes
10. Fix 19: Ensure bidi base direction on mixed Arabic/English titles
11. Fix 9: Collision detection (logging only)

**RTL direction rule (CRITICAL — Round 2 fix):**
- `rtl='1'` is ONLY set on paragraphs that contain Arabic characters
- English-only paragraphs get `rtl='0'` explicitly
- Setting `rtl='1'` on English text causes OOXML bidi algorithm to reorder characters, move periods to the start of lines, and corrupt word order

**Text translation algorithm:**
1. For each paragraph, gather text across all runs
2. Look up in translation map (fuzzy matching: exact → stripped → case-insensitive → 80% prefix match for texts > 40 chars)
3. If found: put Arabic text in first run, clear subsequent runs (preserves first-run formatting)
4. Set `lang='ar-SA'` on all run `<a:rPr>` elements
5. Set `rtl='1'` and `algn` on `<a:pPr>`

**Alignment rules:**
- Footer placeholders (`ftr`, `sldNum`, `dt`) → always `'l'`
- Title placeholders (except `subTitle`) → `'ctr'`
- Everything else with Arabic content → `'r'`

**Table RTL transformation:**
1. Translate cell text (using same fuzzy lookup)
2. Reverse column widths in `<a:tblGrid>`
3. Deep copy and reverse `<a:tc>` elements within each `<a:tr>`
4. Set `rtl='1'` on `<a:tblPr>`
5. Set RTL paragraph properties on all cell text
   - Numeric cells (>80% digits/currency) → `algn='l'` (numbers read LTR)
   - All other cells → `algn='r'`

**Chart RTL transformation:**
1. Reverse category axis: set `<c:crosses val="max"/>` and `<c:orientation val="maxMin"/>` on all `<c:catAx>`
2. Mirror legend position: `r↔l`, `tr↔tl`
3. Move value axis to opposite side: `l↔r`, `t↔b`
4. Reverse value axis direction for bar charts (Fix 7): toggle `<c:orientation val="maxMin/minMax">`
5. Translate chart category labels and series names (Round 3)
6. Apply month name corrections for known Google Translate errors (e.g., `'يمشي'` → `'مارس'` for "March")
7. Set RTL on chart title text

**Freeform shape mirroring decision (`_should_mirror_shape`):**
- Full-width shapes (> 90% slide width): never mirror (background decoratives)
- Footer-zone shapes (bottom > 88% of slide height): always mirror
- `secHead`/`title` layouts: only skip shapes > 50% slide width; mirror smaller shapes
- Default: mirror all content shapes

**Overlap detection for placeholder inheritance:**
Two-tier check before removing xfrm:
- **Tier 1:** Large shapes (> 30% slide width) that would overlap > 20% of the placeholder area
- **Tier 2:** Small images in the title zone (top 25%) whose mirrored position would overlap the placeholder

If dangerous overlap is detected: compute a "safe zone" between all logos, shrink the placeholder to fit, and enable auto-fit (normAutofit + direct font scaling for LibreOffice compatibility).

**Dependencies:** `utils.py`, `lxml`, `copy.deepcopy`.

---

### 2.8 `typography.py` (924 lines)

**Purpose:** Phase 4 typography normalization for Arabic. Runs AFTER Phase 3 has inserted Arabic text and repositioned shapes.

**Key Class:** `TypographyNormalizer`

**Constructor:**
```python
TypographyNormalizer(presentation)
```

**Main Entry Point:**
```python
report = TypographyNormalizer(prs).normalize_all()
```

**Operations per shape:**

1. **Font mapping** (`_apply_font_mapping`):
   - Replaces `<a:latin typeface="...">` in every `<a:rPr>` with the Arabic-capable substitute
   - For Arabic-content runs: ensures `<a:cs>` (complex script) element exists with an Arabic-capable font
   - Sets `lang='ar-SA'` on `<a:rPr>` for Arabic runs
   - Lookup: exact → case-insensitive → prefix match (e.g., `'Calibri Bold'` matches `'Calibri'`)

2. **Text frame margins** (`_set_text_frame_margins`):
   - Only applied to shapes containing Arabic text
   - Sets `lIns/rIns = 27432 EMU (~0.03")` and `tIns/bIns = 45720 EMU (~0.05")`
   - Arabic glyphs are taller (larger ascenders/descenders) than Latin equivalents

3. **Bidirectional formatting** (`_apply_bidi_formatting`):
   - Pure Arabic (>70% Arabic chars): `rtl='1'`, `algn='r'`, `lang='ar-SA'` on all runs
   - Mixed bidi (Arabic + Latin): `rtl='1'`, `algn='r'`, inject RLM (U+200F) at paragraph start
   - Latin in Arabic-context frame: `rtl='1'`, `algn='r'` (visual consistency)
   - Pure LTR in non-Arabic frame: `rtl='0'`, `algn='l'`
   - Footer placeholders: `rtl='0'`, `algn='l'`
   - `ctrTitle`: `rtl='1'`, `algn='ctr'`
   - Sets `rtlCol='1'` on `<a:bodyPr>` for frames containing Arabic

4. **Overflow detection** (`_check_text_overflow`):
   Heuristic: estimates `character_count × char_width_factor × font_size` per line, wraps at frame width, compares total estimated height to frame height with 5% tolerance.
   - Arabic char width factor: 0.65 × font_size
   - Latin char width factor: 0.55 × font_size

5. **Font size reduction** (`_reduce_font_size_to_fit`):
   If overflow detected: applies single proportional reduction (default max 20%). Enforces floors:
   - Title placeholders: min 14pt
   - Body / non-placeholder: min 10pt
   - Table cells: min 9pt (with max 15% reduction)
   - After reduction, if still overflowing: logs warning for manual review (no infinite loops)

**Arabic expansion factors (per-font):**

| Font | Factor |
|------|--------|
| Calibri | 1.20 |
| Arial | 1.22 |
| Times New Roman | 1.28 |
| Sakkal Majalla | 1.15 |
| Dubai | 1.18 |
| Traditional Arabic | 1.30 |
| Default | 1.25 |

**Font map:** See `ARABIC_FONT_MAP` constant (40+ entries mapping Western fonts to Arabic-capable equivalents).

**Dependencies:** `utils.py`, `rtl_transforms.TransformReport`, `lxml`.

---

### 2.9 `structural_validator.py` (505 lines)

**Purpose:** Phase 5 read-only validation. Checks the transformed presentation and reports issues without modifying anything. Replaces the VQA fix loops from v1.

**Key Class:** `StructuralValidator`

**Main Entry Point:**
```python
report = StructuralValidator(prs, resolved_prs).validate()
# Returns: ValidationReport
# report.passed == True if no errors (warnings are acceptable)
```

**Issue Types (`IssueType`):**

| Type | Severity | What it checks |
|------|----------|---------------|
| `rtl_missing` | error (>50% Arabic) / warning (mixed) | Arabic paragraph without `rtl='1'` |
| `alignment_wrong` | error (left) / warning (center in body) | Arabic text with wrong alignment |
| `shape_out_of_bounds` | error (content) / warning (decorative) | Shape extends past slide boundaries (±100,000 EMU tolerance) |
| `shape_overlap` | error (both content) / warning (other) | Shapes overlapping 10–90% (ignores intentional layering >90%) |
| `font_too_small` | error (<8pt) / warning (<10pt body, <14pt title) | Font below minimum readable size |
| `table_not_rtl` | error | Table missing `rtl='1'` on `<a:tblPr>` |
| `chart_not_rtl` | error | Chart category axis not reversed (`<c:orientation val="maxMin">` missing) |
| `placeholder_mismatch` | warning | Placeholder position differs from layout by >10,000 EMU |
| `master_no_rtl` | warning | Master `txStyles` level 1 paragraph missing `rtl='1'` |

**Overlap detection:**
- Computes intersection area of all non-line, non-background shape pairs
- Reports overlap as issue only when 10% < overlap% < 90% AND both shapes are content (text/table/chart/picture)
- Overlap < 10%: likely accidental minor overlap, ignored
- Overlap > 90%: likely intentional layering (text on colored box), ignored

**Pass/fail criterion:** `ValidationReport.passed` is True if there are zero errors. Warnings do not fail the pipeline.

**Dependencies:** python-pptx, `re`, standard library.

---

### 2.10 `smartart_translator.py` (217 lines)

**Purpose:** Translates text in SmartArt/diagram shapes in a saved PPTX file. SmartArt text lives in separate XML parts (`ppt/diagrams/dataX.xml` and `ppt/diagrams/drawingX.xml`) not exposed by the python-pptx shape API.

**Key Function:**
```python
count = translate_smartart_in_pptx(pptx_path, translations)
```

**Algorithm:**
1. Opens PPTX as a ZIP file (PPTX is a ZIP archive)
2. Finds all files in `ppt/diagrams/` ending in `.xml` that contain `/data` or `/drawing` in their path
3. For each diagram XML: iterates all `<a:p>` paragraph elements
4. Gathers multi-run text into single paragraph text (handles split runs like `"Weaponised D" + "rones"`)
5. Looks up translation using fuzzy matching (exact → stripped → case-insensitive)
6. If found: puts Arabic text in first run, clears subsequent runs, sets `lang='ar-SA'` on all `<a:rPr>`, sets `rtl='1'` on `<a:pPr>`
7. Writes modified parts back by rewriting the ZIP with `ZIP_STORED` compression
8. Uses a temporary file to prevent corruption on write failure

**Note:** Works on a file path (not a python-pptx object) because python-pptx does not expose SmartArt XML parts through its API.

**Dependencies:** `zipfile`, `lxml`, `shutil`, `pathlib`.

---

### 2.11 `pipeline.py` (322 lines)

**Purpose:** The orchestrating pipeline class. Defines `SlideShiftV2Pipeline` with `PipelineConfig` and `PipelineResult`. In practice, `process_single_deck.py` bypasses this class and calls the individual transformers directly (see section 10).

**Key Classes:**

**`PipelineConfig`:**
- `input_path`: Source PPTX file
- `output_path`: Output PPTX file
- `translate_fn`: Optional callable `(List[str]) → Dict[str, str]`
- `skip_translation`: Skip Phase 1 entirely
- `max_font_reduction_pct`: Max font shrinkage in Phase 4 (default 20%)
- `log_level`: Logging level string

**`PipelineResult`:**
- `success`: bool
- `output_path`: str or None
- `phase_reports`: Dict of per-phase reports
- `validation_report`: Phase 5 result
- `total_duration_ms`: Wall-clock time
- `error`: Error message if failed

**Phase execution:**
Each phase is wrapped in try/except. If a module isn't available (e.g., during testing), it gracefully stubs out with a warning rather than crashing. The pipeline does NOT fail on validation errors — it reports them and saves regardless.

---

### 2.12 `test_harness.py` (1,503 lines)

**Purpose:** QA and structural comparison testing. Runs the full pipeline on one or more input PPTX files and produces per-deck HTML comparison reports with side-by-side slide images.

**Usage:**
```bash
python slideshift_v2/test_harness.py input.pptx
python slideshift_v2/test_harness.py --dir /path/to/test_decks/
```

**Output per deck:**
```
test_results/
└── myfile/
    ├── analysis.json              # Full structural analysis
    ├── myfile_v2_structural.pptx  # Transformed output
    ├── original/slide-01.jpg ...  # Rendered original slides
    ├── v2/slide-01.jpg ...        # Rendered transformed slides
    ├── comparison/slide-01.jpg ... # Side-by-side comparison images
    └── report.html                # Full HTML comparison report
```

**VQA integration:** The test harness uses LibreOffice (`soffice --headless`) to render slides to PNG/JPG for visual comparison. `pdftoppm` is used as an alternative renderer when LibreOffice produces blank output.

**Structural analysis (`_analyze_shape`):** Extracts per-shape data including position (EMU and inches), placeholder info, paragraph text, RTL flags, alignment, font properties, table/chart presence.

---

### 2.13 `llm_translator.py` (1,061 lines)

**Purpose:** Implements the dual-LLM translation backend for Phase 1b. Replaces the curl-based Google Translate calls when activated via `--llm-translate`. Produces a translation map in the same JSON format as existing translation caches, so the downstream pipeline requires no changes.

**Key Classes:**

| Class | Description |
|-------|-------------|
| `TokenProtector` | Pre-processing layer. Replaces protected tokens with `⟦PROTXXXX⟧` placeholders before GPT sees the text, then restores them after translation. |
| `LLMTranslator` | Orchestrates the full three-layer pipeline: TokenProtector → GPT-4o translation → Claude QA pass → domain glossary overrides. |
| `ClaudeQAPass` | Post-processing layer. Submits EN→AR pairs to Claude Sonnet 4.5 for terminology and semantic review. |
| `DomainGlossary` | Deterministic final safety net. Replaces known catastrophic mistranslations via a hardcoded `BAD_TRANSLATIONS` dict. |

**Constructor:**
```python
LLMTranslator(
    openai_api_key: str,
    anthropic_api_key: str | None = None,  # If absent, Claude QA pass is skipped
    enable_qa: bool = True,
)
```

**Main Entry Point:**
```python
translation_map = LLMTranslator(...).translate(strings: List[str]) -> Dict[str, str]
```

**Three Defensive Layers:**

#### Layer 1: `TokenProtector` (pre-processing)

Replaces tokens that must not be translated with opaque `⟦PROTXXXX⟧` placeholders before the text reaches GPT-4o. Restores originals after translation.

**Protected token categories:**

| Category | Examples |
|----------|----------|
| Abbreviations (176 entries) | `HW`, `SW`, `GDPR`, `HIPAA`, `EBITDA`, `CAPEX`, `OPEX` |
| Quarter/year identifiers | `Q1-24`, `Q3 2025`, `FY2025`, `FY26` |
| Number + unit patterns | `17.4M`, `$500K`, `£2.1B`, `500ms`, `16GB` |
| URLs and email addresses | `https://...`, `user@domain.com` |

**Algorithm:**
1. Scan text with regex patterns and exact-match set (in priority order: longest tokens first to prevent partial matches)
2. Replace each match with `⟦PROT0001⟧`, `⟦PROT0002⟧`, etc., maintaining a restoration map
3. After GPT translation returns Arabic text, restore all `⟦PROTXXXX⟧` tokens to originals
4. Guarantees that abbreviations, financial codes, and technical identifiers pass through untranslated

#### Layer 2: `ClaudeQAPass` (post-processing)

Reviews all EN→AR translation pairs in batches of 60 using Claude Sonnet 4.5. Identifies and corrects six issue categories.

**Issue categories checked:**

| Category | Example |
|----------|----------|
| Abbreviation mangling | `OS` → "platforms" (should stay `OS` or "نظام التشغيل") |
| Brand name corruption | `SlideShift` rendered in Arabic script rather than preserved as-is |
| Number/unit errors | `$500K` changed to a different value or unit |
| Semantic inversions | A negation dropped or inverted, changing meaning |
| Terminology inconsistency | Same English term translated differently across slides |
| Professional register | Informal phrasing where formal executive register is required |

**Output format:** Claude returns a JSON array of correction objects — only strings where issues were found are returned. Strings without issues are implicitly approved.

**Live test results (R6_07 MR Business Vision deck, 61 strings):**
- 5 semantic errors caught
- 5 terminology inconsistencies caught
- 4 register/grammar issues caught
- 1 abbreviation issue caught
- Notable corrections: `OS` → "platforms" fixed, omitted word "Modular" restored, gender agreement errors corrected, "at cost"→"with cost" business meaning restored

#### Layer 3: `DomainGlossary` (final safety net)

A hardcoded `BAD_TRANSLATIONS` dictionary that catches known catastrophic Google Translate / GPT failures and replaces them with correct domain-appropriate Arabic. Applied as the final step after Claude QA.

**Example overrides:**

| Bad Arabic (literal meaning) | Correct Arabic (intended meaning) |
|-------------------------------|-----------------------------------|
| `المخلفات الخطرة` (hazardous waste) | `الأجهزة` (hardware) |
| `مؤتمن` (trustworthy) | `سري` (confidential) |
| `الناتج المحلي الإجمالي` (GDP) | `اللائحة العامة لحماية البيانات` (GDPR) |
| `خط أنابيب الإيرادات` (literal pipeline) | `مسار الإيرادات` (revenue pipeline) |

**System Prompts:**

| Prompt | Size | Key instructions |
|--------|------|------------------|
| GPT system prompt | 2,013 chars | Rules for placeholder tokens, brand name preservation, financial terminology, Arabic numeral format, paragraph-level JSON response format |
| Claude QA prompt | 1,536 chars | Issue category definitions, JSON output schema, scoring criteria, instruction to return only corrections |

**Batching & Performance:**

| Parameter | Value |
|-----------|-------|
| GPT batch size | 40 strings per request |
| GPT temperature | 0.1 (near-deterministic) |
| GPT response format | `{ type: "json_object" }` |
| Claude batch size | 60 pairs per request |
| Retry policy | 2 retries per batch, exponential backoff |
| Caching | JSON dict, same format as existing translation caches |
| Approximate cost | ~$0.09/deck (GPT $0.02 + Claude $0.07) |
| Approximate latency | ~73 seconds per deck |

**Translation Quality Benchmarks (R6_07 MR Business Vision):**

| Approach | Quality Score |
|----------|---------------|
| Google Translate (GTX) baseline | 5.5 / 10 |
| GPT-4o only | 8.8 / 10 |
| GPT-4o + Claude QA | **9.5 / 10** |

All Category A failures eliminated (HW→"hazardous waste", brand name corruption, semantic inversions).

**Model selection rationale:**

- **GPT-4o** was chosen as primary translator based on benchmarks: BLEU 54.2, COMET 0.847, Human Accuracy 4.3/5 — strongest overall EN→AR quality among frontier models tested.
- **Claude Sonnet 4.5** was chosen as QA reviewer to provide independent verification. A model reviewing its own output catches fewer errors than a different model reviewing the same output.
- **Deterministic glossary** provides a hard guarantee that both LLMs cannot jointly miss a known catastrophic failure.

**Dependencies:** `openai`, `anthropic`, `json`, `re`, standard library.

---

## 3. RTL Transformation Rules — The 19 Fixes

All 19 fixes are implemented in `SlideContentTransformer` in `rtl_transforms.py`. They are applied in order during `_transform_slide()`.

### Fix 1 — Unconditional RTL Alignment Pass

**Method:** `_set_rtl_alignment_unconditional(shape)`  
**What it detects:** Any text-bearing shape on any slide  
**What it transforms:**
- Paragraphs with Arabic characters: `rtl='1'`, alignment per `_compute_paragraph_alignment()`
- Paragraphs without Arabic: explicit `rtl='0'` (prevents inherited `rtl='1'` from corrupting English text), alignment per `_compute_paragraph_alignment()`
- Sets `rtlCol='1'` on `<a:bodyPr>` only if the frame has Arabic content  
**Applied to:** All slides (called after `_apply_translation` on every shape)  
**Critical note:** This is the fix for the P0 bug where English text with inherited `rtl='1'` had periods displaced to the start of lines and words reordered.

---

### Fix 2 — secHead/Title Layout Freeform Decision

**Method:** `_should_mirror_shape(shape, layout_type)`  
**What it detects:** Freeform shapes on `secHead` or `title` layouts  
**What it transforms:** Only skips shapes > 50% of slide width (large decorative backgrounds). Mirrors all other shapes including text boxes.  
**Applied to:** All slides with `secHead` or `title` layouts  
**Note:** v1 incorrectly skipped ALL freeforms on these layouts; v2 only skips large backgrounds.

---

### Fix 3 — Footer-Zone Shape Mirroring

**Method:** `_should_mirror_shape(shape, layout_type)`  
**What it detects:** Non-placeholder shapes whose bottom edge is below 88% of slide height  
**What it transforms:** Always mirrors these shapes regardless of layout type  
**Applied to:** All slides  
**Note:** Footer zone shapes (page numbers, bottom decoratives) must always mirror to maintain visual alignment.

---

### Fix 4 — Anti-flipH Protection

**Method:** `_ensure_no_content_flip(shape)`  
**What it detects:** `flipH` or `flipV` attributes on any shape's `<a:xfrm>` element  
**What it transforms:** Removes `flipH` and `flipV` attributes after mirroring  
**Applied to:** All slides, called inside `_mirror_freeform_shape()`  
**Note:** Content images (maps, photos) must NEVER be content-flipped — only repositioned.

---

### Fix 5 — Group Shape Handling

**Method:** Group shape branch in `_transform_slide()`  
**What it detects:** Shapes with a `shapes` attribute (group shapes) that are not placeholders  
**What it transforms:**
- Mirrors the group's bounding box position as a single unit
- Recursively processes text children for translation and RTL alignment (without mirroring child positions — child coordinates are relative to the group's local space)
- Handles tables and charts inside groups  
**Applied to:** All slides  
**Note:** v1 iterated group children for position mirroring, causing double-offset bugs.

---

### Fix 6 — Placeholder Position Inheritance Override (Collision-Aware)

**Method:** `_remove_local_position_override(shape, layout)`  
**What it detects:** Placeholder shapes with explicit local `<a:xfrm>` (overriding layout position) where the inherited position would be safe  
**What it transforms:**
- Safe case: removes the xfrm element entirely — shape re-inherits from the (already RTL-mirrored) layout
- Dangerous overlap case: explicitly mirrors the placeholder, then computes a "safe zone" between any logos and shrinks the placeholder to fit  
**Applied to:** All slides  
**Note:** The two-tier overlap check (large shapes >30% width, then small images in title zone) was added in Round 4 to fix title-logo collisions in R6_17.

---

### Fix 7 — Bar Chart Value Axis Reversal

**Method:** `_transform_chart_rtl(shape)` step 4  
**What it detects:** `<c:valAx>` elements in chart XML  
**What it transforms:** Toggles `<c:orientation val="maxMin/minMax">` on value axes  
**Applied to:** All slides with chart shapes  
**Note:** Makes horizontal bars grow right-to-left as expected in Arabic.

---

### Fix 8 — Text Box Width Enforcement

**Method:** `_validate_textbox_width(shape)`  
**What it detects:** Text boxes with short content (≤ 5 characters) narrower than 457,200 EMU (0.5 inches)  
**What it transforms:** Sets `shape.width = 457200` minimum  
**Applied to:** All slides  
**Note:** Prevents two-digit page numbers from wrapping after mirroring narrows the text box.

---

### Fix 9 — Collision Detection (Diagnostic)

**Method:** `_detect_collisions(shapes, slide_number)`  
**What it detects:** Shapes with overlapping bounding boxes (intersection area > 2 sq inches)  
**What it transforms:** Logs warnings only — no modifications  
**Applied to:** All slides with ≤ 50 shapes (performance guard)

---

### Fix 10 — Directional Shape Reversal

**Method:** `_reverse_directional_shape(shape)`  
**What it detects:** Shapes with directional preset geometry (`rightArrow`, `leftArrow`, `chevron`, etc.)  
**What it transforms:**
- Type-swap pairs: `rightArrow` ↔ `leftArrow`, `curvedRightArrow` ↔ `curvedLeftArrow`, etc.
- FlipH-toggle shapes: `chevron`, `homePlate`, `notchedRightArrow`, `bentArrow`, `pentagon`, etc. — toggles `flipH` attribute  
**Applied to:** All slides, all non-placeholder shapes  
**Note:** Unlike Fix 4 (which removes flipH), Fix 10 intentionally sets flipH on arrow/chevron shapes because they are directional symbols, not content images.

---

### Fix 11 — wrap="none" Text Box Fix

**Method:** `_fix_wrap_none_for_arabic(shape)`  
**What it detects:** Text boxes with `<a:bodyPr wrap="none">` that contain Arabic text  
**What it transforms:**
1. Changes `wrap="none"` → `wrap="square"`
2. Removes `spAutoFit`/`noAutofit`, adds `normAutofit` with `fontScale="80000"` (allow 80% shrink)
3. Expands text box width to fill available panel space  
**Applied to:** All slides (second pass after main shape loop)

---

### Fix 12 — Title-Body Vertical Overlap Resolution

**Method:** `_fix_title_body_overlap(shapes, slide_number)`  
**What it detects:** Body-like shapes whose `top` is within a title-like shape's vertical extent (titles: font ≥ 40pt, in top 30% of slide)  
**What it transforms:** Moves the overlapping body shape down so its top = `title_bottom + 91,440 EMU (0.1 inch gap)`  
**Applied to:** All slides with text shapes  
**Note:** In LTR, overlapping title/body boxes coexist because text flows in different horizontal zones. In RTL, both flow from the right edge, causing visual collision.

---

### Fix 13 — Slide-Level Logo Mirror

**Method:** `_mirror_slide_level_logos(shapes)`  
**What it detects:** Small picture elements (`<p:pic>`, width < 20% slide width, non-placeholder, has blipFill) placed directly on slides  
**What it transforms:** Mirrors X position  
**Applied to:** All slides  
**Note:** The `MasterLayoutTransformer` handles master/layout logos. This fix handles logos placed directly on content slides.

---

### Fix 14 — Photo-Background Cover Title Anchor

**Method:** `_fix_cover_title_anchor(shapes, slide_number)`  
**What it detects:** Slides with a large background image (>50% of slide area) AND an Arabic title text box in the LEFT half  
**What it transforms:** Mirrors the title text box to the right half  
**Applied to:** All slides  
**Affected decks:** R6_04, R6_13, R6_14

---

### Fix 15 — Split-Panel Layout Mirror

**Method:** `_mirror_split_panel_layout(shapes, slide_number)`  
**What it detects:** Exactly 1 large shape on the left half AND 1 large shape on the right half (both > 35% slide width, > 50% slide height), where one is an image and the other is not  
**What it transforms:** Mirrors both panels' X positions so the image moves from left to right (RTL convention)  
**Applied to:** All slides  
**Affected decks:** R6_15, R6_16, R6_18 (photo-left/text-right cover slides)

---

### Fix 16 — Timeline Alternation Reversal

**Method:** `_reverse_timeline_alternation(shapes, slide_number)`  
**What it detects:** Slides with a central vertical axis (connector or narrow tall shape in 33–67% horizontal zone) and text shapes alternating on both sides  
**What it transforms:** Pairs text shapes at similar Y positions on opposite sides of the axis and swaps their X positions  
**Applied to:** All slides  
**Affected decks:** R6_10 slide 15

---

### Fix 17 — Logo Row Order Reversal

**Method:** `_reverse_logo_row_order(shapes, slide_number)`  
**What it detects:** Groups of 3+ small image shapes at similar Y positions (within 5% of slide height)  
**What it transforms:** Reverses the X-order of logos within each row (first becomes last, last becomes first), preserving inter-logo gaps  
**Applied to:** All slides  
**Affected decks:** R6_18 slide 1 (client/partner logo row)

---

### Fix 18 — Circular Container Shape Text Centering

**Method:** `_center_text_in_container_shapes(shapes, slide_number)`  
**What it detects:** Shapes with circular/rounded preset geometry (`ellipse`, `roundRect`, `donut`, `hexagon`, etc.) OR nearly square aspect ratio (0.8–1.2) smaller than 35% slide width, containing Arabic text  
**What it transforms:** Sets `algn='ctr'` on all paragraphs within the shape  
**Applied to:** All slides  
**Affected decks:** R6_09 slide 1 (title in circle)

---

### Fix 19 — Bidi Base Direction for Mixed Text

**Method:** `_fix_bidi_base_direction(shapes, slide_number)`  
**What it detects:** Text frames with BOTH Arabic and Latin characters, with font size ≥ 16pt (title-sized text)  
**What it transforms:**
1. Sets `rtlCol='1'` on `<a:bodyPr>` if not already set
2. Sets `rtl='1'` on all paragraph `<a:pPr>` elements if not already set
3. Sets `algn='r'` (if current alignment is not `'r'` or `'ctr'`)  
**Applied to:** All slides  
**Affected decks:** R6_11 slide 1 (mixed Arabic/English title with embedded brand names)  
**Note:** The Unicode bidi algorithm requires an explicit RTL base direction when Arabic text contains embedded English brand names to avoid visual reordering issues.

---

## 4. Pipeline Walkthrough

The following describes exactly how a single deck goes from English PPTX input to Arabic RTL output, as implemented in `process_single_deck.py` (which bypasses `pipeline.py` and calls transformers directly).

### Step 0: Setup

```python
# ZIP_STORED monkey patch applied BEFORE importing python-pptx
_ZipStoredPatch.apply()
from pptx import Presentation
```

The monkey patch replaces `_ZipPkgWriter._zipf` with a version that uses `zipfile.ZIP_STORED` instead of `ZIP_DEFLATED`. This prevents large-file corruption when python-pptx writes ZIP64-format files.

### Step 1: Load

```python
shutil.copy2(input_path, output_path)  # Work on a copy
prs = Presentation(output_path)
slide_width = int(prs.slide_width)
slide_height = int(prs.slide_height)
```

The presentation is copied first. All mutations happen in-place on the copy.

### Step 2: Load Translations

```python
translations = json.load(open(translations_json))  # Dict[str, str]
```

The translations file is a JSON dictionary mapping English paragraph text to Arabic equivalents. Pre-computed by the calling agent via curl-based translation API calls.

### Phase 0: Layout Analysis

```python
analyzer = LayoutAnalyzer(prs)
layout_classifications = analyzer.analyze_all()
# Returns: Dict[int, LayoutClassification] — slide_number → classification
```

Also initializes the TemplateRegistry:
```python
registry = TemplateRegistry(slide_width, slide_height)
```

### Phase 2: Master & Layout Transform

```python
ml_transformer = MasterLayoutTransformer(prs, registry)
master_report = ml_transformer.transform_all_masters()
layout_report = ml_transformer.transform_all_layouts()
```

After this phase: all slide masters have `rtlCol='1'` on bodyPr and `lang='ar-SA'` on defRPr. All slide layout placeholders have mirrored X positions. Two-column layouts have swapped column positions.

### Phase 3: Slide Content Transform

```python
content_transformer = SlideContentTransformer(
    prs,
    template_registry=registry,
    layout_classifications=layout_classifications,
    translations=translations,
)
slide_report = content_transformer.transform_all_slides()
```

After this phase: all slide placeholders inherit from the transformed layout (xfrm removed). All freeform shapes are mirrored. All text has Arabic translations inserted. Tables are reversed. Charts are RTL-configured. All 19 fixes have been applied.

### Phase 4: Typography Normalization

```python
normalizer = TypographyNormalizer(prs)
typo_report = normalizer.normalize_all()
```

After this phase: fonts are mapped to Arabic-capable equivalents. Text frame margins are set. Bidi formatting is applied. Overflowing text has been reduced (max -20%).

### Step 3: Save

```python
prs.save(output_path)
```

Saved with `ZIP_STORED` (from the monkey patch).

### Step 4: Recompress

```python
recompress_pptx(output_path)
```

Re-compresses the ZIP_STORED file to ZIP_DEFLATED to reduce file size. LibreOffice handles compressed PPTX files better.

### Post-Pipeline: SmartArt (optional)

```python
translate_smartart_in_pptx(output_path, translations)
```

If called, this re-opens the saved file as a ZIP and translates SmartArt diagram XML parts.

---

## 5. Template Registry

### Overview

The `TemplateRegistry` is initialized with slide dimensions:
```python
registry = TemplateRegistry(slide_width_emu, slide_height_emu)
```

It builds rules for all 36 standard `ST_SlideLayoutType` values in `_build_default_rules()`.

### Layout Type Categories

| Category | Layout Types |
|----------|-------------|
| **Centered** | `title`, `secHead` |
| **Single content** | `tx`, `obj`, `titleOnly`, `objOnly`, `blank` |
| **Two-column / side-by-side** | `twoColTx`, `txAndChart`, `chartAndTx`, `txAndObj`, `objAndTx`, `txAndClipArt`, `clipArtAndTx`, `txAndMedia`, `mediaAndTx`, `picTx` |
| **Content-stacked** | `objTx`, `txObj`, `objOverTx`, `txOverObj` |
| **Multi-object** | `twoObj`, `fourObj`, `txAndTwoObj`, `twoObjAndTx`, `twoObjOverTx`, `twoTxTwoObj`, `twoObjAndObj`, `objAndTwoObj` |
| **Specialised** | `tbl`, `chart`, `dgm` |
| **Vertical text** | `vertTx`, `vertTitleAndTx`, `vertTitleAndTxOverChart` |
| **Fallback** | `cust` |

### Key Rules Summary

**`title` layout:**
- `ctrTitle`: `keep_centered`, `rtl=True`, `algn='ctr'`
- `subTitle`: `right_align`, `rtl=True`, `algn='r'`
- Freeforms: mirror
- Master elements: do NOT mirror (brand elements stay put)

**`tx` layout (most common content slide):**
- `title`: `keep_position`, `rtl=True`, `algn='r'`
- `body`: `keep_position`, `rtl=True`, `algn='r'`
- Freeforms: mirror
- `swap_columns=False`

**`twoColTx` layout (and all two-column variants):**
- `title`: `keep_position`, `rtl=True`, `algn='r'`
- `body`: `swap_with_partner`, `rtl=True`, `algn='r'`
- `swap_columns=True`

**`tbl` layout:**
- `title`: `keep_position`, `rtl=True`, `algn='r'`
- `tbl`: `reverse_columns`, `rtl=True`
- `table_action='reverse_columns'`

**`chart` layout:**
- `title`: `keep_position`, `rtl=True`, `algn='r'`
- `chart`: `reverse_axes`, `rtl=False`
- `chart_action='reverse_axes'`

**`fourObj` layout:**
- Each `obj`: `mirror`, `rtl=False`, `mirror_x=True` (each object mirrors its own X position)

**`cust` (fallback) layout:**
- Conservative defaults: keep positions, set RTL, right-align text
- `freeform_action='mirror'` (mirror freeform shapes as usual)

### Custom Rule Registration

Enterprise layouts not matching standard types can be registered:
```python
custom = LayoutTransformRules(
    layout_type='my_special_layout',
    description='Custom enterprise two-column with header',
    placeholder_rules={
        'title': PlaceholderAction(action='keep_position', set_rtl=True, set_alignment='r'),
        'body': PlaceholderAction(action='swap_with_partner', set_rtl=True, set_alignment='r'),
    },
    freeform_action='mirror',
    swap_columns=True,
    table_action='reverse_columns',
    chart_action='reverse_axes',
)
registry.register_custom_rule('my_special_layout', custom)
```

---

## 6. Property Resolution

### The 7-Level Chain

The `PropertyResolver` guarantees that every property has a non-None value for every run in the presentation. No property resolution is left to python-pptx defaults or guesswork.

```
For each run in each paragraph in each shape:

  1. Run-level (a:rPr):
     - sz → font size in hundredths of pt
     - b → bold
     - i → italic
     - a:latin/@typeface → font name
     - a:cs/@typeface → complex script font name

  2. Paragraph default (a:pPr/a:defRPr):
     Same attributes, checked if level 1 returned None

  3. Text frame list style (a:txBody/a:lstStyle/a:lvlNpPr/a:defRPr):
     Where N = paragraph.level + 1 (0-indexed level → 1-indexed element)

  4. Shape-level list style (shape element's own txBody lstStyle):
     Used when shape has a distinct lstStyle from its text frame

  5. Layout placeholder (matched by idx then type):
     a:txBody/a:lstStyle/a:lvlNpPr/a:defRPr

  6. Master placeholder (matched by idx then type):
     a:txBody/a:lstStyle/a:lvlNpPr/a:defRPr

  7. Master txStyles (p:txStyles → titleStyle/bodyStyle/otherStyle):
     Which style is used depends on placeholder type:
     - title, ctrTitle → titleStyle
     - body, subTitle, obj, tbl, chart, dgm, media, clipArt → bodyStyle
     - dt, ftr, sldNum, non-placeholder shapes → otherStyle

  Fallback: DEFAULT_FONT_SIZE_PT = 18.0, DEFAULT_FONT_NAME = 'Calibri', etc.
```

### Alignment Resolution

Alignment follows the same chain but the path is slightly different:

1. Explicit `algn` attribute on `<a:pPr>` (paragraph's own pPr, not on a run)
2. Skip level 2 and 4 (no alignment on defRPr)
3. Text frame lstStyle `algn` on `lvlNpPr`
4. Layout placeholder lstStyle `algn` on `lvlNpPr`
5. Master placeholder lstStyle `algn` on `lvlNpPr`
6. Master txStyles `lvlNpPr/@algn`
7. Default: `'l'` (left)

### RTL Resolution

RTL follows a shorter chain:

1. Explicit `rtl` on `<a:pPr>`
2. Layout placeholder lstStyle `rtl` on `lvlNpPr` (note: RTL is on `lvlNpPr`, not `defRPr`)
3. Master placeholder lstStyle `rtl` on `lvlNpPr`
4. Master txStyles `lvlNpPr/@rtl`
5. Default: `False`

### Theme Font References

When a font name is `+mj-lt`, `+mn-lt`, etc., it is a theme reference. The `_resolve_theme_font()` method resolves it by walking:
```
master.theme._element
  → a:themeElements
    → a:fontScheme
      → a:majorFont / a:minorFont
        → a:latin/@typeface
        → a:cs/@typeface
        → a:ea/@typeface
```

Results are cached per master (keyed by `id(master_obj)`) to avoid repeated XML traversal.

### Placeholder Matching

When resolving properties, the resolver must find the matching placeholder on the layout and master. The matching strategy:

1. Try to find a placeholder with the same `idx` attribute (index number)
2. If no idx match, fall back to finding a placeholder with the same `type` attribute

This is critical: placeholders with the same type but different indices (e.g., two body placeholders in a two-column layout) must be matched by index, not type.

### Synthetic Runs

Some paragraphs contain text in field elements (`<a:fld>`) or have only `<a:endParaRPr>` but no explicit `<a:r>` runs. The resolver creates a `ResolvedRun` for these using `_create_synthetic_run()`, which calls the level-2-through-7 chain directly.

---

## 7. Known Limitations

### 7.1 OLE-Embedded Objects

OLE objects (`oleObject` in `p:graphicFrame`) are detected and classified as `'ole'` in `_classify_shape_type()` but receive no transformation. Their content is opaque binary data (Excel spreadsheets, Word documents embedded as OLE). Text within OLE objects is not translated and RTL is not applied. The shape position is mirrored.

**Impact:** Slides with embedded Excel charts or Word tables will have their container shape repositioned but internal content remains LTR English.

### 7.2 SmartArt Requires Post-Save Processing

SmartArt diagram text lives in `ppt/diagrams/dataX.xml` and `ppt/diagrams/drawingX.xml` files inside the PPTX ZIP archive. The python-pptx shape API cannot access or modify these files. The `smartart_translator.py` module handles SmartArt text, but it must be called AFTER `prs.save()` because it re-opens the saved file as a ZIP.

**Consequence:** SmartArt translation is a separate step, not integrated into the main pipeline's in-memory mutation phase.

### 7.3 Nested Group Shapes

The `SlideContentTransformer._collect_all_shapes()` collects top-level shapes only. Group shapes are mirrored as single units. Text within group children is translated and RTL-formatted, but their positions are NOT mirrored individually (child coordinates are relative to the group's local coordinate space).

**Impact:** Nested groups within groups (rare but possible) may have text translated but layout not fully corrected.

### 7.4 Table Merged Cells

The table RTL transformation reverses `<a:tc>` elements within `<a:tr>` rows. OOXML tables support merged cells via `gridSpan` and `rowSpan` attributes. Reversing cells with complex merge patterns may produce invalid table structures. The code does not account for `gridSpan > 1`.

### 7.5 Font Size Reduction Bounds

The typography normalizer applies at most 20% font reduction. If Arabic text still overflows after the maximum reduction, it logs a warning and leaves the text as-is. No infinite shrink loop is attempted. The affected shapes are identified in the `TransformReport.warnings` list for manual review.

### 7.6 Color Scheme References

The `_resolve_color()` method in `PropertyResolver` extracts `<a:srgbClr val="FF0000">` (direct RGB) colors. For `<a:schemeClr>` (theme color references like `tx1`, `accent1`), it returns the scheme name string rather than the resolved RGB value. Full theme color resolution would require walking the theme XML's `<a:fmtScheme>` → `<a:fillStyleLst>` chain, which is not implemented.

### 7.7 Audio/Video Media Shapes

Shapes with embedded audio or video (`media` shape type) are detected but not transformed. Their container position is mirrored, but media playback coordinates and any overlaid text are not adjusted.

### 7.8 Hyperlinks

Run-level hyperlinks (`<a:hlinkClick>`) in `<a:rPr>` elements are preserved during text translation. However, after replacing run text with Arabic, the hyperlink rId still points to the original URL. This is correct behavior — hyperlinks should not change — but developers should be aware that hyperlinks are preserved as-is.

---

## 8. Key Implementation Notes

### 8.1 ZIP_STORED Monkey Patch

**Problem:** When python-pptx saves large PPTX files (>50 MB), it uses `zipfile.ZIP_DEFLATED` compression. With files that contain many embedded images or complex slides, this can produce corrupted ZIP64 archives. LibreOffice sometimes refuses to open these files.

**Fix:** Monkey-patch `pptx.opc.serialized._ZipPkgWriter._zipf` before importing any pptx modules:

```python
import pptx.opc.serialized as _ser
import zipfile

def _zipf_stored(self):
    return zipfile.ZipFile(self._pkg_file, 'w', zipfile.ZIP_STORED)

_ser._ZipPkgWriter._zipf = lazyproperty(_zipf_stored)
```

This must be applied BEFORE `from pptx import Presentation`.

**Trade-off:** ZIP_STORED produces larger files (no compression). The `recompress_pptx()` function re-compresses with `ZIP_DEFLATED` after saving to recover file size, while avoiding the large-file corruption issue.

**Recompression:**
```python
def recompress_pptx(pptx_path):
    with zipfile.ZipFile(pptx_path, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                zout.writestr(item, zin.read(item.filename))
    shutil.move(tmp, pptx_path)
```

---

### 8.2 curl-Based Translation (Python requests Hangs)

**Problem:** The translation API (Google Translate or similar) sometimes causes `requests.post()` to hang indefinitely when called from within the sandboxed Python environment. The hang does not resolve even with `timeout` parameters.

**Workaround:** Call the translation API using `subprocess.run(['curl', ...])` instead of `requests`. `curl` is a system subprocess that does not share Python's network stack and is not affected by the sandbox's socket restrictions.

```python
result = subprocess.run(
    ['curl', '-s', '-X', 'POST', url,
     '-H', 'Content-Type: application/json',
     '-d', json.dumps(payload)],
    capture_output=True, text=True, timeout=30
)
response = json.loads(result.stdout)
```

This approach bypasses the Python `requests` library entirely and has proven reliable across all 16 test decks.

---

### 8.3 LibreOffice Memory Issues with Large Files

**Problem:** LibreOffice (`soffice --headless --convert-to pdf`) exhausts memory and crashes when converting PPTX files with many high-resolution images or complex SmartArt (> ~100 MB uncompressed).

**Workarounds:**
1. **Per-deck subprocess isolation:** `batch_process.py` processes one deck at a time, calling `process_single_deck.py` as a new Python process. This ensures LibreOffice runs in a fresh process with no memory carryover between decks.
2. **ZIP_STORED + recompress:** Reduces peak memory during PPTX write by avoiding real-time compression.
3. **LibreOffice environment variables:** The `soffice` helper sets `SAL_USE_VCLPLUGIN=svp` (headless rendering) and `HOME=/tmp/lo_home` (isolated user profile) to prevent LibreOffice from reading corrupted cached profiles.
4. **Fallback to pdftoppm:** If LibreOffice produces a blank or corrupted PDF, the test harness falls back to `pdftoppm` for slide image extraction.

---

### 8.4 Sandbox Timeout Constraints

**Problem:** The execution sandbox has a per-operation timeout (typically 60 seconds for most operations, 600 seconds for long-running processes). Processing a 50-slide deck with complex tables and charts can exceed 60 seconds.

**Mitigations:**
1. **Batch by deck:** `batch_process.py` processes decks sequentially, not concurrently, so each deck gets its own time budget.
2. **No loops in transformers:** The design principle "phases execute exactly once, no fix loops" prevents time-unbounded processing. v1 had VQA fix loops that could run indefinitely.
3. **Performance guards:** Collision detection (`Fix 9`) is skipped on slides with >50 shapes. Chart label translation only iterates `<c:v>` elements (not every XML node).
4. **Lazy property resolution:** The `PropertyResolver` caches theme font maps per master and layout type inferences per layout object. These avoid re-traversing the same XML subtrees.

---

## 9. Configuration & Dependencies

### Python Packages

| Package | Usage |
|---------|-------|
| `python-pptx` | PPTX parsing and in-memory mutation. All shape/paragraph/run access. |
| `lxml` | Direct XML manipulation for operations beyond python-pptx's API. Required for setting `rtl`, `algn`, chart axis reversal, table column reversal, etc. |
| `Pillow` (`PIL`) | Slide image generation in the test harness (optional; comparison images skipped if unavailable). |

### System Tools

| Tool | Usage |
|------|-------|
| LibreOffice (`soffice`) | Headless PDF/PNG conversion of PPTX slides for visual QA. Required for `test_harness.py`. |
| `pdftoppm` | Fallback PDF-to-image conversion when LibreOffice renders blank output. Part of the `poppler-utils` package. |
| `curl` | HTTP client for translation API calls (replaces `requests` to avoid sandbox hang). |

### Python Version

Requires Python 3.8+ for:
- `from __future__ import annotations` (PEP 563 deferred evaluation)
- `dataclasses.field()` with `compare=False, hash=False`
- `frozenset` literals and `f-strings`

### File/Directory Structure

```
/home/user/workspace/
├── slideshift_v2/
│   ├── __init__.py (implied)
│   ├── models.py
│   ├── utils.py
│   ├── property_resolver.py
│   ├── layout_analyzer.py
│   ├── template_registry.py
│   ├── rtl_transforms.py
│   ├── typography.py
│   ├── structural_validator.py
│   ├── smartart_translator.py
│   ├── llm_translator.py          # Phase 1b — dual-LLM translation backend
│   ├── pipeline.py
│   └── test_harness.py
├── process_single_deck.py
├── batch_process.py
├── translations_cache/
│   └── R6_03_msnai_small.json  # Per-deck translation dictionaries
├── v2_fixed_output/
│   └── R6_03_msnai_small_v2.pptx  # Transformed outputs
└── test_results/
    └── R6_03_msnai_small/
        └── report.html
```

### Environment Variables

**LibreOffice (existing):**
- `SAL_USE_VCLPLUGIN=svp` — headless rendering plugin
- `HOME=/tmp/lo_home` — isolated user profile directory
- `TMPDIR=/tmp` — explicit temp directory

**Dual-LLM translation (Phase 1b):**
- `OPENAI_API_KEY` — required when `--llm-translate` is active; passed to GPT-4o calls
- `ANTHROPIC_API_KEY` — optional; if set, enables the Claude Sonnet 4.5 QA pass. If absent, `LLMTranslator` skips Layer 2 and runs GPT-only mode.

---

## 10. Running the Pipeline

### Single Deck Processing

```bash
# Standard mode (Google Translate / pre-computed JSON cache)
python process_single_deck.py <input.pptx> <output.pptx> <translations.json>

# Dual-LLM mode (Phase 1b — GPT-4o + Claude QA)
python process_single_deck.py <input.pptx> <output.pptx> --llm-translate

# Dual-LLM mode, GPT-4o only (skip Claude QA pass)
python process_single_deck.py <input.pptx> <output.pptx> --llm-translate --no-qa
```

Example (standard mode):
```bash
python process_single_deck.py \
    R6_03_msnai_small.pptx \
    v2_fixed_output/R6_03_msnai_small_v2.pptx \
    translations_cache/R6_03_msnai_small.json
```

The translations JSON must be a flat dictionary `{english_text: arabic_text}`:
```json
{
  "Microsoft AI": "ذكاء اصطناعي مايكروسوفت",
  "Revenue Growth": "نمو الإيرادات",
  "Q1 2024": "الربع الأول 2024"
}
```

**Output to stdout:**
```
============================================================
  Processing: R6_03_msnai_small.pptx
============================================================
  [Patch] ZIP_STORED monkey-patch applied
  Loaded 47 translations from R6_03_msnai_small.json
  Slide dimensions: 9144000x5143500 EMU, 18 slides
  Phase 0: Analyzed 18 slide layouts
  Phase 2: 42 master + 89 layout changes
  Phase 3: 318 slide changes
  Phase 4: 124 typography changes
  Saved: R6_03_msnai_small_v2.pptx (4823 KB)
  [Recompress] Saved compressed: 2847 KB
  DONE: R6_03_msnai_small.pptx
```

### Batch Processing

```bash
# Process all 16 decks (standard mode)
python batch_process.py

# Process all 16 decks with dual-LLM translation
python batch_process.py --llm-translate

# Process all 16 decks with GPT-4o only (no Claude QA pass)
python batch_process.py --llm-translate --no-qa

# Process only decks matching a pattern
python batch_process.py R6_03
python batch_process.py msnai
```

The batch runner processes decks sequentially (one at a time) to avoid LibreOffice memory exhaustion. Results are saved to `v2_fixed_output/batch_results.json`.

**The 16 test decks:**
- R6_03_msnai_small
- R6_04_sales_bd_marketing
- R6_05_iv_reporting_tables
- R6_06_humanoid_investments
- R6_07_mr_business_vision
- R6_08_slideforest_sales_template
- R6_09_slideforest_minimal_template
- R6_10_slideforest_ultimate_template
- R6_11_iv_quarterly_update
- R6_12_msnai_full_deck
- R6_13_mas_proposal
- R6_14_april_iv
- R6_15_ytelecom_draft
- R6_16_ria_tech_ceo
- R6_17_business_review
- R6_18_agtech_final

### Test Harness (Visual QA)

```bash
python slideshift_v2/test_harness.py R6_03_msnai_small.pptx
python slideshift_v2/test_harness.py --dir /home/user/workspace/
```

Outputs HTML comparison reports with side-by-side slide images to `test_results/<deck_name>/report.html`.

### Using the Pipeline Programmatically

If using `pipeline.py` directly (rather than `process_single_deck.py`):

```python
from slideshift_v2.pipeline import SlideShiftV2Pipeline

config = PipelineConfig(
    input_path='input.pptx',
    output_path='output.pptx',
    translate_fn=my_translation_function,  # Optional
    skip_translation=False,
    max_font_reduction_pct=20.0,
    log_level='INFO',
)

pipeline = SlideShiftV2Pipeline(config)
result = pipeline.run()

if result.success:
    print(f"Success: {result.output_path}")
    print(f"Duration: {result.total_duration_ms:.0f}ms")
else:
    print(f"Failed: {result.error}")
```

The `translate_fn` callable signature: `(List[str]) → Dict[str, str]` — takes a list of English text strings, returns a mapping from each English string to its Arabic translation.

### Debugging Individual Phases

To debug a specific phase, import the classes directly:

```python
from pptx import Presentation
from slideshift_v2.property_resolver import PropertyResolver
from slideshift_v2.layout_analyzer import LayoutAnalyzer
from slideshift_v2.template_registry import TemplateRegistry
from slideshift_v2.rtl_transforms import MasterLayoutTransformer, SlideContentTransformer
from slideshift_v2.typography import TypographyNormalizer
from slideshift_v2.structural_validator import StructuralValidator

prs = Presentation('input.pptx')

# Phase 0: Check what properties were resolved
resolver = PropertyResolver(prs)
resolved = resolver.resolve_presentation()
for slide in resolved.slides:
    for shape in slide.shapes:
        print(f"  Shape '{shape.shape_name}':")
        for para in shape.paragraphs:
            for run in para.runs:
                print(f"    Run: {run.text!r} font={run.effective_font_name} "
                      f"size={run.effective_font_size_pt}pt "
                      f"from={run.source_font_size_level}")

# Phase 0: Check layout classifications
analyzer = LayoutAnalyzer(prs)
for slide_num, cls in analyzer.analyze_all().items():
    print(f"Slide {slide_num}: {cls.resolved_type} ({cls.confidence:.0%}) "
          f"[{cls.layout_name}]")

# Phase 5: Run validation only
validator = StructuralValidator(prs)
report = validator.validate()
for issue in report.issues:
    print(f"  [{issue.severity.upper()}] Slide {issue.slide_number} "
          f"'{issue.shape_name}': {issue.message}")
print(f"Pass: {report.passed} | Errors: {report.errors} | Warnings: {report.warnings}")
```

### Translation File Preparation

Translations are not produced by the pipeline itself — they must be pre-computed. The translation workflow:

1. Extract all paragraph texts from the PPTX (can use `PropertyResolver` or a simple paragraph walk)
2. Call the translation API (use `curl` subprocess to avoid `requests` hang):
   ```bash
   curl -s -X POST "https://translation.googleapis.com/language/translate/v2" \
     -H "Content-Type: application/json" \
     -d '{"q": ["Revenue Growth", "Q1 2024"], "target": "ar", "key": "API_KEY"}'
   ```
3. Save the mapping as JSON to `translations_cache/<deck_name>.json`
4. Run `process_single_deck.py` with the translations file

The translation map key is the exact paragraph text as extracted from the PPTX (including leading/trailing whitespace). The fuzzy matcher in `SlideContentTransformer` handles minor whitespace differences.

---

*Documentation generated from complete source code review of SlideShift v2 (12,176 lines). All module descriptions, function signatures, and behavioral specifications are derived directly from the source files.*

---

## 11. Phase 1b — Dual-LLM Translation Backend

**Module:** `slideshift_v2/llm_translator.py` (1,061 lines)

Phase 1b is an optional, drop-in replacement for the Google Translate (GTX) translation step. When activated via `--llm-translate`, it replaces the curl-based GTX calls with a three-layer pipeline using GPT-4o as the primary translator and Claude Sonnet 4.5 as an independent QA reviewer. The output is a translation map in exactly the same JSON format as existing GTX caches, so no other pipeline stage requires modification.

### 11.1 Architecture

```
List[str] (English paragraph texts)
       │
       ▼
┌──────────────────────────────────────────┐
│  Layer 1: TokenProtector (pre-processing)│
│  Replaces 176 abbreviations + regex      │
│  patterns with ⟦PROTXXXX⟧ placeholders   │
└─────────────────┬────────────────────────┘
                  │
       ┌──────────▼───────────────────────────┐
       │  GPT-4o EN→AR Translation           │
       │  Batches of 40 strings              │
       │  Temperature 0.1, JSON mode         │
       │  System prompt: 2,013 chars         │
       └──────────┬───────────────────────────┘
                  │
       ┌──────────▼───────────────────────────┐
       │  TokenProtector: restore placeholders│
       └──────────┬───────────────────────────┘
                  │
       ┌──────────▼───────────────────────────┐
       │  Layer 2: Claude Sonnet 4.5 QA      │
       │  Reviews EN→AR pairs in batches     │
       │  of 60; returns corrections only    │
       │  System prompt: 1,536 chars         │
       └──────────┬───────────────────────────┘
                  │
       ┌──────────▼───────────────────────────┐
       │  Layer 3: DomainGlossary overrides  │
       │  Hardcoded BAD_TRANSLATIONS dict    │
       │  Catches known catastrophic failures│
       └──────────┬───────────────────────────┘
                  │
                  ▼
       Dict[str, str] translation_map
       (same format as GTX cache JSON)
```

### 11.2 Layer 1 — TokenProtector

**Problem it solves:** GPT-4o — despite explicit instructions — will sometimes transliterate abbreviations (`HW` → `هو`), translate financial codes (`EBITDA` → Arabic expansion), or mangle numeric identifiers (`Q1-24` → `الربع الأول 24`). These are Category A failures that produce unintelligible Arabic in executive presentations.

**Solution:** Before GPT sees any text, every token that must pass through untranslated is replaced with an opaque placeholder string `⟦PROT0001⟧`. GPT translates the surrounding Arabic context while placeholders remain byte-for-byte identical. After translation, placeholders are substituted back.

**Protected token inventory:**

| Type | Count | Examples |
|------|-------|---------|
| Business abbreviations | 176 entries | `HW`, `SW`, `OS`, `GDPR`, `HIPAA`, `EBITDA`, `CAPEX`, `OPEX`, `KPI`, `SLA`, `ROI` |
| Fiscal identifiers | regex | `Q1-24`, `Q3 2025`, `FY2025`, `FY26`, `H1 2024` |
| Numeric + unit | regex | `17.4M`, `$500K`, `£2.1B`, `€300M`, `500ms`, `16GB`, `99.9%` |
| URLs | regex | Any token matching `https?://` or `www.` |
| Email addresses | regex | Any token matching `\S+@\S+\.\S+` |

**Matching order:** Longest match first, to prevent `CAPEX` from being partially matched as `CAP`. Exact-match set is checked before regex patterns.

**Restoration guarantee:** `TokenProtector` maintains a per-string restoration map (`{placeholder: original_token}`). If GPT corrupts or drops a placeholder (rare), the restoration step logs a warning and inserts the original token at the end of the string rather than silently losing it.

### 11.3 Layer 2 — Claude QA Pass

**Problem it solves:** Even with TokenProtector active, GPT-4o produces errors in approximately 25% of strings in complex business decks — wrong terminology, gender agreement failures, inconsistent translation of the same term, overly literal phrasing that changes business meaning.

**Solution:** After GPT translation, all EN→AR pairs are submitted to Claude Sonnet 4.5 for independent review. Claude is a different model with different training data and different translation tendencies. Errors that GPT won't self-correct are reliably caught by Claude's independent review pass.

**Batching:** Pairs are submitted 60 at a time. Claude's response is a JSON array containing only the strings where corrections were made. This keeps the response compact — the majority of strings need no correction and are implicitly approved.

**Six issue categories:**

| Category | What Claude checks | Example failure |
|----------|--------------------|-----------------|
| Abbreviation mangling | Abbreviated terms expanded or transliterated | `OS` → "platforms" instead of staying as `OS` |
| Brand name corruption | Product/company names rendered in Arabic script | `SlideShift` → `سلايدشيفت` |
| Number/unit errors | Numeric values or units altered | `$500K` → `500 ألف دولار` (acceptable) vs. a different number |
| Semantic inversions | Negations dropped; meaning reversed | "not available" → "available" |
| Terminology inconsistency | Same source term gets different Arabic translations across slides | "revenue" → `إيرادات` in one place, `دخل` in another |
| Register issues | Informal or colloquial Arabic where formal executive register is required | Colloquial phrases in a board-level financial slide |

**Live test data (R6_07 MR Business Vision, 61 strings):**

| Issue category | Count |
|----------------|-------|
| Semantic errors | 5 |
| Terminology inconsistencies | 5 |
| Register/grammar issues | 4 |
| Abbreviation issues | 1 |
| **Total issues found** | **15** |

Specific corrections made: `OS` → "platforms" reversed; omitted word "Modular" restored in a product name; gender agreement errors on three adjective-noun pairs corrected; "at cost" → "with cost" meaning change reversed to preserve the business idiom.

### 11.4 Layer 3 — Domain Glossary

**Problem it solves:** Both LLMs share certain systematic biases on domain-specific terms. Known failure patterns that have been observed in production are hardcoded and applied deterministically after the Claude pass, providing a hard guarantee that no combination of model errors can produce these specific bad outputs.

**`BAD_TRANSLATIONS` dict entries (sample):**

```python
BAD_TRANSLATIONS = {
    "المخلفات الخطرة":     "الأجهزة",                        # HW: hazardous waste → hardware
    "مؤتمن":              "سري",                             # confidential: trustworthy → secret/confidential
    "الناتج المحلي الإجمالي": "اللائحة العامة لحماية البيانات", # GDPR: GDP → GDPR
    "خط أنابيب الإيرادات":  "مسار الإيرادات",                 # revenue pipeline: literal pipe → pipeline idiom
}
```

Each entry captures a known catastrophic mistranslation (wrong Arabic → correct Arabic). The dict is keyed by the bad output so it can be applied as a final substitution pass over the completed translation map.

### 11.5 System Prompts

**GPT-4o system prompt (2,013 chars) — key rules:**
- Treat `⟦PROTXXXX⟧` tokens as opaque: copy them unchanged into output
- Preserve brand names, product names, and proper nouns in their original Latin script
- Use Modern Standard Arabic (MSA), formal register appropriate for executive presentations
- Financial and technical abbreviations not caught by TokenProtector should be left in English
- Respond with a JSON object: `{"translations": [{"en": "...", "ar": "..."}, ...]}`
- Do not add, remove, or reorder entries from the input list

**Claude QA prompt (1,536 chars) — key rules:**
- Review each EN→AR pair for the six issue categories defined above
- Return only pairs that require correction (omit approved pairs from output)
- Output format: `[{"index": N, "corrected_ar": "...", "issue": "brief reason"}]`
- Do not re-translate from scratch — minimally correct the existing Arabic
- When in doubt, prefer the more formal/conservative option

### 11.6 Performance & Cost

| Metric | Value |
|--------|-------|
| Average deck size (test corpus) | ~61 unique strings |
| GPT-4o batches per deck | ~2 (40 strings/batch) |
| Claude batches per deck | ~1 (60 pairs/batch) |
| GPT-4o cost per deck | ~$0.02 |
| Claude Sonnet 4.5 cost per deck | ~$0.07 |
| **Total cost per deck** | **~$0.09** |
| End-to-end latency (both models) | ~73 seconds |
| Retry policy | 2 retries, exponential backoff (1s, 2s) |

Cost is negligible relative to the reputational risk of bad translations in Saudi/GCC executive presentations. The full GTX+dual-LLM pipeline costs approximately $0.09/deck compared to $0.00 for GTX alone.

### 11.7 Translation Quality Results

Tested on R6_07 (MR Business Vision — a 28-slide business strategy deck):

| Approach | Quality Score (1–10) | Category A failures |
|----------|---------------------|---------------------|
| Google Translate baseline | 5.5 | HW → hazardous waste, brand corruption, semantic inversions |
| GPT-4o only | 8.8 | 0 (TokenProtector eliminates Category A) |
| GPT-4o + Claude QA | **9.5** | 0 |

Category A failures (HW → hazardous waste, brand name corruption, semantic inversions) are eliminated at the TokenProtector and GPT-4o layer. The Claude QA pass drives the score from 8.8 to 9.5 by catching the subtler Category B failures (register, consistency, partial semantic errors).

### 11.8 Integration with Existing Pipeline

**`process_single_deck.py` integration:**

```bash
# Standard (GTX cache)
python process_single_deck.py input.pptx output.pptx translations.json

# Dual-LLM mode (translates from scratch, no cache needed)
python process_single_deck.py input.pptx output.pptx --llm-translate

# GPT-only (skip Claude QA pass)
python process_single_deck.py input.pptx output.pptx --llm-translate --no-qa
```

When `--llm-translate` is passed, `process_single_deck.py` instantiates `LLMTranslator` and calls `.translate()` to produce the translation map before passing it to `SlideContentTransformer`. The translations JSON is also written to `translations_cache/<deck_name>_llm.json` so it can be reused in subsequent runs without re-calling the APIs.

**`batch_process.py` integration:**

```bash
python batch_process.py --llm-translate         # All 16 decks, full dual-LLM
python batch_process.py --llm-translate --no-qa # All 16 decks, GPT-only
```

**Backward compatibility:** The `translations.json` positional argument is still accepted. Existing translation caches produced by GTX continue to work without modification. The `--llm-translate` flag is additive.

**Environment variables:**

| Variable | Required | Effect |
|----------|----------|--------|
| `OPENAI_API_KEY` | Yes (when `--llm-translate`) | Authenticates GPT-4o API calls |
| `ANTHROPIC_API_KEY` | No | If set, enables the Claude QA pass. If absent, `LLMTranslator` runs GPT-only mode regardless of `--no-qa` flag |

### 11.9 Rationale for Dual-Model Approach

Three independent layers provide defense-in-depth against translation failures:

1. **TokenProtector** eliminates all Category A failures deterministically. No matter how GPT behaves, protected tokens cannot be corrupted.
2. **Claude QA** provides independent verification from a different model. Claude and GPT have different training distributions and different systematic biases. An error GPT makes consistently, Claude is unlikely to make — and vice versa. Two independent models reviewing the same translation reduces the joint error rate far more than one model reviewing itself.
3. **Domain Glossary** provides a hard guarantee layer. Even if both LLMs independently make the same mistake on a known failure mode, the deterministic override catches it.

The total cost ($0.09/deck) is negligible. A single bad translation in a GCC executive presentation — "hazardous waste" where "hardware" was intended, or a brand name transliterated into Arabic script — carries reputational risk that dwarfs the translation cost by orders of magnitude.

---

*Documentation updated March 2026 to reflect Phase 1b dual-LLM translation backend (llm_translator.py, 1,061 lines). Total codebase: 12,176 lines.*
