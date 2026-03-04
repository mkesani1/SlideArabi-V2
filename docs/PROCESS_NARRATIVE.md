# SlideShift v2: Process Narrative

**Project:** SlideShift v2 — English-to-Arabic RTL PowerPoint Converter  
**Document type:** Development Process Narrative  
**Audience:** Technical product owner and engineers inheriting or extending this project  

---

## Overview

SlideShift is a tool that converts English PowerPoint presentations into Arabic RTL (right-to-left) format. Version 2 represents a complete architectural rethink of the original codebase — built from scratch, guided by frontier AI models, implemented entirely by Claude Sonnet, and validated through an iterative visual QA loop powered by Gemini vision.

This document tells the story of how v2 was built: the decisions made, the models involved, the challenges encountered, and how they were resolved.

---

## The Problem

Converting a PowerPoint from English to Arabic is not simply a matter of translating text. Arabic is read right-to-left, which means the entire visual logic of a slide must be inverted. A title that sits in the upper-left corner of an English slide belongs in the upper-right of its Arabic equivalent. Chevron arrows that flow left-to-right must reverse direction. Split panels swap sides. Timelines alternate in the opposite direction. Logos that sit on the right edge of a footer now belong on the left.

Layered on top of this is PowerPoint's own complexity: shapes inherit properties from slide layouts, which in turn inherit from slide masters. A shape's apparent font size may not be defined on the shape itself at all — it may live three levels up the inheritance chain. Any conversion tool that ignores this will produce subtly broken output that looks correct in Python but breaks when opened in PowerPoint.

The v1 codebase (available at `mkesani1/slideshift` on GitHub) had already made progress on these problems, but it had limitations in structure and coverage. Version 2 set out to fix those at the architectural level.

---

## Phase 1: Architecture Design

Development began not with writing code, but with understanding what the code needed to become.

Frontier AI models — Claude and GPT — were asked to analyze the existing v1 codebase and propose an architecture for v2. This was a deliberate choice: rather than designing the system in isolation, the goal was to surface patterns in the existing code, identify where the v1 approach strained under edge cases, and get a second opinion on the right structure before a single line of v2 was written.

The models converged on several key recommendations, shaped heavily by one overriding principle from the project owner: **deterministic rules over reasoning**. This meant that slide transformation logic should be expressed as explicit, auditable code — not delegated to a language model at runtime. If a shape needs to be mirrored, a function should mirror it. The system should be predictable, testable, and debuggable without requiring a model call to understand its behavior.

From this principle, a modular pipeline architecture emerged:

| Module | Responsibility |
|---|---|
| `pipeline.py` | Orchestrates the end-to-end conversion process |
| `layout_analysis.py` | Identifies structural patterns in each slide |
| `template_registry.py` | Classifies slides by template type using pattern matching |
| `property_resolver.py` | Resolves PowerPoint's 3-level inheritance hierarchy |
| `rtl_transforms.py` | Applies all RTL layout transformations |
| `validation.py` | Post-conversion checks for common failure modes |

The **template registry** pattern was a particularly important architectural decision. Rather than analyzing each slide independently at runtime, slides are classified up front against a library of known template patterns (cover slide, section divider, split panel, timeline, etc.). This classification then drives which transforms are applied, making the system's behavior explicit and its coverage measurable.

The **property resolver** addresses PowerPoint's inheritance model directly. When a shape's text has no explicit font size, the resolver walks up the hierarchy — checking the slide layout, then the slide master — to find the inherited value. This prevents the common failure mode of shapes that look correct in isolation but render wrong when PowerPoint applies its own inheritance rules.

---

## Phase 2: Implementation

With the architecture defined, implementation was handled entirely by **Claude Sonnet 4**, operating within the Perplexity Computer sandbox.

The full codebase spans **11,115 lines of Python** across the core modules. The three heaviest modules reflect where the complexity lives:

- **`rtl_transforms.py`** — 3,106 lines. This is the core engine. It contains the transformation logic for every recognized layout pattern: position mirroring, alignment flipping, chevron reversal, split-panel handling, timeline alternation, logo repositioning, and more.
- **`property_resolver.py`** — 1,758 lines. Handles the full 3-level inheritance walk (shape → layout → master) for fonts, colors, paragraph properties, and positioning.
- **`template_registry.py`** — 1,193 lines. Defines the pattern library used to classify slides, with matching logic for each known template type.

**Translation** is handled via the Google Translate API. One non-obvious implementation detail: translation calls are made using `curl` via a Python subprocess, rather than the standard `requests` library. The reason is purely environmental — the Python `requests` library hangs indefinitely in the sandbox, while `curl` subprocess calls complete reliably. Translation results are cached to avoid redundant API calls across repeated runs of the same deck.

---

## Phase 1b: Dual-LLM Translation Backend

After the RTL transformation engine reached production quality across all 16 test decks, a separate workstream addressed the translation quality problem.

Google Translate (GTX) — the original translation provider — produces output rated at approximately 5.5/10 for business Arabic. The failure modes in executive presentations are severe: `HW` (hardware) translates to “hazardous waste” in Arabic; `GDPR` becomes “GDP”; `confidential` becomes “trustworthy”; brand names get rendered in Arabic script. For Saudi and GCC executive presentations, these errors are not tolerable.

`llm_translator.py` (1,061 lines) implements a three-layer replacement for the GTX translation step:

**Layer 1 — TokenProtector (pre-processing):** Before GPT-4o sees any text, 176 protected abbreviations and regex-matched tokens (quarter identifiers, numeric units, URLs, emails) are replaced with opaque `⟦PROTXXXX⟧` placeholder strings. GPT translates the surrounding Arabic context while placeholders pass through unchanged. Originals are restored after translation. This eliminates all Category A failures deterministically.

**Layer 2 — GPT-4o translation:** The primary EN→AR translator. Processes 40 strings per batch at temperature 0.1 with a 2,013-character system prompt defining rules for placeholders, brand name preservation, formal Arabic register, and JSON response format.

**Layer 3 — Claude Sonnet 4.5 QA pass:** Reviews all EN→AR translation pairs in batches of 60, checking for abbreviation mangling, brand name corruption, number/unit errors, semantic inversions, terminology inconsistency, and register issues. Returns corrections only for strings where issues were found. Using a different model for QA than for translation is the key design choice: errors GPT makes consistently, Claude is unlikely to make, reducing the joint error rate significantly.

**Deterministic glossary overrides:** A hardcoded `BAD_TRANSLATIONS` dict applied after the Claude pass catches any remaining known catastrophic failures.

**Quality results tested on R6_07 MR Business Vision:**

| Approach | Quality Score |
|----------|---------------|
| Google Translate baseline | 5.5 / 10 |
| GPT-4o only | 8.8 / 10 |
| GPT-4o + Claude QA | **9.5 / 10** |

The Claude pass found 15 issues across 61 strings in live testing, including semantic errors, terminology inconsistencies, gender agreement failures, and a business idiom meaning change. All Category A failures were eliminated at the GPT-4o layer. Cost is approximately $0.09 per deck (GPT $0.02 + Claude $0.07), with an end-to-end latency of approximately 73 seconds.

The module is backward-compatible: existing GTX translation caches continue to work as a positional argument. The dual-LLM mode is activated via `--llm-translate` on both `process_single_deck.py` and `batch_process.py`. The `--no-qa` flag allows GPT-only mode when `ANTHROPIC_API_KEY` is not available.

---

## Phase 3: Test-Driven Iteration

Rather than writing unit tests in isolation, the QA methodology was visual and end-to-end: run the converter against real presentation files and compare the output against the original slide by slide.

A corpus of **16 test PPTX decks** was assembled, covering a representative range of layouts and use cases:

- Sales decks
- Investor reports
- Technical proposals
- Minimal/clean templates
- Complex “ultimate” templates with dense layout patterns

Each iteration of the development cycle followed the same loop:

1. Process all 16 decks through the current version of the converter
2. Render each slide (both original and converted) as images using LibreOffice
3. Present side-by-side comparisons for visual review

**Visual QA** was performed by **Gemini 3 Flash** (vision model), running within Perplexity Computer. Given pairs of original and converted slide images, Gemini was asked to identify layout discrepancies, text overflow, positioning errors, misaligned elements, and any other visual anomalies. This provided a fast, scalable way to scan all 16 decks after each fix cycle without manual review of every slide.

---

## Phase 4: The Fix Cycle

The iterative loop produced **19 targeted fixes**, each addressing a specific class of conversion failure identified through visual QA. The fixes are grouped here by theme.

### Fixes 1–10: Core RTL Foundations

The first ten fixes established the fundamental correctness of the RTL transformation logic:

- **Alignment flipping** — left-aligned text becomes right-aligned; right-aligned becomes left-aligned
- **flipH prevention on images** — horizontal flip transforms were incorrectly being applied to image shapes, distorting photos
- **Footer mirroring** — footer elements (page numbers, dates, company names) are repositioned to their RTL equivalents
- **Logo overlap prevention** — logos were colliding with other footer elements after repositioning; collision detection was introduced
- **Text box width** — text boxes were not being resized to account for Arabic text expansion (Arabic typically runs 20–40% longer than equivalent English)
- **Collision detection** — a general-purpose system for detecting and resolving shape overlaps after transformation

### Fix 11: Arabic Text Wrapping

Arabic text set to `wrap=none` was overflowing its container without wrapping. The fix changes the wrap mode to `wrap=square` for Arabic content, allowing text to reflow within its bounding box.

### Fix 12: Title-Body Overlap

After repositioning, title shapes and body text shapes were occasionally overlapping. A dedicated resolution pass was added to detect this specific collision pattern and adjust vertical positions.

### Fix 13: Slide-Level Logo Mirroring

Logos embedded at the slide level (rather than in the footer area) were not being caught by the footer mirroring logic. A separate pass was added to handle slide-level logo repositioning.

### Fixes 14–19: Layout-Specific Refinements

The final six fixes addressed specific template patterns that required dedicated handling:

- **Cover title positioning** — cover slides have distinct title placement rules that differ from content slides
- **Split-panel mirroring** — slides divided into two vertical panels (e.g., image left / text right) need both content and panel boundaries swapped
- **Timeline alternation reversal** — timelines with alternating above/below content nodes must have their alternation pattern reversed, not just their horizontal positions
- **Logo row ordering** — when multiple logos appear in a row, their left-to-right order must be reversed for RTL
- **Text centering in circular shapes** — centered text in circular or oval shapes was shifting off-center after transformation
- **Bidi base direction** — the Unicode bidirectional base direction attribute was not being set correctly on text runs, causing RTL/LTR mixing in some edge cases

---

## Models Used

| Model | Role |
|---|---|
| **Claude Sonnet 4** (via Perplexity Computer) | Wrote all 11,115 lines of Python code for the core RTL engine. Handled architecture, implementation, debugging, all 19 fix cycles, code review, and refactoring. |
| **Claude Sonnet 4.5** (via Anthropic API) | Phase 1b QA reviewer. Reviews all EN→AR translation pairs post-GPT, catching semantic errors, terminology inconsistencies, register issues, and abbreviation mangling. |
| **GPT-4o** (via OpenAI API) | Phase 1b primary translator. Produces EN→AR translation at BLEU 54.2 / COMET 0.847 / Human Accuracy 4.3/5 — highest quality among frontier models tested for business Arabic. |
| **Gemini 3 Flash** (vision) | Visual QA — compared original and converted slide images to identify layout discrepancies, positioning errors, and text overflow. |
| **Google Translate API** | Original translation provider (GTX baseline, 5.5/10 quality). Superseded by dual-LLM pipeline for production use; cache files remain compatible. |
| **Claude / GPT** (architectural consultation) | Reviewed the v1 codebase and provided architectural recommendations during the design phase. |

The division of labor reflects each model's strengths. Claude Sonnet is the workhorse for long-context code generation and multi-step debugging. Gemini Flash's vision capability made it well-suited for rapid visual comparison across many slide images. GPT-4o leads on EN→AR translation quality based on benchmark results. Claude Sonnet 4.5 provides independent QA verification — a different model checking GPT's output catches more errors than GPT reviewing its own translations.

---

## Key Technical Decisions

### 1. Deterministic Rules Over LLM Reasoning

Every transformation in SlideShift v2 is expressed as explicit Python code. There are no runtime calls to a language model to decide how to transform a shape. This was a deliberate architectural choice: deterministic transforms are testable, auditable, and produce consistent output. An engineer reading the codebase can trace exactly why any given shape ended up where it did.

### 2. Template Registry Pattern Matching

Slides are classified against a registry of known template patterns before any transformation is applied. This makes coverage explicit — the system knows which slide types it handles, and unknown types can be added to the registry as new patterns are encountered. It also avoids the overhead and unpredictability of per-slide LLM classification at runtime.

### 3. Property Resolver for PowerPoint Inheritance

PowerPoint's 3-level inheritance hierarchy (shape → slide layout → slide master) is fully traversed by the property resolver before transformations are applied. This ensures that inherited properties are correctly accounted for, preventing the common failure mode where a shape looks correct in isolation but breaks when PowerPoint applies its own cascade.

### 4. ZIP_STORED Monkey-Patch

Python-pptx's default behavior with large PPTX files can cause issues. A monkey-patch forces ZIP_STORED compression mode during file manipulation to handle files reliably. Note that files saved in ZIP_STORED format must be recompressed to ZIP_DEFLATED before being passed to LibreOffice for rendering — LibreOffice runs out of memory on very large ZIP_STORED files.

### 5. curl Subprocess for Translation

The Google Translate API is called via `curl` subprocess rather than the Python `requests` library. This is a sandbox-specific workaround: `requests` hangs indefinitely in the execution environment, while `curl` subprocess calls complete reliably. Engineers running this outside the sandbox may wish to revisit this choice.

### 6. Side-by-Side Visual Comparison as Primary QA

The primary quality signal throughout development was visual: did the converted slide look correct next to the original? This is the right metric for a layout conversion tool, where the failure modes are visual (misalignment, overflow, wrong positions) rather than logical. The Gemini vision QA loop made this methodology scalable across 16 decks and 19 fix cycles.

### 7. Dual-LLM Translation Architecture

Google Translate is insufficient for high-stakes business Arabic. The dual-LLM pipeline (GPT-4o + Claude Sonnet 4.5 QA) raises translation quality from 5.5/10 to 9.5/10. The key architectural insight is that independent verification from a different model catches more errors than self-review — the two models have different systematic biases, so their error sets are largely non-overlapping. The `TokenProtector` pre-processing layer provides a deterministic guarantee that protected tokens (abbreviations, financial identifiers, numeric units) cannot be corrupted regardless of model behavior.

---

## Challenges and Solutions

### Sandbox Command Timeouts

The Perplexity Computer sandbox kills commands that run longer than approximately 50 seconds. Processing 16 PPTX decks in a single command exceeds this limit. The solution was batch processing via subagents — breaking the work into smaller units that each complete within the time budget.

### Large PPTX Files and LibreOffice OOM

PPTX files saved with ZIP_STORED compression can exceed 25MB, causing LibreOffice to run out of memory during slide rendering. The fix is to recompress these files to ZIP_DEFLATED before passing them to LibreOffice. This step is now part of the rendering pipeline.

### OLE-Embedded Objects

Some presentations contain OLE-embedded objects (embedded Excel charts, Word documents, etc.). These cannot be accessed or translated through the python-pptx API. This is a known limitation, documented as out of scope for v2.

### Arabic Text Expansion

Arabic text is consistently 20–40% longer than its English equivalent. Text boxes sized for English content will overflow when populated with Arabic translations. The property resolver and transform layer account for this by adjusting text box dimensions during conversion, but edge cases remain in tightly constrained layouts.

---

## What v2 Delivers

SlideShift v2 correctly handles the following layout patterns:

- Standard content slides (title + body)
- Cover and section divider slides
- Split-panel slides (image/text side by side)
- Timeline slides (horizontal and alternating)
- Footer layouts (logo, date, page number)
- Slides with circular and decorative shapes
- Chevron/arrow flow diagrams
- Multi-logo rows
- Slides with slide-level (non-footer) logo placement

The 19-fix cycle brought the output quality from rough but functional to visually correct across all 16 test decks.

---

## Handoff Notes for Engineers

If you are picking up this project, a few things worth knowing:

1. **Start with `pipeline.py`** — it is the entry point and orchestrates the full conversion flow. Reading it top to bottom gives you a map of the system.
2. **`rtl_transforms.py` is the core** — most bugs in visual output will trace back here. It is long but organized by transform type.
3. **The template registry drives coverage** — if a new slide layout is not converting correctly, the first question to ask is whether it matches a registered pattern.
4. **The `curl`-based translation calls are intentional** — do not replace them with `requests` without testing in the target environment first.
5. **Visual QA is the right test methodology** — generating slide images and comparing them side by side catches issues that unit tests miss. Keep the 16-deck corpus and extend it when new layout types are encountered.
6. **The dual-LLM translation pipeline is the recommended path for production use** — set `OPENAI_API_KEY` and `ANTHROPIC_API_KEY` and use `--llm-translate`. The GTX cache approach still works but produces significantly lower translation quality. If only `OPENAI_API_KEY` is set, the system runs GPT-only mode (8.8/10 quality, $0.02/deck).

---

*Document updated March 2026 to include Phase 1b dual-LLM translation backend (llm_translator.py). SlideShift v2 core engine by Claude Sonnet 4 via Perplexity Computer. Dual-LLM translation pipeline by GPT-4o (primary) and Claude Sonnet 4.5 (QA).*
