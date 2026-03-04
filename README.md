# SlideArabi V2

Template-first deterministic RTL transformation engine that converts English PowerPoint presentations to Arabic RTL format.

## Architecture

- **Phase 0**: Parse & resolve layout classifications
- **Phase 1**: Dual-LLM translation (GPT-5.2 primary → Claude Sonnet 4.6 QA)
- **Phase 2**: Master & layout RTL transforms
- **Phase 3**: Slide content transforms (placeholder inheritance with size-divergence guard)
- **Phase 4**: Typography normalization
- **Phase 5**: Structural validation
- **Phase 6**: Visual QA (Gemini vision model)

## Translation Stack

| Role | Model | Purpose |
|------|-------|--------|
| Primary | GPT-5.2 | Flagship translation |
| QA | Claude Sonnet 4.6 | Terminology consistency review |

## Key Features

- **Size-divergence guard**: Prevents placeholder geometry collapse during RTL inheritance (symmetric 30% threshold)
- **11-rule Visual QA**: Catches collapsed text, empty slides, overflow, collisions, glyph rendering issues
- **Embedded Excel detection**: Handles OLE-embedded spreadsheet objects
- **SmartArt translation**: Translates SmartArt text within diagram XML

## Setup

```bash
pip install python-pptx lxml Pillow
export OPENAI_API_KEY=your_key
export ANTHROPIC_API_KEY=your_key
```

## Usage

```bash
python process_single_deck.py input.pptx output.pptx --llm-translate
```

## Documentation

- [Engineering Docs](docs/ENGINEERING_DOCS.md) — Full architecture, module reference, all 19 RTL fixes
- [Process Narrative](docs/PROCESS_NARRATIVE.md) — Development history and decisions
