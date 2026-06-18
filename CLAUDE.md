# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

An autonomous Python pipeline that converts markdown documents or existing PowerPoint files into polished, Merck Healthcare KGaA-branded `.pptx` decks. No LLM is needed when a pre-built plan JSON is provided; Claude (via Merck Foundry AIP) is invoked only when converting unstructured source files.

## Environment

- Python 3.11 via conda environment `ds_env`
- Activate with: `eval "$(conda shell.bash hook)" && conda activate ds_env`
- Install dependencies: `pip install -r requirements.txt`

## Running the pipeline

```bash
# Interactive: convert a markdown or existing .pptx to Merck branding (prompts 6 gate questions)
python -m merck_pptx generate source.md output/deck.pptx

# Scripted: provide meta as a JSON file (no prompts)
python -m merck_pptx generate source.pptx output/deck.pptx --meta meta.json

# Direct: build from a pre-built plan JSON (no LLM, fully deterministic)
python -m merck_pptx build plan.json output/deck.pptx
```

## Required environment variables (for LLM features)

| Variable | Purpose |
|---|---|
| `AIP_BASE_URL` | Merck Foundry AIP endpoint (e.g. `https://merck.palantirfoundry.com/api/v1`) |
| `AIP_TOKEN` | Foundry API token |
| `AIP_MODEL` | Optional model override (default: `claude-sonnet-4-6`) |

These are only required for `generate` (not `build`).

## Package structure

```
merck_pptx/
├── __init__.py          generate_deck, build_from_plan, validate_plan, ValidationError
├── __main__.py          CLI entry point
├── merck_layouts.py     6,801-line rendering engine (do not edit)
├── build_from_plan.py   validate + dispatch to build_* + save
├── validate_plan.py     hard errors + warnings
├── llm.py               Claude API client (Foundry) + plan-generation system prompt
├── generate.py          top-level router: md/pptx/dict/json → .pptx
└── templates/
    ├── EU_Merck_Themed.pptx          EU region theme carrier
    └── USA_Merck_Themed_Base_v1.pptx USA/Canada region theme carrier
```

## Architecture

The pipeline has three layers:

1. **Router** (`generate.py` → `generate_deck(source, output_path, meta)`)
   - `dict` or `.json` input → skips LLM, calls `build_from_plan` directly
   - `.md` or `.pptx` input → extracts text, calls `llm.generate_plan()`, then `build_from_plan`

2. **LLM layer** (`llm.py` → `generate_plan(raw_content, meta) -> dict`)
   - Sends source content + meta to Claude via Foundry AIP
   - System prompt contains: full plan schema, all 44 layout descriptions, quality rules
   - Returns a parsed plan dict; `[PLACEHOLDER]` for unmappable content

3. **Builder** (`build_from_plan.py` → `build_from_plan(plan, output_path) -> str`)
   - Validates plan (`validate_plan.py`)
   - Opens the correct `.pptx` theme template (EU or USA) via `open_deck()`
   - Resolves style per slide (inherit → deck_style → auto-promote for executive categories)
   - Dispatches each slide to the matching `build_*` function in `merck_layouts.py`
   - Post-build: speaker notes, agenda hyperlinks, image placement
   - Saves via `save_deck()`

## Key rules in `merck_layouts.py`

- Three visual styles: `merck_executive`, `merck_corporate`, `merck_storytelling`
- Auto-promote to `merck_executive`: any slide with category matching "Executive Summary", "Recommendation", "Decision Request", "Risk", or "Tradeoff"
- EU template must never be used for USA users (legal disclaimer difference)
- `open_deck()` strips all pre-existing slides; only the theme/master is preserved
- `section_number` must be unique sequential integers on all content slides; `null` for Cover, Agenda, Section Divider, Close

## 44 layout keys

```
cover, exec_summary, agenda, section_divider, close,
chart, two_column, three_column, four_column, matrix_2x2,
phase_process, vertical_numbered, waterfall, decision_rows, gantt,
hero_stat, stat_strip, before_after, milestone_timeline, status_table,
hub_spoke, pillar_detail, label_rows, circular_flow, org_chart,
topic_set, arrow_chain, pull_quote, donut_chart, kpi_dashboard,
icon_grid, journey_map, funnel, comparison_table, score_table,
influence_diagram, word_cloud, pyramid, venn, risk_heatmap,
radar_chart, pros_cons, layered_stack, photo_text, fishbone
```

Content schemas for each layout are defined by the corresponding `build_*` function signature in `merck_layouts.py`.

### Most common content schemas

| Layout | Key content fields |
|---|---|
| `cover` | `subtitle`, `authors: [str]`, `key_messages: [str]` |
| `exec_summary` | `key_messages: [{label, body}]` max 5 |
| `agenda` | `chapters: [{number, title}]` |
| `section_divider` | `number`, `title` |
| `two_column` | `left: {header, items:[str]}`, `right: {header, items:[str]}` |
| `chart` | `chart: {type, data: {categories, series:[{name,values}]}}`, `callouts` |
| `decision_rows` | `decisions: [{tone, number, owner, text}]` max 5 |
| `hero_stat` | `stat: {value, label}`, `context: str` |
| `phase_process` | `phases: [{label, title, body, highlighted}]` max 5, `show_arrows` |
| `gantt` | `rows: [{label}]`, `quarters: [str]` |

## Plan JSON schema

```json
{
  "meta": {
    "region":          "EU" | "USA",
    "deck_label":      "string",
    "classification":  "Public" | "Internal" | "Confidential",
    "month_year":      "June 2026",
    "audience":        "string",
    "deck_style":      "merck_executive" | "merck_corporate" | "merck_storytelling",
    "variety_mode":    "default" | "creative",
    "show_disclaimer": false
  },
  "storyline": ["Chapter 1 action title", "Chapter 2 action title"],
  "slides": [
    {
      "page": 1,
      "page_function": "Cover",
      "layout": "cover",
      "action_title": "Deck Title; Department subtitle",
      "section_number": null,
      "style": "inherit",
      "category": null,
      "takeaway": null,
      "source": null,
      "notes": null,
      "content": { "authors": ["Name 1"] }
    }
  ]
}
```
