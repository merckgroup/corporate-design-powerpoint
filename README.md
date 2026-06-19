# Merck Corporate Design PPTX Tool

Convert a markdown brief, an existing PowerPoint, or a hand-crafted slide plan into a polished **Merck Healthcare KGaA-branded `.pptx` deck** — without opening PowerPoint.

## Quick Start

```bash
# From a markdown brief (calls Claude — needs AIP credentials)
python -m merck_pptx generate brief.md output/deck.pptx --defaults

# From an existing PowerPoint
python -m merck_pptx generate old_deck.pptx output/deck_merck.pptx --defaults

# From a finished slide plan (no Claude needed)
python -m merck_pptx build plan.json output/deck.pptx
```

> **Setup and credentials:** [SETUP.md](SETUP.md)

---

## What it does

| Input | What happens |
|---|---|
| **Markdown document** (`.md`) | Claude reads your content, generates a slide plan, and builds the deck |
| **Existing PowerPoint** (`.pptx` / `.ppt`) | Claude extracts the content, re-plans it in Merck style, and builds it fresh |
| **Plan JSON** (`.json` or Python dict) | Builds directly — no Claude, fully deterministic |

Every output deck follows the Merck brand: Merck purple, gold accents, action titles on every slide, takeaway bands, classification badges, and a consistent footer.

### Starting from a markdown brief

Write your content in any markdown document — no special format required. Headings, bullet lists, tables, and free-form prose all work. Claude reads the brief, structures a slide plan, and builds the deck. Use `--save-plan` to capture the generated plan as JSON so you can tweak and rebuild without another LLM call.

A good brief includes: the core message or recommendation, supporting evidence, and key decisions or next steps. The richer the content, the better Claude can structure the narrative.

---

## Three ways to call it

### 1 — CLI, interactive

Run `generate` without `--meta` or `--defaults` and the pipeline asks six questions before calling Claude:

```bash
python -m merck_pptx generate brief.md output/deck.pptx
```

The prompts cover region, classification, month/year, audience, visual style, and variety mode. The default style adapts to your audience: **executive** for Executive leadership and Senior management; **corporate** for all other audiences.

### 2 — CLI, scripted (no prompts)

```bash
# Provide answers as a JSON file
python -m merck_pptx generate brief.md output/deck.pptx --meta meta.json

# Use built-in defaults (EU, Internal, Executive leadership, merck_executive)
python -m merck_pptx generate brief.md output/deck.pptx --defaults
```

Save the generated plan for inspection or re-use:

```bash
python -m merck_pptx generate brief.md output/deck.pptx \
    --meta meta.json \
    --save-plan output/plan.json

# Tweak plan.json, then rebuild without LLM:
python -m merck_pptx build output/plan.json output/deck_revised.pptx
```

### 3 — Python API

```python
from merck_pptx import generate_deck, build_from_plan

# Markdown or pptx → Merck deck (calls Claude)
generate_deck("brief.md", "output/deck.pptx", meta={
    "region":         "EU",
    "deck_label":     "Q2 Finance Review",
    "classification": "Confidential",
    "month_year":     "June 2026",
    "audience":       "Executive leadership",
    "deck_style":     "merck_executive",
})

# Plan dict or file → Merck deck (no Claude)
build_from_plan(plan_dict, "output/deck.pptx")
build_from_plan("plan.json", "output/deck.pptx")
```

---

## Meta fields

Pass deck identity as a JSON file with `--meta`:

```json
{
  "region":         "EU",
  "division":       "merck",
  "color_theme":    "plastic",
  "deck_label":     "Finance Systems Review",
  "classification": "Confidential",
  "month_year":     "June 2026",
  "audience":       "Executive leadership",
  "deck_style":     "merck_executive",
  "variety_mode":    "default",
  "show_disclaimer": false,
  "chrome": {
    "progress_bar":      false,
    "section_circles":   false,
    "takeaway_bands":    false,
    "footer_breadcrumb": false
  }
}
```

| Field | Required | Valid values | Default |
|---|---|---|---|
| `region` | Yes | `EU` · `USA` | — |
| `division` | No | `merck` · `emd_electronics` · `emd_serono` · `millipore_sigma` · `merck_asia` | `merck` |
| `color_theme` | No | `plastic` · `functional` · `organic` · `synthetic` · `technical` · `electronics` | `plastic` |
| `deck_label` | Yes | Any text — shown in the footer of every slide | — |
| `classification` | Yes | `Public` · `Internal` · `Confidential` | — |
| `month_year` | Yes | e.g. `"June 2026"` | — |
| `audience` | Yes | `Executive leadership` · `Senior management` · `Functional team` · `Mixed audience` · `External / client-facing` | — |
| `deck_style` | No | `merck_executive` · `merck_corporate` · `merck_storytelling` | `merck_executive` |
| `variety_mode` | No | `default` (standard layouts) · `creative` (extended set) | `default` |
| `show_disclaimer` | No | `true` for external-facing decks | `false` |
| `chrome` | No | Object — opt-in custom chrome elements (see below) | `{}` = all off |

### `meta.chrome` — opt-in custom chrome elements

By default the pipeline produces **standard empower output**: Merck logo, classification badge, and page number only. Set any flag to `true` to add that element:

| Flag | Effect |
|---|---|
| `progress_bar` | Thin proportional fill strip at the very top of each content slide |
| `section_circles` | Numbered purple circles + spaced-caps category tag (top-left) |
| `takeaway_bands` | Purple takeaway band above the footer; the LLM writes `takeaway` text only when this is `true` |
| `footer_breadcrumb` | `"Deck Label • Category"` left-aligned in the slide footer |

All four are independent. Without this block the output matches standard empower exactly.

**`region` is a compliance issue.** EU decks use the Merck KGaA (Darmstadt) template. USA/Canada decks carry a legal disclaimer restricting them to North America. Never mix the two.

**`Secret` and above block the build.** If `Secret`, `Top Secret`, or `TS/SCI` is entered as classification, the pipeline exits immediately.

> **How `division` and `color_theme` interact:** `division` selects the template *file* (logo and disclaimer). `color_theme` changes the *colour palette* (applied in code, no extra file needed). Set them independently — any combination works.

---

## Templates and visual options

### Region

| Value | Template family | When to use |
|---|---|---|
| `EU` *(default)* | `EU_Merck_Themed.pptx` | Merck KGaA (Darmstadt) — all EU/global audiences |
| `USA` | `USA_Merck_Themed_Base_v1.pptx` | North America only — carries the legal EMD disclaimer |

### Division

`division` selects the sub-brand template file (logo, wordmark, disclaimer). Missing files fall back to the region default. See [SETUP.md](SETUP.md) for the full template file table and instructions for adding new division templates.

### Color theme

Applied programmatically — no extra template file needed.

| Theme | Cover background | Accent | Best for |
|---|---|---|---|
| `plastic` *(default)* | Lime green `#A5CD50` | Pink `#EB3C96` | General-purpose |
| `functional` | Lime green `#A5CD50` | Teal `#2DBECD` | Life science, biology |
| `organic` | Cream `#FFDCB9` | Red `#E61E50` | Healthcare, patient focus |
| `synthetic` | Violet `#503291` | Yellow `#FFC832` | Industrial, chemistry |
| `technical` | Cream `#FFDCB9` | Teal `#2DBECD` | Engineering, IT |
| `electronics` | Violet `#503291` | Yellow `#FFC832` | EMD Electronics (photo cover) |

The `electronics` cover uses a "Title with picture" layout — the image placeholder is left empty for you to fill in PowerPoint after generation.

### Visual style

| Style | Background | Best for |
|---|---|---|
| `merck_executive` | White | Board decks, formal decisions |
| `merck_corporate` | White | Project updates, town halls, general business |
| `merck_storytelling` | Dark purple | Product launches, change management |

**Auto-promotion:** Five slide types are always forced to `merck_executive` regardless of the deck's style: Executive Summary, Recommendation, Decision Request, Risk, and Tradeoff.

---

## Building from a plan

To control every slide yourself, write a plan JSON and pass it to `build`. Claude is not called — the build is fully deterministic.

**Key rules:**

- `section_number` must be a unique sequential integer on every content slide. Structural slides (Cover, Agenda, Section Divider, Close) use `null`.
- `style: "inherit"` resolves to the deck's `deck_style` at build time.
- The **agenda auto-fills**: leave `content.chapters` empty and the pipeline derives the chapter list from your content slides automatically.
- **Appendix slides**: add `"appendix": true` to any slide to exclude it from the main page count. Appendix slides are numbered A1, A2… in the footer.

For the full plan schema, slide field reference, and all 46 layout content payloads, see [`merck_pptx/slide_plan_schema.md`](merck_pptx/slide_plan_schema.md).

---

## Layout catalog

46 layouts across two tiers. `variety_mode: "default"` covers everyday consulting decks; `"creative"` unlocks the full set.

**Cover and navigation** — `cover` · `exec_summary` · `agenda` · `section_divider` · `close`

**Evidence and data** — `chart_slide` · `waterfall_slide` · `stat_strip` · `hero_stat` · `donut_chart` · `radar_chart` · `risk_heatmap` · `kpi_dashboard` · `score_table` · `word_cloud`

**Argument and structure** — `two_column` · `three_column` · `four_column` · `columns` · `vertical_numbered` · `label_rows` · `before_after` · `pros_cons` · `2x2_matrix`

**Process and timelines** — `phase_process` · `gantt` · `milestone_timeline` · `circular_flow` · `arrow_chain` · `funnel` · `journey_map` · `road_to_success`

**Decisions and tables** — `decision_rows` · `status_table` · `comparison_table`

**Organisation and relationships** — `org_chart` · `hub_spoke` · `pillar_detail` · `topic_set` · `icon_grid`

**Visual and narrative** — `pull_quote` · `key_question` · `pyramid` · `venn` · `layered_stack` · `influence_diagram` · `photo_text` · `fishbone`

For layout use-cases, content schemas, and the pre-send checklist, see [Merck_Presentation_Helper.md](Merck_Presentation_Helper.md).

> **Note on Merck branded visual assets (Mercrobes, etc.):** Some empower library elements — including the 3D organic sphere icons ("Mercrobes") — exist only on the empower server and are not embedded by this pipeline. The pipeline generates all shapes programmatically using python-pptx.

---

## Companion documents

| Document | For whom | What it covers |
|---|---|---|
| [SETUP.md](SETUP.md) | First-time setup | Installation, credentials, division template files, troubleshooting |
| [Merck_Presentation_Helper.md](Merck_Presentation_Helper.md) | Humans building a deck | Layout picker, content schemas, color rules, pre-send checklist |
| [merck_pptx/slide_plan_schema.md](merck_pptx/slide_plan_schema.md) | Anyone writing a plan JSON | Full plan schema, all 46 layout payloads, field name rules |
| [LLM_PLAN_GUIDE.md](LLM_PLAN_GUIDE.md) | LLMs generating a plan | Same as above, optimised for LLM consumption |
| [Merck_Presentation_Guidelines.md](Merck_Presentation_Guidelines.md) | Brand governance | Exact color values, typography, accessibility, brand rules |

---

## Origin

This pipeline is the code implementation of the **Merck Slide Agent** MyGPT by **Anoop Kumar (LS-CL-CD)**:
[mygpt-suite.uptimize.merckgroup.com — Merck Slide Agent](https://mygpt-suite.uptimize.merckgroup.com/chat?a=7cdc5dfe-47a8-4009-ac3f-4f95f6a3114e)

The MyGPT agent runs interactively in the Merck internal environment and produces slide plans through a guided conversation. This repository takes that agent's structural discipline — layout catalog, quality rules, brand enforcement, and plan schema — and turns it into a standalone Python pipeline that any tool or script can call programmatically.
