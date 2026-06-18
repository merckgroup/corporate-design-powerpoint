# Merck Corporate Design PPTX Tool

Convert an existing PowerPoint, a markdown brief, or a structured slide plan into a polished **Merck Healthcare KGaA-branded deck** without opening PowerPoint.

> **Installation and setup:** see [SETUP.md](SETUP.md).

---

## Origin

This pipeline is the code implementation of the **Merck Slide Agent** MyGPT by **Anoop Kumar (LS-CL-CD)**:
[mygpt-suite.uptimize.merckgroup.com — Merck Slide Agent](https://mygpt-suite.uptimize.merckgroup.com/chat?a=7cdc5dfe-47a8-4009-ac3f-4f95f6a3114e)

The MyGPT agent runs interactively in the Merck internal environment and produces slide plans through a guided conversation. This repository takes that agent's structural discipline: layout catalog, quality rules, brand enforcement, and plan schema — and turns it into a standalone Python pipeline that any tool or script can call programmatically.

---

## What it does

The pipeline accepts three types of input and always produces a `.pptx` file:

| Input | What happens |
|---|---|
| **Markdown document** (`.md`) | Claude reads your content, generates a slide plan, and builds the deck |
| **Existing PowerPoint** (`.pptx` / `.ppt`) | Claude extracts the content, re-plans it in Merck style, and builds it fresh |
| **Plan JSON** (`.json` or Python dict) | Builds directly, no Claude, fully deterministic |

Every output deck follows the Merck brand: Merck purple, gold accents, action titles on every slide, takeaway bands, classification badges, section circles, and a consistent footer.

---

## Three ways to call it

### 1 — Command line, interactive

Run `generate` without `--meta` or `--defaults` and the pipeline asks six questions before calling Claude:

```bash
python -m merck_pptx generate brief.md output/deck.pptx
```

```
Merck PPTX Pipeline — deck setup

[1/6] Region (EU / USA) [EU]:
[2/6] Classification (Public / Internal / Confidential) [Internal]:
[3/6] Month and year [June 2026]:
[4/6] Audience:
      1. Executive leadership
      2. Senior management
      3. Functional team
      4. Mixed audience
      5. External / client-facing
      Choice [1]:
[5/6] Visual style (executive / corporate / storytelling) [executive]:
[6/6] Variety mode (default / creative) [default]:

Deck label / title (appears in footer) [Merck Presentation]:
```

The interactive prompts cover the core six fields. To also set `division` or `color_theme`, use `--meta` (see below).

### 2 — Command line, scripted (no prompts)

Three ways to skip all prompts:

```bash
# --meta: supply your answers as a JSON file
python -m merck_pptx generate brief.md output/deck.pptx --meta meta.json

# --defaults: use built-in defaults (EU, Internal, Executive leadership, merck_executive)
python -m merck_pptx generate brief.md output/deck.pptx --defaults

# Pipe / CI: stdin not a TTY automatically activates defaults
echo "" | python -m merck_pptx generate brief.md output/deck.pptx
```

Use `--save-plan` to save the LLM-generated slide plan as JSON alongside the deck — useful for inspecting or tweaking the plan and rebuilding without another LLM call:

```bash
python -m merck_pptx generate brief.md output/deck.pptx \
    --meta meta.json \
    --save-plan output/plan.json

# Tweak plan.json, then rebuild without LLM:
python -m merck_pptx build output/plan.json output/deck_revised.pptx
```

Re-brand an existing PowerPoint or build from a finished plan:

```bash
# Re-brand an existing PowerPoint
python -m merck_pptx generate old_deck.pptx output/deck_merck.pptx --defaults

# Build from a finished plan (no Claude required)
python -m merck_pptx build plan.json output/deck.pptx
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

# Optionally save the generated plan for inspection or re-use
generate_deck("brief.md", "output/deck.pptx", meta={...}, save_plan="output/plan.json")

# Plan dict or file → Merck deck (no Claude)
build_from_plan(plan_dict, "output/deck.pptx")
build_from_plan("plan.json", "output/deck.pptx")
```

---

## Meta fields (`meta.json`)

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
  "deck_style":     "merck_executive"
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

**`region` is a compliance issue.** EU decks use the Merck KGaA (Darmstadt) template. USA/Canada decks carry a legal disclaimer restricting them to North America. Never mix the two.

**`Secret` blocks the build.** If `Secret` is entered as classification, the pipeline exits immediately.

> **How `division` and `color_theme` interact:** `division` selects the template *file* (determines which logo and disclaimer appear on every slide). `color_theme` changes the *colour palette* (applied in code, no extra file needed). Set them independently — any combination works.

---

## Template selection — how the tool picks the right file

Three meta fields control the visual output. They are **independent** — any combination works:

```
region      →  which brand template FILE is loaded   (compliance)
division    →  which sub-brand within that region    (logo / disclaimer)
color_theme →  which colour scheme is applied        (design, no extra file needed)
```

### `region` — compliance boundary

| Value | Template family | When to use |
|---|---|---|
| `EU` (default) | `EU_Merck_Themed.pptx` | Merck KGaA, Darmstadt identity — all EU/global audiences |
| `USA` | `USA_Merck_Themed_Base_v1.pptx` | North America only — carries the legal EMD disclaimer |

⚠ **Region is a compliance issue.** USA templates carry a legal disclaimer that restricts them to North America. Never send a USA-region deck to a non-NA audience.

### `division` — brand / logo variant

Selects the template file for your sub-brand. Each file carries the correct logo, wordmark and disclaimer for that division. Missing files fall back to the Merck default for that region — no error.

| Value | Template file (EU) | Template file (USA) |
|---|---|---|
| `merck` *(default)* | `EU_Merck_Themed.pptx` ✅ included | `USA_Merck_Themed_Base_v1.pptx` ✅ included |
| `emd_electronics` | `EU_EMDElectronics_Themed.pptx` | `USA_EMDElectronics_Themed.pptx` |
| `emd_serono` | `EU_EMDSerono_Themed.pptx` | `USA_EMDSerono_Themed.pptx` |
| `millipore_sigma` | `EU_MilliporeSigma_Themed.pptx` | `USA_MilliporeSigma_Themed.pptx` |
| `merck_asia` | `EU_MerckAsia_Themed.pptx` | `USA_MerckAsia_Themed.pptx` |

**Adding a division template:** export the master template from empower (`Corporate Design Templates → Master Templates → {Division}`) as a `.pptx` file, name it according to the table above, and place it in `merck_pptx/templates/`. The pipeline picks it up automatically.

### `color_theme` — colour scheme

Applied **programmatically** — no extra template file needed. The pipeline modifies the loaded template's PowerPoint theme XML in memory before building any slides, so every cover, section divider, and content slide matches the chosen colour palette.

| Value | Background | Accent colour | Best for |
|---|---|---|---|
| `plastic` *(default)* | Lime green `#A5CD50` | Pink `#EB3C96` | General-purpose Merck Plastic cell design |
| `functional` | Lime green `#A5CD50` | Teal `#2DBECD` | Life science, biology, organic cell imagery |
| `organic` | Cream `#FFDCB9` | Red `#E61E50` | Healthcare, patient focus, warm narrative |
| `synthetic` | Violet `#503291` dark | Yellow `#FFC832` | Industrial, chemistry, manufacturing |
| `technical` | Cream `#FFDCB9` | Teal `#2DBECD` | Engineering, IT, technical audiences |
| `electronics` | Violet `#503291` dark | Yellow `#FFC832` | EMD Electronics; cover has an editable photo placeholder |

The `electronics` cover uses the template's "Title with picture" layout. The image placeholder is left empty — fill it in PowerPoint after generation.

### Putting it together — examples

```json
{ "region": "EU",  "division": "merck",            "color_theme": "plastic"     }
{ "region": "EU",  "division": "merck",            "color_theme": "organic"     }
{ "region": "USA", "division": "emd_electronics",  "color_theme": "electronics" }
{ "region": "EU",  "division": "millipore_sigma",  "color_theme": "synthetic"   }
```

Omit any field to get its default (`EU` · `merck` · `plastic`).

---

## Visual styles

Three locked palettes control card backgrounds, typography, and chrome colour on **content slides**. The `deck_style` is orthogonal to `color_theme` — you can use any style with any colour theme.

| Style | Background | Character | Best for |
|---|---|---|---|
| `merck_executive` | White | Grayscale + single accent | CFO updates, board presentations, M&A |
| `merck_corporate` | White | Purple, yellow, aqua, gold | Project updates, town halls, cross-functional |
| `merck_storytelling` | Dark purple (`#3A2468`) | Gold dominant, white text | Product launches, change management |

**Auto-promotion.** Five slide types are always forced to `merck_executive` style, regardless of the deck's chosen style: **Executive Summary**, **Recommendation**, **Decision Request**, **Risk**, and **Tradeoff**.

---

## Building from a plan

To control every slide yourself, write a plan JSON and pass it to `build_from_plan`. Claude is not called — the build is fully deterministic.

### Plan shape

```json
{
  "meta": {
    "region":                    "EU",
    "division":                  "merck",
    "color_theme":               "organic",
    "deck_label":                "Q2 Finance Review",
    "classification":            "Confidential",
    "month_year":                "June 2026",
    "audience":                  "Executive leadership",
    "deck_style":                "merck_executive",
    "deck_objective":            "Brief CFO on Lumina readiness and unblock IT package handover.",
    "single_sentence_takeaway":  "We are on track for Q1 but Q2 risks a 4-week slip.",
    "final_ask":                 "Three approvals: Data Owner, IT date, Fabric capacity."
  },
  "storyline": [
    "Q1 audit readiness is on track at 92% completion.",
    "Q2 data products are gated on the IT package handover.",
    "Three decisions unblock the slip risk."
  ],
  "slides": [ ... ]
}
```

### Slide shape

```json
{
  "page":           5,
  "page_function":  "Diagnosis",
  "layout":         "two_column",
  "action_title":   "Legacy Systems Drive 60% of IT Costs; Migration Is Overdue",
  "section_number": 1,
  "category":       "DIAGNOSIS",
  "takeaway":       "Deferring migration adds €12M in annual maintenance cost.",
  "source":         "Source: IT Cost Audit, Q1 2026",
  "style":          "inherit",
  "content": {
    "left":  {"label": "Current state", "items": ["SAP ECC 6.0 — EOL 2027", "14 siloed databases"]},
    "right": {"label": "Impact",        "items": ["60% of IT opex", "3-week reporting lag"]}
  }
}
```

**Key rules:**

- `section_number` must be a unique sequential integer (1, 2, 3…) on every content slide. Structural slides — Cover, Agenda, Section Divider, Close — use `null`.
- `style: "inherit"` resolves to the deck's `deck_style` at build time. Set it explicitly only to override for a specific slide.
- The **agenda auto-fills**: leave `content.chapters` empty and the pipeline derives the chapter list from your content slides automatically.
- **Appendix slides**: add `"appendix": true` to any slide to exclude it from the main page count. Appendix slides are numbered A1, A2… in the footer.

---

## Action title quality

The action title is the single most important quality signal. Every content slide needs one.

| Rule | Example |
|---|---|
| Always a declarative sentence | `"Two Sites Drag the Portfolio; Filler Replacement Resolves Both"` |
| Never a noun phrase or topic label | ~~`"Portfolio Overview"`~~ |
| Preferred format: Insight; Consequence | `"OEE Improved 5.4 Points to 84%; Premium SKU Mix Is the Driver"` |
| Numbers beat adjectives | `"Revenue Up 8% to €4.2B"` not `"Strong Revenue Growth"` |
| Max 80 characters | — |
| Semicolons, not dashes | ~~`"Revenue grew — best quarter ever"`~~ |

For the highest-impact slides (Executive Summary, Recommendation, Decision Request), you can italicise a run within the title:

```json
"action_title": [
  ["Three decisions today ", false],
  ["unlock the Q3 trajectory.", true]
]
```

---

## Layout catalog

46 layouts across two tiers. The **default** catalog covers everyday consulting decks. Set `variety_mode: "creative"` in meta to unlock the full set.

### Cover and navigation

| Layout | Use for |
|---|---|
| `cover` | Title slide. `action_title` ≤ 60 chars; `subtitle` required |
| `exec_summary` | One-page argument summary with 3–5 key messages |
| `agenda` | Chapter index — auto-fills from content slides if left empty |
| `section_divider` | Visual chapter break |
| `close` | Closing action statement |

### Evidence and data

| Layout | Use for |
|---|---|
| `chart_slide` | Any chart: slope, bar, waterfall, marimekko, dot plot, small multiples |
| `waterfall_slide` | Financial bridge as the primary slide element |
| `stat_strip` | 3–4 headline statistics in a row |
| `hero_stat` | One large KPI with italic context line |
| `donut_chart` | Market share, completion rates, part-of-whole (2–8 segments) |
| `radar_chart` | Multi-axis capability or maturity spider (4–8 axes, 0–100) |
| `risk_heatmap` | 5×5 probability × impact grid with numbered dots |
| `kpi_dashboard` | 4–6 metrics with RAG dots and optional sparklines |
| `score_table` | Filled-dot maturity ratings on a 1–5 scale |
| `word_cloud` | Survey verbatims, topic frequency (weight 1–5) |

### Argument and structure

| Layout | Use for |
|---|---|
| `two_column` | Current state vs target state, problem vs solution |
| `three_column` | Three parallel options, phases, or pillars |
| `four_column` | Four parallel workstreams |
| `columns` | Auto-dispatches to 2/3/4-column by the number of items you provide |
| `vertical_numbered` | 3–5 numbered action items |
| `label_rows` | Named barriers or root-cause taxonomy (3–6 rows) |
| `before_after` | Side-by-side comparison with gold directional arrow |
| `pros_cons` | Green pros panel vs red cons panel |
| `2x2_matrix` | Quadrant analysis — effort × impact, risk × value, etc. |

### Process and timelines

| Layout | Use for |
|---|---|
| `phase_process` | Sequential phases with done / current / future status (max 5) |
| `gantt` | Multi-row timeline across quarters |
| `milestone_timeline` | Chronological event sequence with status dots |
| `circular_flow` | Continuous loop — Plan/Do/Check/Act etc. (2–8 phases) |
| `arrow_chain` | Causal chain: Trigger → Response → Effect → Outcome + consequence box |
| `funnel` | Multiple inputs converging to one outcome |
| `journey_map` | Multi-actor swim-lane: phases across the top, actors down the side |

### Decisions and tables

| Layout | Use for |
|---|---|
| `decision_rows` | Decision log: number, title, description, owner, tone (max 5) |
| `status_table` | RAG status table — RISK / STATUS / HEALTH column auto-detected |
| `comparison_table` | Feature matrix: options across the top, features down the side; yes / no / partial |

### Organisation and relationships

| Layout | Use for |
|---|---|
| `org_chart` | Reporting hierarchy (2–3 levels, named nodes) |
| `hub_spoke` | Four spokes radiating from a central theme |
| `pillar_detail` | One pillar deep-dive: number, owner, and 3–5 sections |
| `topic_set` | 3–6 equal-weight principles, capabilities, or features |
| `icon_grid` | 6 or 9 equal icon-card cells (2- or 3-column grid) |

### Visual and narrative

| Layout | Use for |
|---|---|
| `pull_quote` | Full-bleed hero quote or vision statement |
| `pyramid` | Strategy hierarchy or value tiers (3–5 tiers, `"up"` or `"down"`) |
| `venn` | Overlap of 2–3 audiences or capabilities |
| `layered_stack` | Technology architecture or platform layers (up to 7) |
| `influence_diagram` | Force map from four sides — SWOT alternative |
| `photo_text` | Half image + half text narrative (culture, site visit, story) |
| `fishbone` | Ishikawa root-cause diagram (up to 6 bones, 3 sub-causes each) |

---

## Common mistakes

These are the most frequent errors when writing plan JSON manually. Wrong content keys produce silently empty slides — no error is raised.

| Layout | Use this key | Not this |
|---|---|---|
| `two_column`, `before_after` columns | `items` | `bullets`, `points`, `list` |
| `decision_rows` decisions | `text` (single string body) OR `title` + `desc` | `decision`, `body`, `description` |
| `before_after` panel labels | `before.label` / `after.label` (nested) OR `before_label` / `after_label` (flat) | hard-coded TODAY/TOMORROW |
| `milestone_timeline` milestones | `label` + `description` (or `title` + `body`) | any other key names |
| `milestone_timeline` status | `"upcoming"` · `"completed"` · `"active"` | `"future"` · `"done"` · `"current"` |
| `status_table` | `rows` only — columns auto-derived | explicit `columns` not required |
| `comparison_table` | `headers` + `rows` (matrix) OR `options` + `features` | any other key names |
| `waterfall_slide` bar types | `"total"` · `"positive"` · `"negative"` | `"start"` · `"up"` · `"down"` |
| `risk_heatmap` scores | `likelihood: 1–5` + `impact: 1–5` (integers) | `"high"` · `"low"` strings |
| `journey_map` | `rows` (with `actor` + `steps`) OR `actors` (with `name` + `cells`) | — |
| `chart_slide` data | `chart: {type, data: {categories, series}}` | `items` list format |
| `chart_slide` layout key | `chart_slide` | `chart` |
| `waterfall_slide` layout key | `waterfall_slide` | `waterfall` |
| `2x2_matrix` layout key | `2x2_matrix` | `matrix_2x2` |
| `2x2_matrix` quadrant cells | `items` (list of strings) | `body`, `text`, `content` |
| `phase_process` highlight | `highlighted: true` | `status: "current"` |
| `funnel` | `inputs` + `output` | `stages` |
| `fishbone` | `bones` | `causes` |
| `influence_diagram` forces | `forces` (with `side` + `tone`) | `nodes` |
| `venn` overlap text | `intersection` | `overlap` |
| `hub_spoke` center | `hub: {label, title, subtitle}` | `center: "string"` |
| `close` body | `action_statement` | `statement` |

---

## Full content schema reference

The complete field-by-field schema for all 46 layouts — including content payload shapes, valid values, and worked examples — is in [`merck_pptx/slide_plan_schema.md`](merck_pptx/slide_plan_schema.md).

---

## Installation and setup

See [SETUP.md](SETUP.md).
