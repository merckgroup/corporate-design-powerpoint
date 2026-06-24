# Merck Corporate Design PPTX Tool

Turn your content — a markdown document, an existing PowerPoint, or a detailed slide plan — into a polished **Merck Healthcare KGaA-branded `.pptx` deck**, automatically.

The tool handles all brand rules: the right template for your region, Merck colors and fonts, action titles on every slide, footer badges, and 46 layout types. You provide the content and make six decisions.

---

## What it creates

Every generated deck includes:

- **Cover** with your title, authors, and date
- **Agenda** derived automatically from your chapter structure
- **Section dividers** and a closing slide
- **Content slides** built to the Merck Corporate Design standard — correct template for your region, Merck fonts, purple action titles, classification badge, and page numbers

All layout types are rendered programmatically — no manual PowerPoint editing needed for structure or branding.

→ For the full brand specification: [Merck_Presentation_Guidelines.md](Merck_Presentation_Guidelines.md)

---

## Getting started

**First-time setup** (Python environment, AIP credentials): → [SETUP.md](SETUP.md)

Once set up, the fastest way to generate a deck:

```bash
python -m merck_pptx generate brief.md output/deck.pptx --defaults
```

This uses sensible defaults (EU region, Internal classification, Executive leadership audience, `merck_executive` style) and asks no questions. Drop `--defaults` to answer the six questions interactively instead.

---

## Three input types

| Input | What happens |
|---|---|
| **Markdown document** (`.md`) | Claude reads your content, structures a slide plan, and builds the deck |
| **Existing PowerPoint** (`.pptx` / `.ppt`) | Claude extracts the content, re-plans it in Merck style, and builds it fresh |
| **Plan JSON** (`.json`) | Builds directly — no Claude, fully deterministic |

---

## Writing your brief

Write your content in any markdown file — no special format required. Headings, bullet lists, tables, and prose all work. A good brief includes:

- The core message or recommendation
- Supporting evidence (data, findings, key results)
- Key decisions or next steps

The richer your content, the better the slide structure. A one-sentence brief produces a minimal plan; a detailed 5-page document typically maps to 10–20 slides.

**Inspect and refine the plan before building:**

```bash
python -m merck_pptx generate brief.md output/deck.pptx --save-plan output/plan.json
# Edit plan.json, then rebuild without calling Claude again:
python -m merck_pptx build output/plan.json output/deck_revised.pptx
```

This is useful when you want to adjust layout choices, reorder slides, or fine-tune wording without paying for another LLM call. See [merck_pptx/slide_plan_schema.md](merck_pptx/slide_plan_schema.md) for the plan format.

---

## The five decisions that shape your deck

When you run `generate` without `--defaults`, the tool asks five questions. Here is what each choice actually does to your deck.

### 1. Region and division

`EU` or `USA`. This is a **compliance requirement**, not just a visual option.

- **EU** — uses the Merck KGaA (Darmstadt) template and disclaimer. For all global and European audiences.
- **USA** — carries the EMD/MilliporeSigma legal disclaimer. For North America only.

Never use an EU deck for a USA audience and vice versa.

For **USA**, the tool immediately asks which division:

| # | Division key | Branding |
|---|---|---|
| 1 | `emd_serono` | EMD Serono — Healthcare *(default)* |
| 2 | `millipore_sigma` | MilliporeSigma — Life Science |
| 3 | `emd_electronics` | EMD Electronics |
| 4 | `usa` | US tri-brand — cross-business North America |

For **EU**, division defaults to `merck` (Merck KGaA). If you need Asia/China branding (`merck_asia`), pass it via `--meta` (see below).

> **Note on EU business units:** Merck Healthcare, Merck Life Science, and Merck Electronics do not have separate empower templates — all EU business units share the standard Merck KGaA template set, differentiated by `color_theme` only (e.g. `organic` for Healthcare, `functional` for Life Science, `electronics` for Electronics).

### 2. Audience

Who will read this deck. This is the most impactful setting — it drives tone, layout density, and how Claude structures the narrative throughout the deck.

| Audience | Effect on the deck |
|---|---|
| `Executive leadership` | Tight layouts, concise text, strong action titles. Key slides (Executive Summary, Recommendation, Decision Request, Risk, Tradeoff) automatically switch to the most formal visual style. |
| `Senior management` | Balanced — enough context for informed reading, structured for quick scanning. |
| `Functional team` | More detail allowed, more granular layouts, less hierarchical structure. |
| `Mixed audience` | Broadly readable balance of depth and brevity. |
| `External / client-facing` | Client-appropriate tone; consider enabling `show_disclaimer` for externally shared decks. |

You can type any free text — the values above are the recommended choices.

### 3. Deck style

Controls the visual weight and feel of every **content slide**. Color theme (below) is a separate, independent setting that controls the cover palette — the two do not interact.

| Style | What it looks like | Best for |
|---|---|---|
| `merck_executive` | White background, rich purple headings, tight formal layout | Board decks, C-suite sign-offs, formal decisions |
| `merck_corporate` | White background, standard proportions, moderate density | Project updates, town halls, general business decks |
| `merck_storytelling` | **Dark purple** background, white text, bold and visual | Product launches, change management, communications |

**Some slides always look formal, regardless of your choice.** If your content includes an Executive Summary, a Recommendation, a Decision Request, a Risk, or a Tradeoff slide, those specific slides will always use the tight, authoritative `merck_executive` look — even if the rest of your deck is in `merck_corporate` or `merck_storytelling`. The reasoning: a budget decision or a risk slide always deserves the most serious visual treatment, whatever the surrounding deck looks like.

### 4. Color theme

Controls the **cover slide and section divider palette only** — not the content slides. Think of it as the "opening color" of your deck. Content slide backgrounds are always white (or dark purple for `merck_storytelling`) regardless of which theme you pick.

| Theme | Cover & dividers | Accent | Best for |
|---|---|---|---|
| `plastic` *(default)* | Lime green | Pink | General purpose |
| `functional` | Lime green | Teal | Life science, biology, cells |
| `organic` | Cream | Red | Healthcare, patient-focused |
| `synthetic` | Dark violet | Yellow | Industrial, chemistry |
| `technical` | Cream | Teal | Engineering, IT, data |
| `electronics` | Dark violet | Yellow | EMD Electronics (adds a photo placeholder on the cover) |

For biology and life science presentations, `functional` is the natural fit. For healthcare and patient-facing content, choose `organic`.

### 5. Classification

Sets the badge shown on every slide footer.

| Value | When to use |
|---|---|
| `Public` | Content approved for external audiences |
| `Internal` | For Merck employees only |
| `Confidential` | Sensitive internal content |

`Secret` and above will immediately stop the build — do not use these.

### Variety mode — leave it at `default`

There is a sixth parameter, `variety_mode`, but for most decks you do not need to think about it. Leave it at `default`.

`default` gives Claude 30 layouts — the standard set for business, science, and project presentations. `creative` adds 16 more visually distinctive layouts: word cloud, fishbone (cause-and-effect), pyramid, layered stack, photo panels. Use `creative` only if you are making a communications or research presentation and specifically want those styles.

> The reason this parameter exists: with 46 options available, Claude occasionally reaches for an unusual layout (a word cloud where a two-column would be better). `default` keeps choices conservative and predictable. Your audience and deck style already guide layout density and formality — variety mode is simply a guardrail on the full layout toolkit.

---

## Setting color theme and other parameters via `--meta`

The interactive flow covers the most common settings. One parameter it does **not** ask for is `color_theme` (and for EU, `division` when you need Asia/China branding). To set these, pass a `--meta` JSON file.

The interactive flow also offers two optional narrative context prompts at the end — you can skip them by pressing Enter and Claude will derive them from the source document.

### Full `--meta` reference

Create a JSON file with any combination of the fields below and pass it with `--meta`. Fields you omit fall back to sensible defaults.

```json
{
  "region":          "EU",
  "division":        "merck",
  "classification":  "Internal",
  "month_year":      "June 2026",
  "audience":        "Senior management",
  "deck_style":      "merck_corporate",
  "color_theme":     "organic",
  "variety_mode":    "default",
  "deck_label":      "Q2 Business Review",
  "show_disclaimer": false,

  "topic":                    "Q2 readiness and IT handover decisions",
  "deck_objective":           "Brief the CFO on Q2 readiness and unblock three IT decisions.",
  "single_sentence_takeaway": "We are on track for Q1 audit but Q2 products risk a 4-week slip without three approvals.",
  "final_ask":                "Approve Tom Kistinger as Data Owner, confirm IT date, authorize Fabric capacity."
}
```

The four `topic` / `deck_objective` / `single_sentence_takeaway` / `final_ask` fields are **optional** — Claude always derives them from the source document if they are absent. When you supply them yourself (because you know exactly what the deck should argue and what you want the audience to decide), Claude produces sharper action titles, tighter takeaway lines, and a more coherent slide sequence.

Use this with any input type:

```bash
# From a markdown brief
python -m merck_pptx generate brief.md output/deck.pptx --meta meta.json

# From an existing PowerPoint
python -m merck_pptx generate old_deck.pptx output/deck.pptx --meta meta.json

# Save the generated plan so you can inspect or tweak it before building
python -m merck_pptx generate brief.md output/deck.pptx --meta meta.json --save-plan plan.json
python -m merck_pptx build plan.json output/deck_revised.pptx
```

**EMD Serono example:**

```json
{
  "region":         "USA",
  "division":       "emd_serono",
  "color_theme":    "organic",
  "classification": "Internal",
  "audience":       "Senior management",
  "deck_style":     "merck_corporate",
  "deck_label":     "EMD Serono Q2 Review"
}
```

```bash
python -m merck_pptx generate brief.md output/deck.pptx --meta emd_serono_meta.json
```

---

## Choosing layouts

Claude picks layouts automatically from your brief content — this is separate from the color theme, which only affects the cover palette. Layout choice is about *structure* (how content is arranged on the slide); color theme is about *palette* (what colors appear on the cover and chapter breaks).

You can guide Claude's layout choices by structuring your markdown:

- A table tends to produce a comparison or status table layout
- A numbered list tends to produce a process or vertical-numbered layout
- A single key number tends to produce a hero stat layout
- A section titled "Decisions" tends to produce a decision rows layout

The 46 layouts are grouped by purpose:

**Cover and navigation** — `cover` · `exec_summary` · `agenda` · `section_divider` · `close`

**Evidence and data** — `chart_slide` · `waterfall_slide` · `stat_strip` · `hero_stat` · `donut_chart` · `radar_chart` · `risk_heatmap` · `kpi_dashboard` · `score_table` · `word_cloud`

**Argument and structure** — `two_column` · `three_column` · `four_column` · `columns` · `vertical_numbered` · `label_rows` · `before_after` · `pros_cons` · `2x2_matrix`

**Process and timelines** — `phase_process` · `gantt` · `milestone_timeline` · `circular_flow` · `arrow_chain` · `funnel` · `journey_map` · `road_to_success`

**Decisions and tables** — `decision_rows` · `status_table` · `comparison_table`

**Organisation and relationships** — `org_chart` · `hub_spoke` · `pillar_detail` · `topic_set` · `icon_grid`

**Visual and narrative** — `pull_quote` · `key_question` · `pyramid` · `venn` · `layered_stack` · `influence_diagram` · `photo_text` · `fishbone`

→ For layout use-cases, content examples, and the pre-send checklist: [Merck_Presentation_Helper.md](Merck_Presentation_Helper.md)

> **Note on Merck visual assets:** Some empower library elements — including the 3D organic sphere icons ("Mercrobes") — exist only on the empower server and are not generated by this pipeline. All shapes are built programmatically.

---

## Going further

| Resource | When to open it |
|---|---|
| [SETUP.md](SETUP.md) | First-time installation, credentials, adding division templates |
| [Merck_Presentation_Helper.md](Merck_Presentation_Helper.md) | Layout picker with use-cases, content schemas, color rules, pre-send checklist |
| [Merck_Presentation_Guidelines.md](Merck_Presentation_Guidelines.md) | Brand rules, exact color values, typography, accessibility |
| [merck_pptx/slide_plan_schema.md](merck_pptx/slide_plan_schema.md) | Writing or editing a plan JSON by hand |
| [LLM_PLAN_GUIDE.md](LLM_PLAN_GUIDE.md) | Technical reference for LLMs or scripts generating a plan; Python API |

---

## Origin

This pipeline is the code implementation of the **Merck Slide Agent** MyGPT by **Anoop Kumar (LS-CL-CD)**:
[mygpt-suite.uptimize.merckgroup.com — Merck Slide Agent](https://mygpt-suite.uptimize.merckgroup.com/chat?a=7cdc5dfe-47a8-4009-ac3f-4f95f6a3114e)

The MyGPT agent runs interactively in the Merck internal environment and produces slide plans through a guided conversation. This repository takes that agent's structural discipline — layout catalog, quality rules, brand enforcement, and plan schema — and turns it into a standalone Python pipeline that any tool or script can call programmatically.
