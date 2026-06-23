# Merck Presentation Builder — User Guide

> A practical helper for creating Merck Corporate Design-compliant presentations with the pipeline.
>
> **This document is for humans building a deck.** It is deliberately concise — use it to make layout decisions, check common content schemas, and run the pre-send checklist.
>
> | Need | Go to |
> |---|---|
> | Which layout to use, content schemas, pre-send checklist | **You are here** |
> | How to use the tool, the six settings, layout overview | `README.md` |
> | Exact color rules, typography, accessibility, brand governance | `Merck_Presentation_Guidelines.md` |
> | Complete JSON schema for all 46 layouts (for LLM use) | `LLM_PLAN_GUIDE.md` |

---

## Quick Start

The fastest way to get a compliant deck is to provide a plan JSON with the right `meta` settings.  
The pipeline handles all colors, fonts, and layout rendering — you focus on **content and structure**.

```bash
# Build directly from a plan (no LLM needed)
python -m merck_pptx build my_plan.json output/deck.pptx

# Generate from a markdown file (LLM converts it to a plan)
python -m merck_pptx generate source.md output/deck.pptx
```

---

## 1. Choosing Your Region and Division

Set `meta.region` to match your audience — this controls the legal disclaimer and template file.

| Region | Value | Template | Legal footer |
|---|---|---|---|
| All countries except USA/Canada | `"EU"` | EU_Merck_Themed.pptx | Merck KGaA disclaimer |
| USA and Canada | `"USA"` | USA_Merck_Themed_Base_v1.pptx | EMD/MilliporeSigma disclaimer |

**Rule:** Never use an EU template for a USA audience and vice versa.

Set `meta.division` to select the correct logo and branding for your business unit.

**EU (`region: "EU"`):**

| Division key | Branding |
|---|---|
| `merck` *(default)* | Merck KGaA — cross-divisional |
| `merck_healthcare` | Merck Healthcare |
| `merck_life_science` | Merck Life Science |
| `merck_electronics` | Merck Electronics |
| `merck_asia` | Merck — Asia/China |

**USA (`region: "USA"`):**

| Division key | Branding |
|---|---|
| `emd_serono` | EMD Serono — Healthcare |
| `millipore_sigma` | MilliporeSigma — Life Science |
| `emd_electronics` | EMD Electronics |
| `usa` | US tri-brand (cross-business) |

---

## 2. Choosing a Visual Style

Set `meta.deck_style`. Every slide in the deck inherits this unless overridden per-slide.

| Style | When to use | Background | Primary accent |
|---|---|---|---|
| `merck_executive` | Board decks, C-suite, formal decisions | White | Rich Purple |
| `merck_corporate` | General business presentations | White | Rich Purple |
| `merck_storytelling` | Impactful narrative, bold statements | Purple | White (inverted) |
| `merck_science` | Pharma lab reports, early-research data decks | White | Merck Blue |

**Auto-promotion:** Any slide with category "Executive Summary", "Recommendation", "Decision Request", "Risk", or "Tradeoff" automatically gets `merck_executive` style — except in `merck_science` decks, where auto-promote is suppressed to preserve the data-first visual language.

**`merck_science` specifics:** Title text uses dark ink instead of purple. Section circles are Merck Blue. The cover shows prominent author bylines (no key_messages grid). Section dividers use a minimal programmatic design (no organic blob shapes). Four exclusive layouts are available: `figure_panel`, `methods_box`, `sar_table`, `multi_chart`.

---

## 3. Choosing a Color Theme

Set `meta.color_theme`. This controls the cover/divider background and the brand accent color.

| Theme | Cover background | Accent | Best for |
|---|---|---|---|
| `plastic` | Light green | Pink | Default, general purpose |
| `functional` | Light green | Teal | Life science, organic cells |
| `organic` | Cream | Red | Healthcare, patient focus |
| `synthetic` | Dark violet | Yellow | Industrial, chemistry |
| `technical` | Cream | Teal | Engineering, IT |
| `electronics` | Dark violet | Yellow | EMD Electronics (photo cover) |

**Dark themes** (`synthetic`, `electronics`): white text is used automatically on dark backgrounds.  
**Electronics theme**: The cover has an image placeholder — fill it in PowerPoint after generating.

---

## 4. Building the Slide Structure

### 4.1 Always Start With These Structural Slides

Every deck should have this backbone (in order):

```
1. cover          → section_number: null
2. agenda         → section_number: null
3. section_divider → section_number: null  (one per chapter)
   ... content slides (section_number: 1, 2, 3 ...)
4. close          → section_number: null
```

### 4.2 Section Numbers

- **Structural slides** (cover, agenda, section_divider, close): always `section_number: null`
- **All other slides**: unique, sequential integers starting at 1 — no gaps, no repeats

```jsonc
// Correct
{ "layout": "cover",           "section_number": null },
{ "layout": "agenda",          "section_number": null },
{ "layout": "section_divider", "section_number": null },
{ "layout": "two_column",      "section_number": 1 },
{ "layout": "chart_slide",     "section_number": 2 },
{ "layout": "close",           "section_number": null }

// Wrong — gaps and cover with a number
{ "layout": "cover",      "section_number": 1 },   // ❌
{ "layout": "two_column", "section_number": 3 },   // ❌ gap
```

### 4.3 Action Titles

The `action_title` is the single most important quality signal — the punchy declarative takeaway at the top of every slide. It should answer "so what?" not just label the topic.

| Rule | Example |
|---|---|
| Always a declarative sentence | `"Two Sites Drag the Portfolio; Filler Replacement Resolves Both"` |
| Never a noun phrase or topic label | ~~`"Portfolio Overview"`~~ |
| Preferred format: Insight; Consequence | `"OEE Improved 5.4 Points to 84%; Premium SKU Mix Is the Driver"` |
| Numbers beat adjectives | `"Revenue Up 8% to €4.2B"` not `"Strong Revenue Growth"` |
| Max 120 characters (60 for cover) | Full prose sentence is fine |
| Never truncate with `…` | Only use `…` when the source heading itself exceeds 120 chars |
| Semicolons, not dashes | ~~`"Revenue grew — best quarter ever"`~~ |

For cover slides, use `"Title; Department subtitle"` — the semicolon separates deck title from department.

For high-impact slides (Executive Summary, Recommendation, Decision Request) you can italicise a run within the title:

```json
"action_title": [
  ["Three decisions today ", false],
  ["unlock the Q3 trajectory.", true]
]
```

Italic runs render in gold italic. Use sparingly — no more than one in four slides.

---

## 5. Picking the Right Layout

### Structural (no section_number)
| Layout | Use when |
|---|---|
| `cover` | Opening slide |
| `agenda` | Table of contents |
| `section_divider` | Chapter separator |
| `exec_summary` | Executive summary (max 5 key messages) |
| `close` | Final / thank-you |

### Text & Lists
| Layout | Use when |
|---|---|
| `two_column` | Two balanced blocks of content |
| `three_column` / `four_column` | Three or four equal content columns |
| `topic_set` | Several independent topic blocks |
| `vertical_numbered` | Numbered steps or ranked points |
| `label_rows` | Labeled rows (like a vertical table) |
| `pull_quote` | Large headline statement |
| `pros_cons` | Explicit pros vs. cons |

### Process & Flow
| Layout | Use when |
|---|---|
| `phase_process` | Sequential phases/stages (max 5) |
| `arrow_chain` | Linear connected steps with optional conclusion callout |
| `milestone_timeline` | Timeline with dated milestones |
| `gantt` | Project plan with quarters |
| `decision_rows` | Decisions/recommendations with owner (max 5); optional conclusion callout |
| `circular_flow` | Repeating/cyclical process |
| `funnel` | Narrowing stages |
| `waterfall_slide` | Cascading / waterfall steps |
| `road_to_success` | Transformation roadmap: horizontal path + milestones + 2–4 stage columns |
| `fishbone` | Cause-and-effect (Ishikawa) |

### Data & Charts
| Layout | Use when |
|---|---|
| `chart_slide` | Bar, line, column, area chart |
| `donut_chart` | Part-of-whole breakdown |
| `kpi_dashboard` | Multiple KPI metrics |
| `hero_stat` | Single prominent number with context |
| `stat_strip` | Row of 3–4 statistics |
| `status_table` | RAG status table |
| `score_table` | Scoring matrix (dot-filled 1–5 scale, or Harvey Ball 0.0–1.0 with `rating_type: "harvey"`) |
| `comparison_table` | Feature comparison |
| `risk_heatmap` | Likelihood × impact matrix |
| `radar_chart` | Multi-axis comparison |
| `word_cloud` | Frequency-based word visualization |

### Visual & Diagrams
| Layout | Use when |
|---|---|
| `2x2_matrix` | Strategy/priority quadrant |
| `before_after` | Side-by-side transformation |
| `hub_spoke` | Central idea with radiating topics |
| `pillar_detail` | Column pillars with sub-detail |
| `icon_grid` | Grid of icons with descriptions |
| `journey_map` | Customer/process journey |
| `influence_diagram` | Flow of influences |
| `pyramid` | Hierarchical levels |
| `venn` | Overlapping sets |
| `org_chart` | Organizational hierarchy |
| `photo_text` | Large image with text |
| `layered_stack` | Architecture / technology stack |
| `key_question` | Frame a discussion or decision question visually |

---

## 6. Color Do's and Don'ts

### Do
- Use the six official color themes — they ensure all colors comply with Merck brand
- Use traffic lights exactly as defined: Good = `#149B5F`, Warning = `#FFC832`, Bad = `#E61E50`
- Mix colors from **different rows AND different columns** of the brand matrix when combining
- Use groups of **3 colors** — fewer is monotonous, more is cluttered

### Don't
- Never invent custom colors or approximate brand colors
- Never use gradients — not linear, not radial
- Never use tints/shades in design elements (only in charts and tables is it acceptable)
- Never combine red + green (colorblindness issue)
- Never make the title text color match the logo color on the same slide

---

## 7. Content Schemas (Most Common)

### `cover`
```jsonc
{
  "layout": "cover",
  "action_title": "Department Name; Subtitle",
  "section_number": null,
  "content": {
    "subtitle": "Optional descriptive subtitle",
    "authors": ["Name, Title"],
    "key_messages": ["Key message 1", "Key message 2"]
  }
}
```

### `exec_summary`
```jsonc
{
  "layout": "exec_summary",
  "section_number": null,
  "content": {
    "key_messages": [
      { "label": "Headline", "body": "Supporting detail" },
      { "label": "Headline 2", "body": "Supporting detail 2" }
    ]
  }
}
// Max 5 key messages
```

### `two_column`
```jsonc
{
  "layout": "two_column",
  "content": {
    "left":  { "label": "Left heading", "body": "Key claim or hero sentence (≤300 chars)", "tone": "neutral",  "items": ["Sub-bullet expanding on body"] },
    "right": { "label": "Right heading","body": "Key claim or hero sentence (≤300 chars)", "tone": "positive", "items": ["Sub-bullet expanding on body"] }
  }
}
// Use for EXACTLY 2 items. NEVER use two_column for 4 items — use 2x2_matrix instead.
// tone: "positive" | "negative" | "neutral"
// NEVER put the same content in both body AND items — both render visibly on the slide
```

### `chart_slide`
```jsonc
{
  "layout": "chart_slide",
  "content": {
    "chart": {
      "type": "bar",
      "data": {
        "categories": ["Q1", "Q2", "Q3"],
        "series": [
          { "name": "Revenue", "values": [10, 15, 12] }
        ]
      }
    },
    "callouts": ["Revenue peaked in Q2"]
  }
}
```

### `decision_rows`
```jsonc
{
  "layout": "decision_rows",
  "content": {
    "decisions": [
      {"number": 1, "title": "Approve budget increase",
       "desc": "€12M incremental for FY2026 — budget sign-off required by 30 June",
       "owner": "CFO", "tone": "positive"},
      {"number": 2, "title": "Defer EU launch",
       "desc": "Delay to Q2 2027 pending trial readout data",
       "owner": "CMO", "tone": "neutral"}
    ],
    "callout": {"type": "conclusion", "text": "All three approvals needed before Q3 planning lock"}
  }
}
// CANONICAL keys: title + desc (desc ≤160 chars). Never use "text" as the only body key.
// tone: "positive" | "negative" | "neutral" — Max 5 decisions
// callout is optional: type = "conclusion" | "result" | "next" | "future"
```

### `phase_process`
```jsonc
{
  "layout": "phase_process",
  "content": {
    "show_arrows": true,
    "phases": [
      {"label": "01", "title": "Discovery",  "body": "Research phase",       "status": "done"},
      {"label": "02", "title": "Design",     "body": "Co-creation workshops","status": "current"},
      {"label": "03", "title": "Deploy",     "body": "Pilot in 3 markets",   "status": "future"}
    ]
  }
}
// status: "done" | "current" | "future"  (use status, not the deprecated highlighted: bool)
// Max 5 phases; body ≤100 chars per phase
```

### `hero_stat`
```jsonc
{
  "layout": "hero_stat",
  "section_number": null,
  "content": {
    "stat": { "value": "€4.2B", "label": "FY2025 Revenue" },
    "context": "Up 18% year-on-year, driven by Oncology"
  }
}
```

### `gantt`
```jsonc
{
  "layout": "gantt",
  "content": {
    "quarters": ["Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025"],
    "rows": [
      { "label": "Phase 1: Planning" },
      { "label": "Phase 2: Execution" }
    ]
  }
}
```

### `road_to_success`
```jsonc
{
  "layout": "road_to_success",
  "content": {
    "stages": [
      {"title": "Discover", "body": "Map current state and identify pain points"},
      {"title": "Build",    "body": "Prototype and validate with stakeholders"},
      {"title": "Scale",    "body": "Full rollout — Q4 target"}
    ],
    "milestones": ["Kick-off", "Sprint 1", "Pilot", "Launch"]
  }
}
// stages: 2–4; milestones: 2–6 (omit for a clean path without labels)
```

### `key_question`
```jsonc
{
  "layout": "key_question",
  "section_number": null,
  "content": {
    "question": "Should we scale now or strengthen the foundation first?",
    "context":  "Decision required before Q3 planning cycle"
  }
}
// Use to visually frame a decision before decision_rows or before_after slides.
```

### `score_table` with Harvey Ball
```jsonc
{
  "layout": "score_table",
  "content": {
    "rating_type": "harvey",
    "rows": [
      {"label": "Customer Satisfaction", "score": 0.75, "note": "Above target"},
      {"label": "Cost Efficiency",        "score": 0.50, "note": "On track"},
      {"label": "Innovation Pipeline",    "score": 0.25, "note": "Needs focus"}
    ]
  }
}
// score: 0.0 (empty circle) to 1.0 (full circle) — fraction, not integer
```

---

## 8. Complete Meta Block Reference

```jsonc
{
  "meta": {
    "region":          "EU",               // Required: "EU" | "USA"
    "deck_label":      "Project Alpha",    // Required: appears in footer
    "classification":  "Internal",         // Required: "Public" | "Internal" | "Confidential"
    "month_year":      "June 2026",        // Required: displayed on cover
    "audience":        "Leadership Team",  // Required: informs LLM tone
    "deck_style":      "merck_corporate",  // Required: "merck_executive" | "merck_corporate" | "merck_storytelling" | "merck_science"
    "color_theme":     "plastic",          // Required: "plastic" | "functional" | "organic" | "synthetic" | "technical" | "electronics"
    "variety_mode":    "default",          // Optional: "default" | "creative"
    "show_disclaimer": false,              // Optional: show legal disclaimer text
    "chrome": {                            // Optional: opt-in custom chrome elements
      "progress_bar":         false,       //   Top proportional fill strip
      "section_circles":      false,       //   Numbered circles + category tag
      "takeaway_bands":       false,       //   Purple takeaway band (LLM writes takeaway only if true)
      "footer_breadcrumb":    false,       //   "Deck • Category" left footer
      "classification_badge": false        //   "Classification: INTERNAL" badge top-right (non-standard)
    }
  }
}
```

> **Default = standard empower.** All `chrome` flags default to `false`. Only the Merck logo, classification badge, and page number are rendered unless you explicitly enable a flag.

---

## 9. Typography Summary

The pipeline applies all fonts automatically. For reference:

| Text element | Font | Size | Color |
|---|---|---|---|
| Action title (top takeaway) | Verdana Regular | 22 pt | Rich Purple |
| Slide heading | Verdana Bold | 22 pt | Rich Purple |
| Body text | Verdana Regular | 16 pt | Near-black |
| Labels in boxes | Verdana Regular | 12 pt | Rich Purple |
| Source / reference | Verdana Regular | 8 pt | Medium gray |
| Footer | Verdana Regular | 8 pt | Rich Purple |

**Rules you should follow in your content text:**
- Maximum 3 different font sizes in any slide's content area
- Bullets: no bullet at level 1, filled bullet at level 2, en-dash at level 3
- Left-align continuous text and lists
- Center short labels inside boxes or graphic elements
- Right-align numbers in tables

---

## 10. Common Mistakes to Avoid

| Mistake | What happens | Fix |
|---|---|---|
| Wrong region | Legal disclaimer mismatch | Always set `"EU"` or `"USA"` explicitly |
| section_number on cover/agenda/divider/close | Renders incorrectly | Set `"section_number": null` |
| Gaps in section_number | Out-of-sequence numbering visible | Use 1, 2, 3 ... with no gaps |
| Noun-phrase action_title | Loses the "so what" | Write as a declarative sentence |
| action_title over 120 chars (60 for cover) | Truncated silently | Keep within the limit; prose is fine |
| Using `two_column` for 4 items | Items 3 and 4 silently dropped | Use `2x2_matrix` for 4-item grids |
| Same content in both `body` and `items` | Both render — content duplicated on slide | Use one or the other, not both |
| `icon_grid` with emoji string | Renders as small text symbol, not vector icon | Use named icon string (e.g. `"target"`, `"users"`) |
| More than 5 items in exec_summary | Truncated silently | Keep to ≤ 5 key messages |
| More than 5 decisions in decision_rows | Truncated silently | Keep to ≤ 5 rows |
| More than 5 phases in phase_process | Truncated silently | Keep to ≤ 5 phases |
| `decision_rows` with `"text"` as only body key | Body renders empty | Use `"title"` + `"desc"` (canonical) |
| `phase_process` with `"highlighted": true` | Deprecated — ignored | Use `"status": "current"` |
| `score_table` Harvey Ball with integer score | Score ignores scale — renders 0% | Use float 0.0–1.0 with `rating_type: "harvey"` |
| Using `color_theme` for content slides | Theme is cover/divider only | Theme does not change body slide colors |
| Embedding Excel tables in charts | Bloats file, exposes hidden data | Use `chart_slide` layout with JSON data |

---

## 11. Image Guidelines

When using `photo_text` or any image-bearing layout:
- Images should be **colorful, focused, authentic** — avoid black-and-white, collages, or ClipArt
- After generating the `.pptx`, resize images using **corner drag points only** — never edge points (distorts)
- For the `electronics` theme cover: the image placeholder is intentionally empty — add your photo in PowerPoint

---

## 12. Pre-Send Checklist

Before sharing a generated deck, verify:

- [ ] `meta.region` matches your actual audience (EU vs. USA)
- [ ] `color_theme` matches the visual identity you intend
- [ ] Every content slide has a unique, sequential `section_number`
- [ ] Cover, agenda, section_divider, and close have `section_number: null`
- [ ] Action titles are declarative sentences, not noun phrases
- [ ] All key_messages / decisions / phases are within their limits (max 5)
- [ ] Chart data uses the correct `categories` / `series` structure
- [ ] No custom colors used outside the six official themes

---

## 13. Getting Help

| Question | Where to look |
|---|---|
| Full design rules (colors, fonts, shapes) | `Merck_Presentation_Guidelines.md` |
| How to use the tool, settings guide, layout overview | `README.md` |
| Complete JSON schema for all 46 layouts | `LLM_PLAN_GUIDE.md` |
| Branding questions | branding@merckgroup.com |
| Brand Hub | https://brandhub.merckgroup.com |
