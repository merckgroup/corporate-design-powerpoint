# Merck Presentation Builder — User Guide

> A practical helper for creating Merck Corporate Design-compliant presentations with the pipeline.  
> For full technical rules see `Merck_Presentation_Guidelines.md`. For API/CLI reference see `README.md`.

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

## 1. Choosing Your Region

Set `meta.region` to match your audience — this controls the legal disclaimer and template file.

| Region | Value | Template | Legal footer |
|---|---|---|---|
| All countries except USA/Canada | `"EU"` | EU_Merck_Themed.pptx | Merck KGaA disclaimer |
| USA and Canada | `"USA"` | USA_Merck_Themed_Base_v1.pptx | EMD/MilliporeSigma disclaimer |

**Rule:** Never use an EU template for a USA audience and vice versa.

---

## 2. Choosing a Visual Style

Set `meta.deck_style`. Every slide in the deck inherits this unless overridden per-slide.

| Style | When to use | Background | Primary color |
|---|---|---|---|
| `merck_executive` | Board decks, C-suite, formal decisions | White | Rich Purple |
| `merck_corporate` | General business presentations | White | Rich Purple |
| `merck_storytelling` | Impactful narrative, bold statements | Purple | White (inverted) |

**Auto-promotion:** Any slide with category "Executive Summary", "Recommendation", "Decision Request", "Risk", or "Tradeoff" automatically gets `merck_executive` style — even if the deck default is different.

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

The `action_title` field is the punchy one-line takeaway at the top of each slide.

**Good:** `"Revenue grew 22% YoY; digital channels drove the gain"`  
**Bad:** `"Revenue Overview"` (noun phrase — tells nothing)

- Use declarative sentences, not noun phrases
- Prefer numbers over vague adjectives: "22% growth" not "significant growth"  
- Max ~80 characters — it must fit on one line
- Use `"Title; Subtitle"` for cover slides (semicolon separates title from department)

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
| `arrow_chain` | Linear connected steps |
| `milestone_timeline` | Timeline with dated milestones |
| `gantt` | Project plan with quarters |
| `decision_rows` | Decisions/recommendations with owner (max 5) |
| `circular_flow` | Repeating/cyclical process |
| `funnel` | Narrowing stages |
| `waterfall_slide` | Cascading / waterfall steps |
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
| `score_table` | Scoring matrix |
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
    "left":  { "header": "Left heading", "items": ["Point 1", "Point 2"] },
    "right": { "header": "Right heading", "items": ["Point 1", "Point 2"] }
  }
}
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
      { "tone": "positive", "number": 1, "owner": "CMO", "text": "Approve budget increase" }
    ]
  }
}
// tone: "positive" | "negative" | "neutral" — Max 5 decisions
```

### `phase_process`
```jsonc
{
  "layout": "phase_process",
  "content": {
    "show_arrows": true,
    "phases": [
      { "label": "01", "title": "Discovery", "body": "Research phase", "highlighted": false }
    ]
  }
}
// Max 5 phases
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
    "deck_style":      "merck_corporate",  // Required: "merck_executive" | "merck_corporate" | "merck_storytelling"
    "color_theme":     "plastic",          // Required: "plastic" | "functional" | "organic" | "synthetic" | "technical" | "electronics"
    "variety_mode":    "default",          // Optional: "default" | "creative"
    "show_disclaimer": false               // Optional: show legal disclaimer text
  }
}
```

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
| More than 5 items in exec_summary | Truncated silently | Keep to ≤ 5 key messages |
| More than 5 decisions in decision_rows | Truncated silently | Keep to ≤ 5 rows |
| More than 5 phases in phase_process | Truncated silently | Keep to ≤ 5 phases |
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
| Full design rules (colors, fonts, shapes) | `Guidelines/Merck_Presentation_Guidelines.md` |
| CLI commands and API reference | `README.md` |
| Plan JSON schema details | `CLAUDE.md` (Plan JSON schema section) |
| Branding questions | branding@merckgroup.com |
| Brand Hub | https://brandhub.merckgroup.com |
