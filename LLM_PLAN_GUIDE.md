# Merck PPTX — LLM Plan Generation Guide

> **Audience:** This file is written for LLMs. It is the authoritative reference for producing
> a valid `plan.json` that the `merck_pptx` pipeline will render into a polished, brand-compliant
> Merck Healthcare KGaA PowerPoint deck.
>
> **Ground truth:** `merck_pptx/build_from_plan.py` (dispatcher) and `merck_pptx/merck_layouts.py`
> (renderer). When in doubt, the dispatcher is authoritative.

---

## 1. Plan JSON top-level shape

```json
{
  "meta":      { ... },
  "storyline": ["chapter 1 sentence", "chapter 2 sentence"],
  "slides":    [ { ... }, { ... } ]
}
```

`storyline` is optional. Include 3–5 chapter-level action sentences when the deck
has multiple sections; omit for single-topic decks.

---

## 2. Meta block — all fields

```jsonc
{
  "meta": {
    "region":          "EU",              // REQUIRED. "EU" | "USA"
    "deck_label":      "Project Alpha",   // REQUIRED. Appears in slide footer.
    "classification":  "Internal",        // REQUIRED. "Public" | "Internal" | "Confidential"
    "month_year":      "June 2026",       // REQUIRED. Displayed on cover (e.g. "June 2026").
    "audience":        "Leadership Team", // REQUIRED. Free text; informs LLM tone.
    "deck_style":      "merck_corporate", // REQUIRED. One of the three styles (see §5).
    "color_theme":     "plastic",         // REQUIRED. One of the six themes (see §6).
    "variety_mode":    "default",         // Optional. "default" | "creative"
    "show_disclaimer": false,             // Optional. Show legal disclaimer footer text.
    "division":        "merck",           // Optional. Controls template file (see §7).
    "cover_top_bar":   false,             // Optional. Show a colored bar at top of cover.
    "chrome": {                           // Optional. Opt-in custom chrome (all default false).
      "progress_bar":         false,      //   Proportional fill strip at top of content slides.
      "section_circles":      false,      //   Numbered purple circles + category tag (top-left).
      "takeaway_bands":       false,      //   Purple takeaway band above footer. LLM writes
                                          //   takeaway text ONLY when this is true.
      "footer_breadcrumb":    false,      //   "Deck Label • Category" left footer text.
      "classification_badge": false       //   "Classification: INTERNAL" badge top-right.
                                          //   Not present in standard empower layouts.
    }
  }
}
```

> **Default is standard empower.** Without `chrome` (or with all flags `false`), the output contains only the Merck logo, classification badge, and page number — matching the standard empower template. Set flags to `true` individually to add each element.


### `region`
| Value | Legal footer | Template |
|---|---|---|
| `"EU"` | Merck KGaA disclaimer | EU_Merck_Themed.pptx |
| `"USA"` | EMD / MilliporeSigma disclaimer | USA_Merck_Themed_Base_v1.pptx |

**NEVER use EU template for USA audiences and vice versa.**

### `division` (optional, default: `"merck"`)
Selects the divisional template. Omit unless the deck targets a specific Merck division.
| Value | Division |
|---|---|
| `"merck"` | Merck Healthcare (default) |
| `"emd_serono"` | EMD Serono |
| `"emd_electronics"` | EMD Electronics |
| `"millipore_sigma"` | MilliporeSigma |
| `"merck_asia"` | Merck Asia |

---

## 3. Slide anatomy — all slide-level fields

```jsonc
{
  "page":           1,                    // 1-indexed. Auto-renumbered by pipeline.
  "page_function":  "Evidence",           // See §4 for valid values.
  "layout":         "two_column",         // One of the 44 layout keys (see §9).
  "action_title":   "Sales grew 18% YoY; digital led the gain",  // ≤120 chars (60 for cover)
  "section_number": 3,                    // integer or null (see §8)
  "style":          "inherit",            // "inherit" | "merck_executive" | "merck_corporate" | "merck_storytelling"
  "category":       "EVIDENCE",           // UPPERCASE tag (free text). Auto-promote trigger (see §5).
  "takeaway":       "Act now to capture share before Q3.",        // ≤120 chars
  "source":         "Internal sales data, FY2025",
  "notes":          "Mention the Asia inflection in speaker notes.",
  "subtitle":       null,                 // Optional secondary heading shown below action_title. Always at slide level, never inside content.
  "appendix":       false,                // Set true to move slide to appendix section.
  "content":        { ... }               // Layout-specific payload (see §9).
}
```

### `page_function` valid values
```
Cover | Agenda | Section Divider | Executive Summary | Close |
Framing | Diagnosis | Evidence | Recommendation | Decision Request |
Risk | Tradeoff
```

### `style` resolution order
1. If `deck_style` is `"merck_storytelling"` → all slides become `merck_storytelling` (overrides everything).
2. If `style` is `"inherit"` → use `deck_style`; if no deck_style, fall back to `"merck_executive"`.
3. Per-slide `style` overrides `"inherit"`.
4. Auto-promote overrides step 2–3 (see §5).

---

## 4. Action titles and takeaways

### `action_title` (≤120 chars, ≤60 for cover)
- Must be a **declarative sentence**, never a noun phrase.
- Preferred pattern: `"INSIGHT; CONSEQUENCE"` (semicolon-separated).
- Use numbers over adjectives: `"OEE improved 5.4 pts to 84%"` not `"OEE improved significantly"`.
- Never truncate with `…` unless the source heading itself exceeds 120 chars.
- Cover slides: use `"Deck Title; Department subtitle"` (semicolon separates title from department).
- Cover action_title is truncated to 60 chars — keep it shorter.

**Good:** `"Revenue grew 22% YoY; digital channels drove the gain"`
**Bad:** `"Revenue Overview"` (noun phrase — tells nothing)

### `takeaway` (≤120 chars)
- The one-sentence "so what" of the slide.
- Omit (set `null`) on: `cover`, `agenda`, `section_divider`, `close`, `exec_summary`, `hero_stat`, `pull_quote`.
- Lives at slide level, not in content.

---

## 5. Visual styles and auto-promote

| Style | When to use | Background | Text |
|---|---|---|---|
| `merck_executive` | Board, C-suite, formal decisions | White | Rich Purple |
| `merck_corporate` | General business presentations | White | Rich Purple |
| `merck_storytelling` | Impactful narrative, bold statements | Purple | White (inverted) |

### Auto-promote rule
If `category` (or `page_function`) exactly matches one of these strings, the slide is
automatically promoted to `merck_executive` — even if the deck default differs:
```
"Executive Summary" | "Recommendation" | "Decision Request" | "Risk" | "Tradeoff"
```
For these slides, also set `style: "merck_executive"` explicitly (belt-and-suspenders).

---

## 6. Color themes

Set in `meta.color_theme`. Controls cover/divider background and brand accent color.

| Theme | Cover bg | Accent color | Best for |
|---|---|---|---|
| `plastic` | Light green #A5CD50 | Pink #EB3C96 | Default; general purpose |
| `functional` | Light green #A5CD50 | Teal #2DBECD | Life science, biology |
| `organic` | Cream #FFDCB9 | Red #E61E50 | Healthcare, patient focus |
| `synthetic` | Dark violet #503291 | Yellow #FFC832 | Industrial, chemistry |
| `technical` | Cream #FFDCB9 | Teal #2DBECD | Engineering, IT |
| `electronics` | Dark violet #503291 | Yellow #FFC832 | EMD Electronics; cover has editable photo |

Dark themes (`synthetic`, `electronics`): white text is used automatically on dark backgrounds.

---

## 7. Section numbers

```
Structural slides → section_number: null
Content slides    → section_number: 1, 2, 3 ... (unique sequential integers, no gaps)
```

Structural layouts that always get `null`:
`cover`, `agenda`, `section_divider`, `close`, `exec_summary`, `hero_stat`, `pull_quote`

All other layouts get a unique sequential integer starting at 1.
Never repeat a number. Never leave gaps.

```jsonc
// Correct
{ "layout": "cover",           "section_number": null },
{ "layout": "agenda",          "section_number": null },
{ "layout": "section_divider", "section_number": null },
{ "layout": "two_column",      "section_number": 1 },
{ "layout": "chart_slide",     "section_number": 2 },
{ "layout": "section_divider", "section_number": null },
{ "layout": "phase_process",   "section_number": 3 },
{ "layout": "close",           "section_number": null }
```

---

## 8. Deck structure rules

- First slide: always `cover`.
- Last slide: always `close`.
- Include an `agenda` slide as slide 2 if the deck has ≥ 5 content slides.
- Use `section_divider` before each chapter group.
- Agenda `chapters` can be left empty (`[]`) — the pipeline auto-fills from content slides.
- Maximum slides: 150 (hard pipeline limit).

### Slide count heuristics

| Deck type | Total slides | Content slides |
|---|---|---|
| Executive briefing (20 min) | 8–12 | 4–7 |
| Standard business review (45 min) | 15–20 | 8–14 |
| Deep-dive workshop (60 min) | 20–30 | 12–22 |

Source content of ~1 page (500 words) maps to roughly 2–4 slides.
Source content of ~5 pages maps to roughly 8–15 slides.

Move supporting detail (methodology, data tables, backup analysis) to appendix slides — keep the main deck narrative clean.

### Deck pacing and variety rules

- **Vary layout families.** Do not use more than 2 consecutive slides from the same family (e.g. two `two_column` slides in a row is fine; three is monotonous).
- **Visual breaks.** After every 3–4 dense content slides, insert a visual break: `hero_stat`, `pull_quote`, or `stat_strip`.
- **Content density per chapter.** 2–4 content slides per chapter is ideal. More than 6 suggests the chapter should be split into two sections.
- **Vary text vs. visual.** Each chapter should contain at least one chart/process layout alongside any column or list layouts.

Layout families for pacing purposes:
- **Text/list:** `two_column`, `three_column`, `four_column`, `vertical_numbered`, `label_rows`, `topic_set`
- **Process/timeline:** `phase_process`, `arrow_chain`, `circular_flow`, `gantt`, `milestone_timeline`, `funnel`, `journey_map`, `road_to_success`
- **Chart/data:** `chart_slide`, `waterfall_slide`, `donut_chart`, `radar_chart`, `kpi_dashboard`, `stat_strip`, `hero_stat`
- **Decision/analysis:** `decision_rows`, `before_after`, `2x2_matrix`, `comparison_table`, `pros_cons`, `risk_heatmap`
- **Visual/story:** `pull_quote`, `word_cloud`, `pyramid`, `venn`, `icon_grid`, `fishbone`, `key_question`

---

## 9. All 46 layouts — content schemas

Every `content` field is an object. All keys shown are read from `content` unless noted as
a slide-level field (e.g., `action_title`, `takeaway`, `source` live at the slide root, not in content).

### IMPORTANT: wrong key names produce silently empty slides. Use exact key names below.

---

### Structural layouts

#### `cover`
```jsonc
{
  "layout": "cover",
  "action_title": "Deck Title; Department subtitle",  // slide-level, ≤60 chars for cover
  "subtitle": "Optional descriptive subtitle",         // slide-level ONLY — not inside content
  "section_number": null,
  "content": {
    "authors":      [{"name": "Dr. Ada Lovelace", "title": "VP Strategy"}],
    "key_messages": ["Key message 1", "Key message 2"]  // optional, shown below authors
  }
}
```
> `subtitle` must be at the **slide level**, not inside `content`.
> The semicolon in `action_title` also works: `"Deck Title; Subtitle"`.

#### `exec_summary`
```jsonc
{
  "layout": "exec_summary",
  "section_number": null,
  "content": {
    "key_messages": [
      {"label": "Headline", "body": "Supporting detail ≤120 chars"},
      {"label": "Headline 2", "body": "Supporting detail 2"}
    ]
  }
}
// Max 5 key_messages
```

#### `agenda`
```jsonc
{
  "layout": "agenda",
  "section_number": null,
  "content": {
    "chapters": [
      {"number": "01", "title": "Context & Diagnosis", "subtitle": "Optional subline"}
    ]
  }
}
// Leave chapters: [] to auto-fill from content slides' section_number + action_title
// Max 12 chapters
```

#### `section_divider`
```jsonc
{
  "layout": "section_divider",
  "action_title": "Chapter Title",  // slide-level — the chapter heading ONLY, ≤40 chars
                                    // Good: "Results"  Bad: "Results: Candidate Pairs Found"
                                    // Do NOT include slide topic, a colon prefix, or sub-heading
  "section_number": null,            // always null — structural slide
  "content": {
    "number": "01"                   // the large chapter number displayed on the divider
                                     // omit or leave "" for no chapter number
  }
}
```
> `content.number` drives the large decorative number on the divider.
> `section_number` stays null (section_divider is a structural slide).
> `action_title` must be the chapter heading only — ≤40 chars, no colon-separated subtitle.

#### `close`
```jsonc
{
  "layout": "close",
  "section_number": null,
  "content": {
    "action_statement": "Thank you — next steps: schedule discovery call by July 1"
  }
}
// If action_statement is absent, the slide falls back to action_title.
```

---

### Text & list layouts

#### `two_column`
```jsonc
{
  "layout": "two_column",
  "content": {
    "left":  {"label": "Current state", "body": "Paragraph text ≤300 chars", "tone": "neutral", "items": ["Point 1", "Point 2"]},
    "right": {"label": "Future state",  "body": "Paragraph text ≤300 chars", "tone": "positive", "items": ["Point 1"]}
  }
}
// Use for EXACTLY 2 items side by side. NEVER for 4 items — use 2x2_matrix instead.
// tone: "positive" | "negative" | "neutral" — colors the label strip
// items and body are both optional; use one or both (NEVER put same content in both)
```

**Column card field rules (apply to two_column, three_column, four_column, columns):**
- `label` = column header shown in the bar (e.g. "STEP 1", "THE PROBLEM")
- `body` = bold hero sentence — the key claim or step subtitle (1–2 sentences, ≤300 chars)
- `items` = sub-bullet details EXPANDING on body (≤120 chars each); omit when body is self-contained
- **NEVER put the same content in both `body` AND `items`** — both render visibly on the slide
- Cross-cutting bullets (apply to all columns, not one): add a final `"Key Points"` column with `items[]` rather than using `takeaway` (takeaway is a single short line, not a multi-bullet list)

#### `three_column`
```jsonc
{
  "layout": "three_column",
  "content": {
    "columns": [
      {"label": "Discover",  "body": "Description", "tone": "neutral", "items": ["Point 1"]},
      {"label": "Design",    "body": "Description", "tone": "positive"},
      {"label": "Deliver",   "body": "Description", "tone": "positive"}
    ]
  }
}
// Exactly 3 items in columns[]
// Optional: "framework": "4ps" | "value_disciplines" | "balanced_scorecard" → pre-populates labels
```

#### `four_column`
```jsonc
{
  "layout": "four_column",
  "content": {
    "columns": [
      {"label": "Q1", "body": "Description", "tone": "neutral", "highlighted": false, "items": ["Point"]},
      {"label": "Q2", "body": "Description", "tone": "positive", "highlighted": true},
      {"label": "Q3", "body": "Description", "tone": "neutral"},
      {"label": "Q4", "body": "Description", "tone": "positive"}
    ]
  }
}
// Exactly 4 items. highlighted: true adds visual emphasis to that column.
```

#### `columns` (auto-dispatcher)
Use `"columns"` as the layout key when you don't want to count yourself — the pipeline
dispatches to two/three/four_column based on `columns` array length.
```jsonc
{
  "layout": "columns",
  "content": {
    "columns": [
      {"label": "Option A", "body": "...", "tone": "positive"},
      {"label": "Option B", "body": "...", "tone": "neutral"},
      {"label": "Option C", "body": "...", "tone": "negative"}
    ]
  }
}
// 2 items → two_column, 3 → three_column, 4+ → four_column
```

#### `vertical_numbered`
```jsonc
{
  "layout": "vertical_numbered",
  "content": {
    "items": [
      {"title": "Step 1 heading", "body": "Detail text"},
      {"title": "Step 2 heading", "body": "Detail text"}
    ]
  }
}
// Optional: "framework": "kotter_8" | "adkar" → pre-populates item titles
```

#### `label_rows`
```jsonc
{
  "layout": "label_rows",
  "content": {
    "rows": [
      {"label": "Objective",   "body": "Grow market share to 35% by FY2027", "color": "teal"},
      {"label": "Constraint",  "body": "Budget capped at €12M"}
    ],
    "label_color": null,    // Optional hex or theme color for ALL label strips (global)
    "callout": {"type": "next", "text": "Three decisions are required to proceed"}
  }
}
// body ≤300 chars per row (full prose description is allowed)
// per-row color (optional): "gray" | "teal" | "blue" | "green" | "yellow" | "orange" | "red" | "pink" | "purple"
// callout is optional — same type/text schema as decision_rows callout
```

#### `topic_set`
```jsonc
{
  "layout": "topic_set",
  "content": {
    "topics": [
      {"label": "01", "title": "Digital Health",   "body": "Description ≤120 chars", "icon": "lightbulb"},
      {"label": "02", "title": "Data Integration", "body": "Description ≤120 chars", "icon": "chart_bar"},
      {"label": "03", "title": "Access Programs",  "body": "Description ≤120 chars", "icon": "users"}
    ]
  }
}
// max 6 topics (excess silently dropped)
// icon is optional; if absent, sequential numbers (1, 2, 3…) are shown instead
// body ≤120 chars
// Valid icon names (26 registered): chart_bar, chart_line, chart_pie,
//   arrow_up, arrow_down, arrow_right, check, x, alert, info,
//   target, gear, users, calendar, clock, lightbulb, lock, globe,
//   search, money, trending_up, trending_down, shield, flag, doc
// Unknown icon names fall back to a generic target icon (no error)
```

**Icon selection guide for `topic_set` and `layered_stack`:**

| Concept | Icon name |
|---|---|
| Strategy, goal, objective | `target` |
| Growth, improvement, upward trend | `trending_up` |
| Decline, reduction, downward trend | `trending_down` |
| Finance, revenue, budget | `money` |
| People, team, organization | `users` |
| Calendar, scheduling, planning | `calendar` |
| Deadline, speed, time pressure | `clock` |
| Innovation, idea, insight | `lightbulb` |
| Process, configuration, operations | `gear` |
| Security, compliance, governance | `shield` |
| Global, geography, markets | `globe` |
| Research, discovery, analysis | `search` |
| Document, report, communication | `doc` |
| Milestone, launch, flag event | `flag` |
| Access, permissions, restriction | `lock` |
| Warning, alert, issue | `alert` |
| Information, context, note | `info` |
| Confirmed, approved, complete | `check` |
| Rejected, closed, stopped | `x` |
| Comparison chart, KPI bar | `chart_bar` |
| Trend line, historical data | `chart_line` |
| Breakdown, share, composition | `chart_pie` |
| Upward direction, increase | `arrow_up` |
| Downward direction, decrease | `arrow_down` |
| Next step, forward, progression | `arrow_right` |

#### `pull_quote`
```jsonc
{
  "layout": "pull_quote",
  "section_number": null,
  "content": {
    "quote":       "The market does not reward complexity.",
    "attribution": "CEO, Merck Healthcare",
    "context":     "Investor Day 2025"
  }
}
```

#### `pros_cons`
```jsonc
{
  "layout": "pros_cons",
  "content": {
    "subject":    "Acquisition of Alpha AG",
    "pros":       ["Fast market entry", "Strong IP portfolio", "Synergy in oncology"],
    "cons":       ["Integration risk", "Premium valuation", "Regulatory overlap"],
    "pros_label": "ADVANTAGES",   // optional, default "ADVANTAGES"
    "cons_label": "RISKS"         // optional, default "RISKS"
  }
}
```

---

### Data & chart layouts

#### `chart_slide`
```jsonc
{
  "layout": "chart_slide",
  "content": {
    "chart": {
      "type": "column",
      "data": {
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "series": [
          {"name": "Revenue", "values": [10.2, 11.5, 13.0, 14.8]},
          {"name": "Target",  "values": [10.0, 12.0, 13.5, 15.0]}
        ]
      }
    },
    "callouts": [
      {"x_in": 5.2, "y_in": 2.1, "label": "Q3 inflection", "direction": "left"}
    ]
  }
}
// chart.type: "bar" | "line" | "column" | "area" | "slope" | "dot" | "marimekko"
// callouts are optional; x_in/y_in are slide inches from top-left
// Default chart type if omitted: "column"
```

**Chart type selection guide:**

| Type | When to use | Avoid when |
|---|---|---|
| `column` | Comparing discrete categories at one point in time | Labels are long (use `bar` instead) |
| `bar` | Same as column but category labels are long or there are many (6+) categories | Showing time trends |
| `line` | Trend over multiple time periods (4+ data points) | Comparing unrelated categories |
| `area` | Trend with volume/mass emphasis; stacked proportions over time | Single series with few points |
| `slope` | Before/after comparison for a small set (2–8) of named items across exactly 2 points | More than 2 time points |
| `dot` | Distribution of a single value across many items; clinical endpoint comparison | Showing magnitude or total |
| `marimekko` | Two-dimensional market sizing where width = market size and height = share | Non-proportional data |

Use `waterfall_slide` (not `chart_slide`) for revenue bridges, cost waterfalls, or any step chart.

**Slope chart** — requires a different data structure than bar/line/column:
```jsonc
{
  "layout": "chart_slide",
  "content": {
    "chart": {
      "type": "slope",
      "data": {
        "labels": ["FY2024", "FY2025"],
        "items": [
          ["Product A", 45, 62],
          ["Product B", 72, 68],
          ["Product C", 38, 55]
        ],
        "highlight_indices": [0, 2]
      }
    }
  }
}
// data.items: list of [label_str, before_value, after_value]
// data.labels: [before_col_label, after_col_label]
// data.highlight_indices: list of 0-based indices to highlight in accent color
// Do NOT use categories/series format for slope charts
```

#### `waterfall_slide`
```jsonc
{
  "layout": "waterfall_slide",
  "content": {
    "chart": {
      "type": "waterfall",
      "data": {
        "bars": [
          {"label": "Base",           "value": 100, "type": "total"},
          {"label": "New customers",  "value": 18,  "type": "positive"},
          {"label": "Churn",          "value": 7,   "type": "negative"},
          {"label": "FY2025 Revenue", "value": 111, "type": "total"}
        ]
      }
    }
  }
}
// bar type: "positive" | "negative" | "total"
// "total" = anchored bar (start/end); "positive" = up bar; "negative" = down bar (value must be positive)
```

#### `hero_stat`
```jsonc
{
  "layout": "hero_stat",
  "section_number": null,
  "content": {
    "stat":    {"value": "€4.2B", "label": "FY2025 Revenue"},
    "context": "Up 18% year-on-year, driven by Oncology ≤130 chars"
  }
}
```

#### `stat_strip`
```jsonc
{
  "layout": "stat_strip",
  "content": {
    "stats": [
      {"value": "84%",   "label": "OEE",       "body": "↑ 5.4 pts YoY"},
      {"value": "€2.1B", "label": "Revenue",   "body": "In-line with plan"},
      {"value": "12",    "label": "Countries", "body": "Active markets"}
    ]
  }
}
// 3–4 stats; body is optional context text
```

#### `donut_chart`
```jsonc
{
  "layout": "donut_chart",
  "content": {
    "segments": [
      {"label": "Oncology",        "value": 42},
      {"label": "Neurology",       "value": 28},
      {"label": "Cardiovascular",  "value": 30}
    ],
    "center_value": "€4.2B",
    "center_label": "Total Revenue",
    "legend_title": "Business Units"
  }
}
// values are relative (do not need to sum to 100)
```

#### `radar_chart`
```jsonc
{
  "layout": "radar_chart",
  "content": {
    "axes":   ["Quality", "Speed", "Cost", "Innovation", "Compliance"],
    "series": [
      {"label": "Current",  "values": [80, 60, 70, 55, 90]},
      {"label": "Target",   "values": [90, 80, 75, 85, 95]}
    ]
  }
}
// values: 0–100 per axis
```

#### `risk_heatmap`
```jsonc
{
  "layout": "risk_heatmap",
  "content": {
    "risks": [
      {"label": "Supply disruption", "likelihood": 3, "impact": 4},
      {"label": "Regulatory delay",  "likelihood": 2, "impact": 5},
      {"label": "FX exposure",       "likelihood": 4, "impact": 2}
    ],
    "x_label": "LIKELIHOOD",  // optional, default "LIKELIHOOD"
    "y_label": "IMPACT"        // optional, default "IMPACT"
  }
}
// likelihood and impact: 1–5 (integers)
```

#### `kpi_dashboard`
```jsonc
{
  "layout": "kpi_dashboard",
  "content": {
    "kpis": [
      {
        "label":     "Revenue",
        "value":     "€4.2B",
        "unit":      "FY2025",
        "status":    "green",     // "green" | "amber" | "red"
        "trend":     "up",        // "up" | "down" | "flat"
        "context":   "↑ 18% YoY",
        "sparkline": [10, 11, 12, 13, 14]  // optional numeric series for mini-chart
      }
    ]
  }
}
// max 6 kpis; context ≤130 chars
```

#### `status_table`
```jsonc
{
  "layout": "status_table",
  "content": {
    "columns": ["Program", "Phase", "Status", "Owner", "Due"],
    "rows": [
      {"Program": "Alpha",  "Phase": "Phase 2", "Status": "On Track", "Owner": "J. Müller", "Due": "Q3 2026"},
      {"Program": "Beta",   "Phase": "Phase 1", "Status": "At Risk",  "Owner": "S. Patel",  "Due": "Q4 2026"}
    ]
  }
}
// RAG color auto-applied to columns whose header contains: RISK, STATUS, RAG, HEALTH, SEVERITY, PRIORITY
// "columns" is optional — pipeline derives it from first row's keys
// Auto-derived column order: program → phase → status → rag → health → milestone → owner → due → date → comment → notes → (others alphabetically)
// cell values ≤45 chars
```

#### `comparison_table`
```jsonc
{
  "layout": "comparison_table",
  "content": {
    "options":  ["Option A", "Option B", "Option C"],
    "features": [
      {"label": "Timeline",       "values": ["6 months", "12 months", "9 months"], "highlighted": false},
      {"label": "Cost",           "values": ["€2M", "€5M", "€3M"],                "highlighted": true},
      {"label": "Risk",           "values": ["Low", "High", "Medium"],             "highlighted": false}
    ]
  }
}
// Alternative: use "headers" + "rows" matrix format (auto-normalised):
// "headers": ["Criterion", "Option A", "Option B"]
// "rows":    [["Timeline", "6 months", "12 months"], ["Cost", "€2M", "€5M"]]
```

#### `score_table`
```jsonc
{
  "layout": "score_table",
  "content": {
    "rows": [
      {"label": "Market size",  "score": 4, "category": "Attractiveness", "note": "Large TAM"},
      {"label": "Margin profile","score": 5, "category": "Attractiveness", "note": ""}
    ],
    "scale":       5,           // Max score (default 5); renders as "x / 5"
    "scale_label": "Score"      // Optional label above score column
  }
}
// note ≤80 chars
```

**Harvey Ball variant** — set `rating_type: "harvey"` for consulting-style pie indicators.
Score is a fraction from 0.0 (empty) to 1.0 (fully filled). Scale and scale_label are ignored.
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
```

---

### Process & timeline layouts

#### `phase_process`
```jsonc
{
  "layout": "phase_process",
  "content": {
    "show_arrows": true,
    "phases": [
      {"label": "01", "title": "Discovery",   "body": "Research phase ≤100 chars", "status": "done"},
      {"label": "02", "title": "Design",      "body": "Co-creation workshops",    "status": "current"},
      {"label": "03", "title": "Deploy",      "body": "Pilot in 3 markets",       "status": "future"}
    ],
    "highlight_index": 1   // Optional: 0-based index of the highlighted (current) phase
  }
}
// status: "done" | "current" | "future"  — NOT highlighted: bool (deprecated)
// Max 5 phases; body ≤100 chars per phase
```

#### `gantt`
```jsonc
{
  "layout": "gantt",
  "content": {
    "quarters": ["Q1 2026", "Q2 2026", "Q3 2026", "Q4 2026"],
    "rows": [
      {"label": "Phase 1: Discovery", "start_q": 0, "duration_q": 2, "tone": "neutral"},
      {"label": "Phase 2: Pilot",     "start_q": 2, "duration_q": 1, "tone": "positive"},
      {"label": "Phase 3: Scale",     "start_q": 3, "duration_q": 1, "tone": "positive"}
    ]
  }
}
// start_q: 0-based index into quarters[]; duration_q: number of quarters
// tone: "positive" | "negative" | "neutral" — colors the bar
```

#### `milestone_timeline`
```jsonc
{
  "layout": "milestone_timeline",
  "content": {
    "milestones": [
      {"date": "Jan 2026", "title": "Study start",    "body": "First patient enrolled ≤80 chars", "status": "done"},
      {"date": "Mar 2026", "title": "Interim readout","body": "Data cut for DSMB review",         "status": "current"},
      {"date": "Dec 2026", "title": "Top-line data",  "body": "Primary endpoint results",         "status": "future"}
    ]
  }
}
// status: "done" | "current" | "future"
// Aliases accepted: "completed"→"done", "active"/"in_progress"→"current", "upcoming"/"planned"→"future"
// title alias: "label" → "title"; body alias: "description" → "body"
// body ≤80 chars per milestone
```

#### `circular_flow`
```jsonc
{
  "layout": "circular_flow",
  "content": {
    "phases": [
      {"label": "Plan",    "body": "Set objectives",  "icon": "target"},
      {"label": "Execute", "body": "Run experiments", "icon": "gear"},
      {"label": "Review",  "body": "Assess outcomes", "icon": "chart_bar"},
      {"label": "Adapt",   "body": "Iterate design",  "icon": "trending_up"}
    ]
  }
}
// icon is optional; 3–6 phases recommended
// icon names: same 26-name registry as topic_set (see §9 topic_set for full list)
```

#### `arrow_chain`
```jsonc
{
  "layout": "arrow_chain",
  "content": {
    "steps": [
      {"label": "Identify",  "body": "Define the problem space",   "highlighted": false},
      {"label": "Analyse",   "body": "Root cause analysis",         "highlighted": true},
      {"label": "Solve",     "body": "Develop solution options",   "highlighted": false}
    ],
    "consequence": {"label": "Outcome", "body": "Sustainable OEE gain"},
    "callout": {"type": "result", "text": "Expected saving: €2.4M annually"}
  }
}
// consequence is optional; highlighted: true adds visual emphasis
// callout is optional — same type/text schema as decision_rows callout
```

#### `funnel`
```jsonc
{
  "layout": "funnel",
  "content": {
    "inputs": [
      {"label": "Awareness",    "body": "12,000 leads",    "color": "teal"},
      {"label": "Engagement",   "body": "4,800 qualified", "color": "blue"},
      {"label": "Conversion",   "body": "960 trials",      "color": "green"}
    ],
    "output": {"label": "Revenue", "body": "€9.6M ARR"}
  }
}
// Key is "inputs" + "output" — NOT "stages"
// per-input color (optional): "gray" | "teal" | "blue" | "green" | "yellow" | "orange" | "red" | "pink" | "purple"
```

#### `journey_map`
```jsonc
{
  "layout": "journey_map",
  "content": {
    "phases": ["Awareness", "Consideration", "Trial", "Purchase", "Loyalty"],
    "rows": [
      {"label": "Patient",   "cells": ["Hears from HCP", "Reads leaflet", "Requests sample", "Fills Rx", "Refills"]},
      {"label": "Physician", "cells": ["Attends CME", "Reviews study", "Prescribes", "Follows up", "Advocates"]}
    ]
  }
}
// "actors" is also accepted as an alias for "rows" — both work equally
// Rich format also accepted: rows[i].steps = [{stage, action, emotion}] — action used as cell text
// If steps format is used, phases are auto-derived from the first row's stage values
```

#### `fishbone`
```jsonc
{
  "layout": "fishbone",
  "content": {
    "effect": "High patient dropout rate",
    "bones": [
      {"label": "Logistics",  "causes": ["Travel burden", "Long wait times"]},
      {"label": "Side effects","causes": ["Nausea", "Fatigue"]},
      {"label": "Motivation", "causes": ["Lack of perceived benefit"]}
    ]
  }
}
// Key is "bones" — NOT "causes" at top level
```

---

### Decision & analysis layouts

#### `2x2_matrix`
```jsonc
{
  "layout": "2x2_matrix",
  "content": {
    "x_axis": {"label": "Market Attractiveness", "low": "Low", "high": "High"},
    "y_axis": {"label": "Competitive Strength",  "low": "Weak", "high": "Strong"},
    "quadrants": {
      "top_left":     {"label": "Invest selectively", "items": ["Product B"]},
      "top_right":    {"label": "Invest aggressively", "items": ["Product A"], "highlighted": true},
      "bottom_left":  {"label": "Divest",              "items": ["Product D"]},
      "bottom_right": {"label": "Harvest",             "items": ["Product C"]}
    }
  }
}
// Layout key: "2x2_matrix" — NOT "matrix_2x2"
// Use whenever the source has exactly 4 equal boxes/quadrants in a 2-row grid.
// Do NOT use two_column for 4-item grids — two_column silently drops items 3 and 4.
// Quadrant labels: ≤30 chars each — short noun phrases only, not full sentences.
// Framework presets: "framework": "bcg" | "swot" | "ansoff" | "risk" — auto-fills axis/quadrant labels
```

#### `decision_rows`
```jsonc
{
  "layout": "decision_rows",
  "content": {
    "decisions": [
      {"number": 1, "title": "Approve budget increase", "desc": "Increase R&D budget by €12M for FY2026 ≤160 chars", "owner": "CFO", "tone": "positive"},
      {"number": 2, "title": "Defer market entry",       "desc": "Delay EU launch to Q2 2027 pending trial data",     "owner": "CMO", "tone": "neutral"}
    ],
    "callout": {
      "type": "conclusion",
      "text": "Recommend Option 1 — decision required by 30 June"
    }
  }
}
// CANONICAL keys: "title" + "desc"  (desc ≤160 chars)
// "body" and "description" are also accepted as aliases for "desc"
// tone: "positive" | "negative" | "neutral"
// max 5 decisions
// callout is optional — renders a branded pill above the footer
// callout.type: "conclusion" (>>) | "result" (↓) | "next" (→) | "future" (✓)
```

#### `before_after`
```jsonc
{
  "layout": "before_after",
  "content": {
    "before_label": "TODAY",
    "after_label":  "TOMORROW",
    "before": {"title": "Fragmented processes", "items": ["Manual data entry", "Siloed systems", "Long cycle time"]},
    "after":  {"title": "Integrated platform",  "items": ["Automated workflows", "Single source of truth", "Real-time reporting"]}
  }
}
// Key is "items" only — NEVER "bullets" or "points"
// before_label / after_label optional; defaults "TODAY" / "TOMORROW"
```

#### `influence_diagram`
```jsonc
{
  "layout": "influence_diagram",
  "content": {
    "center": {"label": "Patient Outcomes", "body": "Primary goal"},
    "forces": [
      {"label": "Clinical evidence", "body": "Phase 3 data",   "side": "left",   "tone": "positive"},
      {"label": "Reimbursement",     "body": "HTA decisions",  "side": "right",  "tone": "neutral"},
      {"label": "HCP adoption",      "body": "KOL programs",   "side": "top",    "tone": "positive"},
      {"label": "Patient access",    "body": "Patient support","side": "bottom", "tone": "negative"}
    ]
  }
}
// Key is "forces" — NOT "nodes"
// side: "left" | "right" | "top" | "bottom"
// tone: "positive" | "negative" | "neutral"
```

---

### Organizational layouts

#### `org_chart`
```jsonc
{
  "layout": "org_chart",
  "content": {
    "root": {"name": "Dr. Sarah Chen", "title": "Chief Medical Officer"},
    "children": [
      {
        "name": "Dr. James Kim",  "title": "Head Oncology",
        "reports": [
          {"name": "Dr. Lena Müller", "title": "Senior Scientist"}
        ]
      },
      {"name": "Dr. Ana Patel", "title": "Head Neurology", "reports": []}
    ]
  }
}
```

#### `hub_spoke`
```jsonc
{
  "layout": "hub_spoke",
  "content": {
    "hub": {"label": "01", "title": "Patient-centric", "subtitle": "Core principle"},
    "spokes": [
      {"title": "Digital tools",    "body": "Remote monitoring ≤100 chars"},
      {"title": "Data integration", "body": "Real-world evidence"},
      {"title": "Access programs",  "body": "Patient support"}
    ]
  }
}
// hub is an OBJECT with label/title/subtitle — NOT a plain string
// spoke body ≤100 chars
// Framework preset: "framework": "porters_5" → auto-fills hub and spoke labels
```

#### `pillar_detail`
```jsonc
{
  "layout": "pillar_detail",
  "content": {
    "pillar_number": "02",
    "pillar_label":  "INNOVATION",
    "owner":         {"label": "Pillar Lead", "name": "Dr. Felix Wagner"},
    "sections": [
      {"label": "Objective",    "body": "Launch 3 digital tools by Q4 2026"},
      {"label": "Key actions",  "body": "Partner with MedTech startups"},
      {"label": "KPI",          "body": "NPS ≥ 8.5 among enrolled patients"}
    ]
  }
}
```

#### `icon_grid`
```jsonc
{
  "layout": "icon_grid",
  "content": {
    "columns": 3,
    "items": [
      {"icon": "users",     "title": "Patient First",   "body": "Outcomes over processes", "highlighted": false},
      {"icon": "chart_bar", "title": "Data-driven",     "body": "Decisions backed by evidence"},
      {"icon": "target",    "title": "Speed",           "body": "Bias for action",         "highlighted": true}
    ]
  }
}
// columns: 2 | 3 (default 3); max 9 items (excess silently dropped)
// icon MUST be a named string from the 26-icon registry (anything else renders as a plain dot):
//   alert, arrow_down, arrow_right, arrow_up, calendar,
//   chart_bar, chart_line, chart_pie, check, clock, doc,
//   flag, gear, globe, info, lightbulb, lock, money, search,
//   shield, target, trending_down, trending_up, users, x
// Pick the closest semantic match (see icon selection guide in §9 topic_set for guidance)
// highlighted: true fills the card background with Merck Purple
```

---

### Visual & story layouts

#### `word_cloud`
```jsonc
{
  "layout": "word_cloud",
  "content": {
    "words": [
      {"text": "Innovation", "weight": 5},
      {"text": "Patient",    "weight": 4},
      {"text": "Data",       "weight": 4},
      {"text": "Access",     "weight": 3},
      {"text": "Digital",    "weight": 3},
      {"text": "Pipeline",   "weight": 2}
    ]
  }
}
// weight: 1–5 (5 = largest / most prominent)
```

#### `pyramid`
```jsonc
{
  "layout": "pyramid",
  "content": {
    "orientation": "up",
    "tiers": [
      {"label": "Vision",   "body": "Become the leading rare disease company"},
      {"label": "Strategy", "body": "3 therapeutic area focus"},
      {"label": "Execution","body": "4 flagship programs by 2027"},
      {"label": "Metrics",  "body": "Revenue, NPS, pipeline value"}
    ]
  }
}
// orientation: "up" (wide base) | "down" (inverted, wide top)
// tiers listed top-to-bottom in JSON; rendered bottom-to-top when orientation is "up"
```

#### `venn`
```jsonc
{
  "layout": "venn",
  "content": {
    "circles": [
      {"label": "Efficacy",   "body": "Proven clinical benefit"},
      {"label": "Safety",     "body": "Tolerable side-effect profile"},
      {"label": "Access",     "body": "Affordable and reimbursed"}
    ],
    "intersection": "The winning therapeutic profile"
  }
}
// Key is "intersection" — NOT "overlap"
// 2 or 3 circles
```

#### `layered_stack`
```jsonc
{
  "layout": "layered_stack",
  "content": {
    "orientation": "vertical",
    "layers": [
      {"icon": "globe",     "label": "Platform",     "body": "Cloud-native data fabric"},
      {"icon": "shield",    "label": "Security",     "body": "SOC 2 Type II, GDPR-compliant"},
      {"icon": "chart_bar", "label": "Intelligence", "body": "Real-time AI/ML pipelines"},
      {"icon": "lock",      "label": "Interface",    "body": "Web + mobile applications"}
    ]
  }
}
// orientation: "vertical" | "horizontal"
// icon uses the same 26-name registry as topic_set (see §9 topic_set for icon selection guide)
// Unknown icon names fall back to a generic target icon (no error)
```

#### `photo_text`
```jsonc
{
  "layout": "photo_text",
  "content": {
    "image_path":  "images/lab_scientist.jpg",   // Relative to project working directory
    "image_label": "Phase 2 trial site, Berlin",
    "image_side":  "left",                        // "left" | "right"
    "title":       "Our research team",
    "bullets": [
      "240 scientists across 8 countries",
      "3 dedicated rare disease research hubs",
      "Partnered with 18 academic centers"
    ]
  }
}
// image_path must be inside the project directory (security restriction)
// Allowed extensions: .png .jpg .jpeg .gif .bmp .tiff .webp .emf .wmf
```

#### `key_question`
```jsonc
{
  "layout": "key_question",
  "section_number": null,
  "content": {
    "question": "Should we scale operations now or invest in infrastructure first?",
    "context":  "Decision required before Q3 planning — impacts budget allocation"
  }
}
// Renders: centred circle (?) icon with four directional arrows pointing in,
//          large question text below, smaller context line below that.
// Use as a discussion-framing slide before a decision_rows or before_after slide.
// section_number: null recommended (structural framing slide)
// question ≤200 chars; context ≤130 chars
```

#### `road_to_success`
```jsonc
{
  "layout": "road_to_success",
  "content": {
    "stages": [
      {"title": "Discover",  "body": "Map current state and identify key pain points across the organisation"},
      {"title": "Build",     "body": "Prototype and validate with stakeholders in three agile sprints"},
      {"title": "Scale",     "body": "Full rollout with change management programme — Q4 target"}
    ],
    "milestones": ["Kick-off", "Sprint 1", "Pilot", "Review", "Launch"]
  }
}
// Renders: horizontal accent line (the "road") spanning the slide,
//          small dot per milestone (labels alternate above/below the line),
//          2–4 stage columns with accent subheadline + body text below the road.
// stages: 2–4 items; body ≤130 chars
// milestones: 2–6 short label strings (≤20 chars each); omit for clean timeline
// Use for transformation roadmaps, implementation plans, go-to-market journeys.
```

---

### Additional image placement (any layout)

Any slide can have an image placed as an overlay by adding an `image` key in `content`:

```jsonc
{
  "content": {
    "image": {
      "path":      "images/graphic.png",
      "placement": "right_panel"    // "right_panel" | "hero" | "background"
    }
  }
}
// placement presets (in inches, x y w h):
//   right_panel: 8.50 3.10 4.50 3.50
//   hero:        6.83 1.30 6.00 5.00
//   background:  0.00 0.00 13.33 7.50
```

---

### Appendix slides

Any slide can be moved to an appendix section by setting the `appendix` flag:

```jsonc
{
  "layout": "two_column",
  "appendix": true,
  "action_title": "Supporting detail: methodology",
  "content": { ... }
}
// Appendix slides are numbered A1, A2, ... and rendered after the main deck.
// category is auto-set to "APPENDIX".
// Appendix slides do not count toward section_number sequential validation.
```

---

## 10. Item count caps (excess items silently dropped — no error raised)

| Layout | Field | Max items |
|---|---|---|
| `cover` | `key_messages` | 4 |
| `exec_summary` | `key_messages` | 5 |
| `agenda` | `chapters` | 12 |
| `phase_process` | `phases` | 5 |
| `decision_rows` | `decisions` | 5 |
| `topic_set` | `topics` | 6 |
| `icon_grid` | `items` | 9 |
| `kpi_dashboard` | `kpis` | 6 |
| `cover` | `authors` | 5 |

---

## 11. Character limits (hard limits — overflow is truncated silently)

| Field | Location | Limit |
|---|---|---|
| `action_title` | slide-level | 120 chars (60 for cover) |
| `takeaway` | slide-level | 90 chars |
| `body` | inside content objects | 130 chars (general) |
| `desc` | inside content objects | 160 chars |
| `context` | inside content objects | 130 chars |
| `note` | inside content objects | 80 chars |
| `exec_summary key_messages[].body` | content | 120 chars |
| `two/three/four_column body` | content | 300 chars (full prose allowed) |
| `two/three/four_column items[] each item` | content | 120 chars per bullet |
| `label_rows rows[].body` | content | 300 chars (full prose allowed) |
| `decision_rows decisions[].desc` | content | 160 chars |
| `phase_process phases[].body` | content | 100 chars |
| `milestone_timeline milestones[].body` | content | 80 chars |
| `hub_spoke spokes[].body` | content | 100 chars |
| `topic_set topics[].body` | content | 120 chars |
| `status_table cell values` | content | 45 chars |

When content is dense, prefer fewer items with tighter text over many items with long text.

---

## 12. Framework presets (optional shortcut)

Add `"framework": "<key>"` to `content` to auto-populate axis/quadrant/column labels:

| Layout | Framework key | What it pre-fills |
|---|---|---|
| `2x2_matrix` | `"bcg"` | Growth-Share Matrix axes and quadrant labels |
| `2x2_matrix` | `"swot"` | Strengths/Weaknesses/Opportunities/Threats |
| `2x2_matrix` | `"ansoff"` | Market Penetration / Development / Diversification |
| `2x2_matrix` | `"risk"` | Likelihood × Impact axes |
| `three_column` | `"4ps"` | Product / Price / Place / Promotion |
| `three_column` | `"value_disciplines"` | Operational Excellence / Customer Intimacy / Product Leadership |
| `three_column` | `"balanced_scorecard"` | Financial / Customer / Internal / Learning |
| `four_column` | same as three_column | |
| `vertical_numbered` | `"kotter_8"` | Kotter's 8-step change model |
| `vertical_numbered` | `"adkar"` | Awareness / Desire / Knowledge / Ability / Reinforcement |
| `hub_spoke` | `"porters_5"` | Porter's Five Forces |

When using a framework preset, you can still override individual items by providing the full content alongside `"framework"`.

---

## 13. Accepted aliases (normalised automatically by the dispatcher)

These input formats are accepted and converted at build time — you may use either form:

### `milestone_timeline` — status aliases
| Input string | Normalises to |
|---|---|
| `"COMPLETED"`, `"complete"`, `"done"`, `"finished"` | `"done"` |
| `"active"`, `"current"`, `"in_progress"`, `"in progress"`, `"inprogress"` | `"current"` |
| `"upcoming"`, `"future"`, `"pending"`, `"planned"` | `"future"` |
| Any other string | `"future"` (safe default) |

Labels `"label"` → `"title"` and `"description"` → `"body"` are also auto-aliased.

### `decision_rows` — body key aliases
| Input | Normalises to |
|---|---|
| `"desc"` or `"description"` | canonical (preferred) |
| `"body"` | accepted equally |
| `"text"` (only when desc/body/description are all absent) | mapped to `"body"` |

Always use `"title"` + `"desc"` as the canonical form. `"text"` is a graceful fallback, not a recommended key.

### `journey_map` — row format aliases
Both formats are accepted:
- Native: `{label, cells: [str, ...]}`
- Rich steps: `{name|actor|label, steps: [{stage, action, emotion}, ...]}` — `action` text becomes `cells`; `stage` values of the first row auto-populate `phases`

Also: `"actors"` key accepted as alias for `"rows"` in content. Both work equally; `"rows"` is the native form.

### `comparison_table` — matrix format alias
Two formats accepted (native wins if both present):
- Native: `{options: [str], features: [{label, values: [str]}]}`
- Matrix: `{headers: [str], rows: [[str], ...]}` — first column is the row label

### `waterfall_slide` — bar type aliases
| Input type | Renders as |
|---|---|
| `"positive"` | Up bar |
| `"negative"` | Down bar (value coerced to positive via `abs()`) |
| `"total"` | Anchored bar (first occurrence = baseline, subsequent = closing bar) |

### `label_rows` — label_color format
`label_color` accepts a 6-character hex string (e.g. `"2DBECD"`, `"#EB3C96"`), an RGB tuple, or null (defaults to Merck Purple).

---

## 14. Layout selection guide

Use this table to choose the right layout for your content. When in doubt, prefer layouts that match the **shape** of the information (comparisons → columns, sequences → processes, single numbers → hero_stat).

### Processes and sequences

| Content type | Best layout | Notes |
|---|---|---|
| 2–5 sequential phases with clear before/during/after status | `phase_process` | Show status: done/current/future |
| 3–6 steps in a chain leading to a single outcome | `arrow_chain` | Use `consequence` for the final outcome |
| Steps that repeat in a cycle (continuous improvement, PDCA) | `circular_flow` | 3–6 phases ideal |
| 4–8 ordered steps in list form | `vertical_numbered` | Use framework presets (kotter_8, adkar) when applicable |
| Patient or customer journey with multiple actor rows | `journey_map` | Best with 3–6 phases and 2–4 actor rows |
| Project milestones with specific dates | `milestone_timeline` | Use status: done/current/future |
| Multi-workstream schedule over quarters | `gantt` | Columns = quarters; rows = workstreams |
| Many inputs converging to one output | `funnel` | Use `inputs` (not "stages") |
| Root cause analysis | `fishbone` | Effect + bones (categories) + causes |
| Transformation or implementation roadmap with named milestones | `road_to_success` | 2–4 stages + 2–6 milestone labels |

### Comparisons and options

| Content type | Best layout | Notes |
|---|---|---|
| Current state vs. future state | `before_after` | Use items, not bullets |
| 2 options with label + items | `two_column` | tone: positive/negative/neutral per column |
| 3 options | `three_column` | Exactly 3 columns array items |
| 4 options or time periods (Q1–Q4) | `four_column` | Exactly 4 columns array items |
| N options × N criteria evaluation matrix | `comparison_table` | Highlighted rows for key criteria |
| Alternatives with numeric scoring | `score_table` | Use category field to group rows |
| Explicit pros and cons for a named topic | `pros_cons` | Customize pros_label / cons_label |
| Factors influencing a central outcome | `influence_diagram` | Assign forces to left/right/top/bottom |

### Data and charts

| Content type | Best layout | Notes |
|---|---|---|
| Single dominant number with one line of context | `hero_stat` | section_number: null |
| 3–4 KPIs in a horizontal band | `stat_strip` | body ≤80 chars per stat |
| 4–6 KPIs with RAG status and trend | `kpi_dashboard` | Include sparkline arrays for mini-charts |
| Category comparison (at one point in time) | `chart_slide` type `column` or `bar` | bar when labels are long or >5 categories |
| Trend over 4+ time periods | `chart_slide` type `line` | Use area for stacked/volume emphasis |
| Before/after comparison across 2–8 named items | `chart_slide` type `slope` | Requires different data format (see §9) |
| Revenue/cost bridge | `waterfall_slide` | bar type: total/positive/negative |
| Part-of-whole breakdown | `donut_chart` | center_value + center_label for totals |
| Multi-axis capability or spider assessment | `radar_chart` | Values 0–100 per axis |
| Risk positioning by likelihood and impact | `risk_heatmap` | likelihood/impact: 1–5 integers |
| Project/program status tracking | `status_table` | RAG auto-colored on STATUS/RISK column headers |

### Decisions and strategic analysis

| Content type | Best layout | Notes |
|---|---|---|
| 1–5 explicit decisions for stakeholder approval | `decision_rows` | Auto-promotes to merck_executive |
| Prioritization across 2 dimensions | `2x2_matrix` | Use framework presets: bcg/swot/ansoff/risk |

### Structure and narrative

| Content type | Best layout | Notes |
|---|---|---|
| 3–6 concepts with icons and short descriptions | `topic_set` | Use icon selection guide in §9 |
| Concept grid with emoji icons and cards | `icon_grid` | Emoji only; highlighted cards for emphasis |
| Strategic hierarchy (platform, security, app) | `layered_stack` | orientation: vertical or horizontal |
| One central concept with radiating pillars | `hub_spoke` | hub is an object, not a string |
| Strategic pillars in detail (one per slide) | `pillar_detail` | One slide per pillar |
| Reporting structure | `org_chart` | root + children + reports |
| Label-value pairs (objectives, principles, criteria) | `label_rows` | body ≤130 chars per row |
| Priority or importance hierarchy | `pyramid` | orientation: up (base = foundation) or down (base = outcome) |
| 2–3 overlapping concepts with shared zone | `venn` | Key: intersection, not overlap |
| Opening impactful statement or attributed quote | `pull_quote` | section_number: null |
| Discussion framing or key decision question | `key_question` | Use before decision_rows or before_after; section_number: null recommended |
| Word frequency or theme visualization | `word_cloud` | weight 1–5 |
| Photo with bullet context | `photo_text` | image must be in project directory |
| Key messages for C-suite or board | `exec_summary` | Max 5 messages; section_number: null |

### Edge cases: when two layouts suit the same content

| Situation | Choose |
|---|---|
| 3-step process — but steps are cyclical | `circular_flow` over `phase_process` |
| 3-step process — but steps are strictly linear with an output | `arrow_chain` over `phase_process` |
| 3-step process — but steps have done/current/future status | `phase_process` |
| Before/after with pros and cons | `before_after` when the focus is transformation; `pros_cons` when the focus is trade-offs |
| Many KPIs — but some have trend data and some don't | `kpi_dashboard` (sparkline is optional per KPI) |
| Single stat — but it is also the slide's main insight | `hero_stat` (section_number: null) over `stat_strip` |
| 2-column — but the columns are strict current/future | `before_after` over `two_column` |
| Decisions — but they are really options to compare | `comparison_table` if no clear owner/tone needed; `decision_rows` if each item needs an owner and a tone |

---

## 15. Critical wrong-key warnings

These mistakes produce silently empty or broken slides — no error is raised:

| Layout | WRONG key | CORRECT key |
|---|---|---|
| `decision_rows` | `"text"` as only body key | `"title"` + `"desc"` (canonical) |
| `phase_process` | `"highlighted": true` | `"status": "current"` |
| `funnel` | `"stages"` | `"inputs"` + `"output"` |
| `journey_map` | (either accepted) | `"rows"` or `"actors"` both work |
| `venn` | `"overlap"` | `"intersection"` |
| `fishbone` | `"causes"` at top level | `"bones": [{label, causes}]` |
| `influence_diagram` | `"nodes"` | `"forces"` |
| `before_after` | `"bullets"` or `"points"` | `"items"` |
| `hub_spoke` | `hub: "Center label"` (string) | `hub: {label, title, subtitle}` (object) |
| `chart_slide` | `"chart"` or `"waterfall"` as layout key | `"chart_slide"` or `"waterfall_slide"` |
| `2x2_matrix` | `"matrix_2x2"` as layout key | `"2x2_matrix"` |
| `waterfall_slide` | `"up"` / `"down"` bar types | `"positive"` / `"negative"` |
| `cover` | `subtitle` inside `content` | `subtitle` at slide level |
| `section_divider` | chapter number in `section_number` | chapter number in `content.number` |
| `icon_grid` | emoji string (e.g. `"🎯"`) — renders as text, not vector icon | named icon string (e.g. `"target"`) from the 26-icon registry |
| `topic_set` | emoji string (e.g. `"🎯"`) as icon | named icon string (e.g. `"target"`) |
| `score_table` Harvey Ball | `score: 4` (integer out of scale) | `score: 0.75` (float 0.0–1.0) when `rating_type: "harvey"` |
| `road_to_success` | `"stages"` containing phases without `title`/`body` | `stages: [{title, body}]` |
| callout (decision_rows etc.) | `callout.type: "done"` or `"check"` | `type` must be `"conclusion"` / `"result"` / `"next"` / `"future"` |

---

## 16. Quality rules checklist

1. `action_title` is always a declarative sentence. Never a noun phrase.
2. Use `"INSIGHT; CONSEQUENCE"` semicolon pattern in action titles.
3. Numbers beat adjectives: `"18% growth"` not `"significant growth"`.
4. `section_number` is `null` for structural slides; unique sequential integer for all others.
5. Auto-promote: for `category` matching `"Executive Summary"`, `"Recommendation"`,
   `"Decision Request"`, `"Risk"`, `"Tradeoff"` — set `style: "merck_executive"` explicitly too.
6. First slide: `cover`. Last slide: `close`. Agenda at slide 2 if ≥ 5 content slides.
7. Max items: `exec_summary` ≤ 5, `decision_rows` ≤ 5, `phase_process` ≤ 5, `kpi_dashboard` ≤ 6.
8. When content cannot be inferred from source, write `"[PLACEHOLDER: description]"` as value.
9. Never invent custom colors. All color choices are handled by `meta.color_theme`.
10. `source` field (slide-level, free text) is shown at bottom of slide. Use for data attribution.
11. `subtitle` always lives at slide level. Never put it inside `content`.
12. Section dividers: `section_number: null`; display chapter number via `content.number`.
13. Vary layout families: no more than 2 consecutive slides from the same family.
14. **Chrome and takeaway:** `takeaway` must be `null` on ALL slides unless `chrome.takeaway_bands`
    is explicitly `true`. When `takeaway_bands` is true, write a takeaway on every content slide.
    Omit takeaway regardless on: cover, agenda, section_divider, close, exec_summary, hero_stat, pull_quote.
15. **Text preservation:** Copy all bullet points and body text from the source VERBATIM into
    content fields. Do not paraphrase, summarise, reorder, or rewrite any source text.
    Your only creative licence is: `action_title` (derived from source heading, ≤120 chars),
    `takeaway` (new "so what" sentence you write), and structural fields (layout, page_function,
    style, category, section_number, color_theme). For figure/chart blocks in the source, use
    `"[PLACEHOLDER: <original description>]"` — never invent chart data or axis values.

---

## 17. Complete minimal example plan

```json
{
  "meta": {
    "region":         "EU",
    "deck_label":     "Oncology Pipeline Review",
    "classification": "Internal",
    "month_year":     "June 2026",
    "audience":       "Oncology Leadership Team",
    "deck_style":     "merck_executive",
    "color_theme":    "organic"
  },
  "storyline": [
    "Pipeline momentum accelerated in H1 2026",
    "Three programs ready for Phase 3 investment decision",
    "Resource reallocation is needed to sustain momentum"
  ],
  "slides": [
    {
      "page": 1,
      "page_function": "Cover",
      "layout": "cover",
      "action_title": "Oncology Pipeline Review; June 2026",
      "subtitle": null,
      "section_number": null,
      "style": "merck_executive",
      "category": null,
      "takeaway": null,
      "source": null,
      "notes": null,
      "content": {
        "authors": [{"name": "Dr. Felix Wagner", "title": "Head of Oncology R&D"}]
      }
    },
    {
      "page": 2,
      "page_function": "Agenda",
      "layout": "agenda",
      "action_title": "Three topics for today",
      "section_number": null,
      "style": "merck_executive",
      "category": null,
      "takeaway": null,
      "source": null,
      "notes": null,
      "content": {"chapters": []}
    },
    {
      "page": 3,
      "page_function": "Section Divider",
      "layout": "section_divider",
      "action_title": "Pipeline Status",
      "section_number": null,
      "style": "merck_executive",
      "category": null,
      "takeaway": null,
      "source": null,
      "notes": null,
      "content": {"number": "01"}
    },
    {
      "page": 4,
      "page_function": "Evidence",
      "layout": "phase_process",
      "action_title": "MRK-001 has reached Phase 2b; readout expected Q4 2026",
      "section_number": 1,
      "style": "inherit",
      "category": "EVIDENCE",
      "takeaway": "Enrollment on track; no safety signals to date.",
      "source": "Clinical Operations, May 2026",
      "notes": null,
      "content": {
        "show_arrows": true,
        "phases": [
          {"label": "01", "title": "Phase 1",  "body": "Dose escalation complete",    "status": "done"},
          {"label": "02", "title": "Phase 2a", "body": "ORR 42% at RP2D",             "status": "done"},
          {"label": "03", "title": "Phase 2b", "body": "250 patients enrolled",       "status": "current"},
          {"label": "04", "title": "Phase 3",  "body": "Start pending readout",       "status": "future"}
        ]
      }
    },
    {
      "page": 5,
      "page_function": "Decision Request",
      "layout": "decision_rows",
      "action_title": "Three decisions needed to advance pipeline in H2 2026",
      "section_number": 2,
      "style": "merck_executive",
      "category": "Decision Request",
      "takeaway": "All three decisions have a Q3 2026 deadline.",
      "source": null,
      "notes": "Emphasise the budget decision timeline.",
      "content": {
        "decisions": [
          {"number": 1, "title": "Approve Phase 3 budget", "desc": "€45M incremental investment for MRK-001 Phase 3 start", "owner": "CFO", "tone": "positive"},
          {"number": 2, "title": "Greenlight MRK-003",    "desc": "Proceed to IND filing; assign dedicated CMC team",       "owner": "CMO", "tone": "neutral"},
          {"number": 3, "title": "Pause MRK-007",         "desc": "Suspend enrollment pending competitor data readout",     "owner": "CMO", "tone": "negative"}
        ]
      }
    },
    {
      "page": 6,
      "page_function": "Close",
      "layout": "close",
      "action_title": "Next steps",
      "section_number": null,
      "style": "merck_executive",
      "category": null,
      "takeaway": null,
      "source": null,
      "notes": null,
      "content": {
        "action_statement": "Decisions required by 30 June — please confirm attendance for the sign-off meeting"
      }
    }
  ]
}
```
