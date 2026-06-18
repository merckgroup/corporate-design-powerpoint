# Slide Plan Schema

The agent produces a structured slide plan (JSON-shaped) before generating any code. The plan captures deck-level meta, storyline, and per-slide content. The build step reads the plan and calls layout functions in `merck_layouts.py`.

This schema is the source of truth. The agent does not generate slides outside this structure.

---

## TOP-LEVEL SHAPE

```json
{
  "meta": { ... },
  "storyline": [ ... ],
  "slides": [ ... ]
}
```

---

## META BLOCK

```json
"meta": {
  "deck_label":     "Finance Systems",
  "classification": "Confidential",
  "month_year":     "May 2026",
  "audience":       "Executive leadership",
  "deck_style":     "merck_executive",
  "show_disclaimer": false,
  "topic":          "Project Lumina Q2 readiness",
  "deck_objective": "Brief CFO on Lumina audit readiness and unblock the IT package handover.",
  "single_sentence_takeaway": "We are on track for Q2 data products but the IT package handover risks a 4-week slip.",
  "final_ask": "Approve Tom Kistinger as LS FP&A Data Owner and unblock Packages 3 to 5."
}
```

| Field | Required | Values |
|-------|----------|--------|
| deck_label | yes | Free text, used in footer |
| classification | yes | Public / Internal / Confidential |
| month_year | yes | "Month YYYY" format |
| audience | yes | Executive leadership / Senior management / Functional team / Mixed audience / External |
| deck_style | yes | merck_executive / merck_corporate / merck_storytelling |
| show_disclaimer | yes | Boolean; only true for external-facing |
| topic | yes | Short noun phrase for the deck topic |
| deck_objective | yes | One sentence; what the deck is for |
| single_sentence_takeaway | yes | The whole deck argument in one line |
| final_ask | yes | The decision or action requested |
| light_footer | optional | Boolean; suppresses dark purple footer band |
| cover_top_bar | optional | Boolean; full-width MERCK_GOLD bar at very top of cover |
| region | yes | `"EU"` (uses `Merck_Themed_Merck.pptx`) or `"USA"` (uses `Merck_Themed_Base_v1.pptx`). Default: `"EU"`. Controls template and brand identity. |

---

## STORYLINE BLOCK

3 to 5 chapter action sentences. These become the agenda and the structural backbone of the deck.

```json
"storyline": [
  "Lumina sub-domain governance is in place and on schedule for Q1 audit readiness.",
  "Data products timeline is healthy; IT package handover is the gating risk.",
  "Tom Kistinger is the right candidate for LS FP&A Data Owner.",
  "Three actions unblock Q2 delivery."
]
```

---

## SLIDES BLOCK

Each slide is one object. The order is the deck order.

```json
"slides": [
  {
    "page": 1,
    "page_function": "Cover",
    "layout": "cover",
    "style": "inherit",
    "action_title": "Project Lumina: Q2 Readiness and Decision Asks",
    "subtitle": "CFO Brief, May 2026",
    "content": {}
  }
]
```

### Per-slide fields

| Field | Required | Values |
|-------|----------|--------|
| page | yes | 1-indexed page number |
| section_number | yes (content slides) | UNIQUE sequential number (1, 2, 3, ..., N). Never the page number. Never shared across slides. |
| page_function | yes | Cover / Framing / Diagnosis / Evidence / Options / Recommendation / Roadmap / Decision Request / Risk / Tradeoff / Close / Section Divider / Hero Moment / Executive Summary |
| layout | yes | One of the layout names in the catalog below |
| style | yes | `inherit` OR one of merck_executive / merck_corporate / merck_storytelling |
| action_title | yes | Full declarative sentence (string OR list of `(text, italic_bool)` tuples) |
| category | optional but recommended | Short UPPERCASE tag (e.g. DIAGNOSIS, EVIDENCE, RISK). Also drives auto-promote. |
| subtitle | optional | One-line qualifier under the action title |
| takeaway | yes (content slides) | One-sentence "so what" |
| source | yes (data slides) | "Source: ..." |
| methodology_note | optional | Italic gray line above the source. Placed inside `content`. |
| content | yes | Layout-specific content payload (see catalog) |

### Auto-promote rule

If `category` exactly matches one of the following, both `style` and `page_function` must be set to that value, and `style` is forced to `merck_executive`:

- `"Executive Summary"`
- `"Recommendation"`
- `"Decision Request"`
- `"Risk"`
- `"Tradeoff"`

---

## LAYOUT CATALOG

### Critical field name rules

Wrong keys produce silently empty slides — no error is raised.

| Layout | ✓ Correct key | ✗ Never use |
|--------|--------------|-------------|
| `two_column` left/right | `items` | `bullets`, `points`, `list` |
| `before_after` before/after | `items` | `bullets`, `points` |
| `decision_rows` decisions[] | `title` + `desc` | `decision`, `body`, `text` |
| `2x2_matrix` quadrants | `items` (list of strings) | `body`, `text`, `content` |
| `phase_process` phases[] | `status: "current"` | `highlight_index`, `highlighted: true` |
| `columns` / `three_column` | `tone: "positive"` | `highlighted: true` |
| `funnel` | `inputs` + `output` | `stages` |
| `journey_map` | `actors` | `rows` |
| `fishbone` | `bones` | `causes` |
| `comparison_table` | `options` + `features` | `headers` + `rows` |
| `influence_diagram` | `forces` (with `side` + `tone`) | `nodes` |
| `venn` | `intersection` | `overlap` |
| `hub_spoke` center | `hub: {label, title, subtitle}` | `center: "string"` |
| `close` body | `action_statement` | `statement` |
| Chart layout key | `chart_slide` | `chart` |
| Waterfall layout key | `waterfall_slide` | `waterfall` |
| Matrix layout key | `2x2_matrix` | `matrix_2x2` |

---

### 1. cover

```json
{
  "layout": "cover",
  "action_title": "Project Lumina: Q2 Readiness",
  "subtitle": "audit readiness and decision asks\nGovernance | Data products | Decisions",
  "content": {
    "authors": [
      {"name": "Anoop Kumar", "title": "Head of Financial Systems"}
    ]
  }
}
```

`action_title` ≤ 60 characters. `subtitle` required. Do not include `key_messages`, `phases`, `category`, or `section_number` in cover content.

### 2. exec_summary (auto-promoted)

```json
{
  "layout": "exec_summary",
  "page_function": "Executive Summary",
  "category": "Executive Summary",
  "action_title": "We Are On Track for Q1 Audit, At Risk for Q2 Data Products",
  "takeaway": "Approval of the IT package handover unblocks Q2 delivery.",
  "content": {
    "key_messages": [
      {"label": "Q1 Audit Readiness", "body": "All sub-domain inventories complete."},
      {"label": "Q2 Data Products", "body": "Packages 3–5 not yet handed over; 4-week slip risk."},
      {"label": "Data Owner", "body": "Tom Kistinger nominated; awaiting confirmation."}
    ]
  }
}
```

### 3. agenda

```json
{
  "layout": "agenda",
  "action_title": "Agenda",
  "content": {
    "chapters": [
      {"number": "01", "title": "Q1 Audit Readiness", "subtitle": "On track, 92% complete"},
      {"number": "02", "title": "Q2 Data Products", "subtitle": "Gated on IT handover"},
      {"number": "03", "title": "Decision Asks", "subtitle": "Three items for CFO approval"}
    ]
  }
}
```

Supports up to 12 chapters. Leave `chapters` empty to auto-fill from content slides.

### 4. section_divider

```json
{
  "layout": "section_divider",
  "section_number": "02",
  "action_title": "Q2 Data Products",
  "content": {}
}
```

`section_number` at slide level is the big number rendered on the divider page.

### 5. chart_slide

```json
{
  "layout": "chart_slide",
  "category": "EVIDENCE",
  "action_title": "Margin Compression Tracks Volume Shift to Lower-Tier SKUs",
  "source": "Source: Internal LS Finance close, May 2026",
  "content": {
    "chart": {
      "type": "slope",
      "data": {
        "before_label": "Q1 2025",
        "after_label": "Q1 2026",
        "items": [
          ["Premium SKUs", 42.0, 38.5],
          ["Mid-tier SKUs", 28.0, 32.0],
          ["Entry-tier SKUs", 30.0, 29.5]
        ],
        "highlight_indices": [0]
      }
    },
    "callouts": [
      {"x_in": 9.5, "y_in": 3.2, "label": "350 bps premium-SKU drop", "direction": "up_right"}
    ]
  }
}
```

**Chart types:** `slope`, `dot`, `marimekko`, `waterfall`, `bar`, `small_multiples`

### 6. two_column

```json
{
  "layout": "two_column",
  "content": {
    "left":  {"label": "Current State", "body": "...", "tone": "negative", "items": ["..."]},
    "right": {"label": "Target State",  "body": "...", "tone": "positive", "items": ["..."]}
  }
}
```

### 7. three_column

```json
{
  "layout": "three_column",
  "content": {
    "columns": [
      {"label": "Phase 1", "body": "Prove value", "tone": "neutral", "items": ["..."]},
      {"label": "Phase 2", "body": "Widen scope", "tone": "neutral", "items": ["..."]},
      {"label": "Phase 3", "body": "Enterprise",  "tone": "positive","items": ["..."]}
    ]
  }
}
```

### 8. 2x2_matrix

```json
{
  "layout": "2x2_matrix",
  "content": {
    "x_axis": {"label": "Impact", "low": "Low", "high": "High"},
    "y_axis": {"label": "Effort", "low": "Low", "high": "High"},
    "quadrants": {
      "top_left":     {"label": "High effort, low impact",  "items": ["AtScale full rollout"]},
      "top_right":    {"label": "High effort, high impact", "highlighted": true, "items": ["Snowflake completion"]},
      "bottom_left":  {"label": "Low effort, low impact",  "items": ["Documentation cleanup"]},
      "bottom_right": {"label": "Low effort, high impact", "items": ["Tom as Data Owner"]}
    }
  }
}
```

### 9. phase_process

Use `status: "current"` to highlight a phase. Also valid: `"done"` (muted) and `"future"` (default). Do not use `highlight_index` — deprecated.

```json
{
  "layout": "phase_process",
  "content": {
    "show_arrows": true,
    "phases": [
      {"label": "Phase 1", "title": "Diagnose", "body": "Map current data flows.",     "status": "done",    "milestone": "Inventory complete"},
      {"label": "Phase 2", "title": "Design",   "body": "Define target architecture.", "status": "current", "milestone": "Approved design"},
      {"label": "Phase 3", "title": "Deliver",  "body": "Stand up Snowflake pipeline.","status": "future",  "milestone": "First controller live"}
    ]
  }
}
```

### 10. vertical_numbered

```json
{
  "layout": "vertical_numbered",
  "content": {
    "items": [
      {"title": "Approve Tom Kistinger as Data Owner", "body": "Unblocks Packages 3–5."},
      {"title": "Confirm IT timeline for Snowflake",   "body": "Mid-Q3 completion commitment."},
      {"title": "Authorize Fabric provisioning",       "body": "Power BI scale needs dedicated capacity."}
    ]
  }
}
```

### 11. waterfall_slide

```json
{
  "layout": "waterfall_slide",
  "content": {
    "chart": {
      "type": "waterfall",
      "data": {
        "bars": [
          {"label": "Q1 2025", "value": 4200, "type": "start"},
          {"label": "Volume",  "value":  600, "type": "up"},
          {"label": "Price",   "value": -180, "type": "down"},
          {"label": "Q1 2026", "value": 4620, "type": "end"}
        ]
      }
    }
  }
}
```

### 12. decision_rows (auto-promoted)

```json
{
  "layout": "decision_rows",
  "page_function": "Decision Request",
  "category": "Decision Request",
  "content": {
    "decisions": [
      {"number": "1", "title": "Approve Tom Kistinger as LS FP&A Data Owner", "desc": "Sign-off by 2026-06-15.", "owner": "Ronan Tatibouet", "tone": "neutral"},
      {"number": "2", "title": "Confirm Snowflake completion date", "desc": "Mid-Q3 2026 from IT.", "owner": "Carsten Goepp", "tone": "negative"}
    ]
  }
}
```

`tone`: `"neutral"` / `"positive"` / `"negative"`. Keys: `"title"` + `"desc"` — not `"text"` or `"body"`.

### 13. gantt

```json
{
  "layout": "gantt",
  "content": {
    "quarters": ["Q2 26", "Q3 26", "Q4 26", "Q1 27"],
    "rows": [
      {"label": "Lumina audit",           "start_q": 0.0, "duration_q": 1.0, "tone": "positive"},
      {"label": "Snowflake pipeline",     "start_q": 0.5, "duration_q": 1.5, "tone": "neutral"},
      {"label": "Data products handover", "start_q": 1.0, "duration_q": 1.0, "tone": "negative"}
    ]
  }
}
```

### 14. hero_stat

```json
{
  "layout": "hero_stat",
  "style": "merck_storytelling",
  "content": {
    "stat":    {"value": "87%", "label": "of controllers report data latency as their top friction"},
    "context": "Internal LS Finance survey, May 2026 (n=42)"
  }
}
```

### 15. close

```json
{
  "layout": "close",
  "page_function": "Close",
  "takeaway": "Decision today protects the Q2 delivery date.",
  "source": "Source: Internal LS Finance, May 2026",
  "content": {
    "action_statement": "Approve three decisions to unblock Q2 delivery and protect Lumina audit readiness."
  }
}
```

### 16. stat_strip

```json
{
  "layout": "stat_strip",
  "content": {
    "stats": [
      {"value": "40+",    "label": "CONTROLLERS",         "body": "Active across PS, AS, DS sectors."},
      {"value": "5,000+", "label": "DOCUMENTS / YEAR",    "body": "LBE decks, GSA surveys, QBR captures."},
      {"value": "Deep",   "label": "INSTITUTIONAL MEMORY","body": "Causal commentary behind every variance."}
    ]
  }
}
```

### 17. before_after

```json
{
  "layout": "before_after",
  "content": {
    "before": {"label": "TODAY",    "title": "The numbers can be inferred", "items": ["Retrieve from raw P&L exports", "Reconcile manually"]},
    "after":  {"label": "TOMORROW", "title": "The context must be remembered", "items": ["Structured controller commentary", "Lifecycle tracking"]}
  }
}
```

Always `items` — never `bullets` or `points`. Override labels with `before_label` / `after_label` at content level.

### 18. milestone_timeline

```json
{
  "layout": "milestone_timeline",
  "content": {
    "milestones": [
      {"date": "Apr 2026", "title": "Mismatch surfaces",  "body": "Lonza and Samsung growth split.",        "status": "done"},
      {"date": "May 2026", "title": "CFO ruling",         "body": "Sold-to confirmed as LS dimension.",      "status": "done"},
      {"date": "Today",    "title": "Operational backbone","body": "Curated view, governance, controls.",    "status": "current"}
    ]
  }
}
```

`status`: `"done"` (purple check), `"current"` (gold), `"future"` (gray hollow).

### 19. status_table

```json
{
  "layout": "status_table",
  "content": {
    "columns": ["DEPENDENCY", "OWNER", "RISK", "MITIGATION"],
    "rows": [
      {"dependency": "GSA Sold-to feasibility", "owner": "Daniel Pantazelos", "risk": "AMBER", "mitigation": "Close before publish."},
      {"dependency": "Power BI semantic model", "owner": "Tom Kistinger",     "risk": "RED",   "mitigation": "Fabric capacity escalation."}
    ]
  }
}
```

`risk` values: `AMBER`, `RED`, `GREEN`, `NEUTRAL`. RAG column auto-detected by header name.

### 20. hub_spoke

```json
{
  "layout": "hub_spoke",
  "content": {
    "hub": {"label": "ONE LS", "title": "CUSTOMER VIEW", "subtitle": "the precondition"},
    "spokes": [
      {"title": "Power BI semantic model",      "body": "Customer view as the LS standard dimension."},
      {"title": "6,400 reports rationalization", "body": "Cannot retire reports without one trusted source."},
      {"title": "Lumina LS FP&A sub-domain",    "body": "Pillar 3 controls live here."},
      {"title": "One LS Analytics SteerCo",     "body": "Endorsement forum for objective and publish date."}
    ]
  }
}
```

`hub` is an **object** (`label`, `title`, `subtitle`) — not a plain string.

### 21. pillar_detail

```json
{
  "layout": "pillar_detail",
  "content": {
    "pillar_number": "02",
    "pillar_label":  "PILLAR",
    "owner": {"label": "OWNER", "name": "Natalia Mocan, Future Platform Architecture"},
    "sections": [
      {"label": "WHAT IT IS", "body": "One LS-wide customer view, curated from divisional input."},
      {"label": "STATUS",     "body": "Curation in flight. Ready for endorsement before next SteerCo."}
    ]
  }
}
```

### 22. four_column

```json
{
  "layout": "four_column",
  "content": {
    "columns": [
      {"label": "PILLAR 1", "body": "Governance",   "tone": "positive", "items": ["Sub-domain inventory complete"]},
      {"label": "PILLAR 2", "body": "Curated view", "tone": "neutral",  "highlighted": true, "items": ["Curation in flight"]},
      {"label": "PILLAR 3", "body": "Controls",     "tone": "neutral",  "items": ["KPI framework drafted"]},
      {"label": "PILLAR 4", "body": "Adoption",     "tone": "neutral",  "items": ["Controller pilot"]}
    ]
  }
}
```

### 23. label_rows

```json
{
  "layout": "label_rows",
  "content": {
    "rows": [
      {"label": "PIPELINE LATENCY",   "body": "TM1 refresh lag forces fallback to legacy reporting."},
      {"label": "DATA TRUST",         "body": "Three sources for the same number."},
      {"label": "GOVERNANCE",         "body": "No formal Data Owner — pace of decisions stalls."}
    ]
  }
}
```

3–6 rows. Optional `label_color` overrides the default purple label fill.

### 24. circular_flow

```json
{
  "layout": "circular_flow",
  "content": {
    "phases": [
      {"label": "PLAN",  "body": "Define targets and scope.", "icon": "target"},
      {"label": "DO",    "body": "Execute the workstream."},
      {"label": "CHECK", "body": "Measure outcomes against plan."},
      {"label": "ACT",   "body": "Adjust and re-enter the cycle."}
    ]
  }
}
```

### 25. org_chart

```json
{
  "layout": "org_chart",
  "content": {
    "root": {"name": "Ronan Tatibouet", "title": "Head of LS Finance Data"},
    "children": [
      {
        "name": "Tom Kistinger", "title": "LS FP&A Data Owner",
        "reports": [
          {"name": "Natalia Mocan", "title": "Curation Lead"},
          {"name": "Julia Baer",    "title": "IT Package Lead"}
        ]
      },
      {"name": "Daniel Pantazelos", "title": "PS Data Steward"}
    ]
  }
}
```

### 26. topic_set

```json
{
  "layout": "topic_set",
  "content": {
    "topics": [
      {"label": "01", "title": "Single Source",   "body": "One canonical view per metric.", "icon": "database"},
      {"label": "02", "title": "Named Ownership", "body": "Every domain has a declared Data Owner."},
      {"label": "03", "title": "Audit Trail",     "body": "All changes logged with author and rationale."},
      {"label": "04", "title": "Governed Access", "body": "Role-based access at the semantic layer.", "icon": "lock"}
    ]
  }
}
```

### 27. arrow_chain

```json
{
  "layout": "arrow_chain",
  "content": {
    "steps": [
      {"label": "TRIGGER",  "body": "Snowflake refresh lag exceeds 4 hours."},
      {"label": "RESPONSE", "body": "Controllers fall back to TM1.", "highlighted": true},
      {"label": "EFFECT",   "body": "Two data sources produce conflicting P&L numbers."},
      {"label": "OUTCOME",  "body": "Manual reconciliation consumes 2 days per close cycle."}
    ],
    "consequence": {"label": "RESULT", "body": "Close date slips 3–5 days every quarter."}
  }
}
```

---

## CREATIVE LAYOUTS (variety_mode: "creative")

### 28. pull_quote

```json
{
  "layout": "pull_quote",
  "style": "merck_storytelling",
  "content": {
    "quote": "The best way to predict the future is to invent it.",
    "attribution": "Alan Kay, Xerox PARC",
    "context": "Vision & Direction"
  }
}
```

### 29. donut_chart

```json
{
  "layout": "donut_chart",
  "content": {
    "segments": [
      {"label": "Digital Direct", "value": 38},
      {"label": "Partner Portal", "value": 25},
      {"label": "Field Sales",    "value": 22},
      {"label": "Distributors",   "value": 15}
    ],
    "center_value": "63%",
    "center_label": "Digital Share",
    "legend_title": "Revenue by Channel"
  }
}
```

### 30. kpi_dashboard

```json
{
  "layout": "kpi_dashboard",
  "content": {
    "kpis": [
      {"label": "Net Revenue", "value": "4.2B", "unit": "+8% YoY",   "status": "green", "trend": "up",   "context": "Ahead of plan", "sparkline": [3.6, 3.7, 3.9, 4.0, 4.2]},
      {"label": "Churn Rate",  "value": "4.8%", "unit": "vs 3% tgt", "status": "red",   "trend": "up",   "context": "Accelerating"}
    ]
  }
}
```

`status`: `"green"` / `"amber"` / `"red"`. `trend`: `"up"` / `"down"` / `"flat"`. Max 6 KPIs.

### 31. icon_grid

```json
{
  "layout": "icon_grid",
  "content": {
    "columns": 3,
    "items": [
      {"icon": "L",  "title": "Identity & Access", "body": "Zero-trust SSO and role-based provisioning."},
      {"icon": "AI", "title": "AI Automation",     "body": "LLM processing reduces effort by 70%.", "highlighted": true},
      {"icon": "S",  "title": "Security",          "body": "ISO 27001 certified, SOC 2 Type II."}
    ]
  }
}
```

### 32. journey_map

```json
{
  "layout": "journey_map",
  "content": {
    "phases": ["Discover", "Evaluate", "Purchase", "Onboard", "Expand"],
    "actors": [
      {"name": "Customer", "cells": ["Searches online", "Requests demo", "Signs contract", "Attends kick-off", "Buys add-ons"]},
      {"name": "Sales",    "cells": ["SDR response <2h", "AE demo",      "Legal & close",  "Handoff to CS",   "QBR + upsell"]}
    ]
  }
}
```

Key: `actors` — not `rows`.

### 33. funnel

```json
{
  "layout": "funnel",
  "content": {
    "inputs": [
      {"label": "Market Expansion",    "body": "Enter 3 new geographies in H2."},
      {"label": "Portfolio Refresh",   "body": "Launch 2 new SKUs; retire 4 underperformers."},
      {"label": "Cost Transformation", "body": "Automate 60% of back-office processes."},
      {"label": "Brand Investment",    "body": "Increase digital SOV from 18% to 28%."}
    ],
    "output": {"label": "Outcome", "body": "500M incremental revenue at 65%+ gross margin by FY27."}
  }
}
```

Keys: `inputs` + `output` — not `stages`.

### 34. comparison_table

```json
{
  "layout": "comparison_table",
  "content": {
    "options": ["Build Internal", "SaaS Partner", "Hybrid Model"],
    "features": [
      {"label": "Time to market",   "values": ["no",  "yes", "partial"]},
      {"label": "Data sovereignty", "values": ["yes", "no",  "yes"],    "highlighted": true},
      {"label": "Total cost (5yr)", "values": ["no",  "yes", "partial"]}
    ]
  }
}
```

`values`: `"yes"` (green check), `"no"` (red cross), `"partial"` (amber tilde), or custom text.

### 35. score_table

```json
{
  "layout": "score_table",
  "content": {
    "scale": 5,
    "scale_label": "MATURITY (1-5)",
    "rows": [
      {"label": "Data governance",        "score": 4.0, "category": "Foundation",   "note": "Strong policy framework"},
      {"label": "API-first architecture", "score": 2.0, "category": "Integration",  "note": "Legacy monolith blocks progress"},
      {"label": "AI / ML capability",     "score": 1.5, "category": "Intelligence", "note": "Pilot only"}
    ]
  }
}
```

Keys: `rows` + `scale` — not `items` + `headers`.

### 36. influence_diagram

```json
{
  "layout": "influence_diagram",
  "content": {
    "center": {"label": "Pricing Strategy", "body": "Value-based pricing across all segments by Q1 2027"},
    "forces": [
      {"label": "Competitor Pressure", "body": "Rivals cutting list price 15%.",        "side": "left",   "tone": "negative"},
      {"label": "Customer WTP",        "body": "NPS 72 supports 8% premium.",            "side": "right",  "tone": "positive"},
      {"label": "Regulatory Ceiling",  "body": "EU limits annual price increases.",       "side": "top",    "tone": "negative"},
      {"label": "Innovation Pipeline", "body": "3 new SKUs justify 18% Tier-1 premium.", "side": "bottom", "tone": "positive"}
    ]
  }
}
```

Key: `forces` — not `nodes`. `side`: `"left"` / `"right"` / `"top"` / `"bottom"`.

### 37. word_cloud

```json
{
  "layout": "word_cloud",
  "content": {
    "words": [
      {"text": "Growth",      "weight": 5},
      {"text": "Flexibility", "weight": 4.5},
      {"text": "Burnout",     "weight": 3},
      {"text": "Recognition", "weight": 4}
    ]
  }
}
```

`weight` 1–5 maps to font size 15–35pt. Up to 30 words.

### 38. pyramid

```json
{
  "layout": "pyramid",
  "content": {
    "orientation": "up",
    "tiers": [
      {"label": "Transactions",      "body": "High volume, low margin."},
      {"label": "Solutions",         "body": "Bundled offer with SLA."},
      {"label": "Strategic Partner", "body": "Co-development, IP sharing."}
    ]
  }
}
```

`tiers` ordered bottom-to-top. `orientation`: `"up"` (wide base) or `"down"` (inverted).

### 39. venn

```json
{
  "layout": "venn",
  "content": {
    "circles": [
      {"label": "Scientific Depth", "body": "40yr R&D, 800 patents"},
      {"label": "Global Reach",     "body": "Sales in 65 countries"},
      {"label": "Digital Platform", "body": "12M patient data points"}
    ],
    "intersection": "Unique Position"
  }
}
```

Key: `intersection` — not `overlap`.

### 40. risk_heatmap

```json
{
  "layout": "risk_heatmap",
  "content": {
    "x_label": "LIKELIHOOD",
    "y_label": "IMPACT",
    "risks": [
      {"label": "Cyber security breach",         "likelihood": 3, "impact": 5},
      {"label": "Key talent attrition",          "likelihood": 4, "impact": 4},
      {"label": "Regulatory enforcement action", "likelihood": 2, "impact": 5}
    ]
  }
}
```

`likelihood` and `impact` each 1–5. Max 15 risks.

### 41. radar_chart

```json
{
  "layout": "radar_chart",
  "content": {
    "axes": ["Scientific Depth", "Brand Trust", "Digital Agility", "Commercial Execution", "Talent Density"],
    "series": [
      {"label": "Merck",        "values": [92, 85, 48, 72, 78]},
      {"label": "Competitor A", "values": [70, 65, 82, 80, 70]}
    ]
  }
}
```

4–8 axes. Values 0–100. Max 4 series.

### 42. pros_cons

```json
{
  "layout": "pros_cons",
  "content": {
    "subject": "ACQUISITION OF TECHVENTURE AG",
    "pros_label": "STRATEGIC CASE FOR",
    "cons_label": "RISKS TO MANAGE",
    "pros": [
      "Cuts 18-month build timeline to 6 months",
      "Brings 45 experienced ML engineers immediately"
    ],
    "cons": [
      "280M price implies 8.4x ARR; above sector median",
      "Cultural integration risk: remote vs office-first"
    ]
  }
}
```

### 43. layered_stack

```json
{
  "layout": "layered_stack",
  "content": {
    "orientation": "vertical",
    "layers": [
      {"icon": "U",  "label": "Experience Layer",     "body": "Web, mobile, and API portals."},
      {"icon": "AI", "label": "Intelligence Layer",    "body": "LLM-powered recommendation and NLP."},
      {"icon": "D",  "label": "Data & Integration",   "body": "Snowflake + Kafka + 40 API connectors."},
      {"icon": "C",  "label": "Cloud Infrastructure", "body": "AWS primary, Azure DR, 99.99% SLA."}
    ]
  }
}
```

### 44. photo_text

```json
{
  "layout": "photo_text",
  "content": {
    "image_path":  "/path/to/campus.jpg",
    "image_label": "Darmstadt R&D Campus",
    "image_side":  "left",
    "title":       "A Campus Built for Collaboration",
    "bullets": [
      "130,000 m2 of lab and collaboration space opened in 2019",
      "4,800 scientists and engineers on one connected campus",
      "Carbon-neutral since 2022"
    ]
  }
}
```

Keys: `image_path`, `image_side`, `title`, `bullets` — not `image: {path, position}` + `text`.

### 45. fishbone

```json
{
  "layout": "fishbone",
  "content": {
    "effect": "Late Deliveries",
    "bones": [
      {"label": "Supplier", "causes": ["Long lead times", "Missing specs", "Single-source risk"]},
      {"label": "Planning", "causes": ["Forecast errors", "No buffer stock"]},
      {"label": "Process",  "causes": ["Manual handoffs", "No SLA tracking"]},
      {"label": "People",   "causes": ["Training gaps", "High turnover"]}
    ]
  }
}
```

Key: `bones` — not `causes`. Max 6 bones, 3 sub-causes each.

### 46. columns (auto-dispatch)

Auto-dispatches to `two_column`, `three_column`, or `four_column` based on count.

```json
{
  "layout": "columns",
  "content": {
    "columns": [
      {"label": "LEVER 1", "tone": "positive", "body": "Price realisation", "items": ["..."]},
      {"label": "LEVER 2", "tone": "neutral",  "body": "Volume growth",     "items": ["..."]},
      {"label": "LEVER 3", "tone": "neutral",  "body": "Cost reduction",    "items": ["..."]}
    ]
  }
}
```

---

## Italic-emphasis on action titles

`action_title` can be a list of `(text, italic_bool)` tuples. Italic runs render in MERCK_YELLOW italic Merck Web.

```json
{
  "action_title": [
    ["Three decisions today ", false],
    ["unlock the Q3 trajectory.", true]
  ]
}
```

Use sparingly — no more than 1 in 4 slides.

---

## Footnotes

`chart_slide` and `waterfall_slide` accept a `footnotes` array:

```json
{
  "content": {
    "chart": { ... },
    "footnotes": [
      [1, "PS includes Reagents and Consumables; excludes Pharma."],
      [2, "DS reported through legacy ERP; reconciliation in progress."]
    ]
  }
}
```

---

## Methodology note

Any content layout accepts `methodology_note` inside `content`. Renders as an italic gray line above the source line.

```json
{
  "layout": "two_column",
  "content": {
    "left": { ... },
    "right": { ... },
    "methodology_note": "MECE structure: WHICH (governance) + WHAT (curation) + HOW IT STAYS TRUE (controls)."
  }
}
```
