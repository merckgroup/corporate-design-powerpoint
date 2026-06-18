# Merck Presentation Design Guidelines
> **Purpose:** Standalone reference for an AI agent generating Merck Corporate Design-compliant `.pptx` presentations. No other documents or tools are needed to apply these rules. Sources: *Helpful Advice for PowerPoint (empower9™, Nov 2023)*, *Merck PowerPoint Guideline (Oct 2015)*, and `merck_layouts.py` (authoritative color/style definitions).

---

## Table of Contents
1. [Regional & Brand Identity Rules](#1-regional--brand-identity-rules)
2. [Slide Format & Structure](#2-slide-format--structure)
3. [Three Visual Styles](#3-three-visual-styles)
4. [Six Color Themes](#4-six-color-themes)
5. [Color Palette — Authoritative Hex Codes](#5-color-palette--authoritative-hex-codes)
6. [Color Usage Rules](#6-color-usage-rules)
7. [Typography](#7-typography)
8. [Slide Layouts](#8-slide-layouts)
9. [Section Numbers](#9-section-numbers)
10. [Corporate Design Elements](#10-corporate-design-elements)
11. [Images & Photography](#11-images--photography)
12. [Tables](#12-tables)
13. [Charts & Diagrams](#13-charts--diagrams)
14. [Shapes](#14-shapes)
15. [Accessibility](#15-accessibility)
16. [Pre-Publish Checklist](#16-pre-publish-checklist)
17. [Support & Contacts](#17-support--contacts)

---

## 1. Regional & Brand Identity Rules

### 1.1 Name & Logo by Region
- **Global (all countries except USA & Canada):** Company name = **Merck**. Use EU master templates.
- **USA & Canada only:** Different brand names per business unit:
  - EMD Electronics
  - MilliporeSigma
  - EMD Serono
  Use USA/Canada master templates — they carry different legal disclaimers. **Never use an EU template for a USA audience.**
- **Asian markets (Japan, China, Korea, Taiwan, Hong Kong):** Same Merck branding, different font (see [§7.3](#73-asian-markets)).

### 1.2 Template Files by Region
| Region | Master template file |
|---|---|
| EU — default (plastic theme) | `EU_Merck_Themed.pptx` |
| EU — functional theme | `EU_Merck_Functional.pptx` |
| EU — organic theme | `EU_Merck_Organic.pptx` |
| EU — synthetic theme | `EU_Merck_Synthetic.pptx` |
| EU — technical theme | `EU_Merck_Technical.pptx` |
| EU — electronics theme | `EU_Merck_Electronics.pptx` |
| USA / Canada | `USA_Merck_Themed_Base_v1.pptx` |

If a theme-specific file is missing, fall back to the region default (`EU_Merck_Themed.pptx` or `USA_Merck_Themed_Base_v1.pptx`).

### 1.3 Slide Aspect Ratio
Standard format for all Merck presentations: **16:9**.

---

## 2. Slide Format & Structure

### 2.1 Every Content Slide Has Three Fixed Zones

```
┌──────────────────────────────────────────────┐
│  Action title (1 line)                       │
│  Slide title (1–2 lines)                     │
├──────────────────────────────────────────────┤
│                                              │
│        WORKING AREA  (content)               │
│   1–4 columns of text, images, charts, etc. │
│                                              │
├──────────────────────────────────────────────┤
│  footer text | date          [logo]  [page#] │
└──────────────────────────────────────────────┘
```

- **Action title:** Short punchy takeaway, max 1 line. → Verdana 22 pt, Rich Purple (`#503291`).
- **Slide title:** 1–2 lines. → Verdana Bold 22 pt, Rich Purple (`#503291`).
- **Footer:** Set once globally via Presentation Settings. Never edit per-slide.
- The working area must contain all content. Nothing overflows into the title or footer zones.

### 2.2 Gridlines & Guides
- Grid spacing: **0.2 cm**. Arrow-key nudge = 0.2 cm per keystroke.
- Fixed marks on slide edges define column splits and the working-area boundary.

---

## 3. Three Visual Styles

Every slide in a deck inherits one of three visual styles. The style determines the full color role mapping used by every layout function. The deck-level default is set in `meta.deck_style`; individual slides may override with `style`.

### 3.1 Style: `merck_executive`
White background, Purple primary, formal tone. **Auto-promoted** to this style for any slide whose `category` is one of:
- "Executive Summary"
- "Recommendation"
- "Decision Request"
- "Risk"
- "Tradeoff"

| Role | Color name | Hex |
|---|---|---|
| `bg` (background) | WHITE | `#FFFFFF` |
| `ink` (primary text) | INK_DARK | `#1A1626` |
| `ink_2` (secondary text, sources) | INK_GRAY | `#555D6E` |
| `ink_3` (rules/borders) | LIGHT_GRAY | `#E0E0E0` |
| `accent` (primary brand) | MERCK_PURPLE | `#503291` |
| `accent_2` (dark depth) | PURPLE_DEEP | `#3A2468` |
| `accent_3` (cool contrast) | LY_CYAN | `#2DBECD` |
| `highlight` (emphasis) | MERCK_GOLD (Pink) | `#EB3C96` |
| `hot` (bright call-out) | MERCK_YELLOW | `#FFC832` |
| `rule` (dividers) | LIGHT_GRAY | `#E0E0E0` |
| `panel` (card background) | PANEL_LIGHT | `#F4F2F8` |
| `muted` (subdued element) | PURPLE_MUTED | `#7D74A0` |
| `good` (positive/green) | GOOD_GREEN | `#149B5F` |
| `warn` (warning/amber) | MERCK_YELLOW | `#FFC832` |
| `bad` (negative/red) | BAD_RED | `#E61E50` |
| `lime` (plan/operating) | OP_LIME | `#A5CD50` |

### 3.2 Style: `merck_corporate`
Identical palette to `merck_executive`. General-purpose business presentations.

### 3.3 Style: `merck_storytelling`
Dark/inverted: Purple background, white text. Used for impactful narrative slides.

| Role | Color name | Hex |
|---|---|---|
| `bg` | MERCK_PURPLE | `#503291` |
| `ink` | WHITE | `#FFFFFF` |
| `ink_2` | PANEL_LIGHT | `#F4F2F8` |
| `ink_3` | PURPLE_MUTED | `#7D74A0` |
| `accent` | MERCK_GOLD (Pink) | `#EB3C96` |
| `accent_2` | MERCK_YELLOW | `#FFC832` |
| `accent_3` | LY_CYAN | `#2DBECD` |
| `highlight` | MERCK_GOLD (Pink) | `#EB3C96` |
| `hot` | MERCK_YELLOW | `#FFC832` |
| `rule` | PURPLE_MUTED | `#7D74A0` |
| `panel` | PANEL_LIGHT | `#F4F2F8` |
| `muted` | PURPLE_MUTED | `#7D74A0` |
| `good` | GOOD_GREEN | `#149B5F` |
| `warn` | MERCK_YELLOW | `#FFC832` |
| `bad` | BAD_RED | `#E61E50` |
| `lime` | OP_LIME | `#A5CD50` |

### 3.4 Style Resolution Order
For any given slide, the style is resolved as:
1. Use `slide.style` if it is not `"inherit"`.
2. Otherwise use `meta.deck_style`.
3. Auto-promote to `merck_executive` if `slide.category` matches the list in §3.1, regardless of steps 1–2.

---

## 4. Six Color Themes

The color theme (`meta.color_theme`) controls the cover/divider background color and the accent role used throughout the deck. It does **not** change the three style palettes above — instead, it selects which template `.pptx` file is opened and applies a color patch to the brand accent slots.

| `color_theme` | Cover BG | Accent | Character | Dark bg? |
|---|---|---|---|---|
| `plastic` | Light green `#A5CD50` | Pink `#EB3C96` | Default / general | No |
| `functional` | Light green `#A5CD50` | Teal `#2DBECD` | Life science, organic cells | No |
| `organic` | Cream `#FFDCB9` | Red `#E61E50` | Healthcare, patient focus | No |
| `synthetic` | Violet `#503291` | Yellow `#FFC832` | Industrial, chemistry | Yes |
| `technical` | Cream `#FFDCB9` | Teal `#2DBECD` | Engineering, IT, angular | No |
| `electronics` | Violet `#503291` | Yellow `#FFC832` | EMD Electronics (photo cover) | Yes |

**Dark-background themes** (`synthetic`, `electronics`): use `ink = WHITE (#FFFFFF)` and `ink_2 = PANEL_LIGHT (#F4F2F8)` for text on these slides.

**Electronics theme:** The cover layout is "Title with picture" — the image placeholder is intentionally left empty; the user fills it in PowerPoint after generation.

### 4.1 Per-Theme Palette Overrides
The following table shows the `accent` and other key roles per theme (roles not listed default to the values in §3.1 executive/corporate palette):

| Theme | `accent` | `accent_2` | `accent_3` | `highlight` | `hot` | `muted` | `panel` |
|---|---|---|---|---|---|---|---|
| `plastic` | `#503291` | `#3A2468` | `#2DBECD` | `#EB3C96` | `#A5CD50` | `#7D74A0` | `#F4F2F8` |
| `functional` | `#2DBECD` | `#503291` | `#A5CD50` | `#2DBECD` | `#A5CD50` | `#96D7D2` | `#F0F9FA` |
| `organic` | `#E61E50` | `#503291` | `#FFDCB9` | `#E61E50` | `#FFDCB9` | `#E1C3CD` | `#FDF5ED` |
| `synthetic` | `#FFC832` | `#2DBECD` | `#A5CD50` | `#FFC832` | `#2DBECD` | `#7D74A0` | `#3F2870` |
| `technical` | `#2DBECD` | `#503291` | `#FFDCB9` | `#2DBECD` | `#FFDCB9` | `#96D7D2` | `#F0F9FA` |
| `electronics` | `#FFC832` | `#2DBECD` | `#A5CD50` | `#FFC832` | `#2DBECD` | `#7D74A0` | `#3F2870` |

---

## 5. Color Palette — Authoritative Hex Codes

These are the canonical color constants used by the rendering engine. Use these values for all fills, strokes, and text colors. **Do not invent or approximate colors.**

### 5.1 Core Brand Colors
| Name | Hex | Role |
|---|---|---|
| MERCK_PURPLE | `#503291` | Primary brand / accent / headings |
| MERCK_BLUE | `#0F69AF` | Secondary brand / cool contrast |
| MERCK_GOLD (Pink) | `#EB3C96` | Highlight / warm emphasis |
| PURPLE_DEEP | `#3A2468` | Dark depth / footer backgrounds |
| PURPLE_MUTED | `#7D74A0` | Subdued elements / separators |
| MERCK_YELLOW | `#FFC832` | Hot callout / warning |
| MERCK_AQUA | `#96D7D2` | Sensitive blue / muted accent |
| BAD_RED | `#E61E50` | Negative / alert |
| GOOD_GREEN | `#149B5F` | Positive / success |

### 5.2 Theme Accent Colors
| Name | Hex | Used in |
|---|---|---|
| LY_CYAN | `#2DBECD` | functional, technical, accent_3 in executive/plastic |
| OP_LIME | `#A5CD50` | plastic/functional cover bg, lime role |
| _MC_PINK | `#EB3C96` | plastic highlight (same as MERCK_GOLD) |
| _MC_PALEYELLOW | `#FFDCB9` | organic/technical warm bg tint |
| _MC_PALEPINK | `#E1C3CD` | organic muted |
| _MC_PALEBLUE | `#96D7D2` | functional/technical muted (same as MERCK_AQUA) |
| _MC_PALEGREEN | `#B4DC96` | template internal bg (do not use as a design choice) |

### 5.3 Neutral / Text Colors
| Name | Hex | Role |
|---|---|---|
| WHITE | `#FFFFFF` | Background (light themes), text on dark themes |
| INK_DARK | `#1A1626` | Primary body text on white backgrounds |
| INK_GRAY | `#555D6E` | Secondary text, source lines, footnotes |
| LIGHT_GRAY | `#E0E0E0` | Rules, dividers, borders |
| PANEL_LIGHT | `#F4F2F8` | Card/panel background on white slides |

### 5.4 Traffic Light / Status Colors
Always use this exact triad for RAG/status indicators:
| Signal | Color | Hex |
|---|---|---|
| Good / Positive | GOOD_GREEN | `#149B5F` |
| Warning / Amber | MERCK_YELLOW | `#FFC832` |
| Bad / Negative | BAD_RED | `#E61E50` |

### 5.5 Chart Color Sequence
When cycling through data series in a chart, use this ordered palette:
| # | Hex | Description |
|---|---|---|
| 1 | `#DE7A21` | orange-600 |
| 2 | `#12879D` | teal-600 |
| 3 | `#184251` | cyan-900 |
| 4 | `#FEC731` | amber-400 |
| 5 | `#F7A216` | amber-500 |
| 6 | `#0E69AF` | blue-700 |
| 7 | `#E61E50` | red-600 |
| 8 | `#96D7D2` | sensitive-blue |
| 9 | `#139A5F` | green-500 |
| 10 | `#FFDCB9` | sensitive-yellow |
| 11 | `#9D80E6` | violet-400 |
| 12 | `#8C2235` | red-800 |

---

## 6. Color Usage Rules

### 6.1 Color Matrix (Brand Guidelines)
The 12 brand colors are organized in a 3×4 matrix. **Always mix from different rows AND different columns** for combinations.

| | Passionate | Cool | Positive | Warm |
|---|---|---|---|---|
| **Rich** | Rich Purple `#503291` | Rich Blue `#0F69AF` | Rich Green `#149B5F` | Rich Red `#E61E50` |
| **Vibrant** | Vibrant Magenta `#EB3C96` | Vibrant Cyan `#2DBECD` | Vibrant Green `#A5CD50` | Vibrant Yellow `#FFC832` |
| **Sensitive** | Sensitive Pink `#E1C3CD` | Sensitive Blue `#96D7D2` | Sensitive Green `#B4DC96` | Sensitive Yellow `#FFDCB9` |

### 6.2 Combination Rules
| Rule | Detail |
|---|---|
| **Group size** | Colors work best in **groups of 3**. Fewer = monotonous; more than 3 = cluttered (charts/infographics excepted) |
| **Matrix rule** | Always mix colors from different rows AND different columns |
| **Gradients** | **Never** — neither linear nor radial |
| **Tints/shades** | **Never** in design elements. Only in charts and tables |
| **Sensitive colors** | Only for backgrounds, never as dominant brand elements |
| **Logo** | Only Rich or Vibrant colors. Never Sensitive. Never black/gray/white |
| **High-contrast same-family** | Avoid (e.g., don't combine Rich Purple + Vibrant Purple + Sensitive Purple) |
| **Red + Green** | Avoid — red-green colorblindness issue |
| **Title vs. logo color** | The title text color must never match the logo color on the same slide |

---

## 7. Typography

### 7.1 Fonts
| Region | Body / Headings | Display (titles/dividers only) |
|---|---|---|
| Global (EU + USA/Canada) | **Verdana** Regular + Bold | **Merck** (custom geometric font) |
| Asia (JP, CN, KR, TW, HK) | **Noto** Regular + Bold | **Noto Black** (replaces Merck display font) |

The Merck display font is **only** for large headlines, title slides, and divider slides. **Never** use it for body text.

### 7.2 Font Size Hierarchy (Authoritative — 2023 values)
| Element | Font | Size | Color |
|---|---|---|---|
| Action title | Verdana Regular | **22 pt** | Rich Purple `#503291` |
| Slide heading | Verdana **Bold** | **22 pt** | Rich Purple `#503291` |
| Subheading | Verdana **Bold** | Same as body | Any corporate color |
| Standard body text | Verdana Regular | **16 pt** | INK_DARK `#1A1626` |
| Minimum body text | Verdana Regular | **10 pt** | INK_DARK `#1A1626` |
| Diagram / margin headings | Verdana Regular | **12 pt** | Rich Purple `#503291` |
| Source / reference lines | Verdana Regular | **8 pt** | INK_GRAY `#555D6E` |
| Footer | Verdana Regular | **8 pt** | Rich Purple `#503291` |

- Use a **maximum of 3 different font sizes** in the content area across the entire presentation.
- When presenting on a beamer/projector: increase sizes from the above defaults.
- PowerPoint is **not** a print-product tool — do not design for brochures/flyers.

### 7.3 Asian Markets
Use **Noto** instead of Verdana. Noto Regular/Bold for headings and body; Noto Black replaces the Merck display font on title/divider slides. Do **not** embed Noto (file size too large) — always send Asian-font decks as **PDF**.

### 7.4 Merck Display Font Rules
- Only for large display headlines, titles, and dividers — never body text.
- **Never** color it in black (`#000000`), gray, or white (`#FFFFFF`).
- Single words within a Merck font headline must **not** be colored differently from the rest of the phrase.
- Embed Merck font in `.pptx` files before sharing (File → Options → Save → "Embed fonts in the file" — Windows only; Mac cannot display embedded fonts).

### 7.5 Text Alignment
| Context | Alignment |
|---|---|
| Continuous text and bullet lists | **Left** |
| Short labels in boxes / graphic elements / table cells | **Center** (exception: never center continuous text or lists) |
| Numbers in tables | **Right** |

### 7.6 Paragraph & Line Spacing
- **First-level paragraph:** 6 pt before, 3 pt after.
- **All other levels:** 3 pt before, 3 pt after.
- **Line spacing:** Multiple 1.05.

### 7.7 Bullet Level Hierarchy
```
First level        (no bullet — plain text)
  • Second level   (filled bullet)
    – Third level  (en-dash)
      – Fourth level
        – Fifth level
```

---

## 8. Slide Layouts

### 8.1 Available Layout Keys (44 total)
The following keys map directly to `build_*` functions in the rendering engine. Choose the layout that best matches the content structure:

**Structural slides (no section_number):**
| Key | Purpose |
|---|---|
| `cover` | Opening slide: title, subtitle, authors, key messages |
| `agenda` | Table of contents listing chapters |
| `section_divider` | Chapter separator with chapter number and title |
| `close` | Final / thank-you slide |
| `exec_summary` | Executive summary with key messages (up to 5) |

**Content: Text & Lists**
| Key | Purpose |
|---|---|
| `two_column` | Left + right column, each with header and bullet items |
| `three_column` | Three columns of equal content |
| `four_column` | Four columns |
| `topic_set` | Multiple independent topic blocks |
| `label_rows` | Labeled rows of content (like a vertical table) |
| `vertical_numbered` | Numbered list of steps or points |
| `pull_quote` | Large quote or statement |
| `pros_cons` | Two-column pros vs. cons comparison |

**Content: Process & Flow**
| Key | Purpose |
|---|---|
| `phase_process` | Horizontal phase/stage flow (up to 5 phases, with optional arrows) |
| `arrow_chain` | Linear arrow-connected steps |
| `milestone_timeline` | Timeline with milestone markers |
| `gantt` | Gantt chart with rows (labels) and quarter columns |
| `waterfall` | Waterfall flow / sequential steps |
| `decision_rows` | Decision/recommendation rows with tone/owner (up to 5) |
| `circular_flow` | Circular process diagram |
| `funnel` | Funnel stages |
| `layered_stack` | Stacked layers (architecture, technology stack) |
| `fishbone` | Cause-and-effect / Ishikawa diagram |

**Content: Data & Charts**
| Key | Purpose |
|---|---|
| `chart` | Chart with type, data, and optional callouts |
| `donut_chart` | Donut/pie chart |
| `kpi_dashboard` | KPI metrics dashboard |
| `hero_stat` | Single prominent statistic with context |
| `stat_strip` | Row of 3–4 statistics |
| `score_table` | Scoring matrix / assessment table |
| `comparison_table` | Feature comparison table |
| `status_table` | Status/RAG table with rows and columns |
| `risk_heatmap` | 2D risk likelihood × impact matrix |
| `radar_chart` | Radar/spider chart |
| `word_cloud` | Visual word cloud |

**Content: Layouts & Visual**
| Key | Purpose |
|---|---|
| `matrix_2x2` | 2×2 strategy/priority matrix |
| `before_after` | Side-by-side before vs. after comparison |
| `hub_spoke` | Central concept with radiating topics |
| `pillar_detail` | Pillars (columns) each with sub-detail |
| `icon_grid` | Grid of icons with labels and descriptions |
| `journey_map` | Customer/process journey with stages and experience markers |
| `influence_diagram` | Flow of influences or relationships |
| `pyramid` | Hierarchical pyramid |
| `venn` | 2 or 3-circle Venn / overlap diagram |
| `org_chart` | Organizational chart |
| `photo_text` | Large image with accompanying text |

### 8.2 Key Content Schemas (Most Common)
| Layout | Required content fields |
|---|---|
| `cover` | `subtitle`, `authors: [str]`, `key_messages: [str]` |
| `exec_summary` | `key_messages: [{label, body}]` — max 5 items |
| `agenda` | `chapters: [{number, title}]` |
| `section_divider` | `number` (int), `title` (str) |
| `two_column` | `left: {header, items:[str]}`, `right: {header, items:[str]}` |
| `chart` | `chart: {type, data: {categories, series:[{name,values}]}}`, optional `callouts` |
| `decision_rows` | `decisions: [{tone, number, owner, text}]` — max 5 |
| `hero_stat` | `stat: {value, label}`, `context: str` |
| `phase_process` | `phases: [{label, title, body, highlighted}]` — max 5; `show_arrows: bool` |
| `gantt` | `rows: [{label}]`, `quarters: [str]` |

### 8.3 Placeholder Rules
| Placeholder type | Content | Rule |
|---|---|---|
| Text placeholder | Bullet/body text | Always use the template's bullet level hierarchy |
| Content placeholder | Diagram, Table, SmartArt, Video | Do NOT insert images via this placeholder type |
| Image placeholder | Photos and graphics only | Auto-fills and auto-crops; never distorts |

---

## 9. Section Numbers

`section_number` must be **`null`** for the following slide types:
- `cover`
- `agenda`
- `section_divider`
- `close`

`section_number` must be a **unique sequential integer** on all other content slides, starting at 1 and incrementing across the entire deck without gaps or repeats.

---

## 10. Corporate Design Elements

### 10.1 Cells
Organic, filled color shapes — a core Merck brand element.
- Always **filled** (never used as outlines/hollow).
- Exist on two levels: **underlying** (background) and **overlaying** (foreground).
- **Same-level cells = same color.** Overlapping / nested cells = different colors.
- Cells inside cells continue across into neighboring cells.
- **One cell type per slide or item** — never mix different cell shapes on the same slide.

### 10.2 Strings
Illustrated loop/path elements used as emphasis.
- **Never** use as an image border.
- **Never** fill with solid color (like a text box).
- **One string shape per slide or item** maximum.

### 10.3 Mercrobes
Stylized 3D micro-organism shapes expressing exploration.
- Always use corporate palette colors — **never** white, grayscale, or black.
- **Never** distort the basic shape.
- **Never** add shadow, glow, mirror, or outline effects.
- **Never** combine Mercrobes with Cells on the same slide.
- **Never** combine Mercrobes with Strings on the same slide.
- **Never** mix different Mercrobe types in one view — same type, same color per slide.
- A Mercrobe must not look like a pathogen (e.g., a virus).
- Mercrobes must not be placed in photos to replace real objects.

### 10.4 Vibrant M
Stylized Merck M — the global brand identifier. Two variants:
- **Cell-filled M** — corporate shapes fill the interior.
- **String-outline M** — drawn as a string line.

**Rule:** When cell shapes are used on a slide, always use the **string-based Vibrant M**.
**Never** insert images into the Vibrant M. **Never** combine it directly with the Merck logo.

### 10.5 Logo
- **Never** rotate, compress, stretch, or distort.
- **Never** add shadow, 3D, outline, or pattern effects.
- **Never** change letter spacing.
- **Never** use black/gray/white or Sensitive colors for the logo.
- **Never** use multiple different colors within a single logo.
- The logo **may** be re-colored in Rich or Vibrant colors. If re-colored, apply the same color consistently across the **entire** presentation.
- Title text color must **never** match the logo color.

---

## 11. Images & Photography

### 11.1 Style Principles (What to Do)
- **Colorful** — vivid, saturated images.
- **Focused** — subject-led composition.
- **Authentic** — real or convincingly real.
- **Emotional** — people in focus, transporting emotions.
- **Extraordinary perspectives** — unusual angles.
- **Minimalist** — "less is more."

### 11.2 What to Avoid
- Black-and-white or monochrome images (exception: historical images).
- Collages / multiple images dissolved together.
- Obvious digital manipulation.
- Unrealistic montages.
- Cartoons, ClipArt, or emojis.
- "Screwball" / abstract expressive style.
- Clichéd posed groups or pure stock-photo look.
- Non-photorealistic or graphically altered images.

### 11.3 Inserting & Handling Images
- Insert via the image **placeholder** icon (links to corporate MAM via empower9™).
- For non-library images: select the placeholder → Insert → Pictures.
- Images auto-fill the placeholder and are cropped to fit — **never distorted**.
- To replace: delete existing image, insert new one. Do **not** use "Change Picture" (distortion risk).
- Resize using **corner** drag points only. **Edge** drag points distort.
- Use Crop tool (Picture Tools → Format) to select a section.

### 11.4 Image Compression
- Screen/projector: 150 ppi.
- Print: 220 ppi.
- Use "Compress Pictures" to reduce file size, checking "Delete cropped areas."

---

## 12. Tables

- Use the **Merck Table Template** (from the empower9™ Library or Table Tools gallery).
- Table style: header row in corporate color, banded rows.
- Text in cells: numbers right-aligned, labels left-aligned.
- Apply to existing table: Table Tools → Design → Gallery → "Merck table template."
- Tints of brand colors are acceptable inside table cells (only exception to the no-tints rule).
- **Do not embed Excel tables** — this copies the entire workbook into the file, increasing size and exposing hidden data.

---

## 13. Charts & Diagrams

### 13.1 Formatting Rules
| Element | Font | Size | Color |
|---|---|---|---|
| Value axis | Verdana Regular | 10 pt | INK_DARK `#1A1626`; no axis line, no gridlines |
| Category axis | Verdana Bold | 12 pt | Rich Purple `#503291`; axis line 1 pt |
| Data labels | Verdana Regular | 10 pt | INK_DARK `#1A1626` |
| Legend / key | Verdana Regular | 12 pt | Rich Purple `#503291` |
| Emphasis / callout | Merck display font | 24 pt | Any Rich or Vibrant color |

- Remove all gridlines and chart borders.
- Use only CD-compliant colors.

### 13.2 Chart Color Series Sequence
Cycle through the 12-color chart palette in order (see [§5.5](#55-chart-color-sequence)). Start at color 1 for the first series, 2 for second, etc.

### 13.3 Diagram Color Sequence (Simple Charts)
For simple bar/column charts with few series, use this preferred order:
1. MERCK_PURPLE `#503291`
2. MERCK_BLUE `#0F69AF`
3. GOOD_GREEN `#149B5F`
4. MERCK_YELLOW `#FFC832`
5. MERCK_GOLD / Pink `#EB3C96`
6. LY_CYAN `#2DBECD`

---

## 14. Shapes

- **Prefer rounded shapes** — use rounded rectangles instead of sharp rectangles.
- For non-natively-rounded shapes (e.g., triangles): Format Shape → Line: solid, ≥6 pt, same fill color → Cap type: Round, Join type: Round.
- Block arrows cannot be rounded.
- **Never use:**
  - Drop shadows, reflections, or glow effects.
  - 3D effects.
  - Color gradients or pattern fills.
  - Shading effects.
  - Online graphics or ClipArt.
- Shapes must be solid-filled (full-area fill), not hollow outlines.

---

## 15. Accessibility

| Step | Rule |
|---|---|
| 1. General design | Strong contrast between text and background; minimum 16 pt font; sufficient white space; use CD color combinations |
| 2. Slide titles | Every slide must have a clear title in a heading placeholder (screen readers navigate by slide titles) |
| 3. Reading order | Arrange elements in intended reading order — screen readers follow the order objects were added, not visual order |
| 4. Alt text | Provide descriptive alternative text for all images, charts, tables, and fixed graphics |
| 5. Mark decorative | Non-informative objects (logo, cell shapes, strings, markers) must be marked as "decorative" so screen readers skip them |
| 6. Check before publish | Run Review → Check Accessibility; not every warning requires action |

---

## 16. Pre-Publish Checklist

### Regional & Template
- [ ] Correct regional master applied (EU vs. USA/Canada).
- [ ] `color_theme` matches the selected template file.
- [ ] `meta.region` is `"EU"` or `"USA"` and matches the template.

### Style & Auto-Promotion
- [ ] `meta.deck_style` is one of `merck_executive`, `merck_corporate`, `merck_storytelling`.
- [ ] Slides with categories "Executive Summary", "Recommendation", "Decision Request", "Risk", or "Tradeoff" are set or will auto-promote to `merck_executive`.

### Section Numbers
- [ ] `section_number` is `null` for: cover, agenda, section_divider, close.
- [ ] All other slides have unique, sequential integers with no gaps.

### Typography
- [ ] Maximum 3 font sizes in content area.
- [ ] Merck display font only on title/divider slides — not in body.
- [ ] Merck display font not in black, gray, or white.
- [ ] Fonts embedded before sharing open `.pptx` (Windows only; send PDF for Asian fonts).

### Colors
- [ ] Only CD-compliant hex codes used (from §5).
- [ ] No gradients, tints (except charts/tables), or standard PowerPoint palette colors.
- [ ] Color combinations follow the matrix rule (different rows AND columns).
- [ ] Logo not in Sensitive colors; title color ≠ logo color.
- [ ] Traffic lights use exactly: Good=`#149B5F`, Warn=`#FFC832`, Bad=`#E61E50`.

### Brand Elements
- [ ] No 3D effects, shadows, reflections, or glow on any shape.
- [ ] No ClipArt, online graphics, or emojis.
- [ ] Only one cell type per slide.
- [ ] Only one string per slide.
- [ ] Mercrobes same type and color per slide; not combined with cells or strings.

### Charts & Tables
- [ ] Chart series colors follow §5.5 sequence.
- [ ] No chart gridlines or borders.
- [ ] Tables use Merck table template.
- [ ] No embedded Excel objects.

### Content Structure
- [ ] All content within the working area boundaries.
- [ ] Footer set globally (not per-slide).
- [ ] Layouts chosen match content type (see §8.1).
- [ ] `section_number` convention followed (see §9).

### Accessibility
- [ ] Every slide has a clear title.
- [ ] Alt text on all images, charts, and tables.
- [ ] Decorative elements marked accordingly.

---

## 17. Support & Contacts

| Need | Contact |
|---|---|
| Branding & design questions | branding@merckgroup.com |
| 4:3 master template request | branding@merckgroup.com |
| Brand identity principles / Digital Brand Hub | https://brandhub.merckgroup.com |
| Technical IT support | IT ticket via ServiceNow: https://fire.service-now.com/MerckPortal/ |

---

*Authoritative source for color values: `merck_pptx/merck_layouts.py` constants. Authoritative source for design rules: Merck PowerPoint Guidelines (Nov 2023 empower9™ edition, Oct 2015 edition). Where sources conflict, the 2023 source takes precedence.*
