import re


def _title_text(value: object) -> str:
    """Extract plain text from an action_title value.

    action_title can be a plain str or a list of (text, italic_bool) tuples
    for rich (italicised) text. This normaliser handles both so length checks
    and regex searches work correctly for both formats.
    """
    if isinstance(value, (list, tuple)):
        return "".join(
            seg[0] if isinstance(seg, (list, tuple)) else str(seg)
            for seg in value
        )
    return str(value) if value is not None else ""


_STRUCTURAL_FUNCTIONS = frozenset({
    "Cover", "Agenda", "Section Divider", "Close",
    "Executive Summary", "Hero Stat", "Pull Quote",
})

# Lowercase-normalized version for case-insensitive matching of LLM-generated page_functions.
# Space-normalised too: "section_divider" → "section divider" → matches "Section Divider".
_STRUCTURAL_NORMS = frozenset(fn.lower() for fn in _STRUCTURAL_FUNCTIONS)

# Layouts that are structurally non-content regardless of page_function value.
# Must stay in sync with _NO_CIRCLE_LAYOUTS in build_from_plan.py.
_STRUCTURAL_LAYOUTS = frozenset({
    "cover", "agenda", "section_divider", "close", "exec_summary", "hero_stat", "pull_quote",
})

_VALID_LAYOUTS = frozenset({
    # Core layouts
    "cover", "exec_summary", "agenda", "section_divider", "close",
    # Column / list layouts
    "two_column", "three_column", "four_column", "columns",
    "vertical_numbered", "label_rows",
    # Chart / data layouts  (canonical names match merck_layouts.py)
    "chart_slide", "waterfall_slide", "2x2_matrix",
    "stat_strip", "donut_chart", "radar_chart", "risk_heatmap",
    # Process / timeline
    "phase_process", "gantt", "milestone_timeline",
    "circular_flow", "arrow_chain", "funnel", "journey_map",
    # Decision / analysis
    "decision_rows", "before_after", "comparison_table", "score_table",
    "pros_cons", "influence_diagram",
    # Organizational
    "hub_spoke", "pillar_detail", "org_chart", "topic_set",
    "status_table", "kpi_dashboard", "icon_grid",
    # Visual / story
    "hero_stat", "stat_strip", "pull_quote", "word_cloud",
    "pyramid", "venn", "layered_stack", "photo_text", "fishbone",
})

_COUNT_WORDS = {
    "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8,
}

_FORBIDDEN_CLASSIFICATIONS = frozenset({"secret", "top secret", "ts/sci"})
_VALID_CLASSIFICATIONS    = frozenset({"public", "internal", "confidential"})
_VALID_COLOR_THEMES       = frozenset({
    "functional", "organic", "plastic", "synthetic", "technical", "electronics",
})

MAX_SLIDES = 150


class ValidationError(ValueError):
    pass


def validate_plan(plan: dict) -> list:
    """
    Validate a slide plan dict.

    Raises ValidationError on hard errors. Returns a list of warning strings
    for soft violations (caller is responsible for printing them).
    """
    warnings = []
    meta = plan.get("meta") or {}
    slides = plan.get("slides", [])

    # Hard: forbidden classification — enforced here so the Python API and
    # the build command cannot bypass the Secret guard that the CLI prompts use.
    classification = str(meta.get("classification", "")).strip().lower()
    if classification in _FORBIDDEN_CLASSIFICATIONS:
        raise ValidationError(
            f"Classification '{meta.get('classification')}' is not permitted. "
            f"This pipeline cannot process Secret-classified content."
        )
    if classification and classification not in _VALID_CLASSIFICATIONS:
        raise ValidationError(
            f"Unknown classification '{meta.get('classification')}'. "
            f"Must be Public, Internal, or Confidential."
        )

    # Soft: unknown color_theme — warn but don't block.
    color_theme = str(meta.get("color_theme") or "").strip().lower()
    if color_theme and color_theme not in _VALID_COLOR_THEMES:
        warnings.append(
            f"Unknown color_theme '{meta.get('color_theme')}'. "
            f"Valid values: {sorted(_VALID_COLOR_THEMES)}. "
            f"Defaulting to 'plastic'."
        )

    # Hard: slide count cap — prevents memory exhaustion from crafted plans.
    if len(slides) > MAX_SLIDES:
        raise ValidationError(
            f"Plan contains {len(slides)} slides; maximum permitted is {MAX_SLIDES}."
        )

    seen_nums = set()
    agenda_chapters = 0
    section_divider_count = 0

    for s in slides:
        layout = s.get("layout", "")
        fn_raw = s.get("page_function") or ""
        # Normalise: lowercase + underscores → spaces so "section_divider" == "Section Divider"
        fn_norm = fn_raw.strip().lower().replace("_", " ")
        is_structural = fn_norm in _STRUCTURAL_NORMS or layout in _STRUCTURAL_LAYOUTS

        # Hard: unknown layout
        if layout not in _VALID_LAYOUTS:
            raise ValidationError(
                f"Slide page={s.get('page')}: unknown layout '{layout}'. "
                f"Valid: {sorted(_VALID_LAYOUTS)}"
            )

        # Hard: duplicate / missing section_number
        n = s.get("section_number")
        if n is not None:
            if n in seen_nums:
                raise ValidationError(f"Duplicate section_number: {n!r}")
            seen_nums.add(n)
        elif not is_structural:
            raise ValidationError(
                f"Slide page={s.get('page')} (page_function={fn_raw!r}) "
                f"is a content slide but has no section_number."
            )

        # Warn: action_title length on non-cover slides (max 80 chars).
        # _title_text handles both str and list-of-(text, italic_bool) formats.
        if layout != "cover" and fn_norm != "cover":
            raw_title = _title_text(s.get("action_title"))
            if raw_title and len(raw_title) > 80:
                warnings.append(
                    f"Slide page={s.get('page')}: action_title is {len(raw_title)} chars (max 80)."
                )

        # Hard: cover title > 60 chars or missing subtitle
        if layout == "cover" or fn_norm == "cover":
            raw = s.get("action_title") or ""
            plain = _title_text(raw)
            title_text = plain.split(";", 1)[0].strip() if ";" in plain else plain
            if len(title_text) > 60:
                raise ValidationError(
                    f"Cover title too long ({len(title_text)} chars, max 60): {title_text!r}"
                )
            has_sub = (";" in plain and plain.split(";", 1)[1].strip()) or \
                      (s.get("content") or {}).get("subtitle")
            if not has_sub:
                raise ValidationError("Cover slide is missing a subtitle.")

        # Hard + warn: takeaway — check both the top-level key and the content dict key
        # independently so a long content.takeaway can't bypass the limit via a short top-level one.
        for tw in (s.get("takeaway"), (s.get("content") or {}).get("takeaway")):
            if isinstance(tw, str):
                if len(tw) > 120:
                    raise ValidationError(
                        f"Slide page={s.get('page')}: takeaway is {len(tw)} chars (max 120)."
                    )
                elif len(tw) >= 90:
                    warnings.append(
                        f"Slide page={s.get('page')}: takeaway is {len(tw)} chars "
                        f"(approaching 120-char limit)."
                    )

        # Warn: title count word vs content item count mismatch
        title = _title_text(s.get("action_title")).lower()
        m = re.search(r"\b(two|three|four|five|six|seven|eight)\b", title)
        if m:
            expected = _COUNT_WORDS[m.group(1)]
            content = s.get("content") or {}
            for key in ("columns", "items", "phases", "steps", "stats", "kpis", "decisions"):
                if key in content and isinstance(content[key], list):
                    actual = len(content[key])
                    if actual != expected:
                        warnings.append(
                            f"Slide page={s.get('page')}: title says '{m.group(1)}' "
                            f"but {key} has {actual} items."
                        )
                    break

        # Collect counts for the post-loop agenda warning
        if layout == "agenda" or fn_norm == "agenda":
            chapters = (s.get("content") or {}).get("chapters") or []
            if len(chapters) > agenda_chapters:
                agenda_chapters = len(chapters)
        if layout == "section_divider" or fn_norm == "section divider":
            section_divider_count += 1

    # Warn: agenda chapter count vs section-divider slide count
    # (chapters ≈ sections; a mismatch means the agenda promises a section that has no divider)
    if agenda_chapters and section_divider_count and agenda_chapters != section_divider_count:
        warnings.append(
            f"Agenda has {agenda_chapters} chapters but there are "
            f"{section_divider_count} section divider slides."
        )

    return warnings
