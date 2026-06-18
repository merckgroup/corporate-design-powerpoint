import re

_STRUCTURAL_FUNCTIONS = frozenset({
    "Cover", "Agenda", "Section Divider", "Close",
    "Hero Stat", "Pull Quote",
})

# Lowercase-normalized version for case-insensitive matching of LLM-generated page_functions.
# Space-normalised too: "section_divider" → "section divider" → matches "Section Divider".
_STRUCTURAL_NORMS = frozenset(fn.lower() for fn in _STRUCTURAL_FUNCTIONS)

# Layouts that are structurally non-content regardless of page_function value.
_STRUCTURAL_LAYOUTS = frozenset({
    "cover", "agenda", "section_divider", "close", "hero_stat", "pull_quote",
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


class ValidationError(ValueError):
    pass


def validate_plan(plan: dict) -> list:
    """
    Validate a slide plan dict.

    Raises ValidationError on hard errors. Returns a list of warning strings
    for soft violations (caller is responsible for printing them).
    """
    warnings = []
    slides = plan.get("slides", [])

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

        # Hard: cover title > 60 chars or missing subtitle
        if layout == "cover" or fn_norm == "cover":
            raw = s.get("action_title") or ""
            title_text = raw.split(";", 1)[0].strip() if ";" in raw else raw
            if len(title_text) > 60:
                raise ValidationError(
                    f"Cover title too long ({len(title_text)} chars, max 60): {title_text!r}"
                )
            has_sub = (";" in raw and raw.split(";", 1)[1].strip()) or \
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
        title = (s.get("action_title") or "").lower()
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
