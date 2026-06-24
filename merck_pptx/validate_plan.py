import re
import zipfile


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
    "donut_chart", "radar_chart", "risk_heatmap",
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
    "key_question",
    # Process extensions
    "road_to_success",
    # Science layouts (merck_science style)
    "figure_panel", "methods_box", "sar_table", "multi_chart",
})

_COUNT_WORDS = {
    "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8,
}

_VALID_CHROME_FLAGS = frozenset({
    "progress_bar", "section_circles", "takeaway_bands",
    "footer_breadcrumb", "classification_badge",
})

_FORBIDDEN_CLASSIFICATIONS = frozenset({"secret", "top secret", "ts/sci"})
_VALID_CLASSIFICATIONS    = frozenset({"public", "internal", "confidential"})
_VALID_COLOR_THEMES       = frozenset({
    "functional", "organic", "plastic", "synthetic", "technical", "electronics",
})
_VALID_DIVISIONS          = frozenset({
    "merck", "emd_serono", "emd_electronics", "millipore_sigma", "merck_asia", "usa",
})
_VALID_REGIONS            = frozenset({"eu", "usa"})
_VALID_DECK_STYLES        = frozenset({
    "merck_executive", "merck_corporate", "merck_storytelling", "merck_science",
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

    # Hard: region must be EU or USA (legal compliance — wrong template = wrong disclaimer).
    region_raw = meta.get("region")
    if region_raw is not None:
        region_norm = str(region_raw).strip().lower()
        # Accept common aliases that build_from_plan normalises anyway.
        if region_norm in ("us", "canada", "north america"):
            region_norm = "usa"
        if region_norm not in _VALID_REGIONS:
            raise ValidationError(
                f"meta.region '{region_raw}' is not valid. Must be 'EU' or 'USA'. "
                f"Region controls which legal template is loaded — set this correctly."
            )

    # Hard: deck_style must be one of the three CD-defined styles.
    deck_style = meta.get("deck_style")
    if deck_style is not None:
        if str(deck_style).strip().lower() not in _VALID_DECK_STYLES:
            raise ValidationError(
                f"meta.deck_style '{deck_style}' is not valid. "
                f"Must be one of: {sorted(_VALID_DECK_STYLES)}."
            )

    # Soft: unknown color_theme — warn but don't block.
    color_theme = str(meta.get("color_theme") or "").strip().lower()
    if color_theme and color_theme not in _VALID_COLOR_THEMES:
        warnings.append(
            f"Unknown color_theme '{meta.get('color_theme')}'. "
            f"Valid values: {sorted(_VALID_COLOR_THEMES)}. "
            f"Defaulting to 'plastic'."
        )

    # Soft: unknown chrome flags — a typo silently disables the feature.
    chrome = meta.get("chrome")
    if isinstance(chrome, dict):
        unknown_flags = set(chrome.keys()) - _VALID_CHROME_FLAGS
        if unknown_flags:
            warnings.append(
                f"Unknown chrome flag(s): {sorted(unknown_flags)}. "
                f"Valid flags: {sorted(_VALID_CHROME_FLAGS)}. "
                f"Typos silently disable the feature."
            )

    # Soft: unknown division — warn but don't block.
    division = str(meta.get("division") or "").strip().lower()
    if division and division not in _VALID_DIVISIONS:
        warnings.append(
            f"Unknown division '{meta.get('division')}'. "
            f"Valid values: {sorted(_VALID_DIVISIONS)}. "
            f"Defaulting to 'merck'."
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

        # Warn: action_title length on non-cover slides (max 120 chars).
        # _title_text handles both str and list-of-(text, italic_bool) formats.
        if layout != "cover" and fn_norm != "cover":
            raw_title = _title_text(s.get("action_title"))
            if raw_title and len(raw_title) > 120:
                warnings.append(
                    f"Slide page={s.get('page')}: action_title is {len(raw_title)} chars (max 120)."
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
                      s.get("subtitle") or \
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

    # Hard: section_numbers must be a gapless sequence 1..N.
    # Non-sequential numbers break the agenda hyperlinks and the visible circle numbering.
    if seen_nums:
        sorted_nums = sorted(seen_nums)
        expected    = list(range(sorted_nums[0], sorted_nums[0] + len(sorted_nums)))
        if sorted_nums != expected:
            missing = sorted(set(expected) - seen_nums)
            raise ValidationError(
                f"section_number values are not sequential. "
                f"Found: {sorted_nums}. "
                f"Missing or out-of-order: {missing}."
            )

    return warnings


# ---------------------------------------------------------------------------
# Post-build PPTX structural integrity check
# ---------------------------------------------------------------------------

_NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_NS_P  = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_A  = "http://schemas.openxmlformats.org/drawingml/2006/main"
_INT_RE = re.compile(r"^-?\d+$")


def validate_pptx(path: str) -> tuple:
    """Scan a built .pptx for structural issues that trigger PowerPoint's
    'needs repair' dialog. Returns (errors, warnings).

    Catches:
      - non-integer cx/cy/x/y EMU attribute values (xsd:long violation)
      - duplicate cNvPr shape IDs within a single slide
      - [Content_Types].xml Override entries pointing at missing zip members
      - empty <a:r> runs with no <a:t> child (soft warning)

    Empty errors list means the file should open cleanly in PowerPoint.
    """
    try:
        from lxml import etree as _ET
    except ImportError:
        from xml.etree import ElementTree as _ET  # type: ignore[no-redef]

    errors: list = []
    warnings: list = []

    try:
        zf = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, FileNotFoundError, OSError) as exc:
        return ([f"could not open {path}: {exc}"], [])

    with zf as z:
        names = z.namelist()
        archive_set = {"/" + n for n in names}

        # 1. Content_Types — every Override PartName must resolve inside the zip.
        try:
            ct_root = _ET.fromstring(z.read("[Content_Types].xml"))
            for ov in ct_root.findall("{%s}Override" % _NS_CT):
                part = ov.get("PartName")
                if part and part not in archive_set:
                    errors.append(
                        f"[Content_Types].xml references missing file: {part}"
                    )
        except Exception as exc:
            errors.append(f"could not parse [Content_Types].xml: {exc}")

        # 2. Per-slide checks.
        slide_files = sorted(
            [n for n in names
             if n.startswith("ppt/slides/slide") and n.endswith(".xml")],
            key=lambda s: int(s.split("slide")[-1].split(".")[0]),
        )
        for sp in slide_files:
            slide_n = int(sp.split("slide")[-1].split(".")[0])
            try:
                tree = _ET.fromstring(z.read(sp))
            except Exception as exc:
                errors.append(f"slide {slide_n}: invalid XML — {exc}")
                continue

            # 2a. Non-integer EMU coordinate values (gantt float-EMU bug etc.).
            for el in tree.iter():
                for attr in ("cx", "cy", "x", "y"):
                    val = el.get(attr)
                    if val is not None and not _INT_RE.match(val):
                        tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
                        errors.append(
                            f"slide {slide_n}: <{tag} {attr}={val!r}> "
                            f"is not an integer EMU (triggers PowerPoint repair warning)"
                        )

            # 2b. Duplicate cNvPr shape IDs within the slide.
            ids: list = []
            for cnv in tree.iter("{%s}cNvPr" % _NS_P):
                sid = cnv.get("id")
                if sid:
                    ids.append(sid)
            from collections import Counter as _Counter
            dupes = sorted(sid for sid, cnt in _Counter(ids).items() if cnt > 1)
            if dupes:
                errors.append(f"slide {slide_n}: duplicate shape IDs {dupes}")

            # 2c. Empty <a:r> runs — render as invisible whitespace (warning only).
            empty_runs = sum(
                1 for r in tree.iter("{%s}r" % _NS_A)
                if r.find("{%s}t" % _NS_A) is None
            )
            if empty_runs:
                warnings.append(
                    f"slide {slide_n}: {empty_runs} <a:r> element(s) with no <a:t> text"
                )

    return (errors, warnings)
