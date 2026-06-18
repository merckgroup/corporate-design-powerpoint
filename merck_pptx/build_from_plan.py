"""Canonical Merck deck builder: JSON plan → PowerPoint.

Reads a slide plan (dict or .json file) and calls the correct build_*
function in merck_layouts.py for every slide. Handles style resolution,
auto-promote, agenda auto-fill, appendix slides, hyperlink wiring, and
post-build speaker notes / image placement.

Usage:
    from merck_pptx import build_from_plan
    build_from_plan("plan.json", "output/deck.pptx")

Or from a dict:
    from merck_pptx import build_from_plan
    build_from_plan(plan_dict, "output/deck.pptx")
"""

from __future__ import annotations

import json
import pathlib
import sys
from typing import Optional

from merck_pptx.merck_layouts import (
    open_deck, save_deck, AUTO_PROMOTE_EXECUTIVE,
    add_speaker_notes, add_slide_jump_hyperlink, add_image,
    build_cover, build_exec_summary, build_agenda, build_section_divider,
    build_chart_slide, build_two_column, build_three_column, build_2x2_matrix,
    build_phase_process, build_vertical_numbered, build_waterfall_slide,
    build_decision_rows, build_gantt, build_hero_stat, build_close,
    build_stat_strip, build_before_after, build_milestone_timeline,
    build_status_table, build_hub_spoke, build_pillar_detail,
    build_four_column, build_label_rows,
    build_circular_flow, build_org_chart, build_topic_set, build_arrow_chain,
    build_pull_quote, build_donut_chart, build_kpi_dashboard,
    build_fishbone, build_journey_map,
    build_icon_grid, build_funnel,
    build_comparison_table, build_score_table, build_influence_diagram,
    build_word_cloud, build_pyramid, build_venn, build_risk_heatmap,
    build_radar_chart, build_pros_cons, build_layered_stack,
    build_photo_text,
)
from merck_pptx.validate_plan import validate_plan, ValidationError

_PKG_DIR = pathlib.Path(__file__).parent
_TEMPLATES = {
    "eu":     _PKG_DIR / "templates" / "EU_Merck_Themed.pptx",
    "usa":    _PKG_DIR / "templates" / "USA_Merck_Themed_Base_v1.pptx",
    "us":     _PKG_DIR / "templates" / "USA_Merck_Themed_Base_v1.pptx",
    "canada": _PKG_DIR / "templates" / "USA_Merck_Themed_Base_v1.pptx",
}
_TEMPLATE_DEFAULT = _PKG_DIR / "templates" / "EU_Merck_Themed.pptx"
_TEMPLATE_DIR     = _PKG_DIR / "templates"

# Permitted image file extensions — prevents a crafted plan from embedding
# arbitrary files (e.g. databases, documents) by disguising them as images.
_SAFE_IMAGE_EXTENSIONS = frozenset({
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp", ".emf", ".wmf",
})


class BuildError(RuntimeError):
    """Wraps a layout-builder exception with page + layout context."""


# Layouts that carry no section-number circle.
# Must stay in sync with _STRUCTURAL_LAYOUTS in validate_plan.py.
_NO_CIRCLE_LAYOUTS = {
    "cover", "agenda", "section_divider", "close", "exec_summary", "hero_stat", "pull_quote"
}


def _is_appendix(slide: dict) -> bool:
    return bool(slide.get("appendix"))


def _short_title(text, max_chars: int = 40) -> str:
    if text is None:
        return ""
    if isinstance(text, (list, tuple)):
        text = "".join(seg[0] if isinstance(seg, (list, tuple)) else str(seg) for seg in text)
    s = str(text).strip()
    if len(s) <= max_chars:
        return s
    return s[:max_chars].rsplit(" ", 1)[0] + "…"


def _place_image(slide, img_spec: dict) -> None:
    from pptx.util import Inches as _In
    raw_path = img_spec.get("path")
    if not raw_path:
        return

    # Security: validate the path before embedding.
    # (1) Extension check — prevents arbitrary file types from being embedded.
    # (2) Directory jail — restricts images to the current working directory tree
    #     to prevent a crafted plan from exfiltrating arbitrary local files
    #     (e.g. via prompt injection setting image.path to a sensitive path).
    img_path = pathlib.Path(raw_path).resolve()
    if img_path.suffix.lower() not in _SAFE_IMAGE_EXTENSIONS:
        print(
            f"WARNING: image not placed — '{img_path.name}' is not a recognised image type. "
            f"Allowed: {sorted(_SAFE_IMAGE_EXTENSIONS)}",
            file=sys.stderr,
        )
        return
    _allowed_image_root = pathlib.Path.cwd().resolve()
    if not str(img_path).startswith(str(_allowed_image_root)):
        print(
            f"WARNING: image not placed — '{img_path}' is outside the allowed "
            f"directory ({_allowed_image_root}). Move the image file inside the "
            f"project directory.",
            file=sys.stderr,
        )
        return

    placement = (img_spec.get("placement") or "right_panel").lower()
    presets = {
        "hero":        (6.83, 1.30, 6.00, 5.00),
        "right_panel": (8.50, 3.10, 4.50, 3.50),
        "background":  (0.00, 0.00, 13.33, 7.50),
    }
    x, y, w, h = presets.get(placement, presets["right_panel"])
    w = img_spec.get("w") or w
    h = img_spec.get("h") or h
    try:
        add_image(slide, str(img_path), _In(x), _In(y),
                  w=_In(w) if w else None, h=_In(h) if h else None)
    except Exception as exc:
        print(f"WARNING: image not placed: {exc}", file=sys.stderr)


def _wire_agenda_hyperlinks(prs, ordered_slides: list) -> None:
    sn_to_idx: dict[str, int] = {}
    for i, s in enumerate(ordered_slides):
        if s.get("layout") in _NO_CIRCLE_LAYOUTS or _is_appendix(s):
            continue
        sn = s.get("section_number")
        if sn is None:
            continue
        sn_str = str(sn).strip()
        sn_to_idx[sn_str] = i
        try:
            sn_to_idx[f"{int(sn_str):02d}"] = i
        except (ValueError, TypeError):
            pass

    for i, s in enumerate(ordered_slides):
        if s.get("layout") != "agenda":
            continue
        agenda_slide = prs.slides[i]
        for shp in agenda_slide.shapes:
            name = getattr(shp, "name", "") or ""
            if not name.startswith("AgendaChapter_"):
                continue
            key = name.split("_", 1)[1]
            target_idx = sn_to_idx.get(key) or sn_to_idx.get(key.lstrip("0"))
            if target_idx is None:
                continue
            try:
                add_slide_jump_hyperlink(shp, prs.slides[target_idx])
            except Exception as exc:
                print(f"WARNING: agenda hyperlink '{name}' not set: {exc}", file=sys.stderr)


def _autofill_agenda(plan: dict) -> None:
    slides = plan.get("slides") or []
    content_slides = [
        s for s in slides
        if s.get("layout") not in _NO_CIRCLE_LAYOUTS and not _is_appendix(s)
    ]
    for s in slides:
        if s.get("layout") != "agenda":
            continue
        content = s.get("content") or {}
        if content.get("chapters"):
            continue
        derived = []
        for cs in content_slides:
            sn = cs.get("section_number")
            num = ""
            if sn is not None:
                try:
                    num = f"{int(sn):02d}"
                except (ValueError, TypeError):
                    num = str(sn)
            derived.append({
                "number": num,
                "title": _short_title(cs.get("action_title"), 40),
                "subtitle": _short_title(cs.get("category") or "", 30),
            })
        if derived:
            s["content"] = {**content, "chapters": derived}


# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------

def _resolve_style(slide: dict, meta: dict) -> str:
    style = slide.get("style", "inherit")
    if style == "inherit" or not style:
        style = meta.get("deck_style", "merck_executive")
    if slide.get("page_function") in AUTO_PROMOTE_EXECUTIVE:
        style = "merck_executive"
    return style


def _resolve_category(slide: dict) -> Optional[str]:
    """Auto-set category to match page_function for auto-promote slides.
    The library triggers on category, not page_function."""
    pf = slide.get("page_function")
    if pf in AUTO_PROMOTE_EXECUTIVE:
        return pf
    return slide.get("category")


def _common_kwargs(slide: dict, meta: dict, total) -> dict:
    return {
        "style":          _resolve_style(slide, meta),
        "page":           slide.get("page"),
        "total":          total,
        "section_number": slide.get("section_number"),
        "category":       _resolve_category(slide),
    }


def _content(slide: dict) -> dict:
    c = slide.get("content")
    return c if isinstance(c, dict) else {}


# ---------------------------------------------------------------------------
# Per-layout builder wrappers
# ---------------------------------------------------------------------------

def _build_cover(prs, meta, slide, total):
    c = _content(slide)
    return build_cover(
        prs, meta,
        action_title=slide.get("action_title"),
        subtitle=slide.get("subtitle", ""),
        style=_resolve_style(slide, meta),
        key_messages=c.get("key_messages"),
        phases=c.get("phases"),
        authors=c.get("authors"),
        top_bar=bool(meta.get("cover_top_bar", False)),
        page=slide.get("page"),
        total=total,
    )


def _build_exec_summary(prs, meta, slide, total):
    c = _content(slide)
    return build_exec_summary(
        prs, meta,
        action_title=slide.get("action_title", ""),
        key_messages=c.get("key_messages", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_agenda(prs, meta, slide, total):
    c = _content(slide)
    return build_agenda(
        prs, meta,
        chapters=c.get("chapters", []),
        style=_resolve_style(slide, meta),
        action_title=slide.get("action_title"),
        page=slide.get("page"),
        total=total,
    )


def _build_section_divider(prs, meta, slide, total):
    number = slide.get("section_number") or _content(slide).get("number") or ""
    return build_section_divider(
        prs, meta,
        number=number,
        title=slide.get("action_title"),
        style=_resolve_style(slide, meta),
        page=slide.get("page"),
        total=total,
        takeaway=slide.get("takeaway"),
        source=slide.get("source"),
    )


def _build_chart_slide(prs, meta, slide, total):
    c = _content(slide)
    return build_chart_slide(
        prs, meta,
        action_title=slide.get("action_title", ""),
        chart=c.get("chart", {}),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        callouts=c.get("callouts"),
        subtitle=slide.get("subtitle"),
        footnotes=c.get("footnotes"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_two_column(prs, meta, slide, total):
    c = _content(slide)
    return build_two_column(
        prs, meta,
        action_title=slide.get("action_title", ""),
        left=c.get("left", {}),
        right=c.get("right", {}),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_three_column(prs, meta, slide, total):
    c = _content(slide)
    return build_three_column(
        prs, meta,
        action_title=slide.get("action_title", ""),
        columns=c.get("columns", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_2x2_matrix(prs, meta, slide, total):
    c = _content(slide)
    return build_2x2_matrix(
        prs, meta,
        action_title=slide.get("action_title", ""),
        x_axis=c.get("x_axis", {}),
        y_axis=c.get("y_axis", {}),
        quadrants=c.get("quadrants", {}),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_phase_process(prs, meta, slide, total):
    c = _content(slide)
    return build_phase_process(
        prs, meta,
        action_title=slide.get("action_title", ""),
        phases=c.get("phases", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        highlight_index=c.get("highlight_index"),
        show_arrows=c.get("show_arrows", True),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_vertical_numbered(prs, meta, slide, total):
    c = _content(slide)
    return build_vertical_numbered(
        prs, meta,
        action_title=slide.get("action_title", ""),
        items=c.get("items", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_waterfall_slide(prs, meta, slide, total):
    c = _content(slide)
    chart = c.get("chart") or {}
    # Schema: content.chart.data.bars — extract bars for the layout function.
    bars = chart.get("data", {}).get("bars") or chart.get("bars") or []
    return build_waterfall_slide(
        prs, meta,
        action_title=slide.get("action_title", ""),
        bars=bars,
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        footnotes=c.get("footnotes"),
        methodology_note=c.get("methodology_note"),
        subtitle=slide.get("subtitle"),
        **_common_kwargs(slide, meta, total),
    )


def _build_decision_rows(prs, meta, slide, total):
    c = _content(slide)
    return build_decision_rows(
        prs, meta,
        action_title=slide.get("action_title", ""),
        decisions=c.get("decisions", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_gantt(prs, meta, slide, total):
    c = _content(slide)
    return build_gantt(
        prs, meta,
        action_title=slide.get("action_title", ""),
        quarters=c.get("quarters", []),
        rows=c.get("rows", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_hero_stat(prs, meta, slide, total):
    c = _content(slide)
    return build_hero_stat(
        prs, meta,
        stat=c.get("stat", {}),
        context=c.get("context", ""),
        source=c.get("source") or slide.get("source"),
        style=_resolve_style(slide, meta),
        page=slide.get("page"),
        total=total,
        category=_resolve_category(slide),
    )


def _build_close(prs, meta, slide, total):
    c = _content(slide)
    return build_close(
        prs, meta,
        action_statement=c.get("action_statement") or slide.get("action_title"),
        style=_resolve_style(slide, meta),
        page=slide.get("page"),
        total=total,
        takeaway=slide.get("takeaway"),
        source=slide.get("source"),
    )


def _build_stat_strip(prs, meta, slide, total):
    c = _content(slide)
    return build_stat_strip(
        prs, meta,
        action_title=slide.get("action_title", ""),
        stats=c.get("stats", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_before_after(prs, meta, slide, total):
    c = _content(slide)
    return build_before_after(
        prs, meta,
        action_title=slide.get("action_title", ""),
        before=c.get("before", {}),
        after=c.get("after", {}),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        before_label=c.get("before_label", "TODAY"),
        after_label=c.get("after_label", "TOMORROW"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_milestone_timeline(prs, meta, slide, total):
    c = _content(slide)
    return build_milestone_timeline(
        prs, meta,
        action_title=slide.get("action_title", ""),
        milestones=c.get("milestones", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_status_table(prs, meta, slide, total):
    c = _content(slide)
    return build_status_table(
        prs, meta,
        action_title=slide.get("action_title", ""),
        columns=c.get("columns", []),
        rows=c.get("rows", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_hub_spoke(prs, meta, slide, total):
    c = _content(slide)
    return build_hub_spoke(
        prs, meta,
        action_title=slide.get("action_title", ""),
        hub=c.get("hub", {}),
        spokes=c.get("spokes", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_pillar_detail(prs, meta, slide, total):
    c = _content(slide)
    return build_pillar_detail(
        prs, meta,
        action_title=slide.get("action_title", ""),
        pillar_number=c.get("pillar_number", ""),
        pillar_label=c.get("pillar_label", "PILLAR"),
        owner=c.get("owner"),
        sections=c.get("sections", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_four_column(prs, meta, slide, total):
    c = _content(slide)
    return build_four_column(
        prs, meta,
        action_title=slide.get("action_title", ""),
        columns=c.get("columns", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_label_rows(prs, meta, slide, total):
    c = _content(slide)
    return build_label_rows(
        prs, meta,
        action_title=slide.get("action_title", ""),
        rows=c.get("rows", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        label_color=c.get("label_color"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_circular_flow(prs, meta, slide, total):
    c = _content(slide)
    return build_circular_flow(
        prs, meta,
        action_title=slide.get("action_title", ""),
        phases=c.get("phases", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_org_chart(prs, meta, slide, total):
    c = _content(slide)
    return build_org_chart(
        prs, meta,
        action_title=slide.get("action_title", ""),
        root=c.get("root", {}),
        children=c.get("children", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_topic_set(prs, meta, slide, total):
    c = _content(slide)
    return build_topic_set(
        prs, meta,
        action_title=slide.get("action_title", ""),
        topics=c.get("topics", []),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_arrow_chain(prs, meta, slide, total):
    c = _content(slide)
    return build_arrow_chain(
        prs, meta,
        action_title=slide.get("action_title", ""),
        steps=c.get("steps", []),
        consequence=c.get("consequence"),
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_pull_quote(prs, meta, slide, total):
    c = _content(slide)
    return build_pull_quote(
        prs, meta,
        quote=c.get("quote") or slide.get("quote"),
        attribution=c.get("attribution") or slide.get("attribution"),
        context=c.get("context"),
        action_title=slide.get("action_title"),
        **_common_kwargs(slide, meta, total),
    )


def _build_donut_chart(prs, meta, slide, total):
    c = _content(slide)
    return build_donut_chart(
        prs, meta,
        action_title=slide.get("action_title", ""),
        segments=c.get("segments", []),
        center_value=c.get("center_value"),
        center_label=c.get("center_label"),
        legend_title=c.get("legend_title"),
        **_common_kwargs(slide, meta, total),
    )


def _build_kpi_dashboard(prs, meta, slide, total):
    c = _content(slide)
    return build_kpi_dashboard(
        prs, meta,
        action_title=slide.get("action_title", ""),
        kpis=c.get("kpis", []),
        **_common_kwargs(slide, meta, total),
    )


def _build_icon_grid(prs, meta, slide, total):
    c = _content(slide)
    return build_icon_grid(
        prs, meta,
        action_title=slide.get("action_title", ""),
        items=c.get("items", []),
        columns=c.get("columns", 3),
        **_common_kwargs(slide, meta, total),
    )


def _build_journey_map(prs, meta, slide, total):
    c = _content(slide)
    return build_journey_map(
        prs, meta,
        action_title=slide.get("action_title", ""),
        phases=c.get("phases", []),
        rows=c.get("actors", []),         # schema key: actors; layout key: rows
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_funnel(prs, meta, slide, total):
    c = _content(slide)
    return build_funnel(
        prs, meta,
        action_title=slide.get("action_title", ""),
        inputs=c.get("inputs", []),      # schema key: inputs (not stages)
        output=c.get("output"),
        **_common_kwargs(slide, meta, total),
    )


def _build_comparison_table(prs, meta, slide, total):
    c = _content(slide)
    return build_comparison_table(
        prs, meta,
        action_title=slide.get("action_title", ""),
        options=c.get("options", []),    # schema keys: options + features
        features=c.get("features", []),
        **_common_kwargs(slide, meta, total),
    )


def _build_score_table(prs, meta, slide, total):
    c = _content(slide)
    return build_score_table(
        prs, meta,
        action_title=slide.get("action_title", ""),
        rows=c.get("rows", []),          # schema keys: rows + scale
        scale=c.get("scale", 5),
        scale_label=c.get("scale_label"),
        **_common_kwargs(slide, meta, total),
    )


def _build_influence_diagram(prs, meta, slide, total):
    c = _content(slide)
    return build_influence_diagram(
        prs, meta,
        action_title=slide.get("action_title", ""),
        center=c.get("center", {}),
        forces=c.get("forces", []),      # schema key: forces (not nodes)
        **_common_kwargs(slide, meta, total),
    )


def _build_word_cloud(prs, meta, slide, total):
    c = _content(slide)
    return build_word_cloud(
        prs, meta,
        action_title=slide.get("action_title", ""),
        words=c.get("words", []),
        **_common_kwargs(slide, meta, total),
    )


def _build_pyramid(prs, meta, slide, total):
    c = _content(slide)
    return build_pyramid(
        prs, meta,
        action_title=slide.get("action_title", ""),
        tiers=c.get("tiers", []),
        orientation=c.get("orientation", "up"),
        **_common_kwargs(slide, meta, total),
    )


def _build_venn(prs, meta, slide, total):
    c = _content(slide)
    return build_venn(
        prs, meta,
        action_title=slide.get("action_title", ""),
        circles=c.get("circles", []),
        intersection=c.get("intersection"),  # schema key: intersection
        **_common_kwargs(slide, meta, total),
    )


def _build_risk_heatmap(prs, meta, slide, total):
    c = _content(slide)
    return build_risk_heatmap(
        prs, meta,
        action_title=slide.get("action_title", ""),
        risks=c.get("risks", []),
        x_label=c.get("x_label", "LIKELIHOOD"),
        y_label=c.get("y_label", "IMPACT"),
        **_common_kwargs(slide, meta, total),
    )


def _build_radar_chart(prs, meta, slide, total):
    c = _content(slide)
    return build_radar_chart(
        prs, meta,
        action_title=slide.get("action_title", ""),
        axes=c.get("axes", []),
        series=c.get("series", []),
        **_common_kwargs(slide, meta, total),
    )


def _build_pros_cons(prs, meta, slide, total):
    c = _content(slide)
    return build_pros_cons(
        prs, meta,
        action_title=slide.get("action_title", ""),
        pros=c.get("pros", []),
        cons=c.get("cons", []),
        pros_label=c.get("pros_label", "ADVANTAGES"),
        cons_label=c.get("cons_label", "RISKS"),
        subject=c.get("subject"),
        **_common_kwargs(slide, meta, total),
    )


def _build_layered_stack(prs, meta, slide, total):
    c = _content(slide)
    return build_layered_stack(
        prs, meta,
        action_title=slide.get("action_title", ""),
        layers=c.get("layers", []),
        orientation=c.get("orientation", "vertical"),
        **_common_kwargs(slide, meta, total),
    )


def _build_photo_text(prs, meta, slide, total):
    c = _content(slide)
    return build_photo_text(
        prs, meta,
        action_title=slide.get("action_title", ""),
        image_path=c.get("image_path"),      # schema keys: image_path, image_side
        image_label=c.get("image_label"),
        title=c.get("title"),
        bullets=c.get("bullets", []),
        image_side=c.get("image_side", "left"),
        **_common_kwargs(slide, meta, total),
    )


def _build_fishbone(prs, meta, slide, total):
    c = _content(slide)
    return build_fishbone(
        prs, meta,
        action_title=slide.get("action_title", ""),
        effect=c.get("effect", ""),
        bones=c.get("bones", []),            # schema key: bones (not causes)
        takeaway=slide.get("takeaway", ""),
        source=slide.get("source"),
        subtitle=slide.get("subtitle"),
        methodology_note=c.get("methodology_note"),
        **_common_kwargs(slide, meta, total),
    )


def _build_columns_auto(prs, meta, slide, total):
    """Auto-dispatch to two/three/four_column based on column count."""
    c = _content(slide)
    cols = c.get("columns", [])
    n = len(cols)
    layout = "three_column" if n == 3 else ("two_column" if n <= 2 else "four_column")
    dispatch = {
        "two_column":   _build_two_column,
        "three_column": _build_three_column,
        "four_column":  _build_four_column,
    }
    patched = {**slide, "layout": layout}
    return dispatch[layout](prs, meta, patched, total)


_DISPATCH = {
    "cover":              _build_cover,
    "exec_summary":       _build_exec_summary,
    "agenda":             _build_agenda,
    "section_divider":    _build_section_divider,
    "chart_slide":        _build_chart_slide,
    "two_column":         _build_two_column,
    "three_column":       _build_three_column,
    "2x2_matrix":         _build_2x2_matrix,
    "phase_process":      _build_phase_process,
    "vertical_numbered":  _build_vertical_numbered,
    "waterfall_slide":    _build_waterfall_slide,
    "decision_rows":      _build_decision_rows,
    "gantt":              _build_gantt,
    "hero_stat":          _build_hero_stat,
    "close":              _build_close,
    "stat_strip":         _build_stat_strip,
    "before_after":       _build_before_after,
    "milestone_timeline": _build_milestone_timeline,
    "status_table":       _build_status_table,
    "hub_spoke":          _build_hub_spoke,
    "pillar_detail":      _build_pillar_detail,
    "four_column":        _build_four_column,
    "label_rows":         _build_label_rows,
    "circular_flow":      _build_circular_flow,
    "org_chart":          _build_org_chart,
    "topic_set":          _build_topic_set,
    "arrow_chain":        _build_arrow_chain,
    "columns":            _build_columns_auto,
    "pull_quote":         _build_pull_quote,
    "donut_chart":        _build_donut_chart,
    "kpi_dashboard":      _build_kpi_dashboard,
    "icon_grid":          _build_icon_grid,
    "journey_map":        _build_journey_map,
    "funnel":             _build_funnel,
    "comparison_table":   _build_comparison_table,
    "score_table":        _build_score_table,
    "influence_diagram":  _build_influence_diagram,
    "word_cloud":         _build_word_cloud,
    "pyramid":            _build_pyramid,
    "venn":               _build_venn,
    "risk_heatmap":       _build_risk_heatmap,
    "radar_chart":        _build_radar_chart,
    "pros_cons":          _build_pros_cons,
    "layered_stack":      _build_layered_stack,
    "photo_text":         _build_photo_text,
    "fishbone":           _build_fishbone,
}


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
# Plan sanitization helpers
# ---------------------------------------------------------------------------

def _trunc(s: object, limit: int) -> object:
    """Truncate a string to `limit` chars, appending '...' if trimmed."""
    if not isinstance(s, str) or len(s) <= limit:
        return s
    if limit <= 3:
        return s[:limit]
    return s[:limit - 3] + "..."


def _sanitize_obj(obj: object, rules: dict) -> None:
    """Apply truncation rules to a dict, mutating in place."""
    if not isinstance(obj, dict):
        return
    for field, limit in rules.items():
        if isinstance(obj.get(field), str):
            obj[field] = _trunc(obj[field], limit)


# Per-field character limits (must align with rendering engine constraints).
# Top-level slide fields:
_SLIDE_LIMITS = {"action_title": 80}
# Content fields — applied at every dict level (top-level content and list items).
# Using one dict for both avoids divergence and closes the gap where list-item
# fields like "context" (used in kpi_dashboard) were not being truncated.
_CONTENT_FIELD_LIMITS = {
    "takeaway": 120,
    "body":     130,
    "desc":     160,
    "context":  130,
    "note":     80,
}


def _sanitize_content(content: object) -> None:
    """Recursively sanitize content dict: truncate known long-text fields."""
    if not isinstance(content, dict):
        return
    _sanitize_obj(content, _CONTENT_FIELD_LIMITS)
    for val in content.values():
        if isinstance(val, list):
            for item in val:
                if isinstance(item, dict):
                    _sanitize_obj(item, _CONTENT_FIELD_LIMITS)
                    for nested_val in item.values():
                        if isinstance(nested_val, list):
                            for nested_item in nested_val:
                                if isinstance(nested_item, dict):
                                    _sanitize_obj(nested_item, _CONTENT_FIELD_LIMITS)
        elif isinstance(val, dict):
            _sanitize_content(val)


def _sanitize_plan(plan: dict) -> None:
    """Truncate overflow-prone fields across all slides before validation."""
    meta = plan.get("meta") or {}
    for slide in plan.get("slides") or []:
        # Takeaway lives both at slide level and inside content — handle first
        # so _sanitize_content below doesn't double-truncate it.
        for loc in (slide, slide.get("content") or {}):
            if isinstance(loc.get("takeaway"), str) and len(loc["takeaway"]) > 120:
                loc["takeaway"] = _trunc(loc["takeaway"], 120)

        # Cover slide: ensure a subtitle is present (either "Title; Subtitle" in
        # action_title, or a standalone subtitle field).  The LLM occasionally
        # omits it; auto-patch using the deck_label rather than hard-failing.
        # Do this BEFORE _sanitize_obj so the combined string is truncated together.
        if slide.get("layout") == "cover":
            raw = (slide.get("action_title") or "").rstrip().rstrip(";").rstrip()
            has_sub = (";" in (slide.get("action_title") or "") and
                       (slide.get("action_title") or "").split(";", 1)[1].strip()) or \
                      (slide.get("subtitle") or "").strip() or \
                      ((slide.get("content") or {}).get("subtitle") or "").strip()
            if not has_sub:
                deck_label = meta.get("deck_label") or "Merck Presentation"
                combined = (raw + "; " + deck_label) if raw else deck_label
                # Keep within 60-char cover-title limit (merck_layouts enforces this).
                slide["action_title"] = _trunc(combined, 60)

        _sanitize_obj(slide, _SLIDE_LIMITS)
        _sanitize_content(slide.get("content") or {})


# ---------------------------------------------------------------------------

def build_from_plan(plan, output_path, base_pptx: Optional[str] = None,
                    strict: bool = True) -> str:
    """
    Build a Merck PowerPoint deck from a slide plan.

    Parameters
    ----------
    plan : dict | str | pathlib.Path
        A plan dict, or a path to a JSON file.
    output_path : str | pathlib.Path
        Destination .pptx path.
    base_pptx : str | None
        Explicit path to a Merck theme .pptx. If None, selected from
        meta.region using the package-bundled templates.
    strict : bool
        If True (default), raise ValidationError on hard errors.

    Returns
    -------
    str
        Absolute path of the saved file.
    """
    if isinstance(plan, (str, pathlib.Path)):
        with open(plan, encoding="utf-8") as f:
            plan = json.load(f)

    # Security: validate output path ends with .pptx — sanity guard against
    # accidental overwrites of non-pptx files (e.g. output_path typo).
    output_path = pathlib.Path(output_path)
    if output_path.suffix.lower() != ".pptx":
        raise ValueError(
            f"output_path must end with .pptx, got: '{output_path.suffix}'"
        )

    # Auto-sanitize: truncate fields that overflow fixed-dimension text boxes in the
    # rendering engine. Truncation is a silent safety net — the LLM prompt instructs
    # the model to stay within these limits, but the build should not fail on overshoot.
    _sanitize_plan(plan)

    # Auto-fill agenda chapters before validation sees the plan.
    _autofill_agenda(plan)

    warnings = validate_plan(plan)
    for w in warnings:
        print(f"WARNING: {w}")

    meta = plan.get("meta") or {}
    slides = plan.get("slides") or []
    main_slides = [s for s in slides if not _is_appendix(s)]
    appendix_slides = [s for s in slides if _is_appendix(s)]
    total = len(main_slides)

    # Renumber pages so progress bar is consistent
    for i, s in enumerate(main_slides, start=1):
        s["page"] = i
    for i, s in enumerate(appendix_slides, start=1):
        s["page"] = f"A{i}"

    if base_pptx is None:
        region = str(meta.get("region", "eu")).lower().strip()
        template_path = _TEMPLATES.get(region, _TEMPLATE_DEFAULT)
    else:
        # Security: restrict base_pptx to the package templates directory so
        # a caller cannot force-load an arbitrary or maliciously crafted .pptx.
        bp = pathlib.Path(base_pptx).resolve()
        if not bp.is_file() or bp.suffix.lower() != ".pptx":
            raise ValueError(f"base_pptx must be an existing .pptx file: {bp}")
        if not str(bp).startswith(str(_TEMPLATE_DIR.resolve())):
            raise ValueError(
                f"base_pptx must be inside the package templates directory "
                f"({_TEMPLATE_DIR}). Got: {bp}"
            )
        template_path = bp

    prs = open_deck(str(template_path))
    ordered = main_slides + appendix_slides

    for s in ordered:
        layout = s.get("layout")
        builder = _DISPATCH.get(layout)
        if builder is None:
            raise BuildError(
                f"page {s.get('page', '?')}: unknown layout {layout!r}"
            )

        is_apx = _is_appendix(s)
        if is_apx:
            if not s.get("category"):
                s["category"] = "APPENDIX"
            slide_total = None
        else:
            slide_total = total

        try:
            built_slide = builder(prs, meta, s, slide_total)
        except Exception as exc:
            raise BuildError(
                f"page {s.get('page', '?')} ({layout}): {exc}"
            ) from exc

        # Speaker notes
        notes = s.get("notes")
        if notes and built_slide is not None:
            try:
                add_speaker_notes(built_slide, notes)
            except Exception as exc:
                print(
                    f"WARNING: speaker notes not written for page {s.get('page')}: {exc}",
                    file=sys.stderr,
                )

        # Image placement
        img = _content(s).get("image")
        if img and built_slide is not None and img.get("path"):
            _place_image(built_slide, img)

    _wire_agenda_hyperlinks(prs, ordered)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    save_deck(prs, str(output_path))
    return str(output_path.resolve())
