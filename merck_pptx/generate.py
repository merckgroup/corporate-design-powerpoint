"""
Top-level autonomous pipeline router.

Accepts a markdown file, an existing .pptx, a plan dict, or a plan JSON file
and produces a Merck-branded .pptx.
"""

import pathlib

from merck_pptx.build_from_plan import build_from_plan
from merck_pptx import llm


def _extract_pptx_content(source_path: pathlib.Path) -> str:
    """Extract slide text from an existing .pptx into a structured string."""
    from pptx import Presentation

    prs = Presentation(str(source_path))
    lines = []
    for i, slide in enumerate(prs.slides, start=1):
        lines.append(f"\n### Slide {i}")
        slide_has_title = False  # reset per slide so only the first shape is tagged TITLE
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text.strip()
            if not text:
                continue
            if not slide_has_title:
                tag = "TITLE"
                slide_has_title = True
            else:
                tag = "BODY"
            lines.append(f"{tag}: {text}")
    return "\n".join(lines)


def generate_deck(source, output_path, meta=None) -> str:
    """
    Autonomous end-to-end pipeline. Any input type → Merck-branded .pptx.

    Parameters
    ----------
    source : dict | str | pathlib.Path
        One of:
          - A plan dict  →  built directly (no LLM)
          - A .json file →  loaded and built directly (no LLM)
          - A .md file   →  LLM converts content to plan, then builds
          - A .pptx file →  slides extracted, LLM converts to plan, then builds
    output_path : str | pathlib.Path
        Destination .pptx path.
    meta : dict | None
        Required when source is a .md or .pptx file. Contains the 6 gate answers:
        region, deck_label, classification, month_year, audience, deck_style.
        Ignored when source is a dict or .json (meta is already embedded in the plan).

    Returns
    -------
    str
        Absolute path of the saved .pptx.

    Raises
    ------
    ValueError
        If meta is None when source is .md or .pptx.
    EnvironmentError
        If AIP_BASE_URL or AIP_TOKEN are not set when LLM is required.
    """
    # Plan dict — direct build, no LLM
    if isinstance(source, dict):
        return build_from_plan(source, output_path)

    # Resolve to a Path for all file-based inputs (str or pathlib.Path)
    if isinstance(source, (str, pathlib.Path)):
        source_path = pathlib.Path(source)
    else:
        raise ValueError(
            f"Unsupported source type: {type(source).__name__}. "
            f"Expected a dict, or a path (str or pathlib.Path) to a .json, .md, or .pptx file."
        )

    suffix = source_path.suffix.lower()

    # .json plan file — direct build, no LLM
    if suffix == ".json":
        return build_from_plan(source_path, output_path)

    # .md or .pptx — LLM path
    if meta is None:
        raise ValueError(
            "meta is required when source is a .md or .pptx file. "
            "Provide a dict with region, deck_label, classification, "
            "month_year, audience, and deck_style."
        )

    if suffix == ".md":
        raw_content = source_path.read_text(encoding="utf-8")
    elif suffix in (".pptx", ".ppt"):
        raw_content = _extract_pptx_content(source_path)
    else:
        raise ValueError(
            f"Unsupported file type '{suffix}'. "
            f"Expected .json (plan), .md (markdown), .pptx, or .ppt."
        )

    plan = llm.generate_plan(raw_content, meta)
    return build_from_plan(plan, output_path)
