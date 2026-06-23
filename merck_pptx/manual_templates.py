"""manual_templates.py — User-managed template directory for Mac/Linux users.

Windows users with empower installed get all Merck Corporate Design templates
synced automatically.  Everyone else (Mac, Linux, or Windows without empower)
can place manually-downloaded .pptx files in a local folder and have the
pipeline use them.

Directory:  ~/.merck_pptx/templates/
            (or a custom path via config.yaml: manual_templates.dir)

File naming convention:  {division}_{color_theme}.pptx

Examples:
    merck_organic.pptx
    merck_plastic.pptx
    emd_serono_organic.pptx
    millipore_sigma_functional.pptx
    merck_asia_synthetic.pptx

Valid division keys:
    merck  merck_asia  emd_serono  millipore_sigma  emd_electronics  usa

Valid color_theme keys:
    plastic  functional  organic  synthetic  technical  electronics

How to obtain templates (no empower):
  1. Ask a Windows colleague to open PowerPoint with empower installed.
  2. Go to: empower tab → Corporate Design Templates → Master Templates → Merck
  3. Right-click the desired template → Export to file → save as .pptx.
  4. Rename to {division}_{color_theme}.pptx and place in ~/.merck_pptx/templates/.
  5. Run 'python -m merck_pptx list-templates' to confirm it is recognised.

Note: division='merck' always works without any manual download — the bundled
EU_Merck_Themed.pptx / USA_Merck_Themed_Base_v1.pptx cover all 6 color themes.
"""

from __future__ import annotations

import pathlib
import zipfile
from typing import Optional

# The 6 official Merck Corporate Design color theme keys — canonical source of truth.
# Import this tuple wherever the theme list is needed instead of redefining it.
# Order matters for filename parsing: longer names must precede any name that is
# a suffix of another (none exist today, but guard for future additions).
VALID_THEMES: tuple = (
    "electronics",   # must precede any shorter name that shares a suffix
    "functional",
    "synthetic",
    "technical",
    "organic",
    "plastic",
)


def manual_templates_dir() -> pathlib.Path:
    """Return the user-managed manual templates directory.

    Default: ~/.merck_pptx/templates/
    Override: set 'manual_templates.dir' in config.yaml to a custom path.

    The directory is NOT created automatically — use setup_merck_pptx.py or
    create it manually before placing files there.
    """
    try:
        from .config import get
        custom = get("manual_templates", "dir", None)
        if custom and str(custom).strip():
            return pathlib.Path(str(custom).strip())
    except Exception:
        pass
    return pathlib.Path.home() / ".merck_pptx" / "templates"


def is_empower_template(path: pathlib.Path) -> bool:
    """Return True if path looks like an empower BinaryFile template.

    empower BinaryFile templates always contain exactly 13 slide layouts.
    Uses a fast zip-level scan (no pptx parse) to count them.
    Returns False on any error so the caller falls back to safe defaults.
    """
    try:
        with zipfile.ZipFile(str(path)) as zf:
            count = sum(
                1 for name in zf.namelist()
                if name.startswith("ppt/slideLayouts/slideLayout")
                and name.endswith(".xml")
            )
        return count == 13
    except Exception:
        return False


def resolve_manual_template(division: str, color_theme: str) -> Optional[pathlib.Path]:
    """Return the path to a manual template file if it exists, else None.

    Looks for:  ~/.merck_pptx/templates/{division}_{color_theme}.pptx

    Returns None silently if the directory or file is absent — never raises.
    Both division and color_theme are normalised (lower-case, hyphens → underscores)
    before building the filename, so 'emd-serono' and 'emd_serono' resolve the same.
    """
    d = str(division or "merck").lower().strip().replace("-", "_").replace(" ", "_")
    t = str(color_theme or "plastic").lower().strip()
    try:
        path = manual_templates_dir() / f"{d}_{t}.pptx"
        return path if path.exists() else None
    except Exception:
        return None


def list_manual_templates() -> list:
    """Scan the manual templates directory and return all recognised templates.

    Returns a list of dicts: [{"division": str, "color_theme": str, "path": str}]

    Recognition rule: a .pptx file is included only if its stem ends with
    '_<valid_theme>'.  This correctly handles multi-underscore division names:
        emd_serono_organic.pptx  →  division='emd_serono', color_theme='organic'
        millipore_sigma_plastic.pptx  →  division='millipore_sigma', color_theme='plastic'

    Files not matching the pattern (backups, unrelated files) are silently skipped.
    Returns [] if the directory does not exist.
    """
    d = manual_templates_dir()
    if not d.exists():
        return []

    results = []
    for f in sorted(d.glob("*.pptx")):
        stem = f.stem
        for theme in VALID_THEMES:
            suffix = f"_{theme}"
            if stem.endswith(suffix) and len(stem) > len(suffix):
                division_part = stem[: -len(suffix)]
                results.append({
                    "division":    division_part,
                    "color_theme": theme,
                    "path":        str(f),
                })
                break  # one theme match per file is enough

    return results
