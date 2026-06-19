"""binary_templates.py — BinaryFile-based template registry.

Replaces the static merck_pptx/templates/ files as the primary source of
Merck-branded PPTX templates.  The empower add-in syncs full-featured
theme PPTXs (13 layouts each) to a local BinaryFiles directory; this
module maps (division, color_theme) to those files via a user-editable
JSON registry.

Typical usage (from build_from_plan.py):
    from .binary_templates import resolve_template_path
    path = resolve_template_path("merck", "functional")
    # returns Path or None
"""

from __future__ import annotations
import json
import os
import pathlib
from typing import Optional

# ---------------------------------------------------------------------------
# Registry path
# ---------------------------------------------------------------------------

_REGISTRY_PATH = pathlib.Path(__file__).parent / "binary_registry.json"

# ---------------------------------------------------------------------------
# BinaryFiles directory resolution
# ---------------------------------------------------------------------------

def _binary_dir() -> pathlib.Path:
    """Return the empower BinaryFiles directory.

    Resolution order:
    1. EMPOWER_BINARY_DIR environment variable
    2. Package config.empower_binary_dir()
    3. Windows default: %LOCALAPPDATA%\\empower\\data\\empower\\BinaryFiles
    """
    env = os.environ.get("EMPOWER_BINARY_DIR")
    if env:
        return pathlib.Path(env)
    try:
        from .config import empower_binary_dir
        return pathlib.Path(empower_binary_dir())
    except Exception:
        pass
    local = os.environ.get(
        "LOCALAPPDATA",
        str(pathlib.Path.home() / "AppData" / "Local"),
    )
    return pathlib.Path(local) / "empower" / "data" / "empower" / "BinaryFiles"


# ---------------------------------------------------------------------------
# Registry I/O
# ---------------------------------------------------------------------------

def load_registry() -> dict:
    """Load binary_registry.json; return empty dict if missing or invalid."""
    if not _REGISTRY_PATH.exists():
        return {}
    try:
        with open(_REGISTRY_PATH, encoding="utf-8") as f:
            data = json.load(f)
        # Strip comments key
        return {k: v for k, v in data.items() if not k.startswith("_")}
    except Exception:
        return {}


def save_registry(registry: dict) -> None:
    """Persist registry to binary_registry.json."""
    out = {"_comment": (
        "Maps (division, color_theme) to empower BinaryFile UID. "
        "Run 'python -m merck_pptx discover-templates' to find more UIDs."
    )}
    out.update(registry)
    with open(_REGISTRY_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------------
# UID resolution
# ---------------------------------------------------------------------------

def resolve_uid(division: str, color_theme: str) -> Optional[str]:
    """Return the empower UID for (division, color_theme), or None.

    Lookup order:
    1. Exact (division, color_theme) from registry
    2. ("merck", color_theme) as fallback when division is not registered
    """
    registry = load_registry()
    div = str(division or "merck").lower().strip().replace("-", "_").replace(" ", "_")
    theme = str(color_theme or "plastic").lower().strip()

    uid = (registry.get(div) or {}).get(theme)
    if uid:
        return uid
    # Fallback to merck division for unknown divisions
    if div != "merck":
        uid = (registry.get("merck") or {}).get(theme)
    return uid


def uid_to_path(uid: str) -> Optional[pathlib.Path]:
    """Return the full path to a BinaryFile PPTX, or None if not found."""
    path = _binary_dir() / f"{uid}.pptx"
    return path if path.exists() else None


def resolve_template_path(division: str, color_theme: str) -> Optional[pathlib.Path]:
    """High-level helper: division + color_theme → BinaryFile Path or None."""
    uid = resolve_uid(division, color_theme)
    if not uid:
        return None
    return uid_to_path(uid)


# ---------------------------------------------------------------------------
# Theme classification (color fingerprint)
# ---------------------------------------------------------------------------

def classify_theme_from_colors(title_colors: set) -> str:
    """Infer the empower color theme from a BinaryFile's Title-layout shape fills.

    Parameters
    ----------
    title_colors : set[str]
        Collected from non-placeholder, non-logo shapes in the 'Title' layout:
        hardcoded hex values (e.g. 'B4DC96') plus scheme names prefixed 's:'
        (e.g. 's:accent3').

    Returns
    -------
    str
        One of: plastic, organic, technical, functional, synthetic, electronics,
        or 'unknown'.
    """
    hc = {c for c in title_colors if not c.startswith("s:")}
    sc = {c for c in title_colors if c.startswith("s:")}

    if "B4DC96" in hc and "FFDCB9" not in hc and "E61E50" not in hc:
        return "plastic"
    if "FFDCB9" in hc and "E61E50" in hc:
        return "organic"
    if "FFDCB9" in hc and "E61E50" not in hc:
        return "technical"
    if "s:accent3" in sc and "B4DC96" not in hc and "FFDCB9" not in hc and "E61E50" not in hc:
        return "functional"
    if "s:accent3" in sc and ("A5CD50" in hc or "E61E50" in hc):
        return "functional"
    if "s:accent4" in sc and "E61E50" not in hc and "FFDCB9" not in hc and "2DBECD" not in hc:
        return "synthetic"
    if "s:accent4" in sc and ("E61E50" in hc or "2DBECD" in hc):
        return "electronics"
    return "unknown"


# ---------------------------------------------------------------------------
# BinaryFile discovery
# ---------------------------------------------------------------------------

def discover_templates() -> dict:
    """Scan the empower BinaryFiles directory and catalogue full-featured templates.

    Returns
    -------
    dict
        {uid: {"theme": str, "logo_shapes": int, "path": str}}
        Only includes 13-layout files (full template PPTXs, not icon libraries).
    """
    from pptx import Presentation as _Prs

    NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

    binary_dir = _binary_dir()
    if not binary_dir.exists():
        return {}

    results = {}
    for fname in sorted(binary_dir.iterdir()):
        if fname.suffix.lower() != ".pptx":
            continue
        uid = fname.stem
        try:
            prs = _Prs(str(fname))
            if len(prs.slide_layouts) != 13:
                continue

            title_layout = next(
                (l for l in prs.slide_layouts if l.name == "Title"), None
            )
            if title_layout is None:
                continue

            title_colors: set = set()
            logo_shapes = 0
            for el in title_layout.shapes._spTree:
                tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
                if tag not in ("sp", "grpSp"):
                    continue
                name = ""
                for c in el.iter():
                    if c.tag.split("}")[-1] == "cNvPr":
                        name = c.get("name", "")
                        break
                if "Logo" in name:
                    logo_shapes = len(list(el))
                    continue
                if el.find(f".//{{{NS_P}}}ph") is not None:
                    continue
                for srgb in el.findall(f".//{{{NS_A}}}srgbClr"):
                    title_colors.add(srgb.get("val", ""))
                for scheme in el.findall(f".//{{{NS_A}}}schemeClr"):
                    title_colors.add("s:" + scheme.get("val", "?"))

            results[uid] = {
                "theme":       classify_theme_from_colors(title_colors),
                "logo_shapes": logo_shapes,
                "path":        str(fname),
            }
        except Exception:
            continue

    return results


# ---------------------------------------------------------------------------
# CLI helpers (called from __main__.py)
# ---------------------------------------------------------------------------

def cmd_discover(args=None):
    """Print a grouped overview of all discovered BinaryFile templates."""
    print("Scanning empower BinaryFiles directory...")
    templates = discover_templates()
    if not templates:
        print(f"No templates found in: {_binary_dir()}")
        return

    registry = load_registry()
    # Build reverse map: uid -> (division, theme)
    registered = {}
    for div, themes in registry.items():
        for theme, uid in (themes or {}).items():
            registered[uid] = (div, theme)

    from collections import defaultdict
    by_logo: dict = defaultdict(lambda: defaultdict(list))
    for uid, info in templates.items():
        key = info["logo_shapes"]
        by_logo[key][info["theme"]].append(uid)

    print(f"\nFound {len(templates)} full-featured templates.\n")

    for logo_sz in sorted(by_logo.keys(), reverse=True):
        if logo_sz == 4:
            logo_label = "standard logo (4 shapes)"
        elif logo_sz == 0:
            logo_label = "logo in slide master"
        else:
            logo_label = f"special logo ({logo_sz} shapes)"
        print(f"=== Group: {logo_label} ===")
        for theme in sorted(by_logo[logo_sz].keys()):
            uids = by_logo[logo_sz][theme]
            mapped = [u for u in uids if u in registered]
            unmapped = [u for u in uids if u not in registered]
            status = f"{len(mapped)} mapped" + (
                f", {len(unmapped)} UNMAPPED" if unmapped else ""
            )
            print(f"  {theme:<12} ({len(uids)} files) -> {status}")
            for u in mapped:
                div, _ = registered[u]
                print(f"    {u}  [division: {div}]")
            for u in unmapped:
                print(f"    {u}  [unregistered]")
        print()

    print("To register a template:")
    print("  python -m merck_pptx register-template <uid> <division> <color_theme>")
    print(f"\nRegistry file: {_REGISTRY_PATH}")


def cmd_register(uid: str, division: str, color_theme: str):
    """Register a UID in the registry under (division, color_theme)."""
    path = uid_to_path(uid)
    if path is None:
        print(f"ERROR: BinaryFile not found: {_binary_dir() / uid}.pptx")
        return

    registry = load_registry()
    div = division.lower().strip().replace("-", "_").replace(" ", "_")
    theme = color_theme.lower().strip()

    if div not in registry:
        registry[div] = {}
    registry[div][theme] = uid
    save_registry(registry)
    print(f"Registered: division={div!r}, color_theme={theme!r} -> {uid}")
    print(f"Registry saved to {_REGISTRY_PATH}")
