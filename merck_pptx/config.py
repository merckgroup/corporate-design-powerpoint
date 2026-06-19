"""Runtime configuration loader for the Merck PPTX pipeline.

Config is read from (in priority order):
  1. ~/.merck_pptx/config.yaml   — user-level config (written by setup.py)
  2. <package_root>/config.yaml  — project-level override (gitignored)
  3. Built-in defaults            — sensible fallbacks, no YAML needed

Typical config.yaml:
  empower:
    binary_dir:    C:/Users/JohnDoe/AppData/Local/empower/data/empower/BinaryFiles
    thumbnail_dir: C:/Users/JohnDoe/AppData/Local/empower/data/empower/ThumbnailLarge
    enabled: true

Template variables supported in path values:
  {username}  — current OS username
  {home}      — current user home directory
  {appdata}   — %APPDATA% on Windows, ~/.config on other platforms
  {localappdata} — %LOCALAPPDATA% on Windows, ~/.local on other platforms
"""

from __future__ import annotations

import os
import pathlib
from typing import Any, Dict, Optional

# ---------------------------------------------------------------------------
# Path resolution
# ---------------------------------------------------------------------------

def _expand_path(raw: str) -> str:
    """Expand template variables and environment variables in a path string."""
    username   = os.environ.get("USERNAME") or os.environ.get("USER") or "unknown"
    home       = str(pathlib.Path.home())
    appdata    = os.environ.get("APPDATA", os.path.join(home, ".config"))
    localappdata = os.environ.get("LOCALAPPDATA", os.path.join(home, ".local"))

    result = raw
    result = result.replace("{username}",    username)
    result = result.replace("{home}",        home)
    result = result.replace("{appdata}",     appdata)
    result = result.replace("{localappdata}", localappdata)
    result = os.path.expandvars(result)   # also expand %VARNAME% / $VAR
    result = os.path.expanduser(result)   # expand ~
    return result


# ---------------------------------------------------------------------------
# Default locations
# ---------------------------------------------------------------------------

def _default_empower_root() -> str:
    """Auto-detect the empower local cache root for the current user."""
    localappdata = os.environ.get(
        "LOCALAPPDATA",
        os.path.join(str(pathlib.Path.home()), ".local")
    )
    return os.path.join(localappdata, "empower", "data", "empower")


_DEFAULTS: Dict[str, Any] = {
    "empower": {
        "enabled":       True,
        "binary_dir":    os.path.join(_default_empower_root(), "BinaryFiles"),
        "thumbnail_dir": os.path.join(_default_empower_root(), "ThumbnailLarge"),
    }
}


# ---------------------------------------------------------------------------
# Config loader
# ---------------------------------------------------------------------------

def _load_yaml(path: pathlib.Path) -> Dict[str, Any]:
    """Load a YAML file; return empty dict if absent or yaml not installed."""
    if not path.exists():
        return {}
    try:
        import yaml  # type: ignore
        with open(path, "r", encoding="utf-8") as fh:
            data = yaml.safe_load(fh) or {}
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _deep_merge(base: dict, override: dict) -> dict:
    """Recursively merge override into base (override wins)."""
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


def _load_config() -> Dict[str, Any]:
    """Load and merge config from all sources.

    Priority (highest wins):
      1. <project_root>/config.yaml  — written by setup_merck_pptx.py, gitignored
      2. ~/.merck_pptx/config.yaml   — fallback user-level config
      3. Built-in defaults
    """
    cfg = dict(_DEFAULTS)

    # Fallback: user home directory
    user_cfg = pathlib.Path.home() / ".merck_pptx" / "config.yaml"
    cfg = _deep_merge(cfg, _load_yaml(user_cfg))

    # Primary: project root (gitignored, written by setup_merck_pptx.py)
    project_cfg = pathlib.Path(__file__).parent.parent / "config.yaml"
    cfg = _deep_merge(cfg, _load_yaml(project_cfg))

    # Expand template variables in all string values
    def _expand_dict(d: dict) -> dict:
        out = {}
        for k, v in d.items():
            if isinstance(v, str):
                out[k] = _expand_path(v)
            elif isinstance(v, dict):
                out[k] = _expand_dict(v)
            else:
                out[k] = v
        return out

    return _expand_dict(cfg)


# Module-level singleton — loaded once at import time
_config: Optional[Dict[str, Any]] = None


def get_config() -> Dict[str, Any]:
    """Return the merged runtime configuration (cached after first call)."""
    global _config
    if _config is None:
        _config = _load_config()
    return _config


def get(section: str, key: str, default: Any = None) -> Any:
    """Convenience: get config().get(section, {}).get(key, default)."""
    return get_config().get(section, {}).get(key, default)


# ---------------------------------------------------------------------------
# Convenience accessors
# ---------------------------------------------------------------------------

def empower_binary_dir() -> str:
    """Return the empower BinaryFiles directory path."""
    return get("empower", "binary_dir", _DEFAULTS["empower"]["binary_dir"])


def empower_thumbnail_dir() -> str:
    """Return the empower ThumbnailLarge directory path."""
    return get("empower", "thumbnail_dir", _DEFAULTS["empower"]["thumbnail_dir"])


def empower_enabled() -> bool:
    """Return True if empower integration is enabled in config.

    Handles both dict form (empower: {enabled: false}) and
    scalar form (empower: false) gracefully.
    """
    v = get_config().get("empower")
    if isinstance(v, bool):
        return v
    return bool(get("empower", "enabled", True))
