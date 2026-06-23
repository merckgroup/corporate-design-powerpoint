"""
CLI entry point: python -m merck_pptx <command> [args]

Commands:
  generate <source> <output.pptx> [--meta meta.json | --defaults] [--save-plan plan.json]
      Convert a .md or .pptx file to a Merck deck using Claude.
      --meta      Path to a meta JSON file (skips all prompts).
      --defaults  Use built-in default meta values (skips all prompts).
                  Also activates automatically when stdin is not a TTY (e.g. pipes, CI).

  build <plan.json> <output.pptx>
      Build a Merck deck from a pre-built plan JSON. No LLM.
"""

import argparse
import json
import sys
from datetime import datetime


def _ask_meta() -> dict:
    """Interactively collect the 6 gate answers and return a meta dict."""
    print("\nMerck PPTX Pipeline — deck setup\n")

    region = input("[1/6] Region (EU / USA) [EU]: ").strip() or "EU"
    if region.upper() not in ("EU", "USA"):
        print(f"Unknown region '{region}'. Defaulting to EU.")
        region = "EU"
    region = region.upper()

    classification = input("[2/6] Classification (Public / Internal / Confidential) [Internal]: ").strip() or "Internal"
    if classification.lower() == "secret":
        print("ERROR: Secret-classified decks cannot be generated. Exiting.")
        sys.exit(1)
    # Normalise to title-case before the membership check so "confidential" == "Confidential"
    classification = classification.strip().title()
    if classification not in ("Public", "Internal", "Confidential"):
        print(f"Unknown classification '{classification}'. Defaulting to Internal.")
        classification = "Internal"

    default_date = datetime.now().strftime("%B %Y")
    month_year = input(f"[3/6] Month and year [{default_date}]: ").strip() or default_date

    print("[4/6] Audience:")
    audiences = [
        "Executive leadership",
        "Senior management",
        "Functional team",
        "Mixed audience",
        "External / client-facing",
    ]
    for i, a in enumerate(audiences, 1):
        print(f"      {i}. {a}")
    aud_choice = input("      Choice [1]: ").strip() or "1"
    try:
        audience = audiences[int(aud_choice) - 1]
    except (ValueError, IndexError):
        audience = audiences[0]

    _style_default = {
        "Executive leadership": "executive",
        "Senior management":    "executive",
    }.get(audience, "corporate")
    style_input = input(
        f"[5/6] Visual style (executive / corporate / storytelling / science) [{_style_default}]: "
    ).strip().lower() or _style_default
    style_map = {
        "executive":    "merck_executive",
        "corporate":    "merck_corporate",
        "storytelling": "merck_storytelling",
        "science":      "merck_science",
    }
    deck_style = style_map.get(style_input, "merck_executive")

    variety = input("[6/6] Variety mode (default / creative) [default]: ").strip().lower() or "default"
    if variety not in ("default", "creative"):
        variety = "default"

    deck_label = input("\nDeck label / title (appears in footer) [Merck Presentation]: ").strip() or "Merck Presentation"

    return {
        "region":         region,
        "deck_label":     deck_label,
        "classification": classification,
        "month_year":     month_year,
        "audience":       audience,
        "deck_style":     deck_style,
        "variety_mode":   variety,
        "show_disclaimer": False,
    }


_VALID_META_REGIONS          = {"EU", "USA"}
_VALID_META_CLASSIFICATIONS  = {"Public", "Internal", "Confidential"}
# Use upper-case so that .upper() comparison works for all variants, including
# 'TS/SCI' which .title() would render as 'Ts/Sci' (slash resets capitalisation).
_FORBIDDEN_META_CLASSIFICATIONS = {"SECRET", "TOP SECRET", "TS/SCI"}


def _default_meta() -> dict:
    """Return meta dict with built-in defaults — no prompts required."""
    return {
        "region":          "EU",
        "deck_label":      "Merck Presentation",
        "classification":  "Internal",
        "month_year":      datetime.now().strftime("%B %Y"),
        "audience":        "Mixed audience",
        "deck_style":      "merck_corporate",
        "color_theme":     "organic",
        "variety_mode":    "default",
        "show_disclaimer": False,
    }


def _validate_meta_dict(meta: dict) -> None:
    """Validate a meta dict loaded from --meta file. Raises ValueError on bad values."""
    region = str(meta.get("region", "")).strip().upper()
    # Empty or missing region is an error — silently accepting it causes the wrong
    # template (and legal disclaimer) to be selected.
    if region not in _VALID_META_REGIONS:
        raise ValueError(
            f"meta.region must be EU or USA, got: {region!r}"
        )
    raw_cls = str(meta.get("classification", "")).strip()
    cls_upper = raw_cls.upper()
    if cls_upper in _FORBIDDEN_META_CLASSIFICATIONS:
        print(
            f"ERROR: Classification '{raw_cls}' is not permitted. Exiting.",
            file=sys.stderr,
        )
        sys.exit(1)
    # Empty or missing classification is also an error.
    cls_title = raw_cls.title()
    if cls_title not in _VALID_META_CLASSIFICATIONS:
        raise ValueError(
            f"meta.classification must be Public, Internal, or Confidential, got: {raw_cls!r}"
        )


def _cmd_list_templates():
    """Print a catalogue of all available Merck PPTX templates from every source."""
    import pathlib
    from merck_pptx.build_from_plan import _TEMPLATE_DIR
    from merck_pptx.manual_templates import VALID_THEMES as _ALL_THEMES

    _SEP = "-" * 60

    print()
    print("  Available Merck PPTX templates")
    print(_SEP)

    # ------------------------------------------------------------------ bundled
    print()
    print("  [1] Bundled templates  (always available, no download needed)")
    eu_tmpl  = _TEMPLATE_DIR / "EU_Merck_Themed.pptx"
    usa_tmpl = _TEMPLATE_DIR / "USA_Merck_Themed_Base_v1.pptx"
    _eu_ok  = "[OK]" if eu_tmpl.exists()  else "[MISSING]"
    _usa_ok = "[OK]" if usa_tmpl.exists() else "[MISSING]"
    print(f"    {_eu_ok}  eu  / merck  ->  {eu_tmpl.name}  (all 6 color themes)")
    print(f"    {_usa_ok}  usa / merck  ->  {usa_tmpl.name}  (all 6 color themes)")
    print("       Note: color themes are applied programmatically -- no separate")
    print("             file needed per theme for the 'merck' division.")

    # --------------------------------------------------------- empower BinaryFiles
    print()
    print("  [2] empower BinaryFiles  (Windows, empower add-in required)")
    try:
        from merck_pptx.binary_templates import load_registry, uid_to_path, _binary_dir
        bdir = pathlib.Path(_binary_dir())
        registry = load_registry()
        if not bdir.exists():
            print(f"    [NOT FOUND]  empower cache directory not found:")
            print(f"                 {bdir}")
            print("    Install empower or use Option B (manual download) below.")
        else:
            total_empower = 0
            for div, themes in sorted(registry.items()):
                if div.startswith("_"):
                    continue
                if not isinstance(themes, dict):
                    continue
                present = []
                missing = []
                for theme in _ALL_THEMES:
                    uid = themes.get(theme)
                    if uid:
                        p = uid_to_path(uid)
                        if p is not None:
                            present.append(theme)
                        else:
                            missing.append(theme)
                if present:
                    print(f"    [OK]  {div:<20}  {', '.join(present)}")
                    total_empower += len(present)
                if missing:
                    print(f"    [REG/MISSING] {div:<18}  {', '.join(missing)}  "
                          "(registered but file absent -- reinstall empower?)")
            if total_empower == 0:
                print("    (registry is empty -- run 'discover-templates' to scan)")
            else:
                print(f"\n    {total_empower} empower template(s) available")
    except Exception as exc:
        print(f"    [ERROR] Could not scan empower BinaryFiles: {exc}")

    # ---------------------------------------------------------- manual templates
    print()
    print("  [3] Manual templates  (~/.merck_pptx/templates/)")
    try:
        from merck_pptx.manual_templates import manual_templates_dir, list_manual_templates
        mdir = manual_templates_dir()
        manual = list_manual_templates()
        print(f"    Directory: {mdir}")
        if not mdir.exists():
            print("    [EMPTY]  Directory does not exist yet.")
        elif not manual:
            print("    [EMPTY]  No recognised template files found.")
        else:
            for entry in manual:
                print(f"    [OK]  {entry['division']:<20}  {entry['color_theme']}")
            print(f"\n    {len(manual)} manual template(s) available")
        print()
        print("    To add templates: place .pptx files named {division}_{color_theme}.pptx")
        print(f"    in the directory above.  Example: emd_serono_organic.pptx")
        print("    Then re-run this command to confirm they are recognised.")
    except Exception as exc:
        print(f"    [ERROR] Could not scan manual templates: {exc}")

    print()
    print(_SEP)
    print("  Run 'python setup_merck_pptx.py' to auto-detect empower and")
    print("  get full instructions for downloading templates manually.")
    print()


def main():
    parser = argparse.ArgumentParser(
        prog="python -m merck_pptx",
        description="Merck PPTX Pipeline",
    )
    sub = parser.add_subparsers(dest="command")

    # generate command
    gen = sub.add_parser("generate", help="Convert .md or .pptx to Merck deck (uses Claude)")
    gen.add_argument("source", help="Source file (.md or .pptx)")
    gen.add_argument("output", help="Output .pptx path")
    gen.add_argument("--meta", help="Path to meta JSON file (skips interactive prompts)")
    gen.add_argument("--defaults", action="store_true",
                     help="Use built-in default meta values without prompts")
    gen.add_argument("--save-plan", metavar="PLAN_JSON",
                     help="Save the LLM-generated slide plan to this JSON file")

    # build command
    bld = sub.add_parser("build", help="Build from a plan JSON (no LLM)")
    bld.add_argument("plan", help="Path to plan .json file")
    bld.add_argument("output", help="Output .pptx path")

    # discover-templates command
    sub.add_parser(
        "discover-templates",
        help="Scan empower BinaryFiles and show all available themed templates",
    )

    # list-templates command
    sub.add_parser(
        "list-templates",
        help="Show all available Merck templates (bundled, empower, and manual)",
    )

    # register-template command
    reg = sub.add_parser(
        "register-template",
        help="Register an empower BinaryFile UID in the template registry",
    )
    reg.add_argument("uid",         help="BinaryFile UID (without .pptx extension)")
    reg.add_argument("division",    help="Division key, e.g. merck, emd_serono, millipore_sigma")
    reg.add_argument("color_theme", help="Theme key: plastic | functional | organic | synthetic | technical | electronics")

    args = parser.parse_args()

    if args.command == "build":
        from merck_pptx.build_from_plan import build_from_plan, TemplateNotFoundError
        from merck_pptx.validate_plan import ValidationError
        try:
            out = build_from_plan(args.plan, args.output)
            print(f"Saved: {out}")
        except ValidationError as e:
            print(f"Validation error: {e}", file=sys.stderr)
            sys.exit(2)
        except TemplateNotFoundError as e:
            print(f"Template error: {e}", file=sys.stderr)
            sys.exit(5)

    elif args.command == "generate":
        from merck_pptx.generate import generate_deck
        from merck_pptx.build_from_plan import TemplateNotFoundError
        from merck_pptx.validate_plan import ValidationError

        if args.meta:
            with open(args.meta, encoding="utf-8") as f:
                meta = json.load(f)
            # Security: validate meta fields on load so a malicious or
            # misconfigured meta.json cannot bypass classification guards.
            _validate_meta_dict(meta)
        elif args.defaults or not (sys.stdin is not None and sys.stdin.isatty()):
            # --defaults flag, or non-TTY stdin (CI, pipes, pythonw, PyInstaller).
            # Guard sys.stdin is not None: in some embedded runtimes stdin may be None,
            # in which case .isatty() would raise AttributeError.
            meta = _default_meta()
        else:
            meta = _ask_meta()
            print()

        try:
            out = generate_deck(
                args.source, args.output, meta=meta,
                save_plan=args.save_plan,
            )
            print(f"Saved: {out}")
        except ValidationError as e:
            print(f"Validation error: {e}", file=sys.stderr)
            sys.exit(2)
        except TemplateNotFoundError as e:
            print(f"Template error: {e}", file=sys.stderr)
            sys.exit(5)
        except EnvironmentError as e:
            print(f"Configuration error: {e}", file=sys.stderr)
            sys.exit(3)
        except ValueError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(4)

    elif args.command == "discover-templates":
        from merck_pptx.binary_templates import cmd_discover
        cmd_discover()

    elif args.command == "list-templates":
        _cmd_list_templates()

    elif args.command == "register-template":
        from merck_pptx.binary_templates import cmd_register
        cmd_register(args.uid, args.division, args.color_theme)

    else:
        parser.print_help()
        sys.exit(0)


if __name__ == "__main__":
    main()
