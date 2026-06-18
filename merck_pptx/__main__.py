"""
CLI entry point: python -m merck_pptx <command> [args]

Commands:
  generate <source> <output.pptx> [--meta meta.json]
      Convert a .md or .pptx file to a Merck deck using Claude.
      If --meta is omitted, 6 gate questions are asked interactively.

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
        f"[5/6] Visual style (executive / corporate / storytelling) [{_style_default}]: "
    ).strip().lower() or _style_default
    style_map = {
        "executive":    "merck_executive",
        "corporate":    "merck_corporate",
        "storytelling": "merck_storytelling",
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
_FORBIDDEN_META_CLASSIFICATIONS = {"Secret", "Top Secret", "TS/SCI"}


def _validate_meta_dict(meta: dict) -> None:
    """Validate a meta dict loaded from --meta file. Raises ValueError on bad values."""
    region = str(meta.get("region", "")).strip().upper()
    if region and region not in _VALID_META_REGIONS:
        raise ValueError(
            f"meta.region must be EU or USA, got: {region!r}"
        )
    raw_cls = str(meta.get("classification", "")).strip()
    cls_title = raw_cls.title()
    if cls_title in _FORBIDDEN_META_CLASSIFICATIONS:
        print(
            f"ERROR: Classification '{raw_cls}' is not permitted. Exiting.",
            file=sys.stderr,
        )
        sys.exit(1)
    if raw_cls and cls_title not in _VALID_META_CLASSIFICATIONS:
        raise ValueError(
            f"meta.classification must be Public, Internal, or Confidential, got: {raw_cls!r}"
        )


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
    gen.add_argument("--meta", help="Path to meta JSON file (if omitted, interactive prompts)")

    # build command
    bld = sub.add_parser("build", help="Build from a plan JSON (no LLM)")
    bld.add_argument("plan", help="Path to plan .json file")
    bld.add_argument("output", help="Output .pptx path")

    args = parser.parse_args()

    if args.command == "build":
        from merck_pptx.build_from_plan import build_from_plan
        from merck_pptx.validate_plan import ValidationError
        try:
            out = build_from_plan(args.plan, args.output)
            print(f"Saved: {out}")
        except ValidationError as e:
            print(f"Validation error: {e}", file=sys.stderr)
            sys.exit(2)

    elif args.command == "generate":
        from merck_pptx.generate import generate_deck
        from merck_pptx.validate_plan import ValidationError

        if args.meta:
            with open(args.meta, encoding="utf-8") as f:
                meta = json.load(f)
            # Security: validate meta fields on load so a malicious or
            # misconfigured meta.json cannot bypass classification guards.
            _validate_meta_dict(meta)
        else:
            meta = _ask_meta()
            print()

        try:
            out = generate_deck(args.source, args.output, meta=meta)
            print(f"Saved: {out}")
        except ValidationError as e:
            print(f"Validation error: {e}", file=sys.stderr)
            sys.exit(2)
        except EnvironmentError as e:
            print(f"Configuration error: {e}", file=sys.stderr)
            sys.exit(3)
        except ValueError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(4)

    else:
        parser.print_help()
        sys.exit(0)


if __name__ == "__main__":
    main()
