"""
Allow running slidegen as a package: python -m slidegen

Dispatches to subcommands: create, edit, reconcile.
"""

import os
import sys


def _validate_config():
    """Validate a project config.yaml without generating a deck.

    Called after main() shifts sys.argv, so sys.argv[1] is the first arg
    after the 'validate' subcommand.
    """
    if len(sys.argv) < 2:
        print("Usage: python -m slidegen validate <config.yaml>")
        sys.exit(1)

    yaml_path = sys.argv[1]
    if not os.path.exists(yaml_path):
        print(f"Error: file not found: {yaml_path}")
        sys.exit(1)

    from slidegen.pipeline.project_config import load_project_config

    try:
        config = load_project_config(yaml_path)
    except (ValueError, KeyError) as e:
        print(f"Error loading config: {e}")
        sys.exit(1)

    errors = config.validate()

    # Summary
    print(f"Project:  {config.name}")
    print(f"Wave:     {config.wave}")
    print(f"Extractions: {len(config.extractions)}")
    print(f"Asks:     {len(config.asks)}")
    print(f"Data:     {config.data_source_path}")
    data_exists = os.path.exists(config.data_source_path)
    print(f"  exists: {'yes' if data_exists else 'NO — file not found'}")

    # Check JSON cache
    json_path = config.data_source_path.replace(".xlsx", ".json")
    context_json = os.path.join(
        os.path.dirname(yaml_path),
        config.context_path or "context",
        "source_data.json",
    ) if not os.path.exists(json_path) else json_path
    if os.path.exists(context_json):
        print(f"  JSON cache: {context_json}")
    elif os.path.exists(json_path):
        print(f"  JSON cache: {json_path}")

    if errors:
        print(f"\n{len(errors)} validation error(s):")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)
    else:
        print("\nConfig is valid.")
        sys.exit(0)


def main():
    usage = (
        "Usage: python -m slidegen <command> [args]\n"
        "\n"
        "Commands:\n"
        "  create        Create a demo slide (or import SlideBuilder for custom)\n"
        "  edit          Interactive live editor (requires open PowerPoint)\n"
        "  reconcile     Sync registry from live PowerPoint state\n"
        "  validate      Validate a project config.yaml (no deck generation)\n"
        "  fetch-synapse Fetch banner plan data from Synapse API → source_data.xlsx\n"
        "  fetch-raw     Fetch raw respondent data from Synapse API → source_raw_data.xlsx\n"
        "\n"
        "Examples:\n"
        "  python -m slidegen create\n"
        "  python -m slidegen edit demo_slide.pptx\n"
        "  python -m slidegen reconcile demo_slide.pptx\n"
        "  python -m slidegen validate projects/jnj_rybrevant/config.yaml\n"
        "  python -m slidegen fetch-synapse projects/jnj_rybrevant/config.yaml --url <url>\n"
        "  python -m slidegen fetch-raw projects/jnj_rybrevant/config.yaml\n"
    )

    if len(sys.argv) < 2:
        print(usage)
        sys.exit(1)

    cmd = sys.argv[1]
    # Shift argv so submodule sees its own args
    sys.argv = sys.argv[1:]

    if cmd == "create":
        from slidegen.create import _demo
        _demo()
    elif cmd == "edit":
        from slidegen.edit import _cli
        _cli()
    elif cmd == "reconcile":
        from slidegen.reconcile import _cli
        _cli()
    elif cmd == "validate":
        _validate_config()
    elif cmd == "fetch-synapse":
        from slidegen.pipeline.synapse_fetcher import _cli
        _cli()
    elif cmd == "fetch-raw":
        from slidegen.pipeline.synapse_raw_fetcher import _cli
        _cli()
    else:
        print(f"Unknown command: {cmd}\n")
        print(usage)
        sys.exit(1)


if __name__ == "__main__":
    main()
