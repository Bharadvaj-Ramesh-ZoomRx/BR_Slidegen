"""
Allow running slidegen as a package: python -m slidegen

Dispatches to subcommands: create, edit, reconcile.
"""

import sys


def main():
    usage = (
        "Usage: python -m slidegen <command> [args]\n"
        "\n"
        "Commands:\n"
        "  create        Create a demo slide (or import SlideBuilder for custom)\n"
        "  edit          Interactive live editor (requires open PowerPoint)\n"
        "  reconcile     Sync registry from live PowerPoint state\n"
        "  fetch-synapse Fetch banner plan data from Synapse API → source_data.xlsx\n"
        "  fetch-raw     Fetch raw respondent data from Synapse API → source_raw_data.xlsx\n"
        "\n"
        "Examples:\n"
        "  python -m slidegen create\n"
        "  python -m slidegen edit demo_slide.pptx\n"
        "  python -m slidegen reconcile demo_slide.pptx\n"
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
