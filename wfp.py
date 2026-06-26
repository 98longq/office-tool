"""Run OfficeTool directly from a source checkout."""

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from office_tool.cli import main as cli_main


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] in {"gui", "--gui"}:
        from office_tool.gui import main as gui_main

        return gui_main()
    return cli_main(argv)


if __name__ == "__main__":
    raise SystemExit(main())
