"""Allow running the CLI with ``python -m pfw``."""

import sys

from pfw.cli import main

if __name__ == "__main__":
    sys.exit(main())
