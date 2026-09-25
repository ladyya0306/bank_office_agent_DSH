"""Allow `python -m office_kit <command>` from the package's parent directory."""
from __future__ import annotations

import sys

from .cli import main

if __name__ == "__main__":
    sys.exit(main())
