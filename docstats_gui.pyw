"""Launch the GUI (double-click this file; .pyw keeps the console window hidden)."""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docstats.gui import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
