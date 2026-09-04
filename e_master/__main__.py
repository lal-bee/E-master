"""支持 python -m e_master。"""

from __future__ import annotations

import sys

from e_master.cli import main

if __name__ == "__main__":
    sys.exit(main())
