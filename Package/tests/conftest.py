"""Test setup: test the installed package if there is one, otherwise the one in src.

Run from the repository folder with
    python -m pytest Package
The calculation itself is tested by the tests of the folder Code (python -m pytest Code); these
tests cover what the package adds: the generated copy, the public names, the command line and the
examples of README.md.
"""

import importlib.util
import sys
from pathlib import Path

if importlib.util.find_spec("eiopa_rfr") is None:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))


def pytest_report_header() -> str:
    """Show which copy of the package is tested."""
    import eiopa_rfr
    return f"eiopa_rfr {eiopa_rfr.__version__} from {Path(eiopa_rfr.__file__).parent}"
