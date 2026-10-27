"""Shared by the package's tests: the test data, and running Python in a new process with the package."""

import os
import subprocess
import sys
from pathlib import Path

import eiopa_rfr

DATA_DIR: Path = Path(__file__).resolve().parent / "data"
# Examples/EIOPA_Example/Input, copied by sync_code.py: the workbook's example (GVT, with VA and shocks)
EXAMPLE_INPUT_DIR: Path = DATA_DIR / "eiopa_example"
# Code/tests/irr_workbook_values.csv, copied by sync_code.py: the workbook's shocked curves for the example
WORKBOOK_VALUES_FILE: Path = DATA_DIR / "irr_workbook_values.csv"
README_FILE: Path = Path(__file__).resolve().parents[1] / "README.md"


def run_python(*args: str, cwd: Path) -> "subprocess.CompletedProcess[str]":
    """Run python with the arguments in a new process that imports the same package as the tests.

    Args:
        args: the arguments of python, e.g. "-m", "eiopa_rfr", "--version".
        cwd: the working folder of the process.

    Returns:
        The finished process, with its output as text.
    """
    env: dict[str, str] = dict(os.environ)
    paths: list[str] = [str(Path(eiopa_rfr.__file__).resolve().parents[1]), env.get("PYTHONPATH", "")]
    env["PYTHONPATH"] = os.pathsep.join(path for path in paths if path)
    return subprocess.run([sys.executable, *args], cwd=cwd, env=env, capture_output=True, text=True)
