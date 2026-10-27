"""Tests of the command eiopa-rfr (cli.py and __main__.py) on the EIOPA example.

All files are written to temporary folders.
"""

import contextlib
import csv
import io
import shutil
import tempfile
import unittest
from pathlib import Path

import eiopa_rfr
from eiopa_rfr.cli import main
from support import EXAMPLE_INPUT_DIR, WORKBOOK_VALUES_FILE, run_python

# the results of the EIOPA example: government bonds, with VA and shocks
RESULT_FILES: list[str] = ["government_bonds_curves.csv", "government_bonds_irr_curves.csv",
                           "government_bonds_irr_parameters.csv", "government_bonds_parameters.csv",
                           "government_bonds_va_curves.csv", "government_bonds_va_parameters.csv"]


def read_rows(path: Path) -> list[dict[str, str]]:
    """The rows of a result file."""
    with path.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def read_parameters(path: Path) -> dict[str, str]:
    """A parameters result file as {Parameter: Value}."""
    return {row["Parameter"]: row["Value"] for row in read_rows(path)}


class TestCommandLine(unittest.TestCase):

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.folder: Path = Path(self._tmp.name)
        self.output: Path = self.folder / "Output"

    def run_cli(self, *args: str) -> str:
        """Run the command with the arguments in this process; return what it printed."""
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            main(list(args))
        return buffer.getvalue()

    def run_example(self, output: Path | None = None) -> str:
        """Run the command on the EIOPA example."""
        return self.run_cli("--input", str(EXAMPLE_INPUT_DIR), "--output", str(output or self.output))

    def test_writes_every_result(self) -> None:
        printed: str = self.run_example()
        self.assertEqual(sorted(path.name for path in self.output.iterdir()), RESULT_FILES)
        self.assertTrue(printed.startswith("Instrument: GVT (government bonds)\nFSP  : 20 years\n"))

    def test_results_match_the_workbook(self) -> None:
        self.run_example()
        parameters: dict[str, str] = read_parameters(self.output / "government_bonds_parameters.csv")
        self.assertEqual((parameters["FSP"], float(parameters["LLFR (CC)"])), ("20", 0.03188302797422029))
        va: dict[str, str] = read_parameters(self.output / "government_bonds_va_parameters.csv")
        self.assertEqual((float(va["CSSR"]), float(va["Total VA"])), (0.5, 0.0034))
        # the shocked curves at all 150 maturities, exactly as Excel calculates them; #N/A as empty cells
        curves: list[dict[str, str]] = read_rows(self.output / "government_bonds_irr_curves.csv")
        workbook: list[dict[str, str]] = read_rows(WORKBOOK_VALUES_FILE)
        self.assertEqual(len(curves), len(workbook))
        for row, expected in zip(curves, workbook):
            for column, value in expected.items():
                where: str = f"{column} at maturity {expected['Maturity']}"
                self.assertEqual(row[column] and float(row[column]), value and float(value), where)

    def test_keeps_other_files_and_deletes_earlier_results(self) -> None:
        self.output.mkdir()
        (self.output / "notes.txt").write_text("mine", encoding="utf-8")
        (self.output / "swaps_curves.csv").write_text("earlier run on swaps", encoding="utf-8")
        self.run_example()
        self.assertEqual((self.output / "notes.txt").read_text(encoding="utf-8"), "mine")
        self.assertFalse((self.output / "swaps_curves.csv").exists())

    def test_creates_the_output_folder_and_its_parents(self) -> None:
        output: Path = self.folder / "results" / "2026-05"
        self.run_example(output)
        self.assertEqual(sorted(path.name for path in output.iterdir()), RESULT_FILES)

    def test_missing_input_file(self) -> None:
        with self.assertRaises(SystemExit) as raised:
            self.run_cli("--input", str(self.folder / "nowhere"), "--output", str(self.output))
        self.assertRegex(str(raised.exception.code), r"^Input error: Input file not found: .*parameters\.csv$")

    def test_error_leaves_no_earlier_results(self) -> None:
        self.output.mkdir()
        (self.output / "government_bonds_curves.csv").write_text("earlier run", encoding="utf-8")
        with self.assertRaises(SystemExit):
            self.run_cli("--input", str(self.folder / "nowhere"), "--output", str(self.output))
        self.assertEqual(list(self.output.iterdir()), [])

    def test_output_folder_is_a_file(self) -> None:
        (self.folder / "results.txt").write_text("", encoding="utf-8")
        with self.assertRaises(SystemExit) as raised:
            self.run_example(self.folder / "results.txt")
        self.assertRegex(str(raised.exception.code), r"^Input error: the output folder .* is a file$")


class TestPythonM(unittest.TestCase):
    """python -m eiopa_rfr, in a new process."""

    def test_default_folders_are_in_the_working_folder(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            shutil.copytree(EXAMPLE_INPUT_DIR, Path(tmp) / "Input")
            process = run_python("-m", "eiopa_rfr", cwd=Path(tmp))
            self.assertEqual(process.returncode, 0, process.stderr)
            self.assertIn("FSP  : 20 years", process.stdout)
            self.assertEqual(sorted(path.name for path in (Path(tmp) / "Output").iterdir()), RESULT_FILES)

    def test_version(self) -> None:
        process = run_python("-m", "eiopa_rfr", "--version", cwd=EXAMPLE_INPUT_DIR)
        self.assertEqual(process.stdout.strip(), f"eiopa-rfr {eiopa_rfr.__version__}")

    def test_input_error_exit_code(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            process = run_python("-m", "eiopa_rfr", cwd=Path(tmp))
            self.assertEqual(process.returncode, 1)
            self.assertIn("Input error: Input file not found", process.stderr)


if __name__ == "__main__":
    unittest.main()
