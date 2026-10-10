"""Unit tests for main.py.

All files are written to temporary folders; the project's Input and Output folders are not touched.
The command-line tests run a copy of the Code modules in a temporary project folder, because
main.py always uses the Input and Output folders next to its own Code folder.

Run from the project folder with either
    python -m unittest discover -s Code -v
    python -m pytest Code
"""

import contextlib
import io
import math
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import main
from calculation import SheetInputs, SheetResults, calculate_sheet
from data_io import CURVE_COLUMNS, CURVE_FILE, OUTPUT_FILES, PARAMETERS_FILE, save_table
from main import print_summary, run
from tests.test_calculation import example_inputs


def write_inputs(folder: Path, inputs: SheetInputs) -> None:
    """Write ``inputs`` as curve.csv and parameters.csv to ``folder`` (Coupon Frequency only if it is set)."""
    save_table(folder, CURVE_FILE, CURVE_COLUMNS, [
        [t, inputs["dlt"].get(t, 0), inputs["LLFRweightsIn"].get(t, 0.0), inputs["Rate"].get(t, math.nan)]
        for t in range(1, inputs["MAX_MATURITY"] + 1)])
    save_table(folder, PARAMETERS_FILE, ["Parameter", "Value"], [
        ["Instrument", inputs["Instrument"]],
        *([["Coupon Frequency", inputs["CouponFreq"]]] if inputs["CouponFreq"] is not None else []),
        ["UFR", inputs["UFR"]],
        ["Convergence", inputs["alpha"]],
        ["CRA", inputs["CRA"]],
        ["Max Maturity", inputs["MAX_MATURITY"]],
    ])


def printed(function: object, *args: object) -> tuple[object, str]:
    """Call ``function(*args)`` and return its result and everything it printed."""
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer):
        result = function(*args)  # type: ignore[operator]
    return result, buffer.getvalue()


def summary_lines(output: str) -> list[tuple[int, str, str]]:
    """The maturity rows of print_summary's output, in printed order, as (maturity, bootstrapped, basic RFR)."""
    lines: list[tuple[int, str, str]] = []
    for line in output.splitlines():
        parts = line.split()
        if len(parts) == 3 and parts[0].isdigit():
            lines.append((int(parts[0]), parts[1], parts[2]))
    return lines


def summary_maturities(output: str) -> list[int]:
    """The printed maturities, in order and including any repeats."""
    return [t for t, _, _ in summary_lines(output)]


def summary_rows(output: str) -> dict[int, list[str]]:
    """The maturity rows as {maturity: [bootstrapped, basic RFR]}."""
    return {t: [boot, basic] for t, boot, basic in summary_lines(output)}


class TestPrintSummary(unittest.TestCase):

    def setUp(self) -> None:
        self.res: SheetResults = calculate_sheet(**example_inputs())
        _, self.output = printed(print_summary, self.res)

    def test_fsp_and_llfr(self) -> None:
        lines = self.output.splitlines()
        self.assertEqual(lines[0], "FSP  : 20 years")
        self.assertEqual(lines[1], "LLFR : 3.22488713% (continuously compounded)")

    def test_selected_maturities(self) -> None:
        self.assertEqual(summary_maturities(self.output),
                         [1, 2, 5, 10, 15, 20, 21, 25, 30, 40, 50, 60, 80, 100, 150])

    def test_values_are_formatted_as_percentages(self) -> None:
        rows = summary_rows(self.output)
        self.assertEqual(rows[20], ["3.15865%", "3.209%"])
        self.assertEqual(rows[150][1], "3.287%")

    def test_missing_bootstrapped_rate_is_shown_as_na(self) -> None:
        # the bootstrapped curve ends at the last DLT maturity (50)
        rows = summary_rows(self.output)
        self.assertNotEqual(rows[50][0], "#N/A")
        self.assertEqual(rows[60][0], "#N/A")

    def test_shorter_curve_ends_at_its_max_maturity(self) -> None:
        # 55 keeps the example valid: its DLT points, weights and rates end at 50
        inputs = example_inputs()
        inputs["MAX_MATURITY"] = 55
        _, output = printed(print_summary, calculate_sheet(**inputs))
        self.assertEqual(summary_maturities(output), [1, 2, 5, 10, 15, 20, 21, 25, 30, 40, 50, 55])

    def test_max_maturity_on_the_list_is_shown_once(self) -> None:
        inputs = example_inputs()
        inputs["MAX_MATURITY"] = 100
        _, output = printed(print_summary, calculate_sheet(**inputs))
        self.assertEqual(summary_maturities(output)[-2:], [80, 100])
        self.assertEqual(summary_maturities(output).count(100), 1)


class TestRun(unittest.TestCase):

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.input_dir: Path = Path(self._tmp.name) / "Input"
        self.output_dir: Path = Path(self._tmp.name) / "Output"
        write_inputs(self.input_dir, example_inputs())

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_returns_the_calculated_results(self) -> None:
        res, _ = printed(run, self.input_dir, self.output_dir)
        expected: SheetResults = calculate_sheet(**example_inputs())
        assert isinstance(res, dict)
        self.assertEqual((res["FSP"], res["LLFR"], res["BASIC_RFR"]),
                         (expected["FSP"], expected["LLFR"], expected["BASIC_RFR"]))

    def test_saves_the_results(self) -> None:
        printed(run, self.input_dir, self.output_dir)
        self.assertEqual(sorted(p.name for p in self.output_dir.iterdir()), sorted(OUTPUT_FILES["SWP"]))

    def test_prints_the_summary_and_the_saved_files(self) -> None:
        _, output = printed(run, self.input_dir, self.output_dir)
        self.assertTrue(output.startswith("Instrument: SWP (swaps)\nFSP  : 20 years\n"))
        for name in OUTPUT_FILES["SWP"]:
            self.assertIn(f"Saved {self.output_dir / name}\n", output)

    def test_government_bonds(self) -> None:
        write_inputs(self.input_dir, example_inputs("GVT"))
        self.assertFalse("Coupon Frequency" in (self.input_dir / PARAMETERS_FILE).read_text(encoding="utf-8"))
        res, output = printed(run, self.input_dir, self.output_dir)
        expected: SheetResults = calculate_sheet(**example_inputs("GVT"))
        assert isinstance(res, dict)
        self.assertEqual((res["FSP"], res["LLFR"], res["BASIC_RFR"]),
                         (expected["FSP"], expected["LLFR"], expected["BASIC_RFR"]))
        self.assertTrue(output.startswith("Instrument: GVT (government bonds)\nFSP  : 20 years\n"))
        self.assertEqual(sorted(p.name for p in self.output_dir.iterdir()), sorted(OUTPUT_FILES["GVT"]))

    def test_invalid_inputs_stop_before_anything_is_printed_or_saved(self) -> None:
        inputs = example_inputs()
        inputs["UFR"] = 3.3
        write_inputs(self.input_dir, inputs)
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer), self.assertRaisesRegex(ValueError, "UFR must be a decimal"):
            run(self.input_dir, self.output_dir)
        self.assertEqual(buffer.getvalue(), "")
        self.assertFalse(self.output_dir.exists())

    def test_unreadable_inputs_stop_before_anything_is_saved(self) -> None:
        (self.input_dir / PARAMETERS_FILE).write_text("Parameter,Value\nUFR,abc\n", encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "missing parameter"):
            run(self.input_dir, self.output_dir)
        self.assertFalse(self.output_dir.exists())

    def test_missing_input_folder(self) -> None:
        with self.assertRaises(FileNotFoundError):
            run(Path(self._tmp.name) / "NoSuchFolder", self.output_dir)
        self.assertFalse(self.output_dir.exists())

    def test_default_folders_are_the_project_folders(self) -> None:
        code_dir: Path = Path(main.__file__).resolve().parent
        self.assertEqual(run.__defaults__, (code_dir.parent / "Input", code_dir.parent / "Output"))


class TestCommandLine(unittest.TestCase):
    """python Code/main.py, run on a copy of the modules in a temporary project folder."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.project: Path = Path(self._tmp.name)
        code_copy: Path = self.project / "Code"
        code_copy.mkdir()
        for module in Path(main.__file__).resolve().parent.glob("*.py"):
            shutil.copy(module, code_copy / module.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def run_main(self) -> "subprocess.CompletedProcess[str]":
        # started from another folder, to show the paths do not depend on the working directory
        return subprocess.run([sys.executable, str(self.project / "Code" / "main.py")], cwd=self._tmp.name,
                              capture_output=True, text=True, timeout=120)

    def test_success(self) -> None:
        write_inputs(self.project / "Input", example_inputs())
        process = self.run_main()
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertIn("FSP  : 20 years", process.stdout)
        for name in OUTPUT_FILES["SWP"]:
            self.assertTrue((self.project / "Output" / name).exists())

    def test_success_government_bonds(self) -> None:
        write_inputs(self.project / "Input", example_inputs("GVT"))
        process = self.run_main()
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertIn("Instrument: GVT (government bonds)", process.stdout)
        self.assertIn("LLFR : 3.18830280% (continuously compounded)", process.stdout)
        for name in OUTPUT_FILES["GVT"]:
            self.assertTrue((self.project / "Output" / name).exists())

    def test_invalid_inputs_give_an_input_error(self) -> None:
        inputs = example_inputs()
        inputs["UFR"] = 3.3
        write_inputs(self.project / "Input", inputs)
        process = self.run_main()
        self.assertEqual(process.returncode, 1)
        self.assertTrue(process.stderr.startswith("Input error: invalid inputs"), process.stderr)
        self.assertIn("UFR must be a decimal", process.stderr)
        self.assertNotIn("Traceback", process.stderr)
        self.assertFalse((self.project / "Output").exists())

    def test_missing_inputs_give_an_input_error(self) -> None:
        process = self.run_main()
        self.assertEqual(process.returncode, 1)
        self.assertTrue(process.stderr.startswith("Input error: Input file not found"), process.stderr)
        self.assertNotIn("Traceback", process.stderr)


if __name__ == "__main__":
    unittest.main()
