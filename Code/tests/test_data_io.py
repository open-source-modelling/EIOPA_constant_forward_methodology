"""Unit tests for data_io.py.

All files are written to temporary folders; the project's Input and Output folders are not touched.

Run from the project folder with either
    python -m unittest discover -s Code -v
    python -m pytest Code
"""

import math
import tempfile
import unittest
from pathlib import Path

import data_io
from calculation import SheetResults, SwapInputs, calculate_sheet
from data_io import (CURVE_FILE, CURVES_OUTPUT_FILE, PARAMETERS_FILE, PARAMETERS_OUTPUT_FILE, csv_value,
                     parse_number, read_csv, read_inputs, save_results, save_table)
from tests.test_calculation import example_inputs

PARAMETERS_TEXT: str = "Parameter,Value\nCoupon Frequency,1\nUFR,0.033\nConvergence,0.11\nCRA,10\nMax Maturity,150\n"
CURVE_TEXT: str = ("Maturity,DLT,LLFR Weight,Input Rate\n"
                   "1,1,0,0.02176\n"
                   "2,1,0,0.022621\n"
                   "3,0,0,0.0235\n"
                   "4,0,0,\n"
                   "5,1,1,0.02569\n")


class TempFolderTestCase(unittest.TestCase):
    """Gives each test an empty temporary folder ``self.folder``."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.folder: Path = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def write(self, name: str, text: str, encoding: str = "utf-8") -> Path:
        path: Path = self.folder / name
        path.write_text(text, encoding=encoding)
        return path

    def write_inputs(self, curve: str = CURVE_TEXT, parameters: str = PARAMETERS_TEXT) -> None:
        self.write(CURVE_FILE, curve)
        self.write(PARAMETERS_FILE, parameters)


class TestFolders(unittest.TestCase):

    def test_input_and_output_are_in_the_project_folder(self) -> None:
        code_dir: Path = Path(data_io.__file__).resolve().parent
        self.assertEqual(data_io.PROJECT_DIR, code_dir.parent)
        self.assertEqual(data_io.INPUT_DIR, code_dir.parent / "Input")
        self.assertEqual(data_io.OUTPUT_DIR, code_dir.parent / "Output")


class TestReadCsv(TempFolderTestCase):

    def test_missing_file(self) -> None:
        with self.assertRaises(FileNotFoundError) as context:
            read_csv(self.folder / "missing.csv")
        self.assertIn("missing.csv", str(context.exception))

    def test_comma_separated(self) -> None:
        rows = read_csv(self.write("a.csv", "A,B\n1,2\n3,4\n"))
        self.assertEqual(rows, [{"A": "1", "B": "2"}, {"A": "3", "B": "4"}])

    def test_semicolon_separated(self) -> None:
        rows = read_csv(self.write("a.csv", "A;B\n1,5;2\n"))
        self.assertEqual(rows, [{"A": "1,5", "B": "2"}])

    def test_byte_order_mark_is_removed(self) -> None:
        # Excel's "CSV UTF-8" adds a byte order mark in front of the first column name
        rows = read_csv(self.write("a.csv", "A,B\n1,2\n", encoding="utf-8-sig"))
        self.assertEqual(list(rows[0]), ["A", "B"])

    def test_spaces_are_stripped(self) -> None:
        rows = read_csv(self.write("a.csv", " A , B \n 1 , 2 \n"))
        self.assertEqual(rows, [{"A": "1", "B": "2"}])

    def test_short_rows_get_empty_values(self) -> None:
        rows = read_csv(self.write("a.csv", "A,B,C\n1\n"))
        self.assertEqual(rows, [{"A": "1", "B": "", "C": ""}])

    def test_long_rows_are_an_error(self) -> None:
        with self.assertRaisesRegex(ValueError, "a.csv line 3: more values than column names"):
            read_csv(self.write("a.csv", "A,B\n1,2\n1,2,3\n"))

    def test_empty_file(self) -> None:
        self.assertEqual(read_csv(self.write("a.csv", "")), [])


class TestParseNumber(unittest.TestCase):

    def test_decimal_point(self) -> None:
        self.assertEqual(parse_number("0.033", "here"), 0.033)

    def test_decimal_comma(self) -> None:
        self.assertEqual(parse_number("0,033", "here"), 0.033)

    def test_negative_and_scientific(self) -> None:
        self.assertEqual(parse_number("-0.0125", "here"), -0.0125)
        self.assertEqual(parse_number("1e-3", "here"), 0.001)

    def test_not_a_number_names_the_location(self) -> None:
        with self.assertRaises(ValueError) as context:
            parse_number("abc", "file.csv line 3")
        self.assertEqual(str(context.exception), "file.csv line 3: 'abc' is not a number")


class TestReadInputs(TempFolderTestCase):

    def test_reads_parameters(self) -> None:
        self.write_inputs()
        inputs = read_inputs(self.folder)
        self.assertEqual((inputs["CouponFreq"], inputs["UFR"], inputs["alpha"], inputs["CRA"]), (1, 0.033, 0.11, 10.0))
        self.assertIsInstance(inputs["CouponFreq"], int)
        self.assertEqual(inputs["MAX_MATURITY"], 150)
        self.assertIsInstance(inputs["MAX_MATURITY"], int)

    def test_reads_curve(self) -> None:
        self.write_inputs()
        inputs = read_inputs(self.folder)
        self.assertEqual(inputs["dlt"], {1: 1, 2: 1, 3: 0, 4: 0, 5: 1})
        self.assertEqual(inputs["LLFRweightsIn"], {1: 0.0, 2: 0.0, 3: 0.0, 4: 0.0, 5: 1.0})
        # an empty Input Rate is left out; a rate at a non-DLT maturity is kept (the calculation ignores it)
        self.assertEqual(inputs["SwapRatesInit"], {1: 0.02176, 2: 0.022621, 3: 0.0235, 5: 0.02569})

    def test_empty_dlt_and_weight_count_as_zero(self) -> None:
        self.write_inputs(curve="Maturity,DLT,LLFR Weight,Input Rate\n7,,,0.03\n")
        inputs = read_inputs(self.folder)
        self.assertEqual((inputs["dlt"], inputs["LLFRweightsIn"], inputs["SwapRatesInit"]), ({7: 0}, {7: 0.0}, {7: 0.03}))

    def test_excel_style_files(self) -> None:
        # semicolons, decimal commas and a byte order mark, as saved by Excel with a European locale
        self.write(CURVE_FILE, CURVE_TEXT.replace(",", ";").replace(".", ","), encoding="utf-8-sig")
        self.write(PARAMETERS_FILE, PARAMETERS_TEXT.replace(",", ";").replace(".", ","), encoding="utf-8-sig")
        excel_style = read_inputs(self.folder)
        self.write_inputs()
        self.assertEqual(excel_style, read_inputs(self.folder))

    def test_parameter_names_ignore_case_and_extra_rows(self) -> None:
        self.write_inputs(parameters="Parameter,Value\ncoupon frequency,2\nufr,0.035\nCONVERGENCE,0.4\n"
                                     "Cra,5\nmax maturity,100\nComment,ignored\n")
        inputs = read_inputs(self.folder)
        self.assertEqual((inputs["CouponFreq"], inputs["UFR"], inputs["alpha"], inputs["CRA"], inputs["MAX_MATURITY"]),
                         (2, 0.035, 0.4, 5.0, 100))

    def test_missing_curve_file(self) -> None:
        self.write(PARAMETERS_FILE, PARAMETERS_TEXT)
        with self.assertRaises(FileNotFoundError):
            read_inputs(self.folder)

    def test_missing_parameters_file(self) -> None:
        self.write(CURVE_FILE, CURVE_TEXT)
        with self.assertRaises(FileNotFoundError):
            read_inputs(self.folder)

    def test_missing_column(self) -> None:
        self.write_inputs(curve="Maturity,DLT,Input Rate\n1,1,0.02\n")
        with self.assertRaisesRegex(ValueError, "expected columns Maturity, DLT, LLFR Weight, Input Rate"):
            read_inputs(self.folder)

    def test_curve_without_rows(self) -> None:
        self.write_inputs(curve="Maturity,DLT,LLFR Weight,Input Rate\n")
        with self.assertRaisesRegex(ValueError, "expected columns"):
            read_inputs(self.folder)

    def test_maturity_outside_range(self) -> None:
        for maturity in ("0", "151"):
            self.write_inputs(curve=f"Maturity,DLT,LLFR Weight,Input Rate\n1,1,0,0.02\n{maturity},1,1,0.03\n")
            with self.assertRaisesRegex(ValueError, f"line 3: maturity {maturity} is outside 1..150"):
                read_inputs(self.folder)

    def test_max_maturity_limits_the_curve(self) -> None:
        # CURVE_TEXT has maturities 1..5; with Max Maturity 4 the row for 5 (line 6) is rejected
        self.write_inputs(parameters=PARAMETERS_TEXT.replace("Max Maturity,150", "Max Maturity,4"))
        with self.assertRaisesRegex(ValueError, "line 6: maturity 5 is outside 1..4"):
            read_inputs(self.folder)

    def test_max_maturity_must_be_a_whole_number(self) -> None:
        for value in ("150.5", "inf", "nan"):
            self.write_inputs(parameters=PARAMETERS_TEXT.replace("Max Maturity,150", f"Max Maturity,{value}"))
            with self.assertRaisesRegex(ValueError, "Max Maturity: must be a whole number of years"):
                read_inputs(self.folder)

    def test_max_maturity_must_be_at_least_one(self) -> None:
        # rejected as a parameter, before any curve row is compared against it
        for value in ("0", "-5"):
            self.write_inputs(parameters=PARAMETERS_TEXT.replace("Max Maturity,150", f"Max Maturity,{value}"))
            with self.assertRaisesRegex(ValueError, f"parameters.csv Max Maturity: must be at least 1 \\(is {value}\\)"):
                read_inputs(self.folder)

    def test_dlt_not_zero_or_one(self) -> None:
        self.write_inputs(curve="Maturity,DLT,LLFR Weight,Input Rate\n1,2,0,0.02\n")
        with self.assertRaisesRegex(ValueError, "line 2: DLT must be 0 or 1"):
            read_inputs(self.folder)

    def test_text_in_a_number_column(self) -> None:
        self.write_inputs(curve="Maturity,DLT,LLFR Weight,Input Rate\n1,1,0,2%\n")
        with self.assertRaisesRegex(ValueError, "swap_curve.csv line 2: '2%' is not a number"):
            read_inputs(self.folder)

    def test_missing_and_empty_parameters_are_listed(self) -> None:
        self.write_inputs(parameters="Parameter,Value\nCoupon Frequency,1\nUFR,\n")
        with self.assertRaisesRegex(ValueError, "missing parameter\\(s\\) UFR, Convergence, CRA, Max Maturity"):
            read_inputs(self.folder)

    def test_bad_parameter_value(self) -> None:
        self.write_inputs(parameters=PARAMETERS_TEXT.replace("0.033", "3.3%"))
        with self.assertRaisesRegex(ValueError, "parameters.csv UFR: '3.3%' is not a number"):
            read_inputs(self.folder)

    def test_decimal_comma_in_a_comma_separated_file(self) -> None:
        self.write_inputs(parameters=PARAMETERS_TEXT.replace("0.033", "0,033"))
        with self.assertRaisesRegex(ValueError, "parameters.csv line 3: more values than column names"):
            read_inputs(self.folder)


class TestCsvValue(unittest.TestCase):

    def test_nan_becomes_empty(self) -> None:
        self.assertEqual(csv_value(math.nan), "")

    def test_other_values_are_unchanged(self) -> None:
        self.assertEqual([csv_value(x) for x in (0.1, 0, 7, "SWP", "")], [0.1, 0, 7, "SWP", ""])


class TestSaveTable(TempFolderTestCase):

    def test_creates_the_folder_and_returns_the_path(self) -> None:
        target: Path = self.folder / "Output"
        path: Path = save_table(target, "t.csv", ["A"], [[1]])
        self.assertEqual(path, target / "t.csv")
        self.assertTrue(path.exists())

    def test_content_and_nan(self) -> None:
        path = save_table(self.folder, "t.csv", ["Name", "Value"], [["a", 1], ["b", math.nan]])
        self.assertEqual(path.read_text(encoding="utf-8").splitlines(), ["Name,Value", "a,1", "b,"])

    def test_floats_keep_full_precision(self) -> None:
        value: float = 0.1 + 0.2                       # 0.30000000000000004
        path = save_table(self.folder, "t.csv", ["Value"], [[value]])
        self.assertEqual(float(read_csv(path)[0]["Value"]), value)

    def test_overwrites_an_existing_file(self) -> None:
        save_table(self.folder, "t.csv", ["A"], [[1], [2]])
        path = save_table(self.folder, "t.csv", ["B"], [[3]])
        self.assertEqual(path.read_text(encoding="utf-8").splitlines(), ["B", "3"])

    def test_written_tables_can_be_read_back(self) -> None:
        path = save_table(self.folder, "t.csv", ["A", "B"], [["x", 0.25], ["y", math.nan]])
        self.assertEqual(read_csv(path), [{"A": "x", "B": "0.25"}, {"A": "y", "B": ""}])


class TestSaveResults(TempFolderTestCase):

    def setUp(self) -> None:
        super().setUp()
        self.inputs: SwapInputs = example_inputs()
        self.res: SheetResults = calculate_sheet(**self.inputs)
        self.paths: list[Path] = save_results(self.inputs, self.res, self.folder)

    def test_writes_both_files(self) -> None:
        self.assertEqual(self.paths, [self.folder / CURVES_OUTPUT_FILE, self.folder / PARAMETERS_OUTPUT_FILE])

    def test_curves_file_has_one_row_per_maturity(self) -> None:
        rows = read_csv(self.paths[0])
        self.assertEqual(list(rows[0]), ["Maturity", "DLT", "LLFR Weight", "Input Rate",
                                         "Bootstrapped Zero Rate CC", "Basic RFR"])
        self.assertEqual([int(row["Maturity"]) for row in rows], list(range(1, self.inputs["MAX_MATURITY"] + 1)))

    def test_curves_file_values(self) -> None:
        rows = {int(row["Maturity"]): row for row in read_csv(self.paths[0])}
        for t in range(1, self.inputs["MAX_MATURITY"] + 1):
            self.assertEqual(float(rows[t]["Basic RFR"]), self.res["BASIC_RFR"][t])
        self.assertEqual(float(rows[20]["Bootstrapped Zero Rate CC"]), self.res["BOOTSTRAPPED_CURVE"][20])
        self.assertEqual((rows[20]["DLT"], rows[20]["LLFR Weight"], rows[20]["Input Rate"]), ("1", "0.33", "0.03233"))
        # no input rate at 14 and no bootstrapped rate (#N/A) beyond the last DLT maturity: empty cells
        self.assertEqual((rows[14]["DLT"], rows[14]["Input Rate"]), ("0", ""))
        self.assertEqual(rows[51]["Bootstrapped Zero Rate CC"], "")

    def test_parameters_file(self) -> None:
        values = {row["Parameter"]: row["Value"] for row in read_csv(self.paths[1])}
        self.assertEqual(values["Instrument"], "SWP")
        self.assertEqual(values["FSP"], "20")
        self.assertEqual(float(values["LLFR (CC)"]), self.res["LLFR"])
        self.assertEqual((values["Coupon Frequency"], values["UFR"], values["Convergence"], values["CRA"],
                          values["Max Maturity"]), ("1", "0.033", "0.11", "10", "150"))

    def test_rows_follow_max_maturity(self) -> None:
        inputs: SwapInputs = example_inputs()
        inputs["MAX_MATURITY"] = 60
        paths = save_results(inputs, calculate_sheet(**inputs), self.folder)
        self.assertEqual([int(row["Maturity"]) for row in read_csv(paths[0])], list(range(1, 61)))


if __name__ == "__main__":
    unittest.main()
