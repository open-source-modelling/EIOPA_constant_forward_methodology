"""Reading the inputs from the folder Input and writing the results to the folder Output.

Both folders sit in the project folder, one level above this Code folder, so the scripts
work from any working directory.
"""

import csv
import math
from pathlib import Path

from calculation import Curve, Flags, SheetResults, SwapInputs

PROJECT_DIR: Path = Path(__file__).resolve().parent.parent
INPUT_DIR: Path = PROJECT_DIR / "Input"
OUTPUT_DIR: Path = PROJECT_DIR / "Output"

CURVE_FILE: str = "swap_curve.csv"
PARAMETERS_FILE: str = "parameters.csv"
CURVES_OUTPUT_FILE: str = "swaps_curves.csv"
PARAMETERS_OUTPUT_FILE: str = "swaps_parameters.csv"

CURVE_COLUMNS: list[str] = ["Maturity", "DLT", "LLFR Weight", "Input Rate"]
PARAMETER_NAMES: list[str] = ["Coupon Frequency", "UFR", "Convergence", "CRA", "Max Maturity"]


# --------------------------------------------------------------------------
# Reading (folder Input)
# --------------------------------------------------------------------------

def read_csv(path: Path) -> list[dict[str, str]]:
    """Read a CSV file into a list of rows.

    The separator is ';' if the header line contains ';' and no ',', otherwise ','. A byte order
    mark (Excel "CSV UTF-8") is removed, and spaces around column names and values are stripped.
    Missing cells at the end of a row become empty strings.

    Args:
        path: the CSV file.

    Returns:
        One {column name: cell text} per data row; an empty list for an empty file.

    Raises:
        FileNotFoundError: if the file does not exist.
        ValueError: if a row has more cells than there are column names (typically a decimal
            comma in a file separated by commas); the message names the file and line.
    """
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    with path.open(newline="", encoding="utf-8-sig") as f:
        text: str = f.read()
    first_line: str = text.splitlines()[0] if text else ""
    delimiter: str = ";" if ";" in first_line and "," not in first_line else ","
    rows: list[dict[str, str]] = []
    for line, row in enumerate(csv.DictReader(text.splitlines(), delimiter=delimiter), start=2):
        if None in row:
            raise ValueError(f"{path.name} line {line}: more values than column names "
                             f"(a decimal comma in a file separated by commas?)")
        rows.append({(k or "").strip(): (v or "").strip() for k, v in row.items()})
    return rows


def parse_number(text: str, where: str) -> float:
    """Convert the text of a CSV cell to a number; a decimal comma is accepted.

    Args:
        text: the cell text, e.g. "0.033" or "0,033".
        where: location for the error message, e.g. "swap_curve.csv line 3".

    Returns:
        The number.

    Raises:
        ValueError: "<where>: '<text>' is not a number".
    """
    try:
        return float(text.replace(",", "."))
    except ValueError:
        raise ValueError(f"{where}: '{text}' is not a number") from None


def read_inputs(folder: Path = INPUT_DIR) -> SwapInputs:
    """Read the inputs from parameters.csv and swap_curve.csv.

    The parameters are read first, because Max Maturity limits the maturities in swap_curve.csv.
    This function checks the format of the files; validation.validate_inputs checks whether the
    values make sense.

    Args:
        folder: folder with the two files; the project's Input folder by default.

    Returns:
        SwapInputs with the parameters (CouponFreq, CRA, UFR, alpha, MAX_MATURITY) and the curve
        columns (dlt, LLFRweightsIn, SwapRatesInit). Maturities missing from swap_curve.csv are
        missing from the curve columns too (they count as DLT 0, weight 0 and no rate); an empty
        Input Rate is missing from SwapRatesInit.

    Raises:
        FileNotFoundError: if a file is missing.
        ValueError: if a parameter or column is missing, a value is not a number, Max Maturity is
            not a whole number of at least 1, a maturity lies outside 1..Max Maturity, or a DLT
            flag is not 0 or 1. The message names the file and, for the curve, the line.
    """
    parameters_path: Path = folder / PARAMETERS_FILE
    values: dict[str, str] = {row.get("Parameter", "").lower(): row.get("Value", "")
                              for row in read_csv(parameters_path)}
    missing: list[str] = [name for name in PARAMETER_NAMES if not values.get(name.lower())]
    if missing:
        raise ValueError(f"{parameters_path}: missing parameter(s) {', '.join(missing)}")

    def parameter(name: str) -> float:
        """Value of one parameter from parameters.csv.

        Args:
            name: the parameter name as in the file (case does not matter).

        Returns:
            The value as a number.

        Raises:
            ValueError: if the value is not a number.
        """
        return parse_number(values[name.lower()], f"{parameters_path.name} {name}")

    # checked here, before the curve rows are compared against it
    max_maturity: float = parameter("Max Maturity")
    if not max_maturity.is_integer():
        raise ValueError(f"{parameters_path.name} Max Maturity: must be a whole number of years (is {max_maturity})")
    if max_maturity < 1:
        raise ValueError(f"{parameters_path.name} Max Maturity: must be at least 1 (is {int(max_maturity)})")
    MAX_MATURITY: int = int(max_maturity)

    curve_path: Path = folder / CURVE_FILE
    rows: list[dict[str, str]] = read_csv(curve_path)
    missing = [c for c in CURVE_COLUMNS if rows and c not in rows[0]]
    if not rows or missing:
        raise ValueError(f"{curve_path}: expected columns {', '.join(CURVE_COLUMNS)}")

    dlt: Flags = {}
    LLFRweightsIn: Curve = {}
    SwapRatesInit: Curve = {}
    for line, row in enumerate(rows, start=2):
        where: str = f"{curve_path.name} line {line}"
        t: int = int(parse_number(row["Maturity"], where))
        if not 1 <= t <= MAX_MATURITY:
            raise ValueError(f"{where}: maturity {t} is outside 1..{MAX_MATURITY} (Max Maturity)")
        flag: int = int(parse_number(row["DLT"], where)) if row["DLT"] else 0
        if flag not in (0, 1):
            raise ValueError(f"{where}: DLT must be 0 or 1")
        dlt[t] = flag
        LLFRweightsIn[t] = parse_number(row["LLFR Weight"], where) if row["LLFR Weight"] else 0.0
        if row["Input Rate"]:
            SwapRatesInit[t] = parse_number(row["Input Rate"], where)

    return SwapInputs(
        CouponFreq=int(parameter("Coupon Frequency")),
        CRA=parameter("CRA"),
        UFR=parameter("UFR"),
        alpha=parameter("Convergence"),
        dlt=dlt,
        LLFRweightsIn=LLFRweightsIn,
        SwapRatesInit=SwapRatesInit,
        MAX_MATURITY=MAX_MATURITY,
    )


# --------------------------------------------------------------------------
# Writing (folder Output)
# --------------------------------------------------------------------------

def csv_value(x: str | int | float) -> str | int | float:
    """Prepare a value for a CSV cell.

    Args:
        x: the value to write.

    Returns:
        An empty string for nan (#N/A); any other value unchanged (the csv module writes floats
        with full precision).
    """
    return "" if isinstance(x, float) and math.isnan(x) else x


def save_table(folder: Path, file_name: str, header: list[str], rows: list[list[str | int | float]]) -> Path:
    """Write a table to a CSV file, overwriting an existing file.

    Args:
        folder: target folder; created if missing (its parent folder must exist).
        file_name: name of the CSV file.
        header: the column names.
        rows: one list of values per row; nan is written as an empty cell.

    Returns:
        The path of the written file.
    """
    folder.mkdir(exist_ok=True)
    path: Path = folder / file_name
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows([[csv_value(x) for x in row] for row in rows])
    return path


def save_results(inputs: SwapInputs, res: SheetResults, folder: Path = OUTPUT_DIR) -> list[Path]:
    """Write the results of a calculation to swaps_curves.csv and swaps_parameters.csv.

    swaps_curves.csv has one row per maturity 1..Max Maturity with the inputs (DLT, LLFR Weight,
    Input Rate) and the results (Bootstrapped Zero Rate CC, Basic RFR). swaps_parameters.csv lists
    the parameters, the FSP and the LLFR (CC).

    Args:
        inputs: the inputs of the calculation.
        res: the results of calculation.calculate_sheet for these inputs.
        folder: target folder; the project's Output folder by default.

    Returns:
        The paths of the written files: [swaps_curves.csv, swaps_parameters.csv].
    """
    curves_path: Path = save_table(
        folder, CURVES_OUTPUT_FILE,
        ["Maturity", "DLT", "LLFR Weight", "Input Rate", "Bootstrapped Zero Rate CC", "Basic RFR"],
        [[t, inputs["dlt"].get(t, 0), inputs["LLFRweightsIn"].get(t, 0.0), inputs["SwapRatesInit"].get(t, math.nan),
          res["BOOTSTRAPPED_CURVE"][t], res["BASIC_RFR"][t]] for t in range(1, inputs["MAX_MATURITY"] + 1)])
    parameters_path: Path = save_table(folder, PARAMETERS_OUTPUT_FILE, ["Parameter", "Value"], [
        ["Instrument", "SWP"],
        ["Coupon Frequency", inputs["CouponFreq"]],
        ["UFR", inputs["UFR"]],
        ["Convergence", inputs["alpha"]],
        ["CRA", inputs["CRA"]],
        ["Max Maturity", inputs["MAX_MATURITY"]],
        ["FSP", res["FSP"]],
        ["LLFR (CC)", res["LLFR"]],
    ])
    return [curves_path, parameters_path]
