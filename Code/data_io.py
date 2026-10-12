"""Reading the inputs from the folder Input and writing the results to the folder Output.

Both folders sit in the project folder, one level above this Code folder, so the scripts
work from any working directory.
"""

import csv
import math
import shutil
from pathlib import Path

from calculation import INSTRUMENTS, Curve, Flags, SheetInputs, SheetResults
from volatility_adjustment import CSSRResults, VAInputs, VAResults

PROJECT_DIR: Path = Path(__file__).resolve().parent.parent
INPUT_DIR: Path = PROJECT_DIR / "Input"
OUTPUT_DIR: Path = PROJECT_DIR / "Output"

CURVE_FILE: str = "curve.csv"
PARAMETERS_FILE: str = "parameters.csv"
VA_PARAMETERS_FILE: str = "va_parameters.csv"     # optional: without it, no VA is calculated
# result files per instrument type: (curves, parameters)
OUTPUT_FILES: dict[str, tuple[str, str]] = {
    "SWP": ("swaps_curves.csv", "swaps_parameters.csv"),
    "GVT": ("government_bonds_curves.csv", "government_bonds_parameters.csv"),
}
# VA result files per instrument type: (curves, parameters)
VA_OUTPUT_FILES: dict[str, tuple[str, str]] = {
    "SWP": ("swaps_va_curves.csv", "swaps_va_parameters.csv"),
    "GVT": ("government_bonds_va_curves.csv", "government_bonds_va_parameters.csv"),
}

CURVE_COLUMNS: list[str] = ["Maturity", "DLT", "LLFR Weight", "Input Rate"]
# Coupon Frequency is required for swaps only
PARAMETER_NAMES: list[str] = ["Instrument", "Coupon Frequency", "UFR", "Convergence", "CRA", "Max Maturity"]


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
        where: location for the error message, e.g. "curve.csv line 3".

    Returns:
        The number.

    Raises:
        ValueError: "<where>: '<text>' is not a number".
    """
    try:
        return float(text.replace(",", "."))
    except ValueError:
        raise ValueError(f"{where}: '{text}' is not a number") from None


def read_parameters(path: Path) -> dict[str, str]:
    """Read a parameters file with the columns Parameter and Value.

    Args:
        path: the CSV file.

    Returns:
        {parameter name in lower case: value text}.

    Raises:
        FileNotFoundError: if the file does not exist.
        ValueError: if a row cannot be read, see read_csv.
    """
    return {row.get("Parameter", "").lower(): row.get("Value", "") for row in read_csv(path)}


def read_inputs(folder: Path = INPUT_DIR) -> SheetInputs:
    """Read the inputs from parameters.csv and curve.csv.

    The parameters are read first, because Max Maturity limits the maturities in curve.csv.
    This function checks the format of the files; validation.validate_inputs checks whether the
    values make sense.

    Args:
        folder: folder with the two files; the project's Input folder by default.

    Returns:
        SheetInputs with the parameters (Instrument, CouponFreq, CRA, UFR, alpha, MAX_MATURITY) and
        the curve columns (dlt, LLFRweightsIn, Rate). Instrument is upper case; CouponFreq is None
        for government bonds, which do not use it. Maturities missing from curve.csv are missing
        from the curve columns too (they count as DLT 0, weight 0 and no rate); an empty Input Rate
        is missing from Rate.

    Raises:
        FileNotFoundError: if a file is missing.
        ValueError: if a parameter or column is missing, Instrument is not SWP or GVT, a value is
            not a number, Max Maturity is not a whole number of at least 1, a maturity lies outside
            1..Max Maturity, or a DLT flag is not 0 or 1. The message names the file and, for the
            curve, the line.
    """
    parameters_path: Path = folder / PARAMETERS_FILE
    values: dict[str, str] = read_parameters(parameters_path)
    Instrument: str = values.get("instrument", "").upper()
    missing: list[str] = [name for name in PARAMETER_NAMES if not values.get(name.lower())
                          and not (name == "Coupon Frequency" and Instrument == "GVT")]
    if missing:
        raise ValueError(f"{parameters_path}: missing parameter(s) {', '.join(missing)}")
    if Instrument not in INSTRUMENTS:
        raise ValueError(f"{parameters_path.name} Instrument: must be SWP (swaps) or GVT (government bonds) "
                         f"(is {values['instrument']})")

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
    Rate: Curve = {}
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
            Rate[t] = parse_number(row["Input Rate"], where)

    return SheetInputs(
        Instrument=Instrument,
        CouponFreq=int(parameter("Coupon Frequency")) if Instrument == "SWP" else None,
        CRA=parameter("CRA"),
        UFR=parameter("UFR"),
        alpha=parameter("Convergence"),
        dlt=dlt,
        LLFRweightsIn=LLFRweightsIn,
        Rate=Rate,
        MAX_MATURITY=MAX_MATURITY,
    )


def read_va_inputs(folder: Path = INPUT_DIR) -> VAInputs | None:
    """Read the VA inputs from va_parameters.csv, if the file exists.

    Only Currency RCS is required. Country RCS and Debt to Total Assets are needed for the macro
    VA; Use Macro VA is 1 to add it and 0 (or empty) otherwise; USA is the undertaking-specific
    adjustment, empty if not used. The four values for the CSSR (MV Fixed Income, MV Fixed Income
    CSSR Curve, BEL, BEL CSSR Curve) stay empty until the undertaking has valued its assets and
    liabilities on the CSSR curve. validation.validate_va_inputs checks whether the values make sense.

    Args:
        folder: folder with va_parameters.csv; the project's Input folder by default.

    Returns:
        VAInputs, with None for every empty value; None if the file does not exist.

    Raises:
        ValueError: if Currency RCS is missing, a value is not a number, or Use Macro VA is not 0 or 1.
    """
    path: Path = folder / VA_PARAMETERS_FILE
    if not path.exists():
        return None
    values: dict[str, str] = read_parameters(path)
    if not values.get("currency rcs"):
        raise ValueError(f"{path}: missing parameter Currency RCS")

    def optional(name: str) -> float | None:
        """Value of one parameter from va_parameters.csv.

        Args:
            name: the parameter name as in the file (case does not matter).

        Returns:
            The value as a number; None if the parameter is missing or empty.

        Raises:
            ValueError: if the value is not a number.
        """
        text: str = values.get(name.lower(), "")
        return parse_number(text, f"{path.name} {name}") if text else None

    use_macro_va: float | None = optional("Use Macro VA")
    if use_macro_va not in (None, 0, 1):
        raise ValueError(f"{path.name} Use Macro VA: must be 0 or 1 (is {use_macro_va})")
    return VAInputs(
        RCS_CU=parse_number(values["currency rcs"], f"{path.name} Currency RCS"),
        RCS_CO=optional("Country RCS"),
        DTTA=optional("Debt to Total Assets"),
        MVFI=optional("MV Fixed Income"),
        MVFI_NOTIONAL=optional("MV Fixed Income CSSR Curve"),
        BEL=optional("BEL"),
        BEL_NOTIONAL=optional("BEL CSSR Curve"),
        MACROVA_SELECTED=use_macro_va == 1,
        USA=optional("USA"),
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


def clear_output_folder(folder: Path = OUTPUT_DIR, input_dir: Path = INPUT_DIR) -> None:
    """Delete the output folder and everything in it, so that no results of an earlier run are left behind.

    As a safeguard, a folder that contains the input folder or this Code folder is not deleted.

    Args:
        folder: the output folder; the project's Output folder by default. Nothing happens if it
            does not exist.
        input_dir: the input folder of the run; the project's Input folder by default.

    Returns:
        None.

    Raises:
        ValueError: if ``folder`` is a file, or contains the input folder or the Code folder.
    """
    if not folder.exists():
        return
    if not folder.is_dir():
        raise ValueError(f"the output folder {folder} is a file")
    target: Path = folder.resolve()
    for kept in (input_dir.resolve(), Path(__file__).resolve().parent):
        if kept == target or target in kept.parents:
            raise ValueError(f"the output folder {folder} contains {kept}, so it is not deleted")
    shutil.rmtree(target)


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


def save_results(inputs: SheetInputs, res: SheetResults, folder: Path = OUTPUT_DIR) -> list[Path]:
    """Write the results of a calculation to the result files of its instrument type.

    The files are swaps_curves.csv and swaps_parameters.csv for swaps, and
    government_bonds_curves.csv and government_bonds_parameters.csv for government bonds
    (OUTPUT_FILES), so the results of both can sit side by side. The curves file has one row per
    maturity 1..Max Maturity with the inputs (DLT, LLFR Weight, Input Rate) and the results
    (Bootstrapped Zero Rate CC, Basic RFR). The parameters file lists the parameters (Coupon
    Frequency for swaps only), the FSP and the LLFR (CC) in the order of the sheet.

    Args:
        inputs: the inputs of the calculation.
        res: the results of calculation.calculate_sheet for these inputs.
        folder: target folder; the project's Output folder by default.

    Returns:
        The paths of the written files: [curves file, parameters file].
    """
    curves_file, parameters_file = OUTPUT_FILES[inputs["Instrument"]]
    curves_path: Path = save_table(
        folder, curves_file,
        ["Maturity", "DLT", "LLFR Weight", "Input Rate", "Bootstrapped Zero Rate CC", "Basic RFR"],
        [[t, inputs["dlt"].get(t, 0), inputs["LLFRweightsIn"].get(t, 0.0), inputs["Rate"].get(t, math.nan),
          res["BOOTSTRAPPED_CURVE"][t], res["BASIC_RFR"][t]] for t in range(1, inputs["MAX_MATURITY"] + 1)])
    coupon_frequency: list[list[str | int | float]] = (
        [["Coupon Frequency", inputs["CouponFreq"]]] if inputs["CouponFreq"] is not None else [])
    # in the order of the sheet: H11:H16, Max Maturity (no cell), L22
    parameters_path: Path = save_table(folder, parameters_file, ["Parameter", "Value"], [
        ["Instrument", inputs["Instrument"]],
        *coupon_frequency,
        ["FSP", res["FSP"]],
        ["UFR", inputs["UFR"]],
        ["Convergence", inputs["alpha"]],
        ["CRA", inputs["CRA"]],
        ["Max Maturity", inputs["MAX_MATURITY"]],
        ["LLFR (CC)", res["LLFR"]],
    ])
    return [curves_path, parameters_path]


def save_va_results(inputs: SheetInputs, res: SheetResults, va: VAInputs, cssr: CSSRResults,
                    va_res: VAResults | None, folder: Path = OUTPUT_DIR) -> list[Path]:
    """Write the results of the VA calculation to the VA result files of the instrument type.

    The files are swaps_va_curves.csv and swaps_va_parameters.csv for swaps, and
    government_bonds_va_curves.csv and government_bonds_va_parameters.csv for government bonds
    (VA_OUTPUT_FILES). The curves file has one row per maturity 1..Max Maturity with the Basic RFR,
    the CSSR Curve and the VA Curve. The parameters file lists the VA inputs and the calculated
    values in the order of the workbook: the VA parameters of sheet 'Input Data & Extrapolation',
    then sheet 'CSSR Calculation', then sheet 'VA Term Structures'. Values that are not available,
    such as the CSSR before the valuation on the CSSR curve and the macro VA when it is not
    selected, are written as empty cells.

    Args:
        inputs: the inputs of the basic RFR.
        res: the basic RFR, from calculation.calculate_sheet.
        va: the VA inputs.
        cssr: the results of volatility_adjustment.calculate_cssr_sheet.
        va_res: the results of volatility_adjustment.calculate_va_sheet; None without a CSSR.
        folder: target folder; the project's Output folder by default.

    Returns:
        The paths of the written files: [curves file, parameters file].
    """
    def value(x: float | None) -> float:
        """A value for the parameters file: None becomes nan, which is written as an empty cell."""
        return math.nan if x is None else x

    curves_file, parameters_file = VA_OUTPUT_FILES[inputs["Instrument"]]
    curves_path: Path = save_table(
        folder, curves_file, ["Maturity", "Basic RFR", "CSSR Curve", "VA Curve"],
        [[t, res["BASIC_RFR"][t], cssr["CSSR_CURVE"][t], va_res["VA_CURVE"][t] if va_res else math.nan]
         for t in range(1, inputs["MAX_MATURITY"] + 1)])
    # in the order of the workbook: sheet by sheet, each block top to bottom
    parameters_path: Path = save_table(folder, parameters_file, ["Parameter", "Value"], [
        # Input Data & Extrapolation, VA Parameters L10:L12
        ["Currency RCS", va["RCS_CU"]],
        ["Country RCS", value(va["RCS_CO"])],
        ["Debt to Total Assets", value(va["DTTA"])],
        # CSSR Calculation: H10, L10:L16, K20
        ["Notional VA", cssr["NOTIONAL_VA"]],
        ["MV Fixed Income", value(va["MVFI"])],
        ["MV Fixed Income CSSR Curve", value(va["MVFI_NOTIONAL"])],
        ["PVBP MV Fixed Income", value(cssr["PVBP_MVFI"])],
        ["BEL", value(va["BEL"])],
        ["BEL CSSR Curve", value(va["BEL_NOTIONAL"])],
        ["PVBP BEL", value(cssr["PVBP_BEL"])],
        ["CSSR", value(cssr["CSSR"])],
        ["Notional VA LLFR (CC)", cssr["NOTIONAL_VA_LLFR"]],
        # VA Term Structures: H10:H12 (USA? and USA in one row), L10:L13, K20
        ["Use Macro VA", int(va["MACROVA_SELECTED"])],
        ["USA", value(va["USA"])],
        ["Currency VA", va_res["CURRENCY_VA"] if va_res else math.nan],
        ["Country Adjustment Factor", va_res["W_CO"] if va_res else math.nan],
        ["Macro VA", va_res["MACRO_VA"] if va_res else math.nan],
        ["Total VA", va_res["TOTAL_VA"] if va_res else math.nan],
        ["Total VA LLFR (CC)", va_res["TOTAL_VA_LLFR"] if va_res else math.nan],
    ])
    return [curves_path, parameters_path]
