"""EIOPA basic risk-free interest rate term structure from swaps or government bonds: main script.

Usage (from any folder):
    python Code/main.py

Reads the CSV files in "Input", validates them, calculates the curves and saves them in
"Output". Both folders sit in the project folder, next to the Code folder. Excel is not used.

Modules:
    data_io.py       reading Input, writing Output
    validation.py    validation of the inputs
    calculation.py   the calculation (port of the VBA module mExtrapolation) and shared types
    main.py          this script: brings them together
    tests/           unit tests (python -m unittest discover -s Code, or python -m pytest Code)

Inputs, in the folder "Input":
    curve.csv         columns Maturity, DLT, LLFR Weight, Input Rate (sheet columns G:J).
                      Input Rate is the par swap rate (swaps) or the annually compounded zero
                      rate (government bonds), before the CRA.
                      One row per maturity 1..Max Maturity; missing maturities count as DLT 0,
                      weight 0 and no rate. Input Rate may be empty; only rates with DLT = 1
                      are used.
    parameters.csv    columns Parameter, Value with the rows
                      Instrument (SWP for swaps or GVT for government bonds; sheet cell H11),
                      Coupon Frequency (swaps only), UFR, Convergence, CRA (sheet cells H12,
                      H14:H16; CRA in basis points, UFR and Convergence as decimals) and
                      Max Maturity (longest maturity of the curves in years; 150 on the sheet).
Both files may use ',' or ';' as separator and '.' or ',' as decimal mark.

Results, in the folder "Output" (created if missing, files overwritten on each run), named after
the instrument: swaps_*.csv for swaps, government_bonds_*.csv for government bonds:
    *_curves.csv                maturity 1..Max Maturity, DLT, LLFR weight, input rate,
                                bootstrapped zero rate CC, basic RFR
    *_parameters.csv            parameters used, FSP and LLFR
Rates are decimals with full precision; nan (#N/A) is written as an empty cell.
"""

import math
import sys
from pathlib import Path

from calculation import INSTRUMENTS, SheetInputs, SheetResults, calculate_sheet
from data_io import INPUT_DIR, OUTPUT_DIR, read_inputs, save_results
from validation import validate_inputs


def print_summary(res: SheetResults) -> None:
    """Print the FSP, the LLFR and both curves at a selection of maturities.

    The maturities shown are those of 1, 2, 5, 10, 15, 20, 21, 25, 30, 40, 50, 60, 80 and 100
    below the curves' longest maturity, followed by the longest maturity. Rates are shown as
    percentages; a missing bootstrapped rate as #N/A.

    Args:
        res: the results of calculation.calculate_sheet.

    Returns:
        None; the summary is printed to standard output.
    """
    print(f"FSP  : {res['FSP']} years")
    print(f"LLFR : {res['LLFR']:.8%} (continuously compounded)")
    print(f"{'Maturity':>8} {'Bootstrapped CC':>16} {'Basic RFR':>10}")
    MAX_MATURITY: int = max(res["BASIC_RFR"])
    for t in [t for t in (1, 2, 5, 10, 15, 20, 21, 25, 30, 40, 50, 60, 80, 100) if t < MAX_MATURITY] + [MAX_MATURITY]:
        boot: float = res["BOOTSTRAPPED_CURVE"][t]
        boot_text: str = "#N/A" if math.isnan(boot) else f"{boot:.5%}"
        print(f"{t:>8} {boot_text:>16} {res['BASIC_RFR'][t]:>10.3%}")


def run(input_dir: Path = INPUT_DIR, output_dir: Path = OUTPUT_DIR) -> SheetResults:
    """Read and validate the inputs, calculate the curves, print the instrument and a summary, and save the results.

    Nothing is printed or saved if the inputs cannot be read or are invalid.

    Args:
        input_dir: folder with curve.csv and parameters.csv; the project's Input folder by default.
        output_dir: folder for the results; the project's Output folder by default.

    Returns:
        The results of calculation.calculate_sheet.

    Raises:
        FileNotFoundError: if an input file is missing.
        ValueError: if the inputs cannot be read or are invalid; the message describes the problems.
    """
    inputs: SheetInputs = read_inputs(input_dir)
    validate_inputs(inputs)
    res: SheetResults = calculate_sheet(**inputs)
    print(f"Instrument: {inputs['Instrument']} ({INSTRUMENTS[inputs['Instrument']]})")
    print_summary(res)
    print()
    for saved in save_results(inputs, res, output_dir):
        print(f"Saved {saved}")
    return res


if __name__ == "__main__":
    try:
        run()
    except (FileNotFoundError, ValueError) as error:
        sys.exit(f"Input error: {error}")
