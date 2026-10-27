# Generated from Code/main.py by Package/sync_code.py: edit that file and run the script again.
"""EIOPA risk-free interest rate term structures from swaps or government bonds, with VA and shocks: main script.

Usage (from any folder):
    python Code/main.py

Deletes the folder "Output", then runs three steps (function run):
    run_basic_rfr   reads curve.csv and parameters.csv from "Input", validates them, calculates
                    the basic risk-free rate and saves it in "Output";
    run_va          only if "Input" contains va_parameters.csv: reads and validates it, calculates
                    the CSSR curve and, if the values on it are given, the CSSR, the VA and the
                    curve with VA, and saves them in "Output";
    run_irr         only if "Input" contains irr_parameters.csv: reads and validates it, calculates
                    the curves after the upward and downward interest rate shocks and, if run_va
                    calculated a VA, the shocked curves with VA, and saves them in "Output".
Both folders sit in the project folder, next to the Code folder. Excel is not used.

Modules:
    data_io.py       reading Input, writing Output
    validation.py    validation of the inputs
    calculation.py   the calculation (port of the VBA module mExtrapolation) and shared types
    volatility_adjustment.py  the undertaking-specific VA: CSSR curve, CSSR and curve with VA
    interest_rate_risk.py     the curves after the interest rate shocks, without and with VA
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
    va_parameters.csv optional; columns Parameter, Value with the rows Currency RCS, and
                      optionally Country RCS, Debt to Total Assets, Use Macro VA (0 or 1), USA,
                      MV Fixed Income, MV Fixed Income CSSR Curve, BEL and BEL CSSR Curve.
                      Without it, no VA is calculated. Without the last four values, only the
                      CSSR curve is calculated, on which the assets and liabilities are valued.
    irr_parameters.csv optional; columns Maturity, smup, bmup, smdown, bmdown, rmfloor (sheet
                      Parameters G16:L166): the relative and absolute upward and downward shocks
                      and the floor of the downward shock, one row per maturity 1..Max Maturity
                      (rows beyond are not used). Without it, no shocks are calculated.
All files may use ',' or ';' as separator and '.' or ',' as decimal mark.

Results, in the folder "Output" (deleted and created again on each run), named after
the instrument: swaps_*.csv for swaps, government_bonds_*.csv for government bonds:
    *_curves.csv                maturity 1..Max Maturity, DLT, LLFR weight, input rate,
                                bootstrapped zero rate CC, basic RFR
    *_parameters.csv            parameters used, FSP and LLFR
    *_va_curves.csv             with va_parameters.csv: maturity 1..Max Maturity, basic RFR,
                                CSSR curve, VA curve
    *_va_parameters.csv         with va_parameters.csv: VA inputs, notional VA, CSSR, VA
    *_irr_curves.csv            with irr_parameters.csv: maturity 1..Max Maturity, basic RFR,
                                bootstrapped zero rate AC and, for each shock, the shocked rates
                                AC, the shocked curve and the shocked curve with VA
    *_irr_parameters.csv        with irr_parameters.csv: shocked UFRs, VA, shocked LLFRs
Rates are decimals with full precision; nan (#N/A) is written as an empty cell.
"""

import math
import sys
from pathlib import Path

from .calculation import INSTRUMENTS, SheetInputs, SheetResults, calculate_sheet
from .data_io import (INPUT_DIR, IRR_PARAMETERS_FILE, OUTPUT_DIR, VA_PARAMETERS_FILE, clear_output_folder, read_inputs,
                     read_irr_inputs, read_va_inputs, save_irr_results, save_results, save_va_results)
from .interest_rate_risk import IRRInputs, IRRResults, calculate_irr_down_sheet, calculate_irr_up_sheet
from .validation import validate_inputs, validate_irr_inputs, validate_va_inputs
from .volatility_adjustment import CSSRResults, VAInputs, VAResults, calculate_cssr_sheet, calculate_va_sheet

# maturities shown in the summaries (those below the curves' longest maturity), followed by the longest maturity
SUMMARY_MATURITIES: tuple[int, ...] = (1, 2, 5, 10, 15, 20, 21, 25, 30, 40, 50, 60, 80, 100)


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
    for t in [t for t in SUMMARY_MATURITIES if t < MAX_MATURITY] + [MAX_MATURITY]:
        boot: float = res["BOOTSTRAPPED_CURVE"][t]
        boot_text: str = "#N/A" if math.isnan(boot) else f"{boot:.5%}"
        print(f"{t:>8} {boot_text:>16} {res['BASIC_RFR'][t]:>10.3%}")


def print_va_summary(res: SheetResults, cssr: CSSRResults, va_res: VAResults | None) -> None:
    """Print the notional VA, the CSSR, the VA and the curves with VA at a selection of maturities.

    The maturities are those of print_summary. Without a CSSR (the values on the CSSR curve are not
    yet given), only the CSSR curve is shown, with a note on what to do next.

    Args:
        res: the basic RFR, from calculation.calculate_sheet.
        cssr: the results of volatility_adjustment.calculate_cssr_sheet.
        va_res: the results of volatility_adjustment.calculate_va_sheet; None without a CSSR.

    Returns:
        None; the summary is printed to standard output.
    """
    print(f"Notional VA : {cssr['NOTIONAL_VA'] * 10000:.0f} bp")
    if cssr["CSSR"] is None or va_res is None:
        print(f"CSSR        : not calculated; value the fixed income assets and the best estimate on the CSSR curve "
              f"and enter the four values in {VA_PARAMETERS_FILE}")
    else:
        print(f"CSSR        : {cssr['CSSR']:.6f}")
        print(f"Currency VA : {va_res['CURRENCY_VA'] * 10000:.2f} bp")
        if not math.isnan(va_res["MACRO_VA"]):
            print(f"Macro VA    : {va_res['MACRO_VA'] * 10000:.2f} bp")
        print(f"Total VA    : {va_res['TOTAL_VA'] * 10000:.0f} bp")
    print(f"{'Maturity':>8} {'Basic RFR':>10} {'CSSR Curve':>11} {'VA Curve':>9}")
    MAX_MATURITY: int = max(res["BASIC_RFR"])
    for t in [t for t in SUMMARY_MATURITIES if t < MAX_MATURITY] + [MAX_MATURITY]:
        va_text: str = f"{va_res['VA_CURVE'][t]:.3%}" if va_res else "#N/A"
        print(f"{t:>8} {res['BASIC_RFR'][t]:>10.3%} {cssr['CSSR_CURVE'][t]:>11.3%} {va_text:>9}")


def print_irr_summary(res: SheetResults, up: IRRResults, down: IRRResults) -> None:
    """Print the shocked UFRs and LLFRs and the shocked curves at a selection of maturities.

    The maturities are those of print_summary. Without a VA, the curves with VA are shown as #N/A,
    with a note why.

    Args:
        res: the basic RFR, from calculation.calculate_sheet.
        up: the results of interest_rate_risk.calculate_irr_up_sheet.
        down: the results of interest_rate_risk.calculate_irr_down_sheet.

    Returns:
        None; the summary is printed to standard output.
    """
    with_va: bool = not math.isnan(up["VA"])
    print(f"UFR up      : {up['UFR']:.2%}")
    print(f"UFR down    : {down['UFR']:.2%}")
    print(f"LLFR up     : {up['LLFR']:.8%} (continuously compounded)")
    print(f"LLFR down   : {down['LLFR']:.8%} (continuously compounded)")
    if with_va:
        print(f"VA          : {up['VA'] * 10000:.0f} bp")
        print(f"LLFR* up    : {up['LLFR_STAR']:.8%} (with VA, continuously compounded)")
        print(f"LLFR* down  : {down['LLFR_STAR']:.8%} (with VA, continuously compounded)")
    else:
        print("VA          : not available, so the curves with VA are not calculated")

    def rate(x: float, width: int) -> str:
        """A rate as a percentage, right-aligned; nan as #N/A."""
        return f"{'#N/A' if math.isnan(x) else f'{x:.3%}':>{width}}"

    print(f"{'Maturity':>8} {'Basic RFR':>10} {'Up':>8} {'Up + VA':>8} {'Down':>8} {'Down + VA':>9}")
    MAX_MATURITY: int = max(res["BASIC_RFR"])
    for t in [t for t in SUMMARY_MATURITIES if t < MAX_MATURITY] + [MAX_MATURITY]:
        print(f"{t:>8} {rate(res['BASIC_RFR'][t], 10)} {rate(up['IRR_RFR'][t], 8)} {rate(up['IRR_VA_CURVE'][t], 8)} "
              f"{rate(down['IRR_RFR'][t], 8)} {rate(down['IRR_VA_CURVE'][t], 9)}")


def print_saved(paths: list[Path]) -> None:
    """Print an empty line and the path of every saved file.

    Args:
        paths: the saved files.

    Returns:
        None; the paths are printed to standard output.
    """
    print()
    for path in paths:
        print(f"Saved {path}")


def run_basic_rfr(input_dir: Path = INPUT_DIR, output_dir: Path = OUTPUT_DIR) -> tuple[SheetInputs, SheetResults]:
    """The basic risk-free rate: read and validate curve.csv and parameters.csv, calculate, print and save.

    Nothing is printed or saved if the inputs cannot be read or are invalid.

    Args:
        input_dir: folder with curve.csv and parameters.csv; the project's Input folder by default.
        output_dir: folder for the results; the project's Output folder by default.

    Returns:
        The inputs and the results of calculation.calculate_sheet, which run_va builds on.

    Raises:
        FileNotFoundError: if an input file is missing.
        ValueError: if the inputs cannot be read or are invalid; the message describes the problems.
    """
    inputs: SheetInputs = read_inputs(input_dir)
    validate_inputs(inputs)
    res: SheetResults = calculate_sheet(**inputs)
    print(f"Instrument: {inputs['Instrument']} ({INSTRUMENTS[inputs['Instrument']]})")
    print_summary(res)
    print_saved(save_results(inputs, res, output_dir))
    return inputs, res


def run_va(inputs: SheetInputs, res: SheetResults, input_dir: Path = INPUT_DIR,
           output_dir: Path = OUTPUT_DIR) -> tuple[CSSRResults, VAResults | None]:
    """The volatility adjustment: read and validate va_parameters.csv, calculate, print and save.

    Always calculates the CSSR curve. When the four values on it are given, also the CSSR, the VA
    and the curve with VA; otherwise the summary says what to do next. Nothing is printed or saved
    if the VA inputs cannot be read or are invalid.

    Args:
        inputs: the inputs of the basic RFR, from run_basic_rfr.
        res: the basic RFR, from run_basic_rfr.
        input_dir: folder with va_parameters.csv; the project's Input folder by default.
        output_dir: folder for the results; the project's Output folder by default.

    Returns:
        The results of volatility_adjustment.calculate_cssr_sheet and calculate_va_sheet; the
        latter is None without a CSSR.

    Raises:
        FileNotFoundError: if va_parameters.csv is missing.
        ValueError: if the VA inputs cannot be read or are invalid; the message describes the problems.
    """
    va: VAInputs | None = read_va_inputs(input_dir)
    if va is None:
        raise FileNotFoundError(f"Input file not found: {input_dir / VA_PARAMETERS_FILE}")
    validate_va_inputs(va)
    cssr: CSSRResults = calculate_cssr_sheet(inputs, res, va)
    va_res: VAResults | None = calculate_va_sheet(inputs, res, va, cssr["CSSR"]) if cssr["CSSR"] is not None else None
    print_va_summary(res, cssr, va_res)
    print_saved(save_va_results(inputs, res, va, cssr, va_res, output_dir))
    return cssr, va_res


def run_irr(inputs: SheetInputs, res: SheetResults, TOTAL_VA: float | None = None, input_dir: Path = INPUT_DIR,
            output_dir: Path = OUTPUT_DIR) -> tuple[IRRResults, IRRResults]:
    """The interest rate shocks: read and validate irr_parameters.csv, calculate, print and save.

    Always calculates the shocked curves without VA; with a total VA also the shocked curves with
    VA. Nothing is printed or saved if the shock parameters cannot be read or are invalid.

    Args:
        inputs: the inputs of the basic RFR, from run_basic_rfr.
        res: the basic RFR, from run_basic_rfr.
        TOTAL_VA: the total VA, from run_va; None without a VA.
        input_dir: folder with irr_parameters.csv; the project's Input folder by default.
        output_dir: folder for the results; the project's Output folder by default.

    Returns:
        The results of interest_rate_risk.calculate_irr_up_sheet and calculate_irr_down_sheet.

    Raises:
        FileNotFoundError: if irr_parameters.csv is missing.
        ValueError: if the shock parameters cannot be read or are invalid; the message describes the problems.
    """
    irr: IRRInputs | None = read_irr_inputs(input_dir)
    if irr is None:
        raise FileNotFoundError(f"Input file not found: {input_dir / IRR_PARAMETERS_FILE}")
    validate_irr_inputs(irr, inputs["MAX_MATURITY"])
    up: IRRResults = calculate_irr_up_sheet(inputs, res, irr, TOTAL_VA)
    down: IRRResults = calculate_irr_down_sheet(inputs, res, irr, TOTAL_VA)
    print_irr_summary(res, up, down)
    print_saved(save_irr_results(inputs, res, up, down, output_dir))
    return up, down


def run(input_dir: Path = INPUT_DIR, output_dir: Path = OUTPUT_DIR) -> SheetResults:
    """Run the whole calculation: delete the output folder, then run_basic_rfr, run_va and run_irr.

    run_va only runs with va_parameters.csv and run_irr only with irr_parameters.csv. run_irr gets
    the total VA of run_va, if there is one (va_parameters.csv with the four values for the CSSR);
    otherwise only the shocked curves without VA are calculated.

    The output folder is deleted first, so that it only ever holds the results of this run: if the
    run stops with an error, no results of an earlier run are left behind. Each part validates its
    own inputs; invalid VA inputs stop the run after the basic RFR has been saved, invalid shock
    parameters after the VA has been saved.

    Args:
        input_dir: folder with curve.csv, parameters.csv and optionally va_parameters.csv and
            irr_parameters.csv; the project's Input folder by default.
        output_dir: folder for the results; the project's Output folder by default. It is deleted
            and created again.

    Returns:
        The results of calculation.calculate_sheet.

    Raises:
        FileNotFoundError: if an input file is missing.
        ValueError: if the inputs cannot be read or are invalid, or the output folder cannot be
            deleted safely; the message describes the problems.
    """
    clear_output_folder(output_dir, input_dir)
    inputs, res = run_basic_rfr(input_dir, output_dir)
    va_res: VAResults | None = None
    if (input_dir / VA_PARAMETERS_FILE).exists():
        print()
        _, va_res = run_va(inputs, res, input_dir, output_dir)
    if (input_dir / IRR_PARAMETERS_FILE).exists():
        print()
        run_irr(inputs, res, va_res["TOTAL_VA"] if va_res else None, input_dir, output_dir)
    return res


if __name__ == "__main__":
    try:
        run()
    except (FileNotFoundError, ValueError) as error:
        sys.exit(f"Input error: {error}")
