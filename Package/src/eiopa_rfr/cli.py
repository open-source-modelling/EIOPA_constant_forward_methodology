"""The command eiopa-rfr: the whole calculation of Code/main.py on any input and output folder.

Usage:
    eiopa-rfr [--input FOLDER] [--output FOLDER]
    python -m eiopa_rfr [--input FOLDER] [--output FOLDER]

The folders default to Input and Output in the working folder. The input folder holds curve.csv,
parameters.csv and optionally va_parameters.csv and irr_parameters.csv, as in the repository's
Input folder; the results are written as in its Output folder.

Unlike main.run, which deletes the repository's whole Output folder, the command deletes only the
result files an earlier run may have left in the output folder (swaps_*.csv and
government_bonds_*.csv with the names of main.run's results), so that it can be pointed at a folder
that holds other files too. It does this before reading the inputs, so that a run that stops with
an error leaves no results of an earlier run behind.
"""

import argparse
import sys
from pathlib import Path

from . import __version__
from .data_io import IRR_OUTPUT_FILES, IRR_PARAMETERS_FILE, OUTPUT_FILES, VA_OUTPUT_FILES, VA_PARAMETERS_FILE
from .main import run_basic_rfr, run_irr, run_va
from .volatility_adjustment import VAResults


def remove_previous_results(output_dir: Path) -> None:
    """Delete the result files of an earlier run from the output folder, and nothing else.

    Args:
        output_dir: the output folder. Nothing happens if it does not exist.

    Returns:
        None.

    Raises:
        ValueError: if ``output_dir`` is a file.
    """
    if not output_dir.exists():
        return
    if not output_dir.is_dir():
        raise ValueError(f"the output folder {output_dir} is a file")
    for files in (OUTPUT_FILES, VA_OUTPUT_FILES, IRR_OUTPUT_FILES):
        for names in files.values():
            for name in names:
                (output_dir / name).unlink(missing_ok=True)


def run(input_dir: Path, output_dir: Path) -> None:
    """Delete the previous results, then run_basic_rfr, run_va and run_irr as main.run does.

    run_va only runs with va_parameters.csv and run_irr only with irr_parameters.csv; run_irr gets
    the total VA of run_va, if there is one. The output folder is created if missing, with its
    parent folders.

    Args:
        input_dir: folder with curve.csv, parameters.csv and optionally va_parameters.csv and
            irr_parameters.csv.
        output_dir: folder for the results.

    Returns:
        None; the summaries are printed and the results saved.

    Raises:
        FileNotFoundError: if an input file is missing.
        ValueError: if the inputs cannot be read or are invalid, or the output folder is a file;
            the message describes the problems.
    """
    remove_previous_results(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    inputs, res = run_basic_rfr(input_dir, output_dir)
    va_res: VAResults | None = None
    if (input_dir / VA_PARAMETERS_FILE).exists():
        print()
        _, va_res = run_va(inputs, res, input_dir, output_dir)
    if (input_dir / IRR_PARAMETERS_FILE).exists():
        print()
        run_irr(inputs, res, va_res["TOTAL_VA"] if va_res else None, input_dir, output_dir)


def main(argv: list[str] | None = None) -> None:
    """Run the command; see the module docstring.

    Args:
        argv: the command line arguments; sys.argv[1:] by default.

    Returns:
        None.

    Raises:
        SystemExit: with the message "Input error: ..." if an input file is missing, the inputs are
            invalid or the output folder is a file; with code 0 after --help or --version.
    """
    parser = argparse.ArgumentParser(
        prog="eiopa-rfr",
        description="EIOPA risk-free interest rate term structure from swaps or government bonds, with the "
                    "volatility adjustment and the interest rate shocks.",
        epilog="The input folder holds curve.csv and parameters.csv, and optionally va_parameters.csv (VA) and "
               "irr_parameters.csv (shocks). See https://github.com/open-source-modelling/"
               "EIOPA_constant_forward_methodology#inputs-and-outputs for the columns.")
    parser.add_argument("-i", "--input", type=Path, default=Path("Input"), metavar="FOLDER",
                        help="folder with the input files (default: Input)")
    parser.add_argument("-o", "--output", type=Path, default=Path("Output"), metavar="FOLDER",
                        help="folder for the results; created if missing, and only earlier result files are "
                             "deleted from it (default: Output)")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)
    try:
        run(args.input, args.output)
    except (FileNotFoundError, ValueError) as error:
        sys.exit(f"Input error: {error}")
