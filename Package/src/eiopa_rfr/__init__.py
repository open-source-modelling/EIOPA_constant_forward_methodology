"""EIOPA's Solvency II risk-free interest rate term structures: basic RFR, VA and interest rate shocks.

The package form of the code in the folder Code of
https://github.com/open-source-modelling/EIOPA_constant_forward_methodology, where the method, the
inputs and outputs and the verification against EIOPA's workbook are documented.

Modules (generated from Code by Package/sync_code.py; identical apart from the imports):
    calculation             the basic risk-free rate (port of the VBA module mExtrapolation) and shared types
    volatility_adjustment   the undertaking-specific VA: CSSR curve, CSSR and curve with VA
    interest_rate_risk      the curves after the interest rate shocks, without and with VA
    validation              checks that the inputs are complete and plausible
    data_io                 reading the input files and writing the result files
    main                    run_basic_rfr, run_va and run_irr: read, validate, calculate, print, save
Modules of the package only:
    cli                     the command eiopa-rfr (also python -m eiopa_rfr)

The default folders of data_io and main (INPUT_DIR and OUTPUT_DIR) are those of the repository,
next to its Code folder; in an installed package they point inside the installation. Always pass
the folders to the functions that read or write files.

The names most often needed are available from the package itself, e.g.
    from eiopa_rfr import calculate_sheet, validate_inputs
"""

from .calculation import (DEFAULT_MAX_MATURITY, INSTRUMENTS, Curve, Flags, SheetInputs, SheetResults, bootstrap_curve,
                          bootstrap_swaps, bootstrap_zeros, calculate_sheet, excel_round, excel_round_value, extrapolation,
                          get_llfr, last_liquid_point)
from .data_io import (read_inputs, read_irr_inputs, read_va_inputs, save_irr_results, save_results,
                      save_va_results)
from .interest_rate_risk import (UFR_SHIFT, IRRInputs, IRRResults, apply_floor, bootstrapped_curve_ac,
                                 calculate_irr_down_sheet, calculate_irr_up_sheet, shock_down, shock_up,
                                 shocked_term_structures)
from .validation import validate_inputs, validate_irr_inputs, validate_va_inputs
from .volatility_adjustment import (VA_SHARE, CSSRResults, VAInputs, VAResults, calculate_cssr_sheet,
                                    calculate_va_sheet, country_adjustment_factor, credit_spread_sensitivity_ratio,
                                    notional_va, va_curve)

# the package version; pyproject.toml reads it from here
__version__ = "0.1.0"

__all__ = [
    # calculation
    "DEFAULT_MAX_MATURITY", "INSTRUMENTS", "Curve", "Flags", "SheetInputs", "SheetResults", "bootstrap_curve",
    "bootstrap_swaps", "bootstrap_zeros", "calculate_sheet", "excel_round", "excel_round_value", "extrapolation",
    "get_llfr", "last_liquid_point",
    # volatility_adjustment
    "VA_SHARE", "CSSRResults", "VAInputs", "VAResults", "calculate_cssr_sheet", "calculate_va_sheet",
    "country_adjustment_factor", "credit_spread_sensitivity_ratio", "notional_va", "va_curve",
    # interest_rate_risk
    "UFR_SHIFT", "IRRInputs", "IRRResults", "apply_floor", "bootstrapped_curve_ac", "calculate_irr_down_sheet",
    "calculate_irr_up_sheet", "shock_down", "shock_up", "shocked_term_structures",
    # validation
    "validate_inputs", "validate_irr_inputs", "validate_va_inputs",
    # data_io
    "read_inputs", "read_irr_inputs", "read_va_inputs", "save_irr_results", "save_results", "save_va_results",
    "__version__",
]
