"""Validation of the inputs before the calculation."""

from calculation import INSTRUMENTS, Curve, SheetInputs
from volatility_adjustment import VAInputs, notional_va


def validate_inputs(inputs: SheetInputs) -> None:
    """Check that the inputs can be calculated and look plausible.

    All rules are checked and every problem is reported, so that all of them can be fixed at once.

    Rules:
    - Instrument is SWP (swaps) or GVT (government bonds).
    - Max Maturity is at least 1, and no DLT flag, LLFR weight or Input Rate lies beyond it.
    - For swaps, Coupon Frequency is given and at least 1 (government bonds do not use it).
    - Convergence is positive.
    - UFR and every Input Rate are decimals (between -1 and 1), e.g. 0.033 and not 3.3.
    - At least one maturity has DLT = 1 and an Input Rate.
    - LLFR weights are not negative, at least one is positive, and they sum to 1.
    - Every positive LLFR weight sits on a maturity with DLT = 1 and an Input Rate
      (PDF 8.5.6: the LLFR uses forwards to DLT maturities only).
    - There is a DLT maturity with an Input Rate before the FSP, so that the first
      LLFR forward rate exists (PDF 8.5.6, VBA GetLLFR).

    Args:
        inputs: the inputs, as returned by data_io.read_inputs.

    Returns:
        None if the inputs are valid.

    Raises:
        ValueError: "invalid inputs:" followed by one line per problem found.
    """
    problems: list[str] = []
    if inputs["Instrument"] not in INSTRUMENTS:
        problems.append(f"Instrument must be SWP (swaps) or GVT (government bonds) (is {inputs['Instrument']})")
    MAX_MATURITY: int = inputs["MAX_MATURITY"]
    if MAX_MATURITY < 1:
        problems.append(f"Max Maturity must be at least 1 (is {MAX_MATURITY})")
    beyond: list[int] = sorted({t for t, flag in inputs["dlt"].items() if flag == 1 and t > MAX_MATURITY}
                               | {t for t, w in inputs["LLFRweightsIn"].items() if w != 0 and t > MAX_MATURITY}
                               | {t for t in inputs["Rate"] if t > MAX_MATURITY})
    if beyond:
        problems.append(f"maturities {beyond} lie beyond Max Maturity ({MAX_MATURITY})")
    if inputs["Instrument"] == "SWP" and (inputs["CouponFreq"] is None or inputs["CouponFreq"] < 1):
        problems.append(f"Coupon Frequency must be at least 1 for swaps (is {inputs['CouponFreq']})")
    if not inputs["alpha"] > 0:
        problems.append(f"Convergence must be positive (is {inputs['alpha']})")
    if not -1 < inputs["UFR"] < 1:
        problems.append(f"UFR must be a decimal, e.g. 0.033 for 3.3% (is {inputs['UFR']})")

    bad_rates: list[int] = [t for t, r in inputs["Rate"].items() if not -1 < r < 1]
    if bad_rates:
        problems.append(f"Input Rate must be a decimal, e.g. 0.02176 for 2.176%; check maturities {bad_rates}")

    liquid: list[int] = sorted(t for t, flag in inputs["dlt"].items() if flag == 1 and t in inputs["Rate"])
    if not liquid:
        problems.append("no maturity has DLT = 1 and an Input Rate")

    weight: Curve = inputs["LLFRweightsIn"]                                  # VBA GetLLFR: weight
    negative: list[int] = [t for t, w in weight.items() if w < 0]
    if negative:
        problems.append(f"LLFR weights must not be negative; check maturities {negative}")
    weighted: list[int] = sorted(t for t, w in weight.items() if w > 0)    # VBA GetLLFR: t(1), t(2), ...
    if not weighted:
        problems.append("no maturity has a positive LLFR weight, so there is no FSP")
    else:
        FSP: int = weighted[0]
        sumw: float = sum(weight[t] for t in weighted)
        if abs(sumw - 1) > 1e-9:
            problems.append(f"LLFR weights must sum to 1 (sum is {sumw})")
        not_liquid: list[int] = [t for t in weighted if t not in liquid]
        if not_liquid:
            problems.append(f"LLFR weights are only allowed on maturities with DLT = 1 and an Input Rate; "
                            f"check maturities {not_liquid}")
        LLPbeforeFSP: int = max((t for t in liquid if t < FSP), default=0)
        if LLPbeforeFSP == 0:
            problems.append(f"there must be a maturity with DLT = 1 and an Input Rate before the FSP ({FSP})")

    if problems:
        raise ValueError("invalid inputs:\n  - " + "\n  - ".join(problems))


def validate_va_inputs(va: VAInputs) -> None:
    """Check that the VA inputs can be calculated and look plausible.

    All rules are checked and every problem is reported, so that all of them can be fixed at once.

    Rules:
    - Currency RCS, and Country RCS if given, are decimals (between -1 and 1), e.g. 0.01 and not 100.
    - The four values for the CSSR are either all given or all missing (before the valuation on the
      CSSR curve has been done).
    - If they are given, the notional VA, 85% of the currency RCS rounded to whole basis points,
      is not 0; otherwise the change in value per unit of spread cannot be calculated.
    - The macro VA needs the Country RCS and a Debt to Total Assets ratio between 0 and 1.
    - The undertaking-specific adjustment (USA) is at most 105% (PDF 9.3.5).
    - The USA and the macro VA are not used together (PDF 9.3.4). Difference from the workbook,
      which adds the macro VA to a currency VA with USA if both are switched on.

    Args:
        va: the VA inputs, as returned by data_io.read_va_inputs.

    Returns:
        None if the inputs are valid.

    Raises:
        ValueError: "invalid VA inputs:" followed by one line per problem found.
    """
    problems: list[str] = []
    if not -1 < va["RCS_CU"] < 1:
        problems.append(f"Currency RCS must be a decimal, e.g. 0.01 for 100 basis points (is {va['RCS_CU']})")
    if va["RCS_CO"] is not None and not -1 < va["RCS_CO"] < 1:
        problems.append(f"Country RCS must be a decimal, e.g. 0.014 for 140 basis points (is {va['RCS_CO']})")

    values: dict[str, float | None] = {"MV Fixed Income": va["MVFI"], "MV Fixed Income CSSR Curve": va["MVFI_NOTIONAL"],
                                       "BEL": va["BEL"], "BEL CSSR Curve": va["BEL_NOTIONAL"]}
    missing: list[str] = [name for name, value in values.items() if value is None]
    if missing and len(missing) < len(values):
        problems.append(f"the values for the CSSR must be given all together or not at all; missing {', '.join(missing)}")
    if not missing and notional_va(va["RCS_CU"]) == 0:
        problems.append(f"the notional VA (85% of the Currency RCS, rounded to whole basis points) is 0, "
                        f"so the CSSR cannot be calculated (Currency RCS is {va['RCS_CU']})")

    if va["MACROVA_SELECTED"]:
        if va["RCS_CO"] is None:
            problems.append("the macro VA needs the Country RCS")
        if va["DTTA"] is None:
            problems.append("the macro VA needs the Debt to Total Assets ratio")
        elif not 0 <= va["DTTA"] <= 1:
            problems.append(f"Debt to Total Assets must be between 0 and 1 (is {va['DTTA']})")
        if va["USA"] is not None:
            problems.append("the USA and the macro VA cannot be used together (PDF 9.3.4)")
    if va["USA"] is not None and va["USA"] > 1.05:
        problems.append(f"the USA must be at most 105% (PDF 9.3.5) (is {va['USA']})")

    if problems:
        raise ValueError("invalid VA inputs:\n  - " + "\n  - ".join(problems))
