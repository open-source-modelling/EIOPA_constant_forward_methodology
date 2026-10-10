"""Validation of the inputs before the calculation."""

from calculation import INSTRUMENTS, Curve, SheetInputs


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
