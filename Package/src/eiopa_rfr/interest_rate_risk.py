# Generated from Code/interest_rate_risk.py by Package/sync_code.py: edit that file and run the script again.
"""Interest rate risk: the risk-free term structures after the upward and downward shocks.

Python port of the sheets "IRR Up Shock" and "IRR Down Shock" of
"RFR extrapolation and VA calculation (19 May 2026).xlsm". Both sheets have the same columns; the
two LLFR cells sit in L19 and O19 on 'IRR Up Shock' and one column to the left on 'IRR Down Shock':

    H11        UFR                        UFR + 0.15% (up), UFR - 0.15% (down)
    H12        VA                         total VA of sheet 'VA Term Structures'
    L19 / K19  LLFR IRR                   GetLLFR of the shocked curve (column K)
    O19 / N19  LLFR* IRR                  LLFR IRR + w_FSP * ln(1 + VA)
    I22:I171   Bootstrapped Zero Rate AC  BootstrapCurve as on 'Input Data & Extrapolation',
                                          but annually compounded
    J22:J171   IRR Shock AC               r * (1 + s_up) + b_up, or r * (1 - s_down) - b_down
    K22:K171   IRR Shock CC               ln(1 + J)
    L22:L171   IRR RFR                    ROUND(Extrapolation(K, "Z", ...), 5); down: at least the floor
    M22:M171   Fwd IRR Shock CC           one-year forward: t * K_t - (t - 1) * K_(t-1)
    N22:N171   Fwd + VA CC                M + ln(1 + VA)
    O22:O171   IRR with VA                ROUND(Extrapolation(N, "F", ...), 5); down: at least the floor

The shock parameters are those of sheet Parameters G16:L166: for each maturity the relative
shocks s_up and s_down, the absolute shocks b_up and b_down and the floor of the downward shock.

The shocks apply to the annually compounded bootstrapped zero rates (after the CRA), the
convention in which the basic risk-free rates are defined and published; the formula gives a
different result on continuously compounded rates. Only the liquid part is shocked: the shocked
rates end at the last DLT maturity, as the bootstrapped curve does. They are then converted to
continuous compounding, because the LLFR and the extrapolation work on continuously compounded
rates, and extrapolated with the shocked LLFR and a UFR shifted by 15 basis points. The floor of
the downward shock is applied last, to the rounded extrapolated rates; it does not change the
LLFR or the forwards with VA.

As on the sheets, the curve with VA starts from the unrounded shocked rates (column K), whereas
the curve with VA on sheet 'VA Term Structures' starts from the rounded basic RFR.

Names follow the workbook's named ranges without the R./V. prefix (IRR_SM_UP, RM_FLOOR, ...); nan
plays the role of Excel's #N/A.
"""

import math
from typing import TypedDict

from .calculation import (Curve, Flags, SheetInputs, SheetResults, bootstrap_curve, excel_round, extrapolation,
                         get_llfr, last_liquid_point)

# Shift of the UFR in the shocked curves: H11 is V.UFR + 0.0015 (up) and V.UFR - 0.0015 (down)
UFR_SHIFT: float = 0.0015

# Column headings of the shock parameters on sheet Parameters (H16:L16), per IRRInputs field;
# also the column names of irr_parameters.csv
SHEET_COLUMNS: dict[str, str] = {"IRR_SM_UP": "smup", "IRR_BM_UP": "bmup", "IRR_SM_DOWN": "smdown",
                                 "IRR_BM_DOWN": "bmdown", "RM_FLOOR": "rmfloor"}


class IRRInputs(TypedDict):
    """Parameters of the interest rate shocks (sheet Parameters G16:L166), {maturity: value}."""
    IRR_SM_UP: Curve        # H17:H166 smup     relative upward shock
    IRR_BM_UP: Curve        # I17:I166 bmup     absolute upward shock
    IRR_SM_DOWN: Curve      # J17:J166 smdown   relative downward shock
    IRR_BM_DOWN: Curve      # K17:K166 bmdown   absolute downward shock
    RM_FLOOR: Curve         # L17:L166 rmfloor  floor of the rates after the downward shock


class IRRResults(TypedDict):
    """Calculated cells of sheet 'IRR Up Shock' or 'IRR Down Shock'.

    VA, LLFR_STAR and IRR_VA_CURVE are nan without a VA.
    """
    UFR: float                      # H11 (V.UFR_UP, V.UFR_DOWN)
    VA: float                       # H12
    LLFR: float                     # L19 / K19 (V.LLFR_IRR_UP, V.LLFR_IRR_DOWN)
    LLFR_STAR: float                # O19 / N19 (V.LLFR_STAR_IRR_UP, V.LLFR_STAR_IRR_DOWN)
    BOOTSTRAPPED_CURVE_AC: Curve    # I22:I171; nan after the last DLT maturity
    IRR_SHOCK_AC: Curve             # J22:J171; nan after the last DLT maturity
    IRR_SHOCK_CC: Curve             # K22:K171 (R.IRR_SHOCK_UP_CC, R.IRR_SHOCK_DOWN_CC)
    IRR_RFR: Curve                  # L22:L171
    IRR_VA_CURVE: Curve             # O22:O171


# --------------------------------------------------------------------------
# Building blocks
# --------------------------------------------------------------------------

def shock_up(rfr: float, sm_up: float, bm_up: float) -> float:
    """The rate after the upward shock (IRR Up Shock J22: LET(..., rfr * (1 + sm_up) + bm_up)).

    Args:
        rfr: the bootstrapped zero rate, annually compounded.
        sm_up: relative upward shock of the maturity.
        bm_up: absolute upward shock of the maturity.

    Returns:
        The shocked rate, annually compounded.
    """
    return rfr * (1 + sm_up) + bm_up


def shock_down(rfr: float, sm_down: float, bm_down: float) -> float:
    """The rate after the downward shock (IRR Down Shock J22: LET(..., rfr * (1 - sm_down) - bm_down)).

    Args:
        rfr: the bootstrapped zero rate, annually compounded.
        sm_down: relative downward shock of the maturity.
        bm_down: absolute downward shock of the maturity.

    Returns:
        The shocked rate, annually compounded; it may be negative.
    """
    return rfr * (1 - sm_down) - bm_down


def apply_floor(curve: Curve, RM_FLOOR: Curve) -> Curve:
    """Raise every rate to at least the floor of its maturity (IRR Down Shock L22 and O22).

    The worksheet's BYROW(HSTACK(rates, R.RM_FLOOR), LAMBDA(r, MAX(r))).

    Args:
        curve: {maturity: rate}.
        RM_FLOOR: {maturity: floor}; must contain every maturity of ``curve``.

    Returns:
        {maturity: max(rate, floor)}.
    """
    return {t: max(r, RM_FLOOR[t]) for t, r in curve.items()}


def bootstrapped_curve_ac(inputs: SheetInputs) -> Curve:
    """The bootstrapped zero rates, annually compounded (both sheets, column I).

    The same BootstrapCurve as on sheet 'Input Data & Extrapolation' (column K), with output
    compounding "A" instead of "C". It is calculated afresh rather than converted from the
    continuously compounded curve: for swaps BootstrapSwaps works in annual compounding, and
    exp(ln(1 + z)) - 1 does not always give z back to the last bit.

    Args:
        inputs: the inputs of the basic RFR.

    Returns:
        {maturity: zero rate} for maturities 1..MAX_MATURITY; nan after the last DLT maturity.
    """
    MAX_MATURITY: int = inputs["MAX_MATURITY"]
    DLT: Flags = {t: inputs["dlt"].get(t, 0) for t in range(1, MAX_MATURITY + 1)}
    # I22: LET(LLP, MAX(FILTER(R.MATURITY, R.DLT=1)),
    #          BootstrapCurve(V.INSTRUMENT, ANCHORARRAY(H22), R.DLT, "A", V.CRA, V.COUPON_FREQ, LLP, "A", "Z"))
    return bootstrap_curve(inputs["Instrument"], inputs["Rate"], DLT, "A", inputs["CRA"], inputs["CouponFreq"],
                           last_liquid_point(DLT), "A", "Z", MAX_MATURITY)


def shocked_term_structures(inputs: SheetInputs, res: SheetResults, BOOTSTRAPPED_CURVE_AC: Curve,
                            IRR_SHOCK_AC: Curve, UFR: float, VA: float | None,
                            RM_FLOOR: Curve | None) -> IRRResults:
    """Extrapolate shocked rates, without and with VA (both sheets, columns K:O and row 19).

    The steps follow the worksheet columns:
        K   IRR Shock CC       ln(1 + J)
        L19 LLFR IRR           GetLLFR(K, R.DLT, R.LLFR_WEIGHTS)
        L   IRR RFR            ROUND(Extrapolation(K, "Z", FSP, UFR, LLFR IRR, alpha, "A"), 5)
        M   Fwd IRR Shock CC   t * K_t - (t - 1) * K_(t-1)
        N   Fwd + VA CC        M + ln(1 + VA)
        O19 LLFR* IRR          LLFR IRR + w_FSP * ln(1 + VA)
        O   IRR with VA        ROUND(Extrapolation(N, "F", FSP, UFR, LLFR* IRR, alpha, "A"), 5)
    With a floor (downward shock), L and O are raised to at least the floor of each maturity.

    Args:
        inputs: the inputs of the basic RFR (DLT, LLFR weights, Convergence, Max Maturity are used).
        res: the basic RFR, from calculation.calculate_sheet (the FSP is used).
        BOOTSTRAPPED_CURVE_AC: column I, from bootstrapped_curve_ac.
        IRR_SHOCK_AC: column J, the shocked rates, annually compounded; nan after the last DLT maturity.
        UFR: the shifted UFR of the sheet (H11), annually compounded.
        VA: the total VA as a decimal (0.0034 = 34 basis points); None without a VA.
        RM_FLOOR: {maturity: floor} for the downward shock; None for the upward shock.

    Returns:
        IRRResults; the curves are for maturities 1..MAX_MATURITY.

    Raises:
        ValueError: if a shocked rate is -100% or less (math domain error of ln(1 + J)).
    """
    MAX_MATURITY: int = inputs["MAX_MATURITY"]
    DLT: Flags = {t: inputs["dlt"].get(t, 0) for t in range(1, MAX_MATURITY + 1)}
    LLFR_WEIGHTS: Curve = {t: inputs["LLFRweightsIn"].get(t, 0.0) for t in range(1, MAX_MATURITY + 1)}
    FSP: int = res["FSP"]
    alpha: float = inputs["alpha"]

    # K22: LN(1+J22)
    IRR_SHOCK_CC: Curve = {t: math.log(1 + IRR_SHOCK_AC[t]) for t in range(1, MAX_MATURITY + 1)}
    # L19: GetLLFR(R.IRR_SHOCK_UP_CC, R.DLT, R.LLFR_WEIGHTS)
    LLFR: float = get_llfr(IRR_SHOCK_CC, DLT, LLFR_WEIGHTS)
    # L22: ROUND(Extrapolation(R.IRR_SHOCK_UP_CC, "Z", V.FSP, V.UFR_UP, V.LLFR_IRR_UP, V.ALPHA, "A"), 5),
    # on the down sheet inside BYROW(HSTACK(..., R.RM_FLOOR), LAMBDA(r, MAX(r)))
    IRR_RFR: Curve = excel_round(extrapolation(IRR_SHOCK_CC, "Z", FSP, UFR, LLFR, alpha, "A", MAX_MATURITY), 5)
    if RM_FLOOR is not None:
        IRR_RFR = apply_floor(IRR_RFR, RM_FLOOR)

    LLFR_STAR: float = math.nan
    IRR_VA_CURVE: Curve = {t: math.nan for t in range(1, MAX_MATURITY + 1)}
    if VA is not None:
        # M22: K22; M23: G23*K23-G22*K22
        FWD: Curve = {t: t * IRR_SHOCK_CC[t] - (t - 1) * IRR_SHOCK_CC[t - 1] if t > 1 else IRR_SHOCK_CC[1]
                      for t in range(1, MAX_MATURITY + 1)}
        # N22: M22+LN(1+V.TOTAL_VA)
        FWD_VA: Curve = {t: FWD[t] + math.log(1 + VA) for t in range(1, MAX_MATURITY + 1)}
        # O19: LET(FIRST_LLFR_WEIGHT, INDEX(R.LLFR_WEIGHTS, V.FSP), V.LLFR_IRR_UP + FIRST_LLFR_WEIGHT * LN(1+$H$12))
        LLFR_STAR = LLFR + LLFR_WEIGHTS[FSP] * math.log(1 + VA)
        # O22: ROUND(Extrapolation(N22:N171, "F", V.FSP, V.UFR_UP, V.LLFR_STAR_IRR_UP, V.ALPHA, "A"), 5),
        # on the down sheet inside BYROW(HSTACK(..., R.RM_FLOOR), LAMBDA(r, MAX(r)))
        IRR_VA_CURVE = excel_round(extrapolation(FWD_VA, "F", FSP, UFR, LLFR_STAR, alpha, "A", MAX_MATURITY), 5)
        if RM_FLOOR is not None:
            IRR_VA_CURVE = apply_floor(IRR_VA_CURVE, RM_FLOOR)

    return {"UFR": UFR, "VA": math.nan if VA is None else VA, "LLFR": LLFR, "LLFR_STAR": LLFR_STAR,
            "BOOTSTRAPPED_CURVE_AC": BOOTSTRAPPED_CURVE_AC, "IRR_SHOCK_AC": IRR_SHOCK_AC,
            "IRR_SHOCK_CC": IRR_SHOCK_CC, "IRR_RFR": IRR_RFR, "IRR_VA_CURVE": IRR_VA_CURVE}


# --------------------------------------------------------------------------
# Sheets "IRR Up Shock" and "IRR Down Shock"
# --------------------------------------------------------------------------

def calculate_irr_up_sheet(inputs: SheetInputs, res: SheetResults, irr: IRRInputs,
                           TOTAL_VA: float | None = None) -> IRRResults:
    """Calculate the sheet 'IRR Up Shock': the term structures after the upward shock, without and with VA.

    The steps follow the worksheet formulas:
        H11      UFR + 0.15%
        I        bootstrapped_curve_ac
        J        shock_up with the maturity's s_up and b_up
        K:O, row 19   shocked_term_structures, without a floor

    Args:
        inputs: the inputs of the basic RFR.
        res: the basic RFR, from calculation.calculate_sheet.
        irr: the shock parameters; IRR_SM_UP and IRR_BM_UP are needed for maturities 1..MAX_MATURITY.
        TOTAL_VA: the total VA, from volatility_adjustment.calculate_va_sheet; None for the curve
            without VA only.

    Returns:
        IRRResults.
    """
    BOOTSTRAPPED_CURVE_AC: Curve = bootstrapped_curve_ac(inputs)
    # J22: LET(m, G22, rfr, I22, sm_up, INDEX(R.IRR_SM_UP, m), bm_up, INDEX(R.IRR_BM_UP, m),
    #          rfr * (1 + sm_up) + bm_up)
    IRR_SHOCK_AC: Curve = {m: shock_up(rfr, irr["IRR_SM_UP"][m], irr["IRR_BM_UP"][m])
                           for m, rfr in BOOTSTRAPPED_CURVE_AC.items()}
    # H11: V.UFR + 0.0015
    return shocked_term_structures(inputs, res, BOOTSTRAPPED_CURVE_AC, IRR_SHOCK_AC, inputs["UFR"] + UFR_SHIFT,
                                   TOTAL_VA, None)


def calculate_irr_down_sheet(inputs: SheetInputs, res: SheetResults, irr: IRRInputs,
                             TOTAL_VA: float | None = None) -> IRRResults:
    """Calculate the sheet 'IRR Down Shock': the term structures after the downward shock, without and with VA.

    The steps follow the worksheet formulas:
        H11      UFR - 0.15%
        I        bootstrapped_curve_ac
        J        shock_down with the maturity's s_down and b_down
        K:O, row 19   shocked_term_structures, with the floor RM_FLOOR on columns L and O

    Args:
        inputs: the inputs of the basic RFR.
        res: the basic RFR, from calculation.calculate_sheet.
        irr: the shock parameters; IRR_SM_DOWN, IRR_BM_DOWN and RM_FLOOR are needed for maturities
            1..MAX_MATURITY.
        TOTAL_VA: the total VA, from volatility_adjustment.calculate_va_sheet; None for the curve
            without VA only.

    Returns:
        IRRResults.
    """
    BOOTSTRAPPED_CURVE_AC: Curve = bootstrapped_curve_ac(inputs)
    # J22: LET(m, G22, rfr, I22, sm_down, INDEX(R.IRR_SM_DOWN, m), bm_down, INDEX(R.IRR_BM_DOWN, m),
    #          rfr * (1 - sm_down) - bm_down)
    IRR_SHOCK_AC: Curve = {m: shock_down(rfr, irr["IRR_SM_DOWN"][m], irr["IRR_BM_DOWN"][m])
                           for m, rfr in BOOTSTRAPPED_CURVE_AC.items()}
    # H11: V.UFR-0.0015
    return shocked_term_structures(inputs, res, BOOTSTRAPPED_CURVE_AC, IRR_SHOCK_AC, inputs["UFR"] - UFR_SHIFT,
                                   TOTAL_VA, irr["RM_FLOOR"])
