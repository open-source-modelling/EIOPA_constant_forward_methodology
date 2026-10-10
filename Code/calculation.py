"""Calculation of the EIOPA basic risk-free interest rate term structure from swaps or government bonds.

Python port of the VBA module ``mExtrapolation`` and of the formulas on sheet
"Input Data & Extrapolation" of "RFR extrapolation and VA calculation (19 May 2026).xlsm":

    H11      Instrument Type                       "SWP" (swaps) or "GVT" (government bonds)
    H13      FSP                                   first maturity with a positive LLFR weight
    K25:K174 Bootstrapped Zero Rate CC             VBA BootstrapCurve: BootstrapSwaps for swaps,
                                                   BootstrapZeros for government bonds
    L22      LLFR                                  VBA GetLLFR
    L25:L174 Extrapolated Zero Rate AC / Basic RFR VBA Extrapolation, rounded to 5 decimals

Naming follows the VBA module mExtrapolation (same parameter and local variable names);
VBA arrays (Option Base 1) and Excel ranges are dicts keyed 1, 2, ... (for ranges the key
is the maturity in years), and VBA's #N/A is math.nan. Worksheet-level names follow the
workbook's named ranges without the R./V. prefix.

This module also defines the data types shared by the other modules.
"""

import math
from typing import TypedDict

# Longest maturity of the curve when none is given; the VBA writes the literal 150 instead.
# The value used in a calculation comes from parameters.csv ("Max Maturity", SheetInputs MAX_MATURITY).
DEFAULT_MAX_MATURITY: int = 150

# Instrument types of cell H11 (list Parameters!G9:G10) and what they stand for
INSTRUMENTS: dict[str, str] = {"SWP": "swaps", "GVT": "government bonds"}

Curve = dict[int, float]   # {maturity or array index: rate / value}
Flags = dict[int, int]     # {maturity: 1 or 0}, e.g. the DLT column


class SheetInputs(TypedDict):
    """Inputs on sheet 'Input Data & Extrapolation'.

    Field names are the names of the VBA parameters that receive them.
    """
    Instrument: str         # H11 Instrument Type, "SWP" or "GVT"  (BootstrapCurve Instrument)
    CouponFreq: int | None  # H12 Coupon Frequency; swaps only, None for government bonds
                            #                               (BootstrapCurve CouponFreq)
    CRA: float              # H16 CRA, basis points         (BootstrapCurve CRA)
    UFR: float              # H14 UFR                       (Extrapolation UFR)
    alpha: float            # H15 Convergence               (Extrapolation alpha)
    dlt: Flags              # H25:H174 DLT                  (BootstrapCurve dlt)
    LLFRweightsIn: Curve    # I25:I174 LLFR Weight          (GetLLFR LLFRweightsIn)
    Rate: Curve             # J25:J174 Input Rate: par swap rates (SWP) or zero rates, annually
                            # compounded (GVT); maturities without a rate are missing
                            #                               (BootstrapCurve Rate)
    MAX_MATURITY: int       # longest maturity of the curve (the VBA's literal 150; 150 rows on the sheet)


class SheetResults(TypedDict):
    """Calculated cells of sheet 'Input Data & Extrapolation'.

    FSP and LLFR are VBA names (Extrapolation parameters); the curves have no VBA name of
    their own (both are ``zero`` in the VBA) and use the workbook's named ranges.
    """
    FSP: int                     # H13
    LLFR: float                  # L22
    BOOTSTRAPPED_CURVE: Curve    # K25:K174 (named range R.BOOTSTRAPPED_CURVE)
    BASIC_RFR: Curve             # L25:L174 (named range R.BASIC_RFR)


# --------------------------------------------------------------------------
# VBA module mExtrapolation
# --------------------------------------------------------------------------

def newton_raphson_forward_swap(fwguess: float, swapt1: float, m: float, c: float,
                                tol: float = 1e-15, max_iter: int = 500) -> float:
    """Solve the constant periodic forward rate of a swap interval with Newton-Raphson (VBA NewtonRaphsonForwardSwap).

    Finds fw such that  swapt1 * (1 - (1+fw)^-m) / fw + (1+fw)^-m = c,
    which is the par condition of the swap divided by the discount factor at the start of the
    interval (PDF D.5.9 and D.6). The iteration stops when |swapt1 * (1 - (1+fw)^-m) / fw + (1+fw)^-m - c|
    is below ``tol`` or after ``max_iter`` steps; as in the VBA, the last value is returned even
    if it has not converged.

    Args:
        fwguess: starting value for the periodic forward rate.
        swapt1: periodic swap coupon, (swap rate - CRA) / coupon frequency.
        m: number of coupon periods in the interval (coupon frequency x years).
        c: target value; 1 for the first DLT maturity, otherwise
            (1 - swapt1 * sum of earlier discount factors) / discount factor at the start of the interval.
        tol: convergence tolerance; 1e-15 as in the VBA.
        max_iter: maximum number of iterations; 500 as in the VBA.

    Returns:
        The periodic (not annualised) forward rate of the interval.
    """
    fw: float = fwguess
    for _ in range(max_iter):
        temp: float = (1 + fw) ** -m
        fx: float = swapt1 * (1 - temp) / fw + temp - c
        if abs(fx) < tol:
            break
        temp /= 1 + fw
        dfx: float = swapt1 * ((1 + (m + 1) * fw) * temp - 1) / fw ** 2 - m * temp
        fw -= fx / dfx
    return fw


def bootstrap_swaps(SwapRatesInit: Curve, dlt: Flags, CouponFreq: int, CRA: float, MaxTenor: int,
                    CompoundingOut: str, RateType: str, MAX_MATURITY: int = DEFAULT_MAX_MATURITY) -> Curve:
    """Bootstrap zero rates from par swap rates under the constant forward assumption (VBA BootstrapSwaps).

    Uses the rates at maturities with DLT = 1 and a rate, after deducting the CRA. For the first
    of these maturities the periodic forward equals the periodic coupon (PDF D.4); for each
    following one the constant forward of the interval is solved from the par condition with
    newton_raphson_forward_swap (PDF D.5). If MaxTenor lies beyond the last DLT maturity, the
    last forward is carried forward flat. Maturities after MaxTenor are nan (#N/A).

    Args:
        SwapRatesInit: {maturity: market par swap rate}; an empty cell is a missing key.
        dlt: {maturity: 1 or 0}, DLT flags.
        CouponFreq: coupon payments per year.
        CRA: credit risk adjustment in basis points, deducted from every swap rate.
        MaxTenor: last maturity to calculate.
        CompoundingOut: "C" for continuously compounded zero rates, anything else for annually
            compounded ones.
        RateType: ignored, exactly as in the VBA (zero rates are always returned).
        MAX_MATURITY: not a VBA parameter; replaces the VBA's literal 150 (length of the input
            range and of the result).

    Returns:
        {maturity: zero rate} for maturities 1..MAX_MATURITY; nan after MaxTenor.

    Raises:
        KeyError: if no maturity has DLT = 1 and a rate (prevented by validation.validate_inputs).

    Newton-Raphson starting guess: follows the VBA, not the PDF.
    - First DLT maturity: the periodic swap coupon (s - CRA) / m (same in VBA and PDF D.4.7).
    - Later DLT maturities: VBA uses ``forward(t - 1) / CouponFreq``, where ``t`` is the
      position of the DLT point in the list (1, 2, ...) but ``forward`` is indexed by
      maturity in years and holds annual forwards. So the guess is the annual forward of
      year t - 1 divided by m, not the periodic forward of the previous interval that
      PDF D.6.3 prescribes (e.g. for the 20y point, the 15th DLT point, the guess comes
      from the year-14 forward instead of the 15y-20y one).
    The guess only sets where the iteration starts; the root, and hence the curve, is the
    same either way within the 1e-15 tolerance.
    """
    SwapRates: Curve = {}
    SwapTenors: dict[int, int] = {}
    forward: Curve = {}
    discount: Curve = {}
    zero: Curve = {}

    m: int = 0
    for k in range(1, MAX_MATURITY + 1):
        if SwapRatesInit.get(k) is not None and dlt.get(k) == 1:
            m = m + 1
            SwapTenors[m] = k
            SwapRates[m] = SwapRatesInit[k] - CRA / 10000

    m = SwapTenors[1]
    fwtemp: float = newton_raphson_forward_swap(SwapRates[1] / CouponFreq, SwapRates[1] / CouponFreq, CouponFreq * m, 1)
    for k in range(1, m + 1):
        forward[k] = ((1 + fwtemp) ** CouponFreq) - 1
        if k > 1:
            discount[k] = discount[k - 1] / (1 + forward[k])
        else:
            discount[k] = 1 / (1 + forward[k])
        zero[k] = (1 / discount[k]) ** (1 / k) - 1

    dtemp: float = 1 / (1 + fwtemp)
    sumdiscount: float = (1 - dtemp ** (m * CouponFreq)) / fwtemp
    for t in range(2, len(SwapTenors) + 1):            # For t = 2 To UBound(SwapTenors)
        m = SwapTenors[t] - SwapTenors[t - 1]
        fwtemp = newton_raphson_forward_swap(forward[t - 1] / CouponFreq, SwapRates[t] / CouponFreq, CouponFreq * m,
                                             (1 - SwapRates[t] / CouponFreq * sumdiscount) / discount[SwapTenors[t - 1]])
        dtemp = 1 / (1 + fwtemp)
        sumdiscount = sumdiscount + discount[SwapTenors[t - 1]] * (1 - dtemp ** (CouponFreq * m)) / fwtemp
        for k in range(1, m + 1):
            forward[SwapTenors[t - 1] + k] = (1 + fwtemp) ** CouponFreq - 1
            discount[SwapTenors[t - 1] + k] = discount[SwapTenors[t - 1] + k - 1] / (1 + forward[SwapTenors[t - 1] + k])
            zero[SwapTenors[t - 1] + k] = (1 / discount[SwapTenors[t - 1] + k]) ** (1 / (SwapTenors[t - 1] + k)) - 1
    t: int = len(SwapTenors) + 1                        # value of t after the VBA For loop

    if SwapTenors[t - 1] < MaxTenor:  # eventual extrapolation to the right
        for k in range(1, MaxTenor - SwapTenors[t - 1] + 1):
            forward[SwapTenors[t - 1] + k] = forward[SwapTenors[t - 1] + k - 1]
            discount[SwapTenors[t - 1] + k] = discount[SwapTenors[t - 1] + k - 1] / (1 + forward[SwapTenors[t - 1] + k])
            zero[SwapTenors[t - 1] + k] = (1 / discount[SwapTenors[t - 1] + k]) ** (1 / (SwapTenors[t - 1] + k)) - 1

    if MaxTenor < MAX_MATURITY:                         # If MaxTenor < 150 Then
        for k in range(MaxTenor + 1, MAX_MATURITY + 1):
            forward[k] = math.nan
            zero[k] = math.nan

    if CompoundingOut == "C":
        for k in range(1, MaxTenor + 1):
            forward[k] = math.log(1 + forward[k])
            zero[k] = math.log(1 + zero[k])

    return zero


def bootstrap_zeros(ZeroRatesInit: Curve, dlt: Flags, CompoundingIn: str, CRA: float, MaxTenor: int,
                    CompoundingOut: str, RateType: str, MAX_MATURITY: int = DEFAULT_MAX_MATURITY) -> Curve:
    """Bootstrap zero rates from zero rates under the constant forward assumption (VBA BootstrapZeros).

    Uses the rates at maturities with DLT = 1 and a rate, after deducting the CRA; with
    CompoundingIn "A" they are converted to continuous compounding. Before the first of these
    maturities its rate is held flat; between two of them the continuously compounded forward is
    constant, f = (b * z_b - a * z_a) / (b - a), so the DLT rates are reproduced exactly. If
    MaxTenor lies beyond the last DLT maturity, the last forward is carried forward flat.
    Maturities after MaxTenor are nan (#N/A).

    Args:
        ZeroRatesInit: {maturity: market zero rate}; an empty cell is a missing key.
        dlt: {maturity: 1 or 0}, DLT flags.
        CompoundingIn: "A" for annually compounded input rates, anything else for continuously
            compounded ones.
        CRA: credit risk adjustment in basis points, deducted from every rate before the
            conversion to continuous compounding.
        MaxTenor: last maturity to calculate.
        CompoundingOut: "A" for annually compounded zero rates, anything else for continuously
            compounded ones.
        RateType: ignored, exactly as in the VBA (zero rates are always returned).
        MAX_MATURITY: not a VBA parameter; replaces the VBA's literal 150 (length of the input
            range and of the result).

    Returns:
        {maturity: zero rate} for maturities 1..MAX_MATURITY; nan after MaxTenor.

    Raises:
        KeyError: if no maturity has DLT = 1 and a rate (prevented by validation.validate_inputs).
    """
    ZeroRates: Curve = {}
    ZeroTenors: dict[int, int] = {}
    forward: Curve = {}
    discount: Curve = {}
    zero: Curve = {}

    m: int = 0
    for k in range(1, MAX_MATURITY + 1):                # For k = 1 To ZeroRatesInit.Rows.Count
        if ZeroRatesInit.get(k) is not None and dlt.get(k) == 1:
            m = m + 1
            ZeroTenors[m] = k
            ZeroRates[m] = ZeroRatesInit[k] - CRA / 10000
            if CompoundingIn == "A":
                ZeroRates[m] = math.log(1 + ZeroRates[m])

    # eventual extrapolation to the left
    for k in range(1, ZeroTenors[1] + 1):
        forward[k] = ZeroRates[1]
        zero[k] = forward[k]
        discount[k] = math.exp(-k * zero[k])

    # constant forward interpolation if applicable
    for t in range(2, len(ZeroTenors) + 1):            # For t = 2 To UBound(ZeroTenors)
        m = ZeroTenors[t] - ZeroTenors[t - 1]
        fwtemp: float = (ZeroTenors[t] * ZeroRates[t] - ZeroTenors[t - 1] * ZeroRates[t - 1]) / m
        for k in range(1, m + 1):
            forward[ZeroTenors[t - 1] + k] = fwtemp
            discount[ZeroTenors[t - 1] + k] = discount[ZeroTenors[t - 1] + k - 1] * math.exp(-fwtemp)
            zero[ZeroTenors[t - 1] + k] = -math.log(discount[ZeroTenors[t - 1] + k]) / (ZeroTenors[t - 1] + k)
    t: int = len(ZeroTenors) + 1                        # value of t after the VBA For loop

    if ZeroTenors[t - 1] < MaxTenor:  # eventual extrapolation to the right
        for k in range(1, MaxTenor - ZeroTenors[t - 1] + 1):
            forward[ZeroTenors[t - 1] + k] = forward[ZeroTenors[t - 1] + k - 1]
            discount[ZeroTenors[t - 1] + k] = discount[ZeroTenors[t - 1] + k - 1] * math.exp(-forward[ZeroTenors[t - 1] + k])
            zero[ZeroTenors[t - 1] + k] = -math.log(discount[ZeroTenors[t - 1] + k]) / (ZeroTenors[t - 1] + k)

    if MaxTenor < MAX_MATURITY:                         # If MaxTenor < 150 Then
        for k in range(MaxTenor + 1, MAX_MATURITY + 1):
            forward[k] = math.nan
            zero[k] = math.nan

    if CompoundingOut == "A":
        for k in range(1, MaxTenor + 1):
            forward[k] = math.exp(forward[k]) - 1
            zero[k] = math.exp(zero[k]) - 1

    return zero


def bootstrap_curve(Instrument: str, Rate: Curve, dlt: Flags, CompoundingIn: str, CRA: float, CouponFreq: int | None,
                    MaxTenor: int, CompoundingOut: str, RateType: str, MAX_MATURITY: int = DEFAULT_MAX_MATURITY) -> Curve:
    """Bootstrap with the method of the instrument type (VBA BootstrapCurve).

    "SWP" calls bootstrap_swaps; any other instrument type, i.e. "GVT", calls bootstrap_zeros.

    Args:
        Instrument: "SWP" for swaps, "GVT" for government bonds.
        Rate: {maturity: market rate}: par swap rates or zero rates; an empty cell is a missing key.
        dlt: {maturity: 1 or 0}, DLT flags.
        CompoundingIn: compounding of zero rates, see bootstrap_zeros; not used for swaps.
        CRA: credit risk adjustment in basis points.
        CouponFreq: coupon payments per year; used only for swaps, may be None otherwise.
        MaxTenor: last maturity to calculate.
        CompoundingOut: "C" for continuously compounded zero rates, "A" for annually compounded
            ones (see bootstrap_swaps and bootstrap_zeros for other values).
        RateType: ignored, as in the VBA.
        MAX_MATURITY: length of the input range and of the result.

    Returns:
        {maturity: zero rate} for maturities 1..MAX_MATURITY; nan after MaxTenor.

    Raises:
        ValueError: for swaps without a coupon frequency (prevented by validation.validate_inputs).
    """
    if Instrument == "SWP":
        if CouponFreq is None:
            raise ValueError("swaps need a coupon frequency")
        return bootstrap_swaps(Rate, dlt, CouponFreq, CRA, MaxTenor, CompoundingOut, RateType, MAX_MATURITY)
    else:
        return bootstrap_zeros(Rate, dlt, CompoundingIn, CRA, MaxTenor, CompoundingOut, RateType, MAX_MATURITY)


def get_llfr(InputRates: Curve, DLTin: Flags, LLFRweightsIn: Curve) -> float:
    """Calculate the last liquid forward rate (VBA GetLLFR; PDF 8.3 and 8.5.6).

    The first maturity with a positive weight is the FSP. The LLFR is the weighted average of
    the forward rate from the last DLT maturity before the FSP to the FSP, and of the forward
    rates from the FSP to each later maturity with a positive weight:
        LLFR = w_FSP * f(LLPbeforeFSP, FSP) + sum over k of w_k * f(FSP, t_k),
        with f(a, b) = (b * z_b - a * z_a) / (b - a).

    Args:
        InputRates: {maturity: continuously compounded zero rate}, e.g. the result of bootstrap_curve.
        DLTin: {maturity: 1 or 0}, DLT flags.
        LLFRweightsIn: {maturity: LLFR weight}.

    Returns:
        The LLFR, continuously compounded.

    Raises:
        ValueError: if the weights do not sum to 1. Difference from the VBA: the VBA adds up the
            weights in ``sumw`` but never uses it (the sheet only shows "ERROR" in I22).
        KeyError: if there is no DLT maturity before the FSP (prevented by validation.validate_inputs).
    """
    rates: Curve = InputRates
    weight: Curve = LLFRweightsIn

    w: Curve = {}
    t: dict[int, int] = {}
    k: int = 0
    sumw: float = 0

    # getting the tenors for which weights > 0 are defined
    for i in range(1, max(weight, default=0) + 1):      # For i = 1 To UBound(weight)
        if weight.get(i, 0) > 0:
            k = k + 1
            t[k] = i  # store those tenors i for which the weight(i) > 0
            w[k] = weight[i]  # store the corresponding weights
            sumw = sumw + w[k]  # check for sum of weights equal to 1
    if abs(sumw - 1) > 1e-9:
        raise ValueError("LLFR weights must sum to 1")

    # determine last liquid point before FSP (note: FSP equals t(1), i.e. the first tenor for which a positive weight has been defined)
    dlt: Flags = DLTin
    LLPbeforeFSP: int = 0
    for i in range(1, t[1]):
        if dlt.get(i) == 1:
            LLPbeforeFSP = i

    tempLLFR: float = w[1] * (t[1] * rates[t[1]] - LLPbeforeFSP * rates[LLPbeforeFSP]) / (t[1] - LLPbeforeFSP)
    for i in range(2, len(t) + 1):                      # For i = 2 To UBound(t)
        tempLLFR = tempLLFR + w[i] * (t[i] * rates[t[i]] - t[1] * rates[t[1]]) / (t[i] - t[1])

    return tempLLFR


def extrapolation(InputRatesCC: Curve, RateType: str, FSP: int, UFR: float, LLFR: float, alpha: float,
                  Compounding: str, MAX_MATURITY: int = DEFAULT_MAX_MATURITY) -> Curve:
    """Extend a zero curve beyond the FSP towards the UFR (VBA Extrapolation; PDF 8.5.5, 8.5.7 and 10.2).

    Up to the FSP the zero rates are the input zero rates (RateType "Z"), or the running average
    of the input one-year forward rates (any other RateType). Beyond the FSP, with h = t - FSP:
        f(FSP, t) = ln(1 + UFR) + (LLFR - ln(1 + UFR)) * (1 - exp(-alpha * h)) / (alpha * h)
        z_t       = (FSP * z_FSP + h * f(FSP, t)) / t

    Args:
        InputRatesCC: {maturity: continuously compounded rate}; only maturities 1..FSP are used.
        RateType: "Z" if the input are zero rates, otherwise one-year forward rates
            (the sheet 'Input Data & Extrapolation' uses "Z").
        FSP: first smoothing point, in years.
        UFR: ultimate forward rate, annually compounded.
        LLFR: last liquid forward rate, continuously compounded.
        alpha: convergence parameter.
        Compounding: "A" for annually compounded zero rates, anything else for continuously
            compounded ones.
        MAX_MATURITY: Length of the extrapolated curve in years.

    Returns:
        {maturity: zero rate} for maturities 1..MAX_MATURITY.
    """
    InputRates: Curve = InputRatesCC

    zero: Curve = {}
    ForwardRates: Curve = {}

    zero[1] = InputRates[1]

    if RateType == "Z":
        for i in range(2, FSP + 1):
            zero[i] = InputRates[i]
    else:
        for i in range(2, FSP + 1):
            zero[i] = ((i - 1) * zero[i - 1] + InputRates[i]) / i

    for i in range(FSP + 1, MAX_MATURITY + 1):          # For i = FSP + 1 To MAX_MATURITY
        ForwardRates[i] = math.log(1 + UFR) + (LLFR - math.log(1 + UFR)) * (1 - math.exp(-alpha * (i - FSP))) / (alpha * (i - FSP))
        zero[i] = (FSP * zero[FSP] + (i - FSP) * ForwardRates[i]) / i

    if Compounding == "A":
        for i in range(1, MAX_MATURITY + 1):            # For i = 1 To MAX_MATURITY
            zero[i] = math.exp(zero[i]) - 1

    return zero


# --------------------------------------------------------------------------
# Sheet "Input Data & Extrapolation"
# --------------------------------------------------------------------------

def excel_round(curve: Curve, digits: int = 5) -> Curve:
    """Round every value of a curve, like the worksheet's ROUND(..., 5) applied to a whole column.

    Args:
        curve: {maturity: value}.
        digits: number of decimals; 5 on the sheet.

    Returns:
        {maturity: rounded value}; nan stays nan.
    """
    return {t: round(z, digits) for t, z in curve.items()}


def calculate_sheet(Instrument: str, CouponFreq: int | None, CRA: float, UFR: float, alpha: float, dlt: Flags,
                    LLFRweightsIn: Curve, Rate: Curve, MAX_MATURITY: int) -> SheetResults:
    """Calculate the sheet 'Input Data & Extrapolation': bootstrapped curve, FSP, LLFR and basic RFR.

    The steps follow the worksheet formulas:
        K25:K174  BootstrapCurve  -> bootstrap_curve up to the last DLT maturity, continuously
                                     compounded: bootstrap_swaps for swaps, bootstrap_zeros for
                                     government bonds (input rates annually compounded)
        H13       XMATCH          -> FSP = first maturity with a positive LLFR weight
        L22       GetLLFR         -> get_llfr on the bootstrapped curve
        L25:L174  ROUND(Extrapolation(..., "A"), 5) -> extrapolation and excel_round

    Args:
        Instrument: "SWP" for swaps, "GVT" for government bonds.
        CouponFreq: coupon payments per year of the swaps; not used for government bonds (None).
        CRA: credit risk adjustment in basis points.
        UFR: ultimate forward rate, annually compounded.
        alpha: convergence parameter.
        dlt: {maturity: 1 or 0}, DLT flags.
        LLFRweightsIn: {maturity: LLFR weight}.
        Rate: {maturity: market rate}: par swap rates, or annually compounded zero rates for
            government bonds; only maturities with DLT = 1 are used.
        MAX_MATURITY: Length of the extrapolated curve in years.

    Returns:
        SheetResults with FSP, LLFR (continuously compounded), BOOTSTRAPPED_CURVE (continuously
        compounded, nan after the last DLT maturity) and BASIC_RFR (annually compounded, rounded
        to 5 decimals); both curves for maturities 1..MAX_MATURITY.

    Raises:
        ValueError: if the LLFR weights do not sum to 1, if there is no DLT maturity or no
            positive weight, or for swaps without a coupon frequency.
        KeyError: for other inputs that validation.validate_inputs rejects, e.g. no DLT maturity
            before the FSP. Validate the inputs first.
    """
    BASECURVE: Curve = Rate
    DLT: Flags = {t: dlt.get(t, 0) for t in range(1, MAX_MATURITY + 1)}
    LLFR_WEIGHTS: Curve = {t: LLFRweightsIn.get(t, 0.0) for t in range(1, MAX_MATURITY + 1)}

    # K25: BootstrapCurve(V.INSTRUMENT, R.BASECURVE, R.DLT, "A", V.CRA, V.COUPON_FREQ, LLP, "C", "Z"),
    # which calls BootstrapSwaps for instrument "SWP" and BootstrapZeros otherwise
    LLP: int = max(t for t in DLT if DLT[t] == 1)                     # MAX(FILTER(R.MATURITY, R.DLT=1))
    BOOTSTRAPPED_CURVE: Curve = bootstrap_curve(Instrument, BASECURVE, DLT, "A", CRA, CouponFreq, LLP, "C", "Z",
                                                 MAX_MATURITY)
    # H13: XMATCH(TRUE, R.LLFR_WEIGHTS > 0, 0)
    FSP: int = min(t for t in LLFR_WEIGHTS if LLFR_WEIGHTS[t] > 0)
    # L22: GetLLFR(R.BOOTSTRAPPED_CURVE, R.DLT, R.LLFR_WEIGHTS)
    LLFR: float = get_llfr(BOOTSTRAPPED_CURVE, DLT, LLFR_WEIGHTS)
    # L25: ROUND(Extrapolation(R.BOOTSTRAPPED_CURVE, "Z", V.FSP, V.UFR, V.LLFR, V.ALPHA, "A"), 5)
    BASIC_RFR: Curve = excel_round(extrapolation(BOOTSTRAPPED_CURVE, "Z", FSP, UFR, LLFR, alpha, "A", MAX_MATURITY), 5)

    return {"FSP": FSP, "LLFR": LLFR, "BOOTSTRAPPED_CURVE": BOOTSTRAPPED_CURVE, "BASIC_RFR": BASIC_RFR}
