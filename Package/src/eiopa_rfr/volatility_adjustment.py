# Generated from Code/volatility_adjustment.py by Package/sync_code.py: edit that file and run the script again.
"""Undertaking-specific volatility adjustment (VA): credit spread sensitivity ratio and term structure with VA.

Python port of the sheets "CSSR Calculation" and "VA Term Structures" of
"RFR extrapolation and VA calculation (19 May 2026).xlsm":

    CSSR Calculation
        H10        Notional VA                     ROUND(85% * RCS_CU, 4)
        L10:L11    MVFI, MVFI_NOTIONAL             fixed income assets at market value and on the CSSR curve
        L12        PVBP_MVFI                       (MVFI - MVFI_NOTIONAL) / NOTIONAL_VA
        L13:L14    BEL, BEL_NOTIONAL               best estimate at market value and on the CSSR curve
        L15        PVBP_BEL                        (BEL - BEL_NOTIONAL) / NOTIONAL_VA
        L16        CSSR                            PVBP_MVFI / PVBP_BEL, limited to 0..1
        K20        NOTIONAL_VA_LLFR                LLFR with the notional VA
        L24:L173   CSSR Curve                      basic RFR with the notional VA
    VA Term Structures
        L10        Currency VA                     85% * CSSR * (USA *) RCS_CU
        L11        W_CO                            country adjustment factor
        L12        Macro VA                        85% * CSSR * max(RCS_CO - 1.3 * RCS_CU, 0) * W_CO
        L13        Total VA                        currency VA (+ macro VA), rounded to whole basis points
        K20        TOTAL_VA_LLFR                   LLFR with the total VA
        L23:L172   VA Curve                        basic RFR with the total VA

Methodology: EIOPA-BoS-26/198 "RFR Technical Documentation", sections 9.3, 9.4 and 10.

Both curves are built in the same way (function va_curve): the VA is added to the continuously
compounded one-year forward rates of the basic RFR, which changes the curve up to the FSP, and the
first term of the LLFR is adjusted; beyond the FSP the curve is extrapolated as the basic RFR.

Names follow the workbook's named ranges without the R./V. prefix (RCS_CU, MVFI, CSSR, ...); nan
plays the role of Excel's #N/A. Rates are decimals: an RCS of 0.01 is 100 basis points.
"""

import math
from typing import TypedDict

from .calculation import Curve, SheetInputs, SheetResults, excel_round, excel_round_value, extrapolation

# Share of the risk-corrected spread in the VA (PDF 9.3.1 and 9.4.1)
VA_SHARE: float = 0.85


class VAInputs(TypedDict):
    """Inputs of the sheets 'CSSR Calculation' and 'VA Term Structures'.

    The four values for the CSSR come from the undertaking's own valuation: once on the basic RFR,
    once on the CSSR curve. They are None until that valuation has been done.
    """
    RCS_CU: float                   # Input Data & Extrapolation L10 Currency RCS
    RCS_CO: float | None            # Input Data & Extrapolation L11 Country RCS; macro VA only
    DTTA: float | None              # Input Data & Extrapolation L12 Debt to Total Assets; macro VA only
    MVFI: float | None              # CSSR Calculation L10 fixed income assets, market value
    MVFI_NOTIONAL: float | None     # CSSR Calculation L11 fixed income assets on the CSSR curve
    BEL: float | None               # CSSR Calculation L13 best estimate liabilities on the basic RFR
    BEL_NOTIONAL: float | None      # CSSR Calculation L14 best estimate liabilities on the CSSR curve
    MACROVA_SELECTED: bool          # VA Term Structures H10 Macro VA?
    USA: float | None               # VA Term Structures H12 undertaking-specific adjustment;
                                    # None if not used (H11 USA? = 0)


class CSSRResults(TypedDict):
    """Calculated cells of sheet 'CSSR Calculation'.

    PVBP_MVFI, PVBP_BEL and CSSR are None while the four values for the CSSR are missing.
    """
    NOTIONAL_VA: float              # H10
    NOTIONAL_VA_LLFR: float         # K20
    CSSR_CURVE: Curve               # L24:L173
    PVBP_MVFI: float | None         # L12
    PVBP_BEL: float | None          # L15
    CSSR: float | None              # L16


class VAResults(TypedDict):
    """Calculated cells of sheet 'VA Term Structures'."""
    CURRENCY_VA: float              # L10
    W_CO: float                     # L11; nan without country RCS and debt to total assets
    MACRO_VA: float                 # L12; nan (#N/A) if the macro VA is not selected
    TOTAL_VA: float                 # L13
    TOTAL_VA_LLFR: float            # K20
    VA_CURVE: Curve                 # L23:L172


# --------------------------------------------------------------------------
# Building blocks
# --------------------------------------------------------------------------

def va_curve(BASIC_RFR: Curve, VA: float, FSP: int, LLFR: float, FIRST_LLFR_WEIGHT: float, UFR: float,
             alpha: float) -> tuple[float, Curve]:
    """Build a risk-free term structure with a VA (PDF 10.1.6 and 10.2; both sheets, columns H:L and cell K20).

    The steps follow the worksheet columns:
        H  RFR               the basic RFR (annually compounded, rounded to 5 decimals)
        I  RFR CC            ln(1 + H)
        J  Fwd RFR CC        one-year forward: t * I_t - (t - 1) * I_(t-1)
        K  Fwd RFR CC + VA   J + ln(1 + VA)
        K20 LLFR with VA     LLFR - w_FSP * J_FSP + w_FSP * K_FSP, i.e. LLFR + w_FSP * ln(1 + VA)
        L  curve with VA     ROUND(Extrapolation(K, "F", FSP, UFR, LLFR with VA, alpha, "A"), 5)
    Only the forwards up to the FSP are used, so the VA changes the curve up to the FSP directly
    and beyond it through the LLFR.

    Args:
        BASIC_RFR: {maturity: basic RFR}, annually compounded, maturities 1..MAX_MATURITY.
        VA: the volatility adjustment as a decimal (0.0034 = 34 basis points).
        FSP: first smoothing point, in years.
        LLFR: the LLFR of the basic RFR, continuously compounded.
        FIRST_LLFR_WEIGHT: the LLFR weight at the FSP.
        UFR: ultimate forward rate, annually compounded.
        alpha: convergence parameter.

    Returns:
        (LLFR with VA, {maturity: rate with VA}); the rates are annually compounded and rounded to
        5 decimals, for the same maturities as BASIC_RFR.
    """
    MAX_MATURITY: int = max(BASIC_RFR)
    RFR_CC: Curve = {t: math.log(1 + BASIC_RFR[t]) for t in range(1, MAX_MATURITY + 1)}
    FWD_CURVE: Curve = {t: t * RFR_CC[t] - (t - 1) * RFR_CC[t - 1] if t > 1 else RFR_CC[1]
                        for t in range(1, MAX_MATURITY + 1)}
    FWD_VA_CURVE: Curve = {t: FWD_CURVE[t] + math.log(1 + VA) for t in range(1, MAX_MATURITY + 1)}
    # K20: LET(FIRST_LLFR_FORWARD, INDEX(R.FWD_CURVE, V.FSP), FIRST_LLFR_WEIGHT, INDEX(R.LLFR_WEIGHTS, V.FSP),
    #          V.LLFR - FIRST_LLFR_WEIGHT * FIRST_LLFR_FORWARD + FIRST_LLFR_WEIGHT * INDEX(R.FWD_VA_CURVE, V.FSP))
    LLFR_VA: float = LLFR - FIRST_LLFR_WEIGHT * FWD_CURVE[FSP] + FIRST_LLFR_WEIGHT * FWD_VA_CURVE[FSP]
    curve: Curve = excel_round(extrapolation(FWD_VA_CURVE, "F", FSP, UFR, LLFR_VA, alpha, "A", MAX_MATURITY), 5)
    return LLFR_VA, curve


def notional_va(RCS_CU: float) -> float:
    """The notional VA used for the CSSR curve (PDF 9.3.2 a; CSSR Calculation H10).

    85% of the currency RCS, rounded to whole basis points like the worksheet's ROUND(..., 4).

    Args:
        RCS_CU: risk-corrected currency spread published by EIOPA, as a decimal.

    Returns:
        The notional VA as a decimal.
    """
    return excel_round_value(VA_SHARE * RCS_CU, 4)


def credit_spread_sensitivity_ratio(MVFI: float, MVFI_NOTIONAL: float, BEL: float, BEL_NOTIONAL: float,
                                    NOTIONAL_VA: float) -> tuple[float, float, float]:
    """Calculate the CSSR (PDF 9.3.2 b-d; CSSR Calculation L12, L15 and L16).

    The changes in value are divided by the notional VA, as on the sheet, which gives a value per
    unit of spread but leaves the ratio unchanged:
        PVBP_MVFI = (MVFI - MVFI_NOTIONAL) / NOTIONAL_VA
        PVBP_BEL  = (BEL - BEL_NOTIONAL) / NOTIONAL_VA
        CSSR      = max(min(PVBP_MVFI / PVBP_BEL, 1), 0), and 1 if PVBP_BEL <= 0
    The PDF defines the CSSR as max(min(dFI / dBEL, 1), 0) and does not cover dBEL = 0; the
    workbook, and this function, then use 1, as for any PVBP_BEL <= 0.

    Args:
        MVFI: fixed income assets valued on the basic RFR (market value).
        MVFI_NOTIONAL: the same assets valued on the CSSR curve.
        BEL: best estimate liabilities valued on the basic RFR.
        BEL_NOTIONAL: the same liabilities valued on the CSSR curve.
        NOTIONAL_VA: the notional VA of the CSSR curve, see notional_va.

    Returns:
        (PVBP_MVFI, PVBP_BEL, CSSR).

    Raises:
        ZeroDivisionError: if the notional VA is 0 (prevented by validation.validate_va_inputs).
    """
    PVBP_MVFI: float = (MVFI - MVFI_NOTIONAL) / NOTIONAL_VA
    PVBP_BEL: float = (BEL - BEL_NOTIONAL) / NOTIONAL_VA
    CSSR: float = 1.0 if PVBP_BEL <= 0 else max(min(PVBP_MVFI / PVBP_BEL, 1), 0)
    return PVBP_MVFI, PVBP_BEL, CSSR


def country_adjustment_factor(RCS_CO: float, DTTA: float) -> float:
    """The country adjustment factor of the macro VA (PDF 9.4.2; VA Term Structures L11).

        W_CO = max(min((RCS_CO * DTTA - 0.6%) / 0.3%, 1), 0)

    Args:
        RCS_CO: risk-corrected country spread published by EIOPA, as a decimal.
        DTTA: debt-to-total-assets ratio of the country published by EIOPA, as a decimal.

    Returns:
        The factor, between 0 and 1.
    """
    return max(min((RCS_CO * DTTA - 0.006) / 0.003, 1), 0)


# --------------------------------------------------------------------------
# Sheets "CSSR Calculation" and "VA Term Structures"
# --------------------------------------------------------------------------

def calculate_cssr_sheet(inputs: SheetInputs, res: SheetResults, va: VAInputs) -> CSSRResults:
    """Calculate the sheet 'CSSR Calculation': notional VA, CSSR curve and, if the values are given, the CSSR.

    The CSSR curve is the basic RFR with a notional VA of 85% of the currency RCS. The undertaking
    values its fixed income assets and best estimate liabilities on this curve; with those values
    (MVFI_NOTIONAL, BEL_NOTIONAL) and the values on the basic RFR (MVFI, BEL), the CSSR follows.

    Args:
        inputs: the inputs of the basic RFR (UFR, Convergence and LLFR weights are used).
        res: the basic RFR, from calculation.calculate_sheet.
        va: the VA inputs.

    Returns:
        CSSRResults; PVBP_MVFI, PVBP_BEL and CSSR are None if any of the four values is missing.
    """
    FSP: int = res["FSP"]
    NOTIONAL_VA: float = notional_va(va["RCS_CU"])
    NOTIONAL_VA_LLFR, CSSR_CURVE = va_curve(res["BASIC_RFR"], NOTIONAL_VA, FSP, res["LLFR"],
                                            inputs["LLFRweightsIn"].get(FSP, 0.0), inputs["UFR"], inputs["alpha"])
    PVBP_MVFI: float | None = None
    PVBP_BEL: float | None = None
    CSSR: float | None = None
    if va["MVFI"] is not None and va["MVFI_NOTIONAL"] is not None and va["BEL"] is not None \
            and va["BEL_NOTIONAL"] is not None:
        PVBP_MVFI, PVBP_BEL, CSSR = credit_spread_sensitivity_ratio(va["MVFI"], va["MVFI_NOTIONAL"], va["BEL"],
                                                                    va["BEL_NOTIONAL"], NOTIONAL_VA)
    return {"NOTIONAL_VA": NOTIONAL_VA, "NOTIONAL_VA_LLFR": NOTIONAL_VA_LLFR, "CSSR_CURVE": CSSR_CURVE,
            "PVBP_MVFI": PVBP_MVFI, "PVBP_BEL": PVBP_BEL, "CSSR": CSSR}


def calculate_va_sheet(inputs: SheetInputs, res: SheetResults, va: VAInputs, CSSR: float) -> VAResults:
    """Calculate the sheet 'VA Term Structures': currency VA, macro VA, total VA and the curve with VA.

    The steps follow the worksheet formulas:
        L10 currency VA  85% * CSSR * RCS_CU, or 85% * CSSR * USA * RCS_CU with a USA (PDF 9.3.1, 9.3.5)
        L11 W_CO         country_adjustment_factor (PDF 9.4.2)
        L12 macro VA     85% * CSSR * max(RCS_CO - 1.3 * RCS_CU, 0) * W_CO if selected, else #N/A (PDF 9.4.1)
        L13 total VA     ROUND(currency VA (+ macro VA), 4): whole basis points (PDF 10.1.3)
        K20, L23:L172    va_curve with the total VA (PDF 10)

    Args:
        inputs: the inputs of the basic RFR (UFR, Convergence and LLFR weights are used).
        res: the basic RFR, from calculation.calculate_sheet.
        va: the VA inputs; RCS_CO and DTTA are needed for the macro VA only.
        CSSR: the credit spread sensitivity ratio, from calculate_cssr_sheet.

    Returns:
        VAResults.

    Raises:
        ValueError: if the macro VA is selected without the country RCS and the debt to total
            assets ratio (prevented by validation.validate_va_inputs).
    """
    FSP: int = res["FSP"]
    if va["USA"] is not None:
        CURRENCY_VA: float = VA_SHARE * CSSR * va["USA"] * va["RCS_CU"]
    else:
        CURRENCY_VA = VA_SHARE * CSSR * va["RCS_CU"]
    W_CO: float = (country_adjustment_factor(va["RCS_CO"], va["DTTA"])
                   if va["RCS_CO"] is not None and va["DTTA"] is not None else math.nan)
    MACRO_VA: float = math.nan
    TOTAL_VA: float = excel_round_value(CURRENCY_VA, 4)
    if va["MACROVA_SELECTED"]:
        if va["RCS_CO"] is None or math.isnan(W_CO):
            raise ValueError("the macro VA needs the country RCS and the debt to total assets ratio")
        MACRO_VA = VA_SHARE * CSSR * max(va["RCS_CO"] - 1.3 * va["RCS_CU"], 0) * W_CO
        TOTAL_VA = excel_round_value(CURRENCY_VA + MACRO_VA, 4)
    TOTAL_VA_LLFR, VA_CURVE = va_curve(res["BASIC_RFR"], TOTAL_VA, FSP, res["LLFR"],
                                       inputs["LLFRweightsIn"].get(FSP, 0.0), inputs["UFR"], inputs["alpha"])
    return {"CURRENCY_VA": CURRENCY_VA, "W_CO": W_CO, "MACRO_VA": MACRO_VA, "TOTAL_VA": TOTAL_VA,
            "TOTAL_VA_LLFR": TOTAL_VA_LLFR, "VA_CURVE": VA_CURVE}
