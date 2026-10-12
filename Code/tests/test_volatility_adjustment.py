"""Unit tests for volatility_adjustment.py.

Run from the project folder with either
    python -m unittest discover -s Code -v
    python -m pytest Code
"""

import math
import unittest

from calculation import Curve, SheetInputs, SheetResults, calculate_sheet
from tests.test_calculation import example_inputs
from volatility_adjustment import (VAInputs, calculate_cssr_sheet, calculate_va_sheet, country_adjustment_factor,
                                   credit_spread_sensitivity_ratio, notional_va, va_curve)


def example_va_inputs() -> VAInputs:
    """A fresh copy of the workbook's VA example (safe to modify in a test).

    Sheet 'Input Data & Extrapolation' L10:L12, 'CSSR Calculation' L10:L14 and
    'VA Term Structures' H10:H12 (macro VA off, USA of 80%).
    """
    return VAInputs(RCS_CU=0.01, RCS_CO=0.014, DTTA=0.5, MVFI=500000.0, MVFI_NOTIONAL=490000.0, BEL=870000.0,
                    BEL_NOTIONAL=850000.0, MACROVA_SELECTED=False, USA=0.8)


def forward(curve_ac: Curve, t: int) -> float:
    """Continuously compounded one-year forward rate from t-1 to t of an annually compounded curve."""
    return t * math.log(1 + curve_ac[t]) - (t - 1) * math.log(1 + curve_ac[t - 1]) if t > 1 else math.log(1 + curve_ac[1])


class TestNotionalVA(unittest.TestCase):

    def test_85_percent_of_the_rcs(self) -> None:
        self.assertEqual(notional_va(0.01), 0.0085)

    def test_rounded_to_whole_basis_points(self) -> None:
        # 85% of 37 bp is 31.45 bp: rounded to 31 bp; a half rounds away from zero, as in Excel
        self.assertEqual(notional_va(0.0037), 0.0031)
        self.assertEqual(notional_va(0.003), 0.0026)        # 25.5 bp
        self.assertEqual(notional_va(-0.003), -0.0026)


class TestCreditSpreadSensitivityRatio(unittest.TestCase):

    def test_ratio_of_the_changes_in_value(self) -> None:
        # assets lose 10,000 and the best estimate 20,000 on the CSSR curve: CSSR = 0.5
        PVBP_MVFI, PVBP_BEL, CSSR = credit_spread_sensitivity_ratio(500000, 490000, 870000, 850000, 0.0085)
        self.assertAlmostEqual(PVBP_MVFI, 10000 / 0.0085, places=6)
        self.assertAlmostEqual(PVBP_BEL, 20000 / 0.0085, places=6)
        self.assertEqual(CSSR, 0.5)

    def test_limited_to_one(self) -> None:
        self.assertEqual(credit_spread_sensitivity_ratio(500000, 450000, 870000, 850000, 0.0085)[2], 1)

    def test_limited_to_zero(self) -> None:
        # assets gain on the higher curve: a negative ratio becomes 0
        self.assertEqual(credit_spread_sensitivity_ratio(500000, 510000, 870000, 850000, 0.0085)[2], 0)

    def test_best_estimate_not_falling_gives_one(self) -> None:
        # the workbook's IF(PVBP_BEL <= 0, 1, ...), also for a change of exactly 0
        self.assertEqual(credit_spread_sensitivity_ratio(500000, 490000, 870000, 870000, 0.0085)[2], 1)
        self.assertEqual(credit_spread_sensitivity_ratio(500000, 490000, 870000, 880000, 0.0085)[2], 1)

    def test_negative_notional_va_gives_the_same_ratio(self) -> None:
        # with a negative RCS the CSSR curve lies below the basic RFR and both values rise
        self.assertEqual(credit_spread_sensitivity_ratio(500000, 510000, 870000, 890000, -0.0085)[2], 0.5)


class TestCountryAdjustmentFactor(unittest.TestCase):

    def test_workbook_example(self) -> None:
        # (1.4% * 50% - 0.6%) / 0.3% = 1/3
        self.assertAlmostEqual(country_adjustment_factor(0.014, 0.5), 1 / 3, places=15)

    def test_limited_to_zero_and_one(self) -> None:
        self.assertEqual(country_adjustment_factor(0.01, 0.5), 0)
        self.assertEqual(country_adjustment_factor(0.03, 0.5), 1)


class TestVACurve(unittest.TestCase):

    def setUp(self) -> None:
        self.inputs: SheetInputs = example_inputs()
        self.res: SheetResults = calculate_sheet(**self.inputs)
        self.FSP: int = self.res["FSP"]
        self.w: float = self.inputs["LLFRweightsIn"][self.FSP]

    def curve(self, VA: float) -> tuple[float, Curve]:
        return va_curve(self.res["BASIC_RFR"], VA, self.FSP, self.res["LLFR"], self.w, self.inputs["UFR"],
                        self.inputs["alpha"])

    def test_llfr_rises_by_the_first_weight_times_the_va(self) -> None:
        # PDF 10.2.1: LLFR with VA = LLFR + w_FSP * ln(1 + VA)
        LLFR_VA, _ = self.curve(0.0034)
        self.assertAlmostEqual(LLFR_VA, self.res["LLFR"] + self.w * math.log(1.0034), places=15)

    def test_forwards_up_to_the_fsp_rise_by_the_va(self) -> None:
        # PDF 10.1.6: f with VA = f + ln(1 + VA); 5e-5 allows for the rounding of both curves
        _, curve = self.curve(0.0034)
        for t in range(1, self.FSP + 1):
            self.assertAlmostEqual(forward(curve, t) - forward(self.res["BASIC_RFR"], t), math.log(1.0034),
                                   delta=5e-5 * t, msg=f"maturity {t}")

    def test_zero_va_gives_the_basic_rfr_up_to_the_fsp(self) -> None:
        _, curve = self.curve(0.0)
        for t in range(1, self.FSP + 1):
            self.assertEqual(curve[t], self.res["BASIC_RFR"][t])

    def test_extrapolated_forward_follows_the_formula(self) -> None:
        # PDF 10.2.2: f(FSP, FSP+h) = UFR^c + (LLFR^VA - UFR^c) * B(alpha, h), up to the rounding to 5 decimals
        LLFR_VA, curve = self.curve(0.0034)
        ufr_cc = math.log(1 + self.inputs["UFR"])
        alpha = self.inputs["alpha"]
        for h in (1, 10, 50, 130):
            t = self.FSP + h
            f = (t * math.log(1 + curve[t]) - self.FSP * math.log(1 + curve[self.FSP])) / h
            b = (1 - math.exp(-alpha * h)) / (alpha * h)
            self.assertAlmostEqual(f, ufr_cc + (LLFR_VA - ufr_cc) * b, delta=1e-5 * t / h)

    def test_va_effect_fades_towards_the_ufr(self) -> None:
        _, curve = self.curve(0.0034)
        effect = {t: curve[t] - self.res["BASIC_RFR"][t] for t in (self.FSP, 60, 100, 150)}
        self.assertEqual(list(effect.values()), sorted(effect.values(), reverse=True))
        self.assertGreater(effect[150], 0)

    def test_negative_va_lowers_the_curve(self) -> None:
        _, curve = self.curve(-0.0010)
        for t in (1, self.FSP, 150):
            self.assertLess(curve[t], self.res["BASIC_RFR"][t])

    def test_length_follows_the_basic_rfr(self) -> None:
        inputs = example_inputs()
        inputs["MAX_MATURITY"] = 80
        res = calculate_sheet(**inputs)
        _, curve = va_curve(res["BASIC_RFR"], 0.0034, res["FSP"], res["LLFR"], self.w, inputs["UFR"], inputs["alpha"])
        self.assertEqual(sorted(curve), list(range(1, 81)))


class TestCalculateCSSRSheet(unittest.TestCase):
    """The workbook example (values as calculated by Excel, sheet 'CSSR Calculation')."""

    def setUp(self) -> None:
        self.inputs: SheetInputs = example_inputs()
        self.res: SheetResults = calculate_sheet(**self.inputs)
        self.cssr = calculate_cssr_sheet(self.inputs, self.res, example_va_inputs())

    def test_notional_va_and_llfr(self) -> None:
        self.assertEqual(self.cssr["NOTIONAL_VA"], 0.0085)                           # H10
        self.assertEqual(self.cssr["NOTIONAL_VA_LLFR"], 0.035042017154306604)         # K20

    def test_cssr(self) -> None:
        self.assertEqual(self.cssr["PVBP_MVFI"], 1176470.588235294)                   # L12
        self.assertEqual(self.cssr["PVBP_BEL"], 2352941.176470588)                    # L15
        self.assertEqual(self.cssr["CSSR"], 0.5)                                      # L16

    def test_cssr_curve(self) -> None:
        expected: Curve = {1: 0.02944, 2: 0.03031, 20: 0.04086, 21: 0.04061, 150: 0.03421}   # L24:L173
        for t, rate in expected.items():
            self.assertEqual(self.cssr["CSSR_CURVE"][t], rate, msg=f"maturity {t}")
        self.assertEqual(len(self.cssr["CSSR_CURVE"]), 150)

    def test_without_the_values_only_the_curve_is_calculated(self) -> None:
        # any one of the four values missing is enough (validation rejects a partial set before)
        for missing in ("MVFI", "MVFI_NOTIONAL", "BEL", "BEL_NOTIONAL"):
            va = example_va_inputs()
            va[missing] = None                                       # type: ignore[literal-required]
            cssr = calculate_cssr_sheet(self.inputs, self.res, va)
            self.assertEqual((cssr["PVBP_MVFI"], cssr["PVBP_BEL"], cssr["CSSR"]), (None, None, None), msg=missing)
            self.assertEqual(cssr["CSSR_CURVE"], self.cssr["CSSR_CURVE"])

    def test_government_bonds(self) -> None:
        inputs = example_inputs("GVT")
        res = calculate_sheet(**inputs)
        cssr = calculate_cssr_sheet(inputs, res, example_va_inputs())
        _, expected = va_curve(res["BASIC_RFR"], 0.0085, res["FSP"], res["LLFR"], inputs["LLFRweightsIn"][res["FSP"]],
                               inputs["UFR"], inputs["alpha"])
        self.assertEqual(cssr["CSSR_CURVE"], expected)


class TestCalculateVASheet(unittest.TestCase):
    """The workbook example (values as calculated by Excel, sheet 'VA Term Structures'), and the macro VA."""

    def setUp(self) -> None:
        self.inputs: SheetInputs = example_inputs()
        self.res: SheetResults = calculate_sheet(**self.inputs)

    def va_sheet(self, va: VAInputs, CSSR: float = 0.5):
        return calculate_va_sheet(self.inputs, self.res, va, CSSR)

    def test_currency_va_with_usa(self) -> None:
        # 85% * 0.5 * 80% * 1%
        res = self.va_sheet(example_va_inputs())
        self.assertEqual(res["CURRENCY_VA"], 0.0034000000000000002)                 # L10
        self.assertAlmostEqual(res["W_CO"], 1 / 3, places=15)                        # L11
        self.assertTrue(math.isnan(res["MACRO_VA"]))                                 # L12 #N/A
        self.assertEqual(res["TOTAL_VA"], 0.0034)                                    # L13
        self.assertEqual(res["TOTAL_VA_LLFR"], 0.03336896819074907)                  # K20

    def test_va_curve(self) -> None:
        expected: Curve = {1: 0.02423, 2: 0.0251, 20: 0.0356, 21: 0.03552, 150: 0.0334}    # L23:L172
        res = self.va_sheet(example_va_inputs())
        for t, rate in expected.items():
            self.assertEqual(res["VA_CURVE"][t], rate, msg=f"maturity {t}")

    def test_currency_va_without_usa(self) -> None:
        va = example_va_inputs()
        va["USA"] = None
        res = self.va_sheet(va)
        self.assertAlmostEqual(res["CURRENCY_VA"], 0.85 * 0.5 * 0.01, places=15)
        self.assertEqual(res["TOTAL_VA"], 0.0043)                                    # 42.5 bp rounds up

    def test_macro_va(self) -> None:
        # 85% * 0.5 * max(1.4% - 1.3 * 1%, 0) * 1/3 = 1.42 bp; total 42.5 + 1.42 = 43.92 bp, rounded to 44 bp
        va = example_va_inputs()
        va["USA"] = None
        va["MACROVA_SELECTED"] = True
        res = self.va_sheet(va)
        self.assertAlmostEqual(res["MACRO_VA"], 0.85 * 0.5 * 0.001 / 3, places=15)
        self.assertEqual(res["TOTAL_VA"], 0.0044)

    def test_macro_va_is_zero_below_130_percent_of_the_currency_rcs(self) -> None:
        va = example_va_inputs()
        va.update(USA=None, MACROVA_SELECTED=True, RCS_CO=0.012, DTTA=1.0)
        self.assertEqual(self.va_sheet(va)["MACRO_VA"], 0)

    def test_macro_va_needs_the_country_values(self) -> None:
        for missing in ("RCS_CO", "DTTA"):
            va = example_va_inputs()
            va.update({"USA": None, "MACROVA_SELECTED": True, missing: None})
            with self.assertRaisesRegex(ValueError, "macro VA needs", msg=missing):
                self.va_sheet(va)

    def test_country_adjustment_factor_without_the_country_values(self) -> None:
        # without macro VA the country values are optional; the factor is then #N/A
        for missing in ("RCS_CO", "DTTA"):
            va = example_va_inputs()
            va[missing] = None                                       # type: ignore[literal-required]
            res = self.va_sheet(va)
            self.assertTrue(math.isnan(res["W_CO"]), msg=missing)
            self.assertEqual(res["TOTAL_VA"], 0.0034)

    def test_cssr_of_zero_gives_the_basic_rfr(self) -> None:
        res = self.va_sheet(example_va_inputs(), CSSR=0.0)
        self.assertEqual(res["TOTAL_VA"], 0)
        for t in range(1, self.res["FSP"] + 1):
            self.assertEqual(res["VA_CURVE"][t], self.res["BASIC_RFR"][t])

    def test_negative_rcs_gives_a_negative_va(self) -> None:
        # PDF 9.3.6: the currency VA can be negative
        va = example_va_inputs()
        va["RCS_CU"] = -0.002
        res = self.va_sheet(va)
        self.assertLess(res["TOTAL_VA"], 0)
        self.assertLess(res["VA_CURVE"][1], self.res["BASIC_RFR"][1])


if __name__ == "__main__":
    unittest.main()
