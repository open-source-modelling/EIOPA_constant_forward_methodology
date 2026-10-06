"""Unit tests for calculation.py.

Run from the project folder with either
    python -m unittest discover -s Code -v
    python -m pytest Code
"""

import math
import unittest

from calculation import (DEFAULT_MAX_MATURITY, Curve, Flags, SwapInputs, bootstrap_swaps, calculate_sheet, excel_round,
                         extrapolation, get_llfr, newton_raphson_forward_swap)

# Workbook example ('Input Data & Extrapolation', 19 May 2026): market par swap rates at the DLT maturities
EXAMPLE_RATES: Curve = {1: 0.02176, 2: 0.022621, 3: 0.023795, 4: 0.0248, 5: 0.02569, 6: 0.02651,
                        7: 0.02732, 8: 0.028, 9: 0.028631, 10: 0.02927, 11: 0.02979, 12: 0.03027,
                        13: 0.030779, 15: 0.03143, 20: 0.03233, 25: 0.0325, 30: 0.032431,
                        40: 0.03206, 50: 0.03131}
EXAMPLE_WEIGHTS: Curve = {20: 0.33, 25: 0.12, 30: 0.48, 40: 0.04, 50: 0.03}


def example_inputs() -> SwapInputs:
    """A fresh copy of the workbook example (safe to modify in a test)."""
    return SwapInputs(
        CouponFreq=1,
        CRA=10,
        UFR=0.033,
        alpha=0.11,
        dlt={t: 1 if t in EXAMPLE_RATES else 0 for t in range(1, DEFAULT_MAX_MATURITY + 1)},
        LLFRweightsIn=dict(EXAMPLE_WEIGHTS),
        SwapRatesInit=dict(EXAMPLE_RATES),
        MAX_MATURITY=DEFAULT_MAX_MATURITY,
    )


def flags(maturities: list[int]) -> Flags:
    """DLT flags: 1 at the given maturities, 0 elsewhere."""
    return {t: 1 if t in maturities else 0 for t in range(1, DEFAULT_MAX_MATURITY + 1)}


def discount_factors(zero_cc: Curve, last: int) -> Curve:
    """Discount factors exp(-t * z_t) for t = 1..last."""
    return {t: math.exp(-t * zero_cc[t]) for t in range(1, last + 1)}


def one_year_forward(zero_cc: Curve, t: int) -> float:
    """cc forward rate from t-1 to t."""
    return t * zero_cc[t] - (t - 1) * zero_cc[t - 1] if t > 1 else zero_cc[1]


class TestNewtonRaphsonForwardSwap(unittest.TestCase):

    def test_first_maturity_solution_is_the_coupon(self) -> None:
        # with target 1 the periodic forward equals the periodic coupon (PDF D.4.7)
        self.assertAlmostEqual(newton_raphson_forward_swap(0.03, 0.03, 10, 1), 0.03, places=15)

    def test_solves_the_par_equation(self) -> None:
        coupon, periods, target = 0.031, 5, 1.02
        f = newton_raphson_forward_swap(0.02, coupon, periods, target)
        d = (1 + f) ** -periods
        self.assertAlmostEqual(coupon * (1 - d) / f + d, target, places=14)

    def test_converges_quadratically(self) -> None:
        # with the exact derivative, 4 Newton steps from a guess 1% off reach the tolerance
        f = newton_raphson_forward_swap(0.02, 0.031, 5, 1.02, max_iter=4)
        d = (1 + f) ** -5
        self.assertLess(abs(0.031 * (1 - d) / f + d - 1.02), 1e-15)

    def test_converges_from_a_poor_guess(self) -> None:
        f_good = newton_raphson_forward_swap(0.03, 0.03, 20, 0.98)
        f_poor = newton_raphson_forward_swap(0.001, 0.03, 20, 0.98)
        self.assertAlmostEqual(f_good, f_poor, places=13)


class TestBootstrapSwaps(unittest.TestCase):

    def test_flat_par_curve_gives_flat_zero_curve(self) -> None:
        # for a flat par curve with annual coupons the zero rates equal the swap rate
        s = 0.03
        rates: Curve = {t: s for t in range(1, 31)}
        zero_ac = bootstrap_swaps(rates, flags(list(rates)), 1, 0, 30, "A", "Z")
        zero_cc = bootstrap_swaps(rates, flags(list(rates)), 1, 0, 30, "C", "Z")
        for t in range(1, 31):
            self.assertAlmostEqual(zero_ac[t], s, places=13)
            self.assertAlmostEqual(zero_cc[t], math.log(1 + s), places=13)

    def test_first_dlt_maturity_after_one_year(self) -> None:
        # no DLT point at 1y: the constant forward up to the first DLT maturity also gives the 1y rate
        rates: Curve = {3: 0.025, 5: 0.028, 10: 0.03}
        zero_cc = bootstrap_swaps(rates, flags(list(rates)), 1, 0, 10, "C", "Z")
        d = discount_factors(zero_cc, 10)
        for t, s in rates.items():
            self.assertAlmostEqual(s * sum(d[k] for k in range(1, t + 1)) + d[t], 1.0, places=12, msg=f"maturity {t}")
        for t in (1, 2):
            self.assertAlmostEqual(zero_cc[t], math.log(1.025), places=13)

    def test_flat_semi_annual_par_curve(self) -> None:
        # semi-annual coupons: the annual zero rate is (1 + s/2)^2 - 1
        s = 0.03
        rates: Curve = {t: s for t in (1, 2, 5, 10)}
        zero_ac = bootstrap_swaps(rates, flags(list(rates)), 2, 0, 10, "A", "Z")
        for t in range(1, 11):
            self.assertAlmostEqual(zero_ac[t], (1 + s / 2) ** 2 - 1, places=12)

    def test_cra_is_deducted_in_basis_points(self) -> None:
        with_cra = bootstrap_swaps({t: 0.031 for t in (1, 5, 10)}, flags([1, 5, 10]), 1, 10, 10, "C", "Z")
        without = bootstrap_swaps({t: 0.030 for t in (1, 5, 10)}, flags([1, 5, 10]), 1, 0, 10, "C", "Z")
        for t in range(1, 11):
            self.assertAlmostEqual(with_cra[t], without[t], places=14)

    def test_dlt_swaps_are_priced_at_par(self) -> None:
        # PDF D.3.5: (s - CRA) * sum of discount factors + d_T = 1 for every DLT maturity
        CRA = 10
        zero_cc = bootstrap_swaps(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), 1, CRA, 50, "C", "Z")
        d = discount_factors(zero_cc, 50)
        for t, s in EXAMPLE_RATES.items():
            value = (s - CRA / 10000) * sum(d[k] for k in range(1, t + 1)) + d[t]
            self.assertAlmostEqual(value, 1.0, places=12, msg=f"maturity {t}")

    def test_forward_is_constant_between_dlt_points(self) -> None:
        zero_cc = bootstrap_swaps(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), 1, 10, 50, "C", "Z")
        for first, last in ((15, 20), (20, 25), (30, 40), (40, 50)):
            forwards = [one_year_forward(zero_cc, t) for t in range(first + 1, last + 1)]
            for f in forwards:
                self.assertAlmostEqual(f, forwards[0], places=12, msg=f"interval {first}-{last}")

    def test_rates_at_non_dlt_maturities_are_ignored(self) -> None:
        dlt = flags(list(EXAMPLE_RATES))
        with_extra: Curve = dict(EXAMPLE_RATES)
        with_extra.update({14: 0.5, 16: 0.5, 35: 0.5})
        self.assertEqual(bootstrap_swaps(with_extra, dlt, 1, 10, 50, "C", "Z"),
                         bootstrap_swaps(EXAMPLE_RATES, dlt, 1, 10, 50, "C", "Z"))

    def test_beyond_max_tenor_is_nan(self) -> None:
        zero_cc = bootstrap_swaps(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), 1, 10, 50, "C", "Z")
        self.assertEqual(len(zero_cc), DEFAULT_MAX_MATURITY)
        self.assertFalse(math.isnan(zero_cc[50]))
        self.assertTrue(all(math.isnan(zero_cc[t]) for t in range(51, DEFAULT_MAX_MATURITY + 1)))

    def test_no_nan_when_last_dlt_maturity_is_max_maturity(self) -> None:
        zero_cc = bootstrap_swaps(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), 1, 10, 50, "C", "Z", MAX_MATURITY=50)
        self.assertEqual(sorted(zero_cc), list(range(1, 51)))
        self.assertFalse(any(math.isnan(z) for z in zero_cc.values()))

    def test_result_length_follows_max_maturity(self) -> None:
        zero_cc = bootstrap_swaps(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), 1, 10, 50, "C", "Z", MAX_MATURITY=70)
        self.assertEqual(sorted(zero_cc), list(range(1, 71)))
        self.assertTrue(all(math.isnan(zero_cc[t]) for t in range(51, 71)))

    def test_flat_forward_extension_up_to_max_tenor(self) -> None:
        # MaxTenor beyond the last DLT maturity: the last forward is carried forward flat
        zero_cc = bootstrap_swaps(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), 1, 10, 60, "C", "Z")
        last_forward = one_year_forward(zero_cc, 50)
        for t in range(51, 61):
            self.assertAlmostEqual(one_year_forward(zero_cc, t), last_forward, places=12)
        self.assertTrue(math.isnan(zero_cc[61]))


class TestGetLLFR(unittest.TestCase):

    def setUp(self) -> None:
        # simple cc curve: z_t = 0.02 + 0.0005 * t
        self.zero_cc: Curve = {t: 0.02 + 0.0005 * t for t in range(1, DEFAULT_MAX_MATURITY + 1)}
        self.dlt: Flags = flags([1, 5, 10, 15, 20, 30])

    def forward(self, a: int, b: int) -> float:
        return (b * self.zero_cc[b] - a * self.zero_cc[a]) / (b - a)

    def test_single_weight_is_the_forward_into_the_fsp(self) -> None:
        # PDF 8.3.4: FSP is the last DLT point -> LLFR = forward from the previous DLT point to the FSP
        self.assertAlmostEqual(get_llfr(self.zero_cc, self.dlt, {20: 1.0}), self.forward(15, 20), places=15)

    def test_weighted_average_of_forwards(self) -> None:
        # PDF 8.5.6: w_F * f(previous DLT, FSP) + sum of w_k * f(FSP, t_k)
        weights: Curve = {15: 0.5, 20: 0.3, 30: 0.2}
        expected = 0.5 * self.forward(10, 15) + 0.3 * self.forward(15, 20) + 0.2 * self.forward(15, 30)
        self.assertAlmostEqual(get_llfr(self.zero_cc, self.dlt, weights), expected, places=15)

    def test_uses_the_last_dlt_point_before_the_fsp(self) -> None:
        # maturities 16..19 are not DLT, so the first forward runs from 15 to 20
        self.assertAlmostEqual(get_llfr(self.zero_cc, self.dlt, {20: 0.5, 30: 0.5}),
                               0.5 * self.forward(15, 20) + 0.5 * self.forward(20, 30), places=15)

    def test_weights_not_summing_to_one_raise(self) -> None:
        with self.assertRaisesRegex(ValueError, "LLFR weights must sum to 1"):
            get_llfr(self.zero_cc, self.dlt, {20: 0.5, 30: 0.4})

    def test_no_weights_raise(self) -> None:
        # the weight check, not an error from an empty loop range
        with self.assertRaisesRegex(ValueError, "LLFR weights must sum to 1"):
            get_llfr(self.zero_cc, self.dlt, {})


class TestExtrapolation(unittest.TestCase):

    def setUp(self) -> None:
        self.FSP, self.UFR, self.alpha = 20, 0.033, 0.11
        self.zero_cc: Curve = {t: 0.02 + 0.0005 * t for t in range(1, DEFAULT_MAX_MATURITY + 1)}
        self.LLFR = 0.031

    def test_rates_up_to_fsp_are_kept(self) -> None:
        curve = extrapolation(self.zero_cc, "Z", self.FSP, self.UFR, self.LLFR, self.alpha, "C")
        for t in range(1, self.FSP + 1):
            self.assertEqual(curve[t], self.zero_cc[t])

    def test_rates_after_fsp_are_not_used(self) -> None:
        changed: Curve = dict(self.zero_cc)
        changed.update({t: 0.5 for t in range(self.FSP + 1, DEFAULT_MAX_MATURITY + 1)})
        self.assertEqual(extrapolation(changed, "Z", self.FSP, self.UFR, self.LLFR, self.alpha, "C"),
                         extrapolation(self.zero_cc, "Z", self.FSP, self.UFR, self.LLFR, self.alpha, "C"))

    def test_forward_from_fsp_follows_the_formula(self) -> None:
        # PDF 8.5.5: f(FSP, FSP+h) = ln(1+UFR) + (LLFR - ln(1+UFR)) * (1 - exp(-alpha h)) / (alpha h)
        curve = extrapolation(self.zero_cc, "Z", self.FSP, self.UFR, self.LLFR, self.alpha, "C")
        ufr_cc = math.log(1 + self.UFR)
        for h in (1, 10, 50, 130):
            t = self.FSP + h
            forward = (t * curve[t] - self.FSP * curve[self.FSP]) / h
            b = (1 - math.exp(-self.alpha * h)) / (self.alpha * h)
            self.assertAlmostEqual(forward, ufr_cc + (self.LLFR - ufr_cc) * b, places=14)

    def test_forward_converges_to_ufr(self) -> None:
        curve = extrapolation(self.zero_cc, "Z", self.FSP, self.UFR, self.LLFR, self.alpha, "C")
        ufr_cc = math.log(1 + self.UFR)
        gaps = [abs((t * curve[t] - self.FSP * curve[self.FSP]) / (t - self.FSP) - ufr_cc)
                for t in (21, 40, 80, 150)]
        self.assertEqual(gaps, sorted(gaps, reverse=True))
        self.assertLess(gaps[-1], 0.1 * gaps[0])

    def test_higher_alpha_converges_faster(self) -> None:
        slow = extrapolation(self.zero_cc, "Z", self.FSP, self.UFR, self.LLFR, 0.11, "C")
        fast = extrapolation(self.zero_cc, "Z", self.FSP, self.UFR, self.LLFR, 0.40, "C")
        ufr_cc = math.log(1 + self.UFR)
        self.assertLess(abs(fast[60] - ufr_cc), abs(slow[60] - ufr_cc))

    def test_flat_curve_at_ufr_stays_flat(self) -> None:
        ufr_cc = math.log(1 + self.UFR)
        curve = extrapolation({t: ufr_cc for t in range(1, DEFAULT_MAX_MATURITY + 1)}, "Z", self.FSP, self.UFR, ufr_cc, self.alpha, "A")
        for t in range(1, DEFAULT_MAX_MATURITY + 1):
            self.assertAlmostEqual(curve[t], self.UFR, places=14)

    def test_result_length_follows_max_maturity(self) -> None:
        short = extrapolation(self.zero_cc, "Z", self.FSP, self.UFR, self.LLFR, self.alpha, "C", MAX_MATURITY=60)
        full = extrapolation(self.zero_cc, "Z", self.FSP, self.UFR, self.LLFR, self.alpha, "C")
        self.assertEqual(sorted(short), list(range(1, 61)))
        self.assertEqual(short, {t: full[t] for t in range(1, 61)})

    def test_annual_compounding_output(self) -> None:
        cc = extrapolation(self.zero_cc, "Z", self.FSP, self.UFR, self.LLFR, self.alpha, "C")
        ac = extrapolation(self.zero_cc, "Z", self.FSP, self.UFR, self.LLFR, self.alpha, "A")
        for t in range(1, DEFAULT_MAX_MATURITY + 1):
            self.assertAlmostEqual(ac[t], math.exp(cc[t]) - 1, places=15)

    def test_forward_input_is_averaged_into_zero_rates(self) -> None:
        # RateType "F": zero rates up to the FSP are the running average of the one-year forwards
        forwards: Curve = {t: 0.01 * (1 + (t % 3)) for t in range(1, DEFAULT_MAX_MATURITY + 1)}
        curve = extrapolation(forwards, "F", self.FSP, self.UFR, self.LLFR, self.alpha, "C")
        for t in (1, 2, 7, self.FSP):
            self.assertAlmostEqual(curve[t], sum(forwards[k] for k in range(1, t + 1)) / t, places=15)


class TestExcelRound(unittest.TestCase):

    def test_rounds_every_value_to_five_decimals(self) -> None:
        self.assertEqual(excel_round({1: 0.0207612, 2: 0.0326049}), {1: 0.02076, 2: 0.0326})

    def test_keeps_nan(self) -> None:
        self.assertTrue(math.isnan(excel_round({1: math.nan})[1]))


class TestCalculateSheet(unittest.TestCase):
    """The workbook example (values as calculated by Excel/VBA), and shorter and longer curves."""

    def setUp(self) -> None:
        self.res = calculate_sheet(**example_inputs())

    def test_fsp_and_llfr(self) -> None:
        self.assertEqual(self.res["FSP"], 20)
        self.assertAlmostEqual(self.res["LLFR"], 0.03224887127830391, places=15)

    def test_bootstrapped_curve(self) -> None:
        self.assertAlmostEqual(self.res["BOOTSTRAPPED_CURVE"][1], 0.02054744788766011, places=15)
        self.assertAlmostEqual(self.res["BOOTSTRAPPED_CURVE"][20], 0.031586531710533036, places=15)
        self.assertAlmostEqual(self.res["BOOTSTRAPPED_CURVE"][50], 0.02925667435304452, places=15)
        self.assertTrue(math.isnan(self.res["BOOTSTRAPPED_CURVE"][51]))

    def test_basic_rfr(self) -> None:
        expected: Curve = {1: 0.02076, 2: 0.02163, 10: 0.02863, 20: 0.03209, 21: 0.03212,
                           40: 0.0325, 60: 0.03266, 100: 0.0328, 150: 0.03287}
        for t, rate in expected.items():
            self.assertEqual(self.res["BASIC_RFR"][t], rate, msg=f"maturity {t}")
        self.assertEqual(len(self.res["BASIC_RFR"]), DEFAULT_MAX_MATURITY)

    def test_shorter_max_maturity_cuts_the_curves(self) -> None:
        inputs = example_inputs()
        inputs["MAX_MATURITY"] = 100
        res = calculate_sheet(**inputs)
        self.assertEqual(sorted(res["BASIC_RFR"]), list(range(1, 101)))
        self.assertEqual(sorted(res["BOOTSTRAPPED_CURVE"]), list(range(1, 101)))
        self.assertEqual(res["BASIC_RFR"], {t: self.res["BASIC_RFR"][t] for t in range(1, 101)})
        self.assertEqual((res["FSP"], res["LLFR"]), (self.res["FSP"], self.res["LLFR"]))

    def test_longer_max_maturity_extends_the_curves(self) -> None:
        inputs = example_inputs()
        inputs["MAX_MATURITY"] = 200
        res = calculate_sheet(**inputs)
        self.assertEqual(sorted(res["BASIC_RFR"]), list(range(1, 201)))
        self.assertEqual({t: res["BASIC_RFR"][t] for t in range(1, DEFAULT_MAX_MATURITY + 1)}, self.res["BASIC_RFR"])
        # beyond the default length the curve keeps converging towards the UFR (0.033)
        self.assertLess(abs(res["BASIC_RFR"][200] - 0.033), abs(res["BASIC_RFR"][DEFAULT_MAX_MATURITY] - 0.033))

    def test_rates_at_non_dlt_maturities_do_not_change_the_result(self) -> None:
        inputs = example_inputs()
        inputs["SwapRatesInit"].update({14: 0.031124, 16: 0.031742, 35: 0.0325})
        self.assertEqual(calculate_sheet(**inputs)["BASIC_RFR"], self.res["BASIC_RFR"])


if __name__ == "__main__":
    unittest.main()
