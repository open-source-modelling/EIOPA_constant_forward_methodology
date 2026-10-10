"""Unit tests for calculation.py.

Run from the project folder with either
    python -m unittest discover -s Code -v
    python -m pytest Code
"""

import math
import unittest

from calculation import (DEFAULT_MAX_MATURITY, Curve, Flags, SheetInputs, bootstrap_curve, bootstrap_swaps,
                         bootstrap_zeros, calculate_sheet, excel_round, excel_round_value, extrapolation, get_llfr,
                         newton_raphson_forward_swap)

# Workbook example ('Input Data & Extrapolation', 19 May 2026): market rates at the DLT maturities, read as
# par swap rates (instrument SWP) or as annually compounded government bond zero rates (instrument GVT)
EXAMPLE_RATES: Curve = {1: 0.02176, 2: 0.022621, 3: 0.023795, 4: 0.0248, 5: 0.02569, 6: 0.02651,
                        7: 0.02732, 8: 0.028, 9: 0.028631, 10: 0.02927, 11: 0.02979, 12: 0.03027,
                        13: 0.030779, 15: 0.03143, 20: 0.03233, 25: 0.0325, 30: 0.032431,
                        40: 0.03206, 50: 0.03131}
EXAMPLE_WEIGHTS: Curve = {20: 0.33, 25: 0.12, 30: 0.48, 40: 0.04, 50: 0.03}


def example_inputs(Instrument: str = "SWP") -> SheetInputs:
    """A fresh copy of the workbook example for swaps or government bonds (safe to modify in a test)."""
    return SheetInputs(
        Instrument=Instrument,
        CouponFreq=1 if Instrument == "SWP" else None,
        CRA=10,
        UFR=0.033,
        alpha=0.11,
        dlt={t: 1 if t in EXAMPLE_RATES else 0 for t in range(1, DEFAULT_MAX_MATURITY + 1)},
        LLFRweightsIn=dict(EXAMPLE_WEIGHTS),
        Rate=dict(EXAMPLE_RATES),
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


class TestBootstrapZeros(unittest.TestCase):

    def test_flat_zero_curve_stays_flat(self) -> None:
        s = 0.03
        rates: Curve = {t: s for t in range(1, 31)}
        zero_cc = bootstrap_zeros(rates, flags(list(rates)), "A", 0, 30, "C", "Z")
        zero_ac = bootstrap_zeros(rates, flags(list(rates)), "A", 0, 30, "A", "Z")
        for t in range(1, 31):
            self.assertAlmostEqual(zero_cc[t], math.log(1 + s), places=15)
            self.assertAlmostEqual(zero_ac[t], s, places=15)

    def test_reproduces_the_input_rates_at_dlt_maturities(self) -> None:
        # z_t = ln(1 + r_t - CRA) at every DLT maturity
        zero_cc = bootstrap_zeros(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, 50, "C", "Z")
        for t, r in EXAMPLE_RATES.items():
            self.assertAlmostEqual(zero_cc[t], math.log(1 + r - 0.001), places=15, msg=f"maturity {t}")

    def test_cra_is_deducted_in_basis_points(self) -> None:
        lowered: Curve = {t: r - 0.001 for t, r in EXAMPLE_RATES.items()}
        self.assertEqual(bootstrap_zeros(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, 50, "C", "Z"),
                         bootstrap_zeros(lowered, flags(list(EXAMPLE_RATES)), "A", 0, 50, "C", "Z"))

    def test_continuously_compounded_input_is_used_as_given(self) -> None:
        rates: Curve = {1: 0.02, 5: 0.025, 10: 0.03}
        zero_cc = bootstrap_zeros(rates, flags(list(rates)), "C", 0, 10, "C", "Z")
        for t, r in rates.items():
            self.assertAlmostEqual(zero_cc[t], r, places=15)

    def test_first_rate_is_held_flat_before_the_first_dlt_maturity(self) -> None:
        rates: Curve = {3: 0.025, 5: 0.028, 10: 0.03}
        zero_cc = bootstrap_zeros(rates, flags(list(rates)), "A", 0, 10, "C", "Z")
        for t in (1, 2, 3):
            self.assertEqual(zero_cc[t], math.log(1.025))

    def test_forward_is_constant_between_dlt_points(self) -> None:
        # f = (b * z_b - a * z_a) / (b - a) for every year between DLT maturities a and b
        zero_cc = bootstrap_zeros(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, 50, "C", "Z")
        for a, b in ((13, 15), (15, 20), (20, 25), (30, 40), (40, 50)):
            expected = (b * zero_cc[b] - a * zero_cc[a]) / (b - a)
            for t in range(a + 1, b + 1):
                self.assertAlmostEqual(one_year_forward(zero_cc, t), expected, places=12, msg=f"interval {a}-{b}")

    def test_rates_at_non_dlt_maturities_are_ignored(self) -> None:
        dlt = flags(list(EXAMPLE_RATES))
        with_extra: Curve = dict(EXAMPLE_RATES)
        with_extra.update({14: 0.5, 16: 0.5, 35: 0.5})
        self.assertEqual(bootstrap_zeros(with_extra, dlt, "A", 10, 50, "C", "Z"),
                         bootstrap_zeros(EXAMPLE_RATES, dlt, "A", 10, 50, "C", "Z"))

    def test_dlt_maturity_without_rate_is_skipped(self) -> None:
        without_15: Curve = {t: r for t, r in EXAMPLE_RATES.items() if t != 15}
        self.assertEqual(bootstrap_zeros(without_15, flags(list(EXAMPLE_RATES)), "A", 10, 50, "C", "Z"),
                         bootstrap_zeros(EXAMPLE_RATES, flags(list(without_15)), "A", 10, 50, "C", "Z"))

    def test_beyond_max_tenor_is_nan(self) -> None:
        zero_cc = bootstrap_zeros(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, 50, "C", "Z")
        self.assertEqual(sorted(zero_cc), list(range(1, DEFAULT_MAX_MATURITY + 1)))
        self.assertFalse(any(math.isnan(zero_cc[t]) for t in range(1, 51)))
        self.assertTrue(all(math.isnan(zero_cc[t]) for t in range(51, DEFAULT_MAX_MATURITY + 1)))

    def test_no_nan_when_last_dlt_maturity_is_max_maturity(self) -> None:
        zero_cc = bootstrap_zeros(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, 50, "C", "Z", MAX_MATURITY=50)
        self.assertEqual(sorted(zero_cc), list(range(1, 51)))
        self.assertFalse(any(math.isnan(z) for z in zero_cc.values()))

    def test_result_length_follows_max_maturity(self) -> None:
        zero_cc = bootstrap_zeros(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, 50, "C", "Z", MAX_MATURITY=70)
        self.assertEqual(sorted(zero_cc), list(range(1, 71)))
        self.assertTrue(all(math.isnan(zero_cc[t]) for t in range(51, 71)))

    def test_flat_forward_extension_up_to_max_tenor(self) -> None:
        # MaxTenor beyond the last DLT maturity: the last forward is carried forward flat
        zero_cc = bootstrap_zeros(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, 60, "C", "Z")
        last_forward = one_year_forward(zero_cc, 50)
        for t in range(51, 61):
            self.assertAlmostEqual(one_year_forward(zero_cc, t), last_forward, places=12)
        self.assertTrue(math.isnan(zero_cc[61]))

    def test_annual_output_is_the_continuous_output_converted(self) -> None:
        zero_cc = bootstrap_zeros(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, 50, "C", "Z")
        zero_ac = bootstrap_zeros(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, 50, "A", "Z")
        for t in range(1, 51):
            self.assertAlmostEqual(zero_ac[t], math.exp(zero_cc[t]) - 1, places=15)
        self.assertTrue(math.isnan(zero_ac[51]))


class TestBootstrapCurve(unittest.TestCase):

    def test_swaps_use_bootstrap_swaps(self) -> None:
        self.assertEqual(bootstrap_curve("SWP", EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, 1, 50, "C", "Z"),
                         bootstrap_swaps(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), 1, 10, 50, "C", "Z"))

    def test_government_bonds_use_bootstrap_zeros(self) -> None:
        self.assertEqual(bootstrap_curve("GVT", EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, None, 50, "C", "Z"),
                         bootstrap_zeros(EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, 50, "C", "Z"))

    def test_coupon_frequency_is_not_used_for_government_bonds(self) -> None:
        self.assertEqual(bootstrap_curve("GVT", EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, 4, 50, "C", "Z"),
                         bootstrap_curve("GVT", EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, None, 50, "C", "Z"))

    def test_swaps_without_coupon_frequency_raise(self) -> None:
        with self.assertRaisesRegex(ValueError, "swaps need a coupon frequency"):
            bootstrap_curve("SWP", EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, None, 50, "C", "Z")

    def test_max_maturity_is_passed_on(self) -> None:
        for Instrument, CouponFreq in (("SWP", 1), ("GVT", None)):
            zero_cc = bootstrap_curve(Instrument, EXAMPLE_RATES, flags(list(EXAMPLE_RATES)), "A", 10, CouponFreq, 50,
                                      "C", "Z", MAX_MATURITY=60)
            self.assertEqual(sorted(zero_cc), list(range(1, 61)), msg=Instrument)


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
    """Expected values are the results of Excel's ROUND for the same numbers."""

    def test_rounds_every_value_to_five_decimals(self) -> None:
        self.assertEqual(excel_round({1: 0.0207612, 2: 0.0326049}), {1: 0.02076, 2: 0.0326})

    def test_keeps_nan(self) -> None:
        self.assertTrue(math.isnan(excel_round({1: math.nan})[1]))

    def test_near_ties_round_like_excel(self) -> None:
        # ties at Excel's 15 significant digits, slightly below the tie in binary: Excel rounds up,
        # Python's round() rounds down
        cases: dict[float, float] = {0.018434999999999979: 0.01844, 0.037144999999999984: 0.03715,
                                     0.055684999999999985: 0.05569, 0.020774999999999988: 0.02078}
        for value, expected in cases.items():
            self.assertEqual(excel_round_value(value), expected, msg=repr(value))
            self.assertNotEqual(round(value, 5), expected, msg=repr(value))

    def test_below_a_tie_at_15_digits_rounds_down(self) -> None:
        # 0.0227949999999999 is below the tie even with 15 significant digits
        self.assertEqual(excel_round_value(0.0227949999999999), 0.02279)

    def test_halves_round_away_from_zero(self) -> None:
        cases: dict[tuple[float, int], float] = {(0.031255, 5): 0.03126, (-0.031255, 5): -0.03126,
                                                 (-0.018434999999999979, 5): -0.01844, (2.5, 0): 3.0,
                                                 (-2.5, 0): -3.0, (1234.5, 0): 1235.0}
        for (value, digits), expected in cases.items():
            self.assertEqual(excel_round_value(value, digits), expected, msg=f"ROUND({value!r}, {digits})")

    def test_other_digits_and_large_numbers(self) -> None:
        self.assertEqual(excel_round_value(0.0326049, 3), 0.033)
        self.assertEqual(excel_round_value(1234.5678, -2), 1200.0)
        self.assertEqual(excel_round_value(1e30), 1e30)
        self.assertEqual(excel_round_value(0.30000000000000004), 0.3)

    def test_no_negative_zero(self) -> None:
        # Excel's ROUND(-1E-9, 5) is 0; Python's round() gives -0.0, which a CSV file shows as "-0.0"
        self.assertEqual(math.copysign(1.0, excel_round_value(-1e-9)), 1.0)
        self.assertEqual(str(excel_round({1: -1e-9})[1]), "0.0")

    def test_infinities_are_unchanged(self) -> None:
        self.assertEqual(excel_round_value(math.inf), math.inf)
        self.assertEqual(excel_round_value(-math.inf), -math.inf)


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
        inputs["Rate"].update({14: 0.031124, 16: 0.031742, 35: 0.0325})
        self.assertEqual(calculate_sheet(**inputs)["BASIC_RFR"], self.res["BASIC_RFR"])


class TestCalculateSheetGovernmentBonds(unittest.TestCase):
    """The workbook example with instrument type GVT (values as calculated by Excel/VBA)."""

    def setUp(self) -> None:
        self.res = calculate_sheet(**example_inputs("GVT"))

    def test_fsp_and_llfr(self) -> None:
        self.assertEqual(self.res["FSP"], 20)
        self.assertEqual(self.res["LLFR"], 0.03188302797422029)

    def test_bootstrapped_curve(self) -> None:
        expected: Curve = {1: 0.02054744788766011, 2: 0.021390581515523904, 13: 0.02934421611279172,
                           14: 0.02968277400823564, 15: 0.029976190850953703, 20: 0.03084923141548605,
                           30: 0.030947158417286666, 31: 0.030900737881867854, 50: 0.029859727832683033}
        for t, rate in expected.items():
            self.assertAlmostEqual(self.res["BOOTSTRAPPED_CURVE"][t], rate, places=15, msg=f"maturity {t}")
        self.assertTrue(math.isnan(self.res["BOOTSTRAPPED_CURVE"][51]))

    def test_basic_rfr(self) -> None:
        expected: Curve = {1: 0.02076, 2: 0.02162, 10: 0.02827, 14: 0.03013, 20: 0.03133, 21: 0.03138,
                           25: 0.03157, 30: 0.03176, 40: 0.03204, 50: 0.03223, 60: 0.03235, 100: 0.03261,
                           150: 0.03274}
        for t, rate in expected.items():
            self.assertEqual(self.res["BASIC_RFR"][t], rate, msg=f"maturity {t}")
        self.assertEqual(len(self.res["BASIC_RFR"]), DEFAULT_MAX_MATURITY)

    def test_basic_rfr_up_to_fsp_is_the_input_rate_minus_cra(self) -> None:
        # before rounding: equal up to the conversion annual -> continuous -> annual
        zero = extrapolation(self.res["BOOTSTRAPPED_CURVE"], "Z", self.res["FSP"], 0.033, self.res["LLFR"], 0.11, "A")
        for t, rate in EXAMPLE_RATES.items():
            if t <= self.res["FSP"]:
                self.assertAlmostEqual(zero[t], rate - 0.001, places=15, msg=f"maturity {t}")

    def test_rounding_after_the_compounding_conversion(self) -> None:
        # 3y: 0.023795 - 0.001 = 0.022795 comes back from ln/exp as 0.0227949999999999, below the tie
        # even with Excel's 15 significant digits, so ROUND gives 0.02279 (as in Excel, L27), not 0.0228
        self.assertEqual(self.res["BASIC_RFR"][3], 0.02279)

    def test_basic_rfr_rounds_like_the_workbook(self) -> None:
        # Rates with six decimals ending in 5: up to the FSP the basic RFR (input rate - 10 bp) lands on a
        # 5-decimal tie up to floating-point noise. The workbook, recalculated with these inputs, gives the
        # values below; Python's round() would give 0.02077, 0.02765, 0.02978, 0.03046 and 0.03131.
        inputs = example_inputs("GVT")
        inputs["Rate"].update({1: 0.021775, 2: 0.022615, 3: 0.023825, 4: 0.024795, 5: 0.025725, 6: 0.026535,
                               7: 0.027355, 8: 0.028025, 9: 0.028655, 10: 0.029285, 11: 0.029765, 12: 0.030305,
                               13: 0.030785, 15: 0.031465, 20: 0.032315})
        res = calculate_sheet(**inputs)
        expected: Curve = {1: 0.02078, 9: 0.02766, 13: 0.02979, 15: 0.03047, 20: 0.03132}
        for t, rate in expected.items():
            self.assertEqual(res["BASIC_RFR"][t], rate, msg=f"maturity {t}")

    def test_llfr_is_calculated_from_the_market_rates(self) -> None:
        # a higher 30y rate raises the LLFR and the extrapolated rates, but not the rates up to the FSP
        inputs = example_inputs("GVT")
        inputs["Rate"][30] += 0.001
        res = calculate_sheet(**inputs)
        self.assertGreater(res["LLFR"], self.res["LLFR"])
        for t in range(1, DEFAULT_MAX_MATURITY + 1):
            if t <= 20:
                self.assertEqual(res["BASIC_RFR"][t], self.res["BASIC_RFR"][t], msg=f"maturity {t}")
            else:
                self.assertGreater(res["BASIC_RFR"][t], self.res["BASIC_RFR"][t], msg=f"maturity {t}")

    def test_coupon_frequency_does_not_change_the_result(self) -> None:
        inputs = example_inputs("GVT")
        inputs["CouponFreq"] = 2
        self.assertEqual(calculate_sheet(**inputs)["BASIC_RFR"], self.res["BASIC_RFR"])

    def test_differs_from_swaps(self) -> None:
        # the same rates read as par swap rates give a different curve
        self.assertNotEqual(calculate_sheet(**example_inputs("SWP"))["LLFR"], self.res["LLFR"])


if __name__ == "__main__":
    unittest.main()
