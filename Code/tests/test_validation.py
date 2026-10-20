"""Unit tests for validation.py.

Run from the project folder with either
    python -m unittest discover -s Code -v
    python -m pytest Code
"""

import unittest

from calculation import SheetInputs
from interest_rate_risk import IRRInputs
from tests.test_calculation import example_inputs
from tests.test_interest_rate_risk import example_irr_inputs
from tests.test_volatility_adjustment import example_va_inputs
from validation import maturity_ranges, validate_inputs, validate_irr_inputs, validate_va_inputs
from volatility_adjustment import VAInputs


class TestValidateInputs(unittest.TestCase):

    def assertInvalid(self, inputs: SheetInputs, *messages: str) -> None:
        """validate_inputs raises ValueError and the message contains every given text."""
        with self.assertRaises(ValueError) as context:
            validate_inputs(inputs)
        for message in messages:
            self.assertIn(message, str(context.exception))

    def test_workbook_example_is_valid(self) -> None:
        validate_inputs(example_inputs())

    def test_weights_with_floating_point_noise_are_valid(self) -> None:
        inputs = example_inputs()
        inputs["LLFRweightsIn"] = {20: 0.01, 25: 0.29, 30: 0.7}     # sum() gives 0.9999999999999999
        self.assertNotEqual(sum(inputs["LLFRweightsIn"].values()), 1.0)
        validate_inputs(inputs)

    def test_dlt_maturity_without_rate_is_allowed(self) -> None:
        # the VBA skips DLT points with an empty rate, so this is not an error
        inputs = example_inputs()
        del inputs["Rate"][12]
        validate_inputs(inputs)

    def test_max_maturity_below_one(self) -> None:
        inputs = example_inputs()
        inputs["MAX_MATURITY"] = 0
        self.assertInvalid(inputs, "Max Maturity must be at least 1")

    def test_shorter_max_maturity_is_valid(self) -> None:
        # the example's DLT points, weights and rates all lie at or below 50;
        # its DLT flags of 0 at 61..150 and zero weights beyond the Max Maturity are no problem
        inputs = example_inputs()
        inputs["MAX_MATURITY"] = 60
        inputs["LLFRweightsIn"][100] = 0.0
        validate_inputs(inputs)

    def test_data_beyond_max_maturity(self) -> None:
        # DLT flags, weights and rates at 40 and 50 lie beyond a Max Maturity of 35
        inputs = example_inputs()
        inputs["MAX_MATURITY"] = 35
        self.assertInvalid(inputs, "maturities [40, 50] lie beyond Max Maturity (35)")

    def test_rate_alone_beyond_max_maturity(self) -> None:
        # a rate at a maturity with DLT = 0 and no weight still lies outside the curve
        inputs = example_inputs()
        inputs["MAX_MATURITY"] = 60
        inputs["Rate"][70] = 0.03
        self.assertInvalid(inputs, "maturities [70] lie beyond Max Maturity (60)")

    def test_coupon_frequency_below_one(self) -> None:
        inputs = example_inputs()
        inputs["CouponFreq"] = 0
        self.assertInvalid(inputs, "Coupon Frequency must be at least 1 for swaps (is 0)")

    def test_swaps_without_coupon_frequency(self) -> None:
        inputs = example_inputs()
        inputs["CouponFreq"] = None
        self.assertInvalid(inputs, "Coupon Frequency must be at least 1 for swaps (is None)")

    def test_government_bond_example_is_valid(self) -> None:
        validate_inputs(example_inputs("GVT"))

    def test_government_bonds_do_not_need_a_coupon_frequency(self) -> None:
        # not used for government bonds, so neither a missing nor an unusual value is a problem
        for CouponFreq in (None, 0):
            inputs = example_inputs("GVT")
            inputs["CouponFreq"] = CouponFreq
            validate_inputs(inputs)

    def test_unknown_instrument(self) -> None:
        inputs = example_inputs()
        inputs["Instrument"] = "BND"
        self.assertInvalid(inputs, "Instrument must be SWP (swaps) or GVT (government bonds) (is BND)")

    def test_convergence_not_positive(self) -> None:
        inputs = example_inputs()
        inputs["alpha"] = 0
        self.assertInvalid(inputs, "Convergence must be positive")

    def test_ufr_in_percent(self) -> None:
        inputs = example_inputs()
        inputs["UFR"] = 3.3
        self.assertInvalid(inputs, "UFR must be a decimal")

    def test_input_rate_in_percent(self) -> None:
        inputs = example_inputs()
        inputs["Rate"][10] = 2.927
        self.assertInvalid(inputs, "Input Rate must be a decimal", "[10]")

    def test_rates_of_minus_100_percent_or_less(self) -> None:
        # the lower end of the range: a rate of -100% or less is not a decimal rate either
        inputs = example_inputs()
        inputs["UFR"] = -1.0
        inputs["Rate"][10] = -2.927
        self.assertInvalid(inputs, "UFR must be a decimal", "Input Rate must be a decimal", "[10]")

    def test_negative_rates_are_valid(self) -> None:
        inputs = example_inputs()
        inputs["UFR"] = -0.005
        inputs["Rate"] = {t: r - 0.035 for t, r in inputs["Rate"].items()}
        validate_inputs(inputs)

    def test_no_dlt_maturity_with_a_rate(self) -> None:
        inputs = example_inputs()
        inputs["dlt"] = {t: 0 for t in inputs["dlt"]}
        self.assertInvalid(inputs, "no maturity has DLT = 1 and an Input Rate")

    def test_negative_weight(self) -> None:
        inputs = example_inputs()
        inputs["LLFRweightsIn"] = {20: 1.1, 30: -0.1}
        self.assertInvalid(inputs, "LLFR weights must not be negative", "[30]")

    def test_no_positive_weight(self) -> None:
        inputs = example_inputs()
        inputs["LLFRweightsIn"] = {}
        self.assertInvalid(inputs, "no maturity has a positive LLFR weight")

    def test_weights_not_summing_to_one(self) -> None:
        inputs = example_inputs()
        inputs["LLFRweightsIn"][20] = 0.5
        self.assertInvalid(inputs, "LLFR weights must sum to 1")

    def test_weight_on_non_dlt_maturity(self) -> None:
        inputs = example_inputs()
        inputs["LLFRweightsIn"] = {20: 0.9, 21: 0.1}
        self.assertInvalid(inputs, "LLFR weights are only allowed on maturities with DLT = 1", "[21]")

    def test_weight_on_dlt_maturity_without_rate(self) -> None:
        inputs = example_inputs()
        del inputs["Rate"][25]
        self.assertInvalid(inputs, "LLFR weights are only allowed on maturities with DLT = 1", "[25]")

    def test_no_dlt_point_before_fsp(self) -> None:
        inputs = example_inputs()
        inputs["LLFRweightsIn"] = {1: 1.0}
        self.assertInvalid(inputs, "there must be a maturity with DLT = 1 and an Input Rate before the FSP (1)")

    def test_all_problems_are_reported_together(self) -> None:
        inputs = example_inputs()
        inputs["CouponFreq"] = 0
        inputs["UFR"] = 3.3
        inputs["LLFRweightsIn"][20] = 0.5
        self.assertInvalid(inputs, "Coupon Frequency", "UFR", "sum to 1")



class TestValidateVAInputs(unittest.TestCase):

    def assertInvalid(self, va: VAInputs, *messages: str) -> None:
        """validate_va_inputs raises ValueError and the message contains every given text."""
        with self.assertRaises(ValueError) as context:
            validate_va_inputs(va)
        for message in messages:
            self.assertIn(message, str(context.exception))

    def test_workbook_example_is_valid(self) -> None:
        validate_va_inputs(example_va_inputs())

    def test_without_the_values_for_the_cssr_is_valid(self) -> None:
        va = example_va_inputs()
        va.update(MVFI=None, MVFI_NOTIONAL=None, BEL=None, BEL_NOTIONAL=None)
        validate_va_inputs(va)

    def test_only_currency_rcs_is_valid(self) -> None:
        validate_va_inputs(VAInputs(RCS_CU=0.01, RCS_CO=None, DTTA=None, MVFI=None, MVFI_NOTIONAL=None, BEL=None,
                                    BEL_NOTIONAL=None, MACROVA_SELECTED=False, USA=None))

    def test_negative_rcs_is_valid(self) -> None:
        va = example_va_inputs()
        va["RCS_CU"] = -0.002
        validate_va_inputs(va)

    def test_rcs_in_basis_points(self) -> None:
        va = example_va_inputs()
        va.update(RCS_CU=100, RCS_CO=140)
        self.assertInvalid(va, "Currency RCS must be a decimal", "Country RCS must be a decimal")

    def test_rcs_of_minus_100_percent_or_less(self) -> None:
        va = example_va_inputs()
        va.update(RCS_CU=-1.0, RCS_CO=-140)
        self.assertInvalid(va, "Currency RCS must be a decimal", "Country RCS must be a decimal")

    def test_some_values_for_the_cssr_missing(self) -> None:
        va = example_va_inputs()
        va.update(MVFI_NOTIONAL=None, BEL_NOTIONAL=None)
        self.assertInvalid(va, "all together or not at all", "MV Fixed Income CSSR Curve, BEL CSSR Curve")

    def test_notional_va_of_zero(self) -> None:
        # 85% of 0.5 bp rounds to 0 bp
        va = example_va_inputs()
        va["RCS_CU"] = 0.00005
        self.assertInvalid(va, "the notional VA", "is 0")

    def test_notional_va_of_zero_is_valid_without_the_values(self) -> None:
        va = example_va_inputs()
        va.update(RCS_CU=0.00005, MVFI=None, MVFI_NOTIONAL=None, BEL=None, BEL_NOTIONAL=None)
        validate_va_inputs(va)

    def test_macro_va_needs_the_country_values(self) -> None:
        va = example_va_inputs()
        va.update(USA=None, MACROVA_SELECTED=True, RCS_CO=None, DTTA=None)
        self.assertInvalid(va, "needs the Country RCS", "needs the Debt to Total Assets ratio")

    def test_debt_to_total_assets_in_percent(self) -> None:
        va = example_va_inputs()
        va.update(USA=None, MACROVA_SELECTED=True, DTTA=50)
        self.assertInvalid(va, "Debt to Total Assets must be between 0 and 1")

    def test_negative_debt_to_total_assets(self) -> None:
        va = example_va_inputs()
        va.update(USA=None, MACROVA_SELECTED=True, DTTA=-0.1)
        self.assertInvalid(va, "Debt to Total Assets must be between 0 and 1")

    def test_debt_to_total_assets_of_zero_and_one_are_valid(self) -> None:
        va = example_va_inputs()
        for DTTA in (0.0, 1.0):
            va.update(USA=None, MACROVA_SELECTED=True, DTTA=DTTA)
            validate_va_inputs(va)

    def test_macro_va_and_usa_together(self) -> None:
        va = example_va_inputs()
        va["MACROVA_SELECTED"] = True
        self.assertInvalid(va, "the USA and the macro VA cannot be used together")

    def test_usa_above_105_percent(self) -> None:
        va = example_va_inputs()
        va["USA"] = 1.06
        self.assertInvalid(va, "the USA must be at most 105%")
        va["USA"] = 1.05
        validate_va_inputs(va)


class TestMaturityRanges(unittest.TestCase):

    def test_runs_become_ranges(self) -> None:
        self.assertEqual(maturity_ranges([1, 2, 3, 7, *range(51, 151)]), "1-3, 7, 51-150")

    def test_single_maturities_and_order(self) -> None:
        self.assertEqual(maturity_ranges([9, 5, 7]), "5, 7, 9")
        self.assertEqual(maturity_ranges([4]), "4")
        self.assertEqual(maturity_ranges([]), "")


class TestValidateIRRInputs(unittest.TestCase):

    def assertInvalid(self, irr: IRRInputs, *messages: str, MAX_MATURITY: int = 150) -> None:
        """validate_irr_inputs raises ValueError and the message contains every given text."""
        with self.assertRaises(ValueError) as context:
            validate_irr_inputs(irr, MAX_MATURITY)
        for message in messages:
            self.assertIn(message, str(context.exception))

    def test_workbook_example_is_valid(self) -> None:
        validate_irr_inputs(example_irr_inputs(), 150)

    def test_maturities_beyond_max_maturity_are_allowed(self) -> None:
        validate_irr_inputs(example_irr_inputs(), 55)

    def test_table_shorter_than_the_curve(self) -> None:
        self.assertInvalid(example_irr_inputs(), "smup is missing for maturities 151-200",
                           "rmfloor is missing for maturities 151-200", MAX_MATURITY=200)

    def test_missing_values(self) -> None:
        irr = example_irr_inputs()
        del irr["RM_FLOOR"][7]
        del irr["IRR_BM_DOWN"][3]
        del irr["IRR_BM_DOWN"][4]
        self.assertInvalid(irr, "rmfloor is missing for maturities 7", "bmdown is missing for maturities 3-4")

    def test_relative_shock_in_percent(self) -> None:
        irr = example_irr_inputs()
        irr["IRR_SM_UP"][1] = 61
        irr["IRR_SM_DOWN"][2] = -0.51
        self.assertInvalid(irr, "smup must be a decimal between 0 and 1, e.g. 0.61 for 61%; check maturities 1",
                           "smdown must be a decimal between 0 and 1")

    def test_relative_shocks_of_zero_and_one_are_valid(self) -> None:
        irr = example_irr_inputs()
        irr["IRR_SM_UP"][1] = 0.0
        irr["IRR_SM_DOWN"][1] = 1.0
        validate_irr_inputs(irr, 150)

    def test_absolute_shock_in_basis_points(self) -> None:
        irr = example_irr_inputs()
        irr["IRR_BM_UP"][1] = 214
        irr["IRR_BM_DOWN"][2] = -0.0099
        self.assertInvalid(irr, "bmup must be a decimal between 0 and 1, e.g. 0.0214 for 214 basis points",
                           "bmdown must be a decimal between 0 and 1")

    def test_floor_in_percent(self) -> None:
        irr = example_irr_inputs()
        irr["RM_FLOOR"][1] = -1.25
        self.assertInvalid(irr, "rmfloor must be a decimal, e.g. -0.0125 for -1.25%; check maturities 1")

    def test_values_beyond_max_maturity_are_checked_too(self) -> None:
        irr = example_irr_inputs()
        irr["IRR_SM_UP"][120] = 20
        self.assertInvalid(irr, "smup must be a decimal", "check maturities 120", MAX_MATURITY=55)

    def test_all_problems_are_reported_together(self) -> None:
        irr = example_irr_inputs()
        irr["IRR_SM_UP"][1] = 61
        del irr["RM_FLOOR"][150]
        self.assertInvalid(irr, "invalid interest rate shock inputs", "smup must be a decimal", "rmfloor is missing")


if __name__ == "__main__":
    unittest.main()
