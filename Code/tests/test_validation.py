"""Unit tests for validation.py.

Run from the project folder with either
    python -m unittest discover -s Code -v
    python -m pytest Code
"""

import unittest

from calculation import SwapInputs
from tests.test_calculation import example_inputs
from validation import validate_inputs


class TestValidateInputs(unittest.TestCase):

    def assertInvalid(self, inputs: SwapInputs, *messages: str) -> None:
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
        del inputs["SwapRatesInit"][12]
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
        inputs["SwapRatesInit"][70] = 0.03
        self.assertInvalid(inputs, "maturities [70] lie beyond Max Maturity (60)")

    def test_coupon_frequency_below_one(self) -> None:
        inputs = example_inputs()
        inputs["CouponFreq"] = 0
        self.assertInvalid(inputs, "Coupon Frequency must be at least 1")

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
        inputs["SwapRatesInit"][10] = 2.927
        self.assertInvalid(inputs, "Input Rate must be a decimal", "[10]")

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
        del inputs["SwapRatesInit"][25]
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


if __name__ == "__main__":
    unittest.main()
