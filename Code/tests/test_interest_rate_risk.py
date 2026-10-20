"""Unit tests for interest_rate_risk.py.

Run from the project folder with either
    python -m unittest discover -s Code -v
    python -m pytest Code
"""

import math
import unittest
from pathlib import Path

from calculation import Curve, SheetInputs, SheetResults, calculate_sheet, excel_round_value
from data_io import read_csv, read_irr_inputs
from interest_rate_risk import (UFR_SHIFT, IRRInputs, IRRResults, apply_floor, bootstrapped_curve_ac,
                                calculate_irr_down_sheet, calculate_irr_up_sheet, shock_down, shock_up)
from tests.test_calculation import example_inputs

# irr_parameters.csv of the EIOPA example: sheet Parameters G16:L166 of the workbook
EXAMPLE_INPUT_DIR: Path = Path(__file__).resolve().parents[2] / "Examples" / "EIOPA_Example" / "Input"
# columns I, J, L and O of the sheets 'IRR Up Shock' and 'IRR Down Shock', as calculated by Excel/VBA for the
# workbook example (instrument GVT, total VA of 34 bp); #N/A as empty cells
WORKBOOK_VALUES_FILE: Path = Path(__file__).resolve().parent / "irr_workbook_values.csv"
# total VA of the workbook example (VA Term Structures L13)
EXAMPLE_VA: float = 0.0034


def example_irr_inputs() -> IRRInputs:
    """A fresh copy of the workbook's shock parameters (safe to modify in a test)."""
    irr: IRRInputs | None = read_irr_inputs(EXAMPLE_INPUT_DIR)
    assert irr is not None
    return irr


def workbook_values() -> dict[str, Curve]:
    """The workbook's values per column of irr_workbook_values.csv, {column: {maturity: value}}; #N/A as nan."""
    rows: list[dict[str, str]] = read_csv(WORKBOOK_VALUES_FILE)
    return {column: {int(row["Maturity"]): float(row[column]) if row[column] else math.nan for row in rows}
            for column in rows[0] if column != "Maturity"}


def zero_shocks(floor: float = -0.99) -> IRRInputs:
    """Shock parameters of 0 for every maturity, with a constant floor."""
    zero: Curve = {t: 0.0 for t in range(1, 151)}
    return IRRInputs(IRR_SM_UP=dict(zero), IRR_BM_UP=dict(zero), IRR_SM_DOWN=dict(zero), IRR_BM_DOWN=dict(zero),
                     RM_FLOOR={t: floor for t in range(1, 151)})


class TestShocks(unittest.TestCase):

    def test_upward_shock(self) -> None:
        # 1 year in the workbook: 2.076% * (1 + 61%) + 2.14%
        self.assertAlmostEqual(shock_up(0.02076, 0.61, 0.0214), 0.0548236, places=15)

    def test_downward_shock_may_be_negative(self) -> None:
        # 1 year in the workbook: 2.076% * (1 - 58%) - 1.16%
        self.assertAlmostEqual(shock_down(0.02076, 0.58, 0.0116), -0.0028808, places=15)

    def test_absolute_shock_applies_to_a_rate_of_zero(self) -> None:
        self.assertEqual(shock_up(0.0, 0.61, 0.0214), 0.0214)
        self.assertEqual(shock_down(0.0, 0.58, 0.0116), -0.0116)

    def test_negative_rates_are_shocked_too(self) -> None:
        # no special case for negative rates: the relative part makes them less negative downwards
        self.assertAlmostEqual(shock_down(-0.01, 0.5, 0.001), -0.006, places=15)
        self.assertAlmostEqual(shock_up(-0.01, 0.5, 0.001), -0.014, places=15)


class TestApplyFloor(unittest.TestCase):

    def test_rates_below_the_floor_are_raised(self) -> None:
        self.assertEqual(apply_floor({1: -0.02, 2: -0.01, 3: 0.01}, {1: -0.0125, 2: -0.0125, 3: -0.0125}),
                         {1: -0.0125, 2: -0.01, 3: 0.01})

    def test_floor_per_maturity(self) -> None:
        self.assertEqual(apply_floor({7: -0.012, 20: -0.012}, {7: -0.0125, 20: -0.00893}), {7: -0.012, 20: -0.00893})


class TestBootstrappedCurveAC(unittest.TestCase):

    def test_is_the_bootstrapped_curve_annually_compounded(self) -> None:
        for Instrument in ("SWP", "GVT"):
            inputs = example_inputs(Instrument)
            cc: Curve = calculate_sheet(**inputs)["BOOTSTRAPPED_CURVE"]
            ac: Curve = bootstrapped_curve_ac(inputs)
            for t in range(1, 51):
                self.assertAlmostEqual(ac[t], math.exp(cc[t]) - 1, places=15, msg=f"{Instrument} {t}")

    def test_ends_at_the_last_dlt_maturity(self) -> None:
        ac: Curve = bootstrapped_curve_ac(example_inputs("GVT"))
        self.assertEqual(len(ac), 150)
        self.assertFalse(math.isnan(ac[50]))
        self.assertTrue(all(math.isnan(ac[t]) for t in range(51, 151)))

    def test_government_bond_rates_after_the_cra(self) -> None:
        # at a DLT maturity the zero rate is the input rate minus the CRA of 10 bp
        self.assertAlmostEqual(bootstrapped_curve_ac(example_inputs("GVT"))[20], 0.03233 - 0.001, places=15)


class TestWorkbookExample(unittest.TestCase):
    """The workbook example, instrument GVT with a total VA of 34 bp: values as calculated by Excel/VBA."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.inputs: SheetInputs = example_inputs("GVT")
        cls.res: SheetResults = calculate_sheet(**cls.inputs)
        cls.up: IRRResults = calculate_irr_up_sheet(cls.inputs, cls.res, example_irr_inputs(), EXAMPLE_VA)
        cls.down: IRRResults = calculate_irr_down_sheet(cls.inputs, cls.res, example_irr_inputs(), EXAMPLE_VA)
        cls.workbook: dict[str, Curve] = workbook_values()

    def assertSameCurve(self, curve: Curve, column: str) -> None:
        """Every maturity 1..150 equal to the workbook column, nan where the workbook shows #N/A."""
        expected: Curve = self.workbook[column]
        self.assertEqual(sorted(curve), list(range(1, 151)))
        for t in range(1, 151):
            if math.isnan(expected[t]):
                self.assertTrue(math.isnan(curve[t]), f"{column} {t}: {curve[t]} instead of #N/A")
            else:
                self.assertEqual(curve[t], expected[t], f"{column} {t}")

    def test_ufr(self) -> None:
        self.assertEqual(self.up["UFR"], 0.0345)                                   # IRR Up Shock H11
        self.assertEqual(self.down["UFR"], 0.0315)                                 # IRR Down Shock H11

    def test_va(self) -> None:
        self.assertEqual((self.up["VA"], self.down["VA"]), (0.0034, 0.0034))      # H12

    def test_llfr(self) -> None:
        self.assertEqual(self.up["LLFR"], 0.04008083326240101)                     # IRR Up Shock L19
        self.assertEqual(self.down["LLFR"], 0.009993603597301537)                  # IRR Down Shock K19

    def test_llfr_with_va(self) -> None:
        self.assertEqual(self.up["LLFR_STAR"], 0.04120093017484616)                # IRR Up Shock O19
        self.assertEqual(self.down["LLFR_STAR"], 0.01111370050974669)              # IRR Down Shock N19

    def test_bootstrapped_curve(self) -> None:
        self.assertSameCurve(self.up["BOOTSTRAPPED_CURVE_AC"], "Bootstrapped Zero Rate AC")    # I22:I171
        self.assertSameCurve(self.down["BOOTSTRAPPED_CURVE_AC"], "Bootstrapped Zero Rate AC")

    def test_shocked_rates(self) -> None:
        self.assertSameCurve(self.up["IRR_SHOCK_AC"], "IRR Shock Up AC")           # J22:J171
        self.assertSameCurve(self.down["IRR_SHOCK_AC"], "IRR Shock Down AC")

    def test_shocked_curves(self) -> None:
        self.assertSameCurve(self.up["IRR_RFR"], "IRR Up RFR")                     # L22:L171
        self.assertSameCurve(self.down["IRR_RFR"], "IRR Down RFR")

    def test_shocked_curves_with_va(self) -> None:
        self.assertSameCurve(self.up["IRR_VA_CURVE"], "IRR Up with VA")            # O22:O171
        self.assertSameCurve(self.down["IRR_VA_CURVE"], "IRR Down with VA")

    def test_shock_is_applied_to_annually_compounded_rates(self) -> None:
        # the same shock on the continuously compounded rate would give exp(ln(1.02076) * 1.61 + 0.0214) - 1,
        # 5.599% instead of 5.482% at 1 year
        on_cc: float = math.exp(math.log(1 + self.up["BOOTSTRAPPED_CURVE_AC"][1]) * 1.61 + 0.0214) - 1
        self.assertAlmostEqual(on_cc, 0.05599, places=5)
        self.assertEqual(self.up["IRR_RFR"][1], 0.05482)

    def test_floor_does_not_bind(self) -> None:
        # the lowest rate after the downward shock is -0.288% at 1 year, above the floor of -1.25%
        floor: Curve = example_irr_inputs()["RM_FLOOR"]
        self.assertTrue(all(self.down["IRR_RFR"][t] > floor[t] for t in range(1, 151)))

    def test_up_above_and_down_below_the_basic_rfr(self) -> None:
        for t in range(1, 151):
            self.assertGreater(self.up["IRR_RFR"][t], self.res["BASIC_RFR"][t], t)
            self.assertLess(self.down["IRR_RFR"][t], self.res["BASIC_RFR"][t], t)


class TestShockedCurves(unittest.TestCase):

    def setUp(self) -> None:
        self.inputs: SheetInputs = example_inputs("GVT")
        self.res: SheetResults = calculate_sheet(**self.inputs)

    def test_shocked_rates_end_at_the_last_dlt_maturity(self) -> None:
        for res in (calculate_irr_up_sheet(self.inputs, self.res, example_irr_inputs(), EXAMPLE_VA),
                    calculate_irr_down_sheet(self.inputs, self.res, example_irr_inputs(), EXAMPLE_VA)):
            self.assertFalse(math.isnan(res["IRR_SHOCK_CC"][50]))
            self.assertTrue(all(math.isnan(res["IRR_SHOCK_AC"][t]) and math.isnan(res["IRR_SHOCK_CC"][t])
                                for t in range(51, 151)))
            # the extrapolated curves are complete
            self.assertFalse(any(math.isnan(res["IRR_RFR"][t]) or math.isnan(res["IRR_VA_CURVE"][t])
                                 for t in range(1, 151)))

    def test_shocked_rates_are_extrapolated_beyond_the_fsp(self) -> None:
        # up to the FSP the curve is the rounded shocked rate; beyond it, the extrapolation replaces the
        # shocked market rates (21 years: shocked 4.791%, extrapolated 4.761%)
        up: IRRResults = calculate_irr_up_sheet(self.inputs, self.res, example_irr_inputs())
        self.assertEqual(up["IRR_RFR"][20], excel_round_value(up["IRR_SHOCK_AC"][20], 5))
        self.assertAlmostEqual(up["IRR_SHOCK_AC"][21], 0.04791, places=5)
        self.assertEqual(up["IRR_RFR"][21], 0.04761)

    def test_without_va(self) -> None:
        up: IRRResults = calculate_irr_up_sheet(self.inputs, self.res, example_irr_inputs())
        down: IRRResults = calculate_irr_down_sheet(self.inputs, self.res, example_irr_inputs(), None)
        with_va: IRRResults = calculate_irr_up_sheet(self.inputs, self.res, example_irr_inputs(), EXAMPLE_VA)
        for res in (up, down):
            self.assertTrue(math.isnan(res["VA"]) and math.isnan(res["LLFR_STAR"]))
            self.assertTrue(all(math.isnan(r) for r in res["IRR_VA_CURVE"].values()))
            self.assertEqual(len(res["IRR_VA_CURVE"]), 150)
        # the curve without VA does not depend on the VA
        self.assertEqual((up["LLFR"], up["IRR_RFR"]), (with_va["LLFR"], with_va["IRR_RFR"]))

    def test_va_of_zero_gives_the_curve_without_va(self) -> None:
        for calculate in (calculate_irr_up_sheet, calculate_irr_down_sheet):
            res: IRRResults = calculate(self.inputs, self.res, example_irr_inputs(), 0.0)
            self.assertEqual(res["LLFR_STAR"], res["LLFR"])
            for t in range(1, 151):
                self.assertAlmostEqual(res["IRR_VA_CURVE"][t], res["IRR_RFR"][t], places=12, msg=t)

    def test_llfr_with_va(self) -> None:
        # LLFR* = LLFR + w_FSP * ln(1 + VA), with the weight 0.33 at the FSP of 20 years
        res: IRRResults = calculate_irr_up_sheet(self.inputs, self.res, example_irr_inputs(), 0.01)
        self.assertAlmostEqual(res["LLFR_STAR"], res["LLFR"] + 0.33 * math.log(1.01), places=15)

    def test_zero_shocks_only_shift_the_ufr(self) -> None:
        # without shocks, the curves are the basic RFR up to the FSP; beyond it the UFR is 15 bp higher or lower
        up: IRRResults = calculate_irr_up_sheet(self.inputs, self.res, zero_shocks())
        down: IRRResults = calculate_irr_down_sheet(self.inputs, self.res, zero_shocks())
        self.assertEqual((up["UFR"], down["UFR"]), (self.inputs["UFR"] + UFR_SHIFT, self.inputs["UFR"] - UFR_SHIFT))
        for res in (up, down):
            self.assertAlmostEqual(res["LLFR"], self.res["LLFR"], places=15)
            for t in range(1, 21):
                self.assertEqual(res["IRR_RFR"][t], self.res["BASIC_RFR"][t], t)
        for t in range(22, 151):
            self.assertGreater(up["IRR_RFR"][t], self.res["BASIC_RFR"][t], t)
            self.assertLess(down["IRR_RFR"][t], self.res["BASIC_RFR"][t], t)

    def test_floor_binds_on_low_rates(self) -> None:
        # government bond rates of -0.5%: after the CRA -0.6%, after the downward shock -0.6% * 42% - 1.16% at 1 year
        inputs: SheetInputs = example_inputs("GVT")
        inputs["Rate"] = {t: -0.005 for t in inputs["Rate"]}
        res: SheetResults = calculate_sheet(**inputs)
        floored: IRRResults = calculate_irr_down_sheet(inputs, res, example_irr_inputs(), EXAMPLE_VA)
        without_floor: IRRInputs = example_irr_inputs()
        without_floor["RM_FLOOR"] = {t: -0.99 for t in range(1, 151)}
        unfloored: IRRResults = calculate_irr_down_sheet(inputs, res, without_floor, EXAMPLE_VA)
        floor: Curve = example_irr_inputs()["RM_FLOOR"]
        self.assertLess(unfloored["IRR_RFR"][1], -0.0125)
        self.assertEqual(floored["IRR_RFR"][1], -0.0125)
        for t in range(1, 151):
            self.assertEqual(floored["IRR_RFR"][t], max(unfloored["IRR_RFR"][t], floor[t]), t)
            self.assertEqual(floored["IRR_VA_CURVE"][t], max(unfloored["IRR_VA_CURVE"][t], floor[t]), t)
        # the floor is applied last: the LLFRs and the shocked rates are those without a floor
        self.assertEqual((floored["LLFR"], floored["LLFR_STAR"]), (unfloored["LLFR"], unfloored["LLFR_STAR"]))
        self.assertEqual([floored["IRR_SHOCK_CC"][t] for t in range(1, 51)],
                         [unfloored["IRR_SHOCK_CC"][t] for t in range(1, 51)])

    def test_upward_shock_has_no_floor(self) -> None:
        irr: IRRInputs = example_irr_inputs()
        irr["RM_FLOOR"] = {t: 0.5 for t in range(1, 151)}
        self.assertEqual(calculate_irr_up_sheet(self.inputs, self.res, irr)["IRR_RFR"],
                         calculate_irr_up_sheet(self.inputs, self.res, example_irr_inputs())["IRR_RFR"])

    def test_swaps(self) -> None:
        inputs: SheetInputs = example_inputs("SWP")
        res: SheetResults = calculate_sheet(**inputs)
        up: IRRResults = calculate_irr_up_sheet(inputs, res, example_irr_inputs(), EXAMPLE_VA)
        down: IRRResults = calculate_irr_down_sheet(inputs, res, example_irr_inputs(), EXAMPLE_VA)
        self.assertGreater(up["LLFR"], res["LLFR"])
        self.assertLess(down["LLFR"], res["LLFR"])
        # the shock applies to the bootstrapped zero rates, not to the par swap rates
        zero: float = bootstrapped_curve_ac(inputs)[10]
        self.assertNotAlmostEqual(zero, inputs["Rate"][10] - 0.001, places=6)
        self.assertEqual(up["IRR_RFR"][10], excel_round_value(shock_up(zero, 0.3, 0.0105), 5))

    def test_curves_follow_max_maturity(self) -> None:
        # the regulatory table runs to 150 years; a shorter curve uses its first maturities
        inputs: SheetInputs = example_inputs("GVT")
        inputs["MAX_MATURITY"] = 55
        res: SheetResults = calculate_sheet(**inputs)
        up: IRRResults = calculate_irr_up_sheet(inputs, res, example_irr_inputs(), EXAMPLE_VA)
        down: IRRResults = calculate_irr_down_sheet(inputs, res, example_irr_inputs(), EXAMPLE_VA)
        for curve in (up["IRR_RFR"], up["IRR_VA_CURVE"], down["IRR_RFR"], down["IRR_VA_CURVE"]):
            self.assertEqual(sorted(curve), list(range(1, 56)))
        self.assertEqual(up["IRR_RFR"][20], workbook_values()["IRR Up RFR"][20])


if __name__ == "__main__":
    unittest.main()
