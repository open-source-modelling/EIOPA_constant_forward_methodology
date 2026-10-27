# eiopa-rfr

Solvency II risk-free interest rate term structures with EIOPA's new constant-forward method, from
swap rates or government bond zero rates: bootstrapping up to the first smoothing point, then
extrapolation of forward rates from the last liquid forward rate (LLFR) to the ultimate forward
rate (UFR). It also calculates the undertaking-specific volatility adjustment (VA) and the term
structures after the upward and downward interest rate shocks, without and with VA.

- Reproduces EIOPA's demonstration workbook *RFR extrapolation and VA calculation* (19 May 2026)
  exactly, following the *RFR Technical Documentation* EIOPA-BoS-26/198 (May 2026).
- Pure Python, standard library only, fully typed.
- The same code is kept as plain, readable scripts in the
  [repository](https://github.com/open-source-modelling/EIOPA_constant_forward_methodology), which
  also explains the method step by step.

> **Disclaimer.** This is not an EIOPA product and is not endorsed by EIOPA. It is provided for
> educational purposes, without warranty. EIOPA publishes the official risk-free rate term
> structures; check any result against EIOPA's publications before relying on it.

## Install

```bash
pip install eiopa-rfr
```

Requires Python 3.10 or later.

## Command line

```bash
eiopa-rfr --input Input --output Output
```

The input folder holds:

| File | Content |
| --- | --- |
| `curve.csv` | columns `Maturity`, `DLT`, `LLFR Weight`, `Input Rate` (par swap rates or annually compounded zero rates) |
| `parameters.csv` | `Instrument` (`SWP` or `GVT`), `Coupon Frequency`, `UFR`, `Convergence`, `CRA` (bp), `Max Maturity` |
| `va_parameters.csv` | optional: the VA inputs; without it, no VA is calculated |
| `irr_parameters.csv` | optional: the shock parameters per maturity; without it, no shocks are calculated |

The command prints a summary and writes the curves and parameters to the output folder as
`swaps_*.csv` or `government_bonds_*.csv`. Before a run it deletes only the result files of an
earlier run from the output folder; other files are kept. `python -m eiopa_rfr` does the same.
The columns of every file are described under
[Inputs and outputs](https://github.com/open-source-modelling/EIOPA_constant_forward_methodology#inputs-and-outputs).

## Python

```python
from eiopa_rfr import calculate_sheet, validate_inputs

rates = {1: 0.02176, 2: 0.022621, 3: 0.023795, 4: 0.0248, 5: 0.02569, 6: 0.02651,
         7: 0.02732, 8: 0.028, 9: 0.028631, 10: 0.02927, 11: 0.02979, 12: 0.03027,
         13: 0.030779, 15: 0.03143, 20: 0.03233, 25: 0.0325, 30: 0.032431,
         40: 0.03206, 50: 0.03131}

inputs = {
    "Instrument": "SWP",               # swaps; "GVT" for government bonds
    "CouponFreq": 1,                   # annual coupons; None for government bonds
    "CRA": 10,                         # basis points
    "UFR": 0.033,
    "alpha": 0.11,
    "dlt": {t: 1 for t in rates},      # DLT flags; other maturities count as 0
    "LLFRweightsIn": {20: 0.33, 25: 0.12, 30: 0.48, 40: 0.04, 50: 0.03},
    "Rate": rates,                     # par swap rates, or zero rates for government bonds
    "MAX_MATURITY": 150,
}
validate_inputs(inputs)                # raises ValueError listing every problem
res = calculate_sheet(**inputs)

print(res["FSP"], res["LLFR"])         # 20 0.03224887127830391
print(res["BASIC_RFR"][60])            # 0.03266  (annually compounded, rounded to 5 decimals)
```

The volatility adjustment and the interest rate shocks build on `inputs` and `res`:

```python
from eiopa_rfr import VAInputs, calculate_cssr_sheet, calculate_va_sheet, validate_va_inputs

va = VAInputs(RCS_CU=0.01, RCS_CO=None, DTTA=None, MACROVA_SELECTED=False, USA=None,
              MVFI=500000, MVFI_NOTIONAL=490000, BEL=870000, BEL_NOTIONAL=850000)
validate_va_inputs(va)
cssr = calculate_cssr_sheet(inputs, res, va)
va_res = calculate_va_sheet(inputs, res, va, cssr["CSSR"])
print(cssr["CSSR"], va_res["TOTAL_VA"])   # 0.5 0.0043
print(va_res["VA_CURVE"][20])             # 0.03653
```

`calculate_irr_up_sheet` and `calculate_irr_down_sheet` give the shocked curves, and `read_inputs`,
`read_va_inputs` and `read_irr_inputs` read the input files of a folder. Pass the folder explicitly:
the defaults are the repository's `Input` and `Output` folders.
[Using the algorithm in your own code](https://github.com/open-source-modelling/EIOPA_constant_forward_methodology#using-the-algorithm-in-your-own-code)
describes every input and output.

## Documentation

- [The method](https://github.com/open-source-modelling/EIOPA_constant_forward_methodology#the-method),
  with references to EIOPA's documentation
- [Verification](https://github.com/open-source-modelling/EIOPA_constant_forward_methodology#verification)
  against EIOPA's workbook
- [Differences from the EIOPA workbook and documentation](https://github.com/open-source-modelling/EIOPA_constant_forward_methodology#differences-from-the-eiopa-workbook-and-documentation)
- [Scope and limitations](https://github.com/open-source-modelling/EIOPA_constant_forward_methodology#scope-and-limitations)

## License

MIT. To cite the package, see
[CITATION.cff](https://github.com/open-source-modelling/EIOPA_constant_forward_methodology/blob/main/CITATION.cff).
