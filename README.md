# EIOPA risk-free rate extrapolation in Python

A transparent Python implementation of EIOPA's new method for deriving the Solvency II basic
risk-free interest rate term structure from swap rates: constant-forward bootstrapping up to
the first smoothing point, followed by extrapolation of forward rates from the last liquid
forward rate (LLFR) to the ultimate forward rate (UFR).

It is written for actuaries who want to

- **understand the change**: the method, its formulas and how it differs from what EIOPA's
  documents and Excel tool show, step by step;
- **use the algorithm in their own code**: `Code/calculation.py` depends only on the Python
  standard library and can be copied into another project as it is.

The code is a line-by-line port of the VBA in EIOPA's demonstration workbook and reproduces
that workbook exactly (see [Verification](#verification)).

> **Disclaimer.** This is not an EIOPA product and is not endorsed by EIOPA. It is provided for
> educational purposes, without warranty. EIOPA publishes the official risk-free rate term
> structures; check any result against EIOPA's publications before relying on it.

---

## Contents

- [What changes](#what-changes)
- [The method](#the-method)
- [Quick start](#quick-start)
- [Inputs and outputs](#inputs-and-outputs)
- [Using the algorithm in your own code](#using-the-algorithm-in-your-own-code)
- [Verification](#verification)
- [Differences from the EIOPA workbook and documentation](#differences-from-the-eiopa-workbook-and-documentation)
- [Scope and limitations](#scope-and-limitations)
- [Repository structure](#repository-structure)
- [Sources](#sources)
- [License](#license)

---

## What changes

The Solvency II review, Directive (EU) 2025/2 and Delegated Regulation (EU) 2026/269, changes the
interpolation and extrapolation of the basic risk-free rate term structure. The Smith-Wilson
method used so far is replaced by the methodology implemented in this repository.

- **Interpolation** with a constant forward rate between the deep, liquid and transparent (DLT)
  maturities. For swaps, the zero rates are bootstrapped from the par swap rates.
- **Extrapolation** of the forward rates after the first smoothing point (FSP): they start at the
  last liquid forward rate (LLFR) and converge to the UFR at a speed set by a convergence
  parameter α.

The convergence parameter is α = 11% for all currencies except the Swedish krona (α = 40%).
Undertakings allowed to use the phasing-in mechanism of Article 77a(2) of the Solvency II
Directive use the values of Table 4 of EIOPA's technical documentation:

| Period | Swedish krona | All other currencies |
|---|---:|---:|
| 2027 | 70.0% | 20.0% |
| 2028 | 64.0% | 18.2% |
| 2029 | 58.0% | 16.4% |
| 2030 | 52.0% | 14.6% |
| 2031 | 46.0% | 12.8% |
| 2032 | 40.0% | 11.0% |

The review also changes the volatility adjustment (VA), which becomes undertaking-specific. The VA
is not part of this repository (see [Scope and limitations](#scope-and-limitations)).

---

## The method

References are to EIOPA's *RFR Technical Documentation*, EIOPA-BoS-26/198 (May 2026).
Superscript *c* denotes a continuously compounded rate; *t* is a maturity in whole years.

### 1. Inputs (§8.1)

Par swap rates at the DLT maturities $t_1 < \dots < t_L$, the credit risk adjustment (CRA, in
basis points), the coupon frequency $m$, the UFR, α, and the LLFR weights.

The first smoothing point (FSP) $t_F$ is the longest DLT maturity for which enough bonds of that or
a longer maturity are outstanding (the residual volume criterion, §8.2). It comes from EIOPA's DLT
assessment. The LLFR weights start at the FSP (§8.3.3), so the code, like EIOPA's workbook, takes
the first maturity with a positive LLFR weight as the FSP.

### 2. Bootstrapping (§8.5.2, Annex D.1–D.6)

Each swap is valued at par, with its rate reduced by the CRA:

$$\frac{s_{t_k} - \text{CRA}}{m}\sum_{j=1}^{m\,t_k} d_{j/m} \;+\; d_{t_k} = 1 .$$

Between consecutive DLT maturities the periodic forward rate is constant. For the first DLT
maturity it equals the periodic coupon. For each later one it is the single unknown of the par
equation above, solved with Newton-Raphson. The result is a zero rate for every whole year up to
the last DLT maturity.

### 3. Last liquid forward rate (§8.3, §8.5.6)

With $f^c(a,b) = \dfrac{b\,z^c_b - a\,z^c_a}{b-a}$ the forward rate from $a$ to $b$:

$$\text{LLFR}^c = w_{t_F}\, f^c(t_F - 1, t_F) + \sum_{k=F+1}^{L} w_{t_k}\, f^c(t_F, t_k).$$

The first term is the one-year forward rate into the FSP (§8.5.6, §10.2.1). Because the forward
rate is constant between DLT maturities, it equals the forward rate from the last DLT maturity
before the FSP to the FSP, which is how EIOPA's VBA and this code calculate it. The weights are
each maturity's share of the annual average notional of traded swaps and sum to 1 (§8.3.3). If
the FSP is the last DLT maturity, its weight is 100% (§8.3.4).

### 4. Extrapolation (§8.5.5, §8.5.7, §8.5.8)

For $h = 1, 2, \dots$ years after the FSP:

$$f^c(t_F, t_F+h) = \ln(1+\text{UFR}) + \big(\text{LLFR}^c - \ln(1+\text{UFR})\big)\,\frac{1-e^{-\alpha h}}{\alpha h}$$

$$z^c_{t_F+h} = \frac{t_F\, z^c_{t_F} + h\, f^c(t_F, t_F+h)}{t_F+h}, \qquad z_t = e^{z^c_t} - 1 .$$

The forward rate starts at the LLFR at the FSP and converges to the UFR; a larger α means faster
convergence. Up to the FSP the curve is the bootstrapped curve. Rates beyond the FSP are only
used for the LLFR. EIOPA publishes the result for maturities 1 to 150, annually compounded; the
workbook rounds them to 5 decimals.

---

## Quick start

Requirements: Python 3.10 or later, because the code uses the `X | Y` type syntax. It has been run
and tested with Python 3.13; older versions were only checked with a type checker. The calculation
needs no other packages.

```bash
python Code/main.py
```

This reads the two files in `Input/`, validates them, calculates the curves, prints a summary
and writes the results to `Output/`:

```text
FSP  : 20 years
LLFR : 3.22488713% (continuously compounded)
Maturity  Bootstrapped CC  Basic RFR
       1         2.05474%     2.076%
      ...
      20         3.15865%     3.209%
      50         2.92567%     3.260%
      60             #N/A     3.266%
     150             #N/A     3.287%
```

The bootstrapped curve exists only up to the last DLT maturity (50 years here), hence `#N/A`
beyond it. The basic risk-free rate continues to 150 years through the extrapolation.

Invalid inputs stop the run with a message listing every problem, for example:

```text
Input error: invalid inputs:
  - UFR must be a decimal, e.g. 0.033 for 3.3% (is 3.3)
  - LLFR weights must sum to 1 (sum is 1.0999999999999999)
```

---

## Inputs and outputs

### `Input/parameters.csv`

| Parameter | Meaning | Example |
|---|---|---|
| `Coupon Frequency` | Coupon payments per year of the swaps | `1` |
| `UFR` | Ultimate forward rate, annually compounded, as a decimal | `0.033` |
| `Convergence` | Convergence parameter α, as a decimal | `0.11` |
| `CRA` | Credit risk adjustment in basis points | `10` |
| `Max Maturity` | Longest maturity of the curves, in years | `150` |

### `Input/swap_curve.csv`

One row per maturity with the columns `Maturity, DLT, LLFR Weight, Input Rate`:

```text
Maturity,DLT,LLFR Weight,Input Rate
1,1,0,0.02176
...
19,0,0,0.032272
20,1,0.33,0.03233
...
```

- `DLT` is 1 for the maturities used in the calculation and 0 otherwise. Rates at maturities with
  DLT = 0 are ignored.
- `LLFR Weight` is positive only at the FSP and the DLT maturities after it. The first maturity
  with a positive weight is the FSP.
- `Input Rate` is the market par swap rate before the CRA, as a decimal. It may be empty.
- Missing maturities count as DLT 0, weight 0 and no rate.

Both files may use `,` or `;` as separator and `.` or `,` as decimal mark, so files saved by
Excel with European settings are read correctly.

The DLT maturities, the FSP, the LLFR weights and the UFR are not calculated here. They come from
EIOPA's DLT assessment and UFR methodology, published with the risk-free rates.

### `Output/`

| File | Contents |
|---|---|
| `swaps_curves.csv` | Per maturity: the inputs, `Bootstrapped Zero Rate CC` (continuously compounded; empty after the last DLT maturity) and `Basic RFR` (annually compounded, rounded to 5 decimals) |
| `swaps_parameters.csv` | The parameters used, the FSP and the LLFR (continuously compounded) |

Rates are written with full precision.

---

## Using the algorithm in your own code

`Code/calculation.py` contains the whole calculation and imports only `math` and `typing`. Copy
it into your project, or add the `Code` folder to the import path.

### The complete calculation

```python
import sys
sys.path.insert(0, "Code")            # or copy Code/calculation.py into your project

from calculation import calculate_sheet
from validation import validate_inputs

rates = {1: 0.02176, 2: 0.022621, 3: 0.023795, 4: 0.0248, 5: 0.02569, 6: 0.02651,
         7: 0.02732, 8: 0.028, 9: 0.028631, 10: 0.02927, 11: 0.02979, 12: 0.03027,
         13: 0.030779, 15: 0.03143, 20: 0.03233, 25: 0.0325, 30: 0.032431,
         40: 0.03206, 50: 0.03131}

inputs = {
    "CouponFreq": 1,                   # annual coupons
    "CRA": 10,                         # basis points
    "UFR": 0.033,
    "alpha": 0.11,
    "dlt": {t: 1 for t in rates},      # DLT flags; other maturities count as 0
    "LLFRweightsIn": {20: 0.33, 25: 0.12, 30: 0.48, 40: 0.04, 50: 0.03},
    "SwapRatesInit": rates,
    "MAX_MATURITY": 150,
}
validate_inputs(inputs)                # raises ValueError listing every problem
res = calculate_sheet(**inputs)

print(res["FSP"], res["LLFR"])         # 20 0.03224887127830391
print(res["BASIC_RFR"][60])            # 0.03266  (annually compounded, rounded to 5 decimals)
```

`res` holds `FSP`, `LLFR` (continuously compounded), `BOOTSTRAPPED_CURVE` (continuously
compounded, `nan` after the last DLT maturity) and `BASIC_RFR`. Both curves are dicts
`{maturity: rate}` for maturities 1 to `MAX_MATURITY`.

### Only the extrapolation

If you already have a zero curve, for example from government bonds or your own bootstrapping,
you can use the LLFR and extrapolation steps on their own:

```python
import math
import sys
sys.path.insert(0, "Code")

from calculation import extrapolation, get_llfr

my_zero_curve = {t: 0.025 + 0.0004 * min(t, 30) for t in range(1, 51)}   # annually compounded
zero_cc = {t: math.log(1 + z) for t, z in my_zero_curve.items()}        # continuously compounded
dlt = {t: 1 for t in (1, 2, 3, 5, 7, 10, 15, 20, 30, 50)}

llfr = get_llfr(zero_cc, dlt, {20: 0.4, 30: 0.4, 50: 0.2})              # FSP = 20
curve = extrapolation(zero_cc, "Z", 20, 0.033, llfr, 0.11, "A", MAX_MATURITY=150)
print(round(curve[150], 5))                                             # annually compounded
```

### Conventions

- Maturities are whole years. Curves, flags and weights are dicts keyed by maturity, e.g.
  `{1: 0.02176, 2: 0.022621}`.
- Rates are decimals (0.033 = 3.3%). The CRA is in basis points.
- `nan` plays the role of Excel's `#N/A`.
- Names follow the VBA in EIOPA's workbook to make the two easy to compare: `bootstrap_swaps` is
  `BootstrapSwaps`, `get_llfr` is `GetLLFR`, and `SwapRatesInit`, `CouponFreq`, `LLFRweightsIn`,
  `MaxTenor` and so on keep their VBA names. Every function's docstring describes its inputs,
  outputs and errors, with references to the documentation.

---

## Verification

- **Against EIOPA's workbook.** The results were compared with *RFR extrapolation and VA
  calculation (19 May 2026).xlsm*, recalculated by Excel and its VBA. The FSP, the LLFR, the
  bootstrapped curve and the basic risk-free rate match exactly, with a largest difference of 0,
  in two cases:
  - the workbook's example: annual coupons, CRA 10 bp, UFR 3.3%, α 11%;
  - a variant with quarterly coupons, CRA 30 bp, UFR 4.3%, α 40% and a 1-year swap rate of 2.5%.
- **Unit tests** cover the calculation, the validation, reading and writing, and the main script,
  with full line and branch coverage of the calculation, validation and input/output modules.
  Several tests check mathematical properties rather than stored numbers: every DLT swap reprices
  at par, a flat par curve gives a flat zero curve, forwards are constant between DLT points, and
  the extrapolated forward follows the formula and converges to the UFR.

Run the tests from the repository folder with either:

```bash
python -m unittest discover -s Code
```

```bash
python -m pytest Code
```

`Check/swaps_workbook_check.csv` holds the result of a comparison run. The script that produced
it is not part of this repository. To check the example yourself, compare `Output/swaps_curves.csv`
with the values the workbook in `Sources/` stores for cells H13, L22, K25:K174 and L25:L174 of the
sheet *Input Data & Extrapolation*. The workbook's inputs are the same as those in `Input/`.

---

## Differences from the EIOPA workbook and documentation

None of these change the results.

| Topic | This code and the VBA | EIOPA documentation |
|---|---|---|
| Solver | Newton-Raphson, tolerance 1e-15, at most 500 iterations | EIOPA's production code uses MATLAB `fzero` (D.5.12) |
| Newton starting guess | As the VBA: `forward(t - 1) / CouponFreq`, where `t` counts DLT points, not years | The periodic forward of the previous interval (D.6.3). Only the starting point differs, not the solution. |
| First LLFR forward | From the last DLT maturity before the FSP to the FSP | One-year forward into the FSP (8.5.6); the same number under constant forwards |
| LLFR weights not summing to 1 | Python raises an error; the VBA computes `sumw` but does not use it, and the sheet shows "ERROR" | Weights sum to 100% (8.3.3) |
| Longest maturity | `Max Maturity` input, 150 by default | The VBA writes the literal 150; EIOPA publishes 1–150 (8.1.7) |

One difference from the workbook could in principle change a result: Python's `round` rounds an
exact tie to the even digit, while Excel's `ROUND` rounds it away from zero. A rate would have to
fall exactly halfway between two 5-decimal values, which did not happen in the cases above.

---

## Scope and limitations

Covered: the basic risk-free rate term structure from **swaps**, i.e. the sheet *Input Data &
Extrapolation* of EIOPA's workbook.

Not covered:

- **Government bonds** as input instruments.
- **Currency risk adjustment** for currencies pegged to the euro (§7). Include it in the CRA if needed.
- **Volatility adjustment**, curves with VA, credit spread sensitivity ratio and the interest rate
  risk shocks.
- **DLT assessment, FSP, LLFR weights and UFR**: these are inputs, taken from EIOPA's publications.
- **Phasing-in of α**: choose the α of the relevant year yourself (see the table above).
- Maturities below one year (EIOPA also publishes from one year onwards, §3.3.1) and
  non-integer maturities.

---

## Repository structure

```text
Code/
  main.py              reads Input, validates, calculates, prints a summary, writes Output
  calculation.py       the calculation (port of EIOPA's VBA module mExtrapolation); standard library only
  validation.py        checks that the inputs are complete and plausible
  data_io.py           reading Input/ and writing Output/
  tests/               unit tests
Input/                 swap_curve.csv, parameters.csv
Output/                swaps_curves.csv, swaps_parameters.csv
Check/                 result of a comparison with the workbook
Sources/               EIOPA's technical documentation and demonstration workbook (example inputs)
```

---

## Sources

- EIOPA, *RFR Technical Documentation: the methodology to derive EIOPA's risk-free interest rate
  term structures*, EIOPA-BoS-26/198, May 2026.
- EIOPA, *RFR extrapolation and VA calculation* demonstration workbook, version of 19 May 2026.
- EIOPA's risk-free interest rate term structures page:
  <https://www.eiopa.europa.eu/tools-and-data/risk-free-interest-rate-term-structures_en>

EIOPA's documentation and workbook are © EIOPA. Check EIOPA's terms before redistributing them
in a public repository, or link to them instead.

---

## License

The code is released under the MIT License (see `LICENSE`). The license does not cover EIOPA's
documentation and workbook in `Sources/`.
