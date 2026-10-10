# EIOPA risk-free rate extrapolation in Python

A transparent Python implementation of EIOPA's new method for deriving the Solvency II basic
risk-free interest rate term structure from swap rates or government bond zero rates:
constant-forward bootstrapping up to the first smoothing point, followed by extrapolation of
forward rates from the last liquid forward rate (LLFR) to the ultimate forward rate (UFR).

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
- [Inputs and outputs](#inputs-and-outputs): files, columns and [example results](#example-results)
- [Using the algorithm in your own code](#using-the-algorithm-in-your-own-code): inputs and outputs of
  [`calculate_sheet`](#inputs-of-calculate_sheet) and a [function reference](#function-reference)
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
  maturities. For swaps, the zero rates are bootstrapped from the par swap rates; for government
  bonds, the input zero rates are interpolated directly.
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

Market rates at the DLT maturities $t_1 < \dots < t_L$: par swap rates (instrument `SWP`) or
annually compounded government bond zero rates (instrument `GVT`). Further the credit risk
adjustment (CRA, in basis points), the coupon frequency $m$ of the swaps, the UFR, α, and the
LLFR weights.

The first smoothing point (FSP) $t_F$ is the longest DLT maturity for which enough bonds of that or
a longer maturity are outstanding (the residual volume criterion, §8.2). It comes from EIOPA's DLT
assessment. The LLFR weights start at the FSP (§8.3.3), so the code, like EIOPA's workbook, takes
the first maturity with a positive LLFR weight as the FSP.

### 2. Bootstrapping (§8.5.2, Annex D.1–D.6)

**Swaps.** Each swap is valued at par, with its rate reduced by the CRA:

$$\frac{s_{t_k} - \text{CRA}}{m}\sum_{j=1}^{m \cdot t_k} d_{j/m} + d_{t_k} = 1 .$$

Between consecutive DLT maturities the periodic forward rate is constant. For the first DLT
maturity it equals the periodic coupon. For each later one it is the single unknown of the par
equation above, solved with Newton-Raphson. The result is a zero rate for every whole year up to
the last DLT maturity.

**Government bonds.** The input rates are already zero rates. After deducting the CRA they are
converted to continuous compounding, $z^c_{t_k} = \ln(1 + r_{t_k} - \text{CRA})$, and the forward
rate is constant between consecutive DLT maturities:

$$f^c(t_{k-1}, t_k) = \frac{t_k \cdot z^c_{t_k} - t_{k-1} \cdot z^c_{t_{k-1}}}{t_k - t_{k-1}} .$$

No solver is needed and the DLT rates are reproduced exactly. Before the first DLT maturity its
rate is held flat. The workbook does this in the VBA function `BootstrapZeros`; `BootstrapCurve`
chooses between it and `BootstrapSwaps` from the instrument type (cell H11).

### 3. Last liquid forward rate (§8.3, §8.5.6)

With $f^c(a,b) = \dfrac{b \cdot z^c_b - a \cdot z^c_a}{b-a}$ the forward rate from $a$ to $b$:

$$\text{LLFR}^c = w_{t_F} \cdot f^c(t_F - 1, t_F) + \sum_{k=F+1}^{L} w_{t_k} \cdot f^c(t_F, t_k).$$

The first term is the one-year forward rate into the FSP (§8.5.6, §10.2.1). Because the forward
rate is constant between DLT maturities, it equals the forward rate from the last DLT maturity
before the FSP to the FSP, which is how EIOPA's VBA and this code calculate it. The weights are
each maturity's share of the annual average notional of traded swaps and sum to 1 (§8.3.3). If
the FSP is the last DLT maturity, its weight is 100% (§8.3.4).

### 4. Extrapolation (§8.5.5, §8.5.7, §8.5.8)

For $h = 1, 2, \dots$ years after the FSP:

$$f^c(t_F, t_F+h) = \ln(1+\text{UFR}) + \big(\text{LLFR}^c - \ln(1+\text{UFR})\big) \cdot \frac{1-e^{-\alpha h}}{\alpha h}$$

$$z^c_{t_F+h} = \frac{t_F \cdot z^c_{t_F} + h \cdot f^c(t_F, t_F+h)}{t_F+h}, \qquad z_t = e^{z^c_t} - 1 .$$

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
Instrument: SWP (swaps)
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

To calculate the curve from government bonds instead, set `Instrument` to `GVT` in
`Input/parameters.csv` and put the bond zero rates in `Input/curve.csv`. The run then prints
`Instrument: GVT (government bonds)` and writes `government_bonds_curves.csv` and
`government_bonds_parameters.csv`. [Example results](#example-results) compares both curves.

Invalid inputs stop the run with a message listing every problem, for example:

```text
Input error: invalid inputs:
  - UFR must be a decimal, e.g. 0.033 for 3.3% (is 3.3)
  - LLFR weights must sum to 1 (sum is 1.0999999999999999)
```

---

## Inputs and outputs

`python Code/main.py` reads two files from `Input/` and writes two files to `Output/`. The
instrument in `parameters.csv` decides how the rates in `curve.csv` are read and how the result
files are named.

| Instrument | Input rates in `curve.csv` | Bootstrapping | Result files |
|---|---|---|---|
| `SWP` (swaps) | Par swap rates | `bootstrap_swaps` (VBA `BootstrapSwaps`) | `swaps_curves.csv`, `swaps_parameters.csv` |
| `GVT` (government bonds) | Zero rates, annually compounded | `bootstrap_zeros` (VBA `BootstrapZeros`) | `government_bonds_curves.csv`, `government_bonds_parameters.csv` |

### Input file `Input/parameters.csv`

Columns `Parameter, Value`, one row per parameter. The order of the rows does not matter and the
names are not case sensitive.

| Parameter | Meaning | Format | Required | Sheet cell | Example |
|---|---|---|---|---|---|
| `Instrument` | Instrument type of the input rates | `SWP` (swaps) or `GVT` (government bonds) | Always | H11 | `SWP` |
| `Coupon Frequency` | Coupon payments per year of the swaps | Whole number ≥ 1 | Swaps only; ignored for government bonds | H12 | `1` |
| `UFR` | Ultimate forward rate, annually compounded | Decimal | Always | H14 | `0.033` |
| `Convergence` | Convergence parameter α | Decimal > 0 | Always | H15 | `0.11` |
| `CRA` | Credit risk adjustment, deducted from every input rate | Basis points | Always | H16 | `10` |
| `Max Maturity` | Longest maturity of the curves | Whole number of years ≥ 1 | Always | 150 rows on the sheet | `150` |

### Input file `Input/curve.csv`

One row per maturity, with these columns (sheet *Input Data & Extrapolation*, columns G to J):

| Column | Meaning | Format | Sheet column | Example (20 years) |
|---|---|---|---|---|
| `Maturity` | Maturity | Whole number of years, 1 to `Max Maturity` | G | `20` |
| `DLT` | 1 if the maturity is deep, liquid and transparent and its rate is used; 0 otherwise | `0` or `1`; empty counts as 0 | H | `1` |
| `LLFR Weight` | Weight of the maturity in the LLFR; positive only at the FSP and the DLT maturities after it | Decimal; empty counts as 0; the weights sum to 1 | I | `0.33` |
| `Input Rate` | Market rate before the CRA: the par swap rate (`SWP`) or the annually compounded zero rate (`GVT`) | Decimal; may be empty | J | `0.03233` |

```text
Maturity,DLT,LLFR Weight,Input Rate
1,1,0,0.02176
...
19,0,0,0.032272
20,1,0.33,0.03233
...
```

- Rates at maturities with DLT = 0 are ignored, and so is a DLT maturity without a rate.
- The first maturity with a positive LLFR weight is the FSP.
- Missing maturities count as DLT 0, weight 0 and no rate.

Both files may use `,` or `;` as separator and `.` or `,` as decimal mark, so files saved by
Excel with European settings are read correctly.

The DLT maturities, the FSP, the LLFR weights and the UFR are not calculated here. They come from
EIOPA's DLT assessment and UFR methodology, published with the risk-free rates.

### Output file `Output/swaps_curves.csv` or `Output/government_bonds_curves.csv`

One row per maturity 1 to `Max Maturity`. The first four columns repeat the inputs.

| Column | Meaning | Compounding | Sheet column |
|---|---|---|---|
| `Maturity` | Maturity in years | | G |
| `DLT` | DLT flag from the input | | H |
| `LLFR Weight` | LLFR weight from the input | | I |
| `Input Rate` | Input rate; empty where there is none | As in the input | J |
| `Bootstrapped Zero Rate CC` | Zero rate after the CRA, bootstrapped up to the last DLT maturity; empty after it (Excel's `#N/A`) | Continuous | K |
| `Basic RFR` | Basic risk-free interest rate term structure, rounded to 5 decimals | Annual | L |

### Output file `Output/swaps_parameters.csv` or `Output/government_bonds_parameters.csv`

Columns `Parameter, Value`:

| Parameter | Meaning | Sheet cell |
|---|---|---|
| `Instrument` | `SWP` or `GVT` | H11 |
| `Coupon Frequency` | Coupon payments per year; swaps only | H12 |
| `UFR`, `Convergence`, `CRA`, `Max Maturity` | The parameters used | H14:H16 |
| `FSP` | First smoothing point in years: the first maturity with a positive LLFR weight | H13 |
| `LLFR (CC)` | Last liquid forward rate, calculated from the bootstrapped curve and the LLFR weights, continuously compounded | L22 |

All rates are decimals written with full precision.

### Example results

The workbook's example inputs (`Input/`: CRA 10 bp, UFR 3.3%, α 11%, FSP 20) give these
results, once with the rates read as par swap rates and once as government bond zero rates:

| | Input rate | Swaps: bootstrapped (CC) | Swaps: basic RFR | Government bonds: bootstrapped (CC) | Government bonds: basic RFR |
|---|---:|---:|---:|---:|---:|
| **LLFR (CC)** | | | 3.22489% | | 3.18830% |
| 1 | 2.1760% | 2.05474% | 2.076% | 2.05474% | 2.076% |
| 2 | 2.2621% | 2.13997% | 2.163% | 2.13906% | 2.162% |
| 5 | 2.5690% | 2.44854% | 2.479% | 2.43901% | 2.469% |
| 10 | 2.9270% | 2.82286% | 2.863% | 2.78778% | 2.827% |
| 15 | 3.1430% | 3.06008% | 3.107% | 2.99762% | 3.043% |
| 20 (FSP) | 3.2330% | 3.15865% | 3.209% | 3.08492% | 3.133% |
| 25 | 3.2500% | 3.16758% | 3.224% | 3.10141% | 3.157% |
| 30 | 3.2431% | 3.14607% | 3.235% | 3.09472% | 3.176% |
| 40 | 3.2060% | 3.07181% | 3.250% | 3.05874% | 3.204% |
| 50 | 3.1310% | 2.92567% | 3.260% | 2.98597% | 3.223% |
| 60 | | | 3.266% | | 3.235% |
| 100 | | | 3.280% | | 3.261% |
| 150 | | | 3.287% | | 3.274% |

- **Government bonds.** Up to the FSP the basic RFR is the input rate minus the CRA, for example
  2.569% − 0.10% = 2.469% at 5 years. The bootstrapped rate is the same rate, continuously
  compounded.
- **Swaps.** Only at 1 year is the result the input rate minus the CRA, because there the par rate
  and the zero rate coincide. For this rising curve the zero rates lie above the par rates, so the
  swap curve is higher than the government bond curve.
- **Beyond the FSP** both curves follow the extrapolation from their own LLFR towards the UFR. The
  market rates at 25, 30, 40 and 50 years enter only through the LLFR.

---

## Using the algorithm in your own code

`Code/calculation.py` contains the whole calculation and imports only `math`, `decimal` and
`typing`. Copy it into your project, or add the `Code` folder to the import path.

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

For government bonds, read the same rates as annually compounded zero rates:

```python
inputs["Instrument"] = "GVT"           # government bonds
inputs["CouponFreq"] = None            # not used for government bonds
validate_inputs(inputs)
res = calculate_sheet(**inputs)

print(res["FSP"], res["LLFR"])         # 20 0.03188302797422029
print(res["BASIC_RFR"][60])            # 0.03235
```

### Inputs of `calculate_sheet`

`calculate_sheet` takes the inputs as keyword arguments; together they form the dict type
`SheetInputs`, which `data_io.read_inputs` returns.

| Name | Type | Meaning | Sheet | File |
|---|---|---|---|---|
| `Instrument` | `str` | `"SWP"` for swaps, `"GVT"` for government bonds | H11 | `parameters.csv` `Instrument` |
| `CouponFreq` | `int` or `None` | Coupon payments per year of the swaps; `None` for government bonds | H12 | `Coupon Frequency` |
| `CRA` | `float` | Credit risk adjustment in basis points | H16 | `CRA` |
| `UFR` | `float` | Ultimate forward rate, annually compounded | H14 | `UFR` |
| `alpha` | `float` | Convergence parameter α | H15 | `Convergence` |
| `dlt` | `dict[int, int]` | `{maturity: 1 or 0}`; missing maturities count as 0 | H25:H174 | `curve.csv` `DLT` |
| `LLFRweightsIn` | `dict[int, float]` | `{maturity: LLFR weight}`; missing maturities count as 0 | I25:I174 | `LLFR Weight` |
| `Rate` | `dict[int, float]` | `{maturity: market rate}` before the CRA: par swap rates, or annually compounded zero rates for government bonds; a maturity without a rate is left out | J25:J174 | `Input Rate` |
| `MAX_MATURITY` | `int` | Longest maturity of the curves, in years | 150 rows | `Max Maturity` |

### Outputs of `calculate_sheet`

`calculate_sheet` returns a dict of type `SheetResults`:

| Key | Type | Meaning | Sheet | Output file |
|---|---|---|---|---|
| `FSP` | `int` | First smoothing point in years | H13 | `*_parameters.csv` `FSP` |
| `LLFR` | `float` | Last liquid forward rate, continuously compounded | L22 | `LLFR (CC)` |
| `BOOTSTRAPPED_CURVE` | `dict[int, float]` | `{maturity: zero rate}` for 1 to `MAX_MATURITY`, continuously compounded, after the CRA; `nan` after the last DLT maturity | K25:K174 | `*_curves.csv` `Bootstrapped Zero Rate CC` |
| `BASIC_RFR` | `dict[int, float]` | `{maturity: basic risk-free rate}` for 1 to `MAX_MATURITY`, annually compounded, rounded to 5 decimals | L25:L174 | `Basic RFR` |

### Only the extrapolation

If you already have a continuously compounded zero curve for every year up to the FSP, for
example from your own bootstrapping, you can use the LLFR and extrapolation steps on their own:

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

### Function reference

Inputs and outputs of every function you might call. A curve is a dict `{maturity: rate}`; the
docstrings give the details and the errors raised.

| Function | Module | Workbook | Inputs | Returns |
|---|---|---|---|---|
| `calculate_sheet` | `calculation` | Sheet *Input Data & Extrapolation* | The `SheetInputs` above | `SheetResults` (above) |
| `bootstrap_curve` | `calculation` | VBA `BootstrapCurve` | `Instrument`, `Rate`, `dlt`, `CompoundingIn`, `CRA`, `CouponFreq`, `MaxTenor`, `CompoundingOut`, `RateType`, `MAX_MATURITY` | Zero curve for 1 to `MAX_MATURITY`, `nan` after `MaxTenor`; from `bootstrap_swaps` for `"SWP"`, otherwise from `bootstrap_zeros` |
| `bootstrap_swaps` | `calculation` | VBA `BootstrapSwaps` | `SwapRatesInit` (par swap rates), `dlt`, `CouponFreq`, `CRA`, `MaxTenor`, `CompoundingOut`, `RateType`, `MAX_MATURITY` | Zero curve, as above |
| `bootstrap_zeros` | `calculation` | VBA `BootstrapZeros` | `ZeroRatesInit` (zero rates), `dlt`, `CompoundingIn`, `CRA`, `MaxTenor`, `CompoundingOut`, `RateType`, `MAX_MATURITY` | Zero curve, as above |
| `newton_raphson_forward_swap` | `calculation` | VBA `NewtonRaphsonForwardSwap` | `fwguess`, `swapt1` (periodic coupon), `m` (periods), `c` (target), `tol` = 1e-15, `max_iter` = 500 | Constant periodic forward rate of one swap interval |
| `get_llfr` | `calculation` | VBA `GetLLFR` | `InputRates` (continuously compounded zero curve), `DLTin`, `LLFRweightsIn` | LLFR, continuously compounded |
| `extrapolation` | `calculation` | VBA `Extrapolation` | `InputRatesCC`, `RateType`, `FSP`, `UFR`, `LLFR`, `alpha`, `Compounding`, `MAX_MATURITY` | Zero curve for 1 to `MAX_MATURITY`: the input up to the FSP, extrapolated after it |
| `excel_round` | `calculation` | Worksheet `ROUND(…, 5)` | `curve`, `digits` = 5 | The curve with every rate rounded by `excel_round_value` |
| `excel_round_value` | `calculation` | Worksheet `ROUND` | `value`, `digits` = 5 | One number rounded as Excel does: at 15 significant digits, halves away from zero |
| `validate_inputs` | `validation` | | `SheetInputs` | Nothing; raises `ValueError` listing every problem |
| `read_inputs` | `data_io` | | `folder` (default `Input/`) | `SheetInputs` read from `parameters.csv` and `curve.csv` |
| `save_results` | `data_io` | | `SheetInputs`, `SheetResults`, `folder` (default `Output/`) | Paths of the two result files written |
| `run` | `main` | | `input_dir`, `output_dir` (default `Input/`, `Output/`) | `SheetResults`; also prints a summary and saves the results |

The text options take these values, as in the VBA:

| Option | Value | Meaning |
|---|---|---|
| `CompoundingIn` | `"A"` / `"C"` | Input zero rates are annually / continuously compounded (`bootstrap_zeros` only) |
| `CompoundingOut`, `Compounding` | `"A"` / `"C"` | Return annually / continuously compounded rates |
| `RateType` | `"Z"` / `"F"` | `extrapolation`: the input are zero rates / one-year forward rates. The bootstrap functions ignore it and always return zero rates. |

`calculate_sheet` calls the bootstrap with `CompoundingIn = "A"`, `CompoundingOut = "C"` and
`MaxTenor` = the last DLT maturity, and the extrapolation with `RateType = "Z"` and
`Compounding = "A"`, exactly like the formulas on the sheet.

### Conventions

- Maturities are whole years. Curves, flags and weights are dicts keyed by maturity, e.g.
  `{1: 0.02176, 2: 0.022621}`.
- Rates are decimals (0.033 = 3.3%). The CRA is in basis points.
- `nan` plays the role of Excel's `#N/A`.
- Names follow the VBA in EIOPA's workbook to make the two easy to compare: `bootstrap_curve` is
  `BootstrapCurve`, `bootstrap_swaps` is `BootstrapSwaps`, `bootstrap_zeros` is `BootstrapZeros`,
  `get_llfr` is `GetLLFR`, and `Instrument`, `Rate`, `CouponFreq`, `LLFRweightsIn`, `MaxTenor`
  and so on keep their VBA names. Every function's docstring describes its inputs,
  outputs and errors, with references to the documentation.

---

## Verification

- **Against EIOPA's workbook.** The results were compared with *RFR extrapolation and VA
  calculation (19 May 2026).xlsm*, recalculated by Excel and its VBA. The FSP, the LLFR, the
  bootstrapped curve and the basic risk-free rate match exactly, with a largest difference of 0,
  in four cases:
  - the workbook's example: annual coupons, CRA 10 bp, UFR 3.3%, α 11%;
  - a variant with quarterly coupons, CRA 30 bp, UFR 4.3%, α 40% and a 1-year swap rate of 2.5%;
  - the workbook's example with instrument type `GVT` (government bonds);
  - government bond rates with six decimals ending in 5, where the rounding to 5 decimals decides
    the last digit (see [rounding](#differences-from-the-eiopa-workbook-and-documentation)).
- **Unit tests** cover the calculation, the validation, reading and writing, and the main script,
  with full line and branch coverage of the calculation, validation and input/output modules.
  Several tests check mathematical properties rather than stored numbers: every DLT swap reprices
  at par, every DLT bond rate is reproduced, a flat par curve gives a flat zero curve, forwards are
  constant between DLT points, and the extrapolated forward follows the formula and converges to
  the UFR.

Run the tests from the repository folder with either:

```bash
python -m unittest discover -s Code
```

```bash
python -m pytest Code
```

To check the example against the workbook yourself, compare `Output/swaps_curves.csv`
with the values the workbook in `Sources/` stores for cells H13, L22, K25:K174 and L25:L174 of the
sheet *Input Data & Extrapolation*. The workbook's inputs are the same as those in `Input/`. For
`Output/government_bonds_curves.csv`, set the workbook's Instrument Type (H11) to `GVT` first.

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

---

## Scope and limitations

Covered: the basic risk-free rate term structure from **swaps** or **government bonds**, i.e. the
sheet *Input Data & Extrapolation* of EIOPA's workbook with either instrument type.

Not covered:

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
Input/                 curve.csv, parameters.csv
Output/                swaps_*.csv and government_bonds_*.csv (curves and parameters)
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
