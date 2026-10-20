# EIOPA risk-free rate extrapolation in Python

A transparent Python implementation of EIOPA's new method for deriving the Solvency II basic
risk-free interest rate term structure from swap rates or government bond zero rates:
constant-forward bootstrapping up to the first smoothing point, followed by extrapolation of
forward rates from the last liquid forward rate (LLFR) to the ultimate forward rate (UFR). It also
calculates the undertaking-specific volatility adjustment (VA): the credit spread sensitivity ratio
(CSSR), the VA and the term structure with VA; and the term structures after the upward and
downward interest rate shocks, without and with VA.

It is written for actuaries who want to

- **understand the change**: the method, its formulas and how it differs from what EIOPA's
  documents and Excel tool show, step by step;
- **use the algorithm in their own code**: `Code/calculation.py` depends only on the Python
  standard library and can be copied into another project as it is.

The basic risk-free rate is a line-by-line port of the VBA in EIOPA's demonstration workbook, and
the VA and the interest rate shocks are ports of the workbook's worksheet formulas, for which there
is no VBA. All three reproduce the workbook exactly (see [Verification](#verification)).

> **Disclaimer.** This is not an EIOPA product and is not endorsed by EIOPA. It is provided for
> educational purposes, without warranty. EIOPA publishes the official risk-free rate term
> structures; check any result against EIOPA's publications before relying on it.

---

## Contents

- [What changes](#what-changes)
- [The method](#the-method)
- [Quick start](#quick-start)
- [Inputs and outputs](#inputs-and-outputs): files and columns
- [Using the algorithm in your own code](#using-the-algorithm-in-your-own-code): inputs and outputs of
  [`calculate_sheet`](#inputs-of-calculate_sheet), [the volatility adjustment](#the-volatility-adjustment),
  [the interest rate shocks](#the-interest-rate-shocks)
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

The review also changes the volatility adjustment (VA), which becomes undertaking-specific. EIOPA
no longer publishes a VA for each currency, only the risk-corrected spreads it is built from. Each
undertaking that applies the VA calculates its own credit spread sensitivity ratio (CSSR) from its
fixed income assets and best estimate liabilities, and from that its own VA and term structure with
VA (see [step 5 of the method](#5-volatility-adjustment-93-94-10)).

EIOPA's workbook also derives the term structures for the interest rate risk sub-module of the
standard formula with the new method: the bootstrapped rates are shocked up and down, and the
shocked curves are extrapolated from their own LLFR (see
[step 6 of the method](#6-interest-rate-risk-shocks-sheets-irr-up-shock-and-irr-down-shock)).

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
chooses between it and `BootstrapSwaps` from the instrument type.

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

### 5. Volatility adjustment (§9.3, §9.4, §10)

EIOPA publishes, for each currency, the risk-corrected currency spread $RCS_{cu}$ and, for
each euro area country, the risk-corrected country spread $RCS_{co}$ and the
debt-to-total-assets ratio $DTAR_{co}$. The undertaking adds its own sensitivity.

**Curve with a VA (§10).** A VA is added to the one-year forward rates of the basic RFR, and the
first term of the LLFR is adjusted in the same way:

$$f^{VA,c}_{t-1,t} = f^c_{t-1,t} + \ln(1 + VA), \qquad LLFR^{VA,c} = LLFR^c + w_{t_F} \cdot \ln(1 + VA).$$

Up to the FSP the zero rates are the averages of these forwards; beyond it the curve is
extrapolated as in step 4, from $LLFR^{VA,c}$. Like the workbook, the code takes the forwards
from the basic RFR as published, i.e. rounded to 5 decimals.

**CSSR (§9.3.2).** The *CSSR curve* is the curve with a notional VA of $85\% \cdot \text{RCS}_{cu}$.
The undertaking values its fixed income assets ($FI$) and its best estimate liabilities ($BEL$) on
the basic RFR and on the CSSR curve:

$$\text{CSSR} = \max\left(\min\left(\frac{\Delta FI}{\Delta BEL}, 1\right), 0\right), \qquad \Delta FI = FI_{CSSR} - FI_{MV}, \quad \Delta BEL = BEL_{CSSR} - BEL_{MV}.$$

The CSSR compares how strongly the assets and the liabilities react to the same change in spreads:
the VA is reduced to the extent that the assets react less than the liabilities.

**VA (§9.3.1, §9.3.5, §9.4).** The currency VA is

$$\text{VA}_{cu} = 85\% \cdot \text{CSSR} \cdot \text{RCS}_{cu} .$$

An undertaking that meets the requirements of Article 77d(1c) may use an undertaking-specific
adjustment (USA): $\text{RCS}_{cu}$ is then multiplied by the USA, the ratio of the RCS of its own
bond portfolio to EIOPA's, at most 105%. Euro area undertakings that do not use a USA add the
macro VA of their country (§9.3.4):

$$\text{VA}_{co,macro} = 85\% \cdot \text{CSSR} \cdot \max(\text{RCS}_{co} - 1.3 \cdot \text{RCS}_{euro}, 0) \cdot \omega_{co}, \qquad \omega_{co} = \max\left(\min\left(\frac{\text{RCS}_{co} \cdot \text{DTAR}_{co} - 0.6\%}{0.3\%}, 1\right), 0\right).$$

The total VA is rounded to whole basis points (§10.1.3) and gives the term structure with VA. It can
be negative (§9.3.6).

### 6. Interest rate risk shocks (sheets IRR Up Shock and IRR Down Shock)

The technical documentation does not describe the shocks; this step follows the workbook's
worksheet formulas. For each maturity the sheet *Parameters* gives a relative shock $s$ and an
absolute shock $b$, upward and downward, and a floor $r^{floor}$ for the downward shock.

**Shocked rates.** The shocks apply to the bootstrapped zero rates $r_t$ after the CRA, annually
compounded, for every maturity up to the last DLT maturity:

$$r^{up}_t = r_t \cdot (1 + s^{up}_t) + b^{up}_t, \qquad r^{down}_t = r_t \cdot (1 - s^{down}_t) - b^{down}_t .$$

Annual compounding matters: the formula gives a different result on continuously compounded rates,
because $\ln(1 + r(1+s) + b) \ne (1+s)\ln(1+r) + b$. At 1 year in the workbook's example the upward
shock gives 5.482%; applied to the continuously compounded rate it would give 5.599%. There is no
special case for low or negative rates.

**Extrapolation.** The shocked rates are converted to continuous compounding,
$z^{c,up}_t = \ln(1 + r^{up}_t)$, and treated like the bootstrapped curve in steps 3 and 4: the LLFR
is calculated from them with the same DLT maturities and weights, and the curve is extrapolated from
that LLFR with the UFR shifted by 15 basis points, $\text{UFR} + 0.15\%$ upwards and
$\text{UFR} - 0.15\%$ downwards. Up to the FSP the curve is the shocked rates; between the FSP and
the last DLT maturity the shocked rates enter only the LLFR. The results are rounded to 5 decimals.

**Floor.** After the downward shock, each rounded rate is raised to at least the floor of its
maturity: $\max(z^{down}_t, r^{floor}_t)$. The floor is applied last, so it does not change the LLFR.
In the workbook's table it is −1.25% up to 7 years, rises linearly to −0.893% at 20 years and stays
there. The upward shock has no floor.

**With VA.** As in step 5, but starting from the shocked rates: the VA is added to the one-year
forwards of $z^{c,up}$ (or $z^{c,down}$) and the LLFR becomes
$LLFR^{up,c} + w_{t_F} \cdot \ln(1 + VA)$. The VA is the undertaking's total VA, not recalculated
for the shock. Unlike on the sheet *VA Term Structures*, the forwards come from the unrounded
shocked rates. The floor applies to the downward curve with VA as well.

---

## Quick start

Requirements: Python 3.10 or later.

```bash
python Code/main.py
```

This first deletes the folder `Output/`, so that it only ever holds the results of the latest run.
It then runs three steps:

1. **Basic risk-free rate** (`run_basic_rfr` in `Code/main.py`): reads `curve.csv` and
   `parameters.csv` from `Input/`, validates them, calculates the curves, prints a summary and
   writes the results to `Output/`.
2. **Volatility adjustment** (`run_va`), only if `Input/` contains `va_parameters.csv`: reads and
   validates it, and calculates, prints and writes the VA on top of the basic risk-free rate.
3. **Interest rate shocks** (`run_irr`), only if `Input/` contains `irr_parameters.csv`: reads and
   validates it, and calculates, prints and writes the curves after the upward and downward
   shocks. With a total VA from the second step, also the shocked curves with VA.

The first step prints:

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

Because `Input/` also contains `va_parameters.csv`, the second step follows, here with the
workbook's example values:

```text
Notional VA : 85 bp
CSSR        : 0.500000
Currency VA : 34.00 bp
Total VA    : 34 bp
Maturity  Basic RFR  CSSR Curve  VA Curve
       1     2.076%      2.944%    2.423%
      ...
      20     3.209%      4.086%    3.560%
     150     3.287%      3.421%    3.340%
```

Calculating your own VA takes two runs, because the CSSR depends on a valuation on the CSSR curve:

1. Fill in the currency RCS (and the other VA settings) in `va_parameters.csv` and leave the four
   values for the CSSR empty. The run writes the CSSR curve to `Output/swaps_va_curves.csv` and
   prints `CSSR : not calculated`.
2. Value your fixed income assets and your best estimate on the basic RFR and on the CSSR curve,
   enter the four values in `va_parameters.csv` and run again. The run now calculates the CSSR,
   the VA and the curve with VA.

Without `va_parameters.csv` the second step does not run.

For the third step, copy `Examples/EIOPA_Example/Input/irr_parameters.csv`, the shock parameters
of the workbook's sheet *Parameters*, into `Input/`. With the VA of 34 bp above it prints:

```text
UFR up      : 3.45%
UFR down    : 3.15%
LLFR up     : 4.04702331% (continuously compounded)
LLFR down   : 1.01223960% (continuously compounded)
VA          : 34 bp
LLFR* up    : 4.15903300% (with VA, continuously compounded)
LLFR* down  : 1.12424929% (with VA, continuously compounded)
Maturity  Basic RFR       Up  Up + VA     Down Down + VA
       1     2.076%   5.482%   5.841%  -0.288%    0.051%
      ...
      20     3.209%   4.891%   5.248%   1.105%    1.448%
     150     3.287%   3.682%   3.736%   2.745%    2.798%
```

Without a total VA (no `va_parameters.csv`, or the CSSR not yet calculated) it prints
`VA : not available` and calculates the shocked curves without VA only.

To calculate the curve from government bonds instead, set `Instrument` to `GVT` in
`Input/parameters.csv` and put the bond zero rates in `Input/curve.csv`. The run then prints
`Instrument: GVT (government bonds)` and writes `government_bonds_curves.csv` and
`government_bonds_parameters.csv`.

Invalid inputs stop the run with a message listing every problem, for example:

```text
Input error: invalid inputs:
  - UFR must be a decimal, e.g. 0.033 for 3.3% (is 3.3)
  - LLFR weights must sum to 1 (sum is 1.0999999999999999)
```

Each step checks its own inputs. Invalid inputs for the basic risk-free rate stop the run before
anything is written; `Output/` has already been deleted and is not created again. Invalid VA inputs (`Input error: invalid VA
inputs: ...`) stop it after the first step: the basic risk-free rate is saved, the VA is not.
Invalid shock parameters (`Input error: invalid interest rate shock inputs: ...`) stop it after
the second step.

---

## Inputs and outputs

`python Code/main.py` reads two files from `Input/` and writes two files to `Output/`; with the
optional `va_parameters.csv` it writes two more for the VA, and with the optional
`irr_parameters.csv` two more for the interest rate shocks. `Output/` is deleted at the start of
every run, so copy results you want to keep elsewhere. The instrument in `parameters.csv`
decides how the rates in `curve.csv` are read and how the result files are named.

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

### Input file `Input/va_parameters.csv` (optional)

Columns `Parameter, Value`, like `parameters.csv`. Only `Currency RCS` is required; an empty value
counts as missing.

| Parameter | Meaning | Format | Required | Sheet cell | Example |
|---|---|---|---|---|---|
| `Currency RCS` | Risk-corrected currency spread published by EIOPA | Decimal (0.01 = 100 bp) | Always | Input Data & Extrapolation L10 | `0.01` |
| `Country RCS` | Risk-corrected country spread published by EIOPA | Decimal | Macro VA only | Input Data & Extrapolation L11 | `0.014` |
| `Debt to Total Assets` | Debt-to-total-assets ratio of the country published by EIOPA | Decimal, 0 to 1 | Macro VA only | Input Data & Extrapolation L12 | `0.5` |
| `Use Macro VA` | 1 to add the macro VA of the country | `0` or `1`; empty counts as 0 | No | VA Term Structures H10 | `0` |
| `USA` | Undertaking-specific adjustment; empty if not used | Decimal, at most 1.05 | No | VA Term Structures H11:H12 | `0.8` |
| `MV Fixed Income` | Fixed income assets valued on the basic RFR | Amount | For the CSSR | CSSR Calculation L10 | `500000` |
| `MV Fixed Income CSSR Curve` | The same assets valued on the CSSR curve | Amount | For the CSSR | CSSR Calculation L11 | `490000` |
| `BEL` | Best estimate liabilities valued on the basic RFR | Amount | For the CSSR | CSSR Calculation L13 | `870000` |
| `BEL CSSR Curve` | The same liabilities valued on the CSSR curve | Amount | For the CSSR | CSSR Calculation L14 | `850000` |

The four values for the CSSR are given all together or not at all. The USA and the macro VA cannot
be used together.

### Input file `Input/irr_parameters.csv` (optional)

The shock parameters of the sheet *Parameters* (G16:L166), one row per maturity. The rows must
cover every maturity 1 to `Max Maturity`; rows beyond it are not used, so the table of 150 years
also serves shorter curves. `Examples/EIOPA_Example/Input/irr_parameters.csv` holds the workbook's
table, including the values the workbook interpolates (relative shocks from 50 to 90 years,
absolute shocks from 50 to 60 years, the floor from 7 to 20 years), at full precision.

| Column | Meaning | Format | Sheet column | Example (1 year) |
|---|---|---|---|---|
| `Maturity` | Maturity | Whole number of years ≥ 1 | G | `1` |
| `smup` | Relative upward shock $s^{up}$ | Decimal, 0 to 1 | H | `0.61` |
| `bmup` | Absolute upward shock $b^{up}$ | Decimal, 0 to 1 | I | `0.0214` |
| `smdown` | Relative downward shock $s^{down}$ | Decimal, 0 to 1 | J | `0.58` |
| `bmdown` | Absolute downward shock $b^{down}$ | Decimal, 0 to 1 | K | `0.0116` |
| `rmfloor` | Floor of the rates after the downward shock | Decimal | L | `-0.0125` |

All four files may use `,` or `;` as separator and `.` or `,` as decimal mark, so files saved by
Excel with European settings are read correctly.

The DLT maturities, the FSP, the LLFR weights, the UFR, the risk-corrected spreads and the shock
parameters are not calculated here. They come from EIOPA's DLT assessment, UFR methodology and
representative portfolios, published with the risk-free rates, and from the standard formula.

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

Columns `Parameter, Value`, in the order of the sheet *Input Data & Extrapolation*:

| Parameter | Meaning | Sheet cell |
|---|---|---|
| `Instrument` | `SWP` or `GVT` | H11 |
| `Coupon Frequency` | Coupon payments per year; swaps only | H12 |
| `FSP` | First smoothing point in years: the first maturity with a positive LLFR weight | H13 |
| `UFR`, `Convergence`, `CRA` | The parameters used | H14:H16 |
| `Max Maturity` | Longest maturity of the curves | 150 rows on the sheet |
| `LLFR (CC)` | Last liquid forward rate, calculated from the bootstrapped curve and the LLFR weights, continuously compounded | L22 |

### Output file `Output/swaps_va_curves.csv` or `Output/government_bonds_va_curves.csv`

Written only with `va_parameters.csv`. One row per maturity 1 to `Max Maturity`, all annually
compounded and rounded to 5 decimals.

| Column | Meaning | Sheet column |
|---|---|---|
| `Maturity` | Maturity in years | CSSR Calculation G |
| `Basic RFR` | Basic risk-free interest rate term structure | CSSR Calculation H |
| `CSSR Curve` | Term structure with the notional VA, on which the assets and liabilities are valued for the CSSR | CSSR Calculation L |
| `VA Curve` | Term structure with the total VA; empty until the CSSR can be calculated | VA Term Structures L |

### Output file `Output/swaps_va_parameters.csv` or `Output/government_bonds_va_parameters.csv`

Columns `Parameter, Value`: the inputs of `va_parameters.csv` and the calculated values, in the
order of the workbook, sheet by sheet and each block from top to bottom. Values not yet available
are empty.

| Parameter | Meaning | Sheet cell |
|---|---|---|
| `Currency RCS`, `Country RCS`, `Debt to Total Assets` | The inputs | Input Data & Extrapolation L10:L12 |
| `Notional VA` | 85% of the currency RCS, rounded to whole basis points | CSSR Calculation H10 |
| `MV Fixed Income`, `MV Fixed Income CSSR Curve` | The inputs | CSSR Calculation L10:L11 |
| `PVBP MV Fixed Income` | Change in value of the assets on the CSSR curve divided by the notional VA | CSSR Calculation L12 |
| `BEL`, `BEL CSSR Curve` | The inputs | CSSR Calculation L13:L14 |
| `PVBP BEL` | Change in value of the best estimate on the CSSR curve divided by the notional VA | CSSR Calculation L15 |
| `CSSR` | Credit spread sensitivity ratio | CSSR Calculation L16 |
| `Notional VA LLFR (CC)` | LLFR of the CSSR curve, continuously compounded | CSSR Calculation K20 |
| `Use Macro VA`, `USA` | The inputs; an empty `USA` stands for USA? = 0 | VA Term Structures H10:H12 |
| `Currency VA` | Currency VA, with the USA if given | VA Term Structures L10 |
| `Country Adjustment Factor` | $\omega_{co}$ of the macro VA; empty without the country values | VA Term Structures L11 |
| `Macro VA` | Macro VA; empty if not used (Excel's `#N/A`) | VA Term Structures L12 |
| `Total VA` | Total VA, rounded to whole basis points | VA Term Structures L13 |
| `Total VA LLFR (CC)` | LLFR of the curve with VA, continuously compounded | VA Term Structures K20 |

### Output file `Output/swaps_irr_curves.csv` or `Output/government_bonds_irr_curves.csv`

Written only with `irr_parameters.csv`. One row per maturity 1 to `Max Maturity`, all annually
compounded. Columns from the sheets *IRR Up Shock* (Up) and *IRR Down Shock* (Down):

| Column | Meaning | Sheet column |
|---|---|---|
| `Maturity` | Maturity in years | G |
| `Basic RFR` | Basic risk-free interest rate term structure | Input Data & Extrapolation L |
| `Bootstrapped Zero Rate AC` | Zero rate after the CRA, bootstrapped up to the last DLT maturity; empty after it | I |
| `IRR Shock Up AC` | Bootstrapped rate after the upward shock; empty after the last DLT maturity | Up J |
| `IRR Up RFR` | Term structure after the upward shock, rounded to 5 decimals | Up L |
| `IRR Up with VA` | The same with the total VA; empty without a VA | Up O |
| `IRR Shock Down AC` | Bootstrapped rate after the downward shock; empty after the last DLT maturity | Down J |
| `IRR Down RFR` | Term structure after the downward shock, rounded to 5 decimals, at least the floor | Down L |
| `IRR Down with VA` | The same with the total VA; empty without a VA | Down O |

### Output file `Output/swaps_irr_parameters.csv` or `Output/government_bonds_irr_parameters.csv`

Columns `Parameter, Value`, in the order of the workbook. Values not available without a VA are
empty.

| Parameter | Meaning | Sheet cell |
|---|---|---|
| `UFR Up` | UFR + 0.15% | IRR Up Shock H11 |
| `VA` | Total VA used for the curves with VA | IRR Up Shock H12 (and IRR Down Shock H12) |
| `LLFR IRR Up (CC)` | LLFR of the upward shocked rates, continuously compounded | IRR Up Shock L19 |
| `LLFR* IRR Up (CC)` | The same with the VA | IRR Up Shock O19 |
| `UFR Down` | UFR − 0.15% | IRR Down Shock H11 |
| `LLFR IRR Down (CC)` | LLFR of the downward shocked rates, continuously compounded | IRR Down Shock K19 |
| `LLFR* IRR Down (CC)` | The same with the VA | IRR Down Shock N19 |

All rates are decimals written with full precision.

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

### The volatility adjustment

`Code/volatility_adjustment.py` builds on `calculation.py` and imports only `math` and `typing`
besides it. With `inputs` and `res` of the swap example in
[The complete calculation](#the-complete-calculation):

```python
from validation import validate_va_inputs
from volatility_adjustment import VAInputs, calculate_cssr_sheet, calculate_va_sheet

va = VAInputs(
    RCS_CU=0.01,                       # currency RCS published by EIOPA (100 bp)
    RCS_CO=None, DTTA=None,            # country RCS and debt-to-total-assets ratio: macro VA only
    MACROVA_SELECTED=False,
    USA=None,                          # undertaking-specific adjustment, e.g. 0.8; None if not used
    MVFI=None, MVFI_NOTIONAL=None,     # fixed income assets on the basic RFR and on the CSSR curve
    BEL=None, BEL_NOTIONAL=None,       # best estimate on the basic RFR and on the CSSR curve
)
validate_va_inputs(va)
cssr = calculate_cssr_sheet(inputs, res, va)
cssr_curve = cssr["CSSR_CURVE"]        # value your assets and liabilities on this curve

va.update(MVFI=500000, MVFI_NOTIONAL=490000, BEL=870000, BEL_NOTIONAL=850000)
cssr = calculate_cssr_sheet(inputs, res, va)
print(cssr["CSSR"])                    # 0.5
va_res = calculate_va_sheet(inputs, res, va, cssr["CSSR"])
print(va_res["TOTAL_VA"])              # 0.0043  (42.5 bp, rounded to whole basis points)
print(va_res["VA_CURVE"][20])          # 0.03653
```

`calculate_cssr_sheet` returns `NOTIONAL_VA`, `NOTIONAL_VA_LLFR`, `CSSR_CURVE` and, once the four
values are given, `PVBP_MVFI`, `PVBP_BEL` and `CSSR` (otherwise `None`). `calculate_va_sheet`
returns `CURRENCY_VA`, `W_CO`, `MACRO_VA` (`nan` if not used), `TOTAL_VA`, `TOTAL_VA_LLFR` and
`VA_CURVE`. The building blocks are available on their own as well: `va_curve` adds any VA to a
basic RFR, `notional_va`, `credit_spread_sensitivity_ratio` and `country_adjustment_factor`
calculate the single values.

### The interest rate shocks

`Code/interest_rate_risk.py` builds on `calculation.py` and imports only `math` and `typing`
besides it. With `inputs` and `res` of the swap example in
[The complete calculation](#the-complete-calculation):

```python
from pathlib import Path

from data_io import read_irr_inputs
from interest_rate_risk import calculate_irr_down_sheet, calculate_irr_up_sheet
from validation import validate_irr_inputs

irr = read_irr_inputs(Path("Examples/EIOPA_Example/Input"))   # the workbook's shock parameters
validate_irr_inputs(irr, inputs["MAX_MATURITY"])
up = calculate_irr_up_sheet(inputs, res, irr)                 # without VA
down = calculate_irr_down_sheet(inputs, res, irr, TOTAL_VA=0.0034)
print(up["LLFR"], up["IRR_RFR"][60])         # 0.040470233120777376 0.0403
print(down["IRR_RFR"][60])                   # 0.02144
print(down["IRR_VA_CURVE"][20])              # 0.01448  (with the VA of 34 bp)
```

`IRRInputs` holds the five columns of the parameter table as `{maturity: value}`: `IRR_SM_UP`,
`IRR_BM_UP`, `IRR_SM_DOWN`, `IRR_BM_DOWN` and `RM_FLOOR`. Both functions return an `IRRResults`
with `UFR`, `VA`, `LLFR`, `LLFR_STAR` (with VA), the columns `BOOTSTRAPPED_CURVE_AC`,
`IRR_SHOCK_AC` and `IRR_SHOCK_CC` (`nan` after the last DLT maturity), and the curves `IRR_RFR` and
`IRR_VA_CURVE`. Without `TOTAL_VA`, `VA`, `LLFR_STAR` and `IRR_VA_CURVE` are `nan`. The building
blocks `shock_up`, `shock_down`, `apply_floor`, `bootstrapped_curve_ac` and
`shocked_term_structures` are available on their own as well.

### Conventions

- Maturities are whole years. Curves, flags and weights are dicts keyed by maturity, e.g.
  `{1: 0.02176, 2: 0.022621}`.
- Rates are decimals (0.033 = 3.3%). The CRA is in basis points.
- `nan` plays the role of Excel's `#N/A`.
- Names follow the VBA in EIOPA's workbook to make the two easy to compare: `bootstrap_curve` is
  `BootstrapCurve`, `bootstrap_swaps` is `BootstrapSwaps`, `bootstrap_zeros` is `BootstrapZeros`,
  `get_llfr` is `GetLLFR`, and `Instrument`, `Rate`, `CouponFreq`, `LLFRweightsIn`, `MaxTenor`
  and so on keep their VBA names. The VA and the shocks have no VBA, so their names follow the
  workbook's named ranges without the `R.`/`V.` prefix: `RCS_CU`, `RCS_CO`, `DTTA`, `MVFI`,
  `MVFI_NOTIONAL`, `BEL_NOTIONAL`, `NOTIONAL_VA`, `CSSR`, `TOTAL_VA`, `W_CO`, `IRR_SM_UP`,
  `RM_FLOOR` and so on. Every function's docstring
  describes its inputs, outputs and errors, with references to the documentation.

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
    the last digit.

  For the VA, the sheets *CSSR Calculation* and *VA Term Structures* of the workbook's example
  (currency RCS 100 bp, CSSR 0.5, USA 80%) match exactly as well: the notional VA, the CSSR and
  its two intermediate values, the currency VA, the country adjustment factor, the total VA, both
  LLFRs, and the CSSR curve and the curve with VA at all 150 maturities. The macro VA, switched
  off in the workbook's example, is tested against the formula of §9.4.

  For the shocks, the sheets *IRR Up Shock* and *IRR Down Shock* of the workbook's example
  (saved with instrument type `GVT`, total VA 34 bp) match exactly: the shifted UFRs, the LLFRs
  with and without VA, and the bootstrapped rates, the shocked rates and both shocked curves at
  all 150 maturities (`Code/tests/irr_workbook_values.csv` holds the workbook's values). The floor
  does not bind in the workbook's example; it is tested on a curve with negative rates.
- **Unit tests** cover the calculation, the VA, the shocks, the validation, reading and writing,
  and the main script, with full line and branch coverage of every module (the main script's
  command-line part is run in a separate process).
  Several tests check mathematical properties rather than stored numbers: every DLT swap reprices
  at par, every DLT bond rate is reproduced, a flat par curve gives a flat zero curve, forwards are
  constant between DLT points, and the extrapolated forward follows the formula and converges to
  the UFR. For the VA: the forwards up to the FSP rise by the VA, the LLFR by the first weight
  times the VA, a VA of 0 gives back the basic RFR, and the effect of the VA fades towards the UFR.
  For the shocks: shocks of 0 give back the basic RFR up to the FSP and its LLFR, the floor leaves
  the LLFRs unchanged, and a VA of 0 gives back the shocked curve without VA.

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
For the VA, compare `Output/swaps_va_curves.csv` and `Output/swaps_va_parameters.csv` with the
sheets *CSSR Calculation* and *VA Term Structures*; `Input/va_parameters.csv` holds the workbook's
VA inputs. For the shocks, copy the files of `Examples/EIOPA_Example/Input/` (government bonds,
with the VA inputs and `irr_parameters.csv`) into `Input/`, run, and compare
`Output/government_bonds_irr_curves.csv` and `Output/government_bonds_irr_parameters.csv` with
the sheets *IRR Up Shock* and *IRR Down Shock* as the workbook is saved.

---

## Differences from the EIOPA workbook and documentation

For the basic risk-free rate, none of these change the results.

| Topic | This code and the VBA | EIOPA documentation |
|---|---|---|
| Solver | Newton-Raphson, tolerance 1e-15, at most 500 iterations | EIOPA's production code uses MATLAB `fzero` (D.5.12) |
| Newton starting guess | As the VBA: `forward(t - 1) / CouponFreq`, where `t` counts DLT points, not years | The periodic forward of the previous interval (D.6.3). Only the starting point differs, not the solution. |
| First LLFR forward | From the last DLT maturity before the FSP to the FSP | One-year forward into the FSP (8.5.6); the same number under constant forwards |
| LLFR weights not summing to 1 | Python raises an error; the VBA computes `sumw` but does not use it, and the sheet shows "ERROR" | Weights sum to 100% (8.3.3) |
| Longest maturity | `Max Maturity` input, 150 by default | The VBA writes the literal 150; EIOPA publishes 1–150 (8.1.7) |

For the VA, the code follows the workbook except in the last row. In the first three rows the
workbook settles points the documentation does not spell out, and the results can differ slightly
from a literal reading of the documentation.

| Topic | This code and the workbook | EIOPA documentation |
|---|---|---|
| Notional VA of the CSSR curve | 85% of the currency RCS, rounded to whole basis points | 85% of the currency RCS (9.3.2 a); only the VA itself is in whole basis points (10.1.3) |
| Best estimate not falling on the CSSR curve | CSSR = 1 if `PVBP_BEL` ≤ 0, i.e. for a positive notional VA if $\Delta BEL \ge 0$ | $\max(\min(\Delta FI / \Delta BEL, 1), 0)$ (9.3.2 d): undefined for $\Delta BEL = 0$, and the limited ratio, not 1, for $\Delta BEL > 0$ |
| Forwards for the curve with VA | From the basic RFR rounded to 5 decimals | From the basic RFR (10.1.6), not stated whether rounded |
| USA and macro VA together | Python rejects it; the workbook adds the macro VA to a currency VA with USA | Not allowed (9.3.4) |

For the interest rate shocks, the code follows the workbook in every point; the technical
documentation does not describe them. Points worth knowing when comparing with other
implementations:

| Topic | This code and the workbook |
|---|---|
| Rates that are shocked | The bootstrapped zero rates after the CRA, annually compounded, up to the last DLT maturity; not the input rates and not the extrapolated basic RFR |
| Floor of the downward shock | Applied to the rounded, extrapolated rates, without and with VA; it does not change the LLFR or the forwards with VA |
| Forwards for the shocked curves with VA | From the unrounded shocked rates; the curve with VA of the basic RFR uses the rounded basic RFR |
| VA in the shocked curves | The undertaking's total VA, not recalculated for the shock |
| Shock parameters | The table of the sheet *Parameters* is an input (`irr_parameters.csv`); Python checks that it covers every maturity and holds decimals |

---

## Scope and limitations

Covered:

- the basic risk-free rate term structure from **swaps** or **government bonds**, i.e. the sheet
  *Input Data & Extrapolation* of EIOPA's workbook with either instrument type;
- the undertaking-specific **volatility adjustment**: the CSSR curve, the CSSR, the currency VA
  with or without USA, the macro VA and the term structure with VA, i.e. the sheets *CSSR
  Calculation* and *VA Term Structures*;
- the **interest rate risk shocks**: the term structures after the upward and downward shocks,
  without and with VA, i.e. the sheets *IRR Up Shock* and *IRR Down Shock*.

Not covered:

- **Valuation of assets and liabilities** on the CSSR curve: these four values are inputs from the
  undertaking's own models.
- **Risk-corrected spreads, debt-to-total-assets ratios and the USA**: inputs, taken from EIOPA's
  publications and, for the USA, from the undertaking's own bond portfolio and the conditions of
  Article 77d(1c).
- **Shock parameters**: the relative and absolute shocks and the floor are inputs, taken from the
  workbook's sheet *Parameters*. The **matching adjustment** (§11) is not covered.
- **Currency risk adjustment** for currencies pegged to the euro (§7). Include it in the CRA if needed.
- **DLT assessment, FSP, LLFR weights and UFR**: these are inputs, taken from EIOPA's publications.
- **Phasing-in of α**: choose the α of the relevant year yourself (see the table above).
- Maturities below one year (EIOPA also publishes from one year onwards, §3.3.1) and
  non-integer maturities.

---

## Repository structure

```text
Code/
  main.py              deletes Output, then run_basic_rfr, run_va and run_irr: read Input,
                       validate, calculate, print a summary, write Output
  calculation.py       the calculation (port of EIOPA's VBA module mExtrapolation); standard library only
  volatility_adjustment.py  the VA (port of the sheets CSSR Calculation and VA Term Structures)
  interest_rate_risk.py     the shocks (port of the sheets IRR Up Shock and IRR Down Shock)
  validation.py        checks that the inputs are complete and plausible
  data_io.py           reading Input/ and writing Output/
  tests/               unit tests; irr_workbook_values.csv holds the workbook's shocked curves
Input/                 curve.csv, parameters.csv and the optional va_parameters.csv and irr_parameters.csv
Output/                swaps_*.csv and government_bonds_*.csv (curves and parameters, with VA: va_*,
                       shocks: irr_*)
Examples/
  EIOPA_Example/       the workbook's example: inputs (Input/, with the shock parameters in
                       irr_parameters.csv) and a copy of the workbook
  Example_September_2026/
                       inputs for 30 September 2026 (Input_EUR, Input_IS, Input_UK), EIOPA's
                       published files for that date (EIOPA_RFR_20260930), charts and a copy of
                       the workbook
Sources/               EIOPA's technical documentation and demonstration workbook (example inputs)
```

To run an example, copy the files of its input folder into `Input/`.

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

The code is released under the MIT License. The license does not cover EIOPA's
documentation and workbook in `Sources/`.
