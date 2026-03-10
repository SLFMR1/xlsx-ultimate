# Engineering Domain Reference

## Table of Contents
1. [Mechanical Engineering](#mechanical)
2. [Electrical Engineering](#electrical)
3. [Civil/Structural Engineering](#civil)
4. [Chemical Engineering](#chemical)
5. [Manufacturing & Quality](#manufacturing)
6. [Statistics & Data Science](#statistics)
7. [Systems Engineering](#systems)
8. [Unit Handling & Dimensional Analysis](#units)

---

## Mechanical Engineering {#mechanical}

### Stress Calculations

```
Axial Stress: σ = F / A
Bending Stress: σ = M × y / I
Shear Stress: τ = V × Q / (I × b)
Torsional Shear: τ = T × r / J
Von Mises: σ_vm = √(σ₁² - σ₁σ₂ + σ₂²)
Factor of Safety: FoS = Yield_Strength / Applied_Stress
```

### Beam Deflection (Common Cases)

```
Cantilever, point load at end: δ = P × L³ / (3 × E × I)
Simply supported, center load: δ = P × L³ / (48 × E × I)
Simply supported, uniform load: δ = 5 × w × L⁴ / (384 × E × I)
```

### Fatigue Analysis

```
Endurance Limit (steel): Se' ≈ 0.5 × Sut (for Sut ≤ 200 ksi)
Modified Endurance: Se = ka × kb × kc × kd × ke × Se'
  ka = surface factor
  kb = size factor
  kc = reliability factor
  kd = temperature factor
  ke = miscellaneous factor

Goodman Line: σa/Se + σm/Sut = 1/n
Soderberg Line: σa/Se + σm/Sy = 1/n
```

### Tolerance Stackup

```
Worst Case: Total = Σ|individual tolerances|
RSS (Root Sum Square): Total = √(Σtolerance²)
Monte Carlo: Simulate N random samples from tolerance distributions
```

### Thermal Analysis

```
Linear Expansion: ΔL = α × L₀ × ΔT
Thermal Stress: σ = E × α × ΔT (if constrained)
Heat Transfer (conduction): Q = k × A × ΔT / L
Heat Transfer (convection): Q = h × A × ΔT
```

### Verification for Mechanical

```python
# Always verify units are consistent
# Use pint library for unit tracking
from pint import UnitRegistry
ureg = UnitRegistry()

force = 10000 * ureg.newton
area = 0.005 * ureg.meter**2
stress = (force / area).to('megapascal')

# Check against material properties
assert stress.magnitude < yield_strength_mpa, "FAIL: Stress exceeds yield"

# Factor of safety check
fos = yield_strength_mpa / stress.magnitude
assert fos >= required_fos, f"FAIL: FoS {fos:.2f} < required {required_fos}"
```

---

## Electrical Engineering {#electrical}

### Circuit Analysis

```
Ohm's Law: V = I × R
Power: P = V × I = I² × R = V² / R
Series Resistance: R_total = R1 + R2 + R3
Parallel Resistance: 1/R_total = 1/R1 + 1/R2 + 1/R3
Kirchhoff's Current Law: ΣI_in = ΣI_out
Kirchhoff's Voltage Law: ΣV_loop = 0
```

### Power Calculations (3-Phase)

```
Single Phase Power: P = V × I × PF
Three Phase Power: P = √3 × V_LL × I × PF
Reactive Power: Q = V × I × sin(φ)
Apparent Power: S = V × I
Power Factor: PF = P / S = cos(φ)
```

### Cable Sizing

```
Current Capacity: Based on cable type, installation method, ambient temperature
Voltage Drop: Vd = (2 × L × I × R_per_m) / 1000  (single phase)
Voltage Drop: Vd = (√3 × L × I × R_per_m) / 1000  (three phase)
Max Voltage Drop: Typically ≤ 3% for branch circuits, ≤ 5% total
```

### Load Schedule

```
Columns: Load ID, Description, kW Rating, PF, kVA, Demand Factor, Connected kW, Demand kW
Connected Load: =Rating × Quantity
Demand Load: =Connected × Demand_Factor
Total Demand: =SUM(Demand_kW)
Diversity Factor: =Sum_Individual_Max / System_Max
```

---

## Civil/Structural Engineering {#civil}

### Structural Load Combinations (ASCE 7)

```
1.4D
1.2D + 1.6L + 0.5(Lr or S or R)
1.2D + 1.6(Lr or S or R) + (L or 0.5W)
1.2D + 1.0W + L + 0.5(Lr or S or R)
1.2D + 1.0E + L + 0.2S
0.9D + 1.0W
0.9D + 1.0E
```

### Concrete Design

```
Moment Capacity: Mn = As × fy × (d - a/2)
  where a = As × fy / (0.85 × f'c × b)
Shear Capacity: Vc = 2 × √f'c × b × d
Required Steel Area: As = Mu / (φ × fy × (d - a/2))
Min Steel Ratio: ρ_min = MAX(3×√f'c/fy, 200/fy)
```

### Hydraulic Calculations

```
Manning's Equation: Q = (1/n) × A × R^(2/3) × S^(1/2)
Bernoulli: P₁/γ + V₁²/2g + z₁ = P₂/γ + V₂²/2g + z₂
Hazen-Williams: V = 0.849 × C × R^0.63 × S^0.54
Darcy-Weisbach: hf = f × (L/D) × V²/(2g)
```

### Earthwork Calculations

```
Cut Volume: Average End Area Method
V = (A1 + A2) / 2 × L
Mass Haul: Cumulative cut minus cumulative fill
```

---

## Chemical Engineering {#chemical}

### Mass Balance

```
Accumulation = Input - Output + Generation - Consumption
Steady State: 0 = Input - Output + Generation - Consumption
Conservation: Total mass in = Total mass out (steady state, no reaction)
Component Balance: x_i × F_in = y_i × F_out (per component)
```

### Energy Balance

```
Q = m × Cp × ΔT (sensible heat)
Q = m × λ (latent heat)
Heat Exchanger: Q = U × A × LMTD
LMTD = (ΔT1 - ΔT2) / ln(ΔT1/ΔT2)
```

### Pipe Sizing

```
Flow Rate: Q = V × A = V × π × D² / 4
Reynolds Number: Re = ρ × V × D / μ
Friction Factor (Moody): f = f(Re, ε/D)
Pressure Drop: ΔP = f × (L/D) × ρ × V² / 2
Pump Power: P = Q × ΔP / η
```

### Reaction Kinetics

```
Zero Order: [A] = [A]₀ - k×t
First Order: [A] = [A]₀ × e^(-k×t)
Second Order: 1/[A] = 1/[A]₀ + k×t
Arrhenius: k = A × e^(-Ea/(R×T))
```

### Verification for Chemical Engineering

```python
# Mass balance closure check
total_in = sum(feed_streams)
total_out = sum(product_streams) + sum(waste_streams)
closure = abs(total_in - total_out) / total_in * 100
assert closure < 0.1, f"Mass balance closure {closure:.2f}% > 0.1%"

# Energy balance
q_in = sum(heat_inputs)
q_out = sum(heat_outputs) + sum(heat_losses)
assert abs(q_in - q_out) / q_in < 0.01
```

---

## Manufacturing & Quality {#manufacturing}

### SPC (Statistical Process Control)

```
X-bar Chart:
  Center Line: X̄ = Grand Mean
  UCL = X̄ + A2 × R̄
  LCL = X̄ - A2 × R̄

R Chart:
  Center Line: R̄ = Average Range
  UCL = D4 × R̄
  LCL = D3 × R̄

Constants (n=5): A2=0.577, D3=0, D4=2.114
```

### Process Capability

```
Cp = (USL - LSL) / (6 × σ)
Cpk = MIN((USL - μ) / (3σ), (μ - LSL) / (3σ))
Pp = (USL - LSL) / (6 × s)   [s = sample std dev]
Ppk = MIN((USL - μ) / (3s), (μ - LSL) / (3s))

Target: Cpk ≥ 1.33 (4σ), preferably ≥ 1.67 (5σ)
Six Sigma: Cpk ≥ 2.0
```

### OEE (Overall Equipment Effectiveness)

```
Availability = Run_Time / Planned_Production_Time
Performance = (Ideal_Cycle_Time × Total_Count) / Run_Time
Quality = Good_Count / Total_Count
OEE = Availability × Performance × Quality

World-Class Target: OEE ≥ 85%
```

### GR&R (Gage Repeatability & Reproducibility)

```
Repeatability (Equipment Variation): EV = K1 × R̄
Reproducibility (Appraiser Variation): AV = √((K2 × X̄_diff)² - (EV² / (n×r)))
GR&R = √(EV² + AV²)
%GR&R = (GR&R / Tolerance) × 100

Acceptable: %GR&R < 10%
Marginal: 10% ≤ %GR&R ≤ 30%
Unacceptable: %GR&R > 30%
```

### Cost Estimation

```
Material Cost = Weight × Material_Price_per_kg
Labor Cost = Time × Labor_Rate
Overhead = Labor_Cost × Overhead_Rate
Total Part Cost = Material + Labor + Overhead
Batch Cost = Setup_Cost + (Total_Part_Cost × Quantity)
Unit Cost = Batch_Cost / Quantity
```

---

## Statistics & Data Science {#statistics}

### Descriptive Statistics

```
Mean: =AVERAGE(range)
Median: =MEDIAN(range)
Mode: =MODE(range)
Standard Deviation (sample): =STDEV.S(range)
Standard Deviation (population): =STDEV.P(range)
Variance: =VAR.S(range) or VAR.P(range)
Percentile: =PERCENTILE.INC(range, k)
IQR: =QUARTILE.INC(range, 3) - QUARTILE.INC(range, 1)
Coefficient of Variation: =STDEV.S(range) / AVERAGE(range)
```

### Regression

```
Slope: =SLOPE(known_y, known_x)
Intercept: =INTERCEPT(known_y, known_x)
R-squared: =RSQ(known_y, known_x)
Predicted: =FORECAST(x, known_y, known_x)
Standard Error: =STEYX(known_y, known_x)
```

### Hypothesis Testing

```
T-Test (two sample): =T.TEST(array1, array2, tails, type)
  type: 1=paired, 2=equal variance, 3=unequal variance
Z-Test: =Z.TEST(array, x, sigma)
Chi-Square: =CHISQ.TEST(observed, expected)
F-Test: =F.TEST(array1, array2)

Confidence Interval: =CONFIDENCE.NORM(alpha, stdev, n)
```

### Verification for Statistics

```python
import numpy as np
from scipy import stats

# Verify Excel AVERAGE
python_mean = np.mean(data)
assert math.isclose(excel_mean, python_mean, rel_tol=1e-10)

# Verify Excel STDEV.S
python_std = np.std(data, ddof=1)  # ddof=1 for sample
assert math.isclose(excel_stdev, python_std, rel_tol=1e-10)

# Verify regression
slope, intercept, r, p, se = stats.linregress(x, y)
assert math.isclose(excel_slope, slope, rel_tol=1e-8)
assert math.isclose(excel_r_squared, r**2, rel_tol=1e-8)
```

---

## Systems Engineering {#systems}

### Requirements Traceability Matrix (RTM)

```
Columns: Req ID, Requirement Text, Priority, Source, Design Element,
         Test Case, Verification Method, Status, Trace Links
Verification Methods: Test, Demonstration, Inspection, Analysis
```

### Risk Matrix

```
Probability (1-5) × Impact (1-5) = Risk Score
Color coding:
  Score 1-4: Green (Low)
  Score 5-9: Yellow (Medium)
  Score 10-16: Orange (High)
  Score 20-25: Red (Critical)
```

### Decision Matrix (Pugh / Weighted Scoring)

```
Columns: Criteria, Weight, Option A Score, Option A Weighted,
         Option B Score, Option B Weighted, ...
Weighted Score: =Score × Weight
Total: =SUM(Weighted_Scores)
Winner: Option with highest total
```

---

## Unit Handling & Dimensional Analysis {#units}

### Python Unit Verification with Pint

```python
from pint import UnitRegistry
ureg = UnitRegistry()

# Define quantities with units
force = 50 * ureg.kilonewton
area = 200 * ureg.millimeter**2

# Calculate with automatic unit tracking
stress = (force / area).to('megapascal')
print(f"Stress: {stress:.1f}")  # 250.0 megapascal

# Unit consistency is enforced automatically
# This would raise DimensionalityError:
# mass + force  # Error! Can't add kg and N
```

### Common Unit Conversions

```
1 inch = 25.4 mm
1 psi = 6.895 kPa
1 ksi = 6.895 MPa
1 lb = 4.448 N
1 BTU = 1055.06 J
1 HP = 745.7 W
1 ft³/s = 28.317 L/s
1 gpm = 0.0631 L/s
°F to °C: (°F - 32) × 5/9
```

### Dimensional Analysis Best Practices

1. ALWAYS state units in column headers (e.g., "Force (kN)", "Stress (MPa)")
2. Use consistent unit systems throughout (don't mix metric and imperial)
3. Include unit conversion sheet if multiple systems needed
4. Verify dimensional homogeneity of all formulas
5. Use named ranges with unit suffixes (e.g., `Force_kN`, `Area_mm2`)

### Verification Pattern

```python
# For every engineering formula, verify units cancel correctly
# Example: σ = F/A should give pressure units
force_units = ureg.newton
area_units = ureg.meter**2
result_units = force_units / area_units
assert result_units.dimensionality == ureg.pascal.dimensionality
```
