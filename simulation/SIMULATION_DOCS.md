# Simulation Deep-Dive

**File:** `solar_panel_simulation.xlsx`  
**Engine:** Microsoft Excel + VBA (Visual Basic for Applications)  
**Author:** Roberto Amaya

---

## Why Excel + VBA?

This simulation was built under a hard constraint: **no external software beyond CREO Parametric and Microsoft Excel**. No Python, no MATLAB, no NumPy.

That constraint forced a genuinely interesting engineering challenge — implementing a time-stepped, multi-variable physics simulation with dynamic event logic (cleaning cycles, rain events, real-time chart updates) using only Excel's formula engine and VBA scripting. This document breaks down exactly how that was done.

---

## Workbook Structure

The workbook contains three sheets, each serving a distinct purpose:

### Sheet 1: `Data`

The computation engine. Every row represents a **30-minute interval** in the simulation. Columns:

| Column | Name | Description |
|--------|------|-------------|
| A | Time (hr) | Hour of day (0–18, increments of 0.5) |
| B | Irradiance (W/m²) | Solar intensity calculated via sinusoidal model |
| C | Clean Power (W) | Theoretical output — no soiling |
| D | Dirty Power (W) | Actual output — soiling applied |
| E | Clean Energy (Wh) | Energy produced if panel were clean (P × Δt) |
| F | Dirty Energy (Wh) | Actual energy produced with soiling (P × Δt) |
| G | Dirt Level | Accumulated dirt coefficient (0.0 → 1.0) |
| H | Clean Event | Binary flag — 1 if a cleaning cycle fired this day |
| I | Time (days) | Simulation day index |

The `Data` sheet is populated by VBA on each simulation run. Cells are not manually editable — the macro writes values row by row based on the parameter inputs from the `Display` sheet.

---

### Sheet 2: `Display`

The main user-facing dashboard. Contains:

**Input Controls (user-editable cells):**

| Parameter | What It Does | Typical Range |
|-----------|-------------|---------------|
| Panel Area (m²) | Physical surface area of the PV panel | 1–20 m² |
| Panel Efficiency % | PV cell conversion rate | 15–22% residential |
| Soiling Rate % | Dirt accumulation per day as % of max | 0.1–2.0% |
| Clean Every (Days) | How often the automated cleaner fires | 7–60 days |
| Rain Every (Days) | How often a rain event partially resets dirt | 5–30 days |
| Electricity Rate ($/kWh) | Local energy price for financial loss calculation | $0.10–$0.20 |
| Simulation Length (Days) | Total duration to simulate | 30–365 days |
| Max Loss % | Maximum efficiency degradation at full soiling | 40–80% |

**Output Summary (auto-calculated after run):**

| Output | Description |
|--------|-------------|
| Total Energy (kWh) | What the panel would have produced if always clean |
| Actual Output (kWh) | Real production with dirt accumulation modeled |
| Energy Lost (kWh) | Difference — the cost of soiling |
| Estimated Loss ($) | Dollar value of lost energy at the set electricity rate |

Charts on this sheet update automatically after each simulation run.

---

### Sheet 3: `Display (2)`

A second scenario dashboard, identical in structure to `Display`. Exists to allow **side-by-side scenario comparison** — for example:
- Scenario A: Clean every 20 days, 90-day window
- Scenario B: Clean every 40 days, 40-day window

Both can be configured independently and run separately, letting the user visually compare energy recovery under different maintenance strategies without overwriting results.

---

## The Physics — Formula Breakdown

### Irradiance Model

```
I(t) = 1000 · sin( π(t − 6) / 12 )
```

This approximates the bell-shaped curve of solar irradiance across a standard daylight window (06:00–18:00). Key properties:
- At t = 6 (sunrise): I = 0 W/m²
- At t = 12 (solar noon): I = 1000 W/m² (peak)
- At t = 18 (sunset): I ≈ 0 W/m²
- The 12-hour denominator normalizes the period to one half-cycle of a sine wave

In Excel, this is implemented per row in column B as:
```excel
= 1000 * SIN(PI() * (A2 - 6) / 12)
```

Values below zero (theoretical pre-dawn/post-dusk) are clamped to 0 using `MAX(..., 0)`.

---

### Clean Power Output

```
P_clean = I × A × η
```

- `I` pulled from column B (irradiance)
- `A` and `η` pulled from named cells on the `Display` sheet via cross-sheet reference

```excel
= B2 * Display!PanelArea * (Display!PanelEfficiency / 100)
```

This gives instantaneous power output in Watts for a perfectly clean panel.

---

### Dirt Accumulation Model

Dirt accumulates each simulation day at the configured **Soiling Rate**. The dirt level `D` is a coefficient from 0 (clean) to 1 (maximum soiling):

```
D(day) = D(day - 1) + SoilingRate
D is capped at 1.0 (cannot exceed full soiling)
```

**Rain events** partially reset the dirt level. A rain event applies a partial cleaning — modeled as reducing `D` by a fixed fraction (e.g., 50%) rather than a full reset, since rain doesn't remove baked-on debris as effectively as mechanical cleaning.

**Cleaning cycles** triggered by the "Clean Every N Days" parameter fully reset `D` to 0.

This logic runs inside the VBA simulation loop — see the VBA section below.

---

### Dirty Power Output

```
P_dirty = P_clean × (1 − D × L)
```

- `D` = current dirt level (from column G)
- `L` = Max Loss factor (from `Display` sheet)

At full soiling (D = 1.0) with L = 0.60, the panel operates at only **40% of clean capacity**. At D = 0.20 with L = 0.60, the panel runs at **88% capacity** — showing how even moderate soiling carries a real cost.

---

### Energy per Interval

```
E = P × Δt
```

`Δt` = 0.5 hours (each row is a 30-minute step).

```excel
= C2 * 0.5   ← clean energy (Wh)
= D2 * 0.5   ← dirty energy (Wh)
```

These are summed across all rows to get total kWh figures on the `Display` sheet.

---

### Financial Loss

```
C = (E_clean − E_dirty) × R
```

Where `R` is the electricity rate in $/kWh. Since `E` values are in Wh, the rate is converted:

```excel
= (SUM(CleanEnergy) - SUM(DirtyEnergy)) / 1000 * Display!ElectricityRate
```

---

## VBA — The Automation Engine

The VBA module is where the simulation actually runs. It handles everything that Excel formulas can't do natively: state tracking across days, conditional event firing, and dynamic chart updates.

### Module Overview

```vba
Sub RunSimulation()
    ' 1. Read all parameters from Display sheet
    ' 2. Clear previous Data sheet output
    ' 3. Loop through each day of the simulation
    '    a. For each day, loop through 30-min time steps (t = 6 to 18)
    '    b. Calculate irradiance at each step
    '    c. Calculate clean and dirty power using current dirt level
    '    d. Write all values to Data sheet
    ' 4. Check if today is a cleaning day → reset dirt to 0
    ' 5. Check if today is a rain day → reduce dirt by rain factor
    ' 6. Accumulate dirt by soiling rate (capped at 1.0)
    ' 7. After all rows written, refresh charts
    ' 8. Update summary cells (Total Energy, Actual Output, Lost, $)
End Sub
```

### Key Logic Blocks

**Parameter loading:**
```vba
Dim panelArea As Double
Dim efficiency As Double
Dim soilingRate As Double
Dim cleanInterval As Integer
Dim rainInterval As Integer
Dim electricityRate As Double
Dim simDays As Integer
Dim maxLoss As Double

panelArea     = wsDisplay.Range("PanelArea").Value
efficiency    = wsDisplay.Range("PanelEfficiency").Value / 100
soilingRate   = wsDisplay.Range("SoilingRate").Value / 100
cleanInterval = wsDisplay.Range("CleanEvery").Value
rainInterval  = wsDisplay.Range("RainEvery").Value
electricityRate = wsDisplay.Range("ElectricityRate").Value
simDays       = wsDisplay.Range("SimLength").Value
maxLoss       = wsDisplay.Range("MaxLoss").Value / 100
```

**Simulation loop with event logic:**
```vba
Dim dirtLevel As Double
dirtLevel = 0  ' Start clean

Dim row As Integer
row = 2  ' Start writing from row 2 (row 1 = headers)

Dim day As Integer
For day = 1 To simDays

    ' --- Cleaning event check ---
    Dim cleanEvent As Integer
    cleanEvent = 0
    If day Mod cleanInterval = 0 Then
        dirtLevel = 0       ' Full reset
        cleanEvent = 1      ' Flag this day
    End If

    ' --- Rain event check ---
    If day Mod rainInterval = 0 And cleanEvent = 0 Then
        dirtLevel = dirtLevel * 0.5   ' Partial reset (50%)
    End If

    ' --- Time-step loop (06:00 to 18:00 in 0.5hr steps) ---
    Dim t As Double
    For t = 6 To 18 Step 0.5

        Dim irradiance As Double
        irradiance = 1000 * Sin(WorksheetFunction.Pi() * (t - 6) / 12)
        If irradiance < 0 Then irradiance = 0

        Dim pClean As Double
        Dim pDirty As Double
        pClean = irradiance * panelArea * efficiency
        pDirty = pClean * (1 - dirtLevel * maxLoss)

        Dim eClean As Double
        Dim eDirty As Double
        eClean = pClean * 0.5   ' 0.5hr interval
        eDirty = pDirty * 0.5

        ' Write to Data sheet
        wsData.Cells(row, 1).Value = t              ' Time
        wsData.Cells(row, 2).Value = irradiance     ' Irradiance
        wsData.Cells(row, 3).Value = pClean         ' Clean Power
        wsData.Cells(row, 4).Value = pDirty         ' Dirty Power
        wsData.Cells(row, 5).Value = eClean         ' Clean Energy
        wsData.Cells(row, 6).Value = eDirty         ' Dirty Energy
        wsData.Cells(row, 7).Value = dirtLevel      ' Dirt Level
        wsData.Cells(row, 8).Value = cleanEvent     ' Clean Event flag
        wsData.Cells(row, 9).Value = day            ' Day index

        row = row + 1
    Next t

    ' --- Accumulate dirt for next day ---
    dirtLevel = dirtLevel + soilingRate
    If dirtLevel > 1 Then dirtLevel = 1  ' Cap at 100%

Next day

' --- Refresh charts and update summary ---
Call UpdateSummary
wsDisplay.ChartObjects(1).Chart.Refresh
wsDisplay.ChartObjects(2).Chart.Refresh
```

**Summary update:**
```vba
Sub UpdateSummary()
    Dim totalClean As Double
    Dim totalDirty As Double

    totalClean = WorksheetFunction.Sum(wsData.Columns(5)) / 1000  ' Convert Wh → kWh
    totalDirty = WorksheetFunction.Sum(wsData.Columns(6)) / 1000

    wsDisplay.Range("TotalEnergy").Value  = totalClean
    wsDisplay.Range("ActualOutput").Value = totalDirty
    wsDisplay.Range("EnergyLost").Value   = totalClean - totalDirty
    wsDisplay.Range("EstLoss").Value      = (totalClean - totalDirty) * electricityRate
End Sub
```

---

## Running the Simulation

1. Open `solar_panel_simulation.xlsx` in Microsoft Excel (Windows or Mac)
2. When prompted, click **Enable Macros** — the VBA code will not run without this
3. Navigate to the **Display** sheet
4. Adjust any input parameters you want to test
5. Click the **Run Simulation** button
6. Results populate instantly — summary metrics update and charts refresh automatically
7. To compare scenarios, switch to **Display (2)**, set different parameters, and run again

> **Note:** If macros are blocked by your organization's Excel settings, go to File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros.

---

## Sample Calculations (Manual Verification)

At `t = 10 hours`, `A = 1.6 m²`, `η = 0.20`, `D = 1.0`, `L = 0.20`:

```
I(10)    = 1000 × sin(π × 4 / 12) = 1000 × sin(60°) ≈ 866.0 W/m²
P_clean  = 866.0 × 1.6 × 0.20              ≈ 277.1 W
P_dirty  = 277.1 × (1 − 1.0 × 0.20)       ≈ 221.7 W
E_clean  = 277.1 × 0.5                     = 138.6 Wh
E_dirty  = 221.7 × 0.5                     = 110.9 Wh
C        = (138.6 − 110.9) Wh × $0.00013  ≈ $0.0036 lost per half-hour at peak
```

These match the values in the `Data` sheet exactly, confirming the VBA output is computing correctly.
