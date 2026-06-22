# VBA Walkthrough — How the Automation Was Built

**Author:** Roberto Amaya  
**Tool:** Microsoft Excel — Visual Basic for Applications (VBA)

---

## Background: Why VBA?

VBA (Visual Basic for Applications) is Excel's built-in programming environment. It lets you write code that interacts directly with your spreadsheet — reading cells, writing values, looping through rows, responding to button clicks, and refreshing charts. It's not Python or JavaScript, but it's genuinely powerful for data automation, especially within Microsoft Office.

For this project, VBA was the only way to implement several critical behaviors that Excel formulas simply cannot do:

| Behavior | Why Formulas Can't Handle It | How VBA Solves It |
|----------|------------------------------|-------------------|
| Simulate time passing day-by-day | Formulas are static — they evaluate once | VBA loops through each day sequentially |
| Track dirt accumulation as state | A cell can't remember its previous value in a formula | VBA uses variables that persist across loop iterations |
| Fire cleaning/rain events conditionally | `IF()` can evaluate conditions, but can't trigger actions | VBA checks conditions and resets state variables |
| Write hundreds of output rows automatically | Manual entry or formula drag — not scalable | VBA writes each row in a loop, one cell at a time |
| Refresh charts after data changes | Charts don't auto-update in all configurations | VBA calls `.Chart.Refresh` explicitly |

Learning VBA for this project meant understanding the object model of Excel (Workbooks → Worksheets → Ranges → Cells), how to declare variables, how to structure loops, and how to reference named ranges across sheets.

---

## Module Structure

The VBA project contains one standard module with three procedures:

```
Module1
├── RunSimulation()     ← Main entry point, called by the button on Display sheet
├── UpdateSummary()     ← Calculates and writes KPI totals after simulation
└── ClearData()         ← Wipes the Data sheet before each new run
```

---

## Step-by-Step: What Happens When You Click "Run Simulation"

### Step 1 — Clear Previous Results

```vba
Sub ClearData()
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Sheets("Data")
    
    ' Clear everything below the header row
    wsData.Range("A2:I10000").ClearContents
End Sub
```

Before writing new results, the old data is wiped. This prevents leftover rows from a previous run from mixing with the new output.

---

### Step 2 — Load Parameters from the Display Sheet

```vba
Dim wsDisplay As Worksheet
Dim wsData As Worksheet
Set wsDisplay = ThisWorkbook.Sheets("Display")
Set wsData    = ThisWorkbook.Sheets("Data")

' Read user inputs from named cells
Dim panelArea       As Double  : panelArea     = wsDisplay.Range("PanelArea").Value
Dim efficiency      As Double  : efficiency    = wsDisplay.Range("PanelEfficiency").Value / 100
Dim soilingRate     As Double  : soilingRate   = wsDisplay.Range("SoilingRate").Value / 100
Dim cleanInterval   As Integer : cleanInterval = wsDisplay.Range("CleanEvery").Value
Dim rainInterval    As Integer : rainInterval  = wsDisplay.Range("RainEvery").Value
Dim electricityRate As Double  : electricityRate = wsDisplay.Range("ElectricityRate").Value
Dim simDays         As Integer : simDays       = wsDisplay.Range("SimLength").Value
Dim maxLoss         As Double  : maxLoss       = wsDisplay.Range("MaxLoss").Value / 100
```

Named ranges (like `PanelArea`, `CleanEvery`) were defined in Excel's Name Manager, so the VBA doesn't need to know which specific cell each value lives in. This keeps the code readable and makes the sheet easier to reorganize later.

---

### Step 3 — Initialize State Variables

```vba
Dim dirtLevel  As Double  : dirtLevel = 0.0   ' Panels start clean
Dim row        As Integer : row = 2            ' Start writing at row 2 (row 1 = headers)
```

`dirtLevel` is the key state variable — it carries forward from one day to the next, accumulating, resetting on cleaning events, and partially resetting on rain. This is the core of what makes VBA necessary; no Excel formula can hold state between rows this way.

---

### Step 4 — The Outer Loop: Days

```vba
Dim day As Integer
For day = 1 To simDays

    ' ... (event checks and time-step loop go here)

    ' Accumulate dirt at end of each day
    dirtLevel = dirtLevel + soilingRate
    If dirtLevel > 1.0 Then dirtLevel = 1.0   ' Hard cap

Next day
```

The outer loop runs once per simulated day. At the end of each day, dirt accumulates by the soiling rate. The `If dirtLevel > 1.0` cap prevents the dirt coefficient from exceeding 1.0 (which would break the power output formula).

---

### Step 5 — Event Checks (Cleaning and Rain)

```vba
' --- Check for scheduled cleaning ---
Dim cleanEvent As Integer : cleanEvent = 0

If day Mod cleanInterval = 0 Then
    dirtLevel = 0       ' Full reset — mechanical cleaning
    cleanEvent = 1      ' Flag so the Data sheet records it
End If

' --- Check for rain (only if no cleaning happened today) ---
If day Mod rainInterval = 0 And cleanEvent = 0 Then
    dirtLevel = dirtLevel * 0.5   ' Rain removes ~50% of current dirt
End If
```

`Mod` is the modulo operator — it returns the remainder of division. So `day Mod 20 = 0` is true on days 20, 40, 60, etc. This elegantly handles recurring events without needing a calendar or date logic.

Rain and cleaning are mutually exclusive on the same day — if cleaning already fired, rain is skipped (the `cleanEvent = 0` guard). This models the real scenario where a scheduled cleaning supersedes the rain effect.

---

### Step 6 — The Inner Loop: Time Steps

```vba
Dim t As Double
For t = 6 To 18 Step 0.5
    
    ' Physics calculations
    Dim irradiance As Double
    irradiance = 1000 * Sin(WorksheetFunction.Pi() * (t - 6) / 12)
    If irradiance < 0 Then irradiance = 0
    
    Dim pClean As Double : pClean = irradiance * panelArea * efficiency
    Dim pDirty As Double : pDirty = pClean * (1 - dirtLevel * maxLoss)
    
    Dim eClean As Double : eClean = pClean * 0.5   ' Wh per 30-min step
    Dim eDirty As Double : eDirty = pDirty * 0.5
    
    ' Write one row to Data sheet
    wsData.Cells(row, 1).Value = t
    wsData.Cells(row, 2).Value = irradiance
    wsData.Cells(row, 3).Value = pClean
    wsData.Cells(row, 4).Value = pDirty
    wsData.Cells(row, 5).Value = eClean
    wsData.Cells(row, 6).Value = eDirty
    wsData.Cells(row, 7).Value = dirtLevel
    wsData.Cells(row, 8).Value = cleanEvent
    wsData.Cells(row, 9).Value = day
    
    row = row + 1

Next t
```

For each day, the inner loop steps through every 30-minute interval from 06:00 to 18:00 (`Step 0.5`). Each iteration calculates the physics for that time slot using the current `dirtLevel` and writes one row to the `Data` sheet.

`WorksheetFunction.Pi()` is VBA's way of calling Excel's `PI()` function — VBA doesn't have a built-in Pi constant, so this bridges into Excel's math library.

---

### Step 7 — Summary and Chart Refresh

```vba
Sub UpdateSummary()
    Dim totalClean As Double
    Dim totalDirty As Double

    ' Sum entire columns E and F from Data sheet (already in Wh)
    totalClean = WorksheetFunction.Sum(wsData.Columns(5)) / 1000
    totalDirty = WorksheetFunction.Sum(wsData.Columns(6)) / 1000

    wsDisplay.Range("TotalEnergy").Value  = Round(totalClean, 1)
    wsDisplay.Range("ActualOutput").Value = Round(totalDirty, 1)
    wsDisplay.Range("EnergyLost").Value   = Round(totalClean - totalDirty, 1)
    wsDisplay.Range("EstLoss").Value      = Round((totalClean - totalDirty) * electricityRate, 2)
End Sub

' Refresh charts so they redraw with new data
wsDisplay.ChartObjects(1).Chart.Refresh
wsDisplay.ChartObjects(2).Chart.Refresh
```

After the loops finish, VBA sums the entire energy columns and writes the four KPI values back to named cells on the `Display` sheet. `Round()` keeps the output clean. The chart refresh calls force Excel to redraw the visualizations using the newly written `Data` rows.

---

## What I Learned Building This

Going into this project I had no VBA experience. The learning curve hit a few specific places:

**Object model.** Excel VBA revolves around a hierarchy: `Application → Workbook → Worksheet → Range → Cell`. Getting comfortable with `Set ws = ThisWorkbook.Sheets("Name")` and then `ws.Cells(row, col).Value` took some adjustment, but once it clicked it made the code much more readable.

**Named ranges.** Early versions of the code used hardcoded cell addresses like `wsDisplay.Range("B4").Value`. This broke every time the sheet layout changed. Switching to named ranges (defined in Excel's Name Manager) made the VBA resilient to layout edits.

**State variables in loops.** The `dirtLevel` variable was the central insight. I initially tried to model dirt accumulation purely in spreadsheet formulas, but there's no way to make a formula reference its own previous value without circular reference errors. Moving the state into a VBA variable that persists across loop iterations was the solution.

**Debugging.** `Debug.Print` to the Immediate Window and adding a `MsgBox dirtLevel` inside the loop were the main tools. VBA's debugger also lets you step through code line by line with F8, which was invaluable for catching the off-by-one error in the day indexing.

---

## Performance Note

For a 90-day simulation with 25 time steps per day, the loop writes **2,250 rows** to the Data sheet. Writing cell-by-cell (one `wsData.Cells(row, col).Value = x` call per value) is the straightforward approach but can be slow for very large simulations.

A performance optimization would be to build the output in a 2D array and write it to the sheet in a single operation:

```vba
' Faster approach for large simulations
Dim outputArr() As Variant
ReDim outputArr(1 To totalRows, 1 To 9)

' ... populate array in loop ...

wsData.Range("A2").Resize(totalRows, 9).Value = outputArr
```

This batches all writes into one call, which is significantly faster. It's a good next step for anyone extending the simulation to longer time windows.
