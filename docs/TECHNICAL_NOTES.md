# Technical Notes & Design Decisions

**Project:** Solar Savior — Automated Solar Panel Cleaning System  
**Author:** Roberto Amaya | ENES 100, May 2026

This document captures the engineering rationale behind key decisions made during the design process — the "why" behind the choices that aren't always obvious from the final output alone.

---

## CREO Parametric Design

![CREO Technical Drawing](../assets/creo_design.jpg)

The CREO drawing above shows the complete screw drive rail assembly with annotated dimensions. Three views are provided:

- **Top-left (front elevation):** Full rail length of 365.00 mm with 46.00 mm inner offset. The roller array (Ø15.00 mm) is evenly distributed along the rail, and the 40.00 mm outer wall thickness (O_THICK) provides structural rigidity for outdoor loading.
- **Bottom-left (side profile):** 45.00 mm frame width showing the low-profile mounting geometry designed to sit flush against a panel frame.
- **Right (cross-section):** The threaded screw assembly in cross-section — Ø31.00 mm outer thread diameter, Ø19.76 mm inner shaft, 8.00 mm thread pitch offset, and 35.00 mm drive section length. The spoke-style end cap confirms hollow-core construction for weight reduction.

---

## Why a Screw Drive vs. Belt Drive

The first major mechanical decision was the drive mechanism for the cleaning carriage. Two options were seriously considered:

**Belt-driven system**
- Common in 3D printers and CNC machines
- Lightweight, inexpensive
- *Problem:* Belt tension changes with temperature and moisture. Outdoor environments — especially rooftops — expose belts to rain, UV, thermal cycling, and debris. Slippage on an inclined surface would cause the cleaning arm to drift or stall, delivering inconsistent pressure across the panel.

**Screw drive system (chosen)**
- Converts rotational motor torque directly to linear motion via a threaded rod
- No slip possible — position is mechanically enforced by the thread pitch
- Higher torque delivery at lower RPM, suitable for working against gravity on inclined panels
- More resistant to contamination — debris falls off the thread rather than jamming a belt groove
- Trade-off: heavier and slightly slower than belt-driven, but for this application speed is not the priority

The CREO model validated that the motor torque spec was sufficient to move the carriage up and across the panel at the maximum expected installation angle.

---

## Why a Sinusoidal Irradiance Model

The simulation uses `I(t) = 1000 · sin(π(t−6)/12)` rather than tabulated irradiance data. This was a deliberate simplification with clear tradeoffs:

**What the model captures:**
- The general bell shape of solar intensity across a day
- Zero irradiance at sunrise and sunset
- Peak irradiance (1000 W/m²) at solar noon — consistent with Standard Test Conditions (STC) used in PV panel ratings
- Smooth transitions, avoiding the jagged noise of real weather data

**What it doesn't capture:**
- Cloud cover and atmospheric scattering
- Seasonal variation in day length and sun angle
- Geographic latitude effects
- Panel orientation and tilt angle

For a first-order feasibility analysis and ROI estimate, these simplifications are acceptable. The model is intended to show the *relative* difference between clean and dirty panel performance, not predict absolute energy yield for a specific installation. Real-world validation with measured irradiance data is listed as future work.

---

## Why a 0.5-Hour Time Step

The simulation steps through time in 30-minute intervals. This was chosen as the balance point between:

- **Accuracy:** Irradiance changes relatively slowly across a day; 30-minute resolution captures the shape of the curve without over-sampling
- **Performance:** A 30-minute step produces 25 rows per day. At 90 days, that's 2,250 rows — manageable for Excel without performance degradation
- **Practical relevance:** Energy billing and grid monitoring typically use 15–30 minute intervals, so this resolution aligns with real-world reporting standards

Finer resolution (e.g., 5-minute steps) would produce 180 rows/day × 90 days = 16,200 rows, which starts to stress Excel on older hardware and doesn't meaningfully change the result for this type of analysis.

---

## Why 20-Day vs. 40-Day Cleaning Intervals Were Compared

Two cleaning interval scenarios were modeled to answer a specific question: **does cleaning twice as often actually help proportionally?**

The answer from the simulation is no — it helps *more* than proportionally:

| Cleaning Frequency | Energy Lost (90 days) | Financial Loss |
|----|----|----|
| Every 20 days | 865.7 kWh | ~$45.45 |
| Every 40 days | 479.8 kWh (40-day sim) | ~$25.19 |

The nonlinear relationship exists because dirt accumulation follows a compounding effect — as `dirtLevel` increases, the marginal loss per additional day of soiling increases (since `P_dirty = P_clean × (1 − D × L)` grows worse faster at higher `D`). Cleaning early, while `D` is still low, recovers more energy per cleaning event than cleaning late after the panel is heavily soiled.

This finding directly supports the design case for *automated* cleaning: it's not just about cleaning — it's about cleaning on the right schedule.

---

## Why Rain Events Use 50% Partial Reset

The rain model resets `dirtLevel` by 50% rather than 100%. This reflects real photovoltaic maintenance data:

- Rain is effective at removing loose particulates (dust, pollen, light debris)
- It is ineffective at removing bonded soiling (bird droppings, industrial residue, baked-on mineral deposits)
- Studies suggest rainfall restores approximately 40–60% of soiled panel efficiency depending on the nature of the debris

The 50% factor is a middle-ground approximation. In a more sophisticated model, this factor would vary with rainfall intensity, duration, and local soiling composition. Future work could pull historical precipitation data via an API to make this dynamic.

---

## Constraint: CREO + Excel Only

The most unusual aspect of this project is the tool restriction. In most engineering programs, a performance simulation like this would be built in Python (with NumPy and pandas), MATLAB, or a dedicated simulation tool like EnergyPlus.

Building it entirely in Excel + VBA required:

1. **Rethinking what "simulation" means in a spreadsheet.** Time-stepped simulation with state is not what Excel is designed for. Forcing it to work required understanding VBA's object model, using variables to carry state across loop iterations, and treating each row write as a discrete event.

2. **Learning VBA from scratch.** No prior programming experience in VBA was required for this course. The VBA walkthrough in this repo (`simulation/VBA_WALKTHROUGH.md`) documents the specific concepts learned and the debugging approaches used.

3. **Being deliberate about performance.** Knowing that Excel has limits on how many cells can be written efficiently shaped the choice of time step resolution and simulation length.

The constraint produced a better engineer — the limitations forced creativity and a deeper understanding of the underlying physics, because there was no library function to abstract it away.
