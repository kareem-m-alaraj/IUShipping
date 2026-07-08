# IUShipping — Transportation Cost Optimization Model

A logistics decision-support tool built in Excel/VBA that minimizes weekly shipping costs across a multi-source, multi-destination distribution network in Bosnia & Herzegovina, using classical Operations Research (Linear Programming) methods.

Developed for **IE303 – Operations Research I**, International University of Sarajevo (Spring 2026), under Prof. Ozge Buyukdagli.

---

## Overview

IUShipping models a weekly shipping plan for a consumer goods distributor operating **3 production centres** and **5 warehouses**. It formulates and solves a classical transportation LP problem to determine the cost-minimizing shipment plan, then wraps the model in an interactive Excel dashboard so a non-technical manager can adjust costs, supply/demand, and scenarios without touching the underlying optimization logic.

**Key result:** under normal operating conditions, the optimal weekly transportation cost is **11,371.04 BAM**. The model also supports scenario analysis (busy season, fuel price spikes, and combined stress conditions) to evaluate cost sensitivity under changing operating assumptions.

---

## Problem Formulation

**Decision variables:** `x_ij` = tons shipped from source *i* to destination *j*

**Objective:** minimize total weekly transportation cost

```
Minimize  Z = Σ Σ c_ij * x_ij
```

**Subject to:**
- Supply constraints at each production centre
- Demand constraints at each warehouse
- Non-negativity: `x_ij ≥ 0`

Supply and demand are balanced using a **dummy source/destination** (per Hillier & Lieberman's transportation method), so managers can freely edit supply and demand values in the interface without breaking feasibility.

**Unit cost formula**, computed dynamically from live cost parameters (fuel price, fuel consumption, driver wages, tolls):

```
c_ij = (D_ij * (FL + FE) * Pf + 2*T_ij*W + 2*Toll_ij) / Q + O
```

Where `D` = distance, `FL`/`FE` = loaded/empty fuel consumption, `Pf` = fuel price, `T` = travel time, `W` = driver wage rate, `Toll` = round-trip toll cost, `Q` = truck capacity, `O` = per-ton overhead.

Solved using the **Simplex method** via Excel Solver.

---

## Workbook Structure

| Sheet | Purpose |
|---|---|
| **README** | In-workbook documentation: structure, workflow, scenario definitions, recorded results |
| **DATA** | Formula-driven transportation cost matrix, derived from live cost parameters |
| **MODEL** | Solver-based LP model: decision variable matrix, constraints, objective function |
| **INTERFACE** | Manager-facing dashboard — edit costs/supply/demand, pick scenarios, run optimization, reset to defaults |
| **SCENARIOS** | Scenario definitions and results log (Normal, Busy Season, Fuel Spike, Combined Stress) |
| **RESULTS** | Final shipment plan, cost breakdown, scenario comparison chart, LP-vs-true-cost reconciliation |

---

## Scenario Analysis

| Scenario | Optimal Weekly Cost (BAM) |
|---|---|
| Normal Operations | 11,371.04 |
| Busy Season | 12,584.12 |
| Fuel Spike | 12,319.00 |
| Combined Stress | 13,627.93 |

Scenarios are switchable from the INTERFACE sheet via VBA macros that update cost and demand parameters and re-solve the model automatically.

**LP vs. true operating cost:** because the LP relaxation treats truck loads as continuous, its optimum understates real-world cost (where partial truckloads still require a full truck). A reconciliation block on the RESULTS sheet recalculates true cost using truckload rounding — showing a gap of roughly **10.5%** above the LP optimum for the Normal scenario, largest on low-utilization routes.

---

## Features

- **Dynamic cost matrix** — every cost cell recalculates automatically from named parameter cells (fuel price, wages, tolls, etc.); nothing is hardcoded
- **Interactive dashboard** — editable supply/demand, one-click scenario switching, Reset button, live scenario-comparison chart
- **Solver automation** — VBA-driven Solver runs triggered directly from the interface
- **Cost reconciliation** — LP optimum vs. realistic truckload-rounded cost, with gap analysis

---

## Data Sources

- [JP Autoceste FBiH](https://jpautoceste.ba) — toll rates (FBiH routes)
- JP Autoputevi RS — toll rates (Republika Srpska routes)
- [GlobalPetrolPrices](https://www.globalpetrolprices.com) — fuel price benchmarks
- HIFA Petrol — regional fuel pricing
- Adorio.ba — driver wage benchmarks
- Hillier & Lieberman, *Introduction to Operations Research*, Ch. 9 — transportation problem methodology

---

## Tech Stack

- Microsoft Excel (.xlsm)
- VBA (macros, form controls, Solver automation)
- Excel Solver (Simplex LP)

---

## Files

- `IUShipping_Optimization.xlsm` — full workbook (model, interface, macros)
- `IUShipping_Report.pdf` — project report: methodology, formulation, scenario comparison, insights

---

## Author

Abdalkarim — IE303 Operations Research I, International University of Sarajevo
