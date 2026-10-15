# SMARTIS — Software Requirements Specification (SRS)

**System:** SMARTIS — Strategic Maintenance, Asset Reliability & Tracking Information System
**Client:** Deeprock Ghana Limited
**Document type:** Requirements specification (reverse-engineered from the delivered `index.html` build, v4)
**Status:** Baseline — describes the system as currently built
**Date:** 14 September 2026

---

## 1. Introduction

### 1.1 Purpose
SMARTIS is a single-file, browser-based maintenance and reliability dashboard for Deeprock Ghana Limited's mixed drilling fleet. It gives the maintenance team and management a real-time view of asset health, work orders, preventive-maintenance (PM) compliance, spares, cost and reliability KPIs, plus a costed ROI justification for the reliability program. This document specifies what the system must do (functional requirements) and the qualities it must meet (non-functional requirements).

### 1.2 Scope
The product is a **self-contained HTML dashboard** (`index.html`) with **no backend and no build step**. All logic runs in the viewer's browser. It ingests fleet data from one of three sources — an uploaded Excel/CSV workbook, a connected Google Sheet, or built-in demo data — and renders 14 analysis views. It does **not** write back to any data source; data entry (e.g. new work orders) is produced as copy-and-paste rows for the user to add to their own sheet.

### 1.3 Intended users
| User | Primary use |
|------|-------------|
| Maintenance Officer (e.g. Shaddy) | Daily driving of the dashboard, shift reports, WO creation, ROI justification |
| Technicians | Work-order status, PM calendar, asset health |
| Site / Operations Management | Executive summary, KPIs, cost and ROI views |
| Whole team | View a shared deployed link and load the current dataset |

### 1.4 Definitions
- **RC Rig** — Reverse-Circulation drilling rig; requires an **external compressor** for air supply.
- **DD Rig** — Diamond Drill rig; uses a **bean pump** for water/fluid circulation (no external compressor).
- **Compressor** — Independent asset supplying air to an RC rig; its failure **cascades** to the connected rig.
- **Light Vehicle (LV)** — Bakkies, crew buses, water trucks, fuel bowsers, etc.
- **Cascade failure** — A compressor going DOWN stops the RC rig hosed to it; both are flagged together.
- **PM** — Preventive Maintenance. **WO** — Work Order. **MTBF / MTTR** — Mean Time Between/To Repair.
- **P-F interval** — Time from potential failure (P) to functional failure (F). **RUL** — Remaining Useful Life.
- **Health score** — Per-system condition, scaled **1 (Excellent) → 6 (Critical)**.

---

## 2. Overall description

### 2.1 Product perspective
SMARTIS is a standalone client-side application. It loads one third-party library (SheetJS/XLSX) from a public CDN for spreadsheet parsing; everything else executes locally. Each viewer's data lives only in **their own browser** (via `localStorage`) — uploading data does not push it to a server or to other viewers.

### 2.2 Fleet / asset model
The system tracks four asset classes with distinct architectures:
1. **RC Drilling Rigs** — engine, hydraulics, track, rotation head, RC sample system, electrical; linked to a connected compressor.
2. **Diamond Drill Rigs** — engine, hydraulics, rotation head, bean pump, wireline winch, water pump, rod handling, core barrel, mast & feed, electrical; tracks core recovery.
3. **External Compressors** — engine, air end, oil separator, cooling system, pressure valves, electrical; tracks CFM, working pressure, and the RC rig assigned.
4. **Light Vehicles** — engine, tyres, brakes, electrical, body/chassis; tracks odometer, fuel economy, cost/km, tyre tread, and statutory compliance expiries.

### 2.3 Operating environment
- Any modern desktop or mobile browser with internet access (needed for the CDN library and fonts, and for the live Google Sheet option).
- Deployable as a single static file to Netlify Drop, Cloudflare Pages, GitHub Pages, or any static host.

### 2.4 Constraints
- Single self-contained file; no server, database, or build pipeline.
- Read-only with respect to source data; no write-back.
- Google-Sheet mode requires the sheet to be shared as viewable (gviz CSV export endpoint).
- Spreadsheet **tab names must match exactly**: `RC_Rigs, DD_Rigs, Compressors, LightVehicles, WorkOrders, Spares, KPITargets, Technicians`.

### 2.5 Assumptions & dependencies
- Column headers in each tab match the schema in Section 4.
- SheetJS is reachable from `cdnjs.cloudflare.com`.
- Dates are supplied as either spreadsheet dates or `DD Mon YYYY` strings.

---

## 3. Functional requirements

### 3.1 Data ingestion & session
- **FR-1 Setup screen.** On first load, show a setup modal offering: (a) upload a file, (b) use demo data, (c) advanced — connect a live Google Sheet by ID.
- **FR-2 File upload.** Accept one `.xlsx`/`.xls` workbook containing the 8 tabs, or one or more `.csv` files named after each tab, via click-to-choose or drag-and-drop. Parse with SheetJS, normalise headers/values, coerce dates to `DD Mon YYYY`, and map tabs to the internal model.
- **FR-3 Partial-load reporting.** Report how many of the 8 tabs were loaded and name any missing/empty tabs; proceed with what is present.
- **FR-4 Demo data.** Provide a complete built-in demo dataset covering all asset types so the dashboard is fully explorable without any file.
- **FR-5 Live Google Sheet.** Given a Spreadsheet ID, fetch each tab via the gviz CSV endpoint, parse to objects, and render.
- **FR-6 Auto-refresh.** In live-sheet mode, refresh data automatically on an interval (currently every 5 minutes); provide a manual refresh control.
- **FR-7 Persistence.** Remember the last uploaded dataset, Spreadsheet ID, and theme in the browser so a refresh restores the session. Data remains private to that browser.
- **FR-8 Data-source indicator.** Show current source (DEMO / FILE / live), a status dot, and a "last refreshed" timestamp.

### 3.2 Global UI
- **FR-9 Tabbed navigation** across 14 views: Executive, Overview, RC Rigs, DD Rigs, Compressors, Light Vehicles, Work Orders, PM Calendar, KPIs, Analytics, Technicians, Spares & Cost, ROI, Shift Report.
- **FR-10 Alert pills** in the header for: assets DOWN/BREAKDOWN, PM overdue count, spares stock-outs, and LV compliance expiries.
- **FR-11 Notifications panel** summarising active alerts with a count badge.
- **FR-12 Theme toggle** (dark / light), persisted.
- **FR-13 Live clock** and current date in top bar.
- **FR-14 Print support** — a print-optimised layout (Executive "Print Report", hides chrome).

### 3.3 Executive & Overview
- **FR-15 Executive summary.** Fleet-wide headline KPIs, an all-asset health strip, and auto-generated Critical Issues, Achievements, and Top Recommendations.
- **FR-16 Overview.** Fleet KPI tiles, all-assets list, top downtime causes (Pareto), active alerts, and trend sparklines for Availability %, Avg MTBF, and Reactive WO %.

### 3.4 Asset views (RC / DD / Compressors / Light Vehicles)
- **FR-17 Per-class summary tiles** (counts by status, availability, etc.).
- **FR-18 Asset cards** with status colour, availability, health, and PM-due indicator; selectable to open a detail panel.
- **FR-19 Asset detail** showing per-system health scores (1–6 with Excellent→Critical labels), hours/odometer, last/next PM, hours-to-next-PM (negative = overdue), open WOs, and class-specific metrics (metres drilled, shift target, core recovery, CFM/pressure, fuel economy, cost/hour or cost/km).
- **FR-20 Cascade map (RC).** A live compressor→RC connection visual and table: green link = compressor supplying air, red = compressor down (cascade — rig cannot drill). Both assets flagged simultaneously.
- **FR-21 LV compliance.** Track and colour-code Roadworthy, Insurance, Fire Extinguisher, and First-Aid expiry dates as VALID / EXPIRING SOON (≤30 days) / EXPIRED, plus pre-start compliance and service-by-km.

### 3.5 Work orders
- **FR-22 WO summary** — counts of Open, In Progress, High Priority, Completed.
- **FR-23 WO table** with filters by Asset Type, WO Type (PM/CORRECTIVE/REACTIVE/PdM), Priority, and Status; clearable.
- **FR-24 Create WO.** A form that builds a new WO row and lets the user preview and **copy the row** to paste into the WorkOrders tab (no write-back).

### 3.6 PM calendar
- **FR-25 Monthly calendar** of scheduled PM events derived from next-PM dates, with month navigation and a "Today" control.

### 3.7 KPIs & analytics
- **FR-26 KPI board.** Each KPI shows value, target, trend, and RAG status (green/amber/red) against target.
- **FR-27 Analytics.** Failure-frequency ranking by system, a P-F interval review table (flagging intervals wider than the P-F window), and an RUL estimate per asset from average health score.

### 3.8 Technicians
- **FR-28 Technician board** — on/off shift counts, WOs completed/open, hours logged, specialisation, and a workload/utilisation bar per technician.

### 3.9 Spares & cost
- **FR-29 Spares inventory** — current vs min/max stock, ABC class, unit cost, stock value, supplier, lead time, and status (STOCK-OUT / BELOW MIN / OK); reorder alerts with a "Raise PO" action.
- **FR-30 Cost tracker** — budget vs actual per category, total variance, categories over budget, actual-vs-budget bars, estimated spend split by asset type, maintenance-type split (PM/Reactive/Corrective/PdM) from WO hours, cost-per-hour bar chart flagging assets above fleet average, and overrun alerts.

### 3.10 ROI
- **FR-31 ROI calculator.** Editable inputs (cost/hour of downtime, breakdowns/year, avg hours/event, projected % reduction, implementation cost, annual running cost). Compute downtime hours prevented, net annual savings, payback period, and first-year return, and render a written ROI justification narrative for Deeprock Ghana Limited.

### 3.11 Shift report
- **FR-32 Shift handover.** Shift KPI tiles, completed-this-shift WOs, outstanding/carry-forward WOs, end-of-shift fleet status across all assets, and an editable handover-notes field; print-friendly.

---

## 4. Data requirements (workbook schema)

The workbook must contain these 8 tabs, named exactly. Representative columns (from the delivered model):

| Tab | Key columns |
|-----|-------------|
| **RC_Rigs** | ID, Name, Site, Hours, Availability, Status¹, Engine, Hydraulics, Track, RotationHead, RCSampleSystem, Electrical, ConnectedCompressor, LastPM, NextPM, HrsToNextPM, OpenWOs, MetresDrilled, ShiftTarget, CostPerHour |
| **DD_Rigs** | ID, Name, Site, Hours, Availability, Status¹, Engine, Hydraulics, RotationHead, BeanPump, WirelineWinch, WaterPump, RodHandling, CoreBarrel, MastFeed, Electrical, LastPM, NextPM, HrsToNextPM, OpenWOs, MetresDrilled, ShiftTarget, CostPerHour, CoreRecovery |
| **Compressors** | ID, Name, Make, Model, Site, AssignedTo, Hours, Availability, Status², Engine, AirEnd, OilSeparator, CoolingSystem, PressureValves, Electrical, LastPM, NextPM, HrsToNextPM, OpenWOs, CFM, WorkingPressure, CostPerHour |
| **LightVehicles** | ID, Name, Type, Make, Model, Odometer, Availability, Status³, Engine, Tyres, Brakes, Electrical, BodyChassis, Driver, Site, LastService, NextServiceKm, KmToNextService, RoadworthyExpiry, InsuranceExpiry, FireExtExpiry, FirstAidExpiry, PreStartCompliance, FuelL100km, CostPerKm, TyreTreadMm, OpenWOs |
| **WorkOrders** | ID, AssetID, AssetType⁴, Type⁵, Priority⁶, System, Description, Technician, RaisedDate, DueDate, Hours, Status⁷, RootCause, DowntimeCost |
| **Spares** | Name, PartNumber, Class (A/B/…), CurrentStock, MinStock, MaxStock, Status⁸, UnitCost, Supplier, LeadTimeDays |
| **KPITargets** | Metric, Value, Target, Trend, Status (green/amber/red), Period |
| **Technicians** | Name, Role, WOsCompleted, WOsOpen, HoursLogged, Specialisation, Status (ON SHIFT/OFF SHIFT) |

**Controlled values**
1. RC/DD Status: `DRILLING`, `BREAKDOWN`, `PM DUE`, `STANDBY`
2. Compressor Status: `RUNNING`, `DOWN`, `PM DUE`, `STANDBY`
3. LV Status: `AVAILABLE`, `ON ROAD`, `IN WORKSHOP`
4. WO AssetType: `RC`, `DD`, `COMPRESSOR` (and LV where applicable)
5. WO Type: `PM`, `CORRECTIVE`, `REACTIVE`, `PdM`
6. Priority: `HIGH`, `MEDIUM`, `LOW`
7. WO Status: `OPEN`, `IN PROGRESS`, `COMPLETED`
8. Spares Status: `STOCK-OUT`, `BELOW MIN`, `OK`

- Health scores (Engine, Hydraulics, etc.): integer **1–6** (1 Excellent → 6 Critical).
- `HrsToNextPM` / `KmToNextService`: negative value denotes overdue.
- `ConnectedCompressor` / `AssignedTo` link an RC rig and its compressor (e.g. `DGL-COMP01`).

---

## 5. Non-functional requirements

- **NFR-1 Portability.** Ship as one static file deployable to any static host; no server or build step.
- **NFR-2 Zero install for viewers.** Team members open a URL; no account, no software.
- **NFR-3 Privacy.** Each viewer's data stays in their own browser; no data is transmitted to a server or shared between viewers. (For a single shared always-current dataset, the live Google Sheet option is used.)
- **NFR-4 Offline-capable core.** After first load, file/demo modes work without network except for the CDN library and fonts.
- **NFR-5 Performance.** Views render near-instantly for a small-to-medium fleet (tens of assets, low-hundreds of work orders).
- **NFR-6 Usability.** Consistent tabbed layout, colour-coded status/health, dark/light themes, mobile-responsive, and print-friendly reports.
- **NFR-7 Resilience.** Handle missing tabs, blank cells, and malformed dates gracefully; surface clear errors for failed sheet connections.
- **NFR-8 Maintainability.** Data model driven by named tabs and a tab→model map, so schema-conformant sheets require no code changes.
- **NFR-9 Branding.** Presented for Deeprock Ghana Limited (name, colours, asset naming e.g. `DGL-RC01`).

---

## 6. Out of scope / non-goals
- No write-back to Google Sheets or Excel (WO creation is copy-to-paste only).
- No user authentication, roles, or audit trail.
- No multi-user shared state beyond a common live Google Sheet.
- No automated procurement (the "Raise PO" control is a placeholder action).
- No server-side storage, history, or backups beyond per-browser `localStorage`.

## 7. Future considerations (candidate enhancements)
- Optional backend for shared, writable data and true real-time collaboration.
- Direct WO write-back to the source sheet.
- Role-based access and change history.
- Configurable KPI targets and PM intervals in-app.
- Export of Executive / Shift reports to PDF.

---

## 8. Traceability — modules to source
All requirements above are realised in the single delivered file `index.html` (identical working copy: `SMARTIS Deeprock Ghana v4.html`). Deployment options are described in `HOW_TO_DEPLOY.md`; a schema-correct starter workbook is `SMARTIS Demo Data Template.xlsx`.
