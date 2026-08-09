# Retail Operations Audit & Performance Tracking System

Automated, multi-system audit and performance tracking platform built for a growing multi-store fashion retail chain — replacing manual, paper-based checklists and visit logs with live, self-updating dashboards.

## Overview

Retail field operations — store audits, visual merchandising compliance, and daily opening/closing checklists — were tracked manually across spreadsheets with no real-time visibility, no automated scoring, and no centralized reporting. This project replaces that with three integrated systems, each built on a shared modular architecture, feeding live dashboards that rebuild automatically on every form submission.

## Systems

| System | Purpose |
|---|---|
| **AM Monthly Checklist Scorecard** | Area Manager store audit scoring, tracked monthly across all stores |
| **VM Store Audit** | Visual merchandising adherence scoring (1–5 scale), rolled up into per-area and per-store ratings |
| **SM Daily Opening/Closing Checklist** | Daily store compliance checklist with branching Opening/Closing flows and photo evidence uploads |

## Architecture

```
Google Forms → Google Sheets → Google Apps Script → Live Dashboard
```

Each system shares the same modular file structure:

- `Config.gs` — store list, sheet names, checklist definitions — single source of truth
- `FormBuilder.gs` — programmatic form generation (page-per-section, branching logic for Opening/Closing)
- `SubmitHandler.gs` — parses submissions into Master Log, Issue Log, and (where applicable) Photo/Score Logs
- `Dashboard.gs` — builds and refreshes a live dashboard: KPI cards, trend charts, heat maps, store deep-dive drill-through
- `Export.gs` — one-click PDF export of the live dashboard to Drive
- `MasterData.gs` — reference sheets (Store, Employee, Checklist, Settings) for future filtering and weighting
- `Branding.gs` — applies org logo to Form and Dashboard
- `Menu.gs` — custom UI menu with setup, refresh, testing, and data-repair tools
- `Setup.gs` / `Triggers.gs` — one-time wiring and automated triggers (on-submit + daily refresh)

Dashboards rebuild automatically via `onFormSubmit` and scheduled triggers — no manual reporting required.

## Key Features

- **Automated dashboard refresh** on form submission and manual/scheduled triggers
- **Custom tools menu** per system (Setup, Refresh, Seed Test Data, Reset, Rebuild reference sheets)
- **Config-driven store list** — adding a store is a one-line config change, no code edits
- **Heat map + drill-through dashboards** — failure-by-section heat maps, per-store deep dives, top recurring issues
- **Branching form logic** (SM system) — a single form dynamically routes Opening vs. Closing responders through different question sets
- **Photo evidence capture** (SM system) — Drive-link-based photo uploads tied to specific checklist items
- **One-click PDF export** straight from the dashboard to Drive
- **Data-repair tooling** — utilities to audit form configuration, rebuild derived logs, and safely remove individual bad records without leaving orphaned data

## Tech Stack

Google Forms · Google Sheets · Google Apps Script · Google Drive API

## Scalability

Every system is config-driven — store codes, thresholds, and settings live entirely in `Config.gs`. Adding a new store requires only a config entry, not a code change. Currently deployed across 10 stores (scaling to 18–20 within 2026), with the architecture designed to handle 50+ stores without structural rework.

## Impact

- Digitized three previously manual, paper-based tracking processes across store operations
- Eliminated manual dashboard compilation — real-time visibility for management
- Reduced reporting lag from days to instant, on-submission updates
- Standardized scoring and issue tracking across audit types (checklist adherence, VM ratings, daily compliance)

## Screenshots

### AM Monthly Checklist Dashboard
_Area Manager store audit scorecard — KPIs, score trends, and heat map by section._


![image alt](https://github.com/Ayan-baksi/Retail-Ops-Automation-Suite/blob/main/AM-Monthly-Checklist/AM_MASTER%20DASHBAORD.png?raw=true)

### VM Store Audit Dashboard
_Visual merchandising adherence tracking — area-wise scores, store ranking, and low-score item tracking._

![image alt](https://github.com/Ayan-baksi/Retail-Ops-Automation-Suite/blob/main/VM-CHECKLIST/VM_MASTER%20DASHBAORD.png?raw=true)

### SM Daily Checklist Dashboard
_Daily Opening/Closing compliance — real-time completion status, photo evidence log, and issue tracking._

![image_alt](https://github.com/Ayan-baksi/Retail-Ops-Automation-Suite/blob/main/SM-DAILY_CHECKLIST/SM_MASTER%20DASHBAORD.png?raw=true)

## Notes

Sample data included is dummy/anonymized. Store identifiers and real operational data have been replaced with placeholder values (e.g. `STORE-001`) for confidentiality.
