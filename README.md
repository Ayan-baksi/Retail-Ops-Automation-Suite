# Retail Operations Audit & Performance Tracking System

Automated, multi-system audit and performance tracking platform built for a growing multi-store fashion retail chain — replacing manual, paper-based checklists and visit logs with a live, self-updating reporting layer.

## Overview

Retail field operations (store audits, journey planning, opening/closing compliance) were tracked manually across spreadsheets with no real-time visibility. This project replaces that with four integrated systems, each following a shared modular architecture, feeding live dashboards that rebuild automatically on every form submission.

## Systems

| System | Purpose |
|---|---|
| **AM Monthly Checklist Scorecard** | Area Manager store audit scoring, tracked monthly |
| **PJP (Permanent Journey Plan) Visit Logger** | Logs and tracks scheduled field visits against plan |
| **SM Daily Opening/Closing Checklist** | Daily store compliance checklist for opening and closing procedures |
| **VM Store Audit** | Visual merchandising adherence scoring (1–5 scale) |

## Architecture

Google Forms → Google Sheets → Google Apps Script → Live Dashboard

Each system is built on a shared modular script structure:
- `Config.gs` — store list, sheet names, checklist definitions — single source of truth
- `FormBuilder.gs` — programmatic form generation (one page per section)
- `SubmitHandler.gs` — parses submissions into a Master Log + Issue Log
- `Dashboard.gs` — builds/refreshes a Power BI-style dashboard: KPI cards, trend charts, heat maps, store deep-dive drill-through
- `Export.gs` — one-click PDF export of the live dashboard to Drive
- `MasterData.gs` — reference sheets (Store, Employee, Checklist, Settings) for future filtering/weighting
- `Branding.gs` — applies org logo to Form + Dashboard
- `Setup.gs` — single `setup()` function that wires everything together

Dashboards rebuild automatically via `onFormSubmit` and manual refresh triggers — no manual reporting required.

## Key Features

- **Automated dashboard refresh** on form submission and manual refresh trigger
- **Custom "AM Tools" menu** for setup, refresh, and reset actions
- **Config-driven store list** — adding a store is a one-line config change, no code edits
- **Heat map + drill-through dashboard** — failure-by-section heat map, per-store deep dive, top recurring issues
- **One-click PDF export** straight from the dashboard to Drive
- **Built-in resilience** — handles missing data, permission errors, and stray UI elements gracefully (see inline comments for real debugging notes)

## Tech Stack

Google Forms · Google Sheets · Google Apps Script · Google Drive API

## Scalability

The system is config-driven — store codes, thresholds, and settings live entirely in `Config.gs`. Adding a new store requires only a config entry, not a code change. Currently deployed across 10 stores (scaling to 18-20 within 2026), with the architecture designed to handle 50+ stores without structural rework.

## Impact

- Digitized previously manual, paper-based tracking processes across store operations
- Eliminated manual dashboard compilation — real-time visibility for management
- Reduced reporting lag from days to instant, on-submission updates

## Notes

Sample data included is dummy/anonymized. Store identifiers and real operational data have been replaced with placeholder values (e.g. `STORE-001`) for confidentiality.
