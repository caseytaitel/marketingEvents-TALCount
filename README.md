# Marketing Events TAL Summary

Reads a **Marketing Events Registry** workbook and writes **TAL Accounts @
Marketing Events** (contact attendees, unique accounts, TAL coverage) from
HubSpot.

Business logic is locked in [`docs/SPEC.md`](docs/SPEC.md).

## Setup

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env.local
```

Set `HUBSPOT_TOKEN` in `.env.local` (HubSpot private app token with CRM read
on contacts — including `events_attended` — companies — including
`hs_is_target_account` — and contact–company associations).

## Run

Drop the Marketing Events Registry at the repo root as **`INPUT.xlsx`**, then:

```
python tal_summary.py
```

Default output: `outputs/YYYY-MM-DD/tal_accounts_marketing_events.xlsx` (the dated folder is created on a successful run).

A different registry path or output path still works:

```
python tal_summary.py path\to\other-registry.xlsx --out path\to\custom.xlsx
```

The registry file is never overwritten. Any event with zero HubSpot matches
hard-stops, lists the unmatched Events Attended Appendage values, and writes
nothing.

## Verify

1. Spot-check 1–2 events in HubSpot (contact search on `events_attended`): D
   should match the contact count; E/F should follow Primary Company + TAL.
2. On each sheet, TOTAL D ≤ sum of that sheet’s per-event D (same for E).
   Totals are unions, not sums.
3. A known non-TAL company must not inflate F.
4. Extra internal spaces in registry G still match; a truly unknown G
   hard-stops with no output file.
5. Open the output next to the operator template: sheet names, headers,
   fills, TOTAL placement (sheet 1: no blank before TOTAL; sheet 2: blank
   then TOTAL), and `%` formulas.
