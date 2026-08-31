# Marketing Events TAL Summary — SPEC

**Status:** Locked business logic (2026-08-31). Implement against this document.

Standalone tool: read the operator’s **Marketing Events Registry**, pull
HubSpot attendance / unique accounts / TAL, write a two-sheet **TAL Accounts
@ Marketing Events** workbook.

Reference look-and-feel: the operator’s `TAL Accounts @ Marketing Events.xlsx`
template (sheet names, columns, fonts, fills, formulas, widths).

---

## Goal

One command:

1. Accept a Marketing Events Registry `.xlsx`.
2. Build output columns **A–C** from the Registry tab (**A**, **B**, **G**).
3. Fill **D–G** from HubSpot for every event row.
4. Write a **new** two-sheet `.xlsx`. Never overwrite the registry.

Operator loop: maintain the registry → drop it in → get the finished TAL
workbook.

---

## Entry point

Root-level script (name flexible at implement time; default
`tal_summary.py`):

```
python tal_summary.py path/to/Marketing_Events_Registry.xlsx [--out path/to/output.xlsx]
```

- Auth: `.env.local` with `HUBSPOT_TOKEN` (load explicitly at start).
- Default output: `outputs/YYYY-MM-DD/tal_accounts_marketing_events.xlsx`.
- `--out` optional. Non-`.xlsx` extensions rewritten to `.xlsx`.
- If resolved output path equals the input path → hard-stop (no overwrite).
- Console: per-event D/E/F, plus both union TOTALs.
- Any event with **zero** HubSpot contact matches → hard-stop: list every
  unmatched Events Attended Appendage, write **nothing**, exit nonzero.

Dependencies: Python 3, `openpyxl`, `requests`, `python-dotenv`.

---

## Input: Marketing Events Registry

Workbook with (at least) a sheet named **`Registry`**.

| Registry column | Header (exact) | Role |
| :--- | :--- | :--- |
| A | `Category` | → output A |
| B | `Sub-category` | → output B (display only; never used for match or aggregation) |
| G | `Events Attended Appendage` | → output C **and** HubSpot match key |

Ignore C–F and H–I for this report (List Name, List ID, Event Type, Event
Date, Lead Source, Lead Source Description). Ignore other sheets
(`Naming Convention`, `Lead Source Dropdowns`, …).

### Row kinds on Registry

| Kind | Recognition | Behavior |
| :--- | :--- | :--- |
| Event | G non-empty (after strip) | Include on sheet 1; HubSpot fill D–F; % formula in G |
| Blank separator | A/B/G all empty, appearing **between** event blocks | Copy through on sheet 1 as a blank row |
| Trailing empties | Empty rows after the last event | **Omit** — do not pad the output |
| TOTAL | — | **Not in the registry.** Script appends TOTAL rows |

All categories are in scope (Channel Partner, CISO Society, Trade Shows,
CyAlliance, Happy Hours, IANS, …). No Event Type filter, no date filter,
no list-membership filter.

Display strings for A/B/C stay as stored in the registry (em-dashes,
punctuation, Sub-category placeholders, etc.).

---

## Matching (HubSpot)

Output C / Registry G values **are** the HubSpot contact property
`events_attended` tokens.

Equality after normalizing **both** sides:

1. Strip leading/trailing whitespace.
2. Collapse any run of whitespace to a single space.

HubSpot `events_attended` is `;`-delimited. Split on `;`, normalize each
token, then exact-token equality against the normalized registry appendage.

HubSpot search (`CONTAINS_TOKEN`) is a prefilter only. Client-side
normalized equality is the source of truth.

Zero matches after confirm → hard-stop for that run (not a written `0`).

---

## Metric definitions (output D–G)

### Per event

1. **D `# Contact Attendees`** — count of contact IDs whose normalized
   `events_attended` tokens include this event’s normalized appendage.
2. **E `# Unique Accounts`** — those contacts → **Primary Company** (CRM
   associations v4: `HUBSPOT_DEFINED`, `typeId == 1`) → count of distinct
   company IDs. A contact with no Primary Company counts in D and does
   **not** add to E.
3. **F `# Accounts on TAL`** — of that company set, count where
   `hs_is_target_account == "true"` (string). Blank and `"false"` do not
   count. Internal name is `hs_is_target_account`.
4. **G `% Accounts on TAL`** — **not** a Python value. Write formula
   `=F{row}/E{row}` with number format `0%` on data rows (matches the
   operator template).

### TOTAL — all events (sheet 1)

Pool every contact matched by **any** event row in this run (union):

- D = size of that contact-ID set (one person at three events counts once).
- E = distinct Primary Companies of that union.
- F = TAL count on that company union.
- G = same `=F/E` formula on the TOTAL row.

Expect TOTAL D ≤ sum of per-event D, and TOTAL E ≤ sum of per-event E.
That inequality is correct (union vs sum).

### TOTAL — Channel Partner Events (sheet 2)

Same union rules, restricted to rows whose Category is exactly
`Channel Partner Events`. No extra HubSpot round-trip — slice maps already
built for the full run.

---

## HubSpot properties used

| Object | Internal name | Role |
| :--- | :--- | :--- |
| Contact | `events_attended` | `;`-delimited tokens; selection key |
| Company | `hs_is_target_account` | TAL; only `"true"` counts |
| Association | contacts → companies v4 Primary (`typeId` 1) | Account identity for E / F |

**Not used:** list membership, List ID, `lead_source_description`, meetings,
deals, date bounds on attendance.

---

## Output workbook

Always a **new** `.xlsx`. Two sheets, columns A–G only (no used column H
content; no trailing empty data rows).

| Sheet name (exact) | Contents |
| :--- | :--- |
| `TAL Accounts @Events` | Header; every included registry row (events + mid-block blanks) with D–G filled on event rows; then **TOTAL** (all-events union) immediately after the last content row — **no** blank row before TOTAL |
| `TAL Accounts @Channel Events` | Header; only event rows with Category exactly `Channel Partner Events` (no category blank separators required); then **one blank row**; then **TOTAL** (channel union) |

### Output headers (row 1, exact)

`Category` | `Sub-category` | `Events Attended` | `# Contact Attendees` | `# Unique Accounts` | `# Accounts on TAL` | `% Accounts on TAL`

Note: output column C is labeled **`Events Attended`**, even though the
registry source column is **`Events Attended Appendage`**.

### Formatting (match operator template)

| Element | Rule |
| :--- | :--- |
| Font | Arial throughout |
| Header | Bold; fill `C9DAF8` on A–G |
| Data D–F | Integers; center-aligned |
| Data G | Formula `=F{row}/E{row}`; format `0%`; center-aligned |
| TOTAL label (C) | Bold; fill `FFF2CC`; right-aligned; text `TOTAL` |
| TOTAL D–F | Bold; fill `FFF2CC`; center-aligned |
| TOTAL G | Bold; fill `D9D2E9`; formula `=F{row}/E{row}`; format `0.00%`; center-aligned |
| Column widths | A `18.75`, B `24.38`, C `55.5`, D `17.13`, E `16.13`, F `22.5`, G `16.63` |

Identity cells (A–C) on event rows: plain Arial, no special fill.

---

## Suggested fetch shape (implementer’s guide)

Not normative packaging — efficiency target:

1. Read Registry → list of event appendages (unique, order preserved) +
   full row plan for sheet 1 / sheet 2.
2. One HubSpot contact search over those appendages (normalized confirm).
3. Index contacts by original appendage string.
4. One Primary Company map + one TAL company read over the contact union.
5. Per-event and TOTAL metrics by slicing those maps (no per-event
   re-fetch of companies).

Build HubSpot helpers inside this repo (auth, search, Primary Company,
batch company read, retries). Keep the surface small — only what this
report needs.

---

## Caveats

- `hs_is_target_account` is sparsely populated. F measures what is flagged
  today, not confirmed off-TAL.
- Corporate-family fragmentation: separate Company records inflate E.
- After whitespace normalization, remaining mismatches hard-stop (typos,
  punctuation drift, event never tagged in HubSpot).
- No date filter: a contact tagged for an event counts regardless of when
  they were tagged.
- TOTAL attendees are a **union**. Do not expect them to equal the sum of
  per-event attendees.

---

## Out of scope

- Intro Demos, deals, pipeline, revenue
- Cold Outreach / list `membershipTimestamp`
- Writing back to the Marketing Events Registry
- Category / Sub-category rollups beyond the Channel Partner sheet
- Creating or backfilling HubSpot `events_attended` tokens

---

## Verification

1. Run against a current Marketing Events Registry export.
2. Spot-check 1–2 events in the HubSpot UI (contact search on
   `events_attended`) — D should match the contact count; E/F should be
   explainable from Primary Company + TAL.
3. Confirm each sheet’s TOTAL D ≤ sum of that sheet’s per-event D (same for E).
4. Confirm a known non-TAL company does not inflate F.
5. A registry G value with extra internal spaces still matches; a truly
   unknown G hard-stops and writes no file.
6. Open the output next to the operator template: sheet names, headers,
   fills, TOTAL placement (sheet 1: no blank before TOTAL; sheet 2: blank
   then TOTAL), and `%` formulas.
