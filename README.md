# PSH Project

GitHub development home for the NWS WFO LIX Post-Tropical Cyclone Report (PSH) automation.

## Production model

**GitHub is the source of truth. Google Apps Script remains the operational runtime.**

The deployable artifact is still one file:

`dist/PSH_Automation.gs`

Paste that file into the Apps Script project bound to the PSH Google Sheet. No `clasp` deployment is required.

## Current baseline

- **PSH Automation v0.18**
- Apps Script V8 runtime
- Bound-spreadsheet design
- Hurricane Francine (2024AL06) is the primary regression/reference case
- Conservative QC is intentional: a defensible blank is preferred over a guessed value

## Testing architecture

v0.15 moves rainfall matching, source precedence, CoCoRaHS ID canonicalization, Synoptic alias matching, IEM CSV parsing, and deterministic source resolution into a pure JavaScript module:

`src/core/rainfall_core.js`

Bertha (2026AL02) is now the second, out-of-sample reference case. Its regression deliberately covers weaker winds, marine AWOS/CMAN/WLON, WeatherSTEM, zero reportable rainfall stations, and NOAA/USGS water while preserving manual WeatherFlow/CPRA/USACE/TPCG behavior.

That same module is:

1. concatenated into the production Apps Script artifact by `npm run build`, and
2. executed directly by GitHub tests.

The deterministic Francine rainfall fixture reproduces the automatable edge cases that exposed the earlier test gap:

- `LA-JF-20` remains an issued **9.48 in** reference but is a documented manual historical exception because current live official/IEM sources do not return a defensible automated value.
- `LA-SC-06` must canonicalize to `LA-SC-6` while still preferring the official CoCoRaHS value **9.22 in**.
- `LIX` has an intentionally wrong ACIS fallback value of **4.33 in**, while the exact-window Synoptic bulk fixture contains the issued **7.93 in**. CI fails if ACIS wins.

Run:

```bash
npm test
```

That now executes the Francine rainfall core regression, Apps Script build/syntax guardrails, and committed workbook fixture checks.

A separate live integration workflow tests current external API behavior against IEM and, when a repository `SYNOPTIC_TOKEN` secret is configured, Synoptic:

`.github/workflows/live-integration.yml`

## Repository layout

```
src/
  core/
    rainfall_core.js       pure/testable rainfall logic used by production
  PSH_Automation.gs        Apps Script adapter/runtime

dist/
  PSH_Automation.gs        generated paste-ready Apps Script artifact

tests/
  rainfall-core.test.mjs   executable deterministic Francine rainfall regression
  live-rainfall.mjs        live external API checks
  smoke.mjs                build/syntax/adapter guardrails
  francine_expected.json   issued Francine regression references
  bertha_expected.json     issued Bertha out-of-sample references

fixtures/
  api/francine/
    rainfall_sources.json
    iem_cocorahs_daily.csv
  PSHLIX_testingspreadsheet.xlsx
  PSHLIX_2024AL06_Francine_Data.xlsx
  source_manifest.json

scripts/
  build.mjs
  inspect-fixtures.mjs

.github/workflows/
  ci.yml
  live-integration.yml
```

## Build

```bash
npm run build
```

The build concatenates the tested pure rainfall core with the Apps Script adapter and writes one deployable file to `dist/PSH_Automation.gs`.

## What still requires Google Apps Script

GitHub now tests the core rainfall behavior before deployment. The bound Sheet remains the final acceptance environment for things GitHub cannot reproduce exactly:

- `SpreadsheetApp`
- `PropertiesService`
- bound-sheet menus
- rich-text links and formatting
- formula preservation after writes
- Apps Script-specific `UrlFetchApp` behavior

Those checks should catch runtime/integration issues, not be the first place core rainfall bugs are discovered.

## Data sources

Automated sources include Synoptic, IEM, WeatherSTEM, CoCoRaHS, ACIS, NOAA CO-OPS, and USGS Water Data.

WeatherFlow historical observations and unverified USACE/CPRA/TPCG water mappings remain intentionally manual.

## Design rule

Do not weaken a QC rule merely to improve regression score. If the automation cannot defend station identity, variable semantics, datum, or historical conversion, keep the field blank and log why.

## Bertha independent API findings (2026-10-06)

Live IEM ASOS archive queries, independently executed in GitHub Actions, establish:

- KHSA 2026-07-23 16:47Z reports an internally inconsistent 30-kt sustained wind and 15-kt gust. Individual-report QC now rejects that sustained reading; the highest remaining reported sustained wind is 20 kt (and a 30-kt gust survives).
- KMSY and KNEW maximum available IEM sustained winds were 21 kt and 32 kt respectively, compared with issued 23/34 kt. The archive result remains authoritative for what it actually contains; no upward adjustment is made to match historical issue values.
- Direct NAVD88 USGS series (63160/62020/62615) are eligible. Conversion from historical 00065 stage to NAVD88 using a present-day site altitude is disallowed pending event-effective datum metadata. This may create justified blanks where older versions populated estimates.
- Bertha regression's BPPL1 must-be-blank assertion was removed: whether to populate it depends on the verified source parameter, not the station name.
- Missing CMAN/WLON sea-level pressures remain missing until a verified MSLP source is provided. Do not substitute station pressure.

Executable tests: `tests/airport-wind.test.mjs`, `tests/usgs-datum.test.mjs`. Independent source inspection is retained under `scripts/audit-bertha-iem.mjs` with a manual GitHub Actions workflow.

## Muse audit follow-up, v0.19

- USGS 7-digit spreadsheet IDs are normalized with a leading zero; both numeric and USGS-prefixed URL locators are supported.
- Direct NAVD88 observation retrieval is preserved. A station ID alone never forces a blanket blank.
- For allowlisted stage-only sites, an explicit USGS 00065 existence check distinguishes missing stage observations from missing **verified historical** gage-datum elevation. Neither produces an unvalidated NAVD88 value.
- The modern monitoring-location altitude is **not** accepted as a historical gage-zero elevation; the obsolete helper was removed. The altitude-datums reference list defines datum systems, not a site-specific, event-effective gage-zero offset.
- The noisy COOP-prefixed Synoptic retry is removed; the canonical request, exact-ID bulk recovery, and ACIS fallback remain intact.
- Water QC uses tagged `[AUTO-QC]` annotations that are removed/replaced on subsequent runs while preserving bare human-entered I/E and remarks. Old bare I/E flags cannot be retrospectively classified safely.
- Selected airport per-field source/time provenance is logged, and equality of sustained wind and gust at >=25 kt is flagged for review without overriding or deleting observations.
- MSLP-only, WeatherSTEM sustained-wind, CoCoRaHS legacy-ID, KMSY and KNEW selection logic were not modified.

Stage-to-NAVD88 conversion remains intentionally unavailable until a *site-specific elevation of gage zero valid at the event date* is independently verified, in NAVD88. This audit therefore improves source correctness and diagnostics rather than promising to recreate values where defensible historical metadata have not been found.

## Accuracy-first historical acceptance (v0.20)

After **Run Francine Regression Check** or **Run Bertha Regression Check**, review **`_PSH_Review`** first. Populated numeric agreement is the fraction of compared nonblank automated readings that fall inside the original reference tolerance. Missing values are reported separately, never counted as wrong measurements. The review sheet ranks populated differences by multiples of tolerance and provides the station, observation field, automated/issued values, difference, QC notes, and station source URL. A mismatch is a research/review flag, not proof that the issued report or archived source is wrong. The old combined completion score remains in the log for historical continuity but is explicitly *not* called accuracy. This update changes no observation retrieval, QC tolerances, or issuance data values.
