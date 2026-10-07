# PSH Project

GitHub development home for the NWS WFO LIX Post-Tropical Cyclone Report (PSH) automation.

## Production model

**GitHub is the source of truth. Google Apps Script remains the operational runtime.**

The deployable artifact is still one file:

`dist/PSH_Automation.gs`

Paste that file into the Apps Script project bound to the PSH Google Sheet. No `clasp` deployment is required.

## Current baseline

- **PSH Automation v0.16**
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
