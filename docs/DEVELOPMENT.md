# Development workflow

## Source of truth

Production is assembled from:

- `src/core/rainfall_core.js` — pure rainfall logic
- `src/PSH_Automation.gs` — Apps Script adapter/runtime

Do not hand-edit `dist/PSH_Automation.gs`.

Build it with:

```bash
npm run build
```

## Required change cycle

1. Change pure rainfall behavior in `src/core/rainfall_core.js` when possible.
2. Keep Google-specific reads/writes and HTTP transport in `src/PSH_Automation.gs`.
3. Add or update deterministic fixtures when a real regression case is discovered.
4. Run `npm test`.
5. Run `npm run build`.
6. Confirm CI is green.
7. Only then paste `dist/PSH_Automation.gs` into the bound testing Sheet for final acceptance.

The Sheet is no longer the primary debugger for rainfall logic.

## Deterministic Francine rainfall regression

`tests/rainfall-core.test.mjs` executes the same pure rainfall functions used by production.

The source fixture is:

`fixtures/api/francine/rainfall_sources.json`

with IEM daily data in:

`fixtures/api/francine/iem_cocorahs_daily.csv`

The regression protects the source-selection cases that are actually automatable and documents the remaining manual exception:

- `LA-JF-20 = 9.48` remains an issued reference/manual historical exception; current live official/IEM sources do not defensibly reproduce it.
- `LA-SC-06 = 9.22` from official CoCoRaHS despite zero-padding differences and a lower IEM mirror value.
- `LIX = 7.93` from exact-window Synoptic even though ACIS is deliberately set to 4.33.

If any of those source-selection rules change, `npm test` fails before Apps Script deployment.

## Live integration

`npm run test:live` checks current public external behavior.

The scheduled/manual GitHub workflow:

`.github/workflows/live-integration.yml`

verifies:

- IEM daily CoCoRaHS CSV still parses; the Francine `LA-JF-20` row is informational because the current live endpoint does not return it.
- IEM's `LA_COCORAHS` network catalog still contains `LA-JF-20`.
- Synoptic LIX is checked when the optional repository secret `SYNOPTIC_TOKEN` exists.

Live integration is separated from deterministic CI because external services can be unavailable even when the code is correct.

## Main CI

`.github/workflows/ci.yml` runs `npm test` on every push and pull request.

That includes:

- executable rainfall core regression
- production adapter/core linkage checks
- generated `dist` parity
- JavaScript syntax validation
- credential-string guardrails
- exact XLSX fixture integrity
- WeatherSTEM Francine workbook regression checks

## What still requires the bound Sheet

Only Google-runtime acceptance behavior:

- `SpreadsheetApp`
- `PropertiesService`
- menu wiring
- formatting/rich links/formula preservation
- Apps Script `UrlFetchApp` execution
- final end-to-end workbook population

## Reference authority

`fixtures/PSHLIX_2024AL06_Francine_Data.xlsx` remains the authoritative issued-product reference.

`tests/francine_expected.json` and API fixtures are executable regression representations of that authority. If they conflict with the issued workbook, fix the fixture/test rather than redefining the issued result.

## Current known exception

Tiger Stadium WeatherSTEM gust remains an intentional known discrepancy: current direct retrieval is about 44.32 kt versus the issued 48 kt. Do not loosen the tolerance simply to make that test pass.


## Bertha out-of-sample acceptance

Tropical Storm Bertha (2026AL02) is the second historical reference and should be treated as an out-of-sample check, not as a source of station-specific production hacks.

In the bound testing Sheet:

1. Paste the current `dist/PSH_Automation.gs`.
2. Reload the Sheet.
3. Run **PSH Automation → Load Bertha Test Window**.
4. Run **Run EVERYTHING**.
5. Run **Run Bertha Regression Check**.
6. Review `REGRESSION-BERTHA` and `REGRESSION-BERTHA-SUMMARY` in `_PSH_Log`.

The reference rainfall outcome is zero stations at or above 3 inches. Sub-3-inch populated template rows are logged for review but do not fail this regression by themselves.
