# PSH Project

GitHub development home for the NWS WFO LIX Post-Tropical Cyclone Report (PSH) Google Apps Script automation.

## Production model

**GitHub is the source of truth. Google Apps Script remains the operational runtime.**

Development, review, history, and lightweight automated checks happen here. The finished production artifact is:

`dist/PSH_Automation.gs`

Download that file and paste it into the Apps Script project bound to the PSH Google Sheet.

No `clasp` deployment is required.

## Current baseline

- Current development baseline: **PSH Automation v0.13**
- Apps Script V8 runtime
- Bound-spreadsheet design
- Hurricane Francine (2024AL06) is the primary regression/reference case
- Conservative QC is intentional: a defensible blank is preferred over a guessed value

### Automated data sources

- Synoptic
- IEM METAR/HFMETAR archive
- WeatherSTEM direct station metadata + historical minute endpoint
- CoCoRaHS
- ACIS
- NOAA CO-OPS
- USGS Water Data

### Intentionally manual / not yet automated

- WeatherFlow historical observations
- USACE / LA CPRA / TPCG water levels until mappings/API behavior are verified
- Tornado narratives / EF ratings
- Inland flooding narratives
- Impact narratives
- Event summary narrative

## Repository layout

```
src/
  PSH_Automation.gs        canonical Apps Script source

dist/
  PSH_Automation.gs        paste/download this into Google Apps Script

tests/
  smoke.mjs                static/syntax guardrails
  francine_expected.json   issued-PSH reference values derived from Francine

fixtures/
  PSHLIX_testingspreadsheet.xlsx
  PSHLIX_2024AL06_Francine_Data.xlsx
  source_manifest.json     fingerprints of the committed workbook fixtures
  README.md                fixture authority + workflow

scripts/
  build.mjs                creates dist/ from src/
  inspect-fixtures.mjs     validates and inspects committed XLSX fixtures

docs/
  DEVELOPMENT.md           development + validation workflow
  WEATHERSTEM_FRANCINE_VALIDATION.md
                           v0.10 WeatherSTEM regression validation

.github/workflows/
  ci.yml                   GitHub Actions smoke checks
```

## Local development

Requires Node.js 20+ only for build/smoke checks. The production script itself runs in Google Apps Script and has no Node dependency.

```bash
npm install
npm test
npm run build
```

`npm run build` currently copies the single canonical `.gs` source into `dist/`. This intentionally leaves room to split the source into modules later while continuing to emit one paste-ready Apps Script file.

## Operational validation

`npm test` checks source/dist parity, JavaScript syntax/static guardrails, and the committed XLSX fixture fingerprints/structure.

`fixtures/PSHLIX_2024AL06_Francine_Data.xlsx` is the authoritative Francine regression workbook. The text JSON fixture is secondary and should be updated when it disagrees with the issued workbook.

GitHub CI cannot emulate `SpreadsheetApp`, `PropertiesService`, or the live Google Sheet. A release is not considered operationally validated until it is run in the bound PSH testing spreadsheet.

WeatherSTEM v0.10 sustained-wind validation is documented in `docs/WEATHERSTEM_FRANCINE_VALIDATION.md`. The maximum valid direct minute Anemometer method is retained. Tiger Stadium gust remains a known discrepancy; the issued 48 kt target is intentionally kept rather than loosening tolerance.

v0.13 corrects the COOP request path and adds a conservative batch fallback for CoCoRaHS. COOP stations are requested by their template/NWSLI identifier (for example LIX), while response-side aliases still accept forms such as COOPLIX. If the official CoCoRaHS API leaves a station blank, the IEM daily CoCoRaHS mirror may fill it only for exact 12Z-to-12Z whole-day rainfall windows; it never overrides an official CoCoRaHS value. A fresh bound-Sheet Francine run is still required for operational acceptance.

For Francine:

1. Paste `dist/PSH_Automation.gs` into the bound Apps Script project.
2. Reload the spreadsheet.
3. Run **PSH Automation → Load Francine Test Window**.
4. Run **Run EVERYTHING**.
5. Review generated sheets and `_PSH_Log`.
6. Run **Run Francine Regression Check** only after spot-checking the run.

## Design rule

Do not weaken a QC rule merely to improve regression score. If the automation cannot defend the station identity, variable semantics, datum, or historical conversion, keep the field blank and log why.
