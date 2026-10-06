# PSH Project

GitHub development home for the NWS WFO LIX Post-Tropical Cyclone Report (PSH) Google Apps Script automation.

## Production model

**GitHub is the source of truth. Google Apps Script remains the operational runtime.**

Development, review, history, and lightweight automated checks happen here. The finished production artifact is:

`dist/PSH_Automation.gs`

Download that file and paste it into the Apps Script project bound to the PSH Google Sheet.

No `clasp` deployment is required.

## Current baseline

- Baseline imported from **PSH Automation v0.10**
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
  francine_expected.json   key issued-PSH reference values

fixtures/
  source_manifest.json     fingerprints of the uploaded workbook baselines
  README.md                fixture workflow

scripts/
  build.mjs                creates dist/ from src/

docs/
  DEVELOPMENT.md           development + validation workflow

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

GitHub CI cannot emulate `SpreadsheetApp`, `PropertiesService`, or the live Google Sheet. A release is not considered operationally validated until it is run in the bound PSH testing spreadsheet.

For Francine:

1. Paste `dist/PSH_Automation.gs` into the bound Apps Script project.
2. Reload the spreadsheet.
3. Run **PSH Automation → Load Francine Test Window**.
4. Run **Run EVERYTHING**.
5. Review generated sheets and `_PSH_Log`.
6. Run **Run Francine Regression Check** only after spot-checking the run.

## Design rule

Do not weaken a QC rule merely to improve regression score. If the automation cannot defend the station identity, variable semantics, datum, or historical conversion, keep the field blank and log why.
