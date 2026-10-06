# Development workflow

## Source of truth

Edit `src/PSH_Automation.gs`.

Do not hand-edit `dist/PSH_Automation.gs`; generate it with:

```bash
npm run build
```

The build is deliberately trivial today so the deployed artifact stays exactly equivalent to the canonical source.

## What GitHub can test

The CI test:

- parses the Apps Script as JavaScript by copying it to a temporary `.js` file
- verifies `src` and `dist` are byte-identical
- checks required public entry points
- checks critical conservative-QC helpers are still present
- guards against accidentally committing obvious credential strings
- verifies the committed XLSX fixture hashes/sizes against `fixtures/source_manifest.json`
- reads the XLSX ZIP/XML directly and confirms the Francine/testing workbook structure is readable

These are guardrails, not a replacement for Apps Script/runtime meteorological validation.

## What still requires the Google Sheet

Anything involving:

- `SpreadsheetApp`
- `PropertiesService`
- bound-sheet menus
- formula preservation after writes
- live HTTP behavior through `UrlFetchApp`
- the full operational Francine run and regression function

must still be validated in the bound PSH testing spreadsheet.

## Recommended change cycle

1. Make the change in `src/PSH_Automation.gs`.
2. Run `npm test`.
3. Run `npm run build`.
4. Review the diff.
5. Download `dist/PSH_Automation.gs`.
6. Paste it into the testing spreadsheet's bound Apps Script project.
7. Run the Francine test window.
8. Inspect `Wind and Pressure`, `Rainfall`, `Water Level`, `Summary`, and `_PSH_Log`.
9. Run the Francine regression after manual spot checks.
10. Keep only behavior that can be defended.

## Reference workbooks

The exact binary fixtures are committed under `fixtures/`.

- `PSHLIX_2024AL06_Francine_Data.xlsx` is the authoritative issued-product regression reference.
- `PSHLIX_testingspreadsheet.xlsx` is the working/testing template snapshot.

Run `npm run fixtures:inspect` to verify their fingerprints and print WeatherSTEM rows from the wind/pressure sheet.

For code review, selected reference values also live in `tests/francine_expected.json`; when it conflicts with the issued workbook, update the JSON rather than redefining the workbook.

## Current WeatherSTEM follow-up

v0.10 changed WeatherSTEM to direct station-specific historical retrieval and uses the maximum valid direct minute Anemometer observation for sustained wind. Issue #1 tracks validation of that method against the issued Francine workbook, especially Alex Box and the other WeatherSTEM rows. Regression expectations should change only where the direct station identity and statistic are defensible.
