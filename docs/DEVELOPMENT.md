# Development workflow

## Source of truth

Edit `src/PSH_Automation.gs`.

Do not hand-edit `dist/PSH_Automation.gs`; generate it with:

```bash
npm run build
```

The build is deliberately trivial today so the deployed artifact stays exactly equivalent to the canonical source.

## What GitHub can test

The CI smoke test:

- parses the Apps Script as JavaScript by copying it to a temporary `.js` file
- verifies `src` and `dist` are byte-identical
- checks required public entry points
- checks critical conservative-QC helpers are still present
- guards against accidentally committing obvious credential strings

These are guardrails, not meteorological validation.

## What still requires the Google Sheet

Anything involving:

- `SpreadsheetApp`
- `PropertiesService`
- bound-sheet menus
- template row/column layout
- formula preservation
- live HTTP behavior through `UrlFetchApp`
- actual API results
- the Francine regression function

must be validated in the PSH testing spreadsheet.

## Recommended change cycle

1. Make the change in `src/PSH_Automation.gs`.
2. Run `npm test`.
3. Run `npm run build`.
4. Review the diff.
5. Download `dist/PSH_Automation.gs`.
6. Paste into the testing spreadsheet's Apps Script project.
7. Run the Francine test window.
8. Inspect `Wind and Pressure`, `Rainfall`, `Water Level`, `Summary`, and `_PSH_Log`.
9. Run the Francine regression after manual spot checks.
10. Commit/merge only the behavior we can defend.

## Reference workbooks

The original uploaded XLSX files are binary fixtures. Their exact fingerprints are recorded in `fixtures/source_manifest.json`.

For code review and future automated regression work, prefer extracting only the station/reference values needed into text/JSON fixtures rather than treating the XLSX binary as executable truth.

## Current known follow-up

v0.10 changed WeatherSTEM to direct station-specific historical retrieval and can now produce an independent Alex Box observation. The current Francine regression code still contains the older deliberate-blank Alex Box assertion. That should be reviewed against the issued Francine PSH before changing the regression expectation.
