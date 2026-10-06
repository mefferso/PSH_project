# Workbook fixtures

The exact XLSX workbooks used by this project are committed in this directory:

- `PSHLIX_testingspreadsheet.xlsx` — the working/test template used while iterating on the automation
- `PSHLIX_2024AL06_Francine_Data.xlsx` — the completed/issued Hurricane Francine PSH

## Authority

`PSHLIX_2024AL06_Francine_Data.xlsx` is the authoritative regression reference for Francine. Regression expectations may be extracted from it, but QC must not be weakened merely to reproduce an issued value that cannot be independently defended.

The testing workbook is useful for template/layout validation and for comparing generated values, but it does not override the issued Francine workbook.

## Integrity

`source_manifest.json` records the expected SHA-256, size, repository path, and role of each source fixture.

Run:

```bash
npm run fixtures:inspect
```

The inspector reads the XLSX ZIP/XML directly with Node built-ins, verifies the committed binaries against the manifest, and prints the WeatherSTEM rows from `Wind and Pressure` for regression review.

`npm test` also runs this integrity check in GitHub Actions.

## Text regression fixtures

`tests/francine_expected.json` contains selected issued-product values used by code/tests. It is derived from the issued Francine workbook and should be kept aligned with that workbook. The XLSX remains the authority when the two disagree.
