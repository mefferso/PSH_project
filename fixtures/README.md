# Workbook fixtures

Two workbook baselines were supplied when this repository was initialized:

- `PSHLIX_testingspreadsheet.xlsx` — the working/test template used while iterating on the automation
- `PSHLIX_2024AL06_Francine_Data.xlsx` — the completed Hurricane Francine PSH reference product

Their SHA-256 fingerprints are in `source_manifest.json`.

The GitHub connector used to initialize this repository can write UTF-8 repository files but does not provide a raw-byte upload path for local XLSX binaries. For that reason, the exact binary workbooks are **not silently reconstructed or converted and committed as if they were originals**.

The regression values that matter to code are kept as text in `tests/francine_expected.json`, and the Apps Script itself contains the operational Francine regression suite.

If the exact XLSX binaries are desired in GitHub too, upload them once through the GitHub web UI into this directory; their hashes can then be checked against `source_manifest.json`.
