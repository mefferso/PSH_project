# WeatherSTEM Francine validation

## Scope

This validation compares the issued Hurricane Francine PSH workbook
(`fixtures/PSHLIX_2024AL06_Francine_Data.xlsx`) with the v0.10 direct
WeatherSTEM behavior in the bound PSH testing spreadsheet.

The issued workbook is authoritative. QC is not relaxed to improve the score.

## Validation run

The bound spreadsheet log shows a full Hurricane Francine run beginning
**2026-10-05 19:02:47 local time**. WeatherSTEM rows were populated after the
coordinate-matched Synoptic fallback pass by the direct station-specific
WeatherSTEM historical pass.

v0.10 sustained wind uses the **maximum valid direct minute Anemometer
observation**, converted from mph to kt. Gust remains the native WeatherSTEM
**10 Minute Wind Gust** maximum. Pressure remains from the existing trusted
fallback path; the direct WeatherSTEM pass does not currently replace pressure.

## Priority-station results

| Station | Issued sustained (kt) | v0.10 sustained (kt) | Error (kt) | Issued gust (kt) | v0.10 gust (kt) | Gust error (kt) |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| WSEBRAlexBox | 25 | 25.20 | +0.20 | 30 | 30.41 | +0.41 |
| WSEBRTigerStadium | 44 | 44.32 | +0.32 | 48 | 44.32 | -3.68 |
| WSNOLakefront / Municipal Yacht Harbor | 47 | 46.92 | -0.08 | 54 | 53.88 | -0.12 |
| WSNOMidCIty | 44 | 44.32 | +0.32 | 53 | 53.01 | +0.01 |
| WSNOBayouSauvage | 45 | 45.19 | +0.19 | 59 | 59.09 | +0.09 |
| WSSCEOC | 37 | 37.37 | +0.37 | 46 | 46.06 | +0.06 |
| WSSCLuling | 38 | 38.23 | +0.23 | 46 | 46.06 | +0.06 |

All seven sustained-wind values are within 0.4 kt of the issued PSH values.

## Comparison with the earlier 10-minute rolling mean

The earlier rolling-mean method materially under-represented the issued
Francine sustained winds. Previously documented examples include:

| Station | Issued (kt) | Earlier rolling mean (kt) | Error (kt) | v0.10 max-minute (kt) |
| --- | ---: | ---: | ---: | ---: |
| WSEBRAlexBox | 25 | 18.3 | -6.7 | 25.20 |
| WSEBRTigerStadium | 44 | 27.0 | -17.0 | 44.32 |
| WSNOLakefront / Municipal Yacht Harbor | 47 | 31.4 | -15.6 | 46.92 |
| WSSCEOC | 37 | 22.8 | -14.2 | 37.37 |
| WSSCLuling | 38 | 21.3 | -16.7 | 38.23 |

Across those five directly comparable stations, mean absolute sustained-wind
error improves from about **14.0 kt** with the rolling mean to about **0.24 kt**
with the v0.10 max-minute method.

## Alex Box / Issue #1

The issued workbook contains separate station-specific hyperlinks:

- Alex Box: `https://eastbatonrouge.weatherstem.com/data?refer=/alexbox`
- Tiger Stadium: `https://eastbatonrouge.weatherstem.com/data?refer=/tigerstadium`

The v0.10 direct pass independently produced **25.20 kt sustained / 30.41 kt
gust** at Alex Box, matching the issued **25 / 30 kt** values within rounding.
The old regression assertion that Alex Box must remain blank is therefore no
longer defensible and has been removed.

Alex Box pressure is a separate issue: the issued workbook lists **996.6 mb**,
but the current direct WeatherSTEM pass intentionally does not replace pressure,
and the current run did not produce a trusted fallback pressure there. It
remains blank/manual until a defensible pressure path is validated.

## Remaining WeatherSTEM discrepancy

Tiger Stadium sustained wind validates cleanly, but the current direct native
10-minute gust result is **44.32 kt** versus **48 kt issued**. The regression
keeps **48 kt** as the authoritative target with the normal 1-kt tolerance so
the discrepancy remains visible.

Do not widen the tolerance or substitute the sustained maximum merely to make
that check pass. Investigate the historical gust sensor/semantics separately.

## Decision

Keep v0.10's **maximum valid direct minute Anemometer observation** as the
WeatherSTEM sustained-wind method. It matches the issued Francine PSH
substantially better than the earlier rolling 10-minute mean while preserving
the existing hard plausibility and gust-vs-sustained QC.
