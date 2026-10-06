/**
 * PSH Post-Tropical Cyclone Report automation for
 * Copy of PSHLIX_YYYYALXX_StormName_Data
 *
 * v0.11 - Rainfall source/ID hardening + WeatherSTEM v0.10 validation
 *
 * Adds:
 *   - hard meteorological plausibility QC before values can enter Summary
 *   - separate rainfall collection window (critical for PSH storm-total matching)
 *   - CoCoRaHS direct retrieval + ACIS fallback for COOP/HADS daily rain
 *   - MSLP-only pressure selection (never substitutes station pressure for MSLP)
 *   - retry/backoff for transient 429/5xx/HTML gateway failures
 *   - Synoptic nearest-station fallback for stale/local station aliases
 *   - USGS NAVD88 elevation first; controlled 00065+site-datum fallback with QC flag
 *   - WeatherSTEM sustained-wind correction: uses the maximum valid direct minute Anemometer observation
 *   - WeatherFlow intentionally manual (no geographic substitution)
 *   - cross-variable wind QC (gust must be >= sustained)
 *   - validated USGS stage->NAVD88 conversion allowlist
 *   - dynamic rainfall station discovery (CoCoRaHS + Synoptic)
 *   - compact coverage summaries + expanded Francine regression checks
 *
 * What this version automates:
 *   - Storm metadata in Summary
 *   - Wind / gust / pressure from Synoptic Data API
 *   - Rainfall totals from Synoptic + CoCoRaHS + ACIS fallbacks
 *   - NOAA CO-OPS water levels
 *   - USGS water levels / NAVD88 conversion where defensible
 *   - Summary tornado count + update date
 *   - Source audit + run log
 *
 * What remains intentionally manual / Phase 2:
 *   - USACE / LA CPRA: CWMS public API exists, but RiverGages SID-to-CWMS TSID mapping still needs verified implementation
 *   - Tornado narratives / EF ratings
 *   - Inland Flooding narratives
 *   - Impacts / fatalities unless entered during configuration
 *   - Event Summary narrative
 *
 * IMPORTANT:
 *   Synoptic requires a PUBLIC API TOKEN (not the private key).
 *   Store it with PSH Automation -> Set Synoptic Token.
 */

const PSH = Object.freeze({
  SUMMARY: 'Summary',
  WIND: 'Wind and Pressure',
  RAIN: 'Rainfall',
  WATER: 'Water Level',
  TORNADO: 'Tornadoes',
  LOG: '_PSH_Log',

  PROP_START: 'PSH_START_UTC',
  PROP_END: 'PSH_END_UTC',
  PROP_RAIN_START: 'PSH_RAIN_START_UTC',
  PROP_RAIN_END: 'PSH_RAIN_END_UTC',
  PROP_STORM: 'PSH_STORM_NAME',
  PROP_ATCF: 'PSH_ATCF_ID',
  PROP_SYNOPTIC: 'PSH_SYNOPTIC_TOKEN',
  PROP_WEATHERSTEM: 'PSH_WEATHERSTEM_KEY',
  PROP_USGS: 'PSH_USGS_API_KEY',

  SYNOPTIC_TS: 'https://api.synopticdata.com/v2/stations/timeseries',
  SYNOPTIC_PRECIP: 'https://api.synopticdata.com/v2/stations/precip',
  COOPS: 'https://api.tidesandcurrents.noaa.gov/api/prod/datagetter',
  USGS_IV: 'https://waterservices.usgs.gov/nwis/iv/', // legacy fallback; USGS decommissions in early 2027
  USGS_LOCATIONS: 'https://api.waterdata.usgs.gov/ogcapi/v1/collections/monitoring-locations/items',
  USGS_CONTINUOUS: 'https://api.waterdata.usgs.gov/ogcapi/v1/collections/continuous/items',
  ACIS_STNDATA: 'https://data.rcc-acis.org/StnData',
  SYNOPTIC_META: 'https://api.synopticdata.com/v2/stations/metadata',
  COCORAHS_DAILY: 'https://api2.cocorahs.org/api/DailyPrecipObs',
  WEATHERSTEM: 'https://api.weatherstem.com/api',
  WEATHERSTEM_CDN: 'https://cdn.weatherstem.com/dashboard/data/dynamic/model',
  IEM_ASOS: 'https://mesonet.agron.iastate.edu/cgi-bin/request/asos.py',
  IEM_NETWORK_GEOJSON: 'https://mesonet.agron.iastate.edu/geojson/network.php',

  // Stage->NAVD88 conversions are intentionally conservative. These IDs were
  // validated against the completed Hurricane Francine PSH within 0.15 ft.
  // A direct NAVD88/elevation parameter is still preferred and does not need this list.
  USGS_NAVD88_CONVERSION_ALLOWLIST: Object.freeze({
    'TSPL1': true, 'PPAL1': true, 'EPCM6': true, 'OFBM6': true, 'GTEL1': true,
    'DOSL1': true, 'LBLL1': true, 'CPGL1': true, 'LCLL1': true, 'HACL1': true,
    'RFPL1': true, 'PLHL1': true, 'PSIL1': true, 'SIPL1': true, 'CBDL1': true,
    'EWEL1': true, 'WILL1': true, 'GRPL1': true, '301200090072400': true,
    'MSVL1': true, 'DACL1': true, 'CCOL1': true, 'DCLL1': true
  }),

  // Kept deliberately small to avoid monster URLs / responses.
  SYNOPTIC_CHUNK: 25,
  FETCH_PAUSE_MS: 150,

  // Deliberately generous hard limits. Values outside these are not written.
  QC: Object.freeze({
    WIND_MAX_KT: 220,
    GUST_MAX_KT: 250,
    PRESSURE_MIN_MB: 850,
    PRESSURE_MAX_MB: 1100,
    RAIN_MAX_IN: 50,
    WATER_MIN_FT: -20,
    WATER_MAX_FT: 50,
    GEO_MATCH_MILES: 0.08,
    WEATHERSTEM_EXACT_MATCH_MILES: 0.015,
    RAIN_DISCOVERY_PAD_DEG: 0.20,
    WIND_GUST_EPSILON_KT: 0.5
  })
});

// Per-execution caches for CoCoRaHS legacy-ID resolution. These are intentionally
// mutable globals (not persisted properties) so each run starts from fresh catalog data.
const COCORAHS_RESOLVE_CACHE_ = {};
const COCORAHS_IEM_CATALOG_CACHE_ = {};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PSH Automation')
    .addItem('1. Configure Storm', 'pshConfigureStorm')
    .addItem('2. Configure Rainfall Window', 'pshConfigureRainfallWindow')
    .addItem('3. Set Synoptic Token', 'pshSetSynopticToken')
    .addItem('4. Set WeatherSTEM Key (optional)', 'pshSetWeatherStemKey')
    .addItem('5. Set USGS API Key (optional)', 'pshSetUsgsApiKey')
    .addSeparator()
    .addItem('Run EVERYTHING', 'pshRunAll')
    .addItem('Run Wind + Pressure', 'pshRunWindPressure')
    .addItem('Run Rainfall', 'pshRunRainfall')
    .addItem('Run Water Levels', 'pshRunWaterLevels')
    .addItem('Refresh Summary Metadata', 'pshRefreshSummary')
    .addSeparator()
    .addItem('Audit Source Links', 'pshAuditSources')
    .addItem('Show Automation Log', 'pshShowLog')
    .addItem('Clear Automation Log', 'pshClearLog')
    .addSeparator()
    .addItem('Load Francine Test Window', 'pshLoadFrancineTest')
    .addItem('Run Francine Regression Check', 'pshRunFrancineRegression')
    .addItem('Write Coverage Summary to Log', 'pshCoverageSummary')
    .addItem('Debug WeatherSTEM Sensors', 'pshDebugWeatherStem')
    .addToUi();
}

/** ---------------------------- UI / CONFIG ---------------------------- */

function pshConfigureStorm() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  const storm = promptRequired_(ui, 'Storm name',
    'Enter the report storm name exactly as you want it shown.\nExample: Hurricane Francine');
  if (storm === null) return;

  const atcf = promptOptional_(ui, 'ATCF ID',
    'Optional. Example: AL06 or 2024AL06');
  if (atcf === null) return;

  const startText = promptRequired_(ui, 'Start time (UTC)',
    'Enter the start of the collection window in UTC.\nExamples:\n2024-09-10 00:00\n2024-09-10T00:00Z');
  if (startText === null) return;

  const endText = promptRequired_(ui, 'End time (UTC)',
    'Enter the end of the collection window in UTC.\nExamples:\n2024-09-12 23:59\n2024-09-12T23:59Z');
  if (endText === null) return;

  const start = parseUtc_(startText);
  const end = parseUtc_(endText);
  if (!start || !end || end <= start) {
    ui.alert('Bad time window', 'I could not parse those UTC times, or the end is not after the start.', ui.ButtonSet.OK);
    return;
  }

  const direct = promptOptional_(ui, 'Direct fatalities', 'Enter a number, or leave blank if not known yet.');
  if (direct === null) return;
  const indirect = promptOptional_(ui, 'Indirect fatalities', 'Enter a number, or leave blank if not known yet.');
  if (indirect === null) return;

  props.setProperty(PSH.PROP_STORM, storm);
  props.setProperty(PSH.PROP_ATCF, atcf || '');
  props.setProperty(PSH.PROP_START, start.toISOString());
  props.setProperty(PSH.PROP_END, end.toISOString());
  // New storm = reset rainfall window to the storm window. Override with menu item #2 when needed.
  props.setProperty(PSH.PROP_RAIN_START, start.toISOString());
  props.setProperty(PSH.PROP_RAIN_END, end.toISOString());

  writeSummaryConfig_(storm, start, end, direct, indirect);
  writeRainWindow_(start, end);
  log_('INFO', 'CONFIG', '', `Configured ${storm}; ${start.toISOString()} through ${end.toISOString()}`);
  ui.alert('Configured', 'Storm window saved. Next, set the Synoptic public token if you have not already.', ui.ButtonSet.OK);
}

function pshConfigureRainfallWindow() {
  const ui = SpreadsheetApp.getUi();
  const cfg = getConfig_();
  if (!cfg) return;
  const props = PropertiesService.getDocumentProperties();

  const startText = promptRequired_(ui, 'Rainfall start time (UTC)',
    'Enter the rainfall accumulation start time in UTC.\nExample: 2024-09-10 12:00');
  if (startText === null) return;
  const endText = promptRequired_(ui, 'Rainfall end time (UTC)',
    'Enter the rainfall accumulation end time in UTC.\nExample: 2024-09-12 12:00');
  if (endText === null) return;

  const start = parseUtc_(startText);
  const end = parseUtc_(endText);
  if (!start || !end || end <= start) {
    ui.alert('Bad rainfall window', 'I could not parse those UTC times, or the end is not after the start.', ui.ButtonSet.OK);
    return;
  }
  props.setProperty(PSH.PROP_RAIN_START, start.toISOString());
  props.setProperty(PSH.PROP_RAIN_END, end.toISOString());
  writeRainWindow_(start, end);
  log_('INFO','CONFIG','',`Rainfall window set to ${start.toISOString()} through ${end.toISOString()}`);
  ui.alert('Rainfall window saved', `${start.toISOString()} through ${end.toISOString()}`, ui.ButtonSet.OK);
}

function pshSetSynopticToken() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(
    'Synoptic public API token',
    'Paste your PUBLIC TOKEN here. Do NOT paste the Synoptic private API key.\n\nThe token is stored in this spreadsheet\'s Document Properties, not in a cell.',
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const token = r.getResponseText().trim();
  if (!token) {
    ui.alert('No token entered.');
    return;
  }
  PropertiesService.getDocumentProperties().setProperty(PSH.PROP_SYNOPTIC, token);
  log_('INFO', 'CONFIG', '', 'Synoptic public API token saved.');
  ui.alert('Saved', 'Synoptic public token saved.', ui.ButtonSet.OK);
}


function pshSetWeatherStemKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(
    'WeatherSTEM API key (optional)',
    'WeatherSTEM historical data requires its own API key. Paste it here if you have one. This is separate from Synoptic.\n\nIf you leave WeatherSTEM unconfigured, the script will still use Synoptic and all other sources.',
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const key = r.getResponseText().trim();
  if (!key) {
    PropertiesService.getDocumentProperties().deleteProperty(PSH.PROP_WEATHERSTEM);
    ui.alert('WeatherSTEM key cleared.');
    return;
  }
  PropertiesService.getDocumentProperties().setProperty(PSH.PROP_WEATHERSTEM, key);
  log_('INFO', 'CONFIG', '', 'WeatherSTEM API key saved.');
  ui.alert('Saved', 'WeatherSTEM API key saved.', ui.ButtonSet.OK);
}

function pshSetUsgsApiKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(
    'USGS Water Data API key (optional)',
    'Paste a free USGS Water Data API key here to raise the modern API rate limit. The key is stored in this spreadsheet\'s Document Properties, not in a cell. Leave blank to clear it.',
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const key = r.getResponseText().trim();
  const props = PropertiesService.getDocumentProperties();
  if (!key) {
    props.deleteProperty(PSH.PROP_USGS);
    log_('INFO','CONFIG','','USGS API key cleared.');
    ui.alert('USGS key cleared.');
    return;
  }
  props.setProperty(PSH.PROP_USGS, key);
  log_('INFO','CONFIG','','USGS Water Data API key saved.');
  ui.alert('Saved', 'USGS Water Data API key saved.', ui.ButtonSet.OK);
}

function pshLoadFrancineTest() {
  const props = PropertiesService.getDocumentProperties();
  const start = new Date('2024-09-10T00:00:00Z');
  const end = new Date('2024-09-12T23:59:00Z');
  props.setProperty(PSH.PROP_STORM, 'Hurricane Francine');
  props.setProperty(PSH.PROP_ATCF, '2024AL06');
  const rainStart = new Date('2024-09-10T12:00:00Z');
  const rainEnd = new Date('2024-09-12T12:00:00Z');
  props.setProperty(PSH.PROP_START, start.toISOString());
  props.setProperty(PSH.PROP_END, end.toISOString());
  props.setProperty(PSH.PROP_RAIN_START, rainStart.toISOString());
  props.setProperty(PSH.PROP_RAIN_END, rainEnd.toISOString());
  writeSummaryConfig_('Hurricane Francine', start, end, '0', '0');
  writeRainWindow_(rainStart, rainEnd);
  log_('INFO', 'CONFIG', '', 'Loaded Hurricane Francine storm + rainfall windows.');
  SpreadsheetApp.getUi().alert('Francine test loaded',
    'Storm: 2024-09-10 00:00Z through 2024-09-12 23:59Z.\nRainfall: 2024-09-10 12:00Z through 2024-09-12 12:00Z.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/** ---------------------------- MAIN RUNNERS ---------------------------- */

function pshRunAll() {
  const ui = SpreadsheetApp.getUi();
  const cfg = getConfig_();
  if (!cfg) return;

  const answer = ui.alert(
    'Run PSH automation?',
    `Storm: ${cfg.storm}\nStorm UTC: ${cfg.start.toISOString()} through ${cfg.end.toISOString()}\nRain UTC: ${cfg.rainStart.toISOString()} through ${cfg.rainEnd.toISOString()}\n\nThis will overwrite generated observation fields in Wind/Pressure, Rainfall, and Water Level. Station metadata and source links are preserved.`,
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  pshClearLog();
  log_('INFO', 'RUN', '', `Starting full run for ${cfg.storm}`);

  try { pshRunWindPressure_(cfg); } catch (e) { log_('ERROR', 'WIND', '', e.stack || e.message); }
  try { pshRunRainfall_(cfg); } catch (e) { log_('ERROR', 'RAIN', '', e.stack || e.message); }
  try { pshRunWaterLevels_(cfg); } catch (e) { log_('ERROR', 'WATER', '', e.stack || e.message); }
  try { pshRefreshSummary_(cfg); } catch (e) { log_('ERROR', 'SUMMARY', '', e.stack || e.message); }
  try { pshCoverageSummary(); } catch (e) { log_('ERROR', 'COVERAGE', '', e.stack || e.message); }

  SpreadsheetApp.flush();
  log_('INFO', 'RUN', '', 'Full run finished. Review _PSH_Log and spot-check QC before issuance.');
  ui.alert('Run finished', 'Automation finished. Review the generated data AND the _PSH_Log before using the report.', ui.ButtonSet.OK);
}

function pshRunWindPressure() {
  const cfg = getConfig_();
  if (cfg) pshRunWindPressure_(cfg);
}

function pshRunRainfall() {
  const cfg = getConfig_();
  if (cfg) pshRunRainfall_(cfg);
}

function pshRunWaterLevels() {
  const cfg = getConfig_();
  if (cfg) pshRunWaterLevels_(cfg);
}

function pshRefreshSummary() {
  const cfg = getConfig_();
  if (cfg) pshRefreshSummary_(cfg);
}

/** ---------------------------- WIND / PRESSURE ---------------------------- */

function pshRunWindPressure_(cfg) {
  const token = getSynopticToken_();
  if (!token) throw new Error('Synoptic public API token is not set. Use PSH Automation -> Set Synoptic Token.');

  const sh = mustSheet_(PSH.WIND);
  const lastRow = findLastStationRow_(sh, 1);
  if (lastRow < 2) throw new Error('No station rows found on Wind and Pressure.');

  sh.getRange(2, 11, lastRow - 1, 17).clearContent(); // K:AA

  // Remove only prior automation QC notes; preserve human QC.
  clearAutoQc_(sh, 2, lastRow, 28, 30);

  const rows = sh.getRange(2, 1, lastRow - 1, 10).getValues();
  const requested = [];
  const byApiId = {};
  const metaByRow = {};

  rows.forEach((r, i) => {
    const sheetRow = i + 2;
    const rawId = String(r[0] || '').trim();
    if (!rawId) return;
    const network = String(r[7] || '').trim().toUpperCase();
    const apiId = synopticId_(rawId, network);
    // WeatherFlow/DataScope and WeatherSTEM sheet aliases are generally not valid
    // Synoptic STIDs. Sending them directly causes whole-batch failures and lots of
    // recursive retries. Route those networks straight to geographic/source fallbacks.
    if (!/WEATHERFLOW|WEATHERSTEM/.test(network)) {
      requested.push(apiId);
      if (!byApiId[apiId.toUpperCase()]) byApiId[apiId.toUpperCase()] = [];
      byApiId[apiId.toUpperCase()].push(sheetRow);
    }
    metaByRow[sheetRow] = {
      id: rawId, name: String(r[1] || ''), lat: numeric_(r[2]), lon: numeric_(r[3]),
      network: network, apiId: apiId
    };
  });

  const start = synopticTime_(cfg.start);
  const end = synopticTime_(cfg.end);
  const outByRow = {};

  chunks_(unique_(requested), PSH.SYNOPTIC_CHUNK).forEach(chunk => {
    const params = {
      token,
      stid: chunk.join(','),
      start,
      end,
      vars: 'wind_speed,wind_gust,wind_direction,sea_level_pressure',
      units: 'english,pres|mb,speed|kts',
      obtimezone: 'UTC',
      hfmetars: '1',
      qc: 'on'
    };
    const stations = fetchSynopticStationsResilient_(
      PSH.SYNOPTIC_TS, params, chunk, 'WIND', `Synoptic wind ${chunk[0]}...`
    );

    stations.forEach(st => {
      const stid = String(st.STID || '').toUpperCase();
      const sheetRows = byApiId[stid] || [];
      if (!sheetRows.length) return;
      const parsed = parseSynopticWindStation_(st);
      sheetRows.forEach(rowNum => outByRow[rowNum] = parsed);
    });
    Utilities.sleep(PSH.FETCH_PAUSE_MS);
  });


  // ASOS/AWOS correctness pass using the IEM METAR archive. IEM exposes routine,
  // special, and MADIS HFMETAR/5-minute reports plus PK WND remarks. For airport
  // networks this is a better fit for PSH max sustained/gust than relying on a
  // single Synoptic sensor set. When IEM returns usable values, they override the
  // corresponding Synoptic fields; missing IEM fields fall back to Synoptic.
  for (let rowNum = 2; rowNum <= lastRow; rowNum++) {
    const m = metaByRow[rowNum];
    if (!m || !/^(ASOS|AWOS)$/.test(m.network)) continue;
    try {
      const iem = fetchIemAirportWind_(m.id, cfg.start, cfg.end);
      if (iem) {
        outByRow[rowNum] = mergeParsedWind_(outByRow[rowNum], iem);
        log_('INFO','WIND',m.id,'Airport wind/pressure reconciled with IEM METAR/HFMETAR archive.');
      }
    } catch (e) {
      log_('INFO','WIND',m.id,`IEM airport reconciliation unavailable; kept Synoptic values: ${e.message || e}`);
    }
    Utilities.sleep(PSH.FETCH_PAUSE_MS);
  }

  // Try to rescue stale/local aliases only by a very tight coordinate match.
  // This avoids silently substituting a nearby but different sensor.
  for (let rowNum = 2; rowNum <= lastRow; rowNum++) {
    if (outByRow[rowNum]) continue;
    const m = metaByRow[rowNum];
    if (!m || m.lat === null || m.lon === null) continue;
    if (!/WLON|WEATHERSTEM/.test(m.network)) continue;
    try {
      const rescued = fetchSynopticByLocation_(token, m, cfg.start, cfg.end);
      if (rescued) {
        // For WeatherSTEM, Synoptic is acceptable as a fallback for gust/pressure
        // but not for sustained wind: Francine validation shows generic Synoptic
        // wind_speed and the earlier rolling 10-minute mean systematically
        // under-represent the issued PSH sustained-wind statistic. Keep wind blank
        // unless the direct WeatherSTEM minute Anemometer pass succeeds.
        if (m.network === 'WEATHERSTEM') {
          rescued.parsed.wind = {value:null,time:null,direction:null,index:-1,sensorKey:'WeatherSTEM-manual'};
          rescued.parsed.weatherStemSustainedNeedsDirect = true;
        }
        outByRow[rowNum] = rescued.parsed;
        log_('INFO', 'WIND', m.id, `Recovered via coordinate-matched Synoptic station ${rescued.stid} (${round_(rescued.distance, 2)} mi).`);
      }
    } catch (e) {
      log_('WARN', 'WIND', m.id, `Alias recovery failed: ${e.message || e}`);
    }
  }

  // Direct WeatherSTEM pass. The public station metadata + /data endpoint exposes
  // minute Anemometer observations and the native 10 Minute Wind Gust sensor.
  // Use the maximum valid minute Anemometer observation as PSH sustained wind.
  // The native 10 Minute Wind Gust logic remains separate and unchanged.
  // This runs even without a WeatherSTEM API key.
  for (let rowNum = 2; rowNum <= lastRow; rowNum++) {
    const m = metaByRow[rowNum];
    if (!m || m.network !== 'WEATHERSTEM') continue;
    const link = cellLink_(sh.getRange(rowNum, 1));
    try {
      const direct = fetchWeatherStemWind_(link, cfg.start, cfg.end);
      if (direct) outByRow[rowNum] = mergeParsedWind_(outByRow[rowNum], direct);
    } catch (e) {
      log_('WARN', 'WIND', m.id, `WeatherSTEM direct: ${e.message || e}`);
    }
  }

  // WeatherFlow is intentionally manual. Do not substitute a nearby station.
  for (let rowNum = 2; rowNum <= lastRow; rowNum++) {
    const m = metaByRow[rowNum];
    if (m && m.network === 'WEATHERFLOW') {
      outByRow[rowNum] = null;
      log_('INFO','WIND',m.id,'MANUAL REQUIRED — WeatherFlow historical data is intentionally not auto-retrieved.');
    }
  }

  const values = [];
  const qcWrites = [];
  for (let row = 2; row <= lastRow; row++) {
    const m = metaByRow[row] || {id: sh.getRange(row, 1).getDisplayValue(), network: ''};
    let p = outByRow[row];
    if (!p) {
      values.push(new Array(17).fill(''));
      qcWrites.push(null);
      if (m.network !== 'WEATHERFLOW') {
        log_('WARN', 'WIND', m.id, 'No usable wind/pressure result returned for the requested period.');
      }
      continue;
    }

    const qc = qcWindPressure_(p, m);
    p = qc.parsed;
    values.push([
      valueOrBlank_(p.wind.value),
      valueOrBlank_(p.wind.direction),
      hhmm_(p.wind.time),
      day_(p.wind.time), month_(p.wind.time), year_(p.wind.time),
      valueOrBlank_(p.gust.value),
      valueOrBlank_(p.gust.direction),
      hhmm_(p.gust.time),
      day_(p.gust.time), month_(p.gust.time), year_(p.gust.time),
      valueOrBlank_(p.pressure.value),
      hhmm_(p.pressure.time),
      day_(p.pressure.time), month_(p.pressure.time), year_(p.pressure.time),
    ]);
    qcWrites.push(qc.messages.length ? qc : null);
  }

  sh.getRange(2, 11, values.length, 17).setValues(values);
  applyAutoQc_(sh, qcWrites, 2);
  log_('INFO', 'WIND', '', `Wind/pressure complete: ${values.length} station rows processed.`);
}

function fetchIemAirportWind_(rawId, startDate, endDate) {
  let station = String(rawId || '').trim().toUpperCase();
  // IEM's US airport archive normally uses the 3-character FAA identifier.
  if (/^K[A-Z0-9]{3}$/.test(station)) station = station.substring(1);
  if (!station) return null;

  const params = [
    ['station', station],
    ['sts', startDate.toISOString()],
    ['ets', endDate.toISOString()],
    ['tz', 'Etc/UTC'],
    ['format', 'onlycomma'],
    ['missing', 'M'],
    ['trace', 'T'],
    ['direct', 'no'],
    ['report_type', '1'], // MADIS HFMETAR / 5-minute ASOS
    ['report_type', '3'], // routine
    ['report_type', '4'], // specials
    ['data', 'sknt'],
    ['data', 'drct'],
    ['data', 'gust'],
    ['data', 'peak_wind_gust'],
    ['data', 'peak_wind_drct'],
    ['data', 'peak_wind_time'],
    ['data', 'mslp']
  ];
  const qs = params.map(x => encodeURIComponent(x[0]) + '=' + encodeURIComponent(x[1])).join('&');
  const resp = fetchTextResilient_(PSH.IEM_ASOS + '?' + qs, {}, `IEM airport ${station}`);
  const rows = Utilities.parseCsv(resp);
  if (!rows || rows.length < 2) return null;
  const hdr = rows[0].map(x => String(x || '').trim().toLowerCase());
  const idx = {};
  hdr.forEach((h,i) => idx[h]=i);
  const val = (r,k) => idx[k] === undefined ? '' : r[idx[k]];

  let wind={value:null,time:null,direction:null,index:-1,sensorKey:'IEM-sknt'};
  let gust={value:null,time:null,direction:null,index:-1,sensorKey:'IEM-gust'};
  let pressure={value:null,time:null,index:-1,sensorKey:'IEM-mslp'};

  for (let i=1;i<rows.length;i++) {
    const r=rows[i];
    const valid=parseApiTime_(val(r,'valid'));
    if (!valid || valid < startDate || valid > endDate) continue;

    const spd=numeric_(val(r,'sknt'));
    if (spd !== null && spd >= 0 && spd <= PSH.QC.WIND_MAX_KT && (wind.value===null || spd>wind.value)) {
      const dir=numeric_(val(r,'drct'));
      wind={value:spd,time:valid,direction:(dir!==null&&dir>=0&&dir<=360?Math.round(dir):null),index:i,sensorKey:'IEM-sknt'};
    }

    const g=numeric_(val(r,'gust'));
    if (g !== null && g >= 0 && g <= PSH.QC.GUST_MAX_KT && (gust.value===null || g>gust.value)) {
      const dir=numeric_(val(r,'drct'));
      gust={value:g,time:valid,direction:(dir!==null&&dir>=0&&dir<=360?Math.round(dir):null),index:i,sensorKey:'IEM-gust'};
    }

    const pg=numeric_(val(r,'peak_wind_gust'));
    if (pg !== null && pg >= 0 && pg <= PSH.QC.GUST_MAX_KT && (gust.value===null || pg>gust.value)) {
      const pdir=numeric_(val(r,'peak_wind_drct'));
      const pt=parseIemPeakTime_(val(r,'peak_wind_time'), valid);
      gust={value:pg,time:pt || valid,direction:(pdir!==null&&pdir>=0&&pdir<=360?Math.round(pdir):null),index:i,sensorKey:'IEM-peak_wind_gust'};
    }

    const mslp=numeric_(val(r,'mslp'));
    if (mslp !== null && mslp >= PSH.QC.PRESSURE_MIN_MB && mslp <= PSH.QC.PRESSURE_MAX_MB &&
        (pressure.value===null || mslp<pressure.value)) {
      pressure={value:mslp,time:valid,index:i,sensorKey:'IEM-mslp'};
    }
  }

  if (wind.value===null && gust.value===null && pressure.value===null) return null;
  return {wind,gust,pressure};
}

function parseIemPeakTime_(raw, valid) {
  const s=String(raw || '').trim();
  if (!s || s==='M' || !valid) return null;
  const d=parseApiTime_(s);
  if (d) return d;
  const m=s.match(/^(\d{1,2}):?(\d{2})$/);
  if (!m) return null;
  const out=new Date(valid.getTime());
  out.setUTCHours(parseInt(m[1],10),parseInt(m[2],10),0,0);
  // PK WND time can belong to the previous UTC day relative to the METAR.
  if (out-valid > 6*3600*1000) out.setUTCDate(out.getUTCDate()-1);
  return out;
}

function fetchTextResilient_(url, options, label) {
  let last='';
  for (let attempt=0;attempt<4;attempt++) {
    const resp=UrlFetchApp.fetch(url, Object.assign({muteHttpExceptions:true}, options || {}));
    const code=resp.getResponseCode();
    const txt=resp.getContentText();
    if (code>=200 && code<300) return txt;
    last=`HTTP ${code}; ${String(txt).slice(0,300)}`;
    if (!(code===429 || code>=500)) break;
    Utilities.sleep(500*Math.pow(2,attempt));
  }
  throw new Error(`${label || 'request'} failed: ${last}`);
}

function fetchSynopticByLocation_(token, meta, startDate, endDate) {
  const radius = `${meta.lat},${meta.lon},${PSH.QC.GEO_MATCH_MILES}`;
  const json = fetchJson_(PSH.SYNOPTIC_META, {
    token,
    radius,
    limit: 8,
    complete: 1,
    sensorvars: 1,
    obrange: `${synopticTime_(startDate)},${synopticTime_(endDate)}`
  }, `Synoptic metadata near ${meta.id}`);

  const maxMiles = meta.network === 'WEATHERSTEM' ? PSH.QC.WEATHERSTEM_EXACT_MATCH_MILES : PSH.QC.GEO_MATCH_MILES;
  const candidates = (json.STATION || []).filter(st => {
    const d = numeric_(st.DISTANCE);
    return d !== null && d <= maxMiles;
  });
  if (!candidates.length) return null;

  candidates.sort((a,b) => {
    const an = String(a.NAME || '').toLowerCase(), bn = String(b.NAME || '').toLowerCase();
    const target = String(meta.name || '').toLowerCase();
    const as = nameSimilarity_(target, an), bs = nameSimilarity_(target, bn);
    if (bs !== as) return bs - as;
    return (numeric_(a.DISTANCE) || 999) - (numeric_(b.DISTANCE) || 999);
  });

  for (const st of candidates.slice(0,3)) {
    const stid = String(st.STID || '').trim();
    if (!stid) continue;
    try {
      const data = fetchJson_(PSH.SYNOPTIC_TS, {
        token, stid,
        start: synopticTime_(startDate), end: synopticTime_(endDate),
        vars: 'wind_speed,wind_gust,wind_direction,sea_level_pressure',
        units: 'english,pres|mb,speed|kts', obtimezone: 'UTC', hfmetars: '1', qc: 'on'
      }, `Synoptic alias ${stid}`);
      const station = (data.STATION || [])[0];
      if (!station) continue;
      const parsed = parseSynopticWindStation_(station);
      if (parsed.wind.value !== null || parsed.gust.value !== null || parsed.pressure.value !== null) {
        return {stid, distance: numeric_(st.DISTANCE) || 0, parsed};
      }
    } catch (e) {}
  }
  return null;
}

function parseSynopticWindStation_(st) {
  const obs = st.OBSERVATIONS || {};
  const times = obs.date_time || [];

  const wind = extremumWithTime_(obs, times, ['wind_speed'], 'max', v => v >= 0 && v <= PSH.QC.WIND_MAX_KT);
  const gust = extremumWithTime_(obs, times, ['wind_gust'], 'max', v => v >= 0 && v <= PSH.QC.GUST_MAX_KT);
  // PSH column W is MSLP. Raw station pressure can be tens of mb lower at elevated
  // sites and must NEVER be mixed into the MSLP minimum search.
  const pressure = extremumWithTime_(obs, times, ['sea_level_pressure'], 'min',
    v => v >= PSH.QC.PRESSURE_MIN_MB && v <= PSH.QC.PRESSURE_MAX_MB);

  if (wind.index >= 0) wind.direction = directionAt_(obs, wind.sensorKey, wind.index);
  if (gust.index >= 0) gust.direction = directionAt_(obs, gust.sensorKey, gust.index);

  return { wind, gust, pressure };
}

function extremumWithTime_(obs, times, prefixes, mode, validator) {
  let best = { value: null, time: null, index: -1, sensorKey: '' };
  const keys = Object.keys(obs).filter(k => prefixes.some(p => k === p || k.indexOf(p + '_') === 0));

  keys.forEach(key => {
    const arr = obs[key];
    if (!Array.isArray(arr)) return;
    arr.forEach((raw, i) => {
      const v = numeric_(raw);
      if (v === null) return;
      if (validator && !validator(v)) return;
      if (best.value === null || (mode === 'max' ? v > best.value : v < best.value)) {
        best = { value: v, time: parseApiTime_(times[i]), index: i, sensorKey: key };
      }
    });
  });
  return best;
}

function directionAt_(obs, speedKey, index) {
  if (index < 0) return null;
  const suffix = speedKey.replace(/^wind_(speed|gust)/, '');
  const candidates = [
    'wind_direction' + suffix,
    ...Object.keys(obs).filter(k => k.indexOf('wind_direction') === 0)
  ];
  for (const key of unique_(candidates)) {
    const arr = obs[key];
    if (!Array.isArray(arr) || index >= arr.length) continue;
    const v = numeric_(arr[index]);
    if (v !== null && v >= 0 && v <= 360) return Math.round(v);
  }
  return null;
}

function qcWindPressure_(parsed, meta) {
  const p = JSON.parse(JSON.stringify(parsed));
  // Restore Date objects lost by JSON copy.
  ['wind','gust','pressure'].forEach(k => {
    if (parsed[k] && parsed[k].time) p[k].time = new Date(parsed[k].time);
  });
  const messages = [];
  const vars = [];

  const checks = [
    ['wind', 0, PSH.QC.WIND_MAX_KT, 'S', 'sustained wind'],
    ['gust', 0, PSH.QC.GUST_MAX_KT, 'G', 'gust'],
    ['pressure', PSH.QC.PRESSURE_MIN_MB, PSH.QC.PRESSURE_MAX_MB, 'P', 'pressure']
  ];
  checks.forEach(c => {
    const obj = p[c[0]];
    if (!obj || obj.value === null) return;
    if (obj.value < c[1] || obj.value > c[2]) {
      messages.push(`${c[4]} ${obj.value} failed hard plausibility QC`);
      vars.push(c[3]);
      obj.value = null; obj.time = null; obj.direction = null;
    }
  });

  // Cross-variable consistency: a reported gust lower than sustained wind is
  // physically/semantically inconsistent for the PSH fields and usually means
  // the wrong sensor set was selected. Blank both rather than guessing.
  if (p.wind && p.gust && p.wind.value !== null && p.gust.value !== null &&
      p.gust.value + PSH.QC.WIND_GUST_EPSILON_KT < p.wind.value) {
    messages.push(`gust ${round_(p.gust.value,2)} kt < sustained ${round_(p.wind.value,2)} kt; both blanked`);
    vars.push('S'); vars.push('G');
    p.wind.value=null; p.wind.time=null; p.wind.direction=null;
    p.gust.value=null; p.gust.time=null; p.gust.direction=null;
  }

  if (parsed.weatherStemSustainedNeedsDirect && (!p.wind || p.wind.value === null)) {
    messages.push('WeatherSTEM sustained wind left blank: no valid direct minute Anemometer observation was available');
    vars.push('S');
  }

  // Also catch values that were excluded while parsing when raw series had outliers.
  if (parsed.pressure.value === null && meta && /CWOP|PWS/.test(meta.network)) {
    // Not an error by itself; consumer networks often lack pressure.
  }
  return { parsed: p, messages, vars: unique_(vars) };
}

function nameSimilarity_(a, b) {
  const aw = new Set(String(a || '').replace(/[^a-z0-9 ]/g,' ').split(/\s+/).filter(x => x.length > 2));
  const bw = new Set(String(b || '').replace(/[^a-z0-9 ]/g,' ').split(/\s+/).filter(x => x.length > 2));
  if (!aw.size || !bw.size) return 0;
  let hit = 0; aw.forEach(x => { if (bw.has(x)) hit++; });
  return hit / Math.max(aw.size, bw.size);
}

function weatherStemLinkParts_(link) {
  const m = String(link || '').match(/^https?:\/\/([^.]+)\.weatherstem\.com\/data\?refer=\/([^&#]+)/i);
  if (!m) throw new Error('Could not derive WeatherSTEM network/slug from link.');
  return {network: String(m[1]).toLowerCase(), slug: decodeURIComponent(m[2])};
}

function fetchWeatherStemStationMeta_(network, slug) {
  const url = `${PSH.WEATHERSTEM_CDN}/${encodeURIComponent(network)}/${encodeURIComponent(slug)}/station.json`;
  const resp = UrlFetchApp.fetch(url, {
    method:'get', muteHttpExceptions:true,
    headers:{'User-Agent':'Mozilla/5.0','Accept':'application/json'}
  });
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
    throw new Error(`station.json HTTP ${resp.getResponseCode()}`);
  }
  const json = JSON.parse(resp.getContentText());
  if (!json || !json.id) throw new Error('WeatherSTEM station metadata did not include station id.');
  return json;
}

function weatherStemSensors_(stationMeta) {
  const out = {anemometer:null, gust10:null, vane:null};
  const tx = Array.isArray(stationMeta && stationMeta.transmitters) ? stationMeta.transmitters : [];
  tx.forEach(t => {
    const sensors = Array.isArray(t && t.sensors) ? t.sensors : [];
    sensors.forEach(sensor => {
      const name = String(sensor && (sensor.name || sensor.type) || '').toLowerCase();
      const typeId = numeric_(sensor && sensor.type_id);
      const id = numeric_(sensor && sensor.id);
      if (id === null) return;
      if (!out.gust10 && (typeId === 47 || /10\s*minute\s*wind\s*gust/.test(name))) out.gust10 = id;
      else if (!out.anemometer && (typeId === 15 || name === 'anemometer')) out.anemometer = id;
      else if (!out.vane && (typeId === 17 || name === 'wind vane')) out.vane = id;
    });
  });
  return out;
}

function fetchWeatherStemMinuteData_(network, slug, stationId, sensorIds, start, end) {
  const ids = sensorIds.filter(x => x !== null && x !== undefined).map(x => String(x));
  if (!ids.length) return null;
  const url = `https://${network}.weatherstem.com/data`;
  const payload = {
    timezone_offset: 0,
    id: String(stationId),
    start_date: Utilities.formatDate(start, 'UTC', 'yyyy-MM-dd HH:mm'),
    end_date: Utilities.formatDate(end, 'UTC', 'yyyy-MM-dd HH:mm'),
    operation: 'datapoint',
    interval: 'minute',
    sensors: ids,
    format: 'json',
    timestamp_format: 'standard',
    record_id: '1',
    mysql_mode: false,
    query: ''
  };
  const resp = UrlFetchApp.fetch(url, {
    method:'post',
    payload:JSON.stringify(payload),
    contentType:'application/x-www-form-urlencoded; charset=UTF-8',
    muteHttpExceptions:true,
    headers:{
      'Accept':'application/json, text/javascript, */*; q=0.01',
      'Origin':`https://${network}.weatherstem.com`,
      'Referer':`https://${network}.weatherstem.com/data?refer=/${slug}`,
      'X-Requested-With':'XMLHttpRequest',
      'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/147 Safari/537.36'
    }
  });
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
    throw new Error(`/data HTTP ${resp.getResponseCode()}`);
  }
  const json = JSON.parse(resp.getContentText());
  if (!Array.isArray(json) || !json.length || !Array.isArray(json[0])) {
    throw new Error('WeatherSTEM /data returned unexpected JSON layout.');
  }
  return json;
}

function weatherStemTable_(raw) {
  if (!Array.isArray(raw) || raw.length < 2 || !Array.isArray(raw[0])) return [];
  const hdr = raw[0].map(x => String(x || '').trim());
  return raw.slice(1).map(row => {
    const o = {};
    hdr.forEach((h,i) => { o[h] = row[i]; });
    return o;
  });
}

function weatherStemDirectionNear_(rows, target, directionHeader) {
  if (!target || !directionHeader) return null;
  let best = null, bestDt = Infinity;
  rows.forEach(r => {
    const t = parseApiTime_(r.Timestamp || r.timestamp || r.Time || r.time);
    const d = numeric_(r[directionHeader]);
    if (!t || d === null) return;
    const dt = Math.abs(t.getTime() - target.getTime());
    if (dt < bestDt && dt <= 2 * 60 * 1000) { bestDt = dt; best = d; }
  });
  return best;
}

function weatherStemMaxAnemometer_(rows, windHeader) {
  let best = null;
  rows.forEach(r => {
    const time = parseApiTime_(r.Timestamp || r.timestamp || r.Time || r.time);
    let v = numeric_(r[windHeader]);
    if (!time || v === null) return;
    // WeatherSTEM station metadata for these Davis stations reports mph.
    v *= 0.868976242;
    if (v < 0 || v > PSH.QC.WIND_MAX_KT) return;
    if (!best || v > best.value) best = {value:v, time};
  });
  return best;
}

function fetchWeatherStemWind_(link, start, end) {
  const parts = weatherStemLinkParts_(link);
  const meta = fetchWeatherStemStationMeta_(parts.network, parts.slug);
  const sensors = weatherStemSensors_(meta);
  if (!sensors.anemometer) throw new Error('No Anemometer sensor found in station metadata.');

  const ids = [sensors.anemometer, sensors.gust10, sensors.vane].filter(Boolean);
  const raw = fetchWeatherStemMinuteData_(parts.network, parts.slug, meta.id, ids, start, end);
  const rows = weatherStemTable_(raw);
  const headers = raw[0].map(x => String(x || '').trim());
  const windHeader = headers.find(h => /^anemometer$/i.test(h)) || headers.find(h => /wind speed/i.test(h));
  const gustHeader = headers.find(h => /10\s*minute\s*wind\s*gust/i.test(h)) || headers.find(h => /wind gust/i.test(h));
  const dirHeader = headers.find(h => /^wind vane$/i.test(h)) || headers.find(h => /wind direction/i.test(h));
  if (!windHeader) throw new Error(`WeatherSTEM /data did not return Anemometer header. Headers: ${headers.join(', ')}`);

  const maxWind = weatherStemMaxAnemometer_(rows, windHeader);
  let wind = {value:null,time:null,direction:null,index:-1,sensorKey:'WeatherSTEM'};
  let gust = {value:null,time:null,direction:null,index:-1,sensorKey:'WeatherSTEM'};
  const pressure = {value:null,time:null,index:-1,sensorKey:'WeatherSTEM'}; // preserve trusted Synoptic pressure

  if (maxWind) {
    wind = {
      value:maxWind.value,
      time:maxWind.time,
      direction:weatherStemDirectionNear_(rows, maxWind.time, dirHeader),
      index:-1,
      sensorKey:'WeatherSTEM-anemometer-max'
    };
  }

  if (gustHeader) {
    rows.forEach(r => {
      let v = numeric_(r[gustHeader]);
      const time = parseApiTime_(r.Timestamp || r.timestamp || r.Time || r.time);
      if (v === null || !time) return;
      v *= 0.868976242;
      if (v < 0 || v > PSH.QC.GUST_MAX_KT) return;
      if (gust.value === null || v > gust.value) {
        gust = {
          value:v, time,
          direction:weatherStemDirectionNear_(rows, time, dirHeader),
          index:-1, sensorKey:'WeatherSTEM-10min-gust'
        };
      }
    });
  }
  return {wind,gust,pressure};
}

function mergeParsedWind_(base, direct) {
  const empty = () => ({value:null,time:null,direction:null,index:-1,sensorKey:''});
  const out = base || {wind:empty(),gust:empty(),pressure:empty()};
  ['wind','gust','pressure'].forEach(k => {
    if (direct && direct[k] && direct[k].value !== null && direct[k].value !== undefined) out[k] = direct[k];
    else if (!out[k]) out[k] = empty();
  });
  if (direct && direct.wind && direct.wind.value !== null) out.weatherStemSustainedNeedsDirect = false;
  return out;
}

function flattenWeatherStem_(x, out) {
  if (!x) return;
  if (Array.isArray(x)) { x.forEach(v => flattenWeatherStem_(v,out)); return; }
  if (typeof x !== 'object') return;
  const keys = Object.keys(x);
  if (keys.some(k => /sensor|name|type/i.test(k)) && keys.some(k => /value|reading/i.test(k)) && keys.some(k => /time|date/i.test(k))) {
    out.push(x);
  }
  keys.forEach(k => flattenWeatherStem_(x[k], out));
}

function isWeatherStemSustainedSensor_(sensorText) {
  const s = String(sensorText || '').toLowerCase().replace(/[_]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!s || s.indexOf('gust') >= 0 || s.indexOf('anemometer') >= 0 || s.indexOf('instant') >= 0) return false;
  // Accept only names that explicitly communicate averaging/sustained semantics.
  // Generic "wind speed" remains excluded because it may be instantaneous.
  const hasWind = s.indexOf('wind') >= 0;
  const hasSustainedSemantics = /(sustained|10\s*[- ]?(minute|min)|ten\s*[- ]?minute|avg|average|averaged|mean)/.test(s);
  return hasWind && hasSustainedSemantics;
}

function pshDebugWeatherStem(link) {
  const ui = SpreadsheetApp.getUi();
  if (!link) {
    const r = ui.prompt('WeatherSTEM station link', 'Paste one WeatherSTEM station data link to inspect station/sensor metadata.', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    link = r.getResponseText().trim();
  }
  if (!link) return;
  const parts = weatherStemLinkParts_(link);
  const meta = fetchWeatherStemStationMeta_(parts.network, parts.slug);
  const sensors = weatherStemSensors_(meta);
  log_('INFO','WEATHERSTEM-DEBUG',`${parts.slug}@${parts.network}`,`station_id=${meta.id}; Anemometer=${sensors.anemometer || 'none'}; 10 Minute Wind Gust=${sensors.gust10 || 'none'}; Wind Vane=${sensors.vane || 'none'}`);
  const tx = Array.isArray(meta.transmitters) ? meta.transmitters : [];
  tx.forEach(t => (Array.isArray(t.sensors) ? t.sensors : []).forEach(sensor => {
    const ds = sensor && sensor.data_source ? sensor.data_source : {};
    const unit = sensor && sensor.units ? sensor.units : {};
    log_('INFO','WEATHERSTEM-DEBUG',`${parts.slug}@${parts.network}`,
      `sensor id=${sensor.id}; type_id=${sensor.type_id}; name=${sensor.name || sensor.type || ''}; units=${unit.symbol || unit.name || ''}; source=${ds.metadata || ''}`);
  }));
  ui.alert('WeatherSTEM debug complete', 'Station and sensor metadata written to _PSH_Log.', ui.ButtonSet.OK);
}

/** ---------------------------- RAINFALL ---------------------------- */


function discoverRainStations_(sh, cfg, token) {
  /*
   * PSH guidance says there must be NO blank rows within the data listing.
   * v0.5 inserted discoveries immediately before "Latest Update:", which could
   * leave the template's gray/blank separator above the new rows. First clean
   * any prior auto-discovery residue that sits below the first blank separator,
   * then insert trusted discoveries immediately after the contiguous station list.
   */
  let footer = sh.createTextFinder('Latest Update:').matchCase(false).useRegularExpression(false).findNext();
  const footerRow0 = footer ? footer.getRow() : (sh.getLastRow() + 1);
  const ids0 = sh.getRange(2,1,Math.max(1,footerRow0-2),1).getDisplayValues().flat();
  let firstBlank = 0;
  for (let i=0;i<ids0.length;i++) {
    if (!String(ids0[i] || '').trim()) { firstBlank = i + 2; break; }
  }
  if (!firstBlank) firstBlank = footerRow0;

  // If v0.5 left data below the first separator blank, remove only that invalid
  // trailing block. A correctly formatted PSH sheet must not contain data there.
  let firstTrailingData = 0, lastTrailingData = 0;
  for (let row=firstBlank+1; row<footerRow0; row++) {
    const v = String(sh.getRange(row,1).getDisplayValue() || '').trim();
    if (v) {
      if (!firstTrailingData) firstTrailingData = row;
      lastTrailingData = row;
    }
  }
  if (firstTrailingData) {
    const n = lastTrailingData - firstTrailingData + 1;
    sh.deleteRows(firstTrailingData, n);
    log_('INFO','RAIN','DISCOVERY',`Removed ${n} prior auto-discovery rows that were below the PSH blank separator.`);
    footer = sh.createTextFinder('Latest Update:').matchCase(false).useRegularExpression(false).findNext();
  }

  const footerRow = footer ? footer.getRow() : (sh.getLastRow() + 1);
  const idsNow = sh.getRange(2,1,Math.max(1,footerRow-2),1).getDisplayValues().flat();
  let insertAt = footerRow;
  for (let i=0;i<idsNow.length;i++) {
    if (!String(idsNow[i] || '').trim()) { insertAt = i + 2; break; }
  }

  const existingLast = Math.max(2, insertAt - 1);
  const existing = sh.getRange(2,1,Math.max(1,existingLast-1),7).getValues();
  const ids = new Set();
  const coords = [];
  existing.forEach(r => {
    const id = String(r[0] || '').trim();
    const network = String(r[6] || '').trim().toUpperCase();
    if (id) rainIdAliasKeys_(id, network).forEach(k => ids.add(k));
    const lat = numeric_(r[2]), lon = numeric_(r[3]);
    if (lat !== null && lon !== null) coords.push([lat,lon]);
  });
  if (!coords.length) return;

  let minLat=Math.min.apply(null,coords.map(x=>x[0])), maxLat=Math.max.apply(null,coords.map(x=>x[0]));
  let minLon=Math.min.apply(null,coords.map(x=>x[1])), maxLon=Math.max.apply(null,coords.map(x=>x[1]));
  const pad=PSH.QC.RAIN_DISCOVERY_PAD_DEG;
  minLat-=pad; maxLat+=pad; minLon-=pad; maxLon+=pad;

  const discovered = [];

  // Auto-add only CoCoRaHS discoveries with PSH-significant storm totals (>=3").
  // These have stable station IDs and are retrieved from the official CoCoRaHS API.
  // Lower totals remain available through the pre-populated/template stations.
  ['LA','MS'].forEach(state => {
    try {
      const reps = fetchCocorahsReportsForState_(state, cfg.rainStart, cfg.rainEnd);
      const byStation = {};
      reps.forEach(r => {
        const id = String(r.stationNumber || r.StationNumber || r.station_number || '').trim();
        const lat = numeric_(r.latitude !== undefined ? r.latitude : r.Latitude);
        const lon = numeric_(r.longitude !== undefined ? r.longitude : r.Longitude);
        const t = parseApiTime_(r.obsDateTime || r.ObsDateTime);
        if (!id || lat===null || lon===null || !t) return;
        if (lat<minLat || lat>maxLat || lon<minLon || lon>maxLon) return;
        if (t<=cfg.rainStart || t>cfg.rainEnd) return;
        if (!byStation[id]) byStation[id]=[];
        byStation[id].push(r);
      });
      Object.keys(byStation).forEach(id => {
        if (rainIdAliasKeys_(id, 'COCORAHS').some(k=>ids.has(k))) return;
        const rs=byStation[id];
        let total=0, count=0;
        rs.forEach(r => {
          let v = numeric_(r.gaugeCatch);
          if (v === null) v = numeric_(r.precip);
          if (v === null && (r.gaugeCatchIsTrace || r.precipIsTrace)) v = 0;
          if (v === null || v < 0 || v > 30) return;
          total += v; count++;
        });
        if (!count || total < 3 || total > PSH.QC.RAIN_MAX_IN) return;
        const r=rs[0];
        const lat=numeric_(r.latitude !== undefined ? r.latitude : r.Latitude);
        const lon=numeric_(r.longitude !== undefined ? r.longitude : r.Longitude);
        discovered.push([
          id,
          String(r.stationName || r.StationName || r.station_name || id),
          lat, lon,
          String(r.countyName || r.CountyName || r.county || ''),
          state,
          'CoCoRaHS'
        ]);
        rainIdAliasKeys_(id, 'COCORAHS').forEach(k=>ids.add(k));
      });
    } catch(e) {
      log_('INFO','RAIN','DISCOVERY',`CoCoRaHS ${state} discovery unavailable: ${e.message || e}`);
    }
  });


  // Historical CoCoRaHS fallback: the bulk CoCoRaHS endpoint can miss stations
  // that were valid for an older event but are no longer part of the current
  // reporting roster. IEM retains historical CoCoRaHS network metadata. Use that
  // catalog only to enumerate candidate station IDs, then retrieve the actual
  // rainfall from the official CoCoRaHS API for the configured 12Z-to-12Z window.
  ['LA','MS'].forEach(state => {
    try {
      const features=fetchIemNetworkStations_(`${state}_COCORAHS`);
      // Rank only in-domain, non-template catalog stations before applying the
      // runtime cap. Prefer stations whose IEM archive interval overlaps this
      // storm, especially stations still active. The old first-150 scan could
      // exhaust itself on unrelated/closed stations before reaching valid LIX
      // observers such as Jefferson or St. Charles Parish.
      const candidates=features
        .map(f => {
          const id=String(f.id || (f.properties && (f.properties.sid || f.properties.station)) || '').trim();
          const coords=(f.geometry && f.geometry.coordinates) || [];
          const lon=numeric_(coords[0]), lat=numeric_(coords[1]);
          if (!id || lat===null || lon===null) return null;
          if (lat<minLat || lat>maxLat || lon<minLon || lon>maxLon) return null;
          if (rainIdAliasKeys_(id, 'COCORAHS').some(k=>ids.has(k))) return null;
          const p=f.properties || {};
          const begin=parseApiTime_(p.archive_begin || p.archiveBegin || p.begints || p.start || '');
          const finish=parseApiTime_(p.archive_end || p.archiveEnd || p.endts || p.end || '');
          const overlaps=!(begin && begin>cfg.rainEnd) && !(finish && finish<cfg.rainStart);
          const priority=overlaps ? (finish ? 1 : 0) : 2;
          return {f,id,lat,lon,priority};
        })
        .filter(Boolean)
        .sort((a,b) => a.priority-b.priority || a.id.localeCompare(b.id));

      let tried=0, added=0;
      const maxHistoricalChecks=200;
      for (const candidate of candidates) {
        if (tried >= maxHistoricalChecks) {
          log_('INFO','RAIN','DISCOVERY',`Historical CoCoRaHS fallback capped after ${maxHistoricalChecks} event-prioritized ${state} candidates (${candidates.length} available).`);
          break;
        }
        const f=candidate.f;
        const id=candidate.id;
        const lon=candidate.lon, lat=candidate.lat;
        tried++;
        try {
          const summary=fetchCocorahsStationSummary_(id,cfg.rainStart,cfg.rainEnd);
          if (!summary || summary.total===null || summary.total<3 || summary.total>PSH.QC.RAIN_MAX_IN) continue;
          const props=f.properties || {};
          const county=String(summary.county || props.county || props.county_name || props.countyName || '').trim();
          discovered.push([
            id,
            summary.name || String(props.sname || props.name || id),
            summary.lat!==null ? summary.lat : lat,
            summary.lon!==null ? summary.lon : lon,
            county,
            state,
            'CoCoRaHS'
          ]);
          rainIdAliasKeys_(id, 'COCORAHS').forEach(k=>ids.add(k));
          added++;
        } catch(e) {}
        if (tried % 15===0) Utilities.sleep(80);
      }
      if (added) log_('INFO','RAIN','DISCOVERY',`Historical CoCoRaHS catalog fallback added ${added} ${state} stations with >=3\" event totals.`);
    } catch(e) {
      log_('INFO','RAIN','DISCOVERY',`IEM historical CoCoRaHS catalog unavailable for ${state}: ${e.message || e}`);
    }
  });

  /*
   * Non-template Synoptic rainfall stations are NOT auto-inserted in v0.6.
   * v0.5 proved that broad discovery can pull technically valid but operationally
   * unvetted private/mesonet totals into the Top-10 list. Per PSH guidance, automated
   * output still requires QC for completeness/accuracy. We therefore log high-total
   * candidates for human review without letting them contaminate the issued table.
   */
  try {
    const json = fetchJson_(PSH.SYNOPTIC_PRECIP, {
      token,
      bbox: `${minLon},${minLat},${maxLon},${maxLat}`,
      start: synopticTime_(cfg.rainStart),
      end: synopticTime_(cfg.rainEnd),
      pmode: 'totals',
      units: 'english,precip|in',
      obtimezone: 'UTC',
      all_reports: '0',
      complete: '1'
    }, 'Synoptic rainfall candidate discovery');
    const candidates = [];
    (json.STATION || []).forEach(st => {
      const id=String(st.STID || '').trim();
      const net=normalizeRainNetwork_(st.MNET_SHORTNAME || st.SOURCE || st.MNET_ID || 'Synoptic');
      if (!id || rainIdAliasKeys_(id, net).some(k=>ids.has(k))) return;
      const total=bestPrecipTotal_(st.OBSERVATIONS || {});
      if (total===null || total < 3 || total > PSH.QC.RAIN_MAX_IN) return;
      if (/WEATHERFLOW|PWS|PERSONAL/.test(net)) return;
      candidates.push({id,name:String(st.NAME || id),total,net});
    });
    candidates.sort((a,b)=>b.total-a.total);
    candidates.slice(0,20).forEach(c =>
      log_('INFO','RAIN-CANDIDATE',c.id,`${c.name}: ${round_(c.total,2)} in (${c.net}). Review manually before adding to PSH.`)
    );
    if (candidates.length) {
      log_('INFO','RAIN','DISCOVERY',`${candidates.length} non-template Synoptic rainfall stations >=3" were held for manual QC rather than auto-inserted.`);
    }
  } catch(e) {
    log_('INFO','RAIN','DISCOVERY',`Synoptic candidate discovery unavailable: ${e.message || e}`);
  }

  if (!discovered.length) {
    log_('INFO','RAIN','DISCOVERY','No new trusted CoCoRaHS stations >=3" discovered.');
    return;
  }

  // Insert BEFORE the template separator blank/gray row. This preserves the PSH
  // requirement that no blank rows occur inside the data listing.
  sh.insertRowsBefore(insertAt, discovered.length);
  if (insertAt > 2) {
    sh.getRange(insertAt-1,1,1,9).copyFormatToRange(sh,1,9,insertAt,insertAt+discovered.length-1);
  }
  sh.getRange(insertAt,1,discovered.length,7).setValues(discovered);
  log_('INFO','RAIN','DISCOVERY',`Added ${discovered.length} trusted CoCoRaHS stations >=3" directly above the PSH separator row.`);
}

function fetchIemNetworkStations_(network) {
  const json=fetchJson_(PSH.IEM_NETWORK_GEOJSON,{network},`IEM network ${network}`);
  return (json && json.features) || [];
}

function getIemCocorahsCatalog_(state) {
  const st = String(state || '').trim().toUpperCase();
  if (!st) return [];
  if (!COCORAHS_IEM_CATALOG_CACHE_[st]) {
    COCORAHS_IEM_CATALOG_CACHE_[st] = fetchIemNetworkStations_(`${st}_COCORAHS`);
  }
  return COCORAHS_IEM_CATALOG_CACHE_[st];
}

function resolveCocorahsId_(rawId, lat, lon, name) {
  const raw = String(rawId || '').trim().toUpperCase();
  if (!raw) return raw;
  if (Object.prototype.hasOwnProperty.call(COCORAHS_RESOLVE_CACHE_, raw)) return COCORAHS_RESOLVE_CACHE_[raw];

  const stateMatch = raw.match(/^([A-Z]{2})-/);
  const state = stateMatch ? stateMatch[1] : '';
  if (!state) { COCORAHS_RESOLVE_CACHE_[raw] = raw; return raw; }

  let features = [];
  try { features = getIemCocorahsCatalog_(state); }
  catch (e) {
    log_('WARN','RAIN',raw,`CoCoRaHS ID resolver could not load IEM ${state}_COCORAHS catalog: ${e.message || e}`);
    COCORAHS_RESOLVE_CACHE_[raw] = raw;
    return raw;
  }

  // If the current/historical IEM catalog already knows this exact ID, keep it.
  const exact = features.find(f => String(f.id || (f.properties && (f.properties.sid || f.properties.station)))
    .trim().toUpperCase() === raw);
  if (exact) { COCORAHS_RESOLVE_CACHE_[raw] = raw; return raw; }

  const xlat = numeric_(lat), xlon = numeric_(lon);
  let best = null;
  if (xlat !== null && xlon !== null) {
    features.forEach(f => {
      const coords = (f.geometry && f.geometry.coordinates) || [];
      const flon = numeric_(coords[0]), flat = numeric_(coords[1]);
      if (flat === null || flon === null) return;
      const dlat = flat - xlat, dlon = flon - xlon;
      // Do not relabel a nearby observer as the same CoCoRaHS station. Both
      // coordinates and name must strongly agree before an ID substitution is allowed.
      if (Math.abs(dlat) > 0.005 || Math.abs(dlon) > 0.005) return;
      const dist2 = dlat*dlat + dlon*dlon;
      const props = f.properties || {};
      const candidateName = String(props.sname || props.name || props.station_name || '');
      const sim = nameSimilarity_(String(name || '').toLowerCase(), candidateName.toLowerCase());
      if (sim < 0.75) return;
      if (!best || dist2 < best.dist2 || (Math.abs(dist2-best.dist2) < 1e-10 && sim > best.sim)) {
        best = {f, dist2, sim};
      }
    });
  }

  // If coordinates did not resolve, fall back to a strong station-name match.
  if (!best && name) {
    features.forEach(f => {
      const props = f.properties || {};
      const candidateName = String(props.sname || props.name || props.station_name || '');
      const sim = nameSimilarity_(String(name || '').toLowerCase(), candidateName.toLowerCase());
      if (sim < 0.90) return;
      if (!best || sim > best.sim) best = {f, dist2:null, sim};
    });
  }

  const resolved = best ? String(best.f.id || (best.f.properties && (best.f.properties.sid || best.f.properties.station)) || '').trim().toUpperCase() : '';
  if (resolved && resolved !== raw) {
    COCORAHS_RESOLVE_CACHE_[raw] = resolved;
    log_('WARN','RAIN',raw,`CoCoRaHS legacy/current ID substitution: ${raw} -> ${resolved}.`);
    return resolved;
  }

  // Both old and current CoCoRaHS station numbers can share the same XX-YY-NN
  // shape, so unresolved non-catalog IDs are treated conservatively as legacy.
  COCORAHS_RESOLVE_CACHE_[raw] = null;
  log_('WARN','RAIN',raw,'CoCoRaHS station ID could not be resolved through the IEM historical catalog; row left blank.');
  return null;
}

function fetchCocorahsStationSummary_(stationNumber,start,end) {
  const qStart=new Date(start.getTime()-12*3600*1000);
  const qEnd=new Date(end.getTime()+12*3600*1000);
  const state=String(stationNumber).split('-')[0];
  const json=fetchJson_(PSH.COCORAHS_DAILY,{
    offset:0,limit:100,
    startDate:Utilities.formatDate(qStart,'UTC','yyyy-MM-dd'),
    endDate:Utilities.formatDate(qEnd,'UTC','yyyy-MM-dd'),
    sortField:'ObsDateTime',sortDir:'asc',country:'USA',subdiv1:state,
    stationField:'StationNumber',stationFieldValue:stationNumber,units:'english'
  },`CoCoRaHS historical ${stationNumber}`);
  const reps=(json && json.results) || [];
  let total=0,count=0,first=null;
  reps.forEach(r=>{
    const t=parseApiTime_(r.obsDateTime || r.ObsDateTime);
    if (!t || t<=start || t>end) return;
    let v=numeric_(r.gaugeCatch !== undefined ? r.gaugeCatch : r.GaugeCatch);
    if (v===null) v=numeric_(r.precip !== undefined ? r.precip : r.Precip);
    if (v===null && (r.gaugeCatchIsTrace || r.GaugeCatchIsTrace || r.precipIsTrace || r.PrecipIsTrace)) v=0;
    if (v===null || v<0 || v>30) return;
    if (!first) first=r;
    total+=v; count++;
  });
  if (!count) return null;
  first=first || reps[0] || {};
  return {
    total,
    name:String(first.stationName || first.StationName || stationNumber),
    lat:numeric_(first.latitude !== undefined ? first.latitude : first.Latitude),
    lon:numeric_(first.longitude !== undefined ? first.longitude : first.Longitude),
    county:String(first.countyName || first.CountyName || first.county || '')
  };
}

function fetchCocorahsReportsForState_(state, start, end) {
  const qStart=new Date(start.getTime()-12*3600*1000);
  const qEnd=new Date(end.getTime()+12*3600*1000);
  const out=[];
  let offset=0;
  const requestedLimit=250;
  // The service may silently cap the requested page size. Advance by the number
  // actually returned (not the requested limit) and honor metadata.totalCount;
  // otherwise historical stations can be skipped between pages.
  for (let page=0; page<100 && offset<20000; page++) {
    const json=fetchJson_(PSH.COCORAHS_DAILY,{
      offset, limit:requestedLimit,
      startDate: Utilities.formatDate(qStart,'UTC','yyyy-MM-dd'),
      endDate: Utilities.formatDate(qEnd,'UTC','yyyy-MM-dd'),
      sortField:'ObsDateTime', sortDir:'asc',
      country:'USA', subdiv1:state, units:'english'
    },`CoCoRaHS discovery ${state}`);
    const reps=(json && json.results) || [];
    if (!reps.length) break;
    out.push.apply(out,reps);
    offset += reps.length;
    const rs=json && json.metadata && json.metadata.resultset;
    const total=rs ? numeric_(rs.totalCount) : null;
    if (total !== null && offset >= total) break;
    Utilities.sleep(60);
  }
  return out;
}

function canonicalCocorahsId_(id) {
  const u=String(id || '').trim().toUpperCase();
  const m=u.match(/^([A-Z]{2})-([A-Z]{2})-(\d+)$/);
  if (!m) return u;
  return `${m[1]}-${m[2]}-${Number(m[3])}`;
}

function rainIdAliasKeys_(id, network) {
  const u=String(id || '').trim().toUpperCase();
  const net=String(network || '').trim().toUpperCase();
  const keys=[u];

  // Airport rainfall IDs may differ only by the ICAO K prefix.
  if (/^K[A-Z0-9]{3}$/.test(u)) keys.push(u.substring(1));
  else if (/^[A-Z0-9]{3}$/.test(u)) keys.push('K'+u);

  // Synoptic COOP STIDs commonly carry a COOP prefix while the PSH template uses
  // the NWSLI alone (LIX vs COOPLIX). Treat those as the same station only when
  // the network identifies the row as COOP, or when the candidate already has
  // the explicit COOP prefix.
  if (/^COOP[A-Z0-9]+$/.test(u)) keys.push(u.substring(4));
  if (net === 'COOP' && u && !/^COOP/.test(u)) keys.push('COOP'+u);

  // CoCoRaHS station numbers are sometimes exported with a zero-padded sequence
  // (LA-SC-06) while the live service/IEM catalog uses LA-SC-6. Canonicalize only
  // the numeric station suffix; parish/state components must still match exactly.
  if (net === 'COCORAHS' || /^[A-Z]{2}-[A-Z]{2}-\d+$/.test(u)) {
    keys.push(canonicalCocorahsId_(u));
  }
  return unique_(keys);
}

function normalizeRainNetwork_(x) {
  const s=String(x || '').trim().toUpperCase();
  if (/ASOS/.test(s)) return 'ASOS';
  if (/AWOS/.test(s)) return 'AWOS';
  if (/COCORAHS/.test(s)) return 'CoCoRaHS';
  if (/HADS/.test(s)) return 'HADS';
  if (/COOP/.test(s)) return 'COOP';
  if (/CWOP/.test(s)) return 'CWOP';
  if (/RAWS/.test(s)) return 'RAWS';
  if (/MESONET/.test(s)) return 'Mesonet';
  return s || 'Synoptic';
}

function pshRunRainfall_(cfg) {
  const token = getSynopticToken_();
  if (!token) throw new Error('Synoptic public API token is not set. Use PSH Automation -> Set Synoptic Token.');

  const sh = mustSheet_(PSH.RAIN);

  // Discover reporting stations for the event before reading the fixed template.
  // This prevents a stale station roster from hiding the actual highest totals.
  try { discoverRainStations_(sh, cfg, token); }
  catch (e) { log_('WARN','RAIN','DISCOVERY',`Dynamic station discovery failed; continuing with template rows: ${e.message || e}`); }

  const lastRow = findLastStationRow_(sh, 1);
  if (lastRow < 2) throw new Error('No station rows found on Rainfall.');

  sh.getRange(2, 8, lastRow - 1, 1).clearContent(); // H generated total only
  clearAutoQc_(sh, 2, lastRow, 9, 9);             // I: preserve human I/E flags
  writeRainWindow_(cfg.rainStart, cfg.rainEnd);

  const rows = sh.getRange(2, 1, lastRow - 1, 7).getValues();
  const byApiId = {};
  const requested = [];
  const totalsByRow = {};
  const sourceByRow = {};

  rows.forEach((r, i) => {
    const sheetRow = i + 2;
    const rawId = String(r[0] || '').trim();
    if (!rawId) return;
    const network = String(r[6] || '').trim().toUpperCase();
    if (network === 'COCORAHS') return; // official CoCoRaHS API path below
    const apiId = synopticId_(rawId, network);
    requested.push(apiId);
    if (!byApiId[apiId.toUpperCase()]) byApiId[apiId.toUpperCase()] = [];
    byApiId[apiId.toUpperCase()].push(sheetRow);
  });

  const start = synopticTime_(cfg.rainStart);
  const end = synopticTime_(cfg.rainEnd);

  chunks_(unique_(requested), PSH.SYNOPTIC_CHUNK).forEach(chunk => {
    const params = {
      token,
      stid: chunk.join(','),
      start,
      end,
      pmode: 'totals',
      search: 'nearest',
      window: 60,
      units: 'english,precip|in',
      obtimezone: 'UTC',
      all_reports: '0'
    };
    const stations = fetchSynopticStationsResilient_(
      PSH.SYNOPTIC_PRECIP, params, chunk, 'RAIN', `Synoptic precip ${chunk[0]}...`
    );

    stations.forEach(st => {
      const stid = String(st.STID || '').toUpperCase();
      const targetRows = byApiId[stid] || [];
      const total = bestPrecipTotal_(st.OBSERVATIONS || {});
      if (total === null) return;
      targetRows.forEach(rowNum => { totalsByRow[rowNum] = total; sourceByRow[rowNum] = 'Synoptic'; });
    });
    Utilities.sleep(PSH.FETCH_PAUSE_MS);
  });

  // Official CoCoRaHS daily reports. Unlike v0.3, use the configured rainfall
  // window itself -- not a +/-24 h overlap window that could accidentally count
  // post-event rainfall from the next morning's report.
  for (let i = 0; i < rows.length; i++) {
    const rowNum = i + 2;
    const rawId = String(rows[i][0] || '').trim();
    const network = String(rows[i][6] || '').trim().toUpperCase();
    if (!rawId || network !== 'COCORAHS') continue;
    try {
      const resolvedId = resolveCocorahsId_(rawId, rows[i][2], rows[i][3], rows[i][1]);
      if (!resolvedId) continue;
      const total = fetchCocorahsTotal_(resolvedId, cfg.rainStart, cfg.rainEnd);
      if (total !== null) { totalsByRow[rowNum] = total; sourceByRow[rowNum] = 'CoCoRaHS'; }
    } catch (e) {
      log_('WARN', 'RAIN', rawId, `CoCoRaHS direct: ${e.message || e}`);
    }
    if (i % 15 === 0) Utilities.sleep(75);
  }

  // Synoptic precipitation totals are preferred for COOP/HADS when available
  // because they are computed over the exact configured PSH rainfall window.
  // ACIS is daily/calendar-binned and can straddle a 12Z-to-12Z storm window
  // (Francine LIX was 4.33 in from ACIS vs the issued/window-matched 7.93 in).
  // Keep ACIS as a conservative fallback only when the window-matched Synoptic
  // total is unavailable.
  for (let i = 0; i < rows.length; i++) {
    const rowNum = i + 2;
    const rawId = String(rows[i][0] || '').trim();
    const network = String(rows[i][6] || '').trim().toUpperCase();
    if (!rawId || !/^(COOP|HADS)$/.test(network)) continue;
    const synTotal = totalsByRow[rowNum];
    if (synTotal !== undefined && synTotal !== null) continue;
    try {
      const acis = fetchAcisDailyPrecip_(rawId, network, cfg.rainStart, cfg.rainEnd);
      if (acis !== null) {
        totalsByRow[rowNum] = acis;
        sourceByRow[rowNum] = 'ACIS';
        log_('INFO','RAIN',rawId,`Synoptic window total unavailable; used ACIS daily fallback (${round_(acis,2)} in).`);
      }
    } catch (e) {
      // Missing ACIS station/data is common and not a hard run error.
    }
    if (i % 25 === 0) Utilities.sleep(50);
  }

  const output = [];
  const missingByNetwork = {};
  const countsBySource = {};
  for (let row = 2; row <= lastRow; row++) {
    let total = totalsByRow[row];
    const idx = row - 2;
    const id = String(rows[idx][0] || '').trim();
    const network = String(rows[idx][6] || 'UNKNOWN').trim().toUpperCase() || 'UNKNOWN';
    if (total !== undefined && total !== null && (total < 0 || total > PSH.QC.RAIN_MAX_IN)) {
      log_('WARN', 'RAIN', id, `[AUTO-QC] Rain total ${total} in failed plausibility QC; blanked.`);
      total = null;
    }

    if (total === undefined || total === null) {
      output.push(['', '']);
      if (!missingByNetwork[network]) missingByNetwork[network] = [];
      missingByNetwork[network].push(id);
    } else {
      output.push([round_(total, 2), '']);
      const src = sourceByRow[row] || 'unknown';
      countsBySource[src] = (countsBySource[src] || 0) + 1;
    }
  }

  sh.getRange(2, 8, output.length, 1).setValues(output.map(r => [r[0]]));
  mergeQcColumn_(sh, 2, 9, output.map(r => r[1]));
  Object.keys(missingByNetwork).sort().forEach(net => {
    const ids = missingByNetwork[net];
    log_('INFO','RAIN',net,`${ids.length} rows had no usable total in the configured rainfall window. Examples: ${ids.slice(0,12).join(', ')}${ids.length>12?' ...':''}`);
  });
  log_('INFO', 'RAIN', '', `Rainfall complete: ${output.length} rows; populated by source: ${JSON.stringify(countsBySource)}.`);
}

function fetchCocorahsTotal_(stationNumber, start, end) {
  const qStart = new Date(start.getTime() - 12*3600*1000);
  const qEnd = new Date(end.getTime() + 12*3600*1000);
  const state = String(stationNumber).split('-')[0];

  const json = fetchJson_(PSH.COCORAHS_DAILY, {
    offset: 0, limit: 100,
    startDate: Utilities.formatDate(qStart, 'UTC', 'yyyy-MM-dd'),
    endDate: Utilities.formatDate(qEnd, 'UTC', 'yyyy-MM-dd'),
    sortField: 'ObsDateTime', sortDir: 'asc',
    country: 'USA', subdiv1: state,
    stationField: 'StationNumber',
    stationFieldValue: stationNumber,
    units: 'english'
  }, `CoCoRaHS ${stationNumber}`);

  const reports = (json && json.results) || [];
  if (!reports.length) return null;

  let sum = 0, count = 0;
  reports.forEach(r => {
    const responseStation = String(r.stationNumber || r.StationNumber || r.station_number || '').trim();
    if (!responseStation || responseStation.toUpperCase() !== String(stationNumber).trim().toUpperCase()) return;
    const t = parseApiTime_(r.obsDateTime || r.ObsDateTime);
    if (!t || t <= start || t > end) return;
    let v = numeric_(r.gaugeCatch);
    if (v === null) v = numeric_(r.precip);
    if (v === null) {
      if (r.gaugeCatchIsTrace || r.precipIsTrace) v = 0;
      else return;
    }
    if (v < 0 || v > 30) return;
    sum += v; count++;
  });
  return count ? sum : null;
}

function fetchAcisDailyPrecip_(rawId, network, start, end) {
  let sid = '';
  const net = String(network || '').toUpperCase();
  // NWSLI/SHEF type 7 is the best ACIS mapping for COOP/HADS rows in this template.
  sid = `${String(rawId).toUpperCase()} 7`;

  const json = fetchJson_(PSH.ACIS_STNDATA, {
    sid,
    sdate: Utilities.formatDate(start, 'UTC', 'yyyy-MM-dd'),
    edate: Utilities.formatDate(end, 'UTC', 'yyyy-MM-dd'),
    elems: 'pcpn',
    output: 'json'
  }, `ACIS ${rawId}`);
  if (json && json.error) return null;
  const data = (json && json.data) || [];
  let sum = 0, count = 0;
  data.forEach(r => {
    const raw = Array.isArray(r) ? r[1] : null;
    if (raw === null || raw === undefined || raw === '' || raw === 'M') return;
    if (String(raw).toUpperCase() === 'T') { count++; return; }
    const v = parseFloat(String(raw).replace(/[^0-9.+-]/g,''));
    if (!Number.isFinite(v) || v < 0 || v > 30) return;
    sum += v; count++;
  });
  return count ? sum : null;
}

function bestPrecipTotal_(obs) {
  const list = obs.precipitation;
  if (Array.isArray(list) && list.length) {
    // Prefer the report type with the largest raw observation count instead of simply
    // choosing the largest total (which can bias high when multiple sensors exist).
    const valid = list
      .map(x => ({ total: numeric_(x.total), count: numeric_(x.count) || 0 }))
      .filter(x => x.total !== null);
    if (!valid.length) return null;
    valid.sort((a, b) => b.count - a.count);
    return valid[0].total;
  }

  // Legacy/default response fallback.
  const keys = Object.keys(obs).filter(k => k.indexOf('total_precip_value_') === 0);
  for (const k of keys) {
    const v = numeric_(obs[k]);
    if (v !== null) return v;
  }
  return null;
}

/** ---------------------------- WATER LEVEL ---------------------------- */

function pshRunWaterLevels_(cfg) {
  const sh = mustSheet_(PSH.WATER);
  const lastRow = findLastStationRow_(sh, 1);
  if (lastRow < 2) throw new Error('No station rows found on Water Level.');

  // G and I:L plus N are generated. IMPORTANT: preserve H (Datum).
  // v0.1 accidentally cleared G:L, which wiped the datum column before requests were made.
  sh.getRange(2, 7, lastRow - 1, 1).clearContent();      // G: max level
  sh.getRange(2, 9, lastRow - 1, 4).clearContent();      // I:L: time/date
  clearAutoQc_(sh, 2, lastRow, 14, 14);                 // N: preserve human I/E flags

  repairWaterDatums_(sh, lastRow);
  const rows = sh.getRange(2, 1, lastRow - 1, 14).getValues();
  const results = new Array(lastRow - 1).fill(null);

  for (let i = 0; i < rows.length; i++) {
    const rowNum = i + 2;
    const r = rows[i];
    const stationId = String(r[0] || '').trim();
    const source = String(r[12] || '').trim().toUpperCase();
    const datum = String(r[7] || '').trim() || (source === 'NOS' ? 'MHHW' : 'NAVD88');
    if (!stationId) continue;

    const link = cellLink_(sh.getRange(rowNum, 1));

    try {
      if (source === 'NOS' || /tidesandcurrents\.noaa\.gov/i.test(link)) {
        results[i] = fetchNoaaWater_(stationId, link, datum, cfg.start, cfg.end);
      } else if (source === 'USGS' || /waterdata\.usgs\.gov/i.test(link)) {
        results[i] = fetchUsgsWater_(stationId, link, datum, cfg.start, cfg.end);
      } else if (source.indexOf('TPCG') >= 0) {
        log_('WARN', 'WATER', stationId,
          'TPCG not auto-filled yet. No verified automated source/mapping is configured; left blank rather than guess.');
      } else if (source.indexOf('USACE') >= 0 || source.indexOf('CPRA') >= 0 || /rivergages\.mvr\.usace\.army\.mil/i.test(link)) {
        log_('WARN', 'WATER', stationId,
          `${source || 'USACE/CPRA'} not auto-filled yet. USACE CWMS has a public historical API, but this RiverGages SID still needs a verified SID-to-CWMS-timeseries mapping; left blank rather than guess.`);
      } else {
        log_('WARN', 'WATER', stationId, `Unknown/unsupported source: ${source || '(blank)'}`);
      }
    } catch (e) {
      log_('ERROR', 'WATER', stationId, e.message || String(e));
    }

    if (i % 10 === 0) Utilities.sleep(100);
  }

  const values = results.map((p, idx) => {
    if (!p || p.value === null) return { level: '', time: '', day: '', month: '', year: '', comment: '' };
    if (p.value < PSH.QC.WATER_MIN_FT || p.value > PSH.QC.WATER_MAX_FT) {
      log_('WARN', 'WATER', String(rows[idx][0] || ''), `[AUTO-QC] Water level ${p.value} ft failed plausibility QC; blanked.`);
      return { level: '', time: '', day: '', month: '', year: '', comment: 'E' };
    }
    return {
      level: round_(p.value, 2),
      time: hhmm_(p.time),
      day: day_(p.time),
      month: month_(p.time),
      year: year_(p.time),
      comment: p.comment || ''
    };
  });

  // G = level. H = pre-populated datum (preserve). I:L = time/date.
  // M = source (preserve). N = I/E review flag.
  sh.getRange(2, 7, values.length, 1).setValues(values.map(v => [v.level]));
  sh.getRange(2, 9, values.length, 4).setValues(values.map(v => [v.time, v.day, v.month, v.year]));
  mergeQcColumn_(sh, 2, 14, values.map(v => v.comment));

  log_('INFO', 'WATER', '', `Water-level pass complete: ${rows.length} station rows checked.`);
}

function fetchNoaaWater_(sheetId, link, datum, start, end) {
  const station = extract_(link, /[?&]id=(\d{7})/i) || (/^\d{7}$/.test(sheetId) ? sheetId : null);
  if (!station) throw new Error(`Could not determine NOAA station ID from link: ${link}`);

  const begin = Utilities.formatDate(start, 'UTC', 'yyyyMMdd HH:mm');
  const finish = Utilities.formatDate(end, 'UTC', 'yyyyMMdd HH:mm');
  const apiDatum = normalizeNoaaDatum_(datum);

  const attempts = [
    { product: 'water_level', datum: apiDatum },
    { product: 'hourly_height', datum: apiDatum }
  ];

  let lastError = '';
  for (const a of attempts) {
    try {
      const json = fetchJson_(PSH.COOPS, {
        begin_date: begin,
        end_date: finish,
        station,
        product: a.product,
        datum: a.datum,
        time_zone: 'gmt',
        units: 'english',
        application: 'NWS_LIX_PSH_Automation',
        format: 'json'
      }, `NOAA CO-OPS ${station}`);

      const data = (json && json.data) || [];
      let best = null;
      data.forEach(x => {
        const v = numeric_(x.v);
        const t = parseApiTime_(x.t);
        if (v === null || !t) return;
        if (!best || v > best.value) best = { value: v, time: t, comment: '' };
      });
      if (best) return best;
      lastError = `${a.product}: no data`;
    } catch (e) {
      lastError = `${a.product}: ${e.message}`;
    }
  }
  throw new Error(`NOAA returned no usable water level (${lastError})`);
}

function fetchUsgsWater_(sheetId, link, datum, start, end) {
  const site =
    extract_(link, /monitoring-location\/([0-9A-Za-z]+)/i) ||
    extract_(link, /[?&]site_no=([0-9A-Za-z]+)/i) ||
    (/^\d{8,15}$/.test(sheetId) ? String(sheetId) : null);
  if (!site) throw new Error(`Could not determine USGS site number from link: ${link}`);

  const navdWanted = /NAVD/i.test(datum);
  // Use genuine elevation-above-datum parameters first. 72251 is *water level above
  // marsh*, not NAVD88, so it is intentionally NOT in this list.
  const properParams = navdWanted ? ['63160','62020','62615'] : ['00065'];

  for (const param of properParams) {
    const best = fetchUsgsContinuousMax_(site, param, start, end);
    if (best) return {value:best.value, time:best.time, comment:''};
  }

  if (navdWanted) {
    // Many Louisiana coastal sites historically expose only 00065. Only use a
    // stage->NAVD88 conversion for stations explicitly validated against a known
    // completed PSH. Unvalidated conversions stay blank.
    const allowKey = String(sheetId || '').replace(/\.0$/,'').toUpperCase();
    if (!PSH.USGS_NAVD88_CONVERSION_ALLOWLIST[allowKey]) {
      log_('INFO','WATER',sheetId,
        'USGS 00065 stage exists/possible, but this station is not on the validated NAVD88 conversion allowlist; left blank.');
      return null;
    }
    const stage = fetchUsgsContinuousMax_(site, '00065', start, end);
    if (stage) {
      const meta = fetchUsgsSiteDatum_(site);
      if (meta && meta.altitude !== null && meta.altitude !== 0 && /NAVD\s*88/i.test(meta.datum || '')) {
        const converted = stage.value + meta.altitude;
        if (converted >= PSH.QC.WATER_MIN_FT && converted <= PSH.QC.WATER_MAX_FT) {
          const ageDays = Math.abs(new Date().getTime() - end.getTime()) / 86400000;
          log_('WARN','WATER',sheetId,
            `Converted USGS 00065 gage height to NAVD88 using published site altitude ${meta.altitude} ft (${meta.datum}). ` +
            `${ageDays > 365 ? 'Historical event: verify the gage datum/offset in effect at event time.' : 'Verify datum before issuance.'}`);
          return {value:converted, time:stage.time, comment:'E'};
        }
      }
    }
    log_('INFO', 'WATER', sheetId,
      'No direct USGS NAVD88 elevation series and no defensible NAVD88 gage-datum conversion were available; left blank.');
    return null;
  }
  return null;
}

function fetchUsgsContinuousMax_(site, param, start, end) {
  // Modern USGS Water Data API first. A successful modern response is authoritative,
  // even if it contains no matching observations. Legacy NWIS /iv is used only when
  // the modern request itself fails, which keeps runtime down and avoids doubling calls.
  let modernFailed = false;
  try {
    const json = fetchJson_(PSH.USGS_CONTINUOUS, {
      f: 'json',
      monitoring_location_id: `USGS-${site}`,
      parameter_code: param,
      datetime: `${start.toISOString()}/${end.toISOString()}`,
      limit: 10000,
      api_key: getUsgsApiKey_() || ''
    }, `USGS modern ${site} ${param}`);
    let best = null;
    ((json && json.features) || []).forEach(f => {
      const p = (f && f.properties) || {};
      if (String(p.parameter_code || '') !== String(param)) return;
      let v = numeric_(p.value);
      const t = parseApiTime_(p.time);
      if (v === null || !t) return;
      const unit = String(p.unit_of_measure || '').toLowerCase();
      if (/meter|^m$/.test(unit) && unit.indexOf('ft') < 0) v *= 3.280839895;
      if (v < PSH.QC.WATER_MIN_FT || v > PSH.QC.WATER_MAX_FT) return;
      if (!best || v > best.value) best = {value:v,time:t};
    });
    return best;
  } catch (e) {
    modernFailed = true;
    const msg = String(e.message || e);
    if (msg.indexOf('HTTP 429') >= 0) {
      log_('WARN','WATER',site,`USGS modern API rate-limited for ${param}. Add a free USGS API key with PSH Automation -> Set USGS API Key; legacy fallback attempted.`);
    } else {
      log_('INFO','WATER',site,`Modern USGS API request failed for ${param}; trying legacy fallback: ${msg}`);
    }
  }

  if (!modernFailed) return null;
  try {
    const json = fetchJson_(PSH.USGS_IV, {
      format: 'json', sites: site,
      startDT: start.toISOString(), endDT: end.toISOString(),
      parameterCd: param, siteStatus: 'all'
    }, `USGS legacy ${site} ${param}`);
    const series = (((json || {}).value || {}).timeSeries) || [];
    let best = null;
    series.forEach(ts => {
      const code = ((((ts || {}).variable || {}).variableCode || [])[0] || {}).value || '';
      if (code !== param) return;
      const unit = (((ts || {}).variable || {}).unit || {}).unitCode || 'ft';
      (ts.values || []).forEach(g => (g.value || []).forEach(x => {
        let v = numeric_(x.value);
        const t = parseApiTime_(x.dateTime);
        if (v === null || !t) return;
        if (/^(m|meter|meters)$/i.test(unit)) v *= 3.280839895;
        if (v < PSH.QC.WATER_MIN_FT || v > PSH.QC.WATER_MAX_FT) return;
        if (!best || v > best.value) best = {value:v,time:t};
      }));
    });
    return best;
  } catch (e) {
    return null;
  }
}

function fetchUsgsSiteDatum_(site) {
  // Modern monitoring-locations endpoint replaces legacy /nwis/site metadata.
  try {
    const json = fetchJson_(PSH.USGS_LOCATIONS, {
      f: 'json',
      id: `USGS-${site}`,
      limit: 1,
      api_key: getUsgsApiKey_() || ''
    }, `USGS monitoring-location ${site}`);
    const feature = ((json && json.features) || [])[0];
    const p = (feature && feature.properties) || {};
    const altitude = numeric_(p.altitude);
    const datum = String(p.vertical_datum || p.vertical_datum_name || '');
    if (altitude === null && !datum) return null;
    return {altitude, datum};
  } catch (e) {
    return null;
  }
}

/** ---------------------------- SUMMARY / AUDIT ---------------------------- */

function pshRefreshSummary_(cfg) {
  const sh = mustSheet_(PSH.SUMMARY);
  sh.getRange('B3').setValue(cfg.storm);
  sh.getRange('B7').setValue(`${fmtDate_(cfg.start)} - ${fmtDate_(cfg.end)}`);

  const tor = mustSheet_(PSH.TORNADO);
  const last = Math.max(2, tor.getLastRow());
  const data = tor.getRange(2, 1, Math.max(1, last - 1), 11).getValues();
  const count = data.filter(r => numeric_(r[1]) !== null && numeric_(r[2]) !== null).length;
  sh.getRange('B11').setValue(count);

  const finder = sh.createTextFinder('Report Last Updated on').matchCase(false).useRegularExpression(false).findNext();
  if (finder) finder.setValue(`Report Last Updated on ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy')}:`);

  log_('INFO', 'SUMMARY', '', `Summary metadata refreshed; tornado count=${count}.`);
}

function writeSummaryConfig_(storm, start, end, direct, indirect) {
  const sh = mustSheet_(PSH.SUMMARY);
  sh.getRange('B3').setValue(storm);
  sh.getRange('B7').setValue(`${fmtDate_(start)} - ${fmtDate_(end)}`);

  const d = String(direct || '').trim();
  const ind = String(indirect || '').trim();
  if (d !== '' || ind !== '') {
    sh.getRange('B9').setValue(`${d === '' ? '[unknown]' : d} - Direct\n${ind === '' ? '[unknown]' : ind} - Indirect`);
  }
}

function writeRainWindow_(start, end) {
  const sh = mustSheet_(PSH.RAIN);
  const startCell = sh.createTextFinder('Rainfall Start Time:').matchCase(false).useRegularExpression(false).findNext();
  if (startCell) sh.getRange(startCell.getRow(), startCell.getColumn()+1)
    .setValue(Utilities.formatDate(start, 'UTC', 'HHmm \'UTC\' MMM dd yyyy'));
  const endCell = sh.createTextFinder('Rainfall End Time:').matchCase(false).useRegularExpression(false).findNext();
  if (endCell) sh.getRange(endCell.getRow(), endCell.getColumn()+1)
    .setValue(Utilities.formatDate(end, 'UTC', 'HHmm \'UTC\' MMM dd yyyy'));
}

function pshAuditSources() {
  const ss = SpreadsheetApp.getActive();
  const targets = [PSH.WIND, PSH.RAIN, PSH.WATER];
  const counts = {};

  targets.forEach(name => {
    const sh = mustSheet_(name);
    const last = findLastStationRow_(sh, 1);
    for (let row = 2; row <= last; row++) {
      const link = cellLink_(sh.getRange(row, 1));
      const source = name === PSH.WATER ? String(sh.getRange(row, 13).getDisplayValue()).trim() : '';
      const provider = providerFrom_(link, source);
      counts[provider] = (counts[provider] || 0) + 1;
      if (!link && name !== PSH.RAIN) {
        log_('WARN', 'AUDIT', sh.getRange(row, 1).getDisplayValue(), `${name}: no hyperlink found in Site ID cell.`);
      }
    }
  });

  Object.keys(counts).sort().forEach(k => log_('INFO', 'AUDIT', k, `${counts[k]} station rows`));
  SpreadsheetApp.getUi().alert('Audit complete', 'Provider counts and missing-link warnings were written to _PSH_Log.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function providerFrom_(url, source) {
  const u = String(url || '').toLowerCase();
  const s = String(source || '').toUpperCase();
  if (u.indexOf('synopticdata') >= 0 || u.indexOf('mesowest') >= 0) return 'SYNOPTIC/MESOWEST';
  if (u.indexOf('tidesandcurrents.noaa.gov') >= 0 || s === 'NOS') return 'NOAA CO-OPS';
  if (u.indexOf('waterdata.usgs.gov') >= 0 || s === 'USGS') return 'USGS';
  if (s.indexOf('TPCG') >= 0) return 'TPCG';
  if (u.indexOf('rivergages.mvr.usace.army.mil') >= 0 || s.indexOf('USACE') >= 0) return 'USACE';
  if (s.indexOf('CPRA') >= 0) return 'LA CPRA';
  if (u.indexOf('weatherflow.com') >= 0) return 'WEATHERFLOW';
  if (u.indexOf('weatherstem.com') >= 0) return 'WEATHERSTEM';
  if (!url) return source || 'NO LINK';
  return source || 'OTHER';
}


function pshCoverageSummary() {
  // Wind coverage means at least one reportable PSH field (sustained, gust, or MSLP)
  // is populated. Counting only sustained wind falsely showed WeatherSTEM as 0%.
  const specs = [
    {sheet:PSH.RAIN, networkCol:7, valueCols:[8], label:'RAIN'},
    {sheet:PSH.WATER, networkCol:13, valueCols:[7], label:'WATER'}
  ];

  const wsh = mustSheet_(PSH.WIND);
  const wlast = findLastStationRow_(wsh,1);
  const wids = wsh.getRange(2,1,wlast-1,1).getDisplayValues();
  const wnets = wsh.getRange(2,8,wlast-1,1).getDisplayValues();
  const wvals = wsh.getRange(2,11,wlast-1,13).getValues(); // K:W
  const wg = {};
  for (let i=0;i<wids.length;i++) {
    if (!String(wids[i][0]||'').trim()) continue;
    const n=String(wnets[i][0]||'UNKNOWN').trim().toUpperCase() || 'UNKNOWN';
    if (!wg[n]) wg[n]={total:0,filled:0};
    wg[n].total++;
    const has = numeric_(wvals[i][0]) !== null || numeric_(wvals[i][6]) !== null || numeric_(wvals[i][12]) !== null;
    if (has) wg[n].filled++;
  }
  Object.keys(wg).sort().forEach(n =>
    log_('INFO','COVERAGE',`WIND/${n}`,`${wg[n].filled}/${wg[n].total} rows have sustained/gust/pressure data (${Math.round(100*wg[n].filled/wg[n].total)}%).`)
  );

  specs.forEach(spec => {
    const sh = mustSheet_(spec.sheet);
    const last = findLastStationRow_(sh,1);
    const ids = sh.getRange(2,1,last-1,1).getDisplayValues();
    const nets = sh.getRange(2,spec.networkCol,last-1,1).getDisplayValues();
    const vals = spec.valueCols.map(c => sh.getRange(2,c,last-1,1).getValues());
    const g = {};
    for (let i=0;i<ids.length;i++) {
      if (!String(ids[i][0]||'').trim()) continue;
      const n = String(nets[i][0]||'UNKNOWN').trim().toUpperCase() || 'UNKNOWN';
      if (!g[n]) g[n]={total:0,filled:0};
      g[n].total++;
      if (vals.some(v => numeric_(v[i][0]) !== null)) g[n].filled++;
    }
    Object.keys(g).sort().forEach(n => log_('INFO','COVERAGE',`${spec.label}/${n}`,`${g[n].filled}/${g[n].total} populated (${Math.round(100*g[n].filled/g[n].total)}%).`));
  });
}

/**
 * Synoptic rejects an entire multi-station request when even one STID is invalid,
 * inaccessible, or absent from the account. Split failed batches recursively so
 * one stale station never kills the other 24 good stations in that request.
 */
function fetchSynopticStationsResilient_(endpoint, baseParams, stationIds, section, label) {
  if (!stationIds || !stationIds.length) return [];
  const params = Object.assign({}, baseParams, { stid: stationIds.join(',') });

  try {
    const json = fetchJson_(endpoint, params, label);
    return (json && json.STATION) || [];
  } catch (e) {
    if (stationIds.length === 1) {
      log_('WARN', section, stationIds[0], `Synoptic skipped this station: ${e.message || e}`);
      return [];
    }

    const mid = Math.floor(stationIds.length / 2);
    const left = stationIds.slice(0, mid);
    const right = stationIds.slice(mid);
    return fetchSynopticStationsResilient_(endpoint, baseParams, left, section, `${label} [split]`)
      .concat(fetchSynopticStationsResilient_(endpoint, baseParams, right, section, `${label} [split]`));
  }
}

/**
 * Repair datum cells if an earlier v0.1 run wiped column H.
 * The template convention is MHHW for NOS/NOAA CO-OPS rows and NAVD88 for
 * the remaining water-level networks. Existing nonblank datum values win.
 */
function repairWaterDatums_(sh, lastRow) {
  if (lastRow < 2) return;
  const sourceVals = sh.getRange(2, 13, lastRow - 1, 1).getDisplayValues();
  const datumRange = sh.getRange(2, 8, lastRow - 1, 1);
  const datumVals = datumRange.getDisplayValues();
  let repaired = 0;

  for (let i = 0; i < datumVals.length; i++) {
    if (String(datumVals[i][0] || '').trim()) continue;
    const source = String(sourceVals[i][0] || '').trim().toUpperCase();
    if (!source) continue;
    datumVals[i][0] = source === 'NOS' ? 'MHHW' : 'NAVD88';
    repaired++;
  }

  if (repaired) {
    datumRange.setValues(datumVals);
    log_('INFO', 'WATER', '', `Repaired ${repaired} missing datum cells (MHHW for NOS; NAVD88 otherwise).`);
  }
}


function pshRunFrancineRegression() {
  const cfg = getConfig_();
  if (!cfg || cfg.atcf !== '2024AL06') {
    SpreadsheetApp.getUi().alert('Load the Francine test window first.');
    return;
  }

  // Broad reference suite from the completed WFO LIX Francine PSH. This is
  // intentionally spread across networks/variables so a 100% score means much
  // more than a handful of hand-picked headline values.
  const checks = [
    // ASOS - wind / gust / MSLP
    {g:'ASOS',sheet:PSH.WIND,id:'KBTR',col:11,expected:27,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KBTR',col:17,expected:47,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KBTR',col:23,expected:998.5,tol:0.6},
    {g:'ASOS',sheet:PSH.WIND,id:'KMSY',col:11,expected:43,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KMSY',col:17,expected:68,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KMSY',col:23,expected:991.3,tol:0.6},
    {g:'ASOS',sheet:PSH.WIND,id:'KNEW',col:11,expected:37,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KNEW',col:17,expected:61,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KASD',col:11,expected:29,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KASD',col:17,expected:48,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KBIX',col:17,expected:41,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KGPT',col:11,expected:31,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KGPT',col:17,expected:46,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KPQL',col:11,expected:20,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KPQL',col:17,expected:38,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KMCB',col:11,expected:16,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KMCB',col:17,expected:30,tol:1},
    {g:'ASOS',sheet:PSH.WIND,id:'KMCB',col:23,expected:995.1,tol:0.6},

    // AWOS
    {g:'AWOS',sheet:PSH.WIND,id:'KREG',col:11,expected:34,tol:1},
    {g:'AWOS',sheet:PSH.WIND,id:'KREG',col:17,expected:53,tol:1},
    {g:'AWOS',sheet:PSH.WIND,id:'KHUM',col:11,expected:45,tol:1},
    {g:'AWOS',sheet:PSH.WIND,id:'KHUM',col:17,expected:61,tol:1},
    {g:'AWOS',sheet:PSH.WIND,id:'KMJD',col:11,expected:25,tol:1},
    {g:'AWOS',sheet:PSH.WIND,id:'KMJD',col:17,expected:42,tol:1},
    {g:'AWOS',sheet:PSH.WIND,id:'KPZZ',col:11,expected:69,tol:1},
    {g:'AWOS',sheet:PSH.WIND,id:'KPZZ',col:17,expected:81,tol:1},

    // CWOP / RAWS / WLON
    {g:'CWOP',sheet:PSH.WIND,id:'F7886',col:17,expected:53,tol:1},
    {g:'CWOP',sheet:PSH.WIND,id:'F7886',col:23,expected:994.2,tol:0.5},
    {g:'CWOP',sheet:PSH.WIND,id:'F0324',col:11,expected:27,tol:1},
    {g:'CWOP',sheet:PSH.WIND,id:'F0324',col:17,expected:55,tol:1},
    {g:'CWOP',sheet:PSH.WIND,id:'G2274',col:17,expected:51,tol:1},
    {g:'RAWS',sheet:PSH.WIND,id:'TS947',col:11,expected:36,tol:1},
    {g:'RAWS',sheet:PSH.WIND,id:'TS947',col:17,expected:47,tol:1},
    {g:'WLON',sheet:PSH.WIND,id:'NWCL1',col:11,expected:40,tol:1},
    {g:'WLON',sheet:PSH.WIND,id:'NWCL1',col:17,expected:52,tol:1},
    {g:'WLON',sheet:PSH.WIND,id:'SHBL1',col:11,expected:36,tol:1},
    {g:'WLON',sheet:PSH.WIND,id:'SHBL1',col:17,expected:45,tol:1},

    // WeatherSTEM direct historical validation against the issued Francine PSH.
    // v0.10's maximum valid direct minute Anemometer observation reproduces the
    // issued sustained winds for these station-specific WeatherSTEM links within
    // rounding tolerance. Keep these as real regression targets; do not replace
    // them with nearby Synoptic aliases.
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSEBRAlexBox',col:11,expected:25,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSEBRAlexBox',col:17,expected:30,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSEBRTigerStadium',col:11,expected:44,tol:1},
    // Known unresolved discrepancy: the 2026-10-05 v0.10 run returned ~44.3 kt
    // from the native WeatherSTEM 10 Minute Wind Gust sensor vs 48 kt issued.
    // Keep the authoritative 48 kt target so regression exposes the mismatch.
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSEBRTigerStadium',col:17,expected:48,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSNOLakefront',col:11,expected:47,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSNOLakefront',col:17,expected:54,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSNOMidCIty',col:11,expected:44,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSNOMidCIty',col:17,expected:53,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSNOBayouSauvage',col:11,expected:45,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSNOBayouSauvage',col:17,expected:59,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSSCEOC',col:11,expected:37,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSSCEOC',col:17,expected:46,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSSCLuling',col:11,expected:38,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSSCLuling',col:17,expected:46,tol:1},
    // Retain a few additional issued gust sentinels for broader network coverage.
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSNOOldAurora',col:17,expected:27,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSNOIrishChannel',col:17,expected:55,tol:1},
    {g:'WeatherSTEM',sheet:PSH.WIND,id:'WSNOWarehouse',col:17,expected:52,tol:1},

    // Rainfall
    {g:'RAIN-CoCoRaHS',sheet:PSH.RAIN,id:'LA-ST-11',col:8,expected:9.63,tol:0.12},
    {g:'RAIN-CoCoRaHS',sheet:PSH.RAIN,id:'LA-JF-20',col:8,expected:9.48,tol:0.12},
    {g:'RAIN-CoCoRaHS',sheet:PSH.RAIN,id:'LA-SC-06',col:8,expected:9.22,tol:0.12},
    {g:'RAIN-CoCoRaHS',sheet:PSH.RAIN,id:'LA-JF-21',col:8,expected:8.77,tol:0.12},
    {g:'RAIN-CoCoRaHS',sheet:PSH.RAIN,id:'LA-ST-23',col:8,expected:8.66,tol:0.12},
    {g:'RAIN-COOP',sheet:PSH.RAIN,id:'ABRL1',col:8,expected:8.57,tol:0.20},
    {g:'RAIN-COOP',sheet:PSH.RAIN,id:'LIX',col:8,expected:7.93,tol:0.25},
    {g:'RAIN-HADS',sheet:PSH.RAIN,id:'MSVL1',col:8,expected:8.26,tol:0.10},
    {g:'RAIN-RAWS',sheet:PSH.RAIN,id:'TS947',col:8,expected:8.44,tol:0.10},
    {g:'RAIN-ASOS',sheet:PSH.RAIN,id:'MSY',col:8,expected:7.98,tol:0.15},

    // NOAA and selected USGS water levels
    {g:'WATER-NOS',sheet:PSH.WATER,id:'WYCM6',col:7,expected:4.81,tol:0.08},
    {g:'WATER-NOS',sheet:PSH.WATER,id:'NWCL1',col:7,expected:3.63,tol:0.08},
    {g:'WATER-NOS',sheet:PSH.WATER,id:'PNLM6',col:7,expected:3.13,tol:0.08},
    {g:'WATER-NOS',sheet:PSH.WATER,id:'SHBL1',col:7,expected:2.91,tol:0.10},
    {g:'WATER-NOS',sheet:PSH.WATER,id:'GISL1',col:7,expected:1.78,tol:0.08},
    {g:'WATER-NOS',sheet:PSH.WATER,id:'PTFL1',col:7,expected:1.74,tol:0.08},
    {g:'WATER-USGS',sheet:PSH.WATER,id:'PPAL1',col:7,expected:2.38,tol:0.15},
    {g:'WATER-USGS',sheet:PSH.WATER,id:'MSVL1',col:7,expected:5.26,tol:0.15}
  ];

  let pass=0, fail=0, missing=0;
  const groups={};
  checks.forEach(c => {
    if (!groups[c.g]) groups[c.g]={pass:0,fail:0,missing:0};
    const sh = mustSheet_(c.sheet);
    const row = c.g === 'RAIN-CoCoRaHS'
      ? findRainRowByAlias_(sh, c.id, 'COCORAHS')
      : findRowById_(sh, c.id);
    if (!row) {
      missing++; groups[c.g].missing++;
      log_('WARN','REGRESSION',c.id,`${c.g}: station missing from automated sheet.`);
      return;
    }
    const v = numeric_(sh.getRange(row,c.col).getValue());
    if (c.expectedBlank) {
      if (v===null) { pass++; groups[c.g].pass++; log_('INFO','REGRESSION',c.id,`PASS ${c.g}: intentionally blank (unsafe alias/substitution rejected).`); }
      else { fail++; groups[c.g].fail++; log_('WARN','REGRESSION',c.id,`FAIL ${c.g}: expected blank for unsafe alias guard, got ${v}.`); }
      return;
    }
    if (v === null) {
      missing++; groups[c.g].missing++;
      log_('WARN','REGRESSION',c.id,`${c.g}: no automated value; reference=${c.expected}.`);
      return;
    }
    const err = Math.abs(v-c.expected);
    if (err <= c.tol) {
      pass++; groups[c.g].pass++;
      log_('INFO','REGRESSION',c.id,`PASS ${c.g}: ${v} vs ${c.expected} (tol ${c.tol}).`);
    } else {
      fail++; groups[c.g].fail++;
      log_('WARN','REGRESSION',c.id,`FAIL ${c.g}: ${v} vs ${c.expected}; |error|=${round_(err,2)} > ${c.tol}.`);
    }
  });

  // Hard sanity checks across the full wind table.
  const wsh = mustSheet_(PSH.WIND);
  const wlast = findLastStationRow_(wsh,1);
  const windVals = wsh.getRange(2,11,wlast-1,13).getValues();
  windVals.forEach((r,i) => {
    const id = wsh.getRange(i+2,1).getDisplayValue();
    if (numeric_(r[0]) !== null && numeric_(r[0]) > PSH.QC.WIND_MAX_KT) { fail++; log_('ERROR','REGRESSION',id,'Impossible sustained wind survived QC.'); }
    if (numeric_(r[6]) !== null && numeric_(r[6]) > PSH.QC.GUST_MAX_KT) { fail++; log_('ERROR','REGRESSION',id,'Impossible gust survived QC.'); }
    if (numeric_(r[12]) !== null && (numeric_(r[12]) < PSH.QC.PRESSURE_MIN_MB || numeric_(r[12]) > PSH.QC.PRESSURE_MAX_MB)) { fail++; log_('ERROR','REGRESSION',id,'Impossible pressure survived QC.'); }
    if (numeric_(r[0]) !== null && numeric_(r[6]) !== null && numeric_(r[6]) + PSH.QC.WIND_GUST_EPSILON_KT < numeric_(r[0])) {
      fail++; log_('ERROR','REGRESSION',id,'Gust lower than sustained wind survived cross-variable QC.');
    }
  });

  // DELIBERATE DEVIATION: the issued 2024 Francine PSH reported BPPL1 = 6.36 ft
  // NAVD88 and BDML1 = 2.40 ft NAVD88. These two rows are permanently manual by
  // design because the automation does not have a defensible historical stage->NAVD88
  // conversion for them. Do not "fix" this regression by weakening the datum QC.
  ['BPPL1','BDML1'].forEach(id => {
    const w = mustSheet_(PSH.WATER), row = findRowById_(w,id);
    if (row && numeric_(w.getRange(row,7).getValue()) !== null) {
      fail++; log_('ERROR','REGRESSION',id,'Known unvalidated USGS stage->NAVD88 conversion was populated.');
    }
  });

  Object.keys(groups).sort().forEach(g => {
    const x=groups[g];
    log_('INFO','REGRESSION-SUMMARY',g,`${x.pass} pass, ${x.fail} fail, ${x.missing} missing.`);
  });
  const total = pass+fail+missing;
  const score = total ? Math.round(100*pass/total) : 0;
  log_('INFO','REGRESSION','',`Francine expanded regression: ${pass} pass, ${fail} fail, ${missing} missing; score=${score}%.`);
  SpreadsheetApp.getUi().alert('Francine regression',
    `${pass} passed
${fail} failed
${missing} missing

Expanded check score: ${score}%
See _PSH_Log for network-by-network details.`,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function findRowById_(sh, id) {
  const last = findLastStationRow_(sh,1);
  const vals = sh.getRange(2,1,last-1,1).getDisplayValues();
  const target = String(id).trim().toUpperCase();
  for (let i=0;i<vals.length;i++) if (String(vals[i][0]).trim().toUpperCase()===target) return i+2;
  return 0;
}

function findRainRowByAlias_(sh, id, network) {
  const last = findLastStationRow_(sh,1);
  const vals = sh.getRange(2,1,last-1,7).getDisplayValues();
  const targets = new Set(rainIdAliasKeys_(id, network));
  for (let i=0;i<vals.length;i++) {
    const rowNet=String(vals[i][6] || '').trim().toUpperCase();
    if (network && rowNet && rowNet !== String(network).trim().toUpperCase()) continue;
    const aliases=rainIdAliasKeys_(vals[i][0], rowNet || network);
    if (aliases.some(k => targets.has(k))) return i+2;
  }
  return 0;
}

function clearAutoQc_(sh, startRow, endRow, startCol, endCol) {
  if (endRow < startRow) return;
  const range = sh.getRange(startRow, startCol, endRow-startRow+1, endCol-startCol+1);
  const vals = range.getValues();
  let changed=false;
  vals.forEach(r => {
    for (let i=0;i<r.length;i++) {
      if (String(r[i] || '').indexOf('[AUTO-QC]') >= 0) { r[i]=''; changed=true; }
    }
  });
  if (changed) range.setValues(vals);
}

function mergeQcColumn_(sh, startRow, col, newValues) {
  if (!newValues || !newValues.length) return;
  const range = sh.getRange(startRow, col, newValues.length, 1);
  const existing = range.getValues();
  let changed = false;
  for (let i=0;i<newValues.length;i++) {
    const next = newValues[i];
    const old = String(existing[i][0] === null || existing[i][0] === undefined ? '' : existing[i][0]);
    if (next !== '' && next !== null && next !== undefined) {
      existing[i][0] = next; changed = true;
    } else if (old.indexOf('[AUTO-QC]') >= 0) {
      existing[i][0] = ''; changed = true;
    }
    // Otherwise preserve human-entered I/E/comment content exactly as-is.
  }
  if (changed) range.setValues(existing);
}

function applyAutoQc_(sh, qcWrites, startRow) {
  qcWrites.forEach((qc, i) => {
    if (!qc || !qc.messages || !qc.messages.length) return;
    const row = startRow+i;
    const flagCell = sh.getRange(row,28);
    const varCell = sh.getRange(row,29);
    const remCell = sh.getRange(row,30);
    if (!flagCell.getDisplayValue()) flagCell.setValue('E');
    if (!varCell.getDisplayValue()) varCell.setValue(qc.vars.length ? qc.vars.join('') : 'A');
    const old = remCell.getDisplayValue();
    const msg = `[AUTO-QC] ${qc.messages.join('; ')}`;
    remCell.setValue(old ? `${old} | ${msg}` : msg);
    log_('WARN','WIND',sh.getRange(row,1).getDisplayValue(),msg);
  });
}


/** ---------------------------- LOGGING ---------------------------- */

function pshShowLog() {
  const sh = ensureLogSheet_();
  sh.showSheet();
  SpreadsheetApp.setActiveSheet(sh);
}

function pshClearLog() {
  const sh = ensureLogSheet_();
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 5).clearContent();
}

function ensureLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(PSH.LOG);
  if (!sh) {
    sh = ss.insertSheet(PSH.LOG);
    sh.getRange('A1:E1').setValues([['Timestamp', 'Level', 'Section', 'Station', 'Message']]);
    sh.getRange('A1:E1').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 155);
    sh.setColumnWidth(2, 70);
    sh.setColumnWidth(3, 100);
    sh.setColumnWidth(4, 110);
    sh.setColumnWidth(5, 650);
  }
  return sh;
}

function log_(level, section, station, message) {
  const sh = ensureLogSheet_();
  sh.appendRow([new Date(), level, section, station || '', String(message || '')]);
}

/** ---------------------------- HELPERS ---------------------------- */

function getConfig_() {
  const props = PropertiesService.getDocumentProperties();
  const storm = props.getProperty(PSH.PROP_STORM);
  const startS = props.getProperty(PSH.PROP_START);
  const endS = props.getProperty(PSH.PROP_END);
  if (!storm || !startS || !endS) {
    SpreadsheetApp.getUi().alert('Not configured', 'Run PSH Automation -> Configure Storm first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
  const start = new Date(startS);
  const end = new Date(endS);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) throw new Error('Saved storm times are invalid. Re-run Configure Storm.');
  const rainStartS = props.getProperty(PSH.PROP_RAIN_START) || startS;
  const rainEndS = props.getProperty(PSH.PROP_RAIN_END) || endS;
  const rainStart = new Date(rainStartS);
  const rainEnd = new Date(rainEndS);
  if (isNaN(rainStart.getTime()) || isNaN(rainEnd.getTime()) || rainEnd <= rainStart) {
    throw new Error('Saved rainfall times are invalid. Re-run Configure Rainfall Window.');
  }
  return { storm, atcf: props.getProperty(PSH.PROP_ATCF) || '', start, end, rainStart, rainEnd };
}

function getSynopticToken_() {
  return PropertiesService.getDocumentProperties().getProperty(PSH.PROP_SYNOPTIC) || '';
}

function getWeatherStemKey_() {
  return PropertiesService.getDocumentProperties().getProperty(PSH.PROP_WEATHERSTEM) || '';
}

function getUsgsApiKey_() {
  return PropertiesService.getDocumentProperties().getProperty(PSH.PROP_USGS) || '';
}

function mustSheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error(`Required sheet not found: ${name}`);
  return sh;
}

function findLastStationRow_(sh, col) {
  const last = sh.getLastRow();
  if (last < 2) return last;
  const vals = sh.getRange(1, col, last, 1).getDisplayValues().flat();
  let found = 1;
  for (let i = 1; i < vals.length; i++) {
    const v = String(vals[i] || '').trim();
    if (!v) continue;
    // Stop before obvious footer/update labels on narrative sheets.
    if (/^(Latest Update:|Update Details:|Remarks:)/i.test(v)) break;
    found = i + 1;
  }
  return found;
}

function promptRequired_(ui, title, message) {
  const r = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const s = r.getResponseText().trim();
  if (!s) {
    ui.alert('A value is required.');
    return null;
  }
  return s;
}

function promptOptional_(ui, title, message) {
  const r = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  return r.getResponseText().trim();
}

function parseUtc_(text) {
  let s = String(text || '').trim();
  if (!s) return null;
  if (/^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}$/.test(s)) s = s.replace(' ', 'T') + ':00Z';
  else if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(s)) s += ':00Z';
  else if (/^\d{4}-\d{2}-\d{2}$/.test(s)) s += 'T00:00:00Z';
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function synopticTime_(d) {
  return Utilities.formatDate(d, 'UTC', 'yyyyMMddHHmm');
}

function fmtDate_(d) {
  return Utilities.formatDate(d, 'UTC', 'MM/dd/yyyy');
}

function hhmm_(d) {
  if (!d) return '';
  return Utilities.formatDate(d, 'UTC', 'HHmm');
}
function day_(d) { return d ? Number(Utilities.formatDate(d, 'UTC', 'd')) : ''; }
function month_(d) { return d ? Number(Utilities.formatDate(d, 'UTC', 'M')) : ''; }
function year_(d) { return d ? Number(Utilities.formatDate(d, 'UTC', 'yyyy')) : ''; }

function parseApiTime_(s) {
  if (!s) return null;
  let t = String(s).trim();
  // NOAA commonly returns "YYYY-MM-DD HH:mm" with GMT requested.
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(t)) t = t.replace(' ', 'T') + ':00Z';
  const d = new Date(t);
  return isNaN(d.getTime()) ? null : d;
}

function numeric_(v) {
  if (v === null || v === undefined || v === '' || v === 'M' || v === 'NaN') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function valueOrBlank_(v) { return v === null || v === undefined ? '' : v; }
function round_(v, n) {
  if (v === null || v === undefined || !Number.isFinite(Number(v))) return '';
  const f = Math.pow(10, n);
  return Math.round(Number(v) * f) / f;
}

function chunks_(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}
function unique_(arr) { return [...new Set(arr.filter(Boolean))]; }

function synopticId_(rawId, network) {
  let id = String(rawId || '').trim().toUpperCase();
  const net = String(network || '').toUpperCase();
  // Rainfall airport rows often omit the ICAO K prefix while Synoptic uses it.
  if (/^(ASOS|AWOS)$/.test(net) && /^[A-Z]{3}$/.test(id)) id = 'K' + id;
  // Synoptic exposes many COOP stations with a COOP-prefixed STID (for example
  // template LIX -> Synoptic COOPLIX). Query the API using that canonical STID so
  // the configured rainfall-window total maps back to the existing template row.
  if (net === 'COOP' && id && !/^COOP/.test(id)) id = 'COOP' + id;
  return id;
}

function normalizeNoaaDatum_(datum) {
  const d = String(datum || '').trim().toUpperCase();
  if (d === 'NAVD88' || d === 'NAVD 88') return 'NAVD';
  const allowed = ['STND', 'MHHW', 'MHW', 'MTL', 'MSL', 'MLW', 'MLLW', 'NAVD'];
  return allowed.indexOf(d) >= 0 ? d : 'MHHW';
}

function cellLink_(range) {
  try {
    const rich = range.getRichTextValue();
    if (rich) {
      const direct = rich.getLinkUrl();
      if (direct) return direct;
      const runs = rich.getRuns();
      for (const run of runs) {
        const u = run.getLinkUrl();
        if (u) return u;
      }
    }
  } catch (e) {}

  // Fallback for =HYPERLINK("url", "label") formulas.
  const f = range.getFormula();
  const m = f && f.match(/^=HYPERLINK\("([^"]+)"/i);
  return m ? m[1] : '';
}

function extract_(s, re) {
  const m = String(s || '').match(re);
  return m ? m[1] : '';
}

function buildUrl_(base, params) {
  const query = Object.keys(params || {})
    .filter(k => params[k] !== null && params[k] !== undefined && params[k] !== '')
    .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])))
    .join('&');
  return base + (query ? (base.indexOf('?') >= 0 ? '&' : '?') + query : '');
}

function fetchResponseResilient_(url, label) {
  const maxAttempts = 4;
  let last = null;
  for (let attempt=1; attempt<=maxAttempts; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, {
        method: 'get',
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { 'User-Agent': 'NWS-LIX-PSH-Automation/0.6' }
      });
      const code = resp.getResponseCode();
      if ((code === 429 || code >= 500) && attempt < maxAttempts) {
        Utilities.sleep(400 * Math.pow(2, attempt-1));
        continue;
      }
      return resp;
    } catch (e) {
      last = e;
      if (attempt < maxAttempts) { Utilities.sleep(400 * Math.pow(2, attempt-1)); continue; }
    }
  }
  throw new Error(`${label || 'HTTP request'} failed after retries: ${last ? (last.message || last) : 'unknown error'}`);
}

function fetchText_(base, params, label) {
  const url = buildUrl_(base, params);
  const resp = fetchResponseResilient_(url, label);
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error(`${label || 'HTTP request'} failed: HTTP ${code}; ${body.slice(0,400)}`);
  return body;
}

function fetchJson_(base, params, label) {
  const text = fetchText_(base, params, label);
  let json;
  try { json = JSON.parse(text); }
  catch (e) { throw new Error(`${label || 'HTTP request'} did not return JSON: ${text.slice(0, 400)}`); }

  if (json && json.SUMMARY && Number(json.SUMMARY.RESPONSE_CODE) > 1) {
    throw new Error(`${label || 'API'}: ${json.SUMMARY.RESPONSE_MESSAGE || 'API error'}`);
  }
  // ACIS returns {error:...}; some other services use an error object alongside no data.
  if (json && json.error && !json.data && !json.features && !json.results) {
    const msg = typeof json.error === 'string' ? json.error : JSON.stringify(json.error);
    throw new Error(`${label || 'API'}: ${msg}`);
  }
  return json;
}