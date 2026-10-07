/**
 * Pure rainfall logic shared by GitHub tests and the Apps Script runtime.
 * No SpreadsheetApp, UrlFetchApp, PropertiesService, or Node APIs belong here.
 */
const PSHRainCore = (() => {
  const SOURCE_PRIORITY = Object.freeze({
    Synoptic: 400,
    'CoCoRaHS': 400,
    'IEM-CoCoRaHS': 300,
    ACIS: 100
  });

  function numeric(value) {
    if (value === null || value === undefined || value === '' || value === 'M' || value === 'NaN') return null;
    const n = Number(value);
    return Number.isFinite(n) ? n : null;
  }

  function unique(values) {
    return [...new Set((values || []).filter(Boolean))];
  }

  function canonicalCocorahsId(id) {
    const u = String(id || '').trim().toUpperCase();
    const m = u.match(/^([A-Z]{2})-([A-Z]{2})-(\d+)$/);
    if (!m) return u;
    return \`\${m[1]}-\${m[2]}-\${Number(m[3])}\`;
  }

  function rainIdAliasKeys(id, network) {
    const u = String(id || '').trim().toUpperCase();
    const net = String(network || '').trim().toUpperCase();
    if (!u) return [];
    const keys = [u];

    if (/^K[A-Z0-9]{3}$/.test(u)) keys.push(u.substring(1));
    else if (/^[A-Z0-9]{3}$/.test(u)) keys.push('K' + u);

    if (/^COOP[A-Z0-9]+$/.test(u)) keys.push(u.substring(4));
    if (net === 'COOP' && !/^COOP/.test(u)) keys.push('COOP' + u);

    if (net === 'COCORAHS' || /^[A-Z]{2}-[A-Z]{2}-\d+$/.test(u)) {
      keys.push(canonicalCocorahsId(u));
    }
    return unique(keys);
  }

  function normalizeRainNetwork(value) {
    const s = String(value || '').trim().toUpperCase();
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

  function synopticRequestId(rawId, network) {
    let id = String(rawId || '').trim().toUpperCase();
    const net = String(network || '').trim().toUpperCase();
    if (/^(ASOS|AWOS)$/.test(net) && /^[A-Z]{3}$/.test(id)) id = 'K' + id;
    return id;
  }

  function aliasesOverlap(aId, aNetwork, bId, bNetwork) {
    const a = new Set(rainIdAliasKeys(aId, aNetwork));
    return rainIdAliasKeys(bId, bNetwork).some(key => a.has(key));
  }

  function bestPrecipTotal(obs) {
    const observations = obs || {};
    const list = observations.precipitation;
    if (Array.isArray(list) && list.length) {
      const valid = list
        .map(item => ({ total: numeric(item && item.total), count: numeric(item && item.count) || 0 }))
        .filter(item => item.total !== null);
      if (!valid.length) return null;
      valid.sort((a, b) => b.count - a.count);
      return valid[0].total;
    }
    const keys = Object.keys(observations).filter(k => k.indexOf('total_precip_value_') === 0);
    for (const key of keys) {
      const value = numeric(observations[key]);
      if (value !== null) return value;
    }
    return null;
  }

  function buildSynopticRowIndex(rows) {
    const byAlias = {};
    const requested = [];
    (rows || []).forEach((row, index) => {
      const id = String(row && row.id || '').trim();
      const network = String(row && row.network || '').trim().toUpperCase();
      if (!id || network === 'COCORAHS') return;
      const rowNum = row && row.rowNum !== undefined ? row.rowNum : index;
      const requestId = synopticRequestId(id, network);
      requested.push(requestId);
      rainIdAliasKeys(requestId, network).forEach(key => {
        const u = String(key).toUpperCase();
        if (!byAlias[u]) byAlias[u] = [];
        byAlias[u].push(rowNum);
      });
    });
    Object.keys(byAlias).forEach(key => { byAlias[key] = unique(byAlias[key]); });
    return { requested: unique(requested), byAlias };
  }

  function matchSynopticRows(station, byAlias) {
    const stid = String(station && station.STID || '').trim().toUpperCase();
    if (!stid) return [];
    const returnedNetwork = normalizeRainNetwork(
      station.MNET_SHORTNAME || station.SOURCE || station.MNET_ID || ''
    );
    return unique(
      rainIdAliasKeys(stid, returnedNetwork)
        .flatMap(key => (byAlias && byAlias[String(key).toUpperCase()]) || [])
    );
  }

  function parseCsvLine(line) {
    const out = [];
    let cur = '';
    let quoted = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') {
        if (quoted && line[i + 1] === '"') { cur += '"'; i++; }
        else quoted = !quoted;
      } else if (c === ',' && !quoted) {
        out.push(cur);
        cur = '';
      } else {
        cur += c;
      }
    }
    out.push(cur);
    return out;
  }

  function parseIemDailyCsv(text) {
    const lines = String(text || '').trim().split(/\r?\n/).filter(Boolean);
    const totals = {};
    if (!lines.length) return totals;
    const header = parseCsvLine(lines[0]).map(x => String(x).trim().toLowerCase());
    const stationCol = header.indexOf('station');
    const precipCol = header.indexOf('precip_in');
    if (stationCol < 0 || precipCol < 0) {
      throw new Error('IEM daily CoCoRaHS CSV missing station/precip_in columns.');
    }
    for (let i = 1; i < lines.length; i++) {
      const cols = parseCsvLine(lines[i]);
      const id = canonicalCocorahsId(cols[stationCol]);
      if (!id) continue;
      let value = numeric(cols[precipCol]);
      if (value === null) continue;
      if (Math.abs(value - 0.0001) < 1e-8) value = 0;
      if (value < 0 || value > 30) continue;
      totals[id] = (totals[id] || 0) + value;
    }
    return totals;
  }

  function candidatePriority(source) {
    return SOURCE_PRIORITY[source] || 0;
  }

  function chooseCandidate(existing, candidate) {
    if (!candidate || numeric(candidate.value) === null) return existing || null;
    if (!existing || numeric(existing.value) === null) return candidate;
    return candidatePriority(candidate.source) > candidatePriority(existing.source)
      ? candidate
      : existing;
  }

  function canonicalValueMap(values, cocorahs = false) {
    const out = {};
    Object.entries(values || {}).forEach(([key, value]) => {
      const k = cocorahs ? canonicalCocorahsId(key) : String(key).trim().toUpperCase();
      const v = numeric(value);
      if (k && v !== null) out[k] = v;
    });
    return out;
  }

  function resolveRainfallSources(input) {
    const rows = (input && input.rows || []).map((row, index) => ({
      id: String(row.id || '').trim(),
      network: String(row.network || '').trim().toUpperCase(),
      rowNum: index
    }));
    const index = buildSynopticRowIndex(rows);
    const resolved = new Array(rows.length).fill(null);

    function apply(rowNum, value, source) {
      const v = numeric(value);
      if (v === null) return;
      resolved[rowNum] = chooseCandidate(resolved[rowNum], { value: v, source });
    }

    function applySynoptic(stations, source) {
      (stations || []).forEach(station => {
        const total = bestPrecipTotal(station.OBSERVATIONS || {});
        if (total === null) return;
        matchSynopticRows(station, index.byAlias).forEach(rowNum => apply(rowNum, total, source));
      });
    }

    applySynoptic(input.synopticDirect, 'Synoptic');
    applySynoptic(input.synopticBulk, 'Synoptic');

    const official = canonicalValueMap(input.cocorahsOfficial, true);
    const iem = canonicalValueMap(input.iemCocorahas || input.iemCocorahs, true);
    const acis = canonicalValueMap(input.acis, false);

    rows.forEach((row, rowNum) => {
      if (row.network === 'COCORAHS') {
        const key = canonicalCocorahsId(row.id);
        if (official[key] !== undefined) apply(rowNum, official[key], 'CoCoRaHS');
        if (iem[key] !== undefined) apply(rowNum, iem[key], 'IEM-CoCoRaHS');
      }
      if (/^(COOP|HADS)$/.test(row.network)) {
        const key = row.id.toUpperCase();
        if (acis[key] !== undefined) apply(rowNum, acis[key], 'ACIS');
      }
    });

    const byId = {};
    rows.forEach((row, rowNum) => {
      byId[row.id] = resolved[rowNum]
        ? { value: resolved[rowNum].value, source: resolved[rowNum].source }
        : null;
    });
    return { rows: resolved, byId };
  }

  return Object.freeze({
    SOURCE_PRIORITY,
    numeric,
    unique,
    canonicalCocorahsId,
    rainIdAliasKeys,
    normalizeRainNetwork,
    synopticRequestId,
    aliasesOverlap,
    bestPrecipTotal,
    buildSynopticRowIndex,
    matchSynopticRows,
    parseIemDailyCsv,
    chooseCandidate,
    resolveRainfallSources
  });
})();
