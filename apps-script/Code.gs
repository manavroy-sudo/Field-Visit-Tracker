/**
 * Field Visit Tracker — Google Apps Script backend
 *
 * Deploy this bound to the "Field Visit Tracker Backend" Google Sheet:
 *   https://docs.google.com/spreadsheets/d/1hD_X7C8_A3E489H-fBCse76rG102OAPdrO9JXeplOnE/edit
 *
 * Setup:
 * 1. Open the Sheet above.
 * 2. Extensions -> Apps Script.
 * 3. Delete any starter code, paste this whole file in.
 * 4. Change API_KEY below to a random string of your own.
 * 5. Deploy -> New deployment -> type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 * 6. Copy the deployment URL (ends in /exec).
 * 7. In index.html, set APPS_SCRIPT_URL to that URL and APPS_SCRIPT_KEY to the
 *    same string you used for API_KEY below.
 */

const SHEET_ID = '1hD_X7C8_A3E489H-fBCse76rG102OAPdrO9JXeplOnE';
const API_KEY = 'CHANGE_ME_TO_A_RANDOM_STRING';

const TABS = {
  pm: 'Partner Meet Responses',
  tc: 'Team Connect Responses',
  im: 'Insurer Meet Responses'
};

const LOGIN_TAB = 'Login Log';
const LOGIN_HEADERS = ['timestamp', 'emp_id', 'emp_name', 'role', 'zone', 'region'];

// Live roster + travel plan source — a separate sheet the leaders themselves edit.
// Must be shared (at least Viewer) with whichever Google account this script is deployed as.
const TRAVEL_SHEET_ID = '1wtvnrhCemuwEqHxO_dxhF9NEJ92M3BMV1Gy0zrrCDDg';
const TRAVEL_TAB_NAME = "Aug'26";

// Google Chat space webhooks. CHAT_WEBHOOK_URL is the main/production target
// (the daily 10am trigger posts here); TESTING_WEBHOOK_URL is used only when
// a doPost caller explicitly passes {target:'testing'}, so new report types
// can be verified in the Testing space before going out to the real one.
// Get either from: space name -> Apps & integrations -> Webhooks -> Add webhook.
const CHAT_WEBHOOK_URL = 'PASTE_YOUR_CHAT_WEBHOOK_URL_HERE';
const TESTING_WEBHOOK_URL = 'PASTE_YOUR_TESTING_CHAT_WEBHOOK_URL_HERE';

// A separate spreadsheet (not the backend Sheet) that already collects/fetches
// every visit-form response. Column C holds the leader's full name — counting
// how many times a name appears there gives that leader's total forms filled.
const FORM_RESPONSES_SHEET_ID = '1vAI93lXcQetHZtk7wIbqIt-Zf-jG2S8etuN4-jafgyg';

// Exact field names written/read for each visit type. Order defines column order.
const LOCATION_FIELDS = ['location_lat','location_lng','location_distance_km','location_verified'];
const HEADERS = {
  pm: ['submission_id','ts','emp_id','emp_name','designation','zone','region','planned_date','city',
       'partner_gid','partner_name','partner_contact','exist_new','bucket','status','purpose',
       'purpose_other','active_challenges','challenge_other','trend','confidence','inactive_reason',
       'lob','support','outcome','notes','edited_ts','edited_by'].concat(LOCATION_FIELDS),
  tc: ['submission_id','ts','emp_id','emp_name','designation','zone','region','planned_date','city',
       'mode','branch','bcity','meet_type','objective','team_size','agenda','actions','health',
       'challenge','challenge_other','edited_ts','edited_by'].concat(LOCATION_FIELDS),
  im: ['submission_id','ts','emp_id','emp_name','designation','zone','region','planned_date','city',
       'insurer','purpose','discussion_areas','outcomes_text','outcome','edited_ts','edited_by'].concat(LOCATION_FIELDS)
};

function getSheet_(type) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(TABS[type]);
  if (!sh) {
    sh = ss.insertSheet(TABS[type]);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS[type]);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getLoginSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(LOGIN_TAB);
  if (!sh) sh = ss.insertSheet(LOGIN_TAB);
  if (sh.getLastRow() === 0) sh.appendRow(LOGIN_HEADERS);
  return sh;
}

function logLogin_(data) {
  const sh = getLoginSheet_();
  sh.appendRow([
    data.timestamp || new Date().toISOString(),
    data.emp_id || '',
    data.emp_name || '',
    data.role || '',
    data.zone || '',
    data.region || ''
  ]);
  return json_({ status: 'success' });
}

function getLogins_() {
  const sh = getLoginSheet_();
  const rows = sh.getDataRange().getValues();
  const logins = [];
  if (rows.length < 2) return logins;
  const headers = rows[0];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue;
    const rec = {};
    headers.forEach((h, idx) => { rec[h] = row[idx]; });
    logins.push(rec);
  }
  return logins;
}

// Known misspellings/abbreviations seen in the live Travel Plan sheet, mapped
// to their correct name so geocoding actually finds them. Extend this list as
// new typos show up in future months' plans — matching is case-insensitive.
const CITY_ALIASES = {
  'bhuneshwar': 'Bhubaneswar', 'bhubneshwar': 'Bhubaneswar', 'bhubaneshwar': 'Bhubaneswar',
  'coochbihar': 'Cooch Behar', 'coochbehar': 'Cooch Behar',
  'n 24 pgn': 'North 24 Parganas', 'south 24 pgn': 'South 24 Parganas', 'north 24 pgn': 'North 24 Parganas',
  'fardabad': 'Faridabad', 'faridabd': 'Faridabad', 'fardabd': 'Faridabad',
  'kkamgaon': 'Khamgaon', 'noda': 'Noida', 'nandad': 'Nanded',
  'ferozpur': 'Firozpur', 'bhathinda': 'Bathinda',
  'nalagargh & baddi': 'Nalagarh', 'nalagargh': 'Nalagarh',
  'vihskapatnam': 'Visakhapatnam', 'chithoor': 'Chittoor', 'tirupathi': 'Tirupati',
  'cochin': 'Kochi', 'sitamahri': 'Sitamarhi', 'nababganj': 'Nawabganj',
  'amathi': 'Amethi', 'raibareily': 'Raebareli',
  'krishna nagar': 'Krishnanagar', 'baraipur': 'Baruipur', 'silligudi': 'Siliguri',
  'asansoal': 'Asansol', 'bardhman': 'Bardhaman', 'kareemnagar': 'Karimnagar',
  'doddabllapur karnataka': 'Doddaballapur', 'tiruvalluar': 'Tiruvallur', 'vijianagaram': 'Vizianagaram',
  'pankchkula': 'Panchkula'
};

function normalizeCity_(name) {
  let s = String(name || '').trim().replace(/\s+/g, ' ');
  s = s.replace(/\/.*$/, '').trim(); // strip "/neem ka dhana"-style suffixes
  const key = s.toLowerCase();
  if (CITY_ALIASES[key]) return CITY_ALIASES[key];
  // Title-case so "gurugram"/"GURUGRAM"/"Gurugram" all collapse to one cache
  // entry instead of being geocoded and stored as three separate "cities".
  return key.replace(/\b\w/g, c => c.toUpperCase());
}

function normalizeName_(s) {
  return String(s || '').toLowerCase().replace(/[. ]/g, ' ').replace(/\s+/g, ' ').trim();
}

/**
 * Parses the live "M&H Travel Plan" sheet's Aug'26 tab. It's a manually-edited
 * sheet with merged header cells and inconsistent columns, so this is defensive:
 * it locates headers by text match rather than fixed column letters, and skips
 * (rather than guesses on) any row it can't confidently map to an Employee ID.
 */
function findTravelSheet_(ss) {
  let sh = ss.getSheetByName(TRAVEL_TAB_NAME);
  if (sh) return sh;
  // Tolerate apostrophe-character differences (straight vs curly quote) between
  // what's typed here and what's actually in the live sheet's tab name.
  const all = ss.getSheets();
  for (let i = 0; i < all.length; i++) {
    if (/^aug.?26$/i.test(all[i].getName().replace(/\s+/g, ''))) return all[i];
  }
  return null;
}

function dateCellToKey_(v, tz) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, tz, 'd-MMM');
  }
  const s = String(v || '').trim();
  return /^\d{1,2}-[A-Za-z]{3}$/.test(s) ? s : null;
}

/**
 * Parses the live "M&H Travel Plan" sheet's Aug'26 tab. It's a manually-edited
 * sheet with merged header cells and inconsistent columns, so this is defensive:
 * it locates headers by text match rather than fixed column letters, and skips
 * (rather than guesses on) any row it can't confidently map to an Employee ID.
 */
function getTravelPlanData_() {
  const ss = SpreadsheetApp.openById(TRAVEL_SHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  const sh = findTravelSheet_(ss);
  if (!sh) return { roster: [], plans: {} };
  const values = sh.getDataRange().getValues();

  let headerRow = -1, nameCol = -1;
  for (let r = 0; r < values.length; r++) {
    const idIdx = values[r].indexOf('Employee ID');
    const nameIdx = values[r].indexOf('Full Name');
    if (idIdx > -1 && nameIdx > -1) { headerRow = r; nameCol = nameIdx; break; }
  }
  if (headerRow === -1) return { roster: [], plans: {} };

  const header = values[headerRow];
  const baseCol = header.indexOf('Base Location');
  const roleCol = header.indexOf('Role');
  const zoneCol = header.indexOf('Zone');
  const regionCol = header.indexOf('Region');
  const planCountCol = header.indexOf('Plan');

  // The "Plan" / "Actual" section markers (marking where each date-column
  // block starts) sit a row or two above the header, depending on how many
  // merged label rows precede it — search a small window rather than assume one.
  let planBlockStart = -1, actualBlockStart = -1;
  for (let back = 1; back <= 3 && (planBlockStart === -1 || actualBlockStart === -1); back++) {
    const row = values[headerRow - back] || [];
    if (planBlockStart === -1) planBlockStart = row.indexOf('Plan');
    if (actualBlockStart === -1) actualBlockStart = row.indexOf('Actual');
  }
  if (planBlockStart === -1) planBlockStart = planCountCol;
  if (actualBlockStart === -1) actualBlockStart = header.length;

  const dateColKeys = {};
  const dateColIdx = [];
  for (let c = 0; c < header.length; c++) {
    const key = dateCellToKey_(header[c], tz);
    if (key) { dateColIdx.push(c); dateColKeys[c] = key; }
  }
  const planDateCols = dateColIdx.filter(c => c >= planBlockStart && c < actualBlockStart);

  // Fallback name->id map from the small roster table elsewhere in the sheet,
  // for rows where the Employee ID column is missing/shifted.
  const nameToId = {};
  for (let r = 0; r < values.length; r++) {
    if (values[r][0] === 'Full Name' && values[r][1] === 'Employee Id') {
      for (let rr = r + 1; rr < values.length; rr++) {
        const nm = values[rr][0], id = values[rr][1];
        if (!nm) break;
        nameToId[normalizeName_(nm)] = String(id).trim();
      }
      break;
    }
  }

  const idPattern = /^\d{3,10}$/;
  const roster = [];
  const plans = {};
  const skipped = [];

  for (let r = headerRow + 1; r < values.length; r++) {
    const row = values[r];
    const name = String(row[nameCol] || '').trim();
    if (!name) continue;

    // Some rows have a stray extra ID-like value ahead of the real one (a data
    // entry artifact in the source sheet) — the value closest to the name
    // column is reliably the correct Employee ID, so keep overwriting rather
    // than stopping at the first match.
    let empId = '';
    for (let c = 0; c <= nameCol; c++) {
      const v = String(row[c] || '').trim();
      if (idPattern.test(v)) empId = v;
    }
    if (!empId) empId = nameToId[normalizeName_(name)] || '';
    if (!empId) { skipped.push(name); continue; }

    roster.push({
      id: empId,
      name: name,
      role: String(row[roleCol] || '').trim(),
      zone: String(row[zoneCol] || '').trim(),
      region: String(row[regionCol] || '').trim(),
      base: String(row[baseCol] || '').trim(),
      plan: Number(row[planCountCol]) || 0
    });

    const dayMap = {};
    planDateCols.forEach(c => {
      const val = String(row[c] || '').trim();
      if (val && !/^no travel$/i.test(val) && !/^leave$/i.test(val)) {
        dayMap[dateColKeys[c]] = val;
      }
    });
    plans[empId] = dayMap;
  }

  return { roster: roster, plans: plans, skipped: skipped };
}

// —————————————————————————————————————————————
// CITY GEOCODING — for the 30km location check on submit
// —————————————————————————————————————————————
const CITY_COORDS_TAB = 'City Coordinates';
const CITY_COORDS_HEADERS = ['city', 'lat', 'lng', 'geocoded_at'];
const MAX_GEOCODE_PER_REQUEST = 4; // keep doGet fast; the rest top up over the next few loads

function getCityCoordsSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(CITY_COORDS_TAB);
  if (!sh) sh = ss.insertSheet(CITY_COORDS_TAB);
  if (sh.getLastRow() === 0) sh.appendRow(CITY_COORDS_HEADERS);
  return sh;
}

function getCityCoordsCache_() {
  const sh = getCityCoordsSheet_();
  const rows = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const [city, lat, lng] = rows[i];
    if (city && lat && lng) map[city] = { lat: Number(lat), lng: Number(lng) };
  }
  return map;
}

/**
 * Uses Apps Script's built-in Maps geocoder (first-party, no external HTTP
 * call) as the primary lookup — Nominatim's abuse-prevention frequently
 * blocks Google's shared Apps Script IP ranges, so a raw UrlFetchApp call to
 * it fails silently in practice. Falls back to Nominatim only if Maps can't
 * resolve a name, since it's still worth a second try for odd local names.
 */
function geocodeCity_(cityName) {
  try {
    const geocoder = Maps.newGeocoder().setRegion('in');
    const response = geocoder.geocode(cityName + ', India');
    if (response.status === 'OK' && response.results && response.results.length) {
      const loc = response.results[0].geometry.location;
      return { lat: loc.lat, lng: loc.lng };
    }
    Logger.log('Maps geocoder returned status ' + response.status + ' for "' + cityName + '"');
  } catch (e) {
    Logger.log('Maps geocoder threw for "' + cityName + '": ' + e.toString());
  }

  try {
    const url = 'https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=in&q=' +
      encodeURIComponent(cityName + ', India');
    const res = UrlFetchApp.fetch(url, {
      headers: { 'User-Agent': 'InsuranceDekho-FieldVisitTracker/1.0 (internal tool)' },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      Logger.log('Nominatim HTTP ' + res.getResponseCode() + ' for "' + cityName + '": ' + res.getContentText().slice(0, 200));
      return null;
    }
    const arr = JSON.parse(res.getContentText());
    if (arr && arr.length) {
      return { lat: parseFloat(arr[0].lat), lng: parseFloat(arr[0].lon) };
    }
  } catch (e) {
    Logger.log('Nominatim fallback threw for "' + cityName + '": ' + e.toString());
  }
  return null;
}

/**
 * Geocodes any cities from the current travel plan that aren't cached yet,
 * capped per-request so doGet stays fast. Run backfillCityCoordinates()
 * manually (Apps Script editor -> Run) to catch up the rest in one go.
 */
function topUpCityCoordinates_(plans) {
  const cache = getCityCoordsCache_();
  const uniqueCities = new Set();
  Object.values(plans).forEach(dayMap => {
    Object.values(dayMap).forEach(city => uniqueCities.add(normalizeCity_(city)));
  });
  const missing = [...uniqueCities].filter(c => c && !cache[c]);
  if (!missing.length) return cache;

  const sh = getCityCoordsSheet_();
  let did = 0;
  for (const city of missing) {
    if (did >= MAX_GEOCODE_PER_REQUEST) break;
    const coords = geocodeCity_(city);
    if (coords) {
      sh.appendRow([city, coords.lat, coords.lng, new Date().toISOString()]);
      cache[city] = coords;
    }
    did++;
    if (did < missing.length) Utilities.sleep(1100); // respect Nominatim's 1 req/sec policy
  }
  return cache;
}

/**
 * One-time manual catch-up: run this from the Apps Script editor (select this
 * function, click Run) to geocode ALL missing cities in one go, instead of
 * waiting for them to top up a few at a time across normal app usage.
 */
function backfillCityCoordinates() {
  const travel = getTravelPlanData_();
  const cache = getCityCoordsCache_();
  const uniqueCities = new Set();
  Object.values(travel.plans).forEach(dayMap => {
    Object.values(dayMap).forEach(city => uniqueCities.add(normalizeCity_(city)));
  });
  const missing = [...uniqueCities].filter(c => c && !cache[c]);
  const sh = getCityCoordsSheet_();
  let done = 0;
  missing.forEach(city => {
    const coords = geocodeCity_(city);
    if (coords) {
      sh.appendRow([city, coords.lat, coords.lng, new Date().toISOString()]);
      done++;
    }
    Utilities.sleep(1100);
  });
  Logger.log('Geocoded ' + done + ' of ' + missing.length + ' missing cities.');
}

function doGet(e) {
  try {
    const key = (e.parameter && e.parameter.key) || '';
    if (key !== API_KEY) return json_({ status: 'error', msg: 'unauthorized' });

    const data = {};
    Object.keys(TABS).forEach(type => {
      const sh = getSheet_(type);
      const rows = sh.getDataRange().getValues();
      if (rows.length < 2) return;
      const headers = rows[0];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row[0]) continue; // skip blank rows
        const rec = { type: type };
        headers.forEach((h, idx) => { rec[h] = row[idx]; });
        data[rec.submission_id] = rec;
      }
    });
    let travel = { roster: [], plans: {}, skipped: [] };
    try {
      travel = getTravelPlanData_();
    } catch (travelErr) {
      travel.error = travelErr.toString();
    }

    let cityCoords = {};
    try {
      cityCoords = topUpCityCoordinates_(travel.plans || {});
    } catch (geoErr) {
      // don't let a geocoding hiccup take down the rest of the sync
    }

    return json_({
      status: 'success',
      data: data,
      logins: getLogins_(),
      roster: travel.roster,
      plans: travel.plans,
      travel_skipped: travel.skipped || [],
      travel_error: travel.error || '',
      city_coords: cityCoords
    });
  } catch (err) {
    return json_({ status: 'error', msg: err.toString() });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if ((data.key || '') !== API_KEY) return json_({ status: 'error', msg: 'unauthorized' });

    if (data.action === 'login') {
      return logLogin_(data);
    }

    // Manual/testing hook: trigger the day-wise travel summary on demand for
    // any specific day (data.day, e.g. '4-Aug') instead of waiting for the
    // 10am trigger. Omit data.day to send today's summary.
    // data.target: 'testing' posts to TESTING_WEBHOOK_URL instead of the
    // main CHAT_WEBHOOK_URL — used to verify a new report before it goes out
    // to the real space. Omit target to send to the main webhook.
    if (data.action === 'sendTravelSummary') {
      const payload = data.day
        ? buildTravelSummaryForDay_(data.day, data.day)
        : buildTodaysTravelSummary_();
      postToChat_(payload, data.target === 'testing' ? TESTING_WEBHOOK_URL : undefined);
      return json_({ status: 'success', preview: (payload.count || 0) + ' leader(s) traveling' });
    }

    if (data.action === 'sendFormFillSummary') {
      const payload = buildFormFillSummaryCard_();
      postToChat_(payload, data.target === 'testing' ? TESTING_WEBHOOK_URL : undefined);
      return json_({ status: 'success', preview: (payload.count || 0) + ' leader(s) with responses' });
    }

    // Installs (or re-installs) both daily triggers: travel summary at 8am
    // and form-fill summary at 8pm, both in the project's own time zone.
    if (data.action === 'setupTriggers') {
      setupDailyTravelTrigger();
      setupDailyFormFillTrigger();
      return json_({ status: 'success', msg: 'Daily triggers installed: travel summary ~8am, form-fill summary ~8pm.' });
    }

    // Diagnostic: returns the real distinct values (+counts) for the
    // categorical columns in the Responses sheet, so bucket categories can
    // be designed from actual data instead of guesses.
    if (data.action === 'inspectColumns') {
      return json_({ status: 'success', columns: inspectResponseColumns_() });
    }

    if (data.action === 'sendPartnerIntelSummary') {
      const payload = buildPartnerIntelSummaryCard_();
      postToChat_(payload, data.target === 'testing' ? TESTING_WEBHOOK_URL : undefined);
      return json_({ status: 'success', preview: (payload.count || 0) + ' leader(s) in report' });
    }

    if (data.action === 'sendDailyOpsTracker') {
      const payload = buildDailyOpsTrackerCard_();
      postToChat_(payload, data.target === 'testing' ? TESTING_WEBHOOK_URL : undefined);
      return json_({ status: 'success', preview: (payload.count || 0) + ' leader(s) in tracker' });
    }

    const type = data.visit_type;
    if (!TABS[type]) return json_({ status: 'error', msg: 'invalid visit_type' });
    const sh = getSheet_(type);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

    if (data.action === 'edit') {
      const idCol = headers.indexOf('submission_id') + 1;
      const lastRow = sh.getLastRow();
      if (lastRow < 2) return json_({ status: 'error', msg: 'not found' });
      const ids = sh.getRange(2, idCol, lastRow - 1, 1).getValues();
      let targetRow = -1;
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === data.submission_id) { targetRow = i + 2; break; }
      }
      if (targetRow === -1) return json_({ status: 'error', msg: 'not found' });
      headers.forEach((h, idx) => {
        if (data[h] !== undefined) sh.getRange(targetRow, idx + 1).setValue(data[h]);
      });
      return json_({ status: 'success' });
    }

    const row = headers.map(h => (data[h] !== undefined ? data[h] : ''));
    sh.appendRow(row);
    updateSummary_(data.emp_id, data.emp_name);
    return json_({ status: 'success' });
  } catch (err) {
    return json_({ status: 'error', msg: err.toString() });
  }
}

const SUMMARY_TAB = 'Leader Summary';
const SUMMARY_HEADERS = ['emp_id', 'emp_name', 'partner_meets', 'team_meets', 'insurer_meets', 'total_visits', 'last_updated'];

function getSummarySheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SUMMARY_TAB);
  if (!sh) sh = ss.insertSheet(SUMMARY_TAB);
  if (sh.getLastRow() === 0) sh.appendRow(SUMMARY_HEADERS);
  return sh;
}

function countByEmp_(type, empId) {
  const sh = getSheet_(type);
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return 0;
  const empCol = rows[0].indexOf('emp_id');
  let count = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][empCol]) === String(empId)) count++;
  }
  return count;
}

/**
 * Live tracker of how many partner/team/insurer visits each leader has
 * logged, kept in its own "Leader Summary" tab and refreshed on every new
 * submission (not needed for edits, since those don't change counts).
 */
function updateSummary_(empId, empName) {
  if (!empId) return;
  const sh = getSummarySheet_();
  const pmCount = countByEmp_('pm', empId);
  const tcCount = countByEmp_('tc', empId);
  const imCount = countByEmp_('im', empId);
  const total = pmCount + tcCount + imCount;
  const rows = sh.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(empId)) { targetRow = i + 1; break; }
  }
  const rowData = [empId, empName || '', pmCount, tcCount, imCount, total, new Date().toISOString()];
  if (targetRow === -1) {
    sh.appendRow(rowData);
  } else {
    sh.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  }
}

// —————————————————————————————————————————————
// DAILY TRAVEL SUMMARY — posted to a Google Chat space every morning
// —————————————————————————————————————————————
// Leadership order shown within each zone (highest first); anything not
// listed here sorts after these four.
const ROLE_ORDER_ = { ZH: 0, RH: 1, SH: 2, RM: 3 };
function roleRank_(role) {
  return ROLE_ORDER_.hasOwnProperty(role) ? ROLE_ORDER_[role] : 99;
}

// Hex colors mirror the web app's own role-badge palette, for visual consistency.
const ROLE_COLORS_ = { ZH: '#1a2b4a', RH: '#3b82f6', SH: '#10b981', RM: '#f59e0b' };
// A colored dot per zone, purely as a visual anchor in the section header.
const ZONE_DOTS_ = { North: '🔵', RON: '🟠', South: '🟢', West: '🟣', 'E&C': '🔴' };

/**
 * Builds one leader's row as a 2-column widget — [Role + Name] | City. Chat's
 * Columns widget only reliably keeps 2 columns side-by-side on a phone
 * screen; a 3rd column gets silently dropped rather than wrapping, which is
 * why Role/Name/City as 3 separate columns lost the City column on mobile.
 * Merging Role+Name into one column keeps City guaranteed-visible right
 * next to the name.
 */
function travelCardRow_(role, name, city) {
  const roleColor = ROLE_COLORS_[role] || '#4a5568';
  return {
    columns: {
      columnItems: [
        {
          horizontalSizeStyle: 'FILL_AVAILABLE_SPACE',
          widgets: [{ textParagraph: { text: '<font color="' + roleColor + '"><b>' + role + '</b></font>&nbsp;&nbsp;<b>' + name + '</b>' } }]
        },
        {
          horizontalSizeStyle: 'FILL_AVAILABLE_SPACE',
          widgets: [{ textParagraph: { text: '📍 ' + city } }]
        }
      ]
    }
  };
}

/**
 * Builds a colored Google Chat Card — grouped by zone, with each zone's
 * leaders ordered ZH -> RH -> SH -> RM, each row explicitly labeling its
 * travel destination. dayKey must match the travel plan's own column-key
 * format ('d-MMM', e.g. '4-Aug'); dateLabel is the human-readable heading.
 */
function buildTravelSummaryForDay_(dayKey, dateLabel) {
  const travel = getTravelPlanData_();

  const rows = [];
  travel.roster.forEach(l => {
    const city = (travel.plans[l.id] || {})[dayKey];
    if (city) rows.push({ zone: l.zone, name: l.name, role: l.role, city: city });
  });
  rows.sort((a, b) =>
    a.zone.localeCompare(b.zone) ||
    roleRank_(a.role) - roleRank_(b.role) ||
    a.name.localeCompare(b.name)
  );

  if (!rows.length) {
    return {
      count: 0,
      cardsV2: [{
        cardId: 'travelSummary-' + dayKey,
        card: {
          header: { title: 'Field Visit Tracker', subtitle: 'Travel Plan - ' + dateLabel, imageType: 'CIRCLE' },
          sections: [{ widgets: [{ textParagraph: { text: 'No leaders have a planned city visit on this day.' } }] }]
        }
      }]
    };
  }

  const sections = [];
  let currentZone = null, currentWidgets = null;
  rows.forEach(r => {
    if (r.zone !== currentZone) {
      currentZone = r.zone;
      currentWidgets = [travelCardRow_('ROLE', 'NAME', 'CITY'), { divider: {} }];
      sections.push({ header: (ZONE_DOTS_[r.zone] || '⚪') + ' ' + currentZone, widgets: currentWidgets });
    }
    currentWidgets.push(travelCardRow_(r.role, r.name, r.city));
  });
  sections.push({
    widgets: [{ textParagraph: { text: '<font color="#10b981"><b>Total traveling: ' + rows.length + '</b></font>' } }]
  });

  return {
    count: rows.length,
    cardsV2: [{
      cardId: 'travelSummary-' + dayKey,
      card: {
        header: { title: 'Field Visit Tracker', subtitle: 'Travel Plan - ' + dateLabel, imageType: 'CIRCLE' },
        sections: sections
      }
    }]
  };
}

function buildTodaysTravelSummary_() {
  const tz = SpreadsheetApp.openById(TRAVEL_SHEET_ID).getSpreadsheetTimeZone();
  const now = new Date();
  const todayKey = Utilities.formatDate(now, tz, 'd-MMM');
  const dateLabel = Utilities.formatDate(now, tz, 'EEEE, d MMMM yyyy');
  return buildTravelSummaryForDay_(todayKey, dateLabel);
}

function postToChat_(payload, webhookUrl) {
  const target = webhookUrl || CHAT_WEBHOOK_URL;
  if (!target || target.indexOf('PASTE_YOUR_') === 0) {
    Logger.log('Chat webhook not set — skipping Chat post.');
    return;
  }
  // Only forward fields the Chat webhook actually understands — "count" is
  // bookkeeping for our own doPost preview response, not part of the payload.
  const sendable = payload.cardsV2 ? { cardsV2: payload.cardsV2 } : { text: payload.text };
  const res = UrlFetchApp.fetch(target, {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify(sendable),
    muteHttpExceptions: true
  });
  Logger.log('Chat post response: ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 300));
}

// Column indexes in the "Responses" tab (0-based), per the M&H Visit Tracker
// backend's own header list.
const RESP_COL_TIMESTAMP_ = 1;
const RESP_COL_NAME_ = 2;
const RESP_COL_VISIT_DATE_ = 7;  // H
const RESP_COL_VISIT_TYPE_ = 8;
const RESP_COL_VISIT_CITY_ = 9;  // J
const RESP_COL_PARTNER_TYPE_ = 14;   // O: Existing Partner / New Partner
const RESP_COL_PARTNER_STATUS_ = 17; // R: Active / Inactive (NOT column T, which is Inactive Issues)
const RESP_COL_ACTIVE_ISSUES_ = 18;  // S
const RESP_COL_INACTIVE_ISSUES_ = 19; // T
const RESP_COL_ACTIVATION_PROB_ = 20; // U
const RESP_COL_BUSINESS_OPP_ = 21;    // V
const RESP_COL_SUPPORT_REQUIRED_ = 31; // AF
const RESP_COL_ACTION_OWNER_ = 33;     // AH
const RESP_COL_FOLLOWUP_REQUIRED_ = 34;

/**
 * Diagnostic only: reads a single named column (by its 0-based index) from
 * the Responses tab and tallies distinct values (splitting on '|' for
 * multi-select cells), without ever touching the Notes/Photo URLs columns —
 * those can carry embedded base64 images that make a full-row read huge.
 * Used to inspect real category values before building bucket logic, so
 * buckets are grounded in actual data instead of guessed.
 */
function tallyColumn_(colIndex) {
  const ss = SpreadsheetApp.openById(FORM_RESPONSES_SHEET_ID);
  const sh = ss.getSheetByName('Responses') || ss.getSheets()[0];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  const counts = {};
  values.forEach(row => {
    String(row[0] || '').split('|').map(s => s.trim()).filter(Boolean).forEach(part => {
      counts[part] = (counts[part] || 0) + 1;
    });
  });
  return Object.keys(counts).map(k => ({ value: k, count: counts[k] })).sort((a, b) => b.count - a.count);
}

function inspectResponseColumns_() {
  return {
    partnerType: tallyColumn_(RESP_COL_PARTNER_TYPE_),
    partnerStatus: tallyColumn_(RESP_COL_PARTNER_STATUS_),
    activeIssues: tallyColumn_(RESP_COL_ACTIVE_ISSUES_),
    inactiveIssues: tallyColumn_(RESP_COL_INACTIVE_ISSUES_),
    activationProbability: tallyColumn_(RESP_COL_ACTIVATION_PROB_),
    businessOpportunity: tallyColumn_(RESP_COL_BUSINESS_OPP_),
    supportRequired: tallyColumn_(RESP_COL_SUPPORT_REQUIRED_),
    actionOwner: tallyColumn_(RESP_COL_ACTION_OWNER_)
  };
}

// —————————————————————————————————————————————
// PARTNER INTELLIGENCE — leader-wise + company-wide partner analytics
// —————————————————————————————————————————————
// These mappings were built from the REAL distinct values found in the
// Responses sheet (via inspectResponseColumns_), not guessed. Active Issues
// only had 11 distinct values (not 62) as of this analysis — review/update
// this mapping if new issue strings appear in future responses that aren't
// covered here (they'll fall into "Other" by default, which is safe but
// worth periodically re-checking with inspectColumns).
const ACTIVE_ISSUE_BUCKETS_ = {
  'Payout': 'Payout & Pricing',
  'Pricing': 'Payout & Pricing',
  'Competition': 'Competition & Demand',
  'Low Customer Demand': 'Competition & Demand',
  'RM Support': 'Service & Claims',
  'Claims': 'Service & Claims',
  'Policy Issuance Delay': 'Service & Claims',
  'Product Availability': 'Product',
  'Product Knowledge': 'Product',
  'Technology': 'Technology'
  // Anything else (including the literal value "Other") falls into 'Other'.
};

const INACTIVE_ISSUE_BUCKETS_ = {
  'No Customer Demand': 'Demand & Seasonality',
  'Seasonal Business': 'Demand & Seasonality',
  'Shifted to Competitor': 'Lost to Competitor',
  'Pricing': 'Pricing & Support',
  'RM Support': 'Pricing & Support',
  'Technology': 'Pricing & Support',
  'Business Closed': 'Closed / Relationship',
  'Relationship Issue': 'Closed / Relationship'
};

function bucketOf_(map, rawValue) {
  return map[rawValue] || 'Other';
}

/**
 * Column U (Activation Possibility) turned out to be mostly free-text RM
 * notes ("Every month 50k business with us...") rather than a clean
 * High/Medium/Low dropdown — only a handful of rows are literal short
 * values. Rather than fabricate false precision by guessing sentiment from
 * prose, this only classifies exact literal matches and buckets everything
 * else as "Detailed note (unclassified)" — an honest signal that the form
 * doesn't yet collect this as structured data.
 */
function classifyActivationProbability_(raw) {
  const key = String(raw || '').trim().toLowerCase();
  if (!key) return 'Not specified';
  if (key === 'high' || key === 'yes') return 'High';
  if (key === 'medium' || key === 'slightly') return 'Medium';
  if (key === 'low' || key === 'no') return 'Low';
  return 'Detailed note (unclassified)';
}

function pct_(n, total) { return total ? Math.round((n / total) * 100) : 0; }

function topEntry_(obj) {
  const keys = Object.keys(obj || {});
  if (!keys.length) return null;
  let best = keys[0];
  keys.forEach(k => { if (obj[k] > obj[best]) best = k; });
  return { key: best, value: obj[best] };
}

/**
 * Reads every Partner Meet row (Team Connect/Insurer Meet rows don't have
 * partner-level fields, so they're excluded from partner analytics) and
 * builds both a per-leader breakdown and a company-wide rollup: partner
 * coverage (existing/new), active/inactive status, issue-category mix,
 * business opportunity mix, support-category mix, and which internal team
 * is most needed — everything computed dynamically from current data, not
 * hardcoded.
 */
function getPartnerIntelData_() {
  const travel = getTravelPlanData_();
  const roleByName = {}, zoneByName = {};
  travel.roster.forEach(l => {
    const key = normalizeName_(l.name);
    roleByName[key] = l.role;
    zoneByName[key] = l.zone;
  });

  const ss = SpreadsheetApp.openById(FORM_RESPONSES_SHEET_ID);
  const sh = ss.getSheetByName('Responses') || ss.getSheets()[0];
  const lastRow = sh.getLastRow();

  const leaders = {};
  const overall = {
    total: 0, existing: 0, newPartner: 0, active: 0, inactive: 0,
    issueBuckets: {}, inactiveIssueBuckets: {}, businessOpp: {}, supportBuckets: {},
    actionOwner: {}, activationProbBuckets: {}
  };

  function ensureLeader(key, name) {
    if (!leaders[key]) {
      leaders[key] = {
        name: name, role: roleByName[key] || '', zone: zoneByName[key] || 'Other',
        total: 0, existing: 0, newPartner: 0, active: 0, inactive: 0,
        issueBuckets: {}, businessOpp: {}, supportBuckets: {}, actionOwner: {}
      };
    }
    return leaders[key];
  }

  if (lastRow >= 2) {
    const n = lastRow - 1;
    const col = idx => sh.getRange(2, idx + 1, n, 1).getValues();
    const names = col(RESP_COL_NAME_);
    const visitTypes = col(RESP_COL_VISIT_TYPE_);
    const partnerTypes = col(RESP_COL_PARTNER_TYPE_);
    const partnerStatuses = col(RESP_COL_PARTNER_STATUS_);
    const activeIssues = col(RESP_COL_ACTIVE_ISSUES_);
    const inactiveIssues = col(RESP_COL_INACTIVE_ISSUES_);
    const activationProbs = col(RESP_COL_ACTIVATION_PROB_);
    const businessOpps = col(RESP_COL_BUSINESS_OPP_);
    const supportReqs = col(RESP_COL_SUPPORT_REQUIRED_);
    const actionOwners = col(RESP_COL_ACTION_OWNER_);

    for (let i = 0; i < n; i++) {
      if (String(visitTypes[i][0] || '').trim() !== 'Partner Meet') continue;
      const rawName = String(names[i][0] || '').trim();
      if (!rawName) continue;
      const key = normalizeName_(rawName);
      const leader = ensureLeader(key, rawName);

      leader.total++; overall.total++;

      const pType = String(partnerTypes[i][0] || '').trim();
      if (pType === 'Existing Partner') { leader.existing++; overall.existing++; }
      else if (pType === 'New Partner') { leader.newPartner++; overall.newPartner++; }

      const pStatus = String(partnerStatuses[i][0] || '').trim();
      const isInactive = pStatus === 'Inactive';
      if (pStatus === 'Active') { leader.active++; overall.active++; }
      else if (isInactive) { leader.inactive++; overall.inactive++; }

      String(activeIssues[i][0] || '').split('|').map(s => s.trim()).filter(Boolean).forEach(issue => {
        const bucket = bucketOf_(ACTIVE_ISSUE_BUCKETS_, issue);
        leader.issueBuckets[bucket] = (leader.issueBuckets[bucket] || 0) + 1;
        overall.issueBuckets[bucket] = (overall.issueBuckets[bucket] || 0) + 1;
      });

      const opp = String(businessOpps[i][0] || '').trim();
      if (opp) {
        leader.businessOpp[opp] = (leader.businessOpp[opp] || 0) + 1;
        overall.businessOpp[opp] = (overall.businessOpp[opp] || 0) + 1;
      }

      String(supportReqs[i][0] || '').split('|').map(s => s.trim()).filter(Boolean).forEach(sup => {
        leader.supportBuckets[sup] = (leader.supportBuckets[sup] || 0) + 1;
        overall.supportBuckets[sup] = (overall.supportBuckets[sup] || 0) + 1;
      });

      const owner = String(actionOwners[i][0] || '').trim();
      if (owner) {
        leader.actionOwner[owner] = (leader.actionOwner[owner] || 0) + 1;
        overall.actionOwner[owner] = (overall.actionOwner[owner] || 0) + 1;
      }

      if (isInactive) {
        String(inactiveIssues[i][0] || '').split('|').map(s => s.trim()).filter(Boolean).forEach(issue => {
          const bucket = bucketOf_(INACTIVE_ISSUE_BUCKETS_, issue);
          overall.inactiveIssueBuckets[bucket] = (overall.inactiveIssueBuckets[bucket] || 0) + 1;
        });
        const probBucket = classifyActivationProbability_(activationProbs[i][0]);
        overall.activationProbBuckets[probBucket] = (overall.activationProbBuckets[probBucket] || 0) + 1;
      }
    }
  }

  return { leaders: leaders, overall: overall };
}

/**
 * Builds the Partner Intelligence report as a Chat Card: a company-level
 * overview + dynamic key insights, followed by a condensed leader-wise
 * breakdown grouped by zone (ordered ZH -> RH -> SH -> RM). Per-leader
 * detail is intentionally condensed (coverage/status %, top issue, top
 * support need) rather than the full 6-section breakdown, to keep the card
 * a manageable size in Chat across many leaders — the same underlying data
 * (getPartnerIntelData_) has everything needed for a deeper cut if wanted.
 */
function buildPartnerIntelSummaryCard_() {
  const data = getPartnerIntelData_();
  const overall = data.overall;
  const sections = [];

  const topIssue = topEntry_(overall.issueBuckets);
  const topOpp = topEntry_(overall.businessOpp);
  const topSupport = topEntry_(overall.supportBuckets);
  const topTeam = topEntry_(overall.actionOwner);
  const highProb = overall.activationProbBuckets['High'] || 0;

  const overviewWidgets = [
    { textParagraph: { text: '<b>Company-Level Partner Coverage</b>' } },
    { textParagraph: { text: 'Partners Met: <b>' + overall.total + '</b>  |  Existing: <b>' + overall.existing + '</b> (' + pct_(overall.existing, overall.total) + '%)  New: <b>' + overall.newPartner + '</b> (' + pct_(overall.newPartner, overall.total) + '%)' } },
    { textParagraph: { text: 'Active: <font color="#10b981"><b>' + overall.active + '</b></font> (' + pct_(overall.active, overall.total) + '%)  Inactive: <font color="#ef4444"><b>' + overall.inactive + '</b></font> (' + pct_(overall.inactive, overall.total) + '%)' } },
    { textParagraph: { text: 'Top Problem: <font color="#ef4444"><b>' + (topIssue ? topIssue.key : '-') + '</b></font>' + (topIssue ? ' (' + topIssue.value + ' mentions)' : '') } },
    { textParagraph: { text: 'Top Business Opportunity: <font color="#10b981"><b>' + (topOpp ? topOpp.key : '-') + '</b></font>' + (topOpp ? ' (' + topOpp.value + ' partners)' : '') } },
    { textParagraph: { text: 'Top Support Need: <b>' + (topSupport ? topSupport.key : '-') + '</b>  |  Team Most Needed: <b>' + (topTeam ? topTeam.key : '-') + '</b>' } },
    { textParagraph: { text: '<font color="#4a5568">Reactivation signal (' + overall.inactive + ' inactive partners) — rough, based on RM notes not a dropdown yet: high-confidence mentions = <b>' + highProb + '</b></font>' } }
  ];
  sections.push({ header: 'Overview', widgets: overviewWidgets });

  const insightWidgets = [{ textParagraph: { text: '<b>Key Insights</b>' } }];
  if (overall.total) {
    insightWidgets.push({ textParagraph: { text: '1. ' + pct_(overall.existing, overall.total) + '% of partners met are existing partners, ' + pct_(overall.newPartner, overall.total) + '% are new.' } });
    if (topIssue) insightWidgets.push({ textParagraph: { text: '2. <b>' + topIssue.key + '</b> is the most frequent partner problem (' + topIssue.value + ' mentions).' } });
    if (topTeam) insightWidgets.push({ textParagraph: { text: '3. <b>' + topTeam.key + '</b> carries the highest support load (' + topTeam.value + ' requests).' } });
    insightWidgets.push({ textParagraph: { text: '4. ' + pct_(overall.inactive, overall.total) + '% of partners met are currently inactive — a reactivation target base of ' + overall.inactive + '.' } });
  } else {
    insightWidgets.push({ textParagraph: { text: 'No Partner Meet data found yet.' } });
  }
  sections.push({ widgets: insightWidgets });

  const leaders = Object.keys(data.leaders).map(k => data.leaders[k]).sort((a, b) =>
    a.zone.localeCompare(b.zone) || roleRank_(a.role) - roleRank_(b.role) || a.name.localeCompare(b.name)
  );

  let currentZone = null, currentWidgets = null;
  leaders.forEach(l => {
    if (l.zone !== currentZone) {
      currentZone = l.zone;
      currentWidgets = [];
      sections.push({ header: (ZONE_DOTS_[l.zone] || '⚪') + ' ' + currentZone, widgets: currentWidgets });
    }
    const roleColor = ROLE_COLORS_[l.role] || '#4a5568';
    const lTopIssue = topEntry_(l.issueBuckets);
    const lTopSupport = topEntry_(l.supportBuckets);
    currentWidgets.push({ textParagraph: { text: '<font color="' + roleColor + '"><b>' + (l.role || '-') + '</b></font> <b>' + l.name + '</b>' } });
    currentWidgets.push({ textParagraph: { text: 'Met: <b>' + l.total + '</b> | Existing ' + pct_(l.existing, l.total) + '% / New ' + pct_(l.newPartner, l.total) + '% | Active ' + pct_(l.active, l.total) + '% / Inactive ' + pct_(l.inactive, l.total) + '%' } });
    currentWidgets.push({ textParagraph: { text: 'Top issue: <b>' + (lTopIssue ? lTopIssue.key : '-') + '</b>  |  Top support need: <b>' + (lTopSupport ? lTopSupport.key : '-') + '</b>' } });
    currentWidgets.push({ divider: {} });
  });

  return {
    count: leaders.length,
    cardsV2: [{
      cardId: 'partnerIntel-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmm'),
      card: {
        header: { title: 'Field Visit Tracker', subtitle: 'Partner Intelligence Summary', imageType: 'CIRCLE' },
        sections: sections
      }
    }]
  };
}

// —————————————————————————————————————————————
// DAILY OPS TRACKER — a single condensed table: plan vs. actual, per leader
// —————————————————————————————————————————————
/**
 * Reads the Responses sheet once and returns, per leader (matched on
 * Employee Name): total forms filled (any visit type), Partner Meet count,
 * the set of distinct visit dates (= actual days travelled), and the set of
 * distinct visit cities (= actual cities covered) — all "actual" figures
 * for the Daily Ops Tracker come from here.
 */
function getActualTravelStats_() {
  const ss = SpreadsheetApp.openById(FORM_RESPONSES_SHEET_ID);
  const sh = ss.getSheetByName('Responses') || ss.getSheets()[0];
  const lastRow = sh.getLastRow();
  const stats = {};
  if (lastRow < 2) return stats;
  const n = lastRow - 1;
  const col = idx => sh.getRange(2, idx + 1, n, 1).getValues();
  const names = col(RESP_COL_NAME_);
  const visitTypes = col(RESP_COL_VISIT_TYPE_);
  const visitDates = col(RESP_COL_VISIT_DATE_);
  const visitCities = col(RESP_COL_VISIT_CITY_);
  const tz = Session.getScriptTimeZone();

  for (let i = 0; i < n; i++) {
    const raw = String(names[i][0] || '').trim();
    if (!raw) continue;
    const key = normalizeName_(raw);
    if (!stats[key]) stats[key] = { name: raw, totalForms: 0, partnerMeets: 0, dateSet: new Set(), citySet: new Set() };
    const rec = stats[key];
    rec.totalForms++;
    if (String(visitTypes[i][0] || '').trim() === 'Partner Meet') rec.partnerMeets++;

    const vd = visitDates[i][0];
    const dateKey = vd instanceof Date && !isNaN(vd.getTime()) ? Utilities.formatDate(vd, tz, 'yyyy-MM-dd') : String(vd || '').trim();
    if (dateKey) rec.dateSet.add(dateKey);

    const vc = String(visitCities[i][0] || '').trim();
    if (vc) rec.citySet.add(normalizeCity_(vc));
  }
  return stats;
}

function newZoneTotals_() {
  return { citiesPlanned: 0, daysPlanned: 0, daysActual: 0, forms: 0, partners: 0, citiesActual: 0 };
}
function addToTotals_(totals, r) {
  totals.citiesPlanned += r.citiesPlanned;
  totals.daysPlanned += r.daysPlanned;
  totals.daysActual += r.daysActual;
  totals.forms += r.forms;
  totals.partners += r.partners;
  totals.citiesActual += r.citiesActual;
}
function totalsRow_(label, t, w) {
  return padRight2_(label, w.name + w.today + 1) + ' ' +
    padRight2_(t.citiesPlanned, w.num) + ' ' + padRight2_(t.daysPlanned, w.num) + ' ' + padRight2_(t.daysActual, w.num) + ' ' +
    padRight2_(t.forms, w.num) + ' ' + padRight2_(t.partners, w.num) + ' ' + padRight2_(t.citiesActual, w.num);
}

function padRight2_(val, len) {
  let s = String(val);
  if (s.length > len) return s.slice(0, len - 1) + '…';
  while (s.length < len) s += ' ';
  return s;
}

/**
 * Builds the Daily Ops Tracker as a real monospace table — one header row,
 * one compact line per leader (no repeated labels per row), a PAN India
 * total at the top, and a zone total under each zone. Sent as plain text
 * (not a Card): Chat's Columns widget only reliably keeps 2 columns
 * side-by-side on a phone screen — a 3rd silently vanishes rather than
 * wrapping — so a genuine multi-column data table needs a monospace ```
 * block instead. Column headers are short but not cryptic (CITIES-P,
 * DAYS-P, etc.) so the table stays narrow enough to hold together.
 */
/**
 * Gathers the Daily Ops Tracker's rows (one per leader, plan-vs-actual) plus
 * the report date — the single source of truth consumed by both the Chat
 * table and the Dashboard sheet, so the two views never drift apart.
 */
function getDailyOpsTrackerRows_() {
  const travel = getTravelPlanData_();
  const tz = SpreadsheetApp.openById(TRAVEL_SHEET_ID).getSpreadsheetTimeZone();
  const now = new Date();
  const todayKey = Utilities.formatDate(now, tz, 'd-MMM');
  const todayDayOfMonth = now.getDate();
  const dateLabel = Utilities.formatDate(now, tz, 'EEEE, d MMMM yyyy');

  const actual = getActualTravelStats_();

  const rows = [];
  travel.roster.forEach(l => {
    const key = normalizeName_(l.name);
    const dayMap = travel.plans[l.id] || {};
    let daysPlanned = 0;
    const citiesPlannedSet = new Set();
    Object.keys(dayMap).forEach(dk => {
      const dayNum = parseInt(dk.split('-')[0], 10);
      if (!isNaN(dayNum) && dayNum <= todayDayOfMonth) {
        daysPlanned++;
        citiesPlannedSet.add(normalizeCity_(dayMap[dk]));
      }
    });
    const a = actual[key] || { totalForms: 0, partnerMeets: 0, dateSet: new Set(), citySet: new Set() };
    rows.push({
      zone: l.zone, role: l.role, name: l.name,
      today: dayMap[todayKey] || '-',
      citiesPlanned: citiesPlannedSet.size,
      daysPlanned: daysPlanned,
      daysActual: a.dateSet.size,
      forms: a.totalForms,
      partners: a.partnerMeets,
      citiesActual: a.citySet.size
    });
  });
  rows.sort((a, b) => a.zone.localeCompare(b.zone) || roleRank_(a.role) - roleRank_(b.role) || a.name.localeCompare(b.name));

  return { rows: rows, dateLabel: dateLabel };
}

/**
 * Writes the same Daily Ops Tracker data into a "Dashboard" tab in the
 * form-responses spreadsheet, as a real spreadsheet table — not
 * constrained by Chat's width/column limits. Freezes the Zone and Leader
 * Name columns so they stay visible while scrolling right through the
 * metric columns, including on the Sheets mobile app. Re-run each time the
 * tracker is generated, so it always reflects the latest data.
 */
function writeDailyOpsDashboard_(rows, dateLabel) {
  const ss = SpreadsheetApp.openById(FORM_RESPONSES_SHEET_ID);
  let sh = ss.getSheetByName('Dashboard');
  if (!sh) sh = ss.insertSheet('Dashboard');
  sh.clear();

  const headers = ['Zone', 'Leader Name', 'Role', "Today's Plan", 'Cities Planned (MTD)', 'Days Planned (MTD)', 'Days Actual', 'Forms Filled', 'Partners Met', 'Cities Covered'];
  const data = [];
  data.push(['Updated: ' + dateLabel]);
  data.push(headers);

  const panTotals = newZoneTotals_();
  rows.forEach(r => addToTotals_(panTotals, r));
  data.push(['', 'PAN INDIA TOTAL', '', '', panTotals.citiesPlanned, panTotals.daysPlanned, panTotals.daysActual, panTotals.forms, panTotals.partners, panTotals.citiesActual]);

  let currentZone = null, zoneTotals = null;
  rows.forEach(r => {
    if (r.zone !== currentZone) {
      if (currentZone !== null) {
        data.push(['', 'ZONE TOTAL - ' + currentZone, '', '', zoneTotals.citiesPlanned, zoneTotals.daysPlanned, zoneTotals.daysActual, zoneTotals.forms, zoneTotals.partners, zoneTotals.citiesActual]);
      }
      currentZone = r.zone;
      zoneTotals = newZoneTotals_();
    }
    data.push([r.zone, r.name, r.role, r.today, r.citiesPlanned, r.daysPlanned, r.daysActual, r.forms, r.partners, r.citiesActual]);
    addToTotals_(zoneTotals, r);
  });
  if (currentZone !== null) {
    data.push(['', 'ZONE TOTAL - ' + currentZone, '', '', zoneTotals.citiesPlanned, zoneTotals.daysPlanned, zoneTotals.daysActual, zoneTotals.forms, zoneTotals.partners, zoneTotals.citiesActual]);
  }

  const maxCols = headers.length;
  const paddedData = data.map(row => {
    const r = row.slice();
    while (r.length < maxCols) r.push('');
    return r;
  });
  sh.getRange(1, 1, paddedData.length, maxCols).setValues(paddedData);

  // Not merged across all columns — a merge spanning outside the frozen
  // column range throws "can't freeze columns which contain only part of a
  // merged cell" once setFrozenColumns(2) runs below.
  sh.getRange(1, 1, 1, 1).setFontWeight('bold').setFontColor('#4a5568');
  sh.getRange(2, 1, 1, maxCols).setFontWeight('bold').setBackground('#1a2b4a').setFontColor('#ffffff');
  sh.getRange(3, 1, 1, maxCols).setFontWeight('bold').setBackground('#d1fae5');

  for (let i = 3; i < paddedData.length; i++) {
    if (String(paddedData[i][1]).indexOf('ZONE TOTAL') === 0) {
      sh.getRange(i + 1, 1, 1, maxCols).setFontWeight('bold').setBackground('#e5e7eb');
    }
  }

  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);
  sh.autoResizeColumns(1, maxCols);
}

function buildDailyOpsTrackerCard_() {
  const data = getDailyOpsTrackerRows_();
  const rows = data.rows, dateLabel = data.dateLabel;
  writeDailyOpsDashboard_(rows, dateLabel);

  const w = { name: 13, today: 8, num: 5 };
  const headerLine = padRight2_('NAME', w.name) + ' ' + padRight2_('TODAY', w.today) + ' ' +
    padRight2_('CITYP', w.num) + ' ' + padRight2_('DAYSP', w.num) + ' ' + padRight2_('DAYSA', w.num) + ' ' +
    padRight2_('FORMS', w.num) + ' ' + padRight2_('PTNRS', w.num) + ' ' + padRight2_('CITYA', w.num);
  const totalW = headerLine.length;

  const panIndiaTotals = newZoneTotals_();
  rows.forEach(r => addToTotals_(panIndiaTotals, r));

  const lines = [];
  lines.push('PAN INDIA TOTAL');
  lines.push(totalsRow_('All Zones', panIndiaTotals, w));
  lines.push('');

  let currentZone = null, currentZoneTotals = null;
  rows.forEach(r => {
    if (r.zone !== currentZone) {
      if (currentZone !== null) {
        lines.push('-'.repeat(totalW));
        lines.push(totalsRow_('Zone Total', currentZoneTotals, w));
        lines.push('');
      }
      currentZone = r.zone;
      currentZoneTotals = newZoneTotals_();
      lines.push('[' + currentZone + ']');
      lines.push(headerLine);
      lines.push('-'.repeat(totalW));
    }
    lines.push(
      padRight2_(r.name, w.name) + ' ' + padRight2_(r.today, w.today) + ' ' +
      padRight2_(r.citiesPlanned, w.num) + ' ' + padRight2_(r.daysPlanned, w.num) + ' ' + padRight2_(r.daysActual, w.num) + ' ' +
      padRight2_(r.forms, w.num) + ' ' + padRight2_(r.partners, w.num) + ' ' + padRight2_(r.citiesActual, w.num)
    );
    addToTotals_(currentZoneTotals, r);
  });
  if (currentZone !== null) {
    lines.push('-'.repeat(totalW));
    lines.push(totalsRow_('Zone Total', currentZoneTotals, w));
  }

  const legend = 'CITYP=Cities Planned MTD  DAYSP=Days Planned MTD  DAYSA=Days Actually Travelled  FORMS=Forms Filled  PTNRS=Partners Met  CITYA=Cities Covered (Actual)';
  const text = '*Field Visit Tracker - Daily Ops Tracker*\nPublished: ' + dateLabel + '\n```\n' + lines.join('\n') + '\n```\n' + legend;
  return { count: rows.length, text: text };
}

/**
 * Reads every row of the form-responses spreadsheet's "Responses" tab and
 * aggregates, per leader (matched on Employee Name): total forms filled, a
 * breakdown by visit type (Partner/Team/Insurer), how many still need
 * follow-up, and the most recent submission timestamp — so the daily report
 * can show more than a bare count. Row 1 is assumed to be a header row.
 */
function getFormFillCounts_() {
  const ss = SpreadsheetApp.openById(FORM_RESPONSES_SHEET_ID);
  const sh = ss.getSheetByName('Responses') || ss.getSheets()[0];
  const values = sh.getDataRange().getValues();
  const counts = {};
  for (let r = 1; r < values.length; r++) {
    const raw = String(values[r][RESP_COL_NAME_] || '').trim();
    if (!raw) continue;
    const key = normalizeName_(raw);
    if (!counts[key]) {
      counts[key] = {
        name: raw, count: 0,
        byType: { 'Partner Meet': 0, 'Team Connect': 0, 'Insurer Meet': 0 },
        pendingFollowUps: 0,
        lastVisitAt: null
      };
    }
    const rec = counts[key];
    rec.count++;

    const visitType = String(values[r][RESP_COL_VISIT_TYPE_] || '').trim();
    if (rec.byType.hasOwnProperty(visitType)) rec.byType[visitType]++;

    if (String(values[r][RESP_COL_FOLLOWUP_REQUIRED_] || '').trim() === 'Yes') rec.pendingFollowUps++;

    const ts = values[r][RESP_COL_TIMESTAMP_];
    const tsDate = ts instanceof Date ? ts : new Date(ts);
    if (!isNaN(tsDate.getTime()) && (!rec.lastVisitAt || tsDate > rec.lastVisitAt)) rec.lastVisitAt = tsDate;
  }
  return counts;
}

/** "Today" / "Yesterday" / "N days ago" / "-" for a leader's most recent submission. */
function daysAgoLabel_(date) {
  if (!date) return '-';
  const msPerDay = 24 * 60 * 60 * 1000;
  const days = Math.floor((new Date().setHours(0, 0, 0, 0) - new Date(date).setHours(0, 0, 0, 0)) / msPerDay);
  if (days <= 0) return 'Today';
  if (days === 1) return 'Yesterday';
  return days + ' days ago';
}

// Compliance/recency color thresholds — green on track, amber behind, red well behind.
function complianceColor_(pct) {
  return pct >= 80 ? '#10b981' : (pct >= 50 ? '#f59e0b' : '#ef4444');
}
function recencyColor_(lastVisitAt) {
  if (!lastVisitAt) return '#ef4444';
  const days = Math.floor((new Date().setHours(0, 0, 0, 0) - new Date(lastVisitAt).setHours(0, 0, 0, 0)) / 86400000);
  return days <= 2 ? '#10b981' : (days <= 6 ? '#f59e0b' : '#ef4444');
}

/**
 * Builds one leader's row as a 2-column widget: Role + Name stacked in
 * column 1; in column 2 — filled count + compliance (vs. this leader's
 * planned visits from the roster), a Partner/Team/Insurer breakdown,
 * pending follow-ups, and how recently they last submitted — so the report
 * shows workload mix and where the backlog/staleness actually is, not just
 * a bare count.
 */
function formFillRow_(role, name, filled, planned, byType, pendingFollowUps, lastVisitAt) {
  const roleColor = ROLE_COLORS_[role] || '#4a5568';
  const col2Widgets = [{ textParagraph: { text: 'Filled: <b>' + filled + '</b>' } }];
  if (planned) {
    const pct = Math.round((filled / planned) * 100);
    col2Widgets.push({
      textParagraph: {
        text: '<font color="' + complianceColor_(pct) + '">Compliance: <b>' + filled + '/' + planned + ' (' + pct + '%)</b></font>'
      }
    });
  } else {
    col2Widgets.push({ textParagraph: { text: '<font color="#4a5568">Planned count unavailable</font>' } });
  }
  if (byType) {
    col2Widgets.push({
      textParagraph: {
        text: 'Partner: <b>' + byType['Partner Meet'] + '</b> | Team: <b>' + byType['Team Connect'] + '</b> | Insurer: <b>' + byType['Insurer Meet'] + '</b>'
      }
    });
  }
  if (pendingFollowUps !== undefined) {
    const fColor = pendingFollowUps > 0 ? '#ef4444' : '#10b981';
    col2Widgets.push({
      textParagraph: { text: '<font color="' + fColor + '">Pending follow-ups: <b>' + pendingFollowUps + '</b></font>' }
    });
  }
  if (lastVisitAt !== undefined) {
    col2Widgets.push({
      textParagraph: { text: '<font color="' + recencyColor_(lastVisitAt) + '">Last active: <b>' + daysAgoLabel_(lastVisitAt) + '</b></font>' }
    });
  }
  return {
    columns: {
      columnItems: [
        {
          horizontalSizeStyle: 'FILL_AVAILABLE_SPACE',
          widgets: [
            { textParagraph: { text: '<font color="' + roleColor + '"><b>' + (role || '-') + '</b></font>' } },
            { textParagraph: { text: '<b>' + name + '</b>' } }
          ]
        },
        { horizontalSizeStyle: 'FILL_AVAILABLE_SPACE', widgets: col2Widgets }
      ]
    }
  };
}

/** Header row for the form-fill table — same 2-column shape as the data rows. */
function formFillHeaderRow_() {
  return {
    columns: {
      columnItems: [
        { horizontalSizeStyle: 'FILL_AVAILABLE_SPACE', widgets: [{ textParagraph: { text: '<b>ROLE / NAME</b>' } }] },
        { horizontalSizeStyle: 'FILL_AVAILABLE_SPACE', widgets: [{ textParagraph: { text: '<b>Activity Summary</b>' } }] }
      ]
    }
  };
}

/**
 * Builds a colored Card — same layout as the travel summary — showing every
 * leader who appears in the form-responses sheet, how many forms they've
 * filled so far, and their compliance rate against this month's planned
 * visits (from the live roster), so the report highlights who's on track
 * and who's falling behind instead of just a raw count. Grouped by zone and
 * ordered ZH -> RH -> SH -> RM within each zone; any name not found on the
 * roster is grouped under "Other" with compliance shown as unavailable.
 */
function buildFormFillSummaryCard_() {
  const travel = getTravelPlanData_();
  const roleByName = {}, zoneByName = {}, plannedByName = {};
  travel.roster.forEach(l => {
    const key = normalizeName_(l.name);
    roleByName[key] = l.role;
    zoneByName[key] = l.zone;
    plannedByName[key] = l.plan || 0;
  });

  const counts = getFormFillCounts_();
  const rows = Object.keys(counts).map(key => ({
    name: counts[key].name,
    count: counts[key].count,
    byType: counts[key].byType,
    pendingFollowUps: counts[key].pendingFollowUps,
    lastVisitAt: counts[key].lastVisitAt,
    planned: plannedByName[key] || 0,
    role: roleByName[key] || '',
    zone: zoneByName[key] || 'Other'
  }));
  rows.sort((a, b) =>
    a.zone.localeCompare(b.zone) ||
    roleRank_(a.role) - roleRank_(b.role) ||
    a.name.localeCompare(b.name)
  );

  if (!rows.length) {
    return {
      count: 0,
      cardsV2: [{
        cardId: 'formFillSummary-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmm'),
        card: {
          header: { title: 'Field Visit Tracker', subtitle: 'Form Fill Summary', imageType: 'CIRCLE' },
          sections: [{ widgets: [{ textParagraph: { text: 'No form responses found yet.' } }] }]
        }
      }]
    };
  }

  const sections = [];
  let currentZone = null, currentWidgets = null;
  let totalResponses = 0;
  rows.forEach(r => {
    if (r.zone !== currentZone) {
      currentZone = r.zone;
      currentWidgets = [formFillHeaderRow_(), { divider: {} }];
      sections.push({ header: (ZONE_DOTS_[r.zone] || '⚪') + ' ' + currentZone, widgets: currentWidgets });
    }
    currentWidgets.push(formFillRow_(r.role, r.name, r.count, r.planned, r.byType, r.pendingFollowUps, r.lastVisitAt));
    currentWidgets.push({ divider: {} });
    totalResponses += r.count;
  });

  // Key Insight section: surfaces who's behind, who's leading, and where the
  // follow-up backlog actually sits, instead of leaving the reader to scan
  // every row themselves for the same signal.
  const withPlan = rows.filter(r => r.planned > 0).map(r => ({ r: r, pct: Math.round((r.count / r.planned) * 100) }));
  const behind = withPlan.filter(x => x.pct < 50);
  const totalPendingFollowUps = rows.reduce((sum, r) => sum + (r.pendingFollowUps || 0), 0);
  const insightWidgets = [{ textParagraph: { text: '<b>Key Insight</b>' } }];
  if (withPlan.length) {
    const top = withPlan.reduce((best, x) => (x.pct > best.pct ? x : best), withPlan[0]);
    insightWidgets.push({
      textParagraph: { text: '<font color="#10b981">Top performer: <b>' + top.r.name + '</b> (' + top.pct + '% compliance)</font>' }
    });
    if (behind.length) {
      const names = behind.map(x => x.r.name + ' (' + x.pct + '%)').join(', ');
      insightWidgets.push({
        textParagraph: { text: '<font color="#ef4444"><b>' + behind.length + ' leader(s) below 50% compliance:</b> ' + names + '</font>' }
      });
    } else {
      insightWidgets.push({ textParagraph: { text: '<font color="#10b981">No leaders below 50% compliance.</font>' } });
    }
    if (totalPendingFollowUps > 0) {
      insightWidgets.push({
        textParagraph: { text: '<font color="#f59e0b"><b>' + totalPendingFollowUps + ' follow-up(s)</b> pending across the team.</font>' }
      });
    }
  } else {
    insightWidgets.push({ textParagraph: { text: 'Planned-visit data unavailable for matching leaders.' } });
  }
  sections.push({ widgets: insightWidgets });

  sections.push({
    widgets: [{ textParagraph: { text: '<font color="#10b981"><b>Total leaders: ' + rows.length + ' | Total responses: ' + totalResponses + '</b></font>' } }]
  });

  return {
    count: rows.length,
    cardsV2: [{
      cardId: 'formFillSummary-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmm'),
      card: {
        header: { title: 'Field Visit Tracker', subtitle: 'Form Fill Summary', imageType: 'CIRCLE' },
        sections: sections
      }
    }]
  };
}

/**
 * The daily travel-plan job — install this on a trigger (see
 * setupDailyTravelTrigger) so it runs automatically, or run it manually any
 * time to send an on-demand summary of who's traveling where today. Posts to
 * the main CHAT_WEBHOOK_URL (Agency Warriors).
 */
function sendDailyTravelSummary() {
  postToChat_(buildTodaysTravelSummary_());
}

/**
 * The daily 8pm job — install this on a trigger (see
 * setupDailyFormFillTrigger) so it runs automatically. Posts to the main
 * CHAT_WEBHOOK_URL (Agency Warriors). Sends the Daily Ops Tracker (plan vs.
 * actual, per leader) — this replaced the older Form Fill Summary /
 * Partner Intelligence report as the 8pm send; those report-builders are
 * still available for on-demand testing via their own doPost actions.
 */
function sendDailyFormFillSummary() {
  postToChat_(buildDailyOpsTrackerCard_());
}

/**
 * One-time setup: run this once (from the Apps Script editor, or via the
 * doPost 'setupTriggers' action) to install the daily 8am travel-plan
 * trigger. Re-running it is safe — it clears any previous trigger for this
 * function first so you never end up with duplicates. Apps Script
 * day-timer triggers fire sometime within the chosen hour, not at an exact
 * minute — and use the PROJECT's time zone (Project Settings -> General ->
 * Time zone), so make sure that's set to India Standard Time (Asia/Kolkata)
 * for this to actually mean 8am IST.
 */
function setupDailyTravelTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyTravelSummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyTravelSummary')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
  Logger.log('Daily trigger installed for sendDailyTravelSummary (~8am, project time zone).');
}

/**
 * One-time setup: run this once (from the Apps Script editor, or via the
 * doPost 'setupTriggers' action) to install the daily 8pm form-fill-summary
 * trigger. Re-running it is safe — it clears any previous trigger for this
 * function first. Same project-time-zone caveat as setupDailyTravelTrigger.
 */
function setupDailyFormFillTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyFormFillSummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyFormFillSummary')
    .timeBased()
    .atHour(20)
    .everyDays(1)
    .create();
  Logger.log('Daily trigger installed for sendDailyFormFillSummary (~8pm, project time zone).');
}
