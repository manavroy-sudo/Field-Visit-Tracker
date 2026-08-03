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

// Google Chat space webhook — posts the daily travel summary. Get this from
// the space: name/gear icon -> Apps & integrations -> Webhooks -> Add webhook.
const CHAT_WEBHOOK_URL = 'PASTE_YOUR_CHAT_WEBHOOK_URL_HERE';

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
function padRight_(str, len) {
  str = String(str);
  while (str.length < len) str += ' ';
  return str;
}

function truncate_(str, max) {
  str = String(str);
  return str.length > max ? str.slice(0, max - 1) + '…' : str;
}

// Fixed, capped column widths (not data-driven) so the total line length is
// bounded no matter how long a name or city string gets — this is what
// keeps the table from wrapping/misaligning on a narrow phone screen. Role
// codes (ZH/RH/SH/RM) are always 2 chars so that column stays tiny.
var NAME_W_ = 18, ROLE_W_ = 4, CITY_W_ = 16;

/**
 * Builds a plain-text, monospace ASCII table inside a ``` code block —
 * grouped by zone via an in-table sub-header line rather than a 5th column,
 * to keep every row within the fixed width above.
 */
function buildTodaysTravelSummary_() {
  const travel = getTravelPlanData_();
  const tz = SpreadsheetApp.openById(TRAVEL_SHEET_ID).getSpreadsheetTimeZone();
  const now = new Date();
  const todayKey = Utilities.formatDate(now, tz, 'd-MMM');

  const rows = [];
  travel.roster.forEach(l => {
    const city = (travel.plans[l.id] || {})[todayKey];
    if (city) rows.push({ zone: l.zone, name: l.name, role: l.role, city: city });
  });
  rows.sort((a, b) => a.zone.localeCompare(b.zone) || a.name.localeCompare(b.name));

  const dateLabel = Utilities.formatDate(now, tz, 'EEEE, d MMMM yyyy');

  if (!rows.length) {
    return { text: '*Field Visit Tracker - Today\'s Travel Plan*\n' + dateLabel + '\n\nNo leaders have a planned city visit today.' };
  }

  const totalW = NAME_W_ + ROLE_W_ + CITY_W_ + 2;
  const lines = [];
  lines.push(padRight_('NAME', NAME_W_) + ' ' + padRight_('ROLE', ROLE_W_) + ' ' + 'CITY');
  lines.push('-'.repeat(totalW));

  let currentZone = null;
  rows.forEach(r => {
    if (r.zone !== currentZone) {
      currentZone = r.zone;
      if (lines.length > 2) lines.push('');
      lines.push('[' + currentZone + ']');
    }
    lines.push(
      padRight_(truncate_(r.name, NAME_W_), NAME_W_) + ' ' +
      padRight_(r.role, ROLE_W_) + ' ' +
      truncate_(r.city, CITY_W_)
    );
  });

  const text = '*Field Visit Tracker - Today\'s Travel Plan*\n' + dateLabel + '\n```\n' + lines.join('\n') + '\n```\nTotal traveling today: ' + rows.length;
  return { text: text };
}

function postToChat_(payload) {
  if (!CHAT_WEBHOOK_URL || CHAT_WEBHOOK_URL === 'PASTE_YOUR_CHAT_WEBHOOK_URL_HERE') {
    Logger.log('CHAT_WEBHOOK_URL not set — skipping Chat post.');
    return;
  }
  const res = UrlFetchApp.fetch(CHAT_WEBHOOK_URL, {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log('Chat post response: ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 300));
}

/**
 * The daily job — install this on a trigger (see setupDailyTravelTrigger)
 * so it runs automatically, or run it manually any time to send an
 * on-demand summary of who's traveling where today.
 */
function sendDailyTravelSummary() {
  postToChat_(buildTodaysTravelSummary_());
}

/**
 * One-time setup: run this once from the Apps Script editor (select this
 * function, click Run) to install the daily 10am trigger. Re-running it is
 * safe — it clears any previous trigger for this function first so you never
 * end up with duplicates. Apps Script day-timer triggers fire sometime within
 * the chosen hour, not at an exact minute — and use the PROJECT's time zone
 * (Project Settings -> General -> Time zone), so make sure that's set to
 * India Standard Time (Asia/Kolkata) for this to actually mean 10am IST.
 */
function setupDailyTravelTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyTravelSummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyTravelSummary')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .create();
  Logger.log('Daily trigger installed for sendDailyTravelSummary (~10am, project time zone).');
}
