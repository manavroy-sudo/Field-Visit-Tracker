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

/**
 * Reads Column C of the form-responses spreadsheet (one row per submission)
 * and counts how many times each leader's name appears — that count is how
 * many forms they've filled so far. Row 1 is assumed to be a header row.
 */
function getFormFillCounts_() {
  const ss = SpreadsheetApp.openById(FORM_RESPONSES_SHEET_ID);
  const sh = ss.getSheets()[0];
  const values = sh.getDataRange().getValues();
  const counts = {};
  for (let r = 1; r < values.length; r++) {
    const raw = String(values[r][2] || '').trim(); // Column C
    if (!raw) continue;
    const key = normalizeName_(raw);
    if (!counts[key]) counts[key] = { name: raw, count: 0 };
    counts[key].count++;
  }
  return counts;
}

/**
 * Builds one leader's row as a 2-column widget — [Role + Name] | Count —
 * matching the same layout used for the travel summary.
 */
function formFillRow_(role, name, count) {
  const roleColor = ROLE_COLORS_[role] || '#4a5568';
  return {
    columns: {
      columnItems: [
        {
          horizontalSizeStyle: 'FILL_AVAILABLE_SPACE',
          widgets: [{ textParagraph: { text: '<font color="' + roleColor + '"><b>' + (role || '-') + '</b></font>&nbsp;&nbsp;<b>' + name + '</b>' } }]
        },
        {
          horizontalSizeStyle: 'FILL_AVAILABLE_SPACE',
          widgets: [{ textParagraph: { text: '📝 ' + count } }]
        }
      ]
    }
  };
}

/**
 * Builds a colored Card — same layout as the travel summary — showing every
 * leader who appears in the form-responses sheet and how many forms they've
 * filled so far in total. Grouped by zone (from the live roster) and ordered
 * ZH -> RH -> SH -> RM within each zone; any name not found on the roster is
 * grouped under "Other".
 */
function buildFormFillSummaryCard_() {
  const travel = getTravelPlanData_();
  const roleByName = {}, zoneByName = {};
  travel.roster.forEach(l => {
    const key = normalizeName_(l.name);
    roleByName[key] = l.role;
    zoneByName[key] = l.zone;
  });

  const counts = getFormFillCounts_();
  const rows = Object.keys(counts).map(key => ({
    name: counts[key].name,
    count: counts[key].count,
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
      currentWidgets = [formFillRow_('ROLE', 'NAME', 'COUNT'), { divider: {} }];
      sections.push({ header: (ZONE_DOTS_[r.zone] || '⚪') + ' ' + currentZone, widgets: currentWidgets });
    }
    currentWidgets.push(formFillRow_(r.role, r.name, r.count));
    totalResponses += r.count;
  });
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
 * The daily form-fill job — install this on a trigger (see
 * setupDailyFormFillTrigger) so it runs automatically. Posts to the main
 * CHAT_WEBHOOK_URL (Agency Warriors).
 */
function sendDailyFormFillSummary() {
  postToChat_(buildFormFillSummaryCard_());
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
