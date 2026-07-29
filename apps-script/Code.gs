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

// Exact field names written/read for each visit type. Order defines column order.
const HEADERS = {
  pm: ['submission_id','ts','emp_id','emp_name','designation','zone','region','planned_date','city',
       'partner_gid','partner_name','partner_contact','exist_new','bucket','status','purpose',
       'purpose_other','active_challenges','challenge_other','trend','confidence','inactive_reason',
       'lob','support','outcome','notes','edited_ts','edited_by'],
  tc: ['submission_id','ts','emp_id','emp_name','designation','zone','region','planned_date','city',
       'mode','branch','bcity','meet_type','objective','team_size','agenda','actions','health',
       'challenge','challenge_other','edited_ts','edited_by'],
  im: ['submission_id','ts','emp_id','emp_name','designation','zone','region','planned_date','city',
       'insurer','purpose','discussion_areas','outcomes_text','outcome','edited_ts','edited_by']
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
    return json_({ status: 'success', data: data, logins: getLogins_() });
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
    return json_({ status: 'success' });
  } catch (err) {
    return json_({ status: 'error', msg: err.toString() });
  }
}
