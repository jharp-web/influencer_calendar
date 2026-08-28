// =====================================================================
// MARKETING CALENDAR — BACKEND
// Main "Posts" sheet is the single source of truth.
// Run runImport() ONCE to pull the two external calendars in.
// =====================================================================

var CALENDAR_YEAR = 2026;

// External source sheets (used only by runImport)
var ORGANIC_SHEET_ID = '1pr5uNAH5VwgHXH5LFVDRwqUfhaQjRRYim_bY_zj9-w4';
var ORGANIC_TAB_NAME = 'Content calendar';
var PAID_SHEET_ID    = '1kZd027NQ1wd3c9V73C9fJZZdKT9PMGjyDitJuT9uXKU';
var PAID_TAB_GID     = 1661186045;

var SCHEMA = ['id','team','date','endDate','channel','platforms','postType',
              'influencer','title','notes','status','funnel','format',
              'audience','evergreen','source'];

// =====================================================================
// WEB APP
// =====================================================================
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({posts: getAllPosts()}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'save')   return respond(upsertPost(body.post));
    if (body.action === 'delete') return respond(deletePost(body.id));
    return respond({error: 'Unknown action'});
  } catch(err) {
    return respond({error: err.message});
  } finally {
    lock.releaseLock();
  }
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================================
// MAIN SHEET
// =====================================================================
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Posts');
  if (!sheet) {
    sheet = ss.insertSheet('Posts');
    sheet.appendRow(SCHEMA);
    return sheet;
  }
  ensureSchema(sheet);
  return sheet;
}

// Adds any missing columns from SCHEMA to the header row
function ensureSchema(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                  .map(function(h) { return String(h).trim(); });
  var added = false;
  SCHEMA.forEach(function(h) {
    if (headers.indexOf(h) < 0) {
      headers.push(h);
      sheet.getRange(1, headers.length).setValue(h);
      added = true;
    }
  });
  return added;
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h).trim(); });
}

function getAllPosts() {
  var sheet = getSheet();
  if (sheet.getLastRow() < 2) return [];
  var headers = getHeaders(sheet);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  var posts = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var post = {};
    headers.forEach(function(h, idx) { post[h] = row[idx]; });
    if (!post.id) continue;

    ['platforms','audience'].forEach(function(key) {
      var v = post[key];
      if (Array.isArray(v)) return;
      if (typeof v === 'string' && v) {
        try { post[key] = JSON.parse(v); }
        catch(e) { post[key] = [v]; }
      } else {
        post[key] = [];
      }
    });

    post.date      = toDateStr(post.date);
    post.endDate   = post.endDate ? toDateStr(post.endDate) : '';
    post.evergreen = post.evergreen === true || post.evergreen === 'TRUE' || post.evergreen === 'true';
    post.team      = normalizeTeam(post.team, post.channel);

    if (post.date) posts.push(post);
  }
  return posts;
}

function upsertPost(post) {
  var sheet   = getSheet();
  var headers = getHeaders(sheet);
  function ci(key) { return headers.indexOf(key); }

  var values = {};
  SCHEMA.forEach(function(k) {
    if (k === 'platforms' || k === 'audience') {
      values[k] = JSON.stringify(Array.isArray(post[k]) ? post[k] : []);
    } else if (k === 'evergreen') {
      values[k] = post[k] === true;
    } else {
      values[k] = post[k] == null ? '' : post[k];
    }
  });

  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, ci('id') + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(post.id)) {
        var range = sheet.getRange(i + 2, 1, 1, headers.length);
        var rowVals = range.getValues()[0];
        SCHEMA.forEach(function(k) {
          var idx = ci(k);
          if (idx >= 0) rowVals[idx] = values[k];
        });
        range.setValues([rowVals]);
        return {success: true, updated: true};
      }
    }
  }

  var newRow = new Array(headers.length).fill('');
  SCHEMA.forEach(function(k) {
    var idx = ci(k);
    if (idx >= 0) newRow[idx] = values[k];
  });
  sheet.appendRow(newRow);
  return {success: true, created: true};
}

function deletePost(id) {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {error: 'Not found'};
  var headers = getHeaders(sheet);
  var idCol = headers.indexOf('id') + 1;
  var ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      return {success: true};
    }
  }
  return {error: 'Not found'};
}

// =====================================================================
// DATE PARSING
// Handles: real Dates, "1/9", "Jan 7", "January 5 - 9", typos, blanks
// =====================================================================
var MONTH_MAP = {
  jan:1, january:1, jaunuary:1, janurary:1,
  feb:2, february:2, februray:2, febuary:2,
  mar:3, march:3,
  apr:4, april:4,
  may:5,
  jun:6, june:6,
  jul:7, july:7,
  aug:8, august:8,
  sep:9, sept:9, september:9,
  oct:10, october:10,
  nov:11, november:11,
  dec:12, december:12
};

function pad2(n) { return String(n).padStart(2, '0'); }

function ymd(y, m, d) {
  return y + '-' + pad2(m) + '-' + pad2(d);
}

// Normalize an already-good value (Date object or ISO string) to YYYY-MM-DD
function toDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return ymd(val.getFullYear(), val.getMonth() + 1, val.getDate());
  }
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return parseLooseDate(s).start;
}

// Returns {start:'YYYY-MM-DD'|'', end:'YYYY-MM-DD'|'', isRange:bool}
// Assumes CALENDAR_YEAR when no year is present.
function parseLooseDate(raw) {
  var empty = {start: '', end: '', isRange: false};
  if (raw == null) return empty;

  if (raw instanceof Date) {
    if (isNaN(raw.getTime())) return empty;
    return {start: ymd(raw.getFullYear(), raw.getMonth() + 1, raw.getDate()), end: '', isRange: false};
  }

  var s = String(raw)
    .replace(/\[merged\]/gi, '')   // Drive export artifact, harmless here
    .replace(/\s+/g, ' ')
    .trim();
  if (!s) return empty;

  // Already ISO
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return {start: ymd(+iso[1], +iso[2], +iso[3]), end: '', isRange: false};

  // Strip parenthetical notes: "April 6 - 10 (onsite)"
  var note = s.match(/^(.*?)\s*\(([^)]*)\)\s*$/);
  if (note) s = note[1].trim();

  // Normalize all dash variants to a plain hyphen
  s = s.replace(/[‐-―−]/g, '-');

  // --- "Month D - D"  (range, may cross a month boundary) ---
  var mRange = s.match(/^([A-Za-z]+)\s+(\d{1,2})\s*-\s*(\d{1,2})$/);
  if (mRange) {
    var mo = MONTH_MAP[mRange[1].toLowerCase()];
    if (mo) {
      var d1 = +mRange[2], d2 = +mRange[3];
      var startS = ymd(CALENDAR_YEAR, mo, d1);
      // End day smaller than start day => range rolls into the next month
      var endMo = mo, endYr = CALENDAR_YEAR;
      if (d2 < d1) { endMo = mo + 1; if (endMo > 12) { endMo = 1; endYr++; } }
      return {start: startS, end: ymd(endYr, endMo, d2), isRange: true};
    }
  }

  // --- "Month D - Month D" ---
  var mFull = s.match(/^([A-Za-z]+)\s+(\d{1,2})\s*-\s*([A-Za-z]+)\s+(\d{1,2})$/);
  if (mFull) {
    var mo1 = MONTH_MAP[mFull[1].toLowerCase()];
    var mo2 = MONTH_MAP[mFull[3].toLowerCase()];
    if (mo1 && mo2) {
      var yr2 = mo2 < mo1 ? CALENDAR_YEAR + 1 : CALENDAR_YEAR;
      return {
        start: ymd(CALENDAR_YEAR, mo1, +mFull[2]),
        end:   ymd(yr2, mo2, +mFull[4]),
        isRange: true
      };
    }
  }

  // --- "Month D"  e.g. "Jan 7", "August 10" ---
  var mSingle = s.match(/^([A-Za-z]+)\s+(\d{1,2})$/);
  if (mSingle) {
    var moS = MONTH_MAP[mSingle[1].toLowerCase()];
    if (moS) return {start: ymd(CALENDAR_YEAR, moS, +mSingle[2]), end: '', isRange: false};
  }

  // --- "M/D" or "M/D/YYYY" or "M/D/YY" ---
  var slash = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (slash) {
    var mm = +slash[1], dd = +slash[2];
    var yy = CALENDAR_YEAR;
    if (slash[3]) {
      yy = +slash[3];
      if (yy < 100) yy += 2000;
    }
    if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
      // Sheets epoch sentinel — treat as empty
      if (yy === 1899) return empty;
      return {start: ymd(yy, mm, dd), end: '', isRange: false};
    }
  }

  // --- "M/D - M/D" ---
  var slashRange = s.match(/^(\d{1,2})\/(\d{1,2})\s*-\s*(\d{1,2})\/(\d{1,2})$/);
  if (slashRange) {
    return {
      start: ymd(CALENDAR_YEAR, +slashRange[1], +slashRange[2]),
      end:   ymd(CALENDAR_YEAR, +slashRange[3], +slashRange[4]),
      isRange: true
    };
  }

  return empty;
}

// =====================================================================
// NORMALIZERS
// =====================================================================
function clean(v) {
  return String(v == null ? '' : v).replace(/\[merged\]/gi, '').replace(/\s+/g, ' ').trim();
}

var TEAMS = ['Paid','Organic','LCM','Brand'];

function normalizeTeam(team, channel) {
  var t = clean(team);
  for (var i = 0; i < TEAMS.length; i++) {
    if (TEAMS[i].toLowerCase() === t.toLowerCase()) return TEAMS[i];
  }
  if (/lifecycle|life cycle|crm|lcm/i.test(t)) return 'LCM';
  // Fall back to inferring from channel
  var c = clean(channel).toLowerCase();
  if (!c) return 'Paid';
  if (/newsletter|email|lifecycle/.test(c)) return 'LCM';
  if (/customer story|case study/.test(c))  return 'Brand';
  if (/paid|progo|programmatic|dooh|ooh|search|direct buy|sponsor/.test(c)) return 'Paid';
  if (/linkedin|reddit|facebook|instagram|tiktok|youtube|threads|^x$|twitter/.test(c)) return 'Organic';
  return 'Paid';
}

// Paid channel names
function normalizePaidChannel(val) {
  var c = clean(val).toLowerCase();
  if (!c) return 'Paid Social';
  if (/progo|programmatic/.test(c))        return 'Programmatic';
  if (/ooh|dooh|out of home/.test(c))      return 'DOOH';
  if (/direct buy/.test(c))                return 'Direct Buy';
  if (/paid search|sem|ppc|pmax/.test(c))  return 'Paid Search';
  if (/paid social|social/.test(c))        return 'Paid Social';
  if (/sponsor/.test(c))                   return 'Sponsorships';
  if (/partner/.test(c))                   return 'Partnerships';
  if (/influencer|creator/.test(c))        return 'Influencer / Creator';
  if (/newsletter|email/.test(c))          return 'Newsletter';
  if (/customer|case study/.test(c))       return 'Customer Story';
  return clean(val);
}

// Organic channel names
function normalizeOrganicChannel(val) {
  var c = clean(val).toLowerCase();
  if (!c) return '';
  if (/linkedin/.test(c))            return 'LinkedIn';
  if (/instagram/.test(c))           return 'Instagram';
  if (/facebook/.test(c))            return 'Facebook';
  if (/threads/.test(c))             return 'Threads';
  if (/reddit/.test(c))              return 'Reddit';
  if (/tiktok|tik tok/.test(c))      return 'TikTok';
  if (/youtube|shorts/.test(c))      return 'YouTube';
  if (/^x$|twitter/.test(c))         return 'X';
  return clean(val);
}

function normalizeStatus(val) {
  var s = clean(val).toLowerCase();
  if (!s) return 'Planned';
  if (/cancel/.test(s))                       return 'Cancelled';
  if (/live|complete|published|posted/.test(s)) return 'Live';
  if (/brief|sent/.test(s))                   return 'Brief Sent';
  if (/review/.test(s))                       return 'In Review';
  if (/approv/.test(s))                       return 'Approved';
  if (/produc|build|progress/.test(s))        return 'In Production';
  return 'Planned';
}

// =====================================================================
// ONE-TIME IMPORT — run this from the Apps Script editor
// Safe to re-run: rows are keyed by source + row number, so a second
// run updates rather than duplicates.
// =====================================================================
function runImport() {
  var summary = {organic: 0, paid: 0, skipped: 0, errors: []};

  try {
    var org = readOrganic();
    summary.organic = org.rows.length;
    summary.skipped += org.skipped;
    writeRows(org.rows);
  } catch(e) {
    summary.errors.push('Organic: ' + e.message);
  }

  try {
    var paid = readPaid();
    summary.paid = paid.rows.length;
    summary.skipped += paid.skipped;
    writeRows(paid.rows);
  } catch(e) {
    summary.errors.push('Paid: ' + e.message);
  }

  try { backfillTeams(); } catch(e) { summary.errors.push('Backfill: ' + e.message); }

  Logger.log(JSON.stringify(summary, null, 2));
  console.log('IMPORT COMPLETE');
  console.log('Organic posts imported: ' + summary.organic);
  console.log('Paid posts imported: '    + summary.paid);
  console.log('Rows skipped (no usable date): ' + summary.skipped);
  if (summary.errors.length) console.log('Errors: ' + summary.errors.join(' | '));
  return summary;
}

// Bulk write: builds a lookup of existing ids, updates in place or appends
function writeRows(rows) {
  if (!rows.length) return;
  var sheet   = getSheet();
  var headers = getHeaders(sheet);
  function ci(k) { return headers.indexOf(k); }

  var existing = {};
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, ci('id') + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0]) existing[String(ids[i][0])] = i + 2;
    }
  }

  var appends = [];
  rows.forEach(function(post) {
    var rowVals = new Array(headers.length).fill('');
    SCHEMA.forEach(function(k) {
      var idx = ci(k);
      if (idx < 0) return;
      if (k === 'platforms' || k === 'audience') {
        rowVals[idx] = JSON.stringify(Array.isArray(post[k]) ? post[k] : []);
      } else if (k === 'evergreen') {
        rowVals[idx] = post[k] === true;
      } else {
        rowVals[idx] = post[k] == null ? '' : post[k];
      }
    });
    var at = existing[String(post.id)];
    if (at) {
      sheet.getRange(at, 1, 1, headers.length).setValues([rowVals]);
    } else {
      appends.push(rowVals);
    }
  });

  if (appends.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, headers.length)
         .setValues(appends);
  }
}

// ---------- ORGANIC: Content calendar ----------
function readOrganic() {
  var ss = SpreadsheetApp.openById(ORGANIC_SHEET_ID);
  var sheet = ss.getSheetByName(ORGANIC_TAB_NAME);
  if (!sheet) {
    // fall back to the first tab
    sheet = ss.getSheets()[0];
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {rows: [], skipped: 0};

  var headers = data[0].map(function(h) { return clean(h); });
  var idx = {};
  headers.forEach(function(h, i) { if (idx[h] === undefined) idx[h] = i; });

  function cell(row, name) {
    return idx[name] !== undefined ? row[idx[name]] : '';
  }

  var rows = [], skipped = 0;
  var lastDate = '';   // carries forward through vertically merged date cells

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var channel = normalizeOrganicChannel(cell(row, 'Channel'));
    var topic   = clean(cell(row, 'Topic'));
    var campaign= clean(cell(row, 'Campaign'));
    var rawDate = cell(row, 'Date');

    var parsed = parseLooseDate(rawDate);

    // Week-divider row: has a date range but no channel and no content
    var isDivider = parsed.isRange && !channel && !topic && !campaign;
    if (isDivider) continue;

    // Merged continuation: date cell blank, but there is content
    var dateStr = parsed.start;
    if (dateStr) {
      lastDate = dateStr;
    } else if (lastDate && (channel || topic || campaign)) {
      dateStr = lastDate;
    }

    if (!dateStr) { if (channel || topic || campaign) skipped++; continue; }
    if (!channel && !topic && !campaign) continue;

    // Boosting Flight can act as an end date
    var endStr = '';
    var flight = parseLooseDate(cell(row, 'Boosting Flight'));
    if (flight.start && flight.start > dateStr) endStr = flight.start;
    else if (parsed.isRange && parsed.end && parsed.end > dateStr) endStr = parsed.end;

    var tags = ['Primary Tag','Secondary Tag','Tertiary Tag']
      .map(function(k) { return clean(cell(row, k)); })
      .filter(Boolean).join(', ');
    var boosting = clean(cell(row, 'Boosting (LK)'));

    var notes = [
      campaign ? 'Campaign: ' + campaign : '',
      boosting ? 'Boosting: ' + boosting : '',
      tags     ? 'Tags: ' + tags         : ''
    ].filter(Boolean).join(' | ');

    rows.push({
      id:         'org_' + (r + 1),
      team:       'Organic',
      date:       dateStr,
      endDate:    endStr,
      channel:    channel,
      platforms:  channel ? [channel] : [],
      postType:   clean(cell(row, 'Post Type')),
      influencer: clean(cell(row, 'Brand or Exec')),
      title:      topic || campaign || (channel + ' post'),
      notes:      notes,
      status:     normalizeStatus(cell(row, 'Status')),
      funnel:     '',
      format:     '',
      audience:   [],
      evergreen:  false,
      source:     'Content calendar'
    });
  }
  return {rows: rows, skipped: skipped};
}

// ---------- PAID: 2026 Campaign Calendar ----------
function readPaid() {
  var ss = SpreadsheetApp.openById(PAID_SHEET_ID);
  var sheet = null;
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getSheetId() == PAID_TAB_GID) { sheet = all[i]; break; }
  }
  if (!sheet) sheet = all[0];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {rows: [], skipped: 0};

  var headers = data[0].map(function(h) { return clean(h); });
  var idx = {};
  headers.forEach(function(h, i) { if (h && idx[h] === undefined) idx[h] = i; });

  // First column is "Date" in some copies, "Launch" in others
  var dateKey = idx['Date'] !== undefined ? 'Date'
              : idx['Launch'] !== undefined ? 'Launch' : null;

  function cell(row, name) {
    return idx[name] !== undefined ? row[idx[name]] : '';
  }

  var rows = [], skipped = 0;

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var rawDate = dateKey ? cell(row, dateKey) : row[0];
    var campaign = clean(cell(row, 'Campaign'));
    var rawChannel = clean(cell(row, 'Channel'));

    if (!campaign && !rawChannel) continue;

    var parsed = parseLooseDate(rawDate);
    if (!parsed.start) { skipped++; continue; }

    var platform = clean(cell(row, 'Platform'));
    var icp      = clean(cell(row, 'ICP')).replace(/^GenBiz$/i, 'Gen Biz');
    var segment  = clean(cell(row, 'SS/SCS'));
    var geo      = clean(cell(row, 'Geo'));
    var vendor   = clean(cell(row, 'Ad Vendor'));
    var launchNotes = clean(cell(row, 'Launch Notes'));
    var variants = clean(cell(row, 'Ad Variants'));
    var totalAds = clean(cell(row, 'Total Ads'));

    var notes = [
      launchNotes ? 'Notes: ' + launchNotes : '',
      geo      ? 'Geo: ' + geo             : '',
      vendor   ? 'Vendor: ' + vendor       : '',
      segment  ? 'Segment: ' + segment     : '',
      variants ? 'Variants: ' + variants   : '',
      totalAds ? 'Total ads: ' + totalAds  : ''
    ].filter(Boolean).join(' | ');

    var audience = [];
    if (icp) audience.push(icp);
    if (segment) audience.push(segment);

    rows.push({
      id:         'paid_' + (r + 1),
      team:       'Paid',
      date:       parsed.start,
      endDate:    parsed.isRange && parsed.end ? parsed.end : '',
      channel:    normalizePaidChannel(rawChannel),
      platforms:  platform ? [platform] : [],
      postType:   clean(cell(row, 'Format')),
      influencer: vendor,
      title:      campaign || (rawChannel + ' campaign'),
      notes:      notes,
      status:     normalizeStatus(cell(row, 'Status')),
      funnel:     clean(cell(row, 'Funnel')),
      format:     clean(cell(row, 'Format')),
      audience:   audience,
      evergreen:  false,
      source:     '2026 Campaign Calendar'
    });
  }
  return {rows: rows, skipped: skipped};
}

// Give any pre-existing post without a team a sensible one, inferred
// from its channel. Run automatically by runImport().
function backfillTeams() {
  var sheet = getSheet();
  if (sheet.getLastRow() < 2) return 0;
  var headers = getHeaders(sheet);
  var teamCol = headers.indexOf('team') + 1;
  var chCol   = headers.indexOf('channel') + 1;
  if (!teamCol || !chCol) return 0;

  var n = sheet.getLastRow() - 1;
  var teams    = sheet.getRange(2, teamCol, n, 1).getValues();
  var channels = sheet.getRange(2, chCol,   n, 1).getValues();
  var changed = 0;

  for (var i = 0; i < n; i++) {
    if (!clean(teams[i][0])) {
      teams[i][0] = normalizeTeam('', channels[i][0]);
      changed++;
    }
  }
  if (changed) sheet.getRange(2, teamCol, n, 1).setValues(teams);
  return changed;
}

// Preview what the import will do without writing anything
function previewImport() {
  var org  = readOrganic();
  var paid = readPaid();
  console.log('ORGANIC: ' + org.rows.length + ' rows, ' + org.skipped + ' skipped');
  console.log('  sample: ' + JSON.stringify(org.rows.slice(0, 3), null, 2));
  console.log('PAID: ' + paid.rows.length + ' rows, ' + paid.skipped + ' skipped');
  console.log('  sample: ' + JSON.stringify(paid.rows.slice(0, 3), null, 2));
}
