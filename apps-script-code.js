// ============================================================
// CONFIGURATION — External Sheet IDs
// ============================================================
var EXT_SHEET_1_ID  = '1kZd027NQ1wd3c9V73C9fJZZdKT9PMGjyDitJuT9uXKU';
var EXT_SHEET_1_GID = 1661186045; // Social / Organic calendar tab
var EXT_SHEET_2_ID  = '1pr5uNAH5VwgHXH5LFVDRwqUfhaQjRRYim_bY_zj9-w4';
var EXT_SHEET_2_GID = 0;          // Paid Media calendar tab (first tab)

// ============================================================
// WEB APP ENTRY POINTS
// ============================================================
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify(getAllPosts()))
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
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// MERGE ALL SOURCES
// ============================================================
function getAllPosts() {
  var posts = [];
  try { posts = posts.concat(getMainSheetPosts()); }
  catch(e) { Logger.log('Main sheet error: ' + e.message); }
  try { posts = posts.concat(getExtSheet1Posts()); }
  catch(e) { Logger.log('Ext sheet 1 error: ' + e.message); }
  try { posts = posts.concat(getExtSheet2Posts()); }
  catch(e) { Logger.log('Ext sheet 2 error: ' + e.message); }
  return {posts: posts};
}

// ============================================================
// MAIN POSTS SHEET (editable)
// ============================================================
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Posts');
  if (!sheet) sheet = initSheet();
  return sheet;
}

function initSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.insertSheet('Posts');
  sheet.appendRow(['id','date','endDate','influencer','title','platforms','types','notes','status','color','channel','subType','format','funnel','audience','evergreen']);
  return sheet;
}

function getMainSheetPosts() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var posts = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var post = {};
    headers.forEach(function(h, idx) { post[h] = row[idx]; });
    ['platforms','types','audience'].forEach(function(key) {
      if (typeof post[key] === 'string' && post[key]) {
        try { post[key] = JSON.parse(post[key]); } catch(e) { post[key] = post[key] ? [post[key]] : []; }
      } else if (!Array.isArray(post[key])) {
        post[key] = [];
      }
    });
    post.date    = formatDateStr(post.date);
    post.endDate = post.endDate ? formatDateStr(post.endDate) : '';
    post.evergreen = post.evergreen === true || post.evergreen === 'TRUE' || post.evergreen === 'true';
    posts.push(post);
  }
  return posts;
}

// ============================================================
// EXTERNAL SHEET 1 — Social / Organic Calendar (read-only)
// Columns: Date | Brand or Exec | Channel | Status | Post Type |
//          Campaign | Topic | Boosting (LK) | Boosting Flight |
//          Primary Tag | Secondary Tag | Tertiary Tag
// ============================================================
function getExtSheet1Posts() {
  var ss    = SpreadsheetApp.openById(EXT_SHEET_1_ID);
  var sheet = getSheetByGid(ss, EXT_SHEET_1_GID);
  if (!sheet) throw new Error('Ext Sheet 1: tab with gid=' + EXT_SHEET_1_GID + ' not found');

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  var posts = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateVal = idx['Date'] !== undefined ? row[idx['Date']] : null;
    if (!dateVal) continue;
    var dateStr = formatDateStr(dateVal);
    if (!dateStr) continue;

    // Boosting Flight can serve as an end date for boosted posts
    var endDateStr = '';
    if (idx['Boosting Flight'] !== undefined && row[idx['Boosting Flight']]) {
      var bf = formatDateStr(row[idx['Boosting Flight']]);
      if (bf && bf > dateStr) endDateStr = bf;
    }

    var tags = [
      idx['Primary Tag']   !== undefined ? String(row[idx['Primary Tag']]   || '') : '',
      idx['Secondary Tag'] !== undefined ? String(row[idx['Secondary Tag']] || '') : '',
      idx['Tertiary Tag']  !== undefined ? String(row[idx['Tertiary Tag']]  || '') : ''
    ].filter(Boolean).join(', ');

    var campaign = idx['Campaign']      !== undefined ? String(row[idx['Campaign']]      || '') : '';
    var boosting = idx['Boosting (LK)'] !== undefined ? String(row[idx['Boosting (LK)']] || '') : '';
    var notesArr = [
      campaign ? 'Campaign: ' + campaign : '',
      boosting ? 'Boosting: ' + boosting : '',
      tags     ? 'Tags: '     + tags     : ''
    ].filter(Boolean);

    var title = '';
    if (idx['Topic']    !== undefined && row[idx['Topic']])    title = String(row[idx['Topic']]);
    else if (idx['Campaign'] !== undefined && row[idx['Campaign']]) title = String(row[idx['Campaign']]);

    posts.push({
      id:         'ext1_' + i,
      date:       dateStr,
      endDate:    endDateStr,
      influencer: idx['Brand or Exec'] !== undefined ? String(row[idx['Brand or Exec']] || '') : '',
      title:      title,
      channel:    mapChannel(idx['Channel'] !== undefined ? row[idx['Channel']] : ''),
      subType:    idx['Post Type'] !== undefined ? String(row[idx['Post Type']] || '') : '',
      format:     '',
      funnel:     '',
      audience:   [],
      platforms:  [],
      types:      [],
      notes:      notesArr.join(' | '),
      status:     mapStatus(idx['Status'] !== undefined ? row[idx['Status']] : ''),
      color:      '',
      evergreen:  false,
      readOnly:   true,
      source:     'Social Calendar'
    });
  }
  return posts;
}

// ============================================================
// EXTERNAL SHEET 2 — Paid Media Calendar (read-only)
// Columns: Date | Campaign | Channel | Platform | Funnel |
//          Format | SS/SCS | ICP | Geo | Ad Vendor |
//          Launch Notes | Ad Variants | Total Ads | Status | …
// ============================================================
function getExtSheet2Posts() {
  var ss    = SpreadsheetApp.openById(EXT_SHEET_2_ID);
  var sheet = getSheetByGid(ss, EXT_SHEET_2_GID);
  if (!sheet) throw new Error('Ext Sheet 2: tab with gid=' + EXT_SHEET_2_GID + ' not found');

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  var posts = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateVal = idx['Date'] !== undefined ? row[idx['Date']] : null;
    if (!dateVal) continue;
    var dateStr = formatDateStr(dateVal);
    if (!dateStr) continue;

    var platform = idx['Platform']  !== undefined ? String(row[idx['Platform']]  || '') : '';
    var icp      = idx['ICP']       !== undefined ? String(row[idx['ICP']]       || '') : '';
    var notesArr = [
      idx['Launch Notes'] !== undefined && row[idx['Launch Notes']]  ? 'Notes: '    + row[idx['Launch Notes']]  : '',
      idx['Geo']          !== undefined && row[idx['Geo']]           ? 'Geo: '      + row[idx['Geo']]           : '',
      idx['Ad Vendor']    !== undefined && row[idx['Ad Vendor']]     ? 'Vendor: '   + row[idx['Ad Vendor']]     : '',
      idx['Ad Variants']  !== undefined && row[idx['Ad Variants']]   ? 'Variants: ' + row[idx['Ad Variants']]   : '',
      idx['SS/SCS']       !== undefined && row[idx['SS/SCS']]        ? 'Segment: '  + row[idx['SS/SCS']]        : ''
    ].filter(Boolean);

    posts.push({
      id:         'ext2_' + i,
      date:       dateStr,
      endDate:    '',
      influencer: idx['Ad Vendor']  !== undefined ? String(row[idx['Ad Vendor']]  || '') : '',
      title:      idx['Campaign']   !== undefined ? String(row[idx['Campaign']]   || '') : '',
      channel:    mapChannel(idx['Channel'] !== undefined ? row[idx['Channel']] : ''),
      subType:    platform,
      format:     idx['Format']     !== undefined ? String(row[idx['Format']]     || '') : '',
      funnel:     idx['Funnel']     !== undefined ? String(row[idx['Funnel']]     || '') : '',
      audience:   icp ? [icp] : [],
      platforms:  platform ? [platform] : [],
      types:      [],
      notes:      notesArr.join(' | '),
      status:     mapStatus(idx['Status'] !== undefined ? row[idx['Status']] : ''),
      color:      '',
      evergreen:  false,
      readOnly:   true,
      source:     'Paid Media'
    });
  }
  return posts;
}

// ============================================================
// HELPERS
// ============================================================
function getSheetByGid(spreadsheet, gid) {
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() == gid) return sheets[i];
  }
  return sheets.length ? sheets[0] : null; // fallback to first sheet
}

function formatDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return val.getFullYear() + '-' +
      String(val.getMonth() + 1).padStart(2, '0') + '-' +
      String(val.getDate()).padStart(2, '0');
  }
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return d.getFullYear() + '-' +
      String(d.getMonth() + 1).padStart(2, '0') + '-' +
      String(d.getDate()).padStart(2, '0');
  }
  return '';
}

// Map arbitrary channel text to one of the calendar's 9 channel names
function mapChannel(val) {
  if (!val) return 'Paid Social';
  var c = String(val).toLowerCase().trim();
  if (c.includes('influencer') || c.includes('creator'))                  return 'Influencer / Creator';
  if (c.includes('newsletter') || c.includes('email'))                    return 'Newsletter';
  if (c.includes('programmatic') || c.includes('display') ||
      c.includes('ctv') || c.includes('native') || c.includes('progo'))   return 'Programmatic';
  if (c.includes('sponsor'))                                               return 'Sponsorships';
  if (c.includes('partner'))                                               return 'Partnerships';
  if (c.includes('dooh') || c.includes('out of home') || c.includes('ooh')) return 'DOOH';
  if (c.includes('search') || c.includes('sem') ||
      c.includes('ppc') || c.includes('pmax'))                             return 'Paid Search';
  if (c.includes('customer') || c.includes('case study'))                 return 'Customer Story';
  // Social platform names and generic "paid social"
  if (c.includes('social') || c.includes('linkedin') || c.includes('meta') ||
      c.includes('facebook') || c.includes('instagram') ||
      c.includes('twitter') || c.includes('tiktok') || c.includes('youtube')) return 'Paid Social';
  return val; // keep original if nothing matched — calendar will still render it
}

// Map arbitrary status text to one of the calendar's 5 status names
function mapStatus(val) {
  if (!val) return 'Planned';
  var s = String(val).toLowerCase().trim();
  if (s.includes('brief') || s.includes('sent'))      return 'Brief Sent';
  if (s.includes('review'))                            return 'In Review';
  if (s.includes('approv'))                            return 'Approved';
  if (s.includes('produc') || s.includes('build') ||
      s.includes('launch'))                            return 'In Production';
  if (s.includes('plan')   || s.includes('draft') ||
      s.includes('schedul') || s.includes('upcoming')) return 'Planned';
  return 'Planned';
}

// ============================================================
// WRITE — main sheet only; external sheet posts are read-only
// ============================================================
function upsertPost(post) {
  if (String(post.id || '').match(/^ext[12]_/)) {
    return {error: 'External posts cannot be edited from the calendar.'};
  }

  var sheet   = getSheet();
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];

  var requiredHeaders = ['id','date','endDate','influencer','title','platforms','types','notes','status','color','channel','subType','format','funnel','audience','evergreen'];
  requiredHeaders.forEach(function(h) {
    if (headers.indexOf(h) < 0) {
      sheet.getRange(1, headers.length + 1).setValue(h);
      headers.push(h);
    }
  });

  function colIdx(key) { return headers.indexOf(key); }
  var platformsStr = JSON.stringify(Array.isArray(post.platforms) ? post.platforms : []);
  var typesStr     = JSON.stringify(Array.isArray(post.types)     ? post.types     : []);
  var audienceStr  = JSON.stringify(Array.isArray(post.audience)  ? post.audience  : []);

  for (var i = 1; i < data.length; i++) {
    if (data[i][colIdx('id')] === post.id) {
      var row  = sheet.getRange(i + 1, 1, 1, headers.length);
      var vals = row.getValues()[0];
      vals[colIdx('id')]         = post.id        || '';
      vals[colIdx('date')]       = post.date       || '';
      vals[colIdx('endDate')]    = post.endDate    || '';
      vals[colIdx('influencer')] = post.influencer || '';
      vals[colIdx('title')]      = post.title      || '';
      vals[colIdx('platforms')]  = platformsStr;
      vals[colIdx('types')]      = typesStr;
      vals[colIdx('notes')]      = post.notes      || '';
      vals[colIdx('status')]     = post.status     || '';
      vals[colIdx('color')]      = post.color      || '';
      vals[colIdx('channel')]    = post.channel    || '';
      vals[colIdx('subType')]    = post.subType    || '';
      vals[colIdx('format')]     = post.format     || '';
      vals[colIdx('funnel')]     = post.funnel     || '';
      vals[colIdx('audience')]   = audienceStr;
      vals[colIdx('evergreen')]  = post.evergreen  || false;
      row.setValues([vals]);
      return {success: true};
    }
  }

  var newRow = new Array(headers.length).fill('');
  newRow[colIdx('id')]         = post.id        || '';
  newRow[colIdx('date')]       = post.date       || '';
  newRow[colIdx('endDate')]    = post.endDate    || '';
  newRow[colIdx('influencer')] = post.influencer || '';
  newRow[colIdx('title')]      = post.title      || '';
  newRow[colIdx('platforms')]  = platformsStr;
  newRow[colIdx('types')]      = typesStr;
  newRow[colIdx('notes')]      = post.notes      || '';
  newRow[colIdx('status')]     = post.status     || '';
  newRow[colIdx('color')]      = post.color      || '';
  newRow[colIdx('channel')]    = post.channel    || '';
  newRow[colIdx('subType')]    = post.subType    || '';
  newRow[colIdx('format')]     = post.format     || '';
  newRow[colIdx('funnel')]     = post.funnel     || '';
  newRow[colIdx('audience')]   = audienceStr;
  newRow[colIdx('evergreen')]  = post.evergreen  || false;
  sheet.appendRow(newRow);
  return {success: true};
}

function deletePost(id) {
  if (String(id || '').match(/^ext[12]_/)) {
    return {error: 'External posts cannot be deleted from the calendar.'};
  }
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return {success: true};
    }
  }
  return {error: 'Not found'};
}
