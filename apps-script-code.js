function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify(getAllPosts()))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'save') return respond(upsertPost(body.post));
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

function getAllPosts() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {posts: []};
  var headers = data[0];
  var posts = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue; // skip empty rows
    var post = {};
    headers.forEach(function(h, idx) {
      post[h] = row[idx];
    });
    // Parse JSON arrays
    ['platforms','types','audience'].forEach(function(key) {
      if (typeof post[key] === 'string' && post[key]) {
        try { post[key] = JSON.parse(post[key]); } catch(e) { post[key] = post[key] ? [post[key]] : []; }
      } else if (!Array.isArray(post[key])) {
        post[key] = [];
      }
    });
    // Normalize date
    if (post.date instanceof Date) {
      var d = post.date;
      post.date = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
    } else {
      post.date = String(post.date);
    }
    // Normalize evergreen boolean
    post.evergreen = post.evergreen === true || post.evergreen === 'TRUE' || post.evergreen === 'true';
    posts.push(post);
  }
  return {posts: posts};
}

function upsertPost(post) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  // Ensure all new columns exist
  var requiredHeaders = ['id','date','endDate','influencer','title','platforms','types','notes','status','color','channel','subType','format','funnel','audience','evergreen'];
  requiredHeaders.forEach(function(h) {
    if (headers.indexOf(h) < 0) {
      sheet.getRange(1, headers.length + 1).setValue(h);
      headers.push(h);
    }
  });

  function colIdx(key) { return headers.indexOf(key); }

  // Serialize arrays to JSON strings
  var platformsStr = JSON.stringify(Array.isArray(post.platforms) ? post.platforms : []);
  var typesStr = JSON.stringify(Array.isArray(post.types) ? post.types : []);
  var audienceStr = JSON.stringify(Array.isArray(post.audience) ? post.audience : []);

  // Look for existing row
  for (var i = 1; i < data.length; i++) {
    if (data[i][colIdx('id')] === post.id) {
      var row = sheet.getRange(i + 1, 1, 1, headers.length);
      var vals = row.getValues()[0];
      vals[colIdx('id')]         = post.id || '';
      vals[colIdx('date')]       = post.date || '';
      vals[colIdx('endDate')]    = post.endDate || '';
      vals[colIdx('influencer')] = post.influencer || '';
      vals[colIdx('title')]      = post.title || '';
      vals[colIdx('platforms')]  = platformsStr;
      vals[colIdx('types')]      = typesStr;
      vals[colIdx('notes')]      = post.notes || '';
      vals[colIdx('status')]     = post.status || '';
      vals[colIdx('color')]      = post.color || '';
      vals[colIdx('channel')]    = post.channel || '';
      vals[colIdx('subType')]    = post.subType || '';
      vals[colIdx('format')]     = post.format || '';
      vals[colIdx('funnel')]     = post.funnel || '';
      vals[colIdx('audience')]   = audienceStr;
      vals[colIdx('evergreen')]  = post.evergreen || false;
      row.setValues([vals]);
      return {success: true};
    }
  }

  // New row
  var newRow = new Array(headers.length).fill('');
  newRow[colIdx('id')]         = post.id || '';
  newRow[colIdx('date')]       = post.date || '';
  newRow[colIdx('endDate')]    = post.endDate || '';
  newRow[colIdx('influencer')] = post.influencer || '';
  newRow[colIdx('title')]      = post.title || '';
  newRow[colIdx('platforms')]  = platformsStr;
  newRow[colIdx('types')]      = typesStr;
  newRow[colIdx('notes')]      = post.notes || '';
  newRow[colIdx('status')]     = post.status || '';
  newRow[colIdx('color')]      = post.color || '';
  newRow[colIdx('channel')]    = post.channel || '';
  newRow[colIdx('subType')]    = post.subType || '';
  newRow[colIdx('format')]     = post.format || '';
  newRow[colIdx('funnel')]     = post.funnel || '';
  newRow[colIdx('audience')]   = audienceStr;
  newRow[colIdx('evergreen')]  = post.evergreen || false;
  sheet.appendRow(newRow);
  return {success: true};
}

function deletePost(id) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return {success: true};
    }
  }
  return {error: 'Not found'};
}
