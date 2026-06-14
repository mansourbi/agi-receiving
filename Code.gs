/*************************************************************
 * AGI Receiving — Google Apps Script backend
 * Bind this to the spreadsheet (Extensions ▸ Apps Script),
 * paste this file, then Deploy ▸ New deployment ▸ Web app:
 *    Execute as:  Me
 *    Who has access:  Anyone
 * Copy the /exec Web app URL into the PWA's API_URL.
 *************************************************************/

var USERS_TAB = 'Users';
var DATA_TAB  = 'Data';

// Canonical fields the app sends, in the order used when the
// Data tab is empty and headers need to be created.
var CANON_ORDER = [
  'receiptserial', 'receiptdate', 'user', 'timestampofscan',
  'pieceid', 'ref', 'size', 'date', 'thk', 'glass', 'procc', 'notes'
];

var CANON_LABEL = {
  receiptserial: 'Receipt Serial',
  receiptdate: 'Receipt Date',
  user: 'User',
  timestampofscan: 'Timestamp of Scan',
  pieceid: 'PIECE ID',
  ref: 'REF',
  size: 'SIZE',
  date: 'DATE',
  thk: 'THK',
  glass: 'GLASS',
  procc: 'PROCC',
  notes: 'Notes'
};

// Header / key aliases → canonical key. Lets sheet headers and
// incoming keys be named loosely and still line up correctly.
var ALIASES = {
  receiptserial: ['receiptserial', 'serial', 'serialnumber', 'serialno', 'receiptno', 'receiptnumber'],
  receiptdate:   ['receiptdate'],
  user:          ['user', 'username', 'operator', 'scannedby', 'receivedby'],
  timestampofscan: ['timestampofscan', 'timestamp', 'scantime', 'scannedat', 'time'],
  pieceid:       ['pieceid', 'piece', 'id', 'uid', 'pieceuid'],
  ref:           ['ref', 'reference'],
  size:          ['size', 'sizemm', 'dimensions', 'dimension', 'dim', 'dims'],
  date:          ['date', 'proddate', 'productiondate', 'dt'],
  thk:           ['thk', 'thickness', 'thick'],
  glass:         ['glass', 'glasstype', 'type'],
  procc:         ['procc', 'proc', 'process', 'processes', 'processing'],
  notes:         ['notes', 'note', 'remark', 'remarks', 'comment', 'comments']
};

function norm_(s) {
  return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
}

function canonOf_(rawKey) {
  var n = norm_(rawKey);
  for (var c in ALIASES) {
    if (ALIASES[c].indexOf(n) !== -1) return c;
  }
  return null;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return jsonOut_({ ok: true, service: 'AGI Receiving', ts: new Date().toISOString() });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    if (action === 'login')  return jsonOut_(login_(body.username, body.password));
    if (action === 'append') return jsonOut_(append_(body.username, body.password, body.rows));
    return jsonOut_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function login_(username, password) {
  var sh = SpreadsheetApp.getActive().getSheetByName(USERS_TAB);
  if (!sh) return { ok: false, error: 'Tab "' + USERS_TAB + '" not found' };
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok: false, error: 'No users defined' };

  var headers = values[0].map(function (h) { return norm_(h); });
  var uCol = -1, pCol = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === 'username' || headers[i] === 'user') uCol = i;
    if (headers[i] === 'password' || headers[i] === 'pass') pCol = i;
  }
  if (uCol < 0 || pCol < 0) {
    return { ok: false, error: 'Users tab needs "Username" and "Password" columns' };
  }

  var u = String(username == null ? '' : username).trim();
  var p = String(password == null ? '' : password);
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][uCol]).trim() === u && String(values[r][pCol]) === p) {
      return { ok: true, username: String(values[r][uCol]).trim() };
    }
  }
  return { ok: false, error: 'Invalid username or password' };
}

function append_(username, password, rows) {
  // Re-validate credentials on write so the endpoint can't be used anonymously.
  var auth = login_(username, password);
  if (!auth.ok) return { ok: false, error: 'Auth failed: ' + auth.error };
  if (!rows || !rows.length) return { ok: false, error: 'No rows to insert' };

  var sh = SpreadsheetApp.getActive().getSheetByName(DATA_TAB);
  if (!sh) sh = SpreadsheetApp.getActive().insertSheet(DATA_TAB);

  var canonHeaders = CANON_ORDER.map(function (c) { return CANON_LABEL[c]; });

  // The Data tab is a valid receiving sheet only if its header row contains a
  // PIECE ID column. If it doesn't (empty tab, or one mistakenly set up with
  // other columns such as Username/Password), reset it to the correct columns
  // so scans are never mismapped.
  var existing = [];
  if (sh.getLastRow() >= 1 && sh.getLastColumn() >= 1) {
    existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  }
  var hasPieceId = existing.some(function (h) { return canonOf_(h) === 'pieceid'; });
  if (!hasPieceId) {
    sh.clear();
    sh.getRange(1, 1, 1, canonHeaders.length).setValues([canonHeaders]);
  }

  var lastCol = sh.getLastColumn();
  var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var colCanon = headerRow.map(function (h) { return canonOf_(h); });

  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var src = rows[i];
    // Re-key the incoming object by canonical field.
    var byCanon = {};
    for (var k in src) {
      var c = canonOf_(k);
      if (c) byCanon[c] = src[k];
    }
    // Build the line in the sheet's own column order.
    var line = [];
    for (var col = 0; col < headerRow.length; col++) {
      var c2 = colCanon[col];
      line.push(c2 && byCanon[c2] != null ? byCanon[c2] : '');
    }
    out.push(line);
  }

  // Single bulk write.
  sh.getRange(sh.getLastRow() + 1, 1, out.length, headerRow.length).setValues(out);
  return { ok: true, inserted: out.length };
}
