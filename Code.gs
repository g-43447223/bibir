var SCRIPT_PROP_NAME = 'SCAN_SHEET_ID';

function doGet(e) {
  var p = e && e.parameter ? e.parameter : null;
  if (p && p.action === 'save') {
    try {
      var payload = {
        condition: String(p.condition || ''),
        confidence: Number(p.confidence || 0),
        metrics: parseJson_(p.metrics || '')
      };
      saveScan(payload);
      return jsonp_({ ok: true }, p.callback);
    } catch (err) {
      return jsonp_({ ok: false, error: String(err.message || err) }, p.callback);
    }
  }
  if (p && p.action === 'list') {
    try {
      return jsonp_({ ok: true, scans: getScans() }, p.callback);
    } catch (err) {
      return jsonp_({ ok: false, error: String(err.message || err) }, p.callback);
    }
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover')
    .setTitle('Kesan Condition Bibir');
}

function parseJson_(s) {
  try {
    var o = JSON.parse(s);
    return o && typeof o === 'object' ? o : {};
  } catch (e) {
    return {};
  }
}

function jsonp_(obj, callbackName) {
  var cb = typeof callbackName === 'string' && /^[A-Za-z_][A-Za-z0-9_]*$/.test(callbackName) ? callbackName : 'callback';
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JAVASCRIPT);
  out.setContent(cb + '(' + JSON.stringify(obj) + ');');
  return out;
}

function saveScan(payload) {
  if (!payload || typeof payload.condition !== 'string' || !payload.condition.trim()) {
    throw new Error('Data tidak lengkap.');
  }
  var allowed = [
    'Bibir kering / menggelupas',
    'Bibir kusam / gelap',
    'Bibir sensitif / mudah pedih',
    'Tiada masalah'
  ];
  if (allowed.indexOf(payload.condition) === -1) {
    throw new Error('Condition tidak dibenarkan.');
  }
  var conf = Number(payload.confidence);
  if (!isFinite(conf) || conf < 0 || conf > 1) {
    throw new Error('Confidence mesti nombor antara 0 dan 1.');
  }
  var sheet = getSheet_();
  var metrics = payload.metrics ? JSON.stringify(payload.metrics) : '';
  sheet.appendRow([new Date(), payload.condition, conf, metrics]);
  return { ok: true };
}

function getScans() {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = values.length - 1; i >= 1; i--) {
    if (rows.length >= 20) break;
    var r = values[i];
    var item = {
      created_at: formatDate_(r[0]),
      condition: String(r[1] || ''),
      confidence: Number(r[2] || 0),
      metrics: null
    };
    if (r[3]) {
      try { item.metrics = JSON.parse(r[3]); } catch (e) { item.metrics = null; }
    }
    rows.push(item);
  }
  return rows;
}

function setup() {
  var ss = getSpreadsheet_();
  getSheet_(ss);
  return 'Sedia. Sheet: ' + ss.getName() + '\nBuka: ' + ss.getUrl();
}

function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(SCRIPT_PROP_NAME);
  if (!id) {
    var ss = SpreadsheetApp.create('Kesan Condition Bibir');
    id = ss.getId();
    props.setProperty(SCRIPT_PROP_NAME, id);
  }
  return SpreadsheetApp.openById(id);
}

function getSheet_(ss) {
  if (!ss) ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('Scans');
  if (!sheet) {
    sheet = ss.insertSheet('Scans');
    sheet.appendRow(['Timestamp', 'Condition', 'Confidence', 'Metrics']);
  }
  return sheet;
}

function formatDate_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  return String(d || '');
}