/**
 * 監造抽查紀錄 — Google Apps Script 共編版
 *
 * 網頁畫面直接讀取 GitHub 上的最新程式（GitHub 更新後，這裡會自動跟著更新），
 * 所有人的資料存放在共用試算表「監造抽查資料庫」，照片存在同資料夾的「監造抽查照片」。
 *
 * 部署：部署 → 新增部署作業 → 類型「網頁應用程式」
 *   執行身分：存取網頁應用程式的使用者
 *   誰可以存取：所有已登入 Google 帳戶的使用者
 * 共編：把雲端資料夾「室內裝修抽查表單」共用給同事（編輯者），再把網頁應用程式網址給同事。
 */

var SHEET_ID = '1jjk1q4f-2XBrfZj8RFVUQtT2oX-ZL9aw4tYuRb6O4Vs';   // 監造抽查資料庫
var SITE = 'https://raw.githubusercontent.com/cytus201-collab/inspection/main/';
var ASSETS = ['style.css', 'data.js', 'mat-parse.js', 'cloud.js', 'app.js'];
var CHUNK = 45000, MAX_CHUNKS = 20;
var SUM_COLS = ['紀錄', '工程名稱', '工項', '階段', '編號', '檢查位置', '檢查日期', '缺失數', '照片數', '最後修改者', '最後修改時間'];

/* ---------- 網頁 ---------- */
function doGet(e) {
  var fresh = e && e.parameter && e.parameter.refresh;
  return HtmlService.createHtmlOutput(page_(fresh))
    .setTitle('監造抽查紀錄')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function page_(fresh) {
  var cache = CacheService.getScriptCache();
  if (!fresh) {
    var n = Number(cache.get('page:n') || 0);
    if (n) {
      var parts = cache.getAll(Array.apply(null, Array(n)).map(function (_, i) { return 'page:' + i; }));
      var html = '';
      for (var i = 0; i < n; i++) { if (parts['page:' + i] == null) { html = ''; break; } html += parts['page:' + i]; }
      if (html) return html;
    }
  }
  var bust = '?t=' + Date.now();
  var fetch = function (f) {
    var r = UrlFetchApp.fetch(SITE + f + bust, { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('無法讀取 ' + f + '（' + r.getResponseCode() + '）');
    return r.getContentText('UTF-8');
  };
  var out = fetch('index.html')
    .replace(/<link rel="(manifest|icon|apple-touch-icon)"[^>]*>\s*/g, '')
    .replace(/<meta name="viewport"[^>]*>\s*/g, '');
  ASSETS.forEach(function (f) {
    var body = fetch(f).replace(/<\/script/gi, '<\\/script');
    var name = f.replace('.', '\\.');
    if (/\.css$/.test(f)) out = out.replace(new RegExp('<link rel="stylesheet" href="' + name + '[^"]*">'), function () { return '<style>\n' + body + '\n</style>'; });
    else {
      var tag = new RegExp('<script src="' + name + '[^"]*"></script>');
      if (tag.test(out)) out = out.replace(tag, function () { return '<script>\n' + body + '\n</script>'; });
      else out = out.replace(/(<script src="app\.js[^"]*"><\/script>|<script>\n)/, function (m) { return '<script>\n' + body + '\n</script>\n' + m; });
    }
  });
  try {
    var size = 90000, chunks = {}, k = 0;
    for (var p = 0; p < out.length; p += size) chunks['page:' + (k++)] = out.slice(p, p + size);
    chunks['page:n'] = String(k);
    cache.putAll(chunks, 600);
  } catch (err) {}
  return out;
}

/* ---------- 資料同步 ---------- */
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function me_() { try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; } }
function syncSheet_() {
  var s = ss_(), sh = s.getSheetByName('_sync');
  if (!sh) {
    sh = s.insertSheet('_sync');
    sh.getRange(1, 1, 1, 4).setValues([['key', 't', 'by', 'del']]);
    sh.hideSheet();
  }
  return sh;
}
function sumSheet_() {
  var s = ss_(), sh = s.getSheetByName('抽查紀錄總覽');
  if (!sh) {
    sh = s.insertSheet('抽查紀錄總覽', 0);
    sh.getRange(1, 1, 1, SUM_COLS.length).setValues([SUM_COLS]).setFontWeight('bold').setBackground('#000000').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function pull(since) {
  var sh = syncSheet_(), last = sh.getLastRow(), items = [];
  if (last > 1) {
    var width = 4 + MAX_CHUNKS;
    var vals = sh.getRange(2, 1, last - 1, width).getValues();
    vals.forEach(function (r) {
      if (!r[0] || Number(r[1]) <= Number(since || 0)) return;
      items.push({ key: String(r[0]), t: Number(r[1]), by: String(r[2]), del: !!r[3], json: r[3] ? '' : r.slice(4).join('') });
    });
  }
  return { now: Date.now(), me: me_(), items: items };
}

function push(ops) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = syncSheet_(), now = Date.now(), who = me_();
    var last = sh.getLastRow(), rowOf = {};
    if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { if (r[0]) rowOf[String(r[0])] = i + 2; });
    var appends = [];
    (ops || []).forEach(function (op) {
      var parts = [];
      if (!op.del) {
        for (var i = 0; i < op.json.length; i += CHUNK) parts.push(op.json.slice(i, i + CHUNK));
        if (parts.length > MAX_CHUNKS) throw new Error('資料過大：' + op.key);
      }
      while (parts.length < MAX_CHUNKS) parts.push('');
      var row = [op.key, now, who, op.del ? 1 : ''].concat(parts);
      if (rowOf[op.key]) sh.getRange(rowOf[op.key], 1, 1, row.length).setValues([row]);
      else appends.push(row);
      if (op.key.indexOf('record:') === 0) summary_(op, who, now);
    });
    if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, appends[0].length).setValues(appends);
    return { now: now };
  } finally {
    lock.releaseLock();
  }
}

function summary_(op, who, now) {
  var sh = sumSheet_(), last = sh.getLastRow(), at = 0;
  if (last > 1) {
    var f = sh.getRange(2, 1, last - 1, 1).createTextFinder(op.key).matchEntireCell(true).findNext();
    if (f) at = f.getRow();
  }
  if (op.del) { if (at) sh.deleteRow(at); return; }
  var s = op.sum || {};
  var row = SUM_COLS.map(function (c) { return c === '紀錄' ? op.key : c === '最後修改者' ? who : c === '最後修改時間' ? new Date(now) : (s[c] == null ? '' : s[c]); });
  if (at) sh.getRange(at, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
}

/* ---------- 照片 ---------- */
function photoFolder_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('PHOTO_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var parent = DriveApp.getFileById(SHEET_ID).getParents().next();
  var it = parent.getFoldersByName('監造抽查照片');
  var folder = it.hasNext() ? it.next() : parent.createFolder('監造抽查照片');
  props.setProperty('PHOTO_FOLDER', folder.getId());
  return folder;
}
function putPhoto(id, dataUrl) {
  var folder = photoFolder_(), name = id + '.jpg';
  if (folder.getFilesByName(name).hasNext()) return true;
  var b64 = String(dataUrl).split(',')[1];
  folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name));
  return true;
}
function getPhoto(id) {
  var it = photoFolder_().getFilesByName(id + '.jpg');
  if (!it.hasNext()) return null;
  return 'data:image/jpeg;base64,' + Utilities.base64Encode(it.next().getBlob().getBytes());
}

/* ---------- 公文（收文、發文）：PDF 存在「監造公文／工程名稱／收文或發文」 ---------- */
function docRoot_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('DOC_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var parent = DriveApp.getFileById(SHEET_ID).getParents().next();
  var it = parent.getFoldersByName('監造公文');
  var folder = it.hasNext() ? it.next() : parent.createFolder('監造公文');
  props.setProperty('DOC_FOLDER', folder.getId());
  return folder;
}
function sub_(folder, name) {
  var it = folder.getFoldersByName(name);
  return it.hasNext() ? it.next() : folder.createFolder(name);
}
/* 還原 base64（容許 data: 開頭、空白換行、缺少補位、網址安全字元）；失敗時回報長度與開頭供除錯 */
function b64_(x) {
  var s = String(x == null ? '' : x), i = s.indexOf('base64,');
  if (i >= 0) s = s.slice(i + 7);
  s = s.replace(/[^A-Za-z0-9+\/=_-]/g, '').replace(/=+$/, '');
  if (/[-_]/.test(s)) s = s.replace(/-/g, '+').replace(/_/g, '/');
  while (s.length % 4) s += '=';
  try { return Utilities.base64Decode(s); }
  catch (e) { throw new Error('檔案資料無法還原（長度 ' + s.length + '，開頭「' + String(x).slice(0, 30) + '」）'); }
}
/* putDoc(o) 或 putDoc(data, meta)：o / meta = { proj: 工程名稱, dir: '收文' | '發文', name: 檔名, mime, old: 舊檔 id（重新歸檔時移到垃圾桶） }；data 可為 data URL */
function putDoc(o, meta) {
  if (typeof o === 'string') { meta = meta || {}; meta.data = o; o = meta; }
  var proj = String(o.proj || '未命名工程').replace(/[\\/:*?"<>|]/g, '_').slice(0, 80);
  var folder = sub_(sub_(docRoot_(), proj), o.dir === '收文' ? '收文' : '發文');
  if (o.old) { try { DriveApp.getFileById(o.old).setTrashed(true); } catch (e) {} }
  var f = folder.createFile(Utilities.newBlob(b64_(o.data), o.mime || 'application/pdf', o.name || '公文.pdf'));
  return { id: f.getId(), url: f.getUrl() };
}
function getDoc(id) {
  var f = DriveApp.getFileById(id);
  return 'data:' + f.getMimeType() + ';base64,' + Utilities.base64Encode(f.getBlob().getBytes());
}
/* 掃描公文文字辨識（需在編輯器「服務」加入 Drive API）：
   1. 用 Drive.Files.create 把圖片轉成 Google 文件（Google 官方 OCR，語言提示 zh）
   2. 用 Drive 匯出成純文字取回（不需要另外授權 Google 文件權限），再刪除暫存文件 */
function ocrText_(docId) {
  try {
    var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + docId + '/export?mimeType=text/plain',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) return r.getContentText('UTF-8');
    try { return DocumentApp.openById(docId).getBody().getText(); }
    catch (e) { throw new Error('辨識完成但無法取回文字（' + r.getResponseCode() + '；' + e.message + '）'); }
  } finally { try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {} }
}
function ocrBlob_(blob) {
  if (typeof Drive === 'undefined' || !Drive.Files) throw new Error('Drive API 服務未啟用');
  var G = 'application/vnd.google-apps.document', F = Drive.Files, tries = [], errs = [], made = null;
  if (typeof F.create === 'function') [{ ocr: true, ocrLanguage: 'zh' }, { ocrLanguage: 'zh-TW' }, {}].forEach(function (op) { tries.push(function () { return F.create({ name: 'ocr-tmp', mimeType: G }, blob, op); }); });
  if (typeof F.insert === 'function') [{ ocr: true, ocrLanguage: 'zh-TW' }, { ocr: true }].forEach(function (op) { tries.push(function () { return F.insert({ title: 'ocr-tmp', mimeType: G }, blob, op); }); });
  if (!tries.length) throw new Error('Drive API 服務未啟用');
  for (var i = 0; i < tries.length && !made; i++) {
    try { var r = tries[i](); if (r && r.id) made = r.id; }
    catch (e) { var m = String(e && e.message || e).slice(0, 160); if (errs.indexOf(m) < 0) errs.push(m); }
  }
  if (!made) throw new Error('Google 轉檔失敗：' + errs.join('／'));
  return ocrText_(made);
}
function ocrDoc(id) { return ocrBlob_(DriveApp.getFileById(id).getBlob()); }

/* 收文辨識：網頁把掃描公文的前幾頁轉成 JPEG 傳來（避開雲端 OCR 2MB 限制），暫存後辨識，再刪除暫存檔 */
function ocrImg(o) {
  var data = typeof o === 'string' ? o : o.data;
  return ocrBlob_(Utilities.newBlob(b64_(data), 'image/jpeg', 'ocr-src.jpg'));
}
function renameDoc(o) {
  var f = DriveApp.getFileById(o.id); f.setName(String(o.name || f.getName()).slice(0, 200));
  return { id: f.getId(), url: f.getUrl() };
}

/* 第一次使用前由擁有者在編輯器執行一次，完成授權並建立工作表與照片資料夾 */
function setup() {
  syncSheet_(); sumSheet_(); photoFolder_(); docRoot_();
  Logger.log('完成：已建立「抽查紀錄總覽」「_sync」與照片資料夾');
}
