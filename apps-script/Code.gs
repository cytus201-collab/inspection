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

/* 第一次使用前由擁有者在編輯器執行一次，完成授權並建立工作表與照片資料夾 */
function setup() {
  syncSheet_(); sumSheet_(); photoFolder_();
  Logger.log('完成：已建立「抽查紀錄總覽」「_sync」與照片資料夾');
}
