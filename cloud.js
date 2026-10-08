/* Google Apps Script 共編版同步模組
   只在 Apps Script 網頁應用程式內啟用（google.script.run 存在時），GitHub 版不受影響。
   資料以「工程、抽查紀錄、材料」為單位存入共用試算表；同一筆資料以最後儲存者為準。
   只上傳「使用者在本機實際修改過」的資料（dirty），程式改版造成的格式差異不會被當成修改而上傳，
   避免久未開啟的裝置用舊資料覆蓋雲端。每次開啟都會完整下載一次，讓本機副本與雲端一致。 */
(function () {
  'use strict';
  if (!(window.google && google.script && google.script.run)) return;
  var META_KEY = 'inspect-cloud-meta-v1', PULL_MS = 30000, PUSH_MS = 2000;
  var SCHEMA = 7;   // 工程資料格式版本：較舊的畫面讀到較新的資料時暫停上傳，避免覆蓋掉新欄位（例如圖說）
  var outdated = false;
  var meta = load() || { hash: {}, since: 0, photoQ: [] };
  if (!meta.up) meta.up = {};   // 已確認上傳到雲端的照片
  if (!meta.dirty) meta.dirty = {};   // 本機修改過、尚未上傳的資料
  if (!meta.seen) meta.seen = {};     // 上次觀察到的本機資料雜湊（用來判斷使用者修改了哪些）
  var fullDone = false;
  var swept = false;
  var busy = false, timer = null, lastErr = '', me = '', staleOpen = false, started = false;

  function load() { try { return JSON.parse(localStorage.getItem(META_KEY)); } catch (e) { return null; } }
  function keep() { try { localStorage.setItem(META_KEY, JSON.stringify(meta)); } catch (e) {} }
  function call(fn) {
    var args = Array.prototype.slice.call(arguments, 1);
    return new Promise(function (res, rej) {
      var r = google.script.run.withSuccessHandler(res).withFailureHandler(function (e) { rej(new Error(e && e.message || String(e))); });
      r[fn].apply(r, args);
    });
  }
  function A() { return window.APP; }
  function hash(str) { var h = 5381; for (var i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) | 0; return (h >>> 0).toString(36) + ':' + str.length; }

  /* ---------- 本機資料 → 同步單位 ---------- */
  function snapshot() {
    var out = {}, D = A().D, W = {};
    D.works.forEach(function (w) { W[w.id] = w; }); if (D.matWork) W[D.matWork.id] = D.matWork;
    A().state().projects.forEach(function (p) {
      var pj = JSON.stringify({ _v: SCHEMA, id: p.id, info: p.info, works: p.works || null, wOrder: p.wOrder || null, tpl: p.tpl || null, plans: p.plans || [], dwgs: p.dwgs || [], trash: p.trash || [], phrases: p.phrases || null, matOutAt: p.matOutAt || '', created: p.created });
      out['project:' + p.id] = { json: pj };
      Object.keys(p.recs || {}).forEach(function (wid) {
        (p.recs[wid] || []).forEach(function (r) {
          if (r.draft) return;   // 未確認存檔的新表不同步
          var w = W[wid] || { name: wid, kind: '' }, ng = countNg(r);
          out['record:' + r.id] = {
            json: JSON.stringify({ pid: p.id, wid: wid, rec: r }),
            sum: { '工程名稱': p.info.name || '', '工項': w.name, '階段': w.kind === 'phased' ? ['施工前', '施工中檢查', '施工完成檢查'][r.phase] || '' : '', '編號': r.docNo, '檢查位置': r.location || '', '檢查日期': r.checkDate || '', '缺失數': ng, '照片數': (r.photos || []).length }
          };
        });
      });
      (p.mats || []).forEach(function (m) { out['material:' + m.id] = { json: JSON.stringify({ pid: p.id, mat: m }) }; });
    });
    Object.keys(out).forEach(function (k) { out[k].h = hash(out[k].json); });
    return out;
  }
  function countNg(r) {
    var n = 0;
    Object.keys(r.res || {}).forEach(function (k) { if (r.res[k] && r.res[k].r === 'ng' && !r.res[k].off) n++; });
    Object.keys(r.extra || {}).forEach(function (g) { (r.extra[g] || []).forEach(function (x) { if (x.r === 'ng') n++; }); });
    (r.rows || []).forEach(function (x) { if (x.r === 'ng') n++; });
    return n;
  }
  /* 比對本機資料與上次觀察的結果：mark 為 true 時，變動的項目標記為「待上傳」 */
  function observe(mark, snap) {
    snap = snap || snapshot();
    if (mark) {
      Object.keys(snap).forEach(function (k) { if (meta.seen[k] !== snap[k].h) meta.dirty[k] = 1; });
      Object.keys(meta.seen).forEach(function (k) { if (!snap[k]) meta.dirty[k] = 1; });
    }
    var seen = {}; Object.keys(snap).forEach(function (k) { seen[k] = snap[k].h; }); meta.seen = seen;
    keep(); return snap;
  }
  function pending(snap) {
    snap = snap || snapshot();
    var ops = [];
    Object.keys(meta.dirty).forEach(function (k) {
      if (snap[k]) { if (meta.hash[k] !== snap[k].h) ops.push({ key: k, json: snap[k].json, sum: snap[k].sum || null, h: snap[k].h }); else delete meta.dirty[k]; }
      else if (meta.hash[k]) ops.push({ key: k, del: true });
      else delete meta.dirty[k];
    });
    var rank = function (o) { return o.del ? (o.key.indexOf('project:') === 0 ? 3 : 2) : (o.key.indexOf('project:') === 0 ? 0 : 1); };
    return ops.sort(function (a, b) { return rank(a) - rank(b); });
  }

  /* ---------- 上傳 ---------- */
  function push() {
    if (outdated) return Promise.resolve(0);
    var ops = pending();
    if (!ops.length) return Promise.resolve(0);
    var batch = ops.slice(0, 20);
    return call('push', batch.map(function (o) { return { key: o.key, json: o.json || '', sum: o.sum || null, del: !!o.del }; })).then(function () {
      batch.forEach(function (o) { if (o.del) delete meta.hash[o.key]; else meta.hash[o.key] = o.h; delete meta.dirty[o.key]; });
      keep();
      return ops.length > 20 ? push() : ops.length;
    });
  }

  /* ---------- 下載同事的修改 ---------- */
  function apply(items, full) {
    var S = A().state(), snap = snapshot(), changed = [], onServer = {};
    var byId = {}; S.projects.forEach(function (p) { byId[p.id] = p; });
    var rank = function (x) { return x.del ? 2 : x.key.indexOf('project:') === 0 ? 0 : 1; };
    items.sort(function (a, b) { return rank(a) - rank(b); }).forEach(function (x) {
      var local = snap[x.key];
      if (!x.del) onServer[x.key] = 1;
      if (meta.dirty[x.key]) return;                                // 本機有尚未上傳的修改，以本機為準
      if (!x.del && local && local.json === x.json) { meta.hash[x.key] = local.h; return; }
      var i = x.key.indexOf(':'), type = x.key.slice(0, i), id = x.key.slice(i + 1), d;
      if (x.del) {
        if (type === 'project') S.projects = S.projects.filter(function (p) { return p.id !== id; });
        else S.projects.forEach(function (p) {
          if (type === 'material') p.mats = (p.mats || []).filter(function (m) { return m.id !== id; });
          else Object.keys(p.recs || {}).forEach(function (w) { p.recs[w] = p.recs[w].filter(function (r) { return r.id !== id; }); });
        });
        delete meta.hash[x.key]; changed.push(x.key); return;
      }
      try { d = JSON.parse(x.json); } catch (e) { return; }
      if (type === 'project' && (d._v || 0) > SCHEMA) { outdated = true; return; }
      if (type === 'project') {
        var p = byId[id];
        if (!p) { p = { id: id, info: {}, recs: {}, mats: [], created: d.created }; S.projects.push(p); byId[id] = p; }
        p.info = d.info || {}; if (d.works) p.works = d.works; else delete p.works; if ('wOrder' in d) { if (d.wOrder) p.wOrder = d.wOrder; else delete p.wOrder; }
        if (d.tpl) p.tpl = d.tpl; else delete p.tpl;
        /* 舊版畫面上傳的資料沒有這些欄位時保留本機內容，並標記稍後上傳補回 */
        if ('plans' in d) p.plans = d.plans || []; else if ((p.plans || []).length) meta.dirty[x.key] = 1;
        if ('dwgs' in d) p.dwgs = d.dwgs || []; else if ((p.dwgs || []).length) meta.dirty[x.key] = 1;
        if ('trash' in d) p.trash = d.trash || [];
        if ('phrases' in d) { if (d.phrases) p.phrases = d.phrases; else delete p.phrases; } else if (p.phrases) meta.dirty[x.key] = 1; p.matOutAt = d.matOutAt || '';
      } else {
        var pp = byId[d.pid]; if (!pp) return;
        if (type === 'record') {
          if (A().fixRec) A().fixRec(d.wid, d.rec);   // 舊格式紀錄轉為固定代碼
          var list = pp.recs[d.wid] || (pp.recs[d.wid] = []), k = list.findIndex(function (r) { return r.id === id; });
          if (k >= 0) list[k] = d.rec; else list.push(d.rec);
        } else if (type === 'material') {
          var mats = pp.mats || (pp.mats = []), j = mats.findIndex(function (m) { return m.id === id; });
          if (j >= 0) mats[j] = d.mat; else mats.push(d.mat);
        }
      }
      meta.hash[x.key] = hash(x.json); changed.push(x.key);
    });
    /* 完整下載時：本機有、雲端沒有、且不是本機新增的資料 → 移除（例如雲端以版本記錄還原後） */
    if (full) Object.keys(snap).forEach(function (k) {
      if (onServer[k] || meta.dirty[k]) return;
      var i = k.indexOf(':'), type = k.slice(0, i), id = k.slice(i + 1);
      if (type === 'project') S.projects = S.projects.filter(function (p) { return p.id !== id; });
      else S.projects.forEach(function (p) {
        if (type === 'material') p.mats = (p.mats || []).filter(function (m) { return m.id !== id; });
        else Object.keys(p.recs || {}).forEach(function (w) { p.recs[w] = p.recs[w].filter(function (r) { return r.id !== id; }); });
      });
      delete meta.hash[k]; changed.push(k);
    });
    return changed;
  }
  function pull() {
    var full = !fullDone;
    return call('pull', full ? 0 : (meta.since || 0)).then(function (res) {
      me = res.me || me;
      var changed = apply(res.items || [], full);
      fullDone = true;
      meta.since = Math.max(0, res.now - 10000); keep();
      if (changed.length) {
        if (A().migrate) A().migrate();
        observe(false);
        A().persist();
        var c = A().cur(), editing = /^(form|matItem|preview|matImport|dwg)$/.test(c.v);
        var mine = editing && changed.some(function (k) { return k === 'record:' + c.rid || k === 'material:' + c.mid; });
        if (!editing) A().rerender(); else if (mine) staleOpen = true;
      }
      return changed.length;
    });
  }

  /* ---------- 照片 ---------- */
  function uploadPhoto(id, data) {
    if (meta.photoQ.indexOf(id) < 0) meta.photoQ.push(id); keep();
    return call('putPhoto', id, data).then(function () {
      meta.up[id] = 1;
      meta.photoQ = meta.photoQ.filter(function (x) { return x !== id; }); keep(); status();
    }).catch(function (e) { lastErr = '照片上傳失敗：' + e.message; status(); });
  }
  function retryPhotos() {
    var q = meta.photoQ.slice(); if (!q.length) return Promise.resolve();
    return q.reduce(function (pr, id) {
      return pr.then(function () {
        return new Promise(function (res) {
          var rq = indexedDB.open('inspect-photos', 1);
          rq.onsuccess = function () { var g = rq.result.transaction('img').objectStore('img').get(id); g.onsuccess = function () { res(g.result); }; g.onerror = function () { res(null); }; };
          rq.onerror = function () { res(null); };
        }).then(function (d) { if (d) return uploadPhoto(id, d); meta.photoQ = meta.photoQ.filter(function (x) { return x !== id; }); keep(); });
      });
    }, Promise.resolve());
  }
  function getPhoto(id) { return call('getPhoto', id).catch(function () { return null; }); }
  /* 補傳：這台裝置有、但尚未確認上傳的照片（例如匯入備份檔、或先前上傳失敗的） */
  function idbGet(id) {
    return new Promise(function (res) {
      try {
        var rq = indexedDB.open('inspect-photos', 1);
        rq.onupgradeneeded = function () { rq.result.createObjectStore('img'); };
        rq.onsuccess = function () { try { var g = rq.result.transaction('img').objectStore('img').get(id); g.onsuccess = function () { res(g.result || null); }; g.onerror = function () { res(null); }; } catch (e) { res(null); } };
        rq.onerror = function () { res(null); };
      } catch (e) { res(null); }
    });
  }
  function sweepPhotos() {
    if (swept) return Promise.resolve(); swept = true;
    var ids = [];
    A().state().projects.forEach(function (p) {
      Object.keys(p.recs || {}).forEach(function (w) { (p.recs[w] || []).forEach(function (r) { (r.photos || []).forEach(function (ph) { if (!meta.up[ph.id]) ids.push(ph.id); }); }); });
      (p.dwgs || []).forEach(function (g) { if (g.img && !meta.up[g.img]) ids.push(g.img); });
    });
    return ids.reduce(function (pr, id) {
      return pr.then(function () { return idbGet(id).then(function (d) { if (d) return uploadPhoto(id, d); }); });
    }, Promise.resolve());
  }

  /* ---------- 流程 ---------- */
  function cycle() {
    if (busy) return Promise.resolve();
    busy = true; status();
    var first = !fullDone;
    return (first ? pull() : push().then(pull)).then(function () { return push(); }).then(retryPhotos).then(sweepPhotos).then(function () {
      lastErr = ''; meta.last = Date.now(); keep();
    }).catch(function (e) {
      lastErr = /permission|權限|You do not have|找不到|not found/i.test(e.message) ? '沒有共用資料的存取權限，請管理者分享「室內裝修抽查表單」資料夾給你' : e.message;
    }).then(function () { busy = false; status(); });
  }
  function changed() {
    if (!started) return;
    observe(true);
    var pill = document.getElementById('syncPill');
    if (pill && !busy) { pill.hidden = false; pill.textContent = '有新的修改，稍後儲存到雲端'; pill.className = 'sync-pill noprint wait'; }
    clearTimeout(timer); timer = setTimeout(cycle, PUSH_MS);
  }
  function status() {
    var pill = document.getElementById('syncPill'); if (!pill) return;
    var n = started ? Object.keys(meta.dirty).length + meta.photoQ.length : 0, t;
    if (outdated) t = '系統已更新，請重新整理頁面（此畫面已暫停上傳）';
    else if (lastErr) t = lastErr;
    else if (meta.photoQ.length) t = '照片上傳中，尚有 ' + meta.photoQ.length + ' 張';
    else if (busy) t = '雲端同步中…';
    else if (staleOpen) t = '同事已更新此表，返回後重新開啟可看到';
    else if (n) t = '待上傳 ' + n + ' 筆';
    else t = '已儲存到雲端' + (me ? '（' + me + '）' : '');
    pill.hidden = false; pill.textContent = t;
    pill.className = 'sync-pill noprint' + (lastErr || outdated ? ' err' : n || staleOpen ? ' wait' : '');
  }
  function init() {
    started = true;
    /* 第一次使用新同步方式：只有雲端沒有的本機資料視為待上傳，其餘以雲端為準 */
    var snap0 = snapshot();
    if (meta.v !== 2) { meta.dirty = {}; Object.keys(snap0).forEach(function (k) { if (!meta.hash[k]) meta.dirty[k] = 1; }); meta.v = 2; }
    observe(false, snap0);
    window.addEventListener('popstate', function () { if (staleOpen) { staleOpen = false; setTimeout(function () { A().rerender(); status(); }, 0); } });
    /* 切換分頁、鎖螢幕或關閉前立刻上傳（無痕視窗關閉後本機資料會清空） */
    document.addEventListener('visibilitychange', function () { clearTimeout(timer); cycle(); });
    window.addEventListener('pagehide', function () { clearTimeout(timer); cycle(); });
    setInterval(function () { if (!document.hidden) cycle(); }, PULL_MS);
    var card = document.getElementById('h-backup');
    if (card) card.textContent = '備份檔（手動，雲端版通常不需要）';
    var note = document.getElementById('backupNote');
    if (note) note.textContent = '共編版的資料正本存在雲端試算表與雲端硬碟，本機只是暫存副本。使用無痕視窗時，關閉視窗前請確認右下角顯示「已儲存到雲端」。';
    cycle();
  }

  window.Cloud = { init: init, changed: changed, uploadPhoto: uploadPhoto, getPhoto: getPhoto, syncNow: cycle, pending: function () { return pending(); },
    unsynced: function () { return started && !outdated ? Object.keys(meta.dirty).length + meta.photoQ.length : 0; },
    reload: function () { meta = { hash: {}, since: 0, photoQ: meta.photoQ, up: meta.up, dirty: {}, seen: {}, v: 2 }; keep(); fullDone = false; return cycle(); } };
})();
