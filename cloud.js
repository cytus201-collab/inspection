/* Google Apps Script 共編版同步模組
   只在 Apps Script 網頁應用程式內啟用（google.script.run 存在時），GitHub 版不受影響。
   資料以「工程、抽查紀錄、材料」為單位存入共用試算表；同一筆資料以最後儲存者為準。 */
(function () {
  'use strict';
  if (!(window.google && google.script && google.script.run)) return;
  var META_KEY = 'inspect-cloud-meta-v1', PULL_MS = 30000, PUSH_MS = 2000;
  var meta = load() || { hash: {}, since: 0, photoQ: [] };
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
      var pj = JSON.stringify({ id: p.id, info: p.info, works: p.works || null, tpl: p.tpl || null, matOutAt: p.matOutAt || '', created: p.created });
      out['project:' + p.id] = { json: pj };
      Object.keys(p.recs || {}).forEach(function (wid) {
        (p.recs[wid] || []).forEach(function (r) {
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
  function pending(snap) {
    snap = snap || snapshot();
    var ops = [];
    Object.keys(snap).forEach(function (k) { if (meta.hash[k] !== snap[k].h) ops.push({ key: k, json: snap[k].json, sum: snap[k].sum || null, h: snap[k].h }); });
    Object.keys(meta.hash).forEach(function (k) { if (!snap[k]) ops.push({ key: k, del: true }); });
    var rank = function (o) { return o.del ? (o.key.indexOf('project:') === 0 ? 3 : 2) : (o.key.indexOf('project:') === 0 ? 0 : 1); };
    return ops.sort(function (a, b) { return rank(a) - rank(b); });
  }

  /* ---------- 上傳 ---------- */
  function push() {
    var ops = pending();
    if (!ops.length) return Promise.resolve(0);
    var batch = ops.slice(0, 20);
    return call('push', batch.map(function (o) { return { key: o.key, json: o.json || '', sum: o.sum || null, del: !!o.del }; })).then(function () {
      batch.forEach(function (o) { if (o.del) delete meta.hash[o.key]; else meta.hash[o.key] = o.h; });
      keep();
      return ops.length > 20 ? push() : ops.length;
    });
  }

  /* ---------- 下載同事的修改 ---------- */
  function apply(items) {
    var S = A().state(), snap = snapshot(), changed = [];
    var byId = {}; S.projects.forEach(function (p) { byId[p.id] = p; });
    var rank = function (x) { return x.del ? 2 : x.key.indexOf('project:') === 0 ? 0 : 1; };
    items.sort(function (a, b) { return rank(a) - rank(b); }).forEach(function (x) {
      var local = snap[x.key];
      if (local && meta.hash[x.key] !== local.h) return;          // 本機有尚未上傳的修改，以本機為準
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
      if (type === 'project') {
        var p = byId[id];
        if (!p) { p = { id: id, info: {}, recs: {}, mats: [], created: d.created }; S.projects.push(p); byId[id] = p; }
        p.info = d.info || {}; if (d.works) p.works = d.works; else delete p.works;
        if (d.tpl) p.tpl = d.tpl; else delete p.tpl; p.matOutAt = d.matOutAt || '';
      } else {
        var pp = byId[d.pid]; if (!pp) return;
        if (type === 'record') {
          var list = pp.recs[d.wid] || (pp.recs[d.wid] = []), k = list.findIndex(function (r) { return r.id === id; });
          if (k >= 0) list[k] = d.rec; else list.push(d.rec);
        } else if (type === 'material') {
          var mats = pp.mats || (pp.mats = []), j = mats.findIndex(function (m) { return m.id === id; });
          if (j >= 0) mats[j] = d.mat; else mats.push(d.mat);
        }
      }
      meta.hash[x.key] = hash(x.json); changed.push(x.key);
    });
    return changed;
  }
  function pull() {
    return call('pull', meta.since || 0).then(function (res) {
      me = res.me || me;
      var changed = apply(res.items || []);
      meta.since = Math.max(0, res.now - 10000); keep();
      if (changed.length) {
        A().persist();
        var c = A().cur(), editing = /^(form|matItem|preview|matImport)$/.test(c.v);
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

  /* ---------- 流程 ---------- */
  function cycle() {
    if (busy) return Promise.resolve();
    busy = true; status();
    return push().then(pull).then(function () { return push(); }).then(retryPhotos).then(function () {
      lastErr = ''; meta.last = Date.now(); keep();
    }).catch(function (e) {
      lastErr = /permission|權限|You do not have|找不到|not found/i.test(e.message) ? '沒有共用資料的存取權限，請管理者分享「室內裝修抽查表單」資料夾給你' : e.message;
    }).then(function () { busy = false; status(); });
  }
  function changed() {
    if (!started) return;
    var pill = document.getElementById('syncPill');
    if (pill && !busy) { pill.hidden = false; pill.textContent = '有新的修改，稍後儲存到雲端'; pill.className = 'sync-pill noprint wait'; }
    clearTimeout(timer); timer = setTimeout(cycle, PUSH_MS);
  }
  function status() {
    var pill = document.getElementById('syncPill'); if (!pill) return;
    var n = started ? pending().length + meta.photoQ.length : 0, t;
    if (lastErr) t = lastErr;
    else if (busy) t = '雲端同步中…';
    else if (staleOpen) t = '同事已更新此表，返回後重新開啟可看到';
    else if (n) t = '待上傳 ' + n + ' 筆';
    else t = '已儲存到雲端' + (me ? '（' + me + '）' : '');
    pill.hidden = false; pill.textContent = t;
    pill.className = 'sync-pill noprint' + (lastErr ? ' err' : n || staleOpen ? ' wait' : '');
  }
  function init() {
    started = true;
    window.addEventListener('popstate', function () { if (staleOpen) { staleOpen = false; setTimeout(function () { A().rerender(); status(); }, 0); } });
    document.addEventListener('visibilitychange', function () { if (!document.hidden) cycle(); });
    setInterval(function () { if (!document.hidden) cycle(); }, PULL_MS);
    var card = document.getElementById('h-backup');
    if (card) card.textContent = '備份檔（手動，雲端版通常不需要）';
    cycle();
  }

  window.Cloud = { init: init, changed: changed, uploadPhoto: uploadPhoto, getPhoto: getPhoto, syncNow: cycle, pending: function () { return pending(); } };
})();
