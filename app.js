(function () {
  'use strict';
  var D = window.FORM_DATA;
  var SYM = { ok: '○', ng: '╳', na: '／' };
  var PH_SHORT = ['施工前', '施工中', '施工完成'];
  var KEY = 'inspect-app-v3';
  var CDN = {
    pdf: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
    pdfWorker: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
    pdfCmaps: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/cmaps/',
    xlsx: 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    h2c: 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
    jspdf: 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js'
  };
  var $ = function (id) { return document.getElementById(id); };
  var WORK = {}; D.works.forEach(function (w) { WORK[w.id] = w; });
  if (D.matWork) WORK[D.matWork.id] = D.matWork;
  function isFix(w) { return w.kind === 'phased' || w.kind === 'matform'; }
  var PHASED = D.works.filter(function (w) { return w.kind === 'phased'; });

  var INFO_FIELDS = [
    { k: 'name', l: '工程名稱', req: true, wide: true },
    { k: 'code', l: '工程編號' },
    { k: 'owner', l: '業主' },
    { k: 'supervisor', l: '監造單位' },
    { k: 'contractor', l: '承包商' },
    { k: 'amount', l: '契約金額', im: 'numeric' },
    { k: 'startDate', l: '開工日期', t: 'date' },
    { k: 'durType', l: '工期類型', t: 'select', opts: ['日曆天', '工作天', '限期完工'] },
    { k: 'duration', l: '契約工期（天）', t: 'number' },
    { k: 'planEnd', l: '預定完工日期', t: 'date', calc: true },
    { k: 'changeCount', l: '變更次數', t: 'number' },
    { k: 'extendDays', l: '工期展延天數', t: 'number' },
    { k: 'changedDuration', l: '變更後工期（天）', t: 'number' },
    { k: 'changedAmount', l: '變更後金額', im: 'numeric' },
    { k: 'changedEnd', l: '變更後完工日期', t: 'date' },
    { k: 'inspector', l: '監造人員' },
    { k: 'siteManager', l: '工地主任' },
    { k: 'note', l: '備註', t: 'textarea', wide: true }
  ];
  var MAT_BASIC = [
    { k: 'no', l: '契約詳細表項次' }, { k: 'name', l: '材料／設備名稱', wide: true }, { k: 'qty', l: '契約數量' },
    { k: 'sampleTest', l: '是否取樣試驗', t: 'select', opts: ['', '是', '否'] },
    { k: 'factory', l: '是否驗廠／廠驗', t: 'select', opts: ['', '是', '否'] }
  ];
  var MAT_TRACK = [
    { k: 'planSubmit', l: '預定送審日期', t: 'date' }, { k: 'actualSubmit', l: '實際送審日期', t: 'date' },
    { k: 'factoryDate', l: '驗廠／廠驗日期', t: 'date' },
    { k: 'reviewDate', l: '審查日期', t: 'date' },
    { k: 'reviewResult', l: '審查結果', t: 'select', opts: ['', '核定', '修正後再送', '不核定'] },
    { k: 'archiveNo', l: '歸檔編號' },
    { k: 'arrivalDate', l: '進場日期', t: 'date' },
    { k: 'testDate', l: '取樣試驗日期', t: 'date' },
    { k: 'testResult', l: '試驗結果', t: 'select', opts: ['', '合格', '不合格'] }
  ];
  var DOCS = [['vendor', '協力廠商資料'], ['catalog', '型錄'], ['test', '相關試驗報告'], ['sample', '樣品色票']];
  var MSTAT = {
    todo: '待送審', review: '審查中', approved: '已核定待進場', arrived: '已進場待抽查', ok: '抽查合格', ng: '抽查不合格'
  };
  var MSTAT_ORDER = ['ok', 'ng', 'arrived', 'approved', 'review', 'todo'];

  /* ---------- 工具 ---------- */
  function today() { var d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); }
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function roc(d, sep) {
    var a = (d || '').split('-'); if (a.length !== 3) return '';
    return sep ? (a[0] - 1911) + sep + a[1] + sep + a[2] : (a[0] - 1911) + ' 年 ' + (+a[1]) + ' 月 ' + (+a[2]) + ' 日';
  }
  function addDays(d, n) { var t = new Date(d + 'T00:00:00'); t.setDate(t.getDate() + n); t.setMinutes(t.getMinutes() - t.getTimezoneOffset()); return t.toISOString().slice(0, 10); }
  function load() { try { return JSON.parse(localStorage.getItem(KEY)); } catch (e) { return null; } }
  var saveMsg = '';
  /* 本機儲存：主要存在 IndexedDB（沒有 5MB 上限）；資料不大時另存一份到 localStorage 當備援 */
  var LS_MAX = 1500000, saveTimer = null, idbOK = !!window.indexedDB, stateSize = 0;
  function writeLocal() {
    clearTimeout(saveTimer); saveTimer = null;
    S.savedAt = Date.now();
    var txt = JSON.stringify(S); stateSize = txt.length;
    var lsOK = false;
    try { if (txt.length < LS_MAX) { localStorage.setItem(KEY, txt); lsOK = true; } else localStorage.removeItem(KEY); } catch (e) {}
    if (idbOK) IDB.put('state:v1', txt).then(function () { saveMsg = ''; }, function () { idbOK = false; if (!lsOK) saveMsg = '裝置空間不足或瀏覽器禁止暫存，請匯出備份'; });
    else saveMsg = lsOK ? '' : '裝置空間不足或瀏覽器禁止暫存，請匯出備份';
  }
  function save() {
    clearTimeout(saveTimer); saveTimer = setTimeout(writeLocal, 250);
    if (window.Cloud) window.Cloud.changed();
  }
  function flushLocal() { if (saveTimer) writeLocal(); }
  window.addEventListener('pagehide', flushLocal);
  document.addEventListener('visibilitychange', function () { if (document.hidden) flushLocal(); });
  function loadScript(src) {
    return new Promise(function (res, rej) {
      var s = document.createElement('script'); s.src = src; s.onload = res;
      s.onerror = function () { rej(new Error('無法下載解析元件，請確認網路連線後再試')); };
      document.head.appendChild(s);
    });
  }
  function fieldHtml(prefix, f, v) {
    var id = prefix + f.k, input;
    v = v == null ? '' : v;
    if (f.t === 'select') input = '<select id="' + id + '">' + f.opts.map(function (o) { return '<option value="' + esc(o) + '"' + (o === v ? ' selected' : '') + '>' + (o || '—') + '</option>'; }).join('') + '</select>';
    else if (f.t === 'textarea') input = '<textarea id="' + id + '" rows="3">' + esc(v) + '</textarea>';
    else input = '<input id="' + id + '" type="' + (f.t || 'text') + '"' + (f.im ? ' inputmode="' + f.im + '"' : '') + (f.t === 'number' ? ' min="0" inputmode="numeric"' : '') + ' value="' + esc(v) + '">';
    return '<div class="f' + (f.wide ? ' wide' : '') + '"><span class="f-head"><label for="' + id + '" class="' + (f.req ? 'req' : '') + '">' + f.l + '</label>' +
      (f.calc ? '<button type="button" class="link" id="calcEnd">推算</button>' : '') + '</span>' + input + '</div>';
  }

  var S = load() || { v: 3, projects: [] };
  S.projects.forEach(function (p) { if (!p.mats) p.mats = []; });

  function proj(id) { return S.projects.filter(function (p) { return p.id === id; })[0]; }
  function recs(p, wid) { return p.recs[wid] || (p.recs[wid] = []); }
  function recById(p, wid, rid) { return recs(p, wid).filter(function (r) { return r.id === rid; })[0]; }
  function matById(p, mid) { return p.mats.filter(function (m) { return m.id === mid; })[0]; }
  function cell(r, key) { return r.res[key] || (r.res[key] = { r: '', note: '' }); }
  function k(g, i) { return g + '-' + i; }
  function flowName(i) { return i === 0 ? '施工前' : PH_SHORT[i] + '檢查'; }
  function stats(w, r) {
    if (w.kind === 'photo') { var n = (r.photos || []).length; return { t: n, d: n, ng: 0 }; }
    var t = 0, d = 0, ng = 0;
    if (w.kind === 'irregular') {
      (r.rows || []).forEach(function (x) { t++; if (x.r) d++; if (x.r === 'ng') ng++; });
      return { t: t, d: d, ng: ng };
    }
    effGroups(w, r).forEach(function (g) { g.items.forEach(function (it) { t++; if (it.c.r) d++; if (it.c.r === 'ng') ng++; }); });
    return { t: t, d: d, ng: ng };
  }
  /* 實際使用的抽查項目：預設項目（可改標準、可移除）＋自行新增的項目 */
  /* 抽查項目以固定代碼對應：r.res[項目代碼]、r.extra[分組代碼]。已確認的表使用存檔當下的標準快照（r.snap），法規更新不影響舊表 */
  function hasStd(w) { return !!(w && w.phases && w.phases.length && w.kind !== 'irregular' && w.kind !== 'photo'); }
  function baseGroups(w, r) { return r.snap && !r.draft ? r.snap.groups : w.phases[r.phase].groups; }
  function snapOf(w, r) {
    return { ver: D.version, groups: w.phases[r.phase].groups.map(function (g) {
      return { id: g.id, name: g.name, items: g.items.map(function (it) { return { id: it.id, name: it.name, std: it.std, hint: it.hint || '', ref: it.ref || '' }; }) };
    }) };
  }
  function snapStale(w, r) { return !!(r.snap && JSON.stringify(r.snap.groups) !== JSON.stringify(snapOf(w, r).groups)); }
  function effGroups(w, r) {
    if (!r.extra) r.extra = {};
    return baseGroups(w, r).map(function (g, gi) {
      var items = [], removed = [];
      g.items.forEach(function (it) {
        var key = it.id, c = r.res[key] || { r: '', note: '' };
        if (c.off) { removed.push({ key: key, name: it.name }); return; }
        items.push({ key: key, name: it.name, std: c.std != null ? c.std : it.std, def: it.std, hint: it.hint, ref: it.ref, c: c });
      });
      (r.extra[g.id] || []).forEach(function (x, xi) { items.push({ x: g.id + '|' + xi, name: x.name, std: x.std, hint: '', ref: '', c: x, custom: true }); });
      return { name: g.name, gi: gi, gid: g.id, items: items, removed: removed };
    });
  }
  /* 舊格式（第幾組-第幾項）轉為固定代碼；轉換時的標準即當時填寫的標準 */
  function migrateRec(w, r) {
    if (!hasStd(w) || r.kv === 2) return false;
    var ph = w.phases[r.phase] || w.phases[0], res = {}, extra = {};
    Object.keys(r.res || {}).forEach(function (key) {
      var m = /^(\d+)-(\d+)$/.exec(key), g = m && ph.groups[+m[1]], it = g && g.items[+m[2]];
      res[it ? it.id : key] = r.res[key];
    });
    Object.keys(r.extra || {}).forEach(function (gi) { var g = ph.groups[+gi]; extra[g ? g.id : gi] = r.extra[gi]; });
    r.res = res; r.extra = extra; r.kv = 2;
    if (!r.draft && !r.snap) r.snap = snapOf(w, r);
    return true;
  }
  function migrateTpl(p) {
    if (!p.tpl) return false; var ch = false;
    Object.keys(p.tpl).forEach(function (wid) {
      var w = WORK[wid]; if (!w || !w.phases) return;
      Object.keys(p.tpl[wid] || {}).forEach(function (pi) {
        var t = p.tpl[wid][pi], ph = w.phases[+pi]; if (!t || t.kv === 2 || !ph) return;
        var cells = {}, extra = {};
        Object.keys(t.cells || {}).forEach(function (key) { var m = /^(\d+)-(\d+)$/.exec(key), g = m && ph.groups[+m[1]], it = g && g.items[+m[2]]; cells[it ? it.id : key] = t.cells[key]; });
        Object.keys(t.extra || {}).forEach(function (gi) { var g = ph.groups[+gi]; extra[g ? g.id : gi] = t.extra[gi]; });
        p.tpl[wid][pi] = { cells: cells, extra: extra, kv: 2 }; ch = true;
      });
    });
    return ch;
  }
  function migrateAll() {
    var ch = false;
    S.projects.forEach(function (p) {
      Object.keys(p.recs || {}).forEach(function (wid) { (p.recs[wid] || []).forEach(function (r) { if (migrateRec(WORK[wid], r)) ch = true; }); });
      (p.trash || []).forEach(function (t) { if (t.type === 'rec' && migrateRec(WORK[t.wid], t.data)) ch = true; });
      if (migrateTpl(p)) ch = true;
    });
    return ch;
  }
  function recTitle(w, r) { return w.kind === 'phased' ? w.name + '(' + (r.phase + 1) + ') ' + flowName(r.phase) : w.kind === 'matform' ? '材料進場抽查紀錄' : w.name; }
  /* 新增一筆抽查紀錄（套用本工程預設標準） */
  function newRecord(p, w, pi, extra) {
    var list = recs(p, w.id);
    var seq = list.filter(function (r) { return r.phase === pi; }).reduce(function (m, r) { return Math.max(m, r.seq); }, 0) + 1;
    var r = {
      id: uid(), phase: pi, seq: seq,
      docNo: (p.info.code ? p.info.code + '-' : '') + w.code + (w.kind === 'phased' ? (pi + 1) : '') + '-' + ('00' + seq).slice(-3),
      location: '', checkDate: today(), res: {}, extra: {}, fix: 'none', fixDate: '', fixPerson: '', sig: {}, created: Date.now(), draft: true, kv: 2
    };
    if (!isFix(w)) { r.rows = []; r.result = ''; r.exec = ''; r.deadline = ''; r.method = ''; r.methodNote = ''; r.re = ''; r.reTime = ''; r.remark = ''; }
    var t = p.tpl && p.tpl[w.id] && p.tpl[w.id][pi];
    if (t) { r.res = JSON.parse(JSON.stringify(t.cells || {})); r.extra = JSON.parse(JSON.stringify(t.extra || {})); }
    Object.assign(r, extra || {});
    list.push(r); save();
    return r;
  }
  function matStatus(m) {
    if (m.inspResult === 'ok') return 'ok';
    if (m.inspResult === 'ng') return 'ng';
    if (m.arrivalDate) return 'arrived';
    if (m.reviewResult === '核定') return 'approved';
    if (m.actualSubmit) return 'review';
    return 'todo';
  }

  /* ---------- 導覽 ---------- */
  var cur = { v: 'projects' };
  try { history.scrollRestoration = 'manual'; } catch (e) {}
  function go(state, replace) {
    cur = state;
    try { history[replace ? 'replaceState' : 'pushState'](state, ''); } catch (e) {}
    render();
  }
  window.addEventListener('popstate', function (e) { cur = e.state || { v: 'projects' }; render(); });
  $('navBack').addEventListener('click', function () { history.back(); });
  var VIEWS = ['vProjects', 'vProject', 'vWork', 'vForm', 'vMaterials', 'vMatItem', 'vMatImport', 'vPreview', 'vDwgs', 'vDwg'];
  function show(id) { VIEWS.forEach(function (v) { $(v).hidden = v !== id; }); window.scrollTo(0, 0); }
  function topbar(title, sub, back, right) {
    $('tbTitle').textContent = title; $('tbSub').textContent = sub || '';
    $('navBack').hidden = !back; $('navBack').textContent = back ? '‹ ' + back : '';
    $('tbRight').textContent = right || '';
    document.title = title;
  }
  /* 新增抽查後未按「確定」就離開：不列入紀錄（有填內容時可復原） */
  function touched(r) {
    if (r.location || (r.photos || []).length || (r.rows || []).length || Object.keys(r.sig || {}).length) return true;
    if (Object.keys(r.res || {}).some(function (k) { var c = r.res[k]; return c && (c.r || c.note); })) return true;
    return Object.keys(r.extra || {}).some(function (g) { return (r.extra[g] || []).some(function (x) { return x.r || x.note; }); });
  }
  function purgeDrafts() {
    var keep = cur.v === 'form' ? cur.rid : cur.link ? cur.link.split('|')[1] : null, gone = null, n = 0;
    S.projects.forEach(function (p) {
      Object.keys(p.recs || {}).forEach(function (wid) {
        var list = p.recs[wid];
        for (var i = list.length - 1; i >= 0; i--) {
          var r = list[i]; if (!r.draft || r.id === keep) continue;
          list.splice(i, 1); n++;
          if (touched(r)) gone = { pid: p.id, wid: wid, rec: r }; else (r.photos || []).forEach(function (x) { IDB.del(x.id).catch(function () {}); });
        }
      });
    });
    if (n) save();
    if (gone) setTimeout(function () {
      ask('剛才的抽查表尚未確認存檔', '已不列入紀錄。要回去繼續填寫嗎？', '捨棄', '繼續填寫', function () {
        var p = proj(gone.pid); if (!p) return;
        recs(p, gone.wid).push(gone.rec); save();
        go({ v: 'form', pid: gone.pid, wid: gone.wid, rid: gone.rec.id });
      });
    }, 0);
  }
  function render() {
    purgeDrafts(); purgeTrash();
    var p = cur.pid && proj(cur.pid), w = cur.wid && WORK[cur.wid];
    if (cur.v !== 'projects' && !p) cur = { v: 'projects' };
    else if (/^(work|form|preview)$/.test(cur.v) && !w) cur = { v: 'project', pid: cur.pid };
    else if (/^(form|preview)$/.test(cur.v) && !recById(p, cur.wid, cur.rid)) cur = { v: 'work', pid: cur.pid, wid: cur.wid };
    else if (cur.v === 'matItem' && !matById(p, cur.mid)) cur = { v: 'materials', pid: cur.pid };
    else if (cur.v === 'matImport' && !pendingImport) cur = { v: 'materials', pid: cur.pid };
    else if (/^(dwg|dwgPrint)$/.test(cur.v) && !dwgById(p, cur.did)) cur = { v: 'dwgs', pid: cur.pid };
    $('sumBar').hidden = true;
    ({
      projects: renderProjects, project: renderProject, work: renderWork, form: renderForm, preview: renderPreview,
      materials: renderMaterials, matItem: renderMatItem, matImport: renderMatImport, matPrint: renderMatPrint, sumPrint: renderSumPrint,
      dwgs: renderDwgs, dwg: renderDwg, dwgPrint: renderDwgPrint
    })[cur.v]();
  }

  /* ========== 1 工程清單 ========== */
  /* 個人版／共編版切換：個人版在 GitHub，共編版在 Google Apps Script */
  var PERSONAL_URL = 'https://cytus201-collab.github.io/inspection/', CO_URL_DEFAULT = 'https://script.google.com/macros/s/AKfycbxci9ZymyINwrKblPts7ACxfFevhfgwMF3QT9JNw6hPIrUiNRnbywxC1bOuOtVhSEngDA/exec';
  var IS_CO = !!(window.google && google.script && google.script.run);
  function coUrl() {
    try { var q = new URLSearchParams(location.search).get('co'); if (q && /^https:\/\/script\.google\.com\//.test(q)) { localStorage.setItem('coedit-url', q); history.replaceState(history.state, '', location.pathname); } } catch (e) {}
    try { return localStorage.getItem('coedit-url') || CO_URL_DEFAULT; } catch (e) { return CO_URL_DEFAULT; }
  }
  function myExecUrl() { var m = /https:\/\/script\.google\.com\/(a\/[^/]+\/)?macros\/(u\/\d+\/)?s\/[^/?#]+\/exec/.exec(document.referrer || ''); return m ? m[0] : ''; }
  function renderSwitch() {
    var a = $('msPersonal'), b = $('msCo');
    a.classList.toggle('on', !IS_CO); b.classList.toggle('on', IS_CO);
    if (IS_CO) { var me = myExecUrl(); a.href = PERSONAL_URL + (me ? '?co=' + encodeURIComponent(me) : ''); a.target = '_top'; b.removeAttribute('href'); }
    else { var u = coUrl(); b.href = u || '#'; a.removeAttribute('href'); }
  }
  $('msCo').addEventListener('click', function (e) {
    if (IS_CO) { e.preventDefault(); return; }
    if (coUrl()) return;
    e.preventDefault();
    var u = window.prompt('第一次切換：請貼上共編版網址（script.google.com/macros/s/…/exec）。之後從共編版切回來時會自動記住。');
    if (u && /^https:\/\/script\.google\.com\//.test(u.trim())) { try { localStorage.setItem('coedit-url', u.trim()); } catch (er) {} location.href = u.trim(); }
  });
  $('msPersonal').addEventListener('click', function (e) { if (!IS_CO) e.preventDefault(); });
  function renderProjects() {
    renderSwitch();
    topbar('監造抽查紀錄', '室內裝修工程 · 表單版次 ' + D.version);
    var html = S.projects.map(function (p) {
      var n = 0, ng = 0;
      Object.keys(p.recs).forEach(function (wid) {
        (p.recs[wid] || []).forEach(function (r) { n++; if (WORK[wid] && stats(WORK[wid], r).ng) ng++; });
      });
      var mt = p.mats.length, md = p.mats.filter(function (m) { return m.inspResult; }).length;
      var meta = [p.info.contractor, p.info.startDate && '開工 ' + p.info.startDate, mt && '材料抽查 ' + md + '/' + mt].filter(Boolean).join(' · ') || '尚未填寫基本資料';
      return '<button type="button" class="row-btn" data-pid="' + p.id + '"><span class="rb-main"><span class="rb-title">' +
        esc(p.info.name || '未命名工程') + '</span><span class="rb-meta">' + esc(meta) + '</span></span>' +
        '<span class="rb-right">' + n + ' 份' + (ng ? '<span class="badge-ng">' + ng + ' 缺失</span>' : '') + '</span><span class="chev">›</span></button>';
    }).join('');
    $('projectList').innerHTML = html || '<div class="empty">還沒有工程。按「新增工程」建立第一個，填入基本資料後即可開始抽查。</div>';
    $('backupMsg').textContent = saveMsg;
    renderTrash('trashBoxP', S.trash); $('trashMsgP').textContent = '';
    storeInfo();
    show('vProjects');
  }
  $('projectList').addEventListener('click', function (e) {
    var b = e.target.closest('[data-pid]'); if (b) go({ v: 'project', pid: b.dataset.pid });
  });
  function lastPhrases() { var p = S.projects.filter(function (x) { return x.phrases && x.phrases.length; }).pop(); return p ? JSON.parse(JSON.stringify(p.phrases)) : null; }
  $('btnNewProject').addEventListener('click', function () {
    var p = { id: uid(), info: { durType: '日曆天', changeCount: '0', extendDays: '0' }, recs: {}, mats: [], created: Date.now() };
    var ph = lastPhrases(); if (ph) p.phrases = ph;
    S.projects.unshift(p); save(); go({ v: 'project', pid: p.id, focus: 1 });
  });
  $('btnExport').addEventListener('click', function () {
    $('backupMsg').textContent = '正在打包資料與照片…';
    IDB.all().then(function (imgs) {
      Object.keys(imgs).forEach(function (kk) { if (/^bak_/.test(kk)) delete imgs[kk]; });
      var n = Object.keys(imgs).length;
      download('監造抽查備份_' + today() + '.json', JSON.stringify(Object.assign({}, S, { _images: imgs })), 'application/json');
      $('backupMsg').textContent = '已匯出 ' + S.projects.length + ' 個工程' + (n ? '、' + n + ' 張照片' : '') + '。';
    }).catch(function () {
      download('監造抽查備份_' + today() + '.json', JSON.stringify(S), 'application/json');
      $('backupMsg').textContent = '已匯出 ' + S.projects.length + ' 個工程（照片無法讀取，未包含）。';
    });
  });
  function download(name, text, type) {
    var blob = new Blob([text], { type: type });
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
  }
  $('fileImport').addEventListener('change', function (e) {
    var f = e.target.files[0]; if (!f) return;
    var rd = new FileReader();
    rd.onload = function () {
      try {
        var data = JSON.parse(rd.result);
        if (!data || !Array.isArray(data.projects)) throw 0;
        var imgs = data._images || {}; delete data._images;
        Object.keys(imgs).forEach(function (kk) { IDB.put(kk, imgs[kk]); if (window.Cloud) window.Cloud.uploadPhoto(kk, imgs[kk]); });
        var added = 0, updated = 0;
        data.projects.forEach(function (p) {
          if (!p.mats) p.mats = [];
          var ex = proj(p.id);
          if (ex) { S.projects[S.projects.indexOf(ex)] = p; updated++; } else { S.projects.push(p); added++; }
        });
        save(); renderProjects();
        $('backupMsg').textContent = '匯入完成：新增 ' + added + ' 個、更新 ' + updated + ' 個工程。';
      } catch (err) { $('backupMsg').textContent = '這個檔案不是本系統的備份檔，請選擇「監造抽查備份_日期.json」。'; }
      e.target.value = '';
    };
    rd.readAsText(f);
  });

  /* ========== 2 工程 ========== */
  /* 本工程要抽查的工項：未設定時為全部；已停用但有紀錄的工項仍顯示，避免資料看不到 */
  function enabledWorks(p) {
    return orderedWorks(p).filter(function (w) { return !p.works || p.works.indexOf(w.id) >= 0 || (p.recs[w.id] || []).length; });
  }
  /* 工項排列順序：每個工程可自訂（長按拖曳），新加入的工項排在預設位置 */
  function orderedWorks(p) {
    if (!p.wOrder || !p.wOrder.length) return D.works.slice();
    var pos = {}; p.wOrder.forEach(function (id, i) { pos[id] = i; });
    var base = {}; D.works.forEach(function (w, i) { base[w.id] = i; });
    var rank = function (w) {
      if (pos[w.id] != null) return pos[w.id];
      for (var i = base[w.id] - 1; i >= 0; i--) if (pos[D.works[i].id] != null) return pos[D.works[i].id] + 0.5;   // 未排序的新工項接在原本前一項之後
      return -0.5;
    };
    return D.works.slice().sort(function (a, b) { return rank(a) - rank(b) || base[a.id] - base[b.id]; });
  }
  function renderProject() {
    var p = proj(cur.pid), I = p.info;
    topbar(I.name || '未命名工程', [I.code, I.contractor].filter(Boolean).join(' · '), '工程清單');
    var on = enabledWorks(p);
    $('workPick').innerHTML = orderedWorks(p).map(function (w) {
      return '<label class="chk"><input type="checkbox" data-wsel="' + w.id + '"' + (on.indexOf(w) >= 0 ? ' checked' : '') + '><span>' + esc(w.name) + '</span></label>';
    }).join('');
    $('workGrid').innerHTML = dwgCard(p) + on.map(function (w) {
      var list = p.recs[w.id] || [], ng = 0, meta;
      list.forEach(function (r) { if (stats(w, r).ng) ng++; });
      if (w.kind === 'phased') {
        var cnt = [0, 0, 0]; list.forEach(function (r) { cnt[r.phase]++; });
        meta = '前 ' + cnt[0] + ' · 中 ' + cnt[1] + ' · 完成 ' + cnt[2];
      } else meta = '共 ' + list.length + ' 份';
      return '<button type="button" class="work-card' + (ng ? ' has-ng' : '') + (w.kind !== 'phased' ? ' special' : '') + '" data-wid="' + w.id + '">' +
        '<span class="wc-code">' + w.code + '</span><span class="wc-name">' + esc(w.name) + '</span>' +
        '<span class="wc-meta">' + meta + '</span>' + (ng ? '<span class="wc-ng">' + ng + ' 份有缺失</span>' : '') + '</button>';
    }).join('');

    var c = {}; MSTAT_ORDER.forEach(function (s) { c[s] = 0; }); p.mats.forEach(function (m) { c[matStatus(m)]++; });
    var n = p.mats.length, done = c.ok + c.ng;
    $('matDir').innerHTML = n ?
      '<span class="md-big">' + done + '<small> / ' + n + '</small></span><span class="md-lbl">已完成進場抽查</span>' +
      barHtml(c, n) +
      '<span class="md-meta">待抽查 ' + c.arrived + ' · 待進場 ' + c.approved + ' · 審查中 ' + c.review + ' · 待送審 ' + c.todo + (c.ng ? ' · <b>不合格 ' + c.ng + '</b>' : '') + '</span>' +
      '<span class="md-go">進入材料管制 ›</span>'
      : '<span class="md-lbl">尚未匯入材料設備送審管制總表</span><span class="md-meta">上傳工程會格式的管制總表（PDF、Excel 或 CSV），系統會列出每項材料，追蹤送審、進場與抽查日期。</span><span class="md-go">開始建立 ›</span>';

    renderCal(p);
    $('infoForm').innerHTML = INFO_FIELDS.map(function (f) { return fieldHtml('i_', f, I[f.k]); }).join('');
    $('infoCard').open = !I.name || !!cur.focus;
    $('infoSaved').textContent = saveMsg || '已自動儲存';
    renderTrash('trashBox', p.trash); $('trashMsg').textContent = '';
    show('vProject');
    if (cur.focus) { delete cur.focus; try { history.replaceState(cur, ''); } catch (e) {} setTimeout(function () { $('i_name').focus(); }, 50); }
  }
  function barHtml(c, n) {
    if (!n) return '<span class="progress-bar"></span>';
    return '<span class="progress-bar">' + MSTAT_ORDER.map(function (s) {
      return c[s] ? '<span class="seg-' + s + '" style="width:' + (c[s] / n * 100).toFixed(2) + '%" title="' + MSTAT[s] + ' ' + c[s] + '"></span>' : '';
    }).join('') + '</span>';
  }
  $('workPick').addEventListener('change', function (e) {
    var id = e.target.dataset.wsel; if (!id) return;
    var p = proj(cur.pid), sel = p.works ? p.works.slice() : D.works.map(function (w) { return w.id; });
    if (e.target.checked) { if (sel.indexOf(id) < 0) sel.push(id); } else sel = sel.filter(function (x) { return x !== id; });
    p.works = D.works.map(function (w) { return w.id; }).filter(function (x) { return sel.indexOf(x) >= 0; });
    save(); var open = $('workPickBox').open; renderProject(); $('workPickBox').open = open;
  });
  $('btnWorkOrder').addEventListener('click', function () { var p = proj(cur.pid), sy = window.scrollY; delete p.wOrder; save(); renderProject(); window.scrollTo(0, sy); $('workPickBox').open = true; });
  /* 長按後拖曳排序（手機、電腦皆可）：工項卡片、照片共用 */
  function makeSortable(grid, sel, onEnd, handleSel) {
    var timer = null, st = null, justDragged = false;
    function cards() { return Array.prototype.slice.call(grid.querySelectorAll(sel)); }
    function cancel() { clearTimeout(timer); timer = null; if (st && !st.on) st = null; }
    grid.addEventListener('pointerdown', function (e) {
      var c = e.target.closest(sel); if (!c || e.button > 0 || !grid.contains(c)) return;
      if (handleSel && !e.target.closest(handleSel)) return;
      st = { c: c, x: e.clientX, y: e.clientY, id: e.pointerId, on: false };
      timer = setTimeout(begin, e.pointerType === 'mouse' ? 350 : 450);
    });
    function begin() {
      if (!st) return; st.on = true;
      var r = st.c.getBoundingClientRect();
      st.dx = st.x - r.left; st.dy = st.y - r.top;
      st.ghost = st.c.cloneNode(true); st.ghost.className += ' drag-ghost'; st.ghost.style.width = r.width + 'px'; st.ghost.style.height = r.height + 'px';
      document.body.appendChild(st.ghost); place(st.x, st.y);
      st.c.classList.add('drag-src'); grid.classList.add('sorting');
      try { grid.setPointerCapture(st.id); } catch (err) {}
      if (navigator.vibrate) try { navigator.vibrate(15); } catch (err) {}
    }
    function place(x, y) { st.ghost.style.left = (x - st.dx) + 'px'; st.ghost.style.top = (y - st.dy) + 'px'; }
    grid.addEventListener('pointermove', function (e) {
      if (!st) return;
      if (!st.on) { if (Math.abs(e.clientX - st.x) + Math.abs(e.clientY - st.y) > 8) cancel(); return; }
      e.preventDefault(); place(e.clientX, e.clientY);
      st.ghost.style.visibility = 'hidden';
      var el = document.elementFromPoint(e.clientX, e.clientY); st.ghost.style.visibility = '';
      var t = el && el.closest(sel);
      if (t && t !== st.c && grid.contains(t)) {
        var r = t.getBoundingClientRect(), list = cards(), after = list.indexOf(t) > list.indexOf(st.c);
        var before = (e.clientY < r.top + r.height / 2 && Math.abs(e.clientY - (r.top + r.height / 2)) > r.height / 4) || (Math.abs(e.clientY - (r.top + r.height / 2)) <= r.height / 4 && e.clientX < r.left + r.width / 2);
        if (before && !after) grid.insertBefore(st.c, t); else if (!before && after) grid.insertBefore(st.c, t.nextSibling);
      }
      var vh = window.innerHeight; if (e.clientY < 80) window.scrollBy(0, -12); else if (e.clientY > vh - 60) window.scrollBy(0, 12);
    });
    function end() {
      clearTimeout(timer); timer = null;
      if (!st) return; var s0 = st; st = null; if (!s0.on) return;
      s0.ghost.remove(); s0.c.classList.remove('drag-src'); grid.classList.remove('sorting');
      justDragged = true; setTimeout(function () { justDragged = false; }, 350);
      onEnd(cards());
    }
    grid.addEventListener('pointerup', end); grid.addEventListener('pointercancel', function () { if (st && st.on) end(); else cancel(); });
    grid.addEventListener('dragstart', function (e) { if (e.target.closest(sel)) e.preventDefault(); });
    grid.addEventListener('touchmove', function (e) { if (st && st.on) e.preventDefault(); }, { passive: false });
    grid.addEventListener('contextmenu', function (e) { if (e.target.closest(sel) && (!handleSel || e.target.closest(handleSel))) e.preventDefault(); });
    grid.addEventListener('click', function (e) { if (justDragged) { e.stopImmediatePropagation(); e.preventDefault(); } }, true);
  }
  makeSortable($('workGrid'), '[data-wid]', function (list) {
    var p = proj(cur.pid), shown = list.map(function (c) { return c.dataset.wid; });
    var all = orderedWorks(p).map(function (w) { return w.id; }), hidden = all.filter(function (id) { return shown.indexOf(id) < 0; });
    /* 隱藏的工項維持在原本相鄰位置之後 */
    var next = shown.slice();
    hidden.forEach(function (id) { var i = all.indexOf(id), prev = null; for (var k = i - 1; k >= 0; k--) if (next.indexOf(all[k]) >= 0) { prev = all[k]; break; } next.splice(prev ? next.indexOf(prev) + 1 : 0, 0, id); });
    p.wOrder = next; save(); var sy = window.scrollY; renderProject(); window.scrollTo(0, sy);
  });
  $('btnWorkAll').addEventListener('click', function () { var p = proj(cur.pid); delete p.works; save(); renderProject(); $('workPickBox').open = true; });
  $('workGrid').addEventListener('click', function (e) {
    if (e.target.closest('[data-dwgs]')) { go({ v: 'dwgs', pid: cur.pid }); return; }
    var b = e.target.closest('[data-wid]'); if (b) go({ v: 'work', pid: cur.pid, wid: b.dataset.wid });
  });
  $('matDir').addEventListener('click', function () { go({ v: 'materials', pid: cur.pid }); });
  function onInfo(e) {
    var id = e.target.id; if (!id || id.indexOf('i_') !== 0) return;
    var p = cur.v === 'project' && proj(cur.pid); if (!p) return;
    p.info[id.slice(2)] = e.target.value; save();
    $('infoSaved').textContent = saveMsg || '已自動儲存';
    if (/^i_(name|code|contractor)$/.test(id)) topbar(p.info.name || '未命名工程', [p.info.code, p.info.contractor].filter(Boolean).join(' · '), '工程清單');
  }
  $('infoForm').addEventListener('input', onInfo);
  $('infoForm').addEventListener('change', onInfo);
  $('infoForm').addEventListener('click', function (e) {
    if (e.target.id !== 'calcEnd') return;
    var I = proj(cur.pid).info, n = parseInt(I.duration, 10);
    if (!I.startDate || !n) { $('infoSaved').textContent = '請先填開工日期與契約工期'; return; }
    if (I.durType !== '日曆天') { $('infoSaved').textContent = '工作天需扣除假日，請手動填寫'; return; }
    I.planEnd = addDays(I.startDate, n - 1); $('i_planEnd').value = I.planEnd; save();
    $('infoSaved').textContent = '已推算（開工日算第 1 天）';
  });
  $('btnDelProject').addEventListener('click', function () {
    var p = proj(cur.pid), n = 0; Object.keys(p.recs).forEach(function (w) { n += p.recs[w].length; });
    ask('刪除「' + (p.info.name || '未命名工程') + '」？', '此工程與底下 ' + n + ' 份抽查紀錄、' + p.mats.length + ' 項材料會移到工程清單最下方的「最近刪除」，7 天內可以復原。', '取消', '刪除工程', function () {
      S.projects.splice(S.projects.indexOf(p), 1);
      (S.trash || (S.trash = [])).push({ id: uid(), type: 'proj', at: Date.now(), data: p }); save(); go({ v: 'projects' }, true);
    });
  });

  /* ========== 3 工項紀錄清單 ========== */
  function renderWork() {
    var p = proj(cur.pid), w = WORK[cur.wid], list = recs(p, w.id);
    if (batchWid !== w.id) batchSel = null;
    topbar(w.name, p.info.name || '未命名工程', p.info.name || '工程');
    var blocks = w.kind === 'phased' ? w.phases.map(function (ph, pi) { return { pi: pi, title: '(' + (pi + 1) + ') ' + flowName(pi), n: countItems(ph) + ' 項' }; })
      : [{ pi: 0, title: w.kind === 'photo' ? '督導紀錄' : '抽查紀錄', n: w.kind === 'safety' ? countItems(w.phases[0]) + ' 項' : w.kind === 'photo' ? '照片' : '項目自選' }];
    $('phaseBlocks').innerHTML = blocks.map(function (b) {
      var items = list.filter(function (r) { return r.phase === b.pi; }).sort(function (a, c) { return c.seq - a.seq; });
      var rows = items.map(function (r) {
        var s = stats(w, r);
        var meta = [r.checkDate && roc(r.checkDate, '.'), r.location].filter(Boolean).join(' · ') || '尚未填寫位置與日期';
        var ck = batchSel ? '<span class="bsel' + (batchSel[r.id] ? ' on' : '') + '" aria-hidden="true">' + (batchSel[r.id] ? '✓' : '') + '</span>' : '';
        return '<button type="button" class="rec' + (batchSel && batchSel[r.id] ? ' picked' : '') + '" data-rid="' + r.id + '">' + ck + '<span class="rb-main"><span class="rb-title">' + esc(r.docNo) +
          (s.ng ? '<span class="badge-ng">' + s.ng + ' 缺失</span>' : '') + '</span><span class="rb-meta">' + esc(meta) + '</span></span>' +
          '<span class="rb-right">' + s.d + '/' + s.t + '</span><span class="chev">›</span></button>';
      }).join('');
      return '<section class="phase-block"><div class="pb-head"><span class="pb-title">' + b.title +
        '<span class="pb-n"> · ' + b.n + '</span></span><button type="button" class="btn sm" data-new="' + b.pi + '">＋ 新增抽查</button></div>' +
        (rows || '<div class="pb-empty">尚無紀錄</div>') + '</section>';
    }).join('');
    var n = list.filter(function (r) { return !r.draft; }).length, ns = batchSel ? Object.keys(batchSel).length : 0;
    $('batchBar').hidden = !n;
    $('batchBar').innerHTML = batchSel ? '<span class="bb-n">已選 ' + ns + ' 張</span><button type="button" class="link" data-bb="all">全選</button><button type="button" class="link" data-bb="cancel">取消</button>' +
      '<button type="button" class="btn primary sm" data-bb="go"' + (ns ? '' : ' disabled') + '>預覽輸出</button>'
      : '<button type="button" class="link" data-bb="start">☑ 批次輸出 PDF</button>';
    $('workRefs').textContent = w.kind === 'photo' ? '每次督導新增一筆，加入現場照片並填寫說明後輸出 A4 照片頁（每頁 6 張）。' : w.kind === 'irregular' ? '不定期抽查依承商自主檢查紀錄表隨機抽樣，抽驗項目可從 11 個工項的抽查標準帶入。'
      : '抽查標準依據：' + w.sources.map(function (s) { return D.refs[s]; }).join('；') + '。數值請依本案契約圖說確認。';
    show('vWork');
  }
  function countItems(ph) { return ph.groups.reduce(function (a, g) { return a + g.items.length; }, 0); }
  $('phaseBlocks').addEventListener('click', function (e) {
    var p = proj(cur.pid), w = WORK[cur.wid];
    var nb = e.target.closest('[data-new]');
    if (nb) { var r = newRecord(p, w, +nb.dataset.new); go({ v: 'form', pid: p.id, wid: w.id, rid: r.id }); return; }
    var rb = e.target.closest('[data-rid]');
    if (rb && batchSel) { var id = rb.dataset.rid; if (batchSel[id]) delete batchSel[id]; else batchSel[id] = 1; renderWork(); return; }
    if (rb) go({ v: 'form', pid: p.id, wid: w.id, rid: rb.dataset.rid });
  });
  /* 批次輸出：勾選多張後一起預覽、列印或另存成一個 PDF */
  var batchSel = null, batchWid = null;
  $('batchBar').addEventListener('click', function (e) {
    var b = e.target.closest('[data-bb]'); if (!b) return;
    var p = proj(cur.pid), w = WORK[cur.wid], a = b.dataset.bb;
    if (a === 'start') { batchSel = {}; batchWid = w.id; }
    else if (a === 'cancel') batchSel = null;
    else if (a === 'all') recs(p, w.id).forEach(function (r) { if (!r.draft) batchSel[r.id] = 1; });
    else if (a === 'go') {
      var ids = recs(p, w.id).filter(function (r) { return batchSel[r.id]; }).sort(function (x, y) { return (x.checkDate || '').localeCompare(y.checkDate || '') || x.phase - y.phase || x.seq - y.seq; }).map(function (r) { return r.id; });
      if (!ids.length) return;
      go({ v: 'preview', pid: p.id, wid: w.id, rid: ids[0], batch: ids }); return;
    }
    renderWork();
  });

  /* ========== 4 填寫 ========== */
  function R() { return recById(proj(cur.pid), cur.wid, cur.rid); }
  function W() { return WORK[cur.wid]; }

  function renderForm() {
    var p = proj(cur.pid), w = W(), r = R(), I = p.info, single = !isFix(w), m = w.kind === 'matform' && matById(p, r.mid);
    topbar(recTitle(w, r), m ? m.name : (I.name || '未命名工程'), m ? '材料' : w.name);
    $('autoLine').innerHTML = '自動帶入：工程名稱「' + esc(I.name || '未填') + '」' + (m ? ' · 材料「' + esc(m.name) + '」' : single ? '' : ' · 分項「' + esc(w.name) + '」') + ' · 監造人員「' + esc(I.inspector || '未填') + '」';
    if (!r.draft && hasStd(w) && snapStale(w, r)) $('autoLine').innerHTML += '<div class="draft-note std-note">本表沿用存檔當時（' + esc(r.snap.ver) + '）的抽查標準，與目前版本不同。<button type="button" class="link" id="btnSnapUpdate">更新為最新標準</button></div>';
    var prevR = r.draft && hasStd(w) && w.kind !== 'matform' ? recs(p, w.id).filter(function (x) { return x !== r && !x.draft && x.phase === r.phase; }).sort(function (a, b) { return (b.created || 0) - (a.created || 0); })[0] : null;
    if (prevR) $('autoLine').innerHTML += '<div class="copy-prev"><button type="button" class="link" id="btnCopyPrev" data-rid="' + prevR.id + '">↺ 帶入上一張（' + esc(prevR.docNo) + '）的判定與實測值</button></div>';
    if (r.draft) $('autoLine').innerHTML += '<div class="draft-note">尚未存檔：按下方「確認並產生 A4 表單」後才會列入抽查紀錄；直接返回則不保留。</div>';
    var photoOnly = w.kind === 'photo';
    var canMove = w.kind !== 'matform';
    $('btnMove').hidden = !canMove; $('btnMoveTop').hidden = !canMove;
    $('tplRow').hidden = w.kind === 'irregular' || photoOnly; $('tplMsg').textContent = '';
    $('signCard').hidden = photoOnly;
    ['docNo', 'location', 'checkDate', 'fixDate', 'fixPerson'].forEach(function (id) { $(id).value = r[id] || ''; });
    $('fixPerson').placeholder = I.inspector || '';
    document.querySelectorAll('input[name="fix"]').forEach(function (x) { x.checked = x.value === r.fix; });
    $('sigInspectorLbl').textContent = (single ? '監造現場人員' : '監造人員') + (I.inspector ? '（' + I.inspector + '）' : '');
    $('phasedFix').hidden = single || photoOnly; $('singleResult').hidden = !single || photoOnly;
    $('irrPicker').hidden = w.kind !== 'irregular';
    if (photoOnly) { $('groups').innerHTML = ''; $('tbRight').textContent = (r.photos || []).length + ' 張'; }
    else {
      if (single) fillSingle(r); else fixVis();
      if (w.kind === 'irregular') { fillPickWork(); renderRows(); } else renderGroups();
    }
    $('phMsg').textContent = ''; renderPhotoGrid(); pinInfo();
    show('vForm');
    sizePad('sigInspector'); sizePad('sigFixer');
  }
  ['docNo', 'location', 'checkDate', 'fixDate', 'fixPerson'].forEach(function (id) {
    $(id).addEventListener('input', function () { R()[id] = $(id).value; save(); });
  });

  $('autoLine').addEventListener('click', function (e) {
    var cp = e.target.closest('#btnCopyPrev');
    if (cp) {
      var p = proj(cur.pid), w0 = W(), r0 = R(), src = recById(p, w0.id, cp.dataset.rid); if (!src) return;
      var doCopy = function () {
        r0.res = JSON.parse(JSON.stringify(src.res || {})); r0.extra = JSON.parse(JSON.stringify(src.extra || {}));
        if (!isFix(w0)) ['result', 'exec', 'method', 'methodNote'].forEach(function (k) { r0[k] = src[k] || ''; });
        save(); renderForm();
      };
      if (touched(r0) && Object.keys(r0.res).some(function (k) { return r0.res[k].r || r0.res[k].note; })) ask('覆蓋目前已填的內容？', '會以「' + src.docNo + '」的判定、實測值與抽查標準取代本表目前填寫的內容；檢查位置、日期、照片不受影響。', '取消', '帶入', doCopy);
      else doCopy();
      return;
    }
    if (!e.target.closest('#btnSnapUpdate')) return;
    var w = W(), r = R();
    ask('更新為最新抽查標準？', '已填的判定與實測值會依項目代碼保留；新版已刪除的項目不再顯示，新增的項目需要補填。', '取消', '更新', function () { r.snap = snapOf(w, r); save(); renderForm(); });
  });
  /* 固定清單（分階段工項、勞安衛生） */
  function renderGroups() {
    var w = W(), r = R();
    $('groups').innerHTML = effGroups(w, r).map(function (g) {
      var gi = g.gi;
      var items = g.items.map(function (it) {
        var id = it.custom ? 'x' + it.x : it.key, attr = it.custom ? 'data-x="' + it.x + '"' : 'data-key="' + it.key + '"';
        var name = it.custom ? '<input class="inline-in" data-f="name" placeholder="管理項目" value="' + esc(it.name) + '">' : '<div class="item-name">' + esc(it.name) + '</div>';
        var std = '<textarea class="std-in" data-f="std" rows="2" placeholder="抽查標準（定性定量）" aria-label="抽查標準">' + esc(it.std) + '</textarea>';
        var tools = '<div class="row-tools">' + (!it.custom && it.std !== it.def ? '<button type="button" class="link" data-reset="' + it.key + '">還原預設標準</button>' : '') +
          (it.custom ? '<button type="button" class="link" data-delx="' + it.x + '">刪除此項目</button>' : '<button type="button" class="link" data-off="' + it.key + '">移除此項</button>') + '</div>';
        return _itemHtml(attr, 'r' + id, name + std + tools, '', it.custom ? '' : D.refs[it.ref], it.c.r, it.c.note, it.hint, 'n' + id);
      }).join('');
      var rem = g.removed;
      var remHtml = rem.length ? '<div class="removed">已移除：' + rem.map(function (o) { return '<button type="button" class="link" data-restore="' + o.key + '">' + esc(o.name) + ' ↺</button>'; }).join('') + '</div>' : '';
      return '<details class="group" open><summary>' + esc(g.name) + '<span class="gcount" id="gc' + gi + '"></span></summary>' +
        '<div class="gtools"><button type="button" class="link" data-allok="' + gi + '">本組未判定的全部設為 ○</button></div>' + items +
        '<div class="gfoot"><button type="button" class="btn ghost sm" data-addx="' + g.gid + '">＋ 新增抽查項目</button>' + remHtml + '</div></details>';
    }).join('');
    updateCounts();
  }
  function itemHtml(attr, radioName, nameHtml, stdHtml, ref, val, note, hint, noteId, extra) {
    var b = ['ok', 'ng', 'na'].map(function (v) {
      var lab = { ok: '合格', ng: '缺失', na: '無此項' }[v];
      return '<label class="bubble ' + v + '"><input type="radio" name="' + radioName + '" value="' + v + '"' + (val === v ? ' checked' : '') +
        ' aria-label="' + lab + '"><span>' + SYM[v] + '</span><small>' + lab + '</small></label>';
    }).join('');
    return '<div class="item' + (val === 'ng' ? ' ng' : '') + '" ' + attr + '>' +
      '<div>' + nameHtml + stdHtml + (ref ? '<div class="item-ref">依據：' + esc(ref) + '</div>' : '') + '</div>' +
      '<div class="row"><div class="bubbles" role="radiogroup">' + b + '</div>' +
      '<div class="note"><textarea id="' + noteId + '" rows="1" placeholder="' + esc(hint || '實際抽查情形') + '" aria-label="實際抽查情形">' + esc(note) + '</textarea>' +
      '<div class="phrases">' + phrases().map(function (f, i) { return '<button type="button" class="chip" data-ph="' + i + '" title="長按可編輯定型文">' + esc(f.t) + '</button>'; }).join('') + '</div></div></div>' +
      (extra || '') + '</div>';
  }
  var plainName = function (s) { return '<div class="item-name">' + s + '</div>'; };
  var plainStd = function (s) { return s ? '<div class="item-std">' + s + '</div>' : ''; };
  // 包裝成相同外觀
  var _itemHtml = itemHtml;
  itemHtml = function (attr, rn, name, std, ref, val, note, hint, nid, extra) { return _itemHtml(attr, rn, plainName(name), plainStd(std), ref, val, note, hint, nid, extra); };

  function itemTarget(item) {
    var r = R();
    if (item.dataset.row != null) return r.rows[+item.dataset.row];
    if (item.dataset.x != null) { var a = item.dataset.x.split('|'); return r.extra[a[0]][+a[1]]; }
    return cell(r, item.dataset.key);
  }
  $('groups').addEventListener('change', function (e) {
    var item = e.target.closest('.item'); if (!item || e.target.type !== 'radio') return;
    var v = e.target.value; itemTarget(item).r = v;
    item.classList.toggle('ng', v === 'ng'); save(); updateCounts();
  });
  $('groups').addEventListener('input', function (e) {
    var item = e.target.closest('.item'); if (!item || e.target.type === 'radio') return;
    var obj = itemTarget(item), f = e.target.dataset.f || 'note', v = e.target.value;
    if (f === 'std' && item.dataset.key != null) {
      var def = ''; baseGroups(W(), R()).forEach(function (g) { g.items.forEach(function (it) { if (it.id === item.dataset.key) def = it.std; }); });
      if (v === def) delete obj.std; else obj.std = v;
    } else obj[f] = v;
    save();
  });
  /* 定型文：可自訂（長按或按右鍵任一定型文進入編輯）；存在本機，也會一併匯出到備份檔 */
  function phrases() { var p = cur.pid && proj(cur.pid); return (p && p.phrases && p.phrases.length) ? p.phrases : (S.phrases && S.phrases.length) ? S.phrases : (D.phrases || []); }
  var PH_SYM = { ok: '○', na: '／', ng: '╳' };
  function phToText(list) { return list.map(function (f) { return f.t + (f.r ? ' ' + PH_SYM[f.r] : ''); }).join('\n'); }
  function phFromText(txt) {
    return txt.split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean).map(function (l) {
      var m = /\s*(○|／|\/|╳|x|X)$/.exec(l), r = '';
      if (m) { r = { '○': 'ok', '／': 'na', '/': 'na', '╳': 'ng', x: 'ng', X: 'ng' }[m[1]]; l = l.slice(0, m.index).trim(); }
      return { t: l, r: r };
    }).filter(function (f) { return f.t; });
  }
  var phLong = false, phTimer = null;
  function openPhEdit() { $('phText').value = phToText(phrases()); $('phAll').checked = true; $('phDlg').hidden = false; $('phText').focus(); }
  $('groups').addEventListener('pointerdown', function (e) {
    if (!e.target.closest('[data-ph]')) return;
    clearTimeout(phTimer); phTimer = setTimeout(function () { phLong = true; openPhEdit(); }, 550);
  });
  ['pointerup', 'pointerleave', 'pointercancel'].forEach(function (ev) { $('groups').addEventListener(ev, function () { clearTimeout(phTimer); }); });
  $('groups').addEventListener('contextmenu', function (e) { if (e.target.closest('[data-ph]')) { e.preventDefault(); clearTimeout(phTimer); openPhEdit(); } });
  $('phCancel').addEventListener('click', function () { $('phDlg').hidden = true; });
  $('phReset').addEventListener('click', function () { $('phText').value = phToText(D.phrases || []); });
  $('phSave').addEventListener('click', function () {
    /* 定型文存在工程資料內，共編時同事共用；勾選「套用到所有工程」則每個工程同一套 */
    var list = phFromText($('phText').value), targets = $('phAll').checked ? S.projects : [proj(cur.pid)];
    targets.forEach(function (p) { if (!p) return; if (list.length) p.phrases = JSON.parse(JSON.stringify(list)); else delete p.phrases; });
    delete S.phrases;
    $('phDlg').hidden = true; save();
    if (cur.v === 'form') { var w = W(); if (w.kind === 'irregular') renderRows(); else renderGroups(); }
  });
  /* 定型文：填入實際抽查情形，並帶入對應的判定（若尚未判定） */
  $('groups').addEventListener('click', function (e) {
    var b = e.target.closest('[data-ph]'); if (!b) return;
    if (phLong) { phLong = false; e.stopPropagation(); return; }
    var item = b.closest('.item'), f = phrases()[+b.dataset.ph], ta = item.querySelector('.note textarea');
    var cur0 = ta.value.trim(), prev = phrases().filter(function (x) { return x.t === cur0; })[0];
    if (cur0.split('；').indexOf(f.t) >= 0) return;
    ta.value = !cur0 || (prev && prev.r && f.r) ? f.t : cur0 + '；' + f.t;
    ta.dispatchEvent(new Event('input', { bubbles: true }));
    if (f.r && !item.querySelector('input[type=radio]:checked')) {
      var rd = item.querySelector('input[type=radio][value="' + f.r + '"]');
      if (rd) { rd.checked = true; rd.dispatchEvent(new Event('change', { bubbles: true })); }
    }
    e.stopPropagation();
  }, true);
  $('groups').addEventListener('click', function (e) {
    var t = e.target.closest('button'); if (!t) return;
    var r = R(), ds = t.dataset;
    if (ds.allok != null) {
      effGroups(W(), r)[+ds.allok].items.forEach(function (it) { if (!it.c.r) { if (it.custom) it.c.r = 'ok'; else cell(r, it.key).r = 'ok'; } });
      save(); renderGroups(); return;
    }
    if (ds.delrow != null) { r.rows.splice(+ds.delrow, 1); save(); renderRows(); return; }
    if (ds.addx != null) { (r.extra[ds.addx] || (r.extra[ds.addx] = [])).push({ name: '', std: '', r: '', note: '' }); save(); renderGroups(); focusLast(ds.addx); return; }
    if (ds.delx != null) { var a = ds.delx.split('|'); r.extra[a[0]].splice(+a[1], 1); save(); renderGroups(); return; }
    if (ds.off != null) { cell(r, ds.off).off = true; save(); renderGroups(); return; }
    if (ds.restore != null) { delete cell(r, ds.restore).off; save(); renderGroups(); return; }
    if (ds.reset != null) { delete cell(r, ds.reset).std; save(); renderGroups(); return; }
  });
  function focusLast(gi) {
    var list = $('groups').querySelectorAll('[data-x^="' + gi + '|"] .inline-in'); var el = list[list.length - 1];
    if (el) { el.focus(); el.scrollIntoView({ block: 'center' }); }
  }
  /* 本工程預設標準 */
  $('btnTplSave').addEventListener('click', function () {
    var p = proj(cur.pid), w = W(), r = R(), cells = {}, extra = {};
    Object.keys(r.res).forEach(function (key) { var c = r.res[key]; if (c.std != null || c.off) { cells[key] = { r: '', note: '' }; if (c.std != null) cells[key].std = c.std; if (c.off) cells[key].off = true; } });
    Object.keys(r.extra || {}).forEach(function (gi) { extra[gi] = r.extra[gi].filter(function (x) { return x.name || x.std; }).map(function (x) { return { name: x.name, std: x.std, r: '', note: '' }; }); });
    p.tpl = p.tpl || {}; p.tpl[w.id] = p.tpl[w.id] || {}; p.tpl[w.id][r.phase] = { cells: cells, extra: extra, kv: 2 }; save();
    $('tplMsg').textContent = '已設為本工程預設。之後在本工程新增的「' + recTitle(w, r) + '」會套用這份抽查標準。';
  });
  $('btnTplClear').addEventListener('click', function () {
    var p = proj(cur.pid), w = W(), r = R();
    if (p.tpl && p.tpl[w.id]) delete p.tpl[w.id][r.phase];
    save(); $('tplMsg').textContent = '已恢復原始預設，之後新增的紀錄會使用原本的抽查標準（本表不變）。';
  });
  function updateCounts() {
    var w = W(), r = R(), s = stats(w, r);
    if (w.kind !== 'irregular') {
      effGroups(w, r).forEach(function (g) {
        var d = 0, ng = 0;
        g.items.forEach(function (it) { if (it.c.r) d++; if (it.c.r === 'ng') ng++; });
        var el = $('gc' + g.gi); if (el) el.textContent = (ng ? ng + ' 缺失 · ' : '') + d + '/' + g.items.length;
      });
    }
    $('tbRight').textContent = s.d + '/' + s.t;
    if (!isFix(w)) {
      $('resultHint').textContent = s.t ? (s.ng ? '抽驗項目有 ' + s.ng + ' 項缺失。' : (s.d === s.t ? '抽驗項目皆無缺失。' : '')) + (r.resultManual ? '' : '（結果依抽驗項目自動判定，可手動更改）') : '';
      if (!r.resultManual) {
        r.result = s.ng ? 'ng' : (s.t && s.d === s.t ? 'ok' : '');
        document.querySelectorAll('input[name="s_result"]').forEach(function (x) { x.checked = x.value === r.result; });
        $('ngBlock').hidden = r.result !== 'ng'; save();
      }
    }
  }

  /* 不定期抽查：自選項目 */
  function fillPickWork() {
    var en = enabledWorks(proj(cur.pid));
    $('pickWork').innerHTML = PHASED.filter(function (w) { return en.indexOf(w) >= 0; }).map(function (w) { return '<option value="' + w.id + '">' + esc(w.name) + '</option>'; }).join('');
    fillPickItem();
  }
  function fillPickItem() {
    var w = WORK[$('pickWork').value];
    $('pickItem').innerHTML = w.phases.map(function (ph, pi) {
      return '<optgroup label="' + ph.name + '">' + ph.groups.map(function (g, gi) {
        return g.items.map(function (it, ii) { return '<option value="' + pi + '|' + gi + '|' + ii + '">' + esc(g.name + '／' + it.name) + '</option>'; }).join('');
      }).join('') + '</optgroup>';
    }).join('');
  }
  $('pickWork').addEventListener('change', fillPickItem);
  $('btnPickAdd').addEventListener('click', function () {
    var w = WORK[$('pickWork').value], a = $('pickItem').value.split('|').map(Number);
    var it = w.phases[a[0]].groups[a[1]].items[a[2]];
    R().rows.push({ work: w.name, name: it.name, std: it.std, hint: it.hint, ref: it.ref, note: '', r: '' });
    save(); renderRows(true);
  });
  $('btnCustomAdd').addEventListener('click', function () {
    R().rows.push({ work: '', name: '', std: '', note: '', r: '', custom: true });
    save(); renderRows(true);
  });
  function renderRows(scrollLast) {
    var r = R();
    if (!r.rows.length) { $('groups').innerHTML = '<div class="empty">尚未加入抽驗項目。從上方選擇工項與項目後按「加入此項目」。</div>'; updateCounts(); return; }
    $('groups').innerHTML = '<section class="group"><div class="group-head">抽驗項目</div>' + r.rows.map(function (x, i) {
      var name = x.custom ? '<input class="inline-in" data-f="name" placeholder="項目名稱" value="' + esc(x.name) + '">' : esc((x.work ? x.work + '／' : '') + x.name);
      var std = x.custom ? '<textarea class="inline-in" data-f="std" rows="1" placeholder="抽查標準（選填）">' + esc(x.std) + '</textarea>' : esc(x.std);
      var del = '<div class="row-tools"><button type="button" class="link" data-delrow="' + i + '">移除此項目</button></div>';
      return _itemHtml('data-row="' + i + '"', 'row' + i, '<div class="item-name">' + name + '</div>', std ? '<div class="item-std">' + std + '</div>' : '', x.ref ? D.refs[x.ref] : '', x.r, x.note, x.hint, 'rn' + i, del);
    }).join('') + '</section>';
    updateCounts();
    if (scrollLast) { var items = $('groups').querySelectorAll('.item'); var last = items[items.length - 1]; if (last) last.scrollIntoView({ block: 'center' }); }
  }

  /* 不定期／勞安：結果與缺失處理 */
  var SR = { s_result: 'result', s_exec: 'exec', s_method: 'method', s_re: 're' };
  function fillSingle(r) {
    Object.keys(SR).forEach(function (n) { document.querySelectorAll('input[name="' + n + '"]').forEach(function (x) { x.checked = x.value === r[SR[n]]; }); });
    ['deadline', 'methodNote', 'reTime', 'remark'].forEach(function (f) { $('s_' + f).value = r[f] || ''; });
    $('ngBlock').hidden = r.result !== 'ng';
  }
  Object.keys(SR).forEach(function (n) {
    document.querySelectorAll('input[name="' + n + '"]').forEach(function (x) {
      x.addEventListener('change', function () { var r = R(); r[SR[n]] = x.value; if (n === 's_result') r.resultManual = true; save(); $('ngBlock').hidden = r.result !== 'ng'; });
    });
  });
  ['deadline', 'methodNote', 'reTime', 'remark'].forEach(function (f) {
    $('s_' + f).addEventListener('input', function () { R()[f] = $('s_' + f).value; save(); });
  });

  function fixVis() { var on = R().fix !== 'none'; $('fixFields').hidden = !on; $('fixSignBox').hidden = !on; }
  document.querySelectorAll('input[name="fix"]').forEach(function (x) {
    x.addEventListener('change', function () { R().fix = x.value; save(); fixVis(); sizePad('sigFixer'); });
  });

  /* 簽名 */
  function sizePad(id) {
    var c = $(id), w = c.clientWidth; if (!w) return;
    var dpr = window.devicePixelRatio || 1;
    c.width = w * dpr; c.height = 140 * dpr;
    var ctx = c.getContext('2d'); ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.lineWidth = 2.4; ctx.lineCap = 'round'; ctx.lineJoin = 'round'; ctx.strokeStyle = '#000';
    var src = R() && R().sig[id];
    if (src) { var img = new Image(); img.onload = function () { ctx.drawImage(img, 0, 0, w, 140); }; img.src = src; }
  }
  function initPad(id) {
    var c = $(id), drawing = false, last = null;
    function pt(e) { var b = c.getBoundingClientRect(); return { x: e.clientX - b.left, y: e.clientY - b.top }; }
    c.addEventListener('pointerdown', function (e) { drawing = true; last = pt(e); c.setPointerCapture(e.pointerId); });
    c.addEventListener('pointermove', function (e) {
      if (!drawing) return; var q = pt(e), ctx = c.getContext('2d');
      ctx.beginPath(); ctx.moveTo(last.x, last.y); ctx.lineTo(q.x, q.y); ctx.stroke(); last = q;
    });
    function end() {
      if (!drawing) return; drawing = false;
      var o = document.createElement('canvas'), Wd = 360, H = Math.round(360 * c.height / c.width);
      o.width = Wd; o.height = H; o.getContext('2d').drawImage(c, 0, 0, Wd, H);
      R().sig[id] = o.toDataURL('image/png'); save();
    }
    c.addEventListener('pointerup', end); c.addEventListener('pointercancel', end);
  }
  initPad('sigInspector'); initPad('sigFixer');
  document.querySelectorAll('[data-clear]').forEach(function (b) {
    b.addEventListener('click', function () { delete R().sig[b.dataset.clear]; save(); sizePad(b.dataset.clear); });
  });
  var rt; window.addEventListener('resize', function () { clearTimeout(rt); rt = setTimeout(function () { if (cur.v === 'form') { sizePad('sigInspector'); sizePad('sigFixer'); } }, 200); });

  $('btnDelRec').addEventListener('click', function () {
    var r = R();
    if (r.draft) { var dl = recs(proj(cur.pid), cur.wid); dl.splice(dl.indexOf(r), 1); save(); history.back(); return; }
    ask('刪除紀錄 ' + r.docNo + '？', '會移到工程頁最下方的「最近刪除」，7 天內可以復原。', '取消', '刪除紀錄', function () {
      var p = proj(cur.pid), list = recs(p, cur.wid); list.splice(list.indexOf(r), 1);
      toTrash(p, 'rec', r, { wid: cur.wid }); save();
      history.back();
    });
  });

  $('btnConfirm').addEventListener('click', function () {
    var w = W(), r = R(), miss = [];
    if (w.kind === 'photo') {
      if (!(r.photos || []).length) { $('phMsg').textContent = '請先加入照片。'; $('phMsg').scrollIntoView({ block: 'center' }); return; }
      delete r.draft; save();
      go({ v: 'preview', pid: cur.pid, wid: cur.wid, rid: cur.rid }); return;
    }
    var toPreview = function () {
      delete r.draft; if (hasStd(w) && !r.snap) r.snap = snapOf(w, r); save();
      if (!isFix(w) && !r.result) { r.result = stats(w, r).ng ? 'ng' : 'ok'; save(); }
      if (w.kind === 'matform') { var m = matById(proj(cur.pid), r.mid); if (m) { m.inspDate = r.checkDate || today(); m.inspResult = stats(w, r).ng ? 'ng' : 'ok'; save(); } }
      go({ v: 'preview', pid: cur.pid, wid: cur.wid, rid: cur.rid });
    };
    if (w.kind === 'irregular') {
      if (!r.rows.length) { ask('尚未加入抽驗項目', '請先加入至少一個抽驗項目。', '好', '回去填寫', function () {}); return; }
      r.rows.forEach(function (x) { if (!x.r) miss.push(x.name || '自訂項目'); });
      if (!miss.length) return toPreview();
      return ask('還有 ' + miss.length + ' 項未判定', miss.slice(0, 5).join('、') + '。', '回去填寫', '未判定的標為「／」並繼續', function () {
        r.rows.forEach(function (x) { if (!x.r) x.r = 'na'; }); save(); toPreview();
      });
    }
    effGroups(w, r).forEach(function (g) { g.items.forEach(function (it) { if (!it.c.r) miss.push(it.name || '自訂項目'); }); });
    if (!miss.length) return toPreview();
    ask('還有 ' + miss.length + ' 項未判定', miss.slice(0, 5).join('、') + (miss.length > 5 ? ' 等' : '') + '。', '回去填寫', '未判定的標為「／」並繼續', function () {
      effGroups(w, r).forEach(function (g) { g.items.forEach(function (it) { if (!it.c.r) { if (it.custom) it.c.r = 'na'; else cell(r, it.key).r = 'na'; } }); });
      save(); toPreview();
    });
  });

  /* ========== 預覽列印 ========== */
  /* 輸出內容：表單、照片、抽查位置圖可個別勾選（單張與批次共用） */
  var PARTS = { form: true, photo: true, dwg: true };
  function recPaper(p, w, r, first) {
    bindOut(r, 'outAt', function () { return r.checkDate ? r.checkDate + nowLocal().slice(10) : nowLocal(); });
    var out = '';
    if (w.kind === 'photo') { if (PARTS.photo || PARTS.form) out += photoPages(w, r).replace(/<\/section>$/, foot() + '</section>'); }
    else {
      if (PARTS.form) out += (first ? '' : '<section class="rec-sec">') + (w.kind === 'phased' ? paperPhased(p, w, r) : w.kind === 'matform' ? paperMat(p, w, r) : paperSingle(p, w, r)) + (first ? '' : '</section>');
      if (PARTS.photo) out += photoPages(w, r);
      if (PARTS.dwg) out += recDwgPages(p, w.id, r);
    }
    return out;
  }
  function renderPreview() {
    var p = proj(cur.pid), w = W(), r = R(), batch = (cur.batch || []).map(function (id) { return recById(p, w.id, id); }).filter(Boolean);
    if (!batch.length) batch = [r];
    var multi = batch.length > 1;
    topbar('列印預覽', multi ? w.name + ' · 共 ' + batch.length + ' 張' : recTitle(w, r) + ' ' + r.docNo, multi ? w.name : '返回修改');
    $('btnMovePv').hidden = w.kind === 'matform' || multi;
    $('pvParts').hidden = false;
    document.querySelectorAll('[data-part]').forEach(function (x) { x.checked = !!PARTS[x.dataset.part]; x.closest('label').hidden = w.kind === 'photo' && x.dataset.part !== 'photo'; });
    var paper = $('paper'); paper.className = 'paper';
    var html = batch.map(function (x, i) { return recPaper(p, w, x, i === 0); }).join('').replace(/^<section class="rec-sec">/, '<section>');
    paper.innerHTML = html || '<p class="pv-empty">所選的內容沒有可輸出的頁面（例如這幾張沒有照片或位置標記）。</p>';
    if (!multi) bindOut(r, 'outAt', function () { return r.checkDate ? r.checkDate + nowLocal().slice(10) : nowLocal(); });
    $('outAt').closest('.out-at').hidden = multi;
    fillImgs(paper);
    $('printHint').textContent = '紙張選 A4 直式、邊界「預設」。「另存 PDF」會直接下載檔案，不經過列印視窗。';
    show('vPreview');
    var dt = roc(r.checkDate || today(), '.'), mm = w.kind === 'matform' && matById(p, r.mid);
    var only = PARTS.form ? '' : PARTS.photo && !PARTS.dwg ? '_照片' : PARTS.dwg && !PARTS.photo ? '_位置圖' : '_照片及位置圖';
    if (w.kind === 'photo') only = '';
    if (multi) {
      var ds = batch.map(function (x) { return x.checkDate; }).filter(Boolean).sort(), d1 = roc(ds[0] || today(), '.'), d2 = roc(ds[ds.length - 1] || today(), '.');
      setDocName(w.name + '抽查_' + d1 + (d2 !== d1 ? '-' + d2 : '') + only);
    } else setDocName((w.kind === 'phased' ? w.name + PH_SHORT[r.phase] + '抽查_' + dt : w.kind === 'matform' ? '材料進場抽查_' + (mm ? mm.name : '') + '_' + dt : w.name + '_' + dt) + only);
  }
  $('btnPrint').addEventListener('click', function () { window.print(); });
  $('pvParts').addEventListener('change', function (e) { var x = e.target.closest('[data-part]'); if (!x) return; PARTS[x.dataset.part] = x.checked; renderPreview(); });
  /* 檔名：工項＋階段＋抽查_民國日期（列印另存 PDF 時瀏覽器也會用這個名稱） */
  var docName = '';
  function setDocName(n) { $('pdfMsg').textContent = ''; docName = String(n).replace(/[\\/:*?"<>|]+/g, '_').trim(); document.title = docName; $('pdfName').value = docName; }
  $('pdfName').addEventListener('input', function () { docName = $('pdfName').value.trim() || docName; document.title = docName; });
  /* 直接另存 PDF：每頁依 A4 繪製後組成 PDF 檔下載 */
  $('btnPdf').addEventListener('click', function () {
    var btn = $('btnPdf'), label = btn.textContent;
    btn.disabled = true; btn.textContent = '產生 PDF 中…'; $('pdfMsg').textContent = '';
    (window.html2canvas ? Promise.resolve() : loadScript(CDN.h2c)).then(function () {
      return window.jspdf ? null : loadScript(CDN.jspdf);
    }).then(makePdf).then(function (blob) {
      var a = document.createElement('a'), name = (docName || '抽查表') + '.pdf';
      a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 60000);
      $('pdfMsg').textContent = '已產生「' + name + '」。';
    }).catch(function (err) { $('pdfMsg').textContent = (err && err.message) || 'PDF 產生失敗，請改用「列印」再選「另存為 PDF」。'; })
      .then(function () { btn.disabled = false; btn.textContent = label; });
  });
  function makePdf() {
    var src = $('paper'), land = src.classList.contains('landscape');
    var Wmm = land ? 297 : 210, Hmm = land ? 210 : 297, mx = land ? 8 : 10, my = 8;
    var host = document.createElement('div'); host.className = 'pdf-host'; document.body.appendChild(host);
    /* 依頁拆開：表單本體一頁、照片頁與位置圖各一頁 */
    var pages = [], curPg = null;
    Array.prototype.slice.call(src.childNodes).forEach(function (n) {
      var brk = n.nodeType === 1 && (n.classList.contains('ph-page') || n.classList.contains('rec-sec'));
      if (brk || !curPg) { curPg = document.createElement('div'); curPg.className = src.className + ' pdf-page'; pages.push(curPg); }
      var c = n.cloneNode(true); if (brk) c.style.marginTop = '0'; curPg.appendChild(c);
      if (brk) curPg = null;
    });
    pages = pages.filter(function (pg) { return pg.textContent.trim() || pg.querySelector('img'); });
    pages.forEach(function (pg) { pg.style.width = Wmm + 'mm'; pg.style.padding = my + 'mm ' + mx + 'mm'; host.appendChild(pg); });
    var pdf = new window.jspdf.jsPDF({ orientation: land ? 'landscape' : 'portrait', unit: 'mm', format: 'a4', compress: true });
    var first = true;
    return pages.reduce(function (pr, pg) {
      return pr.then(function () {
        /* 可切頁的位置：表格每一列、標題段落的下緣（避免一列被切成兩半） */
        var top0 = pg.getBoundingClientRect().top, cssW = pg.getBoundingClientRect().width, cuts = [];
        Array.prototype.forEach.call(pg.querySelectorAll('tr, .sum-paper > *'), function (el) { cuts.push(el.getBoundingClientRect().bottom - top0); });
        return window.html2canvas(pg, { scale: 2, backgroundColor: '#ffffff', useCORS: true, logging: false }).then(function (cv) {
          var pxPerMm = cv.width / Wmm, pageH = Math.floor(Hmm * pxPerMm), padT = Math.round(my * pxPerMm), k = cv.width / cssW;
          var bk = cuts.map(function (c) { return Math.round(c * k); }).sort(function (a, b) { return a - b; });
          for (var y = 0, n = 0; y < cv.height - 2; n++) {
            var off = n ? padT : 0, room = pageH - off, h = Math.min(room, cv.height - y);
            if (y + h < cv.height - 2) {
              var best = 0; bk.forEach(function (b) { if (b > y + room * 0.5 && b <= y + room) best = b; });
              if (best) h = best - y + 1;
            }
            var part = document.createElement('canvas');
            part.width = cv.width; part.height = h; part.getContext('2d').drawImage(cv, 0, y, cv.width, h, 0, 0, cv.width, h);
            if (!first) pdf.addPage('a4', land ? 'landscape' : 'portrait'); first = false;
            pdf.addImage(part.toDataURL('image/jpeg', 0.92), 'JPEG', 0, off / pxPerMm, Wmm, h / pxPerMm);
            y += h;
          }
        });
      });
    }, Promise.resolve()).then(function () { host.remove(); return pdf.output('blob'); }, function (e) { host.remove(); throw e; });
  }
  function sigImg(r, id) { return r.sig[id] ? '<img alt="簽名" src="' + r.sig[id] + '">' : ''; }
  /* 表單產出時間：預設為現在，可在預覽工具列自訂，存在該筆紀錄（或工程的材料總表）上 */
  function nowLocal() { var d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 16); }
  function fmtOut(v) { var m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(v || ''); return m ? (+m[1]) + '/' + (+m[2]) + '/' + (+m[3]) + ' ' + m[4] + ':' + m[5] : ''; }
  var outTarget = null;   // { obj, key, def }
  function outValue() { return (outTarget && outTarget.obj[outTarget.key]) || (outTarget && outTarget.def ? outTarget.def() : nowLocal()); }
  function foot() { return '<div class="foot"><span></span><span class="foot-out">產出 ' + fmtOut(outValue()) + '</span></div>'; }
  function bindOut(obj, key, def) {
    outTarget = { obj: obj, key: key, def: def }; $('outAt').value = outValue();
    $('btnOutNow').textContent = def ? '改回檢查日期' : '改為現在';
  }
  function setOut(v) {
    if (!outTarget) return;
    if (v) { outTarget.obj[outTarget.key] = v; save(); }
    var f = document.querySelector('#paper .foot-out'); if (f) f.textContent = '產出 ' + fmtOut(outValue());
    $('outAt').value = outValue();
  }
  $('outAt').addEventListener('change', function () { if ($('outAt').value) setOut($('outAt').value); });
  $('btnOutNow').addEventListener('click', function () {
    if (outTarget && outTarget.def) { delete outTarget.obj[outTarget.key]; save(); setOut(''); } else setOut(nowLocal());
  });
  function box(on) { return on ? '■' : '□'; }

  /* 依基本格式：項目不足一頁時補上空白格，讓表格填滿版面 */
  function rowUnits(a, b, c) { return Math.max(1, Math.ceil((a || '').length / 9), Math.ceil((b || '').length / 20), Math.ceil((c || '').length / 14)); }
  function tableBody(groups, target) {
    var body = '', u = 0;
    groups.forEach(function (g) {
      if (!g.items.length) return;
      body += '<tr class="sub"><td colspan="4">' + esc(g.name) + '</td></tr>'; u += 1;
      g.items.forEach(function (it) {
        body += '<tr><td>' + esc(it.name) + '</td><td>' + esc(it.std) + '</td><td>' + esc(it.c.note) + '</td><td class="res' + (it.c.r === 'ng' ? ' ng' : '') + '">' + (SYM[it.c.r] || '') + '</td></tr>';
        u += rowUnits(it.name, it.std, it.c.note) + 0.25;
      });
    });
    while (u + 1.2 <= target) { body += '<tr class="blank"><td></td><td></td><td></td><td></td></tr>'; u += 1.2; }
    return body;
  }
  var PAD = { phased: 36, matform: 32, safety: 0, irregular: 12 };
  function paperPhased(p, w, r) {
    var I = p.info;
    var flow = PH_SHORT.map(function (_, i) { return '<span class="opt">' + box(i === r.phase) + flowName(i) + '</span>'; }).join('');
    var body = tableBody(effGroups(w, r), PAD.phased);
    var on = r.fix !== 'none';
    var fixTxt = box(r.fix === 'done') + '已完成改善（檢附改善前中後照片）<br>' + box(r.fix === 'track') + '未完成改善，填具「施工品質缺失處理改善暨追蹤紀錄表」進行追蹤改善';
    return '<h3>' + (w.printTitle ? esc(w.printTitle) : esc(w.name) + '施工抽查紀錄表(' + (r.phase + 1) + ')') + '</h3><p class="no">編號：' + esc(r.docNo) + '</p>' +
      headTable(I.name, w.name, r) +
      '<tr><td>施工流程</td><td colspan="3">' + flow + '</td></tr>' +
      '<tr><td>檢查結果</td><td colspan="3"><span class="opt">○檢查合格</span><span class="opt">╳有缺失需改正</span><span class="opt">／無此檢查項目</span></td></tr></table>' +
      '<table style="border-top:0"><colgroup><col style="width:20%"><col style="width:40%"><col style="width:29%"><col style="width:11%"></colgroup>' +
      '<thead><tr><th style="border-top:0">管理項目</th><th style="border-top:0">抽查標準（定性定量）</th><th style="border-top:0">實際抽查情形（敘述抽查值）</th><th style="border-top:0">抽查結果</th></tr></thead><tbody>' + body +
      '<tr><td colspan="4" style="line-height:1.7">缺失複查結果：<br>' + fixTxt + '<br>複查日期：' + (on && roc(r.fixDate) || '　　年　　月　　日') +
      '　　複查人員簽名：' + (on ? esc(r.fixPerson || I.inspector) + ' ' + sigImg(r, 'sigFixer') : '') + '</td></tr>' +
      '<tr><td colspan="4" class="notes"><ol>' + D.notes.map(function (n) { return '<li>' + esc(n) + '</li>'; }).join('') + '</ol></td></tr>' +
      '<tr><td colspan="4" class="sig">監造工地負責（授權）人/監造現場人員簽名：' + esc(I.inspector) + ' ' + sigImg(r, 'sigInspector') + '</td></tr></tbody></table>' + foot();
  }
  function paperMat(p, w, r) {
    var I = p.info, m = matById(p, r.mid) || { name: '', no: '', qty: '' }, on = r.fix !== 'none';
    var fixTxt = box(r.fix === 'done') + '已完成改善（檢附改善前中後照片）<br>' + box(r.fix === 'track') + '未完成改善，填具「施工品質缺失處理改善暨追蹤紀錄表」進行追蹤改善';
    return '<h3>材料進場抽查紀錄表</h3><p class="no">編號：' + esc(r.docNo) + '</p>' +
      '<table><colgroup><col style="width:18%"><col style="width:42%"><col style="width:15%"><col style="width:25%"></colgroup>' +
      '<tr><td>工程名稱</td><td colspan="3">' + esc(I.name) + '</td></tr>' +
      '<tr><td>材料／設備名稱</td><td colspan="3">' + esc(m.name) + '</td></tr>' +
      '<tr><td>契約詳細表項次</td><td>' + esc(m.no) + '</td><td>契約數量</td><td>' + esc(m.qty) + '</td></tr>' +
      '<tr><td>檢查位置</td><td>' + esc(r.location) + '</td><td>檢查日期</td><td>' + roc(r.checkDate, '.') + '</td></tr>' +
      '<tr><td>檢查結果</td><td colspan="3"><span class="opt">○檢查合格</span><span class="opt">╳有缺失需改正</span><span class="opt">／無此檢查項目</span></td></tr></table>' +
      '<table style="border-top:0"><colgroup><col style="width:20%"><col style="width:40%"><col style="width:29%"><col style="width:11%"></colgroup>' +
      '<thead><tr><th style="border-top:0">管理項目</th><th style="border-top:0">抽查標準（定性定量）</th><th style="border-top:0">實際抽查情形（敘述抽查值）</th><th style="border-top:0">抽查結果</th></tr></thead><tbody>' +
      tableBody(effGroups(w, r), PAD.matform) +
      '<tr><td colspan="4" style="line-height:1.7">缺失複查結果：<br>' + fixTxt + '<br>複查日期：' + (on && roc(r.fixDate) || '　　年　　月　　日') +
      '　　複查人員簽名：' + (on ? esc(r.fixPerson || I.inspector) + ' ' + sigImg(r, 'sigFixer') : '') + '</td></tr>' +
      '<tr><td colspan="4" class="notes"><ol>' + D.notes.map(function (n) { return '<li>' + esc(n) + '</li>'; }).join('') + '</ol></td></tr>' +
      '<tr><td colspan="4" class="sig">監造工地負責（授權）人/監造現場人員簽名：' + esc(I.inspector) + ' ' + sigImg(r, 'sigInspector') + '</td></tr></tbody></table>' + foot();
  }
  function headTable(projName, subName, r) {
    return '<table><colgroup><col style="width:18%"><col style="width:42%"><col style="width:15%"><col style="width:25%"></colgroup>' +
      '<tr><td>工程名稱</td><td colspan="3">' + esc(projName) + '</td></tr>' +
      (subName ? '<tr><td>分項工程名稱</td><td colspan="3">' + esc(subName) + '</td></tr>' : '') +
      '<tr><td>檢查位置</td><td>' + esc(r.location) + '</td><td>檢查日期</td><td>' + roc(r.checkDate, '.') + '</td></tr>';
  }
  function paperSingle(p, w, r) {
    var I = p.info, isIrr = w.kind === 'irregular';
    var title = isIrr ? '不定期抽查紀錄' : '勞安衛生抽查紀錄';
    var notes = isIrr ? D.irregularNotes : D.safetyNotes;
    var inner = '<table class="inner"><colgroup><col style="width:19%"><col style="width:46%"><col style="width:26%"><col style="width:9%"></colgroup>' +
      '<tr><th>抽驗項目</th><th>抽查標準</th><th>抽查情形</th><th>結果</th></tr>';
    if (isIrr) {
      var u = 0;
      r.rows.forEach(function (x) {
        inner += '<tr><td>' + esc((x.work ? x.work + '／' : '') + x.name) + '</td><td>' + esc(x.std) + '</td><td>' + esc(x.note) + '</td><td class="res' + (x.r === 'ng' ? ' ng' : '') + '">' + (SYM[x.r] || '') + '</td></tr>';
        u += rowUnits(x.name, x.std, x.note) + 0.25;
      });
      while (u + 1.2 <= PAD.irregular) { inner += '<tr class="blank"><td></td><td></td><td></td><td></td></tr>'; u += 1.2; }
    } else {
      inner += tableBody(effGroups(w, r), PAD.safety);
    }
    inner += '</table>';
    var ng = r.result === 'ng';
    return '<h3>' + title + '</h3><p class="no">編號：' + esc(r.docNo) + '</p>' +
      headTable(I.name, '', r) + '</table>' +
      '<table style="border-top:0" class="single"><tbody>' +
      '<tr><td>一、抽驗說明：<br>' + notes.map(function (n, i) { return '(' + '一二三四'[i] + ')' + esc(n); }).join('<br>') + '</td></tr>' +
      '<tr><td>二、抽驗項目：' + inner +
      '<div class="res-line">' + box(r.result === 'ok') + '合格</div>' +
      '<div class="res-line">' + box(ng) + '不合格　執行 ' + box(ng && r.exec === 'improve') + '改善　' + box(ng && r.exec === 'redo') + '拆除重作　改善期限：' + (ng && r.deadline ? roc(r.deadline) : '') + '</div></td></tr>' +
      '<tr><td>二、缺失處理方法：　' + box(ng && r.method === 'repair') + '修補改善　' + box(ng && r.method === 'redo') + '拆除重作' +
      (ng && r.methodNote ? '<br>' + esc(r.methodNote) : '') + '</td></tr>' +
      '<tr><td style="line-height:1.7">三、複驗結果：<br>' + box(r.re === 'ok') + '合格<br>' + box(r.re === 'ng') + '不合格，填具「施工不合管制總表」進行追蹤改善<br>複查時間：' + (r.reTime ? roc(r.reTime) : '') + '</td></tr>' +
      '<tr><td class="remark">備註：' + esc(r.remark) + '</td></tr></tbody></table>' +
      '<p class="signline">監造現場人員：' + esc(I.inspector) + ' ' + sigImg(r, 'sigInspector') + '</p>' + foot();
  }

  /* ========== 5 材料抽查管制 ========== */
  var matFilter = 'all', flash = '';
  function renderMaterials() {
    var p = proj(cur.pid);
    topbar('材料抽查管制', p.info.name || '未命名工程', p.info.name || '工程');
    var c = {}; MSTAT_ORDER.forEach(function (s) { c[s] = 0; }); p.mats.forEach(function (m) { c[matStatus(m)]++; });
    var n = p.mats.length;
    $('matCount').textContent = n ? '抽查完成 ' + (c.ok + c.ng) + ' / ' + n : '尚無資料';
    $('matBar').outerHTML = barHtml(c, n).replace('<span class="progress-bar">', '<div class="progress-bar" id="matBar" aria-hidden="true">').replace(/<\/span>$/, '</div>');
    $('matLegend').innerHTML = MSTAT_ORDER.map(function (s) { return '<span class="lg"><i class="seg-' + s + '"></i>' + MSTAT[s] + ' ' + c[s] + '</span>'; }).join('');
    var F = [['all', '全部 ' + n], ['todo', '待送審 ' + c.todo], ['review', '審查中 ' + c.review], ['approved', '待進場 ' + c.approved], ['arrived', '待抽查 ' + c.arrived], ['done', '已抽查 ' + (c.ok + c.ng)], ['ng', '不合格 ' + c.ng]];
    $('matFilter').innerHTML = F.map(function (f) { return '<label class="pill"><input type="radio" name="mf" value="' + f[0] + '"' + (matFilter === f[0] ? ' checked' : '') + '><span>' + f[1] + '</span></label>'; }).join('');
    var list = p.mats.filter(function (m) {
      var s = matStatus(m); return matFilter === 'all' || s === matFilter || (matFilter === 'done' && (s === 'ok' || s === 'ng'));
    });
    $('matList').innerHTML = n ? (list.map(function (m) {
      var s = matStatus(m);
      var dates = [m.arrivalDate && '進場 ' + roc(m.arrivalDate, '.'), m.inspDate && '抽查 ' + roc(m.inspDate, '.')].filter(Boolean).join(' · ');
      return '<button type="button" class="row-btn mat-row" data-mid="' + m.id + '"><span class="rb-main"><span class="rb-meta">' + esc(m.no) + '</span>' +
        '<span class="rb-title">' + esc(m.name || '未命名材料') + '</span><span class="rb-meta">' + esc([m.qty, dates].filter(Boolean).join(' · ')) + '</span></span>' +
        '<span class="st st-' + s + '">' + MSTAT[s] + '</span><span class="chev">›</span></button>';
    }).join('') || '<div class="empty">此篩選沒有材料。</div>')
      : '<div class="empty">尚未建立材料清單。按「匯入材料管制表」上傳工程會格式的材料設備送審管制總表（PDF、Excel 或 CSV），或按「新增材料」逐筆建立。</div>';
    $('matMsg').textContent = flash; flash = '';
    show('vMaterials');
  }
  $('matFilter').addEventListener('change', function (e) { matFilter = e.target.value; renderMaterials(); });
  $('matList').addEventListener('click', function (e) { var b = e.target.closest('[data-mid]'); if (b) go({ v: 'matItem', pid: cur.pid, mid: b.dataset.mid }); });
  $('btnMatNew').addEventListener('click', function () {
    var p = proj(cur.pid), m = blankMat({}); p.mats.push(m); save(); go({ v: 'matItem', pid: p.id, mid: m.id, isNew: 1 });
  });
  function blankMat(o) {
    return {
      id: uid(), no: o.no || '', name: o.name || '', qty: o.qty || '', sampleTest: o.sampleTest || '', factory: o.factory || '',
      docs: o.docs || { vendor: false, catalog: false, test: false, sample: false, other: '' },
      planSubmit: '', actualSubmit: '', factoryDate: '', reviewDate: '', reviewResult: '', archiveNo: '', arrivalDate: o.arrivalDate || '',
      testDate: '', testResult: '', inspDate: o.inspDate || '', inspResult: '', inspNote: '', remark: o.remark || ''
    };
  }

  /* 材料明細 */
  function M() { return matById(proj(cur.pid), cur.mid); }
  function renderMatItem() {
    var p = proj(cur.pid), m = M();
    topbar(m.name || '新增材料', m.no || p.info.name, '材料清單');
    $('matBasic').innerHTML = MAT_BASIC.map(function (f) { return fieldHtml('m_', f, m[f.k]); }).join('');
    $('matDocs').innerHTML = DOCS.map(function (d) {
      return '<label class="chk"><input type="checkbox" id="md_' + d[0] + '"' + (m.docs[d[0]] ? ' checked' : '') + '><span>' + d[1] + '</span></label>';
    }).join('') + '<label class="f"><span>其他</span><input id="md_other" value="' + esc(m.docs.other === 'ˇ' ? 'ˇ' : m.docs.other) + '" placeholder="例：製造圖"></label>';
    $('matTrack').innerHTML = MAT_TRACK.map(function (f) { return fieldHtml('m_', f, m[f.k]); }).join('');
    $('m_inspDate').value = m.inspDate || ''; $('m_inspNote').value = m.inspNote || ''; $('m_remark').value = m.remark || '';
    document.querySelectorAll('input[name="m_insp"]').forEach(function (x) { x.checked = x.value === (m.inspResult || ''); });
    var mrecs = recs(p, 'material').filter(function (r) { return r.mid === m.id; }).sort(function (a, b) { return b.seq - a.seq; });
    $('matRecBlock').innerHTML = '<div class="pb-head"><span class="pb-title">材料進場抽查紀錄表<span class="pb-n"> · ' + countItems(D.matWork.phases[0]) + ' 項，標準可自行修改</span></span>' +
      '<button type="button" class="btn sm" id="btnMatRecNew">＋ 新增抽查表</button></div>' +
      (mrecs.map(function (r) {
        var s = stats(D.matWork, r);
        return '<button type="button" class="rec" data-mrid="' + r.id + '"><span class="rb-main"><span class="rb-title">' + esc(r.docNo) + (s.ng ? '<span class="badge-ng">' + s.ng + ' 缺失</span>' : '') + '</span>' +
          '<span class="rb-meta">' + esc([r.checkDate && roc(r.checkDate, '.'), r.location].filter(Boolean).join(' · ')) + '</span></span><span class="rb-right">' + s.d + '/' + s.t + '</span><span class="chev">›</span></button>';
      }).join('') || '<div class="pb-empty">尚無抽查表。進場抽查時按「新增抽查表」，完成後會自動記錄抽查日期與結果。</div>');
    show('vMatItem');
    if (cur.isNew) setTimeout(function () { $('m_no').focus(); }, 50);
  }
  function onMat(e) {
    var t = e.target, m = cur.v === 'matItem' && M(); if (!m) return;
    if (t.name === 'm_insp') { if (e.type !== 'change') return; m.inspResult = t.value; if (t.value && !m.inspDate) { m.inspDate = today(); $('m_inspDate').value = m.inspDate; } save(); return; }
    if (!t.id) return;
    if (t.id.indexOf('md_') === 0) { var kk = t.id.slice(3); m.docs[kk] = t.type === 'checkbox' ? t.checked : t.value; }
    else if (t.id.indexOf('m_') === 0) m[t.id.slice(2)] = t.value;
    save();
  }
  $('vMatItem').addEventListener('input', onMat);
  $('vMatItem').addEventListener('change', onMat);
  $('matRecBlock').addEventListener('click', function (e) {
    var p = proj(cur.pid), m = M();
    if (e.target.closest('#btnMatRecNew')) {
      var r = newRecord(p, D.matWork, 0, { mid: m.id, location: '', checkDate: m.arrivalDate || today() });
      go({ v: 'form', pid: p.id, wid: 'material', rid: r.id }); return;
    }
    var b = e.target.closest('[data-mrid]'); if (b) go({ v: 'form', pid: p.id, wid: 'material', rid: b.dataset.mrid });
  });
  $('btnInspToday').addEventListener('click', function () { var m = M(); m.inspDate = today(); $('m_inspDate').value = m.inspDate; save(); });
  $('btnMatDone').addEventListener('click', function () { history.back(); });
  $('btnMatDel').addEventListener('click', function () {
    var p = proj(cur.pid), m = M();
    ask('刪除「' + (m.name || '未命名材料') + '」？', '會移到工程頁最下方的「最近刪除」，7 天內可以復原。', '取消', '刪除材料', function () {
      p.mats.splice(p.mats.indexOf(m), 1); toTrash(p, 'mat', m); save(); go({ v: 'materials', pid: p.id }, true);
    });
  });

  /* 匯入材料管制表 */
  var pendingImport = null;
  $('matFile').addEventListener('change', function (e) {
    var f = e.target.files[0]; e.target.value = ''; if (!f) return;
    var ext = (f.name.split('.').pop() || '').toLowerCase();
    $('matMsg').textContent = '正在讀取「' + f.name + '」…';
    var rd = new FileReader();
    rd.onload = function () {
      var job;
      if (ext === 'csv') job = Promise.resolve(MatParse.parseRows(parseCsv(new TextDecoder('utf-8').decode(rd.result))));
      else if (ext === 'xlsx' || ext === 'xls') job = (window.XLSX ? Promise.resolve() : loadScript(CDN.xlsx)).then(function () {
        var wb = XLSX.read(new Uint8Array(rd.result), { type: 'array' }), all = [];
        wb.SheetNames.forEach(function (n) { all = all.concat(MatParse.parseRows(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '' }))); });
        return all;
      });
      else if (ext === 'pdf') job = pdfToPages(rd.result).then(MatParse.parsePages);
      else job = Promise.reject(new Error('請選擇 PDF、Excel（.xlsx）或 CSV 檔'));
      job.then(function (items) {
        if (!items.length) throw new Error('沒有讀到材料項目。請確認檔案是「材料設備送審管制總表」格式；掃描檔（圖片）無法讀取，請改用 Excel 或 CSV。');
        pendingImport = { file: f.name, items: items }; $('matMsg').textContent = '';
        go({ v: 'matImport', pid: cur.pid });
      }).catch(function (err) { $('matMsg').textContent = err.message || '讀取失敗'; });
    };
    rd.readAsArrayBuffer(f);
  });
  function pdfToPages(buf) {
    return (window.pdfjsLib ? Promise.resolve() : loadScript(CDN.pdf)).then(function () {
      pdfjsLib.GlobalWorkerOptions.workerSrc = CDN.pdfWorker;
      return pdfjsLib.getDocument({ data: new Uint8Array(buf), cMapUrl: CDN.pdfCmaps, cMapPacked: true }).promise;
    }).then(function (pdf) {
      var jobs = [];
      for (var i = 1; i <= pdf.numPages; i++) jobs.push(pdf.getPage(i).then(function (pg) {
        var vp = pg.getViewport({ scale: 1 });
        return pg.getTextContent().then(function (tc) {
          return {
            items: tc.items.filter(function (it) { return typeof it.str === 'string'; }).map(function (it) {
              var pt = vp.convertToViewportPoint(it.transform[4], it.transform[5]);
              var h = Math.hypot(it.transform[2], it.transform[3]) || it.height || 10;
              return { str: it.str, x: pt[0], y: pt[1] - h, w: it.width };
            })
          };
        });
      }));
      return Promise.all(jobs);
    });
  }
  function parseCsv(text) {
    var rows = [], row = [], f = '', q = false;
    text = text.replace(/^﻿/, '');
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += ch; }
      else if (ch === '"') q = true;
      else if (ch === ',') { row.push(f); f = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(f); rows.push(row); row = []; f = ''; }
      else f += ch;
    }
    if (f || row.length) { row.push(f); rows.push(row); }
    return rows;
  }
  function renderMatImport() {
    var p = proj(cur.pid), items = pendingImport.items;
    topbar('匯入材料管制表', pendingImport.file, '取消');
    var existing = p.mats.length;
    $('impInfo').textContent = '從「' + pendingImport.file + '」讀到 ' + items.length + ' 項材料。' + (existing ? '目前已有 ' + existing + ' 項。' : '') + '請核對名稱與數量，有誤可匯入後在明細中修改。';
    $('impTable').innerHTML = '<thead><tr><th>項次</th><th>材料／設備名稱</th><th>契約數量</th><th>取樣試驗</th><th>驗廠</th><th>送審資料</th></tr></thead><tbody>' +
      items.map(function (m) {
        var d = DOCS.filter(function (x) { return m.docs[x[0]]; }).map(function (x) { return x[1].replace('相關', '').replace('協力', ''); });
        if (m.docs.other) d.push(m.docs.other === 'ˇ' ? '其他' : m.docs.other);
        return '<tr><td>' + esc(m.no) + '</td><td>' + esc(m.name) + '</td><td>' + esc(m.qty) + '</td><td>' + esc(m.sampleTest) + '</td><td>' + esc(m.factory) + '</td><td>' + esc(d.join('、')) + '</td></tr>';
      }).join('') + '</tbody>';
    $('btnImpGo').textContent = '匯入 ' + items.length + ' 項';
    show('vMatImport');
  }
  $('btnImpCancel').addEventListener('click', function () { pendingImport = null; history.back(); });
  $('btnImpGo').addEventListener('click', function () {
    var p = proj(cur.pid), mode = document.querySelector('input[name="impMode"]:checked').value, key = function (m) { return (m.no + '|' + m.name).replace(/\s+/g, ''); };
    var old = {}; p.mats.forEach(function (m) { old[key(m)] = m; });
    var added = 0, kept = 0;
    var next = pendingImport.items.map(function (x) {
      var ex = mode === 'merge' && old[key(x)];
      if (ex) { kept++; ex.qty = x.qty || ex.qty; ex.sampleTest = x.sampleTest || ex.sampleTest; ex.factory = x.factory || ex.factory; ex.docs = x.docs; delete old[key(x)]; return ex; }
      added++; return blankMat(x);
    });
    if (mode === 'merge') Object.keys(old).forEach(function (kk) { next.push(old[kk]); });
    p.mats = next; save(); pendingImport = null;
    flash = '已匯入：新增 ' + added + ' 項' + (kept ? '、更新 ' + kept + ' 項（保留原有日期與抽查結果）' : '') + '。';
    history.back();
  });

  /* 匯出 CSV（可再匯入） */
  $('btnMatCsv').addEventListener('click', function () {
    var p = proj(cur.pid), q = function (v) { v = String(v == null ? '' : v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
    var head = ['契約詳細表項次', '材料/設備名稱', '契約數量', '是否取樣試驗', '是否驗廠', '協力廠商資料', '型錄', '相關試驗報告', '樣品色票', '其他', '預定送審日期', '實際送審日期', '審查日期', '審查結果', '進場日期', '抽查日期', '抽查結果', '抽查情形', '備註'];
    var lines = [head.join(',')].concat(p.mats.map(function (m) {
      return [m.no, m.name, m.qty, m.sampleTest, m.factory, m.docs.vendor ? 'ˇ' : '', m.docs.catalog ? 'ˇ' : '', m.docs.test ? 'ˇ' : '', m.docs.sample ? 'ˇ' : '', m.docs.other,
        m.planSubmit, m.actualSubmit, m.reviewDate, m.reviewResult, m.arrivalDate, m.inspDate, { ok: '合格', ng: '不合格' }[m.inspResult] || '', m.inspNote, m.remark].map(q).join(',');
    }));
    download('材料抽查管制_' + (p.info.name || '工程') + '_' + today() + '.csv', '﻿' + lines.join('\r\n'), 'text/csv');
  });

  /* 列印管制總表（A4 橫式） */
  $('btnMatPrint').addEventListener('click', function () { go({ v: 'matPrint', pid: cur.pid }); });
  function renderMatPrint() {
    var p = proj(cur.pid), I = p.info;
    topbar('列印預覽', '材料設備送審管制總表', '材料清單');
    var d = function (v) { return v ? roc(v, '.') : ''; }, v = function (b) { return b ? 'ˇ' : ''; };
    var rows = p.mats.map(function (m, i) {
      return '<tr><td class="c">' + (i + 1) + '</td><td>' + esc(m.no) + '<br>' + esc(m.name) + '</td><td class="c">' + esc(m.qty) + '</td><td class="c">' + esc(m.sampleTest) + '</td>' +
        '<td class="c">' + d(m.planSubmit) + '<br>' + d(m.actualSubmit) + '</td><td class="c">' + esc(m.factory) + '<br>' + d(m.factoryDate) + '</td>' +
        '<td class="c">' + v(m.docs.vendor) + '</td><td class="c">' + v(m.docs.catalog) + '</td><td class="c">' + v(m.docs.test) + '</td><td class="c">' + v(m.docs.sample) + '</td><td class="c">' + esc(m.docs.other) + '</td>' +
        '<td class="c">' + d(m.reviewDate) + '<br>' + esc(m.reviewResult) + '</td><td class="c">' + d(m.arrivalDate) + '</td>' +
        '<td class="c">' + d(m.inspDate) + '<br><b>' + ({ ok: '○', ng: '╳' }[m.inspResult] || '') + '</b></td><td>' + esc([m.archiveNo, m.remark].filter(Boolean).join(' ')) + '</td></tr>';
    }).join('');
    bindOut(p, 'matOutAt'); $('pvParts').hidden = true; $('btnMovePv').hidden = true; $('outAt').closest('.out-at').hidden = false;
    var paper = $('paper'); paper.className = 'paper landscape';
    paper.innerHTML = '<h3>材料設備送審管制總表</h3><p class="no">工程名稱：' + esc(I.name) + '</p>' +
      '<table class="mat"><colgroup><col style="width:3.5%"><col style="width:17%"><col style="width:8%"><col style="width:5%"><col style="width:8%"><col style="width:7%">' +
      '<col style="width:4.5%"><col style="width:4.5%"><col style="width:4.5%"><col style="width:4.5%"><col style="width:5%"><col style="width:8%"><col style="width:6.5%"><col style="width:7%"><col style="width:7.5%"></colgroup>' +
      '<thead><tr><th rowspan="2">項次</th><th rowspan="2">契約詳細表項次<br>材料/設備名稱</th><th rowspan="2">契約數量</th><th rowspan="2">是否取樣試驗</th><th rowspan="2">預定送審日期<br>實際送審日期</th><th rowspan="2">是否驗廠<br>驗廠日期</th>' +
      '<th colspan="5">送審資料（ˇ）</th><th rowspan="2">審查日期<br>審查結果</th><th rowspan="2">進場日期</th><th rowspan="2">抽查日期<br>結果</th><th rowspan="2">備註<br>(歸檔編號)</th></tr>' +
      '<tr><th>協力廠商資料</th><th>型錄</th><th>相關試驗報告</th><th>樣品色票</th><th>其他</th></tr></thead><tbody>' + rows + '</tbody></table>' +
      '<p class="mat-note">註：本表單於開工後應請廠商檢討提出預定送審及預定進場日期，並由監造單位會同廠商定期檢討辦理情形。</p>' + foot();
    $('printHint').textContent = '紙張選 A4 橫向、邊界「預設」。「另存 PDF」會直接下載檔案，不經過列印視窗。';
    show('vPreview');
    setDocName('材料設備送審管制總表_' + (I.name || '') + '_' + roc(today(), '.'));
  }

  /* ========== 抽查總表：截至某日的材料進場抽驗、分項施工抽查、缺失改善追蹤（格式依月進度會議報告） ========== */
  var sumOut = {};
  $('btnSum').addEventListener('click', function () {
    var to = today(); go({ v: 'sumPrint', pid: cur.pid, to: to, from: sumFromDef(to) });
  });
  function sumFromDef(to) { var d = new Date(to + 'T00:00:00'); d.setMonth(d.getMonth() - 1); d.setDate(d.getDate() + 1); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); }
  function sumDate(e) {
    var v = e.target.value; if (!v) return;
    cur[e.target.id === 'sumFrom' ? 'from' : 'to'] = v;
    if (cur.from > cur.to) { if (e.target.id === 'sumTo') cur.from = sumFromDef(cur.to); else cur.to = cur.from; }
    go(cur, true);
  }
  $('sumFrom').addEventListener('change', sumDate); $('sumTo').addEventListener('change', sumDate);
  /* 整理本工程所有已確認的抽查（不含草稿與照片紀錄） */
  function sumData(p, to) {
    var cons = [], mats = [], defs = [], usedMat = {};
    Object.keys(p.recs || {}).forEach(function (wid) {
      var w = WORK[wid]; if (!w || w.kind === 'photo') return;
      (p.recs[wid] || []).forEach(function (r) {
        if (r.draft || !r.checkDate || r.checkDate > to) return;
        var m = w.kind === 'matform' && matById(p, r.mid), st = recState(w, r), single = !isFix(w);
        if (w.kind === 'matform') usedMat[r.mid] = 1;
        var row = { date: r.checkDate, no: r.docNo || '', name: w.kind === 'matform' ? (m ? m.name : '材料') + '進場抽查' : w.kind === 'phased' ? w.name + PH_SHORT[r.phase] + '抽查' : w.name, loc: r.location || '', st: st };
        (w.kind === 'matform' ? mats : cons).push(row);
        if (st === 'ng' || st === 'fixed') {
          var items = [];
          if (w.kind === 'irregular') (r.rows || []).forEach(function (x) { if (x.r === 'ng') items.push(x.name + (x.note ? '（' + x.note + '）' : '')); });
          else effGroups(w, r).forEach(function (g) { g.items.forEach(function (it) { if (it.c.r === 'ng') items.push(it.name + (it.c.note ? '（' + it.c.note + '）' : '')); }); });
          if (!items.length && r.remark) items.push(r.remark);
          var fd = single ? (r.reTime || '').slice(0, 10) : r.fixDate;
          defs.push({ date: r.checkDate, no: row.no, name: row.name, loc: row.loc, txt: items.join('；'), st: st, fixed: st === 'fixed' ? '已改善' + (fd ? ' ' + roc(fd, '/') : '') : (!single && r.fix === 'track' ? '追蹤改善中' : '未改善') });
        }
      });
    });
    /* 只在材料清單勾選抽查結果、沒有抽查單的材料 */
    (p.mats || []).forEach(function (m) {
      if (usedMat[m.id] || !m.inspDate || !m.inspResult || m.inspDate > to) return;
      mats.push({ date: m.inspDate, no: m.no || '', name: (m.name || '材料') + '進場抽查', loc: '', st: m.inspResult === 'ng' ? 'ng' : 'ok' });
      if (m.inspResult === 'ng') defs.push({ date: m.inspDate, no: m.no || '', name: (m.name || '材料') + '進場抽查', loc: '', txt: m.inspNote || '', st: 'ng', fixed: '未改善' });
    });
    var by = function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : a.no < b.no ? -1 : a.no > b.no ? 1 : 0; };
    return { cons: cons.sort(by), mats: mats.sort(by), defs: defs.sort(by) };
  }
  function renderSumPrint() {
    var p = proj(cur.pid), I = p.info, to = cur.to || today(), from = cur.from || sumFromDef(to), d = sumData(p, to);
    topbar('列印預覽', '抽查總表', I.name);
    $('sumFrom').value = from; $('sumTo').value = to; $('sumBar').hidden = false;
    var RES = { ok: 'OK', ng: '<b class="sum-ng">有缺失</b>', fixed: '缺失已改善', todo: '未完成' };
    var cnt = function (list) { var pre = list.filter(function (x) { return x.date < from; }).length; return { pre: pre, now: list.length - pre, all: list.length }; };
    var head = function (c) { return '<p class="sum-cnt">前期累計：<b>' + c.pre + '</b> 次　　本期：<b>' + c.now + '</b> 次　　本期累計：<b>' + c.all + '</b> 次'; };
    var tbl = function (list, cols, cell) {
      if (!list.length) return '<p class="sum-none">截至本日尚無紀錄。</p>';
      return '<table class="sum"><colgroup>' + cols.map(function (c) { return '<col style="width:' + c[1] + '">'; }).join('') + '</colgroup><thead><tr>' +
        cols.map(function (c) { return '<th>' + c[0] + '</th>'; }).join('') + '</tr></thead><tbody>' +
        list.map(function (x, i) { return '<tr' + (x.date < from ? ' class="prev"' : '') + '><td class="c">' + (i + 1) + '</td>' + cell(x) + '</tr>'; }).join('') + '</tbody></table>';
    };
    var COLS = [['項次', '7%'], ['項目名稱', '37%'], ['位置', '20%'], ['抽驗日期', '16%'], ['檢查結果', '20%']];
    var basic = function (x) { return '<td>' + esc(x.name) + '</td><td class="c">' + esc(x.loc) + '</td><td class="c">' + roc(x.date, '/') + '</td><td class="c">' + RES[x.st] + '</td>'; };
    var cm = cnt(d.mats), cc = cnt(d.cons), cd = cnt(d.defs), fixedN = d.defs.filter(function (x) { return x.st === 'fixed'; }).length;
    var paper = $('paper'); paper.className = 'paper sum-paper';
    paper.innerHTML = '<div class="sum-title"><h3>「' + esc(I.name || '') + '」</h3><p>工程抽查總表</p>' + (I.supervisor ? '<p class="sum-sup">' + esc(I.supervisor) + '</p>' : '') + '</div>' +
      '<p class="sum-asof">截至 ' + roc(to) + '<br><span>本期：' + roc(from, '/') + ' ～ ' + roc(to, '/') + '</span></p>' +
      '<h4>一、材料進場抽驗情形：</h4>' + head(cm) + '</p>' + tbl(d.mats, COLS, basic) +
      '<h4>二、分項施工抽查情形：</h4>' + head(cc) + '</p>' + tbl(d.cons, COLS, basic) +
      '<h4>三、缺失改善追蹤情形：' + (d.defs.length ? '' : '無') + '</h4>' + head(cd) + '　　已改善：<b>' + fixedN + '</b> 次　未改善：<b>' + (d.defs.length - fixedN) + '</b> 次</p>' +
      (d.defs.length ? tbl(d.defs, [['項次', '7%'], ['項目名稱', '22%'], ['位置', '13%'], ['抽驗日期', '12%'], ['缺失內容', '28%'], ['改善情形', '18%']], function (x) {
        return '<td>' + esc(x.name) + '</td><td class="c">' + esc(x.loc) + '</td><td class="c">' + roc(x.date, '/') + '</td><td>' + esc(x.txt) + '</td><td class="c">' + (x.st === 'fixed' ? esc(x.fixed) : '<b class="sum-ng">' + esc(x.fixed) + '</b>') + '</td>';
      }) : '') +
      '<p class="sum-note">註：灰色為前期（' + roc(from, '/') + ' 以前）紀錄；不含尚未確認存檔的草稿。</p>' + foot();
    bindOut(sumOut, 't'); $('pvParts').hidden = true; $('btnMovePv').hidden = true; $('outAt').closest('.out-at').hidden = false;
    $('printHint').textContent = '紙張選 A4 直式、邊界「預設」。「另存 PDF」會直接下載檔案，不經過列印視窗。';
    show('vPreview');
    setDocName('抽查總表_' + (I.name || '') + '_截至' + roc(to, '.'));
  }

  /* ========== 抽查照片：屬於每一張抽查單（照片檔存於 IndexedDB，容量較大） ========== */
  var IDB = (function () {
    var dbp = null;
    function db() {
      return dbp || (dbp = new Promise(function (res, rej) {
        var q = indexedDB.open('inspect-photos', 1);
        q.onupgradeneeded = function () { q.result.createObjectStore('img'); };
        q.onsuccess = function () { res(q.result); }; q.onerror = function () { rej(q.error); };
      }));
    }
    function run(mode, fn) {
      return db().then(function (d) {
        return new Promise(function (res, rej) {
          var t = d.transaction('img', mode), req = fn(t.objectStore('img'));
          t.oncomplete = function () { res(req && req.result); }; t.onerror = function () { rej(t.error); };
        });
      });
    }
    return {
      put: function (key, v) { return run('readwrite', function (st) { return st.put(v, key); }); },
      get: function (key) { return run('readonly', function (st) { return st.get(key); }); },
      del: function (key) { return run('readwrite', function (st) { return st.delete(key); }); },
      keys: function () { return run('readonly', function (st) { return st.getAllKeys ? st.getAllKeys() : null; }).then(function (k) { return k || []; }); },
      all: function () {
        return db().then(function (d) {
          return new Promise(function (res) {
            var out = {}, c = d.transaction('img').objectStore('img').openCursor();
            c.onsuccess = function () { var cu = c.result; if (cu) { out[cu.key] = cu.value; cu.continue(); } else res(out); };
            c.onerror = function () { res(out); };
          });
        });
      }
    };
  })();
  function photoTitle(w) { return w.kind === 'matform' ? '材料進場抽查照片' : w.kind === 'phased' ? w.name + '施工抽查照片' : w.name + '照片'; }
  var thumbCache = {};
  function renderPhotoGrid() {
    var r = R(); if (!r.photos) r.photos = [];
    $('phCount').textContent = r.photos.length ? r.photos.length + ' 張・' + Math.ceil(r.photos.length / 6) + ' 頁' + (r.photos.length > 1 ? '・長按照片可拖曳排序' : '') : '尚無照片';
    $('phGrid').innerHTML = r.photos.map(function (it, i) {
      return '<div class="ph-card" data-pi="' + i + '"><div class="ph-thumb" title="長按照片可拖曳排序"><img alt="照片 ' + (i + 1) + '" data-img="' + it.id + '"></div>' +
        '<label class="f"><span>照片說明</span><textarea rows="2" data-cap="' + i + '" placeholder="例：立柱間距量測">' + esc(it.cap) + '</textarea></label>' +
        '<div class="row-tools"><span class="ph-no">第 ' + (i + 1) + ' 張</span>' +
        '<button type="button" class="link" data-rm="' + i + '">刪除</button></div></div>';
    }).join('');
    fillImgs($('phGrid'));
  }
  function fillImgs(root) {
    root.querySelectorAll('img[data-img]').forEach(function (img) {
      loadImg(img.dataset.img).then(function (u) {
        if (u) { img.src = u; return; }
        var m = document.createElement('span'); m.className = 'ph-miss';
        m.textContent = window.Cloud ? '這張照片還沒上傳到雲端，請在拍照的那台裝置開啟本系統一次' : '這台裝置沒有這張照片（請用含照片的備份檔匯入，或改用雲端共編版）';
        img.replaceWith(m);
      });
    });
  }
  function loadImg(id) {
    if (thumbCache[id]) return Promise.resolve(thumbCache[id]);
    return IDB.get(id).catch(function () { return null; }).then(function (u) {
      if (u || !window.Cloud) return u;
      return window.Cloud.getPhoto(id).then(function (d) { if (d) IDB.put(id, d).catch(function () {}); return d; });
    }).then(function (u) { if (u) thumbCache[id] = u; return u; }).catch(function () { return null; });
  }
  function compress(file) {
    return new Promise(function (res, rej) {
      var url = URL.createObjectURL(file), img = new Image();
      img.onload = function () {
        var sc = Math.min(1, 1280 / Math.max(img.width, img.height)), c = document.createElement('canvas');
        c.width = Math.round(img.width * sc); c.height = Math.round(img.height * sc);
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url);
        res(c.toDataURL('image/jpeg', 0.75));
      };
      img.onerror = function () { URL.revokeObjectURL(url); rej(new Error('無法讀取「' + file.name + '」，請改用 JPG 或 PNG')); };
      img.src = url;
    });
  }
  function addPhotos(files) {
    var r = R(), list = Array.prototype.slice.call(files || []); if (!list.length) return;
    if (!r.photos) r.photos = [];
    $('phMsg').textContent = '正在處理 ' + list.length + ' 張照片…';
    var lastCap = r.photos.length ? r.photos[r.photos.length - 1].cap : '';
    list.reduce(function (pr, f) {
      return pr.then(function () {
        return compress(f).then(function (data) {
          var id = 'ph_' + uid(); thumbCache[id] = data;
          return IDB.put(id, data).then(function () { r.photos.push({ id: id, cap: lastCap }); if (window.Cloud) window.Cloud.uploadPhoto(id, data); });
        });
      });
    }, Promise.resolve()).then(function () {
      save(); renderPhotoGrid(); $('phMsg').textContent = '已加入 ' + list.length + ' 張。';
    }).catch(function (err) { save(); renderPhotoGrid(); $('phMsg').textContent = err.message || '照片加入失敗'; });
  }
  makeSortable($('phGrid'), '.ph-card', function (list) {
    var r = R(), old = r.photos.slice(); r.photos = list.map(function (c) { return old[+c.dataset.pi]; }); save(); renderPhotoGrid();
  }, '.ph-thumb');
  $('phCam').addEventListener('change', function (e) { addPhotos(e.target.files); e.target.value = ''; });
  $('phPick').addEventListener('change', function (e) { addPhotos(e.target.files); e.target.value = ''; });
  $('phGrid').addEventListener('input', function (e) { var i = e.target.dataset.cap; if (i != null) { R().photos[+i].cap = e.target.value; save(); } });
  $('phGrid').addEventListener('click', function (e) {
    var b = e.target.closest('button'); if (!b) return;
    var it = R().photos, d = b.dataset, i;
    if (d.rm != null) { i = +d.rm; var gone = it.splice(i, 1)[0]; IDB.del(gone.id).catch(function () {}); }
    else return;
    save(); renderPhotoGrid();
  });
  /* 列印：接在抽查表後，每頁 2×3，不足補空格 */
  function photoPages(w, r) {
    var items = r.photos || []; if (!items.length) return '';
    var head = esc(photoTitle(w)) + ' ' + roc(r.checkDate, '.'), out = '';
    var img = function (x) { return '<td class="ph-img">' + (x ? '<img alt="" data-img="' + x.id + '">' : '') + '</td>'; };
    var cap = function (x) { return '<td class="ph-cap">' + (x ? esc(x.cap) : '') + '</td>'; };
    for (var i = 0, pn = 1; i < items.length; i += 6, pn++) {
      var pg = items.slice(i, i + 6), rows = '';
      for (var k2 = 0; k2 < 3; k2++) rows += '<tr>' + img(pg[k2 * 2]) + img(pg[k2 * 2 + 1]) + '</tr><tr>' + cap(pg[k2 * 2]) + cap(pg[k2 * 2 + 1]) + '</tr>';
      out += '<section class="ph-page"><h3>' + head + '</h3><table class="ph"><colgroup><col style="width:50%"><col style="width:50%"></colgroup>' + rows + '</table>' +
        '<div class="ph-pno">照片 ' + pn + ' / ' + Math.ceil(items.length / 6) + '　' + esc(r.docNo) + '</div></section>';
    }
    return out;
  }

  /* ========== 移到／複製到其他工程 ========== */
  ['btnMove', 'btnMoveTop', 'btnMovePv'].forEach(function (id) { $(id).addEventListener('click', openMove); });
  function openMove() {
    var p = proj(cur.pid), w = W(), r = R();
    if (S.projects.length < 2) { ask('還沒有其他工程', '請先在工程清單新增要移入的工程，再回來移動或複製這張表。', '好', '前往工程清單', function () { go({ v: 'projects' }); }); return; }
    $('mvInfo').textContent = '「' + recTitle(w, r) + ' ' + r.docNo + '」會放到所選工程的同一工項、同一階段，編號依該工程重新編排。';
    $('mvList').innerHTML = S.projects.filter(function (x) { return x !== p; }).map(function (x, i) {
      return '<label class="pill"><input type="radio" name="mvTo" value="' + x.id + '"' + (i ? '' : ' checked') + '><span>' + esc(x.info.name || '未命名工程') + '</span></label>';
    }).join('');
    $('moveDlg').hidden = false;
  }
  $('mvCancel').addEventListener('click', function () { $('moveDlg').hidden = true; });
  $('mvGo').addEventListener('click', function () {
    var sel = document.querySelector('input[name="mvTo"]:checked'); if (!sel) return;
    var mode = document.querySelector('input[name="mvMode"]:checked').value;
    var from = proj(cur.pid), to = proj(sel.value), w = W(), r = R();
    $('moveDlg').hidden = true;
    var nr = transfer(from, to, w, r, mode);
    if (mode === 'move') { go({ v: 'form', pid: to.id, wid: w.id, rid: nr.id }, true); return; }
    ask('已複製到「' + (to.info.name || '未命名工程') + '」', '新編號 ' + nr.docNo + '，照片也一併複製。', '留在本表', '前往該表', function () {
      go({ v: 'form', pid: to.id, wid: w.id, rid: nr.id });
    });
  });
  function transfer(from, to, w, r, mode) {
    var list = recs(to, w.id), nr = mode === 'move' ? r : JSON.parse(JSON.stringify(r));
    var seq = list.filter(function (x) { return x.phase === r.phase; }).reduce(function (m, x) { return Math.max(m, x.seq); }, 0) + 1;
    if (mode === 'copy') {
      nr.id = uid(); nr.created = Date.now();
      (nr.photos || []).forEach(function (ph) {
        var old = ph.id; ph.id = 'ph_' + uid();
        loadImg(old).then(function (d) { if (!d) return; thumbCache[ph.id] = d; IDB.put(ph.id, d); if (window.Cloud) window.Cloud.uploadPhoto(ph.id, d); });
      });
    } else {
      var src = recs(from, w.id); src.splice(src.indexOf(r), 1);
      (from.plans || []).forEach(function (pl) { if (pl.rid === r.id) delete pl.rid; });
    }
    nr.seq = seq;
    nr.docNo = (to.info.code ? to.info.code + '-' : '') + w.code + (w.kind === 'phased' ? (r.phase + 1) : '') + '-' + ('00' + seq).slice(-3);
    list.push(nr);
    if (to.works && to.works.indexOf(w.id) < 0) to.works.push(w.id);
    save();
    return nr;
  }

  /* ========== 工程日曆：已抽查（讀取表單）、預定抽查（可編輯、可拖曳）、材料進場 ========== */
  var calY, calM, calSel = today();
  function ymd(d) { var t = new Date(d); t.setMinutes(t.getMinutes() - t.getTimezoneOffset()); return t.toISOString().slice(0, 10); }
  function workLabel(w, phase) { return w.kind === 'phased' ? w.name.replace(/工程$/, '') + '（' + ['施工前', '施工中', '完成'][phase] + '）' : w.name; }
  function calEvents(p) {
    var ev = [], linked = {};
    Object.keys(p.recs || {}).forEach(function (wid) {
      var w = WORK[wid]; if (!w) return;
      (p.recs[wid] || []).forEach(function (r) {
        if (!r.checkDate) return;
        var m = w.kind === 'matform' && matById(p, r.mid);
        ev.push({ date: r.checkDate, kind: 'rec', label: m ? '材料抽查：' + m.name : workLabel(w, r.phase), sub: [r.docNo, r.location].filter(Boolean).join(' · '), ng: stats(w, r).ng, wid: wid, rid: r.id });
        linked[r.id] = 1; linked[wid + '|' + (w.kind === 'phased' ? r.phase : 0) + '|' + r.checkDate] = 1;
      });
    });
    (p.mats || []).forEach(function (m) { if (m.arrivalDate) ev.push({ date: m.arrivalDate, kind: 'mat', label: '進場：' + (m.name || '材料'), sub: m.no || '', mid: m.id }); });
    (p.plans || []).forEach(function (pl) {
      if ((pl.rid && linked[pl.rid]) || linked[pl.wid + '|' + (pl.phase || 0) + '|' + pl.date]) return;
      var w = WORK[pl.wid];
      ev.push({ date: pl.date, kind: 'plan', label: (pl.time ? pl.time + ' ' : '') + (w ? workLabel(w, pl.phase || 0) : '抽查'), sub: pl.note || '', plid: pl.id, overdue: pl.date < today() });
    });
    return ev.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : a.label < b.label ? -1 : 1; });
  }
  function renderCal(p) {
    if (calY == null) { var t = new Date(); calY = t.getFullYear(); calM = t.getMonth(); }
    $('calTitle').textContent = (calY - 1911) + ' 年 ' + (calM + 1) + ' 月（' + calY + '）';
    var ev = calEvents(p), byDay = {};
    ev.forEach(function (e) { (byDay[e.date] = byDay[e.date] || []).push(e); });
    var first = new Date(calY, calM, 1), start = new Date(calY, calM, 1 - first.getDay()), td = today(), cells = '';
    ['日', '一', '二', '三', '四', '五', '六'].forEach(function (d) { cells += '<div class="cal-wd">' + d + '</div>'; });
    for (var i = 0; i < 42; i++) {
      var d = new Date(start); d.setDate(start.getDate() + i);
      var key = ymd(d), list = byDay[key] || [], other = d.getMonth() !== calM;
      if (i === 35 && other) break;
      cells += '<button type="button" class="cal-cell' + (other ? ' other' : '') + (key === td ? ' today' : '') + (key === calSel ? ' sel' : '') + '" data-day="' + key + '" aria-label="' + key + ' ' + list.length + ' 筆">' +
        '<span class="cal-n">' + d.getDate() + '</span><span class="cal-evs">' +
        list.slice(0, 3).map(function (e) { return evChip(e, true); }).join('') +
        (list.length > 3 ? '<span class="cal-more">+' + (list.length - 3) + '</span>' : '') + '</span></button>';
    }
    $('calGrid').innerHTML = cells;
    renderCalDay(p, byDay[calSel] || []);
  }
  function evChip(e, small) {
    var cls = 'ev ev-' + e.kind + (e.ng ? ' ev-ng' : '') + (e.overdue ? ' ev-over' : '');
    var attr = e.plid ? ' data-plid="' + e.plid + '" draggable="true"' : '';
    return '<span class="' + cls + '"' + attr + ' title="' + esc(e.label) + '">' + (e.ng ? '╳ ' : e.overdue ? '! ' : '') + esc(e.label) + '</span>';
  }
  var calOpen = null;   // 展開動作列的預定 id
  function renderCalDay(p, list) {
    var on = enabledWorks(p);
    var d = calSel.split('-');
    var tag = { plan: '預定', rec: '已查', mat: '進場' };
    var rows = list.map(function (e) {
      var title = e.label.replace(/^材料抽查：/, '材料：');
      var badge = e.ng ? '<span class="dl-ng">' + e.ng + '缺失</span>' : e.overdue ? '<span class="dl-ng">過期</span>' : '';
      var line = '<span class="dl-tag t-' + e.kind + '">' + tag[e.kind] + '</span>' +
        '<span class="dl-txt"><b>' + esc(title) + '</b>' + (e.sub ? '<small>' + esc(e.sub) + '</small>' : '') + '</span>' + badge;
      if (e.kind === 'plan') {
        var open = calOpen === e.plid, pl = plan(p, e.plid);
        return '<li class="dl-item' + (open ? ' open' : '') + '"><button type="button" class="dl-row" data-toggle="' + e.plid + '" aria-expanded="' + open + '">' + line + '<span class="dl-chev">' + (open ? '▴' : '▾') + '</span></button>' +
          (open ? '<div class="dl-acts"><button type="button" class="btn primary sm" data-start="' + pl.id + '">開始抽查</button>' +
            '<label class="dl-date">改日期<input type="date" data-redate="' + pl.id + '" value="' + pl.date + '"></label>' +
            '<button type="button" class="link" data-delplan="' + pl.id + '">刪除</button></div>' : '') + '</li>';
      }
      var go2 = e.rid ? ' data-openrec="' + e.wid + '|' + e.rid + '"' : ' data-openmat="' + e.mid + '"';
      return '<li class="dl-item"><button type="button" class="dl-row"' + go2 + '>' + line + '<span class="dl-chev">›</span></button></li>';
    }).join('');
    var opts = on.map(function (w) { return '<option value="' + w.id + '">' + esc(w.name) + '</option>'; }).join('');
    $('calDay').innerHTML = '<h3 class="day-h">' + (d[0] - 1911) + ' 年 ' + (+d[1]) + ' 月 ' + (+d[2]) + ' 日' + (calSel === today() ? '（今天）' : '') +
      '<span class="day-n">' + (list.length ? list.length + ' 筆' : '') + '</span></h3>' +
      (rows ? '<ul class="day-list">' + rows + '</ul>' : '<p class="muted day-empty">這天沒有抽查紀錄或預定。</p>') +
      '<details class="plan-add"><summary>＋ 新增 ' + (+d[1]) + '/' + (+d[2]) + ' 預定抽查</summary><div class="plan-form"><div class="grid g3">' +
      '<label class="f">工項<select id="plWork">' + opts + '</select></label>' +
      '<label class="f" id="plPhaseBox">階段<select id="plPhase"><option value="0">施工前</option><option value="1">施工中檢查</option><option value="2">施工完成檢查</option></select></label>' +
      '<label class="f">時間（選填）<input id="plTime" type="time"></label>' +
      '<label class="f wide">備註（選填）<input id="plNote" placeholder="例：3F 東側隔間封板前"></label></div>' +
      '<button type="button" class="btn primary sm" id="btnAddPlan">加入預定</button></div></details>';
    var syncPh = function () { var w = WORK[$('plWork').value]; $('plPhaseBox').hidden = !w || w.kind !== 'phased'; };
    $('plWork').addEventListener('change', syncPh); syncPh();
  }
  function plan(p, id) { return (p.plans || []).filter(function (x) { return x.id === id; })[0]; }
  $('calPrev').addEventListener('click', function () { calM--; if (calM < 0) { calM = 11; calY--; } renderCal(proj(cur.pid)); });
  $('calNext').addEventListener('click', function () { calM++; if (calM > 11) { calM = 0; calY++; } renderCal(proj(cur.pid)); });
  $('calToday').addEventListener('click', function () { var t = new Date(); calY = t.getFullYear(); calM = t.getMonth(); calSel = today(); renderCal(proj(cur.pid)); });
  $('calGrid').addEventListener('click', function (e) {
    var c = e.target.closest('[data-day]'); if (!c) return;
    calSel = c.dataset.day; renderCal(proj(cur.pid));
    if (window.innerWidth < 960) $('calDay').scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  });
  /* 拖曳預定到其他日期（電腦） */
  $('calGrid').addEventListener('dragstart', function (e) { var c = e.target.closest('[data-plid]'); if (c) e.dataTransfer.setData('text/plain', c.dataset.plid); });
  $('calGrid').addEventListener('dragover', function (e) { if (e.target.closest('[data-day]')) e.preventDefault(); });
  $('calGrid').addEventListener('drop', function (e) {
    var c = e.target.closest('[data-day]'), id = e.dataTransfer.getData('text/plain'); if (!c || !id) return;
    e.preventDefault(); var p = proj(cur.pid), pl = plan(p, id); if (!pl) return;
    pl.date = c.dataset.day; calSel = pl.date; save(); renderCal(p);
  });
  $('calDay').addEventListener('click', function (e) {
    var p = proj(cur.pid), b = e.target.closest('button'); if (!b) return;
    var ds = b.dataset;
    if (b.id === 'btnAddPlan') {
      var wid = $('plWork').value, w = WORK[wid];
      (p.plans || (p.plans = [])).push({ id: uid(), date: calSel, time: $('plTime').value, wid: wid, phase: w && w.kind === 'phased' ? +$('plPhase').value : 0, note: $('plNote').value.trim() });
      save(); renderCal(p); return;
    }
    if (ds.toggle) { calOpen = calOpen === ds.toggle ? null : ds.toggle; renderCal(p); return; }
    if (ds.delplan) { p.plans = p.plans.filter(function (x) { return x.id !== ds.delplan; }); save(); renderCal(p); return; }
    if (ds.start) {
      var pl = plan(p, ds.start), w2 = WORK[pl.wid];
      if (!w2) return;
      var r = newRecord(p, w2, w2.kind === 'phased' ? pl.phase : 0, { checkDate: pl.date, location: pl.note || '' });
      pl.rid = r.id; save();
      go({ v: 'form', pid: p.id, wid: w2.id, rid: r.id }); return;
    }
    if (ds.openrec) { var a = ds.openrec.split('|'); go({ v: 'form', pid: p.id, wid: a[0], rid: a[1] }); return; }
    if (ds.openmat) go({ v: 'matItem', pid: p.id, mid: ds.openmat });
  });
  $('calDay').addEventListener('change', function (e) {
    var id = e.target.dataset.redate; if (!id || !e.target.value) return;
    var p = proj(cur.pid), pl = plan(p, id); pl.date = e.target.value; calSel = pl.date;
    var d = new Date(pl.date + 'T00:00:00'); calY = d.getFullYear(); calM = d.getMonth();
    save(); renderCal(p);
  });

  /* ========== 圖說與抽查位置：匯入 PDF／圖片，用色塊標記抽查位置 ========== */
  var DW_COLORS = ['#2f6fed', '#e5383b', '#2a9d4b', '#f08c00', '#8e44ad'];
  var DW_NAMES = ['藍', '紅', '綠', '橘', '紫'];
  var dwColor = 0, dwZoom = 1, dwSel = null, dwMode = 'view', dwLastDid = null;
  function dwgs(p) { return p.dwgs || (p.dwgs = []); }
  function dwgById(p, id) { return dwgs(p).filter(function (g) { return g.id === id; })[0]; }
  function G() { return dwgById(proj(cur.pid), cur.did); }
  function markCount(p) { return dwgs(p).reduce(function (a, g) { return a + g.marks.length; }, 0); }
  function linkInfo(p, link) {
    if (!link) return null;
    var a = link.split('|'), w = WORK[a[0]], r = w && recById(p, a[0], a[1]);
    if (!r) return null;
    return { w: w, r: r, short: r.docNo, long: (w.kind === 'phased' ? w.name + '(' + (r.phase + 1) + ')' : recTitle(w, r)) + ' ' + r.docNo };
  }
  function markText(p, m) { var li = linkInfo(p, m.rid); return m.t || (li ? li.short : ''); }
  function linkedMarks(p, link) {
    var out = []; dwgs(p).forEach(function (g) { g.marks.forEach(function (m) { if (m.rid === link) out.push({ g: g, m: m }); }); }); return out;
  }

  /* 工項目錄中的「圖說」入口 */
  function dwgCard(p) {
    var n = dwgs(p).length;
    return '<button type="button" class="work-card special dwg-card" data-dwgs="1"><span class="wc-code">DWG</span><span class="wc-name">圖說・抽查位置</span>' +
      '<span class="wc-meta">' + (n ? n + ' 張圖說 · ' + markCount(p) + ' 個色塊' : '匯入圖說並標記抽查位置') + '</span></button>';
  }

  function renderDwgs() {
    var p = proj(cur.pid), li = linkInfo(p, cur.link);
    topbar('圖說與抽查位置', p.info.name || '未命名工程', li ? '返回抽查表' : (p.info.name || '工程'));
    $('dwgLinkHint').hidden = !li;
    if (li) $('dwgLinkHint').textContent = '選擇要標記「' + li.long + '」抽查位置的圖說。';
    var list = dwgs(p);
    $('dwgList').innerHTML = list.length ? list.map(function (g) {
      var linked = cur.link ? g.marks.filter(function (m) { return m.rid === cur.link; }).length : 0;
      return '<button type="button" class="dwg-item" data-did="' + g.id + '"><span class="dwg-thumb"><img alt="" data-img="' + g.img + '"></span>' +
        '<span class="dwg-meta"><b>' + esc(g.name) + '</b><small>' + g.marks.length + ' 個色塊' + (linked ? ' · 本表 ' + linked + ' 處' : '') + '</small></span></button>';
    }).join('') : '<div class="empty">尚未匯入圖說。按「匯入圖說」選擇 PDF 或圖片（可多選，PDF 每一頁會成為一張圖說）。</div>';
    fillImgs($('dwgList'));
    show('vDwgs');
    findOrphanDwgs();
  }
  /* 找回：這台裝置存有圖檔、但已不在任何工程清單中的圖說（例如被舊版畫面同步覆蓋） */
  function findOrphanDwgs() {
    $('dwgOrphans').hidden = true;
    IDB.keys().then(function (keys) {
      var used = {}; S.projects.forEach(function (p) { (p.dwgs || []).forEach(function (g) { used[g.img] = 1; }); });
      var lost = keys.filter(function (k) { return /^dw_/.test(k) && !used[k]; });
      if (!lost.length || cur.v !== 'dwgs') return;
      $('dwgOrphans').hidden = false;
      $('dwgOrphanList').innerHTML = lost.map(function (k) {
        return '<label class="orphan"><input type="checkbox" value="' + k + '" checked><span class="dwg-thumb"><img alt="" data-img="' + k + '"></span></label>';
      }).join('');
      $('dwgOrphanMsg').textContent = '這台裝置還留有 ' + lost.length + ' 張先前匯入、但目前不在任何工程清單中的圖說。勾選後可加回本工程（色塊標記無法一併找回，需重新標記）。';
      fillImgs($('dwgOrphanList'));
    }).catch(function () {});
  }
  /* 還原圖說：只加回目前沒有的圖說與色塊，不覆蓋、不刪除現有資料 */
  function parseRestore(text) {
    var out = [], clean = String(text || '').replace(/\t/g, '');   // 整列複製時欄位以 Tab 分隔，JSON 內不含真正的 Tab
    try { var j = JSON.parse(clean.trim()); if (j && j.restoreDwgs) return j.restoreDwgs; if (j && j.dwgs) return [{ pid: j.id, dwgs: j.dwgs }]; } catch (e) {}
    clean.split(/\r?\n/).forEach(function (line) {
      var i = line.indexOf('{"'); if (i < 0) return;
      var body = line.slice(i), k = body.lastIndexOf('}');
      try { var d = JSON.parse(body.slice(0, k + 1)); if (d && d.id && d.dwgs) out.push({ pid: d.id, dwgs: d.dwgs }); } catch (e) {}
    });
    return out;
  }
  function doRestore(text) {
    var list = parseRestore(text), addG = 0, addM = 0, miss = 0, seen = 0;
    if (!list.length) { $('rsMsg').textContent = '沒有讀到圖說資料。請確認貼上的是 _sync 分頁中「project:」開頭那一列的內容。'; return; }
    list.forEach(function (x) {
      var p = proj(x.pid); if (!p) { miss++; return; }
      var cur0 = dwgs(p), byId = {}; cur0.forEach(function (g) { byId[g.id] = g; });
      (x.dwgs || []).forEach(function (g) {
        seen++;
        var have = byId[g.id];
        if (!have) { cur0.push(JSON.parse(JSON.stringify(g))); addG++; addM += (g.marks || []).length; return; }
        var mk = {}; have.marks.forEach(function (m) { mk[m.id] = 1; });
        (g.marks || []).forEach(function (m) { if (!mk[m.id]) { have.marks.push(JSON.parse(JSON.stringify(m))); addM++; } });
      });
    });
    if (addG || addM) save();
    $('rsMsg').textContent = '讀到 ' + seen + ' 張圖說：加回 ' + addG + ' 張圖說、' + addM + ' 個色塊' + (miss ? '；有 ' + miss + ' 個工程在本系統中已不存在，略過' : '') + (addG || addM ? '。' : '（目前都已存在，不需還原）。');
    if (addG || addM) { renderDwgs(); $('dwgRestore').open = true; }
  }
  $('rsGo').addEventListener('click', function () { doRestore($('rsText').value); });
  $('rsFile').addEventListener('change', function (e) {
    var f = e.target.files[0]; e.target.value = ''; if (!f) return;
    var rd = new FileReader(); rd.onload = function () { doRestore(rd.result); }; rd.readAsText(f);
  });
  $('btnOrphanAdd').addEventListener('click', function () {
    var p = proj(cur.pid), ids = Array.prototype.slice.call(document.querySelectorAll('#dwgOrphanList input:checked')).map(function (x) { return x.value; });
    if (!ids.length) return;
    Promise.all(ids.map(function (id, i) {
      return loadImg(id).then(function (u) {
        return new Promise(function (res) {
          var im = new Image(); im.onload = function () { res({ id: uid(), name: '找回的圖說 ' + (dwgs(p).length + i + 1), img: id, w: im.naturalWidth, h: im.naturalHeight, marks: [], created: Date.now() }); };
          im.onerror = function () { res(null); }; im.src = u;
        });
      });
    })).then(function (gs) {
      gs.filter(Boolean).forEach(function (g) { dwgs(p).push(g); if (window.Cloud) window.Cloud.uploadPhoto(g.img, thumbCache[g.img]); });
      save(); renderDwgs();
    });
  });
  $('dwgList').addEventListener('click', function (e) {
    var b = e.target.closest('[data-did]'); if (b) go({ v: 'dwg', pid: cur.pid, did: b.dataset.did, link: cur.link });
  });
  function canvasJpeg(c) { return c.toDataURL('image/jpeg', 0.85); }
  function imgFileToDwg(file) {
    return new Promise(function (res, rej) {
      var url = URL.createObjectURL(file), img = new Image();
      img.onload = function () {
        var sc = Math.min(1, 3200 / Math.max(img.width, img.height)), c = document.createElement('canvas');
        c.width = Math.round(img.width * sc); c.height = Math.round(img.height * sc);
        var x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0, c.width, c.height);
        URL.revokeObjectURL(url); res([{ data: canvasJpeg(c), w: c.width, h: c.height, name: file.name.replace(/\.[^.]+$/, '') }]);
      };
      img.onerror = function () { URL.revokeObjectURL(url); rej(new Error('無法讀取「' + file.name + '」，請改用 PDF、JPG 或 PNG')); };
      img.src = url;
    });
  }
  function pdfFileToDwg(file) {
    return file.arrayBuffer().then(function (buf) {
      return (window.pdfjsLib ? Promise.resolve() : loadScript(CDN.pdf)).then(function () {
        pdfjsLib.GlobalWorkerOptions.workerSrc = CDN.pdfWorker;
        return pdfjsLib.getDocument({ data: new Uint8Array(buf), cMapUrl: CDN.pdfCmaps, cMapPacked: true }).promise;
      });
    }).then(function (pdf) {
      var n = Math.min(pdf.numPages, 30), out = [], base = file.name.replace(/\.pdf$/i, '');
      var seq = Promise.resolve();
      for (var i = 1; i <= n; i++) (function (i) {
        seq = seq.then(function () {
          $('dwgMsg').textContent = '正在轉換「' + file.name + '」第 ' + i + ' / ' + n + ' 頁…';
          return pdf.getPage(i).then(function (pg) {
            var v1 = pg.getViewport({ scale: 1 }), sc = Math.min(6, 3200 / Math.max(v1.width, v1.height)), vp = pg.getViewport({ scale: sc });
            var c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
            var x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
            return pg.render({ canvasContext: x, viewport: vp }).promise.then(function () {
              out.push({ data: canvasJpeg(c), w: c.width, h: c.height, name: base + (n > 1 ? ' 第' + i + '頁' : '') });
              c.width = c.height = 0;
            });
          });
        });
      })(i);
      return seq.then(function () { return out; });
    });
  }
  $('dwgFile').addEventListener('change', function (e) {
    var files = Array.prototype.slice.call(e.target.files || []); e.target.value = ''; if (!files.length) return;
    var p = proj(cur.pid), added = [];
    $('dwgMsg').textContent = '正在讀取圖說…';
    files.reduce(function (pr, f) {
      return pr.then(function () {
        var job = /pdf$/i.test(f.type) || /\.pdf$/i.test(f.name) ? pdfFileToDwg(f) : imgFileToDwg(f);
        return job.then(function (pages) {
          return pages.reduce(function (q, pg) {
            return q.then(function () {
              var id = 'dw_' + uid(); thumbCache[id] = pg.data;
              return IDB.put(id, pg.data).then(function () {
                var g = { id: uid(), name: pg.name, img: id, w: pg.w, h: pg.h, marks: [], created: Date.now() };
                dwgs(p).push(g); added.push(g); save();
                if (window.Cloud) window.Cloud.uploadPhoto(id, pg.data);
              });
            });
          }, Promise.resolve());
        });
      });
    }, Promise.resolve()).then(function () {
      $('dwgMsg').textContent = '已匯入 ' + added.length + ' 張圖說。';
      if (added.length === 1) go({ v: 'dwg', pid: p.id, did: added[0].id, link: cur.link }, !!cur.link); else renderDwgs();
    }).catch(function (err) { $('dwgMsg').textContent = err.message || '匯入失敗'; renderDwgs(); });
  });

  /* 標記編輯器 */
  function colorsHtml(sel, attr) {
    return DW_COLORS.map(function (c, i) { return '<button type="button" class="dw-sw' + (i === sel ? ' on' : '') + '" ' + attr + '="' + i + '" style="--c:' + c + '" aria-label="' + DW_NAMES[i] + '色"></button>'; }).join('');
  }
  /* 色塊狀態：對應抽查紀錄者依結果自動上色 */
  var ST_COLOR = { ok: '#2f6fed', ng: '#e5383b', fixed: '#2a9d4b', todo: '#8c8c8c' };
  var ST_NAME = { ok: '合格', ng: '有缺失', fixed: '已改善', todo: '尚未判定', none: '未對應' };
  function recState(w, r) {
    var single = !isFix(w), s = stats(w, r);
    var ng = single ? (r.result === 'ng' || s.ng > 0) : s.ng > 0;
    if (ng) return (single ? r.re === 'ok' : r.fix === 'done') ? 'fixed' : 'ng';
    return s.d || r.result || w.kind === 'photo' ? 'ok' : 'todo';
  }
  function markState(p, m) { var li = linkInfo(p, m.rid); return li ? recState(li.w, li.r) : 'none'; }
  function colorOf(m) { var p = proj(cur.pid), st = m.rid ? markState(p, m) : 'none'; return st === 'none' ? DW_COLORS[m.c || 0] : ST_COLOR[st]; }
  /* 篩選（從工項目錄進入時）；從抽查單進入時只看本表 */
  var dwF = { period: null, from: '', to: '', st: [], ph: [], wk: [] }, dwOthers = false, dwFOpen = false, dwInfoId = null, dwEditing = false;
  function markDate(p, m) { var li = linkInfo(p, m.rid); return li ? li.r.checkDate || '' : m.created ? ymd(m.created) : ''; }
  function periodFrom(f) {
    var t = today();
    if (f.period === 'week') { var d = new Date(t + 'T00:00:00'); return addDays(t, -((d.getDay() + 6) % 7)); }
    if (f.period === 'month') return t.slice(0, 8) + '01';
    if (f.period === '3m') return addDays(t, -90);
    if (f.period === 'custom') return f.from || '';
    return '';
  }
  function passFilter(p, m) {
    var li = linkInfo(p, m.rid), st = li ? recState(li.w, li.r) : 'none';
    if (dwF.st.length && dwF.st.indexOf(st) < 0) return false;
    if (dwF.wk.length && (!li || dwF.wk.indexOf(li.w.id) < 0)) return false;
    if (dwF.ph.length && (!li || li.w.kind !== 'phased' || dwF.ph.indexOf(String(li.r.phase)) < 0)) return false;
    if (dwF.period && dwF.period !== 'all') {
      var d = markDate(p, m), a = periodFrom(dwF), b = dwF.period === 'custom' ? dwF.to : '';
      if (!d || (a && d < a) || (b && d > b)) return false;
    }
    return true;
  }
  function filterOn() { return !!(dwF.st.length || dwF.wk.length || dwF.ph.length || (dwF.period && dwF.period !== 'all')); }
  /* 每個色塊在目前畫面的狀態：show 正常、ghost 淡灰參考、hide 不顯示 */
  function markVis(p, m) {
    if (cur.link) return m.rid === cur.link ? 'show' : dwOthers ? 'ghost' : 'hide';
    return passFilter(p, m) ? 'show' : 'hide';
  }
  function filterDesc() {
    var out = [];
    if (dwF.period && dwF.period !== 'all') out.push({ week: '本週', month: '本月', '3m': '近 3 個月', custom: (dwF.from ? roc(dwF.from, '.') : '') + '～' + (dwF.to ? roc(dwF.to, '.') : '') }[dwF.period]);
    if (dwF.st.length) out.push(dwF.st.map(function (x) { return ST_NAME[x]; }).join('、'));
    if (dwF.ph.length) out.push(dwF.ph.map(function (x) { return PH_SHORT[+x]; }).join('、'));
    if (dwF.wk.length) out.push(dwF.wk.map(function (x) { return WORK[x] ? WORK[x].name : x; }).join('、'));
    return out.join('；');
  }
  function renderFilter() {
    var p = proj(cur.pid), g = G(), wk = {};
    g.marks.forEach(function (m) { var li = linkInfo(p, m.rid); if (li) wk[li.w.id] = li.w.name; });
    var chip = function (grp, v, label) { var on = dwF[grp].indexOf(v) >= 0; return '<button type="button" class="chip' + (on ? ' on' : '') + '" aria-pressed="' + on + '" data-f="' + grp + '" data-v="' + v + '">' + esc(label) + '</button>'; };
    var per = [['all', '全部'], ['week', '本週'], ['month', '本月'], ['3m', '近 3 個月'], ['custom', '自訂']];
    $('dwFilter').innerHTML =
      '<div class="dwf-row"><span class="dwf-l">期間</span>' + per.map(function (x) { return '<button type="button" class="chip' + ((dwF.period || 'all') === x[0] ? ' on' : '') + '" data-per="' + x[0] + '">' + x[1] + '</button>'; }).join('') +
      (dwF.period === 'custom' ? '<input type="date" id="dwfFrom" value="' + dwF.from + '"> ～ <input type="date" id="dwfTo" value="' + dwF.to + '">' : '') + '</div>' +
      '<div class="dwf-row"><span class="dwf-l">狀態</span>' + ['ng', 'fixed', 'ok', 'todo', 'none'].map(function (x) { return chip('st', x, ST_NAME[x]); }).join('') + '</div>' +
      '<div class="dwf-row"><span class="dwf-l">階段</span>' + [0, 1, 2].map(function (x) { return chip('ph', String(x), PH_SHORT[x]); }).join('') + '</div>' +
      (Object.keys(wk).length ? '<div class="dwf-row"><span class="dwf-l">工項</span>' + Object.keys(wk).map(function (x) { return chip('wk', x, wk[x]); }).join('') + '</div>' : '') +
      '<div class="dwf-row"><button type="button" class="link" id="dwfClear">清除篩選（全部顯示）</button></div>';
  }
  function renderDwg() {
    var p = proj(cur.pid), g = G(), li = linkInfo(p, cur.link);
    if (dwLastDid !== g.id + (cur.link || '')) {
      dwLastDid = g.id + (cur.link || ''); dwZoom = 1; dwSel = null; dwInfoId = null; dwEditing = false; dwMode = li ? 'rect' : 'view'; dwOthers = false;
      if (dwF.period === null) dwF.period = g.marks.length > 20 ? 'month' : 'all';
    }
    topbar(g.name, p.info.name || '未命名工程', li ? '返回' : '圖說');
    document.querySelectorAll('input[name="dwMode"]').forEach(function (x) { x.checked = x.value === dwMode; });
    $('dwColors').innerHTML = colorsHtml(dwColor, 'data-dwc');
    $('dwName').value = g.name;
    $('dwOthersBox').hidden = !li; $('dwOthers').checked = dwOthers;
    $('dwFilterBtn').hidden = !!li; $('dwFilter').hidden = !!li || !dwFOpen; $('dwFilterBtn').setAttribute('aria-expanded', String(dwFOpen && !li));
    if (!li && dwFOpen) renderFilter();
    $('dwTip').textContent = li ? '在圖上拖曳框出「' + li.long + '」的抽查位置，畫好的色塊會自動對應這張抽查表。勾選「顯示其他紀錄」可淡灰顯示其他抽查單的位置作參考。' :
      '點標記可看該抽查單摘要；選 ▭ 方框、◯ 圓圈或 ↗ 箭頭後在圖上拖曳新增，雙指或 Ctrl＋滾輪可縮放。對應抽查紀錄的色塊依結果自動上色：藍＝合格、紅＝有缺失、綠＝已改善、灰＝尚未判定。';
    show('vDwg');
    var img = $('dwImg');
    if (img.dataset.id !== g.img) { img.removeAttribute('src'); img.dataset.id = g.img; loadImg(g.img).then(function (u) { if (u && img.dataset.id === g.img) img.src = u; }); }
    sizeDwg(); drawMarks(); dwPanel();
  }
  function sizeDwg() {
    var g = G(), st = $('dwStage'), cw = Math.max(200, st.clientWidth) * dwZoom;
    $('dwCanvas').style.width = cw + 'px'; $('dwCanvas').style.height = (cw * g.h / g.w) + 'px';
    st.classList.toggle('marking', dwMode !== 'view');
  }
  function rgba(hex, al) { var n = parseInt(hex.slice(1), 16); return 'rgba(' + (n >> 16) + ',' + ((n >> 8) & 255) + ',' + (n & 255) + ',' + al + ')'; }
  function markStyle(m, col) { col = col || colorOf(m); return 'left:' + (m.x * 100) + '%;top:' + (m.y * 100) + '%;width:' + (m.w * 100) + '%;height:' + (m.h * 100) + '%;--c:' + col + ';--bg:' + rgba(col, 0.32); }
  /* 標記形狀：方框（rect）、圓圈（ellipse）、箭頭（arrow：起點 x,y，終點 x+w,y+h） */
  function markHtml(m, o) {
    o = o || {}; var col = o.col || colorOf(m), ratio = o.ratio || 1, tag = o.tag != null ? '<span class="dw-tag"' + '%TAGSTYLE%' + '>' + o.tag + '</span>' : '';
    var attr = (o.id ? ' data-mk="' + m.id + '"' : ''), handle = o.sel ? '<i class="dw-h"></i>' : '';
    var cls = 'dw-mark' + (o.cls ? ' ' + o.cls : '') + (o.sel ? ' sel' : '');
    if (m.s === 'arrow') {
      var dx = m.w, dy = m.h * ratio, len = Math.sqrt(dx * dx + dy * dy), ang = Math.atan2(dy, dx) * 180 / Math.PI;
      return '<div class="' + cls + ' dw-arrow"' + attr + ' style="left:' + (m.x * 100) + '%;top:' + (m.y * 100) + '%;width:' + (len * 100) + '%;transform:rotate(' + ang.toFixed(2) + 'deg);--c:' + col + '">' +
        tag.replace('%TAGSTYLE%', ' style="transform:rotate(' + (-ang).toFixed(2) + 'deg)"') + handle + '</div>';
    }
    return '<div class="' + cls + (m.s === 'ellipse' ? ' dw-ell' : '') + '"' + attr + ' style="' + markStyle(m, col) + '">' + tag.replace('%TAGSTYLE%', '') + handle + '</div>';
  }
  function gRatio() { var g = G(); return g ? g.h / g.w : 1; }
  function drawMarks() {
    var p = proj(cur.pid), g = G(), shown = 0;
    $('dwMarks').innerHTML = g.marks.map(function (m, i) {
      var v = markVis(p, m); if (v === 'hide' && m.id !== dwSel) return '';
      if (v === 'ghost') return markHtml(m, { col: '#9a9a9a', cls: 'ghost', ratio: g.h / g.w });
      shown++;
      return oneMark(p, g, m, i);
    }).join('');
    $('dwCount').textContent = (shown < g.marks.length ? '顯示 ' + shown + ' / ' : '') + g.marks.length + ' 個';
    $('dwFilterN').textContent = filterOn() ? '（' + shown + '/' + g.marks.length + '）' : '';
    $('dwFilterBtn').classList.toggle('on', filterOn());
    $('dwMarkList').innerHTML = g.marks.map(function (m, i) {
      if (markVis(p, m) !== 'show') return '';
      var li = linkInfo(p, m.rid), st = li ? recState(li.w, li.r) : 'none';
      return '<li class="dl-item"><button type="button" class="dl-row" data-pick="' + m.id + '"><span class="dw-dot" style="--c:' + colorOf(m) + '">' + (i + 1) + '</span>' +
        '<span class="dl-txt"><b>' + esc(m.t || (li ? li.short : '未命名標記')) + '</b><small>' + esc(li ? li.long + (li.r.checkDate ? ' · ' + roc(li.r.checkDate, '.') : '') + (li.r.location ? ' · ' + li.r.location : '') : '未對應抽查紀錄') + '</small></span>' +
        (st === 'ng' ? '<span class="dl-ng">缺失</span>' : '') + '<span class="dl-chev">›</span></button></li>';
    }).join('') || '<li class="dl-item"><span class="dl-row muted">' + (g.marks.length ? '目前條件下沒有色塊。' : '尚無色塊。') + '</span></li>';
  }
  function oneMark(p, g, m, i) { var t = markText(p, m); return markHtml(m, { id: true, sel: m.id === dwSel, ratio: g.h / g.w, tag: (i + 1) + (t ? ' ' + esc(t) : '') }); }
  function curMark() { var g = G(); return g && g.marks.filter(function (m) { return m.id === dwSel; })[0]; }
  function recOptions(p, sel) {
    var html = '<option value="">（不對應）</option>';
    enabledWorks(p).concat(D.matWork ? [D.matWork] : []).forEach(function (w) {
      var list = (p.recs[w.id] || []).filter(function (r) { return !r.draft || w.id + '|' + r.id === cur.link || w.id + '|' + r.id === sel; });
      if (!list.length) return;
      html += '<optgroup label="' + esc(w.name) + '">' + list.map(function (r) {
        var v = w.id + '|' + r.id;
        return '<option value="' + v + '"' + (v === sel ? ' selected' : '') + '>' + esc((w.kind === 'phased' ? '(' + (r.phase + 1) + ') ' : '') + r.docNo + (r.location ? ' · ' + r.location : '')) + '</option>';
      }).join('') + '</optgroup>';
    });
    return html;
  }
  function dwPanel() {
    var m = curMark(), p = proj(cur.pid), li = m && linkInfo(p, m.rid);
    var info = !!(m && li && !dwEditing && m.rid !== cur.link);
    dwInfo(info ? m : null);
    $('dwEdit').hidden = !m || info; if (!m || info) return;
    $('dwEditColors').hidden = !!li; $('dwAutoColor').hidden = !li;
    $('dwLabel').value = m.t || ''; $('dwLabel').placeholder = markText(p, Object.assign({}, m, { t: '' })) || '例：抽查位置、C3 柱';
    $('dwRec').innerHTML = recOptions(p, m.rid || '');
    $('dwEditColors').innerHTML = colorsHtml(m.c || 0, 'data-mkc');
    $('dwOpenRec').hidden = !linkInfo(p, m.rid);
  }
  function dwInfo(m) {
    var box = $('dwInfo'); box.hidden = !m; if (!m) return;
    var p = proj(cur.pid), li = linkInfo(p, m.rid), r = li.r, w = li.w, st = recState(w, r), s = stats(w, r), ph = (r.photos || []).slice(0, 3);
    box.innerHTML = '<div class="dwi-head"><span class="dw-dot" style="--c:' + ST_COLOR[st] + '">' + (G().marks.indexOf(m) + 1) + '</span><b>' + esc(li.long) + '</b>' +
      '<button type="button" class="dwi-x" data-dwi="close" aria-label="關閉">✕</button></div>' +
      '<div class="dwi-meta">' + esc([r.checkDate && roc(r.checkDate, '.'), r.location].filter(Boolean).join(' · ') || '尚未填寫日期與位置') + '</div>' +
      '<div class="dwi-meta"><span class="st-chip" style="--c:' + ST_COLOR[st] + '">' + ST_NAME[st] + '</span>' + (s.t ? ' 已判定 ' + s.d + '/' + s.t + (s.ng ? '，缺失 ' + s.ng + ' 項' : '') : '') + (m.t ? ' · 標示：' + esc(m.t) : '') + '</div>' +
      (ph.length ? '<div class="dwi-ph">' + ph.map(function (x) { return '<img alt="" data-img="' + x.id + '">'; }).join('') + ((r.photos || []).length > 3 ? '<span>+' + (r.photos.length - 3) + '</span>' : '') + '</div>' : '') +
      '<div class="dwi-btns"><button type="button" class="btn primary sm" data-dwi="open">開啟抽查單 ›</button><button type="button" class="btn ghost sm" data-dwi="edit">編輯色塊</button></div>';
    fillImgs(box);
  }
  $('dwInfo').addEventListener('click', function (e) {
    var b = e.target.closest('[data-dwi]'); if (!b) return; var a = b.dataset.dwi, m = curMark();
    if (a === 'close') selectMark(null);
    else if (a === 'edit') { dwEditing = true; dwPanel(); $('dwEdit').scrollIntoView({ block: 'nearest', behavior: 'smooth' }); }
    else if (a === 'open' && m) { var x = m.rid.split('|'); go({ v: 'form', pid: cur.pid, wid: x[0], rid: x[1] }); }
  });
  function selectMark(id, scroll) {
    if (id !== dwSel) dwEditing = false;
    dwSel = id; drawMarks(); dwPanel();
    if (scroll) { var el = document.querySelector('[data-mk="' + id + '"]'); if (el) el.scrollIntoView({ block: 'center', inline: 'center', behavior: 'smooth' }); }
  }
  document.querySelectorAll('input[name="dwMode"]').forEach(function (x) { x.addEventListener('change', function () { dwMode = x.value; sizeDwg(); }); });
  /* 縮放：按鈕、Ctrl＋滾輪／觸控板（以游標為中心）、手機雙指縮放 */
  $('dwColors').addEventListener('click', function (e) {
    var b = e.target.closest('[data-dwc]'); if (!b) return; dwColor = +b.dataset.dwc; $('dwColors').innerHTML = colorsHtml(dwColor, 'data-dwc');
  });
  $('dwEditColors').addEventListener('click', function (e) {
    var b = e.target.closest('[data-mkc]'), m = curMark(); if (!b || !m) return;
    m.c = dwColor = +b.dataset.mkc; save(); $('dwColors').innerHTML = colorsHtml(dwColor, 'data-dwc'); drawMarks(); dwPanel();
  });
  $('dwLabel').addEventListener('input', function () { var m = curMark(); if (m) { m.t = $('dwLabel').value; save(); drawMarks(); } });
  $('dwRec').addEventListener('change', function () { var m = curMark(); if (m) { m.rid = $('dwRec').value; save(); drawMarks(); dwPanel(); } });
  $('dwDone').addEventListener('click', function () { selectMark(null); });
  $('dwOthers').addEventListener('change', function () { dwOthers = $('dwOthers').checked; drawMarks(); });
  $('dwFilterBtn').addEventListener('click', function () { dwFOpen = !dwFOpen; $('dwFilter').hidden = !dwFOpen; $('dwFilterBtn').setAttribute('aria-expanded', String(dwFOpen)); if (dwFOpen) renderFilter(); });
  $('dwFilter').addEventListener('click', function (e) {
    var b = e.target.closest('button'); if (!b) return;
    if (b.id === 'dwfClear') dwF = { period: 'all', from: '', to: '', st: [], ph: [], wk: [] };
    else if (b.dataset.per) dwF.period = b.dataset.per;
    else if (b.dataset.f) { var a = dwF[b.dataset.f], i = a.indexOf(b.dataset.v); if (i < 0) a.push(b.dataset.v); else a.splice(i, 1); }
    else return;
    renderFilter(); drawMarks();
  });
  $('dwFilter').addEventListener('change', function (e) { if (e.target.id === 'dwfFrom') dwF.from = e.target.value; if (e.target.id === 'dwfTo') dwF.to = e.target.value; drawMarks(); });
  $('dwDelMark').addEventListener('click', function () { var g = G(); g.marks = g.marks.filter(function (m) { return m.id !== dwSel; }); save(); selectMark(null); });
  $('dwOpenRec').addEventListener('click', function () { var m = curMark(), a = m && m.rid.split('|'); if (a) go({ v: 'form', pid: cur.pid, wid: a[0], rid: a[1] }); });
  $('dwMarkList').addEventListener('click', function (e) { var b = e.target.closest('[data-pick]'); if (b) selectMark(b.dataset.pick, true); });
  $('dwName').addEventListener('input', function () { var g = G(); g.name = $('dwName').value || '未命名圖說'; save(); $('tbTitle').textContent = g.name; });
  function zoomAt(z, ox, oy) {   // ox, oy：以舞台左上角為原點的固定點
    var st = $('dwStage'), cv = $('dwCanvas'), w0 = cv.offsetWidth, h0 = cv.offsetHeight;
    var fx = (st.scrollLeft + ox - cv.offsetLeft) / w0, fy = (st.scrollTop + oy - cv.offsetTop) / h0;
    dwZoom = Math.max(1, Math.min(10, z)); sizeDwg();
    st.scrollLeft = fx * cv.offsetWidth + cv.offsetLeft - ox; st.scrollTop = fy * cv.offsetHeight + cv.offsetTop - oy;
  }
  function zoomTo(z) { var st = $('dwStage'); zoomAt(z, st.clientWidth / 2, st.clientHeight / 2); }
  $('dwZoomIn').addEventListener('click', function () { zoomTo(dwZoom * 1.5); });
  $('dwZoomOut').addEventListener('click', function () { zoomTo(dwZoom / 1.5); });
  $('dwZoomFit').addEventListener('click', function () { zoomTo(1); });
  $('dwStage').addEventListener('wheel', function (e) {
    if (!e.ctrlKey && !e.metaKey) return; e.preventDefault();
    var b = $('dwStage').getBoundingClientRect(); zoomAt(dwZoom * Math.exp(-e.deltaY * 0.0025), e.clientX - b.left, e.clientY - b.top);
  }, { passive: false });
  var pinch = null;
  function tdist(t) { return Math.hypot(t[0].clientX - t[1].clientX, t[0].clientY - t[1].clientY); }
  $('dwStage').addEventListener('touchstart', function (e) {
    if (e.touches.length !== 2) return;
    if (drag) { if (drag.el) drag.el.remove(); if (drag.t === 'move' || drag.t === 'size') drawMarks(); drag = null; }
    var b = $('dwStage').getBoundingClientRect();
    pinch = { d: tdist(e.touches), z: dwZoom, b: b, cx: (e.touches[0].clientX + e.touches[1].clientX) / 2 - b.left, cy: (e.touches[0].clientY + e.touches[1].clientY) / 2 - b.top };
    var st = $('dwStage'), cv = $('dwCanvas');
    pinch.fx = (st.scrollLeft + pinch.cx - cv.offsetLeft) / cv.offsetWidth; pinch.fy = (st.scrollTop + pinch.cy - cv.offsetTop) / cv.offsetHeight;
    e.preventDefault();
  }, { passive: false });
  $('dwStage').addEventListener('touchmove', function (e) {
    if (!pinch || e.touches.length !== 2) return; e.preventDefault();
    var st = $('dwStage'), cv = $('dwCanvas'), mx = (e.touches[0].clientX + e.touches[1].clientX) / 2 - pinch.b.left, my = (e.touches[0].clientY + e.touches[1].clientY) / 2 - pinch.b.top;
    dwZoom = Math.max(1, Math.min(10, pinch.z * tdist(e.touches) / pinch.d)); sizeDwg();
    st.scrollLeft = pinch.fx * cv.offsetWidth + cv.offsetLeft - mx; st.scrollTop = pinch.fy * cv.offsetHeight + cv.offsetTop - my;
  }, { passive: false });
  $('dwStage').addEventListener('touchend', function (e) { if (e.touches.length < 2) pinch = null; });
  window.addEventListener('resize', function () { if (cur.v === 'dwg' && G()) sizeDwg(); });

  /* 拖曳：畫新色塊、移動、調整大小 */
  var drag = null;
  function frac(e) { var b = $('dwCanvas').getBoundingClientRect(); return { x: Math.max(0, Math.min(1, (e.clientX - b.left) / b.width)), y: Math.max(0, Math.min(1, (e.clientY - b.top) / b.height)) }; }
  $('dwCanvas').addEventListener('pointerdown', function (e) {
    if (e.button > 0 || pinch || !e.isPrimary) return;
    var pt = frac(e), mk = e.target.closest('[data-mk]'), m;
    if (e.target.classList.contains('dw-h') && (m = curMark())) drag = { t: 'size', m: m, p0: pt, o: { w: m.w, h: m.h, x: m.x, y: m.y } };
    else if (mk && mk.dataset.mk === dwSel && (m = curMark())) drag = { t: 'move', m: m, p0: pt, o: { x: m.x, y: m.y } };
    else if (mk && dwMode === 'view') { selectMark(mk.dataset.mk); return; }
    else if (dwMode !== 'view') { drag = { t: 'new', s: dwMode, p0: pt, el: document.createElement('div') }; drag.el.className = 'dw-temp'; $('dwMarks').appendChild(drag.el); }
    else { drag = { t: 'tap', p0: pt }; return; }
    e.preventDefault(); $('dwCanvas').setPointerCapture(e.pointerId);
  });
  $('dwCanvas').addEventListener('pointermove', function (e) {
    if (!drag || drag.t === 'tap') return;
    var pt = frac(e), dx = pt.x - drag.p0.x, dy = pt.y - drag.p0.y, m = drag.m, el;
    if (drag.t === 'new') {
      drag.r = drag.s === 'arrow' ? { s: 'arrow', x: drag.p0.x, y: drag.p0.y, w: dx, h: dy, c: dwColor, rid: cur.link || '' }
        : { s: drag.s, x: Math.min(pt.x, drag.p0.x), y: Math.min(pt.y, drag.p0.y), w: Math.abs(dx), h: Math.abs(dy), c: dwColor, rid: cur.link || '' };
      drag.el.innerHTML = markHtml(drag.r, { sel: false, ratio: gRatio() }); return;
    }
    var arrow = m.s === 'arrow';
    if (drag.t === 'move') {
      m.x = Math.max(Math.max(0, -m.w), Math.min(1 - Math.max(0, m.w), drag.o.x + dx)); m.y = Math.max(Math.max(0, -m.h), Math.min(1 - Math.max(0, m.h), drag.o.y + dy));
    } else if (arrow) { m.w = Math.max(-m.x, Math.min(1 - m.x, drag.o.w + dx)); m.h = Math.max(-m.y, Math.min(1 - m.y, drag.o.h + dy)); }
    else { m.w = Math.max(0.005, Math.min(1 - m.x, drag.o.w + dx)); m.h = Math.max(0.005, Math.min(1 - m.y, drag.o.h + dy)); }
    drag.moved = true;
    if ((el = document.querySelector('#dwMarks [data-mk="' + m.id + '"]'))) { var g0 = G(); el.outerHTML = oneMark(proj(cur.pid), g0, m, g0.marks.indexOf(m)); }
  });
  function endDrag(e) {
    if (!drag) return; var d = drag; drag = null;
    if (d.t === 'tap') { var pt = frac(e); if (Math.abs(pt.x - d.p0.x) + Math.abs(pt.y - d.p0.y) < 0.01 && dwSel) selectMark(null); return; }
    if (d.t === 'new') {
      var g = G(), r = d.r;
      if (!r || (Math.abs(r.w) < 0.006 && Math.abs(r.h) < 0.006)) { d.el.remove(); return; }
      var nm = r.s === 'arrow' ? { id: uid(), s: 'arrow', x: r.x, y: r.y, w: r.w, h: r.h, c: dwColor, t: '', rid: cur.link || '', created: Date.now() }
        : { id: uid(), s: r.s === 'ellipse' ? 'ellipse' : undefined, x: r.x, y: r.y, w: Math.max(r.w, 0.006), h: Math.max(r.h, 0.006), c: dwColor, t: '', rid: cur.link || '', created: Date.now() };
      if (!nm.s) delete nm.s;
      d.el.remove();
      g.marks.push(nm); save(); dwSel = nm.id; dwEditing = true; drawMarks(); dwPanel(); return;
    }
    if (d.moved) save();
  }
  $('dwCanvas').addEventListener('pointerup', endDrag);
  $('dwCanvas').addEventListener('pointercancel', function () { if (drag && drag.el) drag.el.remove(); drag = null; });

  $('dwRotate').addEventListener('click', function () {
    var g = G(), old = g.img;
    loadImg(old).then(function (u) {
      if (!u) return;
      var img = new Image();
      img.onload = function () {
        var c = document.createElement('canvas'); c.width = img.height; c.height = img.width;
        var x = c.getContext('2d'); x.translate(c.width, 0); x.rotate(Math.PI / 2); x.drawImage(img, 0, 0);
        var data = canvasJpeg(c), id = 'dw_' + uid(); thumbCache[id] = data;
        IDB.put(id, data).then(function () {
          g.marks.forEach(function (m) {
            if (m.s === 'arrow') { var ax = 1 - m.y, ay = m.x, aw = -m.h, ah = m.w; m.x = ax; m.y = ay; m.w = aw; m.h = ah; return; }
            var nx = 1 - (m.y + m.h), ny = m.x, w = m.w; m.x = nx; m.y = ny; m.w = m.h; m.h = w;
          });
          g.img = id; var t = g.w; g.w = g.h; g.h = t; save();
          IDB.del(old).catch(function () {}); delete thumbCache[old];
          if (window.Cloud) window.Cloud.uploadPhoto(id, data);
          renderDwg();
        });
      };
      img.src = u;
    });
  });
  $('dwDel').addEventListener('click', function () {
    var p = proj(cur.pid), g = G();
    ask('刪除圖說「' + g.name + '」？', '圖說與 ' + g.marks.length + ' 個色塊會移到工程頁最下方的「最近刪除」，7 天內可以復原。', '取消', '刪除圖說', function () {
      p.dwgs = dwgs(p).filter(function (x) { return x !== g; }); toTrash(p, 'dwg', g); save();
      history.back();
    });
  });
  $('dwPrint').addEventListener('click', function () { go({ v: 'dwgPrint', pid: cur.pid, did: cur.did }); });

  /* 列印：整張圖說（附色塊說明），或附在抽查表後的「抽查位置圖」 */
  function dwFigure(g, marks, land, reserve, numOf) {
    var aw = land ? 277 : 186, ah = Math.max(70, (land ? 190 : 270) - reserve);
    var fw = Math.min(aw, ah * g.w / g.h), fh = fw * g.h / g.w;
    return '<div class="dw-fig" style="width:' + fw.toFixed(1) + 'mm;height:' + fh.toFixed(1) + 'mm"><img alt="" data-img="' + g.img + '">' +
      marks.map(function (m) { return markHtml(m, { ratio: g.h / g.w, tag: numOf(m) }); }).join('') + '</div>';
  }
  function stKey() {
    return '<p class="dw-stkey">' + ['ok', 'ng', 'fixed', 'todo'].map(function (k) { return '<span><i class="dw-key" style="--c:' + ST_COLOR[k] + ';--bg:' + rgba(ST_COLOR[k], 0.4) + '"></i>' + ST_NAME[k] + '</span>'; }).join('') + '</p>';
  }
  function dwLegend(p, marks, numOf) {
    return '<table class="dw-legend"><colgroup><col style="width:9%"><col style="width:33%"><col style="width:58%"></colgroup><thead><tr><th>編號</th><th>標示</th><th>對應抽查紀錄</th></tr></thead><tbody>' +
      marks.map(function (m) {
        var li = linkInfo(p, m.rid);
        return '<tr><td class="c"><i class="dw-key" style="--c:' + colorOf(m) + ';--bg:' + rgba(colorOf(m), 0.4) + '"></i>' + numOf(m) + '</td><td>' + esc(m.t || '') + '</td><td>' +
          (li ? esc(li.long) + '（' + ST_NAME[recState(li.w, li.r)] + '）' + esc(li.r.location ? '，' + li.r.location : '') + esc(li.r.checkDate ? '，' + roc(li.r.checkDate, '.') : '') : '') + '</td></tr>';
      }).join('') + '</tbody></table>';
  }
  function renderDwgPrint() {
    var p = proj(cur.pid), g = G(), land = g.w > g.h, num = function (m) { return g.marks.indexOf(m) + 1; };
    var ms = g.marks.filter(function (m) { return passFilter(p, m); }), fd = filterDesc();
    topbar('列印預覽', g.name, '圖說');
    bindOut(g, 'outAt'); $('btnMovePv').hidden = true; $('pvParts').hidden = true; $('outAt').closest('.out-at').hidden = false;
    var paper = $('paper'); paper.className = 'paper dw-paper' + (land ? ' landscape' : '');
    var rows = Math.min(ms.length, 30);
    paper.innerHTML = '<h3 class="dw-h3">抽查位置圖</h3><p class="no">工程名稱：' + esc(p.info.name) + '　圖說：' + esc(g.name) + (fd ? '　篩選：' + esc(fd) : '') + '</p>' +
      dwFigure(g, ms, land, 26 + (rows ? 8 + rows * 5.2 : 0), num) + stKey() + (ms.length ? dwLegend(p, ms, num) : '') + foot();
    fillImgs(paper);
    $('printHint').textContent = '紙張選 A4 ' + (land ? '橫向' : '直式') + '、邊界「預設」。「另存 PDF」會直接下載檔案，不經過列印視窗。';
    show('vPreview');
    setDocName('抽查位置圖_' + g.name + '_' + roc(today(), '.'));
  }
  /* 抽查表列印時，附上標記本表的圖說 */
  function recDwgPages(p, wid, r) {
    var link = wid + '|' + r.id, out = '';
    dwgs(p).forEach(function (g) {
      var ms = g.marks.filter(function (m) { return m.rid === link; }); if (!ms.length) return;
      var num = function (m) { return ms.indexOf(m) + 1; };
      out += '<section class="ph-page dw-page"><h3>抽查位置圖 ' + esc(r.docNo) + '</h3><p class="no">圖說：' + esc(g.name) + (r.location ? '　檢查位置：' + esc(r.location) : '') + '</p>' +
        dwFigure(g, ms, false, 30 + (ms.some(function (m) { return m.t; }) ? 8 + ms.length * 5.2 : 0), num) +
        (ms.some(function (m) { return m.t; }) ? dwLegend(p, ms, num) : '') + '</section>';
    });
    return out;
  }
  $('btnPin').addEventListener('click', function () {
    var p = proj(cur.pid), list = dwgs(p), link = cur.wid + '|' + cur.rid, ls = linkedMarks(p, link);
    var target = ls.length ? ls[0].g : list.length === 1 ? list[0] : null;
    if (target) go({ v: 'dwg', pid: p.id, did: target.id, link: link }); else go({ v: 'dwgs', pid: p.id, link: link });
  });
  function pinInfo() {
    var p = proj(cur.pid), ls = linkedMarks(p, cur.wid + '|' + cur.rid);
    $('pinInfo').textContent = ls.length ? '已標記 ' + ls.length + ' 處，列印時附上抽查位置圖' : dwgs(p).length ? '' : '尚未匯入圖說';
  }

  function storeInfo() {
    var mb = function (n) { return (n / 1048576).toFixed(n < 1048576 ? 2 : 1) + ' MB'; };
    var base = '本機資料 ' + mb(stateSize || JSON.stringify(S).length) + (idbOK ? '（存於 IndexedDB，沒有 5 MB 上限）' : '（瀏覽器不支援 IndexedDB，上限約 5 MB）');
    $('storeInfo').textContent = base;
    if (navigator.storage && navigator.storage.estimate) navigator.storage.estimate().then(function (e) {
      $('storeInfo').textContent = base + '；連同照片與圖說共用 ' + mb(e.usage || 0) + (e.quota ? '／可用 ' + mb(e.quota) : '');
    }).catch(function () {});
  }
  /* ========== 最近刪除（回收桶）：保留 7 天，到期自動永久刪除 ========== */
  var TRASH_DAYS = 7, DAY = 86400000;
  function toTrash(p, type, data, extra) {
    (p.trash || (p.trash = [])).push(Object.assign({ id: uid(), type: type, at: Date.now(), data: data }, extra || {}));
  }
  function trashImgs(t) {
    var d = t.data, ids = [];
    if (t.type === 'rec') (d.photos || []).forEach(function (x) { ids.push(x.id); });
    if (t.type === 'dwg') ids.push(d.img);
    if (t.type === 'proj') {
      Object.keys(d.recs || {}).forEach(function (w) { (d.recs[w] || []).forEach(function (r) { (r.photos || []).forEach(function (x) { ids.push(x.id); }); }); });
      (d.dwgs || []).forEach(function (g) { ids.push(g.img); });
      (d.trash || []).forEach(function (x) { ids = ids.concat(trashImgs(x)); });
    }
    return ids;
  }
  function purgeTrash() {
    var limit = Date.now() - TRASH_DAYS * DAY, changed = false;
    var keep = function (list) { return (list || []).filter(function (t) { if (t.at >= limit) return true; trashImgs(t).forEach(function (id) { IDB.del(id).catch(function () {}); }); changed = true; return false; }); };
    S.trash = keep(S.trash);
    S.projects.forEach(function (p) { if (p.trash && p.trash.length) p.trash = keep(p.trash); });
    if (changed) save();
  }
  function trashLabel(t) {
    var d = t.data, w = t.wid && WORK[t.wid];
    if (t.type === 'rec') return { tag: '抽查表', name: (w ? (w.kind === 'phased' ? w.name + '(' + (d.phase + 1) + ')' : w.name) + ' ' : '') + d.docNo, sub: [d.checkDate && roc(d.checkDate, '.'), d.location].filter(Boolean).join(' · ') };
    if (t.type === 'mat') return { tag: '材料', name: d.name || '未命名材料', sub: d.no || '' };
    if (t.type === 'dwg') return { tag: '圖說', name: d.name, sub: d.marks.length + ' 個色塊' };
    var n = 0; Object.keys(d.recs || {}).forEach(function (k) { n += d.recs[k].length; });
    return { tag: '工程', name: d.info.name || '未命名工程', sub: n + ' 份紀錄' };
  }
  function trashHtml(list) {
    var now = Date.now();
    return list.slice().sort(function (a, b) { return b.at - a.at; }).map(function (t) {
      var L = trashLabel(t), left = Math.max(1, Math.ceil((t.at + TRASH_DAYS * DAY - now) / DAY));
      return '<li class="dl-item"><div class="dl-row tr-row"><span class="dl-tag">' + L.tag + '</span><span class="dl-txt"><b>' + esc(L.name) + '</b><small>' +
        esc([L.sub, left + ' 天後永久刪除'].filter(Boolean).join(' · ')) + '</small></span>' +
        '<button type="button" class="btn ghost sm" data-tr-restore="' + t.id + '">復原</button>' +
        '<button type="button" class="link tr-del" data-tr-del="' + t.id + '" aria-label="永久刪除">✕</button></div></li>';
    }).join('');
  }
  function renderTrash(box, list) {
    var b = $(box); b.hidden = !list || !list.length; if (b.hidden) return;
    b.querySelector('.tr-n').textContent = list.length;
    b.querySelector('ul').innerHTML = trashHtml(list);
  }
  function restoreTrash(p, t) {
    var d = t.data;
    if (t.type === 'rec') {
      var list = recs(p, t.wid);
      if (list.some(function (r) { return r.phase === d.phase && r.docNo === d.docNo; })) {
        d.seq = list.filter(function (r) { return r.phase === d.phase; }).reduce(function (m, r) { return Math.max(m, r.seq); }, 0) + 1;
        var w = WORK[t.wid]; d.docNo = (p.info.code ? p.info.code + '-' : '') + w.code + (w.kind === 'phased' ? (d.phase + 1) : '') + '-' + ('00' + d.seq).slice(-3);
      }
      list.push(d); return '已復原「' + d.docNo + '」';
    }
    if (t.type === 'mat') { p.mats.push(d); return '已復原材料「' + (d.name || '') + '」'; }
    if (t.type === 'dwg') { dwgs(p).push(d); return '已復原圖說「' + d.name + '」'; }
  }
  $('trashBox').addEventListener('click', function (e) {
    var p = proj(cur.pid), b = e.target.closest('[data-tr-restore],[data-tr-del]'); if (!b || !p) return;
    var id = b.dataset.trRestore || b.dataset.trDel, t = (p.trash || []).filter(function (x) { return x.id === id; })[0]; if (!t) return;
    if (b.dataset.trRestore) {
      p.trash = p.trash.filter(function (x) { return x !== t; }); var msg = restoreTrash(p, t); save(); renderProject(); $('trashBox').open = true; $('trashMsg').textContent = msg + '。';
    } else ask('永久刪除？', '「' + trashLabel(t).name + '」刪除後無法再復原。', '取消', '永久刪除', function () {
      p.trash = p.trash.filter(function (x) { return x !== t; }); trashImgs(t).forEach(function (i) { IDB.del(i).catch(function () {}); }); save(); renderProject(); $('trashBox').open = true;
    });
  });
  $('trashBoxP').addEventListener('click', function (e) {
    var b = e.target.closest('[data-tr-restore],[data-tr-del]'); if (!b) return;
    var id = b.dataset.trRestore || b.dataset.trDel, t = (S.trash || []).filter(function (x) { return x.id === id; })[0]; if (!t) return;
    if (b.dataset.trRestore) {
      S.trash = S.trash.filter(function (x) { return x !== t; });
      if (!proj(t.data.id)) S.projects.push(t.data); save(); renderProjects(); $('trashBoxP').open = true; $('trashMsgP').textContent = '已復原工程「' + (t.data.info.name || '') + '」。';
    } else ask('永久刪除？', '「' + trashLabel(t).name + '」與底下所有紀錄刪除後無法再復原。', '取消', '永久刪除', function () {
      S.trash = S.trash.filter(function (x) { return x !== t; }); trashImgs(t).forEach(function (i) { IDB.del(i).catch(function () {}); }); save(); renderProjects(); $('trashBoxP').open = true;
    });
  });

  /* ========== 全域搜尋：Ctrl＋K、⌘K 或 / 開啟；手機按頂列放大鏡 ========== */
  var srItems = [], srAct = 0, srRes = [];
  function srIndex() {
    var out = [];
    S.projects.forEach(function (p) {
      var pn = p.info.name || '未命名工程';
      out.push({ pid: p.id, pn: pn, tag: '工程', title: pn, sub: [p.info.code, p.info.contractor].filter(Boolean).join(' · '), text: [pn, p.info.code, p.info.owner, p.info.contractor, p.info.supervisor].join(' '), go: { v: 'project', pid: p.id } });
      Object.keys(p.recs || {}).forEach(function (wid) {
        var w = WORK[wid]; if (!w) return;
        (p.recs[wid] || []).forEach(function (r) {
          if (r.draft) return;
          var m = w.kind === 'matform' && matById(p, r.mid), notes = [];
          Object.keys(r.res || {}).forEach(function (k) { var c = r.res[k]; if (c && c.note) notes.push(c.note); });
          Object.keys(r.extra || {}).forEach(function (g) { (r.extra[g] || []).forEach(function (x) { notes.push(x.name, x.note); }); });
          (r.rows || []).forEach(function (x) { notes.push(x.name, x.note); });
          (r.photos || []).forEach(function (x) { notes.push(x.cap); });
          notes.push(r.remark, r.methodNote);
          var title = (m ? '材料抽查：' + m.name : w.kind === 'phased' ? w.name + '(' + (r.phase + 1) + ')' : w.name) + ' ' + r.docNo;
          out.push({ pid: p.id, pn: pn, tag: '抽查表', title: title, sub: [r.checkDate && roc(r.checkDate, '.'), r.location].filter(Boolean).join(' · '),
            text: [title, r.location, r.checkDate, roc(r.checkDate || '', '.'), notes.filter(Boolean).join(' ')].join(' '), go: { v: 'form', pid: p.id, wid: wid, rid: r.id } });
        });
      });
      (p.mats || []).forEach(function (m) {
        out.push({ pid: p.id, pn: pn, tag: '材料', title: m.name || '未命名材料', sub: [m.no, MSTAT[matStatus(m)]].filter(Boolean).join(' · '), text: [m.name, m.no, m.qty, m.remark, m.inspNote].join(' '), go: { v: 'matItem', pid: p.id, mid: m.id } });
      });
      (p.dwgs || []).forEach(function (g) {
        var labels = g.marks.map(function (m) { var li = linkInfo(p, m.rid); return [m.t, li && li.short].filter(Boolean).join(' '); }).join(' ');
        out.push({ pid: p.id, pn: pn, tag: '圖說', title: g.name, sub: g.marks.length + ' 個標記', text: [g.name, labels].join(' '), go: { v: 'dwg', pid: p.id, did: g.id } });
      });
    });
    return out;
  }
  function srSnip(text, toks) {
    var low = text.toLowerCase(), i = -1;
    toks.some(function (t) { i = low.indexOf(t); return i >= 0; });
    if (i < 0) return '';
    var a = Math.max(0, i - 12), s0 = text.slice(a, i + 28).replace(/\s+/g, ' ');
    return (a ? '…' : '') + s0 + (i + 28 < text.length ? '…' : '');
  }
  function srRender() {
    var q = $('srInput').value.trim().toLowerCase(), toks = q.split(/\s+/).filter(Boolean), only = $('srScope').checked && cur.pid;
    srRes = !toks.length ? [] : srItems.filter(function (it) {
      if (only && it.pid !== cur.pid) return false;
      var h = (it.title + ' ' + it.sub + ' ' + it.text).toLowerCase();
      return toks.every(function (t) { return h.indexOf(t) >= 0; });
    }).sort(function (a, b) { var ta = toks.every(function (t) { return (a.title + a.sub).toLowerCase().indexOf(t) >= 0; }), tb = toks.every(function (t) { return (b.title + b.sub).toLowerCase().indexOf(t) >= 0; }); return (tb - ta); }).slice(0, 60);
    srAct = 0;
    var last = null, html = '';
    srRes.forEach(function (it, i) {
      if (it.pid !== last && !only) { html += '<li class="sr-grp">' + esc(it.pn) + '</li>'; last = it.pid; }
      var inTitle = toks.every(function (t) { return (it.title + it.sub).toLowerCase().indexOf(t) >= 0; });
      html += '<li class="sr-item' + (i === srAct ? ' act' : '') + '" data-sr="' + i + '" role="option"><span class="dl-tag">' + it.tag + '</span><span class="dl-txt"><b>' + esc(it.title) + '</b><small>' +
        esc(inTitle ? it.sub : srSnip(it.text, toks) || it.sub) + '</small></span></li>';
    });
    $('srList').innerHTML = toks.length ? (html || '<li class="sr-none">找不到「' + esc(q) + '」</li>') : '<li class="sr-none">輸入關鍵字，例如編號「RB2」、位置「3F」、日期「115.09」或備註文字。多個關鍵字以空白分隔。</li>';
  }
  function srMove(d) {
    if (!srRes.length) return;
    srAct = (srAct + d + srRes.length) % srRes.length;
    $('srList').querySelectorAll('.sr-item').forEach(function (x) { x.classList.toggle('act', +x.dataset.sr === srAct); if (+x.dataset.sr === srAct) x.scrollIntoView({ block: 'nearest' }); });
  }
  function srOpen() {
    srItems = srIndex(); $('srScopeBox').hidden = !cur.pid; if (!cur.pid) $('srScope').checked = false;
    $('srDlg').hidden = false; $('srInput').select(); $('srInput').focus(); srRender();
  }
  function srClose() { $('srDlg').hidden = true; }
  function srGo(i) { var it = srRes[i]; if (!it) return; srClose(); go(it.go); }
  $('btnSearch').addEventListener('click', srOpen);
  $('srClose').addEventListener('click', srClose);
  $('srDlg').addEventListener('click', function (e) { if (e.target === $('srDlg')) srClose(); var li = e.target.closest('[data-sr]'); if (li) srGo(+li.dataset.sr); });
  $('srInput').addEventListener('input', srRender);
  $('srScope').addEventListener('change', srRender);
  $('srInput').addEventListener('keydown', function (e) {
    if (e.key === 'ArrowDown') { e.preventDefault(); srMove(1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); srMove(-1); }
    else if (e.key === 'Enter') { e.preventDefault(); srGo(srAct); }
    else if (e.key === 'Escape') { e.preventDefault(); srClose(); }
  });
  document.addEventListener('keydown', function (e) {
    var typing = /^(INPUT|TEXTAREA|SELECT)$/.test((e.target.tagName || '')) || e.target.isContentEditable;
    if ((e.ctrlKey || e.metaKey) && (e.key === 'k' || e.key === 'K')) { e.preventDefault(); if ($('srDlg').hidden) srOpen(); else srClose(); }
    else if (e.key === '/' && !typing && $('srDlg').hidden) { e.preventDefault(); srOpen(); }
    else if (e.key === 'Escape' && !$('srDlg').hidden) srClose();
  });

  /* ---------- 對話框 ---------- */
  var dlgYesFn = null;
  function ask(title, body, no, yes, fn) {
    $('dlgTitle').textContent = title; $('dlgBody').textContent = body;
    $('dlgNo').textContent = no; $('dlgYes').textContent = yes; dlgYesFn = fn;
    $('dlg').hidden = false; $('dlgNo').focus();
  }
  $('dlgNo').addEventListener('click', function () { $('dlg').hidden = true; });
  $('dlgYes').addEventListener('click', function () { $('dlg').hidden = true; if (dlgYesFn) dlgYesFn(); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && !$('dlg').hidden) $('dlg').hidden = true; });

  /* 抽查項目改用固定代碼：第一次開啟時自動轉換舊資料，轉換前先在本機另存一份備份 */
  function migrateBoot() {
    var need = S.projects.some(function (p) { return Object.keys(p.recs || {}).some(function (wid) { return (p.recs[wid] || []).some(function (r) { return hasStd(WORK[wid]) && r.kv !== 2; }); }); });
    if (!need) { if (migrateAll()) writeLocal(); return; }
    try { IDB.put('bak_kv2_' + Date.now(), JSON.stringify(S)).catch(function () {}); } catch (e) {}
    migrateAll(); writeLocal();
  }
  window.APP = {
    state: function () { return S; }, setState: function (x) { S = x; }, D: D, cur: function () { return cur; },
    persist: function () { writeLocal(); },
    rerender: function () { render(); }, ask: ask,
    fixRec: function (wid, r) { return migrateRec(WORK[wid], r); }, migrate: migrateAll
  };
  /* 啟動：讀取 IndexedDB 中較新的資料後再開始 */
  function boot() {
    S.projects.forEach(function (p) { if (!p.mats) p.mats = []; });
    migrateBoot();
    go({ v: 'projects' }, true);
    if (window.Cloud) window.Cloud.init();
  }
  (idbOK ? IDB.get('state:v1').then(function (txt) {
    if (!txt) return;
    var st = JSON.parse(txt);
    if (st && Array.isArray(st.projects) && (st.savedAt || 0) >= (S.savedAt || 0)) S = st;
  }).catch(function () {}) : Promise.resolve()).then(boot, boot);
  window.addEventListener('beforeunload', function (e) {
    flushLocal();
    var n = window.Cloud && window.Cloud.unsynced ? window.Cloud.unsynced() : 0;
    if (n) { e.preventDefault(); e.returnValue = '還有 ' + n + ' 筆資料尚未上傳到雲端'; return e.returnValue; }
  });

  if (!window.Cloud && 'serviceWorker' in navigator && /^https?:/.test(location.protocol)) {
    navigator.serviceWorker.register('sw.js').catch(function () {});
  }
})();
