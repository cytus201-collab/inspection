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
    { k: 'actualEnd', l: '實際完工日期', t: 'date' },
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
    { k: 'planUse', l: '預定施工日期', t: 'date' },
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
  function refText(code) { return code ? D.refs[code] || code : ''; }
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
        items.push({ key: key, name: c.name || it.name, std: c.std != null ? c.std : it.std, def: it.std, hint: c.hint != null ? c.hint : it.hint, ref: c.ref != null ? c.ref : it.ref, c: c });
      });
      (r.extra[g.id] || []).forEach(function (x, xi) { items.push({ x: g.id + '|' + xi, name: x.name, std: x.std, hint: x.hint || '', ref: x.ref || '', c: x, custom: true }); });
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
  var VIEWS = ['vProjects', 'vProject', 'vWork', 'vForm', 'vMaterials', 'vMatItem', 'vMatImport', 'vPreview', 'vDwgs', 'vDwg', 'vStd', 'vDocs', 'vDocEdit', 'vDaily', 'vDailyEdit', 'vField', 'vDrImport', 'vWorks'];
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
    if (!/^(projects|field)$/.test(cur.v) && !p) cur = { v: 'projects' };
    else if (/^(work|form|preview)$/.test(cur.v) && !w) cur = { v: 'project', pid: cur.pid };
    else if (/^(form|preview)$/.test(cur.v) && !recById(p, cur.wid, cur.rid)) cur = { v: 'work', pid: cur.pid, wid: cur.wid };
    else if (cur.v === 'matItem' && !matById(p, cur.mid)) cur = { v: 'materials', pid: cur.pid };
    else if (cur.v === 'matImport' && !pendingImport) cur = { v: 'materials', pid: cur.pid };
    else if (/^(dwg|dwgPrint)$/.test(cur.v) && !dwgById(p, cur.did)) cur = { v: 'dwgs', pid: cur.pid };
    else if (/^(docEdit|docPrint)$/.test(cur.v) && !docById(p, cur.did)) cur = { v: 'docs', pid: cur.pid };
    $('sumBar').hidden = true;
    /* 公文列印需要零邊界（Safari 不支援具名頁面，所以只在公文預覽時加上全域 @page） */
    if ($('gwFontBar') && !/^(docPrint|dailyPrint)$/.test(cur.v)) $('gwFontBar').hidden = true;
    if (cur.v !== 'docPrint') $('gwBar').hidden = true;
    var pgs = $('gwPage'); if (/^(docPrint|dailyPrint)$/.test(cur.v) && !pgs) { pgs = document.createElement('style'); pgs.id = 'gwPage'; pgs.textContent = '@page{size:A4 portrait;margin:0}'; document.head.appendChild(pgs); } else if (!/^(docPrint|dailyPrint)$/.test(cur.v) && pgs) pgs.remove();
    ({
      projects: renderProjects, project: renderProject, work: renderWork, form: renderForm, preview: renderPreview,
      materials: renderMaterials, matItem: renderMatItem, matImport: renderMatImport, matPrint: renderMatPrint, sumPrint: renderSumPrint,
      dwgs: renderDwgs, dwg: renderDwg, dwgPrint: renderDwgPrint, std: renderStd, docs: renderDocs, docEdit: renderDocEdit, docPrint: renderDocPrint, daily: renderDaily, dailyEdit: renderDailyEdit, dailyPrint: renderDailyPrint, field: renderField, works: renderWorks, drImport: renderDrImport
    })[cur.v]();
    coEdit(); markNav();
    if (window.Cloud && window.Cloud.viewChanged) window.Cloud.viewChanged();
  }
  /* 共編：同一張表有同事開著時顯示提醒 */
  function coEdit() {
    var id = cur.v === 'form' ? cur.rid : cur.v === 'matItem' ? cur.mid : '';
    var who = id && window.Cloud && window.Cloud.editors ? window.Cloud.editors(id) : [];
    ['coEditForm', 'coEditMat'].forEach(function (bid) {
      var el = $(bid); if (!el) return;
      el.hidden = !who.length;
      if (who.length) el.textContent = '⚠ ' + who.join('、') + ' 也正在開啟這張表。兩人同時修改時，較晚存檔的會覆蓋較早的，請先協調由一人填寫。';
    });
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
  /* ========== 辦公室／現場兩種模式、導覽列 ========== */
  function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
  function uiMode() { return lsGet('ui-mode') || (window.matchMedia && window.matchMedia('(max-width: 760px)').matches ? 'field' : 'office'); }
  function curProj() { var id = cur.pid || lsGet('last-pid'); return (id && proj(id)) || S.projects[0] || null; }
  function wkStart(d) { var t = new Date(d + 'T00:00:00'), w = (t.getDay() + 6) % 7; return addDays(d, -w); }
  function projDay(p, dt) { var I = drInfo(p); return I.start ? drDays(I.start, dt) + 1 : 0; }
  function lastAct(p) { var xs = dailies(p).filter(function (d) { return d.act !== '' && d.act != null && d.date <= today(); }).sort(function (a, b) { return a.date < b.date ? 1 : -1; }); return xs.length ? +xs[0].act : null; }
  /* 工期內：開工日起，至實際完工日（未完工則持續，逾期也要每日填報） */
  function inPeriod(p, dt) { var I = drInfo(p); return !!I.start && dt >= I.start && (!p.info.actualEnd || dt <= p.info.actualEnd); }
  function plansOn(p, dt) { return (p.plans || []).filter(function (pl) { return pl.date === dt && !(pl.rid && recById(p, pl.wid, pl.rid) && !recById(p, pl.wid, pl.rid).draft); }); }
  function openDefs(p) { return sumData(p, '9999-12-31').defs.filter(function (x) { return x.st !== 'fixed'; }); }
  document.addEventListener('click', function (e) {
    var b = e.target.closest('[data-mode]'); if (!b) return;
    lsSet('ui-mode', b.dataset.mode); go({ v: b.dataset.mode === 'field' ? 'field' : 'projects' });
  });
  function navTo(k) {
    var p = curProj();
    if (k === 'office') return go({ v: 'projects', pid: p && p.id });
    if (k === 'settings') { go({ v: 'projects', pid: p && p.id }); setTimeout(function () { $('ovSettings').scrollIntoView({ block: 'start' }); }, 30); return; }
    if (k === 'field') return go({ v: 'field' });
    if (!p) return go({ v: 'projects' });
    if (k === 'project') return go({ v: 'project', pid: p.id });
    if (k === 'works') return go({ v: 'works', pid: p.id });
    if (k === 'docs') return go({ v: 'docs', pid: p.id });
    if (k === 'daily') return go({ v: 'daily', pid: p.id, ym: today().slice(0, 7) });
    if (k === 'materials') return go({ v: 'materials', pid: p.id });
  }
  ['rail', 'tabs'].forEach(function (id) { $(id).addEventListener('click', function (e) { var b = e.target.closest('[data-nav]'); if (b) navTo(b.dataset.nav); }); });
  function markNav() {
    var v = cur.v, k = v === 'projects' ? 'office' : v === 'field' ? 'field' : /^doc/.test(v) ? 'docs' : /^(daily|drImport)/.test(v) ? 'daily' : /^mat/.test(v) ? 'materials' : /^(works|work|form|preview|std|dwgs|dwg)/.test(v) ? 'works' : 'project';
    document.querySelectorAll('#rail [data-nav], #tabs [data-nav]').forEach(function (b) { b.classList.toggle('on', b.dataset.nav === k); });
    if (cur.pid) lsSet('last-pid', cur.pid);
  }
  function startPlan(p, plid) {
    var pl = plan(p, plid), w2 = pl && WORK[pl.wid]; if (!w2) return;
    var r = newRecord(p, w2, w2.kind === 'phased' ? pl.phase : 0, { checkDate: pl.date, location: pl.note || '' });
    pl.rid = r.id; save(); go({ v: 'form', pid: p.id, wid: w2.id, rid: r.id });
  }
  /* ---- 辦公室總覽 ---- */
  var ovFilter = 'all';
  function projOpts(p) { return S.projects.map(function (x) { return '<option value="' + x.id + '"' + (p && x.id === p.id ? ' selected' : '') + '>' + esc(x.info.name || '未命名工程') + '</option>'; }).join(''); }
  function renderProjects() {
    renderSwitch();
    topbar('監造抽查紀錄', '工程明細 · 表單版次 ' + D.version);
    var td = today(), ws = wkStart(td), we = addDays(ws, 6), ym = td.slice(0, 7), wd = '日一二三四五六'[new Date(td + 'T00:00:00').getDay()];
    $('ovDate').textContent = rocYmd(td) + ' 星期' + wd;
    var p = curProj();
    $('ovProj').innerHTML = projOpts(p); $('ovProj').hidden = !p;
    if (!p) {
      $('ovQuick').innerHTML = ''; $('ovDueF').innerHTML = '';
      $('ovKpi').innerHTML = '<div class="empty">還沒有工程。按「新增工程」建立第一個。</div>'; $('ovDue').innerHTML = '';
    } else {
      $('ovQuick').innerHTML = '<button type="button" class="btn ghost sm" data-oq="docs">＋ 發文</button><button type="button" class="btn ghost sm" data-oq="dailyPrint">整月日報列印</button><button type="button" class="btn ghost sm" data-oq="sum">抽查總表／缺失管制</button><button type="button" class="btn ghost sm" data-oq="materials">材料管制</button>';
      var I = drInfo(p), day = projDay(p, td), total = I.dur ? I.dur + (+I.ext || 0) : 0, late = I.start && total && day > total && !p.info.actualEnd;
      var recW = 0, planW = 0, drDone = 0, drNeed = 0, defNext = '';
      Object.keys(p.recs || {}).forEach(function (wid) { (p.recs[wid] || []).forEach(function (r) { if (!r.draft && r.checkDate >= ws && r.checkDate <= we) recW++; }); });
      planW = (p.plans || []).filter(function (pl) { return pl.date >= ws && pl.date <= we && !pl.rid; }).length;
      monthDays(ym).forEach(function (dt) { if (dt <= td && inPeriod(p, dt)) { drNeed++; if (dailyOf(p, dt)) drDone++; } });
      var defs = openDefs(p);
      defs.forEach(function (x) { var dl = defInfo(x, td).dl; if (dl && (!defNext || dl < defNext)) defNext = dl; });
      var a = lastAct(p);
      var tile = function (k, v, s, dark) { return '<div class="kpi' + (dark ? ' dark' : '') + '"><span>' + k + '</span><b>' + v + '</b><small>' + s + '</small></div>'; };
      $('ovKpi').innerHTML =
        tile('工期', day ? '第 ' + day + '<em> 天</em>' : '—', total ? (late ? '已逾工期（共 ' + total + ' 天）' : '共 ' + total + ' 天') : '未填開工日或工期') +
        tile('實際進度', a != null ? a.toFixed(2) + '<em>%</em>' : '—', a != null ? '依施工日誌' : '尚無日報進度') +
        tile('本週抽查', recW, '預定待抽查 ' + planW) +
        tile('本月日報', drDone + '<em> / ' + drNeed + '</em>', drNeed - drDone ? (drNeed - drDone) + ' 天未填' : '已全部填寫') +
        tile('未改善缺失', defs.length, defNext ? '最近期限 ' + (+defNext.slice(5, 7)) + '/' + (+defNext.slice(8)) : defs.length ? '未設期限' : '無', true);
      /* 近 7 天：本工程的會議、發文、缺失期限、預定抽查 */
      var due = [], end = addDays(td, 7);
      (p.docs || []).forEach(function (d) {
        if (d.kind === 'meet' && d.meet && d.meet.time && d.meet.time.slice(0, 10) >= td && d.meet.time.slice(0, 10) <= end) due.push({ dt: d.meet.time.slice(0, 10), k: 'doc', t: '會議：' + shortSubj(d) + ' ' + d.meet.time.slice(11, 16), s: d.meet.place || '', go: { v: 'docEdit', pid: p.id, did: d.id } });
        else if (d.date >= td && d.date <= end) due.push({ dt: d.date, k: 'doc', t: '發文：' + shortSubj(d), s: d.no || '', go: { v: 'docEdit', pid: p.id, did: d.id } });
      });
      defs.forEach(function (x) { var dl = defInfo(x, td).dl; if (dl && dl <= end) due.push({ dt: dl, k: 'ng', t: '缺失改善期限：' + x.name, s: x.loc || '', late: dl < td, go: x.r && x.w ? { v: 'form', pid: p.id, wid: x.w.id, rid: x.r.id } : { v: 'project', pid: p.id } }); });
      (p.plans || []).forEach(function (pl) { if (!pl.rid && pl.date >= td && pl.date <= end) { var w = WORK[pl.wid]; due.push({ dt: pl.date, k: 'rec', t: '預定抽查：' + (w ? workLabel(w, pl.phase || 0) : ''), s: pl.note || '', go: { v: 'project', pid: p.id } }); } });
      due.sort(function (x, y) { return x.dt < y.dt ? -1 : 1; });
      ovDueItems = due;
      $('ovDueF').innerHTML = [['all', '全部'], ['doc', '公文'], ['ng', '缺失'], ['rec', '抽查']].map(function (f) { return '<button type="button" class="cf' + (ovFilter === f[0] ? ' on' : '') + '" data-of="' + f[0] + '">' + f[1] + '</button>'; }).join('');
      var shown = due.filter(function (x) { return ovFilter === 'all' || x.k === ovFilter; });
      $('ovDue').innerHTML = shown.length ? shown.map(function (x) {
        var i = due.indexOf(x);
        return '<li><button type="button" data-due="' + i + '"><span class="dd' + (x.k === 'ng' ? ' dark' : '') + '">' + (+x.dt.slice(5, 7)) + '/' + (+x.dt.slice(8)) + '</span><span class="dt"><b>' + esc(x.t) + '</b>' + (x.s ? '<small>' + esc(x.s) + '</small>' : '') + '</span></button></li>';
      }).join('') : '<li class="muted ov-none">近 7 天沒有期限或會議。</li>';
    }
    $('backupMsg').textContent = saveMsg;
    renderTrash('trashBoxP', S.trash); $('trashMsgP').textContent = '';
    storeInfo();
    show('vProjects');
  }
  $('ovProj').addEventListener('change', function () { lsSet('last-pid', this.value); go({ v: 'projects', pid: this.value }, true); });
  $('pjSel').addEventListener('change', function () { lsSet('last-pid', this.value); go({ v: 'project', pid: this.value }, true); });
  var ovDueItems = [];
  $('ovDueF').addEventListener('click', function (e) { var b = e.target.closest('[data-of]'); if (b) { ovFilter = b.dataset.of; renderProjects(); } });
  $('ovDue').addEventListener('click', function (e) { var b = e.target.closest('[data-due]'); if (b) go(ovDueItems[+b.dataset.due].go); });
  $('ovQuick').addEventListener('click', function (e) {
    var b = e.target.closest('[data-oq]'), p = curProj(); if (!b || !p) return;
    var k = b.dataset.oq;
    if (k === 'docs') go({ v: 'docs', pid: p.id });
    if (k === 'dailyPrint') go({ v: 'dailyPrint', pid: p.id, ym: today().slice(0, 7) });
    if (k === 'sum') { var t = today(); go({ v: 'sumPrint', pid: p.id, to: t, from: sumFromDef(t), kind: 'ng' }); }
    if (k === 'materials') go({ v: 'materials', pid: p.id });
  });
  /* ---- 今日現場 ---- */
  function renderField() {
    var p = curProj(), td = today();
    topbar('監造抽查紀錄', '現場模式', '');
    $('fdDate').textContent = (+td.slice(5, 7)) + '/' + (+td.slice(8)) + ' 星期' + '日一二三四五六'[new Date(td + 'T00:00:00').getDay()];
    $('fdProj').innerHTML = S.projects.map(function (x) { return '<option value="' + x.id + '"' + (p && x.id === p.id ? ' selected' : '') + '>' + esc(x.info.name || '未命名工程') + '</option>'; }).join('');
    if (!p) { $('fdPlan').innerHTML = '<div class="empty">還沒有工程。請先到「明細」新增工程。</div>'; $('fdTwo').innerHTML = ''; $('fdDone').innerHTML = ''; show('vField'); return; }
    var a = lastAct(p), day = projDay(p, td);
    $('fdProjSub').textContent = [day ? '第 ' + day + ' 天' : '', a != null ? '實際 ' + a.toFixed(1) + '%' : ''].filter(Boolean).join('・');
    var pls = plansOn(p, td);
    $('fdPlan').innerHTML = '<div class="fd-plan-h"><span>今日預定抽查</span><b>' + pls.length + '</b></div>' + (pls.length ? pls.map(function (pl) {
      var w = WORK[pl.wid];
      return '<div class="fd-pl"><span><b>' + esc(w ? workLabel(w, pl.phase || 0) : '抽查') + '</b><small>' + esc([pl.note, pl.time].filter(Boolean).join('・')) + '</small></span><button type="button" data-fstart="' + pl.id + '">開始抽查</button></div>';
    }).join('') : '<p class="fd-none">今天沒有預定抽查。</p>');
    var defs = openDefs(p), dl = ''; defs.forEach(function (x) { var d2 = defInfo(x, td).dl; if (d2 && (!dl || d2 < dl)) dl = d2; });
    var dr = dailyOf(p, td), au = drAuto(p, td), cnt = au.s2.length + au.s3.length + au.s4.length + au.s5.length;
    $('fdTwo').innerHTML = '<button type="button" class="kpi" data-fgo="ng"><span>未改善缺失</span><b>' + defs.length + '</b><small>' + (dl ? '期限 ' + (+dl.slice(5, 7)) + '/' + (+dl.slice(8)) : defs.length ? '未設期限' : '無') + '</small></button>' +
      '<button type="button" class="kpi" data-fgo="daily"><span>今日日報</span><b class="sm">' + (dr ? '已填' : inPeriod(p, td) ? '未填' : '—') + '</b><small>' + (cnt ? '已帶入 ' + cnt + ' 筆' : '尚無自動帶入') + '</small></button>';
    var tag = { rec: '抽查', mat: '材料', doc: '公文', meet: '會議', daily: '日報' };
    var done = calEvents(p).filter(function (x) { return x.date === td && x.kind !== 'plan'; });
    $('fdDone').innerHTML = done.length ? done.map(function (x) {
      var go2 = x.rid ? ' data-openrec="' + x.wid + '|' + x.rid + '"' : x.did ? ' data-opendoc="' + x.did + '"' : x.drd ? ' data-opendr="' + x.drd + '"' : ' data-openmat="' + x.mid + '"';
      return '<button type="button" class="fd-row"' + go2 + '><span class="chip-k">' + (tag[x.kind] || '') + '</span><span class="fd-t">' + esc(x.label) + '</span>' + (x.ng ? '<span class="fd-ng">缺失</span>' : '') + '<span class="chev">›</span></button>';
    }).join('') : '<p class="muted fd-empty">今天還沒有紀錄。</p>';
    show('vField');
  }
  $('fdProj').addEventListener('change', function () { lsSet('last-pid', $('fdProj').value); renderField(); });
  $('vField').addEventListener('click', function (e) {
    var p = curProj(), b = e.target.closest('button'); if (!p || !b) return;
    var ds = b.dataset;
    if (ds.fstart) return startPlan(p, ds.fstart);
    if (ds.fgo === 'ng') { var t = today(); return go({ v: 'sumPrint', pid: p.id, to: t, from: sumFromDef(t), kind: 'ng' }); }
    if (ds.fgo === 'daily') return go({ v: 'dailyEdit', pid: p.id, date: today() });
    if (ds.fq === 'rec') return go({ v: 'works', pid: p.id });
    if (ds.fq === 'mat') return go({ v: 'materials', pid: p.id });
    if (ds.fq === 'plan') { calSel = today(); var d0 = new Date(); calY = d0.getFullYear(); calM = d0.getMonth(); go({ v: 'project', pid: p.id }); setTimeout(function () { var pa = document.querySelector('#calDay .plan-add'); if (pa) { pa.open = true; pa.scrollIntoView({ block: 'center' }); } }, 40); return; }
    if (ds.openrec) { var a2 = ds.openrec.split('|'); return go({ v: 'form', pid: p.id, wid: a2[0], rid: a2[1] }); }
    if (ds.opendoc) return go({ v: 'docEdit', pid: p.id, did: ds.opendoc });
    if (ds.opendr) return go({ v: 'dailyEdit', pid: p.id, date: ds.opendr });
    if (ds.openmat) return go({ v: 'matItem', pid: p.id, mid: ds.openmat });
  });

  $('recDir').addEventListener('click', function () { go({ v: 'works', pid: cur.pid }); });
  function lastPhrases() { var p = S.projects.filter(function (x) { return x.phrases && x.phrases.length; }).pop(); return p ? JSON.parse(JSON.stringify(p.phrases)) : null; }
  $('pjNew').addEventListener('click', function () { $('btnNewProject').click(); });
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
  function renderWorks() {
    var p = proj(cur.pid);
    topbar('施工抽查', p.info.name || '未命名工程', '工程');
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

    show('vWorks');
  }
  function renderProject() {
    var p = proj(cur.pid), I = p.info;
    topbar(I.name || '未命名工程', [I.code, I.contractor].filter(Boolean).join(' · '));
    $('pjSel').innerHTML = projOpts(p);
    /* 施工抽查進度（工項目錄在「施工抽查」頁） */
    var on = enabledWorks(p), tot = 0, ngR = 0, ph = [0, 0, 0], wk = 0, ws0 = wkStart(today()), we0 = addDays(ws0, 6), used = 0;
    on.forEach(function (w) {
      var list = (p.recs[w.id] || []).filter(function (r) { return !r.draft; }); if (list.length) used++;
      list.forEach(function (r) { tot++; if (stats(w, r).ng) ngR++; if (w.kind === 'phased') ph[r.phase]++; if (r.checkDate >= ws0 && r.checkDate <= we0) wk++; });
    });
    var openN = openDefs(p).length, plN = (p.plans || []).filter(function (pl) { return !pl.rid && pl.date >= today(); }).length;
    $('recDir').innerHTML = '<span class="md-big">' + tot + '<small> 份</small></span><span class="md-lbl">已完成施工抽查（' + used + ' 個工項）</span>' +
      '<span class="md-meta">施工前 ' + ph[0] + ' · 施工中 ' + ph[1] + ' · 施工完成 ' + ph[2] + ' · 本週 ' + wk + ' 份</span>' +
      '<span class="md-meta">' + (ngR ? '有缺失 ' + ngR + ' 份 · ' : '') + (openN ? '<b>未改善 ' + openN + ' 項</b>' : '缺失皆已改善') + (plN ? ' · 預定待抽查 ' + plN + ' 項' : '') + '</span>' +
      '<span class="md-go">進入施工抽查 ›</span>';
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
    save(); var open = $('workPickBox').open; renderWorks(); $('workPickBox').open = open;
  });
  $('btnWorkOrder').addEventListener('click', function () { var p = proj(cur.pid), sy = window.scrollY; delete p.wOrder; save(); renderWorks(); window.scrollTo(0, sy); $('workPickBox').open = true; });
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
    p.wOrder = next; save(); var sy = window.scrollY; renderWorks(); window.scrollTo(0, sy);
  });
  $('btnWorkAll').addEventListener('click', function () { var p = proj(cur.pid); delete p.works; save(); renderWorks(); $('workPickBox').open = true; });
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
        return _itemHtml(attr, 'r' + id, name + std + tools, '', refText(it.ref), it.c.r, it.c.note, it.hint, 'n' + id);
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
  /* ========== 抽查標準管理（每個工程各自一套，存在 p.tpl；已存檔的表不受影響） ========== */
  $('btnStd').addEventListener('click', function () { go({ v: 'std', pid: cur.pid }); });
  function stdWorks(p) { return orderedWorks(p).filter(function (w) { return hasStd(w) && w.kind !== 'matform'; }).concat(D.matWork && hasStd(D.matWork) ? [D.matWork] : []); }
  function stdTpl(p, wid, pi, make) {
    var t = p.tpl && p.tpl[wid] && p.tpl[wid][pi];
    if (!t && make) { p.tpl = p.tpl || {}; p.tpl[wid] = p.tpl[wid] || {}; t = p.tpl[wid][pi] = { cells: {}, extra: {}, kv: 2 }; }
    return t;
  }
  function stdClean(p, wid, pi) {
    var t = stdTpl(p, wid, pi); if (!t) return;
    Object.keys(t.cells).forEach(function (k) { var c = t.cells[k]; if (c.std == null && !c.off && !c.name && c.hint == null && c.ref == null) delete t.cells[k]; });
    Object.keys(t.extra).forEach(function (g) { if (!t.extra[g].length) delete t.extra[g]; });
    if (!Object.keys(t.cells).length && !Object.keys(t.extra).length) delete p.tpl[wid][pi];
  }
  function stdLog(p, w, pi, txt) {
    p.stdLog = p.stdLog || [];
    p.stdLog.unshift({ t: Date.now(), by: (window.Cloud && window.Cloud.me && window.Cloud.me()) || '', wid: w.id, ph: pi, txt: txt });
    if (p.stdLog.length > 300) p.stdLog.length = 300;
  }
  var stdSel = { wid: null, pi: 0 };
  function renderStd() {
    var p = proj(cur.pid), ws = stdWorks(p);
    topbar('抽查標準', p.info.name || '未命名工程', '工程');
    if (!ws.some(function (w) { return w.id === stdSel.wid; })) stdSel = { wid: ws[0] && ws[0].id, pi: 0 };
    var w = WORK[stdSel.wid]; if (!w) { $('stdBody').innerHTML = ''; show('vStd'); return; }
    if (!w.phases[stdSel.pi]) stdSel.pi = 0;
    var pi = stdSel.pi, t = stdTpl(p, w.id, pi) || { cells: {}, extra: {} };
    $('stdWork').innerHTML = ws.map(function (x) { var n = p.tpl && p.tpl[x.id] ? Object.keys(p.tpl[x.id]).length : 0; return '<option value="' + x.id + '"' + (x.id === w.id ? ' selected' : '') + '>' + esc(x.name) + (n ? '（已調整）' : '') + '</option>'; }).join('');
    $('stdPh').innerHTML = w.phases.length > 1 ? w.phases.map(function (ph, i) { var m = !!stdTpl(p, w.id, i); return '<label class="pill"><input type="radio" name="stdPh" value="' + i + '"' + (i === pi ? ' checked' : '') + '><span>' + esc(ph.name) + (m ? ' ●' : '') + '</span></label>'; }).join('') : '';
    var refOpts = Object.keys(D.refs).map(function (k) { return '<option value="' + esc(D.refs[k]) + '">'; }).join('');
    var fld = function (f, v, ph, area) { return area ? '<textarea class="inline-in" data-f="' + f + '" rows="2" placeholder="' + ph + '">' + esc(v) + '</textarea>' : '<input class="inline-in" data-f="' + f + '" placeholder="' + ph + '" value="' + esc(v) + '"' + (f === 'ref' ? ' list="stdRefs"' : '') + '>'; };
    var card = function (attr, it, mod, off, custom) {
      return '<div class="std-item' + (off ? ' off' : '') + (mod ? ' mod' : '') + '" ' + attr + '>' +
        '<div class="std-top"><span class="std-code">' + (custom ? '本工程新增' : esc(it.id)) + '</span>' + (mod ? '<span class="std-badge">已調整</span>' : '') +
        (custom ? '<button type="button" class="link" data-sdel>刪除</button>' : (mod ? '<button type="button" class="link" data-sreset>還原公司預設</button>' : '') + '<label class="std-off"><input type="checkbox" data-soff' + (off ? ' checked' : '') + '>本工程不使用</label>') + '</div>' +
        '<label class="std-l">項目名稱' + fld('name', it.name, '項目名稱') + '</label>' +
        '<label class="std-l">抽查標準' + fld('std', it.std, '抽查標準（定性或定量）', true) + '</label>' +
        '<div class="std-row"><label class="std-l">實測提示' + fld('hint', it.hint || '', '例：實測：＿＿cm') + '</label><label class="std-l">依據' + fld('ref', refText(it.ref), '例：公司施工抽查紀錄表') + '</label></div></div>';
    };
    var html = w.phases[pi].groups.map(function (g) {
      var items = g.items.map(function (it) {
        var c = t.cells[it.id] || {}, mod = c.std != null || !!c.name || c.hint != null || c.ref != null;
        var v = { id: it.id, name: c.name || it.name, std: c.std != null ? c.std : it.std, hint: c.hint != null ? c.hint : it.hint, ref: c.ref != null ? c.ref : it.ref };
        return card('data-skey="' + it.id + '"', v, mod, !!c.off, false);
      }).join('') + (t.extra[g.id] || []).map(function (x, xi) { return card('data-sx="' + g.id + '|' + xi + '"', x, false, false, true); }).join('');
      return '<section class="card std-group"><h3 class="h small">' + esc(g.name) + '</h3>' + items + '<button type="button" class="btn ghost sm" data-sadd="' + g.id + '">＋ 新增項目</button></section>';
    }).join('');
    var log = (p.stdLog || []).filter(function (l) { return l.wid === w.id; }).slice(0, 30);
    $('stdBody').innerHTML = '<datalist id="stdRefs">' + refOpts + '</datalist>' + html +
      '<div class="std-foot"><button type="button" class="btn ghost" id="btnStdReset"' + (stdTpl(p, w.id, pi) ? '' : ' disabled') + '>此階段全部恢復公司預設</button></div>' +
      '<details class="card std-log"' + (log.length ? '' : ' hidden') + '><summary>修改紀錄（' + log.length + '）</summary><ul>' + log.map(function (l) {
        return '<li><span class="muted">' + fmtOut(new Date(l.t - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 16)) + (l.by ? ' · ' + esc(l.by.replace(/@.*$/, '')) : '') + ' · ' + esc((w.phases[l.ph] || {}).name || '') + '</span><br>' + esc(l.txt) + '</li>';
      }).join('') + '</ul></details>';
    show('vStd');
  }
  $('stdWork').addEventListener('change', function () { stdSel = { wid: $('stdWork').value, pi: 0 }; renderStd(); });
  $('stdPh').addEventListener('change', function (e) { if (e.target.name === 'stdPh') { stdSel.pi = +e.target.value; renderStd(); } });
  function stdBase(w, pi, key) { var it = null; w.phases[pi].groups.forEach(function (g) { g.items.forEach(function (x) { if (x.id === key) it = x; }); }); return it; }
  var STD_F = { name: '名稱', std: '標準', hint: '提示', ref: '依據' };
  $('stdBody').addEventListener('change', function (e) {
    var el = e.target, box = el.closest('[data-skey],[data-sx]'); if (!box) return;
    var p = proj(cur.pid), w = WORK[stdSel.wid], pi = stdSel.pi, t = stdTpl(p, w.id, pi, true), f = el.dataset.f;
    if (box.dataset.skey) {
      var key = box.dataset.skey, base = stdBase(w, pi, key), c = t.cells[key] || (t.cells[key] = { r: '', note: '' });
      if (el.hasAttribute('data-soff')) { if (el.checked) c.off = true; else delete c.off; stdLog(p, w, pi, (el.checked ? '停用' : '恢復使用') + '「' + (c.name || base.name) + '」'); }
      else if (f) {
        var v = el.value.trim(), bv = f === 'ref' ? refText(base.ref) : (base[f] || ''), old = c[f] != null ? (f === 'ref' ? refText(c.ref) : c[f]) : bv;
        if (v === bv || (f === 'name' && !v)) delete c[f]; else c[f] = f === 'ref' ? (Object.keys(D.refs).filter(function (k) { return D.refs[k] === v; })[0] || v) : v;
        if (v !== old) stdLog(p, w, pi, '「' + (c.name || base.name) + '」' + STD_F[f] + '：' + (old || '（空白）') + ' → ' + (v || '（空白）'));
      }
    } else {
      var a = box.dataset.sx.split('|'), x = (t.extra[a[0]] || [])[+a[1]]; if (!x || !f) return;
      var nv = el.value.trim(); if (f === 'ref') nv = Object.keys(D.refs).filter(function (k) { return D.refs[k] === nv; })[0] || nv;
      if (x[f] !== nv) { stdLog(p, w, pi, '本工程新增項目「' + (f === 'name' ? nv : x.name) + '」' + STD_F[f] + '：' + (x[f] || '（空白）') + ' → ' + (el.value.trim() || '（空白）')); x[f] = nv; }
    }
    stdClean(p, w.id, pi); save();
    if (el.hasAttribute('data-soff')) renderStd();
    else { var b2 = box.querySelector('.std-badge'); if (box.dataset.skey && !b2 && t.cells[box.dataset.skey]) { box.classList.add('mod'); } }
  });
  $('stdBody').addEventListener('click', function (e) {
    var p = proj(cur.pid), w = WORK[stdSel.wid], pi = stdSel.pi, b = e.target;
    if (b.dataset.sadd) { var t = stdTpl(p, w.id, pi, true); (t.extra[b.dataset.sadd] = t.extra[b.dataset.sadd] || []).push({ name: '', std: '', hint: '', ref: '', r: '', note: '' }); save(); renderStd(); var n = document.querySelectorAll('[data-sx^="' + b.dataset.sadd + '|"]'); if (n.length) n[n.length - 1].querySelector('input').focus(); return; }
    var box = b.closest('[data-skey],[data-sx]');
    if (b.hasAttribute('data-sreset') && box) { var t1 = stdTpl(p, w.id, pi), c = t1 && t1.cells[box.dataset.skey]; if (c) { ['std', 'name', 'hint', 'ref'].forEach(function (f) { delete c[f]; }); stdLog(p, w, pi, '「' + stdBase(w, pi, box.dataset.skey).name + '」還原公司預設'); stdClean(p, w.id, pi); save(); renderStd(); } return; }
    if (b.hasAttribute('data-sdel') && box) { var a = box.dataset.sx.split('|'), t2 = stdTpl(p, w.id, pi), x = t2.extra[a[0]][+a[1]]; t2.extra[a[0]].splice(+a[1], 1); if (x.name || x.std) stdLog(p, w, pi, '刪除本工程新增項目「' + (x.name || x.std) + '」'); stdClean(p, w.id, pi); save(); renderStd(); return; }
    if (b.id === 'btnStdReset') ask('恢復公司預設？', '「' + w.name + ' ' + w.phases[pi].name + '」在本工程的調整會全部清除。已存檔的抽查表不受影響。', '取消', '恢復', function () { delete p.tpl[w.id][pi]; stdLog(p, w, pi, '此階段全部恢復公司預設'); save(); renderStd(); });
  });

  /* 本工程預設標準 */
  $('btnTplSave').addEventListener('click', function () {
    var p = proj(cur.pid), w = W(), r = R(), cells = {}, extra = {};
    Object.keys(r.res).forEach(function (key) { var c = r.res[key]; if (c.std != null || c.off || c.name || c.hint != null || c.ref != null) { cells[key] = { r: '', note: '' }; ['std', 'name', 'hint', 'ref'].forEach(function (f) { if (c[f] != null && c[f] !== '' || (f === 'std' && c.std != null)) cells[key][f] = c[f]; }); if (c.off) cells[key].off = true; } });
    Object.keys(r.extra || {}).forEach(function (gi) { extra[gi] = r.extra[gi].filter(function (x) { return x.name || x.std; }).map(function (x) { return { name: x.name, std: x.std, hint: x.hint || '', ref: x.ref || '', r: '', note: '' }; }); });
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
      return _itemHtml('data-row="' + i + '"', 'row' + i, '<div class="item-name">' + name + '</div>', std ? '<div class="item-std">' + std + '</div>' : '', refText(x.ref), x.r, x.note, x.hint, 'rn' + i, del);
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
    fitForms(paper);
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
    var libs = (window.html2canvas ? Promise.resolve() : loadScript(CDN.h2c)).then(function () { return window.jspdf ? null : loadScript(CDN.jspdf); });
    if (cur.v === 'docPrint') {
      libs.then(docPdfAll).then(function (names) { $('pdfMsg').textContent = '已產生 ' + names.length + ' 個檔案：' + names.join('、') + '。' + $('pdfMsg').textContent; })
        .catch(function (err) { $('pdfMsg').textContent = (err && err.message) || 'PDF 產生失敗'; }).then(function () { btn.disabled = false; btn.textContent = label; });
      return;
    }
    libs.then(makePdf).then(function (blob) {
      var a = document.createElement('a'), name = (docName || '抽查表') + '.pdf';
      a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 60000);
      $('pdfMsg').textContent = '已產生「' + name + '」。';
    }).catch(function (err) { $('pdfMsg').textContent = (err && err.message) || 'PDF 產生失敗，請改用「列印」再選「另存為 PDF」。'; })
      .then(function () { btn.disabled = false; btn.textContent = label; });
  });
  function makePdf() {
    /* 公文：確定標楷體與 Times 已載入才轉成圖片 */
    if (($('paper').classList.contains('gw-paper') || $('paper').classList.contains('dr-paper')) && document.fonts && document.fonts.load) {
      return Promise.all([document.fonts.load('16pt "GW Kai"', '測'), document.fonts.load('16pt "GW Times"', 'A1')]).catch(function () {}).then(makePdf2);
    }
    return makePdf2();
  }
  function makePdf2() {
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
    var gw = src.classList.contains('gw-paper') || src.classList.contains('dr-paper');
    pages.forEach(function (pg) { pg.style.width = Wmm + 'mm'; pg.style.padding = gw ? '0' : my + 'mm ' + mx + 'mm'; host.appendChild(pg); });
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
  /* 不定期／勞安：依實際高度補空白列，讓表單不論項目多寡都剛好排滿一張 A4（以列印寬度 190mm 量測） */
  function fitForms(paper) {
    var forms = paper.querySelectorAll('.fit-form'); if (!forms.length) return;
    var mm = 96 / 25.4, target = 276 * mm;
    var host = document.createElement('div'); host.className = 'paper';
    host.style.cssText = 'position:absolute;left:-10000px;top:0;width:190mm;max-width:none;padding:0;border:0;margin:0;visibility:hidden';
    document.body.appendChild(host);
    forms.forEach(function (f) {
      var c = f.cloneNode(true); host.innerHTML = ''; host.appendChild(c);
      var t = c.querySelector('table.inner'), real = f.querySelector('table.inner'); if (!t || !real) return;
      var cols = t.rows[0] ? t.rows[0].cells.length : 4, tb = t.tBodies[t.tBodies.length - 1] || t;
      var mk = function () { var tr = document.createElement('tr'); tr.className = 'blank'; for (var i = 0; i < cols; i++) { var td = document.createElement('td'); td.innerHTML = '&nbsp;'; tr.appendChild(td); } return tr; };
      var n = 0;
      while (n < 60 && c.offsetHeight < target) { tb.appendChild(mk()); n++; }
      if (n && c.offsetHeight > target) n--;
      var rtb = real.tBodies[real.tBodies.length - 1] || real;
      for (var i = 0; i < n; i++) rtb.appendChild(mk());
    });
    host.remove();
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
    } else {
      inner += tableBody(effGroups(w, r), PAD.safety);
    }
    inner += '</table>';
    var ng = r.result === 'ng';
    return '<div class="fit-form"><h3>' + title + '</h3><p class="no">編號：' + esc(r.docNo) + '</p>' +
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
      '<p class="signline">監造現場人員：' + esc(I.inspector) + ' ' + sigImg(r, 'sigInspector') + '</p>' + foot() + '</div>';
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

  /* ========== 公文（發文）：每個工程各自的字號、受文者、聯絡人；公司名稱與公司章全公司共用 ========== */
  var COMPANY_DEF = { name: '周冠宇建築師事務所', addr: '300新竹市東區西大路135巷2號', head: '周冠宇', seal: '' };
  var DOC_LINK = 'https://www.ey.gov.tw/Page/9695ADCD1F0CB9F4';
  var DOC_PH = {
    subj: ['檢送「{工程}」○○，請 查照。', '有關「{工程}」○○乙案，經審尚符契約規定，敬請 貴府准予辦理，請 查照。', '有關「{工程}」○○乙案，詳如說明，請 查照。', '，請 鑒核。', '，請 核備。', '，請 查照。'],
    desc: ['依據○○公司○○年○月○日○○字第○○號函辦理。', '復 貴府○○年○月○日○○字第○○號函。', '隨函檢附○○一式○份如附件。', '請 貴公司於文到○日內改善完成並回覆本所。', '旨揭資料經本所審查符合契約規定。', '詳如附件。']
  };
  function company() { if (!S.company) S.company = JSON.parse(JSON.stringify(COMPANY_DEF)); return S.company; }
  function docCfg(p) { if (!p.docCfg) p.docCfg = { caseNo: '', prefix: '周建字第', last: 0, contacts: [], to: [], cc: [] }; return p.docCfg; }
  function docById(p, id) { return (p.docs || []).filter(function (d) { return d.id === id; })[0]; }
  function rocFull(d) { var a = (d || '').split('-'); return a.length === 3 ? '中華民國' + (a[0] - 1911) + '年' + a[1] + '月' + a[2] + '日' : ''; }
  function docNo(cfg, date, n) { return (cfg.prefix || '') + ((date || today()).slice(0, 4) - 1911) + (cfg.caseNo || '') + ('0000' + n).slice(-5) + '號'; }
  function lines(t) { return String(t || '').split(/\n/).map(function (x) { return x.trim(); }).filter(Boolean); }
  $('btnDocs').addEventListener('click', function () { go({ v: 'docs', pid: cur.pid }); });
  function renderDocs() {
    var p = proj(cur.pid), cfg = docCfg(p), co = company();
    topbar('公文收發', p.info.name || '未命名工程', '工程');
    $('dcCase').value = cfg.caseNo; $('dcPrefix').value = cfg.prefix; $('dcLast').value = cfg.last || 0;
    $('dcNext').textContent = cfg.caseNo ? '下一號：' + docNo(cfg, today(), (cfg.last || 0) + 1) : '請先填寫案件編號';
    $('dcContacts').value = (cfg.contacts || []).map(function (c) { return [c.name, c.tel, c.mail].join('｜'); }).join('\n');
    $('dcTo').value = (cfg.to || []).join('\n'); $('dcCc').value = (cfg.cc || []).join('\n');
    $('coName').value = co.name || ''; $('coAddr').value = co.addr || ''; $('coHead').value = co.head || '';
    $('coSealImg').hidden = !co.seal; if (co.seal) { $('coSealImg').setAttribute('data-img', co.seal); fillImgs($('coSealBox')); }
    if (!cfg.caseNo) $('dcSetBox').open = true;
    var list = (p.docs || []).slice().sort(function (a, b) { return (b.date || '') < (a.date || '') ? -1 : 1; });
    $('docList').innerHTML = list.length ? list.map(function (d) {
      var t = d.kind === 'meet' ? '開會通知單：' + (d.meet && d.meet.reason || '') : d.subj || '';
      return '<button type="button" class="doc-row" data-doc="' + d.id + '"><span class="doc-d">' + roc(d.date, '.') + '</span><span class="doc-t"><b>' + esc(t || '（未填主旨）') + '</b><small>' + esc(d.no) + (d.to ? ' · ' + esc(d.to) : '') + '</small></span>' +
        (d.fileId ? '<span class="doc-ok" title="已歸檔到雲端">已歸檔</span>' : '') + '<span class="chev">›</span></button>';
    }).join('') : '<div class="empty">尚無發文。按上方「＋ 函」或「＋ 開會通知單」開始。</div>';
    show('vDocs');
  }
  function cfgField(id, fn) { $(id).addEventListener('change', function () { var p = proj(cur.pid); fn(docCfg(p), $(id).value, p); save(); renderDocs(); }); }
  cfgField('dcCase', function (c, v) { c.caseNo = v.trim(); });
  cfgField('dcPrefix', function (c, v) { c.prefix = v.trim(); });
  cfgField('dcLast', function (c, v) { c.last = Math.max(0, parseInt(v, 10) || 0); });
  cfgField('dcContacts', function (c, v) { c.contacts = lines(v).map(function (l) { var a = l.split(/[｜|]/); return { name: (a[0] || '').trim(), tel: (a[1] || '').trim(), mail: (a[2] || '').trim() }; }); });
  cfgField('dcTo', function (c, v) { c.to = lines(v); });
  cfgField('dcCc', function (c, v) { c.cc = lines(v); });
  ['coName', 'coAddr', 'coHead'].forEach(function (id) { $(id).addEventListener('change', function () { company()[{ coName: 'name', coAddr: 'addr', coHead: 'head' }[id]] = $(id).value.trim(); save(); }); });
  $('coSealPick').addEventListener('change', function () {
    var f = $('coSealPick').files[0]; if (!f) return;
    var rd = new FileReader();
    rd.onload = function () { var id = 'seal-' + uid(), data = rd.result; IDB.put(id, data).then(function () { thumbCache[id] = data; company().seal = id; save(); if (window.Cloud) window.Cloud.uploadPhoto(id, data); renderDocs(); }); };
    rd.readAsDataURL(f); $('coSealPick').value = '';
  });
  function newDoc(kind) {
    var p = proj(cur.pid), cfg = docCfg(p);
    if (!cfg.caseNo) { ask('請先設定案件編號', '發文字號會由「字號前綴＋民國年＋案件編號＋5 碼流水號」組成，請在「發文設定」填寫本工程的案件編號（例如 197）。', '好', '', function () {}); $('dcSetBox').open = true; return; }
    var n = (cfg.last || 0) + 1, c = (cfg.contacts || [])[0] || {}, d = today();
    var doc = { id: uid(), dir: 'out', kind: kind, seq: n, no: docNo(cfg, d, n), date: d, to: (cfg.to || [])[0] || '', speed: '普通件', secret: '普通', att: '', subj: '', desc: '', main: '', cc: '', cName: c.name || '', cTel: c.tel || '', cMail: c.mail || '', created: Date.now() };
    if (kind === 'meet') doc.meet = { reason: '', time: '', place: '', host: '', ctc: '', att: '', obs: '', note: '' };
    cfg.last = n; (p.docs = p.docs || []).push(doc); save();
    go({ v: 'docEdit', pid: p.id, did: doc.id });
  }
  $('btnNewLetter').addEventListener('click', function () { newDoc('letter'); });
  $('btnNewMeet').addEventListener('click', function () { newDoc('meet'); });
  $('docList').addEventListener('click', function (e) { var b = e.target.closest('[data-doc]'); if (b) go({ v: 'docEdit', pid: cur.pid, did: b.dataset.doc }); });
  var DOC_F = [['no', '發文字號'], ['date', '發文日期', 'date'], ['speed', '速別', ['普通件', '速件', '最速件']], ['secret', '密等及解密條件或保密期限', ['普通', '密', '機密']], ['att', '附件'], ['short', '簡易主旨（檔名用，空白則自動擷取）']];
  var MEET_F = [['reason', '開會事由', 'area'], ['time', '開會時間', 'datetime-local'], ['place', '開會地點'], ['host', '主持人'], ['ctc', '聯絡人及電話'], ['att', '出席者'], ['obs', '列席者'], ['note', '備註（一行一點）', 'area']];
  function renderDocEdit() {
    var p = proj(cur.pid), d = docById(p, cur.did), cfg = docCfg(p);
    topbar(d.kind === 'meet' ? '開會通知單' : '函', d.no, '公文');
    var inp = function (k, l, t, v, list) {
      var val = v == null ? '' : v;
      if (Array.isArray(t)) return '<label class="f">' + l + '<select data-k="' + k + '">' + t.map(function (o) { return '<option' + (o === val ? ' selected' : '') + '>' + o + '</option>'; }).join('') + '</select></label>';
      if (t === 'area') return '<label class="f wide">' + l + '<textarea data-k="' + k + '" rows="3">' + esc(val) + '</textarea></label>';
      return '<label class="f">' + l + '<input data-k="' + k + '"' + (t ? ' type="' + t + '"' : '') + (list ? ' list="' + list + '"' : '') + ' value="' + esc(val) + '"></label>';
    };
    var chips = function (key) { return '<div class="chips doc-chips">' + DOC_PH[key].map(function (t, i) { return '<button type="button" class="chip" data-ins="' + key + '|' + i + '">' + esc(t.replace('{工程}', '工程名稱')) + '</button>'; }).join('') + (key === 'subj' ? '<button type="button" class="chip" data-ins="proj">插入「工程名稱」</button>' : '') + '</div>'; };
    var dl = function (id, arr) { return '<datalist id="' + id + '">' + (arr || []).map(function (x) { return '<option value="' + esc(x) + '">'; }).join('') + '</datalist>'; };
    var contacts = (cfg.contacts || []).map(function (c, i) { return '<option value="' + i + '">' + esc(c.name) + '</option>'; }).join('');
    var body = d.kind === 'meet'
      ? '<section class="card"><h2 class="h small">會議內容</h2><div class="grid g2">' + MEET_F.map(function (f) { return inp('meet.' + f[0], f[1], f[2], d.meet[f[0]], f[0] === 'att' || f[0] === 'obs' ? 'dlCc' : ''); }).join('') + '</div>' + chips('desc') + '</section>'
      : '<section class="card"><h2 class="h small">主旨</h2><textarea class="doc-area" data-k="subj" rows="4" placeholder="主旨（一段，結尾用「請 查照」等期望語）">' + esc(d.subj) + '</textarea>' + chips('subj') +
        '<h2 class="h small" style="margin-top:14px">說明</h2><textarea class="doc-area" data-k="desc" rows="7" placeholder="一行一點，系統自動編「一、二、三」；行首空兩格為下一層「(一)(二)」">' + esc(d.desc) + '</textarea>' + chips('desc') +
        '<p class="muted doc-ref">寫作參考：<a href="' + DOC_LINK + '" target="_blank" rel="noopener">行政院 文書處理手冊（事務管理各手冊）</a></p></section>';
    $('docEditBody').innerHTML = dl('dlTo', cfg.to) + dl('dlCc', cfg.cc) +
      '<section class="card"><h2 class="h small">發文資料</h2><div class="grid g2">' +
      inp('to', '受文者', '', d.to, 'dlTo') + DOC_F.map(function (f) { return inp(f[0], f[1], f[2], d[f[0]]); }).join('') + '</div>' +
      '<div class="grid g3" style="margin-top:8px">' + (contacts ? '<label class="f">帶入聯絡人<select id="docCtc"><option value="">—</option>' + contacts + '</select></label>' : '') +
      inp('cName', '案件聯絡人', '', d.cName) + inp('cTel', '聯絡方式', '', d.cTel) + inp('cMail', '電子信箱', '', d.cMail) + '</div></section>' + body +
      '<section class="card"><div class="grid g2">' + inp('main', '正本', '', d.main, 'dlCc') + inp('cc', '副本', '', d.cc, 'dlCc') + '</div>' +
      '<p class="muted">常用受文者與正副本清單在「公文收發」頁的「發文設定」。</p></section>' +
      '<div class="doc-acts"><button type="button" class="btn primary" id="btnDocPv">預覽／輸出 PDF</button><button type="button" class="btn ghost" id="btnDocDel">刪除此公文</button></div>';
    var sh = $('docEditBody').querySelector('[data-k="short"]'); if (sh) sh.placeholder = '自動：' + shortSubj(Object.assign({}, d, { short: '' }));
    show('vDocEdit');
  }
  function docSet(d, k, v) { if (k.indexOf('meet.') === 0) d.meet[k.slice(5)] = v; else d[k] = v; }
  $('docEditBody').addEventListener('input', function (e) {
    var k = e.target.dataset.k; if (!k) return;
    var p = proj(cur.pid), d = docById(p, cur.did), cfg = docCfg(p);
    /* 改發文日期時，字號裡的民國年跟著更新（流水號不變） */
    if (k === 'date' && d.seq && e.target.value && d.no === docNo(cfg, d.date, d.seq)) { d.no = docNo(cfg, e.target.value, d.seq); var ni = $('docEditBody').querySelector('[data-k="no"]'); if (ni) ni.value = d.no; }
    docSet(d, k, e.target.value); save();
    if (k === 'subj' || k === 'meet.reason') { var sh = $('docEditBody').querySelector('[data-k="short"]'); if (sh) sh.placeholder = '自動：' + shortSubj(Object.assign({}, d, { short: '' })); }
  });
  $('docEditBody').addEventListener('change', function (e) {
    if (e.target.id !== 'docCtc' || e.target.value === '') return;
    var p = proj(cur.pid), d = docById(p, cur.did), c = docCfg(p).contacts[+e.target.value];
    d.cName = c.name; d.cTel = c.tel; d.cMail = c.mail; save(); renderDocEdit();
  });
  $('docEditBody').addEventListener('click', function (e) {
    var b = e.target, p = proj(cur.pid), d = docById(p, cur.did);
    if (b.dataset.ins) {
      var sec = b.closest('.card'), ta = sec.querySelector('textarea[data-k="desc"], textarea[data-k="meet.note"]');
      var txt = b.dataset.ins === 'proj' ? '「' + (p.info.name || '') + '」' : DOC_PH[b.dataset.ins.split('|')[0]][+b.dataset.ins.split('|')[1]].replace('{工程}', p.info.name || '');
      if (b.dataset.ins === 'proj' || b.dataset.ins.indexOf('subj') === 0) ta = sec.querySelector('textarea[data-k="subj"]');
      if (!ta) return;
      var s0 = ta.selectionStart != null && document.activeElement === ta ? ta.selectionStart : ta.value.length, v = ta.value;
      var add = ta.dataset.k === 'subj' || b.dataset.ins === 'proj' ? txt : (v && !/\n$/.test(v.slice(0, s0)) && s0 ? '\n' : '') + txt;
      ta.value = v.slice(0, s0) + add + v.slice(s0); docSet(d, ta.dataset.k, ta.value); save(); ta.focus(); ta.selectionStart = ta.selectionEnd = s0 + add.length; return;
    }
    if (b.id === 'btnDocPv') { go({ v: 'docPrint', pid: p.id, did: d.id }); return; }
    if (b.id === 'btnDocDel') ask('刪除這份公文？', '「' + d.no + '」會從本工程移除（雲端已歸檔的 PDF 不會刪除）。', '取消', '刪除', function () { p.docs = p.docs.filter(function (x) { return x !== d; }); save(); history.back(); });
  });
  /* 說明：一行一點自動編號，行首空兩格為第二層 */
  var CN = '一二三四五六七八九十'.split('');
  function cnNum(n) { return n <= 10 ? CN[n - 1] : n < 20 ? '十' + CN[n - 11] : CN[Math.floor(n / 10) - 1] + '十' + (n % 10 ? CN[n % 10 - 1] : ''); }
  function pointsHtml(t) {
    var raw = String(t || '').split(/\n/).filter(function (x) { return x.trim(); }), out = '', n = 0, m = 0;
    var strip = function (s) { return s.trim().replace(/^([一二三四五六七八九十]+[、.．]|[0-9０-９]+[.、．)）]|[(（][一二三四五六七八九十0-9]+[)）])\s*/, ''); };
    var single = raw.filter(function (x) { return !/^[ 　\t]/.test(x); }).length === 1 && raw.length === 1;
    raw.forEach(function (l) {
      if (/^([ 　]{2,}|\t)/.test(l) && n) { m++; out += '<div class="gw-p2"><span>(' + cnNum(m) + ')</span><span>' + esc(strip(l)) + '</span></div>'; }
      else { n++; m = 0; out += single ? '<div class="gw-p1 one"><span></span><span>' + esc(strip(l)) + '</span></div>' : '<div class="gw-p1"><span>' + cnNum(n) + '.</span><span>' + esc(strip(l)) + '</span></div>'; }
    });
    return out;
  }
  function meetTime(v) {
    var m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(v || ''); if (!m) return '';
    var wd = '日一二三四五六'[new Date(+m[1], m[2] - 1, +m[3]).getDay()];
    return '中華民國' + (m[1] - 1911) + '年' + (+m[2]) + '月' + (+m[3]) + '日（星期' + wd + '）' + m[4] + '時' + m[5] + '分';
  }
  /* ===== 公文輸出：依正本、副本各出一份（受文者不同），內部單位不另出；可跨頁、印章可拖曳 ===== */
  var MM = 96 / 25.4, GW_TOP1 = 63.5, GW_TOPN = 25.5, GW_LIMIT = 271, GW_END = 42.1;
  function splitUnits(t) { return String(t || '').split(/[、，,；;\n]+/).map(function (x) { return x.replace(/[（(][^）)]*[）)]/g, '').trim(); }).filter(Boolean); }
  function internalUnit(n) { return /^(本所|本事務所|本公司|本處)/.test(n); }
  function docCopies(d) {
    var main = d.kind === 'meet' ? splitUnits(d.meet.att) : splitUnits(d.main), cc = splitUnits(d.cc), out = [], seen = {};
    if (!main.length && d.to) main = [d.to];
    main.forEach(function (u) { if (!internalUnit(u) && !seen[u]) { seen[u] = 1; out.push({ unit: u, kind: '正本' }); } });
    cc.forEach(function (u) { if (!internalUnit(u) && !seen[u]) { seen[u] = 1; out.push({ unit: u, kind: '副本' }); } });
    if (!out.length) out.push({ unit: d.to || '', kind: '正本' });
    return out;
  }
  /* 檔名用簡易主旨：去掉「有關」、工程名稱與期望語，只留第一句 */
  function shortSubj(d) {
    if (d.short) return d.short;
    var t = d.kind === 'meet' ? (d.meet.reason || '') : (d.subj || '');
    t = t.replace(/[「『][^」』]*[」』]/g, '').replace(/^(有關|關於|檢陳|為)\s*/, '').replace(/^[案之的，、\s]+/, '');
    t = t.split(/[，,。；]/)[0].replace(/乙案$|一案$|案$/, '').trim();
    return t.slice(0, 30) || (d.kind === 'meet' ? '開會通知' : '函');
  }
  function docFileName(d) { var a = (d.date || '').split('-'); return (a.length === 3 ? (a[0] - 1911) + '.' + a[1] + a[2] : '') + '_' + d.no + '_' + shortSubj(d); }
  function gwBodyItems(d) {
    var meet = d.kind === 'meet', row = function (k, v) { return '<div class="gw-row gw-b"><span class="gw-k">' + k + '：</span><span class="gw-v">' + v + '</span></div>'; };
    var pts = function (t) { return pointsHtml(t).replace(/<div class="gw-p/g, '<div class="gw-b gw-p'); };
    var h = meet
      ? row('開會事由', esc(d.meet.reason)) + row('開會時間', meetTime(d.meet.time)) + row('開會地點', esc(d.meet.place)) + (d.meet.host ? row('主持人', esc(d.meet.host)) : '') + row('聯絡人及電話', esc(d.meet.ctc)) +
        (d.meet.note ? row('備註', '') + pts(d.meet.note) : '')
      : row('主旨', esc(d.subj)) + (lines(d.desc).length ? row('說明', '') + pts(d.desc) : '');
    return h.replace('<div class="gw-row gw-b">', '<div class="gw-row gw-b gw-first">');
  }
  /* 量測每段高度後分頁：第一頁從受文者開始，續頁從上緣 25.5mm 開始；最後一頁底部保留正副本與負責人 */
  function gwPaginate(d) {
    var host = document.createElement('div'); host.className = 'paper gw-paper';
    host.style.cssText = 'position:absolute;left:-10000px;top:0;visibility:hidden';
    host.innerHTML = '<div class="gw gw-measure"><div class="gw-to">受文者：' + esc(d.to) + '</div><div class="gw-meta">' + gwMeta(d) + '</div>' + gwBodyItems(d) + '</div>';
    document.body.appendChild(host);
    var gw = host.firstChild, kids = Array.prototype.slice.call(gw.children), top0 = gw.getBoundingClientRect().top, prevEnd = null, items = [];
    kids.forEach(function (el, i) {
      var r = el.getBoundingClientRect(), mt = parseFloat(getComputedStyle(el).marginTop) || 0;
      var s = (prevEnd == null ? r.top - mt : prevEnd) - top0, e = r.bottom - top0; prevEnd = r.bottom;
      items.push({ html: el.outerHTML, h: (e - s) / MM, head: i < 2 });
    });
    host.remove();
    var pages = [[]], y = GW_TOP1;
    items.forEach(function (it) {
      if (!it.head && y + it.h > GW_LIMIT && pages[pages.length - 1].length > (pages.length === 1 ? 2 : 0)) { pages.push([]); y = GW_TOPN; }
      pages[pages.length - 1].push(it.html); y += it.h;
    });
    if (y + GW_END > GW_LIMIT) pages.push([]);
    return pages;
  }
  function gwMeta(d) { return '發文日期：' + rocFull(d.date) + '<br>發文字號：' + esc(d.no) + '<br>速別：' + esc(d.speed) + '<br>密等及解密條件或保密期限：' + esc(d.secret) + '<br>附件：' + esc(d.att); }
  function gwSealPos(d, n) { var s = d.sealPos || {}; return { pg: Math.min(Math.max(1, s.pg || n), n), x: s.x != null ? s.x : 124.7, y: s.y != null ? s.y : 237 }; }
  function gwCopyHtml(d, co, copy, pages) {
    var meet = d.kind === 'meet', n = pages.length, sp = gwSealPos(d, n);
    var dash = ''; for (var y = 0; y < 245; y += 2.4) dash += '<i style="top:' + y.toFixed(1) + 'mm"></i>';
    var bind = '<div class="gw-bind" aria-hidden="true">' + dash + '<span style="top:72.5mm">裝</span><span style="top:122.3mm">訂</span><span style="top:171mm">線</span></div>';
    var head = '<div class="gw-file">檔　　號：<br>保存年限：</div><h1 class="gw-title">' + esc(co.name) + '　' + (meet ? '開會通知單' : '函') + '</h1>' +
      '<div class="gw-addr">地址：' + esc(co.addr) + '<br>案件聯絡人：' + esc(d.cName) + '<br>聯絡方式：' + esc(d.cTel) + (d.cMail ? '<br>電子信箱：' + esc(d.cMail) : '') + '</div>';
    var cc = meet ? '<div class="gw-cc">出席者：' + esc(d.meet.att) + (d.meet.obs ? '<br>列席者：' + esc(d.meet.obs) : '') + '<br>副本：' + esc(d.cc) + '</div>'
      : '<div class="gw-cc">正本：' + esc(d.main) + '<br>副本：' + esc(d.cc) + '</div>';
    var end = '<div class="gw-end">' + cc + (meet ? '' : '<div class="gw-head">負責人：' + esc(co.head) + '</div>') + '</div>';
    var src = co.seal && thumbCache[co.seal];
    var seal = d.sealOn ? (src ? '<img class="gw-seal-abs" alt="公司章" draggable="false" src="' + src + '" style="left:' + sp.x + 'mm;top:' + sp.y + 'mm">' : '<span class="gw-noseal">（尚未上傳公司章）</span>') : '';
    return pages.map(function (items, i) {
      var first = i === 0, last = i === n - 1;
      var body = items.join('').replace('受文者：' + esc(d.to) + '</div>', '受文者：' + esc(copy.unit) + '</div>');
      return '<section class="gw-sheet' + (first ? '' : ' rec-sec') + '"><div class="gw' + (first ? '' : ' gw-cont') + '" data-pg="' + (i + 1) + '">' + bind + (first ? head : '') + body + (last ? end : '') +
        (i + 1 === sp.pg ? seal : '') + '<div class="gw-pg">第' + (i + 1) + '頁　共' + n + '頁</div></div></section>';
    }).join('');
  }
  var gwPick = null;   // 目前勾選要輸出的份數（以單位名稱記）
  function renderDocPrint() {
    var p = proj(cur.pid), d = docById(p, cur.did), co = company();
    topbar('列印預覽', d.no, '返回修改');
    var copies = docCopies(d), off = d.copiesOff || [];
    var sel = copies.filter(function (c) { return off.indexOf(c.unit) < 0; });
    var paint = function () {
      var pages = gwPaginate(d), paper = $('paper');
      paper.className = 'paper gw-paper';
      paper.innerHTML = (sel.length ? sel : copies.slice(0, 1)).map(function (c) { return gwCopyHtml(d, co, c, pages); }).join('').replace(/^<section class="gw-sheet rec-sec">/, '<section class="gw-sheet">');
      Array.prototype.forEach.call(paper.querySelectorAll('.gw-sheet'), function (s, i) { if (i) s.classList.add('rec-sec'); });
    };
    $('gwBar').hidden = false;
    $('gwBar').innerHTML = '<span class="gw-bar-k">輸出</span>' + copies.map(function (c, i) {
      return '<label><input type="checkbox" data-copy="' + i + '"' + (off.indexOf(c.unit) < 0 ? ' checked' : '') + '>' + c.kind + '：' + esc(c.unit) + '</label>';
    }).join('') + '<span class="gw-bar-k">用印</span><label><input type="checkbox" id="gwSealOn"' + (d.sealOn ? ' checked' : '') + '>電子章</label>' +
      (d.sealOn ? '<button type="button" class="link" id="gwSealReset">印章回到預設位置</button><small class="muted">可直接拖曳印章；跨頁公文可拖到任一頁</small>' : '<small class="muted">不加章：印出後蓋實體章</small>');
    gwPick = { d: d, copies: copies };
    (co.seal && d.sealOn ? loadImg(co.seal) : Promise.resolve()).then(function () { if (cur.v === 'docPrint' && cur.did === d.id) paint(); });
    paint();
    $('pvParts').hidden = true; $('btnMovePv').hidden = true; $('outAt').closest('.out-at').hidden = true;
    $('printHint').textContent = '「另存 PDF」每份各存一個檔案（檔名加上正本／副本與單位）' + (window.Cloud ? '，正本同時歸檔到雲端「監造公文／' + (p.info.name || '') + '／發文」' : '') + '。用瀏覽器列印時，紙張選 A4 直式。';
    gwFontCheck().then(function (ok) { if (cur.v === 'docPrint') gwFontBar(ok); });
    show('vPreview');
    setDocName(docFileName(d));
  }
  $('gwBar').addEventListener('change', function (e) {
    var d = gwPick && gwPick.d; if (!d) return;
    if (e.target.dataset.copy != null) {
      var u = gwPick.copies[+e.target.dataset.copy].unit, off = d.copiesOff = d.copiesOff || [];
      if (e.target.checked) d.copiesOff = off.filter(function (x) { return x !== u; }); else if (off.indexOf(u) < 0) off.push(u);
    }
    if (e.target.id === 'gwSealOn') d.sealOn = e.target.checked;
    save(); renderDocPrint();
  });
  $('gwBar').addEventListener('click', function (e) { if (e.target.id === 'gwSealReset' && gwPick) { delete gwPick.d.sealPos; save(); renderDocPrint(); } });
  /* 拖曳印章：位置以 mm 記在公文上，所有份數共用 */
  (function () {
    var drag = null;
    $('paper').addEventListener('pointerdown', function (e) {
      var img = e.target.closest && e.target.closest('.gw-seal-abs'); if (!img || cur.v !== 'docPrint') return;
      e.preventDefault(); img.setPointerCapture(e.pointerId);
      var r = img.getBoundingClientRect(); drag = { img: img, dx: e.clientX - r.left, dy: e.clientY - r.top };
    });
    $('paper').addEventListener('pointermove', function (e) {
      if (!drag) return;
      var sheets = $('paper').querySelectorAll('.gw'), host = null;
      Array.prototype.forEach.call(sheets, function (g) { var r = g.getBoundingClientRect(); if (e.clientY >= r.top && e.clientY <= r.bottom) host = g; });
      host = host || drag.img.parentNode;
      if (host !== drag.img.parentNode) host.appendChild(drag.img);
      var hr = host.getBoundingClientRect(), k = hr.width / 210;
      var x = Math.max(0, Math.min(210 - 20, (e.clientX - drag.dx - hr.left) / k)), y = Math.max(0, Math.min(297 - 20, (e.clientY - drag.dy - hr.top) / k));
      drag.img.style.left = x.toFixed(1) + 'mm'; drag.img.style.top = y.toFixed(1) + 'mm'; drag.pos = { pg: +host.dataset.pg, x: +x.toFixed(1), y: +y.toFixed(1) };
    });
    var end = function () { if (!drag) return; if (drag.pos && gwPick) { gwPick.d.sealPos = drag.pos; save(); renderDocPrint(); } drag = null; };
    $('paper').addEventListener('pointerup', end); $('paper').addEventListener('pointercancel', end);
  })();
  /* 另存 PDF：每一份各自一個檔案；正本歸檔雲端 */
  function docPdfAll() {
    var p = proj(cur.pid), d = docById(p, cur.did), co = company(), pages = gwPaginate(d), paper = $('paper'), keep = paper.innerHTML;
    var copies = docCopies(d).filter(function (c) { return (d.copiesOff || []).indexOf(c.unit) < 0; });
    if (!copies.length) copies = docCopies(d).slice(0, 1);
    var base = docFileName(d), msgs = [], archived = false;
    return copies.reduce(function (pr, c) {
      return pr.then(function () {
        paper.innerHTML = gwCopyHtml(d, co, c, pages);
        return makePdf().then(function (blob) {
          var name = base + '_' + c.kind + '_' + c.unit + '.pdf', a = document.createElement('a');
          a.href = URL.createObjectURL(blob); a.download = name.replace(/[\\/:*?"<>|]+/g, '_'); document.body.appendChild(a); a.click(); a.remove();
          setTimeout(function () { URL.revokeObjectURL(a.href); }, 60000);
          msgs.push(name);
          if (c.kind === '正本' && !archived) { archived = true; archiveDoc(blob, base + '.pdf'); }
          return new Promise(function (res) { setTimeout(res, 400); });
        });
      });
    }, Promise.resolve()).then(function () { paper.innerHTML = keep; return msgs; });
  }
  function archiveDoc(blob, fname) {
    var p = proj(cur.pid), d = docById(p, cur.did); if (!d || !window.Cloud || !window.Cloud.putDoc) return;
    var rd = new FileReader();
    rd.onload = function () {
      window.Cloud.putDoc({ proj: p.info.name, dir: '發文', name: fname, data: String(rd.result).split(',')[1], mime: 'application/pdf', old: d.fileId || '' }).then(function (res) {
        d.fileId = res.id; d.url = res.url; d.archivedAt = Date.now(); save();
        $('pdfMsg').textContent += ' 正本已歸檔到雲端「監造公文／' + (p.info.name || '') + '／發文」。';
      }).catch(function (e) {
        $('pdfMsg').textContent += /putDoc|not found|找不到|Script function/i.test(e.message) ? ' 雲端歸檔需要先更新 Apps Script（依說明重新部署）。' : ' 雲端歸檔失敗：' + e.message;
      });
    };
    rd.readAsDataURL(blob);
  }
  /* 公文字型：先用本機字型名稱（local()）；Chrome 等瀏覽器讀不到時，改讀取本機標楷體字型檔，存在這台裝置（不上傳） */
  var GW_PS = 'DFKaiShu-SB-Estd-BF', gwFontOk = null;
  /* Windows：標楷體（DFKai-SB）；Mac：標楷體-繁（BiauKai TC，由字體簿下載） */
  var GW_LOCAL = ['DFKaiShu-SB-Estd-BF', 'BiauKaiTC', 'BiauKaiTC-Regular', 'BiauKai TC', 'BiauKai TC Regular', '標楷體-繁', 'BiauKai', '標楷體', 'DFKai-SB', 'BiauKaiHK', 'BiauKai HK', '標楷體-港澳'].map(function (n) { return 'local("' + n + '")'; }).join(', ');
  function gwLocalOk() {
    if (!window.FontFace) return Promise.resolve(false);
    return new FontFace('gw-kai-test', GW_LOCAL).load().then(function () { return true; }, function () { return false; });
  }
  function gwUseBuf(buf) {
    var face = new FontFace('GW Kai', sfntFrom(buf, GW_PS));
    return face.load().then(function (f) { document.fonts.add(f); return true; });
  }
  function gwFontCheck() {
    if (gwFontOk) return Promise.resolve(true);
    return gwLocalOk().then(function (ok) {
      if (ok) return (gwFontOk = true);
      return IDB.get('font:gwkai').then(function (buf) { return buf ? gwUseBuf(buf) : false; }).catch(function () { return false; }).then(function (r) { return (gwFontOk = !!r); });
    });
  }
  /* 字型集合檔（.ttc）取出指定的字型；一般 .ttf/.otf 原樣使用 */
  function sfntFrom(buf, ps) {
    var dv = new DataView(buf);
    if (dv.getUint32(0) !== 0x74746366) return buf;   // 'ttcf'
    var n = dv.getUint32(8), offs = [], pick = null;
    for (var i = 0; i < n; i++) offs.push(dv.getUint32(12 + i * 4));
    var face = function (o) {
      var nt = dv.getUint16(o + 4), tabs = [];
      for (var t = 0; t < nt; t++) { var r = o + 12 + t * 16; tabs.push({ tag: dv.getUint32(r), sum: dv.getUint32(r + 4), off: dv.getUint32(r + 8), len: dv.getUint32(r + 12) }); }
      return { ver: dv.getUint32(o), tabs: tabs };
    };
    var psName = function (f) {
      var nt = f.tabs.filter(function (t) { return t.tag === 0x6e616d65; })[0]; if (!nt) return '';
      var cnt = dv.getUint16(nt.off + 2), so = nt.off + dv.getUint16(nt.off + 4);
      for (var k = 0; k < cnt; k++) {
        var r = nt.off + 6 + k * 12, pid = dv.getUint16(r), id = dv.getUint16(r + 6), len = dv.getUint16(r + 8), off = dv.getUint16(r + 10);
        if (id !== 6) continue;
        var str = ''; if (pid === 1) { for (var c = 0; c < len; c++) str += String.fromCharCode(dv.getUint8(so + off + c)); } else { for (var c2 = 0; c2 < len; c2 += 2) str += String.fromCharCode(dv.getUint16(so + off + c2)); }
        return str;
      }
      return '';
    };
    var faces = offs.map(face);
    pick = faces.filter(function (f) { return psName(f) === ps; })[0] || faces[0];
    var nt2 = pick.tabs.length, head = 12 + nt2 * 16, size = head;
    pick.tabs.forEach(function (t) { size += (t.len + 3) & ~3; });
    var out = new ArrayBuffer(size), ov = new DataView(out), ob = new Uint8Array(out), src = new Uint8Array(buf), pos = head;
    ov.setUint32(0, pick.ver); ov.setUint16(4, nt2);
    var es = Math.floor(Math.log2(nt2)), sr = Math.pow(2, es) * 16; ov.setUint16(6, sr); ov.setUint16(8, es); ov.setUint16(10, nt2 * 16 - sr);
    pick.tabs.forEach(function (t, i) {
      var r = 12 + i * 16; ov.setUint32(r, t.tag); ov.setUint32(r + 4, t.sum); ov.setUint32(r + 8, pos); ov.setUint32(r + 12, t.len);
      ob.set(src.subarray(t.off, t.off + t.len), pos); pos += (t.len + 3) & ~3;
    });
    return out;
  }
  function gwFontBar(ok) {
    var bar = $('gwFontBar'); if (!bar) return;
    bar.hidden = false;
    bar.innerHTML = ok ? '✓ 字型：標楷體（與 Pages 範本相同）'
      : '⚠ 瀏覽器讀不到這台電腦的「標楷體」，目前顯示的是替代字型，輸出會與 Pages 不同。' +
        (window.queryLocalFonts ? '<button type="button" class="btn primary sm" id="gwFontQuery">讀取本機標楷體</button>' : '') +
        '<label class="btn ghost sm">選擇標楷體字型檔…<input type="file" id="gwFontFile" accept=".ttf,.ttc,.otf,font/*" hidden></label>' +
        '<small>只需設定一次，字型檔只存在這台裝置。Mac：「字體簿」找到「標楷體」→ 右鍵「在 Finder 中顯示」即是字型檔；Windows：C:\\Windows\\Fonts\\kaiu.ttf。</small>';
  }
  function gwSaveBuf(buf) {
    return gwUseBuf(buf).then(function () {
      gwFontOk = true; IDB.put('font:gwkai', buf).catch(function () {});
      if (cur.v === 'docPrint') { gwFontBar(true); renderDocPrint(); }
    }).catch(function () { $('gwFontBar').insertAdjacentHTML('beforeend', '<b class="sum-ng">這個檔案無法當作字型使用，請改選標楷體的字型檔。</b>'); });
  }
  document.addEventListener('click', function (e) {
    if (e.target.id !== 'gwFontQuery') return;
    window.queryLocalFonts().then(function (list) {
      var hit = function (re) { return list.filter(function (x) { return re.test(x.postscriptName + ' ' + x.fullName + ' ' + x.family); })[0]; };
      var f = hit(/DFKaiShu|DFKai-SB/) || hit(/BiauKaiTC|BiauKai TC|標楷體-繁/) || hit(/BiauKai|標楷體/);
      if (!f) throw new Error('none');
      return f.blob().then(function (b) { return b.arrayBuffer(); }).then(gwSaveBuf);
    }).catch(function () { $('gwFontBar').insertAdjacentHTML('beforeend', '<b class="sum-ng">讀取失敗（可能未允許或在共編版的框架內無法使用），請改用「選擇標楷體字型檔」。</b>'); });
  });
  document.addEventListener('change', function (e) {
    if (e.target.id !== 'gwFontFile' || !e.target.files[0]) return;
    e.target.files[0].arrayBuffer().then(gwSaveBuf);
  });

  /* ========== 監造日報：版面依同事系統輸出的「公共工程監造報表」Excel 逐格重現（欄寬、列高、合併、框線、字型）；日曆天每日一張 ========== */
  var DR_TPL = {"w":[17.33203125,5.33203125,10.1640625,10.6640625,7.1640625,4.0,9.33203125,4.1640625,18.6640625,18.83203125],"rows":[{"h":24.0,"cells":[{"c":0,"cs":10,"t":"公共工程監造報表","f":12.0,"b":1,"m":1,"h":"center","v":"center","bd":[2,2,2,0]}]},{"h":18.75,"cells":[{"c":0,"t":"報表編號：","f":12.0,"h":"center","v":"center","bd":[2,0,0,0]},{"c":1,"f":12.0,"h":"left","v":"center","k":"no"},{"c":2},{"c":3},{"c":4},{"c":5},{"c":6},{"c":7},{"c":8},{"c":9,"bd":[0,2,0,0]}]},{"h":24.0,"cells":[{"c":0,"t":"本日天氣：","f":12.0,"h":"center","v":"center","bd":[2,0,0,0]},{"c":1,"t":"上午:","f":12.0,"h":"left","v":"center"},{"c":2,"f":12.0,"v":"center","k":"wAm"},{"c":3,"t":"下午：","f":12.0,"v":"center"},{"c":4,"f":12.0,"v":"center","k":"wPm"},{"c":5},{"c":6,"cs":2,"t":"填報日期：","f":12.0,"h":"right","v":"center","bd":[0,0,0,1]},{"c":8,"f":12.0,"h":"left","v":"center","bd":[0,0,0,1],"k":"date"},{"c":9,"f":12.0,"h":"left","v":"center","bd":[0,2,0,1],"k":"wd"}]},{"h":24.75,"cells":[{"c":0,"cs":2,"t":"工程名稱","f":12.0,"h":"center","v":"center","bd":[2,0.5,1,0.5]},{"c":2,"cs":8,"f":12.0,"h":"left","v":"center","w":1,"bd":[0.5,0.5,1,0.5],"k":"name"}]},{"h":23.25,"cells":[{"c":0,"cs":2,"t":"契約工期","f":12.0,"h":"center","v":"center","w":1,"bd":[2,0.5,0.5,0.5]},{"c":2,"cs":2,"t":"開工日期","f":12.0,"h":"center","v":"center","bd":[0.5,0.5,0.5,0.5]},{"c":4,"cs":2,"t":"累計工期","f":12.0,"h":"center","v":"center","bd":[0.5,0.5,0.5,0.5]},{"c":6,"cs":2,"t":"剩餘工期","f":12.0,"h":"center","v":"center","bd":[0.5,0.5,0.5,0.5]},{"c":8,"t":"預定完工日期","f":12.0,"h":"center","v":"center","bd":[0.5,0.5,0.5,0.5]},{"c":9,"t":"實際完工日期","f":12.0,"h":"center","v":"center","bd":[0.5,2,0.5,0.5]}]},{"h":23.25,"cells":[{"c":0,"f":12.0,"h":"right","v":"center","w":1,"bd":[2,0,0.5,0.5],"k":"dur"},{"c":1,"t":"天","f":12.0,"h":"left","v":"center","w":1,"bd":[0,0.5,0.5,0.5]},{"c":2,"cs":2,"f":12.0,"h":"center","v":"center","bd":[0.5,0.5,0.5,0.5],"k":"start"},{"c":4,"f":12.0,"h":"right","v":"center","bd":[0.5,0,0.5,0.5],"k":"elapsed"},{"c":5,"t":"天","f":12.0,"v":"center","bd":[0,0.5,0.5,0.5]},{"c":6,"f":12.0,"h":"right","v":"center","bd":[0.5,0,0.5,0.5],"k":"remain"},{"c":7,"t":"天","f":12.0,"v":"center","bd":[0,0.5,0.5,0.5]},{"c":8,"f":12.0,"h":"center","v":"center","bd":[0.5,0.5,0.5,0.5],"k":"planEnd"},{"c":9,"f":12.0,"h":"center","v":"center","bd":[0.5,2,0.5,0.5],"k":"actEnd"}]},{"h":28.0,"cells":[{"c":0,"t":"契約變更次數：","f":12.0,"h":"right","v":"center","bd":[2,0.5,0.5,0.5]},{"c":1,"f":12.0,"h":"center","v":"center","bd":[0.5,0,0.5,0.5],"k":"changes"},{"c":2,"t":"次","f":12.0,"h":"left","v":"center","bd":[0,0.5,0.5,0.5]},{"c":3,"cs":3,"t":"工期展延天數：","f":12.0,"h":"right","v":"center","bd":[0.5,0.5,0.5,0.5]},{"c":6,"f":12.0,"h":"center","v":"center","bd":[0.5,0,0.5,0.5],"k":"ext"},{"c":7,"t":"天","f":12.0,"h":"left","v":"center","bd":[0,0.5,0.5,0.5]},{"c":8,"rs":3,"t":"契約金額","f":12.0,"h":"center","v":"center","bd":[0.5,0.5,0.5,0.5]},{"c":9,"f":12.0,"h":"left","v":"center","bd":[0.5,2,0.5,0.5],"k":"amount"}]},{"h":14.0,"cells":[{"c":0,"rs":2,"t":"預定進度（％）","f":12.0,"h":"right","v":"center","bd":[2,0.5,0.5,0.5]},{"c":1,"cs":2,"rs":2,"f":12.0,"h":"center","v":"center","bd":[0.5,0.5,0.5,0.5],"k":"plan"},{"c":3,"cs":3,"rs":2,"t":"實際進度（％）","f":12.0,"h":"right","v":"center","bd":[0.5,0.5,0.5,0.5]},{"c":6,"cs":2,"rs":2,"f":12.0,"h":"center","v":"center","bd":[0.5,0.5,0.5,0.5],"k":"act"},{"c":9,"rs":2,"bd":[0.5,2,0.5,0.5]}]},{"h":14.0,"cells":[]},{"h":18.0,"cells":[{"c":0,"cs":10,"t":"一、工程進行情況(含約定之重要施工項目及數量)：(詳施工日誌)","f":11.0,"h":"left","v":"center","bd":[2,0.5,1,0]}]},{"h":15,"cells":[{"c":0,"cs":10,"f":11.0,"h":"left","v":"top","bd":[2,2,0,0],"rs":6,"k":"s1"}]},{"h":15,"cells":[]},{"h":15,"cells":[]},{"h":15,"cells":[]},{"h":15,"cells":[]},{"h":15,"cells":[]},{"h":18.0,"cells":[{"c":0,"cs":10,"t":"二、監督依照設計圖說及核定施工圖說施工（含約定之檢驗停留點及施工抽查等情形）：","f":11.0,"h":"left","v":"center","bd":[2,0.5,1,0]}]},{"h":15,"cells":[{"c":0,"cs":10,"f":11.0,"v":"top","bd":[2,2,0,1],"rs":4,"k":"s2"}]},{"h":16.5,"cells":[]},{"h":15,"cells":[]},{"h":15,"cells":[]},{"h":18.0,"cells":[{"c":0,"cs":10,"t":"三、查核材料規格及品質(含約定之檢驗停留點、材料設備管制及檢(試)驗等抽驗情形)：","f":11.0,"h":"left","v":"center","bd":[2,0.5,1,0]}]},{"h":16.5,"cells":[{"c":0,"cs":10,"f":11.0,"v":"top","bd":[2,2,0,1],"rs":4,"k":"s3"}]},{"h":16.5,"cells":[]},{"h":16.5,"cells":[]},{"h":16.5,"cells":[]},{"h":16.5,"cells":[{"c":0,"cs":10,"t":"四、督導工地職業安全衛生事項：","f":11.0,"v":"center","bd":[2,0.5,1,0]}]},{"h":16.5,"cells":[{"c":0,"cs":5,"t":"（一）施工廠商施工前檢查事項辦理情形：","f":11.0,"h":"left","v":"top","w":1,"bd":[2,0,0,0]},{"c":5,"cs":4,"f":11.0,"h":"left","v":"top","w":1,"k":"x4a"},{"c":9,"bd":[0,2,0,0]}]},{"h":16.5,"cells":[{"c":0,"cs":10,"bd":[2,2,0,0],"k":"x4a2"}]},{"h":16.5,"cells":[{"c":0,"cs":4,"t":"（二）其他工地安全衛生督導事項：","f":11.0,"h":"left","v":"top","w":1,"bd":[2,0,0,0]},{"c":4,"cs":6,"f":11.0,"h":"left","v":"top","w":1,"bd":[0,2,0,0],"k":"x4b"}]},{"h":16.5,"cells":[{"c":0,"cs":10,"bd":[2,2,0,1],"k":"x4b2"}]},{"h":18.0,"cells":[{"c":0,"cs":10,"t":"五、其他約定監造事項(含重要事項紀錄、主管機關指示及通知廠商辦理事項等)：","f":11.0,"h":"left","v":"center","bd":[2,2,1,0]}]},{"h":88.0,"cells":[{"c":0,"cs":10,"f":11.0,"h":"left","v":"top","w":1,"bd":[2,2,0,1],"k":"s5"}]},{"h":130.0,"cells":[{"c":0,"cs":10,"t":"監造單位簽章：","f":11.0,"h":"left","v":"top","bd":[2,2,1,1]}]},{"h":125.0,"cells":[{"c":0,"cs":10,"t":"註：\n1.監造報告表原則應包含上述欄位；惟若上述欄位之內容業詳載於廠商填報之施工日誌，並按時陳報監造單位核備者，則監造報表之該等欄位可載明參詳施工日誌。\n2.本表原則應按日填寫，機關另有規定者，從其規定；若屬委外監造之工程，則一律按日填寫。未達新臺幣五千萬元或工期為九十日曆天以下之工程，得由機關統一訂定內部稽查程序及監造報告表之填報方式與周期。\n3.本監造報告表格式僅供參考，各機關亦得依契約約定事項，自行增訂之。\n4.契約工期如有修正，應填修正後之契約工期，含展延工期及不計工期天數；如有依契約變更設計，預定進度及實際進度應填變更設計後計算之進度。\n5.公共工程屬建築物者，仍應依本表辦理。惟該工程之監造人（建築師），應另依內政部最新訂頒之「建築物（監督、查核）報告表」填報。","f":10.0,"h":"left","v":"top","w":1,"bd":[2,2,1,2]}]},{"h":18.0,"cells":[{"c":0,"bd":[0,0,2,0]},{"c":1,"bd":[0,0,2,0]},{"c":2,"bd":[0,0,2,0]},{"c":3,"bd":[0,0,2,0]},{"c":4,"bd":[0,0,2,0]},{"c":5,"bd":[0,0,2,0]},{"c":6,"bd":[0,0,2,0]},{"c":7,"bd":[0,0,2,0]},{"c":8,"bd":[0,0,2,0]},{"c":9,"bd":[0,0,2,0]}]},{"h":15,"cells":[{"c":0},{"c":1},{"c":2},{"c":3},{"c":4},{"c":5},{"c":6},{"c":7},{"c":8},{"c":9}]}]};
  var DR_WEATHER = ['晴', '陰', '多雲', '雨', '陣雨', '颱風', '停工'];
  function dailies(p) { return p.dailies || (p.dailies = []); }
  function dailyOf(p, date) { return dailies(p).filter(function (x) { return x.date === date; })[0]; }
  function dailyMake(p, date) {
    var d = dailyOf(p, date); if (d) return d;
    var prev = dailies(p).filter(function (x) { return x.date < date; }).sort(function (a, b) { return a.date < b.date ? 1 : -1; })[0];
    d = { id: uid(), date: date, wAm: '', wPm: '', plan: prev ? prev.plan : '', act: prev ? prev.act : '', items: [], x2: '', x3: '', x4a: 'done', x4b: '', x5: '', off: {} };
    dailies(p).push(d); return d;
  }
  function drDays(a, b) { return Math.round((new Date(b + 'T00:00:00') - new Date(a + 'T00:00:00')) / 86400000); }
  function drInfo(p) {
    var I = p.info || {}, dur = +(I.changedDuration || I.duration) || 0;
    return { start: I.startDate || '', dur: dur, planEnd: I.changedEnd || I.planEnd || (I.startDate && dur ? addDays(I.startDate, dur - 1) : ''), amount: I.changedAmount || I.amount || '', changes: I.changeCount || 0, ext: I.extendDays || 0 };
  }
  /* 日期格式與同事 Excel 列印的 PDF 相同（西元） */
  function adSlash(d) { var a = (d || '').split('-'); return a.length === 3 ? a[0] + '/' + (+a[1]) + '/' + (+a[2]) : ''; }
  function adYmd(d) { var a = (d || '').split('-'); return a.length === 3 ? a[0] + '年' + (+a[1]) + '月' + (+a[2]) + '日' : ''; }
  function rocSlash(d) { var a = (d || '').split('-'); return a.length === 3 ? (a[0] - 1911) + '/' + (+a[1]) + '/' + (+a[2]) : ''; }
  function rocYmd(d) { var a = (d || '').split('-'); return a.length === 3 ? (a[0] - 1911) + '年' + (+a[1]) + '月' + (+a[2]) + '日' : ''; }
  function pct(v) { return v === '' || v == null || isNaN(+v) ? '' : (+v).toFixed(2) + '%'; }
  /* 當天自動帶入：二 施工抽查、三 材料、四 勞安、五 公文 */
  function drAuto(p, date) {
    var o = { s2: [], s3: [], s4: [], s5: [] }, matDone = {};
    var res = function (w, r) { var st = recState(w, r); return st === 'ok' ? '合格' : st === 'fixed' ? '缺失已改善' : st === 'ng' ? '有缺失' : '檢查中'; };
    var ngNames = function (w, r) {
      var xs = [];
      if (w.kind === 'irregular') (r.rows || []).forEach(function (x) { if (x.r === 'ng') xs.push(x.name); });
      else effGroups(w, r).forEach(function (g) { g.items.forEach(function (it) { if (it.c.r === 'ng') xs.push(it.name); }); });
      return xs.length ? '（' + xs.join('、') + '）' : '';
    };
    Object.keys(p.recs || {}).forEach(function (wid) {
      var w = WORK[wid]; if (!w || w.kind === 'photo') return;
      (p.recs[wid] || []).forEach(function (r) {
        if (r.draft) return;
        var loc = r.location ? '（' + r.location + '）' : '';
        if (r.checkDate === date) {
          if (w.kind === 'matform') { var m = matById(p, r.mid); matDone[r.mid] = 1; o.s3.push((m ? m.name : '材料') + '進場抽查' + loc + '：' + (res(w, r) === '有缺失' ? '不合格' + ngNames(w, r) : res(w, r))); }
          else if (w.kind === 'safety') o.s4.push('勞安衛生抽查' + loc + '：' + res(w, r) + (res(w, r) === '有缺失' ? ngNames(w, r) : ''));
          else o.s2.push((w.kind === 'phased' ? w.name + PH_SHORT[r.phase] + '抽查' : w.name) + loc + '：' + res(w, r) + (res(w, r) === '有缺失' ? ngNames(w, r) : ''));
        }
        var fd = isFix(w) ? r.fixDate : (r.reTime || '').slice(0, 10);
        if (fd === date && r.checkDate !== date && recState(w, r) === 'fixed') o[w.kind === 'safety' ? 's4' : w.kind === 'matform' ? 's3' : 's2'].push((w.kind === 'phased' ? w.name + PH_SHORT[r.phase] + '抽查' : w.name) + loc + '缺失複查：已完成改善');
      });
    });
    (p.mats || []).forEach(function (m) {
      if (m.arrivalDate === date) o.s3.push((m.name || '材料') + '進場');
      if (!matDone[m.id] && m.inspDate === date && m.inspResult) o.s3.push((m.name || '材料') + '進場抽查：' + (m.inspResult === 'ok' ? '合格' : '不合格'));
    });
    (p.docs || []).forEach(function (d) {
      if (d.date === date) o.s5.push(d.no + ' ' + shortSubj(d));
      if (d.kind === 'meet' && d.meet && (d.meet.time || '').slice(0, 10) === date) o.s5.push(d.meet.time.slice(11, 16) + ' ' + shortSubj(d));
    });
    return o;
  }
  function drCum(p, d, it) {
    if (it.cum !== '' && it.cum != null) return it.cum;
    var s = 0; dailies(p).forEach(function (x) { if (x.date <= d.date) (x.items || []).forEach(function (y) { if (y.name === it.name && y.unit === it.unit) s += +y.qty || 0; }); });
    return Math.round(s * 1000) / 1000;
  }
  function drValues(p, date) {
    var I = drInfo(p), d = dailyOf(p, date) || { items: [], off: {}, x4a: 'done' }, a = drAuto(p, date), off = d.off || {};
    var el = I.start ? drDays(I.start, date) + 1 : '';
    var lines = function (auto, key, extra) { var xs = (off[key] ? [] : auto).concat(lines0(extra)); return xs.length ? xs : ['無']; };
    var s1 = (d.items || []).filter(function (it) { return it.name; }).map(function (it) {
      return it.name + '    本日完成數量：' + (it.qty === '' || it.qty == null ? '' : it.qty) + (it.unit || '') + '    累計完成數量：' + drCum(p, d, it) + (it.unit || '');
    });
    var s4 = lines(a.s4, 's4', d.x4b);
    return {
      no: el ? el + ' ' : '', wAm: d.wAm || '', wPm: d.wPm || '', date: adYmd(date), wd: '星期' + '日一二三四五六'[new Date(date + 'T00:00:00').getDay()],
      name: p.info.name || '', dur: I.dur || '', start: adSlash(I.start), elapsed: el ? el + ' ' : '', remain: el && I.dur ? (I.dur - el) + ' ' : '', planEnd: adSlash(I.planEnd), actEnd: adSlash(p.info.actualEnd || ''),
      changes: +I.changes ? I.changes : '', ext: I.ext, amount: I.amount ? '原契約:' + Number(String(I.amount).replace(/,/g, '')).toLocaleString('en-US') : '', plan: pct(d.plan), act: pct(d.act),
      x4a: d.x4a === 'undone' ? '□完成 ■未完成' : '■完成 □未完成', x4a2: '', x4b: s4.join('；'), x4b2: '',
      s1: s1.length ? s1 : ['無'], s2: lines(a.s2, 's2', d.x2), s3: lines(a.s3, 's3', d.x3), s5: lines(a.s5, 's5', d.x5)
    };
  }
  function lines0(t) { return String(t || '').split('\n').map(function (x) { return x.replace(/\s+$/, ''); }).filter(function (x) { return x.trim(); }); }
  /* Excel 欄寬 → 像素（新細明體 12 的數字寬 8px），整張依「縮放成一頁」放進 A4 可列印範圍（左右 13mm、上下 10mm） */
  var DR_GEO = (function () {
    var px = DR_TPL.w.map(function (w) { return Math.floor(w * 8 + 5); }), wpt = px.reduce(function (s, x) { return s + x; }, 0) * 0.75;
    var hpt = DR_TPL.rows.reduce(function (s, r) { return s + r.h; }, 0), PT = 25.4 / 72;
    var k = Math.min(184 / (wpt * PT), 277 / (hpt * PT));
    /* 列高再乘 1.0436：依同事 Excel 實際列印的 PDF 量測校正（Excel 列高比換算值略高） */
    return { k: k, cols: px.map(function (x) { return x * 0.75 * PT * k; }), rows: DR_TPL.rows.map(function (r) { return r.h * PT * k * 1.0436; }) };
  })();
  function drSheet(p, date) {
    var v = drValues(p, date), G = DR_GEO, BW = { 2: 0.45, 1: 0.25, 0.5: 0.12 };
    var h = '<table class="dr-t" style="width:' + G.cols.reduce(function (s, x) { return s + x; }, 0).toFixed(2) + 'mm"><colgroup>' + G.cols.map(function (w) { return '<col style="width:' + w.toFixed(2) + 'mm">'; }).join('') + '</colgroup>';
    DR_TPL.rows.forEach(function (r, ri) {
      h += '<tr style="height:' + G.rows[ri].toFixed(2) + 'mm">';
      r.cells.forEach(function (c) {
        var st = [], bd = c.bd || [0, 0, 0, 0], sides = ['left', 'right', 'top', 'bottom'];
        bd.forEach(function (b, i) { if (b) st.push('border-' + sides[i] + ':' + BW[b] + 'mm solid #000'); });
        st.push('font-size:' + ((c.f || 12) * G.k * 1.0436).toFixed(2) + 'pt');
        if (c.b) st.push('font-weight:700');
        if (c.m) st.push("font-family:PMingLiU,MingLiU,'Songti TC','Noto Serif CJK TC',serif");
        st.push('text-align:' + (c.h === 'center' ? 'center' : c.h === 'right' ? 'right' : 'left'));
        st.push('vertical-align:' + (c.v === 'top' ? 'top' : c.v === 'bottom' ? 'bottom' : 'middle'));
        var txt = '', cls = '';
        if (c.k) {
          var val = v[c.k];
          if (Array.isArray(val)) {
            var hh = 0; for (var q = ri; q < ri + (c.rs || 1); q++) hh += G.rows[q];
            txt = '<div class="dr-fit" style="height:' + (hh - 0.4).toFixed(2) + 'mm">' + val.map(esc).join('<br>') + '</div>';
          } else txt = esc(val);
        } else if (c.t) txt = esc(c.t).replace(/\n/g, '<br>');
        if (c.w || (c.k && /^(s5|x4b|name)$/.test(c.k))) cls = ' class="dr-w"';
        h += '<td' + (c.cs ? ' colspan="' + c.cs + '"' : '') + (c.rs ? ' rowspan="' + c.rs + '"' : '') + cls + ' style="' + st.join(';') + '">' + txt + '</td>';
      });
      h += '</tr>';
    });
    return h + '</table>';
  }
  /* 內容超過格子時縮小字級，格子不變 */
  function drFit(root) {
    root.querySelectorAll('.dr-fit').forEach(function (el) {
      var f = parseFloat(getComputedStyle(el.parentNode).fontSize), n = 0;
      while (el.scrollHeight > el.clientHeight + 1 && n++ < 30) { f *= 0.94; el.style.fontSize = f + 'px'; el.style.lineHeight = '1.15'; }
    });
  }
  function drPages(p, dates) { return dates.map(function (dt, i) { return '<section class="dr-page' + (i ? ' rec-sec' : '') + '"><div class="dr-sheet">' + drSheet(p, dt) + '</div></section>'; }).join(''); }
  /* ---- 月表 ---- */
  $('btnDaily').addEventListener('click', function () { var t = today(); go({ v: 'daily', pid: cur.pid, ym: t.slice(0, 7) }); });
  function monthDays(ym) { var a = ym.split('-'), n = new Date(+a[0], +a[1], 0).getDate(), out = []; for (var i = 1; i <= n; i++) out.push(ym + '-' + ('0' + i).slice(-2)); return out; }
  function drRange(p, ym) {
    var I = drInfo(p);
    return monthDays(ym).filter(function (d) { return (!I.start || d >= I.start) && (!p.info.actualEnd || d <= p.info.actualEnd || dailyOf(p, d)) && (d <= today() || dailyOf(p, d) || !I.start); });
  }
  function renderDaily() {
    var p = proj(cur.pid), ym = cur.ym || today().slice(0, 7), I = drInfo(p), a = ym.split('-');
    topbar('監造日報', p.info.name || '未命名工程', '工程');
    $('drMonth').textContent = (a[0] - 1911) + ' 年 ' + (+a[1]) + ' 月';
    $('drWarn').hidden = !!(I.start && I.dur);
    var days = drRange(p, ym), td = today();
    $('drList').innerHTML = days.length ? days.map(function (dt) {
      var d = dailyOf(p, dt), au = drAuto(p, dt), n = I.start ? drDays(I.start, dt) + 1 : '';
      var cnt = au.s2.length + au.s3.length + au.s4.length + au.s5.length;
      return '<button type="button" class="dr-row' + (dt === td ? ' today' : '') + '" data-dr="' + dt + '"><span class="dr-d">' + (+dt.slice(5, 7)) + '/' + (+dt.slice(8)) + '<small>' + '日一二三四五六'[new Date(dt + 'T00:00:00').getDay()] + '</small></span>' +
        '<span class="dr-no">' + (n ? '第' + n + '號' : '') + '</span><span class="dr-w">' + esc([d && d.wAm, d && d.wPm].filter(Boolean).join('/') || '—') + '</span>' +
        '<span class="dr-i">' + (d && d.items.length ? d.items.length + ' 項施工' : '無施工日誌') + (cnt ? '・帶入 ' + cnt + ' 筆' : '') + '</span>' +
        '<span class="dr-pc">' + (d && (d.plan || d.act) ? pct(d.plan) + ' / ' + pct(d.act) : '') + '</span><span class="dr-st ' + (d ? 'ok' : '') + '">' + (d ? '✓ 已填' : '○ 未填') + '</span><span class="chev">›</span></button>';
    }).join('') : '<div class="empty">' + (I.start ? '本月不在工期內。' : '請先在工程資料填寫開工日期與契約工期。') + '</div>';
    show('vDaily');
  }
  $('drPrev').addEventListener('click', function () { var a = (cur.ym || today().slice(0, 7)).split('-'), d = new Date(+a[0], +a[1] - 2, 1); cur.ym = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2); go(cur, true); });
  $('drNext').addEventListener('click', function () { var a = (cur.ym || today().slice(0, 7)).split('-'), d = new Date(+a[0], +a[1], 1); cur.ym = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2); go(cur, true); });
  $('drList').addEventListener('click', function (e) { var b = e.target.closest('[data-dr]'); if (b) go({ v: 'dailyEdit', pid: cur.pid, date: b.dataset.dr }); });
  $('drImportBtn').addEventListener('click', function () { go({ v: 'drImport', pid: cur.pid }); });
  $('drPrintMonth').addEventListener('click', function () { go({ v: 'dailyPrint', pid: cur.pid, ym: cur.ym || today().slice(0, 7) }); });
  $('drInfoBtn').addEventListener('click', function () { go({ v: 'project', pid: cur.pid }); setTimeout(function () { var b = $('infoCard'); if (b) { b.open = true; b.scrollIntoView({ block: 'start' }); } }, 50); });
  /* ---- 單日編輯 ---- */
  function renderDailyEdit() {
    var p = proj(cur.pid), dt = cur.date, d = dailyOf(p, dt) || { items: [], off: {}, x4a: 'done', wAm: '', wPm: '', plan: '', act: '', x2: '', x3: '', x4b: '', x5: '' };
    var I = drInfo(p), n = I.start ? drDays(I.start, dt) + 1 : '', au = drAuto(p, dt), off = d.off || {};
    topbar('監造日報' + (n ? '（第' + n + '號）' : ''), rocYmd(dt) + ' 星期' + '日一二三四五六'[new Date(dt + 'T00:00:00').getDay()], '日報月表');
    var wsel = function (k) { return '<select data-dk="' + k + '"><option value=""></option>' + DR_WEATHER.map(function (w) { return '<option' + (d[k] === w ? ' selected' : '') + '>' + w + '</option>'; }).join('') + '</select>'; };
    var auto = function (key, list, label) {
      return '<div class="dr-auto"><div class="dr-auto-h"><span>自動帶入（' + label + '）</span>' + (list.length ? '<label><input type="checkbox" data-doff="' + key + '"' + (off[key] ? '' : ' checked') + '>列入報表</label>' : '') + '</div>' +
        (list.length ? '<ul' + (off[key] ? ' class="off"' : '') + '>' + list.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul>' : '<p class="muted">這天沒有相關紀錄。</p>') + '</div>';
    };
    var items = (d.items || []).map(function (it, i) {
      return '<div class="dr-item" data-di="' + i + '"><input data-if="name" placeholder="工程項目" value="' + esc(it.name) + '"><input data-if="qty" inputmode="decimal" placeholder="本日數量" value="' + esc(it.qty) + '"><input data-if="unit" placeholder="單位" value="' + esc(it.unit) + '">' +
        '<input data-if="cum" inputmode="decimal" placeholder="累計（空白自動加總）" value="' + esc(it.cum) + '"><button type="button" class="link" data-idel="' + i + '">刪除</button></div>';
    }).join('');
    $('drEditBody').innerHTML =
      '<section class="card"><div class="grid g3"><label class="f">天氣（上午）' + wsel('wAm') + '</label><label class="f">天氣（下午）' + wsel('wPm') + '</label><span></span>' +
      '<label class="f">預定進度（%）<input data-dk="plan" inputmode="decimal" value="' + esc(d.plan) + '"></label><label class="f">實際進度（%）<input data-dk="act" inputmode="decimal" value="' + esc(d.act) + '"></label>' +
      '<p class="muted" style="align-self:end">進度以施工日誌為準。</p></div></section>' +
      '<section class="card"><h2 class="h small">一、工程進行情況（含約定之重要施工項目及數量）</h2><div class="dr-items">' + (items || '<p class="muted">尚無施工項目。下一階段可從施工日誌匯入；目前可手動新增。</p>') + '</div>' +
      '<button type="button" class="btn ghost sm" id="drAddItem">＋ 新增施工項目</button></section>' +
      '<section class="card"><h2 class="h small">二、監督依照設計圖說施工（含檢驗停留點及施工抽查）</h2>' + auto('s2', au.s2, '施工抽查') + '<label class="f">補充<textarea data-dk="x2" rows="2">' + esc(d.x2) + '</textarea></label></section>' +
      '<section class="card"><h2 class="h small">三、查核材料規格及品質</h2>' + auto('s3', au.s3, '材料進場與抽查') + '<label class="f">補充<textarea data-dk="x3" rows="2">' + esc(d.x3) + '</textarea></label></section>' +
      '<section class="card"><h2 class="h small">四、督導工地職業安全衛生事項</h2><div class="seg"><label class="pill"><input type="radio" name="drx4a" value="done"' + (d.x4a !== 'undone' ? ' checked' : '') + '><span>（一）施工前檢查：完成</span></label><label class="pill"><input type="radio" name="drx4a" value="undone"' + (d.x4a === 'undone' ? ' checked' : '') + '><span>未完成</span></label></div>' +
      auto('s4', au.s4, '勞安衛生抽查') + '<label class="f">（二）其他督導事項補充<textarea data-dk="x4b" rows="2">' + esc(d.x4b) + '</textarea></label></section>' +
      '<section class="card"><h2 class="h small">五、其他約定監造事項</h2>' + auto('s5', au.s5, '當日公文、會議') + '<label class="f">補充（重要事項、主管機關指示、通知廠商事項）<textarea data-dk="x5" rows="3">' + esc(d.x5) + '</textarea></label></section>' +
      '<div class="doc-acts"><button type="button" class="btn ghost" id="drPrevDay">‹ 前一天</button><button type="button" class="btn primary" id="drPv">預覽／輸出 PDF</button><button type="button" class="btn ghost" id="drNextDay">後一天 ›</button></div>';
    show('vDailyEdit');
  }
  function drCur() { var p = proj(cur.pid); return dailyMake(p, cur.date); }
  $('drEditBody').addEventListener('input', function (e) {
    var t = e.target, d;
    if (t.dataset.dk && t.tagName !== 'SELECT') { d = drCur(); d[t.dataset.dk] = t.value.trim() === '' && /plan|act/.test(t.dataset.dk) ? '' : (/plan|act/.test(t.dataset.dk) ? t.value.trim() : t.value); save(); }
    if (t.dataset.if) { d = drCur(); var it = d.items[+t.closest('[data-di]').dataset.di]; it[t.dataset.if] = t.value.trim(); save(); }
  });
  $('drEditBody').addEventListener('change', function (e) {
    var t = e.target, d = drCur();
    if (t.tagName === 'SELECT' && t.dataset.dk) { d[t.dataset.dk] = t.value; save(); }
    if (t.name === 'drx4a') { d.x4a = t.value; save(); }
    if (t.dataset.doff) { d.off = d.off || {}; if (t.checked) delete d.off[t.dataset.doff]; else d.off[t.dataset.doff] = 1; save(); renderDailyEdit(); }
  });
  $('drEditBody').addEventListener('click', function (e) {
    var t = e.target;
    if (t.id === 'drAddItem') { var d = drCur(); d.items.push({ name: '', qty: '', unit: '', cum: '' }); save(); renderDailyEdit(); var xs = $('drEditBody').querySelectorAll('[data-if="name"]'); if (xs.length) xs[xs.length - 1].focus(); }
    if (t.dataset.idel != null) { var d2 = drCur(); d2.items.splice(+t.dataset.idel, 1); save(); renderDailyEdit(); }
    if (t.id === 'drPv') { drCur(); save(); go({ v: 'dailyPrint', pid: cur.pid, date: cur.date }); }
    if (t.id === 'drPrevDay' || t.id === 'drNextDay') { cur.date = addDays(cur.date, t.id === 'drPrevDay' ? -1 : 1); go(cur, true); }
  });
  /* ---- 輸出 ---- */
  function renderDailyPrint() {
    var p = proj(cur.pid), paper = $('paper'), dates = cur.date ? [cur.date] : drRange(p, cur.ym);
    topbar('列印預覽', cur.date ? rocYmd(cur.date) + ' 監造日報' : (+cur.ym.slice(0, 4) - 1911) + ' 年 ' + (+cur.ym.slice(5)) + ' 月監造日報（' + dates.length + ' 張）', cur.date ? '返回修改' : '日報月表');
    paper.className = 'paper dr-paper';
    paper.innerHTML = dates.length ? drPages(p, dates) : '<p class="pv-empty">這個月沒有在工期內的日期。</p>';
    $('pvParts').hidden = true; $('btnMovePv').hidden = true; $('outAt').closest('.out-at').hidden = true;
    $('printHint').textContent = '版面依公共工程監造報表（A4 直式）。用瀏覽器列印時，紙張選 A4、邊界「無」或「預設」皆可。';
    gwFontCheck().then(function (ok) { if (/^daily/.test(cur.v)) gwFontBar(ok); });
    show('vPreview');
    drFit(paper);
    var a = (cur.date || cur.ym + '-01').split('-');
    setDocName(cur.date ? '監造日報_' + (a[0] - 1911) + '.' + a[1] + a[2] + '_' + (p.info.name || '') : '監造日報_' + (a[0] - 1911) + '年' + (+a[1]) + '月_' + (p.info.name || ''));
  }

  /* ========== 施工日誌匯入（照同事系統：選工程 → 載入 Excel、解析引擎與關鍵字 → 預覽 → 確認寫入） ========== */
  var DRI_DEF = { maxCol: 50, date: '日期,施工日期,中華民國,填表日期', item: '工程項目,施工項目,工作項目,項目名稱', qty: '本日完成數量,數量,今日數量,施作數量', unit: '單位', weather: '天氣,氣候', black: '品質管理,品管費,利潤,稅雜費,稅管費,管理費,保險費,營業稅,小計,合計,總計', zero: true, engine: 'auto', dup: 'over' };
  function driCfg() { var c = {}; try { c = JSON.parse(lsGet('dr-import-cfg') || '{}'); } catch (e) {} return Object.assign({}, DRI_DEF, c); }
  var driFiles = [], driResult = null;
  function renderDrImport() {
    var p = proj(cur.pid), c = driCfg();
    topbar('匯入施工日誌', p.info.name || '', '日報月表');
    $('driProj').innerHTML = S.projects.map(function (x) { return '<option value="' + x.id + '"' + (x.id === p.id ? ' selected' : '') + '>' + esc(x.info.name || '未命名工程') + '</option>'; }).join('');
    $('driDup').value = c.dup; $('driEngine').value = c.engine; $('driMax').value = c.maxCol;
    ['date', 'item', 'qty', 'unit', 'weather', 'black'].forEach(function (k) { $('dri_' + k).value = c[k]; });
    $('driZero').checked = !!c.zero;
    $('driFilesInfo').textContent = driFiles.length ? '已選 ' + driFiles.length + ' 個檔案：' + driFiles.map(function (f) { return f.name; }).join('、') : '';
    $('driGo').disabled = !driFiles.length;
    driPaint();
    show('vDrImport');
  }
  function driSaveCfg() {
    var c = { dup: $('driDup').value, engine: $('driEngine').value, maxCol: +$('driMax').value || 50, zero: $('driZero').checked };
    ['date', 'item', 'qty', 'unit', 'weather', 'black'].forEach(function (k) { c[k] = $('dri_' + k).value; });
    lsSet('dr-import-cfg', JSON.stringify(c)); return Object.assign({}, DRI_DEF, c);
  }
  ['driDup', 'driEngine', 'driMax', 'driZero', 'dri_date', 'dri_item', 'dri_qty', 'dri_unit', 'dri_weather', 'dri_black'].forEach(function (id) { $(id).addEventListener('change', driSaveCfg); });
  $('driProj').addEventListener('change', function () { cur.pid = $('driProj').value; go(cur, true); });
  function driPick(files) { driFiles = Array.prototype.slice.call(files || []).filter(function (f) { return /\.(xlsx|xls|xlsm)$/i.test(f.name); }); driResult = null; renderDrImport(); }
  $('driFile').addEventListener('change', function () { driPick($('driFile').files); $('driFile').value = ''; });
  $('driDrop').addEventListener('dragover', function (e) { e.preventDefault(); $('driDrop').classList.add('over'); });
  $('driDrop').addEventListener('dragleave', function () { $('driDrop').classList.remove('over'); });
  $('driDrop').addEventListener('drop', function (e) { e.preventDefault(); $('driDrop').classList.remove('over'); driPick(e.dataTransfer.files); });
  $('driDrop').addEventListener('click', function (e) { if (e.target.id !== 'driFile') $('driFile').click(); });
  /* ---- 解析 ---- */
  function kws(s) { return String(s || '').split(/[,，、\s]+/).filter(Boolean); }
  function txt(v) { return v == null ? '' : v instanceof Date ? v.toISOString() : String(v).replace(/\s+/g, ''); }
  function hasKw(v, list) { var t = txt(v); return list.some(function (k) { return t.indexOf(k) >= 0; }); }
  function toDate(v) {
    if (v instanceof Date && !isNaN(v)) { if (v.getFullYear() < 1990) return ''; var t = new Date(v.getTime() - v.getTimezoneOffset() * 60000); return t.toISOString().slice(0, 10); }
    if (typeof v === 'number' && v > 20000 && v < 80000) { var d = new Date(Math.round((v - 25569) * 86400000)); return d.toISOString().slice(0, 10); }
    var s = String(v || '').replace(/\s+/g, ''), m;
    if ((m = /(\d{2,4})[年\/.-](\d{1,2})[月\/.-](\d{1,2})/.exec(s))) { var y = +m[1]; if (y < 1911) y += 1911; return y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2); }
    return '';
  }
  function num(v) { if (typeof v === 'number') return v; var s = String(v == null ? '' : v).replace(/[,\s]/g, ''); return s === '' || isNaN(+s) ? null : +s; }
  /* 進度：儲存格若以百分比格式顯示就照顯示值；標題有「(%)」者數值已是百分比；否則 0～1 的小數視為比例 */
  function pctCell(row, ci, label) {
    for (var j = ci + 1; j < row.length; j++) {
      var v = num(row[j]); if (v == null) continue;
      var s = String((row.shown || [])[j] || '');
      if (/%$/.test(s.trim())) return parseFloat(s);
      return /[%％]/.test(label) || v > 1 ? v : v * 100;
    }
    return '';
  }
  function after(row, i, test) { for (var j = i + 1; j < row.length; j++) { if (row[j] !== '' && row[j] != null) { if (!test || test(row[j])) return row[j]; } } return ''; }
  function driBlock(rows, c, log, where) {
    var K = { date: kws(c.date), item: kws(c.item), qty: kws(c.qty), unit: kws(c.unit), weather: kws(c.weather), black: kws(c.black) };
    var out = { date: '', wAm: '', wPm: '', plan: '', act: '', items: [] }, hdr = -1, col = {};
    rows.forEach(function (row, ri) {
      row.forEach(function (v, ci) {
        var t = txt(v); if (!t) return;
        if (!out.date && /填表日期|報表日期|施工日期|^日期/.test(t)) { var dv = after(row, ci, function (x) { return !!toDate(x); }); if (dv) out.date = toDate(dv); }
        if (!out.wAm && /上午/.test(t) && row.slice(0, ci + 1).some(function (x) { return hasKw(x, K.weather); })) out.wAm = String(after(row, ci, function (x) { return !/下午/.test(txt(x)); }) || '').trim();
        if (!out.wPm && /下午/.test(t) && row.slice(0, ci + 1).some(function (x) { return hasKw(x, K.weather); })) out.wPm = String(after(row, ci, function (x) { return !/填表|日期/.test(txt(x)); }) || '').trim();
        /* 標題有「(%)」者數值已是百分比；否則 0～1 的小數視為比例 */
        if (out.plan === '' && /預定進度/.test(t)) out.plan = pctCell(row, ci, t);
        if (out.act === '' && /實際進度/.test(t)) out.act = pctCell(row, ci, t);
      });
      var shortCell = function (v) { var t = txt(v); return t && t.length <= 12; };
      if (hdr < 0 && row.some(function (v) { return shortCell(v) && hasKw(v, K.item); }) && row.some(function (v) { return shortCell(v) && hasKw(v, K.qty) && !hasKw(v, K.item); })) {
        hdr = ri;
        row.forEach(function (v, ci) {
          var t = txt(v);
          if (txt(v).length > 12) return;
          if (col.item == null && hasKw(v, K.item)) col.item = ci;
          if (col.unit == null && hasKw(v, K.unit)) col.unit = ci;
          if (col.contract == null && /契約數量/.test(t)) col.contract = ci;
          if (col.cum == null && /累計/.test(t)) col.cum = ci;
          if (col.qty == null && /本日|今日/.test(t) && hasKw(v, K.qty)) col.qty = ci;
        });
        if (col.qty == null) row.forEach(function (v, ci) { if (col.qty == null && hasKw(v, K.qty) && ci !== col.cum && ci !== col.contract) col.qty = ci; });
      }
    });
    if (out.date === '') { log.push('⚠ ' + where + '：找不到日期，略過'); return null; }
    if (hdr < 0) { log.push('⚠ ' + where + '（' + out.date + '）：找不到施工項目表頭'); return out; }
    for (var r = hdr + 1; r < rows.length; r++) {
      var row = rows[r], name = txt(row[col.item]) ? String(row[col.item]).trim() : '';
      var first = txt(row.find(function (x) { return txt(x); }) || '');
      if (/^[一二三四五六七八九十]、/.test(first) || /營造業專業工程|材料名稱|工別/.test(first)) break;
      if (!name || /^[A-ZＡ-Ｚ][.．]?$/.test(name)) continue;
      if (K.black.some(function (k) { return name.indexOf(k) >= 0; })) continue;
      var q = num(row[col.qty]);
      if (c.zero && !q) continue;
      out.items.push({ name: name.replace(/\s+$/, ''), unit: col.unit != null ? String(row[col.unit] || '').trim() : '', qty: q == null ? '' : String(q), cum: col.cum != null && num(row[col.cum]) != null ? String(num(row[col.cum])) : '', contract: col.contract != null && num(row[col.contract]) != null ? String(num(row[col.contract])) : '' });
    }
    return out;
  }
  function driParseBook(wb, name, c, log) {
    var days = [], X = window.XLSX;
    var sheets = wb.SheetNames.map(function (sn) {
      var ws = wb.Sheets[sn], rows = X.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: true }).map(function (r) { return r.slice(0, c.maxCol); });
      var shown = X.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: true });
      rows.forEach(function (r, i) { r.shown = shown[i] || []; });   // 儲存格顯示的文字（判斷是否為百分比格式）
      return { sn: sn, rows: rows };
    });
    var isDiary = function (rows) { return rows.slice(0, 15).some(function (r) { return r.some(function (v) { return /施工日誌/.test(txt(v)); }); }); };
    var perSheet = sheets.filter(function (s) { return isDiary(s.rows) && s.rows.length < 120; });
    var engine = c.engine === 'auto' ? (perSheet.length >= 2 ? 'sheet' : 'stack') : c.engine;
    log.push('■ ' + name + '：' + (engine === 'sheet' ? '標準單日頁籤' : '垂直堆疊') + '（' + wb.SheetNames.length + ' 張工作表）');
    if (engine === 'sheet') sheets.forEach(function (s) {
      if (!isDiary(s.rows)) { log.push('　略過工作表「' + s.sn + '」（不是施工日誌）'); return; }
      var d = driBlock(s.rows, c, log, '「' + s.sn + '」'); if (d) days.push(d);
    });
    else sheets.forEach(function (s) {
      var starts = []; s.rows.forEach(function (r, i) { var f = txt(r.find(function (x) { return txt(x); }) || ''); if (/施工日誌$/.test(f) && !/監造/.test(f)) starts.push(i); });
      if (!starts.length) return;
      starts.forEach(function (st, k) { var d = driBlock(s.rows.slice(st, starts[k + 1] || s.rows.length), c, log, '「' + s.sn + '」第 ' + (st + 1) + ' 列'); if (d && d.date) days.push(d); });
    });
    return days;
  }
  $('driGo').addEventListener('click', function () {
    var c = driSaveCfg(), log = [], all = [];
    $('driGo').disabled = true; $('driGo').textContent = '解析中…';
    (window.XLSX ? Promise.resolve() : loadScript(CDN.xlsx)).then(function () {
      return driFiles.reduce(function (pr, f) {
        return pr.then(function () { return f.arrayBuffer(); }).then(function (buf) {
          var wb = window.XLSX.read(buf, { type: 'array', cellDates: true });
          all = all.concat(driParseBook(wb, f.name, c, log));
        });
      }, Promise.resolve());
    }).then(function () {
      var by = {}; all.forEach(function (d) { by[d.date] = d; });
      var days = Object.keys(by).sort().map(function (k) { return by[k]; });
      log.push('完成：共 ' + days.length + ' 天、' + days.reduce(function (s, d) { return s + d.items.length; }, 0) + ' 個施工項目。');
      driResult = { days: days, log: log, sel: {} }; days.forEach(function (d) { driResult.sel[d.date] = true; });
    }).catch(function (e) { log.push('✕ 解析失敗：' + (e && e.message || e)); driResult = { days: [], log: log, sel: {} }; })
      .then(function () { $('driGo').disabled = false; $('driGo').textContent = '開始解析'; driPaint(); });
  });
  function driPaint() {
    var R = driResult, p = proj(cur.pid);
    $('driPv').innerHTML = !R ? '<p class="muted">尚未解析任何資料。</p>' : !R.days.length ? '<p class="muted">沒有解析出任何日期。請檢查解析紀錄，或調整關鍵字後再試。</p>' :
      '<div class="dri-tbl"><table><thead><tr><th><input type="checkbox" id="driAll" checked aria-label="全選"></th><th>日期</th><th>天氣</th><th>預定／實際</th><th>施工項目（本日／累計）</th><th>狀態</th></tr></thead><tbody>' +
      R.days.map(function (d) {
        var has = !!dailyOf(p, d.date);
        return '<tr><td><input type="checkbox" data-dsel="' + d.date + '"' + (R.sel[d.date] ? ' checked' : '') + '></td><td class="nw">' + d.date.slice(5).replace('-', '/') + '</td><td class="nw">' + esc([d.wAm, d.wPm].join('/')) + '</td><td class="nw">' + pct(d.plan) + '／' + pct(d.act) + '</td>' +
          '<td>' + (d.items.length ? d.items.map(function (it) { return esc(it.name) + ' <span class="muted">' + esc(it.qty + it.unit) + (it.cum ? '／' + esc(it.cum + it.unit) : '') + '</span>'; }).join('<br>') : '<span class="muted">無</span>') + '</td>' +
          '<td class="nw">' + (has ? ($('driDup').value === 'over' ? '將覆蓋' : '已有，略過') : '新增') + '</td></tr>';
      }).join('') + '</tbody></table></div>';
    $('driWrite').disabled = !(R && R.days.some(function (d) { return R.sel[d.date]; }));
    $('driLog').textContent = R ? R.log.join('\n') : '';
  }
  $('driPv').addEventListener('change', function (e) {
    if (!driResult) return;
    if (e.target.id === 'driAll') driResult.days.forEach(function (d) { driResult.sel[d.date] = e.target.checked; });
    else if (e.target.dataset.dsel) driResult.sel[e.target.dataset.dsel] = e.target.checked;
    driPaint();
  });
  $('driWrite').addEventListener('click', function () {
    var p = proj(cur.pid), over = $('driDup').value === 'over', n = 0, skip = 0;
    driResult.days.forEach(function (d) {
      if (!driResult.sel[d.date]) return;
      var ex = dailyOf(p, d.date); if (ex && !over) { skip++; return; }
      var t = dailyMake(p, d.date);
      t.wAm = d.wAm; t.wPm = d.wPm; t.plan = d.plan === '' ? '' : String(Math.round(d.plan * 10000) / 10000); t.act = d.act === '' ? '' : String(Math.round(d.act * 10000) / 10000);
      t.items = d.items.map(function (it) { return { name: it.name, qty: it.qty, unit: it.unit, cum: it.cum }; }); t.src = '施工日誌匯入'; n++;
    });
    save(); driResult.log.push('已寫入 ' + n + ' 天' + (skip ? '，略過 ' + skip + ' 天（已有日報）' : '') + '。'); driPaint();
    $('driMsg').textContent = '已寫入 ' + n + ' 天的日報。';
  });

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
  $('sumKind').addEventListener('change', function () { cur.kind = $('sumKind').value; go(cur, true); });
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
          defs.push({ date: r.checkDate, no: row.no, name: row.name, loc: row.loc, txt: items.join('；'), st: st, fd: st === 'fixed' ? fd : '', w: w, r: r, single: single, fixed: st === 'fixed' ? '已改善' + (fd ? ' ' + roc(fd, '/') : '') : (!single && r.fix === 'track' ? '追蹤改善中' : '未改善') });
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
  /* 四、材料／設備送審情形：依「截至」日判斷已送審、審查中、退回修正、尚未送審 */
  function sumSubmit(p, to) {
    var ms = p.mats || [], g = { ok: [], rev: [], back: [], todo: [] }, late = 0;
    ms.forEach(function (m) {
      var sent = m.actualSubmit && m.actualSubmit <= to, rd = !m.reviewDate || m.reviewDate <= to;
      if (sent && m.reviewResult === '核定' && rd) g.ok.push(m);
      else if (sent && /修正|不核定/.test(m.reviewResult || '') && rd) g.back.push(m);
      else if (sent) g.rev.push(m);
      else { g.todo.push(m); if (m.planSubmit && m.planSubmit < to) late++; }
    });
    var n = ms.length, sentN = n - g.todo.length, pct = function (a) { return n ? Math.round(a / n * 100) + '%' : '—'; };
    var d = function (v) { return v ? roc(v, '/') : ''; };
    var by = function (a, b) { var x = a.planSubmit || '9', y = b.planSubmit || '9'; return x < y ? -1 : x > y ? 1 : 0; };
    var sub = function (title, list, last) {
      if (!list.length) return '<p class="sum-sub">' + title + '：無</p>';
      return '<p class="sum-sub">' + title + '（' + list.length + ' 項）</p><table class="sum"><colgroup><col style="width:7%"><col style="width:33%"><col style="width:15%"><col style="width:15%"><col style="width:15%"><col style="width:15%"></colgroup>' +
        '<thead><tr><th>編號</th><th>各項材料／設備送審</th><th>預定施工日期</th><th>預定送審日期</th><th>實際送審日期</th><th>' + last[0] + '</th></tr></thead><tbody>' +
        list.slice().sort(by).map(function (m, i) {
          var overdue = !m.actualSubmit || m.actualSubmit > to ? (m.planSubmit && m.planSubmit < to) : false;
          return '<tr><td class="c">' + (i + 1) + '</td><td>' + esc([m.no, m.name].filter(Boolean).join(' ')) + '</td><td class="c">' + d(m.planUse) + '</td>' +
            '<td class="c">' + (overdue ? '<b class="sum-ng">' + d(m.planSubmit) + '</b>' : d(m.planSubmit)) + '</td><td class="c">' + (m.actualSubmit && m.actualSubmit <= to ? d(m.actualSubmit) : '') + '</td><td class="c">' + last[1](m) + '</td></tr>';
        }).join('') + '</tbody></table>';
    };
    return '<h4>四、各項材料／設備送審情形：' + (n ? '' : '尚未建立材料清單') + '</h4>' + (n ?
      '<table class="sum sum-stat"><thead><tr><th>總數量</th><th>已提送</th><th>逾期未送</th><th>已核定</th><th>審查中</th><th>退回修正</th><th>送審比例</th><th>核可比例</th></tr></thead><tbody><tr>' +
      [n, sentN, late ? '<b class="sum-ng">' + late + '</b>' : 0, g.ok.length, g.rev.length, g.back.length, pct(sentN), pct(g.ok.length)].map(function (v) { return '<td class="c">' + v + '</td>'; }).join('') + '</tr></tbody></table>' +
      sub('（一）尚未送審', g.todo, ['備註', function (m) { return m.planSubmit && m.planSubmit < to ? '<b class="sum-ng">逾期未送</b>' : ''; }]) +
      sub('（二）已送審、尚未核定', g.rev.concat(g.back), ['審查結果', function (m) { return g.back.indexOf(m) >= 0 ? '<b class="sum-ng">' + esc(m.reviewResult) + '</b>' + (m.reviewDate ? '<br>' + d(m.reviewDate) : '') : '審查中'; }]) +
      sub('（三）已核定', g.ok, ['核定日期', function (m) { return d(m.reviewDate); }]) : '');
  }
  /* 缺失：處理方式、複查結果（不定期／勞安有改善期限與處理方式；施工抽查只有複查結果） */
  function defInfo(x, to) {
    var r = x.r || {}, M = { repair: '修補改善', redo: '拆除重作' }, E = { improve: '改善', redo: '拆除重作' };
    var method = x.single ? M[r.method] || E[r.exec] || '' : '';
    var dl = x.single ? r.deadline || '' : '';
    var re = x.st === 'fixed' ? '合格' : x.single && r.re === 'ng' ? '不合格，持續追蹤' : !x.single && r.fix === 'track' ? '追蹤改善中' : '待複查';
    return { method: method, dl: dl, re: re, late: x.st !== 'fixed' && dl && dl < to };
  }
  function renderNgPrint(p, to, d) {
    var I = p.info, list = d.defs, fx = list.filter(function (x) { return x.st === 'fixed'; }).length, late = 0;
    var rows = list.map(function (x, i) {
      var f = defInfo(x, to); if (f.late) late++;
      return '<tr><td class="c">' + (i + 1) + '</td><td class="c">' + esc(x.no) + '</td><td class="c">' + roc(x.date, '/') + '</td><td>' + esc(x.name) + '</td><td class="c">' + esc(x.loc) + '</td><td>' + esc(x.txt) + '</td>' +
        '<td class="c">' + esc(f.method) + '</td><td class="c">' + (f.late ? '<b class="sum-ng">' + roc(f.dl, '/') + '<br>逾期</b>' : roc(f.dl, '/')) + '</td><td class="c">' + (x.fd ? roc(x.fd, '/') : '') + '</td>' +
        '<td class="c">' + (x.st === 'fixed' ? esc(f.re) : '<b class="sum-ng">' + esc(f.re) + '</b>') + '</td></tr>';
    }).join('');
    var paper = $('paper'); paper.className = 'paper landscape sum-paper';
    paper.innerHTML = '<div class="sum-title"><h3>施工不合格管制總表</h3></div><p class="no">工程名稱：' + esc(I.name || '') + (I.supervisor ? '　　監造單位：' + esc(I.supervisor) : '') + '</p>' +
      '<p class="sum-asof">截至 ' + roc(to) + '</p>' +
      '<p class="sum-cnt" style="margin-left:0">缺失總數：<b>' + list.length + '</b> 項　　已改善：<b>' + fx + '</b> 項　　未改善：<b>' + (list.length - fx) + '</b> 項　　逾期未改善：<b' + (late ? ' class="sum-ng"' : '') + '>' + late + '</b> 項</p>' +
      (list.length ? '<table class="sum"><colgroup><col style="width:4%"><col style="width:10%"><col style="width:8%"><col style="width:14%"><col style="width:8%"><col style="width:22%"><col style="width:8%"><col style="width:9%"><col style="width:8%"><col style="width:9%"></colgroup>' +
        '<thead><tr><th>項次</th><th>表單編號</th><th>發現日期</th><th>工項名稱</th><th>位置</th><th>缺失內容</th><th>處理方式</th><th>改善期限</th><th>改善完成日</th><th>複查結果</th></tr></thead><tbody>' + rows + '</tbody></table>'
        : '<p class="sum-none">截至本日尚無缺失紀錄。</p>') + foot();
    $('printHint').textContent = '紙張選 A4 橫向、邊界「預設」。「另存 PDF」會直接下載檔案，不經過列印視窗。';
    setDocName('施工不合格管制總表_' + (I.name || '') + '_截至' + roc(to, '.'));
  }
  function renderTrackPrint(p, to, d) {
    var I = p.info, box2 = function (on) { return on ? '■' : '□'; };
    var img = function (x) { return '<td class="ph-img">' + (x ? '<img alt="" data-img="' + x.id + '">' : '') + '</td>'; };
    var cap = function (x) { return '<td class="ph-cap">' + (x ? esc(x.cap) : '') + '</td>'; };
    var pages = d.defs.map(function (x, i) {
      var f = defInfo(x, to), r = x.r || {}, phs = r.photos || [];
      var aft = phs.filter(function (q) { return /改善後|改善完成|完成後|複查/.test(q.cap || ''); }), bef = phs.filter(function (q) { return aft.indexOf(q) < 0; });
      var fixer = x.single ? '' : esc(r.fixPerson || I.inspector || '') + ' ' + (r.sig && r.sig.sigFixer ? sigImg(r, 'sigFixer') : '');
      return '<section' + (i ? ' class="rec-sec"' : '') + '><h3>施工品質缺失處理改善暨追蹤紀錄表</h3><p class="no">編號：' + esc(x.no) + '-NG</p>' +
        '<table class="trk"><colgroup><col style="width:16%"><col style="width:34%"><col style="width:16%"><col style="width:34%"></colgroup>' +
        '<tr><th>工程名稱</th><td colspan="3">' + esc(I.name || '') + '</td></tr>' +
        '<tr><th>工項名稱</th><td>' + esc(x.name) + '</td><th>位置</th><td>' + esc(x.loc) + '</td></tr>' +
        '<tr><th>發現日期</th><td>' + roc(x.date) + '</td><th>改善期限</th><td>' + (f.dl ? roc(f.dl) : '') + '</td></tr>' +
        '<tr><th>缺失內容</th><td colspan="3" class="trk-txt">' + esc(x.txt) + '</td></tr>' +
        '<tr><th>處理方式</th><td colspan="3">' + box2(f.method === '修補改善' || f.method === '改善') + '修補改善　' + box2(f.method === '拆除重作') + '拆除重作</td></tr></table>' +
        '<table class="ph"><colgroup><col style="width:50%"><col style="width:50%"></colgroup><tr><th>改善前</th><th>改善後</th></tr>' +
        '<tr>' + img(bef[0]) + img(aft[0]) + '</tr><tr>' + cap(bef[0]) + cap(aft[0]) + '</tr></table>' +
        '<table class="trk"><colgroup><col style="width:16%"><col style="width:84%"></colgroup>' +
        '<tr><th>複查結果</th><td>' + box2(x.st === 'fixed') + '已完成改善　' + box2(x.st !== 'fixed') + '未完成改善，持續追蹤</td></tr>' +
        '<tr><th>複查日期</th><td>' + (x.fd ? roc(x.fd) : '　　年　　月　　日') + '</td></tr>' +
        '<tr><th>複查人員</th><td class="trk-sig">' + fixer + '</td></tr></table>' + (i === d.defs.length - 1 ? foot() : '') + '</section>';
    }).join('');
    var paper = $('paper'); paper.className = 'paper sum-paper';
    paper.innerHTML = pages || '<p class="pv-empty">截至本日尚無缺失紀錄。</p>';
    fillImgs(paper);
    $('printHint').textContent = '每項缺失一頁。照片說明含「改善後」的會放在右側「改善後」欄。紙張選 A4 直式。';
    setDocName('缺失改善追蹤紀錄表_' + (I.name || '') + '_截至' + roc(to, '.'));
  }
  function renderSumPrint() {
    var p = proj(cur.pid), I = p.info, to = cur.to || today(), from = cur.from || sumFromDef(to), d = sumData(p, to), kind = cur.kind || 'sum';
    topbar('列印預覽', { sum: '抽查總表', ng: '施工不合格管制總表', track: '缺失改善追蹤紀錄表' }[kind], I.name);
    $('sumFrom').value = from; $('sumTo').value = to; $('sumBar').hidden = false; $('sumKind').value = kind;
    $('sumFrom').closest('label').hidden = kind !== 'sum';
    if (kind !== 'sum') {
      bindOut(sumOut, 't'); $('pvParts').hidden = true; $('btnMovePv').hidden = true; $('outAt').closest('.out-at').hidden = false;
      show('vPreview');
      if (kind === 'ng') renderNgPrint(p, to, d); else renderTrackPrint(p, to, d);
      return;
    }
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
      '<p class="sum-asof">截至 ' + roc(to) + '</p>' +
      '<h4>一、材料進場抽驗情形：</h4>' + head(cm) + '</p>' + tbl(d.mats, COLS, basic) +
      '<h4>二、分項施工抽查情形：</h4>' + head(cc) + '</p>' + tbl(d.cons, COLS, basic) +
      '<h4>三、缺失改善追蹤情形：' + (d.defs.length ? '' : '無') + '</h4>' + head(cd) + '　　已改善：<b>' + fixedN + '</b> 次　未改善：<b>' + (d.defs.length - fixedN) + '</b> 次</p>' +
      (d.defs.length ? tbl(d.defs, [['項次', '7%'], ['項目名稱', '22%'], ['位置', '13%'], ['抽驗日期', '12%'], ['缺失內容', '28%'], ['改善情形', '18%']], function (x) {
        return '<td>' + esc(x.name) + '</td><td class="c">' + esc(x.loc) + '</td><td class="c">' + roc(x.date, '/') + '</td><td>' + esc(x.txt) + '</td><td class="c">' + (x.st === 'fixed' ? esc(x.fixed) : '<b class="sum-ng">' + esc(x.fixed) + '</b>') + '</td>';
      }) : '') +
      sumSubmit(p, to) +
      foot();
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
    (p.docs || []).forEach(function (d) {
      if (d.date) ev.push({ date: d.date, kind: 'doc', label: (d.kind === 'meet' ? '開會通知：' : '發文：') + shortSubj(d), sub: d.no, did: d.id });
      if (d.kind === 'meet' && d.meet && d.meet.time) ev.push({ date: d.meet.time.slice(0, 10), kind: 'meet', label: '會議：' + shortSubj(d), sub: [d.meet.time.slice(11, 16), d.meet.place].filter(Boolean).join(' · '), did: d.id });
    });
    (p.dailies || []).forEach(function (d) { ev.push({ date: d.date, kind: 'daily', label: '監造日報' + ((d.items || []).length ? '（' + d.items.length + ' 項施工）' : ''), sub: [d.wAm, d.wPm].filter(Boolean).join('/'), drd: d.date }); });
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
    var CF = { all: null, rec: ['rec'], plan: ['plan'], mat: ['mat'], doc: ['doc', 'meet'], daily: ['daily'] };
    var ev = calEvents(p).filter(function (e) { return !CF[calFilter] || CF[calFilter].indexOf(e.kind) >= 0; }), byDay = {};
    $('calFilter').innerHTML = [['all', '全部'], ['rec', '抽查'], ['plan', '預定'], ['mat', '材料'], ['doc', '公文'], ['daily', '日報']].map(function (f) { return '<button type="button" class="cf' + (calFilter === f[0] ? ' on' : '') + '" data-cf="' + f[0] + '">' + f[1] + '</button>'; }).join('');
    ev.forEach(function (e) { (byDay[e.date] = byDay[e.date] || []).push(e); });
    var first = new Date(calY, calM, 1), start = new Date(calY, calM, 1 - first.getDay()), td = today(), cells = '';
    ['日', '一', '二', '三', '四', '五', '六'].forEach(function (d) { cells += '<div class="cal-wd">' + d + '</div>'; });
    for (var i = 0; i < 42; i++) {
      var d = new Date(start); d.setDate(start.getDate() + i);
      var key = ymd(d), list = byDay[key] || [], other = d.getMonth() !== calM;
      if (i === 35 && other) break;
      cells += '<button type="button" class="cal-cell' + (other ? ' other' : '') + (key === td ? ' today' : '') + (key === calSel ? ' sel' : '') + '" data-day="' + key + '" aria-label="' + key + ' ' + list.length + ' 筆">' +
        '<span class="cal-n">' + d.getDate() + '</span><span class="cal-evs">' + calBadges(list) + '</span></button>';
    }
    $('calGrid').innerHTML = cells;
    renderCalDay(p, byDay[calSel] || []);
  }
  /* 日曆格子只放各類數量徽章（抽／預／材／文／會），內容在點選當天後的清單 */
  var calFilter = 'all';
  var BADGE = [['daily', '報'], ['rec', '抽'], ['plan', '預'], ['mat', '材'], ['doc', '文'], ['meet', '會']];
  function calBadges(list) {
    return BADGE.map(function (b) {
      var xs = list.filter(function (e) { return e.kind === b[0]; }); if (!xs.length) return '';
      var ng = xs.some(function (e) { return e.ng; }), over = xs.some(function (e) { return e.overdue; });
      return '<span class="cb cb-' + b[0] + (ng ? ' cb-ng' : '') + (over ? ' cb-over' : '') + '" title="' + esc(xs.map(function (e) { return e.label; }).join('\n')) + '">' + b[1] + (xs.length > 1 ? xs.length : '') + (ng ? '!' : '') + '</span>';
    }).join('');
  }
  $('calFilter').addEventListener('click', function (e) { var b = e.target.closest('[data-cf]'); if (!b) return; calFilter = b.dataset.cf; renderCal(proj(cur.pid)); });
  function evChip(e, small) {
    var cls = 'ev ev-' + e.kind + (e.ng ? ' ev-ng' : '') + (e.overdue ? ' ev-over' : '');
    var attr = e.plid ? ' data-plid="' + e.plid + '" draggable="true"' : '';
    return '<span class="' + cls + '"' + attr + ' title="' + esc(e.label) + '">' + (e.ng ? '╳ ' : e.overdue ? '! ' : '') + esc(e.label) + '</span>';
  }
  var calOpen = null;   // 展開動作列的預定 id
  function renderCalDay(p, list) {
    var on = enabledWorks(p);
    var d = calSel.split('-');
    var tag = { plan: '預定', rec: '已查', mat: '進場', doc: '發文', meet: '會議', daily: '日報' };
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
      var go2 = e.rid ? ' data-openrec="' + e.wid + '|' + e.rid + '"' : e.did ? ' data-opendoc="' + e.did + '"' : e.drd ? ' data-opendr="' + e.drd + '"' : ' data-openmat="' + e.mid + '"';
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
    if (ds.start) { startPlan(p, ds.start); return; }
    if (ds.openrec) { var a = ds.openrec.split('|'); go({ v: 'form', pid: p.id, wid: a[0], rid: a[1] }); return; }
    if (ds.openmat) go({ v: 'matItem', pid: p.id, mid: ds.openmat });
    if (ds.opendoc) go({ v: 'docEdit', pid: p.id, did: ds.opendoc });
    if (ds.opendr) go({ v: 'dailyEdit', pid: p.id, date: ds.opendr });
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
    rerender: function () { render(); }, ask: ask, coEdit: coEdit,
    fixRec: function (wid, r) { return migrateRec(WORK[wid], r); }, migrate: migrateAll
  };
  /* 啟動：讀取 IndexedDB 中較新的資料後再開始 */
  function boot() {
    try { gwFontCheck(); } catch (e) {}
    S.projects.forEach(function (p) { if (!p.mats) p.mats = []; });
    migrateBoot();
    go({ v: uiMode() === 'field' ? 'field' : 'projects' }, true);
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
