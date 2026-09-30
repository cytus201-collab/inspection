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
    xlsx: 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
  };
  var $ = function (id) { return document.getElementById(id); };
  var WORK = {}; D.works.forEach(function (w) { WORK[w.id] = w; });
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
  function save() { try { localStorage.setItem(KEY, JSON.stringify(S)); saveMsg = ''; } catch (e) { saveMsg = '裝置空間不足或瀏覽器禁止暫存，請匯出備份'; } }
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
    var t = 0, d = 0, ng = 0;
    if (w.kind === 'irregular') {
      (r.rows || []).forEach(function (x) { t++; if (x.r) d++; if (x.r === 'ng') ng++; });
      return { t: t, d: d, ng: ng };
    }
    w.phases[r.phase].groups.forEach(function (g, gi) {
      g.items.forEach(function (_, ii) { var c = (r.res[k(gi, ii)] || {}).r; t++; if (c) d++; if (c === 'ng') ng++; });
    });
    return { t: t, d: d, ng: ng };
  }
  function recTitle(w, r) { return w.kind === 'phased' ? w.name + '(' + (r.phase + 1) + ') ' + flowName(r.phase) : w.name; }
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
  var VIEWS = ['vProjects', 'vProject', 'vWork', 'vForm', 'vMaterials', 'vMatItem', 'vMatImport', 'vPreview'];
  function show(id) { VIEWS.forEach(function (v) { $(v).hidden = v !== id; }); window.scrollTo(0, 0); }
  function topbar(title, sub, back, right) {
    $('tbTitle').textContent = title; $('tbSub').textContent = sub || '';
    $('navBack').hidden = !back; $('navBack').textContent = back ? '‹ ' + back : '';
    $('tbRight').textContent = right || '';
    document.title = title;
  }
  function render() {
    var p = cur.pid && proj(cur.pid), w = cur.wid && WORK[cur.wid];
    if (cur.v !== 'projects' && !p) cur = { v: 'projects' };
    else if (/^(work|form|preview)$/.test(cur.v) && !w) cur = { v: 'project', pid: cur.pid };
    else if (/^(form|preview)$/.test(cur.v) && !recById(p, cur.wid, cur.rid)) cur = { v: 'work', pid: cur.pid, wid: cur.wid };
    else if (cur.v === 'matItem' && !matById(p, cur.mid)) cur = { v: 'materials', pid: cur.pid };
    else if (cur.v === 'matImport' && !pendingImport) cur = { v: 'materials', pid: cur.pid };
    ({
      projects: renderProjects, project: renderProject, work: renderWork, form: renderForm, preview: renderPreview,
      materials: renderMaterials, matItem: renderMatItem, matImport: renderMatImport, matPrint: renderMatPrint
    })[cur.v]();
  }

  /* ========== 1 工程清單 ========== */
  function renderProjects() {
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
    show('vProjects');
  }
  $('projectList').addEventListener('click', function (e) {
    var b = e.target.closest('[data-pid]'); if (b) go({ v: 'project', pid: b.dataset.pid });
  });
  $('btnNewProject').addEventListener('click', function () {
    var p = { id: uid(), info: { durType: '日曆天', changeCount: '0', extendDays: '0' }, recs: {}, mats: [], created: Date.now() };
    S.projects.unshift(p); save(); go({ v: 'project', pid: p.id, focus: 1 });
  });
  $('btnExport').addEventListener('click', function () {
    download('監造抽查備份_' + today() + '.json', JSON.stringify(S), 'application/json');
    $('backupMsg').textContent = '已匯出 ' + S.projects.length + ' 個工程。';
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
  function renderProject() {
    var p = proj(cur.pid), I = p.info;
    topbar(I.name || '未命名工程', [I.code, I.contractor].filter(Boolean).join(' · '), '工程清單');
    $('workGrid').innerHTML = D.works.map(function (w) {
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

    $('infoForm').innerHTML = INFO_FIELDS.map(function (f) { return fieldHtml('i_', f, I[f.k]); }).join('');
    $('infoCard').open = !I.name || !!cur.focus;
    $('infoSaved').textContent = saveMsg || '已自動儲存';
    show('vProject');
    if (cur.focus) setTimeout(function () { $('i_name').focus(); }, 50);
  }
  function barHtml(c, n) {
    if (!n) return '<span class="progress-bar"></span>';
    return '<span class="progress-bar">' + MSTAT_ORDER.map(function (s) {
      return c[s] ? '<span class="seg-' + s + '" style="width:' + (c[s] / n * 100).toFixed(2) + '%" title="' + MSTAT[s] + ' ' + c[s] + '"></span>' : '';
    }).join('') + '</span>';
  }
  $('workGrid').addEventListener('click', function (e) {
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
    ask('刪除「' + (p.info.name || '未命名工程') + '」？', '此工程底下 ' + n + ' 份抽查紀錄與 ' + p.mats.length + ' 項材料管制資料都會刪除，無法復原。', '取消', '刪除工程', function () {
      S.projects.splice(S.projects.indexOf(p), 1); save(); go({ v: 'projects' }, true);
    });
  });

  /* ========== 3 工項紀錄清單 ========== */
  function renderWork() {
    var p = proj(cur.pid), w = WORK[cur.wid], list = recs(p, w.id);
    topbar(w.name, p.info.name || '未命名工程', p.info.name || '工程');
    var blocks = w.kind === 'phased' ? w.phases.map(function (ph, pi) { return { pi: pi, title: '(' + (pi + 1) + ') ' + flowName(pi), n: countItems(ph) + ' 項' }; })
      : [{ pi: 0, title: '抽查紀錄', n: w.kind === 'safety' ? countItems(w.phases[0]) + ' 項' : '項目自選' }];
    $('phaseBlocks').innerHTML = blocks.map(function (b) {
      var items = list.filter(function (r) { return r.phase === b.pi; }).sort(function (a, c) { return c.seq - a.seq; });
      var rows = items.map(function (r) {
        var s = stats(w, r);
        var meta = [r.checkDate && roc(r.checkDate, '.'), r.location].filter(Boolean).join(' · ') || '尚未填寫位置與日期';
        return '<button type="button" class="rec" data-rid="' + r.id + '"><span class="rb-main"><span class="rb-title">' + esc(r.docNo) +
          (s.ng ? '<span class="badge-ng">' + s.ng + ' 缺失</span>' : '') + '</span><span class="rb-meta">' + esc(meta) + '</span></span>' +
          '<span class="rb-right">' + s.d + '/' + s.t + '</span><span class="chev">›</span></button>';
      }).join('');
      return '<section class="phase-block"><div class="pb-head"><span class="pb-title">' + b.title +
        '<span class="pb-n"> · ' + b.n + '</span></span><button type="button" class="btn sm" data-new="' + b.pi + '">＋ 新增抽查</button></div>' +
        (rows || '<div class="pb-empty">尚無紀錄</div>') + '</section>';
    }).join('');
    $('workRefs').textContent = w.kind === 'irregular' ? '不定期抽查依承商自主檢查紀錄表隨機抽樣，抽驗項目可從 11 個工項的抽查標準帶入。'
      : '抽查標準依據：' + w.sources.map(function (s) { return D.refs[s]; }).join('；') + '。數值請依本案契約圖說確認。';
    show('vWork');
  }
  function countItems(ph) { return ph.groups.reduce(function (a, g) { return a + g.items.length; }, 0); }
  $('phaseBlocks').addEventListener('click', function (e) {
    var p = proj(cur.pid), w = WORK[cur.wid];
    var nb = e.target.closest('[data-new]');
    if (nb) {
      var pi = +nb.dataset.new, list = recs(p, w.id);
      var seq = list.filter(function (r) { return r.phase === pi; }).reduce(function (m, r) { return Math.max(m, r.seq); }, 0) + 1;
      var r = {
        id: uid(), phase: pi, seq: seq,
        docNo: (p.info.code ? p.info.code + '-' : '') + w.code + (w.kind === 'phased' ? (pi + 1) : '') + '-' + ('00' + seq).slice(-3),
        location: '', checkDate: today(), res: {}, fix: 'none', fixDate: '', fixPerson: '', sig: {}, created: Date.now()
      };
      if (w.kind !== 'phased') { r.rows = []; r.result = ''; r.exec = ''; r.deadline = ''; r.method = ''; r.methodNote = ''; r.re = ''; r.reTime = ''; r.remark = ''; }
      list.push(r); save(); go({ v: 'form', pid: p.id, wid: w.id, rid: r.id });
      return;
    }
    var rb = e.target.closest('[data-rid]');
    if (rb) go({ v: 'form', pid: p.id, wid: w.id, rid: rb.dataset.rid });
  });

  /* ========== 4 填寫 ========== */
  function R() { return recById(proj(cur.pid), cur.wid, cur.rid); }
  function W() { return WORK[cur.wid]; }

  function renderForm() {
    var p = proj(cur.pid), w = W(), r = R(), I = p.info, single = w.kind !== 'phased';
    topbar(recTitle(w, r), I.name || '未命名工程', w.name);
    $('autoLine').innerHTML = '自動帶入：工程名稱「' + esc(I.name || '未填') + '」' + (single ? '' : ' · 分項「' + esc(w.name) + '」') + ' · 監造人員「' + esc(I.inspector || '未填') + '」';
    ['docNo', 'location', 'checkDate', 'fixDate', 'fixPerson'].forEach(function (id) { $(id).value = r[id] || ''; });
    $('fixPerson').placeholder = I.inspector || '';
    document.querySelectorAll('input[name="fix"]').forEach(function (x) { x.checked = x.value === r.fix; });
    $('sigInspectorLbl').textContent = (single ? '監造現場人員' : '監造人員') + (I.inspector ? '（' + I.inspector + '）' : '');
    $('phasedFix').hidden = single; $('singleResult').hidden = !single;
    $('irrPicker').hidden = w.kind !== 'irregular';
    if (single) fillSingle(r); else fixVis();
    if (w.kind === 'irregular') { fillPickWork(); renderRows(); } else renderGroups();
    show('vForm');
    sizePad('sigInspector'); sizePad('sigFixer');
  }
  ['docNo', 'location', 'checkDate', 'fixDate', 'fixPerson'].forEach(function (id) {
    $(id).addEventListener('input', function () { R()[id] = $(id).value; save(); });
  });

  /* 固定清單（分階段工項、勞安衛生） */
  function renderGroups() {
    var w = W(), r = R();
    $('groups').innerHTML = w.phases[r.phase].groups.map(function (g, gi) {
      var items = g.items.map(function (it, ii) {
        var key = k(gi, ii), c = cell(r, key);
        return itemHtml('data-key="' + key + '"', 'r' + key, esc(it.name), esc(it.std), D.refs[it.ref], c.r, c.note, it.hint, 'n' + key);
      }).join('');
      return '<details class="group" open><summary>' + esc(g.name) + '<span class="gcount" id="gc' + gi + '"></span></summary>' +
        '<div class="gtools"><button type="button" class="link" data-allok="' + gi + '">本組未判定的全部設為 ○</button></div>' + items + '</details>';
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
      '<div class="note"><textarea id="' + noteId + '" rows="1" placeholder="' + esc(hint || '實際抽查情形') + '" aria-label="實際抽查情形">' + esc(note) + '</textarea></div></div>' +
      (extra || '') + '</div>';
  }
  var plainName = function (s) { return '<div class="item-name">' + s + '</div>'; };
  var plainStd = function (s) { return s ? '<div class="item-std">' + s + '</div>' : ''; };
  // 包裝成相同外觀
  var _itemHtml = itemHtml;
  itemHtml = function (attr, rn, name, std, ref, val, note, hint, nid, extra) { return _itemHtml(attr, rn, plainName(name), plainStd(std), ref, val, note, hint, nid, extra); };

  $('groups').addEventListener('change', function (e) {
    var item = e.target.closest('.item'); if (!item || e.target.type !== 'radio') return;
    var r = R(), v = e.target.value;
    if (item.dataset.row != null) r.rows[+item.dataset.row].r = v; else cell(r, item.dataset.key).r = v;
    item.classList.toggle('ng', v === 'ng'); save(); updateCounts();
  });
  $('groups').addEventListener('input', function (e) {
    var item = e.target.closest('.item'); if (!item) return;
    var r = R();
    if (item.dataset.row != null) {
      var row = r.rows[+item.dataset.row];
      if (e.target.tagName === 'TEXTAREA' && e.target.dataset.f) row[e.target.dataset.f] = e.target.value;
      else if (e.target.tagName === 'TEXTAREA') row.note = e.target.value;
      else if (e.target.dataset.f) row[e.target.dataset.f] = e.target.value;
    } else if (e.target.tagName === 'TEXTAREA') cell(r, item.dataset.key).note = e.target.value;
    save();
  });
  $('groups').addEventListener('click', function (e) {
    var t = e.target;
    if (t.dataset && t.dataset.allok != null) {
      var r = R(), gi = t.dataset.allok; W().phases[r.phase].groups[+gi].items.forEach(function (_, ii) { var c = cell(r, k(gi, ii)); if (!c.r) c.r = 'ok'; });
      save(); renderGroups(); return;
    }
    if (t.dataset && t.dataset.delrow != null) { R().rows.splice(+t.dataset.delrow, 1); save(); renderRows(); }
  });
  function updateCounts() {
    var w = W(), r = R(), s = stats(w, r);
    if (w.kind !== 'irregular') {
      w.phases[r.phase].groups.forEach(function (g, gi) {
        var d = 0, ng = 0;
        g.items.forEach(function (_, ii) { var c = cell(r, k(gi, ii)).r; if (c) d++; if (c === 'ng') ng++; });
        $('gc' + gi).textContent = (ng ? ng + ' 缺失 · ' : '') + d + '/' + g.items.length;
      });
    }
    $('tbRight').textContent = s.d + '/' + s.t;
    if (w.kind !== 'phased') {
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
    $('pickWork').innerHTML = PHASED.map(function (w) { return '<option value="' + w.id + '">' + esc(w.name) + '</option>'; }).join('');
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
    ask('刪除紀錄 ' + r.docNo + '？', '刪除後無法復原。', '取消', '刪除紀錄', function () {
      var list = recs(proj(cur.pid), cur.wid); list.splice(list.indexOf(r), 1); save();
      go({ v: 'work', pid: cur.pid, wid: cur.wid }, true);
    });
  });

  $('btnConfirm').addEventListener('click', function () {
    var w = W(), r = R(), miss = [];
    var toPreview = function () {
      if (w.kind !== 'phased' && !r.result) { r.result = stats(w, r).ng ? 'ng' : 'ok'; save(); }
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
    w.phases[r.phase].groups.forEach(function (g, gi) { g.items.forEach(function (it, ii) { if (!cell(r, k(gi, ii)).r) miss.push(it.name); }); });
    if (!miss.length) return toPreview();
    ask('還有 ' + miss.length + ' 項未判定', miss.slice(0, 5).join('、') + (miss.length > 5 ? ' 等' : '') + '。', '回去填寫', '未判定的標為「／」並繼續', function () {
      w.phases[r.phase].groups.forEach(function (g, gi) { g.items.forEach(function (_, ii) { var c = cell(r, k(gi, ii)); if (!c.r) c.r = 'na'; }); });
      save(); toPreview();
    });
  });

  /* ========== 預覽列印 ========== */
  function renderPreview() {
    var p = proj(cur.pid), w = W(), r = R();
    topbar('列印預覽', recTitle(w, r) + ' ' + r.docNo, '返回修改');
    bindOut(r, 'outAt');
    var paper = $('paper'); paper.className = 'paper';
    paper.innerHTML = w.kind === 'phased' ? paperPhased(p, w, r) : paperSingle(p, w, r);
    $('printHint').textContent = '紙張選 A4 直式、邊界「預設」；要存 PDF 就把印表機選「另存為 PDF」。';
    show('vPreview');
    document.title = (w.kind === 'phased' ? w.name + '(' + (r.phase + 1) + ')' : w.name) + '_' + r.docNo;
  }
  $('btnPrint').addEventListener('click', function () { window.print(); });
  function sigImg(r, id) { return r.sig[id] ? '<img alt="簽名" src="' + r.sig[id] + '">' : ''; }
  /* 表單產出時間：預設為現在，可在預覽工具列自訂，存在該筆紀錄（或工程的材料總表）上 */
  function nowLocal() { var d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 16); }
  function fmtOut(v) { var m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(v || ''); return m ? (+m[1]) + '/' + (+m[2]) + '/' + (+m[3]) + ' ' + m[4] + ':' + m[5] : ''; }
  var outTarget = null;   // { obj, key }
  function outValue() { return (outTarget && outTarget.obj[outTarget.key]) || nowLocal(); }
  function foot() { return '<div class="foot"><span>表單版次 ' + esc(D.version) + '</span><span class="foot-out">產出 ' + fmtOut(outValue()) + '</span></div>'; }
  function bindOut(obj, key) { outTarget = { obj: obj, key: key }; $('outAt').value = outValue(); }
  function setOut(v) {
    if (!outTarget) return;
    outTarget.obj[outTarget.key] = v; save();
    var f = document.querySelector('#paper .foot-out'); if (f) f.textContent = '產出 ' + fmtOut(outValue());
    $('outAt').value = outValue();
  }
  $('outAt').addEventListener('change', function () { if ($('outAt').value) setOut($('outAt').value); });
  $('btnOutNow').addEventListener('click', function () { setOut(nowLocal()); });
  function box(on) { return on ? '■' : '□'; }

  function paperPhased(p, w, r) {
    var I = p.info, ph = w.phases[r.phase];
    var flow = PH_SHORT.map(function (_, i) { return '<span class="opt">' + box(i === r.phase) + flowName(i) + '</span>'; }).join('');
    var body = '';
    ph.groups.forEach(function (g, gi) {
      body += '<tr class="sub"><td colspan="4">' + esc(g.name) + '</td></tr>';
      g.items.forEach(function (it, ii) {
        var c = cell(r, k(gi, ii));
        body += '<tr><td>' + esc(it.name) + '</td><td>' + esc(it.std) + '</td><td>' + esc(c.note) + '</td><td class="res' + (c.r === 'ng' ? ' ng' : '') + '">' + (SYM[c.r] || '') + '</td></tr>';
      });
    });
    var on = r.fix !== 'none';
    var fixTxt = box(r.fix === 'done') + '已完成改善（檢附改善前中後照片）<br>' + box(r.fix === 'track') + '未完成改善，填具「施工品質缺失處理改善暨追蹤紀錄表」進行追蹤改善';
    return '<h3>' + esc(w.name) + '施工抽查紀錄表(' + (r.phase + 1) + ')</h3><p class="no">編號：' + esc(r.docNo) + '</p>' +
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
      r.rows.forEach(function (x) {
        inner += '<tr><td>' + esc((x.work ? x.work + '／' : '') + x.name) + '</td><td>' + esc(x.std) + '</td><td>' + esc(x.note) + '</td><td class="res' + (x.r === 'ng' ? ' ng' : '') + '">' + (SYM[x.r] || '') + '</td></tr>';
      });
    } else {
      w.phases[0].groups.forEach(function (g, gi) {
        inner += '<tr class="sub"><td colspan="4">' + esc(g.name) + '</td></tr>';
        g.items.forEach(function (it, ii) {
          var c = cell(r, k(gi, ii));
          inner += '<tr><td>' + esc(it.name) + '</td><td>' + esc(it.std) + '</td><td>' + esc(c.note) + '</td><td class="res' + (c.r === 'ng' ? ' ng' : '') + '">' + (SYM[c.r] || '') + '</td></tr>';
        });
      });
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
  $('btnInspToday').addEventListener('click', function () { var m = M(); m.inspDate = today(); $('m_inspDate').value = m.inspDate; save(); });
  $('btnMatDone').addEventListener('click', function () { history.back(); });
  $('btnMatDel').addEventListener('click', function () {
    var p = proj(cur.pid), m = M();
    ask('刪除「' + (m.name || '未命名材料') + '」？', '此材料的送審、進場與抽查紀錄都會刪除。', '取消', '刪除材料', function () {
      p.mats.splice(p.mats.indexOf(m), 1); save(); go({ v: 'materials', pid: p.id }, true);
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
    bindOut(p, 'matOutAt');
    var paper = $('paper'); paper.className = 'paper landscape';
    paper.innerHTML = '<h3>材料設備送審管制總表</h3><p class="no">工程名稱：' + esc(I.name) + '</p>' +
      '<table class="mat"><colgroup><col style="width:3.5%"><col style="width:17%"><col style="width:8%"><col style="width:5%"><col style="width:8%"><col style="width:7%">' +
      '<col style="width:4.5%"><col style="width:4.5%"><col style="width:4.5%"><col style="width:4.5%"><col style="width:5%"><col style="width:8%"><col style="width:6.5%"><col style="width:7%"><col style="width:7.5%"></colgroup>' +
      '<thead><tr><th rowspan="2">項次</th><th rowspan="2">契約詳細表項次<br>材料/設備名稱</th><th rowspan="2">契約數量</th><th rowspan="2">是否取樣試驗</th><th rowspan="2">預定送審日期<br>實際送審日期</th><th rowspan="2">是否驗廠<br>驗廠日期</th>' +
      '<th colspan="5">送審資料（ˇ）</th><th rowspan="2">審查日期<br>審查結果</th><th rowspan="2">進場日期</th><th rowspan="2">抽查日期<br>結果</th><th rowspan="2">備註<br>(歸檔編號)</th></tr>' +
      '<tr><th>協力廠商資料</th><th>型錄</th><th>相關試驗報告</th><th>樣品色票</th><th>其他</th></tr></thead><tbody>' + rows + '</tbody></table>' +
      '<p class="mat-note">註：本表單於開工後應請廠商檢討提出預定送審及預定進場日期，並由監造單位會同廠商定期檢討辦理情形。</p>' + foot();
    $('printHint').textContent = '紙張選 A4 橫向、邊界「預設」；要存 PDF 就把印表機選「另存為 PDF」。';
    show('vPreview');
    document.title = '材料設備送審管制總表_' + (I.name || '');
  }

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

  go({ v: 'projects' }, true);

  if ('serviceWorker' in navigator && /^https?:/.test(location.protocol)) {
    navigator.serviceWorker.register('sw.js').catch(function () {});
  }
})();
