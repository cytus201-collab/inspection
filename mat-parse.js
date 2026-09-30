/* 材料設備送審管制總表解析（純函式，可在瀏覽器與 Node 測試）
   pages: [{ items: [{ str, x, y }] }]，y 由上往下遞增（單位 pt） */
(function (root) {
  'use strict';
  var ANCHOR = /^[壹貳參肆伍陸柒捌玖拾]+[.．、][^\s]*/;
  var DOC_KEYS = ['vendor', 'catalog', 'test', 'sample', 'other'];

  function clean(s) { return String(s || '').replace(/\(cid:\d+\)/g, '\u0001').replace(/[\t\r\n]+/g, ' ').trim(); }

  function findCols(items) {
    var c = {};
    items.forEach(function (it) {
      var s = it.s;
      if (s === '契約' || s === '數量') c.qty = c.qty == null ? it.x : Math.min(c.qty, it.x);
      else if (s === '取樣') c.sampleTest = it.x;
      else if (s.indexOf('驗廠') === 0) c.factory = it.x;
      else if (s === '協力') c.vendor = it.x;
      else if (s.indexOf('型錄') === 0) c.catalog = it.x;
      else if (s === '相關' || s === '報告') c.test = it.x;
      else if (s === '樣品' || s === '色票') c.sample = it.x;
      else if (s === '其他') c.other = it.x;
      else if (s === '審查結果') c.review = it.x;
    });
    return c;
  }

  function normQty(parts) {
    var q = parts.join(' ').replace(/\s+/g, ' ').trim();
    q = q.replace(/\s*([0-9.+]+)\s*m\s*[\u0001\u0000-\u001f²2]/g, '$1 m²').replace(/[\u0000-\u001f]/g, '');
    q = q.replace(/(\d)\s+(樘|式|組|個|座|只|支|處|套|台|樘)/g, '$1$2');
    q = q.replace(/\(詳契約詳細\s*表\)/, '(詳契約詳細表)');
    return q;
  }

  /* pdf.js 會依字型把一行切成多段：相鄰且間距很小的片段先合併 */
  function mergeRuns(items) {
    var arr = items.filter(function (it) { return it.str != null; }).slice().sort(function (a, b) { return (Math.round(a.y) - Math.round(b.y)) || (a.x - b.x); });
    var out = [];
    arr.forEach(function (it) {
      var last = out[out.length - 1];
      if (last && it.w != null && last.w != null && Math.abs(it.y - last.y) < 1.5 && it.x - (last.x + last.w) < 1.6 && it.x >= last.x) {
        last.str += it.str; last.w = it.x + it.w - last.x;
      } else out.push({ str: it.str, x: it.x, y: it.y, w: it.w });
    });
    return out;
  }

  function parsePages(pages) {
    var out = [], cols = null;
    pages.forEach(function (pg) {
      var items = mergeRuns(pg.items).map(function (it) { return { s: clean(it.str), x: it.x, y: it.y, lead: /^[\t ]/.test(it.str) }; }).filter(function (it) { return it.s; });
      var pc = findCols(items);
      if (pc.vendor != null && pc.qty != null) cols = Object.assign({}, cols || {}, pc);
      if (!cols) return;
      var nameMax = cols.qty - 25;
      var anchors = items.filter(function (it) { return ANCHOR.test(it.s) && it.x < nameMax; }).sort(function (a, b) { return a.y - b.y; });
      var footY = Infinity;
      items.forEach(function (it) { if (/^註[:：]/.test(it.s) && it.y < footY) footY = it.y; });

      var rows = anchors.map(function (a, i) {
        var next = anchors[i + 1];
        var y0 = a.y - 3, y1 = Math.min(next ? next.y - 3 : a.y + 55, footY - 2);
        var band = items.filter(function (it) { return it !== a && it.y >= y0 && it.y < y1 && !ANCHOR.test(it.s); });
        var farRight = cols.review != null ? cols.review - 5 : 1e9;
        var name = band.filter(function (it) { return it.y > a.y + 2 && (it.x < nameMax || (it.lead && it.x >= farRight)); })
          .sort(function (p, q) { return (p.y - q.y) || (p.x - q.x); });
        // 同一行的片段以空格相連，跨行直接相接
        var nm = '', lastY = null;
        name.forEach(function (it) { nm += (lastY != null && Math.abs(it.y - lastY) < 3 ? ' ' : '') + it.s; lastY = it.y; });
        var qtyParts = band.filter(function (it) { return it.x >= nameMax && it.x < (cols.sampleTest || cols.qty + 50) - 8; })
          .sort(function (p, q) { return Math.abs(p.y - q.y) < 4 ? p.x - q.x : p.y - q.y; }).map(function (it) { return it.s; });
        var near = function (colX, it) { return colX != null && Math.abs(it.x - colX) < 14; };
        var yn = function (colX) {
          var hit = band.filter(function (it) { return (it.s === '是' || it.s === '否') && near(colX, it); })[0];
          return hit ? hit.s : '';
        };
        var docs = { vendor: false, catalog: false, test: false, sample: false, other: '' };
        var docCols = DOC_KEYS.map(function (k) { return { k: k, x: cols[k] }; }).filter(function (d) { return d.x != null; });
        band.forEach(function (it) {
          if (it.x < (cols.vendor || 0) - 10) return;
          if (it.x >= farRight) return;
          var best = null, bd = 1e9;
          docCols.forEach(function (d) { var dd = Math.abs(it.x - d.x - 5); if (dd < bd) { bd = dd; best = d.k; } });
          if (!best || bd > 18) return;
          if (/^[ˇvV✓✔√]$/.test(it.s)) { if (best === 'other') docs.other = docs.other || 'ˇ'; else docs[best] = true; }
          else if (best === 'other') docs.other = (docs.other === 'ˇ' ? '' : docs.other) + it.s;
        });
        return { no: a.s, name: nm, qty: normQty(qtyParts), sampleTest: yn(cols.sampleTest), factory: yn(cols.factory), docs: docs, _y: a.y, _gap: next ? next.y - a.y : 99 };
      });

      // 兩個項次共用一列（例如「八.一.(二)」「八.二.(二)」各式電線）：沒有名稱的併入下一列
      for (var i = rows.length - 2; i >= 0; i--) {
        var r = rows[i], n = rows[i + 1];
        if (!r.name && r._gap < 25) {
          n.no = r.no + '、' + n.no;
          if (!n.qty) n.qty = r.qty; else if (r.qty && n.qty.indexOf(r.qty) < 0) n.qty = r.qty + ' ' + n.qty;
          n.qty = normQty([n.qty]);
          if (!n.sampleTest) n.sampleTest = r.sampleTest;
          if (!n.factory) n.factory = r.factory;
          DOC_KEYS.forEach(function (k) { if (!n.docs[k] && r.docs[k]) n.docs[k] = r.docs[k]; });
          rows.splice(i, 1);
        }
      }
      rows.forEach(function (r) { delete r._y; delete r._gap; out.push(r); });
    });
    return out;
  }

  /* Excel / CSV：二維陣列，自動找標題列 */
  function parseRows(rows) {
    var hi = -1, map = {};
    for (var i = 0; i < Math.min(rows.length, 15); i++) {
      var line = rows[i].map(function (c) { return String(c || '').replace(/\s+/g, ''); });
      if (line.some(function (c) { return c.indexOf('材料') >= 0 || c.indexOf('名稱') >= 0; })) {
        hi = i;
        line.forEach(function (c, j) {
          if (/項次/.test(c) && map.no == null) map.no = j;
          if (/材料|名稱/.test(c) && map.name == null) map.name = j;
          if (/數量/.test(c)) map.qty = j;
          if (/取樣|試驗$/.test(c) && map.sampleTest == null && !/報告/.test(c)) map.sampleTest = j;
          if (/驗廠/.test(c) && map.factory == null) map.factory = j;
          if (/協力|廠商資料/.test(c)) map.vendor = j;
          if (/型錄/.test(c)) map.catalog = j;
          if (/試驗報告|報告/.test(c)) map.test = j;
          if (/樣品|色票/.test(c)) map.sample = j;
          if (/^其他/.test(c)) map.other = j;
          if (/進場/.test(c)) map.arrivalDate = j;
          if (/抽查日期/.test(c)) map.inspDate = j;
          if (/備註/.test(c)) map.remark = j;
        });
        break;
      }
    }
    if (hi < 0 || map.name == null) return [];
    var val = function (r, k) { return map[k] == null ? '' : String(r[map[k]] == null ? '' : r[map[k]]).trim(); };
    var tick = function (v) { return /^(ˇ|v|V|✓|✔|√|是|有|y|Y|1|true)$/.test(v); };
    return rows.slice(hi + 1).filter(function (r) { return val(r, 'name'); }).map(function (r) {
      var o = val(r, 'other');
      return {
        no: val(r, 'no'), name: val(r, 'name'), qty: val(r, 'qty'),
        sampleTest: val(r, 'sampleTest'), factory: val(r, 'factory'),
        docs: { vendor: tick(val(r, 'vendor')), catalog: tick(val(r, 'catalog')), test: tick(val(r, 'test')), sample: tick(val(r, 'sample')), other: o },
        arrivalDate: val(r, 'arrivalDate'), inspDate: val(r, 'inspDate'), remark: val(r, 'remark')
      };
    });
  }

  root.MatParse = { parsePages: parsePages, parseRows: parseRows, mergeRuns: mergeRuns };
})(typeof window !== 'undefined' ? window : globalThis);
