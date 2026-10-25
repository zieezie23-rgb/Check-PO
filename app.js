/* Pengecekan PO Expired — antarmuka */
(function () {
  'use strict';
  var C = window.PoCore;
  var $ = function (id) { return document.getElementById(id); };

  var KINDS = [
    { id: 'master', label: 'Master Stock', hint: 'Master Stock*.xlsb', accept: '.xlsb,.xlsx,.xls', test: /master/i },
    { id: 'pom', label: 'POMGG', hint: 'POMGG.csv', accept: '.csv,.txt', test: /pom/i },
    { id: 'stockOut', label: 'Stock Out', hint: 'Stock Out.xlsx', accept: '.xlsx,.xlsb,.xls', test: /stock\s*[-_]?\s*out|stockout/i },
    { id: 'promo', label: 'Promo', hint: 'Promo.xlsx', accept: '.xlsx,.xlsb,.xls', test: /promo/i },
    { id: 'seasonal', label: 'Seasonal', hint: 'Laporan Harian Seasonal*.xlsx (sheet BASIC)', accept: '.xlsx,.xlsb,.xls', test: /seasonal/i }
  ];
  var S = { master: null, pom: null, stockOut: null, promo: null, seasonal: null };
  var META = {};
  var lastRes = null;

  /* ---------- util UI ---------- */
  var toastTimer;
  function toast(msg) {
    var t = $('toast'); t.textContent = msg; t.classList.add('show');
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.classList.remove('show'); }, 2600);
  }
  function tick() { return new Promise(function (r) { setTimeout(r, 30); }); }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function hhmm(ts) { var d = new Date(ts); return String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); }
  function nf(n) { return Number(n).toLocaleString('id-ID'); }

  /* ---------- penyimpanan (IndexedDB) ---------- */
  var dbp = null;
  function db() {
    if (!dbp) dbp = new Promise(function (res, rej) {
      var rq = indexedDB.open('po-expired', 1);
      rq.onupgradeneeded = function () { rq.result.createObjectStore('kv'); };
      rq.onsuccess = function () { res(rq.result); };
      rq.onerror = function () { rej(rq.error); };
    });
    return dbp;
  }
  function idbGet(k) { return db().then(function (d) { return new Promise(function (res, rej) { var r = d.transaction('kv').objectStore('kv').get(k); r.onsuccess = function () { res(r.result); }; r.onerror = function () { rej(r.error); }; }); }); }
  function idbSet(k, v) { return db().then(function (d) { return new Promise(function (res, rej) { var t = d.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = res; t.onerror = function () { rej(t.error); }; }); }); }

  /* ---------- SheetJS (dimuat saat dibutuhkan) ---------- */
  var XLSX_URLS = ['https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'];
  var xlsxP = null;
  function ensureXLSX() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (xlsxP) return xlsxP;
    xlsxP = new Promise(function (res, rej) {
      var i = 0;
      (function next() {
        if (i >= XLSX_URLS.length) { xlsxP = null; return rej(new Error('Pustaka pembaca Excel gagal dimuat. Hubungkan ke internet sekali, lalu coba lagi.')); }
        var s = document.createElement('script'); s.src = XLSX_URLS[i++]; s.crossOrigin = 'anonymous';
        s.onload = function () { window.XLSX ? res(window.XLSX) : next(); };
        s.onerror = function () { s.remove(); next(); };
        document.head.appendChild(s);
      })();
    });
    return xlsxP;
  }

  var SHEET_TEST = {
    master: function (rows) { for (var i = 0; i < Math.min(rows.length, 15); i++) if (rows[i] && String(rows[i][0] == null ? '' : rows[i][0]).trim().toUpperCase() === 'PRDCD') return true; return false; },
    stockOut: function (rows) { for (var i = 0; i < Math.min(rows.length, 15); i++) { var u = (rows[i] || []).map(function (x) { return String(x == null ? '' : x).trim().toUpperCase(); }); if (u.indexOf('PRDCD') >= 0 && u.indexOf('KL') >= 0) return true; } return false; },
    promo: function (rows) { for (var i = 0; i < Math.min(rows.length, 15); i++) if (rows[i] && String(rows[i][0] == null ? '' : rows[i][0]).trim().toUpperCase() === 'PLU') return true; return false; }
  };
  var PARSERS = { master: C.parseMaster, stockOut: C.parseStockOut, promo: C.parsePromo, seasonal: C.parseSeasonal };

  /* file Seasonal besar (banyak sheet, sheet BASIC sangat lebar): baca hanya sheet BASIC dan batasi kolom */
  function readSeasonalRows(X, buf) {
    var names = X.read(buf, { type: 'array', bookSheets: true }).SheetNames;
    var nm = names.filter(function (n) { return /^basic$/i.test(String(n).trim()); })[0];
    if (!nm) throw new Error('Sheet "BASIC" tidak ditemukan di file Seasonal.');
    var wb = X.read(buf, { type: 'array', dense: true, sheets: nm, cellFormula: false, cellStyles: false, cellHTML: false, cellNF: false });
    var ws = wb.Sheets[nm];
    var rg = X.utils.decode_range(ws['!ref']); if (rg.e.c > 45) rg.e.c = 45;
    ws['!ref'] = X.utils.encode_range(rg);
    return X.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  }

  function readWorkbookRows(X, buf, kind) {
    var wb = X.read(buf, { type: 'array', dense: true, cellFormula: false, cellStyles: false, cellHTML: false, cellNF: false });
    for (var i = 0; i < wb.SheetNames.length; i++) {
      var rows = X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[i]], { header: 1, raw: true, defval: null });
      if (SHEET_TEST[kind](rows)) return rows;
    }
    throw new Error('Sheet yang berisi data ' + (KINDS.filter(function (k) { return k.id === kind; })[0].label) + ' tidak ditemukan di file ini.');
  }

  /* ---------- kartu bahan ---------- */
  function buildCards() {
    var box = $('bahan'); box.innerHTML = '';
    KINDS.forEach(function (k) {
      var d = document.createElement('div'); d.className = 'bh'; d.id = 'bh-' + k.id; d.title = k.hint;
      d.innerHTML = '<b>' + k.label + '</b>' +
        '<div class="st"><span class="dot"></span><span class="msg">Belum ada file</span><button type="button">Pilih file</button></div>' +
        '<input type="file" hidden accept="' + k.accept + '">';
      var inp = d.querySelector('input');
      d.querySelector('button').onclick = function () { inp.click(); };
      inp.onchange = function () { if (inp.files[0]) loadFile(k.id, inp.files[0]); inp.value = ''; };
      box.appendChild(d);
    });
  }
  function setCard(kind, state, msg) {
    var d = $('bh-' + kind); d.className = 'bh ' + state; d.querySelector('.msg').textContent = msg;
  }
  function cardOk(kind) {
    var m = META[kind]; if (!m) return;
    setCard(kind, 'ok', m.name + ' · ' + nf(m.count) + ' baris · ' + hhmm(m.at));
  }
  function refreshReady() {
    var ok = !!(S.master && S.pom);
    $('btnRun').disabled = !ok;
    var miss = [], note = [];
    if (!S.stockOut) { miss.push('Stock Out'); note.push('penilaian Stock Out tidak aktif'); }
    if (!S.promo) { miss.push('Promo'); note.push('kolom promo kosong'); }
    if (!S.seasonal) { miss.push('Seasonal'); note.push('teks khusus item Basic tidak aktif'); }
    var w = $('warn');
    if (ok && miss.length) { w.hidden = false; w.textContent = 'Bahan ' + miss.join(', ') + ' belum dimuat — hasil tetap bisa dihitung, tetapi ' + note.join(', ') + '.'; }
    else if (!ok) { w.hidden = false; w.textContent = 'Muat Master Stock dan POMGG dulu supaya No PO bisa dicek.'; }
    else w.hidden = true;
  }

  async function loadFile(kind, file) {
    setCard(kind, 'busy', 'Membaca ' + file.name + '…');
    await tick();
    try {
      var data;
      if (kind === 'pom') {
        var text = new TextDecoder('utf-8').decode(await file.arrayBuffer());
        data = C.parsePOM(text);
      } else {
        var X = await ensureXLSX();
        var buf = await file.arrayBuffer();
        await tick();
        data = PARSERS[kind](kind === 'seasonal' ? readSeasonalRows(X, buf) : readWorkbookRows(X, buf, kind));
      }
      S[kind] = data;
      META[kind] = { name: file.name, at: Date.now(), count: kind === 'master' || kind === 'pom' ? data.count : data.count };
      cardOk(kind);
      if (data.warn && data.warn.length) toast('Master: kolom ' + data.warn.join(', ') + ' tidak ditemukan di header');
      idbSet(kind, { meta: META[kind], data: data }).catch(function () { });
    } catch (e) {
      console.error(e);
      setCard(kind, 'err', e.message || String(e));
      toast('Gagal membaca ' + file.name);
    }
    refreshReady();
  }

  function classify(file) {
    var n = file.name;
    for (var i = 0; i < KINDS.length; i++) if (KINDS[i].test.test(n)) return KINDS[i].id;
    if (/\.csv$/i.test(n)) return 'pom';
    return null;
  }
  async function loadMany(files) {
    var list = Array.prototype.slice.call(files);
    for (var i = 0; i < list.length; i++) {
      var k = classify(list[i]);
      if (!k) { toast('Nama file tidak dikenali: ' + list[i].name); continue; }
      await loadFile(k, list[i]);
    }
  }

  /* ---------- pengaturan ---------- */
  function loadOpts() {
    var o = {}; try { o = JSON.parse(localStorage.getItem('po_opts') || '{}'); } catch (e) { }
    $('optDsi').checked = o.dsiStd !== false;
    var box = $('stdBox'); box.innerHTML = '';
    for (var f = 1; f <= 7; f++) {
      var v = (o.std && o.std[f] != null) ? o.std[f] : C.DEFAULT_STD[f];
      var l = document.createElement('label'); l.innerHTML = 'FK ' + f + ' <input type="number" min="0" data-fk="' + f + '" value="' + v + '">';
      box.appendChild(l);
    }
  }
  function readOpts() {
    var std = {}; document.querySelectorAll('#stdBox input').forEach(function (i) { std[i.dataset.fk] = Number(i.value) || 0; });
    return { dsiStd: $('optDsi').checked, std: std };
  }
  function saveOpts() { try { localStorage.setItem('po_opts', JSON.stringify(readOpts())); } catch (e) { } }

  /* ---------- hasil ---------- */
  function run() {
    var inputs = $('poInput').value.split(/[\s,;]+/).filter(Boolean);
    if (!inputs.length) { toast('Isi No PO dulu'); return; }
    var SS = { master: S.master, pom: S.pom, stockOut: S.stockOut || { map: new Map() }, promo: S.promo || { map: new Map() }, seasonal: S.seasonal || { map: new Map() } };
    var o = readOpts(); saveOpts();
    lastRes = C.compute(inputs, SS, o);
    renderResult();
  }

  function renderSummary() {
    if (!lastRes) return;
    var ta = $('kesimpulan'); ta.value = C.buildSummary(lastRes);
    ta.style.height = 'auto'; ta.style.height = Math.min(Math.max(ta.scrollHeight + 4, 60), 2400) + 'px';
  }

  function renderResult() {
    var r = lastRes, w = $('warn');
    $('secHasil').hidden = false;
    if (r.notFound.length) { w.hidden = false; w.textContent = 'Tidak ditemukan di POMGG: ' + r.notFound.join(', '); } else refreshReady();
    renderSummary();
    var pos = new Set(r.rows.map(function (x) { return x.po; })).size;
    $('stat').textContent = r.rows.length + ' item dari ' + pos + ' PO';
    $('btnJpg').disabled = !r.rows.length;
    $('btnCopy').disabled = !r.rows.length;
    if (!r.rows.length) { $('tableWrap').innerHTML = '<div class="empty">Tidak ada item untuk No PO tersebut.</div>'; return; }

    var cols = C.COLS, h1 = '', h2 = '', i, c;
    var colg = '<colgroup>' + cols.map(function (c) { return '<col style="width:' + c.px + 'px">'; }).join('') + '</colgroup>';
    for (i = 0; i < cols.length; i++) {
      c = cols[i];
      if (c.spacer) { h1 += '<th class="sp" rowspan="2"></th>'; continue; }
      if (c.g === 'promo') {
        if (!cols[i - 1] || cols[i - 1].g !== 'promo') h1 += '<th class="grp" colspan="4">PROMOSI (TERDEKAT)</th>';
        h2 += '<th>' + esc(c.h) + '</th>';
      } else h1 += '<th rowspan="2">' + esc(c.h) + '</th>';
    }
    var body = r.rows.map(function (row) {
      return '<tr>' + cols.map(function (c) {
        if (c.spacer) return '<td class="sp"></td>';
        var t = C.fmtCell(c, row[c.k]);
        return '<td class="' + c.a + (c.k === 'kes' ? ' kes' : '') + '" title="' + esc(t) + '">' + esc(t) + '</td>';
      }).join('') + '</tr>';
    }).join('');
    var wdt = cols.reduce(function (a, c) { return a + c.px; }, 0);
    $('tableWrap').innerHTML = '<table class="sheet" style="width:' + wdt + 'px">' + colg + '<thead><tr>' + h1 + '</tr><tr class="r2">' + h2 + '</tr></thead><tbody>' + body + '</tbody></table>';
    var tr0 = document.querySelector('#tableWrap thead tr'), hh = tr0 ? tr0.getBoundingClientRect().height : 0;
    document.querySelectorAll('#tableWrap thead tr.r2 th').forEach(function (th) { th.style.top = Math.floor(hh) + 'px'; });
    $('secHasil').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  /* ---------- salin ---------- */
  function copyText() {
    var ta = $('kesimpulan'), txt = ta.value;
    function fallback() { ta.focus(); ta.select(); var ok = false; try { ok = document.execCommand('copy'); } catch (e) { } toast(ok ? 'Teks disalin' : 'Gagal menyalin — blok teks lalu salin manual'); }
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(txt).then(function () { toast('Teks disalin'); }, fallback); else fallback();
  }

  /* ---------- export JPEG (digambar langsung ke canvas) ---------- */
  function wrapLines(ctx, text, maxW) {
    var words = String(text).split(/\s+/), lines = [], cur = '';
    words.forEach(function (w) { var t = cur ? cur + ' ' + w : w; if (cur && ctx.measureText(t).width > maxW) { lines.push(cur); cur = w; } else cur = t; });
    if (cur) lines.push(cur); return lines;
  }
  function exportJPEG() {
    if (!lastRes || !lastRes.rows.length) return;
    var cols = C.COLS, rows = lastRes.rows;
    var rowH = 26, headH = 52, grpH = 24, titleH = 32, dateH = 24, pad = 14, FONT = 'Calibri, Carlito, "Segoe UI", Arial, sans-serif';
    var W = cols.reduce(function (a, c) { return a + c.px; }, 0) + pad * 2;
    var H = pad + titleH + dateH + grpH + headH + rows.length * rowH + pad;
    var sc = Math.min(2, Math.sqrt(16e6 / (W * H)), 16000 / Math.max(W, H)); sc = Math.max(sc, 0.4);
    var cv = document.createElement('canvas'); cv.width = Math.round(W * sc); cv.height = Math.round(H * sc);
    var ctx = cv.getContext('2d'); ctx.scale(sc, sc);
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
    ctx.fillStyle = '#000'; ctx.textBaseline = 'middle';
    ctx.font = 'bold 21px ' + FONT; ctx.textAlign = 'left'; ctx.fillText('PO EXPIRED DC BOGOR1', pad, pad + titleH / 2);
    ctx.font = 'bold 15px ' + FONT; ctx.fillText(C.fmtTodayID(new Date()), pad, pad + titleH + dateH / 2);
    var y0 = pad + titleH + dateH, yh = y0 + grpH;
    var xs = [], x = pad; cols.forEach(function (c) { xs.push(x); x += c.px; });

    function cell(x, y, w, h, fill) {
      if (fill) { ctx.fillStyle = fill; ctx.fillRect(x, y, w, h); }
      ctx.strokeStyle = '#000'; ctx.lineWidth = 1; ctx.strokeRect(x + .5, y + .5, w, h);
    }
    // header
    var gi = cols.map(function (c, i) { return c.g === 'promo' ? i : -1; }).filter(function (i) { return i >= 0; });
    if (gi.length) {
      var gx = xs[gi[0]], gw = gi.reduce(function (a, i) { return a + cols[i].px; }, 0);
      cell(gx, y0, gw, grpH, '#C5D9F1');
      ctx.fillStyle = '#000'; ctx.font = 'bold 14px ' + FONT; ctx.textAlign = 'center'; ctx.fillText('PROMOSI (TERDEKAT)', gx + gw / 2, y0 + grpH / 2);
    }
    cols.forEach(function (c, i) {
      if (c.spacer) return;
      var isP = c.g === 'promo', hy = isP ? yh : y0, hh = isP ? headH : grpH + headH;
      cell(xs[i], hy, c.px, hh, '#C5D9F1');
      ctx.fillStyle = '#000'; ctx.font = 'bold 14px ' + FONT; ctx.textAlign = 'center';
      var ls = wrapLines(ctx, c.h, c.px - 8), lh = 16, ty = hy + hh / 2 - (ls.length - 1) * lh / 2;
      ls.forEach(function (l, n) { ctx.fillText(l, xs[i] + c.px / 2, ty + n * lh); });
    });
    // isi
    ctx.font = '14.5px ' + FONT;
    rows.forEach(function (row, ri) {
      var y = yh + headH + ri * rowH;
      cols.forEach(function (c, i) {
        if (c.spacer) return;
        cell(xs[i], y, c.px, rowH, null);
        var t = C.fmtCell(c, row[c.k]); if (!t) return;
        ctx.save(); ctx.beginPath(); ctx.rect(xs[i] + 1, y + 1, c.px - 1, rowH - 1); ctx.clip();
        ctx.fillStyle = '#000'; ctx.font = (c.k === 'kes' ? 'bold ' : '') + '14.5px ' + FONT;
        if (c.a === 'c') { ctx.textAlign = 'center'; ctx.fillText(t, xs[i] + c.px / 2, y + rowH / 2); }
        else if (c.a === 'r') { ctx.textAlign = 'right'; ctx.fillText(t, xs[i] + c.px - 5, y + rowH / 2); }
        else { ctx.textAlign = 'left'; ctx.fillText(t, xs[i] + 5, y + rowH / 2); }
        ctx.restore();
      });
    });
    cv.toBlob(function (b) {
      if (!b) { toast('Gagal membuat gambar — kurangi jumlah PO'); return; }
      var d = new Date(), p = function (n) { return String(n).padStart(2, '0'); };
      var a = document.createElement('a'); a.href = URL.createObjectURL(b);
      a.download = 'PO_Expired_' + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes()) + '.jpg';
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(function () { URL.revokeObjectURL(a.href); }, 4000);
      toast('JPEG tersimpan');
    }, 'image/jpeg', 0.93);
  }

  /* ---------- init ---------- */
  async function restore() {
    for (var i = 0; i < KINDS.length; i++) {
      try {
        var rec = await idbGet(KINDS[i].id);
        if (rec && rec.data) { S[KINDS[i].id] = rec.data; META[KINDS[i].id] = rec.meta; cardOk(KINDS[i].id); }
      } catch (e) { }
    }
    refreshReady();
  }

  function init() {
    $('subtitle').textContent = 'PO EXPIRED DC BOGOR1 · ' + C.fmtTodayID(new Date());
    buildCards(); loadOpts(); refreshReady();
    $('btnPickAll').onclick = function () { $('fileAll').click(); };
    $('fileAll').onchange = function (e) { loadMany(e.target.files); e.target.value = ''; };
    $('btnRun').onclick = run;
    $('btnClear').onclick = function () { $('poInput').value = ''; $('poInput').focus(); };
    $('btnCopy').onclick = copyText;
    $('btnJpg').onclick = exportJPEG;
    $('optDsi').onchange = function () { saveOpts(); if (lastRes && $('poInput').value.trim()) run(); };
    $('stdBox').onchange = function () { saveOpts(); if ($('optDsi').checked && lastRes) run(); };
    $('poInput').addEventListener('keydown', function (e) { if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') run(); });
    var bx = $('secBahan');
    ['dragenter', 'dragover'].forEach(function (n) { bx.addEventListener(n, function (e) { e.preventDefault(); $('bahan').classList.add('drag'); }); });
    ['dragleave', 'drop'].forEach(function (n) { bx.addEventListener(n, function (e) { e.preventDefault(); $('bahan').classList.remove('drag'); }); });
    bx.addEventListener('drop', function (e) { if (e.dataTransfer && e.dataTransfer.files.length) loadMany(e.dataTransfer.files); });

    var dp = null;
    window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); dp = e; $('btnInstall').hidden = false; });
    $('btnInstall').onclick = function () { if (dp) { dp.prompt(); dp = null; $('btnInstall').hidden = true; } };
    if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) navigator.serviceWorker.register('sw.js').catch(function () { });
    restore();
  }
  init();
})();
