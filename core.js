/* Pengecekan PO Expired — logika inti (tanpa DOM, bisa dites di Node) */
(function (root) {
  'use strict';

  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var BLN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  var HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];

  var BLN_ID = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  var DEFAULT_STD = { 1: 21, 2: 15, 3: 10, 4: 7, 5: 7, 6: 7, 7: 7 };

  /* ---------- util ---------- */
  function str(v) { return v == null ? '' : String(v).trim(); }
  function key(v) { return str(v).replace(/\.0+$/, '').replace(/^0+/, ''); }
  function toNum(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (typeof v === 'string') {
      var s = v.trim().replace(/,/g, '');
      if (s === '' || isNaN(Number(s))) return null;
      return Number(s);
    }
    return null;
  }
  function n0(v) { var n = toNum(v); return n == null ? 0 : n; }

  function serialToDate(n) {
    var d = new Date(Math.round((n - 25569) * 86400000));
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }
  function parseDate(v, mdy) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
    if (typeof v === 'number') return (v > 20000 && v < 100000) ? serialToDate(v) : null;
    var s = String(v).trim(), m;
    if ((m = /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/.exec(s))) return new Date(+m[1], +m[2] - 1, +m[3]);
    if ((m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})/.exec(s))) {
      var y = +m[3], a = +m[1], b = +m[2]; if (y < 100) y += 2000;
      /* mdy: format Amerika (bulan/tanggal/tahun) — dipakai kecuali angka pertama jelas bukan bulan (>12) */
      if (mdy && a <= 12) return new Date(y, a - 1, b);
      return new Date(y, b - 1, a);
    }
    if (/^\d{5}(\.\d+)?$/.test(s)) return serialToDate(Number(s));
    var t = Date.parse(s);
    if (!isNaN(t)) { var d = new Date(t); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
    return null;
  }

  /* ---------- CSV ---------- */
  function detectDelim(text) {
    var nl = text.indexOf('\n'); var line = text.substring(0, nl < 0 ? Math.min(text.length, 5000) : nl);
    var best = ',', bc = -1;
    [',', ';', '\t', '|'].forEach(function (d) { var c = line.split(d).length; if (c > bc) { bc = c; best = d; } });
    return best;
  }
  function parseCSV(text, delim, cb) {
    var n = text.length;
    if (text.indexOf('"') < 0) {
      var start = 0;
      while (start < n) {
        var end = text.indexOf('\n', start); if (end < 0) end = n;
        var line = text.substring(start, end);
        if (line.charCodeAt(line.length - 1) === 13) line = line.slice(0, -1);
        start = end + 1;
        if (line === '') continue;
        cb(line.split(delim));
      }
      return;
    }
    var i = 0, rec = [], f = '', inQ = false, dc = delim.charCodeAt(0);
    while (i < n) {
      var ch = text.charCodeAt(i);
      if (inQ) {
        if (ch === 34) { if (text.charCodeAt(i + 1) === 34) { f += '"'; i += 2; continue; } inQ = false; i++; continue; }
        f += text[i]; i++; continue;
      }
      if (ch === 34) { inQ = true; i++; continue; }
      if (ch === dc) { rec.push(f); f = ''; i++; continue; }
      if (ch === 10 || ch === 13) {
        if (ch === 13 && text.charCodeAt(i + 1) === 10) i++;
        i++; rec.push(f); f = '';
        if (rec.length > 1 || rec[0] !== '') cb(rec);
        rec = []; continue;
      }
      f += text[i]; i++;
    }
    if (f !== '' || rec.length) { rec.push(f); cb(rec); }
  }

  /* ---------- parser bahan ---------- */
  function parsePOM(text) {
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    var delim = detectDelim(text);
    var head = null, idx = null, rows = [], maxDate = new Map();
    parseCSV(text, delim, function (rec) {
      if (!head) {
        head = rec.map(function (s) { return s.trim().toUpperCase(); });
        idx = { doc: head.indexOf('DOCNO'), supco: head.indexOf('SUPCO'), prd: head.indexOf('PRDCD'), qty: head.indexOf('QTY'), date: head.indexOf('PO_DATE'), umur: [] };
        head.forEach(function (h, j) { if (/^UMUR/.test(h)) { if (h === 'UMUR_PO') idx.umur.unshift(j); else idx.umur.push(j); } });
        var miss = [];
        if (idx.doc < 0) miss.push('DOCNO'); if (idx.supco < 0) miss.push('SUPCO'); if (idx.prd < 0) miss.push('PRDCD');
        if (idx.qty < 0) miss.push('QTY'); if (idx.date < 0) miss.push('PO_DATE'); if (!idx.umur.length) miss.push('UMUR_PO');
        if (miss.length) throw new Error('Kolom tidak ditemukan di POMGG.csv: ' + miss.join(', '));
        return;
      }
      var doc = str(rec[idx.doc]); if (!doc) return;
      var prd = str(rec[idx.prd]);
      var dt = parseDate(rec[idx.date], true);
      var ms = dt ? dt.getTime() : null;
      var umur = 0;   /* ada beberapa kolom UMUR*: pakai yang pertama berisi (UMUR_PO diutamakan) */
      for (var u = 0; u < idx.umur.length; u++) { var uv = rec[idx.umur[u]]; if (uv != null && String(uv).trim() !== '') { umur = n0(uv); break; } }
      var exp = null;
      if (dt) { var de = new Date(dt.getTime()); de.setDate(de.getDate() + (umur || 0)); exp = de.getTime(); }
      rows.push({ doc: doc, docU: doc.toUpperCase(), supco: str(rec[idx.supco]), prd: prd, qty: n0(rec[idx.qty]), date: exp });
      if (ms != null) { var k = key(prd); var o = maxDate.get(k); if (o == null || ms > o) maxDate.set(k, ms); }
    });
    if (!head) throw new Error('POMGG.csv kosong.');
    return { rows: rows, maxDate: maxDate, count: rows.length };
  }

  var MASTER_F = {
    prd: [1, 'PRDCD'], desc: [2, 'DESC2'], supco: [10, 'SUPCO'], snama: [11, 'SNAMA'], fk: [12, 'FK'], frac: [16, 'FRAC'],
    pkm: [19, 'PKM'], ptag: [21, 'PTAG'], avg: [25, 'AVG'], mean: [29, 'MEAN'], stk: [38, 'STKG113'], poout: [45, 'POG113'],
    keb: [53, 'KEB'], protect: [64, 'PROTECT'], lead: [66, 'LEAD']
  };
  function parseMaster(rows) {
    var h = -1, i;
    for (i = 0; i < Math.min(rows.length, 15); i++) { if (rows[i] && str(rows[i][0]).toUpperCase() === 'PRDCD') { h = i; break; } }
    if (h < 0) throw new Error('Header "PRDCD" tidak ditemukan di kolom A (baris 1-15) pada Master Stock.');
    var head = rows[h].map(function (x) { return str(x).toUpperCase(); });
    var idx = {}, warn = [];
    Object.keys(MASTER_F).forEach(function (k) {
      var pos = MASTER_F[k][0] - 1, name = MASTER_F[k][1], f = head.indexOf(name);
      if (f >= 0) idx[k] = f; else { idx[k] = pos; warn.push(name + ' (pakai kolom ke-' + (pos + 1) + ')'); }
    });
    var map = new Map(), supMap = new Map(), count = 0;
    for (i = h + 1; i < rows.length; i++) {
      var r = rows[i]; if (!r) continue;
      var p = str(r[idx.prd]); if (!/^\d{4,}$/.test(p)) continue;
      count++;
      var k = key(p);
      var sup = str(r[idx.supco]);
      if (sup) { var su = sup.toUpperCase(); if (!supMap.has(su)) supMap.set(su, str(r[idx.snama])); }
      if (map.has(k)) continue;
      var pt = r[idx.ptag];
      map.set(k, {
        desc: str(r[idx.desc]), supco: sup, fk: str(r[idx.fk]), frac: n0(r[idx.frac]), pkm: n0(r[idx.pkm]),
        ptag: (pt == null || pt === '') ? 0 : pt, avg: n0(r[idx.avg]), mean: n0(r[idx.mean]), stk: n0(r[idx.stk]),
        poout: n0(r[idx.poout]), keb: n0(r[idx.keb]), protect: str(r[idx.protect]).toUpperCase(), lead: n0(r[idx.lead])
      });
    }
    if (!count) throw new Error('Tidak ada baris produk (PRDCD) di Master Stock.');
    return { map: map, supMap: supMap, count: count, warn: warn };
  }

  function parseStockOut(rows) {
    var h = -1, ip = -1, ik = -1, i, j;
    for (i = 0; i < Math.min(rows.length, 15) && h < 0; i++) {
      var r = rows[i]; if (!r) continue;
      var up = r.map(function (x) { return str(x).toUpperCase(); });
      var a = up.indexOf('PRDCD'), b = up.indexOf('KL');
      if (a >= 0 && b >= 0) { h = i; ip = a; ik = b; }
    }
    if (h < 0) throw new Error('Header "Prdcd" dan "KL" tidak ditemukan di Stock Out.');
    var map = new Map();
    for (j = h + 1; j < rows.length; j++) {
      var rr = rows[j]; if (!rr) continue;
      var k = key(rr[ip]); if (!k || !/^\d+$/.test(k)) continue;
      if (!map.has(k)) map.set(k, n0(rr[ik]));
    }
    return { map: map, count: map.size };
  }

  function parsePromo(rows) {
    var h = -1, i;
    for (i = 0; i < Math.min(rows.length, 15); i++) { if (rows[i] && str(rows[i][0]).toUpperCase() === 'PLU') { h = i; break; } }
    if (h < 0) throw new Error('Header "PLU" tidak ditemukan di kolom A pada Promo.');
    var map = new Map();
    for (i = h + 1; i < rows.length; i++) {
      var r = rows[i]; if (!r) continue;
      var k = key(r[0]); if (!k || !/^\d+$/.test(k) || map.has(k)) continue;
      var mn = null, mx = null;
      for (var c = 4; c <= 7; c++) {
        var d = parseDate(r[c]); if (!d) continue;
        var t = d.getTime(); if (mn == null || t < mn) mn = t; if (mx == null || t > mx) mx = t;
      }
      var tg = r[2];
      map.set(k, { target: (tg == null || tg === '') ? 0 : tg, jenis: str(r[3]), awal: mn, akhir: mx });
    }
    return { map: map, count: map.size };
  }

  function parseSeasonal(rows) {
    var h = -1, i;
    for (i = 0; i < Math.min(rows.length, 15); i++) {
      if (rows[i] && str(rows[i][0]).toUpperCase() === 'PRDCD' && rows[i].some(function (x) { return str(x).toUpperCase() === 'POLA1'; })) { h = i; break; }
    }
    if (h < 0) throw new Error('Header "PRDCD" dan "POLA1" tidak ditemukan di sheet BASIC (baris 1-15).');
    var head = rows[h].map(function (x) { return str(x).toUpperCase(); });
    var NAMES = { desc: 'DESC', frac: 'FRAC', pola1: 'POLA1', pola2: 'POLA2', tgt: 'TGT (IN QTY)', po: 'PO (IN QTY)', bpb: 'BPB (IN QTY)', poout: 'PO OUT (IN QTY)' };
    var idx = {}, miss = [];
    Object.keys(NAMES).forEach(function (k) { idx[k] = head.indexOf(NAMES[k]); if (idx[k] < 0) miss.push(NAMES[k]); });
    if (miss.length) throw new Error('Kolom tidak ditemukan di sheet BASIC: ' + miss.join(', '));
    var map = new Map();
    for (i = h + 1; i < rows.length; i++) {
      var r = rows[i]; if (!r) continue;
      var k = key(r[0]); if (!k || !/^\d+$/.test(k) || map.has(k)) continue;
      map.set(k, { desc: str(r[idx.desc]), frac: n0(r[idx.frac]), pola1: str(r[idx.pola1]).toUpperCase(), pola2: str(r[idx.pola2]), tgt: n0(r[idx.tgt]), po: n0(r[idx.po]), bpb: n0(r[idx.bpb]), poout: n0(r[idx.poout]) });
    }
    if (!map.size) throw new Error('Tidak ada baris produk di sheet BASIC.');
    return { map: map, count: map.size };
  }

  /* ---------- hitung ---------- */
  function dsi(a, L) { return (typeof a === 'number' && typeof L === 'number' && L !== 0) ? a / (L / 30) : ''; }
  function D(ms) { return ms == null ? '' : new Date(ms); }

  function buildRow(r, S, opts) {
    var k = key(r.prd);
    var m = S.master.map.get(k);
    var promo = S.promo.map.get(k);
    var sn = S.master.supMap.get(r.supco.toUpperCase());
    var o = {
      po: r.doc, tgl: D(r.date), supco: r.supco, snama: sn == null ? 0 : sn, prd: r.prd,
      desc: m ? m.desc : '', frac: m ? m.frac : '', ptag: m ? m.ptag : '', fk: m ? m.fk : '',
      avg: m ? m.avg : '', mean: m ? m.mean : '', qty: r.qty,
      pkm: m ? m.pkm : '', stk: m ? m.stk : '', poout: m ? m.poout : '', keb: m ? m.keb : '', lead: m ? m.lead : ''
    };
    o.dsiPo = dsi(r.qty, o.avg);
    var mx = S.pom.maxDate.get(k); o.maxTgl = mx == null ? '' : new Date(mx);
    o.over = (m ? Math.max(o.stk + o.poout + r.qty - o.pkm, 0) : '');
    o.dsiPkm = dsi(o.pkm, o.avg);
    o.dsiStk = dsi(o.stk, o.avg);
    o.avgPb = typeof o.avg === 'number' ? o.avg / 30 * 7 : '';
    o.promoAwal = promo ? D(promo.awal) : '';
    o.promoAkh = promo ? D(promo.akhir) : '';
    o.target = promo ? (toNum(promo.target) == null ? promo.target : toNum(promo.target)) : '';
    o.jenis = promo ? promo.jenis : '';
    o.ket = m ? (m.protect === 'PROTECT HO' ? 'CMO' : 'Reg') : '';
    var sea = S.seasonal && S.seasonal.map ? S.seasonal.map.get(k) : null;
    o.basic = (sea && sea.pola1 === 'BASIC') ? sea : null;

    var std = 0;
    if (opts.dsiStd) { var sv = (opts.std || DEFAULT_STD)[Number(o.fk)]; std = sv == null ? 0 : sv; }
    if (o.jenis !== '') o.kes = 'Promosi';
    else if (typeof o.poout === 'number' && typeof o.pkm === 'number' && o.poout < o.pkm) o.kes = 'PO Out Aktif kecil/ under PKM';
    else if (S.stockOut.map.get(k) === 1) o.kes = 'Stock Out';
    else if (typeof o.dsiStk === 'number' && o.dsiStk < std) o.kes = 'DSI under Standart';
    else if (opts.dsiStd && typeof o.dsiStk === 'number' && o.dsiStk > std) o.kes = 'DSI stock melebihi standart';
    else o.kes = '';
    return o;
  }

  function compute(inputs, S, opts) {
    opts = opts || {};
    var want = new Map();
    inputs.forEach(function (p) { var t = String(p).trim(), u = t.toUpperCase(); if (u && !want.has(u)) want.set(u, t); });
    var hit = new Set(), rows = [];
    var pr = S.pom.rows;
    for (var i = 0; i < pr.length; i++) {
      if (want.has(pr[i].docU)) { hit.add(pr[i].docU); rows.push(buildRow(pr[i], S, opts)); }
    }
    var notFound = [];
    want.forEach(function (orig, u) { if (!hit.has(u)) notFound.push(orig); });
    return { rows: rows, notFound: notFound, opts: opts };
  }

  /* ---------- kesimpulan teks ---------- */
  var KES_ORDER = ['Promosi', 'Stock Out', 'PO Out Aktif kecil/ under PKM', 'DSI under Standart', 'DSI stock melebihi standart'];

  function nfID(v, dec) {
    var r = Math.round(v * Math.pow(10, dec)) / Math.pow(10, dec), neg = r < 0;
    var p = Math.abs(r).toFixed(dec).split('.');
    var out = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + (p[1] && Number(p[1]) ? ',' + p[1].replace(/0+$/, '') : '');
    return (neg ? '-' : '') + out;
  }
  function pct(a, b) { return b > 0 ? nfID(a / b * 100, 1) + '%' : '-'; }
  function fmtTglID(d) {
    if (!(d instanceof Date) || isNaN(d)) return '...';
    return String(d.getDate()).padStart(2, '0') + ' ' + BLN_ID[d.getMonth()] + ' ' + String(d.getFullYear()).slice(2);
  }

  function aligned(pairs) {
    var w = 0;
    pairs.forEach(function (p) { if (p[0].length > w) w = p[0].length; });
    return pairs.map(function (p) { return p[0] + new Array(w - p[0].length + 1).join(' ') + ' : ' + p[1]; }).join('\n');
  }

  function buildBasic(items) {
    var out = 'Selamat pagi pak @Jonathan C\n\nHari ini di DC BGR1 ada kedatangan item basic:\n';
    items.forEach(function (r) {
      var b = r.basic, f = b.frac || (typeof r.frac === 'number' ? r.frac : 0) || 1;
      function ctn(q) { return q / f; }
      var qtyPo = ctn(r.qty), tgt = ctn(b.tgt), bpb = ctn(b.bpb), pout = ctn(b.poout);
      var ttl = qtyPo + bpb + pout, kurang = tgt - ttl;
      out += '\n' +
        aligned([
          ['Supp', r.supco + (r.snama ? ' - ' + r.snama : '')],
          ['No po', r.po],
          ['Ket', 'PO EXP TGL ' + fmtTglID(r.tgl)]
        ]) + '\n\n' +
        aligned([
          ['Plu', r.prd],
          ['Desc', r.desc || b.desc]
        ]) + '\n\n' +
        aligned([
          ['Qty PO datang (exp) in ctn', nfID(qtyPo, 2)],
          ['Target in ctn', nfID(tgt, 2)],
          ['BPB in ctn', nfID(bpb, 2)],
          ['PO Out in ctn', nfID(pout, 2)],
          ['BPB vs Target', pct(b.bpb, b.tgt)],
          ['PO + BPB vs Target', pct(b.bpb + b.poout, b.tgt)],
          ['Ttl PO Exp + BPB + PO Out', nfID(ttl, 2)]
        ]) + '\n\n' +
        'jika diterima,\n' +
        aligned([
          ['Kekurang In ctn', kurang > 0 ? nfID(kurang, 2) : '0 (melebihi target ' + nfID(-kurang, 2) + ')'],
          ['BPB vs Target', pct(b.bpb + r.qty, b.tgt)]
        ]) + '\n';
    });
    return out + '\nMohon arahannya apakah bisa diterima atau ditolak?';
  }

  function buildStd(rows) {
    function has(v) { return rows.some(function (r) { return r.kes === v; }); }
    var reasons = KES_ORDER.filter(has);
    var out = (has('DSI stock melebihi standart') ? 'Izin ditolak' : 'Izin diterima') + ' Pak @IDM Pak Rian PR,';
    if (reasons.length) out += ' Item ' + reasons.join(', ');
    return out;
  }

  function buildSummary(res) {
    var rows = res.rows; if (!rows.length) return '';
    var basics = rows.filter(function (r) { return r.basic; }), others = rows.filter(function (r) { return !r.basic; });
    var parts = [];
    if (basics.length) parts.push(buildBasic(basics));
    if (others.length) parts.push(buildStd(others));
    return parts.join('\n\n');
  }

  /* ---------- kolom & format tampilan (meniru sheet CHECK C:AD + AF) ---------- */
  var COLS = [
    { k: 'po', h: 'NO PO EXPIRED', w: 13.4, a: 'c', t: 'txt' },
    { k: 'tgl', h: 'TGL PO EKSPIRED', w: 12, a: 'r', t: 'date' },
    { k: 'supco', h: 'SUPCO', w: 9.3, a: 'c', t: 'txt' },
    { k: 'snama', h: 'SNAMA', w: 40, a: 'l', t: 'txt' },
    { k: 'prd', h: 'PRDCD', w: 11.9, a: 'c', t: 'txt' },
    { k: 'desc', h: 'DESC', w: 52, a: 'l', t: 'txt' },
    { k: 'frac', h: 'FRAC', w: 5.9, a: 'r', t: 'int' },
    { k: 'ptag', h: 'PTAG', w: 6.3, a: 'r', t: 'int' },
    { k: 'fk', h: 'FK', w: 5.6, a: 'c', t: 'txt' },
    { k: 'avg', h: 'AVG SLS', w: 8.1, a: 'c', t: 'int' },
    { k: 'mean', h: 'Mean', w: 8.1, a: 'c', t: 'dec' },
    { k: 'qty', h: 'QTY PO EXP', w: 8.6, a: 'r', t: 'int' },
    { k: 'dsiPo', h: 'DSI PO EXP', w: 7.3, a: 'r', t: 'int' },
    { k: 'pkm', h: 'PKM DC', w: 8.3, a: 'r', t: 'int' },
    { k: 'stk', h: 'STOCK DC', w: 9.6, a: 'r', t: 'int' },
    { k: 'poout', h: 'QTY PO OUT', w: 8.9, a: 'r', t: 'int' },
    { k: 'maxTgl', h: 'MAX TGL PO OUT', w: 12.1, a: 'r', t: 'date' },
    { k: 'over', h: 'OVER PKM', w: 8.9, a: 'r', t: 'int' },
    { k: 'keb', h: 'KBTHN TOKO', w: 9.1, a: 'r', t: 'int' },
    { k: 'dsiPkm', h: 'DSI PKM', w: 7.7, a: 'r', t: 'dec' },
    { k: 'dsiStk', h: 'DSI STOCK', w: 8.4, a: 'r', t: 'dec' },
    { k: 'lead', h: 'LEAD TIME', w: 7.4, a: 'r', t: 'int' },
    { k: 'avgPb', h: 'AVG PB TOKO (7_HARI)', w: 10.7, a: 'r', t: 'dec' },
    { k: 'promoAwal', h: 'TGL AWAL', w: 11.7, a: 'r', t: 'date', g: 'promo' },
    { k: 'promoAkh', h: 'TGL AKH', w: 11.7, a: 'r', t: 'date', g: 'promo' },
    { k: 'target', h: 'TARGET', w: 13.7, a: 'r', t: 'int', g: 'promo' },
    { k: 'jenis', h: 'JENIS PROMO', w: 9.3, a: 'l', t: 'txt', g: 'promo' },
    { k: 'ket', h: 'KET', w: 7.5, a: 'c', t: 'txt' },
    { spacer: true, w: 1.3 },
    { k: 'kes', h: 'Kesimpulan', w: 30.1, a: 'l', t: 'txt' }
  ];
  COLS.forEach(function (c) { c.px = Math.round(c.w * 7 + 5); });

  function commas(s) { return s.replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
  function fmtDate(d) {
    if (!(d instanceof Date) || isNaN(d)) return '';
    return d.getDate() + '-' + MON[d.getMonth()] + '-' + String(d.getFullYear()).slice(2);
  }
  function fmtCell(c, v) {
    if (v === '' || v == null) return '';
    if (c.t === 'date') return fmtDate(v);
    if (c.t === 'int' || c.t === 'dec') {
      if (typeof v !== 'number') return String(v);
      var r = c.t === 'int' ? Math.round(v) : Math.round(v * 100) / 100;
      if (r === 0) return '-';
      var neg = r < 0, s = Math.abs(r).toFixed(c.t === 'int' ? 0 : 2).split('.');
      var o = commas(s[0]) + (s[1] ? '.' + s[1] : '');
      return neg ? '(' + o + ')' : o;
    }
    return String(v);
  }
  function fmtTodayID(d) {
    d = d || new Date();
    return HARI[d.getDay()] + ', ' + String(d.getDate()).padStart(2, '0') + ' ' + BLN[d.getMonth()] + ' ' + d.getFullYear();
  }

  var api = {
    key: key, toNum: toNum, parseDate: parseDate, parseCSV: parseCSV, detectDelim: detectDelim,
    parsePOM: parsePOM, parseMaster: parseMaster, parseStockOut: parseStockOut, parsePromo: parsePromo, parseSeasonal: parseSeasonal,
    compute: compute, buildSummary: buildSummary, COLS: COLS, fmtCell: fmtCell, fmtDate: fmtDate, fmtTodayID: fmtTodayID,
    DEFAULT_STD: DEFAULT_STD
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.PoCore = api;
})(typeof window !== 'undefined' ? window : this);
