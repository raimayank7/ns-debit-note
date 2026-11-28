/* debit-note.niveshsahayak.in: "Interest on late payment from your ledger" (byof3-interest, 28 Nov 2026).
   The visitor drops a party ledger (Tally Ledger Account, Busy ledger, any Excel / CSV with Date, Debit, Credit) or a bills
   outstanding list (Tally Bills Receivable / Ledger Outstandings as Excel or XML, Busy outstanding report) and gets the interest per
   bill and per party, at their own rate or at the MSMED Act s.16 rate (3 x RBI bank rate, monthly rests), payments set against the
   oldest bill first (NSOdInt, intcore.js). For each party: a statement to print or save as PDF, a debit note opened in this site's own
   maker (through the data home's 'handoff', like stock's "Make purchase order"), and a reminder to copy or open in WhatsApp.
   Read in the browser; nothing leaves the device and nothing is saved to the NSBiz books (the debit note is saved only if the
   visitor presses Save in the maker). Kept in this browser: localStorage ns-oi (settings, message), ns-oi-lay (column choices).
   English only for now (owner rule 28 Nov): words are oi.* keys in engine/i18n.js EN.
   Designed & Developed by: Nivesh Sahayak (https://www.niveshsahayak.in/) */
(function () {
  'use strict';
  var NS = window.NS, I = window.NSOdInt; if (!NS || !I) return;
  var $ = NS.$, esc = NS.esc;
  var KEY = 'ns-oi', LAY = 'ns-oi-lay', XLJS = '/assets/vendor/xlsx-0.20.3.full.min.js';
  function W(k, v) { return NS.t('oi.' + k, v); }
  var DEF = { mode: 'msme', bankRate: I.BANK_RATE, agreed: 0, rate: 18, compound: false, cdays: 30, grace: 0, side: 'auto', gst: 18, msg: '', business: '' };
  var S = load(), D = null, root, F = null, colsOpen = false, open = {}, busyMsg = '';
  /* D = { file, kind, entries | lines, parties (NSOdInt shape), skipped, side (used), guessed, sample } */

  function load() { var s = {}; try { s = JSON.parse(localStorage.getItem(KEY) || '{}'); } catch (e) {} return Object.assign({}, DEF, s); }
  function save() { try { localStorage.setItem(KEY, JSON.stringify(S)); } catch (e) {} }
  function layouts() { try { return JSON.parse(localStorage.getItem(LAY) || '{}'); } catch (e) { return {}; } }
  function money(n) { return '₹' + I.inr(n); }
  function upto() { var i = root && $('[data-oi=upto]', root); return (i && i.value) || I.today(); }
  function set() { return { mode: S.mode, bankRate: S.bankRate, agreed: +S.agreed || 0, rate: +S.rate || 0, compound: !!S.compound, cdays: +S.cdays || 0, grace: +S.grace || 0, upto: upto() }; }
  function bizName() { var b = window.NSApp && NSApp.biz && NSApp.biz(); return S.business || (b && b.name) || ''; }

  /* ---------- reading the file ---------- */
  function decode(buf) {
    var b = new Uint8Array(buf);
    if (b[0] === 0xFF && b[1] === 0xFE) return new TextDecoder('utf-16le').decode(buf);
    if (b[0] === 0xFE && b[1] === 0xFF) return new TextDecoder('utf-16be').decode(buf);
    if (b.length > 3 && b[1] === 0 && b[3] === 0) return new TextDecoder('utf-16le').decode(buf);
    return new TextDecoder().decode(buf);
  }
  function sheets(f, buf) {
    var b = new Uint8Array(buf);
    return NS.loadJs(XLJS, '').then(function () {
      var X = window.XLSX, csv = !(b[0] === 0x50 && b[1] === 0x4B) && !(b[0] === 0xD0 && b[1] === 0xCF);
      var wb = csv ? X.read(decode(buf), { type: 'string', raw: true }) : X.read(buf, { type: 'array', cellDates: false });
      return wb.SheetNames.map(function (n) { return { name: n, rows: X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '' }).map(function (r) { return r.map(function (c) { return c == null ? '' : String(c).trim(); }); }) }; });
    }, function () {
      if (b[0] === 0xD0 && b[1] === 0xCF) throw new Error(W('xlsold'));
      return NS.xlsx.read(new File([buf], f.name)).then(function (rows) { return [{ name: f.name, rows: rows }]; });
    });
  }
  /* one sheet: a ledger first (Date + Debit + Credit), else an outstanding list */
  function readSheet(name, sh) {
    var lay = layouts(), r = I.readLedger(sh.rows, { name: name, layouts: lay });
    if (!r.err && r.entries.length) return { kind: r.single ? W('k_ledger1', { p: r.single }) : W('k_ledger'), ledger: r };
    var o = I.readOutstanding(sh.rows, { name: name, layouts: lay, asOf: upto() });
    if (!o.err && o.lines.length) return { kind: W('k_outst'), outst: o };
    return { err: r.err && o.err ? W('noread') : (o.err || r.err), map: r.map, row: r.row };
  }
  function readFile(f) {
    busy(W('reading', { f: f.name }));
    return f.arrayBuffer().then(function (buf) {
      var head = decode(buf.slice(0, 4000)).replace(/^﻿/, '').trim();
      if (/\.xml$/i.test(f.name) || /^<(\?xml|ENVELOPE)/i.test(head)) {
        var X = I.readTallyXml(decode(buf), { asOf: upto() });
        if (X.err) throw new Error(X.err);
        return took(f.name, X.kind || W('k_outst'), null, X, null);
      }
      return sheets(f, buf).then(function (book) {
        var best = null, why = '';
        book.forEach(function (sh) { if (best) return; var x = readSheet(f.name, sh); if (x.err) { why = why || x.err; if (!F) F = { name: f.name, rows: sh.rows, row: x.row || 0 }; return; } best = { sh: sh, x: x }; });
        if (!best) { colsOpen = true; F = F || { name: f.name, rows: (book[0] && book[0].rows) || [], row: 0 }; throw new Error(why || W('noread')); }
        F = { name: f.name, rows: best.sh.rows, row: (best.x.ledger || best.x.outst).row || 0 };
        return took(f.name, best.x.kind + (book.length > 1 ? ' (' + best.sh.name + ')' : ''), best.x.ledger, best.x.outst, null);
      });
    }).then(function () { busy(''); }, function (e) { busy(''); NS.toast(e.message || W('noread'), 'bad'); draw(); });
  }
  function took(name, kind, ledger, outst, sample) {
    D = { file: name, kind: kind, sample: !!sample, ledger: ledger, outst: outst, skipped: (ledger || outst).skipped || [] };
    build(); colsOpen = false; open = {}; draw();
    var t0 = $('[data-oi-step="3"]', root); if (t0) t0.scrollIntoView({ block: 'start', behavior: 'smooth' });
  }
  /* the parties from what was read, for the side chosen */
  function build() {
    if (!D) return;
    if (D.ledger) { D.guessed = I.guessSide(D.ledger.entries); D.side = S.side === 'auto' ? D.guessed : S.side; D.parties = I.fromLedger(D.ledger.entries, D.side); }
    else { D.side = 'recv'; D.parties = I.fromOutstanding(D.outst.lines); }
  }
  function trySample() { var rows = I.sampleRows(upto()), r = I.readLedger(rows, { name: 'sample' }); F = { name: 'sample', rows: rows, row: r.row }; took(W('samplename'), W('k_sample'), r, null, true); }
  function sampleBook() { return [['Ledger', I.sampleRows(I.today()), [12, 30, 14, 16, 14, 14]]]; }

  /* the Columns screen for a sheet the reader could not place */
  var CK = [['date', 'c_date'], ['dr', 'c_dr'], ['cr', 'c_cr'], ['party', 'c_party'], ['vno', 'c_vno'], ['vtype', 'c_vtype'], ['due', 'c_due']];
  function colsHtml() {
    if (!F || !F.rows.length) return '';
    var heads = F.rows[F.row] || [], opts = function (sel) { return '<option value="">' + esc(W('c_none')) + '</option>' + heads.map(function (h, i) { return '<option value="' + i + '"' + (sel === i ? ' selected' : '') + '>' + esc((i + 1) + ': ' + (h || '-')) + '</option>'; }).join(''); };
    return '<div class="oi-cols"><h4>' + esc(W('cols_h')) + '</h4><p class="note">' + esc(W('cols_p')) + '</p><div class="oi-grid"><label>' + esc(W('c_row')) + '<input type="number" min="1" data-oi="headrow" value="' + (F.row + 1) + '"></label>' +
      CK.map(function (c) { return '<label>' + esc(W(c[1])) + '<select data-oicol="' + c[0] + '">' + opts(null) + '</select></label>'; }).join('') + '</div>' +
      '<button type="button" class="btn" data-oi-do="cols">' + esc(W('cols_use')) + '</button></div>';
  }
  function rereadCols() {
    var map = {}; [].forEach.call(root.querySelectorAll('[data-oicol]'), function (s) { if (s.value !== '') map[s.getAttribute('data-oicol')] = +s.value; });
    var row = +(($('[data-oi=headrow]', root) || {}).value || 1) - 1;
    if (!('date' in map) || !('dr' in map) || !('cr' in map)) { NS.toast(W('cols_need'), 'bad'); return; }
    var lay = layouts(); lay[NSTpl.sig(F.rows[row] || [])] = map; try { localStorage.setItem(LAY, JSON.stringify(lay)); } catch (e) {}
    var r = I.readLedger(F.rows, { name: F.name, layouts: lay });
    if (r.err || !r.entries.length) { NS.toast(r.err || W('noread'), 'bad'); return; }
    took(F.name, W('k_yours'), r, null, null);
  }

  /* ---------- outputs ---------- */
  function download(name, blob) { var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000); }
  function xlsxBlob(list) { var z = NS.xlsx.write(list); return z instanceof Blob ? z : new Blob([z], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); }
  function csv(rows) { return '﻿' + rows.map(function (r) { return r.map(function (c) { c = String(c == null ? '' : c); return /[",\n]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c; }).join(','); }).join('\r\n'); }
  function result() { return I.work(D.parties, set()); }
  function partyOf(R, k) { return R.parties.filter(function (p) { return p.k === k; })[0]; }
  /* one party's statement in a print window (Print or Save as PDF there) */
  function statement(ps, R) {
    var st = set(), b = bizName(), rows = ps.rows.map(function (x) { return '<tr><td>' + esc(x.no || '') + '</td><td>' + I.dmy(x.date) + '</td><td class="n">' + I.inr(x.amt) + '</td><td>' + I.dmy(x.due) + '</td><td>' + (x.paid ? I.dmy(x.paid) : x.left > 0 ? 'Unpaid ' + I.inr(x.left) : '') + '</td><td class="n">' + x.days + '</td><td class="n">' + I.inr(x.interest) + '</td></tr>'; }).join('');
    var h = '<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Interest statement - ' + esc(ps.name) + '</title><style>@page{size:A4;margin:14mm}body{font:13px/1.45 system-ui,Arial,sans-serif;color:#111}h1{font-size:18px;margin:0 0 2px}table{width:100%;border-collapse:collapse;margin:10px 0}th,td{border:1px solid #999;padding:4px 6px;text-align:left}th{background:#eee}.n{text-align:right;white-space:nowrap}.tot td{font-weight:700}.ft{margin-top:22px;font-size:11px;color:#555}</style></head><body>' +
      (b ? '<p><b>' + esc(b) + '</b></p>' : '') + '<h1>Statement of interest on late payment</h1><p>To: <b>' + esc(ps.name) + '</b><br>Worked out up to ' + I.dmy(st.upto) + ': ' + esc(I.basis(st)) + '.</p>' +
      '<table><thead><tr><th>Bill no.</th><th>Bill date</th><th class="n">Amount (Rs)</th><th>Last day without interest</th><th>Paid on</th><th class="n">Days late</th><th class="n">Interest (Rs)</th></tr></thead><tbody>' + rows +
      '<tr class="tot"><td colspan="6">Total interest</td><td class="n">' + I.inr(ps.interest) + '</td></tr></tbody></table>' +
      '<p>Payments are set against the oldest bill first. Pending on bills: Rs ' + I.inr(ps.pending) + '.</p><p class="ft">Made with the free interest tool at debit-note.niveshsahayak.in</p>' +
      '<script>window.onload=function(){setTimeout(function(){window.print()},300)}<\/script></body></html>';
    var w = window.open('', '_blank'); if (!w) { NS.toast(W('popup'), 'bad'); return; }
    w.document.open(); w.document.write(h); w.document.close();
  }
  /* the debit note: this site's own maker, filled through the data home's hand-over (docs.js ?hand=1) */
  function debitNote(ps) {
    var st = set(), late = ps.rows.filter(function (x) { return x.interest > 0; }).sort(function (a, b) { return (a.date || '').localeCompare(b.date || ''); });
    if (!(ps.interest > 0)) { NS.toast(W('nointerest'), 'bad'); return; }
    var bills = late.map(function (x) { return (x.no || 'Bill') + ' of ' + I.dmy(x.date) + ' (Rs ' + I.inr(x.amt) + ', ' + x.days + ' days late): Rs ' + I.inr(x.interest); });
    var h = { type: 'note', t: Date.now(), reason: 'other', buyer: { name: ps.name }, orig: late[0] ? { no: late[0].no || '', date: late[0].date || '' } : null,
      items: [{ desc: I.noteDesc(ps, st), hsn: '', qty: 1, unit: 'Nos', rate: ps.interest, gst: S.gst === '' ? 0 : +S.gst }],
      notes: 'Interest on late payment. Bills: ' + bills.join('; ') + '.' };
    if (!window.NSStore) { NS.toast(W('nostore'), 'bad'); return; }
    NSStore.ready().then(function () { return NSStore.meta('handoff', h); }).then(function () { location.assign(location.pathname + '?hand=1#tool'); }, function () { NS.toast(W('nostore'), 'bad'); });
  }
  function reminder(ps) { return I.message(S.msg || I.DEFMSG, ps, set(), { business: bizName() }); }
  function copy(text) { (navigator.clipboard ? navigator.clipboard.writeText(text) : Promise.reject()).then(function () { NS.toast(W('copied')); }, function () { window.prompt(W('copyhand'), text); }); }

  /* ---------- drawing ---------- */
  function busy(m) { busyMsg = m; var b = root && $('[data-oi-busy]', root); if (b) { b.textContent = m; b.hidden = !m; } }
  function radio(name, val, label) { return '<label class="oi-rad"><input type="radio" name="oi-' + name + '" data-oiset="' + name + '" value="' + val + '"' + (String(S[name]) === String(val) ? ' checked' : '') + '> ' + esc(label) + '</label>'; }
  function inp(k, label, type, extra) { return '<label>' + esc(label) + '<input type="' + (type || 'number') + '" step="any" data-oiset="' + k + '" value="' + esc(S[k]) + '"' + (extra || '') + '></label>'; }
  function draw() {
    if (!root) return;
    var u = upto(), st = set(), msme = S.mode === 'msme';
    var h = '<h2>' + esc(W('h')) + '</h2><p>' + esc(W('lead')) + '</p><p class="note">' + esc(W('local')) + '</p>' +
      '<div class="card oi-card" data-oi-step="1"><h3>' + esc(W('s1')) + '</h3>' +
      '<div class="oi-row">' + radio('mode', 'msme', W('m_msme')) + radio('mode', 'own', W('m_own')) + '</div>' +
      '<div class="oi-grid">' +
      (msme ? inp('bankRate', W('bankrate')) + '<label>' + esc(W('agreed')) + '<select data-oiset="agreed">' + [[0, W('ag_none')], [15, '15'], [30, '30'], [45, '45'], [60, W('ag_more')]].map(function (o) { return '<option value="' + o[0] + '"' + (+S.agreed === o[0] ? ' selected' : '') + '>' + esc(o[1]) + '</option>'; }).join('') + '</select></label>'
        : inp('rate', W('rate')) + inp('cdays', W('cdays')) + inp('grace', W('grace')) + '<label class="oi-chk"><input type="checkbox" data-oiset="compound"' + (S.compound ? ' checked' : '') + '> ' + esc(W('compound')) + '</label>') +
      '<label>' + esc(W('upto')) + '<input type="date" data-oi="upto" value="' + esc(u) + '"></label>' +
      '<label>' + esc(W('side')) + '<select data-oiset="side">' + [['auto', W('side_auto')], ['recv', W('side_recv')], ['pay', W('side_pay')]].map(function (o) { return '<option value="' + o[0] + '"' + (S.side === o[0] ? ' selected' : '') + '>' + esc(o[1]) + '</option>'; }).join('') + '</select></label>' +
      '</div>' +
      '<p class="note">' + esc(msme ? W('msme_note', { r: I.BANK_RATE, d: I.BANK_RATE_SEEN, x: (3 * (+S.bankRate || 0)).toFixed(2) }) : W('own_note')) + '</p></div>' +
      '<div class="card oi-card" data-oi-step="2"><h3>' + esc(W('s2')) + '</h3>' +
      '<div class="oi-drop" data-oi-drop><b>' + esc(W('drop')) + ' <label class="oi-pick" for="oi-file">' + esc(W('choose')) + '</label></b><span>' + esc(W('drop_p')) + '</span>' +
      '<input type="file" id="oi-file" class="sr" accept=".xlsx,.xls,.csv,.txt,.xml"></div>' +
      '<div class="oi-row"><button type="button" class="btn" data-oi-do="sample">' + esc(W('try')) + '</button><button type="button" class="btn sm" data-oi-do="sampledl">' + esc(W('sampledl')) + '</button></div>' +
      '<p class="note" data-oi-busy' + (busyMsg ? '' : ' hidden') + '>' + esc(busyMsg) + '</p>' + (colsOpen ? colsHtml() : '') + '</div>';
    if (D) {
      var R = result(), sk = D.skipped.filter(function (s) { return !s.keep; });
      h += '<div class="card oi-card" data-oi-step="3"><h3>' + esc(W('s3')) + '</h3>' +
        '<p><b>' + esc(D.file) + '</b>: ' + esc(D.kind) + '. ' + esc(W('readsum', { p: R.parties.length, b: R.rows.length })) + (D.ledger ? ' ' + esc(D.side === 'pay' ? W('as_pay') : W('as_recv')) + (S.side === 'auto' ? ' ' + esc(W('as_auto')) : '') : ' ' + esc(W('outst_note'))) + (D.sample ? ' <b>' + esc(W('issample')) + '</b>' : '') + '</p>' +
        (sk.length ? '<details class="oi-skip"><summary>' + esc(W('skipped', { n: sk.length })) + '</summary><ul>' + sk.slice(0, 300).map(function (s) { return '<li>' + esc(W('row', { r: s.row })) + ': ' + esc(s.why) + '</li>'; }).join('') + '</ul></details>' : '') +
        (R.nodate ? '<p class="note">' + esc(W('nodate', { n: R.nodate })) + '</p>' : '') +
        '<p class="oi-total">' + esc(W('total')) + ' <b>' + money(R.total) + '</b> <small>(' + esc(I.basis(st)) + ', ' + esc(W('to')) + ' ' + I.dmy(st.upto) + ')</small></p>' +
        '<div class="oi-row"><button type="button" class="btn primary" data-oi-do="xlsx">' + esc(W('dl_xlsx')) + '</button><button type="button" class="btn sm" data-oi-do="csv">' + esc(W('dl_csv')) + '</button>' +
        (D.ledger || colsOpen ? '' : '') + (F && !colsOpen ? '<button type="button" class="btn sm" data-oi-do="showcols">' + esc(W('cols_open')) + '</button>' : '') + '</div>' +
        (colsOpen && D ? colsHtml() : '') +
        '<div class="tbl-wrap"><table class="oi-tbl"><thead><tr><th>' + esc(W('t_party')) + '</th><th class="n">' + esc(W('t_late')) + '</th><th class="n">' + esc(W('t_pending')) + '</th><th class="n">' + esc(W('t_days')) + '</th><th class="n">' + esc(W('t_int')) + '</th><th>' + esc(W('t_do')) + '</th></tr></thead><tbody>' +
        R.parties.map(function (p) {
          var o = open[p.k];
          return '<tr><td>' + esc(p.name) + (p.advance > 0 ? '<br><small>' + esc(W('adv', { a: money(p.advance) })) + '</small>' : '') + '</td><td class="n" data-l="' + esc(W('t_late')) + '">' + p.late + ' / ' + p.bills + '</td><td class="n" data-l="' + esc(W('t_pending')) + '">' + money(p.pending) + '</td><td class="n" data-l="' + esc(W('t_days')) + '">' + p.maxDays + '</td><td class="n" data-l="' + esc(W('t_int')) + '"><b>' + money(p.interest) + '</b></td>' +
            '<td class="oi-acts"><button type="button" class="btn sm" data-oi-p="bills" data-k="' + esc(p.k) + '" aria-expanded="' + (o === 'bills') + '">' + esc(W('a_bills')) + '</button><button type="button" class="btn sm" data-oi-p="stmt" data-k="' + esc(p.k) + '">' + esc(W('a_stmt')) + '</button>' +
            (D.side === 'pay' ? '' : '<button type="button" class="btn sm" data-oi-p="note" data-k="' + esc(p.k) + '">' + esc(W('a_note')) + '</button><button type="button" class="btn sm" data-oi-p="msg" data-k="' + esc(p.k) + '" aria-expanded="' + (o === 'msg') + '">' + esc(W('a_msg')) + '</button>') + '</td></tr>' +
            (o === 'bills' ? '<tr class="oi-open"><td colspan="6">' + billsHtml(p) + '</td></tr>' : '') +
            (o === 'msg' ? '<tr class="oi-open"><td colspan="6">' + msgHtml(p) + '</td></tr>' : '');
        }).join('') + '</tbody></table></div>' +
        '<details class="oi-set"><summary>' + esc(W('more')) + '</summary><div class="oi-grid">' + inp('gst', W('gst')) + inp('business', W('business'), 'text', ' placeholder="' + esc(W('business_ph')) + '"') + '</div>' +
        '<label class="oi-tpl">' + esc(W('tpl')) + '<textarea rows="8" data-oiset="msg">' + esc(S.msg || I.DEFMSG) + '</textarea></label><p class="note">' + esc(W('tpl_p')) + '</p></details>' +
        '<p class="note">' + esc(W(D.side === 'pay' ? 'pay_note' : 'gst_note')) + '</p></div>';
    }
    root.innerHTML = h;
  }
  function billsHtml(p) {
    return '<div class="tbl-wrap"><table class="oi-tbl oi-sub"><thead><tr><th>' + esc(W('b_no')) + '</th><th>' + esc(W('b_date')) + '</th><th class="n">' + esc(W('b_amt')) + '</th><th>' + esc(W('b_free')) + '</th><th>' + esc(W('b_paid')) + '</th><th class="n">' + esc(W('t_days')) + '</th><th class="n">' + esc(W('t_int')) + '</th></tr></thead><tbody>' +
      p.rows.map(function (x) { return '<tr><td>' + esc(x.no || '') + (x.open ? ' <small>(' + esc(W('opening')) + ')</small>' : '') + '</td><td data-l="' + esc(W('b_date')) + '">' + I.dmy(x.date) + '</td><td class="n" data-l="' + esc(W('b_amt')) + '">' + I.inr(x.amt) + '</td><td data-l="' + esc(W('b_free')) + '">' + I.dmy(x.due) + '</td><td data-l="' + esc(W('b_paid')) + '">' + (x.paid ? I.dmy(x.paid) : x.left > 0 ? esc(W('unpaid', { a: I.inr(x.left) })) : '-') + '</td><td class="n" data-l="' + esc(W('t_days')) + '">' + x.days + '</td><td class="n" data-l="' + esc(W('t_int')) + '">' + I.inr(x.interest) + '</td></tr>'; }).join('') + '</tbody></table></div>';
  }
  function msgHtml(p) {
    var m = reminder(p);
    return '<pre class="oi-msg">' + esc(m) + '</pre><div class="oi-row"><button type="button" class="btn sm" data-oi-p="copy" data-k="' + esc(p.k) + '">' + esc(W('copy')) + '</button><a class="btn sm" target="_blank" rel="noopener" href="https://wa.me/?text=' + encodeURIComponent(m) + '">' + esc(W('wa')) + '</a></div>';
  }

  function wire() {
    root.addEventListener('click', function (e) {
      var b = e.target.closest('[data-oi-do],[data-oi-p]'); if (!b) return;
      var op = b.getAttribute('data-oi-do');
      if (op === 'sample') return trySample();
      if (op === 'sampledl') return download('Sample party ledger (made-up).xlsx', xlsxBlob(sampleBook()));
      if (op === 'cols') return rereadCols();
      if (op === 'showcols') { colsOpen = true; return draw(); }
      if (!D) return;
      var R = result(), nm = 'Interest on late payment ' + I.dmy(upto());
      if (op === 'xlsx') return download(nm + '.xlsx', xlsxBlob([['Parties', I.partySheet(R), [32, 8, 12, 14, 14, 14, 10, 14, 14]], ['Bills', I.billSheet(R), [30, 16, 12, 14, 12, 14, 12, 14, 10, 10, 14, 30]]]));
      if (op === 'csv') return download(nm + '.csv', new Blob([csv(I.billSheet(R))], { type: 'text/csv' }));
      var k = b.getAttribute('data-k'), ps = partyOf(R, k), p = b.getAttribute('data-oi-p'); if (!ps) return;
      if (p === 'bills' || p === 'msg') { open[k] = open[k] === p ? '' : p; return draw(); }
      if (p === 'stmt') return statement(ps, R);
      if (p === 'note') return debitNote(ps);
      if (p === 'copy') return copy(reminder(ps));
    });
    root.addEventListener('change', function (e) {
      var el = e.target;
      if (el.id === 'oi-file' && el.files && el.files[0]) { readFile(el.files[0]); el.value = ''; return; }
      var k = el.getAttribute('data-oiset');
      if (k) { S[k] = el.type === 'checkbox' ? el.checked : el.type === 'number' ? (el.value === '' ? '' : +el.value) : el.value; if (k === 'msg' && S.msg === I.DEFMSG) S.msg = ''; save(); if (k === 'side') build(); draw(); return; }
      if (el.getAttribute('data-oi') === 'upto') draw();
    });
    root.addEventListener('dragover', function (e) { var z = e.target.closest('[data-oi-drop]'); if (!z) return; e.preventDefault(); z.classList.add('on'); });
    root.addEventListener('dragleave', function (e) { var z = e.target.closest('[data-oi-drop]'); if (z) z.classList.remove('on'); });
    root.addEventListener('drop', function (e) { var z = e.target.closest('[data-oi-drop]'); if (!z) return; e.preventDefault(); z.classList.remove('on'); var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]; if (f) readFile(f); });
    root.addEventListener('click', function (e) { if (e.target.closest('[data-oi-drop]') && !e.target.closest('label,input,button')) { var i = $('#oi-file', root); if (i) i.click(); } });
  }
  function init() {
    root = document.querySelector('[data-intbulk]');
    if (!root) {   /* translated home pages have no section of their own: add it after the maker */
      var tool = document.getElementById('tool'); if (!tool) return;
      root = document.createElement('section'); root.className = 'intbulk wrap'; root.id = 'interest'; root.setAttribute('data-intbulk', '');
      tool.parentNode.insertBefore(root, tool.nextSibling);
    }
    root.classList.add('intbulk'); draw(); wire();
    document.addEventListener('ns-lang', draw);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
  window.NSIntBulk = { draw: draw, read: readFile, sample: trySample };
})();
