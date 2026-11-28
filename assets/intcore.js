/* NSOdInt: interest on overdue bills from a file the visitor already has (byof3-interest, 28 Nov 2026). Pure, no screen; the
   debit-note site's "Interest on late payment" section (intbulk.js) and scripts/test_intcore.js use it.
   Reads:
   - a party LEDGER (Tally "Ledger Account" Excel export, Busy ledger, any sheet with Date + Debit + Credit; one party, or many parties
     with a Party column or "Ledger: <name>" / "Account : <name>" heading rows between them). Bills = the debit side (customer's ledger)
     or the credit side (supplier's ledger); the other side = payments and credit notes, set against the OLDEST unpaid bill first.
   - a BILLS OUTSTANDING list (Tally Bills Receivable / Ledger Outstandings, Busy outstanding report, any sheet with party + pending
     amount + bill date): read with NSRemind.readRows (remcore.js). Only what is still pending is known there, so interest runs on the
     pending amount from the due date to the chosen day (part payments made earlier are already netted off by the software).
   Interest:
   - own rate: simple interest, rate % a year x days / 365 (or compounded each month if the visitor ticks it), after the credit days.
   - MSMED Act, 2006 s.15-16 (micro and small suppliers): the buyer must pay by the agreed day (never more than 45 days from acceptance)
     or, with no written agreement, within 15 days; after that, compound interest with monthly rests at three times the bank rate notified
     by the Reserve Bank of India. RBI Bank Rate 5.75% (rbi.org.in home page, Current Rates, read 28 Nov 2026); the visitor can change it.
     The bill date stands in for the day of acceptance (the file does not know that day).
   Days: interest runs from the day after the last free day up to the day paid (or the chosen day); days = paid - last free day.
   Monthly rests: every full month from the last free day the interest of that month is added to what earns interest; part of a month =
   that month's interest pro rata by days. Interest stops when the bill itself is fully paid (payments go to the bill, not to interest).
   Designed & Developed by: Nivesh Sahayak (https://www.niveshsahayak.in/) */
(function (G) {
  'use strict';
  var Tpl = G.NSTpl || (typeof require === 'function' ? require('./tplcore.js') : null);
  var Rem = G.NSRemind || (typeof require === 'function' ? require('./remcore.js') : null);
  var BANK_RATE = 5.75, BANK_RATE_SEEN = '28 Nov 2026', BANK_RATE_SRC = 'https://www.rbi.org.in/ (Current Rates: Bank Rate)';
  function t(s) { return String(s == null ? '' : s).trim(); }
  function key(s) { return t(s).toLowerCase().replace(/[^a-z0-9ऀ-෿]+/g, ' ').trim(); }
  function r2(n) { return Math.round((+n + Number.EPSILON) * 100) / 100; }
  function utc(d) { return Date.parse(d + 'T00:00:00Z'); }
  function addDays(d, n) { var x = new Date(utc(d) + n * 864e5); return x.toISOString().slice(0, 10); }
  function gap(a, b) { return Math.round((utc(b) - utc(a)) / 864e5); }
  function addMonths(d, n) {   /* 31 Jan + 1 month = 28/29 Feb (the month's last day), like a calendar month */
    var y = +d.slice(0, 4), m = +d.slice(5, 7) - 1 + n, day = +d.slice(8, 10);
    y += Math.floor(m / 12); m = ((m % 12) + 12) % 12;
    var last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
    return y + '-' + ('0' + (m + 1)).slice(-2) + '-' + ('0' + Math.min(day, last)).slice(-2);
  }
  function today() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }

  /* ---------- reading a ledger ---------- */
  var COLS = [
    { k: 'date', h: 'Date', need: 1, a: ['vch date', 'voucher date', 'txn date', 'transaction date', 'bill date', 'entry date', 'doc date', 'dated'] },
    { k: 'dr', h: 'Debit', need: 1, a: ['dr', 'debit amount', 'debit rs', 'debitrs', 'debit (rs.)', 'debit(rs.)', 'debit amt', 'dr amount', 'dr amt', 'debit inr', 'debit ₹'] },
    { k: 'cr', h: 'Credit', need: 1, a: ['cr', 'credit amount', 'credit rs', 'credit (rs.)', 'credit(rs.)', 'credit amt', 'cr amount', 'cr amt', 'credit inr', 'credit ₹'] },
    { k: 'party', h: 'Party', a: ['party name', 'partys name', "party's name", 'customer', 'customer name', 'supplier', 'supplier name', 'ledger', 'ledger name', 'account name', 'name of party', 'debtor', 'creditor'] },
    { k: 'part', h: 'Particulars', a: ['account', 'details', 'description', 'narration', 'short narration', 'remarks', 'head'] },
    { k: 'vtype', h: 'Vch Type', a: ['voucher type', 'vch type', 'type', 'vchtype', 'entry type', 'doc type', 'transaction type'] },
    { k: 'vno', h: 'Vch No.', a: ['voucher no', 'vch no', 'vch/bill no', 'bill no', 'invoice no', 'ref no', 'reference', 'doc no', 'bill/ref', 'vchbill no', 'invoice number', 'bill number', 'number'] },
    { k: 'due', h: 'Due date', a: ['due on', 'due dt', 'due by', 'payment due'] },
    { k: 'cdays', h: 'Credit days', a: ['credit period', 'credit period days', 'payment terms days', 'terms days', 'cr days'] }
  ];
  var TOTAL = /^(grand\s*)?(sub[\s-]*)?total\b|^closing\s*balance|^(to|by|cr|dr)?\s*closing balance|^carried over|^c\/o$|^b\/f$|^brought forward/i;
  var OPEN = /opening\s*balance|^b\/?f$|balance b\/?f|brought forward/i;
  var HEADPARTY = /^\s*(ledger|account|party|a\/c|customer|supplier)\s*(name)?\s*[:\-]\s*(.+)$/i;

  /* rows of one sheet -> { entries: [{ row, party, date, vno, part, vtype, dr, cr, due, cdays, open }], skipped, map, row, single }
     opts { name (file name), layouts (remembered maps by heading signature), from (period start for an opening line without a date) } */
  function readLedger(rows, opts) {
    opts = opts || {}; rows = rows || [];
    var L = null, s;
    for (var r = 0; r < Math.min(rows.length, 40) && opts.layouts; r++) { s = Tpl.sig(rows[r] || []); if (s && opts.layouts[s] && opts.layouts[s].date != null) { L = { row: r, map: opts.layouts[s], heads: rows[r], remembered: true }; break; } }
    if (!L) L = Tpl.matchCols(rows.slice(0, 40), COLS);
    if (!L || !('date' in L.map) || !('dr' in L.map) || !('cr' in L.map)) return { err: 'No Date, Debit and Credit columns were found, so this is not a ledger.', rows: rows, map: L ? L.map : {}, row: L ? L.row : 0 };
    /* the party above the headings: "Ledger: ABC", or the line before "Ledger Account" (Tally), else the file name */
    var single = '', period = '';
    for (var i = 0; i < L.row; i++) {
      var line = (rows[i] || []).map(t).filter(Boolean).join(' '), m = HEADPARTY.exec(line);
      if (m && !single) single = t(m[3]);
      if (/^ledger account$/i.test(line) && !single) { for (var j = i - 1; j >= 0; j--) { var p = (rows[j] || []).map(t).filter(Boolean).join(' '); if (p) { single = p; break; } } }
      var pm = /(\d{1,2}[-\/. ][A-Za-z0-9]{1,9}[-\/. ]\d{2,4})\s*(to|-)\s*(\d{1,2}[-\/. ][A-Za-z0-9]{1,9}[-\/. ]\d{2,4})/i.exec(line);
      if (pm && !period) period = Tpl.date(pm[1]);
    }
    var hasParty = 'party' in L.map;
    if (!single && !hasParty) single = t(String(opts.name || '').replace(/\.[a-z0-9]+$/i, '').replace(/[_-]+/g, ' ')) || 'Party';
    var out = [], skipped = [], cur = single, from = opts.from || period;
    for (var k = L.row + 1; k < rows.length; k++) {
      var row = rows[k] || [], o = { _r: k + 1 }, any = false;
      Object.keys(L.map).forEach(function (f) { var v = t(row[L.map[f]]); o[f] = v; if (v) any = true; });
      var cells = row.map(t).filter(Boolean);
      if (!any) { if (cells.length === 1) { var hp = HEADPARTY.exec(cells[0]); if (hp) cur = t(hp[3]); } continue; }
      var first = cells[0] || '';
      var hp2 = HEADPARTY.exec(first); if (hp2 && !o.dr && !o.cr) { cur = t(hp2[3]); continue; }
      var party = hasParty ? (t(o.party) || cur) : cur;
      if (hasParty && t(o.party)) cur = party;
      var words = [o.part, o.vtype, first].map(t).join(' ');
      if (TOTAL.test(t(o.part)) || TOTAL.test(first) || TOTAL.test(t(o.party))) { skipped.push({ row: o._r, why: 'total or closing line' }); continue; }
      var dr = Math.abs(Tpl.amt(o.dr).v), cr = Math.abs(Tpl.amt(o.cr).v);
      if (!(dr > 0) && !(cr > 0)) { if (!/\d/.test(o.dr + o.cr)) { if (!o.date) { var hp3 = cells.length === 1 ? cells[0] : ''; if (hp3 && !hasParty && !TOTAL.test(hp3)) cur = hp3.replace(HEADPARTY, '$3'); } continue; } skipped.push({ row: o._r, why: 'amount is zero' }); continue; }
      var isOpen = OPEN.test(words), d = Tpl.date(o.date);
      if (!d && isOpen) d = from || '';
      if (!d) { skipped.push({ row: o._r, why: o.date ? 'date not read: ' + o.date : 'no date' }); continue; }
      if (!party) { skipped.push({ row: o._r, why: 'no party name' }); continue; }
      out.push({ row: o._r, party: party, date: d, vno: t(o.vno), part: t(o.part), vtype: t(o.vtype), dr: r2(dr), cr: r2(cr), due: o.due ? Tpl.date(o.due) : '', cdays: o.cdays ? parseInt(String(o.cdays).replace(/[^\d]/g, ''), 10) || 0 : null, open: isOpen });
    }
    return { entries: out, skipped: skipped, map: L.map, row: L.row, heads: L.heads || rows[L.row], sig: Tpl.sig(rows[L.row] || []), single: hasParty ? '' : single, remembered: !!L.remembered };
  }

  /* customer's ledger (bills on the debit side) or supplier's ledger (bills on the credit side): from the voucher words, else the bigger side */
  function guessSide(entries) {
    var s = 0, p = 0, dr = 0, cr = 0;
    (entries || []).forEach(function (e) { var w = (e.vtype + ' ' + e.part).toLowerCase(); if (/sale|invoice|tax inv/.test(w)) s += e.dr ? 1 : 0; if (/purchase|purc/.test(w)) p += e.cr ? 1 : 0; dr += e.dr; cr += e.cr; });
    if (s !== p) return s > p ? 'recv' : 'pay';
    return dr >= cr ? 'recv' : 'pay';
  }

  /* entries -> parties: { k, name, bills: [{ no, date, amt, due, cdays, row, open }], pays: [{ date, amt, no, row }] } */
  function fromLedger(entries, side) {
    var by = {}, order = [];
    (entries || []).forEach(function (e) {
      var k = key(e.party); if (!k) return;
      var p = by[k]; if (!p) { p = by[k] = { k: k, name: e.party, bills: [], pays: [] }; order.push(k); }
      var bill = side === 'pay' ? e.cr : e.dr, pay = side === 'pay' ? e.dr : e.cr;
      if (bill > 0) p.bills.push({ no: e.vno || (e.open ? 'Opening balance' : ''), date: e.date, amt: bill, due: e.due, cdays: e.cdays, row: e.row, open: e.open, what: e.vtype || e.part });
      if (pay > 0) p.pays.push({ date: e.date, amt: pay, no: e.vno, row: e.row, what: e.vtype || e.part });
    });
    return order.map(function (k) { return by[k]; });
  }
  /* outstanding-list lines (NSRemind.readRows) -> parties with pending bills and no payments */
  function fromOutstanding(lines) {
    var by = {}, order = [];
    (lines || []).forEach(function (l) {
      var k = key(l.party); if (!k) return;
      var p = by[k]; if (!p) { p = by[k] = { k: k, name: l.party, bills: [], pays: [], outstanding: true }; order.push(k); }
      if (l.amt > 0) p.bills.push({ no: l.bill, date: l.date || (l.due ? '' : ''), amt: r2(l.amt), due: l.due, row: l.row, days: l.days });
      else if (l.amt < 0) p.pays.push({ date: '', amt: r2(-l.amt), no: l.bill || 'On account', row: l.row, onAcc: true });
    });
    return order.map(function (k) { return by[k]; });
  }

  /* ---------- the last free day of a bill ---------- */
  /* set { mode 'own'|'msme', cdays, grace, agreed (MSMED written agreement days, 0 = none) } */
  function lastFree(b, set) {
    if (!b.date && b.due) return b.due;
    if (!b.date) return '';
    if (set.mode === 'msme') { var a = Math.floor(+set.agreed || 0); return addDays(b.date, a > 0 ? Math.min(a, 45) : 15); }
    var base = b.due || addDays(b.date, b.cdays != null && b.cdays !== '' && set.useFileDays !== false ? +b.cdays : (+set.cdays || 0));
    return addDays(base, +set.grace || 0);
  }

  /* payments set against the oldest unpaid bill first. A payment dated before any bill waits as an advance and goes to the next bill on
     that bill's date. Returns per bill: pays [{ date, amt }], left (unpaid at the end); and the advance left over. */
  function allocate(party) {
    var bills = party.bills.map(function (b, i) { return Object.assign({ i: i, pays: [], left: b.amt }, b); });
    bills.sort(function (a, b) { return (a.date || '').localeCompare(b.date || '') || a.i - b.i; });
    var pays = party.pays.slice().sort(function (a, b) { return (a.date || '').localeCompare(b.date || ''); });
    var adv = [];   /* { date, amt } not used yet */
    function apply(date, amt, upto) {
      for (var j = 0; j < bills.length && amt > 0.004; j++) {
        var b = bills[j]; if (b.left <= 0.004) continue; if (upto && b.date && b.date > upto) break;
        var use = Math.min(b.left, amt); b.left = r2(b.left - use); amt = r2(amt - use);
        b.pays.push({ date: date && b.date && date < b.date ? b.date : date, amt: r2(use) });
      }
      return amt;
    }
    /* walk the days: bills and payments in date order */
    var ev = [];
    bills.forEach(function (b) { ev.push({ d: b.date || '', k: 0, b: b }); });
    pays.forEach(function (p) { ev.push({ d: p.date || '', k: 1, p: p }); });
    ev.sort(function (a, b) { return a.d.localeCompare(b.d) || a.k - b.k; });
    ev.forEach(function (e) {
      if (e.k === 0) { for (var a = 0; a < adv.length && e.b.left > 0.004; a++) { if (adv[a].amt <= 0.004) continue; var u = Math.min(adv[a].amt, e.b.left); e.b.left = r2(e.b.left - u); adv[a].amt = r2(adv[a].amt - u); e.b.pays.push({ date: e.b.date, amt: r2(u), adv: 1 }); } return; }
      var rest = apply(e.p.date, e.p.amt, e.p.date || '');
      if (rest > 0.004) adv.push({ date: e.p.date, amt: rest });
    });
    return { bills: bills, advance: r2(adv.reduce(function (s, a) { return s + a.amt; }, 0)) };
  }

  /* ---------- interest on one bill ---------- */
  /* b { amt, pays [{ date, amt }] }, start = last free day, upto = the chosen day; set { mode, rate (own % a year), compound, bankRate }
     -> { rate (% a year used), days, interest, slices [{ from, to, days, bal, amt }] } */
  function billInterest(b, start, upto, set) {
    var msme = set.mode === 'msme', yr = msme ? 3 * (set.bankRate === '' || set.bankRate == null ? BANK_RATE : +set.bankRate) : +set.rate || 0;
    var res = { rate: r2(yr), days: 0, interest: 0, slices: [], start: start };
    if (!start || !(yr > 0) || !upto || upto <= start) return res;
    var bal = b.amt, pays = (b.pays || []).filter(function (p) { return p.amt > 0; }).slice().sort(function (x, y) { return (x.date || '').localeCompare(y.date || ''); });
    /* what is already paid by the last free day earns nothing */
    pays.forEach(function (p) { if (!p.date || p.date <= start) bal = r2(bal - p.amt); });
    pays = pays.filter(function (p) { return p.date && p.date > start; });
    if (!(bal > 0.004)) return res;
    var compound = msme || !!set.compound, cap = 0, cur = start, acc = 0, tot = 0, n = 0, pi = 0;
    var end = upto;
    while (cur < end && bal > 0.004) {
      var rest = compound ? addMonths(start, n + 1) : '9999-12-31', perDays = compound ? gap(compound ? addMonths(start, n) : cur, rest) : 365;
      var stop = rest < end ? rest : end;
      while (cur < stop && bal > 0.004) {
        var nx = pi < pays.length && pays[pi].date < stop ? pays[pi].date : stop;
        var nd = gap(cur, nx);
        if (nd > 0) {
          var baseAmt = bal + cap, v = compound ? baseAmt * (yr / 1200) * nd / perDays : baseAmt * yr / 100 * nd / 365;
          res.slices.push({ from: addDays(cur, 1), to: nx, days: nd, bal: r2(baseAmt), amt: r2(v) });
          acc += v; res.days += nd;
        }
        cur = nx;
        while (pi < pays.length && pays[pi].date <= cur) { bal = r2(bal - pays[pi].amt); pi++; }
      }
      if (compound && cur === rest) { cap += acc; tot += acc; acc = 0; n++; }
      if (!compound) break;
    }
    tot += acc;
    res.interest = r2(tot);
    return res;
  }

  /* ---------- the whole file ---------- */
  /* parties (fromLedger / fromOutstanding) + set { mode, rate, compound, bankRate, agreed, cdays, grace, upto, min (smallest interest to show) }
     -> { rows: per bill, parties: per party summary, total, nodate } */
  function work(parties, set) {
    set = set || {}; var upto = set.upto || today(), rows = [], sum = [], total = 0, nodate = 0;
    (parties || []).forEach(function (p) {
      var A = allocate(p), ps = { k: p.k, name: p.name, bills: 0, late: 0, billed: 0, pending: 0, overdue: 0, interest: 0, maxDays: 0, advance: A.advance, outstanding: !!p.outstanding, rows: [] };
      A.bills.forEach(function (b) {
        ps.bills++; ps.billed = r2(ps.billed + b.amt); ps.pending = r2(ps.pending + b.left);
        var lf = lastFree(b, set);
        if (!lf) { nodate++; rows.push({ party: p.name, no: b.no, date: '', amt: b.amt, left: b.left, due: '', days: 0, interest: 0, note: 'no bill date or due date: interest not worked out', row: b.row }); return; }
        var I = billInterest(b, lf, upto, set);
        var paidOn = b.left <= 0.004 && b.pays.length ? b.pays[b.pays.length - 1].date : '';
        var lateDays = paidOn ? Math.max(0, gap(lf, paidOn)) : Math.max(0, gap(lf, upto));
        if (b.left > 0.004 && upto > lf) ps.overdue = r2(ps.overdue + b.left);
        if (I.interest > 0) ps.late++;
        ps.interest = r2(ps.interest + I.interest); ps.maxDays = Math.max(ps.maxDays, lateDays);
        var row = { party: p.name, no: b.no, date: b.date, amt: b.amt, left: b.left, due: lf, paid: paidOn, paidAmt: r2(b.amt - b.left), days: lateDays, idays: I.days, rate: I.rate, interest: I.interest, slices: I.slices, open: b.open, row: b.row };
        rows.push(row); ps.rows.push(row);
      });
      total = r2(total + ps.interest);
      sum.push(ps);
    });
    sum.sort(function (a, b) { return b.interest - a.interest; });
    return { rows: rows, parties: sum, total: total, nodate: nodate, upto: upto, rate: set.mode === 'msme' ? r2(3 * (set.bankRate === '' || set.bankRate == null ? BANK_RATE : +set.bankRate)) : r2(+set.rate || 0) };
  }

  /* ---------- texts ---------- */
  function inr(n) { var v = r2(n), neg = v < 0; v = Math.abs(v); var s = v.toFixed(2), i = s.split('.')[0], f = s.split('.')[1], l3 = i.slice(-3), rest = i.slice(0, -3); if (rest) l3 = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + l3; return (neg ? '-' : '') + l3 + (f === '00' ? '' : '.' + f); }
  function dmy(d) { return d ? d.slice(8, 10) + '-' + d.slice(5, 7) + '-' + d.slice(0, 4) : ''; }
  function basis(set) {
    if (set.mode === 'msme') return 'compound interest with monthly rests at ' + r2(3 * (set.bankRate === '' || set.bankRate == null ? BANK_RATE : +set.bankRate)) + '% a year (three times the RBI bank rate of ' + (set.bankRate === '' || set.bankRate == null ? BANK_RATE : +set.bankRate) + '%) under section 16 of the MSMED Act, 2006';
    return (set.compound ? 'interest compounded monthly' : 'simple interest') + ' at ' + r2(+set.rate || 0) + '% a year';
  }
  /* the debit-note line and the reminder for one party summary */
  function noteDesc(ps, set) { return 'Interest on late payment of ' + ps.late + ' bill' + (ps.late === 1 ? '' : 's') + ' up to ' + dmy(set.upto || today()) + ' (' + basis(set) + ')'; }
  var DEFMSG = 'Dear {name},\n\nAs per our terms, interest is payable on bills paid after the due date. Up to {upto}, interest of Rs {interest} is due on {late} bill(s) ({basis}).\n\n{bills}\n\nPlease pay it with the pending amount of Rs {pending}. A statement is attached.\n\nRegards,\n{business}';
  function billLines(ps, max) {
    var r = ps.rows.filter(function (x) { return x.interest > 0; }).sort(function (a, b) { return b.interest - a.interest; }), out = r.slice(0, max || 8).map(function (x) { return '- ' + (x.no || 'Bill') + ' of ' + dmy(x.date) + ': Rs ' + inr(x.amt) + ', ' + x.days + ' days late, interest Rs ' + inr(x.interest); });
    if (r.length > out.length) out.push('- and ' + (r.length - out.length) + ' more');
    return out.join('\n');
  }
  function message(tpl, ps, set, ctx) {
    ctx = ctx || {};
    var v = { name: ps.name, upto: dmy(set.upto || today()), interest: inr(ps.interest), late: ps.late, basis: basis(set), bills: billLines(ps), pending: inr(ps.pending), business: ctx.business || '', days: ps.maxDays };
    return String(tpl || DEFMSG).replace(/\{(\w+)\}/g, function (m, k) { return v[k] != null ? v[k] : m; }).replace(/\n{3,}/g, '\n\n').trim();
  }

  /* ---------- Excel rows ---------- */
  function billSheet(W) {
    var out = [['Party', 'Bill no.', 'Bill date', 'Bill amount (Rs)', 'Paid (Rs)', 'Last day without interest', 'Paid on', 'Still unpaid (Rs)', 'Days late', 'Rate % a year', 'Interest (Rs)', 'Note']];
    W.rows.forEach(function (x) { out.push([x.party, x.no || '', dmy(x.date), x.amt, x.paidAmt || 0, dmy(x.due), dmy(x.paid), x.left, x.days, x.rate || '', x.interest, x.note || (x.open ? 'opening balance' : '')]); });
    return out;
  }
  function partySheet(W) {
    var out = [['Party', 'Bills', 'Bills with interest', 'Billed (Rs)', 'Still unpaid (Rs)', 'Overdue now (Rs)', 'Most days late', 'Advance not used (Rs)', 'Interest (Rs)']];
    W.parties.forEach(function (p) { out.push([p.name, p.bills, p.late, p.billed, p.pending, p.overdue, p.maxDays, p.advance, p.interest]); });
    out.push(['Total', '', '', '', '', '', '', '', W.total]);
    return out;
  }

  /* ---------- the sample (made-up, clearly marked): a Tally-style ledger of two customers ---------- */
  function sampleRows(upto) {
    upto = upto || today();
    function back(n) { return dmy(addDays(upto, -n)); }
    return [
      ['SAMPLE TRADERS (made-up sample data)'], ['Ludhiana'], [],
      ['Ledger: Gill Hardware (sample)'], ['1-Apr-2026 to ' + dmy(upto)], [],
      ['Date', 'Particulars', 'Vch Type', 'Vch No.', 'Debit', 'Credit'],
      ['', 'Opening Balance', '', '', '15,000.00', ''],
      [back(150), 'Sales', 'Sales', 'SI/26-27/011', '42,000.00', ''],
      [back(120), 'HDFC Bank', 'Receipt', 'RC/014', '', '20,000.00'],
      [back(95), 'Sales', 'Sales', 'SI/26-27/044', '18,500.00', ''],
      [back(60), 'HDFC Bank', 'Receipt', 'RC/031', '', '30,000.00'],
      [back(30), 'Sales', 'Sales', 'SI/26-27/092', '9,800.00', ''],
      [back(10), 'Credit Note', 'Credit Note', 'CN/006', '', '1,200.00'],
      ['', 'Closing Balance', '', '', '', '34,100.00'],
      [],
      ['Ledger: Sandhu Builders (sample)'],
      ['Date', 'Particulars', 'Vch Type', 'Vch No.', 'Debit', 'Credit'],
      [back(200), 'Sales', 'Sales', 'SI/26-27/003', '1,05,500.00', ''],
      [back(140), 'Cash', 'Receipt', 'RC/019', '', '50,000.00'],
      [back(70), 'Sales', 'Sales', 'SI/26-27/058', '24,000.00', ''],
      [back(5), 'SBI', 'Receipt', 'RC/044', '', '55,500.00'],
      ['', 'Grand Total', '', '', '1,29,500.00', '1,05,500.00']
    ];
  }

  var API = { COLS: COLS, BANK_RATE: BANK_RATE, BANK_RATE_SEEN: BANK_RATE_SEEN, BANK_RATE_SRC: BANK_RATE_SRC, readLedger: readLedger, guessSide: guessSide, fromLedger: fromLedger, fromOutstanding: fromOutstanding,
    readOutstanding: function (rows, opts) { return Rem ? Rem.readRows(rows, opts) : { err: 'reader missing' }; }, readTallyXml: function (s, o) { return Rem ? Rem.readTallyXml(s, o) : { err: 'reader missing' }; },
    lastFree: lastFree, allocate: allocate, billInterest: billInterest, work: work, message: message, DEFMSG: DEFMSG, noteDesc: noteDesc, basis: basis, billSheet: billSheet, partySheet: partySheet,
    sampleRows: sampleRows, inr: inr, dmy: dmy, addDays: addDays, addMonths: addMonths, gap: gap, today: today, key: key };
  if (typeof module !== 'undefined' && module.exports) module.exports = API; else G.NSOdInt = API;
})(typeof window !== 'undefined' ? window : this);
