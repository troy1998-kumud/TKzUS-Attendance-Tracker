/**
 * TkzUS Attendance: clock in/out and time off, stored in this Google Sheet.
 * Staff open the web app link and sign in with their Employee ID and PIN.
 * HR opens the same link, taps "HR login" and uses the HR PIN from the Settings tab.
 * All times are India Standard Time. A shift that ends after midnight counts toward the day it started.
 */

var TZ = 'Asia/Kolkata';
var SHEETS = { EMP: 'Employees', ATT: 'Attendance', LEAVE: 'Leave', SET: 'Settings' };
var HEAD = {
  Employees: ['Employee ID', 'Full name', 'Department', 'Job title', 'Shift start', 'Shift end', 'PIN', 'Paid leave days', 'Sick leave days', 'Active'],
  Attendance: ['Work date', 'Employee ID', 'Full name', 'Department', 'Clock in', 'Clock out', 'Hours worked', 'Flags', 'Clock in at', 'Clock out at'],
  Leave: ['Request ID', 'Requested at', 'Employee ID', 'Full name', 'Type', 'From', 'To', 'Working days', 'Note', 'Status', 'Decided at'],
  Settings: ['Setting', 'Value']
};
var DEFAULT_SETTINGS = [
  ['Company', 'TkzUS'],
  ['HR PIN', ''],
  ['Default shift start', '18:00'],
  ['Default shift end', '03:00'],
  ['Allowed difference (min)', '45'],
  ['Minimum hours', '8'],
  ['Paid leave days', '14'],
  ['Sick leave days', '5']
];
var STARTING_STAFF = [
  ['TK0040', 'Chandini Bezawada', 'Recruiting', 'IT Recruiter, Associate', '', ''],
  ['TK0048', 'Nikhil Katta', 'Payroll', 'Payroll Advisor', '16:00', '01:00'],
  ['TK0049', 'Sai Sushank Aidagiri', 'Recruiting', '', '', ''],
  ['TK0057', 'Prasanth Dasari', 'Bench Sales', '', '', ''],
  ['TK0064', 'Srikanth Macha', 'Recruiting', '', '', ''],
  ['TK0067', 'Sri Harsha Vardhan Reddy Vanipenta', 'Recruiting', '', '', ''],
  ['TK0071', 'Tharun Kumar Madaraj', 'Recruiting', '', '', ''],
  ['TK0072', 'Mounica Sai', 'Bench Sales', '', '', ''],
  ['TK0074', 'Divya Lakavath', 'Bench Sales', '', '', ''],
  ['TK0075', 'Kumud Srivastava', 'HR Operations', 'Executive Secretary-HR Ops', '16:00', '01:00']
];

// ---------- setup (run once from the Apps Script editor) ----------
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Object.keys(HEAD).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), HEAD[name].length).setNumberFormat('@');
    if (!sh.getRange(1, 1).getValue()) {
      sh.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  var set = ss.getSheetByName(SHEETS.SET);
  if (set.getLastRow() < 2) {
    var rows = DEFAULT_SETTINGS.map(function (r) { return r[0] === 'HR PIN' ? [r[0], randomPin(6)] : r; });
    set.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  var emp = ss.getSheetByName(SHEETS.EMP);
  if (emp.getLastRow() < 2) {
    var people = STARTING_STAFF.map(function (p) { return p.concat([randomPin(4), '', '', 'Yes']); });
    emp.getRange(2, 1, people.length, HEAD.Employees.length).setValues(people);
  }
  var leave = ss.getSheetByName(SHEETS.LEAVE);
  leave.getRange(2, 10, leave.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Pending', 'Approved', 'Declined', 'Cancelled'], true).build());
  emp.getRange(2, 10, emp.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());
  var blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
}

function randomPin(len) { var s = ''; for (var i = 0; i < len; i++) s += Math.floor(Math.random() * 10); return s; }

// ---------- web app ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('TkzUS Attendance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- sheet helpers ----------
function book() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet(name) { return book().getSheetByName(name); }
function rows(name) {
  var sh = sheet(name), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEAD[name].length).getDisplayValues().map(function (r, i) {
    var o = { _row: i + 2 };
    HEAD[name].forEach(function (h, j) { o[h] = String(r[j]).trim(); });
    return o;
  });
}
function writeRow(name, rowNum, obj) {
  var sh = sheet(name), vals = HEAD[name].map(function (h) { return obj[h] == null ? '' : String(obj[h]); });
  if (rowNum) sh.getRange(rowNum, 1, 1, vals.length).setValues([vals]);
  else { var r = sh.getLastRow() + 1; sh.getRange(r, 1, 1, vals.length).setNumberFormat('@').setValues([vals]); }
}
function settings() {
  var out = {};
  rows(SHEETS.SET).forEach(function (r) { out[r.Setting] = r.Value; });
  return {
    company: out['Company'] || 'TkzUS',
    hrPin: out['HR PIN'] || '',
    shiftStart: out['Default shift start'] || '18:00',
    shiftEnd: out['Default shift end'] || '03:00',
    grace: num(out['Allowed difference (min)'], 45),
    minHours: num(out['Minimum hours'], 8),
    paid: num(out['Paid leave days'], 14),
    sick: num(out['Sick leave days'], 5)
  };
}
function num(v, d) { var n = parseFloat(v); return isNaN(n) ? d : n; }
function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- India time ----------
function nowIST() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function toMin(stamp) { // 'yyyy-MM-dd HH:mm' -> minutes since epoch
  var p = stamp.split(/[- :]/).map(Number);
  return Date.UTC(p[0], p[1] - 1, p[2], p[3] || 0, p[4] || 0) / 60000;
}
function hm(t) { var p = String(t || '').split(':'); return (+p[0] || 0) * 60 + (+p[1] || 0); }
function addDays(ds, n) { var p = ds.split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)).toISOString().slice(0, 10); }
function dow(ds) { var p = ds.split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay(); }
function workDate(stamp) { // before 08:00 counts toward the previous day's shift
  var d = stamp.slice(0, 10);
  return hm(stamp.slice(11)) < 8 * 60 ? addDays(d, -1) : d;
}
function fmtDur(min) { min = Math.max(0, Math.round(min)); return Math.floor(min / 60) + 'h ' + (min % 60 < 10 ? '0' : '') + (min % 60) + 'm'; }
function weekdaysBetween(from, to) {
  var n = 0;
  for (var d = from; d <= to; d = addDays(d, 1)) { var w = dow(d); if (w >= 1 && w <= 5) n++; if (n > 366) break; }
  return n;
}

// ---------- people ----------
function profile(r, st) {
  return {
    id: r['Employee ID'], name: r['Full name'], dept: r['Department'], title: r['Job title'],
    shiftStart: r['Shift start'] || st.shiftStart, shiftEnd: r['Shift end'] || st.shiftEnd,
    paid: num(r['Paid leave days'], st.paid), sick: num(r['Sick leave days'], st.sick),
    active: !/^no$/i.test(r['Active'])
  };
}
function authStaff(id, pin) {
  id = String(id || '').trim().toUpperCase(); pin = String(pin || '').trim();
  if (!id || !pin) throw new Error('Enter your Employee ID and PIN.');
  var st = settings();
  var r = rows(SHEETS.EMP).filter(function (x) { return x['Employee ID'].toUpperCase() === id; })[0];
  if (!r || r['PIN'] !== pin) throw new Error('That Employee ID and PIN do not match. Ask HR if you have forgotten your PIN.');
  var p = profile(r, st);
  if (!p.active) throw new Error('This employee is not active. Please contact HR.');
  return p;
}
function authHR(pin) {
  var st = settings();
  if (!st.hrPin || String(pin || '').trim() !== st.hrPin) throw new Error('That HR PIN is not right.');
  return st;
}

// ---------- attendance rules ----------
function evaluate(rec, p, st) {
  var flags = [], wd = rec['Work date'];
  if (dow(wd) === 0 || dow(wd) === 6) flags.push('Weekend');
  if (!rec['Clock in at']) return { flags: flags, worked: 0 };
  var inM = toMin(rec['Clock in at']);
  var startM = toMin(wd + ' ' + p.shiftStart), endM = toMin(wd + ' ' + p.shiftEnd);
  if (endM <= startM) endM += 1440;
  if (inM - startM > st.grace) flags.push('Late ' + fmtDur(inM - startM));
  if (!rec['Clock out at']) return { flags: flags, worked: 0, open: true, inM: inM };
  var outM = toMin(rec['Clock out at']), worked = outM - inM;
  if (endM - outM > st.grace) flags.push('Left early ' + fmtDur(endM - outM));
  if (worked < st.minHours * 60) flags.push('Short day');
  return { flags: flags, worked: worked };
}
function openShift(id, stamp) {
  var wd = workDate(stamp), prev = addDays(wd, -1);
  var mine = rows(SHEETS.ATT).filter(function (r) {
    return r['Employee ID'].toUpperCase() === id.toUpperCase() && (r['Work date'] === wd || r['Work date'] === prev);
  });
  var open = mine.filter(function (r) { return r['Clock in at'] && !r['Clock out at']; }).pop();
  var today = mine.filter(function (r) { return r['Work date'] === wd; }).pop();
  return { open: open, today: today, wd: wd };
}

// ---------- leave rules ----------
function balances(p, year) {
  var used = { Paid: 0, Sick: 0 }, pending = { Paid: 0, Sick: 0 };
  rows(SHEETS.LEAVE).forEach(function (r) {
    if (r['Employee ID'].toUpperCase() !== p.id.toUpperCase() || r['From'].slice(0, 4) !== String(year)) return;
    var t = /sick/i.test(r['Type']) ? 'Sick' : 'Paid', d = num(r['Working days'], 0);
    if (r['Status'] === 'Approved') used[t] += d;
    else if (r['Status'] === 'Pending') pending[t] += d;
  });
  return {
    year: String(year),
    Paid: { allow: p.paid, used: used.Paid, pending: pending.Paid, left: Math.max(0, p.paid - used.Paid - pending.Paid) },
    Sick: { allow: p.sick, used: used.Sick, pending: pending.Sick, left: Math.max(0, p.sick - used.Sick - pending.Sick) }
  };
}

// ---------- staff API ----------
function staffState(id, pin) {
  var p = authStaff(id, pin), st = settings(), stamp = nowIST(), s = openShift(p.id, stamp);
  var recent = rows(SHEETS.ATT).filter(function (r) { return r['Employee ID'].toUpperCase() === p.id.toUpperCase(); })
    .slice(-10).reverse().map(function (r) { return { date: r['Work date'], in: r['Clock in'], out: r['Clock out'], hours: r['Hours worked'], flags: r['Flags'] }; });
  var requests = rows(SHEETS.LEAVE).filter(function (r) { return r['Employee ID'].toUpperCase() === p.id.toUpperCase(); })
    .slice(-10).reverse().map(function (r) { return { id: r['Request ID'], type: r['Type'], from: r['From'], to: r['To'], days: r['Working days'], note: r['Note'], status: r['Status'] }; });
  var status = 'out', since = '';
  if (s.open) { status = 'in'; since = s.open['Clock in at']; }
  else if (s.today && s.today['Clock out']) { status = 'done'; since = s.today['Clock out at']; }
  return {
    company: st.company, now: stamp, workDate: s.wd,
    me: { id: p.id, name: p.name, dept: p.dept, title: p.title, shiftStart: p.shiftStart, shiftEnd: p.shiftEnd },
    status: status, since: since, today: s.today ? { in: s.today['Clock in'], out: s.today['Clock out'], hours: s.today['Hours worked'], flags: s.today['Flags'] } : null,
    balances: balances(p, s.wd.slice(0, 4)), recent: recent, requests: requests
  };
}

function clockIn(id, pin) {
  return withLock(function () {
    var p = authStaff(id, pin), st = settings(), stamp = nowIST(), s = openShift(p.id, stamp);
    if (s.open) throw new Error('You are already clocked in since ' + s.open['Clock in'] + ' on ' + s.open['Work date'] + '. Clock out first.');
    if (s.today && s.today['Clock out']) throw new Error('You already clocked in and out for this shift. Ask HR if something needs fixing.');
    var rec = { 'Work date': s.wd, 'Employee ID': p.id, 'Full name': p.name, 'Department': p.dept, 'Clock in': stamp.slice(11), 'Clock in at': stamp };
    rec['Flags'] = evaluate(rec, p, st).flags.join(', ');
    writeRow(SHEETS.ATT, null, rec);
    return staffState(id, pin);
  });
}

function clockOut(id, pin) {
  return withLock(function () {
    var p = authStaff(id, pin), st = settings(), stamp = nowIST(), s = openShift(p.id, stamp);
    if (!s.open) throw new Error('You are not clocked in right now.');
    var rec = s.open;
    rec['Clock out'] = stamp.slice(11); rec['Clock out at'] = stamp;
    var ev = evaluate(rec, p, st);
    rec['Hours worked'] = fmtDur(ev.worked); rec['Flags'] = ev.flags.join(', ');
    writeRow(SHEETS.ATT, rec._row, rec);
    return staffState(id, pin);
  });
}

function requestLeave(id, pin, type, from, to, note) {
  return withLock(function () {
    var p = authStaff(id, pin);
    type = /sick/i.test(type) ? 'Sick leave' : 'Paid leave';
    to = to || from;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(from || '') || !/^\d{4}-\d{2}-\d{2}$/.test(to)) throw new Error('Pick the dates.');
    if (to < from) throw new Error('The last day is before the first day.');
    if (from.slice(0, 4) !== to.slice(0, 4)) throw new Error('Leave resets on January 1, so make one request up to December 31 and another from January 1.');
    var days = weekdaysBetween(from, to);
    if (!days) throw new Error('Those dates are a weekend. There are no working days to take off.');
    var b = balances(p, from.slice(0, 4))[type === 'Sick leave' ? 'Sick' : 'Paid'];
    if (days > b.left) throw new Error('Only ' + b.left + ' ' + type.toLowerCase() + ' days are left for ' + from.slice(0, 4) + '.');
    writeRow(SHEETS.LEAVE, null, {
      'Request ID': 'L' + Date.now().toString(36).toUpperCase(), 'Requested at': nowIST(), 'Employee ID': p.id, 'Full name': p.name,
      'Type': type, 'From': from, 'To': to, 'Working days': days, 'Note': String(note || '').slice(0, 300), 'Status': 'Pending'
    });
    return staffState(id, pin);
  });
}

function cancelLeave(id, pin, reqId) {
  return withLock(function () {
    var p = authStaff(id, pin);
    var r = rows(SHEETS.LEAVE).filter(function (x) { return x['Request ID'] === reqId && x['Employee ID'].toUpperCase() === p.id.toUpperCase(); })[0];
    if (!r) throw new Error('That request was not found.');
    if (r['Status'] !== 'Pending') throw new Error('Only pending requests can be cancelled.');
    r['Status'] = 'Cancelled'; r['Decided at'] = nowIST();
    writeRow(SHEETS.LEAVE, r._row, r);
    return staffState(id, pin);
  });
}

// ---------- HR API ----------
function hrState(pin, month) {
  var st = authHR(pin), stamp = nowIST(), wd = workDate(stamp);
  month = /^\d{4}-\d{2}$/.test(month || '') ? month : wd.slice(0, 7);
  var people = rows(SHEETS.EMP).map(function (r) { return profile(r, st); }).filter(function (p) { return p.active && p.id; });
  var att = rows(SHEETS.ATT), leave = rows(SHEETS.LEAVE);
  var onLeave = function (id, d) {
    return leave.filter(function (r) { return r['Employee ID'].toUpperCase() === id.toUpperCase() && r['Status'] === 'Approved' && r['From'] <= d && r['To'] >= d; })[0];
  };
  var weekend = dow(wd) === 0 || dow(wd) === 6;
  var today = people.map(function (p) {
    var rec = att.filter(function (r) { return r['Employee ID'].toUpperCase() === p.id.toUpperCase() && r['Work date'] === wd; }).pop();
    var lv = onLeave(p.id, wd), state;
    if (rec && rec['Clock in'] && !rec['Clock out']) state = 'On shift';
    else if (rec && rec['Clock out']) state = 'Done';
    else if (lv) state = lv['Type'];
    else if (weekend) state = 'Weekend';
    else {
      var startM = toMin(wd + ' ' + p.shiftStart);
      state = toMin(stamp) - startM > st.grace ? 'Not in' : 'Not started';
    }
    return { id: p.id, name: p.name, dept: p.dept, shift: p.shiftStart + '–' + p.shiftEnd, state: state,
      in: rec ? rec['Clock in'] : '', out: rec ? rec['Clock out'] : '', hours: rec ? rec['Hours worked'] : '', flags: rec ? rec['Flags'] : '' };
  });
  var pending = leave.filter(function (r) { return r['Status'] === 'Pending'; }).map(function (r) {
    var p = people.filter(function (x) { return x.id.toUpperCase() === r['Employee ID'].toUpperCase(); })[0];
    var b = p ? balances(p, r['From'].slice(0, 4))[/sick/i.test(r['Type']) ? 'Sick' : 'Paid'] : null;
    return { id: r['Request ID'], name: r['Full name'], type: r['Type'], from: r['From'], to: r['To'], days: r['Working days'], note: r['Note'],
      left: b ? b.allow - b.used : '' };
  });
  var decided = leave.filter(function (r) { return r['Status'] !== 'Pending'; }).slice(-8).reverse().map(function (r) {
    return { name: r['Full name'], type: r['Type'], from: r['From'], to: r['To'], days: r['Working days'], status: r['Status'] };
  });
  var monthRows = att.filter(function (r) { return r['Work date'].slice(0, 7) === month; });
  var summary = people.map(function (p) {
    var mine = monthRows.filter(function (r) { return r['Employee ID'].toUpperCase() === p.id.toUpperCase(); });
    var mins = 0, shortN = 0, lateN = 0, earlyN = 0, days = 0;
    mine.forEach(function (r) {
      if (r['Clock in']) days++;
      var m = /(\d+)h (\d+)m/.exec(r['Hours worked']); if (m) mins += +m[1] * 60 + +m[2];
      if (/Short/.test(r['Flags'])) shortN++;
      if (/Late/.test(r['Flags'])) lateN++;
      if (/early/.test(r['Flags'])) earlyN++;
    });
    var lvDays = leave.filter(function (r) { return r['Employee ID'].toUpperCase() === p.id.toUpperCase() && r['Status'] === 'Approved'; })
      .reduce(function (n, r) { var c = 0; for (var d = r['From']; d <= r['To']; d = addDays(d, 1)) if (d.slice(0, 7) === month && dow(d) >= 1 && dow(d) <= 5) c++; return n + c; }, 0);
    var b = balances(p, month.slice(0, 4));
    return { id: p.id, name: p.name, days: days, hours: fmtDur(mins), short: shortN, late: lateN, early: earlyN, leave: lvDays,
      paidLeft: b.Paid.allow - b.Paid.used, sickLeft: b.Sick.allow - b.Sick.used };
  });
  var log = monthRows.slice().reverse().map(function (r) {
    return { date: r['Work date'], name: r['Full name'], in: r['Clock in'], out: r['Clock out'], hours: r['Hours worked'], flags: r['Flags'] };
  });
  return { company: st.company, now: stamp, workDate: wd, month: month, today: today, pending: pending, decided: decided, summary: summary, log: log, sheetUrl: book().getUrl() };
}

function decideLeave(pin, reqId, decision) {
  return withLock(function () {
    authHR(pin);
    var r = rows(SHEETS.LEAVE).filter(function (x) { return x['Request ID'] === reqId; })[0];
    if (!r) throw new Error('That request was not found.');
    if (r['Status'] !== 'Pending') throw new Error('That request was already ' + r['Status'].toLowerCase() + '.');
    r['Status'] = decision === 'approve' ? 'Approved' : 'Declined'; r['Decided at'] = nowIST();
    writeRow(SHEETS.LEAVE, r._row, r);
    return hrState(pin);
  });
}
