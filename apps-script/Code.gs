/**
 * Staff Attendance — backend (Google Apps Script web app).
 *
 * Data lives in the spreadsheet this script is bound to; selfies go to a Drive folder.
 * The phone app (GitHub Pages) talks to this script by POSTing JSON as text/plain
 * (Apps Script cannot answer a CORS preflight, so never send application/json).
 *
 * First time: run setup() from the editor once and allow the permissions.
 */

const SH = { EMP: 'Employees', PUNCH: 'Punches', LOC: 'Locations', SET: 'Settings', HOL: 'Holidays' };

const EMP_COLS = ['EmpID', 'Name', 'PIN', 'Phone', 'Department', 'ShiftStart', 'ShiftEnd',
  'LunchMinutes', 'GraceMinutes', 'WeeklyOff', 'AllowedLocations', 'Active', 'JoinDate'];
const PUNCH_COLS = ['Timestamp', 'WorkDate', 'EmpID', 'Name', 'Action', 'Latitude', 'Longitude',
  'AccuracyM', 'Address', 'Location', 'InsideFence', 'DistanceM', 'PhotoURL', 'PhotoID',
  'ThumbID', 'DeviceTime', 'Device'];
const LOC_COLS = ['Name', 'Latitude', 'Longitude', 'RadiusM'];
const HOL_COLS = ['Date', 'Name'];

const DEFAULT_SETTINGS = [
  ['COMPANY_NAME', 'My Company', 'Shown in the app and on every selfie stamp'],
  ['ADMIN_PASSWORD', '', 'Password for the admin dashboard'],
  ['ENFORCE_GEOFENCE', 'FALSE', 'TRUE = block punches made outside the radius of the Locations tab'],
  ['MAX_ACCURACY_M', '200', 'Reject GPS fixes less accurate than this many metres (0 = accept any)'],
  ['PHOTO_FOLDER_ID', '', 'Drive folder for selfies (created by setup)'],
];

const LABEL = { IN: 'Check In', LUNCH_OUT: 'Lunch Out', LUNCH_IN: 'Lunch In', OUT: 'Check Out' };
const NEXT = { '': ['IN'], IN: ['LUNCH_OUT', 'OUT'], LUNCH_OUT: ['LUNCH_IN'], LUNCH_IN: ['LUNCH_OUT', 'OUT'], OUT: [] };

/* ------------------------------------------------------------------ setup */

function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SH.EMP, EMP_COLS, ['EmpID', 'PIN', 'Phone', 'ShiftStart', 'ShiftEnd', 'WeeklyOff', 'AllowedLocations', 'JoinDate']);
  ensureSheet_(ss, SH.PUNCH, PUNCH_COLS, ['WorkDate', 'EmpID']);
  ensureSheet_(ss, SH.LOC, LOC_COLS, ['Name']);
  ensureSheet_(ss, SH.HOL, HOL_COLS, ['Date']);
  ensureSheet_(ss, SH.SET, ['Key', 'Value', 'Notes'], ['Key', 'Value']);

  const have = settings_();
  DEFAULT_SETTINGS.forEach(function (d) {
    if (!(d[0] in have)) sheet_(SH.SET).appendRow([d[0], txt_(d[1]), d[2]]);
  });
  const s = settings_();
  if (!s.ADMIN_PASSWORD) setSetting_('ADMIN_PASSWORD', 'admin' + Math.floor(100000 + Math.random() * 900000));
  if (!s.PHOTO_FOLDER_ID) setSetting_('PHOTO_FOLDER_ID', DriveApp.createFolder(ss.getName() + ' - Selfies').getId());
  secret_();

  const emp = sheet_(SH.EMP);
  if (emp.getLastRow() < 2) {
    emp.appendRow([txt_('EMP001'), 'Sample Employee', txt_('1234'), '', 'Office', txt_('09:00'), txt_('18:00'), 45, 10, txt_('Sun'), '', 'Yes', txt_(fmt_(new Date(), 'yyyy-MM-dd'))]);
  }
  ['Sheet1', 'Sheet 1'].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh && ss.getSheets().length > 1 && sh.getLastRow() === 0) ss.deleteSheet(sh);
  });
  try { Maps.newGeocoder().reverseGeocode(19.076, 72.8777); } catch (e) { /* only here to request the permission */ }

  const msg = 'Setup complete.\n\nAdmin password: ' + settings_().ADMIN_PASSWORD +
    '\n(Change it in the Settings tab or from the dashboard.)\n\nNow deploy: Deploy > New deployment > Web app,' +
    ' Execute as: Me, Who has access: Anyone.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* run from the editor: see the log */ }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Attendance').addItem('Run setup', 'setup').addToUi();
}

function ensureSheet_(ss, name, cols, textCols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#e6f4f1');
    sh.setFrozenRows(1);
  } else {
    // add any columns that a newer version introduced
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    cols.forEach(function (c) {
      if (head.indexOf(c) < 0) {
        sh.getRange(1, head.length + 1).setValue(c).setFontWeight('bold').setBackground('#e6f4f1');
        head.push(c);
      }
    });
  }
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  textCols.forEach(function (c) {
    const i = head.indexOf(c);
    if (i >= 0) sh.getRange(2, i + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  });
  return sh;
}

/* ------------------------------------------------------------------ web entry points */

const ROUTES = {
  login: login_,
  status: status_,
  geocode: geocode_,
  punch: punch_,
  adminLogin: adminLogin_,
  adminDay: adminDay_,
  adminReport: adminReport_,
  adminEmployees: adminEmployees_,
  adminSaveEmployee: adminSaveEmployee_,
  adminLocations: adminLocations_,
  adminSaveLocation: adminSaveLocation_,
  adminDeleteLocation: adminDeleteLocation_,
  adminSaveSettings: adminSaveSettings_,
  adminThumbs: adminThumbs_,
  adminPhoto: adminPhoto_,
};

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = ROUTES[req.action];
    if (!fn) throw new Error('Unknown action');
    out = Object.assign({ ok: true }, fn(req));
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'attendance', time: Date.now() }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ employee API */

function login_(req) {
  const id = String(req.empId || '').trim();
  const pin = String(req.pin || '').trim();
  if (!id || !pin) throw new Error('Enter your Employee ID and PIN.');
  throttle_('emp:' + id.toUpperCase());
  const e = findEmployee_(id);
  if (!e || e.pin !== pin) {
    failed_('emp:' + id.toUpperCase());
    throw new Error('Wrong Employee ID or PIN.');
  }
  if (!e.active) throw new Error('Your account is inactive. Contact the office.');
  const token = sign_({ r: 'emp', id: e.id, pv: hash_(e.pin), exp: Date.now() + 180 * 864e5 });
  return Object.assign({ token: token }, statusFor_(e));
}

function status_(req) {
  return statusFor_(authEmp_(req.token));
}

function geocode_(req) {
  const e = authEmp_(req.token);
  const lat = Number(req.lat), lng = Number(req.lng);
  if (!isFinite(lat) || !isFinite(lng)) throw new Error('Bad location');
  const f = fence_(e, lat, lng);
  return { address: address_(lat, lng), place: f.name || '', inside: f.inside, distance: f.distance, fenced: f.has, enforce: enforceFence_() };
}

function punch_(req) {
  const e = authEmp_(req.token);
  const type = String(req.type || '');
  if (!LABEL[type]) throw new Error('Unknown punch type');

  const m = /^data:image\/jpeg;base64,(.+)$/.exec(String(req.photo || ''));
  if (!m) throw new Error('A selfie is required.');
  if (m[1].length > 8e6) throw new Error('Photo is too large.');
  const tm = /^data:image\/jpeg;base64,(.+)$/.exec(String(req.thumb || ''));

  const lat = Number(req.lat), lng = Number(req.lng), acc = Math.round(Number(req.acc) || 0);
  if (!isFinite(lat) || !isFinite(lng) || (lat === 0 && lng === 0)) {
    throw new Error('Location is required. Turn on GPS and allow location access.');
  }
  const s = settings_();
  const maxAcc = num_(s.MAX_ACCURACY_M, 0);
  if (maxAcc && acc > maxAcc) {
    throw new Error('GPS is not accurate enough (±' + acc + ' m). Step near a window or outside and try again.');
  }
  const f = fence_(e, lat, lng);
  if (enforceFence_() && f.has && !f.inside) {
    throw new Error('You are ' + f.distance + ' m away from ' + f.name + '. Punch from the work location.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const now = new Date();
    const st = empState_(e, now);
    if (st.next.indexOf(type) < 0) {
      throw new Error(st.next.length
        ? 'You cannot ' + LABEL[type] + ' now. Next: ' + st.next.map(function (t) { return LABEL[t]; }).join(' or ') + '.'
        : 'You have already checked out for this shift.');
    }
    const stamp = fmt_(now, 'yyyy-MM-dd_HHmmss');
    const folder = monthFolder_(now);
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/jpeg',
      e.id + '_' + stamp + '_' + type + '.jpg'));
    file.setDescription(e.name + ' - ' + LABEL[type] + ' - ' + fmt_(now, 'dd MMM yyyy HH:mm:ss') + ' - ' + lat + ',' + lng);
    let thumbId = '';
    if (tm) {
      thumbId = thumbFolder_(folder).createFile(Utilities.newBlob(Utilities.base64Decode(tm[1]), 'image/jpeg',
        e.id + '_' + stamp + '_' + type + '_t.jpg')).getId();
    }
    const address = String(req.address || '').slice(0, 300) || address_(lat, lng);
    sheet_(SH.PUNCH).appendRow([
      now, txt_(st.workDate), txt_(e.id), txt_(e.name), LABEL[type], lat, lng, acc, txt_(address),
      txt_(f.name), f.has ? (f.inside ? 'Yes' : 'No') : '', f.has ? f.distance : '',
      file.getUrl(), file.getId(), thumbId, txt_(String(req.deviceTime || '').slice(0, 40)), txt_(String(req.device || '').slice(0, 150)),
    ]);
  } finally {
    lock.releaseLock();
  }
  return statusFor_(e);
}

function statusFor_(e) {
  const now = new Date();
  const st = empState_(e, now);
  return {
    company: settings_().COMPANY_NAME || '',
    tz: tz_(),
    serverTime: now.getTime(),
    profile: {
      id: e.id, name: e.name, dept: e.dept, shiftStart: e.shiftStart, shiftEnd: e.shiftEnd,
      lunch: e.lunch, grace: e.grace, weeklyOff: e.weeklyOff,
    },
    workDate: st.workDate,
    punches: st.punches.map(publicPunch_),
    next: st.next,
    summary: st.summary,
  };
}

function empState_(e, now) {
  const wd = workDate_(e, now);
  const list = punches_(function (p) { return p.empId === e.id && p.date === wd; });
  const last = list.length ? list[list.length - 1].action : '';
  return { workDate: wd, punches: list, next: NEXT[last] || [], summary: summarize_(e, wd, list, now, holidays_()) };
}

/* ------------------------------------------------------------------ admin API */

function adminLogin_(req) {
  throttle_('admin');
  const pw = String(settings_().ADMIN_PASSWORD || '');
  if (!pw || String(req.password || '') !== pw) {
    failed_('admin');
    throw new Error('Wrong password.');
  }
  return { token: adminToken_(pw), company: settings_().COMPANY_NAME || '', tz: tz_() };
}

function adminDay_(req) {
  authAdmin_(req.token);
  const date = dateParam_(req.date) || fmt_(new Date(), 'yyyy-MM-dd');
  const rep = buildReport_(date, date, '');
  const punches = punches_(function (p) { return p.date === date; }).map(publicPunch_);
  return { date: date, rows: rep.days, punches: punches, settings: publicSettings_(), serverTime: Date.now() };
}

function adminReport_(req) {
  authAdmin_(req.token);
  const from = dateParam_(req.from), to = dateParam_(req.to);
  if (!from || !to || from > to) throw new Error('Choose a valid date range.');
  return buildReport_(from, to, String(req.empId || ''));
}

function adminEmployees_(req) {
  authAdmin_(req.token);
  return { employees: employees_().map(function (e) { delete e.row; return e; }), locations: locations_() };
}

function adminSaveEmployee_(req) {
  authAdmin_(req.token);
  const x = req.employee || {};
  const id = String(x.id || '').trim();
  if (!/^[A-Za-z0-9_-]{1,20}$/.test(id)) throw new Error('Employee ID: letters, numbers, - or _ (max 20).');
  const name = String(x.name || '').trim();
  if (!name) throw new Error('Name is required.');
  const pin = String(x.pin || '').trim();
  if (!/^\d{4,8}$/.test(pin)) throw new Error('PIN must be 4 to 8 digits.');
  const ss = hhmm_(x.shiftStart), se = hhmm_(x.shiftEnd);
  if (!ss || !se) throw new Error('Shift start and end must be times like 09:30.');
  if (ss === se) throw new Error('Shift start and end cannot be the same.');
  const lunch = Math.round(num_(x.lunch, 0)), grace = Math.round(num_(x.grace, 0));
  if (lunch < 0 || lunch > 300 || grace < 0 || grace > 180) throw new Error('Check the lunch / grace minutes.');

  const join = x.joinDate ? dateParam_(x.joinDate) : '';
  if (x.joinDate && !join) throw new Error('Joining date is invalid.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SH.EMP);
    const col = headerMap_(sh);
    const existing = findEmployee_(id);
    if (x.isNew && existing) throw new Error('Employee ID ' + id + ' already exists.');
    const vals = {
      EmpID: txt_(existing ? existing.id : id), Name: txt_(name), PIN: txt_(pin), Phone: txt_(String(x.phone || '').trim()),
      Department: txt_(String(x.dept || '').trim()), ShiftStart: txt_(ss), ShiftEnd: txt_(se), LunchMinutes: lunch,
      GraceMinutes: grace, WeeklyOff: txt_(String(x.weeklyOff || '').trim()),
      AllowedLocations: txt_(String(x.locations || '').trim()), Active: x.active === false ? 'No' : 'Yes',
      JoinDate: txt_(join || (existing ? existing.joinDate : '') || fmt_(new Date(), 'yyyy-MM-dd')),
    };
    const width = sh.getLastColumn();
    const row = existing ? sh.getRange(existing.row, 1, 1, width).getValues()[0] : new Array(width).fill('');
    Object.keys(vals).forEach(function (k) { if (col[k] !== undefined) row[col[k]] = vals[k]; });
    if (existing) sh.getRange(existing.row, 1, 1, width).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return adminEmployees_(req);
}

function adminLocations_(req) {
  authAdmin_(req.token);
  return { locations: locations_() };
}

function adminSaveLocation_(req) {
  authAdmin_(req.token);
  const x = req.location || {};
  const name = String(x.name || '').trim();
  const lat = Number(x.lat), lng = Number(x.lng), radius = Math.round(num_(x.radius, 100));
  if (!name) throw new Error('Location name is required.');
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) throw new Error('Invalid coordinates.');
  if (radius < 10 || radius > 50000) throw new Error('Radius must be between 10 and 50000 metres.');
  const sh = sheet_(SH.LOC);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim().toLowerCase() === name.toLowerCase()) {
      sh.getRange(i + 1, 1, 1, 4).setValues([[txt_(name), lat, lng, radius]]);
      return adminLocations_(req);
    }
  }
  sh.appendRow([txt_(name), lat, lng, radius]);
  return adminLocations_(req);
}

function adminDeleteLocation_(req) {
  authAdmin_(req.token);
  const name = String(req.name || '').trim().toLowerCase();
  const sh = sheet_(SH.LOC);
  const v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]).trim().toLowerCase() === name) sh.deleteRow(i + 1);
  }
  return adminLocations_(req);
}

function adminSaveSettings_(req) {
  authAdmin_(req.token);
  const x = req.settings || {};
  if (x.company !== undefined) setSetting_('COMPANY_NAME', String(x.company).trim().slice(0, 60));
  if (x.enforceGeofence !== undefined) setSetting_('ENFORCE_GEOFENCE', x.enforceGeofence ? 'TRUE' : 'FALSE');
  if (x.maxAccuracy !== undefined) setSetting_('MAX_ACCURACY_M', String(Math.max(0, Math.round(num_(x.maxAccuracy, 0)))));
  const out = { settings: publicSettings_() };
  if (x.newPassword) {
    const pw = String(x.newPassword);
    if (pw.length < 6) throw new Error('New password must be at least 6 characters.');
    setSetting_('ADMIN_PASSWORD', pw);
    out.token = adminToken_(pw);
  }
  return out;
}

function adminThumbs_(req) {
  authAdmin_(req.token);
  const out = {};
  (req.ids || []).slice(0, 60).forEach(function (id) {
    id = String(id);
    if (!/^[\w-]{10,}$/.test(id)) return;
    try {
      const b = DriveApp.getFileById(id).getBlob();
      out[id] = 'data:image/jpeg;base64,' + Utilities.base64Encode(b.getBytes());
    } catch (e) { out[id] = null; }
  });
  return { thumbs: out };
}

function adminPhoto_(req) {
  authAdmin_(req.token);
  const id = String(req.id || '');
  if (!/^[\w-]{10,}$/.test(id)) throw new Error('Bad photo id');
  const b = DriveApp.getFileById(id).getBlob();
  return { photo: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) };
}

function buildReport_(from, to, empId) {
  const now = new Date();
  const dates = [];
  for (let d = from; d <= to; d = addDays_(d, 1)) {
    dates.push(d);
    if (dates.length > 93) throw new Error('Choose at most 3 months at a time.');
  }
  const hol = holidays_();
  const all = punches_(function (p) { return p.date >= from && p.date <= to && (!empId || p.empId === empId); });
  const byKey = {};
  all.forEach(function (p) { (byKey[p.empId + '|' + p.date] = byKey[p.empId + '|' + p.date] || []).push(p); });
  const emps = employees_().filter(function (e) {
    if (empId && e.id !== empId) return false;
    if (e.joinDate && e.joinDate > to) return false;
    return e.active || all.some(function (p) { return p.empId === e.id; });
  });

  const days = [], totals = [];
  emps.forEach(function (e) {
    const t = { id: e.id, name: e.name, dept: e.dept, shift: e.shiftStart + '-' + e.shiftEnd, present: 0, halfDay: 0,
      absent: 0, late: 0, lateMin: 0, earlyMin: 0, workedMin: 0, otMin: 0, lunchOverMin: 0, off: 0, missedOut: 0 };
    dates.forEach(function (d) {
      if (e.joinDate && d < e.joinDate) return; // not employed yet
      const s = summarize_(e, d, byKey[e.id + '|' + d] || [], now, hol);
      s.id = e.id; s.name = e.name; s.dept = e.dept;
      days.push(s);
      if (s.status === 'Present') t.present++;
      if (s.status === 'Half Day') t.halfDay++;
      if (s.status === 'Absent') t.absent++;
      if (s.status === 'Weekly Off' || s.status === 'Holiday') t.off++;
      if (s.status === 'No Check-out') t.missedOut++;
      if (s.late) { t.late++; t.lateMin += s.late; }
      t.earlyMin += s.early; t.workedMin += s.workedMin; t.otMin += s.ot; t.lunchOverMin += s.lunchOver;
    });
    totals.push(t);
  });
  return { from: from, to: to, days: days, totals: totals };
}

/* ------------------------------------------------------------------ attendance logic */

/** Shift in minutes from midnight. Night shifts (end before start) cross midnight. */
function shift_(e) {
  const s = toMin_(e.shiftStart), en = toMin_(e.shiftEnd);
  let len = (en - s + 1440) % 1440;
  if (len === 0) len = 1440;
  // The "day" of a shift starts halfway through the off-hours before it, so a late
  // check-out after midnight or an early check-in still lands on the right shift.
  return { start: s, end: en, len: len, windowStart: s - (1440 - len) / 2 };
}

function workDate_(e, t) {
  return fmt_(new Date(t.getTime() - shift_(e).windowStart * 60000), 'yyyy-MM-dd');
}

function summarize_(e, date, list, now, hol) {
  const sh = shift_(e);
  const schedStart = dateAt_(date, sh.start);
  const schedEnd = new Date(schedStart.getTime() + sh.len * 60000);
  const current = workDate_(e, now) === date;
  const holiday = hol[date] || '';
  const off = !!holiday || isWeeklyOff_(e, date);

  list = list.slice().sort(function (a, b) { return a.ts - b.ts; });
  const first = list.filter(function (p) { return p.action === 'IN'; })[0] || null;
  const outs = list.filter(function (p) { return p.action === 'OUT'; });
  const lastOut = outs.length ? outs[outs.length - 1] : null;
  let breakMs = 0, bo = null, lunchOut = null, lunchIn = null;
  list.forEach(function (p) {
    if (p.action === 'LUNCH_OUT') { bo = p.ts; if (!lunchOut) lunchOut = p.ts; }
    else if (p.action === 'LUNCH_IN' && bo) { breakMs += p.ts - bo; bo = null; lunchIn = p.ts; }
  });
  const onBreak = !!bo && !lastOut;
  if (bo && current && !lastOut) breakMs += now - bo;

  const endT = lastOut ? lastOut.ts : (current && first ? now : null);
  const workedMin = first && endT ? Math.max(0, Math.round((endT - first.ts - breakMs) / 60000)) : 0;
  const expected = Math.max(0, sh.len - num_(e.lunch, 0));
  const lateMin = first ? Math.max(0, Math.round((first.ts - schedStart) / 60000)) : 0;
  const late = first && !off && lateMin > num_(e.grace, 0) ? lateMin : 0;
  const early = lastOut && !off ? Math.max(0, Math.round((schedEnd - lastOut.ts) / 60000)) : 0;
  const breakMin = Math.round(breakMs / 60000);
  const ot = lastOut ? (off ? workedMin : Math.max(0, workedMin - expected)) : 0;

  let status;
  if (!first) {
    if (off) status = holiday ? 'Holiday' : 'Weekly Off';
    else if (now < schedStart) status = 'Upcoming';
    else if (now < schedEnd) status = 'Not In Yet';
    else status = 'Absent';
  } else if (!lastOut) {
    status = current ? (onBreak ? 'On Lunch' : 'Working') : 'No Check-out';
  } else {
    status = !off && workedMin < expected / 2 ? 'Half Day' : 'Present';
  }

  return {
    date: date, status: status, offDay: off, holiday: holiday,
    shift: e.shiftStart + '-' + e.shiftEnd,
    in: first ? first.ts.getTime() : null,
    out: lastOut ? lastOut.ts.getTime() : null,
    lunchOut: lunchOut ? lunchOut.getTime() : null,
    lunchIn: lunchIn ? lunchIn.getTime() : null,
    workedMin: workedMin, expectedMin: expected, late: late, early: early,
    breakMin: breakMin, lunchOver: Math.max(0, breakMin - num_(e.lunch, 0)), ot: ot,
  };
}

function isWeeklyOff_(e, date) {
  const days = String(e.weeklyOff || '').toLowerCase().split(/[\s,;\/]+/).filter(String).map(function (d) { return d.slice(0, 3); });
  if (!days.length) return false;
  return days.indexOf(fmt_(dateAt_(date, 720), 'EEE').toLowerCase()) >= 0;
}

function fence_(e, lat, lng) {
  let locs = locations_();
  const allowed = String(e.locations || '').toLowerCase().split(',').map(function (s) { return s.trim(); }).filter(String);
  if (allowed.length) locs = locs.filter(function (l) { return allowed.indexOf(l.name.toLowerCase()) >= 0; });
  if (!locs.length) return { has: false, inside: true, distance: 0, name: '' };
  let best = null;
  locs.forEach(function (l) {
    const d = Math.round(distance_(lat, lng, l.lat, l.lng));
    if (!best || d - l.radius < best.distance - best.radius) best = { name: l.name, distance: d, radius: l.radius };
  });
  return { has: true, inside: best.distance <= best.radius, distance: best.distance, name: best.name };
}

function distance_(a, b, c, d) {
  const R = 6371000, r = Math.PI / 180;
  const x = Math.sin((c - a) * r / 2), y = Math.sin((d - b) * r / 2);
  return 2 * R * Math.asin(Math.sqrt(x * x + Math.cos(a * r) * Math.cos(c * r) * y * y));
}

function address_(lat, lng) {
  const key = 'geo:' + lat.toFixed(4) + ',' + lng.toFixed(4);
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit !== null) return hit;
  let a = '';
  try {
    const r = Maps.newGeocoder().reverseGeocode(lat, lng);
    a = (r && r.results && r.results[0] && r.results[0].formatted_address) || '';
  } catch (err) { /* quota or network: the stamp falls back to coordinates */ }
  if (a) cache.put(key, a, 21600);
  return a;
}

/* ------------------------------------------------------------------ data access */

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" is missing. Run setup() in the script editor.');
  return sh;
}

function headerMap_(sh) {
  const m = {};
  sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].forEach(function (h, i) { m[String(h).trim()] = i; });
  return m;
}

function employees_() {
  const sh = sheet_(SH.EMP);
  const v = sh.getDataRange().getValues();
  const c = headerMap_(sh);
  const g = function (r, k) { return c[k] === undefined ? '' : r[c[k]]; };
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    const id = String(g(r, 'EmpID')).trim();
    if (!id) continue;
    out.push({
      row: i + 1, id: id, name: String(g(r, 'Name')).trim(), pin: String(g(r, 'PIN')).trim(),
      phone: String(g(r, 'Phone')).trim(), dept: String(g(r, 'Department')).trim(),
      shiftStart: hhmm_(g(r, 'ShiftStart')) || '09:00', shiftEnd: hhmm_(g(r, 'ShiftEnd')) || '18:00',
      lunch: num_(g(r, 'LunchMinutes'), 0), grace: num_(g(r, 'GraceMinutes'), 0),
      weeklyOff: String(g(r, 'WeeklyOff')).trim(), locations: String(g(r, 'AllowedLocations')).trim(),
      active: !/^(no|n|false|0|inactive|left)$/i.test(String(g(r, 'Active')).trim()),
      joinDate: dstr_(g(r, 'JoinDate')),
    });
  }
  return out;
}

function findEmployee_(id) {
  id = String(id).trim().toUpperCase();
  return employees_().filter(function (e) { return e.id.toUpperCase() === id; })[0] || null;
}

function punches_(keep) {
  const sh = sheet_(SH.PUNCH);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const c = headerMap_(sh);
  const v = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const g = function (r, k) { return c[k] === undefined ? '' : r[c[k]]; };
  const out = [];
  v.forEach(function (r) {
    const ts = g(r, 'Timestamp');
    if (!ts) return;
    const p = {
      ts: ts instanceof Date ? ts : new Date(ts), date: dstr_(g(r, 'WorkDate')), empId: String(g(r, 'EmpID')).trim(),
      name: String(g(r, 'Name')), action: actionCode_(g(r, 'Action')), lat: g(r, 'Latitude'), lng: g(r, 'Longitude'),
      acc: g(r, 'AccuracyM'), address: String(g(r, 'Address')), location: String(g(r, 'Location')),
      inside: String(g(r, 'InsideFence')), distance: g(r, 'DistanceM'), photoId: String(g(r, 'PhotoID')),
      thumbId: String(g(r, 'ThumbID')),
    };
    if (keep(p)) out.push(p);
  });
  out.sort(function (a, b) { return a.ts - b.ts; });
  return out;
}

function publicPunch_(p) {
  return {
    ts: p.ts.getTime(), date: p.date, empId: p.empId, name: p.name, action: p.action, label: LABEL[p.action] || p.action,
    lat: p.lat, lng: p.lng, acc: p.acc, address: p.address, location: p.location, inside: p.inside,
    distance: p.distance, photoId: p.photoId, thumbId: p.thumbId,
  };
}

function actionCode_(v) {
  v = String(v).trim();
  if (LABEL[v]) return v;
  for (const k in LABEL) if (LABEL[k].toLowerCase() === v.toLowerCase()) return k;
  return v;
}

function locations_() {
  const v = sheet_(SH.LOC).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const name = String(v[i][0]).trim(), lat = Number(v[i][1]), lng = Number(v[i][2]);
    if (name && isFinite(lat) && isFinite(lng) && v[i][1] !== '') out.push({ name: name, lat: lat, lng: lng, radius: num_(v[i][3], 100) });
  }
  return out;
}

function holidays_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.HOL);
  const out = {};
  if (!sh) return out;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    const d = dstr_(r[0]);
    if (d) out[d] = String(r[1] || 'Holiday');
  });
  return out;
}

function settings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.SET);
  const out = {};
  if (!sh) return out;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    if (String(r[0]).trim()) out[String(r[0]).trim()] = String(r[1]).trim();
  });
  return out;
}

function setSetting_(key, value) {
  const sh = sheet_(SH.SET);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === key) { sh.getRange(i + 1, 2).setValue(txt_(value)); return; }
  }
  sh.appendRow([key, txt_(value), '']);
}

function publicSettings_() {
  const s = settings_();
  return { company: s.COMPANY_NAME || '', enforceGeofence: enforceFence_(), maxAccuracy: num_(s.MAX_ACCURACY_M, 0), tz: tz_() };
}

function enforceFence_() {
  return /^(true|yes|1)$/i.test(settings_().ENFORCE_GEOFENCE || '');
}

function monthFolder_(d) {
  const root = DriveApp.getFolderById(settings_().PHOTO_FOLDER_ID);
  const name = fmt_(d, 'yyyy-MM');
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function thumbFolder_(month) {
  const it = month.getFoldersByName('thumbs');
  return it.hasNext() ? it.next() : month.createFolder('thumbs');
}

/* ------------------------------------------------------------------ auth */

function secret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}

function hmac_(s) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, secret_())).replace(/=+$/, '');
}

function hash_(s) {
  return hmac_('pv:' + s).slice(0, 12);
}

function sign_(payload) {
  const body = Utilities.base64EncodeWebSafe(JSON.stringify(payload)).replace(/=+$/, '');
  return body + '.' + hmac_(body);
}

function verify_(token, role) {
  const parts = String(token || '').split('.');
  if (parts.length !== 2 || hmac_(parts[0]) !== parts[1]) throw new Error('AUTH: Please log in again.');
  let p;
  try {
    let b = parts[0];
    while (b.length % 4) b += '=';
    p = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(b)).getDataAsString());
  } catch (e) { throw new Error('AUTH: Please log in again.'); }
  if (p.r !== role || !(p.exp > Date.now())) throw new Error('AUTH: Session expired. Please log in again.');
  return p;
}

function authEmp_(token) {
  const p = verify_(token, 'emp');
  const e = findEmployee_(p.id);
  if (!e || hash_(e.pin) !== p.pv) throw new Error('AUTH: Please log in again.');
  if (!e.active) throw new Error('AUTH: Your account is inactive.');
  return e;
}

function adminToken_(pw) {
  return sign_({ r: 'admin', pv: hash_(pw), exp: Date.now() + 14 * 864e5 });
}

function authAdmin_(token) {
  const p = verify_(token, 'admin');
  if (p.pv !== hash_(String(settings_().ADMIN_PASSWORD || ''))) throw new Error('AUTH: Please log in again.');
}

function throttle_(key) {
  const n = Number(CacheService.getScriptCache().get('fail:' + key) || 0);
  if (n >= 5) throw new Error('Too many wrong attempts. Try again in 15 minutes.');
}

function failed_(key) {
  const c = CacheService.getScriptCache();
  c.put('fail:' + key, String(Number(c.get('fail:' + key) || 0) + 1), 900);
}

/* ------------------------------------------------------------------ helpers */

function tz_() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
}

function fmt_(d, pattern) {
  return Utilities.formatDate(d, tz_(), pattern);
}

function dateAt_(date, minutes) {
  return new Date(Utilities.parseDate(date + ' 00:00', tz_(), 'yyyy-MM-dd HH:mm').getTime() + minutes * 60000);
}

function addDays_(date, n) {
  return fmt_(new Date(dateAt_(date, 720).getTime() + n * 864e5), 'yyyy-MM-dd');
}

function dstr_(v) {
  if (v instanceof Date) return fmt_(v, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/.exec(s); // dd/mm/yyyy
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

function dateParam_(v) {
  const s = String(v || '');
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function hhmm_(v) {
  if (v instanceof Date) return fmt_(v, 'HH:mm');
  if (typeof v === 'number' && v >= 0 && v < 1) {
    const m = Math.round(v * 1440);
    return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2);
  }
  const s = String(v || '').trim().toUpperCase();
  const m = /^(\d{1,2})[:.](\d{2})(?::\d{2})?\s*(AM|PM)?$/.exec(s);
  if (!m) return '';
  let h = Number(m[1]);
  const mi = Number(m[2]);
  if (m[3] === 'PM' && h < 12) h += 12;
  if (m[3] === 'AM' && h === 12) h = 0;
  if (h > 23 || mi > 59) return '';
  return ('0' + h).slice(-2) + ':' + ('0' + mi).slice(-2);
}

function toMin_(hm) {
  const p = String(hm).split(':');
  return Number(p[0]) * 60 + Number(p[1]);
}

function num_(v, d) {
  const n = Number(v);
  return v === '' || v === null || !isFinite(n) ? d : n;
}

/** Leading apostrophe stops Sheets turning "0123", "09:00" or "2026-10-10" into numbers/dates. */
function txt_(v) {
  v = String(v === undefined || v === null ? '' : v);
  return v === '' ? '' : "'" + v;
}
