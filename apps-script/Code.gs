const VERSION = '2026-10-13.3';

const SH = { EMP: 'Employees', PUNCH: 'Punches', LOC: 'Locations', SET: 'Settings', HOL: 'Holidays', ALERT: 'Alerts', OUT: 'Outside',
  CO: 'Companies', REQ: 'Requests' };

const EMP_COLS = ['EmpID', 'Name', 'PIN', 'Phone', 'Department', 'ShiftStart', 'ShiftEnd',
  'LunchMinutes', 'GraceMinutes', 'WeeklyOff', 'AllowedLocations', 'Active', 'JoinDate', 'CanPunchOthers', 'FaceData', 'FaceRegistered', 'LunchStart', 'HalfDayOff'];
const PUNCH_COLS = ['Timestamp', 'WorkDate', 'EmpID', 'Name', 'Action', 'Latitude', 'Longitude',
  'AccuracyM', 'Address', 'Location', 'InsideFence', 'DistanceM', 'PhotoURL', 'PhotoID',
  'ThumbID', 'DeviceTime', 'Device', 'PunchedBy', 'FaceMatch', 'PunchID'];
const LOC_COLS = ['Name', 'Latitude', 'Longitude', 'RadiusM'];
const HOL_COLS = ['Date', 'Name'];
const OUT_COLS = ['ID', 'WorkDate', 'EmpID', 'Name', 'Start', 'End', 'Minutes', 'MaxDistanceM', 'Place', 'Updated'];
const ALERT_COLS = ['Timestamp', 'WorkDate', 'EmpID', 'Name', 'Type', 'Detail', 'Latitude', 'Longitude'];
const COMPANY_COLS = ['Code', 'Name', 'SheetID', 'Contact', 'Phone', 'Created', 'TrialEnds', 'Active', 'Notes'];
const REQ_COLS = ['Time', 'Company', 'Name', 'Phone', 'City', 'Staff', 'Status', 'Code'];

let CO = null;

const DEFAULT_SETTINGS = [
  ['COMPANY_NAME', 'My Company', 'Shown in the app and on every selfie stamp'],
  ['ADMIN_PASSWORD', '', 'Password for the admin dashboard'],
  ['ENFORCE_GEOFENCE', 'FALSE', 'TRUE = block punches made outside the radius of the Locations tab'],
  ['MAX_ACCURACY_M', '200', 'Reject GPS fixes less accurate than this many metres (0 = accept any)'],
  ['PHOTO_FOLDER_ID', '', 'Drive folder for selfies (created by setup)'],
  ['VERIFY_IN_OUT', 'SELFIE', 'Check In / Check Out: SELFIE (photo only), FACE (face recognition only, no photo) or BOTH'],
  ['VERIFY_LUNCH', 'SELFIE', 'Lunch Out / Lunch In: SELFIE, FACE or BOTH'],
  ['FACE_MODE', 'WARN', 'With BOTH: WARN = record a face mismatch only, REQUIRED = block the punch'],
  ['FACE_THRESHOLD', '0.5', 'Face match strictness: lower = stricter (0.4 strict, 0.5 normal, 0.6 loose)'],
  ['LUNCH_START', '13:00', 'Default lunch start time for reminders (each employee can have their own)'],
  ['AWAY_MINUTES', '15', 'Android app: alert when someone is outside the work location this many minutes in lunch hour without Lunch Out'],
  ['MIN_HALF_DAY_HOURS', '3', 'Worked less than this (hours) = Absent; from this up to half the shift = Half Day'],
  ['LUNCH_MODE', 'SELFIE', 'SELFIE = staff take Lunch Out / Lunch In selfies. AUTO = no lunch selfie, lunch minutes are deducted automatically'],
];

const AUTO_LUNCH_AFTER_MIN = 300;

const LABEL = { IN: 'Check In', LUNCH_OUT: 'Lunch Out', LUNCH_IN: 'Lunch In', OUT: 'Check Out' };
const NEXT = { '': ['IN'], IN: ['LUNCH_OUT', 'OUT'], LUNCH_OUT: ['LUNCH_IN'], LUNCH_IN: ['LUNCH_OUT', 'OUT'], OUT: [] };

function setup() {
  CO = null;
  resetMemo_();
  prepare_(null, '');
  if (!settings_().OWNER_PASSWORD) setSetting_('OWNER_PASSWORD', 'owner' + Math.floor(100000 + Math.random() * 900000));
  try { Maps.newGeocoder().reverseGeocode(19.076, 72.8777); } catch (e) { }
  clearDataCache_();
  const msg = 'Setup complete.\n\nAdmin password: ' + settings_().ADMIN_PASSWORD +
    '\nOwner panel password: ' + settings_().OWNER_PASSWORD +
    '\n(Both are in the Settings tab.)\n\nNow deploy: Deploy > New deployment > Web app,' +
    ' Execute as: Me, Who has access: Anyone.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { }
}

function prepare_(parent, name) {
  const ss = ss_();
  ensureSheet_(ss, SH.EMP, EMP_COLS, ['EmpID', 'PIN', 'Phone', 'ShiftStart', 'ShiftEnd', 'WeeklyOff', 'AllowedLocations', 'JoinDate']);
  ensureSheet_(ss, SH.PUNCH, PUNCH_COLS, ['WorkDate', 'EmpID']);
  ensureSheet_(ss, SH.LOC, LOC_COLS, ['Name']);
  ensureSheet_(ss, SH.HOL, HOL_COLS, ['Date']);
  ensureSheet_(ss, SH.ALERT, ALERT_COLS, ['WorkDate', 'EmpID']);
  ensureSheet_(ss, SH.OUT, OUT_COLS, ['ID', 'WorkDate', 'EmpID']);
  ensureSheet_(ss, SH.SET, ['Key', 'Value', 'Notes'], ['Key', 'Value']);

  const have = settings_();
  DEFAULT_SETTINGS.forEach(function (d) {
    if (!(d[0] in have)) sheet_(SH.SET).appendRow([d[0], txt_(d[0] === 'COMPANY_NAME' && name ? name : d[1]), d[2]]);
  });
  clearDataCache_();
  const s = settings_();
  if (!s.ADMIN_PASSWORD) setSetting_('ADMIN_PASSWORD', 'admin' + Math.floor(100000 + Math.random() * 900000));
  if (!s.PHOTO_FOLDER_ID) setSetting_('PHOTO_FOLDER_ID', (parent || DriveApp).createFolder(ss.getName() + ' - Selfies').getId());
  secret_();

  const emp = sheet_(SH.EMP);
  if (emp.getLastRow() < 2) {
    emp.appendRow([txt_('EMP001'), 'Sample Employee', txt_('1234'), '', 'Office', txt_('09:00'), txt_('18:00'), 45, 10, txt_('Sun'), '', 'Yes', txt_(fmt_(new Date(), 'yyyy-MM-dd'))]);
  }
  ['Sheet1', 'Sheet 1'].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh && ss.getSheets().length > 1 && sh.getLastRow() === 0) ss.deleteSheet(sh);
  });
  clearDataCache_();
}

function ss_() {
  return CO ? CO.ss : SpreadsheetApp.getActive();
}

function cache_() {
  const c = CacheService.getScriptCache(), p = CO ? CO.code + ':' : '';
  return {
    get: function (k) { return c.get(p + k); },
    put: function (k, v, t) { return c.put(p + k, v, t); },
    remove: function (k) { return c.remove(p + k); },
    removeAll: function (ks) { return c.removeAll(ks.map(function (k) { return p + k; })); },
  };
}

function resetMemo_() {
  EMPLOYEES_MEMO = null;
  SETTINGS_MEMO = null;
}

const DATA_CACHE = ['emp_v1', 'set_v1', 'hol_v1', 'loc_v1'];

function cacheGet_(key) {
  try { const v = cache_().get(key); return v ? JSON.parse(v) : null; } catch (e) { return null; }
}

function cachePut_(key, value) {
  try { const v = JSON.stringify(value); if (v.length < 95000) cache_().put(key, v, 600); } catch (e) { }
}

function clearDataCache_() {
  try { cache_().removeAll(DATA_CACHE); } catch (e) { }
  EMPLOYEES_MEMO = null;
  SETTINGS_MEMO = null;
}

function forgetEmployees_() {
  try { cache_().remove('emp_v1'); } catch (e) { }
  EMPLOYEES_MEMO = null;
}

function onEdit() {
  CO = null;
  clearDataCache_();
  try { CacheService.getScriptCache().remove('companies_v1'); } catch (e) { }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Attendance').addItem('Run setup', 'setup').addToUi();
}

function ensureSchema_() {
  const cache = cache_();
  if (cache.get('schema') === '8') return;
  const ss = ss_();
  [[SH.EMP, EMP_COLS], [SH.PUNCH, PUNCH_COLS]].forEach(function (x) {
    const sh = ss.getSheetByName(x[0]);
    if (!sh || sh.getLastRow() === 0) return;
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    x[1].forEach(function (c) {
      if (head.indexOf(c) < 0) {
        sh.getRange(1, head.length + 1).setValue(c).setFontWeight('bold').setBackground('#e6f4f1');
        head.push(c);
      }
    });
  });
  cache.put('schema', '8', 21600);
  if (!ss.getSheetByName(SH.OUT)) ensureSheet_(ss, SH.OUT, OUT_COLS, ['ID', 'WorkDate', 'EmpID']);
  if (!ss.getSheetByName(SH.ALERT)) ensureSheet_(ss, SH.ALERT, ALERT_COLS, ['WorkDate', 'EmpID']);
}

function ensureSheet_(ss, name, cols, textCols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#e6f4f1');
    sh.setFrozenRows(1);
  } else {
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

const ROUTES = {
  login: login_,
  status: status_,
  geocode: geocode_,
  punch: punch_,
  attachPhoto: attachPhoto_,
  adminLogin: adminLogin_,
  adminDay: adminDay_,
  adminSavePunch: adminSavePunch_,
  adminDeletePunch: adminDeletePunch_,
  reportAway: reportAway_,
  reportOutside: reportOutside_,
  adminTestPush: adminTestPush_,
  adminReport: adminReport_,
  adminEmployees: adminEmployees_,
  adminSaveEmployee: adminSaveEmployee_,
  adminLocations: adminLocations_,
  adminSaveLocation: adminSaveLocation_,
  adminDeleteLocation: adminDeleteLocation_,
  adminSaveSettings: adminSaveSettings_,
  adminThumbs: adminThumbs_,
  adminPhoto: adminPhoto_,
  faceRefs: faceRefs_,
  health: health_,
  adminSaveFace: adminSaveFace_,
  adminResetFace: adminResetFace_,
  ownerLogin: ownerLogin_,
  ownerData: ownerData_,
  ownerCreateCompany: ownerCreateCompany_,
  ownerSaveCompany: ownerSaveCompany_,
  ownerResetAdmin: ownerResetAdmin_,
  ownerOpenAdmin: ownerOpenAdmin_,
  ownerShareInfo: ownerShareInfo_,
  ownerSaveRequest: ownerSaveRequest_,
  requestTrial: requestTrial_,
};

function doPost(e) {
  let out;
  try {
    CO = null;
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!ROOT_ACTIONS[req.action]) useCompany_(req.company);
    ensureSchema_();
    const fn = ROUTES[req.action];
    if (!fn) throw new Error('Unknown action');
    out = Object.assign({ ok: true }, fn(req));
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'attendance', version: VERSION, time: Date.now() }))
    .setMimeType(ContentService.MimeType.JSON);
}

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
  const e = authEmp_(req.token);
  const out = statusFor_(e);
  if (req.native) out.notice = takeNotice_(e);
  return out;
}

function geocode_(req) {
  const me = authEmp_(req.token);
  const e = req.forEmpId ? resolveTargets_(me, [req.forEmpId])[0] : me;
  const lat = Number(req.lat), lng = Number(req.lng);
  if (!isFinite(lat) || !isFinite(lng)) throw new Error('Bad location');
  const f = fence_(e, lat, lng);
  return { address: address_(lat, lng), place: f.name || '', inside: f.inside, distance: f.distance, fenced: f.has, enforce: enforceFence_() };
}

function resolveTargets_(me, ids) {
  ids = (ids || []).map(function (x) { return String(x).trim().toUpperCase(); }).filter(String);
  ids = ids.filter(function (x, i) { return ids.indexOf(x) === i; });
  if (!ids.length || (ids.length === 1 && ids[0] === me.id.toUpperCase())) return [me];
  if (!me.canPunch) throw new Error('You are not allowed to take selfies for other employees.');
  if (ids.length > 50) throw new Error('A group selfie can have at most 50 employees.');
  const all = {};
  employees_().forEach(function (x) { all[x.id.toUpperCase()] = x; });
  return ids.map(function (id) {
    const t = all[id];
    if (!t) throw new Error('Employee ' + id + ' not found.');
    if (!t.active) throw new Error(t.name + ' is inactive.');
    return t;
  });
}

function names_(list, max) {
  max = max || 3;
  const n = list.map(function (t) { return t.name; });
  return n.length > max ? n.slice(0, max).join(', ') + ' +' + (n.length - max) + ' more' : n.join(', ');
}

function punch_(req) {
  const me = authEmp_(req.token);
  const type = String(req.type || '');
  if (!LABEL[type]) throw new Error('Unknown punch type');
  const targets = resolveTargets_(me, req.forEmpIds || (req.forEmpId ? [req.forEmpId] : []));
  const forOthers = !(targets.length === 1 && targets[0].id === me.id);

  const method = verifyFor_(type);
  const m = /^data:image\/jpeg;base64,(.+)$/.exec(String(req.photo || ''));
  const photoComing = !!m || !!req.photoPending;
  if (!photoComing && method !== 'FACE') throw new Error('A selfie is required.');
  if (m && m[1].length > 8e6) throw new Error('Photo is too large.');
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
  const fences = targets.map(function (t) { return fence_(t, lat, lng); });
  if (enforceFence_()) {
    const outside = targets.filter(function (t, i) { return fences[i].has && !fences[i].inside; });
    if (outside.length) {
      const f = fences[targets.indexOf(outside[0])];
      throw new Error((forOthers ? names_(outside) + (outside.length > 1 ? ' are ' : ' is ') : 'You are ') + f.distance +
        ' m away from ' + f.name + '. Punch from the work location.');
    }
  }

  const face = checkFaces_(targets, cleanFaces_(req.faces), forOthers, req.faceError, method, photoComing);

  const punchId = Utilities.getUuid();
  let recent;
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const now = new Date();
    recent = recentPunches_(now);
    const hol = holidays_();
    const states = targets.map(function (t) { return empState_(t, now, recent, hol); });
    const bad = targets.filter(function (t, i) { return states[i].next.indexOf(type) < 0; });
    if (bad.length) {
      const st = states[targets.indexOf(bad[0])];
      if (!forOthers) {
        throw new Error(st.next.length
          ? 'You cannot ' + LABEL[type] + ' now. Next: ' + st.next.map(function (t) { return LABEL[t]; }).join(' or ') + '.'
          : 'You have already checked out for this shift.');
      }
      throw new Error('Cannot ' + LABEL[type] + ' for ' + names_(bad) + ' now. Refresh the list and try again.');
    }
    const stamp = fmt_(now, 'yyyy-MM-dd_HHmmss');
    const base = (targets.length > 1 ? 'GROUP' + targets.length : targets[0].id) + '_' + stamp + '_' + type;
    const folder = m ? monthFolder_(now) : null;
    const file = m ? folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/jpeg', base + '.jpg')) : null;
    if (file) file.setDescription(names_(targets, 999) + ' - ' + LABEL[type] + ' - ' + fmt_(now, 'dd MMM yyyy HH:mm:ss') + ' - ' +
      lat + ',' + lng + (forOthers ? ' - taken by ' + me.name + ' (' + me.id + ')' : ''));
    let thumbId = '';
    if (m && tm) thumbId = thumbFolder_(folder).createFile(Utilities.newBlob(Utilities.base64Decode(tm[1]), 'image/jpeg', base + '_t.jpg')).getId();
    const address = String(req.address || '').slice(0, 300) || address_(lat, lng);
    const by = forOthers ? me.id + ' - ' + me.name + (targets.length > 1 ? ' (group of ' + targets.length + ')' : '') : '';
    const rows = targets.map(function (t, i) {
      const f = fences[i];
      return [
        now, txt_(states[i].workDate), txt_(t.id), txt_(t.name), LABEL[type], lat, lng, acc, txt_(address),
        txt_(f.name), f.has ? (f.inside ? 'Yes' : 'No') : '', f.has ? f.distance : '',
        file ? file.getUrl() : (req.photoPending ? 'pending' : ''), file ? file.getId() : '', thumbId,
        txt_(String(req.deviceTime || '').slice(0, 40)), txt_(String(req.device || '').slice(0, 150)), txt_(by),
        face.results[i], punchId,
      ];
    });
    const sh = sheet_(SH.PUNCH);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    targets.forEach(function (t, i) {
      recent.push({ ts: now, date: states[i].workDate, empId: t.id, name: t.name, action: type, lat: lat, lng: lng, acc: acc,
        address: address, location: fences[i].name || '', inside: fences[i].has ? (fences[i].inside ? 'Yes' : 'No') : '',
        distance: fences[i].has ? fences[i].distance : '', photoId: file ? file.getId() : '', thumbId: thumbId, by: by,
        face: face.results[i], photoUrl: file ? file.getUrl() : (req.photoPending ? 'pending' : '') });
    });
    Object.keys(face.enroll).forEach(function (id) { saveFace_(id, face.enroll[id], 'first selfie'); });
    if (face.enroll[me.id]) me.faceData = face.enroll[me.id];
  } finally {
    lock.releaseLock();
  }
  const res = statusFor_(me, recent);
  res.punchId = punchId;
  res.punchedFor = targets.map(function (t, i) { return { id: t.id, name: t.name, face: face.results[i] }; });
  return res;
}

function reminderInfo_(e, workDate) {
  const s = settings_();
  let locs = locations_();
  const allowed = String(e.locations || '').toLowerCase().split(',').map(function (x) { return x.trim(); }).filter(String);
  if (allowed.length) locs = locs.filter(function (l) { return allowed.indexOf(l.name.toLowerCase()) >= 0; });
  return {
    lunchStart: e.lunchStart || defaultLunchStart_(),
    lunchMinutes: num_(e.lunch, 0),
    awayMinutes: Math.max(1, num_(s.AWAY_MINUTES, 15)),
    shiftEnd: workDate && isHalfDayOff_(e, workDate) && !isWeeklyOff_(e, workDate) && !holidays_()[workDate]
      ? halfShiftEnd_(e) : e.shiftEnd,
    autoLunch: autoLunch_(),
    fences: locs.map(function (l) { return { name: l.name, lat: l.lat, lng: l.lng, radius: l.radius }; }),
  };
}

function defaultLunchStart_() {
  const v = settings_().LUNCH_START;
  return v === undefined ? '13:00' : (hhmm_(v) || '');
}

function reportAway_(req) {
  const e = authEmp_(req.token);
  const now = new Date();
  const wd = workDate_(e, now);
  const minutes = Math.max(0, Math.round(num_(req.minutes, 0)));
  const dist = Math.round(num_(req.distance, 0));
  const key = 'away:' + e.id + ':' + wd;
  const cache = cache_();
  if (cache.get(key)) return { duplicate: true };
  cache.put(key, '1', 21600);
  const detail = 'Outside ' + (req.place ? String(req.place).slice(0, 60) + ' ' : 'work location ') + 'for ' + minutes +
    ' min in lunch hour without Lunch Out' + (dist ? ' (' + dist + ' m away)' : '');
  sheet_(SH.ALERT).appendRow([now, txt_(wd), txt_(e.id), txt_(e.name), 'Away in lunch hour', txt_(detail),
    num_(req.lat, ''), num_(req.lng, '')]);
  return { logged: true };
}

function reportOutside_(req) {
  const e = authEmp_(req.token);
  const id = String(req.id || '');
  if (!/^[\w-]{8,64}$/.test(id)) throw new Error('Bad id');
  const now = Date.now();
  const start = Math.min(Number(req.start), now + 300000);
  const end = req.end ? Math.min(Number(req.end), now + 300000) : null;
  if (!isFinite(start) || start < now - 7 * 864e5) throw new Error('Bad start time');
  if (end !== null && (!isFinite(end) || end < start)) throw new Error('Bad end time');
  const row = [txt_(id), txt_(workDate_(e, new Date(start))), txt_(e.id), txt_(e.name), new Date(start), end ? new Date(end) : '',
    end ? Math.round((end - start) / 60000) : '', Math.round(num_(req.maxDistance, 0)), txt_(String(req.place || '').slice(0, 60)), new Date()];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SH.OUT);
    const last = sh.getLastRow();
    let found = 0;
    if (last >= 2) {
      const first = Math.max(2, last - 499);
      const ids = sh.getRange(first, 1, last - first + 1, 3).getValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === id) {
          if (String(ids[i][2]).toUpperCase() !== e.id.toUpperCase()) throw new Error('Not your record');
          found = first + i;
          break;
        }
      }
    }
    if (found) sh.getRange(found, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return { saved: true };
}

function outsideFor_(from, to) {
  const sh = ss_().getSheetByName(SH.OUT);
  if (!sh || sh.getLastRow() < 2) return [];
  const c = headerMap_(sh);
  const last = sh.getLastRow();
  const first = Math.max(2, last - 4999);
  const now = Date.now();
  return sh.getRange(first, 1, last - first + 1, sh.getLastColumn()).getValues().map(function (r) {
    const st = r[c.Start], en = r[c.End];
    const startMs = st instanceof Date ? st.getTime() : 0, endMs = en instanceof Date ? en.getTime() : null;
    return { id: String(r[c.ID]), date: dstr_(r[c.WorkDate]), empId: String(r[c.EmpID]).trim(), start: startMs, end: endMs,
      minutes: Math.max(0, Math.round(((endMs || now) - startMs) / 60000)), maxDistance: num_(r[c.MaxDistanceM], 0), place: String(r[c.Place]) };
  }).filter(function (o) { return o.start && o.date >= from && o.date <= to; });
}

function alertsFor_(date) {
  const sh = ss_().getSheetByName(SH.ALERT);
  if (!sh || sh.getLastRow() < 2) return [];
  const c = headerMap_(sh);
  const last = sh.getLastRow();
  const start = Math.max(2, last - 1999);
  return sh.getRange(start, 1, last - start + 1, sh.getLastColumn()).getValues().filter(function (r) {
    return dstr_(r[c.WorkDate]) === date;
  }).map(function (r) {
    const ts = r[c.Timestamp];
    return { ts: ts instanceof Date ? ts.getTime() : 0, empId: String(r[c.EmpID]), type: String(r[c.Type]), detail: String(r[c.Detail]) };
  });
}

function adminTestPush_(req) {
  authAdmin_(req.token);
  const e = findEmployee_(req.id);
  if (!e) throw new Error('Employee not found.');
  cache_().put('note:' + e.id.toUpperCase(), JSON.stringify({
    id: Utilities.getUuid().slice(0, 8), title: 'Attendance test', body: 'Test from the admin: notifications are working on this phone.' }), 21600);
  return { queued: true };
}

function takeNotice_(e) {
  const cache = cache_();
  const key = 'note:' + e.id.toUpperCase();
  const v = cache.get(key);
  if (!v) return null;
  cache.remove(key);
  try { return JSON.parse(v); } catch (err) { return null; }
}

function attachPhoto_(req) {
  const me = authEmp_(req.token);
  const id = String(req.punchId || '');
  if (!/^[\w-]{20,}$/.test(id)) throw new Error('Bad punch id');
  const m = /^data:image\/jpeg;base64,(.+)$/.exec(String(req.photo || ''));
  if (!m) throw new Error('No photo');
  if (m[1].length > 8e6) throw new Error('Photo is too large.');
  const tm = /^data:image\/jpeg;base64,(.+)$/.exec(String(req.thumb || ''));
  const sh = sheet_(SH.PUNCH);
  const c = headerMap_(sh);
  if (c.PunchID === undefined) throw new Error('Punch not found');
  const width = sh.getLastColumn();
  let found = [], end = sh.getLastRow(), chunk = 300;
  while (end >= 2 && !found.length) {
    const start = Math.max(2, end - chunk + 1);
    sh.getRange(start, 1, end - start + 1, width).getValues().forEach(function (r, i) {
      if (String(r[c.PunchID]) === id) found.push({ row: start + i, r: r });
    });
    end = start - 1;
    chunk *= 2;
  }
  if (!found.length) throw new Error('Punch not found');
  const mine = found.every(function (x) {
    return String(x.r[c.EmpID]).toUpperCase() === me.id.toUpperCase() || String(x.r[c.PunchedBy]).indexOf(me.id + ' - ') === 0;
  });
  if (!mine) throw new Error('Not your punch');
  if (String(found[0].r[c.PhotoID] || '')) return { photoId: String(found[0].r[c.PhotoID]) };

  const cache = cache_();
  if (cache.get('attach:' + id)) return { busy: true };
  cache.put('attach:' + id, '1', 120);
  try {
    return savePunchPhoto_(sh, c, found, m, tm);
  } finally {
    cache.remove('attach:' + id);
  }
}

function savePunchPhoto_(sh, c, found, m, tm) {
  const first = found[0].r;
  const when = first[c.Timestamp] instanceof Date ? first[c.Timestamp] : new Date();
  const base = (found.length > 1 ? 'GROUP' + found.length : String(first[c.EmpID])) + '_' + fmt_(when, 'yyyy-MM-dd_HHmmss') + '_' +
    actionCode_(first[c.Action]);
  const folder = monthFolder_(when);
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/jpeg', base + '.jpg'));
  file.setDescription(found.map(function (x) { return String(x.r[c.Name]); }).join(', ') + ' - ' + String(first[c.Action]) + ' - ' +
    fmt_(when, 'dd MMM yyyy HH:mm:ss') + ' - ' + first[c.Latitude] + ',' + first[c.Longitude] +
    (String(first[c.PunchedBy]) ? ' - taken by ' + String(first[c.PunchedBy]) : ''));
  const thumbId = tm ? thumbFolder_(folder).createFile(Utilities.newBlob(Utilities.base64Decode(tm[1]), 'image/jpeg', base + '_t.jpg')).getId() : '';
  found.forEach(function (x) {
    sh.getRange(x.row, c.PhotoURL + 1).setValue(file.getUrl());
    sh.getRange(x.row, c.PhotoID + 1).setValue(file.getId());
    if (c.ThumbID !== undefined) sh.getRange(x.row, c.ThumbID + 1).setValue(thumbId);
  });
  return { photoId: file.getId() };
}

function statusFor_(e, given) {
  const now = new Date();
  const recent = given || recentPunches_(now);
  const hol = holidays_();
  const st = empState_(e, now, recent, hol);
  const out = {
    company: settings_().COMPANY_NAME || '',
    tz: tz_(),
    serverTime: now.getTime(),
    profile: {
      id: e.id, name: e.name, dept: e.dept, shiftStart: e.shiftStart, shiftEnd: e.shiftEnd,
      lunch: e.lunch, grace: e.grace, weeklyOff: e.weeklyOff, halfDayOff: e.halfDayOff, canPunchOthers: e.canPunch,
      hasFace: !!e.faceData,
    },
    lunchMode: autoLunch_() ? 'AUTO' : 'SELFIE',
    faceMode: faceMode_(),
    verify: { IN: verifyFor_('IN'), OUT: verifyFor_('OUT'), LUNCH_OUT: verifyFor_('LUNCH_OUT'), LUNCH_IN: verifyFor_('LUNCH_IN') },
    faceStrict: faceStrict_(),
    reminders: reminderInfo_(e, st.workDate),
    workDate: st.workDate,
    punches: st.punches.map(publicPunch_),
    next: st.next,
    summary: st.summary,
  };
  if (e.canPunch) out.team = team_(e, now, recent, hol);
  return out;
}

function team_(me, now, recent, hol) {
  return employees_().filter(function (t) { return t.active && t.id !== me.id; }).map(function (t) {
    const st = empState_(t, now, recent, hol);
    const last = st.punches[st.punches.length - 1];
    return {
      id: t.id, name: t.name, dept: t.dept, shift: t.shiftStart + '-' + t.shiftEnd, next: st.next,
      status: st.summary.status, late: st.summary.late,
      last: last ? { label: LABEL[last.action] || last.action, ts: last.ts.getTime() } : null,
    };
  });
}

function recentPunches_(now) {
  return punchesSince_(addDays_(fmt_(now, 'yyyy-MM-dd'), -2), function () { return true; });
}

function empState_(e, now, recent, hol) {
  const wd = workDate_(e, now);
  const list = recent.filter(function (p) { return p.empId === e.id && p.date === wd; });
  const last = list.length ? list[list.length - 1].action : '';
  return { workDate: wd, punches: list, next: next_(last), summary: summarize_(e, wd, list, now, hol) };
}

function next_(last) {
  const n = NEXT[last] || [];
  return autoLunch_() ? n.filter(function (t) { return t !== 'LUNCH_OUT'; }) : n;
}

function autoLunch_() {
  return /^auto/i.test(settings_().LUNCH_MODE || '');
}

const ROOT_ACTIONS = { ownerLogin: 1, ownerData: 1, ownerCreateCompany: 1, ownerSaveCompany: 1, ownerResetAdmin: 1,
  ownerOpenAdmin: 1, ownerShareInfo: 1, ownerSaveRequest: 1, requestTrial: 1 };

function root_() {
  return SpreadsheetApp.getActive();
}

function rootSheet_(name) {
  const sh = root_().getSheetByName(name);
  if (sh) return sh;
  return name === SH.CO
    ? ensureSheet_(root_(), SH.CO, COMPANY_COLS, ['Code', 'SheetID', 'Phone', 'Created', 'TrialEnds'])
    : ensureSheet_(root_(), SH.REQ, REQ_COLS, ['Phone', 'Code']);
}

function companies_() {
  const c = CacheService.getScriptCache();
  try { const v = c.get('companies_v1'); if (v) return JSON.parse(v); } catch (e) { }
  const sh = root_().getSheetByName(SH.CO);
  const out = [];
  if (sh && sh.getLastRow() > 1) {
    const h = headerMap_(sh);
    const g = function (r, k) { return h[k] === undefined ? '' : r[h[k]]; };
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().forEach(function (r, i) {
      const code = String(g(r, 'Code')).trim().toUpperCase();
      if (!code) return;
      out.push({ row: i + 2, code: code, name: String(g(r, 'Name')), sheetId: String(g(r, 'SheetID')).trim(),
        contact: String(g(r, 'Contact')), phone: String(g(r, 'Phone')), created: dstr_(g(r, 'Created')),
        trialEnds: dstr_(g(r, 'TrialEnds')), active: !/^(no|false|0|off)$/i.test(String(g(r, 'Active')).trim()), notes: String(g(r, 'Notes')) });
    });
  }
  try { c.put('companies_v1', JSON.stringify(out), 600); } catch (e) { }
  return out;
}

function companyByCode_(code) {
  code = String(code || '').trim().toUpperCase();
  const c = companies_().filter(function (x) { return x.code === code; })[0];
  if (!c) throw new Error('Company code ' + code + ' not found. Check the code and try again.');
  return c;
}

function useCompany_(code) {
  if (!String(code || '').trim()) return;
  const c = companyByCode_(code);
  if (!c.active) throw new Error('This account is switched off. Contact the app provider.');
  if (c.trialEnds && c.trialEnds < fmt_(new Date(), 'yyyy-MM-dd')) {
    throw new Error('The free trial for ' + c.name + ' has ended. Contact the app provider to continue.');
  }
  enterCompany_(c);
}

function enterCompany_(c) {
  CO = { code: c.code, ss: SpreadsheetApp.openById(c.sheetId) };
  resetMemo_();
}

function leaveCompany_() {
  CO = null;
  resetMemo_();
}

function clip_(v, n) {
  return String(v === undefined || v === null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, n);
}

function newCode_() {
  const a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 6; i++) s += a.charAt(Math.floor(Math.random() * a.length));
  return s;
}

function newPassword_() {
  return 'admin' + Math.floor(100000 + Math.random() * 900000);
}

function companiesFolder_() {
  const id = settings_().COMPANIES_FOLDER_ID;
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  const f = DriveApp.createFolder(root_().getName() + ' - Companies');
  setSetting_('COMPANIES_FOLDER_ID', f.getId());
  return f;
}

function ownerLogin_(req) {
  throttle_('owner');
  const pw = String(settings_().OWNER_PASSWORD || '');
  if (!pw) throw new Error('The owner panel is not set up here. Run setup in the Google Sheet once.');
  if (String(req.password || '') !== pw) {
    failed_('owner');
    throw new Error('Wrong password.');
  }
  return { token: sign_({ r: 'owner', pv: hash_(pw), exp: Date.now() + 30 * 864e5 }) };
}

function authOwner_(token) {
  const p = verify_(token, 'owner');
  const pw = String(settings_().OWNER_PASSWORD || '');
  if (!pw || p.pv !== hash_(pw)) throw new Error('AUTH: Please log in again.');
}

function ownerData_(req) {
  authOwner_(req.token);
  const today = fmt_(new Date(), 'yyyy-MM-dd');
  const companies = companies_().map(function (c) {
    return { code: c.code, name: c.name, contact: c.contact, phone: c.phone, created: c.created, trialEnds: c.trialEnds,
      active: c.active, notes: c.notes, status: !c.active ? 'Off' : (c.trialEnds && c.trialEnds < today ? 'Ended' : 'Active') };
  }).reverse();
  const sh = root_().getSheetByName(SH.REQ);
  const requests = [];
  if (sh && sh.getLastRow() > 1) {
    const h = headerMap_(sh);
    const from = Math.max(2, sh.getLastRow() - 299);
    sh.getRange(from, 1, sh.getLastRow() - from + 1, sh.getLastColumn()).getValues().forEach(function (r, i) {
      const t = r[h.Time];
      requests.push({ row: from + i, time: t instanceof Date ? t.getTime() : new Date(t).getTime() || 0, company: String(r[h.Company]),
        name: String(r[h.Name]), phone: String(r[h.Phone]), city: String(r[h.City]), staff: String(r[h.Staff]),
        status: String(r[h.Status]), code: String(r[h.Code]) });
    });
  }
  return { companies: companies, requests: requests.reverse(), today: today, tz: tz_() };
}

function ownerCreateCompany_(req) {
  authOwner_(req.token);
  const name = clip_(req.name, 80);
  if (!name) throw new Error('Enter the company name.');
  const days = Math.max(0, Math.min(365, Math.round(num_(req.trialDays, 15))));
  const tz = tz_(), today = fmt_(new Date(), 'yyyy-MM-dd');
  const trialEnds = days ? addDays_(today, days) : '';
  let code, pw, book;
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    CacheService.getScriptCache().remove('companies_v1');
    const taken = companies_().map(function (c) { return c.code; });
    do { code = newCode_(); } while (taken.indexOf(code) >= 0);
    const folder = companiesFolder_();
    book = SpreadsheetApp.create('Attendance - ' + name + ' (' + code + ')');
    try { book.setSpreadsheetTimeZone(tz); } catch (e) { }
    try { DriveApp.getFileById(book.getId()).moveTo(folder); } catch (e) { }
    CO = { code: code, ss: book };
    resetMemo_();
    prepare_(folder, name);
    pw = settings_().ADMIN_PASSWORD;
    leaveCompany_();
    rootSheet_(SH.CO).appendRow([txt_(code), txt_(name), txt_(book.getId()), txt_(clip_(req.contact, 60)), txt_(clip_(req.phone, 30)),
      txt_(today), txt_(trialEnds), 'Yes', '']);
    if (req.requestRow) markRequest_(req.requestRow, 'Done', code);
    CacheService.getScriptCache().remove('companies_v1');
  } finally {
    leaveCompany_();
    lock.releaseLock();
  }
  return { created: { code: code, name: name, trialEnds: trialEnds, adminPassword: pw } };
}

function ownerSaveCompany_(req) {
  authOwner_(req.token);
  CacheService.getScriptCache().remove('companies_v1');
  const c = companyByCode_(req.code);
  const sh = rootSheet_(SH.CO), h = headerMap_(sh);
  const set = function (k, v) { if (h[k] !== undefined) sh.getRange(c.row, h[k] + 1).setValue(v); };
  if (req.contact !== undefined) set('Contact', txt_(clip_(req.contact, 60)));
  if (req.phone !== undefined) set('Phone', txt_(clip_(req.phone, 30)));
  if (req.notes !== undefined) set('Notes', txt_(clip_(req.notes, 200)));
  if (req.trialEnds !== undefined) {
    const d = dateParam_(req.trialEnds);
    if (req.trialEnds && !d) throw new Error('Choose a valid trial end date.');
    set('TrialEnds', txt_(d));
  }
  if (req.active !== undefined) set('Active', req.active ? 'Yes' : 'No');
  CacheService.getScriptCache().remove('companies_v1');
  return ownerData_(req);
}

function ownerResetAdmin_(req) {
  authOwner_(req.token);
  const c = companyByCode_(req.code);
  const pw = newPassword_();
  try {
    enterCompany_(c);
    setSetting_('ADMIN_PASSWORD', pw);
  } finally {
    leaveCompany_();
  }
  return { code: c.code, adminPassword: pw };
}

function ownerOpenAdmin_(req) {
  authOwner_(req.token);
  const c = companyByCode_(req.code);
  try {
    enterCompany_(c);
    return { code: c.code, token: adminToken_(String(settings_().ADMIN_PASSWORD || '')) };
  } finally {
    leaveCompany_();
  }
}

function ownerShareInfo_(req) {
  authOwner_(req.token);
  const c = companyByCode_(req.code);
  try {
    enterCompany_(c);
    return { share: { code: c.code, name: c.name, trialEnds: c.trialEnds, adminPassword: String(settings_().ADMIN_PASSWORD || '') } };
  } finally {
    leaveCompany_();
  }
}

function markRequest_(row, status, code) {
  const sh = rootSheet_(SH.REQ), h = headerMap_(sh);
  row = Math.round(Number(row));
  if (!(row >= 2 && row <= sh.getLastRow())) throw new Error('Request not found. Refresh and try again.');
  sh.getRange(row, h.Status + 1).setValue(status);
  if (code !== undefined) sh.getRange(row, h.Code + 1).setValue(txt_(code));
}

function ownerSaveRequest_(req) {
  authOwner_(req.token);
  const st = String(req.status || '');
  if (!/^(New|Done|Ignored)$/.test(st)) throw new Error('Unknown status.');
  markRequest_(req.row, st);
  return ownerData_(req);
}

function requestTrial_(req) {
  const company = clip_(req.companyName, 80), name = clip_(req.name, 60), city = clip_(req.city, 60), staff = clip_(req.staff, 20);
  const phone = String(req.phone || '').replace(/[^\d+]/g, '').slice(0, 15);
  if (!company) throw new Error('Enter your company name.');
  if (phone.replace(/\D/g, '').length < 10) throw new Error('Enter a valid mobile number.');
  const c = CacheService.getScriptCache();
  const n = Number(c.get('trialreq') || 0);
  if (n >= 30) throw new Error('Too many requests right now. Please try again in an hour.');
  c.put('trialreq', String(n + 1), 3600);
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = rootSheet_(SH.REQ);
    const last = sh.getLastRow();
    if (last > 1) {
      const h = headerMap_(sh);
      const from = Math.max(2, last - 199);
      const dup = sh.getRange(from, 1, last - from + 1, sh.getLastColumn()).getValues().some(function (r) {
        return String(r[h.Phone]) === phone && String(r[h.Status]) === 'New';
      });
      if (dup) return { message: 'We already have your request. We will send your free trial login on WhatsApp soon.' };
    }
    sh.appendRow([new Date(), txt_(company), txt_(name), txt_(phone), txt_(city), txt_(staff), 'New', '']);
  } finally {
    lock.releaseLock();
  }
  return { message: 'Thank you! We will send your free trial login on WhatsApp soon.' };
}

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
  const punches = punchesSince_(date, function (p) { return p.date === date; }).map(publicPunch_);
  return { date: date, rows: rep.days, punches: punches, alerts: alertsFor_(date), outside: outsideFor_(date, date),
    settings: publicSettings_(), serverTime: Date.now() };
}

function adminSavePunch_(req) {
  authAdmin_(req.token);
  const e = findEmployee_(req.empId || '');
  if (!e) throw new Error('Employee not found.');
  const type = String(req.type || '');
  if (!LABEL[type]) throw new Error('Choose Check In, Lunch Out, Lunch In or Check Out.');
  const date = dateParam_(req.date);
  if (!date) throw new Error('Choose the date.');
  const hm = hhmm_(req.time);
  if (!hm) throw new Error('Enter the time.');
  if (e.joinDate && date < e.joinDate) throw new Error(e.name + ' joined on ' + e.joinDate + '.');
  const t = manualTime_(e, date, hm);
  const now = new Date();
  if (t.getTime() > now.getTime() + 60000) throw new Error('That time has not come yet.');
  const reason = String(req.reason || '').replace(/\s+/g, ' ').trim().slice(0, 100);
  const id = String(req.id || '');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    let list = punchesSince_(date, function (p) { return p.empId === e.id && p.date === date; });
    if (id) {
      const old = list.filter(function (p) { return p.id === id; })[0];
      if (!old || !old.manual) throw new Error('Only punches added by the admin can be changed.');
      list = list.filter(function (p) { return p !== old; });
    }
    checkOrder_(list.concat([{ ts: t, action: type }]));
    const sh = sheet_(SH.PUNCH);
    const c = headerMap_(sh);
    const row = [t, txt_(date), txt_(e.id), txt_(e.name), LABEL[type], '', '', '', txt_('Added by admin' + (reason ? ': ' + reason : '')),
      '', '', '', '', '', '', txt_(fmt_(now, 'yyyy-MM-dd HH:mm:ss')), txt_('Admin dashboard'), txt_('Admin'), '', id || Utilities.getUuid()];
    const at = id ? findPunchRow_(sh, c, id) : sh.getLastRow() + 1;
    if (!at) throw new Error('Punch not found. Refresh and try again.');
    sh.getRange(at, 1, 1, row.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }
  return { saved: { empId: e.id, date: date, type: type, ts: t.getTime() } };
}

function adminDeletePunch_(req) {
  authAdmin_(req.token);
  const id = String(req.id || '');
  if (!id) throw new Error('Punch not found.');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = sheet_(SH.PUNCH);
    const c = headerMap_(sh);
    const at = findPunchRow_(sh, c, id);
    if (!at) throw new Error('Punch not found. Refresh and try again.');
    const r = sh.getRange(at, 1, 1, sh.getLastColumn()).getValues()[0];
    if (String(r[c.PunchedBy]).trim() !== 'Admin') throw new Error('Only punches added by the admin can be deleted.');
    const date = dstr_(r[c.WorkDate]), empId = String(r[c.EmpID]).trim();
    checkOrder_(punchesSince_(date, function (p) { return p.empId === empId && p.date === date && p.id !== id; }));
    sh.deleteRow(at);
  } finally {
    lock.releaseLock();
  }
  return { deleted: id };
}

function manualTime_(e, date, hm) {
  const ws = shift_(e).windowStart;
  let m = toMin_(hm);
  while (m < ws) m += 1440;
  while (m >= ws + 1440) m -= 1440;
  const t = dateAt_(date, m);
  if (workDate_(e, t) !== date) throw new Error('That time does not fit this shift.');
  return t;
}

function checkOrder_(list) {
  list = list.slice().sort(function (a, b) { return a.ts - b.ts; });
  let prev = '';
  list.forEach(function (p) {
    if ((NEXT[prev] || []).indexOf(p.action) < 0) {
      throw new Error('Not possible. The day would read: ' + list.map(function (q) { return LABEL[q.action] + ' ' + fmt_(q.ts, 'HH:mm'); }).join(', ') +
        '. The order must be Check In, Lunch Out, Lunch In, Check Out.');
    }
    prev = p.action;
  });
}

function findPunchRow_(sh, c, id) {
  const last = sh.getLastRow();
  if (last < 2 || c.PunchID === undefined) return 0;
  const v = sh.getRange(2, c.PunchID + 1, last - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]) === id) return i + 2;
  return 0;
}

function adminReport_(req) {
  authAdmin_(req.token);
  const from = dateParam_(req.from), to = dateParam_(req.to);
  if (!from || !to || from > to) throw new Error('Choose a valid date range.');
  return buildReport_(from, to, String(req.empId || ''));
}

function adminEmployees_(req) {
  authAdmin_(req.token);
  forgetEmployees_();
  return {
    employees: employees_().map(function (e) {
      e.hasFace = !!e.faceData;
      delete e.row; delete e.faceData;
      return e;
    }),
    locations: locations_(),
  };
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

  const halfOff = dayName_(x.halfDayOff);
  if (halfOff && String(x.weeklyOff || '').toLowerCase().indexOf(halfOff.toLowerCase()) >= 0) {
    throw new Error('The half-day off cannot be the same day as a weekly off.');
  }
  const join = x.joinDate ? dateParam_(x.joinDate) : '';
  if (x.joinDate && !join) throw new Error('Joining date is invalid.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    forgetEmployees_();
  const sh = sheet_(SH.EMP);
    const col = headerMap_(sh);
    const existing = findEmployee_(id);
    if (x.isNew && existing) throw new Error('Employee ID ' + id + ' already exists.');
    const vals = {
      EmpID: txt_(existing ? existing.id : id), Name: txt_(name), PIN: txt_(pin), Phone: txt_(String(x.phone || '').trim()),
      Department: txt_(String(x.dept || '').trim()), ShiftStart: txt_(ss), ShiftEnd: txt_(se), LunchMinutes: lunch,
      LunchStart: txt_(hhmm_(x.lunchStart) || ''),
      HalfDayOff: txt_(halfOff),
      GraceMinutes: grace, WeeklyOff: txt_(String(x.weeklyOff || '').trim()),
      AllowedLocations: txt_(String(x.locations || '').trim()), Active: x.active === false ? 'No' : 'Yes',
      CanPunchOthers: x.canPunchOthers ? 'Yes' : 'No',
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
  try { cache_().remove('loc_v1'); } catch (e) { }
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
  try { cache_().remove('loc_v1'); } catch (e) { }
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
  if (x.faceMode !== undefined) setSetting_('FACE_MODE', /^(WARN|REQUIRED)$/.test(x.faceMode) ? x.faceMode : 'OFF');
  const okV = function (v) { return /^(SELFIE|FACE|BOTH)$/.test(v) ? v : 'SELFIE'; };
  if (x.verifyInOut !== undefined) setSetting_('VERIFY_IN_OUT', okV(x.verifyInOut));
  if (x.verifyLunch !== undefined) setSetting_('VERIFY_LUNCH', okV(x.verifyLunch));
  if (x.minHalfDay !== undefined) setSetting_('MIN_HALF_DAY_HOURS', String(Math.min(12, Math.max(0, num_(x.minHalfDay, 3)))));
  if (x.lunchStart !== undefined) setSetting_('LUNCH_START', hhmm_(x.lunchStart) || '');
  if (x.awayMinutes !== undefined) setSetting_('AWAY_MINUTES', String(Math.min(240, Math.max(1, Math.round(num_(x.awayMinutes, 15))))));
  if (x.faceStrict !== undefined) setSetting_('FACE_MODE', x.faceStrict ? 'REQUIRED' : 'WARN');
  if (x.lunchMode !== undefined) setSetting_('LUNCH_MODE', x.lunchMode === 'AUTO' ? 'AUTO' : 'SELFIE');
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
  const all = punchesSince_(from, function (p) { return p.date <= to && (!empId || p.empId === empId); });
  const byKey = {};
  all.forEach(function (p) { (byKey[p.empId + '|' + p.date] = byKey[p.empId + '|' + p.date] || []).push(p); });
  const emps = employees_().filter(function (e) {
    if (empId && e.id !== empId) return false;
    if (e.joinDate && e.joinDate > to) return false;
    return e.active || all.some(function (p) { return p.empId === e.id; });
  });

  const outByKey = {};
  outsideFor_(from, to).forEach(function (o) { (outByKey[o.empId.toUpperCase() + '|' + o.date] = outByKey[o.empId.toUpperCase() + '|' + o.date] || []).push(o); });
  const days = [], totals = [];
  emps.forEach(function (e) {
    const t = { id: e.id, name: e.name, dept: e.dept, shift: e.shiftStart + '-' + e.shiftEnd, present: 0, halfDay: 0,
      absent: 0, late: 0, lateMin: 0, earlyMin: 0, workedMin: 0, otMin: 0, lunchOverMin: 0, off: 0, missedOut: 0,
      outsideCount: 0, outsideMin: 0, manual: 0 };
    dates.forEach(function (d) {
      if (e.joinDate && d < e.joinDate) return;
      const s = summarize_(e, d, byKey[e.id + '|' + d] || [], now, hol);
      s.id = e.id; s.name = e.name; s.dept = e.dept;
      s.manual = (byKey[e.id + '|' + d] || []).filter(function (p) { return p.manual; }).length;
      s.lunchStart = e.lunchStart || defaultLunchStart_(); s.lunchMin = num_(e.lunch, 0);
      t.manual += s.manual;
      const outs = outByKey[e.id.toUpperCase() + '|' + d] || [];
      s.outsideCount = outs.length;
      s.outsideMin = outs.reduce(function (a, o) { return a + o.minutes; }, 0);
      s.outsideNow = outs.some(function (o) { return !o.end; });
      const ls = e.lunchStart || defaultLunchStart_();
      if (ls && outs.length) {
        const lw = dateAt_(d, toMin_(ls)).getTime(), le = lw + num_(e.lunch, 0) * 60000;
        s.outsideLunchMin = outs.reduce(function (a, o) {
          return a + Math.max(0, Math.round((Math.min(o.end || Date.now(), le) - Math.max(o.start, lw)) / 60000));
        }, 0);
      } else s.outsideLunchMin = 0;
      t.outsideCount += s.outsideCount;
      t.outsideMin += s.outsideMin;
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

function verifyFor_(type) {
  const key = /^LUNCH/.test(type) ? 'VERIFY_LUNCH' : 'VERIFY_IN_OUT';
  const v = String(settings_()[key] || '').toUpperCase();
  if (v === 'SELFIE' || v === 'FACE' || v === 'BOTH') return v;
  return rawFaceMode_() === 'OFF' ? 'SELFIE' : 'BOTH';
}

function rawFaceMode_() {
  const m = String(settings_().FACE_MODE || 'OFF').toUpperCase();
  return m === 'WARN' || m === 'REQUIRED' ? m : 'OFF';
}

function faceStrict_() {
  return rawFaceMode_() === 'REQUIRED';
}

function faceMode_() {
  const any = ['IN', 'LUNCH_OUT'].some(function (t) { return verifyFor_(t) !== 'SELFIE'; });
  return !any ? 'OFF' : faceStrict_() ? 'REQUIRED' : 'WARN';
}

function cleanFaces_(faces) {
  if (!Array.isArray(faces)) return [];
  return faces.slice(0, 60).filter(function (f) {
    return Array.isArray(f) && f.length === 128 && f.every(function (v) { return typeof v === 'number' && isFinite(v) && Math.abs(v) < 2; });
  });
}

function parseFace_(v) {
  const s = String(v || '').trim();
  if (!s || s[0] !== '[') return null;
  try { const f = cleanFaces_([JSON.parse(s)]); return f.length ? f[0] : null; } catch (e) { return null; }
}

function faceDist_(a, b) {
  let sum = 0;
  for (let i = 0; i < 128; i++) { const d = a[i] - b[i]; sum += d * d; }
  return Math.sqrt(sum);
}

function facePct_(d, thr) {
  const p = d <= thr ? 100 - (d / thr) * 30 : 70 - ((d - thr) / thr) * 140;
  return Math.max(0, Math.min(100, Math.round(p)));
}

function checkFaces_(targets, faces, forOthers, faceError, method, hasPhoto) {
  const out = { results: targets.map(function () { return ''; }), enroll: {} };
  if (method === 'SELFIE') return out;
  const strict = method === 'FACE' || faceStrict_();
  if (!faces.length && faceError && hasPhoto) {
    out.results = targets.map(function () { return 'Face check failed'; });
    return out;
  }
  if (method === 'FACE') {
    const unreg = targets.filter(function (t) { return !t.faceData; });
    if (unreg.length && !(hasPhoto && targets.length === 1 && faces.length === 1)) {
      throw new Error(names_(unreg) + (unreg.length > 1 ? ' have' : ' has') + ' no registered face yet. Take a selfie once, or ask the admin to register the face.');
    }
  }
  const thr = num_(settings_().FACE_THRESHOLD, 0.5);
  const used = {};
  targets.forEach(function (t, i) {
    if (!faces.length) { out.results[i] = faceError ? 'Face check failed' : 'No face'; return; }
    if (!t.faceData) {
      if (targets.length === 1 && faces.length === 1) { out.enroll[t.id] = faces[0]; out.results[i] = 'Registered'; }
      else out.results[i] = 'Not registered';
      return;
    }
    let best = 99, bj = -1;
    faces.forEach(function (f, j) {
      if (used[j]) return;
      const d = faceDist_(f, t.faceData);
      if (d < best) { best = d; bj = j; }
    });
    if (bj >= 0 && best <= thr) { used[bj] = true; out.results[i] = 'Match ' + facePct_(best, thr) + '%'; }
    else out.results[i] = 'NO MATCH' + (bj >= 0 ? ' ' + facePct_(best, thr) + '%' : '');
  });
  if (strict) {
    if (!faces.length) {
      throw new Error(faceError ? 'Face check could not start. Check the internet and try again.' : 'No face found in the photo. Look straight at the camera in good light.');
    }
    const bad = targets.filter(function (t, i) { return /^NO MATCH/.test(out.results[i]); });
    if (bad.length) {
      throw new Error(forOthers
        ? 'Face does not match for ' + names_(bad) + '. Make sure everyone selected is in the photo.'
        : 'Face does not match your registered face. Try again in good light, or ask the admin to reset your face.');
    }
  }
  return out;
}

function saveFace_(id, desc, how) {
  forgetEmployees_();
  const sh = sheet_(SH.EMP);
  const col = headerMap_(sh);
  const e = findEmployee_(id);
  if (!e || col.FaceData === undefined) return;
  sh.getRange(e.row, col.FaceData + 1).setValue(JSON.stringify(desc.map(function (v) { return Math.round(v * 1e4) / 1e4; })));
  if (col.FaceRegistered !== undefined) sh.getRange(e.row, col.FaceRegistered + 1).setValue(txt_(fmt_(new Date(), 'yyyy-MM-dd HH:mm') + ' (' + how + ')'));
  forgetEmployees_();
}

function health_() {
  const emps = employees_();
  return {
    version: VERSION,
    employees: emps.length,
    active: emps.filter(function (e) { return e.active; }).length,
    facesRegistered: emps.filter(function (e) { return !!e.faceData; }).length,
    faceMode: faceMode_(),
    verifyInOut: verifyFor_('IN'),
    verifyLunch: verifyFor_('LUNCH_OUT'),
    blockOutside: enforceFence_(),
    locations: locations_().map(function (l) { return l.name + ' (' + l.radius + ' m)'; }),
    lunchStart: defaultLunchStart_(),
  };
}

function faceRefs_(req) {
  const me = authEmp_(req.token);
  const refs = {};
  if (faceMode_() !== 'OFF') {
    (me.canPunch ? employees_().filter(function (e) { return e.active; }) : [me]).forEach(function (e) {
      if (e.faceData) refs[e.id] = { name: e.name, d: e.faceData };
    });
  }
  return { refs: refs, threshold: num_(settings_().FACE_THRESHOLD, 0.5), mode: faceMode_() };
}

function adminSaveFace_(req) {
  authAdmin_(req.token);
  const f = cleanFaces_([req.descriptor]);
  if (!f.length) throw new Error('No face data. Take the photo again.');
  if (!findEmployee_(req.id)) throw new Error('Employee not found. Save the employee first.');
  saveFace_(req.id, f[0], 'admin');
  return adminEmployees_(req);
}

function adminResetFace_(req) {
  authAdmin_(req.token);
  forgetEmployees_();
  const e = findEmployee_(req.id);
  if (!e) throw new Error('Employee not found.');
  forgetEmployees_();
  const sh = sheet_(SH.EMP);
  const col = headerMap_(sh);
  if (col.FaceData !== undefined) sh.getRange(e.row, col.FaceData + 1).setValue('');
  if (col.FaceRegistered !== undefined) sh.getRange(e.row, col.FaceRegistered + 1).setValue('');
  return adminEmployees_(req);
}

function shift_(e) {
  const s = toMin_(e.shiftStart), en = toMin_(e.shiftEnd);
  let len = (en - s + 1440) % 1440;
  if (len === 0) len = 1440;
  return { start: s, end: en, len: len, windowStart: s - (1440 - len) / 2 };
}

function workDate_(e, t) {
  return fmt_(new Date(t.getTime() - shift_(e).windowStart * 60000), 'yyyy-MM-dd');
}

function summarize_(e, date, list, now, hol) {
  const sh = shift_(e);
  const schedStart = dateAt_(date, sh.start);
  const halfOff = !hol[date] && !isWeeklyOff_(e, date) && isHalfDayOff_(e, date);
  const dayLen = halfOff ? Math.round(sh.len / 2) : sh.len;
  const schedEnd = new Date(schedStart.getTime() + dayLen * 60000);
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
  let autoLunch = false;
  if (!lunchOut && first && endT && autoLunch_() && (endT - first.ts) / 60000 > AUTO_LUNCH_AFTER_MIN) {
    breakMs = num_(e.lunch, 0) * 60000;
    autoLunch = true;
  }
  const workedMin = first && endT ? Math.max(0, Math.round((endT - first.ts - breakMs) / 60000)) : 0;
  const fullExpected = Math.max(0, sh.len - num_(e.lunch, 0));
  const expected = halfOff ? Math.round(fullExpected / 2) : fullExpected;
  const lateMin = first ? Math.max(0, Math.round((first.ts - schedStart) / 60000)) : 0;
  const late = first && !off && lateMin > num_(e.grace, 0) ? lateMin : 0;
  const early = lastOut && !off ? Math.max(0, Math.round((schedEnd - lastOut.ts) / 60000)) : 0;
  const breakMin = Math.round(breakMs / 60000);
  const ot = lastOut ? (off ? workedMin : Math.max(0, workedMin - expected)) : 0;

  let status, short = false;
  if (!first) {
    if (off) status = holiday ? 'Holiday' : 'Weekly Off';
    else if (now < schedStart) status = 'Upcoming';
    else if (now < schedEnd) status = 'Not In Yet';
    else status = 'Absent';
  } else if (!lastOut) {
    status = current ? (onBreak ? 'On Lunch' : 'Working') : 'No Check-out';
  } else {
    const minHalf = num_(settings_().MIN_HALF_DAY_HOURS, 3) * 60;
    const fullAt = halfOff ? Math.min(minHalf, expected) : expected / 2;
    if (off || workedMin >= fullAt) status = 'Present';
    else if (!halfOff && workedMin >= minHalf) status = 'Half Day';
    else { status = 'Absent'; short = true; }
  }

  return {
    date: date, status: status, offDay: off, holiday: holiday, halfDayOff: halfOff,
    shift: e.shiftStart + '-' + e.shiftEnd,
    in: first ? first.ts.getTime() : null,
    out: lastOut ? lastOut.ts.getTime() : null,
    lunchOut: lunchOut ? lunchOut.getTime() : null,
    lunchIn: lunchIn ? lunchIn.getTime() : null,
    workedMin: workedMin, expectedMin: expected, late: late, early: early,
    breakMin: breakMin, lunchOver: Math.max(0, breakMin - num_(e.lunch, 0)), ot: ot, autoLunch: autoLunch, short: short,
  };
}

function dayName_(v) {
  const d = String(v || '').trim().slice(0, 3).toLowerCase();
  const all = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  for (let i = 0; i < all.length; i++) if (all[i].toLowerCase() === d) return all[i];
  return '';
}

function isHalfDayOff_(e, date) {
  return !!e.halfDayOff && fmt_(dateAt_(date, 720), 'EEE').toLowerCase() === e.halfDayOff.toLowerCase();
}

function halfShiftEnd_(e) {
  const sh = shift_(e);
  const m = (sh.start + Math.round(sh.len / 2)) % 1440;
  return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2);
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
  const cache = cache_();
  const hit = cache.get(key);
  if (hit !== null) return hit;
  let a = '';
  try {
    const r = Maps.newGeocoder().reverseGeocode(lat, lng);
    a = (r && r.results && r.results[0] && r.results[0].formatted_address) || '';
  } catch (err) {   }
  if (a) cache.put(key, a, 21600);
  return a;
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" is missing. Run setup() in the script editor.');
  return sh;
}

function headerMap_(sh) {
  const m = {};
  sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].forEach(function (h, i) { m[String(h).trim()] = i; });
  return m;
}

let EMPLOYEES_MEMO = null;

function employees_() {
  if (!EMPLOYEES_MEMO) EMPLOYEES_MEMO = cacheGet_('emp_v1');
  if (!EMPLOYEES_MEMO) {
    EMPLOYEES_MEMO = readEmployees_();
    cachePut_('emp_v1', EMPLOYEES_MEMO);
  }
  return EMPLOYEES_MEMO.map(function (e) { return Object.assign({}, e); });
}

function readEmployees_() {
  EMPLOYEES_MEMO = null;
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
      canPunch: /^(yes|y|true|1)$/i.test(String(g(r, 'CanPunchOthers')).trim()),
      faceData: parseFace_(g(r, 'FaceData')),
      faceReg: String(g(r, 'FaceRegistered')).trim(),
      lunchStart: hhmm_(g(r, 'LunchStart')),
      halfDayOff: dayName_(g(r, 'HalfDayOff')),
    });
  }
  return out;
}

function findEmployee_(id) {
  id = String(id).trim().toUpperCase();
  return employees_().filter(function (e) { return e.id.toUpperCase() === id; })[0] || null;
}

function punchesSince_(from, keep) {
  const sh = sheet_(SH.PUNCH);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const c = headerMap_(sh);
  const width = sh.getLastColumn();
  const stop = addDays_(from, -2);
  let rows = [], end = last, chunk = 600;
  while (end >= 2) {
    const start = Math.max(2, end - chunk + 1);
    const v = sh.getRange(start, 1, end - start + 1, width).getValues();
    rows = v.concat(rows);
    end = start - 1;
    chunk *= 2;
    let earliest = '';
    for (let i = 0; i < v.length && !earliest; i++) {
      if (c.PunchedBy === undefined || String(v[i][c.PunchedBy]).trim() !== 'Admin') earliest = dstr_(v[i][c.WorkDate]);
    }
    if (earliest && earliest < stop) break;
  }
  return rowsToPunches_(rows, c, function (q) { return q.date >= from && keep(q); });
}

function punches_(keep) {
  const sh = sheet_(SH.PUNCH);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const c = headerMap_(sh);
  return rowsToPunches_(sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues(), c, keep);
}

function rowsToPunches_(v, c, keep) {
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
      thumbId: String(g(r, 'ThumbID')), by: String(g(r, 'PunchedBy')), face: String(g(r, 'FaceMatch')),
      photoUrl: String(g(r, 'PhotoURL')), id: String(g(r, 'PunchID')), devTime: String(g(r, 'DeviceTime')),
    };
    p.manual = p.by.trim() === 'Admin';
    if (keep(p)) out.push(p);
  });
  out.sort(function (a, b) { return a.ts - b.ts; });
  return out;
}

function publicPunch_(p) {
  return {
    ts: p.ts.getTime(), date: p.date, empId: p.empId, name: p.name, action: p.action, label: LABEL[p.action] || p.action,
    lat: p.lat, lng: p.lng, acc: p.acc, address: p.address, location: p.location, inside: p.inside,
    distance: p.distance, photoId: p.photoId, thumbId: p.thumbId, by: p.by, face: p.face,
    photoPending: p.photoUrl === 'pending' && !p.photoId, id: p.id, manual: p.manual, addedAt: p.manual ? p.devTime : '',
  };
}

function actionCode_(v) {
  v = String(v).trim();
  if (LABEL[v]) return v;
  for (const k in LABEL) if (LABEL[k].toLowerCase() === v.toLowerCase()) return k;
  return v;
}

function locations_() {
  const cached = cacheGet_('loc_v1');
  if (cached) return cached;
  const out = readLocations_();
  cachePut_('loc_v1', out);
  return out;
}

function readLocations_() {
  const v = sheet_(SH.LOC).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const name = String(v[i][0]).trim(), lat = Number(v[i][1]), lng = Number(v[i][2]);
    if (name && isFinite(lat) && isFinite(lng) && v[i][1] !== '') out.push({ name: name, lat: lat, lng: lng, radius: num_(v[i][3], 100) });
  }
  return out;
}

function holidays_() {
  const cached = cacheGet_('hol_v1');
  if (cached) return cached;
  const sh = ss_().getSheetByName(SH.HOL);
  const out = {};
  if (!sh) return out;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    const d = dstr_(r[0]);
    if (d) out[d] = String(r[1] || 'Holiday');
  });
  cachePut_('hol_v1', out);
  return out;
}

let SETTINGS_MEMO = null;

function settings_() {
  if (SETTINGS_MEMO) return SETTINGS_MEMO;
  const cached = cacheGet_('set_v1');
  if (cached) { SETTINGS_MEMO = cached; return cached; }
  const sh = ss_().getSheetByName(SH.SET);
  const out = {};
  if (!sh) return out;
  SETTINGS_MEMO = out;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    if (String(r[0]).trim()) out[String(r[0]).trim()] = String(r[1]).trim();
  });
  cachePut_('set_v1', out);
  return out;
}

function setSetting_(key, value) {
  SETTINGS_MEMO = null;
  try { cache_().remove('set_v1'); } catch (e) { }
  const sh = sheet_(SH.SET);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === key) { sh.getRange(i + 1, 2).setValue(txt_(value)); return; }
  }
  sh.appendRow([key, txt_(value), '']);
}

function publicSettings_() {
  const s = settings_();
  return { company: s.COMPANY_NAME || '', enforceGeofence: enforceFence_(), maxAccuracy: num_(s.MAX_ACCURACY_M, 0), tz: tz_(),
    lunchMode: autoLunch_() ? 'AUTO' : 'SELFIE', faceMode: faceMode_(), faceThreshold: num_(s.FACE_THRESHOLD, 0.5),
    verifyInOut: verifyFor_('IN'), verifyLunch: verifyFor_('LUNCH_OUT'), faceStrict: faceStrict_(),
    lunchStart: defaultLunchStart_(), awayMinutes: num_(s.AWAY_MINUTES, 15), minHalfDay: num_(s.MIN_HALF_DAY_HOURS, 3),
    companyCode: CO ? CO.code : '' };
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
  if (CO) payload = Object.assign({ c: CO.code }, payload);
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
  if (p.r !== role) throw new Error('AUTH: Please log in again.');
  if (!(p.exp > Date.now())) throw new Error('AUTH: Session expired. Please log in again.');
  if ((p.c || '') !== (CO ? CO.code : '')) throw new Error('AUTH: Please log in again.');
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
  const n = Number(cache_().get('fail:' + key) || 0);
  if (n >= 5) throw new Error('Too many wrong attempts. Try again in 15 minutes.');
}

function failed_(key) {
  const c = cache_();
  c.put('fail:' + key, String(Number(c.get('fail:' + key) || 0) + 1), 900);
}

function tz_() {
  return ss_().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
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
  m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/.exec(s);
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

function txt_(v) {
  v = String(v === undefined || v === null ? '' : v);
  return v === '' ? '' : "'" + v;
}
