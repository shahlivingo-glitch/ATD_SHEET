// Admin dashboard: daily board with selfies, reports + CSV, employees, locations, settings.

const A = {
  token: store.get('att_admin_token'),
  tab: 'day',
  day: null,
  employees: [],
  locations: [],
  report: null,
  thumbs: {},          // thumbId -> data URL (null = unavailable)
  punchesById: {},
};

const DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

function show(id) { ['v-login', 'v-main'].forEach(function (v) { $('#' + v).classList.toggle('hidden', v !== id); }); }

async function call(action, data) {
  try {
    return await api(action, Object.assign({ token: A.token }, data || {}));
  } catch (e) {
    if (e.auth) logout(e.message);
    throw e;
  }
}

function logout(msg) {
  A.token = null;
  store.del('att_admin_token');
  show('v-login');
  if (msg) toast(msg, true);
}

$('#login-form').addEventListener('submit', async function (ev) {
  ev.preventDefault();
  const btn = this.querySelector('button');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>';
  try {
    if ($('#co-code')) setCompany($('#co-code').value);
    const r = await api('adminLogin', { password: $('#admin-pw').value });
    A.token = r.token;
    store.set('att_admin_token', r.token);
    setTZ(r.tz);
    store.set('att_tz', r.tz);
    $('#admin-pw').value = '';
    start();
  } catch (e) {
    toast(e.message, true);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Log in';
  }
});

$('#btn-refresh').innerHTML = ICONS.refresh;
$('#btn-logout').innerHTML = ICONS.logout;
document.querySelectorAll('[data-close]').forEach(function (b) {
  b.innerHTML = ICONS.close;
  b.addEventListener('click', function () { b.closest('.modal').classList.add('hidden'); });
});
document.querySelectorAll('.modal').forEach(function (m) {
  m.addEventListener('click', function (ev) { if (ev.target === m) m.classList.add('hidden'); });
});
$('#btn-logout').addEventListener('click', function () { logout(); });
$('#btn-refresh').addEventListener('click', function () { loadTab(true); });

$('#tabs').addEventListener('click', function (ev) {
  const b = ev.target.closest('button');
  if (!b) return;
  A.tab = b.dataset.tab;
  document.querySelectorAll('#tabs button').forEach(function (x) { x.classList.toggle('on', x === b); });
  ['day', 'report', 'emps', 'locs'].forEach(function (t) { $('#tab-' + t).classList.toggle('hidden', t !== A.tab); });
  store.set('att_admin_tab', A.tab);
  loadTab();
});

function loadTab(force) {
  if (A.tab === 'day') return loadDay();
  if (A.tab === 'locs') return loadEmployees().then(loadSettingsForm);
  if (force || !A.employees.length) loadEmployees(); else renderEmployees();
  if (A.tab === 'report' && force && A.report) runReport();
}

/* ---------------- daily board ---------------- */

async function loadDay() {
  const date = $('#day-date').value || ymd(Date.now());
  $('#day-date').value = date;
  if (!A.day) $('#day-list').innerHTML = '<div class="empty"><span class="spinner"></span></div>';
  try {
    const r = await call('adminDay', { date: date });
    A.day = r;
    A.settings = r.settings;
    setTZ(r.settings.tz);
    $('#h-company').textContent = (r.settings.company || 'Attendance') + ' - Admin';
    renderDay();
    $('#day-updated').textContent = 'Updated ' + fmtTime(Date.now(), true) + ' · refreshes every minute';
  } catch (e) {
    toast(e.message, true);
  }
}

function renderDay() {
  const r = A.day;
  if (!r) return;
  const byEmp = {};
  A.punchesById = {};
  r.punches.forEach(function (p) {
    (byEmp[p.empId] = byEmp[p.empId] || []).push(p);
    p.key = p.empId + '_' + p.ts + '_' + p.action;
    A.punchesById[p.key] = p;
  });

  const rows = r.rows;
  const count = function (f) { return rows.filter(f).length; };
  const working = ['Working', 'On Lunch'];
  const k = [
    ['Employees', rows.length],
    ['Present', count(function (x) { return x.in; })],
    ['Working now', count(function (x) { return working.indexOf(x.status) >= 0; })],
    ['On lunch', count(function (x) { return x.status === 'On Lunch'; })],
    ['Late', count(function (x) { return x.late; })],
    ['Checked out', count(function (x) { return x.out; })],
    ['Absent / not in', count(function (x) { return !x.in && !x.offDay && x.status !== 'Upcoming'; })],
  ];
  $('#kpis').innerHTML = k.map(function (x) { return '<div class="kpi"><b>' + x[1] + '</b><span>' + x[0] + '</span></div>'; }).join('');

  const filter = $('#day-filter').value;
  const q = $('#day-search').value.trim().toLowerCase();
  const list = rows.filter(function (x) {
    if (q && (x.name + ' ' + x.id + ' ' + x.dept).toLowerCase().indexOf(q) < 0) return false;
    if (filter === 'in') return !!x.in;
    if (filter === 'late') return !!x.late;
    if (filter === 'lunch') return x.status === 'On Lunch';
    if (filter === 'missing') return !x.in && !x.offDay;
    return true;
  });

  const need = [];
  $('#day-list').innerHTML = list.length ? list.map(function (x) {
    const ps = byEmp[x.id] || [];
    const slot = function (type, pick) {
      const arr = ps.filter(function (p) { return p.action === type; });
      const p = pick === 'last' ? arr[arr.length - 1] : arr[0];
      if (p && p.thumbId && !(p.thumbId in A.thumbs)) need.push(p.thumbId);
      const th = p && p.thumbId && A.thumbs[p.thumbId] ? ' style="background-image:url(' + A.thumbs[p.thumbId] + ')"' : '';
      const flag = p && p.inside === 'No' ? ' <span title="Outside location" style="color:var(--out)">!</span>' : '';
      // no photo: blue for a punch the admin added, green for a face-recognition punch,
      // red OUT when the person is absent / not in yet; an empty slot opens "Add missed punch"
      let mark = '', cls = 'th';
      if (p && !p.photoId && p.photoPending) { cls += ' wait'; mark = 'Photo uploading'; }
      else if (p && p.manual) { cls += ' manual'; mark = 'ADMIN'; }
      else if (p && !p.photoId) { cls += ' face'; mark = type === 'IN' ? 'IN' : '&#10003;'; }
      else if (!p && type === 'IN' && (x.status === 'Absent' || x.status === 'Not In Yet')) { cls += ' absent'; mark = 'OUT'; }
      else if (!p) { cls += ' add'; mark = '+'; }
      return '<button class="p has"' + (p ? ' data-photo="' + esc(p.key) + '"' : ' data-add="' + esc(x.id + '|' + type) + '" title="Add missed ' + LABELS[type] + '"') + '>' +
        '<div class="' + cls + '" data-thumb="' + esc(p ? p.thumbId : '') + '"' + th + '>' + mark + '</div>' +
        '<b>' + (p ? fmtTime(p.ts) : '--') + flag + '</b>' + LABELS[type] + '</button>';
    };
    const meta = [];
    if (x.outsideCount) {
      const txt = x.outsideNow
        ? 'Outside now · ' + x.outsideCount + '× ' + dur(x.outsideMin)
        : 'Outside ' + x.outsideCount + '× · ' + dur(x.outsideMin) + (x.outsideLunchMin ? ' (' + dur(x.outsideLunchMin) + ' in lunch hour)' : '');
      meta.push('<button class="pill ' + (x.outsideNow ? 'bad' : 'warn') + '" data-outside="' + esc(x.id) + '" style="border:0;cursor:pointer">' + esc(txt) + '</button>');
    }
    if (x.halfDayOff) meta.push('<span class="pill info">Half-day off</span>');
    if (x.short) meta.push('<span class="pill bad">Worked under ' + ((A.settings && A.settings.minHalfDay) || 3) + 'h</span>');
    if (x.late) meta.push('<span class="pill warn">Late ' + dur(x.late) + '</span>');
    if (x.early) meta.push('<span class="pill warn">Left ' + dur(x.early) + ' early</span>');
    if (x.autoLunch) meta.push('<span class="pill mute">Lunch ' + dur(x.breakMin) + ' auto</span>');
    if (x.lunchOver) meta.push('<span class="pill warn">Lunch +' + dur(x.lunchOver) + '</span>');
    if (x.ot) meta.push('<span class="pill info">OT ' + dur(x.ot) + '</span>');
    if (x.in) meta.push('<span class="pill mute">Worked ' + dur(x.workedMin) + '</span>');
    (r.alerts || []).filter(function (a) { return a.empId === x.id; }).forEach(function (a) {
      meta.push('<span class="pill bad" title="' + esc(a.detail) + '">' + esc(a.type) + ' · ' + fmtTime(a.ts) + '</span>');
    });
    const badFace = ps.filter(function (p) { return /^NO MATCH|^No face|^Face check failed/.test(p.face || ''); }).length;
    if (badFace) meta.push('<span class="pill bad">Face mismatch ' + badFace + '</span>');
    if (ps.some(function (p) { return p.by && !p.manual; })) meta.push('<span class="pill info">Punched on shared phone</span>');
    const manual = ps.filter(function (p) { return p.manual; }).length;
    if (manual) meta.push('<span class="pill info">Added by admin ' + manual + '</span>');
    if (ps.length > 4) meta.push('<span class="pill mute">' + ps.length + ' punches</span>');
    meta.push('<button class="pill mute" data-add="' + esc(x.id + '|') + '" style="border:0;cursor:pointer">+ Add punch</button>');
    return '<div class="emp-card"><div class="head"><div class="who"><b>' + esc(x.name) + '</b>' +
      '<span class="muted small">' + esc(x.id) + (x.dept ? ' · ' + esc(x.dept) : '') + ' · ' + esc(shiftText(x.shift)) + '</span></div>' +
      statusPill(x.status) + '</div>' +
      '<div class="punch-strip">' + slot('IN') + slot('LUNCH_OUT') + slot('LUNCH_IN', 'last') + slot('OUT', 'last') + '</div>' +
      (meta.length ? '<div class="meta">' + meta.join('') + '</div>' : '') + '</div>';
  }).join('') : '<div class="empty">Nobody matches.</div>';

  fetchThumbs(need);
}

function shiftText(s) {
  const p = String(s || '').split('-');
  return p.length === 2 ? fmtHM(p[0]) + ' - ' + fmtHM(p[1]) : s;
}

let thumbBusy = false;
async function fetchThumbs(ids) {
  ids = ids.filter(function (id, i) { return id && ids.indexOf(id) === i && !(id in A.thumbs); });
  if (!ids.length || thumbBusy) return;
  thumbBusy = true;
  try {
    for (let i = 0; i < ids.length; i += 30) {
      const r = await call('adminThumbs', { ids: ids.slice(i, i + 30) });
      Object.keys(r.thumbs).forEach(function (id) {
        A.thumbs[id] = r.thumbs[id];
        document.querySelectorAll('[data-thumb="' + id + '"]').forEach(function (el) {
          if (r.thumbs[id]) el.style.backgroundImage = 'url(' + r.thumbs[id] + ')';
        });
      });
    }
  } catch (e) { /* thumbnails are optional */ } finally { thumbBusy = false; }
}

function openOutside(empId) {
  const r = A.day;
  const row = (r.rows || []).find(function (x) { return x.id === empId; });
  const list = (r.outside || []).filter(function (o) { return o.empId === empId; }).sort(function (a, b) { return a.start - b.start; });
  $('#out-title').textContent = (row ? row.name : empId) + ' · outside the work location';
  $('#out-list').innerHTML = list.length ? '<div class="table-wrap" style="box-shadow:none"><table><thead><tr><th>Left</th><th>Back</th>' +
    '<th class="num">Minutes</th><th class="num">Farthest</th></tr></thead><tbody>' + list.map(function (o) {
      return '<tr><td>' + fmtTime(o.start) + '</td><td>' + (o.end ? fmtTime(o.end) : '<span class="pill bad">still outside</span>') +
        '</td><td class="num">' + o.minutes + '</td><td class="num">' + (o.maxDistance ? o.maxDistance + ' m' : '') + '</td></tr>';
    }).join('') + '</tbody></table></div>' +
    '<p class="muted small" style="margin:10px 0 0">Recorded by the Android app while the employee is working (Check In to Check Out, not during a punched lunch break). Trips under 2 minutes are ignored.</p>'
    : '<div class="empty">No trips outside.</div>';
  $('#outside-modal').classList.remove('hidden');
}

$('#day-list').addEventListener('click', function (ev) {
  const ob = ev.target.closest('[data-outside]');
  if (ob) { openOutside(ob.dataset.outside); return; }
  const ad = ev.target.closest('[data-add]');
  if (ad) { const k = ad.dataset.add.split('|'); openManual(k[0], k[1]); return; }
  const b = ev.target.closest('[data-photo]');
  if (b) openPhoto(A.punchesById[b.dataset.photo]);
});
$('#day-date').addEventListener('change', function () { A.day = null; loadDay(); });
$('#day-filter').addEventListener('change', renderDay);
$('#day-search').addEventListener('input', renderDay);

async function openPhoto(p) {
  if (!p) return;
  $('#ph-title').textContent = p.name + ' · ' + p.label;
  const img = $('#ph-img');
  img.removeAttribute('src');
  img.classList.toggle('hidden', !p.photoId);
  if (p.manual) {
    const why = manualReason(p);
    $('#ph-info').innerHTML = '<div><span class="pill info">Added by admin, no selfie</span></div>' +
      '<div><b>' + fmtDate(p.ts, true) + ', ' + fmtTime(p.ts) + '</b></div>' +
      (why ? '<div>Reason: ' + esc(why) + '</div>' : '') +
      (p.addedAt ? '<div class="muted small">Added on ' + esc(p.addedAt) + '</div>' : '') +
      '<div><button type="button" class="btn sm ghost" id="ph-change">Change time or delete</button></div>';
    $('#ph-change').addEventListener('click', function () { $('#photo-modal').classList.add('hidden'); openManual(p.empId, p.action, p); });
    $('#photo-modal').classList.remove('hidden');
    return;
  }
  if (p.thumbId && A.thumbs[p.thumbId]) img.src = A.thumbs[p.thumbId];
  const map = 'https://www.google.com/maps?q=' + p.lat + ',' + p.lng;
  const fence = p.location
    ? (p.inside === 'No' ? '<span class="pill bad">Outside: ' + p.distance + ' m from ' + esc(p.location) + '</span>' : '<span class="pill ok">At ' + esc(p.location) + '</span>')
    : '';
  $('#ph-info').innerHTML =
    (p.photoId ? '' : p.photoPending
      ? '<div><span class="pill warn">Photo still uploading from the phone</span></div>'
      : '<div><span class="pill info">Face recognition only, no photo saved</span></div>') +
    '<div><b>' + fmtDate(p.ts, true) + ', ' + fmtTime(p.ts, true) + '</b></div>' +
    '<div>' + esc(p.address || 'No address') + '</div>' +
    '<div class="muted small">' + p.lat + ', ' + p.lng + ' · ±' + esc(p.acc) + ' m GPS</div>' +
    (p.by ? '<div><span class="pill warn">Taken by ' + esc(p.by) + '</span></div>' : '') +
    (p.face ? '<div><span class="pill ' + (/^Match|^Registered/.test(p.face) ? 'ok' : /^NO|^No face|^Face check failed/.test(p.face) ? 'bad' : 'mute') + '">Face: ' + esc(p.face) + '</span></div>' : '') +
    '<div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">' + fence +
    '<a class="btn sm ghost" href="' + map + '" target="_blank" rel="noopener">Open in Maps</a></div>';
  $('#photo-modal').classList.remove('hidden');
  if (!p.photoId) return;
  try {
    const r = await call('adminPhoto', { id: p.photoId });
    img.src = r.photo;
  } catch (e) { toast('Could not load the photo: ' + e.message, true); }
}

/* ---------------- missed punch added by the admin ---------------- */

function hm24(ms) {
  return new Intl.DateTimeFormat('en-GB', { timeZone: TZ, hour: '2-digit', minute: '2-digit', hourCycle: 'h23' }).format(new Date(ms));
}

function manualReason(p) { return String(p.address || '').replace(/^Added by admin:?\s*/, ''); }

/** Usual time for that punch, from the person's shift and lunch. */
function usualTime(row, type) {
  const sh = String(row.shift || '').split('-');
  const ls = row.lunchStart || (A.settings && A.settings.lunchStart) || '';
  if (type === 'IN') return sh[0] || '';
  if (type === 'OUT') return sh[1] || '';
  if (type === 'LUNCH_OUT') return ls;
  if (!ls) return '';
  const m = (+ls.split(':')[0] * 60 + +ls.split(':')[1] + (row.lunchMin || 0)) % 1440;
  return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2);
}

function openManual(empId, type, p) {
  const r = A.day;
  const row = r && r.rows.find(function (x) { return x.id === empId; });
  if (!row) return;
  if (!type) {
    const ps = r.punches.filter(function (q) { return q.empId === empId; });
    const last = ps.length ? ps[ps.length - 1].action : '';
    type = { '': 'IN', IN: 'OUT', LUNCH_OUT: 'LUNCH_IN', LUNCH_IN: 'OUT', OUT: 'OUT' }[last] || 'IN';
  }
  A.manual = { empId: empId, date: r.date, id: p ? p.id : '', row: row };
  $('#mp-title').textContent = p ? 'Change ' + LABELS[p.action] : 'Add missed punch';
  $('#mp-who').innerHTML = '<b>' + esc(row.name) + '</b> <span class="muted">(' + esc(row.id) + ') · ' +
    fmtDate(new Date(r.date + 'T12:00:00').getTime(), true) + ' · shift ' + esc(shiftText(row.shift)) + '</span>';
  $('#mp-type').value = type;
  $('#mp-type').disabled = !!p;
  $('#mp-time').value = p ? hm24(p.ts) : usualTime(row, type);
  $('#mp-reason').value = p ? manualReason(p) : '';
  $('#mp-delete').classList.toggle('hidden', !p);
  const sh = String(row.shift || '').split('-');
  $('#mp-hint').textContent = 'Saved without a selfie and marked "Added by admin". ' +
    (sh.length === 2 && sh[1] < sh[0] ? 'Night shift: a time after midnight is saved on the next morning.' : '');
  $('#manual-modal').classList.remove('hidden');
  $('#mp-time').focus();
}

$('#mp-type').addEventListener('change', function () {
  if (A.manual && !A.manual.id) $('#mp-time').value = usualTime(A.manual.row, this.value);
});

$('#manual-form').addEventListener('submit', async function (ev) {
  ev.preventDefault();
  const m = A.manual;
  if (!m) return;
  const btn = this.querySelector('button[type=submit]');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>';
  try {
    await call('adminSavePunch', { empId: m.empId, date: m.date, id: m.id, type: $('#mp-type').value,
      time: $('#mp-time').value, reason: $('#mp-reason').value });
    $('#manual-modal').classList.add('hidden');
    toast(LABELS[$('#mp-type').value] + ' saved for ' + m.row.name);
    await loadDay();
  } catch (e) {
    toast(e.message, true);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Save punch';
  }
});

$('#mp-delete').addEventListener('click', async function () {
  const m = A.manual;
  if (!m || !m.id || !confirm('Delete this ' + LABELS[$('#mp-type').value] + ' for ' + m.row.name + '?')) return;
  const btn = this;
  btn.disabled = true;
  try {
    await call('adminDeletePunch', { id: m.id });
    $('#manual-modal').classList.add('hidden');
    toast('Punch deleted');
    await loadDay();
  } catch (e) {
    toast(e.message, true);
  } finally {
    btn.disabled = false;
  }
});

/* ---------------- reports ---------------- */

$('#rep-run').addEventListener('click', runReport);
async function runReport() {
  const from = $('#rep-from').value, to = $('#rep-to').value;
  if (!from || !to) return toast('Pick both dates', true);
  const btn = $('#rep-run');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>';
  try {
    A.report = await call('adminReport', { from: from, to: to, empId: $('#rep-emp').value });
    renderReport();
    $('#rep-csv').disabled = false;
  } catch (e) {
    toast(e.message, true);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Show';
  }
}

function renderReport() {
  const r = A.report;
  const hrs = function (m) { return (m / 60).toFixed(1); };
  let html = '<h3 style="font-size:16px;margin:4px 0 8px">Summary ' + esc(r.from) + ' to ' + esc(r.to) + '</h3>' +
    '<div class="table-wrap"><table><thead><tr><th>ID</th><th>Name</th><th>Shift</th><th class="num">Present</th><th class="num">Half day</th>' +
    '<th class="num">Absent</th><th class="num">Off</th><th class="num">Late days</th><th class="num">Late total</th><th class="num">Early leave</th>' +
    '<th class="num">Worked hrs</th><th class="num">OT hrs</th><th class="num">Lunch over</th><th class="num">No check-out</th>' +
    '<th class="num">Times outside</th><th class="num">Outside</th><th class="num">Added by admin</th></tr></thead><tbody>' +
    r.totals.map(function (t) {
      return '<tr><td>' + esc(t.id) + '</td><td>' + esc(t.name) + '</td><td>' + esc(shiftText(t.shift)) + '</td><td class="num">' + t.present +
        '</td><td class="num">' + t.halfDay + '</td><td class="num">' + t.absent + '</td><td class="num">' + t.off + '</td><td class="num">' + t.late +
        '</td><td class="num">' + dur(t.lateMin) + '</td><td class="num">' + dur(t.earlyMin) + '</td><td class="num">' + hrs(t.workedMin) +
        '</td><td class="num">' + hrs(t.otMin) + '</td><td class="num">' + dur(t.lunchOverMin) + '</td><td class="num">' + t.missedOut +
        '</td><td class="num">' + (t.outsideCount || 0) + '</td><td class="num">' + dur(t.outsideMin || 0) +
        '</td><td class="num">' + (t.manual || 0) + '</td></tr>';
    }).join('') + '</tbody></table></div>';

  html += '<h3 style="font-size:16px;margin:20px 0 8px">Day by day</h3>' +
    '<div class="table-wrap"><table><thead><tr><th>Date</th><th>ID</th><th>Name</th><th>Status</th><th>In</th><th>Lunch out</th><th>Lunch in</th><th>Out</th>' +
    '<th class="num">Worked</th><th class="num">Lunch</th><th class="num">Late</th><th class="num">Early</th><th class="num">OT</th><th class="num">Outside</th></tr></thead><tbody>' +
    r.days.map(function (d) {
      return '<tr><td>' + esc(d.date) + '</td><td>' + esc(d.id) + '</td><td>' + esc(d.name) + '</td><td>' + statusPill(d.status) +
        (d.manual ? ' <span class="pill info" title="Punches added by the admin">Admin ' + d.manual + '</span>' : '') + '</td><td>' +
        (d.in ? fmtTime(d.in) : '') + '</td><td>' + (d.lunchOut ? fmtTime(d.lunchOut) : '') + '</td><td>' + (d.lunchIn ? fmtTime(d.lunchIn) : '') +
        '</td><td>' + (d.out ? fmtTime(d.out) : '') + '</td><td class="num">' + (d.in ? dur(d.workedMin) : '') + '</td><td class="num">' +
        (d.breakMin ? dur(d.breakMin) : '') + '</td><td class="num">' + (d.late ? dur(d.late) : '') + '</td><td class="num">' +
        (d.early ? dur(d.early) : '') + '</td><td class="num">' + (d.ot ? dur(d.ot) : '') + '</td><td class="num">' +
        (d.outsideCount ? d.outsideCount + '× ' + dur(d.outsideMin) : '') + '</td></tr>';
    }).join('') + '</tbody></table></div>';
  $('#rep-out').innerHTML = html;
}

$('#rep-csv').addEventListener('click', function () {
  const r = A.report;
  if (!r) return;
  const q = function (v) { v = String(v === null || v === undefined ? '' : v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
  const t = function (ms) { return ms ? fmtTime(ms) : ''; };
  const lines = [['Date', 'Employee ID', 'Name', 'Department', 'Shift', 'Status', 'In', 'Lunch Out', 'Lunch In', 'Out',
    'Worked (min)', 'Lunch (min)', 'Late (min)', 'Early leave (min)', 'Overtime (min)', 'Lunch over (min)',
    'Times outside', 'Outside (min)', 'Outside in lunch hour (min)', 'Punches added by admin'].join(',')];
  r.days.forEach(function (d) {
    lines.push([d.date, d.id, d.name, d.dept, d.shift, d.status, t(d.in), t(d.lunchOut), t(d.lunchIn), t(d.out),
      d.workedMin, d.breakMin, d.late, d.early, d.ot, d.lunchOver, d.outsideCount || 0, d.outsideMin || 0, d.outsideLunchMin || 0, d.manual || 0].map(q).join(','));
  });
  const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'attendance_' + r.from + '_to_' + r.to + '.csv';
  document.body.appendChild(a);
  a.click();
  setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
});

/* ---------------- employees ---------------- */

async function loadEmployees() {
  try {
    const r = await call('adminEmployees');
    A.employees = r.employees;
    A.locations = r.locations;
    renderEmployees();
    renderLocations();
  } catch (e) { toast(e.message, true); }
}

function renderEmployees() {
  const code = (A.settings && A.settings.companyCode) || companyCode();
  if (code) {
    const link = location.href.replace(/admin\.html.*$/, '') + '?c=' + code;
    $('#emp-share').innerHTML = 'Staff log in with company code <b>' + esc(code) + '</b> · staff app link: <a href="' + esc(link) + '" target="_blank" rel="noopener">' + esc(link) + '</a>';
  }
  $('#emp-share').classList.toggle('hidden', !code);
  const q = $('#emp-search').value.trim().toLowerCase();
  const list = (A.employees || []).filter(function (e) { return !q || (e.id + ' ' + e.name + ' ' + e.dept).toLowerCase().indexOf(q) >= 0; });
  $('#emp-rows').innerHTML = list.length ? list.map(function (e) {
    return '<tr><td>' + esc(e.id) + '</td><td>' + esc(e.name) +
      (e.canPunch ? ' <span class="pill info" title="Can take selfies for other employees">Shared phone</span>' : '') +
      '</td><td>' + esc(e.dept) + '</td><td>' + fmtHM(e.shiftStart) + ' - ' + fmtHM(e.shiftEnd) +
      '</td><td class="num">' + e.lunch + 'm</td><td class="num">' + e.grace + 'm</td><td>' + esc(e.weeklyOff) +
      (e.halfDayOff ? ' <span class="pill info">½ ' + esc(e.halfDayOff) + '</span>' : '') + '</td><td class="num">' + esc(e.pin) + '</td><td>' +
      (e.hasFace ? '<span class="pill ok" title="' + esc(e.faceReg) + '">Registered</span>' : '<span class="pill mute">Not yet</span>') +
      '</td><td>' + (e.active ? '<span class="pill ok">Active</span>' : '<span class="pill mute">Inactive</span>') +
      '</td><td><button class="btn sm ghost" data-edit="' + esc(e.id) + '">Edit</button></td></tr>';
  }).join('') : '<tr><td colspan="11" class="empty">No employees yet.</td></tr>';

  const sel = $('#rep-emp');
  const cur = sel.value;
  sel.innerHTML = '<option value="">All employees</option>' + A.employees.map(function (e) {
    return '<option value="' + esc(e.id) + '">' + esc(e.name) + ' (' + esc(e.id) + ')</option>';
  }).join('');
  sel.value = cur;
  $('#loc-names').innerHTML = A.locations.map(function (l) { return '<option value="' + esc(l.name) + '">'; }).join('');
}

$('#emp-search').addEventListener('input', renderEmployees);
$('#emp-add').addEventListener('click', function () { openEmp(null); });
$('#emp-rows').addEventListener('click', function (ev) {
  const b = ev.target.closest('[data-edit]');
  if (b) openEmp(A.employees.find(function (e) { return e.id === b.dataset.edit; }));
});

$('#f-off').innerHTML = DAYS.map(function (d) { return '<label><input type="checkbox" value="' + d + '"> ' + d + '</label>'; }).join('');

function openEmp(e) {
  const isNew = !e;
  e = e || { id: '', pin: String(Math.floor(1000 + Math.random() * 9000)), name: '', phone: '', dept: '', locations: '',
    shiftStart: '09:00', shiftEnd: '18:00', lunch: 45, grace: 10, weeklyOff: 'Sun', active: true, joinDate: ymd(Date.now()) };
  $('#emp-modal-title').textContent = isNew ? 'Add employee' : 'Edit ' + e.name;
  $('#f-id').value = e.id;
  $('#f-id').disabled = !isNew;
  $('#f-pin').value = e.pin;
  $('#f-name').value = e.name;
  $('#f-phone').value = e.phone;
  $('#f-dept').value = e.dept;
  $('#f-locs').value = e.locations;
  $('#f-start').value = e.shiftStart;
  $('#f-end').value = e.shiftEnd;
  $('#f-lunch').value = e.lunch;
  $('#f-grace').value = e.grace;
  $('#f-join').value = e.joinDate || '';
  $('#f-lunchstart').value = e.lunchStart || '';
  const off = String(e.weeklyOff).toLowerCase();
  document.querySelectorAll('#f-off input').forEach(function (c) { c.checked = off.indexOf(c.value.toLowerCase()) >= 0; });
  $('#f-active').checked = e.active;
  $('#f-halfoff').value = e.halfDayOff || '';
  $('#f-others').checked = !!e.canPunch;
  A.editing = isNew ? null : e.id;
  renderFaceRow();
  $('#f-push-row').classList.toggle('hidden', isNew);
  $('#f-push-status').innerHTML = '<span class="muted">Rings on their phone the next time the Android app is open.</span>';
  $('#emp-form').dataset.isNew = isNew ? '1' : '';
  $('#emp-modal').classList.remove('hidden');
  (isNew ? $('#f-id') : $('#f-name')).focus();
}

$('#emp-form').addEventListener('submit', async function (ev) {
  ev.preventDefault();
  const btn = this.querySelector('button[type=submit]');
  btn.disabled = true;
  try {
    const r = await call('adminSaveEmployee', { employee: {
      isNew: !!this.dataset.isNew, id: $('#f-id').value.trim(), pin: $('#f-pin').value.trim(), name: $('#f-name').value.trim(),
      phone: $('#f-phone').value.trim(), dept: $('#f-dept').value.trim(), locations: $('#f-locs').value.trim(),
      shiftStart: $('#f-start').value, shiftEnd: $('#f-end').value, lunch: $('#f-lunch').value, grace: $('#f-grace').value, joinDate: $('#f-join').value, lunchStart: $('#f-lunchstart').value,
      weeklyOff: Array.prototype.filter.call(document.querySelectorAll('#f-off input'), function (c) { return c.checked; })
        .map(function (c) { return c.value; }).join(', '),
      active: $('#f-active').checked, canPunchOthers: $('#f-others').checked, halfDayOff: $('#f-halfoff').value,
    } });
    A.employees = r.employees;
    A.locations = r.locations;
    renderEmployees();
    $('#emp-modal').classList.add('hidden');
    A.day = null;
    toast('Employee saved');
  } catch (e) {
    toast(e.message, true);
  } finally {
    btn.disabled = false;
  }
});

/* ---------------- locations & settings ---------------- */

function renderLocations() {
  $('#loc-rows').innerHTML = A.locations.length ? A.locations.map(function (l) {
    return '<tr><td>' + esc(l.name) + '</td><td><a href="https://www.google.com/maps?q=' + l.lat + ',' + l.lng + '" target="_blank" rel="noopener">' +
      l.lat.toFixed(5) + ', ' + l.lng.toFixed(5) + '</a></td><td class="num">' + l.radius + ' m</td>' +
      '<td style="white-space:nowrap"><button class="btn sm ghost" data-edit-loc="' + esc(l.name) + '">Edit</button> ' +
      '<button class="btn sm ghost" data-del="' + esc(l.name) + '">Delete</button></td></tr>';
  }).join('') : '<tr><td colspan="4" class="muted">No locations yet. Punches are still recorded with GPS and address.</td></tr>';
}

/** Fills the form with a location so it can be changed (or reset to the normal "add" form). */
function editLocation(l) {
  A.locEditing = l ? l.name : null;
  $('#loc-name').value = l ? l.name : '';
  $('#loc-lat').value = l ? l.lat : '';
  $('#loc-lng').value = l ? l.lng : '';
  $('#loc-radius').value = l ? l.radius : 100;
  $('#loc-save').textContent = l ? 'Update location' : 'Save location';
  $('#loc-cancel').classList.toggle('hidden', !l);
  $('#loc-editing').classList.toggle('hidden', !l);
  $('#loc-editing').innerHTML = l ? '<span class="pill warn">Editing</span> <b>' + esc(l.name) +
    '</b> - change the values or tap <b>Use my current location</b>, then <b>Update location</b>.' : '';
  if (l) { $('#loc-form').scrollIntoView({ behavior: 'smooth', block: 'center' }); $('#loc-radius').focus(); }
}
$('#loc-cancel').addEventListener('click', function () { editLocation(null); });

$('#loc-rows').addEventListener('click', async function (ev) {
  const eb = ev.target.closest('[data-edit-loc]');
  if (eb) {
    const l = A.locations.find(function (x) { return x.name === eb.dataset.editLoc; });
    if (l) editLocation(l);
    return;
  }
  const b = ev.target.closest('[data-del]');
  if (!b || !confirm('Delete location "' + b.dataset.del + '"?')) return;
  try {
    const r = await call('adminDeleteLocation', { name: b.dataset.del });
    A.locations = r.locations;
    renderLocations();
  } catch (e) { toast(e.message, true); }
});

$('#loc-here').addEventListener('click', function () {
  const btn = this;
  if (!navigator.geolocation) return toast('Location not supported', true);
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Locating';
  navigator.geolocation.getCurrentPosition(function (pos) {
    $('#loc-lat').value = pos.coords.latitude.toFixed(6);
    $('#loc-lng').value = pos.coords.longitude.toFixed(6);
    toast('Got location (±' + Math.round(pos.coords.accuracy) + ' m)');
    btn.disabled = false;
    btn.textContent = 'Use my current location';
  }, function (err) {
    toast(err.message || 'Could not get location', true);
    btn.disabled = false;
    btn.textContent = 'Use my current location';
  }, { enableHighAccuracy: true, timeout: 20000, maximumAge: 0 });
});

$('#loc-form').addEventListener('submit', async function (ev) {
  ev.preventDefault();
  const btn = $('#loc-save');
  const was = A.locEditing, name = $('#loc-name').value.trim();
  btn.disabled = true;
  try {
    let r = await call('adminSaveLocation', { location: {
      name: name, lat: $('#loc-lat').value, lng: $('#loc-lng').value, radius: $('#loc-radius').value,
    } });
    // renamed while editing: the new name was saved, so remove the old entry
    if (was && was.toLowerCase() !== name.toLowerCase()) r = await call('adminDeleteLocation', { name: was });
    A.locations = r.locations;
    renderLocations();
    renderEmployees();
    editLocation(null);
    toast(was ? 'Location updated' : 'Location saved');
  } catch (e) {
    toast(e.message, true);
  } finally {
    btn.disabled = false;
  }
});

async function loadSettingsForm() {
  if (!A.settings) {
    try { A.settings = (await call('adminDay', { date: ymd(Date.now()) })).settings; } catch (e) { return; }
  }
  $('#set-company').value = A.settings.company;
  $('#set-fence').checked = A.settings.enforceGeofence;
  $('#set-lunch').value = A.settings.lunchMode || 'SELFIE';
  $('#set-lunchstart').value = A.settings.lunchStart || '';
  $('#set-away').value = A.settings.awayMinutes || 15;
  $('#set-halfday').value = A.settings.minHalfDay !== undefined ? A.settings.minHalfDay : 3;
  $('#set-v-inout').value = A.settings.verifyInOut || (A.settings.faceMode && A.settings.faceMode !== 'OFF' ? 'BOTH' : 'SELFIE');
  $('#set-v-lunch').value = A.settings.verifyLunch || $('#set-v-inout').value;
  $('#set-strict').value = (A.settings.faceStrict !== undefined ? A.settings.faceStrict : A.settings.faceMode === 'REQUIRED') ? '1' : '0';
  $('#set-acc').value = A.settings.maxAccuracy;
}

$('#set-form').addEventListener('submit', async function (ev) {
  ev.preventDefault();
  try {
    const r = await call('adminSaveSettings', { settings: {
      company: $('#set-company').value, enforceGeofence: $('#set-fence').checked, lunchMode: $('#set-lunch').value, lunchStart: $('#set-lunchstart').value, awayMinutes: $('#set-away').value, minHalfDay: $('#set-halfday').value, verifyInOut: $('#set-v-inout').value, verifyLunch: $('#set-v-lunch').value, faceStrict: $('#set-strict').value === '1', maxAccuracy: $('#set-acc').value,
      newPassword: $('#set-pw').value,
    } });
    A.settings = r.settings;
    if (r.token) { A.token = r.token; store.set('att_admin_token', r.token); }
    $('#set-pw').value = '';
    $('#h-company').textContent = (r.settings.company || 'Attendance') + ' - Admin';
    toast('Settings saved');
  } catch (e) { toast(e.message, true); }
});

/* ---------------- face registration ---------------- */

function renderFaceRow() {
  const row = $('#f-face-row');
  const e = A.editing && (A.employees || []).find(function (x) { return x.id === A.editing; });
  row.classList.toggle('hidden', !e);
  if (!e) return;
  $('#f-face-status').innerHTML = e.hasFace
    ? '<span class="pill ok">Face registered</span> <span class="muted">' + esc(e.faceReg) + '</span>'
    : '<span class="pill mute">Face not registered</span> <span class="muted">(first selfie will register it)</span>';
  $('#f-face-reset').classList.toggle('hidden', !e.hasFace);
}

const FR = { stream: null, facing: 'user' };
$('#fr-close').innerHTML = ICONS.close;

async function frStart() {
  frStop();
  FR.stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: FR.facing, width: { ideal: 960 } }, audio: false });
  const v = $('#fr-video');
  v.srcObject = FR.stream;
  v.style.transform = FR.facing === 'user' ? 'scaleX(-1)' : '';
  await v.play().catch(function () {});
}
function frStop() {
  if (FR.stream) FR.stream.getTracks().forEach(function (t) { t.stop(); });
  FR.stream = null;
}
function frClose() { frStop(); $('#face-modal').classList.add('hidden'); }

$('#f-face-reg').addEventListener('click', async function () {
  const e = A.employees.find(function (x) { return x.id === A.editing; });
  if (!e) return;
  $('#fr-title').textContent = 'Register face: ' + e.name;
  $('#face-modal').classList.remove('hidden');
  loadFace().catch(function (err) { toast(err.message, true); });
  try { await frStart(); } catch (err) { toast('Camera: ' + err.message, true); frClose(); }
});
$('#fr-close').addEventListener('click', frClose);
$('#fr-flip').addEventListener('click', function () {
  FR.facing = FR.facing === 'user' ? 'environment' : 'user';
  frStart().catch(function (err) { toast(err.message, true); });
});
$('#fr-shoot').addEventListener('click', async function () {
  const v = $('#fr-video');
  if (!v.videoWidth) return toast('Camera is still starting', true);
  const btn = this;
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Reading face';
  try {
    const c = document.createElement('canvas');
    c.width = v.videoWidth;
    c.height = v.videoHeight;
    c.getContext('2d').drawImage(v, 0, 0);
    const faces = await faceDescriptors(c);
    if (faces.length !== 1) throw new Error(faces.length ? 'More than one face found. Only the employee should be in the frame.' : 'No face found. Face the camera in good light.');
    let r;
    try {
      r = await call('adminSaveFace', { id: A.editing, descriptor: faces[0] });
    } catch (e) {
      throw new Error('Face read OK, but saving failed: ' + e.message);
    }
    if (Array.isArray(r.employees)) A.employees = r.employees;
    else await loadEmployees();
    renderEmployees();
    renderFaceRow();
    frClose();
    toast('Face registered');
  } catch (err) {
    toast(err.message, true);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Capture face';
  }
});
$('#f-face-reset').addEventListener('click', async function () {
  if (!confirm('Reset this face? The next selfie will register a new one.')) return;
  try {
    const r = await call('adminResetFace', { id: A.editing });
    if (Array.isArray(r.employees)) A.employees = r.employees;
    else await loadEmployees();
    renderEmployees();
    renderFaceRow();
    toast('Face reset');
  } catch (err) { toast(err.message, true); }
});

/* ---------------- push notifications ---------------- */

$('#f-push-test').addEventListener('click', async function () {
  const btn = this;
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Sending';
  try {
    await call('adminTestPush', { id: A.editing });
    toast('Test queued. It rings on their phone when the Android app is open (within about a minute) or next opened.');
  } catch (e) {
    toast(e.message, true);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Send test notification';
  }
});

/* ---------------- boot ---------------- */

function start() {
  show('v-main');
  const today = ymd(Date.now());
  $('#day-date').value = today;
  $('#rep-from').value = today.slice(0, 8) + '01';
  $('#rep-to').value = today;
  const tab = store.get('att_admin_tab');
  const b = tab && document.querySelector('#tabs [data-tab="' + tab + '"]');
  if (b && tab !== 'day') b.click();
  else loadDay();
  if (A.tab !== 'emps') loadEmployees();
}

setInterval(function () {
  if (A.token && A.tab === 'day' && !document.hidden && $('#photo-modal').classList.contains('hidden')) loadDay();
}, 60000);

if (store.get('att_tz')) setTZ(store.get('att_tz'));
if (A.token) start(); else show('v-login');
