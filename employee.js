// Employee phone app: log in, see today's shift, punch with a stamped selfie.

const S = {
  token: store.get('att_token'),
  data: null,          // last status from the server
  offset: 0,           // server clock minus phone clock (ms)
  type: null,          // punch being made
  stream: null,
  watchId: null,
  fix: null,           // best GPS fix {lat, lng, acc}
  place: null,         // geocode result for the fix
  geoErr: '',
  geocodedAt: null,
  photo: null,
  thumb: null,
  busy: false,
  targets: [],         // employees this punch is for; empty = myself
  facing: 'user',      // 'user' = front camera, 'environment' = back camera
  group: false,        // group-selfie selection mode
  sel: {},             // selected employee ids in group mode
};

const COLORS = { IN: '#15803d', LUNCH_OUT: '#b45309', LUNCH_IN: '#1d4ed8', OUT: '#b91c1c' };
const HINTS = {
  IN: 'Start your shift', LUNCH_OUT: 'Going for lunch / break', LUNCH_IN: 'Back from lunch', OUT: 'Going home',
};

function show(id) {
  ['v-login', 'v-home', 'v-camera', 'v-done'].forEach(function (v) { $('#' + v).classList.toggle('hidden', v !== id); });
}

function now() { return Date.now() + S.offset; }

/* ---------------- login / status ---------------- */

$('#login-form').addEventListener('submit', async function (ev) {
  ev.preventDefault();
  const btn = this.querySelector('button');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>';
  try {
    const t0 = Date.now();
    const r = await api('login', { empId: $('#emp-id').value.trim(), pin: $('#emp-pin').value.trim() });
    S.token = r.token;
    store.set('att_token', r.token);
    applyStatus(r, t0);
    show('v-home');
  } catch (e) {
    toast(e.message, true);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Log in';
  }
});

/** Gets fresh data from the server. Resolves true when it worked. */
async function refresh(quiet, timeoutMs) {
  if (!S.token) return false;
  S.lastRetry = false;
  $('#btn-refresh').classList.add('spin');
  const slow = quiet ? null : setTimeout(function () { toast('Server is slow, still updating...'); }, 9000);
  try {
    const t0 = Date.now();
    const r = await api('status', { token: S.token, native: !!window.NativeApp }, timeoutMs);
    applyStatus(r, t0);
    return true;
  } catch (e) {
    if (e.auth) { logout(e.message); return false; }
    S.lastError = e.message;
    S.lastRetry = !!e.retry;
    if (!quiet) toast(e.message, true);
    return false;
  } finally {
    clearTimeout(slow);
    $('#btn-refresh').classList.remove('spin');
  }
}

function applyStatus(r, t0) {
  if (!r || !r.profile || !r.summary) {
    throw new Error('Incomplete reply from the server. Admin: update Code.gs and deploy a New version.');
  }
  if (t0) S.offset = r.serverTime - Math.round((t0 + Date.now()) / 2);
  S.data = r;
  setTZ(r.tz);
  store.set('att_company', r.company);
  // remembered so the next app start shows this screen instantly while it updates
  store.set('att_last', { data: r, offset: S.offset });
  renderHome();
  if (window.NativeApp) NativeApp.onStatus(r);
  if (anyFaceUse()) loadFace().catch(function () {}); // download the face models early
}

function faceText(f) {
  if (f === 'Registered') return 'registered now';
  if (/^Match/.test(f)) return 'matched ' + f.replace('Match ', '');
  if (/^NO MATCH/.test(f)) return 'did not match, admin will review';
  return f.toLowerCase();
}

function logout(msg) {
  S.token = null;
  S.data = null;
  store.del('att_token');
  store.del('att_last');
  if (window.NativeApp) NativeApp.onLogout();
  show('v-login');
  if (msg) toast(msg, true);
}

/* ---------------- home ---------------- */

function renderHome() {
  const d = S.data;
  if (!d) return;
  const p = d.profile;
  $('#h-company').textContent = d.company || 'Attendance';
  document.title = (d.company ? d.company + ' - ' : '') + 'Attendance';
  $('#h-name').textContent = p.name;
  $('#h-meta').textContent = p.id + (p.dept ? ' · ' + p.dept : '');
  $('#h-status').innerHTML = statusPill(d.summary.status);

  const chips = ['<span class="pill info">Shift ' + esc(fmtHM(p.shiftStart)) + ' - ' + esc(fmtHM(p.shiftEnd)) + '</span>'];
  if (p.lunch) chips.push('<span class="pill mute">Lunch ' + p.lunch + ' min' + (d.lunchMode === 'AUTO' ? ' (auto)' : '') + '</span>');
  if (p.weeklyOff) chips.push('<span class="pill mute">Off: ' + esc(p.weeklyOff) + '</span>');
  $('#shift-row').innerHTML = chips.join('');

  const s = d.summary;
  const notes = [];
  if (s.holiday) notes.push('<div class="notice ok">Today is a holiday: ' + esc(s.holiday) + '</div>');
  else if (s.offDay) notes.push('<div class="notice ok">Today is your weekly off.</div>');
  if (s.late) notes.push('<div class="notice warn">You checked in ' + dur(s.late) + ' late.</div>');
  if (s.status === 'On Lunch' && s.lunchOver) notes.push('<div class="notice warn">Lunch is ' + dur(s.lunchOver) + ' over your ' + p.lunch + ' min.</div>');
  $('#notice').innerHTML = notes.join('');

  const acts = d.next.map(function (t) {
    return '<button class="action ' + t + '" data-type="' + t + '">' + ICONS[t] +
      '<div><b>' + LABELS[t] + '</b><span>' + HINTS[t] + ' · take selfie</span></div></button>';
  });
  if (!d.next.length) {
    acts.push('<div class="card" style="text-align:center"><b style="font-size:17px">Shift complete</b>' +
      '<div class="muted">You worked ' + dur(s.workedMin) + ' today. See you next shift.</div></div>');
  }
  $('#actions').innerHTML = acts.join('');

  $('#st-worked').textContent = dur(s.workedMin);
  $('#st-lunch').textContent = dur(s.breakMin);
  $('#st-late').textContent = dur(s.late);

  $('#timeline').innerHTML = d.punches.length ? d.punches.map(function (x) {
    const where = x.location ? x.location + (x.inside === 'No' ? ' (outside, ' + x.distance + ' m)' : '') : (x.address || '');
    return '<li><span class="dot ' + x.action + '"></span><div style="flex:1;min-width:0"><div><span class="t">' + fmtTime(x.ts) +
      '</span> · ' + esc(x.label) + '</div><div class="muted small">' + esc(where) + '</div></div></li>';
  }).join('') : '<li class="muted small">No punches yet for this shift.</li>';

  renderTeam();
}

/* ---------------- other employees (shared phone / group selfie) ---------------- */

function renderTeam() {
  const d = S.data;
  const card = $('#team-card');
  if (!d.team) { card.classList.add('hidden'); return; }
  card.classList.remove('hidden');
  Object.keys(S.sel).forEach(function (id) {
    const t = d.team.find(function (x) { return x.id === id; });
    if (!t || !t.next.length) delete S.sel[id];
  });
  $('#btn-group').textContent = S.group ? 'Cancel group' : 'Group selfie';
  $('#team-hint').textContent = S.group
    ? 'Tick everyone in the photo, then choose the action below.'
    : "Take the selfie for staff who don't have a phone.";

  const q = $('#team-search').value.trim().toLowerCase();
  const list = d.team.filter(function (t) { return !q || (t.name + ' ' + t.id + ' ' + t.dept).toLowerCase().indexOf(q) >= 0; });
  $('#team-list').innerHTML = list.length ? list.map(function (t) {
    const last = t.last ? t.last.label + ' ' + fmtTime(t.last.ts) : 'No punch yet';
    let right;
    if (S.group) {
      right = t.next.length ? '<input type="checkbox"' + (S.sel[t.id] ? ' checked' : '') + '>' : '<span class="muted small">Done</span>';
      return '<li class="' + (t.next.length ? 'pickable' : '') + (S.sel[t.id] ? ' picked' : '') + '" data-pick="' + esc(t.id) + '">' +
        right + '<div class="who"><b>' + esc(t.name) + '</b><span class="muted small">' + esc(t.id) + ' · ' + esc(last) + '</span></div>' +
        statusPill(t.status) + '</li>';
    }
    right = t.next.length
      ? t.next.map(function (x) { return '<button class="mini ' + x + '" data-id="' + esc(t.id) + '" data-type="' + x + '">' + LABELS[x] + '</button>'; }).join('')
      : '<span class="muted small">Shift complete</span>';
    return '<li><div class="who"><b>' + esc(t.name) + '</b><span class="muted small">' + esc(t.id) + ' · ' + esc(last) + '</span></div>' +
      statusPill(t.status) + '<div class="mini-row">' + right + '</div></li>';
  }).join('') : '<li class="muted small">No employees found.</li>';
  renderGroupBar();
}

function selectedTeam() {
  return (S.data.team || []).filter(function (t) { return S.sel[t.id]; });
}

function renderGroupBar() {
  const bar = $('#group-bar');
  if (!S.group) { bar.classList.add('hidden'); return; }
  bar.classList.remove('hidden');
  const sel = selectedTeam();
  let common = sel.length ? sel[0].next.slice() : [];
  sel.forEach(function (t) { common = common.filter(function (x) { return t.next.indexOf(x) >= 0; }); });
  bar.innerHTML = '<div><b>' + sel.length + ' selected</b>' +
    (!sel.length ? ' <span class="muted small">Tick names above</span>' : '') +
    (sel.length && !common.length ? '<div class="small" style="color:var(--bad-ink)">These people are at different steps. Select people who are all checking in, or all going to lunch, etc.</div>' : '') +
    '</div>' + common.map(function (x) {
      return '<button class="mini ' + x + '" data-group="' + x + '">' + LABELS[x] + ' (' + sel.length + ')</button>';
    }).join('');
}

$('#btn-group').addEventListener('click', function () {
  S.group = !S.group;
  S.sel = {};
  renderTeam();
});
$('#team-search').addEventListener('input', renderTeam);
$('#team-list').addEventListener('click', function (ev) {
  const b = ev.target.closest('.mini');
  if (b) {
    const t = S.data.team.find(function (x) { return x.id === b.dataset.id; });
    if (t) openCamera(b.dataset.type, [t]);
    return;
  }
  const li = ev.target.closest('li.pickable');
  if (li && S.group) {
    const id = li.dataset.pick;
    if (S.sel[id]) delete S.sel[id]; else S.sel[id] = true;
    renderTeam();
  }
});
$('#group-bar').addEventListener('click', function (ev) {
  const b = ev.target.closest('[data-group]');
  if (b) openCamera(b.dataset.group, selectedTeam());
});

$('#actions').addEventListener('click', function (ev) {
  const b = ev.target.closest('.action');
  if (b) openCamera(b.dataset.type);
});
$('#btn-refresh').innerHTML = ICONS.refresh;
$('#btn-logout').innerHTML = ICONS.logout;
$('#cam-close').innerHTML = ICONS.close;
$('#btn-refresh').addEventListener('click', function () { refresh(); toast('Updated'); });
$('#btn-logout').addEventListener('click', function () { if (confirm('Log out from this phone?')) logout(); });

function tick() {
  const t = now();
  $('#clock-time').textContent = fmtTime(t, true);
  $('#clock-date').textContent = fmtDate(t, true);
  $('#cam-clock').textContent = fmtTime(t, true);
}
setInterval(tick, 1000);

/* ---------------- camera + location ---------------- */

/** SELFIE (photo), FACE (face recognition only, no photo) or BOTH, as set by the admin. */
function methodFor(type) {
  const v = S.data && S.data.verify;
  if (v && v[type]) return v[type];
  return S.data && S.data.faceMode && S.data.faceMode !== 'OFF' ? 'BOTH' : 'SELFIE'; // older backend
}

function anyFaceUse() {
  return ['IN', 'OUT', 'LUNCH_OUT', 'LUNCH_IN'].some(function (t) { return methodFor(t) !== 'SELFIE'; });
}

function faceStrict() {
  return S.method === 'FACE' || (S.data.verify ? !!S.data.faceStrict : S.data.faceMode === 'REQUIRED');
}

async function openCamera(type, targets) {
  S.type = type;
  S.targets = targets || [];
  S.method = methodFor(type);
  S.auto = false;
  S.autoPaused = false;
  S.offerSelfie = false;
  S.faceFallback = false;
  S.camOpenedAt = Date.now();
  S.facing = S.targets.length ? 'environment' : 'user';
  S.photo = S.thumb = null;
  $('#cam-title').textContent = LABELS[type] + (S.targets.length === 1 ? ' · ' + S.targets[0].name
    : S.targets.length > 1 ? ' · ' + S.targets.length + ' people' : '');
  if (S.targets.length) { S.place = null; S.geocodedAt = null; } // fence depends on who it is for
  $('#preview').classList.add('hidden');
  $('#video').classList.remove('hidden');
  $('#bar-shoot').classList.remove('hidden');
  $('#bar-confirm').classList.add('hidden');
  setAutoUI(S.method === 'FACE' ? 'Starting face check...' : null);
  show('v-camera');
  startGeo();
  // fresh each time: a first selfie may have just registered someone
  const refs = faceOn() ? loadFaceRefs() : null;
  try {
    await startStream();
  } catch (e) {
    closeCamera();
    toast(e.name === 'NotAllowedError'
      ? 'Camera permission is blocked. Allow camera for this site in your browser settings.'
      : 'Could not open the camera: ' + e.message, true);
    return;
  }
  if (S.method === 'FACE') {
    await refs;
    const expected = S.targets.length ? S.targets : [S.data.profile];
    if (!S.targets.length && S.data.profile.hasFace && !FL.refs[S.data.profile.id]) await loadFaceRefs();
    const missing = expected.filter(function (t) { return !FL.refs[t.id]; });
    if (!missing.length) {
      S.auto = true; // punch as soon as the face is recognised, no photo
      setAutoUI('Look at the camera');
    } else if (expected.length === 1 && !(S.targets.length ? false : S.data.profile.hasFace)) {
      setAutoUI(null); // first time: one selfie registers the face
      toast('Face not registered yet. Take one selfie to register it.');
    } else if (expected.length === 1) {
      setAutoUI(null);
      toast('Could not load your registered face' + (FL.refsError ? ': ' + FL.refsError : '') + '. Take a selfie instead.', true);
    } else {
      closeCamera();
      toast('No registered face for ' + missing.map(function (t) { return t.name; }).join(', ') + '. Ask the admin to register them.', true);
    }
  }
}

/** Face-only mode replaces the shutter with a status line. */
function setAutoUI(text, tappable) {
  const hint = $('#auto-hint');
  $('#btn-shoot').classList.toggle('hidden', text !== null);
  hint.classList.toggle('hidden', text === null);
  hint.classList.toggle('tap', !!tappable);
  if (text !== null) hint.textContent = text;
}

async function startStream() {
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) throw new Error('This browser has no camera support. Open the link in Chrome or Safari.');
  stopStream();
  S.stream = await navigator.mediaDevices.getUserMedia({
    video: { facingMode: S.facing, width: { ideal: 1920 }, height: { ideal: 1440 }, aspectRatio: { ideal: 4 / 3 } }, audio: false,
  });
  const v = $('#video');
  v.classList.toggle('mirror', S.facing === 'user');
  v.srcObject = S.stream;
  await v.play().catch(function () {});
  setupZoom();
  updateGeoUI();
  startFaceLoop();
}

/* ---------------- zoom: real camera zoom where the phone supports it, digital otherwise ---------------- */

const Z = { level: 1, hw: null };

function camTrack() { return S.stream && S.stream.getVideoTracks()[0]; }

function setupZoom() {
  const t = camTrack();
  const caps = t && t.getCapabilities ? t.getCapabilities() : {};
  Z.hw = caps && caps.zoom && caps.zoom.max > caps.zoom.min ? { min: caps.zoom.min, max: Math.min(caps.zoom.max, 10) } : null;
  const set = t && t.getSettings ? t.getSettings() : {};
  Z.level = Z.hw ? (set.zoom || Z.hw.min) : 1;
  applyZoomUI();
  if (Z.hw && Z.level > Z.hw.min) setZoom(Z.hw.min); // open in the widest view the camera has
}

function zoomRange() { return Z.hw ? [Z.hw.min, Z.hw.max] : [1, 4]; }
function digitalZoom() { return Z.hw ? 1 : Z.level; }

async function setZoom(z) {
  const r = zoomRange();
  z = Math.min(r[1], Math.max(r[0], Math.round(z * 20) / 20));
  if (z === Z.level) return;
  Z.level = z;
  if (Z.hw) {
    try { await camTrack().applyConstraints({ advanced: [{ zoom: z }] }); } catch (e) { Z.hw = null; Z.level = Math.max(1, z); }
  }
  applyZoomUI();
}

function applyZoomUI() {
  const dz = digitalZoom();
  $('#video').style.transform = (S.facing === 'user' ? 'scaleX(-1) ' : '') + (dz !== 1 ? 'scale(' + dz + ')' : '');
  const r = zoomRange();
  $('#zoom-val').textContent = parseFloat(Z.level.toFixed(2)) + 'x';
  $('#zoom-out').disabled = Z.level <= r[0];
  $('#zoom-in').disabled = Z.level >= r[1];
}

$('#zoom-in').addEventListener('click', function () { setZoom(Z.level + (Z.level < 2 ? 0.25 : 0.5)); });
$('#zoom-out').addEventListener('click', function () { setZoom(Z.level - (Z.level <= 2 ? 0.25 : 0.5)); });

// pinch with two fingers on the camera picture
(function () {
  let d0 = 0, z0 = 1;
  const stage = $('.cam-stage');
  const dist = function (e) { return Math.hypot(e.touches[0].clientX - e.touches[1].clientX, e.touches[0].clientY - e.touches[1].clientY); };
  stage.addEventListener('touchstart', function (e) { if (e.touches.length === 2) { d0 = dist(e); z0 = Z.level; } }, { passive: true });
  stage.addEventListener('touchmove', function (e) {
    if (e.touches.length === 2 && d0 && S.stream) { e.preventDefault(); setZoom(z0 * dist(e) / d0); }
  }, { passive: false });
  stage.addEventListener('touchend', function () { d0 = 0; });
})();

function stopStream() {
  stopFaceLoop();
  if (S.stream) S.stream.getTracks().forEach(function (t) { t.stop(); });
  S.stream = null;
  $('#video').srcObject = null;
}
// the shutter must wait until the video actually has frames
['loadedmetadata', 'playing', 'resize'].forEach(function (ev) { $('#video').addEventListener(ev, updateGeoUI); });

function closeCamera() {
  stopStream();
  if (S.watchId !== null) navigator.geolocation.clearWatch(S.watchId);
  S.watchId = null;
  show('v-home');
}
$('#cam-close').addEventListener('click', closeCamera);
$('#cam-flip').innerHTML = ICONS.flip;
$('#cam-flip').addEventListener('click', async function () {
  if (!$('#bar-confirm').classList.contains('hidden')) return; // reviewing a photo
  S.facing = S.facing === 'user' ? 'environment' : 'user';
  try { await startStream(); } catch (e) { toast('Could not switch camera: ' + e.message, true); }
});

function startGeo() {
  S.geoErr = '';
  // keep a fresh-enough earlier fix so a retake does not wait again
  if (S.fix && Date.now() - S.fix.at > 120000) { S.fix = null; S.place = null; }
  updateGeoUI();
  if (!navigator.geolocation) { S.geoErr = 'Location is not supported on this browser.'; updateGeoUI(); return; }
  if (S.watchId !== null) navigator.geolocation.clearWatch(S.watchId);
  S.watchId = navigator.geolocation.watchPosition(function (pos) {
    const c = pos.coords;
    const acc = Math.round(c.accuracy);
    const fresh = !S.fix || Date.now() - S.fix.at > 30000;
    if (fresh || acc <= S.fix.acc) {
      S.fix = { lat: c.latitude, lng: c.longitude, acc: acc, at: Date.now() };
      S.geoErr = '';
      geocode();
      updateGeoUI();
    }
  }, function (err) {
    S.geoErr = err.code === 1
      ? 'Location permission is blocked. Allow location for this site and turn on GPS.'
      : 'Waiting for GPS... make sure Location is ON.';
    updateGeoUI();
  }, { enableHighAccuracy: true, maximumAge: 0, timeout: 30000 });
}

let geoBusy = false;
async function geocode() {
  const f = S.fix;
  if (!f || geoBusy) return;
  if (S.place && S.geocodedAt && distM(S.geocodedAt, f) < 40) return;
  geoBusy = true;
  try {
    const r = await api('geocode', { token: S.token, lat: f.lat, lng: f.lng, forEmpId: S.targets.length ? S.targets[0].id : '' });
    S.place = r;
    S.geocodedAt = { lat: f.lat, lng: f.lng };
  } catch (e) {
    if (e.auth) { closeCamera(); return logout(e.message); }
  } finally {
    geoBusy = false;
    updateGeoUI();
  }
}

function distM(a, b) {
  const r = Math.PI / 180, R = 6371000;
  const x = Math.sin((b.lat - a.lat) * r / 2), y = Math.sin((b.lng - a.lng) * r / 2);
  return 2 * R * Math.asin(Math.sqrt(x * x + Math.cos(a.lat * r) * Math.cos(b.lat * r) * y * y));
}

function updateGeoUI() {
  const box = $('#geo-box');
  let html;
  if (S.geoErr && !S.fix) {
    html = '<div class="row">' + ICONS.pin.replace('<svg', '<svg width="18" height="18"') + '<span>' + esc(S.geoErr) + '</span></div>';
  } else if (!S.fix) {
    html = '<div class="row"><span class="spinner" style="width:16px;height:16px;border-width:2px"></span><span>Getting your location...</span></div>';
  } else {
    const p = S.place;
    let fence = '';
    if (p && p.fenced) {
      fence = p.inside
        ? '<span class="pill ok">At ' + esc(p.place) + '</span>'
        : '<span class="pill ' + (p.enforce ? 'bad' : 'warn') + '">' + p.distance + ' m from ' + esc(p.place) + '</span>';
    }
    html = '<div class="row" style="justify-content:space-between"><b>±' + S.fix.acc + ' m GPS</b>' + fence + '</div>' +
      '<div style="margin-top:4px;opacity:.9">' + esc(p && p.address ? p.address : S.fix.lat.toFixed(5) + ', ' + S.fix.lng.toFixed(5)) + '</div>';
    if (p && p.fenced && !p.inside) {
      html += p.enforce
        ? '<div class="geo-warn bad">You are outside the work area. Go inside ' + esc(p.place) + ' to punch.</div>'
        : '<div class="geo-warn warn">Outside the work area: this punch will be marked as outside.</div>';
    }
  }
  box.innerHTML = html;
  $('#btn-shoot').disabled = !(S.fix && S.stream && $('#video').videoWidth) || outsideBlocked();
}

/** The admin blocks punches outside the work locations, and this person is outside. */
function outsideBlocked() {
  const p = S.place;
  return !!(p && p.fenced && !p.inside && p.enforce);
}

$('#btn-shoot').addEventListener('click', capture);
$('#btn-retake').addEventListener('click', async function () {
  S.photo = S.thumb = null;
  $('#preview').classList.add('hidden');
  $('#video').classList.remove('hidden');
  $('#bar-shoot').classList.remove('hidden');
  $('#bar-confirm').classList.add('hidden');
  try { await startStream(); } catch (e) { toast(e.message, true); closeCamera(); }
});

function faceOn() { return !!S.method && S.method !== 'SELFIE'; }

/* ---------------- live face boxes on the camera ---------------- */
// green = recognised and expected, amber = recognised but not selected, red = not registered / not the right person

const FL = { on: false, busy: false, timer: null, refs: {}, thr: 0.5, last: [] };

async function loadFaceRefs() {
  FL.refsAt = Date.now();
  try {
    const r = await api('faceRefs', { token: S.token });
    FL.refs = r.refs || {};
    FL.thr = r.threshold || 0.5;
    FL.refsError = '';
  } catch (e) {
    FL.refsError = e.message;
    if (e.auth) { closeCamera(); logout(e.message); }
  }
}

function startFaceLoop() {
  if (!faceOn() || FL.on) return;
  FL.on = true;
  FL.loop = (FL.loop || 0) + 1;
  FL.fails = 0;
  FL.worked = false;
  const fs = $('#face-status');
  fs.classList.remove('hidden');
  fs.style.background = '';
  fs.textContent = faceReady ? 'Starting face check...' : 'Loading face check... (first time can take a minute)';
  faceTick(FL.loop);
}

function stopFaceLoop() {
  FL.on = false;
  clearTimeout(FL.timer);
  const c = $('#face-overlay');
  c.getContext('2d').clearRect(0, 0, c.width, c.height);
  $('#face-status').classList.add('hidden');
  FL.last = [];
}

async function faceTick(loop) {
  if (!FL.on || loop !== FL.loop) return; // an older camera session: let it end
  const v = $('#video');
  if (v.videoWidth) {
    try {
      await loadFace();
      if (!FL.on || loop !== FL.loop) return;
      // read faces from a small copy of the frame: much faster on phones
      const small = FL.small || (FL.small = document.createElement('canvas'));
      const sc = Math.min(1, 480 / Math.max(v.videoWidth, v.videoHeight));
      small.width = Math.round(v.videoWidth * sc);
      small.height = Math.round(v.videoHeight * sc);
      small.getContext('2d').drawImage(v, 0, 0, small.width, small.height);
      const found = await faceJob(function () { return detectFaces(small, 320); }, 12000);
      FL.fails = 0;
      FL.worked = true;
      FL.broken = false;
      if (FL.on && loop === FL.loop) drawFaceBoxes(found, 1 / sc);
    } catch (e) {
      console.error('live face check', e);
      FL.fails++;
      if (FL.fails >= 2) {
        FL.fails = 0;
        try { await nextFaceBackend(); } catch (e2) { FL.broken = true; }
      }
      if (FL.on && loop === FL.loop) $('#face-status').textContent = FL.broken ? 'Face check not working on this phone' : 'Retrying face check...';
    }
    offerSelfieIfStuck();
  }
  if (FL.on && loop === FL.loop) FL.timer = setTimeout(function () { faceTick(loop); }, 200);
}

/** Face-only mode must never trap anyone: if face check can't run, allow a selfie (flagged for the admin). */
function offerSelfieIfStuck() {
  if (!S.auto || S.offerSelfie || S.busy) return;
  if (FL.broken || (!FL.worked && Date.now() - S.camOpenedAt > 25000)) {
    S.offerSelfie = true;
    setAutoUI('Face check not working. Tap to take a selfie', true);
  }
}

function faceDistance(a, b) {
  let s = 0;
  for (let i = 0; i < 128; i++) { const d = a[i] - b[i]; s += d * d; }
  return Math.sqrt(s);
}

/** Who is this face, and is it someone this punch is for? */
function classifyFace(desc) {
  let bestId = null, best = 99;
  Object.keys(FL.refs).forEach(function (id) {
    const d = faceDistance(desc, FL.refs[id].d);
    if (d < best) { best = d; bestId = id; }
  });
  const me = S.data.profile;
  const expected = S.targets.length ? S.targets.map(function (t) { return t.id; }) : [me.id];
  if (bestId && best <= FL.thr) {
    const name = FL.refs[bestId].name;
    if (expected.indexOf(bestId) >= 0) return { color: '#22c55e', label: name, ok: true, id: bestId };
    if (!S.targets.length) return { color: '#ef4444', label: 'Not you (' + name + ')' };
    return { color: '#f59e0b', label: name + ' (not selected)' };
  }
  if (!S.targets.length && !FL.refs[me.id]) {
    // the server says this face is registered: the face list just hasn't arrived yet, so fetch it again
    if (me.hasFace) {
      if (Date.now() - (FL.refsAt || 0) > 4000) loadFaceRefs();
      return { color: '#f59e0b', label: FL.refsError ? 'Face list failed: ' + FL.refsError : 'Loading your face...' };
    }
    return { color: '#ef4444', label: 'Not registered yet' };
  }
  return { color: '#ef4444', label: 'Not registered' };
}

function drawFaceBoxes(found, scale) {
  const sc = scale || 1;
  const v = $('#video'), c = $('#face-overlay');
  const cw = v.clientWidth, ch = v.clientHeight, dpr = window.devicePixelRatio || 1;
  if (c.width !== Math.round(cw * dpr) || c.height !== Math.round(ch * dpr)) {
    c.width = Math.round(cw * dpr);
    c.height = Math.round(ch * dpr);
  }
  const ctx = c.getContext('2d');
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, cw, ch);
  // the video uses object-fit: cover, and the front camera is shown mirrored
  const k = Math.min(cw / v.videoWidth, ch / v.videoHeight); // whole picture is shown (object-fit: contain)
  const ox = (cw - v.videoWidth * k) / 2, oy = (ch - v.videoHeight * k) / 2;
  const mirror = S.facing === 'user';
  FL.last = found.map(function (f) {
    const b = f.detection.box;
    const who = classifyFace(f.descriptor);
    let x = ox + b.x * sc * k;
    let y = oy + b.y * sc * k, w = b.width * sc * k, h = b.height * sc * k;
    if (mirror) x = cw - x - w;
    const dz = digitalZoom();
    if (dz !== 1) {
      x = cw / 2 + (x - cw / 2) * dz;
      y = ch / 2 + (y - ch / 2) * dz;
      w *= dz;
      h *= dz;
    }
    ctx.lineWidth = 3;
    ctx.strokeStyle = who.color;
    roundRect(ctx, x, y, w, h, 10);
    ctx.stroke();
    ctx.font = '600 14px system-ui, -apple-system, "Segoe UI", Roboto, sans-serif';
    const tw = ctx.measureText(who.label).width + 14;
    const ly = Math.max(4, y > 26 ? y - 26 : y + h + 4);
    const lx = Math.max(4, Math.min(x, cw - tw - 4)); // keep the name on screen
    ctx.fillStyle = who.color;
    roundRect(ctx, lx, ly, tw, 22, 6);
    ctx.fill();
    ctx.fillStyle = '#fff';
    ctx.textBaseline = 'middle';
    ctx.fillText(who.label, lx + 7, ly + 11);
    return who;
  });

  const st = $('#face-status');
  const green = FL.last.filter(function (w) { return w.ok; });
  maybeAutoPunch(found, green);
  if (!found.length) {
    st.textContent = 'No face - look at the camera';
    st.style.background = 'rgba(239,68,68,.85)';
  } else if (S.targets.length) {
    const ids = green.map(function (w) { return w.id; });
    const got = S.targets.filter(function (t) { return ids.indexOf(t.id) >= 0; }).length;
    st.textContent = got + ' of ' + S.targets.length + ' selected recognised';
    st.style.background = got === S.targets.length ? 'rgba(34,197,94,.85)' : 'rgba(239,68,68,.85)';
  } else {
    st.textContent = green.length ? 'Face recognised ✓' : FL.last[0].label;
    st.style.background = green.length ? 'rgba(34,197,94,.85)' : 'rgba(239,68,68,.85)';
  }
}

async function capture() {
  const v = $('#video');
  if (!v.videoWidth || !S.fix) { toast('Camera is still starting, try again', true); return; }
  // with digital zoom, keep only the centre part the person saw on screen
  const dz = digitalZoom();
  const sw = v.videoWidth / dz, sh = v.videoHeight / dz;
  const sx = (v.videoWidth - sw) / 2, sy = (v.videoHeight - sh) / 2;
  const MAX = 1080;
  const k = Math.min(dz > 1 ? 3 : 1, MAX / Math.max(sw, sh));
  const W = Math.round(sw * k), H = Math.round(sh * k);

  // freeze the frame un-mirrored; faces are read from this copy
  const raw = document.createElement('canvas');
  raw.width = W;
  raw.height = H;
  raw.getContext('2d').drawImage(v, sx, sy, sw, sh, 0, 0, W, H);

  S.faces = [];
  S.faceError = !!S.faceFallback;
  if (faceOn() && !S.faceFallback) {
    stopFaceLoop(); // one face job at a time
    const shoot = $('#btn-shoot');
    shoot.disabled = true;
    $('#cam-title').dataset.title = $('#cam-title').dataset.title || $('#cam-title').textContent;
    $('#cam-title').textContent = 'Checking face...';
    try {
      S.faces = await faceDescriptors(raw);
    } catch (e) {
      S.faceError = true;
      if (faceStrict()) toast(e.message, true);
    }
    $('#cam-title').textContent = $('#cam-title').dataset.title;
    delete $('#cam-title').dataset.title;
    updateGeoUI();
    const need = Math.max(1, S.targets.length);
    if (faceStrict() && !S.faceError && S.faces.length < need) {
      toast(S.faces.length
        ? 'Only ' + S.faces.length + ' of ' + need + ' faces found. Everyone must face the camera.'
        : 'No face found. Look straight at the camera in good light.', true);
      startFaceLoop();
      return; // camera stays live for another try
    }

  }

  const c = $('#canvas');
  c.width = W;
  c.height = H;
  const ctx = c.getContext('2d');
  ctx.save();
  if (S.facing === 'user') {
    ctx.translate(W, 0);
    ctx.scale(-1, 1); // same mirror view the person saw while taking a selfie
  }
  ctx.drawImage(raw, 0, 0, W, H);
  ctx.restore();
  S.stampTime = now();
  drawStamp(ctx, W, H);
  S.photo = c.toDataURL('image/jpeg', 0.8);

  const t = document.createElement('canvas');
  const tk = 240 / W;
  t.width = 240;
  t.height = Math.round(H * tk);
  t.getContext('2d').drawImage(c, 0, 0, t.width, t.height);
  S.thumb = t.toDataURL('image/jpeg', 0.7);

  stopStream();
  $('#preview').src = S.photo;
  $('#preview').classList.remove('hidden');
  $('#video').classList.add('hidden');
  $('#bar-shoot').classList.add('hidden');
  $('#bar-confirm').classList.remove('hidden');
}

function wrapText(ctx, text, maxW, maxLines) {
  const words = String(text).split(/\s+/);
  const lines = [];
  let line = '';
  words.forEach(function (w) {
    const test = line ? line + ' ' + w : w;
    if (ctx.measureText(test).width > maxW && line) { lines.push(line); line = w; } else line = test;
  });
  if (line) lines.push(line);
  if (lines.length > maxLines) {
    lines.length = maxLines;
    let l = lines[maxLines - 1];
    while (ctx.measureText(l + '...').width > maxW && l.length) l = l.slice(0, -1);
    lines[maxLines - 1] = l + '...';
  }
  return lines;
}

function drawStamp(ctx, W, H) {
  const d = S.data, p = S.place || {}, f = S.fix;
  const u = Math.max(W, H * 0.75) / 34;           // base unit scales with photo size
  const pad = u * 1.1;
  const font = function (wt, size) { return wt + ' ' + Math.round(size) + 'px system-ui, -apple-system, "Segoe UI", Roboto, sans-serif'; };
  const maxW = W - pad * 2;

  ctx.font = font(400, u * 0.95);
  const addr = wrapText(ctx, p.address || 'Address not available', maxW, 3);
  const lineH = u * 1.35;
  const group = S.targets.length > 1;
  ctx.font = font(500, u * 0.95);
  const names = group ? wrapText(ctx, S.targets.map(function (t) { return t.name + ' (' + t.id + ')'; }).join(', '), maxW, 4) : [];
  const byLine = S.targets.length ? 'Taken by ' + d.profile.name + ' (' + d.profile.id + ')' : '';
  const bandH = pad * 2 + u * 1.6 + lineH * 2.2 + lineH * names.length + lineH * addr.length + lineH + (byLine ? lineH : 0);
  const y0 = H - bandH;

  const g = ctx.createLinearGradient(0, y0 - u * 3, 0, H);
  g.addColorStop(0, 'rgba(0,0,0,0)');
  g.addColorStop(0.25, 'rgba(0,0,0,0.6)');
  g.addColorStop(1, 'rgba(0,0,0,0.82)');
  ctx.fillStyle = g;
  ctx.fillRect(0, y0 - u * 3, W, bandH + u * 3);

  ctx.textBaseline = 'top';
  let y = y0 + pad;

  // action pill + company
  const label = LABELS[S.type].toUpperCase();
  ctx.font = font(800, u * 0.95);
  const pw = ctx.measureText(label).width + u * 1.4;
  ctx.fillStyle = COLORS[S.type];
  roundRect(ctx, pad, y, pw, u * 1.5, u * 0.4);
  ctx.fill();
  ctx.fillStyle = '#fff';
  ctx.fillText(label, pad + u * 0.7, y + u * 0.27);
  ctx.font = font(600, u * 0.85);
  ctx.textAlign = 'right';
  ctx.fillStyle = 'rgba(255,255,255,0.85)';
  ctx.fillText(d.company || '', W - pad, y + u * 0.32, W - pad * 3 - pw);
  ctx.textAlign = 'left';
  y += u * 1.6 + u * 0.4;

  // name, date-time
  ctx.fillStyle = '#fff';
  ctx.font = font(700, u * 1.25);
  const who = S.targets.length === 1 ? S.targets[0] : d.profile;
  ctx.fillText(group ? 'Group of ' + S.targets.length + ' employees' : who.name + '  (' + who.id + ')', pad, y, maxW);
  y += lineH * 1.15;
  if (group) {
    ctx.font = font(500, u * 0.95);
    names.forEach(function (l) { ctx.fillText(l, pad, y, maxW); y += lineH; });
    ctx.fillStyle = '#fff';
  }
  ctx.font = font(700, u * 1.05);
  ctx.fillText(fmtDate(S.stampTime, false) + '   ' + fmtTime(S.stampTime, true), pad, y, maxW);
  y += lineH * 1.05;

  // address
  ctx.font = font(400, u * 0.95);
  ctx.fillStyle = 'rgba(255,255,255,0.95)';
  addr.forEach(function (l) { ctx.fillText(l, pad, y, maxW); y += lineH; });

  // coordinates + place
  ctx.font = font(500, u * 0.82);
  ctx.fillStyle = 'rgba(255,255,255,0.8)';
  let geo = f.lat.toFixed(6) + ', ' + f.lng.toFixed(6) + '  ±' + f.acc + 'm';
  if (p.fenced) geo += '  ·  ' + (p.inside ? p.place : p.distance + 'm from ' + p.place);
  ctx.fillText(geo, pad, y + u * 0.1, maxW);
  if (byLine) {
    y += lineH;
    ctx.font = font(600, u * 0.85);
    ctx.fillStyle = '#ffd27a';
    ctx.fillText(byLine, pad, y + u * 0.1, maxW);
  }
}

function roundRect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

/** Sends the punch. photo is null for face-only punches. */
async function sendPunch(photo, thumb, faces, faceError) {
  const p = S.place || {};
  const t0 = Date.now();
  const body = {
    token: S.token, type: S.type, lat: S.fix.lat, lng: S.fix.lng, acc: S.fix.acc,
    address: p.address || '', deviceTime: new Date().toISOString(), device: navigator.userAgent,
    forEmpIds: S.targets.map(function (t) { return t.id; }),
    faces: faces || [], faceError: !!faceError,
    photoPending: !!photo, // the photo follows in the background, so the punch is confirmed fast
  };
  let r;
  try {
    r = await api('punch', body, 60000);
  } catch (e) {
    if (!photo || !/selfie is required/i.test(e.message)) throw e;
    // older backend: send the photo with the punch
    r = await api('punch', Object.assign({}, body, { photoPending: false, photo: photo, thumb: thumb }), 90000);
    photo = null;
  }
  if (photo && r.punchId) queuePhoto(r.punchId, photo, thumb);
  $('#done-photo').textContent = photo && r.punchId ? 'Photo uploading...' : '';
  stopStream();
  if (S.watchId !== null) navigator.geolocation.clearWatch(S.watchId);
  S.watchId = null;
  const forOthers = S.targets.length > 0;
  if (forOthers) { S.group = false; S.sel = {}; }
  applyStatus(r, t0);
  $('#done-img').classList.toggle('hidden', !photo);
  $('#done-check').classList.toggle('hidden', !!photo);
  if (photo) $('#done-img').src = photo;
  if (forOthers) {
    $('#done-title').textContent = LABELS[S.type] + ' at ' + fmtTime(now()) +
      (r.punchedFor.length > 1 ? ' · ' + r.punchedFor.length + ' people' : '');
    $('#done-sub').textContent = 'Saved for ' + r.punchedFor.map(function (t) {
      return t.name + (t.face ? ' (face: ' + faceText(t.face) + ')' : '');
    }).join(', ') + '.';
  } else {
    const last = r.punches[r.punches.length - 1];
    const f = r.punchedFor && r.punchedFor[0] && r.punchedFor[0].face;
    $('#done-title').textContent = LABELS[S.type] + ' at ' + fmtTime(last ? last.ts : now());
    $('#done-sub').textContent = (S.type === 'OUT'
      ? 'Worked ' + dur(r.summary.workedMin) + ' today. Have a good evening!'
      : (S.type === 'IN' && r.summary.late ? 'Marked late by ' + dur(r.summary.late) + '.' : 'Saved successfully.')) +
      (f ? ' Face: ' + faceText(f) + '.' : '');
  }
  S.targets = [];
  S.afterPunch = refresh(true);
  show('v-done');
}

$('#btn-submit').addEventListener('click', async function () {
  if (S.busy || !S.photo) return;
  S.busy = true;
  const btn = this;
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Saving';
  try {
    await sendPunch(S.photo, S.thumb, S.faces, S.faceError);
  } catch (e) {
    if (e.auth) { closeCamera(); return logout(e.message); }
    toast(e.message, true);
  } finally {
    S.busy = false;
    btn.disabled = false;
    btn.textContent = 'Submit';
  }
});

/* ---------------- face recognition only: no button, no photo ---------------- */

function maybeAutoPunch(found, green) {
  if (!S.auto || S.busy || S.autoPaused) return;
  if (!S.fix) { setAutoUI('Waiting for GPS...'); FL.okStreak = 0; return; }
  // give the address lookup a moment, so the server doesn't have to do it while saving
  if (!S.place && Date.now() - S.camOpenedAt < 3000) { setAutoUI('Getting location...'); return; }
  if (outsideBlocked()) { setAutoUI('Outside work area - go inside to punch'); FL.okStreak = 0; return; }
  const ids = green.map(function (w) { return w.id; });
  const ready = S.targets.length
    ? S.targets.every(function (t) { return ids.indexOf(t.id) >= 0; })
    : ids.length > 0;
  FL.okStreak = ready ? (FL.okStreak || 0) + 1 : 0;
  if (!ready) { setAutoUI(found.length ? 'Look straight at the camera' : 'Look at the camera'); return; }
  if (FL.okStreak < 2) { setAutoUI('Hold still...'); return; }
  autoPunch(found);
}

async function autoPunch(found) {
  S.busy = true;
  FL.okStreak = 0;
  setAutoUI('Saving...');
  const faces = found.map(function (f) { return Array.from(f.descriptor, function (v) { return Math.round(v * 1e4) / 1e4; }); });
  try {
    await sendPunch(null, null, faces, false);
  } catch (e) {
    if (e.auth) { closeCamera(); return logout(e.message); }
    toast(e.message, true);
    S.autoPaused = true; // don't retry in a loop; wait for a tap
    setAutoUI('Tap to try again', true);
  } finally {
    S.busy = false;
  }
}

$('#auto-hint').addEventListener('click', function () {
  if (S.offerSelfie) {
    S.auto = false;
    S.faceFallback = true; // the punch is saved with "Face check failed" for the admin to review
    setAutoUI(null);
    updateGeoUI();
    return;
  }
  if (!S.autoPaused) return;
  S.autoPaused = false;
  setAutoUI('Look at the camera');
});

/* ---------------- full-screen wheel while the app gets fresh data ---------------- */

function showLoader(text) {
  $('#loader-text').textContent = text || 'Updating...';
  $('#loader-retry').classList.add('hidden');
  $('#loader').classList.remove('hidden');
  clearTimeout(S.loaderSlow);
  S.loaderSlow = setTimeout(function () { $('#loader-text').textContent = 'Server is slow, still updating...'; }, 8000);
}

function hideLoader() {
  clearTimeout(S.loaderSlow);
  $('#loader').classList.add('hidden');
}

/** Opening the app: wait for fresh data before showing anything. */
async function startupRefresh() {
  showLoader('Updating...');
  let ok = false;
  // short tries with automatic retry: after a deploy Google can be slow or answer 404 for a minute
  for (let attempt = 1; attempt <= 6 && !ok; attempt++) {
    ok = await refresh(true, 20000);
    if (!S.token) { hideLoader(); return; } // logged out
    if (ok || !S.lastRetry) break;
    clearTimeout(S.loaderSlow);
    $('#loader-text').textContent = 'Server is updating, retrying (' + attempt + ' of 5)...';
    if (attempt < 6) await new Promise(function (res) { setTimeout(res, 3000); });
  }
  if (ok) { hideLoader(); processUploads(); return; }
  const last = store.get('att_last');
  if (last && last.data && last.data.profile) {
    // no internet / server down: show the last saved screen, clearly marked
    S.offset = last.offset || 0;
    S.data = last.data;
    setTZ(last.data.tz);
    renderHome();
    hideLoader();
    toast('Could not update (' + (S.lastError || 'no connection') + '). Showing the last saved screen.', true);
    return;
  }
  $('#loader-text').textContent = 'Could not load: ' + (S.lastError || 'no connection');
  $('#loader-retry').classList.remove('hidden');
  clearTimeout(S.loaderSlow);
}
$('#loader-retry').addEventListener('click', startupRefresh);

// after every punch the app refreshes from the server before going back home
$('#btn-done').addEventListener('click', async function () {
  if (S.afterPunch) {
    showLoader('Updating...');
    await S.afterPunch;
    S.afterPunch = null;
    hideLoader();
  }
  show('v-home');
});

/* ---------------- background photo upload ---------------- */
// photos wait on the phone until the server has them, even if the app is closed or the signal drops

let uploadQueue = store.get('att_uploads') || [];
let uploading = false;

function saveQueue() {
  try { localStorage.setItem('att_uploads', JSON.stringify(uploadQueue)); }
  catch (e) { /* phone storage full: keep the photo in memory for this session */ }
  renderUploadNote();
}

function queuePhoto(punchId, photo, thumb) {
  uploadQueue.push({ punchId: punchId, photo: photo, thumb: thumb, tries: 0 });
  saveQueue();
  processUploads();
}

async function processUploads() {
  if (uploading || !uploadQueue.length || !S.token) return;
  uploading = true;
  try {
    while (uploadQueue.length) {
      const job = uploadQueue[0];
      try {
        const r = await api('attachPhoto', { token: S.token, punchId: job.punchId, photo: job.photo, thumb: job.thumb }, 120000);
        if (r.busy) { await new Promise(function (res) { setTimeout(res, 5000); }); continue; }
        uploadQueue.shift();
        saveQueue();
        if ($('#done-photo').textContent) $('#done-photo').textContent = 'Photo saved ✓';
      } catch (e) {
        if (e.auth) break;
        job.tries++;
        if (/not found|not your punch|bad punch/i.test(e.message) || job.tries > 30) { uploadQueue.shift(); saveQueue(); continue; }
        saveQueue();
        break; // try again later
      }
    }
  } finally {
    uploading = false;
  }
}

function renderUploadNote() {
  const el = $('#upload-note');
  if (!el) return;
  el.classList.toggle('hidden', !uploadQueue.length);
  el.textContent = uploadQueue.length === 1 ? '1 photo waiting to upload' : uploadQueue.length + ' photos waiting to upload';
}

window.addEventListener('online', processUploads);
setInterval(function () { if (uploadQueue.length && !document.hidden) processUploads(); }, 30000);

/* ---------------- boot ---------------- */

document.addEventListener('visibilitychange', function () {
  if (!document.hidden && !$('#v-home').classList.contains('hidden')) refresh(true);
});
setInterval(function () { if (!document.hidden && !$('#v-home').classList.contains('hidden')) refresh(true); }, 60000);

(function boot() {
  const company = store.get('att_company');
  if (company) $('#login-company').textContent = company;
  tick();
  if (S.token) {
    show('v-home');
    $('#h-name').textContent = 'Loading...';
    startupRefresh();
    renderUploadNote();
  } else {
    show('v-login');
  }
})();
