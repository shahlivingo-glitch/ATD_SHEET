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

async function refresh(quiet) {
  if (!S.token) return;
  try {
    const t0 = Date.now();
    const r = await api('status', { token: S.token });
    applyStatus(r, t0);
  } catch (e) {
    if (e.auth) return logout(e.message);
    if (!quiet) toast(e.message, true);
  }
}

function applyStatus(r, t0) {
  if (t0) S.offset = r.serverTime - Math.round((t0 + Date.now()) / 2);
  S.data = r;
  setTZ(r.tz);
  store.set('att_company', r.company);
  renderHome();
}

function logout(msg) {
  S.token = null;
  S.data = null;
  store.del('att_token');
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
  if (p.lunch) chips.push('<span class="pill mute">Lunch ' + p.lunch + ' min</span>');
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
}

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

async function openCamera(type) {
  S.type = type;
  S.photo = S.thumb = null;
  $('#cam-title').textContent = LABELS[type];
  $('#preview').classList.add('hidden');
  $('#video').classList.remove('hidden');
  $('#bar-shoot').classList.remove('hidden');
  $('#bar-confirm').classList.add('hidden');
  show('v-camera');
  startGeo();
  try {
    await startStream();
  } catch (e) {
    closeCamera();
    toast(e.name === 'NotAllowedError'
      ? 'Camera permission is blocked. Allow camera for this site in your browser settings.'
      : 'Could not open the camera: ' + e.message, true);
  }
}

async function startStream() {
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) throw new Error('This browser has no camera support. Open the link in Chrome or Safari.');
  S.stream = await navigator.mediaDevices.getUserMedia({
    video: { facingMode: 'user', width: { ideal: 1280 }, height: { ideal: 1280 } }, audio: false,
  });
  const v = $('#video');
  v.srcObject = S.stream;
  await v.play().catch(function () {});
  updateGeoUI();
}

function stopStream() {
  if (S.stream) S.stream.getTracks().forEach(function (t) { t.stop(); });
  S.stream = null;
}

function closeCamera() {
  stopStream();
  if (S.watchId !== null) navigator.geolocation.clearWatch(S.watchId);
  S.watchId = null;
  show('v-home');
}
$('#cam-close').addEventListener('click', closeCamera);

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
    const r = await api('geocode', { token: S.token, lat: f.lat, lng: f.lng });
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
  }
  box.innerHTML = html;
  $('#btn-shoot').disabled = !(S.fix && S.stream);
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

function capture() {
  const v = $('#video');
  if (!v.videoWidth || !S.fix) return;
  const MAX = 1080;
  const k = Math.min(1, MAX / Math.max(v.videoWidth, v.videoHeight));
  const W = Math.round(v.videoWidth * k), H = Math.round(v.videoHeight * k);
  const c = $('#canvas');
  c.width = W;
  c.height = H;
  const ctx = c.getContext('2d');
  ctx.save();
  ctx.translate(W, 0);
  ctx.scale(-1, 1); // same mirror view the person saw while taking it
  ctx.drawImage(v, 0, 0, W, H);
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
  const bandH = pad * 2 + u * 1.6 + lineH * 2.2 + lineH * addr.length + lineH;
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
  ctx.fillText(d.profile.name + '  (' + d.profile.id + ')', pad, y, maxW);
  y += lineH * 1.15;
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

$('#btn-submit').addEventListener('click', async function () {
  if (S.busy || !S.photo) return;
  S.busy = true;
  const btn = this;
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Saving';
  try {
    const p = S.place || {};
    const t0 = Date.now();
    const r = await api('punch', {
      token: S.token, type: S.type, lat: S.fix.lat, lng: S.fix.lng, acc: S.fix.acc,
      address: p.address || '', photo: S.photo, thumb: S.thumb,
      deviceTime: new Date().toISOString(), device: navigator.userAgent,
    });
    if (S.watchId !== null) navigator.geolocation.clearWatch(S.watchId);
    S.watchId = null;
    applyStatus(r, t0);
    const last = r.punches[r.punches.length - 1];
    $('#done-img').src = S.photo;
    $('#done-title').textContent = LABELS[S.type] + ' at ' + fmtTime(last ? last.ts : now());
    $('#done-sub').textContent = S.type === 'OUT'
      ? 'Worked ' + dur(r.summary.workedMin) + ' today. Have a good evening!'
      : (S.type === 'IN' && r.summary.late ? 'Marked late by ' + dur(r.summary.late) + '.' : 'Saved successfully.');
    show('v-done');
  } catch (e) {
    if (e.auth) { closeCamera(); return logout(e.message); }
    toast(e.message, true);
  } finally {
    S.busy = false;
    btn.disabled = false;
    btn.textContent = 'Submit';
  }
});

$('#btn-done').addEventListener('click', function () { show('v-home'); });

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
    refresh();
  } else {
    show('v-login');
  }
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(function () {});
})();
