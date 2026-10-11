// Shared helpers for the employee app and the admin dashboard.

const API_URL = (window.ATTENDANCE_CONFIG || {}).apiUrl || '';

async function api(action, data, timeoutMs) {
  if (!API_URL || API_URL.indexOf('script.google.com') < 0) {
    throw new Error('App is not configured yet: put the Apps Script URL in config.js');
  }
  let res;
  const ctrl = window.AbortController ? new AbortController() : null;
  const timer = ctrl ? setTimeout(function () { ctrl.abort(); }, timeoutMs || 45000) : null;
  try {
    // text/plain keeps this a "simple" request: Apps Script cannot answer CORS preflights.
    res = await fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(Object.assign({ action: action }, data || {})),
      redirect: 'follow',
      signal: ctrl ? ctrl.signal : undefined,
    });
  } catch (e) {
    if (e && e.name === 'AbortError') {
      const t = new Error('The server is taking too long. Check the internet and try again.');
      t.retry = true;
      throw t;
    }
    throw new Error(navigator.onLine === false
      ? 'No internet connection. Please try again.'
      : 'Cannot reach the attendance server. Admin: check the Apps Script deployment (Who has access = Anyone) and the URL in config.js.');
  }
  if (timer) clearTimeout(timer);
  if (res.status === 404 || res.status >= 500) {
    // Google answers 404 for a short while right after a new version is deployed
    const busy = new Error('The server is updating. Please wait a moment and try again.');
    busy.retry = true;
    throw busy;
  }
  let j;
  let text = '';
  try { text = await res.text(); j = JSON.parse(text); } catch (e) {
    throw new Error('Server error (' + res.status + '). Please try again.');
  }
  if (!j.ok) {
    if (j.error === 'Unknown action') {
      throw new Error('This needs the updated backend. Admin: paste the new Code.gs in Apps Script, then Deploy → Manage deployments → New version.');
    }
    const err = new Error(String(j.error || 'Something went wrong').replace(/^AUTH:\s*/, ''));
    err.auth = /^AUTH:/.test(j.error || '');
    throw err;
  }
  return j;
}

const store = {
  get(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* private mode */ } },
  del(k) { try { localStorage.removeItem(k); } catch (e) { /* private mode */ } },
};

let toastTimer;
function toast(msg, isErr) {
  let el = document.getElementById('toast');
  if (!el) { el = document.createElement('div'); el.id = 'toast'; document.body.appendChild(el); }
  el.textContent = msg;
  el.className = 'show' + (isErr ? ' err' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { el.className = ''; }, isErr ? 5000 : 2800);
}

function $(sel, root) { return (root || document).querySelector(sel); }
function esc(s) {
  return String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

let TZ = undefined;
function setTZ(tz) { try { new Intl.DateTimeFormat('en', { timeZone: tz }); TZ = tz; } catch (e) { TZ = undefined; } }
function fmtTime(ms, sec) {
  if (!ms) return '--:--';
  return new Intl.DateTimeFormat('en-IN', { timeZone: TZ, hour: '2-digit', minute: '2-digit', second: sec ? '2-digit' : undefined, hour12: true }).format(new Date(ms));
}
function fmtDate(ms, long) {
  return new Intl.DateTimeFormat('en-IN', { timeZone: TZ, weekday: long ? 'long' : 'short', day: '2-digit', month: 'short', year: 'numeric' }).format(new Date(ms));
}
function ymd(ms) {
  const p = new Intl.DateTimeFormat('en-CA', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date(ms));
  const g = function (t) { return p.find(function (x) { return x.type === t; }).value; };
  return g('year') + '-' + g('month') + '-' + g('day');
}
function dur(min) {
  min = Math.round(min || 0);
  if (!min) return '0m';
  const h = Math.floor(min / 60), m = min % 60;
  return (h ? h + 'h ' : '') + (m ? m + 'm' : '').trim();
}
function fmtHM(hm) {
  const p = String(hm || '').split(':');
  if (p.length < 2) return hm || '';
  let h = +p[0];
  const ap = h >= 12 ? 'PM' : 'AM';
  h = h % 12 || 12;
  return h + ':' + p[1] + ' ' + ap;
}

const STATUS_CLASS = {
  'Present': 'ok', 'Working': 'ok', 'On Lunch': 'warn', 'Half Day': 'warn', 'No Check-out': 'bad',
  'Absent': 'bad', 'Not In Yet': 'mute', 'Upcoming': 'mute', 'Weekly Off': 'info', 'Holiday': 'info',
};
function statusPill(s) { return '<span class="pill ' + (STATUS_CLASS[s] || 'mute') + '">' + esc(s) + '</span>'; }

const ICONS = {
  IN: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><path d="M10 17l5-5-5-5"/><path d="M15 12H3"/></svg>',
  OUT: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="M16 17l5-5-5-5"/><path d="M21 12H9"/></svg>',
  LUNCH_OUT: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 8h1a4 4 0 1 1 0 8h-1"/><path d="M3 8h14v9a4 4 0 0 1-4 4H7a4 4 0 0 1-4-4Z"/><path d="M6 2v2M10 2v2M14 2v2"/></svg>',
  LUNCH_IN: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg>',
  logout: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="M16 17l5-5-5-5"/><path d="M21 12H9"/></svg>',
  refresh: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-3-6.7L21 8"/><path d="M21 3v5h-5"/></svg>',
  close: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg>',
  flip: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 7h-3l-2-3H9L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2Z"/><path d="M9 13a3 3 0 0 1 5.2-2M15 13a3 3 0 0 1-5.2 2"/><path d="M14 9v2h2M10 17v-2H8"/></svg>',
  pin: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0Z"/><circle cx="12" cy="10" r="3"/></svg>',
};
const LABELS = { IN: 'Check In', LUNCH_OUT: 'Lunch Out', LUNCH_IN: 'Lunch In', OUT: 'Check Out' };

/* ---------- face recognition (runs on the phone; only numbers are sent) ---------- */
const FACE_CDN = 'https://cdn.jsdelivr.net/npm/@vladmandic/face-api@1.7.15';
const TF_WASM = 'https://cdn.jsdelivr.net/npm/@tensorflow/tfjs-backend-wasm@4.22.0/dist/';
const FACE_BACKENDS = ['webgl', 'wasm', 'cpu'];
let faceLoading = null;
let faceReady = false;
let faceQueue = Promise.resolve();

function withTimeout(promise, ms, msg) {
  return Promise.race([Promise.resolve(promise), new Promise(function (resolve, reject) {
    setTimeout(function () { reject(new Error(msg || 'Face check took too long')); }, ms);
  })]);
}

/** Face jobs run one at a time; a stuck job is dropped after its time limit so the next one can start. */
function faceJob(fn, ms) {
  const run = faceQueue.then(function () { return withTimeout(fn(), ms); });
  faceQueue = run.then(function () {}, function () {});
  return run;
}

function loadFace() {
  if (faceLoading) return faceLoading;
  faceLoading = new Promise(function (resolve, reject) {
    if (window.faceapi) return resolve();
    const s = document.createElement('script');
    s.src = FACE_CDN + '/dist/face-api.js';
    s.onload = resolve;
    s.onerror = function () { reject(new Error('Could not load face check. Check the internet.')); };
    document.head.appendChild(s);
  }).then(function () {
    const m = FACE_CDN + '/model';
    return Promise.all([
      faceapi.nets.tinyFaceDetector.loadFromUri(m),
      faceapi.nets.faceLandmark68TinyNet.loadFromUri(m),
      faceapi.nets.faceRecognitionNet.loadFromUri(m),
    ]);
  }).then(function () {
    return pickFaceBackend(store.get('att_face_backend'));
  }).then(function () { faceReady = true; });
  faceLoading.catch(function () { faceLoading = null; });
  return faceLoading;
}

/** Runs every face model once on a blank picture: proves this mode works here and makes the first real check fast. */
async function faceWarmUp() {
  const c = document.createElement('canvas');
  c.width = 160;
  c.height = 160;
  const ctx = c.getContext('2d');
  ctx.fillStyle = '#8a8a8a';
  ctx.fillRect(0, 0, 160, 160);
  await faceapi.detectAllFaces(c, new faceapi.TinyFaceDetectorOptions({ inputSize: 160 }));
  await faceapi.nets.faceLandmark68TinyNet.detectLandmarks(c);
  await faceapi.nets.faceRecognitionNet.computeFaceDescriptor(c);
}

/** Phones differ: try the graphics chip, then WebAssembly, then plain CPU, and remember what works. */
async function pickFaceBackend(first) {
  const order = [first].concat(FACE_BACKENDS).filter(function (b, i, a) { return b && FACE_BACKENDS.indexOf(b) >= 0 && a.indexOf(b) === i; });
  for (let i = 0; i < order.length; i++) {
    const be = order[i];
    try {
      if (be === 'wasm' && faceapi.tf.setWasmPaths) faceapi.tf.setWasmPaths(TF_WASM);
      if (!(await faceapi.tf.setBackend(be))) continue;
      await faceapi.tf.ready();
      await withTimeout(faceWarmUp(), be === 'cpu' ? 90000 : 25000, be + ' too slow');
      store.set('att_face_backend', be);
      return be;
    } catch (e) {
      console.error('face mode ' + be + ' failed', e);
    }
  }
  throw new Error('Face check is not supported on this device.');
}

/** After a failure, move on to the next mode in the list. */
async function nextFaceBackend() {
  const cur = faceapi.tf.getBackend();
  const rest = FACE_BACKENDS.slice(FACE_BACKENDS.indexOf(cur) + 1);
  if (!rest.length) throw new Error('Face check is not supported on this device.');
  await pickFaceBackend(rest[0]);
}

async function detectFaces(img, inputSize) {
  return faceapi
    .detectAllFaces(img, new faceapi.TinyFaceDetectorOptions({ inputSize: inputSize, scoreThreshold: 0.45 }))
    .withFaceLandmarks(true)
    .withFaceDescriptors();
}

/** Returns one 128-number descriptor per face found, biggest face first. */
async function faceDescriptors(img) {
  await loadFace();
  let found;
  try {
    found = await faceJob(function () { return detectFaces(img, 416); }, 30000);
  } catch (e) {
    console.error('face check failed, trying the next mode', e);
    try {
      await nextFaceBackend();
      found = await faceJob(function () { return detectFaces(img, 416); }, 60000);
    } catch (e2) {
      console.error(e2);
      throw new Error('Could not read the face on this device (' + (e2 && e2.message || e2) + ').');
    }
  }
  found = Array.isArray(found) ? found : [];
  found.sort(function (a, b) { return b.detection.box.area - a.detection.box.area; });
  return found.map(function (f) { return Array.from(f.descriptor, function (v) { return Math.round(v * 1e4) / 1e4; }); });
}

/* ---------- Android app download link (only on Android phones using the website) ---------- */
document.addEventListener('DOMContentLoaded', function () {
  const native = window.Capacitor && window.Capacitor.isNativePlatform && window.Capacitor.isNativePlatform();
  if (/Android/i.test(navigator.userAgent) && !native) {
    document.querySelectorAll('.android-dl').forEach(function (a) { a.classList.remove('hidden'); });
  }
});

/* ---------- install as an app (home-screen web app) ---------- */
(function setupInstall() {
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(function () {});
  const standalone = window.matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;
  if (standalone) return;
  const ios = /iphone|ipad|ipod/i.test(navigator.userAgent) ||
    (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  let prompt = null;
  const buttons = function () { return document.querySelectorAll('.install-btn'); };
  const showAll = function (on) { buttons().forEach(function (b) { b.classList.toggle('hidden', !on); }); };

  window.addEventListener('beforeinstallprompt', function (e) {
    e.preventDefault();
    prompt = e;
    showAll(true);
  });
  window.addEventListener('appinstalled', function () { prompt = null; showAll(false); toast('App installed. Open it from your home screen.'); });
  document.addEventListener('click', async function (ev) {
    if (!ev.target.closest('.install-btn')) return;
    if (prompt) {
      prompt.prompt();
      const r = await prompt.userChoice;
      prompt = null;
      if (r.outcome === 'accepted') showAll(false);
    } else if (ios) {
      toast('iPhone: open this page in Safari, tap the Share button, then "Add to Home Screen".');
    } else {
      toast('Open the browser menu (⋮) and tap "Install app" or "Add to Home screen".');
    }
  });
  if (ios) document.addEventListener('DOMContentLoaded', function () { showAll(true); });
})();
