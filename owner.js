// Owner panel: create a separate admin panel (its own Google Sheet) for each company, send its login, manage trials.

const O = { token: store.get('att_owner_token'), data: null, fromRequest: null, editing: null };

function show(id) { ['v-login', 'v-main'].forEach(function (v) { $('#' + v).classList.toggle('hidden', v !== id); }); }

async function call(action, data) {
  try {
    // owner actions always run on the main sheet, never inside a company
    return await api(action, Object.assign({ token: O.token, company: '' }, data || {}));
  } catch (e) {
    if (e.auth) logout(e.message);
    throw e;
  }
}

function logout(msg) {
  O.token = null;
  store.del('att_owner_token');
  show('v-login');
  if (msg) toast(msg, true);
}

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
$('#btn-refresh').addEventListener('click', function () { load(); });

async function busy(btn, label, fn) {
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>';
  try { return await fn(); } catch (e) { toast(e.message, true); } finally { btn.disabled = false; btn.textContent = label; }
}

$('#login-form').addEventListener('submit', function (ev) {
  ev.preventDefault();
  busy(this.querySelector('button'), 'Log in', async function () {
    const r = await api('ownerLogin', { password: $('#own-pw').value, company: '' });
    O.token = r.token;
    store.set('att_owner_token', r.token);
    $('#own-pw').value = '';
    show('v-main');
    await load();
  });
});

async function load() {
  try {
    O.data = await call('ownerData');
    setTZ(O.data.tz);
    render();
  } catch (e) { toast(e.message, true); }
}

function base() { return location.href.replace(/owner\.html.*$/, '').replace(/[?#].*$/, ''); }

function waLink(phone, text) {
  let d = String(phone || '').replace(/\D/g, '');
  if (d.length === 10) d = '91' + d;
  return 'https://wa.me/' + (d.length >= 11 ? d : '') + '?text=' + encodeURIComponent(text);
}

function niceDate(ymdStr) { return ymdStr ? fmtDate(new Date(ymdStr + 'T12:00:00').getTime()) : ''; }

function render() {
  const d = O.data;
  const fresh = d.requests.filter(function (r) { return r.status === 'New'; });
  const older = d.requests.length - fresh.length;
  $('#requests').innerHTML = '<h3 style="font-size:17px;margin:20px 0 8px">Free trial requests' + (fresh.length ? ' <span class="pill bad">' + fresh.length + ' new</span>' : '') + '</h3>' +
    (fresh.length ? fresh.map(function (r) {
      return '<div class="card" style="display:flex;gap:10px;align-items:center;flex-wrap:wrap"><div style="flex:1;min-width:200px"><b>' + esc(r.company) + '</b>' +
        '<div class="muted small">' + esc([r.name, r.phone, r.city, r.staff ? r.staff + ' staff' : ''].filter(Boolean).join(' · ')) +
        ' · ' + fmtDate(r.time) + ' ' + fmtTime(r.time) + '</div></div>' +
        '<button class="btn sm" data-make="' + r.row + '">Create panel</button>' +
        '<a class="btn sm ghost" target="_blank" rel="noopener" href="' + esc(waLink(r.phone, 'Hello ' + (r.name || '') + ', thank you for asking for a free trial of the Staff Attendance App.')) + '">WhatsApp</a>' +
        '<button class="btn sm ghost" data-ignore="' + r.row + '">Ignore</button></div>';
    }).join('') : '<p class="muted small" style="margin:0">No new requests.' + (older ? ' ' + older + ' handled before.' : '') + ' They come from the "Get free trial" button on the login page.</p>');

  $('#co-rows').innerHTML = d.companies.length ? d.companies.map(function (c) {
    const pill = c.status === 'Active' ? 'ok' : c.status === 'Ended' ? 'warn' : 'mute';
    return '<tr><td><b>' + esc(c.name) + '</b>' + (c.notes ? '<div class="muted small">' + esc(c.notes) + '</div>' : '') + '</td><td><code>' + esc(c.code) + '</code></td>' +
      '<td>' + esc(c.contact) + (c.phone ? '<div class="muted small">' + esc(c.phone) + '</div>' : '') + '</td>' +
      '<td>' + (c.trialEnds ? niceDate(c.trialEnds) : '<span class="muted">No end</span>') + '</td>' +
      '<td><span class="pill ' + pill + '">' + (c.status === 'Ended' ? 'Trial ended' : c.status === 'Off' ? 'Switched off' : 'Active') + '</span></td>' +
      '<td style="white-space:nowrap"><button class="btn sm" data-share="' + esc(c.code) + '">Share</button> ' +
      '<button class="btn sm ghost" data-open="' + esc(c.code) + '">Open</button> ' +
      '<button class="btn sm ghost" data-edit="' + esc(c.code) + '">Edit</button></td></tr>';
  }).join('') : '<tr><td colspan="6" class="empty">No companies yet. Create the first one above.</td></tr>';
}

$('#requests').addEventListener('click', function (ev) {
  const mk = ev.target.closest('[data-make]');
  if (mk) {
    const r = O.data.requests.find(function (x) { return String(x.row) === mk.dataset.make; });
    O.fromRequest = r.row;
    $('#n-name').value = r.company;
    $('#n-contact').value = r.name;
    $('#n-phone').value = r.phone;
    $('#n-from').textContent = 'From the free trial request of ' + (r.name || r.company) + '.';
    window.scrollTo({ top: 0, behavior: 'smooth' });
    $('#n-name').focus();
    return;
  }
  const ig = ev.target.closest('[data-ignore]');
  if (ig && confirm('Ignore this request?')) {
    call('ownerSaveRequest', { row: Number(ig.dataset.ignore), status: 'Ignored' }).then(function (r) { O.data = r; render(); }).catch(function (e) { toast(e.message, true); });
  }
});

$('#new-form').addEventListener('submit', function (ev) {
  ev.preventDefault();
  const form = this;
  busy(form.querySelector('button[type=submit]'), 'Create admin panel', async function () {
    const contact = $('#n-contact').value, phone = $('#n-phone').value;
    const r = await call('ownerCreateCompany', { name: $('#n-name').value, contact: contact, phone: phone,
      trialDays: $('#n-days').value, requestRow: O.fromRequest || undefined });
    form.reset();
    $('#n-days').value = 15;
    $('#n-from').textContent = '';
    O.fromRequest = null;
    give(r.created, contact, phone, false);
    await load();
  });
});

/** The ready-to-send login message for a company. */
function give(c, contact, phone, reset) {
  const b = base();
  const lines = [
    'Hello' + (contact ? ' ' + contact : '') + ',',
    '',
    reset ? 'Your new admin password for the Staff Attendance App:' : 'Your FREE trial of the Staff Attendance App is ready.',
    '',
    'Company: ' + c.name,
    'Company code: ' + c.code,
    '',
    'ADMIN PANEL (for you)',
    'Link: ' + b + 'admin.html?c=' + c.code,
    'Password: ' + c.adminPassword,
  ];
  if (!reset) {
    lines.push('', 'STAFF APP (for employees)', 'Link: ' + b + '?c=' + c.code, 'Sample login: Employee ID EMP001, PIN 1234',
      'Android app: ' + b + 'Attendance-Trial.apk (enter company code ' + c.code + ' at login)', '');
    if (c.trialEnds) lines.push('Free trial till ' + niceDate(c.trialEnds) + '.');
    lines.push('Add your employees and work location in the admin panel.');
  }
  $('#give-title').textContent = (reset ? 'New password for ' : 'Panel ready for ') + c.name;
  $('#give-text').value = lines.join('\n');
  $('#give-wa').href = waLink(phone, lines.join('\n'));
  $('#give-share').classList.toggle('hidden', !navigator.share);
  $('#give-modal').classList.remove('hidden');
}

$('#give-share').addEventListener('click', function () {
  navigator.share({ title: $('#give-title').textContent, text: $('#give-text').value }).catch(function () { /* closed */ });
});

$('#give-copy').addEventListener('click', async function () {
  const t = $('#give-text');
  try { await navigator.clipboard.writeText(t.value); } catch (e) { t.select(); document.execCommand('copy'); }
  toast('Message copied');
});

function companyOf(code) { return O.data.companies.find(function (c) { return c.code === code; }); }

async function openAdmin(code) {
  try {
    const r = await call('ownerOpenAdmin', { code: code });
    setCompany(r.code);
    store.set('att_admin_token', r.token);
    window.open('admin.html?c=' + encodeURIComponent(r.code), '_blank');
  } catch (e) { toast(e.message, true); }
}

$('#co-rows').addEventListener('click', function (ev) {
  const op = ev.target.closest('[data-open]');
  if (op) { openAdmin(op.dataset.open); return; }
  const sb = ev.target.closest('[data-share]');
  if (sb) {
    const c = companyOf(sb.dataset.share);
    busy(sb, 'Share', async function () {
      const r = await call('ownerShareInfo', { code: c.code });
      give(r.share, c.contact, c.phone, false);
    });
    return;
  }
  const ed = ev.target.closest('[data-edit]');
  if (!ed) return;
  const c = companyOf(ed.dataset.edit);
  O.editing = c.code;
  $('#edit-title').textContent = c.name + ' (' + c.code + ')';
  $('#e-contact').value = c.contact;
  $('#e-phone').value = c.phone;
  $('#e-ends').value = c.trialEnds;
  $('#e-notes').value = c.notes;
  $('#e-active').checked = c.active;
  $('#edit-modal').classList.remove('hidden');
});

$('#edit-form').addEventListener('submit', function (ev) {
  ev.preventDefault();
  busy(this.querySelector('button[type=submit]'), 'Save', async function () {
    O.data = await call('ownerSaveCompany', { code: O.editing, contact: $('#e-contact').value, phone: $('#e-phone').value,
      trialEnds: $('#e-ends').value, notes: $('#e-notes').value, active: $('#e-active').checked });
    render();
    $('#edit-modal').classList.add('hidden');
    toast('Saved');
  });
});

$('#e-open').addEventListener('click', function () { openAdmin(O.editing); });

$('#e-reset').addEventListener('click', function () {
  const c = companyOf(O.editing);
  if (!confirm('Make a new admin password for ' + c.name + '? The old one stops working.')) return;
  busy(this, 'New admin password', async function () {
    const r = await call('ownerResetAdmin', { code: c.code });
    $('#edit-modal').classList.add('hidden');
    give({ code: c.code, name: c.name, adminPassword: r.adminPassword }, c.contact, c.phone, true);
  });
});

if (O.token) { show('v-main'); load(); } else show('v-login');
