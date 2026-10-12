// Android app extras: time reminders and the lunch-hour range check.
// Only active inside the installed Android app (Capacitor); the website ignores this file.

(function () {
  const C = window.Capacitor;
  if (!C || !C.isNativePlatform || !C.isNativePlatform()) return;

  function plugin(name) {
    if (C.Plugins && C.Plugins[name]) return C.Plugins[name];
    return C.registerPlugin ? C.registerPlugin(name) : null;
  }
  const LN = plugin('LocalNotifications');
  const BG = plugin('BackgroundGeolocation');
  const ID = { LUNCH_OUT: 101, LUNCH_IN: 102, CHECK_OUT: 103, AWAY: 110 };
  const CHANNEL = 'attendance-alerts';
  const W = { id: null, outSince: null, status: null, starting: false, error: '', notif: '', notes: {} };

  let ready = null;
  function setup() {
    if (ready) return ready;
    ready = (async function () {
      try { await LN.requestPermissions(); } catch (e) { /* the person can allow it later in settings */ }
      try {
        await LN.createChannel({ id: CHANNEL, name: 'Attendance alerts', description: 'Lunch and check-out reminders',
          importance: 5, visibility: 1, vibration: true, lights: true });
      } catch (e) { /* older Android: no channels */ }
    })();
    return ready;
  }

  function at(date, hhmm, addMin) {
    const p = String(hhmm || '').split(':');
    if (p.length < 2) return null;
    const d = new Date(date + 'T00:00:00');
    d.setHours(+p[0], +p[1] + (addMin || 0), 0, 0);
    return d.getTime();
  }

  function lunchWindow(d) {
    const r = d.reminders;
    if (!r || !r.lunchStart || r.autoLunch) return null;
    const start = at(d.workDate, r.lunchStart, 0);
    return start ? { start: start, end: start + (r.lunchMinutes || 30) * 60000 } : null;
  }

  function shiftEndTime(d) {
    const p = d.profile;
    let end = at(d.workDate, (d.reminders && d.reminders.shiftEnd) || p.shiftEnd, 0);
    if (end !== null && at(d.workDate, p.shiftStart, 0) >= end) end += 864e5; // night shift ends the next day
    return end;
  }

  async function scheduleReminders(d) {
    await setup();
    try { await LN.cancel({ notifications: [{ id: ID.LUNCH_OUT }, { id: ID.LUNCH_IN }, { id: ID.CHECK_OUT }] }); } catch (e) { }
    const now = Date.now(), list = [], next = d.next || [], s = d.summary || {};
    const win = lunchWindow(d);
    const add = function (id, when, title, body) {
      if (when && when > now + 5000) list.push({ id: id, title: title, body: body, channelId: CHANNEL,
        schedule: { at: new Date(when), allowWhileIdle: true } });
    };
    if (win && next.indexOf('LUNCH_OUT') >= 0 && !s.lunchOut) {
      add(ID.LUNCH_OUT, win.start + 15 * 60000, 'Lunch Out not punched', 'Lunch time started 15 minutes ago. If you went for lunch, punch Lunch Out now.');
    }
    if (next.length === 1 && next[0] === 'LUNCH_IN' && s.lunchOut) {
      add(ID.LUNCH_IN, s.lunchOut + ((d.reminders && d.reminders.lunchMinutes) || 30) * 60000 + 5 * 60000,
        'Lunch time is over', 'Punch Lunch In when you are back at work.');
    }
    if (next.indexOf('OUT') >= 0) {
      add(ID.CHECK_OUT, shiftEndTime(d) + 15 * 60000, 'Shift is over', 'Remember to punch Check Out before you go home.');
    }
    if (list.length) { try { await LN.schedule({ notifications: list }); } catch (e) { console.error('schedule', e); } }
  }

  function awayKey(d) { return 'att_away_' + d.profile.id + '_' + d.workDate; }

  function distanceM(a, b, c, e) {
    const r = Math.PI / 180, R = 6371000;
    const x = Math.sin((c - a) * r / 2), y = Math.sin((e - b) * r / 2);
    return 2 * R * Math.asin(Math.sqrt(x * x + Math.cos(a * r) * Math.cos(c * r) * y * y));
  }

  /** How far outside the nearest work location (0 or less = inside). */
  function outsideBy(d, lat, lng) {
    let best = null;
    (d.reminders.fences || []).forEach(function (f) {
      const over = distanceM(lat, lng, f.lat, f.lng) - f.radius;
      if (!best || over < best.over) best = { over: over, name: f.name };
    });
    return best;
  }

  const OUT_MIN_MS = 2 * 60000;       // shorter than this outside = GPS wobble, ignored
  const OT_LIMIT_MS = 6 * 3600000;     // forgot to Check Out: stop recording 6 h after shift end

  /** Checked in and not on a punched lunch break. */
  function working(d) { return (d.next || []).indexOf('OUT') >= 0; }

  function watchNeeded(d) {
    const r = d.reminders;
    if (!r || !(r.fences || []).length || !working(d)) return false;
    return Date.now() < shiftEndTime(d) + OT_LIMIT_MS;
  }

  /* ---- trips outside the work location ---- */
  let trip = store.get('att_trip');            // the trip going on now, kept if the app restarts
  let outQueue = store.get('att_out_q') || [];  // reports waiting for the internet

  function saveTrip() {
    if (trip) store.set('att_trip', trip); else store.del('att_trip');
    render();
  }

  function queueTrip(t, end) {
    outQueue = outQueue.filter(function (q) { return q.id !== t.id; });
    outQueue.push({ id: t.id, start: t.start, end: end || null, maxDistance: Math.round(t.maxDist || 0), place: t.place || '' });
    store.set('att_out_q', outQueue);
    flushTrips();
  }

  let flushing = false;
  async function flushTrips() {
    if (flushing || !outQueue.length || !S.token) return;
    flushing = true;
    try {
      while (outQueue.length) {
        const q = outQueue[0];
        const ok = await post(Object.assign({ action: 'reportOutside', token: S.token }, q));
        if (!ok) break; // try again later
        outQueue.shift();
        store.set('att_out_q', outQueue);
      }
    } finally {
      flushing = false;
    }
  }

  /** POST to the server through Android's own HTTP (not slowed down in the background). */
  async function post(body) {
    try {
      const http = plugin('CapacitorHttp');
      if (http && http.request) {
        const r = await http.request({ method: 'POST', url: API_URL, headers: { 'Content-Type': 'text/plain;charset=utf-8' }, data: JSON.stringify(body) });
        const j = typeof r.data === 'string' ? JSON.parse(r.data) : r.data;
        return !!(j && (j.ok || /not your|bad /i.test(j.error || ''))); // drop records the server can never accept
      }
      await api(body.action, body);
      return true;
    } catch (e) {
      return false;
    }
  }

  function closeTrip(end) {
    if (!trip) return;
    const t = trip;
    trip = null;
    saveTrip();
    queueTrip(t, Math.max(end, t.start));
  }

  async function updateWatch(d) {
    const need = watchNeeded(d);
    if (need && W.id === null && !W.starting && BG) {
      W.starting = true;
      try {
        W.id = await BG.addWatcher({
          backgroundTitle: 'Attendance',
          backgroundMessage: 'Recording when you leave the work area during work hours.',
          requestPermissions: true, stale: false, distanceFilter: 25,
        }, onLocation);
        W.error = '';
      } catch (e) {
        console.error('watcher', e);
        W.error = e && e.code === 'NOT_AUTHORIZED'
          ? 'Location permission is off. Allow location for Attendance in phone Settings.'
          : 'Range check could not start: ' + ((e && e.message) || e);
      } finally {
        W.starting = false;
        render();
      }
    } else if (!need && W.id !== null) {
      stopWatch();
    }
  }

  function stopWatch() {
    if (W.id !== null && BG) BG.removeWatcher({ id: W.id }).catch(function () {});
    W.id = null;
    W.outSince = null;
    closeTrip(Date.now());
    render();
  }

  /** A message the admin queued (for example the test notification): ring now. */
  async function showNotice(note) {
    W.notes[note.id] = true;
    await setup();
    try {
      await LN.schedule({ notifications: [{ id: 2000 + Math.floor(Math.random() * 90000), channelId: CHANNEL,
        title: note.title || 'Attendance', body: note.body || '', schedule: { at: new Date(Date.now() + 500), allowWhileIdle: true } }] });
    } catch (e) { console.error('notice', e); }
  }

  function hm(t) { return new Date(t).toLocaleTimeString('en-IN', { hour: 'numeric', minute: '2-digit', hour12: true }); }

  /** A short line on the home screen so the person (and admin) can see the reminders are working. */
  function render() {
    const d = W.status, card = document.getElementById('native-card'), el = document.getElementById('native-status');
    if (!card || !el || !d) return;
    card.classList.remove('hidden');
    const r = d.reminders;
    const parts = [];
    if (!r) {
      parts.push('<b>Reminders off:</b> the server needs the latest update (ask the admin).');
    } else {
      const win = lunchWindow(d);
      parts.push('<b>Reminders on</b>' + (win ? ' · lunch ' + hm(win.start) + '–' + hm(win.end) : r.autoLunch ? ' · automatic lunch' : ' · no lunch time set'));
      let range;
      const nx = d.next || [];
      if (!(r.fences || []).length) range = 'off: no work location set by the admin';
      else if (W.error) range = W.error;
      else if (W.id !== null) range = trip ? '<b style="color:var(--bad-ink)">you are outside since ' + hm(trip.start) + '</b>' : 'on';
      else if (nx.indexOf('IN') >= 0) range = 'starts after Check In';
      else if (nx.indexOf('LUNCH_IN') >= 0) range = 'paused during lunch break';
      else if (!nx.length) range = 'off (checked out)';
      else if (working(d) && Date.now() >= shiftEndTime(d) + OT_LIMIT_MS) range = 'stopped: shift ended long ago, please Check Out';
      else range = 'starting...';
      parts.push('Location while working: ' + range);
      if (store.get(awayKey(d))) parts.push('Lunch-hour away alert sent today');
      if (outQueue.length) parts.push(outQueue.length + ' location record' + (outQueue.length > 1 ? 's' : '') + ' waiting for internet');
    }
    if (W.notif && W.notif !== 'granted') parts.push('<span style="color:var(--bad-ink)"><b>Notifications are blocked.</b> Allow them for Attendance in phone Settings.</span>');
    el.innerHTML = parts.join('<br>');
  }

  document.addEventListener('click', async function (ev) {
    if (!ev.target.closest('#native-test')) return;
    await setup();
    try {
      await LN.schedule({ notifications: [{ id: 199, channelId: CHANNEL, title: 'Attendance test',
        body: 'Notifications are working on this phone.', schedule: { at: new Date(Date.now() + 2000), allowWhileIdle: true } }] });
      toast('Test notification in 2 seconds');
    } catch (e) {
      toast('Could not send a test notification: ' + ((e && e.message) || e), true);
    }
  });

  function onLocation(loc, err) {
    const d = W.status;
    if (err) { W.error = err.code === 'NOT_AUTHORIZED' ? 'Location permission is off. Allow location for Attendance in phone Settings.' : String(err.message || err); render(); return; }
    if (!loc || !d) return;
    if (!watchNeeded(d)) { stopWatch(); return; }
    const out = outsideBy(d, loc.latitude, loc.longitude);
    if (!out) return;
    const margin = Math.max(25, Math.min(loc.accuracy || 0, 100)); // don't count GPS wobble as leaving
    const now = Date.now();
    if (out.over <= margin) {
      W.outSince = null;
      if (trip) closeTrip(now);
      return;
    }
    if (!W.outSince) W.outSince = now;
    if (!trip && now - W.outSince >= OUT_MIN_MS) {
      trip = { id: Date.now().toString(36) + Math.random().toString(36).slice(2, 10), empId: d.profile.id,
        start: W.outSince, maxDist: Math.round(out.over), place: out.name };
      saveTrip();
      queueTrip(trip, null); // the admin sees "outside now" straight away
    } else if (trip && out.over > trip.maxDist) {
      trip.maxDist = Math.round(out.over);
      store.set('att_trip', trip);
    }
    // lunch-hour alert: outside too long in lunch hour without punching Lunch Out
    const win = lunchWindow(d);
    if (win && (d.next || []).indexOf('LUNCH_OUT') >= 0 && !store.get(awayKey(d)) && now >= win.start && now < win.end + 3600000) {
      const awayMs = now - Math.max(W.outSince, win.start);
      if (awayMs >= d.reminders.awayMinutes * 60000) alertAway(d, out, Math.round(awayMs / 60000), loc);
    }
  }

  async function alertAway(d, out, minutes, loc) {
    store.set(awayKey(d), Date.now());
    render();
    await setup();
    try {
      await LN.schedule({ notifications: [{ id: ID.AWAY, channelId: CHANNEL, title: 'You left without punching Lunch Out',
        body: 'You have been outside ' + out.name + ' for ' + minutes + ' minutes during lunch hour. Open the app and punch Lunch Out.',
        schedule: { at: new Date(Date.now() + 1000), allowWhileIdle: true } }] });
    } catch (e) { console.error('away alert', e); }
    post({ action: 'reportAway', token: S.token, minutes: minutes, distance: Math.round(out.over), place: out.name,
      lat: loc.latitude, lng: loc.longitude });
  }

  window.NativeApp = {
    onStatus: function (d) {
      if (!d || !d.profile) return;
      W.status = d;
      if (trip && (!working(d) || trip.empId !== d.profile.id)) closeTrip(Date.now());
      flushTrips();
      render();
      if (LN && LN.checkPermissions) LN.checkPermissions().then(function (p) { W.notif = p.display; render(); }).catch(function () {});
      scheduleReminders(d).catch(function (e) { console.error(e); });
      if (d.notice && !W.notes[d.notice.id]) showNotice(d.notice);
      updateWatch(d).catch(function (e) { console.error(e); });
    },
    onLogout: function () {
      stopWatch();
      if (LN) LN.cancel({ notifications: [{ id: ID.LUNCH_OUT }, { id: ID.LUNCH_IN }, { id: ID.CHECK_OUT }] }).catch(function () {});
    },
  };
  document.documentElement.classList.add('native-app');
})();
