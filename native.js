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
  const W = { id: null, outSince: null, status: null, starting: false };

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
    let end = at(d.workDate, p.shiftEnd, 0);
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

  function watchNeeded(d) {
    const r = d.reminders, win = lunchWindow(d);
    if (!r || !win || !(r.fences || []).length) return false;
    if ((d.next || []).indexOf('LUNCH_OUT') < 0) return false; // not checked in, already at lunch, or checked out
    if (store.get(awayKey(d))) return false;                  // already alerted today
    return Date.now() < win.end;                                 // from check-in until the end of lunch hour
  }

  async function updateWatch(d) {
    const need = watchNeeded(d);
    if (need && W.id === null && !W.starting && BG) {
      W.starting = true;
      try {
        W.id = await BG.addWatcher({
          backgroundTitle: 'Attendance',
          backgroundMessage: 'Checking the work area during lunch hour.',
          requestPermissions: true, stale: false, distanceFilter: 25,
        }, onLocation);
      } catch (e) {
        console.error('watcher', e);
      } finally {
        W.starting = false;
      }
    } else if (!need && W.id !== null) {
      stopWatch();
    }
  }

  function stopWatch() {
    if (W.id !== null && BG) BG.removeWatcher({ id: W.id }).catch(function () {});
    W.id = null;
    W.outSince = null;
  }

  function onLocation(loc, err) {
    const d = W.status;
    if (err || !loc || !d) return;
    if (!watchNeeded(d)) { stopWatch(); return; }
    const out = outsideBy(d, loc.latitude, loc.longitude);
    if (!out) return;
    const margin = Math.max(25, Math.min(loc.accuracy || 0, 100)); // don't count GPS wobble as leaving
    const now = Date.now();
    if (out.over <= margin) { W.outSince = null; return; }
    if (!W.outSince) W.outSince = now;
    const win = lunchWindow(d);
    const awayMs = now - Math.max(W.outSince, win.start); // only time inside lunch hour counts
    if (now >= win.start && awayMs >= d.reminders.awayMinutes * 60000) alertAway(d, out, Math.round(awayMs / 60000), loc);
  }

  async function alertAway(d, out, minutes, loc) {
    store.set(awayKey(d), Date.now());
    stopWatch();
    await setup();
    try {
      await LN.schedule({ notifications: [{ id: ID.AWAY, channelId: CHANNEL, title: 'You left without punching Lunch Out',
        body: 'You have been outside ' + out.name + ' for ' + minutes + ' minutes during lunch hour. Open the app and punch Lunch Out.',
        schedule: { at: new Date(Date.now() + 1000), allowWhileIdle: true } }] });
    } catch (e) { console.error('away alert', e); }
    const body = { action: 'reportAway', token: S.token, minutes: minutes, distance: Math.round(out.over), place: out.name,
      lat: loc.latitude, lng: loc.longitude };
    try {
      // Android slows down web requests from apps in the background; its own HTTP is not slowed
      const http = plugin('CapacitorHttp');
      if (http && http.request) {
        await http.request({ method: 'POST', url: API_URL, headers: { 'Content-Type': 'text/plain;charset=utf-8' }, data: JSON.stringify(body) });
      } else {
        await api('reportAway', body);
      }
    } catch (e) { /* the phone still alerted the person */ }
  }

  window.NativeApp = {
    onStatus: function (d) {
      if (!d || !d.profile) return;
      W.status = d;
      scheduleReminders(d).catch(function (e) { console.error(e); });
      updateWatch(d).catch(function (e) { console.error(e); });
    },
    onLogout: function () {
      stopWatch();
      if (LN) LN.cancel({ notifications: [{ id: ID.LUNCH_OUT }, { id: ID.LUNCH_IN }, { id: ID.CHECK_OUT }] }).catch(function () {});
    },
  };
  document.documentElement.classList.add('native-app');
})();
