/*
  W&L Athletics Van Reservations
  Front end for GitHub Pages. Data lives in the Van Checkout 2026-27 Google Sheet,
  reached through a Google Apps Script web app (Code.gs). Sign-in uses W&L Microsoft 365.
*/
(function () {
  "use strict";

  const CFG = window.VAN_CONFIG || {};
  const VANS = CFG.vans || [];
  const AUTH = CFG.auth || {};
  const AUTH_ON = !!AUTH.clientId;
  const DEMO = !CFG.apiUrl;
  const DOMAIN = (CFG.emailDomain || "").toLowerCase();

  const DEMO_KEY = "wlu-van-demo-v2";
  const PREFS_KEY = "wlu-van-prefs-v1";
  const PENDING_KEY = "wlu-van-pending-v1";
  const REFRESH_MS = 60000;
  const SORRY = " was just reserved for some of those dates. Please choose another van or different dates.";

  const state = {
    reservations: [],
    view: "list",
    listFrom: todayISO(),
    listTo: todayISO(),
    editing: null,       // reservation being changed, or null for a new booking
    mine: [],
    month: todayISO().slice(0, 8) + "01",
    selectedDay: todayISO(),
    preselectVan: null,
    loadPromise: null
  };

  let msalApp = null;
  let account = null;
  let demoMemory = null;

  /* ================= Helpers ================= */

  function $(sel) { return document.querySelector(sel); }
  function pad(n) { return String(n).padStart(2, "0"); }
  function toISO(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  function fromISO(s) { const p = s.split("-").map(Number); return new Date(p[0], p[1] - 1, p[2]); }
  function todayISO() { return toISO(new Date()); }
  function addDays(s, n) { const d = fromISO(s); d.setDate(d.getDate() + n); return toISO(d); }
  function addMonths(s, n) { const d = fromISO(s); d.setDate(1); d.setMonth(d.getMonth() + n); return toISO(d); }
  function fmt(s, opts) { return fromISO(s).toLocaleDateString("en-US", opts); }
  function fmtLong(s) { return fmt(s, { weekday: "long", month: "long", day: "numeric", year: "numeric" }); }
  function fmtRange(a, b) {
    const full = { month: "short", day: "numeric", year: "numeric" };
    if (a === b) return fmt(a, full);
    const sameYear = a.slice(0, 4) === b.slice(0, 4);
    return fmt(a, sameYear ? { month: "short", day: "numeric" } : full) + " to " + fmt(b, full);
  }
  // "2026-10-14 15:05" -> "Oct 14, 3:05 PM"
  function fmtStamp(s) {
    if (!s) return "";
    const parts = s.split(" ");
    let out = fmt(parts[0], { month: "short", day: "numeric" });
    if (parts[1]) {
      const hm = parts[1].split(":").map(Number);
      out += ", " + ((hm[0] % 12) || 12) + ":" + pad(hm[1]) + (hm[0] < 12 ? " AM" : " PM");
    }
    return out;
  }
  function dayCount(a, b) { return Math.round((fromISO(b) - fromISO(a)) / 86400000) + 1; }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function vanName(id) { const v = VANS.find(function (x) { return x.id === id; }); return v ? v.name : "Van " + id; }

  /* ---- Reservation status ---- */

  function isReturned(r) { return r.status === "Returned" && !!r.returnedAt; }
  function isOverdue(r) { return !isReturned(r) && r.endDate < todayISO(); }

  // Last day a reservation blocks the van. An early return frees the remaining days.
  function blockEnd(r) {
    if (isReturned(r)) {
      const rd = r.returnedAt.slice(0, 10);
      if (rd < r.endDate) return rd < r.startDate ? r.startDate : rd;
    }
    return r.endDate;
  }
  function covers(r, date) { return r.startDate <= date && blockEnd(r) >= date; }
  function reservationOn(vanId, date) {
    return state.reservations.find(function (r) { return r.van === vanId && covers(r, date); });
  }
  function overlapping(vanId, start, end, ignoreId) {
    return state.reservations.filter(function (r) {
      return r.van === vanId && (!ignoreId || r.id !== ignoreId) && r.startDate <= end && blockEnd(r) >= start;
    });
  }
  function isAvailable(vanId, start, end, ignoreId) { return overlapping(vanId, start, end, ignoreId).length === 0; }
  function nextFreeDay(vanId, date) {
    let d = date;
    for (let i = 0; i < 730; i++) { if (!reservationOn(vanId, d)) return d; d = addDays(d, 1); }
    return null;
  }
  // { cls, text } for the colored pill
  function statusOf(r) {
    const t = todayISO();
    if (isReturned(r)) return { cls: "done", text: "Returned" };
    if (r.endDate < t) return { cls: "late", text: "Not returned" };
    if (r.startDate <= t) return { cls: "busy", text: "In use" };
    return { cls: "busy", text: "Reserved" };
  }
  function returnDetails(r) {
    if (!isReturned(r)) return "";
    const bits = ["Returned " + fmtStamp(r.returnedAt)];
    if (r.returnMileage) bits.push(r.returnMileage + " mi");
    if (r.fullTank) bits.push("Full tank: " + r.fullTank);
    let html = '<span class="ret">' + esc(bits.join(" · ")) + "</span>";
    if (r.issues && !/^(none|no|n\/a|na|-)\.?$/i.test(r.issues.trim())) {
      html += '<span class="ret issue">Issue reported: ' + esc(r.issues) + "</span>";
    }
    return html;
  }

  function storageGet(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } }
  function storageSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* storage blocked */ } }
  function storageDel(k) { try { localStorage.removeItem(k); } catch (e) { /* storage blocked */ } }

  let toastTimer = null;
  function toast(msg) {
    const t = $("#toast");
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.hidden = true; }, 7000);
  }
  function setStatus(msg) { $("#status").textContent = msg; }

  /* ================= W&L sign-in (Microsoft 365) ================= */

  function signedInEmail() { return account ? String(account.username || "").toLowerCase() : ""; }
  function redirectUri() { return AUTH.redirectUri || (location.origin + location.pathname); }

  async function initAuth() {
    if (!AUTH_ON) { renderAuth(); return; }
    if (!window.msal) { setStatus("The W&L sign-in service did not load. Refresh the page to try again."); return; }

    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: AUTH.clientId,
        authority: "https://login.microsoftonline.com/" + (AUTH.tenantId || "organizations"),
        redirectUri: redirectUri()
      },
      cache: { cacheLocation: "localStorage" }
    });

    try {
      const result = await msalApp.handleRedirectPromise();
      if (result && result.account) account = result.account;
    } catch (e) {
      console.error(e);
      toast("Sign-in did not finish. Please try again.");
    }
    if (!account) {
      const all = msalApp.getAllAccounts();
      if (all.length) account = all[0];
    }
    if (account) msalApp.setActiveAccount(account);
    renderAuth();

    // Resume a reservation the person started before signing in
    const pending = storageGet(PENDING_KEY);
    if (account && pending) {
      storageDel(PENDING_KEY);
      await load(true);
      openBooking(pending.van, pending.date);
    }
  }

  function signIn(pending) {
    if (pending) storageSet(PENDING_KEY, pending);
    msalApp.loginRedirect({ scopes: ["User.Read"], prompt: "select_account" });
  }

  function signOut() {
    msalApp.logoutRedirect({ account: account, postLogoutRedirectUri: redirectUri() });
  }

  // Token the Google Sheet uses to confirm who is booking (it asks Microsoft directly)
  async function getToken() {
    if (!AUTH_ON || !account) return null;
    const req = { scopes: ["User.Read"], account: account };
    try {
      return (await msalApp.acquireTokenSilent(req)).accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) { msalApp.acquireTokenRedirect(req); }
      throw e;
    }
  }

  function renderAuth() {
    const area = $("#authArea");
    if (!AUTH_ON) { area.innerHTML = ""; return; }
    if (account) {
      area.innerHTML = '<span class="who">Signed in as ' + esc(account.name || account.username) + "</span>" +
        '<button type="button" class="linklike" id="signOutBtn">Sign out</button>';
    } else {
      area.innerHTML = '<button type="button" class="linklike" id="signInBtn">Sign in with W&amp;L</button>';
    }
  }

  /* ================= Data ================= */

  function normalize(r) {
    function s(v) { return v == null ? "" : String(v).trim(); }
    return {
      id: s(r.id),
      van: s(r.van),
      startDate: s(r.startDate).slice(0, 10),
      endDate: s(r.endDate).slice(0, 10),
      team: s(r.team),
      driver: s(r.driver),
      destination: s(r.destination),
      status: s(r.status) || "Reserved",
      returnedAt: s(r.returnedAt),
      returnMileage: s(r.returnMileage),
      fullTank: s(r.fullTank),
      issues: s(r.issues)
    };
  }

  function demoSeed() {
    const t = todayISO();
    return [
      { id: "R0001", van: "206", startDate: addDays(t, -6), endDate: addDays(t, -5), team: "Volleyball", driver: "Sample Coach", destination: "Away match",
        status: "Returned", returnedAt: addDays(t, -5) + " 21:40", returnMileage: "48210", fullTank: "Yes", issues: "None" },
      { id: "R0002", van: "97", startDate: addDays(t, -3), endDate: addDays(t, -2), team: "Men's Soccer", driver: "Sample Coach", destination: "Away game", status: "Reserved" },
      { id: "R0003", van: "208", startDate: t, endDate: t, team: "Men's Tennis", driver: "Sample Coach", destination: "Recruiting", status: "Reserved" },
      { id: "R0004", van: "140", startDate: addDays(t, 1), endDate: addDays(t, 3), team: "Field Hockey", driver: "Sample Coach", destination: "Away game", status: "Reserved" },
      { id: "R0005", van: "203", startDate: addDays(t, 5), endDate: addDays(t, 6), team: "Football", driver: "Sample Coach", destination: "Scouting", status: "Reserved" }
    ];
  }
  function demoLoad() {
    let rows = storageGet(DEMO_KEY);
    if (!Array.isArray(rows)) rows = demoMemory || demoSeed();
    demoMemory = rows;
    storageSet(DEMO_KEY, rows);
    return rows.slice();
  }
  function demoSave(rows) { demoMemory = rows; storageSet(DEMO_KEY, rows); }

  async function fetchReservations() {
    if (DEMO) return demoLoad().map(normalize).filter(function (r) { return r.status !== "Cancelled"; });
    const res = await fetch(CFG.apiUrl + "?action=list&t=" + Date.now());
    if (!res.ok) throw new Error("List request failed (" + res.status + ")");
    const data = await res.json();
    if (data.ok === false) throw new Error(data.message || "List request failed");
    return (data.reservations || []).map(normalize).filter(function (r) { return r.van && r.startDate && r.endDate; });
  }

  // Demo mode: same rules as the Google Sheet, saved only in this browser
  function demoApi(action, p) {
    const all = demoLoad();
    const t = todayISO();
    const clash = function (ignoreId) {
      return all.map(normalize).some(function (r) {
        return r.status !== "Cancelled" && r.van === p.van && r.id !== ignoreId && r.startDate <= p.endDate && blockEnd(r) >= p.startDate;
      });
    };
    if (action === "book") {
      if (clash(null)) return { ok: false, conflict: true };
      const id = "R" + String(all.length + 1).padStart(4, "0");
      all.push(Object.assign({}, p, { id: id, status: "Reserved" }));
      demoSave(all);
      return { ok: true, id: id };
    }
    if (action === "mine") {
      return { ok: true, reservations: all.filter(function (r) {
        return (r.status || "Reserved") === "Reserved" && r.endDate >= t && String(r.email || "").toLowerCase() === p.email;
      }) };
    }
    const r = all.find(function (x) { return x.id === p.id; });
    if (!r) return { ok: false, message: "That reservation was not found." };
    if (action === "cancel") { r.status = "Cancelled"; demoSave(all); return { ok: true, id: r.id }; }
    if (action === "update") {
      if (clash(r.id)) return { ok: false, conflict: true };
      ["van", "startDate", "endDate", "team", "driver", "destination", "notes"].forEach(function (k) { r[k] = p[k]; });
      demoSave(all);
      return { ok: true, id: r.id };
    }
    return { ok: false, message: "Unknown request." };
  }

  // Sends book, update, cancel, or mine to the Google Sheet
  async function api(action, payload) {
    if (DEMO) return demoApi(action, payload);
    const body = Object.assign({ action: action, token: await getToken() }, payload);
    // text/plain keeps this a "simple" request, which Google Apps Script accepts from GitHub Pages
    const res = await fetch(CFG.apiUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(body)
    });
    if (!res.ok) throw new Error("Booking request failed (" + res.status + ")");
    return await res.json();
  }

  function load(quiet) {
    if (state.loadPromise) return state.loadPromise;
    if (!quiet) setStatus("Loading reservations...");
    state.loadPromise = fetchReservations()
      .then(function (rows) {
        rows.sort(function (a, b) { return a.startDate.localeCompare(b.startDate) || Number(a.van) - Number(b.van); });
        state.reservations = rows;
        setStatus("Last updated " + new Date().toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" }));
        render();
      })
      .catch(function (e) {
        console.error(e);
        setStatus("Could not load reservations. Check your connection and refresh the page.");
      })
      .finally(function () { state.loadPromise = null; });
    return state.loadPromise;
  }

  /* ================= Rendering ================= */

  function render() {
    $("#listView").hidden = state.view !== "list";
    $("#calendarView").hidden = state.view !== "calendar";
    document.querySelectorAll(".toggle").forEach(function (b) {
      b.setAttribute("aria-pressed", String(b.dataset.view === state.view));
    });
    if (state.view === "list") renderList(); else { renderCalendar(); renderDay(); }
  }

  function tripItem(r, highlight) {
    const st = statusOf(r);
    return '<li class="' + (highlight ? "now" : "") + '">' +
      '<div class="trip-top"><span>' + esc(fmtRange(r.startDate, r.endDate)) + "</span>" +
      '<span class="res-team">' + esc(r.team) + "</span></div>" +
      '<div class="trip-sub"><span class="pill ' + st.cls + '">' + esc(st.text) + "</span>" +
      (r.id ? '<span class="rid">' + esc(r.id) + "</span>" : "") + returnDetails(r) + "</div></li>";
  }

  function renderList() {
    const from = state.listFrom;
    const to = state.listTo;
    const today = todayISO();
    $("#listFrom").value = from;
    $("#listTo").value = to;
    $("#listTo").min = from;
    const rangeText = from === to
      ? "on " + fmt(from, { weekday: "short", month: "short", day: "numeric" })
      : "for all of " + fmtRange(from, to);

    $("#vanList").innerHTML = VANS.map(function (v) {
      const mine = state.reservations.filter(function (r) { return r.van === v.id; });
      const hits = overlapping(v.id, from, to);
      const cur = hits[0];
      const upcoming = mine.filter(function (r) { return !isReturned(r) && r.endDate >= today; });
      const recent = mine.filter(function (r) { return isReturned(r) || r.endDate < today; })
        .sort(function (a, b) { return b.startDate.localeCompare(a.startDate); }).slice(0, 3);
      const overdue = mine.filter(isOverdue);

      let status;
      if (hits.length) {
        const free = nextFreeDay(v.id, from);
        status =
          '<p class="van-status"><span class="pill busy">' + (from === to ? "Reserved" : "Not available") + "</span>" +
          "<span>" + hits.map(function (r) {
            return "<strong>" + esc(r.team) + "</strong>, " + esc(fmtRange(r.startDate, r.endDate)) + (isReturned(r) ? " (returned)" : "");
          }).join("<br>") + "</span></p>" +
          (free && free > from ? '<p class="hint">Next open day: ' + esc(fmt(free, { weekday: "short", month: "short", day: "numeric" })) + "</p>" : "");
      } else {
        status =
          '<p class="van-status"><span class="pill ok">Available</span>' +
          "<span>" + esc(rangeText) + "</span></p>";
      }

      const warn = overdue.length
        ? '<p class="warn">Return form not received for ' + esc(overdue[0].team) + ", " + esc(fmtRange(overdue[0].startDate, overdue[0].endDate)) + ".</p>"
        : "";

      const upList = upcoming.length
        ? '<ul class="res-list">' + upcoming.map(function (r) { return tripItem(r, r.startDate <= to && blockEnd(r) >= from); }).join("") + "</ul>"
        : '<p class="hint">No upcoming reservations.</p>';
      const recentList = recent.length
        ? '<ul class="res-list">' + recent.map(function (r) { return tripItem(r, false); }).join("") + "</ul>"
        : '<p class="hint">No recent trips.</p>';

      const btn = (!cur && from >= today)
        ? '<button type="button" class="btn small" data-reserve="' + esc(v.id) + '" data-date="' + from + '" data-end="' + to + '">Reserve ' + esc(v.name) + "</button>"
        : "";

      return '<article class="van-card ' + (cur ? "is-busy" : "is-ok") + '">' +
        "<h3>" + esc(v.name) + "</h3>" + status + warn +
        "<h4>Upcoming reservations</h4>" + upList +
        "<h4>Recent trips</h4>" + recentList + btn + "</article>";
    }).join("");
  }

  function renderCalendar() {
    const m = state.month;
    const today = todayISO();
    const first = fromISO(m);
    const daysInMonth = new Date(first.getFullYear(), first.getMonth() + 1, 0).getDate();
    const cellCount = Math.ceil((first.getDay() + daysInMonth) / 7) * 7;
    const start = addDays(m, -first.getDay());

    $("#calHeading").textContent = fmt(m, { month: "long", year: "numeric" });

    let html = "";
    for (let i = 0; i < cellCount; i++) {
      const d = addDays(start, i);
      const res = state.reservations.filter(function (r) { return covers(r, d); });
      const classes = ["day"];
      if (d.slice(0, 7) !== m.slice(0, 7)) classes.push("out");
      if (d < today) classes.push("past");
      if (d === today) classes.push("today");
      if (d === state.selectedDay) classes.push("selected");

      const label = fmtLong(d) + ": " + (res.length
        ? res.length + " of " + VANS.length + " vans reserved, " + res.map(function (r) {
            return vanName(r.van) + " by " + r.team + " (" + statusOf(r).text + ")";
          }).join("; ")
        : "all vans available");

      const chips = res.length
        ? res.map(function (r) {
            const st = statusOf(r);
            return '<span class="chip ' + st.cls + '" title="' + esc(vanName(r.van) + ", " + r.team + ", " + fmtRange(r.startDate, r.endDate) + ", " + st.text) + '">' +
              (st.cls === "done" ? "&#10003; " : "") + esc(r.van) + " " + esc(r.team) + "</span>";
          }).join("")
        : (d >= today ? '<span class="open">All open</span>' : "");

      const count = res.length
        ? '<span class="count busy">' + res.length + " out</span>"
        : (d >= today ? '<span class="count ok">Open</span>' : "");

      html += '<button type="button" class="' + classes.join(" ") + '" data-day="' + d + '" aria-label="' + esc(label) +
        '" aria-pressed="' + (d === state.selectedDay) + '">' +
        '<span class="num" aria-hidden="true">' + fromISO(d).getDate() + "</span>" +
        '<span class="chips" aria-hidden="true">' + chips + "</span>" +
        '<span aria-hidden="true">' + count + "</span></button>";
    }
    $("#calGrid").innerHTML = html;
  }

  function renderDay() {
    const d = state.selectedDay;
    const today = todayISO();
    $("#dayHeading").textContent = fmtLong(d);
    $("#dayDetail").innerHTML = '<ul class="day-list">' + VANS.map(function (v) {
      const r = reservationOn(v.id, d);
      if (r) {
        const st = statusOf(r);
        return '<li><span class="pill ' + st.cls + '">' + esc(st.text) + '</span><span class="dl-van">' + esc(v.name) + "</span>" +
          '<span class="dl-info"><strong>' + esc(r.team) + "</strong>, " + esc(fmtRange(r.startDate, r.endDate)) +
          (r.destination ? " (" + esc(r.destination) + ")" : "") +
          (r.id ? ' <span class="rid">' + esc(r.id) + "</span>" : "") + returnDetails(r) + "</span></li>";
      }
      return '<li><span class="pill ok">Available</span><span class="dl-van">' + esc(v.name) + "</span>" +
        (d >= today ? '<button type="button" class="btn small" data-reserve="' + esc(v.id) + '" data-date="' + d + '">Reserve</button>' : "") +
        "</li>";
    }).join("") + "</ul>";
  }

  /* ================= Booking ================= */

  function fillTeams() {
    const sel = $("#team");
    sel.innerHTML = '<option value="">Choose a team</option>' +
      (CFG.teams || []).map(function (t) { return '<option value="' + esc(t) + '">' + esc(t) + "</option>"; }).join("") +
      '<option value="__other">Other</option>';
  }

  function selectedVan() {
    const c = document.querySelector('input[name="van"]:checked');
    return c ? c.value : null;
  }

  // A trip already under way keeps its van
  function lockVan(vanId) {
    document.querySelectorAll('input[name="van"]').forEach(function (i) {
      if (i.value !== vanId) i.closest("label").hidden = true;
    });
  }

  function updateChoices() {
    const s = $("#startDate").value;
    let e = $("#endDate").value;
    $("#endDate").min = s || todayISO();
    if (s && (!e || e < s)) { e = s; $("#endDate").value = s; }

    const prev = selectedVan() || state.preselectVan;
    if (!s || !e) {
      $("#vanChoices").innerHTML = "";
      $("#dayCount").textContent = "";
      $("#unavailNote").textContent = "Choose your dates to see which vans are open.";
      return;
    }
    const n = dayCount(s, e);
    $("#dayCount").textContent = n + (n === 1 ? " day: " : " days: ") + fmtRange(s, e);

    const ignore = state.editing ? state.editing.id : null;
    const avail = VANS.filter(function (v) { return isAvailable(v.id, s, e, ignore); });
    const taken = VANS.length - avail.length;

    $("#vanChoices").innerHTML = avail.length
      ? avail.map(function (v) {
          return '<label class="choice"><input type="radio" name="van" value="' + esc(v.id) + '"' +
            (v.id === prev ? " checked" : "") + "><span>" + esc(v.name) + "</span></label>";
        }).join("")
      : '<p class="error">No vans are open for these dates. Try different dates.</p>';

    if (state.editing && state.editing.startDate < todayISO()) lockVan(state.editing.van);
    $("#unavailNote").textContent = taken
      ? taken + (taken === 1 ? " van is" : " vans are") + " already reserved for some or all of these dates and not shown."
      : "All vans are open for these dates.";
  }

  function openBooking(vanId, date, endDate, editRes) {
    if (AUTH_ON && !account) { signIn({ van: vanId || null, date: date || null }); return; }

    const today = todayISO();
    const start = date && date >= today ? date : today;
    const end = endDate && endDate >= start ? endDate : start;
    state.editing = editRes || null;
    state.preselectVan = editRes ? editRes.van : (vanId || null);
    const started = !!(editRes && editRes.startDate < today);

    $("#bookForm").reset();
    $("#formError").textContent = "";
    $("#bookTitle").textContent = editRes ? "Change reservation " + editRes.id : "Reserve a van";
    $("#submitBtn").textContent = editRes ? "Save changes" : "Reserve";
    $("#startDate").min = started ? editRes.startDate : today;
    $("#startDate").value = editRes ? editRes.startDate : start;
    $("#endDate").value = editRes ? editRes.endDate : end;
    $("#startDate").disabled = started;
    $("#editNote").hidden = !started;
    $("#editNote").textContent = started ? "This trip has already started, so only the end date and details can be changed." : "";

    const prefs = storageGet(PREFS_KEY) || {};
    if (prefs.team) $("#team").value = prefs.team;
    if (prefs.otherTeam) $("#otherTeam").value = prefs.otherTeam;
    $("#otherTeamWrap").hidden = $("#team").value !== "__other";

    if (AUTH_ON && account) {
      $("#emailWrap").hidden = true;
      $("#signedInAs").hidden = false;
      $("#signedInAs").textContent = "Reserving as " + (account.name || "") + " (" + signedInEmail() + ")";
      $("#driver").value = prefs.driver || account.name || "";
    } else {
      $("#emailWrap").hidden = false;
      $("#signedInAs").hidden = true;
      $("#driver").value = prefs.driver || "";
      $("#email").value = prefs.email || "";
    }

    if (editRes) {
      const known = (CFG.teams || []).indexOf(editRes.team) !== -1;
      $("#team").value = known ? editRes.team : "__other";
      $("#otherTeam").value = known ? "" : editRes.team;
      $("#otherTeamWrap").hidden = known;
      $("#driver").value = editRes.driver || "";
      $("#destination").value = editRes.destination || "";
      $("#notes").value = editRes.notes || "";
      if (!AUTH_ON && editRes.email) $("#email").value = editRes.email;
    }

    updateChoices();
    if (started) lockVan(editRes.van);
    $("#bookDialog").showModal();
  }

  async function submitBooking(ev) {
    ev.preventDefault();
    const err = $("#formError");
    const btn = $("#submitBtn");
    err.textContent = "";

    const s = $("#startDate").value;
    const e = $("#endDate").value;
    const van = selectedVan();
    let team = $("#team").value;
    if (team === "__other") team = $("#otherTeam").value.trim();
    const driver = $("#driver").value.trim();
    const email = AUTH_ON && account ? signedInEmail() : $("#email").value.trim().toLowerCase();
    const today = todayISO();

    const problems = [];
    if (!s || !e) problems.push("Choose a start and end date.");
    else if (s < today && !(state.editing && state.editing.startDate === s)) problems.push("The start date cannot be in the past.");
    else if (e < s) problems.push("The end date must be on or after the start date.");
    else if (CFG.maxDays && dayCount(s, e) > CFG.maxDays) problems.push("Reservations can be at most " + CFG.maxDays + " days.");
    if (!van) problems.push("Choose a van.");
    if (!team) problems.push("Choose your team.");
    if (!driver) problems.push("Enter the driver's name.");
    if (!/^[^@\s]+@[^@\s]+$/.test(email) || (DOMAIN && !email.endsWith("@" + DOMAIN))) {
      problems.push(AUTH_ON ? "Please sign in with your W&L account." : "Enter a valid @" + DOMAIN + " email address.");
    }
    if (problems.length) { err.textContent = problems.join(" "); return; }

    btn.disabled = true;
    btn.textContent = "Checking...";
    try {
      // Get the latest reservations so nobody books over someone else
      await load(true);
      const editing = state.editing;
      if (!isAvailable(van, s, e, editing ? editing.id : null)) {
        err.textContent = "Sorry, " + vanName(van) + SORRY;
        updateChoices();
        return;
      }

      btn.textContent = editing ? "Saving..." : "Reserving...";
      const payload = {
        id: editing ? editing.id : undefined,
        van: van,
        startDate: s,
        endDate: e,
        team: team,
        driver: driver,
        email: email,
        reservedBy: account ? (account.name || email) : driver,
        destination: $("#destination").value.trim(),
        notes: $("#notes").value.trim()
      };
      const result = await api(editing ? "update" : "book", payload);
      if (!result.ok) {
        err.textContent = result.message || ("Sorry, " + vanName(van) + SORRY);
        await load(true);
        updateChoices();
        return;
      }

      storageSet(PREFS_KEY, {
        team: $("#team").value,
        otherTeam: $("#otherTeam").value.trim(),
        driver: driver,
        email: AUTH_ON ? "" : email
      });
      $("#bookDialog").close();
      toast((editing ? "Saved changes to " + editing.id + ": " : "Reserved: ") + vanName(van) + " for " + team + ", " + fmtRange(s, e) +
        (!editing && result.id ? ". Reservation " + result.id + "." : "."));
      await load(true);
      if (editing) openMine();
    } catch (x) {
      console.error(x);
      err.textContent = "Something went wrong and the reservation was not saved. Please try again.";
    } finally {
      btn.disabled = false;
      btn.textContent = state.editing ? "Save changes" : "Reserve";
    }
  }

  /* ================= My reservations (change or cancel) ================= */

  function mineEmail() {
    if (AUTH_ON && account) return signedInEmail();
    return String($("#mineEmail").value || (storageGet(PREFS_KEY) || {}).email || "").trim().toLowerCase();
  }

  function openMine() {
    if (AUTH_ON && !account) { signIn(); return; }
    $("#mineList").innerHTML = "";
    $("#mineStatus").textContent = "";
    $("#mineLookup").hidden = AUTH_ON;
    if (!AUTH_ON) $("#mineEmail").value = (storageGet(PREFS_KEY) || {}).email || "";
    if (!$("#mineDialog").open) $("#mineDialog").showModal();
    if (AUTH_ON || $("#mineEmail").value) refreshMine();
  }

  async function refreshMine() {
    const email = mineEmail();
    if (!AUTH_ON && (!email || (DOMAIN && !email.endsWith("@" + DOMAIN)))) {
      $("#mineStatus").textContent = "Enter the @" + DOMAIN + " email you used when you booked.";
      return;
    }
    $("#mineStatus").textContent = "Loading...";
    try {
      const res = await api("mine", { email: email });
      if (!res.ok) { $("#mineStatus").textContent = res.message || "Could not load your reservations."; return; }
      state.mine = (res.reservations || []).map(function (r) {
        return Object.assign(normalize(r), { notes: r.notes || "", email: r.email || "" });
      });
      renderMine(res.admin);
    } catch (x) {
      console.error(x);
      $("#mineStatus").textContent = "Could not load your reservations. Please try again.";
    }
  }

  function renderMine(admin) {
    const today = todayISO();
    $("#mineStatus").textContent = state.mine.length
      ? (admin ? "Admin view: all upcoming reservations." : "Your upcoming reservations:")
      : "You have no upcoming reservations.";
    $("#mineList").innerHTML = state.mine.map(function (r) {
      const started = r.startDate < today;
      return '<li class="mine-item">' +
        "<div><strong>" + esc(vanName(r.van)) + "</strong>, " + esc(fmtRange(r.startDate, r.endDate)) +
        (started ? ' <span class="pill busy">In use</span>' : "") + "</div>" +
        '<div class="hint">' + esc(r.team) + (r.destination ? ", " + esc(r.destination) : "") +
        ' <span class="rid">' + esc(r.id) + "</span>" + (admin && r.email ? " · " + esc(r.email) : "") + "</div>" +
        '<div class="mine-actions">' +
        '<button type="button" class="btn small" data-edit="' + esc(r.id) + '">Change</button>' +
        (started && !admin ? "" : '<button type="button" class="btn ghost" data-cancel="' + esc(r.id) + '">Cancel reservation</button>') +
        "</div></li>";
    }).join("");
  }

  async function cancelReservation(id) {
    const r = state.mine.find(function (x) { return x.id === id; });
    if (!r) return;
    const ok = window.confirm("Cancel " + vanName(r.van) + " for " + r.team + ", " + fmtRange(r.startDate, r.endDate) +
      "?\n\nThe van will become available to others for those dates.");
    if (!ok) return;
    $("#mineStatus").textContent = "Cancelling...";
    try {
      const res = await api("cancel", { id: id, email: mineEmail() });
      if (!res.ok) { $("#mineStatus").textContent = res.message || "Could not cancel. Please try again."; return; }
      toast("Cancelled " + id + ". " + vanName(r.van) + " is open again for those dates.");
      await load(true);
      await refreshMine();
    } catch (x) {
      console.error(x);
      $("#mineStatus").textContent = "Could not cancel. Please try again.";
    }
  }

  /* ================= Events ================= */

  function wire() {
    document.addEventListener("click", function (ev) {
      const reserve = ev.target.closest("[data-reserve]");
      if (reserve) { openBooking(reserve.dataset.reserve, reserve.dataset.date, reserve.dataset.end); return; }

      const edit = ev.target.closest("[data-edit]");
      if (edit) {
        const r = state.mine.find(function (x) { return x.id === edit.dataset.edit; });
        if (r) { $("#mineDialog").close(); openBooking(r.van, r.startDate, r.endDate, r); }
        return;
      }
      const cancel = ev.target.closest("[data-cancel]");
      if (cancel) { cancelReservation(cancel.dataset.cancel); return; }

      const day = ev.target.closest("[data-day]");
      if (day) {
        state.selectedDay = day.dataset.day;
        if (state.selectedDay.slice(0, 7) !== state.month.slice(0, 7)) state.month = state.selectedDay.slice(0, 8) + "01";
        renderCalendar();
        renderDay();
        const again = document.querySelector('[data-day="' + state.selectedDay + '"]');
        if (again) again.focus();
        return;
      }

      const toggle = ev.target.closest(".toggle");
      if (toggle) {
        state.view = toggle.dataset.view;
        storageSet(PREFS_KEY + "-view", state.view);
        render();
        return;
      }

      if (ev.target.id === "signInBtn") signIn();
      if (ev.target.id === "signOutBtn") signOut();
    });

    $("#reserveBtn").addEventListener("click", function () {
      if (state.view === "calendar") openBooking(null, state.selectedDay);
      else openBooking(null, state.listFrom, state.listTo);
    });
    $("#listFrom").addEventListener("change", function (e) {
      if (!e.target.value) return;
      state.listFrom = e.target.value;
      if (state.listTo < state.listFrom) state.listTo = state.listFrom;
      renderList();
    });
    $("#listTo").addEventListener("change", function (e) {
      if (!e.target.value) return;
      state.listTo = e.target.value < state.listFrom ? state.listFrom : e.target.value;
      renderList();
    });
    $("#mineBtn").addEventListener("click", openMine);
    $("#closeMine").addEventListener("click", function () { $("#mineDialog").close(); });
    $("#mineLookup").addEventListener("submit", function (e) {
      e.preventDefault();
      const prefs = storageGet(PREFS_KEY) || {};
      prefs.email = $("#mineEmail").value.trim().toLowerCase();
      storageSet(PREFS_KEY, prefs);
      refreshMine();
    });
    $("#prevMonth").addEventListener("click", function () { state.month = addMonths(state.month, -1); renderCalendar(); });
    $("#nextMonth").addEventListener("click", function () { state.month = addMonths(state.month, 1); renderCalendar(); });
    $("#todayBtn").addEventListener("click", function () {
      state.selectedDay = todayISO();
      state.month = state.selectedDay.slice(0, 8) + "01";
      renderCalendar();
      renderDay();
    });

    $("#startDate").addEventListener("change", updateChoices);
    $("#endDate").addEventListener("change", updateChoices);
    $("#team").addEventListener("change", function () {
      $("#otherTeamWrap").hidden = $("#team").value !== "__other";
    });
    $("#bookForm").addEventListener("submit", submitBooking);
    $("#cancelBtn").addEventListener("click", function () { $("#bookDialog").close(); });
    $("#closeDialog").addEventListener("click", function () { $("#bookDialog").close(); });
    $("#bookDialog").addEventListener("close", function () {
      state.editing = null;
      $("#startDate").disabled = false;
      $("#editNote").hidden = true;
    });

    // Keep the page current (returns show up within a minute)
    setInterval(function () { if (!document.hidden) load(true); }, REFRESH_MS);
    document.addEventListener("visibilitychange", function () { if (!document.hidden) load(true); });
  }

  /* ================= Start ================= */

  async function start() {
    $("#demoBanner").hidden = !DEMO;
    const savedView = storageGet(PREFS_KEY + "-view");
    if (savedView === "list" || savedView === "calendar") state.view = savedView;
    fillTeams();
    wire();
    render();
    await load(false);
    await initAuth();
  }

  start();
})();
