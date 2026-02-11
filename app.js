/* ====== CONFIG: EDIT THESE ====== */
const CONFIG = {
  tenantId: "bcfdd46a-c2dd-4e71-a9f8-5cd31816ff9e",
  clientId: "cf321f12-ce1d-4067-b44e-05fafad8258d",

  // Folder display names under Inbox (create these folders + Outlook rules)
  folders: {
    outlook: { special: "inbox" },
    slack:   { name: "PTC - Slack Alerts" },
    hubspot: { name: "PTC - HubSpot Alerts" },
    monday:  { name: "PTC - Monday Alerts" }
  },

  // Gauge scaling / redlines (tune this to feel right)
  scale: {
    max:     { outlook: 60, slack: 25, hubspot: 25, monday: 25, total: 100 },
    redline: { outlook: 25, slack: 10, hubspot: 10, monday: 10, total: 35 }
  },

  // App links (purely “open app”, not data)
  links: {
    outlook: "https://outlook.office.com/mail/",
    slack:   "https://app.slack.com/client/",
    hubspot: "https://app.hubspot.com/",
    monday:  "https://monday.com/"
  },

  autoRefreshSeconds: 60
};
/* ====== END CONFIG ====== */


const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["Mail.ReadBasic"]; // Minimal for folder unread counts

let msalApp;
let activeAccount = null;
let autoTimer = null;

const $ = (id) => document.getElementById(id);
const byKey = (key) => ({
  gauge: document.querySelector(`.gauge[data-key="${key}"]`),
  value: $(`val_${key}`),
  meta:  $(`meta_${key}`)
});

function nowStr() {
  return new Date().toLocaleString();
}

function clamp(n, a, b) {
  return Math.max(a, Math.min(b, n));
}

function pct(value, max) {
  if (value == null || !Number.isFinite(value)) return 0;
  return clamp(value / Math.max(1, max), 0, 1);
}

// Map 0..1 -> -120..+120 degrees
function needleAngle(p) {
  const min = -120, max = 120;
  const deg = min + (max - min) * p;
  return `${deg}deg`;
}

function zoneColor(value, redline) {
  if (value == null || !Number.isFinite(value)) return "rgba(255,255,255,0.6)";
  if (value >= redline) return "var(--ringRed)";
  if (value >= Math.ceil(redline * 0.6)) return "var(--ringWarn)";
  return "var(--ringFill)";
}

function setLight(name, on) {
  const el = document.querySelector(`.light[data-light="${name}"]`);
  if (!el) return;
  el.classList.toggle("on", !!on);
}

function setButtons(signedIn) {
  $("btnRefresh").disabled = !signedIn;
  $("btnBaseline").disabled = !signedIn;
  $("btnAuto").disabled = !signedIn;
  $("btnSignOut").disabled = !signedIn;
}

function setMode(text) {
  $("mode").textContent = text;
}

function setDriver(text) {
  $("driver").textContent = text || "—";
}

function setUpdateTime(text) {
  $("lastUpdate").textContent = text || "—";
}

function baselineKey() {
  const id = activeAccount?.homeAccountId || "anon";
  return `ptc_gauge_baseline_${id}`;
}

function getBaseline() {
  try {
    const raw = localStorage.getItem(baselineKey());
    return raw ? JSON.parse(raw) : null;
  } catch { return null; }
}

function setBaseline(obj) {
  localStorage.setItem(baselineKey(), JSON.stringify(obj));
}

function setBaselineUI() {
  const b = getBaseline();
  $("baselineTime").textContent = b?.time ? new Date(b.time).toLocaleString() : "—";
}

async function graphGet(path, token) {
  const res = await fetch(`${GRAPH}${path}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
  return res.json();
}

async function getToken() {
  const req = { scopes: SCOPES, account: activeAccount };
  try {
    const r = await msalApp.acquireTokenSilent(req);
    return r.accessToken;
  } catch {
    // Popup is simplest on Pages
    const r = await msalApp.acquireTokenPopup(req);
    return r.accessToken;
  }
}

async function getInboxUnread(token) {
  const data = await graphGet(`/me/mailFolders('inbox')?$select=unreadItemCount`, token);
  return data.unreadItemCount ?? 0;
}

async function getInboxChildFolders(token) {
  // One call is usually enough; if you have a lot of folders, raise $top or extend pagination.
  const data = await graphGet(
    `/me/mailFolders('inbox')/childFolders?$select=displayName,id,unreadItemCount&$top=200`,
    token
  );
  return data.value || [];
}

function findFolderCount(childFolders, name) {
  const match = childFolders.find(f => (f.displayName || "").toLowerCase() === name.toLowerCase());
  return match?.unreadItemCount ?? 0;
}

function renderGauge(key, value, delta, max, redline) {
  const el = byKey(key);
  if (!el.gauge) return;

  const p = pct(value, max);
  el.gauge.style.setProperty("--pct", p);
  el.gauge.style.setProperty("--angle", needleAngle(p));
  el.gauge.style.setProperty("--fill", zoneColor(value, redline));

  if (el.value) el.value.textContent = (value == null ? "—" : String(value));

  if (el.meta) {
    if (delta == null) el.meta.textContent = (key === "total") ? "New since baseline: —" : "Since baseline: —";
    else {
      const sign = delta > 0 ? "+" : "";
      el.meta.textContent = (key === "total")
        ? `New since baseline: ${sign}${delta}`
        : `Since baseline: ${sign}${delta}`;
    }
  }
}

function updateLights(counts) {
  // Light on if >0, and CHECK if any missing
  setLight("outlook", (counts.outlook ?? 0) > 0);
  setLight("slack",   (counts.slack ?? 0) > 0);
  setLight("hubspot", (counts.hubspot ?? 0) > 0);
  setLight("monday",  (counts.monday ?? 0) > 0);
}

function computeTotals(counts) {
  return (counts.outlook ?? 0) + (counts.slack ?? 0) + (counts.hubspot ?? 0) + (counts.monday ?? 0);
}

function computeDeltas(counts) {
  const b = getBaseline();
  if (!b?.counts) return { deltas: null, baseline: null };

  const deltas = {};
  for (const k of ["outlook","slack","hubspot","monday"]) {
    const base = Number.isFinite(b.counts[k]) ? b.counts[k] : 0;
    const cur = Number.isFinite(counts[k]) ? counts[k] : 0;
    deltas[k] = cur - base;
  }
  deltas.total = (deltas.outlook + deltas.slack + deltas.hubspot + deltas.monday);
  return { deltas, baseline: b };
}

function setAppLinks() {
  $("link_outlook").href = CONFIG.links.outlook;
  $("link_slack").href = CONFIG.links.slack;
  $("link_hubspot").href = CONFIG.links.hubspot;
  $("link_monday").href = CONFIG.links.monday;
}

function updateBusyMode(total) {
  const red = CONFIG.scale.redline.total;
  if (total >= red) setMode("REDLINE");
  else if (total >= Math.ceil(red * 0.55)) setMode("BUSY");
  else if (total > 0) setMode("ACTIVE");
  else setMode("IDLE");
}

async function refreshAll() {
  setLight("check", false);

  try {
    const token = await getToken();

    const inboxCount = await getInboxUnread(token);
    const childFolders = await getInboxChildFolders(token);

    const counts = {
      outlook: inboxCount,
      slack: findFolderCount(childFolders, CONFIG.folders.slack.name),
      hubspot: findFolderCount(childFolders, CONFIG.folders.hubspot.name),
      monday: findFolderCount(childFolders, CONFIG.folders.monday.name)
    };

    const total = computeTotals(counts);
    counts.total = total;

    // deltas
    const { deltas } = computeDeltas(counts);

    // render gauges
    renderGauge("outlook", counts.outlook, deltas?.outlook ?? null, CONFIG.scale.max.outlook, CONFIG.scale.redline.outlook);
    renderGauge("slack", counts.slack, deltas?.slack ?? null, CONFIG.scale.max.slack, CONFIG.scale.redline.slack);
    renderGauge("hubspot", counts.hubspot, deltas?.hubspot ?? null, CONFIG.scale.max.hubspot, CONFIG.scale.redline.hubspot);
    renderGauge("monday", counts.monday, deltas?.monday ?? null, CONFIG.scale.max.monday, CONFIG.scale.redline.monday);
    renderGauge("total", counts.total, deltas?.total ?? null, CONFIG.scale.max.total, CONFIG.scale.redline.total);

    updateLights(counts);
    updateBusyMode(total);

    setUpdateTime(nowStr());
  } catch (e) {
    console.error(e);
    setLight("check", true);
    setMode("CHECK");
    setUpdateTime(nowStr());
  }
}

async function signIn() {
  const loginReq = { scopes: SCOPES };

  // Handle redirect response if any (safe to call even if none)
  await msalApp.handleRedirectPromise().catch(() => null);

  // If already signed in
  const existing = msalApp.getAllAccounts();
  if (existing?.length) {
    activeAccount = existing[0];
    msalApp.setActiveAccount(activeAccount);
    setButtons(true);
    setDriver(activeAccount.username);
    setBaselineUI();
    await refreshAll();
    return;
  }

  // Otherwise login
  try {
    const result = await msalApp.loginPopup(loginReq);
    activeAccount = result.account;
  } catch {
    // Popup blocked? Fall back to redirect
    await msalApp.loginRedirect(loginReq);
    return;
  }

  msalApp.setActiveAccount(activeAccount);
  setButtons(true);
  setDriver(activeAccount.username);
  setBaselineUI();
  await refreshAll();
}

async function signOut() {
  const acc = activeAccount;
  activeAccount = null;

  setButtons(false);
  setDriver("—");
  setMode("IDLE");
  setUpdateTime("—");
  setBaselineUI();

  // reset UI visuals
  for (const k of ["outlook","slack","hubspot","monday","total"]) {
    renderGauge(k, null, null, CONFIG.scale.max[k] ?? 100, CONFIG.scale.redline[k] ?? 10);
  }
  setLight("outlook", false);
  setLight("slack", false);
  setLight("hubspot", false);
  setLight("monday", false);
  setLight("check", false);

  if (autoTimer) {
    clearInterval(autoTimer);
    autoTimer = null;
    $("btnAuto").textContent = "AUTO: OFF";
  }

  if (acc) {
    await msalApp.logoutPopup({ account: acc }).catch(() => {});
  }
}

function setBaselineFromCurrent() {
  // Baseline is just “trip reset”
  // We store the current visible counts.
  const current = {};
  for (const k of ["outlook","slack","hubspot","monday"]) {
    const v = parseInt($(`val_${k}`).textContent, 10);
    current[k] = Number.isFinite(v) ? v : 0;
  }
  setBaseline({ time: Date.now(), counts: current });
  setBaselineUI();
  refreshAll().catch(() => {});
}

function toggleAuto() {
  if (autoTimer) {
    clearInterval(autoTimer);
    autoTimer = null;
    $("btnAuto").textContent = "AUTO: OFF";
    return;
  }
  autoTimer = setInterval(() => refreshAll().catch(() => {}), CONFIG.autoRefreshSeconds * 1000);
  $("btnAuto").textContent = `AUTO: ${CONFIG.autoRefreshSeconds}s`;
}

window.addEventListener("load", () => {
  setAppLinks();

  // MSAL init
  const msalConfig = {
    auth: {
      clientId: CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: "sessionStorage" }
  };

  msalApp = new msal.PublicClientApplication(msalConfig);

  // Buttons
  $("btnIgnition").onclick = () => signIn().catch(err => { console.error(err); setLight("check", true); });
  $("btnRefresh").onclick = () => refreshAll().catch(() => setLight("check", true));
  $("btnBaseline").onclick = () => setBaselineFromCurrent();
  $("btnAuto").onclick = () => toggleAuto();
  $("btnSignOut").onclick = () => signOut().catch(() => {});

  // Default UI state
  setButtons(false);
  setDriver("—");
  setUpdateTime("—");
  setBaselineUI();
  setMode("IDLE");

  // Initialize gauges empty
  for (const k of ["outlook","slack","hubspot","monday","total"]) {
    renderGauge(k, null, null, CONFIG.scale.max[k] ?? 100, CONFIG.scale.redline[k] ?? 10);
  }
});
