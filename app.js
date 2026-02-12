/* ==========================================================
   PTC Comms Telemetry (Static GitHub Pages)
   - Outlook Inbox unread count via Graph
   - Slack/HubSpot/Monday via Outlook folders (unreadItemCount)
   - Counts only (Mail.ReadBasic), no message bodies
   ========================================================== */

const CONFIG = {
  tenantId: "bcfdd46a-c2dd-4e71-a9f8-5cd31816ff9e",
  clientId: "cf321f12-ce1d-4067-b44e-05fafad8258d",

  folders: {
    slack:   { name: "PTC - Slack Alerts" },
    hubspot: { name: "PTC - HubSpot Alerts" },
    monday:  { name: "PTC - Monday Alerts" }
  },

  // Tuned for company-wide use (adjust anytime)
  scale: {
    max: {
      outlook: 300,
      slack: 100,
      hubspot: 100,
      monday: 100,
      total: 500
    },
    redline: {
      outlook: 200,
      slack: 30,
      hubspot: 30,
      monday: 30,
      total: 280
    }
  },

  links: {
    outlook: "https://outlook.office.com/mail/",
    slack:   "https://app.slack.com/client/",
    hubspot: "https://app.hubspot.com/",
    monday:  "https://monday.com/"
  },

  autoRefreshSeconds: 60
};

const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["Mail.ReadBasic", "User.Read"];

let msalApp;
let activeAccount = null;
let autoTimer = null;

const $ = (id) => document.getElementById(id);
function nowStr(){ return new Date().toLocaleString(); }
function clamp(n,a,b){ return Math.max(a, Math.min(b,n)); }
function pct(value, max){ return clamp((value || 0) / Math.max(1,max), 0, 1); }

function setButtons(signedIn){
  $("btnRefresh").disabled = !signedIn;
  $("btnBaseline").disabled = !signedIn;
  $("btnAuto").disabled = !signedIn;
  $("btnSignOut").disabled = !signedIn;
}
function setMode(t){ $("mode").textContent = t; }
function setDriver(t){ $("driver").textContent = t || "—"; }
function setUpdateTime(t){ $("lastUpdate").textContent = t || "—"; }
function setPollMeta(pollText){ $("pollMeta").textContent = pollText; }

function baselineKey(){
  const id = activeAccount?.homeAccountId || "anon";
  return `ptc_comms_baseline_${id}`;
}
function getBaseline(){
  try{
    const raw = localStorage.getItem(baselineKey());
    return raw ? JSON.parse(raw) : null;
  }catch{ return null; }
}
function setBaseline(obj){
  localStorage.setItem(baselineKey(), JSON.stringify(obj));
}
function setBaselineUI(){
  const b = getBaseline();
  $("baselineTime").textContent = b?.time ? new Date(b.time).toLocaleString() : "—";
}
function setBaselineFromCurrent(){
  const current = {};
  for (const k of ["outlook","slack","hubspot","monday"]){
    const v = parseInt($(`val_${k}`).textContent, 10);
    current[k] = Number.isFinite(v) ? v : 0;
  }
  setBaseline({ time: Date.now(), counts: current });
  setBaselineUI();
  refreshAll().catch(() => {});
}

function setAppLinks(){
  $("link_outlook").href = CONFIG.links.outlook;
  $("link_slack").href   = CONFIG.links.slack;
  $("link_hubspot").href = CONFIG.links.hubspot;
  $("link_monday").href  = CONFIG.links.monday;
}

/* =========================
   Graph helpers
   ========================= */
async function graphGet(path, token){
  const res = await fetch(`${GRAPH}${path}`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
  return res.json();
}
async function getToken(){
  const req = { scopes: SCOPES, account: activeAccount };
  try{
    const r = await msalApp.acquireTokenSilent(req);
    return r.accessToken;
  }catch{
    const r = await msalApp.acquireTokenPopup(req);
    return r.accessToken;
  }
}
async function getInboxUnread(token){
  const data = await graphGet(`/me/mailFolders('inbox')?$select=unreadItemCount`, token);
  return data.unreadItemCount ?? 0;
}
async function getFolderCandidates(token){
  const [top, inboxKids] = await Promise.all([
    graphGet(`/me/mailFolders?$select=displayName,id,unreadItemCount&$top=200`, token).catch(() => ({ value: [] })),
    graphGet(`/me/mailFolders('inbox')/childFolders?$select=displayName,id,unreadItemCount&$top=200`, token).catch(() => ({ value: [] }))
  ]);

  const list = [...(top.value || []), ...(inboxKids.value || [])];
  const seen = new Set();
  const out = [];
  for (const f of list){
    if (!f?.id || seen.has(f.id)) continue;
    seen.add(f.id);
    out.push(f);
  }
  return out;
}
function findFolderCount(folders, name){
  const target = (name || "").trim().toLowerCase();
  const match = folders.find(f => (f.displayName || "").trim().toLowerCase() === target);
  return match?.unreadItemCount ?? 0;
}

/* =========================
   Rendering
   ========================= */
function computeTotals(c){
  return (c.outlook||0)+(c.slack||0)+(c.hubspot||0)+(c.monday||0);
}
function computeDeltas(counts){
  const b = getBaseline();
  if (!b?.counts) return { deltas: null };

  const deltas = {};
  for (const k of ["outlook","slack","hubspot","monday"]){
    const base = Number.isFinite(b.counts[k]) ? b.counts[k] : 0;
    const cur  = Number.isFinite(counts[k]) ? counts[k] : 0;
    deltas[k] = cur - base;
  }
  deltas.total = deltas.outlook + deltas.slack + deltas.hubspot + deltas.monday;
  return { deltas };
}

function statusFor(value, redline){
  const v = Number.isFinite(value) ? value : 0;
  const warnAt = Math.ceil(redline * 0.60);

  if (v >= redline) return { level: "bad",  label: "CRITICAL" };
  if (v >= warnAt)  return { level: "warn", label: "ELEVATED" };
  if (v > 0)        return { level: "ok",   label: "OK" };
  return { level: "idle", label: "IDLE" };
}

function setCard(key, value, delta, max, redline, isTotal=false){
  const card = $(`card_${key}`);
  const statusEl = $(`status_${key}`);
  const valEl = $(`val_${key}`);
  const metaEl = $(`meta_${key}`);

  if (!card || !statusEl || !valEl || !metaEl) return;

  const v = Number.isFinite(value) ? value : 0;
  const d = (delta == null) ? null : delta;

  const s = statusFor(v, redline);
  card.dataset.status = s.level;
  statusEl.textContent = s.label;

  card.style.setProperty("--p", pct(v, max));

  valEl.textContent = String(v);

  if (d == null){
    metaEl.textContent = isTotal ? "New since baseline: —" : "Since baseline: —";
  }else{
    const sign = d > 0 ? "+" : "";
    metaEl.textContent = isTotal ? `New since baseline: ${sign}${d}` : `Since baseline: ${sign}${d}`;
  }
}

function updateMode(total){
  const red = CONFIG.scale.redline.total;
  if (total >= red) setMode("REDLINE");
  else if (total >= Math.ceil(red * 0.55)) setMode("BUSY");
  else if (total > 0) setMode("ACTIVE");
  else setMode("IDLE");
}

/* =========================
   Refresh
   ========================= */
async function refreshAll(){
  setPollMeta("POLL: PULLING • API: —");

  const t0 = performance.now();
  try{
    const token = await getToken();

    const inboxCount = await getInboxUnread(token);
    const folders = await getFolderCandidates(token);

    const counts = {
      outlook: inboxCount,
      slack:   findFolderCount(folders, CONFIG.folders.slack.name),
      hubspot: findFolderCount(folders, CONFIG.folders.hubspot.name),
      monday:  findFolderCount(folders, CONFIG.folders.monday.name)
    };
    counts.total = computeTotals(counts);

    const { deltas } = computeDeltas(counts);

    // Mini breakdown under TOTAL
    $("mini_outlook").textContent = String(counts.outlook);
    $("mini_slack").textContent = String(counts.slack);
    $("mini_hubspot").textContent = String(counts.hubspot);
    $("mini_monday").textContent = String(counts.monday);

    // Render cards
    setCard("total", counts.total, deltas?.total ?? null, CONFIG.scale.max.total, CONFIG.scale.redline.total, true);
    setCard("outlook", counts.outlook, deltas?.outlook ?? null, CONFIG.scale.max.outlook, CONFIG.scale.redline.outlook);
    setCard("slack", counts.slack, deltas?.slack ?? null, CONFIG.scale.max.slack, CONFIG.scale.redline.slack);
    setCard("hubspot", counts.hubspot, deltas?.hubspot ?? null, CONFIG.scale.max.hubspot, CONFIG.scale.redline.hubspot);
    setCard("monday", counts.monday, deltas?.monday ?? null, CONFIG.scale.max.monday, CONFIG.scale.redline.monday);

    updateMode(counts.total);
    setUpdateTime(nowStr());

    const ms = Math.round(performance.now() - t0);
    setPollMeta(`POLL: LIVE • API: ${ms}ms`);
  }catch(e){
    console.error(e);
    setMode("CHECK");
    setUpdateTime(nowStr());
    setPollMeta("POLL: ERROR • API: —");
    // Keep last-known values on screen (don’t blank the dashboard)
  }
}

/* =========================
   Auth
   ========================= */
async function signIn(){
  const loginReq = { scopes: SCOPES };

  await msalApp.handleRedirectPromise().catch(() => null);

  const existing = msalApp.getAllAccounts();
  if (existing?.length){
    activeAccount = existing[0];
    msalApp.setActiveAccount(activeAccount);
    setButtons(true);
    setDriver(activeAccount.username);
    setBaselineUI();
    await refreshAll();
    return;
  }

  try{
    const result = await msalApp.loginPopup(loginReq);
    activeAccount = result.account;
  }catch{
    await msalApp.loginRedirect(loginReq);
    return;
  }

  msalApp.setActiveAccount(activeAccount);
  setButtons(true);
  setDriver(activeAccount.username);
  setBaselineUI();
  await refreshAll();
}

async function signOut(){
  const acc = activeAccount;
  activeAccount = null;

  if (autoTimer){
    clearInterval(autoTimer);
    autoTimer = null;
    $("btnAuto").textContent = "AUTO: OFF";
  }

  setButtons(false);
  setDriver("—");
  setUpdateTime("—");
  setBaselineUI();
  setMode("IDLE");
  setPollMeta("POLL: IDLE • API: —");

  // Clear UI
  for (const k of ["total","outlook","slack","hubspot","monday"]){
    const card = $(`card_${k}`);
    const status = $(`status_${k}`);
    const val = $(`val_${k}`);
    const meta = $(`meta_${k}`);
    if (card) card.dataset.status = "idle";
    if (status) status.textContent = "IDLE";
    if (val) val.textContent = "—";
    if (meta) meta.textContent = (k === "total") ? "New since baseline: —" : "Since baseline: —";
    if (card) card.style.setProperty("--p", 0);
  }

  for (const k of ["mini_outlook","mini_slack","mini_hubspot","mini_monday"]){
    const el = $(k);
    if (el) el.textContent = "—";
  }

  if (acc){
    await msalApp.logoutPopup({ account: acc }).catch(() => {});
  }
}

function toggleAuto(){
  if (autoTimer){
    clearInterval(autoTimer);
    autoTimer = null;
    $("btnAuto").textContent = "AUTO: OFF";
    return;
  }
  autoTimer = setInterval(() => refreshAll().catch(() => {}), CONFIG.autoRefreshSeconds * 1000);
  $("btnAuto").textContent = `AUTO: ${CONFIG.autoRefreshSeconds}s`;
}

/* =========================
   Init
   ========================= */
window.addEventListener("load", () => {
  setAppLinks();

  const msalConfig = {
    auth: {
      clientId: CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: "sessionStorage" }
  };
  msalApp = new msal.PublicClientApplication(msalConfig);

  $("btnIgnition").onclick = () => signIn().catch(console.error);
  $("btnRefresh").onclick  = () => refreshAll().catch(console.error);
  $("btnBaseline").onclick = () => setBaselineFromCurrent();
  $("btnAuto").onclick     = () => toggleAuto();
  $("btnSignOut").onclick  = () => signOut().catch(console.error);

  setButtons(false);
  setDriver("—");
  setUpdateTime("—");
  setBaselineUI();
  setMode("IDLE");
  setPollMeta("POLL: IDLE • API: —");
});
