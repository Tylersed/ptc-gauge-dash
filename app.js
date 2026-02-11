/* =========================
   PTC Gauge Cluster — LIVE TELEMETRY (Static GitHub Pages)
   - Reads unread counts only (no message bodies)
   - TOTAL = Inbox unread + folder unread (Slack/HubSpot/Monday folders)
   - Waves always animate; spikes on new notifications
   ========================= */

/* ====== CONFIG ====== */
const CONFIG = {
  tenantId: "bcfdd46a-c2dd-4e71-a9f8-5cd31816ff9e",
  clientId: "cf321f12-ce1d-4067-b44e-05fafad8258d",

  folders: {
    slack:   { name: "PTC - Slack Alerts" },
    hubspot: { name: "PTC - HubSpot Alerts" },
    monday:  { name: "PTC - Monday Alerts" }
  },

  scale: {
    max:     { total: 240 },
    redline: { total: 120 }
  },

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
const SCOPES = ["Mail.ReadBasic", "User.Read"];

let msalApp;
let activeAccount = null;
let autoTimer = null;

const $ = (id) => document.getElementById(id);

function nowStr(){ return new Date().toLocaleString(); }

function setLight(name, on){
  const el = document.querySelector(`.light[data-light="${name}"]`);
  if (!el) return;
  el.classList.toggle("on", !!on);
}
function setButtons(signedIn){
  $("btnRefresh").disabled = !signedIn;
  $("btnBaseline").disabled = !signedIn;
  $("btnAuto").disabled = !signedIn;
  $("btnSignOut").disabled = !signedIn;
}
function setMode(t){ $("mode").textContent = t; }
function setDriver(t){ $("driver").textContent = t || "—"; }
function setUpdateTime(t){ $("lastUpdate").textContent = t || "—"; }

function baselineKey(){
  const id = activeAccount?.homeAccountId || "anon";
  return `ptc_gauge_baseline_${id}`;
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

/* =========================
   GRAPH HELPERS
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

/* Find folders robustly:
   - Top-level folders
   - Inbox childFolders
   (covers most setups)
*/
async function getFolderCandidates(token){
  const [top, inboxKids] = await Promise.all([
    graphGet(`/me/mailFolders?$select=displayName,id,unreadItemCount&$top=200`, token).catch(() => ({ value: [] })),
    graphGet(`/me/mailFolders('inbox')/childFolders?$select=displayName,id,unreadItemCount&$top=200`, token).catch(() => ({ value: [] }))
  ]);
  const list = [...(top.value || []), ...(inboxKids.value || [])];

  // de-dupe by id
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
   TOTAL GAUGE (center)
   ========================= */
function clamp(n,a,b){ return Math.max(a, Math.min(b,n)); }
function pct(value, max){ return clamp(value / Math.max(1,max), 0, 1); }

let _needle = 0; // smoothing
function needleAngle(p){
  _needle = _needle + (p - _needle) * 0.35; // damped
  const min = -120, max = 120;
  return `${min + (max-min) * _needle}deg`;
}
function zoneColor(value, redline){
  if (!Number.isFinite(value)) return "rgba(255,255,255,0.6)";
  if (value >= redline) return "var(--ringRed)";
  if (value >= Math.ceil(redline * 0.6)) return "var(--ringWarn)";
  return "var(--ringFill)";
}
function renderTotalGauge(total, delta){
  const gauge = document.querySelector(`.gauge[data-key="total"]`);
  if (!gauge) return;

  const max = CONFIG.scale.max.total;
  const red = CONFIG.scale.redline.total;
  const p = pct(total, max);

  gauge.style.setProperty("--pct", p);
  gauge.style.setProperty("--angle", needleAngle(p));
  gauge.style.setProperty("--fill", zoneColor(total, red));

  $("val_total").textContent = String(total ?? "—");
  if (delta == null) $("meta_total").textContent = "New since baseline: —";
  else {
    const sign = delta > 0 ? "+" : "";
    $("meta_total").textContent = `New since baseline: ${sign}${delta}`;
  }
}

/* =========================
   LIVE WAVES (always scanning)
   ========================= */

class Wave {
  constructor(key, canvasId, cardId){
    this.key = key;
    this.cv = $(canvasId);
    this.card = $(cardId);
    this.ctx = this.cv.getContext("2d", { alpha: true });

    this.t = 0;
    this.speed = 0.022 + Math.random()*0.01;
    this.ampBase = 0.18 + Math.random()*0.07;   // base motion
    this.noise = 0.05 + Math.random()*0.03;

    this.impulses = []; // {x, h, life}
    this.level = 0;     // visual intensity based on count
    this.lastValue = null;

    this.resize();
    window.addEventListener("resize", () => this.resize());
  }

  resize(){
    const dpr = Math.max(1, window.devicePixelRatio || 1);
    const rect = this.cv.getBoundingClientRect();
    this.cv.width = Math.floor(rect.width * dpr);
    this.cv.height = Math.floor(rect.height * dpr);
    this.ctx.setTransform(dpr,0,0,dpr,0,0); // draw in CSS pixels
  }

  setValue(v){
    // Smoothly influence amplitude (not reading content; just counts)
    const vv = Number.isFinite(v) ? v : 0;
    this.level = this.level + (Math.min(1, vv/40) - this.level) * 0.15;
  }

  spike(delta){
    const d = Math.max(1, Math.min(20, delta|0));
    const h = 0.55 + Math.min(1.2, d * 0.08); // bigger delta = bigger spike
    // create 1–3 impulses depending on delta
    const bursts = d >= 6 ? 3 : (d >= 3 ? 2 : 1);
    for (let i=0;i<bursts;i++){
      this.impulses.push({ x: 1.0 + i*0.03, h: h*(1 - i*0.12), life: 1.0 });
    }
    // card pulse
    if (this.card){
      this.card.classList.remove("pulse");
      // force reflow so animation re-triggers
      void this.card.offsetWidth;
      this.card.classList.add("pulse");
    }
  }

  draw(){
    const ctx = this.ctx;
    const w = this.cv.getBoundingClientRect().width;
    const h = this.cv.getBoundingClientRect().height;

    // Clear
    ctx.clearRect(0,0,w,h);

    // Background glow band
    const mid = h * 0.55;
    ctx.save();
    ctx.globalAlpha = 0.9;
    const grad = ctx.createLinearGradient(0, mid-40, 0, mid+40);
    grad.addColorStop(0, "rgba(96,165,255,0.00)");
    grad.addColorStop(0.5, "rgba(96,165,255,0.06)");
    grad.addColorStop(1, "rgba(96,165,255,0.00)");
    ctx.fillStyle = grad;
    ctx.fillRect(0,0,w,h);
    ctx.restore();

    // Axis line
    ctx.save();
    ctx.strokeStyle = "rgba(255,255,255,0.10)";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(0, mid);
    ctx.lineTo(w, mid);
    ctx.stroke();
    ctx.restore();

    // Wave parameters
    const baseAmp = (h * 0.20) * (this.ampBase + this.level * 0.55);
    const freq = 0.035 + this.level * 0.015;

    // Move impulses left
    for (const imp of this.impulses){
      imp.x -= 0.016;     // left motion
      imp.life -= 0.012;  // decay
    }
    this.impulses = this.impulses.filter(i => i.life > 0 && i.x > -0.2);

    // Draw neon line (glow + core)
    const drawLine = (alpha, width, color) => {
      ctx.save();
      ctx.globalAlpha = alpha;
      ctx.lineWidth = width;
      ctx.strokeStyle = color;
      ctx.beginPath();

      for (let x=0; x<=w; x++){
        const nx = x / w;

        // base wave
        const s1 = Math.sin((this.t + nx*12) * (1/freq)) * 0.55;
        const s2 = Math.sin((this.t*1.8 + nx*18) * (1/(freq*1.2))) * 0.35;

        // noise
        const n = (Math.sin((this.t*3.1 + nx*44)) * this.noise);

        // impulse spike (Gaussian-ish peak)
        let spike = 0;
        for (const imp of this.impulses){
          const dx = (nx - imp.x);
          spike += Math.exp(-(dx*dx) / 0.0025) * imp.h * imp.life;
        }

        const y = mid + (s1+s2+n) * baseAmp - spike * (h*0.28);

        if (x === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      }

      ctx.stroke();
      ctx.restore();
    };

    // color shifts when level is high
    const accent = (this.level > 0.75) ? "rgba(255,204,102,0.95)"
                 : (this.level > 0.90) ? "rgba(255,77,109,0.95)"
                 : "rgba(96,165,255,0.95)";

    // outer glow
    drawLine(0.20, 6, "rgba(96,165,255,0.75)");
    drawLine(0.16, 10, "rgba(96,165,255,0.35)");
    // core
    drawLine(0.95, 2.2, accent);

    // Update time
    this.t += this.speed;
  }
}

const WAVES = {
  outlook: null,
  slack: null,
  hubspot: null,
  monday: null
};

function initWaves(){
  WAVES.outlook = new Wave("outlook", "cv_outlook", "w_outlook");
  WAVES.slack   = new Wave("slack",   "cv_slack",   "w_slack");
  WAVES.hubspot = new Wave("hubspot", "cv_hubspot", "w_hubspot");
  WAVES.monday  = new Wave("monday",  "cv_monday",  "w_monday");
}

function waveLoop(){
  for (const k of ["outlook","slack","hubspot","monday"]){
    WAVES[k]?.draw();
  }
  requestAnimationFrame(waveLoop);
}

/* =========================
   COUNTS + DELTAS + UI
   ========================= */

function computeTotals(c){
  return (c.outlook||0)+(c.slack||0)+(c.hubspot||0)+(c.monday||0);
}
function computeDeltas(counts){
  const b = getBaseline();
  if (!b?.counts) return { deltas: null, baseline: null };
  const deltas = {};
  for (const k of ["outlook","slack","hubspot","monday"]){
    const base = Number.isFinite(b.counts[k]) ? b.counts[k] : 0;
    const cur  = Number.isFinite(counts[k]) ? counts[k] : 0;
    deltas[k] = cur - base;
  }
  deltas.total = deltas.outlook + deltas.slack + deltas.hubspot + deltas.monday;
  return { deltas, baseline: b };
}

function updateLights(counts){
  setLight("outlook", (counts.outlook||0) > 0);
  setLight("slack",   (counts.slack||0) > 0);
  setLight("hubspot", (counts.hubspot||0) > 0);
  setLight("monday",  (counts.monday||0) > 0);
}

function updateBusyMode(total){
  const red = CONFIG.scale.redline.total;
  if (total >= red) setMode("REDLINE");
  else if (total >= Math.ceil(red * 0.55)) setMode("BUSY");
  else if (total > 0) setMode("ACTIVE");
  else setMode("IDLE");
}

function setAppLinks(){
  $("link_outlook").href = CONFIG.links.outlook;
  $("link_slack").href   = CONFIG.links.slack;
  $("link_hubspot").href = CONFIG.links.hubspot;
  $("link_monday").href  = CONFIG.links.monday;
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

/* =========================
   REFRESH LOOP (data pulls)
   ========================= */
let lastCounts = { outlook: null, slack: null, hubspot: null, monday: null };

async function refreshAll(){
  setLight("check", false);
  $("pollState").textContent = "PULLING";

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

    // Write numbers
    $("val_outlook").textContent = String(counts.outlook);
    $("val_slack").textContent = String(counts.slack);
    $("val_hubspot").textContent = String(counts.hubspot);
    $("val_monday").textContent = String(counts.monday);

    // Deltas
    const { deltas } = computeDeltas(counts);
    const fmt = (v, key) => {
      if (v == null) return key === "total" ? "New since baseline: —" : "Since baseline: —";
      const sign = v > 0 ? "+" : "";
      return key === "total" ? `New since baseline: ${sign}${v}` : `Since baseline: ${sign}${v}`;
    };

    $("meta_outlook").textContent = fmt(deltas?.outlook ?? null, "outlook");
    $("meta_slack").textContent = fmt(deltas?.slack ?? null, "slack");
    $("meta_hubspot").textContent = fmt(deltas?.hubspot ?? null, "hubspot");
    $("meta_monday").textContent = fmt(deltas?.monday ?? null, "monday");

    // TOTAL gauge
    renderTotalGauge(counts.total, deltas?.total ?? null);

    // Wave behavior: always animate; spike on NEW notifications
    for (const k of ["outlook","slack","hubspot","monday"]){
      const prev = Number.isFinite(lastCounts[k]) ? lastCounts[k] : null;
      const cur  = counts[k] ?? 0;

      WAVES[k]?.setValue(cur);

      if (prev != null && cur > prev){
        WAVES[k]?.spike(cur - prev);
      }
      lastCounts[k] = cur;
    }

    updateLights(counts);
    updateBusyMode(counts.total);
    setUpdateTime(nowStr());

    const ms = Math.round(performance.now() - t0);
    $("apiMs").textContent = `${ms}ms`;
    $("pollState").textContent = "LIVE";
  }catch(e){
    console.error(e);
    setLight("check", true);
    setMode("CHECK");
    setUpdateTime(nowStr());
    $("pollState").textContent = "ERROR";
    $("apiMs").textContent = "—";
  }
}

/* =========================
   AUTH
   ========================= */
async function signIn(){
  const loginReq = { scopes: SCOPES };

  // Handle any redirect response safely
  await msalApp.handleRedirectPromise().catch(() => null);

  // Already signed in?
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
  setButtons(false);
  setDriver("—");
  setMode("IDLE");
  setUpdateTime("—");
  setBaselineUI();

  // reset UI
  $("val_total").textContent = "—";
  $("meta_total").textContent = "New since baseline: —";
  for (const k of ["outlook","slack","hubspot","monday"]){
    $(`val_${k}`).textContent = "—";
    $(`meta_${k}`).textContent = "Since baseline: —";
    lastCounts[k] = null;
    WAVES[k]?.setValue(0);
  }
  setLight("outlook", false);
  setLight("slack", false);
  setLight("hubspot", false);
  setLight("monday", false);
  setLight("check", false);
  $("pollState").textContent = "IDLE";
  $("apiMs").textContent = "—";

  if (autoTimer){
    clearInterval(autoTimer);
    autoTimer = null;
    $("btnAuto").textContent = "AUTO: OFF";
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
   INIT
   ========================= */
window.addEventListener("load", () => {
  setAppLinks();
  initWaves();
  requestAnimationFrame(waveLoop);

  const msalConfig = {
    auth: {
      clientId: CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: "sessionStorage" }
  };
  msalApp = new msal.PublicClientApplication(msalConfig);

  $("btnIgnition").onclick = () => signIn().catch(err => { console.error(err); setLight("check", true); });
  $("btnRefresh").onclick  = () => refreshAll().catch(() => setLight("check", true));
  $("btnBaseline").onclick = () => setBaselineFromCurrent();
  $("btnAuto").onclick     = () => toggleAuto();
  $("btnSignOut").onclick  = () => signOut().catch(() => {});

  setButtons(false);
  setDriver("—");
  setUpdateTime("—");
  setBaselineUI();
  setMode("IDLE");
  $("pollState").textContent = "IDLE";
  $("apiMs").textContent = "—";
});
