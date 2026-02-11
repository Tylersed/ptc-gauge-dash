/* ==========================================================
   PTC Gauge Cluster — LIVE TELEMETRY (Static GitHub Pages)
   - TOTAL = Outlook Inbox unread + (Slack/HubSpot/Monday Outlook folder unread)
   - Waves = oscilloscope feed, always scrolling right->left
   - Spikes travel left when new notifications arrive
   - Counts only (Mail.ReadBasic), no message bodies
   ========================================================== */

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

  // Polling interval (data pulls). Waves animate continuously regardless.
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
function clamp(n,a,b){ return Math.max(a, Math.min(b,n)); }
function pct(value, max){ return clamp(value / Math.max(1,max), 0, 1); }

/* =========================
   UI helpers
   ========================= */
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
   Center TOTAL gauge
   ========================= */
let _needle = 0;
function needleAngle(p){
  _needle = _needle + (p - _needle) * 0.35;
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

/* ==========================================================
   OSCILLOSCOPE WAVE ENGINE (always scrolling right->left)
   - Uses a ring buffer of samples
   - New samples enter on the RIGHT
   - Old samples drift LEFT
   - Trails + glow + “years of work” look
   ========================================================== */
class OscilloscopeWave {
  constructor(key, canvasId, cardId){
    this.key = key;
    this.cv = $(canvasId);
    this.card = $(cardId);
    this.ctx = this.cv.getContext("2d", { alpha: true });

    this.dpr = Math.max(1, window.devicePixelRatio || 1);
    this.w = 0;
    this.h = 0;

    // ring buffer
    this.N = 300;                   // resolution of the wave
    this.buf = new Float32Array(this.N);
    this.write = 0;

    // timing
    this.acc = 0;
    this.sampleHz = 90;             // how fast new samples enter (controls scroll speed)
    this.sampleDt = 1 / this.sampleHz;

    // motion
    this.phase = Math.random() * 1000;
    this.phase2 = Math.random() * 1000;

    // level drives intensity (based on unread count)
    this.level = 0;                 // 0..1
    this.pendingSpikes = 0;         // integer spikes queued
    this.spikeTail = 0;             // decay tail for nicer spikes

    // accent from CSS
    this.accent = this._getAccent();

    this._resize();
    window.addEventListener("resize", () => {
      this.dpr = Math.max(1, window.devicePixelRatio || 1);
      this._resize();
    });

    // start with a visible line even before any data pulls
    for (let i=0;i<this.N;i++){
      this.buf[i] = (Math.sin(i*0.08) * 0.08);
    }
  }

  _getAccent(){
    try{
      const v = getComputedStyle(this.card).getPropertyValue("--accent").trim();
      return v || "rgba(96,165,255,1)";
    }catch{
      return "rgba(96,165,255,1)";
    }
  }

  _resize(){
    const rect = this.cv.getBoundingClientRect();
    this.w = Math.max(1, rect.width);
    this.h = Math.max(1, rect.height);

    this.cv.width = Math.floor(this.w * this.dpr);
    this.cv.height = Math.floor(this.h * this.dpr);

    // draw in CSS pixels
    this.ctx.setTransform(this.dpr,0,0,this.dpr,0,0);

    // clear hard once on resize
    this.ctx.clearRect(0,0,this.w,this.h);
  }

  setCount(count){
    const c = Number.isFinite(count) ? count : 0;
    const target = clamp(c / 55, 0, 1);           // tune mapping
    this.level = this.level + (target - this.level) * 0.10;
  }

  spike(delta){
    const d = Math.max(1, Math.min(25, delta|0));
    this.pendingSpikes += d;

    if (this.card){
      this.card.classList.remove("pulse");
      void this.card.offsetWidth;
      this.card.classList.add("pulse");
    }
  }

  _pushSample(v){
    this.buf[this.write] = v;
    this.write = (this.write + 1) % this.N;
  }

  _nextSample(dt){
    // baseline motion (always moving even at 0)
    const baseAmp = 0.10 + this.level * 0.26;
    const f1 = 0.90 + this.level * 0.60;
    const f2 = 0.55 + this.level * 0.40;

    this.phase += dt * (4.6 + this.level * 4.2);
    this.phase2 += dt * (2.8 + this.level * 2.0);

    const s1 = Math.sin(this.phase * f1) * 0.62;
    const s2 = Math.sin(this.phase2 * f2) * 0.38;

    // deterministic-ish noise (no RNG flicker)
    const n = Math.sin((this.phase * 3.7) + (this.phase2 * 1.9)) * 0.14;

    // spikes: sharp attack, slow decay, and they travel left as samples scroll
    let spike = 0;
    if (this.pendingSpikes > 0){
      spike = 1.25 + Math.min(1.2, this.pendingSpikes * 0.02);
      this.pendingSpikes -= 1;
      this.spikeTail = Math.max(this.spikeTail, spike);
    }
    this.spikeTail *= 0.88; // decay
    spike += this.spikeTail * 0.20;

    // final sample in -1..1
    // spike is positive => wave goes UP on screen (we subtract later)
    const v = clamp(((s1+s2+n) * baseAmp) + (spike * 0.75), -1.2, 1.6);
    this._pushSample(v);
  }

  step(dt){
    // update accent in case theme changes
    this.accent = this._getAccent();

    // generate new samples at fixed Hz (controls scroll)
    this.acc += dt;
    while (this.acc >= this.sampleDt){
      this._nextSample(this.sampleDt);
      this.acc -= this.sampleDt;
    }
  }

  render(){
    const ctx = this.ctx;
    const w = this.w;
    const h = this.h;

    // TRAILS: fade old frames slightly (professional oscilloscope look)
    ctx.save();
    ctx.globalCompositeOperation = "source-over";
    ctx.fillStyle = "rgba(0,0,0,0.18)";
    ctx.fillRect(0,0,w,h);
    ctx.restore();

    const mid = h * 0.55;
    const ampPx = (h * 0.33) * (0.55 + this.level * 0.75);

    // draw faint baseline
    ctx.save();
    ctx.strokeStyle = "rgba(255,255,255,0.08)";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(0, mid);
    ctx.lineTo(w, mid);
    ctx.stroke();
    ctx.restore();

    // Wave path: oldest on left, newest on right
    const sampleAt = (i) => this.buf[(this.write + i) % this.N];

    // Glow layers (lighter blend)
    const drawWave = (lineW, alpha, color) => {
      ctx.save();
      ctx.globalCompositeOperation = "lighter";
      ctx.globalAlpha = alpha;
      ctx.lineWidth = lineW;
      ctx.strokeStyle = color;
      ctx.lineJoin = "round";
      ctx.lineCap = "round";
      ctx.beginPath();

      for (let x=0; x<=w; x++){
        const t = x / w;
        const idx = Math.floor(t * (this.N - 1));
        const v = sampleAt(idx);

        // v positive should go UP (smaller y)
        const y = mid - (v * ampPx);

        if (x === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      }
      ctx.stroke();
      ctx.restore();
    };

    // outer bloom
    drawWave(10, 0.10, this.accent);
    drawWave(6,  0.16, this.accent);

    // crisp core
    drawWave(2.2, 0.92, "rgba(255,255,255,0.95)");
    drawWave(1.6, 0.92, this.accent);

    // “spark” dot at the newest point (right edge) for wow
    const newest = this.buf[(this.write + this.N - 1) % this.N];
    const yN = mid - (newest * ampPx);

    ctx.save();
    ctx.globalCompositeOperation = "lighter";
    ctx.globalAlpha = 0.85;
    ctx.fillStyle = this.accent;
    ctx.beginPath();
    ctx.arc(w - 3, yN, 3.2, 0, Math.PI*2);
    ctx.fill();

    ctx.globalAlpha = 0.25;
    ctx.beginPath();
    ctx.arc(w - 3, yN, 10, 0, Math.PI*2);
    ctx.fill();
    ctx.restore();
  }
}

/* =========================
   Waves init + animation loop
   ========================= */
const WAVES = {};
let _lastFrame = performance.now();

function initWaves(){
  WAVES.outlook = new OscilloscopeWave("outlook", "cv_outlook", "w_outlook");
  WAVES.slack   = new OscilloscopeWave("slack",   "cv_slack",   "w_slack");
  WAVES.hubspot = new OscilloscopeWave("hubspot", "cv_hubspot", "w_hubspot");
  WAVES.monday  = new OscilloscopeWave("monday",  "cv_monday",  "w_monday");
}

function animate(){
  const now = performance.now();
  const dt = Math.min(0.05, (now - _lastFrame) / 1000);
  _lastFrame = now;

  for (const k of ["outlook","slack","hubspot","monday"]){
    const w = WAVES[k];
    if (!w) continue;
    w.step(dt);
    w.render();
  }
  requestAnimationFrame(animate);
}

/* =========================
   Counts + deltas
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

    // numbers
    $("val_outlook").textContent = String(counts.outlook);
    $("val_slack").textContent = String(counts.slack);
    $("val_hubspot").textContent = String(counts.hubspot);
    $("val_monday").textContent = String(counts.monday);

    // deltas from baseline
    const { deltas } = computeDeltas(counts);
    const fmt = (v, key) => {
      if (v == null) return key === "total" ? "New since baseline: —" : "Since baseline: —";
      const sign = v > 0 ? "+" : "";
      return key === "total" ? `New since baseline: ${sign}${v}` : `Since baseline: ${sign}${v}`;
    };
    $("meta_outlook").textContent = fmt(deltas?.outlook ?? null, "outlook");
    $("meta_slack").textContent   = fmt(deltas?.slack ?? null, "slack");
    $("meta_hubspot").textContent = fmt(deltas?.hubspot ?? null, "hubspot");
    $("meta_monday").textContent  = fmt(deltas?.monday ?? null, "monday");

    // total gauge
    renderTotalGauge(counts.total, deltas?.total ?? null);

    // wave intensity + spikes that travel left
    for (const k of ["outlook","slack","hubspot","monday"]){
      const prev = Number.isFinite(lastCounts[k]) ? lastCounts[k] : null;
      const cur  = counts[k] ?? 0;

      WAVES[k]?.setCount(cur);

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

  setButtons(false);
  setDriver("—");
  setMode("IDLE");
  setUpdateTime("—");
  setBaselineUI();

  $("val_total").textContent = "—";
  $("meta_total").textContent = "New since baseline: —";

  for (const k of ["outlook","slack","hubspot","monday"]){
    $(`val_${k}`).textContent = "—";
    $(`meta_${k}`).textContent = "Since baseline: —";
    lastCounts[k] = null;
    WAVES[k]?.setCount(0);
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
   Init
   ========================= */
window.addEventListener("load", () => {
  setAppLinks();
  initWaves();
  requestAnimationFrame(animate);

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
