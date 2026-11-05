/* ==========================================================================
   GL Trial Balance Generator - app.js   (ES module)

   Needs, in the same folder:  index.html, styles.css, data.js, excel-tools.js

   Data flow
     - Strict authentication: no guest mode. Every sign-in verifies both
       ALLOWED_EMAILS (if set) and real Firestore read permission BEFORE any
       local browser data is loaded or displayed.
     - Local data is stored per user (localStorage key scoped by e-mail).
     - Nothing goes to Firebase until the user presses "Save to Cloud".
     - Save to Cloud pushes ONLY the GLs edited on this browser since the last save.

   Performance
     - ExcelJS / JSZip / jsPDF are loaded on demand (and pre-fetched when idle),
       so the first page load only downloads what the login screen needs.
     - All UI updates go through one requestAnimationFrame scheduler, so many
       edits in a row cause a single repaint, and hidden views are not rendered.
     - Per-GL comparison and validation results are cached and invalidated
       only when that GL changes.
   ========================================================================== */

import {
  numberToWords, parseAmount, toCents, formatMoney,
  normalizeDate, isValidDate, parsePastedRows, parseWorkbookFile,
  buildWorkbookBuffer, buildAndSavePdf,
  plainTextFromHtml, sanitizeRichNoteHtml
} from "./excel-tools.js?v=20261105_5";

const CFG = window.APP_CONFIG;
const MASTER = window.MASTER_GL_LIST.slice().sort(function (a, b) { return a.glNo.localeCompare(b.glNo); });
const GL_BY_NO = window.GL_BY_NO;
const KEYS = CFG.storage;
const FB_VERSION = "10.12.2";

const $ = function (id) { return document.getElementById(id); };
const esc = function (s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
  });
};
const plural = function (n, word) { return n + " " + word + (n === 1 ? "" : "s"); };
const pad2 = function (n) { return String(n).padStart(2, "0"); };
function debounce(fn, ms) {
  let t = 0;
  return function () { const a = arguments, self = this; clearTimeout(t); t = setTimeout(function () { fn.apply(self, a); }, ms); };
}


/* ==========================================================================
   0. ON-DEMAND LIBRARIES
   ========================================================================== */
const LIBS = {
  exceljs:   "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js",
  jszip:     "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js",
  jspdf:     "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js",
  autotable: "https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"
};
const libPromises = {};

function loadLib(name) {
  if (libPromises[name]) return libPromises[name];
  libPromises[name] = new Promise(function (resolve, reject) {
    const s = document.createElement("script");
    s.src = LIBS[name];
    s.async = true;
    s.crossOrigin = "anonymous";
    s.onload = function () { resolve(); };
    s.onerror = function () {
      delete libPromises[name];
      s.remove();
      reject(new Error("Could not load the " + name + " library. Check your internet connection."));
    };
    document.head.appendChild(s);
  });
  return libPromises[name];
}
function ensureExcelLibs() { return Promise.all([loadLib("exceljs"), loadLib("jszip")]); }
function ensurePdfLibs() { return loadLib("jspdf").then(function () { return loadLib("autotable"); }); }

/** Warm the cache in the background once the user is inside the app. */
function prefetchLibsWhenIdle() {
  const idle = window.requestIdleCallback || function (cb) { return setTimeout(cb, 2500); };
  idle(function () {
    ensureExcelLibs().catch(function () {});
    idle(function () { ensurePdfLibs().catch(function () {}); });
  });
}


/* ==========================================================================
   1. STATE + LOCAL STORAGE + 6-MONTH DATABASE HISTORY
   state.gls[glNo] = {
     mode, items:[{date,ref,desc,amount}], override, glBalance, diffNote,
     targetMonthBalance,
     months: {
       "YYYY.MM.DD": { mode, items, glBalance, tbBalance, override, diffNote }
     }
   }
   ========================================================================== */
const MAX_SAVED_MONTHS = Number(CFG.maxSavedMonths) || 6;
const META_DOC_ID = (CFG.firestore && CFG.firestore.metaDocId) || "_meta";

function emptyState() {
  return {
    officerName: CFG.defaults.officerName,
    serviceNo: CFG.defaults.serviceNo,
    tbDate: CFG.defaults.tbDate,
    prevMonthDate: "",
    savedMonths: [],
    selected: {},
    activeGl: null,
    gls: {},
    dirtyGls: {},
    metaDirty: false
  };
}

let state = emptyState();
const ui = {
  view: "schedules",
  filter: "all",          // all | selected | needs_update | issues | zero | balance
  search: "",
  balScope: "selected",
  balFilter: "all",       // all | changed | unchanged | pending
  balSearch: ""
};
let persistTimer = 0;
let isAuthorizedSession = false;
let fb = null;
let currentUser = null;

function storageKeyForUser(baseKey, email) {
  const clean = String(email || "").trim().toLowerCase();
  return clean ? baseKey + "." + clean : baseKey;
}

function cloneItems(items) {
  return (Array.isArray(items) ? items : []).map(function (i) {
    return {
      date: String(i && i.date != null ? i.date : ""),
      ref: String(i && i.ref != null ? i.ref : ""),
      desc: String(i && i.desc != null ? i.desc : ""),
      amount: String(i && i.amount != null ? i.amount : "")
    };
  });
}

function hydrate(saved) {
  const s = Object.assign(emptyState(), saved || {}, {
    selected: (saved && saved.selected) || {},
    gls: (saved && saved.gls) || {},
    dirtyGls: (saved && saved.dirtyGls) || {},
    savedMonths: Array.isArray(saved && saved.savedMonths) ? saved.savedMonths : []
  });
  // Ensure every GL has a valid `months` map and migrate any legacy prevMonthDate/prevMonthBalance into `months`
  Object.keys(s.gls).forEach(function (no) {
    s.gls[no] = normalizeLocalGl(s.gls[no], s.tbDate);
  });
  return s;
}

function normalizeLocalGl(raw, activeTbDate) {
  const g = Object.assign(defaultGl(), raw || {});
  g.items = cloneItems(g.items);
  g.months = (raw && raw.months && typeof raw.months === "object") ? Object.assign({}, raw.months) : {};

  // Migrate legacy prevMonthDate + prevMonthBalance if present and not yet in g.months
  const legPrevDate = normalizeDate(String(g.prevMonthDate || "").trim());
  const legPrevBal = String(g.prevMonthBalance == null ? "" : g.prevMonthBalance).trim();
  if (isValidDate(legPrevDate) && legPrevBal !== "" && !g.months[legPrevDate]) {
    g.months[legPrevDate] = {
      mode: parseAmount(legPrevBal) !== 0 ? "balance" : "zero",
      items: legPrevDate !== activeTbDate ? cloneItems(g.items) : [],
      glBalance: legPrevBal,
      tbBalance: legPrevBal,
      override: false,
      diffNote: ""
    };
  }
  return g;
}

function loadStateForUser(email) {
  const lower = String(email || "").toLowerCase();
  try {
    sessionStorage.removeItem("glTb.demo");
    const userKey = storageKeyForUser(KEYS.state, lower);
    let raw = localStorage.getItem(userKey);

    const ownerKey = KEYS.state + ".owner";
    const legacy = localStorage.getItem(KEYS.state);
    if (legacy) {
      const owner = localStorage.getItem(ownerKey);
      if (!raw && (!owner || owner === lower)) raw = legacy;
      if (!owner || owner === lower) {
        localStorage.removeItem(KEYS.state);
        localStorage.removeItem(ownerKey);
      }
    }
    if (raw) {
      const st = hydrate(JSON.parse(raw));
      pruneToSixMonths(st);
      return st;
    }
  } catch (e) { console.warn("Could not read saved state", e); }
  return emptyState();
}

function persistNow() {
  clearTimeout(persistTimer);
  persistTimer = 0;
  if (!isAuthorizedSession || !currentUser) return;
  try {
    localStorage.setItem(storageKeyForUser(KEYS.state, currentUser.email), JSON.stringify(state));
    localStorage.setItem(KEYS.dirty, hasUnsynced() ? "true" : "false");
  } catch (e) {
    toast("Browser storage is full or blocked. Your latest edits may not be kept.", "error");
  }
}
function persistSoon() { clearTimeout(persistTimer); persistTimer = setTimeout(persistNow, 300); }

function hasUnsynced() {
  if (state.metaDirty) return true;
  for (const k in state.dirtyGls) { if (state.dirtyGls[k]) return true; }
  return false;
}

function defaultGl() {
  return {
    mode: "zero",
    items: [],
    override: false,
    glBalance: "",
    diffNote: "",
    prevMonthDate: "",
    prevMonthBalance: "",
    targetMonthBalance: "",
    months: {}
  };
}
function readGl(glNo) {
  if (!state.gls[glNo]) return defaultGl();
  const g = Object.assign(defaultGl(), state.gls[glNo]);
  g.months = state.gls[glNo].months || {};
  return g;
}
function ensureGl(glNo) {
  if (!state.gls[glNo]) state.gls[glNo] = defaultGl();
  else state.gls[glNo] = Object.assign(defaultGl(), state.gls[glNo], { months: state.gls[glNo].months || {} });
  return state.gls[glNo];
}

/** Call after changing one GL's data. */
function markDirty(glNo) {
  if (isValidDate(state.tbDate)) {
    saveGlToMonthHistory(glNo, state.tbDate);
  }
  state.dirtyGls[glNo] = true;
  invalidate(glNo);
  persistSoon();
  schedule("badge", "counts");
}

/* ---- Month-End Calendar & 6-Month Database History Helpers ---- */

/** Given "YYYY.MM.DD", returns the exact last day of the previous calendar month as "YYYY.MM.DD". */
function prevMonthEndOf(dateStr) {
  if (!isValidDate(dateStr)) return "";
  const p = dateStr.split(".");
  const y = +p[0], m = +p[1];
  const prevLastDay = new Date(y, m - 1, 0);
  return prevLastDay.getFullYear() + "." + pad2(prevLastDay.getMonth() + 1) + "." + pad2(prevLastDay.getDate());
}

function lastMonthEnd() {
  const now = new Date();
  const d = new Date(now.getFullYear(), now.getMonth(), 0);
  return d.getFullYear() + "." + pad2(d.getMonth() + 1) + "." + pad2(d.getDate());
}

/** Returns the last `count` month-end dates (ending with this month-end and preceding months) in descending order. */
function recentMonthEnds(count) {
  count = count || 6;
  const now = new Date();
  const list = [];
  for (let i = 0; i < count; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + 1 - i, 0);
    list.push(d.getFullYear() + "." + pad2(d.getMonth() + 1) + "." + pad2(d.getDate()));
  }
  return list;
}

/** Collects all valid month-end dates saved in `st`, sorted ascending (oldest -> newest). */
function getSavedMonthsList(st) {
  st = st || state;
  const set = new Set();
  (st.savedMonths || []).forEach(function (m) { if (isValidDate(m)) set.add(m); });
  Object.keys(st.gls || {}).forEach(function (no) {
    const months = st.gls[no] && st.gls[no].months;
    if (months && typeof months === "object") {
      Object.keys(months).forEach(function (m) { if (isValidDate(m)) set.add(m); });
    }
  });
  return Array.from(set).sort();
}

/** Prunes saved month-end reports so at most `MAX_SAVED_MONTHS` (6) months are kept in the database. */
function pruneToSixMonths(st) {
  st = st || state;
  const all = getSavedMonthsList(st);
  if (all.length <= MAX_SAVED_MONTHS) {
    st.savedMonths = all;
    return [];
  }
  const keepList = all.slice(all.length - MAX_SAVED_MONTHS);
  const keepSet = new Set(keepList);
  const removed = all.slice(0, all.length - MAX_SAVED_MONTHS);
  st.savedMonths = keepList;
  Object.keys(st.gls || {}).forEach(function (no) {
    const mObj = st.gls[no] && st.gls[no].months;
    if (!mObj) return;
    Object.keys(mObj).forEach(function (m) {
      if (!keepSet.has(m)) delete mObj[m];
    });
  });
  return removed;
}

/** Saves one GL's current working fields into `g.months[monthDate]`. */
function saveGlToMonthHistory(glNo, monthDate) {
  if (!isValidDate(monthDate)) return;
  const g = ensureGl(glNo);
  const c = calcGl(g);
  const sysRaw = getSystemGlRaw(g);
  const hasAnyData = sysRaw !== "" || cleanItems(g).length > 0 || g.mode === "balance" || (g.months && g.months[monthDate]);
  if (!hasAnyData) return;

  g.months[monthDate] = {
    mode: g.mode,
    items: cleanItems(g).map(function (it) { return { date: it.date, ref: it.ref, desc: it.desc, amount: it.amount }; }),
    glBalance: sysRaw,
    tbBalance: String(c.tb / 100),
    override: !!g.override,
    diffNote: String(g.diffNote || "")
  };
  if (state.savedMonths.indexOf(monthDate) === -1) {
    state.savedMonths.push(monthDate);
  }
  pruneToSixMonths(state);
}

/** Looks up the Previous Month-End record strictly from the database (`g.months`). Never manually typed! */
function getPrevMonthFromDb(glNo) {
  const curDate = state.tbDate;
  if (!isValidDate(curDate)) return { hasPrev: false, prevDate: "", prevCents: 0, prevSnap: null };

  const expectedPrev = prevMonthEndOf(curDate); // e.g. "2026.09.30" for "2026.10.31"
  const expectedPrefix = expectedPrev.slice(0, 7); // "2026.09"
  const g = readGl(glNo);
  const months = g.months || {};

  let matchedDate = "";
  if (months[expectedPrev]) {
    matchedDate = expectedPrev;
  } else {
    // Check if there is a date in the exact previous calendar month (YYYY.MM.*)
    const keys = Object.keys(months).filter(function (k) { return k.slice(0, 7) === expectedPrefix; }).sort();
    if (keys.length) matchedDate = keys[keys.length - 1];
  }

  if (!matchedDate || !months[matchedDate]) {
    return { hasPrev: false, prevDate: "", prevCents: 0, prevSnap: null };
  }

  const snap = months[matchedDate];
  const rawBal = String(snap.glBalance != null && String(snap.glBalance).trim() !== "" ? snap.glBalance : (snap.tbBalance || "0")).trim();
  const num = parseAmount(rawBal);
  const cents = isNaN(num) ? 0 : toCents(num);
  return {
    hasPrev: true,
    prevDate: matchedDate,
    prevCents: cents,
    prevSnap: snap
  };
}

/**
 * Switches the active Current Month-End Date (`state.tbDate`).
 * Loads working data for `newMonthDate` if it exists in `g.months`; otherwise starts fresh for `newMonthDate`
 * and carries over previous month's schedule items (if previous month exists in DB) for reuse.
 */
function activateMonthEnd(newMonthDate, markAllDirty) {
  const cleanDate = normalizeDate(String(newMonthDate || "").trim());
  if (!isValidDate(cleanDate)) return false;

  const oldDate = state.tbDate;
  if (isValidDate(oldDate) && oldDate !== cleanDate) {
    MASTER.forEach(function (gl) { saveGlToMonthHistory(gl.glNo, oldDate); });
  }

  state.tbDate = cleanDate;
  state.prevMonthDate = prevMonthEndOf(cleanDate);
  state.metaDirty = true;

  MASTER.forEach(function (gl) {
    const no = gl.glNo;
    const g = ensureGl(no);
    const curSnap = g.months && g.months[cleanDate];
    const prevInfo = getPrevMonthFromDb(no);

    if (curSnap) {
      g.mode = curSnap.mode === "balance" ? "balance" : "zero";
      g.items = cloneItems(curSnap.items);
      g.targetMonthBalance = String(curSnap.glBalance == null ? "" : curSnap.glBalance);
      g.glBalance = g.targetMonthBalance;
      g.override = !!curSnap.override;
      g.diffNote = String(curSnap.diffNote == null ? "" : curSnap.diffNote);
    } else if (oldDate && oldDate !== cleanDate) {
      // Starting a new month-end: Current Month GL Balance must be entered manually in GL Balances tab
      g.targetMonthBalance = "";
      g.glBalance = "";
      g.override = false;
      g.diffNote = "";
      if (prevInfo.hasPrev && prevInfo.prevSnap) {
        g.items = cloneItems(prevInfo.prevSnap.items);
        g.mode = prevInfo.prevSnap.mode === "balance" || g.items.length > 0 ? "balance" : "zero";
      } else {
        g.items = [];
        g.mode = "zero";
      }
    }
    g.prevMonthDate = prevInfo.hasPrev ? prevInfo.prevDate : "";
    g.prevMonthBalance = prevInfo.hasPrev ? String(prevInfo.prevCents / 100) : "";
    if (markAllDirty) state.dirtyGls[no] = true;
  });

  pruneToSixMonths(state);
  invalidate();
  persistSoon();
  return true;
}

function normalizeCloudGl(d, activeMonth) {
  const months = {};
  if (d.months && typeof d.months === "object") {
    Object.keys(d.months).forEach(function (mKey) {
      const m = normalizeDate(mKey);
      const snap = d.months[mKey];
      if (!isValidDate(m) || !snap || typeof snap !== "object") return;
      months[m] = {
        mode: snap.mode === "balance" ? "balance" : "zero",
        items: cloneItems(snap.items),
        glBalance: String(snap.glBalance == null ? "" : snap.glBalance),
        tbBalance: String(snap.tbBalance == null ? "" : snap.tbBalance),
        override: !!snap.override,
        diffNote: String(snap.diffNote == null ? "" : snap.diffNote)
      };
    });
  }

  const topItems = cloneItems(d.items);
  const topMode = d.mode === "balance" ? "balance" : "zero";
  const topGlBal = String(d.targetMonthBalance || d.glBalance || "").trim();
  const docTbDate = normalizeDate(String(d.tbDate || ""));
  const docPrevDate = normalizeDate(String(d.prevMonthDate || ""));
  const docPrevBal = String(d.prevMonthBalance == null ? "" : d.prevMonthBalance).trim();

  // Migrate legacy previous month snapshot if present on the cloud doc
  if (isValidDate(docPrevDate) && docPrevBal !== "" && !months[docPrevDate]) {
    months[docPrevDate] = {
      mode: parseAmount(docPrevBal) !== 0 ? "balance" : "zero",
      items: docPrevDate !== docTbDate ? topItems : [],
      glBalance: docPrevBal,
      tbBalance: docPrevBal,
      override: false,
      diffNote: ""
    };
  }
  // Migrate top-level month snapshot if present on the cloud doc
  if (isValidDate(docTbDate) && !months[docTbDate]) {
    months[docTbDate] = {
      mode: topMode,
      items: topItems,
      glBalance: topGlBal,
      tbBalance: topGlBal,
      override: !!d.override,
      diffNote: String(d.diffNote == null ? "" : d.diffNote)
    };
  }

  const effMonth = isValidDate(activeMonth) ? activeMonth : docTbDate;
  const activeSnap = isValidDate(effMonth) ? months[effMonth] : null;

  return {
    mode: activeSnap ? activeSnap.mode : topMode,
    items: activeSnap ? cloneItems(activeSnap.items) : topItems,
    override: activeSnap ? !!activeSnap.override : !!d.override,
    glBalance: activeSnap ? String(activeSnap.glBalance || "") : topGlBal,
    targetMonthBalance: activeSnap ? String(activeSnap.glBalance || "") : topGlBal,
    diffNote: activeSnap ? String(activeSnap.diffNote || "") : String(d.diffNote == null ? "" : d.diffNote),
    prevMonthDate: docPrevDate,
    prevMonthBalance: docPrevBal,
    months: months
  };
}


/* ==========================================================================
   2. CALCULATIONS, COMPARISON & VALIDATION (cached)
   ========================================================================== */
function getSystemGlRaw(g) {
  const t = String(g.targetMonthBalance == null ? "" : g.targetMonthBalance).trim();
  if (t !== "") return t;
  return String(g.glBalance == null ? "" : g.glBalance).trim();
}

function calcGl(g) {
  if (g.mode !== "balance") {
    return { tb: 0, gl: 0, diff: 0, hasSysGl: false, sysRaw: "", hasDiscrepancy: false };
  }
  let tb = 0;
  const items = g.items || [];
  for (let i = 0; i < items.length; i++) {
    const n = parseAmount(items[i].amount);
    if (!isNaN(n)) tb += toCents(n);
  }
  const sysRaw = getSystemGlRaw(g);
  const sysNum = parseAmount(sysRaw);
  const hasSysGl = sysRaw !== "" && !isNaN(sysNum);
  // System GL Balance comes strictly from the Current Month GL Balance entered in GL Balances tab (fixed from system).
  const gl = hasSysGl ? toCents(sysNum) : (g.override ? 0 : tb);
  const diff = gl - tb;
  const note = plainTextFromHtml(g.diffNote || "");
  const hasDiscrepancy = diff !== 0 || (!!g.override && note !== "");
  return { tb: tb, gl: gl, diff: diff, hasSysGl: hasSysGl, sysRaw: sysRaw, hasDiscrepancy: hasDiscrepancy };
}

function differenceText(g, c) {
  if (g.mode !== "balance") return "NIL";
  if (c.diff !== 0) {
    return numberToWords(Math.abs(c.diff) / 100);
  }
  return "Nil";
}

function cleanItems(g) {
  return (g.items || []).filter(function (it) { return it.date || it.ref || it.desc || it.amount; });
}

const cmpCache = new Map();
const issueCache = new Map();
function invalidate(glNo) {
  if (glNo) { cmpCache.delete(glNo); issueCache.delete(glNo); }
  else { cmpCache.clear(); issueCache.clear(); }
}

/**
 * Compares Previous Month-End Balance (strictly from the database) vs.
 * Current Month-End GL Balance (manually entered in the GL Balances tab).
 */
function getGlComparison(glNo) {
  const hit = cmpCache.get(glNo);
  if (hit) return hit;

  const g = readGl(glNo);
  const c = calcGl(g);
  const prevInfo = getPrevMonthFromDb(glNo);

  const hasPrev = prevInfo.hasPrev;
  const prevDate = prevInfo.prevDate;
  const prevCents = prevInfo.prevCents;

  const targetRaw = getSystemGlRaw(g);
  const targetNum = parseAmount(targetRaw);
  const hasTarget = targetRaw !== "" && !isNaN(targetNum);
  const targetCents = hasTarget ? toCents(targetNum) : 0;

  const diffFromPrev = (hasTarget && hasPrev) ? (targetCents - prevCents) : 0;
  const scheduleCents = g.mode === "balance" ? c.tb : 0;

  let status = "pending";
  if (hasTarget) {
    if (hasPrev) {
      status = diffFromPrev === 0 ? "unchanged" : "changed";
    } else {
      // No previous month data in DB: if non-zero balance entered, officer needs to verify/update TB schedule
      status = targetCents === 0 ? "unchanged" : "changed";
    }
  }

  const res = {
    hasPrev: hasPrev,
    prevDate: prevDate,
    prevCents: prevCents,
    hasTarget: hasTarget,
    targetRaw: targetRaw,
    targetCents: targetCents,
    diffFromPrev: diffFromPrev,
    scheduleCents: scheduleCents,
    glCents: c.gl,
    tbCents: c.tb,
    diffFromSchedule: hasTarget ? targetCents - scheduleCents : 0,
    isScheduleTallied: hasTarget && targetCents === scheduleCents,
    glMatchesTarget: hasTarget && targetCents === c.gl,
    status: status,
    needsHighlight: status === "changed"
  };
  cmpCache.set(glNo, res);
  return res;
}

/** Blocking problems for one GL (empty array = ready to export). */
function glIssues(glNo) {
  const hit = issueCache.get(glNo);
  if (hit) return hit;
  const g = readGl(glNo);
  const p = [];
  if (g.mode === "balance") {
    const items = cleanItems(g);
    const c = calcGl(g);
    const note = plainTextFromHtml(g.diffNote || "");
    if (!items.length && (!c.hasSysGl || c.gl === 0)) {
      p.push("Marked Has Balance but has no schedule items and no System GL Balance.");
    }
    if (items.some(function (it) { return isNaN(parseAmount(it.amount)); })) {
      p.push("A schedule row has a missing or invalid amount.");
    }
    if (c.sysRaw !== "" && isNaN(parseAmount(c.sysRaw))) {
      p.push("System General Ledger Balance is not a valid number.");
    }
    if ((g.override || note !== "") && !c.hasSysGl) {
      p.push("Difference explanation entered, but Current Month GL Balance is missing. Enter it in the GL Balances tab.");
    } else if ((g.override || note !== "") && c.diff === 0) {
      p.push("System GL Balance and Trial Balance are equal (" + formatMoney(c.gl) + "), so a difference cannot exist. Clear the difference note or update the balance in GL Balances.");
    }
  }
  issueCache.set(glNo, p);
  return p;
}

/** Non-blocking warnings for one GL. */
function glWarnings(glNo) {
  const g = readGl(glNo);
  const c = calcGl(g);
  const cmp = getGlComparison(glNo);
  const w = [];
  if (g.mode === "balance" && c.diff !== 0) {
    w.push("Discrepancy of " + numberToWords(Math.abs(c.diff) / 100) + " (" + formatMoney(Math.abs(c.diff)) + ") between System GL (" + formatMoney(c.gl) + ") and Trial Balance (" + formatMoney(c.tb) + "). The certification clause will be omitted.");
  } else if (cmp.hasTarget && !cmp.glMatchesTarget) {
    w.push("GL balance in the schedule (" + formatMoney(cmp.glCents) + ") is not equal to the entered current month-end balance (" + formatMoney(cmp.targetCents) + ").");
  }
  return w;
}


/* ==========================================================================
   3. RENDER SCHEDULER (one repaint per frame, hidden views are skipped)
   ========================================================================== */
const pending = { presets: false, list: false, editor: false, totals: false, balances: false, counts: false, badge: false };
let rafId = 0;

function schedule() {
  for (let i = 0; i < arguments.length; i++) pending[arguments[i]] = true;
  if (!rafId) rafId = requestAnimationFrame(flush);
}

function flush() {
  rafId = 0;
  const p = Object.assign({}, pending);
  for (const k in pending) pending[k] = false;
  if (!isAuthorizedSession) return;

  if (p.presets) renderPresetButtons();
  if (ui.view === "schedules") {
    if (p.list) renderGlList();
    if (p.editor) renderEditor();
    else if (p.totals) updateTotals();
  } else if (p.balances || p.list) {
    renderGlBalancesView();
  }
  if (p.counts || p.list || p.balances) renderCounts();
  if (p.badge) updateBadge();
}

function renderAll() { schedule("presets", "list", "editor", "balances", "counts", "badge"); }


/* ==========================================================================
   4. TOASTS + DIALOGS
   ========================================================================== */
function toast(message, type, action) {
  type = type || "info";
  const icons = { info: "fa-circle-info", success: "fa-circle-check", error: "fa-circle-xmark", warn: "fa-triangle-exclamation" };
  const el = document.createElement("div");
  el.className = "toast" + (type !== "info" ? " toast-" + type : "");
  el.innerHTML = '<i class="fa-solid ' + icons[type] + ' mt-0.5"></i><span>' + esc(message) + "</span>";
  function dismiss() {
    el.classList.add("is-leaving");
    setTimeout(function () { el.remove(); }, 250);
  }
  if (action) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "toast-action";
    b.textContent = action.label;
    b.addEventListener("click", function () { action.fn(); dismiss(); });
    el.appendChild(b);
  }
  $("toastContainer").appendChild(el);
  setTimeout(dismiss, action ? 7000 : (type === "error" ? 6000 : 3500));
}

function openModal(el) { el.classList.remove("hidden"); }
function closeModal(el) { el.classList.add("hidden"); }
function isOpen(id) { return !$(id).classList.contains("hidden"); }

function confirmDialog(opts) {
  return new Promise(function (resolve) {
    $("confirmTitle").textContent = opts.title;
    $("confirmMessage").textContent = opts.message;
    const ok = $("confirmOk"), cancel = $("confirmCancel");
    ok.textContent = opts.okText || "Confirm";
    ok.className = "btn justify-center " + (opts.danger === false ? "btn-primary" : "btn-danger");
    openModal($("confirmModal"));
    cancel.focus();
    function done(v) {
      closeModal($("confirmModal"));
      ok.onclick = null; cancel.onclick = null;
      resolve(v);
    }
    ok.onclick = function () { done(true); };
    cancel.onclick = function () { done(false); };
  });
}

function setBusy(btn, busy, busyLabel) {
  if (!btn) return;
  const label = btn.querySelector("span");
  if (busy) {
    btn.disabled = true;
    btn.dataset.label = label ? label.textContent : "";
    if (label && busyLabel) label.textContent = busyLabel;
  } else {
    btn.disabled = false;
    if (label && btn.dataset.label) label.textContent = btn.dataset.label;
  }
}

function openSheet(mode) {
  $("actionSheet").dataset.mode = mode;
  $("sheetUserEmail").textContent = currentUser ? (currentUser.email || "-") : "-";
  openModal($("actionSheet"));
}


/* ==========================================================================
   5. FIREBASE + STRICT AUTHORISATION
   ========================================================================== */
async function initFirebase() {
  sessionStorage.removeItem("glTb.demo");
  if (!window.isFirebaseConfigured()) {
    $("firebaseWarning").classList.remove("hidden");
    return;
  }
  try {
    const base = "https://www.gstatic.com/firebasejs/" + FB_VERSION + "/";
    const mods = await Promise.all([
      import(base + "firebase-app.js"),
      import(base + "firebase-auth.js"),
      import(base + "firebase-firestore.js")
    ]);
    const app = mods[0].initializeApp(window.FIREBASE_CONFIG);
    fb = { authM: mods[1], fs: mods[2], auth: mods[1].getAuth(app), db: mods[2].getFirestore(app) };
    fb.authM.onAuthStateChanged(fb.auth, onAuthChanged);
  } catch (e) {
    console.error(e);
    showLoginError("Could not load Firebase. Check your internet connection and reload.");
  }
}

function showLoginError(msg) {
  const el = $("loginError");
  el.textContent = msg;
  el.classList.remove("hidden");
}

/** True only if the user is allowed AND Firestore rules let them read the collection. */
async function verifyDatabaseAccess(user) {
  if (!fb || !user) return false;
  const allowed = window.ALLOWED_EMAILS || [];
  const email = (user.email || "").toLowerCase();
  if (allowed.length && allowed.map(function (e) { return e.toLowerCase(); }).indexOf(email) === -1) return false;
  try {
    const fs = fb.fs;
    const q = fs.query(fs.collection(fb.db, CFG.firestore.collection), fs.limit(1));
    await (typeof fs.getDocsFromServer === "function" ? fs.getDocsFromServer(q) : fs.getDocs(q));
    return true;
  } catch (err) {
    console.warn("Database authorisation check failed for", email, err);
    return false;
  }
}

/** Saves the current user's Officer Name & Service Number to Firestore `_meta.officers[email]` */
async function saveOfficerProfileToCloud() {
  if (!isAuthorizedSession || !fb || !currentUser || !navigator.onLine) return;
  const emailKey = String(currentUser.email || "").trim().toLowerCase();
  if (!emailKey) return;
  try {
    const fs = fb.fs;
    const officersUpdate = {};
    officersUpdate[emailKey] = {
      officerName: String(state.officerName || "").trim(),
      serviceNo: String(state.serviceNo || "").trim(),
      email: emailKey,
      updatedAt: fs.serverTimestamp()
    };
    await fs.setDoc(fs.doc(fb.db, CFG.firestore.collection, META_DOC_ID), {
      officers: officersUpdate
    }, { merge: true });
  } catch (e) {
    console.warn("Could not auto-save officer profile to cloud:", e);
  }
}
const saveOfficerProfileSoon = debounce(saveOfficerProfileToCloud, 800);

/** Fetches active month-end metadata & per-Gmail Officer profile (`_meta`) from Firebase on sign-in. */
async function syncFirebaseMeta() {
  if (!fb) return;
  try {
    const fs = fb.fs;
    const getDocFn = typeof fs.getDocFromServer === "function" ? fs.getDocFromServer : fs.getDoc;
    const metaRef = fs.doc(fb.db, CFG.firestore.collection, META_DOC_ID);
    const metaSnap = await getDocFn(metaRef).catch(function () { return fs.getDoc(metaRef); });
    const emailKey = String((currentUser && currentUser.email) || "").trim().toLowerCase();
    if (metaSnap && metaSnap.exists()) {
      const m = metaSnap.data() || {};
      if (m.currentMonthEnd && isValidDate(normalizeDate(m.currentMonthEnd))) {
        const cm = normalizeDate(m.currentMonthEnd);
        if (state.tbDate !== cm) {
          activateMonthEnd(cm, false);
          state.metaDirty = false;
        }
      }
      if (Array.isArray(m.savedMonths)) {
        m.savedMonths.forEach(function (x) {
          const nd = normalizeDate(x);
          if (isValidDate(nd) && state.savedMonths.indexOf(nd) === -1) state.savedMonths.push(nd);
        });
        pruneToSixMonths(state);
      }
      // Load Officer Name & Service Number saved for this specific Gmail account
      const officersMap = (m.officers && typeof m.officers === "object") ? m.officers : {};
      const cloudOfficer = emailKey ? officersMap[emailKey] : null;
      if (cloudOfficer && typeof cloudOfficer === "object") {
        if (typeof cloudOfficer.officerName === "string") state.officerName = cloudOfficer.officerName;
        if (typeof cloudOfficer.serviceNo === "string") state.serviceNo = cloudOfficer.serviceNo;
        persistNow();
      } else if (emailKey && (state.officerName || state.serviceNo)) {
        saveOfficerProfileToCloud();
      }
    } else if (emailKey && (state.officerName || state.serviceNo)) {
      saveOfficerProfileToCloud();
    }
  } catch (e) { /* ignore if _meta does not exist yet */ }
}

async function onAuthChanged(user) {
  if (!user) {
    currentUser = null;
    isAuthorizedSession = false;
    state = emptyState();
    invalidate();
    leaveApp();
    return;
  }

  const btn = $("btnGoogleSignIn");
  setBusy(btn, true, "Verifying access...");
  const ok = await verifyDatabaseAccess(user);
  if (!ok) {
    setBusy(btn, false);
    currentUser = null;
    isAuthorizedSession = false;
    state = emptyState();
    invalidate();
    leaveApp();
    showLoginError("Access Denied: " + (user.email || "This Google account") + " is not authorised to access the GL database. Local data is protected.");
    try { await fb.authM.signOut(fb.auth); } catch (e) { /* ignore */ }
    return;
  }

  $("loginError").classList.add("hidden");
  currentUser = user;
  isAuthorizedSession = true;
  state = loadStateForUser(user.email);
  await syncFirebaseMeta();
  setBusy(btn, false);
  invalidate();
  enterApp(user.email || user.displayName || "Signed in");
}

async function signIn() {
  $("loginError").classList.add("hidden");
  if (!fb) { showLoginError("Firebase is not configured. Add your project keys to data.js."); return; }
  try {
    const provider = new fb.authM.GoogleAuthProvider();
    provider.setCustomParameters({ prompt: "select_account" });
    await fb.authM.signInWithPopup(fb.auth, provider);
  } catch (e) {
    if (e && (e.code === "auth/popup-closed-by-user" || e.code === "auth/cancelled-popup-request")) return;
    if (e && e.code === "auth/popup-blocked") {
      // Phones often block pop-ups: fall back to a full-page redirect.
      try { await fb.authM.signInWithRedirect(fb.auth, new fb.authM.GoogleAuthProvider()); return; } catch (e2) { e = e2; }
    }
    showLoginError("Sign-in failed: " + (e && e.message ? e.message : "unknown error"));
  }
}

async function signOutUser() {
  if (hasUnsynced()) {
    const ok = await confirmDialog({
      title: "Sign out with unsaved cloud changes?",
      message: "Your edits stay in this browser for your account, but they have not been pushed to Firebase yet. Sign out anyway?",
      okText: "Sign Out"
    });
    if (!ok) return;
  }
  persistNow();
  isAuthorizedSession = false;
  currentUser = null;
  state = emptyState();
  invalidate();
  leaveApp();
  if (fb) await fb.authM.signOut(fb.auth);
}

/* ---- Cloud save / load (with 6-month history & _meta config) ---- */
async function saveToCloud() {
  if (!isAuthorizedSession || !fb || !currentUser) { toast("Please sign in with an authorised account to save.", "warn"); return; }
  if (!navigator.onLine) { toast("You are offline. Your work is safe in this browser; save again when you are back online.", "warn"); return; }

  if (isValidDate(state.tbDate)) {
    Object.keys(state.dirtyGls).forEach(function (no) {
      if (state.dirtyGls[no] && GL_BY_NO[no]) saveGlToMonthHistory(no, state.tbDate);
    });
  }
  pruneToSixMonths(state);

  const glNos = Object.keys(state.dirtyGls).filter(function (k) { return state.dirtyGls[k] && GL_BY_NO[k]; });
  if (!glNos.length && !state.metaDirty) {
    toast("Nothing new to save. The cloud is up to date.", "info");
    return;
  }

  const btns = [$("btnSaveCloud"), $("bnSave")];
  btns.forEach(function (b) { setBusy(b, true, "Saving..."); });
  try {
    const fs = fb.fs;

    // 1. Save _meta document (active Current Month-End, 6-month list & per-Gmail Officer profile)
    if (state.metaDirty || glNos.length) {
      const emailKey = String(currentUser.email || "").trim().toLowerCase();
      const metaPayload = {
        currentMonthEnd: state.tbDate || "",
        savedMonths: getSavedMonthsList(state).slice(-MAX_SAVED_MONTHS),
        updatedBy: currentUser.email || "",
        updatedAt: fs.serverTimestamp()
      };
      if (emailKey) {
        metaPayload.officers = {};
        metaPayload.officers[emailKey] = {
          officerName: String(state.officerName || "").trim(),
          serviceNo: String(state.serviceNo || "").trim(),
          email: emailKey,
          updatedAt: fs.serverTimestamp()
        };
      }
      await fs.setDoc(fs.doc(fb.db, CFG.firestore.collection, META_DOC_ID), metaPayload, { merge: true });
      state.metaDirty = false;
    }

    // 2. Save dirty GL documents (replacing `months` with the pruned <=6 months map)
    for (let i = 0; i < glNos.length; i += 400) {
      const chunk = glNos.slice(i, i + 400);
      const batch = fs.writeBatch(fb.db);
      chunk.forEach(function (no) {
        const g = readGl(no);
        const prevInfo = getPrevMonthFromDb(no);
        const prunedMonths = {};
        const keepSet = new Set(getSavedMonthsList(state).slice(-MAX_SAVED_MONTHS));
        Object.keys(g.months || {}).forEach(function (m) {
          if (keepSet.has(m)) prunedMonths[m] = g.months[m];
        });

        batch.set(fs.doc(fb.db, CFG.firestore.collection, no), {
          glNo: no,
          glName: GL_BY_NO[no] ? GL_BY_NO[no].glName : "",
          mode: g.mode,
          items: cleanItems(g).map(function (it) { return { date: it.date, ref: it.ref, desc: it.desc, amount: it.amount }; }),
          override: !!g.override,
          glBalance: getSystemGlRaw(g),
          diffNote: g.diffNote || "",
          tbDate: state.tbDate || "",
          prevMonthDate: prevInfo.hasPrev ? prevInfo.prevDate : "",
          prevMonthBalance: prevInfo.hasPrev ? String(prevInfo.prevCents / 100) : "",
          targetMonthBalance: getSystemGlRaw(g),
          months: prunedMonths,
          updatedBy: currentUser.email || "",
          updatedAt: fs.serverTimestamp()
        });
      });
      await batch.commit();
      chunk.forEach(function (no) { delete state.dirtyGls[no]; });
    }
    persistNow();
    renderMonthSettings();
    toast("Saved to cloud (retaining up to " + MAX_SAVED_MONTHS + " month-end reports).", "success");
  } catch (e) {
    console.error(e);
    toast("Cloud save failed: " + (e && e.message ? e.message : "unknown error") + ". Your data is still safe in this browser.", "error");
  } finally {
    btns.forEach(function (b) { setBusy(b, false); });
    schedule("badge");
  }
}

async function loadFromCloud(replace) {
  if (!isAuthorizedSession || !fb) throw new Error("Not signed in with an authorised account.");
  const fs = fb.fs;
  const snap = await fs.getDocs(fs.collection(fb.db, CFG.firestore.collection));
  if (replace) { state.gls = {}; state.dirtyGls = {}; state.savedMonths = []; }

  let cloudMonthEnd = "";
  const rawDocs = [];
  snap.forEach(function (d) {
    if (d.id === META_DOC_ID) {
      const m = d.data() || {};
      if (m.currentMonthEnd && isValidDate(normalizeDate(m.currentMonthEnd))) {
        cloudMonthEnd = normalizeDate(m.currentMonthEnd);
      }
      if (Array.isArray(m.savedMonths)) {
        m.savedMonths.forEach(function (x) {
          const nd = normalizeDate(x);
          if (isValidDate(nd) && state.savedMonths.indexOf(nd) === -1) state.savedMonths.push(nd);
        });
      }
      const emailKey = String((currentUser && currentUser.email) || "").trim().toLowerCase();
      if (emailKey && m.officers && typeof m.officers === "object" && m.officers[emailKey]) {
        const cloudOfficer = m.officers[emailKey];
        if (typeof cloudOfficer.officerName === "string") state.officerName = cloudOfficer.officerName;
        if (typeof cloudOfficer.serviceNo === "string") state.serviceNo = cloudOfficer.serviceNo;
      }
      return;
    }
    if (GL_BY_NO[d.id]) rawDocs.push({ id: d.id, data: d.data() });
  });

  if (!cloudMonthEnd) {
    // Infer the latest tbDate from GL documents if _meta didn't have one yet
    rawDocs.forEach(function (r) {
      const dt = normalizeDate(String(r.data && r.data.tbDate || ""));
      if (isValidDate(dt) && dt > cloudMonthEnd) cloudMonthEnd = dt;
    });
  }
  if (cloudMonthEnd) state.tbDate = cloudMonthEnd;

  let n = 0;
  rawDocs.forEach(function (r) {
    const g = normalizeCloudGl(r.data, state.tbDate);
    state.gls[r.id] = g;
    delete state.dirtyGls[r.id];
    n++;
  });

  pruneToSixMonths(state);
  if (isValidDate(state.tbDate)) {
    activateMonthEnd(state.tbDate, false);
    state.metaDirty = false;
  }

  localStorage.setItem(storageKeyForUser(KEYS.hasLoadedFromCloud, currentUser.email), "true");
  invalidate();
  persistNow();
  fillOfficerInputs();
  renderMonthSettings();
  renderAll();
  return n;
}


/* ==========================================================================
   6. APP SHELL, MONTH SETTINGS & VIEWS
   ========================================================================== */
function updateActiveMonthHeaderUI() {
  const cur = isValidDate(state.tbDate) ? state.tbDate : "";
  if ($("tbDate")) $("tbDate").value = cur;
  if ($("activeMonthDisplay")) {
    $("activeMonthDisplay").textContent = cur || "Not Set (Click to Set)";
    $("activeMonthDisplay").classList.toggle("text-amber-700", !cur);
  }

  // Check if ANY GL has Previous Month-End data in the database for the current month-end
  let prevDateFound = "";
  if (cur) {
    for (let i = 0; i < MASTER.length; i++) {
      const p = getPrevMonthFromDb(MASTER[i].glNo);
      if (p.hasPrev) { prevDateFound = p.prevDate; break; }
    }
  }
  const pill = $("prevMonthStatusPill");
  if (pill) {
    if (!cur) {
      pill.className = "inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-slate-200/80 text-slate-600 truncate";
      pill.innerHTML = "Prev: —";
    } else if (prevDateFound) {
      pill.className = "inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-blue-100 text-blue-800 border border-blue-300 truncate";
      pill.innerHTML = '<i class="fa-solid fa-database text-[10px]"></i> Prev DB: ' + esc(prevDateFound);
    } else {
      const expPrev = prevMonthEndOf(cur);
      pill.className = "inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-slate-200/80 text-slate-600 truncate";
      pill.innerHTML = "Prev (" + esc(expPrev) + "): No DB Data";
    }
  }
}

function renderMonthSettings() {
  if ($("settingCurrentMonthEnd") && document.activeElement !== $("settingCurrentMonthEnd")) {
    $("settingCurrentMonthEnd").value = state.tbDate || "";
  }
  if ($("quickMonthEndBtns")) {
    const recents = recentMonthEnds(6);
    $("quickMonthEndBtns").innerHTML = recents.map(function (dt) {
      const active = dt === state.tbDate;
      return '<button type="button" data-pick-month="' + dt + '" class="px-2.5 py-1 rounded-lg font-mono text-xs font-semibold border transition ' +
        (active
          ? "bg-blue-700 text-white border-blue-700 shadow-sm"
          : "bg-white text-slate-700 border-slate-300 hover:bg-blue-50 hover:border-blue-400") + '">' +
        dt + (active ? ' <i class="fa-solid fa-check ml-1"></i>' : "") + "</button>";
    }).join("");
  }

  const saved = getSavedMonthsList(state).slice(-MAX_SAVED_MONTHS);
  const expPrev = isValidDate(state.tbDate) ? prevMonthEndOf(state.tbDate) : "";
  if ($("savedMonthsCount")) $("savedMonthsCount").textContent = saved.length;
  if ($("savedMonthsList")) {
    if (!saved.length) {
      $("savedMonthsList").innerHTML = '<span class="text-xs text-slate-500 italic">No month-end reports saved in database yet.</span>';
    } else {
      $("savedMonthsList").innerHTML = saved.map(function (m) {
        const isCur = m === state.tbDate;
        const isPrev = expPrev && m.slice(0, 7) === expPrev.slice(0, 7);
        const badge = isCur
          ? '<span class="ml-1 px-1.5 py-0.2 rounded bg-blue-700 text-white text-[10px]">Current</span>'
          : isPrev
            ? '<span class="ml-1 px-1.5 py-0.2 rounded bg-indigo-600 text-white text-[10px]">Prev Month</span>'
            : "";
        return '<div class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border text-xs font-mono font-semibold ' +
          (isCur ? "bg-blue-50 border-blue-400 text-blue-900" : isPrev ? "bg-indigo-50 border-indigo-300 text-indigo-900" : "bg-white border-slate-200 text-slate-700") + '">' +
          '<i class="fa-regular fa-calendar-check"></i>' + esc(m) + badge +
          '<button type="button" data-del-month="' + esc(m) + '" class="ml-1 text-slate-400 hover:text-rose-600" title="Remove this month-end report from database"><i class="fa-solid fa-trash-can text-[11px]"></i></button>' +
          "</div>";
      }).join("");
    }
  }
}

function fillOfficerInputs() {
  $("officerName").value = state.officerName || "";
  $("serviceNo").value = state.serviceNo || "";
  updateActiveMonthHeaderUI();
  renderMonthSettings();
}

function enterApp(label) {
  $("loginScreen").classList.add("hidden");
  $("appShell").classList.remove("hidden");
  $("userEmail").textContent = label;
  fillOfficerInputs();

  setOfficerBarOpen(!state.officerName || !state.tbDate);

  switchView(ui.view);
  updateNetStatus();
  prefetchLibsWhenIdle();

  const loadedKey = storageKeyForUser(KEYS.hasLoadedFromCloud, currentUser && currentUser.email);
  if (localStorage.getItem(loadedKey) !== "true") openModal($("firstLoginModal"));
}

function leaveApp() {
  $("appShell").classList.add("hidden");
  $("loginScreen").classList.remove("hidden");
  document.body.classList.remove("show-editor");
  ["glList", "itemsBody", "glBalancesBody"].forEach(function (id) { $(id).innerHTML = ""; });
  ["settingsModal", "actionSheet", "validationModal", "shortcutsModal", "firstLoginModal"].forEach(function (id) { closeModal($(id)); });
}

function switchView(view) {
  ui.view = view === "balances" ? "balances" : "schedules";
  const isBal = ui.view === "balances";
  $("schedulesWorkspace").classList.toggle("hidden", isBal);
  $("glBalancesView").classList.toggle("hidden", !isBal);
  [["navTabSchedules", "bnSchedules", false], ["navTabBalances", "bnBalances", true]].forEach(function (t) {
    const on = t[2] === isBal;
    $(t[0]).classList.toggle("is-active", on);
    $(t[1]).classList.toggle("is-active", on);
  });
  renderAll();
  window.scrollTo({ top: 0, behavior: "auto" });
}

function setOfficerBarOpen(open) {
  $("officerBar").classList.toggle("is-open", open);
  $("officerToggle").setAttribute("aria-expanded", open);
  $("officerChevron").classList.toggle("rotate-180", open);
}

function showEditorOnMobile(show) {
  document.body.classList.toggle("show-editor", !!show);
  if (window.matchMedia("(max-width: 1023px)").matches) window.scrollTo({ top: 0, behavior: "auto" });
}

function updateNetStatus() {
  $("netStatus").classList.toggle("hidden", navigator.onLine);
}


/* ==========================================================================
   7. RENDERING
   ========================================================================== */
function updateBadge() {
  const dirty = hasUnsynced();
  const badge = $("memoryBadge");
  badge.dataset.state = dirty ? "unsaved" : "saved";
  badge.title = dirty ? "Unsaved cloud changes - press Save" : "All changes saved in browser memory";
  $("memoryBadgeText").textContent = dirty ? "Unsaved Cloud Changes" : "Saved in Browser Memory";
  badge.querySelector("i").className = "fa-solid " + (dirty ? "fa-cloud-arrow-up" : "fa-circle-check");
  document.body.dataset.dirty = dirty ? "true" : "false";
  $("bnSave").classList.toggle("has-dot", dirty);
}

function selectedList() { return MASTER.filter(function (gl) { return !!state.selected[gl.glNo]; }); }

function renderPresetButtons() {
  const sel = Object.keys(state.selected).filter(function (k) { return state.selected[k]; }).sort().join(",");
  window.PRESETS.forEach(function (p) {
    const nos = window.getPresetGlNumbers(p.id);
    const btn = document.querySelector('.preset-btn[data-preset="' + p.id + '"]');
    const cnt = document.querySelector('[data-preset-count="' + p.id + '"]');
    if (cnt) cnt.textContent = nos.length + " GLs";
    if (btn) btn.classList.toggle("is-active", nos.length > 0 && nos.slice().sort().join(",") === sel);
  });
}

function applyPreset(presetId) {
  const nos = window.getPresetGlNumbers(presetId);
  state.selected = {};
  nos.forEach(function (no) { state.selected[no] = true; });
  if (!state.activeGl || !state.selected[state.activeGl]) state.activeGl = nos.slice().sort()[0] || null;
  persistSoon();
  renderAll();
  toast("Loaded " + plural(nos.length, "GL") + " from " + presetId.replace("PRESET_", "Set ") + ".", "success");
}

function matchesSearch(gl, q) {
  return !q || gl.glNo.indexOf(q) !== -1 || gl.glName.toLowerCase().indexOf(q) !== -1;
}

function visibleGls() {
  const q = ui.search.trim().toLowerCase();
  return MASTER.filter(function (gl) {
    if (!matchesSearch(gl, q)) return false;
    switch (ui.filter) {
      case "selected":     return !!state.selected[gl.glNo];
      case "needs_update": return getGlComparison(gl.glNo).needsHighlight;
      case "issues":       return !!state.selected[gl.glNo] && glIssues(gl.glNo).length > 0;
      case "zero":         return readGl(gl.glNo).mode === "zero";
      case "balance":      return readGl(gl.glNo).mode === "balance";
      default:             return true;
    }
  });
}

/** Counts, badges, progress and the mobile summary line (cheap, uses caches). */
function renderCounts() {
  let zero = 0, bal = 0, needs = 0, selNeeds = 0, sel = 0, ready = 0, issues = 0;
  for (let i = 0; i < MASTER.length; i++) {
    const no = MASTER[i].glNo;
    const g = state.gls[no];
    if (g && g.mode === "balance") bal++; else zero++;
    const isSel = !!state.selected[no];
    if (getGlComparison(no).needsHighlight) { needs++; if (isSel) selNeeds++; }
    if (isSel) {
      sel++;
      if (glIssues(no).length) issues++; else ready++;
    }
  }
  const counts = { all: MASTER.length, selected: sel, needs_update: needs, issues: issues, zero: zero, balance: bal };
  for (const k in counts) {
    const el = document.querySelector('[data-count="' + k + '"]');
    if (el) el.textContent = counts[k];
  }
  $("selectionSummary").textContent = sel + " of " + MASTER.length + " GLs selected";

  const pct = sel ? Math.round((ready / sel) * 100) : 0;
  $("progressText").textContent = ready + " / " + sel + " ready";
  $("progressBar").style.width = pct + "%";
  $("progressBar").classList.toggle("is-complete", sel > 0 && ready === sel);

  const badgeCount = sel ? selNeeds : needs;
  [$("navBalancesBadge"), $("bnBalancesBadge")].forEach(function (b) {
    b.textContent = badgeCount;
    b.classList.toggle("hidden", badgeCount === 0);
  });

  $("officerSummary").textContent =
    (state.officerName || "Officer not set") + " · " + (state.tbDate || "No date") + " · " + plural(sel, "GL");
}

function renderGlList() {
  const list = visibleGls();
  const parts = new Array(list.length);
  for (let i = 0; i < list.length; i++) {
    const gl = list[i];
    const no = gl.glNo;
    const g = readGl(no);
    const sel = !!state.selected[no];
    const cmp = getGlComparison(no);

    let tags = "";
    if (cmp.status === "changed") {
      tags += '<span class="gl-tag is-needs-update"><i class="fa-solid fa-bolt mr-1"></i>Needs Update</span>';
      if (cmp.isScheduleTallied) tags += ' <span class="gl-tag is-tallied">Tallied</span>';
    } else if (cmp.status === "unchanged") {
      tags += '<span class="gl-tag is-reusable"><i class="fa-solid fa-check mr-1"></i>Same TB</span>';
    }
    if (g.mode === "balance") {
      const n = cleanItems(g).length;
      const c = calcGl(g);
      tags += " " + (n ? '<span class="gl-tag is-items">' + plural(n, "item") + "</span>" : '<span class="gl-tag is-empty">No items</span>');
      if (c.hasDiscrepancy) {
        tags += ' <span class="gl-tag is-empty" title="System GL (' + formatMoney(c.gl) + ') ≠ Trial Balance (' + formatMoney(c.tb) + ')"><i class="fa-solid fa-scale-unbalanced mr-1"></i>Diff</span>';
      }
    }
    if (sel) {
      const iss = glIssues(no);
      tags += iss.length
        ? ' <i class="fa-solid fa-circle-exclamation gl-status-issue" title="' + esc(iss.join(" ")) + '"></i>'
        : ' <i class="fa-solid fa-circle-check gl-status-ok" title="Ready to export"></i>';
    }

    const cls = "gl-row" + (sel ? " is-selected" : "") + (no === state.activeGl ? " is-active" : "") +
      (cmp.status === "changed" ? " is-needs-update" : cmp.status === "unchanged" ? " is-reusable" : "");

    parts[i] =
      '<div class="' + cls + '" role="listitem" tabindex="0" data-gl="' + no + '">' +
        '<input type="checkbox" class="gl-check" aria-label="Select GL ' + no + '"' + (sel ? " checked" : "") + ">" +
        '<div class="gl-info"><div class="gl-no flex flex-wrap items-center gap-1.5">' + no + " " + tags + "</div>" +
        '<div class="gl-name">' + esc(gl.glName) + "</div></div>" +
        (sel
          ? '<div class="gl-toggle">' +
              '<button type="button" class="gl-toggle-btn' + (g.mode === "zero" ? " is-on" : "") + '" data-val="zero">Zero</button>' +
              '<button type="button" class="gl-toggle-btn' + (g.mode === "balance" ? " is-on" : "") + '" data-val="balance">Has Balance</button>' +
            "</div>"
          : "") +
      "</div>";
  }
  $("glList").innerHTML = parts.join("");
  $("glListEmpty").classList.toggle("hidden", list.length > 0);
}

/* ---- Editor ---- */
function navList() {
  if (state.activeGl && state.selected[state.activeGl]) return selectedList().map(function (g) { return g.glNo; });
  return visibleGls().map(function (g) { return g.glNo; });
}

function renderEditor() {
  const no = state.activeGl;
  const has = !!no && !!GL_BY_NO[no];
  $("editorEmpty").classList.toggle("hidden", has);
  $("editorPanel").classList.toggle("hidden", !has);
  if (!has) return;

  const g = readGl(no);
  $("activeGlNo").textContent = no;
  $("activeGlName").textContent = GL_BY_NO[no].glName;
  $("modeZero").classList.toggle("is-active", g.mode === "zero");
  $("modeBalance").classList.toggle("is-active", g.mode === "balance");
  $("modeZero").setAttribute("aria-pressed", g.mode === "zero");
  $("modeBalance").setAttribute("aria-pressed", g.mode === "balance");
  $("notSelectedNotice").classList.toggle("hidden", !!state.selected[no]);
  $("zeroPanel").classList.toggle("hidden", g.mode !== "zero");
  $("balancePanel").classList.toggle("hidden", g.mode !== "balance");

  const nav = navList();
  const pos = nav.indexOf(no);
  $("glPosition").textContent = pos >= 0 ? (pos + 1) + " / " + nav.length : "-";
  $("btnPrevGl").disabled = pos <= 0;
  $("btnNextGl").disabled = pos < 0 || pos >= nav.length - 1;

  if (g.mode === "balance") {
    renderItems();
    const c = calcGl(g);
    const notePlain = plainTextFromHtml(g.diffNote || "");
    const showDiffBox = !!g.override || c.diff !== 0 || notePlain !== "";
    $("chkOverride").checked = showDiffBox;
    $("overrideFields").classList.toggle("hidden", !showDiffBox);
    $("glBalanceInput").value = getSystemGlRaw(g);
    if (document.activeElement !== $("diffNote")) {
      $("diffNote").innerHTML = sanitizeRichNoteHtml(g.diffNote || "");
    }
  }
  updateTotals();
}

function renderItems(newIndex) {
  const g = readGl(state.activeGl);
  const parts = new Array(g.items.length);
  for (let i = 0; i < g.items.length; i++) {
    const it = g.items[i];
    const badAmt = String(it.amount).trim() !== "" && isNaN(parseAmount(it.amount));
    parts[i] =
      '<tr data-idx="' + i + '"' + (i === newIndex ? ' class="is-new"' : "") + ">" +
        '<td class="cell-serial">' + (i + 1) + "</td>" +
        '<td data-label="Date"><input class="cell-input cell-date" data-field="date" maxlength="10" inputmode="numeric" placeholder="YYYY.MM.DD" value="' + esc(it.date) + '"></td>' +
        '<td data-label="Reference"><input class="cell-input" data-field="ref" value="' + esc(it.ref) + '"></td>' +
        '<td data-label="Description" class="cell-wide"><input class="cell-input" data-field="desc" value="' + esc(it.desc) + '"></td>' +
        '<td data-label="Amount" class="cell-wide"><input class="cell-input cell-amount' + (badAmt ? " is-invalid" : "") + '" data-field="amount" inputmode="decimal" enterkeyhint="next" placeholder="0.00" value="' + esc(it.amount) + '"></td>' +
        '<td class="cell-actions text-center whitespace-nowrap">' +
          '<button type="button" class="row-dup" data-dup aria-label="Duplicate row ' + (i + 1) + '" title="Duplicate row"><i class="fa-regular fa-copy"></i></button>' +
          '<button type="button" class="row-delete" data-del aria-label="Delete row ' + (i + 1) + '" title="Delete row"><i class="fa-solid fa-trash-can"></i></button>' +
        "</td>" +
      "</tr>";
  }
  $("itemsBody").innerHTML = parts.join("");
  $("itemCount").textContent = plural(g.items.length, "item");
  $("itemsEmpty").classList.toggle("hidden", g.items.length > 0);
}

function updateTotals() {
  const no = state.activeGl;
  if (!no || !GL_BY_NO[no]) return;
  const g = readGl(no);
  const c = calcGl(g);
  const notePlain = plainTextFromHtml(g.diffNote || "");
  const noteHtml = sanitizeRichNoteHtml(g.diffNote || "");

  $("itemsTotal").textContent = formatMoney(c.tb);
  $("tbTotal").textContent = formatMoney(c.tb);

  const glInput = $("glBalanceInput");
  if (glInput) {
    if (document.activeElement !== glInput) {
      glInput.value = getSystemGlRaw(g);
    }
    glInput.placeholder = formatMoney(c.tb);
  }
  if ($("glBalanceSourceHint")) {
    $("glBalanceSourceHint").textContent = c.hasSysGl
      ? "System GL Balance: Rs. " + formatMoney(c.gl)
      : "Defaulting to Trial Balance (" + formatMoney(c.tb) + ") — type here or in GL Balances";
  }

  $("glBalanceAuto").textContent = c.diff === 0
    ? "Nil"
    : numberToWords(Math.abs(c.diff) / 100);
  $("glBalanceAuto").classList.toggle("has-difference", c.hasDiscrepancy);

  if (c.diff === 0 && !c.hasDiscrepancy) {
    $("glBalanceAutoHint").textContent = c.hasSysGl
      ? "System GL (" + formatMoney(c.gl) + ") tallies with Trial Balance."
      : "Using Trial Balance (" + formatMoney(c.tb) + "). Difference: Nil";
  } else if (c.diff !== 0) {
    $("glBalanceAutoHint").textContent = "Diff: " + (c.diff > 0 ? "+" : "-") + formatMoney(Math.abs(c.diff)) + " (GL: " + formatMoney(c.gl) + " vs TB: " + formatMoney(c.tb) + ")";
  } else {
    $("glBalanceAutoHint").textContent = "Enter the actual System GL Balance.";
  }

  const badge = $("reconcileBadge");
  if (badge) {
    if (c.hasDiscrepancy) {
      badge.className = "inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold bg-stamp-tint text-stamp border border-stamp/30";
      badge.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> ' +
        (c.diff !== 0 ? "Difference: " + numberToWords(Math.abs(c.diff) / 100) : "Difference Marked");
    } else {
      badge.className = "inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold bg-ledger-tint text-ledger border border-ledger/20";
      badge.innerHTML = '<i class="fa-solid fa-circle-check"></i> Tallied (GL = TB)';
    }
  }

  const showDiffBox = !!g.override || c.diff !== 0 || notePlain !== "";
  $("chkOverride").checked = showDiffBox;
  $("overrideFields").classList.toggle("hidden", !showDiffBox);

  const warnBox = $("diffLogicWarning");
  if (warnBox) {
    warnBox.classList.toggle("hidden", !((g.override || notePlain !== "") && c.diff === 0));
  }

  $("tbWords").textContent = numberToWords(c.tb / 100);
  $("glWords").textContent = numberToWords(c.gl / 100);
  if ($("previewTbNum")) $("previewTbNum").textContent = "Rs. " + formatMoney(c.tb);
  if ($("previewGlNum")) $("previewGlNum").textContent = "Rs. " + formatMoney(c.gl);

  $("diffPreview").textContent = differenceText(g, c);
  $("diffPreview").classList.toggle("text-stamp", c.hasDiscrepancy);
  $("diffPreview").classList.toggle("font-bold", c.hasDiscrepancy);

  if (!c.hasDiscrepancy) {
    $("certPreview").textContent =
      "We hereby certify that the total balance of GL Account No." + no +
      " has been duly verified and tallied with the total amount of the corresponding Trial Balance as at " + (state.tbDate || "YYYY.MM.DD");
    $("certPreview").classList.remove("hidden");
    if ($("certOmittedNotice")) $("certOmittedNotice").classList.add("hidden");
  } else if (notePlain !== "") {
    $("certPreview").innerHTML = noteHtml;
    $("certPreview").classList.remove("hidden");
    if ($("certOmittedNotice")) $("certOmittedNotice").classList.add("hidden");
  } else {
    $("certPreview").innerHTML = "";
    $("certPreview").classList.add("hidden");
    if ($("certOmittedNotice")) $("certOmittedNotice").classList.remove("hidden");
  }

  $("sigPrepared").textContent = state.officerName || "-";
  $("itemCount").textContent = plural(g.items.length, "item");

  renderEditorCompareBanner(no);
}

function renderEditorCompareBanner(no) {
  const banner = $("editorCompareBanner");
  const cmp = getGlComparison(no);
  if (!cmp.hasPrev && !cmp.hasTarget) { banner.className = "hidden"; banner.innerHTML = ""; return; }

  const prevLabel = cmp.hasPrev ? "Prev Month-End (" + esc(cmp.prevDate) + ")" : "Prev Month-End (DB)";
  const prevValHtml = cmp.hasPrev ? formatMoney(cmp.prevCents) : '<span class="text-slate-400">No DB Data</span>';
  const newLabel = state.tbDate ? "Current Month (" + esc(state.tbDate) + ")" : "Current Month";
  const sign = function (n) { return n > 0 ? "+" : ""; };

  if (cmp.status === "changed") {
    banner.className = "mt-4 rounded-xl border-2 border-amber-400 bg-gradient-to-r from-amber-50 to-orange-50/70 p-3.5 text-sm space-y-2 shadow-sm";
    banner.innerHTML =
      '<div class="flex flex-wrap items-center justify-between gap-2">' +
        '<span class="font-bold text-amber-900 flex items-center gap-1.5"><i class="fa-solid fa-bolt text-amber-600"></i> GL Balance changed — Update Trial Balance items below</span>' +
        (cmp.isScheduleTallied
          ? '<span class="gl-tag is-tallied"><i class="fa-solid fa-circle-check mr-1"></i>Trial Balance tallied with System GL</span>'
          : '<span class="gl-tag is-needs-update">TB Diff: ' + sign(cmp.diffFromSchedule) + formatMoney(cmp.diffFromSchedule) + "</span>") +
      "</div>" +
      '<div class="grid grid-cols-3 gap-2 pt-1 text-xs">' +
        '<div class="bg-white/90 rounded-lg px-2.5 py-1.5 border border-amber-200"><div class="text-slate-500">' + prevLabel + '</div><div class="font-mono font-semibold text-sm">' + prevValHtml + "</div></div>" +
        '<div class="bg-white/90 rounded-lg px-2.5 py-1.5 border border-amber-200"><div class="text-slate-500">' + newLabel + '</div><div class="font-mono font-bold text-sm text-amber-800">' + formatMoney(cmp.targetCents) + "</div></div>" +
        '<div class="bg-white/90 rounded-lg px-2.5 py-1.5 border border-amber-200"><div class="text-slate-500">Change vs Prev</div><div class="font-mono font-bold text-sm text-amber-800">' + (cmp.hasPrev ? sign(cmp.diffFromPrev) + formatMoney(cmp.diffFromPrev) : "—") + "</div></div>" +
      "</div>";
  } else if (cmp.status === "unchanged") {
    banner.className = "mt-4 rounded-xl border border-blue-300 bg-blue-50/80 p-3.5 text-sm flex flex-wrap items-center justify-between gap-3";
    banner.innerHTML =
      '<div class="flex items-center gap-2 text-blue-900 font-medium"><i class="fa-solid fa-circle-check text-blue-600"></i>' +
      (cmp.hasPrev
        ? "<span>Same as previous month-end (" + formatMoney(cmp.prevCents) + "). The previous Trial Balance schedule can be reused.</span>"
        : "<span>Current Month GL Balance is Zero (0.00).</span>") +
      "</div>" +
      (cmp.hasPrev ? '<span class="text-xs font-mono text-blue-800 bg-blue-100 px-2 py-0.5 rounded">Prev DB: ' + esc(cmp.prevDate) + "</span>" : "");
  } else {
    banner.className = "mt-4 rounded-xl border border-slate-200 bg-slate-50 p-3 text-xs flex flex-wrap items-center justify-between gap-2";
    banner.innerHTML =
      "<span><strong>" + prevLabel + ':</strong> <span class="font-mono font-semibold">' + prevValHtml + "</span></span>" +
      '<button type="button" class="link-btn font-semibold" data-banner-action="go-balances">Enter Current Month GL Balance in GL Balances &rarr;</button>';
  }
}

/* ---- GL Balances view ---- */
function renderGlBalancesView() {
  const base = ui.balScope === "selected" ? selectedList() : MASTER;

  let unchanged = 0, changed = 0, pend = 0;
  let commonPrevDate = "";
  for (let i = 0; i < base.length; i++) {
    const cmp = getGlComparison(base[i].glNo);
    if (!commonPrevDate && cmp.hasPrev && cmp.prevDate) commonPrevDate = cmp.prevDate;
    if (cmp.status === "unchanged") unchanged++;
    else if (cmp.status === "changed") changed++;
    else pend++;
  }
  $("balStatSelected").textContent = base.length;
  $("balStatUnchanged").textContent = unchanged;
  $("balStatChanged").textContent = changed;
  $("balStatPending").textContent = pend;
  $("balStatPrevDateHint").textContent = commonPrevDate
    ? "Prev Month DB: " + commonPrevDate
    : "Prev Month DB: No Data (—)";
  $("balHeaderNewDate").textContent = state.tbDate || "Current Month";
  $("balScopeSelected").classList.toggle("is-active", ui.balScope === "selected");
  $("balScopeAll").classList.toggle("is-active", ui.balScope === "all");
  document.querySelectorAll("[data-kpi-filter]").forEach(function (k) {
    k.classList.toggle("is-active", k.dataset.kpiFilter === ui.balFilter && ui.balFilter !== "all");
  });
  document.querySelectorAll("#balFilterPills [data-bfilter]").forEach(function (x) {
    x.classList.toggle("is-active", x.dataset.bfilter === ui.balFilter);
  });

  const q = ui.balSearch.trim().toLowerCase();
  const rows = [];
  for (let i = 0; i < base.length; i++) {
    const gl = base[i];
    if (!matchesSearch(gl, q)) continue;
    const no = gl.glNo;
    const cmp = getGlComparison(no);
    if (ui.balFilter !== "all" && cmp.status !== ui.balFilter) continue;

    const itemCnt = cleanItems(readGl(no)).length;
    const rowCls = "bal-row" + (cmp.status === "changed" ? " is-needs-update" : cmp.status === "unchanged" ? " is-reusable" : "");

    const prevDateHtml = cmp.hasPrev
      ? '<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-slate-100 text-slate-700 font-mono text-xs"><i class="fa-solid fa-database text-[10px] text-blue-600"></i>' + esc(cmp.prevDate) + "</span>"
      : '<span class="text-slate-400 font-mono text-xs">—</span>';

    const prevBalHtml = cmp.hasPrev
      ? '<span class="font-mono font-semibold text-slate-700">' + formatMoney(cmp.prevCents) + "</span>"
      : '<span class="text-slate-400 font-mono">—</span>';

    let diffHtml = '<span class="text-slate-400 font-mono">—</span>';
    if (cmp.hasTarget && cmp.hasPrev) {
      diffHtml = cmp.diffFromPrev === 0
        ? '<span class="font-mono font-semibold text-blue-700">Nil (0.00)</span>'
        : '<span class="font-mono font-bold text-amber-700">' + (cmp.diffFromPrev > 0 ? "+" : "") + formatMoney(cmp.diffFromPrev) + "</span>";
    }

    let statusHtml;
    if (cmp.status === "unchanged") {
      statusHtml = cmp.hasPrev
        ? '<span class="gl-tag is-reusable"><i class="fa-solid fa-circle-check mr-1"></i>Same — Reuse Prev TB</span>'
        : '<span class="gl-tag is-reusable"><i class="fa-solid fa-check mr-1"></i>Zero Balance</span>';
    } else if (cmp.status === "changed") {
      statusHtml = '<div class="space-y-1"><span class="gl-tag is-needs-update"><i class="fa-solid fa-bolt mr-1"></i>Update Trial Balance</span>' +
        (cmp.isScheduleTallied
          ? '<div class="text-[11px] font-semibold text-blue-700"><i class="fa-solid fa-check-double mr-1"></i>TB Tallied with GL</div>'
          : '<div class="text-[11px] font-semibold text-amber-800"><i class="fa-solid fa-scale-unbalanced mr-1"></i>TB Diff: ' + (cmp.diffFromSchedule > 0 ? "+" : "") + formatMoney(cmp.diffFromSchedule) + "</div>") +
        "</div>";
    } else {
      statusHtml = '<span class="inline-flex items-center gap-1 text-xs text-slate-500 font-medium"><i class="fa-regular fa-keyboard"></i> Enter Balance</span>';
    }

    rows.push(
      '<tr class="' + rowCls + '" data-gl="' + no + '">' +
        '<td class="bal-no font-mono font-bold text-blue-800">' + no + "</td>" +
        '<td class="bal-name"><div class="font-medium text-slate-800">' + esc(gl.glName) + '</div><div class="text-xs text-slate-500">' +
          (itemCnt ? plural(itemCnt, "TB item") : "No TB items") + "</div></td>" +
        '<td data-label="Prev Month-End">' + prevDateHtml + "</td>" +
        '<td data-label="Prev Balance (DB)" class="text-right">' + prevBalHtml + "</td>" +
        '<td data-label="Current GL Balance" class="bal-input text-right"><div class="flex items-center justify-end">' +
          '<input type="text" inputmode="decimal" enterkeyhint="next" class="field-input font-mono font-semibold text-right max-w-[165px] py-1.5 px-2.5 text-xs bal-target-input" data-gl="' + no + '" placeholder="Enter GL Balance..." value="' + esc(cmp.targetRaw) + '" aria-label="Current month-end GL balance for GL ' + no + '" />' +
        "</div></td>" +
        '<td data-label="Change vs Prev" class="text-right">' + diffHtml + "</td>" +
        '<td data-label="Trial Balance Status">' + statusHtml + "</td>" +
        '<td class="bal-action text-center"><button type="button" class="btn ' + (cmp.status === "changed" ? "btn-primary" : "btn-ghost") + ' btn-sm" data-open-schedule="' + no + '">' +
          '<i class="fa-solid ' + (cmp.status === "changed" ? "fa-pen-to-square" : "fa-folder-open") + '"></i> ' + (cmp.status === "changed" ? "Update TB" : "Open TB") + "</button></td>" +
      "</tr>"
    );
  }
  $("glBalancesBody").innerHTML = rows.join("");
  $("glBalancesEmpty").classList.toggle("hidden", rows.length > 0);
}


/* ==========================================================================
   8. ACTIONS
   ========================================================================== */
function setMode(glNo, mode) {
  const g = ensureGl(glNo);
  if (g.mode === mode) return;
  g.mode = mode;
  markDirty(glNo);
  schedule("list", state.activeGl === glNo ? "editor" : "counts");
}

function toggleSelect(glNo, on) {
  if (on) state.selected[glNo] = true; else delete state.selected[glNo];
  invalidate(glNo);
  persistSoon();
  schedule("presets", "list", "counts", "balances");
  if (state.activeGl === glNo) schedule("editor");
}

function openGl(glNo) {
  state.activeGl = glNo;
  persistSoon();
  if (ui.view !== "schedules") switchView("schedules");
  schedule("list", "editor");
  showEditorOnMobile(true);
}

function stepGl(dir) {
  const nav = navList();
  if (!nav.length) return;
  const pos = nav.indexOf(state.activeGl);
  const next = pos < 0 ? nav[0] : nav[pos + dir];
  if (next) openGl(next);
}

function addRow() {
  const no = state.activeGl;
  if (!no || ui.view !== "schedules") return;
  const g = ensureGl(no);
  if (g.mode !== "balance") { g.mode = "balance"; }
  g.items.push({ date: "", ref: "", desc: "", amount: "" });
  markDirty(no);
  renderEditor();
  renderItems(g.items.length - 1);
  schedule("list");
  const last = $("itemsBody").querySelector("tr:last-child .cell-date");
  if (last) last.focus();
}

function updateTargetMonthBalance(glNo, rawValue) {
  const g = ensureGl(glNo);
  const clean = String(rawValue).trim();
  g.targetMonthBalance = clean;
  g.glBalance = clean;
  const n = parseAmount(clean);
  if (clean !== "" && !isNaN(n)) {
    const cents = toCents(n);
    const prevInfo = getPrevMonthFromDb(glNo);
    if (prevInfo.hasPrev && cents === prevInfo.prevCents && cleanItems(g).length === 0 && prevInfo.prevSnap) {
      g.items = cloneItems(prevInfo.prevSnap.items);
    }
    if (cents !== 0 && g.mode === "zero") g.mode = "balance";
    else if (cents === 0 && cleanItems(g).length === 0) g.mode = "zero";
  }
  const c = calcGl(g);
  g.override = c.diff !== 0 || plainTextFromHtml(g.diffNote || "") !== "";
  markDirty(glNo);
}

function parseBalancePaste(text) {
  const rows = [];
  let skipped = 0;
  String(text || "").replace(/\r/g, "").split("\n").forEach(function (line) {
    if (!line.trim()) return;
    const m = line.match(/(^|\D)(\d{6})(?!\d)/);
    if (!m) { skipped++; return; }
    const rest = line.slice(m.index + m[0].length);
    let tokens = rest.split(/\t/).map(function (s) { return s.trim(); }).filter(Boolean);
    if (tokens.length < 2) tokens = rest.trim().split(/\s+/);
    let amt = NaN;
    for (let i = tokens.length - 1; i >= 0; i--) {
      const v = parseAmount(tokens[i]);
      if (!isNaN(v)) { amt = v; break; }
    }
    if (isNaN(amt)) { skipped++; return; }
    rows.push({ glNo: m[2], amount: amt });
  });
  return { rows: rows, skipped: skipped };
}

function focusField(id) {
  if (id === "tbDate") {
    openModal($("settingsModal"));
    renderMonthSettings();
    setTimeout(function () { if ($("settingCurrentMonthEnd")) $("settingCurrentMonthEnd").focus(); }, 60);
    return;
  }
  switchView("schedules");
  setOfficerBarOpen(true);
  showEditorOnMobile(false);
  setTimeout(function () { const el = $(id); if (el) { el.focus(); el.scrollIntoView({ block: "center" }); } }, 60);
}


/* ==========================================================================
   9. EVENTS
   ========================================================================== */
function closeTopModal() {
  if (isOpen("confirmModal")) { $("confirmCancel").click(); return true; }
  if (isOpen("validationModal")) { $("btnValidationClose").click(); return true; }
  for (const id of ["shortcutsModal", "actionSheet", "settingsModal"]) {
    if (isOpen(id)) { closeModal($(id)); return true; }
  }
  return false;
}

function bindEvents() {
  /* Login / logout */
  $("btnGoogleSignIn").addEventListener("click", signIn);
  $("btnLogout").addEventListener("click", signOutUser);

  /* Views */
  ["navTabSchedules", "navTabBalances", "bnSchedules", "bnBalances"].forEach(function (id) {
    $(id).addEventListener("click", function () {
      if ($(id).dataset.view === "schedules" && ui.view === "schedules") showEditorOnMobile(false);
      switchView($(id).dataset.view);
    });
  });

  /* Top bar + bottom nav */
  $("btnSaveCloud").addEventListener("click", saveToCloud);
  $("bnSave").addEventListener("click", saveToCloud);
  $("btnDownloadExcel").addEventListener("click", downloadExcel);
  $("btnDownloadPdf").addEventListener("click", downloadPdf);
  $("btnSettings").addEventListener("click", function () { renderMonthSettings(); openModal($("settingsModal")); });
  $("btnShortcuts").addEventListener("click", function () { openModal($("shortcutsModal")); });
  $("btnCloseShortcuts").addEventListener("click", function () { closeModal($("shortcutsModal")); });
  $("bnExport").addEventListener("click", function () { openSheet("export"); });
  $("bnMenu").addEventListener("click", function () { openSheet("menu"); });

  /* Action sheet */
  const sheetAct = function (fn) { return function () { closeModal($("actionSheet")); fn(); }; };
  $("sheetExcel").addEventListener("click", sheetAct(downloadExcel));
  $("sheetPdf").addEventListener("click", sheetAct(downloadPdf));
  $("sheetSettings").addEventListener("click", sheetAct(function () { renderMonthSettings(); openModal($("settingsModal")); }));
  $("sheetShortcuts").addEventListener("click", sheetAct(function () { openModal($("shortcutsModal")); }));
  $("sheetLogout").addEventListener("click", sheetAct(signOutUser));
  $("sheetClose").addEventListener("click", function () { closeModal($("actionSheet")); });

  /* Click on the dark backdrop closes simple dialogs */
  ["settingsModal", "actionSheet", "shortcutsModal"].forEach(function (id) {
    $(id).addEventListener("click", function (e) { if (e.target === $(id)) closeModal($(id)); });
  });
  $("btnCloseSettings").addEventListener("click", function () { closeModal($("settingsModal")); });

  /* Keyboard shortcuts */
  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape") { if (closeTopModal()) e.preventDefault(); return; }
    if (!isAuthorizedSession) return;
    const mod = e.ctrlKey || e.metaKey;
    const t = e.target;
    const typing = t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT" || t.isContentEditable);

    if (mod && e.key.toLowerCase() === "s") { e.preventDefault(); saveToCloud(); return; }
    if (e.altKey && e.key === "ArrowDown") { e.preventDefault(); stepGl(1); return; }
    if (e.altKey && e.key === "ArrowUp") { e.preventDefault(); stepGl(-1); return; }
    if (e.altKey && e.key.toLowerCase() === "n") { e.preventDefault(); addRow(); return; }
    if (typing || mod || e.altKey) return;
    if (e.key === "/") { e.preventDefault(); const s = ui.view === "balances" ? $("balSearch") : $("glSearch"); showEditorOnMobile(false); s.focus(); s.select(); }
    else if (e.key === "?") { openModal($("shortcutsModal")); }
    else if (e.key === "1") { switchView("schedules"); }
    else if (e.key === "2") { switchView("balances"); }
  });

  /* Officer details (saved per Gmail in local storage and Firestore DB) */
  $("officerToggle").addEventListener("click", function () {
    setOfficerBarOpen(!$("officerBar").classList.contains("is-open"));
  });
  $("officerName").addEventListener("input", function (e) {
    state.officerName = e.target.value;
    state.metaDirty = true;
    persistSoon();
    saveOfficerProfileSoon();
    schedule("totals", "counts", "badge");
  });
  $("serviceNo").addEventListener("input", function (e) {
    state.serviceNo = e.target.value;
    state.metaDirty = true;
    persistSoon();
    saveOfficerProfileSoon();
    schedule("counts", "badge");
  });
  $("officerName").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); $("serviceNo").focus(); } });

  function openMonthSettingsDialog() {
    renderMonthSettings();
    openModal($("settingsModal"));
    setTimeout(function () { if ($("settingCurrentMonthEnd")) $("settingCurrentMonthEnd").focus(); }, 60);
  }
  if ($("btnOpenMonthSettings")) {
    $("btnOpenMonthSettings").addEventListener("click", openMonthSettingsDialog);
  }
  if ($("monthDisplayBox")) {
    $("monthDisplayBox").addEventListener("click", openMonthSettingsDialog);
  }

  /* Set Active Current Month-End in Settings */
  function applyMonthEndSetting(rawDate) {
    const v = normalizeDate(String(rawDate || "").trim());
    if (!isValidDate(v)) {
      if ($("settingCurrentMonthEnd")) $("settingCurrentMonthEnd").classList.add("is-invalid");
      toast("Enter a valid Month-End date as YYYY.MM.DD (e.g. 2026.10.31).", "error");
      return;
    }
    if ($("settingCurrentMonthEnd")) {
      $("settingCurrentMonthEnd").value = v;
      $("settingCurrentMonthEnd").classList.remove("is-invalid");
    }
    activateMonthEnd(v, false);
    fillOfficerInputs();
    renderAll();

    // Immediately sync the active Current Month-End to Firebase `_meta` if online
    if (fb && currentUser && navigator.onLine) {
      fb.fs.setDoc(fb.fs.doc(fb.db, CFG.firestore.collection, META_DOC_ID), {
        currentMonthEnd: v,
        savedMonths: getSavedMonthsList(state).slice(-MAX_SAVED_MONTHS),
        updatedBy: currentUser.email || "",
        updatedAt: fb.fs.serverTimestamp()
      }, { merge: true }).then(function () {
        state.metaDirty = false;
        persistSoon();
        schedule("badge");
      }).catch(function () {});
    }

    // Check if previous month data exists in DB
    let prevCount = 0;
    const expPrev = prevMonthEndOf(v);
    MASTER.forEach(function (gl) { if (getPrevMonthFromDb(gl.glNo).hasPrev) prevCount++; });
    toast(
      prevCount
        ? "Current Month-End set to " + v + ". Loaded Previous Month (" + expPrev + ") data from DB for " + plural(prevCount, "GL") + "."
        : "Current Month-End set to " + v + ". No Previous Month (" + expPrev + ") data in DB.",
      "success"
    );
  }

  if ($("btnApplyCurrentMonth")) {
    $("btnApplyCurrentMonth").addEventListener("click", function () {
      applyMonthEndSetting($("settingCurrentMonthEnd") ? $("settingCurrentMonthEnd").value : "");
    });
  }
  if ($("settingCurrentMonthEnd")) {
    $("settingCurrentMonthEnd").addEventListener("keydown", function (e) {
      if (e.key === "Enter") { e.preventDefault(); applyMonthEndSetting(e.target.value); }
    });
  }
  if ($("quickMonthEndBtns")) {
    $("quickMonthEndBtns").addEventListener("click", function (e) {
      const b = e.target.closest("[data-pick-month]");
      if (b) applyMonthEndSetting(b.dataset.pickMonth);
    });
  }
  if ($("savedMonthsList")) {
    $("savedMonthsList").addEventListener("click", async function (e) {
      const delBtn = e.target.closest("[data-del-month]");
      if (!delBtn) return;
      const m = delBtn.dataset.delMonth;
      const ok = await confirmDialog({
        title: "Delete Month-End Report " + m + "?",
        message: "This removes the saved " + m + " report from the 6-month database history.",
        okText: "Delete Report"
      });
      if (!ok) return;
      state.savedMonths = (state.savedMonths || []).filter(function (x) { return x !== m; });
      Object.keys(state.gls).forEach(function (no) {
        if (state.gls[no] && state.gls[no].months && state.gls[no].months[m]) {
          delete state.gls[no].months[m];
          state.dirtyGls[no] = true;
        }
      });
      state.metaDirty = true;
      invalidate();
      persistSoon();
      fillOfficerInputs();
      renderAll();
      toast("Removed " + m + " from saved reports.", "info");
    });
  }

  /* Presets */
  $("presetButtons").addEventListener("click", function (e) {
    const b = e.target.closest(".preset-btn");
    if (b) applyPreset(b.dataset.preset);
  });

  /* Search + filters */
  const onSearch = debounce(function (v) { ui.search = v; schedule("list"); }, 120);
  $("glSearch").addEventListener("input", function (e) { onSearch(e.target.value); });
  $("filterTabs").addEventListener("click", function (e) {
    const t = e.target.closest(".filter-tab");
    if (!t) return;
    ui.filter = t.dataset.filter;
    document.querySelectorAll("#filterTabs .filter-tab").forEach(function (x) { x.classList.toggle("is-active", x === t); });
    schedule("list");
  });
  $("btnSelectVisible").addEventListener("click", function () {
    visibleGls().forEach(function (gl) { state.selected[gl.glNo] = true; });
    persistSoon(); schedule("presets", "list", "counts", "editor");
  });
  $("btnClearSelection").addEventListener("click", async function () {
    if (!Object.keys(state.selected).length) return;
    const ok = await confirmDialog({ title: "Clear all selected GLs?", message: "Only the ticks are removed. GL data is kept.", okText: "Clear", danger: false });
    if (!ok) return;
    state.selected = {};
    persistSoon(); schedule("presets", "list", "counts", "editor");
  });

  /* GL list */
  $("glList").addEventListener("click", function (e) {
    const row = e.target.closest(".gl-row");
    if (!row) return;
    const no = row.dataset.gl;
    if (e.target.closest(".gl-check")) { toggleSelect(no, e.target.checked); return; }
    const tb = e.target.closest(".gl-toggle-btn");
    if (tb) { setMode(no, tb.dataset.val); return; }
    openGl(no);
  });
  $("glList").addEventListener("keydown", function (e) {
    if ((e.key === "Enter" || e.key === " ") && e.target.classList.contains("gl-row")) {
      e.preventDefault();
      openGl(e.target.dataset.gl);
    }
  });

  /* Editor toolbar */
  $("btnBackToList").addEventListener("click", function () { showEditorOnMobile(false); });
  $("btnPrevGl").addEventListener("click", function () { stepGl(-1); });
  $("btnNextGl").addEventListener("click", function () { stepGl(1); });
  $("modeZero").addEventListener("click", function () { if (state.activeGl) setMode(state.activeGl, "zero"); });
  $("modeBalance").addEventListener("click", function () { if (state.activeGl) setMode(state.activeGl, "balance"); });
  $("btnSelectActive").addEventListener("click", function () { if (state.activeGl) toggleSelect(state.activeGl, true); });
  $("editorCompareBanner").addEventListener("click", function (e) {
    const b = e.target.closest("[data-banner-action]");
    if (!b || !state.activeGl) return;
    if (b.dataset.bannerAction === "go-balances") {
      const no = state.activeGl;
      switchView("balances");
      setTimeout(function () {
        const inp = $("glBalancesBody").querySelector('.bal-target-input[data-gl="' + no + '"]');
        if (inp) { inp.focus(); inp.select(); }
      }, 60);
    }
  });
  if ($("btnGoToGlBalances")) {
    $("btnGoToGlBalances").addEventListener("click", function () {
      const no = state.activeGl;
      if (!no) return;
      if (!state.selected[no]) ui.balScope = "all";
      switchView("balances");
      setTimeout(function () {
        const inp = $("glBalancesBody").querySelector('.bal-target-input[data-gl="' + no + '"]');
        if (inp) { inp.focus(); inp.select(); inp.scrollIntoView({ block: "center" }); }
      }, 60);
    });
  }

  /* Quick paste of schedule rows */
  $("pasteToggle").addEventListener("click", function () {
    const open = $("pasteBody").classList.toggle("hidden") === false;
    $("pasteToggle").setAttribute("aria-expanded", open);
    $("pasteChevron").classList.toggle("rotate-180", open);
    if (open) $("pasteArea").focus();
  });
  function applyPaste(replace) {
    const no = state.activeGl;
    if (!no) return;
    const res = parsePastedRows($("pasteArea").value);
    if (!res.items.length) { $("pasteResult").textContent = "No valid rows found. Check the column order."; return; }
    const g = ensureGl(no);
    g.items = replace ? res.items : g.items.concat(res.items);
    markDirty(no);
    $("pasteArea").value = "";
    $("pasteResult").textContent = plural(res.items.length, "row") + (replace ? " loaded" : " added") + (res.skipped ? ", " + res.skipped + " skipped" : "") + ".";
    schedule("editor", "list");
  }
  $("btnPasteReplace").addEventListener("click", function () { applyPaste(true); });
  $("btnPasteAppend").addEventListener("click", function () { applyPaste(false); });

  /* Items table */
  $("btnAddRow").addEventListener("click", addRow);

  $("itemsBody").addEventListener("input", function (e) {
    const input = e.target.closest(".cell-input");
    if (!input) return;
    const idx = +input.closest("tr").dataset.idx;
    const no = state.activeGl;
    const g = ensureGl(no);
    const field = input.dataset.field;
    g.items[idx][field] = input.value;
    if (field === "amount") input.classList.toggle("is-invalid", input.value.trim() !== "" && isNaN(parseAmount(input.value)));
    markDirty(no);
    schedule("totals", "list");
  });
  $("itemsBody").addEventListener("change", function (e) {
    const input = e.target.closest(".cell-date");
    if (!input) return;
    const idx = +input.closest("tr").dataset.idx;
    const v = normalizeDate(input.value);
    input.value = v;
    ensureGl(state.activeGl).items[idx].date = v;
    input.classList.toggle("is-invalid", !!v && !isValidDate(v));
    markDirty(state.activeGl);
  });
  // Enter = next cell; Enter on the last amount adds a new row.
  $("itemsBody").addEventListener("keydown", function (e) {
    if (e.key !== "Enter" || !e.target.classList.contains("cell-input")) return;
    e.preventDefault();
    const inputs = Array.prototype.slice.call($("itemsBody").querySelectorAll(".cell-input"));
    const i = inputs.indexOf(e.target);
    if (i === inputs.length - 1) addRow();
    else if (inputs[i + 1]) { inputs[i + 1].focus(); inputs[i + 1].select(); }
  });
  $("itemsBody").addEventListener("click", function (e) {
    const btn = e.target.closest("[data-del],[data-dup]");
    if (!btn) return;
    const idx = +btn.closest("tr").dataset.idx;
    const no = state.activeGl;
    const g = ensureGl(no);
    if (btn.hasAttribute("data-dup")) {
      g.items.splice(idx + 1, 0, Object.assign({}, g.items[idx]));
      markDirty(no);
      renderItems(idx + 1);
      schedule("totals", "list");
      return;
    }
    const removed = g.items.splice(idx, 1)[0];
    markDirty(no);
    renderItems();
    schedule("totals", "list");
    toast("Row " + (idx + 1) + " deleted.", "info", {
      label: "Undo",
      fn: function () {
        const gg = ensureGl(no);
        gg.items.splice(Math.min(idx, gg.items.length), 0, removed);
        markDirty(no);
        if (state.activeGl === no) schedule("editor");
        schedule("list");
      }
    });
  });

  /* System GL Balance & Rich-Text Optional Note handling */
  if ($("glBalanceInput")) {
    $("glBalanceInput").addEventListener("input", function (e) {
      const no = state.activeGl;
      if (!no) return;
      const g = ensureGl(no);
      const clean = String(e.target.value).trim();
      g.targetMonthBalance = clean;
      g.glBalance = clean;
      const bad = clean !== "" && isNaN(parseAmount(clean));
      e.target.classList.toggle("is-invalid", bad);
      const c = calcGl(g);
      g.override = c.diff !== 0 || plainTextFromHtml(g.diffNote || "") !== "";
      markDirty(no);
      schedule("totals", "list", "balances", "counts");
    });
  }
  $("chkOverride").addEventListener("change", function (e) {
    const no = state.activeGl;
    if (!no) return;
    const g = ensureGl(no);
    g.override = e.target.checked;
    const c = calcGl(g);
    if (!g.override && c.diff === 0) {
      g.diffNote = "";
      $("diffNote").innerHTML = "";
    }
    const showDiffBox = g.override || c.diff !== 0;
    $("overrideFields").classList.toggle("hidden", !showDiffBox);
    if (g.override) $("diffNote").focus();
    markDirty(no);
    schedule("totals", "list");
  });

  /* Rich Text Editor for Optional Note (#diffNote) */
  const RT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18];
  let savedRtRange = null;

  function saveRtSelection() {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const r = sel.getRangeAt(0);
    if ($("diffNote") && $("diffNote").contains(r.commonAncestorContainer)) {
      savedRtRange = r.cloneRange();
    }
  }

  function restoreRtSelection() {
    const ed = $("diffNote");
    if (!ed) return;
    ed.focus();
    if (savedRtRange) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(savedRtRange);
    }
  }

  function currentSelectionPt() {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return 10;
    let node = sel.anchorNode;
    if (node && node.nodeType === 3) node = node.parentElement;
    while (node && node !== $("diffNote")) {
      if (node.getAttribute && node.getAttribute("data-pt")) {
        const pt = parseFloat(node.getAttribute("data-pt"));
        if (!isNaN(pt)) return pt;
      }
      if (node.style && node.style.fontSize) {
        const m = String(node.style.fontSize).match(/^([\d.]+)\s*pt$/i);
        if (m) return parseFloat(m[1]);
      }
      node = node.parentElement;
    }
    return 10;
  }

  function updateRtToolbarState() {
    if (!$("diffNoteToolbar")) return;
    ["bold", "italic", "underline", "insertUnorderedList", "insertOrderedList"].forEach(function (cmd) {
      const btn = $("diffNoteToolbar").querySelector('[data-rt-cmd="' + cmd + '"]');
      if (!btn) return;
      let on = false;
      try { on = document.queryCommandState(cmd); } catch (e) {}
      btn.classList.toggle("is-active", !!on);
    });
    const pt = currentSelectionPt();
    if ($("rtFontSizeSelect") && document.activeElement !== $("rtFontSizeSelect")) {
      $("rtFontSizeSelect").value = String(RT_SIZES.indexOf(pt) !== -1 ? pt : 10);
    }
  }

  function syncDiffNoteFromDom() {
    const no = state.activeGl;
    if (!no) return;
    const g = ensureGl(no);
    const plain = plainTextFromHtml($("diffNote").innerHTML);
    g.diffNote = plain ? sanitizeRichNoteHtml($("diffNote").innerHTML) : "";
    if (plain !== "") g.override = true;
    markDirty(no);
    schedule("totals", "list");
  }

  function applyRtFontSize(pt) {
    const ed = $("diffNote");
    if (!ed) return;
    restoreRtSelection();
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    if (!ed.contains(range.commonAncestorContainer)) return;

    if (range.collapsed) {
      // Insert a styled zero-width space so newly typed characters get this size
      const span = document.createElement("span");
      span.setAttribute("data-pt", String(pt));
      span.style.fontSize = pt + "pt";
      span.innerHTML = "&#8203;";
      range.insertNode(span);
      const nr = document.createRange();
      nr.setStart(span.firstChild, 1);
      nr.collapse(true);
      sel.removeAllRanges();
      sel.addRange(nr);
      savedRtRange = nr.cloneRange();
    } else {
      document.execCommand("fontSize", false, "7");
      const fonts = ed.querySelectorAll('font[size="7"]');
      fonts.forEach(function (f) {
        const span = document.createElement("span");
        span.setAttribute("data-pt", String(pt));
        span.style.fontSize = pt + "pt";
        while (f.firstChild) span.appendChild(f.firstChild);
        f.parentNode.replaceChild(span, f);
      });
      saveRtSelection();
    }
    if ($("rtFontSizeSelect")) $("rtFontSizeSelect").value = String(pt);
    syncDiffNoteFromDom();
  }

  if ($("diffNoteToolbar")) {
    $("diffNoteToolbar").addEventListener("mousedown", function (e) {
      if (e.target.closest("button")) {
        e.preventDefault(); // keep text selection inside #diffNote
      }
    });
    $("diffNoteToolbar").addEventListener("click", function (e) {
      const cmdBtn = e.target.closest("[data-rt-cmd]");
      if (cmdBtn) {
        const cmd = cmdBtn.dataset.rtCmd;
        restoreRtSelection();
        document.execCommand(cmd, false, null);
        if (cmd === "removeFormat") {
          // Also strip custom font-size spans inside selection
          const sel = window.getSelection();
          if (sel && sel.rangeCount) {
            $("diffNote").querySelectorAll("span[data-pt]").forEach(function (sp) {
              if (sel.containsNode(sp, true)) {
                sp.removeAttribute("data-pt");
                sp.style.fontSize = "";
              }
            });
          }
        }
        saveRtSelection();
        updateRtToolbarState();
        syncDiffNoteFromDom();
        return;
      }
      const stepBtn = e.target.closest("[data-rt-size-step]");
      if (stepBtn) {
        const step = +stepBtn.dataset.rtSizeStep;
        const curPt = currentSelectionPt();
        let idx = RT_SIZES.indexOf(curPt);
        if (idx === -1) idx = RT_SIZES.indexOf(10);
        const nextIdx = Math.max(0, Math.min(RT_SIZES.length - 1, idx + step));
        applyRtFontSize(RT_SIZES[nextIdx]);
      }
    });
  }
  if ($("rtFontSizeSelect")) {
    $("rtFontSizeSelect").addEventListener("change", function (e) {
      const pt = parseFloat(e.target.value) || 10;
      applyRtFontSize(pt);
    });
  }

  $("diffNote").addEventListener("input", function () {
    saveRtSelection();
    updateRtToolbarState();
    syncDiffNoteFromDom();
  });
  ["keyup", "mouseup", "focus"].forEach(function (ev) {
    $("diffNote").addEventListener(ev, function () {
      saveRtSelection();
      updateRtToolbarState();
    });
  });

  /* GL Balances view */
  if ($("btnBalUploadExcel")) {
    $("btnBalUploadExcel").addEventListener("click", function () {
      renderMonthSettings();
      openModal($("settingsModal"));
    });
  }
  $("btnBalBackToSchedules").addEventListener("click", function () { switchView("schedules"); });
  $("balScopeSelected").addEventListener("click", function () { ui.balScope = "selected"; schedule("balances"); });
  $("balScopeAll").addEventListener("click", function () { ui.balScope = "all"; schedule("balances"); });
  $("balFilterPills").addEventListener("click", function (e) {
    const pill = e.target.closest("[data-bfilter]");
    if (!pill) return;
    ui.balFilter = pill.dataset.bfilter;
    schedule("balances");
  });
  document.querySelectorAll("[data-kpi-filter]").forEach(function (k) {
    k.addEventListener("click", function () {
      ui.balFilter = ui.balFilter === k.dataset.kpiFilter ? "all" : k.dataset.kpiFilter;
      schedule("balances");
    });
  });
  const onBalSearch = debounce(function (v) { ui.balSearch = v; schedule("balances"); }, 120);
  $("balSearch").addEventListener("input", function (e) { onBalSearch(e.target.value); });

  $("glBalancesBody").addEventListener("change", function (e) {
    const input = e.target.closest(".bal-target-input");
    if (!input) return;
    updateTargetMonthBalance(input.dataset.gl, input.value);
    schedule("balances", "counts");
  });
  // Enter saves the value and jumps to the next GL's balance field.
  $("glBalancesBody").addEventListener("keydown", function (e) {
    if (e.key !== "Enter" || !e.target.classList.contains("bal-target-input")) return;
    e.preventDefault();
    const inputs = Array.prototype.slice.call($("glBalancesBody").querySelectorAll(".bal-target-input"));
    const i = inputs.indexOf(e.target);
    const nextGl = inputs[i + 1] ? inputs[i + 1].dataset.gl : null;
    updateTargetMonthBalance(e.target.dataset.gl, e.target.value);
    renderGlBalancesView();
    renderCounts();
    if (nextGl) {
      const n = $("glBalancesBody").querySelector('.bal-target-input[data-gl="' + nextGl + '"]');
      if (n) { n.focus(); n.select(); }
    }
  });
  $("glBalancesBody").addEventListener("click", function (e) {
    const openBtn = e.target.closest("[data-open-schedule]");
    if (openBtn) {
      const no = openBtn.dataset.openSchedule;
      const cmp = getGlComparison(no);
      if (cmp.status === "changed" && cmp.targetCents !== 0 && readGl(no).mode !== "balance") {
        ensureGl(no).mode = "balance";
        markDirty(no);
      }
      openGl(no);
    }
  });

  /* Bulk paste of balances */
  $("balPasteToggle").addEventListener("click", function () {
    const open = $("balPasteBody").classList.toggle("hidden") === false;
    $("balPasteToggle").setAttribute("aria-expanded", open);
    $("balPasteChevron").classList.toggle("rotate-180", open);
    if (open) $("balPasteArea").focus();
  });
  $("btnBalPasteApply").addEventListener("click", function () {
    const res = parseBalancePaste($("balPasteArea").value);
    let applied = 0, unknown = 0;
    res.rows.forEach(function (r) {
      if (!GL_BY_NO[r.glNo]) { unknown++; return; }
      updateTargetMonthBalance(r.glNo, String(r.amount));
      applied++;
    });
    $("balPasteResult").textContent = applied
      ? plural(applied, "balance") + " applied" + (unknown ? ", " + unknown + " GL not in master list" : "") + (res.skipped ? ", " + res.skipped + " line(s) skipped" : "") + "."
      : "No matching GL balances found. Each line needs a 6-digit GL number and an amount.";
    if (applied) { $("balPasteArea").value = ""; schedule("balances", "counts"); toast(plural(applied, "current month balance") + " applied.", "success"); }
  });

  /* First-login modal */
  $("btnLoadYes").addEventListener("click", async function () {
    const btn = $("btnLoadYes");
    btn.disabled = true;
    try {
      const n = await loadFromCloud(false);
      closeModal($("firstLoginModal"));
      toast(n ? "Loaded " + plural(n, "GL account") + " from the cloud." : "The cloud has no saved data yet.", n ? "success" : "info");
    } catch (e) {
      toast("Could not load from the cloud: " + (e && e.message ? e.message : "unknown error"), "error");
    } finally { btn.disabled = false; }
  });
  $("btnLoadNo").addEventListener("click", function () {
    if (currentUser) localStorage.setItem(storageKeyForUser(KEYS.hasLoadedFromCloud, currentUser.email), "true");
    closeModal($("firstLoginModal"));
  });

  /* Settings: Super User Excel import */
  if ($("importMonthDate")) {
    $("importMonthDate").addEventListener("change", function (e) {
      e.target.value = normalizeDate(e.target.value);
    });
  }
  $("excelUpload").addEventListener("change", function (e) { handleImport(e.target.files); e.target.value = ""; });
  const dz = $("dropZone");
  ["dragenter", "dragover"].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add("is-dragover"); }); });
  ["dragleave", "drop"].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove("is-dragover"); }); });
  dz.addEventListener("drop", function (e) { handleImport(e.dataTransfer.files); });

  /* Settings: backup */
  $("btnBackupDownload").addEventListener("click", downloadBackup);
  $("backupUpload").addEventListener("change", function (e) { restoreBackup(e.target.files[0]); e.target.value = ""; });

  /* Settings: restore + clear */
  $("btnRestoreCloud").addEventListener("click", async function () {
    const ok = await confirmDialog({
      title: "Restore from Firebase Cloud?",
      message: "All GL data in this browser will be replaced by the cloud copy. Unsaved local changes will be lost.",
      okText: "Restore Data"
    });
    if (!ok) return;
    const btn = $("btnRestoreCloud");
    btn.disabled = true;
    try {
      const n = await loadFromCloud(true);
      closeModal($("settingsModal"));
      toast("Restored " + plural(n, "GL account") + " from the cloud.", "success");
    } catch (e) {
      toast("Restore failed: " + (e && e.message ? e.message : "unknown error"), "error");
    } finally { btn.disabled = false; }
  });

  $("btnClearCache").addEventListener("click", async function () {
    const ok = await confirmDialog({
      title: "Clear local browser cache?",
      message: "This removes all GL data, officer details and selections stored in this browser. Anything not saved to the cloud is lost.",
      okText: "Clear Cache"
    });
    if (!ok) return;
    clearTimeout(persistTimer);
    if (currentUser) {
      localStorage.removeItem(storageKeyForUser(KEYS.state, currentUser.email));
      localStorage.removeItem(storageKeyForUser(KEYS.hasLoadedFromCloud, currentUser.email));
    }
    localStorage.removeItem(KEYS.dirty);
    state = emptyState();
    invalidate();
    fillOfficerInputs();
    closeModal($("settingsModal"));
    renderAll();
    toast("Local browser cache cleared.", "success");
    openModal($("firstLoginModal"));
  });

  /* Connectivity + leaving the page */
  window.addEventListener("online", function () { updateNetStatus(); if (isAuthorizedSession) toast("Back online.", "success"); });
  window.addEventListener("offline", function () { updateNetStatus(); if (isAuthorizedSession) toast("You are offline. Edits are still saved in this browser.", "warn"); });
  window.addEventListener("beforeunload", function (e) {
    persistNow();
    if (isAuthorizedSession && hasUnsynced()) { e.preventDefault(); e.returnValue = ""; }
  });
  document.addEventListener("visibilitychange", function () { if (document.visibilityState === "hidden") persistNow(); });

  /* Leaving the phone layout resets the list/editor toggle */
  window.matchMedia("(min-width: 1024px)").addEventListener("change", function (mq) {
    if (mq.matches) document.body.classList.remove("show-editor");
  });
}


/* ==========================================================================
   10. IMPORT MONTH-END EXCEL FILES (Retained up to 6 Months)
   ========================================================================== */
async function handleImport(fileList) {
  const files = Array.prototype.slice.call(fileList || []).filter(function (f) { return /\.xlsx$/i.test(f.name); });
  const report = $("uploadReport");
  report.classList.remove("hidden");
  report.innerHTML = "";
  const line = function (kind, text) {
    const li = document.createElement("li");
    li.className = "is-" + kind;
    li.textContent = text;
    report.appendChild(li);
  };

  if (!files.length) { line("warn", "Please choose one or more .xlsx files."); return; }

  line("ok", "Reading " + plural(files.length, "file") + "...");
  try { await ensureExcelLibs(); }
  catch (e) { report.innerHTML = ""; line("error", e.message); return; }
  report.innerHTML = "";

  const manualMonthEnd = normalizeDate(($("importMonthDate") && $("importMonthDate").value || "").trim());

  let total = 0;
  let importedMonths = new Set();
  for (const file of files) {
    try {
      const blocks = await parseWorkbookFile(file);
      if (!blocks.length) { line("warn", file.name + ": no 'GL Account Number and Name:' blocks found."); continue; }
      blocks.forEach(function (b) {
        if (!GL_BY_NO[b.glNo]) { line("warn", file.name + ": GL " + b.glNo + " is not in the master list, skipped."); return; }
        const items = cloneItems(b.items);
        let tb = 0;
        items.forEach(function (it) { const n = parseAmount(it.amount); if (!isNaN(n)) tb += toCents(n); });
        const glCents = b.glBalance != null ? toCents(b.glBalance) : (b.tbBalance != null ? toCents(b.tbBalance) : tb);
        const override = glCents !== tb;
        const autoDiffWords = numberToWords(Math.abs(glCents - tb) / 100);
        const rawNote = String(b.diffNote || b.diffText || "").trim();
        const note = (/^nil$/i.test(rawNote) || rawNote.toLowerCase() === autoDiffWords.toLowerCase() || !isNaN(parseAmount(rawNote)))
          ? ""
          : rawNote;
        const monthEnd = normalizeDate(manualMonthEnd || b.tbDate || "");

        if (!isValidDate(monthEnd)) {
          line("error", file.name + " (GL " + b.glNo + "): Month-End Date not found in sheet. Enter the Month-End Date in the box above and re-upload.");
          return;
        }

        const g = ensureGl(b.glNo);
        const snapMode = items.length || glCents !== 0 || tb !== 0 ? "balance" : "zero";
        const glBalStr = String(glCents / 100);

        // Save into the 6-month database history under `monthEnd`
        g.months[monthEnd] = {
          mode: snapMode,
          items: items,
          glBalance: glBalStr,
          tbBalance: String(tb / 100),
          override: override,
          diffNote: override ? note : ""
        };
        if (state.savedMonths.indexOf(monthEnd) === -1) state.savedMonths.push(monthEnd);
        importedMonths.add(monthEnd);

        // If uploaded report matches current active month-end, update working schedule
        if (state.tbDate === monthEnd) {
          g.mode = snapMode;
          g.items = cloneItems(items);
          g.targetMonthBalance = glBalStr;
          g.glBalance = glBalStr;
          g.override = override;
          g.diffNote = override ? note : "";
        } else if (isValidDate(state.tbDate) && monthEnd.slice(0, 7) === prevMonthEndOf(state.tbDate).slice(0, 7)) {
          // Uploaded report is for the Previous Month-End of the active Current Month-End:
          // Pre-populate schedule items if the current month has no items yet
          if (cleanItems(g).length === 0 && items.length > 0) {
            g.items = cloneItems(items);
            g.mode = snapMode;
          }
        } else if (!isValidDate(state.tbDate)) {
          // No active month set yet: keep items ready
          g.items = cloneItems(items);
          g.mode = snapMode;
        }

        state.dirtyGls[b.glNo] = true;
        total++;
        line("ok", file.name + ": GL " + b.glNo + " saved for Month-End [" + monthEnd + "] — Bal " + formatMoney(glCents) + " (" + plural(items.length, "item") + ")");
      });
    } catch (e) {
      console.error(e);
      line("error", file.name + ": could not be read (" + (e && e.message ? e.message : "unknown error") + ").");
    }
  }

  state.metaDirty = true;
  const pruned = pruneToSixMonths(state);
  if (pruned.length) {
    line("warn", "Auto-pruned oldest month-end report(s) beyond 6-month limit: " + pruned.join(", "));
  }
  invalidate();
  persistNow();
  fillOfficerInputs();
  renderAll();
  if (total) {
    toast("Imported " + plural(total, "GL schedule") + " into database (" + Array.from(importedMonths).join(", ") + "). Press Save to Cloud to sync.", "success");
  }
}


/* ==========================================================================
   11. BACKUP / RESTORE (JSON file)
   ========================================================================== */
function saveBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
}

function downloadBackup() {
  if (!isAuthorizedSession) return;
  persistNow();
  const data = { app: "glTrialBalance", version: 1, exportedAt: new Date().toISOString(), exportedBy: currentUser.email || "", state: state };
  const stamp = new Date().toISOString().slice(0, 10);
  saveBlob(new Blob([JSON.stringify(data, null, 1)], { type: "application/json" }), "GL_TB_Backup_" + stamp + ".json");
  toast("Backup downloaded.", "success");
}

async function restoreBackup(file) {
  if (!file || !isAuthorizedSession) return;
  let data;
  try { data = JSON.parse(await file.text()); }
  catch (e) { toast("This is not a valid backup file.", "error"); return; }
  if (!data || data.app !== "glTrialBalance" || !data.state || typeof data.state.gls !== "object") {
    toast("This is not a GL Trial Balance backup file.", "error");
    return;
  }
  const ok = await confirmDialog({
    title: "Restore this backup?",
    message: "Backup from " + (data.exportedAt ? data.exportedAt.slice(0, 16).replace("T", " ") : "an unknown date") +
      ". All GL data in this browser will be replaced. Restored GLs will be marked for Save to Cloud.",
    okText: "Restore Backup"
  });
  if (!ok) return;
  state = hydrate(data.state);
  state.dirtyGls = {};
  Object.keys(state.gls).forEach(function (k) { state.dirtyGls[k] = true; });
  invalidate();
  persistNow();
  fillOfficerInputs();
  closeModal($("settingsModal"));
  renderAll();
  toast("Backup restored (" + plural(Object.keys(state.gls).length, "GL") + ").", "success");
}


/* ==========================================================================
   12. PRE-EXPORT CHECK + EXCEL / PDF
   ========================================================================== */
function showValidation(errors, warnings) {
  return new Promise(function (resolve) {
    const icon = function (kind) { return kind === "error" ? "fa-circle-xmark text-stamp" : "fa-triangle-exclamation text-amber"; };
    const item = function (kind, p) {
      return '<li class="val-item is-' + kind + '"' + (p.gl ? ' data-gl="' + p.gl + '"' : "") + (p.field ? ' data-field="' + p.field + '"' : "") + ">" +
        '<i class="fa-solid ' + icon(kind) + ' mt-0.5"></i>' +
        "<span>" + (p.gl ? '<strong class="font-mono">' + p.gl + "</strong> — " : "") + esc(p.msg) + "</span>" +
        '<span class="val-go">' + (p.gl ? "Open" : "Fix") + ' <i class="fa-solid fa-chevron-right"></i></span></li>';
    };
    $("validationList").innerHTML =
      errors.map(function (p) { return item("error", p); }).join("") +
      warnings.map(function (p) { return item("warn", p); }).join("");
    $("validationTitle").textContent = errors.length ? "Please fix these before exporting" : "Check before exporting";
    $("validationSummary").textContent =
      (errors.length ? plural(errors.length, "problem") + " must be fixed. " : "") +
      (warnings.length ? plural(warnings.length, "warning") + " (you can still export)." : "");
    $("btnValidationContinue").classList.toggle("hidden", errors.length > 0);

    const modal = $("validationModal");
    openModal(modal);
    const done = function (v) {
      closeModal(modal);
      $("btnValidationContinue").onclick = null;
      $("btnValidationClose").onclick = null;
      $("validationList").onclick = null;
      resolve(v);
    };
    $("btnValidationContinue").onclick = function () { done(true); };
    $("btnValidationClose").onclick = function () { done(false); };
    $("validationList").onclick = function (e) {
      const it = e.target.closest(".val-item");
      if (!it) return;
      done(false);
      if (it.dataset.gl) openGl(it.dataset.gl);
      else if (it.dataset.field) focusField(it.dataset.field);
    };
  });
}

async function collectExportParams() {
  const errors = [], warnings = [];
  if (!String(state.officerName).trim()) errors.push({ field: "officerName", msg: "Enter the Officer Name." });
  if (!String(state.serviceNo).trim()) errors.push({ field: "serviceNo", msg: "Enter the Service Number." });
  if (!isValidDate(state.tbDate)) errors.push({ field: "tbDate", msg: "Enter the Trial Balance Date as YYYY.MM.DD." });

  const chosen = selectedList();
  if (!chosen.length) errors.push({ msg: "Select at least one GL account (tick it in the list or choose a GL Set)." });

  const blocks = chosen.map(function (gl) {
    glIssues(gl.glNo).forEach(function (m) { errors.push({ gl: gl.glNo, msg: m }); });
    glWarnings(gl.glNo).forEach(function (m) { warnings.push({ gl: gl.glNo, msg: m }); });
    const g = readGl(gl.glNo);
    const items = g.mode === "balance" ? cleanItems(g) : [];
    const c = calcGl(g);
    return {
      glNo: gl.glNo,
      glName: gl.glName,
      mode: g.mode,
      items: items.map(function (it) {
        return { date: it.date, ref: it.ref, desc: it.desc, amount: Math.round(parseAmount(it.amount) * 100) / 100 };
      }),
      tbCents: c.tb,
      glCents: c.gl,
      diffText: differenceText(g, c),
      diffNote: String(g.diffNote || "").trim(),
      hasDiscrepancy: c.hasDiscrepancy
    };
  });

  if (errors.length || warnings.length) {
    const go = await showValidation(errors, warnings);
    if (!go || errors.length) return null;
  }
  return { officerName: state.officerName.trim(), serviceNo: state.serviceNo.trim(), tbDate: state.tbDate, blocks: blocks };
}

let exporting = false;

async function downloadExcel() {
  if (exporting) return;
  const params = await collectExportParams();
  if (!params) return;
  exporting = true;
  const btn = $("btnDownloadExcel");
  setBusy(btn, true, "Preparing...");
  try {
    await ensureExcelLibs();
    const buffer = await buildWorkbookBuffer(params);
    saveBlob(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      CFG.excel.fileNamePrefix + "_" + state.tbDate + ".xlsx");
    toast("Downloaded Excel with " + plural(params.blocks.length, "GL schedule") + ".", "success");
  } catch (e) {
    console.error(e);
    toast("Excel export failed: " + (e && e.message ? e.message : "unknown error"), "error");
  } finally {
    setBusy(btn, false);
    exporting = false;
  }
}

async function downloadPdf() {
  if (exporting) return;
  const params = await collectExportParams();
  if (!params) return;
  exporting = true;
  const btn = $("btnDownloadPdf");
  setBusy(btn, true, "Generating...");
  try {
    await ensurePdfLibs();
    await buildAndSavePdf(params, CFG.excel.fileNamePrefix + "_" + state.tbDate + ".pdf");
    toast("Downloaded PDF with " + plural(params.blocks.length, "GL schedule") + ".", "success");
  } catch (e) {
    console.error(e);
    toast("PDF export failed: " + (e && e.message ? e.message : "unknown error"), "error");
  } finally {
    setBusy(btn, false);
    exporting = false;
  }
}


/* ==========================================================================
   13. START
   ========================================================================== */
bindEvents();
initFirebase();
