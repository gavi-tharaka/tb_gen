/* ==========================================================================
   GL Trial Balance Generator - data.js
   Static configuration only. Loaded as a classic <script> before app.js.
   Everything is exposed on `window` so the ES module (app.js) can read it.

   Contents
     1. FIREBASE_CONFIG       - paste your Firebase project keys here
     2. APP_CONFIG            - storage keys, Firestore paths, defaults
     3. PRESETS               - the 4 rotatable GL sets
     4. MASTER_GL_LIST        - all GL accounts (edit here, not in the UI)
     5. Small lookup helpers
   ========================================================================== */


/* --------------------------------------------------------------------------
   1. FIREBASE CONFIG
   Firebase console -> Project settings -> General -> Your apps -> Web app.
   While any value still starts with "YOUR_", the app runs in offline demo mode.
   -------------------------------------------------------------------------- */
window.FIREBASE_CONFIG = {
  apiKey:            "AIzaSyC5-R8KWuw8qYKbyQmIkQxpAeTOAf7kAho",
  authDomain:        "bank-gl-trial-balance.firebaseapp.com",
  projectId:         "bank-gl-trial-balance",
  storageBucket:     "bank-gl-trial-balance.firebasestorage.app",
  messagingSenderId: "142817310526",
  appId:             "1:142817310526:web:582b7ed562e6322dfb9e6f"
};

/**
 * Optional convenience list. If it is not empty, app.js signs out any Google
 * account that is not on it. This is only a UI courtesy: real access control
 * must be enforced in Firestore Security Rules, because anyone can edit this
 * file in their own browser.
 * Example: ["officer1@gmail.com", "officer2@gmail.com"]
 */
window.ALLOWED_EMAILS = [];

/** True once real Firebase keys have been pasted in. */
window.isFirebaseConfigured = function () {
  const c = window.FIREBASE_CONFIG || {};
  const required = ["apiKey", "authDomain", "projectId", "appId"];
  return required.every(function (k) {
    return typeof c[k] === "string" && c[k].trim() !== "" && c[k].indexOf("YOUR_") !== 0;
  });
};


/* --------------------------------------------------------------------------
   2. APP CONFIG
   -------------------------------------------------------------------------- */
window.APP_CONFIG = {
  appName: "GL Trial Balance Generator",

  /* Maximum number of month-end reports kept in the database (older ones are auto-pruned) */
  maxSavedMonths: 6,

  /* localStorage keys. Bump the version suffix if the saved data shape changes. */
  storage: {
    state:             "glTb.state.v1",            // whole working state (officer, selections, items...)
    hasLoadedFromCloud:"glTb.hasLoadedFromCloud",  // "true" after the first-login prompt is answered
    dirty:             "glTb.unsyncedChanges"      // "true" when local data differs from the cloud
  },

  /* Firestore layout:  glTrialBalance (collection) / <glNo> (one document per GL) + _meta */
  firestore: {
    collection: "glTrialBalance",
    metaDocId:  "_meta"
  },

  /* Defaults used when a new browser has nothing saved yet */
  defaults: {
    officerName: "",
    serviceNo:   "",
    tbDate:      ""        // format YYYY.MM.DD
  },

  /* Export settings (used by the ExcelJS builder in app.js) */
  excel: {
    sheetName:    "Sheet1",
    fileNamePrefix: "GL_Trial_Balance",     // final name: GL_Trial_Balance_2026.09.30.xlsx
    columnWidths: [14, 16, 22, 42, 20],     // A..E
    signatureBoxRows: 6,
    gapBetweenBlocks: 3,
    checkedByOps:     "Operations manager",
    checkedByManager: "Manager",
    auditLabel:       "Audit officer :"
  }
};


/* --------------------------------------------------------------------------
   3. PRESETS  (the 4 sets that rotate between officers each month)
   -------------------------------------------------------------------------- */
window.PRESETS = [
  { id: "PRESET_1", label: "Set 1" },
  { id: "PRESET_2", label: "Set 2" },
  { id: "PRESET_3", label: "Set 3" },
  { id: "PRESET_4", label: "Set 4" }
];


/* --------------------------------------------------------------------------
   4. MASTER GL LIST
   One object per GL account:
     glNo          - string, keep as text so leading zeros / formatting never change
     glName        - exactly as it must print in the Excel header
     defaultPreset - PRESET_1 .. PRESET_4

   NOTE: This list currently holds 56 accounts. The brief mentions 60 in total,
   so four GLs are still missing. Append them at the end of the right group:
     { glNo: "......", glName: "........", defaultPreset: "PRESET_x" },
   Nothing else in the app needs to change; counts update automatically.
   -------------------------------------------------------------------------- */
window.MASTER_GL_LIST = [

  /* ---- Group 1 (Page 1) ---- */
  { glNo: "300800", glName: "LOCAL CHEQUES & PURCHASED",                  defaultPreset: "PRESET_1" },
  { glNo: "300910", glName: "FOREIGN CHEQUES & DRAFTS PURCHASED",         defaultPreset: "PRESET_1" },
  { glNo: "393110", glName: "S/A - STAFF TRAVELLING",                     defaultPreset: "PRESET_1" },
  { glNo: "393150", glName: "BANK DEPOSIT RECOVERABLE",                   defaultPreset: "PRESET_1" },
  { glNo: "393380", glName: "S/A - TRANSIT/ JCB CARD STOCKS",             defaultPreset: "PRESET_1" },
  { glNo: "393390", glName: "S/A - DEBIT CARD STOCKS",                    defaultPreset: "PRESET_1" },
  { glNo: "393430", glName: "S/A - MICR CHEQUE BOOKS",                    defaultPreset: "PRESET_1" },
  { glNo: "393460", glName: "S/A - RENT ADVANCE",                         defaultPreset: "PRESET_1" },
  { glNo: "393580", glName: "S/A - ALTERNATION TO PREMISES",              defaultPreset: "PRESET_1" },
  { glNo: "393620", glName: "S/A - STAFF UNIFORM",                        defaultPreset: "PRESET_1" },

  /* ---- Group 2 (Page 2) ---- */
  { glNo: "393630", glName: "S/A - PREMISES",                             defaultPreset: "PRESET_2" },
  { glNo: "393720", glName: "S/A - LEGAL CHARGES",                        defaultPreset: "PRESET_2" },
  { glNo: "393830", glName: "AMO DUE PB EPF O/A STAFF LN DISBURS",        defaultPreset: "PRESET_2" },
  { glNo: "393880", glName: "S/A - MISCELLANEOUS",                        defaultPreset: "PRESET_2" },
  { glNo: "394050", glName: "S/A - LEGAL AND OTHER CHARGES PAST DUE Ods", defaultPreset: "PRESET_2" },
  { glNo: "394340", glName: "S/A - DEBTORS",                              defaultPreset: "PRESET_2" },
  { glNo: "394680", glName: "S/A - REIM. RECVL ON INW REM. CBSL",         defaultPreset: "PRESET_2" },
  { glNo: "394682", glName: "S/A REIM. RCBLE ON TRN COST CBSL",           defaultPreset: "PRESET_2" },
  { glNo: "397200", glName: "S/A - RHO ADVANCES",                         defaultPreset: "PRESET_2" },
  { glNo: "397250", glName: "WESTERN UNION PAYMENT MADE ACC.",            defaultPreset: "PRESET_2" },
  { glNo: "397280", glName: "RIA- PAYMENT MADE ACC.",                     defaultPreset: "PRESET_2" },
  { glNo: "398100", glName: "S/A - INITIAL DEPOSIT ISURU UDANA",          defaultPreset: "PRESET_2" },
  { glNo: "399950", glName: "NP PAWNING SPURIOUS ARTICLES",               defaultPreset: "PRESET_2" },
  { glNo: "399960", glName: "NP PAWNING DISPUTED ARTICLES",               defaultPreset: "PRESET_2" },
  { glNo: "399970", glName: "PAWNING DISPUTES ARTICLES A/C",              defaultPreset: "PRESET_2" },
  { glNo: "399980", glName: "PAWNING SPURIOUS ARTICLES A/C",              defaultPreset: "PRESET_2" },
  { glNo: "399990", glName: "OPERATIONAL LOSS ACCOUNT",                   defaultPreset: "PRESET_2" },

  /* ---- Group 3 (Page 3 - Part A) ---- */
  { glNo: "436010", glName: "MARGIN A/C - LETTER OF CREDIT",              defaultPreset: "PRESET_3" },
  { glNo: "436020", glName: "MARGIN A/C - SHIPPING GUARANTEE ISSUED",     defaultPreset: "PRESET_3" },
  { glNo: "436080", glName: "MARGIN A/C - STAFF HOUSING LOANS",           defaultPreset: "PRESET_3" },
  { glNo: "436090", glName: "MARGIN A/C - LOANS",                         defaultPreset: "PRESET_3" },
  { glNo: "436100", glName: "MARGIN A/C - PAST DUE ADVANCES",             defaultPreset: "PRESET_3" },
  { glNo: "436110", glName: "MARGIN A/C - MISCE. GUARANTEE ISSUED",       defaultPreset: "PRESET_3" },
  { glNo: "436120", glName: "MARGIN A/C - COLLECTION",                    defaultPreset: "PRESET_3" },
  { glNo: "436160", glName: "MARGIN A/C - MISCELLANIOUS",                 defaultPreset: "PRESET_3" },
  { glNo: "436170", glName: "MARGIN A/C - USANCE BILLS",                  defaultPreset: "PRESET_3" },
  { glNo: "467010", glName: "CASHIERS EXCESS AND SHORTAGES",              defaultPreset: "PRESET_3" },
  { glNo: "467030", glName: "DRAFT PAYABLE",                              defaultPreset: "PRESET_3" },
  { glNo: "467040", glName: "PAY ORDER",                                  defaultPreset: "PRESET_3" },
  { glNo: "467070", glName: "UNCLAIMED BALANCE",                          defaultPreset: "PRESET_3" },

  /* ---- Group 4 (Page 3 - Part B and Page 4) ---- */
  { glNo: "467120", glName: "SUSPENSE A/C - CREDITORS",                   defaultPreset: "PRESET_4" },
  { glNo: "467130", glName: "SUSPENSE A/C - CREDITORS PAWNING",           defaultPreset: "PRESET_4" },
  { glNo: "467140", glName: "SUSPENSE A/C - CREDITORS (2)",               defaultPreset: "PRESET_4" },
  { glNo: "467340", glName: "SUSPENSE A/C CREDITORS PAWNING 1",           defaultPreset: "PRESET_4" },
  { glNo: "467350", glName: "SUSPENSE A/C CREDITORS PAWNING 2",           defaultPreset: "PRESET_4" },
  { glNo: "467360", glName: "SUSPENSE A/C CREDITORS PAWNING 3",           defaultPreset: "PRESET_4" },
  { glNo: "467370", glName: "SUSPENSE A/C CREDITORS PAWNING 4",           defaultPreset: "PRESET_4" },
  { glNo: "467380", glName: "SUSPENSE A/C CREDITORS PAWNING 5",           defaultPreset: "PRESET_4" },
  { glNo: "492030", glName: "SUSPENSE A/C - RENT",                        defaultPreset: "PRESET_4" },
  { glNo: "492150", glName: "SUSPENSE A/C - AMOUNT PAYBLE TO CRIBS",      defaultPreset: "PRESET_4" },
  { glNo: "492190", glName: "E- CHANNELING PAYABLE",                      defaultPreset: "PRESET_4" },
  { glNo: "492710", glName: "SUSP A/C-MONEY RETAI O/A CONTR PERFORM",     defaultPreset: "PRESET_4" },
  { glNo: "492940", glName: "PENDING DISB ETF SISU UDANA SCHOL PRJCT",    defaultPreset: "PRESET_4" },
  { glNo: "492950", glName: "PENDING DISB BFE SISU UDANA SCHOL PRJCT",    defaultPreset: "PRESET_4" },
  { glNo: "391200", glName: "Pawning Advance",                            defaultPreset: "PRESET_4" },
  { glNo: "391220", glName: "NP Pawning Advance",                         defaultPreset: "PRESET_4" }

  /* ---- Append the remaining GL accounts here (up to 60) ---- */
];


/* --------------------------------------------------------------------------
   5. LOOKUP HELPERS (read-only conveniences for app.js)
   -------------------------------------------------------------------------- */

/** glNo -> GL object, for O(1) lookups (e.g. while importing Excel files). */
window.GL_BY_NO = (function () {
  const map = {};
  window.MASTER_GL_LIST.forEach(function (gl) { map[gl.glNo] = gl; });
  return map;
})();

/** PRESET_x -> array of glNo strings. */
window.getPresetGlNumbers = function (presetId) {
  return window.MASTER_GL_LIST
    .filter(function (gl) { return gl.defaultPreset === presetId; })
    .map(function (gl) { return gl.glNo; });
};

/* Developer sanity check: warns in the console about duplicate or malformed rows. */
(function validateMasterList() {
  const seen = {};
  const validPresets = window.PRESETS.map(function (p) { return p.id; });
  window.MASTER_GL_LIST.forEach(function (gl, i) {
    if (!/^\d{6}$/.test(gl.glNo)) console.warn("[data.js] Row " + i + ": glNo should be 6 digits ->", gl.glNo);
    if (seen[gl.glNo]) console.warn("[data.js] Duplicate GL number:", gl.glNo);
    if (!gl.glName || !gl.glName.trim()) console.warn("[data.js] Missing glName for", gl.glNo);
    if (validPresets.indexOf(gl.defaultPreset) === -1) console.warn("[data.js] Unknown preset for", gl.glNo, gl.defaultPreset);
    seen[gl.glNo] = true;
  });
  if (window.MASTER_GL_LIST.length !== 60) {
    console.info("[data.js] MASTER_GL_LIST has " + window.MASTER_GL_LIST.length + " accounts (expected 60).");
  }
})();
