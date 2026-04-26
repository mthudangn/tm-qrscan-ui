// APP-SCRIPT BOUNDED BY FILE SEPARATELY /
/*************************************
 * Scan QR using normal phone camera to open UI webapp browser only
 * Then scan BARCODE continuously using UI webapp camera 
 *************************************/

const CFG = {
  // Sheets
  SHEET_PXK: "PXK",
  SHEET_LOG: "LOG",
  SHEET_QUEUE: "QUEUE",
  DOT_HEADER_ROW: 1,

  // prodKey generation (ONLY for PXK)
  PK_HEADER: "prodKey",
  BASEKEY_HEADER: "_BaseKey", // helper col (hidden)
  // Dynamic by HEADER names (order matters)
  // Edit THIS list only. Add/remove headers freely; inserting columns won't break.
  PK_SRC_HEADERS: [
    "Tên Sp", "DH"
  ],

  // PXK anchor column for prodKey header search
  COL_PRODKEY: "A",
  STEP_PREFIX: "Ngay_",
  STEPS: ["GTAM", "CAN", "IN", "CHAP", "DAN", "QC", "KHO"],

  // PXK dynamic header names you care about (edit here only)
  PXK_ID_HEADER: "PXK_ID",
  PXK_LABEL_TO: "Kính gửi:",
  PXK_LABEL_DATE: "Ngày xuất hàng:",
  PXK_H_TENSP: "Tên Sp",
  PXK_H_PRINT: "Print",
  PXK_H_CUSTOMER: "Customer",
  PXK_H_LOAI: "LOẠI",
  PXK_H_D: "D",
  PXK_H_R: "R",
  PXK_H_C: "C",
  PXK_H_SL: "SL",
  PXK_H_GIAY: "Giấy",
  PXK_H_KHO: "Khổ",
  PXK_H_DAIC: "Dài-C",
  PXK_H_SLC: "SL-C",
  PXK_H_CANLAN: "Cán lằn",
  PXK_H_DH: "DH",
  PXK_H_M: "M",
  PXK_H_M2: "M2",
  PXK_H_M3: "M3",
  PXK_H_KHGH: "KHGH",

  // PXK print/export (to send customer)
  PXK_PRINT_LAST_HEADER: "DH",      // in tới cột DH (hoặc đổi sang header khác nếu muốn)
  PXK_PRINT_FOOTER_ROWS: 16,        // số dòng dưới "nhận sản phẩm..." cần lấy luôn (chữ ký)

  // Dates
  TZ: "Asia/Ho_Chi_Minh",
  DATE_FMT: "dd/MM/yyyy",

  // TEM layout
  TEM_COLS_PER_ROW: 3,     // 3 tem / hàng
  TEM_ROWS_PER_LABEL: 7,   // line1 có QR master nhỏ + temField1
  TEM_GAP_ROWS: 1,
  TEM_ROWS_PER_PAGE: 4,    // 4 hàng / trang => 12 tem / A4
};

const UI = {
  GITHUB_UI_BASE: "https://mthudangn.github.io/tm-qrscan-ui/",
};

/*************************************
 * BATCH DOT/TEM (NO TEMPLATE)
 * - DOT_1, DOT_2... created by code (insertSheet)
 * - TEM_1, TEM_2... created by code (insertSheet)
 *************************************/
const BATCH = {
  PROP_KEY: "CURRENT_DOT_BATCH_NO",
  DOT_PREFIX: "DOT_",
  TEM_PREFIX: "TEM_",
  DEFAULT_BATCH: 1,
};

function dotName_(n){ return BATCH.DOT_PREFIX + n; }
function temName_(n){ return BATCH.TEM_PREFIX + n; }

function getCurrentBatchNo_() {
  const p = PropertiesService.getDocumentProperties();
  const n = parseInt(p.getProperty(BATCH.PROP_KEY) || String(BATCH.DEFAULT_BATCH), 10);
  return Number.isInteger(n) && n >= 1 ? n : BATCH.DEFAULT_BATCH;
}
function setCurrentBatchNo_(n) {
  n = parseInt(String(n), 10);
  if (!Number.isInteger(n) || n < 1) throw new Error("Invalid batch number");
  PropertiesService.getDocumentProperties().setProperty(BATCH.PROP_KEY, String(n));
  return n;
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/* onOpen(): adds AUTOMATE menu */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("AUTOMATE")
    .addItem("0) Setup New Deploy URL", "setupNewDeployUrl")
    .addItem("1) Prepare PXK for QR", "preparePXKforQR")
    .addItem("2) Install Queue Auto Worker", "installQueueTrigger")
    .addSeparator()
    .addItem("3) Generate DOT", "generateDotSmart")
    .addItem("4) Generate TEM", "buildTemFromDot")
    .addSeparator()
    .addItem("5) Export TEM", "exportTemPdf")
    .addItem("6) Export PXK", "exportPxkPdf")
    .addToUi();
}

/*************************************
 * A) PXK data boundary (stop before form zone)
 *************************************/
function findPXKHeaderPos_(shPXK) {
  const pos = findCellByText_(shPXK, [CFG.PXK_H_TENSP, "Ten Sp"]);
  if (!pos) throw new Error("PXK: cannot find header 'Tên Sp'.");
  return pos; // {row, col}
}

function findPXKDataEndRow_(shPXK, headerRow, nameCol1Based) {
  const lastRow = shPXK.getLastRow();
  if (lastRow <= headerRow) return headerRow;

  // đọc 1 block: cột A + 5 cột đầu (để bắt text bị merge/đặt lệch cột)
  const n = lastRow - headerRow;
  const maxScanCols = Math.min(8, shPXK.getLastColumn()); // A..H là đủ bắt "nhận sản phẩm"
  const block = shPXK.getRange(headerRow + 1, 1, n, maxScanCols).getDisplayValues();

  const normStop_ = (s) => String(s || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // bỏ dấu
    .replace(/đ/g, "d")
    .replace(/\s+/g, " ");

  // match được cả có dấu & không dấu, và cả dòng bắt đầu bằng ** (kể cả có space)
  const isStopLine_ = (s) => {
    const t = normStop_(s);
    return (
      t.includes("nhan san pham") ||
      t.startsWith("**") ||
      t.startsWith("* *") ||          // phòng trường hợp "**" bị tách
      t.includes("nhan hang") ||      // optional: nếu bạn hay dùng chữ khác
      t.includes("cam on")            // optional
    );
  };

  let end = headerRow;
  let blankRun = 0;

  for (let i = 0; i < n; i++) {
    const r = headerRow + 1 + i;

    // 1) STOP detection: join vài cột đầu để bắt merge/đặt lệch cột
    const joined = block[i].slice(0, 8).join(" ");
    if (isStopLine_(joined)) break;

    // 2) blank-run logic: vẫn dựa vào cột Tên SP (nếu có)
    const nameVal = (nameCol1Based >= 1 && nameCol1Based <= maxScanCols)
      ? String(block[i][nameCol1Based - 1] || "").trim()
      : ""; // fallback nếu nameCol ngoài vùng scan

    if (!nameVal) blankRun++;
    else blankRun = 0;

    if (blankRun >= 20) break;
    end = r;
  }

  return end;
}

/*************************************
 * WEBAPP URL GUARDS (hard stop + /dev warning)
 *************************************/
function assertWebAppExecUrlOrThrow_() {
  const p = PropertiesService.getScriptProperties();
  const saved = String(p.getProperty(WEBAPP.PROP_EXEC_URL) || "").trim();

  if (!saved) {
    throw new Error(
      "CHƯA SET BACKEND URL.\n\n" +
      "👉 Vào menu: AUTOMATE → 0) Setup New Deploy URL\n" +
      "và dán Cloud Run backend URL dạng:\n" +
      "https://tm-barcode-api-1060743715550.asia-southeast1.run.app"
    );
  }

  if (!/^https?:\/\//i.test(saved)) {
    throw new Error(
      "BACKEND URL không hợp lệ.\n\n" +
      "URL hiện tại:\n" + saved + "\n\n" +
      "Backend URL phải bắt đầu bằng https:// hoặc http://"
    );
  }

  return true;
}

/*************************************
 * Utilities: sheet, headers, logging, find
 *************************************/
function mustGetSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Missing sheet: ${name}`);
  return sh;
}
function getHeader_(sh, headerRow = 1) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error(`Sheet '${sh.getName()}' has no columns.`);
  return sh.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || "").trim());
}
function indexHeaders_(headerRow) {
  const idx = {};
  headerRow.forEach((h, i) => { if (h) idx[h] = i; });
  return idx;
}
function index_(h){
  const o = {};
  (h || []).forEach((x,i)=>{ if (x !== "" && x != null) o[String(x).trim()] = i; });
  return o;
}

// Normalize header: trim, lower, remove diacritics, collapse spaces, đ->d
function normHeader_(x) {
  return String(x || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/\s+/g, " ");
}

function makeHeaderIndexMap_(headers) {
  const m = {};
  headers.forEach((h, i) => {
    const raw = String(h || "").trim();
    if (!raw) return;
    const col1 = i + 1;
    if (m[raw] == null) m[raw] = col1;
    const nk = normHeader_(raw);
    if (nk && m[nk] == null) m[nk] = col1;
  });
  return m;
}

function getAvailableStepsFromPXK_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  const shPXK = ss.getSheetByName(CFG.SHEET_PXK);
  if (!shPXK) return [];

  const pxkStart = findStartRowAfterHeader_(shPXK, CFG.PK_HEADER, CFG.COL_PRODKEY);
  if (pxkStart === -1) return [];
  const hr = pxkStart - 1;

  const lastCol = shPXK.getLastColumn();
  const headers = shPXK.getRange(hr, 1, 1, lastCol).getDisplayValues()[0]
    .map(v => String(v || "").trim());
  const pCol = makeHeaderIndexMap_(headers);

  const steps = (CFG.STEPS && CFG.STEPS.length) ? CFG.STEPS : [];
  return steps.filter(s => {
    const colName = `${CFG.STEP_PREFIX}${String(s).trim().toUpperCase()}`;
    return !!findFirstHeaderCol_(pCol, [colName]);
  });
}

function ensureStepColumnsInPXK_(shPXK, headerRow, steps) {
  const stepsU = (steps || [])
    .map(s => String(s).trim().toUpperCase())
    .filter(Boolean);
  if (!stepsU.length) stepsU.push(); // no-op

  const desiredSteps = stepsU.map(s => `${CFG.STEP_PREFIX}${s}`); // Ngay_GTAM...

  // 0) Read current headers
  const readHeaders_ = () => {
    const lastCol = Math.max(1, shPXK.getLastColumn());
    const headers = shPXK.getRange(headerRow, 1, 1, lastCol)
      .getDisplayValues()[0]
      .map(v => String(v || "").trim());
    return { lastCol, headers, hMap: makeHeaderIndexMap_(headers) };
  };

  let { headers, hMap } = readHeaders_();

  // 1) Anchor: Customer must exist
  const customerCol1 =
    findFirstHeaderCol_(hMap, [CFG.PXK_H_CUSTOMER, "Customer", "Khách hàng", "Khach hang", "Client", "KH"]) || 0;
  if (!customerCol1) {
    throw new Error("PXK: cannot find header 'Customer' to place Ngay_* columns after it.");
  }

  // 2) Ensure missing Ngay_* exist (append to end first, then we will move them)
  if (desiredSteps.length) {
    const have = new Set(headers);
    const missing = desiredSteps.filter(h => !have.has(h));
    if (missing.length) {
      const curLast = shPXK.getLastColumn();
      shPXK.insertColumnsAfter(curLast, missing.length);
      shPXK.getRange(headerRow, curLast + 1, 1, missing.length).setValues([missing]);

      // date format for these appended columns
      shPXK.getRange(
        headerRow + 1,
        curLast + 1,
        Math.max(1, shPXK.getMaxRows() - headerRow),
        missing.length
      ).setNumberFormat(CFG.DATE_FMT);

      ({ headers, hMap } = readHeaders_());
    }

    // 3) Move Ngay_* to be contiguous right after Customer, in desired order
    // target position for first step column
    let insertPos = customerCol1 + 1;

    const colOf_ = (name) => findFirstHeaderCol_(hMap, [name]);

    for (const stepHeader of desiredSteps) {
      let cur = colOf_(stepHeader);
      if (!cur) continue;

      // already in place
      if (cur === insertPos) {
        insertPos++;
        continue;
      }

      // move whole column to insertPos
      shPXK.moveColumns(shPXK.getRange(1, cur, shPXK.getMaxRows(), 1), insertPos);

      // refresh header map after each move (indices change)
      ({ headers, hMap } = readHeaders_());
      insertPos++;
    }

    // 4) Ensure date format for final cluster (safety)
    ({ headers, hMap } = readHeaders_());
    const firstStepCol = desiredSteps.length ? findFirstHeaderCol_(hMap, [desiredSteps[0]]) : 0;
    if (firstStepCol && desiredSteps.length) {
      shPXK.getRange(
        headerRow + 1,
        firstStepCol,
        Math.max(1, shPXK.getMaxRows() - headerRow),
        desiredSteps.length
      ).setNumberFormat(CFG.DATE_FMT);
    }
  }

  // 5) Ensure _BaseKey exists AND is rightmost; hide it
  // If missing -> append at end
  ({ headers, hMap } = readHeaders_());
  let baseCol = findFirstHeaderCol_(hMap, [CFG.BASEKEY_HEADER]);

  if (!baseCol) {
    const curLast = shPXK.getLastColumn();
    shPXK.insertColumnsAfter(curLast, 1);
    shPXK.getRange(headerRow, curLast + 1).setValue(CFG.BASEKEY_HEADER);

    ({ headers, hMap } = readHeaders_());
    baseCol = findFirstHeaderCol_(hMap, [CFG.BASEKEY_HEADER]);
  }

  // Move _BaseKey to the far right if not already rightmost
  const lastColNow = shPXK.getLastColumn();
  if (baseCol && baseCol !== lastColNow) {
    shPXK.moveColumns(shPXK.getRange(1, baseCol, shPXK.getMaxRows(), 1), lastColNow + 1);
    baseCol = shPXK.getLastColumn(); // now rightmost
  }

  // Hide _BaseKey column (safe)
  if (baseCol) {
    try { shPXK.hideColumn(shPXK.getRange(1, baseCol)); } catch (e) {}
  }
}

function findFirstHeaderCol_(headerMap, candidates) {
  for (const h of candidates) {
    const raw = String(h || "").trim();
    if (!raw) continue;
    if (headerMap[raw]) return headerMap[raw];
    const nk = normHeader_(raw);
    if (nk && headerMap[nk]) return headerMap[nk];
  }
  return 0;
}

function mustFindHeader_(headerMap, candidates, sheetNameForError) {
  const c = findFirstHeaderCol_(headerMap, candidates);
  if (!c) throw new Error(`${sheetNameForError} missing header: ${candidates.join(" / ")}`);
  return c;
}

/** Convert column letter to index (A->1, B->2...) */
function colLetterToIndex_(col) {
  const s = String(col || "").trim().toUpperCase();
  if (!s) throw new Error("Invalid column letter");
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 65 || c > 90) throw new Error(`Invalid column letter: ${col}`);
    n = n * 26 + (c - 64);
  }
  return n;
}

function colIndexToLetter_(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function norm_(x) { return String(x || "").trim().toLowerCase(); }

/**
 * Find start row after a header text in a given column letter.
 * Returns headerRow + 1, or -1 if not found.
 */
function findStartRowAfterHeader_(sh, headerText, colLetter) {
  const col = colLetterToIndex_(colLetter);
  const lastRow = sh.getLastRow();
  if (lastRow < 1) return -1;

  const vals = sh.getRange(1, col, lastRow, 1).getDisplayValues();
  const target = String(headerText || "").trim().toLowerCase();

  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i][0] || "").trim().toLowerCase();
    if (v === target) return i + 2; // headerRow + 1
  }
  return -1;
}


function findCellByText_(sh, texts) {
  const maxRows = Math.min(sh.getLastRow(), 250);
  const maxCols = Math.min(sh.getLastColumn(), 150);
  if (maxRows < 1 || maxCols < 1) return null;

  const vals = sh.getRange(1, 1, maxRows, maxCols).getDisplayValues();
  const targets = texts.map(t => norm_(t));

  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[0].length; c++) {
      if (targets.includes(norm_(vals[r][c]))) return { row: r + 1, col: c + 1 };
    }
  }
  return null;
}

function findCellContainsText_(sh, texts) {
  const maxRows = Math.min(sh.getLastRow(), 50);
  const maxCols = Math.min(sh.getLastColumn(), 20);
  if (maxRows < 1 || maxCols < 1) return null;

  const vals = sh.getRange(1, 1, maxRows, maxCols).getDisplayValues();
  const targets = texts.map(t => norm_(t));

  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[0].length; c++) {
      const cell = norm_(vals[r][c]);
      if (!cell) continue;

      for (const t of targets) {
        if (cell.includes(t)) return { row: r + 1, col: c + 1 };
      }
    }
  }
  return null;
}

function getOrCreateLog_(ss){
  let sh = ss.getSheetByName(CFG.SHEET_LOG);
  if (!sh) sh = ss.insertSheet(CFG.SHEET_LOG);

  // move LOG to desired position (left side) immediately
  enforceSheetOrder_(ss);

  // keep it simple + stable
  const needHeader = sh.getLastRow() < 1 || String(sh.getRange(1,1).getDisplayValue() || "").trim() !== "TS";
  if (needHeader) {
    sh.clearContents();
    sh.appendRow(["TS","prodKey","Step","Worker","Status","Note"]);
  }
  return sh;
}

function getOrCreateQueue_(ss){
  let sh = ss.getSheetByName(CFG.SHEET_QUEUE);
  if (!sh) sh = ss.insertSheet(CFG.SHEET_QUEUE);
  enforceSheetOrder_(ss);

  const needHeader =
    sh.getLastRow() < 1 ||
    String(sh.getRange(1,1).getDisplayValue() || "").trim() !== "TS";

  if (needHeader) {
    sh.clearContents();
    sh.appendRow([
      "TS","queueId","prodKey","step","worker","pxk",
      "status","note","doneBadgesJson","updatedAt"
    ]);
  }
  return sh;
}

function appendLog_(sh, pk, step, worker, status, note){
  const st = String(status || "").trim().toUpperCase();
  const p  = String(pk || "").trim();

  // Chỉ log OK. (Giữ SYSTEM/INFO nếu muốn trace hệ thống)
  if (st !== "OK" && !(p === "SYSTEM" && st === "INFO")) return;

  sh.appendRow([
    Utilities.formatDate(new Date(), CFG.TZ, "yyyy-MM-dd HH:mm:ss"),
    pk || "", step || "", worker || "", st || "", note || ""
  ]);
}

/*************************************
 * Numeric helpers
 *************************************/
function ceilFromDisplay_(v) {
  const s = String(v ?? "").trim();
  if (!s) return "";
  const x = parseFloat(s.replace(",", "."));
  if (isNaN(x)) return "";
  return String(Math.ceil(x));
}
function to2dec_(v) {
  if (v === "" || v == null) return "";
  const x = (typeof v === "number") ? v : parseFloat(String(v).replace(",", "."));
  if (isNaN(x)) return "";
  return Number(x.toFixed(2));
}

/*************************************
 * BACKEND URL (stable)
 * - Store Cloud Run backend URL in Script Properties
 *************************************/
const WEBAPP = {
  PROP_EXEC_URL: "WEBAPP_EXEC_URL" // keep same key to avoid breaking existing properties
};

/** Normalize backend URL: trim, remove query/fragment, no /exec forcing */
function normalizeWebAppBaseUrl_(url) {
  url = String(url || "").trim();

  // remove query/fragment if user pasted full link with ?...
  url = url.split("#")[0];
  url = url.split("?")[0];

  // Cloud Run / Apps Script compatible:
  // DO NOT force /exec anymore.
  url = url.replace(/\/+$/,"");

  return url;
}

/** Get stable backend URL from Script Properties */
function getWebAppBaseUrl_() {
  const p = PropertiesService.getScriptProperties();
  const saved = String(p.getProperty(WEBAPP.PROP_EXEC_URL) || "").trim();
  if (saved) return normalizeWebAppBaseUrl_(saved);

  const auto = ScriptApp.getService().getUrl(); // fallback only, usually unused after Cloud Run
  return auto ? normalizeWebAppBaseUrl_(auto) : "";
}

function buildOpenCameraUrl_() {
  const backendUrl = getWebAppBaseUrl_();
  if (!backendUrl) return "";

  const api = encodeURIComponent(normalizeWebAppBaseUrl_(backendUrl));
  const sid = encodeURIComponent(SpreadsheetApp.getActive().getId());

  return UI.GITHUB_UI_BASE + "?api=" + api + "&mode=scan&v=1&sid=" + sid;
}

function buildMasterScanUrl_(shPXK) {
  const backendUrl = getWebAppBaseUrl_();
  if (!backendUrl) return "";

  const pxkId = getPxkIdFromSheet_(shPXK);
  const info = getPxkHeaderInfo_(shPXK);

  const api = encodeURIComponent(normalizeWebAppBaseUrl_(backendUrl));
  const sid = encodeURIComponent(shPXK.getParent().getId());
  const co  = encodeURIComponent(String(info.toRaw || "").trim());
  const dt  = encodeURIComponent(formatPxkShortDate_(info.dateDigits));

  return `${UI.GITHUB_UI_BASE}?api=${api}&mode=scan&v=1&sid=${sid}&pxk=${encodeURIComponent(pxkId)}&co=${co}&dt=${dt}`;
}

function formatPxkShortDate_(ddmmyyyy) {
  const s = String(ddmmyyyy || "").trim();
  if (!/^\d{8}$/.test(s)) return "";
  const dd = s.slice(0, 2);
  const mm = s.slice(2, 4);
  return `${dd}/${mm}`;
}

function setupNewDeployUrl() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const p = PropertiesService.getScriptProperties();

  const cur = String(p.getProperty(WEBAPP.PROP_EXEC_URL) || "").trim();

  const resp = ui.prompt(
    "Setup Backend URL for QR Scan UI",
    "Dán Cloud Run backend URL mới.\n\n" +
    "Ví dụ:\n" +
    "https://tm-barcode-api-1060743715550.asia-southeast1.run.app\n\n" +
    "Nếu vẫn dùng Apps Script tạm thời thì vẫn có thể dán link /exec.\n\n" +
    (cur ? ("URL hiện tại: " + cur + "\n\n") : "") +
    "✏️ Backend URL:",
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  let url = String(resp.getResponseText() || "").trim();

  if (!url) {
    ui.alert("❌ Empty URL! Hãy dán Cloud Run backend URL.");
    return;
  }

  url = normalizeWebAppBaseUrl_(url);

  if (!/^https?:\/\//i.test(url)) {
    ui.alert("❌ Sai URL! Backend URL phải bắt đầu bằng https:// hoặc http://");
    return;
  }

  p.setProperty(WEBAPP.PROP_EXEC_URL, url);

  const openCam = buildOpenCameraUrl_();
  const shLog = getOrCreateLog_(ss);

  const note =
    `---Backend--- ${url}\n` +
    `---GitHub--- ${openCam}`;

  appendLog_(shLog, "SYSTEM", "SETUP_BACKEND_URL", "", "INFO", note);

  ui.alert(
    "✅ DONE Setup Backend URL!\n\n" +
    "1) Backend URL:\n" + url + "\n\n" +
    "2) GitHub UI:\n" + openCam + "\n\n" +
    "QR generated from this file will include sid = this PXK master file ID."
  );
}

/*************************************
 * Row lookup
 *************************************/
function findRowIndexByTwoValues_(sh, colA0, valA, colB0, valB, startRow = 2) {
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return -1;

  const a = String(valA || "").trim();
  const b = String(valB || "").trim();

  const vals = sh.getRange(startRow, 1, lastRow - startRow + 1, sh.getLastColumn()).getDisplayValues();
  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    if (
      String(row[colA0] || "").trim() === a &&
      String(row[colB0] || "").trim() === b
    ) {
      return startRow + i;
    }
  }
  return -1;
}

/*************************************
 * 1) Prepare PXK for QR (prodKey + Ngay_STEP)
 *************************************/
function preparePXKforQR() {
  const ss = SpreadsheetApp.getActive();
  const shPXK = mustGetSheet_(ss, CFG.SHEET_PXK);
  const shLog = getOrCreateLog_(ss);

  const pxkId = writePxkIdHeader_(shPXK);   // luôn regenerate theo Kính gửi + Ngày xuất hàng
  const masterUrl = writeMasterQrToPXK_(shPXK);
  generateprodKeyPXK_(shPXK);

  appendLog_(
    shLog,
    "SYSTEM",
    "PREPARE_PXK",
    "",
    "OK",
    `PXK_ID = ${pxkId} - MASTER_QR = ${masterUrl || ""}`
  );

  SpreadsheetApp.getUi().alert(
    `✅ PXK_ID, prodKey and Ngay_STEP generated successfully!\n\nPXK_ID: ${pxkId}`
  );
}

function normalizeAlphaNum_(v) {
  const raw = String(v || "").trim();
  if (!raw) return "";

  const clean = raw
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/Đ/g, "D")
    .replace(/đ/g, "D");

  let words = clean.split(/[^A-Z0-9]+/).filter(Boolean);

  // bỏ các từ generic
  const STOP = new Set(["CTY", "CONGTY", "CONG", "TY", "COMPANY"]);
  words = words.filter(w => !STOP.has(w));

  if (!words.length) return "";

  // nhiều từ => lấy chữ cái đầu
  if (words.length > 1) {
    return words.map(w => w[0]).join("");
  }

  // 1 từ => giữ nguyên
  return words[0];
}

function extractPxkDateDigits_(s) {
  const t = String(s || "").trim();

  // bắt dd/MM/yyyy hoặc d/M/yyyy
  const m = t.match(/(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})/);
  if (!m) return "";

  const dd = m[1].padStart(2, "0");
  const mm = m[2].padStart(2, "0");
  const yyyy = m[3];

  return `${dd}${mm}${yyyy}`;
}

function getPxkHeaderInfo_(shPXK) {
  const toPos = findCellContainsText_(shPXK, ["Kính gửi", "Kinh gui"]);
  if (!toPos) throw new Error("PXK: Không tìm thấy dòng 'Kính gửi'.");

  // ô của bạn đang là: "Kính gửi: CTY MINH NGỌC" cùng 1 ô
  const toCellText = String(shPXK.getRange(toPos.row, toPos.col).getDisplayValue() || "").trim();
  let toVal = toCellText.replace(/^\s*Kính gửi\s*:\s*/i, "").trim();

  // fallback nếu value nằm ô bên phải
  if (!toVal) {
    toVal = String(shPXK.getRange(toPos.row, toPos.col + 1).getDisplayValue() || "").trim();
  }

  const datePos = findCellContainsText_(shPXK, ["Ngày xuất hàng", "Ngay xuat hang"]);
  if (!datePos) throw new Error("PXK: cannot find 'Ngày xuất hàng' line.");

  // ô của bạn đang là: "Ngày xuất hàng: 09/04/2024" cùng 1 ô
  const dateCellText = String(shPXK.getRange(datePos.row, datePos.col).getDisplayValue() || "").trim();
  const dateDigits = extractPxkDateDigits_(dateCellText);

  if (!toVal) throw new Error("PXK: 'Kính gửi' đang trống tên khách hàng.");
  if (!dateDigits) throw new Error("PXK: cannot parse export date from 'Ngày xuất hàng'.");

  return {
    toRaw: toVal,
    toCode: normalizeAlphaNum_(toVal),
    dateDigits: dateDigits
  };
}

function buildPxkId_(shPXK) {
  const info = getPxkHeaderInfo_(shPXK);
  return `${info.toCode}${info.dateDigits}`;
}

function writePxkIdHeader_(shPXK) {
  const pxkId = buildPxkId_(shPXK);

  // đặt tại A5:A6 (bên trái Kính gửi / Ngày xuất hàng)
  const labelCell = shPXK.getRange("A5");
  const valueCell = shPXK.getRange("A6");

  // set text
  labelCell.setValue(CFG.PXK_ID_HEADER);
  valueCell.setValue(pxkId);
  valueCell.setNumberFormat("@");

  // style cho rõ
  const styleRange = shPXK.getRange("A5:A6");

  styleRange
    .setFontFamily("Arial")
    .setFontSize(14)
    .setFontWeight("bold")
    .setFontColor("#FF0000")
    .setBackground("#FFF200") // yellow highlight for PXK_ID
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  return pxkId;
}

function writeMasterQrToPXK_(shPXK) {
  const masterUrl = buildMasterScanUrl_(shPXK);
  if (!masterUrl) return "";

  // Vị trí gợi ý: C5:D10, nằm cạnh PXK_ID
  const titleCell = shPXK.getRange("C5");
  const qrCell = shPXK.getRange("C6");
  const linkCell = shPXK.getRange("C10");

  titleCell.setValue("QR_MASTER");
  titleCell
    .setFontFamily("Arial")
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#D9EAD3");

  const qrImgUrl =
    "https://api.qrserver.com/v1/create-qr-code/?size=220x220&data=" +
    encodeURIComponent(masterUrl);

  qrCell.setFormula(`=IMAGE("${qrImgUrl}")`);

  // link text để văn phòng copy nếu cần
  linkCell.setValue(masterUrl);
  linkCell.setWrap(true);
  linkCell.setFontSize(8);

  return masterUrl;
}

function getPxkIdFromSheet_(shPXK) {
  const header = String(shPXK.getRange("A5").getDisplayValue() || "").trim();
  const value = String(shPXK.getRange("A6").getDisplayValue() || "").trim();

  if (header === CFG.PXK_ID_HEADER && value) return value;

  // fallback: rebuild nếu chưa có
  return writePxkIdHeader_(shPXK);
}

function normalizePxkId_(v) {
  return String(v || "").trim().toUpperCase();
}

function getPxkOpenStatus_(ss, pxkId) {
  const shPXK = mustGetSheet_(ss, CFG.SHEET_PXK);
  const want = normalizePxkId_(pxkId);
  if (!want) {
    return { ok: false, code: "BAD_INPUT", msg: "Missing pxk" };
  }

  // PXK hiện tại của file này nằm ở A6, không phải trong bảng line-item
  const currentPxk = normalizePxkId_(getPxkIdFromSheet_(shPXK));
  if (!currentPxk) {
    return { ok: false, code: "PXK_ID_MISSING", msg: "PXK_ID chưa được tạo ở A6" };
  }

  if (currentPxk !== want) {
    return { ok: false, code: "PXK_NOT_FOUND", msg: "PXK không tồn tại", found: 0, openCount: 0 };
  }

  const pxkStart = findStartRowAfterHeader_(shPXK, CFG.PK_HEADER, CFG.COL_PRODKEY);
  if (pxkStart === -1) {
    return { ok: false, code: "PXK_HEADER_NOT_FOUND", msg: "PXK: cannot find prodKey header." };
  }

  const hr = pxkStart - 1;
  const lastCol = shPXK.getLastColumn();
  const headers = shPXK.getRange(hr, 1, 1, lastCol).getDisplayValues()[0]
    .map(v => String(v || "").trim());
  const hMap = makeHeaderIndexMap_(headers);

  const khoCol1 = findFirstHeaderCol_(hMap, [CFG.STEP_PREFIX + "KHO"]);
  if (!khoCol1) {
    return { ok: false, code: "KHO_COL_MISSING", msg: "PXK missing column 'Ngay_KHO'" };
  }

  const pos = findPXKHeaderPos_(shPXK);
  const pxkEnd = findPXKDataEndRow_(shPXK, pos.row, pos.col);
  if (pxkEnd < pxkStart) {
    return { ok: false, code: "NO_DATA", msg: "PXK has no data rows" };
  }

  const n = pxkEnd - pxkStart + 1;
  const vals = shPXK.getRange(pxkStart, 1, n, lastCol).getDisplayValues();

  let found = 0;
  let openCount = 0;

  for (const row of vals) {
    const pk = String(row[0] || "").trim(); // cột A = prodKey
    if (!pk) continue;

    found++;
    const khoVal = String(row[khoCol1 - 1] || "").trim();
    if (!khoVal) openCount++;
  }

  if (!found) {
    return { ok: false, code: "PXK_NOT_FOUND", msg: "PXK không tồn tại", found: 0, openCount: 0 };
  }

  if (openCount <= 0) {
    return { ok: false, code: "PXK_CLOSED", msg: "PXK này đã hoàn tất nhập kho", found, openCount: 0 };
  }

  return {
    ok: true,
    code: "OPEN",
    msg: "PXK này đang hoạt động",
    found,
    openCount
  };
}

function normalizePkText_(v) {
  return String(v || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/Đ/g, "D")
    .replace(/đ/g, "D")
    .replace(/[^A-Z0-9]/g, "");
}

function buildProdKeyBase_(pxkId, tenSp, dh) {
  const a = String(pxkId || "").trim().toUpperCase();
  const b = normalizePkText_(tenSp).slice(0, 10);
  const c = String(dh || "").trim();

  if (!a || !b || !c) return "";
  return `${a}|${b}|${c}`;
}

function generateprodKeyPXK_(shPXK) {
  if (shPXK.getName() !== CFG.SHEET_PXK) throw new Error("prodKey generator is PXK-only.");

  const pos = findPXKHeaderPos_(shPXK);
  const hr = pos.row;
  const nameCol = pos.col;

  ensureStepColumnsInPXK_(shPXK, hr, CFG.STEPS);

  const pkCol = nameCol - 1;
  if (pkCol < 1) throw new Error("PXK: cannot place prodKey left of 'Tên Sp'.");

  if (String(shPXK.getRange(hr, pkCol).getDisplayValue() || "").trim() !== CFG.PK_HEADER) {
    shPXK.getRange(hr, pkCol).setValue(CFG.PK_HEADER);
  }

  const startRow = hr + 1;
  const endRow = findPXKDataEndRow_(shPXK, hr, nameCol);
  if (endRow < startRow) return;

  const lastCol0 = Math.max(1, shPXK.getLastColumn());
  const headers = shPXK.getRange(hr, 1, 1, lastCol0).getDisplayValues()[0]
    .map(x => String(x || "").trim());
  const hMap = makeHeaderIndexMap_(headers);

  const baseCol = findFirstHeaderCol_(hMap, [CFG.BASEKEY_HEADER]);
  if (!baseCol) throw new Error("PXK: missing _BaseKey (ensureStepColumnsInPXK_ failed).");

  const srcCols1 = [];
  for (const h of CFG.PK_SRC_HEADERS) {
    const col1 = findFirstHeaderCol_(hMap, [h]);
    if (!col1) throw new Error(`PXK missing header for prodKey source: '${h}'`);
    srcCols1.push(col1);
  }

  const colTenSp = srcCols1[0];
  const colDH    = srcCols1[1];

  const pxkId = getPxkIdFromSheet_(shPXK);

  const n = endRow - startRow + 1;
  const vals = shPXK.getRange(startRow, 1, n, lastCol0).getDisplayValues();

  const baseOut = [];
  const pkOut = [];
  const seen = Object.create(null);

  for (let i = 0; i < n; i++) {
    const row = vals[i];

    const tenSp = row[colTenSp - 1];
    const dh = row[colDH - 1];

    const baseKey = buildProdKeyBase_(pxkId, tenSp, dh);
    baseOut.push([baseKey]);

    if (!baseKey) {
      pkOut.push([""]);
      continue;
    }

    seen[baseKey] = (seen[baseKey] || 0) + 1;
    const count = seen[baseKey];

    pkOut.push([count === 1 ? baseKey : `${baseKey}#${count - 1}`]);
  }

  shPXK.getRange(startRow, baseCol, n, 1).clearContent().setValues(baseOut);
  shPXK.getRange(startRow, pkCol, n, 1).clearContent().setValues(pkOut);

  shPXK.getRange(startRow, baseCol, n, 1).setNumberFormat("@");
  shPXK.getRange(startRow, pkCol, n, 1).setNumberFormat("@");

  shPXK.hideColumn(shPXK.getRange(1, baseCol));
}

/*************************************
 * DOT template
 *************************************/
const DOT_HEADERS = [
  "temSTT", "PXK_ID", "prodKey",
  "temField1", "temField2", "temField3", "temField4", "temField5",
  "Tên Sp", "Print", "LOẠI", "D", "R", "C", "SL", "Giấy",
  "Khổ", "Dài-C", "SL-C", "Cán lằn", "DH", "Customer", "M2", "M3", "M"
];
function ensureDotTemplate_(shDot) {
  const hr = CFG.DOT_HEADER_ROW;

  // If sheet is empty -> write headers fresh
  const empty =
    shDot.getLastRow() < hr ||
    shDot.getLastColumn() < 1 ||
    !String(shDot.getRange(hr, 1).getDisplayValue() || "").trim();

  if (empty) {
    shDot.clearContents();
    shDot.getRange(hr, 1, 1, DOT_HEADERS.length).setValues([DOT_HEADERS]);
    shDot.getRange(hr, 1, 1, DOT_HEADERS.length).setFontWeight("bold");
    shDot.setFrozenRows(1);
    return;
  }

  // Otherwise: migrate/reorder existing columns to match DOT_HEADERS
  const lastCol0 = shDot.getLastColumn();

  const readHeaders_ = () =>
    shDot.getRange(hr, 1, 1, shDot.getLastColumn())
      .getDisplayValues()[0]
      .map(v => String(v || "").trim());

  let headers = readHeaders_();
  let idx = indexHeaders_(headers); // 0-based index

  // Append any missing required headers (safe)
  const missing = DOT_HEADERS.filter(h => idx[h] == null);
  if (missing.length) {
    const curLast = shDot.getLastColumn();
    shDot.insertColumnsAfter(curLast, missing.length);
    shDot.getRange(hr, curLast + 1, 1, missing.length).setValues([missing]);
    shDot.getRange(hr, curLast + 1, 1, missing.length).setFontWeight("bold");
    headers = readHeaders_();
    idx = indexHeaders_(headers);
  }

  // Reorder: move each desired header column into the exact position 1..DOT_HEADERS.length
  // Extras not in DOT_HEADERS will remain on the right
  for (let pos1 = 1; pos1 <= DOT_HEADERS.length; pos1++) {
    const h = DOT_HEADERS[pos1 - 1];

    headers = readHeaders_();
    idx = indexHeaders_(headers);

    const curPos1 = (idx[h] != null) ? (idx[h] + 1) : 0;
    if (!curPos1) continue;

    if (curPos1 !== pos1) {
      shDot.moveColumns(shDot.getRange(1, curPos1, shDot.getMaxRows(), 1), pos1);
    }
  }

  // Ensure header row styling on final layout
  shDot.getRange(hr, 1, 1, DOT_HEADERS.length).setFontWeight("bold");
  shDot.setFrozenRows(1);
}

/*************************************
 * 3) Generate DOT
 *************************************/
function generateDotSmart() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  try {
    assertWebAppExecUrlOrThrow_();
  } catch (e) {
    ui.alert("❌ Generate DOT bị chặn\n\n" + e.message);
    return;
  }

  const batches = getAllDotBatches_(ss);
  const nextDefault = (batches.length ? batches[batches.length - 1] : 0) + 1;

  const resp = ui.prompt(
    "Generate DOT (số ĐỢT xuất TEM)",
    `Để trống & bấm OK để tạo sheet mới tiếp theo DOT_${nextDefault}, hoặc;
    Nhập lại số sheet DOT có sẵn để update (VD: Nhập số 1 để tạo DOT_1)
    Lưu ý: Nếu nhập số DOT cũ thì số sheet TEM tương ứng cũng cần generate lại STT\n
    ✏️ Bạn muốn tạo sheet DOT số:`,
      ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const raw = String(resp.getResponseText() || "").trim();
  const targetNo = raw ? parseInt(raw, 10) : nextDefault;
  if (!Number.isInteger(targetNo) || targetNo < 1) {
    ui.alert("❌ Số sheet DOT không hợp lệ.");
    return;
  }

  // Check existed BEFORE creating
  const targetName = dotName_(targetNo);
  const existedBefore = !!ss.getSheetByName(targetName);

  // Ensure DOT sheet
  const shDot = getBatchDotSheet_(ss, targetNo);
  ensureDotTemplate_(shDot);

  // Existed keys = ALL keys đã có ở BẤT KỲ DOT_* nào
  // => re-generate DOT_2 sẽ KHÔNG bao giờ kéo lại key của DOT_1
  const existedPK = collectAllDotprodKeys_(ss);

  // Build new rows from PXK that are NOT in existedPK
  const rows = buildDotRowsFromPXKDelta_(ss, existedPK);

  if (rows.length) {
    const writeRow = shDot.getLastRow() + 1;
    shDot.getRange(writeRow, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Always re-sort + re-number (this is the "generate again" behavior)
  sortDOT_(shDot);
  fillDotSttAfterSort_(shDot);  // step cuối: temSTT = 1..n
  applyDotFormats_(shDot);
  applyDotBorders_(shDot, CFG.DOT_HEADER_ROW, shDot.getLastRow(), shDot.getLastColumn());

  setCurrentBatchNo_(targetNo);

  ui.alert(
    `✅ ${shDot.getName()} generated successfully\n\n Rows: ${rows.length} \n`
  );
}

function dotUniqueKey_(pxkId, prodKey) {
  return `${String(pxkId || "").trim()}||${String(prodKey || "").trim()}`;
}

function buildQrKey_(pxkId, prodKey) {
  return String(prodKey || "").trim();
}

function parseQrKey_(raw) {
  const s = String(raw || "").trim();
  if (!s) return { pxkId: "", prodKey: "" };

  const pos = s.indexOf("|");
  if (pos === -1) {
    return { pxkId: "", prodKey: s };
  }

  return {
    pxkId: s.slice(0, pos).trim(),
    prodKey: s
  };
}

function collectprodKeys_(shDot) {
  const hr = CFG.DOT_HEADER_ROW;
  const headers = getHeader_(shDot, hr);
  const idx = indexHeaders_(headers);

  if (idx["PXK_ID"] == null || idx["prodKey"] == null) return new Set();

  const n = shDot.getLastRow() - hr;
  if (n <= 0) return new Set();

  const vals = shDot.getRange(hr + 1, 1, n, shDot.getLastColumn()).getDisplayValues();
  const out = new Set();

  for (const r of vals) {
    const pxkId = String(r[idx["PXK_ID"]] || "").trim();
    const pk = String(r[idx["prodKey"]] || "").trim();
    if (!pxkId || !pk) continue;
    out.add(dotUniqueKey_(pxkId, pk));
  }
  return out;
}

function collectAllDotprodKeys_(ss) {
  const set = new Set();
  const batches = getAllDotBatches_(ss);
  for (const n of batches) {
    const sh = ss.getSheetByName(dotName_(n));
    if (!sh) continue;
    collectprodKeys_(sh).forEach(k => set.add(k));
  }
  return set;
}

/*************************************
 * Build DOT rows from PXK delta (only rows whose prodKey not in existedPK)
 * - Reads PXK once (displayValues + values)
 * - Builds rows following DOT template headers
 *************************************/
function buildDotRowsFromPXKDelta_(ss, existedPK) {
  const shPXK = mustGetSheet_(ss, CFG.SHEET_PXK);
  const pxkId = getPxkIdFromSheet_(shPXK);

  // ===== PXK boundaries =====
  const pxkStart = findStartRowAfterHeader_(shPXK, CFG.PK_HEADER, CFG.COL_PRODKEY);
  if (pxkStart === -1) throw new Error("PXK: cannot find prodKey header in col A.");
  const pxkHr = pxkStart - 1;

  const pos = findPXKHeaderPos_(shPXK);
  const pxkEnd = findPXKDataEndRow_(shPXK, pos.row, pos.col);
  if (pxkEnd < pxkStart) return [];

  const pxkLastCol = Math.max(1, shPXK.getLastColumn());
  const pxkHeaders = shPXK.getRange(pxkHr, 1, 1, pxkLastCol).getDisplayValues()[0].map(v => String(v || "").trim());
  const p = makeHeaderIndexMap_(pxkHeaders);

  // ===== PXK columns (1-based) =====
  const colPK = mustFindHeader_(p, [CFG.PK_HEADER], "PXK");
  const colTen = mustFindHeader_(p, [CFG.PXK_H_TENSP, "Ten Sp"], "PXK");

  const colPrint = findFirstHeaderCol_(p, [CFG.PXK_H_PRINT, "PRINT"]);
  const colLoai  = findFirstHeaderCol_(p, [CFG.PXK_H_LOAI, "LOAI", "Loai"]);
  const colD     = findFirstHeaderCol_(p, [CFG.PXK_H_D]);
  const colR     = findFirstHeaderCol_(p, [CFG.PXK_H_R]);
  const colC     = findFirstHeaderCol_(p, [CFG.PXK_H_C]);
  const colSL    = findFirstHeaderCol_(p, [CFG.PXK_H_SL]);

  const colGiay  = findFirstHeaderCol_(p, [CFG.PXK_H_GIAY, "Giay"]);
  const colKho   = findFirstHeaderCol_(p, [CFG.PXK_H_KHO, "Kho"]);
  const colDaiC  = findFirstHeaderCol_(p, [CFG.PXK_H_DAIC, "Dai-C"]);
  const colSLC   = findFirstHeaderCol_(p, [CFG.PXK_H_SLC, "SL C", "SLc"]);
  const colCan   = findFirstHeaderCol_(p, [CFG.PXK_H_CANLAN, "Can lan"]);
  const colDH    = findFirstHeaderCol_(p, [CFG.PXK_H_DH, "DH so", "DH số"]);
  const colCus   = findFirstHeaderCol_(p, [CFG.PXK_H_CUSTOMER, "Khách hàng", "Khach hang", "KH"]);
  const colM2    = findFirstHeaderCol_(p, [CFG.PXK_H_M2, "M²", "M_2"]);
  const colM3    = findFirstHeaderCol_(p, [CFG.PXK_H_M3, "M³", "M_3"]);
  const colM     = findFirstHeaderCol_(p, [CFG.PXK_H_M, "Song TM /M", "Sóng TM /M", "M"]);

  const pxkN = pxkEnd - pxkStart + 1;

  // ===== build DOT rows (same schema as ensureDotTemplate_) =====
  const out = [];
  const c = indexHeaders_(DOT_HEADERS);
  const W = DOT_HEADERS.length;

  const pxkVals = shPXK.getRange(pxkStart, 1, pxkN, pxkLastCol).getDisplayValues();

  for (let i = 0; i < pxkN; i++) {
    const rV = pxkVals[i];

    const pk = String(rV[colPK - 1] || "").trim();
    if (!pk) continue;

    // IMPORTANT: dùng FULL prodKey để không làm mất hàng trùng (#0/#1)
    if (existedPK && existedPK.has(dotUniqueKey_(pxkId, pk))) continue;

    const vTen   = String(rV[colTen - 1] || "").trim();
    const vPrint = colPrint ? String(rV[colPrint - 1] || "").trim() : "";
    const vLoai  = colLoai  ? String(rV[colLoai  - 1] || "").trim() : "";
    const vD_    = colD     ? String(rV[colD     - 1] || "").trim() : "";
    const vR_    = colR     ? String(rV[colR     - 1] || "").trim() : "";
    const vC_    = colC     ? String(rV[colC     - 1] || "").trim() : "";
    const vSL_   = colSL    ? String(rV[colSL    - 1] || "").trim() : "";

    const vGiay  = colGiay ? String(rV[colGiay - 1] || "").trim() : "";
    const vKho   = colKho  ? String(rV[colKho  - 1] || "").trim() : "";
    const vDaiC  = colDaiC ? String(rV[colDaiC - 1] || "").trim() : "";
    const vSLC   = colSLC  ? ceilFromDisplay_(rV[colSLC  - 1]) : "";
    const vCan   = colCan  ? String(rV[colCan  - 1] || "").trim() : "";
    const vDH    = colDH   ? String(rV[colDH   - 1] || "").trim() : "";
    const vCus   = colCus  ? String(rV[colCus  - 1] || "").trim() : "";

    const vM2    = colM2 ? to2dec_(rV[colM2 - 1]) : "";
    const vM3    = colM3 ? to2dec_(rV[colM3 - 1]) : "";
    const vM_    = colM  ? to2dec_(rV[colM  - 1]) : "";

    const tem1 = `${vPrint ? `(${vPrint}) ` : ""}${vCus}${vGiay ? `- ${vGiay}-` : ""}${vDH ? `\n${vDH}` : ""}`.trim();
    const tem2 = `${vKho} x ${vDaiC} = ${vSLC}`.replace(/\s+/g, " ").trim();

    const sizeDRC = [vD_, vR_, vC_].filter(x => String(x || "").trim() !== "");
    const tem3 = `${sizeDRC.join(" x ")} = ${vSL_}-- ${vLoai}`.replace(/\s+/g, " ").trim();

    const tem4 = vTen;
    const tem5 = vCan;

    const row = new Array(W).fill("");
    row[c["PXK_ID"]] = pxkId;
    row[c["prodKey"]] = pk;
    row[c["Tên Sp"]] = vTen;

    row[c["Print"]] = vPrint;
    row[c["LOẠI"]] = vLoai;
    row[c["D"]] = vD_;
    row[c["R"]] = vR_;
    row[c["C"]] = vC_;
    row[c["SL"]] = vSL_;

    row[c["Giấy"]] = vGiay;
    row[c["Khổ"]] = vKho;
    row[c["Dài-C"]] = vDaiC;
    row[c["SL-C"]] = vSLC;

    row[c["Cán lằn"]] = vCan;
    row[c["DH"]] = vDH;
    row[c["Customer"]] = vCus;

    row[c["M2"]] = vM2;
    row[c["M3"]] = vM3;
    row[c["M"]]  = vM_;

    row[c["temField1"]] = tem1;
    row[c["temField2"]] = tem2;
    row[c["temField3"]] = tem3;
    row[c["temField4"]] = tem4;
    row[c["temField5"]] = tem5;
    row[c["temSTT"]] = ""; // reset after sort

    out.push(row);
  }

  return out;
}

/*************************************
 * DOT border formatting
 *************************************/
function applyDotBorders_(shDot, hr, lastRow, lastCol) {
  if (lastRow <= hr || lastCol < 1) return;
  const rng = shDot.getRange(hr, 1, lastRow - hr + 1, lastCol);
  rng.setBorder(true, true, true, true, true, true);
}

function applyDotFormats_(shDot) {
  const hr = CFG.DOT_HEADER_ROW;
  const headers = getHeader_(shDot, hr);
  const idx = indexHeaders_(headers);

  const lastRow = shDot.getLastRow();
  if (lastRow <= hr) return;
  const n = lastRow - hr;

  // M2/M3/M: 0.00
  if (idx["M2"] != null && idx["M3"] != null && idx["M"] != null) {
    shDot.getRange(hr + 1, idx["M2"] + 1, n, 3).setNumberFormat("0.00");
  }

  if (idx["Dài-C"] != null) {
    shDot.getRange(hr + 1, idx["Dài-C"] + 1, n, 1).setNumberFormat("@");
  }

  // SL-C: integer
  if (idx["SL-C"] != null) {
    shDot.getRange(hr + 1, idx["SL-C"] + 1, n, 1).setNumberFormat("0");
  }

  // DH as text
  if (idx["DH"] != null) {
    shDot.getRange(hr + 1, idx["DH"] + 1, n, 1).setNumberFormat("@");
  }

  // temSTT: center + middle
  if (idx["temSTT"] != null) {
    shDot.getRange(hr + 1, idx["temSTT"] + 1, n, 1)
      .setHorizontalAlignment("CENTER")
      .setVerticalAlignment("MIDDLE");
  }
}

function sortDOT_(shDot) {
  const hr = CFG.DOT_HEADER_ROW;
  const headers = getHeader_(shDot, hr);
  const idx = indexHeaders_(headers);
  ["Giấy", "Khổ", "Dài-C"].forEach(h => { if (idx[h] == null) throw new Error(`DOT missing header '${h}' for sorting`); });

  const lastRow = shDot.getLastRow();
  const lastCol = shDot.getLastColumn();
  if (lastRow <= hr + 1) return;

  shDot.getRange(hr + 1, 1, lastRow - hr, lastCol).sort([
    { column: idx["Giấy"] + 1, ascending: false }, // Giấy DESC
    { column: idx["Khổ"] + 1, ascending: false },  // Khổ DESC
    { column: idx["Dài-C"] + 1, ascending: false } // Dài-C DESC
  ]);
}

/*************************************
 * Fill temSTT after sort — RESET 1..n for current DOT sheet
 *************************************/
function fillDotSttAfterSort_(shDot) {
  const hr = CFG.DOT_HEADER_ROW;
  const headers = getHeader_(shDot, hr);
  const idx = indexHeaders_(headers);

  if (idx["temSTT"] == null) throw new Error(`${shDot.getName()} missing header 'temSTT'`);

  const lastRow = shDot.getLastRow();
  if (lastRow <= hr) return;

  const n = lastRow - hr;
  const colStt = idx["temSTT"] + 1;

  const out = Array.from({ length: n }, (_, i) => [String(i + 1)]);
  shDot.getRange(hr + 1, colStt, n, 1).setValues(out);
}

/*************************************
 * 6) TEM build from DOT_n (with QR)
 *************************************/
function buildTemFromDot(batchNoInput) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  // --- 1) Resolve batchNo (prompt if not provided) ---
  let batchNo = null;

  if (batchNoInput != null && String(batchNoInput).trim() !== "") {
    batchNo = parseInt(String(batchNoInput).trim(), 10);
  } else {
    const def = String(getCurrentBatchNo_());
    const resp = ui.prompt(
      "Generate TEM",
      `Nhập số sheet TEM muốn tạo dựa trên DOT có sẵn (default OK: TEM_${def})\n VD: Nhập số 1 để tạo TEM_1 từ DOT_1\n
      ✏️ Bạn muốn tạo sheet TEM số:`,
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    const raw = String(resp.getResponseText() || def).trim();
    batchNo = parseInt(raw, 10);
  }

  if (!Number.isInteger(batchNo) || batchNo < 1) {
    ui.alert("❌ Số batch không hợp lệ.");
    return;
  }

  // --- 2) DOT MUST exist (DO NOT auto-create) ---
  const dotNm = dotName_(batchNo);
  const shDot = ss.getSheetByName(dotNm);
  if (!shDot) {
    ui.alert(`❌ Không tìm thấy sheet ${dotNm}. (Generate DOT trước)`);
    return;
  }
  if (shDot.getLastRow() <= CFG.DOT_HEADER_ROW) {
    ui.alert(`❌ ${dotNm} đang trống (không có dữ liệu).`);
    return;
  }

  // TEM may be created
  const shTem = getBatchTemSheet_(ss, batchNo);

  // Keep batch pointer consistent
  setCurrentBatchNo_(batchNo);

  // --- 3) Build PXK map: prodKey(base) -> KHGH raw text ---
  const shPXK = mustGetSheet_(ss, CFG.SHEET_PXK);
  const masterUrl = buildMasterScanUrl_(shPXK);

  const pxkStart = findStartRowAfterHeader_(shPXK, CFG.PK_HEADER, CFG.COL_PRODKEY);
  if (pxkStart === -1) throw new Error("PXK: cannot find 'prodKey' header.");
  const pxkHr = pxkStart - 1;

  const pos = findPXKHeaderPos_(shPXK);
  const pxkEnd = findPXKDataEndRow_(shPXK, pos.row, pos.col);
  const pxkLastCol = Math.max(1, shPXK.getLastColumn());

  const pxkHeaders = shPXK.getRange(pxkHr, 1, 1, pxkLastCol).getDisplayValues()[0].map(v => String(v || "").trim());
  const pCol = makeHeaderIndexMap_(pxkHeaders);

  const pxkPkCol   = mustFindHeader_(pCol, [CFG.PK_HEADER], "PXK");
  const pxkKhghCol = mustFindHeader_(pCol, [CFG.PXK_H_KHGH, "Khgh"], "PXK");

  const pxkMap = new Map();
  if (pxkEnd >= pxkStart) {
    const pxkN = pxkEnd - pxkStart + 1;
    const pxkVals = shPXK.getRange(pxkStart, 1, pxkN, pxkLastCol).getDisplayValues();
    for (const r of pxkVals) {
      const pk = String(r[pxkPkCol - 1] || "").trim();
      const khgh = String(r[pxkKhghCol - 1] || "").trim();
      if (!pk || !khgh) continue;
      const base = pk.split("#")[0];
      if (!pxkMap.has(base)) pxkMap.set(base, khgh);
    }
  }

  // --- 4) Read DOT once, filter rows (pk & qr) ---
  const hr = CFG.DOT_HEADER_ROW;
  const lastRow = shDot.getLastRow();
  const lastCol = shDot.getLastColumn();

  const headers = shDot.getRange(hr, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v || "").trim());
  const idx = indexHeaders_(headers);

  const needed = ["temSTT","prodKey","temField1","temField2","temField3","temField4","temField5"];
  for (const h of needed) if (idx[h] == null) throw new Error(`DOT missing header '${h}'`);

  const nAll = lastRow - hr;
  const allRows = shDot.getRange(hr + 1, 1, nAll, lastCol).getValues();

  const rows = [];
  for (const r of allRows) {
    const pk = String(r[idx["prodKey"]] || "").trim();
    if (pk) rows.push(r);
  }

  // --- 5) Clear TEM + prepare output matrix (single write) ---
  shTem.clear();

  const COLS = CFG.TEM_COLS_PER_ROW || 2;
  const ROWS_PER_PAGE = CFG.TEM_ROWS_PER_PAGE || 5;
  const H0 = CFG.TEM_ROWS_PER_LABEL || 7;
  const GR = CFG.TEM_GAP_ROWS || 0;
  const H  = H0 + GR;

  const total = rows.length;
  if (!total) {
    ui.alert(`❌ ${dotNm} không có dòng hợp lệ (thiếu prodKey).`);
    return;
  }

  const perPage = COLS * ROWS_PER_PAGE;
  const pages = Math.ceil(total / perPage);
  const totalBlocks = pages * ROWS_PER_PAGE;
  const outRows = totalBlocks * H;
  const COLS_PER_LABEL = 3; // qr nhỏ | nội dung chính | gap
  const outCols = COLS * COLS_PER_LABEL - 1;

  const values = Array.from({ length: outRows }, () => Array(outCols).fill(""));
  const formulaAll = Array.from({ length: outRows }, () => Array(outCols).fill("")); // setFormulas 1 lần

  // Fill values + formulas in-memory
  for (let i = 0; i < total; i++) {
    const r = rows[i];

    const page = Math.floor(i / perPage);
    const pos2  = i % perPage;
    const rowInPage = Math.floor(pos2 / COLS);
    const colInRow  = pos2 % COLS;

    const baseR = page * (ROWS_PER_PAGE * H) + rowInPage * H;
    const col2  = colInRow * 3;

    // Row 1 + 2: QR master nằm bên trái, cao bằng 2 hàng đầu
    values[baseR + 0][col2 + 1] = r[idx["temField1"]] || "";
    values[baseR + 1][col2 + 1] = r[idx["temField2"]] || "";

    if (masterUrl) {
      const qrMasterUrl =
        "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" +
        encodeURIComponent(masterUrl);
      formulaAll[baseR + 0][col2] = `=IMAGE("${qrMasterUrl}")`;
    }

    // Rows 3..5: nội dung chính ở cột phải
    values[baseR + 2][col2 + 1] = r[idx["temField3"]] || "";
    values[baseR + 3][col2 + 1] = r[idx["temField4"]] || "";
    values[baseR + 4][col2 + 1] = r[idx["temField5"]] || "";

    // Row 6 = barcode item full width: đặt ở cột phải
    const itemCode = String(r[idx["prodKey"]] || "").trim();
    if (itemCode) {
      const barcodeUrl =
        "https://bwipjs-api.metafloor.com/?bcid=code128" +
        "&scale=4" +
        "&paddingwidth=20" +
        "&includetext" +
        "&text=" + encodeURIComponent(itemCode);
      formulaAll[baseR + 5][col2 + 1] = `=IMAGE("${barcodeUrl}")`;
    }

    // Row 7 = footer ở cột phải
    const pk = String(r[idx["prodKey"]] || "").trim();
    const base = pk.split("#")[0];
    const khghText = String(pxkMap.get(base) || "").trim();
    const stt = String(r[idx["temSTT"]] || "").trim();
    values[baseR + 6][col2 + 1] = khghText ? `${khghText}                     ${stt}` : `${stt}`;
  }

  // Ghi số trang ở góc phải hàng cuối của mỗi page 2x5
  for (let p = 0; p < pages; p++) {
    const footerRow0 = (p + 1) * (ROWS_PER_PAGE * H) - 1; // 0-based: row cuối của page
    values[footerRow0][outCols - 1] = `- - - - - ${p + 1}/${pages} - - - - -`;
  }

  // Write values once
  shTem.getRange(1, 1, outRows, outCols).setValues(values);

  // Write formulas ONLY to cells that actually contain formulas
  // để không đè mất temField1 / text thường
  for (let r = 0; r < outRows; r++) {
    for (let c = 0; c < outCols; c++) {
      const f = formulaAll[r][c];
      if (f) {
        shTem.getRange(r + 1, c + 1).setFormula(f);
      }
    }
  }

  // Formatting
  marktemField1If3L_(shTem, outRows, outCols);
  applyTemFormatting_(shTem, outRows, outCols);
  autoFitTemFonts_(shTem, outRows, outCols);
  forceTemOuterBorders_(shTem, outRows, outCols);

  ui.alert(`✅ ${shTem.getName()} generated successfully from ${dotNm} \n\n Quantity: ${total}\nPages: ${pages}`);
}

/*************************************
 * 7) Export TEM PDF (choose batch)
 *************************************/
function exportTemPdf() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const def = String(getCurrentBatchNo_());
  const resp = ui.prompt(
    "Export TEM PDF",
    `Nhập số sheet TEM muốn export (default OK: TEM_${def})\n VD: Nhập số 1 để tạo TEM_1 từ DOT_1 \n
    ✏️ Bạn muốn export PDF của sheet TEM số:`,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const batchNo = parseInt(String(resp.getResponseText() || def).trim(), 10);
  if (!Number.isInteger(batchNo) || batchNo < 1) {
    ui.alert("❌ Batch number không hợp lệ.");
    return;
  }

  const sh = ss.getSheetByName(temName_(batchNo));
  if (!sh) {
    ui.alert(`❌ Không tìm thấy sheet ${temName_(batchNo)}. Build TEM trước.`);
    return;
  }

  const outRows = sh.getLastRow();
  const outCols = sh.getLastColumn();
  if (outRows < 1 || outCols < 1) {
    ui.alert(`❌ ${sh.getName()} đang trống. Build TEM trước.`);
    return;
  }

  const gid = sh.getSheetId();
  const rangeA1 = `A1:${colIndexToLetter_(outCols)}${outRows}`;

  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export" +
    "?format=pdf" +
    "&gid=" + gid +
    "&range=" + encodeURIComponent(rangeA1) +
    "&portrait=true" +
    "&size=A4" +
    "&scale=2" +
    "&sheetnames=false&pagenumbers=false&gridlines=false&fzr=false" +
    "&top_margin=0.00&bottom_margin=0.00&left_margin=0.00&right_margin=0.00";

  const token = ScriptApp.getOAuthToken();
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + token } }).getBlob();

  const ts = Utilities.formatDate(new Date(), CFG.TZ, "yyyyMMdd_HHmmss");
  const fileName = `${sh.getName()}_${ts}.pdf`;
  const file = DriveApp.createFile(blob.setName(fileName));

  ui.alert("✅ Exported TEM pdf!\n\n" + "📍 Open link below to print " + fileName + "\n\n" + file.getUrl());
}

/*************************************
 * PXK PRINT RANGE (auto) + EXPORT PDF
 * - Auto chọn vùng giống file mẫu, không cần user chọn "vùng in"
 *************************************/

/**
 * Find stop row where footer begins (line contains "nhận sản phẩm", "**", "cảm ơn", ...)
 * Return row index (1-based). If not found, returns -1.
 */
function findPXKStopRow_(shPXK, headerRow) {
  const lastRow = shPXK.getLastRow();
  if (lastRow <= headerRow) return -1;

  const n = lastRow - headerRow;
  const maxScanCols = Math.min(8, shPXK.getLastColumn());
  const block = shPXK.getRange(headerRow + 1, 1, n, maxScanCols).getDisplayValues();

  const normStop_ = (s) => String(s || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/\s+/g, " ");

  const isStopLine_ = (s) => {
    const t = normStop_(s);
    return (
      t.includes("vui long") ||
      t.includes("da nhan du") ||
      t.includes("nhan san pham") ||
      t.startsWith("**") ||
      t.includes("cam on")
    );
  };

  for (let i = 0; i < n; i++) {
    const r = headerRow + 1 + i;
    const joined = block[i].slice(0, 8).join(" ");
    if (isStopLine_(joined)) return r;
  }
  return -1;
}

/**
 * Build PXK export range A1:?? using:
 * - EndRow = stopRow + footerRows (to include signature/footer)
 * - EndCol = header "DH" (configurable) if found, else lastColumn
 */
function getPXKPrintRange_(shPXK) {
  const pos = findPXKHeaderPos_(shPXK);
  const hr = pos.row;

  const lastCol = shPXK.getLastColumn();
  const headers = shPXK.getRange(hr, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(v => String(v || "").trim());
  const hMap = makeHeaderIndexMap_(headers);

  // START COL = Tên Sp
  const startCol1 = mustFindHeader_(hMap, [CFG.PXK_H_TENSP, "Ten Sp", "Tên SP", "Ten SP", "ten sp"], "PXK");

  // ---- END ROW logic (giữ nguyên như bạn đang dùng) ----
  const stopRow = findPXKStopRow_(shPXK, hr);
  const footerRows = Number(CFG.PXK_PRINT_FOOTER_ROWS || 0);

  let endRow;
  if (stopRow !== -1) {
    endRow = Math.min(shPXK.getLastRow(), stopRow + footerRows);
  } else {
    const dataEnd = findPXKDataEndRow_(shPXK, hr, pos.col);
    endRow = Math.min(shPXK.getLastRow(), dataEnd + footerRows);
  }

  // ---- END COL = DH (như config) ----
  const wantHeader = String(CFG.PXK_PRINT_LAST_HEADER || "DH").trim();
  const endCol1 = findFirstHeaderCol_(hMap, [wantHeader, "DH"]) || lastCol;

  const startLetter = colIndexToLetter_(startCol1);
  const endLetter   = colIndexToLetter_(endCol1);

  return `${startLetter}1:${endLetter}${endRow}`;
}

/** Export PXK "Phiếu Xuất Kho" to PDF with auto-range */
function exportPxkPdf() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const shPXK = ss.getSheetByName(CFG.SHEET_PXK);

  if (!shPXK) {
    ui.alert(`❌ Không tìm thấy sheet ${CFG.SHEET_PXK}.`);
    return;
  }

  const shLog = getOrCreateLog_(ss);

  // 1) CLEAN-UP rows rác trước khi in (clear nguyên rows thiếu PK + Tên Sp)
  try {
    pxkCleanupRowsForPrint_(shPXK);
  } catch (e) {
    ui.alert("❌ Clean-up PXK failed.\n\n" + (e?.message || e));
    return;
  }

  // 2) Tạm hide các cột nằm giữa TênSp..LOẠI và Giấy..DH để PDF KHÔNG in
  let restoreState = [];
  try {
    restoreState = pxkTempHideInternalColsForPdf_(shPXK);
  } catch (e) {
    ui.alert("❌ Hide internal columns failed.\n\n" + (e?.message || e));
    return;
  }

  // 3) Export PDF (auto-range)
  let rangeA1 = "";
  try {
    rangeA1 = getPXKPrintRange_(shPXK);
  } catch (e) {
    // restore before exiting
    try { pxkRestoreHiddenState_(shPXK, restoreState); } catch (_) {}
    ui.alert("❌ Không xác định được vùng in PXK.\n\n" + (e?.message || e));
    return;
  }

  try {
    const gid = shPXK.getSheetId();

    const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export" +
      "?format=pdf" +
      "&gid=" + gid +
      "&range=" + encodeURIComponent(rangeA1) +
      "&portrait=true" +
      "&size=A4" +
      "&scale=2" +
      "&sheetnames=false&printtitle=false&pagenum=RIGHT&gridlines=false&fzr=false" +
      "&top_margin=0.4&bottom_margin=0.9&left_margin=0.6&right_margin=0.6" +
      "&horizontal_alignment=CENTER";

    const token = ScriptApp.getOAuthToken();
    const blob = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + token }
    }).getBlob();

    const ts = Utilities.formatDate(new Date(), CFG.TZ, "yyyyMMdd_HHmmss");
    const ssDriveName = DriveApp.getFileById(ss.getId()).getName();
    const fileName = `${ssDriveName}_${ts}.pdf`;
    const file = DriveApp.createFile(blob.setName(fileName));

    appendLog_(shLog, "SYSTEM", "EXPORT_PXK_PDF", "", "OK", `RANGE=${rangeA1} | FILE=${fileName}`);

    ui.alert(
      "✅ Exported PXK pdf!\n\n" +
      "📍 Open link below to print " + ssDriveName + "\n\n" +
      file.getUrl()
    );
  } finally {
    try { pxkRestoreHiddenState_(shPXK, restoreState); } catch (e) {}
  }
}

/**
 * CLEAN-UP trước khi in:
 * - Chỉ clear các dòng DATA (từ startRow đến dataEnd)
 * - Điều kiện clear: thiếu CẢ prodKey và Tên Sp
 * - Clear nguyên row (ALL columns) -> bao gồm cả cột hidden / cột rỗng / cột không in
 * - KHÔNG đụng footer/signature (nằm dưới stop line)
 */
function pxkCleanupRowsForPrint_(shPXK) {
  const pos = findPXKHeaderPos_(shPXK);
  const hr = pos.row;

  const pxkStart = findStartRowAfterHeader_(shPXK, CFG.PK_HEADER, CFG.COL_PRODKEY);
  if (pxkStart === -1) throw new Error("PXK: cannot find prodKey header in col A.");
  const startRow = pxkStart;

  const lastCol = shPXK.getLastColumn();
  const headers = shPXK.getRange(hr, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v || "").trim());
  const hMap = makeHeaderIndexMap_(headers);

  const colPK  = findFirstHeaderCol_(hMap, [CFG.PK_HEADER, "prodKey"]);
  const colTen = findFirstHeaderCol_(hMap, [CFG.PXK_H_TENSP, "Ten Sp"]);
  if (!colPK || !colTen) throw new Error("PXK: missing prodKey or Tên Sp header.");

  // Data end = trước stop line (nếu có), để KHÔNG clear footer/signature
  const stopRow = findPXKStopRow_(shPXK, hr); // row chứa dòng "** Nhận sản phẩm..."
  const dataEnd = (stopRow !== -1) ? (stopRow - 1) : findPXKDataEndRow_(shPXK, hr, pos.col);
  if (dataEnd < startRow) return;

  const n = dataEnd - startRow + 1;

  // Read ONLY 2 cols for condition (fast)
  const pkVals  = shPXK.getRange(startRow, colPK,  n, 1).getDisplayValues();
  const tenVals = shPXK.getRange(startRow, colTen, n, 1).getDisplayValues();

  // Clear whole row across ALL columns
  const rowsToClear = [];
  for (let i = 0; i < n; i++) {
    const pk  = String(pkVals[i][0] || "").trim();
    const ten = String(tenVals[i][0] || "").trim();
    if (!pk && !ten) rowsToClear.push(startRow + i);
  }

  // Batch clear with RangeList (avoid looping clear each cell)
  if (rowsToClear.length) {
    const a1s = rowsToClear.map(r => `A${r}:${colIndexToLetter_(lastCol)}${r}`);
    shPXK.getRangeList(a1s).clearContent();
  }
}

/**
 * Tạm HIDE các cột nằm GIỮA:
 * - Tên Sp ... LOẠI
 * - Giấy ... DH
 * (kể cả cột rỗng)
 * => để PDF KHÔNG in các cột đó
 *
 * Return: [{col, wasHidden}]
 */
function pxkTempHideInternalColsForPdf_(shPXK) {
  const pos = findPXKHeaderPos_(shPXK);
  const hr = pos.row;

  const lastCol = shPXK.getLastColumn();
  const headers = shPXK.getRange(hr, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v || "").trim());
  const hMap = makeHeaderIndexMap_(headers);

  const colTen  = mustFindHeader_(hMap, [CFG.PXK_H_TENSP, "Ten Sp"], "PXK");
  const colLoai = mustFindHeader_(hMap, [CFG.PXK_H_LOAI, "LOAI", "Loai"], "PXK");
  const colGiay = mustFindHeader_(hMap, [CFG.PXK_H_GIAY, "Giay"], "PXK");
  const colDH   = mustFindHeader_(hMap, [CFG.PXK_H_DH, "DH"], "PXK");

  const cols = [];

  // between Tên Sp and LOẠI
  for (let c = colTen + 1; c <= colLoai - 1; c++) cols.push(c);

  // between Giấy and DH
  for (let c = colGiay + 1; c <= colDH - 1; c++) cols.push(c);

  // uniq + valid
  const uniq = [...new Set(cols)].filter(c => c >= 1 && c <= lastCol);

  const state = uniq.map(c => ({ col: c, wasHidden: shPXK.isColumnHiddenByUser(c) }));

  // hide all of them (even if already hidden)
  for (const c of uniq) {
    try { shPXK.hideColumn(shPXK.getRange(1, c)); } catch (e) {}
  }

  return state;
}

/** Restore các cột đã hide tạm (chỉ show lại những cột trước đó KHÔNG hidden) */
function pxkRestoreHiddenState_(shPXK, state) {
  if (!state || !state.length) return;
  for (const it of state) {
    if (!it.wasHidden) {
      try { shPXK.showColumns(it.col); } catch (e) {}
    }
  }
}

/*************************************
 * TEM formatting + auto-fit fonts (keep your existing logic)
 * applyTemFormatting_, autoFitTemFonts_, marktemField1If3L_
 *************************************/
/* Apply same format to TEM same as DOT e.g. bold, red colored */
function applyTemFormatting_(shTem, outRows, outCols) {
  const H0 = CFG.TEM_ROWS_PER_LABEL || 7;
  const GR = CFG.TEM_GAP_ROWS || 0;
  const H = H0 + GR;
  const ROWS_PER_PAGE = CFG.TEM_ROWS_PER_PAGE || 5;

  const rngAll = shTem.getRange(1, 1, outRows, outCols);

  // base
  rngAll.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  rngAll.setVerticalAlignment("MIDDLE");
  rngAll.setHorizontalAlignment("CENTER");
  rngAll.setFontFamily("Arial");

  // row heights per tem line (bản 3x4)
  // const rowHeights = [48, 28, 24, 66, 40, 138, 18, 5];
  const rowHeights = [50, 28, 24, 66, 40, 72, 18, 5];
  for (let r = 1; r <= outRows; r++) {
    const slot = (r - 1) % H;
    shTem.setRowHeightsForced(r, 1, rowHeights[slot]);
  }

  const lastColLetter = colIndexToLetter_(outCols);

  // ---------- Wrap riêng line 4 theo block 3 cột: QR | TEXT | GAP ----------
  // line 4 đã merge 2 cột đầu, nên phải target cột trái của block: 1,4,7...
  {
    const ranges = [];
    for (let base = 1; base <= outRows; base += H) {
      const r = base + 3; // line 4
      for (let c = 1; c <= outCols; c += 3) {
        ranges.push(`${colIndexToLetter_(c)}${r}`);
      }
    }
    if (ranges.length) shTem.getRangeList(ranges).setWrap(true);
  }
  // ---------- Line-specific formatting (batched by RangeList) ----------
  const lineRows = (offset) => {
    const arr = [];
    for (let base = 1; base <= outRows; base += H) {
      const r = base + offset;
      arr.push(`A${r}:${lastColLetter}${r}`);
    }
    return arr;
  };

  const rowsLine1 = lineRows(0);
  const rowsLine2 = lineRows(1);
  const rowsLine3 = lineRows(2);
  const rowsLine4 = lineRows(3);
  const rowsLine5 = lineRows(4);
  const rowsLine7 = lineRows(6);

  if (rowsLine1.length) shTem.getRangeList(rowsLine1).setFontWeight("normal");
  if (rowsLine2.length) shTem.getRangeList(rowsLine2).setFontWeight("bold");
  if (rowsLine3.length) shTem.getRangeList(rowsLine3).setFontWeight("normal");
  if (rowsLine4.length) shTem.getRangeList(rowsLine4).setFontWeight("bold");
  if (rowsLine5.length) shTem.getRangeList(rowsLine5).setFontWeight("normal");
  if (rowsLine7.length) shTem.getRangeList(rowsLine7).setFontSize(12).setFontWeight("normal");

  // safety: redraw outer border after merge để không bị mất viền
  for (let c = 1; c <= outCols; c += 3) {
    for (let base = 1; base <= outRows; base += H) {
      shTem.getRange(base, c, H0, 2).setBorder(
        true, true, true, true, false, false,
        "#000000",
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );
    }
  }

  // ---------- Column widths: qr nhỏ, tem wide, gap narrow ----------
  const W_QR  = 70;
  const W_TEM = 200;
  const W_GAP = 5;

  for (let c = 1; c <= outCols; c++) {
    const mod = (c - 1) % 3;
    if (mod === 0) shTem.setColumnWidth(c, W_QR);       // QR nhỏ
    else if (mod === 1) shTem.setColumnWidth(c, W_TEM); // nội dung chính
    else shTem.setColumnWidth(c, W_GAP);                // gap
  }

  // ---------- Merge theo layout mới ----------
  // line 1+2: QR ở cột trái cao 2 hàng
  // line 1+2 bên phải là text thường
  // line 3..7: merge full-width 2 cột đầu
  for (let c = 1; c <= outCols; c += 3) {
    for (let base = 1; base <= outRows; base += H) {
      // QR bên trái cao 2 hàng
      shTem.getRange(base, c, 2, 1).merge();

      // text line 1 bên phải
      shTem.getRange(base, c + 1, 1, 1).breakApart();
      // text line 2 bên phải
      shTem.getRange(base + 1, c + 1, 1, 1).breakApart();

      // line 3..7 merge full-width
      for (let k = 2; k <= 6; k++) {
        shTem.getRange(base + k, c, 1, 2).merge();
      }
    }
  }

  // Format dòng footer số trang
  for (let base = 1; base <= outRows; base += (ROWS_PER_PAGE * H)) {
    const footerRow = base + (ROWS_PER_PAGE * H) - 1; // hàng cuối của mỗi page
    const cell = shTem.getRange(footerRow, outCols, 1, 1);
    const v = String(cell.getDisplayValue() || "").trim();
    const hasPageNo = !!v;

    shTem.setRowHeight(footerRow, hasPageNo ? 15 : 5);

    cell
      .setHorizontalAlignment("RIGHT")
      .setVerticalAlignment("MIDDLE")
      .setFontSize(hasPageNo ? 10 : 8)
      .setFontWeight("normal");
  }
}

function forceTemOuterBorders_(shTem, outRows, outCols) {
  const H0 = CFG.TEM_ROWS_PER_LABEL || 7;
  const GR = CFG.TEM_GAP_ROWS || 0;
  const H = H0 + GR;

  const THICK_COLOR = "#000000";
  const DASH_COLOR  = "#999999";

  for (let c = 1; c <= outCols; c += 3) {
    const leftCol = c;

    for (let base = 1; base <= outRows; base += H) {
      // 1) force outer border full block trước
      shTem.getRange(base, leftCol, H0, 2).setBorder(
        true, true, true, true, null, null,
        THICK_COLOR,
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );

      // 2) force dashed separators bên trong
      for (let k = 0; k < H0 - 1; k++) {
        shTem.getRange(base + k, leftCol, 1, 2).setBorder(
          null, null, true, null, null, null,
          DASH_COLOR,
          SpreadsheetApp.BorderStyle.DASHED
        );
      }

      // 3) ép lại left + right từng dòng để tránh lủng nét do merge/image
      for (let r = base; r <= base + H0 - 1; r++) {
        shTem.getRange(r, leftCol, 1, 2).setBorder(
          null, true, null, true, null, null,
          THICK_COLOR,
          SpreadsheetApp.BorderStyle.SOLID_THICK
        );
      }

      // 4) ép lại top row và bottom row lần cuối
      shTem.getRange(base, leftCol, 1, 2).setBorder(
        true, null, null, null, null, null,
        THICK_COLOR,
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );

      shTem.getRange(base + H0 - 1, leftCol, 1, 2).setBorder(
        null, null, true, null, null, null,
        THICK_COLOR,
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );
    }
  }
}

function autoFitTemFonts_(shTem, outRows, outCols) {
  const H0 = CFG.TEM_ROWS_PER_LABEL || 7;
  const GR = CFG.TEM_GAP_ROWS || 0;
  const H = H0 + GR;

  // line1..5
  const baseFont = [12, 16, 12, 18, 12];
  const minFont  = [10, 12, 10, 10, 10];

  const limitLen = [33, 18, 28, 200, 60];
  const PER_LINE = 18;

  const vals = shTem.getRange(1, 1, outRows, outCols).getDisplayValues();

  // layout mới: mỗi tem = 3 cột [QR | TEXT | GAP]
  // text anchor column là cột trái của block: 1,4,7...
  for (let c = 1; c <= outCols; c += 3) {
    const c0 = c - 1;
    const fontCol = Array.from({ length: outRows }, () => [10]);

    for (let base0 = 0; base0 < outRows; base0 += H) {
      for (let k = 0; k < 5; k++) {
        const r0 = base0 + k;
        const text = String(vals[r0]?.[c0] || "").trim();

        if (!text) {
          fontCol[r0][0] = baseFont[k];
          continue;
        }

        let fs = baseFont[k];

        // ===== LINE 4 = temField4 =====
        if (k === 3) {
          const parts = text.split("\n").map(s => String(s || "").trim()).filter(Boolean);
          const longest = parts.length ? Math.max(...parts.map(s => s.length)) : text.length;
          const totalLen = text.replace(/\n/g, "").length;

          let lines = 0;
          for (const p of (parts.length ? parts : [text])) {
            lines += Math.max(1, Math.ceil(p.length / PER_LINE));
          }

          if (lines >= 4) fs = 10;
          else if (lines === 3) fs = 11;
          else if (lines === 2) {
            if (longest <= 12 && totalLen <= 18) fs = 16;
            else if (longest <= 18 && totalLen <= 28) fs = 15;
            else fs = 14;
          } else {
            if (longest <= 8  && totalLen <= 8) fs = 22;
            else if (longest <= 12 && totalLen <= 12) fs = 20;
            else if (longest <= 18 && totalLen <= 18) fs = 18;
            else if (longest <= 26 && totalLen <= 26) fs = 16;
            else fs = 14;
          }

          fs = Math.max(minFont[k], fs);
          fontCol[r0][0] = fs;
          continue;
        }

        // ===== các line khác =====
        const L = text.length;
        if (L > limitLen[k]) {
          const over = L - limitLen[k];
          fs = Math.max(minFont[k], fs - Math.ceil(over / 6));
        }
        fontCol[r0][0] = fs;
      }
    }

    // apply cho cột text anchor
    shTem.getRange(1, c, outRows, 1).setFontSizes(fontCol);

    // line 2..7 đang merge 2 cột đầu, nên set luôn cột bên cạnh cho đồng bộ
    if (c + 1 <= outCols) {
      shTem.getRange(1, c + 1, outRows, 1).setFontSizes(fontCol);
    }
  }
}

// Colour/mark TEM 3L (normal for 5L)
function marktemField1If3L_(shTem, outRows, outCols) {
  const H0 = CFG.TEM_ROWS_PER_LABEL || 7;
  const GR = CFG.TEM_GAP_ROWS || 0;
  const H = H0 + GR;

  const bgHit = "#bfbfbf";
  const bgNo = "#ffffff";

  const vals = shTem.getRange(1, 1, outRows, outCols).getDisplayValues();

  const hit = [];
  const no = [];

  for (let base0 = 0; base0 < outRows; base0 += H) {
    const r0 = base0; // line1 (0-based)
    const r1 = r0 + 1; // 1-based row
    for (let c1 = 2; c1 <= outCols; c1 += 3) {
      const v = String(vals[r0][c1 - 1] || "").trim();
      const a1 = `${colIndexToLetter_(c1)}${r1}`;
      if (v.includes("3L")) hit.push(a1);
      else no.push(a1);
    }
  }

  if (hit.length) shTem.getRangeList(hit).setBackground(bgHit);
  if (no.length) shTem.getRangeList(no).setBackground(bgNo);
}

/*************************************
 * 8) WebApp handler: fill Ngay_STEP in DATA (no overwrite) + log + push PXK (no overwrite)
 * NOTE: Client wants NO input fields on webapp anymore.
 * => WebApp should ONLY read prodKey from URL and show step buttons.
 *************************************/

function updatePxkStep_(prodKey, stepName, userEmail) {
  const ss = SpreadsheetApp.getActive();
  const shPXK = mustGetSheet_(ss, CFG.SHEET_PXK);
  const shLog = getOrCreateLog_(ss);

  const pk = String(prodKey || "").trim();
  const step = String(stepName || "").trim().toUpperCase();

  if (!pk || !step) {
    return { ok: false, code: "BAD_INPUT", msg: "Thiếu prodKey hoặc công đoạn" };
  }

  const colName = `${CFG.STEP_PREFIX}${step}`;

  const pxkStart = findStartRowAfterHeader_(shPXK, CFG.PK_HEADER, CFG.COL_PRODKEY);
  if (pxkStart === -1) {
    return { ok: false, code: "PXK_HEADER_NOT_FOUND", msg: "PXK: cannot find 'prodKey' header." };
  }

  const hr = pxkStart - 1;
  const lastCol = shPXK.getLastColumn();
  const headers = shPXK.getRange(hr, 1, 1, lastCol).getDisplayValues()[0]
    .map(v => String(v || "").trim());
  const pCol = makeHeaderIndexMap_(headers);

  const pkCol1 = mustFindHeader_(pCol, [CFG.PK_HEADER], "PXK");
  const stepCol1 = findFirstHeaderCol_(pCol, [colName]);

  if (!stepCol1) {
    return { ok: false, code: "STEP_COL_MISSING", msg: `PXK missing column '${colName}'` };
  }

  const pxkRow = findRow_(shPXK, pkCol1 - 1, pk, pxkStart);
  if (pxkRow === -1) {
    return { ok: false, code: "PK_NOT_FOUND", msg: "prodKey not found in current PXK" };
  }

  const cell = shPXK.getRange(pxkRow, stepCol1);

  if (String(cell.getDisplayValue() || "").trim() !== "") {
    return { ok: false, code: "SKIP", msg: "Đã quét", row: pxkRow, col: stepCol1 };
  }

  const now = new Date();
  cell.setValue(now);
  cell.setNumberFormat(CFG.DATE_FMT);
  SpreadsheetApp.flush();

  const written = String(cell.getDisplayValue() || "").trim();
  if (!written) {
    return {
      ok: false,
      code: "PXK_WRITE_FAILED",
      msg: `Write failed for ${colName} in PXK`,
      row: pxkRow,
      col: stepCol1
    };
  }

  return {
    ok: true,
    code: "OK",
    msg: `OK`,
    row: pxkRow,
    col: stepCol1,
    value: written
  };
}

function getStepStatusFromPXK_(ss, prodKey) {
  const shPXK = ss.getSheetByName(CFG.SHEET_PXK);
  if (!shPXK) return null;

  const pxkStart = findStartRowAfterHeader_(shPXK, CFG.PK_HEADER, CFG.COL_PRODKEY);
  if (pxkStart === -1) return null;
  const hr = pxkStart - 1;

  const lastCol = shPXK.getLastColumn();
  const headers = shPXK.getRange(hr, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v || "").trim());
  const pCol = makeHeaderIndexMap_(headers);

  const pkCol1 = findFirstHeaderCol_(pCol, [CFG.PK_HEADER]);
  if (!pkCol1) return null;

  const row = findRow_(shPXK, pkCol1 - 1, prodKey, pxkStart);
  if (row === -1) return null;

  const steps = (CFG.STEPS && CFG.STEPS.length) ? CFG.STEPS.slice() : [];
  const stepCols = steps
    .map(s => ({ s, col1: findFirstHeaderCol_(pCol, [CFG.STEP_PREFIX + s]) }))
    .filter(it => it.col1); // chỉ giữ step có cột

  // read row once
  const rowVals = shPXK.getRange(row, 1, 1, lastCol).getDisplayValues()[0];

  const done = {};
  for (const it of stepCols) {
    done[it.s] = String(rowVals[it.col1 - 1] || "").trim() !== "";
  }

  const doneBadges = [];
  for (const it of stepCols) {
    if (done[it.s]) doneBadges.push((STEP_LABEL[it.s] || it.s));
  }
  return { steps: stepCols.map(x => x.s), done, doneBadges };
}

/**
 * Read temField1..5 + temSTT from DOT sheet by prodKey.
 * Supports DOT batch naming DOT_1, DOT_2,...
 */
function getTemFromDotByprodKey_(ss, prodKey) {
  const pk = String(prodKey || "").trim();
  const blank = { f1:"", f2:"", f3:"", f4:"", f5:"", f6:{ stt:"", khgh:"" } };
  if (!pk) return blank;

  const shDot = getCurrentDotSheet_(ss);
  if (!shDot) return blank;

  const h = getHeader_(shDot);
  const idx = index_(h);

  const need = ["temSTT","PXK_ID","prodKey","temField1","temField2","temField3","temField4","temField5"];
  for (const k of need) {
    if (idx[k] == null) return blank;
  }

  const lastRow = shDot.getLastRow();
  if (lastRow < 2) return blank;

  const vals = shDot.getRange(2, 1, lastRow - 1, h.length).getDisplayValues();
  let row = null;

  for (const rr of vals) {
    if (String(rr[idx["prodKey"]] || "").trim() === pk) {
      row = rr;
      break;
    }
  }

  if (!row) return blank;

  return {
    f1: String(row[idx["temField1"]] || ""),
    f2: String(row[idx["temField2"]] || ""),
    f3: String(row[idx["temField3"]] || ""),
    f4: String(row[idx["temField4"]] || ""),
    f5: String(row[idx["temField5"]] || ""),
    f6: { stt: String(row[idx["temSTT"]] || ""), khgh:"" }
  };
}

function getDataByprodKey_(prodKey) {
  const pk = String(prodKey || "").trim();
  if (!pk) return null;

  const ss = SpreadsheetApp.getActive();

  // 1) Step status (đã có doneBadges sẵn)
  const st = getStepStatusFromPXK_(ss, pk);
  if (!st) return null;

  // 2) TEM fields from DOT
  const tem = getTemFromDotByprodKey_(ss, pk) || { f6:{ stt:"", khgh:"" } };

  // 3) KHGH dd/MM from PXK
  const khgh = getKhghFromPXKByprodKey_(ss, pk);
  tem.f6 = tem.f6 || {};
  tem.f6.khgh = khgh || "";

  return {
    tem,
    doneBadges: st.doneBadges || []
  };
}

/*************************************
 * 9) Push Ngay_* from DATA -> PXK (no overwrite by default)
 *************************************/

function getAllDotBatches_(ss) {
  return ss.getSheets()
    .map(s => s.getName())
    .filter(n => /^DOT_\d+$/.test(n))
    .map(n => parseInt(n.replace("DOT_", ""), 10))
    .sort((a,b)=>a-b);
}

function getBatchDotSheet_(ss, batchNo) {
  const name = dotName_(batchNo);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  enforceSheetOrder_(ss);
  return sh;
}

function getBatchTemSheet_(ss, batchNo) {
  const name = temName_(batchNo);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  enforceSheetOrder_(ss);
  return sh;
}

/*************************************
 * SHEET ORDER ENFORCER
 * Desired (left->right):
 * can, in, kho, LOG, (DOT_1,TEM_1,DOT_2,TEM_2,...), PXK, (others...)
 *************************************/
function enforceSheetOrder_(ss) {
  const wantFixed = [CFG.SHEET_LOG, CFG.SHEET_QUEUE];

  // Collect DOT_n / TEM_n
  const dots = ss.getSheets()
    .map(s => s.getName())
    .filter(n => /^DOT_\d+$/.test(n))
    .map(n => parseInt(n.replace("DOT_", ""), 10))
    .sort((a,b)=>a-b);

  const pairs = [];
  for (const n of dots) {
    pairs.push(`DOT_${n}`);
    if (ss.getSheetByName(`TEM_${n}`)) pairs.push(`TEM_${n}`);
  }

  const want = [...wantFixed, ...pairs, CFG.SHEET_PXK];
  const wantSet = new Set(want);

  // Move unknown/rest to the FRONT (preserve their current order)
  const rest = ss.getSheets().filter(s => !wantSet.has(s.getName()));
  let idx = 1;
  for (const sh of rest) {
    moveSheetToIndex_(ss, sh, idx);
    idx++;
  }

  // Then move known sheets in desired order right after
  for (const name of want) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;
    moveSheetToIndex_(ss, sh, idx);
    idx++;
  }
}

function moveSheetToIndex_(ss, sh, index1based) {
  if (!ss || !sh) return;
  const sheets = ss.getSheets();
  const cur = sheets.indexOf(sh) + 1;
  if (cur === index1based) return;

  ss.setActiveSheet(sh);
  ss.moveActiveSheet(index1based);
}

/*************************************
 * QR Check-in WebApp (SERVER)
 * - URL: .../exec?prodKey=XXXX
 * - Read: DATA (status) + DOT (tem fields) + PXK (KHGH)
 * - Submit: set Ngay_STEP once (no overwrite) + LOG
 *************************************/

/** Step label shown in UI */
const STEP_LABEL = {
  GTAM: "🤎 GIẤY TẤM",
  CAN:  "💗 CÁN",
  IN:   "🩵 IN",
  CHAP: "💛 CHẠP",
  DAN:  "❤️ DÁN",
  QC:   "✏️ QC",
  KHO:  "🚚 KHO",
};

/*************** API ***************/
function toTemPrint_(tem, meta) {
  const t = tem || {};
  const f6 = t.f6 || {};

  return {
    l1: String(t.f1 || ""),
    l2: String(t.f2 || ""),
    l3: String(t.f3 || ""),
    l4: String(t.f4 || ""),
    l5: String(t.f5 || ""),
    footer: {
      khgh: String(f6.khgh || "").trim(),
      stt:  String(f6.stt  || "").trim()
    },
    meta: meta || {}
  };
}

/*************** API (NEW SPLIT) ***************/
/**
 * 1) COMMIT = chỉ ghi PXK (lock + recheck + update)
 * - Trả ok:true chỉ khi CHẮC CHẮN đã ghi thành công
 * - Không đọc DOT / tem / KHGH để phản hồi nhanh nhất
 */
function submitStepCommit(payload) {
  const prodKey = String(payload?.prodKey || "").trim();
  const pxk = String(payload?.pxk || "").trim();
  const step = String(payload?.step || "").trim();
  const workerInfo = String(payload?.workerInfo || "").trim().toLowerCase();

  const stepKey = step.toUpperCase();
  const stepLabel = STEP_LABEL[stepKey] || stepKey;

  if (!prodKey || !pxk || !step) return { ok:false, code:"BAD_INPUT", msg:"Missing input" };
  if (!workerInfo) {
    return { ok:false, code:"NO_WORKER", msg:"Vui lòng nhập Email / Tên / SĐT để biết ai báo cáo." };
  }

  const ss = SpreadsheetApp.getActive();

  const availableSteps = getAvailableStepsFromPXK_(ss);
  if (!availableSteps.includes(stepKey)) {
    return { ok:false, code:"NOT_CONFIG", msg:`Công đoạn chưa được cấu hình: ${stepLabel}\nBáo văn phòng.` };
  }

  // precheck nhanh để queue có xác suất thành công cao
  const before = getStepStatusFromPXK_(ss, prodKey);
  if (!before) return { ok:false, code:"NOT_FOUND", msg:"Mã TEM không tồn tại trong PXK" };

  if (before.done[stepKey]) {
    return {
      ok:false,
      code:"SKIP",
      msg:`Công đoạn đã được ghi nhận: ${stepLabel}`,
      doneBadges: before.doneBadges || []
    };
  }

  return enqueueCommit_({ prodKey, step: stepKey, workerInfo, pxk });
}

/**
 * 2) FETCH = lấy TEM + doneBadges (đọc DOT + KHGH + PXK status)
 * - Không ghi gì
 * - Gọi sau COMMIT để render UI (không chặn scan)
 */
function submitStepFetch(payload) {
  const prodKey = String(payload?.prodKey || "").trim();
  const pxk = String(payload?.pxk || "").trim();
  const step = String(payload?.step || "").trim();
  const workerInfo = String(payload?.workerInfo || "").trim().toLowerCase();

  const stepKey   = String(step || "").trim().toUpperCase();
  const stepLabel = STEP_LABEL[stepKey] || stepKey;

  if (!prodKey) return { ok:false, code:"BAD_INPUT", msg:"Missing prodKey" };

  const ss = SpreadsheetApp.getActive();
  const st = getStepStatusFromPXK_(ss, prodKey);
  const tem = getTemFromDotByprodKey_(ss, prodKey) || { f1:"", f2:"", f3:"", f4:"", f5:"", f6:{ stt:"", khgh:"" } };

  // ép footer lấy trực tiếp
  const khgh = getKhghFromPXKByprodKey_(ss, prodKey) || "";
  const stt  = String(tem?.f6?.stt || "").trim();

  tem.f6 = tem.f6 || {};
  tem.f6.khgh = khgh;
  tem.f6.stt  = stt;

  const temPrint = toTemPrint_(tem, {
    stepLabel,
    workerInfo,
    today: Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy"),
  });

  return {
    ok: true,
    code: "OK",
    temPrint,
    doneBadges: st?.doneBadges || []
  };
}

/* PROCESS QUEUE functions */
function installQueueTrigger() {
  const fn = "queueWorkerTick";

  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === fn) {
      ScriptApp.deleteTrigger(t);
    }
  }

  ScriptApp.newTrigger(fn)
    .timeBased()
    .everyMinutes(1)
    .create();

  SpreadsheetApp.getUi().alert("✅ Queue trigger installed: every 1 minute");
}

function queueWorkerTick() {
  processQueueBatch_(100);
}

function processQueueBatch_(limit = 20) {
  const ss = SpreadsheetApp.getActive();
  const shQ = getOrCreateQueue_(ss);
  const lastRow = shQ.getLastRow();
  if (lastRow < 2) return { ok:true, processed:0 };

  const vals = shQ.getRange(2, 1, lastRow - 1, shQ.getLastColumn()).getDisplayValues();
  let processed = 0;

  const COL = {
    TS: 1, queueId: 2, prodKey: 3, step: 4, worker: 5,
    pxk: 6, status: 7, note: 8, doneBadgesJson: 9, updatedAt: 10
  };

  const markSkip_ = (rowNo, prodKey, note) => {
    const st = getStepStatusFromPXK_(ss, prodKey);
    shQ.getRange(rowNo, COL.status).setValue("SKIP");
    shQ.getRange(rowNo, COL.doneBadgesJson).setValue(JSON.stringify(st?.doneBadges || []));
    shQ.getRange(rowNo, COL.note).setValue(note || "Bỏ qua");
    shQ.getRange(rowNo, COL.updatedAt).setValue(Utilities.formatDate(new Date(), CFG.TZ, "yyyy-MM-dd HH:mm:ss"));
  };

  const shPXK = mustGetSheet_(ss, CFG.SHEET_PXK);
  const pxkStart = findStartRowAfterHeader_(shPXK, CFG.PK_HEADER, CFG.COL_PRODKEY);
  if (pxkStart === -1) return { ok:true, processed:0 };

  const hr = pxkStart - 1;
  const lastCol = shPXK.getLastColumn();
  const headers = shPXK.getRange(hr, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v || "").trim());
  const pCol = makeHeaderIndexMap_(headers);
  const prodKeyCol1 = mustFindHeader_(pCol, [CFG.PK_HEADER], "PXK");

  const stepColMap = {};
  for (const s of CFG.STEPS) {
    const c = findFirstHeaderCol_(pCol, [CFG.STEP_PREFIX + s]);
    if (c) stepColMap[s] = c;
  }

  const prodKeyToRow = {};
  const lastPxkRow = shPXK.getLastRow();
  const pkVals = shPXK.getRange(pxkStart, prodKeyCol1, lastPxkRow - pxkStart + 1, 1).getDisplayValues();
  for (let i = 0; i < pkVals.length; i++) {
    const pk = String(pkVals[i][0] || "").trim();
    if (pk && prodKeyToRow[pk] == null) prodKeyToRow[pk] = pxkStart + i;
  }

  for (let i = 0; i < vals.length; i++) {
    if (processed >= limit) break;

    const rowNo = i + 2;
    const status = String(vals[i][COL.status - 1] || "").trim();
    if (status !== "PENDING") continue;

    const prodKey = String(vals[i][COL.prodKey - 1] || "").trim();
    const step = String(vals[i][COL.step - 1] || "").trim().toUpperCase();
    const worker = String(vals[i][COL.worker - 1] || "").trim();

    shQ.getRange(rowNo, COL.status).setValue("PROCESSING");
    shQ.getRange(rowNo, COL.updatedAt).setValue(Utilities.formatDate(new Date(), CFG.TZ, "yyyy-MM-dd HH:mm:ss"));

    try {
      const pxkRow = prodKeyToRow[prodKey] || -1;
      const stepCol1 = stepColMap[step];

      if (pxkRow === -1 || !stepCol1) {
        markSkip_(rowNo, prodKey, "Bỏ qua");
        processed++;
        continue;
      }

      const currentVal = String(shPXK.getRange(pxkRow, stepCol1).getDisplayValue() || "").trim();
      if (currentVal) {
        markSkip_(rowNo, prodKey, "Đã quét");
        processed++;
        continue;
      }

      const writeRes = updatePxkStep_(prodKey, step, worker);
      refreshPXKKhghCache();

      if (!writeRes || !writeRes.ok) {
        markSkip_(rowNo, prodKey, "Bỏ qua");
        processed++;
        continue;
      }

      const stAfter = getStepStatusFromPXK_(ss, prodKey);
      shQ.getRange(rowNo, COL.status).setValue("DONE");
      shQ.getRange(rowNo, COL.doneBadgesJson).setValue(JSON.stringify(stAfter?.doneBadges || []));
      shQ.getRange(rowNo, COL.note).setValue(writeRes.msg || "Ghi nhận thành công");
      shQ.getRange(rowNo, COL.updatedAt).setValue(Utilities.formatDate(new Date(), CFG.TZ, "yyyy-MM-dd HH:mm:ss"));
      processed++;

    } catch (e) {
      markSkip_(rowNo, prodKey, "Bỏ qua");
      processed++;
    }
  }

  return { ok:true, processed };
}

function getQueueStatusBulk_(items) {
  const ss = SpreadsheetApp.getActive();
  const shQ = getOrCreateQueue_(ss);
  const lastRow = shQ.getLastRow();
  if (lastRow < 2) return [];

  const vals = shQ.getRange(2, 1, lastRow - 1, shQ.getLastColumn()).getDisplayValues();

  const wanted = new Map();
  for (const it of (items || [])) {
    const prodKey = String(it?.prodKey || "").trim();
    const step = String(it?.step || "").trim().toUpperCase();
    const workerInfo = String(it?.workerInfo || "").trim().toLowerCase();
    if (!prodKey || !step || !workerInfo) continue;

    wanted.set(`${prodKey}||${step}||${workerInfo}`, {
      prodKey,
      step,
      workerInfo
    });
  }

  // Quan trọng:
  // Nếu có nhiều queue rows cho cùng 1 tem/công đoạn,
  // DONE phải thắng SKIP để UI không bị OK -> Đã quét.
  const rank = {
    DONE: 50,
    ERROR: 40,
    PROCESSING: 30,
    PENDING: 20,
    SKIP: 10
  };

  const best = new Map();

  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];

    const prodKey = String(row[2] || "").trim();
    const step = String(row[3] || "").trim().toUpperCase();
    const workerInfo = String(row[4] || "").trim().toLowerCase();
    const key = `${prodKey}||${step}||${workerInfo}`;

    if (!wanted.has(key)) continue;

    const status = String(row[6] || "").trim().toUpperCase();
    const note = String(row[7] || "").trim();
    const doneBadgesJson = String(row[8] || "[]").trim();

    const cur = {
      i,
      score: rank[status] || 0,
      prodKey,
      step,
      workerInfo,
      status,
      note,
      doneBadgesJson
    };

    const old = best.get(key);

    if (
      !old ||
      cur.score > old.score ||
      (cur.score === old.score && cur.i > old.i)
    ) {
      best.set(key, cur);
    }
  }

  const out = [];

  for (const rec of best.values()) {
    let doneBadges = [];
    try {
      doneBadges = JSON.parse(rec.doneBadgesJson || "[]");
    } catch(e) {
      doneBadges = [];
    }

    let temStt = "";
    let temPrint = {
      l4: rec.prodKey,
      footer: { khgh: "", stt: "" },
      meta: { today: "" }
    };

    // Chỉ DONE / ERROR mới cần đọc DOT để render full tem.
    // SKIP / PENDING / PROCESSING chỉ cần mini card cho nhanh.
    if (rec.status === "DONE" || rec.status === "ERROR") {
      const tem = getTemFromDotByprodKey_(ss, rec.prodKey) || {
        f1:"", f2:"", f3:"", f4:"", f5:"", f6:{ stt:"", khgh:"" }
      };

      const khgh = getKhghFromPXKByprodKey_(ss, rec.prodKey) || "";
      const stt = String(tem?.f6?.stt || "").trim();

      tem.f6 = tem.f6 || {};
      tem.f6.khgh = khgh;
      tem.f6.stt = stt;

      temStt = stt;

      temPrint = toTemPrint_(tem, {
        stepLabel: STEP_LABEL[rec.step] || rec.step,
        workerInfo: rec.workerInfo,
        today: Utilities.formatDate(new Date(), CFG.TZ, "dd/MM/yyyy"),
      });
    }

    out.push({
      prodKey: rec.prodKey,
      step: rec.step,
      workerInfo: rec.workerInfo,
      status: rec.status,
      note: rec.status === "SKIP" ? "" : rec.note,
      temStt,
      doneBadges: rec.status === "SKIP" ? [] : doneBadges,
      temPrint
    });
  }

  return out;
}

/*** Get KHGH from PXK by prodKey (FAST via CacheService)*/
// Return KHGH RAW text from PXK (no date formatting)
// Lookup order: exact prodKey -> baseKey fallback
function getKhghFromPXKByprodKey_(ss, prodKey) {
  const pk = String(prodKey || "").trim();
  if (!pk) return "";

  const map = getPXKKhghMapCached_(ss);

  // lookup exact PK first, fallback to baseKey
  const base = pk.split("#")[0];
  const raw = map[pk] || map[base] || "";

  return String(raw || "").trim();
}

/** scan 1 column to find exact value, return sheet row number or -1 */
function findRow_(sh, colIdx0, value, startRow=2){
  const last = sh.getLastRow();
  if (last < startRow) return -1;
  const a = sh.getRange(startRow, colIdx0 + 1, last - startRow + 1, 1).getDisplayValues();
  const target = String(value || "").trim();
  for (let i=0;i<a.length;i++){
    if (String(a[i][0] || "").trim() === target) return startRow + i;
  }
  return -1;
}

function enqueueCommit_(payload) {
  const ss = SpreadsheetApp.getActive();
  const shQ = getOrCreateQueue_(ss);

  const prodKey = String(payload?.prodKey || "").trim();
  const step = String(payload?.step || "").trim().toUpperCase();
  const worker = String(payload?.workerInfo || "").trim();
  const pxk = String(payload?.pxk || "").trim().toUpperCase();

  if (!prodKey || !step || !worker || !pxk) {
    return { ok:false, code:"BAD_INPUT", msg:"Missing fields" };
  }

  const queueId =
    "Q" +
    Utilities.formatDate(new Date(), CFG.TZ, "yyyyMMddHHmmss") +
    "_" +
    Math.floor(Math.random() * 100000);

  shQ.appendRow([
    Utilities.formatDate(new Date(), CFG.TZ, "yyyy-MM-dd HH:mm:ss"),
    queueId,
    prodKey,
    step,
    worker,
    pxk,
    "PENDING",
    "",
    "[]",
    Utilities.formatDate(new Date(), CFG.TZ, "yyyy-MM-dd HH:mm:ss")
  ]);

  return {
    ok: true,
    code: "QUEUED",
    msg: "Barcode đã vào hàng chờ",
    queueId,
    status: "PENDING"
  };
}

/*************************************
 * PXK KHGH cache helpers
 *************************************/
// Cache keys + tuning
const PXK_CACHE = {
  KEY_META: "PXK_KHGH_MAP_META_V1",    // stores {chunks, ts}
  KEY_CHUNK_PREFIX: "PXK_KHGH_MAP_C_", // chunk keys
  TTL_SEC: 6 * 60 * 60,               // 6 hours
  CHUNK_SIZE: 85000                  // safe chunk size for CacheService (~100KB limit)
};

/**
 * Return map {prodKey: KHGH_raw} from cache, building it if needed.
 * Uses ScriptCache (shared across users).
 */
function getPXKKhghMapCached_(ss) {
  const cache = CacheService.getScriptCache();

  // meta tells us how many chunks exist
  const metaStr = cache.get(PXK_CACHE.KEY_META);
  if (metaStr) {
    try {
      const meta = JSON.parse(metaStr);
      if (meta && meta.chunks) {
        const parts = [];
        for (let i = 0; i < meta.chunks; i++) {
          const s = cache.get(PXK_CACHE.KEY_CHUNK_PREFIX + i);
          if (!s) throw new Error("missing chunk");
          parts.push(s);
        }
        return JSON.parse(parts.join("")); // full JSON string
      }
    } catch (e) {
      // fallthrough -> rebuild
    }
  }

  // build fresh map and cache it
  const map = buildPXKKhghMap_(ss);
  cachePXKKhghMap_(cache, map);
  return map;
}

/**
 * Build {prodKey: KHGH_raw} by reading PXK like your buildTemFromDot logic.
 * - Find header "prodKey" in column A
 * - Header row = startRow - 1
 * - Use DISPLAY values (consistent with sheet)
 */
function buildPXKKhghMap_(ss) {
  const shPXK = ss.getSheetByName(CFG.SHEET_PXK);
  if (!shPXK) return {};

  const pxkStart = findStartRowAfterHeader_(shPXK, "prodKey", "A");
  if (pxkStart === -1) return {};

  const pxkHr = pxkStart - 1;
  const lastCol = shPXK.getLastColumn();

  const headers = shPXK.getRange(pxkHr, 1, 1, lastCol).getDisplayValues()[0]
    .map(v => String(v || "").trim());

  const pkCol = headers.indexOf("prodKey") + 1;
  const khghCol = headers.indexOf("KHGH") + 1;
  if (!pkCol || !khghCol) return {};

  const lastRow = shPXK.getLastRow();
  if (lastRow < pxkStart) return {};

  const n = lastRow - pxkStart + 1;
  const vals = shPXK.getRange(pxkStart, 1, n, lastCol).getDisplayValues();

  const out = {};
  for (const r of vals) {
    const pk = String(r[pkCol - 1] || "").trim();
    const khgh = String(r[khghCol - 1] || "").trim();
    // keep first occurrence only (same as your PXK map logic)
    if (pk && khgh && out[pk] == null) out[pk] = khgh;
  }
  return out;
}

/**
 * Store big JSON map into CacheService in chunks.
 */
function cachePXKKhghMap_(cache, mapObj) {
  // stringify once
  const json = JSON.stringify(mapObj || {});
  const chunks = [];
  for (let i = 0; i < json.length; i += PXK_CACHE.CHUNK_SIZE) {
    chunks.push(json.slice(i, i + PXK_CACHE.CHUNK_SIZE));
  }

  // write chunks
  for (let i = 0; i < chunks.length; i++) {
    cache.put(PXK_CACHE.KEY_CHUNK_PREFIX + i, chunks[i], PXK_CACHE.TTL_SEC);
  }

  // meta for reconstruction
  const meta = JSON.stringify({ chunks: chunks.length, ts: Date.now() });
  cache.put(PXK_CACHE.KEY_META, meta, PXK_CACHE.TTL_SEC);
}

/**
 * Manual refresh if you want (optional).
 * Call this from editor when PXK changed a lot.
 */
function refreshPXKKhghCache() {
  const ss = SpreadsheetApp.getActive();
  const cache = CacheService.getScriptCache();

  // best-effort clear old meta; chunks will expire anyway
  cache.remove(PXK_CACHE.KEY_META);

  const map = buildPXKKhghMap_(ss);
  cachePXKKhghMap_(cache, map);

  return `PXK KHGH cache refreshed: ${Object.keys(map).length} keys`;
}

/**
 * Decide which DOT to read:
 * - Prefer DOT_<currentBatchNo> if function exists & sheet exists
 */
function getCurrentDotSheet_(ss) {
  if (!ss) ss = SpreadsheetApp.getActive();

  const n = getCurrentBatchNo_();
  const nm = dotName_(n);
  const sh = ss.getSheetByName(nm);

  // không fallback sheet DOT
  return sh || null;
}

/***************
 * EXTERNAL API (for hosted UI)
 * - POST JSON: { token, action: "meta"|"commit"|"fetchTem", ... }
 * - GET: health check
 ***************/
function getApiToken_() {
  // LƯU token ở Script Properties để copy sheet/project qua công ty mới không bị lộ trong code
  const p = PropertiesService.getScriptProperties();
  const t = String(p.getProperty("API_TOKEN") || "").trim();
  return t;
}

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const action = String(p.action || "").trim();

  // Health check (still JSON to make client happy)
  if (!action) return json_({ ok: true, msg: "API OK" });

  // 1) META / GET STEPS (HTML cũ gọi getSteps)
  if (action === "meta" || action === "getSteps") {
    const ss = SpreadsheetApp.getActive();
    const steps = getAvailableStepsFromPXK_(ss) || [];
    // trả về dạng [{id,label}] để UI render đẹp
    return json_({
      ok: true,
      steps: steps.map(s => ({ id: s, label: (STEP_LABEL[s] || s) }))
    });
  }

  // 2) COMMIT (GET) - để GitHub Pages gọi không bị CORS preflight
  if (action === "commit") {
    const prodKey = String(p.prodKey || "").trim();
    const step = String(p.step || "").trim();
    const workerInfo = String(p.workerInfo || "").trim();
    const pxk = normalizePxkId_(p.pxk || "");

    if (!prodKey || !step || !workerInfo || !pxk) {
      return json_({ ok: false, code: "BAD_INPUT", msg: "Missing fields" });
    }
    const c = submitStepCommit({ prodKey, step, workerInfo, pxk });
    return json_(c);
  }

  // 3) FETCH (HTML cũ gọi fetch; code server mới gọi fetchTem)
  if (action === "fetch" || action === "fetchTem") {
    const prodKey = String(p.prodKey || "").trim();
    const step = String(p.step || "").trim();
    const workerInfo = String(p.workerInfo || "").trim();
    const pxk = normalizePxkId_(p.pxk || "");

    if (!prodKey || !pxk) return json_({ ok: false, code: "BAD_INPUT", msg: "Missing prodKey or pxk" });
    const f = submitStepFetch({ prodKey, step, workerInfo, pxk });
    return json_(f);
  }

  // 4) CHECK PXK_ID
  if (action === "checkPxk") {
    const ss = SpreadsheetApp.getActive();
    const pxk = normalizePxkId_(p.pxk || "");
    if (!pxk) return json_({ ok: false, code: "BAD_INPUT", msg: "Missing pxk" });

    const st = getPxkOpenStatus_(ss, pxk);
    return json_({
      ok: st.ok,
      code: st.code,
      msg: st.msg,
      found: st.found,
      openCount: st.openCount,
      pxk: pxk,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheetName: CFG.SHEET_PXK,
      version: "PXK_OPEN_BY_KHO_V1"
    });
  }

  // 5) BULK STATUS FOR LIVE QUEUE
  if (action === "queueStatusBulk") {
    let items = [];
    try {
      items = JSON.parse(String(p.items || "[]"));
    } catch(e) {
      return json_({ ok:false, code:"BAD_INPUT", msg:"Invalid items json" });
    }

    // opportunistic processing nhẹ để queue tiến lên ngay cả khi chưa có trigger
    processQueueBatch_(10);

    return json_({
      ok: true,
      items: getQueueStatusBulk_(items)
    });
  }

  return json_({ ok: false, msg: "Unknown action" });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) ? e.postData.contents : "{}");

    const token = String(body.token || "").trim();
    const expect = getApiToken_();

    if (!expect) return json_({ ok:false, msg:"Missing API_TOKEN in Script Properties" });
    if (token !== expect) return json_({ ok:false, msg:"Không có quyền truy cập" });

    const action = String(body.action || "").trim();

    // ===== 1) META =====
    if (action === "meta") {
      const ss = SpreadsheetApp.getActive();
      const steps = getAvailableStepsFromPXK_(ss); // chỉ trả step có cột Ngay_STEP thật sự tồn tại trong PXK
      return json_({
        ok: true,
        steps,
        stepLabel: STEP_LABEL || {}
      });
    }

    // ===== 2) COMMIT =====
    if (action === "commit") {
      const prodKey = String(body.prodKey || "").trim();
      const step = String(body.step || "").trim();
      const workerInfo = String(body.workerInfo || "").trim();
      const pxk = normalizePxkId_(body.pxk || "");

      if (!prodKey || !step || !workerInfo || !pxk) {
        return json_({ ok:false, msg:"Missing fields" });
      }

      const c = submitStepCommit({ prodKey, step, workerInfo, pxk });
      return json_(c);
    }

    // ===== 3) FETCH TEM/BADGES =====
    if (action === "fetchTem") {
      const prodKey = String(body.prodKey || "").trim();
      const step = String(body.step || "").trim();
      const workerInfo = String(body.workerInfo || "").trim();
      const pxk = normalizePxkId_(body.pxk || "");

      if (!prodKey || !pxk) return json_({ ok:false, msg:"Missing prodKey or pxk" });

      const f = submitStepFetch({ prodKey, step, workerInfo, pxk });
      return json_(f);
    }

  // 4) CHECK PXK_ID (OPEN / CLOSED by Ngay_KHO)
  if (action === "checkPxk") {
    const ss = SpreadsheetApp.getActive();
    const pxk = normalizePxkId_(body.pxk || "");
    if (!pxk) return json_({ ok: false, code: "BAD_INPUT", msg: "Missing pxk" });

    const st = getPxkOpenStatus_(ss, pxk);
    return json_({
      ok: st.ok,
      code: st.code,
      msg: st.msg,
      found: st.found,
      openCount: st.openCount,
      pxk: pxk,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheetName: CFG.SHEET_PXK,
      version: "PXK_OPEN_BY_KHO_V1"
    });
  }

    return json_({ ok:false, msg:"Unknown action" });

  } catch (err) {
    return json_({ ok:false, msg:String((err && err.message) || err) });
  }
}

function json_(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.SHEET_PXK) return;

    const val = String(e.range.getDisplayValue() || "");

    // CHỈ trigger khi sửa đúng 2 ô header
    if (
      val.includes("Kính gửi") ||
      val.includes("Ngày xuất hàng")
    ) {
      writePxkIdHeader_(sh);
    }

  } catch (err) {
    Logger.log(err);
  }
}
