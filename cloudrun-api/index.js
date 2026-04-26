const express = require("express");
const { google } = require("googleapis");

const app = express();
app.use(express.json({ limit: "2mb" }));

// =====================================================
// CONFIG — Cloud Run backend only
// =====================================================
const CFG = {
  SHEET_PXK: "PXK",
  SHEET_QUEUE: "QUEUE",

  PK_HEADER: "prodKey",
  PXK_ID_HEADER: "PXK_ID",

  STEP_PREFIX: "Ngay_",
  STEPS: ["GTAM", "CAN", "IN", "CHAP", "DAN", "QC", "KHO"],

  PXK_H_TENSP: "Tên Sp",
  PXK_H_KHGH: "KHGH",

  DOT_PREFIX: "DOT_",
  DOT_HEADER_ROW: 1,

  TZ: "Asia/Ho_Chi_Minh",
  DATE_FMT: "dd/MM/yyyy",

  VERSION: "CLOUD_RUN_SCAN_API_V1"
};

const STEP_LABEL = {
  GTAM: "🤎 GIẤY TẤM",
  CAN:  "💗 CÁN",
  IN:   "🩵 IN",
  CHAP: "💛 CHẠP",
  DAN:  "❤️ DÁN",
  QC:   "✏️ QC",
  KHO:  "🚚 KHO"
};

const QUEUE_HEADER = [
  "TS",
  "queueId",
  "prodKey",
  "step",
  "worker",
  "pxk",
  "status",
  "note",
  "doneBadgesJson",
  "updatedAt"
];

// =====================================================
// CORS — GitHub Pages UI
// =====================================================
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type,Authorization");
  if (req.method === "OPTIONS") return res.status(204).send("");
  next();
});

// =====================================================
// BASIC HELPERS
// =====================================================
function clean_(v) {
  return String(v ?? "").trim();
}

function upper_(v) {
  return clean_(v).toUpperCase();
}

function norm_(v) {
  return clean_(v)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/\s+/g, " ");
}

function normalizePxkId_(v) {
  return upper_(v);
}

function colLetter_(n1) {
  let s = "";
  let n = Number(n1);
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function todayVN_() {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: CFG.TZ,
    day: "2-digit",
    month: "2-digit",
    year: "numeric"
  }).formatToParts(new Date());

  const get = (type) => parts.find(p => p.type === type)?.value || "";
  return `${get("day")}/${get("month")}/${get("year")}`;
}

function nowVN_() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: CFG.TZ,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  }).formatToParts(new Date());

  const m = {};
  for (const p of parts) m[p.type] = p.value;

  return `${m.year}-${m.month}-${m.day} ${m.hour}:${m.minute}:${m.second}`;
}

function headerMap0_(headers) {
  const m = {};
  (headers || []).forEach((h, i) => {
    const raw = clean_(h);
    if (!raw) return;

    if (m[raw] == null) m[raw] = i;

    const nk = norm_(raw);
    if (nk && m[nk] == null) m[nk] = i;
  });
  return m;
}

function findHeader0_(headerMap, candidates) {
  for (const h of candidates || []) {
    const raw = clean_(h);
    if (!raw) continue;

    if (headerMap[raw] != null) return headerMap[raw];

    const nk = norm_(raw);
    if (nk && headerMap[nk] != null) return headerMap[nk];
  }
  return -1;
}

function parseProdKey_(raw) {
  const s = clean_(raw);
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

function isStopLine_(row) {
  const t = norm_((row || []).slice(0, 10).join(" "));

  return (
    t.includes("nhan san pham") ||
    t.startsWith("**") ||
    t.startsWith("* *") ||
    t.includes("nhan hang") ||
    t.includes("cam on") ||
    t.includes("mong quy khach hang quet ma qr")
  );
}

function findPxkHeader_(rows) {
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    if (norm_(row[0]) === norm_(CFG.PK_HEADER)) {
      return {
        headerRow0: r,
        header: row.map(clean_),
        hMap: headerMap0_(row.map(clean_))
      };
    }
  }
  return null;
}

function findPxkDataEndRow0_(rows, headerRow0, tenSpCol0) {
  let end0 = headerRow0;
  let blankRun = 0;

  for (let r = headerRow0 + 1; r < rows.length; r++) {
    const row = rows[r] || [];

    if (isStopLine_(row)) break;

    const nameVal = tenSpCol0 >= 0 ? clean_(row[tenSpCol0]) : "";
    if (!nameVal) blankRun++;
    else blankRun = 0;

    if (blankRun >= 20) break;
    end0 = r;
  }

  return end0;
}

function getDoneBadges_(row, hMap) {
  const out = [];

  for (const step of CFG.STEPS) {
    const col0 = findHeader0_(hMap, [`${CFG.STEP_PREFIX}${step}`]);
    if (col0 < 0) continue;

    const val = clean_(row[col0]);
    if (val) out.push(STEP_LABEL[step] || step);
  }

  return out;
}

function jsonOk_(res, obj = {}) {
  return res.json({ ok: true, ...obj });
}

function jsonFail_(res, obj = {}, status = 200) {
  return res.status(status).json({ ok: false, ...obj });
}

function parseItems_(raw) {
  if (Array.isArray(raw)) return raw;

  const s = clean_(raw);
  if (!s) return [];

  try {
    const parsed = JSON.parse(s);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

// =====================================================
// GOOGLE SHEETS CLIENT
// =====================================================
async function getSheetsClient_() {
  const auth = new google.auth.GoogleAuth({
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive.metadata.readonly"
    ]
  });

  return google.sheets({ version: "v4", auth });
}

function qSheet_(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

async function getSpreadsheetMeta_(sheets, sid) {
  const result = await sheets.spreadsheets.get({
    spreadsheetId: sid,
    fields: "spreadsheetId,properties.title,sheets.properties(title,sheetId)"
  });

  return {
    spreadsheetId: result.data.spreadsheetId,
    title: result.data.properties.title,
    sheets: (result.data.sheets || []).map(s => ({
      title: s.properties.title,
      sheetId: s.properties.sheetId
    }))
  };
}

async function listSheetNames_(sheets, sid) {
  const meta = await getSpreadsheetMeta_(sheets, sid);
  return meta.sheets.map(s => s.title);
}

async function sheetExists_(sheets, sid, sheetName) {
  const names = await listSheetNames_(sheets, sid);
  return names.includes(sheetName);
}

async function readRange_(sheets, sid, range) {
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: sid,
    range
  });

  return result.data.values || [];
}

async function readPxkRows_(sheets, sid) {
  return readRange_(sheets, sid, `${qSheet_(CFG.SHEET_PXK)}!A1:ZZ5000`);
}

async function updateCell_(sheets, sid, sheetName, row1, col1, value) {
  const a1 = `${qSheet_(sheetName)}!${colLetter_(col1)}${row1}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId: sid,
    range: a1,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[value]]
    }
  });
}

async function updateRowRange_(sheets, sid, sheetName, row1, startCol1, values) {
  const start = `${colLetter_(startCol1)}${row1}`;
  const end = `${colLetter_(startCol1 + values.length - 1)}${row1}`;
  const range = `${qSheet_(sheetName)}!${start}:${end}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId: sid,
    range,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [values]
    }
  });
}

async function appendRow_(sheets, sid, sheetName, values) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: sid,
    range: `${qSheet_(sheetName)}!A:Z`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [values]
    }
  });
}

async function ensureSheet_(sheets, sid, sheetName, header) {
  const exists = await sheetExists_(sheets, sid, sheetName);

  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sid,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: sheetName
              }
            }
          }
        ]
      }
    });
  }

  if (header && header.length) {
    const vals = await readRange_(sheets, sid, `${qSheet_(sheetName)}!A1:${colLetter_(header.length)}1`);
    const first = vals[0] || [];
    const needHeader = clean_(first[0]) !== clean_(header[0]);

    if (needHeader) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: sid,
        range: `${qSheet_(sheetName)}!A1:${colLetter_(header.length)}1`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [header]
        }
      });
    }
  }
}

async function ensureQueue_(sheets, sid) {
  await ensureSheet_(sheets, sid, CFG.SHEET_QUEUE, QUEUE_HEADER);
}

// =====================================================
// PXK STATUS / LOOKUP
// =====================================================
function getPxkIdFromRows_(rows) {
  const label = clean_(rows?.[4]?.[0]);
  const value = clean_(rows?.[5]?.[0]);

  if (label === CFG.PXK_ID_HEADER && value) return value;
  return "";
}

function getPxkRowStatusFromRows_(rows, prodKey) {
  const found = findPxkHeader_(rows);
  if (!found) return null;

  const tenSpCol0 = findHeader0_(found.hMap, [CFG.PXK_H_TENSP, "Ten Sp"]);
  const end0 = findPxkDataEndRow0_(rows, found.headerRow0, tenSpCol0);

  for (let r = found.headerRow0 + 1; r <= end0; r++) {
    const row = rows[r] || [];
    const pk = clean_(row[0]);

    if (pk === prodKey) {
      return {
        row0: r,
        row1: r + 1,
        row,
        headerRow0: found.headerRow0,
        header: found.header,
        hMap: found.hMap,
        doneBadges: getDoneBadges_(row, found.hMap)
      };
    }
  }

  return null;
}

function getPxkOpenStatusFromRows_(rows, pxkId) {
  const want = normalizePxkId_(pxkId);
  if (!want) {
    return { ok: false, code: "BAD_INPUT", msg: "Missing pxk" };
  }

  const currentPxk = normalizePxkId_(getPxkIdFromRows_(rows));
  if (!currentPxk) {
    return { ok: false, code: "PXK_ID_MISSING", msg: "PXK_ID chưa được tạo ở A6" };
  }

  if (currentPxk !== want) {
    return {
      ok: false,
      code: "PXK_NOT_FOUND",
      msg: "PXK không tồn tại",
      found: 0,
      openCount: 0
    };
  }

  const found = findPxkHeader_(rows);
  if (!found) {
    return {
      ok: false,
      code: "PXK_HEADER_NOT_FOUND",
      msg: "PXK: cannot find prodKey header.",
      found: 0,
      openCount: 0
    };
  }

  const khoCol0 = findHeader0_(found.hMap, [`${CFG.STEP_PREFIX}KHO`]);
  if (khoCol0 < 0) {
    return {
      ok: false,
      code: "KHO_COL_MISSING",
      msg: "PXK missing column 'Ngay_KHO'",
      found: 0,
      openCount: 0
    };
  }

  const tenSpCol0 = findHeader0_(found.hMap, [CFG.PXK_H_TENSP, "Ten Sp"]);
  const end0 = findPxkDataEndRow0_(rows, found.headerRow0, tenSpCol0);

  if (end0 <= found.headerRow0) {
    return {
      ok: false,
      code: "NO_DATA",
      msg: "PXK has no data rows",
      found: 0,
      openCount: 0
    };
  }

  let foundCount = 0;
  let openCount = 0;

  for (let r = found.headerRow0 + 1; r <= end0; r++) {
    const row = rows[r] || [];
    const pk = clean_(row[0]);
    if (!pk) continue;

    foundCount++;
    const khoVal = clean_(row[khoCol0]);
    if (!khoVal) openCount++;
  }

  if (!foundCount) {
    return {
      ok: false,
      code: "PXK_NOT_FOUND",
      msg: "PXK không tồn tại",
      found: 0,
      openCount: 0
    };
  }

  if (openCount <= 0) {
    return {
      ok: false,
      code: "PXK_CLOSED",
      msg: "PXK này đã hoàn tất nhập kho",
      found: foundCount,
      openCount: 0
    };
  }

  return {
    ok: true,
    code: "OPEN",
    msg: "PXK này đang hoạt động",
    found: foundCount,
    openCount
  };
}

// =====================================================
// STEPS
// =====================================================
async function getAvailableSteps_(sheets, sid) {
  const rows = await readPxkRows_(sheets, sid);
  const found = findPxkHeader_(rows);
  if (!found) return [];

  const steps = [];

  for (const step of CFG.STEPS) {
    const col0 = findHeader0_(found.hMap, [`${CFG.STEP_PREFIX}${step}`]);
    if (col0 >= 0) {
      steps.push({
        id: step,
        label: STEP_LABEL[step] || step
      });
    }
  }

  return steps;
}

// =====================================================
// COMMIT
// =====================================================
async function submitStepCommit_(sheets, sid, payload) {
  const prodKey = clean_(payload.prodKey);
  const step = upper_(payload.step);
  const workerInfo = clean_(payload.workerInfo);
  const pxk = normalizePxkId_(payload.pxk);

  const stepLabel = STEP_LABEL[step] || step;

  if (!prodKey || !step || !workerInfo || !pxk) {
    return { ok: false, code: "BAD_INPUT", msg: "Missing fields" };
  }

  const parsed = parseProdKey_(prodKey);
  if (parsed.pxkId && normalizePxkId_(parsed.pxkId) !== pxk) {
    return {
      ok: false,
      code: "PXK_MISMATCH",
      msg: "Mã TEM không thuộc PXK đang mở"
    };
  }

  const rows = await readPxkRows_(sheets, sid);

  const currentPxk = normalizePxkId_(getPxkIdFromRows_(rows));
  if (currentPxk && currentPxk !== pxk) {
    return {
      ok: false,
      code: "PXK_MISMATCH",
      msg: "PXK đang mở không khớp file hiện tại"
    };
  }

  const st = getPxkRowStatusFromRows_(rows, prodKey);
  if (!st) {
    return {
      ok: false,
      code: "NOT_FOUND",
      msg: "Mã TEM không tồn tại trong PXK"
    };
  }

  const stepCol0 = findHeader0_(st.hMap, [`${CFG.STEP_PREFIX}${step}`]);
  if (stepCol0 < 0) {
    return {
      ok: false,
      code: "NOT_CONFIG",
      msg: `Công đoạn chưa được cấu hình: ${stepLabel}\nBáo văn phòng.`
    };
  }

  const currentVal = clean_(st.row[stepCol0]);
  if (currentVal) {
    return {
      ok: false,
      code: "SKIP",
      msg: `Công đoạn đã được ghi nhận: ${stepLabel}`,
      doneBadges: st.doneBadges || []
    };
  }

  await updateCell_(sheets, sid, CFG.SHEET_PXK, st.row1, stepCol0 + 1, todayVN_());

  const updatedRow = [...st.row];
  updatedRow[stepCol0] = todayVN_();

  return {
    ok: true,
    code: "OK",
    status: "DONE",
    msg: `Đã ghi nhận: ${stepLabel}`,
    doneBadges: getDoneBadges_(updatedRow, st.hMap)
  };
}

// =====================================================
// DOT / TEM FETCH
// =====================================================
function pickDotNamesDesc_(sheetNames) {
  return (sheetNames || [])
    .filter(n => /^DOT_\d+$/i.test(n))
    .map(n => ({
      name: n,
      no: parseInt(String(n).replace(/^DOT_/i, ""), 10)
    }))
    .filter(x => Number.isInteger(x.no))
    .sort((a, b) => b.no - a.no)
    .map(x => x.name);
}

async function getTemFromDotByProdKey_(sheets, sid, prodKey) {
  const sheetNames = await listSheetNames_(sheets, sid);
  const dotNames = pickDotNamesDesc_(sheetNames);

  for (const dotName of dotNames) {
    const rows = await readRange_(sheets, sid, `${qSheet_(dotName)}!A1:ZZ5000`);
    if (!rows.length) continue;

    const header = (rows[0] || []).map(clean_);
    const hMap = headerMap0_(header);

    const colPK = findHeader0_(hMap, ["prodKey"]);
    if (colPK < 0) continue;

    const c = {
      stt: findHeader0_(hMap, ["temSTT"]),
      f1: findHeader0_(hMap, ["temField1"]),
      f2: findHeader0_(hMap, ["temField2"]),
      f3: findHeader0_(hMap, ["temField3"]),
      f4: findHeader0_(hMap, ["temField4"]),
      f5: findHeader0_(hMap, ["temField5"])
    };

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] || [];
      if (clean_(row[colPK]) !== prodKey) continue;

      return {
        f1: c.f1 >= 0 ? clean_(row[c.f1]) : "",
        f2: c.f2 >= 0 ? clean_(row[c.f2]) : "",
        f3: c.f3 >= 0 ? clean_(row[c.f3]) : "",
        f4: c.f4 >= 0 ? clean_(row[c.f4]) : "",
        f5: c.f5 >= 0 ? clean_(row[c.f5]) : "",
        f6: {
          stt: c.stt >= 0 ? clean_(row[c.stt]) : "",
          khgh: ""
        },
        dotName
      };
    }
  }

  return null;
}

function buildKhghMapFromPxkRows_(rows) {
  const map = {};
  const found = findPxkHeader_(rows);
  if (!found) return map;

  const khghCol0 = findHeader0_(found.hMap, [CFG.PXK_H_KHGH, "Khgh"]);
  if (khghCol0 < 0) return map;

  const tenSpCol0 = findHeader0_(found.hMap, [CFG.PXK_H_TENSP, "Ten Sp"]);
  const end0 = findPxkDataEndRow0_(rows, found.headerRow0, tenSpCol0);

  for (let r = found.headerRow0 + 1; r <= end0; r++) {
    const row = rows[r] || [];
    const pk = clean_(row[0]);
    const khgh = clean_(row[khghCol0]);

    if (!pk || !khgh) continue;

    const base = pk.split("#")[0];

    if (map[pk] == null) map[pk] = khgh;
    if (base && map[base] == null) map[base] = khgh;
  }

  return map;
}

function getKhghFromPxkRows_(rows, prodKey) {
  const map = buildKhghMapFromPxkRows_(rows);
  const pk = clean_(prodKey);
  const base = pk.split("#")[0];

  return clean_(map[pk] || map[base] || "");
}

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
      khgh: clean_(f6.khgh),
      stt: clean_(f6.stt)
    },
    meta: meta || {}
  };
}

async function submitStepFetch_(sheets, sid, payload) {
  const prodKey = clean_(payload.prodKey);
  const stepKey = upper_(payload.step);
  const workerInfo = clean_(payload.workerInfo);
  const pxk = normalizePxkId_(payload.pxk);

  const stepLabel = STEP_LABEL[stepKey] || stepKey;

  if (!prodKey || !pxk) {
    return { ok: false, code: "BAD_INPUT", msg: "Missing prodKey or pxk" };
  }

  const rows = await readPxkRows_(sheets, sid);
  const currentPxk = normalizePxkId_(getPxkIdFromRows_(rows));

  if (currentPxk && currentPxk !== pxk) {
    return {
      ok: false,
      code: "PXK_MISMATCH",
      msg: "PXK đang mở không khớp file hiện tại"
    };
  }

  const st = getPxkRowStatusFromRows_(rows, prodKey);
  const tem = await getTemFromDotByProdKey_(sheets, sid, prodKey) || {
    f1: "",
    f2: "",
    f3: "",
    f4: "",
    f5: "",
    f6: { stt: "", khgh: "" }
  };

  const khgh = getKhghFromPxkRows_(rows, prodKey) || "";
  const stt = clean_(tem?.f6?.stt);

  tem.f6 = tem.f6 || {};
  tem.f6.khgh = khgh;
  tem.f6.stt = stt;

  const temPrint = toTemPrint_(tem, {
    stepLabel,
    workerInfo,
    today: todayVN_()
  });

  return {
    ok: true,
    code: "OK",
    temPrint,
    doneBadges: st?.doneBadges || []
  };
}

// =====================================================
// QUEUE SHEET
// =====================================================
function makeQueueId_() {
  const stamp = nowVN_().replace(/[-:\s]/g, "");
  const rand = Math.floor(Math.random() * 100000);
  return `Q${stamp}_${rand}`;
}

async function enqueueCommit_(sheets, sid, payload) {
  const prodKey = clean_(payload?.prodKey);
  const step = upper_(payload?.step);
  const worker = clean_(payload?.workerInfo);
  const pxk = normalizePxkId_(payload?.pxk);

  if (!prodKey || !step || !worker || !pxk) {
    return { ok: false, code: "BAD_INPUT", msg: "Missing fields" };
  }

  await ensureQueue_(sheets, sid);

  const queueId = makeQueueId_();

  await appendRow_(sheets, sid, CFG.SHEET_QUEUE, [
    nowVN_(),
    queueId,
    prodKey,
    step,
    worker,
    pxk,
    "PENDING",
    "",
    "[]",
    nowVN_()
  ]);

  return {
    ok: true,
    code: "QUEUED",
    msg: "Barcode đã vào hàng chờ",
    queueId,
    status: "PENDING"
  };
}

async function readQueueRows_(sheets, sid) {
  await ensureQueue_(sheets, sid);
  return readRange_(sheets, sid, `${qSheet_(CFG.SHEET_QUEUE)}!A1:J5000`);
}

function queueIndex_() {
  return {
    TS: 0,
    queueId: 1,
    prodKey: 2,
    step: 3,
    worker: 4,
    pxk: 5,
    status: 6,
    note: 7,
    doneBadgesJson: 8,
    updatedAt: 9
  };
}

async function processQueueBatch_(sheets, sid, limit = 20) {
  const qRows = await readQueueRows_(sheets, sid);
  if (qRows.length < 2) return { ok: true, processed: 0 };

  const COL = queueIndex_();
  let processed = 0;

  for (let i = 1; i < qRows.length; i++) {
    if (processed >= limit) break;

    const row = qRows[i] || [];
    const row1 = i + 1;

    const status = upper_(row[COL.status]);
    if (status && status !== "PENDING" && status !== "PROCESSING") continue;

    const prodKey = clean_(row[COL.prodKey]);
    const step = upper_(row[COL.step]);
    const workerInfo = clean_(row[COL.worker]);
    const pxk = normalizePxkId_(row[COL.pxk]);

    if (!prodKey || !step || !workerInfo || !pxk) {
      await updateRowRange_(sheets, sid, CFG.SHEET_QUEUE, row1, 7, [
        "ERROR",
        "Missing fields",
        "[]",
        nowVN_()
      ]);
      processed++;
      continue;
    }

    try {
      await updateCell_(sheets, sid, CFG.SHEET_QUEUE, row1, 7, "PROCESSING");
      await updateCell_(sheets, sid, CFG.SHEET_QUEUE, row1, 10, nowVN_());

      const out = await submitStepCommit_(sheets, sid, {
        prodKey,
        step,
        workerInfo,
        pxk
      });

      if (out.ok) {
        await updateRowRange_(sheets, sid, CFG.SHEET_QUEUE, row1, 7, [
          "DONE",
          "",
          JSON.stringify(out.doneBadges || []),
          nowVN_()
        ]);
      } else if (out.code === "SKIP") {
        await updateRowRange_(sheets, sid, CFG.SHEET_QUEUE, row1, 7, [
          "SKIP",
          "",
          JSON.stringify(out.doneBadges || []),
          nowVN_()
        ]);
      } else {
        await updateRowRange_(sheets, sid, CFG.SHEET_QUEUE, row1, 7, [
          "ERROR",
          out.msg || out.error || "Lỗi xử lý",
          JSON.stringify(out.doneBadges || []),
          nowVN_()
        ]);
      }
    } catch (err) {
      await updateRowRange_(sheets, sid, CFG.SHEET_QUEUE, row1, 7, [
        "ERROR",
        err.message || String(err),
        "[]",
        nowVN_()
      ]);
    }

    processed++;
  }

  return { ok: true, processed };
}

function buildQueueLookup_(qRows) {
  const COL = queueIndex_();
  const map = new Map();

  for (let i = 1; i < qRows.length; i++) {
    const row = qRows[i] || [];
    const prodKey = clean_(row[COL.prodKey]);
    if (!prodKey) continue;

    // Keep latest occurrence.
    map.set(prodKey, {
      prodKey,
      step: upper_(row[COL.step]),
      workerInfo: clean_(row[COL.worker]),
      pxk: normalizePxkId_(row[COL.pxk]),
      status: upper_(row[COL.status]) || "PENDING",
      note: clean_(row[COL.note]),
      doneBadgesJson: clean_(row[COL.doneBadgesJson])
    });
  }

  return map;
}

async function getQueueStatusBulk_(sheets, sid, items) {
  const inputItems = Array.isArray(items) ? items : [];
  const qRows = await readQueueRows_(sheets, sid);
  const qMap = buildQueueLookup_(qRows);

  const out = [];

  for (const it of inputItems) {
    const prodKey = clean_(it?.prodKey || it?.pk || it);
    const reqStep = upper_(it?.step);
    const reqWorker = clean_(it?.workerInfo);

    if (!prodKey) continue;

    const rec = qMap.get(prodKey) || {
      prodKey,
      step: reqStep,
      workerInfo: reqWorker,
      pxk: "",
      status: "PENDING",
      note: "",
      doneBadgesJson: "[]"
    };

    let doneBadges = [];
    try {
      doneBadges = JSON.parse(rec.doneBadgesJson || "[]");
      if (!Array.isArray(doneBadges)) doneBadges = [];
    } catch (e) {
      doneBadges = [];
    }

    let temPrint = null;
    let temStt = "";

    if (["DONE", "SKIP", "ERROR", "PROCESSING", "PENDING"].includes(rec.status)) {
      const f = await submitStepFetch_(sheets, sid, {
        prodKey,
        step: rec.step || reqStep,
        workerInfo: rec.workerInfo || reqWorker,
        pxk: rec.pxk || it?.pxk || ""
      });

      if (f && f.ok) {
        temPrint = f.temPrint || null;
        temStt = clean_(f.temPrint?.footer?.stt);
        if (!doneBadges.length && rec.status !== "SKIP") {
          doneBadges = f.doneBadges || [];
        }
      }
    }

    out.push({
      prodKey,
      step: rec.step || reqStep,
      workerInfo: rec.workerInfo || reqWorker,
      status: rec.status,
      note: rec.status === "SKIP" ? "" : rec.note,
      temStt,
      doneBadges: rec.status === "SKIP" ? [] : doneBadges,
      temPrint
    });
  }

  return out;
}

// =====================================================
// ROUTER
// =====================================================
function getPayload_(req) {
  return {
    ...(req.query || {}),
    ...(req.body || {})
  };
}

async function handleRequest_(req, res) {
  try {
    const p = getPayload_(req);
    const action = clean_(p.action);
    const sid = clean_(p.sid);

    if (!action) {
      return jsonOk_(res, {
        service: "TM Barcode API",
        message: "Cloud Run backend is running",
        version: CFG.VERSION
      });
    }

    if (action === "health") {
      return jsonOk_(res, {
        service: "TM Barcode API",
        message: "OK",
        version: CFG.VERSION
      });
    }

    if (!sid) {
      return jsonFail_(res, {
        code: "MISSING_SID",
        msg: "Missing sid"
      }, 400);
    }

    const sheets = await getSheetsClient_();

    if (action === "sheet-test") {
      const meta = await getSpreadsheetMeta_(sheets, sid);
      return jsonOk_(res, meta);
    }

    if (action === "meta" || action === "getSteps") {
      const steps = await getAvailableSteps_(sheets, sid);
      return jsonOk_(res, {
        steps,
        stepLabel: STEP_LABEL
      });
    }

    if (action === "checkPxk") {
      const pxk = normalizePxkId_(p.pxk);
      if (!pxk) {
        return jsonFail_(res, {
          code: "BAD_INPUT",
          msg: "Missing pxk"
        });
      }

      const rows = await readPxkRows_(sheets, sid);
      const st = getPxkOpenStatusFromRows_(rows, pxk);
      const meta = await getSpreadsheetMeta_(sheets, sid);

      return res.json({
        ok: st.ok,
        code: st.code,
        msg: st.msg,
        found: st.found,
        openCount: st.openCount,
        pxk,
        spreadsheetId: meta.spreadsheetId,
        spreadsheetName: meta.title,
        sheetName: CFG.SHEET_PXK,
        version: "PXK_OPEN_BY_KHO_V1"
      });
    }

    // COMMIT = enqueue first, then immediately process a small queue batch
    // This keeps QUEUE logic but makes UI receive DONE/SKIP faster
    if (action === "commit") {
      const out = await enqueueCommit_(sheets, sid, {
        prodKey: p.prodKey,
        step: p.step,
        workerInfo: p.workerInfo,
        pxk: p.pxk
      });

      if (out && out.ok) {
        try {
          await processQueueBatch_(sheets, sid, 5);
        } catch (e) {
          // Do not fail the scan response if immediate queue processing has an issue.
          // queueStatusBulk will process again on next poll.
        }
      }

      return res.json(out);
    }

    // Fetch TEM/card data.
    if (action === "fetch" || action === "fetchTem") {
      const out = await submitStepFetch_(sheets, sid, {
        prodKey: p.prodKey,
        step: p.step,
        workerInfo: p.workerInfo,
        pxk: p.pxk
      });
      return res.json(out);
    }

    // Optional queue insert API. Keep for UI versions using live queue.
    if (action === "enqueue" || action === "queueCommit") {
      const out = await enqueueCommit_(sheets, sid, {
        prodKey: p.prodKey,
        step: p.step,
        workerInfo: p.workerInfo,
        pxk: p.pxk
      });
      return res.json(out);
    }

    // Live queue bulk status. Same intent as old Apps Script:
    // process a small batch first, then return status for requested items
    if (action === "queueStatusBulk") {
      let items = parseItems_(p.items);

      if (!items.length && p.prodKeys) {
        items = clean_(p.prodKeys)
          .split(",")
          .map(s => s.trim())
          .filter(Boolean)
          .map(prodKey => ({
            prodKey,
            step: p.step || "",
            workerInfo: p.workerInfo || "",
            pxk: p.pxk || ""
          }));
      }

      if (!items.length) {
        return jsonFail_(res, {
          code: "BAD_INPUT",
          msg: "Invalid items json"
        });
      }

      await processQueueBatch_(sheets, sid, 10);

      const rows = await getQueueStatusBulk_(sheets, sid, items);
      return jsonOk_(res, { items: rows });
    }

    return jsonFail_(res, {
      code: "UNKNOWN_ACTION",
      msg: `Unknown action: ${action}`
    }, 404);

  } catch (err) {
    return jsonFail_(res, {
      code: err.code || "SERVER_ERROR",
      error: err.message,
      msg: err.message
    }, 500);
  }
}

app.get("/", handleRequest_);
app.post("/", handleRequest_);

app.get("/health", (req, res) => {
  res.json({
    ok: true,
    service: "TM Barcode API",
    version: CFG.VERSION
  });
});

app.get("/sheet-test", async (req, res) => {
  try {
    const sid = clean_(req.query.sid);
    if (!sid) {
      return jsonFail_(res, {
        code: "MISSING_SID",
        error: "Missing sid"
      }, 400);
    }

    const sheets = await getSheetsClient_();
    const meta = await getSpreadsheetMeta_(sheets, sid);
    return jsonOk_(res, meta);
  } catch (err) {
    return jsonFail_(res, {
      code: err.code || "SERVER_ERROR",
      error: err.message
    }, 500);
  }
});

const port = process.env.PORT || 8080;
app.listen(port, () => {
  console.log(`TM Barcode API listening on port ${port}`);
});
