"use strict";

/** -----------------------------
 * DOM
 * ----------------------------- */
const $ = (id) => document.getElementById(id);

const fileInputs = {
  current: $("file-current"),
  A: $("file-a"),
  B: $("file-b"),
  C: $("file-c"),
};

const fileListEls = {
  current: $("file-current-list"),
  A: $("file-a-list"),
  B: $("file-b-list"),
  C: $("file-c-list"),
};

const compareBody = $("compare-body");
const statusBox = $("status-box");
const logBox = $("log-box");

const btnExtract = $("btn-extract");
const btnOpenMapping = $("btn-open-mapping");
const btnCalc = $("btn-calc");
const btnExportCsv = $("btn-export-csv");
const btnReset = $("btn-reset");

const mappingModal = $("mapping-modal");
const mappingBackdrop = $("mapping-backdrop");
const btnCloseMapping = $("btn-close-mapping");
const btnApplySuggestions = $("btn-apply-suggestions");
const btnSaveMapping = $("btn-save-mapping");
const mappingBody = $("mapping-body");

/** -----------------------------
 * 상수
 * ----------------------------- */
const PROJECT_KEYS = ["current", "A", "B", "C"];

const PROJECT_LABELS = {
  current: "현재 프로젝트",
  A: "A 프로젝트",
  B: "B 프로젝트",
  C: "C 프로젝트",
};

const SECTION_DISPLAY_CLASS = {
  APT: "section-row--apt",
  PIT: "section-row--pit",
  주차장: "section-row--parking",
  부대동: "section-row--ancillary",
};

const CATEGORY_OPTIONS = [
  { value: "", label: "선택" },
  { value: "레미콘", label: "레미콘" },
  { value: "거푸집", label: "거푸집" },
  { value: "철근", label: "철근" },
  { value: "제외", label: "제외" },
];

const INCLUDE_OPTIONS = [
  { value: "Y", label: "반영" },
  { value: "N", label: "제외" },
];

const ITEM_CODE_OPTIONS = [
  "",
  "240", "270", "300", "180",
  "3회", "4회", "유로", "알폼", "갱폼", "합벽", "보밑면", "데크", "방수턱",
  "H10", "H13", "H16", "H19", "H22", "H25", "H29"
];

const COMPARE_LAYOUT = [
  { type: "section", section: "APT" },
  { section: "APT", itemCode: "240", item: "레미콘", spec: "25-24-15", category: "레미콘" },
  { section: "APT", itemCode: "270", item: "레미콘", spec: "25-27-15", category: "레미콘" },
  { section: "APT", itemCode: "300", item: "레미콘", spec: "25-30-15", category: "레미콘" },
  { section: "APT", itemCode: "180", item: "레미콘", spec: "25-18-08", category: "레미콘" },
  { section: "APT", itemCode: "3회", item: "거푸집", spec: "3회", category: "거푸집" },
  { section: "APT", itemCode: "4회", item: "거푸집", spec: "4회", category: "거푸집" },
  { section: "APT", itemCode: "유로", item: "거푸집", spec: "유로", category: "거푸집" },
  { section: "APT", itemCode: "알폼", item: "거푸집", spec: "알폼", category: "거푸집" },
  { section: "APT", itemCode: "갱폼", item: "거푸집", spec: "갱폼", category: "거푸집" },
  { section: "APT", itemCode: "합벽", item: "거푸집", spec: "합벽", category: "거푸집" },
  { section: "APT", itemCode: "보밑면", item: "거푸집", spec: "보밑면", category: "거푸집" },
  { section: "APT", itemCode: "데크", item: "거푸집", spec: "데크플레이트", category: "거푸집" },
  { section: "APT", itemCode: "방수턱", item: "거푸집", spec: "방수턱", category: "거푸집" },
  { section: "APT", itemCode: "H10", item: "철근", spec: "H10", category: "철근" },
  { section: "APT", itemCode: "H13", item: "철근", spec: "H13", category: "철근" },
  { section: "APT", itemCode: "H16", item: "철근", spec: "H16", category: "철근" },
  { section: "APT", itemCode: "H19", item: "철근", spec: "H19", category: "철근" },
  { section: "APT", itemCode: "H22", item: "철근", spec: "H22", category: "철근" },
  { section: "APT", itemCode: "H25", item: "철근", spec: "H25", category: "철근" },
  { section: "APT", itemCode: "H29", item: "철근", spec: "H29", category: "철근" },

  { type: "section", section: "PIT" },
  { section: "PIT", itemCode: "240", item: "레미콘", spec: "25-24-15", category: "레미콘" },
  { section: "PIT", itemCode: "270", item: "레미콘", spec: "25-27-15", category: "레미콘" },
  { section: "PIT", itemCode: "300", item: "레미콘", spec: "25-30-15", category: "레미콘" },
  { section: "PIT", itemCode: "180", item: "레미콘", spec: "25-18-08", category: "레미콘" },
  { section: "PIT", itemCode: "3회", item: "거푸집", spec: "3회", category: "거푸집" },
  { section: "PIT", itemCode: "4회", item: "거푸집", spec: "4회", category: "거푸집" },
  { section: "PIT", itemCode: "유로", item: "거푸집", spec: "유로", category: "거푸집" },
  { section: "PIT", itemCode: "알폼", item: "거푸집", spec: "알폼", category: "거푸집" },
  { section: "PIT", itemCode: "갱폼", item: "거푸집", spec: "갱폼", category: "거푸집" },
  { section: "PIT", itemCode: "합벽", item: "거푸집", spec: "합벽", category: "거푸집" },
  { section: "PIT", itemCode: "보밑면", item: "거푸집", spec: "보밑면", category: "거푸집" },
  { section: "PIT", itemCode: "데크", item: "거푸집", spec: "데크플레이트", category: "거푸집" },
  { section: "PIT", itemCode: "방수턱", item: "거푸집", spec: "방수턱", category: "거푸집" },
  { section: "PIT", itemCode: "H10", item: "철근", spec: "H10", category: "철근" },
  { section: "PIT", itemCode: "H13", item: "철근", spec: "H13", category: "철근" },
  { section: "PIT", itemCode: "H16", item: "철근", spec: "H16", category: "철근" },
  { section: "PIT", itemCode: "H19", item: "철근", spec: "H19", category: "철근" },
  { section: "PIT", itemCode: "H22", item: "철근", spec: "H22", category: "철근" },
  { section: "PIT", itemCode: "H25", item: "철근", spec: "H25", category: "철근" },
  { section: "PIT", itemCode: "H29", item: "철근", spec: "H29", category: "철근" },

  { type: "section", section: "주차장" },
  { section: "주차장", itemCode: "240", item: "레미콘", spec: "25-24-15", category: "레미콘" },
  { section: "주차장", itemCode: "270", item: "레미콘", spec: "25-27-15", category: "레미콘" },
  { section: "주차장", itemCode: "300", item: "레미콘", spec: "25-30-15", category: "레미콘" },
  { section: "주차장", itemCode: "180", item: "레미콘", spec: "25-18-08", category: "레미콘" },
  { section: "주차장", itemCode: "3회", item: "거푸집", spec: "3회", category: "거푸집" },
  { section: "주차장", itemCode: "4회", item: "거푸집", spec: "4회", category: "거푸집" },
  { section: "주차장", itemCode: "유로", item: "거푸집", spec: "유로", category: "거푸집" },
  { section: "주차장", itemCode: "알폼", item: "거푸집", spec: "알폼", category: "거푸집" },
  { section: "주차장", itemCode: "갱폼", item: "거푸집", spec: "갱폼", category: "거푸집" },
  { section: "주차장", itemCode: "합벽", item: "거푸집", spec: "합벽", category: "거푸집" },
  { section: "주차장", itemCode: "보밑면", item: "거푸집", spec: "보밑면", category: "거푸집" },
  { section: "주차장", itemCode: "데크", item: "거푸집", spec: "데크플레이트", category: "거푸집" },
  { section: "주차장", itemCode: "방수턱", item: "거푸집", spec: "방수턱", category: "거푸집" },
  { section: "주차장", itemCode: "H10", item: "철근", spec: "H10", category: "철근" },
  { section: "주차장", itemCode: "H13", item: "철근", spec: "H13", category: "철근" },
  { section: "주차장", itemCode: "H16", item: "철근", spec: "H16", category: "철근" },
  { section: "주차장", itemCode: "H19", item: "철근", spec: "H19", category: "철근" },
  { section: "주차장", itemCode: "H22", item: "철근", spec: "H22", category: "철근" },
  { section: "주차장", itemCode: "H25", item: "철근", spec: "H25", category: "철근" },
  { section: "주차장", itemCode: "H29", item: "철근", spec: "H29", category: "철근" },

  { type: "section", section: "부대동" },
  { section: "부대동", itemCode: "240", item: "레미콘", spec: "25-24-15", category: "레미콘" },
  { section: "부대동", itemCode: "270", item: "레미콘", spec: "25-27-15", category: "레미콘" },
  { section: "부대동", itemCode: "300", item: "레미콘", spec: "25-30-15", category: "레미콘" },
  { section: "부대동", itemCode: "180", item: "레미콘", spec: "25-18-08", category: "레미콘" },
  { section: "부대동", itemCode: "3회", item: "거푸집", spec: "3회", category: "거푸집" },
  { section: "부대동", itemCode: "4회", item: "거푸집", spec: "4회", category: "거푸집" },
  { section: "부대동", itemCode: "유로", item: "거푸집", spec: "유로", category: "거푸집" },
  { section: "부대동", itemCode: "알폼", item: "거푸집", spec: "알폼", category: "거푸집" },
  { section: "부대동", itemCode: "갱폼", item: "거푸집", spec: "갱폼", category: "거푸집" },
  { section: "부대동", itemCode: "합벽", item: "거푸집", spec: "합벽", category: "거푸집" },
  { section: "부대동", itemCode: "보밑면", item: "거푸집", spec: "보밑면", category: "거푸집" },
  { section: "부대동", itemCode: "데크", item: "거푸집", spec: "데크플레이트", category: "거푸집" },
  { section: "부대동", itemCode: "방수턱", item: "거푸집", spec: "방수턱", category: "거푸집" },
  { section: "부대동", itemCode: "H10", item: "철근", spec: "H10", category: "철근" },
  { section: "부대동", itemCode: "H13", item: "철근", spec: "H13", category: "철근" },
  { section: "부대동", itemCode: "H16", item: "철근", spec: "H16", category: "철근" },
  { section: "부대동", itemCode: "H19", item: "철근", spec: "H19", category: "철근" },
  { section: "부대동", itemCode: "H22", item: "철근", spec: "H22", category: "철근" },
  { section: "부대동", itemCode: "H25", item: "철근", spec: "H25", category: "철근" },
  { section: "부대동", itemCode: "H29", item: "철근", spec: "H29", category: "철근" },
];

/** -----------------------------
 * 상태
 * ----------------------------- */
const state = {
  rawEntriesByProject: {
    current: [],
    A: [],
    B: [],
    C: [],
  },
  fileSummaryByProject: {
    current: [],
    A: [],
    B: [],
    C: [],
  },
  uniqueNames: [],
  mappingConfig: {},
  lastCompareRows: [],
};

/** -----------------------------
 * 공통 유틸
 * ----------------------------- */
function setStatus(text) {
  statusBox.textContent = text;
}

function setLog(text) {
  logBox.textContent = text;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizeText(value) {
  return String(value ?? "").replace(/\s+/g, "").trim().toUpperCase();
}

function normalizeDisplayText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function toNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  const cleaned = String(value ?? "").replace(/,/g, "").trim();
  if (!cleaned) return 0;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function fmtNumber(value) {
  return Number(value || 0).toLocaleString("ko-KR", {
    maximumFractionDigits: 2,
  });
}

function fmtRatio(value) {
  return Number(value || 0).toLocaleString("ko-KR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function ratioClass(value) {
  if (!value) return "";
  if (value < 0.9) return "bad";
  if (value <= 1.1) return "good";
  return "warn";
}

function updateFileListText() {
  for (const key of PROJECT_KEYS) {
    const files = Array.from(fileInputs[key].files || []);
    if (!files.length) {
      fileListEls[key].textContent = "선택된 파일 없음";
      continue;
    }
    fileListEls[key].textContent = files.map((f) => f.name).join("\n");
  }
}

function openModal() {
  mappingModal.classList.add("is-open");
  mappingModal.setAttribute("aria-hidden", "false");
}

function closeModal() {
  mappingModal.classList.remove("is-open");
  mappingModal.setAttribute("aria-hidden", "true");
}

/** -----------------------------
 * 엑셀 파싱
 * ----------------------------- */
function findAnalysisSheetName(workbook) {
  const names = workbook.SheetNames || [];
  const exact = names.find((n) => normalizeText(n) === normalizeText("분석표B"));
  if (exact) return exact;

  const partial = names.find((n) => normalizeText(n).includes(normalizeText("분석표B")));
  if (partial) return partial;

  return names[0];
}

function getSectionNameFromSheet(sheet) {
  const candidates = ["A5", "B5", "C5", "A4", "B4", "C4"];
  for (const addr of candidates) {
    const cell = sheet[addr];
    const raw = normalizeDisplayText(cell ? cell.v : "");
    if (!raw) continue;
    if (raw.includes("[APT]")) return "APT";
    if (raw.includes("[PIT]")) return "PIT";
    if (raw.includes("[주차장]")) return "주차장";
    if (raw.includes("[부대동]")) return "부대동";
  }
  return "";
}

function parseHeaderNames(rows) {
  const headerRow3 = rows[2] || [];
  const headerRow4 = rows[3] || [];

  const colStart = 1;
  const colEnd = Math.max(headerRow3.length, headerRow4.length) - 1;
  const results = [];

  for (let c = colStart; c <= colEnd; c += 1) {
    const top = normalizeDisplayText(headerRow3[c]);
    const bottom = normalizeDisplayText(headerRow4[c]);

    if (!top && !bottom) continue;

    let name = "";
    if (bottom) {
      name = bottom;
    } else {
      name = top;
    }

    if (!name) continue;

    results.push({
      colIndex: c,
      rawTop: top,
      rawBottom: bottom,
      rawName: name,
      normalizedName: normalizeText(name),
    });
  }

  return results;
}

function isSummaryRow(text) {
  const t = normalizeText(text);
  if (!t) return false;

  const keywords = [
    "소계",
    "할증",
    "계",
    "합계",
    "TOTAL",
    "비율",
    "수량합계",
    "면적합계",
  ];

  return keywords.some((k) => t.includes(normalizeText(k)));
}

function parseSheetEntries(rows, sectionName, projectKey, fileName) {
  const headers = parseHeaderNames(rows);
  const entries = [];

  for (let r = 5; r < rows.length; r += 1) {
    const row = rows[r] || [];
    const rowLabel = normalizeDisplayText(row[0]);

    if (isSummaryRow(rowLabel)) {
      break;
    }

    for (const header of headers) {
      const value = toNumber(row[header.colIndex]);
      if (!value) continue;

      entries.push({
        projectKey,
        projectLabel: PROJECT_LABELS[projectKey],
        section: sectionName || "미확인",
        sourceFile: fileName,
        rowIndex: r + 1,
        rowLabel: rowLabel || "",
        rawName: header.rawName,
        normalizedName: header.normalizedName,
        qty: value,
      });
    }
  }

  return entries;
}

async function parseWorkbookFile(file, projectKey) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = findAnalysisSheetName(workbook);
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    defval: "",
  });

  const sectionName = getSectionNameFromSheet(sheet);
  const entries = parseSheetEntries(rows, sectionName, projectKey, file.name);

  return {
    projectKey,
    fileName: file.name,
    sheetName,
    sectionName,
    rowCount: rows.length,
    entryCount: entries.length,
    entries,
  };
}

/** -----------------------------
 * 명칭 처리 제안
 * ----------------------------- */
function suggestMappingByName(rawName) {
  const t = normalizeText(rawName);

  const result = {
    include: "Y",
    category: "",
    itemCode: "",
    note: "",
  };

  // 레미콘
  if (t.includes("24MPA")) {
    result.category = "레미콘";
    result.itemCode = "240";
    result.note = "24MPA → 240";
    return result;
  }
  if (t.includes("27MPA")) {
    result.category = "레미콘";
    result.itemCode = "270";
    result.note = "27MPA → 270";
    return result;
  }
  if (t.includes("30MPA")) {
    result.category = "레미콘";
    result.itemCode = "300";
    result.note = "30MPA → 300";
    return result;
  }
  if (t.includes("현치/무근") || t.includes("버림")) {
    result.category = "레미콘";
    result.itemCode = "180";
    result.note = "버림/현치무근 → 180";
    return result;
  }
  if (t.includes("현치/구체") || t === "합벽" || t === "합벽채움" || t.includes("ACT-INSIDE")) {
    result.category = "레미콘";
    result.itemCode = "240";
    result.note = "구체/합벽/ACT-INSIDE 계열 → 240";
    return result;
  }

  // 거푸집
  if (t === "3회") {
    result.category = "거푸집";
    result.itemCode = "3회";
    result.note = "3회 → 3회";
    return result;
  }
  if (t === "4회") {
    result.category = "거푸집";
    result.itemCode = "4회";
    result.note = "4회 → 4회";
    return result;
  }
  if (t === "유로") {
    result.category = "거푸집";
    result.itemCode = "유로";
    result.note = "유로 → 유로";
    return result;
  }
  if (t === "원형" || t === "CURVED") {
    result.category = "거푸집";
    result.itemCode = "유로";
    result.note = "원형/CURVED → 유로";
    return result;
  }
  if (t === "알폼-H" || t === "알폼-V" || t === "알폼") {
    result.category = "거푸집";
    result.itemCode = "알폼";
    result.note = "알폼-H/V → 알폼";
    return result;
  }
  if (t === "갱폼" || t === "갱폼-2") {
    result.category = "거푸집";
    result.itemCode = "갱폼";
    result.note = "갱폼/갱폼-2 → 갱폼";
    return result;
  }
  if (t === "합벽") {
    result.category = "거푸집";
    result.itemCode = "합벽";
    result.note = "합벽 → 합벽";
    return result;
  }
  if (t === "보밑면") {
    result.category = "거푸집";
    result.itemCode = "보밑면";
    result.note = "보밑면 → 보밑면";
    return result;
  }
  if (t === "DECK") {
    result.category = "거푸집";
    result.itemCode = "데크";
    result.note = "DECK → 데크";
    return result;
  }
  if (t === "방수턱") {
    result.category = "거푸집";
    result.itemCode = "방수턱";
    result.note = "방수턱 → 방수턱";
    return result;
  }

  // 철근
  if (t === "H10") {
    result.category = "철근";
    result.itemCode = "H10";
    result.note = "H10 → H10";
    return result;
  }
  if (t === "H13") {
    result.category = "철근";
    result.itemCode = "H13";
    result.note = "H13 → H13";
    return result;
  }
  if (t === "H16") {
    result.category = "철근";
    result.itemCode = "H16";
    result.note = "H16 → H16";
    return result;
  }
  if (t === "H19") {
    result.category = "철근";
    result.itemCode = "H19";
    result.note = "H19 → H19";
    return result;
  }
  if (t === "D22" || t === "H22") {
    result.category = "철근";
    result.itemCode = "H22";
    result.note = "D22/H22 → H22";
    return result;
  }
  if (t === "H25") {
    result.category = "철근";
    result.itemCode = "H25";
    result.note = "H25 → H25";
    return result;
  }
  if (t === "H29") {
    result.category = "철근";
    result.itemCode = "H29";
    result.note = "H29 → H29";
    return result;
  }

  result.include = "N";
  result.category = "제외";
  result.itemCode = "";
  result.note = "자동 제안 없음";
  return result;
}

/** -----------------------------
 * 명칭 추출 / 매핑 렌더링
 * ----------------------------- */
function buildUniqueNamesFromEntries() {
  const nameMap = new Map();

  for (const projectKey of PROJECT_KEYS) {
    for (const entry of state.rawEntriesByProject[projectKey]) {
      if (!nameMap.has(entry.normalizedName)) {
        nameMap.set(entry.normalizedName, {
          rawName: entry.rawName,
          normalizedName: entry.normalizedName,
        });
      }
    }
  }

  state.uniqueNames = Array.from(nameMap.values()).sort((a, b) => {
    return a.rawName.localeCompare(b.rawName, "ko");
  });
}

function ensureMappingConfig() {
  for (const item of state.uniqueNames) {
    if (!state.mappingConfig[item.normalizedName]) {
      state.mappingConfig[item.normalizedName] = suggestMappingByName(item.rawName);
    }
  }
}

function optionHtml(list, selectedValue) {
  return list
    .map((opt) => {
      const selected = String(opt.value) === String(selectedValue) ? "selected" : "";
      return `<option value="${escapeHtml(opt.value)}" ${selected}>${escapeHtml(opt.label)}</option>`;
    })
    .join("");
}

function itemCodeOptionHtml(selectedValue) {
  return ITEM_CODE_OPTIONS
    .map((value) => {
      const selected = String(value) === String(selectedValue) ? "selected" : "";
      const label = value || "선택";
      return `<option value="${escapeHtml(value)}" ${selected}>${escapeHtml(label)}</option>`;
    })
    .join("");
}

function renderMappingTable() {
  if (!state.uniqueNames.length) {
    mappingBody.innerHTML = `
      <tr>
        <td colspan="5">추출된 명칭이 없습니다.</td>
      </tr>
    `;
    return;
  }

  const html = state.uniqueNames
    .map((item) => {
      const config = state.mappingConfig[item.normalizedName] || suggestMappingByName(item.rawName);
      const suggestion = suggestMappingByName(item.rawName);

      return `
        <tr data-name-key="${escapeHtml(item.normalizedName)}">
          <td>${escapeHtml(item.rawName)}</td>
          <td>
            <select class="map-include">
              ${optionHtml(INCLUDE_OPTIONS, config.include)}
            </select>
          </td>
          <td>
            <select class="map-category">
              ${optionHtml(CATEGORY_OPTIONS, config.category)}
            </select>
          </td>
          <td>
            <select class="map-itemcode">
              ${itemCodeOptionHtml(config.itemCode)}
            </select>
          </td>
          <td class="mapping-suggest">${escapeHtml(suggestion.note || "-")}</td>
        </tr>
      `;
    })
    .join("");

  mappingBody.innerHTML = html;
}

function saveMappingFromUI() {
  const rows = Array.from(mappingBody.querySelectorAll("tr[data-name-key]"));

  rows.forEach((row) => {
    const key = row.dataset.nameKey;
    state.mappingConfig[key] = {
      include: row.querySelector(".map-include")?.value || "Y",
      category: row.querySelector(".map-category")?.value || "",
      itemCode: row.querySelector(".map-itemcode")?.value || "",
      note: suggestMappingByName(key).note || "",
    };
  });
}

function applyAutoSuggestionsToCurrentMapping() {
  for (const item of state.uniqueNames) {
    state.mappingConfig[item.normalizedName] = suggestMappingByName(item.rawName);
  }
  renderMappingTable();
}

/** -----------------------------
 * 데이터 취합
 * ----------------------------- */
function clearDataState() {
  state.rawEntriesByProject = {
    current: [],
    A: [],
    B: [],
    C: [],
  };
  state.fileSummaryByProject = {
    current: [],
    A: [],
    B: [],
    C: [],
  };
  state.uniqueNames = [];
  state.mappingConfig = {};
  state.lastCompareRows = [];
}

async function extractNamesFromFiles() {
  clearDataState();

  const logs = [];
  let totalFileCount = 0;

  for (const projectKey of PROJECT_KEYS) {
    const files = Array.from(fileInputs[projectKey].files || []);
    totalFileCount += files.length;

    for (const file of files) {
      const parsed = await parseWorkbookFile(file, projectKey);
      state.rawEntriesByProject[projectKey].push(...parsed.entries);
      state.fileSummaryByProject[projectKey].push(parsed);

      logs.push(
        `[${PROJECT_LABELS[projectKey]}] 파일: ${parsed.fileName}`,
        `- 시트명: ${parsed.sheetName}`,
        `- 구분: ${parsed.sectionName || "미확인"}`,
        `- 추출건수: ${parsed.entryCount}`,
        ""
      );
    }
  }

  if (!totalFileCount) {
    throw new Error("업로드된 파일이 없습니다.");
  }

  buildUniqueNamesFromEntries();
  ensureMappingConfig();

  btnOpenMapping.disabled = state.uniqueNames.length === 0;
  btnCalc.disabled = state.uniqueNames.length === 0;

  setLog(logs.join("\n") || "로그가 없습니다.");
  setStatus(`명칭 추출 완료: ${state.uniqueNames.length}개`);
}

/** -----------------------------
 * 비교표 계산
 * ----------------------------- */
function getMappedEntriesByProject() {
  const result = {
    current: [],
    A: [],
    B: [],
    C: [],
  };

  for (const projectKey of PROJECT_KEYS) {
    result[projectKey] = state.rawEntriesByProject[projectKey]
      .map((entry) => {
        const config = state.mappingConfig[entry.normalizedName];
        if (!config) return null;
        if (config.include !== "Y") return null;
        if (!config.category || config.category === "제외") return null;
        if (!config.itemCode) return null;

        return {
          ...entry,
          mappedCategory: config.category,
          mappedItemCode: config.itemCode,
        };
      })
      .filter(Boolean);
  }

  return result;
}

function buildProjectAggregate(mappedEntries) {
  const aggregate = {
    current: {},
    A: {},
    B: {},
    C: {},
  };

  for (const projectKey of PROJECT_KEYS) {
    const bucket = {};

    for (const entry of mappedEntries[projectKey]) {
      const key = `${entry.section}__${entry.mappedCategory}__${entry.mappedItemCode}`;
      bucket[key] = (bucket[key] || 0) + entry.qty;
    }

    aggregate[projectKey] = bucket;
  }

  return aggregate;
}

function calcCompareRows() {
  const mappedEntries = getMappedEntriesByProject();
  const aggregate = buildProjectAggregate(mappedEntries);

  const rows = [];

  for (const row of COMPARE_LAYOUT) {
    if (row.type === "section") {
      rows.push(row);
      continue;
    }

    const key = `${row.section}__${row.category}__${row.itemCode}`;
    const current = aggregate.current[key] || 0;
    const A = aggregate.A[key] || 0;
    const B = aggregate.B[key] || 0;
    const C = aggregate.C[key] || 0;
    const avg = (A + B + C) / 3;
    const ratio = current === 0 ? 0 : avg / current;

    rows.push({
      ...row,
      current,
      A,
      B,
      C,
      avg,
      ratio,
      note: "",
    });
  }

  state.lastCompareRows = rows;
  return rows;
}

function renderCompareTable(rows) {
  if (!rows.length) {
    compareBody.innerHTML = `
      <tr>
        <td colspan="11" class="empty-row">비교표가 아직 생성되지 않았습니다.</td>
      </tr>
    `;
    return;
  }

  const html = rows
    .map((row) => {
      if (row.type === "section") {
        const cls = SECTION_DISPLAY_CLASS[row.section] || "";
        return `
          <tr class="section-row ${cls}">
            <td colspan="11">${escapeHtml(row.section)}</td>
          </tr>
        `;
      }

      return `
        <tr>
          <td>${escapeHtml(row.section)}</td>
          <td>${escapeHtml(row.itemCode)}</td>
          <td>${escapeHtml(row.item)}</td>
          <td>${escapeHtml(row.spec)}</td>
          <td class="num">${fmtNumber(row.current)}</td>
          <td class="num">${fmtNumber(row.A)}</td>
          <td class="num">${fmtNumber(row.B)}</td>
          <td class="num">${fmtNumber(row.C)}</td>
          <td class="num">${fmtNumber(row.avg)}</td>
          <td class="num ratio ${ratioClass(row.ratio)}">${fmtRatio(row.ratio)}</td>
          <td>${escapeHtml(row.note || "")}</td>
        </tr>
      `;
    })
    .join("");

  compareBody.innerHTML = html;
}

function validateBeforeCalc() {
  if (!state.uniqueNames.length) {
    throw new Error("먼저 명칭을 추출해야 합니다.");
  }

  const missing = [];

  for (const item of state.uniqueNames) {
    const config = state.mappingConfig[item.normalizedName];
    if (!config) {
      missing.push(`${item.rawName} : 설정 없음`);
      continue;
    }
    if (config.include === "Y") {
      if (!config.category || config.category === "제외") {
        missing.push(`${item.rawName} : 분류 미설정`);
      }
      if (!config.itemCode) {
        missing.push(`${item.rawName} : 아이템구분 미설정`);
      }
    }
  }

  if (missing.length) {
    throw new Error(
      `명칭 설정이 완료되지 않았습니다.\n\n${missing.slice(0, 30).join("\n")}${missing.length > 30 ? "\n..." : ""}`
    );
  }
}

/** -----------------------------
 * CSV 내보내기
 * ----------------------------- */
function exportCompareCsv() {
  if (!state.lastCompareRows.length) {
    setStatus("내보낼 비교표가 없습니다.");
    return;
  }

  const lines = [
    ["구분", "아이템구분", "품명", "규격", "현재 프로젝트", "A 프로젝트", "B 프로젝트", "C 프로젝트", "평균치(A~C)", "비율", "비고"]
  ];

  state.lastCompareRows.forEach((row) => {
    if (row.type === "section") {
      lines.push([row.section, "", "", "", "", "", "", "", "", "", ""]);
    } else {
      lines.push([
        row.section,
        row.itemCode,
        row.item,
        row.spec,
        row.current,
        row.A,
        row.B,
        row.C,
        row.avg,
        row.ratio,
        row.note || "",
      ]);
    }
  });

  const csv = "\uFEFF" + lines
    .map((line) =>
      line.map((cell) => `"${String(cell ?? "").replaceAll('"', '""')}"`).join(",")
    )
    .join("\r\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "비교표_결과.csv";
  a.click();

  URL.revokeObjectURL(url);
}

/** -----------------------------
 * 초기화
 * ----------------------------- */
function resetAll() {
  for (const key of PROJECT_KEYS) {
    fileInputs[key].value = "";
  }
  updateFileListText();
  clearDataState();

  btnOpenMapping.disabled = true;
  btnCalc.disabled = true;
  btnExportCsv.disabled = true;

  compareBody.innerHTML = `
    <tr>
      <td colspan="11" class="empty-row">비교표가 아직 생성되지 않았습니다.</td>
    </tr>
  `;

  mappingBody.innerHTML = "";
  setStatus("대기 중");
  setLog("로그가 여기에 표시됩니다.");
  closeModal();
}

/** -----------------------------
 * 이벤트
 * ----------------------------- */
for (const key of PROJECT_KEYS) {
  fileInputs[key].addEventListener("change", updateFileListText);
}

btnExtract.addEventListener("click", async () => {
  try {
    setStatus("명칭 추출 중...");
    setLog("파일을 분석하고 있습니다...");
    await extractNamesFromFiles();
    renderMappingTable();
    openModal();
  } catch (error) {
    console.error(error);
    setStatus("오류 발생");
    setLog(error?.message || String(error));
  }
});

btnOpenMapping.addEventListener("click", () => {
  renderMappingTable();
  openModal();
});

btnCloseMapping.addEventListener("click", closeModal);
mappingBackdrop.addEventListener("click", closeModal);

btnApplySuggestions.addEventListener("click", () => {
  applyAutoSuggestionsToCurrentMapping();
});

btnSaveMapping.addEventListener("click", () => {
  saveMappingFromUI();
  closeModal();
  setStatus("명칭 설정 저장 완료");
});

btnCalc.addEventListener("click", () => {
  try {
    if (mappingModal.classList.contains("is-open")) {
      saveMappingFromUI();
    }

    validateBeforeCalc();
    const rows = calcCompareRows();
    renderCompareTable(rows);

    btnExportCsv.disabled = rows.length === 0;
    setStatus("비교표 생성 완료");
  } catch (error) {
    console.error(error);
    setStatus("오류 발생");
    setLog(error?.message || String(error));
  }
});

btnExportCsv.addEventListener("click", exportCompareCsv);
btnReset.addEventListener("click", resetAll);

/** -----------------------------
 * 최초 상태
 * ----------------------------- */
updateFileListText();
setStatus("대기 중");
setLog("로그가 여기에 표시됩니다.");
