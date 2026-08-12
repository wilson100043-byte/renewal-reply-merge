const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const decoder = new TextDecoder("utf-8");
const encoder = new TextEncoder();

export class WorkbookError extends Error {
  constructor(message, code = "WORKBOOK_ERROR") {
    super(message);
    this.name = "WorkbookError";
    this.code = code;
  }
}

function assertZipLibrary() {
  if (!globalThis.fflate?.unzipSync || !globalThis.fflate?.zipSync) {
    throw new WorkbookError("Excel 壓縮元件沒有載入，請重新整理頁面。", "ZIP_LIBRARY_MISSING");
  }
}

function parseXml(xmlText, label) {
  const document = new DOMParser().parseFromString(xmlText, "application/xml");
  const error = document.getElementsByTagName("parsererror")[0];
  if (error) {
    throw new WorkbookError(`${label} 的 XML 無法解析。`, "INVALID_XML");
  }
  return document;
}

function requireFile(files, path) {
  const bytes = files[path];
  if (!bytes) {
    throw new WorkbookError(`Excel 缺少必要檔案：${path}`, "INVALID_XLSX");
  }
  return bytes;
}

function textFile(files, path) {
  return decoder.decode(requireFile(files, path));
}

function normalizePath(baseDir, target) {
  const parts = target.startsWith("/")
    ? target.slice(1).split("/")
    : `${baseDir}/${target}`.split("/");
  const normalized = [];
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") normalized.pop();
    else normalized.push(part);
  }
  return normalized.join("/");
}

function childElements(node, tagName = null) {
  return Array.from(node.childNodes).filter(
    (child) => child.nodeType === Node.ELEMENT_NODE && (!tagName || child.localName === tagName),
  );
}

function firstDescendant(node, localName) {
  return Array.from(node.getElementsByTagNameNS("*", localName))[0] || null;
}

function descendantText(node, localName) {
  return Array.from(node.getElementsByTagNameNS("*", localName))
    .map((item) => item.textContent || "")
    .join("");
}

function cellColumn(reference) {
  return String(reference || "").match(/[A-Za-z]+/)?.[0]?.toUpperCase() || "";
}

function columnNumber(column) {
  let result = 0;
  for (const character of column.toUpperCase()) {
    result = result * 26 + character.charCodeAt(0) - 64;
  }
  return result;
}

export function normalizeIdentifier(value) {
  const normalized = String(value ?? "").trim();
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)[eE][+-]?\d+$/.test(normalized)) {
    const expanded = Number(normalized);
    if (Number.isSafeInteger(expanded)) return String(expanded);
  }
  return normalized.replace(/\.0$/, "");
}

function normalizeReply(value) {
  return String(value ?? "").replace(/\r\n/g, "\n").trim();
}

function resolveReplyPeriod(fileName, sheetName) {
  const normalizedSheet = String(sheetName || "").trim();
  const compactDate = normalizedSheet.match(/(?:^|\D)((?:19|20)\d{2})(0[1-9]|1[0-2])(?:\D|$)/);
  if (compactDate) return `${compactDate[1]}${compactDate[2]}`;

  const rocDate = normalizedSheet.match(/(?:^|\D)(\d{3})\s*年\s*(1[0-2]|0?[1-9])\s*月/);
  if (rocDate) return `${Number(rocDate[1]) + 1911}${String(Number(rocDate[2])).padStart(2, "0")}`;

  const monthMatch = normalizedSheet.match(/(?:^|\D)(1[0-2]|0?[1-9])\s*月(?:\D|$)/);
  if (!monthMatch) return "";

  const sheetYear = normalizedSheet.match(/(?:^|\D)((?:19|20)\d{2})\s*年?/);
  const fileYears = [...String(fileName || "").matchAll(/(?:19|20)\d{2}/g)].map((match) => Number(match[0]));
  const year = Number(sheetYear?.[1]) || Math.max(0, ...fileYears) || new Date().getFullYear();
  const month = String(Number(monthMatch[1])).padStart(2, "0");
  return `${year}${month}`;
}

function joinReplies(current, previous, period = "") {
  const newReply = normalizeReply(current);
  const oldReply = normalizeReply(previous);
  if (!newReply) return oldReply;
  const datedReply = period && !/^\d{6}\s*[:：]/.test(newReply) ? `${period}:${newReply}` : newReply;
  if (!oldReply) return datedReply;
  const separator = /[;；]\s*$/.test(datedReply) ? "" : ";";
  return `${datedReply}${separator}${oldReply}`;
}

function loadSharedStrings(files) {
  const bytes = files["xl/sharedStrings.xml"];
  if (!bytes) return [];
  const document = parseXml(decoder.decode(bytes), "sharedStrings.xml");
  return Array.from(document.getElementsByTagNameNS("*", "si")).map((item) =>
    Array.from(item.getElementsByTagNameNS("*", "t"))
      .map((text) => text.textContent || "")
      .join(""),
  );
}

function getCellValue(cell, sharedStrings) {
  if (!cell) return "";
  const type = cell.getAttribute("t") || "";
  if (type === "inlineStr") return descendantText(cell, "t");
  const value = firstDescendant(cell, "v")?.textContent || "";
  if (type === "s") return sharedStrings[Number(value)] ?? "";
  if (type === "str") return value;
  if (type === "b") return value === "1" ? "TRUE" : "FALSE";
  return value;
}

function rowCells(row, sharedStrings) {
  const result = new Map();
  for (const cell of childElements(row, "c")) {
    result.set(cellColumn(cell.getAttribute("r")), {
      node: cell,
      value: getCellValue(cell, sharedStrings),
      hasFormula: Boolean(firstDescendant(cell, "f")),
    });
  }
  return result;
}

function findHeader(sheetDocument, sharedStrings, expectedHeaders) {
  const rows = Array.from(sheetDocument.getElementsByTagNameNS("*", "row")).slice(0, 20);
  for (const row of rows) {
    const cells = rowCells(row, sharedStrings);
    const headerToColumn = new Map();
    for (const [column, cell] of cells) {
      headerToColumn.set(normalizeReply(cell.value), column);
    }
    if (expectedHeaders.every((header) => headerToColumn.has(header))) {
      return {
        rowNumber: Number(row.getAttribute("r") || 1),
        columns: Object.fromEntries(expectedHeaders.map((header) => [header, headerToColumn.get(header)])),
      };
    }
  }
  throw new WorkbookError(`找不到必要欄位：${expectedHeaders.join("、")}`, "HEADERS_NOT_FOUND");
}

function workbookSheetMap(files) {
  const workbookDocument = parseXml(textFile(files, "xl/workbook.xml"), "workbook.xml");
  const relationsDocument = parseXml(
    textFile(files, "xl/_rels/workbook.xml.rels"),
    "workbook.xml.rels",
  );
  const relations = new Map(
    Array.from(relationsDocument.getElementsByTagNameNS("*", "Relationship")).map((relation) => [
      relation.getAttribute("Id"),
      normalizePath("xl", relation.getAttribute("Target") || ""),
    ]),
  );

  const sheets = [];
  for (const sheet of Array.from(workbookDocument.getElementsByTagNameNS("*", "sheet"))) {
    const relationId = sheet.getAttributeNS(REL_NS, "id") || sheet.getAttribute("r:id");
    const path = relations.get(relationId);
    if (!path || !files[path]) continue;
    sheets.push({
      name: sheet.getAttribute("name") || "",
      path,
      sheetId: sheet.getAttribute("sheetId") || "",
      state: sheet.getAttribute("state") || "visible",
    });
  }
  return sheets;
}

export function openWorkbook(arrayBuffer, fileName = "workbook.xlsx") {
  assertZipLibrary();
  let files;
  try {
    files = globalThis.fflate.unzipSync(new Uint8Array(arrayBuffer));
  } catch {
    throw new WorkbookError("檔案不是有效的 .xlsx，或檔案已損壞。", "INVALID_ZIP");
  }
  const sheets = workbookSheetMap(files);
  if (!sheets.length) {
    throw new WorkbookError("Excel 裡沒有可讀取的工作表。", "NO_SHEETS");
  }
  return {
    fileName,
    files,
    sheets,
    sharedStrings: loadSharedStrings(files),
  };
}

function openSheet(workbook, sheetName) {
  const sheet = workbook.sheets.find((item) => item.name === sheetName);
  if (!sheet) {
    throw new WorkbookError(`找不到工作表「${sheetName}」。`, "SHEET_NOT_FOUND");
  }
  const xmlText = textFile(workbook.files, sheet.path);
  return {
    ...sheet,
    xmlText,
    document: parseXml(xmlText, sheetName),
  };
}

export function listSheetNames(workbook) {
  return workbook.sheets.filter((sheet) => sheet.state !== "hidden").map((sheet) => sheet.name);
}

function readSourceGroups(sourceSheet, workbook, headers, period) {
  const header = findHeader(sourceSheet.document, workbook.sharedStrings, [
    headers.id,
    headers.renewalStatus,
    headers.currentReply,
    headers.previousReply,
  ]);
  const groups = new Map();
  let blankIds = 0;
  let dataRows = 0;

  for (const row of Array.from(sourceSheet.document.getElementsByTagNameNS("*", "row"))) {
    const rowNumber = Number(row.getAttribute("r") || 0);
    if (rowNumber <= header.rowNumber) continue;
    const cells = rowCells(row, workbook.sharedStrings);
    const id = normalizeIdentifier(cells.get(header.columns[headers.id])?.value);
    const renewalStatus = normalizeReply(cells.get(header.columns[headers.renewalStatus])?.value);
    const current = normalizeReply(cells.get(header.columns[headers.currentReply])?.value);
    const previous = normalizeReply(cells.get(header.columns[headers.previousReply])?.value);
    if (!id && !renewalStatus && !current && !previous) continue;
    dataRows += 1;
    if (!id) {
      blankIds += 1;
      continue;
    }
    const merged = joinReplies(current, previous, period);
    const item = { id, renewalStatus, current, previous, merged, rowNumber };
    if (!groups.has(id)) groups.set(id, []);
    groups.get(id).push(item);
  }
  return { groups, blankIds, dataRows, header };
}

function readTargetIndex(targetSheet, workbook, headers) {
  const header = findHeader(targetSheet.document, workbook.sharedStrings, [
    headers.id,
    headers.renewalStatus,
    headers.previousReply,
  ]);
  const index = new Map();
  let blankIds = 0;

  for (const row of Array.from(targetSheet.document.getElementsByTagNameNS("*", "row"))) {
    const rowNumber = Number(row.getAttribute("r") || 0);
    if (rowNumber <= header.rowNumber) continue;
    const cells = rowCells(row, workbook.sharedStrings);
    const id = normalizeIdentifier(cells.get(header.columns[headers.id])?.value);
    if (!id) {
      if (cells.size) blankIds += 1;
      continue;
    }
    const replyCell = cells.get(header.columns[headers.previousReply]);
    const renewalStatusCell = cells.get(header.columns[headers.renewalStatus]);
    const item = {
      id,
      rowNumber,
      oldReply: normalizeReply(replyCell?.value),
      replyHasFormula: Boolean(replyCell?.hasFormula),
      oldRenewalStatus: normalizeReply(renewalStatusCell?.value),
      renewalStatusHasFormula: Boolean(renewalStatusCell?.hasFormula),
    };
    if (!index.has(id)) index.set(id, []);
    index.get(id).push(item);
  }
  return { index, blankIds, header };
}

function conflictRecord(id, records, replyConflict, renewalStatusConflict) {
  const fields = [
    replyConflict ? "回覆" : "",
    renewalStatusConflict ? "續訂/停訂" : "",
  ].filter(Boolean);
  return {
    type: "conflict",
    id,
    rowNumber: records.map((item) => item.rowNumber).join("、"),
    message: `雲端同一編號出現不同${fields.join("與")}結果`,
    oldReply: "",
    newReply: records
      .map((item) => `${item.renewalStatus || "（狀態空白）"}｜${item.merged || "（回覆空白）"}`)
      .join("；"),
  };
}

export function buildPreview({
  sourceWorkbook,
  sourceSheetName,
  targetWorkbook,
  targetSheetName,
  headers,
}) {
  const sourceSheet = openSheet(sourceWorkbook, sourceSheetName);
  const targetSheet = openSheet(targetWorkbook, targetSheetName);
  const replyPeriod = resolveReplyPeriod(sourceWorkbook.fileName, sourceSheetName);
  if (!replyPeriod) {
    throw new WorkbookError(
      `無法從來源檔名與分頁「${sourceSheetName}」判斷月份，請將分頁命名為「6月」或「202606」。`,
      "PERIOD_NOT_FOUND",
    );
  }
  const source = readSourceGroups(sourceSheet, sourceWorkbook, headers, replyPeriod);
  const target = readTargetIndex(targetSheet, targetWorkbook, headers);

  const changes = [];
  const issues = [];
  let unchangedRows = 0;
  let skippedEmpty = 0;

  for (const [id, records] of source.groups) {
    const distinctResults = [...new Set(records.map((item) => item.merged))];
    const distinctRenewalStatuses = [
      ...new Set(records.map((item) => item.renewalStatus).filter(Boolean)),
    ];
    const replyConflict = distinctResults.length > 1;
    const renewalStatusConflict = distinctRenewalStatuses.length > 1;
    if (replyConflict || renewalStatusConflict) {
      issues.push(conflictRecord(id, records, replyConflict, renewalStatusConflict));
      continue;
    }
    const merged = distinctResults[0] || "";
    const renewalStatus = distinctRenewalStatuses[0] || "";
    if (!merged) {
      skippedEmpty += 1;
      if (!renewalStatus) continue;
    }
    const targetRows = target.index.get(id);
    if (!targetRows?.length) {
      issues.push({
        type: "unmatched",
        id,
        rowNumber: "—",
        message: "本機續訂清單找不到此編號",
        oldReply: "",
        newReply: merged,
      });
      continue;
    }
    for (const targetRow of targetRows) {
      let rowChanged = false;
      let rowHasIssue = false;
      const compareField = (field, oldValue, newValue, column, hasFormula) => {
        if (!newValue || oldValue === newValue) return;
        if (hasFormula) {
          rowHasIssue = true;
          issues.push({
            type: "formula",
            id,
            rowNumber: targetRow.rowNumber,
            message: `${field}儲存格含公式，未自動覆蓋`,
            oldReply: oldValue,
            newReply: newValue,
          });
          return;
        }
        rowChanged = true;
        changes.push({
          id,
          field,
          rowNumber: targetRow.rowNumber,
          oldReply: oldValue,
          newReply: newValue,
          column,
        });
      };
      compareField(
        headers.previousReply,
        targetRow.oldReply,
        merged,
        target.header.columns[headers.previousReply],
        targetRow.replyHasFormula,
      );
      compareField(
        headers.renewalStatus,
        targetRow.oldRenewalStatus,
        renewalStatus,
        target.header.columns[headers.renewalStatus],
        targetRow.renewalStatusHasFormula,
      );
      if (!rowChanged && !rowHasIssue) unchangedRows += 1;
    }
  }

  if (source.blankIds) {
    issues.push({
      type: "blank-source-id",
      id: "（空白）",
      rowNumber: "—",
      message: `雲端有 ${source.blankIds} 列缺少編號細目`,
      oldReply: "",
      newReply: "",
    });
  }

  if (target.blankIds) {
    issues.push({
      type: "blank-target-id",
      id: "（空白）",
      rowNumber: "—",
      message: `本機續訂清單有 ${target.blankIds} 列缺少編號細目`,
      oldReply: "",
      newReply: "",
    });
  }

  return {
    changes,
    issues,
    stats: {
      sourceRows: source.dataRows,
      sourceIds: source.groups.size,
      changedRows: new Set(changes.map((change) => change.rowNumber)).size,
      unchangedRows,
      skippedEmpty,
      issueCount: issues.length,
      targetBlankIds: target.blankIds,
    },
    replyPeriod,
    targetSheet,
  };
}

function escapeXmlText(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export function patchWorksheetCell(xmlText, rowNumber, column, value) {
  const cellReference = `${column}${rowNumber}`;
  const tagPrefix = "(?:[A-Za-z_][\\w.-]*:)?";
  const rowPattern = new RegExp(
    `<${tagPrefix}row\\b(?=[^>]*\\br="${rowNumber}")[^>]*>[\\s\\S]*?<\\/${tagPrefix}row>`,
  );
  const rowMatch = xmlText.match(rowPattern);
  if (!rowMatch || rowMatch.index === undefined) {
    throw new WorkbookError(`找不到目標列 ${rowNumber}。`, "ROW_NOT_FOUND");
  }
  const rowXml = rowMatch[0];
  const prefix = rowXml.match(/^<([A-Za-z_][\w.-]*:)?row\b/)?.[1] || "";
  const cellPattern = new RegExp(
    `<${tagPrefix}c\\b(?=[^>]*\\br="${escapeRegExp(cellReference)}")[^>]*(?:\\/>|>[\\s\\S]*?<\\/${tagPrefix}c>)`,
  );
  const existingCell = rowXml.match(cellPattern)?.[0] || "";
  const content = `<${prefix}is><${prefix}t xml:space="preserve">${escapeXmlText(value)}</${prefix}t></${prefix}is>`;
  let updatedRow;

  if (existingCell) {
    const openingTag = existingCell.match(/^<[^>]+>/)?.[0] || "";
    const attributes = openingTag
      .replace(/^<[^\s>]+/, "")
      .replace(/\s*\/?>$/, "")
      .replace(/\s+t=(["'])[^"']*\1/g, "")
      .trim();
    const replacement = `<${prefix}c${attributes ? ` ${attributes}` : ""} t="inlineStr">${content}</${prefix}c>`;
    updatedRow = rowXml.replace(cellPattern, replacement);
  } else {
    const newCell = `<${prefix}c r="${cellReference}" t="inlineStr">${content}</${prefix}c>`;
    const allCellsPattern = new RegExp(
      `<${tagPrefix}c\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/${tagPrefix}c>)`,
      "g",
    );
    const nextCell = [...rowXml.matchAll(allCellsPattern)].find((match) => {
      const reference = match[0].match(/\br="([^"]+)"/)?.[1] || "";
      return columnNumber(cellColumn(reference)) > columnNumber(column);
    });
    const insertionIndex = nextCell?.index ?? rowXml.lastIndexOf(`</${prefix}row>`);
    updatedRow = `${rowXml.slice(0, insertionIndex)}${newCell}${rowXml.slice(insertionIndex)}`;
  }

  return `${xmlText.slice(0, rowMatch.index)}${updatedRow}${xmlText.slice(rowMatch.index + rowXml.length)}`;
}

export function createUpdatedWorkbook(targetWorkbook, preview) {
  if (!preview?.changes) {
    throw new WorkbookError("請先完成更新預覽。", "PREVIEW_REQUIRED");
  }
  if (!preview.changes.length) {
    throw new WorkbookError("沒有需要修改的資料。", "NO_CHANGES");
  }

  let updatedXml = preview.targetSheet.xmlText;
  for (const change of preview.changes) {
    updatedXml = patchWorksheetCell(updatedXml, change.rowNumber, change.column, change.newReply);
  }
  targetWorkbook.files[preview.targetSheet.path] = encoder.encode(updatedXml);
  const bytes = globalThis.fflate.zipSync(targetWorkbook.files, { level: 6 });
  return new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export function makeOutputName(originalName, date = new Date()) {
  const stem = originalName.replace(/\.xlsx$/i, "");
  const stamp = [date.getFullYear(), date.getMonth() + 1, date.getDate()]
    .map((value) => String(value).padStart(2, "0"))
    .join("");
  return `${stem}_${stamp}_已更新.xlsx`;
}

export const mergeReplies = joinReplies;
export const normalizeId = normalizeIdentifier;
