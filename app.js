import {
  WorkbookError,
  buildPreview,
  createUpdatedWorkbook,
  listSheetNames,
  makeOutputName,
  openWorkbook,
} from "./xlsx-engine.js";
import { makeBackupName, writeFileHandle } from "./file-output.js";

const config = globalThis.APP_CONFIG;
const state = {
  accessToken: "",
  sourceWorkbook: null,
  targetWorkbook: null,
  targetFileHandle: null,
  targetOriginalFile: null,
  sourceName: "",
  targetName: "",
  preview: null,
  activeFilter: "changes",
};

const elements = Object.fromEntries(
  [
    "globalMessage",
    "driveUrl",
    "connectDrive",
    "sourceFile",
    "sourceFileName",
    "sourceSheet",
    "sourceStatus",
    "sourceFacts",
    "targetFile",
    "targetFilePicker",
    "targetFileName",
    "targetStatus",
    "targetSheetName",
    "runPreview",
    "emptyPreview",
    "previewContent",
    "countSource",
    "countChanges",
    "countUnchanged",
    "countIssues",
    "previewCaption",
    "previewRows",
    "downloadResult",
    "overwriteOriginal",
  ].map((id) => [id, document.getElementById(id)]),
);

elements.targetSheetName.textContent = config.targetSheetName;
elements.driveUrl.value = localStorage.getItem("renewal.driveUrl") || "";

function showMessage(message, type = "info") {
  elements.globalMessage.textContent = message;
  elements.globalMessage.className = `message${type === "error" ? " is-error" : ""}${type === "success" ? " is-success" : ""}`;
  elements.globalMessage.hidden = false;
}

function clearMessage() {
  elements.globalMessage.hidden = true;
  elements.globalMessage.textContent = "";
}

function setStatus(element, text, type = "idle") {
  element.textContent = text;
  element.className = `status-chip${type === "ready" ? " is-ready" : ""}${type === "error" ? " is-error" : ""}`;
}

function updateSteps() {
  const sourceReady = Boolean(state.sourceWorkbook && elements.sourceSheet.value);
  const targetReady = Boolean(state.targetWorkbook);
  document.querySelectorAll(".step").forEach((step) => {
    step.classList.remove("is-active", "is-complete");
    const name = step.dataset.step;
    if (name === "source") step.classList.add(sourceReady ? "is-complete" : "is-active");
    if (name === "target") {
      if (targetReady) step.classList.add("is-complete");
      else if (sourceReady) step.classList.add("is-active");
    }
    if (name === "preview" && sourceReady && targetReady) step.classList.add("is-active");
  });
  elements.runPreview.disabled = !(sourceReady && targetReady);
}

function friendlyError(error) {
  if (error instanceof WorkbookError) return error.message;
  if (error?.name === "AbortError") return "操作已取消。";
  return error?.message || "發生未預期的錯誤。";
}

function formatBytes(size) {
  if (!Number.isFinite(size)) return "—";
  if (size < 1024) return `${size} B`;
  if (size < 1024 ** 2) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / 1024 ** 2).toFixed(1)} MB`;
}

function extractDriveFileId(url) {
  const input = String(url || "").trim();
  const pathMatch = input.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (pathMatch) return pathMatch[1];
  try {
    const parsed = new URL(input);
    const id = parsed.searchParams.get("id");
    if (id) return id;
  } catch {
    // The user-facing error below covers invalid URLs.
  }
  throw new Error("這不是可辨識的 Google Drive Excel 連結。");
}

function waitForGoogleIdentity(timeoutMs = 8000) {
  if (globalThis.google?.accounts?.oauth2) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const started = Date.now();
    const timer = setInterval(() => {
      if (globalThis.google?.accounts?.oauth2) {
        clearInterval(timer);
        resolve();
      } else if (Date.now() - started > timeoutMs) {
        clearInterval(timer);
        reject(new Error("Google 登入元件載入失敗，請檢查網路後重新整理。"));
      }
    }, 100);
  });
}

async function requestDriveToken() {
  if (!config.googleClientId) {
    throw new Error("尚未設定 Google OAuth Client ID。管理者可先用下方測試檔功能驗證流程。");
  }
  await waitForGoogleIdentity();
  return new Promise((resolve, reject) => {
    const tokenClient = globalThis.google.accounts.oauth2.initTokenClient({
      client_id: config.googleClientId,
      scope: "https://www.googleapis.com/auth/drive.readonly",
      callback: (response) => {
        if (response.error) reject(new Error(`Google 授權失敗：${response.error}`));
        else resolve(response.access_token);
      },
      error_callback: () => reject(new Error("Google 登入視窗已關閉或無法開啟。")),
    });
    tokenClient.requestAccessToken({ prompt: state.accessToken ? "" : "consent" });
  });
}

async function driveRequest(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${state.accessToken}`,
      ...(options.headers || {}),
    },
  });
  if (response.status === 401) {
    state.accessToken = "";
    throw new Error("Google 授權已過期，請重新按「登入並讀取」。");
  }
  if (!response.ok) {
    const details = await response.json().catch(() => ({}));
    throw new Error(details?.error?.message || `Google Drive 讀取失敗（${response.status}）。`);
  }
  return response;
}

async function loadDriveWorkbook() {
  clearMessage();
  const url = elements.driveUrl.value;
  const fileId = extractDriveFileId(url);
  elements.connectDrive.disabled = true;
  elements.connectDrive.textContent = "正在讀取…";
  setStatus(elements.sourceStatus, "連線中");

  try {
    state.accessToken = await requestDriveToken();
    const metadataResponse = await driveRequest(
      `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?fields=id,name,size,mimeType,modifiedTime,capabilities(canDownload)`,
    );
    const metadata = await metadataResponse.json();
    if (metadata.capabilities?.canDownload === false) {
      throw new Error("這個 Drive 檔案不允許下載。請確認共用權限。");
    }
    if (!String(metadata.name || "").toLowerCase().endsWith(".xlsx")) {
      throw new Error("來源必須是 Excel .xlsx 檔案。");
    }
    const fileResponse = await driveRequest(
      `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?alt=media`,
    );
    const buffer = await fileResponse.arrayBuffer();
    setSourceWorkbook(openWorkbook(buffer, metadata.name), {
      name: metadata.name,
      size: Number(metadata.size || buffer.byteLength),
    });
    localStorage.setItem("renewal.driveUrl", url);
    showMessage(`已讀取雲端檔案「${metadata.name}」。`, "success");
  } catch (error) {
    setStatus(elements.sourceStatus, "讀取失敗", "error");
    showMessage(friendlyError(error), "error");
  } finally {
    elements.connectDrive.disabled = false;
    elements.connectDrive.textContent = "登入並讀取";
  }
}

function fillSheetOptions(workbook) {
  const previousSelection = localStorage.getItem("renewal.sourceSheet") || "";
  const names = listSheetNames(workbook);
  elements.sourceSheet.replaceChildren();
  for (const name of names) {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    elements.sourceSheet.appendChild(option);
  }
  const preferred = names.includes(previousSelection)
    ? previousSelection
    : [...names].reverse().find((name) => /月$/.test(name)) || names[0] || "";
  elements.sourceSheet.value = preferred;
  elements.sourceSheet.disabled = names.length === 0;
}

function setSourceWorkbook(workbook, fileInfo) {
  state.sourceWorkbook = workbook;
  state.sourceName = fileInfo.name;
  state.preview = null;
  fillSheetOptions(workbook);
  elements.sourceFileName.textContent = fileInfo.name;
  const facts = elements.sourceFacts.querySelectorAll("strong");
  facts[0].textContent = `${fileInfo.name} · ${formatBytes(fileInfo.size)}`;
  facts[1].textContent = `${listSheetNames(workbook).length} 個可見分頁`;
  setStatus(elements.sourceStatus, "已讀取", "ready");
  resetPreview();
  updateSteps();
}

async function loadLocalSourceFile(file) {
  if (!file) return;
  clearMessage();
  try {
    const workbook = openWorkbook(await file.arrayBuffer(), file.name);
    setSourceWorkbook(workbook, file);
    showMessage(`已讀取測試來源「${file.name}」。`, "success");
  } catch (error) {
    setStatus(elements.sourceStatus, "讀取失敗", "error");
    showMessage(friendlyError(error), "error");
  }
}

async function loadTargetFile(file) {
  if (!file) return;
  clearMessage();
  try {
    const workbook = openWorkbook(await file.arrayBuffer(), file.name);
    if (!listSheetNames(workbook).includes(config.targetSheetName)) {
      throw new Error(`本機 Excel 找不到「${config.targetSheetName}」工作表。`);
    }
    state.targetWorkbook = workbook;
    state.targetOriginalFile = file;
    state.targetName = file.name;
    state.preview = null;
    elements.targetFileName.textContent = `${file.name} · ${formatBytes(file.size)}`;
    setStatus(elements.targetStatus, "已選擇", "ready");
    resetPreview();
    updateSteps();
  } catch (error) {
    state.targetWorkbook = null;
    state.targetFileHandle = null;
    state.targetOriginalFile = null;
    setStatus(elements.targetStatus, "讀取失敗", "error");
    showMessage(friendlyError(error), "error");
    updateSteps();
  }
}

function resetPreview() {
  state.preview = null;
  elements.previewContent.hidden = true;
  elements.emptyPreview.hidden = false;
  elements.previewRows.replaceChildren();
  elements.downloadResult.disabled = true;
  elements.overwriteOriginal.disabled = true;
}

function issueCount(preview) {
  return preview.issues.length;
}

function renderSummary(preview) {
  elements.countSource.textContent = preview.stats.sourceIds.toLocaleString("zh-TW");
  elements.countChanges.textContent = preview.stats.changedRows.toLocaleString("zh-TW");
  elements.countUnchanged.textContent = preview.stats.unchangedRows.toLocaleString("zh-TW");
  elements.countIssues.textContent = issueCount(preview).toLocaleString("zh-TW");
}

function makeCell(text, className = "") {
  const cell = document.createElement("td");
  cell.textContent = text || "—";
  if (className) cell.className = className;
  return cell;
}

function renderPreviewRows() {
  const preview = state.preview;
  if (!preview) return;
  const rows = state.activeFilter === "changes" ? preview.changes : preview.issues;
  const visible = rows.slice(0, config.previewLimit);
  elements.previewRows.replaceChildren();

  if (!visible.length) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 5;
    cell.textContent = state.activeFilter === "changes" ? "目前沒有需要修改的資料。" : "目前沒有異常資料。";
    row.appendChild(cell);
    elements.previewRows.appendChild(row);
  } else {
    for (const item of visible) {
      const row = document.createElement("tr");
      row.appendChild(makeCell(item.id));
      row.appendChild(makeCell(String(item.rowNumber || "—")));
      if (state.activeFilter === "changes") {
        row.appendChild(makeCell(item.field));
        row.appendChild(makeCell(item.oldReply, "cell-reply"));
        row.appendChild(makeCell(item.newReply, "cell-reply"));
      } else {
        const issueCell = document.createElement("td");
        const label = document.createElement("span");
        label.className = "issue-label";
        label.textContent = item.message;
        issueCell.appendChild(label);
        row.appendChild(issueCell);
        row.appendChild(makeCell(item.oldReply, "cell-reply"));
        row.appendChild(makeCell(item.newReply, "cell-reply"));
      }
      elements.previewRows.appendChild(row);
    }
  }
  elements.previewCaption.textContent = `共 ${rows.length.toLocaleString("zh-TW")} 筆，顯示前 ${Math.min(rows.length, config.previewLimit)} 筆`;
}

function runPreview() {
  clearMessage();
  try {
    const sourceSheetName = elements.sourceSheet.value;
    localStorage.setItem("renewal.sourceSheet", sourceSheetName);
    state.preview = buildPreview({
      sourceWorkbook: state.sourceWorkbook,
      sourceSheetName,
      targetWorkbook: state.targetWorkbook,
      targetSheetName: config.targetSheetName,
      headers: config.headers,
    });
    renderSummary(state.preview);
    state.activeFilter = "changes";
    document.querySelectorAll(".filter-tab").forEach((tab) =>
      tab.classList.toggle("is-selected", tab.dataset.filter === "changes"),
    );
    renderPreviewRows();
    elements.emptyPreview.hidden = true;
    elements.previewContent.hidden = false;
    elements.downloadResult.disabled = state.preview.changes.length === 0;
    elements.overwriteOriginal.disabled =
      state.preview.changes.length === 0 || !state.targetFileHandle;
    showMessage(
      state.preview.changes.length
        ? `檢查完成：${state.preview.replyPeriod} 回覆，預計修改 ${state.preview.stats.changedRows} 列、${state.preview.changes.length} 個欄位。`
        : "檢查完成：目前沒有需要修改的資料。",
      state.preview.issues.length ? "info" : "success",
    );
  } catch (error) {
    resetPreview();
    showMessage(friendlyError(error), "error");
  }
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function downloadResult() {
  clearMessage();
  try {
    const blob = createUpdatedWorkbook(state.targetWorkbook, state.preview);
    const outputName = makeOutputName(state.targetName);
    downloadBlob(blob, outputName);
    elements.downloadResult.focus();
    showMessage(`已產生「${outputName}」，原始 Excel 沒有被修改。`, "success");
  } catch (error) {
    showMessage(friendlyError(error), "error");
  }
}

async function chooseTargetFile() {
  if (typeof globalThis.showOpenFilePicker !== "function") {
    elements.targetFile.click();
    return;
  }
  try {
    const [handle] = await globalThis.showOpenFilePicker({
      multiple: false,
      types: [
        {
          description: "Excel 活頁簿",
          accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
          },
        },
      ],
    });
    state.targetFileHandle = handle;
    await loadTargetFile(await handle.getFile());
  } catch (error) {
    if (error?.name !== "AbortError") showMessage(friendlyError(error), "error");
  }
}

async function overwriteOriginal() {
  if (!state.targetFileHandle || !state.targetOriginalFile) {
    showMessage("請重新選擇原檔，才能取得覆寫權限。", "error");
    return;
  }
  const changeCount = state.preview?.changes?.length || 0;
  if (!changeCount) {
    showMessage("沒有需要覆寫的修改。", "error");
    return;
  }
  if (!globalThis.confirm(`將修改 ${state.preview.stats.changedRows} 列、${changeCount} 個欄位並覆寫「${state.targetName}」。\n覆寫前會先下載原檔備份，是否繼續？`)) return;

  clearMessage();
  elements.overwriteOriginal.disabled = true;
  elements.overwriteOriginal.textContent = "正在備份並覆寫…";
  try {
    const blob = createUpdatedWorkbook(state.targetWorkbook, state.preview);
    const backupName = makeBackupName(state.targetName);
    downloadBlob(state.targetOriginalFile, backupName);
    await writeFileHandle(state.targetFileHandle, blob);

    state.targetOriginalFile = new File([blob], state.targetName, { type: blob.type });
    resetPreview();
    showMessage(`已下載「${backupName}」，並完成覆寫「${state.targetName}」。`, "success");
  } catch (error) {
    showMessage(`原檔未完成覆寫：${friendlyError(error)}`, "error");
  } finally {
    elements.overwriteOriginal.textContent = "備份後覆寫原檔";
  }
}

elements.connectDrive.addEventListener("click", loadDriveWorkbook);
elements.sourceFile.addEventListener("change", (event) => loadLocalSourceFile(event.target.files?.[0]));
elements.targetFilePicker.addEventListener("click", (event) => {
  event.preventDefault();
  chooseTargetFile();
});
elements.targetFile.addEventListener("change", (event) => {
  state.targetFileHandle = null;
  loadTargetFile(event.target.files?.[0]);
});
elements.sourceSheet.addEventListener("change", () => {
  localStorage.setItem("renewal.sourceSheet", elements.sourceSheet.value);
  resetPreview();
  updateSteps();
});
elements.runPreview.addEventListener("click", runPreview);
elements.downloadResult.addEventListener("click", downloadResult);
elements.overwriteOriginal.addEventListener("click", overwriteOriginal);

document.querySelectorAll(".filter-tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    state.activeFilter = tab.dataset.filter;
    document.querySelectorAll(".filter-tab").forEach((item) => item.classList.toggle("is-selected", item === tab));
    renderPreviewRows();
  });
});

document.querySelectorAll("label.file-picker[for]").forEach((label) => {
  label.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    event.preventDefault();
    if (label === elements.targetFilePicker) chooseTargetFile();
    else document.getElementById(label.htmlFor)?.click();
  });
});

if (typeof globalThis.showOpenFilePicker !== "function") {
  elements.overwriteOriginal.hidden = true;
}

window.addEventListener("error", (event) => {
  if (event.message?.includes("ResizeObserver")) return;
  showMessage("頁面執行發生錯誤，請重新整理後再試。", "error");
});

updateSteps();
