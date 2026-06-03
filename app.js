const state = {
  file: null,
  sheetName: "",
  rows: [],
  groups: new Map(),
  log: []
};

const els = {
  dropZone: document.getElementById("dropZone"),
  fileInput: document.getElementById("fileInput"),
  fileHint: document.getElementById("fileHint"),
  columnInput: document.getElementById("columnInput"),
  emptyNameInput: document.getElementById("emptyNameInput"),
  headerCheckbox: document.getElementById("headerCheckbox"),
  trimCheckbox: document.getElementById("trimCheckbox"),
  splitButton: document.getElementById("splitButton"),
  resetButton: document.getElementById("resetButton"),
  copyLogButton: document.getElementById("copyLogButton"),
  rowCount: document.getElementById("rowCount"),
  groupCount: document.getElementById("groupCount"),
  sheetName: document.getElementById("sheetName"),
  groupList: document.getElementById("groupList"),
  statusTitle: document.getElementById("statusTitle"),
  statusText: document.getElementById("statusText"),
  logOutput: document.getElementById("logOutput"),
  toast: document.getElementById("toast")
};

function init() {
  bindEvents();
  if ("serviceWorker" in navigator && location.protocol.startsWith("http")) {
    navigator.serviceWorker.register("./sw.js").catch(() => {});
  }
}

function bindEvents() {
  els.fileInput.addEventListener("change", () => {
    const [file] = els.fileInput.files;
    if (file) readFile(file);
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    els.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropZone.classList.add("dragging");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    els.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropZone.classList.remove("dragging");
    });
  });

  els.dropZone.addEventListener("drop", (event) => {
    const [file] = event.dataTransfer.files;
    if (file) readFile(file);
  });

  [els.columnInput, els.emptyNameInput, els.headerCheckbox, els.trimCheckbox].forEach((input) => {
    input.addEventListener("input", refreshGroups);
    input.addEventListener("change", refreshGroups);
  });

  els.splitButton.addEventListener("click", splitWorkbook);
  els.resetButton.addEventListener("click", reset);
  els.copyLogButton.addEventListener("click", copyLog);
}

async function readFile(file) {
  if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
    toast("请选择 Excel 或 CSV 文件");
    return;
  }

  setStatus("正在读取文件", file.name);
  log(`读取文件：${file.name}（${formatBytes(file.size)}）`);

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: "" });

    state.file = file;
    state.sheetName = sheetName;
    state.rows = rows.filter((row) => row.some((cell) => String(cell).trim() !== ""));

    if (!state.rows.length) {
      throw new Error("第一个工作表没有可拆分的数据。");
    }

    els.fileHint.textContent = file.name;
    log(`读取工作表：${sheetName}，有效行 ${state.rows.length} 行`);
    refreshGroups();
  } catch (error) {
    resetDataOnly();
    setStatus("读取失败", error.message);
    log(`读取失败：${error.message}`);
    toast("读取失败，请检查文件");
  }
}

function refreshGroups() {
  if (!state.rows.length) return;

  const columnIndex = columnToIndex(els.columnInput.value);
  if (columnIndex < 0) {
    els.splitButton.disabled = true;
    setStatus("拆分列无效", "请输入类似 A、B、AA 的列名。");
    return;
  }

  const hasHeader = els.headerCheckbox.checked;
  const emptyName = normalizeName(els.emptyNameInput.value || "未填写");
  const dataRows = hasHeader ? state.rows.slice(1) : state.rows.slice();
  const groups = new Map();

  dataRows.forEach((row) => {
    const key = normalizeName(row[columnIndex] ?? "", emptyName);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  });

  state.groups = groups;
  els.rowCount.textContent = dataRows.length;
  els.groupCount.textContent = groups.size;
  els.sheetName.textContent = state.sheetName || "-";
  els.splitButton.disabled = !dataRows.length || !groups.size;
  setStatus("文件已就绪", `${groups.size} 个分组，点击按钮即可下载 ZIP。`);
  renderGroupList();
}

function renderGroupList() {
  if (!state.groups.size) {
    els.groupList.textContent = "没有可拆分的数据。";
    return;
  }

  const filenames = buildOutputNames();
  els.groupList.innerHTML = Array.from(state.groups.entries())
    .sort((a, b) => b[1].length - a[1].length)
    .map(([key, rows]) => {
      const name = filenames.get(key);
      return `<div class="group-item"><span>${escapeHTML(name)}</span><b>${rows.length} 行</b></div>`;
    })
    .join("");
}

async function splitWorkbook() {
  if (!state.groups.size) return;

  setStatus("正在生成文件", "文件较大时请稍等几秒。");
  els.splitButton.disabled = true;

  try {
    const zip = new JSZip();
    const header = els.headerCheckbox.checked ? state.rows[0] : null;
    const filenames = buildOutputNames();

    for (const [key, rows] of state.groups.entries()) {
      const outputRows = header ? [header, ...rows] : rows;
      const worksheet = XLSX.utils.aoa_to_sheet(outputRows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName(key));
      const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      zip.file(filenames.get(key), bytes);
    }

    const zipBlob = await zip.generateAsync({ type: "blob" });
    downloadBlob(zipBlob, `${sanitizeFilePart(stripExt(state.file?.name || "excel"))}_拆分结果.zip`);
    log(`生成完成：${state.groups.size} 个 Excel 文件`);
    setStatus("生成完成", "ZIP 已开始下载。");
    toast("拆分完成");
  } catch (error) {
    setStatus("生成失败", error.message);
    log(`生成失败：${error.message}`);
    toast("生成失败");
  } finally {
    els.splitButton.disabled = !state.groups.size;
  }
}

function buildOutputNames() {
  const used = new Map();
  const names = new Map();

  for (const [key, rows] of state.groups.entries()) {
    const base = `${sanitizeFilePart(key)}_${rows.length}`;
    const count = used.get(base) || 0;
    used.set(base, count + 1);
    names.set(key, `${base}${count ? `_${count + 1}` : ""}.xlsx`);
  }

  return names;
}

function columnToIndex(value) {
  const letters = String(value || "").trim().toUpperCase();
  if (!/^[A-Z]+$/.test(letters)) return -1;
  return letters.split("").reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function normalizeName(value, fallback = "未填写") {
  const text = els.trimCheckbox.checked ? String(value ?? "").trim() : String(value ?? "");
  return text || fallback;
}

function sanitizeFilePart(value) {
  return normalizeName(value)
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .slice(0, 80) || "未填写";
}

function safeSheetName(value) {
  return sanitizeFilePart(value).replace(/[\[\]]/g, "_").slice(0, 31) || "Sheet1";
}

function stripExt(filename) {
  return filename.replace(/\.[^.]+$/, "");
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function reset() {
  els.fileInput.value = "";
  els.fileHint.textContent = "支持单工作表数据表；会读取第一个工作表。";
  resetDataOnly();
  state.log = [];
  log("已清空。");
  setStatus("等待选择文件", "默认按 A 列拆分，首行作为标题行保留到每个子表。");
}

function resetDataOnly() {
  state.file = null;
  state.sheetName = "";
  state.rows = [];
  state.groups = new Map();
  els.rowCount.textContent = "0";
  els.groupCount.textContent = "0";
  els.sheetName.textContent = "-";
  els.groupList.textContent = "选择文件后显示每个输出文件名。";
  els.splitButton.disabled = true;
}

async function copyLog() {
  await navigator.clipboard.writeText(state.log.join("\n"));
  toast("记录已复制");
}

function setStatus(title, text) {
  els.statusTitle.textContent = title;
  els.statusText.textContent = text;
}

function log(message) {
  const line = `[${new Date().toLocaleTimeString()}] ${message}`;
  state.log.push(line);
  els.logOutput.textContent = state.log.join("\n");
}

function toast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("show");
  window.clearTimeout(toast.timer);
  toast.timer = window.setTimeout(() => els.toast.classList.remove("show"), 1800);
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function escapeHTML(value) {
  return String(value).replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[char]);
}

init();
