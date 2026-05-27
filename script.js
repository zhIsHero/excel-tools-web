const sidebar = document.querySelector(".tabs");
const toggleButton = document.querySelector(".collapse-button");
const groupToggles = document.querySelectorAll(".group-toggle");
const subTabButtons = document.querySelectorAll(".sub-tab");
const pageEyebrow = document.querySelector("#pageEyebrow");
const pageTitle = document.querySelector("#pageTitle");
const excelActions = document.querySelector("#excelActions");
const renameActions = document.querySelector("#renameActions");
const excelSplitView = document.querySelector("#excelSplitView");
const fileRenameView = document.querySelector("#fileRenameView");
const placeholderView = document.querySelector("#placeholderView");
const placeholderTitle = document.querySelector("#placeholderTitle");
const placeholderText = document.querySelector("#placeholderText");
const excelInput = document.querySelector("#excelInput");
const splitButton = document.querySelector("#splitButton");
const clearButton = document.querySelector("#clearButton");
const fileName = document.querySelector("#fileName");
const sheetCount = document.querySelector("#sheetCount");
const exportStatus = document.querySelector("#exportStatus");
const statusHint = document.querySelector("#statusHint");
const sheetList = document.querySelector("#sheetList");
const folderButton = document.querySelector("#folderButton");
const renameExcelInput = document.querySelector("#renameExcelInput");
const formatMatchInput = document.querySelector("#formatMatchInput");
const renameButton = document.querySelector("#renameButton");
const renameClearButton = document.querySelector("#renameClearButton");
const folderName = document.querySelector("#folderName");
const renameRuleCount = document.querySelector("#renameRuleCount");
const renameStatus = document.querySelector("#renameStatus");
const renameStatusHint = document.querySelector("#renameStatusHint");
const renameList = document.querySelector("#renameList");

let currentWorkbook = null;
let selectedDirectory = null;
let renameRules = [];
let renamePreviewRows = [];
const viewMeta = {
  "excel-split": {
    eyebrow: "Excel Tool",
    title: "Sheet 拆分工具",
    type: "excel-split",
  },
  "file-rename": {
    eyebrow: "Excel Tool",
    title: "批量修改文件名",
    type: "file-rename",
  },
  "edu-manage": {
    eyebrow: "Management",
    title: "教务管理",
    type: "placeholder",
  },
  "exam-manage": {
    eyebrow: "Management",
    title: "考试管理",
    type: "placeholder",
  },
};

toggleButton.addEventListener("click", () => {
  const isCollapsed = sidebar.dataset.collapsed === "true";
  sidebar.dataset.collapsed = String(!isCollapsed);
  toggleButton.setAttribute("aria-expanded", String(isCollapsed));
  toggleButton.setAttribute("aria-label", isCollapsed ? "折叠页签" : "展开页签");
});

groupToggles.forEach((button) => {
  button.addEventListener("click", () => {
    const group = button.closest(".tab-group");
    const isOpen = group.classList.toggle("open");
    button.setAttribute("aria-expanded", String(isOpen));
  });
});

subTabButtons.forEach((button) => {
  button.addEventListener("click", () => {
    const group = button.closest(".tab-group");
    group.classList.add("open");
    group.querySelector(".group-toggle").setAttribute("aria-expanded", "true");
    subTabButtons.forEach((item) => item.classList.remove("active"));
    button.classList.add("active");
    showView(button.dataset.view);
  });
});

excelInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;

  if (!file) {
    resetExcelState();
    return;
  }

  if (!window.XLSX) {
    setStatus("库未加载", "请确认网络可访问 SheetJS CDN 后刷新页面");
    splitButton.disabled = true;
    return;
  }

  setStatus("读取中", "正在解析 Excel 文件");

  try {
    const buffer = await file.arrayBuffer();
    currentWorkbook = XLSX.read(buffer, { type: "array" });
    fileName.textContent = file.name;
    sheetCount.textContent = currentWorkbook.SheetNames.length;
    renderSheetList(currentWorkbook.SheetNames);
    splitButton.disabled = currentWorkbook.SheetNames.length === 0;
    setStatus("已就绪", "点击拆分导出生成文件");
  } catch (error) {
    console.error(error);
    resetExcelState();
    setStatus("读取失败", "请确认文件是有效的 Excel 工作簿");
  }
});

splitButton.addEventListener("click", () => {
  if (!currentWorkbook || currentWorkbook.SheetNames.length === 0) {
    setStatus("待处理", "请先选择一个 Excel 文件");
    return;
  }

  setStatus("导出中", "浏览器会连续下载多个 Excel 文件");

  const usedFileNames = new Map();

  currentWorkbook.SheetNames.forEach((sheetName) => {
    const workbook = XLSX.utils.book_new();
    const worksheet = currentWorkbook.Sheets[sheetName];
    const exportName = getUniqueFileName(sanitizeFileName(sheetName), usedFileNames);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    XLSX.writeFile(workbook, `${exportName}.xlsx`);
  });

  setStatus("已完成", `已生成 ${currentWorkbook.SheetNames.length} 个文件`);
});

clearButton.addEventListener("click", () => {
  excelInput.value = "";
  resetExcelState();
});

folderButton.addEventListener("click", async () => {
  if (!window.showDirectoryPicker) {
    setRenameStatus("不支持", "请使用新版 Chrome 或 Edge 浏览器");
    return;
  }

  try {
    selectedDirectory = await window.showDirectoryPicker({ mode: "readwrite" });
    folderName.textContent = selectedDirectory.name;
    setRenameStatus("已选择", "继续上传 Excel 规则表");
    await refreshRenamePreview();
  } catch (error) {
    if (error.name !== "AbortError") {
      console.error(error);
      setRenameStatus("选择失败", "浏览器没有获得文件夹访问权限");
    }
  }
});

renameExcelInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;

  if (!file) {
    renameRules = [];
    await refreshRenamePreview();
    return;
  }

  if (!window.XLSX) {
    setRenameStatus("库未加载", "xlsx.full.min.js 没有加载成功");
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    renameRules = parseRenameRules(firstSheet);
    renameRuleCount.textContent = renameRules.length;
    setRenameStatus("规则已读取", "确认预览后点击开始改名");
    await refreshRenamePreview();
  } catch (error) {
    console.error(error);
    renameRules = [];
    renameRuleCount.textContent = "0";
    setRenameStatus("读取失败", "请确认 Excel 第一列和第二列是文件名");
    await refreshRenamePreview();
  }
});

formatMatchInput.addEventListener("change", async () => {
  if (renameRules.length > 0 && selectedDirectory) {
    setRenameStatus("规则已更新", formatMatchInput.checked ? "当前按完整文件名匹配" : "当前忽略扩展名匹配");
  }

  await refreshRenamePreview();
});

renameButton.addEventListener("click", async () => {
  if (!selectedDirectory || renameRules.length === 0) {
    setRenameStatus("待处理", "请先选择文件夹和 Excel");
    return;
  }

  renameButton.disabled = true;
  setRenameStatus("处理中", "正在按匹配结果修改文件名");

  const results = [];

  for (const row of renamePreviewRows) {
    if (!row.canRename) {
      results.push({ ...row, status: row.reason, state: "warning" });
      continue;
    }

    try {
      const oldHandle = await selectedDirectory.getFileHandle(row.matchedName);
      const oldFile = await oldHandle.getFile();
      const newHandle = await selectedDirectory.getFileHandle(row.newName, { create: true });
      const writable = await newHandle.createWritable();
      await writable.write(await oldFile.arrayBuffer());
      await writable.close();
      await selectedDirectory.removeEntry(row.matchedName);
      results.push({ ...row, status: "已改名", state: "success" });
    } catch (error) {
      console.error(error);
      results.push({ ...row, status: "处理失败", state: "warning" });
    }
  }

  renderRenameRows(results);
  const successCount = results.filter((row) => row.state === "success").length;
  setRenameStatus("已完成", `成功 ${successCount} 个，跳过 ${results.length - successCount} 个`);
  await refreshRenamePreview(false);
});

renameClearButton.addEventListener("click", () => {
  selectedDirectory = null;
  renameRules = [];
  renamePreviewRows = [];
  renameExcelInput.value = "";
  formatMatchInput.checked = false;
  folderName.textContent = "未选择";
  renameRuleCount.textContent = "0";
  renameButton.disabled = true;
  renameList.innerHTML = '<div class="empty-state">还没有可执行的改名规则</div>';
  setRenameStatus("待处理", "先选择文件夹和 Excel");
});

function renderSheetList(sheetNames) {
  if (sheetNames.length === 0) {
    sheetList.innerHTML = '<div class="empty-state">这个文件里没有可拆分的 sheet</div>';
    return;
  }

  sheetList.innerHTML = sheetNames
    .map(
      (sheetName, index) => `
        <div class="sheet-row">
          <span class="status blue"></span>
          <div>
            <strong>${escapeHtml(sheetName)}</strong>
            <span>将导出为 ${escapeHtml(sanitizeFileName(sheetName))}.xlsx</span>
          </div>
          <code>#${index + 1}</code>
        </div>
      `,
    )
    .join("");
}

function resetExcelState() {
  currentWorkbook = null;
  fileName.textContent = "未选择";
  sheetCount.textContent = "0";
  splitButton.disabled = true;
  sheetList.innerHTML = '<div class="empty-state">还没有读取到 sheet</div>';
  setStatus("待处理", "选择文件后开始拆分");
}

function setStatus(status, hint) {
  exportStatus.textContent = status;
  statusHint.textContent = hint;
}

function setRenameStatus(status, hint) {
  renameStatus.textContent = status;
  renameStatusHint.textContent = hint;
}

function showView(viewName) {
  const meta = viewMeta[viewName] || {
    eyebrow: "Feature",
    title: "暂未实现",
    type: "placeholder",
  };

  pageEyebrow.textContent = meta.eyebrow;
  pageTitle.textContent = meta.title;
  excelActions.classList.remove("active");
  renameActions.classList.remove("active");
  excelSplitView.classList.remove("active");
  fileRenameView.classList.remove("active");
  placeholderView.classList.remove("active");

  if (meta.type === "excel-split") {
    excelActions.classList.add("active");
    excelSplitView.classList.add("active");
    return;
  }

  if (meta.type === "file-rename") {
    renameActions.classList.add("active");
    fileRenameView.classList.add("active");
    return;
  }

  placeholderView.classList.add("active");
  placeholderTitle.textContent = `${meta.title}暂未实现`;
  placeholderText.textContent = "这个功能入口已经放好，后续可以继续接入实际业务逻辑。";
}

async function refreshRenamePreview(shouldRender = true) {
  renameButton.disabled = true;

  if (!selectedDirectory || renameRules.length === 0) {
    if (shouldRender) {
      renameList.innerHTML = '<div class="empty-state">还没有可执行的改名规则</div>';
    }
    return;
  }

  const existingNames = new Set();
  const nameIndex = new Map();
  const shouldMatchExtension = formatMatchInput.checked;

  for await (const [name, handle] of selectedDirectory.entries()) {
    if (handle.kind === "file") {
      existingNames.add(name);
      const key = shouldMatchExtension ? name : baseFileName(name);
      const matches = nameIndex.get(key) || [];
      matches.push(name);
      nameIndex.set(key, matches);
    }
  }

  const targetCounts = new Map();
  const normalizedRules = renameRules.map((rule) => {
    const matchKey = shouldMatchExtension ? rule.oldName : baseFileName(rule.oldName);
    const matches = nameIndex.get(matchKey) || [];
    const matchedName = matches.length === 1 ? matches[0] : "";
    const newName = normalizeNewFileName(matchedName || rule.oldName, rule.newName);
    targetCounts.set(newName, (targetCounts.get(newName) || 0) + 1);
    return { ...rule, matchedName, matches, newName };
  });

  renamePreviewRows = normalizedRules.map((rule) => {
    const duplicateReason = targetCounts.get(rule.newName) > 1 ? "目标名重复" : "";
    const reason =
      duplicateReason ||
      getRenameBlockReason(rule.oldName, rule.matchedName, rule.matches, rule.newName, existingNames);
    return {
      oldName: rule.oldName,
      matchedName: rule.matchedName,
      newName: rule.newName,
      canRename: !reason,
      reason: reason || "可改名",
      state: reason ? "warning" : "success",
    };
  });

  renameButton.disabled = !renamePreviewRows.some((row) => row.canRename);

  if (shouldRender) {
    renderRenameRows(renamePreviewRows);
  }
}

function parseRenameRules(worksheet) {
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

  return rows
    .map((row) => ({
      oldName: String(row[0] || "").trim(),
      newName: String(row[1] || "").trim(),
    }))
    .filter((row) => row.oldName && row.newName);
}

function normalizeNewFileName(oldName, newName) {
  if (pathExtension(newName)) {
    return newName;
  }

  return `${newName}${pathExtension(oldName)}`;
}

function getRenameBlockReason(oldName, matchedName, matches, newName, existingNames) {
  if (matches.length > 1) {
    return "匹配到多个文件";
  }

  if (!matchedName) {
    return "未找到原文件";
  }

  if (matchedName === newName) {
    return "文件名未变化";
  }

  if (hasInvalidFileNameChars(newName)) {
    return "新文件名非法";
  }

  if (existingNames.has(newName) && newName !== matchedName) {
    return "目标名已存在";
  }

  return "";
}

function renderRenameRows(rows) {
  if (rows.length === 0) {
    renameList.innerHTML = '<div class="empty-state">还没有可执行的改名规则</div>';
    return;
  }

  renameList.innerHTML = rows
    .map(
      (row) => `
        <div class="rename-row ${row.state}">
          <span class="status ${row.state === "success" ? "green" : "orange"}"></span>
          <div>
            <strong>${escapeHtml(row.matchedName || row.oldName)} → ${escapeHtml(row.newName)}</strong>
            <span>${escapeHtml(renameRowHint(row))}</span>
          </div>
          <code>${escapeHtml(row.status || row.reason)}</code>
        </div>
      `,
    )
    .join("");
}

function hasInvalidFileNameChars(name) {
  return /[\\/:*?"<>|]/.test(name) || name.endsWith(".") || name.trim() !== name || name.length === 0;
}

function pathExtension(fileName) {
  const dotIndex = fileName.lastIndexOf(".");
  return dotIndex > 0 ? fileName.slice(dotIndex) : "";
}

function baseFileName(fileName) {
  const dotIndex = fileName.lastIndexOf(".");
  return dotIndex > 0 ? fileName.slice(0, dotIndex) : fileName;
}

function renameRowHint(row) {
  if (row.status) {
    return row.reason === "可改名" ? "已完成处理" : row.reason;
  }

  if (row.matchedName && row.matchedName !== row.oldName) {
    return `${row.reason}，Excel 原值：${row.oldName}`;
  }

  return row.reason;
}

function sanitizeFileName(name) {
  return name.replace(/[\\/:*?"<>|]/g, "_").trim() || "sheet";
}

function getUniqueFileName(name, usedFileNames) {
  const count = usedFileNames.get(name) || 0;
  usedFileNames.set(name, count + 1);
  return count === 0 ? name : `${name}_${count + 1}`;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
