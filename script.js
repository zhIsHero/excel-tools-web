const sidebar = document.querySelector(".tabs");
const toggleButton = document.querySelector(".collapse-button");
const groupToggles = document.querySelectorAll(".group-toggle");
const subTabButtons = document.querySelectorAll(".sub-tab");
const pageEyebrow = document.querySelector("#pageEyebrow");
const pageTitle = document.querySelector("#pageTitle");
const excelActions = document.querySelector("#excelActions");
const excelSplitView = document.querySelector("#excelSplitView");
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

let currentWorkbook = null;
const viewMeta = {
  "excel-split": {
    eyebrow: "Excel Tool",
    title: "Sheet 拆分工具",
    implemented: true,
  },
  "excel-merge": {
    eyebrow: "Excel Tool",
    title: "Excel 合并",
    implemented: false,
  },
  "edu-manage": {
    eyebrow: "Management",
    title: "教务管理",
    implemented: false,
  },
  "exam-manage": {
    eyebrow: "Management",
    title: "考试管理",
    implemented: false,
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

function showView(viewName) {
  const meta = viewMeta[viewName] || {
    eyebrow: "Feature",
    title: "暂未实现",
    implemented: false,
  };

  pageEyebrow.textContent = meta.eyebrow;
  pageTitle.textContent = meta.title;

  if (meta.implemented) {
    excelActions.hidden = false;
    excelSplitView.classList.add("active");
    placeholderView.classList.remove("active");
    return;
  }

  excelActions.hidden = true;
  excelSplitView.classList.remove("active");
  placeholderView.classList.add("active");
  placeholderTitle.textContent = `${meta.title}暂未实现`;
  placeholderText.textContent = "这个功能入口已经放好，后续可以继续接入实际业务逻辑。";
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
