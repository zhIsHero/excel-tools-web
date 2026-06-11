const sidebar = document.querySelector(".tabs");
const toggleButton = document.querySelector(".collapse-button");
const entryGate = document.querySelector("#entryGate");
const entryGateButton = document.querySelector("#entryGateButton");
const groupToggles = document.querySelectorAll(".group-toggle");
const subTabButtons = document.querySelectorAll(".sub-tab");
const pageEyebrow = document.querySelector("#pageEyebrow");
const pageTitle = document.querySelector("#pageTitle");
const excelActions = document.querySelector("#excelActions");
const renameActions = document.querySelector("#renameActions");
const scheduleActions = document.querySelector("#scheduleActions");
const surveyActions = document.querySelector("#surveyActions");
const convertActions = document.querySelector("#convertActions");
const contactsActions = document.querySelector("#contactsActions");
const excelSplitView = document.querySelector("#excelSplitView");
const fileRenameView = document.querySelector("#fileRenameView");
const scheduleModifyView = document.querySelector("#scheduleModifyView");
const surveyView = document.querySelector("#surveyView");
const convertView = document.querySelector("#convertView");
const teacherContactsView = document.querySelector("#teacherContactsView");
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
const scheduleInput = document.querySelector("#scheduleInput");
const scheduleExportButton = document.querySelector("#scheduleExportButton");
const scheduleClearButton = document.querySelector("#scheduleClearButton");
const scheduleFileName = document.querySelector("#scheduleFileName");
const scheduleClassCount = document.querySelector("#scheduleClassCount");
const scheduleStatus = document.querySelector("#scheduleStatus");
const scheduleStatusHint = document.querySelector("#scheduleStatusHint");
const scheduleTitleInput = document.querySelector("#scheduleTitleInput");
const customFooterInput = document.querySelector("#customFooterInput");
const footerLineInputs = document.querySelectorAll(".footer-line-input");
const teacherList = document.querySelector("#teacherList");
const surveySourceInput = document.querySelector("#surveySourceInput");
const surveyTeacherInput = document.querySelector("#surveyTeacherInput");
const surveyTemplateInput = document.querySelector("#surveyTemplateInput");
const surveyExportButton = document.querySelector("#surveyExportButton");
const surveyClearButton = document.querySelector("#surveyClearButton");
const surveySourceName = document.querySelector("#surveySourceName");
const surveyTeacherName = document.querySelector("#surveyTeacherName");
const surveyStatus = document.querySelector("#surveyStatus");
const surveyStatusHint = document.querySelector("#surveyStatusHint");
const surveyEdgeWeightsInput = document.querySelector("#surveyEdgeWeightsInput");
const surveyMiddleWeightsInput = document.querySelector("#surveyMiddleWeightsInput");
const surveyGradeConfigInput = document.querySelector("#surveyGradeConfigInput");
const surveySubjectConfigInput = document.querySelector("#surveySubjectConfigInput");
const convertInput = document.querySelector("#convertInput");
const convertButton = document.querySelector("#convertButton");
const convertClearButton = document.querySelector("#convertClearButton");
const convertFileName = document.querySelector("#convertFileName");
const convertModeLabel = document.querySelector("#convertModeLabel");
const convertStatus = document.querySelector("#convertStatus");
const convertStatusHint = document.querySelector("#convertStatusHint");
const convertModeInputs = document.querySelectorAll('input[name="convertMode"]');
const convertModeCards = document.querySelectorAll(".convert-mode");
const contactsInput = document.querySelector("#contactsInput");
const contactsClearButton = document.querySelector("#contactsClearButton");
const contactsFileName = document.querySelector("#contactsFileName");
const contactsCount = document.querySelector("#contactsCount");
const contactsStatus = document.querySelector("#contactsStatus");
const contactsStatusHint = document.querySelector("#contactsStatusHint");
const contactsList = document.querySelector("#contactsList");

let currentWorkbook = null;
let selectedDirectory = null;
let renameRules = [];
let renamePreviewRows = [];
let scheduleWorkbook = null;
let scheduleSheetNames = [];
let scheduleTemplateWorkbook = null;
let teacherInfoBySheet = new Map();
let surveySourceWorkbook = null;
let surveyTeacherWorkbook = null;
let surveyTemplateBuffer = null;
let convertFile = null;
let teacherContacts = [];
let teacherContactMap = new Map();
const teacherContactsStorageKey = "hangge.teacherContacts.v1";
const defaultScheduleFooterLines = [
  "《习近平新时代中国特色社会主义思想》课程在朝会、班团活动、思政课上分别以1/3节课的时间进行教学。",
  "同一节课有两个科目为单双周上课，单周上第一排科目，双周上第二排科目",
  "保卫科联系电话：0817-3599008。               2025.11.21",
];
const defaultSurveyGradeConfig = [
  { id: "1", teacherLabel: "初2025级", displayName: "初一", outputName: "初2025级" },
  { id: "2", teacherLabel: "初2024级", displayName: "初二", outputName: "初2024级" },
  { id: "3", teacherLabel: "初2023级", displayName: "初三", outputName: "初2023级" },
  { id: "4", teacherLabel: "高2025级", displayName: "高一", outputName: "高2025级" },
  { id: "5", teacherLabel: "高2024级", displayName: "高二", outputName: "高2024级" },
  { id: "6", teacherLabel: "高2023级", displayName: "高三", outputName: "高2023级" },
];
const defaultSurveySubjectConfig = [
  { name: "语文", type: "five", startColumn: 29 },
  { name: "数学", type: "five", startColumn: 34 },
  { name: "英语", teacherSubject: "外语", type: "five", startColumn: 39 },
  { name: "物理", type: "five", startColumn: 44 },
  { name: "化学", type: "five", startColumn: 49 },
  { name: "生物", type: "five", startColumn: 54 },
  { name: "政治", type: "five", startColumn: 59 },
  { name: "历史", type: "five", startColumn: 64 },
  { name: "地理", type: "five", startColumn: 69 },
  { name: "体育", type: "single", startColumn: 74 },
  { name: "美术", type: "single", startColumn: 75 },
  { name: "音乐", type: "single", startColumn: 76 },
  { name: "心理", teacherSubject: "心理、生涯", type: "single", startColumn: 77 },
  { name: "信息", teacherSubject: "信息（通用）", type: "single", startColumn: 78 },
  { name: "创新", type: "single", startColumn: 79 },
  { name: "劳技", type: "single", startColumn: 80 },
  { name: "书法", type: "single", startColumn: 81 },
  { name: "班主任", type: "single", startColumn: 82 },
];

entryGateButton.addEventListener("click", () => {
  entryGate.classList.add("unlocked");
});

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
  "schedule-modify": {
    eyebrow: "Excel Tool",
    title: "班级课表修改",
    type: "schedule-modify",
  },
  "midterm-survey": {
    eyebrow: "Excel Tool",
    title: "中期问卷处理",
    type: "midterm-survey",
  },
  "pdf-word-convert": {
    eyebrow: "Document Tool",
    title: "PDF与Word转换",
    type: "pdf-word-convert",
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
  "teacher-contacts": {
    eyebrow: "Management",
    title: "教师联系方式",
    type: "teacher-contacts",
  },
};

loadStoredTeacherContacts();

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

contactsInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;

  if (!file) {
    return;
  }

  if (!window.XLSX) {
    setContactsStatus("库未加载", "xlsx.full.min.js 没有加载成功");
    return;
  }

  try {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
    const contacts = parseTeacherContactsWorkbook(workbook);

    teacherContacts = contacts;
    rebuildTeacherContactMap();
    saveTeacherContacts(file.name);
    renderTeacherContacts(file.name);
    updateTeacherPhonesFromContacts();
    setContactsStatus("已保存", `已导入 ${contacts.length} 位教师`);
  } catch (error) {
    console.error(error);
    setContactsStatus("导入失败", "请确认表格列顺序为姓名、性别、身份证号、出生日期、年龄、电话、学科");
  }
});

contactsClearButton.addEventListener("click", () => {
  localStorage.removeItem(teacherContactsStorageKey);
  teacherContacts = [];
  teacherContactMap = new Map();
  contactsInput.value = "";
  renderTeacherContacts("");
  updateTeacherPhonesFromContacts();
  setContactsStatus("已清空", "已移除本地保存的教师联系方式");
});

resetScheduleFooterInputs();
surveyGradeConfigInput.value = JSON.stringify(defaultSurveyGradeConfig, null, 2);
surveySubjectConfigInput.value = JSON.stringify(defaultSurveySubjectConfig, null, 2);

customFooterInput.addEventListener("change", () => {
  footerLineInputs.forEach((input) => {
    input.disabled = !customFooterInput.checked;
  });

  if (!customFooterInput.checked) {
    resetScheduleFooterInputs();
  }
});

scheduleInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;

  if (!file) {
    resetScheduleState();
    return;
  }

  if (!window.XLSX) {
    setScheduleStatus("库未加载", "xlsx.full.min.js 没有加载成功");
    return;
  }

  setScheduleStatus("读取中", "正在解析课程表");

  try {
    const buffer = await file.arrayBuffer();
    scheduleWorkbook = XLSX.read(buffer, { type: "array" });
    scheduleSheetNames = scheduleWorkbook.SheetNames.filter((sheetName) => {
      const worksheet = scheduleWorkbook.Sheets[sheetName];
      return worksheet && worksheet["!ref"] && sheetName.toLowerCase() !== "sheet1";
    });
    teacherInfoBySheet = new Map();
    renderTeacherRows();
    scheduleFileName.textContent = file.name;
    scheduleClassCount.textContent = scheduleSheetNames.length;
    scheduleExportButton.disabled = scheduleSheetNames.length === 0;
    setScheduleStatus("已就绪", "点击生成修改版导出 xlsx");
  } catch (error) {
    console.error(error);
    resetScheduleState();
    setScheduleStatus("读取失败", "请确认文件是有效的课程表 Excel");
  }
});

scheduleExportButton.addEventListener("click", async () => {
  if (!scheduleWorkbook || scheduleSheetNames.length === 0) {
    setScheduleStatus("待处理", "请先上传课程表");
    return;
  }

  if (!window.ExcelJS) {
    setScheduleStatus("库未加载", "exceljs.min.js 没有加载成功");
    return;
  }

  scheduleExportButton.disabled = true;
  setScheduleStatus("生成中", "正在套用课程表修改版模板");

  try {
    scheduleTemplateWorkbook = scheduleTemplateWorkbook || (await loadScheduleTemplateWorkbook());
  } catch (error) {
    console.error(error);
    scheduleExportButton.disabled = false;
    setScheduleStatus("模板加载失败", "请确认 schedule-template.xlsx 已随网站一起部署");
    return;
  }

  try {
    const outputWorkbook = await createScheduleOutputWorkbook();
    const outputBuffer = await outputWorkbook.xlsx.writeBuffer();
    const paginatedBuffer = await addScheduleColumnBreaks(outputBuffer);
    downloadBuffer(paginatedBuffer, "课程表修改版.xlsx");
    setScheduleStatus("已完成", `已生成 ${scheduleSheetNames.length} 个班级 sheet`);
  } catch (error) {
    console.error(error);
    setScheduleStatus("生成失败", "请检查输入课程表是否符合模板结构");
  } finally {
    scheduleExportButton.disabled = false;
  }
});

scheduleClearButton.addEventListener("click", () => {
  scheduleInput.value = "";
  resetScheduleState();
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

function setScheduleStatus(status, hint) {
  scheduleStatus.textContent = status;
  scheduleStatusHint.textContent = hint;
}

function setContactsStatus(status, hint) {
  contactsStatus.textContent = status;
  contactsStatusHint.textContent = hint;
}

function resetScheduleState() {
  scheduleWorkbook = null;
  scheduleSheetNames = [];
  teacherInfoBySheet = new Map();
  renderTeacherRows();
  scheduleFileName.textContent = "未选择";
  scheduleClassCount.textContent = "0";
  scheduleExportButton.disabled = true;
  setScheduleStatus("待处理", "上传课程表后生成");
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
  scheduleActions.classList.remove("active");
  surveyActions.classList.remove("active");
  convertActions.classList.remove("active");
  contactsActions.classList.remove("active");
  excelSplitView.classList.remove("active");
  fileRenameView.classList.remove("active");
  scheduleModifyView.classList.remove("active");
  surveyView.classList.remove("active");
  convertView.classList.remove("active");
  teacherContactsView.classList.remove("active");
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

  if (meta.type === "schedule-modify") {
    scheduleActions.classList.add("active");
    scheduleModifyView.classList.add("active");
    return;
  }

  if (meta.type === "midterm-survey") {
    surveyActions.classList.add("active");
    surveyView.classList.add("active");
    return;
  }

  if (meta.type === "pdf-word-convert") {
    convertActions.classList.add("active");
    convertView.classList.add("active");
    return;
  }

  if (meta.type === "teacher-contacts") {
    contactsActions.classList.add("active");
    teacherContactsView.classList.add("active");
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

async function createScheduleOutputWorkbook() {
  const outputWorkbook = new ExcelJS.Workbook();
  const templateSheet = scheduleTemplateWorkbook.worksheets[0];

  scheduleSheetNames.forEach((sheetName) => {
    const worksheet = outputWorkbook.addWorksheet(sheetName.slice(0, 31));
    cloneTemplateWorksheet(templateSheet, worksheet, sheetName);
    fillScheduleTemplateSheet(worksheet, scheduleWorkbook.Sheets[sheetName], sheetName);
  });

  return outputWorkbook;
}

function fillScheduleTemplateSheet(worksheet, sourceSheet, sheetName) {
  const classNumber = extractClassNumber(sheetName);
  const teacher = teacherInfoBySheet.get(sheetName) || {};
  const teacherDetails = [teacher.name, teacher.phone].filter(Boolean).join(":");
  const teacherInfo = teacherDetails ? `${classNumber || sheetName}班主任 ${teacherDetails}` : "";
  const campusAndTeacher = teacherInfo
    ? `南高高坪校区${" ".repeat(17)}${teacherInfo}`
    : "南高高坪校区";
  const title = scheduleTitleInput.value.trim() || "高2025级2025年秋季班级课表11.21";
  const footerText = formatScheduleFooter(Array.from(footerLineInputs, (input) => input.value));

  setCell(worksheet, 0, 0, "     课　　程　　表    ");
  setCell(worksheet, 0, 10, "     课　　程　　表    ");
  setCell(worksheet, 1, 0, title);
  setCell(worksheet, 1, 10, title);
  setCell(worksheet, 2, 0, campusAndTeacher);
  setCell(worksheet, 2, 9, sheetName);
  setCell(worksheet, 2, 10, campusAndTeacher);
  setCell(worksheet, 2, 17, sheetName);

  setScheduleHeaderText(worksheet, "A4", "课　　　　星\n　　程　　期\n\n时间");
  setScheduleHeaderText(worksheet, "K4", "　　课　　　　星\n　　　　程　　期\n时间");

  ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"].forEach((day, index) => {
    setCell(worksheet, 3, 3 + index, day);
  });

  ["星期一", "星期二", "星期三", "星期四", "星期五"].forEach((day, index) => {
    setCell(worksheet, 3, 13 + index, day);
  });

  fillScheduleSection(worksheet, sourceSheet, {
    label: "上\n\n午",
    sourceStartRow: 3,
    outputStartRow: 4,
    count: 5,
    periodStart: 1,
  });
  fillScheduleSection(worksheet, sourceSheet, {
    label: "下\n\n午",
    sourceStartRow: 8,
    outputStartRow: 10,
    count: 4,
    periodStart: 1,
  });
  fillScheduleSection(worksheet, sourceSheet, {
    label: "托管服务",
    sourceStartRow: 12,
    outputStartRow: 15,
    count: 3,
    periodStart: 1,
  });

  setCell(worksheet, 18, 0, footerText);
  setCell(worksheet, 18, 10, footerText);
  worksheet.getRow(19).height = scheduleFooterRowHeight(footerText);
  worksheet.getCell("A19").alignment = {
    ...worksheet.getCell("A19").alignment,
    horizontal: "left",
    vertical: "top",
    wrapText: true,
  };
  worksheet.getCell("K19").alignment = {
    ...worksheet.getCell("K19").alignment,
    horizontal: "left",
    vertical: "top",
    wrapText: true,
  };

  return worksheet;
}

async function loadScheduleTemplateWorkbook() {
  const templateBuffer = await loadScheduleTemplateBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);
  return workbook;
}

async function loadScheduleTemplateBuffer() {
  try {
    const response = await fetch("schedule-template.xlsx");

    if (response.ok) {
      return await response.arrayBuffer();
    }
  } catch (error) {
    console.info("Falling back to embedded schedule template.", error);
  }

  if (!window.SCHEDULE_TEMPLATE_XLSX_BASE64) {
    throw new Error("Missing embedded schedule template.");
  }

  return base64ToArrayBuffer(window.SCHEDULE_TEMPLATE_XLSX_BASE64);
}

function cloneValue(value) {
  return JSON.parse(JSON.stringify(value));
}

function cloneTemplateWorksheet(templateSheet, worksheet, sheetName) {
  const templateModel = cloneValue(templateSheet.model);
  const merges = templateModel.merges || [];
  worksheet.model = {
    ...templateModel,
    name: sheetName.slice(0, 31),
    merges: [],
  };
  worksheet.name = sheetName.slice(0, 31);
  merges.forEach((range) => worksheet.mergeCells(range));
  worksheet.pageSetup = {
    ...cloneValue(templateSheet.pageSetup),
    paperSize: 9,
    orientation: "portrait",
    scale: 100,
    fitToPage: false,
    fitToWidth: undefined,
    fitToHeight: undefined,
    horizontalCentered: true,
    verticalCentered: false,
    margins: {
      left: 1.5 / 2.54,
      right: 1.5 / 2.54,
      top: 1.8 / 2.54,
      bottom: 1.8 / 2.54,
      header: 0.5,
      footer: 0.5,
    },
  };
}

function fillScheduleSection(worksheet, sourceSheet, options) {
  const { label, sourceStartRow, outputStartRow, count, periodStart } = options;
  setCell(worksheet, outputStartRow, 0, label);
  setCell(worksheet, outputStartRow, 10, label);

  for (let index = 0; index < count; index += 1) {
    const sourceRow = sourceStartRow + index;
    const outputRow = outputStartRow + index;
    setCell(worksheet, outputRow, 1, String(periodStart + index));
    setCell(worksheet, outputRow, 11, String(periodStart + index));

    for (let sourceCol = 3; sourceCol <= 9; sourceCol += 1) {
      setCell(worksheet, outputRow, sourceCol, sourceValue(sourceSheet, sourceRow, sourceCol));
    }

    for (let sourceCol = 3; sourceCol <= 7; sourceCol += 1) {
      setCell(worksheet, outputRow, 10 + sourceCol, sourceValue(sourceSheet, sourceRow, sourceCol));
    }
  }
}

function sourceValue(worksheet, row, col) {
  const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: col })];
  return cell ? cell.v ?? cell.w ?? "" : "";
}

function setCell(worksheet, row, col, value) {
  if (value === undefined || value === null) {
    return;
  }

  const cell = worksheet.getCell(row + 1, col + 1);
  cell.value = value;

  if (isScheduleCourseCell(row + 1, col + 1)) {
    updateScheduleCourseCellFit(cell);
  }
}

function isScheduleCourseCell(row, col) {
  const isCourseRow = row >= 5 && row <= 17;
  const isLeftCourseColumn = col >= 4 && col <= 10;
  const isRightCourseColumn = col >= 14 && col <= 18;
  return isCourseRow && (isLeftCourseColumn || isRightCourseColumn);
}

function updateScheduleCourseCellFit(cell) {
  const value = String(cell.value || "");
  const lines = value.split(/\r?\n/);
  const longestLineLength = Math.max(...lines.map((line) => line.trim().length), 0);
  const totalTextLength = lines.join("").trim().length;
  const shouldShrink = lines.length > 2 || longestLineLength > 6 || totalTextLength > 12;

  cell.font = {
    ...cell.font,
    size: shouldShrink ? 9 : 12,
  };
  cell.alignment = {
    ...cell.alignment,
    horizontal: "center",
    vertical: "middle",
    wrapText: true,
    shrinkToFit: shouldShrink,
  };
}

function setScheduleHeaderText(worksheet, address, value) {
  const cell = worksheet.getCell(address);
  cell.value = value;
  cell.font = {
    ...cell.font,
    name: "宋体",
    size: 9,
  };
  cell.alignment = {
    ...cell.alignment,
    horizontal: "left",
    vertical: "top",
    wrapText: true,
  };
  worksheet.getRow(cell.row).height = Math.max(worksheet.getRow(cell.row).height || 0, 52);
}

function resetScheduleFooterInputs() {
  footerLineInputs.forEach((input, index) => {
    input.value = defaultScheduleFooterLines[index] || "";
  });
}

function renderTeacherRows() {
  if (scheduleSheetNames.length === 0) {
    teacherList.innerHTML = '<div class="empty-state">上传课程表后显示班级列表</div>';
    return;
  }

  teacherList.innerHTML = "";

  scheduleSheetNames.forEach((sheetName) => {
    const teacher = teacherInfoBySheet.get(sheetName) || { name: "", phone: "" };
    teacherInfoBySheet.set(sheetName, teacher);

    const row = document.createElement("div");
    row.className = "teacher-row";

    const label = document.createElement("strong");
    label.textContent = sheetName;

    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.placeholder = "班主任姓名";
    nameInput.value = teacher.name;
    nameInput.addEventListener("input", () => {
      teacher.name = nameInput.value.trim();
      const matchedContact = findTeacherContact(teacher.name);
      teacher.phone = matchedContact?.phone || "";
      phoneInput.value = teacher.phone;
    });

    const phoneInput = document.createElement("input");
    phoneInput.type = "text";
    phoneInput.placeholder = "联系电话";
    phoneInput.value = teacher.phone;
    phoneInput.addEventListener("input", () => {
      teacher.phone = phoneInput.value.trim();
    });

    row.append(label, nameInput, phoneInput);
    teacherList.append(row);
  });
}

function parseTeacherContactsWorkbook(workbook) {
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
  const contacts = [];

  rows.forEach((row, index) => {
    const name = String(row[0] || "").trim();

    if (!name || (index === 0 && name.includes("姓名"))) {
      return;
    }

    const phone = String(row[5] || "").trim();
    contacts.push({
      name,
      gender: String(row[1] || "").trim(),
      idNumber: String(row[2] || "").trim(),
      birthDate: formatContactCell(row[3]),
      age: String(row[4] || "").trim(),
      phone,
      subject: String(row[6] || "").trim(),
    });
  });

  if (contacts.length === 0) {
    throw new Error("No contacts found");
  }

  return contacts;
}

function formatContactCell(value) {
  if (value instanceof Date) {
    return value.toLocaleDateString("zh-CN");
  }
  return String(value || "").trim();
}

function rebuildTeacherContactMap() {
  teacherContactMap = new Map();
  teacherContacts.forEach((contact) => {
    if (contact.name) {
      teacherContactMap.set(normalizeTeacherName(contact.name), contact);
    }
  });
}

function normalizeTeacherName(name) {
  return String(name || "").replace(/\s+/g, "").trim();
}

function findTeacherContact(name) {
  return teacherContactMap.get(normalizeTeacherName(name));
}

function saveTeacherContacts(fileName) {
  if (typeof localStorage === "undefined") {
    return;
  }

  localStorage.setItem(
    teacherContactsStorageKey,
    JSON.stringify({
      fileName,
      contacts: teacherContacts,
      savedAt: new Date().toISOString(),
    }),
  );
}

function loadStoredTeacherContacts() {
  if (typeof localStorage === "undefined") {
    renderTeacherContacts("");
    return;
  }

  try {
    const raw = localStorage.getItem(teacherContactsStorageKey);

    if (!raw) {
      renderTeacherContacts("");
      return;
    }

    const saved = JSON.parse(raw);
    teacherContacts = Array.isArray(saved.contacts) ? saved.contacts : [];
    rebuildTeacherContactMap();
    renderTeacherContacts(saved.fileName || "");
    setContactsStatus("已加载", `已自动加载 ${teacherContacts.length} 位教师`);
  } catch (error) {
    console.error(error);
    teacherContacts = [];
    teacherContactMap = new Map();
    renderTeacherContacts("");
    setContactsStatus("加载失败", "本地保存的教师联系方式无法读取");
  }
}

function renderTeacherContacts(fileName) {
  contactsFileName.textContent = fileName || "未选择";
  contactsCount.textContent = String(teacherContacts.length);

  if (teacherContacts.length === 0) {
    contactsList.innerHTML = '<div class="empty-state">还没有导入教师联系方式</div>';
    if (!fileName) {
      setContactsStatus("待导入", "请选择教师联系方式 Excel");
    }
    return;
  }

  contactsList.innerHTML = teacherContacts
    .slice(0, 60)
    .map(
      (contact) => `
        <div class="contact-row">
          <strong>${escapeHtml(contact.name)}</strong>
          <span>${escapeHtml(contact.phone || "无电话")}</span>
          <span>${escapeHtml(contact.subject || "未填学科")}</span>
        </div>
      `,
    )
    .join("");
}

function updateTeacherPhonesFromContacts() {
  teacherInfoBySheet.forEach((teacher) => {
    const matchedContact = findTeacherContact(teacher.name);
    teacher.phone = matchedContact?.phone || "";
  });

  if (scheduleSheetNames.length > 0) {
    renderTeacherRows();
  }
}

function formatScheduleFooter(values) {
  const lines = values
    .map((line) => line.replace(/^\s*(?:备注[:：]?\s*)?\d*[、.．]?\s*/, "").trim())
    .filter(Boolean);
  const footerLines = lines.length > 0 ? lines : defaultScheduleFooterLines;

  return footerLines
    .map((line, index) => `${index === 0 ? "备注： " : "      "}${index + 1}、${line}`)
    .join("\n");
}

function scheduleFooterRowHeight(footerText) {
  return 100;
}

function extractClassNumber(sheetName) {
  const match = String(sheetName).match(/(\d+)\s*班/);
  return match ? match[1] : "";
}

function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return bytes.buffer;
}

function downloadBuffer(buffer, fileName) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

async function addScheduleColumnBreaks(buffer) {
  if (!window.JSZip) {
    throw new Error("jszip.min.js 没有加载成功");
  }

  const zip = await JSZip.loadAsync(buffer);
  const templateZip = await JSZip.loadAsync(base64ToArrayBuffer(window.SCHEDULE_TEMPLATE_XLSX_BASE64));
  const drawingXml = await templateZip.file("xl/drawings/drawing1.xml").async("string");
  const worksheetPaths = Object.keys(zip.files).filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/.test(path));
  const breaks =
    '<colBreaks count="2" manualBreakCount="2">' +
    '<brk id="10" min="0" max="1048575" man="1"/>' +
    '<brk id="18" min="0" max="1048575" man="1"/>' +
    "</colBreaks>";

  await Promise.all(
    worksheetPaths.map(async (path, index) => {
      const sheetNumber = index + 1;
      const worksheetXml = await zip.file(path).async("string");
      const cleanedXml = worksheetXml.replace(/<colBreaks[\s\S]*?<\/colBreaks>/, "");
      const withoutDrawing = cleanedXml.replace(/<drawing[^>]*\/>/, "");
      const printScaledXml = normalizeSchedulePageSetupXml(withoutDrawing);
      const updatedXml = printScaledXml.replace(
        "</worksheet>",
        `${breaks}<drawing r:id="rId1"/></worksheet>`,
      );
      zip.file(path, updatedXml);
      zip.file(`xl/drawings/drawing${sheetNumber}.xml`, drawingXml);
      zip.file(
        `xl/worksheets/_rels/sheet${sheetNumber}.xml.rels`,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" ' +
          `Target="../drawings/drawing${sheetNumber}.xml"/>` +
          "</Relationships>",
      );
    }),
  );

  const contentTypesPath = "[Content_Types].xml";
  let contentTypesXml = await zip.file(contentTypesPath).async("string");
  worksheetPaths.forEach((_, index) => {
    const drawingPart = `/xl/drawings/drawing${index + 1}.xml`;
    if (!contentTypesXml.includes(`PartName="${drawingPart}"`)) {
      contentTypesXml = contentTypesXml.replace(
        "</Types>",
        `<Override PartName="${drawingPart}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`,
      );
    }
  });
  zip.file(contentTypesPath, contentTypesXml);

  return zip.generateAsync({ type: "arraybuffer" });
}

function normalizeSchedulePageSetupXml(worksheetXml) {
  return worksheetXml
    .replace(/(<pageSetUpPr\b[^>]*?)\s+fitToPage="1"/, '$1 fitToPage="0"')
    .replace(/<pageSetup\b[^>]*>/, (pageSetupXml) => {
      let xml = pageSetupXml
        .replace(/\s+fitToWidth="[^"]*"/g, "")
        .replace(/\s+fitToHeight="[^"]*"/g, "");

      if (/\s+scale="[^"]*"/.test(xml)) {
        xml = xml.replace(/\s+scale="[^"]*"/, ' scale="100"');
      } else {
        xml = xml.replace("<pageSetup", '<pageSetup scale="100"');
      }

      return xml;
    });
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

surveySourceInput.addEventListener("change", async (event) => {
  surveySourceWorkbook = await readSurveyWorkbook(event.target.files[0], "问卷数据");
  surveySourceName.textContent = event.target.files[0]?.name || "未选择";
  refreshSurveyReadyState();
});

surveyTeacherInput.addEventListener("change", async (event) => {
  surveyTeacherWorkbook = await readSurveyWorkbook(event.target.files[0], "教师任课表");
  surveyTeacherName.textContent = event.target.files[0]?.name || "未选择";
  refreshSurveyReadyState();
});

surveyTemplateInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  surveyTemplateBuffer = file ? await file.arrayBuffer() : null;
});

surveyClearButton.addEventListener("click", () => {
  surveySourceInput.value = "";
  surveyTeacherInput.value = "";
  surveyTemplateInput.value = "";
  surveySourceWorkbook = null;
  surveyTeacherWorkbook = null;
  surveyTemplateBuffer = null;
  surveySourceName.textContent = "未选择";
  surveyTeacherName.textContent = "未选择";
  refreshSurveyReadyState();
});

surveyExportButton.addEventListener("click", async () => {
  if (!surveySourceWorkbook || !surveyTeacherWorkbook) {
    setSurveyStatus("待处理", "请选择问卷数据和教师任课表");
    return;
  }

  surveyExportButton.disabled = true;
  setSurveyStatus("处理中", "正在统计问卷并套用模板");

  try {
    const config = readSurveyConfig();
    const templateBuffer = await loadSurveyTemplateBuffer();
    const templateWorkbook = new ExcelJS.Workbook();
    await templateWorkbook.xlsx.load(templateBuffer);
    const teacherData = parseSurveyTeacherWorkbook(config);
    const stats = parseSurveySourceWorkbook(config);
    const zip = new JSZip();
    const summary = createSurveySummaryState();

    for (const grade of config.grades) {
      const classNumbers = Array.from(stats.classes.get(grade.id) || []).sort(compareSurveyClassNames);

      if (classNumbers.length === 0) {
        continue;
      }

      const workbook = await createSurveyGradeWorkbook(grade, classNumbers, stats, teacherData, summary, templateWorkbook, config);
      zip.file(`${grade.displayName}.xlsx`, await workbook.xlsx.writeBuffer());
    }

    const summaryWorkbook = createSurveySummaryWorkbook(summary, teacherData, config);
    zip.file("汇总.xlsx", await summaryWorkbook.xlsx.writeBuffer());
    downloadBlob(await zip.generateAsync({ type: "blob" }), "中期问卷处理结果.zip");
    setSurveyStatus("已完成", "已导出各年级文件和汇总.xlsx");
  } catch (error) {
    console.error(error);
    setSurveyStatus("处理失败", error.message || "请检查文件结构与扩展配置");
  } finally {
    surveyExportButton.disabled = false;
  }
});

convertModeInputs.forEach((input) => {
  input.addEventListener("change", () => {
    refreshConvertMode();
    refreshConvertReadyState();
  });
});

convertInput.addEventListener("change", (event) => {
  [convertFile] = event.target.files;
  convertFileName.textContent = convertFile ? convertFile.name : "未选择";
  refreshConvertReadyState();
});

convertClearButton.addEventListener("click", () => {
  convertInput.value = "";
  convertFile = null;
  convertFileName.textContent = "未选择";
  refreshConvertMode();
  refreshConvertReadyState();
});

convertButton.addEventListener("click", async () => {
  if (!convertFile) {
    setConvertStatus("待处理", "请先选择文件");
    return;
  }

  const mode = currentConvertMode();

  if (mode === "pdf-to-word") {
    setConvertStatus("暂不支持", "PDF 保格式转 Word 需要后端服务或 PDF 解析库");
    return;
  }

  if (!convertFile.name.toLowerCase().endsWith(".docx")) {
    setConvertStatus("格式不匹配", "Word 转 PDF 请上传 .docx 文件");
    return;
  }

  convertButton.disabled = true;
  setConvertStatus("处理中", "正在读取 Word 内容");

  try {
    const html = await docxFileToPrintableHtml(convertFile);
    openPrintableDocument(html, baseFileName(convertFile.name));
    setConvertStatus("已打开", "在新窗口中选择打印并保存为 PDF");
  } catch (error) {
    console.error(error);
    setConvertStatus("转换失败", "请确认文件是标准 .docx 文档");
  } finally {
    convertButton.disabled = false;
  }
});

function refreshConvertMode() {
  const mode = currentConvertMode();
  convertModeLabel.textContent = mode === "word-to-pdf" ? "Word 转 PDF" : "PDF 转 Word";
  convertModeCards.forEach((card) => {
    const input = card.querySelector("input");
    card.classList.toggle("active", input.checked);
  });
}

function refreshConvertReadyState() {
  const mode = currentConvertMode();
  const hasFile = Boolean(convertFile);
  convertButton.disabled = !hasFile;

  if (!hasFile) {
    setConvertStatus("待处理", "请选择转换模式和文件");
    return;
  }

  if (mode === "word-to-pdf" && !convertFile.name.toLowerCase().endsWith(".docx")) {
    setConvertStatus("格式不匹配", "Word 转 PDF 请上传 .docx 文件");
    return;
  }

  if (mode === "pdf-to-word") {
    setConvertStatus("需要后端", "纯静态网页暂不能可靠保格式转换 PDF 到 Word");
    return;
  }

  setConvertStatus("已就绪", "点击后打开打印页面");
}

function currentConvertMode() {
  return Array.from(convertModeInputs).find((input) => input.checked)?.value || "word-to-pdf";
}

function setConvertStatus(status, hint) {
  convertStatus.textContent = status;
  convertStatusHint.textContent = hint;
}

async function docxFileToPrintableHtml(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const documentXml = await zip.file("word/document.xml")?.async("string");

  if (!documentXml) {
    throw new Error("Missing document.xml");
  }

  const paragraphs = [];
  const paragraphMatches = documentXml.match(/<w:p[\s\S]*?<\/w:p>/g) || [];

  paragraphMatches.forEach((paragraphXml) => {
    const texts = [];
    const textMatches = paragraphXml.match(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/g) || [];

    textMatches.forEach((textXml) => {
      texts.push(decodeXmlText(textXml.replace(/<[^>]+>/g, "")));
    });

    if (texts.join("").trim()) {
      paragraphs.push(`<p>${escapeHtml(texts.join(""))}</p>`);
    }
  });

  return paragraphs.join("\n") || "<p>未读取到可转换的文字内容。</p>";
}

function decodeXmlText(value) {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function openPrintableDocument(contentHtml, title) {
  const printWindow = window.open("", "_blank");

  if (!printWindow) {
    throw new Error("Popup blocked");
  }

  printWindow.document.write(`<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <title>${escapeHtml(title)}</title>
  <style>
    @page { margin: 18mm; }
    body { color: #111827; font-family: "Microsoft YaHei", "SimSun", sans-serif; font-size: 14px; line-height: 1.8; }
    main { max-width: 760px; margin: 0 auto; }
    p { margin: 0 0 10px; white-space: pre-wrap; }
  </style>
</head>
<body>
  <main>${contentHtml}</main>
  <script>window.addEventListener("load", () => setTimeout(() => window.print(), 200));<\/script>
</body>
</html>`);
  printWindow.document.close();
}

async function readSurveyWorkbook(file, label) {
  if (!file) {
    return null;
  }

  try {
    return XLSX.read(await file.arrayBuffer(), { type: "array" });
  } catch (error) {
    console.error(error);
    setSurveyStatus("读取失败", `${label}不是有效的 Excel 文件`);
    return null;
  }
}

function refreshSurveyReadyState() {
  const isReady = Boolean(surveySourceWorkbook && surveyTeacherWorkbook);
  surveyExportButton.disabled = !isReady;
  setSurveyStatus(isReady ? "已就绪" : "待处理", isReady ? "点击处理并导出 ZIP" : "请选择问卷数据和教师任课表");
}

function setSurveyStatus(status, hint) {
  surveyStatus.textContent = status;
  surveyStatusHint.textContent = hint;
}

function readSurveyConfig() {
  const grades = JSON.parse(surveyGradeConfigInput.value);
  const subjects = JSON.parse(surveySubjectConfigInput.value);
  const edgeWeights = parseSurveyWeights(surveyEdgeWeightsInput.value);
  const middleWeights = parseSurveyWeights(surveyMiddleWeightsInput.value);

  if (!Array.isArray(grades) || grades.length === 0 || !Array.isArray(subjects) || subjects.length === 0) {
    throw new Error("年级配置和学科列配置必须是非空 JSON 数组");
  }

  subjects.forEach((subject) => {
    if (!subject.name || !["five", "single"].includes(subject.type) || !Number.isInteger(subject.startColumn)) {
      throw new Error("学科配置需要包含 name、type 和整数 startColumn");
    }
  });

  return { grades, subjects, edgeWeights, middleWeights };
}

function parseSurveyWeights(value) {
  const weights = String(value)
    .split(",")
    .map((item) => Number(item.trim()));

  if (weights.length !== 4 || weights.some((item) => !Number.isFinite(item))) {
    throw new Error("评分权重需要填写四个用英文逗号分隔的数字");
  }

  return weights;
}

function parseSurveyTeacherWorkbook(config) {
  const rows = XLSX.utils.sheet_to_json(surveyTeacherWorkbook.Sheets[surveyTeacherWorkbook.SheetNames[0]], {
    header: 1,
    defval: "",
  });
  const gradeByLabel = new Map(config.grades.map((grade) => [grade.teacherLabel, grade.id]));
  const teacherMap = new Map();
  const namesBySubject = new Map();
  const nameGradeMap = new Map();
  let gradeId = "";
  let headers = [];

  rows.forEach((row) => {
    const firstCell = String(row[0]).trim();

    if (gradeByLabel.has(firstCell)) {
      gradeId = gradeByLabel.get(firstCell);
      headers = row;
      teacherMap.set(gradeId, new Map());
      return;
    }

    const classNumber = firstCell.match(/^(\d+)/)?.[1];

    if (!gradeId || !classNumber) {
      return;
    }

    const classMap = new Map();
    teacherMap.get(gradeId).set(classNumber, classMap);

    headers.forEach((header, index) => {
      if (index === 0 || !header) {
        return;
      }

      const teacherName = String(row[index] || "").replace(/\d+/g, "").trim();
      const subjectName = String(header).trim();

      if (!teacherName || !subjectName) {
        return;
      }

      classMap.set(subjectName, teacherName);

      if (!namesBySubject.has(subjectName)) {
        namesBySubject.set(subjectName, new Map());
      }

      const subjectTeachers = namesBySubject.get(subjectName);
      if (!subjectTeachers.has(teacherName)) {
        subjectTeachers.set(teacherName, { teacherName, gradeIds: new Set() });
      }
      subjectTeachers.get(teacherName).gradeIds.add(gradeId);
      nameGradeMap.set(teacherName, gradeId);
    });
  });

  return { teacherMap, namesBySubject, nameGradeMap };
}

function parseSurveySourceWorkbook(config) {
  const rows = XLSX.utils.sheet_to_json(surveySourceWorkbook.Sheets[surveySourceWorkbook.SheetNames[0]], {
    header: 1,
    defval: "",
  });
  const subjectStats = new Map(config.subjects.map((subject) => [subject.name, new Map()]));
  const classes = new Map();
  const ethics = new Map();
  const opinions = new Map();

  rows.slice(1).forEach((row) => {
    const gradeId = String(row[7]).trim();
    const classNumber = row.slice(8, 14).map((value) => String(value).trim()).find(Boolean);

    if (!gradeId || !classNumber) {
      return;
    }

    ensureSurveySet(classes, gradeId).add(classNumber);
    const classKey = `${gradeId}|${classNumber}`;

    for (let index = 16; index <= 27; index += 1) {
      if (String(row[index]).trim() === "1" && String(row[index + 1] || "").trim()) {
        ensureSurveyArray(ethics, classKey).push(String(row[index + 1]).trim());
      }
    }

    const opinion = String(row[82] || "").trim();
    if (opinion && opinion !== "无") {
      ensureSurveyArray(opinions, classKey).push(opinion);
    }

    config.subjects.forEach((subject) => {
      const questionCount = subject.type === "five" ? 5 : 1;
      const classStats = ensureSurveyClassSubjectStats(subjectStats.get(subject.name), gradeId, classNumber, questionCount);

      for (let questionIndex = 0; questionIndex < questionCount; questionIndex += 1) {
        const answer = Number(row[subject.startColumn - 1 + questionIndex]);
        if (Number.isInteger(answer) && answer >= 1 && answer <= 4) {
          classStats[questionIndex][answer - 1] += 1;
        }
      }
    });
  });

  return { subjectStats, classes, ethics, opinions };
}

function ensureSurveyClassSubjectStats(subjectMap, gradeId, classNumber, questionCount) {
  if (!subjectMap.has(gradeId)) {
    subjectMap.set(gradeId, new Map());
  }

  const classMap = subjectMap.get(gradeId);
  if (!classMap.has(classNumber)) {
    classMap.set(classNumber, Array.from({ length: questionCount }, () => [0, 0, 0, 0]));
  }

  return classMap.get(classNumber);
}

function ensureSurveySet(map, key) {
  if (!map.has(key)) {
    map.set(key, new Set());
  }
  return map.get(key);
}

function ensureSurveyArray(map, key) {
  if (!map.has(key)) {
    map.set(key, []);
  }
  return map.get(key);
}

async function loadSurveyTemplateBuffer() {
  if (surveyTemplateBuffer) {
    return surveyTemplateBuffer;
  }

  try {
    const response = await fetch("survey-template.xlsx");
    if (response.ok) {
      return await response.arrayBuffer();
    }
  } catch (error) {
    console.info("Falling back to embedded survey template.", error);
  }

  if (!window.SURVEY_TEMPLATE_XLSX_BASE64) {
    throw new Error("缺少 survey-template.xlsx");
  }

  return base64ToArrayBuffer(window.SURVEY_TEMPLATE_XLSX_BASE64);
}

async function createSurveyGradeWorkbook(grade, classNumbers, stats, teacherData, summary, templateWorkbook, config) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "航哥的百宝箱";

  classNumbers.forEach((classNumber) => {
    const worksheet = workbook.addWorksheet(`${grade.displayName}${classNumber}班`.slice(0, 31));
    fillSurveyClassSheet(worksheet, templateWorkbook.worksheets[0], grade, classNumber, stats, teacherData, summary, config);
  });

  return workbook;
}

function fillSurveyClassSheet(worksheet, templateSheet, grade, classNumber, stats, teacherData, summary, config) {
  copySurveyTemplateRange(templateSheet, worksheet, "A1:P1", "A1");
  setSurveyFittedText(worksheet.getCell("A1"), `${grade.outputName}${classNumber}班中期问卷调查统计汇总`, 28);
  let blockIndex = 0;

  config.subjects.forEach((subject) => {
    const gradeStats = stats.subjectStats.get(subject.name)?.get(grade.id);
    const counts = gradeStats?.get(classNumber);

    if (!counts || counts.every((question) => question.every((count) => count === 0))) {
      return;
    }

    const startRow = 2 + blockIndex * 5;
    copySurveyTemplateRange(templateSheet, worksheet, subject.type === "five" ? "A2:P6" : "A7:D11", `A${startRow}`);
    const teacherName = findSurveyTeacherName(teacherData, grade.id, classNumber, subject);
    const rowName = teacherName ? `${subject.name}  (${teacherName})` : subject.name;
    setSurveyFittedText(worksheet.getCell(startRow, 1), rowName, 12);

    counts.forEach((questionCounts, questionIndex) => {
      const score = writeSurveyQuestionScores(worksheet, startRow, questionIndex, questionCounts, subject.type, config);

      if (questionIndex === counts.length - 1 && Number.isFinite(score)) {
        ensureSurveyArray(summary.teacherScores, rowName).push(score);
        summary.scoreRecords.push({ gradeId: grade.id, subjectName: subject.name, teacherName, score });
      }
    });

    blockIndex += 1;
  });

  const descriptionStartRow = 3 + blockIndex * 5;
  copySurveyTemplateRange(templateSheet, worksheet, "A13:K33", `A${descriptionStartRow}`);
  const classKey = `${grade.id}|${classNumber}`;

  (stats.ethics.get(classKey) || []).forEach((text, index) => {
    setSurveyFittedText(worksheet.getCell(descriptionStartRow + 1 + index, 2), text, 28, {
      wrapText: true,
      vertical: "top",
    });
    worksheet.getRow(descriptionStartRow + 1 + index).height = 30;
  });

  (stats.opinions.get(classKey) || []).forEach((text, index) => {
    setSurveyFittedText(worksheet.getCell(descriptionStartRow + 1 + index, 7), text, 28, {
      wrapText: true,
      vertical: "top",
    });
    worksheet.getRow(descriptionStartRow + 1 + index).height = 30;
  });
}

function writeSurveyQuestionScores(worksheet, startRow, questionIndex, counts, subjectType, config) {
  const total = counts.reduce((sum, count) => sum + count, 0);
  if (total === 0) {
    return NaN;
  }

  const weights = subjectType === "five" && questionIndex > 0 && questionIndex < 4 ? config.middleWeights : config.edgeWeights;
  let score = 0;
  const percentColumn = subjectType === "five" ? 3 + questionIndex * 3 : 3;
  const scoreColumn = subjectType === "five" ? 4 + questionIndex * 3 : 4;

  counts.forEach((count, answerIndex) => {
    const ratio = count / total;
    worksheet.getCell(startRow + 1 + answerIndex, percentColumn).value = `${(ratio * 100).toFixed(ratio === 1 ? 0 : 2)}%`;
    score += ratio * weights[answerIndex];
  });

  worksheet.getCell(startRow + 1, scoreColumn).value = score.toFixed(2);
  return score;
}

function findSurveyTeacherName(teacherData, gradeId, classNumber, subject) {
  const classMap = teacherData.teacherMap.get(String(gradeId))?.get(String(classNumber));
  return classMap?.get(subject.teacherSubject || subject.name) || classMap?.get(subject.name) || "";
}

function copySurveyTemplateRange(sourceSheet, destinationSheet, sourceAddress, destinationAddress) {
  const source = XLSX.utils.decode_range(sourceAddress);
  const destination = XLSX.utils.decode_cell(destinationAddress);
  const rowOffset = destination.r - source.s.r;
  const columnOffset = destination.c - source.s.c;

  for (let row = source.s.r; row <= source.e.r; row += 1) {
    const sourceRow = sourceSheet.getRow(row + 1);
    const destinationRow = destinationSheet.getRow(row + 1 + rowOffset);
    destinationRow.height = sourceRow.height;

    for (let column = source.s.c; column <= source.e.c; column += 1) {
      const sourceCell = sourceSheet.getCell(row + 1, column + 1);
      const destinationCell = destinationSheet.getCell(row + 1 + rowOffset, column + 1 + columnOffset);
      destinationCell.value = cloneSurveyValue(sourceCell.value);
      destinationCell.style = cloneSurveyValue(sourceCell.style);
      updateSurveyTextFit(destinationCell, 18);
    }
  }

  for (let column = source.s.c; column <= source.e.c; column += 1) {
    destinationSheet.getColumn(column + 1 + columnOffset).width = sourceSheet.getColumn(column + 1).width;
  }

  (sourceSheet.model.merges || []).forEach((mergeAddress) => {
    const merge = XLSX.utils.decode_range(mergeAddress);
    if (merge.s.r < source.s.r || merge.e.r > source.e.r || merge.s.c < source.s.c || merge.e.c > source.e.c) {
      return;
    }

    const translated = XLSX.utils.encode_range({
      s: { r: merge.s.r + rowOffset, c: merge.s.c + columnOffset },
      e: { r: merge.e.r + rowOffset, c: merge.e.c + columnOffset },
    });

    try {
      destinationSheet.mergeCells(translated);
    } catch (error) {
      console.info("Skipping duplicate merge", translated);
    }
  });
}

function cloneSurveyValue(value) {
  if (value === undefined || value === null) {
    return value;
  }
  return JSON.parse(JSON.stringify(value));
}

function setSurveyFittedText(cell, value, shrinkThreshold, alignment = {}) {
  cell.value = value;
  updateSurveyTextFit(cell, shrinkThreshold, alignment);
}

function updateSurveyTextFit(cell, shrinkThreshold, alignment = {}) {
  const value = cell.value;

  if (typeof value !== "string") {
    return;
  }

  const shouldShrink = value.trim().length > shrinkThreshold;
  cell.alignment = {
    ...cell.alignment,
    ...alignment,
    shrinkToFit: shouldShrink,
  };
}

function createSurveySummaryState() {
  return {
    teacherScores: new Map(),
    scoreRecords: [],
  };
}

function createSurveySummaryWorkbook(summary, teacherData, config) {
  const workbook = new ExcelJS.Workbook();
  const teacherSheet = workbook.addWorksheet("教师");
  const headTeacherSheet = workbook.addWorksheet("班主任");
  const subjectSheet = workbook.addWorksheet("学科");
  const averageSheet = workbook.addWorksheet("平均分");
  appendSurveyRow(teacherSheet, ["教师", "平均分"]);
  appendSurveyRow(headTeacherSheet, ["班主任", "平均分"]);

  Array.from(summary.teacherScores.entries()).forEach(([name, scores]) => {
    const row = [name, surveyAverage(scores)];
    appendSurveyRow(name.includes("班主任") ? headTeacherSheet : teacherSheet, row);
  });

  appendSurveyRow(subjectSheet, ["学科", "教师", "平均分", "年级"]);
  teacherData.namesBySubject.forEach((teachers, subjectName) => {
    if (["通用", "国学", "阅读"].includes(subjectName)) {
      return;
    }
    teachers.forEach(({ teacherName, gradeIds }) => {
      const matched = summary.scoreRecords.filter((record) => record.teacherName === teacherName && record.subjectName === normalizeSurveySubjectName(subjectName));
      appendSurveyRow(subjectSheet, [subjectName, teacherName, matched.length ? surveyAverage(matched.map((record) => record.score)) : "", surveyGradeDisplayNames(gradeIds, config)]);
    });
  });

  const averageSubjects = ["班主任", "语文", "数学", "英语", "政治", "历史", "地理", "物理", "化学", "生物", "体育"];
  appendSurveyRow(averageSheet, ["年级", ...averageSubjects]);
  config.grades.forEach((grade) => {
    appendSurveyRow(averageSheet, [
      grade.displayName,
      ...averageSubjects.map((subjectName) => {
        const records = summary.scoreRecords.filter((record) => record.gradeId === grade.id && record.subjectName === subjectName);
        return records.length ? surveyAverage(records.map((record) => record.score)) : "";
      }),
    ]);
  });

  [teacherSheet, headTeacherSheet, subjectSheet, averageSheet].forEach((sheet) => {
    sheet.getRow(1).font = { bold: true };
    (sheet.columns || []).forEach((column) => {
      column.width = 18;
    });
  });

  return workbook;
}

function appendSurveyRow(worksheet, values) {
  const rowNumber = worksheet.rowCount + 1;
  values.forEach((value, index) => {
    worksheet.getCell(rowNumber, index + 1).value = value;
  });
}

function surveyAverage(values) {
  return (values.reduce((sum, value) => sum + Number(value), 0) / values.length).toFixed(2);
}

function surveyGradeDisplayName(gradeId, config) {
  return config.grades.find((grade) => grade.id === gradeId)?.displayName || "";
}

function surveyGradeDisplayNames(gradeIds, config) {
  return config.grades
    .filter((grade) => gradeIds.has(grade.id))
    .map((grade) => grade.displayName)
    .join("、");
}

function normalizeSurveySubjectName(subjectName) {
  if (subjectName === "外语") {
    return "英语";
  }
  if (subjectName === "心理、生涯") {
    return "心理";
  }
  return subjectName;
}

function compareSurveyClassNames(left, right) {
  return Number(left) - Number(right) || String(left).localeCompare(String(right), "zh-CN");
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
