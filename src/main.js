const state = {
  files: [],
  activeId: null,
  activeTokenEditor: null,
  collapsedSections: {},
};

const SAVE_DIRECTORY_DB_NAME = "json-reader-save-directory";
const SAVE_DIRECTORY_STORE_NAME = "handles";
const SAVE_DIRECTORY_KEY = "lastSaveDirectory";

const WORD_TYPE_OPTIONS = [
  [0, "英语"],
  [1, "日语"],
  [2, "法语"],
  [3, "韩语"],
  [4, "西班牙语"],
  [5, "俄语"],
];

const EXAM_ID_OPTIONS = [
  [1, "中考"],
  [2, "高考"],
  [3, "四级"],
  [4, "六级"],
  [5, "考研"],
  [6, "英语专四"],
  [7, "英语专八"],
  [8, "专升本"],
  [9, "雅思"],
  [10, "新托福"],
  [11, "JLPT"],
  [12, "KET"],
  [13, "小学区考"],
];

const ARTICLE_TYPE_OPTIONS = [
  [1, "真题"],
  [2, "模拟题"],
  [3, "小说文章"],
];

const CONTENT_TYPE_OPTIONS = [
  "clozeQuestion",
  "clozeOption",
  "mcQuestion",
  "mcOption",
  "fibQuestion",
  "tofQuestion",
  "matchingQuestion",
  "matchingOption",
  "roQuestion",
  "instruction",
  "ecQuestion",
  "transQuestion",
];

const fileListEl = document.getElementById("fileList");
const filePickerEl = document.getElementById("filePicker");
const folderPickerEl = document.getElementById("folderPicker");
const addFilesBtnEl = document.getElementById("addFilesBtn");
const addFolderBtnEl = document.getElementById("addFolderBtn");
const previewEl = document.getElementById("preview");
const previewMetaEl = document.getElementById("previewMeta");
const editorEl = document.getElementById("editor");
const highlightEl = document.getElementById("highlight");
const editorStatusEl = document.getElementById("editorStatus");
const spacifyBtnEl = document.getElementById("spacifyBtn");
const exportTxtBtnEl = document.getElementById("exportTxtBtn");
const saveAsBtnEl = document.getElementById("saveAsBtn");
const editMetaBtnEl = document.getElementById("editMetaBtn");
const metaModalEl = document.getElementById("metaModal");
const metaFormEl = document.getElementById("metaForm");
const metaCancelBtnEl = document.getElementById("metaCancelBtn");
const toastContainerEl = document.getElementById("toastContainer");
const metaWordTypeEl = metaFormEl?.elements?.namedItem("wordType");
const metaExamIdEl = metaFormEl?.elements?.namedItem("examId");
const metaArticleTypeEl = metaFormEl?.elements?.namedItem("articleType");

addFilesBtnEl?.addEventListener("click", () => filePickerEl?.click());
addFolderBtnEl?.addEventListener("click", () => folderPickerEl?.click());
spacifyBtnEl?.addEventListener("click", onSpacify);
exportTxtBtnEl?.addEventListener("click", onExportTxt);
saveAsBtnEl?.addEventListener("click", onSaveAs);
editMetaBtnEl?.addEventListener("click", onEditMeta);
metaCancelBtnEl?.addEventListener("click", closeMetaModal);
metaFormEl?.addEventListener("submit", onSaveMeta);
metaModalEl?.addEventListener("click", (event) => {
  if (event.target === metaModalEl) {
    closeMetaModal();
  }
});
filePickerEl.addEventListener("change", onPickFiles);
folderPickerEl.addEventListener("change", onPickFiles);
editorEl.addEventListener("input", onEditorInput);
editorEl.addEventListener("keydown", onEditorKeydown);
editorEl.addEventListener("scroll", syncEditorScroll);

populateMetaSelect(metaWordTypeEl, WORD_TYPE_OPTIONS);
populateMetaSelect(metaExamIdEl, EXAM_ID_OPTIONS);
populateMetaSelect(metaArticleTypeEl, ARTICLE_TYPE_OPTIONS);
setEmptyState();

async function onPickFiles(event) {
  const rawFiles = Array.from(event.target.files || []).filter((file) =>
    file.name.toLowerCase().endsWith(".json")
  );

  if (!rawFiles.length) {
    return;
  }

  for (const file of rawFiles) {
    const text = await file.text();
    const existing = state.files.find((item) => item.path === getFilePath(file));
    const entry = buildFileEntry(file, text);
    if (existing) {
      Object.assign(existing, entry);
    } else {
      state.files.push(entry);
    }
  }

  if (!state.activeId && state.files[0]) {
    state.activeId = state.files[0].id;
  }

  renderFileList();
  syncActiveIntoEditor();
  event.target.value = "";
}

function buildFileEntry(file, text) {
  const parsed = safeParseJson(text);
  const path = getFilePath(file);
  return {
    id: path,
    name: file.name,
    path,
    raw: prettyJson(text),
    parsed,
  };
}

function getFilePath(file) {
  return file.webkitRelativePath || file.name;
}

function renderFileList() {
  fileListEl.innerHTML = "";
  const sorted = [...state.files].sort((a, b) => a.path.localeCompare(b.path));

  for (const file of sorted) {
    const li = document.createElement("li");
    const button = document.createElement("button");
    button.type = "button";
    button.className = file.id === state.activeId ? "active" : "";
    button.addEventListener("click", () => {
      state.activeId = file.id;
      syncActiveIntoEditor();
      renderFileList();
    });

    const name = document.createElement("div");
    name.className = "file-name";
    name.textContent = file.name;

    const path = document.createElement("div");
    path.className = "file-path";
    path.textContent = file.path;

    button.append(name, path);
    li.appendChild(button);
    fileListEl.appendChild(li);
  }
}

function syncActiveIntoEditor() {
  const file = getActiveFile();
  if (!file) {
    setEmptyState();
    return;
  }

  state.activeTokenEditor = null;

  editorEl.value = file.raw;
  renderHighlight(editorEl.value);

  if (file.parsed.ok) {
    setStatusOk();
    renderPreviewSafe(file.parsed.value, file.name);
  } else {
    setStatusError(file.parsed.error);
    previewEl.innerHTML = `<p class="placeholder">JSON 解析错误：${escapeHtml(file.parsed.error)}</p>`;
    previewMetaEl.textContent = file.name;
  }
}

function onEditorInput() {
  const file = getActiveFile();
  if (!file) {
    return;
  }

  state.activeTokenEditor = null;

  file.raw = editorEl.value;
  renderHighlight(file.raw);
  file.parsed = safeParseJson(file.raw);

  if (file.parsed.ok) {
    setStatusOk();
    renderPreviewSafe(file.parsed.value, file.name);
  } else {
    setStatusError(file.parsed.error);
    previewEl.innerHTML = `<p class="placeholder">JSON 解析错误：${escapeHtml(file.parsed.error)}</p>`;
    previewMetaEl.textContent = file.name;
  }
}

function renderPreviewSafe(json, filename) {
  try {
    renderPreview(json, filename);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    previewEl.innerHTML = `<p class="placeholder">渲染失败：${escapeHtml(message)}</p>`;
    previewMetaEl.textContent = filename;
    setStatusError(`渲染失败：${message}`);
  }
}

function renderPreview(json, filename) {
  previewEl.innerHTML = "";

  const sections = extractSections(json);
  const totalRows = sections.reduce((sum, section) => sum + section.rows.length, 0);
  previewMetaEl.textContent = `${filename} · ${totalRows} 行`;

  renderMetaSummary(json);

  if (!sections.length) {
    previewEl.innerHTML += `<p class="placeholder">未找到可渲染内容。期望结构如 { "data": [{ "sectionType": "", "section": [[...]] }] }、{ "data": [[[...]], [[...]]] } 或 { "data": [[...], [...]] }。</p>`;
    return;
  }

  const contentTypeCounters = new Map();
  sections.forEach((section, sectionIndex) => {
    const questionCounters = new Map();
    let clozeOptionGroupIndex = 0;
    const collapsed = isSectionCollapsed(section);

    previewEl.appendChild(buildSectionTypeControl(section, sectionIndex));

    if (collapsed) {
      return;
    }

    section.rows.forEach((row, rowIndex) => {
      const isInstructionOnlyRow = row.every((segment) => String(segment?.contentType || "") === "instruction");
      if (isInstructionOnlyRow) {
        const spacerEl = document.createElement("div");
        spacerEl.className = "instruction-spacer";
        previewEl.appendChild(spacerEl);
      }

      const wrapper = document.createElement("section");
      wrapper.className = "row";

      const choiceOptionType = getChoiceOptionType(row);
      const clozeQuestionNumber = choiceOptionType === "clozeOption" ? ++clozeOptionGroupIndex : null;
      let optionIndex = 0;

      row.forEach((segment, segmentIndex) => {
        const contentType = String(segment?.contentType || "");
        const lower = contentType.toLowerCase();
        const contentTypeOccurrence = (contentTypeCounters.get(contentType) || 0) + 1;
        contentTypeCounters.set(contentType, contentTypeOccurrence);

        let questionNumber = null;
        let optionLabel = "";

        if (lower.includes("question") && contentType !== "clozeQuestion") {
          const next = (questionCounters.get(contentType) || 0) + 1;
          questionCounters.set(contentType, next);
          questionNumber = next;
        }

        if (choiceOptionType && contentType === choiceOptionType) {
          optionIndex += 1;
          optionLabel = `${toAlphabetLabel(optionIndex)}. `;
          if (contentType === "clozeOption") {
            questionNumber = optionIndex === 1 ? clozeQuestionNumber : null;
          } else {
            questionNumber = null;
          }
        }

        const segmentEl = renderSegment(segment, {
          questionNumber,
          optionLabel,
          contentTypeOccurrence,
          sectionIndex,
          rowIndex,
          segmentIndex,
        });
        wrapper.appendChild(segmentEl);
      });

      previewEl.appendChild(wrapper);

      if (section.hasSectionWrapper && rowIndex < section.rows.length - 1) {
        previewEl.appendChild(buildSectionSplitDivider(sectionIndex, rowIndex));
      }
    });
  });
}

function buildSectionTypeControl(section, sectionIndex) {
  const wrapper = document.createElement("div");
  wrapper.className = "section-header";

  if (section.hasSectionWrapper) {
    const toggleBtn = document.createElement("button");
    toggleBtn.type = "button";
    toggleBtn.className = "section-fold-btn";
    toggleBtn.textContent = isSectionCollapsed(section) ? "展开" : "折叠";
    toggleBtn.addEventListener("click", (event) => {
      event.stopPropagation();
      toggleSectionCollapsed(section);
    });
    wrapper.appendChild(toggleBtn);
  }

  const field = document.createElement("div");
  field.className = "section-type-field";

  const input = document.createElement("input");
  input.type = "text";
  input.className = "section-type-input";
  input.value = getDisplaySectionType(section);
  input.placeholder = "输入 sectionType";
  input.title = "编辑 sectionType";
  input.addEventListener("click", (event) => event.stopPropagation());
  input.addEventListener("mousedown", (event) => event.stopPropagation());
  input.addEventListener("change", () => {
    commitSectionTypeInput(sectionIndex, input);
  });
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      commitSectionTypeInput(sectionIndex, input);
      input.blur();
    }
  });

  field.appendChild(input);

  if (!section.hasExplicitSectionType) {
    const badge = document.createElement("span");
    badge.className = "section-type-badge";
    badge.textContent = "无sectionType";
    field.appendChild(badge);
  }

  wrapper.appendChild(field);

  if (section.hasSectionWrapper) {
    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "section-delete-btn";
    deleteBtn.textContent = "删除分节";
    deleteBtn.title = "删除当前 section，并与相邻 section 合并";
    deleteBtn.addEventListener("click", (event) => {
      event.stopPropagation();
      deleteSection(sectionIndex);
    });
    wrapper.appendChild(deleteBtn);
  }

  return wrapper;
}

function commitSectionTypeInput(sectionIndex, inputEl) {
  if (!inputEl) {
    return;
  }
  updateSectionType(sectionIndex, inputEl.value.trim());
}

function buildSectionSplitDivider(sectionIndex, rowIndex) {
  const divider = document.createElement("div");
  divider.className = "section-split-divider";

  const line = document.createElement("div");
  line.className = "section-split-line";

  const addBtn = document.createElement("button");
  addBtn.type = "button";
  addBtn.className = "section-split-btn";
  addBtn.textContent = "新增分节";
  addBtn.title = "在这里创建新的 section";
  addBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    splitSectionAt(sectionIndex, rowIndex);
  });

  divider.appendChild(line);
  divider.appendChild(addBtn);
  return divider;
}

function getSectionCollapseKey(section) {
  const file = getActiveFile();
  const fileId = file?.id || "unknown";
  return `${fileId}::${section.dataIndex}`;
}

function isSectionCollapsed(section) {
  if (!section?.hasSectionWrapper) {
    return false;
  }
  return Boolean(state.collapsedSections[getSectionCollapseKey(section)]);
}

function toggleSectionCollapsed(section) {
  if (!section?.hasSectionWrapper) {
    return;
  }

  const key = getSectionCollapseKey(section);
  state.collapsedSections[key] = !state.collapsedSections[key];

  const file = getActiveFile();
  if (file?.parsed.ok) {
    renderPreviewSafe(file.parsed.value, file.name);
  }
}

function renderSegment(segment, markers = {}) {
  const contentType = String(segment?.contentType || "");
  const lower = contentType.toLowerCase();

  const el = document.createElement("div");
  el.className = "segment";

  const contentWrap = document.createElement("div");
  contentWrap.className = "segment-main";

  if (lower.includes("question") && contentType !== "clozeQuestion") {
    el.classList.add("question");
  }

  if (lower.includes("instruction") || contentType === "instruction") {
    el.classList.add("instruction");
  }

  if (lower.includes("option")) {
    el.classList.add("option");
  }

  if (lower.includes("option") && isCorrectOption(segment?.correct)) {
    el.classList.add("option-correct");
  }

  if (contentType === "pic") {
    const src = segment?.content;
    if (typeof src === "string" && src.trim()) {
      const img = document.createElement("img");
      img.src = src;
      img.alt = "Exam illustration";
      img.style.maxWidth = "100%";
      img.style.borderRadius = "8px";
      img.style.border = "1px solid #d8dee7";
      contentWrap.appendChild(img);
      el.appendChild(contentWrap);
      el.appendChild(buildContentTypeControl(contentType, markers));
      return el;
    }
  }

  if (contentType === "instruction") {
    contentWrap.textContent = `${markers.questionNumber ? `${markers.questionNumber}. ` : ""}${stringifyInstruction(segment)}`;
    el.appendChild(contentWrap);
    el.appendChild(buildContentTypeControl(contentType, markers));
    el.classList.add("clickable");
    el.addEventListener("click", () => jumpToSegmentInEditor(contentType, markers.contentTypeOccurrence || 1));
    return el;
  }

  if (isKnownContentType(contentType)) {
    const text = extractReadableText(segment);
    const hasLeadingQuestionNo = startsWithQuestionNumber(text);
    const canShowAutoQuestionNo = !(lower.includes("question") && hasLeadingQuestionNo);

    if (contentType === "clozeOption") {
      if (markers.questionNumber && canShowAutoQuestionNo) {
        const markerEl = document.createElement("div");
        markerEl.className = "cloze-question-number";
        markerEl.textContent = `${markers.questionNumber}.`;
        contentWrap.appendChild(markerEl);
      }

      const textEl = document.createElement("span");
      textEl.className = "segment-text";
      appendSegmentContent(textEl, segment, markers.optionLabel || "", markers);
      contentWrap.appendChild(textEl);
    } else {
      const prefix = canShowAutoQuestionNo && markers.questionNumber
        ? `${markers.questionNumber}. ${markers.optionLabel || ""}`
        : (markers.optionLabel || "");
      const textEl = document.createElement("span");
      textEl.className = "segment-text";
      appendSegmentContent(textEl, segment, prefix, markers);
      contentWrap.appendChild(textEl);
    }

    el.appendChild(contentWrap);
    el.appendChild(buildContentTypeControl(contentType, markers));

    if (contentType !== "pic") {
      el.classList.add("clickable");
      el.addEventListener("click", () => jumpToSegmentInEditor(contentType, markers.contentTypeOccurrence || 1));
    }
    return el;
  }

  el.classList.add("unknown");
  contentWrap.textContent = JSON.stringify(segment, null, 2);
  el.appendChild(contentWrap);
  el.appendChild(buildContentTypeControl(contentType, markers));
  return el;
}

function buildContentTypeControl(contentType, markers) {
  const select = document.createElement("select");
  select.className = "content-type-select";

  const options = [...CONTENT_TYPE_OPTIONS];
  if (contentType && !options.includes(contentType)) {
    options.unshift(contentType);
  }

  options.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });

  select.value = contentType;
  select.title = "修改 contentType";
  select.addEventListener("mousedown", (event) => event.stopPropagation());
  select.addEventListener("click", (event) => event.stopPropagation());
  select.addEventListener("change", (event) => {
    event.stopPropagation();
    updateSegmentContentType(markers.sectionIndex, markers.rowIndex, markers.segmentIndex, select.value);
  });
  return select;
}

function updateSectionType(sectionIndex, nextType) {
  const file = getActiveFile();
  if (!file || !file.parsed.ok) {
    return;
  }

  const root = file.parsed.value;
  const sections = extractSections(root);
  const target = sections[sectionIndex];
  if (!target) {
    setStatusError("无法定位到要修改的 section。");
    return;
  }

  upgradeRootToSectionObjects(root, sections);
  const targetSection = root?.data?.[sectionIndex];
  if (!targetSection || typeof targetSection !== "object" || Array.isArray(targetSection)) {
    setStatusError("无法升级 section 结构。");
    return;
  }

  targetSection.sectionType = nextType;
  file.raw = JSON.stringify(root, null, 2);
  file.parsed = { ok: true, value: root };

  editorEl.value = file.raw;
  renderHighlight(file.raw);
  renderPreviewSafe(root, file.name);
  setStatusInfo(nextType ? `sectionType 已更新为 ${nextType}` : "sectionType 已清空");
}

function splitSectionAt(sectionIndex, rowIndex) {
  const file = getActiveFile();
  if (!file || !file.parsed.ok) {
    return;
  }

  const root = file.parsed.value;
  const sections = extractSections(root);
  const target = sections[sectionIndex];
  if (!target || !Array.isArray(target.rows)) {
    setStatusError("无法定位到要拆分的 section。");
    return;
  }

  if (rowIndex < 0 || rowIndex >= target.rows.length - 1) {
    setStatusError("当前分隔位置无法新增 section。");
    return;
  }

  upgradeRootToSectionObjects(root, sections);
  const targetSection = root?.data?.[sectionIndex];
  if (!targetSection || !Array.isArray(targetSection.section)) {
    setStatusError("section 结构升级失败。");
    return;
  }

  const movedRows = targetSection.section.splice(rowIndex + 1);
  root.data.splice(sectionIndex + 1, 0, {
    sectionType: "",
    section: movedRows,
  });

  file.raw = JSON.stringify(root, null, 2);
  file.parsed = { ok: true, value: root };
  editorEl.value = file.raw;
  renderHighlight(file.raw);
  renderPreviewSafe(root, file.name);
  setStatusInfo(`已在第 ${rowIndex + 1} 行后新增 section。`);
}

function deleteSection(sectionIndex) {
  const file = getActiveFile();
  if (!file || !file.parsed.ok) {
    return;
  }

  const root = file.parsed.value;
  const sections = extractSections(root);
  if (!sections[sectionIndex] || sections.length <= 1) {
    setStatusError("当前无法删除该分节。");
    return;
  }

  upgradeRootToSectionObjects(root, sections);
  const current = root?.data?.[sectionIndex];
  if (!current || !Array.isArray(current.section)) {
    setStatusError("无法定位到要删除的 section。");
    return;
  }

  const mergeIntoIndex = sectionIndex > 0 ? sectionIndex - 1 : 1;
  const mergeTarget = root?.data?.[mergeIntoIndex];
  if (!mergeTarget || !Array.isArray(mergeTarget.section)) {
    setStatusError("无法定位到相邻 section。");
    return;
  }

  if (sectionIndex > 0) {
    mergeTarget.section.push(...current.section);
  } else {
    mergeTarget.section.unshift(...current.section);
  }
  root.data.splice(sectionIndex, 1);

  file.raw = JSON.stringify(root, null, 2);
  file.parsed = { ok: true, value: root };
  editorEl.value = file.raw;
  renderHighlight(file.raw);
  renderPreviewSafe(root, file.name);
  setStatusInfo("已删除当前分节并合并到相邻 section。");
}

function updateSegmentContentType(sectionIndex, rowIndex, segmentIndex, nextType) {
  const file = getActiveFile();
  if (!file || !file.parsed.ok) {
    return;
  }

  const root = file.parsed.value;
  const segment = getSegmentByLocation(root, sectionIndex, rowIndex, segmentIndex);
  if (!segment || typeof segment !== "object") {
    setStatusError("无法定位到要修改的内容。");
    return;
  }

  segment.contentType = nextType;
  file.raw = JSON.stringify(root, null, 2);
  file.parsed = { ok: true, value: root };

  editorEl.value = file.raw;
  renderHighlight(file.raw);
  renderPreviewSafe(root, file.name);
  setStatusInfo(`contentType 已更新为 ${nextType}`);
}

function extractSections(json) {
  if (!json || typeof json !== "object") {
    return [];
  }

  const data = json.data;
  if (!Array.isArray(data)) {
    return [];
  }

  if (data.every(isSectionObject)) {
    return data.map((item, index) => createNormalizedSection(item.section, index, item, true));
  }

  if (data.every(isSectionArray)) {
    return data.map((rows, index) => createNormalizedSection(rows, index, rows, false));
  }

  if (data.every(isRowArray)) {
    return [createNormalizedSection(data, 0, data, false)];
  }

  const normalized = [];
  data.forEach((item, index) => {
    if (isSectionObject(item)) {
      normalized.push(createNormalizedSection(item.section, index, item, true));
      return;
    }
    if (isSectionArray(item)) {
      normalized.push(createNormalizedSection(item, index, item, false));
      return;
    }
    if (isRowArray(item)) {
      normalized.push(createNormalizedSection([item], index, item, false));
    }
  });
  return normalized;
}

function isSectionArray(value) {
  return Array.isArray(value) && value.every(isRowArray);
}

function isSectionObject(value) {
  return !!value && typeof value === "object" && !Array.isArray(value) && isSectionArray(value.section);
}

function isRowArray(value) {
  return Array.isArray(value) && value.every((item) => item && typeof item === "object" && !Array.isArray(item));
}

function createNormalizedSection(rows, dataIndex, source, hasSectionWrapper) {
  const explicitSectionType = typeof source?.sectionType === "string" ? source.sectionType.trim() : "";
  const inferredSectionType = getInferredSectionType(rows);
  return {
    rows,
    dataIndex,
    source,
    hasSectionWrapper,
    hasExplicitSectionType: Boolean(explicitSectionType),
    sectionType: explicitSectionType,
    inferredSectionType,
  };
}

function extractReadableText(segment) {
  if (!segment || typeof segment !== "object") {
    return String(segment);
  }

  if (Array.isArray(segment.content)) {
    return joinTokens(segment.content);
  }

  if (typeof segment.content === "string") {
    return segment.content;
  }

  if (Array.isArray(segment.instruction)) {
    return joinTokens(segment.instruction);
  }

  if (typeof segment.instruction === "string") {
    return segment.instruction;
  }

  return JSON.stringify(segment);
}

function stringifyInstruction(segment) {
  if (typeof segment?.instruction === "string") {
    return segment.instruction;
  }
  if (Array.isArray(segment?.instruction)) {
    return joinTokens(segment.instruction);
  }
  return extractReadableText(segment);
}

function joinTokens(tokens) {
  return tokens
    .map((token) => {
      if (typeof token === "string") {
        return token;
      }
      if (token && typeof token === "object") {
        return String(token.w ?? "");
      }
      return String(token ?? "");
    })
    .join("");
}

function isKnownContentType(contentType) {
  const lower = String(contentType || "").toLowerCase();
  return (
    lower === "body" ||
    lower === "pic" ||
    lower === "instruction" ||
    lower.includes("question") ||
    lower.includes("option")
  );
}

function isChoiceOptionType(contentType) {
  return new Set(["mcOption", "mapOption", "matchingOption", "clozeOption"]).has(contentType);
}

function toAlphabetLabel(index) {
  let n = index;
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function isCorrectOption(value) {
  if (value === true || value === 1) return true;
  if (typeof value === "string") return value.toLowerCase() === "true";
  return false;
}

function getSectionTitleForRow(row) {
  const mapping = {
    clozeQuestion: "完形填空",
    clozeOption: "完形填空",
    mcQuestion: "选择题",
    mcOption: "选择题",
    fibQuestion: "填空题",
    tofQuestion: "判断题",
    matchingQuestion: "匹配题",
    matchingOption: "匹配题",
    mapQuestion: "地图题",
    mapOption: "地图题",
    roQuestion: "排序题",
    ecQuestion: "改错题",
    transQuestion: "翻译题",
  };

  for (const segment of row) {
    const contentType = String(segment?.contentType || "");
    if (mapping[contentType]) {
      return mapping[contentType];
    }
  }

  return "";
}

function appendSegmentContent(container, segment, prefix = "", markers = {}) {
  if (prefix) {
    container.appendChild(document.createTextNode(prefix));
  }

  if (!segment || typeof segment !== "object") {
    container.appendChild(document.createTextNode(String(segment ?? "")));
    return;
  }

  if (Array.isArray(segment.content)) {
    container.appendChild(renderTokenFragment(segment.content, markers, "content"));
    return;
  }

  if (typeof segment.content === "string") {
    container.appendChild(document.createTextNode(segment.content));
    return;
  }

  if (Array.isArray(segment.instruction)) {
    container.appendChild(renderTokenFragment(segment.instruction, markers, "instruction"));
    return;
  }

  if (typeof segment.instruction === "string") {
    container.appendChild(document.createTextNode(segment.instruction));
    return;
  }

  container.appendChild(document.createTextNode(JSON.stringify(segment)));
}

function renderTokenFragment(tokens, markers = {}, tokenField = "content") {
  const fragment = document.createDocumentFragment();

  tokens.forEach((token, tokenIndex) => {
    if (typeof token === "string") {
      fragment.appendChild(document.createTextNode(token));
      return;
    }

    if (token && typeof token === "object") {
      const wrapper = document.createElement("span");
      wrapper.className = "token-wrapper";

      const span = document.createElement("span");
      span.className = "token-chip";
      span.textContent = String(token.w ?? "");
      if (token.i === 0) {
        span.classList.add("token-index-zero");
      }
      span.addEventListener("click", (event) => {
        event.stopPropagation();
        openTokenEditor(markers, tokenField, tokenIndex);
      });
      wrapper.appendChild(span);

      if (isActiveTokenEditor(markers, tokenField, tokenIndex)) {
        wrapper.appendChild(buildTokenEditor(token, markers, tokenField, tokenIndex));
      }

      fragment.appendChild(wrapper);
      return;
    }

    fragment.appendChild(document.createTextNode(String(token ?? "")));
  });

  return fragment;
}

function openTokenEditor(markers, tokenField, tokenIndex) {
  state.activeTokenEditor = {
    sectionIndex: markers.sectionIndex,
    rowIndex: markers.rowIndex,
    segmentIndex: markers.segmentIndex,
    tokenField,
    tokenIndex,
  };

  const file = getActiveFile();
  if (file?.parsed.ok) {
    renderPreviewSafe(file.parsed.value, file.name);
  }
}

function closeTokenEditor() {
  state.activeTokenEditor = null;

  const file = getActiveFile();
  if (file?.parsed.ok) {
    renderPreviewSafe(file.parsed.value, file.name);
  }
}

function isActiveTokenEditor(markers, tokenField, tokenIndex) {
  const active = state.activeTokenEditor;
  if (!active) {
    return false;
  }

  return (
    active.sectionIndex === markers.sectionIndex &&
    active.rowIndex === markers.rowIndex &&
    active.segmentIndex === markers.segmentIndex &&
    active.tokenField === tokenField &&
    active.tokenIndex === tokenIndex
  );
}

function buildTokenEditor(token, markers, tokenField, tokenIndex) {
  const panel = document.createElement("div");
  panel.className = "token-editor";
  panel.addEventListener("click", (event) => event.stopPropagation());
  panel.addEventListener("mousedown", (event) => event.stopPropagation());

  const preferredKeys = ["w", "i", "bf"];
  const orderedKeys = [
    ...preferredKeys,
    ...Object.keys(token).filter((key) => !preferredKeys.includes(key)),
  ];

  orderedKeys.forEach((key) => {
    const value = Object.prototype.hasOwnProperty.call(token, key) ? token[key] : "";
    const row = document.createElement("label");
    row.className = "token-editor-row";

    const keyEl = document.createElement("span");
    keyEl.className = "token-editor-key";
    keyEl.textContent = key;

    const input = document.createElement("input");
    input.className = "token-editor-input";
    input.name = key;
    input.dataset.tokenKey = key;
    input.value = value === undefined || value === null ? "" : String(value);

    row.appendChild(keyEl);
    row.appendChild(input);

    panel.appendChild(row);
  });

  const actions = document.createElement("div");
  actions.className = "token-editor-actions";

  const confirmBtn = document.createElement("button");
  confirmBtn.type = "button";
  confirmBtn.className = "btn";
  confirmBtn.textContent = "确认";
  confirmBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    saveTokenEdit(markers, tokenField, tokenIndex, token, panel);
  });

  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "btn btn-muted";
  cancelBtn.textContent = "取消";
  cancelBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    closeTokenEditor();
  });

  const insertBtn = document.createElement("button");
  insertBtn.type = "button";
  insertBtn.className = "btn btn-muted";
  insertBtn.textContent = "新增";
  insertBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    insertTokenAfter(markers, tokenField, tokenIndex, token);
  });

  actions.appendChild(insertBtn);
  actions.appendChild(confirmBtn);
  actions.appendChild(cancelBtn);
  panel.appendChild(actions);
  return panel;
}

function saveTokenEdit(markers, tokenField, tokenIndex, originalToken, panel) {
  const file = getActiveFile();
  if (!file || !file.parsed.ok) {
    return;
  }

  const segment = getSegmentByLocation(file.parsed.value, markers.sectionIndex, markers.rowIndex, markers.segmentIndex);
  const token = segment?.[tokenField]?.[tokenIndex];
  if (!token || typeof token !== "object") {
    setStatusError("无法定位到要修改的 token。");
    return;
  }

  const inputMap = new Map(
    Array.from(panel.querySelectorAll(".token-editor-input")).map((input) => [input.dataset.tokenKey, input])
  );

  const preferredKeys = ["w", "i", "bf"];
  const orderedKeys = [
    ...preferredKeys,
    ...Object.keys(originalToken).filter((key) => !preferredKeys.includes(key)),
  ];

  orderedKeys.forEach((key) => {
    const input = inputMap.get(key);
    if (!input) {
      return;
    }
    const originalValue = Object.prototype.hasOwnProperty.call(originalToken, key) ? originalToken[key] : getDefaultTokenValue(key);
    token[key] = castTokenValue(input.value, originalValue);
  });

  syncFileAfterTokenMutation(file, null, "token 已更新");
}

function insertTokenAfter(markers, tokenField, tokenIndex, token) {
  const file = getActiveFile();
  if (!file || !file.parsed.ok) {
    return;
  }

  const segment = getSegmentByLocation(file.parsed.value, markers.sectionIndex, markers.rowIndex, markers.segmentIndex);
  const tokens = segment?.[tokenField];
  if (!Array.isArray(tokens)) {
    setStatusError("无法在当前内容中新增 token。");
    return;
  }

  const nextToken = {
    w: "",
    i: typeof token?.i === "number" ? token.i : 0,
    bf: typeof token?.bf === "string" ? token.bf : "",
  };

  tokens.splice(tokenIndex + 1, 0, nextToken);

  syncFileAfterTokenMutation(
    file,
    {
      sectionIndex: markers.sectionIndex,
      rowIndex: markers.rowIndex,
      segmentIndex: markers.segmentIndex,
      tokenField,
      tokenIndex: tokenIndex + 1,
    },
    "已新增后续 token"
  );
}

function syncFileAfterTokenMutation(file, nextActiveTokenEditor, message) {
  file.raw = JSON.stringify(file.parsed.value, null, 2);
  editorEl.value = file.raw;
  renderHighlight(file.raw);
  state.activeTokenEditor = nextActiveTokenEditor;
  renderPreviewSafe(file.parsed.value, file.name);
  setStatusInfo(message);
}

function showToast(message, tone = "error") {
  if (!toastContainerEl) {
    return;
  }

  const toast = document.createElement("div");
  toast.className = `toast ${tone === "error" ? "toast-error" : ""}`.trim();
  toast.textContent = message;
  toastContainerEl.appendChild(toast);

  window.setTimeout(() => {
    toast.remove();
  }, 4000);
}

function castTokenValue(rawValue, originalValue) {
  if (typeof originalValue === "number") {
    const next = Number(rawValue);
    return Number.isFinite(next) ? next : originalValue;
  }

  if (typeof originalValue === "boolean") {
    if (rawValue === "true") return true;
    if (rawValue === "false") return false;
    return originalValue;
  }

  return rawValue;
}

function getDefaultTokenValue(key) {
  if (key === "i") {
    return 0;
  }
  return "";
}

function getSectionTitleForSection(section) {
  const rows = Array.isArray(section?.rows) ? section.rows : section;
  for (const row of rows) {
    const title = getSectionTitleForRow(row);
    if (title) {
      return title;
    }
  }
  return "";
}

function getInferredSectionType(rows) {
  return getSectionTitleForSection({ rows });
}

function getDisplaySectionType(section) {
  return section.sectionType || section.inferredSectionType || "";
}

function getPrimaryContentType(row) {
  for (const segment of row) {
    const contentType = String(segment?.contentType || "");
    if (contentType && contentType !== "instruction") {
      return contentType;
    }
  }
  return "";
}

function getChoiceOptionType(row) {
  for (const segment of row) {
    const contentType = String(segment?.contentType || "");
    if (isChoiceOptionType(contentType)) {
      return contentType;
    }
  }
  return "";
}

function getDataRow(root, sectionIndex, rowIndex) {
  const sections = extractSections(root);
  const section = sections[sectionIndex];
  if (!section || !Array.isArray(section.rows)) {
    return null;
  }
  return section.rows?.[rowIndex] ?? null;
}

function getSegmentByLocation(root, sectionIndex, rowIndex, segmentIndex) {
  const row = getDataRow(root, sectionIndex, rowIndex);
  if (!Array.isArray(row)) {
    return null;
  }
  return row?.[segmentIndex] ?? null;
}

function upgradeRootToSectionObjects(root, sections = extractSections(root)) {
  if (!root || typeof root !== "object" || Array.isArray(root)) {
    return;
  }

  root.data = sections.map((section) => {
    if (section.hasSectionWrapper && section.source && typeof section.source === "object" && !Array.isArray(section.source)) {
      return {
        ...section.source,
        sectionType: typeof section.source.sectionType === "string" ? section.source.sectionType : "",
        section: section.rows,
      };
    }

    return {
      sectionType: "",
      section: section.rows,
    };
  });
}

function safeParseJson(text) {
  try {
    return { ok: true, value: JSON.parse(text) };
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

function prettyJson(text) {
  const parsed = safeParseJson(text);
  if (!parsed.ok) {
    return text;
  }
  return JSON.stringify(parsed.value, null, 2);
}

function syncEditorScroll() {
  highlightEl.scrollTop = editorEl.scrollTop;
  highlightEl.scrollLeft = editorEl.scrollLeft;
}

function renderHighlight(raw) {
  const escaped = escapeHtml(raw)
    .replace(/("(?:\\u[\da-fA-F]{4}|\\[^u]|[^\\"])*")\s*:/g, '<span class="token-key">$1</span>:')
    .replace(/:\s*("(?:\\u[\da-fA-F]{4}|\\[^u]|[^\\"])*")/g, ': <span class="token-string">$1</span>')
    .replace(/\b(-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)\b/g, '<span class="token-number">$1</span>')
    .replace(/\b(true|false)\b/g, '<span class="token-boolean">$1</span>')
    .replace(/\bnull\b/g, '<span class="token-null">null</span>');

  highlightEl.innerHTML = escaped || " ";
}

function escapeHtml(input) {
  return String(input)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function renderMetaSummary(json) {
  if (!json || typeof json !== "object" || Array.isArray(json)) {
    return;
  }

  const meta = json.meta && typeof json.meta === "object" ? json.meta : {};
  const rows = [
    ["文章ID", meta.articleId ?? 0],
    ["语种", toMappedLabel(meta.wordType, WORD_TYPE_OPTIONS)],
    ["地区", meta.region ?? ""],
    ["考试", toMappedLabel(meta.examId, EXAM_ID_OPTIONS)],
    ["年份", meta.year ?? 0],
    ["月份", meta.month ?? 0],
    ["文章类型", toMappedLabel(meta.articleType, ARTICLE_TYPE_OPTIONS)],
  ];

  const card = document.createElement("section");
  card.className = "meta-card";
  const title = document.createElement("h3");
  title.textContent = "Meta 信息";
  card.appendChild(title);

  rows.forEach(([label, value]) => {
    const line = document.createElement("p");
    line.className = "meta-row";
    line.innerHTML = `<strong>${escapeHtml(label)}</strong>：${escapeHtml(value)}`;
    card.appendChild(line);
  });

  previewEl.appendChild(card);
}

function toMappedLabel(rawValue, options) {
  const value = Number.parseInt(String(rawValue ?? 0), 10);
  const found = options.find(([code]) => code === value);
  if (!found) return String(rawValue ?? 0);
  return `${found[0]} ${found[1]}`;
}

function jumpToSegmentInEditor(contentType, occurrence) {
  const raw = editorEl.value || "";
  const escapedType = escapeRegExp(contentType);
  const pattern = new RegExp(`"contentType"\\s*:\\s*"${escapedType}"`, "g");
  let match;
  let seen = 0;

  while ((match = pattern.exec(raw))) {
    seen += 1;
    if (seen === occurrence) {
      const range = findEnclosingJsonObjectRange(raw, match.index);
      if (range) {
        focusEditorRange(range.start, range.end);
      } else {
        focusEditorAt(match.index);
      }
      setStatusInfo(`已定位并选中 ${contentType} #${occurrence}`);
      return;
    }
  }

  setStatusError(`在源码中未找到 ${contentType} #${occurrence}。`);
}

function focusEditorAt(index) {
  editorEl.focus();
  editorEl.setSelectionRange(index, index);

  const before = editorEl.value.slice(0, index);
  const line = before.split("\n").length - 1;
  const lineHeight = parseFloat(getComputedStyle(editorEl).lineHeight) || 20;
  const target = Math.max(0, line * lineHeight - editorEl.clientHeight / 2);
  editorEl.scrollTop = target;
  syncEditorScroll();
}

function focusEditorRange(start, end) {
  editorEl.focus();
  editorEl.setSelectionRange(start, end);

  const before = editorEl.value.slice(0, start);
  const line = before.split("\n").length - 1;
  const lineHeight = parseFloat(getComputedStyle(editorEl).lineHeight) || 20;
  const target = Math.max(0, line * lineHeight - editorEl.clientHeight / 2);
  editorEl.scrollTop = target;
  syncEditorScroll();
}

function findEnclosingJsonObjectRange(raw, index) {
  const start = findOpeningBrace(raw, index);
  if (start === -1) {
    return null;
  }

  const end = findClosingBrace(raw, start);
  if (end === -1) {
    return null;
  }

  return { start, end: end + 1 };
}

function findOpeningBrace(raw, index) {
  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let i = index; i >= 0; i -= 1) {
    const ch = raw[i];

    if (inString) {
      if (escaped) {
        escaped = false;
      } else if (ch === "\\") {
        escaped = true;
      } else if (ch === "\"") {
        inString = false;
      }
      continue;
    }

    if (ch === "\"") {
      inString = true;
      continue;
    }

    if (ch === "}") {
      depth += 1;
      continue;
    }

    if (ch === "{") {
      if (depth === 0) {
        return i;
      }
      depth -= 1;
    }
  }

  return -1;
}

function findClosingBrace(raw, start) {
  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let i = start; i < raw.length; i += 1) {
    const ch = raw[i];

    if (inString) {
      if (escaped) {
        escaped = false;
      } else if (ch === "\\") {
        escaped = true;
      } else if (ch === "\"") {
        inString = false;
      }
      continue;
    }

    if (ch === "\"") {
      inString = true;
      continue;
    }

    if (ch === "{") {
      depth += 1;
      continue;
    }

    if (ch === "}") {
      depth -= 1;
      if (depth === 0) {
        return i;
      }
    }
  }

  return -1;
}

async function onSpacify() {
  const file = getActiveFile();
  if (!file) {
    return;
  }

  if (!file.parsed.ok) {
    setStatusError("请先修复 JSON 错误，再执行空格修复。");
    return;
  }

  const result = spacifyJson(file.parsed.value);
  if (result.spacesInserted === 0) {
    setStatusInfo("空格修复完成：未检测到缺失空格。");
    return;
  }

  file.parsed = { ok: true, value: result.value };
  file.raw = JSON.stringify(result.value, null, 2);

  editorEl.value = file.raw;
  renderHighlight(file.raw);
  renderPreviewSafe(file.parsed.value, file.name);
  setStatusInfo(`空格修复完成：在 ${result.arraysUpdated} 个句子中插入 ${result.spacesInserted} 个空格 token。`);
}

async function onSaveAs() {
  const file = getActiveFile();
  if (!file) {
    return;
  }

  const parsed = safeParseJson(editorEl.value);
  if (!parsed.ok) {
    setStatusError("JSON 无效，无法另存为。");
    return;
  }

  const output = JSON.stringify(parsed.value, null, 2);
  const suggestedName = file.name.endsWith(".json") ? file.name.replace(/\.json$/i, "-edited.json") : `${file.name}-edited.json`;

  const saved = await saveTextContent({
    suggestedName,
    content: output,
    description: "JSON 文件",
    mimeType: "application/json",
    extensions: [".json"],
    pickerId: "json-reader-save",
  });

  if (saved === "saved") {
    setStatusInfo("文件已保存。");
  } else if (saved === "downloaded") {
    setStatusInfo("已开始下载。");
  }
}

async function onExportTxt() {
  const file = getActiveFile();
  if (!file) {
    return;
  }

  const parsed = safeParseJson(editorEl.value);
  if (!parsed.ok) {
    setStatusError("JSON 无效，无法导出 txt。");
    return;
  }

  const output = renderTxtExport(parsed.value);
  const suggestedName = file.name.endsWith(".json") ? file.name.replace(/\.json$/i, ".txt") : `${file.name}.txt`;

  const saved = await saveTextContent({
    suggestedName,
    content: output,
    description: "TXT 文件",
    mimeType: "text/plain",
    extensions: [".txt"],
    pickerId: "json-reader-save",
  });

  if (saved === "saved") {
    setStatusInfo("TXT 已导出。");
  } else if (saved === "downloaded") {
    setStatusInfo("TXT 已开始下载。");
  }
}

async function saveTextContent({ suggestedName, content, description, mimeType, extensions, pickerId }) {
  try {
    if (window.electronAPI?.isElectron && typeof window.electronAPI.saveFile === "function") {
      const result = await window.electronAPI.saveFile({
        suggestedName,
        content,
      });
      if (result?.ok) {
        return "saved";
      }
      if (result?.message) {
        showToast(`保存失败：${result.message}`, "error");
        setStatusError(`保存失败：${result.message}`);
        return null;
      }
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    showToast(`保存失败：${message}`, "error");
    setStatusError(`保存失败：${message}`);
    return null;
  }

  try {
    if (typeof window.showSaveFilePicker === "function") {
      const pickerOptions = {
        suggestedName,
        id: pickerId,
        types: [{ description, accept: { [mimeType]: extensions } }],
      };
      const lastDirectory = await getLastSaveDirectoryHandle();
      if (lastDirectory) {
        pickerOptions.startIn = lastDirectory;
      }
      const handle = await window.showSaveFilePicker(pickerOptions);
      const writable = await handle.createWritable();
      await writable.write(content);
      await writable.close();
      await storeLastSaveDirectoryHandle(handle);
      return "saved";
    }
  } catch (error) {
    if (error && error.name === "AbortError") {
      return null;
    }
  }

  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = suggestedName;
  anchor.click();
  URL.revokeObjectURL(url);
  return "downloaded";
}

function openSaveDirectoryDb() {
  if (typeof indexedDB === "undefined") {
    return Promise.resolve(null);
  }

  return new Promise((resolve, reject) => {
    const request = indexedDB.open(SAVE_DIRECTORY_DB_NAME, 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(SAVE_DIRECTORY_STORE_NAME);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getLastSaveDirectoryHandle() {
  try {
    const db = await openSaveDirectoryDb();
    if (!db) {
      return null;
    }
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(SAVE_DIRECTORY_STORE_NAME, "readonly");
      const store = tx.objectStore(SAVE_DIRECTORY_STORE_NAME);
      const request = store.get(SAVE_DIRECTORY_KEY);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
  } catch {
    return null;
  }
}

async function storeLastSaveDirectoryHandle(handle) {
  try {
    const db = await openSaveDirectoryDb();
    if (!db) {
      return;
    }
    await new Promise((resolve, reject) => {
      const tx = db.transaction(SAVE_DIRECTORY_STORE_NAME, "readwrite");
      const store = tx.objectStore(SAVE_DIRECTORY_STORE_NAME);
      const request = store.put(handle, SAVE_DIRECTORY_KEY);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  } catch {
    // Ignore storage failures.
  }
}

function renderTxtExport(json) {
  const sections = extractSections(json);
  if (!sections.length) {
    return "";
  }

  const blocks = sections.map((section) => renderTxtSection(section));
  return blocks.filter(Boolean).join("\n\n");
}

function renderTxtSection(section) {
  const lines = [];
  lines.push(`[SECTION_TYPE-${getDisplaySectionType(section)}]`);

  const questionCounters = new Map();
  let clozeOptionGroupIndex = 0;

  section.rows.forEach((row) => {
    const choiceOptionType = getChoiceOptionType(row);
    const clozeQuestionNumber = choiceOptionType === "clozeOption" ? ++clozeOptionGroupIndex : null;
    let optionIndex = 0;

    const rowLines = [];

    row.forEach((segment) => {
      const contentType = String(segment?.contentType || "");
      const lower = contentType.toLowerCase();
      const text = extractReadableText(segment);
      const hasLeadingQuestionNo = startsWithQuestionNumber(text);
      let questionNumber = null;
      let optionLabel = "";

      if (lower.includes("question") && contentType !== "clozeQuestion") {
        const next = (questionCounters.get(contentType) || 0) + 1;
        questionCounters.set(contentType, next);
        questionNumber = next;
      }

      if (choiceOptionType && contentType === choiceOptionType) {
        optionIndex += 1;
        optionLabel = `${toAlphabetLabel(optionIndex)}. `;
        if (contentType === "clozeOption") {
          questionNumber = optionIndex === 1 ? clozeQuestionNumber : null;
        } else {
          questionNumber = null;
        }
      }

      rowLines.push(...renderTxtSegmentLines(segment, {
        questionNumber,
        optionLabel,
        canShowAutoQuestionNo: !(lower.includes("question") && hasLeadingQuestionNo),
      }));
    });

    if (!rowLines.length) {
      return;
    }

    if (shouldJoinRowInline(row)) {
      lines.push(rowLines.join(""));
    } else {
      lines.push(...rowLines);
    }
  });

  return lines.join("\n");
}

function renderTxtSegmentLines(segment, markers = {}) {
  const contentType = String(segment?.contentType || "");

  if (contentType === "instruction") {
    const instruction = stringifyInstruction(segment);
    const prefix = markers.questionNumber ? `${markers.questionNumber}. ` : "";
    return [`${prefix}${instruction}`];
  }

  if (contentType === "clozeOption") {
    const lines = [];
    if (markers.questionNumber && markers.canShowAutoQuestionNo) {
      lines.push(`${markers.questionNumber}.`);
    }
    lines.push(`${markers.optionLabel || ""}${extractReadableText(segment)}`);
    return lines;
  }

  const prefix = markers.canShowAutoQuestionNo && markers.questionNumber
    ? `${markers.questionNumber}. ${markers.optionLabel || ""}`
    : (markers.optionLabel || "");

  return [`${prefix}${extractReadableText(segment)}`];
}

function shouldJoinRowInline(row) {
  const contentTypes = row.map((segment) => String(segment?.contentType || ""));
  if (!contentTypes.length) {
    return true;
  }
  return contentTypes.every((contentType) => {
    return contentType === "body" || contentType === "clozeQuestion" || contentType === "instruction";
  });
}

function populateMetaSelect(selectEl, options) {
  if (!selectEl) return;

  selectEl.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "请选择";
  selectEl.appendChild(placeholder);

  options.forEach(([value, label]) => {
    const option = document.createElement("option");
    option.value = String(value);
    option.textContent = `${value} ${label}`;
    selectEl.appendChild(option);
  });
}

function onEditMeta() {
  const file = getActiveFile();
  if (!file) {
    return;
  }

  const parsed = safeParseJson(editorEl.value);
  if (!parsed.ok) {
    setStatusError("请先修复 JSON 错误，再编辑 Meta。");
    return;
  }

  if (!parsed.value || typeof parsed.value !== "object" || Array.isArray(parsed.value)) {
    setStatusError("Meta 编辑器要求顶层为 JSON 对象。");
    return;
  }

  const currentMeta = parsed.value.meta && typeof parsed.value.meta === "object" ? parsed.value.meta : {};
  fillMetaForm(currentMeta);
  openMetaModal();
}

function fillMetaForm(meta) {
  setMetaField("articleId", toFormValue(meta.articleId));
  setMetaField("wordType", toFormValue(meta.wordType));
  setMetaField("region", toFormValue(meta.region));
  setMetaField("examId", toFormValue(meta.examId));
  setMetaField("exam", toFormValue(meta.exam));
  setMetaField("year", toFormValue(meta.year));
  setMetaField("month", toFormValue(meta.month));
  setMetaField("articleType", toFormValue(meta.articleType));
}

function setMetaField(name, value) {
  const field = metaFormEl?.elements?.namedItem(name);
  if (field) {
    field.value = value;
  }
}

function toFormValue(value) {
  return value === undefined || value === null ? "" : String(value);
}

function openMetaModal() {
  if (!metaModalEl) return;
  metaModalEl.classList.add("open");
  metaModalEl.setAttribute("aria-hidden", "false");
}

function closeMetaModal() {
  if (!metaModalEl) return;
  metaModalEl.classList.remove("open");
  metaModalEl.setAttribute("aria-hidden", "true");
}

function onSaveMeta(event) {
  event.preventDefault();

  const file = getActiveFile();
  if (!file) {
    return;
  }

  const parsed = safeParseJson(editorEl.value);
  if (!parsed.ok) {
    setStatusError("请先修复 JSON 错误，再保存 Meta。");
    return;
  }

  const root = parsed.value;
  if (!root || typeof root !== "object" || Array.isArray(root)) {
    setStatusError("Meta 编辑器要求顶层为 JSON 对象。");
    return;
  }

  const nextMeta = {
    articleId: parseIntField("articleId"),
    wordType: parseIntField("wordType"),
    region: parseStringField("region"),
    examId: parseIntField("examId"),
    exam: parseStringField("exam"),
    year: parseIntField("year"),
    month: parseIntField("month"),
    articleType: parseIntField("articleType"),
  };

  const updatedRoot = withMetaBeforeData(root, nextMeta);
  file.parsed = { ok: true, value: updatedRoot };
  file.raw = JSON.stringify(updatedRoot, null, 2);

  editorEl.value = file.raw;
  renderHighlight(file.raw);
  renderPreviewSafe(updatedRoot, file.name);
  closeMetaModal();
  setStatusInfo("Meta 已更新。");
}

function parseIntField(name) {
  const field = metaFormEl?.elements?.namedItem(name);
  const text = field ? String(field.value).trim() : "";
  if (!text) return 0;
  const number = Number.parseInt(text, 10);
  return Number.isFinite(number) ? number : 0;
}

function parseStringField(name) {
  const field = metaFormEl?.elements?.namedItem(name);
  const text = field ? String(field.value).trim() : "";
  return text || "";
}

function withMetaBeforeData(root, meta) {
  const ordered = { meta };

  if (Object.prototype.hasOwnProperty.call(root, "data")) {
    ordered.data = root.data;
  }

  Object.keys(root).forEach((key) => {
    if (key !== "meta" && key !== "data") {
      ordered[key] = root[key];
    }
  });

  return ordered;
}

function spacifyJson(root) {
  let arraysUpdated = 0;
  let spacesInserted = 0;

  function walk(node) {
    if (Array.isArray(node)) {
      if (isTokenArray(node)) {
        const updated = spacifyTokenArray(node);
        if (updated.spacesInserted > 0) {
          arraysUpdated += 1;
          spacesInserted += updated.spacesInserted;
        }
        return updated.tokens;
      }
      return node.map((item) => walk(item));
    }

    if (node && typeof node === "object") {
      const next = {};
      Object.entries(node).forEach(([key, value]) => {
        next[key] = walk(value);
      });
      return next;
    }

    return node;
  }

  return { value: walk(root), arraysUpdated, spacesInserted };
}

function isTokenArray(value) {
  return Array.isArray(value) && value.length > 0 && value.every((item) => item && typeof item === "object" && "w" in item);
}

function spacifyTokenArray(tokens) {
  const ordered = tokens.map((token) => token);

  if (!isLikelyEnglishTokens(ordered)) {
    return { tokens: ordered, spacesInserted: 0 };
  }

  const rebuilt = [];
  let previousWord = "";
  let previousIsSpace = false;
  let spacesInserted = 0;

  ordered.forEach((originalToken) => {
    const current = { ...originalToken, w: String(originalToken.w ?? "") };

    if (isSpaceToken(current.w)) {
      if (!previousIsSpace && rebuilt.length) {
        rebuilt.push({ ...current, w: " " });
        previousIsSpace = true;
      }
      return;
    }

    if (previousWord && !previousIsSpace && shouldInsertSpace(previousWord, current.w)) {
      rebuilt.push({ w: " ", i: -1 });
      spacesInserted += 1;
    }

    rebuilt.push(current);
    previousWord = current.w;
    previousIsSpace = false;
  });

  const normalized = rebuilt.map((token, index) => ({ ...token, i: index }));
  return { tokens: normalized, spacesInserted };
}

function shouldInsertSpace(leftToken, rightToken) {
  const left = leftToken.trim();
  const right = rightToken.trim();
  if (!left || !right) return false;
  if (containsCjk(left) || containsCjk(right)) return false;
  if (/^'[A-Za-z]/.test(right)) return false;
  if (isNoSpaceBefore(right)) return false;
  if (isNoSpaceAfter(left)) return false;
  if (isSentencePunctuation(left)) return true;
  if (isWordLike(left) && isWordLike(right)) return true;
  if (isClosingBracket(left) && isWordLike(right)) return true;
  if (isWordLike(left) && right === "(") return true;
  return false;
}

function isLikelyEnglishTokens(tokens) {
  let latin = 0;
  let cjk = 0;

  tokens.forEach((token) => {
    const word = String(token.w ?? "");
    if (/[A-Za-z]/.test(word)) latin += 1;
    if (containsCjk(word)) cjk += 1;
  });

  return latin > 0 && latin >= cjk;
}

function isSpaceToken(value) {
  return /^\s+$/.test(value);
}

function isWordLike(value) {
  return /[A-Za-z0-9]/.test(value);
}

function isSentencePunctuation(value) {
  return /^[,.;:!?]+$/.test(value);
}

function isNoSpaceBefore(value) {
  return /^[,.;:!?%)\]}]+$/.test(value);
}

function isNoSpaceAfter(value) {
  return /^[(\[{]+$/.test(value);
}

function isClosingBracket(value) {
  return /^[)\]}]+$/.test(value);
}

function containsCjk(value) {
  return /[\u3400-\u9FBF]/.test(value);
}

function startsWithQuestionNumber(text) {
  return /^\s*\d+([.)]|(、))?\s*/.test(String(text || ""));
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function setEmptyState() {
  fileListEl.innerHTML = "";
  previewEl.innerHTML = '<p class="placeholder">请先加载一个或多个 JSON 文件。</p>';
  previewMetaEl.textContent = "";
  editorEl.value = "";
  renderHighlight("");
  editorStatusEl.textContent = "";
  editorStatusEl.className = "editor-status meta";
}

function setStatusOk() {
  editorStatusEl.textContent = "";
  editorStatusEl.className = "editor-status meta";
}

function setStatusError(message) {
  editorStatusEl.textContent = `JSON 无效 · ${message}`;
  editorStatusEl.className = "editor-status meta status-error";
}

function setStatusInfo(message) {
  editorStatusEl.textContent = message;
  editorStatusEl.className = "editor-status meta status-ok";
}

function onEditorKeydown(event) {
  if (event.key !== "Tab") {
    return;
  }

  event.preventDefault();
  const value = editorEl.value;
  const start = editorEl.selectionStart;
  const end = editorEl.selectionEnd;

  if (event.shiftKey) {
    const lineStart = value.lastIndexOf("\n", start - 1) + 1;
    const selectedText = value.slice(lineStart, end);
    const lines = selectedText.split("\n");

    let removedChars = 0;
    const outdented = lines.map((line) => {
      if (line.startsWith("\t")) {
        removedChars += 1;
        return line.slice(1);
      }
      if (line.startsWith("  ")) {
        removedChars += 2;
        return line.slice(2);
      }
      return line;
    });

    const replacement = outdented.join("\n");
    editorEl.value = value.slice(0, lineStart) + replacement + value.slice(end);
    const newStart = Math.max(lineStart, start - (start > lineStart ? 1 : 0));
    const newEnd = Math.max(newStart, end - removedChars);
    editorEl.setSelectionRange(newStart, newEnd);
  } else if (start !== end) {
    const lineStart = value.lastIndexOf("\n", start - 1) + 1;
    const selectedText = value.slice(lineStart, end);
    const lines = selectedText.split("\n");
    const indented = lines.map((line) => `\t${line}`).join("\n");

    editorEl.value = value.slice(0, lineStart) + indented + value.slice(end);
    const addedChars = lines.length;
    editorEl.setSelectionRange(start + 1, end + addedChars);
  } else {
    editorEl.value = value.slice(0, start) + "\t" + value.slice(end);
    editorEl.setSelectionRange(start + 1, start + 1);
  }

  onEditorInput();
}

function getActiveFile() {
  return state.files.find((file) => file.id === state.activeId) || null;
}
