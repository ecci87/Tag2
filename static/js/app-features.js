function getCloneSelectionPaths() {
  return state.selectedPaths.size > 1 ? [...state.selectedPaths] : [];
}

function getDefaultCloneFolderName() {
  const baseName = getFileLabel(state.folder || "folder") || "folder";
  return `${baseName}-copy`;
}

function getDefaultDuplicateImageName(path = state.previewPath) {
  const fileName = getFileLabel(path || "image") || "image";
  const extension = getImageExtension(fileName);
  const stem = extension ? fileName.slice(0, -extension.length) : fileName;
  return `${stem}-copy${extension}`;
}

async function cloneCurrentFolder() {
  if (!state.folder) {
    showErrorToast("Load a folder first.");
    return;
  }
  if (state.cloning || state.autoCaptioning) {
    showErrorToast("Finish the current operation before cloning.");
    return;
  }

  const selectedPaths = getCloneSelectionPaths();
  const selectionLabel = selectedPaths.length > 1
    ? `${selectedPaths.length} selected media files`
    : "the whole folder";
  const newFolderName = prompt(`Clone ${selectionLabel} into a new sibling folder named:`, getDefaultCloneFolderName());
  if (newFolderName === null) return;
  const trimmedName = String(newFolderName || "").trim();
  if (!trimmedName) {
    showErrorToast("Clone cancelled: folder name is required.");
    return;
  }

  state.cloning = true;
  updateActionButtons();
  state.modelLogLines = [];
  state.modelLogOpen = false;
  resetAutoCaptionProgress();
  updateAutoCaptionProgress({
    visible: true,
    scopeLabel: selectedPaths.length > 1 ? "Clone Selection" : "Clone Folder",
    totalImages: 1,
    processedImages: 0,
    completedImages: 0,
    errors: 0,
    currentMessage: "Preparing clone...",
    currentStepIndex: 0,
    currentStepTotal: 1,
  });
  statusBar.textContent = `Cloning to ${trimmedName}...`;

  try {
    const resp = await fetch("/api/folder/clone/stream", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        source_folder: state.folder,
        new_folder_name: trimmedName,
        image_paths: selectedPaths,
      }),
    });

    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      throw new Error(data.detail || "Failed to clone folder");
    }

    let targetFolder = "";
    await readNdjsonStream(resp, (event) => {
      if (!event || typeof event !== "object") return;
      if (event.type === "start") {
        targetFolder = event.target_folder || targetFolder;
        updateAutoCaptionProgress({
          visible: true,
          scopeLabel: event.mode === "selected" ? "Clone Selection" : "Clone Folder",
          totalImages: Math.max(1, Number(event.total || 0)),
          processedImages: 0,
          completedImages: 0,
          errors: 0,
          currentPath: event.target_folder || "",
          currentMessage: "Creating folder...",
          currentStepIndex: 0,
          currentStepTotal: 1,
        });
        return;
      }
      if (event.type === "progress") {
        updateAutoCaptionProgress({
          visible: true,
          totalImages: Math.max(1, Number(event.total || state.aiProgress.totalImages || 1)),
          processedImages: Number(event.index || 0),
          completedImages: Number(event.index || 0),
          currentPath: event.path || "",
          currentMessage: "Copied",
          currentStepIndex: 1,
          currentStepTotal: 1,
        });
        return;
      }
      if (event.type === "config-copied") {
        targetFolder = event.target_folder || targetFolder;
        updateAutoCaptionProgress({
          currentPath: event.target_folder || "",
          currentMessage: "Copied config",
          currentStepIndex: 1,
          currentStepTotal: 1,
        });
        return;
      }
      if (event.type === "done") {
        targetFolder = event.target_folder || targetFolder;
        updateAutoCaptionProgress({
          visible: true,
          totalImages: Math.max(1, Number(event.total || state.aiProgress.totalImages || 1)),
          processedImages: Number(event.total || state.aiProgress.totalImages || 1),
          completedImages: Number(event.total || state.aiProgress.totalImages || 1),
          currentPath: event.target_folder || "",
          currentMessage: "Clone complete",
          currentStepIndex: 1,
          currentStepTotal: 1,
        });
        return;
      }
      if (event.type === "error") {
        throw new Error(event.message || "Clone failed");
      }
    });

    if (!targetFolder) {
      throw new Error("Clone finished without a target folder");
    }

    folderInput.value = targetFolder;
    await loadFolder();
    statusBar.textContent = `Cloned into ${getFileLabel(targetFolder)}`;
  } catch (err) {
    const message = err?.message || "Failed to clone folder";
    statusBar.textContent = `Clone error: ${message}`;
    showErrorToast(`Clone error: ${message}`);
  } finally {
    state.cloning = false;
    updateActionButtons();
    resetAutoCaptionProgress();
  }
}

async function duplicateCurrentImage() {
  if (!isImageEditAvailable()) {
    showErrorToast("Select a single image first.");
    return;
  }
  if (state.duplicatingImage || state.autoCaptioning || state.cloning || state.uploading) {
    showErrorToast("Finish the current operation before duplicating an image.");
    return;
  }

  const sourcePath = state.previewPath;
  const requestedName = prompt(`Duplicate ${getFileLabel(sourcePath)} as:`, getDefaultDuplicateImageName(sourcePath));
  if (requestedName === null) return;
  const trimmedName = String(requestedName || "").trim();
  if (!trimmedName) {
    showErrorToast("Duplicate cancelled: image name is required.");
    return;
  }

  state.duplicatingImage = true;
  renderMaskEditorUi();
  statusBar.textContent = `Duplicating ${getFileLabel(sourcePath)}...`;
  try {
    const resp = await fetch("/api/image/duplicate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        image_path: sourcePath,
        new_name: trimmedName,
      }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to duplicate image");
    }
    const preserveScrollTop = fileGridContainer.scrollTop;
    await loadFolder({ preserveScrollTop });
    await selectUploadedImages([data.image_path]);
    statusBar.textContent = `Duplicated ${getFileLabel(sourcePath)} as ${getFileLabel(data.image_path)}`;
  } catch (err) {
    showErrorToast(`Duplicate error: ${err.message}`);
    statusBar.textContent = `Duplicate error: ${err.message}`;
  } finally {
    state.duplicatingImage = false;
    renderMaskEditorUi();
  }
}

function createEmptyGroup(name = "") {
  const id = makeUiId("group");
  return { id, name, sentences: [], hidden_sentences: [], _uiId: id };
}

function createEmptySection(name = "") {
  return { name, sentences: [], groups: [], item_order: [], _uiId: makeUiId("section") };
}

function createSentenceOrderItem(sentence) {
  return { type: "sentence", sentence };
}

function createGroupOrderItem(groupId) {
  return { type: "group", group_id: groupId };
}

function getSectionOrderItemKey(item) {
  if (!item || typeof item !== "object") return "";
  if (item.type === "sentence" || item.type === "caption") return `sentence:${item.sentence || item.caption || ""}`;
  if (item.type === "group") return `group:${item.group_id || ""}`;
  return "";
}

function getNormalizedSectionItemOrder(section, sentences, groups) {
  const validSentences = new Set(sentences);
  const validGroupIds = new Set(groups.map(group => group.id).filter(Boolean));
  const orderedItems = [];
  const seenKeys = new Set();

  for (const rawItem of (Array.isArray(section?.item_order) ? section.item_order : [])) {
    if (!rawItem || typeof rawItem !== "object") continue;
    const type = String(rawItem.type || "").trim().toLowerCase();
    if (type === "sentence" || type === "caption") {
      const sentence = String(rawItem.sentence || rawItem.caption || "").trim();
      const item = createSentenceOrderItem(sentence);
      const key = getSectionOrderItemKey(item);
      if (!sentence || !validSentences.has(sentence) || seenKeys.has(key)) continue;
      seenKeys.add(key);
      orderedItems.push(item);
    } else if (type === "group") {
      const groupId = String(rawItem.group_id || "").trim();
      const item = createGroupOrderItem(groupId);
      const key = getSectionOrderItemKey(item);
      if (!groupId || !validGroupIds.has(groupId) || seenKeys.has(key)) continue;
      seenKeys.add(key);
      orderedItems.push(item);
    }
  }

  for (const sentence of sentences) {
    const item = createSentenceOrderItem(sentence);
    const key = getSectionOrderItemKey(item);
    if (!seenKeys.has(key)) {
      seenKeys.add(key);
      orderedItems.push(item);
    }
  }

  for (const group of groups) {
    const item = createGroupOrderItem(group.id);
    const key = getSectionOrderItemKey(item);
    if (!seenKeys.has(key)) {
      seenKeys.add(key);
      orderedItems.push(item);
    }
  }

  return orderedItems;
}

function getOrderedSectionEntries(section) {
  const groups = Array.isArray(section?.groups) ? section.groups : [];
  const groupsById = new Map(groups.map((group, groupIdx) => [group.id, { group, groupIdx }]));
  const entries = [];

  for (const item of getNormalizedSectionItemOrder(section, section.sentences || [], groups)) {
    if (item.type === "sentence") {
      entries.push({ type: "sentence", sentence: item.sentence });
      continue;
    }
    const groupEntry = groupsById.get(item.group_id);
    if (!groupEntry) continue;
    entries.push({
      type: "group",
      groupId: item.group_id,
      group: groupEntry.group,
      groupIdx: groupEntry.groupIdx,
    });
  }

  return entries;
}

function normalizeGroupData(group) {
  const sentences = [];
  const seen = new Set();
  const sourceSentences = Array.isArray(group?.captions) ? group.captions : (Array.isArray(group?.sentences) ? group.sentences : []);
  for (const raw of sourceSentences) {
    const text = String(raw || "").trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    sentences.push(text);
  }
  const hiddenSentences = [];
  const hiddenSeen = new Set();
  const sourceHiddenSentences = Array.isArray(group?.hidden_captions)
    ? group.hidden_captions
    : (Array.isArray(group?.hidden_sentences) ? group.hidden_sentences : []);
  for (const raw of sourceHiddenSentences) {
    const text = String(raw || "").trim();
    if (!text || hiddenSeen.has(text) || !sentences.includes(text)) continue;
    hiddenSeen.add(text);
    hiddenSentences.push(text);
  }
  return {
    id: String(group?.id || group?._uiId || makeUiId("group")).trim(),
    name: String(group?.name || "").trim(),
    sentences,
    hidden_sentences: hiddenSentences,
    _uiId: group?._uiId || makeUiId("group"),
  };
}

function normalizeSectionData(section) {
  const sentences = [];
  const seen = new Set();
  const sourceSentences = Array.isArray(section?.captions) ? section.captions : (Array.isArray(section?.sentences) ? section.sentences : []);
  for (const raw of sourceSentences) {
    const text = String(raw || "").trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    sentences.push(text);
  }
  const groups = (Array.isArray(section?.groups) ? section.groups : []).map(normalizeGroupData);
  return {
    name: String(section?.name || "").trim(),
    sentences,
    groups,
    item_order: getNormalizedSectionItemOrder(section, sentences, groups),
    _uiId: section?._uiId || makeUiId("section"),
  };
}

function normalizeSectionsData(sections) {
  const normalized = (Array.isArray(sections) ? sections : []).map(normalizeSectionData);
  return normalized.length ? normalized : [createEmptySection("")];
}

function serializeSectionsForSave() {
  return normalizeSectionsData(state.sections).map(section => ({
    name: section.name,
    captions: [...section.sentences],
    groups: section.groups.map(group => ({
      id: group.id,
      name: group.name,
      captions: [...group.sentences],
      hidden_captions: [...(group.hidden_sentences || [])],
    })),
    item_order: (section.item_order || []).map(item => (item.type === "sentence" || item.type === "caption")
      ? { type: "caption", caption: item.sentence || item.caption }
      : createGroupOrderItem(item.group_id)),
  }));
}

function getAllConfiguredSentences() {
  return state.sections.flatMap(section => [
    ...(section.sentences || []),
    ...((section.groups || []).flatMap(group => group.sentences || [])),
  ]);
}

function hasConfiguredCaptions() {
  return getAllConfiguredSentences().length > 0;
}

function countSectionSentences(section) {
  return (section.sentences || []).length + ((section.groups || []).reduce((sum, group) => sum + (group.sentences || []).length, 0));
}

function findGroupForSentence(sentence) {
  for (const section of state.sections) {
    for (const group of section.groups || []) {
      if ((group.sentences || []).includes(sentence)) {
        return group;
      }
    }
  }
  return null;
}

function findGroupSentencesForSentence(sentence) {
  return findGroupForSentence(sentence)?.sentences || null;
}

function isSentenceHiddenOnExport(sentence) {
  const group = findGroupForSentence(sentence);
  return !!group && (group.hidden_sentences || []).includes(sentence);
}

function getConfiguredSentenceOrder() {
  const ordered = [];
  for (const section of normalizeSectionsData(state.sections)) {
    for (const entry of getOrderedSectionEntries(section)) {
      if (entry.type === "sentence") {
        ordered.push(entry.sentence);
      } else if (entry.type === "group") {
        ordered.push(...(entry.group?.sentences || []));
      }
    }
  }
  return ordered;
}

function orderEnabledSentences(enabledSentences) {
  const unique = [...new Set((Array.isArray(enabledSentences) ? enabledSentences : []).filter(Boolean))];
  if (unique.length <= 1) return unique;

  const configuredOrder = getConfiguredSentenceOrder();
  if (configuredOrder.length === 0) return unique;

  const orderMap = new Map(configuredOrder.map((sentence, index) => [sentence, index]));
  return [...unique].sort((a, b) => {
    const indexA = orderMap.has(a) ? orderMap.get(a) : Number.MAX_SAFE_INTEGER;
    const indexB = orderMap.has(b) ? orderMap.get(b) : Number.MAX_SAFE_INTEGER;
    return indexA - indexB;
  });
}

function getExportedEnabledSentences(enabledSentences) {
  return orderEnabledSentences(enabledSentences).filter(sentence => !isSentenceHiddenOnExport(sentence));
}

function hasEffectiveCaptionContent(caption) {
  if (!caption) return false;
  const enabledSentences = Array.isArray(caption.enabled_sentences)
    ? caption.enabled_sentences
    : (Array.isArray(caption.enabled_captions) ? caption.enabled_captions : []);
  return getExportedEnabledSentences(enabledSentences).length > 0 || !!(caption.free_text && caption.free_text.trim());
}

function normalizeCaptionCacheEntry(caption) {
  const enabledSentences = Array.isArray(caption?.enabled_captions)
    ? caption.enabled_captions
    : (Array.isArray(caption?.enabled_sentences) ? caption.enabled_sentences : []);
  return {
    enabled_sentences: orderEnabledSentences(enabledSentences),
    free_text: String(caption?.free_text || ""),
  };
}

function createEmptyMetadataCacheEntry() {
  return {
    seed: null,
    min_t: null,
    max_t: null,
    sampling_frequency: null,
  };
}

function normalizeIntegerMetadataValue(value) {
  if (value === null || value === undefined || value === "") return null;
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || !Number.isInteger(parsed)) return null;
  return parsed;
}

function normalizeFloatMetadataValue(value) {
  if (value === null || value === undefined || value === "") return null;
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return null;
  return parsed;
}

function normalizeMetadataCacheEntry(metadata) {
  const normalized = createEmptyMetadataCacheEntry();
  const seed = normalizeIntegerMetadataValue(metadata?.seed);
  const minT = normalizeIntegerMetadataValue(metadata?.min_t);
  const maxT = normalizeIntegerMetadataValue(metadata?.max_t);
  const samplingFrequency = normalizeFloatMetadataValue(metadata?.sampling_frequency);

  if (seed !== null) normalized.seed = seed;
  if (minT !== null) normalized.min_t = minT;
  if (maxT !== null) normalized.max_t = maxT;
  if (samplingFrequency !== null && samplingFrequency >= 0) {
    normalized.sampling_frequency = samplingFrequency;
  }
  return normalized;
}

function applySentenceSelectionToList(enabledSentences, sentence, shouldEnable) {
  let next = Array.isArray(enabledSentences) ? enabledSentences.filter(item => item !== sentence) : [];
  const groupSentences = findGroupSentencesForSentence(sentence);
  if (shouldEnable && groupSentences) {
    next = next.filter(item => !groupSentences.includes(item));
  }
  if (shouldEnable) {
    next.push(sentence);
  }
  return orderEnabledSentences(next);
}

function ensureCaptionCache(path) {
  if (!state.captionCache[path]) {
    state.captionCache[path] = { enabled_sentences: [], free_text: "" };
  }
  return state.captionCache[path];
}

function ensureMetadataCache(path) {
  if (!state.metadataCache[path]) {
    state.metadataCache[path] = createEmptyMetadataCacheEntry();
  }
  return state.metadataCache[path];
}

function getMetadataFieldDefinitions() {
  return [
    {
      key: "seed",
      input: metadataSeedInput,
      parse: () => parseIntegerMetadataInput(metadataSeedInput, "Seed"),
    },
    {
      key: "sampling_frequency",
      input: metadataSamplingFrequencyInput,
      parse: () => parseFloatMetadataInput(metadataSamplingFrequencyInput, "Sampling Frequency"),
    },
    {
      key: "min_t",
      input: metadataMinTInput,
      parse: () => parseIntegerMetadataInput(metadataMinTInput, "Min t"),
    },
    {
      key: "max_t",
      input: metadataMaxTInput,
      parse: () => parseIntegerMetadataInput(metadataMaxTInput, "Max t"),
    },
  ];
}

function formatMetadataFieldValue(fieldName, value) {
  if (value === null || value === undefined || value === "") return "";
  return fieldName === "sampling_frequency" ? `${value}` : `${Math.trunc(value)}`;
}

function parseIntegerMetadataInput(inputEl, label) {
  const text = String(inputEl?.value || "").trim();
  if (!text) return null;
  const parsed = Number(text);
  if (!Number.isFinite(parsed) || !Number.isInteger(parsed)) {
    throw new Error(`${label} must be a whole number`);
  }
  return parsed;
}

function parseFloatMetadataInput(inputEl, label) {
  const text = String(inputEl?.value || "").trim();
  if (!text) return null;
  const parsed = Number(text);
  if (!Number.isFinite(parsed)) {
    throw new Error(`${label} must be a number`);
  }
  if (parsed < 0) {
    throw new Error(`${label} must be greater than or equal to 0`);
  }
  return parsed;
}

function validateMetadataRange(metadata) {
  const minT = metadata?.min_t;
  const maxT = metadata?.max_t;
  if (minT !== null && minT !== undefined && maxT !== null && maxT !== undefined && minT > maxT) {
    throw new Error("Min t must be less than or equal to Max t");
  }
}

function setMetadataFieldState(field, value, placeholder = "Not set") {
  field.input.value = formatMetadataFieldValue(field.key, value);
  field.input.placeholder = placeholder;
}

function setMetadataInputsDisabled(disabled) {
  for (const field of getMetadataFieldDefinitions()) {
    field.input.disabled = disabled;
  }
}

function renderMetadataEditor(options = {}) {
  const { preserveInputs = false } = options;
  const selectedPaths = [...state.selectedPaths];
  const isSingle = selectedPaths.length === 1;
  const isMulti = selectedPaths.length > 1;
  const disabled = selectedPaths.length === 0 || state.metadataSaving;

  setMetadataInputsDisabled(disabled);
  metadataSaveBtn.disabled = disabled;
  metadataSaveBtn.textContent = state.metadataSaving
    ? (isMulti ? "Applying..." : "Saving...")
    : (isMulti ? `Apply to ${selectedPaths.length} Files` : "Save Metadata");

  if (!selectedPaths.length) {
    metadataEditorSummary.textContent = "No media selected";
    metadataEditorNote.textContent = "Select one or more media files to edit metadata sidecars.";
    for (const field of getMetadataFieldDefinitions()) {
      setMetadataFieldState(field, null, "Not set");
    }
    return;
  }

  if (isSingle) {
    const path = selectedPaths[0];
    metadataEditorSummary.textContent = getFileLabel(path);
    metadataEditorNote.textContent = "Blank fields remove those keys from this file's .meta.json sidecar.";
    if (preserveInputs) {
      return;
    }
    const metadata = ensureMetadataCache(path);
    for (const field of getMetadataFieldDefinitions()) {
      setMetadataFieldState(field, metadata[field.key], "Not set");
    }
    return;
  }

  metadataEditorSummary.textContent = `${selectedPaths.length} media files selected`;
  metadataEditorNote.textContent = "Filled fields are applied to all selected files. Blank fields leave existing values unchanged.";
  if (preserveInputs) {
    return;
  }
  for (const field of getMetadataFieldDefinitions()) {
    const values = selectedPaths.map(path => ensureMetadataCache(path)[field.key]);
    const uniqueValues = [...new Set(values.map(value => (value === null ? "__empty__" : String(value))))];
    if (uniqueValues.length === 1) {
      setMetadataFieldState(field, values[0], values[0] === null ? "Not set" : "");
    } else {
      field.input.value = "";
      field.input.placeholder = "Mixed values";
    }
  }
}

function buildSingleMetadataPayload() {
  const metadata = {};
  for (const field of getMetadataFieldDefinitions()) {
    metadata[field.key] = field.parse();
  }
  validateMetadataRange(metadata);
  return metadata;
}

function buildBatchMetadataChanges() {
  const changes = {};
  for (const field of getMetadataFieldDefinitions()) {
    if (!String(field.input.value || "").trim()) continue;
    changes[field.key] = field.parse();
  }
  validateMetadataRange(changes);
  return changes;
}

function setActiveRightPanelTab(tabName) {
  const nextTab = tabName === "metadata" ? "metadata" : "captions";
  state.ui.activeRightPanelTab = nextTab;
  rightPanelTabButtons.forEach((button) => {
    const isActive = button.dataset.rightPanelTab === nextTab;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-selected", isActive ? "true" : "false");
    button.tabIndex = isActive ? 0 : -1;
  });
  rightPanelModePanels.forEach((panel) => {
    const isActive = panel.id === `${nextTab}-editor-panel`;
    panel.classList.toggle("active", isActive);
    panel.hidden = !isActive;
  });
  renderMetadataEditor({ preserveInputs: true });
  updateMultiInfo();
}

function getPreviewEnabledSentences(path = state.previewPath) {
  if (!path) return [];
  const enabledSentences = state.captionCache[path]?.enabled_sentences;
  return orderEnabledSentences(enabledSentences);
}

function getActiveSentenceFilters() {
  return [...state.activeSentenceFilters.keys()].filter(Boolean);
}

function getActiveSentenceFilterEntries() {
  return [...state.activeSentenceFilters.entries()]
    .filter(([sentence, mode]) => !!sentence && (mode === "has" || mode === "missing"));
}

function getSentenceFilterMode(sentence) {
  const mode = state.activeSentenceFilters.get(sentence);
  return mode === "has" || mode === "missing" ? mode : "off";
}

function getActiveSentenceFilterClauses(filters = getActiveSentenceFilterEntries()) {
  const clauses = [];
  const groupedClauses = new Map();

  for (const [sentence, mode] of filters) {
    const location = findSentenceLocation(sentence);
    if (!location?.group) {
      clauses.push({ type: "sentence", sentence, mode });
      continue;
    }

    const groupId = String(location.group.id || `group-${location.sectionIndex}-${location.groupIndex}`);
    const clauseKey = `${location.sectionIndex}:${groupId}`;
    let clause = groupedClauses.get(clauseKey);
    if (!clause) {
      clause = { type: "group", mode: "has", sentences: [] };
      groupedClauses.set(clauseKey, clause);
      clauses.push(clause);
    }
    if (!clause.sentences.includes(sentence)) {
      clause.sentences.push(sentence);
    }
  }

  return clauses;
}

function formatCaptionFilterClause(clause, prefix = "  ") {
  if (!clause) return "";
  if (clause.type === "group") {
    const sentences = (clause.sentences || []).filter(Boolean);
    if (sentences.length === 0) return "";
    if (sentences.length === 1) {
      return `${prefix}"${sentences[0]}"`;
    }
    return `${prefix}(\n${sentences.map(sentence => `${prefix}  "${sentence}"`).join(" or\n")}\n${prefix})`;
  }
  return `${prefix}"${clause.sentence}"`;
}

function buildActiveFilterSummary() {
  const lines = [];
  const captionClauses = getActiveSentenceFilterClauses();
  const hasClauses = captionClauses.filter(clause => clause.type === "group" || clause.mode !== "missing");
  const missingClauses = captionClauses.filter(clause => clause.type !== "group" && clause.mode === "missing");

  if (hasClauses.length > 0) {
    lines.push("Has:");
    hasClauses.forEach((clause, index) => {
      const text = formatCaptionFilterClause(clause, "  ");
      if (!text) return;
      if (index > 0) {
        lines.push("  and");
      }
      lines.push(text);
    });
  }

  if (missingClauses.length > 0) {
    if (lines.length > 0) lines.push("");
    lines.push("Has not:");
    missingClauses.forEach((clause, index) => {
      const text = formatCaptionFilterClause(clause, "  ");
      if (!text) return;
      if (index > 0) {
        lines.push("  or");
      }
      lines.push(text);
    });
  }

  const metaHas = [];
  const metaMissing = [];
  if (state.activeMetaFilters.aspectState === "has") metaHas.push("AR warning");
  if (state.activeMetaFilters.aspectState === "missing") metaMissing.push("AR warning");
  if (state.activeMetaFilters.maskState === "has") metaHas.push("mask sidecar");
  if (state.activeMetaFilters.maskState === "missing") metaMissing.push("mask sidecar");
  if (state.activeMetaFilters.captionState === "has") metaHas.push("TXT file");
  if (state.activeMetaFilters.captionState === "missing") metaMissing.push("TXT file");

  if (metaHas.length > 0) {
    if (lines.length > 0) lines.push("");
    lines.push("Other has:");
    metaHas.forEach((label, index) => {
      if (index > 0) {
        lines.push("  and");
      }
      lines.push(`  ${label}`);
    });
  }

  if (metaMissing.length > 0) {
    if (lines.length > 0) lines.push("");
    lines.push("Other has not:");
    metaMissing.forEach((label, index) => {
      if (index > 0) {
        lines.push("  or");
      }
      lines.push(`  ${label}`);
    });
  }

  return lines.join("\n").trim();
}

function getActiveMetaFilterCount() {
  let count = 0;
  if (state.activeMetaFilters.aspectState !== "any") count += 1;
  if (state.activeMetaFilters.maskState !== "any") count += 1;
  if (state.activeMetaFilters.captionState !== "any") count += 1;
  return count;
}

function getActiveFilterCount() {
  return getActiveSentenceFilterEntries().length + getActiveMetaFilterCount();
}

function hasAnyActiveFilters() {
  return getActiveFilterCount() > 0;
}

function getActiveSentenceFilterKey() {
  return JSON.stringify(
    getActiveSentenceFilterEntries()
      .map(([sentence, mode]) => [sentence, mode])
      .sort((a, b) => a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]))
  );
}

function canApplyActiveSentenceFilters() {
  const filters = getActiveSentenceFilterEntries();
  if (filters.length === 0) return false;
  if (state.filterCaptionCacheKey !== getActiveSentenceFilterKey()) return false;
  return state.images.every(img => !!state.captionCache[img.path]);
}

function getEnabledSentencesForPath(path) {
  const enabledSentences = state.captionCache[path]?.enabled_sentences;
  return orderEnabledSentences(enabledSentences);
}

function imageMatchesMetaFilters(image) {
  if (!image) return false;
  if (state.activeMetaFilters.aspectState === "has" && imageConformsToAspectRatios(image)) {
    return false;
  }
  if (state.activeMetaFilters.aspectState === "missing" && !imageConformsToAspectRatios(image)) {
    return false;
  }
  if (state.activeMetaFilters.maskState === "has" && !image.has_mask) {
    return false;
  }
  if (state.activeMetaFilters.maskState === "missing" && image.has_mask) {
    return false;
  }
  if (state.activeMetaFilters.captionState === "has" && !image.has_caption) {
    return false;
  }
  if (state.activeMetaFilters.captionState === "missing" && image.has_caption) {
    return false;
  }
  return true;
}

function imageMatchesActiveFilters(image) {
  const filters = getActiveSentenceFilterEntries();
  if (!imageMatchesMetaFilters(image)) return false;
  if (filters.length === 0 || !canApplyActiveSentenceFilters()) return true;
  const clauses = getActiveSentenceFilterClauses(filters);
  const enabled = new Set(getEnabledSentencesForPath(image?.path));
  return clauses.every((clause) => {
    if (clause.type === "group") {
      return clause.sentences.some(sentence => enabled.has(sentence));
    }
    return clause.mode === "missing"
      ? !enabled.has(clause.sentence)
      : enabled.has(clause.sentence);
  });
}

function getVisibleImageEntries() {
  return state.images
    .map((img, index) => ({ img, index }))
    .filter(({ img }) => imageMatchesActiveFilters(img));
}

function getVisibleImageIndexByPath(path) {
  if (!path) return -1;
  return getVisibleImageEntries().findIndex(entry => entry.img.path === path);
}

async function fetchCaptionsBulk(paths, sentences = getAllConfiguredSentences()) {
  if (!Array.isArray(paths) || paths.length === 0) return {};
  const resp = await fetch("/api/captions/bulk", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      paths,
      captions: sentences,
    }),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data.detail || "Failed to load captions");
  }
  return data;
}

async function fetchMetadataBulk(paths) {
  if (!Array.isArray(paths) || paths.length === 0) return {};
  const resp = await fetch("/api/media/meta/bulk", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ paths }),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data.detail || "Failed to load metadata");
  }
  return data;
}

async function ensureCaptionCacheLoadedForFiltering() {
  const filters = getActiveSentenceFilters();
  if (filters.length === 0 || state.images.length === 0) return;

  const filterKey = getActiveSentenceFilterKey();
  const needsReload = state.filterCaptionCacheKey !== filterKey || state.images.some(img => !state.captionCache[img.path]);
  if (!needsReload) return;
  if (state.filterLoadingPromise) return state.filterLoadingPromise;

  state.filterLoadingPromise = (async () => {
    const data = await fetchCaptionsBulk(state.images.map(img => img.path));
    for (const [path, caption] of Object.entries(data || {})) {
      state.captionCache[path] = normalizeCaptionCacheEntry(caption);
    }
    state.filterCaptionCacheKey = filterKey;
  })().finally(() => {
    state.filterLoadingPromise = null;
  });

  return state.filterLoadingPromise;
}

function refreshGridForActiveFilters() {
  if (hasAnyActiveFilters()) {
    renderGrid();
  } else {
    updateFileCountDisplay();
    renderFilterActions();
  }
}

function renderFilterActions() {
  const hasFilters = hasAnyActiveFilters();
  const filterCount = getActiveFilterCount();
  const filterSummary = buildActiveFilterSummary();
  clearFiltersBtn.classList.toggle("active", hasFilters);
  clearFiltersBtn.setAttribute("aria-pressed", hasFilters ? "true" : "false");
  applyTriStateFilterButton(filterArBtn, "AR", "No AR", state.activeMetaFilters.aspectState, {
    any: "Aspect ratio filter disabled",
    has: "Showing only media files with the AR warning badge",
    missing: "Showing only media files without the AR warning badge",
  });
  applyTriStateFilterButton(filterMaskBtn, "M", "No M", state.activeMetaFilters.maskState, {
    any: "Mask filter disabled",
    has: "Showing only media files with a mask sidecar",
    missing: "Showing only media files without a mask sidecar",
  });
  applyTriStateFilterButton(filterTxtBtn, "TXT", "No TXT", state.activeMetaFilters.captionState, {
    any: "TXT file filter disabled",
    has: "Showing only media files that have a txt caption file",
    missing: "Showing only media files without a txt caption file",
  });
  clearFiltersBtn.title = hasFilters
    ? `Clear ${filterCount} active filter${filterCount === 1 ? "" : "s"}\n\n${filterSummary}`
    : "No active filters";
  clearFiltersBtn.setAttribute("aria-label", hasFilters
    ? `Clear ${filterCount} active filter${filterCount === 1 ? "" : "s"}. ${filterSummary.replace(/\s+/g, " ").trim()}`
    : "No active filters"
  );
}

function clearSentenceFilters() {
  if (!hasAnyActiveFilters()) return;
  const preferredPath = state.previewPath || state.lastClickedPath || [...state.selectedPaths][0] || null;
  const previousScrollTop = fileGridContainer.scrollTop;
  state.activeSentenceFilters.clear();
  state.activeMetaFilters.aspectState = "any";
  state.activeMetaFilters.maskState = "any";
  state.activeMetaFilters.captionState = "any";
  state.filterCaptionCacheKey = "";
  renderGrid({ preservePath: preferredPath, preserveScrollTop: previousScrollTop });
  renderSentences();
  statusBar.textContent = "Filters cleared";
}

function getNextTriStateFilterValue(currentValue) {
  if (currentValue === "any") return "has";
  if (currentValue === "has") return "missing";
  return "any";
}

function applyTriStateFilterButton(button, positiveLabel, negativeLabel, stateValue, labels) {
  if (!button) return;
  button.textContent = stateValue === "missing" ? negativeLabel : positiveLabel;
  button.classList.toggle("active", stateValue === "has");
  button.classList.toggle("missing", stateValue === "missing");
  button.setAttribute("aria-pressed", stateValue === "any" ? "false" : "true");
  button.title = labels[stateValue] || positiveLabel;
  button.setAttribute("aria-label", labels[stateValue] || positiveLabel);
}

function applyFilterUiUpdate(statusMessage) {
  const preferredPath = state.previewPath || state.lastClickedPath || [...state.selectedPaths][0] || null;
  const previousScrollTop = fileGridContainer.scrollTop;
  renderGrid({ preservePath: preferredPath, preserveScrollTop: previousScrollTop });
  renderSentences();
  if (statusMessage) {
    statusBar.textContent = statusMessage;
  }
}

function toggleAspectFilter() {
  state.activeMetaFilters.aspectState = getNextTriStateFilterValue(state.activeMetaFilters.aspectState);
  let label = "Filters cleared";
  if (state.activeMetaFilters.aspectState === "has") {
    label = "Filtered by wrong aspect ratio";
  } else if (state.activeMetaFilters.aspectState === "missing") {
    label = "Filtered by valid aspect ratio";
  } else if (hasAnyActiveFilters()) {
    const count = getActiveFilterCount();
    label = `Filtered by ${count} filter${count === 1 ? "" : "s"}`;
  }
  applyFilterUiUpdate(label);
}

function toggleCaptionPresenceFilter() {
  state.activeMetaFilters.captionState = getNextTriStateFilterValue(state.activeMetaFilters.captionState);
  let label = "Filters cleared";
  if (state.activeMetaFilters.captionState === "has") {
    label = "Filtered by has txt file";
  } else if (state.activeMetaFilters.captionState === "missing") {
    label = "Filtered by missing txt file";
  } else if (hasAnyActiveFilters()) {
    const count = getActiveFilterCount();
    label = `Filtered by ${count} filter${count === 1 ? "" : "s"}`;
  }
  applyFilterUiUpdate(label);
}

function toggleMaskPresenceFilter() {
  state.activeMetaFilters.maskState = getNextTriStateFilterValue(state.activeMetaFilters.maskState);
  let label = "Filters cleared";
  if (state.activeMetaFilters.maskState === "has") {
    label = "Filtered by has mask";
  } else if (state.activeMetaFilters.maskState === "missing") {
    label = "Filtered by missing mask";
  } else if (hasAnyActiveFilters()) {
    const count = getActiveFilterCount();
    label = `Filtered by ${count} filter${count === 1 ? "" : "s"}`;
  }
  applyFilterUiUpdate(label);
}

async function toggleSentenceFilter(sentence) {
  const preferredPath = state.previewPath || state.lastClickedPath || [...state.selectedPaths][0] || null;
  const previousScrollTop = fileGridContainer.scrollTop;
  const previousFilters = new Map(state.activeSentenceFilters);
  const isGroupSentence = !!findGroupForSentence(sentence);
  const currentMode = getSentenceFilterMode(sentence);
  if (isGroupSentence) {
    if (currentMode === "has") {
      state.activeSentenceFilters.delete(sentence);
    } else {
      state.activeSentenceFilters.set(sentence, "has");
    }
  } else {
    if (currentMode === "off") {
      state.activeSentenceFilters.set(sentence, "has");
    } else if (currentMode === "has") {
      state.activeSentenceFilters.set(sentence, "missing");
    } else {
      state.activeSentenceFilters.delete(sentence);
    }
  }

  try {
    if (state.activeSentenceFilters.size > 0) {
      statusBar.textContent = `Applying ${state.activeSentenceFilters.size} filter${state.activeSentenceFilters.size === 1 ? "" : "s"}...`;
      await ensureCaptionCacheLoadedForFiltering();
      statusBar.textContent = `Filtered by ${state.activeSentenceFilters.size} caption${state.activeSentenceFilters.size === 1 ? "" : "s"}`;
    } else {
      statusBar.textContent = "Filter cleared";
    }
  } catch (err) {
    state.activeSentenceFilters = previousFilters;
    statusBar.textContent = `Filter error: ${err.message}`;
  }

  renderGrid({ preservePath: preferredPath, preserveScrollTop: previousScrollTop });
  renderSentences();
}

function setSectionCollapsed(section, collapsed) {
  state.collapsedSections[section._uiId] = !!collapsed;
}

function isSectionCollapsed(section) {
  return !!state.collapsedSections[section._uiId];
}

function setGroupCollapsed(group, collapsed) {
  state.collapsedGroups[group._uiId] = !!collapsed;
}

function isGroupCollapsed(group) {
  return !!state.collapsedGroups[group._uiId];
}

function findSentenceLocation(sentence) {
  for (const [sectionIndex, section] of state.sections.entries()) {
    if ((section.sentences || []).includes(sentence)) {
      return { sectionIndex, groupIndex: null, section, group: null };
    }
    for (const [groupIndex, group] of (section.groups || []).entries()) {
      if ((group.sentences || []).includes(sentence)) {
        return { sectionIndex, groupIndex, section, group };
      }
    }
  }
  return null;
}

function jumpToSentenceInList(sentence) {
  const location = findSentenceLocation(sentence);
  if (!location) return;

  let rerendered = false;
  if (location.section && isSectionCollapsed(location.section)) {
    setSectionCollapsed(location.section, false);
    rerendered = true;
  }
  if (location.group && isGroupCollapsed(location.group)) {
    setGroupCollapsed(location.group, false);
    rerendered = true;
  }
  if (rerendered) {
    renderSentences();
  }

  const target = sentenceListElements.get(sentence);
  if (!target) return;

  target.scrollIntoView({ behavior: "smooth", block: "center" });
  target.classList.remove("jump-highlight");
  void target.offsetWidth;
  target.classList.add("jump-highlight");
  window.setTimeout(() => target.classList.remove("jump-highlight"), 1200);
}

function hasActiveInlineEditor() {
  const editor = state.ui.activeInlineEditor;
  if (!editor) return false;
  if (!editor.isConnected || editor.contentEditable !== "true") {
    state.ui.activeInlineEditor = null;
    return false;
  }
  return true;
}

function flushQueuedUiRenders(options = {}) {
  const { force = false } = options;
  if (state.ui.renderFrameId) {
    window.cancelAnimationFrame(state.ui.renderFrameId);
    state.ui.renderFrameId = 0;
  }

  if (state.ui.pendingSentenceRender) {
    return renderSentences({ force, includePreview: true });
  }

  if (state.ui.pendingPreviewRender) {
    state.ui.pendingPreviewRender = false;
    renderPreviewCaptionOverlay();
    return true;
  }

  return false;
}

function scheduleUiRender(options = {}) {
  const { sentences = false, preview = false } = options;
  if (sentences) state.ui.pendingSentenceRender = true;
  if (preview) state.ui.pendingPreviewRender = true;
  if (state.ui.renderFrameId) return;
  state.ui.renderFrameId = window.requestAnimationFrame(() => {
    state.ui.renderFrameId = 0;
    flushQueuedUiRenders();
  });
}

function isMaskEditAvailable() {
  return state.selectedPaths.size === 1
    && !!state.previewPath
    && (state.previewMediaType === "image" || state.previewMediaType === "video")
    && !!imgNatW
    && !!imgNatH;
}

function isImageEditAvailable() {
  return state.selectedPaths.size === 1
    && !!state.previewPath
    && state.previewMediaType === "image"
    && !!imgNatW
    && !!imgNatH;
}

function isVideoMaskEditAvailable() {
  return state.selectedPaths.size === 1
    && !!state.previewPath
    && state.previewMediaType === "video"
    && !!imgNatW
    && !!imgNatH;
}

function isMaskEditorVisible() {
  return !!state.maskEditor.active;
}

function isMaskEditorMaskMode() {
  return state.maskEditor.active && state.maskEditor.mode === "mask";
}

function isMaskEditorImageMode() {
  return state.maskEditor.active && state.maskEditor.mode === "image";
}

function getActiveEditCanvas() {
  return isMaskEditorImageMode() ? previewImageEditCanvas : previewMaskCanvas;
}

function hexToRgb(hexColor) {
  const normalized = String(hexColor || "").trim().replace(/^#/, "");
  if (normalized.length !== 6) {
    return { r: 255, g: 90, b: 90 };
  }
  const value = Number.parseInt(normalized, 16);
  if (!Number.isFinite(value)) {
    return { r: 255, g: 90, b: 90 };
  }
  return {
    r: (value >> 16) & 255,
    g: (value >> 8) & 255,
    b: value & 255,
  };
}

function getEffectiveVideoMaskFps(path = state.previewPath) {
  const metaFps = Number(getCurrentVideoMeta(path)?.fps || 0);
  if (metaFps > 0) return metaFps;
  const profileFps = Number(getSelectedVideoTrainingProfileFromState()?.fps || 0);
  return Math.max(1, profileFps || 24);
}

function getVideoMaskKeyframes(path = state.previewPath) {
  const keyframes = getCurrentVideoMeta(path)?.mask_keyframes;
  if (!Array.isArray(keyframes)) return [];
  return [...new Set(keyframes.map((value) => Math.max(0, Number.parseInt(value, 10) || 0)))].sort((a, b) => a - b);
}

function setVideoMaskKeyframes(path, keyframes) {
  if (!path) return;
  if (!state.videoMeta[path]) {
    state.videoMeta[path] = {};
  }
  state.videoMeta[path].mask_keyframes = getVideoMaskKeyframesFromValues(keyframes);
}

function getVideoMaskKeyframesFromValues(keyframes) {
  if (!Array.isArray(keyframes)) return [];
  return [...new Set(keyframes.map((value) => {
    if (value && typeof value === "object") {
      return Math.max(0, Number.parseInt(value.frame_index, 10) || 0);
    }
    return Math.max(0, Number.parseInt(value, 10) || 0);
  }))].sort((a, b) => a - b);
}

function getResolvedVideoMaskKeyframeForFrame(path, requestedFrameIndex, options = {}) {
  const { fallbackToCurrent = true } = options;
  const normalizedRequestedFrameIndex = Math.max(0, Number.parseInt(requestedFrameIndex, 10) || 0);
  let resolvedFrameIndex = null;
  for (const keyframe of getVideoMaskKeyframes(path)) {
    if (keyframe > normalizedRequestedFrameIndex) {
      break;
    }
    resolvedFrameIndex = keyframe;
  }
  if (resolvedFrameIndex == null && fallbackToCurrent && state.maskEditor.mediaType === "video" && state.maskEditor.path === path) {
    return state.maskEditor.frameIndex;
  }
  return resolvedFrameIndex;
}

function getCurrentVideoMaskFrameIndex(path = state.previewPath) {
  const currentTime = Math.max(0, Number(previewVideo.currentTime || 0));
  return Math.max(0, Math.floor((currentTime * getEffectiveVideoMaskFps(path)) + 1e-6));
}

function formatVideoMaskFrameHint(frameIndex, path = state.previewPath) {
  const normalizedFrameIndex = Math.max(0, Number(frameIndex || 0));
  const fps = getEffectiveVideoMaskFps(path);
  const timeSeconds = normalizedFrameIndex / Math.max(1, fps);
  return `frame ${normalizedFrameIndex} (${formatDurationSeconds(timeSeconds)})`;
}

function revokeMaskEditorVideoSnapshot() {
  if (state.maskEditor.videoSnapshotUrl) {
    URL.revokeObjectURL(state.maskEditor.videoSnapshotUrl);
    state.maskEditor.videoSnapshotUrl = null;
  }
}

async function captureCurrentPreviewVideoFrameSnapshot(path = state.previewPath) {
  if (!path || state.previewMediaType !== "video") {
    throw new Error("Video preview is not active");
  }
  if (previewVideo.readyState < 2) {
    throw new Error("Video frame is not ready yet");
  }

  const frameWidth = Math.max(1, Number(previewVideo.videoWidth || getCurrentVideoMeta(path)?.width || imgNatW || 1));
  const frameHeight = Math.max(1, Number(previewVideo.videoHeight || getCurrentVideoMeta(path)?.height || imgNatH || 1));
  const snapshotCanvas = document.createElement("canvas");
  snapshotCanvas.width = frameWidth;
  snapshotCanvas.height = frameHeight;
  const snapshotContext = snapshotCanvas.getContext("2d");
  snapshotContext.drawImage(previewVideo, 0, 0, frameWidth, frameHeight);

  const snapshotBlob = await new Promise((resolve, reject) => {
    snapshotCanvas.toBlob((blob) => {
      if (blob) {
        resolve(blob);
        return;
      }
      reject(new Error("Failed to capture the current video frame"));
    }, "image/jpeg", 0.92);
  });

  const snapshotUrl = URL.createObjectURL(snapshotBlob);
  await new Promise((resolve, reject) => {
    previewImg.onload = () => resolve();
    previewImg.onerror = () => reject(new Error("Failed to load the captured video frame"));
    previewImg.src = snapshotUrl;
  });
  revokeMaskEditorVideoSnapshot();
  state.maskEditor.videoSnapshotUrl = snapshotUrl;
  return { width: frameWidth, height: frameHeight, url: snapshotUrl };
}

function clearMaskCursor() {
  state.maskEditor.cursorClientX = null;
  state.maskEditor.cursorClientY = null;
  if (maskCursorValue) {
    maskCursorValue.textContent = "";
    maskCursorValue.style.fontSize = "";
  }
  maskCursor.classList.remove("visible");
}

function getMaskValuePercentAtMaskPoint(maskX, maskY) {
  if (!isMaskEditorMaskMode() || !previewMaskCanvas.width || !previewMaskCanvas.height) {
    return null;
  }
  const ctx = previewMaskCanvas.getContext("2d", { willReadFrequently: true });
  if (!ctx) {
    return null;
  }
  const sampleX = clamp(Math.floor(maskX), 0, Math.max(0, previewMaskCanvas.width - 1));
  const sampleY = clamp(Math.floor(maskY), 0, Math.max(0, previewMaskCanvas.height - 1));
  const pixel = ctx.getImageData(sampleX, sampleY, 1, 1).data;
  if (!pixel || !pixel.length) {
    return null;
  }
  return clamp((pixel[0] / 255) * 100, 0, 100);
}

function refreshMaskCursorValue() {
  if (!maskCursorValue || !maskCursor.classList.contains("visible") || !isMaskEditorMaskMode()) {
    return;
  }
  const clientX = Number(state.maskEditor.cursorClientX);
  const clientY = Number(state.maskEditor.cursorClientY);
  if (!Number.isFinite(clientX) || !Number.isFinite(clientY)) {
    maskCursorValue.textContent = "";
    return;
  }
  const imagePoint = screenToImage(clientX, clientY);
  const maskPoint = previewPointToMaskPoint(imagePoint);
  const valuePercent = getMaskValuePercentAtMaskPoint(maskPoint.x, maskPoint.y);
  maskCursorValue.textContent = Number.isFinite(valuePercent) ? `${Math.round(valuePercent)}%` : "";
}

function isMaskSignalProbeMode() {
  return isMaskEditorMaskMode() && !!state.maskEditor.latentPreviewEnabled && !!state.maskEditor.signalProbeMode;
}

function getMaskSignalProbeRect() {
  const rect = state.maskEditor.signalProbeRect;
  const maxWidth = Math.max(1, Number(state.maskEditor.imageWidth || previewMaskCanvas.width || 1));
  const maxHeight = Math.max(1, Number(state.maskEditor.imageHeight || previewMaskCanvas.height || 1));
  if (!rect) return null;
  const left = clamp(Number(rect.x || 0), 0, Math.max(0, maxWidth - 1));
  const top = clamp(Number(rect.y || 0), 0, Math.max(0, maxHeight - 1));
  const right = clamp(left + Math.max(1, Number(rect.w || 0)), Math.min(maxWidth, left + 1), maxWidth);
  const bottom = clamp(top + Math.max(1, Number(rect.h || 0)), Math.min(maxHeight, top + 1), maxHeight);
  const width = Math.max(0, Math.round(right - left));
  const height = Math.max(0, Math.round(bottom - top));
  if (width <= 0 || height <= 0) return null;
  return {
    x: Math.round(left),
    y: Math.round(top),
    w: width,
    h: height,
  };
}

function hasMaskSignalProbeRect() {
  return !!getMaskSignalProbeRect();
}

function getMaskSignalLatentRect() {
  const rect = getMaskSignalProbeRect();
  const latentWidth = Math.max(0, Number(state.maskEditor.latentSignalWidth || 0));
  const latentHeight = Math.max(0, Number(state.maskEditor.latentSignalHeight || 0));
  const maskWidth = Math.max(1, Number(previewMaskCanvas.width || state.maskEditor.imageWidth || 1));
  const maskHeight = Math.max(1, Number(previewMaskCanvas.height || state.maskEditor.imageHeight || 1));
  if (!rect || !latentWidth || !latentHeight) return null;
  const scaleX = latentWidth / maskWidth;
  const scaleY = latentHeight / maskHeight;
  const left = clamp(rect.x * scaleX, 0, latentWidth);
  const top = clamp(rect.y * scaleY, 0, latentHeight);
  const right = clamp((rect.x + rect.w) * scaleX, left, latentWidth);
  const bottom = clamp((rect.y + rect.h) * scaleY, top, latentHeight);
  if (right <= left || bottom <= top) return null;
  return { left, top, right, bottom };
}

function getLatentSignalIntegralIndex(x, y, width = state.maskEditor.latentSignalWidth) {
  const stride = Math.max(1, Number(width || 0)) + 1;
  return (y * stride) + x;
}

function queryLatentSignalIntegral(left, top, right, bottom) {
  const integral = state.maskEditor.latentSignalIntegral;
  if (!integral || right <= left || bottom <= top) return 0;
  return integral[getLatentSignalIntegralIndex(right, bottom)]
    - integral[getLatentSignalIntegralIndex(left, bottom)]
    - integral[getLatentSignalIntegralIndex(right, top)]
    + integral[getLatentSignalIntegralIndex(left, top)];
}

function getLatentSignalCellValue(x, y) {
  const values = state.maskEditor.latentSignalValues;
  const width = Math.max(0, Number(state.maskEditor.latentSignalWidth || 0));
  const height = Math.max(0, Number(state.maskEditor.latentSignalHeight || 0));
  if (!values || x < 0 || y < 0 || x >= width || y >= height) {
    return 0;
  }
  return values[(y * width) + x] || 0;
}

function sumLatentSignalRect(left, top, right, bottom) {
  const width = Math.max(0, Number(state.maskEditor.latentSignalWidth || 0));
  const height = Math.max(0, Number(state.maskEditor.latentSignalHeight || 0));
  if (!width || !height) return 0;
  const clampedLeft = clamp(left, 0, width);
  const clampedTop = clamp(top, 0, height);
  const clampedRight = clamp(right, clampedLeft, width);
  const clampedBottom = clamp(bottom, clampedTop, height);
  if (clampedRight <= clampedLeft || clampedBottom <= clampedTop) {
    return 0;
  }

  const fullLeft = Math.ceil(clampedLeft);
  const fullTop = Math.ceil(clampedTop);
  const fullRight = Math.floor(clampedRight);
  const fullBottom = Math.floor(clampedBottom);
  let sum = queryLatentSignalIntegral(fullLeft, fullTop, fullRight, fullBottom);

  const startX = Math.max(0, Math.floor(clampedLeft));
  const endX = Math.min(width - 1, Math.ceil(clampedRight) - 1);
  const startY = Math.max(0, Math.floor(clampedTop));
  const endY = Math.min(height - 1, Math.ceil(clampedBottom) - 1);

  for (let y = startY; y <= endY; y += 1) {
    const overlapY = Math.min(clampedBottom, y + 1) - Math.max(clampedTop, y);
    if (overlapY <= 0) continue;
    const yIsInterior = y >= fullTop && y < fullBottom;
    if (!yIsInterior) {
      for (let x = startX; x <= endX; x += 1) {
        const overlapX = Math.min(clampedRight, x + 1) - Math.max(clampedLeft, x);
        if (overlapX <= 0) continue;
        sum += getLatentSignalCellValue(x, y) * overlapX * overlapY;
      }
      continue;
    }

    const leftBoundaryX = Math.floor(clampedLeft);
    if (leftBoundaryX >= startX && leftBoundaryX <= endX && leftBoundaryX < fullLeft) {
      const overlapX = Math.min(clampedRight, leftBoundaryX + 1) - Math.max(clampedLeft, leftBoundaryX);
      if (overlapX > 0) {
        sum += getLatentSignalCellValue(leftBoundaryX, y) * overlapX * overlapY;
      }
    }

    const rightBoundaryX = Math.ceil(clampedRight) - 1;
    if (rightBoundaryX >= startX && rightBoundaryX <= endX && rightBoundaryX >= fullRight && rightBoundaryX !== leftBoundaryX) {
      const overlapX = Math.min(clampedRight, rightBoundaryX + 1) - Math.max(clampedLeft, rightBoundaryX);
      if (overlapX > 0) {
        sum += getLatentSignalCellValue(rightBoundaryX, y) * overlapX * overlapY;
      }
    }
  }

  return sum;
}

function updateMaskLatentSignalBuffer(sourceCanvas) {
  const canvas = sourceCanvas || previewLatentMaskCanvas;
  if (!canvas || !canvas.width || !canvas.height) {
    state.maskEditor.latentSignalValues = null;
    state.maskEditor.latentSignalIntegral = null;
    state.maskEditor.latentSignalWidth = 0;
    state.maskEditor.latentSignalHeight = 0;
    state.maskEditor.latentSignalTotalValue = 0;
    return;
  }

  const width = canvas.width;
  const height = canvas.height;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const { data } = ctx.getImageData(0, 0, width, height);
  const values = new Float32Array(width * height);
  const integral = new Float64Array((width + 1) * (height + 1));
  let totalValue = 0;

  for (let y = 0; y < height; y += 1) {
    let rowSum = 0;
    for (let x = 0; x < width; x += 1) {
      const value = data[((y * width) + x) * 4] || 0;
      values[(y * width) + x] = value;
      rowSum += value;
      totalValue += value;
      integral[getLatentSignalIntegralIndex(x + 1, y + 1, width)] = integral[getLatentSignalIntegralIndex(x + 1, y, width)] + rowSum;
    }
  }

  state.maskEditor.latentSignalValues = values;
  state.maskEditor.latentSignalIntegral = integral;
  state.maskEditor.latentSignalWidth = width;
  state.maskEditor.latentSignalHeight = height;
  state.maskEditor.latentSignalTotalValue = totalValue;
}

function renderMaskSignalProbeOverlay() {
  const rect = getMaskSignalProbeRect();
  const signalPercent = clamp(Number(state.maskEditor.signalProbePercent || 0), 0, 100);
  const visible = isMaskSignalProbeMode() && !!rect && !!imgNatW && !!imgNatH;
  maskSignalProbeRect.classList.toggle("visible", visible);
  maskSignalProbeRect.classList.toggle("is-good", visible && signalPercent >= 30);
  maskSignalProbeRect.classList.toggle("is-low", visible && signalPercent < 30);
  if (!visible) {
    return;
  }
  const previewScaleX = Math.max(0.0001, Number(state.maskEditor.previewScaleX || 1));
  const previewScaleY = Math.max(0.0001, Number(state.maskEditor.previewScaleY || 1));
  const left = panX + ((rect.x / previewScaleX) * zoomLevel);
  const top = panY + ((rect.y / previewScaleY) * zoomLevel);
  const width = Math.max(1, (rect.w / previewScaleX) * zoomLevel);
  const height = Math.max(1, (rect.h / previewScaleY) * zoomLevel);
  maskSignalProbeRect.style.left = `${left}px`;
  maskSignalProbeRect.style.top = `${top}px`;
  maskSignalProbeRect.style.width = `${width}px`;
  maskSignalProbeRect.style.height = `${height}px`;
  maskSignalProbeRectLabel.textContent = `${signalPercent.toFixed(1)}% of signal`;
}

function renderMaskSignalProbeUi() {
  const active = isMaskEditorMaskMode();
  const showControls = active && !!state.maskEditor.latentPreviewEnabled;
  const interactive = showControls && !state.maskEditor.loading && !state.maskEditor.saving;
  const hasRect = hasMaskSignalProbeRect();
  const signalPercent = clamp(Number(state.maskEditor.signalProbePercent || 0), 0, 100);
  const areaPercent = clamp(Number(state.maskEditor.signalProbeAreaPercent || 0), 0, 100);
  maskSignalProbeControls.classList.toggle("visible", showControls);
  maskSignalProbeBtn.disabled = !interactive;
  maskSignalProbeBtn.setAttribute("aria-pressed", showControls && state.maskEditor.signalProbeMode ? "true" : "false");
  maskSignalProbeBtn.title = state.maskEditor.signalProbeMode
    ? "Probe mode is on. Right-drag in the latent preview to draw or redraw the signal rectangle."
    : "Enable probe mode, then right-drag in the latent preview to measure signal for a detail area.";
  if (!hasRect) {
    maskSignalProbeLabel.textContent = showControls && state.maskEditor.signalProbeMode
      ? "Right-drag an area to measure signal"
      : "No probe area";
  } else {
    maskSignalProbeLabel.textContent = `Signal ${signalPercent.toFixed(1)}% Area ${areaPercent.toFixed(1)}%`;
  }
  maskSignalProbeLabel.title = hasRect
    ? "Signal inside the selected rectangle against the current total latent signal. Aim for roughly 30% or more."
    : "Enable probe mode and right-drag in the latent preview to measure the selected detail area.";
  maskSignalProbeLabel.classList.toggle("has-value", hasRect);
  maskSignalProbeLabel.classList.toggle("is-good", hasRect && signalPercent >= 30);
  maskSignalProbeLabel.classList.toggle("is-low", hasRect && signalPercent < 30);
  renderMaskSignalProbeOverlay();
}

function updateMaskSignalProbeStats() {
  const rect = getMaskSignalProbeRect();
  const latentRect = getMaskSignalLatentRect();
  const totalSignalValue = Math.max(0, Number(state.maskEditor.latentSignalTotalValue || 0));
  if (!isMaskEditorMaskMode() || !rect || !latentRect || !previewMaskCanvas.width || !previewMaskCanvas.height) {
    state.maskEditor.signalProbePercent = 0;
    state.maskEditor.signalProbeAreaPercent = 0;
    renderMaskSignalProbeUi();
    return;
  }
  const rectSignalValue = sumLatentSignalRect(latentRect.left, latentRect.top, latentRect.right, latentRect.bottom);
  const latentArea = Math.max(1, Number(state.maskEditor.latentSignalWidth || 0) * Number(state.maskEditor.latentSignalHeight || 0));
  const rectArea = Math.max(0, (latentRect.right - latentRect.left) * (latentRect.bottom - latentRect.top));
  state.maskEditor.signalProbePercent = totalSignalValue > 0
    ? clamp((rectSignalValue / totalSignalValue) * 100, 0, 100)
    : 0;
  state.maskEditor.signalProbeAreaPercent = clamp((rectArea / latentArea) * 100, 0, 100);
  renderMaskSignalProbeUi();
}

function beginMaskSignalProbeDrag(event) {
  const maxWidth = Math.max(1, Number(state.maskEditor.imageWidth || previewMaskCanvas.width || 1));
  const maxHeight = Math.max(1, Number(state.maskEditor.imageHeight || previewMaskCanvas.height || 1));
  const anchorPoint = previewPointToMaskPoint(screenToImage(event.clientX, event.clientY));
  const anchor = {
    x: clamp(anchorPoint.x, 0, Math.max(0, maxWidth - 1)),
    y: clamp(anchorPoint.y, 0, Math.max(0, maxHeight - 1)),
  };
  state.maskEditor.signalProbeDragging = true;
  state.maskEditor.signalProbeAnchor = anchor;
  state.maskEditor.signalProbeRect = {
    x: Math.round(anchor.x),
    y: Math.round(anchor.y),
    w: 1,
    h: 1,
  };
  updateMaskSignalProbeStats();
}

function updateMaskSignalProbeDrag(clientX, clientY) {
  if (!state.maskEditor.signalProbeDragging || !state.maskEditor.signalProbeAnchor) return;
  const point = previewPointToMaskPoint(screenToImage(clientX, clientY));
  const maxWidth = Math.max(1, Number(state.maskEditor.imageWidth || previewMaskCanvas.width || 1));
  const maxHeight = Math.max(1, Number(state.maskEditor.imageHeight || previewMaskCanvas.height || 1));
  const left = clamp(Math.min(state.maskEditor.signalProbeAnchor.x, point.x), 0, maxWidth);
  const top = clamp(Math.min(state.maskEditor.signalProbeAnchor.y, point.y), 0, maxHeight);
  const right = clamp(Math.max(state.maskEditor.signalProbeAnchor.x, point.x), 0, maxWidth);
  const bottom = clamp(Math.max(state.maskEditor.signalProbeAnchor.y, point.y), 0, maxHeight);
  state.maskEditor.signalProbeRect = {
    x: Math.round(left),
    y: Math.round(top),
    w: Math.max(1, Math.round(right - left)),
    h: Math.max(1, Math.round(bottom - top)),
  };
  updateMaskSignalProbeStats();
}

function stopMaskSignalProbeDrag() {
  if (!state.maskEditor.signalProbeDragging) return;
  state.maskEditor.signalProbeDragging = false;
  state.maskEditor.signalProbeAnchor = null;
  renderMaskSignalProbeUi();
  if (hasMaskSignalProbeRect()) {
    statusBar.textContent = `Signal rectangle ${state.maskEditor.signalProbePercent.toFixed(1)}%`;
  }
}

function toggleMaskSignalProbeMode() {
  if (!isMaskEditorMaskMode()) return;
  state.maskEditor.signalProbeMode = !state.maskEditor.signalProbeMode;
  state.maskEditor.signalProbeDragging = false;
  state.maskEditor.signalProbeAnchor = null;
  clearMaskCursor();
  renderMaskEditorUi();
  statusBar.textContent = state.maskEditor.signalProbeMode
    ? "Signal rectangle mode enabled. Right-drag to measure a detail area."
    : (hasMaskSignalProbeRect()
      ? `Signal rectangle hidden. Live signal ${state.maskEditor.signalProbePercent.toFixed(1)}%`
      : "Signal rectangle mode disabled");
}

function getMaskBrushReferenceSize() {
  return Math.max(1, Number(state.maskEditor.imageWidth || 0), Number(state.maskEditor.imageHeight || 0), Number(imgNatW || 0), Number(imgNatH || 0));
}

function getMaskBrushDiameterMaskPx() {
  const brushSizePercent = clamp(Number(state.maskEditor.brushSizePercent || 6), 0.2, 100);
  return Math.max(1, getMaskBrushReferenceSize() * (brushSizePercent / 100));
}

function getMaskBrushDiameterPreviewPx() {
  const averageScale = ((state.maskEditor.previewScaleX || 1) + (state.maskEditor.previewScaleY || 1)) / 2;
  return Math.max(1, getMaskBrushDiameterMaskPx() / Math.max(averageScale, 0.0001));
}

function syncMaskEditorPreviewScaleFromCurrentImage() {
  if (!state.maskEditor.active || !imgNatW || !imgNatH || !state.maskEditor.imageWidth || !state.maskEditor.imageHeight) {
    return;
  }
  state.maskEditor.previewScaleX = state.maskEditor.imageWidth / Math.max(1, imgNatW);
  state.maskEditor.previewScaleY = state.maskEditor.imageHeight / Math.max(1, imgNatH);
}

function getMaskBrushInfluence(distanceFraction) {
  const coreFraction = clamp(Number(state.maskEditor.brushCore || 30), 0, 95) / 100;
  if (distanceFraction <= coreFraction) {
    return 1;
  }
  const normalizedDistance = clamp((distanceFraction - coreFraction) / Math.max(0.0001, 1 - coreFraction), 0, 1);
  const steepness = clamp(Number(state.maskEditor.brushSteepness || 8), 1, 32);
  const start = 1 / (1 + Math.exp(-steepness * 0.5));
  const end = 1 / (1 + Math.exp(steepness * 0.5));
  const raw = 1 / (1 + Math.exp(steepness * (normalizedDistance - 0.5)));
  return clamp((raw - end) / Math.max(0.0001, start - end), 0, 1);
}

function normalizeMaskLatentBaseWidthPresets(rawPresets) {
  const normalized = [...new Set((Array.isArray(rawPresets) ? rawPresets : [])
    .map((value) => Math.round(Number(value || 0)))
    .filter((value) => Number.isFinite(value) && value >= 64 && value <= 2048))]
    .sort((left, right) => left - right);
  return normalized.length ? normalized : [512, 768, 1024, 1280];
}

function getMaskLatentBaseWidthPresets() {
  state.maskLatentBaseWidthPresets = normalizeMaskLatentBaseWidthPresets(state.maskLatentBaseWidthPresets);
  return state.maskLatentBaseWidthPresets;
}

function getNearestMaskLatentBaseWidthPreset(value, presets = getMaskLatentBaseWidthPresets()) {
  const fallback = presets[0] || 512;
  const numericValue = Math.round(Number(value || 0)) || fallback;
  return presets.reduce((best, preset) => {
    const bestDistance = Math.abs(best - numericValue);
    const presetDistance = Math.abs(preset - numericValue);
    return presetDistance < bestDistance ? preset : best;
  }, fallback);
}

function getMaskLatentBaseWidthPresetIndex(value, presets = getMaskLatentBaseWidthPresets()) {
  const resolvedValue = getNearestMaskLatentBaseWidthPreset(value, presets);
  const index = presets.indexOf(resolvedValue);
  return index >= 0 ? index : 0;
}

function parseMaskLatentBaseWidthPresetsInput(rawValue) {
  const rawEntries = String(rawValue || "")
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);
  if (!rawEntries.length) {
    throw new Error("Mask latent base width presets must contain at least one number.");
  }
  const numericValues = rawEntries
    .map((value) => Math.round(Number(value)))
    .filter((value) => Number.isFinite(value));
  if (!numericValues.length) {
    throw new Error("Mask latent base width presets must contain valid numbers.");
  }
  return normalizeMaskLatentBaseWidthPresets(numericValues);
}

const MASK_LATENT_NOISE_PREVIEW_SEED = 1337;
const MASK_LATENT_NOISE_MAX_TIMESTEP = 999;
const MASK_LATENT_NOISE_BETA_START = 0.00085;
const MASK_LATENT_NOISE_BETA_END = 0.012;
const MASK_LATENT_NOISE_ALPHA_CUMPROD = (() => {
  const alphaCumprod = new Float32Array(MASK_LATENT_NOISE_MAX_TIMESTEP + 1);
  alphaCumprod[0] = 1;
  const betaStartSqrt = Math.sqrt(MASK_LATENT_NOISE_BETA_START);
  const betaEndSqrt = Math.sqrt(MASK_LATENT_NOISE_BETA_END);
  let cumulative = 1;
  for (let timestep = 1; timestep <= MASK_LATENT_NOISE_MAX_TIMESTEP; timestep += 1) {
    const fraction = timestep / MASK_LATENT_NOISE_MAX_TIMESTEP;
    const betaSqrt = betaStartSqrt + (betaEndSqrt - betaStartSqrt) * fraction;
    const beta = betaSqrt * betaSqrt;
    cumulative *= 1 - beta;
    alphaCumprod[timestep] = cumulative;
  }
  return alphaCumprod;
})();

function createMaskLatentNoiseRandom(seed = MASK_LATENT_NOISE_PREVIEW_SEED) {
  let value = seed >>> 0;
  return () => {
    value += 0x6D2B79F5;
    let mixed = value;
    mixed = Math.imul(mixed ^ (mixed >>> 15), mixed | 1);
    mixed ^= mixed + Math.imul(mixed ^ (mixed >>> 7), mixed | 61);
    return ((mixed ^ (mixed >>> 14)) >>> 0) / 4294967296;
  };
}

function getMaskLatentNoiseWeights(timestep) {
  const resolvedTimestep = Math.round(clamp(Number(timestep || 0), 0, MASK_LATENT_NOISE_MAX_TIMESTEP));
  const alphaCumprod = MASK_LATENT_NOISE_ALPHA_CUMPROD[resolvedTimestep] ?? 1;
  return {
    signalScale: Math.sqrt(alphaCumprod),
    noiseScale: Math.sqrt(Math.max(0, 1 - alphaCumprod)),
  };
}

function ensureMaskLatentNoiseBuffer(width, height) {
  if (!width || !height) {
    return null;
  }
  if (
    state.maskEditor.latentNoiseValues
    && state.maskEditor.latentNoiseWidth === width
    && state.maskEditor.latentNoiseHeight === height
  ) {
    return state.maskEditor.latentNoiseValues;
  }
  const noiseValues = new Float32Array(width * height * 3);
  const random = createMaskLatentNoiseRandom();
  for (let index = 0; index < noiseValues.length; index += 2) {
    const u1 = Math.max(random(), 1e-7);
    const u2 = random();
    const radius = Math.sqrt(-2 * Math.log(u1));
    const angle = 2 * Math.PI * u2;
    noiseValues[index] = radius * Math.cos(angle);
    if (index + 1 < noiseValues.length) {
      noiseValues[index + 1] = radius * Math.sin(angle);
    }
  }
  state.maskEditor.latentNoiseValues = noiseValues;
  state.maskEditor.latentNoiseWidth = width;
  state.maskEditor.latentNoiseHeight = height;
  return noiseValues;
}

function applyMaskLatentNoisePreview(ctx, width, height) {
  const timestep = Math.round(clamp(Number(state.maskEditor.latentNoiseTimestep || 0), 0, MASK_LATENT_NOISE_MAX_TIMESTEP));
  if (!ctx || !width || !height || timestep <= 0) {
    return;
  }
  const noiseValues = ensureMaskLatentNoiseBuffer(width, height);
  if (!noiseValues) {
    return;
  }
  const { signalScale, noiseScale } = getMaskLatentNoiseWeights(timestep);
  if (noiseScale <= 0) {
    return;
  }
  const imageData = ctx.getImageData(0, 0, width, height);
  const { data } = imageData;
  const pixelCount = width * height;
  for (let pixelIndex = 0; pixelIndex < pixelCount; pixelIndex += 1) {
    const dataIndex = pixelIndex * 4;
    const noiseIndex = pixelIndex * 3;
    for (let channel = 0; channel < 3; channel += 1) {
      const normalized = (data[dataIndex + channel] / 127.5) - 1;
      const mixed = signalScale * normalized + noiseScale * noiseValues[noiseIndex + channel];
      data[dataIndex + channel] = Math.round(clamp((mixed + 1) * 127.5, 0, 255));
    }
  }
  ctx.putImageData(imageData, 0, 0);
}

function syncMaskLatentBaseWidthFromPresets() {
  const nextValue = getNearestMaskLatentBaseWidthPreset(state.maskEditor.latentBaseWidth || 512);
  const previousValue = Number(state.maskEditor.latentBaseWidth || 0);
  state.maskEditor.latentBaseWidth = nextValue;
  if (state.maskEditor.active && isMaskEditorMaskMode()) {
    updateMaskControlLabels();
    if (previousValue !== nextValue) {
      scheduleMaskLatentPreviewRender({ imageDirty: true });
    }
  }
}

function updateMaskControlLabels() {
  const imageMode = isMaskEditorImageMode();
  const brushSizePercent = clamp(Number(state.maskEditor.brushSizePercent || 6), 0.2, 100);
  const brushValue = Math.max(0, Math.min(100, Number(state.maskEditor.brushValue || 0)));
  const brushCore = clamp(Number(state.maskEditor.brushCore || 30), 0, 95);
  const brushSteepness = clamp(Number(state.maskEditor.brushSteepness || 8), 1, 32);
  const latentNoiseTimestep = Math.round(clamp(Number(state.maskEditor.latentNoiseTimestep || 0), 0, MASK_LATENT_NOISE_MAX_TIMESTEP));
  const brushDiameterMaskPx = getMaskBrushDiameterMaskPx();
  const latentMetrics = getMaskLatentPreviewMetrics();
  const latentBaseWidthPresets = getMaskLatentBaseWidthPresets();
  const brushColor = String(state.maskEditor.brushColor || "#ff5a5a").toLowerCase();
  maskEditorTitle.textContent = imageMode ? "Image" : "Mask";
  maskBrushSizeInput.value = String(brushSizePercent);
  maskBrushValueInput.value = String(Math.round(brushValue));
  maskBrushColorInput.value = brushColor;
  maskBrushCoreInput.value = String(Math.round(brushCore));
  maskBrushSteepnessInput.value = String(brushSteepness);
  maskLatentBaseWidthInput.min = "0";
  maskLatentBaseWidthInput.max = String(Math.max(0, latentBaseWidthPresets.length - 1));
  maskLatentBaseWidthInput.step = "1";
  maskLatentBaseWidthInput.value = String(getMaskLatentBaseWidthPresetIndex(latentMetrics.baseWidth, latentBaseWidthPresets));
  maskLatentDividerInput.value = String(latentMetrics.divider);
  maskLatentNoiseInput.min = "0";
  maskLatentNoiseInput.max = String(MASK_LATENT_NOISE_MAX_TIMESTEP);
  maskLatentNoiseInput.step = "1";
  maskLatentNoiseInput.value = String(latentNoiseTimestep);
  maskBrushValueTitle.textContent = imageMode ? "Strength" : "Value";
  maskBrushSizeLabel.textContent = `${brushSizePercent.toFixed(1)}% · ${Math.round(brushDiameterMaskPx)} px`;
  maskBrushValueLabel.textContent = `${Math.round(brushValue)}%`;
  maskBrushColorLabel.textContent = brushColor;
  maskBrushCoreLabel.textContent = `${Math.round(brushCore)}%`;
  maskBrushSteepnessLabel.textContent = brushSteepness.toFixed(1);
  maskLatentNoiseLabel.textContent = `t=${latentNoiseTimestep}`;
  if (imageMode) {
    maskResetBtn.textContent = "Reset Image";
    maskResetBtn.title = "Restore the image edit overlay to the original image";
    maskEditorStatus.textContent = state.maskEditor.loading
      ? "Loading..."
      : (state.maskEditor.saving ? "Saving..." : `${Math.round(brushValue)}% · ${brushColor}`);
    return;
  }

  maskResetBtn.textContent = `Reset ${Math.round(brushValue)}%`;
  maskResetBtn.title = `Fill the full mask with ${Math.round(brushValue)}%`;
  maskLatentBaseWidthLabel.textContent = `${latentMetrics.baseWidth}px`;
  maskLatentDividerLabel.textContent = `/${latentMetrics.divider}`;
  maskLatentBaseSizeLabel.textContent = `Base ${latentMetrics.baseWidth}×${latentMetrics.baseHeight}`;
  maskLatentGridSizeLabel.textContent = `Latent ${latentMetrics.latentWidth}×${latentMetrics.latentHeight}`;
  maskLatentSignalLabel.textContent = `Signal ${state.maskEditor.latentSignalPercent.toFixed(1)}%`;
  maskLatentReductionLabel.textContent = `Reduction ${state.maskEditor.latentReductionPercent.toFixed(1)}%`;
  maskEditorStatus.textContent = state.maskEditor.loading
    ? "Loading..."
    : (state.maskEditor.saving ? "Saving..." : `${Math.round(brushValue)}%`);
}

function getMaskLatentPreviewMetrics() {
  const sourceWidth = Math.max(1, Number(state.maskEditor.imageWidth || imgNatW || 1));
  const sourceHeight = Math.max(1, Number(state.maskEditor.imageHeight || imgNatH || 1));
  const baseWidth = getNearestMaskLatentBaseWidthPreset(state.maskEditor.latentBaseWidth || 512);
  const baseHeight = Math.max(1, Math.round(baseWidth * (sourceHeight / Math.max(1, sourceWidth))));
  const divider = Math.round(clamp(Number(state.maskEditor.latentDivider || 16), 1, 64));
  const latentWidth = Math.max(1, Math.round(baseWidth / divider));
  const latentHeight = Math.max(1, Math.round(baseHeight / divider));
  return { sourceWidth, sourceHeight, baseWidth, baseHeight, divider, latentWidth, latentHeight };
}

function resizeCanvasTo(canvas, width, height) {
  if (!canvas) return false;
  if (canvas.width === width && canvas.height === height) {
    return false;
  }
  canvas.width = width;
  canvas.height = height;
  return true;
}

function ensureMaskLatentPreviewBuffers() {
  const metrics = getMaskLatentPreviewMetrics();
  if (!state.maskEditor.latentBaseMaskCanvas) {
    state.maskEditor.latentBaseMaskCanvas = document.createElement("canvas");
  }
  if (!state.maskEditor.latentGridCanvas) {
    state.maskEditor.latentGridCanvas = document.createElement("canvas");
  }
  const resizedImageCanvas = resizeCanvasTo(previewLatentImageCanvas, metrics.baseWidth, metrics.baseHeight);
  resizeCanvasTo(state.maskEditor.latentBaseMaskCanvas, metrics.baseWidth, metrics.baseHeight);
  resizeCanvasTo(state.maskEditor.latentGridCanvas, metrics.latentWidth, metrics.latentHeight);
  resizeCanvasTo(previewLatentMaskCanvas, metrics.latentWidth, metrics.latentHeight);
  if (resizedImageCanvas) {
    state.maskEditor.latentImageDirty = true;
  }
  return {
    ...metrics,
    latentBaseMaskCanvas: state.maskEditor.latentBaseMaskCanvas,
    latentGridCanvas: state.maskEditor.latentGridCanvas,
  };
}

function updateMaskLatentSignalStats(sourceCanvas) {
  const canvas = sourceCanvas || previewLatentMaskCanvas;
  if (!canvas || !canvas.width || !canvas.height) {
    state.maskEditor.latentSignalPercent = 0;
    state.maskEditor.latentReductionPercent = 100;
    updateMaskLatentSignalBuffer(null);
    return;
  }
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const { data } = ctx.getImageData(0, 0, canvas.width, canvas.height);
  let sum = 0;
  for (let index = 0; index < data.length; index += 4) {
    sum += data[index];
  }
  const pixelCount = Math.max(1, canvas.width * canvas.height);
  const signalPercent = clamp((sum / (pixelCount * 255)) * 100, 0, 100);
  state.maskEditor.latentSignalPercent = signalPercent;
  state.maskEditor.latentReductionPercent = 100 - signalPercent;
  updateMaskLatentSignalBuffer(canvas);
}

function scheduleMaskLatentPreviewRender(options = {}) {
  const { imageDirty = false } = options;
  if (imageDirty) {
    state.maskEditor.latentImageDirty = true;
  }
  if (!state.maskEditor.active || !isMaskEditorMaskMode()) {
    return;
  }
  if (state.maskEditor.latentPreviewQueued) return;
  state.maskEditor.latentPreviewQueued = true;
  window.requestAnimationFrame(() => {
    state.maskEditor.latentPreviewQueued = false;
    renderMaskLatentPreview();
  });
}

function renderMaskLatentPreview() {
  if (!state.maskEditor.active || !isMaskEditorMaskMode() || !previewMaskCanvas.width || !previewMaskCanvas.height) {
    return;
  }

  const {
    baseWidth,
    baseHeight,
    latentWidth,
    latentHeight,
    latentBaseMaskCanvas,
    latentGridCanvas,
  } = ensureMaskLatentPreviewBuffers();

  if (state.maskEditor.latentImageDirty && previewImg.complete && previewImg.naturalWidth && previewImg.naturalHeight) {
    const latentImageCtx = previewLatentImageCanvas.getContext("2d", { willReadFrequently: true });
    latentImageCtx.setTransform(1, 0, 0, 1, 0, 0);
    latentImageCtx.clearRect(0, 0, previewLatentImageCanvas.width, previewLatentImageCanvas.height);
    latentImageCtx.imageSmoothingEnabled = true;
    latentImageCtx.drawImage(previewImg, 0, 0, baseWidth, baseHeight);
    applyMaskLatentNoisePreview(latentImageCtx, baseWidth, baseHeight);
    state.maskEditor.latentImageDirty = false;
  }

  const latentBaseMaskCtx = latentBaseMaskCanvas.getContext("2d");
  latentBaseMaskCtx.setTransform(1, 0, 0, 1, 0, 0);
  latentBaseMaskCtx.clearRect(0, 0, baseWidth, baseHeight);
  latentBaseMaskCtx.imageSmoothingEnabled = false;
  latentBaseMaskCtx.drawImage(previewMaskCanvas, 0, 0, baseWidth, baseHeight);

  const latentGridCtx = latentGridCanvas.getContext("2d");
  latentGridCtx.setTransform(1, 0, 0, 1, 0, 0);
  latentGridCtx.clearRect(0, 0, latentWidth, latentHeight);
  latentGridCtx.imageSmoothingEnabled = false;
  latentGridCtx.drawImage(latentBaseMaskCanvas, 0, 0, latentWidth, latentHeight);

  const latentMaskCtx = previewLatentMaskCanvas.getContext("2d");
  latentMaskCtx.setTransform(1, 0, 0, 1, 0, 0);
  latentMaskCtx.clearRect(0, 0, previewLatentMaskCanvas.width, previewLatentMaskCanvas.height);
  latentMaskCtx.imageSmoothingEnabled = false;
  latentMaskCtx.drawImage(latentGridCanvas, 0, 0);
  updateMaskLatentSignalStats(previewLatentMaskCanvas);
  updateMaskSignalProbeStats();
  updateMaskControlLabels();
  renderMaskSignalProbeUi();
}

function scheduleMaskMiniPreviewRender() {
  if (state.maskEditor.previewQueued) return;
  state.maskEditor.previewQueued = true;
  window.requestAnimationFrame(() => {
    state.maskEditor.previewQueued = false;
    renderMaskMiniPreview();
  });
}

function renderMaskMiniPreview() {
  const ctx = maskMiniPreview.getContext("2d");
  ctx.clearRect(0, 0, maskMiniPreview.width, maskMiniPreview.height);
  if (!isMaskEditorVisible()) {
    return;
  }
  if (isMaskEditorImageMode()) {
    if (!previewImageEditCanvas.width || !previewImageEditCanvas.height || !state.maskEditor.imageBaseCanvas) {
      return;
    }
    ctx.imageSmoothingEnabled = true;
    ctx.drawImage(state.maskEditor.imageBaseCanvas, 0, 0, maskMiniPreview.width, maskMiniPreview.height);
    ctx.save();
    ctx.globalCompositeOperation = "color";
    ctx.drawImage(previewImageEditCanvas, 0, 0, maskMiniPreview.width, maskMiniPreview.height);
    ctx.restore();
    return;
  }
  if (!previewMaskCanvas.width || !previewMaskCanvas.height) {
    return;
  }
  ctx.imageSmoothingEnabled = true;
  ctx.drawImage(previewMaskCanvas, 0, 0, maskMiniPreview.width, maskMiniPreview.height);
}

function cloneMaskCanvasSnapshot(sourceCanvas = getActiveEditCanvas()) {
  const snapshot = document.createElement("canvas");
  snapshot.width = sourceCanvas.width;
  snapshot.height = sourceCanvas.height;
  if (snapshot.width && snapshot.height) {
    snapshot.getContext("2d").drawImage(sourceCanvas, 0, 0);
  }
  return snapshot;
}

function refreshMaskBaseCanvas() {
  state.maskEditor.baseCanvas = cloneMaskCanvasSnapshot(getActiveEditCanvas());
}

function syncMaskEditorDirtyState() {
  state.maskEditor.dirty = state.maskEditor.historyIndex !== state.maskEditor.cleanHistoryIndex;
}

function resetMaskHistory() {
  state.maskEditor.history = [];
  state.maskEditor.historyIndex = 0;
  state.maskEditor.cleanHistoryIndex = 0;
  syncMaskEditorDirtyState();
}

function collectEditorHistoryTiles(beforeCanvas, afterCanvas, tileKeys) {
  const tileRects = getEditorTileRects(tileKeys, afterCanvas.width, afterCanvas.height);
  if (!tileRects.length) {
    return [];
  }
  const beforeCtx = beforeCanvas.getContext("2d", { willReadFrequently: true });
  const afterCtx = afterCanvas.getContext("2d", { willReadFrequently: true });
  const tiles = [];
  for (const tileRect of tileRects) {
    const beforeImageData = beforeCtx.getImageData(tileRect.left, tileRect.top, tileRect.width, tileRect.height);
    const afterImageData = afterCtx.getImageData(tileRect.left, tileRect.top, tileRect.width, tileRect.height);
    if (areImageDataEqual(beforeImageData, afterImageData)) {
      continue;
    }
    tiles.push({
      x: tileRect.left,
      y: tileRect.top,
      before: beforeImageData,
      after: afterImageData,
    });
  }
  return tiles;
}

function pushMaskHistorySnapshot(options = {}) {
  const {
    beforeCanvas = state.maskEditor.strokeBaseCanvas,
    tileKeys = state.maskEditor.strokeDirtyTiles,
  } = options;
  const activeCanvas = getActiveEditCanvas();
  if (!beforeCanvas || !activeCanvas.width || !activeCanvas.height || !tileKeys?.size) {
    return;
  }
  if (state.maskEditor.cleanHistoryIndex > state.maskEditor.historyIndex) {
    state.maskEditor.cleanHistoryIndex = -1;
  }
  const nextHistory = state.maskEditor.history.slice(0, state.maskEditor.historyIndex);
  const nextEntry = {
    tiles: collectEditorHistoryTiles(beforeCanvas, activeCanvas, tileKeys),
  };
  if (!nextEntry.tiles.length) {
    return;
  }
  nextHistory.push(nextEntry);
  state.maskEditor.history = nextHistory;
  state.maskEditor.historyIndex = nextHistory.length;
  syncMaskEditorDirtyState();
}

function applyEditorHistoryEntry(entry, direction = "after") {
  const activeCanvas = getActiveEditCanvas();
  if (!entry?.tiles?.length || !activeCanvas.width || !activeCanvas.height) return;
  const ctx = activeCanvas.getContext("2d");
  for (const tile of entry.tiles) {
    ctx.putImageData(direction === "before" ? tile.before : tile.after, tile.x, tile.y);
  }
}

function finalizeHistoryPlayback() {
  scheduleMaskMiniPreviewRender();
  if (isMaskEditorMaskMode()) {
    updateMaskSignalProbeStats();
    scheduleMaskLatentPreviewRender();
  }
  renderMaskEditorUi();
  refreshMaskCursorValue();
}

function undoMaskEdit() {
  if (!state.maskEditor.active || state.maskEditor.painting || state.maskEditor.historyIndex <= 0) return;
  const entry = state.maskEditor.history[state.maskEditor.historyIndex - 1];
  applyEditorHistoryEntry(entry, "before");
  state.maskEditor.historyIndex -= 1;
  syncMaskEditorDirtyState();
  finalizeHistoryPlayback();
  statusBar.textContent = "Undid brush stroke";
}

function redoMaskEdit() {
  if (!state.maskEditor.active || state.maskEditor.painting || state.maskEditor.historyIndex >= state.maskEditor.history.length) return;
  const entry = state.maskEditor.history[state.maskEditor.historyIndex];
  applyEditorHistoryEntry(entry, "after");
  state.maskEditor.historyIndex += 1;
  syncMaskEditorDirtyState();
  finalizeHistoryPlayback();
  statusBar.textContent = "Redid brush stroke";
}

function applyMaskViewMode() {
  if (!isMaskEditorMaskMode()) {
    previewMaskCanvas.dataset.viewMode = "overlay";
    previewLatentMaskCanvas.dataset.viewMode = "overlay";
    return;
  }
  const viewMode = state.maskEditor.viewMode === "mask" ? "mask" : "overlay";
  previewMaskCanvas.dataset.viewMode = viewMode;
  previewLatentMaskCanvas.dataset.viewMode = viewMode;
}

function updateMaskViewModeButton() {
  const showMask = state.maskEditor.viewMode !== "mask";
  maskViewModeBtn.textContent = showMask ? "Show Mask" : "Show Overlay";
  maskViewModeBtn.setAttribute("aria-pressed", showMask ? "false" : "true");
  maskViewModeBtn.title = showMask ? "Switch to grayscale mask view" : "Switch to overlay view";
  maskViewModeBtn.disabled = !isMaskEditorMaskMode() || state.maskEditor.loading;
}

function updateMaskHistoryButtons() {
  const active = state.maskEditor.active && !state.maskEditor.loading && !state.maskEditor.saving;
  maskUndoBtn.disabled = !active || state.maskEditor.painting || state.maskEditor.historyIndex <= 0;
  maskRedoBtn.disabled = !active || state.maskEditor.painting || state.maskEditor.historyIndex >= state.maskEditor.history.length;
}

function updateMaskLatentPreviewButton() {
  const enabled = !!state.maskEditor.latentPreviewEnabled;
  maskLatentPreviewBtn.textContent = enabled ? "Hide Latent" : "Latent Preview";
  maskLatentPreviewBtn.setAttribute("aria-pressed", enabled ? "true" : "false");
  maskLatentPreviewBtn.title = enabled
    ? "Switch back to the full-resolution preview"
    : "Show the latent-space preview built from the current mask";
  maskLatentPreviewBtn.disabled = !isMaskEditorMaskMode() || state.maskEditor.loading;
}

function renderPreviewActionBar() {
  const active = isMaskEditorVisible();
  const imageAvailable = isImageEditAvailable();
  const maskAvailable = isMaskEditAvailable();
  duplicateImageBtn.classList.toggle("visible", imageAvailable && !active);
  imageEditBtn.classList.toggle("visible", imageAvailable && !active);
  renderPromptPreviewButton();
  maskEditBtn.classList.toggle("visible", maskAvailable && !active);
  duplicateImageBtn.disabled = !imageAvailable || state.duplicatingImage || state.autoCaptioning || state.cloning || state.uploading;
  imageEditBtn.disabled = !imageAvailable || state.duplicatingImage || state.autoCaptioning || state.cloning || state.uploading;
  maskEditBtn.disabled = !maskAvailable || state.duplicatingImage || state.autoCaptioning || state.cloning || state.uploading;
  duplicateImageBtn.textContent = state.duplicatingImage ? "Duplicating..." : "Duplicate";
  duplicateImageBtn.title = "Duplicate this image with its caption and mask sidecars";
  imageEditBtn.title = "Paint a color overlay that preserves the image detail and shading";
  renderGifConvertButton();
  const visible = !active && [duplicateImageBtn, imageEditBtn, promptPreviewBtn, maskEditBtn, gifConvertBtn].some((button) => button.classList.contains("visible"));
  previewActionBar.classList.toggle("visible", visible);
}

function syncMaskPreviewLayerVisibility() {
  const active = isMaskEditorVisible();
  const imageMode = isMaskEditorImageMode();
  const showLatentPreview = isMaskEditorMaskMode() && !!state.maskEditor.latentPreviewEnabled;
  if (state.previewMediaType === "image") {
    previewImg.style.display = state.previewPath && imgNatW ? "block" : "none";
  } else if (state.previewMediaType === "video") {
    const videoReady = state.previewPath && previewVideo.currentSrc && previewVideo.readyState >= 2;
    previewImg.style.display = active && !!state.maskEditor.videoSnapshotUrl && (showLatentPreview || !videoReady) ? "block" : "none";
    previewVideo.style.display = videoReady && (!active || !showLatentPreview) ? "block" : "none";
    if (previewImg.style.display === "block") {
      applyTransformToElement(previewImg);
    }
    if (previewVideo.style.display === "block") {
      applyTransformToElement(previewVideo);
    }
  }
  previewImageEditCanvas.style.display = imageMode ? "block" : "none";
  previewMaskCanvas.style.display = isMaskEditorMaskMode() && !showLatentPreview ? "block" : "none";
  previewLatentImageCanvas.style.display = showLatentPreview ? "block" : "none";
  previewLatentMaskCanvas.style.display = showLatentPreview ? "block" : "none";
}

function applyTransformToElement(element) {
  if (!element) return;
  element.style.transform = `translate(${panX}px, ${panY}px) scale(${zoomLevel})`;
  element.style.width = imgNatW + "px";
  element.style.height = imgNatH + "px";
}

function renderMaskEditorUi() {
  const maskAvailable = isMaskEditAvailable();
  const imageAvailable = isImageEditAvailable();
  const videoAvailable = isVideoMaskEditAvailable();
  const active = isMaskEditorVisible();
  const imageMode = isMaskEditorImageMode();
  const maskMode = isMaskEditorMaskMode();
  const interactive = active && !state.maskEditor.loading && !state.maskEditor.saving;
  const showLatentPreview = maskMode && !!state.maskEditor.latentPreviewEnabled;
  const currentVideoFrameIndex = videoAvailable ? getCurrentVideoMaskFrameIndex() : null;
  const canCreateVideoKeyframe = videoAvailable && (!active || (!state.maskEditor.loading && !state.maskEditor.saving && !state.maskEditor.painting && !state.maskEditor.switchingKeyframe));
  updateMaskControlLabels();
  applyMaskViewMode();
  updateMaskViewModeButton();
  updateMaskLatentPreviewButton();
  updateMaskHistoryButtons();
  renderMaskSignalProbeUi();
  renderPreviewActionBar();
  maskEditorPanel.classList.toggle("visible", active);
  maskBrushSizeInput.disabled = !interactive;
  maskBrushValueInput.disabled = !interactive;
  maskBrushColorField.classList.toggle("visible", imageMode);
  maskBrushColorInput.disabled = !interactive || !imageMode;
  maskBrushCoreInput.disabled = !interactive;
  maskBrushSteepnessInput.disabled = !interactive;
  maskLatentBaseWidthInput.disabled = !interactive || !showLatentPreview || !maskMode;
  maskLatentDividerInput.disabled = !interactive || !showLatentPreview || !maskMode;
  maskLatentNoiseInput.disabled = !interactive || !showLatentPreview || !maskMode;
  maskApplyBtn.textContent = imageMode ? "Save Image" : "Save Mask";
  maskCancelBtn.textContent = imageMode ? "Cancel Edit" : "Cancel Mask";
  videoMaskAddBtn.classList.toggle("visible", videoAvailable);
  videoMaskAddBtn.classList.toggle("in-editor", active && videoAvailable);
  videoMaskAddBtn.disabled = !canCreateVideoKeyframe;
  maskActionBar.classList.toggle("visible", active);
  maskActionBar.classList.toggle("with-video-key-add", active && videoAvailable);
  maskApplyBtn.classList.toggle("visible", active);
  maskCancelBtn.classList.toggle("visible", active);
  maskUndoBtn.classList.toggle("visible", active);
  maskRedoBtn.classList.toggle("visible", active);
  maskViewModeBtn.classList.toggle("visible", active && maskMode);
  maskLatentPreviewBtn.classList.toggle("visible", active && maskMode);
  maskResetBtn.classList.toggle("visible", active);
  maskLatentPreviewControls.classList.toggle("visible", showLatentPreview);
  previewStage.classList.toggle("mask-signal-probe-mode", isMaskSignalProbeMode());
  if (state.previewMediaType === "video") {
    maskEditBtn.textContent = "Edit Key Mask";
    maskEditBtn.title = videoAvailable
      ? `Edit the active key-frame mask at or before ${formatVideoMaskFrameHint(currentVideoFrameIndex)}`
      : "Edit the key-frame mask at the current video frame";
    videoMaskAddBtn.title = videoAvailable
      ? (active
        ? `Create a new key-frame mask at ${formatVideoMaskFrameHint(currentVideoFrameIndex)} and keep editing`
        : `Add a new key-frame mask at ${formatVideoMaskFrameHint(currentVideoFrameIndex)}`)
      : "Add a new key-frame mask at the current video frame";
  } else {
    maskEditBtn.textContent = "Mask";
    maskEditBtn.title = "Edit the image mask";
    videoMaskAddBtn.title = "Add a new key-frame mask at the current video frame";
  }
  imageEditBtn.textContent = "Edit Image";
  syncMaskPreviewLayerVisibility();
  applyImageEditCanvasTransform();
  applyMaskCanvasTransform();
  applyMaskLatentPreviewTransform();
  if (showLatentPreview) {
    scheduleMaskLatentPreviewRender();
  }
  if (!active) {
    clearMaskCursor();
  }
  renderGifConvertButton();
  renderPreviewCaptionOverlay();
}

function applyMaskCanvasTransform() {
  if (!isMaskEditorMaskMode()) {
    previewMaskCanvas.style.display = "none";
    return;
  }
  previewMaskCanvas.style.display = state.maskEditor.latentPreviewEnabled ? "none" : "block";
  previewMaskCanvas.style.transform = `translate(${panX}px, ${panY}px) scale(${zoomLevel})`;
  previewMaskCanvas.style.width = `${imgNatW}px`;
  previewMaskCanvas.style.height = `${imgNatH}px`;
  scheduleMaskMiniPreviewRender();
}

function applyImageEditCanvasTransform() {
  const visible = isMaskEditorImageMode();
  previewImageEditCanvas.style.display = visible ? "block" : "none";
  if (!visible) {
    return;
  }
  previewImageEditCanvas.style.transform = `translate(${panX}px, ${panY}px) scale(${zoomLevel})`;
  previewImageEditCanvas.style.width = `${imgNatW}px`;
  previewImageEditCanvas.style.height = `${imgNatH}px`;
  scheduleMaskMiniPreviewRender();
}

function applyMaskLatentPreviewTransform() {
  const visible = isMaskEditorMaskMode() && !!state.maskEditor.latentPreviewEnabled;
  previewLatentImageCanvas.style.display = visible ? "block" : "none";
  previewLatentMaskCanvas.style.display = visible ? "block" : "none";
  if (!visible) {
    return;
  }
  const transform = `translate(${panX}px, ${panY}px) scale(${zoomLevel})`;
  previewLatentImageCanvas.style.transform = transform;
  previewLatentImageCanvas.style.width = `${imgNatW}px`;
  previewLatentImageCanvas.style.height = `${imgNatH}px`;
  previewLatentMaskCanvas.style.transform = transform;
  previewLatentMaskCanvas.style.width = `${imgNatW}px`;
  previewLatentMaskCanvas.style.height = `${imgNatH}px`;
}

function isClientInsidePreviewImage(clientX, clientY) {
  const panelRect = previewStage.getBoundingClientRect();
  const px = clientX - panelRect.left;
  const py = clientY - panelRect.top;
  return px >= panX
    && px <= panX + imgNatW * zoomLevel
    && py >= panY
    && py <= panY + imgNatH * zoomLevel;
}

function updateMaskCursor(clientX, clientY) {
  if (isMaskSignalProbeMode()) {
    clearMaskCursor();
    return;
  }
  if (!isMaskEditorVisible() || !imgNatW || !imgNatH || !isClientInsidePreviewImage(clientX, clientY)) {
    clearMaskCursor();
    return;
  }
  const panelRect = previewStage.getBoundingClientRect();
  const diameter = Math.max(4, getMaskBrushDiameterPreviewPx() * zoomLevel);
  state.maskEditor.cursorClientX = clientX;
  state.maskEditor.cursorClientY = clientY;
  maskCursor.style.width = `${diameter}px`;
  maskCursor.style.height = `${diameter}px`;
  maskCursor.style.left = `${clientX - panelRect.left}px`;
  maskCursor.style.top = `${clientY - panelRect.top}px`;
  if (maskCursorValue) {
    maskCursorValue.style.fontSize = `${clamp(diameter * 0.22, 9, 14)}px`;
    maskCursorValue.style.transform = `translate(-50%, calc(-100% - ${Math.max(10, Math.round(diameter * 0.2))}px))`;
  }
  maskCursor.classList.add("visible");
  refreshMaskCursorValue();
}

function getMaskBrushRadius() {
  return Math.max(1, getMaskBrushDiameterMaskPx() / 2);
}

function ensureMaskStrokeCanvases() {
  const activeCanvas = getActiveEditCanvas();
  const width = Math.max(1, activeCanvas.width || 1);
  const height = Math.max(1, activeCanvas.height || 1);

  if (!state.maskEditor.strokeBaseCanvas) {
    state.maskEditor.strokeBaseCanvas = document.createElement("canvas");
  }
  if (state.maskEditor.strokeBaseCanvas.width !== width || state.maskEditor.strokeBaseCanvas.height !== height) {
    state.maskEditor.strokeBaseCanvas.width = width;
    state.maskEditor.strokeBaseCanvas.height = height;
  }

  const pixelCount = width * height;
  if (!state.maskEditor.strokeInfluenceValues || state.maskEditor.strokeInfluenceValues.length !== pixelCount) {
    state.maskEditor.strokeInfluenceValues = new Float32Array(pixelCount);
  }
}

function clearMaskStrokeRenderFrame() {
  if (!state.maskEditor.strokeRenderFrameId) return;
  window.cancelAnimationFrame(state.maskEditor.strokeRenderFrameId);
  state.maskEditor.strokeRenderFrameId = 0;
}

function renderMaskStrokePreview() {
  state.maskEditor.strokeRenderFrameId = 0;
  const activeCanvas = getActiveEditCanvas();
  if (!activeCanvas.width || !activeCanvas.height) return;
  const strokeBaseCanvas = state.maskEditor.strokeBaseCanvas;
  const strokeInfluenceValues = state.maskEditor.strokeInfluenceValues;
  const dirtyTiles = getEditorTileRects(state.maskEditor.strokeDirtyTiles, activeCanvas.width, activeCanvas.height);
  if (!strokeBaseCanvas || !strokeInfluenceValues || !dirtyTiles.length) return;

  const targetValue = clamp(Number(state.maskEditor.brushValue || 0), 0, 100) * 2.55;
  const canvasWidth = activeCanvas.width;
  const outputCtx = activeCanvas.getContext("2d", { willReadFrequently: true });
  const baseCtx = strokeBaseCanvas.getContext("2d", { willReadFrequently: true });
  const imageMode = isMaskEditorImageMode();
  const brushStrength = clamp(Number(state.maskEditor.brushValue || 0), 0, 100) / 100;
  const targetColor = hexToRgb(state.maskEditor.brushColor);

  for (const tile of dirtyTiles) {
    const baseImage = baseCtx.getImageData(tile.left, tile.top, tile.width, tile.height);
    const outputImage = outputCtx.getImageData(tile.left, tile.top, tile.width, tile.height);
    const baseData = baseImage.data;
    const outputData = outputImage.data;

    for (let y = 0; y < tile.height; y += 1) {
      for (let x = 0; x < tile.width; x += 1) {
        const pixelIndex = y * tile.width + x;
        const dataIndex = pixelIndex * 4;
        const influenceIndex = (tile.top + y) * canvasWidth + (tile.left + x);
        const influence = strokeInfluenceValues[influenceIndex] || 0;
        if (influence <= 0) {
          outputData[dataIndex] = baseData[dataIndex];
          outputData[dataIndex + 1] = baseData[dataIndex + 1];
          outputData[dataIndex + 2] = baseData[dataIndex + 2];
          outputData[dataIndex + 3] = imageMode ? baseData[dataIndex + 3] : 255;
          continue;
        }
        if (imageMode) {
          const targetAlpha = clamp(influence * brushStrength, 0, 1);
          const baseAlpha = clamp((baseData[dataIndex + 3] || 0) / 255, 0, 1);
          const outAlpha = targetAlpha + (baseAlpha * (1 - targetAlpha));
          if (outAlpha <= 0) {
            outputData[dataIndex] = 0;
            outputData[dataIndex + 1] = 0;
            outputData[dataIndex + 2] = 0;
            outputData[dataIndex + 3] = 0;
            continue;
          }
          const preservedBaseFactor = baseAlpha * (1 - targetAlpha);
          outputData[dataIndex] = Math.round(((targetColor.r * targetAlpha) + (baseData[dataIndex] * preservedBaseFactor)) / outAlpha);
          outputData[dataIndex + 1] = Math.round(((targetColor.g * targetAlpha) + (baseData[dataIndex + 1] * preservedBaseFactor)) / outAlpha);
          outputData[dataIndex + 2] = Math.round(((targetColor.b * targetAlpha) + (baseData[dataIndex + 2] * preservedBaseFactor)) / outAlpha);
          outputData[dataIndex + 3] = Math.round(outAlpha * 255);
          continue;
        }
        const baseValue = baseData[dataIndex];
        const nextValue = Math.round(baseValue * (1 - influence) + targetValue * influence);
        outputData[dataIndex] = nextValue;
        outputData[dataIndex + 1] = nextValue;
        outputData[dataIndex + 2] = nextValue;
        outputData[dataIndex + 3] = 255;
      }
    }
    outputCtx.putImageData(outputImage, tile.left, tile.top);
  }

  state.maskEditor.dirty = true;
  updateMaskSignalProbeStats();
  scheduleMaskMiniPreviewRender();
  if (isMaskEditorMaskMode()) {
    scheduleMaskLatentPreviewRender();
  }
  refreshMaskCursorValue();
}

function scheduleMaskStrokePreviewRender() {
  if (state.maskEditor.strokeRenderFrameId) return;
  state.maskEditor.strokeRenderFrameId = window.requestAnimationFrame(() => {
    renderMaskStrokePreview();
  });
}

function markMaskStrokeDirtyTiles(maskX, maskY, radius) {
  const activeCanvas = getActiveEditCanvas();
  const minTileX = Math.max(0, Math.floor((maskX - radius - 2) / EDITOR_HISTORY_TILE_SIZE));
  const minTileY = Math.max(0, Math.floor((maskY - radius - 2) / EDITOR_HISTORY_TILE_SIZE));
  const maxTileX = Math.max(0, Math.floor((Math.min(activeCanvas.width, Math.ceil(maskX + radius + 2)) - 1) / EDITOR_HISTORY_TILE_SIZE));
  const maxTileY = Math.max(0, Math.floor((Math.min(activeCanvas.height, Math.ceil(maskY + radius + 2)) - 1) / EDITOR_HISTORY_TILE_SIZE));
  if (!state.maskEditor.strokeDirtyTiles) {
    state.maskEditor.strokeDirtyTiles = new Set();
  }
  for (let tileY = minTileY; tileY <= maxTileY; tileY += 1) {
    for (let tileX = minTileX; tileX <= maxTileX; tileX += 1) {
      state.maskEditor.strokeDirtyTiles.add(getEditorTileKey(tileX, tileY));
    }
  }
}

function previewPointToMaskPoint(point) {
  return {
    x: clamp(point.x * (state.maskEditor.previewScaleX || 1), 0, state.maskEditor.imageWidth || 1),
    y: clamp(point.y * (state.maskEditor.previewScaleY || 1), 0, state.maskEditor.imageHeight || 1),
  };
}

function paintMaskStamp(maskX, maskY) {
  ensureMaskStrokeCanvases();
  const activeCanvas = getActiveEditCanvas();
  const influenceValues = state.maskEditor.strokeInfluenceValues;
  if (!influenceValues) return;
  const radius = getMaskBrushRadius();
  const width = activeCanvas.width;
  const minX = Math.max(0, Math.floor(maskX - radius - 2));
  const minY = Math.max(0, Math.floor(maskY - radius - 2));
  const maxX = Math.min(activeCanvas.width, Math.ceil(maskX + radius + 2));
  const maxY = Math.min(activeCanvas.height, Math.ceil(maskY + radius + 2));

  for (let y = minY; y < maxY; y += 1) {
    for (let x = minX; x < maxX; x += 1) {
      const dx = x + 0.5 - maskX;
      const dy = y + 0.5 - maskY;
      const distance = Math.hypot(dx, dy);
      if (distance >= radius) {
        continue;
      }
      const influence = getMaskBrushInfluence(clamp(distance / Math.max(radius, 1), 0, 1));
      const index = y * width + x;
      if (influence > influenceValues[index]) {
        influenceValues[index] = influence;
      }
    }
  }

  markMaskStrokeDirtyTiles(maskX, maskY, radius);
  scheduleMaskStrokePreviewRender();
}

function paintMaskAtClient(clientX, clientY) {
  const point = previewPointToMaskPoint(screenToImage(clientX, clientY));
  const previous = state.maskEditor.lastPoint;
  if (!previous) {
    paintMaskStamp(point.x, point.y);
    state.maskEditor.lastPoint = point;
    return;
  }

  const step = Math.max(1, getMaskBrushRadius() * 0.3);
  const dx = point.x - previous.x;
  const dy = point.y - previous.y;
  const distance = Math.hypot(dx, dy);
  const count = Math.max(1, Math.ceil(distance / step));
  for (let index = 1; index <= count; index += 1) {
    const ratio = index / count;
    paintMaskStamp(previous.x + dx * ratio, previous.y + dy * ratio);
  }
  state.maskEditor.lastPoint = point;
}

function beginMaskPaint(event) {
  if (state.previewMediaType === "video" && !previewVideo.paused) {
    previewVideo.pause();
    statusBar.textContent = "Paused video preview for painting";
    return;
  }
  ensureMaskStrokeCanvases();
  const activeCanvas = getActiveEditCanvas();
  const strokeBaseCtx = state.maskEditor.strokeBaseCanvas.getContext("2d");
  strokeBaseCtx.setTransform(1, 0, 0, 1, 0, 0);
  strokeBaseCtx.clearRect(0, 0, state.maskEditor.strokeBaseCanvas.width, state.maskEditor.strokeBaseCanvas.height);
  strokeBaseCtx.drawImage(activeCanvas, 0, 0);
  state.maskEditor.strokeInfluenceValues?.fill(0);
  state.maskEditor.strokeDirtyTiles = new Set();
  state.maskEditor.painting = true;
  state.maskEditor.lastPoint = null;
  paintMaskAtClient(event.clientX, event.clientY);
  updateMaskCursor(event.clientX, event.clientY);
}

function stopMaskPaint() {
  clearMaskStrokeRenderFrame();
  if (state.maskEditor.painting) {
    const historyIndexBefore = state.maskEditor.historyIndex;
    renderMaskStrokePreview();
    pushMaskHistorySnapshot();
    if (state.maskEditor.historyIndex === historyIndexBefore) {
      syncMaskEditorDirtyState();
    }
  }
  state.maskEditor.painting = false;
  state.maskEditor.lastPoint = null;
  state.maskEditor.strokeDirtyTiles = null;
  renderMaskEditorUi();
}

function snapshotMaskBaseCanvas() {
  refreshMaskBaseCanvas();
}

async function fetchMaskMetadata(path, ensure = false, options = {}) {
  const { frameIndex = null, createNew = false } = options;
  const params = new URLSearchParams({
    path,
    ensure: ensure ? "true" : "false",
  });
  if (frameIndex != null) {
    params.set("frame_index", String(frameIndex));
  }
  if (createNew) {
    params.set("create_new", "true");
  }
  const resp = await fetch(`/api/mask?${params.toString()}`);
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data.detail || "Failed to load mask metadata");
  }
  return data;
}

function loadMaskImage(url) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Failed to load image asset"));
    image.src = url;
  });
}

function getEditorSourceDimensions(path = state.previewPath) {
  const cropState = path === state.previewPath ? getCurrentCropState() : (state.imageCrops[path] || null);
  const image = state.images.find((item) => item.path === path) || null;
  return {
    width: Math.max(1, Number(cropState?.current_width || image?.width || imgNatW || 1)),
    height: Math.max(1, Number(cropState?.current_height || image?.height || imgNatH || 1)),
  };
}

async function loadImageEditEditorForPath(path) {
  const sourceSize = getEditorSourceDimensions(path);
  const previewSource = previewCache.get(path) || buildImageApiUrl("preview", path);
  const image = await loadMaskImage(previewSource);
  const imageWidth = Math.max(1, Number(image.naturalWidth || image.width || 1));
  const imageHeight = Math.max(1, Number(image.naturalHeight || image.height || 1));
  const previewScaleX = imageWidth / Math.max(1, imgNatW || imageWidth);
  const previewScaleY = imageHeight / Math.max(1, imgNatH || imageHeight);
  const baseCanvas = document.createElement("canvas");
  baseCanvas.width = imageWidth;
  baseCanvas.height = imageHeight;
  baseCanvas.getContext("2d").drawImage(image, 0, 0, imageWidth, imageHeight);

  previewImageEditCanvas.width = imageWidth;
  previewImageEditCanvas.height = imageHeight;
  previewImageEditCanvas.getContext("2d").clearRect(0, 0, imageWidth, imageHeight);

  state.maskEditor.mode = "image";
  state.maskEditor.path = path;
  state.maskEditor.mediaType = "image";
  state.maskEditor.frameIndex = null;
  state.maskEditor.requestedFrameIndex = null;
  state.maskEditor.sourceFrameIndex = null;
  state.maskEditor.sourceWidth = sourceSize.width;
  state.maskEditor.sourceHeight = sourceSize.height;
  state.maskEditor.imageWidth = imageWidth;
  state.maskEditor.imageHeight = imageHeight;
  state.maskEditor.previewScaleX = previewScaleX;
  state.maskEditor.previewScaleY = previewScaleY;
  state.maskEditor.imageBaseCanvas = baseCanvas;
  state.maskEditor.latentPreviewEnabled = false;
  stopMaskPaint();
  refreshMaskBaseCanvas();
  resetMaskHistory();
  applyImageEditCanvasTransform();
  renderMaskEditorUi();
  scheduleMaskMiniPreviewRender();
  return { image_width: imageWidth, image_height: imageHeight };
}

async function loadImageMaskEditorForPath(path) {
  const maskInfo = await fetchMaskMetadata(path, true);
  const sourceWidth = Math.max(1, Number(maskInfo.image_width || 1));
  const sourceHeight = Math.max(1, Number(maskInfo.image_height || 1));
  const workingSize = getCappedEditorDimensions(sourceWidth, sourceHeight);
  const imageWidth = workingSize.width;
  const imageHeight = workingSize.height;
  const previewScaleX = imageWidth / Math.max(1, imgNatW || imageWidth);
  const previewScaleY = imageHeight / Math.max(1, imgNatH || imageHeight);
  const image = await loadMaskImage(buildImageApiUrl("mask/image", path, {
    ensure: true,
    mask_v: maskInfo.mtime || Date.now(),
  }));

  previewMaskCanvas.width = imageWidth;
  previewMaskCanvas.height = imageHeight;
  const ctx = previewMaskCanvas.getContext("2d");
  ctx.clearRect(0, 0, imageWidth, imageHeight);
  ctx.drawImage(image, 0, 0, imageWidth, imageHeight);

  state.maskEditor.mode = "mask";
  state.maskEditor.path = path;
  state.maskEditor.mediaType = "image";
  state.maskEditor.frameIndex = null;
  state.maskEditor.requestedFrameIndex = null;
  state.maskEditor.sourceFrameIndex = null;
  state.maskEditor.sourceWidth = sourceWidth;
  state.maskEditor.sourceHeight = sourceHeight;
  state.maskEditor.imageWidth = imageWidth;
  state.maskEditor.imageHeight = imageHeight;
  state.maskEditor.previewScaleX = previewScaleX;
  state.maskEditor.previewScaleY = previewScaleY;
  state.maskEditor.imageBaseCanvas = null;
  state.maskEditor.latentImageDirty = true;
  stopMaskPaint();
  refreshMaskBaseCanvas();
  resetMaskHistory();
  updateMaskSignalProbeStats();
  setImageMaskPresence(path, true, maskInfo.mtime || Date.now());
  applyMaskCanvasTransform();
  applyMaskLatentPreviewTransform();
  renderMaskEditorUi();
  scheduleMaskLatentPreviewRender({ imageDirty: true });
  return maskInfo;
}

async function loadVideoMaskEditorForPath(path, options = {}) {
  const { createNew = false, requestedFrameIndex: requestedFrameIndexOption = null } = options;
  const requestedFrameIndex = requestedFrameIndexOption == null
    ? getCurrentVideoMaskFrameIndex(path)
    : Math.max(0, Number.parseInt(requestedFrameIndexOption, 10) || 0);
  await captureCurrentPreviewVideoFrameSnapshot(path);
  const maskInfo = await fetchMaskMetadata(path, true, { frameIndex: requestedFrameIndex, createNew });
  const sourceWidth = Math.max(1, Number(maskInfo.image_width || imgNatW || 1));
  const sourceHeight = Math.max(1, Number(maskInfo.image_height || imgNatH || 1));
  const workingSize = getCappedEditorDimensions(sourceWidth, sourceHeight);
  const imageWidth = workingSize.width;
  const imageHeight = workingSize.height;
  const previewScaleX = imageWidth / Math.max(1, imgNatW || imageWidth);
  const previewScaleY = imageHeight / Math.max(1, imgNatH || imageHeight);
  const image = await loadMaskImage(buildImageApiUrl("mask/image", path, {
    ensure: true,
    frame_index: maskInfo.frame_index,
    create_new: createNew,
    mask_v: maskInfo.mtime || Date.now(),
  }));

  previewMaskCanvas.width = imageWidth;
  previewMaskCanvas.height = imageHeight;
  const ctx = previewMaskCanvas.getContext("2d");
  ctx.clearRect(0, 0, imageWidth, imageHeight);
  ctx.drawImage(image, 0, 0, imageWidth, imageHeight);

  state.maskEditor.mode = "mask";
  state.maskEditor.path = path;
  state.maskEditor.mediaType = "video";
  state.maskEditor.frameIndex = Number(maskInfo.frame_index || requestedFrameIndex || 0);
  state.maskEditor.requestedFrameIndex = Number(maskInfo.requested_frame_index || requestedFrameIndex || 0);
  state.maskEditor.sourceFrameIndex = maskInfo.source_frame_index == null ? null : Number(maskInfo.source_frame_index || 0);
  state.maskEditor.switchingKeyframe = false;
  state.maskEditor.sourceWidth = sourceWidth;
  state.maskEditor.sourceHeight = sourceHeight;
  state.maskEditor.imageWidth = imageWidth;
  state.maskEditor.imageHeight = imageHeight;
  state.maskEditor.previewScaleX = previewScaleX;
  state.maskEditor.previewScaleY = previewScaleY;
  state.maskEditor.imageBaseCanvas = null;
  state.maskEditor.latentImageDirty = true;
  stopMaskPaint();
  refreshMaskBaseCanvas();
  resetMaskHistory();
  updateMaskSignalProbeStats();
  setVideoMaskKeyframes(path, maskInfo.keyframes || []);
  setImageMaskPresence(path, true, maskInfo.mtime || Date.now(), maskInfo.mask_count);
  applyMaskCanvasTransform();
  applyMaskLatentPreviewTransform();
  renderMaskEditorUi();
  renderVideoEditPanel();
  scheduleMaskLatentPreviewRender({ imageDirty: true });
  return maskInfo;
}

async function loadMaskEditorForPath(path, options = {}) {
  if (getMediaType(path) === "video") {
    return loadVideoMaskEditorForPath(path, options);
  }
  return loadImageMaskEditorForPath(path);
}

async function enterMaskEditMode(options = {}) {
  const { createNew = false } = options;
  if (!isMaskEditAvailable()) return;
  await clearPromptPreviewDisplay({ preserveView: true });
  if (state.cropDraft || state.cropInteraction) {
    clearCropDraft();
  }
  state.maskEditor.active = true;
  state.maskEditor.loading = true;
  state.maskEditor.path = state.previewPath;
  renderMaskEditorUi();
  try {
    const maskInfo = await loadMaskEditorForPath(state.previewPath, { createNew });
    if (state.maskEditor.mediaType === "video") {
      statusBar.textContent = `Editing key-frame mask for ${getFileLabel(state.previewPath)} at ${formatVideoMaskFrameHint(maskInfo.frame_index, state.previewPath)}`;
    } else {
      statusBar.textContent = `Editing mask for ${getFileLabel(state.previewPath)}`;
    }
  } catch (err) {
    state.maskEditor.active = false;
    previewMaskCanvas.style.display = "none";
    showErrorToast(`Mask error: ${err.message}`);
    statusBar.textContent = `Mask error: ${err.message}`;
  } finally {
    state.maskEditor.loading = false;
    renderMaskEditorUi();
  }
}

async function enterImageEditMode() {
  if (!isImageEditAvailable()) return;
  await clearPromptPreviewDisplay({ preserveView: true });
  if (state.cropDraft || state.cropInteraction) {
    clearCropDraft();
  }
  state.maskEditor.active = true;
  state.maskEditor.mode = "image";
  state.maskEditor.loading = true;
  state.maskEditor.path = state.previewPath;
  renderMaskEditorUi();
  try {
    await loadImageEditEditorForPath(state.previewPath);
    statusBar.textContent = `Editing image for ${getFileLabel(state.previewPath)}`;
  } catch (err) {
    state.maskEditor.active = false;
    state.maskEditor.mode = null;
    previewImageEditCanvas.style.display = "none";
    showErrorToast(`Image edit error: ${err.message}`);
    statusBar.textContent = `Image edit error: ${err.message}`;
  } finally {
    state.maskEditor.loading = false;
    renderMaskEditorUi();
  }
}

async function enterVideoMaskAddMode() {
  if (!isVideoMaskEditAvailable()) return;
  if (state.maskEditor.active && state.maskEditor.mediaType === "video") {
    if (state.maskEditor.loading || state.maskEditor.saving || state.maskEditor.painting || state.maskEditor.switchingKeyframe) {
      return;
    }
    state.maskEditor.switchingKeyframe = true;
    renderMaskEditorUi();
    try {
      if (state.maskEditor.dirty) {
        await saveMaskEdit({ closeAfterSave: false });
      }
      const requestedFrameIndex = getCurrentVideoMaskFrameIndex(state.previewPath);
      const maskInfo = await loadVideoMaskEditorForPath(state.previewPath, {
        createNew: true,
        requestedFrameIndex,
      });
      statusBar.textContent = `Editing key-frame mask for ${getFileLabel(state.previewPath)} at ${formatVideoMaskFrameHint(maskInfo.frame_index, state.previewPath)}`;
    } finally {
      state.maskEditor.switchingKeyframe = false;
      renderMaskEditorUi();
      renderVideoEditPanel();
    }
    return;
  }
  return enterMaskEditMode({ createNew: true });
}

async function syncActiveVideoMaskEditorToSeekPosition() {
  if (!state.maskEditor.active || state.maskEditor.mediaType !== "video" || !state.previewPath) {
    return;
  }
  if (state.maskEditor.loading || state.maskEditor.saving || state.maskEditor.painting || state.maskEditor.switchingKeyframe) {
    return;
  }

  const requestedFrameIndex = getCurrentVideoMaskFrameIndex(state.previewPath);
  state.maskEditor.requestedFrameIndex = requestedFrameIndex;
  const nextFrameIndex = getResolvedVideoMaskKeyframeForFrame(state.previewPath, requestedFrameIndex);
  if (nextFrameIndex == null || Number(nextFrameIndex) === Number(state.maskEditor.frameIndex)) {
    renderMaskEditorUi();
    renderVideoEditPanel();
    return;
  }

  state.maskEditor.switchingKeyframe = true;
  renderMaskEditorUi();
  try {
    if (state.maskEditor.dirty) {
      await saveMaskEdit({ closeAfterSave: false });
    }
    const maskInfo = await loadVideoMaskEditorForPath(state.previewPath, {
      createNew: false,
      requestedFrameIndex,
    });
    statusBar.textContent = `Editing key-frame mask for ${getFileLabel(state.previewPath)} at ${formatVideoMaskFrameHint(maskInfo.frame_index, state.previewPath)}`;
  } finally {
    state.maskEditor.switchingKeyframe = false;
    renderMaskEditorUi();
    renderVideoEditPanel();
  }
}

function closeMaskEditor(options = {}) {
  const { restoreBase = false } = options;
  clearMaskStrokeRenderFrame();
  const activeCanvas = getActiveEditCanvas();
  if (restoreBase && state.maskEditor.baseCanvas && activeCanvas.width && activeCanvas.height) {
    const ctx = activeCanvas.getContext("2d");
    ctx.clearRect(0, 0, activeCanvas.width, activeCanvas.height);
    ctx.drawImage(state.maskEditor.baseCanvas, 0, 0);
  }
  stopMaskPaint();
  state.maskEditor.active = false;
  state.maskEditor.mode = null;
  state.maskEditor.loading = false;
  state.maskEditor.saving = false;
  state.maskEditor.dirty = false;
  state.maskEditor.switchingKeyframe = false;
  state.maskEditor.path = null;
  state.maskEditor.mediaType = null;
  state.maskEditor.frameIndex = null;
  state.maskEditor.requestedFrameIndex = null;
  state.maskEditor.sourceFrameIndex = null;
  state.maskEditor.sourceWidth = 0;
  state.maskEditor.sourceHeight = 0;
  state.maskEditor.imageWidth = 0;
  state.maskEditor.imageHeight = 0;
  state.maskEditor.previewScaleX = 1;
  state.maskEditor.previewScaleY = 1;
  state.maskEditor.history = [];
  state.maskEditor.historyIndex = 0;
  state.maskEditor.cleanHistoryIndex = 0;
  state.maskEditor.signalProbeMode = false;
  state.maskEditor.signalProbeDragging = false;
  state.maskEditor.signalProbeAnchor = null;
  state.maskEditor.signalProbeRect = null;
  state.maskEditor.signalProbePercent = 0;
  state.maskEditor.signalProbeAreaPercent = 0;
  state.maskEditor.imageBaseCanvas = null;
  state.maskEditor.baseCanvas = null;
  state.maskEditor.latentPreviewQueued = false;
  state.maskEditor.latentImageDirty = true;
  state.maskEditor.latentSignalPercent = 50;
  state.maskEditor.latentReductionPercent = 50;
  state.maskEditor.latentNoiseValues = null;
  state.maskEditor.latentNoiseWidth = 0;
  state.maskEditor.latentNoiseHeight = 0;
  state.maskEditor.latentBaseMaskCanvas = null;
  state.maskEditor.latentGridCanvas = null;
  state.maskEditor.latentSignalValues = null;
  state.maskEditor.latentSignalIntegral = null;
  state.maskEditor.latentSignalWidth = 0;
  state.maskEditor.latentSignalHeight = 0;
  state.maskEditor.latentSignalTotalValue = 0;
  state.maskEditor.strokeBaseCanvas = null;
  state.maskEditor.strokeInfluenceValues = null;
  state.maskEditor.strokeDirtyTiles = null;
  revokeMaskEditorVideoSnapshot();
  if (state.previewMediaType === "video") {
    previewImg.removeAttribute("src");
    previewImg.style.display = "none";
  }
  previewImageEditCanvas.width = 0;
  previewImageEditCanvas.height = 0;
  previewImageEditCanvas.style.display = "none";
  previewMaskCanvas.width = 0;
  previewMaskCanvas.height = 0;
  previewMaskCanvas.style.display = "none";
  previewLatentImageCanvas.width = 0;
  previewLatentImageCanvas.height = 0;
  previewLatentImageCanvas.style.display = "none";
  previewLatentMaskCanvas.width = 0;
  previewLatentMaskCanvas.height = 0;
  previewLatentMaskCanvas.style.display = "none";
  renderMaskMiniPreview();
  renderMaskEditorUi();
}

async function composeEditedImageBlob() {
  const targetPath = state.maskEditor.path || state.previewPath;
  if (!targetPath || !previewImageEditCanvas.width || !previewImageEditCanvas.height) {
    throw new Error("Image editor is not ready");
  }
  const sourceImage = await loadMaskImage(buildImageApiUrl("image", targetPath, { v: getImageVersion(targetPath) }));
  const compositeCanvas = document.createElement("canvas");
  compositeCanvas.width = Math.max(1, Number(sourceImage.naturalWidth || sourceImage.width || state.maskEditor.sourceWidth || 1));
  compositeCanvas.height = Math.max(1, Number(sourceImage.naturalHeight || sourceImage.height || state.maskEditor.sourceHeight || 1));
  const ctx = compositeCanvas.getContext("2d");
  ctx.drawImage(sourceImage, 0, 0, compositeCanvas.width, compositeCanvas.height);
  ctx.save();
  ctx.globalCompositeOperation = "color";
  ctx.drawImage(previewImageEditCanvas, 0, 0, compositeCanvas.width, compositeCanvas.height);
  ctx.restore();
  return new Promise((resolve, reject) => {
    compositeCanvas.toBlob((nextBlob) => {
      if (nextBlob) {
        resolve(nextBlob);
        return;
      }
      reject(new Error("Failed to encode edited image"));
    }, "image/png");
  });
}

async function saveImageEdit(options = {}) {
  const { closeAfterSave = true } = options;
  if (!state.maskEditor.active || !state.previewPath || !previewImageEditCanvas.width || !previewImageEditCanvas.height) {
    return;
  }
  if (!state.maskEditor.dirty) {
    if (closeAfterSave) {
      closeMaskEditor();
    }
    return;
  }

  state.maskEditor.saving = true;
  renderMaskEditorUi();
  statusBar.textContent = "Saving image...";
  try {
    const blob = await composeEditedImageBlob();
    const targetPath = state.maskEditor.path || state.previewPath;
    const formData = new FormData();
    formData.append("image_path", targetPath);
    formData.append("image", blob, `${getFileLabel(targetPath)}.png`);

    const resp = await fetch("/api/image/edit", {
      method: "POST",
      body: formData,
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to save image edit");
    }

    state.imageCrops[targetPath] = data.crop || null;
    state.imageVersions[targetPath] = Number(data.mtime || Date.now()) || Date.now();
    invalidateImageCaches(targetPath);
    renderGrid({ preservePath: targetPath, preserveScrollTop: fileGridContainer.scrollTop });
    closeMaskEditor();
    await showPreview(targetPath);
    statusBar.textContent = data.committed_crop
      ? `Saved image edit for ${getFileLabel(targetPath)} and committed the active crop`
      : `Saved image edit for ${getFileLabel(targetPath)}`;
  } catch (err) {
    showErrorToast(`Image edit error: ${err.message}`);
    statusBar.textContent = `Image edit error: ${err.message}`;
    throw err;
  } finally {
    state.maskEditor.saving = false;
    renderMaskEditorUi();
  }
}

async function saveMaskEdit(options = {}) {
  if (isMaskEditorImageMode()) {
    return saveImageEdit(options);
  }
  const { closeAfterSave = true } = options;
  if (!state.maskEditor.active || !state.previewPath || !previewMaskCanvas.width || !previewMaskCanvas.height) {
    return;
  }
  if (!state.maskEditor.dirty) {
    if (closeAfterSave) {
      closeMaskEditor();
    }
    return;
  }

  state.maskEditor.saving = true;
  renderMaskEditorUi();
  statusBar.textContent = "Saving mask...";
  try {
    const blob = await new Promise((resolve, reject) => {
      previewMaskCanvas.toBlob((nextBlob) => {
        if (nextBlob) {
          resolve(nextBlob);
          return;
        }
        reject(new Error("Failed to encode mask PNG"));
      }, "image/png");
    });
    const formData = new FormData();
    const targetPath = state.maskEditor.path || state.previewPath;
    if (state.maskEditor.mediaType === "video") {
      formData.append("media_path", targetPath);
      formData.append("frame_index", String(Math.max(0, Number(state.maskEditor.frameIndex || 0))));
    } else {
      formData.append("image_path", targetPath);
    }
    formData.append("mask", blob, `${getFileLabel(targetPath)}.mask.png`);

    const resp = await fetch("/api/mask", {
      method: "POST",
      body: formData,
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to save mask");
    }
    refreshMaskBaseCanvas();
    state.maskEditor.cleanHistoryIndex = state.maskEditor.historyIndex;
    syncMaskEditorDirtyState();
    if (state.maskEditor.mediaType === "video") {
      setVideoMaskKeyframes(targetPath, data.keyframes || []);
    }
    setImageMaskPresence(targetPath, true, data.mtime || Date.now(), data.mask_count);
    renderVideoEditPanel();
    statusBar.textContent = state.maskEditor.mediaType === "video"
      ? `Saved key-frame mask for ${getFileLabel(targetPath)} at ${formatVideoMaskFrameHint(state.maskEditor.frameIndex, targetPath)}`
      : `Saved mask for ${getFileLabel(targetPath)}`;
    if (closeAfterSave) {
      closeMaskEditor();
    } else {
      renderMaskEditorUi();
    }
  } catch (err) {
    showErrorToast(`Mask error: ${err.message}`);
    statusBar.textContent = `Mask error: ${err.message}`;
    throw err;
  } finally {
    state.maskEditor.saving = false;
    renderMaskEditorUi();
  }
}

function cancelMaskEdit() {
  if (!state.maskEditor.active) return;
  const imageMode = isMaskEditorImageMode();
  closeMaskEditor({ restoreBase: true });
  statusBar.textContent = imageMode ? "Image edit cancelled" : "Mask edit cancelled";
}

function resetMaskEditToDefault() {
  const activeCanvas = getActiveEditCanvas();
  if (!state.maskEditor.active || !activeCanvas.width || !activeCanvas.height) return;
  const beforeCanvas = cloneMaskCanvasSnapshot(activeCanvas);
  const fullTileKeys = getFullCanvasTileKeys(activeCanvas.width, activeCanvas.height);
  if (isMaskEditorImageMode()) {
    clearMaskStrokeRenderFrame();
    activeCanvas.getContext("2d").clearRect(0, 0, activeCanvas.width, activeCanvas.height);
    pushMaskHistorySnapshot({ beforeCanvas, tileKeys: fullTileKeys });
    syncMaskEditorDirtyState();
    scheduleMaskMiniPreviewRender();
    renderMaskEditorUi();
    statusBar.textContent = "Image edit reset";
    return;
  }
  const resetValue = clamp(Number(state.maskEditor.brushValue || 0), 0, 100);
  const resetChannelValue = Math.round(resetValue * 2.55);
  clearMaskStrokeRenderFrame();
  const ctx = previewMaskCanvas.getContext("2d");
  ctx.save();
  ctx.fillStyle = `rgb(${resetChannelValue}, ${resetChannelValue}, ${resetChannelValue})`;
  ctx.fillRect(0, 0, previewMaskCanvas.width, previewMaskCanvas.height);
  ctx.restore();
  pushMaskHistorySnapshot({ beforeCanvas, tileKeys: fullTileKeys });
  syncMaskEditorDirtyState();
  updateMaskSignalProbeStats();
  scheduleMaskMiniPreviewRender();
  scheduleMaskLatentPreviewRender();
  renderMaskEditorUi();
  statusBar.textContent = `Mask reset to ${Math.round(resetValue)}%`;
}

function toggleMaskEditorViewMode() {
  if (!isMaskEditorMaskMode()) return;
  state.maskEditor.viewMode = state.maskEditor.viewMode === "mask" ? "overlay" : "mask";
  applyMaskViewMode();
  updateMaskViewModeButton();
  statusBar.textContent = state.maskEditor.viewMode === "mask"
    ? "Showing grayscale mask view"
    : "Showing mask overlay";
}

function toggleMaskLatentPreview() {
  if (!isMaskEditorMaskMode()) return;
  state.maskEditor.latentPreviewEnabled = !state.maskEditor.latentPreviewEnabled;
  if (state.maskEditor.latentPreviewEnabled) {
    state.maskEditor.latentImageDirty = true;
    scheduleMaskLatentPreviewRender({ imageDirty: true });
    statusBar.textContent = "Showing latent-space mask preview";
  } else {
    statusBar.textContent = "Showing full-resolution mask preview";
  }
  renderMaskEditorUi();
}

function createPreviewCaptionButton(sentence) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "preview-caption-link";
  button.dataset.sentence = sentence;
  button.textContent = sentence;
  button.setAttribute("aria-label", `Jump to caption: ${sentence}`);
  button.addEventListener("click", (e) => {
    e.stopPropagation();
    jumpToSentenceInList(sentence);
  });
  return button;
}

function renderPreviewCaptionOverlay() {
  previewCaptionOverlay.classList.remove("visible");
  previewCaptionOverlay.classList.toggle("collapsed", !!state.previewCaptionOverlayCollapsed);
  previewCaptionToggle.textContent = state.previewCaptionOverlayCollapsed ? "+" : "−";
  previewCaptionToggle.setAttribute("aria-expanded", state.previewCaptionOverlayCollapsed ? "false" : "true");
  previewCaptionToggle.setAttribute("aria-label", state.previewCaptionOverlayCollapsed ? "Expand enabled captions" : "Collapse enabled captions");
  previewCaptionToggle.title = state.previewCaptionOverlayCollapsed ? "Expand enabled captions" : "Collapse enabled captions";

  const sentences = getPreviewEnabledSentences();
  if (isMaskEditorVisible() || !state.previewPath || !imgNatW || !imgNatH || !isPreviewVisible() || sentences.length === 0) {
    previewCaptionList.replaceChildren();
    return;
  }

  const panel = previewStage;
  const maxWidth = Math.max(220, panel.clientWidth - 16);

  previewCaptionOverlay.style.maxWidth = `${maxWidth}px`;

  if (!state.previewCaptionOverlayCollapsed) {
    const existingButtons = new Map(
      [...previewCaptionList.querySelectorAll(".preview-caption-link")].map((button) => [button.dataset.sentence || button.textContent || "", button])
    );
    const nextButtons = sentences.map((sentence) => {
      const existingButton = existingButtons.get(sentence);
      if (existingButton) {
        existingButton.textContent = sentence;
        existingButton.dataset.sentence = sentence;
        existingButton.setAttribute("aria-label", `Jump to caption: ${sentence}`);
        return existingButton;
      }
      return createPreviewCaptionButton(sentence);
    });
    previewCaptionList.replaceChildren(...nextButtons);
  }

  previewCaptionOverlay.classList.add("visible");
}

function togglePreviewCaptionOverlayCollapsed() {
  state.previewCaptionOverlayCollapsed = !state.previewCaptionOverlayCollapsed;
  renderPreviewCaptionOverlay();
}

// ===== THUMBNAIL PRELOADING =====
const thumbBlobCache = new Map();  // path -> blob URL
const thumbLoadQueue = [];
let thumbLoadingCount = 0;
const MAX_CONCURRENT_THUMB_LOADS = 12;
const CROP_SNAP_HYSTERESIS_PX = 10;
const thumbQueuedKeys = new Set();

function getImageVersion(path) {
  return state.imageVersions[path] || 0;
}

function bumpImageVersion(path) {
  state.imageVersions[path] = Date.now();
}

function getMaskVersion(path) {
  return state.imageMaskVersions[path] || 0;
}

function bumpMaskVersion(path, version = Date.now()) {
  state.imageMaskVersions[path] = Number(version || Date.now()) || Date.now();
}

function setImageMaskPresence(path, hasMask, version = null, maskCount = null) {
  const image = state.images.find((item) => item.path === path);
  if (image) {
    image.has_mask = !!hasMask;
    if (maskCount != null) {
      image.mask_count = Math.max(0, Number(maskCount || 0));
    } else if (!hasMask) {
      image.mask_count = 0;
    } else if (!(Number(image.mask_count || 0) > 0)) {
      image.mask_count = 1;
    }
  }
  if (hasMask) {
    bumpMaskVersion(path, version ?? Date.now());
  } else {
    delete state.imageMaskVersions[path];
  }

  if (state.activeMetaFilters.maskState !== "any") {
    const preferredPath = state.previewPath || state.lastClickedPath || [...state.selectedPaths][0] || null;
    const previousScrollTop = fileGridContainer.scrollTop;
    renderGrid({ preservePath: preferredPath, preserveScrollTop: previousScrollTop });
    return;
  }

  const cell = fileGrid.querySelector(`.thumb-cell[data-path="${CSS.escape(path)}"]`);
  if (cell) {
    cell.classList.toggle("has-mask", !!hasMask);
  }
}

function buildImageApiUrl(endpoint, path, extraParams = {}) {
  const params = new URLSearchParams({ path, v: String(getImageVersion(path)) });
  Object.entries(extraParams).forEach(([key, value]) => {
    params.set(key, String(value));
  });
  return `/api/${endpoint}?${params.toString()}`;
}

function queueThumbLoad(path, size, priority = false) {
  const version = getImageVersion(path);
  const key = `${path}:${size}:${version}`;
  if (thumbBlobCache.has(key) || thumbQueuedKeys.has(key)) return false;
  const item = { path, size, key, version };
  thumbQueuedKeys.add(key);
  if (priority) {
    thumbLoadQueue.unshift(item);
  } else {
    thumbLoadQueue.push(item);
  }
  processThumbQueue();
  return true;
}

async function processThumbQueue() {
  while (thumbLoadingCount < MAX_CONCURRENT_THUMB_LOADS && thumbLoadQueue.length > 0) {
    const { path, size, key, version } = thumbLoadQueue.shift();
    if (thumbBlobCache.has(key)) {
      thumbQueuedKeys.delete(key);
      continue;
    }
    thumbLoadingCount++;
    fetch(buildImageApiUrl("thumbnail", path, { size }))
      .then(r => r.ok ? r.blob() : null)
      .then(blob => {
        if (blob) {
          const url = URL.createObjectURL(blob);
          if (version !== getImageVersion(path)) {
            URL.revokeObjectURL(url);
            return;
          }
          thumbBlobCache.set(key, url);
          // Update any visible img with this path
          const imgs = fileGrid.querySelectorAll(`img[data-path="${CSS.escape(path)}"]`);
          const expectedKey = `${path}:${size}:${getImageVersion(path)}`;
          if (expectedKey !== key) {
            return;
          }
          imgs.forEach(img => { img.src = url; });
        }
      })
      .catch(() => {})
      .finally(() => {
        thumbQueuedKeys.delete(key);
        advanceThumbnailProgress();
        thumbLoadingCount--;
        processThumbQueue();
      });
  }
}

// Preview image preload cache (screen-quality, not full-res)
const previewCache = new Map(); // path -> blob URL
const previewLoadingVersions = new Map(); // path -> image version currently being fetched

function capturePreviewViewState() {
  const displayWidth = imgNatW > 0 && zoomLevel > 0 ? imgNatW * zoomLevel : 0;
  const displayHeight = imgNatH > 0 && zoomLevel > 0 ? imgNatH * zoomLevel : 0;
  const panelCenterX = previewStage.clientWidth / 2;
  const panelCenterY = previewStage.clientHeight / 2;
  const imageCenterX = panX + displayWidth / 2;
  const imageCenterY = panY + displayHeight / 2;
  return {
    naturalWidth: imgNatW,
    naturalHeight: imgNatH,
    zoomLevel,
    panX,
    panY,
    displayWidth,
    displayHeight,
    centerOffsetX: panelCenterX - imageCenterX,
    centerOffsetY: panelCenterY - imageCenterY,
    wasUserZoomed: userHasZoomed,
  };
}

function restorePreviewViewState(previousState = null) {
  imgNatW = previewImg.naturalWidth;
  imgNatH = previewImg.naturalHeight;
  syncMaskEditorPreviewScaleFromCurrentImage();
  if (
    previousState
    && previousState.displayWidth > 0
    && previousState.displayHeight > 0
    && imgNatW > 0
    && imgNatH > 0
  ) {
    const panel = previewStage;
    const cx = panel.clientWidth / 2;
    const cy = panel.clientHeight / 2;
    const widthZoom = previousState.displayWidth / imgNatW;
    const heightZoom = previousState.displayHeight / imgNatH;
    const targetZoom = Number.isFinite(widthZoom) && widthZoom > 0
      ? widthZoom
      : heightZoom;
    zoomLevel = Math.max(0.0001, targetZoom || 0.0001);
    const displayWidth = imgNatW * zoomLevel;
    const displayHeight = imgNatH * zoomLevel;
    const centerOffsetX = Number.isFinite(previousState.centerOffsetX) ? previousState.centerOffsetX : 0;
    const centerOffsetY = Number.isFinite(previousState.centerOffsetY) ? previousState.centerOffsetY : 0;
    const imageCenterX = cx - centerOffsetX;
    const imageCenterY = cy - centerOffsetY;
    panX = imageCenterX - displayWidth / 2;
    panY = imageCenterY - displayHeight / 2;
    userHasZoomed = !!previousState.wasUserZoomed;
    applyTransform();
  } else {
    resetZoomPan();
  }
  state.maskEditor.latentImageDirty = true;
  renderMaskEditorUi();
}

function preloadPreview(path) {
  const version = getImageVersion(path);
  if (previewCache.has(path) || previewLoadingVersions.get(path) === version) return;
  previewLoadingVersions.set(path, version);
  fetch(buildImageApiUrl("preview", path))
    .then(r => r.ok ? r.blob() : null)
    .then(blob => {
      if (!blob) return;
      const url = URL.createObjectURL(blob);
      if (version !== getImageVersion(path)) {
        URL.revokeObjectURL(url);
        return;
      }
      if (previewCache.has(path)) {
        URL.revokeObjectURL(previewCache.get(path));
      }
      previewCache.set(path, url);
      if (state.previewPath === path) {
        if (state.previewMediaType === "video") {
          previewVideo.poster = url;
          return;
        }
        if (state.promptPreview.sourcePath === path && state.promptPreview.displayPath && state.promptPreview.displayPath !== path) {
          return;
        }
        const previousViewState = capturePreviewViewState();
        previewImg.onload = () => {
          restorePreviewViewState(previousViewState);
          renderPreviewCaptionOverlay();
        };
        previewImg.src = url;
      }
    })
    .catch(() => {})
    .finally(() => {
      if (previewLoadingVersions.get(path) === version) {
        previewLoadingVersions.delete(path);
      }
    });
}

function updateFileCountDisplay() {
  const total = state.images.length;
  if (total === 0) {
    fileCount.textContent = "No media files";
    return;
  }
  const captioned = state.images.filter(img => img.has_caption).length;
  const visible = getVisibleImageEntries().length;
  const filterCount = getActiveFilterCount();
  const hasCaptionFilters = getActiveSentenceFilterEntries().length > 0;
  if (filterCount > 0 && (!hasCaptionFilters || canApplyActiveSentenceFilters())) {
    fileCount.textContent = `${visible}/${total} media files • ${captioned} captioned • ${filterCount} filter${filterCount === 1 ? "" : "s"}`;
    return;
  }
  fileCount.textContent = `${total} media files • ${captioned} captioned`;
}

function imageConformsToAspectRatios(image, tolerance = 5) {
  const dims = state.thumbnailDimensions[image?.path] || null;
  const width = Number(dims?.width || 0);
  const height = Number(dims?.height || 0);
  if (!width || !height || !state.cropAspectRatios.length) return true;
  return state.cropAspectRatios.some((ratio) => {
    const expectedWidth = height * ratio.value;
    const expectedHeight = width / ratio.value;
    return Math.abs(width - expectedWidth) <= tolerance || Math.abs(height - expectedHeight) <= tolerance;
  });
}

function refreshAspectWarning(path) {
  const image = state.images.find((item) => item.path === path);
  const cell = fileGrid.querySelector(`.thumb-cell[data-path="${CSS.escape(path)}"]`);
  if (!image || !cell) return;
  cell.classList.toggle("aspect-mismatch", !imageConformsToAspectRatios(image));
}

function storeThumbnailDimensions(path, width, height) {
  if (!path || !width || !height) return;
  const image = state.images.find((item) => item.path === path) || null;
  const previousDims = state.thumbnailDimensions[path] || null;
  const previousWidth = Number(previousDims?.width || 0);
  const previousHeight = Number(previousDims?.height || 0);
  if (previousWidth === width && previousHeight === height) {
    refreshAspectWarning(path);
    return;
  }

  const previousAspectMatch = image ? imageMatchesMetaFilters(image) : null;
  state.thumbnailDimensions[path] = { width, height };
  refreshAspectWarning(path);

  if (state.activeMetaFilters.aspectState !== "any" && image) {
    const nextAspectMatch = imageMatchesMetaFilters(image);
    if (previousAspectMatch !== nextAspectMatch) {
      const preferredPath = state.previewPath || state.lastClickedPath || [...state.selectedPaths][0] || null;
      const previousScrollTop = fileGridContainer.scrollTop;
      renderGrid({ preservePath: preferredPath, preserveScrollTop: previousScrollTop });
    }
  }
}

function updateImageDimensions(path, cropState) {
  const image = state.images.find((item) => item.path === path);
  if (!image || !cropState) return;
  if (cropState.current_width) image.width = cropState.current_width;
  if (cropState.current_height) image.height = cropState.current_height;
}

function parseAspectRatioEntry(label) {
  const text = String(label || "").trim();
  const match = text.match(/^(\d+(?:\.\d+)?)\s*[:/]\s*(\d+(?:\.\d+)?)$/);
  if (!match) return null;
  const w = Number(match[1]);
  const h = Number(match[2]);
  if (!w || !h) return null;
  return { label: `${match[1]}:${match[2]}`, value: w / h };
}

function setCropAspectRatios(labels) {
  const parsed = (labels || []).map(parseAspectRatioEntry).filter(Boolean);
  const fallback = ["4:3", "16:9", "3:4", "1:1"].map(parseAspectRatioEntry).filter(Boolean);
  state.cropAspectRatios = parsed.length ? parsed : fallback;
  state.cropAspectRatioLabels = state.cropAspectRatios.map(r => r.label);
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

const EDITOR_MAX_WORKING_EDGE = 2048;
const EDITOR_HISTORY_TILE_SIZE = 256;

function getCappedEditorDimensions(sourceWidth, sourceHeight) {
  const width = Math.max(1, Number(sourceWidth || 1));
  const height = Math.max(1, Number(sourceHeight || 1));
  const scale = Math.min(1, EDITOR_MAX_WORKING_EDGE / Math.max(width, height));
  return {
    width: Math.max(1, Math.round(width * scale)),
    height: Math.max(1, Math.round(height * scale)),
    scale,
  };
}

function getEditorTileKey(tileX, tileY) {
  return `${tileX}:${tileY}`;
}

function parseEditorTileKey(tileKey) {
  const [tileXText, tileYText] = String(tileKey || "").split(":");
  return {
    tileX: Number.parseInt(tileXText, 10) || 0,
    tileY: Number.parseInt(tileYText, 10) || 0,
  };
}

function getEditorTileRects(tileKeys, canvasWidth, canvasHeight) {
  if (!tileKeys || !tileKeys.size || !canvasWidth || !canvasHeight) {
    return [];
  }
  const tileRects = [];
  for (const tileKey of tileKeys) {
    const { tileX, tileY } = parseEditorTileKey(tileKey);
    const left = tileX * EDITOR_HISTORY_TILE_SIZE;
    const top = tileY * EDITOR_HISTORY_TILE_SIZE;
    if (left >= canvasWidth || top >= canvasHeight) {
      continue;
    }
    tileRects.push({
      left,
      top,
      width: Math.min(EDITOR_HISTORY_TILE_SIZE, canvasWidth - left),
      height: Math.min(EDITOR_HISTORY_TILE_SIZE, canvasHeight - top),
    });
  }
  tileRects.sort((leftTile, rightTile) => (leftTile.top - rightTile.top) || (leftTile.left - rightTile.left));
  return tileRects;
}

function getFullCanvasTileKeys(canvasWidth, canvasHeight) {
  const keys = new Set();
  const maxTileX = Math.max(0, Math.ceil(Math.max(1, canvasWidth) / EDITOR_HISTORY_TILE_SIZE) - 1);
  const maxTileY = Math.max(0, Math.ceil(Math.max(1, canvasHeight) / EDITOR_HISTORY_TILE_SIZE) - 1);
  for (let tileY = 0; tileY <= maxTileY; tileY += 1) {
    for (let tileX = 0; tileX <= maxTileX; tileX += 1) {
      keys.add(getEditorTileKey(tileX, tileY));
    }
  }
  return keys;
}

function areImageDataEqual(leftImageData, rightImageData) {
  const leftData = leftImageData?.data;
  const rightData = rightImageData?.data;
  if (!leftData || !rightData || leftData.length !== rightData.length) {
    return false;
  }
  for (let index = 0; index < leftData.length; index += 1) {
    if (leftData[index] !== rightData[index]) {
      return false;
    }
  }
  return true;
}

function invalidateImageCaches(path) {
  if (previewCache.has(path)) {
    URL.revokeObjectURL(previewCache.get(path));
    previewCache.delete(path);
  }
  const timelineCache = state.videoTimelineCache[path];
  if (timelineCache?.frames) {
    timelineCache.frames.forEach((frame) => {
      if (frame?.objectUrl) {
        URL.revokeObjectURL(frame.objectUrl);
      }
    });
  }
  const timelineFetches = state.ui.videoTimelineFetches.get(path);
  if (timelineFetches) {
    timelineFetches.forEach((entry) => entry.controller?.abort());
    state.ui.videoTimelineFetches.delete(path);
  }
  delete state.thumbnailDimensions[path];
  delete state.videoTimelineCache[path];
  delete state.videoTimelineUi[path];
  delete state.videoMeta[path];
  delete state.imageMaskVersions[path];
  for (const key of [...thumbBlobCache.keys()]) {
    if (key.startsWith(path + ":")) {
      URL.revokeObjectURL(thumbBlobCache.get(key));
      thumbBlobCache.delete(key);
    }
  }
}

function clearCropDraft() {
  state.cropDraft = null;
  state.cropDirty = false;
  state.cropInteraction = null;
  renderCropOverlay();
  renderVideoEditPanel();
  renderMaskEditorUi();
}

function canEditCrop() {
  if (isMaskEditorVisible()) {
    return false;
  }
  if (state.selectedPaths.size !== 1 || !state.previewPath || !imgNatW || !imgNatH) {
    return false;
  }
  if (isImageMediaPath(state.previewPath)) {
    return true;
  }
  return isVideoMediaPath(state.previewPath);
}

function getCurrentCropState() {
  return state.previewPath ? (state.imageCrops[state.previewPath] || null) : null;
}

function hasAppliedCrop() {
  return !!getCurrentCropState()?.applied;
}

function getCropSaveScale() {
  const cropState = getCurrentCropState();
  const width = cropState?.current_width || imgNatW || 1;
  const height = cropState?.current_height || imgNatH || 1;
  return {
    x: width / Math.max(1, imgNatW),
    y: height / Math.max(1, imgNatH),
  };
}

function buildCropPayload(crop) {
  const scale = getCropSaveScale();
  return {
    x: Math.round(crop.x * scale.x),
    y: Math.round(crop.y * scale.y),
    w: Math.max(1, Math.round(crop.w * scale.x)),
    h: Math.max(1, Math.round(crop.h * scale.y)),
    ratio: crop.ratio,
  };
}

function clearCropGuide() {
  if (!state.cropGuide) return;
  state.cropGuide = null;
  renderCropOverlay();
}

function setCropGuideToCenter() {
  if (!imgNatW || !imgNatH) return;
  state.cropGuide = {
    x: imgNatW / 2,
    y: imgNatH / 2,
  };
  renderCropOverlay();
}

function updateCropGuideFromClient(clientX, clientY) {
  if (!canEditCrop()) {
    clearCropGuide();
    return;
  }
  const panelRect = previewStage.getBoundingClientRect();
  const px = clientX - panelRect.left;
  const py = clientY - panelRect.top;
  const imgLeft = panX;
  const imgTop = panY;
  const imgRight = panX + imgNatW * zoomLevel;
  const imgBottom = panY + imgNatH * zoomLevel;
  if (px < imgLeft || px > imgRight || py < imgTop || py > imgBottom) {
    clearCropGuide();
    return;
  }
  const point = screenToImage(clientX, clientY);
  state.cropGuide = { x: point.x, y: point.y };
  renderCropOverlay();
}

function updateCropButtons() {
  const editable = canEditCrop();
  const hasDraft = !!state.cropDraft;
  const videoCrop = state.previewMediaType === "video";
  cropRemoveBtn.classList.toggle("visible", !videoCrop && editable && !hasDraft && hasAppliedCrop());
  cropCancelBtn.classList.toggle("visible", editable && hasDraft);
  cropCancelBtn.textContent = videoCrop ? "Clear Crop" : "Cancel";
  cropApplyBtn.classList.toggle("visible", !videoCrop && editable && hasDraft && state.cropDirty);
  cropApplyBtn.textContent = "Apply";
  rotateControls.classList.toggle("visible", !videoCrop && editable && !hasDraft);
  renderGifConvertButton();
}

function renderCropOverlay() {
  const crop = state.cropDraft;
  const guide = state.cropGuide;
  if (guide && canEditCrop()) {
    cropGuideV.classList.add("visible");
    cropGuideH.classList.add("visible");
    cropGuideV.style.left = `${panX + guide.x * zoomLevel}px`;
    cropGuideV.style.top = `${panY}px`;
    cropGuideV.style.height = `${imgNatH * zoomLevel}px`;
    cropGuideH.style.left = `${panX}px`;
    cropGuideH.style.top = `${panY + guide.y * zoomLevel}px`;
    cropGuideH.style.width = `${imgNatW * zoomLevel}px`;
  } else {
    cropGuideV.classList.remove("visible");
    cropGuideH.classList.remove("visible");
  }
  if (!crop || !canEditCrop()) {
    cropBox.classList.remove("active");
    updateCropButtons();
    return;
  }
  cropBox.classList.add("active");
  cropBox.style.left = (panX + crop.x * zoomLevel) + "px";
  cropBox.style.top = (panY + crop.y * zoomLevel) + "px";
  cropBox.style.width = (crop.w * zoomLevel) + "px";
  cropBox.style.height = (crop.h * zoomLevel) + "px";
  cropLabel.textContent = `${crop.ratio || "custom"} • ${crop.w}×${crop.h}`;
  updateCropButtons();
}

function screenToImage(clientX, clientY) {
  const panelRect = previewStage.getBoundingClientRect();
  const px = clientX - panelRect.left;
  const py = clientY - panelRect.top;
  return {
    x: clamp((px - panX) / zoomLevel, 0, imgNatW),
    y: clamp((py - panY) / zoomLevel, 0, imgNatH),
  };
}

function chooseClosestRatio(anchorX, anchorY, currentX, currentY, ratios, stickyChoice = null) {
  const dx = currentX - anchorX;
  const dy = currentY - anchorY;
  const signX = dx >= 0 ? 1 : -1;
  const signY = dy >= 0 ? 1 : -1;
  const targetW = Math.abs(dx);
  const targetH = Math.abs(dy);
  const maxW = signX > 0 ? (imgNatW - anchorX) : anchorX;
  const maxH = signY > 0 ? (imgNatH - anchorY) : anchorY;

  const candidates = [];
  for (const ratio of ratios) {
    const r = ratio.value;
    const widthFromX = Math.min(targetW, maxW, maxH * r);
    const heightFromX = widthFromX / r;
    const cand1 = {
      x: signX > 0 ? anchorX : anchorX - widthFromX,
      y: signY > 0 ? anchorY : anchorY - heightFromX,
      w: widthFromX,
      h: heightFromX,
      ratio: ratio.label,
    };
    const dist1 = Math.hypot((cand1.x + (signX > 0 ? cand1.w : 0)) - currentX, (cand1.y + (signY > 0 ? cand1.h : 0)) - currentY);

    const heightFromY = Math.min(targetH, maxH, maxW / r);
    const widthFromY = heightFromY * r;
    const cand2 = {
      x: signX > 0 ? anchorX : anchorX - widthFromY,
      y: signY > 0 ? anchorY : anchorY - heightFromY,
      w: widthFromY,
      h: heightFromY,
      ratio: ratio.label,
    };
    const dist2 = Math.hypot((cand2.x + (signX > 0 ? cand2.w : 0)) - currentX, (cand2.y + (signY > 0 ? cand2.h : 0)) - currentY);

    for (const [cand, dist, mode] of [[cand1, dist1, "width"], [cand2, dist2, "height"]]) {
      if (cand.w < 1 || cand.h < 1) continue;
      candidates.push({ crop: cand, dist, key: `${ratio.label}:${mode}` });
    }
  }

  if (candidates.length === 0) return null;

  const best = candidates.reduce((winner, candidate) => candidate.dist < winner.dist ? candidate : winner);
  const stickyCandidate = stickyChoice
    ? candidates.find((candidate) => candidate.key === stickyChoice)
    : null;
  const chosen = stickyCandidate && stickyCandidate.dist <= best.dist + CROP_SNAP_HYSTERESIS_PX
    ? stickyCandidate
    : best;

  return {
    crop: chosen.crop,
    stickyChoice: chosen.key,
  };
}

function setCropDraft(crop, dirty = false) {
  if (!crop) {
    clearCropDraft();
    return;
  }
  state.cropDraft = {
    x: Math.round(crop.x),
    y: Math.round(crop.y),
    w: Math.max(1, Math.round(crop.w)),
    h: Math.max(1, Math.round(crop.h)),
    ratio: crop.ratio || state.cropAspectRatioLabels[0] || "1:1",
  };
  state.cropDirty = dirty;
  renderCropOverlay();
}

function getDraftRatio() {
  if (state.cropDraft?.ratio) {
    const found = state.cropAspectRatios.find(r => r.label === state.cropDraft.ratio);
    if (found) return found;
  }
  return state.cropAspectRatios[0] || { label: "1:1", value: 1 };
}

function resizeCropFromHandle(baseCrop, handle, currentX, currentY, stickyChoice = null) {
  const ratio = getDraftRatio();
  const centerX = baseCrop.x + baseCrop.w / 2;
  const centerY = baseCrop.y + baseCrop.h / 2;
  const touchesLeft = baseCrop.x <= 0;
  const touchesRight = baseCrop.x + baseCrop.w >= imgNatW;
  const touchesTop = baseCrop.y <= 0;
  const touchesBottom = baseCrop.y + baseCrop.h >= imgNatH;
  if (["nw", "ne", "sw", "se"].includes(handle)) {
    const anchors = {
      nw: { x: baseCrop.x + baseCrop.w, y: baseCrop.y + baseCrop.h },
      ne: { x: baseCrop.x, y: baseCrop.y + baseCrop.h },
      sw: { x: baseCrop.x + baseCrop.w, y: baseCrop.y },
      se: { x: baseCrop.x, y: baseCrop.y },
    };
    return chooseClosestRatio(
      anchors[handle].x,
      anchors[handle].y,
      currentX,
      currentY,
      state.cropAspectRatios.length ? state.cropAspectRatios : [ratio],
      stickyChoice,
    );
  }

  const r = ratio.value;
  if (handle === "e") {
    const requestedWidth = Math.max(1, currentX - baseCrop.x);
    const maxWidthByRatio = 2 * Math.min(centerY, imgNatH - centerY) * r;
    if (requestedWidth > maxWidthByRatio && (touchesTop || touchesBottom)) {
      const nextX = clamp(currentX - baseCrop.w, 0, imgNatW - baseCrop.w);
      return { crop: { ...baseCrop, x: nextX }, stickyChoice };
    }
    let width = clamp(requestedWidth, 1, imgNatW - baseCrop.x);
    width = Math.min(width, maxWidthByRatio);
    const height = width / r;
    return { crop: { x: baseCrop.x, y: clamp(centerY - height / 2, 0, imgNatH - height), w: width, h: height, ratio: ratio.label }, stickyChoice };
  }
  if (handle === "w") {
    const requestedWidth = Math.max(1, (baseCrop.x + baseCrop.w) - currentX);
    const maxWidthByRatio = 2 * Math.min(centerY, imgNatH - centerY) * r;
    if (requestedWidth > maxWidthByRatio && (touchesTop || touchesBottom)) {
      const nextX = clamp(currentX, 0, imgNatW - baseCrop.w);
      return { crop: { ...baseCrop, x: nextX }, stickyChoice };
    }
    let width = clamp(requestedWidth, 1, baseCrop.x + baseCrop.w);
    width = Math.min(width, maxWidthByRatio);
    const height = width / r;
    return { crop: { x: (baseCrop.x + baseCrop.w) - width, y: clamp(centerY - height / 2, 0, imgNatH - height), w: width, h: height, ratio: ratio.label }, stickyChoice };
  }
  if (handle === "s") {
    const requestedHeight = Math.max(1, currentY - baseCrop.y);
    const maxHeightByRatio = (2 * Math.min(centerX, imgNatW - centerX)) / r;
    if (requestedHeight > maxHeightByRatio && (touchesLeft || touchesRight)) {
      const nextY = clamp(currentY - baseCrop.h, 0, imgNatH - baseCrop.h);
      return { crop: { ...baseCrop, y: nextY }, stickyChoice };
    }
    let height = clamp(requestedHeight, 1, imgNatH - baseCrop.y);
    height = Math.min(height, maxHeightByRatio);
    const width = height * r;
    return { crop: { x: clamp(centerX - width / 2, 0, imgNatW - width), y: baseCrop.y, w: width, h: height, ratio: ratio.label }, stickyChoice };
  }
  if (handle === "n") {
    const requestedHeight = Math.max(1, (baseCrop.y + baseCrop.h) - currentY);
    const maxHeightByRatio = (2 * Math.min(centerX, imgNatW - centerX)) / r;
    if (requestedHeight > maxHeightByRatio && (touchesLeft || touchesRight)) {
      const nextY = clamp(currentY, 0, imgNatH - baseCrop.h);
      return { crop: { ...baseCrop, y: nextY }, stickyChoice };
    }
    let height = clamp(requestedHeight, 1, baseCrop.y + baseCrop.h);
    height = Math.min(height, maxHeightByRatio);
    const width = height * r;
    return { crop: { x: clamp(centerX - width / 2, 0, imgNatW - width), y: (baseCrop.y + baseCrop.h) - height, w: width, h: height, ratio: ratio.label }, stickyChoice };
  }
  return null;
}

async function loadCropData(path) {
  if (!isImageMediaPath(path)) {
    state.imageCrops[path] = null;
    if (state.previewPath === path) {
      updateCropButtons();
    }
    return;
  }
  try {
    const resp = await fetch(`/api/crop?path=${encodeURIComponent(path)}`);
    if (!resp.ok) throw new Error("Failed to load crop");
    const data = await resp.json();
    state.imageCrops[path] = data.crop || null;
    if (state.previewPath === path) {
      updateCropButtons();
    }
  } catch (err) {
    console.error("Failed to load crop:", err);
  }
}

function startCropCreate(event) {
  const start = screenToImage(event.clientX, event.clientY);
  state.cropInteraction = { mode: "create", anchor: start, stickySnap: null };
  const ratio = state.cropAspectRatios[0] || { label: "1:1", value: 1 };
  setCropDraft({ x: start.x, y: start.y, w: 1, h: 1, ratio: ratio.label }, true);
}

function startCropResize(handle, event) {
  if (!state.cropDraft) return;
  event.preventDefault();
  event.stopPropagation();
  state.cropInteraction = { mode: "resize", handle, baseCrop: { ...state.cropDraft }, stickySnap: null };
}

async function applyCropDraft() {
  return applyCropDraftWithOptions();
}

async function applyCropDraftWithOptions(options = {}) {
  if (!state.previewPath || !state.cropDraft || !state.cropDirty) return;
  const { reopenPreview = true } = options;
  const path = state.previewPath;
  if (isVideoMediaPath(path)) {
    state.cropDirty = false;
    renderCropOverlay();
    renderVideoEditPanel();
    statusBar.textContent = "Crop selection updated";
    return;
  }
  statusBar.textContent = "Applying crop...";
  try {
    const resp = await fetch("/api/crop", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ image_path: path, crop: buildCropPayload(state.cropDraft) }),
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.detail || "Failed to apply crop");
    state.imageCrops[path] = data.crop || null;
    updateImageDimensions(path, data.crop);
    clearCropDraft();
    bumpImageVersion(path);
    invalidateImageCaches(path);
    renderGrid();
    userHasZoomed = false;
    zoomLevel = 1;
    panX = 0;
    panY = 0;
    if (reopenPreview) {
      showPreview(path);
    }
    statusBar.textContent = "Crop applied";
  } catch (err) {
    statusBar.textContent = `Crop error: ${err.message}`;
  }
}

function cancelCropEdit() {
  if (!state.cropDraft && !state.cropInteraction) return;
  clearCropDraft();
  statusBar.textContent = "Crop edit cancelled";
}

async function rotatePreviewImage(direction) {
  if (!state.previewPath) return;
  const path = state.previewPath;
  statusBar.textContent = direction === "left" ? "Rotating left..." : "Rotating right...";
  try {
    const resp = await fetch("/api/rotate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ image_path: path, direction }),
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.detail || "Failed to rotate image");
    state.imageCrops[path] = data.crop || null;
    updateImageDimensions(path, data.crop);
    clearCropDraft();
    bumpImageVersion(path);
    invalidateImageCaches(path);
    renderGrid();
    userHasZoomed = false;
    zoomLevel = 1;
    panX = 0;
    panY = 0;
    await showPreview(path);
    statusBar.textContent = direction === "left" ? "Rotated left" : "Rotated right";
  } catch (err) {
    statusBar.textContent = `Rotate error: ${err.message}`;
  }
}

async function removeCrop() {
  if (!state.previewPath || !hasAppliedCrop()) return;
  const path = state.previewPath;
  statusBar.textContent = "Removing crop...";
  try {
    const resp = await fetch("/api/crop", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ image_path: path, crop: null }),
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.detail || "Failed to remove crop");
    state.imageCrops[path] = data.crop || null;
    updateImageDimensions(path, data.crop);
    clearCropDraft();
    bumpImageVersion(path);
    invalidateImageCaches(path);
    renderGrid();
    userHasZoomed = false;
    zoomLevel = 1;
    panX = 0;
    panY = 0;
    showPreview(path);
    statusBar.textContent = "Crop removed";
  } catch (err) {
    statusBar.textContent = `Crop error: ${err.message}`;
  }
}

function updateActionButtons() {
  const hasSelection = state.selectedPaths.size > 0;
  const hasCaptions = hasConfiguredCaptions();
  const selectionSupportsVision = hasSelection;
  const canRunStructured = selectionSupportsVision && hasCaptions && !!state.ollamaModel.trim();
  const canRunFreeTextOnly = selectionSupportsVision && !!state.ollamaModel.trim();
  cloneFolderBtn.disabled = !state.folder || state.autoCaptioning || state.cloning || state.uploading;
  cloneFolderBtn.textContent = state.cloning ? "Cloning..." : "Clone";
  cloneFolderBtn.title = state.selectedPaths.size > 1
    ? "Clone the selected media files into a new sibling folder"
    : "Clone the whole current folder into a new sibling folder";
  autoCaptionBtn.disabled = state.uploading || (!state.autoCaptioning && !canRunStructured);
  addFreeTextNowBtn.disabled = state.uploading || (!state.autoCaptioning && !canRunFreeTextOnly);
  autoCaptionBtn.classList.toggle("running", state.autoCaptioning);
  addFreeTextNowBtn.classList.toggle("running", state.autoCaptioning && state.autoCaptionMode === "free-text-only");
  setGenerateButtonContent(autoCaptionBtn, state.autoCaptioning && state.autoCaptionMode === "full" ? "Stop Auto Caption" : "Auto Caption");
  setGenerateButtonContent(
    addFreeTextNowBtn,
    state.autoCaptioning && state.autoCaptionMode === "free-text-only"
      ? "Stop free-text enhancement"
      : "Ask Ollama to add only free-text details for the selected media",
    { iconOnly: true }
  );
  autoCaptionBtn.title = state.autoCaptioning && state.autoCaptionMode === "full"
    ? "Stop the running auto caption operation"
    : "Use Ollama to verify captions for the selected media";
  addFreeTextNowBtn.title = state.autoCaptioning && state.autoCaptionMode === "free-text-only"
    ? "Stop the running free-text enhancement"
    : "Ask Ollama to add only free-text details for the selected media";
  renderCreatePromptPreviewButton();
  renderPreviewActionBar();
  renderVideoEditPanel();
}

function resetAutoCaptionProgress() {
  state.aiProgress = {
    visible: false,
    scopeLabel: "",
    totalImages: 0,
    processedImages: 0,
    errors: 0,
    completedImages: 0,
    enableFreeText: false,
    freeTextOnly: false,
    currentPath: "",
    currentMessage: "",
    currentStepIndex: 0,
    currentStepTotal: 0,
  };
}

function updateAutoCaptionProgress(patch = {}) {
  state.aiProgress = { ...state.aiProgress, ...patch };
  renderModelLog();
}

function renderToolbarStatusVisibility() {
  const hasActiveUpload = !!(state.uploading || state.uploadQueueCurrentJob || state.uploadQueue.length);
  const hasVideoJobs = !!(state.videoJobs?.activeJob || state.videoJobs?.queuedJobs?.length);
  const hasVisibleProgress = !!state.thumbnailProgress.visible || hasActiveUpload || hasVideoJobs || !!state.aiProgress.visible;
  statusBar.hidden = hasVisibleProgress;
}

function renderAutoCaptionProgress() {
  const progress = state.aiProgress;
  aiProgressPanel.classList.toggle("visible", !!progress.visible);
  renderToolbarStatusVisibility();
  if (!progress.visible) {
    aiProgressSummary.textContent = "";
    aiProgressMetric.textContent = "";
    aiProgressCurrentLabel.textContent = "";
    aiProgressCurrentMetric.textContent = "";
    aiProgressOverallFill.style.width = "0%";
    aiProgressCurrentFill.style.width = "0%";
    return;
  }

  const totalImages = Math.max(1, Number(progress.totalImages || 0));
  const completedImages = Math.max(0, Math.min(totalImages, Number(progress.completedImages || 0)));
  const currentStepTotal = Math.max(0, Number(progress.currentStepTotal || 0));
  const currentStepIndex = Math.max(0, Math.min(currentStepTotal || 0, Number(progress.currentStepIndex || 0)));
  const currentFraction = currentStepTotal > 0 ? currentStepIndex / currentStepTotal : 0;
  const overallPercent = Math.max(0, Math.min(100, ((completedImages + currentFraction) / totalImages) * 100));
  const currentPercent = Math.max(0, Math.min(100, currentStepTotal > 0 ? (currentStepIndex / currentStepTotal) * 100 : 0));

  aiProgressSummary.textContent = `${progress.scopeLabel || "AI"} • ${progress.processedImages}/${progress.totalImages} done${progress.errors ? ` • ${progress.errors} error${progress.errors === 1 ? "" : "s"}` : ""}`;
  aiProgressMetric.textContent = `${Math.round(overallPercent)}%`;
  aiProgressCurrentLabel.textContent = progress.currentPath
    ? `${getFileLabel(progress.currentPath)}${progress.currentMessage ? ` • ${progress.currentMessage}` : ""}`
    : (progress.currentMessage || "Preparing...");
  aiProgressCurrentMetric.textContent = currentStepTotal > 0 ? `${currentStepIndex}/${currentStepTotal}` : "";
  aiProgressOverallFill.style.width = `${overallPercent}%`;
  aiProgressCurrentFill.style.width = `${currentPercent}%`;
}

function renderModelLog() {
  const hasLogUi = state.modelLogLines.length > 0 || !!state.aiProgress.visible;
  renderAutoCaptionProgress();
  modelLogOpenBtn.hidden = !hasLogUi;
  modelLogOpenBtn.disabled = !hasLogUi;
  modelLogOpenBtn.classList.toggle("active", !!state.modelLogOpen && hasLogUi);
  modelLogOpenBtn.textContent = state.modelLogOpen ? "Hide Log" : "Log";
  modelLogOpenBtn.setAttribute("aria-expanded", state.modelLogOpen && hasLogUi ? "true" : "false");
  if (!hasLogUi) {
    state.modelLogOpen = false;
    modelLog.textContent = "";
    modelLogOverlay.classList.remove("open");
    modelLogOverlay.setAttribute("aria-hidden", "true");
    return;
  }
  modelLogOverlay.classList.toggle("open", !!state.modelLogOpen);
  modelLogOverlay.setAttribute("aria-hidden", state.modelLogOpen ? "false" : "true");
  modelLog.innerHTML = state.modelLogLines.join("\n");
  if (state.modelLogOpen) {
    modelLog.scrollTop = modelLog.scrollHeight;
  }
}

function renderAddButtonsVisibility() {
  hideAddButtonsCheckbox.checked = !!state.hideAddButtons;
}

function clearModelLog() {
  state.modelLogLines = [];
  renderModelLog();
}

function toggleModelLogOverlay() {
  if (state.modelLogLines.length === 0 && !state.aiProgress.visible) return;
  state.modelLogOpen = !state.modelLogOpen;
  renderModelLog();
}

function closeModelLogOverlay() {
  if (!state.modelLogOpen) return;
  state.modelLogOpen = false;
  renderModelLog();
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function appendModelLog(text, kind = "") {
  const safe = escapeHtml(text);
  state.modelLogLines.push(kind ? `<span class="${kind}">${safe}</span>` : safe);
  if (state.modelLogLines.length > 400) {
    state.modelLogLines = state.modelLogLines.slice(-400);
  }
  renderModelLog();
}

function slugifyVideoTrainingPresetKey(value, fallback = "profile") {
  const raw = String(value || fallback).trim().toLowerCase();
  let result = "";
  let previousDash = false;
  for (const char of raw) {
    if ((char >= "a" && char <= "z") || (char >= "0" && char <= "9")) {
      result += char;
      previousDash = false;
      continue;
    }
    if (previousDash) continue;
    result += "-";
    previousDash = true;
  }
  return result.replace(/^-+|-+$/g, "") || fallback;
}

function roundVideoTrainingSeconds(value) {
  return Math.round(Math.max(0, Number(value || 0)) * 1000) / 1000;
}

function normalizeEditableVideoTrainingPreset(rawPreset, index, usedKeys) {
  if (!rawPreset || typeof rawPreset !== "object" || Array.isArray(rawPreset)) return null;

  const fallbackKey = `profile-${index + 1}`;
  const baseKey = slugifyVideoTrainingPresetKey(
    rawPreset.key || rawPreset.label || rawPreset.name,
    fallbackKey
  );
  let key = baseKey;
  let counter = 2;
  while (usedKeys.has(key)) {
    key = `${baseKey}-${counter}`;
    counter += 1;
  }
  usedKeys.add(key);

  const label = String(rawPreset.label || rawPreset.name || key).trim() || key;
  let targetFamily = String(rawPreset.target_family || "custom").trim().toLowerCase() || "custom";
  if (!["wan", "ltx", "hunyuan", "custom"].includes(targetFamily)) {
    targetFamily = "custom";
  }

  const numFrames = Math.max(1, Number.parseInt(rawPreset.num_frames, 10) || 1);
  const fps = Math.max(1, Number.parseInt(rawPreset.fps, 10) || 16);
  const shortClipFactor = Math.max(0.25, Math.min(1.0, Number(rawPreset.short_clip_factor || 0.8) || 0.8));
  const longClipFactor = Math.max(shortClipFactor, Math.min(4.0, Number(rawPreset.long_clip_factor || 1.5) || 1.5));
  const preferredExtensions = [];
  for (const extension of Array.isArray(rawPreset.preferred_extensions) ? rawPreset.preferred_extensions : []) {
    let normalizedExtension = String(extension || "").trim().toLowerCase();
    if (!normalizedExtension) continue;
    if (!normalizedExtension.startsWith(".")) {
      normalizedExtension = `.${normalizedExtension}`;
    }
    if (!VIDEO_FILE_EXTENSIONS.has(normalizedExtension) || preferredExtensions.includes(normalizedExtension)) continue;
    preferredExtensions.push(normalizedExtension);
  }

  const idealClipSeconds = roundVideoTrainingSeconds(numFrames / Math.max(1, fps));
  return {
    key,
    label,
    target_family: targetFamily,
    num_frames: numFrames,
    fps,
    shrink_video_to_frames: rawPreset.shrink_video_to_frames !== false,
    short_clip_factor: shortClipFactor,
    long_clip_factor: longClipFactor,
    ideal_clip_seconds: idealClipSeconds,
    min_clip_seconds: roundVideoTrainingSeconds(idealClipSeconds * shortClipFactor),
    max_clip_seconds: roundVideoTrainingSeconds(idealClipSeconds * longClipFactor),
    preferred_extensions: preferredExtensions,
    description: String(rawPreset.description || "").trim(),
  };
}

function normalizeEditableVideoTrainingPresets(rawPresets) {
  const normalized = [];
  const usedKeys = new Set();
  for (const [index, rawPreset] of (Array.isArray(rawPresets) ? rawPresets : []).entries()) {
    const preset = normalizeEditableVideoTrainingPreset(rawPreset, index, usedKeys);
    if (preset) {
      normalized.push(preset);
    }
  }
  return normalized;
}

function parseVideoTrainingPresetsText(rawText) {
  let parsed;
  try {
    parsed = JSON.parse(String(rawText || "").trim());
  } catch {
    throw new Error("Video training presets must be valid JSON.");
  }
  if (!Array.isArray(parsed)) {
    throw new Error("Video training presets JSON must be an array.");
  }
  const normalized = normalizeEditableVideoTrainingPresets(parsed);
  if (!normalized.length) {
    throw new Error("Video training presets JSON must contain at least one preset object.");
  }
  return normalized;
}

function getDefaultVideoTrainingPresetsStatus() {
  return state.folder
    ? "Edit the preset library as JSON. Save to apply the library and the current folder selection."
    : "Edit the preset library as JSON. Load a folder to choose a per-folder profile.";
}

function setVideoTrainingPresetsStatus(message = "", options = {}) {
  const { isError = false } = options;
  settingsVideoPresetsStatus.textContent = String(message || getDefaultVideoTrainingPresetsStatus()).trim();
  settingsVideoPresetsStatus.classList.toggle("is-error", !!isError);
}

function findVideoTrainingProfileByKey(key, presets = state.videoTrainingPresets) {
  const normalizedKey = String(key || "").trim();
  if (!normalizedKey) return null;
  return (Array.isArray(presets) ? presets : []).find((preset) => String(preset?.key || "").trim() === normalizedKey) || null;
}

function populateVideoTrainingProfileSelect(presets = state.videoTrainingPresets, selectedKey = state.videoTrainingProfileKey) {
  const profileList = Array.isArray(presets) ? presets : [];
  const options = [];
  if (state.folder && profileList.length > 0) {
    const resolvedProfile = findVideoTrainingProfileByKey(selectedKey, profileList) || profileList[0] || null;
    for (const preset of profileList) {
      options.push({
        value: String(preset.key || ""),
        label: `${preset.label} (${preset.num_frames}f @ ${preset.fps} fps)`,
      });
    }
    settingsVideoProfileInput.innerHTML = "";
    for (const optionData of options) {
      const option = document.createElement("option");
      option.value = optionData.value;
      option.textContent = optionData.label;
      settingsVideoProfileInput.appendChild(option);
    }
    settingsVideoProfileInput.disabled = false;
    settingsVideoProfileInput.value = resolvedProfile ? resolvedProfile.key : "";
    return;
  }

  settingsVideoProfileInput.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = state.folder ? "No presets configured" : "Load a folder to choose a profile";
  settingsVideoProfileInput.appendChild(placeholder);
  settingsVideoProfileInput.value = "";
  settingsVideoProfileInput.disabled = true;
}

function getSelectedVideoTrainingProfileFromState() {
  return findVideoTrainingProfileByKey(state.videoTrainingProfileKey)
    || findVideoTrainingProfileByKey(state.videoTrainingProfile?.key)
    || state.videoTrainingProfile
    || (Array.isArray(state.videoTrainingPresets) && state.videoTrainingPresets.length > 0 ? state.videoTrainingPresets[0] : null);
}

function getVideoTrainingPreferredExtensionsLabel(profile) {
  const extensions = Array.isArray(profile?.preferred_extensions)
    ? profile.preferred_extensions.map((extension) => String(extension || "").trim()).filter(Boolean)
    : [];
  return extensions.join(", ");
}

function renderVideoTrainingSummary() {
  const visible = state.previewMediaType === "video" && !!state.previewPath && state.selectedPaths.size === 1;
  if (!visible) {
    videoTrainingProfileLabel.textContent = "";
    videoTrainingGuidanceLabel.textContent = "";
    return;
  }

  const profile = getSelectedVideoTrainingProfileFromState();
  if (!profile) {
    videoTrainingProfileLabel.textContent = "No folder video profile selected";
    videoTrainingGuidanceLabel.textContent = "Open Settings > Video Training and choose a profile for this folder.";
    return;
  }

  const duration = getEffectiveVideoDuration(state.previewPath);
  const draft = getVideoClipDraft(state.previewPath);
  const selectionDuration = duration > 0
    ? Math.max(0, duration * Math.max(0, Number(draft?.endFraction || 0) - Number(draft?.startFraction || 0)))
    : 0;
  const minSeconds = Math.max(0, Number(profile.min_clip_seconds || 0));
  const maxSeconds = Math.max(minSeconds, Number(profile.max_clip_seconds || 0));
  const idealSeconds = Math.max(0, Number(profile.ideal_clip_seconds || 0));
  let selectionState = "";
  if (duration > 0) {
    if (selectionDuration < minSeconds) {
      selectionState = "short";
    } else if (selectionDuration > maxSeconds) {
      selectionState = "long";
    } else {
      selectionState = "within target";
    }
  }

  const guidance = [
    `Target ${formatDurationSeconds(idealSeconds)}`,
    `recommended ${formatDurationSeconds(minSeconds)}-${formatDurationSeconds(maxSeconds)}`,
  ];
  if (duration > 0) {
    guidance.push(`selection ${formatDurationSeconds(selectionDuration)}${selectionState ? ` (${selectionState})` : ""}`);
  } else {
    guidance.push("loading duration...");
  }
  const preferredExtensions = getVideoTrainingPreferredExtensionsLabel(profile);
  if (preferredExtensions) {
    guidance.push(`prefer ${preferredExtensions}`);
  }
  if (profile.description) {
    guidance.push(profile.description);
  }

  videoTrainingProfileLabel.textContent = `${profile.label} • ${profile.num_frames}f @ ${profile.fps} fps`;
  videoTrainingGuidanceLabel.textContent = guidance.join(" • ");
}

function isGifMediaPath(path = state.previewPath) {
  return getImageExtension(path) === ".gif";
}

function renderGifConvertButton() {
  const visible = !!state.previewPath
    && state.previewMediaType === "image"
    && state.selectedPaths.size === 1
    && isGifMediaPath(state.previewPath)
    && !state.maskEditor.active;
  const hasVideoJobs = !!(state.videoJobs?.activeJob || state.videoJobs?.queuedJobs?.length);
  gifConvertBtn.classList.toggle("visible", visible);
  gifConvertBtn.disabled = !visible || !!state.cropDraft || state.autoCaptioning || state.cloning || state.uploading || hasVideoJobs;
  gifConvertBtn.title = hasVideoJobs
    ? "Wait for queued video jobs to finish before converting this GIF"
    : "Convert this GIF into an MP4 beside the original GIF";
}

function createPromptPreviewSummary() {
  return {
    total: 0,
    spawned: 0,
    queued: 0,
    running: 0,
    completed: 0,
    failed: 0,
    latest_prompt_id: "",
    latest_output_path: "",
  };
}

function isPromptPreviewSourceActive(sourcePath = state.previewPath) {
  return !!sourcePath && state.promptPreview.sourcePath === sourcePath;
}

function getPromptPreviewFiles(sourcePath = state.previewPath) {
  if (!isPromptPreviewSourceActive(sourcePath)) return [];
  return Array.isArray(state.promptPreview.files) ? state.promptPreview.files.filter(Boolean) : [];
}

function getLatestPromptPreviewFile(sourcePath = state.previewPath, files = getPromptPreviewFiles(sourcePath)) {
  const uniqueFiles = Array.from(new Set((files || []).filter(Boolean)));
  return uniqueFiles.length > 0 ? uniqueFiles[uniqueFiles.length - 1] : "";
}

function getPromptPreviewCyclePaths(sourcePath = state.previewPath, files = getPromptPreviewFiles(sourcePath)) {
  const latestFile = getLatestPromptPreviewFile(sourcePath, files);
  if (!sourcePath) {
    return latestFile ? [latestFile] : [];
  }
  return latestFile ? [sourcePath, latestFile] : [sourcePath];
}

function getPromptPreviewCurrentDisplayPath(sourcePath = state.previewPath) {
  if (isPromptPreviewSourceActive(sourcePath) && state.promptPreview.displayPath) {
    return state.promptPreview.displayPath;
  }
  return sourcePath;
}

function resetPromptPreviewState() {
  state.promptPreview.sourcePath = "";
  state.promptPreview.jobs = [];
  state.promptPreview.summary = createPromptPreviewSummary();
  state.promptPreview.files = [];
  state.promptPreview.displayPath = "";
  state.promptPreview.cycleIndex = -1;
  state.promptPreview.lastFilesKey = "";
  state.promptPreview.loading = false;
  renderPreviewInfo();
  renderPromptPreviewControls();
}

function renderPreviewInfo() {
  if (!state.previewPath) {
    previewInfo.style.display = "none";
    return;
  }

  const media = state.images.find((item) => item.path === state.previewPath);
  if (!media) {
    previewInfo.style.display = "none";
    return;
  }

  const parts = [];
  if (state.previewMediaType === "video") {
    parts.push(`${media.name} • video`);
  } else {
    parts.push(hasAppliedCrop() ? `${media.name} • cropped` : media.name);
  }

  const latestPromptPreviewPath = getLatestPromptPreviewFile(state.previewPath);
  if (latestPromptPreviewPath) {
    parts.push(state.promptPreview.displayPath === latestPromptPreviewPath ? "prompt preview" : "original");
  }

  previewInfo.textContent = parts.join(" • ");
  previewInfo.style.display = "block";
}

function loadPromptPreviewImage(path, { preserveView = true } = {}) {
  if (!path) return;
  if (!(state.imageVersions[path] > 0)) {
    state.imageVersions[path] = Date.now();
  }
  const previousViewState = preserveView ? capturePreviewViewState() : null;
  previewImg.style.display = "none";
  previewImg.onload = () => {
    restorePreviewViewState(previousViewState);
    renderPreviewCaptionOverlay();
    renderPreviewInfo();
  };
  previewImg.src = buildImageApiUrl("preview", path);
  previewPlaceholder.style.display = "none";
}

async function setPromptPreviewDisplayPath(path, options = {}) {
  const { preserveView = true } = options;
  if (!state.previewPath || state.previewMediaType !== "image") return;
  const sourcePath = state.previewPath;
  state.promptPreview.sourcePath = sourcePath;

  if (!path || path === sourcePath) {
    state.promptPreview.displayPath = "";
    state.promptPreview.cycleIndex = getPromptPreviewCyclePaths(sourcePath).indexOf(sourcePath);
    await showPreview(sourcePath, { preserveView });
    renderPreviewInfo();
    renderPromptPreviewControls();
    return;
  }

  state.promptPreview.displayPath = path;
  state.promptPreview.cycleIndex = getPromptPreviewCyclePaths(sourcePath).indexOf(path);
  loadPromptPreviewImage(path, { preserveView });
  renderPreviewInfo();
  renderPromptPreviewControls();
}

async function clearPromptPreviewDisplay(options = {}) {
  if (!state.previewPath || state.promptPreview.displayPath === "") return;
  await setPromptPreviewDisplayPath("", options);
}

function canPromptPreviewCurrentImage() {
  return state.selectedPaths.size === 1
    && !!state.previewPath
    && state.previewMediaType === "image"
    && !!imgNatW
    && !!imgNatH;
}

function hasPromptPreviewActiveJobs(summary = state.promptPreview.summary) {
  return Number(summary?.running || 0) > 0 || Number(summary?.queued || 0) > 0;
}

function getPromptPreviewConfigError(options = {}) {
  const { requireWorkflow = true } = options;
  if (!String(state.comfyuiOutputFolder || "").trim()) {
    return "Set a ComfyUI output folder in Settings first.";
  }
  if (requireWorkflow && !String(state.comfyuiWorkflowPath || "").trim()) {
    return "Set a ComfyUI workflow API JSON path in Settings first.";
  }
  return "";
}

function renderCreatePromptPreviewButton() {
  if (!createPromptPreviewBtn) return;
  const selectionReady = canPromptPreviewCurrentImage();
  const sourcePath = state.previewPath;
  const sourceActive = isPromptPreviewSourceActive(sourcePath);
  const summary = sourceActive ? state.promptPreview.summary : createPromptPreviewSummary();
  const hasActiveJobs = hasPromptPreviewActiveJobs(summary);
  const spawned = Math.max(0, Number(summary.spawned || summary.total || 0));
  const configError = getPromptPreviewConfigError({ requireWorkflow: true });
  const disabled = !selectionReady
    || state.duplicatingImage
    || state.autoCaptioning
    || state.cloning
    || state.uploading
    || state.promptPreview.loading
    || !!configError;

  createPromptPreviewBtn.disabled = disabled;
  createPromptPreviewBtn.classList.toggle("running", hasActiveJobs || state.promptPreview.loading);

  let label = "Create Preview";
  let title = configError || "Queue a new ComfyUI prompt preview for the selected image";
  if (!selectionReady) {
    title = "Select a single image to queue a prompt preview";
  } else if (state.promptPreview.loading && sourceActive) {
    label = "Creating Preview...";
    title = "Queueing a new ComfyUI prompt preview";
  } else if (hasActiveJobs) {
    title = `${spawned} prompt preview job${spawned === 1 ? "" : "s"} queued for ${getFileLabel(sourcePath)}. The latest result will show automatically when it finishes.`;
  }

  setGenerateButtonContent(createPromptPreviewBtn, label);
  createPromptPreviewBtn.title = title;
}

function renderPromptPreviewButton() {
  const visible = canPromptPreviewCurrentImage() && !isMaskEditorVisible();
  const sourcePath = state.previewPath;
  const sourceActive = isPromptPreviewSourceActive(sourcePath);
  const summary = sourceActive ? state.promptPreview.summary : createPromptPreviewSummary();
  const hasActiveJobs = hasPromptPreviewActiveJobs(summary);
  const latestPreviewPath = sourceActive ? getLatestPromptPreviewFile(sourcePath) : "";
  const showingPreview = !!latestPreviewPath && state.promptPreview.displayPath === latestPreviewPath;
  const configError = !latestPreviewPath ? getPromptPreviewConfigError({ requireWorkflow: false }) : "";
  const disabled = !visible
    || state.duplicatingImage
    || state.autoCaptioning
    || state.cloning
    || state.uploading
    || state.promptPreview.loading
    || (!!configError && !latestPreviewPath);

  promptPreviewBtn.classList.toggle("visible", visible);
  promptPreviewBtn.classList.toggle("running", hasActiveJobs);
  promptPreviewBtn.classList.toggle("active", showingPreview);
  promptPreviewBtn.disabled = disabled;
  promptPreviewBtn.setAttribute("aria-busy", String(hasActiveJobs));
  promptPreviewBtn.setAttribute("aria-pressed", showingPreview ? "true" : "false");

  let label = showingPreview ? "Original" : "Preview";
  let title = configError || "Toggle between the original image and the latest generated prompt preview";
  if (showingPreview) {
    title = "Show the original dataset image";
  } else if (state.promptPreview.loading && sourceActive) {
    title = "Queueing a new prompt preview";
  } else if (latestPreviewPath) {
    title = "Show the latest generated prompt preview image";
  } else if (hasActiveJobs) {
    title = "Wait for the latest prompt preview to finish, then click to toggle it";
  } else if (!configError) {
    title = "Load and show the latest generated prompt preview image if one exists";
  }

  promptPreviewBtn.textContent = label;
  promptPreviewBtn.title = title;
}

function renderPromptPreviewControls() {
  renderCreatePromptPreviewButton();
  renderPromptPreviewButton();
}

function applyPromptPreviewSnapshot(sourcePath, payload, options = {}) {
  const { autoDisplayLatest = false } = options;
  const jobs = Array.isArray(payload?.jobs) ? payload.jobs : [];
  const summary = { ...createPromptPreviewSummary(), ...(payload?.summary || {}) };
  const files = Array.from(new Set((Array.isArray(payload?.files) ? payload.files : []).filter(Boolean)));
  const previousCompleted = Number(state.promptPreview.summary.completed || 0);
  const previousFilesKey = state.promptPreview.lastFilesKey || "";
  const nextFilesKey = files.join("\n");
  const filesChanged = nextFilesKey !== previousFilesKey;

  if (filesChanged) {
    const versionBase = Date.now();
    files.forEach((path, index) => {
      state.imageVersions[path] = versionBase + index;
    });
  }

  state.promptPreview.sourcePath = sourcePath;
  state.promptPreview.jobs = jobs;
  state.promptPreview.summary = summary;
  state.promptPreview.files = files;
  state.promptPreview.lastFilesKey = nextFilesKey;

  if (state.promptPreview.displayPath && !files.includes(state.promptPreview.displayPath)) {
    state.promptPreview.displayPath = "";
  }

  const cyclePaths = getPromptPreviewCyclePaths(sourcePath, files);
  const currentDisplayPath = getPromptPreviewCurrentDisplayPath(sourcePath);
  state.promptPreview.cycleIndex = cyclePaths.indexOf(currentDisplayPath);

  renderPromptPreviewControls();
  renderPreviewInfo();

  const completedIncreased = Number(summary.completed || 0) > previousCompleted;
  const shouldAutoDisplayLatest = state.previewPath === sourcePath
    && state.previewMediaType === "image"
    && files.length > 0
    && (autoDisplayLatest || completedIncreased || (previousFilesKey && filesChanged));

  if (shouldAutoDisplayLatest) {
    setPromptPreviewDisplayPath(files[files.length - 1], { preserveView: true }).catch(() => {});
  }
}

async function fetchPromptPreviewStatus(imagePath, options = {}) {
  const { showErrors = true } = options;
  const resp = await fetch(`/api/comfyui/prompt-preview/status?image_path=${encodeURIComponent(imagePath)}`);
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    const message = data.detail || "Failed to load prompt preview status";
    if (showErrors) {
      showErrorToast(`Prompt preview error: ${message}`);
    }
    throw new Error(message);
  }
  return data;
}

async function refreshPromptPreviewStatus(sourcePath = state.previewPath, options = {}) {
  const data = await fetchPromptPreviewStatus(sourcePath, options);
  applyPromptPreviewSnapshot(sourcePath, data, { autoDisplayLatest: !!options.autoDisplayLatest });
  return data;
}

async function queuePromptPreviewFromCurrentCaption(sourcePath = state.previewPath) {
  if (!state.captionCache[sourcePath]) {
    await loadCaptionData(sourcePath);
  }
  const caption = state.captionCache[sourcePath] || { enabled_sentences: [], free_text: "" };
  const resp = await fetch("/api/comfyui/prompt-preview", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      image_path: sourcePath,
      enabled_captions: caption.enabled_sentences || [],
      free_text: caption.free_text || "",
    }),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data.detail || "Failed to queue prompt preview");
  }
  return data;
}

async function runCreatePromptPreviewAction() {
  if (!canPromptPreviewCurrentImage()) {
    showErrorToast("Select a single image first.");
    statusBar.textContent = "Prompt preview requires a single selected image.";
    return;
  }
  const sourcePath = state.previewPath;
  const configError = getPromptPreviewConfigError({ requireWorkflow: true });
  if (configError) {
    showErrorToast(configError);
    statusBar.textContent = configError;
    renderPromptPreviewControls();
    return;
  }

  state.promptPreview.sourcePath = sourcePath;
  state.promptPreview.loading = true;
  statusBar.textContent = `Queueing prompt preview for ${getFileLabel(sourcePath)}...`;
  renderPromptPreviewControls();
  try {
    const queued = await queuePromptPreviewFromCurrentCaption(sourcePath);
    applyPromptPreviewSnapshot(sourcePath, queued, { autoDisplayLatest: false });
    statusBar.textContent = `Queued prompt preview ${Math.max(1, Number(state.promptPreview.summary.spawned || 1))} for ${getFileLabel(sourcePath)}`;
  } catch (err) {
    showErrorToast(`Prompt preview error: ${err.message}`);
    statusBar.textContent = `Prompt preview error: ${err.message}`;
  } finally {
    state.promptPreview.loading = false;
    renderPromptPreviewControls();
  }
}

async function togglePromptPreviewButtonAction() {
  if (!canPromptPreviewCurrentImage()) {
    showErrorToast("Select a single image first.");
    statusBar.textContent = "Prompt preview requires a single selected image.";
    return;
  }

  const sourcePath = state.previewPath;
  let latestPreviewPath = isPromptPreviewSourceActive(sourcePath) ? getLatestPromptPreviewFile(sourcePath) : "";

  if (!latestPreviewPath) {
    const configError = getPromptPreviewConfigError({ requireWorkflow: false });
    if (configError) {
      showErrorToast(configError);
      statusBar.textContent = configError;
      renderPromptPreviewControls();
      return;
    }

    statusBar.textContent = `Checking prompt preview for ${getFileLabel(sourcePath)}...`;
    try {
      const status = await refreshPromptPreviewStatus(sourcePath, { showErrors: false, autoDisplayLatest: false });
      latestPreviewPath = getLatestPromptPreviewFile(sourcePath, Array.isArray(status.files) ? status.files : []);
    } catch (err) {
      showErrorToast(`Prompt preview error: ${err.message}`);
      statusBar.textContent = `Prompt preview error: ${err.message}`;
      renderPromptPreviewControls();
      return;
    }
  }

  if (!latestPreviewPath) {
    showErrorToast("No generated prompt preview file found yet.");
    statusBar.textContent = `No generated prompt preview file found yet for ${getFileLabel(sourcePath)}`;
    renderPromptPreviewControls();
    return;
  }

  try {
    if (state.promptPreview.displayPath === latestPreviewPath) {
      await setPromptPreviewDisplayPath("", { preserveView: true });
      statusBar.textContent = `Showing original image for ${getFileLabel(sourcePath)}`;
    } else {
      await setPromptPreviewDisplayPath(latestPreviewPath, { preserveView: true });
      statusBar.textContent = `Showing latest prompt preview for ${getFileLabel(sourcePath)}`;
    }
  } catch (err) {
    showErrorToast(`Prompt preview error: ${err.message}`);
    statusBar.textContent = `Prompt preview error: ${err.message}`;
  } finally {
    renderPromptPreviewControls();
  }
}

async function revealCurrentPromptPreviewInExplorer() {
  if (!canPromptPreviewCurrentImage()) {
    showErrorToast("Select a single image first.");
    statusBar.textContent = "Prompt preview reveal requires a single selected image.";
    return;
  }
  const sourcePath = state.previewPath;
  const promptPreviewFiles = getPromptPreviewFiles(sourcePath);
  const revealPath = state.promptPreview.displayPath || promptPreviewFiles[promptPreviewFiles.length - 1] || "";
  if (!revealPath || revealPath === sourcePath) {
    showErrorToast("No generated prompt preview file found yet.");
    statusBar.textContent = `No generated prompt preview file found yet for ${getFileLabel(sourcePath)}`;
    return;
  }

  const resp = await fetch(`/api/open-in-explorer?path=${encodeURIComponent(revealPath)}`);
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data.detail || "Failed to reveal prompt preview file");
  }
  statusBar.textContent = `Revealed prompt preview for ${getFileLabel(sourcePath)}`;
}

async function pollPromptPreviewStatus() {
  const sourcePath = state.previewPath;
  if (!sourcePath || state.previewMediaType !== "image") return;
  if (!isPromptPreviewSourceActive(sourcePath)) return;
  if (state.promptPreview.loading) return;
  const summary = state.promptPreview.summary || createPromptPreviewSummary();
  const shouldPoll = Number(summary.running || 0) > 0 || Number(summary.queued || 0) > 0;
  if (!shouldPoll) return;
  try {
    await refreshPromptPreviewStatus(sourcePath, { showErrors: false, autoDisplayLatest: true });
  } catch (err) {
    console.error("Failed to poll prompt preview status:", err);
  }
}

function startPromptPreviewPolling() {
  if (state.ui.promptPreviewPollTimer) return;
  state.ui.promptPreviewPollTimer = window.setInterval(() => {
    pollPromptPreviewStatus().catch(() => {});
  }, 1500);
}

function handleVideoTrainingPresetsInput() {
  const currentSelection = String(settingsVideoProfileInput.value || state.videoTrainingProfileKey || "").trim();
  try {
    const presets = parseVideoTrainingPresetsText(settingsVideoPresetsInput.value);
    populateVideoTrainingProfileSelect(presets, currentSelection);
    setVideoTrainingPresetsStatus("Preset library parsed. Save to apply changes.");
  } catch (err) {
    setVideoTrainingPresetsStatus(err.message || "Invalid video training preset JSON.", { isError: true });
  }
}

function handleVideoTrainingProfileInputChange() {
  if (!state.folder) return;
  setVideoTrainingPresetsStatus("Save to apply the selected profile to the current folder.");
}

function applySettings(settings) {
  const thumbSize = Number(settings.thumb_size || state.thumbSize || 160) || 160;
  state.thumbSize = Math.max(60, Math.min(400, thumbSize));
  thumbSlider.value = String(state.thumbSize);
  document.documentElement.style.setProperty("--thumb-size", state.thumbSize + "px");
  setCropAspectRatios(settings.crop_aspect_ratios || state.cropAspectRatioLabels);
  state.maskLatentBaseWidthPresets = normalizeMaskLatentBaseWidthPresets(settings.mask_latent_base_width_presets || state.maskLatentBaseWidthPresets);
  syncMaskLatentBaseWidthFromPresets();
  state.httpsCertFile = String(settings.https_certfile || "").trim();
  state.httpsKeyFile = String(settings.https_keyfile || "").trim();
  state.httpsPort = Number(settings.https_port || 8900) || 8900;
  state.remoteHttpMode = String(settings.remote_http_mode || "redirect-to-https").trim() || "redirect-to-https";
  state.ffmpegPath = String(settings.ffmpeg_path || "").trim();
  state.ffmpegThreads = Math.max(0, Number(settings.ffmpeg_threads || 0) || 0);
  state.ffmpegHwaccel = String(settings.ffmpeg_hwaccel || "auto").trim() || "auto";
  state.processingReservedCores = Math.max(0, Number(settings.processing_reserved_cores || 0) || 0);
  state.ollamaServer = (settings.ollama_server || "127.0.0.1").trim();
  state.ollamaPort = Number(settings.ollama_port || 11434) || 11434;
  state.ollamaTimeoutSeconds = Number(settings.ollama_timeout_seconds || 20) || 20;
  state.ollamaModel = (settings.ollama_model || "llava").trim() || "llava";
  state.ollamaPromptTemplate = settings.ollama_prompt_template || state.ollamaPromptTemplate;
  state.ollamaGroupPromptTemplate = settings.ollama_group_prompt_template || state.ollamaGroupPromptTemplate;
  state.ollamaEnableFreeText = settings.ollama_enable_free_text ?? state.ollamaEnableFreeText;
  state.ollamaFreeTextPromptTemplate = settings.ollama_free_text_prompt_template || state.ollamaFreeTextPromptTemplate;
  state.comfyuiServer = (settings.comfyui_server || "127.0.0.1").trim() || "127.0.0.1";
  state.comfyuiPort = Number(settings.comfyui_port || 8188) || 8188;
  state.comfyuiWorkflowPath = String(settings.comfyui_workflow_path || "").trim();
  state.comfyuiOutputFolder = String(settings.comfyui_output_folder || "").trim();
  state.comfyuiAutoPreviewEnabled = settings.comfyui_auto_preview ?? state.comfyuiAutoPreviewEnabled;
  if (Object.prototype.hasOwnProperty.call(settings, "video_training_presets")) {
    state.videoTrainingPresets = Array.isArray(settings.video_training_presets) ? settings.video_training_presets : [];
  }
  if (Object.prototype.hasOwnProperty.call(settings, "video_training_profile_key")) {
    state.videoTrainingProfileKey = String(settings.video_training_profile_key || "").trim();
  }
  if (Object.prototype.hasOwnProperty.call(settings, "video_training_profile")) {
    state.videoTrainingProfile = settings.video_training_profile || null;
  }
  if (!findVideoTrainingProfileByKey(state.videoTrainingProfileKey) && Array.isArray(state.videoTrainingPresets) && state.videoTrainingPresets.length > 0) {
    state.videoTrainingProfileKey = String(state.videoTrainingProfile?.key || state.videoTrainingPresets[0].key || "").trim();
  }
  state.videoTrainingProfile = findVideoTrainingProfileByKey(state.videoTrainingProfileKey)
    || state.videoTrainingProfile
    || null;
  if (settings.sections) {
    state.sections = normalizeSectionsData(settings.sections);
  }
  autoFreeTextCheckbox.checked = !!state.ollamaEnableFreeText;
  autoPreviewCheckbox.checked = !!state.comfyuiAutoPreviewEnabled;
  populateVideoTrainingProfileSelect();
  setVideoTrainingPresetsStatus();
  renderVideoTrainingSummary();
  renderGifConvertButton();
  if (state.images.length) {
    renderGrid();
  }
}

function populateOllamaModelSelect(models = [], selectedModel = state.ollamaModel) {
  const uniqueModels = [...new Set((models || []).map(model => String(model || "").trim()).filter(Boolean))];
  state.ollamaAvailableModels = uniqueModels;
  settingsModelInput.innerHTML = "";

  const selectedValue = String(selectedModel || "").trim() || "llava";
  const options = uniqueModels.includes(selectedValue) ? uniqueModels : [selectedValue, ...uniqueModels];

  for (const modelName of options) {
    const option = document.createElement("option");
    option.value = modelName;
    option.textContent = modelName;
    if (modelName === selectedValue) {
      option.selected = true;
    }
    settingsModelInput.appendChild(option);
  }
}

async function refreshOllamaModelOptions() {
  const server = settingsServerInput.value.trim() || "127.0.0.1";
  const port = Number(settingsPortInput.value || "11434") || 11434;
  const currentSelection = settingsModelInput.value || state.ollamaModel || "llava";
  settingsRefreshModelsBtn.disabled = true;
  settingsRefreshModelsBtn.title = "Loading Ollama model list...";
  try {
    const params = new URLSearchParams({ server, port: String(port) });
    const resp = await fetch(`/api/ollama/models?${params.toString()}`);
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to load Ollama models");
    }
    populateOllamaModelSelect(data.models || [], currentSelection);
  } catch (err) {
    populateOllamaModelSelect(state.ollamaAvailableModels || [], currentSelection);
    showErrorToast(`Model list error: ${err.message}`);
  } finally {
    settingsRefreshModelsBtn.disabled = false;
    settingsRefreshModelsBtn.title = "Refresh Ollama model list";
  }
}

function fillSettingsForm() {
  settingsServerInput.value = state.ollamaServer;
  settingsPortInput.value = String(state.ollamaPort);
  settingsTimeoutInput.value = String(state.ollamaTimeoutSeconds);
  populateOllamaModelSelect(state.ollamaAvailableModels || [], state.ollamaModel);
  settingsComfyuiServerInput.value = state.comfyuiServer;
  settingsComfyuiPortInput.value = String(state.comfyuiPort);
  settingsComfyuiWorkflowPathInput.value = state.comfyuiWorkflowPath;
  settingsComfyuiOutputFolderInput.value = state.comfyuiOutputFolder;
  settingsCropAspectRatiosInput.value = state.cropAspectRatioLabels.join(", ");
  settingsMaskLatentBaseWidthPresetsInput.value = getMaskLatentBaseWidthPresets().join(", ");
  settingsHttpsCertInput.value = state.httpsCertFile;
  settingsHttpsKeyInput.value = state.httpsKeyFile;
  settingsHttpsPortInput.value = String(state.httpsPort);
  settingsRemoteHttpModeInput.value = state.remoteHttpMode;
  settingsFfmpegPathInput.value = state.ffmpegPath;
  settingsProcessingReservedCoresInput.value = String(state.processingReservedCores);
  settingsFfmpegThreadsInput.value = String(state.ffmpegThreads);
  settingsFfmpegHwaccelInput.value = state.ffmpegHwaccel;
  settingsVideoPresetsInput.value = JSON.stringify(Array.isArray(state.videoTrainingPresets) ? state.videoTrainingPresets : [], null, 2);
  populateVideoTrainingProfileSelect(state.videoTrainingPresets, state.videoTrainingProfileKey);
  setVideoTrainingPresetsStatus();
  settingsPromptInput.value = state.ollamaPromptTemplate;
  settingsGroupPromptInput.value = state.ollamaGroupPromptTemplate;
  settingsAutoFreeTextEnabled.checked = !!state.ollamaEnableFreeText;
  settingsFreeTextPromptInput.value = state.ollamaFreeTextPromptTemplate;
}

function setActiveSettingsTab(tabName = "auto-captioning") {
  const nextTab = String(tabName || "auto-captioning").trim() || "auto-captioning";
  settingsTabButtons.forEach((button) => {
    const isActive = button.dataset.settingsTab === nextTab;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-selected", isActive ? "true" : "false");
    button.tabIndex = isActive ? 0 : -1;
  });
  settingsPanels.forEach((panel) => {
    const isActive = panel.dataset.settingsPanel === nextTab;
    panel.classList.toggle("active", isActive);
    panel.hidden = !isActive;
  });
}

async function openSettingsModal() {
  fillSettingsForm();
  setActiveSettingsTab("auto-captioning");
  settingsModal.classList.add("open");
  settingsModal.setAttribute("aria-hidden", "false");
  settingsServerInput.focus();
  await refreshOllamaModelOptions();
}

function closeSettingsModal() {
  settingsModal.classList.remove("open");
  settingsModal.setAttribute("aria-hidden", "true");
}

function getFileLabel(path) {
  return String(path || "").split(/[\\/]/).pop() || String(path || "");
}

function getMediaItem(path) {
  return state.images.find((item) => item.path === path) || null;
}

function getMediaType(path = state.previewPath) {
  const item = getMediaItem(path);
  if (item?.media_type) return item.media_type;
  const extension = getImageExtension(path);
  if (IMAGE_FILE_EXTENSIONS.has(extension)) return "image";
  if (VIDEO_FILE_EXTENSIONS.has(extension)) return "video";
  return "other";
}

function isImageMediaPath(path) {
  return getMediaType(path) === "image";
}

function isVideoMediaPath(path) {
  return getMediaType(path) === "video";
}

function isSelectionImagesOnly() {
  if (state.selectedPaths.size === 0) return false;
  return [...state.selectedPaths].every((path) => isImageMediaPath(path));
}

function getActivePreviewElement() {
  return state.previewMediaType === "video" ? previewVideo : previewImg;
}

function isPreviewVisible() {
  const element = getActivePreviewElement();
  return !!element && element.style.display !== "none";
}

function applyPreviewVideoAudioState() {
  previewVideo.muted = !!state.previewVideoMuted;
  previewVideo.volume = clamp(Number(state.previewVideoVolume || 0), 0, 1);
}

function stopPreviewVideo({ clearSource = false } = {}) {
  previewVideo.pause();
  previewVideo.style.display = "none";
  previewVideo.currentTime = 0;
  previewVideo.removeAttribute("poster");
  if (clearSource) {
    previewVideo.removeAttribute("src");
    previewVideo.load();
  }
  syncPreviewVideoPlaybackState();
}

function resetPreviewVideoElement() {
  stopPreviewVideo({ clearSource: true });
  previewVideo.onloadedmetadata = null;
  previewVideo.onloadeddata = null;
  previewVideo.removeAttribute("data-path");
}

function syncPreviewVideoPlaybackState() {
  const isActiveVideo = state.previewMediaType === "video" && !!state.previewPath;
  const meta = isActiveVideo ? getCurrentVideoMeta(state.previewPath) : null;
  const duration = Math.max(0, Number(meta?.duration || previewVideo.duration || 0));
  const currentTime = Math.max(0, Number(previewVideo.currentTime || 0));
  applyPreviewVideoAudioState();
  videoPlayToggleBtn.disabled = !isActiveVideo || previewVideo.readyState < 1;
  videoMuteBtn.disabled = !isActiveVideo || previewVideo.readyState < 1;
  videoVolumeSlider.disabled = !isActiveVideo || previewVideo.readyState < 1;
  videoPlayToggleBtn.textContent = previewVideo.paused ? "Play" : "Pause";
  const isMuted = previewVideo.muted || previewVideo.volume <= 0;
  videoMuteBtn.textContent = isMuted ? "🔇" : "🔊";
  videoMuteBtn.setAttribute("aria-label", isMuted ? "Unmute preview audio" : "Mute preview audio");
  videoMuteBtn.title = isMuted ? "Unmute preview audio" : "Mute preview audio";
  videoMuteBtn.setAttribute("aria-pressed", previewVideo.muted ? "true" : "false");
  videoVolumeSlider.value = String(Math.round(clamp(Number(state.previewVideoVolume || 0), 0, 1) * 100));
  videoPlaybackLabel.textContent = isActiveVideo
    ? `${formatDurationSeconds(currentTime)} / ${formatDurationSeconds(duration)}`
    : "";

  const frames = isActiveVideo ? getVideoTimelineFrames(state.previewPath) : [];
  const frameNodes = videoTimelineStrip.querySelectorAll(".video-timeline-frame");
  renderVideoTimelineOverlay(isActiveVideo ? state.previewPath : null);
  updateVideoTimeRangeLabel(isActiveVideo ? state.previewPath : null);
  if (!frames.length || !frameNodes.length || !(duration > 0)) {
    frameNodes.forEach((node) => node.classList.remove("active"));
    return;
  }

  let activeIndex = 0;
  let bestDistance = Number.POSITIVE_INFINITY;
  frames.forEach((frame, index) => {
    const distance = Math.abs(Number(frame.timeSeconds || 0) - currentTime);
    if (distance < bestDistance) {
      bestDistance = distance;
      activeIndex = index;
    }
  });
  frameNodes.forEach((node, index) => node.classList.toggle("active", index === activeIndex));
}

function getPreviewVideoLoopRange(path = state.previewPath) {
  if (!path || !isVideoMediaPath(path)) return null;
  const duration = getEffectiveVideoDuration(path);
  if (!(duration > 0)) return null;
  const draft = getVideoClipDraft(path);
  const startSeconds = clamp(duration * Number(draft.startFraction || 0), 0, duration);
  const endSeconds = clamp(duration * Number(draft.endFraction || 1), startSeconds, duration);
  return {
    startSeconds,
    endSeconds,
    duration,
  };
}

function snapPreviewVideoIntoLoopRange(options = {}) {
  const { forceStart = false } = options;
  const range = getPreviewVideoLoopRange();
  if (!range || previewVideo.readyState < 1) return null;

  const epsilon = Math.min(0.08, Math.max(0.02, range.duration / 1000));
  const currentTime = Math.max(0, Number(previewVideo.currentTime || 0));
  let nextTime = currentTime;

  if (forceStart) {
    nextTime = range.startSeconds;
  } else if (currentTime < range.startSeconds) {
    nextTime = range.startSeconds;
  } else if (currentTime >= Math.max(range.startSeconds, range.endSeconds - epsilon)) {
    nextTime = range.startSeconds;
  }

  if (Math.abs(nextTime - currentTime) > 0.001) {
    previewVideo.currentTime = nextTime;
  }
  return range;
}

function handlePreviewVideoTimeUpdate() {
  if (state.previewMediaType !== "video" || !state.previewPath) {
    syncPreviewVideoPlaybackState();
    return;
  }
  if (!previewVideo.paused) {
    snapPreviewVideoIntoLoopRange();
  }
  syncPreviewVideoPlaybackState();
  if (
    !previewVideo.paused
    && state.maskEditor.active
    && state.maskEditor.mediaType === "video"
    && !state.maskEditor.loading
    && !state.maskEditor.saving
    && !state.maskEditor.painting
    && !state.maskEditor.switchingKeyframe
  ) {
    const requestedFrameIndex = getCurrentVideoMaskFrameIndex(state.previewPath);
    const nextFrameIndex = getResolvedVideoMaskKeyframeForFrame(state.previewPath, requestedFrameIndex);
    if (nextFrameIndex == null || Number(nextFrameIndex) === Number(state.maskEditor.frameIndex)) {
      return;
    }
    syncActiveVideoMaskEditorToSeekPosition().catch((err) => {
      showErrorToast(`Mask error: ${err.message || err}`);
    });
  }
}

function handlePreviewVideoEnded() {
  const range = snapPreviewVideoIntoLoopRange({ forceStart: true });
  if (range && state.previewMediaType === "video" && state.previewPath) {
    previewVideo.play().catch(() => {
      syncPreviewVideoPlaybackState();
    });
    return;
  }
  syncPreviewVideoPlaybackState();
}

function handlePreviewVideoPause() {
  syncPreviewVideoPlaybackState();
  if (state.maskEditor.active && state.maskEditor.mediaType === "video") {
    syncActiveVideoMaskEditorToSeekPosition().catch((err) => {
      showErrorToast(`Mask error: ${err.message || err}`);
    });
  }
}

function handlePreviewVideoSeeked() {
  syncPreviewVideoPlaybackState();
  if (state.maskEditor.active && state.maskEditor.mediaType === "video") {
    syncActiveVideoMaskEditorToSeekPosition().catch((err) => {
      showErrorToast(`Mask error: ${err.message || err}`);
    });
  }
}

function togglePreviewVideoPlayback() {
  if (state.previewMediaType !== "video" || !state.previewPath || previewVideo.readyState < 1) {
    return;
  }
  if (previewVideo.paused) {
    snapPreviewVideoIntoLoopRange();
    previewVideo.play().catch(() => {
      syncPreviewVideoPlaybackState();
    });
  } else {
    previewVideo.pause();
  }
  syncPreviewVideoPlaybackState();
}

function togglePreviewVideoMute() {
  if (state.previewVideoMuted || state.previewVideoVolume <= 0) {
    const restoredVolume = clamp(Number(state.previewVideoLastVolume || 1), 0.05, 1);
    state.previewVideoVolume = restoredVolume;
    state.previewVideoMuted = false;
  } else {
    state.previewVideoLastVolume = clamp(Number(state.previewVideoVolume || 1), 0.05, 1);
    state.previewVideoMuted = true;
  }
  applyPreviewVideoAudioState();
  syncPreviewVideoPlaybackState();
}

function setPreviewVideoVolumeFromSlider(value) {
  const normalized = clamp(Number(value || 0) / 100, 0, 1);
  state.previewVideoVolume = normalized;
  if (normalized > 0) {
    state.previewVideoLastVolume = normalized;
  }
  state.previewVideoMuted = normalized <= 0;
  applyPreviewVideoAudioState();
  syncPreviewVideoPlaybackState();
}

function seekPreviewVideoTo(timeSeconds) {
  if (state.previewMediaType !== "video" || !state.previewPath || previewVideo.readyState < 1) {
    return;
  }
  const duration = getEffectiveVideoDuration(state.previewPath);
  previewVideo.currentTime = clamp(Number(timeSeconds || 0), 0, Math.max(0, duration));
  syncPreviewVideoPlaybackState();
}

function formatDurationSeconds(totalSeconds) {
  const seconds = Math.max(0, Number(totalSeconds || 0));
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const wholeSeconds = Math.floor(seconds % 60);
  const millis = Math.round((seconds - Math.floor(seconds)) * 1000);
  if (hours > 0) {
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(wholeSeconds).padStart(2, "0")}.${String(millis).padStart(3, "0")}`;
  }
  return `${String(minutes).padStart(2, "0")}:${String(wholeSeconds).padStart(2, "0")}.${String(millis).padStart(3, "0")}`;
}

function getCurrentVideoMeta(path = state.previewPath) {
  return state.videoMeta[path] || null;
}

function getEffectiveVideoDuration(path = state.previewPath) {
  const metaDuration = Number(getCurrentVideoMeta(path)?.duration || 0);
  const elementDuration = Number(previewVideo.duration || 0);
  return Math.max(0, metaDuration || elementDuration || 0);
}

function getVideoTimelineFrames(path = state.previewPath) {
  const cache = state.videoTimelineCache[path];
  if (Array.isArray(cache)) return cache;
  if (cache && Array.isArray(cache.frames)) return cache.frames;
  return [];
}

function getVideoTimelineUi(path = state.previewPath) {
  if (!path) {
    return { zoom: 1, offsetFraction: 0 };
  }
  if (!state.videoTimelineUi[path]) {
    state.videoTimelineUi[path] = { zoom: 1, offsetFraction: 0 };
  }
  const ui = state.videoTimelineUi[path];
  ui.zoom = clamp(Number(ui.zoom || 1), 1, 8);
  ui.offsetFraction = clamp(Number(ui.offsetFraction || 0), 0, 1);
  return ui;
}

function getVideoTimelineMetrics(path = state.previewPath) {
  const viewportWidth = Math.max(videoTimelineViewport?.clientWidth || 0, 280);
  const ui = getVideoTimelineUi(path);
  const trackWidth = Math.max(viewportWidth, Math.round(viewportWidth * ui.zoom));
  const maxOffset = Math.max(0, trackWidth - viewportWidth);
  const offsetPx = maxOffset * ui.offsetFraction;
  return { ui, viewportWidth, trackWidth, maxOffset, offsetPx, zoom: ui.zoom };
}

function getVideoTimelineFrameCount(path = state.previewPath) {
  const metrics = getVideoTimelineMetrics(path);
  return Math.max(10, Math.min(72, Math.ceil(metrics.trackWidth / 72)));
}

function setVideoTimelineOffsetPx(path, nextOffsetPx) {
  const metrics = getVideoTimelineMetrics(path);
  metrics.ui.offsetFraction = metrics.maxOffset > 0
    ? clamp(nextOffsetPx / metrics.maxOffset, 0, 1)
    : 0;
}

function setVideoTimelineZoom(path, nextZoom, anchorViewportRatio = 0.5) {
  const previousMetrics = getVideoTimelineMetrics(path);
  const anchorX = clamp(Number(anchorViewportRatio || 0.5), 0, 1) * previousMetrics.viewportWidth;
  const anchorContentRatio = previousMetrics.trackWidth > 0
    ? clamp((previousMetrics.offsetPx + anchorX) / previousMetrics.trackWidth, 0, 1)
    : 0;

  const ui = getVideoTimelineUi(path);
  ui.zoom = clamp(Number(nextZoom || 1), 1, 8);

  const nextMetrics = getVideoTimelineMetrics(path);
  const targetOffsetPx = clamp(anchorContentRatio * nextMetrics.trackWidth - anchorX, 0, nextMetrics.maxOffset);
  ui.offsetFraction = nextMetrics.maxOffset > 0 ? targetOffsetPx / nextMetrics.maxOffset : 0;
}

function getVideoClipDraft(path = state.previewPath) {
  const existing = state.videoClipDrafts[path];
  if (existing) return existing;
  return { startFraction: 0, endFraction: 1 };
}

function setVideoClipDraft(path, patch = {}) {
  const current = getVideoClipDraft(path);
  const next = {
    startFraction: clamp(Number(patch.startFraction ?? current.startFraction ?? 0), 0, 1),
    endFraction: clamp(Number(patch.endFraction ?? current.endFraction ?? 1), 0, 1),
  };
  if (next.endFraction < next.startFraction) {
    const midpoint = next.startFraction;
    next.startFraction = Math.min(midpoint, next.endFraction);
    next.endFraction = Math.max(midpoint, next.endFraction);
  }
  state.videoClipDrafts[path] = next;
  return next;
}

function buildVideoFrameUrl(path, timeSeconds, width = 160, height = 90) {
  return buildImageApiUrl("video/frame", path, {
    time_seconds: Number(timeSeconds || 0).toFixed(3),
    width,
    height,
  });
}

function getVideoTimelineFetchMap(path) {
  let fetchMap = state.ui.videoTimelineFetches.get(path);
  if (!fetchMap) {
    fetchMap = new Map();
    state.ui.videoTimelineFetches.set(path, fetchMap);
  }
  return fetchMap;
}

function abortInvisibleVideoTimelineFetches(path, keepIndexes = new Set()) {
  const fetchMap = state.ui.videoTimelineFetches.get(path);
  if (!fetchMap) return;
  for (const [index, entry] of fetchMap.entries()) {
    if (keepIndexes.has(index)) continue;
    entry.controller?.abort();
    fetchMap.delete(index);
    const cache = state.videoTimelineCache[path];
    const frame = cache?.frames?.[index];
    if (frame && frame.status === "loading") {
      frame.status = "idle";
    }
  }
  if (fetchMap.size === 0) {
    state.ui.videoTimelineFetches.delete(path);
  }
}

async function ensureVisibleVideoTimelineFrames(path, visibleIndexes = []) {
  const cache = state.videoTimelineCache[path];
  if (!cache?.frames?.length) return;
  const keepIndexes = new Set(visibleIndexes);
  abortInvisibleVideoTimelineFetches(path, keepIndexes);
  const fetchMap = getVideoTimelineFetchMap(path);
  const requestVersion = Number(cache.requestVersion || 0);

  visibleIndexes.forEach((index) => {
    const frame = cache.frames[index];
    if (!frame || frame.status === "loaded" || frame.status === "loading") return;
    const controller = new AbortController();
    frame.status = "loading";
    fetchMap.set(index, { controller, requestVersion });

    fetch(frame.requestUrl, { signal: controller.signal })
      .then((resp) => {
        if (!resp.ok) {
          throw new Error(`Failed to load timeline frame (${resp.status})`);
        }
        return resp.blob();
      })
      .then((blob) => {
        const activeCache = state.videoTimelineCache[path];
        const activeEntry = state.ui.videoTimelineFetches.get(path)?.get(index);
        if (!activeCache || activeCache.requestVersion !== requestVersion || !activeEntry || activeEntry.controller !== controller) {
          return;
        }
        if (frame.objectUrl) {
          URL.revokeObjectURL(frame.objectUrl);
        }
        frame.objectUrl = URL.createObjectURL(blob);
        frame.status = "loaded";
        state.ui.videoTimelineFetches.get(path)?.delete(index);
        if (state.previewPath === path) {
          renderVideoTimelineStrip(path);
          syncPreviewVideoPlaybackState();
        }
      })
      .catch((err) => {
        if (err?.name === "AbortError") {
          return;
        }
        frame.status = "idle";
      })
      .finally(() => {
        const activeMap = state.ui.videoTimelineFetches.get(path);
        if (activeMap?.get(index)?.controller === controller && frame.status !== "loaded") {
          activeMap.delete(index);
        }
        if (activeMap && activeMap.size === 0) {
          state.ui.videoTimelineFetches.delete(path);
        }
      });
  });
}

async function ensureVideoMetaLoaded(path) {
  if (!path || !isVideoMediaPath(path)) return null;
  if (state.videoMeta[path]) return state.videoMeta[path];
  const resp = await fetch(`/api/video/meta?path=${encodeURIComponent(path)}`);
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data.detail || "Failed to load video metadata");
  }
  state.videoMeta[path] = data;
  return data;
}

function renderVideoTimelineStrip(path = state.previewPath) {
  if (!path || !isVideoMediaPath(path)) {
    videoTimelineStrip.replaceChildren();
    videoTimelineStrip.style.width = "0px";
    videoTimelineStrip.style.transform = "translateX(0px)";
    videoTimelineOverlay.style.width = "0px";
    videoTimelineOverlay.style.transform = "translateX(0px)";
    return;
  }
  const frames = getVideoTimelineFrames(path);
  if (!frames.length) {
    videoTimelineStrip.replaceChildren();
    videoTimelineStrip.style.width = "0px";
    videoTimelineStrip.style.transform = "translateX(0px)";
    videoTimelineOverlay.style.width = "0px";
    videoTimelineOverlay.style.transform = "translateX(0px)";
    return;
  }
  const metrics = getVideoTimelineMetrics(path);
  const frameWidth = Math.max(48, Math.ceil(metrics.trackWidth / Math.max(1, frames.length)));
  const startIndex = Math.max(0, Math.floor(metrics.offsetPx / frameWidth) - 2);
  const endIndex = Math.min(frames.length - 1, Math.ceil((metrics.offsetPx + metrics.viewportWidth) / frameWidth) + 2);
  const visibleIndexes = [];
  const nodes = [];
  for (let index = startIndex; index <= endIndex; index += 1) {
    const frame = frames[index];
    if (!frame) continue;
    visibleIndexes.push(index);
    const frameNode = document.createElement("div");
    frameNode.className = "video-timeline-frame";
    frameNode.setAttribute("role", "img");
    frameNode.setAttribute("aria-label", `Timeline frame at ${formatDurationSeconds(frame.timeSeconds)}`);
    frameNode.dataset.timeSeconds = String(frame.timeSeconds);
    frameNode.style.width = `${frameWidth}px`;
    frameNode.style.left = `${index * frameWidth}px`;

    const img = document.createElement("img");
    img.className = "video-timeline-frame-image";
    img.alt = "";
    img.loading = "lazy";
    if (frame.objectUrl) {
      img.src = frame.objectUrl;
      frameNode.classList.add("loaded");
    } else {
      frameNode.classList.add("loading");
    }

    const loader = document.createElement("div");
    loader.className = "video-timeline-frame-loader";
    loader.setAttribute("aria-hidden", "true");

    frameNode.append(img, loader);
    nodes.push(frameNode);
  }
  videoTimelineStrip.replaceChildren(...nodes);
  videoTimelineStrip.style.width = `${metrics.trackWidth}px`;
  videoTimelineStrip.style.transform = `translateX(${-metrics.offsetPx}px)`;
  videoTimelineOverlay.style.width = `${metrics.trackWidth}px`;
  videoTimelineOverlay.style.transform = `translateX(${-metrics.offsetPx}px)`;
  ensureVisibleVideoTimelineFrames(path, visibleIndexes).catch(() => {});
}

async function ensureVideoTimelineLoaded(path) {
  if (!path || !isVideoMediaPath(path)) return;
  const meta = await ensureVideoMetaLoaded(path);
  const frameCount = getVideoTimelineFrameCount(path);
  const cached = state.videoTimelineCache[path];
  if (cached?.frameCount === frameCount) {
    renderVideoTimelineStrip(path);
    return;
  }
  const duration = Math.max(0, Number(meta?.duration || 0));
  const frames = [];
  for (let index = 0; index < frameCount; index += 1) {
    const ratio = frameCount <= 1 ? 0 : index / (frameCount - 1);
    const timeSeconds = duration <= 0 ? 0 : ratio * Math.max(duration - 0.05, 0);
    frames.push({
      timeSeconds,
      requestUrl: buildVideoFrameUrl(path, timeSeconds, 160, 90),
      objectUrl: null,
      status: "idle",
    });
  }
  if (cached?.frames) {
    cached.frames.forEach((frame) => {
      if (frame?.objectUrl) {
        URL.revokeObjectURL(frame.objectUrl);
      }
    });
  }
  abortInvisibleVideoTimelineFetches(path, new Set());
  state.videoTimelineCache[path] = {
    frameCount,
    frames,
    requestVersion: (Number(cached?.requestVersion || 0) + 1),
  };
  renderVideoTimelineStrip(path);
}

function renderVideoTimelineOverlay(path = state.previewPath) {
  if (!path || !isVideoMediaPath(path)) {
    videoTimelineOverlay.querySelectorAll(".video-timeline-mask-key").forEach((node) => node.remove());
    videoTimelineSelection.style.width = "0px";
    videoTimelineStartHandle.style.left = "0px";
    videoTimelineEndHandle.style.left = "0px";
    videoTimelinePlayhead.style.left = "0px";
    return;
  }
  const duration = getEffectiveVideoDuration(path);
  const draft = getVideoClipDraft(path);
  const metrics = getVideoTimelineMetrics(path);
  const startX = draft.startFraction * metrics.trackWidth;
  const endX = draft.endFraction * metrics.trackWidth;
  const playheadFraction = duration > 0
    ? clamp(Number(previewVideo.currentTime || 0) / duration, 0, 1)
    : 0;
  const playheadX = playheadFraction * metrics.trackWidth;
  const requestedFrameIndex = getCurrentVideoMaskFrameIndex(path);
  const currentTargetFrameIndex = getResolvedVideoMaskKeyframeForFrame(path, requestedFrameIndex);
  videoTimelineSelection.style.left = `${startX}px`;
  videoTimelineSelection.style.width = `${Math.max(0, endX - startX)}px`;
  videoTimelineStartHandle.style.left = `${startX}px`;
  videoTimelineEndHandle.style.left = `${endX}px`;
  videoTimelinePlayhead.style.left = `${playheadX}px`;

  const existingMarkers = [...videoTimelineOverlay.querySelectorAll(".video-timeline-mask-key")];
  const keyframes = getVideoMaskKeyframes(path);
  for (let index = 0; index < keyframes.length; index += 1) {
    const frameIndex = keyframes[index];
    const marker = existingMarkers[index] || document.createElement("div");
    marker.className = "video-timeline-mask-key";
    const markerFraction = clamp(
      (Math.max(0, frameIndex) / Math.max(1, getEffectiveVideoMaskFps(path))) / Math.max(duration, 0.001),
      0,
      1,
    );
    marker.style.left = `${markerFraction * metrics.trackWidth}px`;
    marker.title = `Mask key ${formatVideoMaskFrameHint(frameIndex, path)}`;
    marker.classList.toggle("active", state.maskEditor.mediaType === "video" && state.maskEditor.path === path && Number(state.maskEditor.frameIndex) === frameIndex);
    marker.classList.toggle("current-target", currentTargetFrameIndex != null && Number(currentTargetFrameIndex) === frameIndex);
    marker.setAttribute("aria-hidden", "true");
    if (!marker.parentElement) {
      videoTimelineOverlay.appendChild(marker);
    }
  }
  existingMarkers.slice(keyframes.length).forEach((node) => node.remove());
}

function updateVideoTimeRangeLabel(path = state.previewPath) {
  if (!path || !isVideoMediaPath(path)) {
    videoTimeRangeLabel.textContent = "";
    videoTimelineZoomLabel.textContent = "";
    return;
  }
  const duration = getEffectiveVideoDuration(path);
  const draft = getVideoClipDraft(path);
  const startSeconds = duration * draft.startFraction;
  const endSeconds = duration * draft.endFraction;
  videoTimeRangeLabel.textContent = `${formatDurationSeconds(startSeconds)} - ${formatDurationSeconds(endSeconds)}`;
  videoTimelineZoomLabel.textContent = `Zoom ${getVideoTimelineMetrics(path).zoom.toFixed(1)}x`;
}

function getVideoTimelineFractionFromClientX(clientX, path = state.previewPath) {
  if (!path || !isVideoMediaPath(path)) return 0;
  const rect = videoTimelineViewport.getBoundingClientRect();
  const metrics = getVideoTimelineMetrics(path);
  const localX = clamp(clientX - rect.left, 0, rect.width || metrics.viewportWidth);
  return clamp((metrics.offsetPx + localX) / Math.max(1, metrics.trackWidth), 0, 1);
}

function syncVideoTimelineState(path = state.previewPath) {
  updateVideoTimeRangeLabel(path);
  renderVideoTimelineOverlay(path);
  const hasPreviewSize = !!(imgNatW && imgNatH);
  const hasDuration = getEffectiveVideoDuration(path) > 0;
  const draft = getVideoClipDraft(path);
  const clipValid = draft.endFraction > draft.startFraction;
  videoClipBtn.disabled = false;
  videoTimelineStartHandle.disabled = !hasDuration;
  videoTimelineEndHandle.disabled = !hasDuration;
  videoTimelinePlayhead.disabled = !hasDuration;
  videoTimelineViewport.classList.toggle("disabled", !hasDuration);
  renderVideoTrainingSummary();
}

function setVideoClipFractions(path, startFraction, endFraction, seekFraction = null) {
  if (!path || !isVideoMediaPath(path)) return;
  const next = setVideoClipDraft(path, { startFraction, endFraction });
  updateVideoTimeRangeLabel(path);
  renderVideoTimelineOverlay(path);
  syncVideoTimelineState(path);
  if (seekFraction != null) {
    const duration = getEffectiveVideoDuration(path);
    seekPreviewVideoTo(duration * clamp(seekFraction, 0, 1));
  } else if (state.previewPath === path) {
    snapPreviewVideoIntoLoopRange();
    syncPreviewVideoPlaybackState();
  }
}

function beginVideoTimelineInteraction(mode, event) {
  if (!state.previewPath || !isVideoMediaPath(state.previewPath)) return;
  const path = state.previewPath;
  const metrics = getVideoTimelineMetrics(path);
  const draft = getVideoClipDraft(path);
  videoTimelineInteraction = {
    mode,
    path,
    startClientX: event.clientX,
    startOffsetPx: metrics.offsetPx,
    startFraction: getVideoTimelineFractionFromClientX(event.clientX, path),
    startStartFraction: draft.startFraction,
    startEndFraction: draft.endFraction,
    moved: false,
  };
  if (mode === "pan") {
    videoTimelineViewport.classList.add("dragging");
  }
  event.preventDefault();
}

function updateVideoTimelineInteraction(clientX) {
  const interaction = videoTimelineInteraction;
  if (!interaction) return;
  const path = interaction.path;
  const dx = clientX - interaction.startClientX;
  if (!interaction.moved && Math.abs(dx) > 2) {
    interaction.moved = true;
  }

  if (interaction.mode === "pan") {
    setVideoTimelineOffsetPx(path, interaction.startOffsetPx - dx);
    renderVideoTimelineStrip(path);
    syncVideoTimelineState(path);
    syncPreviewVideoPlaybackState();
    return;
  }

  const nextFraction = getVideoTimelineFractionFromClientX(clientX, path);
  if (interaction.mode === "start") {
    setVideoClipFractions(path, nextFraction, interaction.startEndFraction, nextFraction);
    return;
  }
  if (interaction.mode === "end") {
    setVideoClipFractions(path, interaction.startStartFraction, nextFraction, nextFraction);
    return;
  }
  if (interaction.mode === "playhead") {
    const duration = getEffectiveVideoDuration(path);
    seekPreviewVideoTo(duration * nextFraction);
  }
}

function finishVideoTimelineInteraction(clientX = null) {
  const interaction = videoTimelineInteraction;
  if (!interaction) return;
  if (interaction.mode === "pan" && !interaction.moved && clientX != null) {
    const duration = getEffectiveVideoDuration(interaction.path);
    seekPreviewVideoTo(duration * getVideoTimelineFractionFromClientX(clientX, interaction.path));
  }
  videoTimelineViewport.classList.remove("dragging");
  videoTimelineInteraction = null;
}

function renderVideoEditPanel() {
  const videoVisible = state.previewMediaType === "video" && !!state.previewPath && state.selectedPaths.size === 1;
  const visible = videoVisible;
  videoEditPanel.classList.toggle("visible", visible);
  videoEditPanel.classList.toggle("crop-mode", false);
  renderVideoTrainingSummary();
  if (!visible) {
    videoDownloadBtn.disabled = true;
    syncPreviewVideoPlaybackState();
    return;
  }
  videoClipBtn.disabled = false;
  videoDownloadBtn.disabled = false;
  renderVideoTimelineStrip(state.previewPath);
  syncVideoTimelineState(state.previewPath);
  syncPreviewVideoPlaybackState();
}

function downloadCurrentVideo() {
  if (!state.previewPath || !isVideoMediaPath(state.previewPath)) return;
  const path = state.previewPath;
  const link = document.createElement("a");
  link.href = buildImageApiUrl("media", path);
  link.download = getFileLabel(path);
  document.body.appendChild(link);
  link.click();
  link.remove();
  statusBar.textContent = `Downloading ${getFileLabel(path)}...`;
}

async function queueCurrentVideoClip() {
  statusBar.textContent = "Preparing clip...";
  if (!state.previewPath || !isVideoMediaPath(state.previewPath)) return;
  let duration = getEffectiveVideoDuration(state.previewPath);
  if (!(duration > 0)) {
    const meta = getCurrentVideoMeta(state.previewPath) || await ensureVideoMetaLoaded(state.previewPath);
    duration = Math.max(0, Number(meta?.duration || 0));
  }
  if (!(duration > 0)) {
    showErrorToast("Video duration is not ready yet.");
    return;
  }
  const draft = getVideoClipDraft(state.previewPath);
  const startSeconds = duration * draft.startFraction;
  const endSeconds = duration * draft.endFraction;
  const crop = state.cropDraft ? buildCropPayload(state.cropDraft) : null;
  if (!(endSeconds > startSeconds)) {
    showErrorToast("Select a valid clip range first.");
    return;
  }
  statusBar.textContent = crop ? "Queueing crop + clip..." : "Queueing clip...";
  try {
    await enqueueVideoClipJob(state.previewPath, startSeconds, endSeconds, crop);
    await pollVideoJobStatus();
    if (crop) {
      state.cropDirty = false;
    }
    statusBar.textContent = `Queued ${crop ? "crop + clip" : "clip"} for ${getFileLabel(state.previewPath)}`;
  } catch (err) {
    statusBar.textContent = `Clip error: ${err.message}`;
    showErrorToast(`Clip error: ${err.message}`);
  }
}


async function enqueueGifConvertJob(path) {
  const resp = await fetch("/api/media/jobs/convert-gif", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ media_path: path }),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data.detail || "Failed to queue GIF conversion job");
  }
  return data.job || null;
}

async function queueCurrentGifConversion() {
  if (!state.previewPath || !isGifMediaPath(state.previewPath)) return;
  statusBar.textContent = "Queueing GIF conversion...";
  try {
    await enqueueGifConvertJob(state.previewPath);
    await pollVideoJobStatus();
    statusBar.textContent = `Queued GIF conversion for ${getFileLabel(state.previewPath)}`;
  } catch (err) {
    statusBar.textContent = `GIF conversion error: ${err.message}`;
    showErrorToast(`GIF conversion error: ${err.message}`);
  }
}

function getVideoJobPresentTenseLabel(type) {
  if (type === "clip") return "Clipping";
  if (type === "crop") return "Cropping";
  if (type === "gif_to_mp4") return "Converting";
  return "Processing";
}

function getVideoJobPastLabel(type) {
  if (type === "clip") return "Clip";
  if (type === "crop") return "Crop";
  if (type === "gif_to_mp4") return "GIF conversion";
  return "Video";
}
async function enqueueVideoCropJob(path, crop) {
  const resp = await fetch("/api/video/jobs/crop", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ video_path: path, crop }),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data.detail || "Failed to queue video crop job");
  }
  return data.job || null;
}

async function enqueueVideoClipJob(path, startSeconds, endSeconds, crop = null) {
  const resp = await fetch("/api/video/jobs/clip", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      video_path: path,
      start_seconds: startSeconds,
      end_seconds: endSeconds,
      crop,
    }),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data.detail || "Failed to queue video clip job");
  }
  return data.job || null;
}

function syncVideoJobBatchState(unseenFinished = []) {
  const activeJob = state.videoJobs.activeJob;
  const queuedJobs = Array.isArray(state.videoJobs.queuedJobs) ? state.videoJobs.queuedJobs : [];
  const batch = state.videoJobs.batch;
  const currentIds = new Set([
    ...(activeJob?.id ? [activeJob.id] : []),
    ...queuedJobs.map((job) => job?.id).filter(Boolean),
  ]);
  const visible = currentIds.size > 0;

  if (!visible) {
    batch.active = false;
    batch.total = 0;
    batch.completed = 0;
    batch.jobIds.clear();
    return;
  }

  if (!batch.active) {
    batch.active = true;
    batch.total = 0;
    batch.completed = 0;
    batch.jobIds.clear();
  }

  currentIds.forEach((id) => {
    if (!batch.jobIds.has(id)) {
      batch.jobIds.add(id);
      batch.total += 1;
    }
  });

  unseenFinished.forEach((job) => {
    if (!job?.id || !batch.jobIds.has(job.id)) return;
    batch.completed = Math.min(batch.total, batch.completed + 1);
  });
}

function renderVideoJobStatus() {
  const activeJob = state.videoJobs.activeJob;
  const queuedCount = Array.isArray(state.videoJobs.queuedJobs) ? state.videoJobs.queuedJobs.length : 0;
  const running = !!activeJob;
  const visible = running || queuedCount > 0;
  videoJobStatus.hidden = !visible;
  videoJobStatus.classList.toggle("visible", visible);
  renderToolbarStatusVisibility();
  if (!visible) {
    videoJobText.textContent = "";
    videoJobProgressFill.style.width = "0%";
    videoJobProgressFill.classList.remove("active");
    return;
  }

  const total = Math.max(1, Number(state.videoJobs.batch.total || 0));
  const completed = Math.max(0, Number(state.videoJobs.batch.completed || 0));
  const activeFraction = running ? Math.max(0, Math.min(1, Number(activeJob.progress || 0))) : 0;
  const percent = Math.max(0, Math.min(100, ((completed + activeFraction) / total) * 100));
  const parts = [];
  if (activeJob) {
    parts.push(`${getVideoJobPresentTenseLabel(activeJob.type)} ${getFileLabel(activeJob.video_path)}`);
    if (activeJob.message) parts.push(activeJob.message);
  }
  if (queuedCount > 0) {
    parts.push(`${queuedCount} queued`);
  }
  parts.push(`${completed}/${total} done`);
  videoJobText.textContent = parts.join(" • ");
  videoJobProgressFill.style.width = `${percent}%`;
  videoJobProgressFill.classList.toggle("active", running);
  renderGifConvertButton();
}

async function handleCompletedVideoJobs(jobs) {
  const relevantJobs = (jobs || []).filter((job) => String(job.folder || "") === String(state.folder || ""));
  if (!relevantJobs.length) return;
  const generatedOutputPaths = relevantJobs
    .filter((job) => job.status === "completed" && (job.type === "clip" || job.type === "crop" || job.type === "gif_to_mp4") && job.output_path)
    .map((job) => job.output_path);
  if (generatedOutputPaths.length) {
    resetPreviewVideoElement();
    previewImg.style.display = "none";
    previewInfo.style.display = "none";
    previewPlaceholder.style.display = "flex";
  }
  const preserveScrollTop = fileGridContainer.scrollTop;
  await loadFolder({ preserveScrollTop });
  if (generatedOutputPaths.length) {
    await selectUploadedImages([generatedOutputPaths[generatedOutputPaths.length - 1]]);
  }
}

async function pollVideoJobStatus() {
  try {
    const resp = await fetch("/api/video/jobs/status");
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to load video job status");
    }
    state.videoJobs.activeJob = data.active_job || null;
    state.videoJobs.queuedJobs = Array.isArray(data.queued_jobs) ? data.queued_jobs : [];
    state.videoJobs.recentJobs = Array.isArray(data.recent_jobs) ? data.recent_jobs : [];
    state.videoJobs.summary = data.summary || state.videoJobs.summary;

    const unseenFinished = state.videoJobs.recentJobs.filter((job) => {
      if (!job?.id || state.videoJobs.seenFinishedIds.has(job.id)) return false;
      return job.status === "completed" || job.status === "error";
    });
    syncVideoJobBatchState(unseenFinished);
    renderVideoJobStatus();
    unseenFinished.forEach((job) => state.videoJobs.seenFinishedIds.add(job.id));
    if (unseenFinished.length > 0) {
      const completed = unseenFinished.filter((job) => job.status === "completed");
      const failed = unseenFinished.filter((job) => job.status === "error");
      if (failed.length > 0) {
        showErrorToast(failed[0].error || `${getVideoJobPastLabel(failed[0].type)} job failed`);
      }
      if (completed.length > 0) {
        await handleCompletedVideoJobs(completed);
      }
    }
  } catch (err) {
    console.error("Failed to poll video jobs:", err);
  }
}

function startVideoJobPolling() {
  if (state.ui.videoJobPollTimer) return;
  pollVideoJobStatus().catch(() => {});
  state.ui.videoJobPollTimer = window.setInterval(() => {
    pollVideoJobStatus().catch(() => {});
  }, 1200);
}

function getCaptionPathForImage(imagePath) {
  const value = String(imagePath || "");
  const lastDot = value.lastIndexOf(".");
  if (lastDot <= 0) return `${value}.txt`;
  return `${value.slice(0, lastDot)}.txt`;
}

function isEditableElement(element) {
  return !!(
    element?.tagName === "INPUT" ||
    element?.tagName === "TEXTAREA" ||
    element?.isContentEditable
  );
}

async function releasePreviewMediaForDeletion(paths) {
  const targets = new Set((paths || []).filter(Boolean));
  if (!targets.size) return;
  if (!state.previewPath || !targets.has(state.previewPath)) return;

  if (state.previewMediaType === "video") {
    stopPreviewVideo({ clearSource: true });
    previewInfo.style.display = "none";
    previewPlaceholder.style.display = "flex";
    state.previewPath = null;
    state.previewMediaType = null;
    imgNatW = 0;
    imgNatH = 0;
    clearCropDraft();
    renderPreviewCaptionOverlay();
    renderVideoEditPanel();
    await new Promise((resolve) => window.setTimeout(resolve, 50));
    return;
  }

  hidePreview();
}

function getNextSelectionPathAfterDelete(paths) {
  const deletedPaths = new Set((paths || []).filter(Boolean));
  if (!deletedPaths.size) return null;

  const visibleEntries = getVisibleImageEntries();
  const fallbackPath = visibleEntries.find(({ img }) => !deletedPaths.has(img.path))?.img.path || null;
  if (!visibleEntries.length) {
    return fallbackPath;
  }

  const anchorPath = state.lastClickedPath || state.previewPath || [...state.selectedPaths][0] || null;
  let anchorIndex = anchorPath
    ? visibleEntries.findIndex(({ img }) => img.path === anchorPath)
    : -1;

  if (anchorIndex < 0) {
    anchorIndex = visibleEntries.findIndex(({ img }) => deletedPaths.has(img.path));
  }

  if (anchorIndex < 0) {
    return fallbackPath;
  }

  for (let index = anchorIndex + 1; index < visibleEntries.length; index++) {
    const candidatePath = visibleEntries[index].img.path;
    if (!deletedPaths.has(candidatePath)) {
      return candidatePath;
    }
  }

  for (let index = anchorIndex - 1; index >= 0; index--) {
    const candidatePath = visibleEntries[index].img.path;
    if (!deletedPaths.has(candidatePath)) {
      return candidatePath;
    }
  }

  return fallbackPath;
}

async function deleteSelectedImages() {
  if (state.selectedPaths.size === 0) return;

  const selectedPaths = [...state.selectedPaths];
  const count = selectedPaths.length;
  const nextSelectionPath = getNextSelectionPathAfterDelete(selectedPaths);
  const confirmMessage = count === 1
    ? `Delete "${getFileLabel(selectedPaths[0])}"? This also deletes its .txt caption and .meta.json metadata files.`
    : `Delete ${count} selected media files? This also deletes their .txt caption and .meta.json metadata files.`;
  if (!confirm(confirmMessage)) return;

  const preserveScrollTop = fileGridContainer.scrollTop;
  statusBar.textContent = count === 1
    ? `Deleting ${getFileLabel(selectedPaths[0])}...`
    : `Deleting ${count} media files...`;

  try {
    await releasePreviewMediaForDeletion(selectedPaths);
    const resp = await fetch("/api/images/delete", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ image_paths: selectedPaths }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to delete images");
    }

    for (const path of data.deleted_paths || []) {
      invalidateImageCaches(path);
      delete state.captionCache[path];
      delete state.metadataCache[path];
      delete state.imageCrops[path];
      delete state.imageVersions[path];
    }

    await loadFolder({ preserveScrollTop });
    const deletedPaths = new Set((data.deleted_paths || []).filter(Boolean));
    const survivingSelectedPaths = selectedPaths.filter(path => !deletedPaths.has(path));
    if (survivingSelectedPaths.length > 0) {
      await selectUploadedImages(survivingSelectedPaths);
    } else if (deletedPaths.size > 0 && nextSelectionPath) {
      await selectUploadedImages([nextSelectionPath]);
    }

    if (Array.isArray(data.errors) && data.errors.length > 0) {
      const summary = data.errors.length === 1
        ? `${data.deleted_count || 0} deleted, 1 failed: ${data.errors[0].error || "Unknown error"}`
        : `${data.deleted_count || 0} deleted, ${data.errors.length} failed`;
      statusBar.textContent = summary;
      showErrorToast(summary);
      return;
    }

    statusBar.textContent = count === 1 ? "Deleted 1 media file" : `Deleted ${data.deleted_count || count} media files`;
  } catch (err) {
    const message = err?.message || "Failed to delete images";
    statusBar.textContent = `Delete error: ${message}`;
    showErrorToast(`Delete error: ${message}`);
  }
}

async function readNdjsonStream(response, onEvent) {
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";

  while (true) {
    const { value, done } = await reader.read();
    buffer += decoder.decode(value || new Uint8Array(), { stream: !done });

    let newlineIndex;
    while ((newlineIndex = buffer.indexOf("\n")) >= 0) {
      const line = buffer.slice(0, newlineIndex).trim();
      buffer = buffer.slice(newlineIndex + 1);
      if (!line) continue;
      onEvent(JSON.parse(line));
    }

    if (done) break;
  }

  const tail = buffer.trim();
  if (tail) {
    onEvent(JSON.parse(tail));
  }
}

async function saveOllamaSettingsFromForm() {
  const previousNetworkSettings = {
    httpsCertFile: state.httpsCertFile,
    httpsKeyFile: state.httpsKeyFile,
    httpsPort: state.httpsPort,
    remoteHttpMode: state.remoteHttpMode,
  };
  const server = settingsServerInput.value.trim() || "127.0.0.1";
  const port = Number(settingsPortInput.value || "11434") || 11434;
  const timeoutSeconds = Number(settingsTimeoutInput.value || "20") || 20;
  const model = String(settingsModelInput.value || "").trim() || "llava";
  const comfyuiServer = settingsComfyuiServerInput.value.trim() || "127.0.0.1";
  const comfyuiPort = Number(settingsComfyuiPortInput.value || "8188") || 8188;
  const comfyuiWorkflowPath = settingsComfyuiWorkflowPathInput.value.trim();
  const comfyuiOutputFolder = settingsComfyuiOutputFolderInput.value.trim();
  const cropAspectRatios = settingsCropAspectRatiosInput.value.split(",").map(s => s.trim()).filter(Boolean);
  let maskLatentBaseWidthPresets;
  const httpsCertFile = settingsHttpsCertInput.value.trim();
  const httpsKeyFile = settingsHttpsKeyInput.value.trim();
  const httpsPort = Number(settingsHttpsPortInput.value || "8900") || 8900;
  const remoteHttpMode = String(settingsRemoteHttpModeInput.value || "redirect-to-https").trim() || "redirect-to-https";
  const ffmpegPath = settingsFfmpegPathInput.value.trim();
  const processingReservedCores = Math.max(0, Number(settingsProcessingReservedCoresInput.value || "0") || 0);
  const ffmpegThreads = Math.max(0, Number(settingsFfmpegThreadsInput.value || "0") || 0);
  const ffmpegHwaccel = String(settingsFfmpegHwaccelInput.value || "auto").trim() || "auto";
  const promptTemplate = settingsPromptInput.value || state.ollamaPromptTemplate;
  const groupPromptTemplate = settingsGroupPromptInput.value || state.ollamaGroupPromptTemplate;
  const enableFreeText = settingsAutoFreeTextEnabled.checked;
  const freeTextPromptTemplate = settingsFreeTextPromptInput.value || state.ollamaFreeTextPromptTemplate;
  try {
    maskLatentBaseWidthPresets = parseMaskLatentBaseWidthPresetsInput(settingsMaskLatentBaseWidthPresetsInput.value);
  } catch (err) {
    setActiveSettingsTab("editing");
    statusBar.textContent = `Settings error: ${err.message}`;
    return;
  }
  let videoTrainingPresets;
  try {
    videoTrainingPresets = parseVideoTrainingPresetsText(settingsVideoPresetsInput.value);
    populateVideoTrainingProfileSelect(videoTrainingPresets, settingsVideoProfileInput.value || state.videoTrainingProfileKey);
    setVideoTrainingPresetsStatus();
  } catch (err) {
    setActiveSettingsTab("video-training");
    setVideoTrainingPresetsStatus(err.message || "Invalid video training preset JSON.", { isError: true });
    statusBar.textContent = `Settings error: ${err.message}`;
    return;
  }
  const selectedVideoTrainingProfile = state.folder
    ? (findVideoTrainingProfileByKey(settingsVideoProfileInput.value, videoTrainingPresets) || videoTrainingPresets[0] || null)
    : null;
  const settingsPayload = {
    https_certfile: httpsCertFile,
    https_keyfile: httpsKeyFile,
    https_port: httpsPort,
    remote_http_mode: remoteHttpMode,
    ffmpeg_path: ffmpegPath,
    processing_reserved_cores: processingReservedCores,
    ffmpeg_threads: ffmpegThreads,
    ffmpeg_hwaccel: ffmpegHwaccel,
    ollama_server: server,
    ollama_port: port,
    ollama_timeout_seconds: timeoutSeconds,
    ollama_model: model,
    comfyui_server: comfyuiServer,
    comfyui_port: comfyuiPort,
    comfyui_workflow_path: comfyuiWorkflowPath,
    comfyui_output_folder: comfyuiOutputFolder,
    crop_aspect_ratios: cropAspectRatios,
    mask_latent_base_width_presets: maskLatentBaseWidthPresets,
    video_training_presets: videoTrainingPresets,
    ollama_prompt_template: promptTemplate,
    ollama_group_prompt_template: groupPromptTemplate,
    ollama_enable_free_text: enableFreeText,
    ollama_free_text_prompt_template: freeTextPromptTemplate,
  };
  if (state.folder) {
    settingsPayload.folder = state.folder;
    settingsPayload.video_training_profile_key = selectedVideoTrainingProfile?.key || "";
  }
  try {
    const resp = await fetch("/api/settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(settingsPayload),
    });
    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      throw new Error(data.detail || "Failed to save settings");
    }

    let nextSettings = null;
    const refreshUrl = state.folder
      ? `/api/settings?folder=${encodeURIComponent(state.folder)}`
      : "/api/settings";
    const refreshResp = await fetch(refreshUrl);
    if (refreshResp.ok) {
      nextSettings = await refreshResp.json();
    }
    if (!nextSettings) {
      nextSettings = {
        ...settingsPayload,
        video_training_profile: selectedVideoTrainingProfile,
      };
    }
    applySettings(nextSettings);
    const networkChanged = previousNetworkSettings.httpsCertFile !== state.httpsCertFile
      || previousNetworkSettings.httpsKeyFile !== state.httpsKeyFile
      || previousNetworkSettings.httpsPort !== state.httpsPort
      || previousNetworkSettings.remoteHttpMode !== state.remoteHttpMode;
    updateActionButtons();
    closeSettingsModal();
    statusBar.textContent = networkChanged
      ? "Settings saved. Restart the server to apply HTTP/HTTPS changes."
      : `Settings saved. ffmpeg ${state.ffmpegHwaccel === "off" ? "CPU" : "auto GPU"}, threads ${state.ffmpegThreads || "auto"}, reserve ${state.processingReservedCores}`;
  } catch (err) {
    statusBar.textContent = `Settings error: ${err.message}`;
  }
}

function normalizeFolderPathForCompare(value) {
  return String(value || "")
    .trim()
    .replace(/[\\/]+/g, "/")
    .replace(/\/$/, "")
    .toLowerCase();
}

function renderFolderAutocomplete() {
  const autocomplete = state.folderAutocomplete;
  const items = Array.isArray(autocomplete.items) ? autocomplete.items : [];
  const visible = autocomplete.visible && items.length > 0;

  folderSuggestionsList.replaceChildren();

  if (visible) {
    items.forEach((item, index) => {
      const option = document.createElement("div");
      option.id = `folder-suggestion-option-${index}`;
      option.className = "folder-suggestion";
      if (index === autocomplete.highlightedIndex) {
        option.classList.add("active");
      }
      option.setAttribute("role", "option");
      option.setAttribute("aria-selected", index === autocomplete.highlightedIndex ? "true" : "false");

      const name = document.createElement("div");
      name.className = "folder-suggestion-name";
      name.textContent = item.name || item.path || "";
      option.appendChild(name);

      if (item.parent || item.path !== item.name) {
        const meta = document.createElement("div");
        meta.className = "folder-suggestion-meta";
        meta.textContent = item.parent || item.path || "";
        option.appendChild(meta);
      }

      option.addEventListener("mousedown", (event) => {
        event.preventDefault();
      });
      option.addEventListener("mouseenter", () => {
        if (state.folderAutocomplete.highlightedIndex === index) return;
        state.folderAutocomplete.highlightedIndex = index;
        renderFolderAutocomplete();
      });
      option.addEventListener("click", () => {
        applyFolderAutocompleteSelection(index);
      });

      folderSuggestionsList.appendChild(option);
    });
  }

  folderSuggestionsList.hidden = !visible;
  folderSuggestionsList.classList.toggle("open", visible);
  folderInput.setAttribute("aria-expanded", String(visible));

  if (!visible || autocomplete.highlightedIndex < 0 || autocomplete.highlightedIndex >= items.length) {
    folderInput.removeAttribute("aria-activedescendant");
    return;
  }

  const activeOptionId = `folder-suggestion-option-${autocomplete.highlightedIndex}`;
  folderInput.setAttribute("aria-activedescendant", activeOptionId);
  const activeOption = document.getElementById(activeOptionId);
  if (activeOption) {
    activeOption.scrollIntoView({ block: "nearest" });
  }
}

function hideFolderAutocomplete() {
  state.folderAutocomplete.visible = false;
  renderFolderAutocomplete();
}

function clearFolderAutocomplete(options = {}) {
  const { cancelPending = false } = options;
  const autocomplete = state.folderAutocomplete;
  if (cancelPending) {
    if (autocomplete.debounceTimer) {
      window.clearTimeout(autocomplete.debounceTimer);
      autocomplete.debounceTimer = 0;
    }
    if (autocomplete.abortController) {
      autocomplete.abortController.abort();
      autocomplete.abortController = null;
    }
  }
  autocomplete.items = [];
  autocomplete.highlightedIndex = -1;
  autocomplete.visible = false;
  renderFolderAutocomplete();
}

function setFolderAutocompleteItems(items) {
  const autocomplete = state.folderAutocomplete;
  const normalizedItems = Array.isArray(items) ? items.filter(item => item?.path) : [];
  const normalizedInputValue = normalizeFolderPathForCompare(folderInput.value);
  const exactIndex = normalizedItems.findIndex((item) => normalizeFolderPathForCompare(item.path) === normalizedInputValue);
  autocomplete.items = normalizedItems;
  autocomplete.highlightedIndex = normalizedItems.length === 0 ? -1 : Math.max(0, exactIndex);
  autocomplete.visible = normalizedItems.length > 0;
  renderFolderAutocomplete();
}

function moveFolderAutocompleteHighlight(delta) {
  const autocomplete = state.folderAutocomplete;
  const items = autocomplete.items || [];
  if (!items.length) return false;

  let nextIndex = autocomplete.highlightedIndex;
  if (nextIndex < 0) {
    nextIndex = delta >= 0 ? 0 : items.length - 1;
  } else {
    nextIndex = (nextIndex + delta + items.length) % items.length;
  }

  autocomplete.highlightedIndex = nextIndex;
  autocomplete.visible = true;
  renderFolderAutocomplete();
  return true;
}

function applyFolderAutocompleteSelection(index) {
  const item = state.folderAutocomplete.items[index];
  if (!item?.path) return false;
  folderInput.value = item.path;
  hideFolderAutocomplete();
  folderInput.focus();
  return true;
}

async function fetchFolderAutocompleteSuggestions(query) {
  const trimmedQuery = String(query || "").trim();
  if (!trimmedQuery) {
    clearFolderAutocomplete({ cancelPending: true });
    return;
  }

  const autocomplete = state.folderAutocomplete;
  if (autocomplete.abortController) {
    autocomplete.abortController.abort();
    autocomplete.abortController = null;
  }

  const requestSeq = ++autocomplete.requestSeq;
  const controller = new AbortController();
  autocomplete.abortController = controller;

  try {
    const resp = await fetch(`/api/folders/suggest?query=${encodeURIComponent(trimmedQuery)}`, {
      signal: controller.signal,
    });
    const data = await resp.json().catch(() => ({}));
    if (requestSeq !== autocomplete.requestSeq) return;
    if (!resp.ok) {
      clearFolderAutocomplete();
      return;
    }
    setFolderAutocompleteItems(data.suggestions || []);
  } catch (err) {
    if (err?.name === "AbortError") return;
    if (requestSeq !== autocomplete.requestSeq) return;
    clearFolderAutocomplete();
  } finally {
    if (autocomplete.abortController === controller) {
      autocomplete.abortController = null;
    }
  }
}

function scheduleFolderAutocompleteRefresh(options = {}) {
  const { immediate = false } = options;
  const autocomplete = state.folderAutocomplete;
  if (autocomplete.debounceTimer) {
    window.clearTimeout(autocomplete.debounceTimer);
    autocomplete.debounceTimer = 0;
  }

  autocomplete.debounceTimer = window.setTimeout(() => {
    autocomplete.debounceTimer = 0;
    fetchFolderAutocompleteSuggestions(folderInput.value);
  }, immediate ? 0 : 120);
}

function handleFolderInputInput() {
  scheduleFolderAutocompleteRefresh();
}

function handleFolderInputFocus() {
  if (!folderInput.value.trim()) {
    clearFolderAutocomplete({ cancelPending: true });
    return;
  }
  scheduleFolderAutocompleteRefresh({ immediate: true });
}

function handleFolderInputBlur() {
  window.setTimeout(() => {
    if (document.activeElement === folderInput) return;
    hideFolderAutocomplete();
  }, 120);
}

function handleFolderInputKeydown(event) {
  if (event.key === "ArrowDown") {
    if (moveFolderAutocompleteHighlight(1)) {
      event.preventDefault();
    }
    return;
  }

  if (event.key === "ArrowUp") {
    if (moveFolderAutocompleteHighlight(-1)) {
      event.preventDefault();
    }
    return;
  }

  if (event.key === "Escape") {
    if (state.folderAutocomplete.visible) {
      event.preventDefault();
      hideFolderAutocomplete();
    }
    return;
  }

  if (event.key === "Tab") {
    const highlightedItem = state.folderAutocomplete.items[state.folderAutocomplete.highlightedIndex];
    if (state.folderAutocomplete.visible && highlightedItem) {
      applyFolderAutocompleteSelection(state.folderAutocomplete.highlightedIndex);
    }
    return;
  }

  if (event.key !== "Enter") return;

  event.preventDefault();
  const highlightedItem = state.folderAutocomplete.items[state.folderAutocomplete.highlightedIndex];
  if (state.folderAutocomplete.visible && highlightedItem) {
    const currentValue = normalizeFolderPathForCompare(folderInput.value);
    const highlightedValue = normalizeFolderPathForCompare(highlightedItem.path);
    if (currentValue !== highlightedValue) {
      applyFolderAutocompleteSelection(state.folderAutocomplete.highlightedIndex);
      return;
    }
  }

  clearFolderAutocomplete({ cancelPending: true });
  loadFolder();
}

// ===== FOLDER LOADING =====
async function loadFolder(options = {}) {
  const { preserveScrollTop = null } = options;
  const folder = folderInput.value.trim();
  if (!folder) return;
  clearFolderAutocomplete({ cancelPending: true });
  state.folder = folder;
  state.selectedPaths.clear();
  state.lastClickedIndex = -1;
  state.lastClickedPath = null;
  state.previewPath = null;
  state.captionCache = {};
  state.metadataCache = {};
  state.activeSentenceFilters.clear();
  state.activeMetaFilters.aspectState = "any";
  state.activeMetaFilters.maskState = "any";
  state.activeMetaFilters.captionState = "any";
  state.filterCaptionCacheKey = "";
  state.filterLoadingPromise = null;
  state.imageCrops = {};
  state.imageVersions = {};
  state.imageMaskVersions = {};
  state.thumbnailDimensions = {};
  state.metadataSaving = false;
  statusBar.textContent = "Loading...";

  try {
    const resp = await fetch(`/api/list-images?folder=${encodeURIComponent(folder)}`);
    if (!resp.ok) { const e = await resp.json(); throw new Error(e.detail); }
    const data = await resp.json();
    state.images = data.images;
    state.folder = data.folder;
    state.imageVersions = Object.fromEntries((state.images || []).map(img => [img.path, img.mtime || 0]));
    folderInput.value = data.folder;
    updateFileCountDisplay();
    statusBar.textContent = `Loaded ${state.images.length} media files`;

    // Save last_folder and load per-folder sentences
    fetch("/api/settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ last_folder: state.folder }),
    }).catch(() => {});

    const settingsResp = await fetch(`/api/settings?folder=${encodeURIComponent(state.folder)}`);
    if (settingsResp.ok) {
      const settings = await settingsResp.json();
      applySettings(settings);
    }

    renderGrid({ preserveScrollTop });
    hidePreview();
    renderSentences();
    renderMetadataEditor();
    freeText.value = "";
    updateActionButtons();
    freeText.disabled = true;

    // Preload all thumbnails
    const thumbSize = getThumbLoadSize();
    let queuedThumbnailCount = 0;
    for (const img of state.images) {
      if (queueThumbLoad(img.path, thumbSize)) {
        queuedThumbnailCount += 1;
      }
    }
    if (queuedThumbnailCount > 0) {
      startThumbnailProgress(queuedThumbnailCount, `Processing thumbnails 0/${queuedThumbnailCount}`);
    } else {
      resetThumbnailProgress();
    }
  } catch (err) {
    resetThumbnailProgress();
    statusBar.textContent = `Error: ${err.message}`;
  }
}

function getImageExtension(name) {
  const normalizedName = String(name || "").trim().toLowerCase();
  const dotIndex = normalizedName.lastIndexOf(".");
  return dotIndex >= 0 ? normalizedName.slice(dotIndex) : "";
}

function getDroppedImageFiles(fileList) {
  return Array.from(fileList || []).filter(file => MEDIA_FILE_EXTENSIONS.has(getImageExtension(file?.name)));
}

function cancelThumbnailProgressHide() {
  if (thumbnailProgressHideTimer) {
    clearTimeout(thumbnailProgressHideTimer);
    thumbnailProgressHideTimer = null;
  }
}

function updateThumbnailProgressBar() {
  const progress = state.thumbnailProgress;
  const visible = !!progress.visible;
  thumbnailQueueStatus.hidden = !visible;
  thumbnailQueueStatus.classList.toggle("visible", visible);
  renderToolbarStatusVisibility();
  if (!visible) {
    thumbnailQueueText.textContent = "";
    thumbnailQueueProgressFill.style.width = "0%";
    thumbnailQueueProgressFill.classList.remove("active");
    return;
  }

  const total = Math.max(0, Number(progress.total || 0));
  const completed = Math.max(0, Math.min(total, Number(progress.completed || 0)));
  const percent = total > 0 ? (completed / total) * 100 : 0;
  thumbnailQueueText.textContent = progress.label || `Processing thumbnails ${completed}/${total}`;
  thumbnailQueueProgressFill.style.width = `${percent}%`;
  thumbnailQueueProgressFill.classList.toggle("active", completed < total);
}

function startThumbnailProgress(total, label) {
  cancelThumbnailProgressHide();
  state.thumbnailProgress.visible = total > 0;
  state.thumbnailProgress.total = Math.max(0, Number(total || 0));
  state.thumbnailProgress.completed = 0;
  state.thumbnailProgress.label = label || `Processing thumbnails 0/${state.thumbnailProgress.total}`;
  updateThumbnailProgressBar();
}

function advanceThumbnailProgress() {
  if (!state.thumbnailProgress.visible) {
    return;
  }
  state.thumbnailProgress.completed = Math.min(
    state.thumbnailProgress.total,
    state.thumbnailProgress.completed + 1,
  );
  state.thumbnailProgress.label = `Processing thumbnails ${state.thumbnailProgress.completed}/${state.thumbnailProgress.total}`;
  updateThumbnailProgressBar();
  if (state.thumbnailProgress.completed >= state.thumbnailProgress.total) {
    state.thumbnailProgress.label = `Processed ${state.thumbnailProgress.total} thumbnails`;
    updateThumbnailProgressBar();
    cancelThumbnailProgressHide();
    thumbnailProgressHideTimer = window.setTimeout(() => {
      state.thumbnailProgress.visible = false;
      state.thumbnailProgress.label = "";
      updateThumbnailProgressBar();
      thumbnailProgressHideTimer = null;
    }, 1200);
  }
}

function resetThumbnailProgress() {
  cancelThumbnailProgressHide();
  state.thumbnailProgress.visible = false;
  state.thumbnailProgress.label = "";
  state.thumbnailProgress.total = 0;
  state.thumbnailProgress.completed = 0;
  updateThumbnailProgressBar();
}

function setFileDropActive(active) {
  previewStage.classList.toggle("file-drop-active", !!active);
  if (fileDropHint) {
    fileDropHint.hidden = !active;
  }
}

function normalizeClientPath(path) {
  return String(path || "").trim().replace(/\\/g, "/").replace(/\/+/g, "/").toLowerCase();
}

function isSameClientPath(left, right) {
  return normalizeClientPath(left) === normalizeClientPath(right);
}

function getUploadQueuePendingImageCount() {
  return state.uploadQueue.reduce((sum, job) => sum + (job?.imageFiles?.length || 0), 0);
}

function getUploadQueueTotalImageCount() {
  return state.uploadQueueCompletedImages
    + (state.uploadQueueCurrentJob?.imageFiles?.length || 0)
    + getUploadQueuePendingImageCount();
}

function renderUploadQueueStatus() {
  const hasActiveQueue = !!(state.uploading || state.uploadQueueCurrentJob || state.uploadQueue.length);
  uploadQueueStatus.hidden = !hasActiveQueue;
  uploadQueueStatus.classList.toggle("visible", hasActiveQueue);
  renderToolbarStatusVisibility();

  if (!hasActiveQueue) {
    uploadQueueProgressFill.style.width = "0%";
    uploadQueueProgressFill.classList.remove("active");
    uploadQueueText.textContent = "";
    return;
  }

  const currentJob = state.uploadQueueCurrentJob;
  const waitingJobs = state.uploadQueue.length;
  const completedImages = state.uploadQueueCompletedImages;
  const totalImages = getUploadQueueTotalImageCount();
  const progressPercent = totalImages > 0 ? Math.max(0, Math.min(100, (completedImages / totalImages) * 100)) : 0;
  uploadQueueProgressFill.style.width = `${progressPercent}%`;
  uploadQueueProgressFill.classList.toggle("active", !!currentJob);

  const parts = [];
  if (currentJob) {
    parts.push(`Uploading ${currentJob.imageFiles.length} media file${currentJob.imageFiles.length === 1 ? "" : "s"} to ${getFileLabel(currentJob.folder)}`);
  } else if (waitingJobs > 0) {
    parts.push(`Starting next upload job`);
  } else {
    parts.push(state.uploadQueueLastSummary || "Upload queue complete");
  }
  if (waitingJobs > 0) {
    parts.push(`${waitingJobs} queued job${waitingJobs === 1 ? "" : "s"}`);
  }
  if (totalImages > 0) {
    parts.push(`${completedImages}/${totalImages} done`);
  }
  if (state.uploadQueueFailedJobs > 0) {
    parts.push(`${state.uploadQueueFailedJobs} failed`);
  }
  uploadQueueText.textContent = parts.join(" • ");
}

async function selectUploadedImages(paths) {
  const uploadedPaths = Array.from(new Set((paths || []).filter(Boolean)));
  if (!uploadedPaths.length) {
    return;
  }

  const availablePaths = uploadedPaths.filter(path => state.images.some(img => img.path === path));
  if (!availablePaths.length) {
    return;
  }

  state.selectedPaths.clear();
  for (const path of availablePaths) {
    state.selectedPaths.add(path);
  }

  const lastPath = availablePaths[availablePaths.length - 1];
  state.lastClickedPath = lastPath;
  state.lastClickedIndex = state.images.findIndex(img => img.path === lastPath);

  updateGridSelection();
  renderGrid({ preservePath: lastPath, preserveScrollTop: fileGridContainer.scrollTop });

  if (availablePaths.length === 1) {
    await showPreview(lastPath);
    await Promise.all([
      loadCaptionData(lastPath),
      loadCropData(lastPath),
      loadMetadataData(lastPath),
    ]);
    freeText.disabled = false;
  } else {
    await showPreview(lastPath);
    freeText.disabled = true;
    freeText.value = "(Multiple media files selected)";
    await Promise.all([
      loadMultiCaptionState(),
      loadMultiMetadataState(),
    ]);
    clearCropDraft();
  }

  renderMetadataEditor();
  updateMultiInfo();
  renderSentences();
  updateActionButtons();
  preloadAdjacent(lastPath);
}

async function processUploadQueue() {
  if (state.uploading) {
    renderUploadQueueStatus();
    return;
  }
  if (!state.uploadQueue.length) {
    renderUploadQueueStatus();
    updateActionButtons();
    return;
  }

  state.uploading = true;
  renderUploadQueueStatus();
  updateActionButtons();

  while (state.uploadQueue.length) {
    const job = state.uploadQueue.shift();
    state.uploadQueueCurrentJob = job;
    renderUploadQueueStatus();

    try {
      const formData = new FormData();
      formData.append("folder", job.folder);
      for (const file of job.imageFiles) {
        formData.append("files", file, file.name);
      }

      const resp = await fetch("/api/images/upload", {
        method: "POST",
        body: formData,
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        throw new Error(data.detail || "Upload failed");
      }

      const uploadedPaths = Array.isArray(data.uploaded)
        ? data.uploaded.map(item => item?.path).filter(Boolean)
        : [];
      const uploadedCount = Number(data.uploaded_count || uploadedPaths.length || 0);
      const renamedCount = Number(data.renamed_count || 0);
      const serverSkippedCount = Array.isArray(data.skipped) ? data.skipped.length : Number(data.skipped_count || 0);
      const totalSkippedCount = job.clientSkippedCount + serverSkippedCount;

      state.uploadQueueCompletedImages += uploadedCount;

      if (uploadedCount > 0 && isSameClientPath(state.folder, job.folder)) {
        const preserveScrollTop = fileGridContainer.scrollTop;
        folderInput.value = job.folder;
        await loadFolder({ preserveScrollTop });
        await selectUploadedImages(uploadedPaths);
      }

      const summary = [`Uploaded ${uploadedCount} image${uploadedCount === 1 ? "" : "s"} to ${getFileLabel(job.folder)}`];
      summary[0] = `Uploaded ${uploadedCount} media file${uploadedCount === 1 ? "" : "s"} to ${getFileLabel(job.folder)}`;
      if (renamedCount > 0) {
        summary.push(`${renamedCount} renamed`);
      }
      if (totalSkippedCount > 0) {
        summary.push(`${totalSkippedCount} skipped`);
      }
      if (state.uploadQueue.length > 0) {
        summary.push(`${state.uploadQueue.length} queued`);
      }
      state.uploadQueueLastSummary = summary.join(" • ");

      if (uploadedCount === 0 && totalSkippedCount > 0) {
        showErrorToast(`No media files were uploaded to ${getFileLabel(job.folder)}. ${totalSkippedCount} file${totalSkippedCount === 1 ? " was" : "s were"} skipped.`);
      }
    } catch (err) {
      state.uploadQueueFailedJobs += 1;
      state.uploadQueueLastSummary = `Upload error for ${getFileLabel(job.folder)}: ${err.message}`;
      showErrorToast(`Upload failed for ${getFileLabel(job.folder)}: ${err.message}`);
    }

    state.uploadQueueCurrentJob = null;
    renderUploadQueueStatus();
  }

  state.uploading = false;
  renderUploadQueueStatus();
  updateActionButtons();
  statusBar.textContent = state.uploadQueueLastSummary || "Upload queue complete";
  state.uploadQueueCompletedImages = 0;
  state.uploadQueueFailedJobs = 0;
  state.uploadQueueLastSummary = "";
}

async function uploadDroppedFiles(fileList) {
  if (!state.folder) {
    showErrorToast("Load a folder first, then drop media files onto the preview panel.");
    statusBar.textContent = "Upload error: no folder loaded";
    return;
  }

  const droppedFiles = Array.from(fileList || []);
  if (!droppedFiles.length) {
    return;
  }

  const imageFiles = getDroppedImageFiles(droppedFiles);
  const clientSkippedCount = Math.max(0, droppedFiles.length - imageFiles.length);
  if (!imageFiles.length) {
    showErrorToast("No supported media files were dropped.");
    statusBar.textContent = clientSkippedCount > 0
      ? `Upload skipped ${clientSkippedCount} unsupported file${clientSkippedCount === 1 ? "" : "s"}`
      : "Upload skipped";
    return;
  }

  if (!state.uploading && !state.uploadQueue.length && !state.uploadQueueCurrentJob) {
    state.uploadQueueCompletedImages = 0;
    state.uploadQueueFailedJobs = 0;
    state.uploadQueueLastSummary = "";
  }

  state.uploadQueue.push({
    id: ++uploadJobCounter,
    folder: state.folder,
    imageFiles,
    clientSkippedCount,
  });

  renderUploadQueueStatus();
  processUploadQueue().catch((err) => {
    state.uploading = false;
    state.uploadQueueCurrentJob = null;
    state.uploadQueueFailedJobs += 1;
    state.uploadQueueLastSummary = `Upload queue failed: ${err.message}`;
    renderUploadQueueStatus();
    updateActionButtons();
    statusBar.textContent = state.uploadQueueLastSummary;
    showErrorToast(state.uploadQueueLastSummary);
  });
}

function getThumbLoadSize() {
  const s = state.thumbSize;
  if (s <= 80) return 128;
  if (s <= 180) return 256;
  return 400;
}

// ===== GRID RENDERING =====
function restoreGridViewport(path, previousScrollTop = null) {
  const targetPath = path || state.previewPath || state.lastClickedPath || [...state.selectedPaths][0] || null;
  if (targetPath) {
    const cell = fileGrid.querySelector(`.thumb-cell[data-path="${CSS.escape(targetPath)}"]`);
    if (cell) {
      cell.scrollIntoView({ block: "nearest" });
      return;
    }
  }
  if (previousScrollTop !== null && previousScrollTop !== undefined) {
    fileGridContainer.scrollTop = previousScrollTop;
  }
}

function renderGrid(options = {}) {
  const { preservePath = null, preserveScrollTop = null } = options;
  fileGrid.innerHTML = "";
  const size = state.thumbSize;
  const thumbLoadSize = getThumbLoadSize();
  const visibleEntries = getVisibleImageEntries();

  updateFileCountDisplay();
  renderFilterActions();

  for (const { img, index } of visibleEntries) {
    const cell = document.createElement("div");
    let cls = "thumb-cell";
    if (state.selectedPaths.has(img.path)) cls += " selected";
    if (img.has_caption) cls += " has-caption";
    if (img.has_mask) cls += " has-mask";
    if (!imageConformsToAspectRatios(img)) cls += " aspect-mismatch";
    cell.className = cls;
    cell.style.width = size + "px";
    cell.style.height = size + "px";
    cell.dataset.index = index;
    cell.dataset.path = img.path;
    cell.dataset.mediaType = img.media_type || getMediaType(img.path);

    const imgEl = document.createElement("img");
    imgEl.dataset.path = img.path;
    imgEl.loading = "lazy";
    imgEl.addEventListener("load", () => {
      storeThumbnailDimensions(img.path, imgEl.naturalWidth, imgEl.naturalHeight);
    });
    const cachedUrl = thumbBlobCache.get(`${img.path}:${thumbLoadSize}:${getImageVersion(img.path)}`);
    if (cachedUrl) {
      imgEl.src = cachedUrl;
    } else {
      imgEl.src = buildImageApiUrl("thumbnail", img.path, { size: thumbLoadSize });
    }
    cell.appendChild(imgEl);

    const nameEl = document.createElement("div");
    nameEl.className = "thumb-name";
    nameEl.textContent = img.name;
    nameEl.title = img.name;
    cell.appendChild(nameEl);

    const maskBadge = document.createElement("div");
    maskBadge.className = "mask-badge";
    maskBadge.textContent = "M";
    maskBadge.title = `${img.media_type === "video" ? "Video" : "Media"} mask sidecar exists`;
    cell.appendChild(maskBadge);

    const dot = document.createElement("div");
    dot.className = "caption-dot";
    dot.textContent = "TXT";
    dot.title = `${img.media_type === "video" ? "Video" : "Media"} caption text file exists`;
    dot.dataset.captionDot = img.path;
    dot.addEventListener("dblclick", (e) => {
      e.preventDefault();
      e.stopPropagation();
      fetch(`/api/open-file?path=${encodeURIComponent(getCaptionPathForImage(img.path))}`);
    });
    cell.appendChild(dot);

    const aspectWarning = document.createElement("div");
    aspectWarning.className = "aspect-warning";
    aspectWarning.textContent = "AR";
    aspectWarning.title = "Media aspect ratio does not match the configured aspect ratio list";
    cell.appendChild(aspectWarning);

    cell.addEventListener("click", (e) => handleThumbClick(index, e));
    cell.addEventListener("dblclick", (e) => {
      e.preventDefault();
      fetch(`/api/open-in-explorer?path=${encodeURIComponent(img.path)}`);
    });
    fileGrid.appendChild(cell);
  }

  restoreGridViewport(preservePath, preserveScrollTop);
}

function markCaptionIndicator(path, hasCaption) {
  // Update the has_caption flag on the image object
  const img = state.images.find(i => i.path === path);
  if (img) img.has_caption = hasCaption;
  updateFileCountDisplay();
  // Update DOM
  const cell = fileGrid.querySelector(`[data-path="${CSS.escape(path)}"]`);
  if (cell) {
    cell.classList.toggle("has-caption", hasCaption);
  }
}

function updateGridSelection() {
  const cells = fileGrid.querySelectorAll(".thumb-cell");
  cells.forEach(cell => {
    const path = cell.dataset.path;
    cell.classList.toggle("selected", state.selectedPaths.has(path));
  });
}

function refreshVisibleThumbnail(path) {
  const thumbLoadSize = getThumbLoadSize();
  const nextSrc = buildImageApiUrl("thumbnail", path, { size: thumbLoadSize });
  const imgs = fileGrid.querySelectorAll(`img[data-path="${CSS.escape(path)}"]`);
  imgs.forEach((imgEl) => {
    imgEl.src = nextSrc;
  });
}

// ===== SELECTION HANDLING =====
function handleThumbClick(index, event) {
  const img = state.images[index];
  if (!img) return;
  const visibleEntries = getVisibleImageEntries();
  const currentVisibleIndex = visibleEntries.findIndex(entry => entry.img.path === img.path);
  const lastVisibleIndex = visibleEntries.findIndex(entry => entry.img.path === state.lastClickedPath);
  const isPlainReselect = !event.ctrlKey
    && !event.metaKey
    && !event.shiftKey
    && state.selectedPaths.size === 1
    && state.selectedPaths.has(img.path);

  if (event.ctrlKey || event.metaKey) {
    // Toggle individual selection
    if (state.selectedPaths.has(img.path)) {
      state.selectedPaths.delete(img.path);
    } else {
      state.selectedPaths.add(img.path);
    }
    state.lastClickedIndex = index;
    state.lastClickedPath = img.path;
  } else if (event.shiftKey && lastVisibleIndex >= 0 && currentVisibleIndex >= 0) {
    // Range select
    const start = Math.min(lastVisibleIndex, currentVisibleIndex);
    const end = Math.max(lastVisibleIndex, currentVisibleIndex);
    if (!event.ctrlKey) state.selectedPaths.clear();
    for (let i = start; i <= end; i++) {
      state.selectedPaths.add(visibleEntries[i].img.path);
    }
    state.lastClickedIndex = index;
    state.lastClickedPath = img.path;
  } else {
    // Single select
    state.selectedPaths.clear();
    state.selectedPaths.add(img.path);
    state.lastClickedIndex = index;
    state.lastClickedPath = img.path;
  }

  updateGridSelection();

  // Show preview of last clicked (or first selected)
  if (state.selectedPaths.size === 1) {
    const path = [...state.selectedPaths][0];
    if (isPlainReselect) {
      bumpImageVersion(path);
      invalidateImageCaches(path);
      refreshVisibleThumbnail(path);
    }
    showPreview(path, { preserveView: isPlainReselect });
    loadCaptionData(path);
    loadMetadataData(path);
    loadCropData(path);
    freeText.disabled = false;
  } else if (state.selectedPaths.size > 1) {
    // Show preview of clicked image
    showPreview(img.path);
    freeText.disabled = true;
    freeText.value = "(Multiple media files selected)";
    loadMultiCaptionState();
    loadMultiMetadataState();
    clearCropDraft();
  } else {
    hidePreview();
    freeText.disabled = true;
    freeText.value = "";
  }

  renderMetadataEditor();
  updateMultiInfo();
  renderSentences();
  updateActionButtons();

  // Preload adjacent images
  preloadAdjacent(img.path);
}

function preloadAdjacent(path) {
  const visibleEntries = getVisibleImageEntries();
  let index = visibleEntries.findIndex(entry => entry.img.path === path);
  if (index < 0) {
    index = state.images.findIndex(img => img.path === path);
    for (let d = -3; d <= 3; d++) {
      const i = index + d;
      if (i >= 0 && i < state.images.length) {
        preloadPreview(state.images[i].path);
      }
    }
    return;
  }
  for (let d = -3; d <= 3; d++) {
    const i = index + d;
    if (i >= 0 && i < visibleEntries.length) {
      preloadPreview(visibleEntries[i].img.path);
    }
  }
}

// ===== PREVIEW =====
// Zoom/pan state
let zoomLevel = 1;
let panX = 0, panY = 0;
let imgNatW = 0, imgNatH = 0;
let userHasZoomed = false; // tracks if user manually zoomed/panned
let videoTimelineInteraction = null;

function resetZoomPan() {
  userHasZoomed = false;
  fitImageToPanel();
}

function fitImageToPanel() {
  if (!imgNatW || !imgNatH) return;
  const panel = previewStage;
  const pw = panel.clientWidth;
  const ph = panel.clientHeight;
  const scale = Math.min(pw / imgNatW, ph / imgNatH);
  zoomLevel = scale;
  panX = (pw - imgNatW * scale) / 2;
  panY = (ph - imgNatH * scale) / 2;
  applyTransform();
}

function applyTransform() {
  const previewEl = getActivePreviewElement();
  if (previewEl) {
    applyTransformToElement(previewEl);
  }
  if (state.previewMediaType === "video" && state.maskEditor.active && state.maskEditor.videoSnapshotUrl) {
    applyTransformToElement(previewImg);
  }
  applyImageEditCanvasTransform();
  applyMaskCanvasTransform();
  applyMaskLatentPreviewTransform();
  renderMaskSignalProbeOverlay();
  renderCropOverlay();
  renderPreviewCaptionOverlay();
}

async function showPreview(path, options = {}) {
  const { preserveView = false } = options;
  const previousPath = state.previewPath;
  if (state.promptPreview.sourcePath && state.promptPreview.sourcePath !== path) {
    resetPromptPreviewState();
  }
  if (previousPath && previousPath !== path && state.maskEditor.active) {
    try {
      await saveMaskEdit();
    } catch {
      return;
    }
  }
  if (previousPath && previousPath !== path && state.cropDraft) {
    if (isVideoMediaPath(previousPath)) {
      clearCropDraft();
    } else if (state.cropDirty) {
      await applyCropDraftWithOptions({ reopenPreview: false });
    } else {
      clearCropDraft();
    }
  }
  state.previewPath = path;
  state.previewMediaType = getMediaType(path);
  clearCropGuide();

  if (state.previewMediaType === "video") {
    if (state.maskEditor.active) {
      closeMaskEditor();
    }
    clearCropDraft();
  }

  const cachedPreview = previewCache.get(path);
  const thumbLoadSize = getThumbLoadSize();
  const cachedThumb = thumbBlobCache.get(`${path}:${thumbLoadSize}:${getImageVersion(path)}`);

  if (state.previewMediaType === "video") {
    previewImg.style.display = "none";
    const videoSrc = buildImageApiUrl("media", path);
    const posterSrc = cachedPreview || cachedThumb || buildImageApiUrl("preview", path);
    previewVideo.style.display = "none";
    previewVideo.removeAttribute("poster");
    previewVideo.dataset.path = path;
    const finishVideoPreviewLoad = () => {
      if (state.previewPath !== path || state.previewMediaType !== "video") {
        return;
      }
      previewVideo.style.display = "block";
      syncPreviewVideoPlaybackState();
      renderPreviewCaptionOverlay();
    };
    previewVideo.onloadedmetadata = () => {
      if (state.previewPath !== path || state.previewMediaType !== "video") {
        return;
      }
      imgNatW = previewVideo.videoWidth || 0;
      imgNatH = previewVideo.videoHeight || 0;
      resetZoomPan();
      renderMaskEditorUi();
      ensureVideoMetaLoaded(path)
        .then(() => ensureVideoTimelineLoaded(path))
        .then(() => renderVideoEditPanel())
        .catch((err) => showErrorToast(err.message || "Failed to load video editing metadata"));
      syncPreviewVideoPlaybackState();
      renderPreviewCaptionOverlay();
      if (previewVideo.readyState >= 2) {
        finishVideoPreviewLoad();
      }
    };
    previewVideo.onloadeddata = finishVideoPreviewLoad;
    if (previewVideo.currentSrc !== videoSrc && previewVideo.src !== videoSrc) {
      previewVideo.pause();
      previewVideo.removeAttribute("src");
      previewVideo.load();
      previewVideo.poster = posterSrc;
      previewVideo.src = videoSrc;
      previewVideo.load();
    } else if (previewVideo.readyState >= 1) {
      previewVideo.poster = posterSrc;
      imgNatW = previewVideo.videoWidth || 0;
      imgNatH = previewVideo.videoHeight || 0;
      resetZoomPan();
      if (previewVideo.readyState >= 2) {
        finishVideoPreviewLoad();
      }
    }
    preloadPreview(path);
    ensureVideoMetaLoaded(path)
      .then(() => ensureVideoTimelineLoaded(path))
      .then(() => renderVideoEditPanel())
      .catch((err) => showErrorToast(err.message || "Failed to load video editing metadata"));
  } else {
    stopPreviewVideo({ clearSource: true });

    const loadSrc = (src) => {
      if (previewImg.src === src && imgNatW) {
        return;
      }
      const previousViewState = preserveView ? capturePreviewViewState() : null;
      previewImg.style.display = "none";
      previewImg.onload = () => {
        restorePreviewViewState(previousViewState);
        renderPreviewCaptionOverlay();
      };
      previewImg.src = src;
    };

    if (cachedPreview) {
      loadSrc(cachedPreview);
    } else {
      if (cachedThumb) {
        loadSrc(cachedThumb);
      } else {
        loadSrc(buildImageApiUrl("preview", path));
      }
      preloadPreview(path);
    }
  }

  previewPlaceholder.style.display = "none";
  renderPreviewInfo();
  renderPreviewCaptionOverlay();
  renderVideoEditPanel();
  renderMaskEditorUi();
  renderGifConvertButton();
  if (state.promptPreview.sourcePath === path && state.promptPreview.displayPath && state.promptPreview.displayPath !== path && state.previewMediaType === "image") {
    loadPromptPreviewImage(state.promptPreview.displayPath, { preserveView });
  }
}

function hidePreview() {
  if (state.maskEditor.active) {
    closeMaskEditor();
  }
  state.previewPath = null;
  state.previewMediaType = null;
  resetPromptPreviewState();
  previewImg.style.display = "none";
  stopPreviewVideo({ clearSource: true });
  previewPlaceholder.style.display = "flex";
  previewInfo.style.display = "none";
  imgNatW = 0;
  imgNatH = 0;
  renderPreviewCaptionOverlay();
  clearCropDraft();
  renderVideoEditPanel();
  renderGifConvertButton();
}

// ===== ZOOM & PAN =====
{
  const panel = previewStage;
  let fileDragDepth = 0;

  const eventHasFiles = (event) => {
    const types = event?.dataTransfer?.types;
    return Array.isArray(types) ? types.includes("Files") : !!types?.includes?.("Files");
  };

  window.addEventListener("dragover", (e) => {
    if (!eventHasFiles(e)) return;
    e.preventDefault();
  });

  window.addEventListener("drop", (e) => {
    if (!eventHasFiles(e)) return;
    e.preventDefault();
  });

  panel.addEventListener("dragenter", (e) => {
    if (!eventHasFiles(e)) return;
    e.preventDefault();
    fileDragDepth += 1;
    setFileDropActive(true);
  });

  panel.addEventListener("dragover", (e) => {
    if (!eventHasFiles(e)) return;
    e.preventDefault();
    if (e.dataTransfer) {
      e.dataTransfer.dropEffect = state.folder ? "copy" : "none";
    }
    setFileDropActive(true);
  });

  panel.addEventListener("dragleave", (e) => {
    if (!eventHasFiles(e)) return;
    e.preventDefault();
    fileDragDepth = Math.max(0, fileDragDepth - 1);
    if (fileDragDepth === 0) {
      setFileDropActive(false);
    }
  });

  panel.addEventListener("drop", async (e) => {
    if (!eventHasFiles(e)) return;
    e.preventDefault();
    fileDragDepth = 0;
    setFileDropActive(false);
    await uploadDroppedFiles(e.dataTransfer?.files);
  });

  panel.addEventListener("contextmenu", (e) => {
    if (canEditCrop() || isMaskEditorVisible()) e.preventDefault();
  });

  // Mouse wheel zoom — zooms toward cursor position
  panel.addEventListener("wheel", (e) => {
    if (!imgNatW) return;
    e.preventDefault();
    userHasZoomed = true;
    const rect = panel.getBoundingClientRect();
    const mx = e.clientX - rect.left;
    const my = e.clientY - rect.top;

    const oldZoom = zoomLevel;
    const factor = e.deltaY < 0 ? 1.15 : 1 / 1.15;
    zoomLevel *= factor;
    // Clamp zoom
    const panelW = panel.clientWidth;
    const panelH = panel.clientHeight;
    const minZoom = Math.min(panelW / imgNatW, panelH / imgNatH, 1) * 0.5;
    zoomLevel = Math.max(minZoom, Math.min(zoomLevel, 30));

    // Adjust pan so the point under the cursor stays fixed
    const realFactor = zoomLevel / oldZoom;
    panX = mx - realFactor * (mx - panX);
    panY = my - realFactor * (my - panY);
    applyTransform();
  }, { passive: false });

  // Drag to pan
  let isDragging = false;
  let dragMoved = false;
  let dragStartX, dragStartY, dragPanStartX, dragPanStartY;

  panel.addEventListener("mousedown", (e) => {
    if (isMaskSignalProbeMode() && e.button === 2) {
      if (e.target.closest("#video-edit-panel, #preview-caption-overlay, #mask-editor-panel, #preview-action-bar, #mask-action-bar, #mask-edit-btn, #image-edit-btn, #duplicate-image-btn, #mask-apply-btn, #mask-cancel-btn, #mask-undo-btn, #mask-redo-btn, #mask-view-mode-btn, #mask-latent-preview-btn, #mask-reset-btn")) {
        return;
      }
      if (!isClientInsidePreviewImage(e.clientX, e.clientY)) {
        return;
      }
      beginMaskSignalProbeDrag(e);
      e.preventDefault();
      return;
    }
    if (isMaskEditorVisible() && e.button === 2) {
      if (e.target.closest("#video-edit-panel, #preview-caption-overlay, #mask-editor-panel, #preview-action-bar, #mask-action-bar, #mask-edit-btn, #image-edit-btn, #duplicate-image-btn, #mask-apply-btn, #mask-cancel-btn, #mask-undo-btn, #mask-redo-btn, #mask-view-mode-btn, #mask-latent-preview-btn, #mask-reset-btn")) {
        return;
      }
      if (!isClientInsidePreviewImage(e.clientX, e.clientY)) {
        return;
      }
      beginMaskPaint(e);
      e.preventDefault();
      return;
    }
    if (e.button === 2 && canEditCrop()) {
      startCropCreate(e);
      updateCropGuideFromClient(e.clientX, e.clientY);
      e.preventDefault();
      return;
    }
    if (e.button !== 0) return;
    if (e.target.closest("#video-edit-panel, #preview-caption-overlay, #mask-editor-panel, #preview-action-bar, #mask-action-bar, #crop-apply-btn, #crop-cancel-btn, #crop-remove-btn, #mask-edit-btn, #image-edit-btn, #duplicate-image-btn, #mask-apply-btn, #mask-cancel-btn, #mask-undo-btn, #mask-redo-btn, #mask-view-mode-btn, #mask-latent-preview-btn, #mask-reset-btn, #rotate-controls")) {
      return;
    }
    isDragging = true;
    dragMoved = false;
    state.ui.suppressVideoClick = false;
    userHasZoomed = true;
    dragStartX = e.clientX;
    dragStartY = e.clientY;
    dragPanStartX = panX;
    dragPanStartY = panY;
    panel.classList.add("dragging");
    e.preventDefault();
  });

  panel.addEventListener("mousemove", (e) => {
    updateMaskCursor(e.clientX, e.clientY);
    updateCropGuideFromClient(e.clientX, e.clientY);
  });

  panel.addEventListener("mouseleave", () => {
    clearCropGuide();
    clearMaskCursor();
  });

  window.addEventListener("mousemove", (e) => {
    if (videoTimelineInteraction) {
      updateVideoTimelineInteraction(e.clientX);
      return;
    }
    updateMaskCursor(e.clientX, e.clientY);
    if (state.maskEditor.signalProbeDragging && isMaskSignalProbeMode()) {
      updateMaskSignalProbeDrag(e.clientX, e.clientY);
      return;
    }
    if (state.maskEditor.painting && isMaskEditorVisible()) {
      paintMaskAtClient(e.clientX, e.clientY);
      return;
    }
    if (state.cropInteraction && canEditCrop()) {
      updateCropGuideFromClient(e.clientX, e.clientY);
      const point = screenToImage(e.clientX, e.clientY);
      if (state.cropInteraction.mode === "create") {
        const result = chooseClosestRatio(
          state.cropInteraction.anchor.x,
          state.cropInteraction.anchor.y,
          point.x,
          point.y,
          state.cropAspectRatios,
          state.cropInteraction.stickySnap,
        );
        if (result?.crop) {
          state.cropInteraction.stickySnap = result.stickyChoice;
          setCropDraft(result.crop, true);
        }
      } else if (state.cropInteraction.mode === "resize") {
        const result = resizeCropFromHandle(
          state.cropInteraction.baseCrop,
          state.cropInteraction.handle,
          point.x,
          point.y,
          state.cropInteraction.stickySnap,
        );
        if (result?.crop) {
          state.cropInteraction.stickySnap = result.stickyChoice;
          setCropDraft(result.crop, true);
        }
      }
      return;
    }
    if (!isDragging) return;
    if (!dragMoved && (Math.abs(e.clientX - dragStartX) > 2 || Math.abs(e.clientY - dragStartY) > 2)) {
      dragMoved = true;
    }
    panX = dragPanStartX + (e.clientX - dragStartX);
    panY = dragPanStartY + (e.clientY - dragStartY);
    applyTransform();
  });

  window.addEventListener("mouseup", (e) => {
    if (videoTimelineInteraction) {
      finishVideoTimelineInteraction(e.clientX);
      return;
    }
    if (state.maskEditor.signalProbeDragging) {
      stopMaskSignalProbeDrag();
    }
    if (state.maskEditor.painting) {
      stopMaskPaint();
    }
    if (state.cropInteraction) {
      state.cropInteraction = null;
    }
    if (isDragging) {
      state.ui.suppressVideoClick = dragMoved;
      isDragging = false;
      panel.classList.remove("dragging");
    }
  });

  // Double-click to reset zoom
  panel.addEventListener("dblclick", (e) => {
    if (!imgNatW) return;
    e.preventDefault();
    userHasZoomed = false;
    resetZoomPan();
  });

  // Re-fit on window resize
  window.addEventListener("resize", () => {
    if (imgNatW && zoomLevel) {
      // Only reset if at fit-zoom (not manually zoomed)
      const panelW = panel.clientWidth;
      const panelH = panel.clientHeight;
      const fitScale = Math.min(panelW / imgNatW, panelH / imgNatH, 1);
      if (Math.abs(zoomLevel - fitScale) < 0.01 || zoomLevel < fitScale) {
        fitImageToPanel();
      }
    }
    if (state.previewMediaType === "video" && state.previewPath) {
      ensureVideoTimelineLoaded(state.previewPath)
        .then(() => renderVideoEditPanel())
        .catch((err) => showErrorToast(err.message || "Failed to resize video timeline"));
    }
  });
}

cropBox.querySelectorAll("[data-handle]").forEach(handleEl => {
  handleEl.addEventListener("mousedown", (e) => {
    if (e.button !== 0 || !canEditCrop()) return;
    startCropResize(handleEl.dataset.handle, e);
  });
});
maskEditBtn.addEventListener("click", enterMaskEditMode);
imageEditBtn.addEventListener("click", enterImageEditMode);
createPromptPreviewBtn.addEventListener("click", () => {
  runCreatePromptPreviewAction().catch(() => {});
});
promptPreviewBtn.addEventListener("click", () => {
  togglePromptPreviewButtonAction().catch(() => {});
});
duplicateImageBtn.addEventListener("click", duplicateCurrentImage);
videoMaskAddBtn.addEventListener("click", enterVideoMaskAddMode);
maskApplyBtn.addEventListener("click", () => saveMaskEdit().catch(() => {}));
maskCancelBtn.addEventListener("click", cancelMaskEdit);
maskUndoBtn.addEventListener("click", undoMaskEdit);
maskRedoBtn.addEventListener("click", redoMaskEdit);
maskViewModeBtn.addEventListener("click", toggleMaskEditorViewMode);
maskLatentPreviewBtn.addEventListener("click", toggleMaskLatentPreview);
maskSignalProbeBtn.addEventListener("click", toggleMaskSignalProbeMode);
maskResetBtn.addEventListener("click", resetMaskEditToDefault);
maskBrushSizeInput.addEventListener("input", (e) => {
  state.maskEditor.brushSizePercent = clamp(Number(e.target.value || 6), 0.2, 100);
  updateMaskControlLabels();
});
maskBrushValueInput.addEventListener("input", (e) => {
  state.maskEditor.brushValue = clamp(Number(e.target.value || 100), 0, 100);
  updateMaskControlLabels();
});
maskBrushColorInput.addEventListener("input", (e) => {
  state.maskEditor.brushColor = String(e.target.value || "#ff5a5a").trim() || "#ff5a5a";
  updateMaskControlLabels();
});
maskBrushCoreInput.addEventListener("input", (e) => {
  state.maskEditor.brushCore = clamp(Number(e.target.value || 30), 0, 95);
  updateMaskControlLabels();
});
maskBrushSteepnessInput.addEventListener("input", (e) => {
  state.maskEditor.brushSteepness = clamp(Number(e.target.value || 8), 1, 32);
  updateMaskControlLabels();
});
maskLatentBaseWidthInput.addEventListener("input", (e) => {
  const presets = getMaskLatentBaseWidthPresets();
  const index = clamp(Math.round(Number(e.target.value || 0)), 0, Math.max(0, presets.length - 1));
  state.maskEditor.latentBaseWidth = presets[index] || presets[0] || 512;
  updateMaskControlLabels();
  scheduleMaskLatentPreviewRender({ imageDirty: true });
});
maskLatentDividerInput.addEventListener("input", (e) => {
  state.maskEditor.latentDivider = Math.round(clamp(Number(e.target.value || 16), 1, 64));
  updateMaskControlLabels();
  scheduleMaskLatentPreviewRender();
});
maskLatentNoiseInput.addEventListener("input", (e) => {
  state.maskEditor.latentNoiseTimestep = Math.round(clamp(Number(e.target.value || 0), 0, MASK_LATENT_NOISE_MAX_TIMESTEP));
  updateMaskControlLabels();
  scheduleMaskLatentPreviewRender({ imageDirty: true });
});
cropApplyBtn.addEventListener("click", applyCropDraft);
cropCancelBtn.addEventListener("click", cancelCropEdit);
cropRemoveBtn.addEventListener("click", removeCrop);
rotateLeftBtn.addEventListener("click", () => rotatePreviewImage("left"));
rotateRightBtn.addEventListener("click", () => rotatePreviewImage("right"));
previewVideo.addEventListener("click", (e) => {
  if (state.previewMediaType !== "video" || !state.previewPath) return;
  if (state.ui.suppressVideoClick) {
    state.ui.suppressVideoClick = false;
    return;
  }
  e.preventDefault();
  togglePreviewVideoPlayback();
});
previewVideo.addEventListener("play", syncPreviewVideoPlaybackState);
previewVideo.addEventListener("pause", handlePreviewVideoPause);
previewVideo.addEventListener("timeupdate", handlePreviewVideoTimeUpdate);
previewVideo.addEventListener("seeked", handlePreviewVideoSeeked);
previewVideo.addEventListener("ended", handlePreviewVideoEnded);
previewVideo.addEventListener("loadedmetadata", syncPreviewVideoPlaybackState);
videoPlayToggleBtn.addEventListener("click", togglePreviewVideoPlayback);
videoMuteBtn.addEventListener("click", togglePreviewVideoMute);
videoVolumeSlider.addEventListener("input", (e) => setPreviewVideoVolumeFromSlider(e.target.value));
videoTimelineViewport.addEventListener("mousedown", (e) => {
  if (e.button !== 0 || state.previewMediaType !== "video" || !state.previewPath) return;
  if (e.target === videoTimelineStartHandle) {
    beginVideoTimelineInteraction("start", e);
    return;
  }
  if (e.target === videoTimelineEndHandle) {
    beginVideoTimelineInteraction("end", e);
    return;
  }
  if (e.target === videoTimelinePlayhead) {
    beginVideoTimelineInteraction("playhead", e);
    return;
  }
  beginVideoTimelineInteraction("pan", e);
});
videoTimelineViewport.addEventListener("wheel", (e) => {
  if (state.previewMediaType !== "video" || !state.previewPath) return;
  e.preventDefault();
  e.stopPropagation();
  const rect = videoTimelineViewport.getBoundingClientRect();
  const anchorRatio = rect.width > 0 ? clamp((e.clientX - rect.left) / rect.width, 0, 1) : 0.5;
  const currentZoom = getVideoTimelineUi(state.previewPath).zoom;
  const nextZoom = currentZoom * (e.deltaY < 0 ? 1.2 : 1 / 1.2);
  setVideoTimelineZoom(state.previewPath, nextZoom, anchorRatio);
  renderVideoTimelineStrip(state.previewPath);
  syncVideoTimelineState(state.previewPath);
  syncPreviewVideoPlaybackState();
  ensureVideoTimelineLoaded(state.previewPath)
    .then(() => renderVideoEditPanel())
    .catch((err) => showErrorToast(err.message || "Failed to zoom video timeline"));
}, { passive: false });

// ===== CAPTION DATA =====
async function loadCaptionData(path) {
  const allSentences = getAllConfiguredSentences();
  const sentencesJson = JSON.stringify(allSentences);
  try {
    const resp = await fetch(`/api/caption?path=${encodeURIComponent(path)}&captions=${encodeURIComponent(sentencesJson)}`);
    if (resp.ok) {
      const data = await resp.json();
      state.captionCache[path] = normalizeCaptionCacheEntry(data);
      if (state.activeSentenceFilters.size > 0) {
        state.filterCaptionCacheKey = getActiveSentenceFilterKey();
      }
      // Update UI if still selected
      if (state.selectedPaths.size === 1 && state.selectedPaths.has(path)) {
        freeText.value = state.captionCache[path].free_text || "";
        scheduleUiRender({ sentences: true });
      }
      if (state.previewPath === path) {
        scheduleUiRender({ preview: true });
      }
      refreshGridForActiveFilters();
    }
  } catch (err) {
    console.error("Failed to load caption:", err);
  }
}

async function loadMetadataData(path) {
  try {
    const resp = await fetch(`/api/media/meta?path=${encodeURIComponent(path)}`);
    if (resp.ok) {
      const data = await resp.json();
      state.metadataCache[path] = normalizeMetadataCacheEntry(data);
      if (state.selectedPaths.size === 1 && state.selectedPaths.has(path)) {
        renderMetadataEditor();
      }
    }
  } catch (err) {
    console.error("Failed to load metadata:", err);
  }
}

async function loadMultiCaptionState() {
  const paths = [...state.selectedPaths];

  // Batch load captions
  try {
    const data = await fetchCaptionsBulk(paths);
    for (const [path, caption] of Object.entries(data)) {
      state.captionCache[path] = normalizeCaptionCacheEntry(caption);
    }
    if (state.activeSentenceFilters.size > 0) {
      state.filterCaptionCacheKey = getActiveSentenceFilterKey();
    }
    refreshGridForActiveFilters();
    scheduleUiRender({ sentences: true, preview: true });
  } catch (err) {
    console.error("Failed to load bulk captions:", err);
  }
}

async function loadMultiMetadataState() {
  const paths = [...state.selectedPaths];

  try {
    const data = await fetchMetadataBulk(paths);
    for (const [path, metadata] of Object.entries(data)) {
      state.metadataCache[path] = normalizeMetadataCacheEntry(metadata);
    }
    renderMetadataEditor();
  } catch (err) {
    console.error("Failed to load bulk metadata:", err);
  }
}

function updateMultiInfo() {
  if (state.selectedPaths.size > 1) {
    multiInfo.style.display = "block";
    if (state.ui.activeRightPanelTab === "metadata") {
      multiInfo.textContent = `${state.selectedPaths.size} media files selected — Apply updates the filled metadata fields on all selected files`;
    } else {
      multiInfo.textContent = `${state.selectedPaths.size} media files selected — toggling captions applies to all`;
    }
  } else {
    multiInfo.style.display = "none";
  }
}

// ===== SENTENCES =====
let _saveSentencesTimeout = null;
function saveSectionsToStorage() {
  // Debounce saves to avoid hammering the server during rapid edits
  if (_saveSentencesTimeout) clearTimeout(_saveSentencesTimeout);
  _saveSentencesTimeout = setTimeout(() => {
    if (!state.folder) return;
    fetch("/api/settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sections: serializeSectionsForSave(), folder: state.folder }),
    }).catch(err => console.error("Failed to save sections:", err));
  }, 300);
}

function applyRemovedSentencesToLocalState(removedSentences = []) {
  const removed = new Set((removedSentences || []).filter(Boolean));
  if (removed.size === 0) return;
  for (const sentence of [...state.activeSentenceFilters.keys()]) {
    if (removed.has(sentence)) {
      state.activeSentenceFilters.delete(sentence);
    }
  }
  for (const caption of Object.values(state.captionCache)) {
    if (!caption?.enabled_sentences) continue;
    caption.enabled_sentences = caption.enabled_sentences.filter(sentence => !removed.has(sentence));
    if (typeof caption.free_text === "string") {
      caption.free_text = caption.free_text
        .split(/\r?\n/)
        .filter(line => !removed.has(line.trim()))
        .join("\n");
    }
  }
  state.filterCaptionCacheKey = getActiveSentenceFilterKey();
}

async function refreshCaptionsAfterSchemaChange() {
  refreshGridForActiveFilters();
  if (state.selectedPaths.size === 1) {
    const path = [...state.selectedPaths][0];
    await loadCaptionData(path);
    return;
  }
  if (state.selectedPaths.size > 1) {
    await loadMultiCaptionState();
    return;
  }
  renderSentences();
}

const SECTION_DRAG_TYPE = "application/x-tag2-section";
const GROUP_DRAG_TYPE = "application/x-tag2-group";
const SENTENCE_DRAG_TYPE = "application/x-tag2-sentence";
const SECTION_ITEM_DRAG_TYPE = "application/x-tag2-section-item";

function setDragPayload(event, type, payload) {
  if (!event.dataTransfer) return;
  event.dataTransfer.setData(type, JSON.stringify(payload));
  event.dataTransfer.effectAllowed = "move";
}

function getDragPayload(event, type) {
  const raw = event.dataTransfer?.getData(type);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function normalizeGroupIndex(groupIdx) {
  return Number.isInteger(groupIdx) && groupIdx >= 0 ? groupIdx : null;
}

function getSentenceContainer(secIdx, groupIdx = null) {
  const section = state.sections[secIdx];
  if (!section) return null;
  if (groupIdx === null) return section.sentences;
  return section.groups?.[groupIdx]?.sentences || null;
}

function reorderList(list, fromIdx, toIdx) {
  if (!Array.isArray(list)) return false;
  if (fromIdx < 0 || toIdx < 0 || fromIdx >= list.length || toIdx >= list.length || fromIdx === toIdx) {
    return false;
  }
  const [moved] = list.splice(fromIdx, 1);
  list.splice(toIdx, 0, moved);
  return true;
}

function moveSentenceWithinContainer(secIdx, groupIdx, fromIdx, toIdx) {
  const sentences = getSentenceContainer(secIdx, groupIdx);
  if (!reorderList(sentences, fromIdx, toIdx)) return;
  saveSectionsToStorage();
  renderSentences();
}

function createTopLevelSentenceDragRef(sentence) {
  return { type: "sentence", sentence };
}

function createTopLevelGroupDragRef(groupId) {
  return { type: "group", group_id: groupId };
}

function isValidTopLevelDragRef(ref) {
  return !!ref && typeof ref === "object" && (
    (ref.type === "sentence" && !!ref.sentence) ||
    (ref.type === "group" && !!ref.group_id)
  );
}

function moveTopLevelSectionItem(secIdx, fromRef, toRef) {
  const section = state.sections[secIdx];
  if (!section || !isValidTopLevelDragRef(fromRef) || !isValidTopLevelDragRef(toRef)) return;
  const nextOrder = [...(section.item_order || [])];
  const fromKey = getSectionOrderItemKey(fromRef);
  const toKey = getSectionOrderItemKey(toRef);
  const fromIdx = nextOrder.findIndex(item => getSectionOrderItemKey(item) === fromKey);
  const toIdx = nextOrder.findIndex(item => getSectionOrderItemKey(item) === toKey);
  if (fromIdx < 0 || toIdx < 0 || fromIdx === toIdx) return;
  const [moved] = nextOrder.splice(fromIdx, 1);
  nextOrder.splice(toIdx, 0, moved);
  section.item_order = nextOrder;
  saveSectionsToStorage();
  renderSentences();
}

function moveTopLevelSectionItemToEnd(secIdx, fromRef) {
  const section = state.sections[secIdx];
  if (!section || !isValidTopLevelDragRef(fromRef)) return;
  const nextOrder = [...(section.item_order || [])];
  const fromKey = getSectionOrderItemKey(fromRef);
  const fromIdx = nextOrder.findIndex(item => getSectionOrderItemKey(item) === fromKey);
  if (fromIdx < 0 || fromIdx === nextOrder.length - 1) return;
  const [moved] = nextOrder.splice(fromIdx, 1);
  nextOrder.push(moved);
  section.item_order = nextOrder;
  saveSectionsToStorage();
  renderSentences();
}

function moveTopLevelSentenceIntoGroup(secIdx, sentence, groupIdx, targetIdx = null) {
  const section = state.sections[secIdx];
  const group = section?.groups?.[groupIdx];
  if (!section || !group || !sentence) return;
  const sourceIdx = (section.sentences || []).indexOf(sentence);
  if (sourceIdx < 0) return;

  section.sentences.splice(sourceIdx, 1);
  section.item_order = (section.item_order || []).filter(item => !(item?.type === "sentence" && item.sentence === sentence));

  const insertAt = Number.isInteger(targetIdx)
    ? Math.max(0, Math.min(targetIdx, group.sentences.length))
    : group.sentences.length;
  group.sentences.splice(insertAt, 0, sentence);

  saveSectionsToStorage();
  renderSentences();
}

function addSentenceToSection(secIdx, text, groupIdx = null) {
  if (!text) return;
  const allSentences = getAllConfiguredSentences();
  if (allSentences.includes(text)) return;
  if (secIdx >= 0 && secIdx < state.sections.length) {
    if (groupIdx === null) {
      state.sections[secIdx].sentences.push(text);
      state.sections[secIdx].item_order = [
        ...(state.sections[secIdx].item_order || []),
        createSentenceOrderItem(text),
      ];
    } else if (groupIdx >= 0 && groupIdx < (state.sections[secIdx].groups || []).length) {
      state.sections[secIdx].groups[groupIdx].sentences.push(text);
    }
  }
  saveSectionsToStorage();
  renderSentences();
}

async function removeSentence(sentence) {
  if (!state.folder) {
    statusBar.textContent = "No folder loaded";
    return;
  }

  statusBar.textContent = "Deleting caption...";
  try {
    const resp = await fetch("/api/caption/delete-preset", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        folder: state.folder,
        caption: sentence,
      }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to delete caption");
    }
    state.sections = normalizeSectionsData(data.sections || state.sections);
    applyRemovedSentencesToLocalState(data.removed_captions || data.removed_sentences || [sentence]);
    await refreshCaptionsAfterSchemaChange();
    statusBar.textContent = "Caption deleted";
  } catch (err) {
    statusBar.textContent = `Delete caption error: ${err.message}`;
    renderSentences();
  }
}

function toggleSentenceHiddenOnExport(sentence) {
  const group = findGroupForSentence(sentence);
  if (!group) return;
  const hidden = new Set(group.hidden_sentences || []);
  if (hidden.has(sentence)) {
    hidden.delete(sentence);
  } else {
    hidden.add(sentence);
  }
  group.hidden_sentences = (group.sentences || []).filter(item => hidden.has(item));
  saveSectionsToStorage();
  renderSentences();
}

async function renameSentence(oldSentence, newSentence) {
  const nextSentence = String(newSentence || "").trim();
  if (!nextSentence || nextSentence === oldSentence) {
    renderSentences();
    return;
  }
  if (getAllConfiguredSentences().includes(nextSentence)) {
    statusBar.textContent = "Caption already exists";
    renderSentences();
    return;
  }
  if (!state.folder) {
    statusBar.textContent = "No folder loaded";
    renderSentences();
    return;
  }

  statusBar.textContent = "Renaming caption...";
  try {
    const resp = await fetch("/api/caption/rename-preset", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        folder: state.folder,
        old_caption: oldSentence,
        new_caption: nextSentence,
      }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to rename caption");
    }
    state.sections = normalizeSectionsData(data.sections || state.sections);
    for (const caption of Object.values(state.captionCache)) {
      if (!caption?.enabled_sentences) continue;
      caption.enabled_sentences = [...new Set(caption.enabled_sentences.map(sentence => sentence === oldSentence ? nextSentence : sentence))];
    }
    if (state.activeSentenceFilters.has(oldSentence)) {
      const mode = state.activeSentenceFilters.get(oldSentence);
      state.activeSentenceFilters.delete(oldSentence);
      if (mode === "has" || mode === "missing") {
        state.activeSentenceFilters.set(nextSentence, mode);
      }
    }
    state.filterCaptionCacheKey = getActiveSentenceFilterKey();
    refreshGridForActiveFilters();
    if (state.selectedPaths.size === 1) {
      const path = [...state.selectedPaths][0];
      await loadCaptionData(path);
    } else if (state.selectedPaths.size > 1) {
      await loadMultiCaptionState();
    } else {
      renderSentences();
    }
    statusBar.textContent = "Caption renamed";
  } catch (err) {
    statusBar.textContent = `Rename error: ${err.message}`;
    renderSentences();
  }
}

async function renameSection(oldName, newName) {
  const nextName = String(newName || "").trim();
  const currentName = String(oldName || "").trim();
  if (nextName === currentName) {
    renderSentences();
    return;
  }
  if (!state.folder) {
    statusBar.textContent = "No folder loaded";
    renderSentences();
    return;
  }

  statusBar.textContent = "Renaming section...";
  try {
    const resp = await fetch("/api/section/rename", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        folder: state.folder,
        old_name: currentName,
        new_name: nextName,
      }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to rename section");
    }
    state.sections = normalizeSectionsData(data.sections || state.sections);
    renderSentences();
    statusBar.textContent = "Section renamed";
  } catch (err) {
    statusBar.textContent = `Section rename error: ${err.message}`;
    renderSentences();
  }
}

function addSection() {
  const name = prompt("Section name:");
  addSectionWithName(name);
}

function addSectionWithName(name) {
  if (!name || !name.trim()) return;
  state.sections.push(createEmptySection(name.trim()));
  saveSectionsToStorage();
  renderSentences();
}

function addGroup(secIdx) {
  const name = prompt("Group name:");
  addGroupWithName(secIdx, name);
}

function addGroupWithName(secIdx, name) {
  if (!name || !name.trim()) return;
  if (secIdx < 0 || secIdx >= state.sections.length) return;
  const group = createEmptyGroup(name.trim());
  state.sections[secIdx].groups.push(group);
  state.sections[secIdx].item_order = [
    ...(state.sections[secIdx].item_order || []),
    createGroupOrderItem(group.id),
  ];
  saveSectionsToStorage();
  renderSentences();
}

async function deleteSection(index) {
  const sec = state.sections[index];
  const label = sec.name || "(General)";
  if (!confirm(`Delete section "${label}" and its ${countSectionSentences(sec)} caption(s)?`)) return;
  if (!state.folder) {
    statusBar.textContent = "No folder loaded";
    return;
  }

  statusBar.textContent = "Deleting section...";
  try {
    const resp = await fetch("/api/section/delete", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        folder: state.folder,
        section_index: index,
      }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to delete section");
    }
    state.sections = normalizeSectionsData(data.sections || state.sections);
    applyRemovedSentencesToLocalState(data.removed_sentences || []);
    await refreshCaptionsAfterSchemaChange();
    statusBar.textContent = "Section deleted";
  } catch (err) {
    statusBar.textContent = `Delete section error: ${err.message}`;
    renderSentences();
  }
}

async function deleteGroup(secIdx, groupIdx) {
  const group = state.sections[secIdx]?.groups?.[groupIdx];
  if (!group) return;
  const label = group.name || "(Group)";
  if (!confirm(`Delete group "${label}" and its ${(group.sentences || []).length} caption(s)?`)) return;
  if (!state.folder) {
    statusBar.textContent = "No folder loaded";
    return;
  }

  statusBar.textContent = "Deleting group...";
  try {
    const resp = await fetch("/api/group/delete", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        folder: state.folder,
        section_index: secIdx,
        group_index: groupIdx,
      }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to delete group");
    }
    state.sections = normalizeSectionsData(data.sections || state.sections);
    applyRemovedSentencesToLocalState(data.removed_sentences || []);
    await refreshCaptionsAfterSchemaChange();
    statusBar.textContent = "Group deleted";
  } catch (err) {
    statusBar.textContent = `Delete group error: ${err.message}`;
    renderSentences();
  }
}

function beginEditableRename(targetEl, currentValue, onCommit, highlightColor = "var(--accent)") {
  if (targetEl.contentEditable === "true") {
    return;
  }
  const draggableContainer = targetEl.closest(".section-header");
  const restoreDraggable = draggableContainer ? draggableContainer.draggable : null;
  const handleDragStart = (e) => e.preventDefault();

  state.ui.activeInlineEditor = targetEl;
  targetEl.contentEditable = "true";
  targetEl.style.borderColor = highlightColor;
  targetEl.draggable = false;
  targetEl.addEventListener("dragstart", handleDragStart);
  if (draggableContainer) {
    draggableContainer.draggable = false;
  }
  targetEl.focus();
  const range = document.createRange();
  range.selectNodeContents(targetEl);
  range.collapse(false);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);

  const handleKeyDown = (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      targetEl.blur();
      return;
    }
    if (e.key === "Escape") {
      e.preventDefault();
      targetEl.textContent = currentValue || "";
      targetEl.blur();
    }
  };

  const handleBlur = () => {
    targetEl.contentEditable = "false";
    targetEl.style.borderColor = "";
    if (state.ui.activeInlineEditor === targetEl) {
      state.ui.activeInlineEditor = null;
    }
    targetEl.removeEventListener("dragstart", handleDragStart);
    if (draggableContainer && restoreDraggable !== null) {
      draggableContainer.draggable = restoreDraggable;
    }
    const nextValue = targetEl.textContent.trim();
    if (nextValue !== currentValue) {
      onCommit(nextValue);
    } else {
      targetEl.textContent = currentValue || targetEl.textContent;
    }
    targetEl.removeEventListener("keydown", handleKeyDown);
    targetEl.removeEventListener("blur", handleBlur);
    flushQueuedUiRenders({ force: true });
  };

  targetEl.addEventListener("blur", handleBlur);
  targetEl.addEventListener("keydown", handleKeyDown);
}

function getSentenceSelectionState(sentence, selectedPaths) {
  let enabledCount = 0;
  const totalCount = selectedPaths.length;
  for (const path of selectedPaths) {
    const cap = state.captionCache[path];
    if (cap?.enabled_sentences?.includes(sentence)) {
      enabledCount++;
    }
  }
  return {
    enabledCount,
    totalCount,
    isChecked: totalCount > 0 && enabledCount === totalCount,
    isPartial: totalCount > 0 && enabledCount > 0 && enabledCount < totalCount,
  };
}

function createSentenceListItem(sentence, selectedPaths, options = {}) {
  const { isExclusive = false, allowSuppressToggle = false } = options;
  const secIdx = Number.isInteger(options.secIdx) ? options.secIdx : -1;
  const groupIdx = normalizeGroupIndex(options.groupIdx);
  const sentenceIdx = Number.isInteger(options.sentenceIdx) ? options.sentenceIdx : -1;
  const topLevelMix = !!options.topLevelMix;
  const canRefreshSentence = groupIdx === null;
  const isGroupSentence = groupIdx !== null;
  const filterMode = getSentenceFilterMode(sentence);
  const isFilterActive = filterMode !== "off";
  const { isChecked, isPartial } = getSentenceSelectionState(sentence, selectedPaths);
  const li = document.createElement("li");
  li.dataset.sentence = sentence;
  sentenceListElements.set(sentence, li);
  const topLevelRef = topLevelMix ? createTopLevelSentenceDragRef(sentence) : null;

  const dragHandle = document.createElement("span");
  dragHandle.className = "drag-handle";
  dragHandle.title = "Drag to reorder caption";
  dragHandle.draggable = true;
  dragHandle.addEventListener("click", (e) => e.stopPropagation());
  dragHandle.addEventListener("mousedown", (e) => e.stopPropagation());
  dragHandle.addEventListener("dragstart", (e) => {
    e.stopPropagation();
    if (topLevelMix) {
      setDragPayload(e, SECTION_ITEM_DRAG_TYPE, { secIdx, item: topLevelRef });
    } else {
      setDragPayload(e, SENTENCE_DRAG_TYPE, {
        secIdx,
        groupIdx: groupIdx ?? -1,
        sentenceIdx,
      });
    }
    li.classList.add("dragging-item");
  });
  dragHandle.addEventListener("dragend", () => {
    li.classList.remove("dragging-item");
    document.querySelectorAll(".section-sentences li.drag-over").forEach(item => item.classList.remove("drag-over"));
    document.querySelectorAll(".section-sentences.drag-over-end").forEach(list => list.classList.remove("drag-over-end"));
  });

  const checkBox = document.createElement("div");
  checkBox.className = `check-box${isExclusive ? " radio" : ""}`;
  if (isChecked) checkBox.classList.add("checked");
  else if (isPartial) checkBox.classList.add("partial");

  const textSpan = document.createElement("span");
  textSpan.className = "sentence-text editable";
  if (isPartial) textSpan.classList.add("partial");
  if (allowSuppressToggle && isSentenceHiddenOnExport(sentence)) textSpan.classList.add("hidden-export");
  textSpan.textContent = sentence;
  textSpan.title = "Click to rename caption";
  textSpan.addEventListener("click", (e) => {
    e.stopPropagation();
    beginEditableRename(textSpan, sentence, (newName) => {
      renameSentence(sentence, newName);
    });
  });

  const rmBtn = document.createElement("button");
  rmBtn.className = "remove-btn";
  rmBtn.textContent = "\u00D7";
  rmBtn.title = "Remove sentence";
  rmBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    removeSentence(sentence);
  });

  let refreshBtn = null;
  if (canRefreshSentence) {
    refreshBtn = document.createElement("button");
    refreshBtn.type = "button";
    refreshBtn.className = "sentence-refresh-btn generate-btn generate-btn-icon";
    refreshBtn.title = "Refresh this caption with auto captioning";
    refreshBtn.disabled = state.autoCaptioning || state.selectedPaths.size === 0 || !state.ollamaModel.trim();
    setGenerateButtonContent(refreshBtn, "Refresh this caption with auto captioning", { iconOnly: true });
    refreshBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      refreshSentenceCaption(sentence);
    });
  }

  const filterToggleBtn = document.createElement("button");
  filterToggleBtn.type = "button";
  filterToggleBtn.className = `filter-toggle-btn${filterMode === "has" ? " active" : ""}${filterMode === "missing" ? " missing" : ""}`;
  if (isGroupSentence) {
    filterToggleBtn.title = isFilterActive
      ? "Remove this group caption from the thumbnail filter"
      : "Filter thumbnails to images that include this group caption";
    filterToggleBtn.setAttribute("aria-label", isFilterActive
      ? "Remove this group caption from the thumbnail filter"
      : "Filter thumbnails to images that include this group caption");
  } else {
    const spokenMode = filterMode === "missing" ? "Doesn't Have" : filterMode === "has" ? "Has" : "Off";
    const buttonTitle = filterMode === "has"
      ? "Filter thumbnails to images that include this caption"
      : filterMode === "missing"
        ? "Filter thumbnails to images that do not include this caption"
        : "Caption filter is off";
    filterToggleBtn.title = `${buttonTitle}. Click to cycle Off, Has, Doesn't Have.`;
    filterToggleBtn.setAttribute("aria-label", `Caption filter state: ${spokenMode}`);
  }
  filterToggleBtn.setAttribute("aria-pressed", isFilterActive ? "true" : "false");
  filterToggleBtn.appendChild(createFilterIcon());
  filterToggleBtn.addEventListener("click", async (e) => {
    e.stopPropagation();
    await toggleSentenceFilter(sentence);
  });

  if (allowSuppressToggle) {
    const hiddenOnExport = isSentenceHiddenOnExport(sentence);
    const exportToggleBtn = document.createElement("button");
    exportToggleBtn.type = "button";
    exportToggleBtn.className = `export-toggle-btn${hiddenOnExport ? "" : " active"}`;
    exportToggleBtn.textContent = "TXT";
    exportToggleBtn.title = hiddenOnExport
      ? "Do not write this option to the txt output when selected"
      : "Write this option to the txt output when selected";
    exportToggleBtn.setAttribute("aria-pressed", hiddenOnExport ? "false" : "true");
    exportToggleBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      toggleSentenceHiddenOnExport(sentence);
    });
    li.appendChild(dragHandle);
    li.appendChild(checkBox);
    li.appendChild(textSpan);
    if (refreshBtn) li.appendChild(refreshBtn);
    li.appendChild(filterToggleBtn);
    li.appendChild(exportToggleBtn);
    li.appendChild(rmBtn);
  } else {
    li.appendChild(dragHandle);
    li.appendChild(checkBox);
    li.appendChild(textSpan);
    if (refreshBtn) li.appendChild(refreshBtn);
    li.appendChild(filterToggleBtn);
    li.appendChild(rmBtn);
  }
  li.addEventListener("dragover", (e) => {
    if (topLevelMix) {
      const payload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
      if (!payload || payload.secIdx !== secIdx || !isValidTopLevelDragRef(payload.item)) return;
      if (getSectionOrderItemKey(payload.item) === getSectionOrderItemKey(topLevelRef)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      li.classList.add("drag-over");
      return;
    }
    const topLevelPayload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (
      topLevelPayload
      && topLevelPayload.secIdx === secIdx
      && topLevelPayload.item?.type === "sentence"
      && groupIdx !== null
    ) {
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      return;
    }
    const payload = getDragPayload(e, SENTENCE_DRAG_TYPE);
    if (!payload) return;
    if (payload.secIdx !== secIdx || normalizeGroupIndex(payload.groupIdx) !== groupIdx) return;
    if (payload.sentenceIdx === sentenceIdx) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    li.classList.add("drag-over");
  });
  li.addEventListener("dragleave", () => {
    li.classList.remove("drag-over");
  });
  li.addEventListener("drop", (e) => {
    if (topLevelMix) {
      const payload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
      if (!payload || payload.secIdx !== secIdx || !isValidTopLevelDragRef(payload.item)) return;
      e.preventDefault();
      li.classList.remove("drag-over");
      moveTopLevelSectionItem(secIdx, payload.item, topLevelRef);
      return;
    }
    const topLevelPayload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (
      topLevelPayload
      && topLevelPayload.secIdx === secIdx
      && topLevelPayload.item?.type === "sentence"
      && groupIdx !== null
    ) {
      e.preventDefault();
      li.classList.remove("drag-over");
      return;
    }
    const payload = getDragPayload(e, SENTENCE_DRAG_TYPE);
    if (!payload) return;
    const payloadGroupIdx = normalizeGroupIndex(payload.groupIdx);
    if (payload.secIdx !== secIdx || payloadGroupIdx !== groupIdx) return;
    e.preventDefault();
    li.classList.remove("drag-over");
    moveSentenceWithinContainer(secIdx, groupIdx, payload.sentenceIdx, sentenceIdx);
  });
  li.addEventListener("click", (e) => {
    e.stopPropagation();
    toggleSentence(sentence, isChecked, isPartial);
  });
  return li;
}

function createGroupListItem(section, group, secIdx, groupIdx, selectedPaths) {
  const wrapper = document.createElement("li");
  wrapper.className = "group-item-wrap";
  const topLevelRef = createTopLevelGroupDragRef(group.id);

  const groupBlock = document.createElement("div");
  groupBlock.className = "group-block";

  const groupHeader = document.createElement("div");
  groupHeader.className = "group-header";

  const groupDragHandle = document.createElement("span");
  groupDragHandle.className = "drag-handle";
  groupDragHandle.title = "Drag to reorder group";
  groupDragHandle.draggable = true;
  groupDragHandle.addEventListener("click", (e) => e.stopPropagation());
  groupDragHandle.addEventListener("mousedown", (e) => e.stopPropagation());
  groupDragHandle.addEventListener("dragstart", (e) => {
    e.stopPropagation();
    setDragPayload(e, SECTION_ITEM_DRAG_TYPE, { secIdx, item: topLevelRef });
    wrapper.classList.add("dragging-item");
  });
  groupDragHandle.addEventListener("dragend", () => {
    wrapper.classList.remove("dragging-item");
    document.querySelectorAll(".section-sentences li.drag-over").forEach(item => item.classList.remove("drag-over"));
    document.querySelectorAll(".section-sentences.drag-over-end").forEach(list => list.classList.remove("drag-over-end"));
  });
  groupHeader.appendChild(groupDragHandle);

  const groupCollapseBtn = document.createElement("button");
  groupCollapseBtn.type = "button";
  groupCollapseBtn.className = "collapse-btn";
  groupCollapseBtn.textContent = isGroupCollapsed(group) ? "▸" : "▾";
  groupCollapseBtn.title = isGroupCollapsed(group) ? "Expand group" : "Collapse group";
  groupCollapseBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    setGroupCollapsed(group, !isGroupCollapsed(group));
    renderSentences();
  });
  groupHeader.appendChild(groupCollapseBtn);

  const groupName = document.createElement("span");
  groupName.className = "group-name";
  groupName.draggable = false;
  groupName.textContent = group.name || "(Group)";
  groupName.addEventListener("click", (e) => {
    e.stopPropagation();
    beginEditableRename(groupName, group.name || "", (newName) => {
      group.name = newName;
      saveSectionsToStorage();
      renderSentences();
    }, "var(--accent-warn)");
  });
  groupHeader.appendChild(groupName);

  const groupRefreshBtn = document.createElement("button");
  groupRefreshBtn.type = "button";
  groupRefreshBtn.className = "group-refresh-btn generate-btn generate-btn-icon";
  groupRefreshBtn.title = "Refresh this group with auto captioning";
  setGenerateButtonContent(groupRefreshBtn, "Refresh this group with auto captioning", { iconOnly: true });
  groupRefreshBtn.disabled = state.autoCaptioning || state.selectedPaths.size === 0 || !state.ollamaModel.trim();
  groupRefreshBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    refreshGroupCaptions(secIdx, groupIdx);
  });
  groupHeader.appendChild(groupRefreshBtn);

  const groupRemoveBtn = document.createElement("button");
  groupRemoveBtn.type = "button";
  groupRemoveBtn.className = "group-remove-btn";
  groupRemoveBtn.textContent = "\u00D7";
  groupRemoveBtn.title = "Delete group";
  groupRemoveBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    deleteGroup(secIdx, groupIdx);
  });
  groupHeader.appendChild(groupRemoveBtn);

  groupHeader.addEventListener("dragover", (e) => {
    const payload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (!payload || payload.secIdx !== secIdx || payload.item?.type !== "sentence") return;
    e.preventDefault();
    e.stopPropagation();
    e.dataTransfer.dropEffect = "move";
    groupHeader.classList.add("drag-over");
  });
  groupHeader.addEventListener("dragleave", () => {
    groupHeader.classList.remove("drag-over");
  });
  groupHeader.addEventListener("drop", (e) => {
    const payload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (!payload || payload.secIdx !== secIdx || payload.item?.type !== "sentence") return;
    e.preventDefault();
    e.stopPropagation();
    groupHeader.classList.remove("drag-over");
    moveTopLevelSentenceIntoGroup(secIdx, payload.item.sentence, groupIdx);
  });

  groupBlock.appendChild(groupHeader);

  const groupBody = document.createElement("div");
  groupBody.className = `group-body${isGroupCollapsed(group) ? " collapsed" : ""}`;
  const groupSentences = document.createElement("ul");
  groupSentences.className = "section-sentences group-sentences";
  groupSentences.addEventListener("dragover", (e) => {
    const topLevelPayload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (topLevelPayload && topLevelPayload.secIdx === secIdx && topLevelPayload.item?.type === "sentence") {
      e.preventDefault();
      e.stopPropagation();
      e.dataTransfer.dropEffect = "move";
      if (e.target === groupSentences) {
        groupSentences.classList.add("drag-over-end");
      }
      return;
    }
    const payload = getDragPayload(e, SENTENCE_DRAG_TYPE);
    if (!payload) return;
    if (payload.secIdx !== secIdx || normalizeGroupIndex(payload.groupIdx) !== groupIdx) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    if (e.target === groupSentences) {
      groupSentences.classList.add("drag-over-end");
    }
  });
  groupSentences.addEventListener("dragleave", (e) => {
    if (!groupSentences.contains(e.relatedTarget)) {
      groupSentences.classList.remove("drag-over-end");
    }
  });
  groupSentences.addEventListener("drop", (e) => {
    const topLevelPayload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (topLevelPayload && topLevelPayload.secIdx === secIdx && topLevelPayload.item?.type === "sentence") {
      if (e.target !== groupSentences) return;
      e.preventDefault();
      e.stopPropagation();
      groupSentences.classList.remove("drag-over-end");
      moveTopLevelSentenceIntoGroup(secIdx, topLevelPayload.item.sentence, groupIdx);
      return;
    }
    const payload = getDragPayload(e, SENTENCE_DRAG_TYPE);
    if (!payload) return;
    if (payload.secIdx !== secIdx || normalizeGroupIndex(payload.groupIdx) !== groupIdx) return;
    if (e.target !== groupSentences) return;
    e.preventDefault();
    groupSentences.classList.remove("drag-over-end");
    const sentences = getSentenceContainer(secIdx, groupIdx) || [];
    const targetIdx = Math.max(0, sentences.length - 1);
    moveSentenceWithinContainer(secIdx, groupIdx, payload.sentenceIdx, targetIdx);
  });
  for (const [sentenceIdx, sentence] of (group.sentences || []).entries()) {
    groupSentences.appendChild(createSentenceListItem(sentence, selectedPaths, {
      isExclusive: true,
      allowSuppressToggle: true,
      secIdx,
      groupIdx,
      sentenceIdx,
    }));
  }
  groupBody.appendChild(groupSentences);
  groupBody.appendChild(createAddSentenceRow(secIdx, groupIdx, "Add caption to group..."));
  groupBlock.appendChild(groupBody);
  wrapper.appendChild(groupBlock);

  wrapper.addEventListener("dragover", (e) => {
    const payload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (!payload || payload.secIdx !== secIdx || !isValidTopLevelDragRef(payload.item)) return;
    if (getSectionOrderItemKey(payload.item) === getSectionOrderItemKey(topLevelRef)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    wrapper.classList.add("drag-over");
  });
  wrapper.addEventListener("dragleave", () => {
    wrapper.classList.remove("drag-over");
  });
  wrapper.addEventListener("drop", (e) => {
    const payload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (!payload || payload.secIdx !== secIdx || !isValidTopLevelDragRef(payload.item)) return;
    e.preventDefault();
    wrapper.classList.remove("drag-over");
    moveTopLevelSectionItem(secIdx, payload.item, topLevelRef);
  });

  return wrapper;
}

function createAddSentenceRow(secIdx, groupIdx = null, placeholder = "Add caption...") {
  if (state.hideAddButtons) return document.createComment("add-row-hidden");
  const addRow = document.createElement("div");
  addRow.className = `section-add-row${groupIdx === null ? "" : " group-add-row"}`;
  const addInput = document.createElement("input");
  addInput.type = "text";
  addInput.placeholder = placeholder;
  addInput.spellcheck = false;
  const addBtn = document.createElement("button");
  addBtn.type = "button";
  addBtn.appendChild(createAddPlusIcon());
  const submit = () => {
    const text = addInput.value.trim();
    if (!text) return;
    addSentenceToSection(secIdx, text, groupIdx);
    addInput.value = "";
  };
  addBtn.addEventListener("click", submit);
  addInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") submit();
  });
  addRow.appendChild(addInput);
  addRow.appendChild(addBtn);
  if (groupIdx === null) {
    const addGroupBtn = document.createElement("button");
    addGroupBtn.type = "button";
    addGroupBtn.className = "section-add-group-btn";
    const addGroupPlus = createAddPlusIcon();
    const addGroupSuffix = document.createElement("span");
    addGroupSuffix.className = "add-btn-suffix";
    addGroupSuffix.textContent = "G";
    addGroupBtn.appendChild(addGroupPlus);
    addGroupBtn.appendChild(addGroupSuffix);
    addGroupBtn.title = "Create group from this text";
    addGroupBtn.addEventListener("click", () => {
      const text = addInput.value.trim();
      if (!text) return;
      addGroupWithName(secIdx, text);
      addInput.value = "";
    });
    addRow.appendChild(addGroupBtn);
  }
  return addRow;
}

function createAddSectionRow() {
  if (state.hideAddButtons) return document.createComment("add-section-row-hidden");
  const addRow = document.createElement("div");
  addRow.className = "section-add-row section-add-section-row";
  const addInput = document.createElement("input");
  addInput.type = "text";
  addInput.placeholder = "Section name...";
  addInput.spellcheck = false;
  const addBtn = document.createElement("button");
  addBtn.type = "button";
  addBtn.textContent = "+ Section";
  addBtn.title = "Add new section";
  const submit = () => {
    const text = addInput.value.trim();
    if (!text) return;
    addSectionWithName(text);
    addInput.value = "";
  };
  addBtn.addEventListener("click", submit);
  addInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") submit();
  });
  addRow.appendChild(addInput);
  addRow.appendChild(addBtn);
  return addRow;
}

function renderSentences(options = {}) {
  const { force = false, includePreview = true } = options;
  if (!force && hasActiveInlineEditor()) {
    state.ui.pendingSentenceRender = true;
    if (includePreview) {
      state.ui.pendingPreviewRender = true;
    }
    return false;
  }

  state.ui.pendingSentenceRender = false;
  const shouldRenderPreview = includePreview || state.ui.pendingPreviewRender;
  state.ui.pendingPreviewRender = false;

  sectionContainer.innerHTML = "";
  sentenceListElements.clear();
  const selectedPaths = [...state.selectedPaths];

  state.sections = normalizeSectionsData(state.sections);

  state.sections.forEach((section, secIdx) => {
    const sectionBlock = document.createElement("div");
    sectionBlock.className = "section-group";
    sectionBlock.dataset.sectionIdx = secIdx;

    const header = document.createElement("div");
    header.className = "section-header";
    header.draggable = true;

    const dragHandle = document.createElement("span");
    dragHandle.className = "drag-handle";
    dragHandle.title = "Drag to reorder section";
    header.appendChild(dragHandle);

    const collapseBtn = document.createElement("button");
    collapseBtn.type = "button";
    collapseBtn.className = "collapse-btn";
    collapseBtn.textContent = isSectionCollapsed(section) ? "▸" : "▾";
    collapseBtn.title = isSectionCollapsed(section) ? "Expand section" : "Collapse section";
    collapseBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      setSectionCollapsed(section, !isSectionCollapsed(section));
      renderSentences();
    });
    header.appendChild(collapseBtn);

    const nameSpan = document.createElement("span");
    nameSpan.className = "section-name";
    nameSpan.draggable = false;
    nameSpan.textContent = section.name || "(General)";
    nameSpan.addEventListener("click", (e) => {
      e.stopPropagation();
      beginEditableRename(nameSpan, section.name || "", (newName) => {
        renameSection(section.name || "", newName);
      });
    });
    nameSpan.addEventListener("mousedown", (e) => {
      if (nameSpan.contentEditable === "true") e.stopPropagation();
    });
    header.appendChild(nameSpan);

    const sectionRefreshBtn = document.createElement("button");
    sectionRefreshBtn.type = "button";
    sectionRefreshBtn.className = "section-refresh-btn generate-btn generate-btn-icon";
    sectionRefreshBtn.title = "Refresh this section with auto captioning";
    sectionRefreshBtn.disabled = state.autoCaptioning || state.selectedPaths.size === 0 || !state.ollamaModel.trim();
    setGenerateButtonContent(sectionRefreshBtn, "Refresh this section with auto captioning", { iconOnly: true });
    sectionRefreshBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      refreshSectionCaptions(secIdx);
    });
    header.appendChild(sectionRefreshBtn);

    const removeBtn = document.createElement("button");
    removeBtn.className = "section-remove-btn";
    removeBtn.textContent = "\u00D7";
    removeBtn.title = "Delete section";
    removeBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      deleteSection(secIdx);
    });
    header.appendChild(removeBtn);

    header.addEventListener("dragstart", (e) => {
      if (nameSpan.contentEditable === "true") { e.preventDefault(); return; }
      setDragPayload(e, SECTION_DRAG_TYPE, { secIdx });
      header.classList.add("dragging-section");
    });
    header.addEventListener("dragend", () => {
      header.classList.remove("dragging-section");
      document.querySelectorAll(".section-header.drag-over").forEach(h => h.classList.remove("drag-over"));
    });
    header.addEventListener("dragover", (e) => {
      const payload = getDragPayload(e, SECTION_DRAG_TYPE);
      if (!payload) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      header.classList.add("drag-over");
    });
    header.addEventListener("dragleave", () => {
      header.classList.remove("drag-over");
    });
    header.addEventListener("drop", (e) => {
      const payload = getDragPayload(e, SECTION_DRAG_TYPE);
      if (!payload) return;
      e.preventDefault();
      header.classList.remove("drag-over");
      const fromIdx = Number.parseInt(payload.secIdx, 10);
      const toIdx = secIdx;
      if (Number.isNaN(fromIdx) || fromIdx === toIdx) return;
      const [moved] = state.sections.splice(fromIdx, 1);
      state.sections.splice(toIdx, 0, moved);
      saveSectionsToStorage();
      renderSentences();
    });

    sectionBlock.appendChild(header);

    const body = document.createElement("div");
    body.className = `section-body${isSectionCollapsed(section) ? " collapsed" : ""}`;

    const sectionSentences = document.createElement("ul");
    sectionSentences.className = "section-sentences section-mixed-list";
    sectionSentences.addEventListener("dragover", (e) => {
      const payload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
      if (!payload || payload.secIdx !== secIdx || !isValidTopLevelDragRef(payload.item)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      if (e.target === sectionSentences) {
        sectionSentences.classList.add("drag-over-end");
      }
    });
    sectionSentences.addEventListener("dragleave", (e) => {
      if (!sectionSentences.contains(e.relatedTarget)) {
        sectionSentences.classList.remove("drag-over-end");
      }
    });
    sectionSentences.addEventListener("drop", (e) => {
      const payload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
      if (!payload || payload.secIdx !== secIdx || !isValidTopLevelDragRef(payload.item)) return;
      if (e.target !== sectionSentences) return;
      e.preventDefault();
      sectionSentences.classList.remove("drag-over-end");
      moveTopLevelSectionItemToEnd(secIdx, payload.item);
    });
    for (const entry of getOrderedSectionEntries(section)) {
      if (entry.type === "sentence") {
        sectionSentences.appendChild(createSentenceListItem(entry.sentence, selectedPaths, {
          secIdx,
          topLevelMix: true,
        }));
      } else if (entry.type === "group") {
        sectionSentences.appendChild(createGroupListItem(section, entry.group, secIdx, entry.groupIdx, selectedPaths));
      }
    }
    body.appendChild(sectionSentences);

    body.appendChild(createAddSentenceRow(secIdx, null, "Add caption to section..."));
    sectionBlock.appendChild(body);
    sectionContainer.appendChild(sectionBlock);
  });

  sectionContainer.appendChild(createAddSectionRow());

  updateActionButtons();
  renderAddButtonsVisibility();
  if (shouldRenderPreview) {
    renderPreviewCaptionOverlay();
  }
  return true;
}

function getAutoCaptionScope({ freeTextOnly = false, targetSectionIndex = null, targetGroupIndex = null, targetSentence = null } = {}) {
  if (freeTextOnly) return "free-text-only";
  if (targetSentence && String(targetSentence).trim()) return "sentence";
  if (targetSectionIndex !== null && targetGroupIndex !== null) return "group";
  if (targetSectionIndex !== null) return "section";
  return "full";
}

function getAutoCaptionScopeLabel(scope) {
  switch (scope) {
    case "free-text-only": return "Add Now";
    case "sentence": return "Caption refresh";
    case "group": return "Group refresh";
    case "section": return "Section refresh";
    default: return "Auto Caption";
  }
}

function getAutoCaptionScopeLogLabel(scope) {
  return scope === "free-text-only" ? "Free-text enhancement" : getAutoCaptionScopeLabel(scope);
}

function getAutoCaptionStatusIntro(scope, count) {
  const suffix = `for ${count} media file${count === 1 ? "" : "s"}...`;
  switch (scope) {
    case "free-text-only": return `Adding free text ${suffix}`;
    case "sentence": return `Refreshing caption ${suffix}`;
    case "group": return `Refreshing group ${suffix}`;
    case "section": return `Refreshing section ${suffix}`;
    default: return `Auto captioning ${suffix}`;
  }
}

function getAutoCaptionStartLog(scope, event) {
  const count = event.count;
  const target = `${count} media file(s) using ${event.model} @ ${event.host}`;
  switch (scope) {
    case "free-text-only": return `Starting free-text enhancement for ${target}`;
    case "sentence": return `Starting caption refresh for ${target}`;
    case "group": return `Starting group refresh for ${target}`;
    case "section": return `Starting section refresh for ${target}`;
    default: return `Starting auto caption for ${target}`;
  }
}

function getAutoCaptionImageStartLog(scope, event, targetSentence) {
  const fileLabel = getFileLabel(event.path);
  if (scope === "free-text-only") {
    return `[${fileLabel}] Generating free-text details...`;
  }
  if (scope === "sentence") {
    return `[${fileLabel}] Refreshing caption ${targetSentence || ""}...`;
  }
  if (event.target_scope?.type === "group") {
    return `[${fileLabel}] Refreshing group ${event.target_scope.group_name || "(Group)"}...`;
  }
  if (event.target_scope?.type === "section") {
    return `[${fileLabel}] Refreshing section ${event.target_scope.section_name || "(General)"}...`;
  }
  return `[${fileLabel}] Checking ${event.total_targets || event.total_sentences} caption target(s)...`;
}

function getAutoCaptionCompleteLog(scope, event) {
  const fileLabel = getFileLabel(event.path);
  switch (scope) {
    case "free-text-only":
      return `[${fileLabel}] Free-text enhancement complete`;
    case "sentence":
      return `[${fileLabel}] Caption refresh complete`;
    case "group":
      return `[${fileLabel}] Group refresh complete`;
    case "section":
      return `[${fileLabel}] Section refresh complete`;
    default:
      return `[${fileLabel}] Completed with ${((event.enabled_captions || event.enabled_sentences) || []).length} matched captions`;
  }
}

function getAutoCaptionProgressText(scope, processed, total, errors = 0, finished = false) {
  const label = getAutoCaptionScopeLabel(scope);
  const base = finished
    ? `${label} finished: ${processed}/${total} processed`
    : `${label}: ${processed}/${total} done`;
  if (!errors) return base;
  if (finished) return `${base}, ${errors} errors`;
  return `${base}, ${errors} error${errors === 1 ? "" : "s"}`;
}

async function runAutoCaptionStream({ freeTextOnly = false, targetSectionIndex = null, targetGroupIndex = null, targetSentence = null, enableFreeText = null } = {}) {
  if (state.autoCaptioning) {
    stopAutoCaption();
    return;
  }
  if (state.selectedPaths.size === 0 || state.autoCaptioning) return;
  const paths = [...state.selectedPaths];
  const model = state.ollamaModel.trim();
  if (!model) {
    statusBar.textContent = "Configure an Ollama model first";
    openSettingsModal();
    return;
  }
  if (!freeTextOnly && !hasConfiguredCaptions()) {
    statusBar.textContent = "No captions configured for this folder";
    return;
  }
  const scope = getAutoCaptionScope({ freeTextOnly, targetSectionIndex, targetGroupIndex, targetSentence });
  const autoPreviewConfigError = state.comfyuiAutoPreviewEnabled ? getPromptPreviewConfigError({ requireWorkflow: true }) : "";
  const autoPreviewEnabled = !!state.comfyuiAutoPreviewEnabled && !autoPreviewConfigError;

  state.autoCaptioning = true;
  state.autoCaptionMode = scope;
  state.autoCaptionAbortController = new AbortController();
  resetAutoCaptionProgress();
  updateActionButtons();
  clearModelLog();
  statusBar.textContent = getAutoCaptionStatusIntro(scope, paths.length);

  try {
    const resp = await fetch("/api/auto-caption/stream", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal: state.autoCaptionAbortController.signal,
      body: JSON.stringify({
        media_paths: paths,
        model,
        prompt_template: state.ollamaPromptTemplate,
        group_prompt_template: state.ollamaGroupPromptTemplate,
        enable_free_text: enableFreeText ?? state.ollamaEnableFreeText,
        free_text_only: freeTextOnly,
        target_section_index: targetSectionIndex,
        target_group_index: targetGroupIndex,
        target_caption: targetSentence,
        free_text_prompt_template: state.ollamaFreeTextPromptTemplate,
        timeout_seconds: state.ollamaTimeoutSeconds,
      }),
    });
    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      throw new Error(data.detail || "Auto caption failed");
    }

    let processed = 0;
    let errors = 0;
    await readNdjsonStream(resp, (event) => {
      if (event.type === "start") {
        updateAutoCaptionProgress({
          visible: true,
          scopeLabel: getAutoCaptionScopeLabel(scope),
          totalImages: event.count || paths.length,
          processedImages: 0,
          errors: 0,
          completedImages: 0,
          enableFreeText: !!event.enable_free_text,
          freeTextOnly: !!event.free_text_only,
          currentMessage: "Preparing media...",
        });
        appendModelLog(getAutoCaptionStartLog(scope, event), "log-dim");
        if (event.enable_free_text || event.free_text_only) {
          appendModelLog("Free-text enhancement is enabled", "log-dim");
        }
        if (state.comfyuiAutoPreviewEnabled) {
          appendModelLog(
            autoPreviewEnabled
              ? "Auto Preview is enabled"
              : `Auto Preview skipped for this run: ${autoPreviewConfigError}`,
            autoPreviewEnabled ? "log-dim" : "log-warn"
          );
        }
      } else if (event.type === "image-start") {
        const baseSteps = event.free_text_only
          ? 1
          : Math.max(0, event.total_targets || 0) + (state.aiProgress.enableFreeText ? 1 : 0);
        updateAutoCaptionProgress({
          currentPath: event.path,
          currentMessage: event.free_text_only ? "Generating free-text details" : "Checking captions",
          currentStepIndex: 0,
          currentStepTotal: Math.max(1, baseSteps),
        });
        const cache = ensureCaptionCache(event.path);
        cache.enabled_sentences = orderEnabledSentences(event.enabled_captions || event.enabled_sentences || cache.enabled_sentences || []);
        cache.free_text = event.free_text || cache.free_text || "";
        markCaptionIndicator(event.path, hasEffectiveCaptionContent(cache));
        refreshGridForActiveFilters();
        scheduleUiRender({ sentences: true });
        appendModelLog(getAutoCaptionImageStartLog(scope, event, targetSentence), "log-dim");
      } else if (event.type === "caption-check") {
        const currentCaption = event.caption || event.sentence || "";
        updateAutoCaptionProgress({
          currentPath: event.path,
          currentMessage: currentCaption,
          currentStepIndex: event.index || state.aiProgress.currentStepIndex,
          currentStepTotal: Math.max(state.aiProgress.currentStepTotal || 0, (event.total || 0) + (state.aiProgress.enableFreeText ? 1 : 0)),
        });
        const verdict = event.enabled ? "YES" : "NO";
        appendModelLog(`[${getFileLabel(event.path)}] ${event.index}/${event.total} ${currentCaption} -> ${verdict} | ${event.answer || verdict}`, event.enabled ? "log-ok" : "log-dim");
        const cache = ensureCaptionCache(event.path);
        cache.enabled_sentences = orderEnabledSentences(event.enabled_captions || event.enabled_sentences || cache.enabled_sentences || []);
        cache.free_text = event.free_text || cache.free_text || "";
        refreshGridForActiveFilters();
        if (state.selectedPaths.has(event.path)) {
          scheduleUiRender({ sentences: true });
        }
        const hasContent = hasEffectiveCaptionContent(cache);
        markCaptionIndicator(event.path, !!hasContent);
      } else if (event.type === "group-selection") {
        const selectedCaption = event.selected_caption || event.selected_sentence || "(no valid selection)";
        updateAutoCaptionProgress({
          currentPath: event.path,
          currentMessage: event.group_name || "Group selection",
          currentStepIndex: event.index || state.aiProgress.currentStepIndex,
          currentStepTotal: Math.max(state.aiProgress.currentStepTotal || 0, (event.total || 0) + (state.aiProgress.enableFreeText ? 1 : 0)),
        });
        appendModelLog(
          `[${getFileLabel(event.path)}] ${event.index}/${event.total} ${event.group_name || "Group"} -> ${selectedCaption}${event.selected_hidden ? " (ignored in txt)" : ""} | ${event.answer || ""}`,
          (event.selected_caption || event.selected_sentence) ? "log-ok" : "log-warn"
        );
        const cache = ensureCaptionCache(event.path);
        cache.enabled_sentences = orderEnabledSentences(event.enabled_captions || event.enabled_sentences || cache.enabled_sentences || []);
        cache.free_text = event.free_text || cache.free_text || "";
        refreshGridForActiveFilters();
        if (state.selectedPaths.has(event.path)) {
          scheduleUiRender({ sentences: true });
        }
        const hasContent = hasEffectiveCaptionContent(cache);
        markCaptionIndicator(event.path, !!hasContent);
      } else if (event.type === "free-text") {
        updateAutoCaptionProgress({
          currentPath: event.path,
          currentMessage: "Adding free text",
          currentStepIndex: state.aiProgress.currentStepTotal || 1,
          currentStepTotal: state.aiProgress.currentStepTotal || 1,
        });
        const cache = ensureCaptionCache(event.path);
        cache.enabled_sentences = orderEnabledSentences(event.enabled_captions || event.enabled_sentences || cache.enabled_sentences || []);
        cache.free_text = event.free_text || cache.free_text || "";
        if ((event.added_lines || []).length > 0) {
          appendModelLog(`[${getFileLabel(event.path)}] Added free text: ${(event.added_lines || []).join(" | ")}`, "log-ok");
        } else {
          appendModelLog(`[${getFileLabel(event.path)}] Free text output: ${event.answer || "NONE"}`, "log-warn");
        }
        if (state.selectedPaths.size === 1 && state.selectedPaths.has(event.path)) {
          freeText.value = cache.free_text || "";
        }
        markCaptionIndicator(event.path, hasEffectiveCaptionContent(cache));
      } else if (event.type === "image-complete") {
        processed++;
        updateAutoCaptionProgress({
          processedImages: processed,
          completedImages: processed + errors,
          currentPath: event.path,
          currentMessage: "Completed",
          currentStepIndex: state.aiProgress.currentStepTotal || 1,
        });
        state.captionCache[event.path] = normalizeCaptionCacheEntry(event);
        refreshGridForActiveFilters();
        const hasContent = hasEffectiveCaptionContent(state.captionCache[event.path]);
        markCaptionIndicator(event.path, !!hasContent);
        appendModelLog(getAutoCaptionCompleteLog(scope, event), "log-ok");
        if (state.selectedPaths.size === 1 && state.selectedPaths.has(event.path)) {
          freeText.value = event.free_text || "";
        }
        if (autoPreviewEnabled && hasContent) {
          queuePromptPreviewFromCurrentCaption(event.path)
            .then((queued) => {
              appendModelLog(`[${getFileLabel(event.path)}] Auto Preview queued`, "log-dim");
              if (state.previewPath === event.path && state.previewMediaType === "image") {
                state.promptPreview.sourcePath = event.path;
                applyPromptPreviewSnapshot(event.path, queued, { autoDisplayLatest: false });
                renderPromptPreviewControls();
              }
            })
            .catch((queueErr) => {
              appendModelLog(`[${getFileLabel(event.path)}] Auto Preview failed: ${queueErr.message}`, "log-warn");
            });
        }
        scheduleUiRender({ sentences: true });
        statusBar.textContent = getAutoCaptionProgressText(scope, processed, paths.length, errors, false);
      } else if (event.type === "error") {
        errors++;
        updateAutoCaptionProgress({
          errors,
          completedImages: processed + errors,
          currentPath: event.path || state.aiProgress.currentPath,
          currentMessage: "Error",
          currentStepIndex: state.aiProgress.currentStepTotal || 0,
        });
        appendModelLog(`[${getFileLabel(event.path)}] ${event.message}`, "log-err");
        statusBar.textContent = getAutoCaptionProgressText(scope, processed, paths.length, errors, false);
      } else if (event.type === "done") {
        processed = event.processed ?? processed;
        errors = event.errors ?? errors;
        updateAutoCaptionProgress({
          processedImages: processed,
          errors,
          completedImages: processed + errors,
          currentMessage: "Finished",
          currentStepIndex: 0,
          currentStepTotal: 0,
        });
        appendModelLog(`Finished: ${processed}/${event.count} processed, ${errors} error${errors === 1 ? "" : "s"}`, errors ? "log-warn" : "log-ok");
        statusBar.textContent = getAutoCaptionProgressText(scope, processed, event.count, errors, true);
      }
    });

    if (state.selectedPaths.size > 1) {
      scheduleUiRender({ sentences: true });
    }
  } catch (err) {
    if (err.name === "AbortError") {
      updateAutoCaptionProgress({ currentMessage: "Stopped" });
      statusBar.textContent = `${getAutoCaptionScopeLabel(scope)} stopped`;
      appendModelLog(`${getAutoCaptionScopeLogLabel(scope)} stopped by user`, "log-warn");
    } else {
      updateAutoCaptionProgress({ currentMessage: "Failed" });
      const scopeLabel = getAutoCaptionScopeLabel(scope);
      const logLabel = getAutoCaptionScopeLogLabel(scope);
      statusBar.textContent = `${scopeLabel} failed: ${err.message}`;
      appendModelLog(`${logLabel} failed: ${err.message}`, "log-err");
    }
  } finally {
    state.autoCaptioning = false;
    state.autoCaptionMode = null;
    state.autoCaptionAbortController = null;
    resetAutoCaptionProgress();
    if (hasActiveInlineEditor()) {
      scheduleUiRender({ sentences: true, preview: true });
    } else {
      flushQueuedUiRenders({ force: true });
      renderSentences({ force: true });
    }
    renderModelLog();
    updateActionButtons();
  }
}

async function autoCaptionSelected() {
  return runAutoCaptionStream({ freeTextOnly: false });
}

async function addFreeTextNow() {
  return runAutoCaptionStream({ freeTextOnly: true });
}

async function refreshGroupCaptions(sectionIndex, groupIndex) {
  return runAutoCaptionStream({
    targetSectionIndex: sectionIndex,
    targetGroupIndex: groupIndex,
    enableFreeText: false,
  });
}

async function refreshSectionCaptions(sectionIndex) {
  return runAutoCaptionStream({
    targetSectionIndex: sectionIndex,
    enableFreeText: false,
  });
}

async function refreshSentenceCaption(sentence) {
  return runAutoCaptionStream({
    targetSentence: sentence,
    enableFreeText: false,
  });
}

function stopAutoCaption() {
  if (!state.autoCaptioning || !state.autoCaptionAbortController) return;
  state.autoCaptionAbortController.abort();
}

async function toggleSentence(sentence, wasChecked, wasPartial) {
  const selectedPaths = [...state.selectedPaths];
  if (selectedPaths.length === 0) return;

  const shouldEnable = !wasChecked && !wasPartial;

  statusBar.textContent = "Saving...";

  try {
    const resp = await fetch("/api/caption/batch-toggle", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        image_paths: selectedPaths,
        caption: sentence,
        enabled: shouldEnable,
      }),
    });
    if (resp.ok) {
      // Update local cache
      for (const path of selectedPaths) {
        const cap = ensureCaptionCache(path);
        cap.enabled_sentences = applySentenceSelectionToList(cap.enabled_sentences, sentence, shouldEnable);
      }
      refreshGridForActiveFilters();
      renderSentences();
      renderPreviewCaptionOverlay();
      // Update caption indicators on thumbnails
      for (const path of selectedPaths) {
        const cap = state.captionCache[path];
        const hasContent = hasEffectiveCaptionContent(cap);
        markCaptionIndicator(path, hasContent);
      }
      statusBar.textContent = "Saved";
    }
  } catch (err) {
    statusBar.textContent = `Error: ${err.message}`;
  }
}

async function saveMetadataForSelection() {
  const selectedPaths = [...state.selectedPaths];
  if (!selectedPaths.length || state.metadataSaving) return;

  let singlePath = null;
  let singleMetadata = null;
  let batchChanges = null;

  try {
    if (selectedPaths.length === 1) {
      singlePath = selectedPaths[0];
      singleMetadata = buildSingleMetadataPayload();
    } else {
      batchChanges = buildBatchMetadataChanges();
      if (Object.keys(batchChanges).length === 0) {
        throw new Error("Enter at least one metadata value to apply");
      }
    }
  } catch (err) {
    const message = err?.message || "Failed to save metadata";
    statusBar.textContent = `Metadata error: ${message}`;
    showErrorToast(`Metadata error: ${message}`);
    return;
  }

  state.metadataSaving = true;
  renderMetadataEditor({ preserveInputs: true });
  statusBar.textContent = selectedPaths.length === 1
    ? "Saving metadata..."
    : `Applying metadata to ${selectedPaths.length} media files...`;

  try {
    if (selectedPaths.length === 1) {
      const resp = await fetch("/api/media/meta/save", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ path: singlePath, metadata: singleMetadata }),
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        throw new Error(data.detail || "Failed to save metadata");
      }
      state.metadataCache[singlePath] = normalizeMetadataCacheEntry(data.metadata);
      statusBar.textContent = "Metadata saved";
    } else {
      const resp = await fetch("/api/media/meta/apply", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ paths: selectedPaths, changes: batchChanges }),
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        throw new Error(data.detail || "Failed to apply metadata");
      }

      let updatedCount = 0;
      const errors = [];
      for (const result of (data.results || [])) {
        if (result?.ok) {
          state.metadataCache[result.path] = normalizeMetadataCacheEntry(result.metadata);
          updatedCount += 1;
        } else if (result?.path || result?.error) {
          errors.push(result);
        }
      }

      if (errors.length > 0) {
        const summary = errors.length === 1
          ? `${updatedCount} updated, 1 failed: ${errors[0].error || "Unknown error"}`
          : `${updatedCount} updated, ${errors.length} failed`;
        statusBar.textContent = summary;
        showErrorToast(summary);
      } else {
        statusBar.textContent = `Applied metadata to ${updatedCount} media file${updatedCount === 1 ? "" : "s"}`;
      }
    }
  } catch (err) {
    const message = err?.message || "Failed to save metadata";
    statusBar.textContent = `Metadata error: ${message}`;
    showErrorToast(`Metadata error: ${message}`);
  } finally {
    state.metadataSaving = false;
    renderMetadataEditor();
    updateMultiInfo();
  }
}

function showSentenceContextMenu() { /* removed - using inline remove buttons now */ }

function showSectionContextMenu() { /* removed - using inline controls now */ }

// ===== FREE TEXT =====
let freeTextSaveTimeout = null;
freeText.addEventListener("input", () => {
  if (state.selectedPaths.size !== 1) return;
  const path = [...state.selectedPaths][0];

  // Update cache
  if (!state.captionCache[path]) {
    state.captionCache[path] = { enabled_sentences: [], free_text: "" };
  }
  state.captionCache[path].free_text = freeText.value;

  // Debounced save
  if (freeTextSaveTimeout) clearTimeout(freeTextSaveTimeout);
  freeTextSaveTimeout = setTimeout(() => saveFreeText(path), 400);
});

async function saveFreeText(path) {
  const cap = state.captionCache[path];
  if (!cap) return;
  statusBar.textContent = "Saving...";
  try {
    await fetch("/api/caption/save", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        image_path: path,
        enabled_captions: cap.enabled_sentences || [],
        free_text: cap.free_text || "",
      }),
    });
    statusBar.textContent = "Saved";
    // Update caption indicator
    const hasContent = hasEffectiveCaptionContent(cap);
    markCaptionIndicator(path, !!hasContent);
  } catch (err) {
    statusBar.textContent = `Error saving: ${err.message}`;
  }
}

// ===== THUMBNAIL SIZE SLIDER =====
thumbSlider.addEventListener("input", () => {
  state.thumbSize = parseInt(thumbSlider.value);
  document.documentElement.style.setProperty("--thumb-size", state.thumbSize + "px");

  // Re-render with new size
  const cells = fileGrid.querySelectorAll(".thumb-cell");
  cells.forEach(cell => {
    cell.style.width = state.thumbSize + "px";
    cell.style.height = state.thumbSize + "px";
  });
});

// When slider released, reload thumbnails at appropriate resolution
thumbSlider.addEventListener("change", () => {
  fetch("/api/settings", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ thumb_size: state.thumbSize }),
  }).catch(err => console.error("Failed to save thumbnail size:", err));

  const thumbLoadSize = getThumbLoadSize();
  let queuedThumbnailCount = 0;
  for (const img of state.images) {
    if (queueThumbLoad(img.path, thumbLoadSize)) {
      queuedThumbnailCount += 1;
    }
  }
  if (queuedThumbnailCount > 0) {
    startThumbnailProgress(queuedThumbnailCount, `Processing thumbnails 0/${queuedThumbnailCount}`);
  }
});

// ===== RESIZE HANDLES =====
