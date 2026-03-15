function getCloneSelectionPaths() {
  return state.selectedPaths.size > 1 ? [...state.selectedPaths] : [];
}

function getDefaultCloneFolderName() {
  const baseName = getFileLabel(state.folder || "folder") || "folder";
  return `${baseName}-copy`;
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
  if (item.type === "sentence") return `sentence:${item.sentence || ""}`;
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
    if (type === "sentence") {
      const sentence = String(rawItem.sentence || "").trim();
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
  for (const raw of Array.isArray(group?.sentences) ? group.sentences : []) {
    const text = String(raw || "").trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    sentences.push(text);
  }
  const hiddenSentences = [];
  const hiddenSeen = new Set();
  for (const raw of Array.isArray(group?.hidden_sentences) ? group.hidden_sentences : []) {
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
  for (const raw of Array.isArray(section?.sentences) ? section.sentences : []) {
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
    sentences: [...section.sentences],
    groups: section.groups.map(group => ({
      id: group.id,
      name: group.name,
      sentences: [...group.sentences],
      hidden_sentences: [...(group.hidden_sentences || [])],
    })),
    item_order: (section.item_order || []).map(item => item.type === "sentence"
      ? createSentenceOrderItem(item.sentence)
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
  return getExportedEnabledSentences(caption.enabled_sentences).length > 0 || !!(caption.free_text && caption.free_text.trim());
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
      sentences,
    }),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data.detail || "Failed to load captions");
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
      state.captionCache[path] = caption;
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
  if (!state.previewPath || !imgNatW || !imgNatH || !isPreviewVisible() || sentences.length === 0) {
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

function buildImageApiUrl(endpoint, path, extraParams = {}) {
  const params = new URLSearchParams({ path, v: String(getImageVersion(path)) });
  Object.entries(extraParams).forEach(([key, value]) => {
    params.set(key, String(value));
  });
  return `/api/${endpoint}?${params.toString()}`;
}

function queueThumbLoad(path, size, priority = false) {
  const key = `${path}:${size}:${getImageVersion(path)}`;
  if (thumbBlobCache.has(key) || thumbQueuedKeys.has(key)) return false;
  const item = { path, size, key };
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
    const { path, size, key } = thumbLoadQueue.shift();
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
          thumbBlobCache.set(key, url);
          // Update any visible img with this path
          const imgs = fileGrid.querySelectorAll(`img[data-path="${CSS.escape(path)}"]`);
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
let previewLoadingSet = new Set(); // paths currently being fetched

function preloadPreview(path) {
  if (previewCache.has(path) || previewLoadingSet.has(path)) return;
  previewLoadingSet.add(path);
  fetch(buildImageApiUrl("preview", path))
    .then(r => r.ok ? r.blob() : null)
    .then(blob => {
      if (blob) {
        const url = URL.createObjectURL(blob);
        previewCache.set(path, url);
        // If currently selected, swap in the higher-quality preview
        if (state.previewPath === path) {
          if (state.previewMediaType === "video") {
            previewVideo.poster = url;
            return;
          }
          const oldNatW = imgNatW;
          const oldNatH = imgNatH;
          const oldZoom = zoomLevel;
          const oldPanX = panX;
          const oldPanY = panY;
          const wasUserZoomed = userHasZoomed;
          previewImg.onload = () => {
            imgNatW = previewImg.naturalWidth;
            imgNatH = previewImg.naturalHeight;
            if (wasUserZoomed && oldNatW > 0 && oldNatH > 0) {
              // Preserve user's zoom/pan relative to old dimensions
              const scaleRatio = oldNatW / imgNatW;
              zoomLevel = oldZoom * scaleRatio;
              // Adjust pan: the image center that was on screen should stay there
              const panel = previewStage;
              const pw = panel.clientWidth;
              const ph = panel.clientHeight;
              // The point at the center of the panel in old image coords:
              const cx = pw / 2;
              const cy = ph / 2;
              // Map: panelPoint = pan + imgPoint * zoom
              // oldImgPoint = (cx - oldPanX) / oldZoom
              // newPan = cx - oldImgPoint * (oldZoom * scaleRatio) / scaleRatio  -- simplifies to same
              // Actually: old image coord = (panelPt - oldPan) / oldZoom
              // new image coord = oldImgCoord * (newNatW / oldNatW) = oldImgCoord / scaleRatio
              // new pan = panelPt - newImgCoord * newZoom
              const oldImgX = (cx - oldPanX) / oldZoom;
              const oldImgY = (cy - oldPanY) / oldZoom;
              const newImgX = oldImgX / scaleRatio;
              const newImgY = oldImgY / scaleRatio;
              panX = cx - newImgX * zoomLevel;
              panY = cy - newImgY * zoomLevel;
              userHasZoomed = true;
              applyTransform();
            } else {
              resetZoomPan();
            }
            previewImg.style.display = "block";
          };
          previewImg.src = url;
        }
      }
    })
    .catch(() => {})
    .finally(() => previewLoadingSet.delete(path));
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
}

function canEditCrop() {
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

function applySettings(settings) {
  const thumbSize = Number(settings.thumb_size || state.thumbSize || 160) || 160;
  state.thumbSize = Math.max(60, Math.min(400, thumbSize));
  thumbSlider.value = String(state.thumbSize);
  document.documentElement.style.setProperty("--thumb-size", state.thumbSize + "px");
  setCropAspectRatios(settings.crop_aspect_ratios || state.cropAspectRatioLabels);
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
  if (settings.sections) {
    state.sections = normalizeSectionsData(settings.sections);
  }
  autoFreeTextCheckbox.checked = !!state.ollamaEnableFreeText;
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
  settingsCropAspectRatiosInput.value = state.cropAspectRatioLabels.join(", ");
  settingsHttpsCertInput.value = state.httpsCertFile;
  settingsHttpsKeyInput.value = state.httpsKeyFile;
  settingsHttpsPortInput.value = String(state.httpsPort);
  settingsRemoteHttpModeInput.value = state.remoteHttpMode;
  settingsFfmpegPathInput.value = state.ffmpegPath;
  settingsProcessingReservedCoresInput.value = String(state.processingReservedCores);
  settingsFfmpegThreadsInput.value = String(state.ffmpegThreads);
  settingsFfmpegHwaccelInput.value = state.ffmpegHwaccel;
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
  videoTimelineSelection.style.left = `${startX}px`;
  videoTimelineSelection.style.width = `${Math.max(0, endX - startX)}px`;
  videoTimelineStartHandle.style.left = `${startX}px`;
  videoTimelineEndHandle.style.left = `${endX}px`;
  videoTimelinePlayhead.style.left = `${playheadX}px`;
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
    parts.push(`${activeJob.type === "clip" ? "Clipping" : "Cropping"} ${getFileLabel(activeJob.video_path)}`);
    if (activeJob.message) parts.push(activeJob.message);
  }
  if (queuedCount > 0) {
    parts.push(`${queuedCount} queued`);
  }
  parts.push(`${completed}/${total} done`);
  videoJobText.textContent = parts.join(" • ");
  videoJobProgressFill.style.width = `${percent}%`;
  videoJobProgressFill.classList.toggle("active", running);
}

async function handleCompletedVideoJobs(jobs) {
  const relevantJobs = (jobs || []).filter((job) => String(job.folder || "") === String(state.folder || ""));
  if (!relevantJobs.length) return;
  const generatedOutputPaths = relevantJobs
    .filter((job) => job.status === "completed" && (job.type === "clip" || job.type === "crop") && job.output_path)
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
        showErrorToast(failed[0].error || `${failed[0].type === "clip" ? "Clip" : "Crop"} job failed`);
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

async function deleteSelectedImages() {
  if (state.selectedPaths.size === 0) return;

  const selectedPaths = [...state.selectedPaths];
  const count = selectedPaths.length;
  const confirmMessage = count === 1
    ? `Delete "${getFileLabel(selectedPaths[0])}"? This also deletes its .txt caption file.`
    : `Delete ${count} selected media files? This also deletes their .txt caption files.`;
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
      delete state.imageCrops[path];
      delete state.imageVersions[path];
    }

    await loadFolder({ preserveScrollTop });

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
  const cropAspectRatios = settingsCropAspectRatiosInput.value.split(",").map(s => s.trim()).filter(Boolean);
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
    const resp = await fetch("/api/settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
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
        crop_aspect_ratios: cropAspectRatios,
        ollama_prompt_template: promptTemplate,
        ollama_group_prompt_template: groupPromptTemplate,
        ollama_enable_free_text: enableFreeText,
        ollama_free_text_prompt_template: freeTextPromptTemplate,
      }),
    });
    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      throw new Error(data.detail || "Failed to save settings");
    }
    applySettings({
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
      crop_aspect_ratios: cropAspectRatios,
      ollama_prompt_template: promptTemplate,
      ollama_group_prompt_template: groupPromptTemplate,
      ollama_enable_free_text: enableFreeText,
      ollama_free_text_prompt_template: freeTextPromptTemplate,
    });
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

// ===== FOLDER LOADING =====
async function loadFolder(options = {}) {
  const { preserveScrollTop = null } = options;
  const folder = folderInput.value.trim();
  if (!folder) return;
  state.folder = folder;
  state.selectedPaths.clear();
  state.lastClickedIndex = -1;
  state.lastClickedPath = null;
  state.previewPath = null;
  state.captionCache = {};
  state.activeSentenceFilters.clear();
  state.activeMetaFilters.aspectState = "any";
  state.activeMetaFilters.captionState = "any";
  state.filterCaptionCacheKey = "";
  state.filterLoadingPromise = null;
  state.imageCrops = {};
  state.imageVersions = {};
  state.thumbnailDimensions = {};
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
    await loadCaptionData(lastPath);
    await loadCropData(lastPath);
    freeText.disabled = false;
  } else {
    await showPreview(lastPath);
    freeText.disabled = true;
    freeText.value = "(Multiple images selected)";
    await loadMultiCaptionState();
    clearCropDraft();
  }

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

// ===== SELECTION HANDLING =====
function handleThumbClick(index, event) {
  const img = state.images[index];
  if (!img) return;
  const visibleEntries = getVisibleImageEntries();
  const currentVisibleIndex = visibleEntries.findIndex(entry => entry.img.path === img.path);
  const lastVisibleIndex = visibleEntries.findIndex(entry => entry.img.path === state.lastClickedPath);

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
    showPreview(path);
    loadCaptionData(path);
    loadCropData(path);
    freeText.disabled = false;
  } else if (state.selectedPaths.size > 1) {
    // Show preview of clicked image
    showPreview(img.path);
    freeText.disabled = true;
    freeText.value = "(Multiple images selected)";
    loadMultiCaptionState();
    clearCropDraft();
  } else {
    hidePreview();
    freeText.disabled = true;
    freeText.value = "";
  }

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
  if (!previewEl) return;
  previewEl.style.transform = `translate(${panX}px, ${panY}px) scale(${zoomLevel})`;
  previewEl.style.width = imgNatW + "px";
  previewEl.style.height = imgNatH + "px";
  renderCropOverlay();
  renderPreviewCaptionOverlay();
}

async function showPreview(path) {
  const previousPath = state.previewPath;
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
      previewImg.style.display = "none";
      previewImg.onload = () => {
        imgNatW = previewImg.naturalWidth;
        imgNatH = previewImg.naturalHeight;
        resetZoomPan();
        previewImg.style.display = "block";
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
  const img = state.images.find(i => i.path === path);
  if (img) {
    previewInfo.textContent = state.previewMediaType === "video"
      ? `${img.name} • video`
      : (hasAppliedCrop() ? `${img.name} • cropped` : img.name);
    previewInfo.style.display = "block";
  }
  renderPreviewCaptionOverlay();
  renderVideoEditPanel();
}

function hidePreview() {
  state.previewPath = null;
  state.previewMediaType = null;
  previewImg.style.display = "none";
  stopPreviewVideo({ clearSource: true });
  previewPlaceholder.style.display = "flex";
  previewInfo.style.display = "none";
  imgNatW = 0;
  imgNatH = 0;
  renderPreviewCaptionOverlay();
  clearCropDraft();
  renderVideoEditPanel();
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
    if (canEditCrop()) e.preventDefault();
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
    if (e.button === 2 && canEditCrop()) {
      startCropCreate(e);
      updateCropGuideFromClient(e.clientX, e.clientY);
      e.preventDefault();
      return;
    }
    if (e.button !== 0) return;
    if (e.target.closest("#video-edit-panel, #preview-caption-overlay, #crop-apply-btn, #crop-cancel-btn, #crop-remove-btn, #rotate-controls")) {
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
    updateCropGuideFromClient(e.clientX, e.clientY);
  });

  panel.addEventListener("mouseleave", () => {
    clearCropGuide();
  });

  window.addEventListener("mousemove", (e) => {
    if (videoTimelineInteraction) {
      updateVideoTimelineInteraction(e.clientX);
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
previewVideo.addEventListener("pause", syncPreviewVideoPlaybackState);
previewVideo.addEventListener("timeupdate", handlePreviewVideoTimeUpdate);
previewVideo.addEventListener("seeked", syncPreviewVideoPlaybackState);
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
    const resp = await fetch(`/api/caption?path=${encodeURIComponent(path)}&sentences=${encodeURIComponent(sentencesJson)}`);
    if (resp.ok) {
      const data = await resp.json();
      state.captionCache[path] = data;
      if (state.activeSentenceFilters.size > 0) {
        state.filterCaptionCacheKey = getActiveSentenceFilterKey();
      }
      // Update UI if still selected
      if (state.selectedPaths.size === 1 && state.selectedPaths.has(path)) {
        freeText.value = data.free_text || "";
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

async function loadMultiCaptionState() {
  const paths = [...state.selectedPaths];

  // Batch load captions
  try {
    const data = await fetchCaptionsBulk(paths);
    for (const [path, caption] of Object.entries(data)) {
      state.captionCache[path] = caption;
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

function updateMultiInfo() {
  if (state.selectedPaths.size > 1) {
    multiInfo.style.display = "block";
    multiInfo.textContent = `${state.selectedPaths.size} images selected — toggling captions applies to all`;
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
        sentence,
      }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to delete caption");
    }
    state.sections = normalizeSectionsData(data.sections || state.sections);
    applyRemovedSentencesToLocalState(data.removed_sentences || [sentence]);
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
        old_sentence: oldSentence,
        new_sentence: nextSentence,
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
      return `[${fileLabel}] Completed with ${(event.enabled_sentences || []).length} matched captions`;
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
        target_sentence: targetSentence,
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
        cache.enabled_sentences = orderEnabledSentences(event.enabled_sentences || cache.enabled_sentences || []);
        cache.free_text = event.free_text || cache.free_text || "";
        markCaptionIndicator(event.path, hasEffectiveCaptionContent(cache));
        refreshGridForActiveFilters();
        scheduleUiRender({ sentences: true });
        appendModelLog(getAutoCaptionImageStartLog(scope, event, targetSentence), "log-dim");
      } else if (event.type === "caption-check") {
        updateAutoCaptionProgress({
          currentPath: event.path,
          currentMessage: event.sentence,
          currentStepIndex: event.index || state.aiProgress.currentStepIndex,
          currentStepTotal: Math.max(state.aiProgress.currentStepTotal || 0, (event.total || 0) + (state.aiProgress.enableFreeText ? 1 : 0)),
        });
        const verdict = event.enabled ? "YES" : "NO";
        appendModelLog(`[${getFileLabel(event.path)}] ${event.index}/${event.total} ${event.sentence} -> ${verdict} | ${event.answer || verdict}`, event.enabled ? "log-ok" : "log-dim");
        const cache = ensureCaptionCache(event.path);
        cache.enabled_sentences = orderEnabledSentences(event.enabled_sentences || cache.enabled_sentences || []);
        cache.free_text = event.free_text || cache.free_text || "";
        refreshGridForActiveFilters();
        if (state.selectedPaths.has(event.path)) {
          scheduleUiRender({ sentences: true });
        }
        const hasContent = hasEffectiveCaptionContent(cache);
        markCaptionIndicator(event.path, !!hasContent);
      } else if (event.type === "group-selection") {
        updateAutoCaptionProgress({
          currentPath: event.path,
          currentMessage: event.group_name || "Group selection",
          currentStepIndex: event.index || state.aiProgress.currentStepIndex,
          currentStepTotal: Math.max(state.aiProgress.currentStepTotal || 0, (event.total || 0) + (state.aiProgress.enableFreeText ? 1 : 0)),
        });
        const selectedSentence = event.selected_sentence || "(no valid selection)";
        appendModelLog(
          `[${getFileLabel(event.path)}] ${event.index}/${event.total} ${event.group_name || "Group"} -> ${selectedSentence}${event.selected_hidden ? " (ignored in txt)" : ""} | ${event.answer || ""}`,
          event.selected_sentence ? "log-ok" : "log-warn"
        );
        const cache = ensureCaptionCache(event.path);
        cache.enabled_sentences = orderEnabledSentences(event.enabled_sentences || cache.enabled_sentences || []);
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
        cache.enabled_sentences = orderEnabledSentences(event.enabled_sentences || cache.enabled_sentences || []);
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
        state.captionCache[event.path] = {
          enabled_sentences: event.enabled_sentences || [],
          free_text: event.free_text || "",
        };
        refreshGridForActiveFilters();
        const hasContent = hasEffectiveCaptionContent(event);
        markCaptionIndicator(event.path, !!hasContent);
        appendModelLog(getAutoCaptionCompleteLog(scope, event), "log-ok");
        if (state.selectedPaths.size === 1 && state.selectedPaths.has(event.path)) {
          freeText.value = event.free_text || "";
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
        sentence: sentence,
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
        enabled_sentences: cap.enabled_sentences || [],
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
