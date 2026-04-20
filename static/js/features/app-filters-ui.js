function getPreviewEnabledSentences(path = state.previewPath) {
  if (!path) return [];
  const enabledSentences = state.captionCache[path]?.enabled_sentences;
  return orderEnabledSentences(enabledSentences);
}

function normalizeMediaSearchQuery(value) {
  return String(value || "").trim().toLowerCase();
}

function hasActiveSearchQuery() {
  return !!normalizeMediaSearchQuery(state.mediaSearchQuery);
}

function getCaptionSearchableText(path) {
  const caption = state.captionCache[path];
  if (!caption) return "";
  const enabledSentences = Array.isArray(caption.enabled_sentences) ? caption.enabled_sentences : [];
  const freeText = String(caption.free_text || "");
  return `${enabledSentences.join("\n")}\n${freeText}`.toLowerCase();
}

function imageMatchesSearchQuery(image) {
  const query = normalizeMediaSearchQuery(state.mediaSearchQuery);
  if (!query) return true;
  const fileName = String(image?.name || "").trim().toLowerCase();
  if (fileName.includes(query)) return true;
  return getCaptionSearchableText(image?.path).includes(query);
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
  const searchQuery = String(state.mediaSearchQuery || "").trim();

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

  if (searchQuery) {
    if (lines.length > 0) lines.push("");
    lines.push(`Search: \"${searchQuery}\"`);
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
  return getActiveSentenceFilterEntries().length + getActiveMetaFilterCount() + (hasActiveSearchQuery() ? 1 : 0);
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
  if (filters.length > 0 && canApplyActiveSentenceFilters()) {
    const clauses = getActiveSentenceFilterClauses(filters);
    const enabled = new Set(getEnabledSentencesForPath(image?.path));
    const matchesSentenceFilters = clauses.every((clause) => {
      if (clause.type === "group") {
        return clause.sentences.some(sentence => enabled.has(sentence));
      }
      return clause.mode === "missing"
        ? !enabled.has(clause.sentence)
        : enabled.has(clause.sentence);
    });
    if (!matchesSentenceFilters) return false;
  }
  return imageMatchesSearchQuery(image);
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
  const searchActive = hasActiveSearchQuery();
  if ((filters.length === 0 && !searchActive) || state.images.length === 0) return;

  const filterKey = getActiveSentenceFilterKey();
  const needsReload = (filters.length > 0 && state.filterCaptionCacheKey !== filterKey) || state.images.some(img => !state.captionCache[img.path]);
  if (!needsReload) return;
  if (state.filterLoadingPromise) return state.filterLoadingPromise;

  state.filterLoadingPromise = (async () => {
    const data = await fetchCaptionsBulk(state.images.map(img => img.path));
    for (const [path, caption] of Object.entries(data || {})) {
      state.captionCache[path] = normalizeCaptionCacheEntry(caption);
    }
    if (filters.length > 0) {
      state.filterCaptionCacheKey = filterKey;
    }
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
  if (mediaSearchInput) {
    const searchQuery = String(state.mediaSearchQuery || "").trim();
    mediaSearchInput.classList.toggle("active", !!searchQuery);
    mediaSearchInput.title = searchQuery
      ? `Searching filenames and caption text for \"${searchQuery}\"`
      : "Search filenames and caption text";
  }
}

function clearSentenceFilters() {
  if (!hasAnyActiveFilters()) return;
  const preferredPath = state.previewPath || state.lastClickedPath || [...state.selectedPaths][0] || null;
  const previousScrollTop = fileGridContainer.scrollTop;
  state.activeSentenceFilters.clear();
  state.activeMetaFilters.aspectState = "any";
  state.activeMetaFilters.maskState = "any";
  state.activeMetaFilters.captionState = "any";
  state.mediaSearchQuery = "";
  if (mediaSearchInput) {
    mediaSearchInput.value = "";
  }
  state.filterCaptionCacheKey = "";
  renderGrid({ preservePath: preferredPath, preserveScrollTop: previousScrollTop });
  renderSentences();
  statusBar.textContent = "Filters cleared";
}

async function updateMediaSearchQuery(value) {
  const nextQuery = String(value || "");
  if (nextQuery === state.mediaSearchQuery) return;
  state.mediaSearchQuery = nextQuery;

  const preferredPath = state.previewPath || state.lastClickedPath || [...state.selectedPaths][0] || null;
  const previousScrollTop = fileGridContainer.scrollTop;
  renderGrid({ preservePath: preferredPath, preserveScrollTop: previousScrollTop });
  renderSentences();

  const normalizedQuery = normalizeMediaSearchQuery(nextQuery);
  if (!normalizedQuery) {
    const count = getActiveFilterCount();
    statusBar.textContent = count > 0
      ? `Filtered by ${count} filter${count === 1 ? "" : "s"}`
      : "Search cleared";
    return;
  }

  statusBar.textContent = `Searching ${state.images.length} media file${state.images.length === 1 ? "" : "s"}...`;
  try {
    await ensureCaptionCacheLoadedForFiltering();
    if (normalizeMediaSearchQuery(state.mediaSearchQuery) !== normalizedQuery) return;
    renderGrid({ preservePath: preferredPath, preserveScrollTop: previousScrollTop });
    renderSentences();
    const visibleCount = getVisibleImageEntries().length;
    statusBar.textContent = `Search matched ${visibleCount} media file${visibleCount === 1 ? "" : "s"}`;
  } catch (err) {
    if (normalizeMediaSearchQuery(state.mediaSearchQuery) !== normalizedQuery) return;
    statusBar.textContent = `Search error: ${err.message}`;
  }
}

function handleMediaSearchInput(event) {
  updateMediaSearchQuery(event?.target?.value || "").catch((err) => {
    statusBar.textContent = `Search error: ${err.message}`;
  });
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

