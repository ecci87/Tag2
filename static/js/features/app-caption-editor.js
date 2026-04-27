async function loadCaptionData(path) {
  const allSentences = getAllConfiguredSentences();
  const sentencesJson = JSON.stringify(allSentences);
  try {
    const resp = await fetch(`/api/caption?path=${encodeURIComponent(path)}&captions=${encodeURIComponent(sentencesJson)}`);
    if (resp.ok) {
      const data = await resp.json();
      const nextCaption = normalizeCaptionCacheEntry(data);
      if (shouldPreserveLocalFreeTextDraft(path, nextCaption)) {
        nextCaption.free_text = getLocalDraftFreeText(path);
      }
      state.captionCache[path] = nextCaption;
      if (state.activeSentenceFilters.size > 0) {
        state.filterCaptionCacheKey = getActiveSentenceFilterKey();
      }
      // Update UI if still selected
      if (state.selectedPaths.size === 1 && state.selectedPaths.has(path)) {
        freeText.value = state.captionCache[path].free_text || "";
        syncFreeTextHighlightState();
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

function getLocalDraftFreeText(path) {
  if (state.selectedPaths.size === 1 && state.selectedPaths.has(path)) {
    return String(freeText.value || "");
  }
  return String(state.captionCache[path]?.free_text || "");
}

function shouldPreserveLocalFreeTextDraft(path, nextCaption) {
  if (!state.captionDraftPaths.has(path)) return false;
  return getLocalDraftFreeText(path) !== String(nextCaption?.free_text || "");
}

async function loadMultiCaptionState() {
  const paths = [...state.selectedPaths];

  // Batch load captions
  try {
    const data = await fetchCaptionsBulk(paths);
    for (const [path, caption] of Object.entries(data)) {
      const nextCaption = normalizeCaptionCacheEntry(caption);
      if (shouldPreserveLocalFreeTextDraft(path, nextCaption)) {
        nextCaption.free_text = getLocalDraftFreeText(path);
      }
      state.captionCache[path] = nextCaption;
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
      multiInfo.textContent = `${state.selectedPaths.size} media files selected - filled metadata fields and changed checkbox state apply to all`;
    } else {
      multiInfo.textContent = `${state.selectedPaths.size} media files selected - toggling captions applies to all`;
    }
  } else {
    multiInfo.style.display = "none";
  }
}

function hasPendingMetadataChanges() {
  return !!state.metadataEditorDirty;
}

async function savePendingMetadataChangesBeforeContextChange() {
  if (!hasPendingMetadataChanges()) return true;
  return saveMetadataForSelection({ quietNoChanges: true });
}

// ===== SENTENCES =====
let _saveSentencesTimeout = null;
let _saveSectionsRequestToken = 0;

function formatCaptionFileUpdateStatus(baseMessage, touchedCaptionFiles) {
  const count = Number(touchedCaptionFiles);
  if (!Number.isFinite(count) || count <= 0) {
    return baseMessage;
  }
  const fileLabel = count === 1 ? "image caption file" : "image caption files";
  return `${baseMessage}, updated ${count} ${fileLabel}`;
}

function noteCaptionLibraryChanged() {
  syncFreeTextHighlightState();
}

function persistSectionsToStorage(requestToken, options = {}) {
  const {
    rewriteCaptionFiles = true,
    savingMessage = "Saving caption library...",
    successMessage = "Caption library saved",
  } = options;
  if (!state.folder) {
    return Promise.resolve({ ok: false, skipped: true });
  }
  statusBar.textContent = savingMessage;
  return fetch("/api/settings", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      sections: serializeSectionsForSave(),
      folder: state.folder,
      rewrite_caption_files: rewriteCaptionFiles,
    }),
  })
    .then(async (resp) => {
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        throw new Error(data.detail || "Failed to save caption library");
      }
      if (requestToken === _saveSectionsRequestToken) {
        statusBar.textContent = formatCaptionFileUpdateStatus(successMessage, data.touched_caption_files);
      }
      return { ok: true, data };
    })
    .catch((err) => {
      if (requestToken === _saveSectionsRequestToken) {
        console.error("Failed to save sections:", err);
        statusBar.textContent = `Caption library save error: ${err.message}`;
      }
      return { ok: false, error: err };
    });
}

function saveSectionsToStorage(options = {}) {
  noteCaptionLibraryChanged();
  const { immediate = false } = options;
  // Debounce saves to avoid hammering the server during rapid edits
  if (_saveSentencesTimeout) clearTimeout(_saveSentencesTimeout);
  const requestToken = ++_saveSectionsRequestToken;
  if (immediate) {
    _saveSentencesTimeout = null;
    return persistSectionsToStorage(requestToken, options);
  }
  _saveSentencesTimeout = setTimeout(() => {
    _saveSentencesTimeout = null;
    persistSectionsToStorage(requestToken, options);
  }, 300);
  return Promise.resolve({ ok: true, deferred: true });
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

function clearCaptionDragIndicators() {
  document.querySelectorAll(".section-sentences li.drag-over").forEach(item => item.classList.remove("drag-over"));
  document.querySelectorAll(".section-sentences.drag-over-end").forEach(list => list.classList.remove("drag-over-end"));
  document.querySelectorAll(".group-header.drag-over").forEach(header => header.classList.remove("drag-over"));
  document.querySelectorAll(".section-sentences > li.group-item-wrap.drag-over").forEach(item => item.classList.remove("drag-over"));
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

function moveSentenceWithinContainerBefore(secIdx, groupIdx, fromIdx, beforeIdx) {
  const sentences = getSentenceContainer(secIdx, groupIdx);
  if (!Array.isArray(sentences) || sentences.length === 0) return;
  const clampedBeforeIdx = Math.max(0, Math.min(sentences.length, Number.isInteger(beforeIdx) ? beforeIdx : sentences.length));
  let targetIdx = clampedBeforeIdx;
  if (fromIdx < targetIdx) {
    targetIdx -= 1;
  }
  targetIdx = Math.max(0, Math.min(sentences.length - 1, targetIdx));
  if (fromIdx === targetIdx) return;
  moveSentenceWithinContainer(secIdx, groupIdx, fromIdx, targetIdx);
}

function createTopLevelSentenceDragRef(sentence, secIdx = null, groupIdx = null) {
  return {
    type: "sentence",
    sentence,
    sourceSecIdx: Number.isInteger(secIdx) ? secIdx : -1,
    sourceGroupIdx: Number.isInteger(groupIdx) ? groupIdx : -1,
  };
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

function isSentenceSectionItemRef(ref) {
  return !!ref && typeof ref === "object" && ref.type === "sentence" && !!ref.sentence;
}

function getSentenceSectionItemSource(ref) {
  if (!isSentenceSectionItemRef(ref)) return null;
  const sourceSecIdx = Number.parseInt(ref.sourceSecIdx, 10);
  if (!Number.isInteger(sourceSecIdx) || sourceSecIdx < 0) return null;
  const rawGroupIdx = Number.parseInt(ref.sourceGroupIdx, 10);
  return {
    secIdx: sourceSecIdx,
    groupIdx: normalizeGroupIndex(rawGroupIdx),
    sentence: ref.sentence,
  };
}

function isSameSentenceContainer(ref, secIdx, groupIdx) {
  const source = getSentenceSectionItemSource(ref);
  if (!source) return false;
  return source.secIdx === secIdx && source.groupIdx === normalizeGroupIndex(groupIdx);
}

function isSameSentenceLocation(ref, secIdx, groupIdx, sentence) {
  const source = getSentenceSectionItemSource(ref);
  if (!source) return false;
  return source.secIdx === secIdx && source.groupIdx === normalizeGroupIndex(groupIdx) && source.sentence === sentence;
}

function removeSentenceOrderItem(section, sentence) {
  section.item_order = (section.item_order || []).filter((item) => !(
    item
    && (item.type === "sentence" || item.type === "caption")
    && (item.sentence || item.caption) === sentence
  ));
}

function removeGroupOrderItem(section, groupId) {
  section.item_order = (section.item_order || []).filter((item) => !(
    item
    && item.type === "group"
    && item.group_id === groupId
  ));
}

function insertSentenceOrderItem(section, sentence, beforeRef = null) {
  const nextOrder = [...(section.item_order || [])];
  const newItem = createSentenceOrderItem(sentence);
  const beforeKey = beforeRef ? getSectionOrderItemKey(beforeRef) : "";
  const insertIndex = beforeKey
    ? nextOrder.findIndex(item => getSectionOrderItemKey(item) === beforeKey)
    : -1;
  if (insertIndex >= 0) {
    nextOrder.splice(insertIndex, 0, newItem);
  } else {
    nextOrder.push(newItem);
  }
  section.item_order = nextOrder;
}

function insertGroupOrderItem(section, groupId, beforeRef = null) {
  const nextOrder = [...(section.item_order || [])];
  const newItem = createGroupOrderItem(groupId);
  const beforeKey = beforeRef ? getSectionOrderItemKey(beforeRef) : "";
  const insertIndex = beforeKey
    ? nextOrder.findIndex(item => getSectionOrderItemKey(item) === beforeKey)
    : -1;
  if (insertIndex >= 0) {
    nextOrder.splice(insertIndex, 0, newItem);
  } else {
    nextOrder.push(newItem);
  }
  section.item_order = nextOrder;
}

function extractSentenceFromSectionLocation(secIdx, groupIdx, sentence) {
  const section = state.sections[secIdx];
  const normalizedGroupIdx = normalizeGroupIndex(groupIdx);
  const owner = normalizedGroupIdx === null ? section : section?.groups?.[normalizedGroupIdx];
  const sentences = owner?.sentences || [];
  const sourceIndex = sentences.indexOf(sentence);
  if (!section || !owner || sourceIndex < 0) return null;

  const preserveSkip = (owner.skip_sentences || []).includes(sentence);
  const preserveHidden = normalizedGroupIdx !== null && (owner.hidden_sentences || []).includes(sentence);

  sentences.splice(sourceIndex, 1);
  owner.skip_sentences = sentences.filter(item => (owner.skip_sentences || []).includes(item));

  if (normalizedGroupIdx === null) {
    removeSentenceOrderItem(section, sentence);
  } else {
    owner.hidden_sentences = sentences.filter(item => (owner.hidden_sentences || []).includes(item));
  }

  return { preserveSkip, preserveHidden };
}

function extractGroupFromSection(secIdx, groupId) {
  const section = state.sections[secIdx];
  const groups = Array.isArray(section?.groups) ? section.groups : null;
  if (!section || !groups || !groupId) return null;
  const sourceIndex = groups.findIndex(group => group?.id === groupId);
  if (sourceIndex < 0) return null;
  const [group] = groups.splice(sourceIndex, 1);
  removeGroupOrderItem(section, groupId);
  return group;
}

function insertSentenceIntoTopLevelSection(secIdx, sentence, beforeRef = null, options = {}) {
  const { preserveSkip = false } = options;
  const section = state.sections[secIdx];
  if (!section) return false;
  if (!(section.sentences || []).includes(sentence)) {
    section.sentences.push(sentence);
  }
  removeSentenceOrderItem(section, sentence);
  insertSentenceOrderItem(section, sentence, beforeRef);
  if (preserveSkip) {
    const skipped = new Set(section.skip_sentences || []);
    skipped.add(sentence);
    section.skip_sentences = (section.sentences || []).filter(item => skipped.has(item));
  }
  return true;
}

function insertSentenceIntoGroup(secIdx, groupIdx, sentence, targetIdx = null, options = {}) {
  const { preserveSkip = false, preserveHidden = false } = options;
  const group = state.sections[secIdx]?.groups?.[groupIdx];
  if (!group) return false;
  const insertAt = Number.isInteger(targetIdx)
    ? Math.max(0, Math.min(targetIdx, group.sentences.length))
    : group.sentences.length;
  group.sentences.splice(insertAt, 0, sentence);

  const skipped = new Set(group.skip_sentences || []);
  if (preserveSkip) skipped.add(sentence);
  group.skip_sentences = group.sentences.filter(item => skipped.has(item));

  const hidden = new Set(group.hidden_sentences || []);
  if (preserveHidden) hidden.add(sentence);
  group.hidden_sentences = group.sentences.filter(item => hidden.has(item));
  return true;
}

function insertGroupIntoTopLevelSection(secIdx, group, beforeRef = null) {
  const section = state.sections[secIdx];
  if (!section || !group?.id) return false;
  if (!Array.isArray(section.groups)) {
    section.groups = [];
  }
  if (!section.groups.some(existingGroup => existingGroup?.id === group.id)) {
    section.groups.push(group);
  }
  removeGroupOrderItem(section, group.id);
  insertGroupOrderItem(section, group.id, beforeRef);
  return true;
}

function createTopLevelEntryDragRef(entry, secIdx) {
  if (!entry || typeof entry !== "object") return null;
  if (entry.type === "sentence") return createTopLevelSentenceDragRef(entry.sentence, secIdx, null);
  if (entry.type === "group") return createTopLevelGroupDragRef(entry.group?.id || entry.groupId);
  return null;
}

function getFirstTopLevelEntryRef(secIdx, excludedKey = "") {
  const section = state.sections[secIdx];
  if (!section) return null;
  for (const entry of getOrderedSectionEntries(section)) {
    const ref = createTopLevelEntryDragRef(entry, secIdx);
    if (!ref) continue;
    if (excludedKey && getSectionOrderItemKey(ref) === excludedKey) continue;
    return ref;
  }
  return null;
}

function getGroupInsertIndexFromPointer(listEl, clientY) {
  const sentenceItems = [...(listEl?.children || [])].filter((item) => item instanceof HTMLElement && item.dataset?.sentence);
  for (let index = 0; index < sentenceItems.length; index += 1) {
    const rect = sentenceItems[index].getBoundingClientRect();
    if (clientY < rect.top + (rect.height / 2)) {
      return index;
    }
  }
  return sentenceItems.length;
}

function applyLocalSectionSchemaChange() {
  state.sections = normalizeSectionsData(state.sections);
  normalizeCaptionCacheForCurrentSections();
  state.filterCaptionCacheKey = getActiveSentenceFilterKey();
  refreshGridForActiveFilters();
  renderSentences();
  renderPreviewCaptionOverlay();
}

function moveSentenceToTopLevelSection(sourceSecIdx, sourceGroupIdx, sentence, targetSecIdx, beforeRef = null) {
  const extracted = extractSentenceFromSectionLocation(sourceSecIdx, sourceGroupIdx, sentence);
  if (!extracted) return;
  if (!insertSentenceIntoTopLevelSection(targetSecIdx, sentence, beforeRef, { preserveSkip: extracted.preserveSkip })) return;
  saveSectionsToStorage();
  applyLocalSectionSchemaChange();
}

function moveSentenceToGroupContainer(sourceSecIdx, sourceGroupIdx, sentence, targetSecIdx, targetGroupIdx, targetIdx = null) {
  const extracted = extractSentenceFromSectionLocation(sourceSecIdx, sourceGroupIdx, sentence);
  if (!extracted) return;
  if (!insertSentenceIntoGroup(targetSecIdx, targetGroupIdx, sentence, targetIdx, extracted)) return;
  saveSectionsToStorage();
  applyLocalSectionSchemaChange();
}

function moveSentenceToSectionHeader(sourceSecIdx, sourceGroupIdx, sentence, targetSecIdx) {
  const sourceRef = createTopLevelSentenceDragRef(sentence, sourceSecIdx, sourceGroupIdx);
  const firstRef = getFirstTopLevelEntryRef(
    targetSecIdx,
    sourceSecIdx === targetSecIdx && normalizeGroupIndex(sourceGroupIdx) === null
      ? getSectionOrderItemKey(sourceRef)
      : ""
  );

  if (sourceSecIdx === targetSecIdx && normalizeGroupIndex(sourceGroupIdx) === null) {
    if (!firstRef) return;
    moveTopLevelSectionItem(targetSecIdx, sourceRef, firstRef);
    return;
  }

  moveSentenceToTopLevelSection(sourceSecIdx, sourceGroupIdx, sentence, targetSecIdx, firstRef);
}

function moveGroupToTopLevelSection(sourceSecIdx, groupId, targetSecIdx, beforeRef = null) {
  const sourceRef = createTopLevelGroupDragRef(groupId);
  if (sourceSecIdx === targetSecIdx) {
    if (beforeRef) {
      moveTopLevelSectionItem(targetSecIdx, sourceRef, beforeRef);
    } else {
      moveTopLevelSectionItemToEnd(targetSecIdx, sourceRef);
    }
    return;
  }

  const extractedGroup = extractGroupFromSection(sourceSecIdx, groupId);
  if (!extractedGroup) return;
  if (!insertGroupIntoTopLevelSection(targetSecIdx, extractedGroup, beforeRef)) return;
  saveSectionsToStorage();
  applyLocalSectionSchemaChange();
}

function moveGroupToSectionHeader(sourceSecIdx, groupId, targetSecIdx) {
  const sourceRef = createTopLevelGroupDragRef(groupId);
  const firstRef = getFirstTopLevelEntryRef(
    targetSecIdx,
    sourceSecIdx === targetSecIdx ? getSectionOrderItemKey(sourceRef) : ""
  );

  if (sourceSecIdx === targetSecIdx) {
    if (!firstRef) return;
    moveTopLevelSectionItem(targetSecIdx, sourceRef, firstRef);
    return;
  }

  moveGroupToTopLevelSection(sourceSecIdx, groupId, targetSecIdx, firstRef);
}

function moveSentenceToGroupHeader(sourceSecIdx, sourceGroupIdx, sentence, targetSecIdx, targetGroupIdx) {
  if (sourceSecIdx === targetSecIdx && normalizeGroupIndex(sourceGroupIdx) === targetGroupIdx) {
    const sentences = getSentenceContainer(targetSecIdx, targetGroupIdx) || [];
    const sourceIndex = sentences.indexOf(sentence);
    if (sourceIndex <= 0) return;
    moveSentenceWithinContainer(targetSecIdx, targetGroupIdx, sourceIndex, 0);
    return;
  }

  moveSentenceToGroupContainer(sourceSecIdx, sourceGroupIdx, sentence, targetSecIdx, targetGroupIdx, 0);
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
  moveSentenceToGroupContainer(secIdx, null, sentence, secIdx, groupIdx, targetIdx);
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
    noteCaptionLibraryChanged();
    applyRemovedSentencesToLocalState(data.removed_captions || data.removed_sentences || [sentence]);
    await refreshCaptionsAfterSchemaChange();
    statusBar.textContent = formatCaptionFileUpdateStatus("Caption deleted", data.touched_caption_files);
  } catch (err) {
    statusBar.textContent = `Delete caption error: ${err.message}`;
    renderSentences();
  }
}

async function toggleSentenceHiddenOnExport(sentence) {
  const group = findGroupForSentence(sentence);
  if (!group) return;
  const previousHiddenSentences = [...(group.hidden_sentences || [])];
  const hidden = new Set(previousHiddenSentences);
  const willHide = !hidden.has(sentence);
  if (willHide) {
    const confirmed = confirm(
      `Remove "${sentence}" from every existing caption file in this folder?\n\n`
      + "This rewrites all caption files in the current folder and removes this entry wherever it appears. "
      + "Re-enabling TXT later will not add it back automatically."
    );
    if (!confirmed) return;
    hidden.add(sentence);
  } else {
    hidden.delete(sentence);
  }
  group.hidden_sentences = (group.sentences || []).filter(item => hidden.has(item));
  renderSentences();

  const result = await saveSectionsToStorage({
    immediate: true,
    rewriteCaptionFiles: willHide,
    savingMessage: willHide
      ? `Removing "${sentence}" from caption files...`
      : `Enabling "${sentence}" for TXT output...`,
    successMessage: willHide
      ? `Removed "${sentence}" from caption files`
      : `Enabled "${sentence}" for TXT output`,
  });
  if (result?.ok) {
    return;
  }

  group.hidden_sentences = previousHiddenSentences;
  renderSentences();
}

function toggleSentenceSkipAutoCaption(sentence) {
  const location = findSentenceLocation(sentence);
  if (!location) return;
  const owner = location.group || location.section;
  if (!owner) return;
  const skipped = new Set(owner.skip_sentences || []);
  if (skipped.has(sentence)) {
    skipped.delete(sentence);
  } else {
    skipped.add(sentence);
  }
  const sourceSentences = location.group ? (location.group.sentences || []) : (location.section?.sentences || []);
  owner.skip_sentences = sourceSentences.filter(item => skipped.has(item));
  saveSectionsToStorage();
  renderSentences();
}

function toggleGroupSkipAutoCaption(secIdx, groupIdx) {
  const group = state.sections[secIdx]?.groups?.[groupIdx];
  if (!group) return;
  group.skip_auto_caption = !group.skip_auto_caption;
  saveSectionsToStorage();
  renderSentences();
}

function toggleSectionSkipAutoCaption(secIdx) {
  const section = state.sections[secIdx];
  if (!section) return;
  section.skip_auto_caption = !section.skip_auto_caption;
  saveSectionsToStorage();
  renderSentences();
}

function createAutoCaptionSkipButton(skipped, activeTitle, inactiveTitle, onClick) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "auto-caption-skip-btn";

  const icon = document.createElement("span");
  icon.className = "auto-caption-skip-icon";
  icon.setAttribute("aria-hidden", "true");

  const iconText = document.createElement("span");
  iconText.className = "auto-caption-skip-icon-text";
  iconText.textContent = "AI";

  const iconBan = document.createElement("span");
  iconBan.className = "auto-caption-skip-icon-ban";

  icon.appendChild(iconText);
  icon.appendChild(iconBan);
  button.appendChild(icon);
  syncAutoCaptionSkipButtonState(button, skipped, activeTitle, inactiveTitle);

  button.addEventListener("click", (e) => {
    e.stopPropagation();
    onClick(e);
  });
  return button;
}

function syncAutoCaptionSkipButtonState(button, skipped, activeTitle, inactiveTitle) {
  if (!button) return;
  const active = !!skipped;
  const title = active ? activeTitle : inactiveTitle;
  button.classList.toggle("active", active);
  button.title = title;
  button.setAttribute("aria-label", title);
  button.setAttribute("aria-pressed", active ? "true" : "false");
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
    noteCaptionLibraryChanged();
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
    statusBar.textContent = formatCaptionFileUpdateStatus("Caption renamed", data.touched_caption_files);
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
    noteCaptionLibraryChanged();
    renderSentences();
    statusBar.textContent = formatCaptionFileUpdateStatus("Section renamed", data.touched_caption_files);
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
    noteCaptionLibraryChanged();
    applyRemovedSentencesToLocalState(data.removed_sentences || []);
    await refreshCaptionsAfterSchemaChange();
    statusBar.textContent = formatCaptionFileUpdateStatus("Section deleted", data.touched_caption_files);
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
    noteCaptionLibraryChanged();
    applyRemovedSentencesToLocalState(data.removed_sentences || []);
    await refreshCaptionsAfterSchemaChange();
    statusBar.textContent = formatCaptionFileUpdateStatus("Group deleted", data.touched_caption_files);
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

function getFolderSentenceCountState() {
  const counts = new Map();
  let hasMissingCaptions = false;
  for (const image of state.images) {
    const caption = state.captionCache[image.path];
    if (!caption) {
      hasMissingCaptions = true;
      continue;
    }
    for (const sentence of (caption.enabled_sentences || [])) {
      counts.set(sentence, (counts.get(sentence) || 0) + 1);
    }
  }
  return { counts, hasMissingCaptions };
}

function queueFolderSentenceCountLoad() {
  if (state.captionLibraryLoadingPromise) return;
  if (!state.images.some(img => !state.captionCache[img.path])) return;
  if (typeof ensureCaptionCacheLoadedForCurrentFolder !== "function") return;

  state.captionLibraryLoadingPromise = Promise.resolve(ensureCaptionCacheLoadedForCurrentFolder())
    .then(() => {
      scheduleUiRender({ sentences: true });
    })
    .catch((err) => {
      console.error("Failed to load folder caption counts:", err);
    })
    .finally(() => {
      state.captionLibraryLoadingPromise = null;
    });
}

function createSentenceListItem(sentence, selectedPaths, options = {}) {
  const { isExclusive = false, allowSuppressToggle = false } = options;
  const secIdx = Number.isInteger(options.secIdx) ? options.secIdx : -1;
  const groupIdx = normalizeGroupIndex(options.groupIdx);
  const sentenceIdx = Number.isInteger(options.sentenceIdx) ? options.sentenceIdx : -1;
  const topLevelMix = !!options.topLevelMix;
  const folderSentenceCounts = options.folderSentenceCounts instanceof Map ? options.folderSentenceCounts : null;
  const folderCountsPending = !!options.folderCountsPending;
  const sentenceDragRef = createTopLevelSentenceDragRef(sentence, secIdx, groupIdx);
  const topLevelRef = topLevelMix ? sentenceDragRef : null;
  const canRefreshSentence = groupIdx === null;
  const isGroupSentence = groupIdx !== null;
  const autoCaptionSkipped = isSentenceSkippedForAutoCaption(sentence);
  const filterMode = getSentenceFilterMode(sentence);
  const isFilterActive = filterMode !== "off";
  const folderCount = folderSentenceCounts?.get(sentence) || 0;
  const { isChecked, isPartial } = getSentenceSelectionState(sentence, selectedPaths);
  const li = document.createElement("li");
  li.dataset.sentence = sentence;
  sentenceListElements.set(sentence, li);

  const dragHandle = document.createElement("span");
  dragHandle.className = "drag-handle";
  dragHandle.title = "Drag to reorder caption";
  dragHandle.draggable = true;
  dragHandle.addEventListener("click", (e) => e.stopPropagation());
  dragHandle.addEventListener("mousedown", (e) => e.stopPropagation());
  dragHandle.addEventListener("dragstart", (e) => {
    e.stopPropagation();
    setDragPayload(e, SECTION_ITEM_DRAG_TYPE, { secIdx, item: sentenceDragRef });
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
    clearCaptionDragIndicators();
  });
  const checkBox = document.createElement("div");
  checkBox.className = `check-box${isExclusive ? " radio" : ""}`;
  if (isChecked) checkBox.classList.add("checked");
  else if (isPartial) checkBox.classList.add("partial");

  const textSpan = document.createElement("span");
  textSpan.className = "sentence-text editable";
  if (isPartial) textSpan.classList.add("partial");
  if (autoCaptionSkipped) textSpan.classList.add("skip-auto-caption");
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

  const folderCountBadge = document.createElement("span");
  folderCountBadge.className = `sentence-folder-count${folderCountsPending ? " loading" : ""}`;
  folderCountBadge.textContent = folderCountsPending ? "\u2026" : String(folderCount);
  folderCountBadge.title = folderCountsPending
    ? "Loading caption count for this folder"
    : `${folderCount} file${folderCount === 1 ? "" : "s"} in this folder include this caption`;
  folderCountBadge.setAttribute("aria-label", folderCountsPending
    ? "Loading caption count for this folder"
    : `Caption appears in ${folderCount} file${folderCount === 1 ? "" : "s"} in this folder`);

  const autoCaptionSkipBtn = createAutoCaptionSkipButton(
    autoCaptionSkipped,
    "This caption will be skipped during Auto Caption",
    "Check this caption during Auto Caption",
    () => {
      toggleSentenceSkipAutoCaption(sentence);
    }
  );

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
    li.appendChild(folderCountBadge);
    if (refreshBtn) li.appendChild(refreshBtn);
    li.appendChild(filterToggleBtn);
    li.appendChild(autoCaptionSkipBtn);
    li.appendChild(exportToggleBtn);
    li.appendChild(rmBtn);
  } else {
    li.appendChild(dragHandle);
    li.appendChild(checkBox);
    li.appendChild(textSpan);
    li.appendChild(folderCountBadge);
    li.appendChild(filterToggleBtn);
    if (refreshBtn) li.appendChild(refreshBtn);
    li.appendChild(autoCaptionSkipBtn);
    li.appendChild(rmBtn);
  }
  li.addEventListener("dragover", (e) => {
    if (topLevelMix) {
      const payload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
      if (!payload || !isValidTopLevelDragRef(payload.item)) return;
      if (payload.item.type === "group" && getSectionOrderItemKey(payload.item) === getSectionOrderItemKey(topLevelRef)) return;
      if (isSameSentenceLocation(payload.item, secIdx, null, sentence)) return;
      e.preventDefault();
      e.stopPropagation();
      e.dataTransfer.dropEffect = "move";
      li.classList.add("drag-over");
      return;
    }
    const topLevelPayload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (
      topLevelPayload
      && isSentenceSectionItemRef(topLevelPayload.item)
      && !isSameSentenceContainer(topLevelPayload.item, secIdx, groupIdx)
    ) {
      e.preventDefault();
      e.stopPropagation();
      e.dataTransfer.dropEffect = "move";
      li.classList.add("drag-over");
      return;
    }
    const payload = getDragPayload(e, SENTENCE_DRAG_TYPE);
    if (!payload) return;
    if (payload.secIdx !== secIdx || normalizeGroupIndex(payload.groupIdx) !== groupIdx) return;
    if (payload.sentenceIdx === sentenceIdx) return;
    e.preventDefault();
    e.stopPropagation();
    e.dataTransfer.dropEffect = "move";
    li.classList.add("drag-over");
  });
  li.addEventListener("dragleave", () => {
    li.classList.remove("drag-over");
  });
  li.addEventListener("drop", (e) => {
    if (topLevelMix) {
      const payload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
      if (!payload || !isValidTopLevelDragRef(payload.item)) return;
      e.preventDefault();
      e.stopPropagation();
      li.classList.remove("drag-over");
      if (payload.item.type === "group") {
        moveGroupToTopLevelSection(payload.secIdx, payload.item.group_id, secIdx, topLevelRef);
        return;
      }
      const source = getSentenceSectionItemSource(payload.item);
      if (!source) return;
      if (source.secIdx === secIdx && source.groupIdx === null) {
        moveTopLevelSectionItem(secIdx, payload.item, topLevelRef);
        return;
      }
      moveSentenceToTopLevelSection(source.secIdx, source.groupIdx, source.sentence, secIdx, topLevelRef);
      return;
    }
    const topLevelPayload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (
      topLevelPayload
      && isSentenceSectionItemRef(topLevelPayload.item)
      && !isSameSentenceContainer(topLevelPayload.item, secIdx, groupIdx)
    ) {
      const source = getSentenceSectionItemSource(topLevelPayload.item);
      if (!source) return;
      e.preventDefault();
      e.stopPropagation();
      li.classList.remove("drag-over");
      moveSentenceToGroupContainer(source.secIdx, source.groupIdx, source.sentence, secIdx, groupIdx, sentenceIdx);
      return;
    }
    const payload = getDragPayload(e, SENTENCE_DRAG_TYPE);
    if (!payload) return;
    const payloadGroupIdx = normalizeGroupIndex(payload.groupIdx);
    if (payload.secIdx !== secIdx || payloadGroupIdx !== groupIdx) return;
    e.preventDefault();
    e.stopPropagation();
    li.classList.remove("drag-over");
    moveSentenceWithinContainerBefore(secIdx, groupIdx, payload.sentenceIdx, sentenceIdx);
  });
  li.addEventListener("click", (e) => {
    e.stopPropagation();
    toggleSentence(sentence, isChecked, isPartial);
  });
  return li;
}

function createGroupListItem(section, group, secIdx, groupIdx, selectedPaths, options = {}) {
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
    clearCaptionDragIndicators();
  });
  groupHeader.appendChild(groupDragHandle);

  const groupCollapseBtn = document.createElement("button");
  groupCollapseBtn.type = "button";
  groupCollapseBtn.className = "collapse-btn";
  groupCollapseBtn.textContent = isGroupCollapsed(group) ? "\u25B8" : "\u25BE";
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

  const groupSkipBtn = createAutoCaptionSkipButton(
    isGroupSkippedForAutoCaption(group),
    "This group will be skipped during Auto Caption",
    "Check this group during Auto Caption",
    () => {
      toggleGroupSkipAutoCaption(secIdx, groupIdx);
    }
  );
  groupHeader.appendChild(groupSkipBtn);

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
    if (!payload || !isSentenceSectionItemRef(payload.item)) return;
    const source = getSentenceSectionItemSource(payload.item);
    if (!source) return;
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
    if (!payload || !isSentenceSectionItemRef(payload.item)) return;
    const source = getSentenceSectionItemSource(payload.item);
    if (!source) return;
    e.preventDefault();
    e.stopPropagation();
    groupHeader.classList.remove("drag-over");
    moveSentenceToGroupHeader(source.secIdx, source.groupIdx, source.sentence, secIdx, groupIdx);
  });

  groupBlock.appendChild(groupHeader);

  const groupBody = document.createElement("div");
  groupBody.className = `group-body${isGroupCollapsed(group) ? " collapsed" : ""}`;
  const groupSentences = document.createElement("ul");
  groupSentences.className = "section-sentences group-sentences";
  groupSentences.addEventListener("dragover", (e) => {
    const topLevelPayload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (topLevelPayload && isSentenceSectionItemRef(topLevelPayload.item) && !isSameSentenceContainer(topLevelPayload.item, secIdx, groupIdx)) {
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
    e.stopPropagation();
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
    if (topLevelPayload && isSentenceSectionItemRef(topLevelPayload.item) && !isSameSentenceContainer(topLevelPayload.item, secIdx, groupIdx)) {
      const source = getSentenceSectionItemSource(topLevelPayload.item);
      if (!source) return;
      e.preventDefault();
      e.stopPropagation();
      groupSentences.classList.remove("drag-over-end");
      moveSentenceToGroupContainer(
        source.secIdx,
        source.groupIdx,
        source.sentence,
        secIdx,
        groupIdx,
        getGroupInsertIndexFromPointer(groupSentences, e.clientY)
      );
      return;
    }
    const payload = getDragPayload(e, SENTENCE_DRAG_TYPE);
    if (!payload) return;
    if (payload.secIdx !== secIdx || normalizeGroupIndex(payload.groupIdx) !== groupIdx) return;
    e.preventDefault();
    e.stopPropagation();
    groupSentences.classList.remove("drag-over-end");
    const targetIdx = getGroupInsertIndexFromPointer(groupSentences, e.clientY);
    moveSentenceWithinContainerBefore(secIdx, groupIdx, payload.sentenceIdx, targetIdx);
  });
  const visibleSentences = Array.isArray(options.visibleSentences)
    ? options.visibleSentences
    : (group.sentences || []).map((sentence, sentenceIdx) => ({ sentence, sentenceIdx }));
  for (const { sentence, sentenceIdx } of visibleSentences) {
    groupSentences.appendChild(createSentenceListItem(sentence, selectedPaths, {
      isExclusive: true,
      allowSuppressToggle: true,
      secIdx,
      groupIdx,
      sentenceIdx,
      folderSentenceCounts: options.folderSentenceCounts,
      folderCountsPending: options.folderCountsPending,
    }));
  }
  groupBody.appendChild(groupSentences);
  groupBody.appendChild(createAddSentenceRow(secIdx, groupIdx, "Add caption to group..."));
  groupBlock.appendChild(groupBody);
  wrapper.appendChild(groupBlock);

  wrapper.addEventListener("dragover", (e) => {
    const payload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (!payload || !isValidTopLevelDragRef(payload.item)) return;
    if (payload.item.type === "group" && getSectionOrderItemKey(payload.item) === getSectionOrderItemKey(topLevelRef)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    wrapper.classList.add("drag-over");
  });
  wrapper.addEventListener("dragleave", () => {
    wrapper.classList.remove("drag-over");
  });
  wrapper.addEventListener("drop", (e) => {
    const payload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
    if (!payload || !isValidTopLevelDragRef(payload.item)) return;
    e.preventDefault();
    wrapper.classList.remove("drag-over");
    if (payload.item.type === "group") {
      moveGroupToTopLevelSection(payload.secIdx, payload.item.group_id, secIdx, topLevelRef);
      return;
    }
    const source = getSentenceSectionItemSource(payload.item);
    if (!source) return;
    if (source.secIdx === secIdx && source.groupIdx === null) {
      moveTopLevelSectionItem(secIdx, payload.item, topLevelRef);
      return;
    }
    moveSentenceToTopLevelSection(source.secIdx, source.groupIdx, source.sentence, secIdx, topLevelRef);
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

function normalizeCaptionSearchQuery(value) {
  return String(value || "").trim().toLowerCase();
}

function captionMatchesSearchQuery(sentence, normalizedQuery) {
  if (!normalizedQuery) return true;
  return String(sentence || "").toLowerCase().includes(normalizedQuery);
}

function sentenceMatchesCaptionLibraryFilters(sentence, normalizedQuery) {
  if (state.captionSkipFilterActive && !isSentenceSkippedForAutoCaption(sentence)) {
    return false;
  }
  return captionMatchesSearchQuery(sentence, normalizedQuery);
}

function groupMatchesCaptionLibraryFilters(group, normalizedQuery) {
  if (state.captionSkipFilterActive && !isGroupSkippedForAutoCaption(group)) {
    return false;
  }
  return captionMatchesSearchQuery(group?.name || "(Group)", normalizedQuery);
}

function sectionMatchesCaptionLibraryFilters(section, normalizedQuery) {
  if (state.captionSkipFilterActive && !isSectionSkippedForAutoCaption(section)) {
    return false;
  }
  return captionMatchesSearchQuery(section?.name || "(General)", normalizedQuery);
}

function getVisibleGroupSentenceMatches(group, normalizedQuery) {
  const groupSentences = Array.isArray(group?.sentences) ? group.sentences : [];
  const matches = [];
  groupSentences.forEach((sentence, sentenceIdx) => {
    if (sentenceMatchesCaptionLibraryFilters(sentence, normalizedQuery)) {
      matches.push({ sentence, sentenceIdx });
    }
  });
  return matches;
}

function getVisibleSectionEntries(section, normalizedQuery) {
  const orderedEntries = getOrderedSectionEntries(section);
  const sectionMatched = sectionMatchesCaptionLibraryFilters(section, normalizedQuery);
  const skipFilterActive = !!state.captionSkipFilterActive;
  if (!normalizedQuery && !skipFilterActive) {
    return {
      sectionMatched,
      entries: orderedEntries,
    };
  }

  if (sectionMatched) {
    return {
      sectionMatched,
      entries: orderedEntries,
    };
  }

  const visibleEntries = [];
  for (const entry of orderedEntries) {
    if (entry.type === "sentence") {
      if (sentenceMatchesCaptionLibraryFilters(entry.sentence, normalizedQuery)) {
        visibleEntries.push(entry);
      }
      continue;
    }

    if (groupMatchesCaptionLibraryFilters(entry.group, normalizedQuery)) {
      visibleEntries.push(entry);
      continue;
    }

    const visibleSentences = getVisibleGroupSentenceMatches(entry.group, normalizedQuery);
    if (visibleSentences.length > 0) {
      visibleEntries.push({ ...entry, visibleSentences });
    }
  }
  return {
    sectionMatched,
    entries: visibleEntries,
  };
}

function createCaptionSearchEmptyState(searchQuery) {
  const emptyState = document.createElement("div");
  emptyState.className = "section-empty-state";
  const trimmedQuery = String(searchQuery || "").trim();
  if (trimmedQuery && state.captionSkipFilterActive) {
    emptyState.textContent = `No disabled AI captions match "${trimmedQuery}".`;
  } else if (trimmedQuery) {
    emptyState.textContent = `No captions match "${trimmedQuery}".`;
  } else if (state.captionSkipFilterActive) {
    emptyState.textContent = "No captions are disabled for Auto Caption.";
  } else {
    emptyState.textContent = "No captions available.";
  }
  return emptyState;
}

function syncCaptionSearchInputState() {
  if (!captionSearchInput) return;
  const searchQuery = String(state.captionSearchQuery || "");
  const normalizedQuery = normalizeCaptionSearchQuery(searchQuery);
  if (captionSearchInput.value !== searchQuery) {
    captionSearchInput.value = searchQuery;
  }
  captionSearchInput.classList.toggle("active", !!normalizedQuery);
  captionSearchInput.title = normalizedQuery
    ? `Filtering captions for "${searchQuery.trim()}"`
    : "Filter captions in the library";
}

function syncCaptionSkipFilterButtonState() {
  if (!captionSkipFilterBtn) return;
  syncAutoCaptionSkipButtonState(
    captionSkipFilterBtn,
    state.captionSkipFilterActive,
    "Showing only captions disabled for Auto Caption",
    "Show only captions disabled for Auto Caption",
  );
}

function updateCaptionSearchQuery(value) {
  const nextQuery = String(value || "");
  if (nextQuery === state.captionSearchQuery) {
    syncCaptionSearchInputState();
    return;
  }
  state.captionSearchQuery = nextQuery;
  syncCaptionSearchInputState();
  renderSentences({ includePreview: false });
}

function handleCaptionSearchInput(event) {
  updateCaptionSearchQuery(event?.target?.value || "");
}

function toggleCaptionSkipFilter() {
  state.captionSkipFilterActive = !state.captionSkipFilterActive;
  syncCaptionSkipFilterButtonState();
  renderSentences({ includePreview: false });
  if (state.captionSkipFilterActive) {
    statusBar.textContent = "Showing only captions disabled for Auto Caption";
  } else if (normalizeCaptionSearchQuery(state.captionSearchQuery)) {
    statusBar.textContent = `Filtering captions for "${String(state.captionSearchQuery || "").trim()}"`;
  } else {
    statusBar.textContent = "Showing all captions";
  }
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
  const normalizedCaptionSearchQuery = normalizeCaptionSearchQuery(state.captionSearchQuery);
  let hasVisibleEntries = false;

  state.sections = normalizeSectionsData(state.sections);
  const { counts: folderSentenceCounts, hasMissingCaptions } = getFolderSentenceCountState();
  if (hasMissingCaptions) {
    queueFolderSentenceCountLoad();
  }
  syncCaptionSearchInputState();
  syncCaptionSkipFilterButtonState();

  state.sections.forEach((section, secIdx) => {
    const { sectionMatched, entries: visibleEntries } = getVisibleSectionEntries(section, normalizedCaptionSearchQuery);
    if ((normalizedCaptionSearchQuery || state.captionSkipFilterActive) && !sectionMatched && visibleEntries.length === 0) {
      return;
    }
    hasVisibleEntries = true;

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
    collapseBtn.textContent = isSectionCollapsed(section) ? "\u25B8" : "\u25BE";
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

    const sectionSkipBtn = createAutoCaptionSkipButton(
      isSectionSkippedForAutoCaption(section),
      "This section will be skipped during Auto Caption",
      "Check this section during Auto Caption",
      () => {
        toggleSectionSkipAutoCaption(secIdx);
      }
    );
    header.appendChild(sectionSkipBtn);

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
      const topLevelPayload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
      if (topLevelPayload && isValidTopLevelDragRef(topLevelPayload.item)) {
        e.preventDefault();
        e.stopPropagation();
        e.dataTransfer.dropEffect = "move";
        header.classList.add("drag-over");
        return;
      }
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
      const topLevelPayload = getDragPayload(e, SECTION_ITEM_DRAG_TYPE);
      if (topLevelPayload && isValidTopLevelDragRef(topLevelPayload.item)) {
        e.preventDefault();
        e.stopPropagation();
        header.classList.remove("drag-over");
        if (topLevelPayload.item.type === "group") {
          moveGroupToSectionHeader(topLevelPayload.secIdx, topLevelPayload.item.group_id, secIdx);
          return;
        }
        const source = getSentenceSectionItemSource(topLevelPayload.item);
        if (!source) return;
        moveSentenceToSectionHeader(source.secIdx, source.groupIdx, source.sentence, secIdx);
        return;
      }
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
      if (!payload || !isValidTopLevelDragRef(payload.item)) return;
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
      if (!payload || !isValidTopLevelDragRef(payload.item)) return;
      if (e.target !== sectionSentences) return;
      e.preventDefault();
      sectionSentences.classList.remove("drag-over-end");
      if (payload.item.type === "group") {
        moveGroupToTopLevelSection(payload.secIdx, payload.item.group_id, secIdx);
        return;
      }
      const source = getSentenceSectionItemSource(payload.item);
      if (!source) return;
      if (source.secIdx === secIdx && source.groupIdx === null) {
        moveTopLevelSectionItemToEnd(secIdx, payload.item);
        return;
      }
      moveSentenceToTopLevelSection(source.secIdx, source.groupIdx, source.sentence, secIdx);
    });
    for (const entry of visibleEntries) {
      if (entry.type === "sentence") {
        sectionSentences.appendChild(createSentenceListItem(entry.sentence, selectedPaths, {
          secIdx,
          topLevelMix: true,
          folderSentenceCounts,
          folderCountsPending: hasMissingCaptions,
        }));
      } else if (entry.type === "group") {
        sectionSentences.appendChild(createGroupListItem(section, entry.group, secIdx, entry.groupIdx, selectedPaths, {
          folderSentenceCounts,
          folderCountsPending: hasMissingCaptions,
          visibleSentences: entry.visibleSentences,
        }));
      }
    }
    body.appendChild(sectionSentences);

    body.appendChild(createAddSentenceRow(secIdx, null, "Add caption to section..."));
    sectionBlock.appendChild(body);
    sectionContainer.appendChild(sectionBlock);
  });

  if (normalizedCaptionSearchQuery && !hasVisibleEntries) {
    sectionContainer.appendChild(createCaptionSearchEmptyState(state.captionSearchQuery));
  }

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

function getOllamaAnswerLogSuffix(event) {
  if (!event?.answer_incomplete) return "";
  return event.answer_done_reason ? ` [partial: ${event.answer_done_reason}]` : " [partial]";
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
  if (state.uploading || state.cloning || state.moving || state.extractingFrame) {
    showErrorToast("Finish the current operation before starting auto captioning.");
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
        max_output_tokens: state.ollamaMaxOutputTokens,
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
        if (event.max_output_tokens) {
          appendModelLog(`Max output tokens: ${event.max_output_tokens}`, "log-dim");
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
        if (event.skipped) {
          const reason = event.skip_reason ? ` (${event.skip_reason})` : "";
          appendModelLog(`[${getFileLabel(event.path)}] ${event.index}/${event.total} ${currentCaption} -> SKIP${reason}`, "log-warn");
        } else {
          const verdict = event.enabled ? "YES" : "NO";
          appendModelLog(`[${getFileLabel(event.path)}] ${event.index}/${event.total} ${currentCaption} -> ${verdict} | ${event.answer || verdict}${getOllamaAnswerLogSuffix(event)}`, event.enabled ? "log-ok" : "log-dim");
        }
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
        if (event.skipped) {
          const reason = event.skip_reason ? ` (${event.skip_reason})` : "";
          appendModelLog(
            `[${getFileLabel(event.path)}] ${event.index}/${event.total} ${event.group_name || "Group"} -> SKIP${reason}`,
            "log-warn"
          );
        } else {
          appendModelLog(
            `[${getFileLabel(event.path)}] ${event.index}/${event.total} ${event.group_name || "Group"} -> ${selectedCaption}${event.selected_hidden ? " (ignored in txt)" : ""} | ${event.answer || ""}${getOllamaAnswerLogSuffix(event)}`,
            (event.selected_caption || event.selected_sentence) ? "log-ok" : "log-warn"
          );
        }
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
        state.captionDraftPaths.delete(event.path);
        cache.enabled_sentences = orderEnabledSentences(event.enabled_captions || event.enabled_sentences || cache.enabled_sentences || []);
        cache.free_text = event.free_text || cache.free_text || "";
        if ((event.added_lines || []).length > 0) {
          appendModelLog(`[${getFileLabel(event.path)}] Added free text: ${(event.added_lines || []).join(" | ")}${getOllamaAnswerLogSuffix(event)}`, "log-ok");
        } else {
          appendModelLog(`[${getFileLabel(event.path)}] Free text output: ${event.answer || "NONE"}${getOllamaAnswerLogSuffix(event)}`, "log-warn");
        }
        if (state.selectedPaths.size === 1 && state.selectedPaths.has(event.path)) {
          freeText.value = cache.free_text || "";
          syncFreeTextHighlightState();
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
        state.captionDraftPaths.delete(event.path);
        state.captionCache[event.path] = normalizeCaptionCacheEntry(event);
        refreshGridForActiveFilters();
        const hasContent = hasEffectiveCaptionContent(state.captionCache[event.path]);
        markCaptionIndicator(event.path, !!hasContent);
        appendModelLog(getAutoCaptionCompleteLog(scope, event), "log-ok");
        if (state.selectedPaths.size === 1 && state.selectedPaths.has(event.path)) {
          freeText.value = event.free_text || "";
          syncFreeTextHighlightState();
        }
        if (autoPreviewEnabled && hasContent) {
          queuePromptPreviewFromCurrentCaption(event.path)
            .then((queued) => {
              appendModelLog(`[${getFileLabel(event.path)}] Auto Preview queued`, "log-dim");
              if (state.previewPath === event.path && state.previewMediaType === "image") {
                state.promptPreview.sourcePath = event.path;
                applyPromptPreviewSnapshot(event.path, queued, {
                  autoDisplayLatest: false,
                  allowCompletedAutodisplay: false,
                  allowFileChangeAutodisplay: false,
                });
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

function buildRegionDescriptionCaptionContext(path) {
  const cache = ensureCaptionCache(path);
  const enabledSentences = orderEnabledSentences(Array.isArray(cache.enabled_sentences) ? cache.enabled_sentences : []);
  const currentFreeText = state.selectedPaths.size === 1 && state.selectedPaths.has(path)
    ? String(freeText.value || "")
    : String(cache.free_text || "");
  const parts = [];
  if (enabledSentences.length > 0) {
    parts.push(enabledSentences.join("\n"));
  }
  if (currentFreeText.trim()) {
    parts.push(currentFreeText.trim());
  }
  return {
    enabledSentences,
    currentFreeText,
    captionText: parts.join("\n\n"),
  };
}

async function toggleFreeTextRegionPicker() {
  if (state.aiRegionPicker.loading) {
    stopAiRegionPicker();
    return;
  }
  if (state.aiRegionPicker.active) {
    stopAiRegionPicker();
    return;
  }
  if (!state.ollamaModel.trim()) {
    statusBar.textContent = "Configure an Ollama model first";
    openSettingsModal();
    return;
  }
  beginAiRegionPicker();
}

async function describeFreeTextRegion() {
  if (!state.aiRegionPicker.active || state.aiRegionPicker.loading) {
    return;
  }
  if (state.selectedPaths.size !== 1 || !state.previewPath || !isImageMediaPath(state.previewPath)) {
    stopAiRegionPicker({ keepStatus: true });
    statusBar.textContent = "Select a single image first";
    return;
  }
  if (!state.cropDraft) {
    statusBar.textContent = "Draw a region first";
    return;
  }
  const path = state.previewPath;
  const controller = new AbortController();
  const { enabledSentences, currentFreeText, captionText } = buildRegionDescriptionCaptionContext(path);
  state.aiRegionPicker.loading = true;
  state.aiRegionPicker.abortController = controller;
  renderCropOverlay();
  updateActionButtons();
  statusBar.textContent = "Describing selected region...";

  try {
    const resp = await fetch("/api/auto-caption/describe-region", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal: controller.signal,
      body: JSON.stringify({
        image_path: path,
        crop: buildCropPayload(state.cropDraft),
        model: state.ollamaModel.trim(),
        caption_text: captionText,
        free_text: currentFreeText,
        enabled_captions: enabledSentences,
        free_text_prompt_template: state.ollamaFreeTextPromptTemplate,
        timeout_seconds: state.ollamaTimeoutSeconds,
        max_output_tokens: state.ollamaMaxOutputTokens,
      }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to describe selected region");
    }
    const nextFreeText = String(data.free_text || currentFreeText);
    if (nextFreeText !== currentFreeText) {
      setFreeTextDraftForPath(path, nextFreeText, {
        syncEditor: state.selectedPaths.size === 1 && state.selectedPaths.has(path),
      });
    }
    const addedCount = Array.isArray(data.added_lines) ? data.added_lines.length : 0;
    statusBar.textContent = addedCount > 0
      ? `Added ${addedCount} region detail${addedCount === 1 ? "" : "s"}`
      : "No new region details found";
  } catch (err) {
    if (err?.name === "AbortError") {
      statusBar.textContent = "Region description cancelled";
    } else {
      const message = err?.message || "Failed to describe selected region";
      statusBar.textContent = message;
      showErrorToast(message);
    }
  } finally {
    stopAiRegionPicker({ abort: false, keepStatus: true });
  }
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

async function saveMetadataForSelection(options = {}) {
  const { quietNoChanges = false } = options;
  if (state.metadataSavePromise) {
    return state.metadataSavePromise;
  }

  const selectedPaths = [...state.selectedPaths];
  if (!selectedPaths.length) return true;

  const runSave = async () => {
    let singlePath = null;
    let singleMetadata = null;
    let batchChanges = null;

    try {
      if (selectedPaths.length === 1) {
        singlePath = selectedPaths[0];
        singleMetadata = buildSingleMetadataPayload();
      } else {
        batchChanges = buildBatchMetadataChanges();
      }
    } catch (err) {
      const message = err?.message || "Failed to save metadata";
      statusBar.textContent = `Metadata error: ${message}`;
      showErrorToast(`Metadata error: ${message}`);
      return false;
    }

    if (selectedPaths.length > 1 && Object.keys(batchChanges || {}).length === 0) {
      renderMetadataEditor();
      if (!quietNoChanges) {
        statusBar.textContent = "Metadata unchanged";
      }
      return true;
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
        return true;
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
          return false;
        }

        statusBar.textContent = `Applied metadata to ${updatedCount} media file${updatedCount === 1 ? "" : "s"}`;
        return true;
      }
    } catch (err) {
      const message = err?.message || "Failed to save metadata";
      statusBar.textContent = `Metadata error: ${message}`;
      showErrorToast(`Metadata error: ${message}`);
      return false;
    } finally {
      state.metadataSaving = false;
      renderMetadataEditor();
      updateMultiInfo();
    }
  };

  state.metadataSavePromise = runSave();
  try {
    return await state.metadataSavePromise;
  } finally {
    state.metadataSavePromise = null;
  }
}

function showSentenceContextMenu() { /* removed - using inline remove buttons now */ }

function showSectionContextMenu() { /* removed - using inline controls now */ }

// ===== FREE TEXT =====
const FREE_TEXT_COLOR_SENTINEL = "rgb(1, 2, 3)";
const FREE_TEXT_NON_COLOR_KEYWORDS = new Set([
  "transparent",
  "currentcolor",
  "inherit",
  "initial",
  "unset",
  "revert",
  "revert-layer",
]);
const freeTextColorProbeContext = document.createElement("canvas").getContext("2d");
const freeTextResolvedColorCache = new Map();
const freeTextColorSupportProbe = document.createElement("span");

function escapeFreeTextHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;");
}

function parseCssColorToRgb(value) {
  const normalized = String(value || "").trim();
  if (!normalized) return null;

  const hexMatch = normalized.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
  if (hexMatch) {
    const hex = hexMatch[1];
    if (hex.length === 3) {
      return {
        r: Number.parseInt(`${hex[0]}${hex[0]}`, 16),
        g: Number.parseInt(`${hex[1]}${hex[1]}`, 16),
        b: Number.parseInt(`${hex[2]}${hex[2]}`, 16),
      };
    }
    return {
      r: Number.parseInt(hex.slice(0, 2), 16),
      g: Number.parseInt(hex.slice(2, 4), 16),
      b: Number.parseInt(hex.slice(4, 6), 16),
    };
  }

  const rgbMatch = normalized.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  if (rgbMatch) {
    return {
      r: Math.max(0, Math.min(255, Number.parseInt(rgbMatch[1], 10) || 0)),
      g: Math.max(0, Math.min(255, Number.parseInt(rgbMatch[2], 10) || 0)),
      b: Math.max(0, Math.min(255, Number.parseInt(rgbMatch[3], 10) || 0)),
    };
  }

  return null;
}

function getRelativeLuminance({ r, g, b }) {
  const toLinear = (channel) => {
    const value = Math.max(0, Math.min(255, Number(channel || 0))) / 255;
    return value <= 0.03928 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
  };
  return (0.2126 * toLinear(r)) + (0.7152 * toLinear(g)) + (0.0722 * toLinear(b));
}

function getContrastRatio(leftRgb, rightRgb) {
  if (!leftRgb || !rightRgb) return Number.POSITIVE_INFINITY;
  const leftLum = getRelativeLuminance(leftRgb);
  const rightLum = getRelativeLuminance(rightRgb);
  const lighter = Math.max(leftLum, rightLum);
  const darker = Math.min(leftLum, rightLum);
  return (lighter + 0.05) / (darker + 0.05);
}

function getFreeTextEditorBackgroundRgb() {
  const background = getComputedStyle(freeTextEditor || freeText).backgroundColor;
  return parseCssColorToRgb(background);
}

function isSupportedFreeTextCssColor(value) {
  const normalized = String(value || "").trim();
  if (!normalized) return false;
  if (typeof CSS !== "undefined" && typeof CSS.supports === "function") {
    return CSS.supports("color", normalized);
  }
  freeTextColorSupportProbe.style.color = "";
  freeTextColorSupportProbe.style.color = normalized;
  return freeTextColorSupportProbe.style.color !== "";
}

function resolveFreeTextColorToken(word) {
  const normalized = String(word || "").trim().toLowerCase().replace(/\s+/g, " ");
  if (!normalized || !/^[a-z -]+$/.test(normalized)) {
    return null;
  }
  const normalizedCompact = normalized.replace(/[\s-]+/g, "");
  if (!normalizedCompact || FREE_TEXT_NON_COLOR_KEYWORDS.has(normalizedCompact)) {
    return null;
  }
  if (!isSupportedFreeTextCssColor(normalizedCompact)) {
    return null;
  }
  if (freeTextResolvedColorCache.has(normalizedCompact)) {
    return freeTextResolvedColorCache.get(normalizedCompact);
  }
  if (!freeTextColorProbeContext) {
    freeTextResolvedColorCache.set(normalizedCompact, null);
    return null;
  }

  freeTextColorProbeContext.fillStyle = FREE_TEXT_COLOR_SENTINEL;
  freeTextColorProbeContext.fillStyle = normalizedCompact;
  const resolved = freeTextColorProbeContext.fillStyle;
  const nextValue = resolved && resolved !== FREE_TEXT_COLOR_SENTINEL ? resolved : null;
  freeTextResolvedColorCache.set(normalizedCompact, nextValue);
  return nextValue;
}

function getFreeTextWordMatches(text) {
  return Array.from(String(text || "").matchAll(/[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*/g));
}

function getFreeTextColorPhraseMatch(text, wordMatches, startIndex) {
  const startMatch = wordMatches[startIndex];
  if (!startMatch) return null;

  const maxPhraseWords = 4;
  const longestIndex = Math.min(wordMatches.length - 1, startIndex + maxPhraseWords - 1);
  for (let endIndex = longestIndex; endIndex >= startIndex; endIndex -= 1) {
    const endMatch = wordMatches[endIndex];
    const phraseStart = Number(startMatch.index || 0);
    const phraseEnd = Number(endMatch.index || 0) + endMatch[0].length;
    const phraseText = String(text || "").slice(phraseStart, phraseEnd);
    if (endIndex > startIndex) {
      const betweenWords = String(text || "").slice(
        Number(startMatch.index || 0) + startMatch[0].length,
        phraseEnd - endMatch[0].length,
      );
      if (!/^[\s-]+(?:[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*[\s-]+)*$/.test(betweenWords)) {
        continue;
      }
    }
    const resolvedColor = resolveFreeTextColorToken(phraseText);
    if (resolvedColor) {
      return {
        text: phraseText,
        start: phraseStart,
        end: phraseEnd,
        endIndex,
        resolvedColor,
      };
    }
  }
  return null;
}

function buildFreeTextColorTokenMarkup(word, backgroundRgb, resolvedColor = null) {
  const effectiveResolvedColor = resolvedColor || resolveFreeTextColorToken(word);
  if (!effectiveResolvedColor) {
    return escapeFreeTextHtml(word);
  }

  const classNames = [];
  let style = "";
  if (effectiveResolvedColor) {
    classNames.push("free-text-color-token");
    const colorRgb = parseCssColorToRgb(effectiveResolvedColor);
    style = `color: ${effectiveResolvedColor};`;
    if (colorRgb && backgroundRgb && getContrastRatio(colorRgb, backgroundRgb) < 2.6) {
      const colorLuminance = getRelativeLuminance(colorRgb);
      const backgroundLuminance = getRelativeLuminance(backgroundRgb);
      const shadowColor = colorLuminance <= backgroundLuminance
        ? "rgba(255,255,255,0.85)"
        : "rgba(96,96,96,0.9)";
      style += ` text-shadow: 0 0 1px ${shadowColor};`;
    }
  }
  const classAttribute = classNames.length ? ` class="${classNames.join(" ")}"` : "";
  const styleAttribute = style ? ` style="${style}"` : "";
  return `<span${classAttribute}${styleAttribute}>${escapeFreeTextHtml(word)}</span>`;
}

function buildFreeTextHighlightMarkup(value) {
  const text = String(value || "");
  const backgroundRgb = getFreeTextEditorBackgroundRgb();
  const wordMatches = getFreeTextWordMatches(text);
  let html = "";
  let lastIndex = 0;

  for (let index = 0; index < wordMatches.length; index += 1) {
    const match = wordMatches[index];
    const matchIndex = Number(match.index || 0);
    const phraseMatch = getFreeTextColorPhraseMatch(text, wordMatches, index);
    const nextStart = phraseMatch ? phraseMatch.start : matchIndex;
    html += escapeFreeTextHtml(text.slice(lastIndex, nextStart));
    if (phraseMatch) {
      html += buildFreeTextColorTokenMarkup(
        phraseMatch.text,
        backgroundRgb,
        phraseMatch.resolvedColor,
      );
      lastIndex = phraseMatch.end;
      index = phraseMatch.endIndex;
      continue;
    }
    html += buildFreeTextColorTokenMarkup(
      match[0],
      backgroundRgb,
      null,
    );
    lastIndex = matchIndex + match[0].length;
  }

  html += escapeFreeTextHtml(text.slice(lastIndex));
  if (text.endsWith("\n")) {
    html += "&#8203;";
  }
  return html;
}

function wrapFreeTextHighlightMarkup(markup) {
  return `<div class="free-text-highlight-content">${markup}</div>`;
}

function syncFreeTextHighlightMetrics() {
  if (!freeText || !freeTextEditor) return;
  const verticalScrollbarGutter = Math.max(0, freeText.offsetWidth - freeText.clientWidth);
  freeTextEditor.style.setProperty("--free-text-scrollbar-gutter", `${verticalScrollbarGutter}px`);
}

function syncFreeTextHighlightScroll() {
  if (!freeTextHighlight || !freeText) return;
  syncFreeTextHighlightMetrics();
  freeTextEditor.style.setProperty("--free-text-scroll-top", `${Math.max(0, freeText.scrollTop)}px`);
  freeTextEditor.style.setProperty("--free-text-scroll-left", `${Math.max(0, freeText.scrollLeft)}px`);
}

function syncFreeTextHighlightState() {
  if (!freeText || !freeTextEditor || !freeTextHighlight) return;
  const value = String(freeText.value || "");
  const isEmpty = !value;
  freeTextEditor.classList.toggle("is-disabled", !!freeText.disabled);
  freeTextHighlight.classList.toggle("placeholder", isEmpty);
  if (isEmpty) {
    freeTextHighlight.innerHTML = wrapFreeTextHighlightMarkup(escapeFreeTextHtml(freeText.placeholder || ""));
  } else {
    freeTextHighlight.innerHTML = wrapFreeTextHighlightMarkup(buildFreeTextHighlightMarkup(value));
  }
  syncFreeTextHighlightScroll();
}

const freeTextHighlightResizeObserver = typeof ResizeObserver === "function"
  ? new ResizeObserver(() => {
    syncFreeTextHighlightMetrics();
    syncFreeTextHighlightScroll();
  })
  : null;
if (freeTextHighlightResizeObserver && freeText) {
  freeTextHighlightResizeObserver.observe(freeText);
}

let freeTextSaveTimeout = null;
function queueFreeTextSave(path) {
  if (freeTextSaveTimeout) clearTimeout(freeTextSaveTimeout);
  freeTextSaveTimeout = setTimeout(() => {
    freeTextSaveTimeout = null;
    saveFreeText(path);
  }, 400);
}

function setFreeTextDraftForPath(path, value, options = {}) {
  const { syncEditor = false } = options;
  const nextFreeText = String(value || "");
  if (!state.captionCache[path]) {
    state.captionCache[path] = { enabled_sentences: [], free_text: "" };
  }
  state.captionCache[path].free_text = nextFreeText;
  state.captionDraftPaths.add(path);
  if (syncEditor && state.selectedPaths.size === 1 && state.selectedPaths.has(path)) {
    freeText.value = nextFreeText;
    syncFreeTextHighlightState();
  }
  queueFreeTextSave(path);
}

freeText.addEventListener("input", () => {
  syncFreeTextHighlightState();
  if (state.selectedPaths.size !== 1) return;
  const path = [...state.selectedPaths][0];
  setFreeTextDraftForPath(path, freeText.value);
});
freeText.addEventListener("scroll", syncFreeTextHighlightScroll);
syncFreeTextHighlightState();

async function saveFreeText(path) {
  const cap = state.captionCache[path];
  if (!cap) return;
  const savedFreeText = String(cap.free_text || "");
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
    if (String(state.captionCache[path]?.free_text || "") === savedFreeText) {
      state.captionDraftPaths.delete(path);
    }
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
    cell.style.width = cell.classList.contains("thumb-cell-double")
      ? `${(state.thumbSize * 2) + 8}px`
      : state.thumbSize + "px";
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

Object.assign(globalThis, {
  loadCaptionData,
  loadMetadataData,
  loadMultiCaptionState,
  loadMultiMetadataState,
  updateMultiInfo,
  renderSentences,
  autoCaptionSelected,
  addFreeTextNow,
  toggleFreeTextRegionPicker,
  describeFreeTextRegion,
  saveMetadataForSelection,
  refreshGroupCaptions,
  refreshSectionCaptions,
  refreshSentenceCaption,
});

if (Array.isArray(state.sections) && state.sections.length > 0 && sectionContainer && !sectionContainer.childElementCount) {
  renderSentences({ force: true, includePreview: false });
}

// ===== RESIZE HANDLES =====
