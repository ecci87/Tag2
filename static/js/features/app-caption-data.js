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

