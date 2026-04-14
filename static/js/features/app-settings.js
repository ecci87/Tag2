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

  videoTrainingProfileLabel.textContent = `${profile.label} \u2022 ${profile.num_frames}f @ ${profile.fps} fps`;
  videoTrainingGuidanceLabel.textContent = guidance.join(" \u2022 ");
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
  state.ollamaMaxOutputTokens = Number(settings.ollama_max_output_tokens || 64) || 64;
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
  if (Object.prototype.hasOwnProperty.call(settings, "sections")) {
    state.sections = normalizeSectionsData(settings.sections);
    if (typeof renderSentences === "function") {
      renderSentences({ force: true, includePreview: false });
    }
  }
  autoFreeTextCheckbox.checked = !!state.ollamaEnableFreeText;
  autoPreviewCheckbox.checked = !!state.comfyuiAutoPreviewEnabled;
  populateVideoTrainingProfileSelect();
  setVideoTrainingPresetsStatus();
  renderVideoTrainingSummary();
  if (typeof renderGifConvertButton === "function") {
    renderGifConvertButton();
  }
  if (state.images.length && typeof renderGrid === "function") {
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
  settingsMaxOutputTokensInput.value = String(state.ollamaMaxOutputTokens);
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
  const maxOutputTokens = Number(settingsMaxOutputTokensInput.value || "64") || 64;
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
    ollama_max_output_tokens: maxOutputTokens,
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

