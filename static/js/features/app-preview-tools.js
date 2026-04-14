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

function resolveLatestPromptPreviewFile(files = [], summary = null) {
  const latestOutputPath = String(summary?.latest_output_path || "").trim();
  if (latestOutputPath) {
    return latestOutputPath;
  }
  const uniqueFiles = Array.from(new Set((files || []).filter(Boolean)));
  return uniqueFiles.length > 0 ? uniqueFiles[uniqueFiles.length - 1] : "";
}

function getLatestPromptPreviewFile(
  sourcePath = state.previewPath,
  files = getPromptPreviewFiles(sourcePath),
  summary = isPromptPreviewSourceActive(sourcePath) ? state.promptPreview.summary : null,
) {
  if (!isPromptPreviewSourceActive(sourcePath)) return "";
  return resolveLatestPromptPreviewFile(files, summary);
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
    parts.push(`${media.name} \u2022 video`);
  } else {
    parts.push(hasAppliedCrop() ? `${media.name} \u2022 cropped` : media.name);
  }

  const latestPromptPreviewPath = getLatestPromptPreviewFile(state.previewPath);
  if (latestPromptPreviewPath) {
    parts.push(state.promptPreview.displayPath === latestPromptPreviewPath ? "prompt preview" : "original");
  }

  previewInfo.textContent = parts.join(" \u2022 ");
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

function getPromptPreviewCreateSelectionPaths() {
  if (state.selectedPaths.size === 0) return [];
  return [...state.selectedPaths].filter((path) => isImageMediaPath(path));
}

function canCreatePromptPreviewSelection() {
  return state.selectedPaths.size > 0 && isSelectionImagesOnly();
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
  const selectedPaths = getPromptPreviewCreateSelectionPaths();
  const selectionReady = canCreatePromptPreviewSelection();
  const selectionCount = selectedPaths.length;
  const sourcePath = selectedPaths.includes(state.previewPath) ? state.previewPath : "";
  const sourceActive = sourcePath ? isPromptPreviewSourceActive(sourcePath) : false;
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
  let title = configError || (selectionCount === 1
    ? "Queue a new ComfyUI prompt preview for the selected image"
    : `Queue new ComfyUI prompt preview jobs for ${selectionCount} selected images`);
  if (!selectionReady) {
    title = state.selectedPaths.size > 0
      ? "Prompt preview currently supports image selections only"
      : "Select one or more images to queue prompt previews";
  } else if (state.promptPreview.loading && sourceActive) {
    label = "Creating Preview...";
    title = selectionCount === 1
      ? "Queueing a new ComfyUI prompt preview"
      : `Queueing ComfyUI prompt previews for ${selectionCount} images`;
  } else if (hasActiveJobs) {
    title = `${spawned} prompt preview job${spawned === 1 ? "" : "s"} queued for ${getFileLabel(sourcePath)}. The latest result will show automatically when the next job finishes.`;
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

function renderPromptPreviewGridToggle() {
  if (!promptPreviewGridToggleBtn) return;
  const active = !!state.showPromptPreviewThumbnails;
  promptPreviewGridToggleBtn.classList.toggle("active", active);
  promptPreviewGridToggleBtn.setAttribute("aria-pressed", active ? "true" : "false");
  promptPreviewGridToggleBtn.title = active
    ? "Hide prompt preview thumbnails beside source thumbnails"
    : "Show prompt preview thumbnails beside source thumbnails when available";
}

function seedPromptPreviewSnapshotFromFile(sourcePath, previewPath) {
  if (!sourcePath || !previewPath) return;
  const sameSource = isPromptPreviewSourceActive(sourcePath);
  const baseSummary = sameSource
    ? { ...createPromptPreviewSummary(), ...(state.promptPreview.summary || {}) }
    : createPromptPreviewSummary();
  const files = Array.from(new Set([
    ...(sameSource ? getPromptPreviewFiles(sourcePath) : []),
    previewPath,
  ]));
  applyPromptPreviewSnapshot(sourcePath, {
    jobs: sameSource ? state.promptPreview.jobs : [],
    summary: {
      ...baseSummary,
      total: Math.max(1, Number(baseSummary.total || 0)),
      spawned: Math.max(1, Number(baseSummary.spawned || 0)),
      completed: Math.max(1, Number(baseSummary.completed || 0)),
      latest_output_path: previewPath,
    },
    files,
  }, {
    autoDisplayLatest: false,
    allowCompletedAutodisplay: false,
    allowFileChangeAutodisplay: false,
  });
}

function applyPromptPreviewSnapshot(sourcePath, payload, options = {}) {
  const {
    autoDisplayLatest = false,
    allowCompletedAutodisplay = true,
    allowFileChangeAutodisplay = true,
  } = options;
  const sameSource = isPromptPreviewSourceActive(sourcePath);
  const jobs = Array.isArray(payload?.jobs) ? payload.jobs : [];
  const summary = { ...createPromptPreviewSummary(), ...(payload?.summary || {}) };
  const files = Array.from(new Set((Array.isArray(payload?.files) ? payload.files : []).filter(Boolean)));
  const previousSummary = sameSource
    ? { ...createPromptPreviewSummary(), ...(state.promptPreview.summary || {}) }
    : createPromptPreviewSummary();
  const previousCompleted = Number(previousSummary.completed || 0);
  const previousFilesKey = sameSource ? state.promptPreview.lastFilesKey || "" : "";
  const previousLatestPreviewPath = sameSource
    ? resolveLatestPromptPreviewFile(state.promptPreview.files, previousSummary)
    : "";
  const nextFilesKey = files.join("\n");
  const filesChanged = nextFilesKey !== previousFilesKey;
  const latestPreviewPath = resolveLatestPromptPreviewFile(files, summary);
  summary.latest_output_path = latestPreviewPath;
  const completedIncreased = sameSource && Number(summary.completed || 0) > Number(previousSummary.completed || 0);
  const latestPreviewChanged = !!latestPreviewPath && latestPreviewPath !== previousLatestPreviewPath;
  const latestPreviewRefreshed = completedIncreased
    && !!latestPreviewPath
    && latestPreviewPath === previousLatestPreviewPath;

  if (filesChanged) {
    const versionBase = Date.now();
    files.forEach((path, index) => {
      state.imageVersions[path] = versionBase + index;
    });
  }
  if (latestPreviewRefreshed) {
    bumpImageVersion(latestPreviewPath);
    invalidateImageCaches(latestPreviewPath);
  }

  state.promptPreview.sourcePath = sourcePath;
  state.promptPreview.jobs = jobs;
  state.promptPreview.summary = summary;
  state.promptPreview.files = files;
  state.promptPreview.lastFilesKey = nextFilesKey;

  setImagePromptPreviewPath(sourcePath, latestPreviewPath);
  if (latestPreviewPath && (filesChanged || latestPreviewRefreshed)) {
    refreshVisibleThumbnail(latestPreviewPath);
  }

  if (state.promptPreview.displayPath && !files.includes(state.promptPreview.displayPath)) {
    state.promptPreview.displayPath = "";
  }

  const cyclePaths = getPromptPreviewCyclePaths(sourcePath, files);
  const currentDisplayPath = getPromptPreviewCurrentDisplayPath(sourcePath);
  state.promptPreview.cycleIndex = cyclePaths.indexOf(currentDisplayPath);

  renderPromptPreviewControls();
  renderPromptPreviewGridToggle();
  renderPreviewInfo();

  const shouldAutoDisplayLatest = sameSource
    && state.previewPath === sourcePath
    && state.previewMediaType === "image"
    && (latestPreviewChanged || latestPreviewRefreshed)
    && (
      autoDisplayLatest
      || (allowCompletedAutodisplay && completedIncreased)
      || (allowFileChangeAutodisplay && previousFilesKey && filesChanged)
    );

  if (shouldAutoDisplayLatest) {
    setPromptPreviewDisplayPath(latestPreviewPath, { preserveView: true }).catch(() => {});
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
  const {
    autoDisplayLatest = false,
    allowCompletedAutodisplay = true,
    allowFileChangeAutodisplay = true,
  } = options;
  const data = await fetchPromptPreviewStatus(sourcePath, options);
  applyPromptPreviewSnapshot(sourcePath, data, {
    autoDisplayLatest: !!autoDisplayLatest,
    allowCompletedAutodisplay,
    allowFileChangeAutodisplay,
  });
  return data;
}

function getPromptPreviewCaptionState(sourcePath = state.previewPath) {
  const caption = state.captionCache[sourcePath] || { enabled_sentences: [], free_text: "" };
  const activeFreeText = state.previewPath === sourcePath
    && state.selectedPaths.size === 1
    && state.selectedPaths.has(sourcePath)
    && !freeText.disabled;
  return {
    enabled_sentences: orderEnabledSentences(caption.enabled_sentences || []),
    free_text: activeFreeText ? String(freeText.value || "") : String(caption.free_text || ""),
  };
}

async function queuePromptPreviewFromCurrentCaption(sourcePath = state.previewPath) {
  if (!state.captionCache[sourcePath]) {
    await loadCaptionData(sourcePath);
  }
  const caption = getPromptPreviewCaptionState(sourcePath);
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
  const selectedPaths = getPromptPreviewCreateSelectionPaths();
  if (!canCreatePromptPreviewSelection()) {
    const message = state.selectedPaths.size > 0
      ? "Prompt preview currently supports image selections only."
      : "Select one or more images first.";
    showErrorToast(message);
    statusBar.textContent = message;
    return;
  }
  const sourcePath = selectedPaths.includes(state.previewPath) ? state.previewPath : selectedPaths[0] || "";
  const configError = getPromptPreviewConfigError({ requireWorkflow: true });
  if (configError) {
    showErrorToast(configError);
    statusBar.textContent = configError;
    renderPromptPreviewControls();
    return;
  }

  if (sourcePath) {
    state.promptPreview.sourcePath = sourcePath;
  }
  state.promptPreview.loading = true;
  statusBar.textContent = selectedPaths.length === 1
    ? `Queueing prompt preview for ${getFileLabel(sourcePath)}...`
    : `Queueing prompt previews for ${selectedPaths.length} images...`;
  renderPromptPreviewControls();
  let queuedCount = 0;
  let failedCount = 0;
  let firstErrorMessage = "";
  try {
    for (let index = 0; index < selectedPaths.length; index += 1) {
      const path = selectedPaths[index];
      if (selectedPaths.length > 1) {
        statusBar.textContent = `Queueing prompt previews ${index + 1}/${selectedPaths.length}: ${getFileLabel(path)}...`;
      }
      try {
        const queued = await queuePromptPreviewFromCurrentCaption(path);
        queuedCount += 1;
        if (path === sourcePath) {
          applyPromptPreviewSnapshot(path, queued, {
            autoDisplayLatest: false,
            allowCompletedAutodisplay: false,
            allowFileChangeAutodisplay: false,
          });
        }
      } catch (err) {
        failedCount += 1;
        if (!firstErrorMessage) {
          firstErrorMessage = err.message || "Failed to queue prompt preview";
        }
        console.error(`Failed to queue prompt preview for ${path}:`, err);
      }
    }

    if (queuedCount > 0 && failedCount === 0) {
      statusBar.textContent = queuedCount === 1
        ? `Queued prompt preview for ${getFileLabel(sourcePath)}`
        : `Queued ${queuedCount} prompt preview jobs`;
    } else if (queuedCount > 0) {
      const message = `Queued ${queuedCount} prompt preview job${queuedCount === 1 ? "" : "s"}; ${failedCount} failed.`;
      showErrorToast(message);
      statusBar.textContent = message;
    } else {
      throw new Error(firstErrorMessage || "Failed to queue prompt previews");
    }
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
      const status = await refreshPromptPreviewStatus(sourcePath, {
        showErrors: false,
        autoDisplayLatest: false,
        allowCompletedAutodisplay: false,
        allowFileChangeAutodisplay: false,
      });
      latestPreviewPath = resolveLatestPromptPreviewFile(
        Array.isArray(status.files) ? status.files : [],
        status.summary || null,
      );
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
  const revealPath = state.promptPreview.displayPath || getLatestPromptPreviewFile(sourcePath) || "";
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

