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
  state.captionDraftPaths.clear();
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
    state.imageVersions = Object.fromEntries(
      (state.images || []).flatMap((img) => {
        const versions = [[img.path, img.mtime || 0]];
        if (img.prompt_preview_path) {
          versions.push([img.prompt_preview_path, img.prompt_preview_mtime || 0]);
        }
        return versions;
      })
    );
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
  uploadQueueText.textContent = parts.join(" \u2022 ");
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
      state.uploadQueueLastSummary = summary.join(" \u2022 ");

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

function getGridPreservePath() {
  return state.previewPath || state.lastClickedPath || [...state.selectedPaths][0] || null;
}

function rerenderGridPreservingView() {
  renderGrid({ preservePath: getGridPreservePath(), preserveScrollTop: fileGridContainer.scrollTop });
}

function setImagePromptPreviewPath(sourcePath, previewPath) {
  const image = state.images.find((item) => item.path === sourcePath);
  if (!image) return false;
  const nextPreviewPath = String(previewPath || "").trim();
  if (String(image.prompt_preview_path || "") === nextPreviewPath) {
    return false;
  }
  image.prompt_preview_path = nextPreviewPath;
  image.prompt_preview_mtime = nextPreviewPath ? Number(getImageVersion(nextPreviewPath) || Date.now()) : 0;
  if (state.showPromptPreviewThumbnails) {
    rerenderGridPreservingView();
  }
  return true;
}

function createThumbnailImageElement(imagePath, thumbLoadSize) {
  const imgEl = document.createElement("img");
  imgEl.dataset.path = imagePath;
  imgEl.loading = "lazy";
  imgEl.addEventListener("load", () => {
    storeThumbnailDimensions(imagePath, imgEl.naturalWidth, imgEl.naturalHeight);
  });
  const cachedUrl = thumbBlobCache.get(`${imagePath}:${thumbLoadSize}:${getImageVersion(imagePath)}`);
  imgEl.src = cachedUrl || buildImageApiUrl("thumbnail", imagePath, { size: thumbLoadSize });
  return imgEl;
}

function appendStandardThumbnailBadges(cell, img) {
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
}

function createThumbCell(img, index, size, thumbLoadSize, options = {}) {
  const {
    imagePath = img.path,
    label = img.name,
    title = img.name,
    sourcePath = img.path,
    isPromptPreview = false,
  } = options;
  const cell = document.createElement("div");
  let cls = "thumb-cell";
  if (state.selectedPaths.has(sourcePath)) cls += " selected";
  if (!isPromptPreview) {
    if (img.has_caption) cls += " has-caption";
    if (img.has_mask) cls += " has-mask";
    if (!imageConformsToAspectRatios(img)) cls += " aspect-mismatch";
  } else {
    cls += " prompt-preview-thumb";
  }
  cell.className = cls;
  cell.style.width = size + "px";
  cell.style.height = size + "px";
  cell.dataset.index = index;
  cell.dataset.path = imagePath;
  if (sourcePath !== imagePath) {
    cell.dataset.sourcePath = sourcePath;
  }
  cell.dataset.mediaType = isPromptPreview ? "image" : (img.media_type || getMediaType(img.path));

  cell.appendChild(createThumbnailImageElement(imagePath, thumbLoadSize));

  const nameEl = document.createElement("div");
  nameEl.className = "thumb-name";
  nameEl.textContent = label;
  nameEl.title = title;
  cell.appendChild(nameEl);

  if (isPromptPreview) {
    const previewBadge = document.createElement("div");
    previewBadge.className = "prompt-preview-badge";
    previewBadge.textContent = "PV";
    previewBadge.title = `Prompt preview for ${img.name}`;
    cell.appendChild(previewBadge);
    cell.addEventListener("click", (e) => handlePromptPreviewThumbClick(index, sourcePath, imagePath, e));
    cell.addEventListener("dblclick", (e) => {
      e.preventDefault();
      fetch(`/api/open-in-explorer?path=${encodeURIComponent(imagePath)}`);
    });
    return cell;
  }

  appendStandardThumbnailBadges(cell, img);

  cell.addEventListener("click", (e) => handleThumbClick(index, e));
  cell.addEventListener("dblclick", (e) => {
    e.preventDefault();
    fetch(`/api/open-in-explorer?path=${encodeURIComponent(img.path)}`);
  });
  return cell;
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
    const previewPath = state.showPromptPreviewThumbnails && img.media_type === "image"
      ? String(img.prompt_preview_path || "").trim()
      : "";
    if (previewPath && previewPath !== img.path) {
      const pair = document.createElement("div");
      pair.className = "thumb-pair";
      pair.appendChild(createThumbCell(img, index, size, thumbLoadSize));
      pair.appendChild(createThumbCell(img, index, size, thumbLoadSize, {
        imagePath: previewPath,
        label: "Preview",
        title: `Prompt preview for ${img.name}`,
        sourcePath: img.path,
        isPromptPreview: true,
      }));
      fileGrid.appendChild(pair);
      continue;
    }
    fileGrid.appendChild(createThumbCell(img, index, size, thumbLoadSize));
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
    const selectionPath = cell.dataset.sourcePath || cell.dataset.path;
    cell.classList.toggle("selected", state.selectedPaths.has(selectionPath));
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

function handlePromptPreviewThumbClick(index, sourcePath, previewPath, event) {
  handleThumbClick(index, event);
  if (event.ctrlKey || event.metaKey || event.shiftKey) return;
  if (!previewPath) return;
  if (state.selectedPaths.size !== 1 || !state.selectedPaths.has(sourcePath)) return;
  seedPromptPreviewSnapshotFromFile(sourcePath, previewPath);
  setPromptPreviewDisplayPath(previewPath, { preserveView: true }).catch(() => {});
  statusBar.textContent = `Showing prompt preview for ${getFileLabel(sourcePath)}`;
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

  // Mouse wheel zoom - zooms toward cursor position
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
