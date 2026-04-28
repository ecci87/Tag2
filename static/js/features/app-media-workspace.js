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

  const cells = fileGrid.querySelectorAll(`.thumb-cell[data-path="${CSS.escape(path)}"]`);
  cells.forEach((cell) => {
    cell.classList.toggle("has-mask", !!hasMask);
  });
  if (typeof syncVisibleThumbnailMaskOverlay === "function") {
    syncVisibleThumbnailMaskOverlay(path);
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
    fileCount.textContent = `${visible}/${total} media files \u2022 ${captioned} captioned \u2022 ${filterCount} filter${filterCount === 1 ? "" : "s"}`;
    return;
  }
  fileCount.textContent = `${total} media files \u2022 ${captioned} captioned`;
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

async function ensureImageCropScaleReady(path) {
  if (!path || !isImageMediaPath(path)) return true;
  const cropState = state.imageCrops[path];
  if (cropState?.current_width && cropState?.current_height) {
    return true;
  }
  await loadCropData(path);
  const nextCropState = state.imageCrops[path];
  return !!(nextCropState?.current_width && nextCropState?.current_height);
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

function canUseRegionDescription() {
  return state.selectedPaths.size === 1
    && !!state.previewPath
    && !!imgNatW
    && !!imgNatH
    && isImageMediaPath(state.previewPath);
}

function isAiRegionPickerActive() {
  return !!state.aiRegionPicker.active;
}

function isAiRegionPickerVisible() {
  if (!isAiRegionPickerActive()) {
    return false;
  }
  if (isMaskEditorVisible()) {
    return false;
  }
  return canUseRegionDescription();
}

function canEditAiRegionPicker() {
  return isAiRegionPickerVisible();
}

function getAiRegionQueuePendingCount() {
  return Array.isArray(state.aiRegionPicker.queue) ? state.aiRegionPicker.queue.length : 0;
}

function updateAiRegionQueueProgress() {
  if (state.autoCaptioning) {
    return;
  }
  const pending = getAiRegionQueuePendingCount();
  const running = state.aiRegionPicker.loading && state.aiRegionPicker.currentJob ? 1 : 0;
  if (!(pending > 0 || running > 0)) {
    resetAutoCaptionProgress();
    renderModelLog();
    return;
  }
  const totalJobs = Math.max(
    Number(state.aiRegionPicker.totalJobs || 0),
    Number(state.aiRegionPicker.completedJobs || 0) + Number(state.aiRegionPicker.failedJobs || 0) + pending + running,
  );
  updateAutoCaptionProgress({
    visible: true,
    scopeLabel: "Region Queue",
    totalImages: Math.max(1, totalJobs),
    processedImages: Math.max(0, Number(state.aiRegionPicker.completedJobs || 0)),
    errors: Math.max(0, Number(state.aiRegionPicker.failedJobs || 0)),
    completedImages: Math.max(0, Number(state.aiRegionPicker.completedJobs || 0) + Number(state.aiRegionPicker.failedJobs || 0)),
    queuedItems: pending,
    enableFreeText: false,
    freeTextOnly: false,
    currentPath: state.aiRegionPicker.currentJob?.path || state.previewPath || "",
    currentMessage: state.aiRegionPicker.currentJob
      ? `Processing${pending > 0 ? ` • ${pending} queued` : ""}`
      : `${pending} queued`,
    currentStepIndex: 0,
    currentStepTotal: state.aiRegionPicker.currentJob ? 1 : 0,
  });
}

function canInteractPreviewRect() {
  return canEditCrop() || canEditAiRegionPicker();
}

function beginAiRegionPicker() {
  if (!canUseRegionDescription()) {
    const message = state.selectedPaths.size !== 1
      ? "Select a single image first"
      : "Region description is only available for images";
    statusBar.textContent = message;
    showErrorToast(message);
    return false;
  }
  const hasBackgroundQueueWork = !!(
    state.aiRegionPicker.loading
    || state.aiRegionPicker.currentJob
    || getAiRegionQueuePendingCount() > 0
  );
  if (state.cropDraft || state.cropInteraction) {
    const message = "Finish or cancel the current crop edit first";
    statusBar.textContent = message;
    showErrorToast(message);
    return false;
  }
  state.aiRegionPicker.active = true;
  if (!hasBackgroundQueueWork) {
    state.aiRegionPicker.loading = false;
    state.aiRegionPicker.abortController = null;
    state.aiRegionPicker.queue = [];
    state.aiRegionPicker.currentJob = null;
    state.aiRegionPicker.completedJobs = 0;
    state.aiRegionPicker.failedJobs = 0;
    state.aiRegionPicker.totalJobs = 0;
  }
  state.cropDraft = null;
  state.cropDirty = false;
  state.cropInteraction = null;
  setCropGuideToCenter();
  updateActionButtons();
  updateAiRegionQueueProgress();
  statusBar.textContent = "Drag a rectangle over the image to queue a region description";
  return true;
}

function stopAiRegionPicker(options = {}) {
  const { abort = false, clearJobs = abort, keepStatus = false } = options;
  const controller = state.aiRegionPicker.abortController;
  const hasBackgroundQueueWork = !!(
    (state.aiRegionPicker.loading && state.aiRegionPicker.currentJob)
    || getAiRegionQueuePendingCount() > 0
  );
  state.aiRegionPicker.active = false;
  state.cropDraft = null;
  state.cropDirty = false;
  state.cropInteraction = null;
  if (abort && controller) {
    controller.abort();
  }
  if (clearJobs) {
    state.aiRegionPicker.queue = [];
  }
  if (abort) {
    state.aiRegionPicker.loading = false;
    state.aiRegionPicker.abortController = null;
    state.aiRegionPicker.currentJob = null;
  }
  if (clearJobs && !state.aiRegionPicker.loading) {
    state.aiRegionPicker.completedJobs = 0;
    state.aiRegionPicker.failedJobs = 0;
    state.aiRegionPicker.totalJobs = 0;
  }
  clearCropGuide();
  renderCropOverlay();
  renderVideoEditPanel();
  renderMaskEditorUi();
  updateActionButtons();
  updateAiRegionQueueProgress();
  if (!keepStatus) {
    statusBar.textContent = hasBackgroundQueueWork && !abort && !clearJobs
      ? "Region picker closed; queued descriptions continue"
      : "Region picker cancelled";
  }
}

function resetAiRegionPickerDraft(options = {}) {
  const { keepStatus = false } = options;
  if (!state.aiRegionPicker.active) {
    return;
  }
  state.cropDraft = null;
  state.cropDirty = false;
  state.cropInteraction = null;
  setCropGuideToCenter();
  renderCropOverlay();
  renderVideoEditPanel();
  renderMaskEditorUi();
  updateActionButtons();
  updateAiRegionQueueProgress();
  if (!keepStatus) {
    statusBar.textContent = "Drag a rectangle over the image to queue a region description";
  }
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

function isCropEditActive() {
  return !isAiRegionPickerActive() && canEditCrop() && !!(state.cropDraft || state.cropInteraction);
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
  if (!canInteractPreviewRect()) {
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
  const aiRegionVisible = isAiRegionPickerVisible();
  const cropEditable = canEditCrop();
  const overlayVisible = aiRegionVisible || cropEditable;
  const hasDraft = !!state.cropDraft;
  const videoCrop = state.previewMediaType === "video";
  cropRemoveBtn.classList.toggle("visible", !aiRegionVisible && !videoCrop && cropEditable && !hasDraft && hasAppliedCrop());
  cropCancelBtn.classList.toggle("visible", aiRegionVisible ? overlayVisible : cropEditable && hasDraft);
  cropCancelBtn.textContent = aiRegionVisible ? "Cancel" : (videoCrop ? "Clear Crop" : "Cancel");
  cropCancelBtn.disabled = false;
  cropApplyBtn.classList.toggle("visible", !aiRegionVisible && !videoCrop && cropEditable && hasDraft && state.cropDirty);
  cropApplyBtn.textContent = aiRegionVisible ? (state.aiRegionPicker.loading ? "Describing..." : "Describe") : "Apply";
  cropApplyBtn.disabled = aiRegionVisible ? (!hasDraft || !!state.aiRegionPicker.loading) : false;
  rotateControls.classList.toggle("visible", !aiRegionVisible && !videoCrop && cropEditable && !hasDraft);
  renderGifConvertButton();
}

function clampRegionFraction(value) {
  return Math.max(0, Math.min(1, Number(value || 0)));
}

function describeAiRegionCenterOffset(centerX, centerY) {
  let horizontal = null;
  let vertical = null;
  if (centerX < 0.46) {
    horizontal = "left";
  } else if (centerX > 0.54) {
    horizontal = "right";
  }
  if (centerY < 0.46) {
    vertical = "above";
  } else if (centerY > 0.54) {
    vertical = "below";
  }

  if (!horizontal && !vertical) {
    return "center";
  }
  if (horizontal && vertical) {
    const horizontalPhrase = horizontal === "left" ? "to the left" : "to the right";
    return `slightly ${vertical} and ${horizontalPhrase} of the center`;
  }
  if (horizontal) {
    const horizontalPhrase = horizontal === "left" ? "to the left" : "to the right";
    return `slightly ${horizontalPhrase} of the center`;
  }
  return `slightly ${vertical} the center`;
}

function describeAiRegionPosition(crop, imageWidth = imgNatW, imageHeight = imgNatH) {
  const safeWidth = Math.max(1, Number(imageWidth || 1));
  const safeHeight = Math.max(1, Number(imageHeight || 1));
  const left = clampRegionFraction(Number(crop?.x || 0) / safeWidth);
  const top = clampRegionFraction(Number(crop?.y || 0) / safeHeight);
  const right = clampRegionFraction((Number(crop?.x || 0) + Number(crop?.w || 0)) / safeWidth);
  const bottom = clampRegionFraction((Number(crop?.y || 0) + Number(crop?.h || 0)) / safeHeight);
  const widthFraction = Math.max(0, right - left);
  const heightFraction = Math.max(0, bottom - top);
  const centerX = (left + right) / 2;
  const centerY = (top + bottom) / 2;

  if (widthFraction >= 0.58 && heightFraction >= 0.40) {
    if (top <= 0.12 && centerY <= 0.34) {
      return "upper half of the image";
    }
    if (bottom >= 0.88 && centerY >= 0.66) {
      return "lower half of the image";
    }
  }
  if (heightFraction >= 0.58 && widthFraction >= 0.40) {
    if (left <= 0.12 && centerX <= 0.34) {
      return "left half of the image";
    }
    if (right >= 0.88 && centerX >= 0.66) {
      return "right half of the image";
    }
  }

  const touchesLeft = left <= 0.14;
  const touchesRight = right >= 0.86;
  const touchesTop = top <= 0.14;
  const touchesBottom = bottom >= 0.86;
  const compactCorner = widthFraction <= 0.38 && heightFraction <= 0.38;
  if (compactCorner && touchesTop && touchesLeft && centerX <= 0.34 && centerY <= 0.34) {
    return "upper left corner";
  }
  if (compactCorner && touchesTop && touchesRight && centerX >= 0.66 && centerY <= 0.34) {
    return "upper right corner";
  }
  if (compactCorner && touchesBottom && touchesLeft && centerX <= 0.34 && centerY >= 0.66) {
    return "lower left corner";
  }
  if (compactCorner && touchesBottom && touchesRight && centerX >= 0.66 && centerY >= 0.66) {
    return "lower right corner";
  }

  const farLeft = touchesLeft && centerX <= 0.18 && widthFraction <= 0.24;
  const farRight = touchesRight && centerX >= 0.82 && widthFraction <= 0.24;
  const farTop = touchesTop && centerY <= 0.18 && heightFraction <= 0.24;
  const farBottom = touchesBottom && centerY >= 0.82 && heightFraction <= 0.24;
  const tallVerticalBand = heightFraction >= 0.55 && widthFraction <= 0.5;
  if (tallVerticalBand) {
    if (farLeft) {
      return "far left side";
    }
    if (farRight) {
      return "far right side";
    }
    if (touchesLeft && widthFraction >= 0.38) {
      return "left half of the image";
    }
    if (touchesRight && widthFraction >= 0.38) {
      return "right half of the image";
    }
    if (centerX <= 0.34) {
      return "left side";
    }
    if (centerX >= 0.66) {
      return "right side";
    }
    return "middle vertical strip";
  }

  const wideHorizontalBand = widthFraction >= 0.55 && heightFraction <= 0.5;
  if (wideHorizontalBand) {
    if (touchesTop && heightFraction >= 0.38) {
      return "upper half of the image";
    }
    if (touchesBottom && heightFraction >= 0.38) {
      return "lower half of the image";
    }
    if (farTop) {
      return "far upper part of the image";
    }
    if (farBottom) {
      return "far lower part of the image";
    }
    if (centerY <= 0.34) {
      return "upper part of the image";
    }
    if (centerY >= 0.66) {
      return "lower part of the image";
    }
    return "middle horizontal strip";
  }

  if (centerY >= 0.28 && centerY <= 0.72) {
    if (farLeft) {
      return "far left center";
    }
    if (farRight) {
      return "far right center";
    }
  }

  const horizontal = centerX < 0.34 ? "left" : (centerX > 0.66 ? "right" : "center");
  const vertical = centerY < 0.34 ? "upper" : (centerY > 0.66 ? "lower" : "center");
  if (vertical === "center" && horizontal === "center") {
    return describeAiRegionCenterOffset(centerX, centerY);
  }
  if (vertical === "center") {
    if (touchesLeft && horizontal === "left") {
      return "left center";
    }
    if (touchesRight && horizontal === "right") {
      return "right center";
    }
    return `center ${horizontal} side`;
  }
  if (horizontal === "center") {
    if (farTop) {
      return "far upper center";
    }
    if (farBottom) {
      return "far lower center";
    }
    if (touchesTop && vertical === "upper") {
      return "upper center";
    }
    if (touchesBottom && vertical === "lower") {
      return "lower center";
    }
    return `${vertical} center`;
  }
  return `${vertical} ${horizontal}`;
}

function renderCropOverlay() {
  const crop = state.cropDraft;
  const guide = state.cropGuide;
  const overlayVisible = canEditCrop() || isAiRegionPickerVisible();
  previewStage.classList.toggle("ai-region-picker-mode", canEditAiRegionPicker());
  if (guide && overlayVisible) {
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
  if (!crop || !overlayVisible) {
    cropBox.classList.remove("active");
    updateCropButtons();
    if (typeof renderPreviewActionBar === "function") {
      renderPreviewActionBar();
    }
    return;
  }
  cropBox.classList.add("active");
  cropBox.style.left = (panX + crop.x * zoomLevel) + "px";
  cropBox.style.top = (panY + crop.y * zoomLevel) + "px";
  cropBox.style.width = (crop.w * zoomLevel) + "px";
  cropBox.style.height = (crop.h * zoomLevel) + "px";
  if (isAiRegionPickerVisible()) {
    const regionPosition = describeAiRegionPosition(crop);
    cropLabel.textContent = regionPosition;
    cropLabel.title = `${regionPosition} \u2022 ${crop.w}\u00D7${crop.h}`;
  } else {
    const cropSummary = `${crop.ratio || "custom"} \u2022 ${crop.w}\u00D7${crop.h}`;
    cropLabel.textContent = cropSummary;
    cropLabel.title = cropSummary;
  }
  updateCropButtons();
  if (typeof renderPreviewActionBar === "function") {
    renderPreviewActionBar();
  }
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

function createFreeformCrop(anchorX, anchorY, currentX, currentY) {
  const left = clamp(Math.min(anchorX, currentX), 0, Math.max(0, imgNatW - 1));
  const top = clamp(Math.min(anchorY, currentY), 0, Math.max(0, imgNatH - 1));
  const right = clamp(Math.max(anchorX, currentX), Math.min(imgNatW, left + 1), imgNatW);
  const bottom = clamp(Math.max(anchorY, currentY), Math.min(imgNatH, top + 1), imgNatH);
  return {
    crop: {
      x: left,
      y: top,
      w: Math.max(1, right - left),
      h: Math.max(1, bottom - top),
      ratio: "freeform",
    },
    stickyChoice: null,
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

function resizeFreeformCropFromHandle(baseCrop, handle, currentX, currentY) {
  const minSize = 1;
  let left = baseCrop.x;
  let top = baseCrop.y;
  let right = baseCrop.x + baseCrop.w;
  let bottom = baseCrop.y + baseCrop.h;

  if (handle.includes("w")) {
    left = clamp(currentX, 0, imgNatW - minSize);
  }
  if (handle.includes("e")) {
    right = clamp(currentX, minSize, imgNatW);
  }
  if (handle.includes("n")) {
    top = clamp(currentY, 0, imgNatH - minSize);
  }
  if (handle.includes("s")) {
    bottom = clamp(currentY, minSize, imgNatH);
  }

  if (left > right - minSize) {
    if (handle.includes("w")) {
      left = right - minSize;
    } else {
      right = left + minSize;
    }
  }
  if (top > bottom - minSize) {
    if (handle.includes("n")) {
      top = bottom - minSize;
    } else {
      bottom = top + minSize;
    }
  }

  return {
    crop: {
      x: left,
      y: top,
      w: Math.max(minSize, right - left),
      h: Math.max(minSize, bottom - top),
      ratio: baseCrop.ratio || "freeform",
    },
    stickyChoice: null,
  };
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
    updateImageDimensions(path, data.crop);
    if (state.previewPath === path) {
      updateCropButtons();
    }
  } catch (err) {
    console.error("Failed to load crop:", err);
  }
}

function startCropCreate(event) {
  const start = screenToImage(event.clientX, event.clientY);
  const freeform = isAiRegionPickerActive();
  state.cropInteraction = { mode: "create", anchor: start, stickySnap: null, freeform, button: event.button };
  const ratio = state.cropAspectRatios[0] || { label: "1:1", value: 1 };
  setCropDraft({ x: start.x, y: start.y, w: 1, h: 1, ratio: freeform ? "freeform" : ratio.label }, true);
}

function startCropResize(handle, event) {
  if (!state.cropDraft) return;
  event.preventDefault();
  event.stopPropagation();
  state.cropInteraction = {
    mode: "resize",
    handle,
    baseCrop: { ...state.cropDraft },
    stickySnap: null,
    freeform: isAiRegionPickerActive(),
    button: event.button,
  };
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
    const cropScaleReady = await ensureImageCropScaleReady(path);
    if (!cropScaleReady) {
      throw new Error("Image dimensions are not ready yet");
    }
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
  if (isAiRegionPickerActive()) {
    stopAiRegionPicker();
    return;
  }
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
    await showPreview(path, { preserveView: true });
    statusBar.textContent = direction === "left" ? "Rotated left" : "Rotated right";
  } catch (err) {
    statusBar.textContent = `Rotate error: ${err.message}`;
  }
}

async function flipPreviewImage(axis) {
  if (!state.previewPath) return;
  const path = state.previewPath;
  const horizontal = axis === "horizontal";
  statusBar.textContent = horizontal ? "Flipping horizontally..." : "Flipping vertically...";
  try {
    const resp = await fetch("/api/flip", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ image_path: path, axis }),
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.detail || "Failed to flip image");
    state.imageCrops[path] = data.crop || null;
    updateImageDimensions(path, data.crop);
    clearCropDraft();
    bumpImageVersion(path);
    invalidateImageCaches(path);
    renderGrid();
    await showPreview(path, { preserveView: true });
    statusBar.textContent = horizontal ? "Flipped horizontally" : "Flipped vertically";
  } catch (err) {
    statusBar.textContent = `Flip error: ${err.message}`;
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
  const canRunRegionDescribe = canUseRegionDescription() && !!state.ollamaModel.trim();
  const hasBlockingWorkspaceAction = !!(state.cloning || state.moving || state.extractingFrame || state.uploading);
  const pickerActive = isAiRegionPickerActive();
  cloneFolderBtn.disabled = !state.folder || state.autoCaptioning || pickerActive || hasBlockingWorkspaceAction;
  cloneFolderBtn.textContent = state.cloning ? "Cloning..." : "Clone Folder";
  cloneFolderBtn.title = state.selectedPaths.size > 1
    ? "Clone the selected media files into a new sibling folder"
    : "Clone the whole current folder into a new sibling folder";
  moveSelectedBtn.disabled = !state.folder || !hasSelection || state.autoCaptioning || pickerActive || hasBlockingWorkspaceAction;
  moveSelectedBtn.textContent = state.moving ? "Copying..." : "Copy";
  moveSelectedBtn.title = hasSelection
    ? `Copy ${state.selectedPaths.size} selected media file${state.selectedPaths.size === 1 ? "" : "s"} into another folder`
    : "Select one or more media files to copy into another folder";
  autoCaptionBtn.disabled = hasBlockingWorkspaceAction || pickerActive || (!state.autoCaptioning && !canRunStructured);
  addFreeTextNowBtn.disabled = hasBlockingWorkspaceAction || pickerActive || (!state.autoCaptioning && !canRunFreeTextOnly);
  autoCaptionBtn.classList.toggle("running", state.autoCaptioning);
  addFreeTextNowBtn.classList.toggle("running", state.autoCaptioning && state.autoCaptionMode === "free-text-only");
  if (addRegionFreeTextBtn) {
    addRegionFreeTextBtn.disabled = state.autoCaptioning || (!pickerActive && (hasBlockingWorkspaceAction || !canRunRegionDescribe));
    addRegionFreeTextBtn.classList.toggle("running", !!(state.aiRegionPicker.loading || getAiRegionQueuePendingCount() > 0));
    addRegionFreeTextBtn.classList.toggle("picker-active", pickerActive);
    setGenerateButtonContent(
      addRegionFreeTextBtn,
      pickerActive
        ? "Cancel region picker"
        : "Describe a selected image region and append the result to free text",
      { iconOnly: true, iconFactory: createGenerateRegionPickerIcon }
    );
    addRegionFreeTextBtn.title = pickerActive
      ? "Cancel region picker"
      : "Describe a selected image region and append the result to free text";
  }
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
  if (typeof renderMaskEditorUi === "function") {
    renderMaskEditorUi();
  }
}

function resetAutoCaptionProgress() {
  state.aiProgress = {
    visible: false,
    scopeLabel: "",
    totalImages: 0,
    processedImages: 0,
    errors: 0,
    completedImages: 0,
    queuedItems: 0,
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
  const queuedItems = Math.max(0, Number(progress.queuedItems || 0));

  aiProgressSummary.textContent = `${progress.scopeLabel || "AI"} \u2022 ${progress.processedImages}/${progress.totalImages} done${progress.errors ? ` \u2022 ${progress.errors} error${progress.errors === 1 ? "" : "s"}` : ""}${queuedItems ? ` \u2022 ${queuedItems} queued` : ""}`;
  aiProgressMetric.textContent = `${Math.round(overallPercent)}%`;
  aiProgressCurrentLabel.textContent = progress.currentPath
    ? `${getFileLabel(progress.currentPath)}${progress.currentMessage ? ` \u2022 ${progress.currentMessage}` : ""}`
    : (progress.currentMessage || "Preparing...");
  aiProgressCurrentMetric.textContent = currentStepTotal > 0 ? `${currentStepIndex}/${currentStepTotal}` : "";
  aiProgressOverallFill.style.width = `${overallPercent}%`;
  aiProgressCurrentFill.style.width = `${currentPercent}%`;
}

function renderModelLog() {
  const hasLogUi = state.modelLogLines.length > 0 || !!state.aiProgress.visible;
  renderAutoCaptionProgress();
  modelLogOpenBtn.hidden = false;
  modelLogOpenBtn.disabled = false;
  modelLogOpenBtn.classList.toggle("active", !!state.modelLogOpen);
  modelLogOpenBtn.textContent = state.modelLogOpen ? "Hide Log" : "Log";
  modelLogOpenBtn.setAttribute("aria-expanded", state.modelLogOpen ? "true" : "false");
  modelLogOverlay.classList.toggle("open", !!state.modelLogOpen);
  modelLogOverlay.setAttribute("aria-hidden", state.modelLogOpen ? "false" : "true");
  modelLog.innerHTML = hasLogUi
    ? state.modelLogLines.join("\n")
    : '<span class="log-dim">No log entries yet. Run Auto Caption or describe a selected region to inspect prompt and response details.</span>';
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

