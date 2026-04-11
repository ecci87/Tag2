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
  maskBrushSizeLabel.textContent = `${brushSizePercent.toFixed(1)}% \u00B7 ${Math.round(brushDiameterMaskPx)} px`;
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
      : (state.maskEditor.saving ? "Saving..." : `${Math.round(brushValue)}% \u00B7 ${brushColor}`);
    return;
  }

  maskResetBtn.textContent = `Reset ${Math.round(brushValue)}%`;
  maskResetBtn.title = `Fill the full mask with ${Math.round(brushValue)}%`;
  maskLatentBaseWidthLabel.textContent = `${latentMetrics.baseWidth}px`;
  maskLatentDividerLabel.textContent = `/${latentMetrics.divider}`;
  maskLatentBaseSizeLabel.textContent = `Base ${latentMetrics.baseWidth}\u00D7${latentMetrics.baseHeight}`;
  maskLatentGridSizeLabel.textContent = `Latent ${latentMetrics.latentWidth}\u00D7${latentMetrics.latentHeight}`;
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
  previewCaptionToggle.textContent = state.previewCaptionOverlayCollapsed ? "+" : "-";
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

