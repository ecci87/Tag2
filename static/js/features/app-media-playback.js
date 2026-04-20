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
  videoMuteBtn.textContent = isMuted ? "\uD83D\uDD07" : "\uD83D\uDD0A";
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

async function stepPreviewVideoFrames(frameDelta) {
  if (state.previewMediaType !== "video" || !state.previewPath || previewVideo.readyState < 1) {
    return false;
  }

  const step = Math.sign(Number(frameDelta || 0));
  if (!step) {
    return false;
  }

  const path = state.previewPath;
  let meta = getCurrentVideoMeta(path);
  if (!meta) {
    try {
      meta = await ensureVideoMetaLoaded(path);
    } catch {
      meta = null;
    }
  }

  const fps = Math.max(1, Number(meta?.fps || getEffectiveVideoMaskFps(path) || 0));
  const duration = Math.max(0, Number(meta?.duration || previewVideo.duration || 0));
  const currentFrameIndex = Math.max(0, Math.floor((Math.max(0, Number(previewVideo.currentTime || 0)) * fps) + 1e-6));
  const maxFrameIndex = duration > 0
    ? Math.max(0, Math.floor((Math.max(0, duration - (1 / fps)) * fps) + 1e-6))
    : Math.max(0, currentFrameIndex + step);
  const nextFrameIndex = Math.max(0, Math.min(maxFrameIndex, currentFrameIndex + step));

  if (!previewVideo.paused) {
    previewVideo.pause();
  }
  seekPreviewVideoTo(nextFrameIndex / fps);
  return true;
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
  const hasDuration = getEffectiveVideoDuration(path) > 0;
  updateVideoActionButtons(path, { hasDuration });
  videoTimelineStartHandle.disabled = !hasDuration;
  videoTimelineEndHandle.disabled = !hasDuration;
  videoTimelinePlayhead.disabled = !hasDuration;
  videoTimelineViewport.classList.toggle("disabled", !hasDuration);
  renderVideoTrainingSummary();
}

function updateVideoActionButtons(path = state.previewPath, options = {}) {
  const { hasDuration: hasDurationOption = null } = options;
  const videoVisible = state.previewMediaType === "video" && !!state.previewPath && state.selectedPaths.size === 1 && !!path;
  const hasDuration = hasDurationOption == null ? getEffectiveVideoDuration(path) > 0 : !!hasDurationOption;
  const hasBlockingAction = !!(
    state.duplicatingImage
    || state.extractingFrame
    || state.autoCaptioning
    || state.cloning
    || state.moving
    || state.uploading
  );
  const maskEditing = !!state.maskEditor.active;

  videoClipBtn.disabled = !videoVisible || !hasDuration || hasBlockingAction || maskEditing;
  videoDownloadBtn.disabled = !videoVisible || state.extractingFrame;
  videoExtractFrameBtn.disabled = !videoVisible || hasBlockingAction || maskEditing;
  videoExtractFrameBtn.textContent = state.extractingFrame ? "Extracting..." : "Extract Frame";
  videoExtractFrameBtn.title = maskEditing
    ? "Finish mask editing before extracting a frame"
    : "Save the current video frame as a JPG beside the video and copy caption, metadata, and the active key-frame mask when available";
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
    updateVideoActionButtons(null, { hasDuration: false });
    syncPreviewVideoPlaybackState();
    return;
  }
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

async function extractCurrentVideoFrame() {
  if (!state.previewPath || !isVideoMediaPath(state.previewPath) || state.selectedPaths.size !== 1) {
    showErrorToast("Select a single video first.");
    return;
  }
  if (state.duplicatingImage || state.extractingFrame || state.autoCaptioning || state.cloning || state.moving || state.uploading || state.maskEditor.active) {
    showErrorToast("Finish the current operation before extracting a frame.");
    return;
  }

  const sourcePath = state.previewPath;
  const timeSeconds = Math.max(0, Number(previewVideo.currentTime || 0));
  let frameIndex = typeof getCurrentVideoMaskFrameIndex === "function"
    ? getCurrentVideoMaskFrameIndex(sourcePath)
    : 0;
  try {
    await ensureVideoMetaLoaded(sourcePath);
  } catch {
    // Frame extraction still works without video metadata; this only affects mask resolution.
  }
  if (typeof getResolvedVideoMaskKeyframeForFrame === "function") {
    const resolvedMaskFrameIndex = getResolvedVideoMaskKeyframeForFrame(sourcePath, frameIndex, { fallbackToCurrent: false });
    if (resolvedMaskFrameIndex != null) {
      frameIndex = resolvedMaskFrameIndex;
    }
  }

  state.extractingFrame = true;
  updateActionButtons();
  statusBar.textContent = `Extracting frame from ${getFileLabel(sourcePath)}...`;
  try {
    const resp = await fetch("/api/video/extract-frame", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        video_path: sourcePath,
        time_seconds: timeSeconds,
        frame_index: frameIndex,
      }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to extract video frame");
    }
    if (!data.image_path) {
      throw new Error("Frame extraction finished without an output image path");
    }

    const preserveScrollTop = fileGridContainer.scrollTop;
    await loadFolder({ preserveScrollTop });
    await selectUploadedImages([data.image_path]);

    const copiedParts = [];
    if (data.caption_copied) copiedParts.push("caption");
    if (data.metadata_copied) copiedParts.push("metadata");
    if (data.mask_copied) copiedParts.push("mask");
    statusBar.textContent = copiedParts.length > 0
      ? `Extracted ${getFileLabel(data.image_path)} at ${formatDurationSeconds(data.time_seconds || timeSeconds)} with ${copiedParts.join(", ")}`
      : `Extracted ${getFileLabel(data.image_path)} at ${formatDurationSeconds(data.time_seconds || timeSeconds)}`;
  } catch (err) {
    const message = err?.message || "Failed to extract video frame";
    statusBar.textContent = `Frame extract error: ${message}`;
    showErrorToast(`Frame extract error: ${message}`);
  } finally {
    state.extractingFrame = false;
    updateActionButtons();
  }
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
  videoJobText.textContent = parts.join(" \u2022 ");
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
  if (state.autoCaptioning || state.cloning || state.moving || state.extractingFrame || state.uploading) {
    showErrorToast("Finish the current operation before deleting media.");
    return;
  }

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

