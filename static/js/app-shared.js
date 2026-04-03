const state = {
  folder: "",
  images: [],            // [{name, path, size, mtime}, ...]
  imageCrops: {},        // path -> {x, y, w, h, ratio}
  imageVersions: {},     // path -> cache-busting token
  imageMaskVersions: {}, // path -> cache-busting token for image sidecar masks
  thumbnailDimensions: {}, // path -> {width, height}
  selectedPaths: new Set(),
  lastClickedIndex: -1,
  lastClickedPath: null,
  sections: [],          // [{name: "", sentences: [...], groups: [{name: "Car", sentences: [...]}]}, ...]
  // Per-image caption cache: path -> {enabled_sentences: [...], free_text: ""}
  captionCache: {},
  activeSentenceFilters: new Map(),
  activeMetaFilters: {
    aspectState: "any",
    maskState: "any",
    captionState: "any",
  },
  filterCaptionCacheKey: "",
  filterLoadingPromise: null,
  previewPath: null,
  previewMediaType: null,
  previewVideoMuted: false,
  previewVideoLastVolume: 1,
  previewVideoVolume: 1,
  videoMeta: {},         // path -> {width, height, duration}
  videoClipDrafts: {},   // path -> {startFraction, endFraction}
  videoTimelineCache: {}, // path -> [{timeSeconds, url}, ...]
  videoTimelineUi: {},   // path -> {zoom, offsetFraction}
  saving: false,
  autoCaptioning: false,
  autoCaptionMode: null,
  autoCaptionAbortController: null,
  cropAspectRatioLabels: ["4:3", "16:9", "3:4", "1:1", "9:16", "2:3", "3:2"],
  cropAspectRatios: [],
  cropDraft: null,
  cropGuide: null,
  cropDirty: false,
  cropInteraction: null,
  maskEditor: {
    active: false,
    loading: false,
    saving: false,
    dirty: false,
    viewMode: "overlay",
    latentPreviewEnabled: false,
    latentBaseWidth: 512,
    latentDivider: 16,
    latentSignalPercent: 50,
    latentReductionPercent: 50,
    latentPreviewQueued: false,
    latentImageDirty: false,
    latentBaseMaskCanvas: null,
    latentGridCanvas: null,
    history: [],
    historyIndex: -1,
    cleanHistoryIndex: -1,
    path: null,
    imageWidth: 0,
    imageHeight: 0,
    previewScaleX: 1,
    previewScaleY: 1,
    brushSizePercent: 6,
    brushValue: 100,
    brushCore: 30,
    brushSteepness: 8,
    painting: false,
    lastPoint: null,
    baseCanvas: null,
    strokeBaseCanvas: null,
    strokeInfluenceValues: null,
    previewQueued: false,
    strokeRenderFrameId: 0,
    strokeDirtyRect: null,
  },
  collapsedSections: {},
  collapsedGroups: {},
  httpsCertFile: "",
  httpsKeyFile: "",
  httpsPort: 8900,
  remoteHttpMode: "redirect-to-https",
  ffmpegPath: "",
  ffmpegThreads: 0,
  ffmpegHwaccel: "auto",
  processingReservedCores: 4,
  ollamaServer: "127.0.0.1",
  ollamaPort: 11434,
  ollamaTimeoutSeconds: 20,
  ollamaModel: "llava",
  ollamaPromptTemplate: "You are verifying a caption for one media item. Reply with exactly one word: YES or NO. Reply YES only if the caption is clearly correct for the media. Reply NO if it is wrong, uncertain, too specific, or not clearly visible.\n\nCaption: {caption}\nAnswer:",
  ollamaGroupPromptTemplate: "You are selecting the single best caption for one media item from a numbered list. Reply with exactly one number from 1 to {count}. Pick the most likely correct caption for the media.\n\nGroup: {group_name}\n{options}\n\nAnswer:",
  ollamaEnableFreeText: true,
  ollamaFreeTextPromptTemplate: "You are improving a media caption file. The caption text below already covers known details and must not be repeated. Look at the media and return only notable, important visual details that are still missing. Return either NONE or one short line per missing detail, with no bullets or numbering.\n\nCurrent caption text:\n{caption_text}\n\nAnswer:",
  ollamaAvailableModels: [],
  modelLogLines: [],
  modelLogOpen: false,
  cloning: false,
  aiProgress: {
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
  },
  previewCaptionOverlayCollapsed: false,
  hideAddButtons: false,
  thumbSize: 160,
  uploading: false,
  thumbnailProgress: {
    visible: false,
    label: "",
    total: 0,
    completed: 0,
  },
  uploadQueue: [],
  uploadQueueCurrentJob: null,
  uploadQueueCompletedImages: 0,
  uploadQueueFailedJobs: 0,
  uploadQueueLastSummary: "",
  videoJobs: {
    activeJob: null,
    queuedJobs: [],
    recentJobs: [],
    summary: { total: 0, completed: 0, failed: 0, queued: 0, running: 0 },
    seenFinishedIds: new Set(),
    batch: {
      active: false,
      total: 0,
      completed: 0,
      jobIds: new Set(),
    },
  },
  ui: {
    activeInlineEditor: null,
    renderFrameId: 0,
    pendingSentenceRender: false,
    pendingPreviewRender: false,
    videoJobPollTimer: null,
    suppressVideoClick: false,
    videoTimelineFetches: new Map(),
  },
};

const IMAGE_FILE_EXTENSIONS = new Set([".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff", ".tif"]);
const VIDEO_FILE_EXTENSIONS = new Set([".mp4", ".m4v", ".mov", ".webm", ".mkv", ".avi"]);
const MEDIA_FILE_EXTENSIONS = new Set([...IMAGE_FILE_EXTENSIONS, ...VIDEO_FILE_EXTENSIONS]);
let uploadJobCounter = 0;
let thumbnailProgressHideTimer = null;

// ===== DOM REFS =====
const $ = (s) => document.querySelector(s);
const folderInput = $("#folder-input");
const loadBtn = $("#load-btn");
const cloneFolderBtn = $("#clone-folder-btn");
const settingsBtn = $("#settings-btn");
const thumbSlider = $("#thumb-size-slider");
const statusBar = $("#status-bar");
const thumbnailQueueStatus = $("#thumbnail-queue-status");
const thumbnailQueueText = $("#thumbnail-queue-text");
const thumbnailQueueProgressFill = $("#thumbnail-queue-progress-fill");
const uploadQueueStatus = $("#upload-queue-status");
const uploadQueueText = $("#upload-queue-text");
const uploadQueueProgressFill = $("#upload-queue-progress-fill");
const videoJobStatus = $("#video-job-status");
const videoJobText = $("#video-job-text");
const videoJobProgressFill = $("#video-job-progress-fill");
const fileCount = $("#file-count");
const filterArBtn = $("#filter-ar-btn");
const filterMaskBtn = $("#filter-mask-btn");
const filterTxtBtn = $("#filter-txt-btn");
const clearFiltersBtn = $("#clear-filters-btn");
const fileGrid = $("#file-grid");
const fileGridContainer = $("#file-grid-container");
const centerPanel = $("#center-panel");
const previewStage = $("#preview-stage");
const previewImg = $("#preview-img");
const previewVideo = $("#preview-video");
const previewMaskCanvas = $("#preview-mask-canvas");
const previewLatentImageCanvas = $("#preview-latent-image-canvas");
const previewLatentMaskCanvas = $("#preview-latent-mask-canvas");
const fileDropHint = $("#file-drop-hint");
const maskEditorPanel = $("#mask-editor-panel");
const maskEditorStatus = $("#mask-editor-status");
const maskBrushSizeInput = $("#mask-brush-size");
const maskBrushSizeLabel = $("#mask-brush-size-label");
const maskBrushValueInput = $("#mask-brush-value");
const maskBrushValueLabel = $("#mask-brush-value-label");
const maskBrushCoreInput = $("#mask-brush-core");
const maskBrushCoreLabel = $("#mask-brush-core-label");
const maskBrushSteepnessInput = $("#mask-brush-steepness");
const maskBrushSteepnessLabel = $("#mask-brush-steepness-label");
const maskLatentPreviewControls = $("#mask-latent-preview-controls");
const maskLatentBaseWidthInput = $("#mask-latent-base-width");
const maskLatentBaseWidthLabel = $("#mask-latent-base-width-label");
const maskLatentDividerInput = $("#mask-latent-divider");
const maskLatentDividerLabel = $("#mask-latent-divider-label");
const maskLatentBaseSizeLabel = $("#mask-latent-base-size-label");
const maskLatentGridSizeLabel = $("#mask-latent-grid-size-label");
const maskLatentSignalLabel = $("#mask-latent-signal-label");
const maskLatentReductionLabel = $("#mask-latent-reduction-label");
const maskMiniPreview = $("#mask-mini-preview");
const maskCursor = $("#mask-cursor");
const cropOverlay = $("#crop-overlay");
const cropGuideV = $("#crop-guide-v");
const cropGuideH = $("#crop-guide-h");
const cropBox = $("#crop-box");
const cropLabel = cropBox.querySelector(".crop-label");
const cropApplyBtn = $("#crop-apply-btn");
const cropCancelBtn = $("#crop-cancel-btn");
const cropRemoveBtn = $("#crop-remove-btn");
const maskEditBtn = $("#mask-edit-btn");
const maskActionBar = $("#mask-action-bar");
const maskApplyBtn = $("#mask-apply-btn");
const maskCancelBtn = $("#mask-cancel-btn");
const maskUndoBtn = $("#mask-undo-btn");
const maskRedoBtn = $("#mask-redo-btn");
const maskViewModeBtn = $("#mask-view-mode-btn");
const maskLatentPreviewBtn = $("#mask-latent-preview-btn");
const maskResetBtn = $("#mask-reset-btn");
const rotateControls = $("#rotate-controls");
const rotateLeftBtn = $("#rotate-left-btn");
const rotateRightBtn = $("#rotate-right-btn");
const videoEditPanel = $("#video-edit-panel");
const videoPlayToggleBtn = $("#video-play-toggle-btn");
const videoMuteBtn = $("#video-mute-btn");
const videoVolumeSlider = $("#video-volume-slider");
const videoPlaybackLabel = $("#video-playback-label");
const videoClipBtn = $("#video-clip-btn");
const videoDownloadBtn = $("#video-download-btn");
const videoTimelineViewport = $("#video-timeline-viewport");
const videoTimeRangeLabel = $("#video-time-range-label");
const videoTimelineZoomLabel = $("#video-timeline-zoom-label");
const videoTimelineStrip = $("#video-timeline-strip");
const videoTimelineOverlay = $("#video-timeline-overlay");
const videoTimelineSelection = $("#video-timeline-selection");
const videoTimelineStartHandle = $("#video-timeline-start-handle");
const videoTimelineEndHandle = $("#video-timeline-end-handle");
const videoTimelinePlayhead = $("#video-timeline-playhead");
const previewPlaceholder = $("#preview-placeholder");
const previewInfo = $("#preview-info");
const previewCaptionOverlay = $("#preview-caption-overlay");
const previewCaptionToggle = $("#preview-caption-toggle");
const previewCaptionList = $("#preview-caption-list");
const sectionContainer = $("#section-container");
const sentencesSection = $("#sentences-section");
const rightHorizontalResize = $("#right-horizontal-resize");
const autoCaptionBtn = $("#auto-caption-btn");
const modelLogOpenBtn = $("#model-log-open-btn");
const autoFreeTextCheckbox = $("#auto-free-text-checkbox");
const addFreeTextNowBtn = $("#add-free-text-now-btn");
const hideAddButtonsCheckbox = $("#hide-add-buttons-checkbox");
const freeText = $("#free-text");
const multiInfo = $("#multi-info");
const aiProgressPanel = $("#ai-progress-panel");
const aiProgressSummary = $("#ai-progress-summary");
const aiProgressMetric = $("#ai-progress-metric");
const aiProgressOverallFill = $("#ai-progress-overall-fill");
const aiProgressCurrentLabel = $("#ai-progress-current-label");
const aiProgressCurrentMetric = $("#ai-progress-current-metric");
const aiProgressCurrentFill = $("#ai-progress-current-fill");
const modelLogOverlay = $("#model-log-overlay");
const modelLogDialog = $("#model-log-dialog");
const modelLog = $("#model-log");
const modelLogClearBtn = $("#model-log-clear-btn");
const modelLogCloseBtn = $("#model-log-close-btn");
const errorToast = $("#error-toast");
const errorToastMessage = $("#error-toast-message");
const errorToastClose = $("#error-toast-close");
const settingsModal = $("#settings-modal");
const settingsForm = $("#settings-form");
const settingsCloseBtn = $("#settings-close-btn");
const settingsCancelBtn = $("#settings-cancel-btn");
const settingsSaveBtn = $("#settings-save-btn");
const settingsTabButtons = [...document.querySelectorAll(".settings-tab-btn")];
const settingsPanels = [...document.querySelectorAll(".settings-panel")];
const settingsServerInput = $("#settings-ollama-server");
const settingsPortInput = $("#settings-ollama-port");
const settingsTimeoutInput = $("#settings-ollama-timeout");
const settingsModelInput = $("#settings-ollama-model");
const settingsRefreshModelsBtn = $("#settings-refresh-models-btn");
const settingsCropAspectRatiosInput = $("#settings-crop-aspect-ratios");
const settingsHttpsCertInput = $("#settings-https-certfile");
const settingsHttpsKeyInput = $("#settings-https-keyfile");
const settingsHttpsPortInput = $("#settings-https-port");
const settingsRemoteHttpModeInput = $("#settings-remote-http-mode");
const settingsFfmpegPathInput = $("#settings-ffmpeg-path");
const settingsProcessingReservedCoresInput = $("#settings-processing-reserved-cores");
const settingsFfmpegThreadsInput = $("#settings-ffmpeg-threads");
const settingsFfmpegHwaccelInput = $("#settings-ffmpeg-hwaccel");
const settingsPromptInput = $("#settings-ollama-prompt");
const settingsGroupPromptInput = $("#settings-ollama-group-prompt");
const settingsAutoFreeTextEnabled = $("#settings-auto-free-text-enabled");
const settingsFreeTextPromptInput = $("#settings-ollama-free-text-prompt");
const sentenceListElements = new Map();
let errorToastTimer = null;

function hideErrorToast() {
  if (errorToastTimer) {
    clearTimeout(errorToastTimer);
    errorToastTimer = null;
  }
  errorToast.classList.remove("visible");
}

function showErrorToast(message, options = {}) {
  const text = String(message || "Unexpected error").trim();
  if (!text) return;
  const { autoHideMs = 0 } = options;
  errorToastMessage.textContent = text;
  errorToast.classList.add("visible");
  if (errorToastTimer) clearTimeout(errorToastTimer);
  if (autoHideMs > 0) {
    errorToastTimer = window.setTimeout(() => {
      errorToast.classList.remove("visible");
      errorToastTimer = null;
    }, autoHideMs);
  }
}

errorToastClose.addEventListener("click", hideErrorToast);

window.addEventListener("error", (event) => {
  const message = event?.error?.message || event?.message || "Unexpected error";
  showErrorToast(message);
});

window.addEventListener("unhandledrejection", (event) => {
  const reason = event?.reason;
  const message = reason?.message || (typeof reason === "string" ? reason : "Unexpected async error");
  showErrorToast(message);
});

function makeUiId(prefix) {
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}-${Date.now().toString(36)}`;
}

function createGenerateSparkleIcon() {
  const svgNs = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(svgNs, "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("fill", "none");
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("focusable", "false");

  const star = document.createElementNS(svgNs, "path");
  star.setAttribute("d", "M12 7.1L12.95 9.75L15.6 10.7L12.95 11.65L12 14.3L11.05 11.65L8.4 10.7L11.05 9.75L12 7.1Z");
  star.setAttribute("fill", "#d8b1ff");
  svg.appendChild(star);

  const sparkleA = document.createElementNS(svgNs, "path");
  sparkleA.setAttribute("d", "M16.7 6.1L17.1 7.25L18.25 7.65L17.1 8.05L16.7 9.2L16.3 8.05L15.15 7.65L16.3 7.25L16.7 6.1Z");
  sparkleA.setAttribute("fill", "#aff3ff");
  svg.appendChild(sparkleA);

  const sparkleB = document.createElementNS(svgNs, "path");
  sparkleB.setAttribute("d", "M8 14.75L8.45 16L9.7 16.45L8.45 16.9L8 18.15L7.55 16.9L6.3 16.45L7.55 16L8 14.75Z");
  sparkleB.setAttribute("fill", "#ffc4f0");
  svg.appendChild(sparkleB);

  return svg;
}

function createFilterIcon() {
  const svgNs = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(svgNs, "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("fill", "none");
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("focusable", "false");

  const path = document.createElementNS(svgNs, "path");
  path.setAttribute("d", "M4 6H20L14 13V18L10 20V13L4 6Z");
  path.setAttribute("stroke", "currentColor");
  path.setAttribute("stroke-width", "1.8");
  path.setAttribute("stroke-linejoin", "round");
  path.setAttribute("stroke-linecap", "round");
  svg.appendChild(path);

  return svg;
}

function createAddPlusIcon() {
  const svgNs = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(svgNs, "svg");
  svg.setAttribute("viewBox", "0 0 10 10");
  svg.setAttribute("fill", "none");
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("focusable", "false");
  svg.classList.add("add-btn-plus");

  const vertical = document.createElementNS(svgNs, "path");
  vertical.setAttribute("d", "M5 1.5V8.5");
  vertical.setAttribute("stroke", "currentColor");
  vertical.setAttribute("stroke-width", "1.6");
  vertical.setAttribute("stroke-linecap", "round");
  svg.appendChild(vertical);

  const horizontal = document.createElementNS(svgNs, "path");
  horizontal.setAttribute("d", "M1.5 5H8.5");
  horizontal.setAttribute("stroke", "currentColor");
  horizontal.setAttribute("stroke-width", "1.6");
  horizontal.setAttribute("stroke-linecap", "round");
  svg.appendChild(horizontal);

  return svg;
}

function initializeFilterButtons() {
  if (!clearFiltersBtn) return;
  clearFiltersBtn.replaceChildren(createFilterIcon());
}

function setGenerateButtonContent(button, label, options = {}) {
  if (!button) return;
  const { iconOnly = false } = options;
  button.replaceChildren();

  if (iconOnly) {
    const icon = document.createElement("span");
    icon.className = "btn-icon";
    icon.appendChild(createGenerateSparkleIcon());
    button.appendChild(icon);
  } else {
    const content = document.createElement("span");
    content.className = "btn-content";

    const icon = document.createElement("span");
    icon.className = "btn-icon";
    icon.appendChild(createGenerateSparkleIcon());

    const labelSpan = document.createElement("span");
    labelSpan.className = "btn-label";
    labelSpan.textContent = label || "";

    content.appendChild(icon);
    content.appendChild(labelSpan);
    button.appendChild(content);
  }

  if (iconOnly && label) {
    button.setAttribute("aria-label", label);
  } else if (label) {
    button.removeAttribute("aria-label");
  }
}

