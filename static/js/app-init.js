function setupResize(handleEl, panelEl, side) {
  let startX, startW;
  handleEl.addEventListener("mousedown", (e) => {
    e.preventDefault();
    startX = e.clientX;
    startW = panelEl.offsetWidth;
    handleEl.classList.add("active");
    const onMove = (e2) => {
      const dx = side === "left" ? (e2.clientX - startX) : (startX - e2.clientX);
      const newW = Math.max(200, Math.min(window.innerWidth * 0.5, startW + dx));
      panelEl.style.width = newW + "px";
    };
    const onUp = () => {
      handleEl.classList.remove("active");
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  });
}

function setupVerticalResize(handleEl, topEl, bottomEl, containerEl) {
  let startY, startTopHeight;
  handleEl.addEventListener("mousedown", (e) => {
    e.preventDefault();
    startY = e.clientY;
    startTopHeight = topEl.getBoundingClientRect().height;
    handleEl.classList.add("active");

    const onMove = (e2) => {
      const containerHeight = containerEl.getBoundingClientRect().height;
      const handleHeight = handleEl.getBoundingClientRect().height;
      const minTop = 120;
      const minBottom = 120;
      const maxTop = Math.max(minTop, containerHeight - minBottom - handleHeight);
      const nextTop = Math.max(minTop, Math.min(maxTop, startTopHeight + (e2.clientY - startY)));
      topEl.style.flex = `0 0 ${nextTop}px`;
      bottomEl.style.flex = "1 1 auto";
    };

    const onUp = () => {
      handleEl.classList.remove("active");
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    };

    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  });
}

function invokeGlobalFunction(functionName, ...args) {
  const fn = globalThis[functionName];
  if (typeof fn !== "function") {
    console.error(`${functionName} is not defined`);
    return undefined;
  }
  return fn(...args);
}

function addSafeClickListener(element, functionName) {
  if (!element) return;
  element.addEventListener("click", () => {
    invokeGlobalFunction(functionName);
  });
}

function handleMetadataEditorInputChange() {
  if (typeof updateMetadataEditorDirtyState === "function") {
    updateMetadataEditorDirtyState();
  }
}

function handleMetadataEditorFieldKeydown(event) {
  if (event.key !== "Enter") return;
  event.preventDefault();
  invokeGlobalFunction("saveMetadataForSelection");
}

setupResize($("#left-resize"), $("#left-panel"), "left");
setupResize($("#right-resize"), $("#right-panel"), "right");
setupVerticalResize(rightHorizontalResize, captionsSection, freeTextSection, captionsEditorPanel);

// ===== EVENT LISTENERS =====
loadBtn.addEventListener("click", loadFolder);
cloneFolderBtn.addEventListener("click", cloneCurrentFolder);
moveSelectedBtn.addEventListener("click", openMoveSelectedDialog);
folderInput.addEventListener("input", handleFolderInputInput);
folderInput.addEventListener("focus", handleFolderInputFocus);
folderInput.addEventListener("blur", handleFolderInputBlur);
folderInput.addEventListener("keydown", handleFolderInputKeydown);
moveSelectedForm.addEventListener("submit", submitMoveSelectedDialog);
moveSelectedCloseBtn.addEventListener("click", closeMoveSelectedDialog);
moveSelectedCancelBtn.addEventListener("click", closeMoveSelectedDialog);
moveTargetFolderInput.addEventListener("input", handleMoveTargetFolderInput);
moveTargetFolderInput.addEventListener("focus", handleMoveTargetFolderFocus);
moveTargetFolderInput.addEventListener("blur", handleMoveTargetFolderBlur);
moveTargetFolderInput.addEventListener("keydown", handleMoveTargetFolderKeydown);
settingsBtn.addEventListener("click", openSettingsModal);
settingsRefreshModelsBtn.addEventListener("click", refreshOllamaModelOptions);
clearFiltersBtn.addEventListener("click", clearSentenceFilters);
filterArBtn.addEventListener("click", toggleAspectFilter);
filterMaskBtn.addEventListener("click", toggleMaskPresenceFilter);
filterTxtBtn.addEventListener("click", toggleCaptionPresenceFilter);
if (mediaSearchInput) {
  mediaSearchInput.addEventListener("input", handleMediaSearchInput);
  mediaSearchInput.addEventListener("search", handleMediaSearchInput);
}
if (promptPreviewGridToggleBtn) {
  promptPreviewGridToggleBtn.addEventListener("click", () => {
    state.showPromptPreviewThumbnails = !state.showPromptPreviewThumbnails;
    renderPromptPreviewGridToggle();
    if (typeof rerenderGridPreservingView === "function") {
      rerenderGridPreservingView();
    } else if (typeof renderGrid === "function") {
      renderGrid();
    }
    statusBar.textContent = state.showPromptPreviewThumbnails
      ? "Showing prompt preview thumbnails when available"
      : "Showing source thumbnails only";
  });
}
addSafeClickListener(autoCaptionBtn, "autoCaptionSelected");
addSafeClickListener(addFreeTextNowBtn, "addFreeTextNow");
addSafeClickListener(metadataCancelBtn, "cancelMetadataChanges");
addSafeClickListener(metadataSaveBtn, "saveMetadataForSelection");
addSafeClickListener(videoClipBtn, "queueCurrentVideoClip");
addSafeClickListener(videoExtractFrameBtn, "extractCurrentVideoFrame");
addSafeClickListener(gifConvertBtn, "queueCurrentGifConversion");
addSafeClickListener(videoDownloadBtn, "downloadCurrentVideo");
[metadataSeedInput, metadataSamplingFrequencyInput, metadataMinTInput, metadataMaxTInput].forEach((inputEl) => {
  if (!inputEl) return;
  inputEl.addEventListener("input", handleMetadataEditorInputChange);
  inputEl.addEventListener("change", handleMetadataEditorInputChange);
  inputEl.addEventListener("keydown", handleMetadataEditorFieldKeydown);
});
if (metadataCaptionDropoutInput) {
  metadataCaptionDropoutInput.addEventListener("input", handleMetadataEditorInputChange);
  metadataCaptionDropoutInput.addEventListener("change", handleMetadataEditorInputChange);
  metadataCaptionDropoutInput.addEventListener("keydown", (event) => {
    if (!(event.ctrlKey || event.metaKey) || event.key !== "Enter") return;
    event.preventDefault();
    invokeGlobalFunction("saveMetadataForSelection");
  });
}
if (metadataCaptionDropoutEnabledInput) {
  metadataCaptionDropoutEnabledInput.addEventListener("change", handleMetadataEditorInputChange);
}
if (captionSearchInput) {
  captionSearchInput.addEventListener("input", handleCaptionSearchInput);
  captionSearchInput.addEventListener("search", handleCaptionSearchInput);
}
hideAddButtonsCheckbox.addEventListener("change", () => {
  state.hideAddButtons = hideAddButtonsCheckbox.checked;
  renderSentences();
});
autoFreeTextCheckbox.addEventListener("change", async () => {
  state.ollamaEnableFreeText = autoFreeTextCheckbox.checked;
  try {
    await fetch("/api/settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ollama_enable_free_text: state.ollamaEnableFreeText }),
    });
  } catch (err) {
    console.error("Failed to save free-text toggle:", err);
  }
});
autoPreviewCheckbox.addEventListener("change", async () => {
  state.comfyuiAutoPreviewEnabled = autoPreviewCheckbox.checked;
  try {
    await fetch("/api/settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ comfyui_auto_preview: state.comfyuiAutoPreviewEnabled }),
    });
  } catch (err) {
    console.error("Failed to save auto-preview toggle:", err);
  }
});
settingsCloseBtn.addEventListener("click", closeSettingsModal);
settingsCancelBtn.addEventListener("click", closeSettingsModal);
settingsSaveBtn.addEventListener("click", saveOllamaSettingsFromForm);
settingsVideoPresetsInput.addEventListener("input", handleVideoTrainingPresetsInput);
settingsVideoProfileInput.addEventListener("change", handleVideoTrainingProfileInputChange);
settingsTabButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setActiveSettingsTab(button.dataset.settingsTab || "auto-captioning");
  });
  button.addEventListener("keydown", (e) => {
    if (e.key !== "ArrowLeft" && e.key !== "ArrowRight" && e.key !== "Home" && e.key !== "End") return;
    e.preventDefault();
    const currentIndex = settingsTabButtons.indexOf(button);
    if (currentIndex < 0) return;
    let nextIndex = currentIndex;
    if (e.key === "ArrowRight") nextIndex = (currentIndex + 1) % settingsTabButtons.length;
    if (e.key === "ArrowLeft") nextIndex = (currentIndex - 1 + settingsTabButtons.length) % settingsTabButtons.length;
    if (e.key === "Home") nextIndex = 0;
    if (e.key === "End") nextIndex = settingsTabButtons.length - 1;
    const nextButton = settingsTabButtons[nextIndex];
    setActiveSettingsTab(nextButton.dataset.settingsTab || "auto-captioning");
    nextButton.focus();
  });
});
rightPanelTabButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setActiveRightPanelTab(button.dataset.rightPanelTab || "captions");
  });
  button.addEventListener("keydown", (e) => {
    if (e.key !== "ArrowLeft" && e.key !== "ArrowRight" && e.key !== "Home" && e.key !== "End") return;
    e.preventDefault();
    const currentIndex = rightPanelTabButtons.indexOf(button);
    if (currentIndex < 0) return;
    let nextIndex = currentIndex;
    if (e.key === "ArrowRight") nextIndex = (currentIndex + 1) % rightPanelTabButtons.length;
    if (e.key === "ArrowLeft") nextIndex = (currentIndex - 1 + rightPanelTabButtons.length) % rightPanelTabButtons.length;
    if (e.key === "Home") nextIndex = 0;
    if (e.key === "End") nextIndex = rightPanelTabButtons.length - 1;
    const nextButton = rightPanelTabButtons[nextIndex];
    setActiveRightPanelTab(nextButton.dataset.rightPanelTab || "captions");
    nextButton.focus();
  });
});
modelLogOpenBtn.addEventListener("click", toggleModelLogOverlay);
previewCaptionToggle.addEventListener("click", (e) => {
  e.stopPropagation();
  togglePreviewCaptionOverlayCollapsed();
});
modelLogClearBtn.addEventListener("click", clearModelLog);
modelLogCloseBtn.addEventListener("click", closeModelLogOverlay);
settingsForm.addEventListener("submit", (e) => {
  e.preventDefault();
  saveOllamaSettingsFromForm();
});
modelLogOverlay.addEventListener("click", (e) => {
  if (e.target === modelLogOverlay) closeModelLogOverlay();
});
settingsModal.addEventListener("click", (e) => {
  if (e.target === settingsModal) closeSettingsModal();
});
moveSelectedModal.addEventListener("click", (e) => {
  if (e.target === moveSelectedModal) closeMoveSelectedDialog();
});
document.addEventListener("mousedown", (e) => {
  if (folderInputWrap?.contains(e.target)) return;
  if (moveTargetFolderInputWrap?.contains(e.target)) return;
  hideFolderAutocomplete();
  hideMoveFolderAutocomplete();
});
renderAddButtonsVisibility();
initializeFilterButtons();

function scrollPreviewSelectionIntoView(path) {
  if (!path) return;
  const cell = fileGrid.querySelector(`.thumb-cell[data-path="${CSS.escape(path)}"]`)
    || fileGrid.querySelector(`[data-path="${CSS.escape(path)}"]`);
  if (cell) {
    cell.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }
}

async function navigateVisiblePreviewSelection(direction, options = {}) {
  const { imagesOnly = false } = options;
  const step = direction > 0 ? 1 : -1;
  const visibleEntries = getVisibleImageEntries().filter(({ img }) => !imagesOnly || img.media_type === "image");
  if (!visibleEntries.length) {
    return false;
  }

  const anchorPath = state.previewPath || state.lastClickedPath || [...state.selectedPaths][0] || null;
  const currentIndex = anchorPath
    ? visibleEntries.findIndex(({ img }) => img.path === anchorPath)
    : -1;
  let nextIndex = currentIndex + step;
  if (currentIndex < 0) {
    nextIndex = step > 0 ? 0 : visibleEntries.length - 1;
  }
  nextIndex = Math.max(0, Math.min(visibleEntries.length - 1, nextIndex));
  if (nextIndex === currentIndex) {
    return false;
  }

  const nextPath = visibleEntries[nextIndex]?.img?.path;
  if (!nextPath) {
    return false;
  }

  await selectUploadedImages([nextPath]);
  scrollPreviewSelectionIntoView(nextPath);
  return true;
}

// Keyboard shortcuts
document.addEventListener("keydown", (e) => {
  if (state.maskEditor.active && !isEditableElement(document.activeElement) && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") {
    e.preventDefault();
    if (e.shiftKey) {
      redoMaskEdit();
    } else {
      undoMaskEdit();
    }
    return;
  }
  if (state.maskEditor.active && !isEditableElement(document.activeElement) && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "y") {
    e.preventDefault();
    redoMaskEdit();
    return;
  }
  if (e.key === "Escape" && modelLogOverlay.classList.contains("open")) {
    closeModelLogOverlay();
    return;
  }
  if (e.key === "Escape" && settingsModal.classList.contains("open")) {
    closeSettingsModal();
    return;
  }
  if (e.key === "Escape" && moveSelectedModal.classList.contains("open")) {
    closeMoveSelectedDialog();
    return;
  }
  if (e.key === "Escape" && state.maskEditor.active) {
    e.preventDefault();
    cancelMaskEdit();
    return;
  }
  if (e.key === "Escape" && (state.cropDraft || state.cropInteraction)) {
    e.preventDefault();
    cancelCropEdit();
    return;
  }
  if (settingsModal.classList.contains("open") || moveSelectedModal.classList.contains("open")) return;
  if (e.key === " " && !isEditableElement(document.activeElement) && state.previewMediaType === "video" && state.previewPath) {
    e.preventDefault();
    togglePreviewVideoPlayback();
    return;
  }
  if (e.key === "Delete" && !isEditableElement(document.activeElement)) {
    e.preventDefault();
    deleteSelectedImages();
    return;
  }
  // Ctrl+A to select all (when not in text input)
  if ((e.ctrlKey || e.metaKey) && e.key === "a" && !isEditableElement(document.activeElement)) {
    e.preventDefault();
    const visibleEntries = getVisibleImageEntries();
    const nextPaths = visibleEntries.map(({ img }) => img.path).filter(Boolean);
    if (nextPaths.length > 0) {
      selectUploadedImages(nextPaths).catch((err) => {
        showErrorToast(err?.message || "Failed to select visible media");
      });
    }
    return;
  }

  // Arrow keys for navigation
  if (e.key === "ArrowRight" || e.key === "ArrowLeft" || e.key === "ArrowUp" || e.key === "ArrowDown") {
    const activeEl = document.activeElement;
    if (activeEl === freeText || activeEl === folderInput || isEditableElement(activeEl)) return;
    const dir = (e.key === "ArrowRight" || e.key === "ArrowDown") ? 1 : -1;

    if ((e.key === "ArrowRight" || e.key === "ArrowLeft") && state.previewMediaType === "video" && state.previewPath) {
      e.preventDefault();
      stepPreviewVideoFrames(dir).catch((err) => {
        showErrorToast(err?.message || "Failed to step video frame");
      });
      return;
    }

    if (state.previewMediaType === "image" && state.previewPath) {
      e.preventDefault();
      navigateVisiblePreviewSelection(dir, { imagesOnly: true }).catch((err) => {
        showErrorToast(err?.message || "Failed to change preview");
      });
      return;
    }

    e.preventDefault();
    const visibleEntries = getVisibleImageEntries();
    if (visibleEntries.length === 0) return;
    let currentIdx = getVisibleImageIndexByPath(state.lastClickedPath);
    if (currentIdx < 0) {
      currentIdx = getVisibleImageIndexByPath(state.previewPath);
    }
    let newIdx = currentIdx + dir;
    if (currentIdx < 0) {
      newIdx = dir > 0 ? 0 : visibleEntries.length - 1;
    }
    if (newIdx < 0) newIdx = 0;
    if (newIdx >= visibleEntries.length) newIdx = visibleEntries.length - 1;
    if (newIdx >= 0 && newIdx < visibleEntries.length) {
      const nextEntry = visibleEntries[newIdx];
      handleThumbClick(nextEntry.index, e);
      // Scroll into view
      const cell = fileGrid.querySelector(`[data-path="${CSS.escape(nextEntry.img.path)}"]`);
      if (cell) cell.scrollIntoView({ block: "nearest", behavior: "smooth" });
    }
  }
});

// Load settings from server on startup
async function initFromSettings() {
  try {
    const resp = await fetch("/api/settings");
    if (resp.ok) {
      const settings = await resp.json();
      if (settings.last_folder) {
        folderInput.value = settings.last_folder;
      }
      applySettings(settings);
    }
  } catch (err) {
    console.error("Failed to load settings:", err);
  }
}
initFromSettings();
startVideoJobPolling();
startPromptPreviewPolling();

// Initial render of sentences
setCropAspectRatios(state.cropAspectRatioLabels);
if (typeof renderSentences === "function") renderSentences();
setActiveRightPanelTab(state.ui.activeRightPanelTab);
renderMetadataEditor();
renderMaskEditorUi();
updateActionButtons();
renderFilterActions();
renderPromptPreviewGridToggle();
