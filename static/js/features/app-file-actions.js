function getCloneSelectionPaths() {
  return state.selectedPaths.size > 1 ? [...state.selectedPaths] : [];
}

function getMoveSelectionPaths() {
  return [...state.selectedPaths];
}

function getDefaultCloneFolderName() {
  const baseName = getFileLabel(state.folder || "folder") || "folder";
  return `${baseName}-copy`;
}

function getDefaultDuplicateImageName(path = state.previewPath) {
  const fileName = getFileLabel(path || "image") || "image";
  const extension = getImageExtension(fileName);
  const stem = extension ? fileName.slice(0, -extension.length) : fileName;
  return `${stem}-copy${extension}`;
}

function getDefaultMoveTargetFolderValue() {
  const rememberedTarget = String(state.moveDialog?.lastTargetFolder || "").trim();
  if (rememberedTarget) return rememberedTarget;

  const folder = String(state.folder || "").trim().replace(/[\\/]+$/, "");
  if (!folder) return "";
  if (/^[A-Za-z]:$/.test(folder)) return `${folder}\\`;

  const separatorIndex = Math.max(folder.lastIndexOf("\\"), folder.lastIndexOf("/"));
  if (separatorIndex < 0) return folder;
  return folder.slice(0, separatorIndex + 1);
}

function isMoveDialogOpen() {
  return !!moveSelectedModal?.classList.contains("open");
}

function setMoveDialogStatus(message = "") {
  const text = String(message || "").trim();
  moveSelectedStatus.textContent = text;
  moveSelectedStatus.hidden = !text;
}

function renderMoveDialogSummary(selectedPaths = getMoveSelectionPaths()) {
  if (!moveSelectedSummary) return;
  const count = selectedPaths.length;
  if (!count) {
    moveSelectedSummary.textContent = "Select one or more media files to copy into another folder.";
    return;
  }
  if (count === 1) {
    moveSelectedSummary.textContent = `Copy ${getFileLabel(selectedPaths[0])} from ${getFileLabel(state.folder || "current folder")} into:`;
    return;
  }
  moveSelectedSummary.textContent = `Copy ${count} selected media files from ${getFileLabel(state.folder || "current folder")} into:`;
}

function hideMoveFolderAutocomplete() {
  hideFolderAutocompleteFor("move-target");
}

function clearMoveFolderAutocomplete(options = {}) {
  clearFolderAutocompleteFor("move-target", options);
}

function scheduleMoveFolderAutocompleteRefresh(options = {}) {
  scheduleFolderAutocompleteRefreshFor("move-target", options);
}

function handleMoveTargetFolderInput() {
  setMoveDialogStatus("");
  handleFolderAutocompleteInputFor("move-target");
}

function handleMoveTargetFolderFocus() {
  handleFolderAutocompleteFocusFor("move-target");
}

function handleMoveTargetFolderBlur() {
  handleFolderAutocompleteBlurFor("move-target");
}

function handleMoveTargetFolderKeydown(event) {
  handleFolderAutocompleteKeydownFor("move-target", event, () => {
    submitMoveSelectedDialog();
  });
}

function openMoveSelectedDialog() {
  if (!state.folder) {
    showErrorToast("Load a folder first.");
    return;
  }
  if (state.autoCaptioning || state.cloning || state.moving || state.extractingFrame || state.uploading || state.duplicatingImage) {
    showErrorToast("Finish the current operation before copying media.");
    return;
  }

  const selectedPaths = getMoveSelectionPaths();
  if (!selectedPaths.length) {
    showErrorToast("Select one or more media files first.");
    return;
  }

  renderMoveDialogSummary(selectedPaths);
  setMoveDialogStatus("");
  moveTargetFolderInput.value = getDefaultMoveTargetFolderValue();
  clearMoveFolderAutocomplete({ cancelPending: true });
  moveSelectedModal.classList.add("open");
  moveSelectedModal.setAttribute("aria-hidden", "false");
  window.requestAnimationFrame(() => {
    moveTargetFolderInput.focus();
    moveTargetFolderInput.select();
    if (moveTargetFolderInput.value.trim()) {
      scheduleMoveFolderAutocompleteRefresh({ immediate: true });
    }
  });
}

function closeMoveSelectedDialog() {
  moveSelectedModal.classList.remove("open");
  moveSelectedModal.setAttribute("aria-hidden", "true");
  setMoveDialogStatus("");
  clearMoveFolderAutocomplete({ cancelPending: true });
}

async function submitMoveSelectedDialog(event) {
  if (event) event.preventDefault();
  await moveSelectedMediaToFolder(moveTargetFolderInput.value);
}

async function moveSelectedMediaToFolder(rawTargetFolder) {
  if (!state.folder) {
    showErrorToast("Load a folder first.");
    return;
  }
  if (state.autoCaptioning || state.cloning || state.moving || state.extractingFrame || state.uploading || state.duplicatingImage) {
    showErrorToast("Finish the current operation before copying media.");
    return;
  }

  const selectedPaths = getMoveSelectionPaths();
  if (!selectedPaths.length) {
    setMoveDialogStatus("Select one or more media files first.");
    return;
  }

  const targetFolder = String(rawTargetFolder || "").trim();
  if (!targetFolder) {
    setMoveDialogStatus("Select an existing target folder.");
    moveTargetFolderInput.focus();
    return;
  }
  if (normalizeFolderPathForCompare(targetFolder) === normalizeFolderPathForCompare(state.folder)) {
    setMoveDialogStatus("Target folder must be different from the current folder.");
    moveTargetFolderInput.focus();
    moveTargetFolderInput.select();
    return;
  }

  closeMoveSelectedDialog();
  state.moving = true;
  updateActionButtons();
  resetAutoCaptionProgress();
  updateAutoCaptionProgress({
    visible: true,
    scopeLabel: "Copy Selection",
    totalImages: Math.max(1, selectedPaths.length),
    processedImages: 0,
    completedImages: 0,
    errors: 0,
    currentPath: targetFolder,
    currentMessage: "Preparing copy...",
    currentStepIndex: 0,
    currentStepTotal: 1,
  });
  statusBar.textContent = selectedPaths.length === 1
    ? `Copying ${getFileLabel(selectedPaths[0])}...`
    : `Copying ${selectedPaths.length} media files...`;

  const preserveScrollTop = fileGridContainer.scrollTop;
  try {
    const resp = await fetch("/api/media/copy/stream", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        source_folder: state.folder,
        target_folder: targetFolder,
        image_paths: selectedPaths,
      }),
    });

    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      throw new Error(data.detail || "Failed to copy selected media");
    }

    let copiedCount = 0;
    await readNdjsonStream(resp, (event) => {
      if (!event || typeof event !== "object") return;
      if (event.type === "start") {
        updateAutoCaptionProgress({
          visible: true,
          scopeLabel: "Copy Selection",
          totalImages: Math.max(1, Number(event.total || selectedPaths.length || 1)),
          processedImages: 0,
          completedImages: 0,
          currentPath: event.target_folder || targetFolder,
          currentMessage: "Preparing copy...",
          currentStepIndex: 0,
          currentStepTotal: 1,
        });
        return;
      }
      if (event.type === "progress") {
        copiedCount = Number(event.index || copiedCount || 0);
        updateAutoCaptionProgress({
          visible: true,
          totalImages: Math.max(1, Number(event.total || selectedPaths.length || 1)),
          processedImages: copiedCount,
          completedImages: copiedCount,
          currentPath: event.target_path || event.source_path || "",
          currentMessage: "Moved",
          currentStepIndex: 1,
          currentStepTotal: 1,
        });
        return;
      }
      if (event.type === "config-updated") {
        updateAutoCaptionProgress({
          currentPath: event.target_folder || targetFolder,
          currentMessage: event.sections_changed ? "Merged caption library" : "Updated folder config",
          currentStepIndex: 1,
          currentStepTotal: 1,
        });
        return;
      }
      if (event.type === "done") {
        copiedCount = Number(event.copied || event.moved || copiedCount || selectedPaths.length || 0);
        updateAutoCaptionProgress({
          visible: true,
          totalImages: Math.max(1, Number(event.total || selectedPaths.length || 1)),
          processedImages: Math.max(1, copiedCount),
          completedImages: Math.max(1, copiedCount),
          currentPath: event.target_folder || targetFolder,
          currentMessage: "Copy complete",
          currentStepIndex: 1,
          currentStepTotal: 1,
        });
        return;
      }
      if (event.type === "error") {
        throw new Error(event.message || "Copy failed");
      }
    });

    state.moveDialog.lastTargetFolder = targetFolder;
    await loadFolder({ preserveScrollTop });
    statusBar.textContent = copiedCount === 1
      ? `Copied 1 media file to ${getFileLabel(targetFolder)}`
      : `Copied ${copiedCount || selectedPaths.length} media files to ${getFileLabel(targetFolder)}`;
  } catch (err) {
    const message = err?.message || "Failed to copy selected media";
    statusBar.textContent = `Copy error: ${message}`;
    showErrorToast(`Copy error: ${message}`);
  } finally {
    state.moving = false;
    updateActionButtons();
    resetAutoCaptionProgress();
  }
}

async function cloneCurrentFolder() {
  if (!state.folder) {
    showErrorToast("Load a folder first.");
    return;
  }
  if (state.cloning || state.moving || state.extractingFrame || state.autoCaptioning || state.uploading) {
    showErrorToast("Finish the current operation before cloning.");
    return;
  }

  const selectedPaths = getCloneSelectionPaths();
  const selectionLabel = selectedPaths.length > 1
    ? `${selectedPaths.length} selected media files`
    : "the whole folder";
  const newFolderName = prompt(`Clone ${selectionLabel} into a new sibling folder named:`, getDefaultCloneFolderName());
  if (newFolderName === null) return;
  const trimmedName = String(newFolderName || "").trim();
  if (!trimmedName) {
    showErrorToast("Clone cancelled: folder name is required.");
    return;
  }

  state.cloning = true;
  updateActionButtons();
  state.modelLogLines = [];
  state.modelLogOpen = false;
  resetAutoCaptionProgress();
  updateAutoCaptionProgress({
    visible: true,
    scopeLabel: selectedPaths.length > 1 ? "Clone Selection" : "Clone Folder",
    totalImages: 1,
    processedImages: 0,
    completedImages: 0,
    errors: 0,
    currentMessage: "Preparing clone...",
    currentStepIndex: 0,
    currentStepTotal: 1,
  });
  statusBar.textContent = `Cloning to ${trimmedName}...`;

  try {
    const resp = await fetch("/api/folder/clone/stream", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        source_folder: state.folder,
        new_folder_name: trimmedName,
        image_paths: selectedPaths,
      }),
    });

    if (!resp.ok) {
      const data = await resp.json().catch(() => ({}));
      throw new Error(data.detail || "Failed to clone folder");
    }

    let targetFolder = "";
    await readNdjsonStream(resp, (event) => {
      if (!event || typeof event !== "object") return;
      if (event.type === "start") {
        targetFolder = event.target_folder || targetFolder;
        updateAutoCaptionProgress({
          visible: true,
          scopeLabel: event.mode === "selected" ? "Clone Selection" : "Clone Folder",
          totalImages: Math.max(1, Number(event.total || 0)),
          processedImages: 0,
          completedImages: 0,
          errors: 0,
          currentPath: event.target_folder || "",
          currentMessage: "Creating folder...",
          currentStepIndex: 0,
          currentStepTotal: 1,
        });
        return;
      }
      if (event.type === "progress") {
        updateAutoCaptionProgress({
          visible: true,
          totalImages: Math.max(1, Number(event.total || state.aiProgress.totalImages || 1)),
          processedImages: Number(event.index || 0),
          completedImages: Number(event.index || 0),
          currentPath: event.path || "",
          currentMessage: "Copied",
          currentStepIndex: 1,
          currentStepTotal: 1,
        });
        return;
      }
      if (event.type === "config-copied") {
        targetFolder = event.target_folder || targetFolder;
        updateAutoCaptionProgress({
          currentPath: event.target_folder || "",
          currentMessage: "Copied config",
          currentStepIndex: 1,
          currentStepTotal: 1,
        });
        return;
      }
      if (event.type === "done") {
        targetFolder = event.target_folder || targetFolder;
        updateAutoCaptionProgress({
          visible: true,
          totalImages: Math.max(1, Number(event.total || state.aiProgress.totalImages || 1)),
          processedImages: Number(event.total || state.aiProgress.totalImages || 1),
          completedImages: Number(event.total || state.aiProgress.totalImages || 1),
          currentPath: event.target_folder || "",
          currentMessage: "Clone complete",
          currentStepIndex: 1,
          currentStepTotal: 1,
        });
        return;
      }
      if (event.type === "error") {
        throw new Error(event.message || "Clone failed");
      }
    });

    if (!targetFolder) {
      throw new Error("Clone finished without a target folder");
    }

    folderInput.value = targetFolder;
    await loadFolder();
    statusBar.textContent = `Cloned into ${getFileLabel(targetFolder)}`;
  } catch (err) {
    const message = err?.message || "Failed to clone folder";
    statusBar.textContent = `Clone error: ${message}`;
    showErrorToast(`Clone error: ${message}`);
  } finally {
    state.cloning = false;
    updateActionButtons();
    resetAutoCaptionProgress();
  }
}

async function duplicateCurrentImage() {
  if (!isImageEditAvailable()) {
    showErrorToast("Select a single image first.");
    return;
  }
  if (state.duplicatingImage || state.autoCaptioning || state.cloning || state.moving || state.extractingFrame || state.uploading) {
    showErrorToast("Finish the current operation before duplicating an image.");
    return;
  }

  const sourcePath = state.previewPath;
  const requestedName = prompt(`Duplicate ${getFileLabel(sourcePath)} as:`, getDefaultDuplicateImageName(sourcePath));
  if (requestedName === null) return;
  const trimmedName = String(requestedName || "").trim();
  if (!trimmedName) {
    showErrorToast("Duplicate cancelled: image name is required.");
    return;
  }

  state.duplicatingImage = true;
  renderMaskEditorUi();
  statusBar.textContent = `Duplicating ${getFileLabel(sourcePath)}...`;
  try {
    const resp = await fetch("/api/image/duplicate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        image_path: sourcePath,
        new_name: trimmedName,
      }),
    });
    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      throw new Error(data.detail || "Failed to duplicate image");
    }
    const preserveScrollTop = fileGridContainer.scrollTop;
    await loadFolder({ preserveScrollTop });
    await selectUploadedImages([data.image_path]);
    statusBar.textContent = `Duplicated ${getFileLabel(sourcePath)} as ${getFileLabel(data.image_path)}`;
  } catch (err) {
    showErrorToast(`Duplicate error: ${err.message}`);
    statusBar.textContent = `Duplicate error: ${err.message}`;
  } finally {
    state.duplicatingImage = false;
    renderMaskEditorUi();
  }
}
