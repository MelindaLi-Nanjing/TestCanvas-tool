(function () {
  const storageKey = "test-record-tool-state-v1";
  const mediaDbName = "test-record-tool-media-v1";
  const mediaStoreName = "videos";
  const snapshotStoreName = "snapshots";
  const recognitionCtor = window.SpeechRecognition || window.webkitSpeechRecognition;

  const defaultState = {
    runName: "",
    ownerName: "",
    queryCount: 0,
    versionDetails: "",
    presentationMode: false,
    reviewMode: false,
    autoCapture: false,
    lastSavedAt: "",
    tests: [],
      templates: [],
      setup: {
        textHtml: "",
        screenshots: [],
        links: ""
      }
  };

  let state = loadState();
  normalizeTestStatusesInState();
  let pendingScreenshotTarget = null;
  let selectedDropTarget = null;
  let pendingFileTarget = null;
  let activeVoiceContext = null;
  let drawingSession = null;
  let drawingClipboard = null;
  let recordingSession = null;
  let isClosingDrawingAfterSave = false;
  let suppressDialogReopenUntil = 0;
  const mediaObjectUrls = new Map();
  const mediaDbPromise = openMediaDatabase();

  const toolShortcutMap = {
    KeyP: "pen",
    KeyA: "arrow",
    KeyR: "rectangle",
    KeyH: "highlight",
    KeyT: "text",
    KeyC: "callout"
  };

  const elements = {
    newRunButton: document.getElementById("newRunButton"),
    exportButton: document.getElementById("exportButton"),
    exportSummaryExcelButton: document.getElementById("exportSummaryExcelButton"),
    addTestButton: document.getElementById("addTestButton"),
    bottomAddTestButton: document.getElementById("bottomAddTestButton"),
    updateSummaryTableButton: document.getElementById("updateSummaryTableButton"),
    saveRunTopButton: document.getElementById("saveRunTopButton"),
    importJsonButton: document.getElementById("importJsonButton"),
    saveRunButton: document.getElementById("saveRunButton"),
    restoreSnapshotButton: document.getElementById("restoreSnapshotButton"),
    exportJsonButton: document.getElementById("exportJsonButton"),
    importExcelButton: document.getElementById("importExcelButton"),
    downloadExcelTemplateButton: document.getElementById("downloadExcelTemplateButton"),
    exportMediaButton: document.getElementById("exportMediaButton"),
    reviewModeButton: document.getElementById("reviewModeButton"),
    addScreenshotButton: document.getElementById("addScreenshotButton"),
    pasteScreenshotButton: document.getElementById("pasteScreenshotButton"),
    startRecordingButton: document.getElementById("startRecordingButton"),
    pauseRecordingButton: document.getElementById("pauseRecordingButton"),
    stopRecordingButton: document.getElementById("stopRecordingButton"),
    recordingStatus: document.getElementById("recordingStatus"),
    floatingRecordingControls: document.getElementById("floatingRecordingControls"),
    floatingRecordingLabel: document.getElementById("floatingRecordingLabel"),
    floatingPauseButton: document.getElementById("floatingPauseButton"),
    floatingStopButton: document.getElementById("floatingStopButton"),
    screenshotInput: document.getElementById("screenshotInput"),
    jsonInput: document.getElementById("jsonInput"),
    excelInput: document.getElementById("excelInput"),
    testsList: document.getElementById("testsList"),
    reviewOverview: document.getElementById("reviewOverview"),
    floatingSummaryNav: document.getElementById("floatingSummaryNav"),
    floatingSummaryNavBody: document.getElementById("floatingSummaryNavBody"),
    updateFloatingSummaryButton: document.getElementById("updateFloatingSummaryButton"),
    summaryCards: document.getElementById("summaryCards"),
    saveStatus: document.getElementById("saveStatus"),
    saveWarningBanner: document.getElementById("saveWarningBanner"),
    runName: document.getElementById("runName"),
    ownerName: document.getElementById("ownerName"),
    queryCount: document.getElementById("queryCount"),
    versionDetails: document.getElementById("versionDetails"),
    presentationModeToggle: document.getElementById("presentationModeToggle"),
    autoCaptureToggle: document.getElementById("autoCaptureToggle"),
    testCardTemplate: document.getElementById("testCardTemplate"),
    stepTemplate: document.getElementById("stepTemplate"),
    videoTemplate: document.getElementById("videoTemplate"),
    shotTemplate: document.getElementById("shotTemplate"),
    reviewDialog: document.getElementById("reviewDialog"),
    reviewSummary: document.getElementById("reviewSummary"),
    mediaPreviewDialog: document.getElementById("mediaPreviewDialog"),
    mediaPreviewTitle: document.getElementById("mediaPreviewTitle"),
    mediaPreviewImage: document.getElementById("mediaPreviewImage"),
    mediaPreviewVideo: document.getElementById("mediaPreviewVideo"),
    closeMediaPreviewButton: document.getElementById("closeMediaPreviewButton"),
    setupText: document.getElementById("setupText"),
    setupScreenshots: document.getElementById("setupScreenshots"),
    addSetupScreenshotButton: document.getElementById("addSetupScreenshotButton"),
    setupLinks: document.getElementById("setupLinks"),
    drawingDialog: document.getElementById("drawingDialog"),
    closeDrawingDialogButton: document.getElementById("closeDrawingDialogButton"),
    drawingTitle: document.getElementById("drawingTitle"),
    drawingImage: document.getElementById("drawingImage"),
    drawingCanvas: document.getElementById("drawingCanvas"),
    markupTool: document.getElementById("markupTool"),
    penColor: document.getElementById("penColor"),
    penSize: document.getElementById("penSize"),
    textLabel: document.getElementById("textLabel"),
    textInput: document.getElementById("textInput"),
    selectedTextLabel: document.getElementById("selectedTextLabel"),
    selectedTextInput: document.getElementById("selectedTextInput"),
    toggleHelpButton: document.getElementById("toggleHelpButton"),
    drawingHelpPanel: document.getElementById("drawingHelpPanel"),
    duplicateItemButton: document.getElementById("duplicateItemButton"),
    copyItemButton: document.getElementById("copyItemButton"),
    pasteItemButton: document.getElementById("pasteItemButton"),
    undoStrokeButton: document.getElementById("undoStrokeButton"),
    clearDrawingButton: document.getElementById("clearDrawingButton"),
    saveDrawingButton: document.getElementById("saveDrawingButton"),
    mainBackToTopButton: document.getElementById("mainBackToTopButton"),
    fileInput: document.getElementById("fileInput")
  };

  bindTopLevelEvents();
  render();
  void initializeMediaStorage();

  function loadState() {
    try {
      const raw = localStorage.getItem(storageKey);
      if (!raw) {
        return structuredClone(defaultState);
      }
      return { ...structuredClone(defaultState), ...JSON.parse(raw) };
    } catch {
      return structuredClone(defaultState);
    }
  }

  function normalizeTestStatus(status) {
    if (status === "not-started") return "not-start";
    if (status === "pass") return "passed";
    if (status === "fail") return "failed";
    if (status === "not-start" || status === "passed" || status === "query" || status === "failed" || status === "blocked" || status === "cancelled") {
      return status;
    }
    return "not-start";
  }

  function normalizeTestStatusesInState() {
    state.tests = (state.tests || []).map((test) => ({
      ...test,
      status: normalizeTestStatus(test.status)
    }));
  }

  function openMediaDatabase() {
    return new Promise((resolve, reject) => {
      if (!window.indexedDB) {
        reject(new Error("IndexedDB is not available."));
        return;
      }
      const request = window.indexedDB.open(mediaDbName, 2);
      request.onupgradeneeded = () => {
        const database = request.result;
        if (!database.objectStoreNames.contains(mediaStoreName)) {
          database.createObjectStore(mediaStoreName, { keyPath: "id" });
        }
        if (!database.objectStoreNames.contains(snapshotStoreName)) {
          database.createObjectStore(snapshotStoreName, { keyPath: "id" });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Failed to open media database."));
    });
  }

  async function withMediaStore(mode, handler) {
    const database = await mediaDbPromise;
    return new Promise((resolve, reject) => {
      const transaction = database.transaction(mediaStoreName, mode);
      const store = transaction.objectStore(mediaStoreName);
      let result;
      try {
        result = handler(store, transaction);
      } catch (error) {
        reject(error);
        return;
      }
      transaction.oncomplete = () => resolve(result);
      transaction.onerror = () => reject(transaction.error || new Error("IndexedDB transaction failed."));
      transaction.onabort = () => reject(transaction.error || new Error("IndexedDB transaction aborted."));
    });
  }

  async function saveMediaBlob(mediaId, blob) {
    await withMediaStore("readwrite", (store) => {
      store.put({ id: mediaId, blob, updatedAt: Date.now() });
    });
  }

  async function getMediaBlob(mediaId) {
    const database = await mediaDbPromise;
    return new Promise((resolve, reject) => {
      const transaction = database.transaction(mediaStoreName, "readonly");
      const store = transaction.objectStore(mediaStoreName);
      const request = store.get(mediaId);
      request.onsuccess = () => resolve(request.result?.blob || null);
      request.onerror = () => reject(request.error || new Error("Failed to read media blob."));
    });
  }

  function revokeMediaPlaybackUrl(mediaId) {
    const objectUrl = mediaObjectUrls.get(mediaId);
    if (!objectUrl) {
      return;
    }
    URL.revokeObjectURL(objectUrl);
    mediaObjectUrls.delete(mediaId);
  }

  async function deleteMediaBlob(mediaId) {
    revokeMediaPlaybackUrl(mediaId);
    await withMediaStore("readwrite", (store) => {
      store.delete(mediaId);
    });
  }

  async function saveVideoBlob(videoId, blob) {
    await saveMediaBlob(videoId, blob);
  }

  async function getVideoBlob(videoId) {
    return getMediaBlob(videoId);
  }

  async function deleteVideoBlob(videoId) {
    await deleteMediaBlob(videoId);
  }

  async function initializeMediaStorage() {
    try {
      await mediaDbPromise;
      await syncStateFromLatestSnapshotIfNewer();
      await migrateInlineVideosToIndexedDb();
      await migrateInlineScreenshotsAndFilesToIndexedDb();
      await updateSnapshotButtonState();
    } catch {
      if (elements.saveStatus) elements.saveStatus.textContent = "IndexedDB unavailable - video storage may be limited";
    }
  }

  async function syncStateFromLatestSnapshotIfNewer() {
    const latestSnapshot = await getLatestSnapshot();
    if (!latestSnapshot?.stateJson) {
      return;
    }
    let snapshotState;
    try {
      snapshotState = JSON.parse(latestSnapshot.stateJson);
    } catch {
      return;
    }

    const localTs = Date.parse(state?.lastSavedAt || "");
    const snapshotTs = Date.parse(snapshotState?.lastSavedAt || latestSnapshot.createdAt || "");
    const useSnapshot = Number.isFinite(snapshotTs) && (!Number.isFinite(localTs) || snapshotTs > localTs);

    if (!useSnapshot) {
      return;
    }

    state = {
      ...structuredClone(defaultState),
      ...snapshotState,
      tests: Array.isArray(snapshotState.tests) ? snapshotState.tests : []
    };
    normalizeTestStatusesInState();
    render();
    saveState();
    if (elements.saveStatus) {
      elements.saveStatus.textContent = "Loaded latest autosaved snapshot";
    }
  }

  function hasMeaningfulState(snapshotState = state) {
    return Boolean(
      snapshotState?.runName
      || snapshotState?.ownerName
      || snapshotState?.versionDetails
      || snapshotState?.queryCount
      || (snapshotState?.tests || []).length
    );
  }

  async function getAllSnapshots() {
    const database = await mediaDbPromise;
    return new Promise((resolve, reject) => {
      const transaction = database.transaction(snapshotStoreName, "readonly");
      const store = transaction.objectStore(snapshotStoreName);
      const request = store.getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error || new Error("Failed to read run snapshots."));
    });
  }

  async function getLatestSnapshot() {
    const snapshots = await getAllSnapshots();
    return snapshots.sort((left, right) => String(right.createdAt).localeCompare(String(left.createdAt)))[0] || null;
  }

  async function pruneSnapshots(maxSnapshots = 5) {
    const snapshots = await getAllSnapshots();
    const staleSnapshots = snapshots
      .sort((left, right) => String(right.createdAt).localeCompare(String(left.createdAt)))
      .slice(maxSnapshots);
    if (!staleSnapshots.length) {
      return;
    }
    const database = await mediaDbPromise;
    await new Promise((resolve, reject) => {
      const transaction = database.transaction(snapshotStoreName, "readwrite");
      const store = transaction.objectStore(snapshotStoreName);
      staleSnapshots.forEach((snapshot) => store.delete(snapshot.id));
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error || new Error("Failed to trim old run snapshots."));
      transaction.onabort = () => reject(transaction.error || new Error("Failed to trim old run snapshots."));
    });
  }

  async function saveRunSnapshot(reason = "manual") {
    if (!hasMeaningfulState()) {
      return null;
    }
    const snapshot = {
      id: crypto.randomUUID(),
      createdAt: new Date().toISOString(),
      reason,
      runName: state.runName || "Untitled run",
      ownerName: state.ownerName || "",
      testCount: (state.tests || []).length,
      stateJson: JSON.stringify(state)
    };
    const database = await mediaDbPromise;
    await new Promise((resolve, reject) => {
      const transaction = database.transaction(snapshotStoreName, "readwrite");
      const store = transaction.objectStore(snapshotStoreName);
      store.put(snapshot);
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error || new Error("Failed to save a run snapshot."));
      transaction.onabort = () => reject(transaction.error || new Error("Failed to save a run snapshot."));
    });
    await pruneSnapshots();
    await updateSnapshotButtonState();
    return snapshot;
  }

  async function updateSnapshotButtonState() {
    if (!elements.restoreSnapshotButton) {
      return;
    }
    try {
      const latestSnapshot = await getLatestSnapshot();
      elements.restoreSnapshotButton.disabled = !latestSnapshot;
      elements.restoreSnapshotButton.title = latestSnapshot
        ? `Restore ${latestSnapshot.runName || "run"} saved ${new Date(latestSnapshot.createdAt).toLocaleString()}`
        : "No run snapshot available yet";
    } catch {
      elements.restoreSnapshotButton.disabled = true;
      elements.restoreSnapshotButton.title = "Run snapshots are unavailable in this browser";
    }
  }

  async function restoreLatestSnapshot() {
    const latestSnapshot = await getLatestSnapshot();
    if (!latestSnapshot) {
      window.alert("No saved run snapshot is available yet.");
      return false;
    }
    const parsed = JSON.parse(latestSnapshot.stateJson);
    state = {
      ...structuredClone(defaultState),
      ...parsed,
      tests: Array.isArray(parsed.tests) ? parsed.tests : []
    };
    normalizeTestStatusesInState();
    render();
    await migrateInlineVideosToIndexedDb();
    await migrateInlineScreenshotsAndFilesToIndexedDb();
    saveState();
    if (elements.saveStatus) elements.saveStatus.textContent = `Restored snapshot from ${new Date(latestSnapshot.createdAt).toLocaleString()}`;
    return true;
  }

  async function migrateInlineVideosToIndexedDb() {
    const videosNeedingMigration = collectAllVideos().filter((video) => video.dataUrl?.startsWith("data:"));
    if (!videosNeedingMigration.length) {
      return;
    }
    await Promise.all(videosNeedingMigration.map(async (video) => {
      const blob = dataUrlToBlob(video.dataUrl);
      await saveVideoBlob(video.id, blob);
      video.mimeType = video.mimeType || blob.type || video.mimeType;
      video.sizeBytes = video.sizeBytes || blob.size;
      video.storage = "indexeddb";
      delete video.dataUrl;
    }));
    saveState();
    render();
  }

  async function migrateInlineScreenshotsAndFilesToIndexedDb() {
    const shotsNeedingMigration = collectAllScreenshots().filter((shot) => shot.dataUrl?.startsWith("data:"));
    const filesNeedingMigration = collectAllFileAttachments().filter((file) => file.dataUrl?.startsWith("data:"));
    if (!shotsNeedingMigration.length && !filesNeedingMigration.length) {
      return;
    }

    await Promise.all([
      ...shotsNeedingMigration.map(async (shot) => {
        const blob = dataUrlToBlob(shot.dataUrl);
        await saveMediaBlob(shot.id, blob);
        shot.mimeType = shot.mimeType || blob.type || "image/png";
        shot.sizeBytes = shot.sizeBytes || blob.size;
        shot.storage = "indexeddb";
        delete shot.dataUrl;
      }),
      ...filesNeedingMigration.map(async (file) => {
        const blob = dataUrlToBlob(file.dataUrl);
        await saveMediaBlob(file.id, blob);
        file.type = file.type || blob.type || "application/octet-stream";
        file.sizeBytes = file.sizeBytes || blob.size;
        file.storage = "indexeddb";
        delete file.dataUrl;
      })
    ]);

    saveState();
    render();
  }

  function collectAllVideos() {
    return (state.tests || []).flatMap((test) => [
      ...(test.videos || []),
      ...((test.steps || []).flatMap((step) => step.videos || []))
    ]);
  }

  function collectAllScreenshots() {
    return (state.tests || []).flatMap((test) => [
      ...(test.screenshots || []),
      ...((test.steps || []).flatMap((step) => step.screenshots || []))
    ]);
  }

  function collectAllFileAttachments() {
    return (state.tests || []).flatMap((test) => [
      ...(test.files || []),
      ...((test.steps || []).flatMap((step) => step.files || []))
    ]);
  }

  async function deleteVideos(videos) {
    await Promise.all((videos || []).map((video) => deleteVideoBlob(video.id)));
  }

  async function deleteVideosForTests(tests) {
    const videos = (tests || []).flatMap((test) => [
      ...(test.videos || []),
      ...((test.steps || []).flatMap((step) => step.videos || []))
    ]);
    await deleteVideos(videos);
  }

  async function deleteBinaryMediaForTests(tests) {
    const mediaIds = (tests || []).flatMap((test) => [
      ...(test.screenshots || []).map((shot) => shot.id),
      ...(test.files || []).map((file) => file.id),
      ...((test.steps || []).flatMap((step) => [
        ...(step.screenshots || []).map((shot) => shot.id),
        ...(step.files || []).map((file) => file.id)
      ]))
    ]).filter(Boolean);
    await Promise.all(mediaIds.map((mediaId) => deleteMediaBlob(mediaId)));
  }

  async function exportMedia() {
    const videos = collectAllVideos();
    if (!videos.length) {
      window.alert("No recorded videos are available to export.");
      return;
    }

    if (window.showDirectoryPicker) {
      try {
        const directoryHandle = await window.showDirectoryPicker({ mode: "readwrite" });
        for (const video of videos) {
          const blob = await getVideoBlobForVideo(video);
          if (!blob) {
            continue;
          }
          const fileHandle = await directoryHandle.getFileHandle(video.name || `${video.id}.${extensionFromMimeType(video.mimeType)}`, { create: true });
          const writable = await fileHandle.createWritable();
          await writable.write(blob);
          await writable.close();
        }
        window.alert(`Exported ${videos.length} media file(s) to the selected folder.`);
        return;
      } catch (error) {
        if (error?.name !== "AbortError") {
          window.alert("Folder export was interrupted. Falling back to browser downloads.");
        } else {
          return;
        }
      }
    }

    for (const video of videos) {
      const blob = await getVideoBlobForVideo(video);
      if (!blob) {
        continue;
      }
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = video.name || `${video.id}.${extensionFromMimeType(video.mimeType)}`;
      anchor.click();
      URL.revokeObjectURL(url);
    }
  }

  function saveState() {
    state.lastSavedAt = new Date().toISOString();
    try {
      localStorage.setItem(storageKey, JSON.stringify(state));
      if (elements.saveStatus) elements.saveStatus.textContent = "Saved locally";
      if (elements.saveWarningBanner) {
        elements.saveWarningBanner.hidden = true;
        elements.saveWarningBanner.textContent = "";
      }
    } catch {
      if (elements.saveStatus) elements.saveStatus.textContent = "Storage full - export JSON to keep this run";
      if (elements.saveWarningBanner) {
        elements.saveWarningBanner.textContent = "Warning: Browser local storage is full. Your latest edits might not persist after reload. Use Export Run JSON now, then clear old browser data if needed.";
        elements.saveWarningBanner.hidden = false;
      }
      void saveRunSnapshot("autosave-fallback").catch(() => {});
    }
    window.clearTimeout(saveState.timeoutId);
    saveState.timeoutId = window.setTimeout(() => {
      if (elements.saveStatus) elements.saveStatus.textContent = "Ready";
    }, 1200);
  }

  async function saveRunManually() {
    saveState();
    try {
      const snapshot = await saveRunSnapshot("manual-save");
      if (snapshot && elements.saveStatus) {
        elements.saveStatus.textContent = `Saved checkpoint ${new Date(snapshot.createdAt).toLocaleTimeString()}`;
      }
    } catch {
      if (elements.saveStatus) elements.saveStatus.textContent = "Saved locally (snapshot unavailable)";
    }
    window.clearTimeout(saveState.timeoutId);
    saveState.timeoutId = window.setTimeout(() => {
      if (elements.saveStatus) elements.saveStatus.textContent = "Ready";
    }, 1600);
  }

  function bindTopLevelEvents() {
      // SetUp section: text
      if (elements.setupText) {
        elements.setupText.innerHTML = state.setup?.textHtml || "";
        elements.setupText.addEventListener("input", () => {
          state.setup.textHtml = elements.setupText.innerHTML;
          saveState();
        });
      }

      // SetUp section: links
      if (elements.setupLinks) {
        elements.setupLinks.value = state.setup?.links || "";
        elements.setupLinks.addEventListener("input", () => {
          state.setup.links = elements.setupLinks.value;
          saveState();
        });
      }

      // SetUp section: screenshots
      if (elements.addSetupScreenshotButton && elements.screenshotInput && elements.setupScreenshots) {
        elements.addSetupScreenshotButton.addEventListener("click", () => {
          elements.screenshotInput.dataset.setup = "true";
          elements.screenshotInput.click();
        });
        // Render existing screenshots
        renderSetupScreenshots();
      }

      // Screenshot input handler (extend for SetUp)
      if (elements.screenshotInput) {
        elements.screenshotInput.addEventListener("change", async (event) => {
          const files = Array.from(event.target.files || []);
          if (elements.screenshotInput.dataset.setup === "true") {
            for (const file of files) {
              const reader = new FileReader();
              reader.onload = (e) => {
                state.setup.screenshots.push({
                  id: crypto.randomUUID(),
                  name: file.name,
                  dataUrl: e.target.result,
                  createdAt: new Date().toISOString()
                });
                renderSetupScreenshots();
                saveState();
              };
              reader.readAsDataURL(file);
            }
            elements.screenshotInput.value = "";
            delete elements.screenshotInput.dataset.setup;
            return;
          }
          // ...existing code for test screenshots...
        }, true);
      }

      function renderSetupScreenshots() {
        if (!elements.setupScreenshots) return;
        elements.setupScreenshots.innerHTML = "";
        (state.setup?.screenshots || []).forEach((shot, idx) => {
          const wrapper = document.createElement("div");
          wrapper.style.position = "relative";
          wrapper.style.display = "inline-block";
          wrapper.style.border = "1px solid #d3c7af";
          wrapper.style.borderRadius = "8px";
          wrapper.style.overflow = "hidden";
          wrapper.style.background = "#fff";
          wrapper.style.maxWidth = "80px";
          wrapper.style.maxHeight = "80px";
          wrapper.style.margin = "2px";
          const img = document.createElement("img");
          img.src = shot.dataUrl;
          img.alt = shot.name || `Screenshot ${idx + 1}`;
          img.style.width = "80px";
          img.style.height = "80px";
          img.style.objectFit = "cover";
          img.title = shot.name || "Screenshot";
          wrapper.appendChild(img);
          const delBtn = document.createElement("button");
          delBtn.textContent = "×";
          delBtn.type = "button";
          delBtn.style.position = "absolute";
          delBtn.style.top = "2px";
          delBtn.style.right = "2px";
          delBtn.style.background = "#fff";
          delBtn.style.border = "none";
          delBtn.style.borderRadius = "50%";
          delBtn.style.width = "20px";
          delBtn.style.height = "20px";
          delBtn.style.cursor = "pointer";
          delBtn.style.fontWeight = "bold";
          delBtn.addEventListener("click", () => {
            state.setup.screenshots.splice(idx, 1);
            renderSetupScreenshots();
            saveState();
          });
          wrapper.appendChild(delBtn);
          elements.setupScreenshots.appendChild(wrapper);
        });
      }
    elements.newRunButton.addEventListener("click", async () => {
      if (!window.confirm("Start a new test run? This clears the current local record.")) {
        return;
      }
      try {
        await saveRunSnapshot("before-new-run");
      } catch {
        window.alert("Unable to save a recovery snapshot before clearing this run.");
        return;
      }
      await deleteVideosForTests(state.tests || []);
      await deleteBinaryMediaForTests(state.tests || []);
      state = structuredClone(defaultState);
      render();
      saveState();
    });

    elements.addTestButton.addEventListener("click", () => {
      state.tests.push(createTest());
      render();
      saveState();
    });

    elements.bottomAddTestButton.addEventListener("click", () => {
      state.tests.push(createTest());
      render();
      saveState();
      const lastTestCard = document.querySelector(".test-card:last-of-type");
      if (lastTestCard) {
        lastTestCard.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    });

    if (elements.updateSummaryTableButton) {
      elements.updateSummaryTableButton.addEventListener("click", (event) => {
        event.preventDefault();
        renderReviewOverview();
        if (elements.saveStatus) {
          elements.saveStatus.textContent = "Summary updated";
          window.clearTimeout(saveState.timeoutId);
          saveState.timeoutId = window.setTimeout(() => {
            if (elements.saveStatus) elements.saveStatus.textContent = "Ready";
          }, 1200);
        }
      });
    }

    if (elements.updateFloatingSummaryButton) {
      elements.updateFloatingSummaryButton.addEventListener("click", () => {
        renderReviewOverview();
      });
    }

    elements.exportButton.addEventListener("click", exportRun);
    if (elements.exportSummaryExcelButton) {
      elements.exportSummaryExcelButton.addEventListener("click", () => {
        void exportSummarySpreadsheet();
      });
    }
    elements.saveRunButton.addEventListener("click", () => {
      void saveRunManually();
    });
    if (elements.saveRunTopButton) {
      elements.saveRunTopButton.addEventListener("click", () => {
        void saveRunManually();
      });
    }
    elements.exportJsonButton.addEventListener("click", exportJson);
    elements.exportMediaButton.addEventListener("click", exportMedia);
    if (elements.importJsonButton) elements.importJsonButton.addEventListener("click", () => elements.jsonInput.click());
    if (elements.importExcelButton && elements.excelInput) {
      elements.importExcelButton.addEventListener("click", () => elements.excelInput.click());
    }
    if (elements.downloadExcelTemplateButton) {
      elements.downloadExcelTemplateButton.addEventListener("click", () => {
        void downloadSpreadsheetTemplate();
      });
    }
    elements.restoreSnapshotButton.addEventListener("click", async () => {
      await restoreLatestSnapshot();
    });
    if (elements.reviewModeButton) elements.reviewModeButton.addEventListener("click", toggleReviewMode);
    if (elements.startRecordingButton) elements.startRecordingButton.addEventListener("click", () => startScreenRecording());
    if (elements.pauseRecordingButton) elements.pauseRecordingButton.addEventListener("click", togglePauseRecording);
    if (elements.stopRecordingButton) elements.stopRecordingButton.addEventListener("click", stopScreenRecording);
    if (elements.floatingPauseButton) {
      elements.floatingPauseButton.addEventListener("click", togglePauseRecording);
    }
    if (elements.floatingStopButton) {
      elements.floatingStopButton.addEventListener("click", stopScreenRecording);
    }

    elements.addScreenshotButton.addEventListener("click", () => {
      const latestTestId = state.tests.at(-1)?.id || null;
      pendingScreenshotTarget = latestTestId ? { testId: latestTestId, stepId: null } : null;
      if (!pendingScreenshotTarget?.testId) {
        window.alert("Create a test first, then attach screenshots to it.");
        return;
      }
      elements.screenshotInput.click();
    });

    elements.pasteScreenshotButton.addEventListener("click", () => {
      pasteScreenshotFromClipboard();
    });

    elements.screenshotInput.addEventListener("change", async (event) => {
      const files = Array.from(event.target.files || []);
      if (!files.length || !pendingScreenshotTarget?.testId) {
        return;
      }
      await attachScreenshots(pendingScreenshotTarget.testId, files, pendingScreenshotTarget.stepId || null);
      pendingScreenshotTarget = null;
      elements.screenshotInput.value = "";
      render();
      saveState();
    });

    elements.jsonInput.addEventListener("change", async (event) => {
      const [file] = Array.from(event.target.files || []);
      if (!file) {
        return;
      }
      const text = await file.text();
      importJson(text);
      elements.jsonInput.value = "";
    });

    if (elements.excelInput) {
      elements.excelInput.addEventListener("change", async (event) => {
        const [file] = Array.from(event.target.files || []);
        if (!file) {
          return;
        }
        await importSpreadsheet(file);
        elements.excelInput.value = "";
      });
    }

    ["runName", "ownerName", "queryCount", "versionDetails"].forEach((field) => {
      if (!elements[field]) return; // Skip if element doesn't exist
      elements[field].addEventListener("input", () => {
        state[field] = field === "queryCount" ? Number(elements[field].value || 0) : elements[field].value;
        renderSummary();
        saveState();
      });
    });

    if (elements.presentationModeToggle) {
      elements.presentationModeToggle.addEventListener("change", () => {
        state.presentationMode = elements.presentationModeToggle.checked;
        saveState();
      });
    }

    elements.autoCaptureToggle.addEventListener("change", () => {
      state.autoCapture = elements.autoCaptureToggle.checked;
      saveState();
    });

    const handleUndoAction = (event) => {
      event?.preventDefault();
      event?.stopPropagation();
      undoStroke();
    };
    elements.undoStrokeButton.addEventListener("pointerdown", handleUndoAction);
    elements.clearDrawingButton.addEventListener("click", clearDrawing);
    elements.saveDrawingButton.addEventListener("click", saveDrawingMarkup);
    elements.markupTool.addEventListener("change", syncToolUi);
    elements.toggleHelpButton.addEventListener("click", () => {
      elements.drawingHelpPanel.classList.toggle("hidden");
    });
    elements.selectedTextInput.addEventListener("input", updateSelectedItemText);
    elements.duplicateItemButton.addEventListener("click", duplicateSelectedItem);
    elements.copyItemButton.addEventListener("click", copySelectedItem);
    elements.pasteItemButton.addEventListener("click", pasteClipboardItem);
    bindDrawingCanvasEvents();
    bindDrawingDialogShortcuts();
    if (elements.drawingDialog) {
      const closeDrawingDialog = () => {
        const activePointerId = drawingSession?.activePointerId;
        if (typeof activePointerId === "number") {
          releaseDrawingPointerCapture(activePointerId);
        }
        suppressDialogReopenUntil = Date.now() + 260;
        if (!elements.drawingDialog.open) {
          return;
        }
        elements.drawingDialog.close();
      };
      if (elements.closeDrawingDialogButton) {
        const handleCloseDrawingAction = (event) => {
          event.preventDefault();
          event.stopPropagation();
          closeDrawingDialog();
        };
        elements.closeDrawingDialogButton.addEventListener("click", handleCloseDrawingAction);
      }
      const drawingShell = elements.drawingDialog.querySelector(".drawing-shell");
      if (drawingShell) {
        drawingShell.addEventListener("submit", (event) => {
          event.preventDefault();
          closeDrawingDialog();
        });
      }
      elements.drawingDialog.addEventListener("cancel", (event) => {
        event.preventDefault();
        closeDrawingDialog();
      });
      elements.drawingDialog.addEventListener("click", (event) => {
        if (event.target === elements.drawingDialog) {
          closeDrawingDialog();
        }
      });
    }
    elements.drawingDialog.addEventListener("close", () => {
      if (isClosingDrawingAfterSave || !drawingSession) {
        return;
      }
      persistDrawingMarkup({ closeDialog: false });
    });
    if (elements.mediaPreviewDialog) {
      const closeMediaPreviewDialog = () => {
        if (elements.mediaPreviewDialog.open) {
          elements.mediaPreviewDialog.close();
          return;
        }
        closeMediaPreview();
      };
      const mediaPreviewShell = elements.mediaPreviewDialog.querySelector(".media-preview-shell");
      if (mediaPreviewShell) {
        mediaPreviewShell.addEventListener("submit", (event) => {
          event.preventDefault();
          closeMediaPreviewDialog();
        });
      }
      if (elements.closeMediaPreviewButton) {
        elements.closeMediaPreviewButton.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          closeMediaPreviewDialog();
        });
      }
      elements.mediaPreviewDialog.addEventListener("cancel", (event) => {
        event.preventDefault();
        closeMediaPreviewDialog();
      });
      elements.mediaPreviewDialog.addEventListener("click", (event) => {
        if (event.target === elements.mediaPreviewDialog) {
          closeMediaPreviewDialog();
        }
      });
      elements.mediaPreviewDialog.addEventListener("close", () => {
        closeMediaPreview();
      });
    }
    document.addEventListener("click", (event) => {
      const path = event.composedPath ? event.composedPath() : [];
      if (elements.drawingDialog?.open && elements.closeDrawingDialogButton && path.includes(elements.closeDrawingDialogButton)) {
        event.preventDefault();
        event.stopPropagation();
        const activePointerId = drawingSession?.activePointerId;
        if (typeof activePointerId === "number") {
          releaseDrawingPointerCapture(activePointerId);
        }
        suppressDialogReopenUntil = Date.now() + 260;
        elements.drawingDialog.close();
        return;
      }
      if (elements.mediaPreviewDialog?.open && elements.closeMediaPreviewButton && path.includes(elements.closeMediaPreviewButton)) {
        event.preventDefault();
        event.stopPropagation();
        elements.mediaPreviewDialog.close();
      }
    }, true);

    document.addEventListener("pointerdown", (event) => {
      const isPointInsideRect = (element) => {
        if (!element) {
          return false;
        }
        const rect = element.getBoundingClientRect();
        return event.clientX >= rect.left
          && event.clientX <= rect.right
          && event.clientY >= rect.top
          && event.clientY <= rect.bottom;
      };

      if (elements.drawingDialog?.open) {
        if (isPointInsideRect(elements.undoStrokeButton)) {
          event.preventDefault();
          event.stopPropagation();
          undoStroke();
          return;
        }
      }

      if (elements.mediaPreviewDialog?.open && isPointInsideRect(elements.closeMediaPreviewButton)) {
        event.preventDefault();
        event.stopPropagation();
        elements.mediaPreviewDialog.close();
      }
    }, true);
    document.addEventListener("paste", handleGlobalImagePaste);

    if (elements.mainBackToTopButton) {
      window.addEventListener("scroll", () => {
        elements.mainBackToTopButton.hidden = window.scrollY < 300;
      });
      elements.mainBackToTopButton.addEventListener("click", () => {
        window.scrollTo({ top: 0, behavior: "smooth" });
      });
    }

    if (elements.fileInput) {
      elements.fileInput.addEventListener("change", async (event) => {
        const files = Array.from(event.target.files || []);
        if (!files.length || !pendingFileTarget) {
          return;
        }
        await attachFiles(pendingFileTarget, files);
        pendingFileTarget = null;
        elements.fileInput.value = "";
        render();
        saveState();
      });
    }

    document.addEventListener("keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.code === "KeyS") {
        event.preventDefault();
        void saveRunManually();
        return;
      }

      if (!((event.ctrlKey || event.metaKey) && event.shiftKey)) {
        return;
      }
      if (event.code === "KeyP") {
        event.preventDefault();
        togglePauseRecording();
        return;
      }
      if (event.code === "KeyX") {
        event.preventDefault();
        stopScreenRecording();
        return;
      }
      if (event.code === "KeyV") {
        event.preventDefault();
        startVoiceAnnotation(state.tests.at(-1)?.id);
      }
      if (event.code === "KeyS") {
        event.preventDefault();
        const latestTestId = state.tests.at(-1)?.id;
        if (!latestTestId) {
          window.alert("Create a test first, then attach screenshots to it.");
          return;
        }
        pendingScreenshotTarget = { testId: latestTestId, stepId: null };
        elements.screenshotInput.click();
      }
    });
  }

  function createTest() {
    return {
      id: crypto.randomUUID(),
      title: "",
      titleHtml: "",
      status: "not-start",
      reviewStatus: "not-reviewed",
      reviewComment: "",
      notes: "",
      notesHtml: "",
      steps: [],
      videos: [],
      screenshots: [],
      files: []
    };
  }

  function createStep() {
    return {
      id: crypto.randomUUID(),
      text: "",
      textHtml: "",
      screenshots: [],
      videos: [],
      files: []
    };
  }

  async function createFileAttachment(file) {
    const attachment = {
      id: crypto.randomUUID(),
      name: file.name,
      type: file.type || "application/octet-stream",
      sizeBytes: file.size,
      storage: "indexeddb",
      createdAt: new Date().toLocaleString()
    };
    await saveMediaBlob(attachment.id, file);
    return attachment;
  }

  async function attachFiles(testId, files) {
    const test = state.tests.find((item) => item.id === testId);
    if (!test) {
      return;
    }
    test.files = Array.isArray(test.files) ? test.files : [];
    const attachments = await Promise.all(files.map((file) => createFileAttachment(file)));
    test.files.push(...attachments);
  }

  function fileTypeIcon(mimeType) {
    if (!mimeType) return "📄";
    if (mimeType.startsWith("image/")) return "🖼️";
    if (mimeType === "application/pdf") return "📕";
    if (mimeType.includes("word") || mimeType.includes("document")) return "📝";
    if (mimeType.includes("spreadsheet") || mimeType.includes("excel")) return "📊";
    if (mimeType.includes("presentation") || mimeType.includes("powerpoint")) return "📑";
    if (mimeType.includes("zip") || mimeType.includes("archive")) return "🗜️";
    if (mimeType.startsWith("text/")) return "📄";
    return "📎";
  }

  function linkifyText(text) {
    if (!text) return "";
    const escaped = text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
    return escaped.replace(
      /(https?:\/\/[^\s<>"']+)/g,
      '<a href="$1" target="_blank" rel="noopener noreferrer" style="color:#0c6c61;text-decoration:underline;cursor:pointer;">$1</a>'
    );
  }

  let activeRichEditor = null;
  let savedRichRange = null;

  function plainTextToRichHtml(text) {
    return escapeHtml(text || "").replace(/\n/g, "<br>");
  }

  function richHtmlToPlainText(html) {
    const container = document.createElement("div");
    container.innerHTML = html || "";
    return (container.innerText || container.textContent || "").replace(/\u00a0/g, " ").trim();
  }

  function getRichTextHtml(htmlValue, plainText) {
    if (htmlValue) return htmlValue;
    return plainTextToRichHtml(plainText || "");
  }

  function getRichDisplayHtml(htmlValue, plainText, fallbackText = "") {
    if ((htmlValue || "").trim()) return htmlValue;
    return plainTextToRichHtml(plainText || fallbackText);
  }

  function updateRichFieldPlaceholder(editor) {
    if (!editor) return;
    editor.dataset.empty = richHtmlToPlainText(editor.innerHTML).trim() ? "false" : "true";
  }

  function setRichEditorValue(editor, htmlValue, plainText) {
    editor.innerHTML = getRichTextHtml(htmlValue, plainText);
    updateRichFieldPlaceholder(editor);
  }

  function rememberRichSelection(editor) {
    const selection = window.getSelection();
    if (!selection || !selection.rangeCount) return;
    const range = selection.getRangeAt(0);
    if (!editor.contains(range.commonAncestorContainer)) return;
    activeRichEditor = editor;
    savedRichRange = range.cloneRange();
  }

  function restoreRichSelection(editor) {
    if (!editor || activeRichEditor !== editor || !savedRichRange) return;
    const selection = window.getSelection();
    if (!selection) return;
    selection.removeAllRanges();
    selection.addRange(savedRichRange);
  }

  function applyRichTextCommand(editor, command, value = null) {
    if (!editor) return;
    editor.focus();
    restoreRichSelection(editor);
    document.execCommand("styleWithCSS", false, true);
    document.execCommand(command, false, value);
    rememberRichSelection(editor);
    updateRichFieldPlaceholder(editor);
  }

  function bindRichEditor(editor, options) {
    const { colorInput, boldButton, htmlValue, plainText, onChange } = options;
    setRichEditorValue(editor, htmlValue, plainText);

    const emitChange = () => {
      const html = editor.innerHTML;
      const text = richHtmlToPlainText(html);
      updateRichFieldPlaceholder(editor);
      onChange(html, text);
    };

    ["focus", "keyup", "mouseup"].forEach((eventName) => {
      editor.addEventListener(eventName, () => rememberRichSelection(editor));
    });

    editor.addEventListener("input", emitChange);
    editor.addEventListener("paste", (event) => {
      event.preventDefault();
      const text = event.clipboardData?.getData("text/plain") || "";
      document.execCommand("insertText", false, text);
      emitChange();
    });

    if (colorInput) {
      colorInput.addEventListener("input", () => {
        applyRichTextCommand(editor, "foreColor", colorInput.value);
        emitChange();
      });
    }

    if (boldButton) {
      boldButton.addEventListener("mousedown", (event) => event.preventDefault());
      boldButton.addEventListener("click", () => {
        applyRichTextCommand(editor, "bold");
        emitChange();
      });
    }
  }

  async function createScreenshot(file) {
    const shot = {
      id: crypto.randomUUID(),
      name: file.name,
      mimeType: file.type || "image/png",
      sizeBytes: file.size,
      previewWidth: 180,
      storage: "indexeddb",
      flattenedDataUrl: "",
      naturalWidth: 0,
      naturalHeight: 0,
      drawingItems: [],
      annotation: "",
      reviewComment: "",
      suggestedAnnotation: true,
      createdAt: new Date().toLocaleString()
    };
    await saveMediaBlob(shot.id, file);
    return shot;
  }

  async function attachScreenshots(testId, files, stepId = null) {
    const test = state.tests.find((item) => item.id === testId);
    if (!test) {
      return;
    }
    const targetStep = stepId ? (test.steps || []).find((step) => step.id === stepId) : null;
    if (targetStep) {
      targetStep.screenshots = Array.isArray(targetStep.screenshots) ? targetStep.screenshots : [];
    } else {
      test.screenshots = Array.isArray(test.screenshots) ? test.screenshots : [];
    }
    const shots = await Promise.all(files.map((file) => createScreenshot(file)));
    if (targetStep) {
      targetStep.screenshots.push(...shots);
      return;
    }
    test.screenshots.push(...shots);
  }

  function resolveScreenshotTarget(preferredTarget = null) {
    const fromPreferred = typeof preferredTarget === "string"
      ? { testId: preferredTarget, stepId: null }
      : (preferredTarget && preferredTarget.testId ? preferredTarget : null);
    const fromPending = pendingScreenshotTarget?.testId ? pendingScreenshotTarget : null;
    const fromSelected = selectedDropTarget?.kind === "shot" ? selectedDropTarget : null;
    const fromLatest = state.tests.at(-1)?.id ? { testId: state.tests.at(-1).id, stepId: null } : null;
    const target = fromPreferred || fromPending || fromSelected || fromLatest;
    if (!target?.testId) {
      return null;
    }
    const test = state.tests.find((item) => item.id === target.testId);
    if (!test) {
      return null;
    }
    const stepId = target.stepId && (test.steps || []).some((step) => step.id === target.stepId)
      ? target.stepId
      : null;
    return { testId: target.testId, stepId };
  }

  async function pasteScreenshotFromClipboard(preferredTarget = null) {
    const target = resolveScreenshotTarget(preferredTarget);
    if (!target?.testId) {
      window.alert("Create a test first, then paste screenshots into it.");
      return;
    }

    if (!navigator.clipboard?.read) {
      window.alert("Clipboard read is not available. Use Ctrl+V directly while focused in the app.");
      return;
    }

    try {
      const clipboardItems = await navigator.clipboard.read();
      const imageFiles = [];
      for (const item of clipboardItems) {
        const imageType = item.types.find((type) => type.startsWith("image/"));
        if (!imageType) {
          continue;
        }
        const blob = await item.getType(imageType);
        imageFiles.push(new File([blob], `pasted-${Date.now()}.png`, { type: blob.type || "image/png" }));
      }
      if (!imageFiles.length) {
        window.alert("No image found in clipboard.");
        return;
      }
      await attachScreenshots(target.testId, imageFiles, target.stepId || null);
      render();
      saveState();
    } catch {
      window.alert("Clipboard access was blocked. Try Ctrl+V directly in the app.");
    }
  }

  async function handleGlobalImagePaste(event) {
    const clipboard = event.clipboardData;
    if (!clipboard?.items?.length) {
      return;
    }
    const imageItems = Array.from(clipboard.items).filter((item) => item.kind === "file" && item.type.startsWith("image/"));
    if (!imageItems.length) {
      return;
    }

    // Resolve target: prefer explicitly selected drop zone, then focused test card, then fallback to latest test
    const activeCard = document.activeElement?.closest?.(".test-card");
    const activeCardTarget = activeCard?.dataset?.testId ? { testId: activeCard.dataset.testId, stepId: null } : null;
    const target = resolveScreenshotTarget(activeCardTarget);
    if (!target?.testId) {
      return;
    }

    // Only suppress default paste behaviour when we are actually handling an image
    event.preventDefault();
    const files = imageItems.map((item) => item.getAsFile()).filter((file) => file && file.size > 0);
    if (!files.length) {
      return;
    }

    await attachScreenshots(target.testId, files, target.stepId || null);
    render();
    saveState();
  }

  function render() {
    document.body.classList.toggle("review-mode", Boolean(state.reviewMode));
    if (elements.reviewModeButton) elements.reviewModeButton.textContent = state.reviewMode ? "Exit Review Mode" : "Review Mode";
    elements.runName.value = state.runName;
    elements.ownerName.value = state.ownerName;
    elements.queryCount.value = state.queryCount || 0;
    if (elements.versionDetails) elements.versionDetails.value = state.versionDetails;
    if (elements.presentationModeToggle) elements.presentationModeToggle.checked = Boolean(state.presentationMode);
    elements.autoCaptureToggle.checked = state.autoCapture;
    renderSummary();
    renderTemplates();
    renderReviewOverview();
    renderTests();
    refreshSelectedDropTargetUi();
    syncRecordingControls();
    // SetUp section
    if (elements.setupText) elements.setupText.innerHTML = state.setup?.textHtml || "";
    if (elements.setupLinks) elements.setupLinks.value = state.setup?.links || "";
    if (typeof renderSetupScreenshots === "function") renderSetupScreenshots();
  }

  function renderSummary() {
    const counts = state.tests.reduce((accumulator, test) => {
      accumulator.total += 1;
      if (test.status === "passed") accumulator.passed += 1;
      if (test.status === "query") accumulator.query += 1;
      if (test.status === "failed") accumulator.failed += 1;
      if (test.status === "blocked") accumulator.blocked += 1;
      if (test.status === "cancelled") accumulator.cancelled += 1;
      accumulator.screenshots += collectTestScreenshots(test).length;
      return accumulator;
    }, { total: 0, passed: 0, query: 0, failed: 0, blocked: 0, cancelled: 0, screenshots: 0 });

    const cards = [
      [counts.total, "Tests"],
      [counts.passed, "Passed"],
      [counts.query, "Query"],
      [counts.failed, "Failed"],
      [counts.blocked, "Blocked"],
      [counts.cancelled, "Cancelled"],
      [state.queryCount || 0, "Queries"],
      [counts.screenshots, "Screenshots"],
      [recognitionCtor ? "On" : "Off", "Voice capture"]
    ];

    if (elements.summaryCards) {
      elements.summaryCards.innerHTML = "";
      cards.forEach(([value, label]) => {
        const card = document.createElement("div");
        card.className = "summary-card";
        card.innerHTML = `<strong>${value}</strong><span>${label}</span>`;
        elements.summaryCards.appendChild(card);
      });
    }
  }

  function renderTemplates() {
    const templatesList = document.getElementById("templatesList");
    const noTemplates = document.getElementById("noTemplates");
    if (!templatesList || !noTemplates) return;

    const templates = state.templates || [];
    templatesList.innerHTML = "";
    noTemplates.hidden = templates.length > 0;

    templates.forEach((template) => {
      const btn = document.createElement("button");
      btn.className = "secondary";
      btn.style.fontSize = "0.9rem";
      btn.style.padding = "8px 12px";
      btn.textContent = template.templateName || template.name || "Unnamed";
      btn.type = "button";
      btn.title = `${template.steps?.length || 0} steps`;
      btn.addEventListener("click", () => {
        const newTest = structuredClone(template);
        newTest.id = crypto.randomUUID();
        newTest.status = "not-start";
        newTest.reviewStatus = "not-reviewed";
        newTest.reviewComment = "";
        newTest.videos = [];
        newTest.screenshots = [];
        newTest.files = [];
        state.tests.push(newTest);
        render();
        saveState();
        window.alert(`Created test from template: ${template.templateName || template.name || "Unnamed"}`);
      });

      const deleteBtn = document.createElement("button");
      deleteBtn.className = "ghost";
      deleteBtn.style.fontSize = "0.8rem";
      deleteBtn.style.padding = "4px 8px";
      deleteBtn.textContent = "×";
      deleteBtn.type = "button";
      deleteBtn.title = "Delete template";
      deleteBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        if (window.confirm(`Delete template "${template.templateName || template.name || "Unnamed"}"?`)) {
          state.templates = (state.templates || []).filter((t) => t.id !== template.id);
          render();
          saveState();
        }
      });

      const wrapper = document.createElement("div");
      wrapper.style.display = "flex";
      wrapper.style.gap = "6px";
      wrapper.appendChild(btn);
      wrapper.appendChild(deleteBtn);
      templatesList.appendChild(wrapper);
    });
  }

  function renderReviewOverview() {
    if (!elements.reviewOverview) {
      console.warn("reviewOverview element not found");
      return;
    }
    const tests = state.tests || [];
    renderFloatingSummaryNav(tests);

    elements.reviewOverview.classList.toggle("hidden", tests.length === 0);
    const tableRows = tests.map((test, i) => `
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;text-align:center;"><a href="#test-card-${escapeHtml(test.id)}" style="color:#0c6c61;font-weight:700;text-decoration:none;">${i + 1}</a></td>
        <td style="padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;font-weight:600;">${getRichDisplayHtml(test.titleHtml, test.title, "Untitled test")}</td>
        <td style="${statusCellStyle(test.status)}">${escapeHtml(formatTestStatus(test.status))}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;white-space:pre-wrap;">${getRichDisplayHtml(test.notesHtml, test.notes, "-")}</td>
      </tr>`).join("");
    elements.reviewOverview.innerHTML = `
      <h3>Consolidated Test Summary</h3>
      <div style="overflow-x:auto;margin-top:10px;">
        <table style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #e2d7c0;border-radius:12px;overflow:hidden;">
          <thead>
            <tr style="background:#f6f1e7;text-align:left;">
              <th style="padding:10px 12px;border-bottom:1px solid #e2d7c0;width:60px;">No.</th>
              <th style="padding:10px 12px;border-bottom:1px solid #e2d7c0;">Test Scenario</th>
              <th style="padding:10px 12px;border-bottom:1px solid #e2d7c0;width:120px;">Status</th>
              <th style="padding:10px 12px;border-bottom:1px solid #e2d7c0;">UAT Results</th>
            </tr>
          </thead>
          <tbody>${tableRows}</tbody>
        </table>
      </div>`;
  }

  function renderFloatingSummaryNav(tests) {
    if (!elements.floatingSummaryNav || !elements.floatingSummaryNavBody) {
      document.body.classList.remove("has-floating-summary-nav");
      return;
    }

    elements.floatingSummaryNav.classList.toggle("hidden", !tests.length);
    document.body.classList.toggle("has-floating-summary-nav", tests.length > 0);
    if (!tests.length) {
      elements.floatingSummaryNavBody.innerHTML = '<p class="floating-summary-nav-empty">No tests yet.</p>';
      return;
    }

    const tableRows = tests.map((test, i) => {
      const statusClass = `status-${normalizeTestStatus(test.status)}`;
      return `
        <tr>
          <td><a href="#test-card-${escapeHtml(test.id)}">${i + 1}</a></td>
          <td><a href="#test-card-${escapeHtml(test.id)}">${getRichDisplayHtml(test.titleHtml, test.title, "Untitled test")}</a></td>
          <td><span class="floating-summary-nav-status ${statusClass}">${escapeHtml(formatTestStatus(test.status))}</span></td>
        </tr>`;
    }).join("");

    elements.floatingSummaryNavBody.innerHTML = `
      <table class="floating-summary-nav-table">
        <thead>
          <tr>
            <th>NO</th>
            <th>Test Scenario</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>${tableRows}</tbody>
      </table>`;
  }

  function renderTests() {
    if (!elements.testsList) {
      console.warn("testsList element not found");
      return;
    }
    elements.testsList.innerHTML = "";

    if (!state.tests.length) {
      const empty = document.createElement("div");
      empty.className = "summary-card";
      empty.innerHTML = "<strong>No tests yet</strong><span>Create one and start capturing steps, screenshots, and voice notes.</span>";
      elements.testsList.appendChild(empty);
      return;
    }

    if (!elements.testCardTemplate) {
      console.warn("testCardTemplate not found");
      return;
    }

    state.tests.forEach((test, testIndex) => {
      const fragment = elements.testCardTemplate.content.cloneNode(true);
      const card = fragment.querySelector(".test-card");
      const testSequenceNumber = fragment.querySelector(".test-seq-number");
      const title = fragment.querySelector(".test-title");
      const titleColor = fragment.querySelector(".test-title-color");
      const titleBold = fragment.querySelector(".test-title-bold");
      const status = fragment.querySelector(".test-status");
      const notes = fragment.querySelector(".test-notes");
      const notesColor = fragment.querySelector(".test-notes-color");
      const notesBold = fragment.querySelector(".test-notes-bold");
      const notesPreview = fragment.querySelector(".test-notes-preview");
      const stepsList = fragment.querySelector(".steps-list");
      const videosList = fragment.querySelector(".videos-list");
      const shotsList = fragment.querySelector(".shots-list");
      const reviewRow = fragment.querySelector(".review-row");
      const reviewStatus = fragment.querySelector(".review-status");
      const reviewStatusPill = fragment.querySelector(".review-status-pill");
      const testReviewComment = fragment.querySelector(".test-review-comment");

      card.dataset.testId = test.id;
      card.id = `test-card-${test.id}`;
      bindTestDragEvents(card, test.id);

      testSequenceNumber.textContent = `Test ${testIndex + 1}`;
      status.value = normalizeTestStatus(test.status);
      applyTestStatusSelectStyle(status, status.value);
      bindRichEditor(title, {
        colorInput: titleColor,
        boldButton: titleBold,
        htmlValue: test.titleHtml,
        plainText: test.title,
        onChange: (html, text) => {
          test.titleHtml = html;
          test.title = text;
          renderSummary();
          saveState();
        }
      });
      bindRichEditor(notes, {
        colorInput: notesColor,
        boldButton: notesBold,
        htmlValue: test.notesHtml,
        plainText: test.notes,
        onChange: (html, text) => {
          test.notesHtml = html;
          test.notes = text;
          updateTestNotesLinkPreview(text);
          renderSummary();
          saveState();
        }
      });
      function updateTestNotesLinkPreview(text) {
        if (!notesPreview) return;
        const hasLink = /https?:\/\//.test(text || "");
        notesPreview.hidden = !hasLink;
        if (hasLink) notesPreview.innerHTML = linkifyText(text);
      }
      updateTestNotesLinkPreview(test.notes || "");
      reviewRow.classList.toggle("hidden", !state.reviewMode);
      reviewStatus.value = test.reviewStatus || "not-reviewed";
      reviewStatusPill.textContent = formatReviewStatus(test.reviewStatus || "not-reviewed");
      reviewStatusPill.className = `review-status-pill ${test.reviewStatus || "not-reviewed"}`;
      testReviewComment.classList.toggle("hidden", !state.reviewMode);
      testReviewComment.value = test.reviewComment || "";

      status.addEventListener("change", () => {
        applyTestStatusSelectStyle(status, status.value);
        updateTest(test.id, { status: status.value });
      });
      reviewStatus.addEventListener("change", () => updateTest(test.id, { reviewStatus: reviewStatus.value }));
      testReviewComment.addEventListener("input", () => updateTest(test.id, { reviewComment: testReviewComment.value }));

      fragment.querySelector(".add-step").addEventListener("click", () => {
        test.steps.push(createStep());
        render();
        saveState();
      });

      fragment.querySelector(".record-video").addEventListener("click", () => {
        startScreenRecording({ testId: test.id });
      });

      const recordPauseBtn = fragment.querySelector(".record-pause");
      const recordStopBtn = fragment.querySelector(".record-stop");
      const recordVideoBtn = fragment.querySelector(".record-video");

      recordPauseBtn.addEventListener("click", () => {
        togglePauseRecording();
        const state = recordingSession?.recorder?.state || "inactive";
        recordPauseBtn.textContent = state === "paused" ? "Resume" : "Pause";
      });

      recordStopBtn.addEventListener("click", () => {
        stopScreenRecording();
      });

      fragment.querySelector(".attach-shot").addEventListener("click", () => {
        const preferredTarget = selectedDropTarget?.kind === "shot" && selectedDropTarget.testId === test.id
          ? { testId: selectedDropTarget.testId, stepId: selectedDropTarget.stepId || null }
          : { testId: test.id, stepId: null };
        pendingScreenshotTarget = preferredTarget;
        elements.screenshotInput.click();
      });

      fragment.querySelector(".paste-shot").addEventListener("click", () => {
        pasteScreenshotFromClipboard(test.id);
      });

      fragment.querySelector(".record-voice").addEventListener("click", () => {
        startVoiceAnnotation(test.id);
      });

      fragment.querySelector(".remove-test").addEventListener("click", () => {
        void (async () => {
          await deleteVideosForTests([test]);
          await deleteBinaryMediaForTests([test]);
          state.tests = state.tests.filter((item) => item.id !== test.id);
          render();
          saveState();
        })();
      });

      const saveTemplateBtn = document.createElement("button");
      saveTemplateBtn.className = "secondary";
      saveTemplateBtn.textContent = "Save as Template";
      saveTemplateBtn.type = "button";
      saveTemplateBtn.addEventListener("click", () => {
        const templateName = window.prompt("Enter a name for this template:", test.title || "Untitled Template");
        if (!templateName) return;
        const template = structuredClone(test);
        template.templateName = templateName;
        state.templates = (state.templates || []).concat(template);
        render();
        saveState();
        window.alert(`Saved template: ${templateName}`);
      });
      fragment.querySelector(".test-actions").appendChild(saveTemplateBtn);

      bindAttachmentDropZone(shotsList, { kind: "shot", testId: test.id, stepId: null, emptyLabel: "Drop screenshots here" });
      bindAttachmentDropZone(videosList, { kind: "video", testId: test.id, stepId: null, emptyLabel: "Drop videos here" });

      test.steps.forEach((step, index) => {
        const stepFragment = elements.stepTemplate.content.cloneNode(true);
        const row = stepFragment.querySelector(".step-row");
        const stepLinkPreview = stepFragment.querySelector(".step-link-preview");
        const stepShots = stepFragment.querySelector(".step-shots");
        const stepVideos = stepFragment.querySelector(".step-videos");
        row.dataset.stepId = step.id;
        bindStepDragEvents(row, test.id, step.id);
        row.querySelector(".step-index").textContent = String(index + 1);
        const input = row.querySelector(".step-input");
        const stepColor = row.querySelector(".step-color");
        const stepBold = row.querySelector(".step-bold");
        step.screenshots = Array.isArray(step.screenshots) ? step.screenshots : [];
        step.videos = Array.isArray(step.videos) ? step.videos : [];
        bindAttachmentDropZone(stepShots, { kind: "shot", testId: test.id, stepId: step.id, emptyLabel: "Drop screenshots here" });
        bindAttachmentDropZone(stepVideos, { kind: "video", testId: test.id, stepId: step.id, emptyLabel: "Drop videos here" });
        bindRichEditor(input, {
          colorInput: stepColor,
          boldButton: stepBold,
          htmlValue: step.textHtml,
          plainText: step.text,
          onChange: (html, text) => {
            step.textHtml = html;
            step.text = text;
            updateStepLinkPreview(text);
            saveState();
          }
        });
        function updateStepLinkPreview(text) {
          if (!stepLinkPreview) return;
          const hasLink = /https?:\/\//.test(text || "");
          stepLinkPreview.hidden = !hasLink;
          if (hasLink) stepLinkPreview.innerHTML = linkifyText(text);
        }
        updateStepLinkPreview(step.text || "");
        row.querySelector(".record-step").addEventListener("click", () => {
          startScreenRecording({ testId: test.id, stepId: step.id });
        });
        row.querySelector(".add-step-inline").addEventListener("click", () => {
          const currentStepIndex = test.steps.findIndex((item) => item.id === step.id);
          const insertAt = currentStepIndex >= 0 ? currentStepIndex + 1 : test.steps.length;
          test.steps.splice(insertAt, 0, createStep());
          render();
          saveState();
        });
        row.querySelector(".remove-step").addEventListener("click", () => {
          void (async () => {
            await deleteVideos(step.videos || []);
            test.steps = test.steps.filter((item) => item.id !== step.id);
            render();
            saveState();
          })();
        });

        step.screenshots.forEach((shot) => {
          appendShotCard(stepShots, shot, { testId: test.id, stepId: step.id });
        });

        step.videos.forEach((video) => {
          appendVideoCard(stepVideos, video, {
            onNotes: (value) => {
              video.notes = value;
              saveState();
            },
            onRemove: () => {
              void (async () => {
                await deleteVideoBlob(video.id);
                step.videos = step.videos.filter((item) => item.id !== video.id);
                render();
                saveState();
              })();
            }
          }, { testId: test.id, stepId: step.id });
        });

        stepsList.appendChild(stepFragment);
      });

      (test.videos || []).forEach((video) => {
        appendVideoCard(videosList, video, {
          onNotes: (value) => {
            video.notes = value;
            saveState();
          },
          onRemove: () => {
            void (async () => {
              await deleteVideoBlob(video.id);
              test.videos = (test.videos || []).filter((item) => item.id !== video.id);
              render();
              saveState();
            })();
          }
        }, { testId: test.id, stepId: null });
      });

      test.screenshots.forEach((shot) => {
        appendShotCard(shotsList, shot, { testId: test.id, stepId: null });
      });

      const filesList = fragment.querySelector(".files-list");
      (test.files || []).forEach((attachment) => {
        appendFileCard(filesList, attachment, test.id);
      });

      fragment.querySelector(".attach-file").addEventListener("click", () => {
        pendingFileTarget = test.id;
        elements.fileInput.click();
      });

      elements.testsList.appendChild(fragment);
    });
  }

  function updateTest(testId, changes) {
    const test = state.tests.find((item) => item.id === testId);
    if (!test) {
      return;
    }
    Object.assign(test, changes);
    renderSummary();
    saveState();
  }

  function appendShotCard(container, shot, placement) {
    const shotFragment = elements.shotTemplate.content.cloneNode(true);
    const shotCard = shotFragment.querySelector(".shot-card");
    shotCard.dataset.shotId = shot.id;
    bindShotDragEvents(shotCard, placement.testId, shot.id, placement.stepId || null);
    const normalizeShotPreviewWidth = (value) => {
      const numericValue = Number(value);
      if (!Number.isFinite(numericValue)) {
        return 180;
      }
      return Math.min(520, Math.max(120, Math.round(numericValue / 10) * 10));
    };
    const applyShotPreviewWidth = (value) => {
      const safeWidth = normalizeShotPreviewWidth(value);
      shot.previewWidth = safeWidth;
      shotCard.style.gridTemplateColumns = `${safeWidth}px minmax(0, 1fr)`;
      return safeWidth;
    };
    const initialPreviewWidth = applyShotPreviewWidth(shot.previewWidth);
    const shotImage = shotFragment.querySelector(".shot-image");
    void (async () => {
      const src = await getShotDisplayUrl(shot);
      if (src) {
        shotImage.src = src;
      }
    })();
    shotImage.addEventListener("dblclick", async () => {
      const src = await getShotDisplayUrl(shot);
      if (!src) {
        window.alert("Screenshot data is not available in this browser.");
        return;
      }
      openMediaPreview({
        kind: "image",
        src,
        title: shot.name || "Screenshot"
      });
    });
    shotFragment.querySelector(".shot-name").textContent = shot.name;
    shotFragment.querySelector(".shot-meta").textContent = `Attached ${shot.createdAt}`;
    const annotation = shotFragment.querySelector(".shot-annotation");
    const annotationPreview = shotFragment.querySelector(".link-preview");
    annotation.disabled = state.reviewMode;
    annotation.value = shot.annotation;
    function updateShotLinkPreview(text) {
      if (!annotationPreview) return;
      const hasLink = /https?:\/\//.test(text);
      annotationPreview.hidden = !hasLink;
      if (hasLink) annotationPreview.innerHTML = linkifyText(text);
    }
    updateShotLinkPreview(shot.annotation);
    annotation.addEventListener("input", () => {
      shot.annotation = annotation.value;
      shot.suggestedAnnotation = false;
      updateShotLinkPreview(annotation.value);
      saveState();
    });
    const suggestion = shotFragment.querySelector(".suggestion");
    suggestion.classList.toggle("hidden", !shot.suggestedAnnotation);
    const reviewComment = shotFragment.querySelector(".review-comment");
    reviewComment.classList.toggle("hidden", !state.reviewMode);
    reviewComment.value = shot.reviewComment || "";
    reviewComment.addEventListener("input", () => {
      shot.reviewComment = reviewComment.value;
      saveState();
    });
    const shotSizeRange = shotFragment.querySelector(".shot-size-range");
    const shotSizeValue = shotFragment.querySelector(".shot-size-value");
    if (shotSizeRange && shotSizeValue) {
      shotSizeRange.value = String(initialPreviewWidth);
      shotSizeValue.textContent = `${initialPreviewWidth}px`;
      shotSizeRange.addEventListener("input", () => {
        const safeWidth = applyShotPreviewWidth(shotSizeRange.value);
        shotSizeValue.textContent = `${safeWidth}px`;
        saveState();
      });
    }
    shotFragment.querySelector(".edit-shot").addEventListener("click", () => {
      openDrawingDialog(placement.testId, shot.id);
    });
    shotFragment.querySelector(".edit-shot").disabled = state.reviewMode;
    shotFragment.querySelector(".remove-shot").addEventListener("click", () => {
      void deleteMediaBlob(shot.id);
      removeShotFromTest(placement.testId, shot.id);
      render();
      saveState();
    });
    bindMoveControl(shotFragment, {
      testId: placement.testId,
      currentStepId: placement.stepId || null,
      kind: "shot",
      itemId: shot.id
    });
    container.appendChild(shotFragment);
  }

  function appendVideoCard(container, video, handlers, placement) {
    const videoFragment = elements.videoTemplate.content.cloneNode(true);
    const preview = videoFragment.querySelector(".video-preview");
    const videoCard = videoFragment.querySelector(".video-card");
    videoCard.draggable = true;
    bindVideoDragEvents(videoCard, placement.testId, video.id, placement.stepId || null);
    preview.preload = "metadata";
    preview.playsInline = true;
    void loadVideoPreview(preview, video);
    preview.addEventListener("dblclick", async () => {
      const src = await getVideoPlaybackUrl(video);
      openMediaPreview({
        kind: "video",
        src,
        title: video.name || "Video"
      });
    });
    videoFragment.querySelector(".video-name").textContent = video.name;
    videoFragment.querySelector(".video-meta").textContent = formatVideoMeta(video);
    // Subject field with color
    const subjectInput = videoFragment.querySelector(".video-subject");
    const subjectColor = videoFragment.querySelector(".video-subject-color");
    const subjectBold = videoFragment.querySelector(".video-subject-bold");
    bindRichEditor(subjectInput, {
      colorInput: subjectColor,
      boldButton: subjectBold,
      htmlValue: video.subjectHtml,
      plainText: video.subject,
      onChange: (html, text) => {
        video.subjectHtml = html;
        video.subject = text;
        saveState();
      }
    });
    // Notes field with color
    const notesInput = videoFragment.querySelector(".video-notes");
    const notesColor = videoFragment.querySelector(".video-notes-color");
    const notesBold = videoFragment.querySelector(".video-notes-bold");
    const notesPreview = videoFragment.querySelector(".link-preview");
    bindRichEditor(notesInput, {
      colorInput: notesColor,
      boldButton: notesBold,
      htmlValue: video.notesHtml,
      plainText: video.notes,
      onChange: (html, text) => {
        video.notesHtml = html;
        handlers.onNotes(text);
        updateVideoLinkPreview(text);
        saveState();
      }
    });
    function updateVideoLinkPreview(text) {
      if (!notesPreview) return;
      const hasLink = /https?:\/\//.test(text);
      notesPreview.hidden = !hasLink;
      if (hasLink) notesPreview.innerHTML = linkifyText(text);
    }
    updateVideoLinkPreview(video.notes || "");
    videoFragment.querySelector(".remove-video").addEventListener("click", () => {
      handlers.onRemove();
    });
    bindMoveControl(videoFragment, {
      testId: placement.testId,
      currentStepId: placement.stepId || null,
      kind: "video",
      itemId: video.id
    });
    container.appendChild(videoFragment);
  }

  function closeMediaPreview() {
    if (elements.mediaPreviewVideo) {
      elements.mediaPreviewVideo.pause();
      elements.mediaPreviewVideo.removeAttribute("src");
      elements.mediaPreviewVideo.load();
      elements.mediaPreviewVideo.classList.add("hidden");
    }
    if (elements.mediaPreviewImage) {
      elements.mediaPreviewImage.removeAttribute("src");
      elements.mediaPreviewImage.classList.add("hidden");
    }
  }

  function openMediaPreview({ kind, src, title }) {
    if (!elements.mediaPreviewDialog || !src) {
      return;
    }
    if (elements.mediaPreviewTitle) {
      elements.mediaPreviewTitle.textContent = title || "Preview";
    }
    if (kind === "video") {
      elements.mediaPreviewImage?.classList.add("hidden");
      if (elements.mediaPreviewVideo) {
        elements.mediaPreviewVideo.src = src;
        elements.mediaPreviewVideo.classList.remove("hidden");
      }
    } else {
      elements.mediaPreviewVideo?.classList.add("hidden");
      if (elements.mediaPreviewImage) {
        elements.mediaPreviewImage.src = src;
        elements.mediaPreviewImage.classList.remove("hidden");
      }
    }
    elements.mediaPreviewDialog.showModal();
  }

  function appendFileCard(container, attachment, testId) {
    const card = document.createElement("div");
    card.className = "file-card";
    card.innerHTML = `
      <span class="file-icon">${fileTypeIcon(attachment.type)}</span>
      <div class="file-info">
        <a class="file-name" href="#">${escapeHtml(attachment.name)}</a>
        <span class="file-meta">${escapeHtml(formatBytes(attachment.sizeBytes))} &mdash; ${escapeHtml(attachment.createdAt)}</span>
      </div>
      <button class="ghost remove-file" type="button" title="Remove">&#x2715;</button>
    `;
    const fileLink = card.querySelector(".file-name");
    fileLink.addEventListener("click", async (event) => {
      event.preventDefault();
      const blob = await getFileBlobForAttachment(attachment);
      if (!blob) {
        window.alert("File data is not available in this browser.");
        return;
      }
      const objectUrl = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = objectUrl;
      anchor.download = attachment.name || "attachment";
      anchor.click();
      setTimeout(() => URL.revokeObjectURL(objectUrl), 500);
    });
    card.querySelector(".remove-file").addEventListener("click", () => {
      const test = state.tests.find((item) => item.id === testId);
      if (!test) return;
      void deleteMediaBlob(attachment.id);
      test.files = (test.files || []).filter((item) => item.id !== attachment.id);
      render();
      saveState();
    });
    container.appendChild(card);
  }

  function bindMoveControl(fragment, moveContext) {
    const select = fragment.querySelector(".move-target");
    const button = fragment.querySelector(".move-media");
    if (!select || !button) {
      return;
    }
    populateMoveOptions(select, moveContext.testId, moveContext.currentStepId);
    button.addEventListener("click", () => {
      const val = select.value;
      const separatorIndex = val.indexOf("::");
      if (separatorIndex < 0) return;
      const targetTestId = val.slice(0, separatorIndex);
      const targetStepId = val.slice(separatorIndex + 2) || null;
      if (
        targetTestId === moveContext.testId &&
        (targetStepId || null) === (moveContext.currentStepId || null)
      ) {
        return;
      }
      moveAttachmentCrossTest(
        moveContext.testId,
        targetTestId,
        moveContext.kind,
        moveContext.currentStepId || null,
        moveContext.itemId,
        targetStepId
      );
      render();
      saveState();
    });
  }

  function populateMoveOptions(select, currentTestId, currentStepId) {
    select.innerHTML = "";
    state.tests.forEach((test) => {
      const group = document.createElement("optgroup");
      group.label = test.title || "Untitled Test";
      const testOpt = document.createElement("option");
      testOpt.value = test.id + "::";
      testOpt.textContent = "\u2014 Test level (no step)";
      group.appendChild(testOpt);
      (test.steps || []).forEach((step, index) => {
        const opt = document.createElement("option");
        opt.value = test.id + "::" + step.id;
        opt.textContent = "Step " + (index + 1) + (step.text ? " - " + step.text.slice(0, 28) : "");
        group.appendChild(opt);
      });
      select.appendChild(group);
    });
    select.value = currentTestId + "::" + (currentStepId || "");
  }

  function formatVideoMeta(video) {
    const segments = [
      `Recorded ${video.createdAt || "-"}`,
      video.durationSeconds != null ? `Duration ${formatDuration(video.durationSeconds)}` : null,
      video.sizeBytes != null ? `Size ${formatBytes(video.sizeBytes)}` : null
    ].filter(Boolean);
    return segments.join(" | ");
  }

  function formatDuration(totalSeconds) {
    const seconds = Math.max(0, Math.round(totalSeconds));
    const minutes = Math.floor(seconds / 60);
    const remainder = seconds % 60;
    return `${String(minutes).padStart(2, "0")}:${String(remainder).padStart(2, "0")}`;
  }

  function formatBytes(bytes) {
    if (!Number.isFinite(bytes) || bytes < 0) {
      return "-";
    }
    const units = ["B", "KB", "MB", "GB"];
    let value = bytes;
    let unitIndex = 0;
    while (value >= 1024 && unitIndex < units.length - 1) {
      value /= 1024;
      unitIndex += 1;
    }
    return `${value.toFixed(unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
  }

  function collectTestScreenshots(test) {
    return [
      ...(test?.screenshots || []),
      ...((test?.steps || []).flatMap((step) => step.screenshots || []))
    ];
  }

  function getStepById(test, stepId) {
    return (test?.steps || []).find((step) => step.id === stepId) || null;
  }

  function getAttachmentCollection(test, kind, stepId = null) {
    if (!test) {
      return null;
    }
    const property = kind === "shot" ? "screenshots" : "videos";
    if (!stepId) {
      test[property] = Array.isArray(test[property]) ? test[property] : [];
      return test[property];
    }
    const step = getStepById(test, stepId);
    if (!step) {
      return null;
    }
    step[property] = Array.isArray(step[property]) ? step[property] : [];
    return step[property];
  }

  function removeShotFromTest(testId, shotId) {
    const test = state.tests.find((item) => item.id === testId);
    if (!test) {
      return;
    }
    test.screenshots = (test.screenshots || []).filter((item) => item.id !== shotId);
    (test.steps || []).forEach((step) => {
      step.screenshots = (step.screenshots || []).filter((item) => item.id !== shotId);
    });
  }

  function moveAttachment(testId, kind, sourceStepId, sourceId, targetStepId, targetId = null) {
    const test = state.tests.find((item) => item.id === testId);
    if (!test) {
      return;
    }
    const sourceCollection = getAttachmentCollection(test, kind, sourceStepId);
    const targetCollection = getAttachmentCollection(test, kind, targetStepId);
    if (!sourceCollection || !targetCollection) {
      return;
    }
    const sourceIndex = sourceCollection.findIndex((item) => item.id === sourceId);
    if (sourceIndex < 0) {
      return;
    }
    const [item] = sourceCollection.splice(sourceIndex, 1);
    if (!targetId) {
      targetCollection.push(item);
      return;
    }
    const targetIndex = targetCollection.findIndex((entry) => entry.id === targetId);
    if (targetIndex < 0) {
      targetCollection.push(item);
      return;
    }
    targetCollection.splice(targetIndex, 0, item);
  }

  function moveAttachmentCrossTest(sourceTestId, targetTestId, kind, sourceStepId, sourceId, targetStepId) {
    if (sourceTestId === targetTestId) {
      moveAttachment(sourceTestId, kind, sourceStepId, sourceId, targetStepId);
      return;
    }
    const sourceTest = state.tests.find((item) => item.id === sourceTestId);
    const targetTest = state.tests.find((item) => item.id === targetTestId);
    if (!sourceTest || !targetTest) {
      return;
    }
    const sourceCollection = getAttachmentCollection(sourceTest, kind, sourceStepId);
    const targetCollection = getAttachmentCollection(targetTest, kind, targetStepId);
    if (!sourceCollection || !targetCollection) {
      return;
    }
    const sourceIndex = sourceCollection.findIndex((item) => item.id === sourceId);
    if (sourceIndex < 0) {
      return;
    }
    const [item] = sourceCollection.splice(sourceIndex, 1);
    targetCollection.push(item);
  }

  async function loadVideoPreview(element, video) {
    const playbackUrl = await getVideoPlaybackUrl(video);
    if (!playbackUrl) {
      element.removeAttribute("src");
      element.load();
      return;
    }
    element.src = playbackUrl;
    element.load();
  }

  async function getVideoPlaybackUrl(video) {
    if (!video?.id) {
      return video?.dataUrl || "";
    }
    if (mediaObjectUrls.has(video.id)) {
      return mediaObjectUrls.get(video.id);
    }
    const blob = await getVideoBlobForVideo(video);
    if (!blob) {
      return video?.dataUrl || "";
    }
    const objectUrl = URL.createObjectURL(blob);
    mediaObjectUrls.set(video.id, objectUrl);
    return objectUrl;
  }

  function revokeVideoPlaybackUrl(videoId) {
    revokeMediaPlaybackUrl(videoId);
  }

  function dataUrlToBlob(dataUrl) {
    const [header, base64] = dataUrl.split(",");
    const mimeType = header.match(/data:(.*?);base64/)?.[1] || "application/octet-stream";
    const binary = window.atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return new Blob([bytes], { type: mimeType });
  }

  async function getVideoBlobForVideo(video) {
    if (!video) {
      return null;
    }
    if (video.dataUrl?.startsWith("data:")) {
      return dataUrlToBlob(video.dataUrl);
    }
    try {
      return await getVideoBlob(video.id);
    } catch {
      return null;
    }
  }

  async function getVideoDataUrlForExport(video) {
    const blob = await getVideoBlobForVideo(video);
    if (!blob) {
      return "";
    }
    return readBlobAsDataUrl(blob);
  }

  async function getScreenshotBlobForShot(shot) {
    if (!shot) {
      return null;
    }
    if (shot.dataUrl?.startsWith("data:")) {
      return dataUrlToBlob(shot.dataUrl);
    }
    try {
      return await getMediaBlob(shot.id);
    } catch {
      return null;
    }
  }

  async function getFileBlobForAttachment(attachment) {
    if (!attachment) {
      return null;
    }
    if (attachment.dataUrl?.startsWith("data:")) {
      return dataUrlToBlob(attachment.dataUrl);
    }
    try {
      return await getMediaBlob(attachment.id);
    } catch {
      return null;
    }
  }

  async function getShotDataUrlForExport(shot) {
    const blob = await getScreenshotBlobForShot(shot);
    if (!blob) {
      return "";
    }
    return readBlobAsDataUrl(blob);
  }

  async function getAttachmentDataUrlForExport(attachment) {
    const blob = await getFileBlobForAttachment(attachment);
    if (!blob) {
      return "";
    }
    return readBlobAsDataUrl(blob);
  }

  async function getShotDisplayUrl(shot) {
    if (!shot) {
      return "";
    }
    if (shot.flattenedDataUrl) {
      return shot.flattenedDataUrl;
    }
    if (shot.dataUrl?.startsWith("data:")) {
      return shot.dataUrl;
    }
    if (shot.id && mediaObjectUrls.has(shot.id)) {
      return mediaObjectUrls.get(shot.id);
    }
    const blob = await getScreenshotBlobForShot(shot);
    if (!blob) {
      return "";
    }
    const objectUrl = URL.createObjectURL(blob);
    mediaObjectUrls.set(shot.id, objectUrl);
    return objectUrl;
  }

  function startVoiceAnnotation(testId) {
    if (!testId) {
      window.alert("Create a test and attach a screenshot before starting voice annotation.");
      return;
    }
    const test = state.tests.find((item) => item.id === testId);
    const lastShot = collectTestScreenshots(test).at(-1);
    if (!test || !lastShot) {
      window.alert("Attach a screenshot first. Voice annotation is attached to the latest screenshot.");
      return;
    }
    if (!recognitionCtor) {
      return;
    }

    const recognition = new recognitionCtor();
    recognition.lang = "en-AU";
    recognition.interimResults = false;
    recognition.maxAlternatives = 1;

    activeVoiceContext = { testId, shotId: lastShot.id };
    if (elements.saveStatus) elements.saveStatus.textContent = "Listening...";

    recognition.onresult = (event) => {
      const transcript = event.results[0][0].transcript.trim();
      const targetTest = state.tests.find((item) => item.id === activeVoiceContext.testId);
      const targetShot = targetTest?.screenshots.find((item) => item.id === activeVoiceContext.shotId);
      if (!targetShot) {
        return;
      }
      targetShot.annotation = transcript;
      targetShot.suggestedAnnotation = true;
      render();
      saveState();
    };

    recognition.onerror = () => {
      if (elements.saveStatus) elements.saveStatus.textContent = "Voice capture failed";
    };

    recognition.onend = () => {
      activeVoiceContext = null;
      if (elements.saveStatus) elements.saveStatus.textContent = "Saved locally";
    };

    recognition.start();
  }

  function bindDrawingCanvasEvents() {
    const canvas = elements.drawingCanvas;
    canvas.addEventListener("pointerdown", handleDrawStart);
    canvas.addEventListener("pointermove", handleDrawMove);
    canvas.addEventListener("pointerup", handleDrawEnd);
    canvas.addEventListener("pointerleave", handleDrawEnd);
    canvas.addEventListener("pointercancel", handleDrawEnd);
  }

  async function openDrawingDialog(testId, shotId) {
    if (Date.now() < suppressDialogReopenUntil) {
      return;
    }
    const shot = findShot(testId, shotId);
    if (!shot) {
      return;
    }

    const sourceDataUrl = await getShotDataUrlForExport(shot);
    if (!sourceDataUrl) {
      window.alert("Screenshot data is not available in this browser.");
      return;
    }

    const imageMeta = await loadImageDimensions(sourceDataUrl);
    shot.naturalWidth = shot.naturalWidth || imageMeta.width;
    shot.naturalHeight = shot.naturalHeight || imageMeta.height;

    drawingSession = {
      testId,
      shotId,
      naturalWidth: shot.naturalWidth,
      naturalHeight: shot.naturalHeight,
      items: structuredClone(shot.drawingItems || shot.drawingPaths || []),
      draftItem: null,
      activePointerId: null,
      selectedIndex: -1,
      interactionMode: null,
      activeHandle: null,
      dragOrigin: null,
      snapGuides: [],
      dirty: false,
      baseImageDataUrl: sourceDataUrl
    };

    elements.drawingTitle.textContent = `Annotate ${shot.name}`;
    elements.drawingImage.src = sourceDataUrl;
    elements.markupTool.value = "pen";
    elements.textInput.value = "";
    elements.selectedTextInput.value = "";
    elements.drawingHelpPanel.classList.add("hidden");
    syncToolUi();
    elements.drawingDialog.showModal();
    await whenImageReady(elements.drawingImage);
    await new Promise((resolve) => window.requestAnimationFrame(resolve));
    sizeDrawingSurface();
    redrawDrawingCanvas();
  }

  function sizeDrawingSurface() {
    if (!drawingSession) {
      return;
    }
    const imageRect = elements.drawingImage.getBoundingClientRect();
    const canvas = elements.drawingCanvas;
    canvas.width = Math.round(imageRect.width);
    canvas.height = Math.round(imageRect.height);
    canvas.style.width = `${Math.round(imageRect.width)}px`;
    canvas.style.height = `${Math.round(imageRect.height)}px`;
  }

  function redrawDrawingCanvas() {
    const canvas = elements.drawingCanvas;
    const context = canvas.getContext("2d");
    context.clearRect(0, 0, canvas.width, canvas.height);
    if (!drawingSession) {
      return;
    }
    drawingSession.items.forEach((item) => drawMarkupItem(context, item, drawingSession));
    if (drawingSession.draftItem) {
      drawMarkupItem(context, drawingSession.draftItem, drawingSession);
    }
    if (drawingSession.selectedIndex >= 0 && drawingSession.items[drawingSession.selectedIndex]) {
      drawSelectionOverlay(context, drawingSession.items[drawingSession.selectedIndex], drawingSession);
    }
    drawSnapGuides(context);
    syncSelectedItemControls();
  }

  function drawMarkupItem(context, item, session) {
    if (!item) {
      return;
    }
    const scaleX = context.canvas.width / session.naturalWidth;
    const scaleY = context.canvas.height / session.naturalHeight;
    context.lineCap = "round";
    context.lineJoin = "round";

    if (!item.type || item.type === "pen") {
      drawPenPath(context, item, scaleX, scaleY);
      return;
    }

    if (item.type === "rectangle") {
      context.strokeStyle = item.color;
      context.lineWidth = item.size * scaleX;
      const width = (item.end.x - item.start.x) * scaleX;
      const height = (item.end.y - item.start.y) * scaleY;
      context.strokeRect(item.start.x * scaleX, item.start.y * scaleY, width, height);
      return;
    }

    if (item.type === "highlight") {
      const width = (item.end.x - item.start.x) * scaleX;
      const height = (item.end.y - item.start.y) * scaleY;
      context.fillStyle = hexToRgba(item.color, 0.28);
      context.fillRect(item.start.x * scaleX, item.start.y * scaleY, width, height);
      context.strokeStyle = item.color;
      context.lineWidth = Math.max(1, item.size * 0.4 * scaleX);
      context.strokeRect(item.start.x * scaleX, item.start.y * scaleY, width, height);
      return;
    }

    if (item.type === "arrow") {
      drawArrow(context, item, scaleX, scaleY);
      return;
    }

    if (item.type === "text") {
      drawWrappedText(context, item, scaleX, scaleY);
      return;
    }

    if (item.type === "callout") {
      drawCallout(context, item, scaleX, scaleY);
    }
  }

  function drawCallout(context, item, scaleX, scaleY) {
    const x = item.start.x * scaleX;
    const y = item.start.y * scaleY;
    const radius = Math.max(14, item.size * 3 * scaleX);
    context.fillStyle = item.color;
    context.beginPath();
    context.arc(x, y, radius, 0, Math.PI * 2);
    context.fill();
    context.fillStyle = "#ffffff";
    context.font = `bold ${Math.max(12, item.size * 3 * scaleX)}px Segoe UI`;
    context.textAlign = "center";
    context.textBaseline = "middle";
    context.fillText(String(item.number || "1"), x, y);
    context.textAlign = "start";
    context.textBaseline = "top";
    if (item.text) {
      context.fillStyle = item.color;
      context.font = `${Math.max(14, item.size * 3.6 * scaleX)}px Segoe UI`;
      drawMultilineText(context, item.text, x + radius + 10, y - radius / 2, Math.max(120, (item.width || 180) * scaleX), Math.max(16, item.size * 4 * scaleX));
    }
  }

  function drawWrappedText(context, item, scaleX, scaleY) {
    const fontSize = Math.max(14, item.size * 4 * scaleX);
    context.fillStyle = item.color;
    context.font = `${fontSize}px Segoe UI`;
    context.textBaseline = "top";
    drawMultilineText(context, item.text || "Text", item.start.x * scaleX, item.start.y * scaleY, Math.max(100, (item.width || 220) * scaleX), fontSize * 1.35);
  }

  function drawMultilineText(context, text, x, y, maxWidth, lineHeight) {
    wrapText(context, text, maxWidth).forEach((line, index) => {
      context.fillText(line, x, y + index * lineHeight);
    });
  }

  function drawSelectionOverlay(context, item, session) {
    const bounds = getItemBounds(item, session, context);
    if (!bounds) {
      return;
    }
    context.save();
    context.strokeStyle = "#084f47";
    context.setLineDash([8, 6]);
    context.lineWidth = 1.5;
    context.strokeRect(bounds.x, bounds.y, bounds.width, bounds.height);
    context.setLineDash([]);
    getHandleRects(bounds).forEach((handle) => {
      context.fillStyle = "#ffffff";
      context.strokeStyle = "#084f47";
      context.lineWidth = 1.5;
      context.fillRect(handle.x, handle.y, handle.size, handle.size);
      context.strokeRect(handle.x, handle.y, handle.size, handle.size);
    });
    context.restore();
  }

  function drawPenPath(context, item, scaleX, scaleY) {
    if (!item.points?.length) {
      return;
    }
    context.strokeStyle = item.color;
    context.lineWidth = item.size * scaleX;
    context.beginPath();
    item.points.forEach((point, index) => {
      const x = point.x * scaleX;
      const y = point.y * scaleY;
      if (index === 0) {
        context.moveTo(x, y);
      } else {
        context.lineTo(x, y);
      }
    });
    if (item.points.length === 1) {
      context.lineTo(item.points[0].x * scaleX + 0.01, item.points[0].y * scaleY + 0.01);
    }
    context.stroke();
  }

  function drawArrow(context, item, scaleX, scaleY) {
    const startX = item.start.x * scaleX;
    const startY = item.start.y * scaleY;
    const endX = item.end.x * scaleX;
    const endY = item.end.y * scaleY;
    const angle = Math.atan2(endY - startY, endX - startX);
    const headLength = Math.max(10, item.size * 3 * scaleX);

    context.strokeStyle = item.color;
    context.fillStyle = item.color;
    context.lineWidth = item.size * scaleX;
    context.beginPath();
    context.moveTo(startX, startY);
    context.lineTo(endX, endY);
    context.stroke();

    context.beginPath();
    context.moveTo(endX, endY);
    context.lineTo(endX - headLength * Math.cos(angle - Math.PI / 6), endY - headLength * Math.sin(angle - Math.PI / 6));
    context.lineTo(endX - headLength * Math.cos(angle + Math.PI / 6), endY - headLength * Math.sin(angle + Math.PI / 6));
    context.closePath();
    context.fill();
  }

  function handleDrawStart(event) {
    if (!drawingSession) {
      return;
    }
    event.preventDefault();
    const point = getNaturalPoint(event);
    const tool = elements.markupTool.value;

    if (tryStartSelectionInteraction(event, point)) {
      redrawDrawingCanvas();
      return;
    }

    if (tool === "text") {
      drawingSession.draftItem = {
        type: "text",
        color: elements.penColor.value,
        size: Number(elements.penSize.value),
        start: point,
        text: elements.textInput.value.trim() || "Text",
        width: 220
      };
      drawingSession.items.push(drawingSession.draftItem);
      drawingSession.selectedIndex = drawingSession.items.length - 1;
      drawingSession.draftItem = null;
      drawingSession.dirty = true;
      redrawDrawingCanvas();
      return;
    }

    if (tool === "callout") {
      drawingSession.items.push({
        type: "callout",
        color: elements.penColor.value,
        size: Number(elements.penSize.value),
        start: point,
        number: getNextCalloutNumber(),
        text: elements.textInput.value.trim(),
        width: 180
      });
      drawingSession.selectedIndex = drawingSession.items.length - 1;
      drawingSession.dirty = true;
      redrawDrawingCanvas();
      return;
    }

    drawingSession.draftItem = createDraftItem(tool, point);
    drawingSession.selectedIndex = -1;
    redrawDrawingCanvas();
  }

  function handleDrawMove(event) {
    if (!drawingSession) {
      return;
    }
    event.preventDefault();
    const point = getNaturalPoint(event);
    if (drawingSession.interactionMode === "move") {
      updateSelectedItemPosition(point);
      redrawDrawingCanvas();
      return;
    }
    if (drawingSession.interactionMode === "resize") {
      resizeSelectedItem(point);
      redrawDrawingCanvas();
      return;
    }
    if (!drawingSession.draftItem) {
      return;
    }
    updateDraftItem(drawingSession.draftItem, point);
    redrawDrawingCanvas();
  }

  function handleDrawEnd(event) {
    if (!drawingSession) {
      return;
    }
    event.preventDefault();
    releaseDrawingPointerCapture(event.pointerId);
    const point = getNaturalPoint(event);
    if (drawingSession.interactionMode === "move" || drawingSession.interactionMode === "resize") {
      if (drawingSession.interactionMode === "move") {
        updateSelectedItemPosition(point);
      } else {
        resizeSelectedItem(point);
      }
      drawingSession.interactionMode = null;
      drawingSession.activeHandle = null;
      drawingSession.dragOrigin = null;
      drawingSession.snapGuides = [];
      drawingSession.dirty = true;
      redrawDrawingCanvas();
      return;
    }
    if (!drawingSession.draftItem) {
      return;
    }
    updateDraftItem(drawingSession.draftItem, point);
    drawingSession.items.push(drawingSession.draftItem);
    drawingSession.selectedIndex = drawingSession.items.length - 1;
    drawingSession.draftItem = null;
    drawingSession.dirty = true;
    redrawDrawingCanvas();
  }

  function releaseDrawingPointerCapture(pointerId) {
    const canvas = elements.drawingCanvas;
    if (!canvas || typeof canvas.releasePointerCapture !== "function") {
      return;
    }
    if (typeof pointerId === "number") {
      try {
        if (!canvas.hasPointerCapture || canvas.hasPointerCapture(pointerId)) {
          canvas.releasePointerCapture(pointerId);
        }
      } catch {
        // Ignore pointer capture release issues from stale pointer ids.
      }
    }
    if (drawingSession) {
      drawingSession.activePointerId = null;
    }
  }

  function createDraftItem(tool, point) {
    const base = {
      type: tool,
      color: elements.penColor.value,
      size: Number(elements.penSize.value)
    };

    if (tool === "pen") {
      return { ...base, points: [point] };
    }

    return { ...base, start: point, end: point };
  }

  function updateDraftItem(item, point) {
    if (item.type === "pen") {
      item.points.push(point);
      return;
    }
    if (item.type === "text" || item.type === "callout") {
      item.start = point;
      return;
    }
    item.end = point;
  }

  function getNaturalPoint(event) {
    const rect = elements.drawingCanvas.getBoundingClientRect();
    const x = (event.clientX - rect.left) / rect.width;
    const y = (event.clientY - rect.top) / rect.height;
    return {
      x: Math.max(0, Math.min(drawingSession.naturalWidth, x * drawingSession.naturalWidth)),
      y: Math.max(0, Math.min(drawingSession.naturalHeight, y * drawingSession.naturalHeight))
    };
  }

  function undoStroke() {
    if (!drawingSession) {
      return;
    }
    if (drawingSession.draftItem) {
      drawingSession.draftItem = null;
      drawingSession.snapGuides = [];
      drawingSession.dirty = true;
      redrawDrawingCanvas();
      return;
    }
    if (!drawingSession.items.length) {
      if (elements.saveStatus) {
        elements.saveStatus.textContent = "Nothing to undo";
      }
      return;
    }
    drawingSession.items.pop();
    drawingSession.selectedIndex = Math.min(drawingSession.selectedIndex, drawingSession.items.length - 1);
    drawingSession.snapGuides = [];
    drawingSession.dirty = true;
    redrawDrawingCanvas();
  }

  function clearDrawing() {
    if (!drawingSession) {
      return;
    }
    drawingSession.items = [];
    drawingSession.draftItem = null;
    drawingSession.selectedIndex = -1;
    drawingSession.snapGuides = [];
    drawingSession.dirty = true;
    redrawDrawingCanvas();
  }

  async function saveDrawingMarkup() {
    await persistDrawingMarkup({ closeDialog: true });
  }

  async function persistDrawingMarkup({ closeDialog = false } = {}) {
    if (!drawingSession) {
      return;
    }
    const session = drawingSession;
    const shot = findShot(session.testId, session.shotId);
    if (!shot) {
      return;
    }
    shot.drawingItems = structuredClone(session.items);
    delete shot.drawingPaths;
    shot.naturalWidth = session.naturalWidth;
    shot.naturalHeight = session.naturalHeight;
    shot.flattenedDataUrl = shot.drawingItems.length ? await buildAnnotatedImage(shot) : "";
    drawingSession = null;
    render();
    saveState();
    if (closeDialog && elements.drawingDialog.open) {
      isClosingDrawingAfterSave = true;
      elements.drawingDialog.close();
      isClosingDrawingAfterSave = false;
    }
  }

  function findShot(testId, shotId) {
    return collectTestScreenshots(state.tests.find((item) => item.id === testId)).find((item) => item.id === shotId);
  }

  function buildAnnotatedImage(shot) {
    const canvas = document.createElement("canvas");
    const width = shot.naturalWidth || 1;
    const height = shot.naturalHeight || 1;
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    return getShotDataUrlForExport(shot).then((sourceDataUrl) => {
      if (!sourceDataUrl) {
        return "";
      }
      return renderImageToCanvas(sourceDataUrl, canvas, context).then(() => {
        (shot.drawingItems || shot.drawingPaths || []).forEach((item) => drawMarkupItem(context, item, { naturalWidth: width, naturalHeight: height, items: [] }));
        return canvas.toDataURL("image/png");
      });
    });
  }

  function exportRun() {
    void saveRunSnapshot("before-export-summary").catch(() => null);
    const portablePayloadPromise = buildPortableExportState();
    const counts = state.tests.reduce((accumulator, test) => {
      if (test.status === "not-start") accumulator.notStart += 1;
      if (test.status === "passed") accumulator.passed += 1;
      if (test.status === "query") accumulator.query += 1;
      if (test.status === "failed") accumulator.failed += 1;
      if (test.status === "blocked") accumulator.blocked += 1;
      if (test.status === "cancelled") accumulator.cancelled += 1;
      return accumulator;
    }, { notStart: 0, passed: 0, query: 0, failed: 0, blocked: 0, cancelled: 0 });

    const confetti = counts.failed === 0 ? "<div style=\"font-size:40px;margin-bottom:12px;\">🎉</div>" : "";
    Promise.all([portablePayloadPromise, ...state.tests.map(async (test, testIndex) => {
      const testSequenceNumber = testIndex + 1;
      const attachments = await Promise.all((test.screenshots || []).map(async (shot) => {
        const hasMarkup = Boolean((shot.drawingItems || shot.drawingPaths || []).length);
        const imageUrl = shot.flattenedDataUrl || (hasMarkup ? await buildAnnotatedImage(shot) : await getShotDataUrlForExport(shot));
        if (!imageUrl) {
          return `<div style="margin:0 0 14px;padding:12px;border:1px solid #d3c7af;border-radius:12px;background:#fbf8f1;color:#8a3026;">Screenshot data is not available in this browser.</div>`;
        }
        const reviewMarkup = shot.reviewComment ? `<div style="margin-top:8px;padding:10px 12px;border-left:4px solid #c47f00;background:#fff4d8;border-radius:8px;"><strong>Reviewer comment</strong><br>${escapeHtml(shot.reviewComment)}</div>` : "";
        return `
          <figure style="margin:0 0 14px;">
            <img src="${imageUrl}" alt="${escapeHtml(shot.name)}" style="max-width:100%;border-radius:12px;border:1px solid #d3c7af;">
            <figcaption style="margin-top:8px;"><strong>${escapeHtml(shot.name)}</strong><br><span class="editable-field" data-edit-id="shot-annotation-${escapeHtml(test.id)}-${escapeHtml(shot.id)}" contenteditable="true">${linkifyText(shot.annotation || "No annotation")}</span>${reviewMarkup}</figcaption>
          </figure>`;
      }));
      const videosMarkup = (await Promise.all((test.videos || []).map(async (video) => {
        const videoDataUrl = await getVideoDataUrlForExport(video);
        return `
          <div style="margin:0 0 14px;padding:12px;border:1px solid #d3c7af;border-radius:12px;background:#fbf8f1;">
            <strong>${escapeHtml(video.name)}</strong><br>
            <small>${escapeHtml(formatVideoMeta(video))}</small>
            ${videoDataUrl ? `<video controls style="width:100%;margin-top:10px;border-radius:10px;background:#111;" src="${videoDataUrl}"></video>` : `<div style="margin-top:10px;color:#8a3026;">Video data is not available in this browser.</div>`}
            ${video.notes ? `<div class="editable-field" data-edit-id="video-notes-${escapeHtml(test.id)}-${escapeHtml(video.id)}" contenteditable="true" style="margin-top:8px;">${getRichDisplayHtml(video.notesHtml, video.notes)}</div>` : ""}
          </div>`;
      }))).join("");
      const attachmentsMarkup = `${videosMarkup}${attachments.join("")}` || "<p>No attachments.</p>";
      const stepMarkup = test.steps.length
        ? (await Promise.all(test.steps.map(async (step) => {
          const stepVideosMarkup = (await Promise.all((step.videos || []).map(async (video) => {
            const videoDataUrl = await getVideoDataUrlForExport(video);
            return `
            <div style="margin:10px 0 0;padding:10px;border:1px solid #d3c7af;border-radius:10px;background:#fbf8f1;">
              <strong>${escapeHtml(video.name)}</strong><br>
              <small>${escapeHtml(formatVideoMeta(video))}</small>
              ${videoDataUrl ? `<video controls style="width:100%;margin-top:10px;border-radius:10px;background:#111;" src="${videoDataUrl}"></video>` : `<div style="margin-top:10px;color:#8a3026;">Video data is not available in this browser.</div>`}
              ${video.notes ? `<div class="editable-field" data-edit-id="video-notes-${escapeHtml(test.id)}-${escapeHtml(step.id)}-${escapeHtml(video.id)}" contenteditable="true" style="margin-top:8px;">${getRichDisplayHtml(video.notesHtml, video.notes)}</div>` : ""}
            </div>`;
          }))).join("");
          const stepShotsMarkup = await Promise.all((step.screenshots || []).map(async (shot) => {
            const hasMarkup = Boolean((shot.drawingItems || shot.drawingPaths || []).length);
            const imageUrl = shot.flattenedDataUrl || (hasMarkup ? await buildAnnotatedImage(shot) : await getShotDataUrlForExport(shot));
            if (!imageUrl) {
              return `<div style="margin:10px 0 0;padding:10px;border:1px solid #d3c7af;border-radius:10px;background:#fbf8f1;color:#8a3026;">Screenshot data is not available in this browser.</div>`;
            }
            const reviewMarkup = shot.reviewComment ? `<div style="margin-top:8px;padding:10px 12px;border-left:4px solid #c47f00;background:#fff4d8;border-radius:8px;"><strong>Reviewer comment</strong><br>${escapeHtml(shot.reviewComment)}</div>` : "";
            return `<figure style="margin:10px 0 0;"><img src="${imageUrl}" alt="${escapeHtml(shot.name)}" style="max-width:100%;border-radius:12px;border:1px solid #d3c7af;"><figcaption style="margin-top:8px;"><strong>${escapeHtml(shot.name)}</strong><br><span class="editable-field" data-edit-id="shot-annotation-${escapeHtml(test.id)}-${escapeHtml(step.id)}-${escapeHtml(shot.id)}" contenteditable="true">${linkifyText(shot.annotation || "No annotation")}</span>${reviewMarkup}</figcaption></figure>`;
          }));
          return `<li><div class="editable-field export-step-field" data-edit-id="step-${escapeHtml(test.id)}-${escapeHtml(step.id)}" contenteditable="true">${getRichDisplayHtml(step.textHtml, step.text, "(No step text)")}</div>${stepShotsMarkup.join("")}${stepVideosMarkup}</li>`;
        }))).join("")
        : "<li>No steps recorded</li>";
      const filesMarkup = (await Promise.all((test.files || []).map(async (fileAttachment) => {
        const fileDataUrl = await getAttachmentDataUrlForExport(fileAttachment);
        if (!fileDataUrl) {
          return `<div style="display:flex;align-items:center;gap:10px;padding:8px 12px;margin-bottom:6px;border:1px solid #d3c7af;border-radius:10px;background:#fbf8f1;color:#8a3026;"><span style="font-size:1.3rem;">${escapeHtml(fileTypeIcon(fileAttachment.type))}</span><span>${escapeHtml(fileAttachment.name || "Attachment")}</span><small style="color:#8a826f;">Unavailable in this browser</small></div>`;
        }
        return `<div style="display:flex;align-items:center;gap:10px;padding:8px 12px;margin-bottom:6px;border:1px solid #d3c7af;border-radius:10px;background:#fbf8f1;"><span style="font-size:1.3rem;">${escapeHtml(fileTypeIcon(fileAttachment.type))}</span><a href="${fileDataUrl}" download="${escapeHtml(fileAttachment.name)}" style="color:#0c6c61;font-weight:600;">${escapeHtml(fileAttachment.name)}</a><small style="color:#8a826f;">${escapeHtml(formatBytes(fileAttachment.sizeBytes))}</small></div>`;
      }))).join("");

      return `
      <section class="export-test-section" id="test-section-${escapeHtml(test.id)}" data-test-id="${escapeHtml(test.id)}" data-test-seq="${testSequenceNumber}" style="border:1px solid #d3c7af;border-radius:16px;padding:16px;margin-bottom:16px;page-break-inside:avoid;">
        <h2 style="margin:0 0 8px;"><span style="margin-right:8px;color:#6d6a63;">Test ${testSequenceNumber}.</span><span class="editable-field export-scenario-field" data-edit-id="test-title-${escapeHtml(test.id)}" contenteditable="true">${getRichDisplayHtml(test.titleHtml, test.title, "Untitled test")}</span></h2>
        <p><strong>Status:</strong> <span class="export-test-status" data-test-id="${escapeHtml(test.id)}" data-status-value="${escapeHtml(formatTestStatus(test.status))}">${escapeHtml(formatTestStatus(test.status))}</span></p>
        <p><strong>Review Outcome:</strong> ${escapeHtml(formatReviewStatus(test.reviewStatus || "not-reviewed"))}</p>
        <p><strong>UAT Results:</strong> <span class="editable-field export-uat-results-field" data-edit-id="test-notes-${escapeHtml(test.id)}" contenteditable="true">${getRichDisplayHtml(test.notesHtml, test.notes, "-")}</span></p>
        ${test.reviewComment ? `<div style="margin:10px 0 14px;padding:10px 12px;border-left:4px solid #c47f00;background:#fff4d8;border-radius:8px;"><strong>Reviewer comment</strong><br>${linkifyText(test.reviewComment)}</div>` : ""}
        <div class="reviewer-decision-block" data-test-id="${escapeHtml(test.id)}" style="margin:10px 0 14px;padding:10px 14px;border:3px solid #d3c7af;background:#fffdf8;border-radius:10px;">
          <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
            <strong>Reviewer Comment</strong>
            <select class="reviewer-export-status" data-test-id="${escapeHtml(test.id)}" style="border:1px solid #aaa;border-radius:8px;padding:8px 10px;font-weight:700;min-width:110px;">
              <option value="" selected>-- Select --</option>
              <option value="approved">Approved</option>
              <option value="query">Query</option>
              <option value="reject">Reject</option>
            </select>
          </div>
          <textarea class="reviewer-export-comment" data-test-id="${escapeHtml(test.id)}" placeholder="Add reviewer details for Query or Reject." style="width:100%;margin-top:8px;min-height:90px;border:1px solid #d3c7af;border-radius:8px;padding:10px;display:none;box-sizing:border-box;"></textarea>
          <small class="reviewer-export-note" style="display:block;margin-top:6px;color:#6d6a63;">Auto-saved in this browser for this exported report.</small>
        </div>
        <h3>Steps</h3>
        <ol>${stepMarkup}</ol>
        <h3>Attachments</h3>
        ${attachmentsMarkup}
        ${(test.files || []).length ? `<h3>Files</h3><div>${filesMarkup}</div>` : ""}
      </section>`;
    })]).then(([portablePayload, ...sections]) => {
      const embeddedRunJson = serializePortablePayloadForHtml(portablePayload);
      const testMarkup = sections.join("");
      const floatingNavStyles = `:root{--export-floating-nav-reserved-left:404px;} body.has-export-floating-nav{max-width:none !important;margin:24px 18px 40px var(--export-floating-nav-reserved-left) !important;} .export-floating-nav{position:fixed;left:18px;top:18px;width:min(356px,calc(100vw - 36px));max-height:calc(100vh - 36px);padding:14px;border-radius:16px;border:1px solid rgba(255,255,255,0.7);background:#fffdf8;box-shadow:0 10px 28px rgba(31,36,31,.12);backdrop-filter:blur(18px);z-index:1200;display:flex;flex-direction:column;gap:10px;overflow:hidden;} .export-floating-nav h3{margin:0;font-size:0.98rem;} .export-floating-nav-header{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-shrink:0;} .export-floating-nav-body{flex:1;min-height:0;overflow-y:auto;padding-right:6px;} .export-floating-nav-table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #d3c7af;border-radius:10px;overflow:hidden;font-size:0.9rem;} .export-floating-nav-table thead{background:#f6f1e7;} .export-floating-nav-table th{padding:8px 10px;border-bottom:1px solid #d3c7af;text-align:left;font-weight:700;color:#1f241f;} .export-floating-nav-table th:first-child{width:50px;text-align:center;} .export-floating-nav-table th:last-child{width:100px;text-align:center;} .export-floating-nav-table td{padding:8px 10px;border-bottom:1px solid #eee3cf;vertical-align:top;} .export-floating-nav-table td:first-child{text-align:center;font-weight:600;width:50px;} .export-floating-nav-table td:last-child{text-align:center;width:100px;} .export-floating-nav-table a{color:#0c6c61;font-weight:700;text-decoration:none;} .export-floating-nav-table a:hover{text-decoration:underline;} .export-floating-nav-status{font-size:0.8rem;font-weight:700;border-radius:4px;padding:4px 6px;display:inline-block;} .export-floating-nav-status.status-passed{background:#e7f3ef;color:#0b5a4c;} .export-floating-nav-status.status-query{background:#fff1d6;color:#9a6200;} .export-floating-nav-status.status-failed{background:#fbe7e4;color:#8a3026;} .export-floating-nav-status.status-blocked,.export-floating-nav-status.status-cancelled{background:#ece8dd;color:#6d6a63;} @media (max-width:1560px){body.has-export-floating-nav{max-width:980px !important;margin:24px auto !important;padding:0 16px !important;} .export-floating-nav{position:static;left:auto;top:auto;width:auto;max-height:none;margin:0 0 16px 0;}} @media print{.export-floating-nav{display:none;} body.has-export-floating-nav{margin:24px auto !important;}}`;
      const exportStyles = state.presentationMode
        ? floatingNavStyles + "body{font-family:Segoe UI,Tahoma,sans-serif;max-width:1080px;margin:28px auto;padding:0 24px;color:#1f241f;line-height:1.65;} h1{font-size:2.5rem;margin-bottom:8px;} h2{font-size:1.6rem;margin-top:28px;} h3{font-size:1.15rem;margin-top:22px;} section{box-shadow:0 8px 24px rgba(69,52,31,.08);background:#fffdf8;} figure{padding-bottom:10px;border-bottom:1px solid #e2d7c0;} figcaption{font-size:1rem;line-height:1.55;} a{cursor:pointer;color:#0c6c61;} .editable-field{display:inline-block;min-width:12px;padding:2px 4px;border-radius:6px;outline:1px dashed rgba(12,108,97,.35);outline-offset:2px;background:rgba(255,255,255,.55);} .export-scenario-field{display:inline-block;padding:8px 10px;background:#e8f3ff;border:1px solid #9cc4f2;border-radius:8px;} .export-uat-results-field{padding:4px 8px;background:#eef7f5;border:1px solid #9bcab9;border-radius:8px;} .export-step-field{display:inline-block;padding:6px 10px;background:#f1ffd6;border:1px solid #b7e27a;border-radius:8px;}"
        : floatingNavStyles + "body{font-family:Segoe UI,Tahoma,sans-serif;max-width:980px;margin:24px auto;padding:0 16px;color:#1f241f;} a{cursor:pointer;color:#0c6c61;} .editable-field{display:inline-block;min-width:12px;padding:2px 4px;border-radius:6px;outline:1px dashed rgba(12,108,97,.35);outline-offset:2px;background:rgba(255,255,255,.55);} .export-scenario-field{display:inline-block;padding:8px 10px;background:#e8f3ff;border:1px solid #9cc4f2;border-radius:8px;} .export-uat-results-field{padding:4px 8px;background:#eef7f5;border:1px solid #9bcab9;border-radius:8px;} .export-step-field{display:inline-block;padding:6px 10px;background:#f1ffd6;border:1px solid #b7e27a;border-radius:8px;}";
      const overviewMarkup = `
        <h3>Consolidated Test Summary</h3>
        <div style="overflow-x:auto;margin-top:10px;">
          <table style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #e2d7c0;border-radius:12px;overflow:hidden;">
            <thead>
              <tr style="background:#f6f1e7;text-align:left;">
                <th style="padding:10px 12px;border-bottom:1px solid #e2d7c0;width:60px;">No.</th>
                <th style="padding:10px 12px;border-bottom:1px solid #e2d7c0;">Test Scenario</th>
                <th style="padding:10px 12px;border-bottom:1px solid #e2d7c0;width:120px;">Status</th>
                <th style="padding:10px 12px;border-bottom:1px solid #e2d7c0;">UAT Results</th>
                <th style="padding:10px 12px;border-bottom:1px solid #e2d7c0;">Reviewer Comment</th>
              </tr>
            </thead>
            <tbody id="consolidatedSummaryList">${state.tests.map((test, testIndex) => `<tr>
              <td style="padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;text-align:center;"><a href="#test-section-${escapeHtml(test.id)}" style="color:#0c6c61;font-weight:700;text-decoration:none;">${testIndex + 1}</a></td>
              <td style="padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;font-weight:600;">${getRichDisplayHtml(test.titleHtml, test.title, "Untitled test")}</td>
              <td style="${statusCellStyle(test.status)}">${escapeHtml(formatTestStatus(test.status))}</td>
              <td style="padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;white-space:pre-wrap;">${getRichDisplayHtml(test.notesHtml, test.notes, "-")}</td>
              <td style="padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;white-space:pre-wrap;">${linkifyText(test.reviewComment || "-")}</td>
            </tr>`).join("")}</tbody>
          </table>
        </div>`;

      const floatingNavMarkup = `<aside class="export-floating-nav">
        <div class="export-floating-nav-header">
          <h3>Consolidated Test Summary</h3>
        </div>
        <div class="export-floating-nav-body">
          <table class="export-floating-nav-table">
            <thead>
              <tr>
                <th>NO</th>
                <th>Test Scenario</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>${state.tests.map((test, testIndex) => {
              const statusClass = `status-${normalizeTestStatus(test.status)}`;
              return `<tr>
                <td><a href="#test-section-${escapeHtml(test.id)}">${testIndex + 1}</a></td>
                <td><a href="#test-section-${escapeHtml(test.id)}">${getRichDisplayHtml(test.titleHtml, test.title, "Untitled test")}</a></td>
                <td><span class="export-floating-nav-status ${statusClass}">${escapeHtml(formatTestStatus(test.status))}</span></td>
              </tr>`;
            }).join("")}</tbody>
          </table>
        </div>
      </aside>`;

      const html = `<!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="utf-8">
        <title>${escapeHtml(state.runName || "Test run export")}</title>
        <style>${exportStyles}</style>
      </head>
      <body class="has-export-floating-nav">
        ${floatingNavMarkup}
        ${confetti}
        <h1 class="editable-field" data-edit-id="run-name" contenteditable="true">${escapeHtml(state.runName || "Test Run Summary")}</h1>
        <p><strong>Owner:</strong> <span class="editable-field" data-edit-id="owner-name" contenteditable="true">${escapeHtml(state.ownerName || "-")}</span></p>
        <p><strong>Not Start:</strong> ${counts.notStart} | <strong>Passed:</strong> ${counts.passed} | <strong>Query:</strong> ${counts.query} | <strong>Failed:</strong> ${counts.failed} | <strong>Blocked:</strong> ${counts.blocked} | <strong>Cancelled:</strong> ${counts.cancelled} | <strong>Queries Raised:</strong> ${Number(state.queryCount || 0)}</p>
        <h2>Version Details</h2>
        <pre class="editable-field" data-edit-id="version-details" contenteditable="true" style="white-space:pre-wrap;background:#f6f1e7;border-radius:12px;padding:14px;border:1px solid #d3c7af;">${escapeHtml(state.versionDetails || "-")}</pre>
        <section style="margin:24px 0 24px 0;padding:18px 14px;border:1px solid #d3c7af;border-radius:14px;background:#f8f7f3;">
          <h2 style="margin-top:0;">SetUp</h2>
          <div style="margin-bottom:10px;"><strong>Description / Steps:</strong><br><div style="background:#fff;border-radius:8px;border:1px solid #e2d7c0;padding:10px 12px;min-height:40px;">${state.setup?.textHtml || "<span style='color:#aaa'>(No setup description)</span>"}</div></div>
          <div style="margin-bottom:10px;"><strong>Screenshots:</strong><br><div style="display:flex;gap:8px;flex-wrap:wrap;">${(state.setup?.screenshots||[]).map(shot => `<img src="${shot.dataUrl}" alt="Setup Screenshot" style="width:80px;height:80px;object-fit:cover;border:1px solid #d3c7af;border-radius:8px;background:#fff;">`).join("") || "<span style='color:#aaa'>(No screenshots)</span>"}</div></div>
          <div><strong>Related Links:</strong><br>${(state.setup?.links||"").split(/[,;\n]+/).filter(Boolean).map(link => `<a href="${link.trim()}" target="_blank" rel="noopener">${link.trim()}</a>`).join("<br>") || "<span style='color:#aaa'>(No links)</span>"}</div>
        </section>
        ${overviewMarkup}
        <div style="display:flex;justify-content:flex-end;margin:0 0 14px;">
          <button id="downloadReviewedCopyButton" type="button" style="border:0;border-radius:999px;padding:10px 16px;background:#0c6c61;color:#fff;cursor:pointer;">Download Reviewed Copy</button>
        </div>
        ${testMarkup}
        <div style="margin:28px 0 18px;text-align:center;font-size:1.05rem;font-weight:700;letter-spacing:0.08em;color:#8a826f;">END</div>
        <button id="backToTopButton" type="button" aria-label="Back to top" style="position:fixed;right:18px;bottom:18px;width:44px;height:44px;border:0;border-radius:50%;background:#0c6c61;color:#fff;font-size:20px;line-height:1;cursor:pointer;box-shadow:0 10px 18px rgba(12,108,97,.35);">↑</button>
        <script id="embeddedRunJson" type="application/json">${embeddedRunJson}</script>
        <script>
          window.addEventListener("load", function() {
            try {
            var runSlug = "${slugify(state.runName || "test-run")}";
            var storageKey = "export-review-comments-" + runSlug;
            var editStorageKey = "export-editable-fields-" + runSlug;
            var comments = {};
            var editableContent = {};
            try {
              comments = JSON.parse(localStorage.getItem(storageKey) || "{}");
            } catch(e) {
              comments = {};
            }
            try {
              editableContent = JSON.parse(localStorage.getItem(editStorageKey) || "{}");
            } catch(e) {
              editableContent = {};
            }

            function saveComments() {
              try { localStorage.setItem(storageKey, JSON.stringify(comments)); } catch(e) {}
            }

            function saveEditableFields() {
              try { localStorage.setItem(editStorageKey, JSON.stringify(editableContent)); } catch(e) {}
            }

            function shouldShowComment(status) {
              return status === "query" || status === "reject";
            }

            function applyDecisionStyle(testId, status) {
              var block = document.querySelector('.reviewer-decision-block[data-test-id="' + testId + '"]');
              var textarea = document.querySelector('.reviewer-export-comment[data-test-id="' + testId + '"]');
              if (!block) return;
              if (status === "") {
                block.style.borderColor = "#ffb300";
                block.style.background = "#fff3bf";
              } else if (status === "reject") {
                block.style.borderColor = "#b3261e";
                block.style.background = "#fdeceb";
              } else if (status === "query") {
                block.style.borderColor = "#c47f00";
                block.style.background = "#fff6df";
              } else {
                block.style.borderColor = "#2e7d32";
                block.style.background = "#eaf7ed";
              }
              if (textarea) {
                textarea.style.borderColor = block.style.borderColor;
              }
            }

            function normalizeSavedEntry(value) {
              if (typeof value === "string") {
                return { status: value.trim() ? "query" : "", comment: value };
              }
              if (!value || typeof value !== "object") {
                return { status: "", comment: "" };
              }
              var status = value.status === "query" || value.status === "reject" || value.status === "approved" || value.status === ""
                ? value.status
                : "";
              return { status: status, comment: typeof value.comment === "string" ? value.comment : "" };
            }

            function persistEntry(testId, status, comment) {
              comments[testId] = { status: status, comment: comment };
              saveComments();
            }

            function formatReviewerSummary(status, comment) {
              var trimmedComment = typeof comment === "string" ? comment.trim() : "";
              if (status === "approved") {
                return trimmedComment ? "Approved: " + trimmedComment : "Approved";
              }
              if (status === "query") {
                return trimmedComment ? "Query: " + trimmedComment : "Query";
              }
              if (status === "reject") {
                return trimmedComment ? "Reject: " + trimmedComment : "Reject";
              }
              return trimmedComment || "Not Set";
            }

            function collectSummaryFromPage() {
              return Array.from(document.querySelectorAll(".export-test-section")).map(function(section) {
                var testId = section.getAttribute("data-test-id") || "";
                var testSequenceNumber = section.getAttribute("data-test-seq") || "";
                var scenario = section.querySelector('.export-scenario-field');
                var status = section.querySelector('.export-test-status');
                var uat = section.querySelector('.export-uat-results-field');
                var reviewComment = section.querySelector('.reviewer-export-comment');
                var reviewStatus = section.querySelector('.reviewer-export-status');
                var reviewerStatusValue = reviewStatus ? reviewStatus.value : "";
                var reviewerCommentValue = reviewComment ? reviewComment.value : "";
                return {
                  testId: testId,
                  testSequenceNumber: testSequenceNumber,
                  scenario: scenario ? scenario.textContent.trim() : "Untitled test",
                  status: status ? status.textContent.trim() : "Not Start",
                  uatResults: uat ? uat.textContent.trim() : "-",
                  reviewComment: formatReviewerSummary(reviewerStatusValue, reviewerCommentValue)
                };
              });
            }

            function refreshConsolidatedSummary() {
              var summary = collectSummaryFromPage();
              var listEl = document.getElementById("consolidatedSummaryList");
              if (listEl) {
                listEl.innerHTML = summary.map(function(item) {
                  var rawStatus = (item.statusRaw || item.status || '').toLowerCase().replace(/ /g, '-');
                  var statusStyle = statusCellStyleInline(rawStatus);
                  return '<tr>'
                    + '<td style="padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;text-align:center;"><a href="#test-section-' + escapeHtmlInline(item.testId) + '" style="color:#0c6c61;font-weight:700;text-decoration:none;">' + escapeHtmlInline(item.testSequenceNumber || '') + '</a></td>'
                    + '<td style="padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;font-weight:600;">' + escapeHtmlInline(item.scenario || 'Untitled test') + '</td>'
                    + '<td style="' + statusStyle + '">' + escapeHtmlInline(item.status || 'Not Start') + '</td>'
                    + '<td style="padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;white-space:pre-wrap;">' + escapeHtmlInline(item.uatResults || '-') + '</td>'
                    + '<td style="padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;white-space:pre-wrap;">' + escapeHtmlInline(item.reviewComment || '-') + '</td>'
                    + '</tr>';
                }).join('');
              }
            }
            function escapeHtmlInline(s) {
                          function statusCellStyleInline(status) {
                            var base = 'padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;font-weight:700;border-radius:4px;';
                            var map = { 'passed':'background:#e7f3ef;color:#0b5a4c;', 'query':'background:#fff1d6;color:#9a6200;', 'failed':'background:#fbe7e4;color:#8a3026;', 'blocked':'background:#ece8dd;color:#6d6a63;', 'cancelled':'background:#ece8dd;color:#6d6a63;' };
                            return base + (map[status] || '');
                          }
              return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;");
            }
            function decisionLabel(status) {
              if (status === "") return "Not Set";
              if (status === "query") return "Query";
              if (status === "reject") return "Reject";
              return "Approved";
            }

            document.querySelectorAll(".editable-field[data-edit-id]").forEach(function(node) {
              var fieldId = node.getAttribute("data-edit-id");
              if (fieldId && typeof editableContent[fieldId] === "string") {
                node.innerHTML = editableContent[fieldId];
              }
              node.addEventListener("input", function() {
                if (!fieldId) return;
                editableContent[fieldId] = node.innerHTML;
                saveEditableFields();
              });
            });

            document.querySelectorAll(".reviewer-export-comment").forEach(function(textarea) {
              var testId = textarea.dataset.testId;
              var select = document.querySelector('.reviewer-export-status[data-test-id="' + testId + '"]');
              var saved = normalizeSavedEntry(comments[testId]);
              if (select) { select.value = saved.status; }
              textarea.value = saved.comment;
              textarea.style.display = shouldShowComment(saved.status) ? "block" : "none";
              applyDecisionStyle(testId, saved.status);

              if (select) {
                select.addEventListener("change", function() {
                  var status = select.value;
                  textarea.style.display = shouldShowComment(status) ? "block" : "none";
                  if (!shouldShowComment(status)) { textarea.value = ""; }
                  applyDecisionStyle(testId, status);
                  persistEntry(testId, status, textarea.value);
                });
              }

              textarea.addEventListener("input", function() {
                var status = select ? select.value : "";
                persistEntry(testId, status, textarea.value);
              });
            });

            // Allow clicking hyperlinks inside contenteditable fields
            document.addEventListener("click", function(e) {
              var anchor = e.target.closest("a[href]");
              if (anchor && anchor.closest(".editable-field")) {
                e.preventDefault();
                window.open(anchor.getAttribute("href"), "_blank", "noopener,noreferrer");
              }
            });

            var topButton = document.getElementById("backToTopButton");
            if (topButton) {
              topButton.addEventListener("click", function() {
                window.scrollTo({ top: 0, behavior: "smooth" });
              });
            }

            var downloadButton = document.getElementById("downloadReviewedCopyButton");
            if (downloadButton) {
              downloadButton.addEventListener("click", function() {
                try {
                  refreshConsolidatedSummary();
                  var clonedRoot = document.documentElement.cloneNode(true);
                  clonedRoot.querySelectorAll("script").forEach(function(node) { node.remove(); });
                  clonedRoot.querySelectorAll("#downloadReviewedCopyButton").forEach(function(node) { node.remove(); });
                  clonedRoot.querySelectorAll("#backToTopButton").forEach(function(node) { node.remove(); });
                  clonedRoot.querySelectorAll("#updateSummaryButton").forEach(function(node) { node.remove(); });
                  clonedRoot.querySelectorAll(".editable-field").forEach(function(node) {
                    node.removeAttribute("contenteditable");
                    node.style.outline = "none";
                    node.style.background = "transparent";
                    node.style.padding = "0";
                  });

                  clonedRoot.querySelectorAll(".reviewer-export-status").forEach(function(select) {
                    var badge = select.ownerDocument.createElement("span");
                    badge.style.display = "inline-block";
                    badge.style.padding = "7px 10px";
                    badge.style.borderRadius = "999px";
                    badge.style.fontWeight = "600";
                    var status = select.value;
                    if (status === "reject") {
                      badge.style.background = "#fdeceb";
                      badge.style.color = "#7a1f1a";
                      badge.style.border = "1px solid #b3261e";
                    } else if (status === "query") {
                      badge.style.background = "#fff6df";
                      badge.style.color = "#6f4b00";
                      badge.style.border = "1px solid #c47f00";
                    } else if (status === "") {
                      badge.style.background = "#fff3bf";
                      badge.style.color = "#7a5a00";
                      badge.style.border = "1px solid #ffb300";
                    } else {
                      badge.style.background = "#eaf7ed";
                      badge.style.color = "#245e27";
                      badge.style.border = "1px solid #2e7d32";
                    }
                    badge.textContent = decisionLabel(status);
                    select.replaceWith(badge);
                  });

                  clonedRoot.querySelectorAll(".reviewer-export-comment").forEach(function(textarea) {
                    var value = textarea.value.trim();
                    var block = textarea.ownerDocument.createElement("div");
                    block.style.marginTop = "8px";
                    block.style.padding = "10px";
                    block.style.border = "1px solid #d3c7af";
                    block.style.borderRadius = "8px";
                    block.style.background = "#fff";
                    block.innerHTML = value ? escapeHtmlInline(value).split(String.fromCharCode(10)).join("<br>") : "<em>No reviewer comment</em>";
                    textarea.replaceWith(block);
                  });

                  clonedRoot.querySelectorAll(".reviewer-export-note").forEach(function(node) {
                    node.textContent = "Snapshot of reviewer decisions and comments.";
                  });

                  var reviewedHtml = "<!DOCTYPE html>" + String.fromCharCode(10) + clonedRoot.outerHTML;
                  var blob = new Blob([reviewedHtml], { type: "text/html" });
                  var filename = runSlug + "-summary-reviewed.html";
                  var url = URL.createObjectURL(blob);
                  var anchor = document.createElement("a");
                  anchor.href = url;
                  anchor.download = filename;
                  anchor.click();
                  window.setTimeout(function() {
                    URL.revokeObjectURL(url);
                  }, 500);
                  downloadButton.textContent = "Downloaded";
                  window.setTimeout(function() {
                    downloadButton.textContent = "Download Reviewed Copy";
                  }, 1200);
                } catch (error) {
                  console.error(error);
                  window.alert("Unable to prepare reviewed copy in this browser context.");
                }
              });
            }

            refreshConsolidatedSummary();
            } catch(e) { console.error("Reviewer script error:", e); }
          });
        </script>
      </body>
      </html>`;

      const blob = new Blob([html], { type: "text/html" });
      const filename = `${slugify(state.runName || "test-run")}-summary.html`;
      saveFileWithBrowserDialog(blob, filename);
    });
  }

  function readFileAsDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  function readBlobAsDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  let xlsxLibraryPromise = null;

  function normalizeSpreadsheetHeader(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function parseCsvLine(line) {
    const cells = [];
    let current = "";
    let inQuotes = false;
    for (let i = 0; i < line.length; i += 1) {
      const char = line[i];
      const next = line[i + 1];
      if (char === '"') {
        if (inQuotes && next === '"') {
          current += '"';
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
        continue;
      }
      if (char === "," && !inQuotes) {
        cells.push(current.trim());
        current = "";
        continue;
      }
      current += char;
    }
    cells.push(current.trim());
    return cells;
  }

  function parseCsvToRows(text) {
    return String(text || "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => line.length)
      .map(parseCsvLine);
  }

  function looksLikeHeaderRow(row) {
    if (!Array.isArray(row)) return false;
    return row.some((cell) => /(test|scenario|title|step|status|uat|result|id|no|number)/i.test(String(cell || "")));
  }

  async function loadXlsxLibrary() {
    if (window.XLSX) {
      return window.XLSX;
    }
    if (xlsxLibraryPromise) {
      return xlsxLibraryPromise;
    }
    xlsxLibraryPromise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
      script.async = true;
      script.onload = () => {
        if (window.XLSX) {
          resolve(window.XLSX);
        } else {
          reject(new Error("Sheet parser did not load."));
        }
      };
      script.onerror = () => reject(new Error("Unable to load spreadsheet parser from CDN."));
      document.head.appendChild(script);
    });
    return xlsxLibraryPromise;
  }

  function splitStepText(raw) {
    const text = String(raw || "").trim();
    if (!text) return [];
    const newlineParts = text.split(/\r?\n+/).map((item) => item.trim()).filter(Boolean);
    if (newlineParts.length > 1) {
      return newlineParts;
    }
    return text
      .split(/\s*;\s*/)
      .map((item) => item.trim())
      .filter(Boolean);
  }

  function normalizeImportedStatus(value) {
    const raw = String(value || "").trim().toLowerCase();
    if (!raw) return "not-start";
    if (raw === "passed" || raw === "pass") return "passed";
    if (raw === "failed" || raw === "fail") return "failed";
    if (raw === "query") return "query";
    if (raw === "blocked") return "blocked";
    if (raw === "cancelled" || raw === "canceled") return "cancelled";
    if (raw === "not-start" || raw === "not started" || raw === "not-started") return "not-start";
    return normalizeTestStatus(raw);
  }

  function getSpreadsheetCellValue(rowObj, keys) {
    for (const key of keys) {
      if (Object.prototype.hasOwnProperty.call(rowObj, key) && String(rowObj[key] || "").trim()) {
        return String(rowObj[key] || "").trim();
      }
    }
    return "";
  }

  function buildTestsFromSpreadsheetRows(rows) {
    if (!rows.length) {
      return [];
    }

    const headerRow = looksLikeHeaderRow(rows[0]) ? rows[0] : [];
    const headers = (headerRow.length ? headerRow : rows[0]).map((item, index) => {
      const normalized = normalizeSpreadsheetHeader(item);
      return normalized || `column ${index + 1}`;
    });
    const startIndex = headerRow.length ? 1 : 0;

    const testMap = new Map();

    for (let rowIndex = startIndex; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex] || [];
      if (!row.some((cell) => String(cell || "").trim())) {
        continue;
      }
      const rowObj = {};
      headers.forEach((header, index) => {
        rowObj[header] = row[index] ?? "";
      });

      const title = getSpreadsheetCellValue(rowObj, [
        "test scenario",
        "scenario",
        "test title",
        "title",
        "test case",
        "testcase",
        "name",
        "description",
        "column 1"
      ]);
      if (!title) {
        continue;
      }

      const scenarioKey = getSpreadsheetCellValue(rowObj, ["test id", "scenario id", "id", "test no", "no", "number"])
        || title.toLowerCase();
      const status = getSpreadsheetCellValue(rowObj, ["status", "result", "outcome"]);
      const notes = getSpreadsheetCellValue(rowObj, ["uat results", "uat", "results", "notes", "comment", "remarks", "expected result"]);
      const stepText = getSpreadsheetCellValue(rowObj, ["step", "steps", "test step", "action", "procedure"]);

      if (!testMap.has(scenarioKey)) {
        const test = createTest();
        test.title = title;
        test.titleHtml = plainTextToRichHtml(title);
        test.status = normalizeImportedStatus(status);
        test.notes = notes;
        test.notesHtml = notes ? plainTextToRichHtml(notes) : "";
        test.steps = [];
        testMap.set(scenarioKey, test);
      }

      const test = testMap.get(scenarioKey);
      if (notes && !test.notes) {
        test.notes = notes;
        test.notesHtml = plainTextToRichHtml(notes);
      }
      if (status && test.status === "not-start") {
        test.status = normalizeImportedStatus(status);
      }

      const stepParts = splitStepText(stepText);
      stepParts.forEach((stepPart) => {
        if (test.steps.some((step) => step.text === stepPart)) {
          return;
        }
        const step = createStep();
        step.text = stepPart;
        step.textHtml = plainTextToRichHtml(stepPart);
        test.steps.push(step);
      });
    }

    return Array.from(testMap.values());
  }

  async function readSpreadsheetRows(file) {
    const fileName = (file?.name || "").toLowerCase();
    if (fileName.endsWith(".csv") || file.type === "text/csv") {
      return parseCsvToRows(await file.text());
    }

    const XLSX = await loadXlsxLibrary();
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    if (!workbook.SheetNames.length) {
      return [];
    }
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "", raw: false });
  }

  async function importSpreadsheet(file) {
    try {
      const rows = await readSpreadsheetRows(file);
      const importedTests = buildTestsFromSpreadsheetRows(rows);
      if (!importedTests.length) {
        window.alert("No test scenarios were found. Use columns like Test Scenario, Step, Status, and UAT Results.");
        return;
      }

      const replaceExisting = window.confirm("Replace all current tests with imported scenarios? Click Cancel to append instead.");
      if (replaceExisting) {
        await deleteVideosForTests(state.tests || []);
        await deleteBinaryMediaForTests(state.tests || []);
        state.tests = [];
      }
      state.tests.push(...importedTests);
      render();
      saveState();
      window.alert(`Imported ${importedTests.length} test scenario(s) from ${file.name}.`);
    } catch (error) {
      console.error(error);
      window.alert("Unable to import this spreadsheet. For best results, use .xlsx or .csv with a header row.");
    }
  }

  function buildSpreadsheetTemplateRows() {
    return [
      ["Test ID", "Test Scenario", "Step", "Status", "UAT Results"],
      ["TS-001", "Login with valid user", "Open login page", "Passed", "User lands on dashboard"],
      ["TS-001", "Login with valid user", "Enter username and password", "Passed", ""],
      ["TS-001", "Login with valid user", "Click Sign In", "Passed", ""],
      ["TS-002", "Login with invalid password", "Open login page", "Failed", "Error message is shown"],
      ["TS-002", "Login with invalid password", "Enter username and wrong password", "Failed", ""],
      ["TS-002", "Login with invalid password", "Click Sign In", "Failed", ""],
      ["TS-003", "Search with special characters", "Enter query: #@$%^", "Query", "Need confirmation on expected filtering"],
      ["TS-003", "Search with special characters", "Press Enter", "Query", ""]
    ];
  }

  function rowsToCsv(rows) {
    return rows
      .map((row) => row.map((cell) => {
        const value = String(cell ?? "");
        if (/[",\n\r]/.test(value)) {
          return `"${value.replace(/"/g, '""')}"`;
        }
        return value;
      }).join(","))
      .join("\r\n");
  }

  async function downloadSpreadsheetTemplate() {
    const rows = buildSpreadsheetTemplateRows();
    const datePart = new Date().toISOString().slice(0, 10);
    try {
      const XLSX = await loadXlsxLibrary();
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(workbook, worksheet, "TestScenarios");
      const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });
      saveFileWithBrowserDialog(blob, `test-scenarios-template-${datePart}.xlsx`);
      return;
    } catch (error) {
      console.warn("Falling back to CSV template download.", error);
    }

    const csv = rowsToCsv(rows);
    const csvBlob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    saveFileWithBrowserDialog(csvBlob, `test-scenarios-template-${datePart}.csv`);
    window.alert("Excel library is unavailable right now. Downloaded CSV template instead.");
  }

  function buildSummaryRowsForSpreadsheet() {
    const rows = [["No.", "Test Scenario", "Status", "UAT Results"]];
    (state.tests || []).forEach((test, index) => {
      rows.push([
        index + 1,
        richHtmlToPlainText(test.titleHtml) || test.title || "Untitled test",
        formatTestStatus(test.status),
        richHtmlToPlainText(test.notesHtml) || test.notes || "-"
      ]);
    });
    return rows;
  }

  function buildRunInfoRowsForSpreadsheet() {
    const counts = state.tests.reduce((accumulator, test) => {
      if (test.status === "not-start") accumulator.notStart += 1;
      if (test.status === "passed") accumulator.passed += 1;
      if (test.status === "query") accumulator.query += 1;
      if (test.status === "failed") accumulator.failed += 1;
      if (test.status === "blocked") accumulator.blocked += 1;
      if (test.status === "cancelled") accumulator.cancelled += 1;
      return accumulator;
    }, { notStart: 0, passed: 0, query: 0, failed: 0, blocked: 0, cancelled: 0 });

    return [
      ["Field", "Value"],
      ["Run Name", state.runName || "Test Run Summary"],
      ["Owner", state.ownerName || "-"],
      ["Queries Raised", Number(state.queryCount || 0)],
      ["Not Start", counts.notStart],
      ["Passed", counts.passed],
      ["Query", counts.query],
      ["Failed", counts.failed],
      ["Blocked", counts.blocked],
      ["Cancelled", counts.cancelled],
      ["Version Details", state.versionDetails || "-"],
      ["Exported At", new Date().toLocaleString()]
    ];
  }

  async function exportSummarySpreadsheet() {
    void saveRunSnapshot("before-export-summary-excel").catch(() => null);
    const runSlug = slugify(state.runName || "test-run");
    const summaryRows = buildSummaryRowsForSpreadsheet();

    try {
      const XLSX = await loadXlsxLibrary();
      const workbook = XLSX.utils.book_new();
      const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
      const infoSheet = XLSX.utils.aoa_to_sheet(buildRunInfoRowsForSpreadsheet());
      XLSX.utils.book_append_sheet(workbook, summarySheet, "ConsolidatedSummary");
      XLSX.utils.book_append_sheet(workbook, infoSheet, "RunInfo");
      const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });
      saveFileWithBrowserDialog(blob, `${runSlug}-summary.xlsx`);
      return;
    } catch (error) {
      console.warn("Falling back to CSV summary export.", error);
    }

    const csvBlob = new Blob([rowsToCsv(summaryRows)], { type: "text/csv;charset=utf-8" });
    saveFileWithBrowserDialog(csvBlob, `${runSlug}-summary.csv`);
    window.alert("Excel library is unavailable right now. Exported CSV summary instead.");
  }

  function getPickerFileTypeOptions(suggestedFilename, blobType) {
    const extension = suggestedFilename.includes(".")
      ? `.${suggestedFilename.split(".").pop().toLowerCase()}`
      : "";
    const mimeType = blobType || "application/octet-stream";
    if (!extension) {
      return [];
    }
    return [{
      description: `${extension.toUpperCase().slice(1)} file`,
      accept: {
        [mimeType]: [extension]
      }
    }];
  }

  function fallbackDownload(blob, suggestedFilename) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = suggestedFilename;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 500);
  }

  // Uses native picker when available so users can choose save location/name.
  function saveFileWithBrowserDialog(blob, suggestedFilename) {
    if (typeof window.showSaveFilePicker !== "function") {
      fallbackDownload(blob, suggestedFilename);
      return;
    }

    const pickerTypes = getPickerFileTypeOptions(suggestedFilename, blob.type);
    void (async () => {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName: suggestedFilename,
          types: pickerTypes.length ? pickerTypes : undefined
        });
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
      } catch (error) {
        if (error?.name === "AbortError") {
          return;
        }
        console.warn("Native save picker failed, using browser download fallback.", error);
        fallbackDownload(blob, suggestedFilename);
      }
    })();
  }

  async function buildPortableExportState() {
    const exportState = structuredClone(state);
    const allVideos = [];
    const allScreenshots = [];
    const allFiles = [];

    (exportState.tests || []).forEach((test) => {
      (test.videos || []).forEach((video) => allVideos.push(video));
      (test.screenshots || []).forEach((shot) => allScreenshots.push(shot));
      (test.files || []).forEach((file) => allFiles.push(file));
      (test.steps || []).forEach((step) => {
        (step.videos || []).forEach((video) => allVideos.push(video));
        (step.screenshots || []).forEach((shot) => allScreenshots.push(shot));
        (step.files || []).forEach((file) => allFiles.push(file));
      });
    });

    if (allVideos.length) {
      for (const video of allVideos) {
        const videoDataUrl = await getVideoDataUrlForExport(video);
        if (!videoDataUrl) {
          continue;
        }
        video.dataUrl = videoDataUrl;
        video.storage = "inline";
      }
    }

    if (allScreenshots.length) {
      for (const shot of allScreenshots) {
        if (shot.dataUrl?.startsWith("data:")) {
          shot.storage = "inline";
          continue;
        }
        const shotDataUrl = await getShotDataUrlForExport(shot);
        if (!shotDataUrl) {
          continue;
        }
        shot.dataUrl = shotDataUrl;
        shot.storage = "inline";
      }
    }

    if (allFiles.length) {
      for (const fileAttachment of allFiles) {
        if (fileAttachment.dataUrl?.startsWith("data:")) {
          fileAttachment.storage = "inline";
          continue;
        }
        const fileDataUrl = await getAttachmentDataUrlForExport(fileAttachment);
        if (!fileDataUrl) {
          continue;
        }
        fileAttachment.dataUrl = fileDataUrl;
        fileAttachment.storage = "inline";
      }
    }

    return {
      ...exportState,
      exportedAt: new Date().toISOString(),
      mediaStorage: "inline",
      mediaIncludedInJson: true
    };
  }

  function serializePortablePayloadForHtml(payload) {
    return JSON.stringify(payload)
      .replace(/</g, "\\u003c")
      .replace(/\u2028/g, "\\u2028")
      .replace(/\u2029/g, "\\u2029");
  }

  function extractPortablePayloadFromHtml(text) {
    const match = text.match(/<script[^>]*id=["']embeddedRunJson["'][^>]*>([\s\S]*?)<\/script>/i);
    return match ? match[1].trim() : "";
  }

  async function exportJson() {
    void saveRunSnapshot("before-export-json").catch(() => null);
    const exportPayload = await buildPortableExportState();
    const blob = new Blob([JSON.stringify(exportPayload, null, 2)], { type: "application/json" });
    const filename = `${slugify(state.runName || "test-run")}.json`;
    saveFileWithBrowserDialog(blob, filename);
  }

  async function startScreenRecording(target = null) {
    const resolvedTarget = resolveRecordingTarget(target);
    if (!resolvedTarget) {
      window.alert("Create a test first, then start recording.");
      return;
    }
    if (!navigator.mediaDevices?.getDisplayMedia || typeof MediaRecorder === "undefined") {
      window.alert("This browser does not support screen recording. Use a recent Edge or Chrome build.");
      return;
    }
    if (recordingSession) {
      return;
    }

    try {
      const stream = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: false });
      const mimeType = [
        "video/webm;codecs=vp8",
        "video/webm;codecs=vp9",
        "video/webm",
        "video/mp4;codecs=avc1.42E01E,mp4a.40.2",
        "video/mp4"
      ].find((candidate) => MediaRecorder.isTypeSupported(candidate));
      const recorder = mimeType
        ? new MediaRecorder(stream, { mimeType })
        : new MediaRecorder(stream);
      const chunks = [];

      recorder.ondataavailable = (event) => {
        if (event.data.size > 0) {
          chunks.push(event.data);
        }
      };

      recorder.onpause = syncRecordingControls;
      recorder.onresume = syncRecordingControls;
      recorder.onstop = async () => {
        try {
          const recordingMimeType = recorder.mimeType || mimeType || "video/webm";
          const blob = new Blob(chunks, { type: recordingMimeType });
          if (!blob.size) {
            window.alert("Recording finished but no playable media was captured. Please try again.");
            return;
          }
          const durationSeconds = await getVideoDurationSeconds(blob);
          const savedTarget = resolveRecordingTarget(recordingSession?.target || resolvedTarget);
          if (!savedTarget) {
            return;
          }
          const videoId = crypto.randomUUID();
          await saveVideoBlob(videoId, blob);
          const video = {
            id: videoId,
            name: `Recording ${new Date().toLocaleTimeString()}.${extensionFromMimeType(recordingMimeType)}`,
            notes: "",
            createdAt: new Date().toLocaleString(),
            durationSeconds,
            sizeBytes: blob.size,
            mimeType: recordingMimeType,
            storage: "indexeddb"
          };
          if (savedTarget.step) {
            savedTarget.step.videos = savedTarget.step.videos || [];
            savedTarget.step.videos.push(video);
          } else {
            savedTarget.test.videos = savedTarget.test.videos || [];
            savedTarget.test.videos.push(video);
          }
        } finally {
          stream.getTracks().forEach((track) => track.stop());
          recordingSession = null;
          render();
          saveState();
        }
      };

      stream.getVideoTracks().forEach((track) => {
        track.addEventListener("ended", () => {
          if (recorder.state !== "inactive") {
            recorder.requestData();
            recorder.stop();
          }
        }, { once: true });
      });

      recorder.start(1000);
      recordingSession = { recorder, stream, target: { testId: resolvedTarget.test.id, stepId: resolvedTarget.step?.id || null } };
      const activeCard = document.getElementById(`test-card-${resolvedTarget.test.id}`);
      if (activeCard) {
        activeCard.scrollIntoView({ behavior: "smooth", block: "center" });
      }
      syncRecordingControls();
    } catch {
      window.alert("Screen recording was cancelled or could not start.");
    }
  }

  function resolveRecordingTarget(target) {
    const fallbackTest = state.tests.at(-1);
    const requestedTestId = target?.testId || fallbackTest?.id;
    const test = state.tests.find((item) => item.id === requestedTestId);
    if (!test) {
      return null;
    }
    test.steps = Array.isArray(test.steps) ? test.steps : [];
    if (!test.steps.length) {
      test.steps.push(createStep());
    }

    const requestedStepId = target?.stepId || test.steps.at(-1)?.id;
    const step = test.steps.find((item) => item.id === requestedStepId) || null;
    if (step) {
      step.videos = Array.isArray(step.videos) ? step.videos : [];
    }
    return { test, step };
  }

  function getVideoDurationSeconds(blob) {
    return new Promise((resolve) => {
      const url = URL.createObjectURL(blob);
      const video = document.createElement("video");
      video.preload = "metadata";
      video.onloadedmetadata = () => {
        const value = Number.isFinite(video.duration) ? video.duration : null;
        URL.revokeObjectURL(url);
        resolve(value);
      };
      video.onerror = () => {
        URL.revokeObjectURL(url);
        resolve(null);
      };
      video.src = url;
    });
  }

  function extensionFromMimeType(mimeType) {
    if (!mimeType) {
      return "webm";
    }
    if (mimeType.includes("mp4")) {
      return "mp4";
    }
    if (mimeType.includes("webm")) {
      return "webm";
    }
    return "video";
  }

  function togglePauseRecording() {
    if (!recordingSession) {
      return;
    }

    const recorder = recordingSession.recorder;
    try {
      if (recorder.state === "recording") {
        if (typeof recorder.pause !== "function") {
          window.alert("Pause is not supported by this browser for the current capture mode.");
          return;
        }
        recorder.pause();
      } else if (recorder.state === "paused") {
        if (typeof recorder.resume !== "function") {
          window.alert("Resume is not supported by this browser for the current capture mode.");
          return;
        }
        recorder.resume();
      }
    } catch {
      window.alert("Pause/Resume could not be applied in this recording mode. Try Ctrl+Shift+X to stop and start a new recording.");
    }

    syncRecordingControls();
  }

  function stopScreenRecording() {
    if (!recordingSession) {
      return;
    }

    if (recordingSession.recorder.state !== "inactive") {
      recordingSession.recorder.requestData();
      recordingSession.recorder.stop();
    }

    syncRecordingControls();
  }

  function syncRecordingControls() {
    const recorderState = recordingSession?.recorder?.state || "inactive";
    if (elements.startRecordingButton) elements.startRecordingButton.disabled = recorderState !== "inactive";
    if (elements.pauseRecordingButton) {
      elements.pauseRecordingButton.disabled = recorderState === "inactive";
      elements.pauseRecordingButton.textContent = recorderState === "paused" ? "Resume" : "Pause";
    }
    if (elements.stopRecordingButton) elements.stopRecordingButton.disabled = recorderState === "inactive";
    if (elements.floatingPauseButton) elements.floatingPauseButton.textContent = recorderState === "paused" ? "Resume" : "Pause";

    if (elements.floatingRecordingControls) {
      const isRecording = recorderState !== "inactive";
      elements.floatingRecordingControls.hidden = !isRecording;
      if (elements.floatingRecordingLabel && isRecording) {
        const currentTarget = resolveRecordingTarget(recordingSession?.target || null);
        const label = currentTarget?.step
          ? `Recording test ${state.tests.findIndex((item) => item.id === currentTarget.test.id) + 1}, step ${currentTarget.test.steps.findIndex((step) => step.id === currentTarget.step.id) + 1}`
          : `Recording test ${state.tests.findIndex((item) => item.id === currentTarget?.test?.id) + 1}`;
        elements.floatingRecordingLabel.textContent = recorderState === "paused" ? `${label} (paused)` : label;
      }
    }

    // Update inline per-test record controls
    elements.testsList?.querySelectorAll(".test-card").forEach((card) => {
      const isActiveCard = recordingSession?.target?.testId === card.dataset.testId;
      const recordVideoBtn = card.querySelector(".record-video");
      const recordPauseBtn = card.querySelector(".record-pause");
      const recordStopBtn = card.querySelector(".record-stop");
      const isRecording = recorderState !== "inactive";
      if (recordVideoBtn) recordVideoBtn.style.display = (isActiveCard && isRecording) ? "none" : "";
      if (recordPauseBtn) {
        recordPauseBtn.style.display = (isActiveCard && isRecording) ? "" : "none";
        recordPauseBtn.textContent = recorderState === "paused" ? "Resume" : "Pause";
      }
      if (recordStopBtn) {
        recordStopBtn.style.display = (isActiveCard && isRecording) ? "" : "none";
      }
    });

    if (recorderState === "inactive") {
      if (elements.recordingStatus) elements.recordingStatus.textContent = "Recorder idle";
      return;
    }
    const currentTarget = resolveRecordingTarget(recordingSession?.target || null);
    const scopeLabel = currentTarget?.step
      ? `test step ${currentTarget.test.steps.findIndex((step) => step.id === currentTarget.step.id) + 1}`
      : "test";
    if (elements.recordingStatus) elements.recordingStatus.textContent = recorderState === "paused"
      ? `Recording paused for ${scopeLabel}`
      : `Recording in progress for ${scopeLabel}`;
  }

  function toggleReviewMode() {
    state.reviewMode = !state.reviewMode;
    if (state.reviewMode) {
      const shotCommentCount = state.tests.flatMap((test) => test.screenshots).filter((shot) => shot.reviewComment).length;
      const testCommentCount = state.tests.filter((test) => test.reviewComment).length;
      const reviewedCount = state.tests.filter((test) => test.reviewStatus && test.reviewStatus !== "not-reviewed").length;
      elements.reviewSummary.innerHTML = `<strong>${reviewedCount}</strong><span>tests reviewed</span><br><strong>${testCommentCount}</strong><span>test comments</span><br><strong>${shotCommentCount}</strong><span>screenshot comments</span>`;
      elements.reviewDialog.showModal();
    } else if (elements.reviewDialog.open) {
      elements.reviewDialog.close();
    }
    render();
    saveState();
  }

  async function importJson(text) {
    try {
      let parsed;
      try {
        parsed = JSON.parse(text);
      } catch {
        const embeddedPayload = extractPortablePayloadFromHtml(text);
        if (!embeddedPayload) {
          throw new Error("No embedded payload");
        }
        parsed = JSON.parse(embeddedPayload);
      }
      await saveRunSnapshot("before-import").catch(() => null);
      state = {
        ...structuredClone(defaultState),
        ...parsed,
        tests: Array.isArray(parsed.tests) ? parsed.tests : [],
        setup: parsed.setup || { textHtml: "", screenshots: [], links: "" }
      };
      normalizeTestStatusesInState();
      render();
      await migrateInlineVideosToIndexedDb();
      await migrateInlineScreenshotsAndFilesToIndexedDb();
      saveState();
    } catch {
      window.alert("The selected file is not a valid run JSON or exported HTML file.");
    }
  }

  function loadImageDimensions(dataUrl) {
    return new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight });
      image.onerror = reject;
      image.src = dataUrl;
    });
  }

  function whenImageReady(image) {
    return new Promise((resolve, reject) => {
      if (image.complete && image.naturalWidth > 0) {
        resolve();
        return;
      }
      image.onload = () => resolve();
      image.onerror = reject;
    });
  }

  function renderImageToCanvas(dataUrl, canvas, context) {
    return new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => {
        context.drawImage(image, 0, 0, canvas.width, canvas.height);
        resolve();
      };
      image.onerror = reject;
      image.src = dataUrl;
    });
  }

  function syncToolUi() {
    const isText = elements.markupTool.value === "text" || elements.markupTool.value === "callout";
    elements.textLabel.classList.toggle("hidden", !isText);
  }

  function bindTestDragEvents(card, testId) {
    card.addEventListener("dragstart", (event) => {
      if (event.target.closest("input, textarea, select, button, a, [contenteditable='true']")) {
        event.preventDefault();
        return;
      }
      event.dataTransfer.setData("text/test-id", testId);
      card.classList.add("dragging");
    });
    card.addEventListener("dragend", () => {
      card.classList.remove("dragging");
      elements.testsList.querySelectorAll(".drop-target").forEach((element) => element.classList.remove("drop-target"));
    });
    card.addEventListener("dragover", (event) => {
      event.preventDefault();
      card.classList.add("drop-target");
    });
    card.addEventListener("dragleave", () => card.classList.remove("drop-target"));
    card.addEventListener("drop", (event) => {
      event.preventDefault();
      card.classList.remove("drop-target");
      const sourceId = event.dataTransfer.getData("text/test-id");
      if (!sourceId || sourceId === testId) {
        return;
      }
      reorderItem(state.tests, sourceId, testId);
      render();
      saveState();
    });
  }

  function bindShotDragEvents(card, testId, shotId, stepId = null) {
    card.addEventListener("dragstart", (event) => {
      event.dataTransfer.setData("text/shot", JSON.stringify({ testId, shotId, stepId }));
      card.classList.add("dragging");
    });
    card.addEventListener("dragend", () => {
      card.classList.remove("dragging");
      elements.testsList.querySelectorAll(".drop-target").forEach((element) => element.classList.remove("drop-target"));
    });
    card.addEventListener("dragover", (event) => {
      event.preventDefault();
      card.classList.add("drop-target");
    });
    card.addEventListener("dragleave", () => card.classList.remove("drop-target"));
    card.addEventListener("drop", (event) => {
      event.preventDefault();
      event.stopPropagation();
      card.classList.remove("drop-target");
      const payload = event.dataTransfer.getData("text/shot");
      if (!payload) {
        return;
      }
      const source = JSON.parse(payload);
      if (source.testId !== testId || source.shotId === shotId) {
        return;
      }
      moveAttachment(testId, "shot", source.stepId || null, source.shotId, stepId, shotId);
      render();
      saveState();
    });
  }

  function bindVideoDragEvents(card, testId, videoId, stepId = null) {
    card.addEventListener("dragstart", (event) => {
      event.dataTransfer.setData("text/video", JSON.stringify({ testId, videoId, stepId }));
      card.classList.add("dragging");
    });
    card.addEventListener("dragend", () => {
      card.classList.remove("dragging");
      elements.testsList.querySelectorAll(".drop-target").forEach((element) => element.classList.remove("drop-target"));
    });
    card.addEventListener("dragover", (event) => {
      event.preventDefault();
      card.classList.add("drop-target");
    });
    card.addEventListener("dragleave", () => card.classList.remove("drop-target"));
    card.addEventListener("drop", (event) => {
      event.preventDefault();
      event.stopPropagation();
      card.classList.remove("drop-target");
      const payload = event.dataTransfer.getData("text/video");
      if (!payload) {
        return;
      }
      const source = JSON.parse(payload);
      if (source.testId !== testId || source.videoId === videoId) {
        return;
      }
      moveAttachment(testId, "video", source.stepId || null, source.videoId, stepId, videoId);
      render();
      saveState();
    });
  }

  function setSelectedDropTarget(target) {
    selectedDropTarget = {
      kind: target.kind,
      testId: target.testId,
      stepId: target.stepId || null
    };
    refreshSelectedDropTargetUi();
  }

  function isDropTargetSelected(target) {
    return Boolean(
      selectedDropTarget
      && selectedDropTarget.kind === target.kind
      && selectedDropTarget.testId === target.testId
      && (selectedDropTarget.stepId || null) === (target.stepId || null)
    );
  }

  function refreshSelectedDropTargetUi() {
    document.querySelectorAll(".media-drop-zone").forEach((zone) => {
      zone.classList.remove("selected-target");
    });
    if (!selectedDropTarget) {
      return;
    }
    const stepId = selectedDropTarget.stepId || "";
    const selector = `.media-drop-zone[data-kind="${selectedDropTarget.kind}"][data-test-id="${selectedDropTarget.testId}"][data-step-id="${stepId}"]`;
    const targetZone = document.querySelector(selector);
    if (targetZone) {
      targetZone.classList.add("selected-target");
    }
  }

  async function attachVideoFiles(testId, files, stepId = null) {
    const test = state.tests.find((item) => item.id === testId);
    if (!test) {
      return;
    }
    const targetStep = stepId ? (test.steps || []).find((step) => step.id === stepId) : null;
    if (targetStep) {
      targetStep.videos = Array.isArray(targetStep.videos) ? targetStep.videos : [];
    } else {
      test.videos = Array.isArray(test.videos) ? test.videos : [];
    }

    const videos = [];
    for (const file of files) {
      if (!file || file.size <= 0) {
        continue;
      }
      const videoId = crypto.randomUUID();
      await saveVideoBlob(videoId, file);
      const durationSeconds = await getVideoDurationSeconds(file).catch(() => 0);
      videos.push({
        id: videoId,
        name: file.name || `Video ${new Date().toLocaleTimeString()}.${extensionFromMimeType(file.type || "video/webm")}`,
        subject: "",
        subjectHtml: "",
        notes: "",
        notesHtml: "",
        createdAt: new Date().toLocaleString(),
        durationSeconds,
        sizeBytes: file.size,
        mimeType: file.type || "video/webm",
        storage: "indexeddb"
      });
    }

    if (!videos.length) {
      return;
    }
    if (targetStep) {
      targetStep.videos.push(...videos);
      return;
    }
    test.videos.push(...videos);
  }

  function bindAttachmentDropZone(element, options) {
    element.dataset.kind = options.kind;
    element.dataset.testId = options.testId;
    element.dataset.stepId = options.stepId || "";
    element.dataset.emptyLabel = options.emptyLabel;
    element.classList.add("media-drop-zone");
    element.addEventListener("click", () => {
      setSelectedDropTarget(options);
    });
    if (isDropTargetSelected(options)) {
      element.classList.add("selected-target");
    }
    element.addEventListener("dragover", (event) => {
      const payloadType = options.kind === "shot" ? "text/shot" : "text/video";
      const hasInternalPayload = event.dataTransfer.types.includes(payloadType);
      const droppedFiles = Array.from(event.dataTransfer.files || []);
      const hasMatchingFileType = options.kind === "shot"
        ? droppedFiles.some((file) => file.type.startsWith("image/"))
        : droppedFiles.some((file) => file.type.startsWith("video/"));
      if (!hasInternalPayload && !hasMatchingFileType) {
        return;
      }
      event.preventDefault();
      element.classList.add("drop-target");
    });
    element.addEventListener("dragleave", (event) => {
      if (element.contains(event.relatedTarget)) {
        return;
      }
      element.classList.remove("drop-target");
    });
    element.addEventListener("drop", async (event) => {
      const payloadType = options.kind === "shot" ? "text/shot" : "text/video";
      const payload = event.dataTransfer.getData(payloadType);
      event.preventDefault();
      element.classList.remove("drop-target");
      setSelectedDropTarget(options);

      if (payload) {
        const source = JSON.parse(payload);
        const sourceId = options.kind === "shot" ? source.shotId : source.videoId;
        if (source.testId !== options.testId) {
          return;
        }
        moveAttachment(options.testId, options.kind, source.stepId || null, sourceId, options.stepId || null);
        render();
        saveState();
        return;
      }

      const droppedFiles = Array.from(event.dataTransfer.files || []).filter((file) => file && file.size > 0);
      if (!droppedFiles.length) {
        return;
      }

      if (options.kind === "shot") {
        const imageFiles = droppedFiles.filter((file) => file.type.startsWith("image/"));
        if (!imageFiles.length) {
          return;
        }
        await attachScreenshots(options.testId, imageFiles, options.stepId || null);
      } else {
        const videoFiles = droppedFiles.filter((file) => file.type.startsWith("video/"));
        if (!videoFiles.length) {
          return;
        }
        await attachVideoFiles(options.testId, videoFiles, options.stepId || null);
      }
      render();
      saveState();
    });
  }

  function bindStepDragEvents(row, testId, stepId) {
    row.addEventListener("dragstart", (event) => {
      event.dataTransfer.setData("text/step", JSON.stringify({ testId, stepId }));
      row.classList.add("dragging");
    });
    row.addEventListener("dragend", () => {
      row.classList.remove("dragging");
      row.parentElement?.querySelectorAll(".drop-target").forEach((element) => element.classList.remove("drop-target"));
    });
    row.addEventListener("dragover", (event) => {
      event.preventDefault();
      row.classList.add("drop-target");
    });
    row.addEventListener("dragleave", () => row.classList.remove("drop-target"));
    row.addEventListener("drop", (event) => {
      event.preventDefault();
      row.classList.remove("drop-target");
      const payload = event.dataTransfer.getData("text/step");
      if (!payload) {
        return;
      }
      const source = JSON.parse(payload);
      if (source.testId !== testId || source.stepId === stepId) {
        return;
      }
      const test = state.tests.find((item) => item.id === testId);
      if (!test) {
        return;
      }
      reorderItem(test.steps, source.stepId, stepId);
      render();
      saveState();
    });
  }

  function reorderItem(items, sourceId, targetId) {
    const sourceIndex = items.findIndex((item) => item.id === sourceId);
    const targetIndex = items.findIndex((item) => item.id === targetId);
    if (sourceIndex < 0 || targetIndex < 0) {
      return;
    }
    const [sourceItem] = items.splice(sourceIndex, 1);
    items.splice(targetIndex, 0, sourceItem);
  }

  function formatReviewStatus(status) {
    return {
      "not-reviewed": "Not Reviewed",
      approved: "Approved",
      "needs-changes": "Needs Changes",
      question: "Question"
    }[status] || "Not Reviewed";
  }

  function formatTestStatus(status) {
    return {
      "not-start": "Not Start",
      passed: "Passed",
      query: "Query",
      failed: "Failed",
      blocked: "Blocked",
      cancelled: "Cancelled"
    }[normalizeTestStatus(status)] || "Not Start";
  }

  function statusCellStyle(status) {
    const base = "padding:10px 12px;border-bottom:1px solid #eee3cf;vertical-align:top;font-weight:700;border-radius:4px;";
    const styles = {
      passed:    "background:#e7f3ef;color:#0b5a4c;",
      query:     "background:#fff1d6;color:#9a6200;",
      failed:    "background:#fbe7e4;color:#8a3026;",
      blocked:   "background:#ece8dd;color:#6d6a63;",
      cancelled: "background:#ece8dd;color:#6d6a63;",
    };
    return base + (styles[normalizeTestStatus(status)] || "");
  }

  function applyTestStatusSelectStyle(selectElement, status) {
    if (!selectElement) {
      return;
    }
    selectElement.classList.remove(
      "status-not-start",
      "status-passed",
      "status-query",
      "status-failed",
      "status-blocked",
      "status-cancelled"
    );
    const normalized = normalizeTestStatus(status);
    selectElement.classList.add(`status-${normalized}`);
  }

  function bindDrawingDialogShortcuts() {
    elements.drawingDialog.addEventListener("keydown", (event) => {
      if (!drawingSession) {
        return;
      }
      const targetTag = event.target.tagName;
      const isTyping = targetTag === "INPUT" || targetTag === "TEXTAREA" || targetTag === "SELECT";
      if (!isTyping && toolShortcutMap[event.code]) {
        event.preventDefault();
        elements.markupTool.value = toolShortcutMap[event.code];
        syncToolUi();
      }
      if (event.code === "Delete" || event.code === "Backspace") {
        if (drawingSession.selectedIndex >= 0 && !isTyping) {
          event.preventDefault();
          drawingSession.items.splice(drawingSession.selectedIndex, 1);
          drawingSession.selectedIndex = -1;
          drawingSession.dirty = true;
          redrawDrawingCanvas();
        }
      }
      if ((event.ctrlKey || event.metaKey) && event.code === "KeyC") {
        event.preventDefault();
        copySelectedItem();
      }
      if ((event.ctrlKey || event.metaKey) && event.code === "KeyV") {
        event.preventDefault();
        pasteClipboardItem();
      }
      if ((event.ctrlKey || event.metaKey) && event.code === "KeyD") {
        event.preventDefault();
        duplicateSelectedItem();
      }
      if ((event.ctrlKey || event.metaKey) && event.code === "KeyZ") {
        event.preventDefault();
        undoStroke();
      }
      if (event.code === "F1") {
        event.preventDefault();
        elements.drawingHelpPanel.classList.toggle("hidden");
      }
    });
  }

  function tryStartSelectionInteraction(event, point) {
    const context = elements.drawingCanvas.getContext("2d");
    const selectedItem = drawingSession.items[drawingSession.selectedIndex];
    if (selectedItem) {
      const selectedBounds = getItemBounds(selectedItem, drawingSession, context);
      const handle = selectedBounds ? findHandleAtPoint(selectedBounds, point, drawingSession) : null;
      if (handle) {
        drawingSession.interactionMode = "resize";
        drawingSession.activeHandle = handle;
        drawingSession.dragOrigin = point;
        return true;
      }
    }

    const hitIndex = findItemAtPoint(point, context);
    if (hitIndex >= 0) {
      drawingSession.selectedIndex = hitIndex;
      drawingSession.interactionMode = "move";
      drawingSession.dragOrigin = point;
      return true;
    }

    drawingSession.selectedIndex = -1;
    drawingSession.interactionMode = null;
    drawingSession.dragOrigin = null;
    return false;
  }

  function findItemAtPoint(point, context) {
    for (let index = drawingSession.items.length - 1; index >= 0; index -= 1) {
      const bounds = getItemBounds(drawingSession.items[index], drawingSession, context, true);
      if (!bounds) {
        continue;
      }
      if (point.x >= bounds.naturalX && point.x <= bounds.naturalX + bounds.naturalWidth && point.y >= bounds.naturalY && point.y <= bounds.naturalY + bounds.naturalHeight) {
        return index;
      }
    }
    return -1;
  }

  function getItemBounds(item, session, context, includeNatural) {
    const scaleX = context.canvas.width / session.naturalWidth;
    const scaleY = context.canvas.height / session.naturalHeight;
    let naturalX;
    let naturalY;
    let naturalWidth;
    let naturalHeight;

    if (!item.type || item.type === "pen") {
      if (!item.points?.length) {
        return null;
      }
      const xs = item.points.map((point) => point.x);
      const ys = item.points.map((point) => point.y);
      naturalX = Math.min(...xs) - item.size;
      naturalY = Math.min(...ys) - item.size;
      naturalWidth = Math.max(...xs) - Math.min(...xs) + item.size * 2;
      naturalHeight = Math.max(...ys) - Math.min(...ys) + item.size * 2;
    } else if (item.type === "text") {
      const fontSize = Math.max(14, item.size * 4);
      context.font = `${fontSize * scaleX}px Segoe UI`;
      const wrapped = wrapText(context, item.text || "Text", Math.max(100, (item.width || 220) * scaleX));
      naturalX = item.start.x;
      naturalY = item.start.y;
      naturalWidth = Math.max(100, item.width || 220);
      naturalHeight = Math.max(fontSize * 1.35, wrapped.length * fontSize * 1.35);
    } else if (item.type === "callout") {
      const radius = Math.max(14, item.size * 3);
      naturalX = item.start.x - radius;
      naturalY = item.start.y - radius;
      naturalWidth = radius * 2 + (item.text ? Math.max(100, item.width || 180) : 0);
      naturalHeight = Math.max(radius * 2, item.text ? item.size * 9 : radius * 2);
    } else {
      naturalX = Math.min(item.start.x, item.end.x);
      naturalY = Math.min(item.start.y, item.end.y);
      naturalWidth = Math.abs(item.end.x - item.start.x);
      naturalHeight = Math.abs(item.end.y - item.start.y);
    }

    const bounds = {
      x: naturalX * scaleX,
      y: naturalY * scaleY,
      width: Math.max(naturalWidth * scaleX, 14),
      height: Math.max(naturalHeight * scaleY, 14),
      naturalX,
      naturalY,
      naturalWidth: Math.max(naturalWidth, 6),
      naturalHeight: Math.max(naturalHeight, 6)
    };
    return includeNatural ? bounds : bounds;
  }

  function getHandleRects(bounds) {
    const size = 10;
    const handles = [
      { id: "nw", x: bounds.x - size / 2, y: bounds.y - size / 2, size },
      { id: "ne", x: bounds.x + bounds.width - size / 2, y: bounds.y - size / 2, size },
      { id: "sw", x: bounds.x - size / 2, y: bounds.y + bounds.height - size / 2, size },
      { id: "se", x: bounds.x + bounds.width - size / 2, y: bounds.y + bounds.height - size / 2, size }
    ];
    handles.push({ id: "e", x: bounds.x + bounds.width - size / 2, y: bounds.y + bounds.height / 2 - size / 2, size });
    return handles;
  }

  function findHandleAtPoint(bounds, point, session) {
    const scaleX = elements.drawingCanvas.width / session.naturalWidth;
    const scaleY = elements.drawingCanvas.height / session.naturalHeight;
    const screenX = point.x * scaleX;
    const screenY = point.y * scaleY;
    return getHandleRects(bounds).find((handle) => screenX >= handle.x && screenX <= handle.x + handle.size && screenY >= handle.y && screenY <= handle.y + handle.size)?.id || null;
  }

  function updateSelectedItemPosition(point) {
    const item = drawingSession.items[drawingSession.selectedIndex];
    if (!item || !drawingSession.dragOrigin) {
      return;
    }
    const dx = point.x - drawingSession.dragOrigin.x;
    const dy = point.y - drawingSession.dragOrigin.y;
    const snapped = getSnappedDelta(item, dx, dy);
    translateItem(item, snapped.dx, snapped.dy);
    drawingSession.dragOrigin = { x: drawingSession.dragOrigin.x + snapped.dx, y: drawingSession.dragOrigin.y + snapped.dy };
  }

  function translateItem(item, dx, dy) {
    if (!item.type || item.type === "pen") {
      item.points = item.points.map((point) => ({ x: point.x + dx, y: point.y + dy }));
      return;
    }
    if (item.type === "text" || item.type === "callout") {
      item.start = { x: item.start.x + dx, y: item.start.y + dy };
      return;
    }
    item.start = { x: item.start.x + dx, y: item.start.y + dy };
    item.end = { x: item.end.x + dx, y: item.end.y + dy };
  }

  function resizeSelectedItem(point) {
    const item = drawingSession.items[drawingSession.selectedIndex];
    if (!item || !drawingSession.activeHandle) {
      return;
    }
    if (!item.type || item.type === "pen") {
      translateItem(item, point.x - drawingSession.dragOrigin.x, point.y - drawingSession.dragOrigin.y);
      drawingSession.dragOrigin = point;
      return;
    }
    if (item.type === "text" || item.type === "callout") {
      if (drawingSession.activeHandle === "e") {
        item.width = Math.max(100, point.x - item.start.x);
      } else {
        item.size = Math.max(2, Math.round(distanceBetween(item.start, point) / 12));
      }
      return;
    }
    const nextStart = { ...item.start };
    const nextEnd = { ...item.end };
    if (drawingSession.activeHandle.includes("n")) nextStart.y = point.y;
    if (drawingSession.activeHandle.includes("s")) nextEnd.y = point.y;
    if (drawingSession.activeHandle.includes("w")) nextStart.x = point.x;
    if (drawingSession.activeHandle.includes("e")) nextEnd.x = point.x;
    item.start = nextStart;
    item.end = nextEnd;
  }

  function getNextCalloutNumber() {
    const numbers = drawingSession.items
      .filter((item) => item.type === "callout")
      .map((item) => Number(item.number) || 0);
    return (numbers.length ? Math.max(...numbers) : 0) + 1;
  }

  function distanceBetween(firstPoint, secondPoint) {
    return Math.hypot(secondPoint.x - firstPoint.x, secondPoint.y - firstPoint.y);
  }

  function updateSelectedItemText() {
    if (!drawingSession || drawingSession.selectedIndex < 0) {
      return;
    }
    const item = drawingSession.items[drawingSession.selectedIndex];
    if (!item || (item.type !== "text" && item.type !== "callout")) {
      return;
    }
    item.text = elements.selectedTextInput.value;
    drawingSession.dirty = true;
    redrawDrawingCanvas();
  }

  function syncSelectedItemControls() {
    const selectedItem = drawingSession?.items?.[drawingSession.selectedIndex];
    const supportsText = selectedItem && (selectedItem.type === "text" || selectedItem.type === "callout");
    elements.selectedTextLabel.classList.toggle("hidden", !supportsText);
    elements.selectedTextInput.value = supportsText ? (selectedItem.text || "") : "";
    elements.duplicateItemButton.disabled = !selectedItem;
    elements.copyItemButton.disabled = !selectedItem;
    elements.pasteItemButton.disabled = !drawingClipboard;
  }

  function copySelectedItem() {
    if (!drawingSession || drawingSession.selectedIndex < 0) {
      return;
    }
    const item = drawingSession.items[drawingSession.selectedIndex];
    if (!item) {
      return;
    }
    drawingClipboard = structuredClone(item);
    syncSelectedItemControls();
  }

  function pasteClipboardItem() {
    if (!drawingSession || !drawingClipboard) {
      return;
    }
    const clone = structuredClone(drawingClipboard);
    offsetItem(clone, 18, 18);
    if (clone.type === "callout") {
      clone.number = getNextCalloutNumber();
    }
    drawingSession.items.push(clone);
    drawingSession.selectedIndex = drawingSession.items.length - 1;
    drawingSession.dirty = true;
    redrawDrawingCanvas();
  }

  function duplicateSelectedItem() {
    copySelectedItem();
    pasteClipboardItem();
  }

  function offsetItem(item, dx, dy) {
    if (!item.type || item.type === "pen") {
      item.points = item.points.map((point) => ({ x: point.x + dx, y: point.y + dy }));
      return;
    }
    if (item.type === "text" || item.type === "callout") {
      item.start = { x: item.start.x + dx, y: item.start.y + dy };
      return;
    }
    item.start = { x: item.start.x + dx, y: item.start.y + dy };
    item.end = { x: item.end.x + dx, y: item.end.y + dy };
  }

  function getSnappedDelta(item, dx, dy) {
    const context = elements.drawingCanvas.getContext("2d");
    const currentBounds = getItemBounds(item, drawingSession, context, true);
    if (!currentBounds) {
      return { dx, dy };
    }
    const movingBounds = {
      ...currentBounds,
      naturalX: currentBounds.naturalX + dx,
      naturalY: currentBounds.naturalY + dy
    };
    const threshold = 8;
    const guides = [];
    const candidatesX = [drawingSession.naturalWidth / 2];
    const candidatesY = [drawingSession.naturalHeight / 2];
    drawingSession.items.forEach((otherItem, index) => {
      if (index === drawingSession.selectedIndex) {
        return;
      }
      const bounds = getItemBounds(otherItem, drawingSession, context, true);
      if (!bounds) {
        return;
      }
      candidatesX.push(bounds.naturalX, bounds.naturalX + bounds.naturalWidth, bounds.naturalX + bounds.naturalWidth / 2);
      candidatesY.push(bounds.naturalY, bounds.naturalY + bounds.naturalHeight, bounds.naturalY + bounds.naturalHeight / 2);
    });
    let snappedDx = dx;
    let snappedDy = dy;
    const movingXs = [movingBounds.naturalX, movingBounds.naturalX + movingBounds.naturalWidth, movingBounds.naturalX + movingBounds.naturalWidth / 2];
    const movingYs = [movingBounds.naturalY, movingBounds.naturalY + movingBounds.naturalHeight, movingBounds.naturalY + movingBounds.naturalHeight / 2];
    for (const candidate of candidatesX) {
      const hit = movingXs.find((moving) => Math.abs(candidate - moving) <= threshold);
      if (hit !== undefined) {
        snappedDx += candidate - hit;
        guides.push({ axis: "x", value: candidate });
        break;
      }
    }
    for (const candidate of candidatesY) {
      const hit = movingYs.find((moving) => Math.abs(candidate - moving) <= threshold);
      if (hit !== undefined) {
        snappedDy += candidate - hit;
        guides.push({ axis: "y", value: candidate });
        break;
      }
    }
    drawingSession.snapGuides = guides;
    return { dx: snappedDx, dy: snappedDy };
  }

  function drawSnapGuides(context) {
    if (!drawingSession?.snapGuides?.length) {
      return;
    }
    const scaleX = context.canvas.width / drawingSession.naturalWidth;
    const scaleY = context.canvas.height / drawingSession.naturalHeight;
    context.save();
    context.strokeStyle = "rgba(12, 108, 97, 0.55)";
    context.setLineDash([5, 5]);
    drawingSession.snapGuides.forEach((guide) => {
      context.beginPath();
      if (guide.axis === "x") {
        const x = guide.value * scaleX;
        context.moveTo(x, 0);
        context.lineTo(x, context.canvas.height);
      } else {
        const y = guide.value * scaleY;
        context.moveTo(0, y);
        context.lineTo(context.canvas.width, y);
      }
      context.stroke();
    });
    context.restore();
  }

  function wrapText(context, text, maxWidth) {
    const segments = String(text || "").split(/\n/);
    const lines = [];
    segments.forEach((segment) => {
      const words = segment.split(/\s+/).filter(Boolean);
      if (!words.length) {
        lines.push("");
        return;
      }
      let currentLine = words.shift();
      words.forEach((word) => {
        const trial = `${currentLine} ${word}`;
        if (context.measureText(trial).width <= maxWidth) {
          currentLine = trial;
        } else {
          lines.push(currentLine);
          currentLine = word;
        }
      });
      lines.push(currentLine);
    });
    return lines;
  }

  function hexToRgba(hex, alpha) {
    const normalized = hex.replace("#", "");
    const value = normalized.length === 3
      ? normalized.split("").map((char) => char + char).join("")
      : normalized;
    const red = parseInt(value.slice(0, 2), 16);
    const green = parseInt(value.slice(2, 4), 16);
    const blue = parseInt(value.slice(4, 6), 16);
    return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function slugify(value) {
    return String(value).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "");
  }
})();