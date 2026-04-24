const QUEUE_STAGE_PROGRESS = {
  output_folder_selected: {
    label: "Starting",
    floor: 4,
    cap: 12,
    baseRate: 18,
    easeRate: 0.32,
    tone: "active",
  },
  parse_started: {
    label: "Preparing",
    floor: 12,
    cap: 28,
    baseRate: 10,
    easeRate: 0.24,
    tone: "active",
  },
  filename_built: {
    label: "Naming",
    floor: 24,
    cap: 42,
    baseRate: 8,
    easeRate: 0.18,
    tone: "active",
  },
  pdf_pipeline_started: {
    label: "Loading",
    floor: 38,
    cap: 58,
    baseRate: 7,
    easeRate: 0.15,
    tone: "active",
  },
  pipeline_selected: {
    label: "Creating PDF",
    floor: 54,
    cap: 84,
    baseRate: 2.4,
    easeRate: 0.08,
    tone: "active",
  },
  pdf_written: {
    label: "PDF ready",
    floor: 82,
    cap: 93,
    baseRate: 5.2,
    easeRate: 0.18,
    tone: "active",
  },
  deliver_started: {
    label: "Saving",
    floor: 92,
    cap: 98,
    baseRate: 3.6,
    easeRate: 0.22,
    tone: "active",
  },
  complete: {
    label: "Saved",
    floor: 100,
    cap: 100,
    baseRate: 1,
    easeRate: 1,
    tone: "complete",
  },
  failed: {
    label: "Failed",
    floor: 100,
    cap: 100,
    baseRate: 1,
    easeRate: 1,
    tone: "failed",
  },
};

const STAGE_LABELS = {
  drop_received: "Queued",
  files_accepted: "Queued",
  outlook_extract_started: "Importing from Outlook",
  output_folder_selected: "Starting",
  parse_started: "Preparing",
  filename_built: "Naming",
  pdf_pipeline_started: "Loading tools",
  pipeline_selected: "Creating PDF",
  pdf_written: "PDF ready",
  deliver_started: "Saving",
  complete: "Saved",
  failed: "Needs attention",
};

const DEFAULT_DROP_COPY = "Supports Outlook drags and .msg files. Click to browse if you prefer.";
const SERVER_DROP_COPY = "Supports .msg files. Click to browse if you prefer.";

let queueProgressAnimationFrame = 0;

const state = {
  maxFiles: 25,
  items: [],
  queueProgressByTaskId: {},
  pendingRemovalsById: {},
  outputDir: "",
  outputDirLabel: "",
  activeConvertIds: [],
  celebratoryPulseTimer: 0,
  isBusy: false,
  explainerOpen: false,
  latestStatus: null,
  serverMode: false,
  capabilities: {
    nativeOutputPicker: true,
    outlookImport: true,
  },
};

const elements = {
  clearButton: document.getElementById("clear-button"),
  convertButton: document.getElementById("convert-button"),
  dropzone: document.getElementById("dropzone"),
  dropzoneCopy: document.getElementById("dropzone-copy"),
  appEyebrow: document.getElementById("app-eyebrow"),
  explainerBackdrop: document.getElementById("explainer-backdrop"),
  explainerClose: document.getElementById("explainer-close"),
  explainerCopy: document.getElementById("explainer-copy"),
  explainerModal: document.getElementById("explainer-modal"),
  fileInput: document.getElementById("file-input"),
  helperRow: document.getElementById("helper-row"),
  howItWorksButton: document.getElementById("how-it-works-button"),
  queueCountBadge: document.getElementById("queue-count-badge"),
  queueEmpty: document.getElementById("queue-empty"),
  queueList: document.getElementById("queue-list"),
  queuePanel: document.querySelector(".queue-panel"),
  shell: document.querySelector(".shell"),
  simpleStatus: document.getElementById("simple-status"),
  statusDetail: document.getElementById("status-detail"),
  statusHeadline: document.getElementById("status-headline"),
  topbarBody: document.getElementById("topbar-body"),
  workflowBody: document.getElementById("workflow-body"),
};

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function defaultStatusDetail() {
  if (state.outputDir) {
    return `Server save folder: ${state.outputDir}`;
  }
  if (!state.capabilities.nativeOutputPicker) {
    return "The hosted deployment expects a server-managed save location.";
  }
  return "You will choose a save folder after clicking Convert to PDF.";
}

function refreshHostingCopy() {
  if (elements.appEyebrow) {
    elements.appEyebrow.textContent = state.serverMode ? "Hosted PDF Converter" : "Local PDF Converter";
  }
  if (elements.topbarBody) {
    elements.topbarBody.innerHTML = state.capabilities.outlookImport
      ? "Drop Outlook emails or <code>.msg</code> files, then convert them to PDFs with the correct filename prefix."
      : "Upload <code>.msg</code> files, then convert them to PDFs with the correct filename prefix.";
  }
  if (elements.workflowBody) {
    elements.workflowBody.textContent = state.outputDir
      ? `Upload .msg files or drag them into the page. Converted PDFs will be saved to ${state.outputDir}.`
      : state.capabilities.outlookImport
        ? "Drag from Outlook or click to browse. When you are ready, click Convert to PDF and choose where to save the files."
        : "Click to browse for .msg files or drag them into the page, then convert them to PDFs.";
  }
  if (elements.explainerCopy) {
    elements.explainerCopy.textContent = state.outputDir
      ? `Drop one email or a whole batch, click Convert to PDF, and the app saves each PDF to ${state.outputDir} with the correct filename prefix.`
      : "Drop one email or a whole batch, click Convert to PDF, choose a save folder, and the app creates a separate PDF for each message with the correct filename prefix.";
  }
}

function lastPathSegment(value) {
  const parts = String(value || "")
    .split(/[/\\]/)
    .filter(Boolean);
  return parts[parts.length - 1] || value || "";
}

function formatSource(source) {
  if (source === "outlook") {
    return "From Outlook";
  }
  if (source === "upload") {
    return "Added from browser";
  }
  return "Queued locally";
}

function formatQueuePercent(progress) {
  return `${Math.round(progress?.percent || 0)}%`;
}

function queueStageLabel(stage) {
  return STAGE_LABELS[stage] || "Queued";
}

function queueDocumentLabel(item) {
  if (item?.stage === "complete") {
    return "PDF";
  }
  if (item?.stage === "failed") {
    return "ERR";
  }
  return "MSG";
}

function queueVisualState(item, progress) {
  const stage = item?.stage || "";
  if (stage === "complete" || progress?.tone === "complete") {
    return { tone: "complete", label: "Saved" };
  }
  if (stage === "failed" || progress?.tone === "failed") {
    return { tone: "failed", label: "Needs attention" };
  }
  if (progress?.active) {
    return { tone: "active", label: progress.label || "Converting" };
  }
  return { tone: "queued", label: "Ready" };
}

function deriveItemSummary(item, progress) {
  if (item.stage === "failed") {
    return item.error || "This file could not be converted.";
  }
  if (item.stage === "complete") {
    return item.outputPath ? `Saved as ${lastPathSegment(item.outputPath)}` : "PDF saved.";
  }
  if (progress?.active) {
    return `${progress.label || queueStageLabel(item.stage)} ${formatQueuePercent(progress)}`;
  }
  return formatSource(item.source);
}

function buildQueueProgressStateFromSnapshot(item) {
  if (!item?.stage || !QUEUE_STAGE_PROGRESS[item.stage]) {
    return null;
  }
  const model = QUEUE_STAGE_PROGRESS[item.stage];
  const isTerminal = item.stage === "complete" || item.stage === "failed";
  return {
    active: true,
    stage: item.stage,
    label: model.label,
    tone: model.tone,
    percent: isTerminal ? model.cap : model.floor,
    floorPercent: model.floor,
    capPercent: model.cap,
    baseRate: model.baseRate,
    easeRate: model.easeRate,
    lastFrameAt: performance.now(),
  };
}

function queueProgressStateForItem(item) {
  return state.queueProgressByTaskId[item.taskId] || buildQueueProgressStateFromSnapshot(item);
}

function startQueueProgressLoop() {
  if (queueProgressAnimationFrame) {
    return;
  }

  const tick = () => {
    const now = performance.now();
    let shouldContinue = false;
    let changed = false;
    const nextState = { ...state.queueProgressByTaskId };

    Object.entries(nextState).forEach(([taskId, progress]) => {
      if (!progress?.active) {
        return;
      }

      const deltaSeconds = Math.max(0, now - (progress.lastFrameAt || now)) / 1000;
      let nextPercent = progress.percent;

      if (progress.percent < progress.floorPercent) {
        const floorRemaining = progress.floorPercent - progress.percent;
        nextPercent += deltaSeconds * (progress.baseRate * 1.6 + floorRemaining * 0.6);
      } else if (progress.percent < progress.capPercent) {
        const remaining = progress.capPercent - progress.percent;
        nextPercent += deltaSeconds * (progress.baseRate + remaining * progress.easeRate);
      }

      const clampedPercent = Math.min(progress.capPercent, Math.max(progress.percent, nextPercent));

      if (Math.abs(clampedPercent - progress.percent) >= 0.02) {
        nextState[taskId] = {
          ...progress,
          percent: clampedPercent,
          lastFrameAt: now,
        };
        changed = true;
      }

      if (clampedPercent < progress.capPercent - 0.02) {
        shouldContinue = true;
      }
    });

    if (changed) {
      state.queueProgressByTaskId = nextState;
      renderQueue();
    }

    if (shouldContinue) {
      queueProgressAnimationFrame = requestAnimationFrame(tick);
      return;
    }

    queueProgressAnimationFrame = 0;
  };

  queueProgressAnimationFrame = requestAnimationFrame(tick);
}

function activateQueueProgress(taskId, stage = "output_folder_selected") {
  const model = QUEUE_STAGE_PROGRESS[stage] || QUEUE_STAGE_PROGRESS.output_folder_selected;
  state.queueProgressByTaskId = {
    ...state.queueProgressByTaskId,
    [taskId]: {
      active: true,
      stage,
      label: model.label,
      tone: model.tone,
      percent: 0,
      floorPercent: model.floor,
      capPercent: model.cap,
      baseRate: model.baseRate,
      easeRate: model.easeRate,
      lastFrameAt: performance.now(),
    },
  };
}

function updateQueueProgress(taskId, stage) {
  if (!taskId || !stage) {
    return;
  }

  const previous = state.queueProgressByTaskId[taskId];
  const model = QUEUE_STAGE_PROGRESS[stage];
  if (!model && !previous?.active) {
    return;
  }

  const currentPercent = previous?.percent || 0;
  const nextFloor = Math.max(currentPercent, model?.floor || currentPercent);
  const nextCap = Math.max(nextFloor, model?.cap || currentPercent);

  state.queueProgressByTaskId = {
    ...state.queueProgressByTaskId,
    [taskId]: {
      active: true,
      stage,
      label: model?.label || queueStageLabel(stage),
      tone: model?.tone || previous?.tone || "active",
      percent: currentPercent,
      floorPercent: nextFloor,
      capPercent: nextCap,
      baseRate: model?.baseRate || previous?.baseRate || 6,
      easeRate: model?.easeRate || previous?.easeRate || 0.14,
      lastFrameAt: performance.now(),
    },
  };

  if (state.items.some((item) => item.taskId === taskId)) {
    renderQueue();
  }
  startQueueProgressLoop();
}

function setStatus(headline, detail = "", tone = "neutral") {
  state.latestStatus = { headline, detail, tone };
  renderStatus();
}

function addStatus(headline, detail = "", tone = "neutral") {
  setStatus(headline, detail, tone);
}

function renderStatus() {
  if (!elements.helperRow || !elements.simpleStatus || !elements.statusHeadline || !elements.statusDetail) {
    return;
  }
  const hasVisibleStatus = Boolean(state.latestStatus?.headline || state.latestStatus?.detail);
  const headline = state.latestStatus?.headline || "Ready for files.";
  const detail = state.latestStatus?.detail || defaultStatusDetail();
  const tone = state.latestStatus?.tone || "neutral";
  elements.helperRow.hidden = !(hasVisibleStatus || state.serverMode || state.outputDir || !state.capabilities.nativeOutputPicker);
  if (elements.helperRow.hidden) {
    return;
  }
  elements.simpleStatus.className = `simple-status is-${tone}`;
  elements.statusHeadline.textContent = headline;
  elements.statusDetail.textContent = detail;
}

function renderExplainer() {
  if (!elements.explainerModal || !elements.explainerBackdrop || !elements.howItWorksButton) {
    return;
  }

  const isOpen = state.explainerOpen;
  elements.explainerModal.hidden = !isOpen;
  elements.explainerBackdrop.hidden = !isOpen;
  elements.howItWorksButton.setAttribute("aria-expanded", isOpen ? "true" : "false");
  document.body.classList.toggle("explainer-open", isOpen);
}

function openExplainer() {
  state.explainerOpen = true;
  renderExplainer();
  elements.explainerClose?.focus();
}

function closeExplainer() {
  state.explainerOpen = false;
  renderExplainer();
  elements.howItWorksButton?.focus();
}

function celebrateQueueCompletion() {
  if (!elements.queuePanel) {
    return;
  }
  if (state.celebratoryPulseTimer) {
    clearTimeout(state.celebratoryPulseTimer);
  }
  elements.queuePanel.classList.add("is-celebrating");
  state.celebratoryPulseTimer = window.setTimeout(() => {
    elements.queuePanel.classList.remove("is-celebrating");
    state.celebratoryPulseTimer = 0;
  }, 1400);
}

function updateActionState() {
  const queuedCount = state.items.length;
  const convertibleCount = state.items.filter((item) => item.stage !== "complete").length;
  const hasItems = queuedCount > 0;
  const busy = state.isBusy;
  const outputReady = Boolean(state.outputDir);
  const completedDuringRun = state.activeConvertIds.length
    ? state.activeConvertIds.filter((id) => {
        const item = state.items.find((candidate) => candidate.id === id);
        return Boolean(item && (item.stage === "complete" || item.stage === "failed"));
      }).length
    : 0;

  elements.queueCountBadge.textContent = `${queuedCount} queued`;
  elements.clearButton.disabled = busy || !hasItems;
  elements.convertButton.disabled = busy || !convertibleCount;

  if (busy && state.activeConvertIds.length) {
    elements.convertButton.textContent = `Converting ${completedDuringRun}/${state.activeConvertIds.length}...`;
  } else if (!convertibleCount) {
    elements.convertButton.textContent = "Convert to PDF";
  } else if (!outputReady && !state.capabilities.nativeOutputPicker) {
    elements.convertButton.textContent = "Output Folder Required";
  } else if (!outputReady) {
    elements.convertButton.textContent = "Choose Save Folder";
  } else {
    elements.convertButton.textContent = `Convert ${convertibleCount} File${convertibleCount === 1 ? "" : "s"} to PDF`;
  }
}

function renderQueue() {
  const queuedCount = state.items.length;
  elements.queueEmpty.hidden = queuedCount > 0;
  elements.queueList.hidden = queuedCount === 0;

  if (!queuedCount) {
    elements.queueEmpty.textContent = "No emails queued yet.";
    elements.queueList.innerHTML = "";
    updateActionState();
    return;
  }

  elements.queueList.innerHTML = state.items
    .map((item) => {
      const progress = queueProgressStateForItem(item);
      const visual = queueVisualState(item, progress);
      const canRemove = item.stage !== "complete";
      const canRetry = item.stage === "failed";
      const isRemoving = Boolean(state.pendingRemovalsById[item.id]);
      const summary = deriveItemSummary(item, progress);

      return `
        <div class="queue-item is-${visual.tone}" style="--queue-progress: ${progress?.percent || 0}">
          <div class="queue-main">
            <span class="queue-file-glyph" aria-hidden="true">${escapeHtml(queueDocumentLabel(item))}</span>
            <div class="queue-copy">
              <div class="queue-name" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</div>
              <div class="queue-summary-line">${escapeHtml(summary)}</div>
            </div>
          </div>
          <div class="queue-actions">
            <span class="queue-state-pill is-${visual.tone}">${escapeHtml(visual.label)}</span>
            ${canRetry ? `<button class="queue-retry" data-retry-id="${escapeHtml(item.id)}" type="button">Retry</button>` : ""}
            ${
              canRemove
                ? `<button class="queue-remove ${isRemoving ? "is-pending" : ""}" data-remove-id="${escapeHtml(item.id)}" type="button" aria-busy="${isRemoving ? "true" : "false"}">${isRemoving ? "Removing..." : "Remove"}</button>`
                : `<span class="queue-complete-badge">Done</span>`
            }
          </div>
          <div class="queue-progress-track" aria-hidden="true">
            <span class="queue-progress-fill"></span>
          </div>
        </div>
      `;
    })
    .join("");

  updateActionState();
}

function applyQueueSnapshot(items) {
  const nextItems = items || [];
  const nextProgress = {};
  const idSet = new Set(nextItems.map((item) => item.id));

  nextItems.forEach((item) => {
    const existing = state.queueProgressByTaskId[item.taskId];
    if (existing?.active) {
      nextProgress[item.taskId] = existing;
      return;
    }
    const snapshotProgress = buildQueueProgressStateFromSnapshot(item);
    if (snapshotProgress) {
      nextProgress[item.taskId] = snapshotProgress;
    }
  });

  state.queueProgressByTaskId = nextProgress;
  state.pendingRemovalsById = Object.fromEntries(Object.entries(state.pendingRemovalsById).filter(([id]) => idSet.has(id)));
  state.activeConvertIds = state.activeConvertIds.filter((id) => idSet.has(id));
  state.items = nextItems;
  renderQueue();
}

function mergeQueueEvent(payload) {
  if (!payload.taskId) {
    return;
  }

  const outputPath =
    payload.outputPath ||
    (payload.meta && typeof payload.meta.outputPath === "string" ? payload.meta.outputPath : "") ||
    "";

  let changed = false;
  state.items = state.items.map((item) => {
    if (item.taskId !== payload.taskId) {
      return item;
    }
    changed = true;
    return {
      ...item,
      stage: payload.stage || item.stage,
      pipeline: payload.pipeline || item.pipeline,
      error: payload.error || item.error,
      success: typeof payload.success === "boolean" ? payload.success : item.success,
      outputPath: outputPath || item.outputPath,
    };
  });

  if (changed) {
    renderQueue();
  }
}

function describeEvent(payload) {
  const fileName = payload.fileName || "File";
  const stageLabel = queueStageLabel(payload.stage);
  if (payload.stage === "failed") {
    return {
      headline: `${fileName} needs attention.`,
      detail: payload.error || "The file could not be converted.",
      tone: "error",
    };
  }
  if (payload.stage === "complete") {
    return {
      headline: `${fileName} saved successfully.`,
      detail: payload.meta?.outputPath ? lastPathSegment(payload.meta.outputPath) : "PDF saved.",
      tone: "success",
    };
  }
  return {
    headline: `${fileName}: ${stageLabel}.`,
    detail: "Working through the queued files.",
    tone: "neutral",
  };
}

function handleBrokerMessage(payload) {
  if (payload.stage) {
    mergeQueueEvent(payload);
    updateQueueProgress(payload.taskId, payload.stage);
    const description = describeEvent(payload);
    addStatus(description.headline, description.detail, description.tone);
    return;
  }

  if (payload.message) {
    addStatus(payload.message, "", "neutral");
  }
}

async function api(path, options = {}) {
  const response = await fetch(path, options);
  const isJson = response.headers.get("content-type")?.includes("application/json");
  const payload = isJson ? await response.json() : null;
  if (!response.ok) {
    throw new Error(payload?.detail || response.statusText || "Request failed");
  }
  return payload;
}

async function loadHealth() {
  const payload = await api("/api/health");
  state.maxFiles = payload.maxFiles || state.maxFiles;
}

async function loadQueue() {
  const payload = await api("/api/queue");
  state.maxFiles = payload.maxFiles || state.maxFiles;
  applyQueueSnapshot(payload.items || []);
}

function detectUploadSourceFromDrop(dataTransfer) {
  if (!state.capabilities.outlookImport) {
    return "upload";
  }
  const types = Array.from(dataTransfer?.types || []).map((value) => String(value).toLowerCase());
  if (types.some((value) => value.includes("filegroupdescriptor") || value.includes("outlook"))) {
    return "outlook";
  }
  return "upload";
}

async function loadSettings() {
  const payload = await api("/api/settings");
  state.maxFiles = payload.maxFiles || state.maxFiles;
  state.serverMode = Boolean(payload.serverMode);
  state.capabilities = {
    nativeOutputPicker: Boolean(payload.capabilities?.nativeOutputPicker ?? true),
    outlookImport: Boolean(payload.capabilities?.outlookImport ?? true),
  };
  state.outputDir = payload.defaultOutputDir || state.outputDir;
  state.outputDirLabel = payload.defaultOutputDirLabel || state.outputDirLabel;
  refreshHostingCopy();
  setDropzoneCopy("default");
  renderStatus();
  updateActionState();
}

async function uploadFiles(files, { sourceHint = "upload" } = {}) {
  if (!files?.length) {
    return;
  }
  const formData = new FormData();
  [...files].forEach((file) => formData.append("files", file));
  formData.append("source_hint", sourceHint);
  const payload = await api("/api/upload", { method: "POST", body: formData });
  applyQueueSnapshot(payload.items || []);
  if (payload.accepted?.length) {
    addStatus(`Queued ${payload.accepted.length} file(s).`, "Click Convert to PDF when you are ready.", "success");
  }
  if (payload.rejectedCount) {
    addStatus(
      `${payload.rejectedCount} file(s) were skipped.`,
      `Only .msg files are accepted and the queue limit is ${state.maxFiles}.`,
      "error",
    );
  }
}

async function chooseOutputFolder({ silentCancel = false } = {}) {
  const payload = await api("/api/choose-output-folder", { method: "POST" });
  if (!payload.outputDir) {
    if (payload.disabled) {
      if (!silentCancel) {
        addStatus("Save folder is managed on the server.", payload.reason || defaultStatusDetail(), "neutral");
      }
      return false;
    }
    if (!silentCancel) {
      addStatus("Save folder selection was cancelled.", "Choose a folder when you are ready to convert.", "neutral");
    }
    return false;
  }
  state.outputDir = payload.outputDir;
  state.outputDirLabel = payload.outputDirLabel || "";
  refreshHostingCopy();
  addStatus("Save folder selected.", state.outputDir, "success");
  updateActionState();
  return payload;
}

async function convertQueue() {
  const convertibleItems = state.items.filter((item) => item.stage !== "complete");
  if (!convertibleItems.length) {
    addStatus("Add Outlook messages before converting.", "Drop emails into the window or browse for .msg files.", "error");
    return;
  }

  if (!state.outputDir) {
    const selected = await chooseOutputFolder({ silentCancel: false });
    if (!selected) {
      addStatus("Choose a save folder to continue.", "Conversion has not started yet.", "error");
      return;
    }
  }

  const ids = convertibleItems.map((item) => item.id);
  state.activeConvertIds = ids;
  convertibleItems.forEach((item) => activateQueueProgress(item.taskId));
  renderQueue();
  startQueueProgressLoop();
  addStatus(`Starting conversion for ${ids.length} file(s).`, state.outputDir, "neutral");

  try {
    const payload = await api("/api/convert", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ids, output_dir: state.outputDir }),
    });
    await loadQueue();
    if (payload.convertedFiles?.length) {
      addStatus(`Converted ${payload.convertedFiles.length} file(s).`, "Your PDFs have been saved.", "success");
      celebrateQueueCompletion();
    }
    if (payload.errors?.length) {
      payload.errors.forEach((error) => addStatus("A file could not be converted.", error, "error"));
    }
  } finally {
    state.activeConvertIds = [];
    updateActionState();
  }
}

async function retryItem(id) {
  const target = state.items.find((item) => item.id === id);
  if (!target || target.stage !== "failed") {
    return;
  }

  if (!state.outputDir) {
    const selected = await chooseOutputFolder({ silentCancel: false });
    if (!selected) {
      addStatus("Choose a save folder before retrying.", "Retry has not started yet.", "error");
      return;
    }
  }

  state.activeConvertIds = [id];
  activateQueueProgress(target.taskId);
  renderQueue();
  startQueueProgressLoop();
  addStatus(`Retrying ${target.name}.`, state.outputDir, "neutral");

  try {
    const payload = await api("/api/convert", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ids: [id], output_dir: state.outputDir }),
    });
    await loadQueue();
    if (payload.convertedFiles?.length) {
      addStatus("Retry succeeded.", `${target.name} was converted successfully.`, "success");
      celebrateQueueCompletion();
    }
    if (payload.errors?.length) {
      payload.errors.forEach((error) => addStatus("Retry failed.", error, "error"));
    }
  } finally {
    state.activeConvertIds = [];
    updateActionState();
  }
}

async function removeItem(id) {
  const payload = await api("/api/remove", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ids: [id] }),
  });
  applyQueueSnapshot(payload.items || []);
  addStatus("Removed a file from the queue.", "", "neutral");
}

async function clearQueue() {
  await api("/api/clear", { method: "POST" });
  applyQueueSnapshot([]);
  addStatus("Cleared the queue.", "Drop more emails to keep going.", "neutral");
}

async function runBusy(task) {
  if (state.isBusy) {
    return;
  }
  state.isBusy = true;
  elements.shell.dataset.busy = "true";
  updateActionState();
  try {
    await task();
  } finally {
    state.isBusy = false;
    elements.shell.dataset.busy = "false";
    updateActionState();
  }
}

function installQueueEvents() {
  elements.queueList.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const retryButton = target.closest("[data-retry-id]");
    if (retryButton instanceof HTMLButtonElement) {
      if (state.isBusy) {
        addStatus("Please wait for the current conversion to finish.", "Then try again.", "neutral");
        return;
      }
      try {
        await runBusy(() => retryItem(retryButton.dataset.retryId || ""));
      } catch (error) {
        addStatus("Retry failed.", error.message, "error");
      }
      return;
    }

    const removeButton = target.closest("[data-remove-id]");
    if (removeButton instanceof HTMLButtonElement) {
      if (state.isBusy) {
        addStatus("Please wait for the current conversion to finish.", "Then remove files if needed.", "neutral");
        return;
      }
      const removeId = removeButton.dataset.removeId || "";
      if (!removeId || state.pendingRemovalsById[removeId]) {
        return;
      }
      try {
        state.pendingRemovalsById = {
          ...state.pendingRemovalsById,
          [removeId]: true,
        };
        renderQueue();
        await removeItem(removeId);
      } catch (error) {
        addStatus("Could not remove that file.", error.message, "error");
      } finally {
        const nextPending = { ...state.pendingRemovalsById };
        delete nextPending[removeId];
        state.pendingRemovalsById = nextPending;
        renderQueue();
      }
    }
  });
}

function setDropzoneCopy(mode) {
  if (!elements.dropzoneCopy) {
    return;
  }
  if (!state.capabilities.outlookImport) {
    elements.dropzoneCopy.textContent = SERVER_DROP_COPY;
    return;
  }
  if (mode === "outlook") {
    elements.dropzoneCopy.textContent = "Drop Outlook emails here.";
    return;
  }
  if (mode === "upload") {
    elements.dropzoneCopy.textContent = "Drop .msg files here.";
    return;
  }
  elements.dropzoneCopy.textContent = DEFAULT_DROP_COPY;
}

function installDropzoneEvents() {
  const prevent = (event) => {
    event.preventDefault();
    event.stopPropagation();
  };

  ["dragenter", "dragover"].forEach((eventName) => {
    elements.dropzone.addEventListener(eventName, (event) => {
      prevent(event);
      elements.dropzone.classList.add("is-dragover");
      const sourceHint = detectUploadSourceFromDrop(event.dataTransfer);
      setDropzoneCopy(sourceHint === "outlook" ? "outlook" : "upload");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    elements.dropzone.addEventListener(eventName, (event) => {
      prevent(event);
      elements.dropzone.classList.remove("is-dragover");
      setDropzoneCopy("default");
    });
  });

  elements.dropzone.addEventListener("drop", async (event) => {
    try {
      const sourceHint = detectUploadSourceFromDrop(event.dataTransfer);
      await runBusy(() => uploadFiles(event.dataTransfer?.files, { sourceHint }));
    } catch (error) {
      addStatus("Upload failed.", error.message, "error");
    }
  });
}

function installExplainerEvents() {
  elements.howItWorksButton?.addEventListener("click", () => {
    if (state.explainerOpen) {
      closeExplainer();
      return;
    }
    openExplainer();
  });

  elements.explainerClose?.addEventListener("click", () => {
    closeExplainer();
  });

  elements.explainerBackdrop?.addEventListener("click", () => {
    closeExplainer();
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && state.explainerOpen) {
      closeExplainer();
    }
  });
}

function connectEvents() {
  const eventSource = new EventSource("/api/events");
  eventSource.onmessage = (event) => {
    if (!event.data) {
      return;
    }
    try {
      handleBrokerMessage(JSON.parse(event.data));
    } catch (error) {
      addStatus("Could not read a live update.", error.message, "error");
    }
  };
}

function installActionEvents() {
  elements.fileInput.addEventListener("change", async () => {
    try {
      await runBusy(() => uploadFiles(elements.fileInput.files));
      elements.fileInput.value = "";
    } catch (error) {
      addStatus("Upload failed.", error.message, "error");
    }
  });

  elements.convertButton.addEventListener("click", async () => {
    try {
      await runBusy(convertQueue);
    } catch (error) {
      addStatus("Conversion failed.", error.message, "error");
    }
  });

  elements.clearButton.addEventListener("click", async () => {
    try {
      await runBusy(clearQueue);
    } catch (error) {
      addStatus("Could not clear the queue.", error.message, "error");
    }
  });
}

async function bootstrap() {
  renderStatus();
  renderExplainer();
  setDropzoneCopy("default");
  renderQueue();
  installQueueEvents();
  installDropzoneEvents();
  installExplainerEvents();
  installActionEvents();
  connectEvents();
  await loadSettings();
  await loadHealth();
  await loadQueue();
}

bootstrap().catch((error) => {
  addStatus("Could not start the browser UI.", error.message, "error");
});
