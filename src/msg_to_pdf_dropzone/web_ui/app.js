const STAGE_FLOW = [
  "drop_received",
  "parse_started",
  "filename_built",
  "pipeline_selected",
  "pdf_written",
  "complete",
];

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
    label: "Complete",
    floor: 98,
    cap: 100,
    baseRate: 1.3,
    easeRate: 0.55,
    tone: "complete",
  },
  failed: {
    label: "Failed",
    floor: 100,
    cap: 100,
    baseRate: 100,
    easeRate: 1,
    tone: "failed",
  },
};

let queueProgressAnimationFrame = 0;

const STAGE_CONFIG = {
  idle: {
    label: "Idle",
    status: "Idle",
    summary: "Waiting for files. Queue messages or run the preview to watch the mailroom cycle.",
    clerkX: "6%",
    clerkY: "152px",
    clerkRot: "0deg",
    docX: "13%",
    docY: "146px",
    docRot: "-8deg",
    docScale: "0.84",
    docAsset: "msg",
  },
  drop_received: {
    label: "File received",
    status: "Intake",
    summary: "The message file has entered the local queue and is moving into the mailroom.",
    clerkX: "6%",
    clerkY: "150px",
    clerkRot: "-2deg",
    docX: "10%",
    docY: "118px",
    docRot: "-7deg",
    docScale: "1",
    docAsset: "msg",
  },
  files_accepted: {
    label: "Queued",
    status: "Queued",
    summary: "The file is staged locally and ready for conversion when the batch starts.",
    clerkX: "14%",
    clerkY: "150px",
    clerkRot: "-2deg",
    docX: "20%",
    docY: "130px",
    docRot: "-3deg",
    docScale: "1",
    docAsset: "msg",
  },
  outlook_extract_started: {
    label: "Importing from Outlook",
    status: "Importing",
    summary: "The selected Outlook message is being copied into the local queue before conversion starts.",
    clerkX: "10%",
    clerkY: "150px",
    clerkRot: "-1deg",
    docX: "16%",
    docY: "124px",
    docRot: "-6deg",
    docScale: "1",
    docAsset: "msg",
  },
  output_folder_selected: {
    label: "Save folder confirmed",
    status: "Ready",
    summary: "The destination folder is set, so the finished PDF now has a place to go.",
    clerkX: "18%",
    clerkY: "148px",
    clerkRot: "0deg",
    docX: "24%",
    docY: "126px",
    docRot: "-2deg",
    docScale: "1",
    docAsset: "msg",
  },
  parse_started: {
    label: "Preparing message",
    status: "Preparing",
    summary: "The Outlook message is being prepared for conversion into a clean PDF.",
    clerkX: "22%",
    clerkY: "144px",
    clerkRot: "0deg",
    docX: "31%",
    docY: "120px",
    docRot: "0deg",
    docScale: "1",
    docAsset: "msg",
  },
  filename_built: {
    label: "Building filename",
    status: "Naming",
    summary: "The final PDF name is being built from the thread date and message subject.",
    clerkX: "28%",
    clerkY: "146px",
    clerkRot: "0deg",
    docX: "38%",
    docY: "118px",
    docRot: "-4deg",
    docScale: "1",
    docAsset: "msg",
  },
  pdf_pipeline_started: {
    label: "Loading tools",
    status: "Loading",
    summary: "The PDF tools are loading for this file before the final creation pass begins.",
    clerkX: "38%",
    clerkY: "148px",
    clerkRot: "1deg",
    docX: "47%",
    docY: "132px",
    docRot: "-2deg",
    docScale: "0.96",
    docAsset: "msg",
  },
  pipeline_selected: {
    label: "Creating PDF",
    status: "Creating",
    summary: "The selected conversion route is turning the Outlook message into a finished PDF.",
    clerkX: "47%",
    clerkY: "150px",
    clerkRot: "2deg",
    docX: "54%",
    docY: "142px",
    docRot: "0deg",
    docScale: "0.82",
    docAsset: "msg",
  },
  pdf_written: {
    label: "PDF created",
    status: "PDF ready",
    summary: "The PDF has been created and is about to be placed in the selected folder.",
    clerkX: "58%",
    clerkY: "150px",
    clerkRot: "2deg",
    docX: "68%",
    docY: "124px",
    docRot: "-4deg",
    docScale: "1",
    docAsset: "pdf",
  },
  deliver_started: {
    label: "Saving file",
    status: "Saving",
    summary: "The finished PDF is being delivered into the selected folder.",
    clerkX: "67%",
    clerkY: "150px",
    clerkRot: "1deg",
    docX: "79%",
    docY: "124px",
    docRot: "-2deg",
    docScale: "1",
    docAsset: "pdf",
  },
  complete: {
    label: "Saved",
    status: "Complete",
    summary: "The PDF has been saved successfully. The mailroom stays ready for the next file.",
    clerkX: "69%",
    clerkY: "152px",
    clerkRot: "0deg",
    docX: "82%",
    docY: "120px",
    docRot: "0deg",
    docScale: "1",
    docAsset: "pdf",
  },
  failed: {
    label: "Needs attention",
    status: "Failed",
    summary: "This file could not be converted, so it was routed to the reject lane instead of saved output.",
    clerkX: "72%",
    clerkY: "154px",
    clerkRot: "0deg",
    docX: "88%",
    docY: "118px",
    docRot: "-8deg",
    docScale: "1",
    docAsset: "failed",
  },
};

const ASSET_PATHS = {
  msg: "/assets/mailroom/msg-document.png",
  pdf: "/assets/mailroom/pdf-document.png",
  failed: "/assets/mailroom/failed-document.png",
};

const PIPELINE_LABELS = {
  outlook_edge: "Outlook Edge",
  edge_html: "Edge HTML",
  reportlab: "ReportLab",
};

const state = {
  maxFiles: 25,
  items: [],
  queueProgressByTaskId: {},
  statuses: [],
  isBusy: false,
  mailroom: {
    stage: "idle",
    fileName: "No batch loaded",
    summary: STAGE_CONFIG.idle.summary,
    pipeline: "outlook_edge",
  },
};

const elements = {
  clearButton: document.getElementById("clear-button"),
  connectionDot: document.getElementById("connection-dot"),
  connectionStatus: document.getElementById("connection-status"),
  convertButton: document.getElementById("convert-button"),
  dropzone: document.getElementById("dropzone"),
  fileInput: document.getElementById("file-input"),
  queueCountBadge: document.getElementById("queue-count-badge"),
  queueEmpty: document.getElementById("queue-empty"),
  queueList: document.getElementById("queue-list"),
  queueSummary: document.getElementById("queue-summary"),
  shell: document.querySelector(".shell"),
  statusLog: document.getElementById("status-log"),
};

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function pipelineLabel(pipeline) {
  return PIPELINE_LABELS[pipeline] || "Idle";
}

function formatSource(source) {
  if (source === "outlook") {
    return "Outlook";
  }
  if (source === "upload") {
    return "Browser";
  }
  return source || "Local";
}

function formatTimestamp(value) {
  if (!value) {
    return "";
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return new Intl.DateTimeFormat(undefined, {
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
  }).format(date);
}

function queueStageLabel(stage) {
  if (stage === "failed") {
    return "Failed";
  }
  return STAGE_CONFIG[stage]?.status || STAGE_CONFIG[stage]?.label || "Queued";
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
  const now = performance.now();
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
      lastFrameAt: now,
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
  const now = performance.now();
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
      lastFrameAt: now,
    },
  };
  if (state.items.some((item) => item.taskId === taskId)) {
    renderQueue();
  }
  startQueueProgressLoop();
}

function setConnectionState(mode) {
  elements.connectionDot.classList.remove("is-live", "is-offline");
  if (mode === "live") {
    elements.connectionDot.classList.add("is-live");
    elements.connectionStatus.textContent = "Live";
    return;
  }
  if (mode === "offline") {
    elements.connectionDot.classList.add("is-offline");
    elements.connectionStatus.textContent = "Reconnecting...";
    return;
  }
  elements.connectionStatus.textContent = "Connecting...";
}

function addStatus(message, { tone = "info", meta = "" } = {}) {
  state.statuses.unshift({
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    message,
    meta,
    tone,
  });
  state.statuses = state.statuses.slice(0, 8);
  renderStatusLog();
}

function renderStatusLog() {
  if (!state.statuses.length) {
    elements.statusLog.innerHTML = `
      <div class="status-line">
        <strong>Ready for intake.</strong>
        <span class="status-line-meta">Drop emails here or click the intake field to browse for .msg files.</span>
      </div>
    `;
    return;
  }

  elements.statusLog.innerHTML = state.statuses
    .map((entry) => {
      const toneClass =
        entry.tone === "error" ? "is-error" : entry.tone === "success" ? "is-success" : entry.tone === "preview" ? "is-preview" : "";
      return `
        <div class="status-line ${toneClass}">
          <strong>${escapeHtml(entry.message)}</strong>
          <span class="status-line-meta">${escapeHtml(entry.meta || "Local workstation session")}</span>
        </div>
      `;
    })
    .join("");
}

function setBusy(isBusy) {
  state.isBusy = isBusy;
  elements.shell.dataset.busy = isBusy ? "true" : "false";
  updateActionState();
}

function updateActionState() {
  const queuedCount = state.items.length;
  const convertibleCount = state.items.filter((item) => item.stage !== "complete").length;
  const hasItems = queuedCount > 0;
  const hasConvertibleItems = convertibleCount > 0;
  const busy = state.isBusy;

  elements.clearButton.disabled = busy || !hasItems;
  elements.convertButton.disabled = busy || !hasConvertibleItems;
  elements.convertButton.textContent = `Convert ${convertibleCount} File${convertibleCount === 1 ? "" : "s"} to PDF`;

  if (!hasConvertibleItems) {
    elements.convertButton.textContent = "Convert to PDF";
  }
}

function renderQueue() {
  const queuedCount = state.items.length;
  elements.queueCountBadge.textContent = `${queuedCount} queued`;

  elements.queueEmpty.hidden = queuedCount > 0;
  elements.queueList.hidden = queuedCount === 0;

  if (!queuedCount) {
    elements.queueList.innerHTML = "";
    updateActionState();
    return;
  }

  elements.queueList.innerHTML = state.items
    .map(
      (item) => {
        const progress = queueProgressStateForItem(item);
        const disabled = state.isBusy ? "disabled" : "";
        const hasProgress = Boolean(progress?.active);
        const progressLabel = progress?.label || "Queued";
        const progressPercent = Math.round(progress?.percent || 0);
        const classes = ["queue-item"];
        if (hasProgress) {
          classes.push("has-progress", `is-${progress.tone}`);
        }
        return `
        <div class="${classes.join(" ")}" ${hasProgress ? `style="--queue-progress: ${progress.percent}"` : ""}>
          <div class="queue-main">
            <span class="source-pill">${escapeHtml(formatSource(item.source))}</span>
            <div class="queue-copy">
              <div class="queue-name" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</div>
              <div class="queue-stage-row" ${hasProgress ? "" : "hidden"}>
                <span class="queue-stage-label">${escapeHtml(progressLabel)}</span>
                <span class="queue-stage-percent">${progressPercent}%</span>
              </div>
            </div>
          </div>
          <button class="remove-pill" data-remove-id="${escapeHtml(item.id)}" type="button" ${disabled}>Remove</button>
        </div>
      `;
      },
    )
    .join("");

  updateActionState();
}

function renderMailroom() {
  return;
}

function applyQueueSnapshot(items) {
  const nextItems = items || [];
  const nextProgress = {};
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
  state.items = nextItems;
  renderQueue();
}

function describeEvent(event) {
  const stageLabel = (STAGE_CONFIG[event.stage] || STAGE_CONFIG.idle).label;
  const fileName = event.fileName || "Current batch";
  const pipeline = event.pipeline ? `${pipelineLabel(event.pipeline)} route` : "";
  const meta = event.timestamp ? `Event at ${formatTimestamp(event.timestamp)}` : "Live activity";

  if (event.stage === "failed") {
    return {
      message: `${fileName} needs attention.`,
      tone: "error",
      meta: event.error || "The file could not be converted.",
    };
  }

  if (event.stage === "complete") {
    return {
      message: `${fileName} saved successfully.`,
      tone: "success",
      meta: pipeline || "PDF saved",
    };
  }

  return {
    message: `${fileName}: ${stageLabel}.`,
    tone: "preview",
    meta: pipeline || meta,
  };
}

function mergeQueueEvent(payload) {
  if (!payload.taskId) {
    return;
  }
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
    };
  });
  if (changed) {
    renderQueue();
  }
}

function handleBrokerMessage(payload) {
  if (payload.stage) {
    mergeQueueEvent(payload);
    updateQueueProgress(payload.taskId, payload.stage);
    state.mailroom.stage = payload.stage;
    state.mailroom.fileName = payload.fileName || state.mailroom.fileName;
    state.mailroom.pipeline = payload.pipeline || state.mailroom.pipeline;
    state.mailroom.summary = payload.error
      ? `${(STAGE_CONFIG[payload.stage] || STAGE_CONFIG.idle).summary} ${payload.error}`
      : (STAGE_CONFIG[payload.stage] || STAGE_CONFIG.idle).summary;
    renderMailroom();
    const description = describeEvent(payload);
    addStatus(description.message, { tone: description.tone, meta: description.meta });
    return;
  }

  if (payload.message) {
    addStatus(payload.message, { tone: "info", meta: "Server status" });
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

async function uploadFiles(files) {
  if (!files?.length) {
    return;
  }
  const formData = new FormData();
  [...files].forEach((file) => formData.append("files", file));
  const payload = await api("/api/upload", { method: "POST", body: formData });
  applyQueueSnapshot(payload.items || []);
  if (payload.accepted?.length) {
    addStatus(`Queued ${payload.accepted.length} file(s).`, { tone: "success", meta: "Files added to the local queue" });
  }
  if (payload.rejectedCount) {
    addStatus(`${payload.rejectedCount} file(s) were skipped.`, {
      tone: "error",
      meta: `Only .msg files are accepted and the queue limit is ${state.maxFiles}.`,
    });
  }
}

async function chooseOutputFolder({ silentCancel = false } = {}) {
  const payload = await api("/api/choose-output-folder", { method: "POST" });
  if (!payload.outputDir) {
    if (!silentCancel) {
      addStatus("Save folder selection was cancelled.", { tone: "info", meta: "Folder chooser" });
    }
    return false;
  }
  addStatus("Save folder selected.", { tone: "success", meta: payload.outputDir });
  return payload;
}

async function convertQueue() {
  const convertibleItems = state.items.filter((item) => item.stage !== "complete");
  if (!convertibleItems.length) {
    addStatus("Add Outlook messages before converting.", { tone: "error", meta: "Queue is empty" });
    return;
  }

  const selected = await chooseOutputFolder({ silentCancel: true });
  if (!selected) {
    addStatus("Choose a save folder to continue.", { tone: "error", meta: "Conversion paused" });
    return;
  }

  const ids = convertibleItems.map((item) => item.id);
  convertibleItems.forEach((item) => activateQueueProgress(item.taskId));
  renderQueue();
  startQueueProgressLoop();
  addStatus(`Starting conversion for ${ids.length} file(s).`, { tone: "preview", meta: selected.outputDir });
  const payload = await api("/api/convert", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ids, output_dir: selected.outputDir }),
  });
  await loadQueue();
  if (payload.convertedFiles?.length) {
    addStatus(`Converted ${payload.convertedFiles.length} file(s).`, { tone: "success", meta: "PDF batch complete" });
  }
  if (payload.errors?.length) {
    payload.errors.forEach((error) => addStatus(error, { tone: "error", meta: "Conversion error" }));
  }
}

async function removeItem(id) {
  const payload = await api("/api/remove", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ids: [id] }),
  });
  applyQueueSnapshot(payload.items || []);
  addStatus("Removed a file from the queue.", { tone: "info", meta: "Queue updated" });
}

async function clearQueue() {
  await api("/api/clear", { method: "POST" });
  applyQueueSnapshot([]);
  addStatus("Cleared the queue.", { tone: "info", meta: "Queue updated" });
}

async function runBusy(task) {
  if (state.isBusy) {
    return;
  }
  setBusy(true);
  try {
    await task();
  } finally {
    setBusy(false);
  }
}

function installQueueEvents() {
  elements.queueList.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const button = target.closest("[data-remove-id]");
    if (!button) {
      return;
    }
    try {
      await runBusy(() => removeItem(button.dataset.removeId || ""));
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Remove failed" });
    }
  });
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
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    elements.dropzone.addEventListener(eventName, (event) => {
      prevent(event);
      elements.dropzone.classList.remove("is-dragover");
    });
  });

  elements.dropzone.addEventListener("drop", async (event) => {
    try {
      await runBusy(() => uploadFiles(event.dataTransfer?.files));
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Upload failed" });
    }
  });
}

function connectEvents() {
  const eventSource = new EventSource("/api/events");
  setConnectionState("connecting");

  eventSource.onopen = () => {
    setConnectionState("live");
  };

  eventSource.onerror = () => {
    setConnectionState("offline");
  };

  eventSource.onmessage = (event) => {
    if (!event.data) {
      return;
    }
    try {
      handleBrokerMessage(JSON.parse(event.data));
      setConnectionState("live");
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Event parsing failed" });
    }
  };
}

function installActionEvents() {
  elements.fileInput.addEventListener("change", async () => {
    try {
      await runBusy(() => uploadFiles(elements.fileInput.files));
      elements.fileInput.value = "";
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Upload failed" });
    }
  });

  elements.convertButton.addEventListener("click", async () => {
    try {
      await runBusy(convertQueue);
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Conversion failed" });
    }
  });

  elements.clearButton.addEventListener("click", async () => {
    try {
      await runBusy(clearQueue);
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Clear failed" });
    }
  });
}

async function bootstrap() {
  renderMailroom();
  renderStatusLog();
  renderQueue();
  installQueueEvents();
  installDropzoneEvents();
  installActionEvents();
  connectEvents();
  await loadHealth();
  await loadQueue();
  addStatus("Local workstation ready.", { tone: "info", meta: "Files stay local and completed items stay in the queue until you clear them." });
}

bootstrap().catch((error) => {
  addStatus(error.message, { tone: "error", meta: "Failed to initialize browser UI" });
});
