const STAGE_FLOW = [
  "drop_received",
  "parse_started",
  "filename_built",
  "pipeline_selected",
  "pdf_written",
  "complete",
];

const STAGE_CONFIG = {
  idle: {
    label: "Idle",
    status: "Idle",
    summary: "Waiting for files. Drop messages or run the preview to watch the mailroom work.",
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
  outputDir: "",
  outputDirLabel: "",
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
  addFilesButton: document.getElementById("add-files-button"),
  chooseOutputButton: document.getElementById("choose-output-button"),
  clearButton: document.getElementById("clear-button"),
  connectionDot: document.getElementById("connection-dot"),
  connectionStatus: document.getElementById("connection-status"),
  convertButton: document.getElementById("convert-button"),
  dropzone: document.getElementById("dropzone"),
  fileInput: document.getElementById("file-input"),
  importOutlookButton: document.getElementById("import-outlook-button"),
  mailroomFile: document.getElementById("mailroom-file"),
  mailroomScene: document.getElementById("mailroom-scene"),
  mailroomStage: document.getElementById("mailroom-stage"),
  mailroomSummary: document.getElementById("mailroom-summary"),
  nextStepCopy: document.getElementById("next-step-copy"),
  nextStepTitle: document.getElementById("next-step-title"),
  outputFolderChip: document.getElementById("output-folder-chip"),
  outputFolderLabel: document.getElementById("output-folder-label"),
  pipelinePill: document.getElementById("pipeline-pill"),
  previewButton: document.getElementById("preview-button"),
  previewPipeline: document.getElementById("preview-pipeline"),
  queueCountBadge: document.getElementById("queue-count-badge"),
  queueEmpty: document.getElementById("queue-empty"),
  queueList: document.getElementById("queue-list"),
  queueSummary: document.getElementById("queue-summary"),
  sceneDoc: document.getElementById("scene-doc"),
  sceneStatus: document.getElementById("scene-status"),
  shell: document.querySelector(".shell"),
  statusLog: document.getElementById("status-log"),
  timeline: document.getElementById("timeline"),
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

function formatBytes(sizeBytes) {
  if (!Number.isFinite(sizeBytes) || sizeBytes <= 0) {
    return "0 KB";
  }
  const units = ["B", "KB", "MB", "GB"];
  let value = sizeBytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(value >= 10 || unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
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

function formatSource(source) {
  if (source === "outlook") {
    return "Outlook";
  }
  if (source === "upload") {
    return "Browser";
  }
  return source || "Local";
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
        <strong>Ready for files.</strong>
        <span class="status-line-meta">Drop messages or import from Classic Outlook.</span>
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
          <span class="status-line-meta">${escapeHtml(entry.meta || "Local browser session")}</span>
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

function setOutputFolder(outputDir, outputDirLabel) {
  state.outputDir = outputDir || "";
  state.outputDirLabel = outputDirLabel || "";
  elements.outputFolderLabel.textContent = state.outputDir || "Choose a folder.";
  elements.outputFolderChip.textContent = state.outputDir ? state.outputDirLabel || "Folder selected" : "Folder not selected";
  elements.chooseOutputButton.textContent = state.outputDir ? "Change Save Folder" : "Choose Save Folder";
  updateActionState();
}

function updateActionState() {
  const queuedCount = state.items.length;
  const hasItems = queuedCount > 0;
  const busy = state.isBusy;

  elements.addFilesButton.disabled = busy;
  elements.importOutlookButton.disabled = busy;
  elements.chooseOutputButton.disabled = busy;
  elements.previewButton.disabled = busy;
  elements.clearButton.disabled = busy || !hasItems;
  elements.convertButton.disabled = busy || !hasItems;

  if (!hasItems) {
    elements.nextStepTitle.textContent = "Add one or more Outlook messages";
    elements.nextStepCopy.textContent = "Drop files here or import from Classic Outlook.";
    elements.convertButton.textContent = "Convert to PDF";
    return;
  }

  if (!state.outputDir) {
    elements.nextStepTitle.textContent = "Choose where the PDFs should be saved";
    elements.nextStepCopy.textContent = "The main button opens the folder picker.";
    elements.convertButton.textContent = "Choose Folder to Continue";
    return;
  }

  elements.nextStepTitle.textContent = "Convert the queued batch";
  elements.nextStepCopy.textContent = "The files are ready to convert and save.";
  elements.convertButton.textContent = `Convert ${queuedCount} File${queuedCount === 1 ? "" : "s"} to PDF`;
}

function renderQueue() {
  const queuedCount = state.items.length;
  elements.queueCountBadge.textContent = `${queuedCount} queued`;
  elements.queueSummary.textContent = queuedCount
    ? `${queuedCount} file${queuedCount === 1 ? "" : "s"} ready for conversion.`
    : "No files queued yet.";

  elements.queueEmpty.hidden = queuedCount > 0;
  elements.queueList.hidden = queuedCount === 0;

  if (!queuedCount) {
    elements.queueList.innerHTML = "";
    updateActionState();
    return;
  }

  elements.queueList.innerHTML = state.items
    .map(
      (item) => `
        <div class="queue-item">
          <div class="queue-main">
            <div class="queue-name">${escapeHtml(item.name)}</div>
            <div class="queue-meta">
              <span class="source-pill">${escapeHtml(formatSource(item.source))}</span>
              <span>${escapeHtml(formatBytes(item.sizeBytes))}</span>
              <span>Queued ${escapeHtml(formatTimestamp(item.createdAt))}</span>
            </div>
          </div>
          <button class="remove-pill" data-remove-id="${escapeHtml(item.id)}" type="button">Remove</button>
        </div>
      `,
    )
    .join("");

  updateActionState();
}

function timelineKeyForStage(stage) {
  if (["drop_received", "files_accepted", "outlook_extract_started", "output_folder_selected"].includes(stage)) {
    return "drop_received";
  }
  if (["pdf_pipeline_started", "pipeline_selected"].includes(stage)) {
    return "pipeline_selected";
  }
  if (["deliver_started", "complete"].includes(stage)) {
    return "complete";
  }
  if (stage === "failed") {
    return "pipeline_selected";
  }
  return stage;
}

function renderTimeline(stage) {
  const activeKey = timelineKeyForStage(stage);
  const activeIndex = STAGE_FLOW.indexOf(activeKey);
  elements.timeline.querySelectorAll("span").forEach((node) => {
    const index = STAGE_FLOW.indexOf(node.dataset.stage);
    node.classList.toggle("is-active", node.dataset.stage === activeKey && stage !== "failed");
    node.classList.toggle("is-complete", activeIndex > index || (stage === "failed" && index < STAGE_FLOW.indexOf("pipeline_selected")));
  });
}

function sceneGlowForStage(stage, pipeline) {
  if (stage === "failed") {
    return "rgba(212, 84, 84, 0.22)";
  }
  if (stage === "complete") {
    return "rgba(87, 196, 128, 0.22)";
  }
  if (stage === "pipeline_selected" || stage === "pdf_pipeline_started") {
    if (pipeline === "edge_html") {
      return "rgba(83, 146, 230, 0.22)";
    }
    if (pipeline === "reportlab") {
      return "rgba(228, 163, 80, 0.2)";
    }
    return "rgba(48, 197, 209, 0.2)";
  }
  return "rgba(35, 101, 171, 0.14)";
}

function renderMailroom() {
  const stage = state.mailroom.stage in STAGE_CONFIG ? state.mailroom.stage : "idle";
  const config = STAGE_CONFIG[stage];
  elements.mailroomScene.dataset.stage = stage;
  elements.mailroomScene.style.setProperty("--clerk-x", config.clerkX);
  elements.mailroomScene.style.setProperty("--clerk-y", config.clerkY);
  elements.mailroomScene.style.setProperty("--clerk-rot", config.clerkRot);
  elements.mailroomScene.style.setProperty("--doc-x", config.docX);
  elements.mailroomScene.style.setProperty("--doc-y", config.docY);
  elements.mailroomScene.style.setProperty("--doc-rot", config.docRot);
  elements.mailroomScene.style.setProperty("--doc-scale", config.docScale);
  elements.mailroomScene.style.setProperty("--press-glow", sceneGlowForStage(stage, state.mailroom.pipeline));

  elements.sceneDoc.src = ASSET_PATHS[config.docAsset];
  elements.sceneStatus.textContent = config.status;
  elements.mailroomFile.textContent = state.mailroom.fileName || "No batch loaded";
  elements.mailroomStage.textContent = config.label;
  elements.mailroomSummary.textContent = state.mailroom.summary || config.summary;
  elements.pipelinePill.textContent = stage === "idle" ? "Idle" : `${pipelineLabel(state.mailroom.pipeline)} route`;
  renderTimeline(stage);
}

function applyQueueSnapshot(items) {
  state.items = items || [];
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

function handleBrokerMessage(payload) {
  if (payload.stage) {
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

async function importOutlookSelection() {
  const payload = await api("/api/import-outlook", { method: "POST" });
  applyQueueSnapshot(payload.items || []);
  if (payload.accepted?.length) {
    addStatus(`Imported ${payload.accepted.length} Outlook message(s).`, { tone: "success", meta: "Classic Outlook selection" });
    return;
  }
  addStatus("No Outlook messages were imported.", {
    tone: "error",
    meta: "Make sure a Classic Outlook selection is available.",
  });
}

async function chooseOutputFolder({ silentCancel = false } = {}) {
  const payload = await api("/api/choose-output-folder", { method: "POST" });
  if (!payload.outputDir) {
    if (!silentCancel) {
      addStatus("Save folder selection was cancelled.", { tone: "info", meta: "Folder chooser" });
    }
    return false;
  }
  setOutputFolder(payload.outputDir, payload.outputDirLabel);
  addStatus("Save folder selected.", { tone: "success", meta: payload.outputDir });
  return true;
}

async function convertQueue() {
  if (!state.items.length) {
    addStatus("Add Outlook messages before converting.", { tone: "error", meta: "Queue is empty" });
    return;
  }

  if (!state.outputDir) {
    const selected = await chooseOutputFolder({ silentCancel: true });
    if (!selected) {
      addStatus("Choose a save folder to continue.", { tone: "error", meta: "Conversion paused" });
      return;
    }
  }

  const ids = state.items.map((item) => item.id);
  addStatus(`Starting conversion for ${ids.length} file(s).`, { tone: "preview", meta: state.outputDir });
  const payload = await api("/api/convert", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ids, output_dir: state.outputDir }),
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

async function previewMailroom() {
  const pipeline = elements.previewPipeline.value;
  await api("/api/preview-mailroom", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ pipeline }),
  });
  addStatus("Running the mailroom preview.", { tone: "preview", meta: pipelineLabel(pipeline) });
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
  elements.addFilesButton.addEventListener("click", () => elements.fileInput.click());

  elements.fileInput.addEventListener("change", async () => {
    try {
      await runBusy(() => uploadFiles(elements.fileInput.files));
      elements.fileInput.value = "";
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Upload failed" });
    }
  });

  elements.importOutlookButton.addEventListener("click", async () => {
    try {
      await runBusy(importOutlookSelection);
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Outlook import failed" });
    }
  });

  elements.chooseOutputButton.addEventListener("click", async () => {
    try {
      await runBusy(() => chooseOutputFolder());
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Folder chooser failed" });
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

  elements.previewButton.addEventListener("click", async () => {
    try {
      await runBusy(previewMailroom);
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Preview failed" });
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
  setOutputFolder("", "");
  addStatus("Browser workspace ready.", { tone: "info", meta: "Files stay local and the mailroom companion is online." });
}

bootstrap().catch((error) => {
  addStatus(error.message, { tone: "error", meta: "Failed to initialize browser UI" });
});
