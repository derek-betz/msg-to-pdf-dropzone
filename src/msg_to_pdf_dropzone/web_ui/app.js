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
    summary: "Waiting for a batch. Drop files to drive the live sequence, or run the preview to inspect the animation pass.",
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
    label: "Drop received",
    status: "Intake",
    summary: "The batch has landed in intake. The clerk acknowledges the message and begins the handoff.",
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
    label: "Files accepted",
    status: "Accepted",
    summary: "The queue is locked in. The clerk moves the accepted message toward the work surface.",
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
    label: "Outlook import",
    status: "Importing",
    summary: "A Classic Outlook selection is being extracted into the local batch queue before conversion begins.",
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
    label: "Output chosen",
    status: "Ready",
    summary: "The destination folder is set. The mailroom now knows where the finished PDFs should be delivered.",
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
    label: "Parsing message",
    status: "Parsing",
    summary: "The message is being read and normalized into email record data before any PDF work starts.",
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
    label: "Filename built",
    status: "Stamped",
    summary: "Thread-aware naming is complete. The batch now has the target PDF output name for this message.",
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
    label: "Press loading",
    status: "Loading press",
    summary: "The mailroom feeds the message into the PDF pipeline and prepares the active renderer.",
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
    label: "Pipeline selected",
    status: "Press running",
    summary: "The press is running the selected renderer. This is the core conversion pass that turns the message into a PDF.",
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
    label: "PDF written",
    status: "PDF ready",
    summary: "The press has produced the PDF. The payload switches from MSG to PDF before final delivery.",
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
    label: "Delivering",
    status: "Delivering",
    summary: "The converted PDF is being moved into the destination folder selected for this batch.",
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
    label: "Delivered",
    status: "Complete",
    summary: "The PDF has been delivered successfully. The mailroom stays green until the next file arrives.",
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
    label: "Failed",
    status: "Rejected",
    summary: "The file hit a conversion failure and was routed to reject instead of the output slot.",
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
  selected: new Set(),
  outputDir: "",
  outputDirLabel: "",
  mailroom: {
    stage: "idle",
    fileName: "No batch loaded",
    summary: STAGE_CONFIG.idle.summary,
    pipeline: "outlook_edge",
  },
  statuses: [],
};

const elements = {
  addFilesButton: document.getElementById("add-files-button"),
  chooseOutputButton: document.getElementById("choose-output-button"),
  clearButton: document.getElementById("clear-button"),
  connectionStatus: document.getElementById("connection-status"),
  convertButton: document.getElementById("convert-button"),
  dropzone: document.getElementById("dropzone"),
  fileInput: document.getElementById("file-input"),
  importOutlookButton: document.getElementById("import-outlook-button"),
  mailroomFile: document.getElementById("mailroom-file"),
  mailroomScene: document.getElementById("mailroom-scene"),
  mailroomStage: document.getElementById("mailroom-stage"),
  mailroomSummary: document.getElementById("mailroom-summary"),
  maxFilesLabel: document.getElementById("max-files-label"),
  outputFolderChip: document.getElementById("output-folder-chip"),
  outputFolderLabel: document.getElementById("output-folder-label"),
  pipelinePill: document.getElementById("pipeline-pill"),
  previewButton: document.getElementById("preview-button"),
  previewPipeline: document.getElementById("preview-pipeline"),
  queueBody: document.getElementById("queue-body"),
  queueCountBadge: document.getElementById("queue-count-badge"),
  queueEmpty: document.getElementById("queue-empty"),
  queueSummary: document.getElementById("queue-summary"),
  queueTableWrap: document.getElementById("queue-table-wrap"),
  removeButton: document.getElementById("remove-button"),
  sceneDoc: document.getElementById("scene-doc"),
  sceneStatus: document.getElementById("scene-status"),
  selectAll: document.getElementById("select-all"),
  selectionSummary: document.getElementById("selection-summary"),
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

function addStatus(message, { tone = "info", meta = "" } = {}) {
  state.statuses.unshift({
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    message,
    meta,
    tone,
  });
  state.statuses = state.statuses.slice(0, 10);
  renderStatusLog();
}

function renderStatusLog() {
  if (!state.statuses.length) {
    elements.statusLog.innerHTML = '<div class="status-line"><strong>Browser UI ready.</strong><span class="status-line-meta">Waiting for queue activity.</span></div>';
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

function activeIds() {
  return state.selected.size ? [...state.selected] : state.items.map((item) => item.id);
}

function setOutputFolder(outputDir, outputDirLabel) {
  state.outputDir = outputDir || "";
  state.outputDirLabel = outputDirLabel || "";
  elements.outputFolderLabel.textContent = state.outputDir || "Choose a folder before converting.";
  elements.outputFolderChip.textContent = state.outputDir ? state.outputDirLabel || "Selected" : "Not selected";
  updateButtonState();
}

function updateButtonState() {
  const hasItems = state.items.length > 0;
  elements.removeButton.disabled = state.selected.size === 0;
  elements.clearButton.disabled = !hasItems;
  elements.convertButton.disabled = !hasItems || !state.outputDir;
  elements.selectAll.disabled = !hasItems;
}

function renderQueue() {
  const queuedCount = state.items.length;
  const selectedCount = state.selected.size;
  elements.queueCountBadge.textContent = `${queuedCount} queued`;
  elements.queueSummary.textContent = queuedCount ? `${queuedCount} file${queuedCount === 1 ? "" : "s"} staged locally.` : "No files queued yet.";
  elements.selectionSummary.textContent = `${selectedCount} selected`;
  elements.selectAll.checked = queuedCount > 0 && selectedCount === queuedCount;
  elements.queueEmpty.hidden = queuedCount > 0;
  elements.queueTableWrap.hidden = queuedCount === 0;

  elements.queueBody.innerHTML = state.items
    .map((item) => {
      const isSelected = state.selected.has(item.id);
      return `
        <tr class="${isSelected ? "is-selected" : ""}">
          <td><input class="row-check" data-id="${escapeHtml(item.id)}" type="checkbox" ${isSelected ? "checked" : ""}></td>
          <td><div class="queue-name">${escapeHtml(item.name)}</div></td>
          <td><span class="source-pill">${escapeHtml(formatSource(item.source))}</span></td>
          <td>${escapeHtml(formatTimestamp(item.createdAt))}</td>
          <td>${escapeHtml(formatBytes(item.sizeBytes))}</td>
        </tr>
      `;
    })
    .join("");
  updateButtonState();
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
  elements.pipelinePill.textContent = stage === "idle" ? "Idle" : `${pipelineLabel(state.mailroom.pipeline)} pipeline`;
  renderTimeline(stage);
}

function applyQueueSnapshot(items, { preserveSelection = true } = {}) {
  state.items = items || [];
  const validIds = new Set(state.items.map((item) => item.id));
  state.selected = preserveSelection ? new Set([...state.selected].filter((id) => validIds.has(id))) : new Set();
  renderQueue();
}

function describeEvent(event) {
  const stageLabel = (STAGE_CONFIG[event.stage] || STAGE_CONFIG.idle).label;
  const fileName = event.fileName || "Current batch";
  const pipeline = event.pipeline ? ` via ${pipelineLabel(event.pipeline)}` : "";
  if (event.stage === "failed") {
    return { message: `${fileName} failed during ${stageLabel.toLowerCase()}.`, tone: "error", meta: event.error || "Conversion error" };
  }
  if (event.stage === "complete") {
    return { message: `${fileName} delivered successfully${pipeline}.`, tone: "success", meta: stageLabel };
  }
  return { message: `${fileName}: ${stageLabel}${pipeline}.`, tone: "preview", meta: event.timestamp ? `Event at ${formatTimestamp(event.timestamp)}` : "Live event" };
}

function handleBrokerMessage(payload) {
  if (payload.stage) {
    state.mailroom.stage = payload.stage;
    state.mailroom.fileName = payload.fileName || state.mailroom.fileName;
    state.mailroom.pipeline = payload.pipeline || state.mailroom.pipeline;
    state.mailroom.summary = payload.error ? `${(STAGE_CONFIG[payload.stage] || STAGE_CONFIG.idle).summary} ${payload.error}` : (STAGE_CONFIG[payload.stage] || STAGE_CONFIG.idle).summary;
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
  elements.maxFilesLabel.textContent = `${state.maxFiles} files`;
}

async function loadQueue() {
  const payload = await api("/api/queue");
  state.maxFiles = payload.maxFiles || state.maxFiles;
  elements.maxFilesLabel.textContent = `${state.maxFiles} files`;
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
    state.selected = new Set(payload.accepted.map((item) => item.id));
    renderQueue();
    addStatus(`Queued ${payload.accepted.length} file(s).`, { tone: "success", meta: "Browser upload" });
  }
  if (payload.rejectedCount) {
    addStatus(`${payload.rejectedCount} file(s) were skipped.`, { tone: "error", meta: "Only .msg files are accepted and the queue is capped." });
  }
}

async function importOutlookSelection() {
  const payload = await api("/api/import-outlook", { method: "POST" });
  applyQueueSnapshot(payload.items || []);
  if (payload.accepted?.length) {
    state.selected = new Set(payload.accepted.map((item) => item.id));
    renderQueue();
    addStatus(`Imported ${payload.accepted.length} Outlook message(s).`, { tone: "success", meta: "Classic Outlook selection" });
    return;
  }
  addStatus("No Outlook messages were imported.", { tone: "error", meta: "Make sure a Classic Outlook selection is available." });
}

async function chooseOutputFolder() {
  const payload = await api("/api/choose-output-folder", { method: "POST" });
  if (!payload.outputDir) {
    addStatus("Output folder selection was cancelled.", { tone: "info", meta: "Folder chooser" });
    return;
  }
  setOutputFolder(payload.outputDir, payload.outputDirLabel);
  addStatus("Output folder selected.", { tone: "success", meta: payload.outputDir });
}

async function convertSelection() {
  if (!state.outputDir) {
    addStatus("Choose an output folder before converting.", { tone: "error", meta: "Conversion blocked" });
    return;
  }
  const ids = activeIds();
  if (!ids.length) {
    addStatus("Queue is empty.", { tone: "error", meta: "Nothing to convert" });
    return;
  }
  addStatus(`Starting conversion for ${ids.length} file(s).`, { tone: "preview", meta: state.outputDir });
  const payload = await api("/api/convert", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ids, output_dir: state.outputDir }),
  });
  await loadQueue();
  if (payload.convertedFiles?.length) {
    addStatus(`Converted ${payload.convertedFiles.length} file(s).`, { tone: "success", meta: payload.convertedFiles.join(" | ") });
  }
  if (payload.errors?.length) {
    payload.errors.forEach((error) => addStatus(error, { tone: "error", meta: "Conversion error" }));
  }
}

async function removeSelection() {
  if (!state.selected.size) {
    return;
  }
  const payload = await api("/api/remove", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ids: [...state.selected] }),
  });
  state.selected.clear();
  applyQueueSnapshot(payload.items || [], { preserveSelection: false });
  addStatus("Removed the selected files from the queue.", { tone: "info", meta: "Queue updated" });
}

async function clearQueue() {
  await api("/api/clear", { method: "POST" });
  state.selected.clear();
  applyQueueSnapshot([], { preserveSelection: false });
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

function installQueueEvents() {
  elements.queueBody.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || !target.classList.contains("row-check")) {
      return;
    }
    if (target.checked) {
      state.selected.add(target.dataset.id || "");
    } else {
      state.selected.delete(target.dataset.id || "");
    }
    renderQueue();
  });

  elements.selectAll.addEventListener("change", () => {
    state.selected = elements.selectAll.checked ? new Set(state.items.map((item) => item.id)) : new Set();
    renderQueue();
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
      await uploadFiles(event.dataTransfer?.files);
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Upload failed" });
    }
  });
}

function connectEvents() {
  const eventSource = new EventSource("/api/events");
  eventSource.onopen = () => {
    elements.connectionStatus.textContent = "Live";
  };
  eventSource.onerror = () => {
    elements.connectionStatus.textContent = "Reconnecting...";
  };
  eventSource.onmessage = (event) => {
    if (!event.data) {
      return;
    }
    try {
      handleBrokerMessage(JSON.parse(event.data));
      elements.connectionStatus.textContent = "Live";
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Event parsing failed" });
    }
  };
}

function installActionEvents() {
  elements.addFilesButton.addEventListener("click", () => elements.fileInput.click());
  elements.fileInput.addEventListener("change", async () => {
    try {
      await uploadFiles(elements.fileInput.files);
      elements.fileInput.value = "";
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Upload failed" });
    }
  });
  elements.importOutlookButton.addEventListener("click", async () => {
    try {
      await importOutlookSelection();
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Outlook import failed" });
    }
  });
  elements.chooseOutputButton.addEventListener("click", async () => {
    try {
      await chooseOutputFolder();
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Folder chooser failed" });
    }
  });
  elements.convertButton.addEventListener("click", async () => {
    try {
      await convertSelection();
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Conversion failed" });
    }
  });
  elements.removeButton.addEventListener("click", async () => {
    try {
      await removeSelection();
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Remove failed" });
    }
  });
  elements.clearButton.addEventListener("click", async () => {
    try {
      await clearQueue();
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Clear failed" });
    }
  });
  elements.previewButton.addEventListener("click", async () => {
    try {
      await previewMailroom();
    } catch (error) {
      addStatus(error.message, { tone: "error", meta: "Preview failed" });
    }
  });
}

async function bootstrap() {
  renderMailroom();
  renderStatusLog();
  installQueueEvents();
  installDropzoneEvents();
  installActionEvents();
  connectEvents();
  await loadHealth();
  await loadQueue();
  setOutputFolder("", "");
  addStatus("Browser workspace ready.", { tone: "info", meta: "Queue, preview, and live event bridge are online." });
}

bootstrap().catch((error) => {
  addStatus(error.message, { tone: "error", meta: "Failed to initialize browser UI" });
});
