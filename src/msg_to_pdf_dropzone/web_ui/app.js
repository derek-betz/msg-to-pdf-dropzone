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

const QUEUE_FILTERS = ["all", "waiting", "converting", "failed", "converted"];
const DEFAULT_DROP_COPY = "Supports Outlook drags and .msg files. Click to browse manually if needed.";

const state = {
  maxFiles: 25,
  items: [],
  queueProgressByTaskId: {},
  pendingRemovalsById: {},
  expandedQueueItemsById: {},
  queueFilter: "all",
  outputDir: "",
  activeConvertIds: [],
  celebratoryPulseTimer: 0,
  statuses: [],
  statusLogExpanded: false,
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
  dropzoneCopy: document.getElementById("dropzone-copy"),
  fileInput: document.getElementById("file-input"),
  queueCountBadge: document.getElementById("queue-count-badge"),
  queueEmpty: document.getElementById("queue-empty"),
  queueFilters: document.getElementById("queue-filters"),
  queueList: document.getElementById("queue-list"),
  queuePanel: document.querySelector(".queue-panel"),
  queueSummary: document.getElementById("queue-summary"),
  shell: document.querySelector(".shell"),
  statusPanel: document.getElementById("status-panel"),
  statusLog: document.getElementById("status-log"),
  statusToggle: document.getElementById("status-toggle"),
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

function queueVisualState(item, progress) {
  const stage = item?.stage || "";
  if (stage === "complete" || progress?.tone === "complete") {
    return {
      tone: "complete",
      icon: "v",
      iconLabel: "Conversion complete",
    };
  }
  if (stage === "failed" || progress?.tone === "failed") {
    return {
      tone: "failed",
      icon: "x",
      iconLabel: "Conversion failed",
    };
  }
  if (progress?.active) {
    return {
      tone: "active",
      icon: "~",
      iconLabel: "Converting",
    };
  }
  return {
    tone: "queued",
    icon: ".",
    iconLabel: "Queued and waiting",
  };
}

function isQueueItemTerminal(item) {
  return item.stage === "complete" || item.stage === "failed";
}

function isQueueItemConverting(item) {
  const progress = queueProgressStateForItem(item);
  return Boolean(progress?.active && progress.tone === "active");
}

function matchesQueueFilter(item, filter) {
  if (filter === "all") {
    return true;
  }
  if (filter === "converted") {
    return item.stage === "complete";
  }
  if (filter === "failed") {
    return item.stage === "failed";
  }
  if (filter === "converting") {
    return isQueueItemConverting(item);
  }
  if (filter === "waiting") {
    return !isQueueItemTerminal(item) && !isQueueItemConverting(item);
  }
  return true;
}

function filterQueueItems(items) {
  return items.filter((item) => matchesQueueFilter(item, state.queueFilter));
}

function queueFilterLabel(filter) {
  if (filter === "all") {
    return "All";
  }
  if (filter === "waiting") {
    return "Waiting";
  }
  if (filter === "converting") {
    return "Converting";
  }
  if (filter === "failed") {
    return "Failed";
  }
  return "Converted";
}

function deriveOutputNamePreview(item) {
  if (item.outputPath) {
    const parts = String(item.outputPath).split(/[/\\]/).filter(Boolean);
    return parts[parts.length - 1] || item.outputPath;
  }
  return "Generated when conversion runs";
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
        </div>
      `;
    })
    .join("");
}

function setStatusLogExpanded(expanded) {
  if (!elements.statusPanel || !elements.statusToggle) {
    return;
  }
  state.statusLogExpanded = Boolean(expanded);
  elements.shell.classList.toggle("layout-log-collapsed", !state.statusLogExpanded);
  elements.statusPanel.classList.toggle("is-collapsed", !state.statusLogExpanded);
  elements.statusToggle.textContent = state.statusLogExpanded ? "Hide Activity" : "Show Activity";
  elements.statusToggle.setAttribute("aria-expanded", state.statusLogExpanded ? "true" : "false");
}

function setBusy(isBusy) {
  state.isBusy = isBusy;
  elements.shell.dataset.busy = isBusy ? "true" : "false";
  updateActionState();
}

function renderQueueFilters() {
  if (!elements.queueFilters) {
    return;
  }
  const counts = {
    all: state.items.length,
    waiting: state.items.filter((item) => matchesQueueFilter(item, "waiting")).length,
    converting: state.items.filter((item) => matchesQueueFilter(item, "converting")).length,
    failed: state.items.filter((item) => matchesQueueFilter(item, "failed")).length,
    converted: state.items.filter((item) => matchesQueueFilter(item, "converted")).length,
  };
  elements.queueFilters.innerHTML = QUEUE_FILTERS.map((filter) => {
    const selected = filter === state.queueFilter;
    return `
      <button class="queue-filter-chip ${selected ? "is-selected" : ""}" type="button" data-filter="${filter}" role="tab" aria-selected="${selected ? "true" : "false"}">
        ${queueFilterLabel(filter)} <span>${counts[filter]}</span>
      </button>
    `;
  }).join("");
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
  const hasConvertibleItems = convertibleCount > 0;
  const busy = state.isBusy;
  const outputReady = Boolean(state.outputDir);
  const completedDuringRun = state.activeConvertIds.length
    ? state.activeConvertIds.filter((id) => {
        const item = state.items.find((candidate) => candidate.id === id);
        return Boolean(item && isQueueItemTerminal(item));
      }).length
    : 0;

  elements.clearButton.disabled = busy || !hasItems;
  elements.convertButton.disabled = busy || !hasConvertibleItems;
  if (busy && state.activeConvertIds.length) {
    elements.convertButton.textContent = `Converting ${completedDuringRun}/${state.activeConvertIds.length}...`;
  } else if (!hasConvertibleItems) {
    elements.convertButton.textContent = "Convert to PDF";
  } else if (!outputReady) {
    elements.convertButton.textContent = "Choose Save Folder";
  } else {
    elements.convertButton.textContent = `Convert ${convertibleCount} File${convertibleCount === 1 ? "" : "s"} to PDF`;
  }
  renderQueueFilters();
}

function renderQueue() {
  const filteredItems = filterQueueItems(state.items);
  const queuedCount = filteredItems.length;
  elements.queueCountBadge.textContent = `${state.items.length} queued`;

  elements.queueEmpty.hidden = queuedCount > 0;
  elements.queueList.hidden = queuedCount === 0;

  if (!queuedCount) {
    elements.queueEmpty.textContent = state.items.length ? "No files match this filter." : "Add Outlook messages to start.";
    elements.queueList.innerHTML = "";
    updateActionState();
    return;
  }
  elements.queueEmpty.textContent = "Add Outlook messages to start.";

  elements.queueList.innerHTML = filteredItems
    .map(
      (item) => {
        const progress = queueProgressStateForItem(item);
        const canRemove = item.stage !== "complete";
        const canRetry = item.stage === "failed";
        const isExpanded = Boolean(state.expandedQueueItemsById[item.id]);
        const isRemoving = Boolean(state.pendingRemovalsById[item.id]);
        const hasProgress = Boolean(progress?.active);
        const progressLabel = progress?.label || queueStageLabel(item.stage);
        const progressPercent = Math.round(progress?.percent || 0);
        const progressMetric = hasProgress ? `${progressPercent}%` : "Waiting";
        const visual = queueVisualState(item, progress);
        const classes = ["queue-item", `is-${visual.tone}`];
        return `
        <div class="${classes.join(" ")}" style="--queue-progress: ${hasProgress ? progress.percent : 0}">
          <div class="queue-main">
            <span class="queue-file-glyph" aria-hidden="true">@</span>
            <div class="queue-copy">
              <div class="queue-name" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</div>
              <div class="queue-stage-row">
                <span class="queue-stage-label">${escapeHtml(progressLabel)}</span>
                <span class="queue-stage-percent">${escapeHtml(progressMetric)}</span>
              </div>
            </div>
          </div>
          <div class="queue-actions">
            <span class="queue-state-icon is-${visual.tone}" title="${escapeHtml(visual.iconLabel)}" aria-label="${escapeHtml(visual.iconLabel)}">${escapeHtml(visual.icon)}</span>
            <button class="queue-details-toggle ${isExpanded ? "is-open" : ""}" data-details-id="${escapeHtml(item.id)}" type="button" title="Toggle details">${isExpanded ? "Hide" : "Details"}</button>
            ${canRetry ? `<button class="queue-retry" data-retry-id="${escapeHtml(item.id)}" type="button" title="Retry this file">Retry</button>` : ""}
            ${
              canRemove
                ? `<button class="queue-remove ${isRemoving ? "is-pending" : ""}" data-remove-id="${escapeHtml(item.id)}" type="button" title="Remove from queue" aria-busy="${isRemoving ? "true" : "false"}">${isRemoving ? "Removing..." : "Remove"}</button>`
                : `<span class="queue-complete-badge" title="Already converted">Converted</span>`
            }
          </div>
          <div class="queue-details ${isExpanded ? "is-visible" : ""}" ${isExpanded ? "" : "hidden"}>
            <span><strong>Pipeline:</strong> ${escapeHtml(item.pipeline ? pipelineLabel(item.pipeline) : "Pending")}</span>
            <span><strong>Output:</strong> ${escapeHtml(deriveOutputNamePreview(item))}</span>
            ${item.stage === "failed" ? `<span class="queue-error-line"><strong>Issue:</strong> ${escapeHtml(item.error || "No failure details were returned.")}</span>` : ""}
          </div>
          <div class="queue-progress-track" aria-hidden="true">
            <span class="queue-progress-fill"></span>
          </div>
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
  state.expandedQueueItemsById = Object.fromEntries(Object.entries(state.expandedQueueItemsById).filter(([id]) => idSet.has(id)));
  state.activeConvertIds = state.activeConvertIds.filter((id) => idSet.has(id));
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

function detectUploadSourceFromDrop(dataTransfer) {
  const types = Array.from(dataTransfer?.types || []).map((value) => String(value).toLowerCase());
  if (types.some((value) => value.includes("filegroupdescriptor") || value.includes("outlook"))) {
    return "outlook";
  }
  return "upload";
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
  state.outputDir = payload.outputDir;
  addStatus("Save folder selected.", { tone: "success", meta: payload.outputDir });
  updateActionState();
  return payload;
}

async function convertQueue() {
  const convertibleItems = state.items.filter((item) => item.stage !== "complete");
  if (!convertibleItems.length) {
    addStatus("Add Outlook messages before converting.", { tone: "error", meta: "Queue is empty" });
    return;
  }

  if (!state.outputDir) {
    const selected = await chooseOutputFolder({ silentCancel: false });
    if (!selected) {
      addStatus("Choose a save folder to continue.", { tone: "error", meta: "Conversion paused" });
      return;
    }
  }

  const ids = convertibleItems.map((item) => item.id);
  state.activeConvertIds = ids;
  convertibleItems.forEach((item) => activateQueueProgress(item.taskId));
  renderQueue();
  startQueueProgressLoop();
  addStatus(`Starting conversion for ${ids.length} file(s).`, { tone: "preview", meta: state.outputDir });

  try {
    const payload = await api("/api/convert", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ids, output_dir: state.outputDir }),
    });
    await loadQueue();
    if (payload.convertedFiles?.length) {
      addStatus(`Converted ${payload.convertedFiles.length} file(s).`, { tone: "success", meta: "PDF batch complete" });
      celebrateQueueCompletion();
    }
    if (payload.errors?.length) {
      payload.errors.forEach((error) => addStatus(error, { tone: "error", meta: "Conversion error" }));
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
      addStatus("Choose a save folder before retrying.", { tone: "error", meta: "Retry paused" });
      return;
    }
  }
  state.activeConvertIds = [id];
  activateQueueProgress(target.taskId);
  renderQueue();
  startQueueProgressLoop();
  addStatus(`Retrying ${target.name}.`, { tone: "preview", meta: state.outputDir });
  try {
    const payload = await api("/api/convert", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ids: [id], output_dir: state.outputDir }),
    });
    await loadQueue();
    if (payload.convertedFiles?.length) {
      addStatus("Retry succeeded.", { tone: "success", meta: target.name });
      celebrateQueueCompletion();
    }
    if (payload.errors?.length) {
      payload.errors.forEach((error) => addStatus(error, { tone: "error", meta: "Retry error" }));
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
    const detailsToggle = target.closest("[data-details-id]");
    if (detailsToggle instanceof HTMLButtonElement) {
      const detailsId = detailsToggle.dataset.detailsId || "";
      if (!detailsId) {
        return;
      }
      if (state.expandedQueueItemsById[detailsId]) {
        const nextExpanded = { ...state.expandedQueueItemsById };
        delete nextExpanded[detailsId];
        state.expandedQueueItemsById = nextExpanded;
      } else {
        state.expandedQueueItemsById = {
          ...state.expandedQueueItemsById,
          [detailsId]: true,
        };
      }
      renderQueue();
      return;
    }

    const retryButton = target.closest("[data-retry-id]");
    if (retryButton instanceof HTMLButtonElement) {
      if (state.isBusy) {
        addStatus("Wait for the current conversion before retrying.", { tone: "info", meta: "Retry locked during conversion" });
        return;
      }
      try {
        await runBusy(() => retryItem(retryButton.dataset.retryId || ""));
      } catch (error) {
        addStatus(error.message, { tone: "error", meta: "Retry failed" });
      }
      return;
    }

    const removeButton = target.closest("[data-remove-id]");
    if (removeButton instanceof HTMLButtonElement) {
      if (state.isBusy) {
        addStatus("Please wait for the current conversion to finish before removing files.", { tone: "info", meta: "Queue locked during conversion" });
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
        addStatus(error.message, { tone: "error", meta: "Remove failed" });
      } finally {
        const nextPending = { ...state.pendingRemovalsById };
        delete nextPending[removeId];
        state.pendingRemovalsById = nextPending;
        renderQueue();
      }
    }
  });

  if (elements.queueFilters) {
    elements.queueFilters.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      const filterButton = target.closest("[data-filter]");
      if (!(filterButton instanceof HTMLButtonElement)) {
        return;
      }
      const filter = filterButton.dataset.filter || "all";
      if (!QUEUE_FILTERS.includes(filter)) {
        return;
      }
      state.queueFilter = filter;
      renderQueue();
    });
  }
}

function setDropzoneCopy(mode) {
  if (!elements.dropzoneCopy) {
    return;
  }
  if (mode === "outlook") {
    elements.dropzoneCopy.textContent = "Drop to queue Outlook messages from your desktop client.";
    return;
  }
  if (mode === "upload") {
    elements.dropzoneCopy.textContent = "Drop to queue .msg files from your workstation.";
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
      if (!state.outputDir) {
        await chooseOutputFolder({ silentCancel: false });
        return;
      }
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

  if (elements.statusToggle) {
    elements.statusToggle.addEventListener("click", () => {
      setStatusLogExpanded(!state.statusLogExpanded);
    });
  }
}

async function bootstrap() {
  setStatusLogExpanded(false);
  setDropzoneCopy("default");
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
