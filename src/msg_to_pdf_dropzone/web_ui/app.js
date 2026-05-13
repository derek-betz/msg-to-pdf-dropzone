import { createDropzoneController } from "./dropzone_controller.js?v=filename-panel-1";

const QUEUE_STAGE_PROGRESS = {
  drop_received: {
    label: "Received",
    floor: 3,
    cap: 6,
    baseRate: 10,
    easeRate: 0.28,
    tone: "queued",
  },
  outlook_extract_started: {
    label: "Importing",
    floor: 5,
    cap: 10,
    baseRate: 8,
    easeRate: 0.22,
    tone: "queued",
  },
  files_accepted: {
    label: "Ready",
    floor: 8,
    cap: 12,
    baseRate: 7,
    easeRate: 0.2,
    tone: "queued",
  },
  output_folder_selected: {
    label: "Starting",
    floor: 8,
    cap: 18,
    baseRate: 22,
    easeRate: 0.32,
    tone: "active",
  },
  parse_started: {
    label: "Preparing",
    floor: 18,
    cap: 38,
    baseRate: 13,
    easeRate: 0.28,
    tone: "active",
  },
  filename_built: {
    label: "Naming",
    floor: 36,
    cap: 54,
    baseRate: 10,
    easeRate: 0.22,
    tone: "active",
  },
  pdf_pipeline_started: {
    label: "Loading",
    floor: 52,
    cap: 74,
    baseRate: 12,
    easeRate: 0.24,
    tone: "active",
  },
  pipeline_selected: {
    label: "Rendering PDF",
    floor: 74,
    cap: 97,
    baseRate: 17,
    easeRate: 0.34,
    tone: "active",
  },
  pdf_written: {
    label: "PDF ready",
    floor: 97,
    cap: 99,
    baseRate: 10,
    easeRate: 0.3,
    tone: "active",
  },
  deliver_started: {
    label: "Saving",
    floor: 99,
    cap: 99.4,
    baseRate: 8,
    easeRate: 0.35,
    tone: "active",
  },
  complete: {
    label: "Saved",
    floor: 100,
    cap: 100,
    baseRate: 55,
    easeRate: 0.5,
    tone: "complete",
  },
  failed: {
    label: "Failed",
    floor: 100,
    cap: 100,
    baseRate: 55,
    easeRate: 0.5,
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
const TERMINAL_STAGES = new Set(["complete", "failed"]);

const DEFAULT_DROP_COPY = "Supports Outlook drags and .msg files. Click to browse if you prefer.";
const SERVER_DROP_COPY = "Supports .msg files. Click to browse if you prefer.";
const RECENT_DESTINATIONS_KEY = "msg-to-pdf-dropzone.recent-destinations";
const FILENAME_STYLE_KEY = "msg-to-pdf-dropzone.filename-style";
const DEFAULT_FILENAME_STYLE = "date_subject";
const FILENAME_STYLE_OPTIONS = {
  date_subject: {
    detail: "Best for project correspondence that will be combined newest-first.",
    example: "2026-05-07_Project Update.pdf",
  },
  subject: {
    detail: "Best when the subject is the easiest way to recognize each PDF.",
    example: "Project Update.pdf",
  },
  sender_subject: {
    detail: "Best when the person who sent the email matters more than the date.",
    example: "Jane Smith_Project Update.pdf",
  },
  date_sender_subject: {
    detail: "Best when each PDF needs date, sender, and subject in the filename.",
    example: "2026-05-07_Jane Smith_Project Update.pdf",
  },
};
const TIMELINE_STAGES = [
  { key: "files_accepted", label: "Files queued and ready" },
  { key: "parse_started", label: "Preparing and reading messages" },
  { key: "pipeline_selected", label: "Creating the PDFs" },
  { key: "deliver_started", label: "Saving into the destination folder" },
  { key: "complete", label: "Batch complete" },
];

let queueProgressAnimationFrame = 0;
let dropzoneController = null;

const state = {
  maxFiles: 25,
  items: [],
  queueProgressByTaskId: {},
  pendingRemovalsById: {},
  outputDir: "",
  outputDirLabel: "",
  browserOutputDirectoryHandle: null,
  browserOutputDirLabel: "",
  recentDestinations: [],
  recentlyQueuedIds: new Set(),
  activeConvertIds: [],
  celebratoryPulseTimer: 0,
  dropFeedbackTimer: 0,
  queueHandoffTimer: 0,
  newRowsTimer: 0,
  isBusy: false,
  tutorialOpen: false,
  feedbackOpen: false,
  latestStatus: null,
  filenameStyle: DEFAULT_FILENAME_STYLE,
  queueDetailsExpanded: false,
  serverMode: false,
  capabilities: {
    nativeOutputPicker: true,
    outlookImport: true,
  },
};

const elements = {
  batchProgressDetail: document.getElementById("batch-progress-detail"),
  batchProgressFill: document.getElementById("batch-progress-fill"),
  batchProgressHeadline: document.getElementById("batch-progress-headline"),
  batchProgressTrack: document.getElementById("batch-progress-track"),
  batchStatusPill: document.getElementById("batch-status-pill"),
  chooseFolderButton: document.getElementById("choose-folder-button"),
  clearButton: document.getElementById("clear-button"),
  convertButton: document.getElementById("convert-button"),
  destinationInlineValue: document.getElementById("destination-inline-value"),
  dropFeedbackDetail: document.getElementById("drop-feedback-detail"),
  dropFeedbackTitle: document.getElementById("drop-feedback-title"),
  dropRippleCanvas: document.getElementById("drop-ripple-canvas"),
  dropzone: document.getElementById("dropzone"),
  dropzoneCopy: document.getElementById("dropzone-copy"),
  appEyebrow: document.getElementById("app-eyebrow"),
  fileInput: document.getElementById("file-input"),
  feedbackBackdrop: document.getElementById("feedback-backdrop"),
  feedbackButton: document.getElementById("feedback-button"),
  feedbackCategory: document.getElementById("feedback-category"),
  feedbackClose: document.getElementById("feedback-close"),
  feedbackHelpful: document.getElementById("feedback-helpful"),
  feedbackImprove: document.getElementById("feedback-improve"),
  feedbackMessage: document.getElementById("feedback-message"),
  feedbackModal: document.getElementById("feedback-modal"),
  feedbackSend: document.getElementById("feedback-send"),
  filenameStyleDetail: document.getElementById("filename-style-detail"),
  filenameStyleExampleValue: document.getElementById("filename-style-example-value"),
  filenameStyleSelect: document.getElementById("filename-style-select"),
  helperRow: document.getElementById("helper-row"),
  historyList: document.getElementById("history-list"),
  historyStrip: document.getElementById("history-strip"),
  metricComplete: document.getElementById("metric-complete"),
  metricFailed: document.getElementById("metric-failed"),
  metricQueued: document.getElementById("metric-queued"),
  metricReady: document.getElementById("metric-ready"),
  modeBannerDetail: document.getElementById("mode-banner-detail"),
  modeBannerTitle: document.getElementById("mode-banner-title"),
  openOutputFolderButton: document.getElementById("open-output-folder-button"),
  queueCountBadge: document.getElementById("queue-count-badge"),
  queueEmpty: document.getElementById("queue-empty"),
  queueList: document.getElementById("queue-list"),
  queuePanel: document.querySelector(".queue-panel"),
  readinessConversion: document.getElementById("readiness-conversion"),
  readinessConversionValue: document.getElementById("readiness-conversion-value"),
  readinessDestination: document.getElementById("readiness-destination"),
  readinessDestinationValue: document.getElementById("readiness-destination-value"),
  readinessFiles: document.getElementById("readiness-files"),
  readinessFilesValue: document.getElementById("readiness-files-value"),
  resultBanner: document.getElementById("result-banner"),
  resultDetail: document.getElementById("result-detail"),
  resultFailed: document.getElementById("result-failed"),
  resultGuidance: document.getElementById("result-guidance"),
  resultHeadline: document.getElementById("result-headline"),
  resultOpenOutputButton: document.getElementById("result-open-output-button"),
  resultReview: document.getElementById("result-review"),
  resultReviewDestination: document.getElementById("result-review-destination"),
  resultReviewList: document.getElementById("result-review-list"),
  resultRetryFailedButton: document.getElementById("result-retry-failed-button"),
  resultSaved: document.getElementById("result-saved"),
  resultStartNewButton: document.getElementById("result-start-new-button"),
  saveCard: document.querySelector(".save-card"),
  shell: document.querySelector(".shell"),
  simpleStatus: document.getElementById("simple-status"),
  statusDetail: document.getElementById("status-detail"),
  statusHeadline: document.getElementById("status-headline"),
  timelineList: document.getElementById("timeline-list"),
  topbarBody: document.getElementById("topbar-body"),
  tutorialBackdrop: document.getElementById("tutorial-backdrop"),
  tutorialButton: document.getElementById("tutorial-button"),
  tutorialClose: document.getElementById("tutorial-close"),
  tutorialModal: document.getElementById("tutorial-modal"),
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

function loadRecentDestinations() {
  try {
    const rawValue = window.localStorage.getItem(RECENT_DESTINATIONS_KEY);
    if (!rawValue) {
      return [];
    }
    const parsed = JSON.parse(rawValue);
    return Array.isArray(parsed) ? parsed.filter((value) => typeof value === "string" && value.trim()) : [];
  } catch {
    return [];
  }
}

function saveRecentDestinations() {
  try {
    window.localStorage.setItem(RECENT_DESTINATIONS_KEY, JSON.stringify(state.recentDestinations.slice(0, 4)));
  } catch {
    // Ignore persistence failures.
  }
}

function normalizeFilenameStyle(value) {
  return Object.prototype.hasOwnProperty.call(FILENAME_STYLE_OPTIONS, value) ? value : DEFAULT_FILENAME_STYLE;
}

function loadFilenameStyle() {
  try {
    return normalizeFilenameStyle(window.localStorage.getItem(FILENAME_STYLE_KEY));
  } catch {
    return DEFAULT_FILENAME_STYLE;
  }
}

function saveFilenameStyle() {
  try {
    window.localStorage.setItem(FILENAME_STYLE_KEY, state.filenameStyle);
  } catch {
    // Ignore persistence failures.
  }
}

function rememberDestination(pathValue) {
  const normalized = String(pathValue || "").trim();
  if (!normalized) {
    return;
  }
  state.recentDestinations = [normalized, ...state.recentDestinations.filter((value) => value !== normalized)].slice(0, 4);
  saveRecentDestinations();
}

function isHostedBrowserOutputMode() {
  return state.serverMode && !state.capabilities.nativeOutputPicker;
}

function browserDirectoryPickerAvailable() {
  return isHostedBrowserOutputMode() && typeof window.showDirectoryPicker === "function" && window.isSecureContext !== false;
}

function outputDestinationReady() {
  if (state.capabilities.nativeOutputPicker) {
    return Boolean(state.outputDir);
  }
  if (browserDirectoryPickerAvailable()) {
    return Boolean(state.outputDir && state.browserOutputDirectoryHandle);
  }
  return Boolean(state.outputDir);
}

function serverCanOpenOutputFolder() {
  return Boolean(state.outputDir) && !isHostedBrowserOutputMode();
}

function defaultStatusDetail() {
  if (isHostedBrowserOutputMode()) {
    if (state.browserOutputDirLabel) {
      return `Saving copies to ${state.browserOutputDirLabel}.`;
    }
    if (browserDirectoryPickerAvailable()) {
      return "Click Convert to PDF and choose a local save folder in your browser.";
    }
    return "Converted PDFs will be available from the result list.";
  }
  if (state.outputDir) {
    return `Saving to ${state.outputDir}`;
  }
  if (!state.capabilities.nativeOutputPicker) {
    return "The hosted deployment expects a server-managed save location.";
  }
  return "Click Convert to PDF and choose the destination folder when prompted.";
}

function formatDestinationLabel() {
  if (state.browserOutputDirLabel) {
    return state.browserOutputDirLabel;
  }
  if (isHostedBrowserOutputMode()) {
    return browserDirectoryPickerAvailable()
      ? "Choose a local folder before conversion."
      : "PDFs will be available for download.";
  }
  if (state.outputDirLabel || state.outputDir) {
    return state.outputDirLabel || state.outputDir;
  }
  return "Choose a save folder before conversion.";
}

function refreshHostingCopy() {
  if (elements.appEyebrow) {
    elements.appEyebrow.textContent = state.serverMode ? "Hosted PDF Converter" : "Local PDF Converter";
  }
  if (elements.topbarBody) {
    elements.topbarBody.innerHTML = state.capabilities.outlookImport
      ? "Drop Outlook emails or <code>.msg</code> files, then convert them to PDFs with the selected filename style."
      : "Upload <code>.msg</code> files, then convert them to PDFs with the selected filename style.";
  }
  if (elements.workflowBody) {
    elements.workflowBody.textContent = isHostedBrowserOutputMode()
      ? state.browserOutputDirLabel
        ? `Upload .msg files or drag them into the page. Converted PDFs will be copied to ${state.browserOutputDirLabel}.`
        : browserDirectoryPickerAvailable()
          ? "Upload .msg files or drag them into the page. When you are ready, click Convert to PDF and choose a local save folder."
          : "Upload .msg files or drag them into the page. Converted PDFs will be available from the result list."
      : state.outputDir
      ? `Upload .msg files or drag them into the page. Converted PDFs will be saved to ${state.outputDir}.`
      : state.capabilities.outlookImport
        ? "Drag from Outlook or click to browse. When you are ready, click Convert to PDF and choose where to save the files."
        : "Click to browse for .msg files or drag them into the page, then convert them to PDFs.";
  }
}

function refreshFilenameStyleCopy() {
  if (elements.filenameStyleSelect) {
    elements.filenameStyleSelect.value = state.filenameStyle;
  }
  if (elements.filenameStyleDetail) {
    elements.filenameStyleDetail.textContent = FILENAME_STYLE_OPTIONS[state.filenameStyle]?.detail || FILENAME_STYLE_OPTIONS[DEFAULT_FILENAME_STYLE].detail;
  }
  refreshFilenameStyleExample();
}

function filenameStyleSample() {
  return FILENAME_STYLE_OPTIONS[state.filenameStyle]?.example || FILENAME_STYLE_OPTIONS[DEFAULT_FILENAME_STYLE].example;
}

function liveFilenamePreviewItem() {
  return (
    state.items.find((item) => !TERMINAL_STAGES.has(item?.stage) && (item?.outputPath || item?.outputName)) ||
    state.items.find((item) => !TERMINAL_STAGES.has(item?.stage) && item?.name) ||
    null
  );
}

function refreshFilenameStyleExample() {
  if (!elements.filenameStyleExampleValue) {
    return;
  }
  const previewItem = liveFilenamePreviewItem();
  const previewName = previewItem ? queueOutputName(previewItem) : filenameStyleSample();
  elements.filenameStyleExampleValue.textContent = previewName;
  elements.filenameStyleExampleValue.title = previewName;
}

function lastPathSegment(value) {
  const parts = String(value || "")
    .split(/[/\\]/)
    .filter(Boolean);
  return parts[parts.length - 1] || value || "";
}

function parentPath(value) {
  const rawValue = String(value || "").trim();
  if (!rawValue) {
    return "";
  }
  const lastSeparator = Math.max(rawValue.lastIndexOf("\\"), rawValue.lastIndexOf("/"));
  if (lastSeparator <= 0) {
    return "";
  }
  return rawValue.slice(0, lastSeparator);
}

function formatSource(source) {
  if (source === "outlook") {
    return "Outlook drag";
  }
  if (source === "upload") {
    return "Browser upload";
  }
  return "Local queue";
}

function formatQueuePercent(progress) {
  return `${Math.round(progress?.percent || 0)}%`;
}

function formatByteSize(bytes) {
  const value = Number(bytes || 0);
  if (!Number.isFinite(value) || value <= 0) {
    return "";
  }
  if (value >= 1024 * 1024) {
    return `${(value / (1024 * 1024)).toFixed(1)} MB`;
  }
  if (value >= 1024) {
    return `${Math.round(value / 1024)} KB`;
  }
  return `${value} B`;
}

function formatTimelineTime(value) {
  const date = value ? new Date(value) : new Date();
  if (Number.isNaN(date.getTime())) {
    return "Now";
  }
  return date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
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

function queueOutputName(item) {
  if (item.outputPath) {
    return lastPathSegment(item.outputPath);
  }
  if (item.outputName) {
    return item.outputName;
  }
  const baseName = String(item.name || "").replace(/\.msg$/i, "");
  return `${baseName}.pdf`;
}

function reviewFilenameParts(value) {
  const outputName = String(value || "");
  const stem = outputName.replace(/\.pdf$/i, "");
  const extension = outputName.toLowerCase().endsWith(".pdf") ? "PDF" : "";
  const senderSubjectSplit = stem.lastIndexOf("__");
  if (senderSubjectSplit > 0 && senderSubjectSplit < stem.length - 2) {
    return {
      prefixLabel: "From",
      prefixValue: stem.slice(0, senderSubjectSplit).replace(/_+$/g, "").trim(),
      name: stem.slice(senderSubjectSplit + 2).trim() || stem,
      extension,
    };
  }

  const [firstPart, ...remainingParts] = stem.split("_");
  const hasUsefulSplit = remainingParts.length > 0 && firstPart.trim().length > 0;
  const isDatePrefix = /^\d{4}-\d{2}-\d{2}$/.test(firstPart);
  if (!hasUsefulSplit) {
    return { prefixLabel: "", prefixValue: "", name: stem || outputName, extension };
  }

  return {
    prefixLabel: isDatePrefix ? "Date" : "From",
    prefixValue: firstPart.trim(),
    name: remainingParts.join("_").trim() || stem,
    extension,
  };
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
    return { tone: progress.tone || "active", label: progress.label || "Converting" };
  }
  return { tone: "queued", label: "Ready" };
}

function deriveItemSummary(item, progress) {
  if (item.stage === "failed") {
    return failureExplanation(item.error).reason;
  }
  if (item.stage === "complete") {
    return item.outputPath ? `Saved as ${lastPathSegment(item.outputPath)}` : "PDF saved.";
  }
  if (progress?.tone === "queued") {
    return "Queued and ready to convert.";
  }
  if (progress?.active) {
    return `${progress.label || queueStageLabel(item.stage)} ${formatQueuePercent(progress)}`;
  }
  return formatSource(item.source);
}

function compactErrorLabel(error) {
  const normalized = String(error || "").replace(/\s+/g, " ").trim();
  const lowered = normalized.toLowerCase();
  if (!normalized) {
    return "The file did not finish converting.";
  }
  if (lowered.includes("timed out") || lowered.includes("timeout")) {
    return "Conversion timed out. Retry this file.";
  }
  if (lowered.includes("permission") || lowered.includes("access is denied")) {
    return "Save folder needs permission.";
  }
  if (lowered.includes("not found") || lowered.includes("missing") || lowered.includes("no such file")) {
    return "Original file is missing.";
  }
  if (lowered.includes("ole2") || lowered.includes("structured storage")) {
    return "This is not a valid Outlook .msg email.";
  }
  return normalized.length > 86 ? `${normalized.slice(0, 83).trimEnd()}...` : normalized;
}

function failureExplanation(error) {
  const normalized = String(error || "").replace(/\s+/g, " ").trim();
  const lowered = normalized.toLowerCase();
  if (!normalized) {
    return {
      reason: "The file did not finish converting.",
      action: "Retry it once. If it fails again, re-export the email from Outlook and drop the fresh copy here.",
    };
  }
  if (lowered.includes("timed out") || lowered.includes("timeout")) {
    return {
      reason: "The conversion took too long.",
      action: "Retry this file by itself. Large attachments or embedded images can need another pass.",
    };
  }
  if (lowered.includes("permission") || lowered.includes("access is denied")) {
    return {
      reason: "The save folder could not be written to.",
      action: "Choose a folder you can edit, then retry the failed file.",
    };
  }
  if (lowered.includes("not found") || lowered.includes("missing") || lowered.includes("no such file")) {
    return {
      reason: "The original email file is no longer available.",
      action: "Drop the email into the app again, then retry.",
    };
  }
  if (lowered.includes("ole2") || lowered.includes("structured storage")) {
    return {
      reason: "This is not a valid Outlook .msg email.",
      action: "Drop the original email from Outlook again, then retry. Files renamed to .msg will not convert.",
    };
  }
  if (lowered.includes("parse") || lowered.includes("extract") || lowered.includes("read")) {
    return {
      reason: "The email could not be read cleanly.",
      action: "Retry it once. If it still fails, re-export the message from Outlook and try the new copy.",
    };
  }
  return {
    reason: compactErrorLabel(error),
    action: "Retry this file. Saved PDFs from the same batch will be left alone.",
  };
}

function outputPreviewLabel(item) {
  if (item.outputPath) {
    return lastPathSegment(item.outputPath);
  }
  if (item.outputName) {
    return item.outputName;
  }
  const baseName = String(item.name || "").replace(/\.msg$/i, "");
  return `${baseName}.pdf with date prefix`;
}

function buildQueueProgressStateFromSnapshot(item) {
  if (!item?.stage || !QUEUE_STAGE_PROGRESS[item.stage]) {
    return null;
  }
  const model = QUEUE_STAGE_PROGRESS[item.stage];
  const isTerminal = TERMINAL_STAGES.has(item.stage);
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

function progressPercentForItem(item) {
  if (item.stage === "complete" || item.stage === "failed") {
    const progress = queueProgressStateForItem(item);
    if (progress?.active && progress.stage === item.stage && progress.percent < 100) {
      return progress.percent;
    }
    return 100;
  }
  const progress = queueProgressStateForItem(item);
  if (progress?.active) {
    return progress.percent || 0;
  }
  if (item.stage === "files_accepted" || !item.stage) {
    return 8;
  }
  return 0;
}

function delay(milliseconds) {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}

function summarizeBatch() {
  const queued = state.items.length;
  const complete = state.items.filter((item) => item.stage === "complete").length;
  const failed = state.items.filter((item) => item.stage === "failed").length;
  const activeItems = state.items.filter((item) => !TERMINAL_STAGES.has(item.stage));
  const ready = activeItems.filter((item) => !item.stage || item.stage === "files_accepted").length;
  const activeProgress = activeItems
    .map((item) => queueProgressStateForItem(item))
    .filter((progress) => progress?.active);
  const activelyProcessing = activeItems.filter((item) => !["", "drop_received", "files_accepted"].includes(item.stage || ""));
  const totalProgress = queued
    ? Math.round(state.items.reduce((sum, item) => sum + progressPercentForItem(item), 0) / queued)
    : 0;
  const activeAverageProgress = activeProgress.length
    ? Math.round(activeProgress.reduce((sum, progress) => sum + (progress.percent || 0), 0) / activeProgress.length)
    : totalProgress;
  const successPercent = queued ? Math.round((complete / queued) * 100) : 0;

  return {
    queued,
    ready,
    complete,
    failed,
    activeCount: activeItems.length,
    activelyProcessingCount: activelyProcessing.length,
    averageProgress: totalProgress,
    activeAverageProgress,
    successPercent,
  };
}

function renderBatchReview({ shouldShowResult, destinationKnown }) {
  if (!elements.resultReview || !elements.resultReviewList || !elements.resultReviewDestination) {
    return;
  }

  const reviewItems = state.items.filter((item) => item.stage === "complete" || item.stage === "failed");
  elements.resultReview.hidden = !shouldShowResult || reviewItems.length === 0;
  if (elements.resultReview.hidden) {
    elements.resultReviewList.innerHTML = "";
    return;
  }

  elements.resultReviewDestination.textContent = destinationKnown
    ? formatDestinationLabel()
    : "Output folder not available";

  elements.resultReviewList.innerHTML = reviewItems
    .map((item) => {
      const isSaved = item.stage === "complete";
      const primary = isSaved ? queueOutputName(item) : item.name;
      const nameParts = reviewFilenameParts(primary);
      const failure = isSaved ? null : failureExplanation(item.error);
      const detail = isSaved ? "Saved PDF" : failure.reason;
      const action = failure?.action || "";
      const revealAction = isSaved && item.outputPath && !isHostedBrowserOutputMode()
        ? `<button class="result-review-reveal button button-ghost" type="button" data-result-reveal-path="${escapeHtml(item.outputPath)}" data-result-reveal-name="${escapeHtml(primary)}" data-testid="result-reveal-file" aria-label="Show ${escapeHtml(primary)} in output folder"><span class="folder-open-icon" aria-hidden="true"></span><span data-result-reveal-label>Show</span></button>`
        : "";
      const downloadAction = isSaved && item.outputPath && isHostedBrowserOutputMode()
        ? `<button class="result-review-reveal result-review-download button button-ghost" type="button" data-result-download-id="${escapeHtml(item.id)}" data-result-download-name="${escapeHtml(primary)}" data-testid="result-download-file" aria-label="Download ${escapeHtml(primary)}"><span class="folder-open-icon" aria-hidden="true"></span><span data-result-reveal-label>Download</span></button>`
        : "";
      return `
        <div class="result-review-row is-${isSaved ? "saved" : "retry"}">
          <span class="result-review-status">${isSaved ? "Saved" : "Failed"}</span>
          <div class="result-review-file">
            <div class="result-review-name-row" title="${escapeHtml(primary)}">
              ${nameParts.prefixValue ? `<span class="result-review-prefix"><span class="result-review-prefix-label">${escapeHtml(nameParts.prefixLabel)}</span><span class="result-review-prefix-value">${escapeHtml(nameParts.prefixValue)}</span></span>` : ""}
              <strong class="result-review-name">${escapeHtml(nameParts.name)}</strong>
              ${nameParts.extension ? `<span class="result-review-extension">${escapeHtml(nameParts.extension)}</span>` : ""}
            </div>
            <span class="result-review-detail">${escapeHtml(detail)}</span>
            ${action ? `<small>${escapeHtml(action)}</small>` : ""}
          </div>
          ${revealAction || downloadAction}
        </div>
      `;
    })
    .join("");
}

function buildQueueTerminalSummary(summary, { isTerminalBatch, visibleRows }) {
  if (!isTerminalBatch) {
    return "";
  }
  const expanded = state.queueDetailsExpanded;
  const hasFailures = summary.failed > 0;
  const headline = hasFailures
    ? `${summary.failed} file${summary.failed === 1 ? "" : "s"} need retry`
    : `${summary.complete} file${summary.complete === 1 ? "" : "s"} converted successfully`;
  const detail = expanded
    ? "Detailed queue rows are shown below."
    : hasFailures
      ? "Saved PDFs are summarized in Batch Result. Failed rows stay visible below."
      : "Review the saved PDFs in Batch Result, or expand the queue details if needed.";
  const actionLabel = expanded
    ? "Hide details"
    : hasFailures
      ? "Show saved details"
      : "Show queue details";

  return `
    <div class="queue-terminal-summary ${hasFailures ? "is-attention" : "is-success"}">
      <div class="queue-terminal-copy">
        <strong>${escapeHtml(headline)}</strong>
        <span>${escapeHtml(detail)}</span>
      </div>
      <button class="queue-details-toggle" data-queue-details-toggle type="button" aria-expanded="${expanded ? "true" : "false"}">
        ${escapeHtml(actionLabel)}
      </button>
    </div>
    ${visibleRows.length ? "" : `<div class="queue-terminal-empty">Queue rows are hidden because the Batch Result has the review list.</div>`}
  `;
}

function renderOperations() {
  const summary = summarizeBatch();
  const destinationKnown = outputDestinationReady();
  const hasQueuedItems = summary.queued > 0;
  const isProcessing = state.isBusy || summary.activelyProcessingCount > 0;
  const isComplete = hasQueuedItems && summary.complete === summary.queued && summary.failed === 0;
  const isAttention = summary.failed > 0 && !isProcessing;
  const mode = isComplete ? "complete" : isAttention ? "attention" : isProcessing ? "processing" : hasQueuedItems ? "queued" : "idle";

  elements.metricQueued.textContent = String(summary.queued);
  elements.metricReady.textContent = String(summary.ready);
  elements.metricComplete.textContent = String(summary.complete);
  elements.metricFailed.textContent = String(summary.failed);
  elements.shell.dataset.mode = mode;
  if (elements.saveCard) {
    elements.saveCard.hidden = isComplete || isAttention;
  }

  elements.destinationInlineValue.textContent = destinationKnown
    ? formatDestinationLabel()
    : isHostedBrowserOutputMode()
      ? browserDirectoryPickerAvailable()
        ? "Choose a local folder when you convert."
        : "Download PDFs after conversion."
    : !state.capabilities.nativeOutputPicker
      ? "Server-managed output folder."
      : "Choose a destination when you convert.";

  updateReadinessSignals(summary, { destinationKnown, isProcessing, isComplete, isAttention });

  if (elements.historyStrip && elements.historyList) {
    const visibleDestinations = state.recentDestinations.filter((value) => value !== state.outputDir);
    elements.historyStrip.hidden = visibleDestinations.length === 0;
    elements.historyList.innerHTML = visibleDestinations
      .map((value) => `<button class="history-chip" type="button" data-destination="${escapeHtml(value)}">${escapeHtml(lastPathSegment(value) || value)}</button>`)
      .join("");
  }

  if (elements.modeBannerTitle && elements.modeBannerDetail) {
    if (!hasQueuedItems) {
      elements.modeBannerTitle.textContent = "Ready for intake";
      elements.modeBannerDetail.textContent = "Drop emails to start a new conversion batch.";
    } else if (isProcessing) {
      elements.modeBannerTitle.textContent = "Conversion is underway";
      elements.modeBannerDetail.textContent = `${summary.complete} saved so far. Keep this page open while the batch finishes.`;
    } else if (isComplete) {
      elements.modeBannerTitle.textContent = "Batch complete";
      elements.modeBannerDetail.textContent = destinationKnown
        ? `All queued files were saved to ${formatDestinationLabel()}.`
        : "All queued files were saved successfully.";
    } else if (isAttention) {
      elements.modeBannerTitle.textContent = "Review the failed files";
      elements.modeBannerDetail.textContent = `${summary.failed} file(s) need another attempt. The saved PDFs are already finished.`;
    } else {
      elements.modeBannerTitle.textContent = "Batch is staged";
      elements.modeBannerDetail.textContent = destinationKnown
        ? `Destination confirmed. ${summary.ready} file(s) are ready to convert.`
        : `${summary.ready} file(s) are queued. Click Convert to choose a destination and start the batch.`;
    }
  }

  if (elements.resultBanner && elements.resultHeadline && elements.resultDetail && elements.resultSaved && elements.resultFailed) {
    elements.resultSaved.textContent = String(summary.complete);
    elements.resultFailed.textContent = String(summary.failed);
    const shouldShowResult = isComplete || isAttention;
    elements.resultBanner.hidden = !shouldShowResult;
    elements.resultBanner.classList.toggle("is-success", isComplete);
    elements.resultBanner.classList.toggle("is-attention", isAttention);
    renderBatchReview({ shouldShowResult, destinationKnown });
    if (elements.resultOpenOutputButton) {
      const canOpenOutput = destinationKnown && summary.complete > 0 && serverCanOpenOutputFolder();
      elements.resultOpenOutputButton.hidden = !canOpenOutput;
      elements.resultOpenOutputButton.disabled = !canOpenOutput;
    }
    if (elements.resultRetryFailedButton) {
      elements.resultRetryFailedButton.hidden = !isAttention;
      elements.resultRetryFailedButton.disabled = !isAttention || state.isBusy;
    }
    if (elements.resultStartNewButton) {
      elements.resultStartNewButton.disabled = state.isBusy;
    }
    if (shouldShowResult) {
      if (isComplete) {
        elements.resultHeadline.textContent = `${summary.complete} PDF${summary.complete === 1 ? "" : "s"} saved`;
        elements.resultDetail.textContent = destinationKnown
          ? `Every queued email has been converted and saved to ${formatDestinationLabel()}.`
          : "Every queued email has been converted successfully.";
        if (elements.resultGuidance) {
          elements.resultGuidance.textContent = isHostedBrowserOutputMode()
            ? state.browserOutputDirLabel
              ? "PDFs were copied to your selected folder. Use Download if you need another copy."
              : "Use Download on each saved PDF to keep a local copy."
            : destinationKnown
            ? "Open the output folder to review the PDFs, or start a new batch when you are ready."
            : "The full batch is complete. Start a new batch when you are ready.";
        }
      } else {
        elements.resultHeadline.textContent = `${summary.complete} saved, ${summary.failed} need retry`;
        elements.resultDetail.textContent = "The saved PDFs are finished. Only the failed files below need another attempt.";
        if (elements.resultGuidance) {
          elements.resultGuidance.textContent = "Use Retry Failed to convert only the failed files again. The saved PDFs will not be duplicated.";
        }
      }
    }
  }

  if (elements.openOutputFolderButton) {
    const resultPanelOwnsOutputAction = isComplete || isAttention;
    const shouldShowOpenOutput = destinationKnown && serverCanOpenOutputFolder() && !isProcessing && summary.complete > 0 && !resultPanelOwnsOutputAction;
    elements.openOutputFolderButton.hidden = !shouldShowOpenOutput;
    elements.openOutputFolderButton.disabled = !shouldShowOpenOutput;
  }

  if (elements.convertButton) {
    elements.convertButton.hidden = isComplete;
  }

  const setBatchProgress = (percent) => {
    const clamped = Math.max(0, Math.min(100, Math.round(percent || 0)));
    elements.batchProgressFill.style.width = `${clamped}%`;
    elements.batchProgressTrack?.setAttribute("aria-valuenow", String(clamped));
  };

  if (!summary.queued) {
    elements.batchStatusPill.textContent = "Ready";
    elements.batchProgressHeadline.textContent = "No files in progress";
    elements.batchProgressDetail.textContent = "Add files to see the live conversion timeline.";
    setBatchProgress(0);
  } else if (state.isBusy || summary.activelyProcessingCount) {
    elements.batchStatusPill.textContent = summary.failed ? "In progress with failed files" : "Converting";
    elements.batchProgressHeadline.textContent = `${summary.complete} of ${summary.queued} files saved`;
    elements.batchProgressDetail.textContent = summary.failed
      ? `${summary.failed} failed file(s) so far. Progress is estimated from live conversion stages.`
      : `${summary.activelyProcessingCount} file(s) active. Progress is estimated from live conversion stages.`;
    setBatchProgress(summary.averageProgress);
  } else if (summary.failed) {
    elements.batchStatusPill.textContent = "Needs attention";
    elements.batchProgressHeadline.textContent = `${summary.failed} file(s) need a retry`;
    elements.batchProgressDetail.textContent = "Review the failed rows below and retry only those files.";
    setBatchProgress(100);
  } else if (summary.complete === summary.queued) {
    elements.batchStatusPill.textContent = "Complete";
    elements.batchProgressHeadline.textContent = "Every queued file has been saved";
    elements.batchProgressDetail.textContent = "You can clear the queue or keep the rows as a quick audit trail.";
    setBatchProgress(100);
  } else {
    elements.batchStatusPill.textContent = "Ready";
    elements.batchProgressHeadline.textContent = `${summary.ready} file(s) ready to convert`;
    elements.batchProgressDetail.textContent = destinationKnown
      ? `Destination confirmed: ${formatDestinationLabel()}`
      : "Click Convert to choose a destination and start the batch.";
    setBatchProgress(Math.max(8, Math.min(14, summary.averageProgress || 8)));
  }

  const completedStages = new Set();
  const activeStages = new Set();
  for (const item of state.items) {
    const stage = item.stage || "files_accepted";
    const currentIndex = TIMELINE_STAGES.findIndex((entry) => entry.key === stage);
    if (currentIndex === -1) {
      continue;
    }
    TIMELINE_STAGES.forEach((entry, index) => {
      if (index < currentIndex) {
        completedStages.add(entry.key);
      }
    });
    if (stage === "complete") {
      completedStages.add("complete");
      continue;
    }
    if (stage === "failed") {
      activeStages.add("pipeline_selected");
      continue;
    }
    activeStages.add(stage);
  }

  if (!state.items.length) {
    elements.timelineList.innerHTML = `
      <div class="timeline-row is-active">
        <span class="timeline-time">Now</span>
        <span class="timeline-dot"></span>
        <span class="timeline-copy">Queue idle. Waiting for files.</span>
      </div>
    `;
    return;
  }

  const recentItems = [...state.items].slice(-6).reverse();
  elements.timelineList.innerHTML = recentItems.map((item) => {
    const progress = queueProgressStateForItem(item);
    const visual = queueVisualState(item, progress);
    const tone = visual.tone === "complete"
      ? "is-complete"
      : visual.tone === "failed"
        ? "is-failed"
        : progress?.active
          ? "is-active"
          : "";
    const detail = item.stage === "complete"
      ? "Converted successfully"
      : item.stage === "failed"
        ? failureExplanation(item.error).reason
        : progress?.active
          ? `${progress.label || queueStageLabel(item.stage)} ${formatQueuePercent(progress)}`
          : "Queued and ready";
    return `
      <div class="timeline-row ${tone}">
        <span class="timeline-time">${escapeHtml(formatTimelineTime(item.createdAt))}</span>
        <span class="timeline-dot"></span>
        <span class="timeline-copy">${escapeHtml(item.name)}<small>${escapeHtml(detail)}</small></span>
      </div>
    `;
  }).join("");
}

function setReadinessStep(stepElement, valueElement, tone, value) {
  if (!stepElement || !valueElement) {
    return;
  }
  stepElement.className = `readiness-step is-${tone}`;
  valueElement.textContent = value;
}

function pluralizeCount(count, singular, plural = `${singular}s`) {
  return `${count} ${count === 1 ? singular : plural}`;
}

function updateReadinessSignals(summary, { destinationKnown, isProcessing, isComplete, isAttention }) {
  if (!elements.readinessFiles || !elements.readinessDestination || !elements.readinessConversion) {
    return;
  }

  if (!summary.queued) {
    setReadinessStep(elements.readinessFiles, elements.readinessFilesValue, "needed", "Waiting");
  } else if (summary.failed && !isProcessing) {
    setReadinessStep(elements.readinessFiles, elements.readinessFilesValue, "warning", pluralizeCount(summary.failed, "issue"));
  } else if (isComplete) {
    setReadinessStep(elements.readinessFiles, elements.readinessFilesValue, "done", `${summary.complete} saved`);
  } else {
    setReadinessStep(elements.readinessFiles, elements.readinessFilesValue, "done", `${summary.ready || summary.queued} ready`);
  }

  if (!state.capabilities.nativeOutputPicker) {
    if (browserDirectoryPickerAvailable()) {
      if (state.browserOutputDirectoryHandle) {
        setReadinessStep(elements.readinessDestination, elements.readinessDestinationValue, "done", lastPathSegment(state.browserOutputDirLabel) || "Selected");
      } else if (summary.queued && !isComplete) {
        setReadinessStep(elements.readinessDestination, elements.readinessDestinationValue, "active", "Choose now");
      } else {
        setReadinessStep(elements.readinessDestination, elements.readinessDestinationValue, "pending", "Choose later");
      }
    } else {
      setReadinessStep(elements.readinessDestination, elements.readinessDestinationValue, "done", "Download");
    }
  } else if (destinationKnown) {
    setReadinessStep(elements.readinessDestination, elements.readinessDestinationValue, "done", lastPathSegment(formatDestinationLabel()) || "Selected");
  } else if (summary.queued && !isComplete) {
    setReadinessStep(elements.readinessDestination, elements.readinessDestinationValue, "active", "Choose now");
  } else {
    setReadinessStep(elements.readinessDestination, elements.readinessDestinationValue, "pending", "Choose later");
  }

  if (isProcessing) {
    setReadinessStep(elements.readinessConversion, elements.readinessConversionValue, "active", `${Math.round(summary.averageProgress)}%`);
  } else if (isComplete) {
    setReadinessStep(elements.readinessConversion, elements.readinessConversionValue, "done", "Saved");
  } else if (isAttention) {
    setReadinessStep(elements.readinessConversion, elements.readinessConversionValue, "warning", "Retry");
  } else if (summary.queued) {
    setReadinessStep(elements.readinessConversion, elements.readinessConversionValue, "active", "Ready");
  } else {
    setReadinessStep(elements.readinessConversion, elements.readinessConversionValue, "pending", "Not started");
  }
}

async function openOutputFolder() {
  if (isHostedBrowserOutputMode()) {
    addStatus(
      state.browserOutputDirLabel ? "PDFs were saved to your selected folder." : "Output folders cannot be opened from the hosted page.",
      state.browserOutputDirLabel || "Use the Download button beside a saved PDF.",
      "neutral",
    );
    return;
  }
  if (!state.outputDir) {
    addStatus("No output folder is selected yet.", "Convert a batch first, then open the saved destination.", "error");
    return;
  }

  await api("/api/open-output-folder", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ output_dir: state.outputDir }),
  });
  addStatus("Output folder opened.", state.outputDir, "success");
}

async function revealOutputFile(outputPath, outputName = "PDF") {
  const normalizedPath = String(outputPath || "").trim();
  if (!normalizedPath) {
    addStatus("This PDF cannot be opened yet.", "Use Open Output Folder to review the saved files.", "error");
    return;
  }

  await api("/api/reveal-output-file", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ output_path: normalizedPath }),
  });
  addStatus("PDF shown in folder.", outputName, "success");
}

function outputFileDownloadPath(itemId) {
  return `/api/output-file/${encodeURIComponent(itemId)}`;
}

function safeOutputFileName(item, fallbackName = "converted.pdf") {
  const rawName = String(queueOutputName(item) || fallbackName || "converted.pdf").trim();
  const fileName = lastPathSegment(rawName).replace(/[\\/:*?"<>|]+/g, "_").trim();
  return fileName || "converted.pdf";
}

async function fetchOutputBlob(itemId) {
  const response = await fetch(outputFileDownloadPath(itemId));
  if (!response.ok) {
    let detail = response.statusText || "PDF download failed";
    try {
      const payload = await response.json();
      detail = payload?.detail || detail;
    } catch {
      // Keep the HTTP status text.
    }
    throw new Error(detail);
  }
  return response.blob();
}

function triggerBlobDownload(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30000);
}

async function saveOutputItemToBrowserFolder(item) {
  if (!state.browserOutputDirectoryHandle) {
    throw new Error("Choose a browser save folder before saving PDFs.");
  }
  const fileName = safeOutputFileName(item);
  const blob = await fetchOutputBlob(item.id);
  const fileHandle = await state.browserOutputDirectoryHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  try {
    await writable.write(blob);
  } finally {
    await writable.close();
  }
  return fileName;
}

async function downloadOutputFile(itemId, outputName = "PDF") {
  const item = state.items.find((candidate) => candidate.id === itemId);
  const fileName = item ? safeOutputFileName(item, outputName) : String(outputName || "converted.pdf");
  const blob = await fetchOutputBlob(itemId);
  triggerBlobDownload(blob, fileName);
  addStatus("PDF download started.", fileName, "success");
}

async function deliverConvertedFiles(ids) {
  if (!isHostedBrowserOutputMode()) {
    return { mode: "server", saved: 0, failed: 0 };
  }

  const idSet = new Set(ids);
  const completedItems = state.items.filter((item) => idSet.has(item.id) && item.stage === "complete" && item.outputPath);
  if (!completedItems.length) {
    return { mode: "browser", saved: 0, failed: 0 };
  }

  if (!state.browserOutputDirectoryHandle) {
    return { mode: "download", saved: 0, failed: 0 };
  }

  let saved = 0;
  const failures = [];
  for (const item of completedItems) {
    try {
      await saveOutputItemToBrowserFolder(item);
      saved += 1;
    } catch (error) {
      failures.push({ item, error });
    }
  }

  if (saved) {
    addStatus(
      `${saved} PDF${saved === 1 ? "" : "s"} copied to your save folder.`,
      state.browserOutputDirLabel || "Selected browser folder",
      "success",
    );
  }
  failures.forEach(({ item, error }) => {
    addStatus("A PDF could not be copied to your folder.", `${safeOutputFileName(item)}: ${error.message}`, "error");
  });

  return { mode: "browser-folder", saved, failed: failures.length };
}

function setRevealButtonFeedback(button, { label, tone }) {
  if (!button) {
    return;
  }
  const labelElement = button.querySelector("[data-result-reveal-label]");
  button.classList.remove("is-confirmed", "is-error");
  if (tone) {
    button.classList.add(`is-${tone}`);
  }
  if (labelElement) {
    labelElement.textContent = label;
  }
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
  const isTerminal = TERMINAL_STAGES.has(stage);
  const shouldEaseTerminal = isTerminal && previous?.active && currentPercent < nextCap - 0.02;

  state.queueProgressByTaskId = {
    ...state.queueProgressByTaskId,
    [taskId]: {
      active: true,
      stage,
      label: model?.label || queueStageLabel(stage),
      tone: model?.tone || previous?.tone || "active",
      percent: shouldEaseTerminal ? currentPercent : isTerminal ? nextCap : currentPercent,
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

function flashDropzone(tone, message, detail = "") {
  if (!elements.dropzone) {
    return;
  }
  if (state.dropFeedbackTimer) {
    clearTimeout(state.dropFeedbackTimer);
  }
  elements.dropzone.classList.remove("is-drop-accepted", "is-drop-rejected", "is-drop-receiving");
  const className = tone === "pending" ? "is-drop-receiving" : tone === "error" ? "is-drop-rejected" : "is-drop-accepted";
  elements.dropzone.classList.add(className);
  elements.dropzone.dataset.dropFeedback = message;
  if (elements.dropFeedbackTitle) {
    elements.dropFeedbackTitle.textContent = message;
  }
  if (elements.dropFeedbackDetail) {
    elements.dropFeedbackDetail.textContent = detail || (tone === "pending"
      ? "Preparing the queue preview."
      : tone === "error"
        ? "Only Outlook .msg files can be staged."
        : "Date-prefixed PDF name prepared.");
  }
  if (tone === "pending") {
    return;
  }
  state.dropFeedbackTimer = window.setTimeout(() => {
    elements.dropzone.classList.remove("is-drop-accepted", "is-drop-rejected", "is-drop-receiving");
    delete elements.dropzone.dataset.dropFeedback;
    state.dropFeedbackTimer = 0;
  }, 2400);
}

function clearDropzoneFeedback() {
  if (!elements.dropzone) {
    return;
  }
  if (state.dropFeedbackTimer) {
    clearTimeout(state.dropFeedbackTimer);
  }
  dropzoneController?.clear();
  elements.dropzone.classList.remove("is-drop-accepted", "is-drop-rejected", "is-drop-receiving", "is-drop-splash");
  delete elements.dropzone.dataset.dropFeedback;
  if (elements.dropFeedbackTitle) {
    elements.dropFeedbackTitle.textContent = "Email staged";
  }
  if (elements.dropFeedbackDetail) {
    elements.dropFeedbackDetail.textContent = "Date-prefixed PDF name prepared.";
  }
  state.dropFeedbackTimer = 0;
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

function renderFeedbackModal() {
  if (!elements.feedbackModal || !elements.feedbackBackdrop || !elements.feedbackButton) {
    return;
  }

  const isOpen = state.feedbackOpen;
  elements.feedbackModal.hidden = !isOpen;
  elements.feedbackBackdrop.hidden = !isOpen;
  elements.feedbackButton.setAttribute("aria-expanded", isOpen ? "true" : "false");
  document.body.classList.toggle("feedback-open", isOpen);
}

function renderTutorialModal() {
  if (!elements.tutorialModal || !elements.tutorialBackdrop || !elements.tutorialButton) {
    return;
  }

  const isOpen = state.tutorialOpen;
  elements.tutorialModal.hidden = !isOpen;
  elements.tutorialBackdrop.hidden = !isOpen;
  elements.tutorialButton.setAttribute("aria-expanded", isOpen ? "true" : "false");
  document.body.classList.toggle("tutorial-open", isOpen);
}

function openTutorial() {
  state.tutorialOpen = true;
  renderTutorialModal();
  elements.tutorialClose?.focus();
}

function closeTutorial() {
  state.tutorialOpen = false;
  renderTutorialModal();
  elements.tutorialButton?.focus();
}

function openFeedback() {
  state.feedbackOpen = true;
  renderFeedbackModal();
  elements.feedbackImprove?.focus();
}

function closeFeedback() {
  state.feedbackOpen = false;
  renderFeedbackModal();
  elements.feedbackButton?.focus();
}

function buildFeedbackContext() {
  return {
    appName: "Email-PDF Converter",
    url: window.location.href,
    path: window.location.pathname,
    queuedCount: state.items.length,
    activeQueueCount: state.items.filter((item) => item.stage !== "complete").length,
    completedCount: state.items.filter((item) => item.stage === "complete").length,
    failedCount: state.items.filter((item) => item.stage === "failed").length,
    serverMode: state.serverMode,
    outputDir: state.outputDir || null,
    isBusy: state.isBusy,
    activeConvertIds: state.activeConvertIds,
    capabilities: state.capabilities,
    userAgent: navigator.userAgent,
  };
}

async function submitFeedback() {
  const category = elements.feedbackCategory?.value || "other";
  const improve = elements.feedbackImprove?.value.trim() || "";
  const helpful = elements.feedbackHelpful?.value.trim() || "";
  const message = elements.feedbackMessage?.value.trim() || "";
  if (!improve && !helpful && !message) {
    addStatus("Add a note before sending feedback.", "", "error");
    return;
  }

  if (elements.feedbackSend) {
    elements.feedbackSend.disabled = true;
    elements.feedbackSend.textContent = "Sending...";
  }
  try {
    const data = await api("/api/feedback", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        category,
        improve,
        helpful,
        message,
        context: buildFeedbackContext(),
      }),
    });
    if (elements.feedbackCategory) elements.feedbackCategory.value = "other";
    if (elements.feedbackImprove) elements.feedbackImprove.value = "";
    if (elements.feedbackHelpful) elements.feedbackHelpful.value = "";
    if (elements.feedbackMessage) elements.feedbackMessage.value = "";
    closeFeedback();
    if (data?.emailSent === false) {
      addStatus("Feedback saved.", "Email delivery needs follow-up, but the note is stored locally.", "success");
    } else {
      addStatus("Feedback sent.", "Thank you for helping improve the converter.", "success");
    }
  } catch (error) {
    addStatus("Feedback could not be sent.", error.message, "error");
  } finally {
    if (elements.feedbackSend) {
      elements.feedbackSend.disabled = false;
      elements.feedbackSend.textContent = "Send feedback";
    }
  }
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
  const outputReady = outputDestinationReady();
  const completedDuringRun = state.activeConvertIds.length
    ? state.activeConvertIds.filter((id) => {
        const item = state.items.find((candidate) => candidate.id === id);
        return Boolean(item && (item.stage === "complete" || item.stage === "failed"));
      }).length
    : 0;

  elements.queueCountBadge.textContent = `${queuedCount} queued`;
  elements.clearButton.disabled = busy || !hasItems;
  elements.convertButton.disabled = busy || !convertibleCount;
  if (elements.chooseFolderButton) {
    const canChooseFolder = state.capabilities.nativeOutputPicker || browserDirectoryPickerAvailable();
    elements.chooseFolderButton.disabled = busy || !canChooseFolder;
    elements.chooseFolderButton.textContent = canChooseFolder ? "Choose Save Folder" : "Browser Downloads";
  }

  if (busy && state.activeConvertIds.length) {
    elements.convertButton.textContent = `Converting ${completedDuringRun}/${state.activeConvertIds.length}...`;
  } else if (!convertibleCount) {
    elements.convertButton.textContent = "Convert to PDF";
  } else if (!outputReady && !state.capabilities.nativeOutputPicker && !browserDirectoryPickerAvailable()) {
    elements.convertButton.textContent = "Output Folder Required";
  } else {
    elements.convertButton.textContent = `Convert ${convertibleCount} File${convertibleCount === 1 ? "" : "s"} to PDF`;
  }
}

function renderQueue() {
  const queuedCount = state.items.length;
  const summary = summarizeBatch();
  const isTerminalBatch = queuedCount > 0 && summary.activeCount === 0 && !state.isBusy;
  const shouldCollapseTerminalQueue = isTerminalBatch && !state.queueDetailsExpanded;
  const visibleItems = shouldCollapseTerminalQueue && summary.failed > 0
    ? state.items.filter((item) => item.stage === "failed")
    : shouldCollapseTerminalQueue
      ? []
      : state.items;
  refreshFilenameStyleExample();
  elements.queuePanel?.classList.toggle("is-terminal", isTerminalBatch);
  elements.queuePanel?.classList.toggle("is-terminal-collapsed", shouldCollapseTerminalQueue);
  elements.queueEmpty.hidden = queuedCount > 0;
  elements.queueList.hidden = queuedCount === 0;

  if (!queuedCount) {
    elements.queueList.innerHTML = "";
    state.queueDetailsExpanded = false;
    renderOperations();
    updateActionState();
    return;
  }

  const rows = visibleItems
    .map((item, index) => {
      const progress = queueProgressStateForItem(item);
      const visual = queueVisualState(item, progress);
      const canRemove = item.stage !== "complete";
      const canRetry = item.stage === "failed";
      const isRemoving = Boolean(state.pendingRemovalsById[item.id]);
      const isNew = state.recentlyQueuedIds.has(item.id);
      const summary = deriveItemSummary(item, progress);
      const rawPercent = progressPercentForItem(item);
      const percent = Math.round(rawPercent);
      const sizeLabel = formatByteSize(item.sizeBytes);
      const outputDetail = item.stage === "failed"
        ? `<span class="queue-summary-line is-error">${escapeHtml(compactErrorLabel(item.error))}</span>`
        : isNew
          ? `<span class="queue-summary-line"><span class="queue-arrival-chip">Filename prepared</span></span>`
          : `<span class="queue-summary-line">${escapeHtml(formatSource(item.source))}</span>`;

      return `
        <div class="queue-item is-${visual.tone} ${isNew ? "is-new" : ""}" style="--queue-progress: ${rawPercent}; --row-index: ${index}">
          <div class="queue-main">
            <span class="queue-file-glyph is-${visual.tone}" aria-hidden="true">${escapeHtml(queueDocumentLabel(item))}</span>
            <div class="queue-copy">
              <div class="queue-name" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</div>
              <div class="queue-summary-line">${escapeHtml(sizeLabel || formatSource(item.source))}</div>
            </div>
          </div>
          <div class="queue-output">
            <span class="queue-output-glyph" aria-hidden="true"></span>
            <div class="queue-copy">
              <div class="queue-output-name" title="${escapeHtml(outputPreviewLabel(item))}">${escapeHtml(queueOutputName(item))}</div>
              ${outputDetail}
            </div>
          </div>
          <div class="queue-progress-cell">
            <span class="queue-percent">${percent}%</span>
            <div class="queue-progress-track" role="progressbar" aria-label="${escapeHtml(item.name)} conversion progress" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${percent}">
              <span class="queue-progress-fill" style="width: ${rawPercent}%"></span>
            </div>
          </div>
          <div class="queue-actions">
            <span class="queue-state-pill is-${visual.tone}" title="${escapeHtml(summary)}">${escapeHtml(visual.label)}</span>
            ${canRetry ? `<button class="queue-retry" data-retry-id="${escapeHtml(item.id)}" type="button">Retry</button>` : ""}
            ${
              canRemove
                ? `<button class="queue-remove ${isRemoving ? "is-pending" : ""}" data-remove-id="${escapeHtml(item.id)}" type="button" aria-busy="${isRemoving ? "true" : "false"}">${isRemoving ? "Removing..." : "Remove"}</button>`
                : `<span class="queue-complete-badge">Done</span>`
            }
          </div>
        </div>
      `;
    })
    .join("");

  const terminalSummary = buildQueueTerminalSummary(summary, { isTerminalBatch, visibleRows: visibleItems });
  const tableHeader = visibleItems.length
    ? `
      <div class="queue-table-header" aria-hidden="true">
        <span>File Name</span>
        <span>Output PDF</span>
        <span>Progress</span>
        <span>Status</span>
      </div>
    `
    : "";

  elements.queueList.innerHTML = `
    ${terminalSummary}
    ${tableHeader}
    ${rows}
  `;

  renderOperations();
  updateActionState();
}

function applyQueueSnapshot(items) {
  const nextItems = items || [];
  const nextProgress = {};
  const idSet = new Set(nextItems.map((item) => item.id));
  const inferredOutputDir = state.outputDir || parentPath(nextItems.find((item) => item.outputPath)?.outputPath);

  nextItems.forEach((item) => {
    const existing = state.queueProgressByTaskId[item.taskId];
    if (TERMINAL_STAGES.has(item.stage)) {
      if (existing?.active && existing.stage === item.stage && existing.percent < 100) {
        nextProgress[item.taskId] = existing;
        return;
      }
      nextProgress[item.taskId] = buildQueueProgressStateFromSnapshot(item) || existing;
      return;
    }
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
  if (inferredOutputDir && !state.outputDir) {
    state.outputDir = inferredOutputDir;
    state.outputDirLabel = lastPathSegment(inferredOutputDir);
  }
  state.items = nextItems;
  renderQueue();
}

function triggerQueueHandoff() {
  if (!elements.queuePanel) {
    return;
  }
  elements.queuePanel.classList.remove("is-handoff");
  void elements.queuePanel.offsetWidth;
  elements.queuePanel.classList.add("is-handoff");
  if (state.queueHandoffTimer) {
    clearTimeout(state.queueHandoffTimer);
  }
  state.queueHandoffTimer = window.setTimeout(() => {
    elements.queuePanel.classList.remove("is-handoff");
    state.queueHandoffTimer = 0;
  }, 2500);
}

function mergeQueueEvent(payload) {
  if (!payload.taskId) {
    return;
  }

  const outputPath =
    payload.outputPath ||
    (payload.meta && typeof payload.meta.outputPath === "string" ? payload.meta.outputPath : "") ||
    "";
  const outputName =
    payload.outputName ||
    (payload.meta && typeof payload.meta.outputName === "string" ? payload.meta.outputName : "") ||
    "";
  const eventOutputDir =
    (payload.meta && typeof payload.meta.outputDir === "string" ? payload.meta.outputDir : "") ||
    parentPath(outputPath);
  const eventOutputDirLabel =
    (payload.meta && typeof payload.meta.outputDirLabel === "string" ? payload.meta.outputDirLabel : "") ||
    lastPathSegment(eventOutputDir);

  if (eventOutputDir && !state.outputDir) {
    state.outputDir = eventOutputDir;
    state.outputDirLabel = eventOutputDirLabel;
    rememberDestination(eventOutputDir);
    refreshHostingCopy();
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
      outputName: outputName || item.outputName,
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
    const failure = failureExplanation(payload.error);
    return {
      headline: `${fileName} needs attention.`,
      detail: failure.reason,
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

async function refreshFilenameStylePreview() {
  const payload = await api("/api/filename-style-preview", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ filename_style: state.filenameStyle }),
  });
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
  if (!isHostedBrowserOutputMode()) {
    state.browserOutputDirectoryHandle = null;
    state.browserOutputDirLabel = "";
  }
  refreshHostingCopy();
  setDropzoneCopy("default");
  renderStatus();
  renderOperations();
  updateActionState();
}

async function uploadFiles(files, { sourceHint = "upload" } = {}) {
  if (!files?.length) {
    flashDropzone("error", "No MSG files detected", "Drop Outlook emails or .msg files into the workspace.");
    addStatus("No compatible files were detected.", "Drop Outlook emails or .msg files into the workspace.", "error");
    return;
  }
  const fileCount = files.length;
  flashDropzone(
    "pending",
    fileCount === 1 ? "Receiving email..." : `Receiving ${fileCount} emails...`,
    "Reading metadata and preparing PDF names.",
  );
  const formData = new FormData();
  [...files].forEach((file) => formData.append("files", file));
  formData.append("source_hint", sourceHint);
  formData.append("filename_style", state.filenameStyle);
  const [payload] = await Promise.all([
    api("/api/upload", { method: "POST", body: formData }),
    delay(700),
  ]);
  const acceptedIds = new Set((payload.accepted || []).map((item) => item.id).filter(Boolean));
  state.recentlyQueuedIds = acceptedIds;
  state.queueDetailsExpanded = false;
  applyQueueSnapshot(payload.items || []);
  if (payload.accepted?.length) {
    const count = payload.accepted.length;
    flashDropzone(
      "success",
      `${count} email${count === 1 ? "" : "s"} staged`,
      "PDF filenames are ready in the queue.",
    );
    addStatus(
      `Ready to convert: ${count} email${count === 1 ? "" : "s"} staged.`,
      "PDF filenames are prepared in the queue.",
      "success",
    );
    triggerQueueHandoff();
    if (state.newRowsTimer) {
      clearTimeout(state.newRowsTimer);
    }
    state.newRowsTimer = window.setTimeout(() => {
      state.recentlyQueuedIds = new Set();
      state.newRowsTimer = 0;
      renderQueue();
    }, 2600);
  }
  if (payload.rejectedCount) {
    if (!payload.accepted?.length) {
      flashDropzone("error", "No compatible files accepted", "Only Outlook .msg files can be staged.");
    }
    addStatus(
      `${payload.rejectedCount} file(s) were skipped.`,
      `Only .msg files are accepted and the queue limit is ${state.maxFiles}.`,
      "error",
    );
  }
}

async function requestBrowserDirectoryWritePermission(directoryHandle) {
  const options = { mode: "readwrite" };
  if (typeof directoryHandle.queryPermission === "function") {
    const queryResult = await directoryHandle.queryPermission(options);
    if (queryResult === "granted") {
      return true;
    }
  }
  if (typeof directoryHandle.requestPermission === "function") {
    return (await directoryHandle.requestPermission(options)) === "granted";
  }
  return true;
}

async function chooseBrowserOutputFolder({ silentCancel = false } = {}) {
  if (!state.outputDir) {
    if (!silentCancel) {
      addStatus("Server output storage is not configured.", "Conversion cannot start until the hosted output location is available.", "error");
    }
    return false;
  }

  if (!browserDirectoryPickerAvailable()) {
    if (!silentCancel) {
      addStatus("Browser folder selection is unavailable.", "Converted PDFs will be available from the result list.", "neutral");
    }
    return { outputDir: state.outputDir, outputDirLabel: state.outputDirLabel };
  }

  let directoryHandle = null;
  try {
    directoryHandle = await window.showDirectoryPicker({ mode: "readwrite" });
  } catch (error) {
    if (!silentCancel) {
      const isCancel = error?.name === "AbortError";
      addStatus(
        isCancel ? "Save folder selection was cancelled." : "Could not choose the save folder.",
        isCancel ? "Choose a folder when you are ready to convert." : error.message,
        isCancel ? "neutral" : "error",
      );
    }
    return false;
  }

  const hasPermission = await requestBrowserDirectoryWritePermission(directoryHandle);
  if (!hasPermission) {
    if (!silentCancel) {
      addStatus("Folder permission was not granted.", "Choose a folder and allow write access before converting.", "error");
    }
    return false;
  }

  state.browserOutputDirectoryHandle = directoryHandle;
  state.browserOutputDirLabel = directoryHandle.name || "Selected browser folder";
  refreshHostingCopy();
  renderOperations();
  addStatus("Save folder selected.", state.browserOutputDirLabel, "success");
  updateActionState();
  return { outputDir: state.outputDir, outputDirLabel: state.browserOutputDirLabel };
}

async function chooseOutputFolder({ silentCancel = false } = {}) {
  if (isHostedBrowserOutputMode()) {
    return chooseBrowserOutputFolder({ silentCancel });
  }

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
  rememberDestination(state.outputDir);
  refreshHostingCopy();
  renderOperations();
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

  if (state.capabilities.nativeOutputPicker || browserDirectoryPickerAvailable()) {
    const selected = await chooseOutputFolder({ silentCancel: false });
    if (!selected) {
      addStatus("Choose a destination to continue.", "Conversion has not started yet.", "error");
      return;
    }
  } else if (!state.outputDir) {
    addStatus("Output folder is required.", "The hosted deployment needs a server-managed destination.", "error");
    return;
  }

  const ids = convertibleItems.map((item) => item.id);
  state.activeConvertIds = ids;
  state.queueDetailsExpanded = false;
  convertibleItems.forEach((item) => activateQueueProgress(item.taskId));
  renderQueue();
  startQueueProgressLoop();
  addStatus(`Starting conversion for ${ids.length} file(s).`, formatDestinationLabel(), "neutral");

  try {
    const payload = await api("/api/convert", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ids, output_dir: state.outputDir, filename_style: state.filenameStyle }),
    });
    await loadQueue();
    const delivery = await deliverConvertedFiles(ids);
    if (payload.convertedFiles?.length) {
      const detail = delivery.mode === "browser-folder"
        ? `${delivery.saved} local cop${delivery.saved === 1 ? "y was" : "ies were"} saved to ${formatDestinationLabel()}.`
        : delivery.mode === "download"
          ? "Use the Download button beside each saved PDF to keep a local copy."
          : "Your PDFs have been saved.";
      addStatus(`Converted ${payload.convertedFiles.length} file(s).`, detail, "success");
      celebrateQueueCompletion();
    }
    if (payload.errors?.length) {
      payload.errors.forEach((error) => addStatus("A file could not be converted.", failureExplanation(error).reason, "error"));
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

  if (!outputDestinationReady()) {
    const selected = await chooseOutputFolder({ silentCancel: false });
    if (!selected) {
      addStatus("Choose a destination before retrying.", "Retry has not started yet.", "error");
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
      body: JSON.stringify({ ids: [id], output_dir: state.outputDir, filename_style: state.filenameStyle }),
    });
    await loadQueue();
    const delivery = await deliverConvertedFiles([id]);
    if (payload.convertedFiles?.length) {
      const detail = delivery.mode === "browser-folder"
        ? `${target.name} was converted and copied to ${formatDestinationLabel()}.`
        : `${target.name} was converted successfully.`;
      addStatus("Retry succeeded.", detail, "success");
      celebrateQueueCompletion();
    }
    if (payload.errors?.length) {
      payload.errors.forEach((error) => addStatus("Retry failed.", failureExplanation(error).reason, "error"));
    }
  } finally {
    state.activeConvertIds = [];
    updateActionState();
  }
}

async function retryFailedItems() {
  const failedItems = state.items.filter((item) => item.stage === "failed");
  if (!failedItems.length) {
    addStatus("No failed files to retry.", "Every visible file is already clear.", "neutral");
    return;
  }

  if (!outputDestinationReady()) {
    const selected = await chooseOutputFolder({ silentCancel: false });
    if (!selected) {
      addStatus("Choose a destination before retrying.", "Retry has not started yet.", "error");
      return;
    }
  }

  const ids = failedItems.map((item) => item.id);
  state.activeConvertIds = ids;
  failedItems.forEach((item) => activateQueueProgress(item.taskId));
  renderQueue();
  startQueueProgressLoop();
  addStatus(`Retrying ${ids.length} failed file${ids.length === 1 ? "" : "s"}.`, "Only failed rows are being converted again.", "neutral");

  try {
    const payload = await api("/api/convert", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ids, output_dir: state.outputDir, filename_style: state.filenameStyle }),
    });
    await loadQueue();
    const delivery = await deliverConvertedFiles(ids);
    if (payload.convertedFiles?.length) {
      const detail = delivery.mode === "browser-folder"
        ? `Recovered PDFs were copied to ${formatDestinationLabel()}.`
        : "Recovered PDFs have been saved.";
      addStatus(`Recovered ${payload.convertedFiles.length} file(s).`, detail, "success");
      celebrateQueueCompletion();
    }
    if (payload.errors?.length) {
      payload.errors.forEach((error) => addStatus("A failed file still needs attention.", failureExplanation(error).reason, "error"));
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
  renderOperations();
  try {
    await task();
  } finally {
    state.isBusy = false;
    elements.shell.dataset.busy = "false";
    updateActionState();
    renderOperations();
  }
}

function installQueueEvents() {
  elements.queueList.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    const detailsToggle = target.closest("[data-queue-details-toggle]");
    if (detailsToggle instanceof HTMLButtonElement) {
      state.queueDetailsExpanded = !state.queueDetailsExpanded;
      renderQueue();
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

  elements.historyList?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }
    const button = target.closest("[data-destination]");
    if (!(button instanceof HTMLButtonElement)) {
      return;
    }
    const selectedDestination = button.dataset.destination || "";
    if (!selectedDestination) {
      return;
    }
    state.outputDir = selectedDestination;
    state.outputDirLabel = selectedDestination;
    refreshHostingCopy();
    renderOperations();
    updateActionState();
    addStatus("Save destination restored.", selectedDestination, "success");
  });
}

function setDropzoneCopy(mode) {
  if (!elements.dropzoneCopy) {
    return;
  }
  elements.dropzone.classList.remove("is-outlook");
  if (!state.capabilities.outlookImport) {
    elements.dropzoneCopy.textContent = SERVER_DROP_COPY;
    return;
  }
  if (mode === "outlook") {
    elements.dropzoneCopy.textContent = "Drop Outlook emails here. The queue will keep their source visible.";
    elements.dropzone.classList.add("is-outlook");
    return;
  }
  if (mode === "upload") {
    elements.dropzoneCopy.textContent = "Drop .msg files here. The batch preview will show the PDF output naming.";
    return;
  }
  elements.dropzoneCopy.textContent = DEFAULT_DROP_COPY;
}

function previewDropIntent(sourceHint) {
  const mode = sourceHint === "outlook" ? "outlook" : "upload";
  setDropzoneCopy(mode);
  if (elements.dropFeedbackTitle) {
    elements.dropFeedbackTitle.textContent = mode === "outlook" ? "Release to stage emails" : "Release to stage files";
  }
  if (elements.dropFeedbackDetail) {
    elements.dropFeedbackDetail.textContent = mode === "outlook"
      ? "Outlook is preparing the drag. Keep holding and release here."
      : "The queue will inspect the dropped .msg files.";
  }
}

function clearDropIntentPreview() {
  setDropzoneCopy("default");
  if (elements.dropFeedbackTitle) {
    elements.dropFeedbackTitle.textContent = "Email staged";
  }
  if (elements.dropFeedbackDetail) {
    elements.dropFeedbackDetail.textContent = "Date-prefixed PDF name prepared.";
  }
}

function installDropzoneEvents() {
  dropzoneController = createDropzoneController({
    canvas: elements.dropRippleCanvas,
    dropzone: elements.dropzone,
    onDragIntent: ({ sourceHint } = {}) => previewDropIntent(sourceHint),
    onDragEnd: clearDropIntentPreview,
    onDrop: (files, options) => runBusy(() => uploadFiles(files, options)),
    onError: (error) => {
      clearDropzoneFeedback();
      flashDropzone("error", "Upload failed", error.message);
      addStatus("Upload failed.", error.message, "error");
    },
    onFinally: () => {
      setDropzoneCopy("default");
    },
    sourceHintFromDrop: detectUploadSourceFromDrop,
  });
}

function installFeedbackEvents() {
  const openOutputFolderFromButton = () => {
    openOutputFolder().catch((error) => {
      addStatus("The output folder could not be opened.", error.message, "error");
    });
  };
  elements.openOutputFolderButton?.addEventListener("click", openOutputFolderFromButton);
  elements.resultOpenOutputButton?.addEventListener("click", openOutputFolderFromButton);
  elements.resultReviewList?.addEventListener("click", async (event) => {
    const clickTarget = event.target instanceof Element ? event.target : null;
    const downloadButton = clickTarget?.closest("[data-result-download-id]");
    if (downloadButton instanceof HTMLButtonElement && !state.isBusy) {
      downloadButton.disabled = true;
      try {
        await downloadOutputFile(downloadButton.dataset.resultDownloadId, downloadButton.dataset.resultDownloadName);
        setRevealButtonFeedback(downloadButton, { label: "Downloaded", tone: "confirmed" });
      } catch (error) {
        setRevealButtonFeedback(downloadButton, { label: "Unavailable", tone: "error" });
        addStatus("The PDF could not be downloaded.", error.message, "error");
      } finally {
        window.setTimeout(() => {
          downloadButton.disabled = false;
          setRevealButtonFeedback(downloadButton, { label: "Download", tone: "" });
        }, 1800);
      }
      return;
    }

    const revealButton = clickTarget?.closest("[data-result-reveal-path]");
    if (!revealButton || state.isBusy) {
      return;
    }
    revealButton.disabled = true;
    try {
      await revealOutputFile(revealButton.dataset.resultRevealPath, revealButton.dataset.resultRevealName);
      revealButton.disabled = false;
      setRevealButtonFeedback(revealButton, { label: "Shown", tone: "confirmed" });
    } catch (error) {
      revealButton.disabled = false;
      setRevealButtonFeedback(revealButton, { label: "Unavailable", tone: "error" });
      addStatus("The PDF could not be shown.", error.message, "error");
    } finally {
      window.setTimeout(() => {
        revealButton.disabled = false;
        setRevealButtonFeedback(revealButton, { label: "Show", tone: "" });
      }, 1800);
    }
  });
  elements.resultRetryFailedButton?.addEventListener("click", async () => {
    if (state.isBusy) {
      addStatus("Please wait for the current conversion to finish.", "Then retry failed files if any remain.", "neutral");
      return;
    }
    try {
      await runBusy(retryFailedItems);
    } catch (error) {
      addStatus("Could not retry failed files.", error.message, "error");
    }
  });
  elements.resultStartNewButton?.addEventListener("click", async () => {
    if (state.isBusy) {
      addStatus("Please wait for the current conversion to finish.", "Then start a new batch.", "neutral");
      return;
    }
    try {
      await runBusy(clearQueue);
    } catch (error) {
      addStatus("Could not start a new batch.", error.message, "error");
    }
  });
  elements.tutorialButton?.addEventListener("click", () => {
    if (state.tutorialOpen) {
      closeTutorial();
      return;
    }
    openTutorial();
  });
  elements.tutorialClose?.addEventListener("click", closeTutorial);
  elements.tutorialBackdrop?.addEventListener("click", closeTutorial);
  elements.feedbackButton?.addEventListener("click", () => {
    if (state.feedbackOpen) {
      closeFeedback();
      return;
    }
    openFeedback();
  });
  elements.feedbackClose?.addEventListener("click", closeFeedback);
  elements.feedbackBackdrop?.addEventListener("click", closeFeedback);
  elements.feedbackSend?.addEventListener("click", submitFeedback);
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && state.tutorialOpen) {
      closeTutorial();
    }
    if (event.key === "Escape" && state.feedbackOpen) {
      closeFeedback();
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
  elements.filenameStyleSelect?.addEventListener("change", async () => {
    state.filenameStyle = normalizeFilenameStyle(elements.filenameStyleSelect.value);
    saveFilenameStyle();
    refreshFilenameStyleCopy();
    if (!state.items.some((item) => item.stage !== "complete")) {
      return;
    }
    try {
      await runBusy(refreshFilenameStylePreview);
      addStatus("Filename style updated.", "Queued PDF names were refreshed.", "success");
    } catch (error) {
      addStatus("Could not refresh filename previews.", error.message, "error");
    }
  });

  elements.fileInput.addEventListener("change", async () => {
    try {
      await runBusy(() => uploadFiles(elements.fileInput.files));
      elements.fileInput.value = "";
    } catch (error) {
      clearDropzoneFeedback();
      flashDropzone("error", "Upload failed", error.message);
      addStatus("Upload failed.", error.message, "error");
    }
  });

  elements.chooseFolderButton?.addEventListener("click", async () => {
    try {
      await runBusy(() => chooseOutputFolder({ silentCancel: false }));
    } catch (error) {
      addStatus("Could not choose the save folder.", error.message, "error");
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
  state.recentDestinations = loadRecentDestinations();
  state.filenameStyle = loadFilenameStyle();
  refreshFilenameStyleCopy();
  renderStatus();
  renderTutorialModal();
  renderFeedbackModal();
  renderOperations();
  setDropzoneCopy("default");
  renderQueue();
  installQueueEvents();
  installDropzoneEvents();
  installFeedbackEvents();
  installActionEvents();
  connectEvents();
  await loadSettings();
  await loadHealth();
  await loadQueue();
  if (state.filenameStyle !== DEFAULT_FILENAME_STYLE && state.items.some((item) => item.stage !== "complete")) {
    await refreshFilenameStylePreview();
  }
}

bootstrap().catch((error) => {
  addStatus("Could not start the browser UI.", error.message, "error");
});
