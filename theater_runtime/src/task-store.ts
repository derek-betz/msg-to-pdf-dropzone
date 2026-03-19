import { TERMINAL_STAGES, type TaskEvent, type TaskMetaValue, type TaskSnapshot } from "./types";

export class TaskStore {
  private readonly tasks = new Map<string, TaskSnapshot>();
  private readonly order: string[] = [];

  ingest(events: TaskEvent[]): TaskSnapshot[] {
    const changedTaskIds = new Set<string>();

    for (const event of events) {
      let snapshot = this.tasks.get(event.taskId);
      if (!snapshot) {
        snapshot = {
          taskId: event.taskId,
          taskType: event.taskType,
          stage: event.stage,
          firstTimestamp: event.timestamp,
          lastTimestamp: event.timestamp,
          meta: {},
          events: [],
          hasPdfWritten: false,
          terminal: false,
          failed: false
        };
        this.tasks.set(event.taskId, snapshot);
        this.order.push(event.taskId);
      }

      snapshot.stage = event.stage;
      snapshot.lastTimestamp = event.timestamp;
      snapshot.events.push(event);

      if (event.fileName) {
        snapshot.fileName = event.fileName;
      }
      if (event.pipeline) {
        snapshot.pipeline = event.pipeline;
      }
      if (event.success !== undefined) {
        snapshot.success = event.success;
      }
      if (event.error) {
        snapshot.error = event.error;
      }
      if (event.meta) {
        Object.assign(snapshot.meta, event.meta);
      }

      const outputName = snapshot.meta.outputName;
      const outputPath = snapshot.meta.outputPath;
      const outputDir = snapshot.meta.outputDir;
      const outputDirLabel = snapshot.meta.outputDirLabel;
      const batchId = snapshot.meta.batchId;
      const batchSize = snapshot.meta.batchSize;
      const batchIndex = snapshot.meta.batchIndex;

      if (typeof outputName === "string") {
        snapshot.outputName = outputName;
      }
      if (typeof outputPath === "string") {
        snapshot.outputPath = outputPath;
      }
      if (typeof outputDir === "string") {
        snapshot.outputDir = outputDir;
      }
      if (typeof outputDirLabel === "string") {
        snapshot.outputDirLabel = outputDirLabel;
      }
      if (typeof batchId === "string") {
        snapshot.batchId = batchId;
      }
      if (typeof batchSize === "number") {
        snapshot.batchSize = batchSize;
      }
      if (typeof batchIndex === "number") {
        snapshot.batchIndex = batchIndex;
      }

      if (event.stage === "pdf_written") {
        snapshot.hasPdfWritten = true;
      }
      if (TERMINAL_STAGES.has(event.stage)) {
        snapshot.terminal = true;
        snapshot.failed = event.stage === "failed";
      }

      changedTaskIds.add(event.taskId);
    }

    return this.order
      .filter((taskId) => changedTaskIds.has(taskId))
      .map((taskId) => this.tasks.get(taskId))
      .filter((snapshot): snapshot is TaskSnapshot => snapshot !== undefined);
  }

  getTask(taskId: string): TaskSnapshot | undefined {
    return this.tasks.get(taskId);
  }

  getOrderedTasks(): TaskSnapshot[] {
    return this.order
      .map((taskId) => this.tasks.get(taskId))
      .filter((snapshot): snapshot is TaskSnapshot => snapshot !== undefined);
  }
}

export function asStringMeta(
  meta: Record<string, TaskMetaValue>,
  key: string
): string | undefined {
  const value = meta[key];
  return typeof value === "string" ? value : undefined;
}
