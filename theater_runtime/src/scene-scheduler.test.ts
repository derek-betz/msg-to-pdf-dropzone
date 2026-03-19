import { describe, expect, it } from "vitest";

import { SceneScheduler, collapsePendingEvents, type SceneAdapter } from "./scene-scheduler";
import { TaskStore } from "./task-store";
import type { TaskEvent, TaskSnapshot } from "./types";

function event(
  taskId: string,
  stage: TaskEvent["stage"],
  offsetSeconds: number,
  overrides: Partial<TaskEvent> = {}
): TaskEvent {
  return {
    taskId,
    taskType: "msg-to-pdf",
    stage,
    timestamp: new Date(Date.UTC(2026, 2, 19, 12, 0, offsetSeconds)).toISOString(),
    ...overrides
  };
}

function flushScheduler(): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, 0);
  });
}

describe("collapsePendingEvents", () => {
  it("keeps the readable beats when backlog is present", () => {
    const events: TaskEvent[] = [
      event("task-1", "drop_received", 0),
      event("task-1", "outlook_extract_started", 1),
      event("task-1", "files_accepted", 2),
      event("task-1", "output_folder_selected", 3),
      event("task-1", "parse_started", 4),
      event("task-1", "filename_built", 5),
      event("task-1", "pdf_pipeline_started", 6),
      event("task-1", "pipeline_selected", 7, { pipeline: "outlook_edge" }),
      event("task-1", "pdf_written", 8),
      event("task-1", "deliver_started", 9),
      event("task-1", "complete", 10, { success: true })
    ];

    const collapsed = collapsePendingEvents(events, 2);

    expect(collapsed.map((item) => item.stage)).toEqual([
      "drop_received",
      "files_accepted",
      "output_folder_selected",
      "parse_started",
      "filename_built",
      "pdf_pipeline_started",
      "pipeline_selected",
      "pdf_written",
      "deliver_started",
      "complete"
    ]);
  });
});

describe("SceneScheduler", () => {
  it("keeps one hero task on stage and advances to the next queued task", async () => {
    const calls: {
      heroes: Array<string | null>;
      applied: Array<string[]>;
      queueSizes: number[];
      cleared: number;
    } = {
      heroes: [],
      applied: [],
      queueSizes: [],
      cleared: 0
    };

    const scene: SceneAdapter = {
      setHeroTask(snapshot: TaskSnapshot | null) {
        calls.heroes.push(snapshot?.taskId ?? null);
      },
      setQueue(tasks) {
        calls.queueSizes.push(tasks.length);
      },
      async applyEvents(events) {
        calls.applied.push(events.map((item) => item.stage));
      },
      async holdTerminal() {
        return;
      },
      clearHeroTask() {
        calls.cleared += 1;
      }
    };

    const store = new TaskStore();
    const scheduler = new SceneScheduler(scene, store);

    scheduler.accept([
      event("task-1", "drop_received", 0, { fileName: "one.msg" }),
      event("task-1", "files_accepted", 1, {
        meta: { batchId: "msg-batch-1", batchSize: 2, batchIndex: 1 }
      }),
      event("task-1", "pdf_pipeline_started", 2),
      event("task-1", "pipeline_selected", 3, { pipeline: "edge_html" }),
      event("task-1", "pdf_written", 4),
      event("task-1", "deliver_started", 5),
      event("task-1", "complete", 6, { success: true }),
      event("task-2", "drop_received", 7, { fileName: "two.msg" }),
      event("task-2", "files_accepted", 8, {
        meta: { batchId: "msg-batch-1", batchSize: 2, batchIndex: 2 }
      }),
      event("task-2", "pdf_pipeline_started", 9),
      event("task-2", "pipeline_selected", 10, { pipeline: "reportlab" }),
      event("task-2", "pdf_written", 11),
      event("task-2", "deliver_started", 12),
      event("task-2", "complete", 13, { success: true })
    ]);

    await flushScheduler();
    await flushScheduler();

    expect(calls.heroes).toContain("task-1");
    expect(calls.heroes).toContain("task-2");
    expect(calls.applied[0]).toContain("complete");
    expect(calls.applied[1]).toContain("complete");
    expect(calls.cleared).toBe(1);
    expect(calls.queueSizes.some((size) => size >= 1)).toBe(true);
  });

  it("keeps the current batch on stage while waiting for the next file to arrive", async () => {
    const calls: {
      heroes: Array<string | null>;
      cleared: number;
    } = {
      heroes: [],
      cleared: 0
    };

    const scene: SceneAdapter = {
      setHeroTask(snapshot: TaskSnapshot | null) {
        calls.heroes.push(snapshot?.taskId ?? null);
      },
      setQueue() {
        return;
      },
      async applyEvents() {
        return;
      },
      async holdTerminal() {
        return;
      },
      clearHeroTask() {
        calls.cleared += 1;
      }
    };

    const store = new TaskStore();
    const scheduler = new SceneScheduler(scene, store);

    scheduler.accept([
      event("task-1", "drop_received", 0, { fileName: "one.msg" }),
      event("task-1", "files_accepted", 1, {
        meta: { batchId: "msg-batch-2", batchSize: 2, batchIndex: 1 }
      }),
      event("task-1", "pdf_pipeline_started", 2),
      event("task-1", "pipeline_selected", 3, { pipeline: "edge_html" }),
      event("task-1", "pdf_written", 4),
      event("task-1", "deliver_started", 5),
      event("task-1", "complete", 6, { success: true })
    ]);

    await flushScheduler();
    await flushScheduler();

    expect(calls.heroes).toContain("task-1");
    expect(calls.heroes).not.toContain(null);
    expect(calls.cleared).toBe(0);

    scheduler.accept([
      event("task-2", "drop_received", 7, { fileName: "two.msg" }),
      event("task-2", "files_accepted", 8, {
        meta: { batchId: "msg-batch-2", batchSize: 2, batchIndex: 2 }
      }),
      event("task-2", "pdf_pipeline_started", 9),
      event("task-2", "pipeline_selected", 10, { pipeline: "reportlab" }),
      event("task-2", "pdf_written", 11),
      event("task-2", "deliver_started", 12),
      event("task-2", "complete", 13, { success: true })
    ]);

    await flushScheduler();
    await flushScheduler();

    expect(calls.heroes).toContain("task-2");
    expect(calls.cleared).toBe(1);
  });

  it("does not reset the scene when the next batch file arrives before its batch metadata", async () => {
    const calls: {
      heroes: Array<string | null>;
      cleared: number;
      applied: Array<string[]>;
    } = {
      heroes: [],
      cleared: 0,
      applied: []
    };

    const scene: SceneAdapter = {
      setHeroTask(snapshot: TaskSnapshot | null) {
        calls.heroes.push(snapshot?.taskId ?? null);
      },
      setQueue() {
        return;
      },
      async applyEvents(events) {
        calls.applied.push(events.map((item) => item.stage));
        return;
      },
      async holdTerminal() {
        return;
      },
      clearHeroTask() {
        calls.cleared += 1;
      }
    };

    const store = new TaskStore();
    const scheduler = new SceneScheduler(scene, store);

    scheduler.accept([
      event("task-1", "drop_received", 0, { fileName: "one.msg" }),
      event("task-1", "files_accepted", 1, {
        meta: { batchId: "msg-batch-3", batchSize: 2, batchIndex: 1 }
      }),
      event("task-1", "pdf_pipeline_started", 2),
      event("task-1", "pipeline_selected", 3, { pipeline: "edge_html" }),
      event("task-1", "pdf_written", 4),
      event("task-1", "deliver_started", 5),
      event("task-1", "complete", 6, { success: true })
    ]);

    await flushScheduler();
    await flushScheduler();

    scheduler.accept([event("task-2", "drop_received", 7, { fileName: "two.msg" })]);

    await flushScheduler();
    await flushScheduler();

    expect(calls.heroes).toEqual(["task-1"]);
    expect(calls.cleared).toBe(0);
    expect(calls.applied).toHaveLength(1);
    expect(calls.applied[0]).toContain("complete");

    scheduler.accept([
      event("task-2", "files_accepted", 8, {
        meta: { batchId: "msg-batch-3", batchSize: 2, batchIndex: 2 }
      }),
      event("task-2", "pdf_pipeline_started", 9),
      event("task-2", "pipeline_selected", 10, { pipeline: "reportlab" }),
      event("task-2", "pdf_written", 11),
      event("task-2", "deliver_started", 12),
      event("task-2", "complete", 13, { success: true })
    ]);

    await flushScheduler();
    await flushScheduler();

    expect(calls.heroes.slice(0, 2)).toEqual(["task-1", "task-2"]);
    expect(calls.heroes.at(-1)).toBe(null);
    expect(calls.cleared).toBe(1);
    expect(calls.applied).toHaveLength(2);
    expect(calls.applied[1]).toEqual([
      "drop_received",
      "files_accepted",
      "pdf_pipeline_started",
      "pipeline_selected",
      "pdf_written",
      "deliver_started",
      "complete"
    ]);
  });

  it("does not clear the scene when the next file's drop event arrives before the prior file finishes", async () => {
    const calls: {
      heroes: Array<string | null>;
      cleared: number;
      applied: Array<string[]>;
    } = {
      heroes: [],
      cleared: 0,
      applied: []
    };

    const scene: SceneAdapter = {
      setHeroTask(snapshot: TaskSnapshot | null) {
        calls.heroes.push(snapshot?.taskId ?? null);
      },
      setQueue() {
        return;
      },
      async applyEvents(events) {
        calls.applied.push(events.map((item) => item.stage));
      },
      async holdTerminal() {
        return;
      },
      clearHeroTask() {
        calls.cleared += 1;
      }
    };

    const store = new TaskStore();
    const scheduler = new SceneScheduler(scene, store);

    scheduler.accept([
      event("task-1", "drop_received", 0, { fileName: "one.msg" }),
      event("task-1", "files_accepted", 1, {
        meta: { batchId: "msg-batch-4", batchSize: 2, batchIndex: 1 }
      }),
      event("task-1", "pdf_pipeline_started", 2),
      event("task-1", "pipeline_selected", 3, { pipeline: "edge_html" }),
      event("task-1", "pdf_written", 4),
      event("task-1", "deliver_started", 5),
      event("task-1", "complete", 6, { success: true }),
      event("task-2", "drop_received", 7, { fileName: "two.msg" })
    ]);

    await flushScheduler();
    await flushScheduler();

    expect(calls.heroes).toEqual(["task-1"]);
    expect(calls.cleared).toBe(0);
    expect(calls.applied).toHaveLength(1);

    scheduler.accept([
      event("task-2", "files_accepted", 8, {
        meta: { batchId: "msg-batch-4", batchSize: 2, batchIndex: 2 }
      }),
      event("task-2", "pdf_pipeline_started", 9),
      event("task-2", "pipeline_selected", 10, { pipeline: "reportlab" }),
      event("task-2", "pdf_written", 11),
      event("task-2", "deliver_started", 12),
      event("task-2", "complete", 13, { success: true })
    ]);

    await flushScheduler();
    await flushScheduler();

    expect(calls.heroes.slice(0, 2)).toEqual(["task-1", "task-2"]);
    expect(calls.cleared).toBe(1);
    expect(calls.applied).toHaveLength(2);
  });
});
