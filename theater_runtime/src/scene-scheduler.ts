import { IMPORTANT_STAGES, type TaskEvent, type TaskSnapshot } from "./types";
import type { TaskStore } from "./task-store";

export interface AnimationOptions {
  compressed: boolean;
  backlogCount: number;
  speed: number;
}

export interface SceneAdapter {
  setHeroTask(snapshot: TaskSnapshot | null): void;
  setQueue(tasks: TaskSnapshot[], overflowCount: number): void;
  applyEvents(
    events: TaskEvent[],
    snapshot: TaskSnapshot,
    options: AnimationOptions
  ): Promise<void>;
  holdTerminal(snapshot: TaskSnapshot, options: AnimationOptions): Promise<void>;
  clearHeroTask(): void;
}

export function collapsePendingEvents(
  events: TaskEvent[],
  backlogCount: number
): TaskEvent[] {
  if (events.length <= 2 && backlogCount === 0) {
    return events;
  }

  if (events.length <= 4 && backlogCount === 0) {
    return events;
  }

  const keepIndexes = new Set<number>([0, events.length - 1]);

  events.forEach((event, index) => {
    if (IMPORTANT_STAGES.has(event.stage)) {
      keepIndexes.add(index);
    }
  });

  return events.filter((_, index) => keepIndexes.has(index));
}

export class SceneScheduler {
  private heroTaskId: string | null = null;
  private readonly presentedEventCounts = new Map<string, number>();
  private readonly clearedTaskIds = new Set<string>();
  private awaitingBatchId: string | null = null;
  private draining = false;
  private destroyed = false;

  constructor(
    private readonly scene: SceneAdapter,
    private readonly store: TaskStore
  ) {}

  accept(events: TaskEvent[]): void {
    if (events.length === 0) {
      return;
    }
    this.store.ingest(events);
    this.updateQueue();
    void this.drain();
  }

  stop(): void {
    this.destroyed = true;
  }

  private updateQueue(): void {
    const waiting = this.store
      .getOrderedTasks()
      .filter(
        (task) =>
          !this.clearedTaskIds.has(task.taskId) && task.taskId !== this.heroTaskId
      );
    this.scene.setQueue(waiting.slice(0, 10), Math.max(0, waiting.length - 10));
  }

  private backlogCount(): number {
    return this.store
      .getOrderedTasks()
      .filter(
        (task) =>
          !this.clearedTaskIds.has(task.taskId) && task.taskId !== this.heroTaskId
      ).length;
  }

  private pickNextHero(): TaskSnapshot | null {
    return (
      this.store.getOrderedTasks().find((task) => {
        if (this.clearedTaskIds.has(task.taskId)) {
          return false;
        }
        return task.taskId !== this.heroTaskId;
      }) ?? null
    );
  }

  private peekNextHeroAfterCurrent(): TaskSnapshot | null {
    return this.store.getOrderedTasks().find((task) => {
      if (this.clearedTaskIds.has(task.taskId)) {
        return false;
      }
      return task.taskId !== this.heroTaskId;
    }) ?? null;
  }

  private isSameBatch(left: TaskSnapshot | null, right: TaskSnapshot | null): boolean {
    if (!left || !right) {
      return false;
    }
    if (!left.batchId || !right.batchId) {
      return false;
    }
    return left.batchId === right.batchId;
  }

  private isAwaitingBatchContinuation(snapshot: TaskSnapshot): boolean {
    if (!snapshot.batchId) {
      return false;
    }
    if (
      typeof snapshot.batchIndex !== "number" ||
      typeof snapshot.batchSize !== "number"
    ) {
      return false;
    }
    return snapshot.batchIndex < snapshot.batchSize;
  }

  private isBatchContinuationWithoutMeta(snapshot: TaskSnapshot): boolean {
    return this.awaitingBatchId !== null && !snapshot.batchId;
  }

  private buildAnimationOptions(
    snapshot: TaskSnapshot,
    pendingEvents: TaskEvent[],
    collapsedEvents: TaskEvent[]
  ): AnimationOptions {
    const backlogCount = this.backlogCount();
    const compressed =
      backlogCount > 0 ||
      collapsedEvents.length < pendingEvents.length ||
      snapshot.terminal;

    let speed = 1;
    if (backlogCount >= 2) {
      speed = 0.52;
    } else if (compressed) {
      speed = 0.68;
    }

    return {
      compressed,
      backlogCount,
      speed
    };
  }

  private async drain(): Promise<void> {
    if (this.draining || this.destroyed) {
      return;
    }

    this.draining = true;
    try {
      while (!this.destroyed) {
        if (this.heroTaskId === null) {
          const nextHero = this.pickNextHero();
          if (nextHero === null) {
            if (this.awaitingBatchId !== null) {
              this.updateQueue();
              return;
            }
            this.scene.setHeroTask(null);
            this.updateQueue();
            return;
          }
          this.heroTaskId = nextHero.taskId;
          if (!this.isBatchContinuationWithoutMeta(nextHero)) {
            this.awaitingBatchId = null;
            this.scene.setHeroTask(nextHero);
          }
          this.updateQueue();
        }

        const snapshot = this.heroTaskId
          ? this.store.getTask(this.heroTaskId)
          : undefined;
        if (!snapshot) {
          this.heroTaskId = null;
          continue;
        }

        if (
          this.awaitingBatchId !== null &&
          snapshot.batchId === this.awaitingBatchId
        ) {
          this.awaitingBatchId = null;
          this.scene.setHeroTask(snapshot);
        }

        if (this.isBatchContinuationWithoutMeta(snapshot)) {
          return;
        }

        const presentedCount =
          this.presentedEventCounts.get(snapshot.taskId) ?? 0;
        const pendingEvents = snapshot.events.slice(presentedCount);

        if (pendingEvents.length === 0) {
          if (snapshot.terminal) {
            const options = this.buildAnimationOptions(snapshot, [], []);
            await this.scene.holdTerminal(snapshot, options);
            this.clearedTaskIds.add(snapshot.taskId);
            const nextHero = this.peekNextHeroAfterCurrent();
            const awaitingContinuation =
              this.isAwaitingBatchContinuation(snapshot) &&
              !this.isSameBatch(snapshot, nextHero);
            this.heroTaskId = null;
            this.awaitingBatchId = awaitingContinuation
              ? snapshot.batchId ?? null
              : null;
            if (!awaitingContinuation && !this.isSameBatch(snapshot, nextHero)) {
              this.scene.clearHeroTask();
            }
            this.updateQueue();
            continue;
          }
          return;
        }

        const collapsedEvents = collapsePendingEvents(
          pendingEvents,
          this.backlogCount()
        );
        const nextPresentedCount = presentedCount + pendingEvents.length;
        const options = this.buildAnimationOptions(
          snapshot,
          pendingEvents,
          collapsedEvents
        );
        await this.scene.applyEvents(collapsedEvents, snapshot, options);
        this.presentedEventCounts.set(snapshot.taskId, nextPresentedCount);
      }
    } finally {
      this.draining = false;
    }
  }
}
