import type { EventBatchResponse, EventSourceStatus, TaskEvent } from "./types";

type EventsListener = (events: TaskEvent[]) => void;
type StatusListener = (status: EventSourceStatus) => void;

function isTaskEvent(value: unknown): value is TaskEvent {
  if (typeof value !== "object" || value === null) {
    return false;
  }

  const candidate = value as Partial<TaskEvent>;
  return (
    typeof candidate.taskId === "string" &&
    typeof candidate.taskType === "string" &&
    typeof candidate.stage === "string" &&
    typeof candidate.timestamp === "string"
  );
}

export class JsonlEventSource {
  private readonly eventsListeners = new Set<EventsListener>();
  private readonly statusListeners = new Set<StatusListener>();
  private afterOffset = 0;
  private running = false;
  private timerId: number | null = null;
  private status: EventSourceStatus = "connecting";

  constructor(
    private readonly endpoint = "/api/events",
    private readonly pollIntervalMs = 700
  ) {}

  onEvents(listener: EventsListener): () => void {
    this.eventsListeners.add(listener);
    return () => this.eventsListeners.delete(listener);
  }

  onStatus(listener: StatusListener): () => void {
    this.statusListeners.add(listener);
    listener(this.status);
    return () => this.statusListeners.delete(listener);
  }

  start(): void {
    if (this.running) {
      return;
    }
    this.running = true;
    void this.poll();
  }

  stop(): void {
    this.running = false;
    if (this.timerId !== null) {
      window.clearTimeout(this.timerId);
      this.timerId = null;
    }
  }

  private setStatus(nextStatus: EventSourceStatus): void {
    if (this.status === nextStatus) {
      return;
    }
    this.status = nextStatus;
    for (const listener of this.statusListeners) {
      listener(nextStatus);
    }
  }

  private scheduleNext(delayMs: number): void {
    if (!this.running) {
      return;
    }
    this.timerId = window.setTimeout(() => {
      void this.poll();
    }, delayMs);
  }

  private async poll(): Promise<void> {
    if (!this.running) {
      return;
    }

    try {
      const response = await fetch(`${this.endpoint}?after=${this.afterOffset}`, {
        cache: "no-store"
      });
      if (!response.ok) {
        throw new Error(`Event poll failed with ${response.status}`);
      }

      const payload = (await response.json()) as Partial<EventBatchResponse>;
      const events = Array.isArray(payload.events)
        ? payload.events.filter(isTaskEvent)
        : [];
      const nextOffset =
        typeof payload.nextOffset === "number"
          ? payload.nextOffset
          : this.afterOffset + events.length;

      this.afterOffset = Math.max(this.afterOffset, nextOffset);
      this.setStatus("live");

      if (events.length > 0) {
        for (const listener of this.eventsListeners) {
          listener(events);
        }
      }

      this.scheduleNext(events.length > 0 ? 160 : this.pollIntervalMs);
    } catch (error) {
      console.error("Failed to poll task events", error);
      this.setStatus("error");
      this.scheduleNext(Math.max(this.pollIntervalMs, 1600));
    }
  }
}
