import "./styles.css";

import { Application, Color, Container, Graphics, Text } from "pixi.js";

import {
  STAGE_LABELS,
  type EventSourceStatus,
  type RenderPipeline,
  type TaskEvent,
  type TaskSnapshot
} from "./types";
import type { AnimationOptions, SceneAdapter } from "./scene-scheduler";

export interface DemoSettings {
  fileCount: number;
  failureCount: number;
}

interface Zone {
  x: number;
  y: number;
  width: number;
  height: number;
  label: string;
}

type BundleLocation = "entry" | "intake" | "station" | "destination" | "error";
type TerminalDisplayLocation = Extract<BundleLocation, "destination" | "error">;
type StationPile = "incoming" | "outgoing";

class StageActor {
  readonly view: Container;
  private baseY: number;
  private bobOffset = Math.random() * Math.PI * 2;

  constructor(name: string, color: number, x: number, y: number) {
    this.view = new Container();
    this.baseY = y;

    const shadow = new Graphics();
    shadow.ellipse(0, 18, 22, 8).fill({ color: 0x05070d, alpha: 0.28 });
    shadow.y = 12;

    const body = new Graphics();
    body.roundRect(-18, -34, 36, 44, 14).fill(color);
    body
      .roundRect(-12, -48, 24, 20, 10)
      .fill(Color.shared.setValue(color).multiply(1.12));
    body.rect(-10, 10, 8, 18).fill(Color.shared.setValue(color).multiply(0.9));
    body.rect(2, 10, 8, 18).fill(Color.shared.setValue(color).multiply(0.9));

    const badge = new Graphics();
    badge.roundRect(-14, -8, 28, 10, 5).fill(0xf7f2e7);

    const label = new Text({
      text: name,
      style: {
        fill: 0xf7f2e7,
        fontFamily: "Georgia, serif",
        fontSize: 11
      }
    });
    label.anchor.set(0.5, 0);
    label.y = 34;

    this.view.position.set(x, y);
    this.view.addChild(shadow, body, badge, label);
  }

  update(deltaSeconds: number): void {
    this.bobOffset += deltaSeconds * 2.2;
    this.view.y = this.baseY + Math.sin(this.bobOffset) * 1.7;
  }

  setPosition(x: number, y: number): void {
    this.view.position.set(x, y);
    this.baseY = y;
  }
}

class DocumentProp {
  readonly view: Container;
  private readonly stack: Graphics;
  private readonly frame: Graphics;
  private readonly label: Text;
  private readonly tag: Text;
  private readonly countBubble: Graphics;
  private readonly countLabel: Text;
  private stackCount = 1;

  constructor(kind: "MSG" | "PDF") {
    this.view = new Container();
    this.stack = new Graphics();
    this.frame = new Graphics();
    this.label = new Text({
      text: kind,
      style: {
        fill: 0x1f2937,
        fontFamily: "Courier New, monospace",
        fontSize: 18,
        fontWeight: "700"
      }
    });
    this.tag = new Text({
      text: "",
      style: {
        fill: 0x9ca3af,
        fontFamily: "Courier New, monospace",
        fontSize: 10
      }
    });
    this.countBubble = new Graphics();
    this.countLabel = new Text({
      text: "",
      style: {
        fill: 0x1a130d,
        fontFamily: "Courier New, monospace",
        fontSize: 10,
        fontWeight: "700"
      }
    });
    this.countLabel.anchor.set(0.5);
    this.countLabel.position.set(60, 10);

    this.view.addChild(
      this.stack,
      this.frame,
      this.label,
      this.tag,
      this.countBubble,
      this.countLabel
    );
    this.setKind(kind);
  }

  setKind(kind: "MSG" | "PDF"): void {
    const paperColor = kind === "MSG" ? 0xfff7e8 : 0xe9f1ff;
    const accent = kind === "MSG" ? 0xdc7b2b : 0x3072f6;

    this.frame.clear();
    this.frame
      .roundRect(0, 0, 72, 92, 10)
      .fill(paperColor)
      .stroke({ color: 0x1d2734, width: 2 });
    this.frame.roundRect(10, 14, 52, 12, 6).fill(accent);
    this.frame.rect(10, 38, 50, 5).fill({ color: 0x475569, alpha: 0.45 });
    this.frame.rect(10, 50, 44, 5).fill({ color: 0x475569, alpha: 0.3 });
    this.frame.rect(10, 62, 36, 5).fill({ color: 0x475569, alpha: 0.22 });

    this.label.text = kind;
    this.label.anchor.set(0.5);
    this.label.position.set(36, 20);

    this.tag.anchor.set(0.5);
    this.tag.position.set(36, 80);
    this.redrawStack();
  }

  setStackCount(count: number): void {
    this.stackCount = Math.max(1, Math.round(count));
    this.redrawStack();
  }

  setTag(text: string): void {
    this.tag.text = text;
  }

  setNormalTag(): void {
    this.tag.style.fill = 0x9ca3af;
  }

  setFailedTag(): void {
    this.tag.text = "FAILED";
    this.tag.style.fill = 0xc73e2d;
  }

  private redrawStack(): void {
    this.stack.clear();

    const extraSheets = Math.min(4, Math.max(0, this.stackCount - 1));
    for (let index = extraSheets; index >= 1; index -= 1) {
      const offset = index * 5;
      const alpha = 0.14 + index * 0.05;
      this.stack
        .roundRect(offset, offset, 72, 92, 10)
        .fill({ color: 0xffffff, alpha })
        .stroke({ color: 0x1d2734, width: 1, alpha: 0.16 });
    }

    if (this.stackCount > 1) {
      this.countBubble.clear();
      this.countBubble.circle(60, 10, 12).fill(0xf4c16f);
      this.countLabel.text = `x${this.stackCount}`;
      this.countLabel.visible = true;
      return;
    }

    this.countBubble.clear();
    this.countLabel.text = "";
    this.countLabel.visible = false;
  }
}

export class WorkshopScene implements SceneAdapter {
  readonly app: Application;

  private readonly root: HTMLElement;
  private readonly stage: Container;
  private readonly view: HTMLCanvasElement;
  private readonly zones: Record<Exclude<BundleLocation, "entry">, Zone>;
  private readonly carrier: StageActor;
  private readonly document: DocumentProp;
  private readonly stationIncoming: DocumentProp;
  private readonly stationOutgoing: DocumentProp;
  private readonly stationTransfer: DocumentProp;
  private readonly errorPile: DocumentProp;
  private readonly station: Graphics;
  private readonly stationGlow: Graphics;
  private readonly destinationLight: Graphics;
  private readonly errorBin: Graphics;
  private readonly currentStageEl: HTMLElement;
  private readonly currentFileEl: HTMLElement;
  private readonly outputNameEl: HTMLElement;
  private readonly pipelineEl: HTMLElement;
  private readonly queueCountEl: HTMLElement;
  private readonly connectionEl: HTMLElement;
  private readonly eventListEl: HTMLElement;
  private readonly demoFormEl: HTMLFormElement;
  private readonly demoFileCountEl: HTMLInputElement;
  private readonly demoFailureCountEl: HTMLInputElement;
  private readonly demoSubmitEl: HTMLButtonElement;
  private readonly demoStatusEl: HTMLElement;

  private lastPipeline: RenderPipeline | null = null;
  private currentBatchId: string | null = null;
  private currentBatchSize = 1;
  private completedInBatch = 0;
  private failedInBatch = 0;
  private convertedInBatch = 0;
  private errorPileCount = 0;
  private bundleLocation: BundleLocation = "entry";
  private terminalDisplayLocation: TerminalDisplayLocation | null = null;
  private stationPilesActive = false;

  constructor(root: HTMLElement) {
    this.root = root;
    this.root.innerHTML = `
      <div class="shell">
        <section class="stage-panel">
          <div class="panel-header">
            <div>
              <p class="eyebrow">Sprite Task Theater</p>
              <h1>msg-to-pdf carrier workshop</h1>
              <p class="panel-copy">One carrier walks each batch from the intake desk to the conversion station, then on to the destination dock or error bin.</p>
            </div>
            <div class="status-cluster">
              <p class="pill connection-pill is-connecting">connecting</p>
              <p class="pill queue-pill">0 queued</p>
            </div>
          </div>
          <div class="canvas-frame">
            <div class="canvas-host"></div>
          </div>
        </section>
        <aside class="side-panel">
          <section class="card">
            <p class="card-label">Demo Controls</p>
            <form class="demo-form">
              <div class="field-grid">
                <label class="field">
                  <span class="field-label">Dummy .msg files</span>
                  <input class="number-input demo-file-count" type="number" min="1" max="10" step="1" value="5" />
                </label>
                <label class="field">
                  <span class="field-label">Files that fail</span>
                  <input class="number-input demo-failure-count" type="number" min="0" max="5" step="1" value="0" />
                </label>
              </div>
              <p class="demo-hint">The last N files fail and finish in the error bin. All earlier files complete normally.</p>
              <div class="button-row">
                <button class="demo-run-button" type="submit">Run Demo</button>
                <p class="demo-status">Ready to stage a fresh batch.</p>
              </div>
            </form>
          </section>
          <section class="card">
            <p class="card-label">Current Stage</p>
            <h2 class="stage-value">Idle</h2>
            <div class="meta-grid">
              <div>
                <p class="card-label">Current File</p>
                <p class="meta-value file-value">Waiting for work</p>
              </div>
              <div>
                <p class="card-label">Pipeline</p>
                <p class="meta-value pipeline-value">none</p>
              </div>
              <div>
                <p class="card-label">Output Name</p>
                <p class="meta-value output-value">not built yet</p>
              </div>
            </div>
          </section>
          <section class="card">
            <p class="card-label">Live Event Feed</p>
            <ol class="event-feed"></ol>
          </section>
        </aside>
      </div>
    `;

    this.currentStageEl = this.query(".stage-value");
    this.currentFileEl = this.query(".file-value");
    this.outputNameEl = this.query(".output-value");
    this.pipelineEl = this.query(".pipeline-value");
    this.queueCountEl = this.query(".queue-pill");
    this.connectionEl = this.query(".connection-pill");
    this.eventListEl = this.query(".event-feed");
    this.demoFormEl = this.query(".demo-form");
    this.demoFileCountEl = this.query(".demo-file-count");
    this.demoFailureCountEl = this.query(".demo-failure-count");
    this.demoSubmitEl = this.query(".demo-run-button");
    this.demoStatusEl = this.query(".demo-status");

    this.app = new Application();
    this.stage = new Container();
    this.zones = {
      intake: { x: 86, y: 286, width: 180, height: 132, label: "Intake Desk" },
      station: { x: 334, y: 230, width: 214, height: 188, label: "Conversion Station" },
      destination: { x: 620, y: 278, width: 166, height: 140, label: "Destination Dock" },
      error: { x: 620, y: 444, width: 170, height: 72, label: "Error Bin" }
    };

    this.carrier = new StageActor("Carrier", 0x2f8f83, 40, 452);
    this.document = new DocumentProp("MSG");
    this.stationIncoming = new DocumentProp("MSG");
    this.stationOutgoing = new DocumentProp("PDF");
    this.stationTransfer = new DocumentProp("MSG");
    this.errorPile = new DocumentProp("MSG");
    this.station = new Graphics();
    this.stationGlow = new Graphics();
    this.destinationLight = new Graphics();
    this.errorBin = new Graphics();

    this.stationIncoming.view.scale.set(0.84);
    this.stationOutgoing.view.scale.set(0.84);
    this.stationTransfer.view.scale.set(0.78);
    this.errorPile.view.scale.set(0.72);

    this.view = document.createElement("canvas");
    this.query(".canvas-host").appendChild(this.view);
  }

  async init(): Promise<void> {
    await this.app.init({
      canvas: this.view,
      width: 880,
      height: 540,
      backgroundAlpha: 0,
      antialias: true,
      preference: "webgl",
      resolution: Math.min(window.devicePixelRatio || 1, 2)
    });

    this.app.stage.addChild(this.stage);
    this.drawBackdrop();
    this.drawZones();
    this.drawStation();
    this.drawDestinationDock();
    this.drawErrorBin();

    this.document.view.visible = false;
    this.hideStationPiles();
    this.resetErrorPile();
    this.stage.addChild(
      this.stationIncoming.view,
      this.stationOutgoing.view,
      this.stationTransfer.view,
      this.errorPile.view,
      this.document.view,
      this.carrier.view
    );

    this.app.ticker.add((ticker) => {
      const deltaSeconds = ticker.deltaMS / 1000;
      this.carrier.update(deltaSeconds);
      this.pulseStation();
    });

    window.addEventListener("resize", () => this.handleResize());
    this.syncDemoFailureLimit();
    this.handleResize();
  }

  bindDemoLauncher(
    launcher: (settings: DemoSettings) => Promise<void> | void
  ): void {
    this.demoFileCountEl.addEventListener("input", () => {
      this.syncDemoFailureLimit();
    });
    this.demoFailureCountEl.addEventListener("input", () => {
      this.syncDemoFailureLimit();
    });
    this.demoFormEl.addEventListener("submit", (event) => {
      event.preventDefault();
      const settings = this.readDemoSettings();
      void launcher(settings);
    });
  }

  setDemoStatus(message: string, tone: "idle" | "busy" | "error" = "idle"): void {
    this.demoStatusEl.textContent = message;
    this.demoStatusEl.className = `demo-status is-${tone}`;
  }

  setDemoBusy(isBusy: boolean): void {
    this.demoSubmitEl.disabled = isBusy;
    this.demoFileCountEl.disabled = isBusy;
    this.demoFailureCountEl.disabled = isBusy;
  }

  setConnectionStatus(status: EventSourceStatus): void {
    this.connectionEl.textContent = status;
    this.connectionEl.className = `pill connection-pill is-${status}`;
  }

  setHeroTask(snapshot: TaskSnapshot | null): void {
    if (!snapshot) {
      this.currentStageEl.textContent = "Idle";
      this.currentFileEl.textContent = "Waiting for work";
      this.outputNameEl.textContent = "not built yet";
      this.pipelineEl.textContent = "none";
      return;
    }

    const nextBatchId = this.resolveBatchId(snapshot);
    if (this.currentBatchId !== nextBatchId) {
      this.resetHeroScene();
      this.currentBatchId = nextBatchId;
      this.currentBatchSize = snapshot.batchSize ?? 1;
      this.completedInBatch = 0;
      this.failedInBatch = 0;
      this.convertedInBatch = 0;
      this.document.setStackCount(this.currentBatchSize);
    } else if (snapshot.batchSize && snapshot.batchSize > this.currentBatchSize) {
      this.currentBatchSize = snapshot.batchSize;
      this.document.setStackCount(this.currentBatchSize);
    }

    this.currentStageEl.textContent = STAGE_LABELS[snapshot.stage];
    this.currentFileEl.textContent = snapshot.fileName ?? snapshot.taskId;
    this.outputNameEl.textContent = snapshot.outputName ?? "not built yet";
    this.pipelineEl.textContent = snapshot.pipeline ?? "pending";
  }

  setQueue(tasks: TaskSnapshot[], overflowCount: number): void {
    this.queueCountEl.textContent = `${tasks.length + overflowCount} queued`;
  }

  async applyEvents(
    events: TaskEvent[],
    snapshot: TaskSnapshot,
    options: AnimationOptions
  ): Promise<void> {
    for (const event of events) {
      await this.applyEvent(event, snapshot, options);
    }
  }

  async holdTerminal(
    snapshot: TaskSnapshot,
    options: AnimationOptions
  ): Promise<void> {
    if (snapshot.failed) {
      await sleep(Math.max(260, Math.round(520 * options.speed)));
      return;
    }

    const holdMs =
      this.bundleLocation === "destination"
        ? Math.max(320, Math.round(760 * options.speed))
        : Math.max(180, Math.round(320 * options.speed));
    await sleep(holdMs);
  }

  clearHeroTask(): void {
    this.currentStageEl.textContent = "Waiting";
    this.currentFileEl.textContent = "Queueing next task";
    this.outputNameEl.textContent = "not built yet";
    this.pipelineEl.textContent = "none";
    const keepTerminalBundleVisible = this.terminalDisplayLocation !== null;
    if (!keepTerminalBundleVisible) {
      this.document.view.visible = false;
    }
    this.currentBatchId = null;
    this.currentBatchSize = 1;
    this.completedInBatch = 0;
    this.failedInBatch = 0;
    this.convertedInBatch = 0;
    this.hideStationPiles();
    if (!keepTerminalBundleVisible) {
      this.bundleLocation = "entry";
    } else if (this.terminalDisplayLocation) {
      const terminalPoint = this.bundlePoint(this.terminalDisplayLocation);
      const terminalStop = this.actorStop(this.terminalDisplayLocation);
      this.document.view.visible = true;
      this.document.view.position.set(terminalPoint.x, terminalPoint.y);
      this.carrier.setPosition(terminalStop.x, terminalStop.y);
      this.bundleLocation = this.terminalDisplayLocation;
    }
    this.destinationLight.visible = false;
    this.lastPipeline = null;
  }

  private query<T extends HTMLElement>(selector: string): T {
    const element = this.root.querySelector(selector);
    if (!(element instanceof HTMLElement)) {
      throw new Error(`Expected element for selector ${selector}`);
    }
    return element as T;
  }

  private handleResize(): void {
    const frame = this.query(".canvas-frame");
    const width = frame.clientWidth - 20;
    const scale = Math.min(width / 880, 1);
    this.view.style.width = `${880 * scale}px`;
    this.view.style.height = `${540 * scale}px`;
  }

  private drawBackdrop(): void {
    const wall = new Graphics();
    wall.rect(0, 0, 880, 540).fill(0x121924);

    const canopy = new Graphics();
    canopy.rect(0, 0, 880, 220).fill({ color: 0x253347, alpha: 0.78 });

    const floor = new Graphics();
    floor.rect(0, 392, 880, 148).fill(0x161d27);
    floor.rect(0, 390, 880, 2).fill({ color: 0xffffff, alpha: 0.08 });

    const glow = new Graphics();
    glow.circle(210, 126, 120).fill({ color: 0xd4894c, alpha: 0.08 });
    glow.circle(430, 96, 128).fill({ color: 0x6d8cff, alpha: 0.08 });

    this.stage.addChild(wall, canopy, glow, floor);
  }

  private drawZones(): void {
    for (const zone of Object.values(this.zones)) {
      const panel = new Graphics();
      panel
        .roundRect(zone.x, zone.y, zone.width, zone.height, 20)
        .fill({
          color: 0xf3e4c7,
          alpha: zone.label === "Error Bin" ? 0.08 : 0.12
        });
      panel.stroke({ color: 0xffffff, alpha: 0.12, width: 1.5 });

      const title = new Text({
        text: zone.label,
        style: {
          fill: 0xf4e7cd,
          fontFamily: "Georgia, serif",
          fontSize: 16,
          fontWeight: "600"
        }
      });
      title.position.set(zone.x + 16, zone.y - 28);
      this.stage.addChild(panel, title);
    }
  }

  private drawStation(): void {
    this.stationGlow.circle(442, 326, 88).fill({ color: 0x4f78ff, alpha: 0.1 });

    this.station
      .roundRect(360, 248, 162, 150, 26)
      .fill(0x243042)
      .stroke({ color: 0xdcb57d, alpha: 0.28, width: 2 });
    this.station
      .roundRect(374, 292, 60, 82, 18)
      .fill({ color: 0xffffff, alpha: 0.06 })
      .stroke({ color: 0xffffff, alpha: 0.08, width: 1.5 });
    this.station
      .roundRect(452, 292, 60, 82, 18)
      .fill({ color: 0xffffff, alpha: 0.06 })
      .stroke({ color: 0xffffff, alpha: 0.08, width: 1.5 });
    this.station
      .circle(442, 322, 44)
      .fill({ color: 0x8ba0d4, alpha: 0.12 })
      .stroke({ color: 0xffffff, alpha: 0.16, width: 2 });
    this.station.circle(442, 322, 18).fill({ color: 0xffffff, alpha: 0.18 });

    const incomingLabel = new Text({
      text: "Incoming",
      style: {
        fill: 0xdcccb0,
        fontFamily: "Courier New, monospace",
        fontSize: 11,
        fontWeight: "700"
      }
    });
    incomingLabel.anchor.set(0.5);
    incomingLabel.position.set(404, 270);

    const outgoingLabel = new Text({
      text: "Outgoing",
      style: {
        fill: 0xc7d7ff,
        fontFamily: "Courier New, monospace",
        fontSize: 11,
        fontWeight: "700"
      }
    });
    outgoingLabel.anchor.set(0.5);
    outgoingLabel.position.set(482, 270);

    this.stage.addChild(this.stationGlow, this.station, incomingLabel, outgoingLabel);
  }

  private drawDestinationDock(): void {
    this.destinationLight
      .roundRect(656, 316, 96, 78, 24)
      .fill({ color: 0x65db8d, alpha: 0.1 });
    this.destinationLight.visible = false;

    const dock = new Graphics();
    dock
      .roundRect(648, 306, 112, 94, 24)
      .fill(0x233426)
      .stroke({ color: 0xa6f0bd, alpha: 0.26, width: 2 });
    dock.roundRect(676, 340, 56, 12, 6).fill({ color: 0xa6f0bd, alpha: 0.2 });

    this.stage.addChild(this.destinationLight, dock);
  }

  private drawErrorBin(): void {
    this.errorBin
      .roundRect(648, 462, 114, 36, 14)
      .fill({ color: 0x5a1f1f, alpha: 0.52 })
      .stroke({ color: 0xffb0a7, alpha: 0.18, width: 2 });
    this.stage.addChild(this.errorBin);
  }

  private pulseStation(): void {
    const strength =
      this.lastPipeline === "outlook_edge"
        ? 0.18
        : this.lastPipeline === "edge_html"
          ? 0.16
          : this.lastPipeline === "reportlab"
            ? 0.14
            : 0.08;
    const time = performance.now() / 1000;
    this.stationGlow.alpha = strength + Math.sin(time * 3.1) * 0.03;
  }

  private resolveBatchId(snapshot: TaskSnapshot): string {
    return snapshot.batchId ?? snapshot.taskId;
  }

  private batchSizeFor(snapshot: TaskSnapshot): number {
    return snapshot.batchSize ?? this.currentBatchSize ?? 1;
  }

  private isFinalBatchTask(snapshot: TaskSnapshot): boolean {
    if (!this.isBatchMode(snapshot)) {
      return true;
    }
    if (
      typeof snapshot.batchIndex === "number" &&
      typeof snapshot.batchSize === "number"
    ) {
      return snapshot.batchIndex >= snapshot.batchSize;
    }
    return this.completedInBatch + this.failedInBatch + 1 >= this.currentBatchSize;
  }

  private isBatchMode(snapshot: TaskSnapshot): boolean {
    return this.batchSizeFor(snapshot) > 1;
  }

  private remainingBatchItems(): number {
    return Math.max(0, this.currentBatchSize - this.completedInBatch - this.failedInBatch);
  }

  private isFinalBatchItem(projectedCompleted = this.completedInBatch): boolean {
    return projectedCompleted + this.failedInBatch >= this.currentBatchSize;
  }

  private conversionHoldLocation(): BundleLocation {
    return this.bundleLocation === "entry" || this.bundleLocation === "intake"
      ? "intake"
      : "station";
  }

  private stationPilePoint(pile: StationPile): { x: number; y: number } {
    if (pile === "incoming") {
      return { x: 382, y: 300 };
    }
    return { x: 462, y: 300 };
  }

  private stationTransferMidPoint(): { x: number; y: number } {
    return { x: 426, y: 284 };
  }

  private bundlePoint(location: BundleLocation): { x: number; y: number } {
    switch (location) {
      case "entry":
        return { x: -52, y: 302 };
      case "intake":
        return { x: this.zones.intake.x + 54, y: this.zones.intake.y + 24 };
      case "station":
        return { x: this.zones.station.x + 72, y: this.zones.station.y + 56 };
      case "destination":
        return { x: this.zones.destination.x + 36, y: this.zones.destination.y + 28 };
      case "error":
        return { x: this.zones.error.x + 20, y: this.zones.error.y - 36 };
    }
  }

  private actorStop(location: BundleLocation): { x: number; y: number } {
    switch (location) {
      case "entry":
        return { x: 34, y: 452 };
      case "intake":
        return { x: this.zones.intake.x + 62, y: 452 };
      case "station":
        return { x: this.zones.station.x + 74, y: 452 };
      case "destination":
        return { x: this.zones.destination.x + 40, y: 452 };
      case "error":
        return { x: this.zones.error.x + 38, y: 452 };
    }
  }

  private incomingPileCount(): number {
    return Math.max(0, this.currentBatchSize - this.failedInBatch - this.convertedInBatch);
  }

  private outgoingPileCount(): number {
    return Math.max(0, this.convertedInBatch);
  }

  private errorPilePoint(): { x: number; y: number } {
    return { x: this.zones.error.x + 48, y: this.zones.error.y - 62 };
  }

  private hideStationPiles(): void {
    this.stationIncoming.view.visible = false;
    this.stationOutgoing.view.visible = false;
    this.stationTransfer.view.visible = false;
    this.stationPilesActive = false;
  }

  private resetErrorPile(): void {
    this.errorPileCount = 0;
    this.errorPile.view.visible = false;
  }

  private syncErrorPile(kind: "MSG" | "PDF" = "MSG"): void {
    if (this.errorPileCount <= 0) {
      this.errorPile.view.visible = false;
      return;
    }

    this.errorPile.setKind(kind);
    this.errorPile.setStackCount(this.errorPileCount);
    this.errorPile.setFailedTag();
    this.errorPile.setTag(`${this.errorPileCount} FAILED`);
    const point = this.errorPilePoint();
    this.stage.addChild(this.errorPile.view);
    this.errorPile.view.position.set(point.x, point.y);
    this.errorPile.view.visible = true;
  }

  private depositErrorFile(kind: "MSG" | "PDF"): void {
    this.errorPileCount += 1;
    this.syncErrorPile(kind);
    this.document.view.visible = false;
  }

  private syncStationPiles(): void {
    if (!this.stationPilesActive) {
      this.hideStationPiles();
      return;
    }

    const incomingCount = this.incomingPileCount();
    const outgoingCount = this.outgoingPileCount();

    if (incomingCount > 0) {
      this.stationIncoming.setKind("MSG");
      this.stationIncoming.setNormalTag();
      this.stationIncoming.setStackCount(incomingCount);
      this.stationIncoming.setTag(`${incomingCount} LEFT`);
      const point = this.stationPilePoint("incoming");
      this.stationIncoming.view.position.set(point.x, point.y);
      this.stationIncoming.view.visible = true;
    } else {
      this.stationIncoming.view.visible = false;
    }

    if (outgoingCount > 0) {
      this.stationOutgoing.setKind("PDF");
      this.stationOutgoing.setNormalTag();
      this.stationOutgoing.setStackCount(outgoingCount);
      this.stationOutgoing.setTag(`${outgoingCount} READY`);
      const point = this.stationPilePoint("outgoing");
      this.stationOutgoing.view.position.set(point.x, point.y);
      this.stationOutgoing.view.visible = true;
    } else {
      this.stationOutgoing.view.visible = false;
    }
  }

  private activateStationPiles(): void {
    this.stationPilesActive = true;
    this.document.view.visible = false;
    this.bundleLocation = "station";
    this.syncStationPiles();
  }

  private async animateStationTransfer(speed: number): Promise<void> {
    if (!this.stationPilesActive) {
      this.activateStationPiles();
    }

    const nextConverted = Math.min(this.currentBatchSize, this.convertedInBatch + 1);
    const incomingCount = Math.max(0, this.currentBatchSize - this.failedInBatch - nextConverted);
    const outgoingCount = nextConverted;
    const start = this.stationPilePoint("incoming");
    const mid = this.stationTransferMidPoint();
    const end = this.stationPilePoint("outgoing");

    this.stationTransfer.setKind("MSG");
    this.stationTransfer.setNormalTag();
    this.stationTransfer.setStackCount(1);
    this.stationTransfer.setTag("PROCESS");
    this.stationTransfer.view.position.set(start.x, start.y);
    this.stationTransfer.view.visible = true;

    if (incomingCount > 0) {
      this.stationIncoming.setStackCount(incomingCount);
      this.stationIncoming.setTag(`${incomingCount} LEFT`);
    } else {
      this.stationIncoming.view.visible = false;
    }

    await this.moveView(this.stationTransfer.view, mid.x, mid.y, scaleDuration(260, speed));
    this.stationTransfer.setKind("PDF");
    this.stationTransfer.setNormalTag();
    this.stationTransfer.setTag("READY");
    await this.moveView(this.stationTransfer.view, end.x, end.y, scaleDuration(260, speed));

    this.convertedInBatch = outgoingCount;
    this.syncStationPiles();
    this.stationTransfer.view.visible = false;
  }

  private prepareOutgoingBundleForCarry(): void {
    const outgoingCount = Math.max(1, this.outgoingPileCount());
    const outgoingPoint = this.stationPilePoint("outgoing");
    this.document.setKind("PDF");
    this.document.setNormalTag();
    this.document.setStackCount(outgoingCount);
    this.document.setTag(`${outgoingCount} PDF`);
    this.document.view.visible = true;
    this.stage.addChild(this.document.view);
    this.document.view.position.set(outgoingPoint.x, outgoingPoint.y);
    this.stationOutgoing.view.visible = false;
    this.stationTransfer.view.visible = false;
    this.bundleLocation = "station";
  }

  private resetHeroScene(): void {
    this.currentStageEl.textContent = "Queued";
    this.outputNameEl.textContent = "not built yet";
    this.pipelineEl.textContent = "pending";
    this.lastPipeline = null;
    this.destinationLight.visible = false;
    this.destinationLight.alpha = 0.1;

    const entryStop = this.actorStop("entry");
    this.carrier.setPosition(entryStop.x, entryStop.y);

    this.convertedInBatch = 0;
    this.document.setKind("MSG");
    this.document.setNormalTag();
    this.document.setStackCount(1);
    this.document.setTag("");
    this.document.view.visible = false;
    this.stage.addChild(this.document.view);
    const entryPoint = this.bundlePoint("entry");
    this.document.view.position.set(entryPoint.x, entryPoint.y);
    this.bundleLocation = "entry";
    this.terminalDisplayLocation = null;
    this.hideStationPiles();
    this.resetErrorPile();

    this.stationGlow.clear();
    this.stationGlow.circle(442, 326, 88).fill({ color: 0x4f78ff, alpha: 0.1 });
  }

  private async applyEvent(
    event: TaskEvent,
    snapshot: TaskSnapshot,
    options: AnimationOptions
  ): Promise<void> {
    if (
      typeof event.meta?.batchId === "string" &&
      (this.currentBatchId === null || this.currentBatchId === snapshot.taskId)
    ) {
      this.currentBatchId = event.meta.batchId;
    }
    if (snapshot.batchSize && snapshot.batchSize > this.currentBatchSize) {
      this.currentBatchSize = snapshot.batchSize;
    }

    const batchMode = this.isBatchMode(snapshot);
    const finalBatchTask = this.isFinalBatchTask(snapshot);

    this.currentStageEl.textContent = STAGE_LABELS[event.stage];
    this.currentFileEl.textContent = event.fileName ?? snapshot.fileName ?? snapshot.taskId;

    if (typeof event.meta?.outputName === "string") {
      this.outputNameEl.textContent = event.meta.outputName;
    } else if (snapshot.outputName) {
      this.outputNameEl.textContent = snapshot.outputName;
    }
    if (event.pipeline) {
      this.pipelineEl.textContent = event.pipeline;
    } else if (snapshot.pipeline) {
      this.pipelineEl.textContent = snapshot.pipeline;
    }

    this.addFeedLine(event, options.compressed);

    switch (event.stage) {
      case "drop_received": {
        if (this.stationPilesActive) {
          break;
        }
        if (batchMode && this.bundleLocation !== "entry" && this.document.view.visible) {
          break;
        }
        this.document.setKind("MSG");
        this.document.setNormalTag();
        this.document.setTag("");
        this.document.setStackCount(this.batchSizeFor(snapshot));
        const entryPoint = this.bundlePoint("entry");
        this.document.view.position.set(entryPoint.x, entryPoint.y);
        this.document.view.visible = true;
        this.bundleLocation = "entry";
        break;
      }

      case "outlook_extract_started":
        await this.paceAt("entry", options.speed);
        break;

      case "files_accepted":
        if (this.stationPilesActive) {
          await this.paceAt("station", options.speed);
          break;
        }
        this.document.setKind("MSG");
        this.document.setNormalTag();
        this.document.setStackCount(this.batchSizeFor(snapshot));
        this.document.setTag(batchMode ? `${this.currentBatchSize} MSG` : "MSG");
        if (!batchMode || this.bundleLocation === "entry") {
          await this.carryBundle(this.bundleLocation, "intake", options.speed);
        }
        break;

      case "output_folder_selected":
        this.destinationLight.visible = true;
        this.destinationLight.alpha = 0.22;
        await this.paceAt(this.conversionHoldLocation(), options.speed);
        break;

      case "parse_started":
        await this.paceAt(this.conversionHoldLocation(), options.speed);
        break;

      case "filename_built":
        this.document.setNormalTag();
        this.document.setTag(
          batchMode
            ? `TAG ${this.completedInBatch + this.failedInBatch + 1}/${this.currentBatchSize}`
            : "DATED"
        );
        await this.paceAt(this.conversionHoldLocation(), options.speed);
        break;

      case "pdf_pipeline_started":
        if (this.stationPilesActive) {
          await this.paceAt("station", options.speed);
          break;
        }
        if (this.bundleLocation === "entry") {
          await this.carryBundle("entry", "intake", options.speed);
        }
        if (this.bundleLocation === "station") {
          await this.paceAt("station", options.speed);
          break;
        }
        await this.carryBundle(this.bundleLocation, "station", options.speed);
        this.activateStationPiles();
        break;

      case "pipeline_selected":
        this.lastPipeline = event.pipeline ?? null;
        this.pipelineEl.textContent = event.pipeline ?? "none";
        this.recolorStation(event.pipeline ?? null);
        await this.paceAt("station", options.speed);
        break;

      case "pdf_written": {
        if (!this.stationPilesActive) {
          this.activateStationPiles();
        }
        await this.animateStationTransfer(options.speed);
        await this.paceAt("station", options.speed);
        break;
      }

      case "deliver_started":
        if (batchMode && !finalBatchTask) {
          await this.paceAt("station", options.speed);
          break;
        }
        this.prepareOutgoingBundleForCarry();
        await this.carryBundle(this.bundleLocation, "destination", options.speed);
        break;

      case "complete":
        this.completedInBatch += 1;
        if (batchMode && !this.isFinalBatchItem(this.completedInBatch)) {
          this.syncStationPiles();
          break;
        }
        this.destinationLight.visible = true;
        this.destinationLight.alpha = 0.4;
        if (batchMode || this.convertedInBatch > 0) {
          this.document.setKind("PDF");
          this.document.setNormalTag();
          this.document.setStackCount(Math.max(1, this.convertedInBatch));
          this.document.setTag(`${Math.max(1, this.convertedInBatch)} PDF`);
          const destinationPoint = this.bundlePoint("destination");
          this.document.view.position.set(destinationPoint.x, destinationPoint.y);
          this.bundleLocation = "destination";
        }
        this.hideStationPiles();
        this.terminalDisplayLocation = "destination";
        break;

      case "failed": {
        this.failedInBatch += 1;
        this.pipelineEl.textContent = event.pipeline ?? this.pipelineEl.textContent;
        if (this.stationPilesActive) {
          this.syncStationPiles();
        }
        const failedKind = snapshot.hasPdfWritten ? "PDF" : "MSG";
        this.document.setKind(failedKind);
        this.document.setFailedTag();
        this.document.setStackCount(1);
        if (this.stationPilesActive) {
          const incomingPoint = this.stationPilePoint("incoming");
          this.document.view.visible = true;
          this.stage.addChild(this.document.view);
          this.document.view.position.set(incomingPoint.x, incomingPoint.y);
          this.bundleLocation = "station";
        }
        await this.carryBundle(this.bundleLocation, "error", options.speed);
        this.errorBin.alpha = 1;
        await sleep(scaleDuration(220, options.speed));
        this.depositErrorFile(failedKind);

        if (this.incomingPileCount() > 0) {
          this.bundleLocation = "station";
          await this.moveActor(
            this.carrier,
            this.actorStop("station").x,
            this.actorStop("station").y,
            scaleDuration(420, options.speed)
          );
          break;
        }

        if (this.outgoingPileCount() > 0) {
          this.prepareOutgoingBundleForCarry();
          this.destinationLight.visible = true;
          this.destinationLight.alpha = 0.34;
          await this.carryBundle("station", "destination", options.speed);
          this.hideStationPiles();
          this.terminalDisplayLocation = "destination";
          break;
        }
        this.hideStationPiles();
        this.terminalDisplayLocation = null;
        break;
      }
    }
  }

  private recolorStation(pipeline: RenderPipeline | null): void {
    let glow = 0x4f78ff;
    if (pipeline === "edge_html") {
      glow = 0xdb9a38;
    } else if (pipeline === "reportlab") {
      glow = 0xb55b34;
    }

    this.stationGlow.clear();
    this.stationGlow.circle(442, 326, 88).fill({ color: glow, alpha: 0.16 });
  }

  private addFeedLine(event: TaskEvent, compressed: boolean): void {
    const item = document.createElement("li");
    const pieces = [event.fileName ?? compactLabel(event.taskId), STAGE_LABELS[event.stage]];
    if (event.pipeline) {
      pieces.push(`via ${event.pipeline}`);
    }
    if (compressed) {
      pieces.push("(compressed)");
    }
    item.textContent = pieces.join(" • ");
    this.eventListEl.appendChild(item);

    while (this.eventListEl.children.length > 18) {
      this.eventListEl.removeChild(this.eventListEl.firstElementChild as Node);
    }
    this.eventListEl.scrollTop = this.eventListEl.scrollHeight;
  }

  private async carryBundle(
    from: BundleLocation,
    to: BundleLocation,
    speed: number
  ): Promise<void> {
    if (from === to) {
      await this.paceAt(to, speed);
      return;
    }

    const pickupStop = this.actorStop(from);
    const dropStop = this.actorStop(to);
    await this.moveActor(this.carrier, pickupStop.x, pickupStop.y, scaleDuration(260, speed));
    await this.attachDocumentToActor(this.carrier, -20, -112, speed);
    await this.moveActor(this.carrier, dropStop.x, dropStop.y, scaleDuration(700, speed));
    const dropPoint = this.bundlePoint(to);
    this.detachDocument(dropPoint.x, dropPoint.y);
    this.bundleLocation = to;
  }

  private async paceAt(location: BundleLocation, speed: number): Promise<void> {
    const anchor = this.actorStop(location);
    const firstX = anchor.x + 22;
    const secondX = anchor.x - 16;
    await this.moveActor(this.carrier, anchor.x, anchor.y, scaleDuration(180, speed));
    await this.moveActor(this.carrier, firstX, anchor.y, scaleDuration(180, speed));
    await this.moveActor(this.carrier, secondX, anchor.y, scaleDuration(150, speed));
    await this.moveActor(this.carrier, anchor.x, anchor.y, scaleDuration(150, speed));
  }

  private async moveActor(
    actor: StageActor,
    x: number,
    y: number,
    durationMs: number
  ): Promise<void> {
    const fromX = actor.view.x;
    const fromY = actor.view.y;
    const start = performance.now();

    return new Promise((resolve) => {
      const tick = () => {
        const elapsed = performance.now() - start;
        const t = Math.min(1, elapsed / durationMs);
        const eased = easeInOutCubic(t);
        actor.setPosition(lerp(fromX, x, eased), lerp(fromY, y, eased));
        if (t < 1) {
          requestAnimationFrame(tick);
          return;
        }
        resolve();
      };
      tick();
    });
  }

  private async moveView(
    view: Container,
    x: number,
    y: number,
    durationMs: number
  ): Promise<void> {
    const fromX = view.x;
    const fromY = view.y;
    const start = performance.now();

    return new Promise((resolve) => {
      const tick = () => {
        const elapsed = performance.now() - start;
        const t = Math.min(1, elapsed / durationMs);
        const eased = easeInOutCubic(t);
        view.position.set(lerp(fromX, x, eased), lerp(fromY, y, eased));
        if (t < 1) {
          requestAnimationFrame(tick);
          return;
        }
        resolve();
      };
      tick();
    });
  }

  private async attachDocumentToActor(
    actor: StageActor,
    offsetX: number,
    offsetY: number,
    speed: number
  ): Promise<void> {
    actor.view.addChild(this.document.view);
    this.document.view.position.set(offsetX, offsetY);
    await sleep(scaleDuration(120, speed));
  }

  private detachDocument(x: number, y: number): void {
    this.stage.addChild(this.document.view);
    this.document.view.position.set(x, y);
  }

  private readDemoSettings(): DemoSettings {
    this.syncDemoFailureLimit();
    return {
      fileCount: clampInteger(this.demoFileCountEl.value, 5, 1, 10),
      failureCount: clampInteger(
        this.demoFailureCountEl.value,
        0,
        0,
        clampInteger(this.demoFileCountEl.value, 5, 1, 10)
      )
    };
  }

  private syncDemoFailureLimit(): void {
    const fileCount = clampInteger(this.demoFileCountEl.value, 5, 1, 10);
    this.demoFileCountEl.value = `${fileCount}`;
    this.demoFailureCountEl.max = `${fileCount}`;
    const failureCount = clampInteger(
      this.demoFailureCountEl.value,
      0,
      0,
      fileCount
    );
    this.demoFailureCountEl.value = `${failureCount}`;
  }
}

function compactLabel(value: string): string {
  const stem = value.replace(/\.msg$/i, "").replace(/\.pdf$/i, "");
  return stem.length <= 10 ? stem : `${stem.slice(0, 9)}…`;
}

function clampInteger(
  value: string | number,
  fallback: number,
  min: number,
  max: number
): number {
  const parsed = Number.parseInt(`${value}`, 10);
  if (!Number.isFinite(parsed)) {
    return fallback;
  }
  return Math.max(min, Math.min(max, parsed));
}

function scaleDuration(baseMs: number, speed: number): number {
  return Math.max(120, Math.round(baseMs * speed));
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

function lerp(from: number, to: number, t: number): number {
  return from + (to - from) * t;
}

function easeInOutCubic(t: number): number {
  return t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2;
}
