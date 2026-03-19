import { JsonlEventSource } from "./event-source";
import { SceneScheduler } from "./scene-scheduler";
import { TaskStore } from "./task-store";
import { WorkshopScene } from "./workshop-scene";

const root = document.querySelector("#app");

if (!(root instanceof HTMLElement)) {
  throw new Error("Expected #app root element.");
}

const scene = new WorkshopScene(root);
void scene.init().catch((error) => {
  console.error("Failed to initialize theater renderer", error);
});

const store = new TaskStore();
const scheduler = new SceneScheduler(scene, store);
const source = new JsonlEventSource();

source.onStatus((status) => {
  scene.setConnectionStatus(status);
});

source.onEvents((events) => {
  scheduler.accept(events);
});

scene.bindDemoLauncher(async ({ fileCount, failureCount }) => {
  scene.setDemoBusy(true);
  scene.setDemoStatus("Staging a fresh demo batch...", "busy");
  try {
    const response = await fetch("/api/demo", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        fileCount,
        failureCount
      })
    });
    if (!response.ok) {
      throw new Error(`Demo request failed with ${response.status}`);
    }
    scene.setDemoStatus("Reloading the theater...", "busy");
    window.setTimeout(() => {
      window.location.reload();
    }, 160);
  } catch (error) {
    console.error("Failed to start theater demo", error);
    scene.setDemoBusy(false);
    scene.setDemoStatus("Could not start the demo batch.", "error");
  }
});

source.start();

window.addEventListener("beforeunload", () => {
  source.stop();
  scheduler.stop();
});
