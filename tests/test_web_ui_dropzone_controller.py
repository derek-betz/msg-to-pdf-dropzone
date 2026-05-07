from __future__ import annotations

import os
from pathlib import Path
import shutil
import subprocess
import textwrap

import pytest


REPO_ROOT = Path(__file__).resolve().parents[1]
CONTROLLER_PATH = REPO_ROOT / "src" / "msg_to_pdf_dropzone" / "web_ui" / "dropzone_controller.js"


def _node_executable() -> str | None:
    configured = os.environ.get("NODE_EXE")
    if configured and Path(configured).exists():
        return configured
    from_path = shutil.which("node")
    if from_path:
        return from_path
    bundled = (
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
        / "node"
        / "bin"
        / "node.exe"
    )
    if bundled.exists():
        return str(bundled)
    return None


def test_dropzone_controller_event_contract(tmp_path: Path) -> None:
    node = _node_executable()
    if node is None:
        pytest.skip("Node.js is not available for web UI controller regression tests.")

    harness = tmp_path / "dropzone-controller-harness.mjs"
    controller_url = CONTROLLER_PATH.as_uri()
    harness.write_text(
        textwrap.dedent(
            f"""
            import assert from "node:assert/strict";

            const rafQueue = [];
            let rafId = 0;
            const timeoutQueue = [];
            const timeoutIds = new Set();
            let timeoutId = 0;

            function makeEventTarget() {{
              const listeners = new Map();
              return {{
                listeners,
                addEventListener(type, callback) {{
                  if (!listeners.has(type)) listeners.set(type, new Set());
                  listeners.get(type).add(callback);
                }},
                removeEventListener(type, callback) {{
                  listeners.get(type)?.delete(callback);
                }},
                dispatchEvent(event) {{
                  event.target = event.target || this;
                  for (const callback of Array.from(listeners.get(event.type) || [])) {{
                    callback(event);
                  }}
                  return !event.defaultPrevented;
                }},
              }};
            }}

            class ClassList {{
              constructor() {{
                this.values = new Set();
              }}
              add(...names) {{
                names.forEach((name) => this.values.add(name));
              }}
              remove(...names) {{
                names.forEach((name) => this.values.delete(name));
              }}
              toggle(name, force) {{
                if (force) this.values.add(name);
                else this.values.delete(name);
              }}
              contains(name) {{
                return this.values.has(name);
              }}
            }}

            function makeEvent(type, props = {{}}) {{
              return {{
                type,
                clientX: 20,
                clientY: 20,
                defaultPrevented: false,
                propagationStopped: false,
                preventDefault() {{ this.defaultPrevented = true; }},
                stopPropagation() {{ this.propagationStopped = true; }},
                ...props,
              }};
            }}

            const contexts = [];

            function makeContext() {{
              const gradient = {{ addColorStop() {{}} }};
              const context = {{
                _shadowBlur: 0,
                maxShadowBlur: 0,
                setTransformCalls: [],
                ellipseCalls: 0,
                arcCalls: 0,
                lineToCalls: 0,
                radialGradientCalls: 0,
                linearGradientCalls: 0,
                set shadowBlur(value) {{
                  this._shadowBlur = value;
                  this.maxShadowBlur = Math.max(this.maxShadowBlur, Number(value) || 0);
                }},
                get shadowBlur() {{
                  return this._shadowBlur;
                }},
                setTransform(...args) {{ this.setTransformCalls.push(args); }},
                clearRect() {{}},
                save() {{}},
                restore() {{}},
                beginPath() {{}},
                ellipse() {{ this.ellipseCalls += 1; }},
                stroke() {{}},
                fillRect() {{}},
                createRadialGradient() {{ this.radialGradientCalls += 1; return gradient; }},
                createLinearGradient() {{ this.linearGradientCalls += 1; return gradient; }},
                moveTo() {{}},
                lineTo() {{ this.lineToCalls += 1; }},
                fill() {{}},
                arc() {{ this.arcCalls += 1; }},
              }};
              contexts.push(context);
              return context;
            }}

            globalThis.window = {{
              devicePixelRatio: 3,
              matchMedia() {{
                return {{ matches: false }};
              }},
              requestAnimationFrame(callback) {{
                rafQueue.push(callback);
                return ++rafId;
              }},
              cancelAnimationFrame() {{}},
              setTimeout(callback, delay = 0) {{
                const id = ++timeoutId;
                timeoutIds.add(id);
                timeoutQueue.push({{ id, callback, delay }});
                return id;
              }},
            }};
            globalThis.clearTimeout = (id) => timeoutIds.delete(id);
            globalThis.document = makeEventTarget();

            const {{
              captureDropFiles,
              createDropzoneController,
              isPointerInsideElement,
            }} = await import("{controller_url}");

            const fallbackFile = {{ name: "fallback.msg" }};
            assert.deepEqual(captureDropFiles({{ files: [fallbackFile] }}), [fallbackFile]);

            const dropzone = makeEventTarget();
            dropzone.classList = new ClassList();
            dropzone.offsetWidth = 640;
            dropzone.getBoundingClientRect = () => ({{ left: 10, top: 10, right: 510, bottom: 310, width: 500, height: 300 }});

            const canvas = {{
              width: 0,
              height: 0,
              style: {{}},
              getContext() {{ return makeContext(); }},
            }};

            const droppedFile = {{ name: "message.msg" }};
            let getAsFileCalled = false;
            let onDropCalled = false;
            let sourceHint = "";
            const dragIntentHints = [];
            let dragEnded = 0;
            const controller = createDropzoneController({{
              canvas,
              dropzone,
              onDragIntent({{ sourceHint }}) {{
                dragIntentHints.push(sourceHint);
              }},
              onDragEnd() {{
                dragEnded += 1;
              }},
              sourceHintFromDrop(dataTransfer) {{
                return Array.from(dataTransfer.types || []).includes("OutlookItem") ? "outlook" : "upload";
              }},
              async onDrop(files, options) {{
                onDropCalled = true;
                sourceHint = options.sourceHint;
                assert.equal(getAsFileCalled, true);
                assert.deepEqual(files, [droppedFile]);
              }},
            }});

            const dragOver = makeEvent("dragover", {{
              dataTransfer: {{ types: ["Files", "OutlookItem"] }},
            }});
            dropzone.dispatchEvent(dragOver);
            assert.equal(dragOver.defaultPrevented, true);
            assert.equal(dropzone.classList.contains("is-dragover"), true);
            assert.equal(dropzone.classList.contains("is-drop-primed"), true);
            assert.equal(dropzone.classList.contains("is-drop-splash"), false);
            assert.deepEqual(dragIntentHints, ["outlook"]);

            dropzone.dispatchEvent(makeEvent("dragleave", {{ clientX: 100, clientY: 100 }}));
            assert.equal(dropzone.classList.contains("is-dragover"), true);
            assert.equal(dropzone.classList.contains("is-drop-primed"), true);
            assert.equal(dragEnded, 0);

            dropzone.dispatchEvent(makeEvent("dragleave", {{ clientX: 1, clientY: 1 }}));
            assert.equal(dropzone.classList.contains("is-dragover"), false);
            assert.equal(dropzone.classList.contains("is-drop-primed"), false);
            assert.equal(dragEnded, 1);

            assert.equal(isPointerInsideElement(makeEvent("dragover", {{ clientX: 11, clientY: 11 }}), dropzone), true);
            assert.equal(isPointerInsideElement(makeEvent("dragover", {{ clientX: 9, clientY: 11 }}), dropzone), false);

            dropzone.dispatchEvent(makeEvent("dragover", {{
              dataTransfer: {{ types: ["Files", "OutlookItem"] }},
            }}));
            const drop = makeEvent("drop", {{
              dataTransfer: {{
                types: ["Files", "OutlookItem"],
                items: [
                  {{
                    kind: "file",
                    getAsFile() {{
                      getAsFileCalled = true;
                      return droppedFile;
                    }},
                  }},
                ],
                files: [],
              }},
            }});
            dropzone.dispatchEvent(drop);

            assert.equal(drop.defaultPrevented, true);
            assert.equal(drop.propagationStopped, true);
            assert.equal(dropzone.classList.contains("is-dragover"), false);
            assert.equal(dropzone.classList.contains("is-drop-primed"), false);
            assert.equal(dropzone.classList.contains("is-drop-splash"), true);
            assert.equal(getAsFileCalled, true);
            assert.equal(onDropCalled, true);
            assert.equal(sourceHint, "outlook");
            assert.equal(canvas.width, 675);
            assert.equal(canvas.height, 405);
            const rippleContext = contexts.find((context) => context.setTransformCalls.length);
            assert.ok(rippleContext);
            assert.equal(rippleContext.setTransformCalls[0][0], 1.35);

            let fakeNow = performance.now();
            while (rafQueue.length) {{
              const callback = rafQueue.shift();
              fakeNow += 80;
              callback(fakeNow);
              while (timeoutQueue.some((timer) => timer.delay === 0 && timeoutIds.has(timer.id))) {{
                const timerIndex = timeoutQueue.findIndex((timer) => timer.delay === 0 && timeoutIds.has(timer.id));
                const timer = timeoutQueue.splice(timerIndex, 1)[0];
                timeoutIds.delete(timer.id);
                timer.callback();
              }}
              await Promise.resolve();
            }}
            await Promise.resolve();
            assert.ok(rippleContext.maxShadowBlur <= 10, `max shadowBlur was ${{rippleContext.maxShadowBlur}}`);
            assert.ok(rippleContext.ellipseCalls <= 70, `ellipse calls were ${{rippleContext.ellipseCalls}}`);
            assert.ok(rippleContext.arcCalls <= 100, `arc calls were ${{rippleContext.arcCalls}}`);
            assert.ok(rippleContext.lineToCalls <= 60, `lineTo calls were ${{rippleContext.lineToCalls}}`);

            const outsideDragOver = makeEvent("dragover");
            document.dispatchEvent(outsideDragOver);
            assert.equal(outsideDragOver.defaultPrevented, true);

            dropzone.classList.add("is-dragover");
            const outsideDrop = makeEvent("drop");
            document.dispatchEvent(outsideDrop);
            assert.equal(outsideDrop.defaultPrevented, true);
            assert.equal(dropzone.classList.contains("is-dragover"), false);

            controller.destroy();
            dropzone.dispatchEvent(makeEvent("dragover"));
            assert.equal(dropzone.classList.contains("is-dragover"), false);
            """
        ),
        encoding="utf-8",
    )

    result = subprocess.run(
        [node, str(harness)],
        cwd=REPO_ROOT,
        text=True,
        capture_output=True,
        timeout=20,
    )
    assert result.returncode == 0, result.stdout + result.stderr
