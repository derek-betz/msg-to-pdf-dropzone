# Web Dropzone Contract

This note records the frontend/backend behavior that future UI work must preserve.

## Dropzone event contract

- `dragenter` and `dragover` may show a hover affordance, but they must not trigger the committed drop animation.
- `drop` is the only event that starts the ripple/splash effect.
- Dropped files must be captured synchronously from `DataTransfer.items` or `DataTransfer.files` before any `await`, timer, or framework state boundary.
- `dragleave` must not clear hover state when the pointer is still inside the dropzone rectangle. Crossing icons, labels, overlays, or the ripple canvas should not cause flicker.
- Document-level `dragover` and `drop` handlers should prevent default navigation if a file is released outside the dropzone.
- The dropzone controller owns event listeners, visual drag/drop classes, ripple cleanup, and file capture. Application code owns upload, queue state, API calls, and status messaging through callbacks.

## Queue filename contract

The queue must display the server-provided `outputName` when it is available. That value is not a simple `message-name.pdf` transform.

For `.msg` files, the backend parses each staged message, groups messages by normalized thread subject, finds the latest sent date in that thread, and builds the preview filename as:

```text
YYYY-MM-DD_<message subject>.pdf
```

That means an older reply in a thread can show a PDF name prefixed with the date of the newest message in the same thread. The frontend fallback name is only a temporary display value for cases where metadata parsing has not produced `outputName` yet.

Relevant backend functions:

- `StageStore.refresh_output_previews` in `src/msg_to_pdf_dropzone/web_server.py`
- `get_latest_thread_dates` and `build_pdf_filename` in `src/msg_to_pdf_dropzone/thread_logic.py`

Relevant frontend functions:

- `queueOutputName` in `src/msg_to_pdf_dropzone/web_ui/app.js`
- `outputPreviewLabel` in `src/msg_to_pdf_dropzone/web_ui/app.js`
- `createDropzoneController` in `src/msg_to_pdf_dropzone/web_ui/dropzone_controller.js`

## Regression checklist

- Hard-refresh the browser after JavaScript changes.
- Drag a real `.msg` file slowly over the dropzone, including over the icon, text, and ripple canvas.
- Confirm hover state does not flicker while the cursor moves.
- Confirm the ripple starts only after drop, not on hover.
- Confirm the queue row uses the date-prefixed `outputName` once the upload response returns.
- Confirm dropping outside the dropzone does not navigate away from the app.
