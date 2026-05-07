function easeOutCubic(value) {
  return 1 - Math.pow(1 - value, 3);
}

function easeOutQuart(value) {
  return 1 - Math.pow(1 - value, 4);
}

const RIPPLE_PIXEL_RATIO_CAP = 1.35;
const RIPPLE_MAX_CANVAS_AREA = 520000;
const RIPPLE_TARGET_FRAME_MS = 1000 / 30;

export function captureDropFiles(dataTransfer) {
  const items = Array.from(dataTransfer?.items || []);
  const itemFiles = items
    .filter((item) => item.kind === "file")
    .map((item) => item.getAsFile())
    .filter(Boolean);
  if (itemFiles.length) {
    return itemFiles;
  }
  return Array.from(dataTransfer?.files || []);
}

export function isPointerInsideElement(event, element) {
  const rect = element.getBoundingClientRect();
  return event.clientX >= rect.left && event.clientX <= rect.right && event.clientY >= rect.top && event.clientY <= rect.bottom;
}

function resizeRippleCanvas(dropzone, canvas) {
  if (!canvas || !dropzone) {
    return null;
  }
  const rect = dropzone.getBoundingClientRect();
  const rawPixelRatio = Math.min(window.devicePixelRatio || 1, RIPPLE_PIXEL_RATIO_CAP);
  const rawWidth = Math.max(1, Math.floor(rect.width * rawPixelRatio));
  const rawHeight = Math.max(1, Math.floor(rect.height * rawPixelRatio));
  const areaScale = Math.min(1, Math.sqrt(RIPPLE_MAX_CANVAS_AREA / Math.max(1, rawWidth * rawHeight)));
  const pixelRatio = Math.max(1, rawPixelRatio * areaScale);
  canvas.width = Math.max(1, Math.floor(rect.width * pixelRatio));
  canvas.height = Math.max(1, Math.floor(rect.height * pixelRatio));
  canvas.style.width = `${rect.width}px`;
  canvas.style.height = `${rect.height}px`;
  const context = canvas.getContext("2d");
  context.imageSmoothingEnabled = false;
  context.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
  return { context, width: rect.width, height: rect.height };
}

function drawRippleEllipse(context, centerX, centerY, radiusX, radiusY, alpha, lineWidth, color) {
  context.save();
  context.globalAlpha = alpha;
  context.lineWidth = lineWidth;
  context.strokeStyle = color;
  context.shadowBlur = 8;
  context.shadowColor = color;
  context.beginPath();
  context.ellipse(centerX, centerY, radiusX, radiusY, 0, 0, Math.PI * 2);
  context.stroke();
  context.restore();
}

function drawDropRippleFrame(context, width, height, progress, impactPoint = null) {
  const centerX = impactPoint?.x ?? width / 2;
  const centerY = impactPoint?.y ?? height * 0.52;
  context.clearRect(0, 0, width, height);

  const surfaceFade = Math.max(0, 1 - progress);
  const surface = context.createRadialGradient(centerX, centerY, 8, centerX, centerY, width * 0.42);
  surface.addColorStop(0, `rgba(255, 239, 198, ${0.2 * surfaceFade})`);
  surface.addColorStop(0.32, `rgba(157, 211, 109, ${0.08 * surfaceFade})`);
  surface.addColorStop(1, "rgba(0, 0, 0, 0)");
  context.fillStyle = surface;
  context.fillRect(0, 0, width, height);

  const flashLocal = Math.min(1, progress / 0.22);
  const flashAlpha = Math.max(0, 1 - flashLocal) * 0.42;
  if (flashAlpha > 0.01) {
    const flash = context.createRadialGradient(centerX, centerY, 0, centerX, centerY, Math.max(width, height) * 0.72);
    flash.addColorStop(0, `rgba(255, 243, 212, ${flashAlpha})`);
    flash.addColorStop(0.22, `rgba(243, 180, 89, ${flashAlpha * 0.46})`);
    flash.addColorStop(1, "rgba(0, 0, 0, 0)");
    context.fillStyle = flash;
    context.fillRect(0, 0, width, height);
  }

  const glintProgress = Math.min(1, progress / 0.44);
  const glintAlpha = Math.max(0, Math.sin(glintProgress * Math.PI)) * 0.7;
  if (glintAlpha > 0.01) {
    const glint = context.createLinearGradient(centerX - width * 0.28, centerY, centerX + width * 0.28, centerY);
    glint.addColorStop(0, "rgba(255, 244, 218, 0)");
    glint.addColorStop(0.5, `rgba(255, 244, 218, ${glintAlpha})`);
    glint.addColorStop(1, "rgba(157, 211, 109, 0)");
    context.save();
    context.lineWidth = 2;
    context.strokeStyle = glint;
    context.shadowBlur = 8;
    context.shadowColor = "rgba(255, 224, 160, 0.72)";
    context.beginPath();
    context.moveTo(centerX - width * 0.28 * glintProgress, centerY);
    context.lineTo(centerX + width * 0.28 * glintProgress, centerY);
    context.stroke();
    context.restore();
  }

  [
    { delay: 0, speed: 1, color: "rgba(255, 246, 226, 0.95)", width: 2.4 },
    { delay: 0.08, speed: 0.92, color: "rgba(157, 211, 109, 0.78)", width: 2.1 },
    { delay: 0.18, speed: 0.78, color: "rgba(255, 205, 118, 0.6)", width: 1.6 },
  ].forEach((ring) => {
    const local = Math.max(0, Math.min(1, (progress - ring.delay) / (1 - ring.delay)));
    if (!local) {
      return;
    }
    const eased = easeOutQuart(local);
    const alpha = Math.pow(1 - local, 1.28) * ring.speed;
    drawRippleEllipse(
      context,
      centerX,
      centerY,
      24 + eased * width * 0.37,
      7 + eased * height * 0.18,
      alpha,
      ring.width + (1 - local) * 1.4,
      ring.color,
    );
  });

  [
    { angle: -0.96, length: 0.32, delay: 0.02, color: "rgba(255, 244, 218, 0.88)" },
    { angle: -0.18, length: 0.3, delay: 0.08, color: "rgba(157, 211, 109, 0.66)" },
    { angle: 0.58, length: 0.34, delay: 0.06, color: "rgba(255, 244, 218, 0.72)" },
  ].forEach((beam) => {
    const local = Math.max(0, Math.min(1, (progress - beam.delay) / 0.5));
    if (!local || local >= 1) {
      return;
    }
    const eased = easeOutQuart(local);
    const distance = Math.min(width, height) * beam.length * eased;
    const startDistance = 18 + distance * 0.18;
    const endDistance = 32 + distance;
    context.save();
    context.globalAlpha = Math.sin(local * Math.PI) * 0.9;
    context.lineWidth = 1.6 + (1 - local) * 2.4;
    context.strokeStyle = beam.color;
    context.shadowBlur = 10;
    context.shadowColor = beam.color;
    context.beginPath();
    context.moveTo(centerX + Math.cos(beam.angle) * startDistance, centerY + Math.sin(beam.angle) * startDistance);
    context.lineTo(centerX + Math.cos(beam.angle) * endDistance, centerY + Math.sin(beam.angle) * endDistance);
    context.stroke();
    context.restore();
  });

  const crownLocal = Math.min(1, progress / 0.48);
  const crownAlpha = Math.max(0, Math.sin(crownLocal * Math.PI)) * 0.95;
  if (crownAlpha > 0.01) {
    context.save();
    context.globalAlpha = crownAlpha;
    context.shadowBlur = 8;
    context.shadowColor = "rgba(255, 226, 164, 0.9)";
    context.fillStyle = "rgba(255, 244, 218, 0.95)";
    const plumeHeight = 44 * Math.sin(crownLocal * Math.PI);
    context.beginPath();
    context.ellipse(centerX, centerY - plumeHeight * 0.35, 9 + crownLocal * 8, 15 + plumeHeight * 0.25, 0, 0, Math.PI * 2);
    context.fill();
    context.restore();
  }

  [
    [-118, -64, 0],
    [-82, -108, 0.03],
    [-36, -128, 0.07],
    [38, -130, 0.07],
    [86, -104, 0.03],
    [122, -62, 0],
    [0, -56, 0.1],
  ].forEach(([offsetX, offsetY, delay]) => {
    const local = Math.max(0, Math.min(1, (progress - delay) / 0.58));
    if (!local || local >= 1) {
      return;
    }
    const arc = easeOutCubic(local);
    const x = centerX + offsetX * arc;
    const y = centerY + offsetY * arc + 46 * local * local;
    const alpha = Math.sin(local * Math.PI);
    context.save();
    context.globalAlpha = alpha;
    context.fillStyle = "rgba(255, 240, 204, 0.95)";
    context.shadowBlur = 6;
    context.shadowColor = "rgba(255, 224, 160, 0.8)";
    context.beginPath();
    context.arc(x, y, 3 + (1 - local) * 3, 0, Math.PI * 2);
    context.fill();
    context.restore();
  });
}

export function createDropzoneController({
  canvas,
  dropzone,
  onDrop,
  onDragEnd,
  onDragIntent,
  onError,
  onFinally,
  sourceHintFromDrop,
}) {
  let rippleAnimationFrame = 0;
  let splashTimer = 0;
  let dragIntentActive = false;

  const clearRipple = () => {
    if (rippleAnimationFrame) {
      window.cancelAnimationFrame(rippleAnimationFrame);
      rippleAnimationFrame = 0;
    }
    const context = canvas?.getContext("2d");
    if (context && canvas) {
      context.clearRect(0, 0, canvas.width, canvas.height);
    }
  };

  const playRipple = (event = null) => {
    if (window.matchMedia?.("(prefers-reduced-motion: reduce)")?.matches) {
      clearRipple();
      return;
    }
    const setup = resizeRippleCanvas(dropzone, canvas);
    if (!setup) {
      return;
    }
    const rect = dropzone.getBoundingClientRect();
    const impactPoint = event
      ? {
          x: Math.max(0, Math.min(setup.width, event.clientX - rect.left)),
          y: Math.max(0, Math.min(setup.height, event.clientY - rect.top)),
        }
      : null;
    if (rippleAnimationFrame) {
      window.cancelAnimationFrame(rippleAnimationFrame);
    }
    const duration = 1180;
    const start = performance.now();
    let lastFrameAt = start - RIPPLE_TARGET_FRAME_MS;
    const draw = (now) => {
      if (now - lastFrameAt < RIPPLE_TARGET_FRAME_MS && now - start < duration) {
        rippleAnimationFrame = window.requestAnimationFrame(draw);
        return;
      }
      lastFrameAt = now;
      const progress = Math.min(1, (now - start) / duration);
      drawDropRippleFrame(setup.context, setup.width, setup.height, progress, impactPoint);
      if (progress < 1) {
        rippleAnimationFrame = window.requestAnimationFrame(draw);
        return;
      }
      setup.context.clearRect(0, 0, setup.width, setup.height);
      rippleAnimationFrame = 0;
    };
    draw(start);
    rippleAnimationFrame = window.requestAnimationFrame(draw);
  };

  const splash = (event = null) => {
    if (splashTimer) {
      clearTimeout(splashTimer);
    }
    dropzone.classList.remove("is-drop-splash");
    void dropzone.offsetWidth;
    dropzone.classList.add("is-drop-splash");
    playRipple(event);
    splashTimer = window.setTimeout(() => {
      dropzone.classList.remove("is-drop-splash");
      splashTimer = 0;
    }, 1320);
  };

  const setDragOver = (isOver) => {
    dropzone.classList.toggle("is-dragover", isOver);
  };

  const clearDragIntent = () => {
    if (!dragIntentActive) {
      return;
    }
    dragIntentActive = false;
    dropzone.classList.remove("is-drop-primed");
    onDragEnd?.();
  };

  const primeDragIntent = (event) => {
    if (dragIntentActive) {
      return;
    }
    dragIntentActive = true;
    dropzone.classList.add("is-drop-primed");
    const sourceHint = sourceHintFromDrop?.(event.dataTransfer) || null;
    onDragIntent?.({ sourceHint });
  };

  const isPointInsideDropzone = (event) => isPointerInsideElement(event, dropzone);

  const prevent = (event) => {
    event.preventDefault();
    event.stopPropagation();
  };

  const handleDragEnterOrOver = (event) => {
    prevent(event);
    setDragOver(true);
    primeDragIntent(event);
  };

  const handleDragLeave = (event) => {
    prevent(event);
    if (!isPointInsideDropzone(event)) {
      setDragOver(false);
      clearDragIntent();
    }
  };

  const handleDrop = async (event) => {
    prevent(event);
    setDragOver(false);
    dragIntentActive = false;
    dropzone.classList.remove("is-drop-primed");
    try {
      const dataTransfer = event.dataTransfer;
      splash(event);
      const sourceHint = sourceHintFromDrop?.(dataTransfer) || null;
      const files = captureDropFiles(dataTransfer);
      await onDrop(files, { sourceHint });
    } catch (error) {
      onError?.(error);
    } finally {
      onFinally?.();
    }
  };

  const handleDocumentDragEnd = () => {
    setDragOver(false);
    clearDragIntent();
  };
  const handleDocumentDragOver = (event) => {
    event.preventDefault();
  };
  const handleDocumentDrop = (event) => {
    event.preventDefault();
    setDragOver(false);
    clearDragIntent();
  };

  dropzone.addEventListener("dragenter", handleDragEnterOrOver);
  dropzone.addEventListener("dragover", handleDragEnterOrOver);
  dropzone.addEventListener("dragleave", handleDragLeave);
  dropzone.addEventListener("drop", handleDrop);
  document.addEventListener("dragend", handleDocumentDragEnd);
  document.addEventListener("dragover", handleDocumentDragOver);
  document.addEventListener("drop", handleDocumentDrop);

  return {
    clear() {
      if (splashTimer) {
        clearTimeout(splashTimer);
        splashTimer = 0;
      }
      dropzone.classList.remove("is-dragover", "is-drop-splash");
      clearDragIntent();
      clearRipple();
    },
    destroy() {
      dropzone.removeEventListener("dragenter", handleDragEnterOrOver);
      dropzone.removeEventListener("dragover", handleDragEnterOrOver);
      dropzone.removeEventListener("dragleave", handleDragLeave);
      dropzone.removeEventListener("drop", handleDrop);
      document.removeEventListener("dragend", handleDocumentDragEnd);
      document.removeEventListener("dragover", handleDocumentDragOver);
      document.removeEventListener("drop", handleDocumentDrop);
      this.clear();
    },
  };
}
