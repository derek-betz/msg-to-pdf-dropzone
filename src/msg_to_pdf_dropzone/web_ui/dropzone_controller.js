function nextAnimationFrame() {
  return new Promise((resolve) => window.requestAnimationFrame(() => resolve()));
}

function easeOutCubic(value) {
  return 1 - Math.pow(1 - value, 3);
}

function easeOutQuart(value) {
  return 1 - Math.pow(1 - value, 4);
}

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
  const pixelRatio = Math.min(window.devicePixelRatio || 1, 2);
  canvas.width = Math.max(1, Math.floor(rect.width * pixelRatio));
  canvas.height = Math.max(1, Math.floor(rect.height * pixelRatio));
  canvas.style.width = `${rect.width}px`;
  canvas.style.height = `${rect.height}px`;
  const context = canvas.getContext("2d");
  context.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
  return { context, width: rect.width, height: rect.height };
}

function drawRippleEllipse(context, centerX, centerY, radiusX, radiusY, alpha, lineWidth, color) {
  context.save();
  context.globalAlpha = alpha;
  context.lineWidth = lineWidth;
  context.strokeStyle = color;
  context.shadowBlur = 18;
  context.shadowColor = color;
  context.beginPath();
  context.ellipse(centerX, centerY, radiusX, radiusY, 0, 0, Math.PI * 2);
  context.stroke();
  context.restore();
}

function drawDropRippleFrame(context, width, height, progress) {
  const centerX = width / 2;
  const centerY = height * 0.52;
  context.clearRect(0, 0, width, height);

  const surfaceFade = Math.max(0, 1 - progress);
  const surface = context.createRadialGradient(centerX, centerY, 8, centerX, centerY, width * 0.42);
  surface.addColorStop(0, `rgba(255, 239, 198, ${0.2 * surfaceFade})`);
  surface.addColorStop(0.32, `rgba(157, 211, 109, ${0.08 * surfaceFade})`);
  surface.addColorStop(1, "rgba(0, 0, 0, 0)");
  context.fillStyle = surface;
  context.fillRect(0, 0, width, height);

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
    context.shadowBlur = 18;
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
    { delay: 0.16, speed: 0.84, color: "rgba(255, 205, 118, 0.65)", width: 1.8 },
    { delay: 0.25, speed: 0.72, color: "rgba(255, 246, 226, 0.42)", width: 1.2 },
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

  const crownLocal = Math.min(1, progress / 0.48);
  const crownAlpha = Math.max(0, Math.sin(crownLocal * Math.PI)) * 0.95;
  if (crownAlpha > 0.01) {
    context.save();
    context.globalAlpha = crownAlpha;
    context.shadowBlur = 18;
    context.shadowColor = "rgba(255, 226, 164, 0.9)";
    context.fillStyle = "rgba(255, 244, 218, 0.95)";
    const plumeHeight = 44 * Math.sin(crownLocal * Math.PI);
    context.beginPath();
    context.ellipse(centerX, centerY - plumeHeight * 0.35, 9 + crownLocal * 8, 15 + plumeHeight * 0.25, 0, 0, Math.PI * 2);
    context.fill();
    context.restore();
  }

  [
    [-96, -52, 0],
    [-58, -82, 0.04],
    [0, -96, 0.07],
    [58, -82, 0.04],
    [96, -52, 0],
    [-28, -42, 0.1],
    [28, -42, 0.1],
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
    context.shadowBlur = 14;
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
  onError,
  onFinally,
  sourceHintFromDrop,
}) {
  let rippleAnimationFrame = 0;
  let splashTimer = 0;

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

  const playRipple = () => {
    const setup = resizeRippleCanvas(dropzone, canvas);
    if (!setup) {
      return;
    }
    if (rippleAnimationFrame) {
      window.cancelAnimationFrame(rippleAnimationFrame);
    }
    const duration = 980;
    const start = performance.now();
    const draw = (now) => {
      const progress = Math.min(1, (now - start) / duration);
      drawDropRippleFrame(setup.context, setup.width, setup.height, progress);
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

  const splash = () => {
    if (splashTimer) {
      clearTimeout(splashTimer);
    }
    dropzone.classList.remove("is-drop-splash");
    void dropzone.offsetWidth;
    dropzone.classList.add("is-drop-splash");
    playRipple();
    splashTimer = window.setTimeout(() => {
      dropzone.classList.remove("is-drop-splash");
      splashTimer = 0;
    }, 1180);
  };

  const setDragOver = (isOver) => {
    dropzone.classList.toggle("is-dragover", isOver);
  };

  const isPointInsideDropzone = (event) => isPointerInsideElement(event, dropzone);

  const prevent = (event) => {
    event.preventDefault();
    event.stopPropagation();
  };

  const handleDragEnterOrOver = (event) => {
    prevent(event);
    setDragOver(true);
  };

  const handleDragLeave = (event) => {
    prevent(event);
    if (!isPointInsideDropzone(event)) {
      setDragOver(false);
    }
  };

  const handleDrop = async (event) => {
    prevent(event);
    setDragOver(false);
    try {
      const dataTransfer = event.dataTransfer;
      splash();
      const files = captureDropFiles(dataTransfer);
      const sourceHint = sourceHintFromDrop?.(dataTransfer) || null;
      await nextAnimationFrame();
      await onDrop(files, { sourceHint });
    } catch (error) {
      onError?.(error);
    } finally {
      onFinally?.();
    }
  };

  const handleDocumentDragEnd = () => setDragOver(false);
  const handleDocumentDragOver = (event) => {
    event.preventDefault();
  };
  const handleDocumentDrop = (event) => {
    event.preventDefault();
    setDragOver(false);
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
