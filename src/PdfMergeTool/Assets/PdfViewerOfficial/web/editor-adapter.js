(() => {
  const state = {
    mode: "select",
    edits: [],
    selectedId: null,
    fonts: [],
    fontName: "Malgun Gothic",
    textSize: 14,
    color: "#111827",
    dirty: false
  };

  let nextEditId = 1;
  let toolbar;
  let historyStack = [];
  let redoStack = [];
  let copiedEdit = null;

  function postMessage(message) {
    window.chrome?.webview?.postMessage(message);
  }

  function postDirty() {
    postMessage({
      type: "editorStateChanged",
      isDirty: state.dirty,
      editCount: state.edits.length
    });
  }

  function snapshotEdits() {
    return JSON.stringify(state.edits);
  }

  function recordHistory() {
    historyStack.push(snapshotEdits());
    if (historyStack.length > 100) {
      historyStack.shift();
    }
    redoStack = [];
  }

  function restoreSnapshot(snapshot) {
    state.edits = JSON.parse(snapshot);
    nextEditId = state.edits.reduce((nextId, edit) => {
      const match = /^edit-(\d+)$/.exec(edit.id || "");
      return match ? Math.max(nextId, Number(match[1]) + 1) : nextId;
    }, 1);
    state.selectedId = null;
    document.querySelectorAll(".asta-editor-item").forEach(element => element.remove());
    for (const edit of state.edits) {
      renderEdit(edit);
    }
    state.dirty = true;
    postDirty();
  }

  function getApp() {
    return window.PDFViewerApplication ?? null;
  }

  function getPageView(pageNumber) {
    return getApp()?.pdfViewer?.getPageView(pageNumber - 1) ?? null;
  }

  function getPageSize(pageNumber) {
    const viewport = getPageView(pageNumber)?.pdfPage?.getViewport({ scale: 1 });
    return {
      width: viewport?.width || 612,
      height: viewport?.height || 792
    };
  }

  function getPageElementFromEvent(event) {
    return event.target?.closest?.(".page") ?? null;
  }

  function getPageNumber(pageElement) {
    return Number(pageElement?.dataset?.pageNumber) || 1;
  }

  function getCurrentPageNumber() {
    return Number(getApp()?.pdfViewer?.currentPageNumber) || 1;
  }

  function ensureStyle() {
    if (document.getElementById("asta-editor-style")) return;
    const style = document.createElement("style");
    style.id = "asta-editor-style";
    style.textContent = `
      #astaEditorToolbar {
        position: fixed;
        z-index: 10000;
        right: 18px;
        top: 56px;
        display: block;
        font: 12px system-ui, sans-serif;
      }
      #astaEditorToolbar .asta-editor-tool-toggle {
        height: 30px;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        background: #ffffff;
        color: #0f172a;
        padding: 0 10px;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.14);
        cursor: pointer;
      }
      #astaEditorToolbar .asta-editor-tool-panel {
        position: absolute;
        top: 36px;
        right: 0;
        display: none;
        align-items: center;
        gap: 6px;
        flex-wrap: wrap;
        width: min(760px, calc(100vw - 40px));
        max-height: min(70vh, 520px);
        overflow: auto;
        padding: 8px;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        background: #ffffff;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.18);
      }
      #astaEditorToolbar.expanded .asta-editor-tool-panel {
        display: flex;
      }
      #astaEditorToolbar button {
        height: 28px;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        background: #f8fafc;
        color: #0f172a;
        padding: 0 8px;
        cursor: pointer;
      }
      #astaEditorToolbar button.active {
        border-color: #2563eb;
        background: #dbeafe;
        color: #1d4ed8;
      }
      #astaEditorToolbar select,
      #astaEditorToolbar input {
        height: 28px;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        background: #fff;
      }
      #astaEditorToolbar select { width: 160px; }
      #astaEditorToolbar input[type="number"] { width: 56px; padding: 0 4px; }
      #astaEditorToolbar input[type="color"] { width: 34px; padding: 1px; }
      #astaEditorToolbar .asta-editor-toolbar-separator {
        width: 1px;
        height: 24px;
        background: #cbd5e1;
      }
      .asta-editor-layer {
        position: absolute;
        inset: 0;
        z-index: 6;
        pointer-events: auto;
      }
      .asta-editor-item {
        position: absolute;
        min-width: 20px;
        min-height: 16px;
        box-sizing: border-box;
        cursor: move;
        user-select: none;
      }
      .asta-editor-item.selected {
        outline: 2px solid #2563eb;
        outline-offset: 2px;
      }
      .asta-editor-resize-handle {
        position: absolute;
        right: -7px;
        bottom: -7px;
        width: 12px;
        height: 12px;
        border: 2px solid #ffffff;
        border-radius: 50%;
        background: #2563eb;
        box-shadow: 0 1px 4px rgba(15, 23, 42, 0.3);
        cursor: nwse-resize;
        display: none;
      }
      .asta-editor-item.selected .asta-editor-resize-handle {
        display: block;
      }
      .asta-editor-text {
        padding: 2px 4px;
        white-space: pre-wrap;
        line-height: 1.25;
        background: rgba(255, 255, 255, 0.01);
      }
      .asta-editor-textReplace {
        padding: 0;
        white-space: pre-wrap;
        line-height: 1.25;
        background: #ffffff;
      }
      .asta-editor-textHighlight {
        background: #facc15;
        opacity: 0.38;
        mix-blend-mode: multiply;
      }
      .asta-editor-whiteout {
        background: #ffffff;
        border: 1px dashed #94a3b8;
      }
      .asta-editor-rectangle {
        border: 2px solid #2563eb;
        background: rgba(37, 99, 235, 0.08);
      }
      .asta-editor-line,
      .asta-editor-arrow {
        min-height: 20px;
      }
      .asta-editor-line svg,
      .asta-editor-arrow svg,
      .asta-editor-ink svg {
        position: absolute;
        inset: 0;
        width: 100%;
        height: 100%;
        overflow: visible;
      }
      .asta-editor-image img,
      .asta-editor-signature img {
        display: block;
        width: 100%;
        height: 100%;
        object-fit: contain;
        pointer-events: none;
      }
      .asta-editor-stamp {
        display: flex;
        align-items: center;
        justify-content: center;
        border: 3px solid #dc2626;
        border-radius: 6px;
        color: #dc2626;
        font: 700 20px system-ui, sans-serif;
        letter-spacing: 0.08em;
        background: rgba(255, 255, 255, 0.02);
        text-transform: uppercase;
      }
      .asta-editor-text-content {
        display: block;
        min-width: 100%;
        min-height: 100%;
        outline: none;
      }
      .asta-editor-item.inline-editing {
        cursor: text;
      }
      .asta-editor-item.inline-editing .asta-editor-text-content {
        cursor: text;
        user-select: text;
        background: rgba(37, 99, 235, 0.08);
      }
      .asta-editor-ink {
        min-width: 4px;
        min-height: 4px;
        background: transparent;
      }
      .asta-editor-highlight {
        mix-blend-mode: multiply;
      }
      .asta-editor-crosshair .asta-editor-layer {
        cursor: crosshair;
      }
    `;
    document.head.appendChild(style);
  }

  function createToolbar() {
    if (toolbar) return;
    ensureStyle();
    toolbar = document.createElement("div");
    toolbar.id = "astaEditorToolbar";
    toolbar.innerHTML = `
      <button type="button" class="asta-editor-tool-toggle" data-action="toggleTools" aria-expanded="false" title="Show editor tools">Edit Tools</button>
      <div class="asta-editor-tool-panel" role="menu" aria-label="Editor tools">
        <button type="button" data-mode="select" title="Select">Select</button>
        <button type="button" data-mode="text" title="Add text">Text</button>
        <button type="button" data-mode="replaceText" title="Replace existing text">Replace</button>
        <button type="button" data-mode="whiteout" title="Cover an area with white">Whiteout</button>
        <button type="button" data-mode="redact" title="Cover an area with black redaction">Redact</button>
        <button type="button" data-mode="rectangle" title="Add rectangle">Rect</button>
        <button type="button" data-mode="ellipse" title="Add ellipse">Ellipse</button>
        <button type="button" data-mode="line" title="Add line">Line</button>
        <button type="button" data-mode="arrow" title="Add arrow">Arrow</button>
        <button type="button" data-mode="pen" title="Draw freehand pen">Pen</button>
        <button type="button" data-mode="highlight" title="Draw highlight">Highlight</button>
        <button type="button" data-mode="underline" title="Underline selected text">Underline</button>
        <button type="button" data-mode="strikeout" title="Strike out selected text">Strike</button>
        <button type="button" data-mode="stamp" title="Add stamp">Stamp</button>
        <select data-role="font" title="Font"></select>
        <input data-role="size" type="number" min="6" max="96" value="14" title="Text size" />
        <input data-role="color" type="color" value="#111827" title="Color" />
        <span class="asta-editor-toolbar-separator"></span>
        <input data-role="strokeWidth" type="number" min="1" max="24" value="2" title="Line or border width" />
        <input data-role="fillColor" type="color" value="#ffffff" title="Fill color" />
        <input data-role="opacity" type="number" min="5" max="100" step="5" value="100" title="Opacity percent" />
        <input data-role="rotation" type="number" min="-180" max="180" step="5" value="0" title="Rotation degrees" />
        <button type="button" data-action="copy" title="Copy selected">Copy</button>
        <button type="button" data-action="paste" title="Paste copied">Paste</button>
        <button type="button" data-action="duplicate" title="Duplicate selected">Duplicate</button>
        <button type="button" data-action="bringForward" title="Bring selected forward">Forward</button>
        <button type="button" data-action="sendBackward" title="Send selected backward">Backward</button>
        <button type="button" data-action="bringToFront" title="Bring selected to front">Front</button>
        <button type="button" data-action="sendToBack" title="Send selected to back">Back</button>
        <button type="button" data-action="delete" title="Delete selected">Delete</button>
      </div>
    `;
    document.body.appendChild(toolbar);
    toolbar.addEventListener("click", event => {
      const button = event.target.closest("button");
      if (!button) return;
      if (button.dataset.action === "toggleTools") {
        setToolbarExpanded(!toolbar.classList.contains("expanded"));
      } else if (button.dataset.mode) {
        setMode(button.dataset.mode);
        setToolbarExpanded(false);
      } else if (button.dataset.action === "copy") {
        copySelected();
        setToolbarExpanded(false);
      } else if (button.dataset.action === "paste") {
        pasteCopiedEdit();
        setToolbarExpanded(false);
      } else if (button.dataset.action === "duplicate") {
        duplicateSelected();
        setToolbarExpanded(false);
      } else if (button.dataset.action === "bringForward") {
        changeSelectedLayerOrder("forward");
        setToolbarExpanded(false);
      } else if (button.dataset.action === "sendBackward") {
        changeSelectedLayerOrder("backward");
        setToolbarExpanded(false);
      } else if (button.dataset.action === "bringToFront") {
        changeSelectedLayerOrder("front");
        setToolbarExpanded(false);
      } else if (button.dataset.action === "sendToBack") {
        changeSelectedLayerOrder("back");
        setToolbarExpanded(false);
      } else if (button.dataset.action === "delete") {
        deleteSelected();
        setToolbarExpanded(false);
      }
    });
    bindToolbarPropertyInput("font", event => {
      state.fontName = event.target.value || state.fontName;
      applySelectedProperties();
    });
    bindToolbarPropertyInput("size", event => {
      state.textSize = Number(event.target.value) || state.textSize;
      applySelectedProperties();
    });
    bindToolbarPropertyInput("color", event => {
      state.color = event.target.value || state.color;
      applySelectedProperties();
    });
    bindToolbarPropertyInput("strokeWidth", () => {
      applySelectedProperties();
    });
    bindToolbarPropertyInput("fillColor", () => {
      applySelectedProperties();
    });
    bindToolbarPropertyInput("opacity", () => {
      applySelectedProperties();
    });
    bindToolbarPropertyInput("rotation", () => {
      applySelectedProperties();
    });
    setFonts([]);
    setMode("select");
  }

  function setToolbarExpanded(expanded) {
    toolbar?.classList.toggle("expanded", expanded);
    toolbar?.querySelector("[data-action='toggleTools']")?.setAttribute("aria-expanded", String(expanded));
  }

  function bindToolbarPropertyInput(role, handler) {
    const control = toolbar?.querySelector(`[data-role='${role}']`);
    if (!control) return;
    let lastAppliedValue = control.value;
    control.dataset.lastAppliedValue = lastAppliedValue;
    const guardedHandler = event => {
      lastAppliedValue = control.dataset.lastAppliedValue ?? lastAppliedValue;
      const nextValue = event.target.value;
      if (nextValue === lastAppliedValue) return;
      lastAppliedValue = nextValue;
      control.dataset.lastAppliedValue = lastAppliedValue;
      handler(event);
    };
    control.addEventListener("input", guardedHandler);
    control.addEventListener("change", guardedHandler);
  }

  function setToolbarValue(control, value) {
    if (!control) return;
    control.value = value;
    control.dataset.lastAppliedValue = control.value;
  }

  function setMode(mode) {
    state.mode = mode;
    document.body.classList.toggle("asta-editor-crosshair", mode !== "select");
    toolbar?.querySelectorAll("button[data-mode]").forEach(button => {
      button.classList.toggle("active", button.dataset.mode === mode);
    });
    if (mode === "replaceText") {
      setTimeout(() => addSelectedTextReplacementEdit(), 0);
    } else if (mode === "whiteout") {
      setTimeout(() => {
        if (addSelectedTextWhiteoutEdits()) {
          setMode("select");
        }
      }, 0);
    } else if (mode === "redact") {
      setTimeout(() => {
        if (addSelectedTextRedactEdits()) {
          setMode("select");
        }
      }, 0);
    } else if (mode === "highlight") {
      setTimeout(() => {
        if (addSelectedTextHighlightEdits()) {
          setMode("select");
        }
      }, 0);
    } else if (mode === "underline" || mode === "strikeout") {
      setTimeout(() => {
        if (addSelectedTextLineMarkupEdits(mode)) {
          setMode("select");
        }
      }, 0);
    }
  }

  function setFonts(fonts) {
    state.fonts = Array.isArray(fonts) ? fonts : [];
    const select = toolbar?.querySelector("[data-role='font']");
    if (!select) return;
    const current = state.fontName;
    const names = state.fonts.length > 0
      ? state.fonts.map(font => font.name || font.family).filter(Boolean)
      : ["Malgun Gothic", "Arial"];
    select.replaceChildren(...names.map(name => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      return option;
    }));
    state.fontName = names.includes(current) ? current : names[0];
    select.value = state.fontName;
  }

  function ensureLayer(pageElement) {
    if (!pageElement) return null;
    pageElement.style.position ||= "relative";
    let layer = pageElement.querySelector(":scope > .asta-editor-layer");
    if (!layer) {
      layer = document.createElement("div");
      layer.className = "asta-editor-layer";
      layer.addEventListener("pointerdown", onLayerPointerDown);
      pageElement.appendChild(layer);
    }
    return layer;
  }

  function ensureLayers() {
    document.querySelectorAll(".page").forEach(ensureLayer);
  }

  function pagePoint(event, pageElement) {
    const rect = pageElement.getBoundingClientRect();
    return {
      x: event.clientX - rect.left,
      y: event.clientY - rect.top
    };
  }

  function onLayerPointerDown(event) {
    if (event.target !== event.currentTarget) return;
    const pageElement = getPageElementFromEvent(event);
    if (!pageElement || state.mode === "select") {
      selectEdit(null);
      return;
    }
    const point = pagePoint(event, pageElement);
    if (state.mode === "text") {
      const text = window.prompt("Text to add", "");
      if (!text) return;
      recordHistory();
      addEdit({
        type: "text",
        page: getPageNumber(pageElement),
        x: point.x,
        y: point.y,
        width: 180,
        height: Math.max(state.textSize * 1.6, 24),
        text,
        fontName: state.fontName,
        size: state.textSize,
        color: state.color
      });
    } else if (state.mode === "replaceText") {
      addTextReplacementEdit(event, pageElement);
    } else if (state.mode === "whiteout") {
      recordHistory();
      addEdit({
        type: "whiteout",
        page: getPageNumber(pageElement),
        x: point.x,
        y: point.y,
        width: 180,
        height: 90,
        fillColor: "#ffffff",
        borderColor: "#94a3b8",
        borderWidth: 1
      });
    } else if (state.mode === "redact") {
      recordHistory();
      addEdit({
        type: "whiteout",
        variant: "redact",
        page: getPageNumber(pageElement),
        x: point.x,
        y: point.y,
        width: 180,
        height: 90,
        fillColor: "#000000",
        borderColor: "#000000",
        borderWidth: 0
      });
    } else if (state.mode === "rectangle") {
      recordHistory();
      addEdit({
        type: "rectangle",
        page: getPageNumber(pageElement),
        x: point.x,
        y: point.y,
        width: 120,
        height: 70,
        borderColor: state.color,
        fillColor: null,
        borderWidth: 2
      });
    } else if (state.mode === "ellipse") {
      recordHistory();
      addEdit({
        type: "ellipse",
        page: getPageNumber(pageElement),
        x: point.x,
        y: point.y,
        width: 120,
        height: 70,
        borderColor: state.color,
        fillColor: null,
        borderWidth: 2
      });
    } else if (state.mode === "line") {
      recordHistory();
      addEdit({
        type: "line",
        page: getPageNumber(pageElement),
        x: point.x,
        y: point.y,
        width: 140,
        height: 24,
        borderColor: state.color,
        borderWidth: 2
      });
    } else if (state.mode === "arrow") {
      recordHistory();
      addEdit({
        type: "arrow",
        page: getPageNumber(pageElement),
        x: point.x,
        y: point.y,
        width: 140,
        height: 36,
        borderColor: state.color,
        borderWidth: 2
      });
    } else if (state.mode === "pen" || state.mode === "highlight") {
      startInk(event, pageElement);
    } else if (state.mode === "stamp") {
      const text = window.prompt("Stamp text", "APPROVED");
      if (!text) return;
      recordHistory();
      addEdit({
        type: "stamp",
        page: getPageNumber(pageElement),
        x: point.x,
        y: point.y,
        width: 180,
        height: 70,
        text,
        fontName: state.fontName,
        size: Math.max(state.textSize, 20),
        borderColor: "#dc2626",
        color: "#dc2626",
        borderWidth: 3
      });
    }
  }

  function addTextReplacementEdit(event, pageElement) {
    const textElement = findTextLayerElementAt(event.clientX, event.clientY);
    if (!textElement || !pageElement.contains(textElement)) {
      postMessage({
        type: "viewerDiagnostic",
        level: "info",
        message: "No selectable PDF text found at the clicked position."
      });
      return;
    }

    const originalText = textElement.textContent?.trim() || "";
    if (!originalText) return;
    const replacementText = window.prompt("Replacement text", originalText);
    if (replacementText === null) return;

    const pageRect = pageElement.getBoundingClientRect();
    const textRect = textElement.getBoundingClientRect();
    const metrics = getTextReplacementMetrics(textElement, textRect);
    createTextReplacementEdit(pageElement, textRect, originalText, replacementText, metrics);
  }

  function createTextReplacementEdit(pageElement, textRect, originalText, replacementText, metrics) {
    const pageRect = pageElement.getBoundingClientRect();
    recordHistory();
    addEdit({
      type: "textReplace",
      page: getPageNumber(pageElement),
      x: textRect.left - pageRect.left,
      y: textRect.top - pageRect.top,
      width: Math.max(textRect.width, 20),
      height: Math.max(textRect.height, metrics.lineHeight),
      text: replacementText,
      originalText,
      fontName: state.fontName,
      size: metrics.fontSize,
      lineHeight: metrics.lineHeight,
      textInsetX: metrics.textInsetX,
      textInsetY: metrics.textInsetY,
      whiteoutPadding: metrics.whiteoutPadding,
      color: state.color,
      fillColor: "#ffffff"
    });
  }

  function addSelectedTextReplacementEdit() {
    const target = getSelectedTextReplacementTarget();
    if (!target) return false;

    const replacementText = window.prompt("Replacement text", target.originalText);
    if (replacementText === null) return true;

    createTextReplacementEdit(
      target.pageElement,
      target.textRect,
      target.originalText,
      replacementText,
      target.metrics
    );
    window.getSelection()?.removeAllRanges();
    return true;
  }

  function addSelectedTextHighlightEdits() {
    const targets = getSelectedTextHighlightTargets();
    if (targets.length === 0) return false;

    recordHistory();
    for (const target of targets) {
      addEdit({
        type: "textHighlight",
        page: getPageNumber(target.pageElement),
        x: target.x,
        y: target.y,
        width: Math.max(target.width, 1),
        height: Math.max(target.height, 1),
        fillColor: "#facc15",
        opacity: 0.38,
        borderWidth: 0
      });
    }
    window.getSelection()?.removeAllRanges();
    return true;
  }

  function addSelectedTextWhiteoutEdits() {
    const targets = getSelectedTextWhiteoutTargets();
    if (targets.length === 0) return false;

    recordHistory();
    for (const target of targets) {
      addEdit({
        type: "whiteout",
        page: getPageNumber(target.pageElement),
        x: target.x,
        y: target.y,
        width: Math.max(target.width, 1),
        height: Math.max(target.height, 1),
        fillColor: "#ffffff",
        borderColor: "#ffffff",
        borderWidth: 0
      });
    }
    window.getSelection()?.removeAllRanges();
    return true;
  }

  function addSelectedTextRedactEdits() {
    const targets = getSelectedTextRedactTargets();
    if (targets.length === 0) return false;

    recordHistory();
    for (const target of targets) {
      addEdit({
        type: "whiteout",
        variant: "redact",
        page: getPageNumber(target.pageElement),
        x: target.x,
        y: target.y,
        width: Math.max(target.width, 1),
        height: Math.max(target.height, 1),
        fillColor: "#000000",
        borderColor: "#000000",
        borderWidth: 0
      });
    }
    window.getSelection()?.removeAllRanges();
    return true;
  }

  function addSelectedTextLineMarkupEdits(markupType) {
    const targets = getSelectedTextLineMarkupTargets(markupType);
    if (targets.length === 0) return false;

    recordHistory();
    for (const target of targets) {
      addEdit({
        type: "line",
        page: getPageNumber(target.pageElement),
        x: target.x,
        y: target.y,
        width: Math.max(target.width, 1),
        height: 1,
        borderColor: state.color,
        borderWidth: Math.max(Number(toolbar?.querySelector("[data-role='strokeWidth']")?.value) || 2, 1)
      });
    }
    window.getSelection()?.removeAllRanges();
    return true;
  }

  function getSelectedTextHighlightTargets() {
    const selection = window.getSelection?.();
    if (!selection || selection.isCollapsed || selection.rangeCount === 0) return [];

    const range = selection.getRangeAt(0);
    const clientRects = [...range.getClientRects()]
      .filter(rect => rect.width > 0 && rect.height > 0);
    if (clientRects.length === 0) return [];

    const pageElement = getPageElementForSelection(range, clientRects[0]);
    if (!pageElement) return [];

    const pageRect = pageElement.getBoundingClientRect();
    const selectedRects = clientRects.filter(rect => {
      const centerX = rect.left + rect.width / 2;
      const centerY = rect.top + rect.height / 2;
      const element = document.elementFromPoint(centerX, centerY);
      return pageElement.contains(element);
    });

    if (selectedRects.length !== clientRects.length) {
      postMessage({
        type: "viewerDiagnostic",
        level: "info",
        message: "Selected text highlight is limited to one PDF page at a time."
      });
      return [];
    }

    return selectedRects.map(rect => ({
      pageElement,
      x: rect.left - pageRect.left,
      y: rect.top - pageRect.top,
      width: rect.width,
      height: rect.height
    }));
  }

  function getSelectedTextLineMarkupTargets(markupType) {
    const selection = window.getSelection?.();
    if (!selection || selection.isCollapsed || selection.rangeCount === 0) return [];

    const range = selection.getRangeAt(0);
    const clientRects = [...range.getClientRects()]
      .filter(rect => rect.width > 0 && rect.height > 0);
    if (clientRects.length === 0) return [];

    const pageElement = getPageElementForSelection(range, clientRects[0]);
    if (!pageElement) return [];

    const pageRect = pageElement.getBoundingClientRect();
    const selectedRects = clientRects.filter(rect => {
      const centerX = rect.left + rect.width / 2;
      const centerY = rect.top + rect.height / 2;
      const element = document.elementFromPoint(centerX, centerY);
      return pageElement.contains(element);
    });

    if (selectedRects.length !== clientRects.length) {
      postMessage({
        type: "viewerDiagnostic",
        level: "info",
        message: "Selected text underline and strikeout markup is limited to one PDF page at a time."
      });
      return [];
    }

    return selectedRects.map(rect => {
      const y = markupType === "strikeout"
        ? rect.top + rect.height * 0.52
        : rect.bottom - Math.max(1, rect.height * 0.08);
      return {
        pageElement,
        x: rect.left - pageRect.left,
        y: y - pageRect.top,
        width: rect.width
      };
    });
  }

  function getSelectedTextRedactTargets() {
    const selection = window.getSelection?.();
    if (!selection || selection.isCollapsed || selection.rangeCount === 0) return [];

    const range = selection.getRangeAt(0);
    const clientRects = [...range.getClientRects()]
      .filter(rect => rect.width > 0 && rect.height > 0);
    if (clientRects.length === 0) return [];

    const pageElement = getPageElementForSelection(range, clientRects[0]);
    if (!pageElement) return [];

    const pageRect = pageElement.getBoundingClientRect();
    const selectedRects = clientRects.filter(rect => {
      const centerX = rect.left + rect.width / 2;
      const centerY = rect.top + rect.height / 2;
      const element = document.elementFromPoint(centerX, centerY);
      return pageElement.contains(element);
    });

    if (selectedRects.length !== clientRects.length) {
      postMessage({
        type: "viewerDiagnostic",
        level: "info",
        message: "Selected text redaction is limited to one PDF page at a time."
      });
      return [];
    }

    return selectedRects.map(rect => ({
      pageElement,
      x: rect.left - pageRect.left,
      y: rect.top - pageRect.top,
      width: rect.width,
      height: rect.height
    }));
  }

  function getSelectedTextWhiteoutTargets() {
    const selection = window.getSelection?.();
    if (!selection || selection.isCollapsed || selection.rangeCount === 0) return [];

    const range = selection.getRangeAt(0);
    const clientRects = [...range.getClientRects()]
      .filter(rect => rect.width > 0 && rect.height > 0);
    if (clientRects.length === 0) return [];

    const pageElement = getPageElementForSelection(range, clientRects[0]);
    if (!pageElement) return [];

    const pageRect = pageElement.getBoundingClientRect();
    const selectedRects = clientRects.filter(rect => {
      const centerX = rect.left + rect.width / 2;
      const centerY = rect.top + rect.height / 2;
      const element = document.elementFromPoint(centerX, centerY);
      return pageElement.contains(element);
    });

    if (selectedRects.length !== clientRects.length) {
      postMessage({
        type: "viewerDiagnostic",
        level: "info",
        message: "Selected text whiteout is limited to one PDF page at a time."
      });
      return [];
    }

    return selectedRects.map(rect => ({
      pageElement,
      x: rect.left - pageRect.left,
      y: rect.top - pageRect.top,
      width: rect.width,
      height: rect.height
    }));
  }

  function getSelectedTextReplacementTarget() {
    const selection = window.getSelection?.();
    if (!selection || selection.isCollapsed || selection.rangeCount === 0) return null;

    const originalText = selection.toString().trim();
    if (!originalText) return null;

    const range = selection.getRangeAt(0);
    const clientRects = [...range.getClientRects()]
      .filter(rect => rect.width > 0 && rect.height > 0);
    if (clientRects.length === 0) return null;

    const pageElement = getPageElementForSelection(range, clientRects[0]);
    if (!pageElement) return null;

    const pageRect = pageElement.getBoundingClientRect();
    const selectedRects = clientRects.filter(rect => {
      const centerX = rect.left + rect.width / 2;
      const centerY = rect.top + rect.height / 2;
      const element = document.elementFromPoint(centerX, centerY);
      return pageElement.contains(element);
    });

    if (selectedRects.length !== clientRects.length) {
      postMessage({
        type: "viewerDiagnostic",
        level: "info",
        message: "Selected text replacement is limited to one PDF page at a time."
      });
      return null;
    }

    const union = selectedRects.reduce((bounds, rect) => ({
      left: Math.min(bounds.left, rect.left),
      top: Math.min(bounds.top, rect.top),
      right: Math.max(bounds.right, rect.right),
      bottom: Math.max(bounds.bottom, rect.bottom)
    }), {
      left: selectedRects[0].left,
      top: selectedRects[0].top,
      right: selectedRects[0].right,
      bottom: selectedRects[0].bottom
    });

    const textRect = {
      left: union.left,
      top: union.top,
      width: union.right - union.left,
      height: union.bottom - union.top
    };
    const textElement = getSelectionTextElement(range, selectedRects[0]) ?? pageElement.querySelector(".textLayer span, .textLayer [role='presentation'], .textLayer [dir]");
    const metrics = getTextReplacementMetrics(textElement ?? pageElement, textRect);
    metrics.whiteoutPadding = Math.max(metrics.whiteoutPadding, 2);

    return {
      pageElement,
      textRect,
      originalText,
      metrics
    };
  }

  function getPageElementForSelection(range, firstRect) {
    const commonElement = getElementFromNode(range.commonAncestorContainer);
    const commonPage = commonElement?.closest?.(".page");
    if (commonPage) return commonPage;

    const centerX = firstRect.left + firstRect.width / 2;
    const centerY = firstRect.top + firstRect.height / 2;
    return document.elementFromPoint(centerX, centerY)?.closest?.(".page") ?? null;
  }

  function getSelectionTextElement(range, firstRect) {
    const commonElement = getElementFromNode(range.commonAncestorContainer);
    const textElement = commonElement?.closest?.(".textLayer span, .textLayer [role='presentation'], .textLayer [dir]");
    if (textElement) return textElement;

    const centerX = firstRect.left + firstRect.width / 2;
    const centerY = firstRect.top + firstRect.height / 2;
    return document.elementFromPoint(centerX, centerY)?.closest?.(".textLayer span, .textLayer [role='presentation'], .textLayer [dir]") ?? null;
  }

  function getElementFromNode(node) {
    if (!node) return null;
    return node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
  }

  function getTextReplacementMetrics(textElement, textRect) {
    const computedStyle = window.getComputedStyle(textElement);
    const fontSize = Number.parseFloat(computedStyle.fontSize) || state.textSize;
    const parsedLineHeight = Number.parseFloat(computedStyle.lineHeight);
    const lineHeight = Number.isFinite(parsedLineHeight)
      ? parsedLineHeight
      : Math.max(textRect.height, fontSize * 1.25);

    return {
      fontSize,
      lineHeight: Math.max(lineHeight, fontSize),
      textInsetX: 0,
      textInsetY: 0,
      whiteoutPadding: 1
    };
  }

  function findTextLayerElementAt(clientX, clientY) {
    const editorLayer = document.elementFromPoint(clientX, clientY)?.closest?.(".asta-editor-layer");
    if (editorLayer) {
      editorLayer.style.pointerEvents = "none";
    }
    try {
      const elements = document.elementsFromPoint(clientX, clientY);
      return elements.find(element =>
        element.matches?.(".textLayer span, .textLayer [role='presentation'], .textLayer [dir]")
        && element.textContent?.trim());
    } finally {
      if (editorLayer) {
        editorLayer.style.pointerEvents = "";
      }
    }
  }

  function addEdit(edit) {
    edit.id = edit.id || `edit-${nextEditId++}`;
    edit.zIndex = Number.isFinite(Number(edit.zIndex)) ? Number(edit.zIndex) : getNextLayerIndex(edit.page);
    state.edits.push(edit);
    renderEdit(edit);
    state.dirty = true;
    postDirty();
    selectEdit(edit.id);
  }

  function getNextLayerIndex(pageNumber) {
    const page = Number(pageNumber) || getCurrentPageNumber();
    return state.edits
      .filter(edit => Number(edit.page) === page)
      .reduce((next, edit) => Math.max(next, Number(edit.zIndex) || 0), 0) + 10;
  }

  function ensureLayerIndex(edit) {
    if (!Number.isFinite(Number(edit.zIndex))) {
      edit.zIndex = getNextLayerIndex(edit.page);
    }
    return Number(edit.zIndex) || 0;
  }

  function renderEdit(edit) {
    const pageElement = document.querySelector(`.page[data-page-number="${edit.page}"]`);
    const layer = ensureLayer(pageElement);
    if (!layer) return;
    if (edit.type === "ink") {
      normalizeInkBounds(edit);
    }
    let element = layer.querySelector(`[data-edit-id="${edit.id}"]`);
    if (!element) {
      element = document.createElement("div");
      element.dataset.editId = edit.id;
      element.className = `asta-editor-item asta-editor-${edit.type}`;
      if (edit.type === "textReplace") {
        element.classList.add("asta-editor-text-replace");
      }
      element.addEventListener("pointerdown", event => startDrag(event, edit.id));
      element.addEventListener("dblclick", event => beginInlineTextEdit(edit.id, event));
      layer.appendChild(element);
    }
    element.classList.toggle("asta-editor-text-replace", edit.type === "textReplace");
    element.classList.toggle("asta-editor-highlight", edit.type === "ink" && edit.tool === "highlight");
    element.classList.toggle("asta-editor-textHighlight", edit.type === "textHighlight");
    element.classList.toggle("asta-editor-whiteout", edit.type === "whiteout");
    element.classList.toggle("asta-editor-redact", edit.type === "whiteout" && edit.variant === "redact");
    element.style.left = `${edit.x}px`;
    element.style.top = `${edit.y}px`;
    element.style.width = `${edit.width}px`;
    element.style.height = `${edit.height}px`;
    element.style.zIndex = String(ensureLayerIndex(edit));
    element.style.transform = `rotate(${normalizeRotation(edit.rotate, 0)}deg)`;
    element.style.transformOrigin = "center center";
    if (edit.type === "text" || edit.type === "textReplace") {
      setTextElementContent(element, edit.text ?? "");
      element.style.fontFamily = `"${edit.fontName || state.fontName}", sans-serif`;
      element.style.fontSize = `${edit.size || state.textSize}px`;
      element.style.lineHeight = edit.lineHeight ? `${edit.lineHeight}px` : "1.25";
      element.style.color = edit.color || state.color;
      element.style.background = edit.type === "textReplace" ? (edit.fillColor || "#ffffff") : "rgba(255, 255, 255, 0.01)";
      element.style.padding = edit.type === "textReplace"
        ? `${edit.textInsetY || 0}px ${edit.textInsetX || 0}px`
        : "";
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "rectangle") {
      element.style.border = `${edit.borderWidth || 2}px solid ${edit.borderColor || state.color}`;
      element.style.background = edit.fillColor || "rgba(37, 99, 235, 0.08)";
      element.style.borderColor = edit.borderColor || state.color;
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "ellipse") {
      element.style.border = `${edit.borderWidth || 2}px solid ${edit.borderColor || state.color}`;
      element.style.borderRadius = "50%";
      element.style.background = edit.fillColor || "rgba(37, 99, 235, 0.08)";
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "textHighlight") {
      element.style.border = "0";
      element.style.background = edit.fillColor || "#facc15";
      element.style.opacity = String(edit.opacity ?? 0.38);
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "whiteout") {
      element.style.border = `${edit.borderWidth || 1}px dashed ${edit.borderColor || "#94a3b8"}`;
      element.style.background = edit.fillColor || "#ffffff";
      element.style.opacity = "1";
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "line") {
      renderLineSvg(element, edit, false);
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "arrow") {
      renderLineSvg(element, edit, true);
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "ink") {
      renderInkElement(element, edit);
    } else if (edit.type === "image" || edit.type === "signature") {
      renderImageElement(element, edit);
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "stamp") {
      setTextElementContent(element, edit.text ?? "STAMP");
      element.style.border = `${edit.borderWidth || 3}px solid ${edit.borderColor || "#dc2626"}`;
      element.style.color = edit.color || "#dc2626";
      element.style.fontFamily = `"${edit.fontName || state.fontName}", sans-serif`;
      element.style.fontSize = `${edit.size || Math.max(state.textSize, 20)}px`;
      ensureResizeHandle(element, edit.id);
    }
    element.style.opacity = edit.type === "ink" ? "1" : String(getEditOpacity(edit));
    element.classList.toggle("selected", state.selectedId === edit.id);
  }

  function normalizeInkBounds(edit) {
    const points = Array.isArray(edit.points) ? edit.points : [];
    if (points.length === 0) return;
    const strokeWidth = Math.max(Number(edit.borderWidth) || 2, 1);
    const padding = Math.max(strokeWidth * 2, 4);
    const bounds = points.reduce((current, point) => ({
      left: Math.min(current.left, Number(point.x) || 0),
      top: Math.min(current.top, Number(point.y) || 0),
      right: Math.max(current.right, Number(point.x) || 0),
      bottom: Math.max(current.bottom, Number(point.y) || 0)
    }), {
      left: Number(points[0].x) || 0,
      top: Number(points[0].y) || 0,
      right: Number(points[0].x) || 0,
      bottom: Number(points[0].y) || 0
    });
    edit.x = bounds.left - padding;
    edit.y = bounds.top - padding;
    edit.width = Math.max(bounds.right - bounds.left + padding * 2, padding * 2);
    edit.height = Math.max(bounds.bottom - bounds.top + padding * 2, padding * 2);
  }

  function ensureResizeHandle(element, editId) {
    let handle = element.querySelector(":scope > .asta-editor-resize-handle");
    if (!handle) {
      handle = document.createElement("span");
      handle.className = "asta-editor-resize-handle";
      handle.addEventListener("pointerdown", event => startResize(event, editId));
      element.appendChild(handle);
    }
  }

  function setTextElementContent(element, text) {
    let content = element.querySelector(":scope > .asta-editor-text-content");
    if (!content) {
      content = document.createElement("span");
      content.className = "asta-editor-text-content";
      element.prepend(content);
    }
    if (content.isContentEditable) return;
    content.textContent = text;
  }

  function renderImageElement(element, edit) {
    let image = element.querySelector(":scope > img");
    if (!image) {
      image = document.createElement("img");
      image.alt = edit.type === "signature" ? "signature overlay" : "image overlay";
      element.prepend(image);
    }
    image.src = edit.imageDataUrl || "";
  }

  function renderLineSvg(element, edit, withArrowHead) {
    const color = edit.borderColor || state.color;
    const width = Math.max(Number(edit.width) || 1, 1);
    const height = Math.max(Number(edit.height) || 1, 1);
    const strokeWidth = Math.max(Number(edit.borderWidth) || 2, 2);
    element.style.border = "0";
    element.style.background = "transparent";
    element.querySelector(":scope > svg")?.remove();

    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    svg.setAttribute("preserveAspectRatio", "none");

    const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
    line.setAttribute("x1", "0");
    line.setAttribute("y1", String(height / 2));
    line.setAttribute("x2", String(width));
    line.setAttribute("y2", String(height / 2));
    line.setAttribute("stroke", color);
    line.setAttribute("stroke-width", String(strokeWidth));
    line.setAttribute("stroke-linecap", "round");
    svg.appendChild(line);

    if (withArrowHead) {
      const headSize = Math.max(strokeWidth * 4, 10);
      const head = document.createElementNS("http://www.w3.org/2000/svg", "path");
      head.setAttribute("d", `M ${width} ${height / 2} L ${Math.max(width - headSize, 0)} ${Math.max(height / 2 - headSize / 2, 0)} M ${width} ${height / 2} L ${Math.max(width - headSize, 0)} ${Math.min(height / 2 + headSize / 2, height)}`);
      head.setAttribute("stroke", color);
      head.setAttribute("stroke-width", String(strokeWidth));
      head.setAttribute("stroke-linecap", "round");
      head.setAttribute("fill", "none");
      svg.appendChild(head);
    }

    element.prepend(svg);
  }

  function renderInkElement(element, edit) {
    element.style.border = "0";
    element.style.background = "transparent";
    element.querySelector(":scope > svg")?.remove();
    const points = Array.isArray(edit.points) ? edit.points : [];
    if (points.length === 0) return;

    const width = Math.max(Number(edit.width) || 1, 1);
    const height = Math.max(Number(edit.height) || 1, 1);
    const strokeWidth = Math.max(Number(edit.borderWidth) || 2, 1);
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    svg.setAttribute("preserveAspectRatio", "none");

    const polyline = document.createElementNS("http://www.w3.org/2000/svg", "polyline");
    polyline.setAttribute("points", points.map(point =>
      `${(Number(point.x) || 0) - edit.x},${(Number(point.y) || 0) - edit.y}`
    ).join(" "));
    polyline.setAttribute("stroke", edit.borderColor || (edit.tool === "highlight" ? "#facc15" : state.color));
    polyline.setAttribute("stroke-width", String(strokeWidth));
    polyline.setAttribute("stroke-linecap", "round");
    polyline.setAttribute("stroke-linejoin", "round");
    polyline.setAttribute("fill", "none");
    polyline.setAttribute("opacity", String(getEditOpacity(edit)));
    svg.appendChild(polyline);
    element.prepend(svg);
  }

  function startInk(event, pageElement) {
    event.preventDefault();
    event.stopPropagation();
    const layer = ensureLayer(pageElement);
    const firstPoint = pagePoint(event, pageElement);
    const isHighlight = state.mode === "highlight";
    recordHistory();
    const edit = {
      type: "ink",
      tool: isHighlight ? "highlight" : "pen",
      page: getPageNumber(pageElement),
      points: [firstPoint],
      x: firstPoint.x,
      y: firstPoint.y,
      width: 1,
      height: 1,
      borderColor: isHighlight ? "#facc15" : state.color,
      borderWidth: isHighlight ? 12 : 2,
      opacity: isHighlight ? 0.38 : 1
    };
    addEdit(edit);

    const pointerId = event.pointerId;
    layer.setPointerCapture(pointerId);
    const appendPoint = moveEvent => {
      const point = pagePoint(moveEvent, pageElement);
      const previous = edit.points[edit.points.length - 1];
      if (Math.hypot(point.x - previous.x, point.y - previous.y) < 1) return;
      edit.points.push(point);
      renderEdit(edit);
    };
    const finish = upEvent => {
      if (edit.points.length === 1) {
        edit.points.push({ x: edit.points[0].x + 0.5, y: edit.points[0].y + 0.5 });
      }
      layer.releasePointerCapture(pointerId);
      layer.removeEventListener("pointermove", appendPoint);
      layer.removeEventListener("pointerup", finish);
      renderEdit(edit);
      state.dirty = true;
      postDirty();
    };

    layer.addEventListener("pointermove", appendPoint);
    layer.addEventListener("pointerup", finish, { once: true });
  }

  function startDrag(event, editId) {
    if (event.target.closest?.(".asta-editor-text-content")?.isContentEditable) return;
    event.preventDefault();
    event.stopPropagation();
    selectEdit(editId);
    const edit = state.edits.find(item => item.id === editId);
    if (!edit) return;
    const startX = event.clientX;
    const startY = event.clientY;
    const initialX = edit.x;
    const initialY = edit.y;
    const initialPoints = Array.isArray(edit.points)
      ? edit.points.map(point => ({ x: Number(point.x) || 0, y: Number(point.y) || 0 }))
      : null;
    const pointerId = event.pointerId;
    let historyRecorded = false;
    let changed = false;
    const recordDragHistory = () => {
      if (historyRecorded) return;
      recordHistory();
      historyRecorded = true;
    };
    event.currentTarget.setPointerCapture(pointerId);

    function move(moveEvent) {
      const deltaX = moveEvent.clientX - startX;
      const deltaY = moveEvent.clientY - startY;
      if (Math.abs(deltaX) < 1 && Math.abs(deltaY) < 1) return;
      recordDragHistory();
      changed = true;
      if (initialPoints) {
        edit.points = initialPoints.map(point => ({
          x: Math.max(0, point.x + deltaX),
          y: Math.max(0, point.y + deltaY)
        }));
      } else {
        edit.x = Math.max(0, initialX + deltaX);
        edit.y = Math.max(0, initialY + deltaY);
      }
      renderEdit(edit);
    }

    function up() {
      event.currentTarget.releasePointerCapture(pointerId);
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      if (!changed) return;
      state.dirty = true;
      postDirty();
    }

    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up, { once: true });
  }

  function startResize(event, editId) {
    event.preventDefault();
    event.stopPropagation();
    selectEdit(editId);
    const edit = state.edits.find(item => item.id === editId);
    if (!edit) return;
    const startX = event.clientX;
    const startY = event.clientY;
    const initialWidth = Number(edit.width) || 1;
    const initialHeight = Number(edit.height) || 1;
    const pointerId = event.pointerId;
    const target = event.currentTarget;
    let historyRecorded = false;
    let changed = false;
    const recordResizeHistory = () => {
      if (historyRecorded) return;
      recordHistory();
      historyRecorded = true;
    };
    target.setPointerCapture(pointerId);

    function move(moveEvent) {
      const nextWidth = Math.max(20, initialWidth + moveEvent.clientX - startX);
      const nextHeight = Math.max(edit.type === "line" || edit.type === "arrow" ? 20 : 16, initialHeight + moveEvent.clientY - startY);
      if (nextWidth === edit.width && nextHeight === edit.height) return;
      recordResizeHistory();
      changed = true;
      edit.width = nextWidth;
      edit.height = nextHeight;
      renderEdit(edit);
    }

    function up() {
      target.releasePointerCapture(pointerId);
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      if (!changed) return;
      state.dirty = true;
      postDirty();
    }

    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up, { once: true });
  }

  function selectEdit(editId) {
    state.selectedId = editId;
    document.querySelectorAll(".asta-editor-item").forEach(element => {
      element.classList.toggle("selected", element.dataset.editId === editId);
    });
    const edit = state.edits.find(item => item.id === editId);
    syncToolbarFromEdit(edit);
  }

  function syncToolbarFromEdit(edit) {
    if (!toolbar || !edit) return;
    const fontInput = toolbar.querySelector("[data-role='font']");
    const sizeInput = toolbar.querySelector("[data-role='size']");
    const colorInput = toolbar.querySelector("[data-role='color']");
    const strokeWidthInput = toolbar.querySelector("[data-role='strokeWidth']");
    const fillColorInput = toolbar.querySelector("[data-role='fillColor']");
    const opacityInput = toolbar.querySelector("[data-role='opacity']");
    const rotationInput = toolbar.querySelector("[data-role='rotation']");

    if (edit.type === "text" || edit.type === "textReplace" || edit.type === "stamp") {
      setToolbarValue(fontInput, edit.fontName || state.fontName);
      setToolbarValue(sizeInput, edit.size || state.textSize);
      setToolbarValue(colorInput, edit.color || state.color);
    } else {
      setToolbarValue(colorInput, edit.borderColor || state.color);
    }
    setToolbarValue(strokeWidthInput, edit.borderWidth || (edit.type === "stamp" ? 3 : 2));
    setToolbarValue(fillColorInput, edit.fillColor || "#ffffff");
    setToolbarValue(opacityInput, Math.round(getEditOpacity(edit) * 100));
    setToolbarValue(rotationInput, normalizeRotation(edit.rotate, 0));
  }

  function getToolbarProperties() {
    return {
      fontName: toolbar.querySelector("[data-role='font']").value || state.fontName,
      size: Number(toolbar.querySelector("[data-role='size']").value) || state.textSize,
      color: toolbar.querySelector("[data-role='color']").value || state.color,
      strokeWidth: Math.max(Number(toolbar.querySelector("[data-role='strokeWidth']").value) || 2, 1),
      fillColor: toolbar.querySelector("[data-role='fillColor']").value || "#ffffff",
      opacity: normalizeOpacity(toolbar.querySelector("[data-role='opacity']").value),
      rotate: normalizeRotation(toolbar.querySelector("[data-role='rotation']").value, 0)
    };
  }

  function normalizeOpacity(value, fallback = 1) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return fallback;
    const normalized = numeric > 1 ? numeric / 100 : numeric;
    return Math.min(Math.max(normalized, 0.05), 1);
  }

  function normalizeRotation(value, fallback = 0) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return fallback;
    return Math.min(Math.max(numeric, -180), 180);
  }

  function getEditOpacity(edit) {
    return normalizeOpacity(edit?.opacity, edit?.type === "textHighlight" || edit?.tool === "highlight" ? 0.38 : 1);
  }

  function buildSelectedPropertySnapshot(edit, properties = null) {
    const snapshot = {
      opacity: properties ? properties.opacity : getEditOpacity(edit),
      rotate: properties ? properties.rotate : normalizeRotation(edit.rotate, 0)
    };
    if (edit.type === "text" || edit.type === "textReplace") {
      snapshot.fontName = properties ? properties.fontName : edit.fontName ?? "";
      snapshot.size = properties ? properties.size : Number(edit.size) || 0;
      snapshot.color = properties ? properties.color : edit.color ?? "";
      if (edit.type === "textReplace") {
        snapshot.fillColor = properties ? properties.fillColor : edit.fillColor ?? "";
      }
    } else if (edit.type === "stamp") {
      snapshot.fontName = properties ? properties.fontName : edit.fontName ?? "";
      snapshot.size = properties ? properties.size : Number(edit.size) || 0;
      snapshot.color = properties ? properties.color : edit.color ?? "";
      snapshot.borderColor = properties ? properties.color : edit.borderColor ?? "";
      snapshot.borderWidth = properties ? properties.strokeWidth : Number(edit.borderWidth) || 0;
    } else if (edit.type === "rectangle" || edit.type === "ellipse" || edit.type === "textHighlight" || edit.type === "whiteout") {
      snapshot.borderColor = properties ? properties.color : edit.borderColor ?? "";
      snapshot.borderWidth = properties ? properties.strokeWidth : Number(edit.borderWidth) || 0;
      snapshot.fillColor = properties ? properties.fillColor : edit.fillColor ?? "";
    } else if (edit.type === "line" || edit.type === "arrow" || edit.type === "ink") {
      snapshot.borderColor = properties ? properties.color : edit.borderColor ?? "";
      snapshot.borderWidth = properties ? properties.strokeWidth : Number(edit.borderWidth) || 0;
    }
    return snapshot;
  }

  function arePropertySnapshotsEqual(left, right) {
    return JSON.stringify(left) === JSON.stringify(right);
  }

  function applySelectedProperties() {
    const properties = getToolbarProperties();
    state.fontName = properties.fontName;
    state.textSize = properties.size;
    state.color = properties.color;

    const edit = state.edits.find(item => item.id === state.selectedId);
    if (!edit) return;
    const beforeProperties = buildSelectedPropertySnapshot(edit);
    const afterProperties = buildSelectedPropertySnapshot(edit, properties);
    if (arePropertySnapshotsEqual(beforeProperties, afterProperties)) return;
    recordHistory();
    edit.opacity = properties.opacity;
    edit.rotate = properties.rotate;
    if (edit.type === "text" || edit.type === "textReplace") {
      edit.fontName = properties.fontName;
      edit.size = properties.size;
      edit.color = properties.color;
      if (edit.type === "textReplace") {
        edit.fillColor = properties.fillColor;
      }
    } else if (edit.type === "stamp") {
      edit.fontName = properties.fontName;
      edit.size = properties.size;
      edit.color = properties.color;
      edit.borderColor = properties.color;
      edit.borderWidth = properties.strokeWidth;
    } else if (edit.type === "rectangle" || edit.type === "ellipse" || edit.type === "textHighlight" || edit.type === "whiteout") {
      edit.borderColor = properties.color;
      edit.borderWidth = properties.strokeWidth;
      edit.fillColor = properties.fillColor;
    } else if (edit.type === "line" || edit.type === "arrow" || edit.type === "ink") {
      edit.borderColor = properties.color;
      edit.borderWidth = properties.strokeWidth;
    }
    state.dirty = true;
    renderEdit(edit);
    postDirty();
  }

  function editText(editId) {
    beginInlineTextEdit(editId);
  }

  function beginInlineTextEdit(editId, event = null) {
    event?.preventDefault();
    event?.stopPropagation();
    const edit = state.edits.find(item => item.id === editId);
    if (!edit || (edit.type !== "text" && edit.type !== "textReplace" && edit.type !== "stamp")) return;
    const element = document.querySelector(`[data-edit-id="${editId}"]`);
    const content = element?.querySelector(":scope > .asta-editor-text-content");
    if (!element || !content || content.isContentEditable) return;

    selectEdit(editId);
    const originalText = edit.text ?? "";
    element.classList.add("inline-editing");
    content.contentEditable = "true";
    content.dataset.originalText = originalText;
    content.focus();
    selectContent(content);

    const finish = commit => {
      content.removeEventListener("keydown", onKeyDown);
      content.removeEventListener("blur", onBlur);
      commitInlineTextEdit(editId, commit);
    };
    const onKeyDown = keyEvent => {
      keyEvent.stopPropagation();
      if (keyEvent.isComposing) return;
      if (keyEvent.key === "Escape") {
        keyEvent.preventDefault();
        finish(false);
        return;
      }
      if (keyEvent.key === "Enter" && !keyEvent.shiftKey) {
        keyEvent.preventDefault();
        finish(true);
      }
    };
    const onBlur = () => finish(true);

    content.addEventListener("keydown", onKeyDown);
    content.addEventListener("blur", onBlur);
  }

  function commitInlineTextEdit(editId, commit) {
    const edit = state.edits.find(item => item.id === editId);
    const element = document.querySelector(`[data-edit-id="${editId}"]`);
    const content = element?.querySelector(":scope > .asta-editor-text-content");
    if (!edit || !element || !content) return;

    const originalText = content.dataset.originalText ?? edit.text ?? "";
    const nextText = normalizeEditableText(content.innerText);
    content.contentEditable = "false";
    delete content.dataset.originalText;
    element.classList.remove("inline-editing");

    if (!commit || nextText === originalText) {
      content.textContent = originalText;
      return;
    }

    recordHistory();
    edit.text = nextText;
    state.dirty = true;
    renderEdit(edit);
    postDirty();
  }

  function normalizeEditableText(text) {
    return String(text ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n$/, "");
  }

  function selectContent(element) {
    const range = document.createRange();
    range.selectNodeContents(element);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  }

  function deleteSelected() {
    if (!state.selectedId) return;
    recordHistory();
    const editId = state.selectedId;
    state.edits = state.edits.filter(edit => edit.id !== editId);
    document.querySelector(`[data-edit-id="${editId}"]`)?.remove();
    state.selectedId = null;
    state.dirty = true;
    postDirty();
  }

  function copySelected() {
    const edit = state.edits.find(item => item.id === state.selectedId);
    if (!edit) return false;
    copiedEdit = JSON.parse(JSON.stringify(edit));
    return true;
  }

  function createPastedEdit(sourceEdit, targetPage = null) {
    const edit = JSON.parse(JSON.stringify(sourceEdit));
    const offset = 18;
    edit.id = `edit-${nextEditId++}`;
    edit.page = Math.max(1, Number(targetPage || sourceEdit.page || getCurrentPageNumber()) || 1);
    if (Array.isArray(edit.points)) {
      edit.points = edit.points.map(point => ({
        x: Math.max(0, (Number(point.x) || 0) + offset),
        y: Math.max(0, (Number(point.y) || 0) + offset)
      }));
    } else {
      edit.x = Math.max(0, (Number(edit.x) || 0) + offset);
      edit.y = Math.max(0, (Number(edit.y) || 0) + offset);
    }
    return edit;
  }

  function pasteCopiedEdit(targetPage = null) {
    if (!copiedEdit) return false;
    recordHistory();
    const edit = createPastedEdit(copiedEdit, targetPage || getCurrentPageNumber());
    state.edits.push(edit);
    renderEdit(edit);
    selectEdit(edit.id);
    state.dirty = true;
    postDirty();
    return true;
  }

  function duplicateSelected() {
    const edit = state.edits.find(item => item.id === state.selectedId);
    if (!edit) return false;
    copiedEdit = JSON.parse(JSON.stringify(edit));
    return pasteCopiedEdit(edit.page);
  }

  function getOrderedPageEdits(pageNumber) {
    return state.edits
      .filter(edit => Number(edit.page) === Number(pageNumber))
      .sort((left, right) => ensureLayerIndex(left) - ensureLayerIndex(right));
  }

  function changeSelectedLayerOrder(direction) {
    const edit = state.edits.find(item => item.id === state.selectedId);
    if (!edit) return false;
    const ordered = getOrderedPageEdits(edit.page);
    const index = ordered.findIndex(item => item.id === edit.id);
    if (index < 0) return false;

    recordHistory();
    if (direction === "forward" && index < ordered.length - 1) {
      const next = ordered[index + 1];
      [edit.zIndex, next.zIndex] = [ensureLayerIndex(next), ensureLayerIndex(edit)];
    } else if (direction === "backward" && index > 0) {
      const previous = ordered[index - 1];
      [edit.zIndex, previous.zIndex] = [ensureLayerIndex(previous), ensureLayerIndex(edit)];
    } else if (direction === "front") {
      edit.zIndex = ordered.reduce((max, item) => Math.max(max, ensureLayerIndex(item)), 0) + 10;
    } else if (direction === "back") {
      edit.zIndex = ordered.reduce((min, item) => Math.min(min, ensureLayerIndex(item)), ensureLayerIndex(edit)) - 10;
    } else {
      historyStack.pop();
      return false;
    }

    getOrderedPageEdits(edit.page).forEach((item, layerIndex) => {
      item.zIndex = (layerIndex + 1) * 10;
      renderEdit(item);
    });
    selectEdit(edit.id);
    state.dirty = true;
    postDirty();
    return true;
  }

  function undo() {
    if (historyStack.length === 0) return false;
    redoStack.push(snapshotEdits());
    restoreSnapshot(historyStack.pop());
    return true;
  }

  function redo() {
    if (redoStack.length === 0) return false;
    historyStack.push(snapshotEdits());
    restoreSnapshot(redoStack.pop());
    return true;
  }

  function toRgbArray(hex) {
    const value = String(hex || "#000000").replace("#", "");
    const numeric = Number.parseInt(value.length === 3
      ? value.split("").map(char => char + char).join("")
      : value, 16);
    if (!Number.isFinite(numeric)) return [0, 0, 0];
    return [
      ((numeric >> 16) & 255) / 255,
      ((numeric >> 8) & 255) / 255,
      (numeric & 255) / 255
    ];
  }

  function exportEdits() {
    return state.edits.map(edit => {
      const pageElement = document.querySelector(`.page[data-page-number="${edit.page}"]`);
      const rect = pageElement?.getBoundingClientRect();
      const pageSize = getPageSize(edit.page);
      const scaleX = rect?.width ? pageSize.width / rect.width : 1;
      const scaleY = rect?.height ? pageSize.height / rect.height : 1;
      const zIndex = ensureLayerIndex(edit);
      const opacity = getEditOpacity(edit);
      if (edit.type === "text" || edit.type === "textReplace") {
        return {
          type: edit.type,
          page: edit.page,
          zIndex,
          opacity,
          x: edit.x * scaleX,
          y: edit.y * scaleY,
          width: edit.width * scaleX,
          height: edit.height * scaleY,
          text: edit.text ?? "",
          originalText: edit.originalText ?? "",
          fontName: edit.fontName || state.fontName,
          size: (edit.size || state.textSize) * scaleY,
          lineHeight: (edit.lineHeight || (edit.size || state.textSize) * 1.25) * scaleY,
          textInsetX: (edit.textInsetX || 0) * scaleX,
          textInsetY: (edit.textInsetY || 0) * scaleY,
          whiteoutPadding: (edit.whiteoutPadding || 0) * Math.max(scaleX, scaleY),
          color: toRgbArray(edit.color),
          fillColor: toRgbArray(edit.fillColor || "#ffffff"),
          borderWidth: (edit.borderWidth || 0) * Math.max(scaleX, scaleY),
          rotate: normalizeRotation(edit.rotate, 0)
        };
      }
      if (edit.type === "line" || edit.type === "arrow") {
        return {
          type: edit.type,
          page: edit.page,
          zIndex,
          opacity,
          x: edit.x * scaleX,
          y: edit.y * scaleY,
          width: edit.width * scaleX,
          height: edit.height * scaleY,
          borderColor: toRgbArray(edit.borderColor),
          borderWidth: (edit.borderWidth || 2) * Math.max(scaleX, scaleY),
          rotate: normalizeRotation(edit.rotate, 0)
        };
      }
      if (edit.type === "ink") {
        return {
          type: "ink",
          tool: edit.tool || "pen",
          page: edit.page,
          zIndex,
          points: (Array.isArray(edit.points) ? edit.points : []).map(point => ({
            x: (Number(point.x) || 0) * scaleX,
            y: (Number(point.y) || 0) * scaleY
          })),
          borderColor: toRgbArray(edit.borderColor || state.color),
          borderWidth: (edit.borderWidth || 2) * Math.max(scaleX, scaleY),
          opacity
        };
      }
      if (edit.type === "textHighlight") {
        return {
          type: "textHighlight",
          page: edit.page,
          zIndex,
          x: edit.x * scaleX,
          y: edit.y * scaleY,
          width: edit.width * scaleX,
          height: edit.height * scaleY,
          fillColor: toRgbArray(edit.fillColor || "#facc15"),
          opacity,
          rotate: normalizeRotation(edit.rotate, 0)
        };
      }
      if (edit.type === "whiteout") {
        return {
          type: "whiteout",
          variant: edit.variant,
          page: edit.page,
          zIndex,
          x: edit.x * scaleX,
          y: edit.y * scaleY,
          width: edit.width * scaleX,
          height: edit.height * scaleY,
          fillColor: toRgbArray(edit.fillColor || "#ffffff"),
          opacity,
          rotate: normalizeRotation(edit.rotate, 0)
        };
      }
      if (edit.type === "image" || edit.type === "signature") {
        return {
          type: edit.type,
          page: edit.page,
          zIndex,
          x: edit.x * scaleX,
          y: edit.y * scaleY,
          width: edit.width * scaleX,
          height: edit.height * scaleY,
          imageDataUrl: edit.imageDataUrl,
          imageMimeType: edit.imageMimeType,
          opacity,
          rotate: normalizeRotation(edit.rotate, 0)
        };
      }
      if (edit.type === "stamp") {
        return {
          type: "stamp",
          page: edit.page,
          zIndex,
          x: edit.x * scaleX,
          y: edit.y * scaleY,
          width: edit.width * scaleX,
          height: edit.height * scaleY,
          text: edit.text ?? "STAMP",
          color: toRgbArray(edit.color || "#dc2626"),
          borderColor: toRgbArray(edit.borderColor || "#dc2626"),
          borderWidth: (edit.borderWidth || 3) * Math.max(scaleX, scaleY),
          fontName: edit.fontName || state.fontName,
          size: (edit.size || state.textSize) * scaleY,
          opacity,
          rotate: normalizeRotation(edit.rotate, 0)
        };
      }
      return {
        type: "rectangle",
        shape: edit.type,
        page: edit.page,
        zIndex,
        x: edit.x * scaleX,
        y: edit.y * scaleY,
        width: edit.width * scaleX,
        height: edit.height * scaleY,
        borderColor: toRgbArray(edit.borderColor),
        borderWidth: (edit.borderWidth || 2) * Math.max(scaleX, scaleY),
        fillColor: edit.fillColor ? toRgbArray(edit.fillColor) : undefined,
        opacity,
        rotate: normalizeRotation(edit.rotate, 0)
      };
    });
  }

  function getUsedFontNames() {
    return [...new Set(state.edits
      .filter(edit => edit.type === "text" || edit.type === "textReplace" || edit.type === "stamp")
      .map(edit => edit.fontName || state.fontName)
      .filter(Boolean))];
  }

  function markClean() {
    state.dirty = false;
    postDirty();
  }

  function clear() {
    state.edits = [];
    state.selectedId = null;
    state.dirty = false;
    historyStack = [];
    redoStack = [];
    document.querySelectorAll(".asta-editor-item").forEach(element => element.remove());
    postDirty();
  }

  function isEditorTypingTarget(target) {
    return Boolean(target && (
      target.closest?.(".asta-editor-text-content")?.isContentEditable ||
      target.matches?.("input, textarea, select")
    ));
  }

  function initialize() {
    createToolbar();
    ensureLayers();
    const observer = new MutationObserver(() => ensureLayers());
    observer.observe(document.getElementById("viewer") || document.body, {
      childList: true,
      subtree: true
    });
    window.addEventListener("keydown", event => {
      if (event.defaultPrevented || event.isComposing || isEditorTypingTarget(event.target)) return;
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c" && state.selectedId) {
        if (copySelected()) event.preventDefault();
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v") {
        if (pasteCopiedEdit()) event.preventDefault();
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d") {
        if (duplicateSelected()) event.preventDefault();
        return;
      }
      if (event.key === "Delete") {
        deleteSelected();
      }
    });
  }

  window.EditorAdapter = {
    initialize,
    setMode,
    setFonts,
    clear,
    markClean,
    deleteSelected,
    copySelected,
    pasteCopiedEdit,
    duplicateSelected,
    changeSelectedLayerOrder,
    undo,
    redo,
    hasDirtyEdits: () => state.dirty,
    getEdits: exportEdits,
    getUsedFontNames,
    collectState: requestId => {
      postMessage({
        type: "editorStateCollected",
        requestId,
        edits: exportEdits(),
        usedFontNames: getUsedFontNames(),
        isDirty: state.dirty
      });
    }
  };
})();
