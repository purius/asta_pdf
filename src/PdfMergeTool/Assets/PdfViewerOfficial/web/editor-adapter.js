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
        display: flex;
        align-items: center;
        gap: 6px;
        padding: 6px;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        background: #ffffff;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.18);
        font: 12px system-ui, sans-serif;
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
      .asta-editor-rectangle {
        border: 2px solid #2563eb;
        background: rgba(37, 99, 235, 0.08);
      }
      .asta-editor-line,
      .asta-editor-arrow {
        min-height: 20px;
      }
      .asta-editor-line svg,
      .asta-editor-arrow svg {
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
      <button type="button" data-mode="select" title="Select">Select</button>
      <button type="button" data-mode="text" title="Add text">Text</button>
      <button type="button" data-mode="rectangle" title="Add rectangle">Rect</button>
      <button type="button" data-mode="ellipse" title="Add ellipse">Ellipse</button>
      <button type="button" data-mode="line" title="Add line">Line</button>
      <button type="button" data-mode="arrow" title="Add arrow">Arrow</button>
      <button type="button" data-mode="image" title="Add image">Image</button>
      <button type="button" data-mode="stamp" title="Add stamp">Stamp</button>
      <button type="button" data-mode="signature" title="Add signature image">Sign</button>
      <select data-role="font" title="Font"></select>
      <input data-role="size" type="number" min="6" max="96" value="14" title="Text size" />
      <input data-role="color" type="color" value="#111827" title="Color" />
      <button type="button" data-action="delete" title="Delete selected">Delete</button>
    `;
    document.body.appendChild(toolbar);
    toolbar.addEventListener("click", event => {
      const button = event.target.closest("button");
      if (!button) return;
      if (button.dataset.mode) {
        setMode(button.dataset.mode);
      } else if (button.dataset.action === "delete") {
        deleteSelected();
      }
    });
    toolbar.querySelector("[data-role='font']").addEventListener("change", event => {
      state.fontName = event.target.value || state.fontName;
      updateSelectedStyle();
    });
    toolbar.querySelector("[data-role='size']").addEventListener("change", event => {
      state.textSize = Number(event.target.value) || state.textSize;
      updateSelectedStyle();
    });
    toolbar.querySelector("[data-role='color']").addEventListener("change", event => {
      state.color = event.target.value || state.color;
      updateSelectedStyle();
    });
    setFonts([]);
    setMode("select");
  }

  function setMode(mode) {
    state.mode = mode;
    document.body.classList.toggle("asta-editor-crosshair", mode !== "select");
    toolbar?.querySelectorAll("button[data-mode]").forEach(button => {
      button.classList.toggle("active", button.dataset.mode === mode);
    });
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
    } else if (state.mode === "image" || state.mode === "signature") {
      addImageEdit(state.mode, pageElement, point);
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
        borderColor: "#dc2626",
        color: "#dc2626",
        borderWidth: 3
      });
    }
  }

  function addImageEdit(type, pageElement, point) {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/png,image/jpeg";
    input.addEventListener("change", () => {
      const file = input.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.addEventListener("load", () => {
        const dataUrl = String(reader.result || "");
        if (!dataUrl.startsWith("data:image/")) return;
        recordHistory();
        addEdit({
          type,
          page: getPageNumber(pageElement),
          x: point.x,
          y: point.y,
          width: type === "signature" ? 220 : 180,
          height: type === "signature" ? 90 : 130,
          imageDataUrl: dataUrl,
          imageMimeType: file.type || dataUrl.slice(5, dataUrl.indexOf(";"))
        });
      });
      reader.readAsDataURL(file);
    }, { once: true });
    input.click();
  }

  function addEdit(edit) {
    edit.id = edit.id || `edit-${nextEditId++}`;
    state.edits.push(edit);
    renderEdit(edit);
    state.dirty = true;
    postDirty();
    selectEdit(edit.id);
  }

  function renderEdit(edit) {
    const pageElement = document.querySelector(`.page[data-page-number="${edit.page}"]`);
    const layer = ensureLayer(pageElement);
    if (!layer) return;
    let element = layer.querySelector(`[data-edit-id="${edit.id}"]`);
    if (!element) {
      element = document.createElement("div");
      element.dataset.editId = edit.id;
      element.className = `asta-editor-item asta-editor-${edit.type}`;
      element.addEventListener("pointerdown", event => startDrag(event, edit.id));
      element.addEventListener("dblclick", () => editText(edit.id));
      layer.appendChild(element);
    }
    element.style.left = `${edit.x}px`;
    element.style.top = `${edit.y}px`;
    element.style.width = `${edit.width}px`;
    element.style.height = `${edit.height}px`;
    if (edit.type === "text") {
      setTextElementContent(element, edit.text ?? "");
      element.style.fontFamily = `"${edit.fontName || state.fontName}", sans-serif`;
      element.style.fontSize = `${edit.size || state.textSize}px`;
      element.style.color = edit.color || state.color;
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "rectangle") {
      element.style.border = `${edit.borderWidth || 2}px solid ${edit.borderColor || state.color}`;
      element.style.background = "rgba(37, 99, 235, 0.08)";
      element.style.borderColor = edit.borderColor || state.color;
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "ellipse") {
      element.style.border = `${edit.borderWidth || 2}px solid ${edit.borderColor || state.color}`;
      element.style.borderRadius = "50%";
      element.style.background = "rgba(37, 99, 235, 0.08)";
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "line") {
      renderLineSvg(element, edit, false);
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "arrow") {
      renderLineSvg(element, edit, true);
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "image" || edit.type === "signature") {
      renderImageElement(element, edit);
      ensureResizeHandle(element, edit.id);
    } else if (edit.type === "stamp") {
      setTextElementContent(element, edit.text ?? "STAMP");
      element.style.border = `${edit.borderWidth || 3}px solid ${edit.borderColor || "#dc2626"}`;
      element.style.color = edit.color || "#dc2626";
      ensureResizeHandle(element, edit.id);
    }
    element.classList.toggle("selected", state.selectedId === edit.id);
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

  function startDrag(event, editId) {
    event.preventDefault();
    event.stopPropagation();
    selectEdit(editId);
    const edit = state.edits.find(item => item.id === editId);
    if (!edit) return;
    const startX = event.clientX;
    const startY = event.clientY;
    const initialX = edit.x;
    const initialY = edit.y;
    const pointerId = event.pointerId;
    recordHistory();
    event.currentTarget.setPointerCapture(pointerId);

    function move(moveEvent) {
      edit.x = Math.max(0, initialX + moveEvent.clientX - startX);
      edit.y = Math.max(0, initialY + moveEvent.clientY - startY);
      renderEdit(edit);
    }

    function up() {
      event.currentTarget.releasePointerCapture(pointerId);
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
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
    recordHistory();
    target.setPointerCapture(pointerId);

    function move(moveEvent) {
      edit.width = Math.max(20, initialWidth + moveEvent.clientX - startX);
      edit.height = Math.max(edit.type === "line" || edit.type === "arrow" ? 20 : 16, initialHeight + moveEvent.clientY - startY);
      renderEdit(edit);
    }

    function up() {
      target.releasePointerCapture(pointerId);
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
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
    if (edit?.type === "text") {
      toolbar.querySelector("[data-role='font']").value = edit.fontName || state.fontName;
      toolbar.querySelector("[data-role='size']").value = edit.size || state.textSize;
      toolbar.querySelector("[data-role='color']").value = edit.color || state.color;
    }
  }

  function updateSelectedStyle() {
    const edit = state.edits.find(item => item.id === state.selectedId);
    if (!edit) return;
    recordHistory();
    if (edit.type === "text") {
      edit.fontName = state.fontName;
      edit.size = state.textSize;
      edit.color = state.color;
    } else if (edit.type === "rectangle" || edit.type === "ellipse" || edit.type === "line" || edit.type === "arrow" || edit.type === "stamp") {
      edit.borderColor = state.color;
      if (edit.type === "stamp") {
        edit.color = state.color;
      }
    }
    state.dirty = true;
    renderEdit(edit);
    postDirty();
  }

  function editText(editId) {
    const edit = state.edits.find(item => item.id === editId);
    if (!edit || (edit.type !== "text" && edit.type !== "stamp")) return;
    const text = window.prompt("Edit text", edit.text ?? "");
    if (text === null) return;
    recordHistory();
    edit.text = text;
    state.dirty = true;
    renderEdit(edit);
    postDirty();
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
      if (edit.type === "text") {
        return {
          type: "text",
          page: edit.page,
          x: edit.x * scaleX,
          y: edit.y * scaleY,
          width: edit.width * scaleX,
          height: edit.height * scaleY,
          text: edit.text ?? "",
          fontName: edit.fontName || state.fontName,
          size: (edit.size || state.textSize) * scaleY,
          color: toRgbArray(edit.color)
        };
      }
      if (edit.type === "line" || edit.type === "arrow") {
        return {
          type: edit.type,
          page: edit.page,
          x: edit.x * scaleX,
          y: edit.y * scaleY,
          width: edit.width * scaleX,
          height: edit.height * scaleY,
          borderColor: toRgbArray(edit.borderColor),
          borderWidth: (edit.borderWidth || 2) * scaleX
        };
      }
      if (edit.type === "image" || edit.type === "signature") {
        return {
          type: edit.type,
          page: edit.page,
          x: edit.x * scaleX,
          y: edit.y * scaleY,
          width: edit.width * scaleX,
          height: edit.height * scaleY,
          imageDataUrl: edit.imageDataUrl,
          imageMimeType: edit.imageMimeType
        };
      }
      if (edit.type === "stamp") {
        return {
          type: "stamp",
          page: edit.page,
          x: edit.x * scaleX,
          y: edit.y * scaleY,
          width: edit.width * scaleX,
          height: edit.height * scaleY,
          text: edit.text ?? "STAMP",
          color: toRgbArray(edit.color || "#dc2626"),
          borderColor: toRgbArray(edit.borderColor || "#dc2626"),
          borderWidth: (edit.borderWidth || 3) * scaleX
        };
      }
      return {
        type: "rectangle",
        shape: edit.type,
        page: edit.page,
        x: edit.x * scaleX,
        y: edit.y * scaleY,
        width: edit.width * scaleX,
        height: edit.height * scaleY,
        borderColor: toRgbArray(edit.borderColor),
        borderWidth: (edit.borderWidth || 2) * scaleX
      };
    });
  }

  function getUsedFontNames() {
    return [...new Set(state.edits
      .filter(edit => edit.type === "text")
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

  function initialize() {
    createToolbar();
    ensureLayers();
    const observer = new MutationObserver(() => ensureLayers());
    observer.observe(document.getElementById("viewer") || document.body, {
      childList: true,
      subtree: true
    });
    window.addEventListener("keydown", event => {
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
