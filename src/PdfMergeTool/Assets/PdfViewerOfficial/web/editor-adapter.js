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
      .asta-editor-text {
        padding: 2px 4px;
        white-space: pre-wrap;
        line-height: 1.25;
        background: rgba(255, 255, 255, 0.01);
      }
      .asta-editor-rect {
        border: 2px solid #2563eb;
        background: rgba(37, 99, 235, 0.08);
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
      <button type="button" data-mode="select" title="선택">선택</button>
      <button type="button" data-mode="text" title="텍스트 추가">텍스트</button>
      <button type="button" data-mode="rectangle" title="사각형 추가">사각형</button>
      <button type="button" data-mode="ellipse" title="원/타원 추가">타원</button>
      <button type="button" data-mode="line" title="선 추가">선</button>
      <select data-role="font" title="글꼴"></select>
      <input data-role="size" type="number" min="6" max="96" value="14" title="글자 크기" />
      <input data-role="color" type="color" value="#111827" title="색상" />
      <button type="button" data-action="delete" title="선택 삭제">삭제</button>
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
      const text = window.prompt("추가할 텍스트", "");
      if (!text) return;
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
      addEdit({
        type: "line",
        page: getPageNumber(pageElement),
        x: point.x,
        y: point.y,
        width: 140,
        height: 0,
        borderColor: state.color,
        borderWidth: 2
      });
    }
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
      element.textContent = edit.text ?? "";
      element.style.fontFamily = `"${edit.fontName || state.fontName}", sans-serif`;
      element.style.fontSize = `${edit.size || state.textSize}px`;
      element.style.color = edit.color || state.color;
    } else if (edit.type === "rectangle") {
      element.style.borderColor = edit.borderColor || state.color;
    } else if (edit.type === "ellipse") {
      element.style.border = `${edit.borderWidth || 2}px solid ${edit.borderColor || state.color}`;
      element.style.borderRadius = "50%";
      element.style.background = "rgba(37, 99, 235, 0.08)";
    } else if (edit.type === "line") {
      element.style.height = `${Math.max(edit.borderWidth || 2, 2)}px`;
      element.style.minHeight = `${Math.max(edit.borderWidth || 2, 2)}px`;
      element.style.background = edit.borderColor || state.color;
    }
    element.classList.toggle("selected", state.selectedId === edit.id);
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
    if (edit.type === "text") {
      edit.fontName = state.fontName;
      edit.size = state.textSize;
      edit.color = state.color;
    } else if (edit.type === "rectangle" || edit.type === "ellipse" || edit.type === "line") {
      edit.borderColor = state.color;
    }
    state.dirty = true;
    renderEdit(edit);
    postDirty();
  }

  function editText(editId) {
    const edit = state.edits.find(item => item.id === editId);
    if (!edit || edit.type !== "text") return;
    const text = window.prompt("텍스트 수정", edit.text ?? "");
    if (text === null) return;
    edit.text = text;
    state.dirty = true;
    renderEdit(edit);
    postDirty();
  }

  function deleteSelected() {
    if (!state.selectedId) return;
    const editId = state.selectedId;
    state.edits = state.edits.filter(edit => edit.id !== editId);
    document.querySelector(`[data-edit-id="${editId}"]`)?.remove();
    state.selectedId = null;
    state.dirty = true;
    postDirty();
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
      if (edit.type === "line") {
        return {
          type: "line",
          page: edit.page,
          x: edit.x * scaleX,
          y: edit.y * scaleY,
          width: edit.width * scaleX,
          height: edit.height * scaleY,
          borderColor: toRgbArray(edit.borderColor),
          borderWidth: (edit.borderWidth || 2) * scaleX
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
