const AppBridge = (() => {
  let pageOrder = [];
  let pageRotations = {};
  let selectedPages = new Set();
  let firstPageRendered = false;
  let pageStateDirty = false;
  let thumbnailScale = 1;

  function postMessage(message) {
    window.chrome?.webview?.postMessage(message);
  }

  function postDiagnostic(level, message, details = {}) {
    postMessage({ type: "viewerDiagnostic", level, message, details });
  }

  function getApp() {
    return window.PDFViewerApplication ?? null;
  }

  function getTotalPages() {
    return getApp()?.pagesCount ?? getApp()?.pdfDocument?.numPages ?? 0;
  }

  function rebuildPageOrder(totalPages = getTotalPages()) {
    pageOrder = Array.from({ length: totalPages }, (_, index) => index + 1);
    pageRotations = {};
    selectedPages = new Set(pageOrder.length > 0 ? [pageOrder[0]] : []);
    postPageOrder();
  }

  function postPageOrder() {
    const activePage = getApp()?.page ?? pageOrder[0] ?? 1;
    applyPageStatePresentation();
    postMessage({
      type: "pageOrderChanged",
      pageOrder,
      rotations: pageRotations,
      selectedPages: [...selectedPages],
      activePage,
      isDirty: pageStateDirty || Boolean(window.EditorAdapter?.hasDirtyEdits?.())
    });
  }

  async function waitForApplication() {
    while (!getApp()?.initializedPromise) {
      await new Promise(resolve => setTimeout(resolve, 20));
    }
    await getApp().initializedPromise;
    return getApp();
  }

  async function openPdf(data) {
    const app = await waitForApplication();
    firstPageRendered = false;
    pageStateDirty = false;
    window.EditorAdapter?.clear?.();

    let args;
    if (data.url) {
      args = { url: data.url };
    } else if (typeof data === "string") {
      args = { data: Uint8Array.from(atob(data), char => char.charCodeAt(0)) };
    } else if (data.base64) {
      args = { data: Uint8Array.from(atob(data.base64), char => char.charCodeAt(0)) };
    } else {
      throw new Error("Unsupported PDF source.");
    }

    await app.open(args);
  }

  function goToPage(pageNumber) {
    const app = getApp();
    if (!app) return;
    const totalPages = getTotalPages();
    app.page = Math.min(Math.max(pageNumber, 1), Math.max(totalPages, 1));
  }

  function goRelative(delta) {
    goToPage((getApp()?.page ?? 1) + delta);
  }

  function setScale(value) {
    const app = getApp();
    if (!app?.pdfViewer) return;
    app.pdfViewer.currentScaleValue = value;
  }

  function zoom(delta) {
    const app = getApp();
    if (!app?.pdfViewer) return;
    const current = Number(app.pdfViewer.currentScale) || 1;
    app.pdfViewer.currentScale = Math.min(Math.max(current + delta, 0.25), 4);
  }

  function rotateSelectedPages(delta) {
    const pages = selectedPages.size > 0 ? [...selectedPages] : [getApp()?.page ?? 1];
    for (const page of pages) {
      pageRotations[page] = ((pageRotations[page] ?? 0) + delta + 360) % 360;
    }
    pageStateDirty = true;
    postPageOrder();
  }

  function setThumbnailScale(nextScale) {
    thumbnailScale = Math.min(Math.max(nextScale, 0.75), 1.8);
    const thumbnailsView = document.getElementById("thumbnailsView");
    thumbnailsView?.style.setProperty("--thumbnail-width", `${Math.round(126 * thumbnailScale)}px`);
  }

  function applyPageStatePresentation() {
    const visiblePages = new Set(pageOrder);
    document.querySelectorAll(".page[data-page-number]").forEach(page => {
      const pageNumber = Number(page.dataset.pageNumber);
      page.hidden = visiblePages.size > 0 && !visiblePages.has(pageNumber);
    });
    document.querySelectorAll("#thumbnailsView .thumbnail[page-number]").forEach(thumbnail => {
      const pageNumber = Number(thumbnail.getAttribute("page-number"));
      thumbnail.hidden = visiblePages.size > 0 && !visiblePages.has(pageNumber);
    });
  }

  function deleteSelectedPages() {
    if (pageOrder.length <= 1 || selectedPages.size === 0) return;
    pageOrder = pageOrder.filter(page => !selectedPages.has(page));
    selectedPages = new Set([pageOrder[0]]);
    pageStateDirty = true;
    postPageOrder();
    goToPage(pageOrder[0]);
  }

  function reversePageOrder() {
    if (pageOrder.length <= 1) return;
    pageOrder = [...pageOrder].reverse();
    selectedPages = new Set([pageOrder[0]]);
    pageStateDirty = true;
    postPageOrder();
    goToPage(pageOrder[0]);
  }

  async function exportOverlayPdf(data) {
    if (!window.PdfLibAdapter?.createOverlayPdf) {
      throw new Error("PDF editor save adapter is not loaded.");
    }

    const pdfBase64 = await window.PdfLibAdapter.createOverlayPdf({
      ...data,
      edits: Array.isArray(data.edits) ? data.edits : window.EditorAdapter?.getEdits?.() ?? []
    });
    postMessage({
      type: "overlayPdfExported",
      requestId: data.requestId ?? null,
      pdfBase64
    });
  }

  function handleCommand(command) {
    switch (command) {
      case "nextPage":
        goRelative(1);
        break;
      case "prevPage":
        goRelative(-1);
        break;
      case "firstPage":
        goToPage(1);
        break;
      case "lastPage":
        goToPage(getTotalPages());
        break;
      case "mainZoomIn":
        zoom(0.15);
        break;
      case "mainZoomOut":
        zoom(-0.15);
        break;
      case "mainZoomReset":
        setScale("auto");
        break;
      case "fitPage":
        setScale("page-fit");
        break;
      case "thumbZoomIn":
        setThumbnailScale(thumbnailScale + 0.15);
        break;
      case "thumbZoomOut":
        setThumbnailScale(thumbnailScale - 0.15);
        break;
      case "thumbZoomReset":
        setThumbnailScale(1);
        break;
      case "rotateSelectedClockwise":
        rotateSelectedPages(90);
        break;
      case "rotateSelectedCounterClockwise":
        rotateSelectedPages(-90);
        break;
      case "rotateSelected180":
        rotateSelectedPages(180);
        break;
      case "deleteSelectedPages":
        deleteSelectedPages();
        break;
      case "reversePageOrder":
        reversePageOrder();
        break;
      case "undo":
        if (!window.EditorAdapter?.undo?.()) {
          postDiagnostic("info", "No editor action to undo.");
        }
        break;
      case "redo":
        if (!window.EditorAdapter?.redo?.()) {
          postDiagnostic("info", "No editor action to redo.");
        }
        break;
      case "editorSelect":
        window.EditorAdapter?.setMode?.("select");
        break;
      case "editorText":
        window.EditorAdapter?.setMode?.("text");
        break;
      case "editorRectangle":
        window.EditorAdapter?.setMode?.("rectangle");
        break;
      case "editorEllipse":
        window.EditorAdapter?.setMode?.("ellipse");
        break;
      case "editorLine":
        window.EditorAdapter?.setMode?.("line");
        break;
      case "editorArrow":
        window.EditorAdapter?.setMode?.("arrow");
        break;
      case "editorImage":
        window.EditorAdapter?.setMode?.("image");
        break;
      case "editorStamp":
        window.EditorAdapter?.setMode?.("stamp");
        break;
      case "editorSignature":
        window.EditorAdapter?.setMode?.("signature");
        break;
      case "editorDeleteSelection":
        window.EditorAdapter?.deleteSelected?.();
        break;
      case "markClean":
        pageStateDirty = false;
        window.EditorAdapter?.markClean?.();
        postPageOrder();
        break;
      default:
        postDiagnostic("info", "Unsupported official viewer command", { command });
        break;
    }
  }

  async function handleHostMessage(event) {
    const data = event.data ?? {};
    try {
      if (data.type === "loadPdf") {
        const source = data.url ? { url: data.url } : data.base64;
        await openPdf(source);
      } else if (data.type === "command") {
        handleCommand(data.command);
      } else if (data.type === "setEditorFonts") {
        window.EditorAdapter?.setFonts?.(data.fonts);
      } else if (data.type === "collectEditorState") {
        window.EditorAdapter?.collectState?.(data.requestId ?? null);
      } else if (data.type === "exportOverlayPdf") {
        await exportOverlayPdf(data);
      }
    } catch (error) {
      postDiagnostic("error", error?.message ?? String(error), { type: data.type });
    }
  }

  async function initialize() {
    const app = await waitForApplication();
    const eventBus = app.eventBus;
    window.EditorAdapter?.initialize?.();

    eventBus?._on("pagesinit", () => {
      rebuildPageOrder(getTotalPages());
    });

    eventBus?._on("pagesloaded", event => {
      rebuildPageOrder(event.pagesCount ?? getTotalPages());
      setTimeout(applyPageStatePresentation, 0);
    });

    eventBus?._on("pagerendered", event => {
      setTimeout(applyPageStatePresentation, 0);
      if (!firstPageRendered) {
        firstPageRendered = true;
        postMessage({
          type: "viewerFirstPageRendered",
          pageNumber: event.pageNumber ?? app.page ?? 1
        });
      }
    });

    eventBus?._on("pagechanging", event => {
      const activePage = event.pageNumber ?? app.page ?? 1;
      selectedPages = new Set([activePage]);
      postMessage({ type: "activePageChanged", activePage });
      postPageOrder();
    });

    window.chrome?.webview?.addEventListener("message", handleHostMessage);
    postMessage({ type: "viewerReady" });
  }

  return { initialize };
})();

AppBridge.initialize();
