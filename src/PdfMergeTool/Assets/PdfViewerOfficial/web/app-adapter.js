const AppBridge = (() => {
  let firstPageRendered = false;
  let latestRequestedLoadId = 0;
  let currentDocumentLoadId = 0;
  let pdfOpenQueue = Promise.resolve();
  const pendingPageNavigations = new Map();

  function postMessage(message) {
    window.chrome?.webview?.postMessage(message);
  }

  function postDiagnostic(level, message, details = {}) {
    postMessage({ type: "viewerDiagnostic", level, message, details });
  }

  function postDocumentMessage(message) {
    if (!Number.isInteger(currentDocumentLoadId) || currentDocumentLoadId < 1) {
      return;
    }

    postMessage({ ...message, loadId: currentDocumentLoadId });
  }

  function getApp() {
    return window.PDFViewerApplication ?? null;
  }

  function getTotalPages() {
    return getApp()?.pagesCount ?? getApp()?.pdfDocument?.numPages ?? 0;
  }

  async function waitForApplication() {
    while (!getApp()?.initializedPromise) {
      await new Promise(resolve => setTimeout(resolve, 20));
    }
    await getApp().initializedPromise;
    return getApp();
  }

  function configurePreviewOnlyViewer() {
    const app = getApp();
    app?.viewsManager?.close?.();

    const toggle = document.getElementById("viewsManagerToggleButton");
    toggle?.setAttribute("hidden", "true");
    toggle?.setAttribute("aria-hidden", "true");
    document.getElementById("viewsManager")?.setAttribute("hidden", "true");
  }

  function parseLoadId(value) {
    const loadId = Number(value);
    return Number.isInteger(loadId) && loadId > 0 ? loadId : null;
  }

  function queuePdfOpen(data) {
    const loadId = parseLoadId(data.loadId);
    if (!loadId) {
      postDiagnostic("error", "PDF load request is missing a valid load id.");
      return;
    }

    if (loadId < latestRequestedLoadId) {
      return;
    }

    latestRequestedLoadId = loadId;
    for (const pendingLoadId of pendingPageNavigations.keys()) {
      if (pendingLoadId < loadId) {
        pendingPageNavigations.delete(pendingLoadId);
      }
    }

    pdfOpenQueue = pdfOpenQueue
      .catch(() => undefined)
      .then(() => openPdf(data, loadId))
      .catch(error => {
        if (loadId === latestRequestedLoadId) {
          postMessage({
            type: "viewerLoadFailed",
            loadId,
            message: error?.message ?? String(error)
          });
        }
      });
  }

  async function openPdf(data, loadId) {
    if (loadId !== latestRequestedLoadId) {
      return;
    }

    const app = await waitForApplication();
    if (loadId !== latestRequestedLoadId) {
      return;
    }

    firstPageRendered = false;
    currentDocumentLoadId = 0;
    window.AstaViewerLoadId = 0;
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
    if (loadId !== latestRequestedLoadId) {
      return;
    }

    currentDocumentLoadId = loadId;
    window.AstaViewerLoadId = loadId;
    configurePreviewOnlyViewer();
    const deferredPage = pendingPageNavigations.get(loadId);
    pendingPageNavigations.delete(loadId);
    const initialPage = deferredPage ?? data.initialPage;
    if (Number.isInteger(initialPage) && initialPage > 0) {
      goToPage(initialPage);
    }
  }

  function goToPage(pageNumber) {
    const app = getApp();
    if (!app) return;

    const totalPages = Math.max(getTotalPages(), 1);
    const parsedPage = Number(pageNumber) || 1;
    app.page = Math.min(Math.max(parsedPage, 1), totalPages);
  }

  function goRelative(delta) {
    goToPage((getApp()?.page ?? 1) + delta);
  }

  function goToBoundary(last) {
    goToPage(last ? getTotalPages() : 1);
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

  async function exportOverlayPdf(data) {
    try {
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
        loadId: data.loadId,
        pdfBase64
      });
    } catch (error) {
      postMessage({
        type: "overlayPdfExportFailed",
        requestId: data.requestId ?? null,
        loadId: data.loadId,
        message: error?.message ?? String(error)
      });
      throw error;
    }
  }

  function handleCommand(command, options = {}, loadId = null) {
    const commandLoadId = parseLoadId(loadId);
    if (commandLoadId && commandLoadId !== currentDocumentLoadId) {
      if (commandLoadId === latestRequestedLoadId && command === "goToPage") {
        const pageNumber = Number(options?.pageNumber);
        if (Number.isInteger(pageNumber) && pageNumber > 0) {
          pendingPageNavigations.set(commandLoadId, pageNumber);
        }
      }

      return;
    }

    switch (command) {
      case "goToPage":
        goToPage(options?.pageNumber);
        break;
      case "nextPage":
        goRelative(1);
        break;
      case "prevPage":
        goRelative(-1);
        break;
      case "firstPage":
        goToBoundary(false);
        break;
      case "lastPage":
        goToBoundary(true);
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
      case "thumbZoomOut":
      case "thumbZoomReset":
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
      case "editorReplaceText":
        window.EditorAdapter?.setMode?.("replaceText");
        break;
      case "editorWhiteout":
        window.EditorAdapter?.setMode?.("whiteout");
        break;
      case "editorRedact":
        window.EditorAdapter?.setMode?.("redact");
        break;
      case "editorUnderline":
        window.EditorAdapter?.setMode?.("underline");
        break;
      case "editorStrikeout":
        window.EditorAdapter?.setMode?.("strikeout");
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
      case "editorPen":
        window.EditorAdapter?.setMode?.("pen");
        break;
      case "editorHighlight":
        window.EditorAdapter?.setMode?.("highlight");
        break;
      case "editorStamp":
        window.EditorAdapter?.setMode?.("stamp");
        break;
      case "editorDeleteSelection":
        window.EditorAdapter?.deleteSelected?.();
        break;
      case "editorCopySelection":
        window.EditorAdapter?.copySelected?.();
        break;
      case "editorPasteSelection":
        window.EditorAdapter?.pasteCopiedEdit?.();
        break;
      case "editorDuplicateSelection":
        window.EditorAdapter?.duplicateSelected?.();
        break;
      case "editorBringForward":
        window.EditorAdapter?.changeSelectedLayerOrder?.("forward");
        break;
      case "editorSendBackward":
        window.EditorAdapter?.changeSelectedLayerOrder?.("backward");
        break;
      case "editorBringToFront":
        window.EditorAdapter?.changeSelectedLayerOrder?.("front");
        break;
      case "editorSendToBack":
        window.EditorAdapter?.changeSelectedLayerOrder?.("back");
        break;
      case "markClean":
        window.EditorAdapter?.markClean?.();
        break;
      default:
        postDiagnostic("info", "Unsupported preview viewer command", { command });
        break;
    }
  }

  async function handleHostMessage(event) {
    const data = event.data ?? {};
    try {
      if (data.type === "loadPdf") {
        queuePdfOpen(data);
      } else if (data.type === "command") {
        handleCommand(data.command, data.options, data.loadId);
      } else if (data.type === "setEditorFonts") {
        window.EditorAdapter?.setFonts?.(data.fonts);
      } else if (data.type === "collectEditorState") {
        if (parseLoadId(data.loadId) === currentDocumentLoadId) {
          window.EditorAdapter?.collectState?.(data.requestId ?? null);
        }
      } else if (data.type === "exportOverlayPdf") {
        if (parseLoadId(data.loadId) === currentDocumentLoadId) {
          await exportOverlayPdf(data);
        }
      }
    } catch (error) {
      postDiagnostic("error", error?.message ?? String(error), { type: data.type });
    }
  }

  async function initialize() {
    const app = await waitForApplication();
    const eventBus = app.eventBus;
    window.EditorAdapter?.initialize?.();
    configurePreviewOnlyViewer();

    eventBus?._on("pagesinit", configurePreviewOnlyViewer);
    eventBus?._on("pagesloaded", configurePreviewOnlyViewer);
    eventBus?._on("pagerendered", event => {
      if (!firstPageRendered && currentDocumentLoadId > 0) {
        firstPageRendered = true;
        postDocumentMessage({
          type: "viewerFirstPageRendered",
          pageNumber: event.pageNumber ?? app.page ?? 1
        });
      }
    });
    eventBus?._on("pagechanging", event => {
      postDocumentMessage({
        type: "activePageChanged",
        activePage: event.pageNumber ?? app.page ?? 1
      });
    });

    window.chrome?.webview?.addEventListener("message", handleHostMessage);
    postMessage({ type: "viewerReady" });
  }

  return { initialize };
})();

AppBridge.initialize();
