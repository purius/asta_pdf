const AppBridge = (() => {
  let pageOrder = [];
  let pageRotations = {};
  let selectedPages = new Set();
  let firstPageRendered = false;

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
    postMessage({
      type: "pageOrderChanged",
      pageOrder,
      rotations: pageRotations,
      selectedPages: [...selectedPages],
      activePage,
      isDirty: false
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
    postPageOrder();
  }

  function deleteSelectedPages() {
    if (pageOrder.length <= 1 || selectedPages.size === 0) return;
    pageOrder = pageOrder.filter(page => !selectedPages.has(page));
    selectedPages = new Set([pageOrder[0]]);
    postPageOrder();
    goToPage(pageOrder[0]);
  }

  async function exportOverlayPdf(data) {
    if (!window.PdfLibAdapter?.createOverlayPdf) {
      throw new Error("PDF editor save adapter is not loaded.");
    }

    const pdfBase64 = await window.PdfLibAdapter.createOverlayPdf(data);
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
      case "markClean":
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

    eventBus?._on("pagesinit", () => {
      rebuildPageOrder(getTotalPages());
    });

    eventBus?._on("pagesloaded", event => {
      rebuildPageOrder(event.pagesCount ?? getTotalPages());
    });

    eventBus?._on("pagerendered", event => {
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
