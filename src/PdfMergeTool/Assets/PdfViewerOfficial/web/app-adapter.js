const AppBridge = (() => {
  let pageOrder = [];
  let pageRotations = {};
  let selectedPages = new Set();
  let firstPageRendered = false;
  let pageStateDirty = false;
  let thumbnailScale = 1;
  let explicitNavigationTarget = null;
  let explicitNavigationExpiresAt = 0;
  let explicitNavigationSettledUntil = 0;
  let lastAcceptedExplicitNavigationPage = null;
  let protectedExplicitNavigationPage = null;
  let lastUserPageChangeIntentAt = 0;

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

  function resetExplicitNavigationTracking() {
    explicitNavigationTarget = null;
    explicitNavigationExpiresAt = 0;
    explicitNavigationSettledUntil = 0;
    lastAcceptedExplicitNavigationPage = null;
    protectedExplicitNavigationPage = null;
    lastUserPageChangeIntentAt = 0;
  }

  function rebuildPageOrder(totalPages = getTotalPages()) {
    resetExplicitNavigationTracking();
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
    resetExplicitNavigationTracking();
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
    if (Number.isInteger(data.initialPage) && data.initialPage > 0) {
      goToPage(data.initialPage);
    }
  }

  function goToPage(pageNumber) {
    const app = getApp();
    if (!app) return;
    const totalPages = getTotalPages();
    const targetPage = Math.min(Math.max(pageNumber, 1), Math.max(totalPages, 1));
    beginExplicitPageNavigation(targetPage);
    app.page = targetPage;
  }

  function getVisiblePageOrder() {
    const totalPages = getTotalPages();
    const validPages = pageOrder.filter(page => page >= 1 && page <= totalPages);
    return validPages.length > 0
      ? validPages
      : Array.from({ length: totalPages }, (_, index) => index + 1);
  }

  function goRelativeInPageOrder(delta) {
    const orderedPages = getVisiblePageOrder();
    if (orderedPages.length === 0) {
      return;
    }

    const currentPage = getApp()?.page ?? orderedPages[0];
    const currentIndex = orderedPages.indexOf(currentPage);
    const nextIndex = Math.min(
      Math.max((currentIndex >= 0 ? currentIndex : 0) + delta, 0),
      orderedPages.length - 1);
    goToPage(orderedPages[nextIndex]);
  }

  function goToPageOrderBoundary(last) {
    const orderedPages = getVisiblePageOrder();
    if (orderedPages.length === 0) {
      return;
    }

    if (last) {
      goToPage(orderedPages[orderedPages.length - 1]);
      return;
    }

    goToPage(orderedPages[0]);
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

  function ensureDropIndicatorStyle() {
    if (document.getElementById("asta-drop-indicator-style")) {
      return;
    }

    const style = document.createElement("style");
    style.id = "asta-drop-indicator-style";
    style.textContent = `
      #thumbnailsView .thumbnail.asta-drop-target::before,
      #thumbnailsView .thumbnail.asta-drop-target::after {
        content: "";
        position: absolute;
        top: 6px;
        bottom: 6px;
        width: 3px;
        border-radius: 2px;
        background: #2563eb;
        box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.18);
      }
      #thumbnailsView .thumbnail.asta-drop-before::before {
        left: 2px;
      }
      #thumbnailsView .thumbnail.asta-drop-after::after {
        right: 2px;
      }
    `;
    document.head.appendChild(style);
  }

  function beginExplicitPageNavigation(pageNumber) {
    explicitNavigationTarget = Number(pageNumber) || null;
    explicitNavigationExpiresAt = Date.now() + 900;
    explicitNavigationSettledUntil = 0;
    lastAcceptedExplicitNavigationPage = null;
    protectedExplicitNavigationPage = null;
  }

  function markUserPageChangeIntent() {
    lastUserPageChangeIntentAt = Date.now();
  }

  function markKeyboardPageChangeIntent(event) {
    if (event.defaultPrevented || event.isComposing) return;
    const navigationKeys = new Set([
      "ArrowUp",
      "ArrowDown",
      "ArrowLeft",
      "ArrowRight",
      "PageUp",
      "PageDown",
      "Home",
      "End",
      " "
    ]);
    if (navigationKeys.has(event.key)) {
      markUserPageChangeIntent();
    }
  }

  function markNativeViewerNavigationIntent(event) {
    if (event.defaultPrevented || event.isComposing) return;
    if ("button" in event && event.button !== 0) return;
    const target = event.target?.closest?.("#previous, #next, #firstPage, #lastPage, #pageNumber");
    if (target) {
      markUserPageChangeIntent();
    }
  }

  function hasRecentUserPageChangeIntent() {
    return Date.now() - lastUserPageChangeIntentAt <= 1200;
  }

  function shouldAcceptPageChange(pageNumber) {
    if (lastAcceptedExplicitNavigationPage && Date.now() > explicitNavigationSettledUntil) {
      explicitNavigationSettledUntil = 0;
      lastAcceptedExplicitNavigationPage = null;
    }

    if (
      lastAcceptedExplicitNavigationPage &&
      Date.now() <= explicitNavigationSettledUntil &&
      Number(pageNumber) !== lastAcceptedExplicitNavigationPage
    ) {
      return false;
    }

    if (
      protectedExplicitNavigationPage &&
      Number(pageNumber) !== protectedExplicitNavigationPage &&
      !hasRecentUserPageChangeIntent()
    ) {
      return false;
    }

    if (
      protectedExplicitNavigationPage &&
      Number(pageNumber) !== protectedExplicitNavigationPage &&
      hasRecentUserPageChangeIntent()
    ) {
      protectedExplicitNavigationPage = null;
    }

    if (!explicitNavigationTarget) {
      return true;
    }

    if (Date.now() > explicitNavigationExpiresAt) {
      explicitNavigationTarget = null;
      explicitNavigationExpiresAt = 0;
      explicitNavigationSettledUntil = 0;
      lastAcceptedExplicitNavigationPage = null;
      protectedExplicitNavigationPage = null;
      return true;
    }

    if (Number(pageNumber) === explicitNavigationTarget) {
      lastAcceptedExplicitNavigationPage = Number(pageNumber);
      protectedExplicitNavigationPage = Number(pageNumber);
      explicitNavigationSettledUntil = Date.now() + 450;
      explicitNavigationTarget = null;
      explicitNavigationExpiresAt = 0;
      return true;
    }

    return false;
  }

  function findNavigationTargetPage(element) {
    const thumbnail = element?.closest?.("#thumbnailsView .thumbnail[page-number], #thumbnailView .thumbnail[page-number]");
    if (thumbnail) {
      return Number(thumbnail.getAttribute("page-number"));
    }

    const link = element?.closest?.("a[href*='#page=']");
    const match = /[#&]page=(\d+)/.exec(link?.getAttribute("href") || "");
    return match ? Number(match[1]) : null;
  }

  function captureExplicitNavigationIntent(event) {
    if (event.defaultPrevented) return;
    if (event.type === "pointerdown" && event.isPrimary === false) return;
    if ("button" in event && event.button !== 0) return;
    if (event.type === "click" && event.detail === 0) return;

    const targetPage = findNavigationTargetPage(event.target);
    if (targetPage) {
      beginExplicitPageNavigation(targetPage);
    }
  }

  function getThumbnailFromPoint(clientX, clientY) {
    return document
      .elementFromPoint(clientX, clientY)
      ?.closest?.("#thumbnailsView .thumbnail[page-number]") ?? null;
  }

  function getDropInsertionIndex(clientX, clientY) {
    const thumbnail = getThumbnailFromPoint(clientX, clientY);
    if (!thumbnail) {
      return pageOrder.length;
    }

    const targetPage = Number(thumbnail.getAttribute("page-number"));
    const targetIndex = pageOrder.indexOf(targetPage);
    if (targetIndex < 0) {
      return pageOrder.length;
    }

    const rect = thumbnail.getBoundingClientRect();
    return clientX < rect.left + rect.width / 2 ? targetIndex : targetIndex + 1;
  }

  function syncPageOrderFromPagesMapper(pagesMapper, selectedPagePositions = []) {
    if (!pagesMapper?.pagesNumber || typeof pagesMapper.getPrevPageNumber !== "function") {
      return;
    }

    const previousOrder = pageOrder;
    const currentPage = getApp()?.page ?? previousOrder[0] ?? 1;
    const nextOrder = [];
    for (let index = 1; index <= pagesMapper.pagesNumber; index++) {
      const sourcePage = pagesMapper.getPrevPageNumber(index);
      if (sourcePage > 0) {
        nextOrder.push(sourcePage);
      }
    }

    if (nextOrder.length === 0) {
      return;
    }

    const targetPage = resolvePageAfterPageOrderChange(previousOrder, nextOrder, currentPage);
    pageOrder = nextOrder;
    pageStateDirty = true;
    const movedPages = selectedPagePositions
      .map(position => nextOrder[Number(position) - 1])
      .filter(page => Number.isInteger(page));
    selectedPages = new Set(movedPages.length > 0 ? movedPages : [targetPage]);
    postPageOrder();
    goToPage(targetPage);
  }

  function resolvePageAfterPageOrderChange(previousOrder, nextOrder, currentPage = getApp()?.page ?? previousOrder[0] ?? 1) {
    if (nextOrder.includes(currentPage)) {
      return currentPage;
    }

    const previousIndex = previousOrder.indexOf(currentPage);
    const fallbackIndex = previousIndex >= 0
      ? Math.min(previousIndex, nextOrder.length - 1)
      : 0;
    return nextOrder[Math.max(fallbackIndex, 0)] ?? 1;
  }

  function clearDropIndicators() {
    document.querySelectorAll("#thumbnailsView .thumbnail.asta-drop-target").forEach(thumbnail => {
      thumbnail.classList.remove("asta-drop-target", "asta-drop-before", "asta-drop-after");
    });
  }

  function updateNativeDropIndicator(clientX, clientY) {
    ensureDropIndicatorStyle();
    clearDropIndicators();

    const thumbnail = getThumbnailFromPoint(clientX, clientY);
    if (!thumbnail) {
      return;
    }

    const rect = thumbnail.getBoundingClientRect();
    const placement = clientX < rect.left + rect.width / 2 ? "asta-drop-before" : "asta-drop-after";
    thumbnail.classList.add("asta-drop-target", placement);
  }

  function handleNativePageTransferDrop(data) {
    clearDropIndicators();
    const raw = data.payload;
    if (!raw) {
      return;
    }

    try {
      const payload = typeof raw === "string" ? JSON.parse(raw) : raw;
      postMessage({
        type: "insertExternalPages",
        sourcePath: payload.sourcePath,
        pages: Array.isArray(payload.pages) ? payload.pages : [],
        insertionIndex: getDropInsertionIndex(data.clientX, data.clientY)
      });
    } catch (error) {
      postDiagnostic("error", error?.message ?? "Dropped page data could not be read.");
    }
  }

  function handleNativeFileDrop(data) {
    clearDropIndicators();
    if (!Array.isArray(data.paths) || data.paths.length === 0) {
      return;
    }

    postMessage({
      type: "insertExternalFiles",
      paths: data.paths,
      insertionIndex: getDropInsertionIndex(data.clientX, data.clientY)
    });
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
    const previousOrder = pageOrder;
    const currentPage = getApp()?.page ?? previousOrder[0] ?? 1;
    pageOrder = pageOrder.filter(page => !selectedPages.has(page));
    const targetPage = resolvePageAfterPageOrderChange(previousOrder, pageOrder, currentPage);
    selectedPages = new Set([targetPage]);
    pageStateDirty = true;
    postPageOrder();
    goToPage(targetPage);
  }

  function reversePageOrder() {
    if (pageOrder.length <= 1) return;
    const previousOrder = pageOrder;
    const currentPage = getApp()?.page ?? previousOrder[0] ?? 1;
    pageOrder = [...pageOrder].reverse();
    const targetPage = resolvePageAfterPageOrderChange(previousOrder, pageOrder, currentPage);
    selectedPages = new Set([targetPage]);
    pageStateDirty = true;
    postPageOrder();
    goToPage(targetPage);
  }

  function undoPageEdit() {
    const undoButton = document.getElementById("viewsManagerStatusUndoButton");
    const undoBar = document.getElementById("viewsManagerStatusUndo");
    if (!undoButton || !undoBar || undoBar.classList.contains("hidden")) {
      return false;
    }

    undoButton.click();
    return true;
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
        pdfBase64
      });
    } catch (error) {
      postMessage({
        type: "overlayPdfExportFailed",
        requestId: data.requestId ?? null,
        message: error?.message ?? String(error)
      });
      throw error;
    }
  }

  function handleCommand(command) {
    switch (command) {
      case "nextPage":
        goRelativeInPageOrder(1);
        break;
      case "prevPage":
        goRelativeInPageOrder(-1);
        break;
      case "firstPage":
        goToPageOrderBoundary(false);
        break;
      case "lastPage":
        goToPageOrderBoundary(true);
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
        if (!undoPageEdit() && !window.EditorAdapter?.undo?.()) {
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
      } else if (data.type === "nativePageTransferDragOver" || data.type === "nativeFileDragOver") {
        updateNativeDropIndicator(data.clientX, data.clientY);
      } else if (data.type === "nativePageTransferDragLeave") {
        clearDropIndicators();
      } else if (data.type === "nativePageTransferDrop") {
        handleNativePageTransferDrop(data);
      } else if (data.type === "nativeFileDrop") {
        handleNativeFileDrop(data);
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

    eventBus?._on("pagesedited", event => {
      syncPageOrderFromPagesMapper(event.pagesMapper, Array.isArray(event.selectedPages) ? event.selectedPages : []);
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
      if (!shouldAcceptPageChange(activePage)) {
        return;
      }

      selectedPages = new Set([activePage]);
      postMessage({
        type: "activePageChanged",
        activePage,
        selectedPages: [...selectedPages]
      });
    });

    window.chrome?.webview?.addEventListener("message", handleHostMessage);
    document.addEventListener("pointerdown", captureExplicitNavigationIntent, true);
    document.addEventListener("click", captureExplicitNavigationIntent, true);
    document.addEventListener("wheel", markUserPageChangeIntent, { passive: true });
    document.addEventListener("touchstart", markUserPageChangeIntent, { passive: true });
    document.addEventListener("keydown", markKeyboardPageChangeIntent, true);
    document.addEventListener("click", markNativeViewerNavigationIntent, true);
    document.addEventListener("change", markNativeViewerNavigationIntent, true);
    document.addEventListener("input", markNativeViewerNavigationIntent, true);
    postMessage({ type: "viewerReady" });
  }

  return { initialize };
})();

AppBridge.initialize();
