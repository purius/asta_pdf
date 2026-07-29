import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";

const viewer = await readFile(
  new URL("../src/PdfMergeTool/Assets/PdfViewerOfficial/web/viewer.mjs", import.meta.url),
  "utf8"
);

assert.match(viewer, /enableSplitMerge:\s*\{\s*value:\s*false,/s);

const adapter = await readFile(
  new URL("../src/PdfMergeTool/Assets/PdfViewerOfficial/web/app-adapter.js", import.meta.url),
  "utf8"
);

assert.doesNotMatch(adapter, /type:\s*"pageOrderChanged"/);
assert.doesNotMatch(adapter, /function undoPageEdit\(/);
assert.doesNotMatch(adapter, /nativePageTransferDrop/);
assert.doesNotMatch(adapter, /nativeFileDrop/);
assert.match(adapter, /function configurePreviewOnlyViewer\(/);
assert.match(adapter, /case "goToPage":\s*goToPage\(options\?\.pageNumber\);/s);
assert.match(adapter, /let latestRequestedLoadId = 0;/);
assert.match(adapter, /let currentDocumentLoadId = 0;/);
assert.match(adapter, /function queuePdfOpen\(/);
assert.match(adapter, /loadId:\s*currentDocumentLoadId/);
assert.match(adapter, /loadId:\s*data\.loadId/);
assert.match(adapter, /currentDocumentLoadId = 0;\s*window\.AstaViewerLoadId = 0;\s*window\.EditorAdapter\?\.clear/s);

const mainWindow = await readFile(
  new URL("../src/PdfMergeTool/MainWindow.xaml.cs", import.meta.url),
  "utf8"
);
const xaml = await readFile(
  new URL("../src/PdfMergeTool/MainWindow.xaml", import.meta.url),
  "utf8"
);

assert.match(mainWindow, /EditorDocumentState\? _pageOrganizerState/);
assert.match(mainWindow, /void ApplyPageOrganizerState\(/);
assert.match(mainWindow, /OnPageOrganizerCheckBoxPreviewMouseLeftButtonDown/);
assert.match(mainWindow, /OnPageOrganizerItemPreviewMouseLeftButtonDown/);
assert.match(mainWindow, /void HandleNativePageTransferDrop\(/);
assert.match(mainWindow, /void HandleNativeFileDrop\(/);
assert.match(mainWindow, /int _pendingLoadGeneration/);
assert.match(mainWindow, /loadId = loadGeneration/);
assert.match(mainWindow, /bool IsCurrentViewerLoadMessage\(/);
assert.match(mainWindow, /AreBrowserAcceleratorKeysEnabled = false/);
assert.match(mainWindow, /DocumentOperationCoordinator _documentOperations/);
assert.match(mainWindow, /RunCurrentDocumentMutationAsync\(/);
assert.match(mainWindow, /SetDocumentMutationUiState\(/);
assert.match(mainWindow, /PageOrganizerList\.IsHitTestVisible = false/);
assert.match(mainWindow, /PdfViewer\.IsHitTestVisible = false/);
assert.match(mainWindow, /OnViewerPreviewKeyDown/);
assert.match(mainWindow, /MinimumHorizontalDragDistance/);
assert.match(mainWindow, /Mouse\.Capture\(PageOrganizerList, CaptureMode\.SubTree\)/);
assert.match(mainWindow, /AdjustPageOrganizerThumbnailHeight/);
assert.match(mainWindow, /OnPageOrganizerZoomSliderValueChanged/);
assert.match(mainWindow, /SetPageOrganizerThumbnailHeight\(/);
assert.match(mainWindow, /ThumbnailCardWidth/);
assert.match(mainWindow, /GroupBy\(entry => Math\.Round\(entry\.Origin\.Y\)\)/);
assert.match(mainWindow, /OnThumbZoomInClick\(sender, e\);/);
assert.doesNotMatch(mainWindow, /SendViewerCommand\("thumbZoomIn"\)/);
assert.match(mainWindow, /_fallbackRenderService\.OpenDocument\(sourcePath\)/);
assert.doesNotMatch(mainWindow, /exportA4PageImages/);
assert.match(xaml, /x:Name="PageOrganizerList"/);
assert.match(xaml, /PreviewMouseLeftButtonDown="OnPageOrganizerItemPreviewMouseLeftButtonDown"/);
assert.match(xaml, /PreviewMouseLeftButtonUp="OnPageOrganizerItemPreviewMouseLeftButtonUp"/);
assert.match(xaml, /PreviewKeyDown="OnViewerPreviewKeyDown"/);
assert.match(xaml, /x:Name="PageOrganizerZoomSlider"/);
assert.match(xaml, /<WrapPanel Orientation="Horizontal"\s*\/>/);
assert.match(xaml, /ScrollViewer\.HorizontalScrollBarVisibility="Disabled"/);
assert.match(xaml, /Property="Width" Value="\{Binding ThumbnailCardWidth\}"/);
assert.match(xaml, /Width="\{Binding ThumbnailWidth\}"/);
assert.match(xaml, /Height="\{Binding ThumbnailHeight\}"/);

const emittedMessages = [];
const eventHandlers = new Map();
let hostMessageHandler;
const openRequests = [];
const app = {
  initializedPromise: Promise.resolve(),
  pagesCount: 8,
  pdfDocument: { numPages: 8 },
  page: 1,
  pdfViewer: { currentScale: 1 },
  viewsManager: { close() {} },
  eventBus: {
    _on(eventName, handler) {
      eventHandlers.set(eventName, handler);
    }
  },
  open({ url }) {
    return new Promise(resolve => openRequests.push({ url, resolve }));
  }
};
const windowMock = {
  PDFViewerApplication: app,
  EditorAdapter: { clear() {}, initialize() {} },
  chrome: {
    webview: {
      addEventListener(eventName, handler) {
        if (eventName === "message") {
          hostMessageHandler = handler;
        }
      },
      postMessage(message) {
        emittedMessages.push(message);
      }
    }
  }
};
vm.runInNewContext(adapter, {
  window: windowMock,
  document: { getElementById: () => ({ setAttribute() {} }) },
  Promise,
  Uint8Array,
  setTimeout
});

const flushAsync = () => new Promise(resolve => setTimeout(resolve, 0));
await flushAsync();
assert.equal(typeof hostMessageHandler, "function");

hostMessageHandler({ data: { type: "loadPdf", loadId: 1, url: "first.pdf" } });
await flushAsync();
assert.deepEqual(openRequests.map(request => request.url), ["first.pdf"]);

hostMessageHandler({ data: { type: "loadPdf", loadId: 2, url: "second.pdf" } });
hostMessageHandler({ data: { type: "command", command: "goToPage", loadId: 2, options: { pageNumber: 4 } } });
openRequests[0].resolve();
await flushAsync();
assert.deepEqual(openRequests.map(request => request.url), ["first.pdf", "second.pdf"]);

eventHandlers.get("pagerendered")({ pageNumber: 1 });
assert.equal(
  emittedMessages.filter(message => message.type === "viewerFirstPageRendered").length,
  0
);
openRequests[1].resolve();
await flushAsync();
eventHandlers.get("pagerendered")({ pageNumber: 4 });
eventHandlers.get("pagechanging")({ pageNumber: 4 });

assert.equal(app.page, 4);
assert.deepEqual(
  emittedMessages
    .filter(message => message.type === "viewerFirstPageRendered")
    .map(({ type, pageNumber, loadId }) => ({ type, pageNumber, loadId })),
  [{ type: "viewerFirstPageRendered", pageNumber: 4, loadId: 2 }]
);
assert.deepEqual(
  emittedMessages
    .filter(message => message.type === "activePageChanged")
    .map(({ type, activePage, loadId }) => ({ type, activePage, loadId })),
  [{ type: "activePageChanged", activePage: 4, loadId: 2 }]
);
