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
const thumbnailScheduler = await readFile(
  new URL("../src/PdfMergeTool/Services/PageOrganizerThumbnailScheduler.cs", import.meta.url),
  "utf8"
);
const pageOrganizerViewport = await readFile(
  new URL("../src/PdfMergeTool/Services/PageOrganizerViewport.cs", import.meta.url),
  "utf8"
);

const pageOrganizerThumbnailCapPattern =
  /pageNumbers\s*\.\s*Take\s*\(\s*96\s*\)/;
const thumbnailLoadingOverlayPattern =
  /<ProgressBar(?=\s|>)[^>]*\bIsHitTestVisible\s*=\s*"False"[^>]*>[\s\S]*?<ProgressBar\.Style>[\s\S]*?<Setter\s+Property\s*=\s*"Visibility"\s+Value\s*=\s*"Collapsed"\s*\/>[\s\S]*?<DataTrigger\s+Binding\s*=\s*"\{Binding\s+IsThumbnailLoading\}"\s+Value\s*=\s*"True"\s*>[\s\S]*?<Setter\s+Property\s*=\s*"Visibility"\s+Value\s*=\s*"Visible"\s*\/>[\s\S]*?<\/DataTrigger>[\s\S]*?<\/ProgressBar>/;
const thumbnailRetryOverlayPattern =
  /<Button(?=\s|>)[^>]*\bClick\s*=\s*"OnPageOrganizerThumbnailRetryClick"[^>]*>[\s\S]*?<Button\.Style>[\s\S]*?<Setter\s+Property\s*=\s*"Visibility"\s+Value\s*=\s*"Collapsed"\s*\/>[\s\S]*?<DataTrigger\s+Binding\s*=\s*"\{Binding\s+IsThumbnailFailed\}"\s+Value\s*=\s*"True"\s*>[\s\S]*?<Setter\s+Property\s*=\s*"Visibility"\s+Value\s*=\s*"Visible"\s*\/>[\s\S]*?<\/DataTrigger>[\s\S]*?<iconPacks:PackIconMaterial(?=\s|\/|>)[^>]*\bKind\s*=\s*"Refresh"[^>]*\/>[\s\S]*?<\/Button>/;
const thumbnailImagePattern =
  /<Image(?=\s|\/|>)[^>]*\bSource\s*=\s*"\{Binding\s+Thumbnail\}"[^>]*>/;

function escapeRegExp(value) {
  return value.replace(/[\\^$.*+?()\[\]{}|]/g, "\\$&");
}

function getElementBlocks(markup, elementName) {
  const escapedElementName = escapeRegExp(elementName);
  const tagPattern = new RegExp(
    `<\\/?${escapedElementName}(?=\\s|/|>)[^>]*>`,
    "g"
  );
  const blocks = [];
  const openOffsets = [];

  for (const match of markup.matchAll(tagPattern)) {
    const tag = match[0];
    const isClosingTag = tag.startsWith("</");
    const isSelfClosingTag = /\/\s*>$/.test(tag);

    if (isClosingTag) {
      const openOffset = openOffsets.pop();
      if (openOffset !== undefined) {
        blocks.push(markup.slice(openOffset, match.index + tag.length));
      }
    } else if (isSelfClosingTag) {
      if (openOffsets.length === 0) {
        blocks.push(tag);
      }
    } else {
      openOffsets.push(match.index);
    }
  }

  return blocks;
}

function getThumbnailDataTemplate(markup) {
  const thumbnailDataTemplate = getElementBlocks(markup, "DataTemplate").find(
    template => thumbnailImagePattern.test(template)
  );
  assert.ok(
    thumbnailDataTemplate,
    "Thumbnail DataTemplate containing Image Source=\"{Binding Thumbnail}\" is missing."
  );
  return thumbnailDataTemplate;
}

function getOpeningTag(block) {
  return block.slice(0, block.indexOf(">") + 1);
}

function getMethodBlock(source, methodName, nextMethodName) {
  const start = source.indexOf(methodName);
  const end = source.indexOf(nextMethodName, start);
  assert.notEqual(start, -1, `Missing ${methodName}.`);
  assert.notEqual(end, -1, `Missing method after ${methodName}.`);
  return source.slice(start, end);
}

function hasAttribute(block, name, value) {
  return new RegExp(
    "\\b" + escapeRegExp(name) + "\\s*=\\s*\"" + escapeRegExp(value) + "\""
  ).test(getOpeningTag(block));
}

function assertThumbnailOverlayContract(markup) {
  const thumbnailDataTemplate = getThumbnailDataTemplate(markup);
  const thumbnailLoadingOverlay = getElementBlocks(
    thumbnailDataTemplate,
    "ProgressBar"
  ).find(block => hasAttribute(block, "IsHitTestVisible", "False"));
  assert.ok(
    thumbnailLoadingOverlay,
    "Thumbnail loading ProgressBar is missing from the thumbnail DataTemplate."
  );
  assert.match(thumbnailLoadingOverlay, thumbnailLoadingOverlayPattern);

  const thumbnailRetryOverlay = getElementBlocks(
    thumbnailDataTemplate,
    "Button"
  ).find(block =>
    hasAttribute(block, "Click", "OnPageOrganizerThumbnailRetryClick")
  );
  assert.ok(
    thumbnailRetryOverlay,
    "Thumbnail retry Button is missing from the thumbnail DataTemplate."
  );
  assert.doesNotMatch(
    getOpeningTag(thumbnailRetryOverlay),
    /\bIsHitTestVisible\s*=\s*"False"/i,
    "Thumbnail retry Button must remain interactive."
  );
  assert.match(thumbnailRetryOverlay, thumbnailRetryOverlayPattern);
}

assert.match(mainWindow, /EditorDocumentState\? _pageOrganizerState/);
assert.match(mainWindow, /void ApplyPageOrganizerState\(/);
assert.match(mainWindow, /OnPageOrganizerCheckBoxPreviewMouseLeftButtonDown/);
assert.match(mainWindow, /OnPageOrganizerItemPreviewMouseLeftButtonDown/);
assert.match(
  mainWindow,
  /private void OnPageOrganizerItemPreviewMouseLeftButtonDown\(object sender, MouseButtonEventArgs e\)[\s\S]*?FindVisualAncestor<CheckBox>\(source\) is not null\s*\|\|\s*FindVisualAncestor<Button>\(source\) is not null\s*\)\s*\{\s*return;\s*\}/
);
assert.match(
  mainWindow,
  /private void OnPageOrganizerItemPreviewMouseLeftButtonUp\(object sender, MouseButtonEventArgs e\)\s*\{\s*if\s*\(\s*e\.OriginalSource is DependencyObject source\s*&&\s*FindVisualAncestor<Button>\(source\) is not null\s*\)\s*\{\s*return;\s*\}/s
);
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
assert.match(mainWindow, /UpdatePageOrganizerDropIndicator\(GetPageOrganizerInsertionIndex\(e\)\)/);
assert.match(mainWindow, /ClearPageOrganizerDropIndicator\(\)/);
assert.match(mainWindow, /IsDropBefore/);
assert.match(mainWindow, /IsDropAfter/);
assert.match(mainWindow, /OnThumbZoomInClick\(sender, e\);/);
assert.doesNotMatch(mainWindow, /SendViewerCommand\("thumbZoomIn"\)/);
assert.match(mainWindow, /_fallbackRenderService\.OpenDocument\(sourcePath\)/);
assert.doesNotMatch(mainWindow, /exportA4PageImages/);
assert.match(xaml, /x:Name="PageOrganizerList"/);
assert.match(xaml, /PreviewMouseLeftButtonDown="OnPageOrganizerItemPreviewMouseLeftButtonDown"/);
assert.match(xaml, /PreviewMouseLeftButtonUp="OnPageOrganizerItemPreviewMouseLeftButtonUp"/);
assert.match(xaml, /DragLeave="OnPageOrganizerDragLeave"/);
assert.match(xaml, /PreviewKeyDown="OnViewerPreviewKeyDown"/);
assert.match(xaml, /x:Name="PageOrganizerZoomSlider"/);
assert.match(xaml, /<WrapPanel Orientation="Horizontal"\s*\/>/);
assert.match(xaml, /ScrollViewer\.HorizontalScrollBarVisibility="Disabled"/);
assert.match(xaml, /Property="Width" Value="\{Binding ThumbnailCardWidth\}"/);
assert.match(xaml, /Width="\{Binding ThumbnailWidth\}"/);
assert.match(xaml, /Height="\{Binding ThumbnailHeight\}"/);
assert.match(xaml, /x:Name="DropBeforeIndicator"/);
assert.match(xaml, /x:Name="DropAfterIndicator"/);
assert.match(xaml, /Binding="\{Binding IsDropBefore\}"/);
assert.match(xaml, /Binding="\{Binding IsDropAfter\}"/);
const thumbnailDataTemplate = getThumbnailDataTemplate(xaml);
const thumbnailLoadingOverlay = getElementBlocks(
  thumbnailDataTemplate,
  "ProgressBar"
).find(block => thumbnailLoadingOverlayPattern.test(block));
const thumbnailRetryOverlay = getElementBlocks(
  thumbnailDataTemplate,
  "Button"
).find(block => thumbnailRetryOverlayPattern.test(block));
assert.ok(thumbnailLoadingOverlay, "Expected loading overlay fixture in thumbnail DataTemplate.");
assert.ok(thumbnailRetryOverlay, "Expected retry overlay fixture in thumbnail DataTemplate.");

const unrelatedOverlayTemplate = `<DataTemplate x:Key="UnrelatedThumbnailOverlayTemplate">${thumbnailLoadingOverlay}${thumbnailRetryOverlay}</DataTemplate>`;
const thumbnailOverlaysRemovedXaml = xaml
  .replace(thumbnailLoadingOverlay, "")
  .replace(thumbnailRetryOverlay, "")
  .replace(
    "</Window.Resources>",
    `${unrelatedOverlayTemplate}</Window.Resources>`
  );
assert.throws(
  () => assertThumbnailOverlayContract(thumbnailOverlaysRemovedXaml),
  /thumbnail DataTemplate/
);
const disabledRetryXaml = xaml.replace(
  thumbnailRetryOverlay,
  thumbnailRetryOverlay.replace(
    "<Button ",
    '<Button IsHitTestVisible="False" '
  )
);
assert.throws(
  () => assertThumbnailOverlayContract(disabledRetryXaml),
  /interactive/
);
assertThumbnailOverlayContract(xaml);
assert.match(mainWindow, /QueueActivePageFollow\(/);
assert.match(mainWindow, /FollowActivePageOrganizerItem\(/);
assert.match(mainWindow, /PageOrganizerViewport\.GetVerticalOffsetToReveal/);
assert.match(mainWindow, /OnPageOrganizerPreviewKeyDown/);
assert.match(mainWindow, /PageOrganizerList\.Focus\(\)/);
assert.match(mainWindow, /Key\.PageUp/);
assert.match(mainWindow, /Key\.PageDown/);
assert.match(mainWindow, /Key\.Home/);
assert.match(mainWindow, /Key\.End/);
assert.match(xaml, /x:Name="PageOrganizerList"[\s\S]*?Focusable="True"[\s\S]*?PreviewKeyDown="OnPageOrganizerPreviewKeyDown"/);
assert.match(mainWindow, /PageOrganizerThumbnailScheduler/);
assert.match(mainWindow, /ObservableCollection<PageOrganizerRow> PageOrganizerRows/);
assert.match(mainWindow, /Dictionary<int, PageOrganizerItem> _pageOrganizerItemsByPageNumber/);
assert.match(mainWindow, /void RebuildPageOrganizerRows\(/);
assert.match(mainWindow, /void OnPageOrganizerListSizeChanged\(/);
assert.match(mainWindow, /scheduler\.UpdateCacheOrder\(state\.PageNumbers\)/);
assert.match(mainWindow, /_pageOrganizerThumbnailViewportRefreshQueued/);
assert.match(mainWindow, /OnPageOrganizerThumbnailScrollChanged/);
assert.match(mainWindow, /RefreshPageOrganizerThumbnailViewport/);
assert.match(mainWindow, /OnPageOrganizerThumbnailRetryClick/);
assert.match(mainWindow, /PageOrganizerThumbnailRenderState.Failed/);
for (const restoredThumbnailCap of ["pageNumbers.Take(96)", "pageNumbers . Take ( 96 )"]) {
  assert.match(restoredThumbnailCap, pageOrganizerThumbnailCapPattern);
}
assert.doesNotMatch(mainWindow, pageOrganizerThumbnailCapPattern);
assert.match(thumbnailScheduler, /const int MaximumCacheWindowSize = 96/);
assert.match(thumbnailScheduler, /void UpdateCacheOrder\(/);
assert.match(mainWindow, /_pageOrganizerThumbnailCacheWindow = cacheWindow\.ToHashSet\(\)/);
assert.match(mainWindow, /item\.ThumbnailRenderState = PageOrganizerThumbnailRenderState\.Loading/);
assert.match(
  mainWindow,
  /item\.ThumbnailRenderState is PageOrganizerThumbnailRenderState\.Pending or\s+PageOrganizerThumbnailRenderState\.Evicted/s
);
assert.doesNotMatch(mainWindow, /item\.ThumbnailRenderState != PageOrganizerThumbnailRenderState\.Failed/);
assert.match(mainWindow, /scheduler\.Request\(item\.PageNumber, priority: false\)/);
assert.match(
  mainWindow,
  /scheduler\.Prioritize\(cacheWindow\);\s*scheduler\.Prioritize\(visiblePageNumbers\);/s
);
assert.match(
  mainWindow,
  /if \(orderChanged\)[\s\S]*?PageOrganizerItems\.Add\(item\);[\s\S]*?QueuePageOrganizerThumbnailViewportRefresh\(\);/s
);
assert.match(
  mainWindow,
  /private void QueuePageOrganizerThumbnailViewportRefresh\(\)[\s\S]*?IsCurrentPageOrganizerThumbnailRequest\([\s\S]*?DispatcherPriority\.Loaded[\s\S]*?IsCurrentPageOrganizerThumbnailRequest\(/s
);

const organizerList = getElementBlocks(xaml, "ListBox").find(block =>
  hasAttribute(block, "x:Name", "PageOrganizerList")
);
assert.ok(organizerList, "Page Organizer ListBox is missing.");
assert.match(organizerList, /ItemsSource="\{Binding PageOrganizerRows,/);
assert.match(organizerList, /ScrollViewer\.CanContentScroll="True"/);
assert.match(organizerList, /<VirtualizingStackPanel[\s\S]*?VirtualizationMode="Recycling"/);

const queueViewportRefresh = getMethodBlock(
  mainWindow,
  "private void QueuePageOrganizerThumbnailViewportRefresh()",
  "private void RefreshDirtyState()"
);
assert.match(queueViewportRefresh, /if \(_pageOrganizerThumbnailViewportRefreshQueued\)/);
assert.match(queueViewportRefresh, /_pageOrganizerThumbnailViewportRefreshQueued = true/);

const activePageFollow = getMethodBlock(
  mainWindow,
  "private void FollowActivePageOrganizerItem(int pageNumber, int followRevision)",
  "private bool TryRevealPageOrganizerRowIfRealized"
);
assert.doesNotMatch(activePageFollow, /ScrollIntoView/);
assert.match(activePageFollow, /ScrollToEstimatedPageOrganizerRow\(rowIndex\)/);
assert.match(activePageFollow, /followRevision == _pageOrganizerFollowRevision/);
assert.match(
  pageOrganizerViewport,
  /GetVerticalOffsetToRevealIndexedRow\(\s*double currentOffset,\s*double viewportHeight,\s*int rowIndex,\s*double estimatedRowHeight,\s*double scrollableHeight\s*\)/s
);

const pageOrganizerColumnCount = getMethodBlock(
  mainWindow,
  "private int GetPageOrganizerColumnCount()",
  "private void QueuePageOrganizerThumbnailViewportRefresh()"
);
assert.match(
  pageOrganizerColumnCount,
  /var pageOrganizerList = PageOrganizerList;\s*if \(pageOrganizerList is null\)\s*\{\s*return Math\.Max\(1, _pageOrganizerColumnCount\);\s*\}/s
);
assert.match(pageOrganizerColumnCount, /FindVisualDescendant<ScrollViewer>\(pageOrganizerList\)/);
assert.match(pageOrganizerColumnCount, /: pageOrganizerList\.ActualWidth/);

const thumbnailScrollHandler = getMethodBlock(
  mainWindow,
  "private void OnPageOrganizerThumbnailScrollChanged(",
  "private IReadOnlyList<int> GetVisiblePageOrganizerThumbnailNumbers()"
);
assert.match(thumbnailScrollHandler, /QueuePageOrganizerThumbnailViewportRefresh\(\)/);
assert.doesNotMatch(thumbnailScrollHandler, /RefreshPageOrganizerThumbnailViewport\(\);/);

const visiblePageDiscovery = getMethodBlock(
  mainWindow,
  "private IReadOnlyList<int> GetVisiblePageOrganizerThumbnailNumbers()",
  "private void RefreshPageOrganizerThumbnailViewport()"
);
assert.match(visiblePageDiscovery, /GetRealizedPageOrganizerItemContainers\(\)/);
assert.doesNotMatch(visiblePageDiscovery, /PageOrganizerItems\.Count/);
assert.doesNotMatch(visiblePageDiscovery, /ItemContainerGenerator\.ContainerFromIndex\(index\)/);

const viewportRefresh = getMethodBlock(
  mainWindow,
  "private void RefreshPageOrganizerThumbnailViewport()",
  "private void UpdateWindowTitle()"
);
assert.match(viewportRefresh, /previousCacheWindow\.Except\(cacheWindow\)/);
assert.match(viewportRefresh, /cacheWindow\.Except\(previousCacheWindow\)/);
assert.doesNotMatch(viewportRefresh, /foreach \(var item in PageOrganizerItems\)/);

const cacheWindowMethod = getMethodBlock(
  thumbnailScheduler,
  "public IReadOnlySet<int> GetCacheWindow(IReadOnlyCollection<int> visiblePageNumbers)",
  "private bool TryStart(int pageNumber, out int nextPageNumber)"
);
assert.match(cacheWindowMethod, /maximumWindowSize/);
assert.doesNotMatch(cacheWindowMethod, /foreach \(var visibleIndex/);

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
