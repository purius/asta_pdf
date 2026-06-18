(() => {
  const DEFAULT_TEXT_SIZE = 12;
  const DEFAULT_STROKE_WIDTH = 1;

  function getPdfLib() {
    const pdfLib = window.PDFLib;
    if (!pdfLib?.PDFDocument) {
      throw new Error("pdf-lib is not loaded.");
    }
    return pdfLib;
  }

  function getFontkit() {
    return window.fontkit ?? window.Fontkit ?? null;
  }

  function base64ToBytes(base64) {
    const normalized = String(base64 ?? "").replace(/^data:application\/pdf;base64,/, "");
    const binary = atob(normalized);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return bytes;
  }

  function bytesToBase64(bytes) {
    const chunkSize = 0x8000;
    const chunks = [];
    for (let index = 0; index < bytes.length; index += chunkSize) {
      chunks.push(String.fromCharCode(...bytes.subarray(index, index + chunkSize)));
    }
    return btoa(chunks.join(""));
  }

  function colorFromArray(pdfLib, color, fallback) {
    const [r, g, b] = Array.isArray(color) ? color : fallback;
    return pdfLib.rgb(
      Math.min(Math.max(Number(r) || 0, 0), 1),
      Math.min(Math.max(Number(g) || 0, 0), 1),
      Math.min(Math.max(Number(b) || 0, 0), 1)
    );
  }

  async function resolveFont(pdfDoc, pdfLib, edit, fonts) {
    const fontName = edit.fontName || "default";
    const fontSource = fonts?.[fontName] ?? edit.fontBytes ?? edit.fontBase64;

    if (fontSource) {
      const fontkit = getFontkit();
      if (!fontkit) {
        throw new Error("fontkit is required to embed custom Windows fonts.");
      }
      pdfDoc.registerFontkit(fontkit);
      const fontBytes = typeof fontSource === "string" ? base64ToBytes(fontSource) : fontSource;
      return pdfDoc.embedFont(fontBytes, { subset: true });
    }

    return pdfDoc.embedFont(pdfLib.StandardFonts.Helvetica);
  }

  function drawArrow(page, pdfLib, edit, pageHeight) {
    const x = Number(edit.x) || 0;
    const yFromTop = Number(edit.y) || 0;
    const width = Number(edit.width) || 0;
    const height = Number(edit.height) || 0;
    const thickness = Number(edit.borderWidth) || DEFAULT_STROKE_WIDTH;
    const color = colorFromArray(pdfLib, edit.borderColor, [0, 0, 0]);
    const start = { x, y: pageHeight - yFromTop };
    const end = { x: x + width, y: pageHeight - yFromTop - height };
    const angle = Math.atan2(end.y - start.y, end.x - start.x);
    const headLength = Math.max(thickness * 6, 12);
    const wingAngle = Math.PI / 7;

    page.drawLine({ start, end, color, thickness });
    for (const direction of [-1, 1]) {
      page.drawLine({
        start: end,
        end: {
          x: end.x - headLength * Math.cos(angle - direction * wingAngle),
          y: end.y - headLength * Math.sin(angle - direction * wingAngle)
        },
        color,
        thickness
      });
    }
  }

  async function createOverlayPdf(options) {
    const pdfLib = getPdfLib();
    let sourceBytes = options.sourceBytes;
    if (!sourceBytes && options.sourceUrl) {
      sourceBytes = new Uint8Array(await (await fetch(options.sourceUrl)).arrayBuffer());
    }
    sourceBytes ??= base64ToBytes(options.sourceBase64);
    const pdfDoc = await pdfLib.PDFDocument.load(sourceBytes);
    const edits = Array.isArray(options.edits) ? options.edits : [];
    const fonts = options.fonts ?? {};

    for (const edit of edits) {
      const pageIndex = Math.max((Number(edit.page) || 1) - 1, 0);
      const page = pdfDoc.getPage(pageIndex);
      const pageHeight = page.getHeight();
      const x = Number(edit.x) || 0;
      const yFromTop = Number(edit.y) || 0;

      if (edit.type === "text") {
        const font = await resolveFont(pdfDoc, pdfLib, edit, fonts);
        const size = Number(edit.size) || DEFAULT_TEXT_SIZE;
        page.drawText(String(edit.text ?? ""), {
          x,
          y: pageHeight - yFromTop - size,
          size,
          font,
          color: colorFromArray(pdfLib, edit.color, [0, 0, 0]),
          rotate: edit.rotate ? pdfLib.degrees(Number(edit.rotate) || 0) : undefined
        });
      } else if (edit.type === "rectangle") {
        const height = Number(edit.height) || 0;
        const width = Number(edit.width) || 0;
        const options = {
          x,
          y: pageHeight - yFromTop - height,
          width,
          height,
          borderColor: colorFromArray(pdfLib, edit.borderColor, [0, 0, 0]),
          borderWidth: Number(edit.borderWidth) || DEFAULT_STROKE_WIDTH,
          color: edit.fillColor ? colorFromArray(pdfLib, edit.fillColor, [1, 1, 1]) : undefined,
          opacity: edit.opacity === undefined ? undefined : Number(edit.opacity)
        };
        if (edit.shape === "ellipse") {
          page.drawEllipse({
            x: x + width / 2,
            y: pageHeight - yFromTop - height / 2,
            xScale: width / 2,
            yScale: height / 2,
            borderColor: options.borderColor,
            borderWidth: options.borderWidth,
            color: options.color,
            opacity: options.opacity
          });
        } else {
          page.drawRectangle(options);
        }
      } else if (edit.type === "line") {
        page.drawLine({
          start: { x, y: pageHeight - yFromTop },
          end: {
            x: x + (Number(edit.width) || 0),
            y: pageHeight - yFromTop - (Number(edit.height) || 0)
          },
          color: colorFromArray(pdfLib, edit.borderColor, [0, 0, 0]),
          thickness: Number(edit.borderWidth) || DEFAULT_STROKE_WIDTH
        });
      } else if (edit.type === "arrow") {
        drawArrow(page, pdfLib, edit, pageHeight);
      }
    }

    const savedBytes = await pdfDoc.save({ useObjectStreams: false });
    return bytesToBase64(savedBytes);
  }

  window.PdfLibAdapter = {
    hasPdfLib: () => Boolean(window.PDFLib?.PDFDocument),
    hasFontkit: () => Boolean(getFontkit()),
    createOverlayPdf
  };
})();
