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
    const normalized = String(base64 ?? "").replace(/^data:[^;]+;base64,/, "");
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

  async function embedImage(pdfDoc, imageDataUrl, imageMimeType) {
    const normalizedMimeType = String(imageMimeType || imageDataUrl?.match(/^data:([^;]+);/)?.[1] || "").toLowerCase();
    const imageBytes = base64ToBytes(imageDataUrl);
    if (normalizedMimeType.includes("png")) {
      return pdfDoc.embedPng(imageBytes);
    }
    if (normalizedMimeType.includes("jpeg") || normalizedMimeType.includes("jpg")) {
      return pdfDoc.embedJpg(imageBytes);
    }
    throw new Error("Only PNG and JPEG overlay images are supported.");
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

  function drawInkPath(page, pdfLib, edit, pageHeight) {
    const points = Array.isArray(edit.points) ? edit.points : [];
    if (points.length < 2) return;
    const color = colorFromArray(pdfLib, edit.borderColor, edit.tool === "highlight" ? [0.98, 0.8, 0.08] : [0, 0, 0]);
    const thickness = Math.max(Number(edit.borderWidth) || (edit.tool === "highlight" ? 12 : 2), 1);
    const opacity = edit.opacity === undefined ? (edit.tool === "highlight" ? 0.38 : 1) : Number(edit.opacity);

    for (let index = 1; index < points.length; index += 1) {
      const previous = points[index - 1];
      const current = points[index];
      page.drawLine({
        start: {
          x: Number(previous.x) || 0,
          y: pageHeight - (Number(previous.y) || 0)
        },
        end: {
          x: Number(current.x) || 0,
          y: pageHeight - (Number(current.y) || 0)
        },
        color,
        thickness,
        opacity
      });
    }
  }

  async function drawTextReplacement(pdfDoc, page, pdfLib, edit, pageHeight, fonts) {
    const x = Number(edit.x) || 0;
    const yFromTop = Number(edit.y) || 0;
    const width = Number(edit.width) || 0;
    const height = Number(edit.height) || 0;
    const size = Number(edit.size) || DEFAULT_TEXT_SIZE;
    const lineHeight = Number(edit.lineHeight) || size * 1.25;
    const textInsetX = Number(edit.textInsetX) || 0;
    const textInsetY = Number(edit.textInsetY) || 0;
    const whiteoutPadding = Number(edit.whiteoutPadding) || 0;
    const font = await resolveFont(pdfDoc, pdfLib, edit, fonts);

    page.drawRectangle({
      x: x - whiteoutPadding,
      y: pageHeight - yFromTop - height - whiteoutPadding,
      width: width + whiteoutPadding * 2,
      height: height + whiteoutPadding * 2,
      color: colorFromArray(pdfLib, edit.fillColor, [1, 1, 1])
    });

    const color = colorFromArray(pdfLib, edit.color, [0, 0, 0]);
    const lines = String(edit.text ?? "").split(/\r\n|\r|\n/);
    lines.forEach((line, index) => {
      page.drawText(line, {
        x: x + textInsetX,
        y: pageHeight - yFromTop - textInsetY - size - index * lineHeight,
        size,
        font,
        color
      });
    });
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
      } else if (edit.type === "textReplace") {
        await drawTextReplacement(pdfDoc, page, pdfLib, edit, pageHeight, fonts);
      } else if (edit.type === "textHighlight") {
        const height = Number(edit.height) || 0;
        const width = Number(edit.width) || 0;
        page.drawRectangle({
          x,
          y: pageHeight - yFromTop - height,
          width,
          height,
          color: colorFromArray(pdfLib, edit.fillColor, [0.98, 0.8, 0.08]),
          opacity: edit.opacity === undefined ? 0.38 : Number(edit.opacity)
        });
      } else if (edit.type === "whiteout") {
        const height = Number(edit.height) || 0;
        const width = Number(edit.width) || 0;
        page.drawRectangle({
          x,
          y: pageHeight - yFromTop - height,
          width,
          height,
          color: colorFromArray(pdfLib, edit.fillColor, [1, 1, 1])
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
      } else if (edit.type === "ink") {
        drawInkPath(page, pdfLib, edit, pageHeight);
      } else if (edit.type === "image" || edit.type === "signature") {
        const width = Number(edit.width) || 0;
        const height = Number(edit.height) || 0;
        const image = await embedImage(pdfDoc, edit.imageDataUrl, edit.imageMimeType);
        page.drawImage(image, {
          x,
          y: pageHeight - yFromTop - height,
          width,
          height
        });
      } else if (edit.type === "stamp") {
        const width = Number(edit.width) || 0;
        const height = Number(edit.height) || 0;
        const borderWidth = Number(edit.borderWidth) || 3;
        const color = colorFromArray(pdfLib, edit.color, [0.86, 0.15, 0.15]);
        const borderColor = colorFromArray(pdfLib, edit.borderColor, [0.86, 0.15, 0.15]);
        const font = await resolveFont(pdfDoc, pdfLib, edit, fonts);
        const size = Number(edit.size) || Math.max(Math.min(height * 0.38, 24), 10);
        const text = String(edit.text ?? "STAMP").toUpperCase();
        const textWidth = font.widthOfTextAtSize(text, size);
        page.drawRectangle({
          x,
          y: pageHeight - yFromTop - height,
          width,
          height,
          borderColor,
          borderWidth
        });
        page.drawText(text, {
          x: x + Math.max((width - textWidth) / 2, borderWidth + 2),
          y: pageHeight - yFromTop - height / 2 - size / 3,
          size,
          font,
          color
        });
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
