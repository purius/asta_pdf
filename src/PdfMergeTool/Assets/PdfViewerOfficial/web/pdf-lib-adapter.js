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

  function editOpacity(edit, fallback = 1) {
    const numeric = Number(edit?.opacity);
    if (!Number.isFinite(numeric)) return fallback;
    return Math.min(Math.max(numeric, 0.05), 1);
  }

  function editRotation(pdfLib, edit) {
    const numeric = Number(edit?.rotate);
    if (!Number.isFinite(numeric)) return undefined;
    const degrees = Math.min(Math.max(numeric, -180), 180);
    return degrees === 0 ? undefined : pdfLib.degrees(-degrees);
  }

  function editRotationDegrees(edit) {
    const numeric = Number(edit?.rotate);
    if (!Number.isFinite(numeric)) return 0;
    return Math.min(Math.max(numeric, -180), 180);
  }

  function rotatePointAroundCenter(point, center, degrees) {
    if (!degrees) return point;
    const radians = -degrees * Math.PI / 180;
    const cos = Math.cos(radians);
    const sin = Math.sin(radians);
    const dx = point.x - center.x;
    const dy = point.y - center.y;
    return {
      x: center.x + dx * cos - dy * sin,
      y: center.y + dx * sin + dy * cos
    };
  }

  function getBox(edit, pageHeight, padding = 0) {
    const x = (Number(edit.x) || 0) - padding;
    const yFromTop = (Number(edit.y) || 0) - padding;
    const width = (Number(edit.width) || 0) + padding * 2;
    const height = (Number(edit.height) || 0) + padding * 2;
    const y = pageHeight - yFromTop - height;
    return {
      x,
      y,
      width,
      height,
      center: {
        x: x + width / 2,
        y: y + height / 2
      }
    };
  }

  function getRotatedBoxPoint(edit, pageHeight, localX, localY, padding = 0) {
    const box = getBox(edit, pageHeight, padding);
    return rotatePointAroundCenter(
      { x: box.x + localX, y: box.y + localY },
      box.center,
      editRotationDegrees(edit)
    );
  }

  function getRotatedBoxOrigin(edit, pageHeight, padding = 0) {
    return getRotatedBoxPoint(edit, pageHeight, 0, 0, padding);
  }

  function getLineEndpoints(edit, pageHeight) {
    const x = Number(edit.x) || 0;
    const yFromTop = Number(edit.y) || 0;
    const width = Number(edit.width) || 0;
    const height = Number(edit.height) || 0;
    const centerY = pageHeight - yFromTop - height / 2;
    const center = { x: x + width / 2, y: centerY };
    const degrees = editRotationDegrees(edit);
    return {
      start: rotatePointAroundCenter({ x, y: centerY }, center, degrees),
      end: rotatePointAroundCenter({ x: x + width, y: centerY }, center, degrees)
    };
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
    const thickness = Number(edit.borderWidth) || DEFAULT_STROKE_WIDTH;
    const color = colorFromArray(pdfLib, edit.borderColor, [0, 0, 0]);
    const opacity = editOpacity(edit, 1);
    const { start, end } = getLineEndpoints(edit, pageHeight);
    const angle = Math.atan2(end.y - start.y, end.x - start.x);
    const headLength = Math.max(thickness * 6, 12);
    const wingAngle = Math.PI / 7;

    page.drawLine({ start, end, color, thickness, opacity });
    for (const direction of [-1, 1]) {
      page.drawLine({
        start: end,
        end: {
          x: end.x - headLength * Math.cos(angle - direction * wingAngle),
          y: end.y - headLength * Math.sin(angle - direction * wingAngle)
        },
        color,
        thickness,
        opacity
      });
    }
  }

  function drawInkPath(page, pdfLib, edit, pageHeight) {
    const points = Array.isArray(edit.points) ? edit.points : [];
    if (points.length < 2) return;
    const color = colorFromArray(pdfLib, edit.borderColor, edit.tool === "highlight" ? [0.98, 0.8, 0.08] : [0, 0, 0]);
    const thickness = Math.max(Number(edit.borderWidth) || (edit.tool === "highlight" ? 12 : 2), 1);
    const opacity = editOpacity(edit, edit.tool === "highlight" ? 0.38 : 1);

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
    const width = Number(edit.width) || 0;
    const height = Number(edit.height) || 0;
    const size = Number(edit.size) || DEFAULT_TEXT_SIZE;
    const lineHeight = Number(edit.lineHeight) || size * 1.25;
    const textInsetX = Number(edit.textInsetX) || 0;
    const textInsetY = Number(edit.textInsetY) || 0;
    const whiteoutPadding = Number(edit.whiteoutPadding) || 0;
    const font = await resolveFont(pdfDoc, pdfLib, edit, fonts);
    const whiteoutBox = getBox(edit, pageHeight, whiteoutPadding);
    const whiteoutOrigin = getRotatedBoxOrigin(edit, pageHeight, whiteoutPadding);

    page.drawRectangle({
      x: whiteoutOrigin.x,
      y: whiteoutOrigin.y,
      width: whiteoutBox.width,
      height: whiteoutBox.height,
      color: colorFromArray(pdfLib, edit.fillColor, [1, 1, 1]),
      opacity: editOpacity(edit, 1),
      rotate: editRotation(pdfLib, edit)
    });

    const color = colorFromArray(pdfLib, edit.color, [0, 0, 0]);
    const lines = String(edit.text ?? "").split(/\r\n|\r|\n/);
    lines.forEach((line, index) => {
      const textOrigin = getRotatedBoxPoint(
        edit,
        pageHeight,
        textInsetX,
        height - textInsetY - size - index * lineHeight
      );
      page.drawText(line, {
        x: textOrigin.x,
        y: textOrigin.y,
        size,
        font,
        color,
        opacity: editOpacity(edit, 1),
        rotate: editRotation(pdfLib, edit)
      });
    });
  }

  async function drawTextOverlay(pdfDoc, page, pdfLib, edit, pageHeight, fonts) {
    const size = Number(edit.size) || DEFAULT_TEXT_SIZE;
    const lineHeight = Number(edit.lineHeight) || size * 1.25;
    const height = Number(edit.height) || lineHeight;
    const font = await resolveFont(pdfDoc, pdfLib, edit, fonts);
    const color = colorFromArray(pdfLib, edit.color, [0, 0, 0]);
    const lines = String(edit.text ?? "").split(/\r\n|\r|\n/);

    lines.forEach((line, index) => {
      const textOrigin = getRotatedBoxPoint(edit, pageHeight, 0, height - size - index * lineHeight);
      page.drawText(line, {
        x: textOrigin.x,
        y: textOrigin.y,
        size,
        font,
        color,
        opacity: editOpacity(edit, 1),
        rotate: editRotation(pdfLib, edit)
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
    const edits = (Array.isArray(options.edits) ? options.edits : [])
      .slice()
      .sort((left, right) => (Number(left.zIndex) || 0) - (Number(right.zIndex) || 0));
    const fonts = options.fonts ?? {};

    for (const edit of edits) {
      const pageIndex = Math.max((Number(edit.page) || 1) - 1, 0);
      const page = pdfDoc.getPage(pageIndex);
      const pageHeight = page.getHeight();
      const x = Number(edit.x) || 0;
      const yFromTop = Number(edit.y) || 0;

      if (edit.type === "text") {
        await drawTextOverlay(pdfDoc, page, pdfLib, edit, pageHeight, fonts);
      } else if (edit.type === "textReplace") {
        await drawTextReplacement(pdfDoc, page, pdfLib, edit, pageHeight, fonts);
      } else if (edit.type === "textHighlight") {
        const height = Number(edit.height) || 0;
        const width = Number(edit.width) || 0;
        const origin = getRotatedBoxOrigin(edit, pageHeight);
        page.drawRectangle({
          x: origin.x,
          y: origin.y,
          width,
          height,
          color: colorFromArray(pdfLib, edit.fillColor, [0.98, 0.8, 0.08]),
          opacity: editOpacity(edit, 0.38),
          rotate: editRotation(pdfLib, edit)
        });
      } else if (edit.type === "whiteout") {
        const height = Number(edit.height) || 0;
        const width = Number(edit.width) || 0;
        const isRedaction = edit.variant === "redact";
        const origin = getRotatedBoxOrigin(edit, pageHeight);
        page.drawRectangle({
          x: origin.x,
          y: origin.y,
          width,
          height,
          color: colorFromArray(pdfLib, isRedaction ? [0, 0, 0] : edit.fillColor, isRedaction ? [0, 0, 0] : [1, 1, 1]),
          opacity: isRedaction ? 1 : editOpacity(edit, 1),
          rotate: editRotation(pdfLib, edit)
        });
      } else if (edit.type === "rectangle") {
        const height = Number(edit.height) || 0;
        const width = Number(edit.width) || 0;
        const origin = getRotatedBoxOrigin(edit, pageHeight);
        const options = {
          x: origin.x,
          y: origin.y,
          width,
          height,
          borderColor: colorFromArray(pdfLib, edit.borderColor, [0, 0, 0]),
          borderWidth: Number(edit.borderWidth) || DEFAULT_STROKE_WIDTH,
          color: edit.fillColor ? colorFromArray(pdfLib, edit.fillColor, [1, 1, 1]) : undefined,
          opacity: editOpacity(edit, 1),
          rotate: editRotation(pdfLib, edit)
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
            opacity: options.opacity,
            rotate: editRotation(pdfLib, edit)
          });
        } else {
          page.drawRectangle(options);
        }
      } else if (edit.type === "line") {
        const { start, end } = getLineEndpoints(edit, pageHeight);
        page.drawLine({
          start,
          end,
          color: colorFromArray(pdfLib, edit.borderColor, [0, 0, 0]),
          thickness: Number(edit.borderWidth) || DEFAULT_STROKE_WIDTH,
          opacity: editOpacity(edit, 1)
        });
      } else if (edit.type === "arrow") {
        drawArrow(page, pdfLib, edit, pageHeight);
      } else if (edit.type === "ink") {
        drawInkPath(page, pdfLib, edit, pageHeight);
      } else if (edit.type === "image" || edit.type === "signature") {
        const width = Number(edit.width) || 0;
        const height = Number(edit.height) || 0;
        const image = await embedImage(pdfDoc, edit.imageDataUrl, edit.imageMimeType);
        const imageOrigin = getRotatedBoxOrigin(edit, pageHeight);
        page.drawImage(image, {
          x: imageOrigin.x,
          y: imageOrigin.y,
          width,
          height,
          opacity: editOpacity(edit, 1),
          rotate: editRotation(pdfLib, edit)
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
        const stampOrigin = getRotatedBoxOrigin(edit, pageHeight);
        page.drawRectangle({
          x: stampOrigin.x,
          y: stampOrigin.y,
          width,
          height,
          borderColor,
          borderWidth,
          opacity: editOpacity(edit, 1),
          rotate: editRotation(pdfLib, edit)
        });
        const textOrigin = getRotatedBoxPoint(
          edit,
          pageHeight,
          Math.max((width - textWidth) / 2, borderWidth + 2),
          height / 2 - size / 3
        );
        page.drawText(text, {
          x: textOrigin.x,
          y: textOrigin.y,
          size,
          font,
          color,
          opacity: editOpacity(edit, 1),
          rotate: editRotation(pdfLib, edit)
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
